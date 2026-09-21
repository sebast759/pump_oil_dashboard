"""Download Brent futures OHLC data from Investing.com's JSON chart API."""

from __future__ import annotations

import argparse
import json
import math
import re
import time
from pathlib import Path
from typing import Callable, Iterable
from urllib.parse import parse_qs, urlparse

import pandas as pd
from bs4 import BeautifulSoup
from curl_cffi import requests

PAGE_URL = "https://ca.investing.com/commodities/brent-oil-streaming-chart"
API_URL = "https://api.investing.com/api/financialdata/{instrument_id}/historical/chart/"
BROWSERS = ("safari", "chrome", "edge")
ID_PATTERNS = (
    r'"(?:instrument_id|instrumentId|pair_ID|pairId|pair_id)"\s*:\s*"?(\d+)"?',
    r"(?:instrument_id|instrumentId|pair_ID|pairId|pair_id)\s*[=:]\s*['\"]?(\d+)",
)
PRIMARY_ID_PATTERNS = (
    r'"identifiers"\s*:\s*\{[^{{}}]*?"instrument_id"\s*:\s*"?(\d+)"?',
    r'"commodityStore".*?"instrument"\s*:\s*\{.*?"base"\s*:\s*\{[^{{}}]*?"id"\s*:\s*"?(\d+)"?',
)


class InvestingDataError(RuntimeError):
    """Raised when Investing.com does not return usable Brent data."""


def _bounded_request(
    getter: Callable,
    url: str,
    *,
    attempts: int = 3,
    backoff: float = 1.0,
    retry_statuses: Iterable[int] = (403, 429, 500, 502, 503, 504),
    sleep: Callable[[float], None] = time.sleep,
    **kwargs,
):
    """GET with bounded exponential backoff for explicitly retryable statuses."""
    response = None
    for attempt in range(attempts):
        response = getter(url, **kwargs)
        if response.status_code not in retry_statuses or attempt == attempts - 1:
            response.raise_for_status()
            return response
        sleep(backoff * (2**attempt))
    raise InvestingDataError(f"No response received from {url}")


def fetch_chart_page(*, timeout: float = 30.0, sleep=time.sleep) -> str:
    """Fetch the chart page, changing browser fingerprint only after a 403."""
    last_response = None
    for index, browser in enumerate(BROWSERS):
        response = requests.get(PAGE_URL, impersonate=browser, timeout=timeout)
        last_response = response
        if response.status_code != 403:
            response.raise_for_status()
            return response.text
        if index < len(BROWSERS) - 1:
            sleep(2**index)
    assert last_response is not None
    last_response.raise_for_status()
    raise InvestingDataError("Investing.com returned no chart page")


def extract_instrument_id(html: str) -> int:
    """Extract the rolling Brent instrument ID from HTML or an embedded iframe."""
    soup = BeautifulSoup(html, "html.parser")
    page_text = soup.get_text(" ", strip=True).lower()
    if "brent" not in page_text and "brent" not in html.lower():
        raise InvestingDataError("Page does not identify itself as Brent")

    primary: list[int] = []
    for pattern in PRIMARY_ID_PATTERNS:
        primary.extend(int(value) for value in re.findall(pattern, html, re.I | re.S))
    primary = list(dict.fromkeys(value for value in primary if value > 0))
    if len(primary) == 1:
        return primary[0]
    if len(primary) > 1:
        raise InvestingDataError(f"Ambiguous primary Brent instrument IDs: {primary}")

    candidates: list[int] = []
    for pattern in ID_PATTERNS:
        candidates.extend(int(value) for value in re.findall(pattern, html, re.I))

    for iframe in soup.find_all("iframe", src=True):
        parsed = urlparse(iframe["src"])
        query = parse_qs(parsed.query)
        for key in ("instrument_id", "instrumentId", "pair_ID", "pairId", "pair_id"):
            for value in query.get(key, []):
                if value.isdigit():
                    candidates.append(int(value))

    positive = list(dict.fromkeys(value for value in candidates if value > 0))
    if not positive:
        raise InvestingDataError("Could not find a Brent instrument ID in the chart page")
    # Targeted patterns normally resolve to one value. Multiple values are
    # unsafe because the page also contains related instruments and adverts.
    if len(positive) != 1:
        raise InvestingDataError(f"Ambiguous Brent instrument IDs: {positive}")
    return positive[0]


def parse_chart_json(payload: object) -> pd.DataFrame:
    """Validate and convert Investing chart rows into a UTC-indexed DataFrame."""
    if not isinstance(payload, dict) or not isinstance(payload.get("data"), list):
        raise InvestingDataError("Chart response must contain a data list")
    rows = payload["data"]
    if not rows:
        raise InvestingDataError("Chart response contained no observations")

    parsed = []
    for index, row in enumerate(rows):
        if not isinstance(row, list) or len(row) < 6:
            raise InvestingDataError(f"Invalid chart row at position {index}")
        values = row[:6]
        if any(isinstance(value, bool) or not isinstance(value, (int, float)) for value in values):
            raise InvestingDataError(f"Non-numeric chart value at position {index}")
        if any(not math.isfinite(float(value)) for value in values):
            raise InvestingDataError(f"Non-finite chart value at position {index}")
        timestamp, open_, high, low, close, volume = values
        if timestamp <= 0 or min(open_, high, low, close) <= 0 or volume < 0:
            raise InvestingDataError(f"Out-of-range chart value at position {index}")
        parsed.append((timestamp, open_, high, low, close, volume))

    frame = pd.DataFrame(parsed, columns=["timestamp", "open", "high", "low", "close", "volume"])
    frame["timestamp"] = pd.to_datetime(frame["timestamp"], unit="ms", utc=True)
    for column in ("open", "high", "low", "close", "volume"):
        frame[column] = pd.to_numeric(frame[column], errors="raise")
    return frame.drop_duplicates("timestamp", keep="last").sort_values("timestamp").reset_index(drop=True)


def fetch_history(
    instrument_id: int,
    *,
    interval: str = "P1D",
    period: str = "MAX",
    point_count: int = 5000,
    timeout: float = 30.0,
    attempts: int = 3,
    sleep=time.sleep,
) -> pd.DataFrame:
    """Fetch and parse OHLC history for a validated Investing instrument ID."""
    if instrument_id <= 0 or point_count < 1:
        raise ValueError("instrument_id and point_count must be positive")
    url = API_URL.format(instrument_id=instrument_id)
    # The endpoint currently accepts chart selector sizes 60, 70 and 120.
    # With period=MAX, 120 still returns up to 5,000 daily observations; sending
    # pointscount=5000 itself causes a server-side 500. Keep the public option
    # configurable while adapting it to the live endpoint's accepted selector.
    api_point_count = point_count if point_count in {60, 70, 120} else 120
    response = _bounded_request(
        requests.get,
        url,
        params={"interval": interval, "period": period, "pointscount": api_point_count},
        headers={"Referer": PAGE_URL, "Accept": "application/json"},
        impersonate="chrome",
        timeout=timeout,
        attempts=attempts,
        sleep=sleep,
    )
    try:
        payload = response.json()
    except (json.JSONDecodeError, ValueError) as exc:
        raise InvestingDataError("Chart endpoint returned invalid JSON") from exc
    return parse_chart_json(payload)


def download_history(*, request_pause: float = 0.75, **kwargs) -> tuple[int, pd.DataFrame]:
    """Resolve the current Brent ID and download its requested history."""
    instrument_id = extract_instrument_id(fetch_chart_page())
    time.sleep(request_pause)
    return instrument_id, fetch_history(instrument_id, **kwargs)


def update_history_cache(
    path: str | Path,
    *,
    interval: str = "P1D",
    point_count: int = 5000,
    comparison_days: int = 7,
    request_pause: float = 0.75,
) -> tuple[int, pd.DataFrame, bool]:
    """Refresh the tail and redownload MAX history when overlap was revised.

    Returns ``(instrument_id, frame, full_refresh_performed)``.
    """
    cache_path = Path(path)
    instrument_id = extract_instrument_id(fetch_chart_page())
    time.sleep(request_pause)
    recent = fetch_history(
        instrument_id,
        interval=interval,
        period="P1M",
        point_count=120,
    )
    full_refresh = not cache_path.exists()
    cached = None
    if not full_refresh:
        try:
            cached = pd.read_csv(cache_path, parse_dates=["timestamp"])
            cached["timestamp"] = pd.to_datetime(cached["timestamp"], utc=True)
            required = {"timestamp", "open", "high", "low", "close", "volume"}
            if not required.issubset(cached.columns) or cached.empty:
                full_refresh = True
            else:
                cutoff = recent["timestamp"].max() - pd.Timedelta(days=comparison_days)
                # Do not compare the newest session: its close can still move
                # while the market is open. Compare only settled overlap.
                stable_end = recent["timestamp"].max() - pd.Timedelta(days=1)
                old_overlap = cached.loc[
                    (cached["timestamp"] >= cutoff) & (cached["timestamp"] <= stable_end),
                    ["timestamp", "close"],
                ]
                new_overlap = recent.loc[
                    (recent["timestamp"] >= cutoff) & (recent["timestamp"] <= stable_end),
                    ["timestamp", "close"],
                ]
                comparison = old_overlap.merge(new_overlap, on="timestamp", suffixes=("_old", "_new"))
                if comparison.empty:
                    full_refresh = True
                else:
                    full_refresh = not all(
                        math.isclose(old, new, rel_tol=0.0, abs_tol=1e-9)
                        for old, new in zip(comparison["close_old"], comparison["close_new"])
                    )
        except (OSError, ValueError, KeyError):
            full_refresh = True

    if full_refresh:
        time.sleep(request_pause)
        result = fetch_history(
            instrument_id,
            interval=interval,
            period="MAX",
            point_count=point_count,
        )
    else:
        assert cached is not None
        result = pd.concat([cached, recent], ignore_index=True)
        result = result.drop_duplicates("timestamp", keep="last").sort_values("timestamp").reset_index(drop=True)

    cache_path.parent.mkdir(parents=True, exist_ok=True)
    temporary = cache_path.with_suffix(cache_path.suffix + ".tmp")
    result.to_csv(temporary, index=False)
    temporary.replace(cache_path)
    return instrument_id, result, full_refresh


def main() -> None:
    parser = argparse.ArgumentParser(description="Download Investing.com Brent futures OHLC history")
    parser.add_argument("--interval", default="P1D")
    parser.add_argument("--period", default="MAX")
    parser.add_argument("--point-count", type=int, default=5000)
    parser.add_argument("--output", "-o", type=Path, default=Path("investing_brent.csv"))
    args = parser.parse_args()
    instrument_id, frame = download_history(
        interval=args.interval,
        period=args.period,
        point_count=args.point_count,
    )
    args.output.parent.mkdir(parents=True, exist_ok=True)
    frame.to_csv(args.output, index=False)
    print(f"Saved {len(frame)} rows for instrument {instrument_id} to {args.output}")


if __name__ == "__main__":
    main()
