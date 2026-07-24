"""Build a continuous Brent spot-price series from FRED and Yahoo Finance."""

from __future__ import annotations

import logging
from io import StringIO
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Literal

import pandas as pd
import requests
import yfinance as yf

LOGGER = logging.getLogger(__name__)

FRED_SERIES = "DCOILBRENTEU"
YAHOO_TICKER = "BZ=F"
FRED_CSV_URL = (
    "https://fred.stlouisfed.org/graph/fredgraph.csv?id=DCOILBRENTEU"
)
AdjustmentMethod = Literal["constant", "rolling"]


class BrentDataError(RuntimeError):
    """Raised when a valid continuous Brent series cannot be constructed."""


def _cache_is_fresh(path: Path, ttl: timedelta) -> bool:
    if not path.exists():
        return False
    age = datetime.now().timestamp() - path.stat().st_mtime
    return age <= ttl.total_seconds()


def _normalise_series(series: pd.Series, name: str) -> pd.Series:
    """Return a numeric, sorted Series with unique, timezone-naive dates."""
    result = pd.Series(
        pd.to_numeric(series, errors="coerce").to_numpy(),
        index=pd.to_datetime(series.index, errors="coerce"),
        name=name,
        dtype="float64",
    )
    result = result[~result.index.isna()].dropna()
    if isinstance(result.index, pd.DatetimeIndex) and result.index.tz is not None:
        result.index = result.index.tz_convert(None)
    result.index = result.index.normalize()
    result.index.name = None
    return result[~result.index.duplicated(keep="last")].sort_index()


def _read_cached_series(path: Path, name: str) -> pd.Series:
    frame = pd.read_csv(path, parse_dates=["date"])
    if "value" not in frame:
        raise BrentDataError(f"Invalid cache file: {path}")
    return _normalise_series(frame.set_index("date")["value"], name)


def _write_cached_series(series: pd.Series, path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix(path.suffix + ".tmp")
    series.rename("value").rename_axis("date").to_csv(temporary)
    temporary.replace(path)


def _download_fred_spot(
    start: pd.Timestamp | None = None,
    end: pd.Timestamp | None = None,
    timeout: float = 30.0,
) -> pd.Series:
    """Download the official daily Brent spot series from FRED."""
    LOGGER.info("Downloading FRED series %s", FRED_SERIES)
    try:
        params: dict[str, str] = {}
        if start is not None:
            params["cosd"] = start.date().isoformat()
        if end is not None:
            params["coed"] = end.date().isoformat()
        response = requests.get(FRED_CSV_URL, params=params, timeout=timeout)
        response.raise_for_status()
        frame = pd.read_csv(StringIO(response.text))
    except (requests.RequestException, OSError, ValueError) as exc:
        raise BrentDataError(f"Unable to download FRED {FRED_SERIES}") from exc

    date_column = "DATE" if "DATE" in frame.columns else "observation_date"
    if date_column not in frame or FRED_SERIES not in frame:
        raise BrentDataError("FRED response did not contain the expected columns")
    series = frame.set_index(date_column)[FRED_SERIES]
    return _normalise_series(series, "Official Brent spot")


def _download_yahoo_futures(
    start: pd.Timestamp,
    end: pd.Timestamp,
) -> pd.Series:
    """Download unadjusted daily BZ=F closes from Yahoo Finance."""
    LOGGER.info("Downloading Yahoo Finance ticker %s", YAHOO_TICKER)
    try:
        frame = yf.download(
            YAHOO_TICKER,
            start=start.date().isoformat(),
            # yfinance treats end as exclusive.
            end=(end + pd.Timedelta(days=1)).date().isoformat(),
            auto_adjust=False,
            progress=False,
            threads=False,
        )
    except Exception as exc:  # yfinance exposes several backend exception types
        raise BrentDataError(f"Unable to download Yahoo ticker {YAHOO_TICKER}") from exc

    if frame.empty or "Close" not in frame:
        raise BrentDataError(f"Yahoo returned no close prices for {YAHOO_TICKER}")
    close = frame["Close"]
    if isinstance(close, pd.DataFrame):
        if close.shape[1] != 1:
            raise BrentDataError("Yahoo returned an ambiguous Close-price table")
        close = close.iloc[:, 0]
    return _normalise_series(close, "Yahoo Brent futures")


def _load_or_download(
    cache_path: Path,
    name: str,
    ttl: timedelta,
    refresh: bool,
    downloader,
) -> pd.Series:
    cached: pd.Series | None = None
    if cache_path.exists():
        try:
            cached = _read_cached_series(cache_path, name)
        except (OSError, ValueError, BrentDataError):
            LOGGER.warning("Ignoring invalid cache file %s", cache_path)

    if not refresh and _cache_is_fresh(cache_path, ttl):
        LOGGER.info("Using cached %s data from %s", name, cache_path)
        if cached is not None:
            return cached

    try:
        update = downloader(cached)
        series = update if cached is None else pd.concat([cached, update])
        series = _normalise_series(series, name)
        _write_cached_series(series, cache_path)
        return series
    except BrentDataError:
        if cached is not None:
            LOGGER.warning("Download failed; using stale cache %s", cache_path)
            return cached
        raise


def get_continuous_brent_spot(
    start: str | date | pd.Timestamp | None = None,
    end: str | date | pd.Timestamp | None = None,
    *,
    adjustment: AdjustmentMethod = "constant",
    rolling_window: int = 20,
    cache_dir: str | Path = ".cache/brent",
    cache_ttl: timedelta = timedelta(hours=24),
    yahoo_cache_ttl: timedelta = timedelta(0),
    refresh_lookback_days: int = 45,
    refresh: bool = False,
) -> pd.Series:
    """Return a continuous daily Brent spot-price series in USD per barrel.

    Official FRED observations (``DCOILBRENTEU``) are retained wherever they
    exist. Adjusted Yahoo ``BZ=F`` closes fill missing official trading days
    and extend the series after the last common FRED/Yahoo trading date.

    Args:
        start: Optional inclusive first date.
        end: Optional inclusive last date. Defaults to today.
        adjustment: ``"constant"`` uses the ratio on the final overlapping
            trading date. ``"rolling"`` uses the mean ratio over the last
            ``rolling_window`` common observations.
        rolling_window: Number of overlapping observations used by the rolling
            adjustment.
        cache_dir: Directory for provider-level CSV caches.
        cache_ttl: Maximum age of the FRED cache before attempting a refresh.
        yahoo_cache_ttl: Maximum age of the Yahoo cache. The default of zero
            refreshes Yahoo's recent tail on every call while retaining cached
            history.
        refresh_lookback_days: On subsequent downloads, re-fetch this many
            calendar days before the final cached observation. This captures
            delayed FRED releases and revisions without downloading the full
            history on every run.
        refresh: Force fresh provider downloads. A stale cache remains a
            fallback if a provider is temporarily unavailable.

    Returns:
        A sorted, duplicate-free ``pandas.Series`` named ``"Brent Spot"``.
        Weekends and exchange holidays are not manufactured; gaps are filled
        only where Yahoo has an actual observation.

    Raises:
        ValueError: If arguments are invalid.
        BrentDataError: If provider data is unavailable or cannot be aligned.
    """
    if adjustment not in ("constant", "rolling"):
        raise ValueError("adjustment must be 'constant' or 'rolling'")
    if rolling_window < 1:
        raise ValueError("rolling_window must be at least 1")
    if cache_ttl.total_seconds() < 0:
        raise ValueError("cache_ttl cannot be negative")
    if yahoo_cache_ttl.total_seconds() < 0:
        raise ValueError("yahoo_cache_ttl cannot be negative")
    if refresh_lookback_days < 1:
        raise ValueError("refresh_lookback_days must be at least 1")

    end_date = pd.Timestamp(end or date.today()).normalize()
    start_date = pd.Timestamp(start).normalize() if start is not None else None
    if start_date is not None and start_date > end_date:
        raise ValueError("start must be on or before end")

    cache_root = Path(cache_dir)
    official = _load_or_download(
        cache_root / "fred_DCOILBRENTEU.csv",
        "Official Brent spot",
        cache_ttl,
        refresh,
        lambda cached: _download_fred_spot(
            (
                cached.index.max() - pd.Timedelta(days=refresh_lookback_days)
                if cached is not None and not cached.empty
                else start_date
            ),
            end_date,
        ),
    )

    # BZ=F history begins much later than the official spot series. Starting
    # at 2000 captures all available Yahoo history and ample overlap.
    yahoo_start = min(start_date, pd.Timestamp("2000-01-01")) if start_date else pd.Timestamp("2000-01-01")
    yahoo = _load_or_download(
        cache_root / "yahoo_BZ_F.csv",
        "Yahoo Brent futures",
        yahoo_cache_ttl,
        refresh,
        lambda cached: _download_yahoo_futures(
            max(
                yahoo_start,
                cached.index.max() - pd.Timedelta(days=refresh_lookback_days),
            )
            if cached is not None and not cached.empty
            else yahoo_start,
            end_date,
        ),
    )

    common_dates = official.index.intersection(yahoo.index)
    if common_dates.empty:
        raise BrentDataError(
            "FRED and Yahoo have no common trading dates; adjustment is impossible"
        )

    overlap_end = common_dates.max()
    ratios = (official.loc[common_dates] / yahoo.loc[common_dates]).replace(
        [float("inf"), float("-inf")], pd.NA
    ).dropna()
    if ratios.empty:
        raise BrentDataError("No valid FRED/Yahoo price ratios were available")

    if adjustment == "constant":
        factor_history = ratios
    else:
        factor_history = ratios.rolling(
            window=rolling_window,
            min_periods=1,
        ).mean()
    factor = float(factor_history.loc[overlap_end])
    if factor <= 0:
        raise BrentDataError(f"Invalid non-positive adjustment factor: {factor}")

    LOGGER.info(
        "Extending after %s with %s adjustment factor %.6f",
        overlap_end.date(),
        adjustment,
        factor,
    )
    # Apply the latest factor known on or before each Yahoo trading date. This
    # avoids look-ahead bias when filling an internal FRED gap. Beyond the
    # final overlap, the final factor naturally carries forward.
    dated_factors = factor_history.reindex(
        factor_history.index.union(yahoo.index)
    ).sort_index().ffill().reindex(yahoo.index)
    adjusted_yahoo = yahoo * dated_factors
    adjusted_yahoo = adjusted_yahoo.loc[adjusted_yahoo.index >= common_dates.min()]

    # combine_first gives every official observation priority while allowing
    # adjusted Yahoo values to fill official missing dates and extend the tail.
    continuous = official.combine_first(adjusted_yahoo).sort_index()
    continuous = continuous.loc[:end_date]
    if start_date is not None:
        continuous = continuous.loc[start_date:]
    if continuous.empty:
        raise BrentDataError("No Brent observations exist in the requested date range")
    continuous.name = "Brent Spot"
    return continuous


__all__ = ["BrentDataError", "get_continuous_brent_spot"]
