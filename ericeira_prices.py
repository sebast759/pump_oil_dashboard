"""Fetch the latest reported prices for Intermarché Ericeira from FuelFlash."""

from __future__ import annotations

import json
import re
import time
from datetime import datetime, timezone
from pathlib import Path

from bs4 import BeautifulSoup
from curl_cffi import requests

STATION_URL = (
    "https://www.fuelflash.eu/pt/intermarch%C3%A9-ericeira-ericeira-portugal-"
    "estrada-nacional-116-km-1-edif%C3%ADcio-intermarch%C3%A9-64861/"
)
CACHE_PATH = Path(__file__).parent / ".cache" / "ericeira_prices.json"
HISTORY_URL = "https://www.fuelflash.eu/chart/64861/28/"


class EriceiraPriceError(RuntimeError):
    pass


def parse_station_page(content: bytes) -> dict:
    soup = BeautifulSoup(content, "html.parser")
    heading = soup.find("h1")
    if heading is None or "ericeira" not in heading.get_text(" ", strip=True).lower():
        raise EriceiraPriceError("FuelFlash page is not the Ericeira station")

    prices = []
    for card in soup.select("div.preis_gross"):
        name_node = card.find("div")
        whole = card.select_one(".preis_part1")
        decimal = card.select_one(".preis_part2")
        if not name_node or not whole or not decimal:
            continue
        name = name_node.get_text(" ", strip=True)
        raw_price = whole.get_text(strip=True).replace(",", ".") + decimal.get_text(strip=True)
        try:
            price = float(raw_price)
        except ValueError as exc:
            raise EriceiraPriceError(f"Invalid price for {name}: {raw_price}") from exc
        text = card.get_text(" ", strip=True)
        report = re.search(r"(\d{{2}}\.\d{{2}}\.\s+\d{{2}}:\d{{2}})", text)
        prices.append({"fuel": name, "price": round(price, 3), "reported": report.group(1) if report else None})

    if not prices:
        raise EriceiraPriceError("No fuel prices found on the Ericeira page")
    return {
        "station": "Intermarché Ericeira",
        "address": "Estrada Nacional 116, Km 1 · 2655-139 Ericeira",
        "url": STATION_URL,
        "prices": prices,
        "fetched_at": datetime.now().astimezone().isoformat(timespec="minutes"),
    }


def fetch_prices(timeout: float = 30.0) -> dict:
    response = requests.get(STATION_URL, impersonate="chrome", timeout=timeout)
    response.raise_for_status()
    result = parse_station_page(response.content)
    time.sleep(0.5)
    history_response = requests.get(
        HISTORY_URL,
        impersonate="chrome",
        timeout=timeout,
        headers={"Referer": STATION_URL},
    )
    history_response.raise_for_status()
    result["history"] = parse_history_script(history_response.text)
    return result


def parse_history_script(script: str) -> list[dict]:
    columns = re.findall(r"data\.addColumn\('number',\s*'([^']+)'\s*\)", script)
    if not columns:
        raise EriceiraPriceError("No fuel columns found in FuelFlash history")
    rows = []
    pattern = r"data\.addRow\(\[new Date \((\d+)\),([^\]]+)\]\)"
    for timestamp, values_text in re.findall(pattern, script):
        values = [value.strip() for value in values_text.split(",")]
        if len(values) != len(columns):
            raise EriceiraPriceError("FuelFlash history row has an unexpected shape")
        point = {
            "date": datetime.fromtimestamp(int(timestamp) / 1000, timezone.utc).date().isoformat()
        }
        for name, value in zip(columns, values):
            try:
                point[name] = round(float(value), 3)
            except ValueError as exc:
                raise EriceiraPriceError(f"Invalid historical price: {value}") from exc
        rows.append(point)
    if not rows:
        raise EriceiraPriceError("No historical FuelFlash prices found")
    return rows


def load_prices(*, local: bool = False, cache_path: Path = CACHE_PATH) -> dict | None:
    if not local:
        try:
            result = fetch_prices()
            cache_path.parent.mkdir(parents=True, exist_ok=True)
            temporary = cache_path.with_suffix(".tmp")
            temporary.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
            temporary.replace(cache_path)
            return result
        except Exception as exc:
            print(f"  WARNING: Ericeira price download failed: {exc} — using cache")
    if cache_path.exists():
        return json.loads(cache_path.read_text(encoding="utf-8"))
    return None
