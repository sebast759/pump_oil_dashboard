#!/usr/bin/env python3
"""
EU Weekly Oil Bulletin Dashboard Generator
==========================================
Reads the EC Weekly Oil Bulletin Excel file and produces a self-contained
HTML dashboard with:
  - 36-month historical price lines (EU, FR, DE, NL, ES, IT, PT)
  - YTD % change bars
  - Tax vs pre-tax breakdown
  - Consumption mix by country

Usage:
    python generate_oil_dashboard.py                     # uses file in same folder
    python generate_oil_dashboard.py path/to/file.xlsx   # explicit path
    python generate_oil_dashboard.py --output my.html    # custom output name

Requirements:
    pip install openpyxl
"""

import os
import sys
import json
import math
import base64
import argparse
from pathlib import Path
from datetime import datetime, date, timezone

try:
    import openpyxl
except ImportError:
    print("Missing dependency. Run:  pip install openpyxl")
    sys.exit(1)

# ---------------------------------------------------------------------------
# CONFIG
# ---------------------------------------------------------------------------

# Controlled via --local CLI flag (see main()). Do not edit here.

DEFAULT_XLSX = "Weekly_Oil_Bulletin_Prices_History_maticni_4web.xlsx"

# These URLs are stable — the EC overwrites the same file in-place every Thursday.
# Verified: same UUID appears in official newsletter emails from Apr 2024 → Aug 2025.
HISTORY_URL = (
    "https://energy.ec.europa.eu/document/download/"
    "906e60ca-8b6a-44e7-8589-652854d2fd3f_en"
    "?filename=Weekly_Oil_Bulletin_Prices_History_maticni_4web.xlsx"
)
DUTIES_URL = (
    "https://energy.ec.europa.eu/document/download/"
    "ccdc6e96-6792-40cb-b0b4-b6609f1e30d0_en"
    "?filename=Oil_Bulletin_Duties_and_taxes.xlsx"
)


def download_file(url: str, dest: Path, label: str = "file") -> bool:
    """Download url to dest. Returns True on success."""
    try:
        import urllib.request
        print(f"  Downloading {label} ...")
        headers = {"User-Agent": "Mozilla/5.0 (compatible; oil-bulletin-dashboard/1.0)"}
        req = urllib.request.Request(url, headers=headers)
        with urllib.request.urlopen(req, timeout=60) as resp, open(dest, "wb") as fh:
            downloaded = 0
            while True:
                block = resp.read(65536)
                if not block:
                    break
                fh.write(block)
                downloaded += len(block)
        print(f"  Saved {downloaded/1024:.0f} KB -> {dest}")
        return True
    except Exception as e:
        print(f"  Download failed: {e}")
        return False


def resolve_xlsx(input_arg: str, force_download: bool, cache_dir: Path, local: bool = False) -> Path:
    """
    Resolution order:
      1. Explicit path that exists on disk  -> use it directly.
      2. --download flag                    -> always re-fetch from EC.
      3. Cached file < 7 days old           -> use cache.
      4. No file found                      -> auto-download.
    """
    # Explicit file path provided
    if input_arg != DEFAULT_XLSX:
        p = Path(input_arg)
        if p.exists():
            return p
        print(f"ERROR: File not found: {p}")
        sys.exit(1)

    cache_path = cache_dir / DEFAULT_XLSX
    cache_dir.mkdir(parents=True, exist_ok=True)

    if local:
        if cache_path.exists():
            print(f"  [LOCAL] Skipping download — using cached file -> {cache_path}")
            return cache_path
        print("ERROR: --local set but no cached xlsx found. Run once without --local.")
        sys.exit(1)

    if force_download:
        print("Force-downloading latest file from EC ...")
        if download_file(HISTORY_URL, cache_path, "Oil Bulletin history"):
            return cache_path
        sys.exit(1)

    # Check cache freshness (7 days)
    if cache_path.exists():
        age_days = (datetime.now() - datetime.fromtimestamp(cache_path.stat().st_mtime)).days
        if age_days < 7:
            print(f"  Using cached file (age: {age_days}d) -> {cache_path}")
            return cache_path
        else:
            print(f"  Cached file is {age_days} days old -- refreshing ...")

    # Auto-download
    if download_file(HISTORY_URL, cache_path, "Oil Bulletin history"):
        return cache_path

    # Fallback to stale cache if download failed
    if cache_path.exists():
        print("  WARNING: Download failed -- using stale cache.")
        return cache_path

    print("ERROR: No local file and download failed. Check your internet connection.")
    sys.exit(1)


BRENT_CACHE_CSV = Path(__file__).parent / "brent_cache.csv"


def _save_brent_cache(date_strs: list, aligned: list, brent_latest: dict | None):
    """Save weekly Brent prices + latest daily quote to CSV cache."""
    import csv
    with open(BRENT_CACHE_CSV, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["date", "price"])
        for ds, price in zip(date_strs, aligned):
            w.writerow([ds, price if price is not None else ""])
        if brent_latest:
            w.writerow([f"latest:{brent_latest['date']}", brent_latest["price"]])


def _load_brent_cache(date_strs: list) -> tuple:
    """Load Brent prices from CSV cache. Returns (aligned, brent_ytd, brent_latest)."""
    import csv
    if not BRENT_CACHE_CSV.exists():
        return [None] * len(date_strs), None, None
    brent_map = {}
    brent_latest = None
    with open(BRENT_CACHE_CSV, newline="", encoding="utf-8") as f:
        for row in csv.reader(f):
            if len(row) < 2 or row[0] == "date":
                continue
            key, val = row[0], row[1]
            if not val:
                continue
            price = float(val)
            if key.startswith("latest:"):
                brent_latest = {"price": price, "date": key[7:]}
            else:
                brent_map[key] = price
    aligned = [brent_map.get(ds) for ds in date_strs]
    latest_yr = int(date_strs[-1][:4])
    jan1_str  = f"{latest_yr}-01-01"
    jan_idx   = next((i for i, ds in enumerate(date_strs) if ds >= jan1_str), 0)
    base = next((v for v in aligned[jan_idx:] if v), None)
    last = next((v for v in reversed(aligned) if v), None)
    brent_ytd = round((last / base - 1) * 100, 2) if base and last else None
    found = sum(1 for v in aligned if v is not None)
    print(f"  Brent (cache): {found}/{len(date_strs)} weeks loaded")
    return aligned, brent_ytd, brent_latest


def fetch_brent(date_strs: list, local: bool = False) -> tuple:
    """
    Download weekly Brent crude (BZ=F) from Yahoo Finance.
    Saves results to brent_cache.csv; falls back to cache on failure.
    Returns (aligned_prices, brent_ytd_pct, brent_latest).
    """
    if local:
        print("  [LOCAL] Skipping Yahoo Finance — loading Brent from cache ...")
        return _load_brent_cache(date_strs)

    import urllib.request, json
    from datetime import datetime, timedelta

    HEADERS = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
        "Accept": "application/json",
    }

    def yf_get(path: str) -> dict:
        """Try query1 then query2 Yahoo Finance hosts."""
        for host in ("query1.finance.yahoo.com", "query2.finance.yahoo.com"):
            url = f"https://{host}{path}"
            try:
                req = urllib.request.Request(url, headers=HEADERS)
                with urllib.request.urlopen(req, timeout=30) as r:
                    return json.loads(r.read())
            except Exception as e:
                print(f"  {host} failed: {e}")
        raise RuntimeError(f"Both Yahoo Finance hosts failed for {path}")

    try:
        start_dt = datetime.strptime(date_strs[0], "%Y-%m-%d") - timedelta(days=14)
        end_ts   = int(datetime.now().timestamp())
        start_ts = int(start_dt.timestamp())
        path = f"/v8/finance/chart/BZ=F?period1={start_ts}&period2={end_ts}&interval=1wk"
        raw = yf_get(path)
        if raw.get("chart", {}).get("error"):
            raise RuntimeError(f"Yahoo API error: {raw['chart']['error']}")
        result  = raw["chart"]["result"][0]
        tss     = result["timestamp"]
        closes  = result["indicators"]["quote"][0]["close"]

        # Build lookup with ±4-day window so Monday EC dates match Friday Yahoo bars
        brent_map = {}
        for ts, price in zip(tss, closes):
            if price is None:
                continue
            d = datetime.fromtimestamp(ts, tz=timezone.utc).replace(tzinfo=None)
            for offset in range(-4, 5):
                key = (d + timedelta(days=offset)).strftime("%Y-%m-%d")
                brent_map.setdefault(key, round(price, 2))

        aligned = [brent_map.get(ds) for ds in date_strs]
        found   = sum(1 for v in aligned if v is not None)
        print(f"  Brent (Yahoo BZ=F): {found}/{len(date_strs)} weeks matched")

        # YTD %
        latest_yr = int(date_strs[-1][:4])
        jan1_str  = f"{latest_yr}-01-01"
        jan_idx   = next((i for i, ds in enumerate(date_strs) if ds >= jan1_str), 0)
        base = next((v for v in aligned[jan_idx:] if v), None)
        last = next((v for v in reversed(aligned) if v), None)
        brent_ytd = round((last / base - 1) * 100, 2) if base and last else None

        # Latest daily quote
        brent_latest = None
        daily_raw = yf_get("/v8/finance/chart/BZ=F?range=5d&interval=1d")
        d_res    = daily_raw["chart"]["result"][0]
        d_tss    = d_res["timestamp"]
        d_closes = d_res["indicators"]["quote"][0]["close"]
        for ts2, price2 in zip(reversed(d_tss), reversed(d_closes)):
            if price2 is not None:
                d2 = datetime.fromtimestamp(ts2, tz=timezone.utc).strftime("%Y-%m-%d")
                brent_latest = {"price": round(price2, 2), "date": d2}
                print(f"  Brent latest daily: ${price2:.2f} ({d2})")
                break

        _save_brent_cache(date_strs, aligned, brent_latest)
        return aligned, brent_ytd, brent_latest

    except Exception as e:
        print(f"  WARNING: Brent download failed: {e} — using cache")
        return _load_brent_cache(date_strs)


COUNTRIES    = ["EU", "FR", "DE", "NL", "ES", "IT", "PT"]
COUNTRIES    = ["FR", "DE", "NL", "ES", "IT", "PT"]
WEEKS_BACK   = 522          # ≈ 10 years of weekly data

# Colours
COLORS = {
    "EU": "#f59e0b", "FR": "#3b82f6", "DE": "#10b981",
    "NL": "#f97316", "ES": "#ec4899", "IT": "#8b5cf6", "PT": "#06b6d4",
}
FUEL_COLORS = {
    "Gasoline":    "#f59e0b",
    "Diesel":      "#3b82f6",
    "Heating Oil": "#10b981",
    "Fuel Oil":    "#a1a1aa",
    "LPG":         "#f97316",
}

# ---------------------------------------------------------------------------
# HELPERS
# ---------------------------------------------------------------------------
def safe_float(v):
    try:
        return float(v) if v is not None else None
    except (TypeError, ValueError):
        return None


def col_indices(headers, country, sheet_type="with_tax"):
    """Return dict {product: col_index} for a given country in a price sheet."""
    prefix_map = {
        "with_tax": f"{country}_price_with_tax_",
        "wo_tax":   f"{country}_price_wo_tax_",
        "cons":     f"{country}_consumption_",
    }
    prefix = prefix_map[sheet_type]
    result = {}
    for i, h in enumerate(headers):
        if h and str(h).startswith(prefix):
            # Normalise heating oil column name variants
            product = str(h)[len(prefix):]
            product = product.replace(f"he{country}ing_oil", "heating_oil")
            result[product] = i
    return result


def fmt_date(d):
    """datetime → 'Mon YY' label for chart axis."""
    months = ["","Jan","Feb","Mar","Apr","May","Jun",
              "Jul","Aug","Sep","Oct","Nov","Dec"]
    return f"{months[d.month]} {str(d.year)[2:]}"


def compute_sensitivity(brent: list, fuel: list, window: int = 4, lag: int = 1):
    """OLS slope through origin, split by direction of Brent move.

    Unit: millieuros/L per $/bbl = eurocents/L per $10/bbl (chart display unit).
    Returns (up_slope, down_slope).
    """
    up_xx = up_xy = dn_xx = dn_xy = 0.0
    for t in range(window + lag, len(brent)):
        bx  = brent[t - lag]
        bx0 = brent[t - lag - window]
        fy  = fuel[t]
        fy0 = fuel[t - window]
        if bx is None or bx0 is None or fy is None or fy0 is None:
            continue
        dx = bx - bx0   # $/bbl
        dy = fy  - fy0  # millieuros/L
        if   dx > 0: up_xx += dx * dx; up_xy += dx * dy
        elif dx < 0: dn_xx += dx * dx; dn_xy += dx * dy
    return (
        round(up_xy / up_xx, 2) if up_xx else 0.0,
        round(dn_xy / dn_xx, 2) if dn_xx else 0.0,
    )


# ---------------------------------------------------------------------------
# DATA EXTRACTION
# ---------------------------------------------------------------------------
def extract_data(xlsx_path: Path, local: bool = False) -> dict:
    print(f"Reading: {xlsx_path}")
    wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)

    # ---- PRICES WITH TAX ------------------------------------------------
    ws_tax = wb["Prices with taxes"]
    tax_rows = list(ws_tax.iter_rows(values_only=True))
    tax_headers = tax_rows[0]

    # Data rows start at index 3 (index 0=col names, 1=labels, 2=units, 3+=data)
    data_rows = tax_rows[3:]
    # Filter to rows that have a valid date
    data_rows = [(r[0], r) for r in data_rows if isinstance(r[0], datetime)]
    # Sort oldest→newest
    data_rows.sort(key=lambda x: x[0])
    # Take last WEEKS_BACK
    data_rows = data_rows[-WEEKS_BACK:]

    dates       = [r[0] for _, r in data_rows]
    date_labels = [fmt_date(d) for d in dates]
    date_strs   = [d.strftime("%Y-%m-%d") for d in dates]

    # Column maps for each country
    col_tax = {c: col_indices(tax_headers, c, "with_tax") for c in COUNTRIES}

    # ---- PRICES WITHOUT TAX ----------------------------------------------
    ws_notax = wb["Prices wo taxes"]
    notax_rows_raw = list(ws_notax.iter_rows(values_only=True))
    notax_headers  = notax_rows_raw[0]
    notax_by_date  = {}
    for r in notax_rows_raw[3:]:
        if isinstance(r[0], datetime):
            notax_by_date[r[0].strftime("%Y-%m-%d")] = r

    col_notax = {c: col_indices(notax_headers, c, "wo_tax") for c in COUNTRIES}

    # ---- BUILD PRICE SERIES ----------------------------------------------
    countries_data = {}
    for c in COUNTRIES:
        euro95, diesel, euro95_nt, diesel_nt = [], [], [], []
        for d, r in data_rows:
            ds = d.strftime("%Y-%m-%d")
            # With tax
            v95  = safe_float(r[col_tax[c].get("euro95")])   if col_tax[c].get("euro95")  is not None else None
            vdie = safe_float(r[col_tax[c].get("diesel")])   if col_tax[c].get("diesel")  is not None else None
            euro95.append(round(v95,  2) if v95  is not None else None)
            diesel.append(round(vdie, 2) if vdie is not None else None)
            # Without tax
            nr = notax_by_date.get(ds)
            if nr:
                vnt95  = safe_float(nr[col_notax[c].get("euro95")])  if col_notax[c].get("euro95")  is not None else None
                vntdie = safe_float(nr[col_notax[c].get("diesel")])  if col_notax[c].get("diesel")  is not None else None
                euro95_nt.append(round(vnt95,  2) if vnt95  is not None else None)
                diesel_nt.append(round(vntdie, 2) if vntdie is not None else None)
            else:
                euro95_nt.append(None)
                diesel_nt.append(None)

        # Tax rate estimate from latest non-null pair
        tax_rate = None
        for i in range(len(euro95)-1, -1, -1):
            if euro95[i] and euro95_nt[i] and euro95[i] > 0:
                tax_rate = round(1 - euro95_nt[i] / euro95[i], 3)
                break

        countries_data[c] = {
            "euro95":       euro95,
            "diesel":       diesel,
            "euro95_notax": euro95_nt,
            "diesel_notax": diesel_nt,
            "tax_rate":     tax_rate or 0.55,
        }

    # ---- YTD % CHANGE ---------------------------------------------------
    ytd = {}
    jan1 = date(dates[-1].year, 1, 1)
    # Find first date >= Jan 1 of latest year
    jan_idx = next((i for i, d in enumerate(dates) if d.date() >= jan1), 0)

    for c in COUNTRIES:
        s95  = countries_data[c]["euro95"]
        sdie = countries_data[c]["diesel"]
        def ytd_pct(series):
            base = next((v for v in series[jan_idx:] if v), None)
            last = next((v for v in reversed(series) if v), None)
            if base and last and base > 0:
                return round((last / base - 1) * 100, 2)
            return None
        def ytd_abs(series):
            base = next((v for v in series[jan_idx:] if v), None)
            last = next((v for v in reversed(series) if v), None)
            if base and last:
                return round((last - base) / 1000, 4)  # €/L
            return None
        ytd[c] = {
            "euro95_ytd": ytd_pct(s95), "diesel_ytd": ytd_pct(sdie),
            "euro95_abs": ytd_abs(s95), "diesel_abs": ytd_abs(sdie),
        }

    # ---- CONSUMPTION ----------------------------------------------------
    ws_cons  = wb["Consumption"]
    cons_rows = list(ws_cons.iter_rows(values_only=True))
    cons_headers = cons_rows[0]

    cons_ctrs = [c for c in COUNTRIES if c != "EU"]
    col_cons  = {c: col_indices(cons_headers, c, "cons") for c in cons_ctrs}

    # Latest year row (row index 2 after skipping headers)
    latest_cons_row = cons_rows[2]  # most recent year

    consumption = {}
    for c in cons_ctrs:
        vals = {}
        for prod, idx in col_cons[c].items():
            v = safe_float(latest_cons_row[idx])
            vals[prod] = round(v, 1) if v else 0.0
        # Rename for display
        consumption[c] = {
            "Gasoline":    vals.get("euro95",      0),
            "Diesel":      vals.get("diesel",       0),
            "Heating Oil": vals.get("heating_oil",  0),
            "Fuel Oil":    vals.get("fuel_oil_1",   0) + vals.get("fuel_oil_2", 0),
            "LPG":         vals.get("LPG",          0),
        }

    latest_year = latest_cons_row[0] if latest_cons_row[0] else "latest"

    brent_prices, brent_ytd, brent_latest = fetch_brent(date_strs, local=local)

    # ---- SENSITIVITY (OLS, computed once here rather than in JS) -----------
    _SENS_WINDOWS = [4, 26]
    _SENS_LAGS    = [0, 1, 2]
    sensitivity   = {}
    for fuel_key in ("euro95", "diesel"):
        # Per-country slopes (default window=4, lag=1) for the bar charts
        per_country = {}
        for c in COUNTRIES:
            up, dn = compute_sensitivity(brent_prices, countries_data[c][fuel_key])
            per_country[c] = {"up": up, "down": dn}
        # Research grid: average slope across all countries for each (window, lag)
        research = {}
        for win in _SENS_WINDOWS:
            for lag in _SENS_LAGS:
                ups, dns = [], []
                for c in COUNTRIES:
                    u, d = compute_sensitivity(
                        brent_prices, countries_data[c][fuel_key], win, lag
                    )
                    ups.append(u); dns.append(d)
                research[f"{win}_{lag}"] = {
                    "up":   round(sum(ups) / len(ups), 2),
                    "down": round(sum(dns) / len(dns), 2),
                }
        sensitivity[fuel_key] = {"per_country": per_country, "research": research}

    print(f"  Dates: {date_strs[0]} to {date_strs[-1]} ({len(date_strs)} weeks)")
    print(f"  YTD ({dates[-1].year}): {ytd}")
    print(f"  Consumption year: {latest_year}")

    return {
        "dates":       date_strs,
        "labels":      date_labels,
        "countries":   countries_data,
        "ytd":         ytd,
        "consumption": consumption,
        "latest_year": str(latest_year),
        "latest_date": date_strs[-1],
        "brent":        brent_prices,
        "brent_ytd":    brent_ytd,
        "brent_latest": brent_latest,
        "sensitivity":  sensitivity,
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
    }


# ---------------------------------------------------------------------------
# HTML GENERATION
# ---------------------------------------------------------------------------
CHART_JS_CDN = "https://cdn.jsdelivr.net/npm/chart.js@4.4.2/dist/chart.umd.min.js"


def build_html(data: dict) -> str:
    colors_js      = json.dumps(COLORS)
    fuel_colors_js = json.dumps(FUEL_COLORS)
    countries_js   = json.dumps(COUNTRIES)
    data_js        = json.dumps(data)

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>EU Weekly Oil Bulletin Dashboard</title>
<link rel="icon" href="data:image/svg+xml,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 100 100'><text y='.9em' font-size='90'>🛢️</text></svg>">
<script src="{CHART_JS_CDN}"></script>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700;800&family=DM+Mono:wght@400;500&display=swap" rel="stylesheet">
<style>
* {{ box-sizing: border-box; margin: 0; padding: 0; }}
body {{
  background: #020818;
  color: #e2e8f0;
  font-family: 'DM Sans', sans-serif;
  min-height: 100vh;
}}
.mono {{ font-family: 'DM Mono', monospace; }}
/* HEADER */
.header {{
  background: linear-gradient(135deg,#0c1428 0%,#0f2040 100%);
  border-bottom: 1px solid #1e3a5f;
  padding: 20px 32px 0;
}}
.header-top {{
  display: flex; align-items: flex-start;
  justify-content: space-between; margin-bottom: 16px; flex-wrap: wrap; gap: 12px;
}}
.header-title {{ display: flex; align-items: center; gap: 10px; }}
.logo {{
  width: 36px; height: 36px;
  background: linear-gradient(135deg,#f59e0b,#d97706);
  border-radius: 8px; display: flex; align-items: center;
  justify-content: center; font-size: 18px;
}}
h1 {{ font-size: 22px; font-weight: 800; letter-spacing: -0.5px; color: #f1f5f9; }}
.subtitle {{ font-size: 12px; color: #94a3b8; margin-top: 3px; }}
.price-badges {{ display: flex; gap: 8px; flex-wrap: wrap; justify-content: flex-end; }}
.badge {{
  background: #0c1a30; border-radius: 8px;
  padding: 6px 12px; text-align: center; min-width: 68px;
}}
.badge-label {{ font-size: 10px; font-weight: 700; margin-bottom: 2px; }}
.badge-price {{ font-size: 13px; font-weight: 800; color: #f1f5f9; }}
.badge-unit {{ font-size: 9px; color: #94a3b8; margin-top: 1px; }}
/* TABS */
.tabs {{ display: flex; }}
.tab-btn {{
  background: none; border: none; cursor: pointer;
  padding: 10px 20px; font-size: 13px; font-weight: 600;
  color: #94a3b8; border-bottom: 2px solid transparent;
  transition: all 0.2s; font-family: 'DM Sans', sans-serif;
}}
.tab-btn.active {{ color: #f59e0b; border-bottom-color: #f59e0b; }}
/* PANELS */
.content {{ padding: 24px 32px; max-width: 1100px; margin: 0 auto; }}
.panel {{ display: none; }}
.panel.active {{ display: block; }}
/* CARDS / BOXES */
.card {{
  background: #060e1e; border-radius: 12px;
  border: 1px solid #1e293b; overflow: hidden;
}}
.card-header {{
  padding: 14px 20px; border-bottom: 1px solid #1e293b;
  display: flex; justify-content: space-between; align-items: center;
}}
.card-title {{ font-size: 13px; font-weight: 700; color: #f1f5f9; }}
.card-sub {{ font-size: 11px; color: #94a3b8; }}
/* GRID */
.grid-2 {{ display: grid; grid-template-columns: 1fr 1fr; gap: 20px; }}
.grid-7 {{ display: grid; grid-template-columns: repeat(auto-fill, minmax(110px, 1fr)); gap: 10px; margin-bottom: 20px; }}
/* YTD cards */
.ytd-card {{
  background: #060e1e; border-radius: 10px;
  border: 1px solid #1e3a5f; padding: 12px;
}}
.ytd-ctr {{ font-size: 11px; font-weight: 700; margin-bottom: 8px; }}
.ytd-val {{ font-size: 13px; font-weight: 800; line-height: 1.3; }}
.ytd-sub {{ font-size: 9px; color: #94a3b8; margin-bottom: 2px; }}
.ytd-val2 {{ font-size: 12px; font-weight: 700; margin-top: 4px; line-height: 1.3; }}
.up {{ color: #ef4444; }}
.dn {{ color: #10b981; }}
/* TABLE */
table {{ width: 100%; border-collapse: collapse; font-size: 12px; }}
th {{
  padding: 8px 16px; text-align: right; color: #94a3b8;
  font-weight: 600; border-bottom: 1px solid #1e293b;
  background: #0c1428;
}}
th:first-child {{ text-align: left; }}
td {{ padding: 8px 16px; text-align: right; }}
td:first-child {{ text-align: left; }}
tr:nth-child(even) {{ background: #080f1a; }}
.dot {{
  display: inline-block; width: 8px; height: 8px;
  border-radius: 50%; margin-right: 8px;
}}
.pill {{
  background: #f59e0b20; color: #f59e0b;
  padding: 2px 8px; border-radius: 12px;
  font-size: 11px;
}}
/* FUEL toggle */
.toggle-row {{ display: flex; gap: 8px; }}
.toggle-btn {{
  padding: 6px 14px; border-radius: 20px; font-size: 12px;
  font-weight: 600; border: 1px solid #1e3a5f;
  background: #0c1428; color: #94a3b8;
  cursor: pointer; font-family: 'DM Sans', sans-serif;
  transition: all 0.2s;
}}
.toggle-btn.active {{
  border-color: #f59e0b; background: rgba(245,158,11,0.1); color: #f59e0b;
}}
/* TAX bars */
.tax-bar-row {{
  display: flex; align-items: center; gap: 12px; margin-bottom: 10px;
}}
.tax-bar-label {{ width: 32px; font-size: 12px; font-weight: 700; }}
.tax-bar-track {{
  flex: 1; background: #0c1428; border-radius: 4px; height: 26px; overflow: hidden;
}}
.tax-bar-fill {{
  height: 100%; border-radius: 4px;
  display: flex; align-items: center; justify-content: flex-end; padding-right: 10px;
}}
.tax-bar-pct {{ font-size: 11px; font-weight: 700; color: #fff; }}
.tax-bar-aside {{ width: 100px; font-size: 11px; color: #94a3b8; text-align: right; }}
/* Chart containers */
.chart-wrap {{ padding: 16px 12px 8px; }}
canvas {{ max-width: 100%; }}
/* Info box */
.info-box {{
  margin-top: 12px; padding: 12px 16px;
  background: #0c1428; border-radius: 8px;
  font-size: 11px; color: #94a3b8;
}}
.section-title {{
  font-size: 16px; font-weight: 700; color: #f1f5f9; margin-bottom: 4px;
}}
.section-sub {{ font-size: 12px; color: #94a3b8; margin-bottom: 20px; }}
/* Responsive */
@media (max-width: 900px) {{
  .grid-2 {{ grid-template-columns: 1fr; }}
  .grid-7 {{ grid-template-columns: repeat(4,1fr); }}
  .price-badges {{ display: none; }}
  .content {{ padding: 16px; }}
  .header {{ padding: 16px 16px 0; }}
}}
</style>
</head>
<body>

<div class="header">
  <div class="header-top">
    <div>
      <div class="header-title">
        <div class="logo">🛢️</div>
        <div>
          <h1>EU Weekly Oil Bulletin</h1>
          <div class="subtitle">European Commission · Consumer Petroleum Prices inc. taxes · Latest: <span id="latest-date"></span></div>
          <div class="subtitle"><a href="https://energy.ec.europa.eu/data-and-analysis/weekly-oil-bulletin_en" target="_blank" rel="noopener" style="color:#f59e0b;text-decoration:none;font-weight:600;" onmouseover="this.style.color='#fde68a'" onmouseout="this.style.color='#f59e0b'">Data source ↗</a></div>
        </div>
      </div>
    </div>
    <div style="text-align:right;">
      <div style="display:flex;align-items:center;gap:10px;justify-content:flex-end;">
        <div id="badge-date-label" style="font-size:11px;color:#94a3b8;white-space:nowrap;line-height:1.4;"></div>
        <div class="price-badges" id="price-badges"></div>
      </div>
      <div id="brent-info"  style="font-size:10px;color:#94a3b8;margin-top:6px;"></div>
      <div id="brent-info2" style="font-size:10px;color:#94a3b8;margin-top:2px;"></div>
    </div>
  </div>
  <div class="tabs">
    <button class="tab-btn active" onclick="showTab(0)">Historical Prices</button>
    <button class="tab-btn" onclick="showTab(1)">YTD Performance</button>
    <button class="tab-btn" onclick="showTab(2)">Tax Analysis</button>
    <button class="tab-btn" onclick="showTab(3)">Consumption</button>
    <button class="tab-btn" onclick="showTab(4)">Sensitivity</button>
  </div>
</div>

<div class="content">

  <!-- TAB 0: Historical -->
  <div class="panel active" id="tab0">
    <div style="margin-bottom:20px;">
      <div class="section-title" style="text-align:center;">Pump Price History (EUR/L)</div>
      <div class="section-sub" style="text-align:center;margin-bottom:14px;">Weekly consumer pump prices inclusive of taxes and duties</div>
      <div style="display:flex;flex-direction:column;gap:8px;align-items:center;">
        <div class="toggle-row">
          <button class="toggle-btn" id="btn95" onclick="switchFuel('euro95')">Euro-95</button>
          <button class="toggle-btn active" id="btnD" onclick="switchFuel('diesel')">Diesel</button>
        </div>
        <div class="toggle-row">
          <button class="toggle-btn" id="btn1Y" onclick="setRange(52)">1Y</button>
          <button class="toggle-btn" id="btn3Y" onclick="setRange(156)">3Y</button>
          <button class="toggle-btn" id="btn5Y" onclick="setRange(260)">5Y</button>
          <button class="toggle-btn active" id="btnAll" onclick="setRange(0)">ALL</button>
        </div>
      </div>
    </div>
    <div class="card">
      <div class="chart-wrap" style="height:400px;">
        <canvas id="histChart"></canvas>
      </div>
    </div>
    <div class="card" style="margin-top:20px;">
      <div class="card-header">
        <span class="card-title">Latest Prices — <span id="table-date"></span></span>
        <span class="card-sub">€ / litre</span>
      </div>
      <div class="chart-wrap" style="height:300px;">
        <canvas id="priceChart"></canvas>
      </div>
    </div>
  </div>

  <!-- TAB 1: YTD -->
  <div class="panel" id="tab1">
    <div class="section-title">YTD Performance <span id="ytd-year"></span></div>
    <div class="section-sub">% change from Jan 1 through latest data point</div>
    <div class="grid-7" id="ytd-cards"></div>
    <div class="card">
      <div class="chart-wrap" style="height:360px;">
        <canvas id="ytdChart"></canvas>
      </div>
    </div>
  </div>

  <!-- TAB 2: Tax Analysis -->
  <div class="panel" id="tab2">
    <div class="section-title">Tax & Duty Breakdown</div>
    <div class="section-sub">Pre-tax product cost vs total duty burden · EUR/litre · Latest data point</div>
    <div class="grid-2">
      <div class="card">
        <div class="card-header"><span class="card-title">Diesel — Pre-tax vs Taxes (€/L)</span></div>
        <div class="chart-wrap" style="height:280px;"><canvas id="taxDChart"></canvas></div>
        <div style="padding:0 16px 16px;" id="taxDTable"></div>
      </div>
      <div class="card">
        <div class="card-header"><span class="card-title">Euro-95 — Pre-tax vs Taxes (€/L)</span></div>
        <div class="chart-wrap" style="height:280px;"><canvas id="tax95Chart"></canvas></div>
        <div style="padding:0 16px 16px;" id="tax95Table"></div>
      </div>
    </div>
    <div class="card" style="margin-top:20px;">
      <div class="card-header"><span class="card-title">Effective Tax Rate on Euro-95 (ranked)</span></div>
      <div style="padding:16px 20px;" id="tax-bars"></div>
    </div>
  </div>

  <!-- TAB 3: Consumption -->
  <div class="panel" id="tab3">
    <div style="margin-bottom:20px;">
      <div class="section-title" style="text-align:center;">Petroleum Consumption — <span id="cons-year"></span></div>
      <div class="section-sub" style="text-align:center;margin-bottom:14px;">Source: EC Weekly Oil Bulletin · kt (1,000 tonnes)</div>
      <div style="display:flex;justify-content:center;">
        <div class="toggle-row">
          <button class="toggle-btn active" id="btnAbsolute" onclick="switchCons('absolute')">By Volume (kt)</button>
          <button class="toggle-btn" id="btnMix" onclick="switchCons('mix')">Consumption Mix %</button>
        </div>
      </div>
    </div>
    <!-- Sub-panel: absolute volumes -->
    <div id="cons-absolute">
      <div class="card">
        <div class="chart-wrap" style="height:400px;"><canvas id="consAbsChart"></canvas></div>
      </div>
    </div>
    <!-- Sub-panel: percentage mix -->
    <div id="cons-mix" style="display:none;">
      <div class="card">
        <div class="chart-wrap" style="height:400px;"><canvas id="consMixChart"></canvas></div>
      </div>
    </div>
    <!-- Shared table -->
    <div class="card" style="margin-top:20px;">
      <div class="card-header"><span class="card-title">Absolute Volumes (kt)</span></div>
      <table>
        <thead><tr id="cons-head"></tr></thead>
        <tbody id="cons-body"></tbody>
      </table>
    </div>
  </div>

  <!-- TAB 4: Sensitivity -->
  <div class="panel" id="tab4">
    <div style="margin-bottom:20px;">
      <div class="section-title" style="text-align:center;">Brent → Pump Price Sensitivity</div>
      <div style="text-align:center;font-size:20px;font-weight:300;color:#f59e0b;margin:6px 0 2px;">
        Suppliers are <span style="text-decoration:underline;">quick</span> to pass on crude price increases,<br>but <span style="text-decoration:underline;">slow</span> to adjust downward
      </div>
    </div>
    <div class="section-sub" style="text-align:center;margin-bottom:0;">
        Pump price rise / decline (€ cents/L) per $10 Brent move
    </div>
  
    <div class="grid-2">
      <div>
        <div class="card">
          <div class="card-header">
            <span class="card-title">Diesel</span>
            <span class="card-sub" id="sens-diesel-slope"></span>
          </div>
          <div class="chart-wrap" style="height:340px;"><canvas id="sensDieselChart"></canvas></div>
        </div>
        <div class="card" style="margin-top:12px;" id="sens-diesel-table"></div>
      </div>
      <div>
        <div class="card">
          <div class="card-header">
            <span class="card-title">Euro-95</span>
            <span class="card-sub" id="sens-95-slope"></span>
          </div>
          <div class="chart-wrap" style="height:340px;"><canvas id="sens95Chart"></canvas></div>
        </div>
        <div class="card" style="margin-top:12px;" id="sens-95-table"></div>
      </div>
    </div>
    <div class="info-box" style="margin-top:16px;">
      <strong style="color:#f1f5f9;">How to read:</strong> Each bar = OLS slope of 4-week fuel change vs 4-week Brent change (lagged 1 week),
      split by direction of Brent move. <strong style="color:#f59e0b;">Solid bar</strong> = weeks when Brent rose ·
      <strong style="color:#94a3b8;">Light bar</strong> = weeks when Brent fell.
      A taller solid bar than light bar confirms the well-known asymmetry:
      <em>suppliers are quick to pass on crude price increases but slow to adjust downward</em>.
    </div>
    <div style="margin-top:24px;">
      <div class="section-title" style="text-align:center;margin-bottom:4px;">Window &amp; Lag Research</div>
      <div style="text-align:center;font-size:12px;color:#94a3b8;margin-bottom:14px;">
        Pump price rise / decline (€ cents/L) per $10 Brent move &nbsp;·&nbsp; avg across all countries &nbsp;·&nbsp;
        <strong style="color:#f59e0b;">★ = current model choice (4W, lag 1W)</strong>
      </div>
      <div class="grid-2">
        <div class="card">
          <div class="card-header"><span class="card-title">Diesel</span></div>
          <div id="sens-research-diesel"></div>
        </div>
        <div class="card">
          <div class="card-header"><span class="card-title">Euro-95</span></div>
          <div id="sens-research-95"></div>
        </div>
      </div>
      <div class="info-box" style="margin-top:16px;font-size:12px;line-height:1.7;">
        <strong style="color:#f1f5f9;display:block;margin-bottom:6px;">Why 4-week window &amp; 1-week lag?</strong>
        <ul style="margin:0;padding-left:18px;color:#94a3b8;">
          <li>
            <strong style="color:#cbd5e1;">4-week window (short-run pass-through).</strong>
            Borenstein, Cameron &amp; Gilbert (1997, <em>QJE</em>) establish the 4-week horizon as
            the standard for measuring asymmetric retail-fuel price adjustment — long enough to
            capture a full pricing cycle, short enough to isolate behavioural response from
            seasonal and demand-side noise. Confirmed in the ECB's cross-country replication
            (Gelos &amp; Ustyugova, 2017) and in Aucremanne &amp; Dhyne (2005) for euro-area retail prices.
          </li>
          <li style="margin-top:6px;">
            <strong style="color:#cbd5e1;">1-week lag on Brent.</strong>
            EC prices are collected on Mondays and reflect purchasing decisions from the prior
            week. A 1-week lag aligns the crude signal with the actual procurement window, as
            recommended by Granger &amp; Lee (1989, <em>Oxford Bulletin of Econ &amp; Stats</em>) for
            error-correction models with weekly data, and applied in Bacon (1991,
            <em>Oxford Energy Studies</em>) to EU petrol markets.
          </li>
          <li style="margin-top:6px;">
            <strong style="color:#cbd5e1;">Why not 26W?</strong>
            Longer windows accumulate multiple Brent cycles; the growing gap above reflects
            compounded short-run asymmetry, not a stronger behavioural effect. With fewer
            independent observations and greater exposure to structural breaks (COVID, Ukraine),
            the 26W slope is less reliable for identifying pricing behaviour
            (cf. Peltzman 2000, <em>Journal of Political Economy</em>).
          </li>
        </ul>
      </div>
    </div>
  </div>

</div>

<div style="text-align:center;padding:14px 32px;font-size:10px;color:#94a3b8;border-top:1px solid #1e293b;">
  Generated <span id="gen-datetime"></span> · Source: European Commission Weekly Oil Bulletin &amp; Yahoo Finance
</div>

<script>
// ---- EMBEDDED DATA -------------------------------------------------------
const DATA = {data_js};
const COLORS = {colors_js};
const FUEL_COLORS = {fuel_colors_js};
const CTRS = {countries_js};
const FUEL_TYPES = ["Gasoline","Diesel","Heating Oil","Fuel Oil","LPG"];

// ---- UTILS ---------------------------------------------------------------
const $ = id => document.getElementById(id);
let histChart, ytdChart, tax95Chart, taxDChart, consAbsChart, consMixChart;
let currentFuel  = 'diesel';
let currentRange = 0;          // 52 = 1Y; 156 = 3Y; 260 = 5Y; 0 = ALL
let currentCons  = 'absolute';

function fmtVal(v, decimals=1) {{
  if (v == null) return '—';
  return v.toLocaleString('en', {{minimumFractionDigits: decimals, maximumFractionDigits: decimals}});
}}

// ---- INIT ----------------------------------------------------------------
document.addEventListener('DOMContentLoaded', () => {{
  const latest = DATA.dates[DATA.dates.length - 1];
  $('latest-date').textContent = latest;
  $('table-date').textContent  = latest;

  buildBadges();
  buildHistChart();
  buildPriceChart();
  buildYTD();
  buildTaxCharts();
  buildConsumption();
  buildSensitivity();
  $('gen-datetime').textContent = DATA.generated_at;
}});

function showTab(n) {{
  document.querySelectorAll('.panel').forEach((p,i) => p.classList.toggle('active', i===n));
  document.querySelectorAll('.tab-btn').forEach((b,i) => b.classList.toggle('active', i===n));
}}

// ---- BADGES (header) -----------------------------------------------------
function fmtDateLabel(iso) {{
  const d = new Date(iso + 'T12:00:00Z');
  return d.toLocaleDateString('en-GB', {{ day: 'numeric', month: 'long', year: 'numeric', timeZone: 'UTC' }});
}}

function brentInfoLine(price, date, diffPrice, diffRef, col) {{
  const sign  = diffPrice >= 0 ? '+' : '';
  const arrow = diffPrice >= 0 ? '▲' : '▼';
  const pct   = (diffPrice / diffRef * 100).toFixed(1);
  return `<span class="mono" style="color:#cbd5e1">$${{price.toFixed(2)}}/bbl</span>`
    + ` <span style="color:#94a3b8">(${{date}})</span>`
    + ` &nbsp;·&nbsp; <span style="color:${{col}}">${{arrow}} ${{sign}}$${{Math.abs(diffPrice).toFixed(2)}} (${{sign}}${{pct}}%)</span>`;
}}

function buildBadges() {{
  const last       = DATA.dates.length - 1;
  const latestDate = DATA.dates[last];
  const wrap       = $('price-badges');

  // Date label to the left of badges
  $('badge-date-label').innerHTML = fmtDateLabel(latestDate).replace(' ', '<br>');

  // Country badges (no per-badge date)
  CTRS.forEach(c => {{
    const v  = DATA.countries[c].euro95[last];
    const el = document.createElement('div');
    el.className = 'badge';
    el.style.border = `1px solid ${{COLORS[c]}}40`;
    el.innerHTML = `
      <div class="badge-label" style="color:${{COLORS[c]}}">${{c}}</div>
      <div class="badge-price mono">€${{v != null ? (v/1000).toFixed(3) : '—'}}</div>
      <div class="badge-unit">Euro-95/L</div>`;
    wrap.appendChild(el);
  }});

  // Brent badge at latest EC date
  const brentVal = DATA.brent[last];
  const bEl = document.createElement('div');
  bEl.className = 'badge';
  bEl.style.border = '1px solid #94a3b840';
  bEl.innerHTML = `
    <div class="badge-label" style="color:#94a3b8">Brent</div>
    <div class="badge-price mono" style="color:#94a3b8">${{brentVal != null ? '$'+brentVal.toFixed(2) : '—'}}</div>
    <div class="badge-unit">$/barrel</div>`;
  wrap.appendChild(bEl);

  // ---- Line 1: Brent at EC date vs previous EC week ----------------------
  let cur = null, prev = null, curDate = latestDate;
  for (let i = DATA.brent.length - 1; i >= 0; i--) {{
    if (DATA.brent[i] != null) {{
      if (cur === null) {{ cur = DATA.brent[i]; curDate = DATA.dates[i]; }}
      else              {{ prev = DATA.brent[i]; break; }}
    }}
  }}
  if (cur != null && prev != null) {{
    const col = (cur - prev) >= 0 ? '#ef4444' : '#10b981';
    $('brent-info').innerHTML =
      `Brent: ` + brentInfoLine(cur, curDate, cur - prev, prev, col)
      + ` <span style="color:#94a3b8">vs prev. week</span>`;
  }}

  // ---- Line 2: latest daily Yahoo quote vs EC-date Brent -----------------
  const bl = DATA.brent_latest;
  if (bl && bl.price != null && cur != null && bl.date !== curDate) {{
    const diff2 = bl.price - cur;
    const col2  = diff2 >= 0 ? '#ef4444' : '#10b981';
    $('brent-info2').innerHTML =
      brentInfoLine(bl.price, bl.date + ' · today', diff2, cur, col2)
      + ` <span style="color:#94a3b8">vs EC date</span>`;
  }}
}}

// ---- HISTORICAL CHART ----------------------------------------------------
function buildHistChart() {{
  const ctx = $('histChart').getContext('2d');
  const datasets = CTRS.map(c => ({{
    label: c,
    data: DATA.countries[c][currentFuel].map(v => v != null ? +(v/1000).toFixed(4) : null),
    borderColor: COLORS[c],
    backgroundColor: 'transparent',
    borderWidth: c === 'EU' ? 2.5 : 1.8,
    pointRadius: 0,
    tension: 0.3,
    spanGaps: true,
    yAxisID: 'y',
  }}));

  // Brent crude — left axis (rose)
  datasets.push({{
    label: 'Brent',
    data: DATA.brent.map(v => v != null ? +v.toFixed(2) : null),
    borderColor: '#fdffbf',
    backgroundColor: 'transparent',
    borderWidth: 1.5,
    borderDash: [5, 4],
    pointRadius: 0,
    tension: 0.3,
    spanGaps: true,
    yAxisID: 'y2',
  }});

  histChart = new Chart(ctx, {{
    type: 'line',
    data: {{ labels: DATA.labels, datasets }},
    options: {{
      responsive: true, maintainAspectRatio: false, devicePixelRatio: window.devicePixelRatio,
      plugins: {{
        legend: {{
          labels: {{ color: '#94a3b8', font: {{ size: 11 }}, boxWidth: 16 }}
        }},
        tooltip: {{
          backgroundColor: '#0f172a', borderColor: '#334155', borderWidth: 1,
          titleColor: '#94a3b8', bodyColor: '#e2e8f0',
          callbacks: {{
            label: ctx => ctx.dataset.label === 'Brent'
              ? ` Brent: $${{ctx.raw?.toFixed(2) ?? '—'}}/bbl`
              : ` ${{ctx.dataset.label}}: €${{ctx.raw?.toFixed(3) ?? '—'}}/L`
          }}
        }}
      }},
      scales: {{
        x: {{
          ticks: {{ color: '#e2e8f0', maxTicksLimit: 18, font: {{ size: 13, weight: '600' }} }},
          grid: {{ color: '#1e293b' }}
        }},
        y: {{
          position: 'right',
          ticks: {{ color: '#e2e8f0', callback: v => `€${{v.toFixed(2)}}`, font: {{ size: 13, weight: '600' }} }},
          grid: {{ color: '#1e293b' }},
          title: {{ display: true, text: 'Pump price (€/L)', color: '#e2e8f0', font: {{ size: 13, weight: '600' }} }},
        }},
        y2: {{
          position: 'left',
          ticks: {{ color: '#fdffbf99', callback: v => `$${{v.toFixed(0)}}`, font: {{ size: 13, weight: '600' }} }},
          grid: {{ display: false }},
          title: {{ display: true, text: 'Brent ($/bbl)', color: '#fdffbf99', font: {{ size: 13, weight: '600' }} }},
        }}
      }}
    }}
  }});
}}

function updateHistChart() {{
  const total = DATA.dates.length;
  const start = currentRange === 0 ? 0 : Math.max(0, total - currentRange);
  histChart.data.labels = DATA.labels.slice(start);
  histChart.data.datasets.forEach((ds, i) => {{
    if (ds.label === 'Brent') {{
      ds.data = DATA.brent.slice(start).map(v => v != null ? +v.toFixed(2) : null);
    }} else {{
      const c = CTRS[i];
      ds.data = DATA.countries[c][currentFuel].slice(start).map(v => v != null ? +(v/1000).toFixed(4) : null);
    }}
  }});
  histChart.update();
}}

function switchFuel(fuel) {{
  currentFuel = fuel;
  $('btn95').classList.toggle('active', fuel === 'euro95');
  $('btnD' ).classList.toggle('active', fuel === 'diesel');
  updateHistChart();
}}

function setRange(n) {{
  currentRange = n;
  $('btn1Y' ).classList.toggle('active', n === 52);
  $('btn3Y' ).classList.toggle('active', n === 156);
  $('btn5Y' ).classList.toggle('active', n === 260);
  $('btnAll').classList.toggle('active', n === 0);
  updateHistChart();
}}

// ---- PRICE CHART ---------------------------------------------------------
function buildPriceChart() {{
  const last      = DATA.dates.length - 1;
  const dieselData = CTRS.map(c => +((DATA.countries[c].diesel[last] || 0) / 1000).toFixed(4));
  const euro95Data = CTRS.map(c => +((DATA.countries[c].euro95[last] || 0) / 1000).toFixed(4));

  new Chart($('priceChart').getContext('2d'), {{
    type: 'bar',
    data: {{
      labels: CTRS,
      datasets: [
        {{
          label: 'Diesel',
          data: dieselData,
          backgroundColor: CTRS.map(c => COLORS[c] + 'cc'),
          borderRadius: 4,
        }},
        {{
          label: 'Euro-95',
          data: euro95Data,
          backgroundColor: CTRS.map(c => COLORS[c] + '55'),
          borderRadius: 4,
        }},
      ],
    }},
    options: {{
      responsive: true, maintainAspectRatio: false, devicePixelRatio: window.devicePixelRatio,
      plugins: {{
        legend: {{ display: false }},
        tooltip: {{
          backgroundColor: '#0f172a', borderColor: '#334155', borderWidth: 1,
          titleColor: '#94a3b8', bodyColor: '#e2e8f0',
          callbacks: {{ label: ctx => ` ${{ctx.dataset.label}}: €${{ctx.raw?.toFixed(4)}}/L` }}
        }},
      }},
      scales: {{
        x: {{ ticks: {{ color: '#e2e8f0', font: {{ size: 12, weight: 700 }} }}, grid: {{ color: '#1e293b' }} }},
        y: {{
          min: 1.30,
          ticks: {{ color: '#e2e8f0', callback: v => `€${{v.toFixed(2)}}`, font: {{ size: 13, weight: '600' }} }},
          grid: {{ color: '#1e293b' }},
          title: {{ display: true, text: '€/L', color: '#e2e8f0', font: {{ size: 13, weight: '600' }} }},
        }},
      }},
    }},
    plugins: [{{
      id: 'priceLabels',
      afterDatasetsDraw(chart) {{
        const ctx2 = chart.ctx;
        const fuelLabels = ['Diesel', '95'];
        chart.data.datasets.forEach((ds, i) => {{
          chart.getDatasetMeta(i).data.forEach((bar, j) => {{
            const val = ds.data[j];
            if (val == null) return;
            ctx2.save();
            ctx2.textAlign = 'center';
            // fuel type
            ctx2.font = '10px DM Sans, sans-serif';
            ctx2.fillStyle = '#cbd5e1';
            ctx2.textBaseline = 'bottom';
            ctx2.fillText(fuelLabels[i], bar.x, bar.y - 15);
            // value
            ctx2.font = 'bold 11px DM Sans, sans-serif';
            ctx2.fillStyle = '#f8fafc';
            ctx2.fillText(`€${{val.toFixed(2)}}`, bar.x, bar.y - 3);
            ctx2.restore();
          }});
        }});
      }}
    }}],
  }});
}}

// ---- YTD -----------------------------------------------------------------
function buildYTD() {{
  const yr = DATA.dates[DATA.dates.length-1].slice(0,4);
  $('ytd-year').textContent = yr;

  // Cards
  const wrap = $('ytd-cards');
  CTRS.forEach(c => {{
    const y = DATA.ytd[c];
    const el = document.createElement('div');
    el.className = 'ytd-card';
    el.style.borderColor = COLORS[c] + '40';
    const cls95 = (y.euro95_ytd ?? 0) < 0 ? 'dn' : 'up';
    const clsD  = (y.diesel_ytd  ?? 0) < 0 ? 'dn' : 'up';
    function fmtLine(abs, pct) {{
      if (abs == null && pct == null) return '—';
      const absStr = abs != null ? `${{abs>=0?'+':''}}${{abs.toFixed(2)}} €` : '';
      const pctStr = pct != null ? `(${{pct>=0?'+':''}}${{pct.toFixed(1)}}%)` : '';
      return absStr && pctStr ? `${{absStr}} ${{pctStr}}` : absStr || pctStr;
    }}
    el.innerHTML = `
      <div class="ytd-ctr" style="color:${{COLORS[c]}}">${{c}}</div>
      <div class="ytd-sub">Diesel</div>
      <div class="ytd-val ${{clsD}} mono">${{fmtLine(y.diesel_abs, y.diesel_ytd)}}</div>
      <div class="ytd-sub" style="margin-top:6px">Euro-95</div>
      <div class="ytd-val2 ${{cls95}} mono">${{fmtLine(y.euro95_abs, y.euro95_ytd)}}</div>`;
    wrap.appendChild(el);
  }});

  // Brent card
  if (DATA.brent_ytd != null) {{
    const clsB = DATA.brent_ytd < 0 ? 'dn' : 'up';
    const el = document.createElement('div');
    el.className = 'ytd-card';
    el.style.borderColor = '#94a3b840';
    el.innerHTML = `
      <div class="ytd-ctr" style="color:#94a3b8">Brent</div>
      <div class="ytd-sub">$/barrel</div>
      <div class="ytd-val ${{clsB}} mono">${{(DATA.brent_ytd>0?'+':'')+DATA.brent_ytd.toFixed(2)+'%'}}</div>`;
    wrap.appendChild(el);
  }}

  // Bar chart
  const ytdData95   = CTRS.map(c => DATA.ytd[c].euro95_ytd);
  const ytdDataD    = CTRS.map(c => DATA.ytd[c].diesel_ytd);
  const chartLabels = [...CTRS, 'Brent'];

  ytdChart = new Chart($('ytdChart').getContext('2d'), {{
    type: 'bar',
    data: {{
      labels: chartLabels,
      datasets: [
        {{
          label: 'Diesel YTD%',
          data: [...ytdDataD, null],
          backgroundColor: [...CTRS.map(c => COLORS[c] + 'cc'), 'transparent'],
          borderRadius: 4,
          yAxisID: 'y',
        }},
        {{
          label: 'Euro-95 YTD%',
          data: [...ytdData95, null],
          backgroundColor: [...CTRS.map(c => COLORS[c] + '55'), 'transparent'],
          borderRadius: 4,
          yAxisID: 'y',
        }},
        {{
          label: 'Brent YTD%',
          data: [...CTRS.map(() => null), DATA.brent_ytd],
          backgroundColor: '#94a3b880',
          borderColor: '#94a3b8',
          borderWidth: 1,
          borderRadius: 4,
          yAxisID: 'y2',
          order: 0,
        }},
      ]
    }},
    options: {{
      responsive: true, maintainAspectRatio: false, devicePixelRatio: window.devicePixelRatio,
      plugins: {{
        legend: {{ labels: {{ color: '#94a3b8', font: {{ size: 11 }} }} }},
        tooltip: {{
          backgroundColor: '#0f172a', borderColor: '#334155', borderWidth: 1,
          titleColor: '#94a3b8', bodyColor: '#e2e8f0',
          callbacks: {{ label: ctx => ` ${{ctx.dataset.label}}: ${{ctx.raw?.toFixed(2)}}%` }}
        }}
      }},
      scales: {{
        x: {{ ticks: {{ color: '#e2e8f0', font: {{ size: 12, weight: 700 }} }}, grid: {{ color: '#1e293b' }} }},
        y: {{
          ticks: {{ color: '#e2e8f0', callback: v => v+'%', font: {{ size: 13, weight: '600' }} }},
          grid: {{ color: '#1e293b' }},
          title: {{ display: true, text: 'Pump price YTD (%)', color: '#e2e8f0', font: {{ size: 13, weight: '600' }} }},
        }},
        y2: {{
          type: 'linear',
          position: 'right',
          ticks: {{ color: '#e2e8f0', callback: v => v+'%', font: {{ size: 13, weight: '600' }} }},
          grid: {{ display: false }},
          title: {{ display: true, text: 'Brent YTD (%)', color: '#e2e8f0', font: {{ size: 13, weight: '600' }} }},
        }}
      }}
    }},
    plugins: [{{
      id: 'ytdLabels',
      afterDatasetsDraw(chart) {{
        const ctx2 = chart.ctx;
        chart.data.datasets.forEach((ds, i) => {{
          chart.getDatasetMeta(i).data.forEach((bar, j) => {{
            const val = ds.data[j];
            if (val == null) return;
            const text = (val >= 0 ? '+' : '') + val.toFixed(1) + '%';
            ctx2.save();
            ctx2.fillStyle = '#e2e8f0';
            ctx2.font = 'bold 10px DM Sans, sans-serif';
            ctx2.textAlign = 'center';
            ctx2.textBaseline = val >= 0 ? 'bottom' : 'top';
            ctx2.fillText(text, bar.x, val >= 0 ? bar.y - 3 : bar.y + 3);
            ctx2.restore();
          }});
        }});
      }}
    }}, {{
      id: 'ytdDivider',
      afterDraw(chart) {{
        const xScale = chart.scales.x;
        const sepIdx = CTRS.length - 1;  // between PT (last country) and Brent
        const x1 = xScale.getPixelForValue(sepIdx);
        const x2 = xScale.getPixelForValue(sepIdx + 1);
        const midX = (x1 + x2) / 2;
        const {{ top, bottom }} = chart.chartArea;
        const ctx2 = chart.ctx;
        ctx2.save();
        ctx2.strokeStyle = '#94a3b8';
        ctx2.lineWidth = 1;
        ctx2.setLineDash([5, 4]);
        ctx2.beginPath();
        ctx2.moveTo(midX, top);
        ctx2.lineTo(midX, bottom);
        ctx2.stroke();
        ctx2.setLineDash([]);
        ctx2.restore();
      }}
    }}]
  }});
}}

// ---- TAX CHARTS ----------------------------------------------------------
function buildTaxCharts() {{
  const last = DATA.dates.length - 1;
  const ctrs = CTRS.filter(c => c !== 'EU');

  // Diagonal-stripe hatch pattern for the Tax & Duties segment
  function makeHatch(color) {{
    const sz = 8;
    const cv = document.createElement('canvas');
    cv.width = sz; cv.height = sz;
    const cx = cv.getContext('2d');
    cx.fillStyle = color + 'aa';
    cx.fillRect(0, 0, sz, sz);
    cx.strokeStyle = color;
    cx.lineWidth = 2;
    cx.beginPath();
    cx.moveTo(-1, sz + 1); cx.lineTo(sz + 1, -1);
    cx.moveTo(-1 - sz, sz + 1); cx.lineTo(sz + 1 - sz, -1);
    cx.moveTo(-1 + sz, sz + 1); cx.lineTo(sz + 1 + sz, -1);
    cx.stroke();
    return cx.createPattern(cv, 'repeat');
  }}

  function taxDatasets(fuelKey, fuelKeyNt) {{
    return [
      {{
        label: 'Pre-tax',
        data: ctrs.map(c => {{
          const v = DATA.countries[c][fuelKeyNt][last];
          return v != null ? +(v/1000).toFixed(4) : null;
        }}),
        backgroundColor: ctrs.map(c => COLORS[c]),
        borderRadius: [0,0,0,0],
        stack: 'a',
      }},
      {{
        label: 'Tax & Duties',
        data: ctrs.map(c => {{
          const gross = DATA.countries[c][fuelKey][last];
          const net   = DATA.countries[c][fuelKeyNt][last];
          if (gross == null || net == null) return null;
          return +((gross - net)/1000).toFixed(4);
        }}),
        backgroundColor: ctrs.map(c => COLORS[c] + '55'),
        borderRadius: [4,4,0,0],
        stack: 'a',
      }},
    ];
  }}

  // Plugin: value labels inside each stacked segment
  const stackLabels = {{
    id: 'stackLabels',
    afterDatasetsDraw(chart) {{
      const ctx2 = chart.ctx;
      chart.data.datasets.forEach((ds, i) => {{
        chart.getDatasetMeta(i).data.forEach((bar, j) => {{
          const val = ds.data[j];
          if (val == null || val < 0.05) return;
          const segH = Math.abs(bar.base - bar.y);
          if (segH < 14) return;
          const midY = (bar.y + bar.base) / 2;
          ctx2.save();
          ctx2.fillStyle = '#fff';
          ctx2.font = 'bold 10px DM Sans, sans-serif';
          ctx2.textAlign = 'center';
          ctx2.textBaseline = 'middle';
          ctx2.fillText(`€${{val.toFixed(2)}}`, bar.x, midY);
          ctx2.restore();
        }});
      }});
    }}
  }};

  const baseOpts = {{
    responsive: true, maintainAspectRatio: false, devicePixelRatio: window.devicePixelRatio,
    plugins: {{
      legend: {{ labels: {{ color:'#94a3b8', font:{{ size:10 }} }} }},
      tooltip: {{
        backgroundColor:'#0f172a', borderColor:'#334155', borderWidth:1,
        titleColor:'#94a3b8', bodyColor:'#e2e8f0',
        callbacks: {{ label: ctx => ` ${{ctx.dataset.label}}: €${{ctx.raw?.toFixed(3) ?? '—'}}/L` }}
      }}
    }},
    scales: {{
      x: {{ stacked:true, ticks:{{ color:'#e2e8f0', font:{{ size:11 }} }}, grid:{{ color:'#1e293b' }} }},
      y: {{ stacked:true,
            ticks:{{ color:'#e2e8f0', callback: v=>`€${{v.toFixed(2)}}`, font:{{ size:10 }} }},
            grid:{{ color:'#1e293b' }} }}
    }}
  }};

  taxDChart = new Chart($('taxDChart').getContext('2d'), {{
    type:'bar',
    data: {{ labels: ctrs, datasets: taxDatasets('diesel','diesel_notax') }},
    options: baseOpts,
    plugins: [stackLabels],
  }});
  tax95Chart = new Chart($('tax95Chart').getContext('2d'), {{
    type:'bar',
    data: {{ labels: ctrs, datasets: taxDatasets('euro95','euro95_notax') }},
    options: baseOpts,
    plugins: [stackLabels],
  }});

  // Tables under each chart
  function buildTaxTable(tableId, fuelKey, fuelKeyNt) {{
    const wrap = $(tableId);
    const tbl = document.createElement('table');
    tbl.style.cssText = 'width:100%;border-collapse:collapse;font-size:11px;margin-top:8px;';
    tbl.innerHTML = `<thead><tr>
      <th style="text-align:left;padding:4px 8px;color:#94a3b8;font-weight:600;border-bottom:1px solid #1e293b;">Country</th>
      <th style="text-align:right;padding:4px 8px;color:#94a3b8;font-weight:600;border-bottom:1px solid #1e293b;">Pre-tax</th>
      <th style="text-align:right;padding:4px 8px;color:#94a3b8;font-weight:600;border-bottom:1px solid #1e293b;">Tax & Duties</th>
      <th style="text-align:right;padding:4px 8px;color:#94a3b8;font-weight:600;border-bottom:1px solid #1e293b;">Total</th>
      <th style="text-align:right;padding:4px 8px;color:#94a3b8;font-weight:600;border-bottom:1px solid #1e293b;">Tax Rate</th>
    </tr></thead><tbody></tbody>`;
    const tbody = tbl.querySelector('tbody');
    ctrs.forEach(c => {{
      const gross  = DATA.countries[c][fuelKey][last];
      const net    = DATA.countries[c][fuelKeyNt][last];
      const pretax = net   != null ? `€${{(net/1000).toFixed(3)}}` : '—';
      const tax    = (gross && net) ? `€${{((gross-net)/1000).toFixed(3)}}` : '—';
      const total  = gross != null ? `€${{(gross/1000).toFixed(3)}}` : '—';
      const rate   = (gross && net) ? Math.round((gross-net)/gross*100)+'%' : '—';
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td style="padding:4px 8px;color:${{COLORS[c]}};font-weight:700;">${{c}}</td>
        <td class="mono" style="text-align:right;padding:4px 8px;color:#e2e8f0;">${{pretax}}</td>
        <td class="mono" style="text-align:right;padding:4px 8px;color:#e2e8f0;">${{tax}}</td>
        <td class="mono" style="text-align:right;padding:4px 8px;color:#f1f5f9;font-weight:700;">${{total}}</td>
        <td class="mono" style="text-align:right;padding:4px 8px;color:#94a3b8;">${{rate}}</td>`;
      tbody.appendChild(tr);
    }});
    wrap.appendChild(tbl);
  }}
  buildTaxTable('taxDTable', 'diesel', 'diesel_notax');
  buildTaxTable('tax95Table', 'euro95', 'euro95_notax');

  // Tax rate horizontal bars
  const barWrap = $('tax-bars');
  const sorted = [...ctrs].sort((a,b) => DATA.countries[b].tax_rate - DATA.countries[a].tax_rate);
  sorted.forEach(c => {{
    const pct = Math.round(DATA.countries[c].tax_rate * 100);
    const last_95 = DATA.countries[c].euro95[DATA.dates.length-1];
    const last_nt = DATA.countries[c].euro95_notax[DATA.dates.length-1];
    const taxAmt  = (last_95 && last_nt) ? ((last_95 - last_nt)/1000).toFixed(3) : '—';
    const el = document.createElement('div');
    el.className = 'tax-bar-row';
    el.innerHTML = `
      <div class="tax-bar-label" style="color:${{COLORS[c]}}">${{c}}</div>
      <div class="tax-bar-track">
        <div class="tax-bar-fill" style="width:${{pct}}%;background:linear-gradient(90deg,${{COLORS[c]}}60,${{COLORS[c]}})">
          <span class="tax-bar-pct mono">${{pct}}%</span>
        </div>
      </div>
      <div class="tax-bar-aside mono">€${{taxAmt}}/L tax</div>`;
    barWrap.appendChild(el);
  }});
}}

// ---- CONSUMPTION ---------------------------------------------------------
function switchCons(mode) {{
  currentCons = mode;
  $('btnAbsolute').classList.toggle('active', mode === 'absolute');
  $('btnMix'     ).classList.toggle('active', mode === 'mix');
  $('cons-absolute').style.display = mode === 'absolute' ? '' : 'none';
  $('cons-mix'     ).style.display = mode === 'mix'      ? '' : 'none';
}}

// Shared legend options for both cons charts (no generateLabels override — let Chart.js do it correctly)
function consLegendOpts(tooltipFmt) {{
  return {{
    responsive: true, maintainAspectRatio: false, devicePixelRatio: window.devicePixelRatio,
    plugins: {{
      legend: {{
        position: 'top',
        labels: {{
          color: '#94a3b8',
          font: {{ size: 11 }},
          boxWidth: 14,
          boxHeight: 14,
          padding: 16,
          usePointStyle: false,
        }}
      }},
      tooltip: {{
        backgroundColor: '#0f172a', borderColor: '#334155', borderWidth: 1,
        titleColor: '#94a3b8', bodyColor: '#e2e8f0',
        callbacks: {{ label: tooltipFmt }}
      }}
    }},
    scales: {{
      x: {{ stacked: true, ticks: {{ color:'#e2e8f0', font:{{ size:13, weight:'700' }} }}, grid:{{ color:'#1e293b' }} }},
      y: {{ stacked: true, ticks: {{ color:'#e2e8f0', font:{{ size:10 }} }},               grid:{{ color:'#1e293b' }} }}
    }}
  }};
}}

function buildConsumption() {{
  const ctrs = CTRS.filter(c => c !== 'EU');
  $('cons-year').textContent = DATA.latest_year;

  // ---- Absolute volumes chart -------------------------------------------
  const absOpts = consLegendOpts(ctx => ` ${{ctx.dataset.label}}: ${{ctx.raw != null ? ctx.raw.toLocaleString('en',{{maximumFractionDigits:0}}) : '—'}} kt`);
  absOpts.scales.y.ticks.callback = v => v >= 1000 ? (v/1000).toFixed(0)+'k' : v;

  consAbsChart = new Chart($('consAbsChart').getContext('2d'), {{
    type: 'bar',
    data: {{
      labels: ctrs,
      datasets: FUEL_TYPES.map(f => ({{
        label: f,
        data: ctrs.map(c => DATA.consumption[c][f] || 0),
        backgroundColor: FUEL_COLORS[f],
        stack: 'a',
      }}))
    }},
    options: absOpts,
  }});

  // ---- Percentage mix chart ---------------------------------------------
  const pctData = ctrs.map(c => {{
    const vals  = DATA.consumption[c];
    const total = Object.values(vals).reduce((a,b)=>a+b, 0);
    const row   = {{}};
    FUEL_TYPES.forEach(f => {{ row[f] = total > 0 ? +((vals[f]||0)/total*100).toFixed(1) : 0; }});
    return row;
  }});

  const mixOpts = consLegendOpts(ctx => ` ${{ctx.dataset.label}}: ${{ctx.raw?.toFixed(1) ?? '—'}}%`);
  mixOpts.scales.y.min = 0;
  mixOpts.scales.y.max = 100;
  mixOpts.scales.y.ticks.callback = v => v + '%';

  consMixChart = new Chart($('consMixChart').getContext('2d'), {{
    type: 'bar',
    data: {{
      labels: ctrs,
      datasets: FUEL_TYPES.map(f => ({{
        label: f,
        data: pctData.map(r => r[f]),
        backgroundColor: FUEL_COLORS[f],
        stack: 'a',
      }}))
    }},
    options: mixOpts,
  }});

  // ---- Absolute table ---------------------------------------------------
  const thead = $('cons-head');
  ['Country', ...FUEL_TYPES, 'Total'].forEach(h => {{
    const th = document.createElement('th');
    th.textContent = h;
    th.style.color = FUEL_COLORS[h] || '#94a3b8';
    if (h === 'Country') th.style.textAlign = 'left';
    thead.appendChild(th);
  }});
  const tbody = $('cons-body');
  ctrs.forEach(c => {{
    const v     = DATA.consumption[c];
    const total = Object.values(v).reduce((a,b) => a+b, 0);
    const tr    = document.createElement('tr');
    tr.innerHTML = `
      <td><span class="dot" style="background:${{COLORS[c]}}"></span>
          <span style="color:${{COLORS[c]}};font-weight:700">${{c}}</span></td>
      ${{FUEL_TYPES.map(f => `<td class="mono" style="color:#94a3b8">${{fmtVal(v[f]||0, 0)}}</td>`).join('')}}
      <td class="mono" style="font-weight:700;color:#f1f5f9">${{fmtVal(total, 0)}}</td>`;
    tbody.appendChild(tr);
  }});
}}
// ---- SENSITIVITY ---------------------------------------------------------
function buildSensitivity() {{
  // All OLS slopes are pre-computed in Python; JS only handles rendering.
  function makeBarChart(canvasId, slopeId, fuelKey) {{
    const sens = DATA.sensitivity[fuelKey].per_country;
    const upData   = CTRS.map(c => sens[c].up);
    const downData = CTRS.map(c => sens[c].down);
    const avgUp   = (upData.reduce((s, v) => s + v, 0) / upData.length).toFixed(1);
    const avgDown = (downData.reduce((s, v) => s + v, 0) / downData.length).toFixed(1);
    $(slopeId).textContent = `avg ▲ ${{avgUp}}  ▼ ${{avgDown}}  €cts/L per $10 Brent`;

    return new Chart($(canvasId).getContext('2d'), {{
      type: 'bar',
      data: {{
        labels: CTRS,
        datasets: [
          {{
            label: 'Brent ▲ (rising weeks)',
            data: upData,
            backgroundColor: CTRS.map(c => COLORS[c] + 'cc'),
            borderRadius: 4,
          }},
          {{
            label: 'Brent ▼ (falling weeks)',
            data: downData,
            backgroundColor: CTRS.map(c => COLORS[c] + '55'),
            borderRadius: 4,
          }},
        ],
      }},
      options: {{
        responsive: true, maintainAspectRatio: false, devicePixelRatio: window.devicePixelRatio,
        plugins: {{
          legend: {{ labels: {{ color: '#94a3b8', font: {{ size: 11 }} }} }},
          tooltip: {{
            backgroundColor: '#0f172a', borderColor: '#334155', borderWidth: 1,
            titleColor: '#94a3b8', bodyColor: '#e2e8f0',
            callbacks: {{ label: ctx => ` ${{ctx.dataset.label}}: ${{ctx.raw?.toFixed(1)}} €cts/L per $10 Brent` }}
          }}
        }},
        scales: {{
          x: {{ ticks: {{ color: '#e2e8f0', font: {{ size: 12, weight: 700 }} }}, grid: {{ color: '#1e293b' }} }},
          y: {{
            min: 0, max: 10,
            ticks: {{ color: '#e2e8f0', callback: v => v.toFixed(1), font: {{ size: 13, weight: '600' }} }},
            grid: {{ color: '#1e293b' }},
            title: {{ display: true, text: '€ cents/L · $10 Brent', color: '#e2e8f0', font: {{ size: 13, weight: '600' }} }},
          }},
        }},
      }},
      plugins: [{{
        id: 'sensLabels',
        afterDatasetsDraw(chart) {{
          const ctx2 = chart.ctx;
          chart.data.datasets.forEach((ds, i) => {{
            chart.getDatasetMeta(i).data.forEach((bar, j) => {{
              const val = ds.data[j];
              if (val == null || val === 0) return;
              ctx2.save();
              ctx2.fillStyle = '#e2e8f0';
              ctx2.font = 'bold 10px DM Sans, sans-serif';
              ctx2.textAlign = 'center';
              ctx2.textBaseline = 'bottom';
              ctx2.fillText(val.toFixed(1), bar.x, bar.y - 3);
              ctx2.restore();
            }});
          }});
        }}
      }}],
    }});
  }}

  makeBarChart('sensDieselChart', 'sens-diesel-slope', 'diesel');
  makeBarChart('sens95Chart',     'sens-95-slope',     'euro95');

  // Asymmetry tables (one per fuel, aligned under each chart)
  const sensD  = DATA.sensitivity.diesel.per_country;
  const sens95 = DATA.sensitivity.euro95.per_country;
  const fmtGap = v => (v >= 0 ? '+' : '') + v.toFixed(1);

  function buildAsymTable(wrapperId, sensData, upLabel, dnLabel) {{
    const wrap = $(wrapperId);
    const tbl = document.createElement('table');
    tbl.style.cssText = 'width:100%;border-collapse:collapse;font-size:12px;';
    tbl.innerHTML = `<thead><tr>
      <th style="text-align:left;padding:8px 16px;color:#94a3b8;font-weight:600;border-bottom:1px solid #1e293b;">Country</th>
      <th style="text-align:right;padding:8px 12px;color:#94a3b8;font-weight:600;border-bottom:1px solid #1e293b;">${{upLabel}}</th>
      <th style="text-align:right;padding:8px 12px;color:#94a3b8;font-weight:600;border-bottom:1px solid #1e293b;">${{dnLabel}}</th>
      <th style="text-align:right;padding:8px 12px;color:#94a3b8;font-weight:600;border-bottom:1px solid #1e293b;">Gap</th>
    </tr></thead><tbody></tbody>`;
    const tbody = tbl.querySelector('tbody');
    CTRS.filter(c => c !== 'EU').forEach(c => {{
      const up = sensData[c].up; const dn = sensData[c].down; const gap = up - dn;
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td style="padding:6px 16px;color:${{COLORS[c]}};font-weight:700;">${{c}}</td>
        <td class="mono" style="text-align:right;padding:6px 12px;color:#e2e8f0;">${{up.toFixed(1)}}</td>
        <td class="mono" style="text-align:right;padding:6px 12px;color:#e2e8f0;">${{dn.toFixed(1)}}</td>
        <td class="mono" style="text-align:right;padding:6px 12px;color:#e2e8f0;font-weight:700;">${{fmtGap(gap)}}</td>`;
      tbody.appendChild(tr);
    }});
    wrap.appendChild(tbl);
  }}
  buildAsymTable('sens-diesel-table', sensD,  'Brent ▲', 'Brent ▼');
  buildAsymTable('sens-95-table',     sens95, 'Brent ▲', 'Brent ▼');

  // ---- Window × Lag research grid (data pre-computed in Python) ----------
  function buildSensResearch() {{
    const WINS = [4, 26];
    const LAGS = [0, 1, 2];

    function fillResearchTable(wrapperId, fuelKey) {{
      const research = DATA.sensitivity[fuelKey].research;
      const wrap = $(wrapperId);
      const tbl  = document.createElement('table');
      tbl.style.cssText = 'width:100%;border-collapse:collapse;font-size:12px;';
      tbl.innerHTML = `<thead><tr>
        <th style="padding:6px 12px;color:#94a3b8;font-weight:600;border-bottom:1px solid #1e293b;text-align:center;">Window</th>
        <th style="padding:6px 12px;color:#94a3b8;font-weight:600;border-bottom:1px solid #1e293b;text-align:center;">Lag</th>
        <th style="padding:6px 12px;color:#f59e0b;font-weight:600;border-bottom:1px solid #1e293b;text-align:right;">Brent ▲</th>
        <th style="padding:6px 12px;color:#64748b;font-weight:600;border-bottom:1px solid #1e293b;text-align:right;">Brent ▼</th>
        <th style="padding:6px 12px;color:#94a3b8;font-weight:600;border-bottom:1px solid #1e293b;text-align:right;">Gap</th>
      </tr></thead><tbody></tbody>`;
      const tbody = tbl.querySelector('tbody');
      WINS.forEach((w, wi) => {{
        if (wi > 0) {{
          const sep = document.createElement('tr');
          sep.innerHTML = '<td colspan="5" style="padding:0;border-top:2px solid #334155;"></td>';
          tbody.appendChild(sep);
        }}
        LAGS.forEach(l => {{
          const s         = research[`${{w}}_${{l}}`];
          const gap       = s.up - s.down;
          const isCurrent = (w === 4 && l === 1);
          const tr = document.createElement('tr');
          if (isCurrent) tr.style.cssText = 'background:rgba(245,158,11,0.12);';
          tr.innerHTML = `
            <td style="text-align:center;padding:5px 12px;color:#e2e8f0;font-weight:${{isCurrent?700:400}};">
              ${{w}}W${{isCurrent?' ★':''}}
            </td>
            <td style="text-align:center;padding:5px 12px;color:#94a3b8;">lag ${{l}}W</td>
            <td class="mono" style="text-align:right;padding:5px 12px;color:#f59e0b;">${{s.up.toFixed(1)}}</td>
            <td class="mono" style="text-align:right;padding:5px 12px;color:#64748b;">${{s.down.toFixed(1)}}</td>
            <td class="mono" style="text-align:right;padding:5px 12px;color:#e2e8f0;font-weight:700;">${{(gap>=0?'+':'')}}${{gap.toFixed(1)}}</td>`;
          tbody.appendChild(tr);
        }});
      }});
      wrap.appendChild(tbl);
    }}

    fillResearchTable('sens-research-diesel', 'diesel');
    fillResearchTable('sens-research-95',     'euro95');
  }}
  buildSensResearch();
}}
</script>
</body>
</html>"""


# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------
WEBPAGE_NAME  = "oil_dashboard"
GHPAGES_REPO  = Path(__file__).resolve().parent.parent / "sebast759.github.io"
DEFAULT_OUT   = str(GHPAGES_REPO / WEBPAGE_NAME / "index.html")


def push_to_github_pages(html: str, token: str):
    """Push HTML directly to GitHub Pages via the GitHub Contents API.

    No local git repo required — works on Heroku's ephemeral filesystem.

    Env vars (all optional except GITHUB_TOKEN):
      GITHUB_TOKEN     — personal access token with repo write scope
      GITHUB_REPO      — owner/repo  (default: sebast759/sebast759.github.io)
      GITHUB_FILE_PATH — path in repo (default: oil_dashboard/index.html)
    """
    import urllib.request
    import urllib.error

    repo      = os.environ.get("GITHUB_REPO",      "sebast759/sebast759.github.io")
    file_path = os.environ.get("GITHUB_FILE_PATH", "oil_dashboard/index.html")
    api_url   = f"https://api.github.com/repos/{repo}/contents/{file_path}"
    headers   = {
        "Authorization": f"token {token}",
        "Accept":        "application/vnd.github.v3+json",
        "Content-Type":  "application/json",
        "User-Agent":    "oil-dashboard-scheduler",
    }

    # Fetch current file SHA — required by the API for updates (not for first push)
    sha = None
    try:
        with urllib.request.urlopen(
            urllib.request.Request(api_url, headers=headers)
        ) as resp:
            sha = json.loads(resp.read())["sha"]
    except urllib.error.HTTPError as e:
        if e.code != 404:
            raise

    payload = {
        "message": f"chore: update oil dashboard ({datetime.now().strftime('%Y-%m-%d %H:%M')})",
        "content": base64.b64encode(html.encode()).decode(),
    }
    if sha:
        payload["sha"] = sha

    req = urllib.request.Request(
        api_url,
        data=json.dumps(payload).encode(),
        headers=headers,
        method="PUT",
    )
    with urllib.request.urlopen(req):
        pass

    owner = repo.split("/")[0]
    print(f"  Live at https://{owner}.github.io/{file_path.split('/')[0]}/")


def git_push(ghpages_repo: Path, webpage_name: str):
    """Pull, stage, commit and push if there are changes."""
    import subprocess
    print("\nPushing to GitHub Pages ...")
    subprocess.run(["git", "pull"], cwd=ghpages_repo, check=True)
    subprocess.run(["git", "add", f"{webpage_name}/index.html"],
                   cwd=ghpages_repo, check=True)
    result = subprocess.run(["git", "diff", "--cached", "--quiet"],
                             cwd=ghpages_repo)
    if result.returncode != 0:
        subprocess.run(["git", "commit", "-m", f"Update {webpage_name} dashboard"],
                       cwd=ghpages_repo, check=True)
        subprocess.run(["git", "push"], cwd=ghpages_repo, check=True)
        print(f"  Live at https://sebast759.github.io/{webpage_name}/")
    else:
        print("  No changes to push.")


def main():
    parser = argparse.ArgumentParser(
        description="Generate EU Oil Bulletin Dashboard",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python generate_oil_dashboard.py                  # auto-download if needed, use cache if fresh
  python generate_oil_dashboard.py --download       # force re-download from EC website
  python generate_oil_dashboard.py my_file.xlsx     # use a specific local file
  python generate_oil_dashboard.py --no-push        # generate only, skip git push
  python generate_oil_dashboard.py --local          # skip all network calls, use cache

Output (default):
  """ + DEFAULT_OUT
    )
    parser.add_argument("input", nargs="?", default=DEFAULT_XLSX,
                        help="Path to xlsx (optional — auto-downloads if omitted)")
    parser.add_argument("--output", "-o", default=DEFAULT_OUT,
                        help=f"Output HTML file (default: {DEFAULT_OUT})")
    parser.add_argument("--download", "-d", action="store_true",
                        help="Force re-download the latest file from EC website")
    parser.add_argument("--cache-dir", default=".",
                        help="Directory for cached xlsx (default: current folder)")
    parser.add_argument("--no-push", action="store_true",
                        help="Skip git push — just generate the HTML locally")
    parser.add_argument("--local", action="store_true",
                        help="Skip all network calls; use cached xlsx and brent_cache.csv")
    args = parser.parse_args()

    xlsx_path = resolve_xlsx(args.input, args.download, Path(args.cache_dir), local=args.local)

    data = extract_data(xlsx_path, local=args.local)
    html = build_html(data)

    github_token = os.environ.get("PAGES_TOKEN")

    if github_token and not args.no_push:
        # Heroku / CI: push directly via GitHub API — no local git repo needed
        print("\nPushing to GitHub Pages via API ...")
        push_to_github_pages(html, github_token)
    else:
        # Local: write file to disk, optionally git-push
        out_path = Path(args.output)
        out_path.parent.mkdir(parents=True, exist_ok=True)
        out_path.write_text(html, encoding="utf-8")
        print(f"\n  Dashboard saved to: {out_path.resolve()}")

        if not args.no_push and not args.local:
            if GHPAGES_REPO.exists():
                git_push(GHPAGES_REPO, WEBPAGE_NAME)
            else:
                print(f"  WARNING: GitHub Pages repo not found at {GHPAGES_REPO}")
                print(f"  Run with --no-push or set GITHUB_TOKEN for API push.")


if __name__ == "__main__":
    main()
