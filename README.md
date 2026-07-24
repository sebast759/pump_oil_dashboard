# EU Weekly Oil Bulletin Dashboard

Auto-generated dashboard tracking EU retail fuel prices (diesel & Euro-95) across France, Germany, Netherlands, Spain, Italy and Portugal — with Brent crude overlay, YTD performance, tax breakdown, consumption mix and pass-through sensitivity analysis.

**Live dashboard → [sebast759.github.io/oil_dashboard](https://sebast759.github.io/oil_dashboard/)**

---

## What it does

Each week, the script:
1. Downloads the latest [EC Weekly Oil Bulletin](https://energy.ec.europa.eu/data-and-analysis/weekly-oil-bulletin_en) Excel file
2. Fetches weekly Brent crude prices from Yahoo Finance
3. Generates a self-contained HTML dashboard
4. Pushes it to GitHub Pages via the GitHub API

The dashboard updates automatically every **Thursday at 14:00 UTC**, shortly after the EC publishes new data.

---

## Dashboard tabs

| Tab | Content |
|-----|---------|
| **Historical Prices** | 36-month price lines per country + Brent overlay; YTD bar chart |
| **YTD Performance** | Year-to-date % change cards and bar chart |
| **Tax Analysis** | Pre-tax vs tax/duties stacked bars per country |
| **Consumption** | Fuel mix by country (gasoline, diesel, heating oil, LPG) |
| **Sensitivity** | OLS pass-through asymmetry: Brent rises passed on faster than declines |

---

## Sensitivity methodology

The sensitivity tab measures how much pump prices (€ cents/L) move per $10 Brent move, split by direction:

- **Window**: 4 weeks — standard in the pass-through literature (Borenstein, Cameron & Gilbert 1997, *QJE*)
- **Lag**: 1 week — aligns Brent signal with EC procurement window (Bacon 1991)
- **Model**: OLS slope through origin, computed in Python at generation time

A research table cross-tabulates 4W / 26W windows × 0–2 week lags to validate the parameter choice.

---

## Running locally

```bash
pip install openpyxl

# Full run — downloads fresh data, writes to ../sebast759.github.io/oil_dashboard/index.html
python generate_oil_dashboard.py

# Force re-download even if cache is fresh
python generate_oil_dashboard.py --download

# Use cached files only, skip all network calls
python generate_oil_dashboard.py --local

# Generate without pushing
python generate_oil_dashboard.py --no-push

# Custom output path
python generate_oil_dashboard.py --output ./preview.html
```

---

## Automated deployment (GitHub Actions)

The workflow at `.github/workflows/update_dashboard.yml` runs every Thursday at 14:00 UTC.

### First-time setup

1. Create a GitHub [Personal Access Token](https://github.com/settings/tokens) with **Contents: read & write** access on `sebast759/sebast759.github.io`
2. Add it as a secret in this repo:
   - Settings → Secrets and variables → Actions → New secret
   - Name: `PAGES_TOKEN`

The workflow then runs `generate_oil_dashboard.py` with `GITHUB_TOKEN` set, which triggers the GitHub API push path (no local git repo required).

### Manual trigger

Actions tab → **Update Oil Dashboard** → **Run workflow**

---

## Project structure

```
pump_oil_dashboard/
├── generate_oil_dashboard.py   # main script — data extraction, OLS, HTML generation
├── requirements.txt            # openpyxl only
├── .gitignore                  # excludes *.xlsx and brent_cache.csv
└── .github/
    └── workflows/
        └── update_dashboard.yml
```

### Runtime cache (local only, git-ignored)

| File | Purpose |
|------|---------|
| `*.xlsx` | EC bulletin — re-downloaded if older than 7 days |
| `brent_cache.csv` | Weekly Brent prices — fallback if Yahoo Finance is unavailable |

---

## Data sources

| Data | Source | Frequency |
|------|--------|-----------|
| Fuel prices (with/without tax) | [EC Weekly Oil Bulletin](https://energy.ec.europa.eu/data-and-analysis/weekly-oil-bulletin_en) | Weekly (Thursday) |
| Brent crude | Yahoo Finance (`BZ=F`) | Weekly + latest daily |
| Consumption mix | EC Oil Bulletin (Consumption sheet) | Annual |

## Continuous Brent spot series

`brent_spot.py` combines the official FRED `DCOILBRENTEU` spot series with
Yahoo Finance `BZ=F` closes. FRED observations take priority. Level-adjusted
Yahoo values fill missing FRED trading days and extend the series after the
last common trading date.

```python
from brent_spot import get_continuous_brent_spot

brent = get_continuous_brent_spot(
    start="2020-01-01",
    adjustment="rolling",  # or "constant"
    rolling_window=20,
)
```

Provider downloads are cached under `.cache/brent`. After the first complete
download, refreshes retrieve only the latest 45 calendar days and merge any
FRED revisions into the local history. The dashboard caches FRED for six
hours, but refreshes Yahoo's recent tail on every online run. GitHub Actions
persists the historical cache between scheduled runs.
Run the tests and comparison plot with:

```bash
python -m unittest discover -s tests -v
python -m examples.plot_brent_spot
```
