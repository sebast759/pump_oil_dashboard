# Fuel Forecast

Fuel Forecast tracks weekly diesel and Euro 95 pump prices across Europe and
uses daily Brent spot prices to provide an indicative next week signal.

**Live dashboard: [fuelforecast.eu](https://fuelforecast.eu/)**

## Run locally

Install the dependencies:

```powershell
pip install -r requirements.txt
```

Generate the complete local website in `site/` using fresh online data:

```powershell
python generate_oil_dashboard.py
```

Other useful modes:

```powershell
python generate_oil_dashboard.py --local
python generate_oil_dashboard.py --download
python generate_oil_dashboard.py --output preview.html
python generate_oil_dashboard.py --push
```

`--local` uses the cached workbook and Brent data without network calls.
Both normal and local runs write `site/index.html` and copy its required
images, favicons and discovery files into the same `site/` directory.
`--push` retains the optional legacy publishing path, but is not used by the
normal GitHub Pages deployment.

## Automated deployment

`.github/workflows/update_dashboard.yml` generates the site and deploys the
`site/` artifact directly to GitHub Pages. It runs daily at 14:30 UTC, also
runs Monday at 07:00 UTC, and can be started manually from the Actions tab.

## Brent data

`brent_spot.py` creates a continuous daily series from:

* official Brent spot observations from FRED `DCOILBRENTEU`
* adjusted Yahoo Finance `BZ=F` observations for missing and newer dates

FRED observations take priority. Downloads are cached under `.cache/brent`;
online runs refresh Yahoo's latest market data while avoiding a full history
download each time.

## Brent aggregation analysis

The research script compares the latest daily Brent print with trailing
3 day, 5 day and 7 day averages. It calculates directional OLS sensitivity,
holdout errors and country level estimates, then creates an HTML report and
CSV tables.

```powershell
python tests/brent_aggregation.py
```

Generated research outputs are written to `tests/reports/`. That directory is
ignored by Git because the files can be recreated.

The command also runs two built in regression checks before starting the full
analysis. They protect trailing date alignment and the holdout sensitivity
calculation, so no separate aggregation test script is required.

Run the complete test suite with:

```powershell
python -m unittest discover -s tests -v
```

## Project structure

```text
oil_dashboard/
|-- generate_oil_dashboard.py
|-- brent_spot.py
|-- requirements.txt
|-- assets/
|-- examples/
|-- tests/
|   |-- brent_aggregation.py
|   |-- test_brent_spot.py
|   `-- reports/                 generated and ignored
|-- .github/workflows/
|   `-- update_dashboard.yml
`-- site/                       generated local website, ignored by Git
```

The generated `site/` directory is a disposable build artifact and is excluded
from Git. Source images remain under `assets/`. The workbook, `.cache/`,
`brent_cache.csv` and `__pycache__/` are also local runtime files excluded from
Git.

## Data sources

* [European Commission Weekly Oil Bulletin](https://energy.ec.europa.eu/data-and-analysis/weekly-oil-bulletin_en)
* [FRED Brent Europe spot series](https://fred.stlouisfed.org/series/DCOILBRENTEU)
* [Yahoo Finance Brent futures](https://finance.yahoo.com/quote/BZ%3DF/)
