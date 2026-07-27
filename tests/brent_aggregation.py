#!/usr/bin/env python3
"""Run the complete Brent aggregation analysis and generate its report.

Usage:
    python tests/brent_aggregation.py

Default outputs:
    tests/reports/brent_aggregation_report.html
    tests/reports/brent_aggregation_backtest.csv
    tests/reports/brent_sensitivity_by_country.csv

The single command above:
    1. Runs the built in regression checks.
    2. Loads weekly EC diesel and Euro 95 prices.
    3. Loads the cached continuous daily Brent series.
    4. Compares the latest print with 3, 5 and 7 trading day means.
    5. Fits rising and falling directional OLS sensitivities.
    6. Evaluates every method on a chronological holdout period.
    7. Produces console tables, CSV tables, charts and a standalone HTML report.

Use ``python tests/brent_aggregation.py --help`` for optional input, window and
output overrides.
"""

from __future__ import annotations

import argparse
import base64
import html
import io
import logging
import math
import sys
from dataclasses import dataclass
from datetime import timedelta
from pathlib import Path
from typing import Iterable

import pandas as pd
import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from brent_spot import get_continuous_brent_spot
from generate_oil_dashboard import extract_data

LOGGER = logging.getLogger("brent-aggregation-test")
FUELS = ("diesel", "euro95")
REPORT_DIR = Path(__file__).resolve().parent / "reports"
DEFAULT_WORKBOOK = (
    PROJECT_ROOT / "Weekly_Oil_Bulletin_Prices_History_maticni_4web.xlsx"
)
DEFAULT_CACHE_DIR = PROJECT_ROOT / ".cache" / "brent"
DEFAULT_RESULTS_CSV = REPORT_DIR / "brent_aggregation_backtest.csv"
DEFAULT_SENSITIVITY_CSV = REPORT_DIR / "brent_sensitivity_by_country.csv"
DEFAULT_HTML_REPORT = REPORT_DIR / "brent_aggregation_report.html"


@dataclass(frozen=True)
class BacktestResult:
    """Holdout metrics for one Brent aggregation method and fuel."""

    observations: int
    fuel: str
    train_rows: int
    test_rows: int
    up_slope: float
    down_slope: float
    mae_cents: float
    rmse_cents: float
    direction_accuracy: float


def aggregate_weekly_brent(
    daily: pd.Series,
    weekly_dates: Iterable[str | pd.Timestamp],
    observations: int,
) -> pd.Series:
    """Return Brent values aligned to weekly dates.

    Args:
        daily: Daily trading observations with a DatetimeIndex.
        weekly_dates: EC weekly observation dates.
        observations: Number of latest trading observations to average.
            ``1`` reproduces the dashboard's latest-print method.
    """
    if observations < 1:
        raise ValueError("observations must be at least 1")
    series = daily.dropna().sort_index()
    if series.empty:
        raise ValueError("daily Brent series is empty")

    values: list[float] = []
    index: list[pd.Timestamp] = []
    for raw_date in weekly_dates:
        weekly_date = pd.Timestamp(raw_date).normalize()
        history = series.loc[:weekly_date].tail(observations)
        values.append(float(history.mean()) if len(history) == observations else math.nan)
        index.append(weekly_date)
    return pd.Series(values, index=index, name=f"Brent mean {observations}")


def build_change_samples(
    weekly_brent: pd.Series,
    countries: dict,
    dates: list[str],
    fuel: str,
    *,
    change_window: int = 4,
    lag: int = 1,
) -> pd.DataFrame:
    """Build country-week Brent and pump-price change observations."""
    rows: list[dict[str, object]] = []
    brent = weekly_brent.to_numpy()
    for country, country_data in countries.items():
        pump = country_data[fuel]
        for position in range(change_window + lag, len(dates)):
            values = (
                brent[position - lag],
                brent[position - lag - change_window],
                pump[position],
                pump[position - change_window],
            )
            if any(pd.isna(value) for value in values):
                continue
            rows.append(
                {
                    "date": pd.Timestamp(dates[position]),
                    "country": country,
                    "brent_change": float(values[0] - values[1]),
                    "pump_change": float(values[2] - values[3]),
                }
            )
    return pd.DataFrame(rows)


def fit_directional_slopes(samples: pd.DataFrame) -> tuple[float, float]:
    """Fit through-origin slopes separately for Brent rises and declines."""

    def slope(frame: pd.DataFrame) -> float:
        denominator = float((frame["brent_change"] ** 2).sum())
        if denominator == 0:
            return 0.0
        numerator = float(
            (frame["brent_change"] * frame["pump_change"]).sum()
        )
        return numerator / denominator

    return (
        slope(samples[samples["brent_change"] > 0]),
        slope(samples[samples["brent_change"] < 0]),
    )


def evaluate_method(
    samples: pd.DataFrame,
    observations: int,
    fuel: str,
    train_fraction: float,
) -> BacktestResult:
    """Fit on the early period and evaluate on the chronological holdout."""
    unique_dates = sorted(samples["date"].unique())
    split_position = max(1, min(len(unique_dates) - 1, int(len(unique_dates) * train_fraction)))
    cutoff = unique_dates[split_position]
    train = samples[samples["date"] < cutoff].copy()
    test = samples[samples["date"] >= cutoff].copy()
    if train.empty or test.empty:
        raise ValueError("not enough observations for chronological holdout")

    up_slope, down_slope = fit_directional_slopes(train)
    slopes = test["brent_change"].map(
        lambda value: up_slope if value >= 0 else down_slope
    )
    predicted = test["brent_change"] * slopes
    error = predicted - test["pump_change"]
    actual_direction = test["pump_change"].map(lambda value: 1 if value > 0 else -1 if value < 0 else 0)
    predicted_direction = predicted.map(lambda value: 1 if value > 0 else -1 if value < 0 else 0)

    return BacktestResult(
        observations=observations,
        fuel=fuel,
        train_rows=len(train),
        test_rows=len(test),
        up_slope=up_slope,
        down_slope=down_slope,
        mae_cents=float(error.abs().mean() / 10),
        rmse_cents=float(math.sqrt((error ** 2).mean()) / 10),
        direction_accuracy=float((actual_direction == predicted_direction).mean()),
    )


def run_regression_checks() -> None:
    """Run fast checks that protect date alignment and holdout calculations."""
    daily = pd.Series(
        [70.0, 72.0, 74.0, 100.0],
        index=pd.to_datetime(
            ["2026-01-01", "2026-01-02", "2026-01-05", "2026-01-06"]
        ),
    )
    aligned = aggregate_weekly_brent(daily, ["2026-01-05"], 3)
    if not math.isclose(float(aligned.iloc[0]), 72.0):
        raise RuntimeError(
            "Regression check failed: trailing mean used a future observation"
        )

    dates = pd.date_range("2025-01-01", periods=20, freq="W-MON")
    changes = [float(value) for value in range(-10, 10)]
    samples = pd.DataFrame(
        {
            "date": dates,
            "country": ["FR"] * len(dates),
            "brent_change": changes,
            "pump_change": [value * 5.0 for value in changes],
        }
    )
    result = evaluate_method(samples, 3, "diesel", 0.7)
    expected = (
        math.isclose(result.up_slope, 5.0)
        and math.isclose(result.down_slope, 5.0)
        and math.isclose(result.mae_cents, 0.0, abs_tol=1e-12)
        and math.isclose(result.rmse_cents, 0.0, abs_tol=1e-12)
    )
    if not expected:
        raise RuntimeError(
            "Regression check failed: holdout sensitivity calculation changed"
        )
    LOGGER.info("Built in regression checks passed")


def run_backtest(
    xlsx_path: Path,
    cache_dir: Path,
    windows: list[int],
    train_fraction: float,
) -> tuple[pd.DataFrame, pd.DataFrame, dict[int, pd.Series]]:
    """Run every aggregation window against both dashboard fuel series."""
    if not 0.5 <= train_fraction < 1:
        raise ValueError("train_fraction must be between 0.5 and 1")
    data = extract_data(xlsx_path, local=True)
    daily = get_continuous_brent_spot(
        start=pd.Timestamp(data["dates"][0]) - pd.Timedelta(days=21),
        end=pd.Timestamp(data["dates"][-1]),
        cache_dir=cache_dir,
        cache_ttl=timedelta(days=36500),
        yahoo_cache_ttl=timedelta(days=36500),
    )

    results: list[BacktestResult] = []
    sensitivity_rows: list[dict[str, object]] = []
    weekly_series: dict[int, pd.Series] = {}
    for observations in windows:
        weekly = aggregate_weekly_brent(daily, data["dates"], observations)
        weekly_series[observations] = weekly
        for fuel in FUELS:
            samples = build_change_samples(
                weekly,
                data["countries"],
                data["dates"],
                fuel,
            )
            results.append(
                evaluate_method(samples, observations, fuel, train_fraction)
            )
            for country, country_samples in samples.groupby("country"):
                up_slope, down_slope = fit_directional_slopes(country_samples)
                sensitivity_rows.append(
                    {
                        "observations": observations,
                        "fuel": fuel,
                        "country": country,
                        "up_slope": up_slope,
                        "down_slope": down_slope,
                        "rows": len(country_samples),
                    }
                )
    return (
        pd.DataFrame(result.__dict__ for result in results),
        pd.DataFrame(sensitivity_rows),
        weekly_series,
    )


def print_report(results: pd.DataFrame) -> None:
    """Print detailed results and a cross-fuel ranking."""
    display = results.copy()
    display["method"] = display["observations"].map(
        lambda value: "latest print" if value == 1 else f"{value} day mean"
    )
    display["direction_accuracy"] *= 100
    columns = [
        "method",
        "fuel",
        "up_slope",
        "down_slope",
        "mae_cents",
        "rmse_cents",
        "direction_accuracy",
        "test_rows",
    ]
    print("\nChronological holdout results")
    print(
        display[columns].to_string(
            index=False,
            formatters={
                "up_slope": "{:.2f}".format,
                "down_slope": "{:.2f}".format,
                "mae_cents": "{:.2f}".format,
                "rmse_cents": "{:.2f}".format,
                "direction_accuracy": "{:.1f}%".format,
            },
        )
    )

    ranking = (
        results.groupby("observations", as_index=False)
        .agg(
            mean_mae_cents=("mae_cents", "mean"),
            mean_rmse_cents=("rmse_cents", "mean"),
            mean_direction_accuracy=("direction_accuracy", "mean"),
        )
        .sort_values(["mean_mae_cents", "mean_rmse_cents"])
    )
    ranking["method"] = ranking["observations"].map(
        lambda value: "latest print" if value == 1 else f"{value} day mean"
    )
    ranking["mean_direction_accuracy"] *= 100
    print("\nCross-fuel ranking")
    print(
        ranking[
            ["method", "mean_mae_cents", "mean_rmse_cents", "mean_direction_accuracy"]
        ].to_string(
            index=False,
            formatters={
                "mean_mae_cents": "{:.2f}".format,
                "mean_rmse_cents": "{:.2f}".format,
                "mean_direction_accuracy": "{:.1f}%".format,
            },
        )
    )
    winner = ranking.iloc[0]
    print(
        f"\nLowest holdout MAE: {winner['method']} "
        f"({winner['mean_mae_cents']:.2f} cents/L)."
    )


def method_name(observations: int) -> str:
    """Return a concise display name for an aggregation window."""
    return "Latest print" if observations == 1 else f"{observations} day mean"


def figure_data_uri(figure: plt.Figure) -> str:
    """Encode a Matplotlib figure as an inline PNG data URI."""
    buffer = io.BytesIO()
    figure.savefig(buffer, format="png", dpi=150, bbox_inches="tight")
    plt.close(figure)
    encoded = base64.b64encode(buffer.getvalue()).decode("ascii")
    return f"data:image/png;base64,{encoded}"


def plot_error_metrics(results: pd.DataFrame) -> str:
    """Plot holdout MAE and RMSE for every method and fuel."""
    frame = results.copy()
    frame["method"] = frame["observations"].map(method_name)
    labels = [
        f"{method}\n{fuel.title()}"
        for method, fuel in zip(frame["method"], frame["fuel"])
    ]
    positions = list(range(len(frame)))
    figure, axis = plt.subplots(figsize=(11, 5.2))
    axis.bar(
        [value - 0.2 for value in positions],
        frame["mae_cents"],
        width=0.4,
        label="MAE",
        color="#2563eb",
    )
    axis.bar(
        [value + 0.2 for value in positions],
        frame["rmse_cents"],
        width=0.4,
        label="RMSE",
        color="#94a3b8",
    )
    axis.set_xticks(positions, labels)
    axis.set_ylabel("Holdout error, cents/L")
    axis.set_title("Forecast error by Brent aggregation method")
    axis.grid(axis="y", alpha=0.2)
    axis.legend()
    figure.tight_layout()
    return figure_data_uri(figure)


def plot_direction_accuracy(results: pd.DataFrame) -> str:
    """Plot cross-fuel mean direction accuracy by aggregation method."""
    summary = (
        results.groupby("observations", as_index=False)["direction_accuracy"]
        .mean()
        .sort_values("observations")
    )
    labels = [method_name(value) for value in summary["observations"]]
    values = summary["direction_accuracy"] * 100
    figure, axis = plt.subplots(figsize=(9, 4.8))
    bars = axis.bar(labels, values, color=["#2563eb", "#0f766e", "#d97706", "#7c3aed"][:len(labels)])
    axis.set_ylabel("Correct direction, %")
    axis.set_title("Direction accuracy on the chronological holdout")
    axis.set_ylim(max(0, float(values.min()) - 8), min(100, float(values.max()) + 6))
    axis.grid(axis="y", alpha=0.2)
    axis.bar_label(bars, labels=[f"{value:.1f}%" for value in values], padding=3)
    figure.tight_layout()
    return figure_data_uri(figure)


def plot_sensitivity(
    sensitivity: pd.DataFrame,
    fuel: str,
) -> str:
    """Plot average rising and falling sensitivities for one fuel."""
    summary = (
        sensitivity[sensitivity["fuel"] == fuel]
        .groupby("observations", as_index=False)[["up_slope", "down_slope"]]
        .mean()
        .sort_values("observations")
    )
    labels = [method_name(value) for value in summary["observations"]]
    positions = list(range(len(summary)))
    figure, axis = plt.subplots(figsize=(9, 4.8))
    axis.bar(
        [value - 0.2 for value in positions],
        summary["up_slope"],
        width=0.4,
        label="Brent rose",
        color="#f97316",
    )
    axis.bar(
        [value + 0.2 for value in positions],
        summary["down_slope"],
        width=0.4,
        label="Brent fell",
        color="#fdba74",
        alpha=0.65,
    )
    axis.set_xticks(positions, labels)
    axis.set_ylabel("Pump move, cents/L per $10 Brent")
    axis.set_title(f"{fuel.title()} pass through sensitivity")
    axis.grid(axis="y", alpha=0.2)
    axis.legend()
    figure.tight_layout()
    return figure_data_uri(figure)


def plot_weekly_brent(weekly_series: dict[int, pd.Series]) -> str:
    """Plot the recent Brent signals produced by each aggregation method."""
    figure, axis = plt.subplots(figsize=(11, 5.2))
    for observations, series in sorted(weekly_series.items()):
        recent = series.dropna().tail(104)
        axis.plot(
            recent.index,
            recent.values,
            label=method_name(observations),
            linewidth=1.8 if observations == 1 else 1.3,
        )
    axis.set_ylabel("USD per barrel")
    axis.set_title("Weekly Brent inputs, latest two years")
    axis.grid(alpha=0.2)
    axis.legend(ncol=2)
    figure.autofmt_xdate()
    figure.tight_layout()
    return figure_data_uri(figure)


def dataframe_html(
    frame: pd.DataFrame,
    *,
    formats: dict[str, str] | None = None,
) -> str:
    """Render a compact, escaped HTML table with optional numeric formats."""
    display = frame.copy()
    for column, format_string in (formats or {}).items():
        if column in display:
            display[column] = display[column].map(
                lambda value: format_string.format(value)
            )
    return display.to_html(index=False, border=0, classes="report-table", escape=True)


def write_html_report(
    results: pd.DataFrame,
    sensitivity: pd.DataFrame,
    weekly_series: dict[int, pd.Series],
    destination: Path,
    train_fraction: float,
) -> None:
    """Write a self-contained HTML report with tables and inline charts."""
    performance = results.copy()
    performance.insert(0, "method", performance["observations"].map(method_name))
    performance["fuel"] = performance["fuel"].map(
        {"diesel": "Diesel", "euro95": "Euro 95"}
    )
    performance["direction_accuracy"] *= 100
    performance = performance[
        [
            "method",
            "fuel",
            "up_slope",
            "down_slope",
            "mae_cents",
            "rmse_cents",
            "direction_accuracy",
            "train_rows",
            "test_rows",
        ]
    ]

    ranking = (
        results.groupby("observations", as_index=False)
        .agg(
            mean_mae_cents=("mae_cents", "mean"),
            mean_rmse_cents=("rmse_cents", "mean"),
            direction_accuracy=("direction_accuracy", "mean"),
        )
        .sort_values(["mean_mae_cents", "mean_rmse_cents"])
    )
    ranking.insert(0, "method", ranking["observations"].map(method_name))
    ranking["direction_accuracy"] *= 100
    winner = ranking.iloc[0]
    direction_winner = ranking.sort_values(
        "direction_accuracy", ascending=False
    ).iloc[0]

    sensitivity_summary = (
        sensitivity.groupby(["observations", "fuel"], as_index=False)
        .agg(
            up_slope=("up_slope", "mean"),
            down_slope=("down_slope", "mean"),
        )
    )
    sensitivity_summary.insert(
        0, "method", sensitivity_summary["observations"].map(method_name)
    )
    sensitivity_summary["fuel"] = sensitivity_summary["fuel"].map(
        {"diesel": "Diesel", "euro95": "Euro 95"}
    )
    sensitivity_detail = sensitivity.copy()
    sensitivity_detail.insert(
        0, "method", sensitivity_detail["observations"].map(method_name)
    )
    sensitivity_detail["fuel"] = sensitivity_detail["fuel"].map(
        {"diesel": "Diesel", "euro95": "Euro 95"}
    )

    charts = {
        "errors": plot_error_metrics(results),
        "direction": plot_direction_accuracy(results),
        "diesel": plot_sensitivity(sensitivity, "diesel"),
        "euro95": plot_sensitivity(sensitivity, "euro95"),
        "brent": plot_weekly_brent(weekly_series),
    }
    generated = pd.Timestamp.now().strftime("%Y-%m-%d %H:%M")
    train_percent = int(round(train_fraction * 100))
    test_percent = 100 - train_percent
    ranking_table = dataframe_html(
        ranking[
            ["method", "mean_mae_cents", "mean_rmse_cents", "direction_accuracy"]
        ],
        formats={
            "mean_mae_cents": "{:.2f}",
            "mean_rmse_cents": "{:.2f}",
            "direction_accuracy": "{:.1f}%",
        },
    )
    performance_table = dataframe_html(
        performance,
        formats={
            "up_slope": "{:.2f}",
            "down_slope": "{:.2f}",
            "mae_cents": "{:.2f}",
            "rmse_cents": "{:.2f}",
            "direction_accuracy": "{:.1f}%",
        },
    )
    sensitivity_summary_table = dataframe_html(
        sensitivity_summary[["method", "fuel", "up_slope", "down_slope"]],
        formats={"up_slope": "{:.2f}", "down_slope": "{:.2f}"},
    )
    sensitivity_detail_table = dataframe_html(
        sensitivity_detail[
            ["method", "fuel", "country", "up_slope", "down_slope", "rows"]
        ],
        formats={"up_slope": "{:.2f}", "down_slope": "{:.2f}"},
    )
    report = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Brent Aggregation Backtest Report</title>
<style>
body {{ margin:0; background:#f8fafc; color:#172033; font:15px/1.6 Arial,sans-serif; }}
main {{ max-width:1180px; margin:0 auto; padding:32px 24px 60px; }}
h1 {{ font-size:30px; margin:0 0 4px; }}
h2 {{ margin-top:34px; padding-top:20px; border-top:1px solid #cbd5e1; }}
h3 {{ margin-top:24px; }}
.muted {{ color:#64748b; }}
.summary {{ background:#e8f1ff; border-left:5px solid #2563eb; padding:16px 18px; margin:22px 0; }}
.note {{ background:#fff7ed; border-left:5px solid #f97316; padding:14px 18px; }}
.chart {{ width:100%; background:#fff; border:1px solid #e2e8f0; border-radius:10px; margin:12px 0 24px; }}
.grid {{ display:grid; grid-template-columns:1fr 1fr; gap:20px; }}
.report-table {{ width:100%; border-collapse:collapse; font-size:13px; background:#fff; }}
.report-table th {{ background:#e2e8f0; text-align:right; padding:8px; white-space:nowrap; }}
.report-table td {{ border-bottom:1px solid #e2e8f0; text-align:right; padding:7px 8px; }}
.report-table th:first-child,.report-table td:first-child {{ text-align:left; }}
.table-wrap {{ overflow-x:auto; border:1px solid #e2e8f0; border-radius:8px; }}
code {{ background:#e2e8f0; padding:2px 5px; border-radius:4px; }}
@media(max-width:800px) {{ .grid {{ grid-template-columns:1fr; }} }}
</style>
</head>
<body><main>
<h1>Brent aggregation backtest</h1>
<div class="muted">Generated {html.escape(generated)} · Weekly EC pump prices · Daily continuous Brent spot</div>
<div class="summary">
<strong>Result:</strong> {html.escape(str(winner["method"]))} has the lowest average
holdout MAE at {winner["mean_mae_cents"]:.2f} cents/L.
{html.escape(str(direction_winner["method"]))} has the best direction accuracy at
{direction_winner["direction_accuracy"]:.1f}%.
</div>

<h2>Question tested</h2>
<p>Does replacing the latest Brent print available on each EC weekly date with
an average of the latest 3, 5 or 7 trading observations improve the pump price
signal?</p>
<p>The first {train_percent}% of weekly dates fit the directional slopes. The
final {test_percent}% form a chronological holdout and are never used to fit
the model. The pump response uses a 4 week change and Brent uses a 1 week lag,
matching the dashboard sensitivity method.</p>

<h2>Overall ranking</h2>
<div class="table-wrap">{ranking_table}</div>
<div class="grid">
  <img class="chart" src="{charts["errors"]}" alt="Forecast error comparison">
  <img class="chart" src="{charts["direction"]}" alt="Direction accuracy comparison">
</div>

<h2>Detailed holdout performance</h2>
<div class="table-wrap">{performance_table}</div>

<h2>Sensitivity estimates</h2>
<div class="note"><strong>How to read:</strong> Each bar is the OLS slope of the
4 week fuel change against the 4 week Brent change with a 1 week lag, split by
the direction of the Brent move. The solid bar represents weeks when Brent rose.
The light bar represents weeks when Brent fell. A taller solid bar than light
bar indicates asymmetric pass through: suppliers are quicker to pass on crude
price increases than to adjust downward.</div>
<div class="grid">
  <img class="chart" src="{charts["diesel"]}" alt="Diesel sensitivity">
  <img class="chart" src="{charts["euro95"]}" alt="Euro 95 sensitivity">
</div>
<h3>Average sensitivity across countries</h3>
<div class="table-wrap">{sensitivity_summary_table}</div>
<h3>Country sensitivity detail</h3>
<div class="table-wrap">{sensitivity_detail_table}</div>

<h2>How the Brent inputs differ</h2>
<img class="chart" src="{charts["brent"]}" alt="Weekly Brent aggregation comparison">
<p>The averages damp individual trading sessions. This may improve direction
stability, but it can also delay or reduce the measured size of a rapid move.</p>

<h2>Interpretation limits</h2>
<ul>
  <li>This is a historical comparison, not proof that one method will remain best.</li>
  <li>Country observations share the same Brent series and are not fully independent.</li>
  <li>Taxes, exchange rates, refining margins and local competition are not modelled separately.</li>
  <li>MAE measures the size error. Direction accuracy only measures rise, fall or flat.</li>
</ul>
</main></body></html>"""
    destination.parent.mkdir(parents=True, exist_ok=True)
    destination.write_text(report, encoding="utf-8")


def parse_args() -> argparse.Namespace:
    """Parse optional overrides for the complete report generation run."""
    parser = argparse.ArgumentParser(
        description=(
            "Calculate Brent aggregation sensitivities and holdout metrics, "
            "then generate the complete CSV and HTML report."
        ),
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument(
        "xlsx",
        nargs="?",
        default=DEFAULT_WORKBOOK,
        type=Path,
        help="EC Weekly Oil Bulletin workbook",
    )
    parser.add_argument(
        "--windows",
        nargs="+",
        type=int,
        default=[1, 3, 5, 7],
        help="Trading-observation counts to compare",
    )
    parser.add_argument(
        "--train-fraction",
        type=float,
        default=0.7,
        help="Fraction of early dates used to fit slopes",
    )
    parser.add_argument(
        "--cache-dir",
        type=Path,
        default=DEFAULT_CACHE_DIR,
        help="Brent provider cache directory",
    )
    parser.add_argument(
        "--csv",
        type=Path,
        default=DEFAULT_RESULTS_CSV,
        help="Path for detailed holdout results",
    )
    parser.add_argument(
        "--sensitivity-csv",
        type=Path,
        default=DEFAULT_SENSITIVITY_CSV,
        help="Path for country sensitivity estimates",
    )
    parser.add_argument(
        "--report",
        type=Path,
        default=DEFAULT_HTML_REPORT,
        help="Path for the self-contained HTML report",
    )
    return parser.parse_args()


def main() -> None:
    """Run all calculations and generate every report artifact."""
    args = parse_args()
    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
    LOGGER.info("Running complete Brent aggregation analysis")
    run_regression_checks()
    try:
        results, sensitivity, weekly_series = run_backtest(
            args.xlsx,
            args.cache_dir,
            sorted(set(args.windows)),
            args.train_fraction,
        )
    except (FileNotFoundError, ValueError, OSError) as exc:
        raise SystemExit(f"Backtest failed: {exc}") from exc
    print_report(results)
    if args.csv:
        args.csv.parent.mkdir(parents=True, exist_ok=True)
        results.to_csv(args.csv, index=False)
        LOGGER.info("Saved detailed results to %s", args.csv)
    if args.sensitivity_csv:
        args.sensitivity_csv.parent.mkdir(parents=True, exist_ok=True)
        sensitivity.to_csv(args.sensitivity_csv, index=False)
        LOGGER.info("Saved sensitivity results to %s", args.sensitivity_csv)
    if args.report:
        write_html_report(
            results,
            sensitivity,
            weekly_series,
            args.report,
            args.train_fraction,
        )
        LOGGER.info("Saved HTML report to %s", args.report)
    LOGGER.info("Analysis complete")


if __name__ == "__main__":
    main()
