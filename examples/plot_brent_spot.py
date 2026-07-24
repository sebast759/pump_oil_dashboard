"""Plot official, raw-futures, and adjusted continuous Brent prices."""

from __future__ import annotations

import logging

import matplotlib.pyplot as plt
import pandas as pd
import yfinance as yf

from brent_spot import get_continuous_brent_spot


def main() -> None:
    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
    start = "2025-01-01"
    continuous = get_continuous_brent_spot(
        start=start,
        adjustment="rolling",
        rolling_window=20,
    )

    official = pd.read_csv(
        " with sources",
        parse_dates=["observation_date"],
        index_col="observation_date",
    )["DCOILBRENTEU"].pipe(pd.to_numeric, errors="coerce").loc[start:]

    raw_yahoo = yf.download(
        "BZ=F",
        start=start,
        auto_adjust=False,
        progress=False,
    )["Close"]
    if isinstance(raw_yahoo, pd.DataFrame):
        raw_yahoo = raw_yahoo.iloc[:, 0]

    ax = official.plot(figsize=(12, 6), label="Official spot (FRED)", linewidth=2)
    raw_yahoo.plot(ax=ax, label="Raw Yahoo BZ=F", alpha=0.65)
    continuous.plot(
        ax=ax,
        label="Adjusted continuous spot",
        linewidth=2,
        linestyle="--",
    )
    ax.set(title="Continuous Brent spot price", ylabel="USD per barrel", xlabel="")
    ax.grid(alpha=0.25)
    ax.legend()
    plt.tight_layout()
    plt.show()


if __name__ == "__main__":
    main()
