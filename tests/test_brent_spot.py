from __future__ import annotations

import tempfile
import unittest
from datetime import timedelta
from pathlib import Path
from unittest.mock import patch

import pandas as pd

import brent_spot


def _series(values: dict[str, float], name: str) -> pd.Series:
    series = pd.Series(values, dtype=float, name=name)
    series.index = pd.to_datetime(series.index)
    return series


class ContinuousBrentSpotTests(unittest.TestCase):
    def setUp(self) -> None:
        self.official = _series(
            {
                "2026-07-16": 88.0,
                "2026-07-17": 90.0,
                "2026-07-20": 100.0,
            },
            "Official Brent spot",
        )
        self.yahoo = _series(
            {
                "2026-07-16": 86.0,
                "2026-07-17": 90.0,
                "2026-07-20": 80.0,
                "2026-07-21": 84.0,
                "2026-07-22": 88.0,
            },
            "Yahoo Brent futures",
        )
        self.temporary_directory = tempfile.TemporaryDirectory()
        self.cache_dir = Path(self.temporary_directory.name)

    def tearDown(self) -> None:
        self.temporary_directory.cleanup()

    def _provider_patches(self):
        return (
            patch("brent_spot._download_fred_spot", return_value=self.official),
            patch(
                "brent_spot._download_yahoo_futures",
                return_value=self.yahoo,
            ),
        )

    def test_constant_adjustment_and_official_priority(self) -> None:
        fred_patch, yahoo_patch = self._provider_patches()
        with fred_patch, yahoo_patch:
            result = brent_spot.get_continuous_brent_spot(
                end="2026-07-22",
                cache_dir=self.cache_dir,
                refresh=True,
            )

        self.assertEqual(result.name, "Brent Spot")
        self.assertTrue(result.index.is_unique)
        self.assertAlmostEqual(result.loc["2026-07-20"], 100.0)
        # Final overlap factor = 100 / 80 = 1.25.
        self.assertAlmostEqual(result.loc["2026-07-21"], 105.0)
        self.assertAlmostEqual(result.loc["2026-07-22"], 110.0)
        # Weekend dates are not manufactured.
        self.assertNotIn(pd.Timestamp("2026-07-18"), result.index)

    def test_rolling_adjustment(self) -> None:
        fred_patch, yahoo_patch = self._provider_patches()
        with fred_patch, yahoo_patch:
            result = brent_spot.get_continuous_brent_spot(
                end="2026-07-22",
                adjustment="rolling",
                rolling_window=2,
                cache_dir=self.cache_dir,
                refresh=True,
            )

        # Mean of the last two ratios: mean(90/90, 100/80) = 1.125.
        self.assertAlmostEqual(result.loc["2026-07-21"], 94.5)
        self.assertAlmostEqual(result.loc["2026-07-22"], 99.0)

    def test_adjusted_yahoo_fills_internal_official_gap(self) -> None:
        official = _series(
            {
                "2026-05-22": 106.90,
                "2026-05-26": 102.75,
            },
            "Official Brent spot",
        )
        yahoo = _series(
            {
                "2026-05-22": 100.00,
                "2026-05-25": 98.00,
                "2026-05-26": 97.50,
            },
            "Yahoo Brent futures",
        )
        with (
            patch("brent_spot._download_fred_spot", return_value=official),
            patch("brent_spot._download_yahoo_futures", return_value=yahoo),
        ):
            result = brent_spot.get_continuous_brent_spot(
                end="2026-05-26",
                cache_dir=self.cache_dir,
                refresh=True,
            )

        # The 22 May factor (106.90 / 100) is the latest factor available on
        # the missing 25 May date. The official 26 May value remains untouched.
        self.assertAlmostEqual(result.loc["2026-05-25"], 98.0 * 1.069)
        self.assertAlmostEqual(result.loc["2026-05-26"], 102.75)

    def test_fresh_cache_avoids_download(self) -> None:
        fred_patch, yahoo_patch = self._provider_patches()
        with fred_patch, yahoo_patch:
            expected = brent_spot.get_continuous_brent_spot(
                end="2026-07-22",
                cache_dir=self.cache_dir,
                refresh=True,
            )

        with (
            patch(
                "brent_spot._download_fred_spot",
                side_effect=AssertionError("FRED should not be downloaded"),
            ),
            patch(
                "brent_spot._download_yahoo_futures",
                side_effect=AssertionError("Yahoo should not be downloaded"),
            ),
        ):
            cached = brent_spot.get_continuous_brent_spot(
                end="2026-07-22",
                cache_dir=self.cache_dir,
                cache_ttl=timedelta(days=1),
                yahoo_cache_ttl=timedelta(days=1),
            )
        pd.testing.assert_series_equal(cached, expected, check_freq=False)

    def test_yahoo_refreshes_by_default_even_when_fred_cache_is_fresh(self) -> None:
        fred_patch, yahoo_patch = self._provider_patches()
        with fred_patch, yahoo_patch:
            brent_spot.get_continuous_brent_spot(
                end="2026-07-22",
                cache_dir=self.cache_dir,
                refresh=True,
            )

        with (
            patch(
                "brent_spot._download_fred_spot",
                side_effect=AssertionError("Fresh FRED cache should be reused"),
            ),
            patch(
                "brent_spot._download_yahoo_futures",
                return_value=self.yahoo,
            ) as yahoo_refresh,
        ):
            brent_spot.get_continuous_brent_spot(
                end="2026-07-22",
                cache_dir=self.cache_dir,
                cache_ttl=timedelta(days=1),
            )
        yahoo_refresh.assert_called_once()

    def test_rejects_non_overlapping_providers(self) -> None:
        official = _series(
            {"2026-07-20": 100.0},
            "Official Brent spot",
        )
        yahoo = _series(
            {"2026-07-21": 90.0},
            "Yahoo Brent futures",
        )
        with (
            patch("brent_spot._download_fred_spot", return_value=official),
            patch("brent_spot._download_yahoo_futures", return_value=yahoo),
            self.assertRaisesRegex(
                brent_spot.BrentDataError,
                "no common trading dates",
            ),
        ):
            brent_spot.get_continuous_brent_spot(
                end="2026-07-22",
                cache_dir=self.cache_dir,
                refresh=True,
            )

    def test_rejects_invalid_arguments(self) -> None:
        with self.assertRaisesRegex(ValueError, "adjustment"):
            brent_spot.get_continuous_brent_spot(adjustment="bad")  # type: ignore[arg-type]
        with self.assertRaisesRegex(ValueError, "rolling_window"):
            brent_spot.get_continuous_brent_spot(rolling_window=0)
        with self.assertRaisesRegex(ValueError, "yahoo_cache_ttl"):
            brent_spot.get_continuous_brent_spot(
                yahoo_cache_ttl=timedelta(seconds=-1)
            )
        with self.assertRaisesRegex(ValueError, "start"):
            brent_spot.get_continuous_brent_spot(
                start="2026-07-23",
                end="2026-07-22",
            )


if __name__ == "__main__":
    unittest.main()
