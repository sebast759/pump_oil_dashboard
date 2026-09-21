from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import pandas as pd

import investing_brent


class InvestingBrentParserTests(unittest.TestCase):
    def test_extracts_pair_id_from_script(self) -> None:
        html = '<html><title>Brent Oil</title><script>window.x={"pair_ID":"8833"}</script></html>'
        self.assertEqual(investing_brent.extract_instrument_id(html), 8833)

    def test_extracts_pair_id_from_iframe(self) -> None:
        html = '<html><body>Brent Oil<iframe src="https://example.test/chart?pair_ID=12345"></iframe></body></html>'
        self.assertEqual(investing_brent.extract_instrument_id(html), 12345)

    def test_rejects_ambiguous_ids(self) -> None:
        html = '<html>Brent Oil<script>{"pair_ID":1,"instrumentId":2}</script></html>'
        with self.assertRaisesRegex(investing_brent.InvestingDataError, "Ambiguous"):
            investing_brent.extract_instrument_id(html)

    def test_parses_chart_rows_as_utc(self) -> None:
        frame = investing_brent.parse_chart_json(
            {"data": [[1_725_235_200_000, 78.1, 79.2, 77.5, 78.8, 12345, "ignored"]]}
        )
        self.assertEqual(list(frame.columns), ["timestamp", "open", "high", "low", "close", "volume"])
        self.assertEqual(str(frame.loc[0, "timestamp"].tz), "UTC")
        self.assertEqual(frame.loc[0, "close"], 78.8)

    def test_rejects_invalid_or_empty_payload(self) -> None:
        for payload in ({}, {"data": []}, {"data": [[1, 2, 3]]}, {"data": [[1, 2, 3, 4, "x", 6]]}):
            with self.subTest(payload=payload), self.assertRaises(investing_brent.InvestingDataError):
                investing_brent.parse_chart_json(payload)

    def test_changed_overlap_triggers_full_refresh(self) -> None:
        old = pd.DataFrame({
            "timestamp": pd.to_datetime(["2026-09-17", "2026-09-18"], utc=True),
            "open": [1, 1], "high": [2, 2], "low": [0.5, 0.5],
            "close": [103.0, 104.0], "volume": [10, 10],
        })
        recent = old.copy()
        recent.loc[0, "close"] = 103.5
        full = old.copy()
        full.loc[0, "close"] = 103.5
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "cache.csv"
            old.to_csv(path, index=False)
            with (
                patch("investing_brent.fetch_chart_page", return_value='<html>Brent {"pair_ID":8833}</html>'),
                patch("investing_brent.fetch_history", side_effect=[recent, full]) as fetch,
            ):
                _, result, refreshed = investing_brent.update_history_cache(path, request_pause=0)
        self.assertTrue(refreshed)
        self.assertEqual(fetch.call_args_list[-1].kwargs["period"], "MAX")
        self.assertEqual(result.loc[0, "close"], 103.5)


if __name__ == "__main__":
    unittest.main()
