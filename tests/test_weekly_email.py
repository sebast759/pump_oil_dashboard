from __future__ import annotations

import unittest
from datetime import date
from unittest.mock import Mock, patch

import pandas as pd

import weekly_email


class AdviceTests(unittest.TestCase):
    def test_brent_rise_means_fill_up_before_monday(self) -> None:
        advice = weekly_email.advise("diesel", 5.0)
        self.assertEqual(advice.action, "fill_up")
        self.assertEqual(advice.cents, 5)          # 0.09 * 5 * 10 = 4.5, halves round up
        self.assertEqual(advice.tank_saving, 2.3)

    def test_brent_fall_means_wait(self) -> None:
        advice = weekly_email.advise("euro95", -6.0)
        self.assertEqual(advice.action, "wait")
        self.assertEqual(advice.cents, 3)          # 0.05 * -6 * 10 = -3.0

    def test_small_move_means_any_day(self) -> None:
        self.assertEqual(weekly_email.advise("diesel", 0.2).action, "any_day")

    def test_falling_coefficients_differ_from_rising(self) -> None:
        self.assertAlmostEqual(weekly_email.expected_change_cents("diesel", 1.0), 0.9)
        self.assertAlmostEqual(weekly_email.expected_change_cents("diesel", -1.0), -0.6)


class SignalTests(unittest.TestCase):
    def test_prefers_brent_latest(self) -> None:
        data = {"brent": [90.0, 91.0], "latest_date": "2026-09-14",
                "brent_latest": {"price": 99.0, "previous_price": 95.0, "date": "2026-09-21"}}
        self.assertEqual(weekly_email.signal_from_data(data),
                         {"brent_price": 99.0, "brent_previous": 95.0, "brent_date": "2026-09-21",
                          "this_week_move": None, "next_week_outlook": None})

    def test_falls_back_to_weekly_series(self) -> None:
        data = {"brent": [90.0, None, 92.0], "latest_date": "2026-09-14", "brent_latest": None}
        self.assertEqual(weekly_email.signal_from_data(data),
                         {"brent_price": 92.0, "brent_previous": 90.0, "brent_date": "2026-09-14",
                          "this_week_move": None, "next_week_outlook": None})

    def test_stale_data_are_not_sent(self) -> None:
        signal = {"brent_price": 99.0, "brent_previous": 95.0, "brent_date": "2026-09-01"}
        self.assertIn("days old", weekly_email.signal_problem(signal, today=date(2026, 9, 24)))

    def test_fresh_data_are_sent(self) -> None:
        signal = {"brent_price": 99.0, "brent_previous": 95.0, "brent_date": "2026-09-21"}
        self.assertIsNone(weekly_email.signal_problem(signal, today=date(2026, 9, 24)))


class EmailTests(unittest.TestCase):
    signal = {"brent_price": 105.0, "brent_previous": 100.0, "brent_date": "2026-09-24"}

    def test_rise_subject_and_body(self) -> None:
        subject, body = weekly_email.build_email(self.signal)
        self.assertEqual(subject, "Fuel tip: fill up before Monday")
        self.assertIn("Diesel: fill up before Monday", body)
        self.assertIn(weekly_email.SITE_URL, body)

    @patch("weekly_email.requests.post")
    def test_email_targets_only_this_sites_subscribers(self, post: Mock) -> None:
        post.return_value = Mock(ok=True, json=lambda: {"id": "x"})
        weekly_email.send_email("s", "b", "key")
        payload = post.call_args.kwargs["json"]
        self.assertEqual(payload["status"], "about_to_send")
        self.assertEqual(payload["filters"]["filters"][0], {
            "operator": "contains", "field": "subscriber.tags",
            "value": weekly_email.NEWSLETTER_SOURCE,
        })

    @patch("weekly_email.requests.post")
    def test_draft_is_not_sent(self, post: Mock) -> None:
        post.return_value = Mock(ok=True, json=lambda: {"id": "x"})
        weekly_email.send_email("s", "b", "key", draft=True)
        self.assertEqual(post.call_args.kwargs["json"]["status"], "draft")


class WeeklyBrentSignalTests(unittest.TestCase):
    def _dates_prices(self, start: str, values: list[float]) -> tuple[list[str], list[float]]:
        dates = pd.date_range(start, periods=len(values), freq="D")
        return [d.strftime("%Y-%m-%d") for d in dates], values

    def test_matches_the_station_rise_brent_fall_scenario(self) -> None:
        # Mon 7 - Sun 13 Sept: avg 101.67 (prior completed week)
        # Mon 14 - Sun 20 Sept: avg 105.39 (last completed week) -> stations raise
        # Mon 21 - Tue 22 Sept (partial): avg 99.94 -> next update should fall
        dates = (
            ["2026-09-07", "2026-09-08", "2026-09-09", "2026-09-10", "2026-09-11", "2026-09-12", "2026-09-13"]
            + ["2026-09-14", "2026-09-15", "2026-09-16", "2026-09-17", "2026-09-18", "2026-09-19", "2026-09-20"]
            + ["2026-09-21", "2026-09-22"]
        )
        prices = [101.67] * 7 + [105.39] * 7 + [100.34, 99.53]
        signal = weekly_email.weekly_brent_signal(dates, prices, today=date(2026, 9, 22))
        self.assertAlmostEqual(signal["this_week_move"], 3.72, places=2)
        self.assertAlmostEqual(signal["next_week_outlook"], -105.39 + (100.34 + 99.53) / 2, places=2)
        self.assertEqual(signal["last_completed_week_end"], "2026-09-20")

    def test_none_with_no_daily_data(self) -> None:
        self.assertIsNone(weekly_email.weekly_brent_signal([], []))

    def test_none_fields_with_only_one_completed_week(self) -> None:
        dates, prices = self._dates_prices("2026-09-14", [100.0] * 7)
        signal = weekly_email.weekly_brent_signal(dates, prices, today=date(2026, 9, 20))
        self.assertIsNone(signal["this_week_move"])   # no prior week to compare
        self.assertIsNone(signal["next_week_outlook"])  # that week is already complete, no partial week yet


class WeeklyContextLineTests(unittest.TestCase):
    def test_conflicting_moves_produce_a_note(self) -> None:
        line = weekly_email.weekly_context_line(3.72, -5.45)
        self.assertIn("raised", line)
        self.assertIn("fall", line)

    def test_agreeing_moves_produce_no_note(self) -> None:
        self.assertIsNone(weekly_email.weekly_context_line(3.0, 4.0))

    def test_tiny_move_produces_no_note(self) -> None:
        self.assertIsNone(weekly_email.weekly_context_line(0.2, -3.0))

    def test_missing_data_produces_no_note(self) -> None:
        self.assertIsNone(weekly_email.weekly_context_line(None, -3.0))


if __name__ == "__main__":
    unittest.main()
