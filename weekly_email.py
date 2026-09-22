"""Weekly Thursday "fill up before Monday?" email, sent through Buttondown.

The generator writes .cache/weekly_signal.json on every build. This script
turns that signal into the same recommendation the dashboard shows and, unless
--dry-run is given, sends it to all confirmed subscribers.

    python weekly_email.py --dry-run          # print the email, send nothing
    BUTTONDOWN_API_KEY=... python weekly_email.py

The forecast coefficients below are injected into the dashboard JavaScript by
generate_oil_dashboard.py, so the site and the email cannot drift apart.
"""

from __future__ import annotations

import argparse
import json
import math
import os
import sys
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path

import requests

SITE_URL = "https://fuelforecast.eu/"
SIGNAL_PATH = Path(".cache") / "weekly_signal.json"
BUTTONDOWN_EMAILS_URL = "https://api.buttondown.com/v1/emails"

# The Buttondown account is shared with another site, so every signup from this
# one carries a tag and a source. The Thursday email is sent to that tag only.
# Empty hides the signup form on the site. Set to "erireal" once Buttondown has
# finished verifying the account.
NEWSLETTER_BUTTONDOWN_USERNAME = ""
NEWSLETTER_SOURCE = "fuelforecast"

# Expected pump-price change in cents/L = coefficient * Brent move ($/bbl) * 10,
# using the first value when Brent rose and the second when it fell.
FORECAST_COEFFICIENTS = {
    "diesel": (0.09, 0.06),
    "euro95": (0.07, 0.05),
}
FUEL_LABELS = {"diesel": "Diesel", "euro95": "Euro-95"}
NO_CHANGE_CENTS = 2       # below this, "any day is fine"
TANK_LITRES = 50
MAX_BRENT_AGE_DAYS = 10   # older data are too stale to advise on


@dataclass(frozen=True)
class Advice:
    fuel: str
    expected_cents: float
    action: str        # "fill_up", "wait" or "any_day"
    cents: int
    tank_saving: float


def expected_change_cents(fuel: str, brent_move: float) -> float:
    rising, falling = FORECAST_COEFFICIENTS[fuel]
    return (rising if brent_move >= 0 else falling) * brent_move * 10


def advise(fuel: str, brent_move: float) -> Advice:
    expected = expected_change_cents(fuel, brent_move)
    # Match the dashboard's Math.round (halves round up), not Python's round().
    cents = abs(math.floor(expected + 0.5))
    saving = math.floor(abs(expected / 100 * TANK_LITRES) * 10 + 0.5) / 10
    if abs(expected) < NO_CHANGE_CENTS:
        action = "any_day"
    elif expected > 0:
        action = "fill_up"
    else:
        action = "wait"
    return Advice(fuel, expected, action, cents, saving)


def weekly_brent_signal(
    daily_dates: list[str], daily_prices: list[float], today: date | None = None
) -> dict | None:
    """Compare Monday-Sunday weekly average Brent prices.

    Stations set prices off the *previous completed week's average*, with a
    lag — not off any single day. Comparing single points (as this dashboard
    used to) is noisy: a mid-week spike can make it look like Brent fell,
    even in a week where its average — and so the price stations just set —
    rose. Returns:

      this_week_move:     last completed week's avg minus the week before —
                           what stations most likely just priced in.
      next_week_outlook:  the current, still-incomplete week's avg so far
                           minus the last completed week's avg — where the
                           *next* update is trending.

    Either field is None when there isn't enough data yet. Returns None
    entirely when there is no daily data at all.
    """
    if not daily_dates:
        return None
    import pandas as pd

    today = today or date.today()
    series = pd.Series(daily_prices, index=pd.to_datetime(daily_dates))
    series = series[~series.index.duplicated(keep="last")].sort_index()
    weekly = series.resample("W-SUN").mean().dropna()
    if weekly.empty:
        return None

    today_ts = pd.Timestamp(today)
    if weekly.index[-1] > today_ts:
        # The last bucket's Sunday hasn't happened yet: it's the current,
        # still-incomplete week (averaging whatever days have landed so far).
        current_avg, current_end = float(weekly.iloc[-1]), weekly.index[-1]
        completed = weekly.iloc[:-1]
    else:
        current_avg, current_end = None, None
        completed = weekly

    last_avg = float(completed.iloc[-1]) if len(completed) >= 1 else None
    last_end = completed.index[-1] if len(completed) >= 1 else None
    prior_avg = float(completed.iloc[-2]) if len(completed) >= 2 else None

    return {
        "this_week_move": (
            round(last_avg - prior_avg, 2)
            if last_avg is not None and prior_avg is not None else None
        ),
        "next_week_outlook": (
            round(current_avg - last_avg, 2)
            if current_avg is not None and last_avg is not None else None
        ),
        "last_completed_week_end": last_end.strftime("%Y-%m-%d") if last_end is not None else None,
        "current_week_end": current_end.strftime("%Y-%m-%d") if current_end is not None else None,
    }


def next_update_date(current_week_end: str | None, latest_pump_date: str) -> str:
    """The Monday stations next reset prices on: the day after the current,
    still-incomplete week ends, or latest_pump_date + 7 days if that isn't
    known yet. Mirrors the site's nextUpdateDateLabel()."""
    if current_week_end:
        base = datetime.strptime(current_week_end, "%Y-%m-%d").date()
        return (base + timedelta(days=1)).strftime("%Y-%m-%d")
    return (datetime.strptime(latest_pump_date, "%Y-%m-%d").date() + timedelta(days=7)).strftime("%Y-%m-%d")


def format_date_short(iso: str) -> str:
    """'2026-09-28' -> 'Mon 28 Sep', avoiding strftime's non-portable '%-d'."""
    d = datetime.strptime(iso, "%Y-%m-%d").date()
    return f"{d.strftime('%a')} {d.day} {d.strftime('%b')}"


def signal_from_data(data: dict) -> dict:
    """Latest Brent observation and the weekly-average move, as the dashboard sees them."""
    latest = data.get("brent_latest") or {}
    price, brent_date = latest.get("price"), latest.get("date")
    if price is None:
        price = next((v for v in reversed(data["brent"]) if v is not None), None)
        brent_date = data["latest_date"]
    previous = latest.get("previous_price")
    if previous is None:
        previous = next((v for v in reversed(data["brent"][:-1]) if v is not None), None)
    weekly = weekly_brent_signal(data.get("brent_daily_dates") or [], data.get("brent_daily") or [])
    return {
        "brent_price": price,
        "brent_previous": previous,
        "brent_date": brent_date,
        "this_week_move": (weekly or {}).get("this_week_move"),
        "next_week_outlook": (weekly or {}).get("next_week_outlook"),
        "next_update_date": next_update_date((weekly or {}).get("current_week_end"), data["latest_date"]),
    }


def load_signal(path: Path = SIGNAL_PATH) -> dict:
    return json.loads(path.read_text(encoding="utf-8"))


def signal_problem(signal: dict, today: date | None = None) -> str | None:
    """Return why we should not send, or None when the data are usable."""
    today = today or date.today()
    if signal.get("brent_price") is None or signal.get("brent_previous") is None:
        return "Brent price or its previous observation is missing"
    brent_date = datetime.strptime(signal["brent_date"], "%Y-%m-%d").date()
    age = (today - brent_date).days
    if age > MAX_BRENT_AGE_DAYS:
        return f"Brent data are {age} days old (limit {MAX_BRENT_AGE_DAYS})"
    return None


def _fuel_section(advice: Advice) -> str:
    label = FUEL_LABELS[advice.fuel]
    if advice.action == "fill_up":
        return (
            f"## {label}: fill up before Monday\n\n"
            f"Pump prices are expected to rise about **{advice.cents} cents/L** "
            f"at the next update. Filling up before then could save about "
            f"**€{advice.tank_saving:.2f}** on a {TANK_LITRES}L tank."
        )
    if advice.action == "wait":
        return (
            f"## {label}: wait until next week\n\n"
            f"Pump prices are expected to fall about **{advice.cents} cents/L** "
            f"at the next update. Waiting could save about "
            f"**€{advice.tank_saving:.2f}** on a {TANK_LITRES}L tank."
        )
    return (
        f"## {label}: any day is fine\n\n"
        "No meaningful price change is expected at the next update."
    )


def weekly_context_line(
    this_week_move: float | None, next_week_outlook: float | None, next_update: str | None = None
) -> str | None:
    """Explain a move that already happened this week, but only when it
    conflicts with where next week's update is heading — e.g. stations just
    raised prices off last week's average, even though this week's Brent has
    already reversed. When both agree, the headline advice already covers it."""
    if this_week_move is None or next_week_outlook is None:
        return None
    if abs(this_week_move) < 1 or this_week_move * next_week_outlook >= 0:
        return None
    this_direction = "raised" if this_week_move > 0 else "lowered"
    next_direction = "another rise" if next_week_outlook > 0 else "a fall"
    when = f", on {format_date_short(next_update)}" if next_update else ""
    return (
        f"Stations likely already {this_direction} prices this week, based on last "
        f"week's Brent average. But this week's Brent points to {next_direction} "
        f"at the next update{when}."
    )


def build_email(signal: dict) -> tuple[str, str]:
    # The raw week-on-week price move is just for the factual Brent line below.
    # The recommendation itself is driven by the smoothed weekly-average
    # outlook when available (see weekly_brent_signal), falling back to the
    # raw move only when there isn't enough daily history for it yet.
    raw_move = signal["brent_price"] - signal["brent_previous"]
    move = signal.get("next_week_outlook")
    if move is None:
        move = raw_move
    advices = [advise(fuel, move) for fuel in ("diesel", "euro95")]

    if any(a.action == "fill_up" for a in advices):
        subject = "Fuel tip: fill up before Monday"
    elif any(a.action == "wait" for a in advices):
        subject = "Fuel tip: wait until next week"
    else:
        subject = "Fuel tip: any day is fine next week"

    context_line = weekly_context_line(
        signal.get("this_week_move"), signal.get("next_week_outlook"), signal.get("next_update_date")
    )
    direction = "up" if raw_move >= 0 else "down"
    body = "\n\n".join([
        "Here is your Thursday fuel forecast.",
        *([context_line] if context_line else []),
        *(_fuel_section(a) for a in advices),
        f"Brent crude is at ${signal['brent_price']:.2f}/bbl "
        f"({direction} ${abs(raw_move):.2f} on last week, as of {signal['brent_date']}).",
        f"[See the full dashboard]({SITE_URL})",
        "*Indicative forecast based on Brent crude moves. Individual stations vary.*",
    ])
    return subject, body


def send_email(subject: str, body: str, api_key: str, draft: bool = False) -> dict:
    response = requests.post(
        BUTTONDOWN_EMAILS_URL,
        headers={"Authorization": f"Token {api_key}"},
        json={
            "subject": subject,
            "body": body,
            "status": "draft" if draft else "about_to_send",
            "filters": {
                "predicate": "and",
                "groups": [],
                "filters": [{
                    "operator": "contains",
                    "field": "subscriber.tags",
                    "value": NEWSLETTER_SOURCE,
                }],
            },
        },
        timeout=30,
    )
    if not response.ok:
        raise SystemExit(f"Buttondown rejected the email ({response.status_code}): {response.text}")
    return response.json()


def main() -> None:
    parser = argparse.ArgumentParser(description="Send the weekly Thursday fuel email")
    parser.add_argument("--signal", type=Path, default=SIGNAL_PATH,
                        help=f"Signal file written by the generator (default: {SIGNAL_PATH})")
    parser.add_argument("--dry-run", action="store_true",
                        help="Print the email instead of sending it")
    parser.add_argument("--draft", action="store_true",
                        help="Create the email in Buttondown as a draft without sending it")
    args = parser.parse_args()

    if not args.signal.exists():
        raise SystemExit(f"{args.signal} not found. Run generate_oil_dashboard.py first.")
    signal = load_signal(args.signal)
    problem = signal_problem(signal)
    if problem:
        raise SystemExit(f"Not sending: {problem}")

    subject, body = build_email(signal)
    if args.dry_run:
        print(f"Subject: {subject}\n\n{body}")
        return

    api_key = os.environ.get("BUTTONDOWN_API_KEY", "").strip()
    if not api_key:
        raise SystemExit("BUTTONDOWN_API_KEY is not set")
    created = send_email(subject, body, api_key, draft=args.draft)
    print(f"{'Drafted' if args.draft else 'Sent'}: {subject} (id {created.get('id')})")


if __name__ == "__main__":
    sys.exit(main())
