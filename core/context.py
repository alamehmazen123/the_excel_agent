"""Known real-world events that explain otherwise-alarming numbers.

Some dips/anomalies in the data are NOT operational problems — they reflect
known events (a closure, a strike, a system migration). This module lets the
insight engine recognise those periods and EXPLAIN them ("low revenue here is
the 2026 closure"), instead of flagging them as red findings.

Currently this is a small, hard-coded list for Sahel General Hospital. It is
pure ``core/`` data — no UI, no Excel.
"""
from __future__ import annotations

import calendar
import datetime as _dt
from dataclasses import dataclass


@dataclass(frozen=True)
class KnownEvent:
    name: str
    start: _dt.date
    end: _dt.date
    short: str            # one-line label, e.g. "closure (war crisis)"
    note: str             # full explanation shown on the Insights sheet

    def covers_date(self, d: _dt.date) -> bool:
        return self.start <= d <= self.end

    def covers_period(self, period: str) -> bool:
        """True if a 'YYYY-MM' month overlaps this event at all."""
        try:
            y, m = int(period[:4]), int(period[5:7])
        except Exception:
            return False
        month_start = _dt.date(y, m, 1)
        month_end = _dt.date(y, m, calendar.monthrange(y, m)[1])
        return not (month_end < self.start or month_start > self.end)


# --- Sahel General Hospital known events ---------------------------------- #
KNOWN_EVENTS: list[KnownEvent] = [
    KnownEvent(
        name="2024 closure (Lebanon war crisis)",
        start=_dt.date(2024, 9, 28),
        end=_dt.date(2024, 11, 26),
        short="2024 closure (war crisis)",
        note=("Due to the 2024 war crisis in Lebanon, the hospital officially "
              "shut down from 28 September 2024 to 26 November 2024, keeping only "
              "Emergency, Lab, Radiology, minor OR surgeries and a few Dialysis "
              "patients running; regular operations resumed on 27 November 2024. "
              "Low revenue, admissions, load and expenses in this window reflect "
              "the closure — they are expected, not an operational problem."),
    ),
    KnownEvent(
        name="2026 closure (Lebanon war crisis)",
        start=_dt.date(2026, 3, 5),
        end=_dt.date(2026, 4, 16),
        short="2026 closure (war crisis)",
        note=("Due to the 2026 war crisis in Lebanon, the hospital officially "
              "shut down from 5 March 2026 to 16 April 2026, keeping only "
              "Emergency, Lab, Radiology, minor OR surgeries and a few Dialysis "
              "patients running; regular operations resumed on 17 April 2026. "
              "Low revenue, admissions, load and expenses in this window reflect "
              "the closure — they are expected, not an operational problem."),
    ),
]


def event_for_period(period: str) -> KnownEvent | None:
    """The known event overlapping a 'YYYY-MM' month, if any."""
    for ev in KNOWN_EVENTS:
        if ev.covers_period(period):
            return ev
    return None


def events_overlapping(periods: list[str]) -> list[KnownEvent]:
    """Distinct known events overlapping any of the given 'YYYY-MM' periods."""
    out: list[KnownEvent] = []
    for ev in KNOWN_EVENTS:
        if any(ev.covers_period(p) for p in periods) and ev not in out:
            out.append(ev)
    return out
