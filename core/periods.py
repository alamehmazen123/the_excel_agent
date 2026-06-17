"""Period discovery engine — auto-detect all natural date periods in the data.

When a workbook is loaded, this module scans the date column and returns a
structured view of every meaningful period boundary: months, quarters, years,
half-years, and gaps. The UI uses this to offer smart period comparisons
without the user ever typing a date.

Design is pure ``core/`` — no UI, no Excel.
"""
from __future__ import annotations

import datetime as _dt
import calendar
from dataclasses import dataclass, field
from typing import Optional

from .context import event_for_period, events_overlapping
from .models import ColumnProfile, TableProfile, WorkbookProfile


@dataclass
class PeriodInfo:
    """A single discovered period (month, quarter, or year)."""
    key: str               # e.g. "2026-03", "2026-Q1", "2026"
    label: str             # e.g. "Mar 2026", "Q1 2026", "2026"
    start_date: _dt.date
    end_date: _dt.date
    is_complete: bool = True      # True if all expected data is present
    record_count: int = 0
    total_value: float = 0.0


@dataclass
class PeriodDiscovery:
    """Result of scanning a workbook's date column for all meaningful periods."""
    date_column: str
    date_range_min: Optional[_dt.date] = None
    date_range_max: Optional[_dt.date] = None

    # All discovered periods at each granularity.
    months: list[PeriodInfo] = field(default_factory=list)
    quarters: list[PeriodInfo] = field(default_factory=list)
    halfyears: list[PeriodInfo] = field(default_factory=list)
    years: list[PeriodInfo] = field(default_factory=list)

    # Gaps in the timeline (with known-event awareness).
    gaps: list[dict] = field(default_factory=list)

    # Suggested "focus periods" for comparison.
    suggested_focus: list[dict] = field(default_factory=list)
    suggested_compare: list[dict] = field(default_factory=list)

    @property
    def has_multiple_periods(self) -> bool:
        return len(self.months) >= 2

    @property
    def has_multiple_quarters(self) -> bool:
        return len(self.quarters) >= 2

    @property
    def has_multiple_years(self) -> bool:
        return len(self.years) >= 2

    @property
    def latest_month(self) -> Optional[str]:
        return self.months[-1].key if self.months else None

    @property
    def latest_quarter(self) -> Optional[str]:
        return self.quarters[-1].key if self.quarters else None

    @property
    def latest_year(self) -> Optional[str]:
        return self.years[-1].key if self.years else None


_MONTH_NAMES = ["", "Jan", "Feb", "Mar", "Apr", "May", "Jun",
                "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


def _month_label(key: str) -> str:
    """'2026-03' -> 'Mar 2026'."""
    try:
        y, m = key.split("-")
        return f"{_MONTH_NAMES[int(m)]} {y}"
    except Exception:
        return key


def _quarter_label(key: str) -> str:
    """'2026-Q1' -> 'Q1 2026'."""
    try:
        parts = key.split("-")
        return f"{parts[1]} {parts[0]}"
    except Exception:
        return key


def discover_periods(profile: WorkbookProfile,
                     primary_measure: Optional[ColumnProfile] = None) -> PeriodDiscovery:
    """Scan the date column of the primary table and return all periods.

    This is called once after profile loading. The UI then uses the result to
    populate period comparison dropdowns.
    """
    table = profile.primary
    if table is None or not table.date_columns:
        return PeriodDiscovery(date_column="")

    date_col = table.date_columns[0]
    discovery = PeriodDiscovery(date_column=date_col.name)

    # Collect all unique date values.
    dates: list[_dt.date] = []
    month_periods: dict[str, list[_dt.date]] = {}
    for row in table.rows:
        v = row.get(date_col.name)
        if isinstance(v, _dt.datetime):
            d = v.date()
        elif isinstance(v, _dt.date):
            d = v
        else:
            continue
        dates.append(d)
        mk = f"{d.year:04d}-{d.month:02d}"
        if mk not in month_periods:
            month_periods[mk] = []
        month_periods[mk].append(d)

    if not dates:
        return discovery

    discovery.date_range_min = min(dates)
    discovery.date_range_max = max(dates)

    # Build month periods.
    sorted_months = sorted(month_periods.keys())
    for mk in sorted_months:
        m_dates = month_periods[mk]
        y, m = int(mk[:4]), int(mk[5:7])
        start_d = _dt.date(y, m, 1)
        end_d = _dt.date(y, m, calendar.monthrange(y, m)[1])
        total_val = 0.0
        if primary_measure is not None:
            for row in table.rows:
                rk = _month_key(row.get(date_col.name))
                if rk == mk:
                    v = row.get(primary_measure.name)
                    if isinstance(v, (int, float)):
                        total_val += float(v)
        discovery.months.append(PeriodInfo(
            key=mk, label=_month_label(mk),
            start_date=start_d, end_date=end_d,
            record_count=len(m_dates),
            total_value=total_val,
        ))

    # Build quarter periods.
    q_data: dict[str, list[_dt.date]] = {}
    for d in dates:
        q = (d.month - 1) // 3 + 1
        qk = f"{d.year:04d}-Q{q}"
        if qk not in q_data:
            q_data[qk] = []
        q_data[qk].append(d)
    for qk in sorted(q_data.keys()):
        y = int(qk[:4])
        q = int(qk[6])
        start_m = (q - 1) * 3 + 1
        end_m = q * 3
        start_d = _dt.date(y, start_m, 1)
        end_d = _dt.date(y, end_m, calendar.monthrange(y, end_m)[1])
        discovery.quarters.append(PeriodInfo(
            key=qk, label=_quarter_label(qk),
            start_date=start_d, end_date=end_d,
            record_count=len(q_data[qk]),
        ))

    # Build year periods.
    y_data: dict[str, list[_dt.date]] = {}
    for d in dates:
        yk = f"{d.year:04d}"
        if yk not in y_data:
            y_data[yk] = []
        y_data[yk].append(d)
    for yk in sorted(y_data.keys()):
        y = int(yk)
        discovery.years.append(PeriodInfo(
            key=yk, label=yk,
            start_date=_dt.date(y, 1, 1),
            end_date=_dt.date(y, 12, 31),
            record_count=len(y_data[yk]),
        ))

    # Build half-year periods.
    h_data: dict[str, list[_dt.date]] = {}
    for d in dates:
        hk = f"{d.year:04d}-{'H1' if d.month <= 6 else 'H2'}"
        if hk not in h_data:
            h_data[hk] = []
        h_data[hk].append(d)
    for hk in sorted(h_data.keys()):
        y = int(hk[:4])
        is_h1 = hk.endswith("H1")
        start_d = _dt.date(y, 1, 1) if is_h1 else _dt.date(y, 7, 1)
        end_d = _dt.date(y, 6, 30) if is_h1 else _dt.date(y, 12, 31)
        discovery.halfyears.append(PeriodInfo(
            key=hk, label=hk,
            start_date=start_d, end_date=end_d,
            record_count=len(h_data[hk]),
        ))

    # Detect gaps.
    for i in range(len(sorted_months) - 1):
        cur = sorted_months[i]
        nxt = sorted_months[i + 1]
        cur_y, cur_m = int(cur[:4]), int(cur[5:7])
        nxt_y, nxt_m = int(nxt[:4]), int(nxt[5:7])
        months_between = (nxt_y - cur_y) * 12 + (nxt_m - cur_m)
        if months_between > 1:
            gap_label = f"between {_month_label(cur)} and {_month_label(nxt)}"
            ev = event_for_period(cur) or event_for_period(nxt)
            discovery.gaps.append({
                "from": cur, "to": nxt,
                "months_missing": months_between - 1,
                "label": gap_label,
                "explained": ev is not None,
                "event": ev.short if ev else None,
            })

    # Build suggested focus periods.
    if discovery.months:
        # Last 12 months.
        discovery.suggested_focus.append({
            "label": "Last 12 months",
            "type": "rolling",
            "periods": sorted_months[-12:] if len(sorted_months) >= 12 else sorted_months,
        })
        # Latest complete quarter.
        if discovery.quarters:
            discovery.suggested_focus.append({
                "label": f"Latest quarter ({discovery.quarters[-1].label})",
                "type": "quarter",
                "periods": [discovery.quarters[-1].key],
            })
        # Latest complete year.
        if discovery.years:
            discovery.suggested_focus.append({
                "label": f"Year {discovery.years[-1].key}",
                "type": "year",
                "periods": [discovery.years[-1].key],
            })
        # Year to date.
        if discovery.date_range_max:
            ytd_start = _dt.date(discovery.date_range_max.year, 1, 1)
            ytd_months = [mk for mk in sorted_months
                          if _dt.date(int(mk[:4]), int(mk[5:7]), 1) >= ytd_start]
            if ytd_months:
                discovery.suggested_focus.append({
                    "label": f"Year to date ({discovery.date_range_max.year})",
                    "type": "ytd",
                    "periods": ytd_months,
                })

    # Build suggested compare periods.
    if discovery.has_multiple_periods:
        discovery.suggested_compare.append({
            "label": "Previous period",
            "type": "prev_period",
        })
    if discovery.has_multiple_quarters:
        discovery.suggested_compare.append({
            "label": "Previous quarter",
            "type": "prev_quarter",
        })
    if discovery.has_multiple_years:
        discovery.suggested_compare.append({
            "label": "Same period last year",
            "type": "prev_year",
        })

    return discovery


def _month_key(value) -> Optional[str]:
    if isinstance(value, _dt.datetime):
        return f"{value.year:04d}-{value.month:02d}"
    if isinstance(value, _dt.date):
        return f"{value.year:04d}-{value.month:02d}"
    return None
