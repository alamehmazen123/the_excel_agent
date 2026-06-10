"""Grouping / aggregation helpers shared by the analyzers."""
from __future__ import annotations

import datetime as _dt
from collections import defaultdict
from typing import Any, Optional

from .models import ColumnProfile, TableProfile


def _num(v: Any) -> Optional[float]:
    if isinstance(v, bool) or v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    return None


def group_sum(table: TableProfile, dim: ColumnProfile, measure: ColumnProfile,
              top_n: int = 10) -> list[tuple[str, float]]:
    """Sum ``measure`` grouped by ``dim``; return top_n groups by descending sum."""
    totals: dict[str, float] = defaultdict(float)
    for row in table.rows:
        key = row.get(dim.name)
        val = _num(row.get(measure.name))
        if key is None or key == "" or val is None:
            continue
        totals[str(key)] += val
    ranked = sorted(totals.items(), key=lambda kv: kv[1], reverse=True)
    return ranked[:top_n]


def _period_key(value: Any) -> Optional[str]:
    """Bucket a date value into a 'YYYY-MM' period string."""
    if isinstance(value, _dt.datetime):
        return f"{value.year:04d}-{value.month:02d}"
    if isinstance(value, _dt.date):
        return f"{value.year:04d}-{value.month:02d}"
    return None


def time_series(table: TableProfile, date_col: ColumnProfile,
                measure: ColumnProfile) -> list[tuple[str, float]]:
    """Sum ``measure`` by month period, sorted chronologically."""
    totals: dict[str, float] = defaultdict(float)
    for row in table.rows:
        period = _period_key(row.get(date_col.name))
        val = _num(row.get(measure.name))
        if period is None or val is None:
            continue
        totals[period] += val
    return sorted(totals.items(), key=lambda kv: kv[0])


def period_over_period_growth(series: list[tuple[str, float]]) -> Optional[float]:
    """Return fractional growth between the last two periods, or None."""
    if len(series) < 2:
        return None
    prev, last = series[-2][1], series[-1][1]
    if prev == 0:
        return None
    return (last - prev) / abs(prev)
