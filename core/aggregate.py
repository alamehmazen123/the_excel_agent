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


def group_period_dim(table: TableProfile, date_col: ColumnProfile,
                     dim: ColumnProfile, measure: ColumnProfile,
                     top_n: int = 40) -> list[tuple[str, str, float, int]]:
    """Sum ``measure`` and count rows by (Year-Month period, dim value).

    Returns the top ``top_n`` (period, label, total, count) tuples ranked by
    total descending, then re-sorted chronologically for readable display. This
    backs the date-grouped Smart Tables (BASIC RULE: every table carries a date).
    """
    totals: dict[tuple[str, str], float] = defaultdict(float)
    counts: dict[tuple[str, str], int] = defaultdict(int)
    for row in table.rows:
        period = _period_key(row.get(date_col.name))
        key = row.get(dim.name)
        val = _num(row.get(measure.name))
        if period is None or key is None or key == "" or val is None:
            continue
        k = (period, str(key))
        totals[k] += val
        counts[k] += 1
    top = sorted(totals.items(), key=lambda kv: abs(kv[1]), reverse=True)[:top_n]
    top.sort(key=lambda kv: (kv[0][0], -kv[1]))   # chronological, value desc
    return [(p, label, tot, counts[(p, label)]) for (p, label), tot in top]


def crosstab_period(table: TableProfile, date_col: ColumnProfile,
                    dim: ColumnProfile, measure: ColumnProfile,
                    top_n: int = 12, max_periods: int = 18) -> tuple:
    """Cross-tabulate ``measure`` as ``dim`` (rows) × month period (columns).

    Returns ``(periods, rows)`` where ``periods`` is the chronological list of
    'YYYY-MM' columns (capped to the most recent ``max_periods``) and ``rows`` is
    ``[(label, {period: total}, row_total), …]`` for the Top-N labels by total,
    with the remaining labels rolled into an 'Others' row. This is the
    months-across layout a manager reads left-to-right to compare months."""
    cells: dict[tuple[str, str], float] = defaultdict(float)
    label_totals: dict[str, float] = defaultdict(float)
    period_set: set[str] = set()
    for row in table.rows:
        period = _period_key(row.get(date_col.name))
        key = row.get(dim.name)
        val = _num(row.get(measure.name))
        if period is None or key is None or key == "" or val is None:
            continue
        k = str(key)
        cells[(k, period)] += val
        label_totals[k] += val
        period_set.add(period)

    periods = sorted(period_set)[-max_periods:]
    ranked = sorted(label_totals.items(), key=lambda kv: kv[1], reverse=True)
    top = ranked[:top_n]
    others = ranked[top_n:]

    rows: list[tuple] = []
    for label, tot in top:
        rows.append((label, {p: cells.get((label, p), 0.0) for p in periods}, tot))
    if others:
        other_cells = {p: sum(cells.get((lbl, p), 0.0) for lbl, _ in others)
                       for p in periods}
        rows.append(("Others", other_cells, sum(v for _, v in others)))
    return periods, rows


def period_over_period_growth(series: list[tuple[str, float]]) -> Optional[float]:
    """Return fractional growth between the last two periods, or None."""
    if len(series) < 2:
        return None
    prev, last = series[-2][1], series[-1][1]
    if prev == 0:
        return None
    return (last - prev) / abs(prev)
