"""Column type inference and statistics over detected tables.

Kept separate from loading so the same logic can profile data that arrives
from any source (xlsx now; a DataFrame or API payload later).
"""
from __future__ import annotations

import datetime as _dt
from collections import Counter
from typing import Any

from .models import ColumnProfile, ColumnType, TableProfile

# A column is treated as categorical (a dimension) rather than free text when
# its distinct-value ratio is at or below this threshold.
_CATEGORICAL_RATIO = 0.5
_CATEGORICAL_MAX_DISTINCT = 50


def _is_number(v: Any) -> bool:
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def _is_date(v: Any) -> bool:
    return isinstance(v, (_dt.date, _dt.datetime))


def _looks_like_percent(name: str, values: list[Any]) -> bool:
    if "%" in name or "percent" in name.lower() or "rate" in name.lower():
        nums = [v for v in values if _is_number(v)]
        if nums and all(-1.5 <= v <= 1.5 for v in nums):
            return True
        if "%" in name or "percent" in name.lower():
            return True
    return False


def _looks_like_currency(name: str) -> bool:
    n = name.lower()
    keys = ("price", "cost", "revenue", "sales", "amount", "total", "$",
            "usd", "eur", "gbp", "income", "expense", "profit", "budget", "spend")
    return any(k in n for k in keys)


def profile_column(name: str, index: int, values: list[Any]) -> ColumnProfile:
    """Infer the type and compute statistics for a single column."""
    non_null = [v for v in values if v is not None and v != ""]
    count = len(non_null)
    nulls = len(values) - count
    distinct = len(set(non_null))

    prof = ColumnProfile(
        name=str(name), index=index, ctype=ColumnType.EMPTY,
        count=count, nulls=nulls, distinct=distinct,
    )
    if count == 0:
        return prof

    n_numbers = sum(1 for v in non_null if _is_number(v))
    n_dates = sum(1 for v in non_null if _is_date(v))
    frac_numeric = n_numbers / count
    frac_date = n_dates / count

    if frac_date >= 0.8:
        prof.ctype = ColumnType.DATE
        prof.top_values = Counter(non_null).most_common(10)
    elif frac_numeric >= 0.8:
        nums = [float(v) for v in non_null if _is_number(v)]
        prof.minimum = min(nums)
        prof.maximum = max(nums)
        prof.total = sum(nums)
        prof.mean = prof.total / len(nums)
        if _looks_like_percent(name, non_null):
            prof.ctype = ColumnType.PERCENT
        elif _looks_like_currency(name):
            prof.ctype = ColumnType.CURRENCY
        else:
            prof.ctype = ColumnType.NUMERIC
    else:
        ratio = distinct / count
        if ratio <= _CATEGORICAL_RATIO and distinct <= _CATEGORICAL_MAX_DISTINCT:
            prof.ctype = ColumnType.CATEGORICAL
        else:
            prof.ctype = ColumnType.TEXT
        prof.top_values = Counter(str(v) for v in non_null).most_common(10)

    return prof


def profile_table(table: TableProfile) -> None:
    """Populate ``table.columns`` in place from its ``rows``."""
    if not table.rows:
        return
    names = list(table.rows[0].keys())
    for idx, name in enumerate(names):
        col_values = [row.get(name) for row in table.rows]
        table.columns.append(profile_column(name, idx, col_values))
