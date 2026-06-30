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


_CURRENCY_SYMBOLS = ("$", "€", "£", "¥", "₹", "₽", "[$")


def _format_is_currency(fmt: str) -> bool:
    f = fmt or ""
    return any(sym in f for sym in _CURRENCY_SYMBOLS)


def _format_is_percent(fmt: str) -> bool:
    return "%" in (fmt or "")


def _looks_like_percent(name: str) -> bool:
    n = name.lower()
    return "%" in n or "percent" in n or "pct" in n or n.endswith(" rate")


def _looks_like_currency(name: str) -> bool:
    # Header HINTS only -- the cell number format is the stronger signal.
    n = name.lower()
    keys = ("price", "cost", "revenue", "sales", "amount", "total", "$",
            "usd", "usdt", "eur", "gbp", "pnl", "income", "expense", "profit",
            "budget", "spend", "balance", "fee", "pay", "salary", "wage",
            "value", "deposit", "withdraw", "charge", "premium")
    return any(k in n for k in keys)


_ID_TOKENS = {"id", "no", "no.", "#", "index", "idx", "serial", "seq",
              "sequence", "row", "rownum", "s/n", "sn", "sr", "sr.", "key"}


def _looks_like_identifier(name: str, is_first: bool, all_integer: bool,
                           distinct: int, count: int, is_currency_fmt: bool,
                           value_keyword: bool) -> bool:
    """A numeric row id / code column that must never be treated as a value.

    Works on ANY workbook: an integer column that is (near-)unique and isn't
    money-formatted or value-named is almost certainly an identifier/code.
    """
    if not all_integer or count == 0:
        return False
    if is_currency_fmt or value_keyword:
        return False
    n = name.strip().lower()
    header_id = (
        n in _ID_TOKENS
        or n.endswith(" id") or n.endswith("_id") or n.startswith("id ")
        or "id" in n.split() or "number" in n.split() or n == "number"
        or n.endswith(" no") or n.endswith(" #") or n.endswith("code")
    )
    unique_ratio = distinct / count
    if header_id:
        return True
    # First column of unique integers -> a row number.
    if is_first and unique_ratio > 0.9:
        return True
    # Any (near-)unique integer column with no value meaning -> a code/id.
    if unique_ratio > 0.98 and distinct >= 10:
        return True
    return False


def profile_column(name: str, index: int, values: list[Any],
                   number_format: str = "General") -> ColumnProfile:
    """Infer the type and statistics for a single column.

    The cell ``number_format`` is used as the PRIMARY signal for currency /
    percent (data-driven), with header keywords as a fallback hint -- so the
    agent self-identifies value columns on workbooks with any header names.
    """
    non_null = [v for v in values if v is not None and v != ""]
    count = len(non_null)
    nulls = len(values) - count
    distinct = len(set(non_null))

    prof = ColumnProfile(
        name=str(name), index=index, ctype=ColumnType.EMPTY,
        count=count, nulls=nulls, distinct=distinct, number_format=number_format,
    )
    if count == 0:
        return prof

    n_numbers = sum(1 for v in non_null if _is_number(v))
    n_dates = sum(1 for v in non_null if _is_date(v))
    frac_numeric = n_numbers / count
    frac_date = n_dates / count

    fmt_currency = _format_is_currency(number_format)
    fmt_percent = _format_is_percent(number_format)

    if frac_date >= 0.8:
        prof.ctype = ColumnType.DATE
        prof.top_values = Counter(non_null).most_common(10)
    elif frac_numeric >= 0.8:
        nums = [float(v) for v in non_null if _is_number(v)]
        prof.minimum = min(nums)
        prof.maximum = max(nums)
        prof.total = sum(nums)
        prof.mean = prof.total / len(nums)
        all_integer = all(float(v).is_integer() for v in nums)
        value_kw = _looks_like_currency(name)
        if _looks_like_identifier(name, index == 0, all_integer, distinct,
                                  count, fmt_currency, value_kw):
            prof.ctype = ColumnType.IDENTIFIER
        elif fmt_percent or _looks_like_percent(name):
            prof.ctype = ColumnType.PERCENT
        elif fmt_currency or value_kw:
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


def profile_table(table: TableProfile, formats: dict[str, str] | None = None) -> None:
    """Populate ``table.columns`` in place from its ``rows``."""
    if not table.rows:
        return
    formats = formats or {}
    names = list(table.rows[0].keys())
    for idx, name in enumerate(names):
        col_values = [row.get(name) for row in table.rows]
        table.columns.append(
            profile_column(name, idx, col_values, formats.get(name, "General")))
