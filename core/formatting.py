"""Human-friendly value formatting shared by analyzers."""
from __future__ import annotations

from .models import ColumnProfile, ColumnType


def fmt_number(v: float) -> str:
    if v is None:
        return "-"
    if abs(v) >= 1_000_000_000:
        return f"{v / 1_000_000_000:.2f}B"
    if abs(v) >= 1_000_000:
        return f"{v / 1_000_000:.2f}M"
    if abs(v) >= 1_000:
        return f"{v / 1_000:.1f}K"
    if float(v).is_integer():
        return f"{int(v):,}"
    return f"{v:,.2f}"


def fmt_currency(v: float) -> str:
    if v is None:
        return "-"
    return "$" + fmt_number(v)


def fmt_percent(v: float) -> str:
    if v is None:
        return "-"
    # Values already in 0-1 are scaled; values like 12.3 are treated as percent.
    pct = v * 100 if abs(v) <= 1.5 else v
    return f"{pct:.1f}%"


def fmt_measure(col: ColumnProfile, v: float) -> str:
    if col.ctype == ColumnType.CURRENCY:
        return fmt_currency(v)
    if col.ctype == ColumnType.PERCENT:
        return fmt_percent(v)
    return fmt_number(v)
