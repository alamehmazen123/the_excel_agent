"""Human-friendly value formatting shared by analyzers.

Currency rule (hospital default): all monetary values are Lebanese Pounds (LBP)
UNLESS the column is explicitly in dollars (its header or source cell format says
so). This mirrors the COM PivotTable formats in ``pivot_plan`` so every sheet —
tiles, dashboards, summaries, pivots — is consistent.
"""
from __future__ import annotations

from .models import ColumnProfile, ColumnType


def is_dollar_column(col: ColumnProfile) -> bool:
    """True only when the column is explicitly in dollars (header or cell format).

    Everything else monetary is treated as LBP. Shared with pivot_plan so the
    static sheets and the active PivotTables label currency identically.
    """
    if col is None:
        return False
    fmt = col.number_format or ""
    if "$" in fmt or "USD" in fmt:
        return True
    n = (col.name or "").lower()
    return ("$" in n or "usd" in n or "usdt" in n or "dollar" in n)


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


def fmt_currency(v: float, dollar: bool = False) -> str:
    """Money string. LBP by default ('… LBP'); '$…' only when ``dollar`` is True."""
    if v is None:
        return "-"
    if dollar:
        return "$" + fmt_number(v)
    return fmt_number(v) + " LBP"


def fmt_percent(v: float) -> str:
    if v is None:
        return "-"
    # Values already in 0-1 are scaled; values like 12.3 are treated as percent.
    pct = v * 100 if abs(v) <= 1.5 else v
    return f"{pct:.1f}%"


def fmt_measure(col: ColumnProfile, v: float) -> str:
    if col.ctype == ColumnType.CURRENCY:
        return fmt_currency(v, dollar=is_dollar_column(col))
    if col.ctype == ColumnType.PERCENT:
        return fmt_percent(v)
    return fmt_number(v)
