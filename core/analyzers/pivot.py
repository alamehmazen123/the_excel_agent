"""Pivot analyzer (static openpyxl fallback): decoded, date-grouped tables.

Used when Excel is unavailable (or could not build active PivotTables). It now
honors the BASIC RULE — every category breakdown is grouped by Year/Month — and
uses the decoded helper dimensions + Lebanese-Pound formatting so it matches the
active PivotTables and the Smart Tables sheet.
"""
from __future__ import annotations

from typing import Optional

from ..aggregate import group_period_dim, group_sum, time_series
from ..constants import SHEET_PIVOT
from ..formatting import is_dollar_column
from ..models import ColumnProfile, WorkbookProfile
from ..render import (ChartKind, ChartSpec, DataTable, NumberFormat, SheetSpec,
                      TextBlock)
from .base import Analyzer

_MONTHS = ["", "Jan", "Feb", "Mar", "Apr", "May", "Jun",
           "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


def _fmt_period(period: str) -> str:
    try:
        y, m = period.split("-")
        return f"{_MONTHS[int(m)]}-{y[2:]}"
    except Exception:
        return period


class PivotAnalyzer(Analyzer):
    key = "pivot"
    sheet_name = SHEET_PIVOT

    def applies_to(self, profile: WorkbookProfile) -> bool:
        t = profile.primary
        return bool(t and t.dimensions and t.measures)

    def _fmt_for(self, measure: ColumnProfile) -> NumberFormat:
        if measure.ctype.name == "PERCENT":
            return NumberFormat.PERCENT
        return (NumberFormat.CURRENCY if is_dollar_column(measure)
                else NumberFormat.LBP)

    def run(self, profile: WorkbookProfile) -> Optional[SheetSpec]:
        table = profile.primary
        if table is None or not (table.dimensions and table.measures):
            return None

        spec = SheetSpec(
            name=SHEET_PIVOT, heading="Pivot Analysis",
            subheading=f"Source: {table.sheet_name}  •  totals by month and category",
        )

        measure = table.measures[0]
        fmt = self._fmt_for(measure)
        date_col = table.date_columns[0] if table.date_columns else None

        # Time trend first (always carries the date).
        if date_col is not None:
            series = time_series(table, date_col, measure)
            if series:
                spec.tables.append(DataTable(
                    title=f"Total {measure.name} by {date_col.name} (Month)",
                    headers=[f"{date_col.name} (Month)", f"Total {measure.name}"],
                    rows=[[_fmt_period(p), round(v, 2)] for p, v in series],
                    formats=[NumberFormat.GENERAL, fmt],
                ))
                spec.charts.append(ChartSpec(
                    kind=ChartKind.LINE, title=f"{measure.name} over time",
                    categories=[_fmt_period(p) for p, _ in series],
                    series_name=f"Total {measure.name}",
                    values=[round(v, 2) for _, v in series]))

        # One breakdown per dimension, grouped by month (BASIC RULE), cap at 3.
        for dim in table.dimensions[:3]:
            if date_col is not None:
                rows = group_period_dim(table, date_col, dim, measure, top_n=20)
                if not rows:
                    continue
                spec.tables.append(DataTable(
                    title=f"{measure.name} by {date_col.name} (Month) & {dim.name}",
                    headers=[f"{date_col.name} (Month)", dim.name,
                             f"Total {measure.name}", "Records"],
                    rows=[[_fmt_period(p), label, round(t, 2), c]
                          for p, label, t, c in rows],
                    formats=[NumberFormat.GENERAL, NumberFormat.GENERAL, fmt,
                             NumberFormat.INTEGER],
                ))
            else:
                ranked = group_sum(table, dim, measure, top_n=15)
                if not ranked:
                    continue
                spec.tables.append(DataTable(
                    title=f"{measure.name} by {dim.name}",
                    headers=[dim.name, f"Total {measure.name}", "Share"],
                    rows=self._with_share(ranked),
                    formats=[NumberFormat.GENERAL, fmt, NumberFormat.PERCENT],
                ))
                spec.charts.append(ChartSpec(
                    kind=ChartKind.BAR, title=f"{measure.name} by {dim.name}",
                    categories=[k for k, _ in ranked],
                    series_name=f"Total {measure.name}",
                    values=[round(v, 2) for _, v in ranked]))

        if not spec.tables:
            return None
        spec.text_blocks.append(TextBlock(
            title="About these pivots",
            paragraphs=[
                f"Each table sums '{measure.name}' by month and category, decoded "
                "to real names where the reference library applies. Values are in "
                "Lebanese Pounds unless the column is explicitly in dollars.",
            ],
        ))
        return spec

    @staticmethod
    def _with_share(ranked: list[tuple[str, float]]) -> list[list]:
        total = sum(v for _, v in ranked) or 1.0
        return [[k, round(v, 2), v / total] for k, v in ranked]
