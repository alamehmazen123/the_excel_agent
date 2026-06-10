"""Dashboard analyzer: a curated single sheet of KPI tiles + charts."""
from __future__ import annotations

from typing import Optional

from ..aggregate import group_sum, time_series
from ..constants import SHEET_DASHBOARD
from ..formatting import fmt_measure
from ..models import WorkbookProfile
from ..render import ChartKind, ChartSpec, KpiTile, SheetSpec, TextBlock
from .base import Analyzer

ORG_NAME = "SAHEL GENERAL HOSPITAL"
PRODUCT_NAME = "Excel Intelligence Agent"

_MONTHS = ["", "Jan", "Feb", "Mar", "Apr", "May", "Jun",
           "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


def _fmt_period(period: str) -> str:
    """'2026-01' -> 'Jan-26' to match the PivotTables' month/year grouping."""
    try:
        y, m = period.split("-")
        return f"{_MONTHS[int(m)]}-{y[2:]}"
    except Exception:
        return period


class DashboardAnalyzer(Analyzer):
    key = "dashboard"
    sheet_name = SHEET_DASHBOARD

    def applies_to(self, profile: WorkbookProfile) -> bool:
        t = profile.primary
        return bool(t and t.key_measures and t.row_count > 0)

    def run(self, profile: WorkbookProfile) -> Optional[SheetSpec]:
        table = profile.primary
        if table is None or not table.key_measures:
            return None

        spec = SheetSpec(
            name=SHEET_DASHBOARD, heading=ORG_NAME,
            subheading=(f"{PRODUCT_NAME}   •   Source: {table.sheet_name}  •  "
                        f"{table.row_count:,} records"),
        )

        # Precisely report any sheets skipped because they already hold a pivot.
        if profile.pivot_sheets:
            names = ", ".join(profile.pivot_sheets)
            spec.text_blocks.append(TextBlock(
                title="Sheets skipped (already contain a PivotTable)",
                paragraphs=[
                    f"The following sheet(s) were detected as already containing a "
                    f"pivot table and were left untouched: {names}.",
                ],
            ))

        # Headline KPI tiles.
        spec.kpi_tiles.append(KpiTile("Records", f"{table.row_count:,}"))
        for m in table.key_measures[:3]:
            spec.kpi_tiles.append(KpiTile(
                f"Total {m.name}", fmt_measure(m, m.total or 0.0),
                caption=f"avg {fmt_measure(m, m.mean or 0.0)}",
            ))

        measure = table.primary_value_measure or table.key_measures[0]

        # Chart 1: top categories (bar) for the leading dimension.
        if table.dimensions:
            ranked = group_sum(table, table.dimensions[0], measure, top_n=8)
            if ranked:
                spec.charts.append(ChartSpec(
                    kind=ChartKind.BAR,
                    title=f"Top {table.dimensions[0].name} by {measure.name}",
                    categories=[k for k, _ in ranked],
                    series_name=measure.name,
                    values=[round(v, 2) for _, v in ranked],
                ))

        # Chart 2: trend over time (line) if a date column exists.
        if table.date_columns:
            series = time_series(table, table.date_columns[0], measure)
            if len(series) >= 2:
                spec.charts.append(ChartSpec(
                    kind=ChartKind.LINE,
                    title=f"{measure.name} over time",
                    categories=[_fmt_period(p) for p, _ in series],
                    series_name=measure.name,
                    values=[round(v, 2) for _, v in series],
                ))

        # Chart 3: composition (pie) of a second dimension, if present.
        if len(table.dimensions) >= 2:
            ranked2 = group_sum(table, table.dimensions[1], measure, top_n=6)
            if ranked2:
                spec.charts.append(ChartSpec(
                    kind=ChartKind.PIE,
                    title=f"{measure.name} share by {table.dimensions[1].name}",
                    categories=[k for k, _ in ranked2],
                    series_name=measure.name,
                    values=[round(v, 2) for _, v in ranked2],
                ))

        return spec
