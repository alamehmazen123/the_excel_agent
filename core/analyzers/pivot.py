"""Pivot analyzer: dimension x measure aggregations with a native chart."""
from __future__ import annotations

from typing import Optional

from ..aggregate import group_sum
from ..constants import SHEET_PIVOT
from ..models import WorkbookProfile
from ..render import (ChartKind, ChartSpec, DataTable, NumberFormat, SheetSpec,
                      TextBlock)
from .base import Analyzer


class PivotAnalyzer(Analyzer):
    key = "pivot"
    sheet_name = SHEET_PIVOT

    def applies_to(self, profile: WorkbookProfile) -> bool:
        t = profile.primary
        return bool(t and t.dimensions and t.measures)

    def run(self, profile: WorkbookProfile) -> Optional[SheetSpec]:
        table = profile.primary
        if table is None or not (table.dimensions and table.measures):
            return None

        spec = SheetSpec(
            name=SHEET_PIVOT, heading="Pivot Analysis",
            subheading=f"Source: {table.sheet_name}  •  aggregated totals by category",
        )

        measure = table.measures[0]
        # One pivot per dimension (cap at 3 to keep the sheet readable).
        for dim in table.dimensions[:3]:
            ranked = group_sum(table, dim, measure, top_n=15)
            if not ranked:
                continue
            spec.tables.append(DataTable(
                title=f"{measure.name} by {dim.name}",
                headers=[dim.name, f"Total {measure.name}", "Share"],
                rows=self._with_share(ranked),
                formats=[NumberFormat.GENERAL, NumberFormat.DECIMAL, NumberFormat.PERCENT],
            ))
            spec.charts.append(ChartSpec(
                kind=ChartKind.BAR, title=f"{measure.name} by {dim.name}",
                categories=[k for k, _ in ranked],
                series_name=f"Total {measure.name}",
                values=[round(v, 2) for _, v in ranked],
            ))

        if not spec.tables:
            return None
        spec.text_blocks.append(TextBlock(
            title="About these pivots",
            paragraphs=[
                f"Each table sums '{measure.name}' across the values of a "
                "category column, ranked from largest to smallest with each "
                "group's share of the total.",
            ],
        ))
        return spec

    @staticmethod
    def _with_share(ranked: list[tuple[str, float]]) -> list[list]:
        total = sum(v for _, v in ranked) or 1.0
        return [[k, round(v, 2), v / total] for k, v in ranked]
