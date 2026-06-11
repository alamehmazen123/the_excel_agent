"""KPI analyzer: headline metrics, per-measure stats, period growth, top-N."""
from __future__ import annotations

from typing import Optional

from ..aggregate import group_sum, period_over_period_growth, time_series
from ..constants import SHEET_KPI
from ..formatting import fmt_measure, fmt_number, fmt_percent
from ..models import WorkbookProfile
from ..render import DataTable, KpiTile, NumberFormat, SheetSpec, TextBlock
from .base import Analyzer


class KpiAnalyzer(Analyzer):
    key = "kpi"
    sheet_name = SHEET_KPI

    def __init__(self, include_tables: bool = True) -> None:
        # When Excel is available the COM finalizer builds Measure Statistics and
        # the date breakdown as real PivotTables, so we emit only the tiles here.
        self.include_tables = include_tables

    def applies_to(self, profile: WorkbookProfile) -> bool:
        t = profile.primary
        return bool(t and t.key_measures and t.row_count > 0)

    def run(self, profile: WorkbookProfile) -> Optional[SheetSpec]:
        table = profile.primary
        if table is None or not table.measures:
            return None

        spec = SheetSpec(
            name=SHEET_KPI, heading="KPI Analysis",
            subheading=f"Source: {table.sheet_name}  •  {table.row_count:,} records",
        )

        # Headline tiles: record count + total/avg of up to 3 key measures
        # (the user's Custom picks when provided, else auto).
        spec.kpi_tiles.append(KpiTile("Total Records", f"{table.row_count:,}"))
        date_cols = table.date_columns
        tile_measures = table.measures_for(profile.preferred_measure_names) or table.measures
        for measure in tile_measures[:3]:
            total = measure.total or 0.0
            tile = KpiTile(
                label=f"Total {measure.name}",
                value=fmt_measure(measure, total),
                caption=f"avg {fmt_measure(measure, measure.mean or 0.0)}",
            )
            # Period-over-period growth, if a date column exists.
            if date_cols:
                series = time_series(table, date_cols[0], measure)
                growth = period_over_period_growth(series)
                if growth is not None:
                    tile.caption = f"{'+' if growth >= 0 else ''}{fmt_percent(growth)} vs prior period"
                    tile.good = growth >= 0
            spec.kpi_tiles.append(tile)

        if not self.include_tables:
            # Pivots (Measure Statistics, date breakdown) are added by Excel below.
            return spec

        # --- static-table fallback (no Excel) ---
        stat_rows = []
        for m in table.measures:
            stat_rows.append([
                m.name, m.count,
                round(m.total or 0, 2), round(m.mean or 0, 2),
                round(m.minimum or 0, 2), round(m.maximum or 0, 2),
            ])
        spec.tables.append(DataTable(
            title="Measure Statistics",
            headers=["Measure", "Count", "Total", "Average", "Min", "Max"],
            rows=stat_rows,
            formats=[NumberFormat.GENERAL, NumberFormat.INTEGER, NumberFormat.DECIMAL,
                     NumberFormat.DECIMAL, NumberFormat.DECIMAL, NumberFormat.DECIMAL],
        ))
        if table.dimensions:
            dim = table.dimensions[0]
            measure = table.measures[0]
            ranked = group_sum(table, dim, measure, top_n=10)
            if ranked:
                spec.tables.append(DataTable(
                    title=f"Top {dim.name} by {measure.name}",
                    headers=[dim.name, f"Total {measure.name}"],
                    rows=[[k, round(v, 2)] for k, v in ranked],
                    formats=[NumberFormat.GENERAL, NumberFormat.DECIMAL],
                ))
        return spec
