"""Dashboard analyzer — a real one-page dashboard of KPI tiles + charts.

Built with openpyxl so it is ALWAYS present (with or without Excel) and never
left blank: headline tiles, a monthly trend line, top decoded dimensions as
bars, a composition donut/pie, and a Pareto of the leading dimension. Names are
decoded via the library so the GM sees real departments/payers, not codes.
"""
from __future__ import annotations

from collections import defaultdict
from typing import Optional

from ..aggregate import group_sum, time_series
from ..constants import SHEET_DASHBOARD
from ..decode import friendly_name
from ..formatting import fmt_measure, fmt_percent
from ..models import ColumnProfile, TableProfile, WorkbookProfile
from ..render import ChartKind, ChartSpec, KpiTile, SheetSpec, TextBlock
from .base import Analyzer

_MONTHS = ["", "Jan", "Feb", "Mar", "Apr", "May", "Jun",
           "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


def _fmt_period(period: str) -> str:
    try:
        y, m = period.split("-")
        return f"{_MONTHS[int(m)]}-{y}"
    except Exception:
        return period


class DashboardAnalyzer(Analyzer):
    key = "dashboard"
    sheet_name = SHEET_DASHBOARD

    def applies_to(self, profile: WorkbookProfile) -> bool:
        t = profile.primary
        return bool(t and t.key_measures and t.row_count > 0)

    def _dimensions(self, table: TableProfile):
        helpers = [c for c in table.columns if c.is_decoded_helper]
        seen = {c.name for c in helpers}
        plain = [c for c in table.pivot_dimensions
                 if not c.is_decoded_helper and not c.decoded_helper
                 and c.name not in seen]
        return helpers + plain

    def run(self, profile: WorkbookProfile) -> Optional[SheetSpec]:
        table = profile.primary
        if table is None or not table.key_measures:
            return None

        spec = SheetSpec(
            name=SHEET_DASHBOARD, heading="Executive Dashboard",
            subheading=f"Source: {table.sheet_name}  •  {table.row_count:,} records")

        measure = (table.value_for(profile.preferred_value_name)
                   or table.measures_for(profile.preferred_measure_names)[0])
        mname = friendly_name(measure.name)
        date_col = table.date_columns[0] if table.date_columns else None
        series = time_series(table, date_col, measure) if date_col else []
        dims = self._dimensions(table)

        # -- KPI tiles --------------------------------------------------------
        spec.kpi_tiles.append(KpiTile("Records", f"{table.row_count:,}"))
        spec.kpi_tiles.append(KpiTile(f"Total {mname}",
                                      fmt_measure(measure, measure.total or 0.0)))
        if series:
            best = max(series, key=lambda kv: kv[1])
            spec.kpi_tiles.append(KpiTile("Best Month", _fmt_period(best[0]),
                                          caption=fmt_measure(measure, best[1]), good=True))
            if len(series) >= 2:
                prev, last = series[-2][1], series[-1][1]
                mom = (last - prev) / abs(prev) if prev else None
                if mom is not None:
                    spec.kpi_tiles.append(KpiTile(
                        "Month-over-Month",
                        f"{'+' if mom >= 0 else ''}{fmt_percent(mom)}",
                        caption=f"{_fmt_period(series[-1][0])} vs prior", good=mom >= 0))
        if dims:
            top = group_sum(table, dims[0], measure, top_n=1)
            if top:
                spec.kpi_tiles.append(KpiTile(f"Top {friendly_name(dims[0].name)}",
                                              str(top[0][0])))

        # -- Chart 1: monthly trend (line) -----------------------------------
        if len(series) >= 2:
            spec.charts.append(ChartSpec(
                kind=ChartKind.LINE, title=f"{mname} by month",
                categories=[_fmt_period(p) for p, _ in series],
                series_name=mname, values=[round(v, 2) for _, v in series]))

        # -- Chart 2: top decoded dimension (bar) ----------------------------
        if dims:
            ranked = [(k, v) for k, v in group_sum(table, dims[0], measure, top_n=8) if v != 0]
            if ranked:
                spec.charts.append(ChartSpec(
                    kind=ChartKind.BAR,
                    title=f"{mname} by {friendly_name(dims[0].name)}",
                    categories=[str(k) for k, _ in ranked],
                    series_name=mname, values=[round(v, 2) for _, v in ranked]))

        # -- Chart 3: composition of a second dimension (pie) ----------------
        if len(dims) >= 2:
            ranked2 = [(k, v) for k, v in group_sum(table, dims[1], measure, top_n=6) if v > 0]
            if ranked2:
                spec.charts.append(ChartSpec(
                    kind=ChartKind.PIE,
                    title=f"{mname} share by {friendly_name(dims[1].name)}",
                    categories=[str(k) for k, _ in ranked2],
                    series_name=mname, values=[round(v, 2) for _, v in ranked2]))

        # -- Chart 4: Pareto of the leading dimension ------------------------
        if dims:
            items = [(k, v) for k, v in group_sum(table, dims[0], measure, top_n=10) if v > 0]
            tot = sum(v for _, v in items)
            if tot > 0 and len(items) >= 3:
                cum, cum_pct = 0.0, []
                for _, v in items:
                    cum += v
                    cum_pct.append(round(cum / tot * 100, 1))
                spec.charts.append(ChartSpec(
                    kind=ChartKind.PARETO,
                    title=f"{mname} concentration by {friendly_name(dims[0].name)}",
                    categories=[str(k) for k, _ in items],
                    series_name=mname, values=[round(v, 2) for _, v in items],
                    line_values=cum_pct, line_name="Cumulative %"))

        # -- Chart 5: year-over-year (when 2+ years) -------------------------
        if series:
            yearly = defaultdict(float)
            for p, v in series:
                yearly[p.split("-")[0]] += v
            if len(yearly) >= 2:
                yrs = sorted(yearly)
                spec.charts.append(ChartSpec(
                    kind=ChartKind.BAR, title=f"{mname} by year",
                    categories=yrs, series_name=mname,
                    values=[round(yearly[y], 2) for y in yrs]))

        if profile.pivot_sheets:
            spec.text_blocks.append(TextBlock(
                title="Sheets skipped (already contain a PivotTable)",
                paragraphs=["Left untouched: " + ", ".join(profile.pivot_sheets) + "."]))
        return spec
