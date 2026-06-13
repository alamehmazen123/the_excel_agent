"""KPI analyzer — a broad executive scorecard, not just a few tiles.

For a GM who wants the whole picture at a glance: headline tiles (total, average
per month, best/worst month, month-over-month and year-over-year growth, active
months, leaders), a month-by-month trend table with Δ%/share/cumulative and a
trend chart, top breakdowns by the decoded dimensions, and a currency split.
Everything is computed offline from the rows; the COM finalizer still adds a
real "Measure Statistics" PivotTable below.
"""
from __future__ import annotations

from collections import defaultdict
from typing import Optional

from ..aggregate import group_sum, time_series
from ..constants import SHEET_KPI
from ..decode import friendly_name
from ..formatting import fmt_measure, fmt_number, fmt_percent, is_dollar_column
from ..models import ColumnProfile, TableProfile, WorkbookProfile
from ..render import (ChartKind, ChartSpec, DataTable, KpiTile, NumberFormat,
                      SheetSpec, TextBlock)
from .base import Analyzer

_MONTHS = ["", "Jan", "Feb", "Mar", "Apr", "May", "Jun",
           "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


def _month_label(period: str) -> str:
    try:
        y, m = period.split("-")
        return f"{_MONTHS[int(m)]}-{y}"
    except Exception:
        return period


class KpiAnalyzer(Analyzer):
    key = "kpi"
    sheet_name = SHEET_KPI

    def __init__(self, include_tables: bool = True) -> None:
        # include_tables=False when Excel is present (COM adds the stats pivot);
        # the broad scorecard + trend below are ALWAYS emitted either way.
        self.include_tables = include_tables

    def applies_to(self, profile: WorkbookProfile) -> bool:
        t = profile.primary
        return bool(t and t.key_measures and t.row_count > 0)

    # -- helpers ------------------------------------------------------------- #
    def _primary(self, table: TableProfile, profile: WorkbookProfile) -> Optional[ColumnProfile]:
        chosen = table.measures_for(profile.preferred_measure_names)
        vals = [c for c in chosen if c.is_value] or table.value_measures
        return vals[0] if vals else (table.measures[0] if table.measures else None)

    def _dimensions(self, table: TableProfile):
        helpers = [c for c in table.columns if c.is_decoded_helper]
        seen = {c.name for c in helpers}
        plain = [c for c in table.pivot_dimensions
                 if not c.is_decoded_helper and not c.decoded_helper
                 and c.name not in seen]
        return helpers + plain

    def _scorecard(self, table, measure, series) -> list[KpiTile]:
        tiles = [KpiTile("Total Records", f"{table.row_count:,}")]
        if measure is None:
            return tiles
        total = measure.total or 0.0
        tiles.append(KpiTile(f"Total {friendly_name(measure.name)}",
                             fmt_measure(measure, total)))
        if series:
            vals = [v for _, v in series]
            avg = sum(vals) / len(vals)
            best = max(series, key=lambda kv: kv[1])
            worst = min(series, key=lambda kv: kv[1])
            tiles.append(KpiTile("Active Months", f"{len(series)}",
                                 caption=f"avg {fmt_measure(measure, avg)}/mo"))
            tiles.append(KpiTile("Best Month", _month_label(best[0]),
                                 caption=fmt_measure(measure, best[1]), good=True))
            tiles.append(KpiTile("Lowest Month", _month_label(worst[0]),
                                 caption=fmt_measure(measure, worst[1]), good=False))
            if len(series) >= 2:
                prev, last = vals[-2], vals[-1]
                mom = (last - prev) / abs(prev) if prev else None
                if mom is not None:
                    tiles.append(KpiTile("Month-over-Month",
                                         f"{'+' if mom >= 0 else ''}{fmt_percent(mom)}",
                                         caption=f"{_month_label(series[-1][0])} vs prior",
                                         good=mom >= 0))
            # Year-over-year (when two+ years are present).
            yearly = defaultdict(float)
            for p, v in series:
                yearly[p.split("-")[0]] += v
            yrs = sorted(yearly)
            if len(yrs) >= 2:
                py, cy = yearly[yrs[-2]], yearly[yrs[-1]]
                yoy = (cy - py) / abs(py) if py else None
                if yoy is not None:
                    tiles.append(KpiTile("Year-over-Year",
                                         f"{'+' if yoy >= 0 else ''}{fmt_percent(yoy)}",
                                         caption=f"{yrs[-1]} vs {yrs[-2]}", good=yoy >= 0))
        return tiles

    def _monthly_trend(self, measure, series) -> Optional[DataTable]:
        if not series:
            return None
        total = sum(v for _, v in series) or 1.0
        rows, cum = [], 0.0
        prev = None
        for p, v in series:
            cum += v
            delta = ((v - prev) / abs(prev) * 100) if prev else None
            rows.append([_month_label(p), round(v, 2),
                         round(delta, 1) if delta is not None else "—",
                         round(v / total * 100, 1), round(cum, 2)])
            prev = v
        return DataTable(
            title=f"Monthly trend — {friendly_name(measure.name)}",
            headers=["Month", "Total", "Δ% vs prior", "% of total", "Cumulative"],
            rows=rows,
            formats=[NumberFormat.GENERAL, self._fmt(measure), NumberFormat.GENERAL,
                     NumberFormat.GENERAL, self._fmt(measure)],
            bar_columns=[1])

    def _fmt(self, measure) -> NumberFormat:
        return NumberFormat.CURRENCY if is_dollar_column(measure) else NumberFormat.LBP

    # -- run ----------------------------------------------------------------- #
    def run(self, profile: WorkbookProfile) -> Optional[SheetSpec]:
        table = profile.primary
        if table is None or not table.measures:
            return None

        spec = SheetSpec(
            name=SHEET_KPI, heading="KPI Analysis",
            subheading=f"Source: {table.sheet_name}  •  {table.row_count:,} records")

        measure = self._primary(table, profile)
        date_col = table.date_columns[0] if table.date_columns else None
        series = time_series(table, date_col, measure) if (date_col and measure) else []

        spec.kpi_tiles = self._scorecard(table, measure, series)

        # Month-by-month trend table + a trend line chart.
        trend = self._monthly_trend(measure, series) if measure else None
        if trend is not None:
            spec.tables.append(trend)
            spec.charts.append(ChartSpec(
                kind=ChartKind.LINE, title=f"{friendly_name(measure.name)} by month",
                categories=[_month_label(p) for p, _ in series],
                series_name=friendly_name(measure.name),
                values=[round(v, 2) for _, v in series]))

        # Broad breakdowns: each decoded dimension's top contributors (with share).
        if measure is not None:
            grand = measure.total or 0.0
            for dim in self._dimensions(table)[:3]:
                ranked = group_sum(table, dim, measure, top_n=10)
                ranked = [(k, v) for k, v in ranked if v != 0]
                if not ranked:
                    continue
                base = grand if grand > 0 else sum(v for _, v in ranked) or 1.0
                spec.tables.append(DataTable(
                    title=f"Top {friendly_name(dim.name)} by {friendly_name(measure.name)}",
                    headers=[friendly_name(dim.name),
                             f"Total {friendly_name(measure.name)}", "% of total"],
                    rows=[[k, round(v, 2), round(v / base * 100, 1)] for k, v in ranked],
                    formats=[NumberFormat.GENERAL, self._fmt(measure), NumberFormat.GENERAL],
                    bar_columns=[1]))

        # Currency split (LBP vs USD) when a dollar column is present.
        usd = next((c for c in table.value_measures if is_dollar_column(c)), None)
        lbp = next((c for c in table.value_measures if not is_dollar_column(c)), None)
        if usd is not None and lbp is not None:
            spec.tables.append(DataTable(
                title="Currency overview",
                headers=["Currency", "Total"],
                rows=[["Lebanese Pounds (LBP)", round(lbp.total or 0, 2)],
                      ["US Dollars ($)", round(usd.total or 0, 2)]],
                formats=[NumberFormat.GENERAL, NumberFormat.DECIMAL]))

        # Static Measure Statistics only when Excel is absent (else COM adds the
        # real pivot below the scorecard).
        if self.include_tables:
            spec.tables.append(DataTable(
                title="Measure Statistics",
                headers=["Measure", "Count", "Total", "Average", "Min", "Max"],
                rows=[[friendly_name(m.name), m.count, round(m.total or 0, 2),
                       round(m.mean or 0, 2), round(m.minimum or 0, 2),
                       round(m.maximum or 0, 2)] for m in table.measures],
                formats=[NumberFormat.GENERAL, NumberFormat.INTEGER, NumberFormat.DECIMAL,
                         NumberFormat.DECIMAL, NumberFormat.DECIMAL, NumberFormat.DECIMAL]))

        spec.text_blocks.append(TextBlock(
            "About this scorecard",
            ["Figures are month-grouped and shown in Lebanese Pounds (LBP) unless a "
             "column is in USD. Δ% compares each month with the one before it; "
             "year-over-year compares the latest year with the previous one."],
            style="normal"))
        return spec
