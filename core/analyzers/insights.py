"""Insights analyzer — the headline sheet leadership sees first.

This is the product's brain made visible. It runs the semantic layer and the
insight engine, then renders a one-glance briefing:

* a **KPI scorecard** of the headline numbers with red/amber/green status;
* a ranked **"What to look at"** list — the few findings that actually matter;
* **risks** (declines, concentration, ageing) called out in amber;
* a compact **Findings** table (priority · finding · measure) the writer turns
  into a Smart Table with data bars.

Every number is computed by :mod:`core.insights`; this analyzer only arranges
them. It always produces something useful — when nothing notable is detected it
still shows the scorecard and a "stable, keep monitoring" note.
"""
from __future__ import annotations

from typing import Optional

from ..aggregate import time_series
from ..constants import SHEET_INSIGHTS
from ..formatting import fmt_measure, fmt_number
from ..insights import Insight, InsightKind, Severity, detect_insights
from ..library import get_library
from ..models import WorkbookProfile
from ..render import (ChartKind, ChartSpec, DataTable, KpiTile, NumberFormat,
                      SheetSpec, TextBlock)
from ..semantic import MeasureSemantic, MetricKind, ReportType, SemanticModel, analyze
from .base import Analyzer

_REPORT_LABEL = {
    ReportType.FINANCIAL: "Financial / general-ledger report",
    ReportType.RECEIVABLES: "Receivables / ageing report",
    ReportType.CENSUS: "Admissions / census report",
    ReportType.OPERATIONS: "Operations / volume report",
    ReportType.GENERIC: "Data report",
}

_PRIORITY = {Severity.HIGH: "● High", Severity.WATCH: "▲ Watch", Severity.INFO: "○ Info"}


class InsightsAnalyzer(Analyzer):
    key = "insights"
    sheet_name = SHEET_INSIGHTS

    def __init__(self) -> None:
        # Exposed to the pipeline after run() (like ExecutiveSummary.used_llm).
        self.insights: list[Insight] = []
        self.semantic: Optional[SemanticModel] = None

    def applies_to(self, profile: WorkbookProfile) -> bool:
        t = profile.primary
        return bool(t and t.row_count > 0 and (t.value_measures or t.percent_measures))

    # -- scorecard ---------------------------------------------------------- #
    def _scorecard(self, profile: WorkbookProfile, sem: SemanticModel,
                   insights: list[Insight]) -> list[KpiTile]:
        table = profile.primary
        tiles: list[KpiTile] = []

        pm = sem.primary_money
        if pm is not None and pm.column.total is not None:
            tiles.append(KpiTile(
                label=f"Total {pm.meaning}",
                value=fmt_measure(pm.column, pm.column.total),
                caption=f"across {table.row_count:,} records", good=None))

        # Month-over-month for the primary money measure (RAG by direction).
        var = next((i for i in insights if i.kind == InsightKind.VARIANCE
                    and i.measure == (pm.name if pm else None)), None)
        if var is not None:
            prev = var.evidence.get("prev", 0)
            last = var.evidence.get("last", 0)
            pct = abs(last - prev) / abs(prev) * 100 if prev else 0
            arrow = "▲" if last >= prev else "▼"
            tiles.append(KpiTile(
                label=f"{pm.meaning} MoM" if pm else "Change",
                value=f"{arrow} {pct:.0f}%",
                caption=f"{var.period} vs prior month", good=var.good))
        elif pm is not None:
            tiles.append(self._trend_tile(table, sem))

        # Concentration leader.
        conc = next((i for i in insights if i.kind == InsightKind.CONCENTRATION), None)
        if conc is not None:
            tiles.append(KpiTile(
                label=f"Top {conc.dimension}",
                value=str(conc.evidence.get("leader", "—")),
                caption=f"{conc.evidence.get('leader_share', 0) * 100:.0f}% of "
                        f"{sem.primary_money.meaning if sem.primary_money else 'total'}",
                good=False if conc.severity == Severity.HIGH else None))

        # Ageing (receivables) OR record-count fallback.
        aging = next((i for i in insights if i.kind == InsightKind.AGING), None)
        if aging is not None:
            buckets = aging.evidence.get("buckets", {})
            over90 = buckets.get("90+", 0)
            tiles.append(KpiTile(
                label="90+ day balance",
                value=fmt_number(over90),
                caption=aging.title, good=False))
        else:
            tiles.append(KpiTile(label="Records", value=f"{table.row_count:,}",
                                 caption=f"{len(table.columns)} fields", good=None))
        return tiles[:4]

    def _trend_tile(self, table, sem: SemanticModel) -> KpiTile:
        pm = sem.primary_money
        series = time_series(table, table.date_columns[0], pm.column) if table.date_columns else []
        if len(series) >= 2:
            first, last = series[0][1], series[-1][1]
            arrow = "▲" if last >= first else "▼"
            return KpiTile(label=f"{pm.meaning} trend",
                           value=f"{arrow} {fmt_measure(pm.column, last)}",
                           caption=f"{series[0][0]} → {series[-1][0]}",
                           good=None)
        return KpiTile(label=pm.meaning, value=fmt_measure(pm.column, pm.column.total or 0),
                       caption="period total", good=None)

    # -- narrative bits ----------------------------------------------------- #
    def _actions(self, insights: list[Insight], sem: SemanticModel) -> list[str]:
        actions: list[str] = []
        for ins in insights:
            if ins.kind == InsightKind.VARIANCE and ins.good is False:
                drv = ins.evidence.get("driver")
                actions.append(
                    f"Investigate the drop in {ins.measure}"
                    + (f", driven by {drv}" if drv else "")
                    + ", before the next close.")
            elif ins.kind == InsightKind.CONCENTRATION and ins.severity == Severity.HIGH:
                actions.append(
                    f"Reduce dependence on {ins.evidence.get('leader')} "
                    f"({ins.evidence.get('leader_share', 0) * 100:.0f}% of "
                    f"{ins.measure}) by growing the next tier of {ins.dimension}.")
            elif ins.kind == InsightKind.AGING:
                actions.append(
                    "Launch focused collection on the 90+ day receivables to "
                    "protect cash before they become bad debt.")
            elif ins.kind == InsightKind.LOSS:
                actions.append(
                    f"Audit the negative {ins.measure} records and add a control "
                    "to prevent recurrence.")
            if len(actions) >= 4:
                break
        if not actions:
            actions.append("No urgent action: the figures are stable. Keep "
                           "tracking the scorecard monthly against targets.")
        # de-dup, keep order
        seen, out = set(), []
        for a in actions:
            if a not in seen:
                seen.add(a); out.append(a)
        return out

    # -- charts ------------------------------------------------------------- #
    def _charts(self, table, sem: SemanticModel,
                insights: list[Insight]) -> list[ChartSpec]:
        charts: list[ChartSpec] = []

        # Pareto of the strongest concentration: bars (value) + cumulative % line.
        conc = next((i for i in insights if i.kind == InsightKind.CONCENTRATION), None)
        if conc is not None:
            items = conc.evidence.get("items", [])[:8]
            total = sum(v for _, v in items) or 1.0
            cats = [str(k) for k, _ in items]
            vals = [round(float(v), 2) for _, v in items]
            cum, cum_pct = 0.0, []
            for v in vals:
                cum += v
                cum_pct.append(round(cum / total * 100, 1))
            charts.append(ChartSpec(
                kind=ChartKind.PARETO,
                title=f"{sem.primary_money.meaning if sem.primary_money else 'Value'} "
                      f"concentration by {conc.dimension}",
                categories=cats, series_name="Value", values=vals,
                line_values=cum_pct, line_name="Cumulative %"))

        # Trend + forecast of the primary money measure.
        trend = next((i for i in insights if i.kind == InsightKind.TREND), None)
        pm = sem.primary_money
        if pm is not None and table.date_columns:
            series = time_series(table, table.date_columns[0], pm.column)
            if len(series) >= 3:
                cats = [p for p, _ in series]
                vals = [round(float(v), 2) for _, v in series]
                if trend is not None and "forecast" in trend.evidence:
                    cats = cats + ["→ next"]
                    vals = vals + [round(float(trend.evidence["forecast"]), 2)]
                charts.append(ChartSpec(
                    kind=ChartKind.LINE,
                    title=f"{pm.meaning} trend"
                          + (" + forecast" if trend is not None else ""),
                    categories=cats, series_name=pm.meaning, values=vals))
        return charts

    # -- run ---------------------------------------------------------------- #
    def run(self, profile: WorkbookProfile) -> Optional[SheetSpec]:
        table = profile.primary
        if table is None:
            return None

        sem = analyze(profile, get_library())
        insights = detect_insights(profile, sem)
        self.insights = insights
        self.semantic = sem

        purpose_txt = (f"Purpose: {sem.purpose} report" if sem.purpose
                       else _REPORT_LABEL.get(sem.report_type))
        spec = SheetSpec(
            name=SHEET_INSIGHTS, heading="Insights",
            subheading=(f"{purpose_txt}  •  "
                        f"Source: {table.sheet_name}  •  {table.row_count:,} records"),
        )
        spec.kpi_tiles = self._scorecard(profile, sem, insights)

        # State the detected purpose up front, with the evidence behind it.
        if sem.purpose and sem.category_totals:
            top_cats = sorted(sem.category_totals.items(),
                              key=lambda kv: kv[1], reverse=True)[:3]
            cats = ", ".join(c for c, _ in top_cats)
            spec.text_blocks.append(TextBlock(
                "What this workbook is about",
                [f"This looks like a {sem.purpose.upper()} report — the account "
                 f"codes resolve mainly to: {cats}. The analysis below is framed "
                 f"accordingly."],
                style="highlight"))

        # Bottom line = the single most important finding.
        if insights:
            top = insights[0]
            spec.text_blocks.append(TextBlock(
                "Bottom Line", [f"{top.title}. {top.detail}"], style="highlight"))

        # What to look at — the ranked findings (skip the one used as Bottom Line).
        look = [f"{i.title} — {i.detail}" for i in insights[1:6]]
        if look:
            spec.text_blocks.append(TextBlock(
                "What to look at", look, style="highlight"))

        # Risks — anything bad at Watch+ severity.
        risks = [i.detail for i in insights
                 if i.good is False and i.severity.rank >= Severity.WATCH.rank]
        if risks:
            spec.text_blocks.append(TextBlock(
                "Risks & Watch-outs", risks[:5], style="warn"))

        spec.text_blocks.append(TextBlock(
            "Recommended actions", self._actions(insights, sem), style="recommend"))

        # Findings table (the writer renders it as a Smart Table with data bars).
        if insights:
            spec.tables.append(DataTable(
                title="All findings (ranked by priority)",
                headers=["Priority", "Finding", "Measure", "Impact"],
                rows=[[_PRIORITY[i.severity], i.title, i.measure or "—",
                       round(i.score * 100, 0)] for i in insights],
                formats=[NumberFormat.GENERAL, NumberFormat.GENERAL,
                         NumberFormat.GENERAL, NumberFormat.INTEGER],
                bar_columns=[3],          # data bar on the Impact score
            ))

        spec.charts = self._charts(table, sem, insights)

        if not insights:
            spec.text_blocks.append(TextBlock(
                "Note", ["No material variance, concentration or ageing was "
                         "detected in this period — the operation looks stable. "
                         "Use the scorecard above to keep monitoring."],
                style="normal"))
        return spec
