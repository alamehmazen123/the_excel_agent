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

from ..aggregate import group_sum, time_series
from ..constants import SHEET_INSIGHTS
from ..derived_metrics import find_applicable_metrics
from ..decode import friendly_name
from ..formatting import fmt_measure, fmt_number, fmt_percent, is_dollar_column
from ..insights import Insight, InsightKind, Severity, detect_insights
from ..library import get_library
from ..models import WorkbookProfile
from ..render import (ChartKind, ChartSpec, DataTable, KpiTile, NumberFormat,
                      SheetSpec, TextBlock)
from ..semantic import MeasureSemantic, MetricKind, ReportType, SemanticModel, analyze
from .base import Analyzer

_MONTHS = ["", "Jan", "Feb", "Mar", "Apr", "May", "Jun",
           "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

def _fmt_period(period: str) -> str:
    try:
        y, m = period.split("-")
        return f"{_MONTHS[int(m)]}-{y}"
    except Exception:
        return period

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

    # -- type-tailored scorecard (Phase 1.1) --------------------------------- #
    def _scorecard(self, profile: WorkbookProfile, sem: SemanticModel,
                   insights: list[Insight]) -> list[KpiTile]:
        table = profile.primary
        tiles: list[KpiTile] = []
        report = sem.report_type
        pm = sem.primary_money

        # --- REVENUE / FINANCIAL scorecard ---
        if report in (ReportType.FINANCIAL, ReportType.GENERIC) and pm is not None:
            if pm.column.total is not None:
                tiles.append(KpiTile(
                    label=f"Total {pm.meaning}",
                    value=fmt_measure(pm.column, pm.column.total),
                    caption=f"across {table.row_count:,} records", good=None))
            # Derived KPIs for revenue books
            dm_results = find_applicable_metrics(sem, table)
            for dm, avg_val in dm_results[:2]:
                tiles.append(KpiTile(
                    label=dm.name,
                    value=f"{avg_val:.1f}" if dm.unit == "number" else
                           f"{avg_val:,.0f} LBP" if dm.unit == "currency" else
                           f"{avg_val:.1%}" if dm.unit == "percent" else f"{avg_val:.1f}",
                    caption=dm.description, good=None))
            # MoM change.
            var = next((i for i in insights if i.kind == InsightKind.VARIANCE
                        and i.measure == (pm.name if pm else None)), None)
            if var is not None:
                prev = var.evidence.get("prev", 0)
                last = var.evidence.get("last", 0)
                pct = abs(last - prev) / abs(prev) * 100 if prev else 0
                arrow = "▲" if last >= prev else "▼"
                tiles.append(KpiTile(
                    label=f"{pm.meaning} MoM",
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
                            f"{pm.meaning if pm else 'total'}",
                    good=False if conc.severity == Severity.HIGH else None))
            return tiles[:4]

        # --- CENSUS / ADMISSIONS scorecard ---
        if report == ReportType.CENSUS:
            vols = sem.of_kind(MetricKind.VOLUME)
            if vols:
                v = vols[0]
                if v.column.total is not None:
                    tiles.append(KpiTile(
                        label=f"Total {v.meaning}",
                        value=f"{v.column.total:,.0f}",
                        caption=f"across {table.row_count:,} records", good=None))
            if len(vols) >= 2:
                v2 = vols[1]
                if v2.column.total is not None:
                    tiles.append(KpiTile(
                        label=f"Total {v2.meaning}",
                        value=f"{v2.column.total:,.0f}",
                        caption=f"across {table.row_count:,} records", good=None))
            dm_results = find_applicable_metrics(sem, table)
            for dm, avg_val in dm_results[:3]:
                unit_suffix = " days" if dm.unit == "days" else "" if dm.unit == "number" else "%"
                tiles.append(KpiTile(
                    label=dm.name,
                    value=f"{avg_val:.1f}{unit_suffix}",
                    caption=dm.description, good=None))
            if not tiles:
                tiles.append(KpiTile(label="Records", value=f"{table.row_count:,}",
                                     caption=f"{len(table.columns)} fields", good=None))
            return tiles[:4]

        # --- RECEIVABLES scorecard ---
        if report == ReportType.RECEIVABLES:
            bal = sem.balance
            if bal is not None and bal.column.total is not None:
                tiles.append(KpiTile(
                    label=f"Total {bal.meaning}",
                    value=fmt_measure(bal.column, bal.column.total),
                    caption="outstanding balance", good=None))
            aging = next((i for i in insights if i.kind == InsightKind.AGING), None)
            if aging is not None:
                buckets = aging.evidence.get("buckets", {})
                over90 = buckets.get("90+", 0)
                total_ar = sum(buckets.values()) or 1
                tiles.append(KpiTile(
                    label="90+ day aging",
                    value=f"{over90 / total_ar * 100:.0f}%",
                    caption=fmt_measure(bal.column if bal else None, over90) if bal else "",
                    good=False))
            dm_results = find_applicable_metrics(sem, table)
            for dm, avg_val in dm_results[:2]:
                tiles.append(KpiTile(
                    label=dm.name,
                    value=f"{avg_val:.1f}%" if dm.unit == "percent" else f"{avg_val:,.0f}",
                    caption=dm.description, good=None))
            if not tiles:
                tiles.append(KpiTile(label="Records", value=f"{table.row_count:,}",
                                     caption=f"{len(table.columns)} fields", good=None))
            return tiles[:4]

        # --- OPERATIONS scorecard ---
        if report == ReportType.OPERATIONS:
            vols = sem.of_kind(MetricKind.VOLUME)
            if vols:
                v = vols[0]
                if v.column.total is not None:
                    tiles.append(KpiTile(
                        label=f"Total {v.meaning}",
                        value=f"{v.column.total:,.0f}",
                        caption=f"across {table.row_count:,} records", good=None))
            if pm is not None and pm.column.total is not None:
                tiles.append(KpiTile(
                    label=f"Total {pm.meaning}",
                    value=fmt_measure(pm.column, pm.column.total),
                    caption=None, good=None))
            if not tiles:
                tiles.append(KpiTile(label="Records", value=f"{table.row_count:,}",
                                     caption=f"{len(table.columns)} fields", good=None))
            return tiles[:4]

        # --- GENERIC fallback ---
        if pm is not None and pm.column.total is not None:
            tiles.append(KpiTile(
                label=f"Total {pm.meaning}",
                value=fmt_measure(pm.column, pm.column.total),
                caption=f"across {table.row_count:,} records", good=None))
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
        conc = next((i for i in insights if i.kind == InsightKind.CONCENTRATION), None)
        if conc is not None:
            tiles.append(KpiTile(
                label=f"Top {conc.dimension}",
                value=str(conc.evidence.get("leader", "—")),
                caption=f"{conc.evidence.get('leader_share', 0) * 100:.0f}% of "
                        f"{sem.primary_money.meaning if sem.primary_money else 'total'}",
                good=False if conc.severity == Severity.HIGH else None))
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

    # -- account-category roll-up + data quality ---------------------------- #
    def _category_table(self, sem: SemanticModel) -> Optional[DataTable]:
        if not sem.category_totals:
            return None
        rows = sorted(sem.category_totals.items(), key=lambda kv: kv[1], reverse=True)
        total = sum(v for _, v in rows) or 1.0
        return DataTable(
            title="By account category (revenues vs expenses)",
            headers=["Account category", "Total", "% of total"],
            rows=[[c.title(), round(v, 2), round(v / total * 100, 1)] for c, v in rows],
            formats=[NumberFormat.GENERAL, NumberFormat.LBP, NumberFormat.GENERAL],
            bar_columns=[1])

    # -- payer mix analysis (Phase 1.2) ------------------------------------- #
    def _payer_mix(self, profile: WorkbookProfile, sem: SemanticModel) -> Optional[DataTable]:
        """When a decodeable payer field exists (FLD1 / guarantor), show a
        breakdown of revenue by payer with share %."""
        table = profile.primary
        if table is None:
            return None
        lib = get_library()
        pm = sem.primary_money
        if pm is None:
            return None
        # Find a decoded payer/guarantor helper column.
        payer_helper = None
        for c in table.columns:
            if c.is_decoded_helper and ("FLD1" in c.name or "fld1" in c.name):
                payer_helper = c
                break
        if payer_helper is None:
            for c in table.columns:
                if c.is_decoded_helper:
                    vals = set()
                    for r in table.rows:
                        v = r.get(c.name)
                        if v and isinstance(v, str) and v.strip():
                            vals.add(v.strip())
                            if len(vals) >= 5:
                                break
                    payer_keywords = ("insurance", "social security", "army", "private",
                                      "nssf", "cooperative", "union", "employer")
                    for v in vals:
                        vl = v.lower()
                        if any(k in vl for k in payer_keywords):
                            payer_helper = c
                            break
                if payer_helper is not None:
                    break

        if payer_helper is None:
            return None

        ranked = group_sum(table, payer_helper, pm.column, top_n=10)
        ranked = [(k, v) for k, v in ranked if v > 0]
        if not ranked:
            return None
        total = sum(v for _, v in ranked) or 1.0
        rows = []
        for k, v in ranked:
            share = v / total * 100
            rows.append([k, round(v, 2), round(share, 1)])
        return DataTable(
            title="Revenue by payer / guarantor",
            headers=[friendly_name(payer_helper.name), f"Total {pm.meaning}", "% of total"],
            rows=rows,
            formats=[NumberFormat.GENERAL, NumberFormat.LBP, NumberFormat.GENERAL],
            bar_columns=[1])

    # -- type-aware data quality (Phase 1.6) --------------------------------- #
    def _data_quality(self, profile: WorkbookProfile, sem: SemanticModel) -> TextBlock:
        import datetime as _dt  # noqa: PLC0415
        table = profile.primary
        lib = get_library()
        lines = [f"Records: {table.row_count:,}   ·   Fields: {len(table.columns)}"]

        if table.date_columns:
            dates = [r.get(table.date_columns[0].name) for r in table.rows]
            dts = [d for d in dates if isinstance(d, (_dt.date, _dt.datetime))]
            if dts:
                lines.append(f"Date range: {min(dts).strftime('%b %Y')} → "
                             f"{max(dts).strftime('%b %Y')}")
                periods_set = set()
                for d in dts:
                    periods_set.add(f"{d.year}-{d.month:02d}")
                lines.append(f"Periods covered: {len(periods_set)} months")

        # Type-aware quality checks.
        report = sem.report_type

        if report == ReportType.FINANCIAL:
            # Account code decode coverage.
            if sem.account_column:
                col = table.column(sem.account_column)
                if col is not None:
                    seen, unknown = set(), set()
                    for r in table.rows:
                        v = r.get(col.name)
                        if v in (None, ""):
                            continue
                        k = str(v)
                        if k in seen:
                            continue
                        seen.add(k)
                        name = lib.decode("account", v)
                        if name == k:
                            unknown.add(k)
                    total = len(seen)
                    known = total - len(unknown)
                    pct = (known / total * 100) if total else 100
                    lines.append(f"Account codes decoded: {known}/{total} ({pct:.0f}%).")
                    if unknown:
                        sample = ", ".join(sorted(unknown)[:8])
                        lines.append(f"⚠ {len(unknown)} unrecognised code(s): {sample}"
                                     + (" …" if len(unknown) > 8 else "")
                                     + ". Send these to extend the library.")
            # Check for negative values in revenue.
            rev = sem.revenue
            if rev is not None:
                neg_count = 0
                for r in table.rows:
                    v = r.get(rev.name)
                    if isinstance(v, (int, float)) and v < 0:
                        neg_count += 1
                if neg_count:
                    lines.append(f"⚠ {neg_count} negative revenue record(s) "
                                 f"({neg_count / max(1, table.row_count) * 100:.0f}%) "
                                 "— revenue sign convention is applied.")

        elif report == ReportType.CENSUS:
            if table.date_columns:
                dates_only = sorted([d for d in dates if isinstance(d, (_dt.date, _dt.datetime))])
                if len(dates_only) >= 2:
                    min_d, max_d = min(dates_only), max(dates_only)
                    expected_months = ((max_d.year - min_d.year) * 12 +
                                       (max_d.month - min_d.month) + 1)
                    if len(periods_set) < expected_months:
                        missing = expected_months - len(periods_set)
                        lines.append(f"⚠ {missing} missing month(s) of data "
                                     f"({len(periods_set)}/{expected_months} months present).")

        elif report == ReportType.RECEIVABLES:
            payer_col = next((c for c in table.dimensions
                              if any(k in c.name.lower() for k in
                                     ("payer", "guarantor", "insurance", "fld1"))), None)
            if payer_col:
                missing = sum(1 for r in table.rows if r.get(payer_col.name) in (None, ""))
                if missing:
                    lines.append(f"⚠ {missing} record(s) missing payer code "
                                 f"({missing / max(1, table.row_count) * 100:.0f}%).")

        elif report == ReportType.OPERATIONS:
            if table.date_columns and sem.of_kind(MetricKind.VOLUME):
                vol = sem.of_kind(MetricKind.VOLUME)[0]
                series = time_series(table, table.date_columns[0], vol.column)
                if len(series) >= 3:
                    vals = [v for _, v in series]
                    avg_v = sum(vals) / len(vals)
                    max_v = max(vals)
                    min_v = min(vals)
                    if avg_v > 0:
                        lines.append(f"Volume range: {min_v:,.0f} – {max_v:,.0f} "
                                     f"(avg {avg_v:,.0f}/month)")

        lines.append(f"Detected type: {_REPORT_LABEL.get(report, 'Data report')}")
        return TextBlock("Data quality", lines, style="normal")

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

        # Known-event context (e.g. the 2026 closure) — explain low figures up front.
        if table.date_columns and sem.primary_money is not None:
            from ..context import events_overlapping  # noqa: PLC0415
            periods = [p for p, _ in time_series(
                table, table.date_columns[0], sem.primary_money.column)]
            for ev in events_overlapping(periods):
                spec.text_blocks.append(TextBlock(
                    "Important context", [ev.note], style="warn"))

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

        # Payer mix analysis (Phase 1.2) — when a decodeable payer field exists.
        payer = self._payer_mix(profile, sem)
        if payer is not None:
            spec.tables.append(payer)

        # Account-category roll-up (revenues vs expense categories) from the
        # library categories — a quick P&L-style view.
        cat = self._category_table(sem)
        if cat is not None:
            spec.tables.append(cat)

        # Data-quality panel so the GM can trust the numbers at a glance.
        spec.text_blocks.append(self._data_quality(profile, sem))

        spec.charts = self._charts(table, sem, insights)

        if not insights:
            spec.text_blocks.append(TextBlock(
                "Note", ["No material variance, concentration or ageing was "
                         "detected in this period — the operation looks stable. "
                         "Use the scorecard above to keep monitoring."],
                style="normal"))
        return spec
