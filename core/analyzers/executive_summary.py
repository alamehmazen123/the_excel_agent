"""Executive Summary analyzer -- a consultant-grade briefing.

Builds a RICH metrics payload (leaders/laggards, concentration, trend, negative
contributors), asks the Groq LLM for a structured JSON briefing, and renders it
into styled sections. If the LLM is unavailable, a strong deterministic briefing
is produced from the same metrics, so the sheet is ALWAYS useful.
"""
from __future__ import annotations

import json
from typing import Any, Callable, Optional

from ..aggregate import group_sum, period_over_period_growth, time_series
from ..constants import SHEET_SUMMARY
from ..formatting import fmt_measure
from ..models import ColumnProfile, ColumnType, TableProfile, WorkbookProfile
from ..render import SheetSpec, TextBlock
from .base import Analyzer

Narrator = Callable[[dict[str, Any]], Optional[str]]


def _unit(col: ColumnProfile) -> str:
    if col.ctype == ColumnType.CURRENCY:
        return "currency"
    if col.ctype == ColumnType.PERCENT:
        return "percent"
    return "number"


def _negatives(table: TableProfile, col: ColumnProfile) -> tuple[int, float]:
    cnt = 0
    tot = 0.0
    for row in table.rows:
        v = row.get(col.name)
        if isinstance(v, (int, float)) and not isinstance(v, bool) and v < 0:
            cnt += 1
            tot += float(v)
    return cnt, tot


class ExecutiveSummaryAnalyzer(Analyzer):
    key = "executive_summary"
    sheet_name = SHEET_SUMMARY

    def __init__(self, narrator: Optional[Narrator] = None) -> None:
        self._narrator = narrator
        self.used_llm = False
        self.note: Optional[str] = None

    def applies_to(self, profile: WorkbookProfile) -> bool:
        return profile.primary is not None and profile.primary.row_count > 0

    # -- metrics --------------------------------------------------------------
    def build_metrics(self, profile: WorkbookProfile) -> dict[str, Any]:
        table = profile.primary
        metrics: dict[str, Any] = {
            "source_sheet": table.sheet_name,
            "record_count": table.row_count,
            "column_count": len(table.columns),
            "measures": [],
            "breakdowns": [],
        }

        for m in table.key_measures[:5]:
            entry: dict[str, Any] = {
                "name": m.name, "unit": _unit(m),
                "total": round(m.total or 0, 2),
                "total_display": fmt_measure(m, m.total or 0),
                "average_display": fmt_measure(m, m.mean or 0),
                "min_display": fmt_measure(m, m.minimum or 0),
                "max_display": fmt_measure(m, m.maximum or 0),
            }
            if table.date_columns:
                series = time_series(table, table.date_columns[0], m)
                growth = period_over_period_growth(series)
                if growth is not None:
                    entry["period_growth_pct"] = round(growth * 100, 1)
            ncnt, nsum = _negatives(table, m)
            if ncnt:
                entry["negative_count"] = ncnt
                entry["negative_total_display"] = fmt_measure(m, nsum)
            metrics["measures"].append(entry)

        measure = table.primary_value_measure
        if measure is not None:
            grand = measure.total or 0.0
            for dim in table.pivot_dimensions[:3]:
                ranked = group_sum(table, dim, measure, top_n=10_000)
                ranked.sort(key=lambda kv: kv[1], reverse=True)
                if not ranked:
                    continue
                top = ranked[:5]
                top1 = ranked[0]
                top3sum = sum(v for _, v in ranked[:3])

                # Share-of-total is only meaningful when the grand total is
                # positive and large vs the leader. For measures that net to ~0
                # or go negative (e.g. PnL), use the positive mass as the base
                # so percentages stay sane (0-100) instead of exploding.
                positives = sum(v for _, v in ranked if v > 0)
                if grand > 0 and grand >= 0.5 * max(top1[1], 1e-9):
                    base = grand
                else:
                    base = positives

                def share(v: float, _base: float = base) -> Optional[float]:
                    if _base <= 0 or v < 0:
                        return None
                    s = round(v / _base * 100, 1)
                    return s if 0 <= s <= 100 else None

                metrics["breakdowns"].append({
                    "dimension": dim.name, "measure": measure.name,
                    "distinct": dim.distinct,
                    "top_items": [{"label": k, "value_display": fmt_measure(measure, v),
                                   "share_pct": share(v)} for k, v in top],
                    "bottom_item": {"label": ranked[-1][0],
                                    "value_display": fmt_measure(measure, ranked[-1][1])},
                    "leader_share_pct": share(top1[1]),
                    "top3_share_pct": share(top3sum),
                })

            if table.date_columns:
                series = time_series(table, table.date_columns[0], measure)
                if len(series) >= 2:
                    best = max(series, key=lambda kv: kv[1])
                    worst = min(series, key=lambda kv: kv[1])
                    metrics["trend"] = {
                        "measure": measure.name, "periods": len(series),
                        "first": {"period": series[0][0],
                                  "value_display": fmt_measure(measure, series[0][1])},
                        "last": {"period": series[-1][0],
                                 "value_display": fmt_measure(measure, series[-1][1])},
                        "best": {"period": best[0],
                                 "value_display": fmt_measure(measure, best[1])},
                        "worst": {"period": worst[0],
                                  "value_display": fmt_measure(measure, worst[1])},
                    }
        return metrics

    # -- deterministic briefing (offline fallback) ----------------------------
    def _deterministic(self, metrics: dict[str, Any]) -> dict[str, Any]:
        ms = metrics["measures"]
        bds = metrics.get("breakdowns", [])
        pm = ms[0] if ms else None
        out: dict[str, Any] = {"headline": "", "overview": "", "findings": [],
                               "risks": [], "opportunities": [], "actions": []}

        if pm:
            out["headline"] = (f"{pm['name']} totals {pm['total_display']} "
                               f"across {metrics['record_count']:,} records.")
        out["overview"] = (
            f"This briefing covers {metrics['record_count']:,} records and "
            f"{metrics['column_count']} fields from '{metrics['source_sheet']}'. "
            + (f"The primary measure, {pm['name']}, totals {pm['total_display']} "
               f"(avg {pm['average_display']}, range {pm['min_display']}–{pm['max_display']})."
               if pm else ""))

        for m in ms:
            line = (f"{m['name']}: total {m['total_display']}, average "
                    f"{m['average_display']}.")
            if "period_growth_pct" in m:
                g = m["period_growth_pct"]
                line += f" {'Up' if g >= 0 else 'Down'} {abs(g)}% vs the prior period."
            out["findings"].append(line)

        for b in bds[:2]:
            lead = b["top_items"][0]
            line = (f"In {b['dimension']}, '{lead['label']}' leads {b['measure']} at "
                    f"{lead['value_display']}")
            if lead.get("share_pct") is not None:
                line += f" ({lead['share_pct']}% of total)"
            if b.get("top3_share_pct") is not None:
                line += f"; the top 3 account for {b['top3_share_pct']}%"
            out["findings"].append(line + ".")
            if b.get("leader_share_pct") and b["leader_share_pct"] >= 40:
                out["risks"].append(
                    f"Concentration risk: '{lead['label']}' alone is "
                    f"{b['leader_share_pct']}% of {b['measure']} — heavy dependence "
                    f"on a single {b['dimension']}.")
            out["opportunities"].append(
                f"Lift under-performers in {b['dimension']} (e.g. '{b['bottom_item']['label']}' "
                f"at {b['bottom_item']['value_display']}) toward the leader's level.")

        for m in ms:
            if m.get("period_growth_pct", 0) < 0:
                out["risks"].append(
                    f"{m['name']} declined {abs(m['period_growth_pct'])}% versus the "
                    f"prior period — investigate the cause.")
            if m.get("negative_count"):
                out["risks"].append(
                    f"{m['name']} has {m['negative_count']} negative records totalling "
                    f"{m['negative_total_display']} — review these loss-making items.")

        tr = metrics.get("trend")
        if tr:
            out["findings"].append(
                f"{tr['measure']} ranged from {tr['worst']['value_display']} "
                f"({tr['worst']['period']}) to {tr['best']['value_display']} "
                f"({tr['best']['period']}) across {tr['periods']} periods.")
            out["opportunities"].append(
                f"Study what drove the best period ({tr['best']['period']}) and "
                f"replicate it.")

        # Prioritized action plan.
        if bds:
            b = bds[0]
            out["actions"].append(
                f"Double down on the top {b['dimension']} ('{b['top_items'][0]['label']}') "
                f"while diversifying to reduce concentration.")
        for m in ms:
            if m.get("period_growth_pct", 0) < 0:
                out["actions"].append(
                    f"Launch a focused review of {m['name']} to reverse its decline.")
                break
        for m in ms:
            if m.get("negative_count"):
                out["actions"].append(
                    f"Audit the {m['negative_count']} negative {m['name']} records and "
                    f"set controls to prevent recurrence.")
                break
        out["actions"].append(
            "Use the Pivot Analysis and Dashboard sheets to drill into the drivers "
            "behind each figure before the next planning cycle.")
        out["actions"].append(
            "Set targets for the lagging categories and track them monthly.")
        # de-duplicate / trim
        for k in ("findings", "risks", "opportunities", "actions"):
            seen, uniq = set(), []
            for s in out[k]:
                if s not in seen:
                    seen.add(s); uniq.append(s)
            out[k] = uniq[:5]
        if not out["risks"]:
            out["risks"].append("No major concentration or decline detected in the "
                                 "current data; continue monitoring monthly.")
        return out

    # -- rendering ------------------------------------------------------------
    def _spec_from_briefing(self, metrics: dict[str, Any],
                            b: dict[str, Any]) -> SheetSpec:
        spec = SheetSpec(
            name=SHEET_SUMMARY, heading="Executive Summary",
            subheading=f"Source: {metrics['source_sheet']}  •  "
                       f"{metrics['record_count']:,} records",
        )
        if b.get("headline"):
            spec.text_blocks.append(TextBlock("Bottom Line", [b["headline"]],
                                              style="highlight"))
        if b.get("overview"):
            spec.text_blocks.append(TextBlock("Overview", [b["overview"]],
                                              style="normal"))
        if b.get("findings"):
            spec.text_blocks.append(TextBlock("Key Findings", b["findings"],
                                              style="highlight"))
        if b.get("risks"):
            spec.text_blocks.append(TextBlock("Risks & Watch-outs", b["risks"],
                                              style="warn"))
        if b.get("opportunities"):
            spec.text_blocks.append(TextBlock("Opportunities", b["opportunities"],
                                              style="highlight"))
        if b.get("actions"):
            spec.text_blocks.append(TextBlock("Recommended Action Plan", b["actions"],
                                              style="recommend"))
        return spec

    def run(self, profile: WorkbookProfile) -> Optional[SheetSpec]:
        if profile.primary is None:
            return None
        metrics = self.build_metrics(profile)

        briefing: Optional[dict[str, Any]] = None
        if self._narrator is not None:
            try:
                content = self._narrator(metrics)
                if content:
                    parsed = json.loads(content)
                    if isinstance(parsed, dict) and parsed.get("actions"):
                        briefing = parsed
            except Exception:
                briefing = None

        if briefing is not None:
            self.used_llm = True
        else:
            self.used_llm = False
            self.note = ("AI briefing unavailable — this professional summary was "
                         "generated from the computed metrics.")
            briefing = self._deterministic(metrics)

        spec = self._spec_from_briefing(metrics, briefing)
        if self.note:
            spec.text_blocks.append(TextBlock("Note", [self.note], style="normal"))
        return spec
