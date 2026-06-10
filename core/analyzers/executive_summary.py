"""Executive Summary analyzer.

Builds a compact metrics payload from the profiled data, then asks an optional
``narrator`` (the Groq LLM) to turn it into prose. If no narrator is supplied or
it fails, a deterministic, stats-driven template is used instead so the sheet is
ALWAYS produced.
"""
from __future__ import annotations

from typing import Any, Callable, Optional

from ..aggregate import group_sum, period_over_period_growth, time_series
from ..constants import SHEET_SUMMARY
from ..formatting import fmt_measure, fmt_percent
from ..models import WorkbookProfile
from ..render import SheetSpec, TextBlock
from .base import Analyzer

# narrator(metrics) -> narrative text, or None on failure.
Narrator = Callable[[dict[str, Any]], Optional[str]]


class ExecutiveSummaryAnalyzer(Analyzer):
    key = "executive_summary"
    sheet_name = SHEET_SUMMARY

    def __init__(self, narrator: Optional[Narrator] = None) -> None:
        self._narrator = narrator
        self.used_llm = False
        self.note: Optional[str] = None

    def applies_to(self, profile: WorkbookProfile) -> bool:
        return profile.primary is not None and profile.primary.row_count > 0

    def build_metrics(self, profile: WorkbookProfile) -> dict[str, Any]:
        table = profile.primary
        metrics: dict[str, Any] = {
            "source_sheet": table.sheet_name,
            "record_count": table.row_count,
            "column_count": len(table.columns),
            "measures": [],
            "top_breakdowns": [],
        }
        for m in table.measures[:5]:
            entry = {
                "name": m.name, "total": round(m.total or 0, 2),
                "average": round(m.mean or 0, 2),
                "min": round(m.minimum or 0, 2), "max": round(m.maximum or 0, 2),
            }
            if table.date_columns:
                series = time_series(table, table.date_columns[0], m)
                growth = period_over_period_growth(series)
                if growth is not None:
                    entry["period_growth_pct"] = round(growth * 100, 1)
            metrics["measures"].append(entry)

        if table.dimensions and table.measures:
            dim, measure = table.dimensions[0], table.measures[0]
            ranked = group_sum(table, dim, measure, top_n=5)
            metrics["top_breakdowns"].append({
                "dimension": dim.name, "measure": measure.name,
                "items": [{"label": k, "value": round(v, 2)} for k, v in ranked],
            })
        return metrics

    def _template_narrative(self, profile: WorkbookProfile,
                            metrics: dict[str, Any]) -> list[str]:
        table = profile.primary
        paras = [
            f"This report analyzes {metrics['record_count']:,} records across "
            f"{metrics['column_count']} fields from the '{metrics['source_sheet']}' sheet."
        ]
        for m, col in zip(metrics["measures"], table.measures):
            line = (f"{m['name']}: total {fmt_measure(col, m['total'])}, "
                    f"averaging {fmt_measure(col, m['average'])} per record "
                    f"(range {fmt_measure(col, m['min'])}–{fmt_measure(col, m['max'])}).")
            if "period_growth_pct" in m:
                g = m["period_growth_pct"] / 100
                direction = "increased" if g >= 0 else "decreased"
                line += f" It {direction} {fmt_percent(abs(g))} versus the prior period."
            paras.append(line)

        for b in metrics["top_breakdowns"]:
            if b["items"]:
                top = b["items"][0]
                paras.append(
                    f"The leading {b['dimension']} by {b['measure']} is "
                    f"'{top['label']}'. See the Pivot Analysis sheet for the full breakdown."
                )
        paras.append(
            "Recommended next steps: focus on the top contributors above, "
            "investigate any period-over-period declines, and review the "
            "Dashboard sheet for visual trends."
        )
        return paras

    def run(self, profile: WorkbookProfile) -> Optional[SheetSpec]:
        if profile.primary is None:
            return None
        metrics = self.build_metrics(profile)

        narrative: Optional[str] = None
        if self._narrator is not None:
            try:
                narrative = self._narrator(metrics)
            except Exception:
                narrative = None

        spec = SheetSpec(
            name=SHEET_SUMMARY, heading="Executive Summary",
            subheading=f"Source: {metrics['source_sheet']}  •  {metrics['record_count']:,} records",
        )

        if narrative:
            self.used_llm = True
            paragraphs = [p.strip() for p in narrative.split("\n") if p.strip()]
            spec.text_blocks.append(TextBlock(title="Overview", paragraphs=paragraphs))
        else:
            self.used_llm = False
            self.note = ("AI narrative unavailable — summary generated from "
                         "computed metrics.")
            spec.text_blocks.append(TextBlock(
                title="Overview",
                paragraphs=self._template_narrative(profile, metrics),
            ))
            spec.text_blocks.append(TextBlock(title="Note", paragraphs=[self.note]))
        return spec
