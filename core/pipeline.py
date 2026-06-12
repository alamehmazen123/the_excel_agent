"""The engine's single public entry point.

Every front-end (desktop now; Ribbon/web/API later) calls ``Engine.run``.
It knows nothing about how it was invoked.
"""
from __future__ import annotations

from typing import Optional

from .analyzers.base import Analyzer
from .analyzers.dashboard import DashboardAnalyzer
from .analyzers.executive_summary import ExecutiveSummaryAnalyzer, Narrator
from .analyzers.kpi import KpiAnalyzer
from .analyzers.pivot import PivotAnalyzer
from .analyzers.smart_tables import SmartTablesAnalyzer
from .constants import SHEET_KPI, SHEET_PIVOT
from .excel_com import ExcelFinalizer, excel_available
from .loader import load_workbook_profile
from .pivot_plan import build_pivot_plan
from .models import (AnalysisOptions, AnalysisResult, ProgressCallback,
                     TableProfile, WorkbookProfile, _noop_progress)
from .writer import append_sheets, write_results


class Engine:
    """Profile a workbook and append the requested analysis sheets."""

    def __init__(self, narrator: Optional[Narrator] = None) -> None:
        # narrator is the optional Groq callable; None => deterministic summaries.
        self._narrator = narrator

    # -- introspection used by the UI to enable/disable options ---------------
    def profile(self, path: str) -> WorkbookProfile:
        return load_workbook_profile(path)

    def applicable_options(self, profile: WorkbookProfile) -> dict[str, bool]:
        """Which analyses are meaningful for this workbook."""
        return {a.key: a.applies_to(profile) for a in self._all_analyzers()}

    def describe_columns(self, profile: WorkbookProfile,
                         sheet_name: Optional[str] = None) -> dict:
        """Describe a table's columns for the Custom wizard.

        Returns the data sheets, plus the groupable dimensions and the value
        measures of the chosen table, with a recommended pre-selection.
        """
        from .models import ColumnType  # noqa: PLC0415
        sheets = [t.sheet_name for t in profile.tables]
        table = None
        if sheet_name:
            table = next((t for t in profile.tables if t.sheet_name == sheet_name), None)
        table = table or profile.primary
        if table is None:
            return {"sheets": sheets, "sheet": None, "dimensions": [], "measures": []}

        recommended_dims = {c.name for c in table.pivot_dimensions[:6]}
        dims = []
        for c in table.pivot_dimensions:
            if c.ctype == ColumnType.DATE:
                kind, detail = "date", "grouped by month/year"
            elif TableProfile.is_wide_dimension(c):
                kind, detail = "wide", f"{c.distinct} values · top 20"
            else:
                kind, detail = "category", f"{c.distinct} values"
            dims.append({"name": c.name, "kind": kind, "detail": detail,
                         "recommended": c.name in recommended_dims})
        # dates are also offered as dimensions (grouped) -> include any not already
        for c in table.date_columns:
            if not any(d["name"] == c.name for d in dims):
                dims.append({"name": c.name, "kind": "date",
                             "detail": "grouped by month/year", "recommended": True})

        rec_measures = {c.name for c in table.key_measures[:3]}
        measures = []
        for c in table.value_measures + table.percent_measures:
            unit = ("percent" if c.ctype == ColumnType.PERCENT
                    else "currency" if c.ctype == ColumnType.CURRENCY else "number")
            measures.append({"name": c.name, "unit": unit,
                             "recommended": c.name in rec_measures})
        return {"sheets": sheets, "sheet": table.sheet_name,
                "dimensions": dims, "measures": measures}

    # -- main run -------------------------------------------------------------
    def run(self, workbook_path: str, options: AnalysisOptions,
            progress_cb: Optional[ProgressCallback] = None) -> AnalysisResult:
        progress = progress_cb or _noop_progress
        if not options.any_selected():
            raise ValueError("No analysis options were selected.")

        progress(0.05, "Opening workbook…")
        profile = load_workbook_profile(workbook_path)

        # Custom mode: target the chosen sheet and remember the user's measure
        # picks so the KPI/Dashboard/Summary sheets use them too.
        custom = options.custom if (options.custom and options.custom.is_valid()) else None
        if custom is not None:
            if custom.sheet_name:
                for i, t in enumerate(profile.tables):
                    if t.sheet_name == custom.sheet_name:
                        profile.primary_table_index = i
                        break
            profile.preferred_measure_names = [m.name for m in custom.measures]
            profile.preferred_value_name = next(
                (m.name for m in custom.measures
                 if (profile.primary.column(m.name) and
                     profile.primary.column(m.name).is_value)), None)

        progress(0.20, "Detecting tables and column types…")
        use_com = excel_available()
        analyzers = self._selected_analyzers(options, use_com)

        specs = []
        summary_analyzer: Optional[ExecutiveSummaryAnalyzer] = None
        total = len(analyzers)
        for i, analyzer in enumerate(analyzers):
            frac = 0.20 + 0.50 * (i / max(1, total))
            progress(frac, f"Building {analyzer.sheet_name}…")
            # When Excel is available, real PivotTables are built by the COM
            # finalizer, so skip the static openpyxl pivot sheet.
            if use_com and isinstance(analyzer, PivotAnalyzer):
                continue
            if not analyzer.applies_to(profile):
                continue
            spec = analyzer.run(profile)
            if spec is not None:
                specs.append(spec)
            if isinstance(analyzer, ExecutiveSummaryAnalyzer):
                summary_analyzer = analyzer

        if not specs and not (use_com and options.pivot):
            raise ValueError(
                "None of the selected analyses could be produced from this "
                "workbook's data."
            )

        progress(0.72, "Writing sheets into workbook…")
        created = write_results(workbook_path, specs)

        result = AnalysisResult(
            output_path=workbook_path,
            sheets_created=created,
            warnings=list(profile.warnings),
        )

        com_pivots = 0
        if use_com:
            progress(0.86, "Building active pivot tables in Excel…")
            # Build the pivot plan, keeping only pivots whose target sheet exists.
            plan = build_pivot_plan(profile, custom, options.add_dollar)
            plan = [p for p in plan
                    if (p.target_sheet == SHEET_PIVOT and options.pivot)
                    or (p.target_sheet == SHEET_KPI and options.kpi)]
            finalizer = ExcelFinalizer()
            try:
                finalizer.finalize(workbook_path, profile, plan)
                com_pivots = finalizer.pivots_built
                if options.pivot and com_pivots > 0 and SHEET_PIVOT not in created:
                    created.append(SHEET_PIVOT)
                result.notes.extend(finalizer.notes)
            except Exception as exc:                       # noqa: BLE001
                com_pivots = getattr(finalizer, "pivots_built", 0)
                result.notes.append(
                    f"Active PivotTables could not be built in Excel ({exc}).")

        # Fallback: if Excel is absent OR built no pivots, write a STATIC Pivot
        # Analysis sheet with openpyxl so the user always gets pivot tables.
        if options.pivot and SHEET_PIVOT not in created:
            spec = PivotAnalyzer().run(profile)
            if spec is not None:
                created.extend(append_sheets(workbook_path, [spec]))
                result.notes.append(
                    "Pivot Analysis was produced as static tables"
                    + (" (Excel could not build active pivots on this workbook)."
                       if use_com else
                       " (install Microsoft Excel for active PivotTables)."))

        if summary_analyzer is not None:
            result.summary_used_llm = summary_analyzer.used_llm
            if summary_analyzer.note:
                result.notes.append(summary_analyzer.note)

        progress(1.0, "Done.")
        return result

    # -- analyzer wiring ------------------------------------------------------
    def _all_analyzers(self) -> list[Analyzer]:
        return [
            DashboardAnalyzer(),
            PivotAnalyzer(),
            KpiAnalyzer(),
            ExecutiveSummaryAnalyzer(self._narrator),
            SmartTablesAnalyzer(),
        ]

    def _selected_analyzers(self, options: AnalysisOptions,
                            use_com: bool = False) -> list[Analyzer]:
        chosen: list[Analyzer] = []
        if options.kpi:
            # With Excel, KPI stats/date tables become real pivots (built by the
            # COM finalizer), so the KPI sheet only needs its headline tiles.
            chosen.append(KpiAnalyzer(include_tables=not use_com))
        if options.pivot:
            chosen.append(PivotAnalyzer())
        if options.dashboard:
            chosen.append(DashboardAnalyzer())
        if options.executive_summary:
            chosen.append(ExecutiveSummaryAnalyzer(self._narrator))
        if options.smart_tables:
            chosen.append(SmartTablesAnalyzer())
        return chosen
