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
from .constants import SHEET_PIVOT
from .excel_com import ExcelFinalizer, excel_available
from .loader import load_workbook_profile
from .models import (AnalysisOptions, AnalysisResult, ProgressCallback,
                     WorkbookProfile, _noop_progress)
from .writer import write_results


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

    # -- main run -------------------------------------------------------------
    def run(self, workbook_path: str, options: AnalysisOptions,
            progress_cb: Optional[ProgressCallback] = None) -> AnalysisResult:
        progress = progress_cb or _noop_progress
        if not options.any_selected():
            raise ValueError("No analysis options were selected.")

        progress(0.05, "Opening workbook…")
        profile = load_workbook_profile(workbook_path)

        progress(0.20, "Detecting tables and column types…")
        use_com = excel_available()
        analyzers = self._selected_analyzers(options)

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

        if use_com:
            progress(0.86, "Building active pivot tables in Excel…")
            finalizer = ExcelFinalizer()
            try:
                finalizer.finalize(workbook_path, profile, build_pivots=options.pivot)
                if options.pivot and SHEET_PIVOT not in created:
                    created.append(SHEET_PIVOT)
                result.notes.extend(finalizer.notes)
            except Exception as exc:                       # noqa: BLE001
                result.notes.append(
                    f"Excel finishing step failed ({exc}). The analysis sheets "
                    "were still produced as static tables.")
        elif options.pivot:
            result.notes.append(
                "Microsoft Excel was not found on this PC, so pivot results are "
                "static tables. Install Excel for active, refreshable PivotTables.")

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
        ]

    def _selected_analyzers(self, options: AnalysisOptions) -> list[Analyzer]:
        chosen: list[Analyzer] = []
        if options.kpi:
            chosen.append(KpiAnalyzer())
        if options.pivot:
            chosen.append(PivotAnalyzer())
        if options.dashboard:
            chosen.append(DashboardAnalyzer())
        if options.executive_summary:
            chosen.append(ExecutiveSummaryAnalyzer(self._narrator))
        return chosen
