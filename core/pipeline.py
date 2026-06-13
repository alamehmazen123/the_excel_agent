"""The engine's single public entry point.

Every front-end (desktop now; Ribbon/web/API later) calls ``Engine.run``.
It knows nothing about how it was invoked.
"""
from __future__ import annotations

from typing import Optional

from .analyzers.base import Analyzer
from .analyzers.dashboard import DashboardAnalyzer
from .analyzers.executive_summary import ExecutiveSummaryAnalyzer, Narrator
from .analyzers.insights import InsightsAnalyzer
from .analyzers.kpi import KpiAnalyzer
from .analyzers.pivot import PivotAnalyzer
from .analyzers.smart_tables import SmartTablesAnalyzer
from .constants import SHEET_INSIGHTS, SHEET_KPI, SHEET_PIVOT
from .excel_com import ExcelFinalizer, excel_available
from .loader import load_workbook_profile
from .pivot_plan import build_pivot_plan
from .models import (AnalysisOptions, AnalysisResult, ProgressCallback,
                     TableProfile, WorkbookProfile, _noop_progress)
from .writer import append_sheets, write_results


def _apply_revenue_sign(table: TableProfile) -> list[tuple[str, str]]:
    """Flip negative LBP money columns to positive (Lebanese revenue books store
    revenue as a negative). USD/$ columns are already positive and left untouched.

    Mutates the in-memory rows + stats so the openpyxl sheets read positive, and
    assigns each column a hidden POSITIVE helper (``"<col> (+)"``) so the COM
    PivotTables aggregate a real positive column. The original sheet cells are
    NOT changed. Returns ``[(source_col, helper_col), …]`` for the writer to
    inject onto the sheet."""
    from .formatting import is_dollar_column  # noqa: PLC0415
    targets = []
    for col in table.value_measures:
        if is_dollar_column(col):
            continue
        negative = ((col.total is not None and col.total < 0)
                    or (col.mean is not None and col.mean < 0))
        if negative:
            targets.append(col)
    if not targets:
        return []
    names = [c.name for c in targets]
    for row in table.rows:
        for n in names:
            v = row.get(n)
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                row[n] = -v
    helpers: list[tuple[str, str]] = []
    for c in targets:
        c.positive_helper = f"{c.name} (+)"
        helpers.append((c.name, c.positive_helper))
        if c.total is not None:
            c.total = -c.total
        if c.mean is not None:
            c.mean = -c.mean
        lo, hi = c.minimum, c.maximum
        c.minimum = -hi if hi is not None else None
        c.maximum = -lo if lo is not None else None
    return helpers


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
        from .library import get_library  # noqa: PLC0415
        lib = get_library()

        def _describe(name: str) -> str:
            """Library meaning of a header, or '' if it adds nothing."""
            meaning = lib.meaning_of(name)
            return meaning if meaning and meaning != name else ""

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
                         "description": _describe(c.name),
                         "recommended": c.name in recommended_dims})
        # dates are also offered as dimensions (grouped) -> include any not already
        for c in table.date_columns:
            if not any(d["name"] == c.name for d in dims):
                dims.append({"name": c.name, "kind": "date",
                             "detail": "grouped by month/year",
                             "description": _describe(c.name), "recommended": True})

        rec_measures = {c.name for c in table.key_measures[:3]}
        measures = []
        for c in table.value_measures + table.percent_measures:
            unit = ("percent" if c.ctype == ColumnType.PERCENT
                    else "currency" if c.ctype == ColumnType.CURRENCY else "number")
            measures.append({"name": c.name, "unit": unit,
                             "description": _describe(c.name),
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

        # Library decoding: detect code columns the reference library can decode
        # and inject hidden decoded-name helper columns so EVERY sheet (tiles,
        # dashboard, pivots, smart tables, summary) groups by readable names.
        from .decode import find_decodable, apply_to_profile  # noqa: PLC0415
        from .library import get_library  # noqa: PLC0415
        library = get_library()
        decodes = find_decodable(profile.primary, library) if profile.primary else []
        if decodes:
            apply_to_profile(profile.primary, decodes, library)

        # Detect the workbook's PURPOSE from the account-code categories, and for
        # a revenue book (Lebanese convention stores revenue NEGATIVE) flip the
        # LBP money columns to positive so every sheet reads naturally.
        from .semantic import analyze as _analyze  # noqa: PLC0415
        semantic = _analyze(profile, library)
        value_helpers: list = []
        if semantic.is_revenue_report and profile.primary is not None:
            value_helpers = _apply_revenue_sign(profile.primary)

        # If the sheet already carries a USD / $ column, never add another dollar
        # column or run the LBP→USD calc — the dollars are already provided.
        from .formatting import is_dollar_column  # noqa: PLC0415
        has_usd = bool(profile.primary and
                       any(is_dollar_column(c) for c in profile.primary.value_measures))
        add_dollar = bool(options.add_dollar and not has_usd)
        if has_usd and options.add_dollar:
            result_note_usd = ("A USD column was detected, so no extra dollar "
                               "column was added.")
        else:
            result_note_usd = None

        use_com = excel_available()
        analyzers = self._selected_analyzers(options, use_com)

        specs = []
        summary_analyzer: Optional[ExecutiveSummaryAnalyzer] = None
        insights_analyzer: Optional[InsightsAnalyzer] = None
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
            if isinstance(analyzer, InsightsAnalyzer):
                insights_analyzer = analyzer

        if not specs and not (use_com and options.pivot):
            raise ValueError(
                "None of the selected analyses could be produced from this "
                "workbook's data."
            )

        progress(0.72, "Writing sheets into workbook…")
        inject = (profile.primary, decodes, library) if decodes else None
        vhelpers = (profile.primary, value_helpers) if value_helpers else None
        created = write_results(workbook_path, specs, inject=inject,
                                value_helpers=vhelpers)

        result = AnalysisResult(
            output_path=workbook_path,
            sheets_created=created,
            warnings=list(profile.warnings),
        )
        if insights_analyzer is not None:
            result.insights = insights_analyzer.insights
        if semantic.purpose:
            result.notes.append(
                f"Detected purpose: this looks like a {semantic.purpose.upper()} "
                f"report (from the account-code categories).")
        if result_note_usd:
            result.notes.append(result_note_usd)
        if decodes:
            helpers = ", ".join(f"'{d.helper_name}'" for d in decodes)
            result.notes.append(
                f"Added {len(decodes)} HIDDEN helper column(s) of decoded names "
                f"to '{profile.primary.sheet_name}' ({helpers}). Original columns "
                "are unchanged; unhide them in Excel to see the decoded names.")

        com_pivots = 0
        if use_com:
            progress(0.86, "Building active pivot tables in Excel…")
            # Build the pivot plan, keeping only pivots whose target sheet exists.
            plan = build_pivot_plan(profile, custom, add_dollar)
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
            InsightsAnalyzer(),
            DashboardAnalyzer(),
            PivotAnalyzer(),
            KpiAnalyzer(),
            ExecutiveSummaryAnalyzer(self._narrator),
            SmartTablesAnalyzer(),
        ]

    def _selected_analyzers(self, options: AnalysisOptions,
                            use_com: bool = False) -> list[Analyzer]:
        chosen: list[Analyzer] = []
        if options.insights:
            # First so the Insights sheet renders before everything else.
            chosen.append(InsightsAnalyzer())
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
