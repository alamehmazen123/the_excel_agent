"""Tests for the reference library and the Smart Tables analyzer.

These run without Excel/COM -- the library is pure Python + JSON and the
analyzer emits a SheetSpec, so they are fast and deterministic.
"""
from __future__ import annotations

import datetime as _dt

from core.analyzers.smart_tables import SmartTablesAnalyzer
from core.decode import apply_to_profile, find_decodable
from core.library.store import CodeMap, HeaderEntry, Library
from core.models import (ColumnProfile, ColumnType, TableProfile,
                         WorkbookProfile)


def _library() -> Library:
    dept = CodeMap(name="department", label="Department", entries={
        "001": "Cardiology", "002": "Radiology", "003": "Emergency",
    }).reindex()
    lib = Library(
        headers={"DEPT": HeaderEntry("DEPT", "Department", category="department")},
        code_maps={"department": dept},
    )
    return lib


def test_code_lookup_is_tolerant() -> None:
    cm = _library().code_maps["department"]
    assert cm.lookup("001") == "Cardiology"
    assert cm.lookup(1) == "Cardiology"        # Excel-read integer
    assert cm.lookup(" 002 ") == "Radiology"   # whitespace
    assert cm.lookup("999") is None


def test_coverage_and_best_map() -> None:
    lib = _library()
    cov = lib.code_maps["department"].coverage([1, 2, 3, 999])
    assert abs(cov - 0.75) < 1e-9
    assert lib.best_map_for_values([1, 2, 3]).name == "department"
    assert lib.best_map_for_values(["zzz", "qqq"]) is None


def _profile_with_codes() -> WorkbookProfile:
    col_dept = ColumnProfile(name="DEPT", index=0, ctype=ColumnType.CATEGORICAL,
                             count=4, distinct=3)
    col_dept.top_values = [(1, 2), (2, 1), (3, 1)]
    col_val = ColumnProfile(name="Revenue", index=1, ctype=ColumnType.CURRENCY,
                            count=4, total=400.0)
    # A date column is now required (BASIC RULE: every smart table is grouped by
    # month) — two months so the grouping is exercised.
    col_date = ColumnProfile(name="Date", index=2, ctype=ColumnType.DATE, count=4)
    jan, feb = _dt.datetime(2026, 1, 5), _dt.datetime(2026, 2, 9)
    table = TableProfile(
        sheet_name="Data", header_row=1, first_data_row=2, last_data_row=5,
        first_col=1, last_col=3, columns=[col_dept, col_val, col_date],
        rows=[
            {"DEPT": 1, "Revenue": 100.0, "Date": jan},
            {"DEPT": 1, "Revenue": 150.0, "Date": feb},
            {"DEPT": 2, "Revenue": 90.0, "Date": jan},
            {"DEPT": 3, "Revenue": 60.0, "Date": feb},
        ],
    )
    return WorkbookProfile(path="x.xlsx", sheet_names=["Data"], tables=[table])


def test_smart_tables_decodes_and_groups() -> None:
    lib = _library()
    profile = _profile_with_codes()
    # The pipeline injects decoded helper columns before analyzers run.
    table = profile.primary
    apply_to_profile(table, find_decodable(table, lib), lib)

    analyzer = SmartTablesAnalyzer(library=lib)
    assert analyzer.applies_to(profile)            # a helper now exists
    spec = analyzer.run(profile)
    assert spec is not None and spec.tables

    # A months-across cross-tab for the decoded DEPT names should exist; its title
    # uses the FRIENDLY header ("Department"), not the raw helper "DEPT (Name)".
    decoded = [t for t in spec.tables if "Department" in t.title]
    assert decoded, [t.title for t in spec.tables]
    labels = [r[0] for r in decoded[0].rows]       # col 0 = decoded dimension
    assert "Cardiology" in labels and "001" not in labels
    # Months run ACROSS the columns now (headers like 'Jan-26'), plus a Total.
    assert any("-" in str(h) for h in decoded[0].headers[1:])
    assert decoded[0].headers[-1] == "Total"


def test_empty_library_still_applies_with_raw_names() -> None:
    # Smart Tables now apply to any value+date+dimension workbook; without a
    # library the dimensions just stay as their raw (undecoded) names. The
    # checkbox must NOT silently un-tick on unrecognised files.
    analyzer = SmartTablesAnalyzer(library=Library())
    profile = _profile_with_codes()                 # has Date + DEPT + Revenue
    assert analyzer.applies_to(profile)
    spec = analyzer.run(profile)
    assert spec is not None and spec.tables


def test_currency_defaults_to_lbp() -> None:
    from core.formatting import fmt_measure
    lbp = ColumnProfile(name="ORG_AMOUNT", index=0, ctype=ColumnType.CURRENCY)
    usd = ColumnProfile(name="USD", index=1, ctype=ColumnType.CURRENCY)
    assert fmt_measure(lbp, 1_500_000).endswith(" LBP")
    assert fmt_measure(usd, 1_500_000).startswith("$")


def test_every_pivot_is_date_grouped() -> None:
    from core.pivot_plan import build_pivot_plan
    from core.constants import SHEET_PIVOT
    profile = _profile_with_codes()           # has Date + DEPT + Revenue
    plan = build_pivot_plan(profile)
    cat_pivots = [p for p in plan if p.target_sheet == SHEET_PIVOT]
    assert cat_pivots
    # BASIC RULE: when a date column exists, no category pivot is dateless.
    assert all(p.group_date_field for p in cat_pivots), \
        [p.title for p in cat_pivots if not p.group_date_field]
