"""Tests for the reference library and the Smart Tables analyzer.

These run without Excel/COM -- the library is pure Python + JSON and the
analyzer emits a SheetSpec, so they are fast and deterministic.
"""
from __future__ import annotations

from core.analyzers.smart_tables import SmartTablesAnalyzer
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
    table = TableProfile(
        sheet_name="Data", header_row=1, first_data_row=2, last_data_row=5,
        first_col=1, last_col=2, columns=[col_dept, col_val],
        rows=[
            {"DEPT": 1, "Revenue": 100.0},
            {"DEPT": 1, "Revenue": 150.0},
            {"DEPT": 2, "Revenue": 90.0},
            {"DEPT": 3, "Revenue": 60.0},
        ],
    )
    return WorkbookProfile(path="x.xlsx", sheet_names=["Data"], tables=[table])


def test_smart_tables_decodes_and_groups() -> None:
    analyzer = SmartTablesAnalyzer(library=_library())
    profile = _profile_with_codes()
    assert analyzer.applies_to(profile)
    spec = analyzer.run(profile)
    assert spec is not None and spec.tables
    table = spec.tables[0]
    assert "Department" in table.title          # decoded header meaning
    labels = [r[0] for r in table.rows]
    assert "Cardiology" in labels and "001" not in labels
    # Cardiology (100+150=250) ranks first.
    assert table.rows[0][0] == "Cardiology"
    assert abs(table.rows[0][1] - 250.0) < 1e-9


def test_empty_library_does_not_apply() -> None:
    analyzer = SmartTablesAnalyzer(library=Library())
    assert not analyzer.applies_to(_profile_with_codes())
