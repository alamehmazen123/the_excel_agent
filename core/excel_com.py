"""Excel COM finalizer (Windows + Excel only).

After the openpyxl writer has produced the analysis sheets, this module drives a
real Excel instance to add the Excel-native features openpyxl cannot:

  * convert the source data to a real Excel Table (ListObject) if it isn't one
  * build genuine, active PivotTables (with dates grouped by Month + Year)
  * conditional-format the top value of each numeric column (excluding totals)
  * AutoFit every column on every sheet
  * set PivotCaches to auto-refresh when the file is opened
  * save back to the original file, in place

If Excel is not installed, ``excel_available()`` returns False and the caller
falls back to the static openpyxl output.
"""
from __future__ import annotations

from typing import Optional

from .constants import (SHEET_DASHBOARD, SHEET_KPI, SHEET_PIVOT, SHEET_SUMMARY)
from .models import ColumnType, WorkbookProfile

# --- Excel constants (hard-coded so we don't depend on the typelib cache) -----
XL_SRC_RANGE = 1
XL_YES = 1
XL_DATABASE = 1
XL_ROW_FIELD = 1
XL_SUM = -4157
XL_TOP10_TOP = 1
# Group periods order: [Seconds, Minutes, Hours, Days, Months, Quarters, Years]
GROUP_MONTH_YEAR = (False, False, False, False, True, False, True)

# Highlight colours (BGR ints for COM .Color)
_HL_FILL = 198 + 239 * 256 + 206 * 65536     # light green
_HL_FONT = 0 + 97 * 256 + 0 * 65536          # dark green


def excel_available() -> bool:
    """True only if win32com is importable AND Excel COM is registered."""
    try:
        import winreg  # noqa: PLC0415
        import win32com.client  # noqa: F401,PLC0415
    except Exception:
        return False
    try:
        import winreg  # noqa: PLC0415
        winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "Excel.Application")
        return True
    except Exception:
        return False


class ExcelFinalizer:
    """Apply Excel-native finishing to a workbook that already has the sheets."""

    def __init__(self) -> None:
        self.notes: list[str] = []

    # -- public ---------------------------------------------------------------
    def finalize(self, path: str, profile: WorkbookProfile, build_pivots: bool) -> None:
        import os  # noqa: PLC0415
        import pythoncom  # noqa: PLC0415
        import win32com.client as win32  # noqa: PLC0415

        path = os.path.abspath(path)   # Excel COM requires an absolute path
        pythoncom.CoInitialize()
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        wb = None
        try:
            skip_sheets = set(profile.pivot_sheets)   # leave existing pivots alone
            wb = excel.Workbooks.Open(path)
            source_table = self._ensure_source_table(wb, profile)
            source_table_name = source_table.Name if source_table is not None else ""
            if build_pivots and source_table is not None:
                self._build_pivots(excel, wb, profile, source_table)
            self._conditional_format_tables(excel, wb, skip={source_table_name},
                                            skip_sheets=skip_sheets)
            self._conditional_format_pivots(excel, wb, skip_sheets=skip_sheets)
            self._autofit_all(wb, skip_sheets=skip_sheets)
            self._set_auto_refresh(wb)
            wb.Save()                      # in place, same path / name
        finally:
            if wb is not None:
                wb.Close(SaveChanges=False)
            excel.Quit()
            pythoncom.CoUninitialize()

    # -- req 5: source -> Table ----------------------------------------------
    def _ensure_source_table(self, wb, profile: WorkbookProfile):
        primary = profile.primary
        if primary is None:
            return None
        try:
            ws = wb.Worksheets(primary.sheet_name)
        except Exception:
            return None
        # Detect: is the data already inside a Table?
        if ws.ListObjects.Count > 0:
            self.notes.append("Source data was already an Excel Table.")
            return ws.ListObjects(1)
        lo = ws.ListObjects.Add(XL_SRC_RANGE, ws.UsedRange, None, XL_YES)
        lo.Name = self._safe_name("SourceData", wb)
        lo.TableStyle = "TableStyleMedium2"
        self.notes.append("Converted source data into an Excel Table.")
        return lo

    # -- req 1 & 2: real pivots, dates grouped by month + year ----------------
    def _build_pivots(self, excel, wb, profile: WorkbookProfile, source_table) -> None:
        primary = profile.primary
        measures = primary.measures
        dims = primary.dimensions
        if not measures or not dims:
            self.notes.append("Not enough columns to build pivot tables.")
            return
        measure = measures[0].name

        # Fresh Pivot Analysis sheet.
        pvs = self._replace_sheet(wb, SHEET_PIVOT)
        pvs.Cells(1, 1).Value = "Pivot Analysis"
        pvs.Cells(1, 1).Font.Size = 16
        pvs.Cells(1, 1).Font.Bold = True

        row = 3
        for i, dim in enumerate(dims):
            name = f"PT_{i}_{self._token(dim.name)}"
            pc = wb.PivotCaches().Create(SourceType=XL_DATABASE,
                                         SourceData=source_table.Name)
            pt = pc.CreatePivotTable(TableDestination=pvs.Cells(row, 1),
                                     TableName=name)
            pt.PivotFields(dim.name).Orientation = XL_ROW_FIELD
            pt.AddDataField(pt.PivotFields(measure), f"Sum of {measure}", XL_SUM)

            # Group date dimensions by Month + Year.
            if dim.ctype == ColumnType.DATE:
                try:
                    cell = pt.PivotFields(dim.name).DataRange.Cells(1, 1)
                    cell.Group(Start=True, End=True, Periods=GROUP_MONTH_YEAR)
                except Exception as exc:           # noqa: BLE001
                    self.notes.append(f"Could not group dates for {dim.name}: {exc}")

            pt.PivotCache().RefreshOnFileOpen = True
            try:
                pt.ColumnGrand = True
                pt.RowGrand = True
            except Exception:
                pass
            # Advance the cursor below this pivot (+ a gap).
            used = pt.TableRange2
            row = used.Row + used.Rows.Count + 2

    # -- req 3: top-1 per numeric column on every Table -----------------------
    def _conditional_format_tables(self, excel, wb, skip: set, skip_sheets: set) -> None:
        for ws in wb.Worksheets:
            if ws.Name in skip_sheets:     # sheet already has a pivot -> leave it
                continue
            for lo in ws.ListObjects:
                if lo.Name in skip:        # don't highlight the raw source data
                    continue
                for col in lo.ListColumns:
                    body = col.DataBodyRange          # excludes header + totals row
                    if body is None:
                        continue
                    if self._is_numeric_range(excel, body):
                        self._apply_top1(body)

    def _conditional_format_pivots(self, excel, wb, skip_sheets: set) -> None:
        for ws in wb.Worksheets:
            if ws.Name in skip_sheets:
                continue
            for pt in ws.PivotTables():
                body = pt.DataBodyRange
                if body is None:
                    continue
                # A grand-total ROW exists only with row fields; a grand-total
                # COLUMN only with column fields. Subtract those so the max
                # highlight excludes grand totals.
                has_row = self._count(pt, "RowFields") > 0
                has_col = self._count(pt, "ColumnFields") > 0
                nrows = body.Rows.Count - (1 if self._truthy(pt, "RowGrand") and has_row else 0)
                ncols = body.Columns.Count - (1 if self._truthy(pt, "ColumnGrand") and has_col else 0)
                if nrows < 1 or ncols < 1:
                    continue
                for j in range(1, ncols + 1):
                    colrng = ws.Range(body.Cells(1, j), body.Cells(nrows, j))
                    self._apply_top1(colrng)

    def _apply_top1(self, rng) -> None:
        try:
            fc = rng.FormatConditions.AddTop10()
            fc.TopBottom = XL_TOP10_TOP
            fc.Rank = 1
            fc.Percent = False
            fc.Interior.Color = _HL_FILL
            fc.Font.Color = _HL_FONT
            fc.Font.Bold = True
        except Exception:
            pass

    # -- req 4: autofit, auto-refresh -----------------------------------------
    def _autofit_all(self, wb, skip_sheets: set) -> None:
        for ws in wb.Worksheets:
            if ws.Name in skip_sheets:
                continue
            try:
                ws.Columns.AutoFit()
            except Exception:
                pass

    def _set_auto_refresh(self, wb) -> None:
        for ws in wb.Worksheets:
            for pt in ws.PivotTables():
                try:
                    pt.PivotCache().RefreshOnFileOpen = True
                except Exception:
                    pass

    # -- helpers --------------------------------------------------------------
    def _is_numeric_range(self, excel, rng) -> bool:
        try:
            count_num = excel.WorksheetFunction.Count(rng)
            count_all = rng.Cells.Count
            return count_all > 0 and count_num >= count_all * 0.6
        except Exception:
            return False

    def _replace_sheet(self, wb, name: str):
        """Delete an existing sheet of this name and add a fresh one at the end."""
        for ws in list(wb.Worksheets):
            if ws.Name == name:
                ws.Delete()
                break
        ws = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = name
        return ws

    @staticmethod
    def _truthy(obj, attr) -> bool:
        try:
            return bool(getattr(obj, attr))
        except Exception:
            return False

    @staticmethod
    def _count(pt, attr) -> int:
        try:
            return int(getattr(pt, attr).Count)
        except Exception:
            return 0

    @staticmethod
    def _token(s: str) -> str:
        return "".join(ch for ch in str(s) if ch.isalnum())[:20] or "Dim"

    @staticmethod
    def _safe_name(base: str, wb) -> str:
        existing = set()
        for ws in wb.Worksheets:
            for lo in ws.ListObjects:
                existing.add(lo.Name)
        name, i = base, 2
        while name in existing:
            name = f"{base}{i}"
            i += 1
        return name
