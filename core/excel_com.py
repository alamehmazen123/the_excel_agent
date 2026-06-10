"""Excel COM finalizer (Windows + Excel only).

After the openpyxl writer creates the KPI/Dashboard/Executive sheets, this module
drives a real Excel instance to add everything openpyxl cannot:

  * convert the source data to a real Excel Table (if not already one)         [req 5]
  * build genuine PivotTables from a declarative plan, with dates grouped by
    Month + Year, number formats inherited from the source, and the Pivot
    Analysis sheet placed right after the data sheet                      [req 1,2,3,4]
  * top-1 conditional formatting per value field, scoped so it automatically
    excludes grand totals AND subtotals                                     [req 6,7]
  * AutoFit every column on every produced sheet                              [req 9]
  * set PivotCaches to auto-refresh on open; save back in place

Falls back (caller's job) to static openpyxl tables when Excel is absent.
"""
from __future__ import annotations

from collections import defaultdict
from typing import Optional

from .constants import SHEET_PIVOT
from .models import ColumnType, WorkbookProfile
from .pivot_plan import PivotSpec

# --- Excel constants (hard-coded; no dependence on the typelib cache) ---------
XL_SRC_RANGE = 1
XL_YES = 1
XL_DATABASE = 1
XL_ROW_FIELD = 1
XL_TOP10_TOP = 1
XL_DATA_FIELD_SCOPE = 2          # XlPivotConditionScope.xlDataFieldScope
# Group periods order: [Seconds, Minutes, Hours, Days, Months, Quarters, Years]
GROUP_MONTH_YEAR = (False, False, False, False, True, False, True)

_NAVY = 0x1F3864 & 0xFFFFFF
_NAVY_BGR = 0x64381F            # navy as BGR for COM .Color
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
    def __init__(self) -> None:
        self.notes: list[str] = []
        self._pt_counter = 0

    # -- public ---------------------------------------------------------------
    def finalize(self, path: str, profile: WorkbookProfile,
                 pivot_plan: list[PivotSpec]) -> None:
        import os  # noqa: PLC0415
        import pythoncom  # noqa: PLC0415
        import win32com.client as win32  # noqa: PLC0415

        path = os.path.abspath(path)        # Excel COM requires an absolute path
        pythoncom.CoInitialize()
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        wb = None
        try:
            skip_sheets = set(profile.pivot_sheets)
            wb = excel.Workbooks.Open(path)
            data_sheet = profile.primary.sheet_name if profile.primary else None
            source_table = self._ensure_source_table(wb, profile)
            src_name = source_table.Name if source_table is not None else ""

            if pivot_plan and source_table is not None:
                self._build_plan(excel, wb, pivot_plan, source_table, data_sheet)

            self._conditional_format_tables(excel, wb, skip={src_name},
                                            skip_sheets=skip_sheets)
            self._autofit_all(wb, skip_sheets=skip_sheets)
            self._set_auto_refresh(wb)
            wb.Save()
        finally:
            if wb is not None:
                wb.Close(SaveChanges=False)
            excel.Quit()
            pythoncom.CoUninitialize()

    # -- source -> Table (req 5 detection) ------------------------------------
    def _ensure_source_table(self, wb, profile: WorkbookProfile):
        primary = profile.primary
        if primary is None:
            return None
        try:
            ws = wb.Worksheets(primary.sheet_name)
        except Exception:
            return None
        if ws.ListObjects.Count > 0:
            self.notes.append("Source data was already an Excel Table.")
            return ws.ListObjects(1)
        lo = ws.ListObjects.Add(XL_SRC_RANGE, ws.UsedRange, None, XL_YES)
        lo.Name = self._safe_table_name("SourceData", wb)
        lo.TableStyle = "TableStyleMedium2"
        self.notes.append("Converted source data into an Excel Table.")
        return lo

    # -- build all pivots from the plan (req 1,2,3,4) -------------------------
    def _build_plan(self, excel, wb, plan: list[PivotSpec], source_table,
                    data_sheet: Optional[str]) -> None:
        by_sheet: dict[str, list[PivotSpec]] = defaultdict(list)
        for spec in plan:
            by_sheet[spec.target_sheet].append(spec)

        for sheet_name, specs in by_sheet.items():
            if sheet_name == SHEET_PIVOT:
                ws = self._replace_sheet(wb, sheet_name)
                # Place the Pivot Analysis sheet directly after the data sheet.
                if data_sheet:
                    try:
                        ws.Move(After=wb.Worksheets(data_sheet))
                    except Exception:
                        pass
                ws.Cells(1, 1).Value = "Pivot Analysis"
                ws.Cells(1, 1).Font.Size = 16
                ws.Cells(1, 1).Font.Bold = True
                ws.Cells(1, 1).Font.Color = _NAVY_BGR
                cursor = 3
            else:
                # Existing sheet (e.g. KPI Analysis) -> append below its content.
                ws = wb.Worksheets(sheet_name)
                used = ws.UsedRange
                cursor = used.Row + used.Rows.Count + 2

            for spec in specs:
                cursor = self._build_one_pivot(excel, wb, ws, spec,
                                               source_table, cursor)

    def _build_one_pivot(self, excel, wb, ws, spec: PivotSpec, source_table,
                         cursor: int) -> int:
        # Title above the pivot.
        tcell = ws.Cells(cursor, 1)
        tcell.Value = spec.title
        tcell.Font.Bold = True
        tcell.Font.Size = 12
        tcell.Font.Color = _NAVY_BGR
        dest_row = cursor + 1

        self._pt_counter += 1
        name = f"PT_{self._pt_counter}"
        pc = wb.PivotCaches().Create(SourceType=XL_DATABASE,
                                     SourceData=source_table.Name)
        pt = pc.CreatePivotTable(TableDestination=ws.Cells(dest_row, 1),
                                 TableName=name)

        # Add row fields so the grouped DATE stays OUTERMOST: add+group the date
        # first, then the remaining dimensions become inner fields.
        remaining = list(spec.row_fields)
        date_field = spec.group_date_field
        if date_field and date_field in remaining:
            pt.PivotFields(date_field).Orientation = XL_ROW_FIELD
            try:
                cell = pt.PivotFields(date_field).DataRange.Cells(1, 1)
                cell.Group(Start=True, End=True, Periods=GROUP_MONTH_YEAR)
            except Exception as exc:                  # noqa: BLE001
                self.notes.append(f"Could not group dates for {date_field}: {exc}")
            remaining = [r for r in remaining if r != date_field]
        for rf in remaining:
            pt.PivotFields(rf).Orientation = XL_ROW_FIELD

        for df in spec.data_fields:
            fld = pt.AddDataField(pt.PivotFields(df.source_field),
                                  df.caption, df.func)
            try:
                fld.NumberFormat = df.number_format
            except Exception:
                pass

        # Sort + Top-N limiting on the last row field (e.g. wide TRIGGER DETAIL).
        if spec.row_fields:
            self._sort_and_limit(pt, spec.row_fields[-1], spec.sort_field,
                                 spec.visible_items)

        try:
            pt.RowGrand = True
            pt.ColumnGrand = True
        except Exception:
            pass
        try:
            pt.PivotCache().RefreshOnFileOpen = True
        except Exception:
            pass

        self._cf_pivot(pt)

        used = pt.TableRange2
        return used.Row + used.Rows.Count + 2

    def _sort_and_limit(self, pt, field_name: str, sort_field: Optional[str],
                        visible_items: Optional[list]) -> None:
        """Sort a row field descending and, for wide fields, show only Top-N items."""
        try:
            pf = pt.PivotFields(field_name)
        except Exception:
            return
        # Hide non-top items FIRST (PivotItems must be CALLED), then sort.
        if visible_items:
            allow = {str(v) for v in visible_items}
            try:
                items = pf.PivotItems()
                count = items.Count
            except Exception:
                count = 0
            for i in range(1, count + 1):
                try:
                    it = items.Item(i)
                    want = str(it.Name) in allow
                    if bool(it.Visible) != want:
                        it.Visible = want
                except Exception:
                    continue
        if sort_field:
            try:
                pf.AutoSort(2, sort_field)        # xlDescending
            except Exception:
                pass

    # -- conditional formatting (req 6,7): top-1 per value field --------------
    def _cf_pivot(self, pt) -> None:
        """Highlight the single largest value of each data field.

        Using a Top10 rule with ScopeType=xlDataFieldScope makes Excel apply it
        across all detail cells of that data field, automatically EXCLUDING
        subtotals and grand totals -- exactly the requirement.
        """
        try:
            fields = list(pt.DataFields)
        except Exception:
            return
        for fld in fields:
            try:
                rng = fld.DataRange
                cell = rng.Cells(1, 1)
                fc = cell.FormatConditions.AddTop10()
                try:
                    fc.ScopeType = XL_DATA_FIELD_SCOPE
                except Exception:
                    pass
                fc.TopBottom = XL_TOP10_TOP
                fc.Rank = 1
                fc.Percent = False
                fc.Interior.Color = _HL_FILL
                fc.Font.Color = _HL_FONT
                fc.Font.Bold = True
            except Exception:
                continue

    # Static Tables (ListObjects) -- DataBodyRange already excludes header/totals.
    def _conditional_format_tables(self, excel, wb, skip: set, skip_sheets: set) -> None:
        for ws in wb.Worksheets:
            if ws.Name in skip_sheets:
                continue
            for lo in ws.ListObjects:
                if lo.Name in skip:
                    continue
                for col in lo.ListColumns:
                    body = col.DataBodyRange
                    if body is None:
                        continue
                    if self._is_numeric_range(excel, body):
                        self._apply_top1(body)

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

    # -- req 9: autofit -------------------------------------------------------
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
        for ws in list(wb.Worksheets):
            if ws.Name == name:
                ws.Delete()
                break
        ws = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = name
        return ws

    @staticmethod
    def _safe_table_name(base: str, wb) -> str:
        existing = set()
        for ws in wb.Worksheets:
            for lo in ws.ListObjects:
                existing.add(lo.Name)
        name, i = base, 2
        while name in existing:
            name = f"{base}{i}"
            i += 1
        return name
