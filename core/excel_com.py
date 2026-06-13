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
XL_COLUMN_FIELD = 2
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
        self.pivots_built = 0       # how many pivots were actually created

    # -- public ---------------------------------------------------------------
    def finalize(self, path: str, profile: WorkbookProfile,
                 pivot_plan: list[PivotSpec]) -> None:
        import os  # noqa: PLC0415
        import pythoncom  # noqa: PLC0415
        import win32com.client as win32  # noqa: PLC0415

        path = os.path.abspath(path)        # Excel COM requires an absolute path
        self.pivots_built = 0
        pythoncom.CoInitialize()
        # DispatchEx spins up a SEPARATE, dedicated Excel process. Plain Dispatch
        # would attach to the user's already-open Excel — then Visible=False hides
        # their window and any modal dialog they have open (or that we trigger)
        # blocks us forever, which looks like a frozen progress bar. An isolated
        # instance can't be disturbed by, or disturb, the user's interactive Excel.
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        # Never block on link-update / add-in / macro prompts during automation.
        try:
            excel.AskToUpdateLinks = False
            excel.EnableEvents = False
            excel.AutomationSecurity = 3        # msoAutomationSecurityForceDisable
            excel.Application.DisplayAlerts = False
            # CRITICAL for big books: manual calc + no screen updates so a 80k-row
            # sheet doesn't recalculate after every pivot field is added (that was
            # the "stuck for an hour" hang). Restored implicitly when Excel quits.
            excel.Calculation = -4135           # xlCalculationManual
        except Exception:
            pass
        wb = None
        saved = False
        try:
            skip_sheets = set(profile.pivot_sheets)
            # UpdateLinks=0 → don't prompt/stall on external-reference workbooks.
            wb = excel.Workbooks.Open(path, UpdateLinks=0, IgnoreReadOnlyRecommended=True)
            data_sheet = profile.primary.sheet_name if profile.primary else None

            source_table = None
            try:
                source_table = self._ensure_source_table(wb, profile)
            except Exception as exc:                       # noqa: BLE001
                self.notes.append(f"Could not prepare the source table: {exc}")
            src_name = source_table.Name if source_table is not None else ""

            # Each step is isolated so one failure can't lose the rest, and we
            # ALWAYS try to save whatever was built (resilient on odd workbooks).
            if pivot_plan and source_table is not None:
                try:
                    self._build_plan(excel, wb, pivot_plan, source_table, data_sheet)
                except Exception as exc:                   # noqa: BLE001
                    self.notes.append(f"Some pivot tables could not be built: {exc}")
                # NOTE: the Dashboard charts are built reliably by the openpyxl
                # DashboardAnalyzer and survive the COM re-save. We deliberately do
                # NOT rebuild them from pivots here — that path used to DELETE the
                # openpyxl charts and, if the pivots were empty, left the Dashboard
                # blank. Keeping the openpyxl charts guarantees a populated sheet.
            try:
                self._conditional_format_tables(excel, wb, skip={src_name},
                                                skip_sheets=skip_sheets)
            except Exception:
                pass
            try:
                self._autofit_all(wb, skip_sheets=skip_sheets, data_sheet=data_sheet)
            except Exception:
                pass
            try:
                self._hide_helper_columns(wb, data_sheet)   # keep helpers hidden
            except Exception:
                pass
            try:
                self._order_sheets(wb)        # canonical tab order before saving
            except Exception:
                pass
            try:
                wb.Save()
                saved = True
            except Exception as exc:                       # noqa: BLE001
                self.notes.append(f"Could not save Excel changes: {exc}")
        finally:
            if wb is not None:
                wb.Close(SaveChanges=False)
            excel.Quit()
            pythoncom.CoUninitialize()
        if not saved:
            # Nothing persisted -> let the caller fall back to static tables.
            raise RuntimeError("Excel step did not save any changes.")

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

        # ONE shared PivotCache for EVERY pivot. Previously each pivot created its
        # own cache — on an 80k-row book that meant ~20 full copies of the data,
        # a ~45 MB file and a very slow run. A single shared cache is fast and
        # small. (Date grouping is a cache-level op, applied once and inherited.)
        shared_cache = wb.PivotCaches().Create(SourceType=XL_DATABASE,
                                               SourceData=source_table.Name)
        self._date_grouped = False

        built: list = []          # (pivot, spec) -> sorted/CF'd in a second pass
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
                try:
                    pt, cursor = self._build_one_pivot(excel, wb, ws, spec,
                                                       shared_cache, cursor)
                    if pt is not None:
                        built.append((pt, spec))
                        self.pivots_built += 1
                except Exception as exc:                   # noqa: BLE001
                    # Skip just this pivot; keep building the others.
                    self.notes.append(f"Skipped pivot '{spec.title}': {exc}")
                    cursor += 18

        # SECOND PASS: apply sorting + conditional formatting only after every
        # pivot exists. Date grouping refreshes the shared cache and would wipe
        # these if done inline, so they must come last.
        # Split into sub-passes: subtotal removal can refresh the shared cache,
        # so do ALL of those first, then sort, then conditional-format -- this
        # way nothing later wipes an earlier pivot's sort/CF.
        for pt, _spec in built:
            try:
                self._disable_subtotals(pt)
            except Exception:
                continue
        for pt, spec in built:
            try:
                last = spec.row_fields[-1] if spec.row_fields else None
                if last and last != spec.group_date_field:
                    self._sort_and_limit(pt, last, spec.visible_items)
            except Exception:
                continue
        for pt, _spec in built:
            try:
                self._cf_pivot(pt)
            except Exception:
                continue

    def _disable_subtotals(self, pt) -> None:
        """Turn off subtotals so each data field's DataRange is pure detail
        cells (no subtotal/grand-total rows) -- this is what 'exclude subtotals'
        means and keeps the Top-1 conditional format honest."""
        try:
            for pf in pt.RowFields:
                try:
                    pf.Subtotals = tuple([False] * 12)
                except Exception:
                    continue
        except Exception:
            pass

    def _build_one_pivot(self, excel, wb, ws, spec: PivotSpec, shared_cache,
                         cursor: int):
        # Title above the pivot.
        tcell = ws.Cells(cursor, 1)
        tcell.Value = spec.title
        tcell.Font.Bold = True
        tcell.Font.Size = 12
        tcell.Font.Color = _NAVY_BGR
        dest_row = cursor + 1

        self._pt_counter += 1
        name = f"PT_{self._pt_counter}"
        pt = shared_cache.CreatePivotTable(TableDestination=ws.Cells(dest_row, 1),
                                           TableName=name)
        # Defer recomputation until ALL fields are placed — otherwise Excel
        # re-aggregates the whole 80k-row source after every single field, which
        # is the difference between seconds and an hour on a big book.
        try:
            pt.ManualUpdate = True
        except Exception:
            pass

        # Date goes in the COLUMN panel (months across, beside the row titles)
        # WHENEVER there is a row dimension; a date-only pivot keeps it in ROWS.
        date_field = spec.group_date_field
        other_rows = [r for r in spec.row_fields if r != date_field]
        date_in_columns = bool(date_field and other_rows)

        # Place the date in ROWS FIRST. Grouping needs a rendered range, and an
        # UNgrouped date in the COLUMN panel would briefly explode into hundreds
        # of day-columns; in rows it's cheap. We group it, then (if needed) move
        # the small grouped band to the columns. Dimensions go straight to rows.
        if date_field and date_field in spec.row_fields:
            try:
                pt.PivotFields(date_field).Orientation = XL_ROW_FIELD
            except Exception:
                pass
        for rf in other_rows:
            try:
                pt.PivotFields(rf).Orientation = XL_ROW_FIELD
            except Exception:
                pass

        for df in spec.data_fields:
            # A USD-converted measure is a pivot CalculatedField (= base / rate).
            # Create it on demand; if it fails, skip just this dollar column.
            if getattr(df, "calc_formula", None):
                try:
                    pt.CalculatedFields().Add(df.source_field, df.calc_formula)
                except Exception:
                    continue
            try:
                fld = pt.AddDataField(pt.PivotFields(df.source_field),
                                      df.caption, df.func)
            except Exception:
                continue
            if getattr(df, "calculation", None) is not None:
                try:
                    fld.Calculation = df.calculation     # e.g. % of grand total
                except Exception:
                    pass
            try:
                fld.NumberFormat = df.number_format
            except Exception:
                pass

        try:
            pt.RowGrand = True
            pt.ColumnGrand = True
        except Exception:
            pass

        # Render the pivot ONCE now that every field is placed.
        try:
            pt.ManualUpdate = False
        except Exception:
            pass

        # Group the date by Month + Year (needs the rendered range). On the shared
        # cache this happens effectively once; later pivots inherit the grouping.
        if date_field and date_field in spec.row_fields:
            try:
                pt.PivotFields(date_field).DataRange.Cells(1, 1).Group(
                    Start=True, End=True, Periods=GROUP_MONTH_YEAR)
                self._date_grouped = True
            except Exception:
                pass
            # Move the now-small grouped band (Year → Month) to the COLUMN panel
            # when a row dimension exists; otherwise keep it in ROWS.
            target = XL_COLUMN_FIELD if date_in_columns else XL_ROW_FIELD
            try:
                pt.PivotFields(date_field).Orientation = target
            except Exception:
                pass
            try:
                yf = pt.PivotFields("Years")
                yf.Orientation = target
                yf.Position = 1
            except Exception:
                pass

        # NOTE: sorting + conditional formatting are applied in a SECOND pass
        # (see _build_plan) once every pivot exists, and we deliberately do NOT
        # enable RefreshOnFileOpen -- both a later date-group on the shared cache
        # and an on-open refresh would otherwise wipe the sort order and CF.

        used = pt.TableRange2
        return pt, used.Row + used.Rows.Count + 2

    def _sort_and_limit(self, pt, field_name: str,
                        visible_items: Optional[list]) -> None:
        """Sort a row field by its value field (descending) and, for wide fields,
        show only the Top-N items. AutoSort is a field-definition property, so it
        survives the cache refreshes that reset manual PivotItem positions."""
        try:
            pf = pt.PivotFields(field_name)
        except Exception:
            return

        # Hide non-top items (PivotItems must be CALLED). Each toggle is a COM
        # round-trip, so on a field with thousands of distinct values this would
        # crawl — skip the per-item hide above a sane cap and rely on AutoSort.
        if visible_items:
            allow = {str(v) for v in visible_items}
            try:
                items = pf.PivotItems()
                if items.Count <= 400:
                    for i in range(1, items.Count + 1):
                        it = items.Item(i)
                        want = str(it.Name) in allow
                        if bool(it.Visible) != want:
                            it.Visible = want
            except Exception:
                pass

        # AutoSort the field descending by the (first) data field.
        try:
            cap = pt.DataFields.Item(1).Name
            pf.AutoSort(2, cap)            # xlDescending
        except Exception:
            pass

    # -- Dashboard: one chart per category pivot ------------------------------
    def _build_dashboard_charts(self, wb, max_charts: int = 6) -> None:
        """Add a column chart to the Dashboard for each single-dimension value
        pivot on the Pivot Analysis sheet, so the dashboard reflects the pivots."""
        from .constants import SHEET_DASHBOARD  # noqa: PLC0415
        try:
            dash = wb.Worksheets(SHEET_DASHBOARD)
            pv = wb.Worksheets(SHEET_PIVOT)
        except Exception:
            return
        # Clear the old single static chart so we can lay out fresh ones.
        try:
            for ch in list(dash.ChartObjects()):
                ch.Delete()
        except Exception:
            pass

        used = dash.UsedRange
        top0 = used.Top + used.Height + 12
        left0 = dash.Cells(1, 1).Left + 6
        W, H, GAP = 360, 220, 16
        made = 0
        XL_COL_CLUSTERED = 51
        for pt in pv.PivotTables():
            if made >= max_charts:
                break
            try:
                rfields = list(pt.RowFields)
                if len(rfields) != 1:                 # single dimension only
                    continue
                dim_name = rfields[0].Name
                df = pt.DataFields.Item(1)
                if "%" in (df.NumberFormat or ""):    # chart the $ / value pivots
                    continue
                if pt.PivotFields(dim_name).PivotItems().Count > 25:   # skip wide
                    continue
                col = made % 2
                rowi = made // 2
                left = left0 + col * (W + GAP)
                top = top0 + rowi * (H + GAP)
                shp = dash.Shapes.AddChart2(-1, XL_COL_CLUSTERED, left, top, W, H)
                chart = shp.Chart
                chart.SetSourceData(pt.TableRange1)
                try:
                    chart.HasLegend = False
                except Exception:
                    pass
                try:
                    chart.HasTitle = True
                    cap = pv.Cells(pt.TableRange2.Row - 1, 1).Value
                    chart.ChartTitle.Text = str(cap) if cap else f"{df.Name} by {dim_name}"
                except Exception:
                    pass
                made += 1
            except Exception:
                continue

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
                # A data field's DataRange is the DETAIL value cells only (Excel
                # excludes grand totals); with subtotals disabled it has no
                # subtotal rows either, so a plain Top-1 rule highlights exactly
                # the single largest value of that column.
                rng = fld.DataRange
                rng.FormatConditions.Delete()
                fc = rng.FormatConditions.AddTop10()
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

    # -- keep helper columns hidden -------------------------------------------
    def _hide_helper_columns(self, wb, data_sheet) -> None:
        """Re-hide the engine's helper columns on the data sheet — decoded-name
        ``"… (Name)"`` and sign-flipped ``"… (+)"`` columns. openpyxl marks them
        hidden, but re-saving through Excel/COM can surface them; this guarantees
        the user never sees the internal helpers."""
        if not data_sheet:
            return
        try:
            ws = wb.Worksheets(data_sheet)
        except Exception:
            return
        used = ws.UsedRange
        first_row = used.Row
        ncols = used.Columns.Count
        base = used.Column
        for i in range(ncols):
            c = base + i
            try:
                header = ws.Cells(first_row, c).Value
            except Exception:
                continue
            name = str(header) if header is not None else ""
            if name.endswith(" (Name)") or name.endswith(" (+)"):
                try:
                    ws.Columns(c).Hidden = True
                except Exception:
                    pass

    # -- canonical tab order --------------------------------------------------
    def _order_sheets(self, wb) -> None:
        """Sequence the analysis tabs AFTER the data sheet(s):
        KPI → Pivot → Smart Tables → Insights → Executive Summary → Dashboard.
        Moving each present sheet to the end in this order leaves them in
        sequence at the tail, with the data sheet(s) untouched at the front."""
        from .constants import ordered_output_layout  # noqa: PLC0415
        for base in ordered_output_layout():
            try:
                ws = wb.Worksheets(base)
            except Exception:
                continue            # this sheet wasn't produced this run
            try:
                ws.Move(After=wb.Worksheets(wb.Worksheets.Count))
            except Exception:
                pass

    # -- req 9: autofit -------------------------------------------------------
    def _autofit_all(self, wb, skip_sheets: set, data_sheet=None) -> None:
        """AutoFit ONLY the (small) output sheets. AutoFit on the raw data sheet
        of an 80k-row book scans millions of cells and can take an hour — that
        was a primary cause of the 'stuck on Auto-Generate' hang. The data sheet
        and any large sheet are skipped."""
        for ws in wb.Worksheets:
            if ws.Name in skip_sheets or ws.Name == data_sheet:
                continue
            try:
                if ws.UsedRange.Rows.Count > 3000:
                    continue            # large data sheet -> AutoFit is too slow
            except Exception:
                pass
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
