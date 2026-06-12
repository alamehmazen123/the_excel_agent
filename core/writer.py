"""Render SheetSpecs into the workbook using openpyxl.

Contract: the original sheets are never modified -- we only append new sheets.
Output is written to a temp file and atomically swapped in to avoid leaving a
corrupted workbook if the process is interrupted.
"""
from __future__ import annotations

import os
import tempfile
from typing import Optional

from openpyxl import load_workbook
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.worksheet import Worksheet

from .constants import output_sheet_names

from .render import (ChartKind, ChartSpec, DataTable, NumberFormat, SheetSpec,
                     TextBlock)

# Palette
_NAVY = "1F3864"
_BLUE = "2E5496"
_LIGHT = "D6E0F0"
_GREEN = "2E7D32"
_RED = "C62828"
_GREY = "595959"
_WHITE = "FFFFFF"

_THIN = Side(style="thin", color="BFBFBF")
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)

_NUMFMT = {
    NumberFormat.GENERAL: "General",
    NumberFormat.INTEGER: "#,##0",
    NumberFormat.DECIMAL: "#,##0.00",
    NumberFormat.CURRENCY: "$#,##0.00",
    NumberFormat.PERCENT: "0.0%",
    NumberFormat.DATE: "yyyy-mm-dd",
}


def _remove_existing_outputs(wb) -> None:
    """Delete any analysis sheets from a previous run so we regenerate cleanly."""
    targets = set(output_sheet_names())
    for ws in list(wb.worksheets):
        name = ws.title
        # match exact names and versioned variants like "Dashboard (2)"
        base = name.split(" (")[0]
        if base in targets and len(wb.worksheets) > 1:
            del wb[name]


def _unique_sheet_name(wb, base: str) -> str:
    if base not in wb.sheetnames:
        return base
    i = 2
    while f"{base} ({i})" in wb.sheetnames:
        i += 1
    return f"{base} ({i})"


_BRAND_BLUE = "0070C0"   # vivid blue for the heading / hospital name


def _write_heading(ws: Worksheet, spec: SheetSpec, row: int) -> int:
    # Heading text rendered BOLD + BLUE on white (e.g. "SAHEL GENERAL HOSPITAL").
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
    c = ws.cell(row=row, column=1, value=spec.heading)
    c.font = Font(name="Calibri", size=20, bold=True, color=_BRAND_BLUE)
    c.alignment = Alignment(vertical="center", indent=1)
    ws.row_dimensions[row].height = 32
    row += 1
    if spec.subheading:
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
        s = ws.cell(row=row, column=1, value=spec.subheading)
        s.font = Font(size=11, bold=True, color=_NAVY)
        s.alignment = Alignment(indent=1)
        row += 1
    return row + 1


def _write_kpi_tiles(ws: Worksheet, spec: SheetSpec, row: int) -> int:
    col = 1
    span = 2          # each tile spans 2 columns, 3 rows
    for tile in spec.kpi_tiles:
        # Label
        ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=col + span - 1)
        lc = ws.cell(row=row, column=col, value=tile.label.upper())
        lc.font = Font(size=9, bold=True, color=_WHITE)
        lc.fill = PatternFill("solid", fgColor=_BLUE)
        lc.alignment = Alignment(horizontal="center", vertical="center")
        # Value
        ws.merge_cells(start_row=row + 1, start_column=col, end_row=row + 1, end_column=col + span - 1)
        vc = ws.cell(row=row + 1, column=col, value=tile.value)
        vc.font = Font(size=16, bold=True, color=_NAVY)
        vc.fill = PatternFill("solid", fgColor=_LIGHT)
        vc.alignment = Alignment(horizontal="center", vertical="center")
        # Caption
        ws.merge_cells(start_row=row + 2, start_column=col, end_row=row + 2, end_column=col + span - 1)
        cap_color = _GREY if tile.good is None else (_GREEN if tile.good else _RED)
        cc = ws.cell(row=row + 2, column=col, value=tile.caption)
        cc.font = Font(size=9, color=cap_color)
        cc.fill = PatternFill("solid", fgColor=_LIGHT)
        cc.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[row + 1].height = 24
        col += span
        if col > 8:       # wrap to next tile row after 4 tiles
            col = 1
            row += 4
    # advance past the last (possibly partial) tile row
    return row + (4 if col != 1 else 0) + 1


def _write_table(ws: Worksheet, table: DataTable, row: int) -> tuple[int, dict]:
    """Render a table; return (next_row, anchor info for charting)."""
    tc = ws.cell(row=row, column=1, value=table.title)
    tc.font = Font(size=12, bold=True, color=_NAVY)
    row += 1
    header_row = row
    n_cols = len(table.headers)
    for j, h in enumerate(table.headers, start=1):
        hc = ws.cell(row=row, column=j, value=h)
        hc.font = Font(bold=True, color=_WHITE)
        hc.fill = PatternFill("solid", fgColor=_BLUE)
        hc.alignment = Alignment(horizontal="center")
        hc.border = _BORDER
    row += 1
    first_data = row
    for r in table.rows:
        for j, val in enumerate(r, start=1):
            cell = ws.cell(row=row, column=j, value=val)
            cell.border = _BORDER
            if table.formats and j - 1 < len(table.formats):
                cell.number_format = _NUMFMT[table.formats[j - 1]]
            if row % 2 == 0:
                cell.fill = PatternFill("solid", fgColor="F2F5FB")
        row += 1
    info = {
        "header_row": header_row, "first_data": first_data,
        "last_data": row - 1, "n_cols": n_cols,
    }
    return row + 1, info


def _add_chart(ws: Worksheet, chart: ChartSpec, anchor: str,
               data_start_row: int) -> None:
    """Write the chart's data to a helper area and attach a native chart."""
    # Place chart source data in columns J/K (10/11), out of the main view.
    cat_col, val_col = 10, 11
    ws.cell(row=data_start_row, column=cat_col, value="Category")
    ws.cell(row=data_start_row, column=val_col, value=chart.series_name)
    for i, (cat, val) in enumerate(zip(chart.categories, chart.values), start=1):
        ws.cell(row=data_start_row + i, column=cat_col, value=cat)
        ws.cell(row=data_start_row + i, column=val_col, value=val)
    n = len(chart.values)
    if n == 0:
        return

    if chart.kind == ChartKind.BAR:
        ch = BarChart(); ch.type = "col"
    elif chart.kind == ChartKind.LINE:
        ch = LineChart()
    else:
        ch = PieChart()
    ch.title = chart.title          # title sits ABOVE the chart by default
    ch.height = 7.5
    ch.width = 15

    data_ref = Reference(ws, min_col=val_col, min_row=data_start_row,
                         max_row=data_start_row + n)
    cats_ref = Reference(ws, min_col=cat_col, min_row=data_start_row + 1,
                         max_row=data_start_row + n)
    ch.add_data(data_ref, titles_from_data=True)
    ch.set_categories(cats_ref)

    # --- req 8: chart formatting ---
    if chart.kind in (ChartKind.BAR, ChartKind.LINE):
        ch.legend = None                                   # LEGEND = None
        # show both primary axes
        ch.x_axis.delete = False
        ch.y_axis.delete = False
        ch.x_axis.majorTickMark = "out"
        ch.y_axis.majorTickMark = "out"
        # keep category (date) labels readable
        ch.x_axis.tickLblPos = "low"
    if chart.kind == ChartKind.BAR:
        # Data labels = Outside End, value shown.
        ch.dataLabels = DataLabelList()
        ch.dataLabels.showVal = True
        ch.dataLabels.showSerName = False
        ch.dataLabels.showCatName = False
        ch.dataLabels.showLegendKey = False
        ch.dataLabels.dLblPos = "outEnd"
    ws.add_chart(ch, anchor)


def _write_text(ws: Worksheet, block: TextBlock, row: int) -> int:
    style = getattr(block, "style", "normal")
    _AMBER = "C05621"
    # Title styling per block style.
    tc = ws.cell(row=row, column=1, value=block.title)
    if style == "recommend":
        tc.font = Font(size=13, bold=True, color=_RED, underline="single")
    elif style == "warn":
        tc.font = Font(size=12, bold=True, color=_AMBER, underline="single")
    elif style == "highlight":
        tc.font = Font(size=12, bold=True, color=_NAVY)
    else:
        tc.font = Font(size=11, bold=True, color=_NAVY)
    row += 1

    for para in block.paragraphs:
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
        text = para
        if style == "recommend":
            text = f"➤  {para}"
            font = Font(size=11, bold=True, color=_RED)
        elif style == "warn":
            text = f"⚠  {para}"
            font = Font(size=11, bold=True, color=_AMBER)
        elif style == "highlight":
            text = f"•  {para}"
            font = Font(size=11, bold=True, color=_NAVY)
        else:
            font = Font(size=10, color=_GREY)
        pc = ws.cell(row=row, column=1, value=text)
        pc.font = font
        pc.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[row].height = max(15, 15 * (len(text) // 90 + 1))
        row += 1
    return row + 1


def _add_listobject(wb, ws: Worksheet, info: dict) -> None:
    """Register the cell block as a real Excel Table (ListObject).

    This lets the Excel COM pass find every table generically (e.g. to apply
    conditional formatting). A Table's DataBodyRange excludes the header row and
    any totals row, which is exactly what 'top-1 per column excluding total' needs.
    """
    if info["last_data"] < info["first_data"] or info["n_cols"] < 1:
        return
    ref = (f"{get_column_letter(1)}{info['header_row']}:"
           f"{get_column_letter(info['n_cols'])}{info['last_data']}")
    # Unique, valid display name across the whole workbook.
    existing = {t for s in wb.worksheets for t in getattr(s, "tables", {})}
    base = f"Tbl_{''.join(ch for ch in ws.title if ch.isalnum()) or 'S'}"
    name, k = base, 1
    while name in existing:
        k += 1
        name = f"{base}{k}"
    try:
        tbl = Table(displayName=name, ref=ref)
        # No visible table style -> keep our custom header/stripe formatting.
        tbl.tableStyleInfo = TableStyleInfo(
            name="", showFirstColumn=False, showLastColumn=False,
            showRowStripes=False, showColumnStripes=False)
        ws.add_table(tbl)
    except Exception:
        pass   # never let table registration break sheet creation


def render_sheet(wb, spec: SheetSpec) -> str:
    """Create and populate one worksheet from ``spec``. Returns its final name."""
    name = _unique_sheet_name(wb, spec.name)
    ws = wb.create_sheet(title=name)
    ws.sheet_view.showGridLines = False
    for col in range(1, 9):
        ws.column_dimensions[get_column_letter(col)].width = 16

    row = 1
    row = _write_heading(ws, spec, row)
    if spec.kpi_tiles:
        row = _write_kpi_tiles(ws, spec, row)
    for block in spec.text_blocks:
        row = _write_text(ws, block, row)

    chart_helper_row = 1
    for i, table in enumerate(spec.tables):
        row, info = _write_table(ws, table, row)
        _add_listobject(wb, ws, info)
        # Attach a matching chart (if provided by the analyzer) beside the data.
        if i < len(spec.charts):
            anchor = f"M{max(2, row - 12)}"
            _add_chart(ws, spec.charts[i], anchor, chart_helper_row)
            chart_helper_row += len(spec.charts[i].values) + 3

    # Charts with no paired table (e.g. dashboard) get stacked at the bottom.
    for j in range(len(spec.tables), len(spec.charts)):
        anchor = f"A{row}"
        _add_chart(ws, spec.charts[j], anchor, chart_helper_row)
        chart_helper_row += len(spec.charts[j].values) + 3
        row += 16
    return name


def _atomic_save(wb, source_path: str) -> None:
    folder = os.path.dirname(os.path.abspath(source_path))
    fd, tmp = tempfile.mkstemp(suffix=".xlsx", dir=folder)
    os.close(fd)
    try:
        wb.save(tmp)
        wb.close()
        os.replace(tmp, source_path)
    except Exception:
        if os.path.exists(tmp):
            os.remove(tmp)
        raise


def write_results(source_path: str, specs: list[SheetSpec]) -> list[str]:
    """Regenerate the analysis sheets (removes prior output sheets first).

    Returns the list of created sheet names. Original data sheets are untouched.
    """
    try:
        wb = load_workbook(source_path)   # full fidelity (keeps formulas/values)
    except Exception as exc:
        raise ValueError(f"Could not open workbook for writing: {exc}") from exc

    _remove_existing_outputs(wb)
    created = [render_sheet(wb, s) for s in specs if s is not None]
    _atomic_save(wb, source_path)
    return created


def append_sheets(source_path: str, specs: list[SheetSpec]) -> list[str]:
    """Add sheets WITHOUT removing the other analysis sheets. Only a same-named
    sheet (e.g. an empty 'Pivot Analysis' left by a failed Excel step) is
    replaced. Used for the static pivot fallback so KPI/Dashboard/Summary stay."""
    try:
        wb = load_workbook(source_path)
    except Exception as exc:
        raise ValueError(f"Could not open workbook for writing: {exc}") from exc

    created: list[str] = []
    for spec in specs:
        if spec is None:
            continue
        for ws in list(wb.worksheets):
            if ws.title.split(" (")[0] == spec.name and len(wb.worksheets) > 1:
                del wb[ws.title]
        created.append(render_sheet(wb, spec))
    _atomic_save(wb, source_path)
    return created
