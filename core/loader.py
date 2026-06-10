"""Open a workbook and auto-detect its data tables.

Strategy per worksheet:
  1. Find the densest contiguous block of cells (the data region).
  2. Pick the header row: the first mostly-text row at the top of that block
     that is followed by typed data.
  3. Build a TableProfile of {column_name: value} row dicts and profile it.

We deliberately keep this tolerant: messy real-world files have title rows,
blank spacer rows, and multiple tables. We grab the largest well-formed table
per sheet and treat the biggest overall as the workbook's primary table.
"""
from __future__ import annotations

import os
from typing import Any, Optional

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from .constants import output_sheet_names
from .models import TableProfile, WorkbookProfile
from .pivot_detect import detect_pivot_sheets
from .profiler import profile_table


def _cell_is_empty(v: Any) -> bool:
    return v is None or (isinstance(v, str) and v.strip() == "")


def _detect_header_row(ws, min_row: int, max_row: int,
                       min_col: int, max_col: int) -> Optional[int]:
    """Return the 1-based row that best looks like a header, or None."""
    best_row = None
    for r in range(min_row, min(max_row, min_row + 25) + 1):
        cells = [ws.cell(row=r, column=c).value for c in range(min_col, max_col + 1)]
        non_empty = [c for c in cells if not _cell_is_empty(c)]
        if len(non_empty) < max(2, (max_col - min_col + 1) * 0.5):
            continue
        text_like = sum(1 for c in non_empty if isinstance(c, str))
        # A header row is mostly strings, and the row below it has some data.
        if text_like >= len(non_empty) * 0.6:
            below = [ws.cell(row=r + 1, column=c).value
                     for c in range(min_col, max_col + 1)]
            if any(not _cell_is_empty(c) for c in below):
                best_row = r
                break
    return best_row


def _dedupe_headers(raw: list[Any], min_col: int) -> list[str]:
    names: list[str] = []
    seen: dict[str, int] = {}
    for i, v in enumerate(raw):
        name = str(v).strip() if not _cell_is_empty(v) else \
            f"Column_{get_column_letter(min_col + i)}"
        if name in seen:
            seen[name] += 1
            name = f"{name}_{seen[name]}"
        else:
            seen[name] = 0
        names.append(name)
    return names


def _extract_table(ws) -> Optional[TableProfile]:
    if ws.max_row is None or ws.max_row < 2:
        return None
    min_col, max_col = 1, ws.max_column or 1
    min_row, max_row = 1, ws.max_row or 1

    header_row = _detect_header_row(ws, min_row, max_row, min_col, max_col)
    if header_row is None:
        return None

    raw_headers = [ws.cell(row=header_row, column=c).value
                   for c in range(min_col, max_col + 1)]
    # Trim trailing empty header columns.
    while raw_headers and _cell_is_empty(raw_headers[-1]):
        raw_headers.pop()
        max_col -= 1
    if not raw_headers:
        return None
    names = _dedupe_headers(raw_headers, min_col)

    rows: list[dict[str, Any]] = []
    # Capture the number format of each column from the first data cell that has one.
    fmt_by_name: dict[str, str] = {}
    last_data_row = header_row
    for r in range(header_row + 1, max_row + 1):
        cells = [ws.cell(row=r, column=c) for c in range(min_col, max_col + 1)]
        values = [c.value for c in cells]
        if all(_cell_is_empty(v) for v in values):
            continue
        for i, cell in enumerate(cells):
            nm = names[i]
            if nm not in fmt_by_name and not _cell_is_empty(cell.value):
                fmt = getattr(cell, "number_format", "General") or "General"
                fmt_by_name[nm] = fmt
        rows.append({names[i]: values[i] for i in range(len(names))})
        last_data_row = r

    if not rows:
        return None

    table = TableProfile(
        sheet_name=ws.title, header_row=header_row,
        first_data_row=header_row + 1, last_data_row=last_data_row,
        first_col=min_col, last_col=max_col, rows=rows,
    )
    profile_table(table, fmt_by_name)   # formats drive currency/percent detection
    return table


def load_workbook_profile(path: str) -> WorkbookProfile:
    """Open ``path`` and return a fully profiled :class:`WorkbookProfile`."""
    if not os.path.isfile(path):
        raise FileNotFoundError(f"Workbook not found: {path}")

    try:
        wb = load_workbook(path, data_only=True, read_only=True)
    except Exception as exc:  # password-protected, corrupt, wrong format
        raise ValueError(
            "Could not open this workbook. It may be password-protected, "
            "corrupted, or not a valid .xlsx file."
        ) from exc

    profile = WorkbookProfile(path=path, sheet_names=list(wb.sheetnames))
    output_names = set(output_sheet_names())

    # Sheets that already hold a PivotTable are left completely untouched.
    pivot_sheets = set(detect_pivot_sheets(path))
    profile.pivot_sheets = [s for s in wb.sheetnames if s in pivot_sheets]

    for ws in wb.worksheets:
        if ws.title in output_names:
            # Skip sheets we previously produced so re-runs don't analyze them.
            continue
        if ws.title in pivot_sheets:
            # Already contains a pivot table -> do not analyze or modify it.
            continue
        try:
            table = _extract_table(ws)
        except Exception as exc:  # noqa: BLE001 - one bad sheet shouldn't kill the run
            profile.warnings.append(f"Skipped sheet '{ws.title}': {exc}")
            continue
        if table is not None and table.measures:
            profile.tables.append(table)
        elif table is not None:
            # Keep tables without measures too (still useful for counts/pivots).
            profile.tables.append(table)

    wb.close()

    if not profile.tables:
        raise ValueError(
            "No analyzable data tables were found in this workbook. "
            "Make sure at least one sheet has a header row and rows of data."
        )

    # Primary table = the one with the most data cells.
    profile.primary_table_index = max(
        range(len(profile.tables)),
        key=lambda i: profile.tables[i].row_count * len(profile.tables[i].columns),
    )
    return profile
