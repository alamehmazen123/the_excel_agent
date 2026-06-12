"""Smart Tables analyzer: library-decoded, date-grouped summary tables.

Unlike the Pivot Analysis sheet (real Excel PivotTables) this analyzer emits
plain, formatted DataTables rendered by openpyxl. Its value comes from the
:mod:`core.library` brain: code columns (guarantor / department / supplier /
account …) are decoded from cryptic codes to real names via the hidden helper
columns the engine injects, abbreviated headers are expanded to their meaning,
and money is shown in Lebanese Pounds (the hospital default).

It is a SCENARIO GENERATOR: many tables covering each readable dimension against
each value measure, always grouped by Year/Month (BASIC RULE: never a table
without a date). When the library is empty or the data has no date column, the
analyzer does not apply, so the sheet simply does not appear.
"""
from __future__ import annotations

from typing import Optional

from ..aggregate import group_period_dim, time_series
from ..constants import SHEET_SMART
from ..formatting import is_dollar_column
from ..library import Library, get_library
from ..models import ColumnProfile, TableProfile, WorkbookProfile
from ..render import DataTable, NumberFormat, SheetSpec, TextBlock
from .base import Analyzer

# Keep the sheet readable: cap the number of scenario tables and rows per table.
_MAX_TABLES = 14
_ROWS_PER_TABLE = 30

_MONTHS = ["", "Jan", "Feb", "Mar", "Apr", "May", "Jun",
           "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


def _fmt_period(period: str) -> str:
    """'2026-01' -> 'Jan-26'."""
    try:
        y, m = period.split("-")
        return f"{_MONTHS[int(m)]}-{y[2:]}"
    except Exception:
        return period


class SmartTablesAnalyzer(Analyzer):
    key = "smart_tables"
    sheet_name = SHEET_SMART

    def __init__(self, library: Optional[Library] = None) -> None:
        self._library = library

    @property
    def library(self) -> Library:
        if self._library is None:
            self._library = get_library()
        return self._library

    def applies_to(self, profile: WorkbookProfile) -> bool:
        t = profile.primary
        if not t or t.row_count == 0 or not t.value_measures or not t.date_columns:
            return False
        if self.library.is_empty:
            return False
        # Only meaningful once the library has DECODED something in this workbook
        # (a hidden helper column was injected). Otherwise the Pivot/Dashboard
        # sheets already cover plain categorical breakdowns.
        return any(c.is_decoded_helper for c in t.columns)

    # -- internals ---------------------------------------------------------- #
    def _dimensions(self, table: TableProfile) -> list[ColumnProfile]:
        """Readable grouping columns: decoded helpers first, then plain
        categoricals (already human-readable name columns)."""
        helpers = [c for c in table.columns if c.is_decoded_helper]
        cats = [c for c in table.dimensions if not c.is_decoded_helper]
        return helpers + cats

    def _fmt_for(self, measure: ColumnProfile) -> NumberFormat:
        return (NumberFormat.CURRENCY if is_dollar_column(measure)
                else NumberFormat.LBP)

    def run(self, profile: WorkbookProfile) -> Optional[SheetSpec]:
        table = profile.primary
        if table is None or not table.date_columns or not table.value_measures:
            return None

        date_col = table.date_columns[0]
        dims = self._dimensions(table)
        chosen = table.measures_for(profile.preferred_measure_names)
        value_measures = [m for m in chosen if m.is_value] or table.value_measures

        spec = SheetSpec(
            name=SHEET_SMART, heading="Smart Tables",
            subheading=(f"Source: {table.sheet_name}  •  {table.row_count:,} records"
                        f"  •  decoded via the reference library, grouped by month"),
        )

        produced = 0

        # Scenario A: each value measure totalled by month (the time trend).
        for m in value_measures[:2]:
            series = time_series(table, date_col, m)
            if not series:
                continue
            spec.tables.append(DataTable(
                title=f"Total {self.library.meaning_of(m.name)} by {date_col.name} (Month)",
                headers=[f"{date_col.name} (Month)", f"Total {m.name}"],
                rows=[[_fmt_period(p), round(v, 2)] for p, v in series],
                formats=[NumberFormat.GENERAL, self._fmt_for(m)],
                bar_columns=[1],          # data bar on the monthly total
            ))
            produced += 1

        # Scenario B: each readable dimension x each value measure, month-grouped.
        for dim in dims:
            if produced >= _MAX_TABLES:
                break
            dim_title = self.library.meaning_of(dim.name)
            for m in value_measures[:2]:
                if produced >= _MAX_TABLES:
                    break
                rows = group_period_dim(table, date_col, dim, m,
                                        top_n=_ROWS_PER_TABLE)
                if not rows:
                    continue
                spec.tables.append(DataTable(
                    title=f"Total {m.name} by {date_col.name} (Month) & {dim_title}",
                    headers=[f"{date_col.name} (Month)", dim_title,
                             f"Total {m.name}", "Records"],
                    rows=[[_fmt_period(p), label, round(t, 2), c]
                          for p, label, t, c in rows],
                    formats=[NumberFormat.GENERAL, NumberFormat.GENERAL,
                             self._fmt_for(m), NumberFormat.INTEGER],
                    bar_columns=[2],          # data bar on the value total
                ))
                produced += 1

        if not spec.tables:
            return None

        spec.text_blocks.append(TextBlock(
            title="About these tables",
            paragraphs=[
                "Codes are translated to their real names using the hospital "
                "reference library, money is shown in Lebanese Pounds (LBP), and "
                "every table is grouped by month. Tables are sorted by value.",
            ],
            style="normal",
        ))

        helpers = [c.name for c in table.columns if c.is_decoded_helper]
        if helpers:
            spec.text_blocks.append(TextBlock(
                title="Hidden helper columns added by the agent",
                paragraphs=[
                    "To let the PivotTables and these tables show real names, the "
                    "agent added these HIDDEN columns to the data sheet (originals "
                    "untouched): " + ", ".join(helpers) + ".",
                ],
                style="highlight",
            ))
        return spec
