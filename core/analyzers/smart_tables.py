"""Smart Tables analyzer: library-decoded summary tables (no pivots).

Unlike the Pivot Analysis sheet (real Excel PivotTables) this analyzer emits
plain, formatted DataTables rendered by openpyxl. Its value comes from the
:mod:`core.library` brain: code columns (guarantor / department / supplier /
doctor ...) are decoded from cryptic codes to real names, and abbreviated
headers are expanded to their real meaning, so the hospital sees readable
breakdowns like "Total Revenue by Department" with named departments.

When the library is empty the analyzer does not apply, so the sheet simply does
not appear until reference files have been ingested.
"""
from __future__ import annotations

from collections import defaultdict
from typing import Any, Optional

from ..constants import SHEET_SMART
from ..library import Library, get_library
from ..models import ColumnProfile, ColumnType, TableProfile, WorkbookProfile
from ..render import DataTable, NumberFormat, SheetSpec, TextBlock
from .base import Analyzer

# Don't build a breakdown wider than this many decoded groups.
_MAX_GROUPS = 25
# A column qualifies as a "code column" only if the library decodes at least
# this fraction of its distinct values.
_MIN_COVERAGE = 0.5


class SmartTablesAnalyzer(Analyzer):
    key = "smart_tables"
    sheet_name = SHEET_SMART

    def __init__(self, library: Optional[Library] = None) -> None:
        # Injectable for tests; defaults to the cached on-disk library.
        self._library = library

    @property
    def library(self) -> Library:
        if self._library is None:
            self._library = get_library()
        return self._library

    def applies_to(self, profile: WorkbookProfile) -> bool:
        t = profile.primary
        if not t or t.row_count == 0 or not t.value_measures:
            return False
        # Only meaningful once the library knows something this workbook uses.
        return not self.library.is_empty and bool(self._decodable_columns(t))

    def run(self, profile: WorkbookProfile) -> Optional[SheetSpec]:
        table = profile.primary
        if table is None:
            return None
        decodable = self._decodable_columns(table)
        if not decodable:
            return None

        measure = (table.value_for(profile.preferred_value_name)
                   or table.primary_value_measure)
        if measure is None:
            return None

        spec = SheetSpec(
            name=SHEET_SMART, heading="Smart Tables",
            subheading=(f"Source: {table.sheet_name}  •  {table.row_count:,} records"
                        "  •  decoded via the reference library"),
        )

        for col, cmap_name in decodable:
            title = self.library.meaning_of(col.name)
            grouped = self._grouped_decoded(table, col, measure, cmap_name)
            if not grouped:
                continue
            rows = [[label, round(total, 2), count]
                    for label, total, count in grouped]
            spec.tables.append(DataTable(
                title=f"Total {measure.name} by {title}",
                headers=[title, f"Total {measure.name}", "Records"],
                rows=rows,
                formats=[NumberFormat.GENERAL, NumberFormat.DECIMAL,
                         NumberFormat.INTEGER],
            ))

        if not spec.tables:
            return None

        spec.text_blocks.append(TextBlock(
            title="About these tables",
            paragraphs=[
                "Codes are translated to their real names using the hospital "
                "reference library. Tables are sorted by value, highest first.",
            ],
            style="normal",
        ))
        return spec

    # -- internals ---------------------------------------------------------- #
    def _decodable_columns(
            self, table: TableProfile) -> list[tuple[ColumnProfile, str]]:
        """Columns whose values the library can decode, with the map to use.

        A column is matched either via its glossary category (header -> domain)
        or by value-overlap auto-detection against every code map.
        """
        out: list[tuple[ColumnProfile, str]] = []
        lib = self.library
        for col in table.columns:
            if col.ctype in (ColumnType.CURRENCY, ColumnType.PERCENT,
                             ColumnType.DATE, ColumnType.EMPTY):
                continue
            # 1) Explicit: header is in the glossary with a category.
            entry = lib.header(col.name)
            if entry and entry.category:
                cm = lib.map_for_category(entry.category)
                if cm and cm.entries:
                    out.append((col, cm.name))
                    continue
            # 2) Auto-detect by value overlap on this column's distinct values.
            values = [v for v, _ in col.top_values] or self._distinct(table, col)
            cm = lib.best_map_for_values(values, _MIN_COVERAGE)
            if cm is not None:
                out.append((col, cm.name))
        return out

    @staticmethod
    def _distinct(table: TableProfile, col: ColumnProfile,
                  limit: int = 50) -> list[Any]:
        seen: list[Any] = []
        uniq: set[str] = set()
        for row in table.rows:
            v = row.get(col.name)
            k = str(v)
            if v is None or k in uniq:
                continue
            uniq.add(k)
            seen.append(v)
            if len(seen) >= limit:
                break
        return seen

    def _grouped_decoded(self, table: TableProfile, col: ColumnProfile,
                         measure: ColumnProfile, cmap_name: str
                         ) -> list[tuple[str, float, int]]:
        """Sum ``measure`` and count rows by the DECODED value of ``col``."""
        totals: dict[str, float] = defaultdict(float)
        counts: dict[str, int] = defaultdict(int)
        for row in table.rows:
            raw = row.get(col.name)
            if raw is None or raw == "":
                continue
            val = row.get(measure.name)
            if not isinstance(val, (int, float)) or isinstance(val, bool):
                continue
            label = self.library.decode(cmap_name, raw)
            totals[label] += float(val)
            counts[label] += 1
        ranked = sorted(totals.items(), key=lambda kv: kv[1], reverse=True)
        return [(k, v, counts[k]) for k, v in ranked[:_MAX_GROUPS]]
