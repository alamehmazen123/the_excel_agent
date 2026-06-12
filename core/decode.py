"""Library decoding shared across the whole engine.

The reference library (``core/library``) knows the hospital's code -> name maps
(account, fld1, …) and the header glossary. This module makes that knowledge
available to EVERY sheet, not just Smart Tables, by:

1. Detecting which data columns are decodable code columns.
2. Injecting a HIDDEN helper column of decoded names next to each one, so both
   the openpyxl analyzers AND the real Excel PivotTables can group by readable
   names (Excel pivots cannot look up codes themselves).

The original columns are never altered; only new hidden columns are appended.
"""
from __future__ import annotations

from collections import Counter
from dataclasses import dataclass
from typing import Any, Optional

from .library import Library, get_library
from .models import ColumnProfile, ColumnType, TableProfile

# A column qualifies as a decodable code column only if the library decodes at
# least this fraction of its distinct values.
_MIN_COVERAGE = 0.5
# Column types that may hold codes (skip money/dates/percent/blank).
_CODE_TYPES = (ColumnType.CATEGORICAL, ColumnType.TEXT, ColumnType.IDENTIFIER)


@dataclass
class DecodeCol:
    source_name: str          # e.g. "FLD1"
    cmap_name: str            # e.g. "fld1"
    helper_name: str          # e.g. "FLD1 (Name)"
    meaning: str              # glossary meaning of the source header, if any


def _distinct_values(table: TableProfile, col: ColumnProfile, limit: int = 200) -> list[Any]:
    seen: list[Any] = []
    uniq: set[str] = set()
    for row in table.rows:
        v = row.get(col.name)
        if v is None or v == "":
            continue
        k = str(v)
        if k in uniq:
            continue
        uniq.add(k)
        seen.append(v)
        if len(seen) >= limit:
            break
    return seen


def _unique_name(base: str, taken: set[str]) -> str:
    name = base
    i = 2
    while name in taken:
        name = f"{base} {i}"
        i += 1
    return name


def find_decodable(table: TableProfile, library: Optional[Library] = None,
                   min_coverage: float = _MIN_COVERAGE) -> list[DecodeCol]:
    """Code columns in ``table`` the library can decode, with the map to use."""
    lib = library if library is not None else get_library()
    if lib.is_empty:
        return []

    taken = {c.name for c in table.columns}
    out: list[DecodeCol] = []
    for col in table.columns:
        if col.is_decoded_helper or col.ctype not in _CODE_TYPES:
            continue
        cmap = None
        # 1) Explicit: header is in the glossary with a category -> that map.
        entry = lib.header(col.name)
        if entry and entry.category:
            cmap = lib.map_for_category(entry.category)
        # 2) Auto-detect by value overlap against every code map.
        if cmap is None or not cmap.entries:
            cmap = lib.best_map_for_values(_distinct_values(table, col), min_coverage)
        if cmap is None or not cmap.entries:
            continue
        helper = _unique_name(f"{col.name} (Name)", taken)
        taken.add(helper)
        meaning = entry.meaning if entry else ""
        out.append(DecodeCol(col.name, cmap.name, helper, meaning))
    return out


def apply_to_profile(table: TableProfile, decodes: list[DecodeCol],
                     library: Optional[Library] = None) -> None:
    """Add decoded values to ``table.rows`` and a CATEGORICAL helper column to
    ``table.columns`` for each decode; flag the source columns as decoded."""
    if not decodes:
        return
    lib = library if library is not None else get_library()
    by_name = {c.name: c for c in table.columns}

    for dc in decodes:
        src = by_name.get(dc.source_name)
        if src is None:
            continue
        decoded_values: list[str] = []
        for row in table.rows:
            raw = row.get(dc.source_name)
            name = lib.decode(dc.cmap_name, raw) if raw not in (None, "") else None
            row[dc.helper_name] = name
            if name:
                decoded_values.append(name)
        counts = Counter(decoded_values)
        helper = ColumnProfile(
            name=dc.helper_name, index=len(table.columns),
            ctype=ColumnType.CATEGORICAL,
            count=len(decoded_values), nulls=len(table.rows) - len(decoded_values),
            distinct=len(counts), number_format="General",
            top_values=counts.most_common(10), is_decoded_helper=True,
        )
        table.columns.append(helper)
        src.decoded_helper = dc.helper_name


def decoded_values_for(table: TableProfile, dc: DecodeCol,
                       library: Optional[Library] = None) -> dict[Any, str]:
    """Map each raw code value present in the column to its decoded name
    (used by the writer to fill the hidden helper column in the sheet)."""
    lib = library if library is not None else get_library()
    out: dict[Any, str] = {}
    for row in table.rows:
        raw = row.get(dc.source_name)
        if raw in (None, "") or raw in out:
            continue
        out[raw] = lib.decode(dc.cmap_name, raw)
    return out
