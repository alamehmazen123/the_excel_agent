"""Load, query, and persist the reference library (the engine's 'brain').

The on-disk form is three small JSON files in ``data/``:

* ``headers.json`` -- the abbreviation glossary.
* ``codes.json``   -- the per-domain code maps.
* ``meta.json``    -- bookkeeping (version, ingested sources).

Everything is keyed by NORMALIZED forms so matching is tolerant of casing,
whitespace, and the int-vs-zero-padded ambiguity ("1" == "001") that arises
because Excel often reads a code column as integers.
"""
from __future__ import annotations

import json
import os
import re
from dataclasses import dataclass, field
from datetime import datetime, timezone
from typing import Any, Optional

_DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
_HEADERS_FILE = os.path.join(_DATA_DIR, "headers.json")
_CODES_FILE = os.path.join(_DATA_DIR, "codes.json")
_META_FILE = os.path.join(_DATA_DIR, "meta.json")


# --------------------------------------------------------------------------- #
# Normalization helpers (the heart of tolerant matching)                      #
# --------------------------------------------------------------------------- #
def norm_header(value: Any) -> str:
    """Canonical key for a header/abbreviation: upper, trimmed, single-spaced."""
    s = str(value or "").strip().upper()
    return re.sub(r"\s+", " ", s)


def _squash(value: Any) -> str:
    """Aggressive header key: alphanumerics only (for fuzzy header matching)."""
    return re.sub(r"[^A-Z0-9]", "", norm_header(value))


def norm_code(value: Any) -> str:
    """Canonical key for a code value: trimmed string, no trailing ``.0``."""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return str(value).strip()


def code_variants(value: Any) -> list[str]:
    """All plausible lookup keys for a code cell (handles 1 == 001 == '001 ')."""
    base = norm_code(value)
    if not base:
        return []
    variants = {base, base.upper()}
    # Numeric forms: int(001) -> "1"; also re-pad is impossible without a width,
    # so we index both literal and stripped-int forms at ingest time instead.
    if re.fullmatch(r"\d+", base):
        variants.add(str(int(base)))          # "001" -> "1"
    return list(variants)


# --------------------------------------------------------------------------- #
# Data model                                                                  #
# --------------------------------------------------------------------------- #
@dataclass
class HeaderEntry:
    abbrev: str                  # normalized key, e.g. "ADTH"
    meaning: str                 # "Date of Admission"
    category: str = ""           # links to a CodeMap name, e.g. "department"
    notes: str = ""


@dataclass
class CodeMap:
    """A domain dictionary of code -> definition (e.g. all guarantors).

    Optionally each code also carries a CATEGORY (e.g. account 713… → "revenues",
    601… → "purchases"). The category lets the engine infer what a workbook is
    ABOUT (revenue vs expense) from the codes it contains."""
    name: str                    # "guarantor", "department", ...
    label: str = ""              # human label for the domain
    entries: dict[str, str] = field(default_factory=dict)   # canonical -> definition
    categories: dict[str, str] = field(default_factory=dict)  # canonical -> category
    # Lookup indices built from ``entries`` / ``categories`` over every variant.
    _index: dict[str, str] = field(default_factory=dict, repr=False)
    _cat_index: dict[str, str] = field(default_factory=dict, repr=False)

    def reindex(self) -> "CodeMap":
        self._index = {}
        for code, definition in self.entries.items():
            for v in code_variants(code):
                # First write wins so an explicit literal beats an int collision.
                self._index.setdefault(v, definition)
        self._cat_index = {}
        for code, category in self.categories.items():
            for v in code_variants(code):
                self._cat_index.setdefault(v, category)
        return self

    def lookup(self, value: Any) -> Optional[str]:
        for v in code_variants(value):
            hit = self._index.get(v)
            if hit is not None:
                return hit
        return None

    def category_of(self, value: Any) -> Optional[str]:
        """The financial category of a code (e.g. 'revenues'), if known."""
        for v in code_variants(value):
            hit = self._cat_index.get(v)
            if hit is not None:
                return hit
        return None

    def coverage(self, values: list[Any]) -> float:
        """Fraction of distinct ``values`` this map can decode (0.0-1.0)."""
        seen = [v for v in values if norm_code(v)]
        if not seen:
            return 0.0
        hits = sum(1 for v in seen if self.lookup(v) is not None)
        return hits / len(seen)


@dataclass
class Library:
    headers: dict[str, HeaderEntry] = field(default_factory=dict)   # by norm_header
    code_maps: dict[str, CodeMap] = field(default_factory=dict)     # by map name
    meta: dict[str, Any] = field(default_factory=dict)

    # -- introspection ------------------------------------------------------ #
    @property
    def is_empty(self) -> bool:
        return not self.headers and not any(m.entries for m in self.code_maps.values())

    def header(self, name: Any) -> Optional[HeaderEntry]:
        key = norm_header(name)
        entry = self.headers.get(key)
        if entry is not None:
            return entry
        # fall back to alphanumeric-squashed match
        squashed = _squash(name)
        for k, e in self.headers.items():
            if _squash(k) == squashed:
                return e
        return None

    def meaning_of(self, name: Any) -> str:
        """Real meaning of a header, or the original name if unknown."""
        entry = self.header(name)
        return entry.meaning if entry else str(name)

    def map_for_category(self, category: str) -> Optional[CodeMap]:
        if not category:
            return None
        return self.code_maps.get(category) or self.code_maps.get(category.lower())

    def decode(self, map_name: str, value: Any) -> str:
        """Decode a single code via a named map, or return the raw value."""
        cm = self.code_maps.get(map_name)
        hit = cm.lookup(value) if cm else None
        return hit if hit is not None else str(value)

    def best_map_for_values(self, values: list[Any],
                            min_coverage: float = 0.5) -> Optional[CodeMap]:
        """Auto-detect which code map a column of raw values belongs to."""
        best: Optional[CodeMap] = None
        best_cov = min_coverage
        for cm in self.code_maps.values():
            if not cm.entries:
                continue
            cov = cm.coverage(values)
            if cov >= best_cov:
                best, best_cov = cm, cov
        return best


# --------------------------------------------------------------------------- #
# Persistence                                                                 #
# --------------------------------------------------------------------------- #
def _read_json(path: str, default: Any) -> Any:
    try:
        with open(path, "r", encoding="utf-8") as fh:
            return json.load(fh)
    except (FileNotFoundError, ValueError):
        return default


def load_library() -> Library:
    """Load the library from ``core/library/data`` (empty if files are absent)."""
    headers_raw = _read_json(_HEADERS_FILE, {}).get("entries", {})
    headers = {
        norm_header(k): HeaderEntry(
            abbrev=norm_header(k),
            meaning=v.get("meaning", ""),
            category=v.get("category", ""),
            notes=v.get("notes", ""),
        )
        for k, v in headers_raw.items()
    }

    maps_raw = _read_json(_CODES_FILE, {}).get("maps", {})
    code_maps: dict[str, CodeMap] = {}
    for name, m in maps_raw.items():
        entries = {norm_code(k): str(val) for k, val in m.get("entries", {}).items()}
        categories = {norm_code(k): str(val).lower()
                      for k, val in m.get("categories", {}).items()}
        code_maps[name] = CodeMap(
            name=name, label=m.get("label", name), entries=entries,
            categories=categories,
        ).reindex()

    meta = _read_json(_META_FILE, {})
    return Library(headers=headers, code_maps=code_maps, meta=meta)


def save_library(lib: Library) -> None:
    """Persist the library back to disk (used by the ingest tool)."""
    os.makedirs(_DATA_DIR, exist_ok=True)
    headers_out = {
        "version": 1,
        "entries": {
            e.abbrev: {"meaning": e.meaning, "category": e.category, "notes": e.notes}
            for e in lib.headers.values()
        },
    }
    codes_out = {
        "version": 1,
        "maps": {
            name: ({"label": cm.label, "entries": cm.entries}
                   | ({"categories": cm.categories} if cm.categories else {}))
            for name, cm in lib.code_maps.items()
        },
    }
    lib.meta["updated"] = datetime.now(timezone.utc).isoformat()
    with open(_HEADERS_FILE, "w", encoding="utf-8") as fh:
        json.dump(headers_out, fh, ensure_ascii=False, indent=2, sort_keys=True)
    with open(_CODES_FILE, "w", encoding="utf-8") as fh:
        json.dump(codes_out, fh, ensure_ascii=False, indent=2, sort_keys=True)
    with open(_META_FILE, "w", encoding="utf-8") as fh:
        json.dump(lib.meta, fh, ensure_ascii=False, indent=2, sort_keys=True)


# Process-wide cache so each run loads the JSON once.
_CACHE: Optional[Library] = None


def get_library(refresh: bool = False) -> Library:
    global _CACHE
    if _CACHE is None or refresh:
        _CACHE = load_library()
    return _CACHE
