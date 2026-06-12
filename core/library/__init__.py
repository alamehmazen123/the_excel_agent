"""The Library: a persistent reference 'brain' for the engine.

It holds two kinds of hospital-specific knowledge, fed in over time from the
Excel files the user provides:

* a **header glossary** mapping abbreviated column headers to their real meaning
  (e.g. ``ADTH -> Date of Admission``), optionally tagged with a *category* that
  links the column to a code map;
* **code maps**, one per domain (guarantor, department, supplier, doctor, ...),
  mapping the hospital's numeric/string codes to their definitions
  (e.g. ``001 -> Private patient``).

The library is stored as JSON under ``core/library/data`` so it is committed to
the repo and bundled into the installer by PyInstaller. It is loaded read-only
at runtime and consulted by the Smart Tables analyzer (and, in future, the
Executive Summary / KPI analyzers) to produce hospital-relevant output.

This package imports NOTHING from the UI -- it is part of the engine.
"""
from __future__ import annotations

from .store import (CodeMap, HeaderEntry, Library, get_library, load_library,
                    norm_header)

__all__ = [
    "CodeMap",
    "HeaderEntry",
    "Library",
    "get_library",
    "load_library",
    "norm_header",
]
