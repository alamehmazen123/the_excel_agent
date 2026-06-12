"""Merge a hospital reference Excel file into the library (the engine 'brain').

Run this once per file the user provides; it updates the JSON under
``core/library/data`` in place, which the next build then ships to colleagues.

Usage:
    python tools/ingest_library.py <file.xlsx> [--kind headers|codes]
                                               [--category department]
    python tools/ingest_library.py --show          # print library stats

Examples:
    # A glossary of abbreviated headers -> meanings (kind auto-detected):
    python tools/ingest_library.py "Headers Glossary.xlsx"

    # A codes file for one domain (category from sheet name unless given):
    python tools/ingest_library.py "Guarantor Codes.xlsx" --kind codes --category guarantor
"""
from __future__ import annotations

import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core.library import get_library                       # noqa: E402
from core.library.ingest import ingest_excel               # noqa: E402


def _show() -> None:
    lib = get_library(refresh=True)
    print(f"Headers in glossary : {len(lib.headers)}")
    print(f"Code maps           : {len(lib.code_maps)}")
    for name, cm in sorted(lib.code_maps.items()):
        print(f"  - {name:<18} {len(cm.entries):>6} codes")
    sources = lib.meta.get("sources", [])
    print(f"Ingested sources    : {len(sources)}")


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(description="Ingest a reference Excel into the library.")
    p.add_argument("file", nargs="?", help="Path to the .xlsx reference file")
    p.add_argument("--kind", choices=["headers", "codes"],
                   help="Force the file kind (otherwise auto-detected per sheet)")
    p.add_argument("--category", help="Code domain for a codes file (e.g. department)")
    p.add_argument("--show", action="store_true", help="Print library stats and exit")
    args = p.parse_args(argv)

    if args.show or not args.file:
        _show()
        return 0

    if not os.path.exists(args.file):
        print(f"File not found: {args.file}", file=sys.stderr)
        return 2

    report = ingest_excel(args.file, kind=args.kind, category=args.category)
    print(f"Ingested {report.source}: {report.summary()}")
    print(f"  sheets: {', '.join(report.sheets) or '(none)'}")
    print()
    _show()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
