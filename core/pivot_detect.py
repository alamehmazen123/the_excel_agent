"""Detect which worksheets already contain a PivotTable.

Works by inspecting the .xlsx package directly (a zip of XML parts), so it needs
neither Excel nor openpyxl's limited pivot support. A worksheet "has a pivot" if
its relationships file references a pivotTable part.

Mapping:  workbook.xml (sheet name -> r:id)
          workbook.xml.rels (r:id -> worksheets/sheetN.xml)
          worksheets/_rels/sheetN.xml.rels (-> ../pivotTables/pivotTableM.xml)
"""
from __future__ import annotations

import posixpath
import zipfile
from xml.etree import ElementTree as ET

_R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _localname(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def detect_pivot_sheets(path: str) -> list[str]:
    """Return the names of worksheets that already contain a PivotTable."""
    try:
        with zipfile.ZipFile(path) as z:
            names = set(z.namelist())
            if "xl/workbook.xml" not in names:
                return []

            # 1) sheet name -> r:id
            wb_root = ET.fromstring(z.read("xl/workbook.xml"))
            sheets: list[tuple[str, str]] = []
            for el in wb_root.iter():
                if _localname(el.tag) == "sheet":
                    rid = el.attrib.get(f"{{{_R_NS}}}id") or el.attrib.get("id")
                    name = el.attrib.get("name", "")
                    if rid and name:
                        sheets.append((name, rid))

            # 2) r:id -> worksheet part path
            rid_to_target: dict[str, str] = {}
            rels_name = "xl/_rels/workbook.xml.rels"
            if rels_name in names:
                rels_root = ET.fromstring(z.read(rels_name))
                for rel in rels_root:
                    rid = rel.attrib.get("Id", "")
                    target = rel.attrib.get("Target", "")
                    if rid and target:
                        # normalise to a full path inside xl/
                        full = posixpath.normpath(posixpath.join("xl", target))
                        rid_to_target[rid] = full

            # 3) for each sheet, look at its rels for a pivotTable reference
            pivot_sheets: list[str] = []
            for name, rid in sheets:
                part = rid_to_target.get(rid)
                if not part:
                    continue
                folder, fname = posixpath.split(part)
                sheet_rels = posixpath.join(folder, "_rels", fname + ".rels")
                if sheet_rels not in names:
                    continue
                rel_xml = z.read(sheet_rels).decode("utf-8", "ignore")
                if "pivotTable" in rel_xml:
                    pivot_sheets.append(name)
            return pivot_sheets
    except (zipfile.BadZipFile, ET.ParseError, KeyError, OSError):
        return []
