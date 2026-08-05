#!/usr/bin/env python3
"""Inspect XLSX workbook structure without third-party dependencies."""

from __future__ import annotations

import argparse
import collections
import re
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

NS = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
REL_NS = {"r": "http://schemas.openxmlformats.org/package/2006/relationships"}
RID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"


def shared_strings(archive: zipfile.ZipFile) -> list[str]:
    try:
        root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    except KeyError:
        return []
    return ["".join(si.itertext()) for si in root.findall("x:si", NS)]


def sheet_paths(archive: zipfile.ZipFile) -> list[tuple[str, str]]:
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    targets = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels}
    result = []
    for sheet in workbook.find("x:sheets", NS) or []:
        target = targets[sheet.attrib[RID]].lstrip("/")
        if not target.startswith("xl/"):
            target = f"xl/{target}"
        result.append((sheet.attrib["name"], target))
    return result


def cell_value(cell: ET.Element, strings: list[str]) -> str:
    value = cell.findtext("x:v", default="", namespaces=NS)
    if cell.attrib.get("t") == "s" and value:
        return strings[int(value)]
    if cell.attrib.get("t") == "inlineStr":
        return "".join(cell.itertext())
    return value


def inspect(path: Path, sample_rows: int) -> None:
    print(f"\nWORKBOOK: {path}")
    with zipfile.ZipFile(path) as archive:
        strings = shared_strings(archive)
        for name, target in sheet_paths(archive):
            root = ET.fromstring(archive.read(target))
            dimension = root.find("x:dimension", NS)
            formula_counts: collections.Counter[str] = collections.Counter()
            rows: list[list[str]] = []
            formula_total = 0
            for row in root.findall(".//x:sheetData/x:row", NS):
                row_values = []
                for cell in row.findall("x:c", NS):
                    address = cell.attrib["r"]
                    formula = cell.findtext("x:f", default="", namespaces=NS)
                    value = cell_value(cell, strings)
                    if formula:
                        formula_total += 1
                        function = re.match(r"(?:_xlfn\.)?([A-Z][A-Z0-9.]*)\(", formula)
                        formula_counts[function.group(1) if function else "(expression)"] += 1
                    if value or formula:
                        shown = f"{address}={value}"
                        if formula:
                            shown += f" [{formula}]"
                        row_values.append(shown)
                if row_values and len(rows) < sample_rows:
                    rows.append(row_values)
            dim = dimension.attrib.get("ref", "?") if dimension is not None else "?"
            print(f"\n  SHEET: {name!r} range={dim} formulas={formula_total}")
            if formula_counts:
                print(f"    formula types: {formula_counts.most_common(12)}")
            for row in rows:
                print("    " + " | ".join(row[:12]))


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("paths", nargs="+", type=Path)
    parser.add_argument("--sample-rows", type=int, default=12)
    args = parser.parse_args()
    for path in args.paths:
        inspect(path, args.sample_rows)


if __name__ == "__main__":
    main()
