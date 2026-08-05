"""Small read-only XLSX reader for the fixed source model.

The project deliberately has no runtime dependency on Excel or openpyxl.
"""

from __future__ import annotations

import re
import xml.etree.ElementTree as ET
import zipfile
from dataclasses import dataclass
from pathlib import Path

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
OFFICE_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
NS = {"x": MAIN_NS}


@dataclass(frozen=True)
class Cell:
    value: str | None
    formula: str | None

    def number(self) -> float:
        if self.value in (None, ""):
            raise ValueError("Cell has no numeric value")
        return float(self.value)


class Workbook:
    def __init__(self, path: str | Path):
        self.path = Path(path)
        self._archive = zipfile.ZipFile(self.path)
        self._strings = self._read_shared_strings()
        self._sheets = self._read_sheet_paths()
        self._cache: dict[str, dict[str, Cell]] = {}

    def close(self) -> None:
        self._archive.close()

    def __enter__(self) -> "Workbook":
        return self

    def __exit__(self, *_: object) -> None:
        self.close()

    def _read_shared_strings(self) -> list[str]:
        try:
            root = ET.fromstring(self._archive.read("xl/sharedStrings.xml"))
        except KeyError:
            return []
        return ["".join(node.itertext()) for node in root.findall("x:si", NS)]

    def _read_sheet_paths(self) -> dict[str, str]:
        workbook = ET.fromstring(self._archive.read("xl/workbook.xml"))
        rels = ET.fromstring(self._archive.read("xl/_rels/workbook.xml.rels"))
        targets = {node.attrib["Id"]: node.attrib["Target"] for node in rels}
        result = {}
        relation_id = f"{{{OFFICE_REL_NS}}}id"
        for sheet in workbook.find("x:sheets", NS) or []:
            target = targets[sheet.attrib[relation_id]].lstrip("/")
            result[sheet.attrib["name"]] = target if target.startswith("xl/") else f"xl/{target}"
        return result

    def sheet(self, name: str) -> dict[str, Cell]:
        if name in self._cache:
            return self._cache[name]
        root = ET.fromstring(self._archive.read(self._sheets[name]))
        result = {}
        for node in root.findall(".//x:sheetData/x:row/x:c", NS):
            raw = node.findtext("x:v", default=None, namespaces=NS)
            cell_type = node.attrib.get("t")
            if cell_type == "s" and raw is not None:
                raw = self._strings[int(raw)]
            elif cell_type == "inlineStr":
                inline = node.find("x:is", NS)
                raw = "".join(inline.itertext()) if inline is not None else ""
            result[node.attrib["r"]] = Cell(
                value=raw,
                formula=node.findtext("x:f", default=None, namespaces=NS),
            )
        self._cache[name] = result
        return result


CELL_REF = re.compile(r"([A-Z]+)(\d+)")


def column_cells(sheet: dict[str, Cell], column: str, first: int, last: int) -> list[Cell]:
    return [sheet[f"{column}{row}"] for row in range(first, last + 1)]
