#!/usr/bin/env python3
"""Set scalar XLSX cells without recalculating formulas."""

from __future__ import annotations

import argparse
import shutil
import sys
import tempfile
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from model.xlsx_reader import MAIN_NS, Workbook

ET.register_namespace("", MAIN_NS)
NS = {"x": MAIN_NS}


def set_cells(path: Path, sheet_name: str, updates: dict[str, float]) -> None:
    with Workbook(path) as workbook:
        sheet_path = workbook._sheets[sheet_name]  # Internal path is the intended use here.
    with zipfile.ZipFile(path, "r") as source:
        xml = ET.fromstring(source.read(sheet_path))
        for address, value in updates.items():
            cell = xml.find(f".//x:c[@r='{address}']", NS)
            if cell is None:
                raise KeyError(f"Cell {sheet_name}!{address} does not exist")
            formula = cell.find("x:f", NS)
            if formula is not None:
                cell.remove(formula)
            value_node = cell.find("x:v", NS)
            if value_node is None:
                value_node = ET.SubElement(cell, f"{{{MAIN_NS}}}v")
            value_node.text = repr(value)
            cell.attrib.pop("t", None)

        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False, dir=path.parent) as temp:
            temp_path = Path(temp.name)
        try:
            with zipfile.ZipFile(temp_path, "w", zipfile.ZIP_DEFLATED) as target:
                for item in source.infolist():
                    data = ET.tostring(xml, encoding="utf-8", xml_declaration=True) if item.filename == sheet_path else source.read(item.filename)
                    target.writestr(item, data)
            shutil.move(temp_path, path)
        finally:
            temp_path.unlink(missing_ok=True)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("path", type=Path)
    parser.add_argument("sheet")
    parser.add_argument("updates", nargs="+", help="CELL=NUMBER")
    args = parser.parse_args()
    updates = {}
    for update in args.updates:
        address, value = update.split("=", 1)
        updates[address] = float(value)
    set_cells(args.path, args.sheet, updates)


if __name__ == "__main__":
    main()
