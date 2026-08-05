#!/usr/bin/env python3
"""Add the fixed LEND solar export price to the household Excel model."""

from __future__ import annotations

import shutil
import sys
import tempfile
import xml.etree.ElementTree as ET
import zipfile
from copy import deepcopy
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from model.xlsx_reader import MAIN_NS, Workbook

NS = {"x": MAIN_NS}
TAG = f"{{{MAIN_NS}}}"
ET.register_namespace("", MAIN_NS)


def inline_cell(address: str, text: str, style: str | None = None) -> ET.Element:
    attributes = {"r": address, "t": "inlineStr"}
    if style is not None:
        attributes["s"] = style
    cell = ET.Element(f"{TAG}c", attributes)
    inline = ET.SubElement(cell, f"{TAG}is")
    ET.SubElement(inline, f"{TAG}t").text = text
    return cell


def number_cell(address: str, value: float, style: str | None = None) -> ET.Element:
    attributes = {"r": address}
    if style is not None:
        attributes["s"] = style
    cell = ET.Element(f"{TAG}c", attributes)
    ET.SubElement(cell, f"{TAG}v").text = repr(value)
    return cell


def migrate(path: Path) -> None:
    with Workbook(path) as workbook:
        paths = {
            name: workbook._sheets[name]
            for name in ("Aannames", "Model batterij LEND", "Kostenoverzicht")
        }

    with zipfile.ZipFile(path, "r") as source:
        documents = {
            name: ET.fromstring(source.read(sheet_path))
            for name, sheet_path in paths.items()
        }

        assumptions = documents["Aannames"]
        sheet_data = assumptions.find("x:sheetData", NS)
        if sheet_data is None:
            raise ValueError("Aannames has no sheetData")
        existing = assumptions.find(".//x:c[@r='B38']", NS)
        if existing is None:
            reference_row = assumptions.find(".//x:row[@r='34']", NS)
            styles = {}
            if reference_row is not None:
                for cell in reference_row.findall("x:c", NS):
                    styles[cell.attrib["r"][0]] = cell.attrib.get("s")
            row = ET.SubElement(sheet_data, f"{TAG}row", {"r": "38"})
            row.extend(
                (
                    inline_cell(
                        "A38",
                        "Vaste verkoopprijs zonnepanelen LEND",
                        styles.get("A"),
                    ),
                    number_cell("B38", 0.025, styles.get("B")),
                    inline_cell(
                        "C38",
                        "€/kWh; alleen voor directe teruglevering van zonne-energie in het LEND-model",
                        styles.get("C"),
                    ),
                )
            )
        dimension = assumptions.find("x:dimension", NS)
        if dimension is not None:
            dimension.attrib["ref"] = "A1:C38"

        lend = documents["Model batterij LEND"]
        changed = 0
        for row_number in range(2, 8762):
            formula = lend.find(f".//x:c[@r='N{row_number}']/x:f", NS)
            if formula is None:
                raise KeyError(f"Missing formula N{row_number}")
            old = f"-(F{row_number}-H{row_number})*E{row_number}"
            new = f"-(F{row_number}-H{row_number})*Aannames!$B$38"
            if old in (formula.text or ""):
                formula.text = formula.text.replace(old, new)
                changed += 1
        if changed != 8760:
            raise ValueError(f"Expected 8760 changed LEND formulas, got {changed}")

        costs = documents["Kostenoverzicht"]
        explanation = costs.find(".//x:c[@r='G9']/x:is/x:t", NS)
        if explanation is not None:
            explanation.text = (
                "Netladen gebruikt een 24-uurs zonnevooruitblik en reserveert "
                "capaciteit voor verwacht zonne-overschot. In het LEND-model "
                "gebruikt batterijexport de vaste batterijprijs en directe "
                "zonne-export de vaste zonnepanelenprijs."
            )

        with tempfile.NamedTemporaryFile(
            suffix=".xlsx", delete=False, dir=path.parent
        ) as temp:
            temp_path = Path(temp.name)
        try:
            by_path = {paths[name]: root for name, root in documents.items()}
            with zipfile.ZipFile(temp_path, "w", zipfile.ZIP_DEFLATED) as target:
                for item in source.infolist():
                    root = by_path.get(item.filename)
                    data = (
                        ET.tostring(root, encoding="utf-8", xml_declaration=True)
                        if root is not None
                        else source.read(item.filename)
                    )
                    target.writestr(item, data)
            shutil.move(temp_path, path)
        finally:
            temp_path.unlink(missing_ok=True)


if __name__ == "__main__":
    if len(sys.argv) != 2:
        raise SystemExit(f"Usage: {sys.argv[0]} WORKBOOK.xlsx")
    migrate(Path(sys.argv[1]))
