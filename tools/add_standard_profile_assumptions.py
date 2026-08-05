#!/usr/bin/env python3
"""Add a standard-profile assumptions sheet to the household Excel model."""

from __future__ import annotations

import argparse
import csv
import re
import shutil
import tempfile
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
OFFICE_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CONTENT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
APP_NS = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
VT_NS = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"

NS = {"x": MAIN_NS, "rel": REL_NS, "ct": CONTENT_NS, "app": APP_NS, "vt": VT_NS}
TAG = f"{{{MAIN_NS}}}"
REL_TAG = f"{{{REL_NS}}}"
CT_TAG = f"{{{CONTENT_NS}}}"
APP_TAG = f"{{{APP_NS}}}"
VT_TAG = f"{{{VT_NS}}}"
RID = f"{{{OFFICE_REL_NS}}}id"
SHEET_NAME = "Aannames stdprofielen"

ET.register_namespace("r", OFFICE_REL_NS)
ET.register_namespace("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006")
ET.register_namespace("x15", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main")
ET.register_namespace("xr", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision")
ET.register_namespace("xr6", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision6")
ET.register_namespace("xr10", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision10")
ET.register_namespace("xr2", "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2")
ET.register_namespace("xcalcf", "http://schemas.microsoft.com/office/spreadsheetml/2018/calcfeatures")
ET.register_namespace("vt", VT_NS)


def column_name(index: int) -> str:
    result = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        result = chr(65 + remainder) + result
    return result


def inline_cell(address: str, text: str) -> ET.Element:
    cell = ET.Element(f"{TAG}c", {"r": address, "t": "inlineStr"})
    inline = ET.SubElement(cell, f"{TAG}is")
    ET.SubElement(inline, f"{TAG}t").text = text
    return cell


def number_cell(address: str, value: float) -> ET.Element:
    cell = ET.Element(f"{TAG}c", {"r": address})
    ET.SubElement(cell, f"{TAG}v").text = f"{value:.15g}"
    return cell


def row(row_number: int, values: list[str | float | int | None]) -> ET.Element:
    node = ET.Element(f"{TAG}row", {"r": str(row_number)})
    for index, value in enumerate(values, start=1):
        if value is None:
            continue
        address = f"{column_name(index)}{row_number}"
        if isinstance(value, str):
            node.append(inline_cell(address, value))
        else:
            node.append(number_cell(address, float(value)))
    return node


def profile_columns(csv_path: Path) -> tuple[list[str], dict[str, list[float]], dict[str, dict[str, str]], list[tuple[str, str]]]:
    rows = list(csv.reader(csv_path.open(encoding="latin-1", newline=""), delimiter=";"))
    profile_count = len(rows[2]) - 3
    versions = rows[0][3 : 3 + profile_count]
    categories = rows[2][3 : 3 + profile_count]
    types = rows[3][3 : 3 + profile_count]
    directions = rows[4][3 : 3 + profile_count]

    data_rows = rows[7:]
    intervals = [(data_rows[i][1], data_rows[i + 3][2]) for i in range(0, len(data_rows), 4)]
    profiles: dict[str, list[float]] = {}
    metadata: dict[str, dict[str, str]] = {}
    for offset, version in enumerate(versions):
        column = 3 + offset
        quarters = [float(data[column]) for data in data_rows]
        profiles[version] = [
            sum(quarters[start : start + 4]) for start in range(0, len(quarters), 4)
        ]
        metadata[version] = {
            "categorie": categories[offset],
            "type": types[offset],
            "richting": directions[offset],
        }
    return versions, profiles, metadata, intervals


def make_sheet(csv_path: Path, usage_profile: str, injection_profile: str) -> ET.Element:
    _, profiles, metadata, intervals = profile_columns(csv_path)
    if usage_profile not in profiles:
        raise KeyError(f"Unknown usage profile {usage_profile}")
    if injection_profile not in profiles:
        raise KeyError(f"Unknown injection profile {injection_profile}")

    usage = profiles[usage_profile]
    injection = profiles[injection_profile]
    if len(usage) != 8760 or len(injection) != 8760:
        raise ValueError("Expected 8760 hourly values")

    root = ET.Element(f"{TAG}worksheet")
    ET.SubElement(root, f"{TAG}dimension", {"ref": "A1:G8775"})
    sheet_views = ET.SubElement(root, f"{TAG}sheetViews")
    ET.SubElement(sheet_views, f"{TAG}sheetView", {"workbookViewId": "0"})
    ET.SubElement(root, f"{TAG}sheetFormatPr", {"defaultRowHeight": "15"})
    cols = ET.SubElement(root, f"{TAG}cols")
    for min_col, max_col, width in (
        (1, 1, "8"),
        (2, 3, "19"),
        (4, 7, "18"),
    ):
        ET.SubElement(
            cols,
            f"{TAG}col",
            {
                "min": str(min_col),
                "max": str(max_col),
                "width": width,
                "customWidth": "1",
            },
        )
    sheet_data = ET.SubElement(root, f"{TAG}sheetData")

    usage_meta = metadata[usage_profile]
    injection_meta = metadata[injection_profile]
    metadata_rows: list[list[str | float | int | None]] = [
        ["Standaardprofiel-aannames elektriciteit 2026"],
        ["Bron", "Energiedatawijzer - Profielen elektriciteit 2026"],
        ["Bronbestand", csv_path.name],
        ["CSV-versie gebruik", usage_profile],
        ["Toepassingsjaar", 2026],
        ["Aggregatie", "Kwartierwaarden uit de CSV zijn per 4 opgeteld naar uurfracties; kWh-kolommen zijn genormaliseerd op de jaarwaarde."],
        ["Afnameprofiel", usage_profile],
        ["Afname categorie/type/richting", f"{usage_meta['categorie']} / {usage_meta['type']} / {usage_meta['richting']}"],
        ["Jaarverbruik voor schaal", 4000, "kWh"],
        ["Invoedingsprofiel", injection_profile],
        ["Invoeding categorie/type/richting", f"{injection_meta['categorie']} / {injection_meta['type']} / {injection_meta['richting']}"],
        ["Jaarinvoeding voor schaal", 4500, "kWh"],
        ["Let op", "Richting I is een standaardprofiel voor invoeding, niet hetzelfde als bruto PV-opwek achter de meter."],
        ["Gebruik in model", "Referentieblad; de bestaande berekeningen verwijzen nog naar het tabblad Model zonder batterij."],
    ]
    for row_number, values in enumerate(metadata_rows, start=1):
        sheet_data.append(row(row_number, values))

    header_row = 15
    sheet_data.append(
        row(
            header_row,
            [
                "Uur",
                "Van",
                "Tot",
                "Afname fractie",
                "Afname kWh bij 4000",
                "Invoeding fractie",
                "Invoeding kWh bij 4500",
            ],
        )
    )
    usage_total = sum(usage)
    injection_total = sum(injection)
    for index, ((start, end), usage_fraction, injection_fraction) in enumerate(
        zip(intervals, usage, injection), start=1
    ):
        sheet_data.append(
            row(
                header_row + index,
                [
                    index,
                    start,
                    end,
                    usage_fraction,
                    usage_fraction / usage_total * 4000,
                    injection_fraction,
                    injection_fraction / injection_total * 4500,
                ],
            )
        )

    ET.SubElement(root, f"{TAG}pageMargins", {"left": "0.7", "right": "0.7", "top": "0.75", "bottom": "0.75", "header": "0.3", "footer": "0.3"})
    return root


def workbook_xml(root: ET.Element) -> bytes:
    """Serialize workbook.xml with prefixes required by mc:Ignorable values."""
    text = ET.tostring(root, encoding="unicode", xml_declaration=True)
    required_namespaces = {
        "x15": "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main",
        "xr": "http://schemas.microsoft.com/office/spreadsheetml/2014/revision",
        "xr6": "http://schemas.microsoft.com/office/spreadsheetml/2016/revision6",
        "xr10": "http://schemas.microsoft.com/office/spreadsheetml/2016/revision10",
        "xr2": "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2",
    }
    if "Ignorable=" in text:
        declarations = " ".join(
            f'xmlns:{prefix}="{namespace}"'
            for prefix, namespace in required_namespaces.items()
            if f"xmlns:{prefix}=" not in text
        )
        if declarations:
            text = text.replace("<ns0:workbook ", f"<ns0:workbook {declarations} ", 1)
    return text.encode("utf-8")


def max_rid(root: ET.Element) -> int:
    result = 0
    for rel in root.findall("rel:Relationship", NS):
        match = re.fullmatch(r"rId(\d+)", rel.attrib["Id"])
        if match:
            result = max(result, int(match.group(1)))
    return result


def add_or_replace_sheet(workbook_path: Path, csv_path: Path, usage_profile: str, injection_profile: str) -> None:
    sheet_xml = ET.tostring(make_sheet(csv_path, usage_profile, injection_profile), encoding="utf-8", xml_declaration=True)
    with zipfile.ZipFile(workbook_path, "r") as source:
        workbook = ET.fromstring(source.read("xl/workbook.xml"))
        rels = ET.fromstring(source.read("xl/_rels/workbook.xml.rels"))
        content_types = ET.fromstring(source.read("[Content_Types].xml"))
        app = ET.fromstring(source.read("docProps/app.xml"))

        sheets = workbook.find("x:sheets", NS)
        if sheets is None:
            raise ValueError("Workbook has no sheets element")

        existing = None
        for sheet in sheets.findall("x:sheet", NS):
            if sheet.attrib["name"] == SHEET_NAME:
                existing = sheet
                break

        if existing is None:
            existing_sheet_ids = [int(sheet.attrib["sheetId"]) for sheet in sheets.findall("x:sheet", NS)]
            sheet_id = max(existing_sheet_ids) + 1
            rid = f"rId{max_rid(rels) + 1}"
            existing_paths = {
                rel.attrib["Target"]
                for rel in rels.findall("rel:Relationship", NS)
                if rel.attrib["Type"].endswith("/worksheet")
            }
            sheet_number = 1
            while f"worksheets/sheet{sheet_number}.xml" in existing_paths:
                sheet_number += 1
            sheet_target = f"worksheets/sheet{sheet_number}.xml"
            sheet_path = f"xl/{sheet_target}"
            ET.SubElement(sheets, f"{TAG}sheet", {"name": SHEET_NAME, "sheetId": str(sheet_id), RID: rid})
            ET.SubElement(
                rels,
                f"{REL_TAG}Relationship",
                {
                    "Id": rid,
                    "Type": f"{OFFICE_REL_NS}/worksheet",
                    "Target": sheet_target,
                },
            )
        else:
            rid = existing.attrib[RID]
            rel = rels.find(f"rel:Relationship[@Id='{rid}']", NS)
            if rel is None:
                raise KeyError(f"Missing relationship for {rid}")
            target = rel.attrib["Target"]
            sheet_path = target if target.startswith("xl/") else f"xl/{target}"

        part_name = f"/{sheet_path}"
        if content_types.find(f"ct:Override[@PartName='{part_name}']", NS) is None:
            ET.SubElement(
                content_types,
                f"{CT_TAG}Override",
                {
                    "PartName": part_name,
                    "ContentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
                },
            )

        heading_count = app.find("app:HeadingPairs/vt:vector/vt:variant[2]/vt:i4", NS)
        titles = app.find("app:TitlesOfParts/vt:vector", NS)
        if heading_count is not None and titles is not None:
            title_values = [node.text for node in titles.findall("vt:lpstr", NS)]
            if SHEET_NAME not in title_values:
                heading_count.text = str(int(heading_count.text or "0") + 1)
                titles.attrib["size"] = str(int(titles.attrib["size"]) + 1)
                ET.SubElement(titles, f"{VT_TAG}lpstr").text = SHEET_NAME

        replacements = {
            "xl/workbook.xml": workbook_xml(workbook),
            "xl/_rels/workbook.xml.rels": ET.tostring(rels, encoding="utf-8", xml_declaration=True),
            "[Content_Types].xml": ET.tostring(content_types, encoding="utf-8", xml_declaration=True),
            "docProps/app.xml": ET.tostring(app, encoding="utf-8", xml_declaration=True),
            sheet_path: sheet_xml,
        }

        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False, dir=workbook_path.parent) as temp:
            temp_path = Path(temp.name)
        try:
            written = set()
            with zipfile.ZipFile(temp_path, "w", zipfile.ZIP_DEFLATED) as target:
                for item in source.infolist():
                    data = replacements.get(item.filename, source.read(item.filename))
                    target.writestr(item, data)
                    written.add(item.filename)
                if sheet_path not in written:
                    target.writestr(sheet_path, sheet_xml)
            shutil.move(temp_path, workbook_path)
        finally:
            temp_path.unlink(missing_ok=True)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("workbook", type=Path)
    parser.add_argument("csv", type=Path)
    parser.add_argument("--usage-profile", default="1.00_E1A_AZI_A")
    parser.add_argument("--injection-profile", default="1.00_E1A_AZI_I")
    args = parser.parse_args()
    add_or_replace_sheet(args.workbook, args.csv, args.usage_profile, args.injection_profile)


if __name__ == "__main__":
    main()
