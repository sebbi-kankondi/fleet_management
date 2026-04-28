#!/usr/bin/env python3
"""Quick inspection for workbook sheet/cell/formula counts."""

from __future__ import annotations

import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

NS_MAIN = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
NS_REL = {"r": "http://schemas.openxmlformats.org/package/2006/relationships"}


def inspect_workbook(path: Path) -> list[tuple[str, int, int]]:
    with zipfile.ZipFile(path) as zf:
        workbook = ET.fromstring(zf.read("xl/workbook.xml"))
        sheets = workbook.findall("a:sheets/a:sheet", NS_MAIN)

        rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
        rid_to_target = {
            rel.attrib["Id"]: rel.attrib["Target"] for rel in rels.findall("r:Relationship", NS_REL)
        }

        rows: list[tuple[str, int, int]] = []
        for sheet in sheets:
            name = sheet.attrib["name"]
            rid = sheet.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]
            target = rid_to_target[rid]
            xml = ET.fromstring(zf.read(f"xl/{target}"))
            cell_count = len(xml.findall(".//a:c", NS_MAIN))
            formula_count = len(xml.findall(".//a:f", NS_MAIN))
            rows.append((name, cell_count, formula_count))

        return rows


def main() -> None:
    workbook_path = Path("data/financial_projections.xlsx")
    for sheet, cells, formulas in inspect_workbook(workbook_path):
        print(f"{sheet}: cells={cells}, formulas={formulas}")


if __name__ == "__main__":
    main()
