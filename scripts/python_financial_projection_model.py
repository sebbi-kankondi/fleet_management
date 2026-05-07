#!/usr/bin/env python3
"""Create financial_projections_final from financial_projections.

The script copies the source workbook to a separate final workbook path, then
optionally applies user-provided assumption changes to the copied workbook. This
keeps ``data/financial_projections.xlsx`` as the template/source of truth while
letting you generate a changed ``financial_projections_final.xlsx`` without
opening Excel manually.

Examples:
    python scripts/python_financial_projection_model.py --list-assumptions
    python scripts/python_financial_projection_model.py --set "Fuel expense per operating car (monthly)=12000"
    python scripts/python_financial_projection_model.py --assumptions-file scenario.json
    python scripts/python_financial_projection_model.py --interactive
"""

from __future__ import annotations

import argparse
import json
import re
import tempfile
import xml.etree.ElementTree as ET
import zipfile
from dataclasses import dataclass
from pathlib import Path
from shutil import copyfile, move
from typing import Iterable


DEFAULT_INPUT = Path("data/financial_projections.xlsx")
DEFAULT_OUTPUT_NAME = "financial_projections_final.xlsx"
ASSUMPTIONS_SHEET_NAME = "Assumptions"

NS_MAIN_URI = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL_URI = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_OFFICE_REL_URI = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_CONTENT_TYPES_URI = "http://schemas.openxmlformats.org/package/2006/content-types"

NS_MAIN = {"a": NS_MAIN_URI}
NS_REL = {"r": NS_REL_URI}

ET.register_namespace("", NS_MAIN_URI)
ET.register_namespace("r", NS_OFFICE_REL_URI)
ET.register_namespace("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006")
ET.register_namespace("x15", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main")
ET.register_namespace("xr", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision")
ET.register_namespace("xr6", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision6")
ET.register_namespace("xr10", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision10")
ET.register_namespace("xr2", "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2")
ET.register_namespace("ct", NS_CONTENT_TYPES_URI)


@dataclass(frozen=True)
class AssumptionRow:
    """A discovered assumption row in the Assumptions worksheet."""

    label: str
    value_cell: str
    value: str | None
    units: str | None
    notes: str | None


@dataclass(frozen=True)
class WorkbookPartMap:
    """Workbook part locations needed for low-level xlsx editing."""

    assumptions_sheet_part: str
    workbook_part: str = "xl/workbook.xml"
    workbook_rels_part: str = "xl/_rels/workbook.xml.rels"
    content_types_part: str = "[Content_Types].xml"


def copy_financial_projection_workbook(input_path: Path, output_path: Path) -> None:
    """Copy all workbook formulas and non-formula values to the final file."""
    output_path.parent.mkdir(parents=True, exist_ok=True)
    copyfile(input_path, output_path)


def normalize_assumption_key(value: str) -> str:
    """Normalize an assumption label for forgiving CLI lookups."""
    return re.sub(r"[^a-z0-9]+", " ", value.lower()).strip()


def col_letters(cell_ref: str) -> str:
    """Return the column letters from a cell reference such as B12."""
    match = re.match(r"([A-Z]+)", cell_ref)
    if not match:
        raise ValueError(f"Invalid cell reference: {cell_ref}")
    return match.group(1)


def row_number(cell_ref: str) -> int:
    """Return the 1-based row number from a cell reference such as B12."""
    match = re.search(r"(\d+)$", cell_ref)
    if not match:
        raise ValueError(f"Invalid cell reference: {cell_ref}")
    return int(match.group(1))


def load_shared_strings(zf: zipfile.ZipFile) -> list[str]:
    """Return shared-string table values from an xlsx archive."""
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []

    shared_strings_xml = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    values: list[str] = []
    for string_item in shared_strings_xml.findall("a:si", NS_MAIN):
        text_fragments = [text_node.text or "" for text_node in string_item.findall(".//a:t", NS_MAIN)]
        values.append("".join(text_fragments))
    return values


def cell_text(cell: ET.Element | None, shared_strings: list[str]) -> str | None:
    """Read a displayable cell value from worksheet XML."""
    if cell is None:
        return None

    formula = cell.find("a:f", NS_MAIN)
    if formula is not None:
        return f"={formula.text or ''}"

    value = cell.find("a:v", NS_MAIN)
    if value is None:
        inline_text = cell.find("a:is/a:t", NS_MAIN)
        return inline_text.text if inline_text is not None else None

    if cell.attrib.get("t") == "s":
        return shared_strings[int(value.text or 0)]
    return value.text


def find_workbook_parts(zf: zipfile.ZipFile) -> WorkbookPartMap:
    """Resolve the Assumptions worksheet part from workbook metadata."""
    workbook_xml = ET.fromstring(zf.read("xl/workbook.xml"))
    rels_xml = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    rel_targets = {
        rel.attrib["Id"]: rel.attrib["Target"]
        for rel in rels_xml.findall("r:Relationship", NS_REL)
    }

    relationship_key = f"{{{NS_OFFICE_REL_URI}}}id"
    for sheet in workbook_xml.findall("a:sheets/a:sheet", NS_MAIN):
        if sheet.attrib.get("name") == ASSUMPTIONS_SHEET_NAME:
            relationship_id = sheet.attrib[relationship_key]
            target = rel_targets[relationship_id]
            target = target.lstrip("/")
            assumptions_part = target if target.startswith("xl/") else f"xl/{target}"
            return WorkbookPartMap(assumptions_sheet_part=assumptions_part)

    raise ValueError(f"Workbook does not contain a {ASSUMPTIONS_SHEET_NAME!r} sheet.")


def get_cell(row: ET.Element, cell_ref: str) -> ET.Element | None:
    """Return a cell by reference within an XML row."""
    for cell in row.findall("a:c", NS_MAIN):
        if cell.attrib.get("r") == cell_ref:
            return cell
    return None


def list_assumptions(workbook_path: Path) -> list[AssumptionRow]:
    """Read assumptions from the workbook's Assumptions sheet."""
    with zipfile.ZipFile(workbook_path) as zf:
        parts = find_workbook_parts(zf)
        shared_strings = load_shared_strings(zf)
        worksheet_xml = ET.fromstring(zf.read(parts.assumptions_sheet_part))

        assumptions: list[AssumptionRow] = []
        for row in worksheet_xml.findall("a:sheetData/a:row", NS_MAIN):
            row_index = int(row.attrib["r"])
            label = cell_text(get_cell(row, f"A{row_index}"), shared_strings)
            if not label or label in {"Assumption", "Namibia Rent-to-Own Fleet Model - Key Assumptions (Editable)"}:
                continue

            assumptions.append(
                AssumptionRow(
                    label=label,
                    value_cell=f"B{row_index}",
                    value=cell_text(get_cell(row, f"B{row_index}"), shared_strings),
                    units=cell_text(get_cell(row, f"C{row_index}"), shared_strings),
                    notes=cell_text(get_cell(row, f"D{row_index}"), shared_strings),
                )
            )
        return assumptions


def parse_assumption_value(raw_value: object) -> tuple[str, str | None]:
    """Convert user input to xlsx cell type and serialized value/formula text."""
    if isinstance(raw_value, bool):
        return "b", "1" if raw_value else "0"
    if isinstance(raw_value, int | float):
        return "n", str(raw_value)

    value = str(raw_value).strip()
    if not value:
        raise ValueError("Assumption values cannot be blank.")
    if value.startswith("="):
        return "formula", value[1:]

    try:
        float(value.replace(",", ""))
    except ValueError:
        return "str", value
    return "n", value.replace(",", "")


def set_cell_value(cell: ET.Element, raw_value: object) -> None:
    """Set a worksheet cell to a numeric, boolean, string, or formula value."""
    value_type, serialized_value = parse_assumption_value(raw_value)

    for child in list(cell):
        if child.tag in {f"{{{NS_MAIN_URI}}}f", f"{{{NS_MAIN_URI}}}v", f"{{{NS_MAIN_URI}}}is"}:
            cell.remove(child)

    cell.attrib.pop("t", None)
    if value_type == "formula":
        formula = ET.SubElement(cell, f"{{{NS_MAIN_URI}}}f")
        formula.text = serialized_value
        return

    if value_type == "str":
        cell.attrib["t"] = "inlineStr"
        inline_string = ET.SubElement(cell, f"{{{NS_MAIN_URI}}}is")
        text = ET.SubElement(inline_string, f"{{{NS_MAIN_URI}}}t")
        text.text = serialized_value
        return

    if value_type == "b":
        cell.attrib["t"] = "b"

    value = ET.SubElement(cell, f"{{{NS_MAIN_URI}}}v")
    value.text = serialized_value


def ensure_cell(row: ET.Element, cell_ref: str) -> ET.Element:
    """Return an existing cell or create it in roughly column order."""
    existing = get_cell(row, cell_ref)
    if existing is not None:
        return existing

    new_cell = ET.Element(f"{{{NS_MAIN_URI}}}c", {"r": cell_ref})
    target_col = col_letters(cell_ref)
    inserted = False
    for index, cell in enumerate(row.findall("a:c", NS_MAIN)):
        if col_letters(cell.attrib["r"]) > target_col:
            row.insert(index, new_cell)
            inserted = True
            break
    if not inserted:
        row.append(new_cell)
    return new_cell


def write_direct_assumption_updates(
    worksheet_xml: ET.Element,
    assumption_rows: list[AssumptionRow],
    updates: dict[str, object],
) -> dict[str, str]:
    """Apply explicit user-provided assumption updates to worksheet XML."""
    by_exact = {row.label: row for row in assumption_rows}
    by_normalized = {normalize_assumption_key(row.label): row for row in assumption_rows}
    applied: dict[str, str] = {}

    rows_by_index = {
        int(row.attrib["r"]): row for row in worksheet_xml.findall("a:sheetData/a:row", NS_MAIN)
    }

    for key, value in updates.items():
        target_row = None
        if re.fullmatch(r"B\d+", key.strip(), flags=re.IGNORECASE):
            cell_ref = key.strip().upper()
            row_index = row_number(cell_ref)
            matching_label = next(
                (assumption.label for assumption in assumption_rows if assumption.value_cell == cell_ref),
                cell_ref,
            )
            target_row = AssumptionRow(matching_label, cell_ref, None, None, None)
        elif key in by_exact:
            target_row = by_exact[key]
        else:
            normalized_key = normalize_assumption_key(key)
            matches = [row for normalized, row in by_normalized.items() if normalized == normalized_key]
            if len(matches) != 1:
                available = ", ".join(row.label for row in assumption_rows)
                raise KeyError(f"Unknown assumption {key!r}. Available assumptions: {available}")
            target_row = matches[0]

        xml_row = rows_by_index[row_number(target_row.value_cell)]
        set_cell_value(ensure_cell(xml_row, target_row.value_cell), value)
        applied[target_row.label] = str(value)

    return applied


def read_numeric_assumption_values(worksheet_xml: ET.Element, assumption_rows: list[AssumptionRow]) -> dict[str, float]:
    """Read numeric assumption values after XML edits."""
    values: dict[str, float] = {}
    rows_by_index = {
        int(row.attrib["r"]): row for row in worksheet_xml.findall("a:sheetData/a:row", NS_MAIN)
    }
    for assumption in assumption_rows:
        cell = get_cell(rows_by_index[row_number(assumption.value_cell)], assumption.value_cell)
        if cell is None:
            continue
        raw = cell_text(cell, [])
        if raw is None or raw.startswith("="):
            continue
        try:
            values[assumption.label] = float(raw.replace(",", ""))
        except ValueError:
            continue
    return values


def update_calculated_assumption_cells(worksheet_xml: ET.Element, assumption_rows: list[AssumptionRow]) -> dict[str, float]:
    """Refresh known calculated assumption rows for immediate final-file display."""
    numeric_values = read_numeric_assumption_values(worksheet_xml, assumption_rows)
    calculations = {
        "Cost of sales": [
            "Fuel expense per operating car (monthly)",
            "Airtime expense per operating car (monthly)",
            "Carwash expense per operating car (monthly)",
            "Maintenance expense per active car (monthly)",
            "Driver subsistence",
            "Incidental repair reserve",
            "Tracking device expense",
        ],
    }

    calculated_updates: dict[str, float] = {}
    if all(label in numeric_values for label in calculations["Cost of sales"]):
        calculated_updates["Cost of sales"] = sum(numeric_values[label] for label in calculations["Cost of sales"])
    if "Monthly Gross revenue per car" in numeric_values and "Cost of sales" in calculated_updates:
        calculated_updates["Monthly Gross profit per car"] = (
            numeric_values["Monthly Gross revenue per car"] - calculated_updates["Cost of sales"]
        )
    if "Salary expense per operating car (monthly)" in numeric_values and "Cost of sales" in calculated_updates:
        calculated_updates["Total operating expenses per operating car (monthly)"] = (
            calculated_updates["Cost of sales"] + numeric_values["Salary expense per operating car (monthly)"]
        )
    if (
        "Monthly Gross profit per car" in calculated_updates
        and "Salary expense per operating car (monthly)" in numeric_values
    ):
        calculated_updates["Operating Profit"] = (
            calculated_updates["Monthly Gross profit per car"]
            - numeric_values["Salary expense per operating car (monthly)"]
        )

    if calculated_updates:
        write_direct_assumption_updates(worksheet_xml, assumption_rows, calculated_updates)
    return calculated_updates


def force_full_workbook_recalculation(workbook_xml: ET.Element) -> None:
    """Ask Excel-compatible applications to recalculate formulas when opened."""
    calc_pr = workbook_xml.find("a:calcPr", NS_MAIN)
    if calc_pr is None:
        calc_pr = ET.SubElement(workbook_xml, f"{{{NS_MAIN_URI}}}calcPr")
    calc_pr.attrib.update({"calcMode": "auto", "fullCalcOnLoad": "1", "forceFullCalc": "1"})


def remove_calc_chain_references(part_name: str, xml_bytes: bytes) -> bytes:
    """Remove stale calcChain references when a workbook is edited outside Excel."""
    if part_name == "xl/_rels/workbook.xml.rels":
        rels_xml = ET.fromstring(xml_bytes)
        for rel in list(rels_xml.findall("r:Relationship", NS_REL)):
            if rel.attrib.get("Type", "").endswith("/calcChain"):
                rels_xml.remove(rel)
        return ET.tostring(rels_xml, encoding="utf-8", xml_declaration=True)

    if part_name == "[Content_Types].xml":
        content_types_xml = ET.fromstring(xml_bytes)
        calc_chain_part_name = "/xl/calcChain.xml"
        for override in list(content_types_xml):
            if override.attrib.get("PartName") == calc_chain_part_name:
                content_types_xml.remove(override)
        return ET.tostring(content_types_xml, encoding="utf-8", xml_declaration=True)

    return xml_bytes


def apply_assumption_updates(workbook_path: Path, updates: dict[str, object]) -> dict[str, str]:
    """Apply assumption updates in-place to an xlsx workbook copy."""
    if not updates:
        return {}

    assumption_rows = list_assumptions(workbook_path)
    applied: dict[str, str] = {}
    calculated: dict[str, float] = {}

    with zipfile.ZipFile(workbook_path, "r") as source_zip:
        parts = find_workbook_parts(source_zip)
        worksheet_xml = ET.fromstring(source_zip.read(parts.assumptions_sheet_part))
        applied = write_direct_assumption_updates(worksheet_xml, assumption_rows, updates)
        calculated = update_calculated_assumption_cells(worksheet_xml, assumption_rows)

        workbook_xml = ET.fromstring(source_zip.read(parts.workbook_part))
        force_full_workbook_recalculation(workbook_xml)

        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx", dir=workbook_path.parent) as temp_file:
            temp_path = Path(temp_file.name)

        try:
            with zipfile.ZipFile(temp_path, "w", compression=zipfile.ZIP_DEFLATED) as target_zip:
                for item in source_zip.infolist():
                    if item.filename == "xl/calcChain.xml":
                        continue
                    if item.filename == parts.assumptions_sheet_part:
                        data = ET.tostring(worksheet_xml, encoding="utf-8", xml_declaration=True)
                    elif item.filename == parts.workbook_part:
                        data = ET.tostring(workbook_xml, encoding="utf-8", xml_declaration=True)
                    else:
                        data = remove_calc_chain_references(item.filename, source_zip.read(item.filename))
                    target_zip.writestr(item, data)
            move(temp_path, workbook_path)
        finally:
            if temp_path.exists():
                temp_path.unlink()

    return applied | {label: str(value) for label, value in calculated.items() if label not in applied}


def parse_set_arguments(set_arguments: Iterable[str]) -> dict[str, str]:
    """Parse repeated --set KEY=VALUE arguments."""
    updates: dict[str, str] = {}
    for item in set_arguments:
        if "=" not in item:
            raise ValueError(f"Invalid --set value {item!r}. Use --set 'Assumption name=123'.")
        key, value = item.split("=", 1)
        key = key.strip()
        if not key:
            raise ValueError(f"Invalid --set value {item!r}. Assumption name cannot be blank.")
        updates[key] = value.strip()
    return updates


def load_assumption_file(path: Path) -> dict[str, object]:
    """Load assumption overrides from a JSON file."""
    with path.open("r", encoding="utf-8") as file_handle:
        data = json.load(file_handle)
    if not isinstance(data, dict):
        raise ValueError("Assumptions file must contain a JSON object of assumption names to values.")
    return data


def prompt_for_assumptions(workbook_path: Path) -> dict[str, str]:
    """Collect assumption overrides from standard input."""
    print("Available assumptions:")
    for assumption in list_assumptions(workbook_path):
        units = f" ({assumption.units})" if assumption.units else ""
        print(f"- {assumption.label}{units}: current value = {assumption.value}")

    print("\nEnter changes as Assumption name=value. Press Enter on a blank line when done.")
    updates: dict[str, str] = {}
    while True:
        try:
            line = input("> ").strip()
        except EOFError:
            break
        if not line:
            break
        if "=" not in line:
            print("Please use Assumption name=value format.")
            continue
        key, value = line.split("=", 1)
        updates[key.strip()] = value.strip()
    return updates


def print_assumption_list(workbook_path: Path) -> None:
    """Print editable assumptions for CLI users."""
    for assumption in list_assumptions(workbook_path):
        units = f" [{assumption.units}]" if assumption.units else ""
        print(f"{assumption.label} ({assumption.value_cell}) = {assumption.value}{units}")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Generate financial_projections_final by copying all formulas and "
            "all non-formula values from financial_projections, with an optional "
            "input layer for assumption overrides."
        )
    )
    parser.add_argument(
        "--input",
        default=str(DEFAULT_INPUT),
        help="Path to the source financial_projections workbook.",
    )
    parser.add_argument(
        "--output",
        help="Path to the output financial_projections_final workbook.",
    )
    parser.add_argument(
        "--set",
        dest="set_values",
        action="append",
        default=[],
        metavar="ASSUMPTION=VALUE",
        help=(
            "Override one assumption in the generated final workbook. Repeat for "
            "multiple changes. Use the exact assumption name from --list-assumptions "
            "or a value cell such as B8."
        ),
    )
    parser.add_argument(
        "--assumptions-file",
        help="JSON file containing assumption-name-to-value overrides.",
    )
    parser.add_argument(
        "--interactive",
        action="store_true",
        help="Prompt for assumption changes before generating the final workbook.",
    )
    parser.add_argument(
        "--list-assumptions",
        action="store_true",
        help="List editable assumptions from the input workbook and exit.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    input_path = Path(args.input)
    if not input_path.exists():
        raise FileNotFoundError(f"Input workbook not found: {input_path}")

    if args.list_assumptions:
        print_assumption_list(input_path)
        return

    updates: dict[str, object] = {}
    if args.assumptions_file:
        updates.update(load_assumption_file(Path(args.assumptions_file)))
    updates.update(parse_set_arguments(args.set_values))
    if args.interactive:
        updates.update(prompt_for_assumptions(input_path))

    output_path = Path(args.output) if args.output else input_path.with_name(DEFAULT_OUTPUT_NAME)
    copy_financial_projection_workbook(input_path=input_path, output_path=output_path)
    applied_updates = apply_assumption_updates(output_path, updates)

    print(f"Financial projections final workbook generated: {output_path}")
    if applied_updates:
        print("Applied assumption updates:")
        for label, value in applied_updates.items():
            print(f"- {label}: {value}")


if __name__ == "__main__":
    main()
