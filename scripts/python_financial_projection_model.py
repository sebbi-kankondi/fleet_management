#!/usr/bin/env python3
"""Create financial_projections_final from financial_projections.

The script copies the source workbook to a separate final workbook path. Because
an .xlsx file stores formulas and static cell values in the workbook package,
this preserves every formula cell as a formula and every non-formula cell value
without modifying the source financial_projections file.
"""

import argparse
from pathlib import Path
from shutil import copyfile


DEFAULT_INPUT = Path("data/financial_projections.xlsx")
DEFAULT_OUTPUT_NAME = "financial_projections_final.xlsx"


def copy_financial_projection_workbook(input_path: Path, output_path: Path) -> None:
    """Copy all workbook formulas and non-formula values to the final file."""
    output_path.parent.mkdir(parents=True, exist_ok=True)
    copyfile(input_path, output_path)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Generate financial_projections_final by copying all formulas and "
            "all non-formula values from financial_projections."
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
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    input_path = Path(args.input)
    if not input_path.exists():
        raise FileNotFoundError(f"Input workbook not found: {input_path}")

    output_path = Path(args.output) if args.output else input_path.with_name(DEFAULT_OUTPUT_NAME)
    copy_financial_projection_workbook(input_path=input_path, output_path=output_path)
    print(f"Financial projections final workbook generated: {output_path}")


if __name__ == "__main__":
    main()
