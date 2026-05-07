# How to run the financial projection workbook generator

This guide shows how to generate `data/financial_projections_final.xlsx` from
`data/financial_projections.xlsx` and how to change assumptions without opening
the source workbook manually.

## 1) Open the project in VS Code
1. Open VS Code.
2. Go to **File -> Open Folder**.
3. Select the project root folder: `fleet_management`.

## 2) Open PowerShell or a terminal
Change to the project root first. Example:

```powershell
cd C:\path\to\fleet_management
```

Confirm you are in the project root:

```powershell
pwd
dir
```

You should see folders such as `data/`, `docs/`, and `scripts/`.

## 3) Validate script syntax (optional but recommended)

The current script uses only Python's standard library, so you do not need to
install `openpyxl` for the assumption input layer.

```powershell
python -m py_compile scripts/python_financial_projection_model.py
```

## 4) See the editable assumptions

Run this command to list every assumption name, value cell, current value, and
unit from the `Assumptions` sheet:

```powershell
python scripts/python_financial_projection_model.py --list-assumptions
```

Use these names with `--set` or in a JSON assumptions file.

## 5) Generate the final workbook with no assumption changes

```powershell
python scripts/python_financial_projection_model.py `
  --input data/financial_projections.xlsx `
  --output data/financial_projections_final.xlsx
```

This copies all formulas and non-formula values from the source workbook into the
final workbook.

## 6) Generate the final workbook with command-line assumption changes

Use one or more `--set "Assumption name=value"` arguments. The source workbook
is not edited; only the generated final workbook receives these changes.

```powershell
python scripts/python_financial_projection_model.py `
  --input data/financial_projections.xlsx `
  --output data/financial_projections_final.xlsx `
  --set "Fuel expense per operating car (monthly)=12000" `
  --set "Monthly Gross revenue per car=32000"
```

You can also use a value cell from the `Assumptions` sheet, such as `B8`, if you
prefer:

```powershell
python scripts/python_financial_projection_model.py `
  --input data/financial_projections.xlsx `
  --output data/financial_projections_final.xlsx `
  --set "B8=12000"
```

## 7) Generate the final workbook from a JSON assumptions file

Create a JSON file such as `scenario.json`:

```json
{
  "Fuel expense per operating car (monthly)": 12000,
  "Monthly Gross revenue per car": 32000,
  "Bank nominal annual interest rate": 0.11
}
```

Then run:

```powershell
python scripts/python_financial_projection_model.py `
  --input data/financial_projections.xlsx `
  --output data/financial_projections_final.xlsx `
  --assumptions-file scenario.json
```

## 8) Use the interactive input layer

If you prefer to type changes one by one, run:

```powershell
python scripts/python_financial_projection_model.py `
  --input data/financial_projections.xlsx `
  --output data/financial_projections_final.xlsx `
  --interactive
```

The script will list the assumptions and prompt for entries in this format:

```text
Assumption name=value
```

Press **Enter** on a blank line when you are done.

## 9) What the script updates

When you provide input changes, the script:

- copies `financial_projections.xlsx` to `financial_projections_final.xlsx`;
- writes your changed values into the `Assumptions` sheet in the final workbook;
- refreshes the visible calculated assumption totals for:
  - `Cost of sales`;
  - `Monthly Gross profit per car`;
  - `Total operating expenses per operating car (monthly)`;
  - `Operating Profit`;
- marks the workbook for full formula recalculation when opened in Excel or a
  compatible spreadsheet application.

The original `data/financial_projections.xlsx` remains unchanged.

## 10) Troubleshooting

### Error: input file not found
Check the path passed to `--input` and ensure the file exists.

### Error: unknown assumption
Run `--list-assumptions` and copy the exact assumption name, or use the value
cell shown next to the assumption, such as `B8`.

### Output looks unchanged
Ensure you opened the generated output file (`data/financial_projections_final.xlsx`),
not the original source workbook (`data/financial_projections.xlsx`). If formulas
look stale in a non-Excel viewer, open the file in Excel or a compatible
spreadsheet application that honors full-workbook recalculation on load.
