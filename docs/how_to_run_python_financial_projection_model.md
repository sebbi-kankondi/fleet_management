# How to run `python_financial_projection_model_logic.py` (step-by-step)

## 1) Open the project in VS Code
1. Open VS Code.
2. Go to **File -> Open Folder**.
3. Select the project root folder: `fleet_management`.

## 2) Open a terminal in VS Code
1. In VS Code, click **Terminal -> New Terminal**.
2. Confirm you are in the project root (you should see files like `data/` and `python_financial_projection_model_logic.py`).

## 3) Create and activate a virtual environment

### Linux / macOS
```bash
python -m venv .venv
source .venv/bin/activate
```

### Windows (PowerShell)
```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
```

## 4) Install dependencies
The script uses `openpyxl` for Excel read/write.

```bash
pip install openpyxl
```

## 5) Validate script syntax (optional but recommended)
```bash
python -m py_compile python_financial_projection_model_logic.py
```

## 6) Run the model generator
Run with explicit input and output paths:

```bash
python python_financial_projection_model_logic.py \
  --input data/financial_projections.xlsx \
  --output data/financial_projections_generated.xlsx
```

If successful, it prints a completion message with the output file path.

## 7) What the script updates
The script updates/creates required assumptions and regenerates these sheets dynamically:
- `Income_Statement`
- `Cash_Flow`
- `Loan_Amortisation`
- `Balance_Sheet`

It reads fleet drivers from:
- `Fleet_Schedule`

And it enforces requested assumptions:
- Driver subsistence allocation = 1400 (cost of sales)
- Maintenance = 1250 (monthly per active car)
- Incidental repair reserve = 500 (cost of sales)
- Airtime = 290 (monthly per operating car)
- Tracking device expense = 950 (cost of sales)
- Cost of sales as a calculated assumption

## 8) Run with auto-generated output filename
If you omit `--output`, the script creates a timestamped output in `data/`:

```bash
python python_financial_projection_model_logic.py --input data/financial_projections.xlsx
```

Example output filename format:
- `data/financial_projections_generated_YYYYMMDD_HHMMSS.xlsx`

## 9) Troubleshooting

### Error: `ModuleNotFoundError: No module named 'openpyxl'`
Install dependency:
```bash
pip install openpyxl
```

### Error: input file not found
Check the path passed to `--input` and ensure the file exists.

### Output looks unchanged
Ensure you opened the generated output file (the `--output` path or timestamped file), not the original source file.
