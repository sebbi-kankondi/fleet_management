# How to run `python_financial_projection_model_logic.py` (step-by-step)

## 1) Open the project in VS Code
1. Open VS Code.
2. Go to **File -> Open Folder**.
3. Select the project root folder: `fleet_management`.

## 2) Open a terminal in VS Code
1. In VS Code, click **Terminal -> New Terminal**.
2. Confirm you are in the project root (you should see files like `data/` and `python_financial_projection_model.py`).

3. If not in the project root, change directory first (example):
   ```powershell
   cd C:\path\to\fleet_management

## 3) Create and activate a virtual environment

### Windows (PowerShell)
```powershell
python -m venv .venv
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
.venv\Scripts\Activate.ps1
```

If you created/activated the venv in **PowerShell**, keep using the **same PowerShell session** for the next commands.
There is **no need to switch to Git Bash** after activation.

## 4) Install dependencies
The script uses `openpyxl` for Excel read/write.

```powershell
pip install openpyxl
```

## 5) Validate script syntax (optional but recommended)
```powershell
python -m scripts/py_compile python_financial_projection_model.py
```

## 6) Run the model generator
Run with explicit input and output paths:

```powershell
python scripts/python_financial_projection_model.py `
  --input data/financial_projections.xlsx `
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

```powershell
python python_financial_projection_model.py --input data/financial_projections.xlsx
```

Example output filename format:
- `data/financial_projections_generated_YYYYMMDD_HHMMSS.xlsx`

## 9) Troubleshooting

### Error: `ModuleNotFoundError: No module named 'openpyxl'`
Install dependency:
```powershell
pip install openpyxl
```

### Error: input file not found
Check the path passed to `--input` and ensure the file exists.

### Output looks unchanged
Ensure you opened the generated output file (the `--output` path or timestamped file), not the original source file.
