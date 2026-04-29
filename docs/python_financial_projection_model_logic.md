# Python (VS Code) coding logic for Molenzicht CC financial projection generator

## 1) Recommended implementation approach
Build a **Python-based projection engine** that reads assumptions, recomputes all dependent schedules in order, and writes outputs into the existing workbook tabs:

1. `Assumptions`
2. `Fleet_Schedule`
3. `Loan_Amortisation`
4. `Income_Statement`
5. `Cash_Flow`
6. `Balance_Sheet`

This order respects the current dependency chain implied by the workbook layout.

---

## 2) Current workbook structure (confirmed)
Use these existing layouts as the base contract for your script:

- `Assumptions` has `Assumption`, `Value`, `Units`, `Notes / Source` columns.
- `Fleet_Schedule` tracks operating cars per month (`Cars in Rent-to-Own Operation` in column G).
- `Income_Statement` currently includes direct-cost lines (`Fuel`, `Airtime`, `Carwash`, `Maintenance`) and `Cost of Sales`.
- `Cash_Flow` has `Operating Expenses (cash)` in column E.
- `Balance_Sheet` references fleet counts, capex, and cumulative metrics.
- `Loan_Amortisation` is the monthly debt schedule.

---

## 3) New assumption requirements to include
Add/update these assumptions in `Assumptions` and use them in monthly calculations.

### 3.1 Required assumption values
1. **Driver Subsistence Allocation amount** = `1400` (monthly, per operating car) -> **Cost of Sales**.
2. **Cost of Sales calculation as an assumption** -> define a formula rule assumption (see section 5.2).
3. **Maintenance expense per active car (monthly)** = `1250`.
4. **Incidental Repair Reserve (monthly)** = `500` (per operating car) -> **Cost of Sales**.
5. **Airtime expense per operating car (monthly)** = `290` (reduced from 580).
6. **Tracking device expense (monthly)** = `950` (per operating car) -> **Cost of Sales**.

### 3.2 Suggested rows in `Assumptions`
Use the same 4-column pattern already present in the sheet:

- `Maintenance expense per active car (monthly)` -> `1250`
- `Airtime expense per operating car (monthly)` -> `290`
- `Driver subsistence allocation per operating car (monthly)` -> `1400`
- `Incidental repair reserve per operating car (monthly)` -> `500`
- `Tracking device expense per operating car (monthly)` -> `950`
- `Cost of sales per operating car (monthly)` -> **calculated assumption**

> Keep all direct-cost assumptions in `N$ / car / month` units for consistency.

---

## 4) Suggested Python project layout in VS Code

```text
fleet_management/
  data/
    financial_projections.xlsx
  src/
    config.py
    assumptions.py
    fleet.py
    loan.py
    income_statement.py
    cash_flow.py
    balance_sheet.py
    workbook_io.py
    main.py
  tests/
    test_assumptions.py
    test_income_statement.py
    test_balancing.py
```

Use a virtual environment and run from VS Code terminal:

```bash
python -m venv .venv
source .venv/bin/activate
pip install openpyxl pydantic pytest
```

---

## 5) Calculation logic (core)

## 5.1 Load assumptions into a typed model
Create a typed assumptions object (e.g., `pydantic` model or dataclass) with fields such as:

- `monthly_gross_revenue_per_car`
- `fuel_per_car`
- `airtime_per_car`
- `carwash_per_car`
- `maintenance_per_car`
- `driver_subsistence_per_car`
- `incidental_repair_reserve_per_car`
- `tracking_device_per_car`
- `salary_per_car`
- loan parameters, tax rate, batch sizes, etc.

## 5.2 Cost of sales as a calculated assumption
Define a reusable function/assumption rule:

```python
cost_of_sales_per_car = (
    fuel_per_car
    + airtime_per_car
    + carwash_per_car
    + maintenance_per_car
    + driver_subsistence_per_car
    + incidental_repair_reserve_per_car
    + tracking_device_per_car
)
```

Then, for each month:

```python
monthly_cost_of_sales = cars_in_operation * cost_of_sales_per_car
```

## 5.3 Income statement monthly logic
For each month `m` using `cars = fleet[m].cars_in_operation`:

- `gross_revenue = cars * monthly_gross_revenue_per_car`
- `fuel = cars * fuel_per_car`
- `airtime = cars * airtime_per_car` (use `290`)
- `carwash = cars * carwash_per_car`
- `maintenance = cars * maintenance_per_car` (use `1250`)
- `driver_subsistence = cars * driver_subsistence_per_car` (use `1400`)
- `incidental_repair_reserve = cars * incidental_repair_reserve_per_car` (use `500`)
- `tracking_device = cars * tracking_device_per_car` (use `950`)
- `cost_of_sales = sum(all direct costs above)`
- `gross_profit = gross_revenue - cost_of_sales`
- `salary = cars * salary_per_car`
- `total_operating_expenses = salary` (or include other opex lines if policy changes)
- `ebit = gross_profit - total_operating_expenses`

## 5.4 Sheet column updates (important)
The current income statement has:

`Year #, Month #, Cars in Operation, Monthly Gross Revenue, Fuel, Airtime, Carwash, Maintenance, Cost of Sales, Gross Profit, Salary, Total Operating Expenses`

To reflect your new requirements clearly, insert/add columns in `Income_Statement` for:

- `Driver Subsistence Allocation`
- `Incidental Repair Reserve`
- `Tracking Device`

Then compute `Cost of Sales` from all seven direct cost lines:

`Fuel + Airtime + Carwash + Maintenance + Driver Subsistence + Incidental Repair Reserve + Tracking Device`

## 5.5 Cash flow impact
Update `Cash_Flow` `Operating Expenses (cash)` to include:

- `Cost of Sales` (from income statement)
- `Salary`
- any additional cash opex policy items

At minimum, ensure new cost-of-sales components flow into `Operating Expenses (cash)` so cash and P&L stay aligned.

## 5.6 Balance sheet integrity checks
After monthly updates, validate:

- `assets == liabilities + equity`
- closing cash tie-out with cumulative net cash flow
- loan closing balance tie-out with amortisation table

---

## 6) Workbook write-back logic
Use `openpyxl` to:

1. Open template workbook.
2. Update values in `Assumptions` (including new rows).
3. Rebuild all monthly tables in dependency order.
4. Write full monthly ranges for each target sheet.
5. Save as a new versioned output file (e.g., timestamped) or overwrite by config.

Prefer writing data values from Python rather than relying on fragile in-sheet formulas for generated outputs.

---

## 7) Pseudocode orchestration

```python
def run_projection(input_xlsx: str, output_xlsx: str) -> None:
    wb = load_workbook(input_xlsx)

    assumptions = load_assumptions(wb["Assumptions"])
    assumptions = apply_required_overrides(
        assumptions,
        maintenance_per_car=1250,
        airtime_per_car=290,
        driver_subsistence_per_car=1400,
        incidental_repair_reserve_per_car=500,
        tracking_device_per_car=950,
    )
    assumptions.cost_of_sales_per_car = calc_cost_of_sales_per_car(assumptions)

    fleet_rows = build_fleet_schedule(assumptions)
    loan_rows = build_loan_amortisation(assumptions)
    income_rows = build_income_statement(assumptions, fleet_rows)
    cash_rows = build_cash_flow(assumptions, income_rows, loan_rows, fleet_rows)
    balance_rows = build_balance_sheet(assumptions, fleet_rows, cash_rows, loan_rows)

    write_assumptions(wb["Assumptions"], assumptions)
    write_fleet(wb["Fleet_Schedule"], fleet_rows)
    write_loan(wb["Loan_Amortisation"], loan_rows)
    write_income_statement(wb["Income_Statement"], income_rows)
    write_cash_flow(wb["Cash_Flow"], cash_rows)
    write_balance_sheet(wb["Balance_Sheet"], balance_rows)

    run_validation_checks(income_rows, cash_rows, balance_rows, loan_rows)
    wb.save(output_xlsx)
```

---

## 8) Testing checklist
1. **Assumption override test**: confirms maintenance=1250, airtime=290, driver subsistence=1400, incidental reserve=500, tracking=950.
2. **Cost of sales formula test**: verifies monthly CoS equals per-car CoS × cars in operation.
3. **Income-to-cash tie-out test**: ensures cost-of-sales additions are included in cash `Operating Expenses`.
4. **Balance test**: verifies accounting equation each month.
5. **Regression scenario test**: compare base vs changed assumptions and confirm outputs change dynamically.

---

## 9) Practical note on R vs Python
Since you asked for Python coding logic in VS Code, the above design is Python-native. If needed, the same dependency order and formulas can be ported to R, but Python remains better suited for robust Excel model generation, CI tests, and automation pipelines.
