#!/usr/bin/env python3
# Define the script interpreter so the file can be executed directly on Unix-like systems.

"""Generate dynamic financial projection sheets from the Assumptions sheet.

This implementation follows the logic documented in docs/python_financial_projection_model_logic.md.
"""

# Import argparse to parse command-line flags such as input and output file paths.
import argparse
# Import dataclass so assumptions and row records can be represented with typed structures.
from dataclasses import dataclass
# Import datetime to create default output filenames with timestamps.
from datetime import datetime
# Import pathlib.Path for safe, cross-platform filesystem path handling.
from pathlib import Path
# Import typing helpers to make function contracts explicit and easier to maintain.
from typing import Dict, List

# Import openpyxl workbook loader; this is the core Excel IO dependency for reading/writing sheets.
from openpyxl import load_workbook


# Map assumptions by exact label text used in the workbook Assumptions column A.
ASSUMPTION_KEYS = {
    # Revenue model assumptions.
    "monthly_gross_revenue_per_car": "Monthly Gross revenue per car",
    # Direct-cost assumptions per operating car.
    "fuel_per_car": "Fuel expense per operating car (monthly)",
    "airtime_per_car": "Airtime expense per operating car (monthly)",
    "carwash_per_car": "Carwash expense per operating car (monthly)",
    "maintenance_per_car": "Maintenance expense per active car (monthly)",
    "salary_per_car": "Salary expense per operating car (monthly)",
    # Corporate assumptions.
    "tax_rate": "Corporate income tax rate",
    "car_unit_cost": "Car unit cost",
    # Financing assumptions.
    "owner_equity": "Initial owner/investor equity",
    "bank_draw": "Bank loan draw (batch 2)",
    "bank_draw_month": "Bank loan draw month (operating month index)",
    "bank_instalment": "Bank monthly instalment (principal + interest)",
    "bank_annual_interest": "Bank nominal annual interest rate",
}



REQUIRED_ASSUMPTION_LABELS = {
    ASSUMPTION_KEYS["monthly_gross_revenue_per_car"],
    ASSUMPTION_KEYS["fuel_per_car"],
    ASSUMPTION_KEYS["airtime_per_car"],
    ASSUMPTION_KEYS["carwash_per_car"],
    "Maintenance expense per active car (monthly)",
    "Driver subsistence allocation per operating car (monthly)",
    "Incidental repair reserve per operating car (monthly)",
    "Tracking device expense per operating car (monthly)",
    ASSUMPTION_KEYS["salary_per_car"],
    ASSUMPTION_KEYS["tax_rate"],
    ASSUMPTION_KEYS["car_unit_cost"],
    ASSUMPTION_KEYS["owner_equity"],
    ASSUMPTION_KEYS["bank_draw"],
    ASSUMPTION_KEYS["bank_draw_month"],
    ASSUMPTION_KEYS["bank_instalment"],
    ASSUMPTION_KEYS["bank_annual_interest"],
}

# Define a typed assumptions object so all downstream calculations consume consistent fields.
@dataclass
class Assumptions:
    # Monthly revenue per car in N$.
    monthly_gross_revenue_per_car: float
    # Monthly direct costs per car in N$.
    fuel_per_car: float
    airtime_per_car: float
    carwash_per_car: float
    maintenance_per_car: float
    driver_subsistence_per_car: float
    incidental_repair_reserve_per_car: float
    tracking_device_per_car: float
    # Monthly salary per car in N$.
    salary_per_car: float
    # Tax and capex assumptions.
    tax_rate: float
    car_unit_cost: float
    # Financing assumptions.
    owner_equity: float
    bank_draw: float
    bank_draw_month: int
    bank_instalment: float
    bank_annual_interest: float


# Define a typed record for fleet rows to avoid brittle index-based dictionary access.
@dataclass
class FleetRow:
    # Operating month index.
    month: int
    # Fiscal year index.
    year: int
    # Number of cars purchased in the month.
    cars_purchased: float
    # Number of cars currently operating in the month.
    cars_in_operation: float


# Define a typed row for the income statement sheet.
@dataclass
class IncomeRow:
    # Time dimensions.
    year: int
    month: int
    cars_in_operation: float
    # Revenue and direct cost lines.
    monthly_gross_revenue: float
    fuel: float
    airtime: float
    carwash: float
    maintenance: float
    driver_subsistence: float
    incidental_repair_reserve: float
    tracking_device: float
    cost_of_sales: float
    gross_profit: float
    salary: float
    total_operating_expenses: float
    ebit: float
    income_tax: float
    net_profit: float


# Define a typed row for loan amortisation results.
@dataclass
class LoanRow:
    # Loan month counter starting at 1 for first payment month.
    loan_month: int
    # Opening loan balance for the month.
    opening_balance: float
    # Interest portion of payment.
    interest: float
    # Principal portion of payment.
    principal: float
    # Total payment amount.
    payment: float
    # Closing balance after payment.
    closing_balance: float


# Define a typed row for monthly cash flow results.
@dataclass
class CashFlowRow:
    # Time dimension.
    month: int
    # Financing and operating cash line items.
    owner_injection: float
    bank_draw: float
    total_cash_in: float
    operating_expenses_cash: float
    interest: float
    loan_principal: float
    income_tax_paid: float
    investor_payout: float
    net_cash_before_capex: float
    capex: float
    net_cash_flow: float


# Define a typed row for balance sheet outputs.
@dataclass
class BalanceRow:
    # Time dimensions.
    year: int
    month: int
    # Equity/debt/cash indicators.
    total_owner_capital_invested: float
    net_owner_capital_outstanding: float
    loan_closing_balance: float
    cars_in_operation: float
    total_cars_operation_cumulative: float
    vehicles_in_use_cost: float
    total_cars_cost_cumulative: float
    net_cash_before_capex: float
    capex: float
    gross_revenue_cumulative: float


# Small numeric helper to keep workbook output readable and consistent.
def r2(value: float) -> float:
    # Round to two decimals for currency outputs.
    return round(value, 2)


# Locate an assumption row by label in column A and return its row number if found.
def find_assumption_row(assumptions_ws, label: str):
    # Iterate through all used rows in the assumptions sheet.
    for row in range(1, assumptions_ws.max_row + 1):
        # Read the label from column A.
        cell_value = assumptions_ws.cell(row=row, column=1).value
        # If the label matches exactly, return this row index.
        if cell_value == label:
            return row
    # Return None when the label is not present.
    return None


# Ensure required assumptions exist and force the new requested values.
def ensure_required_assumptions(assumptions_ws):
    # Define required overrides/additions based on user requirements.
    required = {
        "Maintenance expense per active car (monthly)": (1250, "N$ / car / month", "Updated requirement"),
        "Airtime expense per operating car (monthly)": (290, "N$ / car / month", "Reduced requirement"),
        "Driver subsistence allocation per operating car (monthly)": (1400, "N$ / car / month", "Cost of sales component"),
        "Incidental repair reserve per operating car (monthly)": (500, "N$ / car / month", "Cost of sales component"),
        "Tracking device expense per operating car (monthly)": (950, "N$ / car / month", "Cost of sales component"),
    }

    # Process each required assumption one-by-one.
    for label, (value, units, note) in required.items():
        # Attempt to find existing row.
        row = find_assumption_row(assumptions_ws, label)
        # If row exists, update only value/units/note fields.
        if row is not None:
            assumptions_ws.cell(row=row, column=2, value=value)
            assumptions_ws.cell(row=row, column=3, value=units)
            assumptions_ws.cell(row=row, column=4, value=note)
        else:
            # If row does not exist, append at end of sheet.
            new_row = assumptions_ws.max_row + 1
            assumptions_ws.cell(row=new_row, column=1, value=label)
            assumptions_ws.cell(row=new_row, column=2, value=value)
            assumptions_ws.cell(row=new_row, column=3, value=units)
            assumptions_ws.cell(row=new_row, column=4, value=note)

    # Build or refresh calculated cost-of-sales-per-car assumption.
    cos_label = "Cost of sales per operating car (monthly)"
    # Compute the assumed per-car cost of sales from direct-cost components.
    cost_of_sales_per_car = 13000 + 290 + 240 + 1250 + 1400 + 500 + 950
    # Find existing row for this calculated assumption.
    cos_row = find_assumption_row(assumptions_ws, cos_label)
    # Update existing row or append a new row.
    target_row = cos_row if cos_row is not None else assumptions_ws.max_row + 1
    # Write calculated assumption fields.
    assumptions_ws.cell(row=target_row, column=1, value=cos_label)
    assumptions_ws.cell(row=target_row, column=2, value=cost_of_sales_per_car)
    assumptions_ws.cell(row=target_row, column=3, value="N$ / car / month")
    assumptions_ws.cell(row=target_row, column=4, value="Calculated = Fuel + Airtime + Carwash + Maintenance + Driver Subsistence + Incidental Reserve + Tracking")


# Read assumptions from sheet into dictionary keyed by label.
def read_assumption_values(assumptions_ws, assumptions_values_ws=None) -> Dict[str, float]:
    # Parse assumption numeric values from numbers or formatted strings (currency/percent).
    def parse_assumption_numeric(value) -> float:
        # Fast-path for native numeric cells.
        if isinstance(value, (int, float)):
            return float(value)
        # Reject unsupported empty values.
        if value is None:
            raise ValueError("Assumption value is empty.")

        # Normalize text values (e.g. '15%', 'N$ 12,500').
        text_value = str(value).strip()
        if text_value == "":
            raise ValueError("Assumption value is blank.")

        is_percent = "%" in text_value
        normalized = (
            text_value.replace("N$", "")
            .replace("$", "")
            .replace("%", "")
            .strip()
        )

        # Validate comma usage so decimal commas fail fast instead of being misparsed.
        if "," in normalized:
            unsigned = normalized.lstrip("+-")
            integer_part, dot, _fractional_part = unsigned.partition(".")
            comma_groups = integer_part.split(",")
            # Allow commas only as thousands separators in the integer part.
            if len(comma_groups) == 1:
                raise ValueError("Ambiguous comma in assumption value; use '.' for decimals.")
            if any(group == "" for group in comma_groups):
                raise ValueError("Invalid comma placement in assumption value.")
            if not (1 <= len(comma_groups[0]) <= 3 and all(len(group) == 3 for group in comma_groups[1:])):
                raise ValueError("Invalid comma placement in assumption value.")
            if dot and "," in _fractional_part:
                raise ValueError("Invalid comma placement in assumption value.")

        cleaned = normalized.replace(",", "")

        # Parse cleaned numeric text.
        numeric_value = float(cleaned)
        return numeric_value / 100.0 if is_percent else numeric_value

    # Create an empty dictionary for values.
    values = {}
    # Iterate all rows in assumptions sheet.
    for row in range(1, assumptions_ws.max_row + 1):
        # Read label and numeric value cells.
        label = assumptions_ws.cell(row=row, column=1).value
        value = assumptions_ws.cell(row=row, column=2).value
        # Prefer cached formula result from a data_only workbook when available.
        if isinstance(value, str) and value.strip().startswith("=") and assumptions_values_ws is not None:
            cached_value = assumptions_values_ws.cell(row=row, column=2).value
            if cached_value is not None:
                value = cached_value
        # Skip rows with missing labels or values.
        if label is None or value is None:
            continue
        label_str = str(label).strip()

        # Skip header rows (e.g. "Assumption" / "Value") if present in exported templates.
        label_text = label_str.lower()
        value_text = str(value).strip().lower()
        if label_text in {"assumption", "label"} and value_text in {"value", "amount"}:
            continue

        # Read only labels the model actually consumes; ignore decorative or formula helper rows.
        if label_str not in REQUIRED_ASSUMPTION_LABELS:
            continue

        # Store normalized numeric values by label text.
        try:
            values[label_str] = parse_assumption_numeric(value)
        except ValueError as exc:
            raise ValueError(f"Could not parse assumption value for label '{label}' at row {row}: {value!r}") from exc
    # Return mapping for downstream field extraction.
    return values


# Convert label-keyed assumption map into typed Assumptions object with overrides.
def build_assumptions(assumption_values: Dict[str, float]) -> Assumptions:
    # Define helper function to fetch required labels or fail fast with clear message.
    def get(label: str) -> float:
        # Raise a descriptive error if an expected assumption is missing.
        if label not in assumption_values:
            raise KeyError(f"Missing required assumption: {label}")
        # Return mapped numeric value.
        return assumption_values[label]

    # Create and return a fully populated assumptions dataclass.
    return Assumptions(
        monthly_gross_revenue_per_car=get(ASSUMPTION_KEYS["monthly_gross_revenue_per_car"]),
        fuel_per_car=get(ASSUMPTION_KEYS["fuel_per_car"]),
        airtime_per_car=get(ASSUMPTION_KEYS["airtime_per_car"]),
        carwash_per_car=get(ASSUMPTION_KEYS["carwash_per_car"]),
        maintenance_per_car=get("Maintenance expense per active car (monthly)"),
        driver_subsistence_per_car=get("Driver subsistence allocation per operating car (monthly)"),
        incidental_repair_reserve_per_car=get("Incidental repair reserve per operating car (monthly)"),
        tracking_device_per_car=get("Tracking device expense per operating car (monthly)"),
        salary_per_car=get(ASSUMPTION_KEYS["salary_per_car"]),
        tax_rate=get(ASSUMPTION_KEYS["tax_rate"]),
        car_unit_cost=get(ASSUMPTION_KEYS["car_unit_cost"]),
        owner_equity=get(ASSUMPTION_KEYS["owner_equity"]),
        bank_draw=get(ASSUMPTION_KEYS["bank_draw"]),
        bank_draw_month=int(get(ASSUMPTION_KEYS["bank_draw_month"])),
        bank_instalment=get(ASSUMPTION_KEYS["bank_instalment"]),
        bank_annual_interest=get(ASSUMPTION_KEYS["bank_annual_interest"]),
    )


# Read fleet schedule inputs from Fleet_Schedule sheet.
def read_fleet_schedule(fleet_ws, fleet_values_ws=None) -> List[FleetRow]:
    # Resolve formula text using cached values from optional data-only workbook.
    def resolve_cell_value(value, *, row_number: int, column_number: int, column_name: str):
        resolved_value = value
        if isinstance(resolved_value, str) and resolved_value.strip().startswith("="):
            if fleet_values_ws is not None:
                cached = fleet_values_ws.cell(row=row_number, column=column_number).value
                if cached is not None:
                    resolved_value = cached
            if isinstance(resolved_value, str) and resolved_value.strip().startswith("="):
                raise ValueError(
                    f"Fleet schedule {column_name} at row {row_number} is a formula without cached value. "
                    "Recalculate and save the workbook, then rerun the script."
                )
        return resolved_value

    # Parse integer-like cells and handle formula cells via optional data-only workbook.
    def parse_int_cell(value, *, row_number: int, column_number: int, column_name: str) -> int:
        resolved_value = resolve_cell_value(
            value, row_number=row_number, column_number=column_number, column_name=column_name
        )

        # Accept native numeric values.
        if isinstance(resolved_value, (int, float)):
            return int(float(resolved_value))

        # Parse string numerics such as '12' or '12.0'.
        if isinstance(resolved_value, str):
            text = resolved_value.strip()
            if text == "":
                return 0
            return int(float(text))

        # Default None/missing values to zero.
        if resolved_value is None:
            return 0

        raise ValueError(f"Unsupported fleet schedule {column_name} value at row {row_number}: {resolved_value!r}")

    # Parse float-like cells and handle formula cells via optional data-only workbook.
    def parse_float_cell(value, *, row_number: int, column_number: int, column_name: str) -> float:
        resolved_value = resolve_cell_value(
            value, row_number=row_number, column_number=column_number, column_name=column_name
        )

        # Parse numeric and string values with blanks defaulting to zero.
        if resolved_value is None:
            return 0.0
        if isinstance(resolved_value, (int, float)):
            return float(resolved_value)
        if isinstance(resolved_value, str):
            text = resolved_value.strip()
            if text == "":
                return 0.0
            return float(text.replace(",", ""))
        raise ValueError(f"Unsupported fleet schedule {column_name} value at row {row_number}: {resolved_value!r}")

    # Allocate list to collect month rows.
    rows: List[FleetRow] = []
    # Iterate from first data row to end of used range.
    for row in range(3, fleet_ws.max_row + 1):
        # Extract month/year/purchase/operating columns from known layout.
        month = fleet_ws.cell(row=row, column=1).value
        year = fleet_ws.cell(row=row, column=2).value
        cars_purchased = fleet_ws.cell(row=row, column=4).value
        cars_in_operation = fleet_ws.cell(row=row, column=7).value
        # Stop when month index is missing (including blank formula/text cells).
        resolved_month = resolve_cell_value(month, row_number=row, column_number=1, column_name="month")
        if resolved_month is None or (isinstance(resolved_month, str) and resolved_month.strip() == ""):
            break
        # Append typed row with safe defaults for blanks.
        rows.append(
            FleetRow(
                month=parse_int_cell(resolved_month, row_number=row, column_number=1, column_name="month"),
                year=parse_int_cell(year, row_number=row, column_number=2, column_name="year"),
                cars_purchased=parse_float_cell(cars_purchased, row_number=row, column_number=4, column_name="cars_purchased"),
                cars_in_operation=parse_float_cell(cars_in_operation, row_number=row, column_number=7, column_name="cars_in_operation"),
            )
        )
    # Return extracted fleet rows.
    return rows


# Compute monthly cost-of-sales-per-car from assumptions.
def cost_of_sales_per_car(a: Assumptions) -> float:
    # Return the sum of all direct-cost components per operating car.
    return (
        a.fuel_per_car
        + a.airtime_per_car
        + a.carwash_per_car
        + a.maintenance_per_car
        + a.driver_subsistence_per_car
        + a.incidental_repair_reserve_per_car
        + a.tracking_device_per_car
    )


# Build income statement rows from assumptions and fleet counts.
def build_income_statement_rows(a: Assumptions, fleet_rows: List[FleetRow]) -> List[IncomeRow]:
    # Prepare output list for monthly income statement rows.
    out: List[IncomeRow] = []
    # Loop through each month in the fleet schedule.
    for fr in fleet_rows:
        # Convenience local for operating car count in month.
        cars = fr.cars_in_operation
        # Compute revenue and each direct-cost line.
        revenue = cars * a.monthly_gross_revenue_per_car
        fuel = cars * a.fuel_per_car
        airtime = cars * a.airtime_per_car
        carwash = cars * a.carwash_per_car
        maintenance = cars * a.maintenance_per_car
        driver_subsistence = cars * a.driver_subsistence_per_car
        incidental_reserve = cars * a.incidental_repair_reserve_per_car
        tracking = cars * a.tracking_device_per_car
        # Compute aggregate Cost of Sales from all direct-cost lines.
        cost_of_sales = fuel + airtime + carwash + maintenance + driver_subsistence + incidental_reserve + tracking
        # Compute gross profit and operating expense section.
        gross_profit = revenue - cost_of_sales
        salary = cars * a.salary_per_car
        total_opex = salary
        ebit = gross_profit - total_opex
        # Calculate income tax only when EBIT is positive.
        tax = max(0.0, ebit * a.tax_rate)
        # Compute net profit after tax.
        net_profit = ebit - tax
        # Append fully populated typed row.
        out.append(
            IncomeRow(
                year=fr.year,
                month=fr.month,
                cars_in_operation=cars,
                monthly_gross_revenue=r2(revenue),
                fuel=r2(fuel),
                airtime=r2(airtime),
                carwash=r2(carwash),
                maintenance=r2(maintenance),
                driver_subsistence=r2(driver_subsistence),
                incidental_repair_reserve=r2(incidental_reserve),
                tracking_device=r2(tracking),
                cost_of_sales=r2(cost_of_sales),
                gross_profit=r2(gross_profit),
                salary=r2(salary),
                total_operating_expenses=r2(total_opex),
                ebit=r2(ebit),
                income_tax=r2(tax),
                net_profit=r2(net_profit),
            )
        )
    # Return computed income rows.
    return out


# Build loan amortisation schedule for model horizon.
def build_loan_rows(a: Assumptions, months: int) -> List[LoanRow]:
    # Prepare list for amortisation output rows.
    rows: List[LoanRow] = []
    # Initialize opening balance at zero before draw month.
    balance = 0.0
    # Compute monthly nominal rate from annual rate.
    monthly_rate = a.bank_annual_interest / 12.0
    # Keep separate loan-month index for amortisation table.
    loan_month = 0
    # Iterate over operating months.
    for month in range(1, months + 1):
        # Add loan draw in configured month.
        if month == a.bank_draw_month:
            balance += a.bank_draw
        # If no outstanding balance, skip posting loan payment row.
        if balance <= 0:
            continue
        # Increment loan month once debt exists.
        loan_month += 1
        # Compute monthly interest on opening balance.
        interest = balance * monthly_rate
        # Compute principal as payment less interest, capped at outstanding balance.
        principal = min(max(0.0, a.bank_instalment - interest), balance)
        # Compute effective payment (may be lower in final month).
        payment = interest + principal
        # Compute closing balance after payment.
        closing = max(0.0, balance - principal)
        # Append loan row rounded to currency precision.
        rows.append(
            LoanRow(
                loan_month=loan_month,
                opening_balance=r2(balance),
                interest=r2(interest),
                principal=r2(principal),
                payment=r2(payment),
                closing_balance=r2(closing),
            )
        )
        # Roll closing to next month opening balance.
        balance = closing
    # Return amortisation rows.
    return rows


# Create a month-index map for quick loan lookup by operating month.
def map_loan_by_operating_month(a: Assumptions, months: int, loan_rows: List[LoanRow]) -> Dict[int, LoanRow]:
    # Initialize output dictionary.
    mapping: Dict[int, LoanRow] = {}
    # Track pointer into loan rows.
    lr_idx = 0
    # Track whether loan has started.
    started = False
    # Iterate operating months.
    for m in range(1, months + 1):
        # Start mapping at draw month.
        if m == a.bank_draw_month:
            started = True
        # If loan active and row exists, map month to row.
        if started and lr_idx < len(loan_rows):
            mapping[m] = loan_rows[lr_idx]
            lr_idx += 1
    # Return month-to-loan-row map.
    return mapping


# Build monthly cash flow rows tying income statement and debt schedule together.
def build_cash_flow_rows(a: Assumptions, fleet_rows: List[FleetRow], income_rows: List[IncomeRow], loan_map: Dict[int, LoanRow]) -> List[CashFlowRow]:
    # Prepare cash flow output list.
    out: List[CashFlowRow] = []
    # Iterate month-aligned fleet and income rows in lockstep.
    for fr, ir in zip(fleet_rows, income_rows):
        # Set one-time owner injection in month 1.
        owner_injection = a.owner_equity if fr.month == 1 else 0.0
        # Set one-time bank draw in configured month.
        bank_draw = a.bank_draw if fr.month == a.bank_draw_month else 0.0
        # Compute total operating inflows as revenue plus financing injections.
        total_cash_in = ir.monthly_gross_revenue + owner_injection + bank_draw
        # Load loan values for current month if loan is active.
        loan_row = loan_map.get(fr.month)
        interest = loan_row.interest if loan_row else 0.0
        principal = loan_row.principal if loan_row else 0.0
        # Define operating cash expenses as Cost of Sales + salary.
        operating_expenses = ir.cost_of_sales + ir.salary
        # Define tax as current-month computed tax expense.
        income_tax_paid = ir.income_tax
        # Placeholder investor payout policy set to zero for deterministic generation.
        investor_payout = 0.0
        # Compute net cash before capex.
        net_cash_before_capex = total_cash_in - operating_expenses - interest - principal - income_tax_paid - investor_payout
        # Compute capex from monthly purchases * unit cost.
        capex = fr.cars_purchased * a.car_unit_cost
        # Compute net cash flow after capex.
        net_cash_flow = net_cash_before_capex - capex
        # Append typed monthly cash flow row.
        out.append(
            CashFlowRow(
                month=fr.month,
                owner_injection=r2(owner_injection),
                bank_draw=r2(bank_draw),
                total_cash_in=r2(total_cash_in),
                operating_expenses_cash=r2(operating_expenses),
                interest=r2(interest),
                loan_principal=r2(principal),
                income_tax_paid=r2(income_tax_paid),
                investor_payout=r2(investor_payout),
                net_cash_before_capex=r2(net_cash_before_capex),
                capex=r2(capex),
                net_cash_flow=r2(net_cash_flow),
            )
        )
    # Return computed cash flow rows.
    return out


# Build balance sheet support rows from fleet, cash, income, and loan values.
def build_balance_rows(fleet_rows: List[FleetRow], cash_rows: List[CashFlowRow], income_rows: List[IncomeRow], loan_map: Dict[int, LoanRow]) -> List[BalanceRow]:
    # Initialize output list and cumulative trackers.
    out: List[BalanceRow] = []
    cumulative_owner = 0.0
    cumulative_cars = 0.0
    cumulative_capex = 0.0
    cumulative_revenue = 0.0

    # Iterate aligned monthly rows.
    for fr, cr, ir in zip(fleet_rows, cash_rows, income_rows):
        # Accumulate owner capital injections.
        cumulative_owner += cr.owner_injection
        # Accumulate cumulative cars purchased.
        cumulative_cars += fr.cars_purchased
        # Accumulate cumulative capex and revenue.
        cumulative_capex += cr.capex
        cumulative_revenue += ir.monthly_gross_revenue
        # Pull month-end loan closing balance from loan map.
        loan_closing = loan_map[fr.month].closing_balance if fr.month in loan_map else 0.0
        # Construct and append balance row.
        out.append(
            BalanceRow(
                year=fr.year,
                month=fr.month,
                total_owner_capital_invested=r2(cumulative_owner),
                net_owner_capital_outstanding=r2(max(0.0, cumulative_owner - cumulative_capex)),
                loan_closing_balance=r2(loan_closing),
                cars_in_operation=r2(fr.cars_in_operation),
                total_cars_operation_cumulative=r2(cumulative_cars),
                vehicles_in_use_cost=r2(fr.cars_in_operation * (cumulative_capex / cumulative_cars) if cumulative_cars > 0 else 0.0),
                total_cars_cost_cumulative=r2(cumulative_capex),
                net_cash_before_capex=r2(cr.net_cash_before_capex),
                capex=r2(cr.capex),
                gross_revenue_cumulative=r2(cumulative_revenue),
            )
        )
    # Return monthly balance rows.
    return out


# Rewrite Income_Statement sheet with explicit new columns and regenerated values.
def write_income_statement(ws, rows: List[IncomeRow]):
    # Overwrite header row with updated required columns.
    headers = [
        "Year #",
        "Month #",
        "Cars in Operation",
        "Monthly Gross Revenue",
        "Fuel",
        "Airtime",
        "Carwash",
        "Maintenance",
        "Driver Subsistence Allocation",
        "Incidental Repair Reserve",
        "Tracking Device",
        "Cost of Sales",
        "Gross Profit",
        "Salary",
        "Total Operating Expenses",
        "EBIT",
        "Income Tax",
        "Net Profit",
    ]
    # Write the refreshed headers to row 2.
    for idx, header in enumerate(headers, start=1):
        ws.cell(row=2, column=idx, value=header)

    # Write each monthly row starting at row 3.
    for row_idx, r in enumerate(rows, start=3):
        ws.cell(row=row_idx, column=1, value=r.year)
        ws.cell(row=row_idx, column=2, value=r.month)
        ws.cell(row=row_idx, column=3, value=r.cars_in_operation)
        ws.cell(row=row_idx, column=4, value=r.monthly_gross_revenue)
        ws.cell(row=row_idx, column=5, value=r.fuel)
        ws.cell(row=row_idx, column=6, value=r.airtime)
        ws.cell(row=row_idx, column=7, value=r.carwash)
        ws.cell(row=row_idx, column=8, value=r.maintenance)
        ws.cell(row=row_idx, column=9, value=r.driver_subsistence)
        ws.cell(row=row_idx, column=10, value=r.incidental_repair_reserve)
        ws.cell(row=row_idx, column=11, value=r.tracking_device)
        ws.cell(row=row_idx, column=12, value=r.cost_of_sales)
        ws.cell(row=row_idx, column=13, value=r.gross_profit)
        ws.cell(row=row_idx, column=14, value=r.salary)
        ws.cell(row=row_idx, column=15, value=r.total_operating_expenses)
        ws.cell(row=row_idx, column=16, value=r.ebit)
        ws.cell(row=row_idx, column=17, value=r.income_tax)
        ws.cell(row=row_idx, column=18, value=r.net_profit)


# Rewrite Cash_Flow sheet with regenerated values and operating expense tie-outs.
def write_cash_flow(ws, rows: List[CashFlowRow]):
    # Write headers to row 2.
    headers = [
        "Month #",
        "Owner/Investor Injection",
        "Bank Loan Draw",
        "Total Cash In",
        "Operating Expenses (cash)",
        "Interest",
        "Loan Principal",
        "Income Tax Paid",
        "Investor Payout/Profit Distribution",
        "Net Cash Before Capex",
        "Capex (car purchases)",
        "Net Cash Flow",
    ]
    # Populate header cells.
    for idx, header in enumerate(headers, start=1):
        ws.cell(row=2, column=idx, value=header)

    # Populate monthly rows.
    for row_idx, r in enumerate(rows, start=3):
        ws.cell(row=row_idx, column=1, value=r.month)
        ws.cell(row=row_idx, column=2, value=r.owner_injection)
        ws.cell(row=row_idx, column=3, value=r.bank_draw)
        ws.cell(row=row_idx, column=4, value=r.total_cash_in)
        ws.cell(row=row_idx, column=5, value=r.operating_expenses_cash)
        ws.cell(row=row_idx, column=6, value=r.interest)
        ws.cell(row=row_idx, column=7, value=r.loan_principal)
        ws.cell(row=row_idx, column=8, value=r.income_tax_paid)
        ws.cell(row=row_idx, column=9, value=r.investor_payout)
        ws.cell(row=row_idx, column=10, value=r.net_cash_before_capex)
        ws.cell(row=row_idx, column=11, value=r.capex)
        ws.cell(row=row_idx, column=12, value=r.net_cash_flow)


# Rewrite Loan_Amortisation sheet with regenerated schedule rows.
def write_loan_amortisation(ws, rows: List[LoanRow]):
    # Write standard headers in row 4 to align with template layout.
    headers = ["Loan month", "Opening balance", "Interest", "Principal", "Payment", "Closing balance"]
    # Populate headers.
    for idx, header in enumerate(headers, start=1):
        ws.cell(row=4, column=idx, value=header)

    # Write loan rows starting at row 5.
    for row_idx, r in enumerate(rows, start=5):
        ws.cell(row=row_idx, column=1, value=r.loan_month)
        ws.cell(row=row_idx, column=2, value=r.opening_balance)
        ws.cell(row=row_idx, column=3, value=r.interest)
        ws.cell(row=row_idx, column=4, value=r.principal)
        ws.cell(row=row_idx, column=5, value=r.payment)
        ws.cell(row=row_idx, column=6, value=r.closing_balance)


# Rewrite Balance_Sheet with regenerated support metrics.
def write_balance_sheet(ws, rows: List[BalanceRow]):
    # Define headers matching model output requirements.
    headers = [
        "Year #",
        "Month #",
        "Total Owner/Investor Capital Invested",
        "Net Owner/Investor Capital Outstanding",
        "Loan Closing Balance",
        "Cars in Operation (count)",
        "Total Cars Operation (Cummulative)",
        "Vehicles In Use - Cost",
        "Total Cars Cost (Cummulative)",
        "Net Cash Before Capex",
        "Capex (car purchases)",
        "Gross Revenue (Cummulative)",
    ]
    # Write headers on row 2.
    for idx, header in enumerate(headers, start=1):
        ws.cell(row=2, column=idx, value=header)

    # Write row values starting at row 3.
    for row_idx, r in enumerate(rows, start=3):
        ws.cell(row=row_idx, column=1, value=r.year)
        ws.cell(row=row_idx, column=2, value=r.month)
        ws.cell(row=row_idx, column=3, value=r.total_owner_capital_invested)
        ws.cell(row=row_idx, column=4, value=r.net_owner_capital_outstanding)
        ws.cell(row=row_idx, column=5, value=r.loan_closing_balance)
        ws.cell(row=row_idx, column=6, value=r.cars_in_operation)
        ws.cell(row=row_idx, column=7, value=r.total_cars_operation_cumulative)
        ws.cell(row=row_idx, column=8, value=r.vehicles_in_use_cost)
        ws.cell(row=row_idx, column=9, value=r.total_cars_cost_cumulative)
        ws.cell(row=row_idx, column=10, value=r.net_cash_before_capex)
        ws.cell(row=row_idx, column=11, value=r.capex)
        ws.cell(row=row_idx, column=12, value=r.gross_revenue_cumulative)


# Basic validation checks to catch accounting and shape anomalies early.
def run_validations(income_rows: List[IncomeRow], cash_rows: List[CashFlowRow], fleet_rows: List[FleetRow]):
    # Ensure all core lists have matching month counts.
    if not (len(income_rows) == len(cash_rows) == len(fleet_rows)):
        raise ValueError("Income, Cash, and Fleet row counts are not aligned")

    # Check each month that cost-of-sales ties to direct line-item sum.
    for row in income_rows:
        expected_cos = r2(
            row.fuel
            + row.airtime
            + row.carwash
            + row.maintenance
            + row.driver_subsistence
            + row.incidental_repair_reserve
            + row.tracking_device
        )
        if r2(row.cost_of_sales) != expected_cos:
            raise ValueError(f"Cost-of-sales tie-out failed in month {row.month}")


# Main projection pipeline that orchestrates read -> compute -> write.
def run_projection(input_path: Path, output_path: Path):
    # Load workbook template from input path.
    wb = load_workbook(filename=input_path)
    # Load data-only workbook for cached formula results in assumption cells.
    wb_values = load_workbook(filename=input_path, data_only=True)
    # Get sheet handles used by the model.
    assumptions_ws = wb["Assumptions"]
    assumptions_values_ws = wb_values["Assumptions"]
    fleet_ws = wb["Fleet_Schedule"]
    fleet_values_ws = wb_values["Fleet_Schedule"]
    income_ws = wb["Income_Statement"]
    cash_ws = wb["Cash_Flow"]
    loan_ws = wb["Loan_Amortisation"]
    balance_ws = wb["Balance_Sheet"]

    # Apply required assumption updates/additions.
    ensure_required_assumptions(assumptions_ws)
    # Read assumptions from sheet after updates.
    assumption_values = read_assumption_values(assumptions_ws, assumptions_values_ws)
    # Convert assumptions into typed object.
    assumptions = build_assumptions(assumption_values)

    # Read fleet schedule source rows.
    fleet_rows = read_fleet_schedule(fleet_ws, fleet_values_ws)
    # Build monthly income statement rows.
    income_rows = build_income_statement_rows(assumptions, fleet_rows)
    # Build loan amortisation rows using model horizon.
    loan_rows = build_loan_rows(assumptions, len(fleet_rows))
    # Create month-to-loan lookup map.
    loan_map = map_loan_by_operating_month(assumptions, len(fleet_rows), loan_rows)
    # Build monthly cash flow rows.
    cash_rows = build_cash_flow_rows(assumptions, fleet_rows, income_rows, loan_map)
    # Build balance sheet rows from model outputs.
    balance_rows = build_balance_rows(fleet_rows, cash_rows, income_rows, loan_map)

    # Execute validation checks before writing output.
    run_validations(income_rows, cash_rows, fleet_rows)

    # Write regenerated outputs back into workbook sheets.
    write_income_statement(income_ws, income_rows)
    write_cash_flow(cash_ws, cash_rows)
    write_loan_amortisation(loan_ws, loan_rows)
    write_balance_sheet(balance_ws, balance_rows)

    # Save completed workbook to output path.
    wb.save(output_path)


# Parse command-line arguments for input/output workbook paths.
def parse_args():
    # Create argument parser with concise help text.
    parser = argparse.ArgumentParser(description="Generate financial projections from assumptions.")
    # Add required input Excel workbook argument.
    parser.add_argument("--input", required=True, help="Path to source financial_projections.xlsx")
    # Add optional output path argument.
    parser.add_argument("--output", required=False, help="Path to output workbook")
    # Parse and return namespace.
    return parser.parse_args()


# Standard script entrypoint.
def main():
    # Parse incoming CLI flags.
    args = parse_args()
    # Resolve input path to pathlib object.
    input_path = Path(args.input)
    # Fail early if source workbook does not exist.
    if not input_path.exists():
        raise FileNotFoundError(f"Input workbook not found: {input_path}")

    # Build default output filename when user does not pass --output.
    if args.output:
        output_path = Path(args.output)
    else:
        stamp = datetime.utcnow().strftime("%Y%m%d_%H%M%S")
        output_path = input_path.parent / f"financial_projections_generated_{stamp}.xlsx"

    # Run end-to-end projection generation.
    run_projection(input_path=input_path, output_path=output_path)
    # Print completion message for terminal visibility.
    print(f"Projection model generated successfully: {output_path}")


# Execute main only when run as a script, not when imported as a module.
if __name__ == "__main__":
    # Call main entrypoint.
    main()
