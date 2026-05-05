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
from openpyxl.formula.translate import Translator


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
    "owner_second_equity": "Owner/investor equity injection (batch 2)",
    "owner_second_month": "Owner second injection month (operating month index)",
    "initial_cars": "Initial cars in operation (month 1)",
    "procurement_lead_time": "Procurement lead time",
    "model_horizon": "Model horizon",
    "bank_draw": "Bank loan draw (batch 2)",
    "bank_draw_month": "Bank loan draw month (operating month index)",
    "bank_instalment": "Bank monthly instalment (principal + interest)",
    "bank_annual_interest": "Bank nominal annual interest rate",
    "bank_loan_term": "Bank loan term",
    "vehicle_disposal_trigger": "Vehicle disposal trigger",
}



REQUIRED_ASSUMPTION_LABELS = {
    ASSUMPTION_KEYS["monthly_gross_revenue_per_car"],
    ASSUMPTION_KEYS["fuel_per_car"],
    ASSUMPTION_KEYS["airtime_per_car"],
    ASSUMPTION_KEYS["carwash_per_car"],
    "Maintenance expense per active car (monthly)",
    "Driver subsistence",
    "Incidental repair reserve",
    "Tracking device expense",
    ASSUMPTION_KEYS["salary_per_car"],
    ASSUMPTION_KEYS["tax_rate"],
    ASSUMPTION_KEYS["car_unit_cost"],
    ASSUMPTION_KEYS["owner_equity"],
    ASSUMPTION_KEYS["owner_second_equity"],
    ASSUMPTION_KEYS["owner_second_month"],
    ASSUMPTION_KEYS["initial_cars"],
    ASSUMPTION_KEYS["procurement_lead_time"],
    ASSUMPTION_KEYS["model_horizon"],
    ASSUMPTION_KEYS["bank_draw"],
    ASSUMPTION_KEYS["bank_draw_month"],
    ASSUMPTION_KEYS["bank_instalment"],
    ASSUMPTION_KEYS["bank_annual_interest"],
    ASSUMPTION_KEYS["bank_loan_term"],
    ASSUMPTION_KEYS["vehicle_disposal_trigger"],
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
    owner_second_equity: float
    owner_second_month: int
    initial_cars: int
    procurement_lead_time: int
    model_horizon: int
    bank_draw: float
    bank_draw_month: int
    bank_payment_start_month: int
    bank_instalment: float
    bank_annual_interest: float
    bank_loan_term: int
    vehicle_disposal_trigger_years: int


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
    # Delete any deprecated investor payout assumption row.
    investor_payout_row = find_assumption_row(assumptions_ws, "Investor monthly payout per N$250,000 tranche")
    if investor_payout_row is not None:
        assumptions_ws.delete_rows(investor_payout_row, 1)

    # Delete deprecated monthly operating profit-per-car assumption rows.
    for deprecated_label in (
        "Assumption Monthly operating profit per car",
        "Monthly operating profit per car",
    ):
        deprecated_row = find_assumption_row(assumptions_ws, deprecated_label)
        if deprecated_row is not None:
            assumptions_ws.delete_rows(deprecated_row, 1)

    # Keep key cost assumptions aligned with required overrides used in downstream calculations.
    for label, value in (
        (ASSUMPTION_KEYS["airtime_per_car"], 290),
        ("Maintenance expense per active car (monthly)", 1250),
    ):
        row = find_assumption_row(assumptions_ws, label)
        if row is not None:
            assumptions_ws.cell(row=row, column=2, value=value)

    # Force row 6 content to the requested vehicle disposal trigger values.
    assumptions_ws.cell(row=6, column=1, value="Vehicle disposal trigger")
    assumptions_ws.cell(row=6, column=2, value=2)
    assumptions_ws.cell(row=6, column=2).number_format = "0"
    assumptions_ws.cell(row=6, column=3, value="Years")
    assumptions_ws.cell(row=6, column=4, value="car trigger for the disposal (with no disposal revenue) of all vehicles.")

    # Insert four rows at row 12, moving the existing row 12 and all rows below to row 16+.
    assumptions_ws.insert_rows(12, amount=4)

    # Write the requested assumptions into rows 12-15.
    assumptions_ws.cell(row=12, column=1, value="Driver subsistence")
    assumptions_ws.cell(row=12, column=2, value=1400)
    assumptions_ws.cell(row=12, column=2).number_format = '"N$" #,##0.00'
    assumptions_ws.cell(row=12, column=3, value="N$ / car / month")
    assumptions_ws.cell(row=12, column=4, value="For client upkeep")

    assumptions_ws.cell(row=13, column=1, value="Incidental repair reserve")
    assumptions_ws.cell(row=13, column=2, value=500)
    assumptions_ws.cell(row=13, column=2).number_format = '"N$" #,##0.00'
    assumptions_ws.cell(row=13, column=3, value="N$ / car / month")
    assumptions_ws.cell(row=13, column=4, value="Internal repair insurance")

    assumptions_ws.cell(row=14, column=1, value="Tracking device expense")
    assumptions_ws.cell(row=14, column=2, value=950)
    assumptions_ws.cell(row=14, column=2).number_format = '"N$" #,##0.00'
    assumptions_ws.cell(row=14, column=3, value="N$ / car / month")
    assumptions_ws.cell(row=14, column=4, value="Monthly payment for 3 year contract. 1 tracking device & 1 dashcam.")

    assumptions_ws.cell(row=15, column=1, value="Cost of sales")
    assumptions_ws.cell(row=15, column=2, value=None)
    assumptions_ws.cell(row=15, column=2).number_format = '"N$" #,##0.00'
    assumptions_ws.cell(row=15, column=3, value="N$ / car / month")
    assumptions_ws.cell(row=15, column=4, value="fuel+airtime+carwash+maintenance+subsistence+repairs+tracking.")

    # Recalculate Cost of sales from rows 8-11 plus inserted subsistence/repair/tracking values.
    def get_value(label: str) -> float | None:
        row = find_assumption_row(assumptions_ws, label)
        if row is None:
            raise KeyError(f"Required assumption row not found for update: {label}")
        cell_value = assumptions_ws.cell(row=row, column=2).value
        if cell_value is None:
            return 0.0
        if isinstance(cell_value, str):
            text_value = cell_value.strip()
            if text_value.startswith("="):
                return None
            normalized = text_value.replace("N$", "").replace("$", "").replace(",", "").strip()
            return float(normalized)
        return float(cell_value)

    cost_component_labels = (
        ASSUMPTION_KEYS["fuel_per_car"],
        ASSUMPTION_KEYS["airtime_per_car"],
        ASSUMPTION_KEYS["carwash_per_car"],
        "Maintenance expense per active car (monthly)",
        "Driver subsistence",
        "Incidental repair reserve",
        "Tracking device expense",
    )
    cost_component_values = [get_value(label) for label in cost_component_labels]
    if all(value is not None for value in cost_component_values):
        assumptions_ws.cell(row=15, column=2, value=sum(cost_component_values))
    cost_of_sales_value = float(assumptions_ws.cell(row=15, column=2).value or 0.0)

    # Ensure Monthly Gross profit per car is ordered above Salary and Total operating expenses rows.
    gross_profit_row = find_assumption_row(assumptions_ws, "Monthly Gross profit per car")
    salary_row = find_assumption_row(assumptions_ws, ASSUMPTION_KEYS["salary_per_car"])
    total_opex_row = find_assumption_row(assumptions_ws, "Total operating expenses per operating car (monthly)")
    if gross_profit_row is None or salary_row is None or total_opex_row is None:
        raise KeyError("Could not find one of required rows: Monthly Gross profit, Salary, or Total operating expenses.")

    # Collect values first, then remove old rows and reinsert in requested order.
    row_payloads = []
    for row_idx in (gross_profit_row, salary_row, total_opex_row):
        row_payloads.append([assumptions_ws.cell(row=row_idx, column=col).value for col in range(1, 5)])
    for row_idx in sorted((gross_profit_row, salary_row, total_opex_row), reverse=True):
        assumptions_ws.delete_rows(row_idx, 1)

    insert_at = min(gross_profit_row, salary_row, total_opex_row)
    assumptions_ws.insert_rows(insert_at, amount=3)
    for offset, payload in enumerate(row_payloads):
        for col_idx, value in enumerate(payload, start=1):
            assumptions_ws.cell(row=insert_at + offset, column=col_idx, value=value)

    gross_profit_row = insert_at
    salary_row = insert_at + 1
    total_opex_row = insert_at + 2

    # Total operating expenses per car should include Cost of sales plus Salary expense.
    salary_value = float(assumptions_ws.cell(row=salary_row, column=2).value or 0.0)
    assumptions_ws.cell(row=salary_row, column=2).number_format = '"N$" #,##0.00'
    assumptions_ws.cell(row=total_opex_row, column=2, value=cost_of_sales_value + salary_value)
    assumptions_ws.cell(row=total_opex_row, column=2).number_format = '"N$" #,##0.00'
    assumptions_ws.cell(row=total_opex_row, column=4, value='Calculated = Cost of sales + Salary expense per operating car (monthly).')

    # Monthly Gross profit per car should be monthly gross revenue per car minus cost of sales.
    revenue_row = find_assumption_row(assumptions_ws, ASSUMPTION_KEYS["monthly_gross_revenue_per_car"])
    revenue_value = float(assumptions_ws.cell(row=revenue_row, column=2).value or 0.0) if revenue_row else 0.0
    assumptions_ws.cell(row=gross_profit_row, column=2, value=revenue_value - cost_of_sales_value)
    assumptions_ws.cell(row=gross_profit_row, column=2).number_format = '"N$" #,##0.00'

    # Ensure there is exactly one Operating Profit assumption row.
    operating_profit_rows = [
        row
        for row in range(1, assumptions_ws.max_row + 1)
        if assumptions_ws.cell(row=row, column=1).value == "Operating Profit"
    ]
    if not operating_profit_rows:
        assumptions_ws.insert_rows(total_opex_row + 1, amount=1)
        operating_profit_row = total_opex_row + 1
        assumptions_ws.cell(row=operating_profit_row, column=1, value="Operating Profit")
        assumptions_ws.cell(row=operating_profit_row, column=3, value="N$ / car / month")
        assumptions_ws.cell(
            row=operating_profit_row,
            column=4,
            value="Gross profit - Value for Operating Expense.",
        )
        operating_profit_rows = [operating_profit_row]

    for row in reversed(operating_profit_rows[1:]):
        assumptions_ws.delete_rows(row, 1)

    operating_profit_row = operating_profit_rows[0]
    assumptions_ws.cell(row=operating_profit_row, column=2, value=(revenue_value - cost_of_sales_value) - salary_value)
    assumptions_ws.cell(row=operating_profit_row, column=2).number_format = '"N$" #,##0.00'

    # Rename owner second injection month label when present.
    owner_injection_month_row = find_assumption_row(assumptions_ws, "Owner injection month (operating month index)")
    if owner_injection_month_row is not None:
        assumptions_ws.cell(
            row=owner_injection_month_row,
            column=1,
            value="Owner second injection month (operating month index)",
        )

    # Set Batch 2 investor payout start month to 8 when the row exists.
    batch_2_payout_start_row = find_assumption_row(assumptions_ws, "Batch 2 investor payout start month")
    if batch_2_payout_start_row is not None:
        assumptions_ws.cell(row=batch_2_payout_start_row, column=2, value=8)
        assumptions_ws.cell(row=batch_2_payout_start_row, column=2).number_format = "0"

    # Set monthly bank interest rate from annual nominal rate / 12.
    bank_annual_interest_row = find_assumption_row(assumptions_ws, "Bank nominal annual interest rate")
    bank_monthly_interest_row = find_assumption_row(assumptions_ws, "Bank monthly interest rate")
    if bank_annual_interest_row is not None and bank_monthly_interest_row is not None:
        annual_cell = assumptions_ws.cell(row=bank_annual_interest_row, column=2).coordinate
        assumptions_ws.cell(
            row=bank_monthly_interest_row,
            column=2,
            value=f"={annual_cell}/12",
        )
        assumptions_ws.cell(row=bank_monthly_interest_row, column=2).number_format = "0.00%"

    # Set bank instalment from PMT(monthly interest rate, loan term, loan draw batch 2).
    bank_instalment_row = find_assumption_row(assumptions_ws, "Bank monthly instalment (principal + interest)")
    bank_loan_term_row = find_assumption_row(assumptions_ws, "Bank loan term")
    bank_draw_row = find_assumption_row(assumptions_ws, "Bank loan draw (batch 2)")
    if (
        bank_instalment_row is not None
        and bank_monthly_interest_row is not None
        and bank_loan_term_row is not None
        and bank_draw_row is not None
    ):
        monthly_interest_cell = assumptions_ws.cell(row=bank_monthly_interest_row, column=2).coordinate
        loan_term_cell = assumptions_ws.cell(row=bank_loan_term_row, column=2).coordinate
        bank_draw_cell = assumptions_ws.cell(row=bank_draw_row, column=2).coordinate
        assumptions_ws.cell(
            row=bank_instalment_row,
            column=2,
            value=f"=-PMT({monthly_interest_cell},{loan_term_cell},{bank_draw_cell})",
        )
        assumptions_ws.cell(row=bank_instalment_row, column=2).number_format = '"N$" #,##0.00'


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

    # If the instalment value is still missing (typically due to an unevaluated formula),
    # calculate it from the same assumptions inputs used in the Assumptions sheet PMT formula.
    bank_instalment_label = ASSUMPTION_KEYS["bank_instalment"]
    if bank_instalment_label not in values:
        annual_rate_label = ASSUMPTION_KEYS["bank_annual_interest"]
        loan_term_label = ASSUMPTION_KEYS["bank_loan_term"]
        loan_draw_label = ASSUMPTION_KEYS["bank_draw"]
        if (
            annual_rate_label in values
            and loan_term_label in values
            and loan_draw_label in values
        ):
            monthly_rate = values[annual_rate_label] / 12.0
            loan_term = int(values[loan_term_label])
            loan_draw = values[loan_draw_label]
            if loan_term > 0 and loan_draw > 0:
                if abs(monthly_rate) < 1e-12:
                    values[bank_instalment_label] = loan_draw / loan_term
                else:
                    growth = (1 + monthly_rate) ** loan_term
                    values[bank_instalment_label] = loan_draw * (monthly_rate * growth) / (growth - 1)
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
        driver_subsistence_per_car=get("Driver subsistence"),
        incidental_repair_reserve_per_car=get("Incidental repair reserve"),
        tracking_device_per_car=get("Tracking device expense"),
        salary_per_car=get(ASSUMPTION_KEYS["salary_per_car"]),
        tax_rate=get(ASSUMPTION_KEYS["tax_rate"]),
        car_unit_cost=get(ASSUMPTION_KEYS["car_unit_cost"]),
        owner_equity=get(ASSUMPTION_KEYS["owner_equity"]),
        owner_second_equity=get(ASSUMPTION_KEYS["owner_second_equity"]),
        owner_second_month=int(get(ASSUMPTION_KEYS["owner_second_month"])),
        initial_cars=int(get(ASSUMPTION_KEYS["initial_cars"])),
        procurement_lead_time=int(get(ASSUMPTION_KEYS["procurement_lead_time"])),
        model_horizon=int(get(ASSUMPTION_KEYS["model_horizon"])),
        bank_draw=get(ASSUMPTION_KEYS["bank_draw"]),
        bank_draw_month=int(get(ASSUMPTION_KEYS["bank_draw_month"])),
        bank_payment_start_month=int(assumption_values.get("Bank payment start month", get(ASSUMPTION_KEYS["bank_draw_month"]))),
        bank_instalment=get(ASSUMPTION_KEYS["bank_instalment"]),
        bank_annual_interest=get(ASSUMPTION_KEYS["bank_annual_interest"]),
        bank_loan_term=int(get(ASSUMPTION_KEYS["bank_loan_term"])),
        vehicle_disposal_trigger_years=int(get(ASSUMPTION_KEYS["vehicle_disposal_trigger"])),
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


# Recalculate fleet schedule columns from purchases onward using assumptions.
def recalculate_fleet_rows(a: Assumptions, source_rows: List[FleetRow]) -> List[FleetRow]:
    months = a.model_horizon if a.model_horizon > 0 else len(source_rows)
    disposal_lag_months = max(0, int(a.vehicle_disposal_trigger_years * 12))
    purchases = {r.month: r.cars_purchased for r in source_rows}
    rows: List[FleetRow] = []
    deliveries_by_month: Dict[int, float] = {}
    disposals_by_month: Dict[int, float] = {}
    active = float(a.initial_cars)
    cumulative_disposed = 0.0

    # Initial owner-funded fleet starts operating in month 1 and must also be disposed on the trigger horizon.
    if disposal_lag_months > 0 and a.initial_cars > 0:
        initial_disposal_month = 1 + disposal_lag_months
        disposals_by_month[initial_disposal_month] = disposals_by_month.get(initial_disposal_month, 0.0) + float(a.initial_cars)

    for month in range(1, months + 1):
        purchases_m = float(purchases.get(month, 0.0))
        if purchases_m > 0 and a.procurement_lead_time >= 0:
            delivery_month = month + a.procurement_lead_time
            deliveries_by_month[delivery_month] = deliveries_by_month.get(delivery_month, 0.0) + purchases_m
        deliveries = deliveries_by_month.get(month, 0.0)
        if deliveries > 0 and disposal_lag_months > 0:
            disposal_month = month + disposal_lag_months
            disposals_by_month[disposal_month] = disposals_by_month.get(disposal_month, 0.0) + deliveries
        disposals = disposals_by_month.get(month, 0.0)
        active = max(0.0, active + deliveries - disposals)
        cumulative_disposed += disposals
        rows.append(FleetRow(month=month, year=((month - 1) // 12) + 1, cars_purchased=purchases_m, cars_in_operation=active))
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




# Resolve bank payment start month using explicit assumption when present, else draw month fallback.
def get_bank_payment_start_month(a: Assumptions) -> int:
    # Payment start month must be at least 1.
    return max(1, int(getattr(a, "bank_payment_start_month", a.bank_draw_month) or a.bank_draw_month))


# Build loan amortisation schedule for model horizon.
def build_loan_rows(a: Assumptions, months: int, cash_rows: List[CashFlowRow]) -> List[LoanRow]:
    # Rebuild amortisation rows from Cash_Flow Interest and Loan Principal values using payment start month alignment.
    rows: List[LoanRow] = []
    payment_start = get_bank_payment_start_month(a)
    balance = float(a.bank_draw)
    term_months = max(0, int(a.bank_loan_term))
    for loan_month in range(1, term_months + 1):
        operating_month = payment_start + loan_month - 1
        if operating_month > months:
            break
        cash_row = cash_rows[operating_month - 1]
        interest = max(0.0, float(cash_row.interest))
        principal = max(0.0, min(float(cash_row.loan_principal), balance))
        payment = interest + principal
        opening = balance
        closing = max(0.0, opening - principal)
        rows.append(LoanRow(
            loan_month=loan_month,
            opening_balance=r2(opening),
            interest=r2(interest),
            principal=r2(principal),
            payment=r2(payment),
            closing_balance=r2(closing),
        ))
        balance = closing
    return rows


# Build monthly cash flow rows tying income statement and debt schedule together.
def build_cash_flow_rows(a: Assumptions, fleet_rows: List[FleetRow], income_rows: List[IncomeRow]) -> List[CashFlowRow]:
    out: List[CashFlowRow] = []
    monthly_rate = a.bank_annual_interest / 12.0
    payment_start = get_bank_payment_start_month(a)
    term_end_exclusive = payment_start + max(0, int(a.bank_loan_term))
    opening_cash_balance = 0.0
    opening_loan_balance = 0.0

    for fr, ir in zip(fleet_rows, income_rows):
        owner_injection = a.owner_second_equity if fr.month == a.owner_second_month else 0.0
        bank_draw = a.bank_draw if fr.month == a.bank_draw_month else 0.0
        opening_loan_balance += bank_draw
        total_cash_in = ir.monthly_gross_revenue + owner_injection + bank_draw

        in_payment_window = payment_start <= fr.month < term_end_exclusive
        interest = (opening_loan_balance * monthly_rate) if in_payment_window else 0.0
        principal = min(max(0.0, opening_loan_balance), max(0.0, a.bank_instalment - interest)) if in_payment_window else 0.0
        closing_loan_balance = max(0.0, opening_loan_balance - principal)

        operating_expenses = ir.cost_of_sales + ir.salary
        income_tax_paid = ir.income_tax
        investor_payout = 0.0
        net_cash_before_capex = total_cash_in - operating_expenses - interest - principal - income_tax_paid - investor_payout
        capex = fr.cars_purchased * a.car_unit_cost
        net_cash_flow = net_cash_before_capex - capex

        out.append(CashFlowRow(
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
        ))
        opening_cash_balance = net_cash_before_capex
        opening_loan_balance = closing_loan_balance
    return out


# Create a month-index map for quick loan lookup by operating month.
def map_loan_by_operating_month(a: Assumptions, months: int, loan_rows: List[LoanRow]) -> Dict[int, LoanRow]:
    mapping: Dict[int, LoanRow] = {}
    payment_start = get_bank_payment_start_month(a)
    for lr in loan_rows:
        operating_month = payment_start + lr.loan_month - 1
        if 1 <= operating_month <= months:
            mapping[operating_month] = lr
    return mapping


# Build balance sheet support rows from fleet, cash, income, and loan values.
def build_balance_rows(a: Assumptions, fleet_rows: List[FleetRow], cash_rows: List[CashFlowRow], income_rows: List[IncomeRow], loan_map: Dict[int, LoanRow]) -> List[BalanceRow]:
    # Initialize output list and cumulative trackers.
    out: List[BalanceRow] = []
    cumulative_owner = float(a.owner_equity)
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


# Rewrite Fleet_Schedule starting from cars purchased and derived operational counts.
def write_fleet_schedule(ws, rows: List[FleetRow], a: Assumptions):
    cumulative_in_operation = 0.0
    disposal_lag_months = max(0, int(a.vehicle_disposal_trigger_years * 12))
    deliveries_by_month: Dict[int, float] = {}
    disposals_by_month: Dict[int, float] = {}
    in_pipeline = 0.0
    total_cars = float(a.initial_cars)
    cumulative_disposed = 0.0

    # Initial owner-funded cars are already in operation at month 1 and should be disposed on trigger month.
    if disposal_lag_months > 0 and a.initial_cars > 0:
        initial_disposal_month = 1 + disposal_lag_months
        disposals_by_month[initial_disposal_month] = disposals_by_month.get(initial_disposal_month, 0.0) + float(a.initial_cars)

    template_formulas = {}
    for col in range(1, 12):
        template_value = ws.cell(row=3, column=col).value
        if isinstance(template_value, str) and template_value.startswith("="):
            template_formulas[col] = template_value

    for row_idx, r in enumerate(rows, start=3):
        deliveries_by_month[r.month + a.procurement_lead_time] = deliveries_by_month.get(r.month + a.procurement_lead_time, 0.0) + r.cars_purchased
        deliveries = deliveries_by_month.get(r.month, 0.0)
        if deliveries > 0 and disposal_lag_months > 0:
            disposal_month = r.month + disposal_lag_months
            disposals_by_month[disposal_month] = disposals_by_month.get(disposal_month, 0.0) + deliveries
        disposals = disposals_by_month.get(r.month, 0.0)
        in_pipeline = max(0.0, in_pipeline + r.cars_purchased - deliveries)
        total_cars = max(0.0, total_cars + r.cars_purchased - disposals)
        cumulative_in_operation += deliveries
        if r.month == 1:
            cumulative_in_operation = r.cars_in_operation
        else:
            cumulative_in_operation = max(cumulative_in_operation - disposals, r.cars_in_operation)
        cumulative_disposed += disposals

        ws.cell(row=row_idx, column=1, value=r.month)
        ws.cell(row=row_idx, column=2, value=r.year)
        ws.cell(row=row_idx, column=3, value=((r.month - 1) % 12) + 1)
        ws.cell(row=row_idx, column=4, value=r.cars_purchased)
        ws.cell(row=row_idx, column=5, value=deliveries)
        ws.cell(row=row_idx, column=6, value=in_pipeline)
        ws.cell(row=row_idx, column=7, value=r.cars_in_operation)
        ws.cell(row=row_idx, column=8, value=total_cars)
        if row_idx == 3:
            ws.cell(row=row_idx, column=9, value="=Assumptions!$B$29")
        else:
            ws.cell(row=row_idx, column=9, value=f"=$I{row_idx-1}+$E{row_idx}")
        ws.cell(row=row_idx, column=10, value=disposals)
        ws.cell(row=row_idx, column=11, value=cumulative_disposed)

        for col, formula in template_formulas.items():
            if col == 9:
                continue
            ws.cell(row=row_idx, column=col, value=Translator(formula, origin=f"{ws.cell(row=3, column=col).coordinate}").translate_formula(f"{ws.cell(row=row_idx, column=col).coordinate}"))

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
        ws.cell(row=row_idx, column=6, value=f"=IF(AND($A{row_idx}>=Assumptions!$B$24,$A{row_idx}<Assumptions!$B$24+Assumptions!$B$31),($O{row_idx}+$C{row_idx})*Assumptions!$B$32,0)")
        ws.cell(row=row_idx, column=7, value=f"=IF(AND($A{row_idx}>=Assumptions!$B$24,$A{row_idx}<Assumptions!$B$24+Assumptions!$B$31),MIN(($O{row_idx}+$C{row_idx}),Assumptions!$B$23-$F{row_idx}),0)")
        ws.cell(row=row_idx, column=8, value=r.income_tax_paid)
        ws.cell(
            row=row_idx,
            column=9,
            value=(
                f"=IF($A{row_idx}<Assumptions!$B$28,Assumptions!$B$34*Fleet_Schedule!$G{row_idx},"
                f"ROUND(Assumptions!$B$34*(Fleet_Schedule!$G{row_idx}-(Assumptions!$B$23/Assumptions!$B$4)),0))"
            ),
        )
        ws.cell(row=row_idx, column=10, value=f"=D{row_idx}-E{row_idx}-F{row_idx}-G{row_idx}-H{row_idx}-I{row_idx}")
        ws.cell(row=row_idx, column=11, value=r.capex)
        ws.cell(row=row_idx, column=12, value=f"=J{row_idx}-K{row_idx}")


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
        ws.cell(row=row_idx, column=2, value=("=Assumptions!$B$23" if row_idx == 5 else f"=F{row_idx-1}"))
        ws.cell(row=row_idx, column=3, value=f"=INDEX(Cash_Flow!$F:$F,MATCH(Assumptions!$B$28+A{row_idx}-1,Cash_Flow!$A:$A,0))")
        ws.cell(row=row_idx, column=4, value=f"=INDEX(Cash_Flow!$G:$G,MATCH(Assumptions!$B$28+A{row_idx}-1,Cash_Flow!$A:$A,0))")
        ws.cell(row=row_idx, column=5, value=f"=C{row_idx}+D{row_idx}")
        ws.cell(row=row_idx, column=6, value=f"=MAX(0,B{row_idx}-D{row_idx})")


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
        ws.cell(row=row_idx, column=3, value=f"=IF($A{row_idx}>1,0,Assumptions!$B$22+IF($B{row_idx}>4,Assumptions!$B$27,0))")
        ws.cell(row=row_idx, column=4, value=f"=IF($U{row_idx}>(Assumptions!$B$22 + Assumptions!$B$23),0,IF($B{row_idx}>4,(Assumptions!$B$22+Assumptions!$B$27)-$U{row_idx},Assumptions!$B$22-$U{row_idx}))")
        ws.cell(row=row_idx, column=5, value=f"=Cash_Flow!$P{row_idx}")
        ws.cell(row=row_idx, column=6, value=r.cars_in_operation)
        ws.cell(row=row_idx, column=7, value=("=Assumptions!$B$29" if row_idx == 3 else f"=Fleet_Schedule!$I{row_idx}"))
        ws.cell(row=row_idx, column=8, value=f"=$F{row_idx}*Assumptions!$B$4")
        ws.cell(row=row_idx, column=9, value=f"=$G{row_idx}*Assumptions!$B$4")
        ws.cell(row=row_idx, column=10, value=f"=Cash_Flow!$J{row_idx}")
        ws.cell(row=row_idx, column=11, value=f"=Cash_Flow!$K{row_idx}")
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

    # Read fleet schedule source rows and recalculate operational fleet dynamics.
    fleet_rows = recalculate_fleet_rows(assumptions, read_fleet_schedule(fleet_ws, fleet_values_ws))
    # Build monthly income statement rows.
    income_rows = build_income_statement_rows(assumptions, fleet_rows)
    # Build monthly cash flow rows using required Interest/Loan Principal formulas.
    cash_rows = build_cash_flow_rows(assumptions, fleet_rows, income_rows)
    # Build loan amortisation rows from cash flow Interest/Loan Principal values.
    loan_rows = build_loan_rows(assumptions, len(fleet_rows), cash_rows)
    # Create month-to-loan lookup map.
    loan_map = map_loan_by_operating_month(assumptions, len(fleet_rows), loan_rows)
    # Build balance sheet rows from model outputs.
    balance_rows = build_balance_rows(assumptions, fleet_rows, cash_rows, income_rows, loan_map)

    # Execute validation checks before writing output.
    run_validations(income_rows, cash_rows, fleet_rows)

    # Write regenerated outputs back into workbook sheets.
    write_fleet_schedule(fleet_ws, fleet_rows, assumptions)
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
