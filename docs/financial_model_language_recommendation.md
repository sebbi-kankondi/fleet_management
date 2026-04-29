# Financial projection model language recommendation (Molenzicht CC)

## Inputs reviewed
- `data/financial_projections.xlsx`
- `data/molenzicht_business_plan.pdf` (metadata and file presence only in this environment)

## Observations from `financial_projections.xlsx`
The workbook is already structured as an assumptions-driven model with dedicated output sheets:

- Assumptions
- Fleet_Schedule
- Cash_Flow
- Income_Statement
- Balance_Sheet
- Loan_Amortisation

Cell/formula density indicates a highly formula-linked spreadsheet model:

- Assumptions: 128 cells, 7 formulas
- Fleet_Schedule: 683 cells, 656 formulas
- Cash_Flow: 1061 cells, 958 formulas
- Income_Statement: 1546 cells, 1140 formulas
- Balance_Sheet: 1415 cells, 1380 formulas
- Loan_Amortisation: 342 cells, 324 formulas

## Recommendation
**Python is the best primary language for generation and maintenance of this model**, with Excel kept as the reporting surface.

### Why Python over R for this use-case
1. **Excel-first automation fit**  
   Python has stronger ecosystem support for creating/updating multi-sheet Excel workbooks with formulas, styles, named ranges, and templates.
2. **Operational maintainability**  
   For production finance models, Python is typically easier to integrate into scheduled jobs, APIs, and CI pipelines.
3. **Testing and model governance**  
   Python tooling makes regression tests and assumption-sensitivity tests straightforward to automate.
4. **Future extensibility**  
   Integrating with ERP/accounting systems and cloud workflows is typically simpler with Python.

## If you strongly prefer R
R can absolutely work, especially with `readxl`, `openxlsx`, and tidyverse-based pipelines.  
A pragmatic approach is:

- Keep business logic in an R script/package.
- Recalculate projection tables from `Assumptions`.
- Write outputs back to the five target sheets in `financial_projections.xlsx`.

This is viable, but for long-term robustness of Excel-centric financial model engineering, Python remains the safer default.

## Implementation pattern (language-agnostic)
1. Read `Assumptions` into a typed parameter object.
2. Recompute each output in dependency order:
   - Fleet_Schedule
   - Loan_Amortisation
   - Income_Statement
   - Cash_Flow
   - Balance_Sheet
3. Validate accounting identities (e.g., Assets = Liabilities + Equity).
4. Write all output sheets atomically.
5. Run regression checks against a baseline scenario.
