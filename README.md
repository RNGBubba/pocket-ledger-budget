# Pocket Ledger Budget

Pocket Ledger Budget is an original personal budget workbook template generated as a portable `.xlsx` file. It gives a household or freelancer three calm views:

- Dashboard: planned income, planned expenses, logged spending, and category rollups.
- Transactions: 200 rows for dates, descriptions, categories, amounts, type, and notes.
- Monthly Plan: income and expense targets with remaining and logged-expense formulas.

## Use

```bash
python budget_workbook.py pocket-ledger-budget.xlsx
```

Open the generated workbook in Excel, LibreOffice Calc, or another XLSX-compatible spreadsheet application. Enter positive amounts and choose `Expense` or `Income` in the Transactions sheet. Set targets on Monthly Plan; formulas recalculate when the workbook opens.

The template is educational and does not provide financial advice. It has no network calls, account sign-in, payment collection, macros, or external data connections.

## Verify

```bash
python -m pytest -q
python budget_workbook.py pocket-ledger-budget.xlsx --months 12
```

The integration smoke test checks the generated ZIP package and parses its workbook XML parts with the standard library.
