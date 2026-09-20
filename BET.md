# Bet: Pocket Ledger Budget

## Offer

An original personal budget workbook for households and freelancers who want a focused monthly plan plus a lightweight transaction log. The generated XLSX is usable in common spreadsheet software without add-ins, macros, network access, or account setup.

## Price

Suggested first-dollar offer: $12 for the base workbook, or $39 for a customized version with a buyer's categories and planning horizon. No payment rail was used in this run.

## 30-day path

1. Publish the new public repository and README.
2. Share it only in relevant, non-spam spreadsheet and personal-finance communities where templates are welcome.
3. Offer paid customization of categories, number of months, labels, or a branded cover.
4. Turn repeated customization requests into documented variants.

## Human click

A human must approve external promotion or a customer sale. This artifact does not collect credentials, send messages, or make financial decisions for the user.

## Artifact

- `budget_workbook.py` — dependency-free XLSX generator using the standard library.
- `pocket-ledger-budget.xlsx` — generated workbook with Dashboard, Transactions, and Monthly Plan sheets.
- `tests/test_budget_workbook.py` — real ZIP/XML integration tests covering sheets, formulas, and validation.
- `README.md` — usage, safety, and verification notes.
- `artifacts/pytest.txt` and `artifacts/qa-smoke.txt` — execution evidence.

## Verification

- `python -m pytest -q` — 3 passed.
- `python budget_workbook.py pocket-ledger-budget.xlsx --months 12` — generated a 9,256-byte workbook.
- ZIP/XML smoke check — workbook package and all worksheet XML parts parsed successfully.
- DoneMeans receipt `t_af4c5f846fe0` verified and packed at `/home/vboxuser/projects/donemeans/packs/t_af4c5f846fe0.json`; status `self-paid`.

## GitHub

Published as a brand-new public repository: https://github.com/RNGBubba/pocket-ledger-budget

The pushed HEAD is `2449ffece1b11acb35d21b3ffc10867662699a95`.

## Constraints

Original work only. No old RNGBubba product names, private repositories, secrets, spending, games, or external services are included.
