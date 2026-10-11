# numbeers – household finances tracker

`Household_Finances.xlsx` is a budgeting, spending and savings workbook for two people
with their own accounts and a joint account. Open it in Excel, Google Sheets, Numbers or LibreOffice.

| Tab | What it's for |
|---|---|
| **How to use** | Quick guide and colour key |
| **Dashboard** | Pick a month: income, spending, savings and what's left for you, your partner, Joint and the household; every category against its budget; year overview and charts |
| **Where to Save** | Biggest discretionary categories, where you keep going over budget, subscription costs, and what a % cut would save over a year |
| **Settings** | Names, accounts (and who owns each), categories, monthly budgets, essential vs non-essential |
| **Transactions** | Your inputs: every incoming and outgoing (drop-downs for account, type and category) |
| **Regular Bills** | Your inputs: recurring income, bills and subscriptions, converted to monthly and yearly cost |
| **Savings / Savings Log** | Savings goals with targets and dates; log deposits and see progress and the amount needed each month |

Yellow cells are inputs and grey cells are calculated. Rows marked *Example – delete* show the format.
Clear their contents before you start using it for real.

## Rebuilding

`python3 build_workbook.py` (needs `openpyxl`) generates a fresh copy. **This overwrites
`Household_Finances.xlsx`**, so copy the file somewhere else before you enter real data in it.
