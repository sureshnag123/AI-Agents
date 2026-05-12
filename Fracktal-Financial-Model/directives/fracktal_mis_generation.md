# MIS Master Sheet Generation — Fracktal Works Private Limited

## Purpose

Generate a comprehensive, investor/director-ready MIS Master workbook from Tally-exported Trial Balance data. All report numbers flow from **Excel formulas** — zero hardcoded values. When TB is updated with a new month, all reports auto-update.

## When to Use

- Monthly MIS reporting to Investors & Directors
- Board meeting / investor presentations
- Quarterly financial analysis & review
- Budget planning and variance analysis

---

## Script

```bash
python execution/fracktal_mis_generator.py --source "path/to/TrialBalance.xls"
python execution/fracktal_mis_generator.py --source "path/to/TrialBalance.xls" --output "MIS_Output.xlsx"
```

## Inputs

| Input                      | Required | Description                                             |
| -------------------------- | -------- | ------------------------------------------------------- |
| Trial Balance (.xls/.xlsx) | ✅       | Tally-exported Trial Balance file                       |
| Output path                | ❌       | Defaults to `.tmp/Fracktal_MIS_Master_<timestamp>.xlsx` |

### Source File Format (Tally Export)

The Tally Trial Balance export must have:

- **Rows 0-6**: Company header info (name, address, CIN, etc.)
- **Row 7**: Company name repeated across month columns
- **Row 8**: Period headers (e.g., "1-Apr-25 to 30-Apr-25")
- **Row 9**: "(in INR)" labels
- **Row 10**: "Closing Balance" labels
- **Row 11**: "Debit" / "Credit" headers
- **Rows 12-79**: Account data (68 accounts with Dr/Cr per period)

Each monthly period has 2 columns: Debit and Credit.

## Output Sheets (8 tabs)

| Sheet                 | Type    | Description                                                       |
| --------------------- | ------- | ----------------------------------------------------------------- |
| `TB`                  | INPUT   | Trial Balance — paste monthly Tally data here (yellow cells)      |
| `P&L`                 | FORMULA | Monthly P&L Statement — all cells are `=TB!{cell}` references     |
| `Cash Flow`           | FORMULA | Monthly cash inflows / outflows / opening-closing balances        |
| `OPEX Schedule`       | FORMULA | Detailed breakup of every indirect expense item                   |
| `Balance Sheet`       | FORMULA | Monthly BS with Liabilities vs Assets + difference check          |
| `Performance Summary` | FORMULA | Quarterly aggregation (Q1-Q4) with FY totals for board review     |
| `KPIs`                | FORMULA | Key financial ratios with benchmarks & plain-English explanations |
| `Dashboard`           | CHARTS  | Revenue vs EBITDA bar chart, expense pie, GP trend line           |

Every sheet includes plain-English explanatory notes at the bottom for non-finance readers.

### TB Sheet Structure

| Column | Content                                             |
| ------ | --------------------------------------------------- |
| A      | Particulars (Account Name)                          |
| B      | APR (Net value: Dr-Cr or Cr-Dr depending on nature) |
| C      | MAY                                                 |
| ...    | ...                                                 |
| M      | MAR                                                 |
| N      | FY Total (=SUM(B:M))                                |

- **Yellow cells**: Input cells for monthly pasting
- **Blue rows**: Section headers (Revenue, COGS, Direct Expenses, OPEX, Balance Sheet Items)
- **Green rows**: Calculated totals (SUM formulas)
- **Dark blue rows**: Grand totals (EBITDA, PBT, PAT)
- Data starts at Row 4 after title/subtitle/headers
- Net values are stored (Dr-Cr for expenses/assets, Cr-Dr for revenue/liabilities)

### P&L Structure (CA Standard Format)

```
I. REVENUE FROM OPERATIONS
   Sale of Products | Sale of Services | Export Sales | Printsticks
   Total Revenue from Operations

II. OTHER INCOME
   Interest Income | Discount Received | Other Income
   Total Other Income

III. TOTAL INCOME (I + II)

IV. EXPENSES
  A. Cost of Goods Sold (COGS)
     Opening Stock (manual input — physical stock at start of month)
     Purchase of RM (Domestic) | Import of RM | Other Purchases
     Total Purchases
     Less: Closing Stock (manual input — physical stock at end of month)
     **Total COGS = Opening Stock + Purchases − Closing Stock**
  B. Manufacturing & Direct Expenses
     Salaries (Production) | Overtime | Electricity | Freight Inward | etc.
  GROSS PROFIT (with margin %)

V. OPERATING EXPENSES
  C. Employee Benefit Expenses (Payroll)
  D. Selling & Distribution (Advertisement, Freight Outward)
  E. Administration & General (Office, Professional, Travel, R&D, etc.)
  TOTAL OPERATING EXPENSES

EBITDA (with margin %)
Less: Finance Cost
PROFIT BEFORE TAX (with margin %)
Less: Tax Provision (manual input)
NET PROFIT AFTER TAX (with margin %)
```

### Balance Sheet Structure

```
EQUITY & LIABILITIES
  I. Shareholders' Funds (Share Capital, Reserves, P&L A/c)
  II. Non-Current Liabilities (Unsecured Loans)
  III. Current Liabilities (Creditors, Duties & Taxes, Salaries, Reimbursements)
  TOTAL EQUITY & LIABILITIES

ASSETS
  I. Non-Current Assets (Fixed Assets, Deferred Tax, Deposits)
  II. Current Assets (Stock, Debtors, Cash, Bank, Advances, TDS etc.)
  Suspense Account
  TOTAL ASSETS

BALANCE CHECK (should be 0)
```

### Dashboard Metrics

- Revenue & Profitability (Revenue, GP, EBITDA, PBT, PAT with margins)
- Balance Sheet Snapshot (SH Funds, FA, CA, CL)
- Key Ratios (Current Ratio, Debt-to-Equity, Working Capital, Burn Rate)
- Expense Breakdown (Material, Direct, Employee, Selling, Admin, Finance)
- Charts: Revenue vs EBITDA bar/line chart

## User Workflow (Monthly Update)

1. Export monthly Trial Balance from Tally (single month period)
2. Open the MIS Master Excel workbook
3. Go to **TB sheet** → paste Debit column values into the correct month's Dr column
4. Paste Credit column values into the correct month's Cr column
5. All report sheets auto-update — verify the BS Balance Check = 0
6. Review Dashboard → present to investors/directors

## Account Mapping (68 Accounts)

| Main Group          | Accounts (Sub-headings)                                                                                                                                            |
| ------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| Capital Account     | Reserves & Surplus, Share Capital (Issued & Paid Up)                                                                                                               |
| Loans (Liability)   | Unsecured Loans                                                                                                                                                    |
| Current Liabilities | Duties & Taxes, Sundry Creditors, Reimbursements, Salaries                                                                                                         |
| Fixed Assets        | Tangible, Intangible, Machinery, Fracktal Studio, SLS PAMM, WIP                                                                                                    |
| Current Assets      | Stock, Deposits, L&A, Debtors, Cash, Bank, Deferred, DTA, TDS                                                                                                      |
| Suspense            | Suspense Account                                                                                                                                                   |
| Revenue             | Export Sales, Printsticks, Sale of Products, Sale of Service                                                                                                       |
| Purchases           | Import of RM, Purchase of RM                                                                                                                                       |
| Direct Expenses     | Other Purchases, Discount, Electricity (2), Freight In, Loading, Overtime, Salaries (Production)                                                                   |
| Indirect Incomes    | Discount Received, Interest Income, Other Income                                                                                                                   |
| Indirect Expenses   | Advt/Marketing, Finance Cost, Office/Admin, Payroll, Rates & Taxes, R&D, Travel, Forex, Freight Out, Professional, Razorpay, Round Off, Tender Fee, Write Off/Back |
| Profit & Loss       | Profit & Loss A/c (accumulated balance)                                                                                                                            |

## Formula Reference

- **P&L Revenue items**: `=IFERROR(TB!{CreditCol}{row},0)` — pulls Credit side from TB
- **P&L Expense items**: `=IFERROR(TB!{DebitCol}{row},0)` — pulls Debit side from TB
- **BS Items (Net)**: `=IFERROR(TB!Dr,0)-IFERROR(TB!Cr,0)` for assets; reverse for liabilities
- **Quarterly/FY**: SUM formulas aggregating monthly columns; BS uses Q-end closing balance
- **All reports**: No hardcoded numbers — change TB values and everything recalculates

## Edge Cases & Notes

- **Net accounts**: Some accounts have both Debit and Credit (e.g., Sundry Debtors, Duties & Taxes). BS formulas compute Net = Dr-Cr for assets, Cr-Dr for liabilities.
- **Round Off & Write Off**: These can have credit nature. Formulas use Net (Dr-Cr) to handle sign correctly.
- **Deferred Tax Asset**: Shows Credit balance in TB but is classified under Current Assets. The Net formula handles this correctly (will show negative, reducing total assets).
- **Tax Provision**: Manual input row in P&L (yellow highlighted). Enter monthly tax amounts if applicable.
- **Group totals**: TB includes group total rows (e.g., "Capital Account", "Current Assets") for verification. Report formulas reference individual sub-items, not group totals.
- **New ledgers**: If new accounts are added in Tally, add matching rows in TB sheet and update relevant report SUM ranges.
- **FY Year**: Currently hardcoded as 2025-26; update in script if needed for future years.

## Dependencies

- Python 3.12+
- pandas, openpyxl, xlrd (for .xls files)

## Learnings

- Tally exports `.xls` files that are actually `.xlsx` format internally — use `openpyxl` engine with pandas
- For BS items: each monthly column shows closing balance at month-end; cumulative column = December closing
- For P&L items: each monthly column shows that month's activity only; cumulative = sum of all months
- Capital Account / Reserves show same values across all months (no transactions)
- P&L A/c row in TB is a transfer account — shows opening accumulated P&L for each monthly period
- The Grand Total (Dr = Cr) always balances in the Tally TB export
