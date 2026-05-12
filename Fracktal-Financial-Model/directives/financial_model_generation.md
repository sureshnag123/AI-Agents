# Financial Model Generation — THOTA HOSPITALITY LLP

## Purpose

Generate comprehensive, partner-ready financial models and MIS reports from Trial Balance data. Two modes available:

1. **MIS Master (Formula-Linked)** — `mis_master_generator.py` — Trial Balance is the ONLY input; all reports auto-calculate via Excel formulas.
2. **Financial Model (Hardcoded)** — `financial_model_generator.py` — Generates reports from pre-existing P&L source Excel data.

## When to Use

- **Monthly MIS** → Use `mis_master_generator.py` with latest Trial Balance
- Board meeting / investor presentations
- Quarterly financial analysis
- Budget planning and variance analysis
- Financial health KPI tracking

---

## MODE 1: MIS Master Sheet (Formula-Linked — Recommended)

### Overview

Creates a master Excel workbook where the **TB (Trial Balance) sheet is the only input**. All other sheets (P&L, Balance Sheet, Cash Flow, KPIs, Dashboard) use **Excel formulas** — zero hardcoded numbers. When user updates the TB, all reports auto-update.

### Inputs

| Input | Required | Description |
|-------|----------|-------------|
| Trial Balance Excel | ✅ | Tally-exported TB file (e.g., `TrialBal 31.01.26.xlsx`) |
| Output path | ❌ | Defaults to `.tmp/MIS_Master_<timestamp>.xlsx` |
| Company name | ❌ | Defaults to "THOTA HOSPITALITY LLP" |

### Command

```bash
python execution/mis_master_generator.py --source "path/to/TrialBal.xlsx"
python execution/mis_master_generator.py --source "path/to/TrialBal.xlsx" --output "MIS_Master.xlsx"
```

### Output Sheets

| Sheet | Type | Description |
|-------|------|-------------|
| `TB` | INPUT | Monthly cumulative Trial Balance values (user fills this) |
| `PnL` | FORMULA | Monthly P&L with quarterly and annual totals |
| `Balance Sheet` | FORMULA | Monthly BS snapshots with balance check |
| `Cash Flow` | FORMULA | Indirect method — derived from P&L and BS |
| `KPIs` | FORMULA | Profitability, liquidity, operational ratios |
| `Dashboard` | FORMULA | Executive summary with charts |
| `Quarterly PnL` | FORMULA | Compact quarterly view for board presentations |
| `Budget vs Actuals` | FORMULA | Variance analysis with budget input cells |

### User Workflow (Monthly)

1. Export cumulative Trial Balance from Tally (Apr 1 to current month end)
2. Open the MIS Master Excel
3. Go to `TB` sheet → paste values in the month column (D=Apr, E=May, ... O=Mar)
4. All report sheets auto-update — no manual changes needed
5. Review Dashboard → present to management

### TB Input Convention

- All values entered as **positive numbers**
- Debit-nature accounts (Assets, Expenses): enter the debit balance
- Credit-nature accounts (Liabilities, Revenue): enter the credit balance
- Net accounts (e.g., Duties & Taxes with both Dr/Cr): enter Credit - Debit
- P&L items are **cumulative** (Tally default). Monthly extraction is done via formulas.
- BS items show **closing balance** for that month.

### Account Mapping (33 accounts)

| Category | Accounts |
|----------|----------|
| Equity | Partner's Capital, P&L A/c (Opening) |
| Loans | Akshatha, Tejas, TVS |
| Current Liabilities | Duties & Taxes, Creditors, Salary Payable, Reimbursements |
| Fixed Assets | Intangible, Tangible |
| Current Assets | Advances, Debtors, Cash, Bank |
| Revenue | Hospitality Services, Studio Rental |
| COGS | Kitchen, Decor, Studio Purchase |
| Direct Expenses | Fuel, Transportation |
| Indirect Income | Discount Received |
| OPEX | Admin, HR, Finance, Marketing, Professional, R&M, etc. |
| Depreciation | User-input depreciation charge |

### Edge Cases Learned

- **Duplicate TB names**: Tally exports can have "Thota Kitchen" at both group and sub-item level. Script uses FIRST occurrence (group subtotal).
- **Credit-nature expenses**: Rates & Taxes and Round Off are credit balances. Formulas NEGATE these in the P&L so they reduce total OPEX.
- **Net accounts**: Some accounts have both Debit and Credit columns in TB. Use NET values on the TB input sheet.
- **April formula**: First month = cumulative value directly (no subtraction needed).
- **Depreciation**: Not always in TB. Added as separate input row for user to enter.

---

## MODE 2: Financial Model (Hardcoded from P&L Source)

### Inputs

| Input | Required | Description |
|-------|----------|-------------|
| Source Excel file | ✅ | Raw financial data (e.g., `FW_P&L_2025-26 - V2.xlsx`) |
| Company name | ❌ | Defaults to "THOTA HOSPITALITY LLP" |
| Output path | ❌ | Defaults to `.tmp/Financial_Model_<timestamp>.xlsx` |

### Source File Expected Sheets

The source Excel file should contain some or all of these sheets:

| Sheet Name | Content |
|-----------|---------|
| `Performance Summary` | Quarterly P&L (Revenue, COGS, GP, EBITDA, PAT) |
| `P&L Detailed` | Monthly P&L with line-item detail |
| `Schedule_OPEX` | Operating expense breakdown (HR, Admin, Marketing, etc.) |
| `BalanceSheet Summary` | Assets, Liabilities, Equity |
| `Cashflow` | Monthly cash inflows/outflows |
| `Fundflow_Post Investment` | Fund flow with investment tracking |
| `Segment_wise_Revenue` | Revenue by segment (Morning Glory, Al Fresco, etc.) |
| `FC` | Fixed cost structure (payroll, rent, etc.) |
| `Revenue Projection` | Sales pipeline deals |

## Execution

### Script

```bash
python execution/financial_model_generator.py --source "<path_to_source_excel>" --company "THOTA HOSPITALITY LLP"
```

### Options

| Flag | Description |
|------|-------------|
| `--source` | Path to the source Excel file (REQUIRED) |
| `--output` | Custom output path (optional) |
| `--company` | Company name for headers (optional) |

## Output

A single `.xlsx` workbook with these sheets:

### 1. Performance Summary
- Quarterly P&L overview (Q1-Q4 + FY Total)
- Revenue, COGS, Gross Profit, OPEX, EBITDA, PAT
- Margin percentages (GP%, EBITDA%, PAT%)
- **Charts:** Revenue/Profitability bar chart + Margin trend line chart

### 2. P&L Statement (Detailed)
- 12-month detailed P&L (Apr–Mar)
- Revenue segments: Morning Glory, Al Fresco, Sunset Soiree, Studio
- COGS breakdown: Purchase Cost, Direct Expenses
- OPEX breakdown: HR, Admin, Marketing, Professional Services, R&M

### 3. Balance Sheet
- Liabilities: Capital Account, Loans, Current Liabilities, P&L Account
- Assets: Fixed Assets (Tangible/Intangible), Current Assets (Debtors, Cash, Bank)
- Professional formatting with sub-totals

### 4. Cash Flow Statement
- Monthly Opening/Closing balances
- Inflows and Outflows with categories
- Net cash position
- Negative values highlighted in red

### 5. Revenue Analysis
- Monthly segment-wise revenue table
- Revenue mix percentages
- Revenue pipeline (top deals with probability and expected revenue)
- **Charts:** Revenue mix pie chart

### 6. KPIs & Financial Ratios
- **Profitability:** GP Margin, EBITDA Margin, PAT %, ROA, ROE
- **Liquidity:** Current Ratio, Quick Ratio, Cash Ratio, Working Capital
- **Operational:** Monthly Breakeven, DSO, Cash Runway, Debt/Equity, Asset Turnover
- Quarterly trend table
- Benchmarks and notes for each ratio

### 7. Budget vs Actuals
- Actuals (Q1-Q3) vs Budget/Projection (Q4)
- YoY growth target (30%)
- Variance analysis (amount and %)
- **Charts:** Actuals vs Target clustered bar chart

### 8. Fund Flow Statement
- Previous vs Current period comparison
- Cash receipts, COGS, Operating Expenses, Other Payments breakdown
- Cash position detail (Bank, Cash, Overdraft, FD)
- Period-over-period change calculation

## Styling

All sheets use professional formatting:
- **Header:** Navy blue (#2F5496) with white text
- **Sub-headers:** Light blue (#D6E4F0)
- **Totals:** Light green (#E2EFDA) with double border
- **Currency:** Indian Rupee format (₹#,##0)
- **Percentages:** 0.0% format
- **Negatives:** Red color (#C00000)
- **Font:** Calibri throughout

## Edge Cases & Notes

1. **#REF! values:** Some source cells may have broken references — these are safely converted to 0
2. **Projected months:** FEB and MAR may be projections (marked with "P" suffix)
3. **Empty sheets:** If a source sheet is missing, that section is skipped gracefully
4. **Large files:** Revenue-Detailed can have 1000+ invoice rows — pipeline shows top 25
5. **Partner columns:** Q4 data in Performance Summary includes projections for partner review

## Learnings

- Source file uses mixed date formats (datetime objects vs strings)
- Some cells contain non-breaking spaces (`\xa0`) that need to be handled
- The `Cashflow Projection` sheet has many `#REF!` errors — use `Cashflow` sheet instead
- Balance Sheet may not fully balance due to "Excess Profit not shown in books" line
- Revenue segments: Morning Glory (breakfast), Al Fresco (outdoor dining), Sunset Soiree (events)

## Dependencies

```
openpyxl>=3.1.0
```
