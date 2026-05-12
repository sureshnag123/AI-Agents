# HR Compliance Filing Automation — Thota Hospitality LLP

## Company Details
| Field | Value |
|-------|-------|
| Company | THOTA HOSPITALITY LLP |
| PAN | AAWFT5455G |
| EPFO Establishment ID | BGBNG3648720000 |
| PF Office | BANGALORE [BNG] |
| State | Karnataka |
| Payroll File | `C:/Users/User/Downloads/Thota_Payroll 2025.xlsx` |

## Goal
Each month, after payroll Excel is finalised, generate:
1. **EPFO ECR 2.0 file** — upload to EPFO Unified Portal for PF filing + challan payment

## Inputs
| Input | Details |
|-------|---------|
| Payroll Excel | `C:/Users/User/Downloads/Thota_Payroll 2025.xlsx` |
| Sheet name | `Salary Stmt-MAR26`, `Salary Stmt-FEB26`, etc. |
| EPFO Estab Code | `BGBNG3648720000` (also in `directives/thota_config.json`) |

## Payroll Excel Structure (Thota format — FY 2025-26)
- **Row 1**: Company name (THOTA HOSPITALITY LLP)
- **Row 2**: "SALARY STATEMENT FOR THE MONTH OF [MON]-[YEAR]"
- **Row 3**: Column headers
- **Row 4**: Sub-headers (PF/ESI under Employee/Employer share)
- **Rows 5 onwards**: Employee data (regular employees)
- **After TOTAL row**: Consultant row (Suresh N) — exclude from statutory filings

### Column mapping (0-indexed) — FY 2025-26 sheets
| Col | Field |
|-----|-------|
| 0 | Sl No |
| 1 | Position |
| 2 | Employee Name |
| 3 | UAN NO |
| 4 | ESI NO |
| 5 | DOJ |
| 6 | Calendar Days |
| 7 | Present Days |
| 8 | Holiday |
| 9 | SL/EL/CL |
| 10 | LOP / Absent |
| 11 | Total Days |
| 12 | Paid Days |
| 13 | LOP Days |
| 14 | Gross Salary |
| 15 | LOP Amount |
| 16 | Eff. Gross = Gross - LOP |
| 17 | Basic |
| 18 | HRA |
| 19 | Conveyance Allowance |
| 20 | Medical Allowance |
| 21 | Spl. Allowance |
| 22 | Incentive (Variable) |
| 23 | Over Time |
| 24 | Total Eff. Gross |
| 25 | Employee PF |
| 26 | Employee ESI |
| 27 | Employer PF |
| 28 | Employer ESI |
| 29 | PT |
| 30 | TDS |
| 31 | Salary Advance |
| 32 | Total Deductions |
| 33 | Arrears |
| 34 | Bonus |
| 35 | Net Salary |
| 36 | Mode of Pay |
| 37 | CTC |

Note: Old MAR25 sheet (`Salary stmt MAR25`) has a different layout with extra
columns for Dept/Designation. Use the FY26 format above for all FY 2025-26 sheets.

## UAN & ESI Reference Sheet
The workbook contains a `PF UAN & ESI IP` sheet with:
- Section 1 (rows 2–14): UAN → Member Name mapping
- Section 2 (rows 17+): ESI IP numbers (empe_name → empe_ip_number)

These match the UAN/ESI columns in the salary sheets but serve as a cross-reference.

## Step-by-step Process

### Step 1 — Generate EPFO ECR file
```bash
cd d:/Agents_Suresh/extracted/DOE-Framework

python execution/generate_ecr_thota.py \
  --file "C:/Users/User/Downloads/Thota_Payroll 2025.xlsx" \
  --sheet "Salary Stmt-MAR26" \
  --wage-month 03 \
  --wage-year 2026 \
  --establishment-id BGBNG3648720000
```
Output: `outputs/hr_compliance/ECR_202603_BGBNG3648720000.txt`
Also produces: `outputs/hr_compliance/PF_Summary_Thota_202603.xlsx`

### Step 2 — Upload ECR to EPFO Portal
1. Login: EPFO Unified Portal → Employer Login
2. Navigate: ECR / IW → Upload ECR
3. Upload the `.txt` file
4. Verify member-wise data on screen
5. Generate challan and pay online

## Changing the Month
Just update `--sheet`, `--wage-month`, `--wage-year`:

| Month | Sheet Name | --wage-month | --wage-year |
|-------|-----------|--------------|------------|
| March 2026 | Salary Stmt-MAR26 | 03 | 2026 |
| February 2026 | Salary Stmt-FEB26 | 02 | 2026 |
| January 2026 | Salary Stmt-JAN26 | 01 | 2026 |
| December 2025 | Salary Stmt-DEC25 | 12 | 2025 |
| November 2025 | Salary Stmt-NOV25 | 11 | 2025 |

## Compliance Deadlines
| Filing | Actual Due | Target (1 day early) |
|--------|-----------|----------------------|
| EPFO ECR + Challan | 15th of following month | **14th** |

**Policy: complete filing 1 day before the actual deadline.**

## PF Contribution Rules Applied
- Employee PF: 12% of EPF wages (capped at ₹15,000 → max ₹1,800/month)
- Employer EPF: 12% of EPF wages
  - EPS portion: 8.33% of EPS wages (capped ₹15,000) → max ₹1,250/month
  - EPF diff (3.67%): Employer EPF − EPS
- EDLI wages: same as EPF wages, capped at ₹15,000

## Known Data Issues
1. **Gokul (Purchase)** — PF deducted (₹1,170 in Mar 2026) but UAN is blank in the
   payroll sheet. He is **excluded from ECR** until UAN is added.
   → Action: Get UAN from Gokul or EPFO portal and fill col D in the sheet.
2. **Consultant row (Suresh N)** appears below the TOTAL row — no PF deducted,
   excluded from ECR automatically (sl_no > total row).
3. **Some employees have no ESI number** in the salary sheet (e.g. Nivedita,
   Nanda Kumar, Krupa N) — check `PF UAN & ESI IP` tab for cross-reference.

## Output Files
All files land in `outputs/hr_compliance/`:
| File | Purpose |
|------|---------|
| `ECR_YYYYMM_BGBNG3648720000.txt` | Upload to EPFO portal |
| `PF_Summary_Thota_YYYYMM.xlsx` | Internal PF register |
