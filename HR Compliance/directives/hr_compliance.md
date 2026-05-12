# HR Compliance Filing Automation — Fracktal Works Pvt Ltd

## Goal
Each month, after the payroll Excel is finalised, generate:
1. **EPFO ECR 2.0 file** — upload to EPFO Unified Portal
2. **ESIC contribution file** — upload to ESIC Employer Portal
3. **Karnataka PT Challan** — pay on Karnataka PT Portal

## Inputs
| Input | Details |
|-------|---------|
| Payroll Excel | `C:/Users/User/Downloads/Payroll25-26_March26.xlsx` (or newer file each month) |
| Sheet name | `Salary Sheet_Mar`, `Salary Sheet_Apr`, etc. |
| State | Karnataka |
| EPFO Estab Code | Fill in `.env` as `EPFO_ESTAB_CODE` |
| ESIC Employer Code | Fill in `.env` as `ESIC_EMPLOYER_CODE` |

## Payroll Excel Structure (Fracktal Works format)
- **Row 1**: Company name
- **Row 2**: "SALARY STATEMENT FOR THE MONTH OF [MONTH]- [YEAR]"
- **Row 3**: Column headers
- **Row 4**: Sub-headers (PF/ESI under Employee/Employer share)
- **Rows 5+**: Employee data

### Sections within each sheet
| Section label in col C | Section type | PF/ESI/PT |
|------------------------|--------------|-----------|
| (first group, no label) | `regular` | Yes |
| `DIRECTOR REMUNERATION` | `director` | PF yes, PT ₹300 |
| `INTERNSHIPS/ Under Contract` | `intern` | No |
| `PROFESSIONAL SERVICE` | `professional_service` | No |

### Column mapping (0-indexed)
- 4: UAN, 5: PAN, 6: ESI Insurance No, 7: Bank Account, 8: DOJ
- 15: Paid Days, 16: LOP Days, 17: CTC, 19: Eff. Gross, 20: Basic
- 27: Total Eff. Gross, 28: Employee PF, 29: Employee ESI
- 30: Employer PF, 31: Employer ESI, 32: PT, 33: TDS
- 38: Net Salary

## Step-by-step Process

### Step 1 — Parse payroll
```bash
cd d:/Agents_Suresh/extracted/DOE-Framework

python execution/parse_payroll.py \
  --file "C:/Users/User/Downloads/Payroll25-26_March26.xlsx" \
  --sheet "Salary Sheet_Mar" \
  > .tmp/payroll_Salary_Sheet_Mar.json
```
Check stderr for validation warnings (missing UAN, ESI numbers, etc.)

### Step 2 — Generate EPFO ECR file
```bash
python execution/generate_ecr.py \
  --json .tmp/payroll_Salary_Sheet_Mar.json \
  --wage-month 03 \
  --wage-year 2026 \
  --establishment-id KARBN001234
```
Output: `outputs/hr_compliance/ECR_202603_KARBN001234.txt`
Upload this .txt file to: EPFO Unified Portal → ECR / IW > ECR Upload

### Step 3 — Generate ESIC file
```bash
python execution/generate_esic_file.py \
  --json .tmp/payroll_Salary_Sheet_Mar.json \
  --wage-month 03 \
  --wage-year 2026 \
  --employer-code <ESIC_CODE>
```
Output: `outputs/hr_compliance/ESIC_202603.xlsx`
Upload to: esic.in → Employer Login → Monthly Contribution

**Known gap**: ESI Insurance Number column (col G) is currently blank for all
employees. Before ESIC can be filed electronically, each eligible employee
needs their ESIC Insurance Number filled in the payroll Excel.
Workaround: File manually on ESIC portal for now; add ESI numbers to Excel.

### Step 4 — Generate PT Challan
```bash
python execution/generate_pt_challan.py \
  --json .tmp/payroll_Salary_Sheet_Mar.json \
  --wage-month 03 \
  --wage-year 2026 \
  --verify-slabs
```
Output: `outputs/hr_compliance/PT_Challan_202603.xlsx`
Pay on: https://ptax.karnataka.gov.in → Pay Tax Online

## Compliance Deadlines

| Filing | Actual Due | Target (1 day early) |
|--------|-----------|----------------------|
| EPFO ECR + Challan | 15th of following month | **14th** |
| ESIC Contribution | 15th of following month | **14th** |
| Karnataka PT (monthly filer) | 20th of following month | **19th** |

**Fracktal has ~24 employees → monthly PT filer.**
**Policy: complete all filings 1 day before the actual deadline.**

## Karnataka PT Slabs (FY 2025-26)
| Monthly Gross | PT per Month |
|---------------|-------------|
| Up to ₹15,000 | Nil |
| ₹15,001 – ₹24,999 | ₹150 |
| ₹25,000 and above | ₹200 |

Note: The ₹300 seen for Directors in the Excel is the employer PT registration
fee, not the employee deduction slab. The scripts use Excel values as
authoritative.

## PF Contribution Rules Applied
- Employee PF: 12% of EPF wages (capped at ₹15,000 basic = max ₹1,800/month)
- Employer EPF: 12% of EPF wages
  - EPS portion: 8.33% of EPS wages (capped ₹15,000) → max ₹1,250/month
  - EPF diff: 3.67% of EPF wages
- Employees with CTC ≥ ₹15,000 basic: PF capped at ₹1,800

## ESIC Rules Applied
- ESI threshold: gross wages ≤ ₹21,000/month
- Employee: 0.75% of gross wages
- Employer: 3.25% of gross wages
- Only Lakshmipathy is currently ESI-eligible (CTC ₹21,672 — borderline; verify)

## Known Data Issues (as of March 2026)
1. **ESI Insurance Numbers missing** for all employees (col G = blank)
   → Action: Get ESIC insurance numbers from ESIC portal and fill col G
2. **Some UAN numbers have leading spaces** (e.g., Impa HL: " 100515384354")
   → parse_payroll.py strips these automatically
3. **Vijay Raghav Varada (CEO)** has no UAN and zero PF
   → Likely opted out or above threshold; confirm with CA

## Output Files
All files land in `outputs/hr_compliance/`:
| File | Purpose |
|------|---------|
| `ECR_YYYYMM_<estab>.txt` | Upload to EPFO portal |
| `PF_Summary_YYYYMM.xlsx` | Internal PF register |
| `ESIC_YYYYMM.xlsx` | Upload to ESIC portal |
| `PT_Challan_YYYYMM.xlsx` | Karnataka PT register + challan |

## Adding New Months
Just change `--sheet` and `--wage-month` / `--wage-year`. All scripts are
stateless — re-run anytime.

## Edge Cases
- **FNF employees**: Sheets like `Geo_FNF`, `Rohit_FNF` are full-and-final
  settlements. Run separately if needed for that month's ECR.
- **Mid-month joiners**: LOP days will reduce EPF wages proportionally
  (already reflected in Basic column).
- **Overtime**: Included in Total Eff. Gross (EPFO uses gross wages including OT).
