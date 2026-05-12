# Payroll Payslip Generation

## Goal
Read payroll data from `Thota_Payroll 2025.xlsx` and generate individual PDF payslips for each employee, per month. Payslips should be issued on or before the 5th of each month.

## Company
**THOTA HOSPITALITY LLP**

## Input
- Excel file: `Thota_Payroll 2025.xlsx` (path provided by user)
- Sheet selection: User specifies which monthly sheet to process (e.g., "Salary Stmt-JAN26")
- If no sheet specified, process the latest available month

## Excel Structure
Each monthly salary sheet follows this layout:
- **Row 1**: Company name ("THOTA HOSPITALITY LLP")
- **Row 2**: Month title (e.g., "SALARY STATEMENT FOR THE MONTH OF JAN-2025")
- **Row 3**: Column headers
- **Row 4**: Sub-headers (PF/ESI labels under Employee/Employer Share)
- **Row 5+**: Employee data rows (until first empty Sl No)

### Column Mapping (varies slightly across sheets)
**Employee Info:** Sl No, Position, Emp Name, Designation, Department, UAN NO, PF NO, ESI NO, BRANCH, BANK ACCOUNT NO, DOJ

**Attendance:** Calendar days, Present day, Holiday, SL/EL/CL, LOP/Absent, Total, Paid Days, LOP Days

**Earnings:** Gross Salary/CTC, LOP Amount, Effective Gross, Basic, HRA, Conveyance Allowance, Special/Medical Allowance, Incentive, Over Time, Total Eff. Gross

**Deductions:** Employee PF, Employee ESI, Employer PF, Employer ESI, PT, TDS, Salary Advance, Total Deductions

**Net:** Arrears, Bonus/Attn Bonus, Net Salary, Mode of Pay, CTC

### Known Column Variations
- Earlier sheets (MAR25, APR25): Use "Gross" / "Spl. Allow" / "Attn Bonus"
- Later sheets (JULY25+): Use "Gross Salary" / "Medical Allowance" / "Spl. Allow" / "Incentive (Variable)" / "Bonus"
- The reader script handles both formats automatically

## Tools
1. `execution/payroll_reader.py` — Reads Excel, normalizes column names, returns structured employee data
2. `execution/payslip_pdf_generator.py` — Generates a professional PDF payslip per employee
3. `execution/payroll_pipeline.py` — Main pipeline: reads data → generates PDFs → outputs to folder

## Output
- Individual PDF payslips saved to `.tmp/payslips/<MONTH>/` folder
- Filename format: `Payslip_<EmpName>_<Month>.pdf`
- Console summary of generated payslips

## Usage
```bash
# Generate payslips for a specific month
python execution/payroll_pipeline.py --excel "C:\path\to\Thota_Payroll 2025.xlsx" --sheet "Salary Stmt-JAN26"

# Generate payslips for all months
python execution/payroll_pipeline.py --excel "C:\path\to\Thota_Payroll 2025.xlsx" --all

# Generate payslips for the latest month
python execution/payroll_pipeline.py --excel "C:\path\to\Thota_Payroll 2025.xlsx"
```

## Edge Cases
- Some employees have None/missing values for Position, Designation, Department — show "N/A" on payslip
- LOP days can be 0 or None — treat as 0
- Some deductions show '-' instead of 0 — normalize to 0
- Incentive and Over Time may be 0 or None
- Employee/Employer PF/ESI may be None if not applicable
- Sheet names are inconsistent (some have spaces, hyphens, different capitalization) — match flexibly
- Salary Advance may have values — include in deductions section
- Arrears and Bonus should appear if non-zero

## Scheduling Note
Payslips should be generated on or before the 5th of each month. The tool itself is on-demand; scheduling can be done via Windows Task Scheduler or cron.
