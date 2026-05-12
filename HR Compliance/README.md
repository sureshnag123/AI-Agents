# HR Compliance — Fracktal Works Private Limited

Automated monthly statutory filings for PF (EPFO ECR 2.0), ESIC, and Professional Tax (Karnataka).

## What This Does

| Filing | Output File | Portal | Deadline |
|--------|-------------|--------|----------|
| EPFO ECR 2.0 | `outputs/ECR_YYYYMM_<EPFO_ID>.txt` | Unified Member Portal | 14th of next month |
| ESIC Monthly Contribution | `outputs/ESIC_YYYYMM.xlsx` | ESIC Portal | 14th of next month |
| Karnataka PT Deduction Register | `outputs/PT_Challan_YYYYMM.xlsx` | Karnataka PT Portal | 20th of next month |

## Setup

```bash
# Install dependencies
pip install -r requirements.txt

# Copy and fill in environment variables
cp .env.example .env
```

## Monthly Run

```bash
# 1. Parse payroll Excel
python execution/parse_payroll.py --file "C:/Users/User/Downloads/Payroll25-26_MarXX.xlsx" --sheet "Salary Sheet_Mar"

# 2. Generate EPFO ECR
python execution/generate_ecr.py --sheet "Salary Sheet_Mar" --wage-month 3 --wage-year 2026

# 3. Generate ESIC file
python execution/generate_esic_file.py --sheet "Salary Sheet_Mar" --wage-month 3 --wage-year 2026

# 4. Generate PT Challan
python execution/generate_pt_challan.py --sheet "Salary Sheet_Mar" --wage-month 3 --wage-year 2026
```

## Company

- **Company:** FRACKTAL WORKS PRIVATE LIMITED
- **EPFO Establishment ID:** PYKRP1426103000
- **ESIC Employer Code:** 49000552030001099
- **PT Registration No (RCN):** 356662288
- **State:** Karnataka
