"""Bridge between the web app and the existing execution/generate_payslip.py
CLI logic — imports its extraction functions directly so both the CLI and
the web app share one source of truth for the payroll calculations."""

import sys
from pathlib import Path

WEBAPP_DIR = Path(__file__).parent
PROJECT_ROOT = WEBAPP_DIR.parent
EXECUTION_DIR = PROJECT_ROOT / "execution"
sys.path.insert(0, str(EXECUTION_DIR))

import generate_payslip  # noqa: E402  (path insert must happen first)
import generate_ecr_file  # noqa: E402  (exposes build_ecr_lines(rows))

import db

EMPLOYEE_SHEET_NAME = "Employee_Details"


def _company_dict():
    settings = db.get_company_settings()
    return {
        "address": settings["address"],
        "cin": settings["cin"],
        "tan": settings["tan"],
        "pf_code": settings["pf_code"],
        "esi_code": settings["esi_code"],
        "pt_reg": settings["pt_reg"],
    }


def discover_salary_sheet(workbook_path):
    """Return the first sheet name starting with 'Salary Sheet' (the
    convention this company's payroll export already uses), or None."""
    import openpyxl
    wb = openpyxl.load_workbook(workbook_path, read_only=True)
    for name in wb.sheetnames:
        if name.strip().lower().startswith("salary sheet"):
            return name
    return None


def load_rows_for_excel(excel_path, salary_sheet_name):
    """Returns (month_label, rows, unmatched, warnings) — see
    generate_payslip.load_payslip_rows for the row/data shape."""
    company = _company_dict()
    return generate_payslip.load_payslip_rows(excel_path, salary_sheet_name, EMPLOYEE_SHEET_NAME, company)


def load_rows_for_month(month_row):
    return load_rows_for_excel(month_row["excel_path"], month_row["salary_sheet_name"])


def compliance_totals(rows):
    """Sum Employee/Employer PF, ESI, and PT across all employee rows for a month."""
    totals = {
        "employee_pf": 0, "employer_pf": 0,
        "employee_esi": 0, "employer_esi": 0,
        "pt": 0,
    }
    for _, data in rows:
        totals["employee_pf"] += data["employee_pf"] or 0
        totals["employer_pf"] += data["employer_pf"] or 0
        totals["employee_esi"] += data["employee_esi"] or 0
        totals["employer_esi"] += data["employer_esi"] or 0
        totals["pt"] += data["pt"] or 0
    return totals


def month_summary(rows):
    """Full aggregate snapshot for a month — persisted to db.month_summary at
    upload time so the dashboard doesn't need to re-parse the Excel file."""
    totals = compliance_totals(rows)
    total_ctc = sum(data["gross_earnings"] or 0 for _, data in rows)
    total_deductions = sum(data["total_deductions"] or 0 for _, data in rows)
    total_net_pay = sum(data["net_pay"] or 0 for _, data in rows)
    return {
        "employee_count": len(rows),
        "total_ctc": total_ctc,
        "total_deductions": total_deductions,
        "total_net_pay": total_net_pay,
        **totals,
    }


def build_ecr_text(rows):
    """Returns (text, included_count, skipped_count, warnings) for the
    monthly EPFO ECR 2.0 file — see generate_ecr_file.build_ecr_lines for
    the exact field rules."""
    settings = db.get_company_settings()
    raw = settings.get("eps_excluded_uans") or ""
    eps_excluded_uans = {u.strip() for u in raw.replace("\n", ",").split(",") if u.strip()}
    lines, skipped, warnings = generate_ecr_file.build_ecr_lines(rows, eps_excluded_uans)
    text = "\n".join(lines) + ("\n" if lines else "")
    return text, len(lines), skipped, warnings


def logo_path():
    """Absolute filesystem path to the uploaded logo (for PDF embedding),
    or None. Stored in the DB as a path relative to webapp/static/ so the
    same value also works directly with Flask's url_for('static', ...)."""
    settings = db.get_company_settings()
    rel = settings.get("logo_path")
    if not rel:
        return None
    full = WEBAPP_DIR / "static" / rel
    return str(full) if full.exists() else None
