#!/usr/bin/env python3
"""
EPFO ECR 2.0 File Generator — Fracktal Works Private Limited

Generates the monthly PF Electronic Challan cum Return (ECR 2.0) upload file
directly from the real payroll Excel (a "Salary Sheet_<Month> <Year>" sheet
+ "Employee_Details"), reusing generate_payslip.load_payslip_rows() — the
same extraction/reconciliation logic the payslips and the webapp's
Compliance tab already use — so the ECR file, the payslips, and the
Compliance totals always agree on every number.

ECR 2.0 field format (one line per PF-enrolled employee):
    UAN #~# MEMBER_NAME #~# GROSS_WAGES #~# EPF_WAGES #~# EPS_WAGES #~#
    EDLI_WAGES #~# EPF_CONTRIBUTION #~# EPS_CONTRIBUTION #~# EPF_EPS_DIFF #~#
    NCP_DAYS #~# REFUND_OF_ADVANCES

Rules applied (per EPFO compliance standard):
  - Delimiter is "#~#" between every field — NOT a single "~"
  - No establishment header line in the file (entered separately on the
    EPFO Unified Portal against the Establishment Code / TRRN)
  - Only employees with a valid 12-digit UAN AND a nonzero Employee PF
    deduction on their payslip are included — this naturally excludes
    interns/professional-service consultants and anyone marked PF-exempt,
    without needing a separate "PF Applicable" flag column
  - GROSS_WAGES = EPF_WAGES (PF-eligible wages, not full CTC gross)
  - EPF_WAGES = MIN(Basic, ₹15,000); EPS_WAGES = EDLI_WAGES = EPF_WAGES
  - EPF_CONTRIBUTION is taken directly from the payslip's own Employee PF
    figure (not recomputed) so the ECR file always matches what was
    actually deducted and shown to the employee
  - EPS_CONTRIBUTION = ROUND(8.33% x EPS_WAGES), hard-capped at ₹1,250
  - EPF_EPS_DIFF = EPF_CONTRIBUTION - EPS_CONTRIBUTION
  - NCP_DAYS = LOP days for the month
  - REFUND_OF_ADVANCES = 0 (not tracked separately)

Usage:
    python generate_ecr_file.py --workbook ".tmp/Payroll26-27_FracktalWorks_JUNE.xlsx" \
        --salary-sheet "Salary Sheet_June 2026"
"""

import sys
import argparse
from pathlib import Path
from datetime import datetime

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")

SCRIPT_DIR = Path(__file__).parent.resolve()
PROJECT_ROOT = SCRIPT_DIR.parent
TMP_DIR = PROJECT_ROOT / ".tmp"
sys.path.insert(0, str(SCRIPT_DIR))

import generate_payslip  # noqa: E402  (path insert must happen first)

DELIMITER = "#~#"
EPF_WAGE_CEILING = 15000
EPS_RATE = 0.0833
EPS_CONTRIBUTION_CAP = 1250


def _norm_uan(value):
    return "".join(ch for ch in str(value or "").strip() if ch.isdigit())


def build_ecr_lines(rows, eps_excluded_uans=None):
    """rows is the (row_id, data) list from generate_payslip.load_payslip_rows.
    `eps_excluded_uans` is an optional set of 12-digit UAN strings for members
    the EPFO portal has recorded as EPS-excluded (e.g. joined after 1 Sept
    2014 on wages above the then-ceiling under a prior employer) — their EPS
    wages/contribution are reported as zero, with the full EPF contribution
    going to the EPF-share diff instead. Returns (lines, skipped_count, warnings).

    Warnings flag employees who actually have a nonzero Employee PF deduction
    on their payslip but no valid 12-digit UAN, so that PF amount is silently
    missing from this ECR file and needs manual resolution before filing
    (e.g. a very recent joiner who exited before their UAN was allotted)."""
    eps_excluded_uans = eps_excluded_uans or set()
    lines = []
    skipped = 0
    warnings = []
    for _row_id, data in rows:
        uan = _norm_uan(data.get("uan"))
        employee_pf = data.get("employee_pf") or 0
        if len(uan) != 12 or not employee_pf:
            skipped += 1
            if employee_pf:
                warnings.append(
                    f"{data['name']}: Rs {employee_pf:,.2f} Employee PF deducted this month but no valid "
                    f"12-digit UAN on file (raw value: {data.get('uan')!r}) — NOT included in this ECR file."
                )
            continue

        # EPF wage base per the company's confirmed PF policy: Basic +
        # Conveyance + Medical Allowance + Special Allowance only; if that
        # total exceeds Rs 15,000 the statutory ceiling applies, otherwise
        # the actual total is used as-is.
        pf_wage_components = (
            (data.get("basic") or 0) + (data.get("conveyance") or 0)
            + (data.get("medical_allowance") or 0) + (data.get("special_allowance") or 0)
        )
        epf_wages = min(int(round(pf_wage_components)), EPF_WAGE_CEILING)
        edli_wages = epf_wages
        gross_wages = epf_wages

        epf_contribution = int(round(employee_pf))
        expected_contribution = int(round(epf_wages * 0.12))
        if abs(expected_contribution - epf_contribution) > 1:
            warnings.append(
                f"{data['name']}: PF wage basis (Basic+Conveyance+Medical+Special, capped Rs 15,000) "
                f"computes to Rs {epf_wages:,} -> expected Employee PF Rs {expected_contribution:,}, but "
                f"the payslip shows Rs {epf_contribution:,} — check this row in the source sheet before filing."
            )

        if uan in eps_excluded_uans:
            eps_wages = 0
            eps_contribution = 0
        else:
            eps_wages = epf_wages
            eps_contribution = min(int(round(eps_wages * EPS_RATE)), EPS_CONTRIBUTION_CAP)
        epf_eps_diff = epf_contribution - eps_contribution
        ncp_days = int(data.get("lop_days") or 0)

        fields = [
            uan,
            str(data["name"]).strip().upper(),
            str(gross_wages), str(epf_wages), str(eps_wages), str(edli_wages),
            str(epf_contribution), str(eps_contribution), str(epf_eps_diff),
            str(ncp_days), "0",
        ]
        lines.append(DELIMITER.join(fields))

    return lines, skipped, warnings


def validate_cross_checks(lines):
    """Verify EPS + DIFF = EPF for every line, and totals reconcile."""
    total_epf, total_eps_plus_diff = 0, 0
    errors = []
    for i, line in enumerate(lines, start=1):
        parts = line.split(DELIMITER)
        epf_contribution = int(parts[6])
        eps_contribution = int(parts[7])
        epf_eps_diff = int(parts[8])
        if eps_contribution + epf_eps_diff != epf_contribution:
            errors.append(f"Line {i}: EPS ({eps_contribution}) + DIFF ({epf_eps_diff}) != EPF ({epf_contribution})")
        total_epf += epf_contribution
        total_eps_plus_diff += eps_contribution + epf_eps_diff
    return errors, total_epf, total_eps_plus_diff


def generate_ecr(workbook_path, salary_sheet_name, output_path=None, eps_excluded_uans=None):
    company = {"address": "", "cin": "", "tan": "", "pf_code": "", "esi_code": "", "pt_reg": ""}
    month_label, rows, _unmatched, _warnings = generate_payslip.load_payslip_rows(
        workbook_path, salary_sheet_name, "Employee_Details", company)

    print(f"\n{'='*60}")
    print("GENERATING PF ECR FILE")
    print(f"Workbook: {workbook_path}")
    print(f"Sheet: {salary_sheet_name}")
    print(f"{'='*60}\n")

    lines, skipped, ecr_warnings = build_ecr_lines(rows, eps_excluded_uans)
    print(f"Employees included: {len(lines)}")
    print(f"Employees skipped (no valid UAN or no PF deduction): {skipped}")
    if ecr_warnings:
        print("\n⚠ WARNINGS:")
        for w in ecr_warnings:
            print(f"  - {w}")

    if not lines:
        print("\nERROR: No PF-enrolled employees with valid 12-digit UANs found.")
        sys.exit(1)

    errors, total_epf, total_eps_diff = validate_cross_checks(lines)
    if errors:
        print("\n✗ Cross-check FAILED:")
        for e in errors:
            print(f"  - {e}")
        sys.exit(1)
    else:
        print(f"\n✓ Cross-check passed: total EPF contribution (₹{total_epf}) "
              f"= total EPS+DIFF (₹{total_eps_diff})")

    TMP_DIR.mkdir(parents=True, exist_ok=True)
    if output_path is None:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        safe_month = month_label.replace(" ", "_").replace("/", "-")
        output_path = str(TMP_DIR / f"ECR_{safe_month}_{timestamp}.txt")

    text = "\n".join(lines) + "\n"
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(text)

    print(f"\n{'='*60}")
    print(f"✓ ECR FILE SAVED: {output_path}")
    print(f"  Delimiter: '{DELIMITER}'  |  Lines: {len(lines)}")
    print("  Upload this file against your Establishment Code on the EPFO Unified Portal.")
    print("  The establishment header/TRRN is entered separately on the portal — NOT in this file.")
    print(f"{'='*60}\n")

    return output_path


def main():
    parser = argparse.ArgumentParser(description="Generate an EPFO ECR 2.0 file from the real payroll Excel")
    parser.add_argument("--workbook", required=True, help="Path to the payroll .xlsx workbook")
    parser.add_argument("--salary-sheet", required=True, help="e.g. 'Salary Sheet_June 2026'")
    parser.add_argument("--output", default=None, help="Output .txt path (default: .tmp/ECR_<month>_<timestamp>.txt)")
    parser.add_argument("--eps-excluded-uans", default="",
                         help="Comma-separated 12-digit UANs the EPFO portal has recorded as EPS-excluded")
    args = parser.parse_args()

    excluded = {u.strip() for u in args.eps_excluded_uans.split(",") if u.strip()}
    generate_ecr(args.workbook, args.salary_sheet, args.output, excluded)


if __name__ == "__main__":
    main()
