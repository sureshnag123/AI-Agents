"""
payroll_pipeline.py — Main pipeline: reads payroll Excel → generates PDF payslips.

Usage:
    # Generate payslips for a specific month sheet
    python payroll_pipeline.py --excel "C:\\path\\to\\Thota_Payroll 2025.xlsx" --sheet "Salary Stmt-JAN26"

    # Generate payslips for ALL monthly sheets
    python payroll_pipeline.py --excel "C:\\path\\to\\Thota_Payroll 2025.xlsx" --all

    # Generate payslips for the latest month (default)
    python payroll_pipeline.py --excel "C:\\path\\to\\Thota_Payroll 2025.xlsx"

Output:
    PDFs saved to .tmp/payslips/<MONTH>/ folder
"""

import argparse
import os
import sys
import time

# Add parent dir to path so we can import sibling modules
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from payroll_reader import (
    load_workbook, list_salary_sheets, read_payroll_sheet,
    read_pf_esi_data, enrich_employees_with_pf_esi,
)
from payslip_pdf_generator import generate_payslip

# Months to process (July 2025 to January 2026)
JULY_TO_JAN_SHEETS = [
    "Salary Stmt-JULY25",
    "Salary Stmt-AUG25",
    "Salary Stmt-SEPT25",
    "Salary Stmt-OCT25",
    "Salary Stmt-NOV25",
    "Salary Stmt-DEC25",
    "Salary Stmt-JAN26",
]


def get_output_root() -> str:
    """Return the .tmp/payslips directory relative to the project root."""
    project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(project_root, ".tmp", "payslips")


def process_sheet(excel_path: str, sheet_name: str, wb=None, pf_esi_data=None) -> list[str]:
    """Process one sheet and return list of generated PDF paths."""
    print(f"\n{'='*60}")
    print(f"  Processing: {sheet_name}")
    print(f"{'='*60}")

    data = read_payroll_sheet(excel_path, sheet_name, wb=wb)
    company = data["company"]
    month = data["month"]
    employees = data["employees"]

    # Enrich with PF UAN, ESI IP numbers and default DOJ
    if pf_esi_data:
        employees = enrich_employees_with_pf_esi(employees, pf_esi_data)

    if not employees:
        print("  ⚠  No employees found in this sheet.")
        return []

    output_dir = os.path.join(get_output_root(), month)
    os.makedirs(output_dir, exist_ok=True)

    generated = []
    for emp in employees:
        try:
            path = generate_payslip(emp, company, month, output_dir)
            generated.append(path)
            print(f"  ✓ {emp['emp_name']:25s} → {os.path.basename(path)}")
        except Exception as e:
            print(f"  ✗ {emp['emp_name']:25s} → ERROR: {e}")

    print(f"\n  Generated {len(generated)}/{len(employees)} payslips → {output_dir}")
    return generated


def main():
    parser = argparse.ArgumentParser(
        description="Generate PDF payslips from Thota Hospitality payroll Excel."
    )
    parser.add_argument(
        "--excel", required=True,
        help="Path to the payroll Excel file (e.g., Thota_Payroll 2025.xlsx)",
    )
    parser.add_argument(
        "--sheet", default=None,
        help="Specific sheet name to process (e.g., 'Salary Stmt-JAN26')",
    )
    parser.add_argument(
        "--all", action="store_true", dest="all_sheets",
        help="Process ALL salary sheets in the workbook",
    )
    args = parser.parse_args()

    if not os.path.isfile(args.excel):
        print(f"ERROR: File not found: {args.excel}")
        sys.exit(1)

    # Load workbook ONCE (avoids repeated slow loads of the large file)
    print("Loading workbook …")
    wb = load_workbook(args.excel)

    sheets = list_salary_sheets(wb=wb)
    if not sheets:
        print("ERROR: No salary statement sheets found in the workbook.")
        wb.close()
        sys.exit(1)

    # Read PF UAN & ESI IP data once
    print("Reading PF UAN & ESI IP data …")
    pf_esi_data = read_pf_esi_data(wb=wb)
    print(f"  UAN records: {len(pf_esi_data['uan_map'])}, ESI records: {len(pf_esi_data['esi_map'])}")

    print(f"╔{'═'*58}╗")
    print(f"║  THOTA HOSPITALITY LLP — Payslip Generator{' '*13}║")
    print(f"╚{'═'*58}╝")
    print(f"\nExcel  : {args.excel}")
    print(f"Sheets : {', '.join(sheets)}")

    start_time = time.time()

    if args.sheet:
        # Process specific sheet
        all_sheet_names = wb.sheetnames
        if args.sheet in all_sheet_names:
            sheets_to_process = [args.sheet]
        else:
            print(f"ERROR: Sheet '{args.sheet}' not found.")
            print(f"Available sheets: {all_sheet_names}")
            wb.close()
            sys.exit(1)
    elif args.all_sheets:
        # Filter to only July 2025 - January 2026
        sheets_to_process = [s for s in JULY_TO_JAN_SHEETS if s in wb.sheetnames]
        if not sheets_to_process:
            # Fallback: use all salary sheets
            sheets_to_process = sheets
    else:
        # Default: process the first sheet (latest)
        sheets_to_process = [sheets[0]]
        print(f"\nNo --sheet specified. Using first sheet: {sheets[0]}")

    total_generated = []
    for sheet in sheets_to_process:
        generated = process_sheet(args.excel, sheet, wb=wb, pf_esi_data=pf_esi_data)
        total_generated.extend(generated)

    wb.close()

    elapsed = time.time() - start_time

    print(f"\n{'─'*60}")
    print(f"  SUMMARY")
    print(f"{'─'*60}")
    print(f"  Total payslips generated : {len(total_generated)}")
    print(f"  Output directory         : {get_output_root()}")
    print(f"  Time elapsed             : {elapsed:.1f}s")
    print(f"{'─'*60}")

    if total_generated:
        print(f"\n  Payslips are ready to download from:")
        # Show unique output dirs
        dirs = sorted(set(os.path.dirname(p) for p in total_generated))
        for d in dirs:
            count = sum(1 for p in total_generated if os.path.dirname(p) == d)
            print(f"    📁 {d}  ({count} files)")

    print()


if __name__ == "__main__":
    main()
