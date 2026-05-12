#!/usr/bin/env python3
"""
Payment Reminder Report Generator

Creates a summary report of all outstanding invoices, grouped by aging
buckets (Current, 1-30, 31-60, 61-90, 90+ days).

Usage:
    # Print aging report to console
    python reminder_report.py

    # Export to CSV
    python reminder_report.py --csv

    # Export to JSON
    python reminder_report.py --json
"""

import os
import sys
import csv
import json
import argparse
from datetime import date
from collections import defaultdict

from dotenv import load_dotenv

load_dotenv()

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from odoo_connector import OdooConnector


AGING_BUCKETS = [
    ("Current (not yet due)", -999, 0),
    ("1-30 days overdue", 1, 30),
    ("31-60 days overdue", 31, 60),
    ("61-90 days overdue", 61, 90),
    ("90+ days overdue", 91, 9999),
]


def get_all_unpaid_invoices(odoo: OdooConnector) -> list[dict]:
    """Fetch all unpaid/partially paid customer invoices."""
    domain = [
        ("move_type", "=", "out_invoice"),
        ("state", "=", "posted"),
        ("payment_state", "in", ["not_paid", "partial"]),
    ]
    fields = [
        "id", "name", "partner_id", "invoice_date", "invoice_date_due",
        "amount_total", "amount_residual", "currency_id", "payment_state",
        "invoice_payment_term_id", "user_id",
    ]
    return odoo.search_read("account.move", domain, fields, order="invoice_date_due asc")


def classify_aging(due_date_str: str) -> tuple[str, int]:
    """Return (bucket_name, days_overdue) for a given due date."""
    due = date.fromisoformat(str(due_date_str))
    days = (date.today() - due).days  # positive = overdue

    for label, low, high in AGING_BUCKETS:
        if low <= days <= high:
            return label, days

    return "Unknown", days


def generate_report(odoo: OdooConnector) -> dict:
    """Generate the full aging report."""
    invoices = get_all_unpaid_invoices(odoo)

    # Classify each invoice
    report = {
        "date": str(date.today()),
        "total_invoices": len(invoices),
        "total_outstanding": 0.0,
        "currency": "",
        "buckets": defaultdict(lambda: {"invoices": [], "total": 0.0, "count": 0}),
        "by_customer": defaultdict(lambda: {"invoices": [], "total": 0.0}),
    }

    for inv in invoices:
        bucket_name, days = classify_aging(inv["invoice_date_due"])
        currency = inv["currency_id"][1] if inv.get("currency_id") else ""
        report["currency"] = currency
        report["total_outstanding"] += inv["amount_residual"]

        inv_data = {
            "name": inv["name"],
            "customer": inv["partner_id"][1] if inv.get("partner_id") else "Unknown",
            "customer_id": inv["partner_id"][0] if inv.get("partner_id") else 0,
            "invoice_date": str(inv["invoice_date"]),
            "due_date": str(inv["invoice_date_due"]),
            "amount_total": inv["amount_total"],
            "amount_due": inv["amount_residual"],
            "days_overdue": days,
            "bucket": bucket_name,
            "payment_state": inv["payment_state"],
        }

        report["buckets"][bucket_name]["invoices"].append(inv_data)
        report["buckets"][bucket_name]["total"] += inv["amount_residual"]
        report["buckets"][bucket_name]["count"] += 1

        cust_name = inv_data["customer"]
        report["by_customer"][cust_name]["invoices"].append(inv_data)
        report["by_customer"][cust_name]["total"] += inv["amount_residual"]

    return report


def print_report(report: dict):
    """Print formatted aging report to console."""
    currency = report["currency"]

    print("\n" + "=" * 70)
    print(f"  ACCOUNTS RECEIVABLE AGING REPORT — {report['date']}")
    print("=" * 70)
    print(f"  Total Invoices: {report['total_invoices']}")
    print(f"  Total Outstanding: {currency} {report['total_outstanding']:,.2f}")
    print("=" * 70)

    for label, _, _ in AGING_BUCKETS:
        bucket = report["buckets"].get(label, {"count": 0, "total": 0, "invoices": []})
        print(f"\n  📋 {label}  ({bucket['count']} invoices — {currency} {bucket['total']:,.2f})")
        print(f"  {'Invoice':<20} {'Customer':<25} {'Due Date':<12} {'Amount Due':>12}")
        print(f"  {'-'*20} {'-'*25} {'-'*12} {'-'*12}")

        for inv in bucket.get("invoices", []):
            print(
                f"  {inv['name']:<20} {inv['customer'][:25]:<25} "
                f"{inv['due_date']:<12} {currency} {inv['amount_due']:>9,.2f}"
            )

    print("\n" + "=" * 70)
    print("  TOP DEBTORS")
    print("=" * 70)
    sorted_customers = sorted(
        report["by_customer"].items(),
        key=lambda x: x[1]["total"],
        reverse=True,
    )
    for cust, data in sorted_customers[:10]:
        print(f"  {cust:<35} {currency} {data['total']:>12,.2f}  ({len(data['invoices'])} invoices)")

    print("=" * 70)


def export_csv(report: dict, filepath: str):
    """Export aging report to CSV."""
    all_invoices = []
    for bucket in report["buckets"].values():
        all_invoices.extend(bucket.get("invoices", []))

    all_invoices.sort(key=lambda x: x["days_overdue"], reverse=True)

    with open(filepath, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=[
            "name", "customer", "invoice_date", "due_date",
            "amount_total", "amount_due", "days_overdue", "bucket", "payment_state"
        ])
        writer.writeheader()
        writer.writerows(all_invoices)

    print(f"Exported {len(all_invoices)} invoices to {filepath}")


def export_json(report: dict, filepath: str):
    """Export aging report to JSON."""
    # Convert defaultdicts to regular dicts
    output = {
        "date": report["date"],
        "total_invoices": report["total_invoices"],
        "total_outstanding": report["total_outstanding"],
        "currency": report["currency"],
        "buckets": {k: dict(v) for k, v in report["buckets"].items()},
        "by_customer": {k: dict(v) for k, v in report["by_customer"].items()},
    }

    with open(filepath, "w", encoding="utf-8") as f:
        json.dump(output, f, indent=2, default=str)

    print(f"Exported report to {filepath}")


def main():
    parser = argparse.ArgumentParser(description="Generate AR aging report from Odoo")
    parser.add_argument("--csv", action="store_true", help="Export to CSV")
    parser.add_argument("--json", action="store_true", help="Export to JSON")
    args = parser.parse_args()

    odoo = OdooConnector()
    print(f"Connected to Odoo {odoo.version().get('server_version', '?')}")

    report = generate_report(odoo)
    print_report(report)

    tmp_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), ".tmp")
    os.makedirs(tmp_dir, exist_ok=True)

    if args.csv:
        export_csv(report, os.path.join(tmp_dir, f"aging_report_{date.today()}.csv"))

    if args.json:
        export_json(report, os.path.join(tmp_dir, f"aging_report_{date.today()}.json"))


if __name__ == "__main__":
    main()
