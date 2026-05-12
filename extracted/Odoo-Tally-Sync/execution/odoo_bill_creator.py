#!/usr/bin/env python3
"""
Odoo Bill Creator

Creates a vendor bill (account.move, move_type=in_invoice) in Odoo 19
from invoice data extracted by pdf_ocr_extractor.
Attaches the original PDF to the created bill.

Usage:
    from odoo_bill_creator import create_vendor_bill
    result = create_vendor_bill(extracted_data, expense_ledger="Office Supplies")

    # CLI dry-run test
    python odoo_bill_creator.py --test-ledger "Office Supplies"
"""

import sys
import json
import base64
import argparse
from pathlib import Path

from dotenv import load_dotenv

SCRIPT_DIR = Path(__file__).parent.resolve()
PROJECT_ROOT = SCRIPT_DIR.parent
load_dotenv(PROJECT_ROOT / ".env")

sys.path.insert(0, str(SCRIPT_DIR))
from odoo_connector import OdooConnector


# ── Vendor helpers ────────────────────────────────────────────────────────────

def find_or_create_vendor(odoo: OdooConnector, vendor_name: str, gstin: str = None) -> int:
    """Find existing vendor by name (case-insensitive), or create a new one."""
    # Try exact ilike match first (most reliable)
    ids = odoo.search(
        "res.partner",
        [("name", "ilike", vendor_name), ("is_company", "=", True)],
        limit=1,
    )
    if ids:
        rec = odoo.read("res.partner", ids, ["name"])[0]
        print(f"  Vendor matched: '{rec['name']}' (ID {ids[0]})")
        return ids[0]

    # Partial match on first 20 characters
    short = vendor_name[:20].strip()
    ids = odoo.search("res.partner", [("name", "ilike", short)], limit=1)
    if ids:
        rec = odoo.read("res.partner", ids, ["name"])[0]
        print(f"  Vendor partial match: '{vendor_name}' → '{rec['name']}' (ID {ids[0]})")
        return ids[0]

    # Create new vendor
    vals = {"name": vendor_name, "is_company": True, "supplier_rank": 1}
    if gstin:
        vals["vat"] = gstin
    new_id = odoo._execute("res.partner", "create", vals)
    print(f"  Vendor created: '{vendor_name}' (ID {new_id})")
    return new_id


# ── Account helpers ───────────────────────────────────────────────────────────

def find_expense_account(odoo: OdooConnector, ledger_name: str) -> int:
    """Find Odoo account by name. Raises if not found."""
    ids = odoo.search("account.account", [("name", "ilike", ledger_name)], limit=5)
    if not ids:
        raise ValueError(
            f"Expense ledger not found: '{ledger_name}'\n"
            f"Run: python odoo_connector.py --list-accounts  (check .tmp/odoo_accounts.json)"
        )
    rec = odoo.read("account.account", ids, ["id", "name", "code"])[0]
    print(f"  Ledger matched: '{rec['name']}' ({rec['code']}) (ID {rec['id']})")
    return rec["id"]


# ── Tax helpers ───────────────────────────────────────────────────────────────

def find_tax(odoo: OdooConnector, tax_rate: float) -> int | None:
    """Find an active purchase tax by percentage rate. Returns None if not found."""
    if not tax_rate:
        return None
    ids = odoo.search(
        "account.tax",
        [("amount", "=", tax_rate), ("type_tax_use", "=", "purchase"), ("active", "=", True)],
        limit=1,
    )
    if ids:
        rec = odoo.read("account.tax", ids, ["name", "amount"])[0]
        print(f"  Tax matched: {rec['name']} ({tax_rate}%) (ID {ids[0]})")
        return ids[0]
    print(f"  Warning: No active purchase tax found for {tax_rate}% — line will have no tax")
    return None


# ── Journal helpers ───────────────────────────────────────────────────────────

def find_purchase_journal(odoo: OdooConnector) -> int:
    """Return the ID of the first purchase journal."""
    ids = odoo.search("account.journal", [("type", "=", "purchase")], limit=1)
    if not ids:
        raise ValueError("No purchase journal found in Odoo.")
    rec = odoo.read("account.journal", ids, ["name"])[0]
    print(f"  Journal: '{rec['name']}' (ID {ids[0]})")
    return ids[0]


# ── Duplicate check ───────────────────────────────────────────────────────────

def find_duplicate(odoo: OdooConnector, vendor_id: int, invoice_number: str) -> int | None:
    """Return existing bill ID if vendor + invoice_number already exists, else None."""
    if not invoice_number:
        return None
    ids = odoo.search(
        "account.move",
        [
            ("move_type", "=", "in_invoice"),
            ("partner_id", "=", vendor_id),
            ("ref", "=", invoice_number),
        ],
        limit=1,
    )
    return ids[0] if ids else None


# ── Main creator ──────────────────────────────────────────────────────────────

def create_vendor_bill(extracted_data: dict, expense_ledger: str) -> dict:
    """
    Create a vendor bill in Odoo from extracted invoice data.

    Args:
        extracted_data : dict returned by pdf_ocr_extractor.extract_invoice_data()
        expense_ledger : Odoo account name from Excel column A

    Returns:
        {
            "bill_id":   int,
            "bill_name": str,   e.g. "BILL/2026/0042"
            "status":    "created" | "duplicate"
        }
    """
    odoo = OdooConnector()

    vendor_name    = (extracted_data.get("vendor_name") or "Unknown Vendor").strip()
    invoice_number = (extracted_data.get("invoice_number") or "").strip()
    invoice_date   = extracted_data.get("invoice_date")   or None
    due_date       = extracted_data.get("due_date")       or None
    gstin          = extracted_data.get("gstin")          or None
    line_items     = extracted_data.get("line_items")     or []
    tax_rate       = float(extracted_data.get("tax_rate") or 0)
    subtotal       = float(extracted_data.get("subtotal") or extracted_data.get("total") or 0)
    pdf_bytes      = extracted_data.get("_pdf_bytes")
    drive_url      = extracted_data.get("_drive_url", "")
    notes          = extracted_data.get("_notes", "")

    # ── Lookup Odoo records ───────────────────────────────────────────────────
    vendor_id  = find_or_create_vendor(odoo, vendor_name, gstin)
    account_id = find_expense_account(odoo, expense_ledger)
    journal_id = find_purchase_journal(odoo)

    # ── Duplicate guard ───────────────────────────────────────────────────────
    dup_id = find_duplicate(odoo, vendor_id, invoice_number)
    if dup_id:
        dup = odoo.read("account.move", [dup_id], ["name"])[0]
        print(f"  Duplicate: bill '{dup['name']}' already exists — skipping")
        return {"bill_id": dup_id, "bill_name": dup["name"], "status": "duplicate"}

    # ── Build invoice lines ───────────────────────────────────────────────────
    invoice_lines = []

    if line_items:
        for item in line_items:
            item_tax_rate = float(item.get("tax_rate") or tax_rate or 0)
            item_tax_id   = find_tax(odoo, item_tax_rate) if item_tax_rate else None
            line = {
                "name":       (item.get("description") or "Purchase").strip(),
                "quantity":   float(item.get("quantity")   or 1),
                "price_unit": float(item.get("unit_price") or item.get("amount") or 0),
                "account_id": account_id,
            }
            if item_tax_id:
                line["tax_ids"] = [(6, 0, [item_tax_id])]
            invoice_lines.append((0, 0, line))
    else:
        # No line items — create a single summary line
        default_tax_id = find_tax(odoo, tax_rate) if tax_rate else None
        line = {
            "name":       "Purchase (PDF import)",
            "quantity":   1.0,
            "price_unit": subtotal,
            "account_id": account_id,
        }
        if default_tax_id:
            line["tax_ids"] = [(6, 0, [default_tax_id])]
        invoice_lines.append((0, 0, line))

    # ── Build narration ───────────────────────────────────────────────────────
    narration_parts = ["Uploaded via PDF Purchase Agent."]
    if drive_url:
        narration_parts.append(f"Source PDF: {drive_url}")
    if notes:
        narration_parts.append(f"Notes: {notes}")
    narration = " | ".join(narration_parts)

    # ── Create the bill (draft) ───────────────────────────────────────────────
    bill_vals = {
        "move_type":        "in_invoice",
        "partner_id":       vendor_id,
        "journal_id":       journal_id,
        "ref":              invoice_number,
        "narration":        narration,
        "invoice_line_ids": invoice_lines,
    }
    if invoice_date:
        bill_vals["invoice_date"] = invoice_date
    if due_date:
        bill_vals["invoice_date_due"] = due_date

    bill_id  = odoo._execute("account.move", "create", bill_vals)
    bill_rec = odoo.read("account.move", [bill_id], ["name"])[0]
    bill_name = bill_rec["name"]
    print(f"  Bill created: {bill_name} (ID {bill_id}) — status: Draft")

    # ── Attach original PDF ───────────────────────────────────────────────────
    if pdf_bytes:
        safe_inv = (invoice_number or bill_name).replace("/", "-").replace("\\", "-")
        filename = f"Invoice_{safe_inv}.pdf"
        att_vals = {
            "name":      filename,
            "type":      "binary",
            "datas":     base64.b64encode(pdf_bytes).decode("utf-8"),
            "res_model": "account.move",
            "res_id":    bill_id,
            "mimetype":  "application/pdf",
        }
        att_id = odoo._execute("ir.attachment", "create", att_vals)
        print(f"  PDF attached: {filename} (attachment ID {att_id})")

    return {"bill_id": bill_id, "bill_name": bill_name, "status": "created"}


# ── CLI ───────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Odoo Bill Creator — test helpers")
    parser.add_argument("--test-ledger", metavar="NAME",
                        help="Check if a ledger name exists in Odoo")
    parser.add_argument("--list-taxes", action="store_true",
                        help="List all purchase taxes in Odoo")
    args = parser.parse_args()

    odoo = OdooConnector()

    if args.test_ledger:
        try:
            acc_id = find_expense_account(odoo, args.test_ledger)
            print(f"OK — account ID: {acc_id}")
        except ValueError as e:
            print(f"NOT FOUND: {e}")

    elif args.list_taxes:
        taxes = odoo.search_read(
            "account.tax",
            [("type_tax_use", "=", "purchase"), ("active", "=", True)],
            ["id", "name", "amount", "amount_type"],
            order="amount asc",
        )
        print(f"\n{'ID':<8} {'Rate':>8}  {'Name':<40} {'Type':<15}")
        print("─" * 75)
        for t in taxes:
            print(f"{t['id']:<8} {t['amount']:>7.1f}%  {t['name']:<40} {t['amount_type']:<15}")

    else:
        parser.print_help()


if __name__ == "__main__":
    main()
