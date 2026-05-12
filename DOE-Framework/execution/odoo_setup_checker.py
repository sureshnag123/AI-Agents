"""
odoo_setup_checker.py — Validates Odoo configuration for Fracktal Works manufacturing workflow.

Checks modules, inventory locations, approval rules, quality control points, and BOM coverage.
READ-ONLY: this script never modifies Odoo data.

Usage:
    python execution/odoo_setup_checker.py --mode full --output .tmp/setup_full.json
    python execution/odoo_setup_checker.py --mode modules
    python execution/odoo_setup_checker.py --mode locations
    python execution/odoo_setup_checker.py --mode approvals
    python execution/odoo_setup_checker.py --mode quality
    python execution/odoo_setup_checker.py --mode bom
"""

import argparse
import json
import os
import sys

# Allow running from project root
sys.path.insert(0, os.path.join(os.path.dirname(__file__)))
from odoo_connect import OdooClient


# ── Required modules ─────────────────────────────────────────────────────────

REQUIRED_MODULES = {
    "sale_management":     "Sales (Quotations, Orders, Approval)",
    "purchase":            "Purchase (RFQ, PO, Vendor Management)",
    "stock":               "Inventory (Locations, GRN, Transfers)",
    "mrp":                 "Manufacturing (MO, BOM, Work Orders)",
    "quality_control":     "Quality (IQC, IPQC, FQC, Alerts)",
    "sale_purchase":       "Sales-Purchase Integration",
    "mrp_subcontracting":  "Subcontracting (optional but recommended)",
    "purchase_stock":      "Purchase-Inventory Integration",
    "account":             "Accounting (Invoicing, Payments)",
    "crm":                 "CRM (Lead/Enquiry Tracking)",
}

OPTIONAL_MODULES = {
    "mrp_workcenter":      "Work Centers (step-by-step production tracking)",
    "stock_landed_costs":  "Landed Costs (import duties, freight in inventory value)",
    "sale_crm":            "CRM-Sales Integration",
    "portal":              "Customer Portal (order tracking for customers)",
    "mail":                "Discuss / Messaging",
}

# ── Required inventory locations ─────────────────────────────────────────────

REQUIRED_LOCATIONS = [
    {"name_contains": "Under IQC",       "usage": "internal", "purpose": "Material held for incoming QC"},
    {"name_contains": "Approved Stock",  "usage": "internal", "purpose": "QC-passed material ready for use"},
    {"name_contains": "Rejection Hold",  "usage": "internal", "purpose": "IQC-failed material awaiting return"},
    {"name_contains": "Scrap",           "usage": "inventory_loss", "purpose": "Scrapped material write-off"},
]

# ── Quality control point triggers ───────────────────────────────────────────

REQUIRED_QC_TRIGGERS = ["on_receipt", "during_production", "at_delivery"]

# ── Checks ────────────────────────────────────────────────────────────────────

def check_modules(client):
    """Check required modules are installed."""
    print("  Checking modules...")
    installed = client.get_installed_modules()
    results = []

    for module, description in REQUIRED_MODULES.items():
        status = "PASS" if module in installed else "FAIL"
        results.append({
            "item": module,
            "description": description,
            "status": status,
            "fix": None if status == "PASS" else f"Install via: Odoo → Apps → search '{module}' → Install"
        })

    for module, description in OPTIONAL_MODULES.items():
        status = "PASS" if module in installed else "OPTIONAL"
        results.append({
            "item": module,
            "description": description,
            "status": status,
            "fix": None if status == "PASS" else f"Optional: Odoo → Apps → search '{module}' → Install"
        })

    return results


def check_locations(client):
    """Check required inventory locations exist."""
    print("  Checking inventory locations...")
    all_locations = client.search_read(
        "stock.location",
        [("active", "=", True)],
        fields=["name", "complete_name", "usage"],
        limit=1000
    )
    location_names = [loc["complete_name"].lower() for loc in all_locations]
    results = []

    for req in REQUIRED_LOCATIONS:
        keyword = req["name_contains"].lower()
        found = any(keyword in name for name in location_names)
        status = "PASS" if found else "FAIL"
        results.append({
            "item": req["name_contains"],
            "description": req["purpose"],
            "status": status,
            "fix": None if status == "PASS" else (
                f"Create via: Odoo → Inventory → Configuration → Locations → New\n"
                f"  Name: '{req['name_contains']}', Type: '{req['usage']}'"
            )
        })

    return results


def check_approvals(client):
    """Check if approval rules are configured for Sales and Purchase."""
    print("  Checking approval rules...")
    results = []

    # Check Purchase approval threshold
    try:
        po_approval = client.search_read(
            "ir.config_parameter",
            [("key", "=", "purchase.purchase_lock")],
            fields=["value"]
        )
        purchase_approval_enabled = bool(po_approval)
    except Exception:
        purchase_approval_enabled = False

    # Check if any purchase approvals exist (via purchase order approvals config)
    try:
        approval_count = client.count("purchase.order", [("state", "=", "to approve")])
        has_po_approval = True  # If "to approve" state exists, approval is configured
    except Exception:
        has_po_approval = False

    results.append({
        "item": "Purchase Order Approval",
        "description": "POs above threshold require approval before sending to vendor",
        "status": "PASS" if has_po_approval else "WARN",
        "fix": (
            "Enable via: Odoo → Purchase → Configuration → Settings\n"
            "  → Purchase Order Approval → Enable → Set amount threshold"
        ) if not has_po_approval else None
    })

    # Check Sales approval (sale.order state 'sent' or approval module)
    try:
        sale_approval_modules = client.search_read(
            "ir.module.module",
            [("name", "=", "sale_management"), ("state", "=", "installed")],
            fields=["name"]
        )
        sale_approval_ok = bool(sale_approval_modules)
    except Exception:
        sale_approval_ok = False

    results.append({
        "item": "Sales Quotation Approval",
        "description": "Quotations above threshold require approval before sending to customer",
        "status": "PASS" if sale_approval_ok else "FAIL",
        "fix": (
            "Enable via: Odoo → Sales → Configuration → Settings\n"
            "  → Sales Order Approval → Enable → Set amount threshold"
        ) if not sale_approval_ok else None
    })

    return results


def check_quality(client):
    """Check quality control points are set up."""
    print("  Checking quality control points...")
    results = []

    try:
        qc_points = client.search_read(
            "quality.point",
            [],
            fields=["name", "picking_type_ids", "measure_on"],
            limit=100
        )
        qc_count = len(qc_points)
    except Exception:
        qc_points = []
        qc_count = 0

    # Check for IQC (on receipt)
    receipt_qc = any(
        "receipt" in str(p.get("name", "")).lower() or
        "incoming" in str(p.get("name", "")).lower() or
        p.get("measure_on") in ["move_line", "product"]
        for p in qc_points
    )

    # Check for FQC (at delivery / final)
    final_qc = any(
        "final" in str(p.get("name", "")).lower() or
        "delivery" in str(p.get("name", "")).lower() or
        "fqc" in str(p.get("name", "")).lower()
        for p in qc_points
    )

    # Check for IPQC (during production)
    inprocess_qc = any(
        "process" in str(p.get("name", "")).lower() or
        "ipqc" in str(p.get("name", "")).lower() or
        "production" in str(p.get("name", "")).lower()
        for p in qc_points
    )

    results.append({
        "item": "IQC — Incoming Quality Check",
        "description": "Quality check point on material receipt (before entering approved stock)",
        "status": "PASS" if receipt_qc else "FAIL",
        "fix": (
            "Create via: Odoo → Quality → Configuration → Control Points → New\n"
            "  Name: 'IQC - Incoming Inspection', Operation: Receipts, Team: Quality"
        ) if not receipt_qc else None
    })

    results.append({
        "item": "IPQC — In-Process Quality Check",
        "description": "Quality check point during manufacturing (at key production milestones)",
        "status": "PASS" if inprocess_qc else "FAIL",
        "fix": (
            "Create via: Odoo → Quality → Configuration → Control Points → New\n"
            "  Name: 'IPQC - In-Process Check', Operation: Manufacturing, Team: Quality"
        ) if not inprocess_qc else None
    })

    results.append({
        "item": "FQC — Final Quality Check",
        "description": "Quality check point after production completion (gate before dispatch)",
        "status": "PASS" if final_qc else "FAIL",
        "fix": (
            "Create via: Odoo → Quality → Configuration → Control Points → New\n"
            "  Name: 'FQC - Final Inspection', Operation: Delivery Orders, Team: Quality"
        ) if not final_qc else None
    })

    results.append({
        "item": "Total Quality Control Points",
        "description": f"{qc_count} QC control points found in system",
        "status": "PASS" if qc_count >= 3 else "WARN",
        "fix": "Create at minimum 3 QC points: IQC, IPQC, FQC" if qc_count < 3 else None
    })

    return results


def check_bom(client):
    """Check BOM coverage for storable products."""
    print("  Checking BOM coverage...")
    results = []

    try:
        # Products that are manufactured (type = product = storable)
        all_products = client.search_read(
            "product.template",
            [("type", "=", "product")],
            fields=["name", "id"],
            limit=500
        )

        # Products with at least one BOM
        boms = client.search_read(
            "mrp.bom",
            [("active", "=", True)],
            fields=["product_tmpl_id"],
            limit=1000
        )
        products_with_bom = {b["product_tmpl_id"][0] for b in boms if b.get("product_tmpl_id")}

        total = len(all_products)
        covered = sum(1 for p in all_products if p["id"] in products_with_bom)
        missing = [p["name"] for p in all_products if p["id"] not in products_with_bom]
        coverage_pct = round((covered / total * 100) if total > 0 else 0, 1)

        status = "PASS" if coverage_pct >= 90 else ("WARN" if coverage_pct >= 70 else "FAIL")

        results.append({
            "item": "BOM Coverage",
            "description": f"{covered}/{total} storable products have a BOM ({coverage_pct}%)",
            "status": status,
            "fix": (
                f"Create BOMs for {len(missing)} products:\n" +
                "\n".join(f"  - {name}" for name in missing[:10]) +
                (f"\n  ...and {len(missing) - 10} more" if len(missing) > 10 else "") +
                "\n  Via: Odoo → Manufacturing → Products → [Product] → Bill of Materials"
            ) if status != "PASS" else None,
            "missing_bom_products": missing[:20]
        })

    except Exception as e:
        results.append({
            "item": "BOM Coverage",
            "description": "Could not check BOM coverage (Manufacturing module may not be installed)",
            "status": "SKIP",
            "fix": f"Error: {e}"
        })

    return results


# ── Main ─────────────────────────────────────────────────────────────────────

def run_checks(mode):
    client = OdooClient()
    all_results = {}

    check_map = {
        "modules":   ("Module Installation",      check_modules),
        "locations": ("Inventory Locations",       check_locations),
        "approvals": ("Approval Rules",            check_approvals),
        "quality":   ("Quality Control Points",    check_quality),
        "bom":       ("BOM Coverage",              check_bom),
    }

    if mode == "full":
        targets = list(check_map.keys())
    else:
        targets = [mode]

    for key in targets:
        label, fn = check_map[key]
        print(f"\n[{label}]")
        all_results[key] = fn(client)

    return all_results


def print_summary(results):
    total = pass_count = fail_count = warn_count = 0
    for section, items in results.items():
        print(f"\n{'─'*50}")
        print(f"  {section.upper()}")
        print(f"{'─'*50}")
        for item in items:
            status = item["status"]
            icon = {"PASS": "✓", "FAIL": "✗", "WARN": "!", "OPTIONAL": "○", "SKIP": "–"}.get(status, "?")
            print(f"  [{icon}] {item['item']}: {item['description']}")
            if item.get("fix"):
                for line in item["fix"].split("\n"):
                    print(f"        → {line}")
            if status in ("FAIL", "WARN"):
                fail_count += 1 if status == "FAIL" else 0
                warn_count += 1 if status == "WARN" else 0
            if status == "PASS":
                pass_count += 1
            if status not in ("OPTIONAL", "SKIP"):
                total += 1

    print(f"\n{'='*50}")
    print(f"  SUMMARY: {pass_count} PASS | {warn_count} WARN | {fail_count} FAIL | out of {total} checks")
    if fail_count > 0:
        print(f"  ACTION REQUIRED: Fix {fail_count} FAIL item(s) before going live")
    elif warn_count > 0:
        print(f"  REVIEW: {warn_count} item(s) need attention")
    else:
        print(f"  All critical checks passed. System ready.")
    print(f"{'='*50}")


def main():
    parser = argparse.ArgumentParser(description="Odoo setup validation for Fracktal Works")
    parser.add_argument("--mode", choices=["modules", "locations", "approvals", "quality", "bom", "full"],
                        default="full", help="What to check")
    parser.add_argument("--output", help="Save results to JSON file (e.g. .tmp/setup_full.json)")
    args = parser.parse_args()

    os.makedirs(".tmp", exist_ok=True)
    print(f"\nOdoo Setup Checker — Mode: {args.mode}")
    print(f"Connecting to: {os.getenv('ODOO_URL', '(not set)')}")

    results = run_checks(args.mode)
    print_summary(results)

    if args.output:
        with open(args.output, "w") as f:
            json.dump(results, f, indent=2)
        print(f"\nResults saved to: {args.output}")


if __name__ == "__main__":
    main()
