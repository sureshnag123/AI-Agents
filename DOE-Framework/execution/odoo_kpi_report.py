"""
odoo_kpi_report.py — Pulls live KPI metrics from Odoo for Fracktal Works.

Covers: Procurement, Inventory, Production, QC, and Dispatch.
READ-ONLY: never modifies Odoo data.

Usage:
    python execution/odoo_kpi_report.py --mode full --days 30 --output .tmp/kpi_report.json
    python execution/odoo_kpi_report.py --mode procurement --days 30
    python execution/odoo_kpi_report.py --mode inventory --days 7
    python execution/odoo_kpi_report.py --mode production --days 7
    python execution/odoo_kpi_report.py --mode dispatch --days 7
    python execution/odoo_kpi_report.py --mode qc --days 30
"""

import argparse
import json
import os
import sys
from datetime import datetime, timedelta

sys.path.insert(0, os.path.dirname(__file__))
from odoo_connect import OdooClient


def date_str(days_ago=0):
    return (datetime.now() - timedelta(days=days_ago)).strftime("%Y-%m-%d 00:00:00")


def today_str():
    return datetime.now().strftime("%Y-%m-%d")


# ── PROCUREMENT KPIs ──────────────────────────────────────────────────────────

def get_procurement_kpis(client, days):
    print("  Pulling procurement KPIs...")
    since = date_str(days)
    report = {}

    # Total POs in period
    total_pos = client.count("purchase.order", [
        ("state", "in", ["purchase", "done"]),
        ("date_approve", ">=", since)
    ])
    report["total_confirmed_pos"] = total_pos

    # PO Compliance Rate: POs vs total receipts (estimate)
    total_receipts = client.count("stock.picking", [
        ("picking_type_code", "=", "incoming"),
        ("state", "=", "done"),
        ("date_done", ">=", since)
    ])
    receipts_with_po = client.count("stock.picking", [
        ("picking_type_code", "=", "incoming"),
        ("state", "=", "done"),
        ("date_done", ">=", since),
        ("purchase_id", "!=", False)
    ])
    non_po_receipts = total_receipts - receipts_with_po
    po_compliance = round((receipts_with_po / total_receipts * 100) if total_receipts > 0 else 0, 1)
    report["total_receipts"] = total_receipts
    report["receipts_with_po"] = receipts_with_po
    report["non_po_receipts"] = non_po_receipts
    report["po_compliance_rate_pct"] = po_compliance
    report["po_compliance_status"] = (
        "CRITICAL" if po_compliance < 85 else
        "WARNING" if po_compliance < 95 else "OK"
    )

    # Overdue POs (promised date passed, not received)
    overdue_pos = client.search_read(
        "purchase.order",
        [
            ("state", "=", "purchase"),
            ("date_planned", "<", today_str())
        ],
        fields=["name", "partner_id", "date_planned", "amount_total"],
        limit=50
    )
    report["overdue_pos"] = len(overdue_pos)
    report["overdue_po_list"] = [
        {
            "po": p["name"],
            "vendor": p["partner_id"][1] if p.get("partner_id") else "Unknown",
            "due_date": str(p.get("date_planned", "")),
            "value": p.get("amount_total", 0)
        }
        for p in overdue_pos
    ]

    # POs awaiting approval
    pending_approval = client.count("purchase.order", [("state", "=", "to approve")])
    report["pos_pending_approval"] = pending_approval

    # Vendor OTD (on-time delivery): POs received on or before promised date
    done_pos = client.search_read(
        "purchase.order",
        [("state", "=", "done"), ("date_approve", ">=", since)],
        fields=["name", "date_planned", "effective_date"],
        limit=500
    )
    if done_pos:
        on_time = sum(
            1 for p in done_pos
            if p.get("effective_date") and p.get("date_planned") and
            str(p["effective_date"])[:10] <= str(p["date_planned"])[:10]
        )
        otd = round(on_time / len(done_pos) * 100, 1)
    else:
        otd = None
    report["vendor_otd_pct"] = otd
    report["vendor_otd_status"] = (
        "NO DATA" if otd is None else
        "CRITICAL" if otd < 60 else
        "WARNING" if otd < 80 else "OK"
    )

    return report


# ── INVENTORY KPIs ────────────────────────────────────────────────────────────

def get_inventory_kpis(client, days):
    print("  Pulling inventory KPIs...")
    since = date_str(days)
    report = {}

    # Items under IQC (not yet approved)
    iqc_items = client.search_read(
        "stock.quant",
        [("location_id.name", "ilike", "IQC")],
        fields=["product_id", "quantity", "in_date"],
        limit=200
    )
    report["items_under_iqc"] = len(iqc_items)
    report["iqc_queue"] = [
        {
            "product": q["product_id"][1] if q.get("product_id") else "Unknown",
            "qty": q.get("quantity", 0),
            "since": str(q.get("in_date", ""))[:16]
        }
        for q in iqc_items
    ]

    # Items in Rejection Hold
    rejection_items = client.count(
        "stock.quant",
        [("location_id.name", "ilike", "Rejection")]
    )
    report["items_in_rejection_hold"] = rejection_items

    # Scrap in period
    scrap_moves = client.search_read(
        "stock.scrap",
        [("date_done", ">=", since), ("state", "=", "done")],
        fields=["product_id", "scrap_qty", "date_done", "origin"],
        limit=200
    )
    report["scrap_count"] = len(scrap_moves)
    report["scrap_items"] = [
        {
            "product": s["product_id"][1] if s.get("product_id") else "Unknown",
            "qty": s.get("scrap_qty", 0),
            "date": str(s.get("date_done", ""))[:10],
            "ref": s.get("origin", "")
        }
        for s in scrap_moves
    ]

    # Products below reorder point
    try:
        below_reorder = client.search_read(
            "stock.warehouse.orderpoint",
            [("qty_on_hand", "<", "product_min_qty")],
            fields=["product_id", "qty_on_hand", "product_min_qty", "product_max_qty"],
            limit=100
        )
        report["stockout_risk_items"] = len(below_reorder)
        report["stockout_risk_list"] = [
            {
                "product": r["product_id"][1] if r.get("product_id") else "Unknown",
                "on_hand": r.get("qty_on_hand", 0),
                "min_qty": r.get("product_min_qty", 0)
            }
            for r in below_reorder
        ]
    except Exception:
        # Reorder rules may not be set up yet
        report["stockout_risk_items"] = "N/A - reorder rules not configured"
        report["stockout_risk_list"] = []

    return report


# ── PRODUCTION KPIs ───────────────────────────────────────────────────────────

def get_production_kpis(client, days):
    print("  Pulling production KPIs...")
    since = date_str(days)
    report = {}

    # MOs by status
    for state, label in [
        ("confirmed", "planned"),
        ("progress", "in_progress"),
        ("done", "completed"),
        ("cancel", "cancelled"),
    ]:
        count = client.count("mrp.production", [
            ("state", "=", state),
            ("create_date", ">=", since)
        ])
        report[f"mo_{label}"] = count

    # Total active MOs
    report["mo_total_active"] = client.count(
        "mrp.production",
        [("state", "in", ["confirmed", "progress"])]
    )

    # Overdue MOs
    overdue_mos = client.search_read(
        "mrp.production",
        [
            ("state", "in", ["confirmed", "progress"]),
            ("date_planned_finished", "<", today_str())
        ],
        fields=["name", "product_id", "date_planned_finished", "date_planned_start"],
        limit=50
    )
    report["mo_overdue"] = len(overdue_mos)
    report["mo_overdue_list"] = [
        {
            "mo": m["name"],
            "product": m["product_id"][1] if m.get("product_id") else "Unknown",
            "due": str(m.get("date_planned_finished", ""))[:10]
        }
        for m in overdue_mos
    ]

    # MOs completed on time in period (planned finish >= actual finish)
    completed_mos = client.search_read(
        "mrp.production",
        [("state", "=", "done"), ("date_finished", ">=", since)],
        fields=["name", "date_planned_finished", "date_finished"],
        limit=500
    )
    if completed_mos:
        on_time = sum(
            1 for m in completed_mos
            if m.get("date_finished") and m.get("date_planned_finished") and
            str(m["date_finished"])[:10] <= str(m["date_planned_finished"])[:10]
        )
        otp = round(on_time / len(completed_mos) * 100, 1)
    else:
        otp = None
    report["on_time_production_pct"] = otp
    report["on_time_production_status"] = (
        "NO DATA" if otp is None else
        "CRITICAL" if otp < 70 else
        "WARNING" if otp < 85 else "OK"
    )

    # MOs awaiting material (availability issue)
    mo_blocked = client.count(
        "mrp.production",
        [("state", "=", "confirmed"), ("reservation_state", "=", "waiting")]
    )
    report["mo_waiting_for_material"] = mo_blocked

    return report


# ── QC KPIs ───────────────────────────────────────────────────────────────────

def get_qc_kpis(client, days):
    print("  Pulling QC KPIs...")
    since = date_str(days)
    report = {}

    try:
        # Quality checks performed
        total_checks = client.count(
            "quality.check",
            [("create_date", ">=", since)]
        )
        passed_checks = client.count(
            "quality.check",
            [("create_date", ">=", since), ("quality_state", "=", "pass")]
        )
        failed_checks = client.count(
            "quality.check",
            [("create_date", ">=", since), ("quality_state", "=", "fail")]
        )
        report["qc_total_checks"] = total_checks
        report["qc_passed"] = passed_checks
        report["qc_failed"] = failed_checks
        report["qc_first_pass_rate_pct"] = round(
            (passed_checks / total_checks * 100) if total_checks > 0 else 0, 1
        )

        # Quality alerts (failures requiring action)
        open_alerts = client.count(
            "quality.alert",
            [("stage_id.name", "not in", ["Closed", "Done"])]
        )
        report["open_quality_alerts"] = open_alerts

        # Recent failures
        failures = client.search_read(
            "quality.check",
            [("create_date", ">=", since), ("quality_state", "=", "fail")],
            fields=["name", "product_id", "point_id", "create_date"],
            limit=20
        )
        report["recent_failures"] = [
            {
                "check": f["name"],
                "product": f["product_id"][1] if f.get("product_id") else "Unknown",
                "point": f["point_id"][1] if f.get("point_id") else "Unknown",
                "date": str(f.get("create_date", ""))[:10]
            }
            for f in failures
        ]

    except Exception as e:
        report["error"] = f"Quality module may not be installed: {e}"

    return report


# ── DISPATCH KPIs ─────────────────────────────────────────────────────────────

def get_dispatch_kpis(client, days):
    print("  Pulling dispatch KPIs...")
    since = date_str(days)
    report = {}

    # Deliveries completed in period
    completed_deliveries = client.count("stock.picking", [
        ("picking_type_code", "=", "outgoing"),
        ("state", "=", "done"),
        ("date_done", ">=", since)
    ])
    report["deliveries_completed"] = completed_deliveries

    # Ready to dispatch (validated, not yet done) — FQC passed
    ready_to_dispatch = client.search_read(
        "stock.picking",
        [
            ("picking_type_code", "=", "outgoing"),
            ("state", "=", "assigned")
        ],
        fields=["name", "partner_id", "scheduled_date", "sale_id"],
        limit=100
    )
    report["orders_ready_to_dispatch"] = len(ready_to_dispatch)
    report["ready_to_dispatch_list"] = [
        {
            "delivery": p["name"],
            "customer": p["partner_id"][1] if p.get("partner_id") else "Unknown",
            "scheduled": str(p.get("scheduled_date", ""))[:10],
            "so": p["sale_id"][1] if p.get("sale_id") else "N/A"
        }
        for p in ready_to_dispatch
    ]

    # Overdue deliveries (scheduled date < today, not done)
    overdue_deliveries = client.search_read(
        "stock.picking",
        [
            ("picking_type_code", "=", "outgoing"),
            ("state", "in", ["confirmed", "assigned"]),
            ("scheduled_date", "<", today_str())
        ],
        fields=["name", "partner_id", "scheduled_date"],
        limit=50
    )
    report["overdue_deliveries"] = len(overdue_deliveries)
    report["overdue_delivery_list"] = [
        {
            "delivery": p["name"],
            "customer": p["partner_id"][1] if p.get("partner_id") else "Unknown",
            "due": str(p.get("scheduled_date", ""))[:10]
        }
        for p in overdue_deliveries
    ]

    # On-time delivery rate
    done_deliveries = client.search_read(
        "stock.picking",
        [
            ("picking_type_code", "=", "outgoing"),
            ("state", "=", "done"),
            ("date_done", ">=", since)
        ],
        fields=["scheduled_date", "date_done"],
        limit=500
    )
    if done_deliveries:
        on_time = sum(
            1 for p in done_deliveries
            if p.get("date_done") and p.get("scheduled_date") and
            str(p["date_done"])[:10] <= str(p["scheduled_date"])[:10]
        )
        otd = round(on_time / len(done_deliveries) * 100, 1)
    else:
        otd = None
    report["on_time_dispatch_pct"] = otd
    report["on_time_dispatch_status"] = (
        "NO DATA" if otd is None else
        "CRITICAL" if otd < 75 else
        "WARNING" if otd < 90 else "OK"
    )

    # Pending invoices (delivered but not invoiced)
    try:
        pending_invoices = client.count("sale.order", [
            ("invoice_status", "=", "to invoice"),
            ("state", "=", "sale")
        ])
        report["pending_invoices"] = pending_invoices
    except Exception:
        report["pending_invoices"] = "N/A"

    # Overdue customer payments
    try:
        overdue_payments = client.search_read(
            "account.move",
            [
                ("move_type", "=", "out_invoice"),
                ("state", "=", "posted"),
                ("payment_state", "in", ["not_paid", "partial"]),
                ("invoice_date_due", "<", today_str())
            ],
            fields=["name", "partner_id", "invoice_date_due", "amount_residual"],
            limit=50
        )
        report["overdue_payments"] = len(overdue_payments)
        report["overdue_payment_list"] = [
            {
                "invoice": p["name"],
                "customer": p["partner_id"][1] if p.get("partner_id") else "Unknown",
                "due": str(p.get("invoice_date_due", "")),
                "outstanding": p.get("amount_residual", 0)
            }
            for p in overdue_payments
        ]
    except Exception:
        report["overdue_payments"] = "N/A"

    return report


# ── PRINT REPORT ──────────────────────────────────────────────────────────────

def print_kpi_report(report, days):
    print(f"\n{'='*60}")
    print(f"  FRACKTAL WORKS — KPI REPORT (Last {days} days)")
    print(f"  Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    print(f"{'='*60}")

    section_labels = {
        "procurement": "PROCUREMENT",
        "inventory": "INVENTORY",
        "production": "PRODUCTION",
        "qc": "QUALITY CONTROL",
        "dispatch": "DISPATCH & INVOICING"
    }

    alert_fields = {
        "po_compliance_status", "vendor_otd_status",
        "on_time_production_status", "on_time_dispatch_status"
    }

    for section, label in section_labels.items():
        if section not in report:
            continue
        data = report[section]
        print(f"\n  [{label}]")
        for key, value in data.items():
            if key.endswith("_list") or key == "missing_bom_products":
                continue  # Skip detail lists in summary
            status_indicator = ""
            if key in alert_fields:
                icons = {"CRITICAL": " ⚠ CRITICAL", "WARNING": " ! WARNING", "OK": " ✓", "NO DATA": " –"}
                status_indicator = icons.get(str(value), "")
            print(f"    {key.replace('_', ' ').title():<40} {value}{status_indicator}")

    print(f"\n{'='*60}")

    # Surface critical alerts
    critical = []
    warnings = []
    for section, data in report.items():
        if isinstance(data, dict):
            for key, value in data.items():
                if value == "CRITICAL":
                    metric = key.replace("_status", "").replace("_", " ").title()
                    critical.append(f"  ✗ CRITICAL: {metric} in {section.upper()}")
                elif value == "WARNING":
                    metric = key.replace("_status", "").replace("_", " ").title()
                    warnings.append(f"  ! WARNING:  {metric} in {section.upper()}")

    if critical:
        print("\n  ACTION REQUIRED:")
        for msg in critical:
            print(msg)
    if warnings:
        print("\n  REVIEW:")
        for msg in warnings:
            print(msg)
    if not critical and not warnings:
        print("\n  No critical issues detected.")

    print(f"{'='*60}\n")


# ── MAIN ──────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Odoo KPI Report for Fracktal Works")
    parser.add_argument("--mode", choices=["procurement", "inventory", "production", "qc", "dispatch", "full"],
                        default="full")
    parser.add_argument("--days", type=int, default=30, help="Number of days to look back")
    parser.add_argument("--output", help="Save JSON output to file (e.g. .tmp/kpi_report.json)")
    args = parser.parse_args()

    os.makedirs(".tmp", exist_ok=True)
    print(f"\nOdoo KPI Report — Mode: {args.mode}, Period: last {args.days} days")
    print(f"Connecting to: {os.getenv('ODOO_URL', '(not set)')}\n")

    client = OdooClient()
    report = {}

    mode_fn_map = {
        "procurement": get_procurement_kpis,
        "inventory":   get_inventory_kpis,
        "production":  get_production_kpis,
        "qc":          get_qc_kpis,
        "dispatch":    get_dispatch_kpis,
    }

    targets = list(mode_fn_map.keys()) if args.mode == "full" else [args.mode]

    for key in targets:
        try:
            report[key] = mode_fn_map[key](client, args.days)
        except Exception as e:
            report[key] = {"error": str(e)}
            print(f"  Error in {key}: {e}")

    report["meta"] = {
        "generated_at": datetime.now().isoformat(),
        "period_days": args.days,
        "odoo_url": os.getenv("ODOO_URL", "")
    }

    print_kpi_report(report, args.days)

    if args.output:
        with open(args.output, "w") as f:
            json.dump(report, f, indent=2, default=str)
        print(f"Report saved to: {args.output}")


if __name__ == "__main__":
    main()
