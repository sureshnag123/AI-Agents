"""
odoo_barcode.py — Barcode-based inventory operations for Fracktal Works.

Works with any USB or Bluetooth barcode scanner (they act as keyboard input).
Also works via manual entry or piped input for testing.

Scan input accepts THREE formats (auto-detected):
  1. Barcode number  — scanned from product barcode label (e.g. 8901234567890)
  2. SKU             — typed or scanned from SKU label    (e.g. FW-RM-0042)
  3. Internal Ref    — any other internal reference in Odoo default_code field

Modes:
  lookup   — Look up a product by barcode or SKU (info only)
  grn      — Goods Receipt: scan items arriving against a PO
  issue    — Material Issue: scan items being issued against a Manufacturing Order
  count    — Cycle Count: scan items and enter qty to update stock
  dispatch — Dispatch Check: scan packed items to validate a Delivery Order

Usage:
  python execution/odoo_barcode.py --mode lookup --barcode 8901234567890
  python execution/odoo_barcode.py --mode lookup --barcode FW-RM-0042
  python execution/odoo_barcode.py --mode grn --ref PO/2024/0042
  python execution/odoo_barcode.py --mode issue --ref MO/2024/0017
  python execution/odoo_barcode.py --mode count --location "Approved Stock"
  python execution/odoo_barcode.py --mode dispatch --ref WH/OUT/00123
  python execution/odoo_barcode.py --mode grn --interactive   (scanner session)

Notes:
  - All write operations require explicit confirmation before committing to Odoo.
  - GRN and Issue create stock moves via Odoo API (uses immediate_transfer).
  - Dispatch validates the delivery order picking.
  - Cycle count uses Odoo inventory adjustment (stock.inventory or quant direct write).
  - SKUs are in format FW-[CATEGORY]-[0001]. Manage with odoo_sku_manager.py.
"""

import argparse
import json
import os
import sys
from datetime import datetime

sys.path.insert(0, os.path.dirname(__file__))
from odoo_connect import OdooClient


# ── SKU Helpers ───────────────────────────────────────────────────────────────

SKU_PREFIX = "FW"

def is_sku(scan_input):
    """Return True if the scanned string looks like a Fracktal SKU (FW-XX-NNNN)."""
    parts = scan_input.strip().split("-")
    return len(parts) == 3 and parts[0].upper() == SKU_PREFIX and parts[2].isdigit()


def _build_product_dict(p, source="product.product"):
    """Normalize a product record into a consistent dict."""
    if source == "product.product":
        return {
            "id": p["id"],
            "tmpl_id": p["product_tmpl_id"][0] if p.get("product_tmpl_id") else None,
            "name": p["name"],
            "sku": p.get("default_code") or "— NOT ASSIGNED —",
            "barcode": p.get("barcode") or "—",
            "internal_ref": p.get("default_code") or "",
            "uom": p["uom_id"][1] if p.get("uom_id") else "Units",
            "on_hand": p.get("qty_available", 0),
            "forecasted": p.get("virtual_available", 0),
            "type": p.get("type", "")
        }
    else:  # product.template
        return {
            "id": None,
            "tmpl_id": p["id"],
            "name": p["name"],
            "sku": p.get("default_code") or "— NOT ASSIGNED —",
            "barcode": p.get("barcode") or "—",
            "internal_ref": p.get("default_code") or "",
            "uom": p["uom_id"][1] if p.get("uom_id") else "Units",
            "on_hand": p.get("qty_available", 0),
            "forecasted": p.get("virtual_available", 0),
            "type": p.get("type", "")
        }


# ── Product Lookup ────────────────────────────────────────────────────────────

PRODUCT_FIELDS_VARIANT = ["name", "barcode", "default_code", "uom_id",
                           "product_tmpl_id", "virtual_available", "qty_available", "type"]
PRODUCT_FIELDS_TMPL = ["name", "barcode", "default_code", "uom_id",
                        "virtual_available", "qty_available", "type"]


def lookup_product(client, scan_input):
    """
    Find a product by barcode number OR SKU (FW-XX-NNNN) OR internal reference.
    Search order:
      1. Barcode field on product.product (variant)
      2. SKU / Internal Reference (default_code) on product.product
      3. Barcode field on product.template
      4. SKU / Internal Reference on product.template
    Returns normalized product dict or None.
    """
    scan_input = scan_input.strip()

    # 1. Barcode match on product.product
    records = client.search_read(
        "product.product",
        [("barcode", "=", scan_input)],
        fields=PRODUCT_FIELDS_VARIANT, limit=1
    )
    if records:
        return _build_product_dict(records[0], "product.product")

    # 2. SKU / Internal Reference match on product.product
    records = client.search_read(
        "product.product",
        [("default_code", "=", scan_input)],
        fields=PRODUCT_FIELDS_VARIANT, limit=1
    )
    if records:
        return _build_product_dict(records[0], "product.product")

    # 3. Barcode match on product.template
    records = client.search_read(
        "product.template",
        [("barcode", "=", scan_input)],
        fields=PRODUCT_FIELDS_TMPL, limit=1
    )
    if records:
        return _build_product_dict(records[0], "product.template")

    # 4. SKU / Internal Reference match on product.template
    records = client.search_read(
        "product.template",
        [("default_code", "=", scan_input)],
        fields=PRODUCT_FIELDS_TMPL, limit=1
    )
    if records:
        return _build_product_dict(records[0], "product.template")

    return None


def print_product(product, indent="  "):
    sku_warn = "  ← ASSIGN SKU (run odoo_sku_manager.py)" if "NOT ASSIGNED" in product["sku"] else ""
    print(f"{indent}SKU          : {product['sku']}{sku_warn}")
    print(f"{indent}Name         : {product['name']}")
    print(f"{indent}Barcode      : {product['barcode']}")
    print(f"{indent}Unit         : {product['uom']}")
    print(f"{indent}On Hand      : {product['on_hand']}")
    print(f"{indent}Forecasted   : {product['forecasted']}")


def format_scan_match(product):
    """One-line scan confirmation with SKU front and centre."""
    sku = product["sku"]
    return f"[{sku}] {product['name']}  ({product['on_hand']} {product['uom']} on hand)"


# ── GRN — Goods Receipt via Barcode ──────────────────────────────────────────

def mode_grn(client, ref=None, interactive=False):
    """
    GRN flow:
    1. Load an open PO (by ref) or let user enter PO reference.
    2. Scan barcodes one by one → match to PO lines.
    3. Confirm quantities → validate receipt in Odoo.
    """
    print("\n── GRN: GOODS RECEIPT ──────────────────────────────")

    # Load PO
    if not ref:
        ref = input("  Enter PO Reference (e.g. PO/2024/0042): ").strip()

    pos = client.search_read(
        "purchase.order",
        [("name", "=", ref), ("state", "in", ["purchase", "done"])],
        fields=["name", "partner_id", "order_line", "picking_ids"],
        limit=1
    )
    if not pos:
        print(f"  ERROR: PO '{ref}' not found or not confirmed.")
        return

    po = pos[0]
    print(f"\n  PO     : {po['name']}")
    print(f"  Vendor : {po['partner_id'][1] if po.get('partner_id') else 'Unknown'}")

    # Load PO lines
    po_lines = client.search_read(
        "purchase.order.line",
        [("order_id", "=", po["id"]), ("qty_received", "<", "product_qty")],
        fields=["product_id", "product_qty", "qty_received", "product_uom"],
        limit=100
    )

    if not po_lines:
        print("  All items already received on this PO.")
        return

    # Enrich PO lines with SKU
    po_product_ids = [l["product_id"][0] for l in po_lines if l.get("product_id")]
    po_sku_map = {}
    if po_product_ids:
        po_prods = client.search_read(
            "product.product",
            [("id", "in", po_product_ids)],
            fields=["id", "default_code"],
            limit=200
        )
        po_sku_map = {p["id"]: (p.get("default_code") or "NO-SKU") for p in po_prods}

    print(f"\n  Pending items on PO:")
    for line in po_lines:
        prod_id = line["product_id"][0] if line.get("product_id") else None
        name = line["product_id"][1] if line.get("product_id") else "Unknown"
        sku = po_sku_map.get(prod_id, "NO-SKU")
        ordered = line.get("product_qty", 0)
        received = line.get("qty_received", 0)
        pending = ordered - received
        print(f"    [{sku}] {name}: {pending} {line['product_uom'][1] if line.get('product_uom') else 'Units'} pending")

    # Find the incoming picking for this PO
    picking_ids = po.get("picking_ids", [])
    if not picking_ids:
        print("  ERROR: No delivery receipt linked to this PO. Validate PO in Odoo first.")
        return

    picking = client.search_read(
        "stock.picking",
        [("id", "in", picking_ids), ("state", "in", ["assigned", "confirmed", "waiting"])],
        fields=["name", "state", "move_line_ids_without_package"],
        limit=1
    )
    if not picking:
        print("  ERROR: No open receipt found. May already be received.")
        return

    picking = picking[0]
    print(f"\n  Receipt : {picking['name']} ({picking['state']})")

    # Load move lines (items to receive)
    move_lines = client.search_read(
        "stock.move.line",
        [("picking_id", "=", picking["id"]), ("state", "not in", ["done", "cancel"])],
        fields=["product_id", "product_uom_qty", "qty_done", "lot_id", "product_uom_id"],
        limit=100
    )

    # Build barcode → move line map
    product_ids = [ml["product_id"][0] for ml in move_lines if ml.get("product_id")]
    product_barcodes = {}
    if product_ids:
        prods = client.search_read(
            "product.product",
            [("id", "in", product_ids)],
            fields=["id", "barcode", "name", "default_code"],
            limit=200
        )
        product_barcodes = {p["barcode"]: p for p in prods if p.get("barcode")}

    print("\n  Ready to scan. Scan each item barcode (or type 'done' to finish, 'skip' to cancel):\n")

    scanned = {}  # move_line_id → qty_done

    while True:
        raw = input("  SCAN > ").strip()
        if raw.lower() == "done":
            break
        if raw.lower() == "skip":
            print("  GRN cancelled.")
            return

        product = lookup_product(client, raw)
        if not product:
            print(f"    ✗ '{raw}' not found. Try scanning barcode or type SKU (e.g. FW-RM-0042).")
            continue

        # Match to move line
        matched_line = next(
            (ml for ml in move_lines if ml.get("product_id") and ml["product_id"][0] == product["id"]),
            None
        )
        if not matched_line:
            print(f"    ✗ [{product['sku']}] {product['name']} — NOT on this PO. Unexpected item.")
            continue

        expected = matched_line.get("product_uom_qty", 0)
        already_scanned = scanned.get(matched_line["id"], 0)
        remaining = expected - already_scanned

        print(f"    ✓ [{product['sku']}] {product['name']}")
        print(f"      Expected: {expected} | Scanned: {already_scanned} | Remaining: {remaining}")

        if remaining <= 0:
            print(f"      ✗ Already fully received.")
            continue

        qty_input = input(f"      Qty received (Enter for 1): ").strip()
        qty = float(qty_input) if qty_input else 1.0
        qty = min(qty, remaining)

        scanned[matched_line["id"]] = already_scanned + qty
        print(f"      → Recorded: {qty} {product['uom']}")

    if not scanned:
        print("\n  Nothing scanned. GRN aborted.")
        return

    # Summary before commit — enrich with SKU
    ml_prod_ids = [m["product_id"][0] for m in move_lines if m.get("product_id")]
    sku_map = {}
    if ml_prod_ids:
        sku_prods = client.search_read(
            "product.product", [("id", "in", ml_prod_ids)],
            fields=["id", "default_code"], limit=200
        )
        sku_map = {p["id"]: (p.get("default_code") or "NO-SKU") for p in sku_prods}

    print("\n  ── GRN SUMMARY ─────────────────────────────────")
    print(f"  {'SKU':<15} {'Qty':>6}  Product")
    print(f"  {'─'*15} {'─'*6}  {'─'*35}")
    for ml_id, qty in scanned.items():
        ml = next(m for m in move_lines if m["id"] == ml_id)
        prod_id = ml["product_id"][0] if ml.get("product_id") else None
        name = ml["product_id"][1] if ml.get("product_id") else "Unknown"
        sku = sku_map.get(prod_id, "NO-SKU")
        print(f"  {sku:<15} {qty:>6}  {name}")

    confirm = input("\n  Confirm and validate GRN in Odoo? (yes/no): ").strip().lower()
    if confirm != "yes":
        print("  GRN discarded. Nothing was saved.")
        return

    # Write qty_done to move lines and validate picking
    for ml_id, qty in scanned.items():
        client.execute("stock.move.line", "write", [ml_id], {"qty_done": qty})

    # Validate the picking
    try:
        client.execute("stock.picking", "button_validate", [picking["id"]])
        print(f"\n  GRN validated successfully: {picking['name']}")
        print(f"  Stock updated in Odoo. IQC check should now be triggered.")
    except Exception as e:
        if "immediate_transfer" in str(e) or "wizard" in str(e).lower():
            # Odoo may require immediate transfer wizard — handle backorder
            print(f"  Note: Odoo may prompt for backorder. Open {picking['name']} in Odoo to complete.")
        else:
            print(f"  ERROR validating receipt: {e}")


# ── MATERIAL ISSUE — Indent via Barcode ──────────────────────────────────────

def mode_issue(client, ref=None, interactive=False):
    """
    Material Issue flow:
    1. Load an open Manufacturing Order (by MO ref).
    2. Show BOM-required items.
    3. Scan items being picked from stores.
    4. Validate transfer in Odoo (records material consumption against MO).
    """
    print("\n── MATERIAL ISSUE (INDENT) ──────────────────────────")

    if not ref:
        ref = input("  Enter MO Reference (e.g. MO/2024/0017): ").strip()

    mos = client.search_read(
        "mrp.production",
        [("name", "=", ref), ("state", "in", ["confirmed", "progress"])],
        fields=["name", "product_id", "product_qty", "move_raw_ids", "picking_type_id"],
        limit=1
    )
    if not mos:
        print(f"  ERROR: MO '{ref}' not found or not in valid state.")
        return

    mo = mos[0]
    print(f"\n  MO      : {mo['name']}")
    print(f"  Product : {mo['product_id'][1] if mo.get('product_id') else 'Unknown'}")
    print(f"  Qty     : {mo.get('product_qty', 0)}")

    # Load raw material moves (BOM components)
    raw_moves = client.search_read(
        "stock.move",
        [("raw_material_production_id", "=", mo["id"]), ("state", "not in", ["done", "cancel"])],
        fields=["product_id", "product_uom_qty", "quantity_done", "reserved_availability"],
        limit=100
    )

    if not raw_moves:
        print("  All materials already issued for this MO.")
        return

    # Enrich raw moves with SKU
    raw_prod_ids = [mv["product_id"][0] for mv in raw_moves if mv.get("product_id")]
    raw_sku_map = {}
    if raw_prod_ids:
        raw_prods = client.search_read(
            "product.product", [("id", "in", raw_prod_ids)],
            fields=["id", "default_code"], limit=200
        )
        raw_sku_map = {p["id"]: (p.get("default_code") or "NO-SKU") for p in raw_prods}

    print(f"\n  BOM components required:")
    print(f"  {'SKU':<15} {'Need':>6} {'Avail':>7}  Product")
    print(f"  {'─'*15} {'─'*6} {'─'*7}  {'─'*35}")
    for mv in raw_moves:
        prod_id = mv["product_id"][0] if mv.get("product_id") else None
        name = mv["product_id"][1] if mv.get("product_id") else "Unknown"
        sku = raw_sku_map.get(prod_id, "NO-SKU")
        needed = mv.get("product_uom_qty", 0) - mv.get("quantity_done", 0)
        reserved = mv.get("reserved_availability", 0)
        avail_icon = "✓" if reserved >= needed else "!"
        print(f"  [{avail_icon}] {sku:<15} {needed:>6.1f} {reserved:>7.1f}  {name}")

    print("\n  Scan barcode or type SKU (e.g. FW-BT-0012) to pick each item:\n")

    scanned = {}  # move_id → qty

    while True:
        raw = input("  SCAN > ").strip()
        if raw.lower() == "done":
            break
        if raw.lower() == "skip":
            print("  Material issue cancelled.")
            return

        product = lookup_product(client, raw)
        if not product:
            print(f"    ✗ '{raw}' not found. Try scanning barcode or type SKU (e.g. FW-BT-0012).")
            continue

        matched_move = next(
            (mv for mv in raw_moves if mv.get("product_id") and mv["product_id"][0] == product["id"]),
            None
        )
        if not matched_move:
            print(f"    ✗ [{product['sku']}] {product['name']} — NOT in BOM. Extra material requires supervisor approval.")
            log_extra = input("      Log as extra issue anyway? (yes/no): ").strip().lower()
            if log_extra != "yes":
                continue

        needed = matched_move.get("product_uom_qty", 0) if matched_move else 0
        already = scanned.get(matched_move["id"] if matched_move else "extra", 0)

        print(f"    ✓ [{product['sku']}] {product['name']}")
        print(f"      BOM qty: {needed} | Already logged: {already} | To issue: {needed - already:.1f}")
        qty_input = input(f"      Qty issuing (Enter for {needed - already:.1f}): ").strip()
        qty = float(qty_input) if qty_input else (needed - already)

        if matched_move:
            scanned[matched_move["id"]] = already + qty
        print(f"      → Logged: {qty} {product['uom']}")

    if not scanned:
        print("\n  Nothing scanned. Issue aborted.")
        return

    print("\n  ── ISSUE SUMMARY ────────────────────────────────")
    print(f"  {'SKU':<15} {'Qty':>6}  Product")
    print(f"  {'─'*15} {'─'*6}  {'─'*35}")
    for mv_id, qty in scanned.items():
        mv = next((m for m in raw_moves if m["id"] == mv_id), None)
        prod_id = mv["product_id"][0] if mv and mv.get("product_id") else None
        name = mv["product_id"][1] if mv and mv.get("product_id") else "Extra item"
        sku = raw_sku_map.get(prod_id, "NO-SKU")
        print(f"  {sku:<15} {qty:>6.1f}  {name}")

    confirm = input("\n  Confirm and record material issue in Odoo? (yes/no): ").strip().lower()
    if confirm != "yes":
        print("  Issue discarded. Nothing saved.")
        return

    for mv_id, qty in scanned.items():
        client.execute("stock.move", "write", [mv_id], {"quantity_done": qty})

    print(f"\n  Material issue recorded against {mo['name']}.")
    print(f"  Inventory deducted. Production team can proceed.")


# ── CYCLE COUNT — Scan and count ─────────────────────────────────────────────

def mode_count(client, location_name=None):
    """
    Cycle Count flow:
    1. Select inventory location.
    2. Scan item barcode → enter physical count.
    3. Commit adjustments to Odoo.
    """
    print("\n── CYCLE COUNT ─────────────────────────────────────")

    if not location_name:
        location_name = input("  Location to count (e.g. 'Approved Stock'): ").strip()

    # Find location
    locations = client.search_read(
        "stock.location",
        [("complete_name", "ilike", location_name), ("usage", "=", "internal")],
        fields=["id", "complete_name"],
        limit=5
    )
    if not locations:
        print(f"  ERROR: Location '{location_name}' not found.")
        return
    if len(locations) > 1:
        print("  Multiple locations found:")
        for i, loc in enumerate(locations):
            print(f"    {i+1}. {loc['complete_name']}")
        choice = int(input("  Choose (number): ").strip()) - 1
        location = locations[choice]
    else:
        location = locations[0]

    print(f"\n  Counting location: {location['complete_name']}")
    print("  Scan items and enter physical count. Type 'done' when finished.\n")

    counts = []  # list of {product_id, location_id, qty}

    while True:
        raw = input("  SCAN > ").strip()
        if raw.lower() == "done":
            break
        if raw.lower() == "skip":
            print("  Count cancelled.")
            return

        product = lookup_product(client, raw)
        if not product:
            print(f"    ✗ '{raw}' not found. Try scanning barcode or type SKU (e.g. FW-CS-0005).")
            continue

        if not product["id"]:
            print(f"    ✗ Product variant not found for '{raw}'.")
            continue

        # Show current Odoo qty at this location
        quants = client.search_read(
            "stock.quant",
            [("product_id", "=", product["id"]), ("location_id", "=", location["id"])],
            fields=["quantity", "reserved_quantity"],
            limit=1
        )
        system_qty = quants[0]["quantity"] if quants else 0.0

        print(f"    SKU          : {product['sku']}")
        print(f"    Product      : {product['name']}")
        print(f"    System Stock : {system_qty} {product['uom']}")

        qty_input = input(f"    Physical Count: ").strip()
        if not qty_input:
            print("    Skipped.")
            continue

        physical_qty = float(qty_input)
        diff = physical_qty - system_qty
        diff_str = f"+{diff:.2f}" if diff >= 0 else f"{diff:.2f}"
        status = "OK" if diff == 0 else ("EXCESS" if diff > 0 else "SHORT")
        print(f"    → Variance: {diff_str} {product['uom']}  [{status}]")

        counts.append({
            "product_id": product["id"],
            "sku": product["sku"],
            "product_name": product["name"],
            "location_id": location["id"],
            "physical_qty": physical_qty,
            "system_qty": system_qty,
            "variance": diff
        })

    if not counts:
        print("\n  Nothing counted. Aborted.")
        return

    print("\n  ── COUNT SUMMARY ────────────────────────────────")
    variances = [c for c in counts if c["variance"] != 0]
    print(f"  Items counted      : {len(counts)}")
    print(f"  Items with variance: {len(variances)}")
    if variances:
        print(f"\n  {'SKU':<15} {'System':>7} {'Physical':>9} {'Variance':>9}  Product")
        print(f"  {'─'*15} {'─'*7} {'─'*9} {'─'*9}  {'─'*30}")
        for c in variances:
            print(f"  {c.get('sku','NO-SKU'):<15} {c['system_qty']:>7.1f} {c['physical_qty']:>9.1f} {c['variance']:>+9.2f}  {c['product_name']}")

    if not variances:
        print("  No variances found. Inventory is accurate.")
        return

    confirm = input("\n  Apply adjustments to Odoo? (yes/no): ").strip().lower()
    if confirm != "yes":
        print("  Adjustments discarded.")
        return

    # Apply inventory adjustments via stock.quant
    adjusted = 0
    for c in variances:
        try:
            quants = client.search_read(
                "stock.quant",
                [("product_id", "=", c["product_id"]), ("location_id", "=", c["location_id"])],
                fields=["id"],
                limit=1
            )
            if quants:
                client.execute("stock.quant", "write", [quants[0]["id"]], {"inventory_quantity": c["physical_qty"]})
                client.execute("stock.quant", "action_apply_inventory", [quants[0]["id"]])
            else:
                # Create new quant for this product/location
                quant_id = client.execute("stock.quant", "create", {
                    "product_id": c["product_id"],
                    "location_id": c["location_id"],
                    "inventory_quantity": c["physical_qty"]
                })
                client.execute("stock.quant", "action_apply_inventory", [quant_id])
            adjusted += 1
        except Exception as e:
            print(f"    ERROR adjusting {c['product_name']}: {e}")

    print(f"\n  {adjusted}/{len(variances)} adjustments applied to Odoo.")
    print(f"  Inventory updated at: {location['complete_name']}")

    # Save count log
    log_path = f".tmp/cycle_count_{datetime.now().strftime('%Y%m%d_%H%M')}.json"
    os.makedirs(".tmp", exist_ok=True)
    with open(log_path, "w") as f:
        json.dump({
            "date": datetime.now().isoformat(),
            "location": location["complete_name"],
            "counts": counts
        }, f, indent=2)
    print(f"  Count log saved: {log_path}")


# ── DISPATCH — Scan items for Delivery Order ──────────────────────────────────

def mode_dispatch(client, ref=None):
    """
    Dispatch Scan flow:
    1. Load a Delivery Order (stock.picking outgoing).
    2. Scan each item being packed.
    3. Validate that all items are accounted for.
    4. Confirm dispatch → validate picking in Odoo.
    """
    print("\n── DISPATCH SCAN ────────────────────────────────────")

    if not ref:
        ref = input("  Enter Delivery Order ref (e.g. WH/OUT/00123): ").strip()

    pickings = client.search_read(
        "stock.picking",
        [("name", "=", ref), ("picking_type_code", "=", "outgoing"),
         ("state", "in", ["assigned", "confirmed"])],
        fields=["name", "partner_id", "move_line_ids_without_package", "scheduled_date"],
        limit=1
    )
    if not pickings:
        print(f"  ERROR: Delivery Order '{ref}' not found or not ready.")
        return

    picking = pickings[0]
    print(f"\n  Delivery : {picking['name']}")
    print(f"  Customer : {picking['partner_id'][1] if picking.get('partner_id') else 'Unknown'}")
    print(f"  Scheduled: {str(picking.get('scheduled_date', ''))[:10]}")

    # Load move lines
    move_lines = client.search_read(
        "stock.move.line",
        [("picking_id", "=", picking["id"]), ("state", "not in", ["done", "cancel"])],
        fields=["product_id", "product_uom_qty", "qty_done"],
        limit=100
    )

    if not move_lines:
        print("  No items to dispatch on this order.")
        return

    # Enrich move lines with SKU
    disp_prod_ids = [ml["product_id"][0] for ml in move_lines if ml.get("product_id")]
    disp_sku_map = {}
    if disp_prod_ids:
        disp_prods = client.search_read(
            "product.product", [("id", "in", disp_prod_ids)],
            fields=["id", "default_code"], limit=200
        )
        disp_sku_map = {p["id"]: (p.get("default_code") or "NO-SKU") for p in disp_prods}

    print(f"\n  Items to dispatch:")
    print(f"  {'SKU':<15} {'Qty':>6}  Product")
    print(f"  {'─'*15} {'─'*6}  {'─'*35}")
    for ml in move_lines:
        prod_id = ml["product_id"][0] if ml.get("product_id") else None
        name = ml["product_id"][1] if ml.get("product_id") else "Unknown"
        sku = disp_sku_map.get(prod_id, "NO-SKU")
        print(f"  {sku:<15} {ml.get('product_uom_qty', 0):>6.1f}  {name}")

    print("\n  Scan barcode or type SKU of each item being packed:\n")

    scanned = {}  # ml_id → qty_scanned

    while True:
        raw = input("  SCAN > ").strip()
        if raw.lower() == "done":
            break
        if raw.lower() == "skip":
            print("  Dispatch scan cancelled.")
            return

        product = lookup_product(client, raw)
        if not product:
            print(f"    ✗ '{raw}' not found. Try scanning barcode or type SKU (e.g. FW-FG-0003).")
            continue

        matched_line = next(
            (ml for ml in move_lines if ml.get("product_id") and ml["product_id"][0] == product["id"]),
            None
        )
        if not matched_line:
            print(f"    ✗ [{product['sku']}] {product['name']} — NOT on this delivery. Do NOT ship.")
            continue

        expected = matched_line.get("product_uom_qty", 0)
        already = scanned.get(matched_line["id"], 0)
        print(f"    ✓ [{product['sku']}] {product['name']}")
        print(f"      Expected: {expected} | Scanned so far: {already}")

        qty_input = input(f"      Qty (Enter for 1): ").strip()
        qty = float(qty_input) if qty_input else 1.0
        scanned[matched_line["id"]] = already + qty
        print(f"      → Total scanned: {scanned[matched_line['id']]}")

    if not scanned:
        print("\n  Nothing scanned. Dispatch aborted.")
        return

    # Validation: check all items fully scanned
    print("\n  ── DISPATCH VALIDATION ──────────────────────────")
    print(f"  {'SKU':<15} {'Expected':>9} {'Scanned':>8}  Status   Product")
    print(f"  {'─'*15} {'─'*9} {'─'*8}  {'─'*8}  {'─'*30}")
    all_ok = True
    for ml in move_lines:
        prod_id = ml["product_id"][0] if ml.get("product_id") else None
        name = ml["product_id"][1] if ml.get("product_id") else "Unknown"
        sku = disp_sku_map.get(prod_id, "NO-SKU")
        expected = ml.get("product_uom_qty", 0)
        actual = scanned.get(ml["id"], 0)
        status = "✓ OK" if actual >= expected else "✗ SHORT"
        if actual < expected:
            all_ok = False
        print(f"  {sku:<15} {expected:>9.1f} {actual:>8.1f}  {status:<8}  {name}")
        print(f"    [{status}] {name}: expected={expected}, scanned={actual}")

    if not all_ok:
        print("\n  WARNING: Some items are short-shipped.")
        proceed = input("  Proceed with partial dispatch? (yes/no): ").strip().lower()
        if proceed != "yes":
            print("  Dispatch aborted. Resolve shortfall before dispatching.")
            return

    confirm = input("\n  Confirm dispatch and validate in Odoo? (yes/no): ").strip().lower()
    if confirm != "yes":
        print("  Dispatch discarded.")
        return

    # Write qty_done to move lines and validate
    for ml_id, qty in scanned.items():
        client.execute("stock.move.line", "write", [ml_id], {"qty_done": qty})

    try:
        client.execute("stock.picking", "button_validate", [picking["id"]])
        print(f"\n  Delivery Order {picking['name']} validated.")
        print(f"  Stock reduced. Invoice can now be generated.")
    except Exception as e:
        print(f"  ERROR validating delivery: {e}")
        print(f"  Open {picking['name']} in Odoo to complete manually.")


# ── LOOKUP MODE ───────────────────────────────────────────────────────────────

def mode_lookup(client, barcode):
    """Simple barcode lookup — no writes."""
    if not barcode:
        barcode = input("  Scan or enter barcode: ").strip()

    product = lookup_product(client, barcode)
    if not product:
        print(f"  No product found for barcode: {barcode}")
        print("  Check: Odoo → Products → [Product] → Barcode field")
        return

    print(f"\n  ── PRODUCT FOUND ────────────────────────────────")
    print_product(product)

    # Also show location-wise stock
    if product["id"]:
        quants = client.search_read(
            "stock.quant",
            [("product_id", "=", product["id"]), ("location_id.usage", "=", "internal")],
            fields=["location_id", "quantity", "reserved_quantity"],
            limit=20
        )
        if quants:
            print(f"\n  Stock by Location:")
            for q in quants:
                loc = q["location_id"][1] if q.get("location_id") else "Unknown"
                available = q.get("quantity", 0) - q.get("reserved_quantity", 0)
                print(f"    {loc:<30} On Hand: {q.get('quantity',0):>8}  Available: {available:>8}")


# ── MAIN ──────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Barcode-based inventory operations for Fracktal Works Odoo"
    )
    parser.add_argument(
        "--mode",
        choices=["lookup", "grn", "issue", "count", "dispatch"],
        required=True,
        help="Operation mode"
    )
    parser.add_argument("--barcode", help="Barcode string (for lookup mode)")
    parser.add_argument("--ref", help="PO/MO/Delivery Order reference")
    parser.add_argument("--location", help="Location name (for count mode)")
    parser.add_argument("--interactive", action="store_true",
                        help="Run interactive scanner session")
    args = parser.parse_args()

    print(f"\nFracktal Works — Barcode Scanner [{args.mode.upper()}]")
    print(f"Connecting to Odoo: {os.getenv('ODOO_URL', '(not set)')}")

    client = OdooClient()
    _ = client.uid  # Trigger auth early, fail fast

    if args.mode == "lookup":
        mode_lookup(client, args.barcode)
    elif args.mode == "grn":
        mode_grn(client, ref=args.ref, interactive=args.interactive)
    elif args.mode == "issue":
        mode_issue(client, ref=args.ref, interactive=args.interactive)
    elif args.mode == "count":
        mode_count(client, location_name=args.location)
    elif args.mode == "dispatch":
        mode_dispatch(client, ref=args.ref)


if __name__ == "__main__":
    main()
