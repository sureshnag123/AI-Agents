"""
odoo_sku_manager.py — SKU generation and management for Fracktal Works inventory.

SKU Format: FW-[CATEGORY]-[0001]

Categories:
  RM  = Raw Material        (steel, aluminium, rods, sheets)
  BT  = Bought-Out Part     (bearings, motors, fasteners, purchased components)
  SF  = Semi-Finished       (machined parts, fabricated sub-parts)
  AS  = Assembly            (sub-assemblies built in-house)
  FG  = Finished Goods      (saleable end products)
  SP  = Spare Part          (service/replacement parts)
  CS  = Consumable          (welding rods, cutting fluid, sandpaper, tape)
  PK  = Packaging           (boxes, foam, straps)

Each product gets a UNIQUE, PERMANENT SKU. Once assigned, never change it.
SKU = Odoo Internal Reference (default_code field on product.template)

Usage:
  python execution/odoo_sku_manager.py --mode list-missing
  python execution/odoo_sku_manager.py --mode list-all --output .tmp/sku_list.json
  python execution/odoo_sku_manager.py --mode preview --category RM
  python execution/odoo_sku_manager.py --mode assign --product-id 42 --sku FW-RM-0001
  python execution/odoo_sku_manager.py --mode assign-bulk --category RM --dry-run
  python execution/odoo_sku_manager.py --mode assign-bulk --category RM
  python execution/odoo_sku_manager.py --mode export --output .tmp/sku_master.csv
"""

import argparse
import csv
import json
import os
import sys
from datetime import datetime

sys.path.insert(0, os.path.dirname(__file__))
from odoo_connect import OdooClient


# ── SKU Configuration ─────────────────────────────────────────────────────────

SKU_PREFIX = "FW"

CATEGORY_CODES = {
    "RM": "Raw Material",
    "BT": "Bought-Out Part",
    "SF": "Semi-Finished",
    "AS": "Assembly / Sub-Assembly",
    "FG": "Finished Goods",
    "SP": "Spare Part",
    "CS": "Consumable",
    "PK": "Packaging",
}

# Odoo product category name fragments → SKU category mapping (case-insensitive)
CATEGORY_AUTO_MAP = {
    "raw":         "RM",
    "material":    "RM",
    "steel":       "RM",
    "aluminium":   "RM",
    "aluminum":    "RM",
    "bought":      "BT",
    "purchase":    "BT",
    "component":   "BT",
    "bearing":     "BT",
    "fastener":    "BT",
    "motor":       "BT",
    "semi":        "SF",
    "machined":    "SF",
    "fabricat":    "SF",
    "assembly":    "AS",
    "sub-asm":     "AS",
    "finished":    "FG",
    "product":     "FG",
    "saleable":    "FG",
    "spare":       "SP",
    "service":     "SP",
    "consumable":  "CS",
    "consumab":    "CS",
    "packaging":   "PK",
    "pack":        "PK",
}


def format_sku(category, number):
    return f"{SKU_PREFIX}-{category}-{number:04d}"


def parse_sku(sku):
    """Parse 'FW-RM-0042' → ('RM', 42). Returns None if invalid."""
    parts = sku.split("-")
    if len(parts) == 3 and parts[0] == SKU_PREFIX and parts[1] in CATEGORY_CODES:
        try:
            return parts[1], int(parts[2])
        except ValueError:
            pass
    return None


def auto_detect_category(product_name, category_name):
    """Guess SKU category from product or Odoo category name."""
    text = f"{product_name} {category_name}".lower()
    for keyword, cat in CATEGORY_AUTO_MAP.items():
        if keyword in text:
            return cat
    return None  # Cannot auto-detect


def get_next_sku_number(client, category):
    """Find the highest existing SKU number for this category and return next."""
    prefix = f"{SKU_PREFIX}-{category}-"
    existing = client.search_read(
        "product.template",
        [("default_code", "like", prefix)],
        fields=["default_code"],
        limit=1000
    )
    numbers = []
    for p in existing:
        code = p.get("default_code", "")
        if code.startswith(prefix):
            suffix = code[len(prefix):]
            try:
                numbers.append(int(suffix))
            except ValueError:
                pass
    return max(numbers, default=0) + 1


# ── Modes ─────────────────────────────────────────────────────────────────────

def list_missing(client):
    """List all products with no SKU (Internal Reference is empty)."""
    print("\nProducts without SKU (Internal Reference):\n")
    products = client.search_read(
        "product.template",
        [("default_code", "in", [False, ""])],
        fields=["id", "name", "type", "categ_id"],
        limit=500
    )
    if not products:
        print("  All products have SKUs assigned.")
        return []

    print(f"  {'ID':<6} {'Type':<10} {'Category':<25} {'Name'}")
    print(f"  {'─'*6} {'─'*10} {'─'*25} {'─'*40}")
    for p in products:
        cat = p["categ_id"][1] if p.get("categ_id") else "—"
        ptype = p.get("type", "—")
        print(f"  {p['id']:<6} {ptype:<10} {cat:<25} {p['name']}")

    print(f"\n  Total missing SKU: {len(products)}")
    print(f"\n  Next step:")
    print(f"    python execution/odoo_sku_manager.py --mode assign-bulk --category RM --dry-run")
    return products


def list_all(client, output=None):
    """List all products with their SKU and barcode."""
    print("\nFull Product SKU / Barcode Master List:\n")
    products = client.search_read(
        "product.template",
        [],
        fields=["id", "name", "default_code", "barcode", "type", "categ_id",
                "qty_available", "uom_id"],
        limit=2000
    )

    print(f"  {'SKU':<15} {'Barcode':<16} {'Type':<10} {'On Hand':>8}  {'Name'}")
    print(f"  {'─'*15} {'─'*16} {'─'*10} {'─'*8}  {'─'*40}")
    for p in products:
        sku = p.get("default_code") or "— NO SKU —"
        barcode = p.get("barcode") or "—"
        ptype = p.get("type", "—")
        qty = p.get("qty_available", 0)
        print(f"  {sku:<15} {barcode:<16} {ptype:<10} {qty:>8.1f}  {p['name']}")

    no_sku = sum(1 for p in products if not p.get("default_code"))
    no_barcode = sum(1 for p in products if not p.get("barcode"))
    print(f"\n  Total products : {len(products)}")
    print(f"  Missing SKU    : {no_sku}")
    print(f"  Missing Barcode: {no_barcode}")

    if output:
        ext = os.path.splitext(output)[1].lower()
        os.makedirs(os.path.dirname(output) if os.path.dirname(output) else ".", exist_ok=True)
        if ext == ".csv":
            with open(output, "w", newline="", encoding="utf-8") as f:
                writer = csv.DictWriter(f, fieldnames=["id", "sku", "barcode", "name", "type", "on_hand"])
                writer.writeheader()
                for p in products:
                    writer.writerow({
                        "id": p["id"],
                        "sku": p.get("default_code") or "",
                        "barcode": p.get("barcode") or "",
                        "name": p["name"],
                        "type": p.get("type", ""),
                        "on_hand": p.get("qty_available", 0)
                    })
        else:
            with open(output, "w", encoding="utf-8") as f:
                json.dump(products, f, indent=2, default=str)
        print(f"\n  Exported to: {output}")

    return products


def preview_bulk_assign(client, category):
    """Show what SKUs would be assigned — dry run."""
    print(f"\nPREVIEW: Auto-assign SKUs — Category: {category} ({CATEGORY_CODES.get(category, '?')})\n")

    products = client.search_read(
        "product.template",
        [("default_code", "in", [False, ""])],
        fields=["id", "name", "categ_id", "type"],
        limit=500
    )

    next_num = get_next_sku_number(client, category)
    eligible = []

    for p in products:
        cat_name = p["categ_id"][1] if p.get("categ_id") else ""
        auto_cat = auto_detect_category(p["name"], cat_name)
        if auto_cat == category or category == "ALL":
            sku = format_sku(category, next_num)
            print(f"  Would assign {sku} → {p['name']}")
            eligible.append((p["id"], sku, p["name"]))
            next_num += 1

    if not eligible:
        print(f"  No products auto-detected for category '{category}'.")
        print(f"  Use --mode assign --product-id [ID] --sku [SKU] to assign manually.")
    else:
        print(f"\n  {len(eligible)} SKUs would be assigned.")
        print(f"  Run without --dry-run to apply.")

    return eligible


def assign_single(client, product_id, sku):
    """Assign a specific SKU to a specific product."""
    # Validate SKU format
    parsed = parse_sku(sku)
    if not parsed:
        print(f"  ERROR: Invalid SKU format '{sku}'.")
        print(f"  Expected format: FW-[CATEGORY]-[NUMBER] e.g. FW-RM-0001")
        print(f"  Valid categories: {', '.join(CATEGORY_CODES.keys())}")
        return False

    # Check SKU not already used
    existing = client.search_read(
        "product.template",
        [("default_code", "=", sku)],
        fields=["id", "name"],
        limit=1
    )
    if existing:
        print(f"  ERROR: SKU '{sku}' already assigned to '{existing[0]['name']}' (ID {existing[0]['id']}).")
        return False

    # Get product info
    product = client.search_read(
        "product.template",
        [("id", "=", product_id)],
        fields=["name", "default_code"],
        limit=1
    )
    if not product:
        print(f"  ERROR: Product ID {product_id} not found.")
        return False

    p = product[0]
    if p.get("default_code"):
        print(f"  WARNING: '{p['name']}' already has SKU: {p['default_code']}")
        confirm = input("  Overwrite? (yes/no): ").strip().lower()
        if confirm != "yes":
            return False

    client.execute("product.template", "write", [product_id], {"default_code": sku})
    print(f"  Assigned: {sku} → {p['name']}")
    return True


def assign_bulk(client, category, dry_run=False):
    """Auto-assign SKUs to unassigned products in a category."""
    eligible = preview_bulk_assign(client, category)

    if not eligible:
        return

    if dry_run:
        print("\n  (Dry run — no changes made)")
        return

    print(f"\n  Assigning {len(eligible)} SKUs...")
    confirm = input("  Confirm bulk assignment? (yes/no): ").strip().lower()
    if confirm != "yes":
        print("  Aborted.")
        return

    assigned = 0
    for product_id, sku, name in eligible:
        try:
            client.execute("product.template", "write", [product_id], {"default_code": sku})
            print(f"  ✓ {sku} → {name}")
            assigned += 1
        except Exception as e:
            print(f"  ✗ Failed for {name}: {e}")

    print(f"\n  Done. {assigned}/{len(eligible)} SKUs assigned.")

    # Auto-export after assignment
    log_path = f".tmp/sku_assignment_{datetime.now().strftime('%Y%m%d_%H%M')}.json"
    os.makedirs(".tmp", exist_ok=True)
    with open(log_path, "w") as f:
        json.dump([{"product_id": pid, "sku": sku, "name": name}
                   for pid, sku, name in eligible], f, indent=2)
    print(f"  Log saved: {log_path}")


def export_master(client, output):
    """Export full SKU + barcode master list for label printing."""
    products = client.search_read(
        "product.template",
        [],
        fields=["id", "name", "default_code", "barcode", "type", "categ_id",
                "uom_id", "qty_available"],
        limit=2000
    )

    os.makedirs(os.path.dirname(output) if os.path.dirname(output) else ".", exist_ok=True)
    with open(output, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["SKU", "Barcode", "Product Name", "Category", "Type",
                         "Unit", "On Hand", "Label Text"])
        for p in products:
            sku = p.get("default_code") or "UNASSIGNED"
            barcode = p.get("barcode") or ""
            cat = p["categ_id"][1] if p.get("categ_id") else ""
            uom = p["uom_id"][1] if p.get("uom_id") else "Units"
            # Label text: what gets printed on physical product label
            label = f"{sku} | {p['name']}"
            writer.writerow([sku, barcode, p["name"], cat, p.get("type", ""),
                             uom, p.get("qty_available", 0), label])

    print(f"\n  SKU Master exported to: {output}")
    print(f"  Use this file to:")
    print(f"    1. Print product labels (import into label printer software)")
    print(f"    2. Verify SKU ↔ barcode mapping before scanning")
    print(f"    3. Share with purchase team for PO line references")


# ── Main ─────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="SKU Manager for Fracktal Works Odoo inventory")
    parser.add_argument("--mode", required=True,
                        choices=["list-missing", "list-all", "preview", "assign", "assign-bulk", "export"],
                        help="Operation mode")
    parser.add_argument("--category", choices=list(CATEGORY_CODES.keys()) + ["ALL"],
                        help="SKU category (RM, BT, SF, AS, FG, SP, CS, PK)")
    parser.add_argument("--product-id", type=int, help="Odoo product.template ID")
    parser.add_argument("--sku", help="SKU to assign (e.g. FW-RM-0001)")
    parser.add_argument("--dry-run", action="store_true", help="Preview without making changes")
    parser.add_argument("--output", help="Output file path (.json or .csv)")
    args = parser.parse_args()

    print(f"\nFracktal Works — SKU Manager")
    print(f"SKU Format: FW-[CATEGORY]-[0001]")
    print(f"Categories: {', '.join(f'{k}={v}' for k, v in CATEGORY_CODES.items())}\n")

    client = OdooClient()
    _ = client.uid

    if args.mode == "list-missing":
        list_missing(client)

    elif args.mode == "list-all":
        list_all(client, output=args.output)

    elif args.mode == "preview":
        if not args.category:
            print("ERROR: --category required for preview mode")
            sys.exit(1)
        preview_bulk_assign(client, args.category)

    elif args.mode == "assign":
        if not args.product_id or not args.sku:
            print("ERROR: --product-id and --sku required for assign mode")
            sys.exit(1)
        assign_single(client, args.product_id, args.sku)

    elif args.mode == "assign-bulk":
        if not args.category:
            print("ERROR: --category required for assign-bulk mode")
            sys.exit(1)
        assign_bulk(client, args.category, dry_run=args.dry_run)

    elif args.mode == "export":
        output = args.output or ".tmp/sku_master.csv"
        export_master(client, output)


if __name__ == "__main__":
    main()
