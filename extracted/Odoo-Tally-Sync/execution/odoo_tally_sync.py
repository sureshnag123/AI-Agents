#!/usr/bin/env python3
"""
Odoo → Tally Prime Sync Pipeline

Main orchestration script that:
1. Fetches invoices/payments from Odoo 19 via XML-RPC
2. Transforms them into Tally Prime XML vouchers
3. Pushes them to Tally via XML Server (port 9000)
4. Tracks sync state to prevent duplicates

Usage:
    # Dry run (preview what will sync)
    python odoo_tally_sync.py --dry-run

    # Sync all voucher types (incremental)
    python odoo_tally_sync.py

    # Sync specific types
    python odoo_tally_sync.py --types "purchase,sales"

    # Full sync for a date range
    python odoo_tally_sync.py --mode full --from 2026-02-01 --to 2026-02-26

    # Generate ledger mapping
    python odoo_tally_sync.py --generate-mapping

    # Rollback last batch
    python odoo_tally_sync.py --rollback --batch-id BATCH_2026_02_26_001
"""

import os
import sys
import json
import argparse
import logging
import xml.etree.ElementTree as ET
from datetime import datetime, date, timedelta
from pathlib import Path
from typing import Optional

from dotenv import load_dotenv

SCRIPT_DIR = Path(__file__).parent.resolve()
PROJECT_ROOT = SCRIPT_DIR.parent
TMP_DIR = PROJECT_ROOT / ".tmp"

load_dotenv(PROJECT_ROOT / ".env")

# Add execution dir to path for local imports
sys.path.insert(0, str(SCRIPT_DIR))

from odoo_connector import OdooConnector
from tally_connector import TallyConnector
from tally_xml_builder import TallyXMLBuilder

# -- Logging --------------------------------------------------------

LOG_DIR = TMP_DIR / "logs"
LOG_DIR.mkdir(parents=True, exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler(LOG_DIR / f"sync_{date.today().isoformat()}.log"),
    ],
)
log = logging.getLogger("odoo_tally_sync")

# -- Sync State Management ------------------------------------------

SYNC_LOG_FILE = TMP_DIR / "sync_log.json"
MAPPING_FILE = TMP_DIR / "ledger_mapping.json"

VOUCHER_TYPE_MAP = {
    "purchase": {
        "description": "Purchase Invoices (Vendor Bills)",
        "odoo_fetch": "get_purchase_invoices",
        "builder_method": "build_purchase_voucher",
        "tally_type": "Purchase",
        "move_type": "in_invoice",
    },
    "purchase_return": {
        "description": "Debit Notes (Purchase Returns)",
        "odoo_fetch": "get_purchase_returns",
        "builder_method": "build_debit_note",
        "tally_type": "Debit Note",
        "move_type": "in_refund",
    },
    "sales": {
        "description": "Sales Invoices",
        "odoo_fetch": "get_sales_invoices",
        "builder_method": "build_sales_voucher",
        "tally_type": "Sales",
        "move_type": "out_invoice",
        # Also check GST INVOICE type to prevent posting same invoice under both types
        "also_check_tally_types": ["GST INVOICE"],
    },
    "gst_invoice": {
        "description": "Sales Invoices (GST INVOICE)",
        "odoo_fetch": "get_sales_invoices",
        "builder_method": "build_gst_invoice_voucher",
        "tally_type": "GST INVOICE",
        "move_type": "out_invoice",
        # Also check Sales type to prevent posting same invoice under both types
        "also_check_tally_types": ["Sales"],
    },
    "sales_return": {
        "description": "Credit Notes (Sales Returns)",
        "odoo_fetch": "get_sales_returns",
        "builder_method": "build_credit_note",
        "tally_type": "Credit Note",
        "move_type": "out_refund",
    },
    "payment": {
        "description": "Outbound Payments (Vendor Payments)",
        "odoo_fetch": "get_outbound_payments",
        "builder_method": "build_payment_voucher",
        "tally_type": "Payment",
    },
    "receipt": {
        "description": "Inbound Payments (Customer Receipts)",
        "odoo_fetch": "get_inbound_payments",
        "builder_method": "build_receipt_voucher",
        "tally_type": "Receipt",
    },
}


def load_sync_log() -> dict:
    """Load the sync state log."""
    if SYNC_LOG_FILE.exists():
        with open(SYNC_LOG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {
        "last_sync": None,
        "batches": [],
        "synced_ids": {},  # {"purchase": [id1, id2], "sales": [...]}
    }


def save_sync_log(sync_log: dict):
    """Persist the sync state log."""
    with open(SYNC_LOG_FILE, "w", encoding="utf-8") as f:
        json.dump(sync_log, f, indent=2, default=str)


def load_mapping() -> dict:
    """Load ledger mapping file."""
    if MAPPING_FILE.exists():
        with open(MAPPING_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {
        "odoo_to_tally_ledgers": {},
        "odoo_to_tally_partners": {},
        "auto_create_missing_ledgers": True,
        "default_group_for_new_creditors": "Sundry Creditors",
        "default_group_for_new_debtors": "Sundry Debtors",
        "default_bank_ledger": "Bank Accounts",
    }


def save_mapping(mapping: dict):
    """Persist ledger mapping."""
    with open(MAPPING_FILE, "w", encoding="utf-8") as f:
        json.dump(mapping, f, indent=2)


# -- Ledger Mapping Generator ---------------------------------------

def generate_mapping(odoo: OdooConnector, tally: TallyConnector):
    """
    Generate an initial ledger mapping by fuzzy-matching
    Odoo accounts/partners with Tally ledgers.
    """
    log.info("Generating ledger mapping...")

    # Fetch from both systems
    odoo_accounts = odoo.get_chart_of_accounts()
    odoo_partners_sup = odoo.get_all_partners(supplier=True)
    odoo_partners_cust = odoo.get_all_partners(customer=True)

    try:
        tally_ledgers = tally.get_all_ledgers()
    except ConnectionError as e:
        log.warning(f"Could not connect to Tally: {e}")
        log.info("Creating mapping with Odoo data only. Fill in Tally names manually.")
        tally_ledgers = []

    tally_names = {l["name"].lower(): l["name"] for l in tally_ledgers}

    mapping = load_mapping()

    # Map accounts
    for acc in odoo_accounts:
        odoo_name = acc.get("name")
        if not odoo_name or not isinstance(odoo_name, str):
            continue
        if odoo_name not in mapping["odoo_to_tally_ledgers"]:
            # Try exact match (case-insensitive)
            if odoo_name.lower() in tally_names:
                mapping["odoo_to_tally_ledgers"][odoo_name] = tally_names[odoo_name.lower()]
            else:
                # Try partial match (require at least 6 chars to avoid false positives)
                matched = False
                odoo_lower = odoo_name.lower()
                for tally_lower, tally_original in tally_names.items():
                    if len(tally_lower) >= 6 and len(odoo_lower) >= 6:
                        if odoo_lower in tally_lower or tally_lower in odoo_lower:
                            mapping["odoo_to_tally_ledgers"][odoo_name] = tally_original
                            matched = True
                            break
                if not matched:
                    # Mark as TODO -- user should fill this in
                    mapping["odoo_to_tally_ledgers"][odoo_name] = f"TODO: {odoo_name}"

    # Map partners (suppliers)
    for partner in odoo_partners_sup:
        p_name = partner.get("name")
        if not p_name or not isinstance(p_name, str):
            continue
        if p_name not in mapping["odoo_to_tally_partners"]:
            if p_name.lower() in tally_names:
                mapping["odoo_to_tally_partners"][p_name] = tally_names[p_name.lower()]
            else:
                mapping["odoo_to_tally_partners"][p_name] = p_name  # Keep same name

    # Map partners (customers)
    for partner in odoo_partners_cust:
        p_name = partner.get("name")
        if not p_name or not isinstance(p_name, str):
            continue
        if p_name not in mapping["odoo_to_tally_partners"]:
            if p_name.lower() in tally_names:
                mapping["odoo_to_tally_partners"][p_name] = tally_names[p_name.lower()]
            else:
                mapping["odoo_to_tally_partners"][p_name] = p_name

    save_mapping(mapping)
    log.info(f"Mapping saved to {MAPPING_FILE}")
    log.info(f"  Account mappings: {len(mapping['odoo_to_tally_ledgers'])}")
    log.info(f"  Partner mappings: {len(mapping['odoo_to_tally_partners'])}")

    # Count TODOs
    todos = sum(1 for v in mapping["odoo_to_tally_ledgers"].values() if v.startswith("TODO:"))
    if todos:
        log.warning(f"  [WARN] {todos} accounts need manual Tally mapping (marked 'TODO:' in the file)")

    return mapping


# -- Ensure Ledger Exists in Tally -----------------------------------

def _enrich_lines_with_product_type(odoo: OdooConnector, lines: list):
    """
    Add '_product_type' key to each line by batch-fetching product info.
    This is used by TallyXMLBuilder to route dynamic accounts like
    '200110 Local Sales' to the correct Tally ledger.
    """
    product_ids = set()
    for line in lines:
        prod = line.get("product_id")
        if isinstance(prod, (list, tuple)):
            product_ids.add(prod[0])
    if not product_ids:
        return
    try:
        products = odoo.read("product.product", list(product_ids), ["id", "type"])
        type_map = {p["id"]: p.get("type", "") for p in products}
    except Exception:
        type_map = {}
    for line in lines:
        prod = line.get("product_id")
        if isinstance(prod, (list, tuple)):
            line["_product_type"] = type_map.get(prod[0], "")


def ensure_ledger_in_tally(tally: TallyConnector, name: str,
                           group: str, mapping: dict,
                           tally_ledger_cache: set,
                           partner_details: dict = None) -> bool:
    """
    Check if a ledger exists in Tally. If not, create it
    (if auto_create_missing_ledgers is True).

    partner_details: Odoo res.partner record dict. When supplied, address,
    state, pincode, phone, mobile, email, GST and PAN are written to the
    new Tally ledger automatically.

    Returns True if the ledger is usable.
    """
    if name.lower() in tally_ledger_cache:
        return True

    if not mapping.get("auto_create_missing_ledgers", False):
        log.warning(f"Ledger '{name}' not found in Tally and auto-create is disabled")
        return False

    # ── Build contact/tax kwargs from Odoo partner record ───────────
    kwargs = {}
    if partner_details:
        p = partner_details

        # Address — build multi-line list: street, street2, city+zip
        addr_lines = []
        if p.get("street"):
            addr_lines.append(p["street"])
        if p.get("street2"):
            addr_lines.append(p["street2"])
        city_zip = " - ".join(filter(None, [p.get("city"), p.get("zip")]))
        if city_zip:
            addr_lines.append(city_zip)
        if addr_lines:
            kwargs["address_lines"] = addr_lines

        # State name (Odoo returns [id, name])
        state_val = p.get("state_id")
        if isinstance(state_val, (list, tuple)) and len(state_val) >= 2:
            kwargs["state"] = state_val[1]

        # Country name
        country_val = p.get("country_id")
        if isinstance(country_val, (list, tuple)) and len(country_val) >= 2:
            kwargs["country"] = country_val[1]

        if p.get("zip"):
            kwargs["pincode"] = p["zip"]
        if p.get("phone"):
            kwargs["phone"] = p["phone"]
        if p.get("mobile"):
            kwargs["mobile"] = p["mobile"]
        if p.get("email"):
            kwargs["email"] = p["email"]

        # GST / Tax registration
        if p.get("vat"):
            kwargs["gst_number"] = p["vat"]
        if p.get("l10n_in_gst_treatment"):
            kwargs["gst_registration_type"] = p["l10n_in_gst_treatment"]

        # PAN (Indian localization field)
        if p.get("l10n_in_pan"):
            kwargs["pan"] = p["l10n_in_pan"]

    log.info(f"Creating missing ledger in Tally: '{name}' under '{group}'")
    result = tally.create_ledger(name, group, **kwargs)

    if result.startswith("OK") or "already" in result.lower():
        # "already exists" response means the ledger is there — treat as success
        tally_ledger_cache.add(name.lower())
        return True
    else:
        log.error(f"Failed to create ledger '{name}': {result}")
        return False


# -- Voucher Preview Helpers ----------------------------------------

def _extract_ledger_entries(xml_str: str) -> list:
    """Return [(ledger_name, amount)] from ALLLEDGERENTRIES.LIST in a voucher XML."""
    try:
        root = ET.fromstring(xml_str)
        entries = []
        for el in root.iter("ALLLEDGERENTRIES.LIST"):
            name = (el.findtext("LEDGERNAME") or "").strip()
            amt_text = (el.findtext("AMOUNT") or "0").strip()
            try:
                amt = float(amt_text)
            except ValueError:
                amt = 0.0
            if name:
                entries.append((name, amt))
        return entries
    except ET.ParseError:
        return []


def _print_verify_table(built_records: list, vtype_desc: str) -> None:
    """Print a preview table of vouchers about to be posted, including ledger breakdown."""
    print(f"\n  {'-'*72}")
    print(f"  VERIFY - {vtype_desc} ({len(built_records)} voucher(s))")
    print(f"  {'-'*72}")
    for record, _lines, xml in built_records:
        name = record.get("name", f"ID-{record['id']}")
        partner = record["partner_id"][1] if record.get("partner_id") else "N/A"
        amt = record.get("amount_total", record.get("amount", 0))
        inv_date = record.get("invoice_date", record.get("date", ""))
        print(f"\n  Voucher : {name}")
        print(f"  Party   : {partner}")
        print(f"  Date    : {inv_date}   Amount: Rs.{amt:,.2f}")
        print(f"  Ledgers :")
        for ledger, amount in _extract_ledger_entries(xml):
            dr_cr = "Dr" if amount < 0 else "Cr"
            print(f"            {dr_cr}  {abs(amount):>12,.2f}   {ledger}")
    print(f"\n  {'-'*72}")


# -- Main Sync Pipeline ---------------------------------------------

def sync_vouchers(
    odoo: OdooConnector,
    tally: TallyConnector,
    builder: TallyXMLBuilder,
    voucher_types: list,
    from_date: Optional[str] = None,
    to_date: Optional[str] = None,
    mode: str = "incremental",
    dry_run: bool = False,
    verify: bool = False,
    batch_size: int = 100,
    voucher_from: Optional[str] = None,
    voucher_to: Optional[str] = None,
) -> dict:
    """
    Main sync pipeline.

    Returns a summary dict with counts per voucher type.
    """
    mapping = load_mapping()
    sync_log = load_sync_log()
    bank_ledger = mapping.get("default_bank_ledger", "Bank Accounts")

    # Build Tally ledger cache for quick lookups
    tally_ledger_cache = set()
    if not dry_run:
        try:
            existing = tally.get_all_ledgers()
            tally_ledger_cache = {l["name"].lower() for l in existing}
        except ConnectionError:
            log.warning("Could not fetch Tally ledgers for caching")

    # Determine date range
    if not from_date:
        if mode == "incremental" and sync_log.get("last_sync"):
            from_date = sync_log["last_sync"][:10]  # YYYY-MM-DD
        else:
            from_date = (date.today() - timedelta(days=7)).isoformat()
    if not to_date:
        to_date = date.today().isoformat()

    log.info(f"{'DRY RUN — ' if dry_run else ''}Sync from {from_date} to {to_date} ({mode} mode)")
    log.info(f"Voucher types: {', '.join(voucher_types)}")

    batch_id = f"BATCH_{datetime.now().strftime('%Y_%m_%d_%H%M%S')}"
    batch_results = {
        "batch_id": batch_id,
        "started_at": datetime.now().isoformat(),
        "mode": mode,
        "from_date": from_date,
        "to_date": to_date,
        "dry_run": dry_run,
        "types": {},
    }

    total_synced = 0
    total_skipped = 0
    total_errors = 0

    for vtype in voucher_types:
        config = VOUCHER_TYPE_MAP.get(vtype)
        if not config:
            log.warning(f"Unknown voucher type: {vtype}")
            continue

        log.info(f"\n{'-'*60}")
        log.info(f"Processing: {config['description']}")

        type_result = {
            "fetched": 0,
            "synced": 0,
            "skipped": 0,
            "errors": 0,
            "details": [],
        }

        # Fetch records from Odoo
        fetch_fn = getattr(odoo, config["odoo_fetch"])

        if vtype in ("payment", "receipt"):
            records = fetch_fn(from_date=from_date, to_date=to_date)
        else:
            records = fetch_fn(from_date=from_date, to_date=to_date)

        type_result["fetched"] = len(records)
        log.info(f"  Fetched {len(records)} records from Odoo")

        # Filter by voucher number range if specified
        if voucher_from or voucher_to:
            before_filter = len(records)
            records = [
                r for r in records
                if (not voucher_from or r.get("name", "") >= voucher_from)
                and (not voucher_to or r.get("name", "") <= voucher_to)
            ]
            log.info(f"  Voucher range filter: {before_filter} -> {len(records)} records "
                     f"({voucher_from or '*'} to {voucher_to or '*'})")

        # Filter already-synced records (incremental mode)
        synced_ids = set(sync_log.get("synced_ids", {}).get(vtype, []))

        # ── Pre-check: fetch existing voucher numbers from Tally ─────
        # Prevents duplicates regardless of sync log state (e.g. after a
        # --mode full re-run). Skips any record whose voucher number is
        # already present in Tally for the given date range.
        tally_existing_numbers = set()
        if not dry_run and config.get("tally_type"):
            tally_from = from_date.replace("-", "")
            tally_to = to_date.replace("-", "")
            try:
                tally_existing_numbers = tally.get_existing_voucher_numbers(
                    tally_from, tally_to, config["tally_type"]
                )
                if tally_existing_numbers:
                    log.info(f"  Tally already has {len(tally_existing_numbers)} "
                             f"'{config['tally_type']}' voucher(s) in this date range — "
                             f"will skip those")
            except Exception as e:
                log.warning(f"  Could not fetch existing Tally vouchers for pre-check: {e}")

            # Cross-type duplicate guard: also check related Tally voucher types
            # (e.g. for gst_invoice, also check "Sales" — prevents posting the same
            # invoice under both types if a previous run used the wrong type)
            for extra_type in config.get("also_check_tally_types", []):
                try:
                    extra_numbers = tally.get_existing_voucher_numbers(
                        tally_from, tally_to, extra_type
                    )
                    if extra_numbers:
                        tally_existing_numbers |= extra_numbers
                        log.info(f"  Cross-type check '{extra_type}': "
                                 f"{len(extra_numbers)} voucher(s) also excluded")
                except Exception as e:
                    log.warning(f"  Could not fetch '{extra_type}' for cross-type check: {e}")

        # ── Phase 1: Build XMLs ──────────────────────────────────────
        build_fn = getattr(builder, config["builder_method"])
        built_records = []   # [(record, lines_or_None, xml)]

        for record in records:
            record_id = record["id"]
            record_name = record.get("name", f"ID-{record_id}")

            # Skip if already synced (incremental)
            if mode == "incremental" and record_id in synced_ids:
                type_result["skipped"] += 1
                continue

            # Skip zero-value invoices — not needed for GST or MIS reports
            amount = record.get("amount_total", record.get("amount", 0))
            if vtype not in ("payment", "receipt") and amount == 0:
                log.info(f"  [SKIP] {record_name}: zero-value invoice, not posted to Tally")
                type_result["skipped"] += 1
                continue

            # Skip if already exists in Tally (duplicate guard)
            if record_name in tally_existing_numbers:
                log.info(f"  [SKIP] {record_name}: already exists in Tally")
                type_result["skipped"] += 1
                synced_ids.add(record_id)  # mark as known-synced
                continue

            try:
                if vtype in ("payment", "receipt"):
                    xml = build_fn(record, mapping, bank_ledger=bank_ledger)
                    lines = None
                else:
                    lines = odoo.get_invoice_lines(record_id)
                    _enrich_lines_with_product_type(odoo, lines)
                    xml = build_fn(record, lines, mapping)
                built_records.append((record, lines, xml))
            except Exception as e:
                log.error(f"  [FAIL] Failed to build XML for {record_name}: {e}")
                type_result["errors"] += 1
                type_result["details"].append({
                    "id": record_id, "name": record_name,
                    "status": "error", "message": f"XML build error: {e}",
                })

        # ── Dry-run: preview only, no posting ────────────────────────
        if dry_run:
            for record, _lines, _xml in built_records:
                record_name = record.get("name", f"ID-{record['id']}")
                partner = record["partner_id"][1] if record.get("partner_id") else "N/A"
                amount = record.get("amount_total", record.get("amount", 0))
                inv_date = record.get("invoice_date", record.get("date", ""))
                log.info(f"  [DRY] {record_name} | {partner} | {inv_date} | Rs.{amount:,.2f}")
                type_result["synced"] += 1
                type_result["details"].append({
                    "id": record["id"], "name": record_name,
                    "status": "dry_run", "partner": partner, "amount": amount,
                })
            continue  # next vtype

        # ── Verify: show ledger breakdown and ask confirmation ───────
        if verify and built_records:
            _print_verify_table(built_records, config["description"])
            answer = input(f"  Post {len(built_records)} voucher(s) to Tally? [y/N]: ").strip().lower()
            if answer != "y":
                log.info(f"  Skipped by user (verify step).")
                type_result["skipped"] += len(built_records)
                continue  # next vtype

        # ── Phase 2: Post to Tally ───────────────────────────────────
        for record, lines, xml in built_records:
            record_id = record["id"]
            record_name = record.get("name", f"ID-{record_id}")

            # Ensure party ledger exists in Tally
            partner_name = record["partner_id"][1] if record.get("partner_id") else None
            if partner_name:
                tally_party = TallyXMLBuilder.resolve_partner(partner_name, mapping)
                group = mapping.get("default_group_for_new_creditors", "Sundry Creditors")
                if vtype in ("sales", "sales_return", "receipt", "gst_invoice"):
                    group = mapping.get("default_group_for_new_debtors", "Sundry Debtors")
                # Fetch full partner details from Odoo so address/GST/contact
                # are written to the Tally ledger on creation
                partner_details = {}
                if tally_party.lower() not in tally_ledger_cache:
                    partner_id_val = record["partner_id"][0] if isinstance(
                        record.get("partner_id"), (list, tuple)) else None
                    if partner_id_val:
                        try:
                            partner_details = odoo.get_partner_details(partner_id_val)
                        except Exception as exc:
                            log.warning(f"Could not fetch partner details for '{partner_name}': {exc}")
                ensure_ledger_in_tally(tally, tally_party, group, mapping,
                                       tally_ledger_cache, partner_details=partner_details)

            # Ensure account/tax ledgers exist in Tally
            if lines is not None:
                for line in lines:
                    acct = line.get("account_id")
                    if not acct:
                        continue
                    acct_type = line.get("account_type", "")
                    if acct_type in ("liability_payable", "asset_receivable"):
                        continue
                    acct_name = acct[1] if isinstance(acct, (list, tuple)) else str(acct)
                    # Use resolve_ledger so numeric-prefixed accounts (e.g. "210501.4 Staff Welfare")
                    # resolve to the bare name matching the existing Tally ledger.
                    tally_name = TallyXMLBuilder.resolve_ledger(acct_name, mapping)
                    if line.get("tax_line_id") or acct_type in ("asset_current", "liability_current"):
                        acct_group = "Duties & Taxes"
                    elif acct_type and acct_type.startswith("income"):
                        acct_group = "Sales Accounts"
                    else:
                        acct_group = "Purchase Accounts"
                    ensure_ledger_in_tally(tally, tally_name, acct_group, mapping, tally_ledger_cache)

            # Push to Tally
            try:
                result = tally.import_voucher_xml(xml)
                total_problems = result["errors"] + result.get("exceptions", 0)
                if total_problems > 0 or result["created"] == 0:
                    err_msg = result['error_message'] or f"created={result['created']}, errors={result['errors']}, exceptions={result.get('exceptions',0)}"
                    log.error(f"  [FAIL] {record_name}: {err_msg}")
                    type_result["errors"] += 1
                    type_result["details"].append({
                        "id": record_id, "name": record_name,
                        "status": "error", "message": err_msg,
                    })
                else:
                    log.info(f"  [OK] {record_name} synced to Tally (created={result['created']})")
                    type_result["synced"] += 1
                    synced_ids.add(record_id)
                    type_result["details"].append({
                        "id": record_id, "name": record_name, "status": "synced",
                    })
            except Exception as e:
                log.error(f"  [FAIL] {record_name}: Tally push failed: {e}")
                type_result["errors"] += 1
                type_result["details"].append({
                    "id": record_id, "name": record_name,
                    "status": "error", "message": str(e),
                })

        # Update sync log
        if not dry_run:
            sync_log.setdefault("synced_ids", {})[vtype] = list(synced_ids)

        total_synced += type_result["synced"]
        total_skipped += type_result["skipped"]
        total_errors += type_result["errors"]

        batch_results["types"][vtype] = type_result

        log.info(f"  Summary: {type_result['fetched']} fetched, "
                 f"{type_result['synced']} synced, "
                 f"{type_result['skipped']} skipped, "
                 f"{type_result['errors']} errors")

    # Finalize
    batch_results["completed_at"] = datetime.now().isoformat()
    batch_results["totals"] = {
        "synced": total_synced,
        "skipped": total_skipped,
        "errors": total_errors,
    }

    if not dry_run:
        sync_log["last_sync"] = datetime.now().isoformat()
        sync_log.setdefault("batches", []).append({
            "batch_id": batch_id,
            "timestamp": datetime.now().isoformat(),
            "mode": mode,
            "from_date": from_date,
            "to_date": to_date,
            "synced": total_synced,
            "errors": total_errors,
        })
        save_sync_log(sync_log)

    # Save batch details
    batch_file = TMP_DIR / f"{batch_id}.json"
    with open(batch_file, "w", encoding="utf-8") as f:
        json.dump(batch_results, f, indent=2, default=str)

    log.info(f"\n{'='*60}")
    log.info(f"{'DRY RUN ' if dry_run else ''}SYNC COMPLETE — Batch {batch_id}")
    log.info(f"  Total synced:  {total_synced}")
    log.info(f"  Total skipped: {total_skipped}")
    log.info(f"  Total errors:  {total_errors}")
    log.info(f"  Batch details: {batch_file}")

    return batch_results


# -- CLI Entry Point -------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Odoo → Tally Prime Sync Pipeline",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python odoo_tally_sync.py --dry-run
  python odoo_tally_sync.py --types "purchase,sales"
  python odoo_tally_sync.py --mode full --from 2026-02-01 --to 2026-02-28
  python odoo_tally_sync.py --generate-mapping
        """
    )
    parser.add_argument("--dry-run", action="store_true",
                        help="Preview without pushing to Tally")
    parser.add_argument("--mode", choices=["incremental", "full"],
                        default="incremental",
                        help="Sync mode (default: incremental)")
    parser.add_argument("--types", type=str, default=None,
                        help="Comma-separated voucher types: purchase,purchase_return,"
                             "sales,sales_return,payment,receipt (default: all)")
    parser.add_argument("--from", dest="from_date", type=str, default=None,
                        help="Start date YYYY-MM-DD")
    parser.add_argument("--to", dest="to_date", type=str, default=None,
                        help="End date YYYY-MM-DD")
    parser.add_argument("--generate-mapping", action="store_true",
                        help="Generate/update ledger mapping file")
    parser.add_argument("--rollback", action="store_true",
                        help="Rollback a sync batch (not yet implemented)")
    parser.add_argument("--batch-id", type=str, default=None,
                        help="Batch ID to rollback")
    parser.add_argument("--batch-size", type=int, default=100,
                        help="Batch size for processing (default: 100)")
    parser.add_argument("--voucher-from", type=str, default=None,
                        help="Start voucher number (e.g. BILL/2026/02/0044)")
    parser.add_argument("--voucher-to", type=str, default=None,
                        help="End voucher number (e.g. BILL/2026/02/0071)")
    parser.add_argument("--verify", action="store_true",
                        help="Show ledger breakdown for each voucher and confirm before posting")

    args = parser.parse_args()
    TMP_DIR.mkdir(parents=True, exist_ok=True)

    # Initialize connectors
    try:
        odoo = OdooConnector()
        log.info(f"[OK] Odoo connected: {odoo.url} (DB: {odoo.db})")
    except (ValueError, ConnectionError) as e:
        log.error(f"[FAIL] Odoo connection failed: {e}")
        sys.exit(1)

    company_name = os.getenv("TALLY_COMPANY_NAME", "")
    if not company_name:
        log.error("[FAIL] TALLY_COMPANY_NAME not set in .env")
        sys.exit(1)

    tally = TallyConnector()

    if not args.dry_run and not args.generate_mapping:
        try:
            info = tally.test_connection()
            log.info(f"[OK] Tally connected: {info['url']} (Company: {company_name})")
        except ConnectionError as e:
            log.error(f"[FAIL] Tally connection failed: {e}")
            sys.exit(1)

    builder = TallyXMLBuilder(company_name=company_name, vch_prefix="")

    # Handle actions
    if args.generate_mapping:
        generate_mapping(odoo, tally)
        return

    if args.rollback:
        log.warning("Rollback feature is planned but not yet implemented.")
        log.info("To manually rollback: open Tally → Day Book → delete the vouchers "
                 f"from batch {args.batch_id or 'last batch'}")
        return

    # Determine voucher types
    if args.types:
        voucher_types = [t.strip().lower() for t in args.types.split(",")]
    else:
        voucher_types = list(VOUCHER_TYPE_MAP.keys())

    # Run sync
    sync_vouchers(
        odoo=odoo,
        tally=tally,
        builder=builder,
        voucher_types=voucher_types,
        from_date=args.from_date,
        to_date=args.to_date,
        mode=args.mode,
        dry_run=args.dry_run,
        verify=args.verify,
        batch_size=args.batch_size,
        voucher_from=args.voucher_from,
        voucher_to=args.voucher_to,
    )


if __name__ == "__main__":
    main()
