#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Push Odoo Sales Invoices → Tally (FY 2025-26)
==============================================
Fetches posted sales invoices from Odoo and pushes only those NOT yet
accounted in Tally, preventing duplicates via a 3-pass live pre-check:

  Pass 1 — Voucher number match  (Odoo invoice name == Tally VOUCHERNUMBER)
  Pass 2 — Party + amount match  (±₹1 tolerance, handles manual Tally entries)
  Pass 3 — Date + amount match   (±₹1 within same month, final safety net)

Tally voucher type used: "GST INVOICE"  (per FY 2025-26 reconciliation rules)

Usage:
    # Dry run — preview only, nothing posted (default)
    python push_sales_to_tally.py

    # Full FY (Apr 2025 – Mar 2026), dry run
    python push_sales_to_tally.py --from 2025-04-01 --to 2026-03-31

    # Actually post to Tally
    python push_sales_to_tally.py --execute

    # Specific month
    python push_sales_to_tally.py --from 2025-07-01 --to 2025-07-31 --execute

    # Verify each voucher before posting
    python push_sales_to_tally.py --execute --verify
"""

import io
import os
import re
import sys
import json
import logging
import argparse
import xml.etree.ElementTree as ET
from datetime import date, datetime
from pathlib import Path

# ── UTF-8 console on Windows ──────────────────────────────────────────
if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
if sys.stderr.encoding and sys.stderr.encoding.lower() != "utf-8":
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

from dotenv import load_dotenv

SCRIPT_DIR   = Path(__file__).parent.resolve()
PROJECT_ROOT = SCRIPT_DIR.parent
TMP_DIR      = PROJECT_ROOT / ".tmp"
TMP_DIR.mkdir(parents=True, exist_ok=True)

load_dotenv(PROJECT_ROOT / ".env")
sys.path.insert(0, str(SCRIPT_DIR))

from odoo_connector    import OdooConnector
from tally_connector   import TallyConnector
from tally_xml_builder import TallyXMLBuilder

# ── Logging ───────────────────────────────────────────────────────────

LOG_DIR = TMP_DIR / "logs"
LOG_DIR.mkdir(parents=True, exist_ok=True)

_stream_handler = logging.StreamHandler()
_stream_handler.stream = open(
    sys.stdout.fileno(), mode="w", encoding="utf-8", buffering=1, closefd=False
) if hasattr(sys.stdout, "fileno") else sys.stdout

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        _stream_handler,
        logging.FileHandler(
            LOG_DIR / f"push_sales_{date.today().isoformat()}.log",
            encoding="utf-8",
        ),
    ],
)
log = logging.getLogger("push_sales")

MAPPING_FILE  = TMP_DIR / "ledger_mapping.json"
TALLY_VCH_TYPE = "GST INVOICE"
# Also check plain "Sales" vouchers to guard against a previous run that used
# the wrong Tally type
ALSO_CHECK_TYPES = ["Sales"]


# ── Helpers ───────────────────────────────────────────────────────────

def load_mapping() -> dict:
    if MAPPING_FILE.exists():
        with open(MAPPING_FILE, encoding="utf-8") as f:
            return json.load(f)
    return {
        "odoo_to_tally_ledgers": {},
        "odoo_to_tally_partners": {},
        "auto_create_missing_ledgers": True,
        "default_group_for_new_creditors": "Sundry Creditors",
        "default_group_for_new_debtors": "Sundry Debtors",
        "default_bank_ledger": "Bank Accounts",
    }


def _norm_party(name: str) -> str:
    """Normalise party name for fuzzy matching (lower, collapse special chars, drop corp suffixes)."""
    name = name.lower().strip()
    # collapse any sequence of non-alphanumeric chars into a single space
    name = re.sub(r"[^a-z0-9]+", " ", name).strip()
    # strip common Indian corporate suffixes (longest match first)
    for suffix in ("pvt ltd", "private limited", "limited", "pvt", "ltd", "llp", "inc", "corp"):
        if name.endswith(" " + suffix):
            name = name[: -(len(suffix) + 1)].strip()
            break
    return " ".join(name.split())


def _tally_amount(raw: str) -> float:
    """Parse a Tally amount string (may be negative) to float."""
    try:
        return abs(float(str(raw).replace(",", "").strip()))
    except (ValueError, TypeError):
        return 0.0


def _already_in_tally(
    odoo_inv: dict,
    tally_by_number: dict,
    tally_by_party_amount: dict,
    tally_by_date_amount: dict,
) -> tuple[bool, str]:
    """
    3-pass duplicate check.
    Returns (True, reason_string) if the invoice is already in Tally.
    """
    inv_name   = odoo_inv.get("name", "")
    inv_amount = round(float(odoo_inv.get("amount_total", 0)), 2)
    inv_date   = (odoo_inv.get("invoice_date") or "").replace("-", "")
    partner    = odoo_inv.get("partner_id", [None, ""])[1] if odoo_inv.get("partner_id") else ""

    # Pass 1 — exact voucher number
    if inv_name and inv_name in tally_by_number:
        return True, f"voucher number '{inv_name}' already in Tally"

    # Pass 2 — normalised party + amount (±₹1)
    party_key = _norm_party(partner)
    for tally_amt in tally_by_party_amount.get(party_key, []):
        if abs(tally_amt - inv_amount) <= 1.0:
            return True, (f"party '{partner}' + amount ₹{inv_amount:,.2f} "
                          f"matched in Tally (tally amt ₹{tally_amt:,.2f})")

    # Pass 3 — same date + amount (±₹1)
    for tally_amt in tally_by_date_amount.get(inv_date, []):
        if abs(tally_amt - inv_amount) <= 1.0:
            return True, (f"date {inv_date} + amount ₹{inv_amount:,.2f} "
                          f"matched in Tally (tally amt ₹{tally_amt:,.2f})")

    return False, ""


def _build_tally_lookup_maps(tally_vouchers: list) -> tuple[dict, dict, dict]:
    """
    Build three lookup structures for fast duplicate checking:
      by_number       : {voucher_number → True}
      by_party_amount : {norm_party_name → [amount, ...]}
      by_date_amount  : {YYYYMMDD → [amount, ...]}
    """
    by_number       = {}
    by_party_amount = {}
    by_date_amount  = {}

    for v in tally_vouchers:
        num    = (v.get("number") or "").strip()
        party  = _norm_party(v.get("party") or "")
        amt    = _tally_amount(v.get("amount") or "0")
        dt     = (v.get("date") or "").strip()

        if num:
            by_number[num] = True
        if party and amt > 0:
            by_party_amount.setdefault(party, []).append(amt)
        if dt and amt > 0:
            by_date_amount.setdefault(dt, []).append(amt)

    return by_number, by_party_amount, by_date_amount


def _enrich_product_types(odoo: OdooConnector, lines: list):
    """Batch-fetch product types and attach as _product_type to each line."""
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


def ensure_ledger(tally: TallyConnector, name: str, group: str,
                  mapping: dict, cache: set, partner_details: dict = None) -> bool:
    """Create a Tally ledger if it doesn't already exist."""
    if name.lower() in cache:
        return True
    if not mapping.get("auto_create_missing_ledgers", False):
        log.warning(f"Ledger '{name}' missing in Tally (auto-create disabled)")
        return False

    kwargs = {}
    if partner_details:
        p = partner_details
        addr = []
        if p.get("street"):  addr.append(p["street"])
        if p.get("street2"): addr.append(p["street2"])
        city_zip = " - ".join(filter(None, [p.get("city"), p.get("zip")]))
        if city_zip: addr.append(city_zip)
        if addr: kwargs["address_lines"] = addr
        state = p.get("state_id")
        if isinstance(state, (list, tuple)) and len(state) >= 2:
            # Strip country-code suffix Odoo appends, e.g. "Karnataka (IN)" → "Karnataka"
            kwargs["state"] = re.sub(r"\s*\([A-Z]{2}\)$", "", state[1]).strip()
        country = p.get("country_id")
        if isinstance(country, (list, tuple)) and len(country) >= 2:
            kwargs["country"] = country[1]
        if p.get("zip"):                    kwargs["pincode"]               = p["zip"]
        if p.get("phone"):                  kwargs["phone"]                 = p["phone"]
        if p.get("email"):                  kwargs["email"]                 = p["email"]
        if p.get("vat"):                    kwargs["gst_number"]            = p["vat"]
        if p.get("l10n_in_pan"):            kwargs["pan"]                   = p["l10n_in_pan"]
        if p.get("l10n_in_gst_treatment"):  kwargs["gst_registration_type"] = p["l10n_in_gst_treatment"]

    log.info(f"  Creating ledger '{name}' under '{group}'")
    result = tally.create_ledger(name, group, **kwargs)
    if result.startswith("OK") or "already" in result.lower():
        cache.add(name.lower())
        return True
    log.error(f"  Failed to create ledger '{name}': {result}")
    return False


def _extract_ledger_entries(xml_str: str) -> list:
    """Return [(ledger_name, amount)] from voucher XML for --verify display."""
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


# ── Core Push Function ────────────────────────────────────────────────

def push_sales(
    odoo: OdooConnector,
    tally: TallyConnector,
    builder: TallyXMLBuilder,
    from_date: str,
    to_date: str,
    execute: bool = False,
    verify: bool = False,
) -> dict:
    """
    Fetch Odoo sales invoices for the date range, pre-check against Tally,
    and push those not yet accounted.

    Returns a summary dict.
    """
    mapping = load_mapping()

    log.info("=" * 65)
    log.info(f"{'LIVE POST' if execute else 'DRY RUN'} — "
             f"Sales invoices  {from_date}  →  {to_date}")
    log.info(f"Tally voucher type: {TALLY_VCH_TYPE}")
    log.info("=" * 65)

    # ── Step 1: Fetch from Odoo ───────────────────────────────────────
    log.info("Fetching sales invoices from Odoo...")
    odoo_invoices = odoo.get_sales_invoices(from_date=from_date, to_date=to_date)
    # Keep only non-zero invoices
    odoo_invoices = [inv for inv in odoo_invoices
                     if float(inv.get("amount_total", 0)) != 0]
    log.info(f"  {len(odoo_invoices)} posted sales invoice(s) found in Odoo")

    # ── Step 2: Fetch existing vouchers from Tally ────────────────────
    tally_all_vouchers = []
    if execute or True:   # always pre-check even in dry-run for accurate preview
        tally_from = from_date.replace("-", "")
        tally_to   = to_date.replace("-", "")
        log.info(f"Fetching existing Tally '{TALLY_VCH_TYPE}' vouchers...")
        try:
            primary_vouchers = tally.get_vouchers(TALLY_VCH_TYPE, tally_from, tally_to)
            # Python-side date filter (Tally TDL date filter can be unreliable)
            primary_vouchers = [v for v in primary_vouchers
                                if tally_from <= v.get("date", "") <= tally_to]
            tally_all_vouchers.extend(primary_vouchers)
            log.info(f"  {len(primary_vouchers)} '{TALLY_VCH_TYPE}' voucher(s) already in Tally")
        except Exception as exc:
            log.warning(f"  Could not fetch Tally '{TALLY_VCH_TYPE}' vouchers: {exc}")

        for extra_type in ALSO_CHECK_TYPES:
            try:
                extra_vouchers = tally.get_vouchers(extra_type, tally_from, tally_to)
                extra_vouchers = [v for v in extra_vouchers
                                  if tally_from <= v.get("date", "") <= tally_to]
                tally_all_vouchers.extend(extra_vouchers)
                if extra_vouchers:
                    log.info(f"  Cross-type '{extra_type}': "
                             f"{len(extra_vouchers)} voucher(s) also excluded")
            except Exception as exc:
                log.warning(f"  Could not fetch '{extra_type}' vouchers: {exc}")

    by_number, by_party_amount, by_date_amount = _build_tally_lookup_maps(tally_all_vouchers)

    # ── Step 3: Classify each Odoo invoice ───────────────────────────
    to_post  = []   # (invoice, reason_to_post)
    skipped  = []   # (invoice, reason_skipped)

    for inv in odoo_invoices:
        already, reason = _already_in_tally(
            inv, by_number, by_party_amount, by_date_amount
        )
        if already:
            skipped.append((inv, reason))
        else:
            to_post.append(inv)

    log.info(f"\nPre-check summary:")
    log.info(f"  Already in Tally (skip) : {len(skipped)}")
    log.info(f"  To be posted            : {len(to_post)}")

    if skipped:
        log.info("\n  Skipped invoices:")
        for inv, reason in skipped:
            name    = inv.get("name", f"ID-{inv['id']}")
            partner = inv["partner_id"][1] if inv.get("partner_id") else "N/A"
            amt     = inv.get("amount_total", 0)
            log.info(f"    [SKIP] {name:<28} {partner:<35} ₹{amt:>12,.2f}  →  {reason}")

    if not to_post:
        log.info("\nAll invoices are already accounted in Tally. Nothing to post.")
        return _summary(skipped=len(skipped), posted=0, errors=0)

    # ── Step 4: Build XMLs ────────────────────────────────────────────
    log.info(f"\nBuilding XML for {len(to_post)} invoice(s)...")
    built = []   # [(invoice, lines, xml)]
    build_errors = 0

    for inv in to_post:
        inv_name = inv.get("name", f"ID-{inv['id']}")
        try:
            lines = odoo.get_invoice_lines(inv["id"])
            _enrich_product_types(odoo, lines)
            xml = builder.build_gst_invoice_voucher(inv, lines, mapping)
            built.append((inv, lines, xml))
        except Exception as exc:
            log.error(f"  [FAIL] XML build for {inv_name}: {exc}")
            build_errors += 1

    # ── Dry-run preview ───────────────────────────────────────────────
    if not execute:
        log.info("\n[DRY RUN] Invoices that WOULD be posted:")
        total_amount = 0.0
        for inv, _lines, _xml in built:
            name    = inv.get("name", f"ID-{inv['id']}")
            partner = inv["partner_id"][1] if inv.get("partner_id") else "N/A"
            amt     = float(inv.get("amount_total", 0))
            dt      = inv.get("invoice_date", "")
            total_amount += amt
            log.info(f"  {name:<28} {partner:<35} {dt}  ₹{amt:>12,.2f}")
        log.info(f"\n  Total to post : {len(built)} invoice(s)  ₹{total_amount:,.2f}")
        log.info("  Re-run with --execute to actually post.")
        return _summary(skipped=len(skipped), posted=0, errors=build_errors, dry_run=True,
                        would_post=len(built))

    # ── Step 5: Build Tally ledger cache ──────────────────────────────
    tally_ledger_cache = set()
    try:
        existing_ledgers = tally.get_all_ledgers()
        tally_ledger_cache = {l["name"].lower() for l in existing_ledgers}
        log.info(f"Tally ledger cache: {len(tally_ledger_cache)} ledgers loaded")
    except Exception as exc:
        log.warning(f"Could not fetch Tally ledgers for cache: {exc}")

    # ── Verify (optional interactive confirmation) ────────────────────
    if verify and built:
        print(f"\n{'─'*70}")
        print(f"  VERIFY — {len(built)} invoice(s) about to be posted")
        print(f"{'─'*70}")
        for inv, _lines, xml in built:
            name    = inv.get("name", f"ID-{inv['id']}")
            partner = inv["partner_id"][1] if inv.get("partner_id") else "N/A"
            amt     = inv.get("amount_total", 0)
            dt      = inv.get("invoice_date", "")
            print(f"\n  Voucher : {name}")
            print(f"  Party   : {partner}")
            print(f"  Date    : {dt}   Amount: ₹{amt:,.2f}")
            print(f"  Ledgers :")
            for ledger, amount in _extract_ledger_entries(xml):
                dr_cr = "Dr" if amount < 0 else "Cr"
                print(f"            {dr_cr}  {abs(amount):>12,.2f}   {ledger}")
        print(f"\n{'─'*70}")
        answer = input(f"  Post {len(built)} invoice(s) to Tally? [y/N]: ").strip().lower()
        if answer != "y":
            log.info("Posting cancelled by user at verify step.")
            return _summary(skipped=len(skipped), posted=0, errors=build_errors)

    # ── Step 6: Post to Tally ─────────────────────────────────────────
    log.info(f"\nPosting {len(built)} invoice(s) to Tally...")
    posted_count = 0
    post_errors  = 0
    results      = []

    for inv, lines, xml in built:
        inv_id   = inv["id"]
        inv_name = inv.get("name", f"ID-{inv_id}")
        partner  = inv["partner_id"][1] if inv.get("partner_id") else None
        amt      = float(inv.get("amount_total", 0))

        # Ensure party (debtor) ledger exists
        if partner:
            tally_party = TallyXMLBuilder.resolve_partner(partner, mapping)
            group = mapping.get("default_group_for_new_debtors", "Sundry Debtors")
            partner_details = {}
            if tally_party.lower() not in tally_ledger_cache:
                pid = inv["partner_id"][0] if isinstance(inv.get("partner_id"), (list, tuple)) else None
                if pid:
                    try:
                        partner_details = odoo.get_partner_details(pid)
                    except Exception as e:
                        log.warning(f"  Could not fetch partner details for ID {pid}: {e}")
            ensure_ledger(tally, tally_party, group, mapping,
                          tally_ledger_cache, partner_details=partner_details)

        # Ensure account/tax ledgers exist
        for line in lines:
            acct = line.get("account_id")
            if not acct:
                continue
            acct_type = line.get("account_type", "")
            if acct_type in ("liability_payable", "asset_receivable"):
                continue
            acct_name  = acct[1] if isinstance(acct, (list, tuple)) else str(acct)
            tally_name = TallyXMLBuilder.resolve_ledger(acct_name, mapping)
            if line.get("tax_line_id") or acct_type in ("asset_current", "liability_current"):
                acct_group = "Duties & Taxes"
            elif acct_type and acct_type.startswith("income"):
                acct_group = "Sales Accounts"
            else:
                acct_group = "Purchase Accounts"
            ensure_ledger(tally, tally_name, acct_group, mapping, tally_ledger_cache)

        # Post
        try:
            result = tally.import_voucher_xml(xml)
            problems = result["errors"] + result.get("exceptions", 0)
            if problems > 0 or result["created"] == 0:
                err_msg = result["error_message"] or (
                    f"created={result['created']}, errors={result['errors']}, "
                    f"exceptions={result.get('exceptions', 0)}"
                )
                log.error(f"  [FAIL] {inv_name:<28} ₹{amt:>12,.2f}  →  {err_msg}")
                post_errors += 1
                results.append({"name": inv_name, "status": "error", "message": err_msg})
            else:
                log.info(f"  [OK]   {inv_name:<28} ₹{amt:>12,.2f}  →  posted (created={result['created']})")
                posted_count += 1
                results.append({"name": inv_name, "status": "posted", "amount": amt})
        except Exception as exc:
            log.error(f"  [FAIL] {inv_name}: {exc}")
            post_errors += 1
            results.append({"name": inv_name, "status": "error", "message": str(exc)})

    # ── Save run report ───────────────────────────────────────────────
    run_id      = f"push_sales_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    report_file = TMP_DIR / f"{run_id}.json"
    report = {
        "run_id":        run_id,
        "timestamp":     datetime.now().isoformat(),
        "from_date":     from_date,
        "to_date":       to_date,
        "tally_type":    TALLY_VCH_TYPE,
        "odoo_fetched":  len(odoo_invoices),
        "skipped":       len(skipped),
        "posted":        posted_count,
        "errors":        post_errors + build_errors,
        "details":       results,
    }
    with open(report_file, "w", encoding="utf-8") as f:
        json.dump(report, f, indent=2, default=str)

    log.info("\n" + "=" * 65)
    log.info("PUSH COMPLETE")
    log.info(f"  Odoo fetched : {len(odoo_invoices)}")
    log.info(f"  Skipped      : {len(skipped)}  (already in Tally)")
    log.info(f"  Posted       : {posted_count}")
    log.info(f"  Errors       : {post_errors + build_errors}")
    log.info(f"  Report       : {report_file}")
    log.info("=" * 65)

    return _summary(skipped=len(skipped), posted=posted_count,
                    errors=post_errors + build_errors)


def _summary(**kwargs) -> dict:
    return kwargs


# ── CLI ───────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Push Odoo Sales Invoices → Tally (FY 2025-26)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python push_sales_to_tally.py                              # dry run, full FY
  python push_sales_to_tally.py --execute                    # post full FY
  python push_sales_to_tally.py --from 2025-07-01 --to 2025-07-31 --execute
  python push_sales_to_tally.py --execute --verify           # confirm each voucher
        """,
    )
    parser.add_argument("--execute", action="store_true",
                        help="Actually post to Tally (default: dry run only)")
    parser.add_argument("--from", dest="from_date", default="2025-04-01",
                        help="Start date YYYY-MM-DD (default: 2025-04-01)")
    parser.add_argument("--to", dest="to_date", default="2026-03-31",
                        help="End date YYYY-MM-DD (default: 2026-03-31)")
    parser.add_argument("--verify", action="store_true",
                        help="Show ledger breakdown and confirm before posting")
    args = parser.parse_args()

    # ── Connect ───────────────────────────────────────────────────────
    try:
        odoo = OdooConnector()
        log.info(f"[OK] Odoo connected: {odoo.url}  DB: {odoo.db}")
    except Exception as exc:
        log.error(f"[FAIL] Odoo connection: {exc}")
        sys.exit(1)

    company_name = os.getenv("TALLY_COMPANY_NAME", "")
    if not company_name:
        log.error("[FAIL] TALLY_COMPANY_NAME not set in .env")
        sys.exit(1)

    tally = TallyConnector()
    if args.execute:
        try:
            info = tally.test_connection()
            log.info(f"[OK] Tally connected: {info['url']}  Company: {company_name}")
        except Exception as exc:
            log.error(f"[FAIL] Tally connection: {exc}")
            sys.exit(1)

    builder = TallyXMLBuilder(company_name=company_name, vch_prefix="")

    push_sales(
        odoo=odoo,
        tally=tally,
        builder=builder,
        from_date=args.from_date,
        to_date=args.to_date,
        execute=args.execute,
        verify=args.verify,
    )


if __name__ == "__main__":
    main()
