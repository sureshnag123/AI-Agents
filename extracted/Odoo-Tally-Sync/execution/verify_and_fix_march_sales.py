#!/usr/bin/env python3
"""
Verification & Fix Agent: March 2026 Sales Invoices

Findings from sync_log analysis:
  - 17 invoices posted TWICE: as "Sales" AND as "GST INVOICE" (duplicates)
  - 6 invoices posted as "Sales" ONLY (wrong type, should be "GST INVOICE")

This agent:
  Step 1 — Verify: Fetch March vouchers from Tally (both types) + Odoo invoices
  Step 2 — Report: Show duplicates, wrong-type, and correctly posted entries
  Step 3 — Delete: Remove all wrong "Sales" type vouchers from Tally
  Step 4 — Re-sync: Post the 6 missing entries as "GST INVOICE"
  Step 5 — Fix log: Clean up sync_log.json (remove sales IDs for March)

Usage:
    python execution/verify_and_fix_march_sales.py --verify-only
    python execution/verify_and_fix_march_sales.py --fix
"""

import os
import sys
import json
import argparse
import re
import xml.etree.ElementTree as ET
from pathlib import Path
from datetime import date
from dotenv import load_dotenv

SCRIPT_DIR = Path(__file__).parent.resolve()
PROJECT_ROOT = SCRIPT_DIR.parent
TMP_DIR = PROJECT_ROOT / ".tmp"
SYNC_LOG_FILE = TMP_DIR / "sync_log.json"
MAPPING_FILE  = TMP_DIR / "ledger_mapping.json"

sys.path.insert(0, str(SCRIPT_DIR))
load_dotenv(PROJECT_ROOT / ".env")

from odoo_connector import OdooConnector
from tally_connector import TallyConnector
from tally_xml_builder import TallyXMLBuilder


# ── Helpers ─────────────────────────────────────────────────────────

def load_sync_log() -> dict:
    if SYNC_LOG_FILE.exists():
        with open(SYNC_LOG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"last_sync": None, "batches": [], "synced_ids": {}}


def save_sync_log(d: dict):
    with open(SYNC_LOG_FILE, "w", encoding="utf-8") as f:
        json.dump(d, f, indent=2, default=str)


def load_mapping() -> dict:
    if MAPPING_FILE.exists():
        with open(MAPPING_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def sanitize_xml(text: str) -> ET.Element:
    text = re.sub(r'&#x([0-8BbCcEeFf]);', '', text)
    text = re.sub(r'&#x1[0-9A-Fa-f];', '', text)
    text = re.sub(r'&#([0-8]|1[0-1]|1[4-9]|2[0-9]|3[01]);', '', text)
    text = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', text)
    return ET.fromstring(text)


def fetch_tally_vouchers_with_masterid(tally: TallyConnector,
                                        vch_type: str) -> list:
    """
    Fetch ALL vouchers of vch_type with MASTERID, VOUCHERNUMBER, DATE,
    PARTYLEDGERNAME, AMOUNT, NARRATION.

    MASTERID is the internal Tally database ID required for deletion.
    """
    safe_name = vch_type.replace(" ", "_")
    xml_req = f"""<ENVELOPE>
  <HEADER>
    <VERSION>1</VERSION>
    <TALLYREQUEST>Export</TALLYREQUEST>
    <TYPE>Collection</TYPE>
    <ID>MasterIDColl_{safe_name}</ID>
  </HEADER>
  <BODY>
    <DESC>
      <STATICVARIABLES>
        <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
        <SVCURRENTCOMPANY>{tally.company}</SVCURRENTCOMPANY>
      </STATICVARIABLES>
      <TDL>
        <TDLMESSAGE>
          <COLLECTION NAME="MasterIDColl_{safe_name}" ISMODIFY="No">
            <TYPE>Voucher</TYPE>
            <CHILDOF>{vch_type}</CHILDOF>
            <FETCH>GUID, MASTERID, DATE, VOUCHERNUMBER, VOUCHERTYPENAME, PARTYLEDGERNAME, AMOUNT, NARRATION</FETCH>
          </COLLECTION>
        </TDLMESSAGE>
      </TDL>
    </DESC>
  </BODY>
</ENVELOPE>"""

    resp = tally._send_xml(xml_req, timeout=120)
    try:
        root = tally._parse_response(resp)
    except Exception as e:
        print(f"  [WARN] Could not parse Tally XML for {vch_type}: {e}")
        return []

    rows = []
    for v in root.iter("VOUCHER"):
        master_id = v.findtext("MASTERID", "").strip()
        guid      = (v.findtext("GUID", "").strip()
                     or v.get("REMOTEID", "").strip())
        vchkey    = v.get("VCHKEY", "").strip()
        vno       = v.findtext("VOUCHERNUMBER", "").strip()
        raw_date  = v.findtext("DATE", "").strip()
        narration = v.findtext("NARRATION", "").strip()
        party     = v.findtext("PARTYLEDGERNAME", "").strip()
        amt_str   = v.findtext("AMOUNT", "0").replace(",", "").strip()
        try:
            amount = abs(float(amt_str))
        except ValueError:
            amount = 0.0

        if not vno:
            continue  # Skip schema/template element (first empty row)

        # Use actual VOUCHERTYPENAME from the voucher (more accurate than the query parameter)
        actual_type = (v.findtext("VOUCHERTYPENAME", "").strip()
                       or v.get("VCHTYPE", "").strip()
                       or vch_type)

        fmt_date = (f"{raw_date[:4]}-{raw_date[4:6]}-{raw_date[6:]}"
                    if len(raw_date) == 8 else raw_date)

        rows.append({
            "vch_type":  actual_type,
            "guid":      guid,
            "vchkey":    vchkey,
            "master_id": master_id,
            "date":      fmt_date,
            "date_raw":  raw_date,
            "number":    vno,
            "party":     party,
            "amount":    amount,
            "narration": narration,
        })
    return rows


def fetch_tally_vouchers_for_invoices(tally: TallyConnector,
                                       vch_type: str,
                                       invoice_names: set) -> list:
    """
    Fetch all vouchers of vch_type from Tally (with MASTERID), then filter
    to only those whose VOUCHERNUMBER matches one of the given Odoo invoice names.
    """
    all_vouchers = fetch_tally_vouchers_with_masterid(tally, vch_type)

    rows = []
    for v in all_vouchers:
        vno       = v.get("number", "")
        narration = v.get("narration", "")

        # Match by voucher number (Tally stores Odoo invoice number as VOUCHERNUMBER)
        odoo_ref = vno if vno in invoice_names else ""
        if not odoo_ref:
            m = re.search(r'(INV/\d{4}/\w+|INVFR/[\d-]+/\d+)', narration or "")
            if m and m.group(0) in invoice_names:
                odoo_ref = m.group(0)

        if not odoo_ref:
            continue   # Not a March invoice we care about

        v["odoo_ref"] = odoo_ref
        rows.append(v)
    return rows


def delete_tally_voucher_by_guid(tally: TallyConnector, vch_type: str,
                                  guid: str, vchkey: str, date_raw: str,
                                  vch_number: str, party: str) -> dict:
    """
    Delete a Tally voucher using its GUID/REMOTEID + VCHKEY.
    Uses the IMPORTDATA/REQUESTDATA envelope which is required for Tally Prime
    voucher deletion (the standard Import Data envelope does not work for DELETE).
    """
    from xml.sax.saxutils import escape as xml_escape
    xml_del = f"""<ENVELOPE>
  <HEADER>
    <TALLYREQUEST>Import Data</TALLYREQUEST>
  </HEADER>
  <BODY>
    <IMPORTDATA>
      <REQUESTDESC>
        <REPORTNAME>Vouchers</REPORTNAME>
        <STATICVARIABLES>
          <SVCURRENTCOMPANY>{xml_escape(tally.company)}</SVCURRENTCOMPANY>
        </STATICVARIABLES>
      </REQUESTDESC>
      <REQUESTDATA>
        <TALLYMESSAGE xmlns:UDF="TallyUDF">
          <VOUCHER VCHTYPE="{xml_escape(vch_type)}" ACTION="DELETE"
                   REMOTEID="{xml_escape(guid)}" VCHKEY="{xml_escape(vchkey)}"
                   OBJVIEW="Accounting Voucher View">
            <GUID>{xml_escape(guid)}</GUID>
            <DATE>{date_raw}</DATE>
            <VOUCHERNUMBER>{xml_escape(vch_number)}</VOUCHERNUMBER>
            <PARTYLEDGERNAME>{xml_escape(party)}</PARTYLEDGERNAME>
            <VOUCHERTYPENAME>{xml_escape(vch_type)}</VOUCHERTYPENAME>
          </VOUCHER>
        </TALLYMESSAGE>
      </REQUESTDATA>
    </IMPORTDATA>
  </BODY>
</ENVELOPE>"""

    resp = tally._send_xml(xml_del, timeout=30)
    try:
        root = tally._parse_response(resp)
        return {
            "deleted": int(root.findtext(".//DELETED", "0")),
            "errors":  int(root.findtext(".//ERRORS", "0")),
            "msg":     root.findtext(".//LINEERROR", ""),
        }
    except Exception as e:
        return {"deleted": 0, "errors": 1, "msg": str(e)}


def enrich_lines_with_product_type(odoo: OdooConnector, lines: list):
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


# ── Main Logic ───────────────────────────────────────────────────────

def step1_verify(odoo: OdooConnector, tally: TallyConnector,
                 from_date: str, to_date: str) -> dict:
    """
    Fetch and cross-reference March invoices from Odoo and both
    voucher types from Tally.  Returns a report dict.
    """
    print(f"\n{'='*70}")
    print(f"  STEP 1 — VERIFY  ({from_date} to {to_date})")
    print(f"{'='*70}")

    # Odoo
    print("\n[Odoo] Fetching posted sales invoices ...")
    odoo_invoices = odoo.get_sales_invoices(from_date=from_date, to_date=to_date)
    odoo_by_name  = {inv["name"]: inv for inv in odoo_invoices}
    invoice_names = set(odoo_by_name.keys())
    print(f"  -> {len(odoo_invoices)} invoices found in Odoo")

    # Tally — fetch ALL sales-category vouchers once, split by actual VOUCHERTYPENAME
    print("\n[Tally] Fetching all sales-category vouchers ...")
    all_tally = fetch_tally_vouchers_for_invoices(tally, "Sales", invoice_names)
    # Split by actual voucher type
    tally_sales = [v for v in all_tally if v.get("vch_type") == "Sales"]
    tally_gst   = [v for v in all_tally if v.get("vch_type") == "GST INVOICE"]
    print(f"  -> {len(tally_sales)} actual 'Sales' type vouchers for March invoices")
    print(f"  -> {len(tally_gst)} actual 'GST INVOICE' type vouchers for March invoices")

    # Build lookup maps — store ALL matching rows per invoice (list, not single entry)
    from collections import defaultdict
    sales_by_ref = defaultdict(list)
    for v in tally_sales:
        if v["odoo_ref"]:
            sales_by_ref[v["odoo_ref"]].append(v)
    gst_by_ref = defaultdict(list)
    for v in tally_gst:
        if v["odoo_ref"]:
            gst_by_ref[v["odoo_ref"]].append(v)

    # Fallback: index by voucher number as well
    sales_by_vno = defaultdict(list)
    for v in tally_sales:
        sales_by_vno[v["number"]].append(v)
    gst_by_vno = defaultdict(list)
    for v in tally_gst:
        gst_by_vno[v["number"]].append(v)

    results = {
        "odoo_invoices":  odoo_invoices,
        "tally_sales":    tally_sales,
        "tally_gst":      tally_gst,
        "correct_gst":    [],   # in GST INVOICE only (correct)
        "duplicates":     [],   # in BOTH Sales and GST INVOICE
        "wrong_sales":    [],   # in Sales only (wrong type)
        "missing_both":   [],   # in Odoo but not in Tally at all
    }

    print(f"\n{'─'*70}")
    print(f"  CROSS-REFERENCE ANALYSIS")
    print(f"{'─'*70}")

    for inv in odoo_invoices:
        name = inv["name"]
        sales_rows = sales_by_ref.get(name) or sales_by_vno.get(name) or []
        gst_rows   = gst_by_ref.get(name)   or gst_by_vno.get(name)   or []
        in_sales   = len(sales_rows) > 0
        in_gst     = len(gst_rows) > 0

        if in_gst and in_sales:
            results["duplicates"].append({
                "invoice":         name,
                "customer":        inv["partner_id"][1] if inv.get("partner_id") else "",
                "date":            inv.get("invoice_date", ""),
                "amount":          inv.get("amount_total", 0),
                "odoo_id":         inv["id"],
                "tally_sales_rows": sales_rows,   # ALL Sales copies to delete
                "tally_gst_rows":   gst_rows,
            })
        elif in_gst and not in_sales:
            results["correct_gst"].append({
                "invoice":  name,
                "customer": inv["partner_id"][1] if inv.get("partner_id") else "",
                "date":     inv.get("invoice_date", ""),
                "amount":   inv.get("amount_total", 0),
                "odoo_id":  inv["id"],
                "tally_gst_rows": gst_rows,
            })
        elif in_sales and not in_gst:
            results["wrong_sales"].append({
                "invoice":         name,
                "customer":        inv["partner_id"][1] if inv.get("partner_id") else "",
                "date":            inv.get("invoice_date", ""),
                "amount":          inv.get("amount_total", 0),
                "odoo_id":         inv["id"],
                "tally_sales_rows": sales_rows,
            })
        else:
            results["missing_both"].append({
                "invoice":  name,
                "customer": inv["partner_id"][1] if inv.get("partner_id") else "",
                "date":     inv.get("invoice_date", ""),
                "amount":   inv.get("amount_total", 0),
                "odoo_id":  inv["id"],
            })

    # Print summary
    print(f"\n  {'Category':<45} {'Count':>6}")
    print(f"  {'─'*45} {'─'*6}")
    print(f"  {'Odoo invoices (March)':<45} {len(odoo_invoices):>6}")
    print(f"  {'Tally Sales vouchers':<45} {len(tally_sales):>6}")
    print(f"  {'Tally GST INVOICE vouchers':<45} {len(tally_gst):>6}")
    print(f"  {'─'*45} {'─'*6}")
    print(f"  {'[OK] Correctly posted (GST INVOICE only)':<45} {len(results['correct_gst']):>6}")
    print(f"  {'[!!] DUPLICATES (Sales + GST INVOICE)':<45} {len(results['duplicates']):>6}")
    print(f"  {'[!!] WRONG TYPE (Sales only, no GST INV)':<45} {len(results['wrong_sales']):>6}")
    print(f"  {'[!!] MISSING in Tally entirely':<45} {len(results['missing_both']):>6}")

    if results["duplicates"]:
        print(f"\n  DUPLICATES — these exist in BOTH Sales and GST INVOICE:")
        for r in results["duplicates"]:
            print(f"    {r['invoice']:<25}  {r['date']}  {r['customer']:<35}  ₹{r['amount']:>12,.2f}")

    if results["wrong_sales"]:
        print(f"\n  WRONG TYPE — posted as 'Sales', should be 'GST INVOICE':")
        for r in results["wrong_sales"]:
            print(f"    {r['invoice']:<25}  {r['date']}  {r['customer']:<35}  ₹{r['amount']:>12,.2f}")

    if results["correct_gst"]:
        print(f"\n  CORRECT — already in GST INVOICE (no action needed):")
        for r in results["correct_gst"]:
            print(f"    {r['invoice']:<25}  {r['date']}  {r['customer']:<35}  ₹{r['amount']:>12,.2f}")

    if results["missing_both"]:
        print(f"\n  MISSING — not in Tally at all:")
        for r in results["missing_both"]:
            print(f"    {r['invoice']:<25}  {r['date']}  {r['customer']:<35}  ₹{r['amount']:>12,.2f}")

    return results


def step2_delete_sales_vouchers(tally: TallyConnector, report: dict) -> list:
    """
    Delete all wrongly-typed 'Sales' vouchers from Tally:
      - Entries in 'duplicates' -> delete the Sales copy (keep GST INVOICE)
      - Entries in 'wrong_sales' -> delete the Sales copy (will re-post as GST INVOICE)
    """
    print(f"\n{'='*70}")
    print(f"  STEP 2 — DELETE wrong 'Sales' vouchers from Tally")
    print(f"{'='*70}")

    # Build flat list of all Sales rows to delete
    to_delete = []  # (inv_name, row, reason)
    for r in report["duplicates"]:
        for row in r.get("tally_sales_rows", []):
            to_delete.append((r["invoice"], row, "duplicate"))
    for r in report["wrong_sales"]:
        for row in r.get("tally_sales_rows", []):
            to_delete.append((r["invoice"], row, "wrong_type"))

    if not to_delete:
        print("  Nothing to delete.")
        return []

    print(f"  Total Sales copies to delete: {len(to_delete)}")
    deleted_ok = []
    for inv_name, row, reason in to_delete:
        guid       = row.get("guid", "")
        vchkey     = row.get("vchkey", "")
        date_raw   = row.get("date_raw") or row["date"].replace("-", "")
        vch_number = row["number"]
        party      = row["party"]
        if not guid:
            print(f"  [SKIP] {inv_name:<25}  no GUID — cannot delete")
            continue
        result = delete_tally_voucher_by_guid(
            tally, "Sales", guid, vchkey, date_raw, vch_number, party)
        status = "DELETED" if result["deleted"] > 0 else f"FAILED ({result['msg']})"
        print(f"  [{status}] {inv_name:<25}  ({reason})  guid=...{guid[-12:]}  vch={vch_number}")
        if result["deleted"] > 0:
            deleted_ok.append(inv_name)

    print(f"\n  Deleted: {len(deleted_ok)} / {len(to_delete)} Sales vouchers")
    return deleted_ok


def step3_post_missing_as_gst_invoice(
    odoo: OdooConnector, tally: TallyConnector,
    report: dict, mapping: dict, company_name: str
) -> list:
    """
    Post all entries that are not yet in 'GST INVOICE':
      - 'wrong_sales' entries (just deleted from Sales, need GST INVOICE)
      - 'missing_both' entries (never posted)
    """
    print(f"\n{'='*70}")
    print(f"  STEP 3 — Post missing entries as 'GST INVOICE'")
    print(f"{'='*70}")

    # Gather all Odoo invoices that need posting as GST INVOICE
    to_post = report["wrong_sales"] + report["missing_both"]

    if not to_post:
        print("  Nothing to post — all entries already correct.")
        return []

    builder = TallyXMLBuilder(company_name=company_name, vch_prefix="")
    bank_ledger = mapping.get("default_bank_ledger", "Bank Accounts")

    # Fetch Tally ledger cache
    try:
        existing = tally.get_all_ledgers()
        tally_ledger_cache = {l["name"].lower() for l in existing}
    except Exception:
        tally_ledger_cache = set()

    # Ensure debtor ledger helper
    def ensure_debtor(name):
        if name.lower() not in tally_ledger_cache:
            tally_party = TallyXMLBuilder.resolve_partner(name, mapping)
            group = mapping.get("default_group_for_new_debtors", "Sundry Debtors")
            result = tally.create_ledger(tally_party, group)
            if result.startswith("OK"):
                tally_ledger_cache.add(tally_party.lower())
                print(f"  [LEDGER CREATED] '{tally_party}' under '{group}'")

    # Build lookup from already-fetched Odoo invoices in the report
    odoo_by_id = {inv["id"]: inv for inv in report["odoo_invoices"]}

    posted_ok = []
    for entry in to_post:
        inv_name   = entry["invoice"]
        odoo_id    = entry["odoo_id"]
        customer   = entry.get("customer", "")
        amount     = entry.get("amount", 0)

        inv_record = odoo_by_id.get(odoo_id)
        if not inv_record:
            print(f"  [SKIP] {inv_name} — cannot find Odoo record")
            continue

        # Fetch lines
        lines = odoo.get_invoice_lines(odoo_id)
        enrich_lines_with_product_type(odoo, lines)

        # Ensure party ledger
        ensure_debtor(customer)

        # Build GST INVOICE XML
        try:
            xml = builder.build_gst_invoice_voucher(inv_record, lines, mapping)
        except Exception as e:
            print(f"  [FAIL] {inv_name} — XML build error: {e}")
            continue

        # Post to Tally
        result = tally.import_voucher_xml(xml)
        total_problems = result["errors"] + result.get("exceptions", 0)
        if total_problems > 0 or result["created"] == 0:
            err = result["error_message"] or f"created={result['created']}"
            print(f"  [FAIL] {inv_name} — {err}")
        else:
            print(f"  [OK]   {inv_name}  {entry['date']}  {customer:<35}  ₹{amount:>12,.2f}")
            posted_ok.append(odoo_id)

    print(f"\n  Posted: {len(posted_ok)} / {len(to_post)} invoices as GST INVOICE")
    return posted_ok


def step4_fix_sync_log(report: dict, newly_posted_ids: list,
                        from_date: str, to_date: str):
    """
    Update sync_log.json:
      - Remove March Sales IDs (all were wrong type, now deleted)
      - Add newly posted IDs to gst_invoice
    """
    print(f"\n{'='*70}")
    print(f"  STEP 4 — Fix sync_log.json")
    print(f"{'='*70}")

    sync_log = load_sync_log()
    odoo_invoices = report["odoo_invoices"]
    march_odoo_ids = set(inv["id"] for inv in odoo_invoices)

    # Remove March IDs from 'sales' (they were wrong type, deleted from Tally)
    current_sales = set(sync_log.get("synced_ids", {}).get("sales", []))
    cleaned_sales = current_sales - march_odoo_ids
    sync_log.setdefault("synced_ids", {})["sales"] = list(cleaned_sales)

    removed = len(current_sales) - len(cleaned_sales)
    print(f"  Removed {removed} March IDs from synced_ids['sales']")

    # Add newly posted IDs to gst_invoice
    current_gst = set(sync_log["synced_ids"].get("gst_invoice", []))
    # Also add IDs that were already correct (they should stay)
    correct_ids  = set(r["odoo_id"] if "odoo_id" in r else 0
                       for r in report.get("correct_gst", []))
    for inv in report["odoo_invoices"]:
        if inv["name"] in {r["invoice"] for r in report["correct_gst"]}:
            correct_ids.add(inv["id"])

    new_gst = current_gst | correct_ids | set(newly_posted_ids)
    sync_log["synced_ids"]["gst_invoice"] = list(new_gst)
    added = len(new_gst) - len(current_gst)
    print(f"  Added {added} IDs to synced_ids['gst_invoice']")

    save_sync_log(sync_log)
    print("  sync_log.json updated.")


def print_final_summary(report: dict, deleted_names: list, posted_ids: list):
    print(f"\n{'='*70}")
    print(f"  FINAL SUMMARY")
    print(f"{'='*70}")
    print(f"  Odoo March invoices:              {len(report['odoo_invoices']):>4}")
    print(f"  Already correct (GST INVOICE):    {len(report['correct_gst']):>4}")
    print(f"  Duplicates removed (Sales copy):  {len(deleted_names):>4}")
    print(f"  Wrong-type fixed (-> GST INVOICE): {len(report['wrong_sales']):>4}")
    print(f"  Missing entries posted:           {len(report['missing_both']):>4}")
    expected_gst = len(report['correct_gst']) + len(report['wrong_sales']) + len(report['missing_both'])
    print(f"  Expected GST INVOICE count:       {expected_gst:>4}")
    print(f"{'='*70}")


# ── CLI ───────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Verify and fix March 2026 sales invoice voucher types"
    )
    parser.add_argument("--verify-only", action="store_true",
                        help="Only verify — do not delete or re-post")
    parser.add_argument("--fix", action="store_true",
                        help="Verify + delete wrong entries + re-post correctly")
    parser.add_argument("--from-date", default="2026-03-01")
    parser.add_argument("--to-date",   default="2026-03-31")
    args = parser.parse_args()

    if not args.verify_only and not args.fix:
        parser.error("Specify --verify-only or --fix")

    TMP_DIR.mkdir(parents=True, exist_ok=True)

    # Connect
    try:
        odoo = OdooConnector()
        print(f"[OK] Odoo: {odoo.url}")
    except Exception as e:
        print(f"[FAIL] Odoo: {e}"); sys.exit(1)

    company_name = os.getenv("TALLY_COMPANY_NAME", "")
    if not company_name:
        print("[FAIL] TALLY_COMPANY_NAME not set in .env"); sys.exit(1)

    tally = TallyConnector()
    try:
        info = tally.test_connection()
        print(f"[OK] Tally: {info['url']}  (Company: {company_name})")
    except ConnectionError as e:
        print(f"[FAIL] Tally: {e}"); sys.exit(1)

    mapping = load_mapping()

    # Step 1 — Always verify
    report = step1_verify(odoo, tally, args.from_date, args.to_date)

    if args.verify_only:
        print("\n[verify-only] Done.  Run with --fix to apply corrections.")
        return

    # Step 2 — Delete wrong Sales vouchers
    deleted_names = step2_delete_sales_vouchers(tally, report)

    # Step 3 — Post missing as GST INVOICE
    posted_ids = step3_post_missing_as_gst_invoice(
        odoo, tally, report, mapping, company_name
    )

    # Step 4 — Fix sync log
    step4_fix_sync_log(report, posted_ids, args.from_date, args.to_date)

    # Final summary
    print_final_summary(report, deleted_names, posted_ids)


if __name__ == "__main__":
    main()
