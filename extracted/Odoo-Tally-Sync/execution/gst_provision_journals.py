#!/usr/bin/env python3
"""
GST Provision Journal Generator

Fetches closing balances of GST input/output ledgers from Tally for every
month-end in the Indian FY (April-March) and generates reversing Journal
vouchers to compute net GST Payable or Input Credit Excess.

Reversal logic:
  - Input GST ledgers with Dr balance  -> reversed to CREDIT in Journal
  - Output GST ledgers with Cr balance -> reversed to DEBIT  in Journal
  - Balancing entry                    -> GST Payable
      Cr (positive) = taxes payable to government
      Dr (negative) = input credit excess over output

Only these 10 ledgers are used - no others:
    GST Payable
    INPUT CGST           Input CGST-RCM
    INPUT IGST           INPUT IGST ON IMPORTS
    INPUT SGST           Input SGST-RCM
    Output CGST @ 9%     Output IGST @ 18%     Output SGST @ 9%

Fetch methods (tried in order):
  1. Ledger Vouchers report  -- per-ledger, most reliable
  2. Trial Balance EXPLODEFLAG -- fallback, may miss deep sub-ledgers
  3. --manual                -- paste values from Tally Balance Sheet screen

Usage:
    python execution/gst_provision_journals.py --dry-run
    python execution/gst_provision_journals.py --year 2025 --months APR --dry-run
    python execution/gst_provision_journals.py --save-xml
    python execution/gst_provision_journals.py --import
    python execution/gst_provision_journals.py --manual APR 2025   # enter values interactively
    python execution/gst_provision_journals.py --diagnose           # dump raw Tally response
"""

import sys
import re
import json
import argparse
import calendar
import time
import xml.etree.ElementTree as ET
from datetime import date
from pathlib import Path
from typing import Dict, List, Optional, Tuple
from xml.sax.saxutils import escape as xml_escape

from dotenv import load_dotenv

SCRIPT_DIR = Path(__file__).parent.resolve()
PROJECT_ROOT = SCRIPT_DIR.parent
TMP_DIR = PROJECT_ROOT / ".tmp"

load_dotenv(PROJECT_ROOT / ".env")

sys.path.insert(0, str(SCRIPT_DIR))
from tally_connector import TallyConnector


# -- Ledger Config ----------------------------------------------------------

GST_INPUT_LEDGERS = [
    "INPUT CGST",
    "Input CGST-RCM",
    "INPUT IGST",
    "INPUT IGST ON IMPORTS",
    "INPUT SGST",
    "Input SGST-RCM",
]

GST_OUTPUT_LEDGERS = [
    "Output CGST @ 9%",
    "Output IGST @ 18%",
    "Output SGST @ 9%",
]

GST_PAYABLE_LEDGER = "GST Payable"
ALL_GST_LEDGERS = GST_INPUT_LEDGERS + GST_OUTPUT_LEDGERS

# Indian FY order: Apr-Mar
FY_MONTHS = [
    ("APR", 4), ("MAY", 5), ("JUN", 6),
    ("JUL", 7), ("AUG", 8), ("SEP", 9),
    ("OCT", 10), ("NOV", 11), ("DEC", 12),
    ("JAN", 1),  ("FEB", 2),  ("MAR", 3),
]

MONTH_NAMES = {
    "APR": "April",   "MAY": "May",      "JUN": "June",
    "JUL": "July",    "AUG": "August",   "SEP": "September",
    "OCT": "October", "NOV": "November", "DEC": "December",
    "JAN": "January", "FEB": "February", "MAR": "March",
}


# -- Date Helpers -----------------------------------------------------------

def fmt_date(d: date) -> str:
    return f"{d.year:04d}{d.month:02d}{d.day:02d}"


def month_end_date(year: int, month: int) -> date:
    last = calendar.monthrange(year, month)[1]
    return date(year, month, last)


def build_fy_months(fy_year: int) -> List[Tuple[str, int, int, date]]:
    months = []
    for label, month_num in FY_MONTHS:
        year = fy_year if month_num >= 4 else fy_year + 1
        months.append((label, year, month_num, month_end_date(year, month_num)))
    return months


# -- Tally XML Helpers ------------------------------------------------------

def sanitize_xml(text: str) -> str:
    text = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', text)
    text = re.sub(r'&#x([0-8BbCcEeFf]);', '', text)
    text = re.sub(r'&#x1[0-9A-Fa-f];', '', text)
    return text


def parse_tally_amount(text: str) -> float:
    """
    Parse a Tally amount string.
    Tally encodes Dr amounts as negative values, Cr as positive.
    Handles formats: '1234.56', '-1234.56', '1234.56 Dr', '1234.56 Cr'
    """
    if not text or not text.strip():
        return 0.0
    text = text.strip().replace(",", "")
    for suffix in (" Cr.", " Cr", " Dr.", " Dr"):
        if text.endswith(suffix):
            text = text[:-len(suffix)].strip()
            break
    try:
        return float(text)
    except ValueError:
        return 0.0


# -- Method 1: Duties & Taxes CHILDOF Collection (primary) -----------------

def fetch_gst_balances_duties_taxes(
    tally: TallyConnector,
    as_of: date,
    timeout: int = 120,
) -> Dict[str, float]:
    """
    Fetch closing balances for all ledgers under the 'Duties & Taxes' group
    using the CHILDOF built-in filter — no custom TDL formulas needed.

    CLOSINGBALANCE from a Ledger collection is cumulative as of SVTODATE
    for balance-sheet ledgers (which GST ledgers always are).

    Returns {ledger_name: net} where negative = Dr, positive = Cr.
    Only ledgers matching our target GST list are returned.
    """
    xml_payload = f"""<ENVELOPE>
  <HEADER>
    <VERSION>1</VERSION>
    <TALLYREQUEST>Export</TALLYREQUEST>
    <TYPE>Collection</TYPE>
    <ID>DutiesTaxLedgers</ID>
  </HEADER>
  <BODY>
    <DESC>
      <STATICVARIABLES>
        <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
        <SVCURRENTCOMPANY>{xml_escape(tally.company)}</SVCURRENTCOMPANY>
        <SVFROMDATE>20000401</SVFROMDATE>
        <SVTODATE>{fmt_date(as_of)}</SVTODATE>
      </STATICVARIABLES>
      <TDL>
        <TDLMESSAGE>
          <COLLECTION NAME="DutiesTaxLedgers" ISMODIFY="No">
            <TYPE>Ledger</TYPE>
            <CHILDOF>Duties &amp; Taxes</CHILDOF>
            <FETCH>NAME, PARENT, CLOSINGBALANCE</FETCH>
          </COLLECTION>
        </TDLMESSAGE>
      </TDL>
    </DESC>
  </BODY>
</ENVELOPE>"""

    try:
        resp = tally._send_xml(xml_payload, timeout=timeout)
    except Exception:
        return {}

    if not resp or not resp.strip():
        return {}

    # Save raw for diagnostics
    raw_path = TMP_DIR / "_duties_taxes_raw.xml"
    try:
        raw_path.write_text(resp, encoding="utf-8")
    except Exception:
        pass

    try:
        root = ET.fromstring(sanitize_xml(resp))
    except ET.ParseError:
        return {}

    target_map = {name.lower(): name for name in ALL_GST_LEDGERS}
    balances: Dict[str, float] = {}

    for ldg in root.iter("LEDGER"):
        name = (ldg.get("NAME") or ldg.findtext(".//NAME", "")).strip()
        if not name:
            continue
        canonical = target_map.get(name.lower())
        if not canonical:
            continue
        closing_text = ldg.findtext("CLOSINGBALANCE", "0").strip()
        net = parse_tally_amount(closing_text)
        balances[canonical] = net

    return balances


# -- Method 2: Ledger Vouchers Report (fallback per-ledger) -----------------

def fetch_ledger_closing_balance(
    tally: TallyConnector,
    ledger_name: str,
    fy_start: date,
    as_of: date,
    timeout: int = 90,
) -> Optional[float]:
    """
    Fetch the closing balance of a single ledger using Tally's built-in
    Ledger Vouchers report.

    This avoids the EXPLODEFLAG sub-group depth limitation.
    Returns net balance: negative = Dr, positive = Cr.
    Returns None on error.

    Tally's Ledger Vouchers report XML contains DSPCLBAL (closing balance)
    as the last DSPACCOP block after all voucher entries.
    """
    xml_payload = f"""<ENVELOPE>
  <HEADER>
    <VERSION>1</VERSION>
    <TALLYREQUEST>Export</TALLYREQUEST>
    <TYPE>Data</TYPE>
    <ID>Ledger Vouchers</ID>
  </HEADER>
  <BODY>
    <DESC>
      <STATICVARIABLES>
        <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
        <SVCURRENTCOMPANY>{xml_escape(tally.company)}</SVCURRENTCOMPANY>
        <SVFROMDATE>{fmt_date(fy_start)}</SVFROMDATE>
        <SVTODATE>{fmt_date(as_of)}</SVTODATE>
        <SVLEDGERNAME>{xml_escape(ledger_name)}</SVLEDGERNAME>
      </STATICVARIABLES>
    </DESC>
  </BODY>
</ENVELOPE>"""

    try:
        resp = tally._send_xml(xml_payload, timeout=timeout)
    except Exception:
        return None

    if not resp or not resp.strip():
        return None

    # Save raw for diagnostics
    raw_path = TMP_DIR / f"_ledger_{re.sub(r'[^a-zA-Z0-9]', '_', ledger_name)}.xml"
    try:
        raw_path.write_text(resp, encoding="utf-8")
    except Exception:
        pass

    try:
        root = ET.fromstring(sanitize_xml(resp))
    except ET.ParseError:
        return None

    # The Ledger Vouchers report has DSPACCOP blocks; the LAST one contains
    # the closing balance in DSPCLBAL. Opening balance is in the FIRST block.
    # Each block has DSPCLDRAMTA (Dr amount, negative) and DSPCLCRAMTA (Cr amount, positive).
    closing_dr = None
    closing_cr = None

    # Try DSPCLBAL tag first (some Tally versions)
    for node in root.iter("DSPCLBAL"):
        text = node.text or ""
        if text.strip():
            closing_dr = parse_tally_amount(text)

    # Try DSPACCOP/DSPACCINFO blocks (standard format)
    if closing_dr is None:
        acc_op_blocks = list(root.iter("DSPACCOP"))
        if acc_op_blocks:
            last_block = acc_op_blocks[-1]
            dr_text = last_block.findtext(".//DSPCLDRAMTA", "")
            cr_text = last_block.findtext(".//DSPCLCRAMTA", "")
            if dr_text or cr_text:
                closing_dr = parse_tally_amount(dr_text)
                closing_cr = parse_tally_amount(cr_text)

    # Also try LDGVCH closing balance tags
    if closing_dr is None:
        for tag in ["CLBAL", "CLOSINGBALANCE", "DSPLDGCL"]:
            nodes = list(root.iter(tag))
            if nodes:
                last = nodes[-1]
                text = last.text or ""
                if text.strip():
                    closing_dr = parse_tally_amount(text)
                    break

    if closing_dr is not None and closing_cr is not None:
        return closing_dr + closing_cr
    if closing_dr is not None:
        return closing_dr
    return None


def fetch_gst_balances_ledger_vouchers(
    tally: TallyConnector,
    fy_start: date,
    as_of: date,
    timeout: int = 90,
) -> Dict[str, float]:
    """
    Fetch closing balances for all GST ledgers using individual
    Ledger Vouchers report queries. One query per ledger (10 total).
    """
    balances: Dict[str, float] = {}
    for ledger in ALL_GST_LEDGERS:
        bal = fetch_ledger_closing_balance(tally, ledger, fy_start, as_of, timeout)
        if bal is not None:
            balances[ledger] = bal
    return balances


# -- Method 2: Trial Balance EXPLODEFLAG (fallback) -------------------------

def fetch_gst_balances_explodeflag(
    tally: TallyConnector,
    fy_start: date,
    as_of: date,
    timeout: int = 180,
) -> Dict[str, float]:
    """
    Fallback: fetch closing balances via Trial Balance EXPLODEFLAG.
    Works for ledgers that appear at top level. May miss deep sub-ledgers.
    """
    xml_payload = f"""<ENVELOPE>
  <HEADER>
    <VERSION>1</VERSION>
    <TALLYREQUEST>Export</TALLYREQUEST>
    <TYPE>Data</TYPE>
    <ID>Trial Balance</ID>
  </HEADER>
  <BODY>
    <DESC>
      <STATICVARIABLES>
        <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
        <SVCURRENTCOMPANY>{xml_escape(tally.company)}</SVCURRENTCOMPANY>
        <SVFROMDATE>{fmt_date(fy_start)}</SVFROMDATE>
        <SVTODATE>{fmt_date(as_of)}</SVTODATE>
        <EXPLODEFLAG>Yes</EXPLODEFLAG>
      </STATICVARIABLES>
    </DESC>
  </BODY>
</ENVELOPE>"""

    try:
        resp = tally._send_xml(xml_payload, timeout=timeout)
    except Exception:
        return {}

    try:
        root = ET.fromstring(sanitize_xml(resp))
    except ET.ParseError:
        return {}

    target_map = {name.lower(): name for name in ALL_GST_LEDGERS}
    balances: Dict[str, float] = {}

    children = list(root)
    i = 0
    while i < len(children):
        if children[i].tag == "DSPACCNAME":
            display_name = (children[i].findtext("DSPDISPNAME") or "").strip()
            if i + 1 < len(children) and children[i + 1].tag == "DSPACCINFO":
                canonical = target_map.get(display_name.lower())
                if canonical:
                    info = children[i + 1]
                    dr = parse_tally_amount(info.findtext(".//DSPCLDRAMTA", ""))
                    cr = parse_tally_amount(info.findtext(".//DSPCLCRAMTA", ""))
                    balances[canonical] = dr + cr
                i += 2
                continue
        i += 1
    return balances


# -- Combined Fetch with Fallback -------------------------------------------

def fetch_gst_closing_balances(
    tally: TallyConnector,
    fy_start: date,
    as_of: date,
) -> Tuple[Dict[str, float], str]:
    """
    Fetch GST ledger closing balances using the best available method.
    Returns (balances_dict, method_used).

    Priority:
      1. Duties & Taxes CHILDOF collection  -- fast, no custom TDL
      2. Ledger Vouchers per-ledger report  -- handles any group structure
      3. Trial Balance EXPLODEFLAG          -- last resort, may miss sub-ledgers
    """
    # Method 1: Duties & Taxes CHILDOF (fastest, no TDL compilation)
    print("  Fetching via Duties & Taxes group (CHILDOF)...", end="", flush=True)
    balances = fetch_gst_balances_duties_taxes(tally, as_of)
    non_zero = sum(1 for v in balances.values() if abs(v) > 0.01)
    if non_zero > 0:
        print(f" OK ({non_zero}/{len(ALL_GST_LEDGERS)} ledgers found)")
        return balances, "DutiesTaxes-CHILDOF"
    print(f" {len(balances)} ledgers found but all zero")

    # Method 2: Ledger Vouchers per-ledger
    print("  Fetching via Ledger Vouchers (per-ledger)...", end="", flush=True)
    balances = fetch_gst_balances_ledger_vouchers(tally, fy_start, as_of)
    non_zero = sum(1 for v in balances.values() if abs(v) > 0.01)
    if non_zero > 0:
        print(f" OK ({non_zero} non-zero ledgers)")
        return balances, "LedgerVouchers"
    print(" returned all zeros")

    # Method 3: EXPLODEFLAG fallback
    print("  Falling back to EXPLODEFLAG...", end="", flush=True)
    balances = fetch_gst_balances_explodeflag(tally, fy_start, as_of)
    non_zero = sum(1 for v in balances.values() if abs(v) > 0.01)
    if non_zero > 0:
        print(f" OK ({non_zero} non-zero ledgers)")
        return balances, "EXPLODEFLAG"
    print(" returned all zeros")

    return {}, "none"


# -- Diagnose: dump raw Ledger Vouchers response ----------------------------

def diagnose_ledger(tally: TallyConnector, ledger_name: str, fy_start: date, as_of: date):
    """
    Dump the raw Tally response for a single ledger to help debug
    which XML tags contain the closing balance.
    """
    print(f"\nDIAGNOSTICS for ledger: {ledger_name!r}")
    print(f"Period: {fy_start} to {as_of}")

    xml_payload = f"""<ENVELOPE>
  <HEADER>
    <VERSION>1</VERSION>
    <TALLYREQUEST>Export</TALLYREQUEST>
    <TYPE>Data</TYPE>
    <ID>Ledger Vouchers</ID>
  </HEADER>
  <BODY>
    <DESC>
      <STATICVARIABLES>
        <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
        <SVCURRENTCOMPANY>{xml_escape(tally.company)}</SVCURRENTCOMPANY>
        <SVFROMDATE>{fmt_date(fy_start)}</SVFROMDATE>
        <SVTODATE>{fmt_date(as_of)}</SVTODATE>
        <SVLEDGERNAME>{xml_escape(ledger_name)}</SVLEDGERNAME>
      </STATICVARIABLES>
    </DESC>
  </BODY>
</ENVELOPE>"""

    print("Sending request...", end="", flush=True)
    try:
        resp = tally._send_xml(xml_payload, timeout=120)
    except Exception as e:
        print(f" FAIL: {e}")
        return

    print(f" OK ({len(resp)} bytes)")
    out_path = TMP_DIR / f"_diag_{re.sub(r'[^a-zA-Z0-9]','_',ledger_name)}.xml"
    out_path.write_text(resp, encoding="utf-8")
    print(f"Raw saved: {out_path}")

    # Print all unique tags and any amount-looking content
    print("\nAll tags with numeric content:")
    try:
        root = ET.fromstring(sanitize_xml(resp))
        seen_tags = {}
        for elem in root.iter():
            if elem.text and elem.text.strip():
                text = elem.text.strip()
                if re.search(r'\d', text):
                    if elem.tag not in seen_tags:
                        seen_tags[elem.tag] = text
        for tag, val in list(seen_tags.items())[:40]:
            print(f"  <{tag}>  {val[:80]}")
    except Exception as e:
        print(f"Parse error: {e}")
        print("First 1000 chars of response:")
        print(resp[:1000])


# -- Journal XML Builder ----------------------------------------------------

def build_provision_journal_xml(
    company: str,
    month_label: str,
    voucher_date: str,
    balances: Dict[str, float],
    voucher_number: str,
) -> Tuple[Optional[str], float, str]:
    """
    Build a Journal voucher XML that reverses all non-zero GST ledger
    closing balances, with GST Payable as the balancing entry.

    Tally Journal XML convention:
      Credit entry: ISDEEMEDPOSITIVE=No,  AMOUNT=+positive
      Debit  entry: ISDEEMEDPOSITIVE=Yes, AMOUNT=-positive (negative value)

    Reversal: reversal_amount = -net  (flips Dr<->Cr)

    Returns (xml_string, gst_payable_amount, status_label)
    """
    entries_xml = ""
    running_sum = 0.0

    for ledger in ALL_GST_LEDGERS:
        net = balances.get(ledger, 0.0)
        if abs(net) < 0.005:
            continue

        reversal_amount = round(-net, 2)
        running_sum += reversal_amount
        is_deemed = "No" if reversal_amount >= 0 else "Yes"

        entries_xml += f"""
        <ALLLEDGERENTRIES.LIST>
          <LEDGERNAME>{xml_escape(ledger)}</LEDGERNAME>
          <ISDEEMEDPOSITIVE>{is_deemed}</ISDEEMEDPOSITIVE>
          <AMOUNT>{reversal_amount}</AMOUNT>
        </ALLLEDGERENTRIES.LIST>"""

    if abs(running_sum) < 0.005 and not entries_xml:
        return None, 0.0, ""

    gst_payable_amount = round(-running_sum, 2)
    if gst_payable_amount > 0:
        gst_deemed = "No"
        status = "PAYABLE (Output > Input)"
    elif gst_payable_amount < 0:
        gst_deemed = "Yes"
        status = "INPUT EXCESS (Input > Output)"
    else:
        gst_deemed = "No"
        status = "NIL (perfectly offset)"

    entries_xml += f"""
        <ALLLEDGERENTRIES.LIST>
          <LEDGERNAME>{xml_escape(GST_PAYABLE_LEDGER)}</LEDGERNAME>
          <ISDEEMEDPOSITIVE>{gst_deemed}</ISDEEMEDPOSITIVE>
          <AMOUNT>{gst_payable_amount}</AMOUNT>
        </ALLLEDGERENTRIES.LIST>"""

    narration = f"GST Provision {month_label} - Input/Output Set-off"

    envelope = f"""<ENVELOPE>
  <HEADER>
    <VERSION>1</VERSION>
    <TALLYREQUEST>Import</TALLYREQUEST>
    <TYPE>Data</TYPE>
    <ID>Vouchers</ID>
  </HEADER>
  <BODY>
    <DESC>
      <STATICVARIABLES>
        <SVCURRENTCOMPANY>{xml_escape(company)}</SVCURRENTCOMPANY>
      </STATICVARIABLES>
    </DESC>
    <DATA>
      <TALLYMESSAGE>
        <VOUCHER VCHTYPE="Journal" ACTION="Create">
          <DATE>{voucher_date}</DATE>
          <VOUCHERTYPENAME>Journal</VOUCHERTYPENAME>
          <PERSISTEDVIEW>Accounting Voucher View</PERSISTEDVIEW>
          <NARRATION>{xml_escape(narration)}</NARRATION>
          <VOUCHERNUMBER>{xml_escape(voucher_number)}</VOUCHERNUMBER>
          {entries_xml}
        </VOUCHER>
      </TALLYMESSAGE>
    </DATA>
  </BODY>
</ENVELOPE>"""

    return envelope, gst_payable_amount, status


# -- Print Journal Preview --------------------------------------------------

def print_journal_preview(month_label: str, balances: Dict[str, float],
                           gst_payable_amount: float, status: str):
    print(f"\n  Journal Entry: {month_label}")
    print(f"  {'Particulars':<35} {'Dr':>14} {'Cr':>14}")
    print(f"  {'-'*63}")
    total_dr = total_cr = 0.0
    for ledger in GST_OUTPUT_LEDGERS:
        net = balances.get(ledger, 0.0)
        if abs(net) < 0.005:
            continue
        amt = abs(net)
        print(f"  {ledger:<35} {amt:>14,.2f} {'':>14}")
        total_dr += amt
    for ledger in GST_INPUT_LEDGERS:
        net = balances.get(ledger, 0.0)
        if abs(net) < 0.005:
            continue
        amt = abs(net)
        print(f"  {ledger:<35} {'':>14} {amt:>14,.2f}")
        total_cr += amt
    if gst_payable_amount >= 0:
        print(f"  {GST_PAYABLE_LEDGER:<35} {'':>14} {gst_payable_amount:>14,.2f}")
        total_cr += gst_payable_amount
    else:
        amt = abs(gst_payable_amount)
        print(f"  {GST_PAYABLE_LEDGER:<35} {amt:>14,.2f} {'':>14}")
        total_dr += amt
    print(f"  {'-'*63}")
    print(f"  {'TOTAL':<35} {total_dr:>14,.2f} {total_cr:>14,.2f}")
    print(f"  Result: {status}")


# -- Manual Input Mode ------------------------------------------------------

def collect_manual_balances(month_label: str, year: int) -> Dict[str, float]:
    """
    Interactively collect GST ledger balances from the user.
    User enters amounts from Tally Balance Sheet.
    Dr balances are entered as positive; the function stores them as negative.
    """
    print(f"\nManual balance entry for {month_label} {year}")
    print("Enter CLOSING BALANCE for each ledger (from Tally Balance Sheet).")
    print("For Dr balance: enter positive number.  For Cr balance: enter negative.")
    print("Press Enter to skip (treat as zero).\n")

    balances: Dict[str, float] = {}

    print("-- Input GST Ledgers (normally Dr balance) --")
    for ledger in GST_INPUT_LEDGERS:
        while True:
            raw = input(f"  {ledger}: ").strip().replace(",", "")
            if not raw:
                break
            try:
                val = float(raw)
                # Dr balance entered as positive -> store as negative (Tally convention)
                balances[ledger] = -abs(val) if val > 0 else val
                break
            except ValueError:
                print("  Invalid number, try again.")

    print("\n-- Output GST Ledgers (normally Cr balance) --")
    for ledger in GST_OUTPUT_LEDGERS:
        while True:
            raw = input(f"  {ledger}: ").strip().replace(",", "")
            if not raw:
                break
            try:
                val = float(raw)
                # Cr balance entered as positive -> store as positive (Tally convention)
                balances[ledger] = abs(val) if val > 0 else val
                break
            except ValueError:
                print("  Invalid number, try again.")

    return balances


# -- Main -------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Generate GST provision Journal vouchers (Apr-Mar)"
    )
    parser.add_argument("--year", type=int, default=None,
                        help="FY start year (e.g. 2025 for FY 2025-26)")
    parser.add_argument("--months", type=str, default=None,
                        help="Comma-separated month labels, e.g. APR,MAY,JUN")
    parser.add_argument("--dry-run", action="store_true",
                        help="Preview journal entries, do NOT post to Tally")
    parser.add_argument("--save-xml", action="store_true",
                        help="Save Journal XMLs to .tmp/gst_provision/")
    parser.add_argument("--import", dest="do_import", action="store_true",
                        help="Post Journal vouchers to Tally Prime")
    parser.add_argument("--manual", nargs=2, metavar=("MONTH", "YEAR"),
                        help="Manual balance entry mode, e.g. --manual APR 2025")
    parser.add_argument("--diagnose", action="store_true",
                        help="Dump raw Tally responses to .tmp/ for debugging")
    parser.add_argument("--timeout", type=int, default=90,
                        help="Per-ledger query timeout in seconds (default: 90)")
    args = parser.parse_args()

    TMP_DIR.mkdir(parents=True, exist_ok=True)

    # -- Determine FY -------------------------------------------------------
    today = date.today()
    fy_year = args.year if args.year else (today.year if today.month >= 4 else today.year - 1)
    fy_start = date(fy_year, 4, 1)
    fy_label = f"FY {fy_year}-{str(fy_year + 1)[2:]}"
    print(f"\nGST Provision Journal Generator -- {fy_label}")

    # -- Manual mode: single month, user-entered balances ------------------
    if args.manual:
        month_label = args.manual[0].upper()
        try:
            manual_year = int(args.manual[1])
        except ValueError:
            print(f"Invalid year: {args.manual[1]}")
            sys.exit(1)

        month_num = dict(FY_MONTHS).get(month_label)
        if not month_num:
            print(f"Unknown month: {month_label}. Use APR, MAY, JUN... MAR")
            sys.exit(1)

        me = month_end_date(manual_year, month_num)
        period_label = f"{month_label} {manual_year}"
        voucher_number = f"GST-PROV-{month_label}-{manual_year}"

        balances = collect_manual_balances(month_label, manual_year)
        if not balances:
            print("No balances entered.")
            return

        xml_str, gst_payable_amount, status = build_provision_journal_xml(
            "", period_label, fmt_date(me), balances, voucher_number
        )
        if xml_str is None:
            print("All balances zero after entry -- nothing to post.")
            return

        print_journal_preview(period_label, balances, gst_payable_amount, status)

        if args.save_xml or args.do_import:
            # Need company name for actual posting
            tally = TallyConnector()
            xml_str, _, _ = build_provision_journal_xml(
                tally.company, period_label, fmt_date(me), balances, voucher_number
            )

        if args.save_xml:
            save_dir = TMP_DIR / "gst_provision"
            save_dir.mkdir(parents=True, exist_ok=True)
            xml_path = save_dir / f"{voucher_number}.xml"
            xml_path.write_text(xml_str, encoding="utf-8")
            print(f"\nSaved: {xml_path}")

        if args.do_import:
            print("\nPosting to Tally...")
            res = tally.import_voucher_xml(xml_str, timeout=120)
            if res["errors"] > 0:
                print(f"FAIL: {res['error_message']}")
            elif res["created"] > 0:
                print(f"OK: Posted {voucher_number} to Tally")
            else:
                print(f"Response: {res}")

        if not args.save_xml and not args.do_import:
            print("\nAdd --save-xml to export XML, or --import to post to Tally.")
        return

    # -- Auto mode: connect to Tally and fetch balances --------------------
    tally = TallyConnector()
    print(f"Connecting to Tally at {tally.base_url}...")
    try:
        info = tally.test_connection()
        print(f"Connected: {info['target_company']}\n")
    except Exception as e:
        print(f"FAIL: {e}")
        print("\nTally is not responding.")
        print("Use --manual APR 2025 to enter balances manually from Tally screen.")
        sys.exit(1)

    # -- Diagnose mode ------------------------------------------------------
    if args.diagnose:
        me = month_end_date(fy_year, 4)  # use April as test
        # Show all ledgers found under Duties & Taxes
        print(f"\nDIAGNOSTICS -- Duties & Taxes CHILDOF query (as of {me})")
        balances = fetch_gst_balances_duties_taxes(tally, me, timeout=120)
        if balances:
            print(f"Found {len(balances)} matching GST ledgers:")
            for name, net in balances.items():
                bal_type = "Cr" if net > 0 else ("Dr" if net < 0 else "Nil")
                print(f"  [{bal_type}] {net:>14,.2f}  {name}")
        else:
            print("No matching GST ledgers found under Duties & Taxes.")
            print("Trying Ledger Vouchers for INPUT CGST...")
            diagnose_ledger(tally, "INPUT CGST", fy_start, me)
            diagnose_ledger(tally, "Output CGST @ 9%", fy_start, me)
        return

    # -- Build month list ---------------------------------------------------
    all_months = build_fy_months(fy_year)
    if args.months:
        filter_labels = {m.strip().upper() for m in args.months.split(",")}
        all_months = [m for m in all_months if m[0] in filter_labels]
    all_months = [m for m in all_months if m[3] <= today]

    if not all_months:
        print("No completed months to process.")
        return

    print(f"Processing: {', '.join(f'{m[0]} {m[1]}' for m in all_months)}\n")

    if args.save_xml:
        save_dir = TMP_DIR / "gst_provision"
        save_dir.mkdir(parents=True, exist_ok=True)

    results = []

    for label, year, month_num, me in all_months:
        period_label = f"{label} {year}"
        print(f"[{period_label}] {me}")

        balances, method = fetch_gst_closing_balances(tally, fy_start, me)

        if not balances:
            print(f"  No balances found -- skipping. Use --manual {label} {year} to enter manually.")
            time.sleep(1)
            continue

        voucher_number = f"GST-PROV-{label}-{year}"
        xml_str, gst_payable_amount, status = build_provision_journal_xml(
            tally.company, period_label, fmt_date(me), balances, voucher_number
        )

        if xml_str is None:
            print(f"  All balances zero -- skipped.")
            time.sleep(1)
            continue

        print_journal_preview(period_label, balances, gst_payable_amount, status)

        results.append({
            "label": period_label,
            "date": str(me),
            "method": method,
            "balances": {k: round(v, 2) for k, v in balances.items()},
            "gst_payable": gst_payable_amount,
            "status": status,
            "voucher_number": voucher_number,
        })

        if args.save_xml:
            xml_path = save_dir / f"{voucher_number}.xml"
            xml_path.write_text(xml_str, encoding="utf-8")
            print(f"  Saved: {xml_path.name}")

        if args.do_import and not args.dry_run:
            print("  Posting to Tally...", end="", flush=True)
            try:
                res = tally.import_voucher_xml(xml_str, timeout=120)
                if res["errors"] > 0:
                    print(f" FAIL: {res['error_message']}")
                elif res["created"] > 0:
                    print(f" OK: {voucher_number}")
                else:
                    print(f" Response: {res}")
            except Exception as e:
                print(f" FAIL: {e}")

        time.sleep(1)

    # -- FY Summary ---------------------------------------------------------
    if not results:
        print("\nNo data collected.")
        return

    print(f"\n{'='*65}")
    print(f"GST PROVISION SUMMARY -- {fy_label}")
    print(f"{'='*65}")
    print(f"{'Month':<16} {'GST Payable':>16}  Status")
    print(f"{'-'*65}")
    total = 0.0
    for r in results:
        print(f"{r['label']:<16} {r['gst_payable']:>16,.2f}  {r['status']}")
        total += r["gst_payable"]
    print(f"{'-'*65}")
    fy_status = "PAYABLE" if total > 0 else ("INPUT EXCESS" if total < 0 else "NIL")
    print(f"{'FY Total':<16} {total:>16,.2f}  {fy_status}")

    summary_path = TMP_DIR / f"gst_provision_summary_{fy_year}.json"
    with open(summary_path, "w", encoding="utf-8") as f:
        json.dump(results, f, indent=2, default=str)
    print(f"\nSummary saved: {summary_path}")

    if args.dry_run:
        print("\n[DRY RUN] Vouchers NOT posted to Tally.")
    elif not args.do_import:
        print("\nAdd --import to post to Tally, or --save-xml to export XMLs.")


if __name__ == "__main__":
    main()
