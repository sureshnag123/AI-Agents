#!/usr/bin/env python3
"""
Repeat Customer Report - CORRECT version
Only counts Sales vouchers where the Party Ledger is under Sundry Debtors.
This filters out expenses, bank entries, purchases etc. that are booked under
the "Sales" voucher type in Tally.
"""

import sys
import re
import json
import xml.etree.ElementTree as ET
from pathlib import Path
from collections import defaultdict

SCRIPT_DIR = Path(__file__).parent.resolve()
PROJECT_ROOT = SCRIPT_DIR.parent
TMP_DIR = PROJECT_ROOT / ".tmp"

sys.path.insert(0, str(SCRIPT_DIR))
from tally_connector import TallyConnector


def sanitize_xml(xml_text: str) -> ET.Element:
    xml_text = re.sub(r'&#x([0-8BbCcEeFf]);', '', xml_text)
    xml_text = re.sub(r'&#x1[0-9A-Fa-f];', '', xml_text)
    xml_text = re.sub(r'&#([0-8]|1[0-1]|1[4-9]|2[0-9]|3[01]);', '', xml_text)
    xml_text = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', xml_text)
    xml_text = re.sub(r'<(/?)(\w+):', r'<\1\2_', xml_text)
    xml_text = re.sub(r'\s+xmlns:\w+="[^"]*"', '', xml_text)
    return ET.fromstring(xml_text)


def main():
    tally = TallyConnector()
    info = tally.test_connection()
    print(f"Connected: {info['target_company']}")

    # ── Step 1: Get all Sundry Debtors ledgers ──
    print("\nFetching Sundry Debtors ledger list...")
    xml_debtors = f"""<ENVELOPE>
  <HEADER>
    <VERSION>1</VERSION>
    <TALLYREQUEST>Export</TALLYREQUEST>
    <TYPE>Collection</TYPE>
    <ID>SundryDebtorsList</ID>
  </HEADER>
  <BODY>
    <DESC>
      <STATICVARIABLES>
        <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
        <SVCURRENTCOMPANY>{tally.company}</SVCURRENTCOMPANY>
      </STATICVARIABLES>
      <TDL>
        <TDLMESSAGE>
          <COLLECTION NAME="SundryDebtorsList" ISMODIFY="No">
            <TYPE>Ledger</TYPE>
            <BELONGSTO>Yes</BELONGSTO>
            <CHILDOF>Sundry Debtors</CHILDOF>
            <FETCH>NAME, PARENT</FETCH>
          </COLLECTION>
        </TDLMESSAGE>
      </TDL>
    </DESC>
  </BODY>
</ENVELOPE>"""

    resp_debtors = tally._send_xml(xml_debtors, timeout=30)
    root_debtors = sanitize_xml(resp_debtors)

    debtor_names = set()
    for ldg in root_debtors.iter("LEDGER"):
        name = (ldg.get("NAME") or ldg.findtext(".//NAME", "")).strip()
        if name:
            debtor_names.add(name)

    print(f"Total Sundry Debtors ledgers: {len(debtor_names)}")

    # ── Step 2: Fetch ALL Sales vouchers for FY 2025-26 ──
    from_date = "20250401"
    to_date = "20260331"

    print(f"\nFetching Sales vouchers ({from_date} to {to_date})...")
    xml_request = f"""<ENVELOPE>
  <HEADER>
    <VERSION>1</VERSION>
    <TALLYREQUEST>Export</TALLYREQUEST>
    <TYPE>Collection</TYPE>
    <ID>SalesVoucherFull</ID>
  </HEADER>
  <BODY>
    <DESC>
      <STATICVARIABLES>
        <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
        <SVCURRENTCOMPANY>{tally.company}</SVCURRENTCOMPANY>
        <SVFROMDATE>{from_date}</SVFROMDATE>
        <SVTODATE>{to_date}</SVTODATE>
      </STATICVARIABLES>
      <TDL>
        <TDLMESSAGE>
          <COLLECTION NAME="SalesVoucherFull" ISMODIFY="No">
            <TYPE>Voucher</TYPE>
            <CHILDOF>Sales</CHILDOF>
            <FETCH>DATE, VOUCHERNUMBER, PARTYLEDGERNAME, NARRATION, AMOUNT, REFERENCE, ALLLEDGERENTRIES.LIST</FETCH>
          </COLLECTION>
        </TDLMESSAGE>
      </TDL>
    </DESC>
  </BODY>
</ENVELOPE>"""

    resp_text = tally._send_xml(xml_request, timeout=180)
    root = sanitize_xml(resp_text)

    all_vouchers = list(root.iter("VOUCHER"))
    print(f"Total vouchers under Sales type: {len(all_vouchers)}")

    # ── Step 3: Filter - only keep vouchers where party OR a ledger entry is a Sundry Debtor ──
    customer_data = defaultdict(lambda: {"invoices": [], "total_amount": 0.0})
    sales_only_count = 0
    non_sales_count = 0
    non_sales_parties = defaultdict(int)

    for v in all_vouchers:
        party = (v.findtext("PARTYLEDGERNAME", "") or "").strip()
        number = v.findtext("VOUCHERNUMBER", "")
        vdate = v.findtext("DATE", "")
        narration = v.findtext("NARRATION", "")
        raw_amount = v.findtext("AMOUNT", "0")

        try:
            amount = abs(float(raw_amount.replace(",", "")))
        except (ValueError, AttributeError):
            amount = 0.0

        # Check if party itself is a Sundry Debtor
        customer_name = None
        if party in debtor_names:
            customer_name = party
        else:
            # Check ledger entries for any Sundry Debtor
            for entry in v.iter("ALLLEDGERENTRIES.LIST"):
                ledger = (entry.findtext("LEDGERNAME", "") or "").strip()
                if ledger in debtor_names:
                    customer_name = ledger
                    # Get the amount from this specific entry
                    entry_amt = entry.findtext("AMOUNT", "0")
                    try:
                        amount = abs(float(entry_amt.replace(",", "")))
                    except (ValueError, AttributeError):
                        pass
                    break

        if customer_name:
            sales_only_count += 1
            customer_data[customer_name]["invoices"].append({
                "number": number,
                "date": vdate,
                "amount": amount,
                "narration": (narration or "")[:100],
            })
            customer_data[customer_name]["total_amount"] += amount
        else:
            non_sales_count += 1
            non_sales_parties[party] += 1

    print(f"\nVouchers with Sundry Debtor (actual sales): {sales_only_count}")
    print(f"Vouchers WITHOUT Sundry Debtor (non-sales): {non_sales_count}")

    # Show top non-sales parties
    print(f"\nTop non-customer parties (excluded):")
    for p, cnt in sorted(non_sales_parties.items(), key=lambda x: -x[1])[:15]:
        print(f"  {p[:55]:<58} {cnt:>5} vouchers")

    # ── Step 4: Separate repeat vs one-time ──
    repeat = {}
    one_time = {}
    for name, data in customer_data.items():
        if len(data["invoices"]) >= 2:
            repeat[name] = data
        else:
            one_time[name] = data

    sorted_repeat = sorted(repeat.items(), key=lambda x: len(x[1]["invoices"]), reverse=True)

    # ── Step 5: Print report ──
    print(f"\n{'='*100}")
    print(f"REPEAT CUSTOMERS (SUNDRY DEBTORS ONLY) - TALLY SALES REGISTER - FY 2025-26")
    print(f"{'='*100}")
    print(f"Total Vouchers under Sales type:       {len(all_vouchers)}")
    print(f"Actual Sales (Sundry Debtor present):  {sales_only_count}")
    print(f"Non-sales entries (excluded):           {non_sales_count}")
    print(f"Unique Customers (Sundry Debtors):     {len(customer_data)}")
    print(f"Repeat Customers (2+ invoices):        {len(repeat)}")
    print(f"One-Time Customers:                    {len(one_time)}")
    print()

    hdr = f"{'#':>4}  {'Customer Name (Sundry Debtor)':<62} {'Inv':>5}  {'Total Amount (Rs)':>18}"
    print(hdr)
    print("-" * len(hdr))

    total_repeat_inv = 0
    total_repeat_amt = 0.0
    for i, (name, data) in enumerate(sorted_repeat, 1):
        inv_count = len(data["invoices"])
        total_amt = data["total_amount"]
        total_repeat_inv += inv_count
        total_repeat_amt += total_amt
        print(f"{i:>4}  {name[:61]:<62} {inv_count:>5}  {total_amt:>18,.2f}")

    print("-" * len(hdr))
    print(f"{'':>4}  {'TOTAL':<62} {total_repeat_inv:>5}  {total_repeat_amt:>18,.2f}")

    # ── Step 6: Save JSON ──
    output = {
        "summary": {
            "report": "Repeat Customers (Sundry Debtors) from Tally Sales Register",
            "period": "FY 2025-26 (Apr 2025 - Mar 2026)",
            "voucher_type": "Sales",
            "filter": "Only vouchers where Party Ledger is under Sundry Debtors group",
            "total_vouchers_under_sales_type": len(all_vouchers),
            "actual_sales_with_debtor": sales_only_count,
            "non_sales_excluded": non_sales_count,
            "unique_customers": len(customer_data),
            "repeat_customers": len(repeat),
            "one_time_customers": len(one_time),
            "total_repeat_invoices": total_repeat_inv,
            "total_repeat_amount": round(total_repeat_amt, 2),
        },
        "repeat_customers": [
            {
                "customer_name": name,
                "invoice_count": len(data["invoices"]),
                "total_amount": round(data["total_amount"], 2),
                "invoices": data["invoices"],
            }
            for name, data in sorted_repeat
        ],
        "excluded_non_customer_parties": dict(
            sorted(non_sales_parties.items(), key=lambda x: -x[1])
        ),
    }

    TMP_DIR.mkdir(parents=True, exist_ok=True)
    out_file = TMP_DIR / "repeat_customers_sundry_debtors_fy2526.json"
    with open(out_file, "w", encoding="utf-8") as f:
        json.dump(output, f, indent=2, ensure_ascii=False)
    print(f"\nReport saved to: {out_file}")


if __name__ == "__main__":
    main()
