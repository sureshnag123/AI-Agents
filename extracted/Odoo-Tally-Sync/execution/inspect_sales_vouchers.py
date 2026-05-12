#!/usr/bin/env python3
"""
Inspect raw Sales vouchers from Tally to understand data structure.
Fetches a few Sales vouchers with full ledger entries to see which
ledger is the actual customer (Sundry Debtor) vs revenue/tax/bank.
"""

import sys
import re
import json
import xml.etree.ElementTree as ET
from pathlib import Path

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

    # Step 1: Fetch a small sample of Sales vouchers with ALLLEDGERENTRIES detail
    # Use a short date range to get a manageable sample
    xml_request = f"""<ENVELOPE>
  <HEADER>
    <VERSION>1</VERSION>
    <TALLYREQUEST>Export</TALLYREQUEST>
    <TYPE>Collection</TYPE>
    <ID>SalesVoucherDetail</ID>
  </HEADER>
  <BODY>
    <DESC>
      <STATICVARIABLES>
        <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
        <SVCURRENTCOMPANY>{tally.company}</SVCURRENTCOMPANY>
        <SVFROMDATE>20250401</SVFROMDATE>
        <SVTODATE>20250410</SVTODATE>
      </STATICVARIABLES>
      <TDL>
        <TDLMESSAGE>
          <COLLECTION NAME="SalesVoucherDetail" ISMODIFY="No">
            <TYPE>Voucher</TYPE>
            <CHILDOF>Sales</CHILDOF>
            <FETCH>DATE, VOUCHERNUMBER, PARTYLEDGERNAME, BASICBUYERNAME, NARRATION, ALLLEDGERENTRIES.LIST</FETCH>
          </COLLECTION>
        </TDLMESSAGE>
      </TDL>
    </DESC>
  </BODY>
</ENVELOPE>"""

    print("\nFetching sample Sales vouchers (Apr 1-10, 2025) with ledger entries...")
    resp_text = tally._send_xml(xml_request, timeout=60)
    
    # Save raw XML for inspection
    raw_file = TMP_DIR / "sample_sales_vouchers_raw.xml"
    with open(raw_file, "w", encoding="utf-8") as f:
        f.write(resp_text)
    print(f"Raw XML saved to: {raw_file}")

    root = sanitize_xml(resp_text)

    vouchers = list(root.iter("VOUCHER"))
    print(f"Found {len(vouchers)} Sales vouchers in sample period")

    # Parse and display structure of first 5 vouchers
    sample = []
    for v in vouchers[:10]:
        vdata = {
            "number": v.findtext("VOUCHERNUMBER", ""),
            "date": v.findtext("DATE", ""),
            "party_ledger": v.findtext("PARTYLEDGERNAME", ""),
            "buyer_name": v.findtext("BASICBUYERNAME", ""),
            "narration": v.findtext("NARRATION", ""),
            "ledger_entries": []
        }

        # Get ALL LEDGER ENTRIES
        for entry in v.iter("ALLLEDGERENTRIES.LIST"):
            ledger_name = entry.findtext("LEDGERNAME", "")
            amount = entry.findtext("AMOUNT", "0")
            is_deemed = entry.findtext("ISDEEMEDPOSITIVE", "")
            vdata["ledger_entries"].append({
                "ledger": ledger_name,
                "amount": amount,
                "is_deemed_positive": is_deemed,
            })

        # Also check LEDGERENTRIES.LIST (Tally sometimes uses this)
        for entry in v.iter("LEDGERENTRIES.LIST"):
            ledger_name = entry.findtext("LEDGERNAME", "")
            amount = entry.findtext("AMOUNT", "0")
            is_deemed = entry.findtext("ISDEEMEDPOSITIVE", "")
            vdata["ledger_entries"].append({
                "ledger": ledger_name,
                "amount": amount,
                "is_deemed_positive": is_deemed,
                "source": "LEDGERENTRIES.LIST",
            })

        sample.append(vdata)

    # Print detailed view
    for i, v in enumerate(sample, 1):
        print(f"\n{'='*80}")
        print(f"Voucher #{i}: {v['number']}  Date: {v['date']}")
        print(f"  PARTYLEDGERNAME: {v['party_ledger']}")
        print(f"  BASICBUYERNAME:  {v['buyer_name']}")
        print(f"  NARRATION:       {v['narration'][:80]}")
        print(f"  Ledger Entries ({len(v['ledger_entries'])}):")
        for e in v['ledger_entries']:
            src = f" [{e.get('source','')}]" if e.get('source') else ""
            print(f"    - {e['ledger']:<50} Amt: {e['amount']:>15}  Deemed+: {e['is_deemed_positive']}{src}")

    # Save sample as JSON
    out = TMP_DIR / "sample_sales_voucher_structure.json"
    with open(out, "w", encoding="utf-8") as f:
        json.dump(sample, f, indent=2, ensure_ascii=False)
    print(f"\nSample saved to: {out}")

    # Step 2: Also fetch the Sundry Debtors ledger list
    print("\n\nFetching Sundry Debtors ledger list...")
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

    resp2 = tally._send_xml(xml_debtors, timeout=30)
    root2 = sanitize_xml(resp2)

    debtors = []
    for ldg in root2.iter("LEDGER"):
        name = (ldg.get("NAME") or ldg.findtext(".//NAME", "")).strip()
        parent = ldg.findtext("PARENT", "").strip()
        if name:
            debtors.append({"name": name, "group": parent})

    print(f"Total Sundry Debtors ledgers: {len(debtors)}")
    debtor_names = {d["name"] for d in debtors}

    # Check how many of the PARTYLEDGERNAME values from sample are actually Sundry Debtors
    print("\nChecking if PARTYLEDGERNAME from Sales vouchers are Sundry Debtors:")
    for v in sample:
        party = v["party_ledger"]
        is_debtor = party in debtor_names
        print(f"  {party[:50]:<52} -> {'SUNDRY DEBTOR' if is_debtor else 'NOT A DEBTOR'}")

        # Also check ledger entries
        for e in v["ledger_entries"]:
            is_d = e["ledger"] in debtor_names
            if is_d:
                print(f"    >> Found debtor in entries: {e['ledger']}")

    # Save debtors list
    out2 = TMP_DIR / "sundry_debtors_list.json"
    with open(out2, "w", encoding="utf-8") as f:
        json.dump(debtors, f, indent=2, ensure_ascii=False)
    print(f"\nSundry Debtors list saved to: {out2}")


if __name__ == "__main__":
    main()
