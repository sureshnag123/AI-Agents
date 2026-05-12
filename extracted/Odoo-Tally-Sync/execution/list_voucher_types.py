#!/usr/bin/env python3
"""
Step 1: List all voucher types in Tally and their parent types.
Step 2: Fetch sample vouchers from GST Invoice / Invoice type.
This helps identify the correct voucher type for revenue sales.
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

    # ── Step 1: Get ALL voucher types ──
    print("\n=== ALL VOUCHER TYPES IN TALLY ===")
    xml_vtypes = f"""<ENVELOPE>
  <HEADER>
    <VERSION>1</VERSION>
    <TALLYREQUEST>Export</TALLYREQUEST>
    <TYPE>Collection</TYPE>
    <ID>VoucherTypeList</ID>
  </HEADER>
  <BODY>
    <DESC>
      <STATICVARIABLES>
        <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
        <SVCURRENTCOMPANY>{tally.company}</SVCURRENTCOMPANY>
      </STATICVARIABLES>
      <TDL>
        <TDLMESSAGE>
          <COLLECTION NAME="VoucherTypeList" ISMODIFY="No">
            <TYPE>VoucherType</TYPE>
            <FETCH>NAME, PARENT, NUMBERINGMETHOD, ISINVOICE</FETCH>
          </COLLECTION>
        </TDLMESSAGE>
      </TDL>
    </DESC>
  </BODY>
</ENVELOPE>"""

    resp = tally._send_xml(xml_vtypes, timeout=30)
    root = sanitize_xml(resp)

    vtypes = []
    for vt in root.iter("VOUCHERTYPE"):
        name = (vt.get("NAME") or vt.findtext(".//NAME", "")).strip()
        parent = vt.findtext("PARENT", "").strip()
        is_invoice = vt.findtext("ISINVOICE", "").strip()
        if name:
            vtypes.append({"name": name, "parent": parent, "is_invoice": is_invoice})

    print(f"\n{'Voucher Type':<45} {'Parent':<25} {'Is Invoice?':<12}")
    print("-" * 85)
    for vt in vtypes:
        print(f"{vt['name']:<45} {vt['parent']:<25} {vt['is_invoice']:<12}")
    print(f"\nTotal: {len(vtypes)} voucher types")

    # Save
    out1 = TMP_DIR / "tally_voucher_types.json"
    with open(out1, "w", encoding="utf-8") as f:
        json.dump(vtypes, f, indent=2, ensure_ascii=False)
    print(f"Saved to: {out1}")

    # ── Step 2: Identify sales/invoice types ──
    print("\n\n=== SALES-RELATED VOUCHER TYPES ===")
    sales_types = [vt for vt in vtypes if
                   vt["parent"].lower() in ("sales", "") and "sale" in vt["name"].lower()
                   or "invoice" in vt["name"].lower()
                   or "gst" in vt["name"].lower()
                   or vt["parent"].lower() == "sales"]
    
    for vt in sales_types:
        print(f"  {vt['name']:<45} parent={vt['parent']:<20} invoice={vt['is_invoice']}")

    # ── Step 3: For each sales-related type, count vouchers in FY 2025-26 ──
    print("\n\n=== VOUCHER COUNTS PER TYPE (FY 2025-26) ===")
    from_date = "20250401"
    to_date = "20260331"

    # Check all types that might be sales-related
    types_to_check = set()
    for vt in vtypes:
        if vt["parent"].lower() == "sales" or vt["name"].lower() == "sales":
            types_to_check.add(vt["name"])
        if "invoice" in vt["name"].lower() or "gst" in vt["name"].lower():
            types_to_check.add(vt["name"])

    print(f"Checking these types: {types_to_check}")

    for vtype in sorted(types_to_check):
        xml_count = f"""<ENVELOPE>
  <HEADER>
    <VERSION>1</VERSION>
    <TALLYREQUEST>Export</TALLYREQUEST>
    <TYPE>Collection</TYPE>
    <ID>VoucherCount_{vtype.replace(' ', '_')}</ID>
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
          <COLLECTION NAME="VoucherCount_{vtype.replace(' ', '_')}" ISMODIFY="No">
            <TYPE>Voucher</TYPE>
            <CHILDOF>{vtype}</CHILDOF>
            <FETCH>VOUCHERNUMBER</FETCH>
          </COLLECTION>
        </TDLMESSAGE>
      </TDL>
    </DESC>
  </BODY>
</ENVELOPE>"""
        try:
            resp = tally._send_xml(xml_count, timeout=60)
            root = sanitize_xml(resp)
            count = len(list(root.iter("VOUCHER")))
            print(f"  {vtype:<45} -> {count} vouchers")
        except Exception as e:
            print(f"  {vtype:<45} -> ERROR: {e}")

    # ── Step 4: Fetch sample GST Invoice vouchers to see their structure ──
    # Try common names
    for test_type in ["GST Invoice", "Tax Invoice", "Sales Invoice"]:
        if test_type in types_to_check or any(test_type.lower() in t.lower() for t in types_to_check):
            print(f"\n\n=== SAMPLE VOUCHERS: {test_type} (first 5) ===")
            xml_sample = f"""<ENVELOPE>
  <HEADER>
    <VERSION>1</VERSION>
    <TALLYREQUEST>Export</TALLYREQUEST>
    <TYPE>Collection</TYPE>
    <ID>SampleInvoice</ID>
  </HEADER>
  <BODY>
    <DESC>
      <STATICVARIABLES>
        <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
        <SVCURRENTCOMPANY>{tally.company}</SVCURRENTCOMPANY>
        <SVFROMDATE>{from_date}</SVFROMDATE>
        <SVTODATE>20250415</SVTODATE>
      </STATICVARIABLES>
      <TDL>
        <TDLMESSAGE>
          <COLLECTION NAME="SampleInvoice" ISMODIFY="No">
            <TYPE>Voucher</TYPE>
            <CHILDOF>{test_type}</CHILDOF>
            <FETCH>DATE, VOUCHERNUMBER, VOUCHERTYPENAME, PARTYLEDGERNAME, BASICBUYERNAME, NARRATION, AMOUNT, ALLLEDGERENTRIES.LIST</FETCH>
          </COLLECTION>
        </TDLMESSAGE>
      </TDL>
    </DESC>
  </BODY>
</ENVELOPE>"""
            try:
                resp = tally._send_xml(xml_sample, timeout=60)
                root = sanitize_xml(resp)
                vouchers = list(root.iter("VOUCHER"))
                print(f"  Found {len(vouchers)} vouchers in sample period")
                for v in vouchers[:5]:
                    print(f"\n  Voucher: {v.findtext('VOUCHERNUMBER', '')}")
                    print(f"    Date:       {v.findtext('DATE', '')}")
                    print(f"    Type:       {v.findtext('VOUCHERTYPENAME', '')}")
                    print(f"    Party:      {v.findtext('PARTYLEDGERNAME', '')}")
                    print(f"    BuyerName:  {v.findtext('BASICBUYERNAME', '')}")
                    print(f"    Amount:     {v.findtext('AMOUNT', '')}")
                    print(f"    Narration:  {v.findtext('NARRATION', '')[:80]}")
                    print(f"    Ledger Entries:")
                    for entry in v.iter("ALLLEDGERENTRIES.LIST"):
                        ln = entry.findtext("LEDGERNAME", "")
                        amt = entry.findtext("AMOUNT", "")
                        print(f"      - {ln:<55} {amt:>15}")
            except Exception as e:
                print(f"  ERROR: {e}")


if __name__ == "__main__":
    main()
