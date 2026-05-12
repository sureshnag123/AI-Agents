#!/usr/bin/env python3
"""
Repeat Customer Report from Tally Sales Register (FY 2025-26)

Fetches only Sales vouchers from Tally, groups by party ledger name (customer),
and identifies repeat customers (2+ sales invoices).
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


def sanitize_and_parse_xml(xml_text: str) -> ET.Element:
    """Sanitize Tally XML that may have unbound namespace prefixes."""
    # Remove invalid XML character references
    xml_text = re.sub(r'&#x([0-8BbCcEeFf]);', '', xml_text)
    xml_text = re.sub(r'&#x1[0-9A-Fa-f];', '', xml_text)
    xml_text = re.sub(r'&#([0-8]|1[0-1]|1[4-9]|2[0-9]|3[01]);', '', xml_text)
    xml_text = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', xml_text)
    # Remove namespace prefixes like UDF:XXXXX -> XXXXX
    xml_text = re.sub(r'<(/?)(\w+):', r'<\1\2_', xml_text)
    # Remove namespace attributes
    xml_text = re.sub(r'\s+xmlns:\w+="[^"]*"', '', xml_text)
    try:
        return ET.fromstring(xml_text)
    except ET.ParseError as e:
        raise ValueError(f"XML parse error: {e}\n{xml_text[:1000]}")


def main():
    tally = TallyConnector()

    # Test connection first
    try:
        info = tally.test_connection()
        print(f"Connected to Tally at {info['url']}")
        print(f"Company: {info['target_company']}")
    except Exception as e:
        print(f"Connection failed: {e}")
        sys.exit(1)

    # FY 2025-26: April 1, 2025 to March 31, 2026
    from_date = "20250401"
    to_date = "20260331"

    print(f"\nFetching Sales vouchers from {from_date} to {to_date} ...")

    # Build Tally XML request for Sales vouchers
    xml_request = f"""<ENVELOPE>
  <HEADER>
    <VERSION>1</VERSION>
    <TALLYREQUEST>Export</TALLYREQUEST>
    <TYPE>Collection</TYPE>
    <ID>Sales Voucher Collection</ID>
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
          <COLLECTION NAME="Sales Voucher Collection" ISMODIFY="No">
            <TYPE>Voucher</TYPE>
            <CHILDOF>Sales</CHILDOF>
            <FETCH>DATE, NARRATION, VOUCHERNUMBER, PARTYLEDGERNAME, AMOUNT, REFERENCE, BASICBUYERNAME</FETCH>
          </COLLECTION>
        </TDLMESSAGE>
      </TDL>
    </DESC>
  </BODY>
</ENVELOPE>"""

    resp_text = tally._send_xml(xml_request, timeout=120)
    root = sanitize_and_parse_xml(resp_text)

    sales_vouchers = []
    for v in root.iter("VOUCHER"):
        sales_vouchers.append({
            "date": v.findtext("DATE", ""),
            "number": v.findtext("VOUCHERNUMBER", ""),
            "party": v.findtext("PARTYLEDGERNAME", ""),
            "buyer": v.findtext("BASICBUYERNAME", ""),
            "amount": v.findtext("AMOUNT", "0"),
            "narration": v.findtext("NARRATION", ""),
            "reference": v.findtext("REFERENCE", ""),
        })

    print(f"Total Sales vouchers fetched: {len(sales_vouchers)}")

    if not sales_vouchers:
        print("No sales vouchers found!")
        sys.exit(0)

    # Group by party (customer) name
    customer_data = defaultdict(lambda: {"invoices": [], "total_amount": 0.0})

    for v in sales_vouchers:
        party = (v.get("party") or "").strip()
        if not party:
            party = "(No Party Name)"

        # Parse amount - Tally amounts for Sales are typically negative (credit)
        try:
            amount = abs(float(v.get("amount", "0").replace(",", "")))
        except (ValueError, AttributeError):
            amount = 0.0

        customer_data[party]["invoices"].append({
            "number": v.get("number", ""),
            "date": v.get("date", ""),
            "amount": amount,
            "reference": v.get("reference", ""),
        })
        customer_data[party]["total_amount"] += amount

    # Separate repeat vs one-time
    repeat_customers = {}
    one_time_customers = {}

    for name, data in customer_data.items():
        count = len(data["invoices"])
        if count >= 2:
            repeat_customers[name] = data
        else:
            one_time_customers[name] = data

    # Sort repeat customers by invoice count descending
    sorted_repeat = sorted(
        repeat_customers.items(),
        key=lambda x: len(x[1]["invoices"]),
        reverse=True
    )

    # Print report
    print(f"\n{'='*100}")
    print(f"REPEAT CUSTOMERS FROM TALLY SALES REGISTER - FY 2025-26")
    print(f"{'='*100}")
    print(f"Total Sales Vouchers:    {len(sales_vouchers)}")
    print(f"Unique Customers:        {len(customer_data)}")
    print(f"Repeat Customers (2+):   {len(repeat_customers)}")
    print(f"One-Time Customers:      {len(one_time_customers)}")
    print()

    print(f"{'#':<5} {'Customer Name (Party Ledger)':<60} {'Inv Count':<10} {'Total Amount (Rs)':>18}")
    print("-" * 98)

    total_repeat_amount = 0.0
    total_repeat_invoices = 0

    for i, (name, data) in enumerate(sorted_repeat, 1):
        inv_count = len(data["invoices"])
        total_amt = data["total_amount"]
        total_repeat_amount += total_amt
        total_repeat_invoices += inv_count
        display_name = name[:59]
        print(f"{i:<5} {display_name:<60} {inv_count:<10} {total_amt:>18,.2f}")

    print("-" * 98)
    print(f"{'TOTAL':<5} {'':<60} {total_repeat_invoices:<10} {total_repeat_amount:>18,.2f}")

    # Save detailed JSON
    output = {
        "summary": {
            "report": "Repeat Customers from Tally Sales Register",
            "period": "FY 2025-26 (Apr 2025 - Mar 2026)",
            "voucher_type": "Sales",
            "total_sales_vouchers": len(sales_vouchers),
            "unique_customers": len(customer_data),
            "repeat_customers": len(repeat_customers),
            "one_time_customers": len(one_time_customers),
            "total_repeat_invoices": total_repeat_invoices,
            "total_repeat_amount": round(total_repeat_amount, 2),
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
    }

    TMP_DIR.mkdir(parents=True, exist_ok=True)
    out_file = TMP_DIR / "repeat_customers_sales_register_fy2526.json"
    with open(out_file, "w", encoding="utf-8") as f:
        json.dump(output, f, indent=2, ensure_ascii=False)
    print(f"\nDetailed report saved to: {out_file}")


if __name__ == "__main__":
    main()
