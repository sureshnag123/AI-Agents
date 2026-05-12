"""Query Tally for Journal/Purchase vouchers in March 2026."""
import requests
import re

BASE_URL = "http://localhost:9000"

# Try querying for Journal vouchers
for vtype in ["Journal", "Purchase"]:
    query = f"""<ENVELOPE>
  <HEADER>
    <VERSION>1</VERSION>
    <TALLYREQUEST>Export</TALLYREQUEST>
    <TYPE>Data</TYPE>
    <ID>{vtype} Register</ID>
  </HEADER>
  <BODY>
    <DESC>
      <STATICVARIABLES>
        <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
        <SVCURRENTCOMPANY>THOTA HOSPITALITY LLP</SVCURRENTCOMPANY>
        <SVFROMDATE>20260301</SVFROMDATE>
        <SVTODATE>20260331</SVTODATE>
      </STATICVARIABLES>
    </DESC>
  </BODY>
</ENVELOPE>"""
    
    print(f"\n=== Querying {vtype} Register (March 2026) ===")
    resp = requests.post(BASE_URL, data=query, headers={"Content-Type": "text/xml"}, timeout=30)
    print(f"Status: {resp.status_code}, Length: {len(resp.text)} chars")
    
    vch_numbers = re.findall(r'<VOUCHERNUMBER>([^<]+)</VOUCHERNUMBER>', resp.text)
    print(f"Voucher numbers found: {len(vch_numbers)}")
    for vn in vch_numbers:
        print(f"  {vn}")
    
    dates = re.findall(r'<DATE>(\d{8})</DATE>', resp.text)
    if dates:
        print(f"Dates: {sorted(set(dates))}")

# Also try a wildcard search for EXL
print("\n=== Searching via TDL report for EXL vouchers ===")
query2 = """<ENVELOPE>
  <HEADER>
    <VERSION>1</VERSION>
    <TALLYREQUEST>Export</TALLYREQUEST>
    <TYPE>Collection</TYPE>
    <ID>EXLVouchers</ID>
  </HEADER>
  <BODY>
    <DESC>
      <STATICVARIABLES>
        <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
        <SVCURRENTCOMPANY>THOTA HOSPITALITY LLP</SVCURRENTCOMPANY>
        <SVFROMDATE>20260101</SVFROMDATE>
        <SVTODATE>20260331</SVTODATE>
      </STATICVARIABLES>
      <TDL>
        <TDLMESSAGE>
          <COLLECTION NAME="EXLVouchers">
            <TYPE>Voucher</TYPE>
            <FETCH>VOUCHERNUMBER, DATE, VOUCHERTYPENAME, NARRATION, AMOUNT</FETCH>
          </COLLECTION>
        </TDLMESSAGE>
      </TDL>
    </DESC>
  </BODY>
</ENVELOPE>"""

resp2 = requests.post(BASE_URL, data=query2, headers={"Content-Type": "text/xml"}, timeout=30)
print(f"Status: {resp2.status_code}, Length: {len(resp2.text)} chars")

vch_numbers2 = re.findall(r'<VOUCHERNUMBER>([^<]+)</VOUCHERNUMBER>', resp2.text)
print(f"Total vouchers found: {len(vch_numbers2)}")
exl_vchs = [v for v in vch_numbers2 if 'EXL' in v]
print(f"EXL vouchers: {len(exl_vchs)}")
for v in exl_vchs:
    print(f"  {v}")

if not exl_vchs:
    # Show all voucher numbers to see what format Tally used
    print("All voucher numbers (first 30):")
    for v in vch_numbers2[:30]:
        print(f"  {v}")
