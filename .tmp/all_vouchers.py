"""Query ALL vouchers from Tally without date restriction."""
import requests
import re

BASE_URL = "http://localhost:9000"

# Query all vouchers via collection
query = """<ENVELOPE>
  <HEADER>
    <VERSION>1</VERSION>
    <TALLYREQUEST>Export</TALLYREQUEST>
    <TYPE>Collection</TYPE>
    <ID>AllVch</ID>
  </HEADER>
  <BODY>
    <DESC>
      <STATICVARIABLES>
        <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
        <SVCURRENTCOMPANY>THOTA HOSPITALITY LLP</SVCURRENTCOMPANY>
      </STATICVARIABLES>
      <TDL>
        <TDLMESSAGE>
          <COLLECTION NAME="AllVch">
            <TYPE>Voucher</TYPE>
            <FETCH>VOUCHERNUMBER, DATE, VOUCHERTYPENAME</FETCH>
          </COLLECTION>
        </TDLMESSAGE>
      </TDL>
    </DESC>
  </BODY>
</ENVELOPE>"""

print("Querying all vouchers (no date filter)...")
resp = requests.post(BASE_URL, data=query, headers={"Content-Type": "text/xml"}, timeout=120)
print(f"Status: {resp.status_code}, Length: {len(resp.text)} chars")

# Save
with open(r"D:\RaviGowda\Thota Documents\AGENT\.tmp\all_vouchers.xml", "w", encoding="utf-8") as f:
    f.write(resp.text)

# Search for EXL
exl = re.findall(r'EXL[^<]*', resp.text)
print(f"\nEXL matches: {len(exl)}")
for e in exl:
    print(f"  {e}")

# All voucher numbers
vch_nums = re.findall(r'<VOUCHERNUMBER>([^<]+)</VOUCHERNUMBER>', resp.text)
print(f"\nTotal voucher numbers: {len(vch_nums)}")
for v in vch_nums[:40]:
    print(f"  {v}")

# All dates
dates = re.findall(r'<DATE>(\d{8})</DATE>', resp.text)
if dates:
    unique = sorted(set(dates))
    print(f"\nUnique dates ({len(unique)}): {unique}")

# All voucher types
vtypes = re.findall(r'<VOUCHERTYPENAME>([^<]+)</VOUCHERTYPENAME>', resp.text)
if vtypes:
    from collections import Counter
    print(f"\nVoucher types:")
    for t, c in Counter(vtypes).most_common():
        print(f"  {t}: {c}")

# Check first 2000 chars of response
print(f"\nFirst 2000 chars:")
print(resp.text[:2000])
