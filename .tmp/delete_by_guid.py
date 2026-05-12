"""Extract GUIDs for duplicate March 2026 vouchers and delete them."""
import re
import requests

# Read the all_vouchers.xml
with open(r"D:\RaviGowda\Thota Documents\AGENT\.tmp\all_vouchers.xml", "r", encoding="utf-8") as f:
    content = f.read()

# Split into voucher blocks
blocks = re.split(r'(?=<VOUCHER\s)', content)
march_vouchers = []

for block in blocks:
    date_match = re.search(r'<DATE[^>]*>202603(\d{2})</DATE>', block)
    if date_match:
        vch_num = re.search(r'<VOUCHERNUMBER>([^<]+)</VOUCHERNUMBER>', block)
        vch_type = re.search(r'VCHTYPE="([^"]+)"', block)
        guid = re.search(r'<GUID>([^<]+)</GUID>', block)
        remoteid = re.search(r'REMOTEID="([^"]+)"', block)
        
        march_vouchers.append({
            'date': '202603' + date_match.group(1),
            'num': vch_num.group(1) if vch_num else '?',
            'type': vch_type.group(1) if vch_type else '?',
            'guid': guid.group(1) if guid else None,
            'remoteid': remoteid.group(1) if remoteid else None,
        })

print(f"March 2026 vouchers: {len(march_vouchers)}")
for v in march_vouchers:
    print(f"  {v['date']} | {v['type']:10s} | #{v['num']:6s} | GUID: {v['guid']}")

# Identify duplicates: the higher-numbered ones are duplicates
# Journal: 1814-1818 = originals, 1819-1823 = duplicates
# Purchase: 886-890 = originals, 891-895 = duplicates
duplicates = []
for v in march_vouchers:
    num = int(v['num'])
    if (v['type'] == 'Journal' and num >= 1819) or (v['type'] == 'Purchase' and num >= 891):
        duplicates.append(v)

print(f"\nDuplicates to delete: {len(duplicates)}")
for v in duplicates:
    print(f"  {v['type']} #{v['num']} GUID={v['guid']}")

# Now delete using GUID
BASE_URL = "http://localhost:9000"
COMPANY = "THOTA HOSPITALITY LLP"

for v in duplicates:
    guid = v['guid']
    remoteid = v['remoteid']
    vtype = v['type']
    vnum = v['num']
    
    delete_xml = f"""<ENVELOPE>
  <HEADER>
    <VERSION>1</VERSION>
    <TALLYREQUEST>Import</TALLYREQUEST>
    <TYPE>Data</TYPE>
    <ID>All Masters and Vouchers</ID>
  </HEADER>
  <BODY>
    <DESC>
      <STATICVARIABLES>
        <SVCURRENTCOMPANY>{COMPANY}</SVCURRENTCOMPANY>
      </STATICVARIABLES>
    </DESC>
    <DATA>
      <TALLYMESSAGE>
        <VOUCHER REMOTEID="{remoteid}" VCHTYPE="{vtype}" ACTION="Delete">
          <GUID>{guid}</GUID>
          <VOUCHERTYPENAME>{vtype}</VOUCHERTYPENAME>
          <VOUCHERNUMBER>{vnum}</VOUCHERNUMBER>
        </VOUCHER>
      </TALLYMESSAGE>
    </DATA>
  </BODY>
</ENVELOPE>"""

    print(f"\n  Deleting {vtype} #{vnum}...", end=" ", flush=True)
    try:
        resp = requests.post(BASE_URL, data=delete_xml.encode("utf-8"),
                           headers={"Content-Type": "text/xml; charset=utf-8"}, timeout=30)
        if resp.status_code == 200:
            deleted = re.search(r'<DELETED>(\d+)</DELETED>', resp.text)
            errors = re.search(r'<ERRORS>(\d+)</ERRORS>', resp.text)
            d = deleted.group(1) if deleted else "0"
            e = errors.group(1) if errors else "0"
            print(f"Deleted={d}, Errors={e}")
            if d == "0" and e == "0":
                print(f"    Full response: {resp.text[:300]}")
        else:
            print(f"HTTP {resp.status_code}")
    except requests.Timeout:
        print("TIMEOUT")
    except Exception as ex:
        print(f"ERROR: {ex}")
