"""Delete duplicate March 2026 vouchers from Tally.

The second batch (Journal 1819-1823, Purchase 891-895) are duplicates.
"""
import requests

BASE_URL = "http://localhost:9000"
COMPANY = "THOTA HOSPITALITY LLP"

# The duplicates are from the second upload:
# Journal: 1819, 1820, 1821, 1822, 1823
# Purchase: 891, 892, 893, 894, 895
duplicates = [
    ("Journal", "1819"), ("Journal", "1820"), ("Journal", "1821"),
    ("Journal", "1822"), ("Journal", "1823"),
    ("Purchase", "891"), ("Purchase", "892"), ("Purchase", "893"),
    ("Purchase", "894"), ("Purchase", "895"),
]

print("Deleting 10 duplicate vouchers from Tally...\n")

success_count = 0
fail_count = 0

for vtype, vnum in duplicates:
    # Build delete XML using REMOTEID approach
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
        <VOUCHER VCHTYPE="{vtype}" ACTION="Delete" VCHACTION="Delete">
          <VOUCHERTYPENAME>{vtype}</VOUCHERTYPENAME>
          <VOUCHERNUMBER>{vnum}</VOUCHERNUMBER>
        </VOUCHER>
      </TALLYMESSAGE>
    </DATA>
  </BODY>
</ENVELOPE>"""

    try:
        resp = requests.post(BASE_URL, data=delete_xml.encode("utf-8"),
                           headers={"Content-Type": "text/xml; charset=utf-8"}, timeout=15)
        if resp.status_code == 200:
            if "DELETED" in resp.text.upper():
                import re
                deleted = re.search(r'<DELETED>(\d+)</DELETED>', resp.text)
                del_count = deleted.group(1) if deleted else "?"
                print(f"  Deleted {vtype} #{vnum} (deleted: {del_count})")
                success_count += 1
            elif "ERROR" in resp.text.upper() or "LINEERROR" in resp.text.upper():
                print(f"  FAILED {vtype} #{vnum}: {resp.text[:200]}")
                fail_count += 1
            else:
                print(f"  {vtype} #{vnum}: Response unclear - {resp.text[:200]}")
                fail_count += 1
        else:
            print(f"  FAILED {vtype} #{vnum}: HTTP {resp.status_code}")
            fail_count += 1
    except Exception as e:
        print(f"  ERROR {vtype} #{vnum}: {e}")
        fail_count += 1

print(f"\nDone: {success_count} deleted, {fail_count} failed")
