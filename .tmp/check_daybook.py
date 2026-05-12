"""Check voucher numbers using Day Book report."""
import requests
import re

TALLY_URL = "http://localhost:9000"

# Use DayBook collection
query_xml = """<ENVELOPE>
  <HEADER>
    <TALLYREQUEST>Export Data</TALLYREQUEST>
  </HEADER>
  <BODY>
    <EXPORTDATA>
      <REQUESTDESC>
        <STATICVARIABLES>
          <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
          <SVCURRENTCOMPANY>THOTA HOSPITALITY LLP</SVCURRENTCOMPANY>
          <SVFROMDATE>20260301</SVFROMDATE>
          <SVTODATE>20260301</SVTODATE>
        </STATICVARIABLES>
        <REPORTNAME>Day Book</REPORTNAME>
      </REQUESTDESC>
    </EXPORTDATA>
  </BODY>
</ENVELOPE>"""

resp = requests.post(TALLY_URL, data=query_xml.encode("utf-8"),
                     headers={"Content-Type": "text/xml"}, timeout=30)

text = resp.text

# Look for ENVTEST
if "ENVTEST" in text:
    print("FOUND 'ENVTEST' in Day Book!")
    idx = text.find("ENVTEST")
    print(text[max(0,idx-300):idx+300])
else:
    print("'ENVTEST' NOT in Day Book export.")

# Look for any voucher number patterns
vch_nums = re.findall(r'<VOUCHERNUMBER>(.*?)</VOUCHERNUMBER>', text)
print(f"\nAll VOUCHERNUMBER tags: {len(vch_nums)}")
for v in vch_nums[-20:]:
    print(f"  {v}")

# Also check for the Fuel narration we just uploaded
if "Fuel" in text:
    idx = text.find("Fuel")
    print(f"\nFound 'Fuel' context:")
    print(text[max(0,idx-400):idx+200])

# Save full response for manual inspection
with open(r".tmp\daybook_response.txt", "w", encoding="utf-8") as f:
    f.write(text[:50000])
print(f"\nFull response saved to .tmp/daybook_response.txt ({len(text)} chars)")
