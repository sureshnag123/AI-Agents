"""Check if Tally stored our custom voucher number."""
import requests

TALLY_URL = "http://localhost:9000"

# Query for Journal vouchers on 01-Mar-2026
query_xml = """<ENVELOPE>
  <HEADER>
    <TALLYREQUEST>Export Data</TALLYREQUEST>
  </HEADER>
  <BODY>
    <EXPORTDATA>
      <REQUESTDESC>
        <REPORTNAME>Vouchers</REPORTNAME>
        <STATICVARIABLES>
          <SVCURRENTCOMPANY>THOTA HOSPITALITY LLP</SVCURRENTCOMPANY>
          <SVFROMDATE>20260301</SVFROMDATE>
          <SVTODATE>20260301</SVTODATE>
          <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
        </STATICVARIABLES>
      </REQUESTDESC>
    </EXPORTDATA>
  </BODY>
</ENVELOPE>"""

resp = requests.post(TALLY_URL, data=query_xml.encode("utf-8"),
                     headers={"Content-Type": "text/xml"}, timeout=30)

# Search for ENVTEST in the response
text = resp.text
if "ENVTEST" in text:
    # Extract the voucher number context
    idx = text.find("ENVTEST")
    start = max(0, idx - 200)
    end = min(len(text), idx + 200)
    print("FOUND 'ENVTEST' in Tally response!")
    print(text[start:end])
else:
    print("'ENVTEST' NOT found in Tally response.")
    # Show voucher numbers found
    import re
    vch_nums = re.findall(r'<VOUCHERNUMBER>(.*?)</VOUCHERNUMBER>', text)
    print(f"\nVoucher numbers found on 01-Mar-2026: {len(vch_nums)}")
    for v in vch_nums[-15:]:
        print(f"  {v}")
