"""Query Tally to find the uploaded test vouchers."""
import requests
import xml.etree.ElementTree as ET

BASE_URL = "http://localhost:9000"
COMPANY = "THOTA HOSPITALITY LLP"

query_xml = """<ENVELOPE>
  <HEADER>
    <VERSION>1</VERSION>
    <TALLYREQUEST>Export</TALLYREQUEST>
    <TYPE>Data</TYPE>
    <ID>Daybook</ID>
  </HEADER>
  <BODY>
    <DESC>
      <STATICVARIABLES>
        <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
        <SVCURRENTCOMPANY>THOTA HOSPITALITY LLP</SVCURRENTCOMPANY>
        <SVFROMDATE>20260301</SVFROMDATE>
        <SVTODATE>20260305</SVTODATE>
      </STATICVARIABLES>
    </DESC>
  </BODY>
</ENVELOPE>"""

print("Querying Tally for March 1-5, 2026 vouchers...")
resp = requests.post(BASE_URL, data=query_xml, headers={"Content-Type": "text/xml"}, timeout=30)
print("Status:", resp.status_code)

if resp.status_code == 200:
    with open(r"D:\RaviGowda\Thota Documents\AGENT\.tmp\tally_query_response.xml", "w", encoding="utf-8") as f:
        f.write(resp.text)
    print("Raw response saved")
    print("Response length:", len(resp.text), "chars")
    
    try:
        root = ET.fromstring(resp.text)
        vouchers = list(root.iter("VOUCHER"))
        print("\nFound", len(vouchers), "voucher(s) in March 1-5, 2026")
        
        for v in vouchers[:20]:
            vch_num = v.findtext("VOUCHERNUMBER", "N/A")
            vch_type = v.get("VCHTYPE", "")
            if not vch_type:
                vch_type = v.findtext("VOUCHERTYPENAME", "N/A")
            date_raw = v.findtext("DATE", "N/A")
            narration = (v.findtext("NARRATION") or "")[:80]
            print("  {} | {:10s} | {} | {}".format(date_raw, vch_type, vch_num, narration))
    except ET.ParseError as e:
        print("Parse error:", e)
        print("First 1000 chars:", resp.text[:1000])
else:
    print("Error:", resp.text[:500])
