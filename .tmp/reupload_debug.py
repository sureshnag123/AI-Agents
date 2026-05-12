"""Re-upload XML to Tally with detailed response logging."""
import requests
import xml.etree.ElementTree as ET

BASE_URL = "http://localhost:9000"

xml_path = r"D:\RaviGowda\Thota Documents\AGENT\.tmp\tally_import_20260304_134359.xml"

with open(xml_path, "r", encoding="utf-8") as f:
    xml_content = f.read()

print("Uploading to Tally...")
resp = requests.post(
    BASE_URL,
    data=xml_content.encode("utf-8"),
    headers={"Content-Type": "text/xml; charset=utf-8"},
    timeout=60,
)

print("Status:", resp.status_code)
print("\n=== FULL RAW RESPONSE ===")
print(resp.text)
print("=== END RESPONSE ===")

# Save response
with open(r"D:\RaviGowda\Thota Documents\AGENT\.tmp\tally_import_response.xml", "w", encoding="utf-8") as f:
    f.write(resp.text)

# Parse response XML
try:
    root = ET.fromstring(resp.text)
    print("\n=== ALL TAGS AND VALUES ===")
    for elem in root.iter():
        if elem.text and elem.text.strip():
            print(f"  {elem.tag}: {elem.text.strip()}")
except ET.ParseError as e:
    print("Parse error:", e)
