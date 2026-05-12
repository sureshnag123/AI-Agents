"""Test the new Import Data envelope format with one voucher."""
import sys
sys.path.insert(0, 'execution')
from excel_to_tally_xml import read_and_validate_excel, convert_to_xml
import requests

TALLY_URL = "http://localhost:9000"

# Read the user's file
rows, errors, warnings = read_and_validate_excel(
    r'.tmp\upload_20260305_073320_upload_20260304_235107_Thota_TEST_March2026_v1.xlsx'
)

# Take just 1 Journal entry to test
test_rows = [r for r in rows if r["voucher_type"].lower() == "journal"][:1]
print(f"Testing with 1 Journal entry: {test_rows[0]['description']}")

xml = convert_to_xml(test_rows, prefix="ENVTEST")

# Show the XML
print("\n=== XML Being Sent ===")
print(xml)
print("=== End XML ===\n")

# Send to Tally
try:
    resp = requests.post(
        TALLY_URL,
        data=xml.encode("utf-8"),
        headers={"Content-Type": "text/xml; charset=utf-8"},
        timeout=30,
    )
    print(f"HTTP Status: {resp.status_code}")
    print(f"Response:\n{resp.text}")
except Exception as e:
    print(f"Error: {e}")
