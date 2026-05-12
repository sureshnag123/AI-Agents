"""Test the complete fix: new envelope + prefix in narration + reference."""
import sys, re
sys.path.insert(0, 'execution')
from excel_to_tally_xml import read_and_validate_excel, convert_to_xml

rows, errors, warnings = read_and_validate_excel(
    r'.tmp\upload_20260305_073320_upload_20260304_235107_Thota_TEST_March2026_v1.xlsx'
)
print(f"Valid: {len(rows)}, Errors: {len(errors)}")

xml = convert_to_xml(rows, prefix="FIX")

# Show sample Journal voucher
journals = re.findall(r'<VOUCHER VCHTYPE="Journal".*?</VOUCHER>', xml, re.DOTALL)
if journals:
    print("\n=== Sample Journal Voucher ===")
    print(journals[0])

# Show sample Purchase voucher
purchases = re.findall(r'<VOUCHER VCHTYPE="Purchase".*?</VOUCHER>', xml, re.DOTALL)
if purchases:
    print("\n=== Sample Purchase Voucher ===")
    print(purchases[0])

# Show envelope header
print("\n=== Envelope Header ===")
header_end = xml.find("<VOUCHER")
print(xml[:header_end])
