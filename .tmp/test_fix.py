import sys, re
sys.path.insert(0, 'execution')
from excel_to_tally_xml import read_and_validate_excel, convert_to_xml

rows, errors, warnings = read_and_validate_excel(
    r'.tmp\upload_20260305_073320_upload_20260304_235107_Thota_TEST_March2026_v1.xlsx'
)
print(f"Valid: {len(rows)}, Errors: {len(errors)}, Warnings: {len(warnings)}")
for e in errors:
    print(f"  ERR: {e}")

purchase_rows = [r for r in rows if r["voucher_type"].lower() == "purchase"]
print(f"\nPurchase entries: {len(purchase_rows)}")
for r in purchase_rows:
    print(f"  Row {r['row_num']}: Ledger={r['expense_ledger']} | Vendor={r['vendor_name']} | Amt={r['amount']}")

xml = convert_to_xml(rows, prefix="FIX")
purchases = re.findall(r'<VOUCHER VCHTYPE="Purchase".*?</VOUCHER>', xml, re.DOTALL)
if purchases:
    print(f"\n--- Sample Purchase Voucher XML ---")
    print(purchases[0])
else:
    print("\nNo Purchase vouchers found in XML")
