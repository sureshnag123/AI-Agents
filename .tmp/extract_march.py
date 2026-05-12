"""Extract March 2026 voucher details from saved response."""
import re

with open(r"D:\RaviGowda\Thota Documents\AGENT\.tmp\all_vouchers.xml", "r", encoding="utf-8") as f:
    content = f.read()

# Find all VOUCHER blocks that contain March 2026 dates
# Split by VOUCHER tags
voucher_blocks = re.split(r'(?=<VOUCHER\s)', content)

march_vouchers = []
for block in voucher_blocks:
    if re.search(r'<DATE[^>]*>202603\d{2}</DATE>', block):
        march_vouchers.append(block)

print(f"Found {len(march_vouchers)} voucher blocks with March 2026 dates\n")

for i, block in enumerate(march_vouchers):
    date = re.search(r'<DATE[^>]*>(\d{8})</DATE>', block)
    vch_num = re.search(r'<VOUCHERNUMBER>([^<]+)</VOUCHERNUMBER>', block)
    vch_type = re.search(r'VCHTYPE="([^"]+)"', block)
    vch_type2 = re.search(r'<VOUCHERTYPENAME>([^<]+)</VOUCHERTYPENAME>', block)
    narration = re.search(r'<NARRATION>([^<]*)</NARRATION>', block)
    
    d = date.group(1) if date else "?"
    n = vch_num.group(1) if vch_num else "?"
    t = vch_type.group(1) if vch_type else (vch_type2.group(1) if vch_type2 else "?")
    nr = narration.group(1)[:80] if narration else "(no narration in this export)"
    
    print(f"  {i+1}. Date: {d} | Type: {t:10s} | VchNo: {n:10s} | {nr}")
