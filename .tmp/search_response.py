"""Search raw Tally response for our uploaded vouchers."""
import re

with open(r"D:\RaviGowda\Thota Documents\AGENT\.tmp\tally_query_response.xml", "r", encoding="utf-8") as f:
    content = f.read()

print("=== Searching for EXL prefix ===")
exl_matches = re.findall(r'EXL[^<]*', content)
if exl_matches:
    for m in exl_matches:
        print(" ", m)
else:
    print("  No 'EXL' found in response")

print("\n=== Searching for VOUCHERNUMBER tags ===")
vch_numbers = re.findall(r'<VOUCHERNUMBER>([^<]+)</VOUCHERNUMBER>', content)
print(f"  Found {len(vch_numbers)} voucher numbers")
for vn in vch_numbers[:30]:
    print(f"  {vn}")

print("\n=== Searching for our narrations ===")
narrations = re.findall(r'<NARRATION>([^<]*(?:Kitchen|Staff|Linen|Water|Pest|Packaging|Maintenance|Restaurant|Grocery|Laundry)[^<]*)</NARRATION>', content, re.IGNORECASE)
if narrations:
    for n in narrations[:10]:
        print(f"  {n[:100]}")
else:
    print("  No matching narrations found")

print("\n=== Voucher types in response ===")
vch_types = re.findall(r'VCHTYPE="([^"]+)"', content)
from collections import Counter
type_counts = Counter(vch_types)
for t, c in type_counts.most_common():
    print(f"  {t}: {c}")

print("\n=== Date range in vouchers ===")
dates = re.findall(r'<DATE>(\d{8})</DATE>', content)
unique_dates = sorted(set(dates))
print(f"  Dates found: {unique_dates}")
