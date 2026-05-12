#!/usr/bin/env python3
"""Export list of duplicate March 2026 Purchase vouchers to a text file."""
import sys
from pathlib import Path
from collections import defaultdict

SCRIPT_DIR = Path(__file__).parent.resolve()
PROJECT_ROOT = SCRIPT_DIR.parent
TMP_DIR = PROJECT_ROOT / ".tmp"
TMP_DIR.mkdir(parents=True, exist_ok=True)

sys.path.insert(0, str(SCRIPT_DIR))
from dotenv import load_dotenv
load_dotenv(PROJECT_ROOT / ".env")
from tally_connector import TallyConnector

t = TallyConnector()
vouchers = t.get_vouchers_with_guid('20260301', '20260331', 'Purchase')
march = [v for v in vouchers
         if '20260301' <= v['date'] <= '20260331'
         and v['number'].startswith('BILL/2026/03/')]

groups = defaultdict(list)
for v in march:
    groups[v['number']].append(v)

duplicates = {k: v for k, v in groups.items() if len(v) > 1}
extra_count = sum(len(v) - 1 for v in duplicates.values())

lines = []
lines.append('TALLY DUPLICATE PURCHASE VOUCHERS - MARCH 2026')
lines.append('=' * 60)
lines.append(f'Duplicated voucher numbers : {len(duplicates)}')
lines.append(f'Extra entries to delete    : {extra_count}')
lines.append('')
lines.append('Instructions:')
lines.append('  1. Open Tally Prime → Gateway of Tally → Day Book')
lines.append('  2. Set date range: 01-Mar-2026 to 31-Mar-2026')
lines.append('  3. Filter by Voucher Type: Purchase')
lines.append('  4. For each "DELETE" row below, find that voucher in the Day Book')
lines.append('     (The duplicate will be the one with a HIGHER date/timestamp or')
lines.append('      matching the REMOTEID shown)')
lines.append('  5. Press Alt+D (or right-click → Delete) to delete it')
lines.append('  NOTE: Go to Company → Alter → Enable "Allow Deletion of Vouchers"')
lines.append('        if deletion is not permitted.')
lines.append('')
lines.append(f'{"Voucher No":<30} {"Date":<12} {"Party":<42} {"Amount":>12}  Action')
lines.append('-' * 115)

for num in sorted(duplicates.keys()):
    entries = duplicates[num]
    for i, v in enumerate(entries):
        d = v['date']
        date_fmt = f"{d[:4]}-{d[4:6]}-{d[6:]}"
        tag = 'KEEP  ' if i == 0 else 'DELETE'
        lines.append(f'{tag}  {v["number"]:<30} {date_fmt:<12} {v["party"]:<42} '
                     f'{v["amount"]:>12}  (REMOTEID: ...{v["guid"][-8:]})')
    lines.append('')

report_file = TMP_DIR / "march_duplicates_to_delete.txt"
with open(report_file, 'w', encoding='utf-8') as f:
    f.write('\n'.join(lines))

for line in lines[:30]:
    print(line.encode('ascii', errors='replace').decode('ascii'))
print(f'\n[...{len(duplicates) - 5} more entries...]')
print(f'\nFull report saved to: {report_file}')
