#!/usr/bin/env python3
"""
Delete Duplicate Vouchers from Tally

Fetches vouchers for a date range, identifies duplicate entries
(same VOUCHERNUMBER appearing more than once), and deletes the
extras — keeping only the first occurrence of each voucher number.

Usage:
    # Preview duplicates without deleting
    python delete_duplicate_vouchers.py --type Purchase --from 2026-03-01 --to 2026-03-31

    # Actually delete the duplicates
    python delete_duplicate_vouchers.py --type Purchase --from 2026-03-01 --to 2026-03-31 --execute
"""

import os
import sys
import argparse
import logging
from pathlib import Path
from collections import defaultdict

from dotenv import load_dotenv

SCRIPT_DIR = Path(__file__).parent.resolve()
PROJECT_ROOT = SCRIPT_DIR.parent
TMP_DIR = PROJECT_ROOT / ".tmp"
TMP_DIR.mkdir(parents=True, exist_ok=True)

load_dotenv(PROJECT_ROOT / ".env")
sys.path.insert(0, str(SCRIPT_DIR))

from tally_connector import TallyConnector

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler()],
)
log = logging.getLogger("delete_duplicates")


def find_and_delete_duplicates(
    tally: TallyConnector,
    voucher_type: str,
    from_date: str,   # YYYY-MM-DD
    to_date: str,     # YYYY-MM-DD
    execute: bool = False,
    number_prefix: str = "",
):
    tally_from = from_date.replace("-", "")
    tally_to = to_date.replace("-", "")

    log.info(f"Fetching '{voucher_type}' vouchers from Tally ({from_date} to {to_date})...")
    all_vouchers = tally.get_vouchers_with_guid(tally_from, tally_to, voucher_type)
    log.info(f"  Total vouchers fetched from Tally: {len(all_vouchers)}")

    # Tally's TDL Collection date filter is unreliable — apply date range in Python
    vouchers = [
        v for v in all_vouchers
        if tally_from <= v["date"] <= tally_to
    ]
    log.info(f"  After date filter ({from_date} to {to_date}): {len(vouchers)}")

    # Optionally restrict to vouchers whose number starts with a given prefix
    # (e.g. "BILL/2026/03/" to only target Odoo-synced bills)
    if number_prefix:
        vouchers = [v for v in vouchers if v["number"].startswith(number_prefix)]
        log.info(f"  After prefix filter '{number_prefix}': {len(vouchers)}")

    # Group by voucher number — preserving insertion order so index 0 = first occurrence
    groups = defaultdict(list)
    for v in vouchers:
        num = v["number"]
        if num:
            groups[num].append(v)

    duplicates = {num: entries for num, entries in groups.items() if len(entries) > 1}

    if not duplicates:
        log.info("No duplicate vouchers found. Nothing to delete.")
        return

    total_extra = sum(len(e) - 1 for e in duplicates.values())
    log.info(f"\nFound {len(duplicates)} voucher number(s) with duplicates "
             f"({total_extra} extra entries to remove):\n")

    header = f"  {'Voucher Number':<30} {'Date':<10} {'Party':<35} {'Amount':>14}  {'GUID'}"
    print(header)
    print("  " + "-" * (len(header) - 2))

    to_delete = []
    for num, entries in sorted(duplicates.items()):
        for i, v in enumerate(entries):
            tag = "  KEEP  " if i == 0 else "  DELETE"
            print(f"{tag}  {v['number']:<30} {v['date']:<10} {v['party']:<35} "
                  f"{v['amount']:>14}  {v['guid']}")
            if i > 0:
                to_delete.append(v)
        print()

    print(f"  Total to delete: {len(to_delete)}")

    if not execute:
        print("\n  [DRY RUN] No changes made. Re-run with --execute to delete.")
        return

    # Confirm before deleting
    answer = input(f"\nDelete {len(to_delete)} duplicate voucher(s) from Tally? [y/N]: ").strip().lower()
    if answer != "y":
        log.info("Aborted by user.")
        return

    deleted = 0
    errors = 0
    for v in to_delete:
        if not v["guid"] or not v["vchkey"]:
            log.error(f"  [SKIP] {v['number']}: missing GUID/VCHKEY, cannot delete")
            errors += 1
            continue
        result = tally.delete_voucher(
            guid=v["guid"],
            vchkey=v["vchkey"],
            date=v["date"],
            voucher_type=voucher_type,
            voucher_number=v["number"],
            party=v["party"],
        )
        if result["deleted"] > 0:
            log.info(f"  [DELETED] {v['number']} (GUID: {v['guid']})")
            deleted += 1
        else:
            err = result.get("error_message") or f"deleted={result['deleted']}, errors={result['errors']}"
            log.error(f"  [FAIL] {v['number']}: {err}")
            errors += 1

    log.info(f"\nDone — {deleted} deleted, {errors} errors.")


def main():
    parser = argparse.ArgumentParser(
        description="Delete duplicate vouchers from Tally Prime",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Dry run — preview only
  python delete_duplicate_vouchers.py --type Purchase --from 2026-03-01 --to 2026-03-31

  # Actually delete
  python delete_duplicate_vouchers.py --type Purchase --from 2026-03-01 --to 2026-03-31 --execute
        """
    )
    parser.add_argument("--type", dest="voucher_type", default="Purchase",
                        help="Tally voucher type name (default: Purchase)")
    parser.add_argument("--from", dest="from_date", required=True,
                        help="Start date YYYY-MM-DD")
    parser.add_argument("--to", dest="to_date", required=True,
                        help="End date YYYY-MM-DD")
    parser.add_argument("--execute", action="store_true",
                        help="Actually delete duplicates (default is dry-run preview)")
    parser.add_argument("--prefix", dest="number_prefix", default="",
                        help="Only consider vouchers whose number starts with this prefix "
                             "(e.g. 'BILL/2026/03/' to target only March Odoo bills)")

    args = parser.parse_args()

    company = os.getenv("TALLY_COMPANY_NAME", "")
    if not company:
        log.error("[FAIL] TALLY_COMPANY_NAME not set in .env")
        sys.exit(1)

    tally = TallyConnector()
    try:
        tally.test_connection()
        log.info(f"[OK] Tally connected ({tally.base_url})")
    except ConnectionError as e:
        log.error(f"[FAIL] {e}")
        sys.exit(1)

    find_and_delete_duplicates(
        tally=tally,
        voucher_type=args.voucher_type,
        from_date=args.from_date,
        to_date=args.to_date,
        execute=args.execute,
        number_prefix=args.number_prefix,
    )


if __name__ == "__main__":
    main()
