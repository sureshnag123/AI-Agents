#!/usr/bin/env python3
"""
Sync Report Generator

Reads sync logs and batch files to produce human-readable reports
of what was synced, what failed, and overall statistics.

Usage:
    python sync_report.py                    # Console summary
    python sync_report.py --detailed         # Line-by-line details
    python sync_report.py --csv              # Export to CSV
    python sync_report.py --failed-only      # Show only failures
    python sync_report.py --last-batch       # Show last batch details
"""

import os
import sys
import json
import csv
import argparse
from datetime import datetime
from pathlib import Path

SCRIPT_DIR = Path(__file__).parent.resolve()
PROJECT_ROOT = SCRIPT_DIR.parent
TMP_DIR = PROJECT_ROOT / ".tmp"
SYNC_LOG_FILE = TMP_DIR / "sync_log.json"


def load_sync_log() -> dict:
    if SYNC_LOG_FILE.exists():
        with open(SYNC_LOG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"last_sync": None, "batches": [], "synced_ids": {}}


def load_batch(batch_id: str) -> dict:
    batch_file = TMP_DIR / f"{batch_id}.json"
    if batch_file.exists():
        with open(batch_file, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def print_summary(sync_log: dict):
    """Print overall sync summary."""
    batches = sync_log.get("batches", [])
    print("\n" + "=" * 70)
    print("ODOO → TALLY SYNC REPORT")
    print("=" * 70)

    print(f"\nLast sync: {sync_log.get('last_sync', 'Never')}")
    print(f"Total batches: {len(batches)}")

    if not batches:
        print("\nNo sync batches found. Run a sync first:")
        print("  python execution/odoo_tally_sync.py --dry-run")
        return

    # Aggregate stats
    total_synced = sum(b.get("synced", 0) for b in batches)
    total_errors = sum(b.get("errors", 0) for b in batches)

    print(f"\nAll-time totals:")
    print(f"  Records synced: {total_synced}")
    print(f"  Errors:         {total_errors}")

    # Per-type breakdown from synced_ids
    synced_ids = sync_log.get("synced_ids", {})
    if synced_ids:
        print(f"\nSynced records by type:")
        for vtype, ids in synced_ids.items():
            print(f"  {vtype:<20} {len(ids)} records")

    # Recent batches
    print(f"\nRecent batches (last 10):")
    print(f"  {'Batch ID':<35} {'Date':<22} {'Synced':>7} {'Errors':>7} {'Mode':<12}")
    print(f"  {'─'*35} {'─'*22} {'─'*7} {'─'*7} {'─'*12}")
    for b in batches[-10:]:
        print(f"  {b.get('batch_id','?'):<35} "
              f"{b.get('timestamp','')[:19]:<22} "
              f"{b.get('synced',0):>7} "
              f"{b.get('errors',0):>7} "
              f"{b.get('mode','?'):<12}")


def print_batch_details(batch: dict, failed_only: bool = False):
    """Print detailed batch results."""
    if not batch:
        print("No batch data found.")
        return

    print(f"\n{'='*70}")
    print(f"BATCH: {batch.get('batch_id', '?')}")
    print(f"{'='*70}")
    print(f"Started:  {batch.get('started_at', '?')}")
    print(f"Completed: {batch.get('completed_at', '?')}")
    print(f"Mode:     {batch.get('mode', '?')}")
    print(f"Range:    {batch.get('from_date', '?')} to {batch.get('to_date', '?')}")
    print(f"Dry Run:  {batch.get('dry_run', False)}")

    totals = batch.get("totals", {})
    print(f"\nTotals: {totals.get('synced', 0)} synced, "
          f"{totals.get('skipped', 0)} skipped, "
          f"{totals.get('errors', 0)} errors")

    for vtype, data in batch.get("types", {}).items():
        print(f"\n  ── {vtype.upper()} ──")
        print(f"  Fetched: {data.get('fetched', 0)}, Synced: {data.get('synced', 0)}, "
              f"Skipped: {data.get('skipped', 0)}, Errors: {data.get('errors', 0)}")

        details = data.get("details", [])
        if failed_only:
            details = [d for d in details if d.get("status") == "error"]

        if details:
            for d in details:
                status = d.get("status", "?")
                icon = "✓" if status == "synced" else "○" if status == "dry_run" else "✗"
                msg = d.get("message", "")
                extra = f" — {msg}" if msg else ""
                partner = d.get("partner", "")
                amount = d.get("amount")
                amount_str = f" | ₹{amount:,.2f}" if amount else ""
                partner_str = f" | {partner}" if partner else ""
                print(f"    {icon} {d.get('name','?')}{partner_str}{amount_str}{extra}")


def export_csv(batch: dict, output_path: Path):
    """Export batch details to CSV."""
    rows = []
    for vtype, data in batch.get("types", {}).items():
        for d in data.get("details", []):
            rows.append({
                "batch_id": batch.get("batch_id", ""),
                "voucher_type": vtype,
                "odoo_id": d.get("id", ""),
                "name": d.get("name", ""),
                "partner": d.get("partner", ""),
                "amount": d.get("amount", ""),
                "status": d.get("status", ""),
                "message": d.get("message", ""),
            })

    with open(output_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=[
            "batch_id", "voucher_type", "odoo_id", "name",
            "partner", "amount", "status", "message"
        ])
        writer.writeheader()
        writer.writerows(rows)

    print(f"Exported {len(rows)} records to {output_path}")


def main():
    parser = argparse.ArgumentParser(description="Odoo-Tally Sync Report")
    parser.add_argument("--detailed", action="store_true",
                        help="Show line-by-line details of last batch")
    parser.add_argument("--last-batch", action="store_true",
                        help="Show details for the last sync batch")
    parser.add_argument("--batch-id", type=str, default=None,
                        help="Show details for a specific batch")
    parser.add_argument("--failed-only", action="store_true",
                        help="Show only failed records")
    parser.add_argument("--csv", action="store_true",
                        help="Export last batch to CSV")

    args = parser.parse_args()
    sync_log = load_sync_log()

    if args.batch_id:
        batch = load_batch(args.batch_id)
        print_batch_details(batch, failed_only=args.failed_only)
    elif args.last_batch or args.detailed:
        batches = sync_log.get("batches", [])
        if batches:
            batch = load_batch(batches[-1]["batch_id"])
            print_batch_details(batch, failed_only=args.failed_only)
        else:
            print("No batches found.")
    elif args.csv:
        batches = sync_log.get("batches", [])
        if batches:
            batch = load_batch(batches[-1]["batch_id"])
            out = TMP_DIR / f"sync_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
            export_csv(batch, out)
        else:
            print("No batches found.")
    else:
        print_summary(sync_log)


if __name__ == "__main__":
    main()
