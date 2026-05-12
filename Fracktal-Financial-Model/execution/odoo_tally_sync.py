"""
Odoo → Tally Prime Master Sync Orchestrator

This is the main entry point that runs the complete sync pipeline:
  1. Extract posted invoices/bills from Odoo
  2. Map Odoo accounts/partners → Tally ledger names
  3. Build Tally-compatible XML files
  4. Push XML to Tally Prime (or save for manual import)

Usage:
    python execution/odoo_tally_sync.py                     # Full sync (last 7 days, dry-run)
    python execution/odoo_tally_sync.py --push              # Full sync + push to Tally
    python execution/odoo_tally_sync.py --days 30           # Last 30 days
    python execution/odoo_tally_sync.py --init              # First-time setup: generate mappings
    python execution/odoo_tally_sync.py --status            # Check connection & sync status
    python execution/odoo_tally_sync.py --type purchase     # Only purchase vouchers
    python execution/odoo_tally_sync.py --create-masters    # Also create missing ledgers in Tally
"""

import os
import sys
import argparse
from datetime import datetime, timedelta
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

# Add execution directory to path
sys.path.insert(0, str(Path(__file__).parent))

from odoo_extractor import connect_odoo, extract_invoices, save_extract
from ledger_mapper import (
    init_account_mapping, init_tax_mapping, init_partner_mapping,
    apply_mapping, check_mapping, connect_odoo as mapper_connect
)
from tally_xml_builder import build_xml_files
from tally_xml_pusher import check_tally_connection, push_xml_files, view_sync_log

TMP_DIR = Path(".tmp")


def run_init():
    """First-time setup: generate all mapping CSVs from Odoo."""
    print("=" * 60)
    print("ODOO → TALLY SYNC: INITIAL SETUP")
    print("=" * 60)

    # Validate env vars
    required_vars = ['ODOO_URL', 'ODOO_DB', 'ODOO_USERNAME', 'ODOO_PASSWORD', 'TALLY_COMPANY_NAME']
    missing = [v for v in required_vars if not os.getenv(v)]
    if missing:
        print(f"\n✗ Missing environment variables in .env:")
        for v in missing:
            print(f"  - {v}")
        print(f"\nPlease fill in .env and try again.")
        return False

    print(f"\n--- Connecting to Odoo ---")
    models, uid = connect_odoo()

    print(f"\n--- Generating Account Mapping ---")
    init_account_mapping(models, uid)

    print(f"\n--- Generating Tax Mapping ---")
    init_tax_mapping(models, uid)

    print(f"\n--- Generating Partner Mapping ---")
    init_partner_mapping(models, uid)

    print(f"\n{'='*60}")
    print("SETUP COMPLETE")
    print(f"{'='*60}")
    print(f"\nNext steps:")
    print(f"  1. Review .tmp/ledger_mapping.csv - fill in Tally ledger names for NEEDS_MAPPING rows")
    print(f"  2. Review .tmp/tax_mapping.csv - verify tax ledger names match your Tally setup")
    print(f"  3. Review .tmp/partner_mapping.csv - verify party ledger names")
    print(f"  4. Enable XML Server in Tally Prime (F12 → Connectivity)")
    print(f"  5. Run: python execution/odoo_tally_sync.py --days 7")
    print(f"     (This will do a dry-run first; add --push to actually send to Tally)")

    return True


def run_status():
    """Check the current status of all components."""
    print("=" * 60)
    print("ODOO → TALLY SYNC: STATUS CHECK")
    print("=" * 60)

    # Check env vars
    print(f"\n--- Environment Variables ---")
    env_vars = {
        'ODOO_URL': os.getenv('ODOO_URL', ''),
        'ODOO_DB': os.getenv('ODOO_DB', ''),
        'ODOO_USERNAME': os.getenv('ODOO_USERNAME', ''),
        'TALLY_HOST': os.getenv('TALLY_HOST', 'localhost'),
        'TALLY_PORT': os.getenv('TALLY_PORT', '9000'),
        'TALLY_COMPANY_NAME': os.getenv('TALLY_COMPANY_NAME', ''),
    }
    for var, val in env_vars.items():
        status = '✓' if val else '✗'
        display = val if var != 'ODOO_PASSWORD' else '***' if val else ''
        print(f"  {status} {var}: {display}")

    # Check Odoo connection
    print(f"\n--- Odoo Connection ---")
    try:
        models, uid = connect_odoo()
        print(f"  ✓ Odoo is reachable")
    except Exception as e:
        print(f"  ✗ Cannot connect to Odoo: {e}")

    # Check Tally connection
    print(f"\n--- Tally Connection ---")
    check_tally_connection()

    # Check mapping files
    print(f"\n--- Mapping Files ---")
    check_mapping()

    # Check sync log
    print(f"\n--- Recent Sync Activity ---")
    view_sync_log(last_n=5)


def run_sync(days=7, date_from=None, date_to=None, move_type='all',
             push=False, create_masters=False, include_synced=False):
    """
    Run the complete Odoo → Tally sync pipeline.

    Args:
        days: Number of days to look back (default: 7)
        date_from: Start date (overrides days)
        date_to: End date
        move_type: 'purchase', 'sales', 'credit', 'debit', or 'all'
        push: If True, push to Tally. If False, dry-run (just build XML)
        create_masters: If True, also create ledger masters in Tally
        include_synced: If True, include already-synced vouchers
    """
    print("=" * 60)
    print(f"ODOO → TALLY SYNC: {'LIVE RUN' if push else 'DRY RUN'}")
    print("=" * 60)

    # Determine date range
    if date_from and date_to:
        pass
    else:
        date_to = datetime.now().strftime('%Y-%m-%d')
        date_from = (datetime.now() - timedelta(days=days)).strftime('%Y-%m-%d')

    print(f"  Date range: {date_from} to {date_to}")
    print(f"  Type: {move_type}")
    print(f"  Mode: {'PUSH to Tally' if push else 'Dry run (XML only)'}")

    # Determine move types
    type_map = {
        'purchase': ['in_invoice'],
        'sales': ['out_invoice'],
        'credit': ['out_refund'],
        'debit': ['in_refund'],
        'all': ['out_invoice', 'in_invoice', 'out_refund', 'in_refund'],
    }
    move_types = type_map.get(move_type, type_map['all'])

    # Step 1: Extract from Odoo
    print(f"\n{'='*40}")
    print("STEP 1: EXTRACT FROM ODOO")
    print(f"{'='*40}")

    try:
        models, uid = connect_odoo()
        vouchers = extract_invoices(
            models, uid, date_from, date_to,
            move_types=move_types,
            skip_synced=not include_synced
        )
    except Exception as e:
        print(f"\n✗ Extraction failed: {e}")
        return False

    if not vouchers:
        print("\n⚠ No new vouchers found. Nothing to sync.")
        return True

    extract_file = save_extract(vouchers)

    # Step 2: Apply ledger mapping
    print(f"\n{'='*40}")
    print("STEP 2: MAP LEDGERS")
    print(f"{'='*40}")

    try:
        mapped = apply_mapping(str(extract_file))
    except Exception as e:
        print(f"\n✗ Mapping failed: {e}")
        print("  Have you run --init first? Run: python execution/odoo_tally_sync.py --init")
        return False

    # Check for unmapped entries
    unmapped = [v for v in mapped if not v.get('mapping_complete', True)]
    if unmapped:
        print(f"\n⚠ {len(unmapped)} vouchers have unmapped entries.")
        print("  They will still be included with best-effort mapping.")
        print("  To fix: update mapping CSVs in .tmp/ and re-run.")

    # Step 3: Build Tally XML
    print(f"\n{'='*40}")
    print("STEP 3: BUILD TALLY XML")
    print(f"{'='*40}")

    try:
        xml_files = build_xml_files(
            single_file=False,
            create_masters=create_masters
        )
    except Exception as e:
        print(f"\n✗ XML build failed: {e}")
        return False

    if not xml_files:
        print("\n✗ No XML files generated")
        return False

    # Step 4: Push to Tally (or dry-run)
    print(f"\n{'='*40}")
    print(f"STEP 4: {'PUSH TO TALLY' if push else 'DRY RUN (review XML files)'}")
    print(f"{'='*40}")

    if push:
        try:
            results = push_xml_files(
                xml_files=xml_files,
                dry_run=False,
                masters_first=create_masters
            )
        except Exception as e:
            print(f"\n✗ Push failed: {e}")
            return False
    else:
        results = push_xml_files(
            xml_files=xml_files,
            dry_run=True,
            masters_first=create_masters
        )
        print(f"\n📂 XML files ready for review in: {TMP_DIR / 'tally_xml'}")
        print(f"   To push to Tally, re-run with --push flag")

    # Final summary
    print(f"\n{'='*60}")
    print("SYNC COMPLETE")
    print(f"{'='*60}")
    print(f"  Extracted: {len(vouchers)} vouchers from Odoo")
    print(f"  XML files: {len(xml_files)}")
    print(f"  Mode: {'Pushed to Tally' if push else 'Dry run (files saved)'}")
    if not push:
        print(f"\n  To push: python execution/odoo_tally_sync.py --push")

    return True


def main():
    parser = argparse.ArgumentParser(
        description="Odoo → Tally Prime Sync Pipeline",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  First-time setup:
    python execution/odoo_tally_sync.py --init

  Check status:
    python execution/odoo_tally_sync.py --status

  Dry run (last 7 days):
    python execution/odoo_tally_sync.py

  Sync last 30 days of purchase bills:
    python execution/odoo_tally_sync.py --days 30 --type purchase --push

  Sync specific date range:
    python execution/odoo_tally_sync.py --from 2026-02-01 --to 2026-02-28 --push

  Sync with ledger master creation:
    python execution/odoo_tally_sync.py --push --create-masters
        """
    )

    parser.add_argument('--init', action='store_true',
                        help='First-time setup: generate mapping CSVs from Odoo')
    parser.add_argument('--status', action='store_true',
                        help='Check connection status and sync history')
    parser.add_argument('--push', action='store_true',
                        help='Actually push to Tally (default is dry-run)')
    parser.add_argument('--days', type=int, default=7,
                        help='Sync last N days (default: 7)')
    parser.add_argument('--from', dest='date_from', type=str,
                        help='Start date (YYYY-MM-DD)')
    parser.add_argument('--to', dest='date_to', type=str,
                        help='End date (YYYY-MM-DD)')
    parser.add_argument('--type', dest='move_type',
                        choices=['purchase', 'sales', 'credit', 'debit', 'all'],
                        default='all', help='Voucher type to sync')
    parser.add_argument('--create-masters', action='store_true',
                        help='Create missing ledger masters in Tally')
    parser.add_argument('--include-synced', action='store_true',
                        help='Re-sync already-synced vouchers')

    args = parser.parse_args()

    if args.init:
        run_init()
    elif args.status:
        run_status()
    else:
        run_sync(
            days=args.days,
            date_from=args.date_from,
            date_to=args.date_to,
            move_type=args.move_type,
            push=args.push,
            create_masters=args.create_masters,
            include_synced=args.include_synced,
        )


if __name__ == "__main__":
    main()
