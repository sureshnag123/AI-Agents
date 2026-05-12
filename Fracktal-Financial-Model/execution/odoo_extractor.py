"""
Odoo Data Extractor: Extracts posted invoices and bills from Odoo 19 via XML-RPC.

Usage:
    python execution/odoo_extractor.py                          # Extract last 7 days
    python execution/odoo_extractor.py --days 30                # Extract last 30 days
    python execution/odoo_extractor.py --from 2026-01-01 --to 2026-01-31  # Date range
    python execution/odoo_extractor.py --type purchase          # Only purchase vouchers
    python execution/odoo_extractor.py --type sales             # Only sales vouchers
"""

import os
import sys
import json
import csv
import xmlrpc.client
import argparse
from datetime import datetime, timedelta
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

# --- Configuration ---
ODOO_URL = os.getenv("ODOO_URL", "http://localhost:8069")
ODOO_DB = os.getenv("ODOO_DB", "")
ODOO_USERNAME = os.getenv("ODOO_USERNAME", "")
ODOO_PASSWORD = os.getenv("ODOO_PASSWORD", "")

TMP_DIR = Path(".tmp")
EXTRACT_FILE = TMP_DIR / "odoo_extract.json"
SYNC_LOG_FILE = TMP_DIR / "sync_log.csv"


def connect_odoo():
    """Establish XML-RPC connection to Odoo 19."""
    common = xmlrpc.client.ServerProxy(f"{ODOO_URL}/xmlrpc/2/common")
    try:
        version = common.version()
        print(f"  Odoo version: {version.get('server_version', 'unknown')}")
    except Exception as e:
        print(f"⚠ Could not fetch version: {e}")

    uid = common.authenticate(ODOO_DB, ODOO_USERNAME, ODOO_PASSWORD, {})
    if not uid:
        raise ConnectionError(
            "Failed to authenticate with Odoo.\n"
            "Please check ODOO_URL, ODOO_DB, ODOO_USERNAME, ODOO_PASSWORD in .env"
        )
    models = xmlrpc.client.ServerProxy(f"{ODOO_URL}/xmlrpc/2/object")
    print(f"✓ Connected to Odoo at {ODOO_URL} (db: {ODOO_DB}, uid: {uid})")
    return models, uid


def get_synced_invoice_ids():
    """Load already-synced invoice IDs from sync log to avoid duplicates."""
    synced = set()
    if SYNC_LOG_FILE.exists():
        with open(SYNC_LOG_FILE, 'r', encoding='utf-8') as f:
            for row in csv.DictReader(f):
                if row.get('status') == 'synced':
                    synced.add(row.get('odoo_invoice_name', ''))
    return synced


def extract_invoices(models, uid, date_from, date_to, move_types=None, skip_synced=True):
    """
    Extract posted invoices/bills from Odoo.

    Args:
        models: Odoo XML-RPC models proxy
        uid: Authenticated user ID
        date_from: Start date (string YYYY-MM-DD)
        date_to: End date (string YYYY-MM-DD)
        move_types: List of move types to extract. Options:
            - 'out_invoice' (Customer Invoice / Sales)
            - 'in_invoice' (Vendor Bill / Purchase)
            - 'out_refund' (Customer Credit Note)
            - 'in_refund' (Vendor Debit Note)
        skip_synced: Skip already-synced invoices
    """
    if move_types is None:
        move_types = ['out_invoice', 'in_invoice', 'out_refund', 'in_refund']

    # Build domain filter
    domain = [
        ('state', '=', 'posted'),
        ('move_type', 'in', move_types),
        ('invoice_date', '>=', date_from),
        ('invoice_date', '<=', date_to),
    ]

    # Get already synced IDs
    synced_names = get_synced_invoice_ids() if skip_synced else set()

    # Fetch invoices
    print(f"\n--- Fetching invoices from {date_from} to {date_to} ---")
    print(f"  Move types: {move_types}")

    invoice_ids = models.execute_kw(
        ODOO_DB, uid, ODOO_PASSWORD,
        'account.move', 'search',
        [domain],
        {'order': 'invoice_date asc'}
    )
    print(f"  Found {len(invoice_ids)} posted invoices/bills")

    # Fetch invoice details in batches
    vouchers = []
    batch_size = 50

    for i in range(0, len(invoice_ids), batch_size):
        batch_ids = invoice_ids[i:i + batch_size]
        invoices = models.execute_kw(
            ODOO_DB, uid, ODOO_PASSWORD,
            'account.move', 'read',
            [batch_ids],
            {'fields': [
                'id', 'name', 'ref', 'move_type', 'state',
                'invoice_date', 'date', 'invoice_date_due',
                'partner_id', 'amount_total', 'amount_untaxed',
                'amount_tax', 'currency_id', 'company_id',
                'invoice_line_ids', 'narration',
                'payment_reference', 'invoice_origin',
            ]}
        )

        for inv in invoices:
            # Skip already synced
            if inv['name'] in synced_names:
                continue

            # Fetch line items
            lines = fetch_invoice_lines(models, uid, inv['invoice_line_ids'])

            # Map move_type to Tally voucher type
            tally_vch_type = {
                'out_invoice': 'Sales',
                'in_invoice': 'Purchase',
                'out_refund': 'Credit Note',
                'in_refund': 'Debit Note',
            }.get(inv['move_type'], 'Journal')

            voucher = {
                'odoo_id': inv['id'],
                'name': inv['name'],
                'ref': inv.get('ref', ''),
                'move_type': inv['move_type'],
                'tally_voucher_type': tally_vch_type,
                'invoice_date': inv['invoice_date'],
                'due_date': inv.get('invoice_date_due', ''),
                'partner_id': inv['partner_id'][0] if inv.get('partner_id') else None,
                'partner_name': inv['partner_id'][1] if inv.get('partner_id') else '',
                'amount_total': inv['amount_total'],
                'amount_untaxed': inv['amount_untaxed'],
                'amount_tax': inv['amount_tax'],
                'currency': inv['currency_id'][1] if inv.get('currency_id') else 'INR',
                'narration': clean_html(inv.get('narration', '') or ''),
                'origin': inv.get('invoice_origin', ''),
                'payment_reference': inv.get('payment_reference', ''),
                'lines': lines,
            }
            vouchers.append(voucher)

        print(f"  Processed batch {i // batch_size + 1}/{(len(invoice_ids) + batch_size - 1) // batch_size}")

    skipped = len(invoice_ids) - len(vouchers)
    print(f"\n  Total extracted: {len(vouchers)}")
    if skipped > 0:
        print(f"  Skipped (already synced): {skipped}")

    return vouchers


def fetch_invoice_lines(models, uid, line_ids):
    """Fetch detailed line items for an invoice."""
    if not line_ids:
        return []

    lines_data = models.execute_kw(
        ODOO_DB, uid, ODOO_PASSWORD,
        'account.move.line', 'read',
        [line_ids],
        {'fields': [
            'id', 'name', 'account_id', 'product_id',
            'quantity', 'price_unit', 'price_subtotal',
            'price_total', 'tax_ids', 'debit', 'credit',
            'balance', 'display_type',
        ]}
    )

    lines = []
    for line in lines_data:
        # Skip section/note lines
        if line.get('display_type') in ('line_section', 'line_note'):
            continue

        # Fetch tax details for this line
        tax_details = []
        if line.get('tax_ids'):
            taxes = models.execute_kw(
                ODOO_DB, uid, ODOO_PASSWORD,
                'account.tax', 'read',
                [line['tax_ids']],
                {'fields': ['id', 'name', 'amount', 'type_tax_use']}
            )
            for tax in taxes:
                tax_amount = (line.get('price_subtotal', 0) * tax['amount']) / 100
                tax_details.append({
                    'tax_id': tax['id'],
                    'tax_name': tax['name'],
                    'tax_rate': tax['amount'],
                    'tax_type': tax['type_tax_use'],
                    'amount': round(tax_amount, 2),
                })

        lines.append({
            'line_id': line['id'],
            'description': line.get('name', ''),
            'account_id': line['account_id'][0] if line.get('account_id') else None,
            'account_name': line['account_id'][1] if line.get('account_id') else '',
            'product_id': line['product_id'][0] if line.get('product_id') else None,
            'product_name': line['product_id'][1] if line.get('product_id') else '',
            'quantity': line.get('quantity', 0),
            'price_unit': line.get('price_unit', 0),
            'price_subtotal': line.get('price_subtotal', 0),
            'price_total': line.get('price_total', 0),
            'debit': line.get('debit', 0),
            'credit': line.get('credit', 0),
            'taxes': tax_details,
        })

    return lines


def clean_html(text):
    """Remove HTML tags from Odoo's rich text fields."""
    if not text:
        return ''
    import re
    clean = re.sub(r'<[^>]+>', '', text)
    clean = clean.replace('&nbsp;', ' ').replace('&amp;', '&')
    return clean.strip()


def save_extract(vouchers, output_file=None):
    """Save extracted vouchers to JSON file."""
    TMP_DIR.mkdir(exist_ok=True)
    if output_file is None:
        output_file = EXTRACT_FILE

    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(vouchers, f, indent=2, default=str)

    print(f"\n✓ Extracted data saved to {output_file}")

    # Summary
    type_counts = {}
    for v in vouchers:
        t = v['tally_voucher_type']
        type_counts[t] = type_counts.get(t, 0) + 1

    total_amount = sum(v['amount_total'] for v in vouchers)

    print(f"\n--- Extraction Summary ---")
    for vtype, count in sorted(type_counts.items()):
        print(f"  {vtype}: {count} vouchers")
    print(f"  Total amount: ₹{total_amount:,.2f}")

    return output_file


def main():
    parser = argparse.ArgumentParser(description="Odoo Data Extractor for Tally Sync")
    parser.add_argument('--days', type=int, default=7, help='Extract last N days (default: 7)')
    parser.add_argument('--from', dest='date_from', type=str, help='Start date (YYYY-MM-DD)')
    parser.add_argument('--to', dest='date_to', type=str, help='End date (YYYY-MM-DD)')
    parser.add_argument('--type', choices=['purchase', 'sales', 'credit', 'debit', 'all'],
                        default='all', help='Voucher type to extract')
    parser.add_argument('--include-synced', action='store_true',
                        help='Include already-synced invoices')
    parser.add_argument('--output', type=str, help='Output file path')
    args = parser.parse_args()

    # Determine date range
    if args.date_from and args.date_to:
        date_from = args.date_from
        date_to = args.date_to
    else:
        date_to = datetime.now().strftime('%Y-%m-%d')
        date_from = (datetime.now() - timedelta(days=args.days)).strftime('%Y-%m-%d')

    # Determine move types
    type_map = {
        'purchase': ['in_invoice'],
        'sales': ['out_invoice'],
        'credit': ['out_refund'],
        'debit': ['in_refund'],
        'all': ['out_invoice', 'in_invoice', 'out_refund', 'in_refund'],
    }
    move_types = type_map.get(args.type, type_map['all'])

    # Connect and extract
    models, uid = connect_odoo()
    vouchers = extract_invoices(
        models, uid,
        date_from, date_to,
        move_types=move_types,
        skip_synced=not args.include_synced
    )

    if vouchers:
        save_extract(vouchers, args.output)
    else:
        print("\n⚠ No new vouchers found for the specified date range and filters.")


if __name__ == "__main__":
    main()
