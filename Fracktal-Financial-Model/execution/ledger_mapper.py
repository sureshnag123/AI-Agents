"""
Ledger Mapper: Maps Odoo accounts/partners to Tally Prime ledger names.

Usage:
    python execution/ledger_mapper.py --init       # Generate initial mapping CSV from Odoo
    python execution/ledger_mapper.py --map FILE   # Apply mapping to extracted Odoo data
    python execution/ledger_mapper.py --check      # Check for unmapped entries
"""

import os
import sys
import csv
import json
import xmlrpc.client
import argparse
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

# --- Configuration ---
ODOO_URL = os.getenv("ODOO_URL", "http://localhost:8069")
ODOO_DB = os.getenv("ODOO_DB", "")
ODOO_USERNAME = os.getenv("ODOO_USERNAME", "")
ODOO_PASSWORD = os.getenv("ODOO_PASSWORD", "")

TMP_DIR = Path(".tmp")
MAPPING_FILE = TMP_DIR / "ledger_mapping.csv"
TAX_MAPPING_FILE = TMP_DIR / "tax_mapping.csv"
PARTNER_MAPPING_FILE = TMP_DIR / "partner_mapping.csv"


def connect_odoo():
    """Establish XML-RPC connection to Odoo 19."""
    common = xmlrpc.client.ServerProxy(f"{ODOO_URL}/xmlrpc/2/common")
    uid = common.authenticate(ODOO_DB, ODOO_USERNAME, ODOO_PASSWORD, {})
    if not uid:
        raise ConnectionError("Failed to authenticate with Odoo. Check credentials in .env")
    models = xmlrpc.client.ServerProxy(f"{ODOO_URL}/xmlrpc/2/object")
    print(f"✓ Connected to Odoo at {ODOO_URL} (db: {ODOO_DB}, uid: {uid})")
    return models, uid


def fetch_odoo_accounts(models, uid):
    """Fetch Chart of Accounts from Odoo."""
    accounts = models.execute_kw(
        ODOO_DB, uid, ODOO_PASSWORD,
        'account.account', 'search_read',
        [[]],
        {'fields': ['id', 'code', 'name', 'account_type'], 'order': 'code'}
    )
    print(f"  Fetched {len(accounts)} accounts from Odoo")
    return accounts


def fetch_odoo_taxes(models, uid):
    """Fetch tax definitions from Odoo."""
    taxes = models.execute_kw(
        ODOO_DB, uid, ODOO_PASSWORD,
        'account.tax', 'search_read',
        [[('active', '=', True)]],
        {'fields': ['id', 'name', 'amount', 'type_tax_use', 'tax_group_id']}
    )
    print(f"  Fetched {len(taxes)} taxes from Odoo")
    return taxes


def fetch_odoo_partners(models, uid):
    """Fetch partners (customers/vendors) from Odoo."""
    partners = models.execute_kw(
        ODOO_DB, uid, ODOO_PASSWORD,
        'res.partner', 'search_read',
        [[('is_company', '=', True), '|', ('customer_rank', '>', 0), ('supplier_rank', '>', 0)]],
        {'fields': ['id', 'name', 'vat', 'customer_rank', 'supplier_rank', 'property_account_receivable_id', 'property_account_payable_id']}
    )
    print(f"  Fetched {len(partners)} partners from Odoo")
    return partners


def init_account_mapping(models, uid):
    """Generate initial ledger_mapping.csv from Odoo Chart of Accounts."""
    TMP_DIR.mkdir(exist_ok=True)
    accounts = fetch_odoo_accounts(models, uid)

    # Suggest Tally ledger names based on Odoo account names
    rows = []
    for acc in accounts:
        suggested_tally_name = suggest_tally_ledger(acc['name'], acc.get('account_type', ''))
        rows.append({
            'odoo_account_id': acc['id'],
            'odoo_account_code': acc['code'],
            'odoo_account_name': acc['name'],
            'odoo_account_type': acc.get('account_type', ''),
            'tally_ledger_name': suggested_tally_name,
            'tally_group': suggest_tally_group(acc.get('account_type', '')),
            'status': 'auto' if suggested_tally_name else 'NEEDS_MAPPING'
        })

    with open(MAPPING_FILE, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=[
            'odoo_account_id', 'odoo_account_code', 'odoo_account_name',
            'odoo_account_type', 'tally_ledger_name', 'tally_group', 'status'
        ])
        writer.writeheader()
        writer.writerows(rows)

    needs_mapping = sum(1 for r in rows if r['status'] == 'NEEDS_MAPPING')
    print(f"\n✓ Ledger mapping saved to {MAPPING_FILE}")
    print(f"  Total accounts: {len(rows)}")
    print(f"  Auto-mapped: {len(rows) - needs_mapping}")
    print(f"  Needs manual mapping: {needs_mapping}")
    if needs_mapping > 0:
        print(f"\n⚠ Please open {MAPPING_FILE} and fill in 'tally_ledger_name' for rows marked 'NEEDS_MAPPING'")

    return rows


def init_tax_mapping(models, uid):
    """Generate initial tax_mapping.csv from Odoo taxes."""
    TMP_DIR.mkdir(exist_ok=True)
    taxes = fetch_odoo_taxes(models, uid)

    rows = []
    for tax in taxes:
        suggested = suggest_tally_tax_ledger(tax['name'], tax['amount'], tax['type_tax_use'])
        rows.append({
            'odoo_tax_id': tax['id'],
            'odoo_tax_name': tax['name'],
            'odoo_tax_rate': tax['amount'],
            'odoo_tax_type': tax['type_tax_use'],
            'tally_ledger_name': suggested,
            'status': 'auto' if suggested else 'NEEDS_MAPPING'
        })

    with open(TAX_MAPPING_FILE, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=[
            'odoo_tax_id', 'odoo_tax_name', 'odoo_tax_rate',
            'odoo_tax_type', 'tally_ledger_name', 'status'
        ])
        writer.writeheader()
        writer.writerows(rows)

    print(f"✓ Tax mapping saved to {TAX_MAPPING_FILE}")
    return rows


def init_partner_mapping(models, uid):
    """Generate initial partner_mapping.csv from Odoo partners."""
    TMP_DIR.mkdir(exist_ok=True)
    partners = fetch_odoo_partners(models, uid)

    rows = []
    for p in partners:
        partner_type = []
        if p.get('customer_rank', 0) > 0:
            partner_type.append('Customer')
        if p.get('supplier_rank', 0) > 0:
            partner_type.append('Vendor')

        rows.append({
            'odoo_partner_id': p['id'],
            'odoo_partner_name': p['name'],
            'odoo_partner_gstin': p.get('vat', ''),
            'odoo_partner_type': '/'.join(partner_type),
            'tally_ledger_name': p['name'],  # Default: same name
            'tally_group': 'Sundry Creditors' if 'Vendor' in partner_type else 'Sundry Debtors',
            'status': 'auto'
        })

    with open(PARTNER_MAPPING_FILE, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=[
            'odoo_partner_id', 'odoo_partner_name', 'odoo_partner_gstin',
            'odoo_partner_type', 'tally_ledger_name', 'tally_group', 'status'
        ])
        writer.writeheader()
        writer.writerows(rows)

    print(f"✓ Partner mapping saved to {PARTNER_MAPPING_FILE}")
    return rows


def suggest_tally_ledger(account_name, account_type):
    """Suggest a Tally ledger name based on Odoo account name and type."""
    # Direct pass-through: Tally can use same names in many cases
    name_lower = account_name.lower()

    # Common mappings
    if 'bank' in name_lower:
        return account_name
    if 'cash' in name_lower:
        return account_name
    if 'receivable' in name_lower:
        return ""  # Will be mapped via partner
    if 'payable' in name_lower:
        return ""  # Will be mapped via partner
    if 'purchase' in name_lower:
        return account_name
    if 'sale' in name_lower or 'revenue' in name_lower or 'income' in name_lower:
        return account_name
    if 'depreciation' in name_lower:
        return account_name
    if 'tax' in name_lower or 'gst' in name_lower or 'cgst' in name_lower or 'sgst' in name_lower or 'igst' in name_lower:
        return account_name
    if 'expense' in name_lower:
        return account_name
    if 'capital' in name_lower:
        return account_name
    if 'stock' in name_lower or 'inventory' in name_lower:
        return account_name

    # Default: use same name (user can override in CSV)
    return account_name


def suggest_tally_group(account_type):
    """Suggest Tally group based on Odoo account type."""
    type_map = {
        'asset_receivable': 'Sundry Debtors',
        'liability_payable': 'Sundry Creditors',
        'asset_cash': 'Cash-in-Hand',
        'asset_current': 'Current Assets',
        'asset_non_current': 'Fixed Assets',
        'asset_fixed': 'Fixed Assets',
        'liability_current': 'Current Liabilities',
        'liability_non_current': 'Loans (Liability)',
        'equity': 'Capital Account',
        'income': 'Sales Accounts',
        'income_other': 'Indirect Income',
        'expense': 'Purchase Accounts',
        'expense_direct_cost': 'Direct Expenses',
        'expense_depreciation': 'Indirect Expenses',
        'off_balance': 'Suspense A/c',
    }
    return type_map.get(account_type, '')


def suggest_tally_tax_ledger(tax_name, rate, tax_type):
    """Suggest Tally tax ledger name based on Odoo tax."""
    name_lower = tax_name.lower()
    suffix = "Output" if tax_type == 'sale' else "Input"

    if 'cgst' in name_lower:
        return f"CGST {suffix} @ {rate}%"
    if 'sgst' in name_lower:
        return f"SGST {suffix} @ {rate}%"
    if 'igst' in name_lower:
        return f"IGST {suffix} @ {rate}%"
    if 'cess' in name_lower:
        return f"Cess {suffix} @ {rate}%"
    if 'gst' in name_lower:
        # Generic GST - try to determine type from rate
        if rate == 9 or rate == 2.5 or rate == 6:
            return f"CGST {suffix} @ {rate}%"  # Likely CGST portion
        elif rate == 18 or rate == 5 or rate == 12 or rate == 28:
            return f"IGST {suffix} @ {rate}%"  # Likely IGST
    if 'tds' in name_lower:
        return f"TDS {tax_name}"
    if 'tcs' in name_lower:
        return f"TCS {tax_name}"

    return ""  # Needs manual mapping


def load_mapping():
    """Load all mapping files and return as dictionaries."""
    mappings = {
        'accounts': {},   # odoo_account_id → tally_ledger_name
        'taxes': {},      # odoo_tax_id → tally_ledger_name
        'partners': {},   # odoo_partner_id → {tally_ledger_name, tally_group}
    }

    if MAPPING_FILE.exists():
        with open(MAPPING_FILE, 'r', encoding='utf-8') as f:
            for row in csv.DictReader(f):
                if row['tally_ledger_name'] and row['status'] != 'NEEDS_MAPPING':
                    mappings['accounts'][int(row['odoo_account_id'])] = row['tally_ledger_name']

    if TAX_MAPPING_FILE.exists():
        with open(TAX_MAPPING_FILE, 'r', encoding='utf-8') as f:
            for row in csv.DictReader(f):
                if row['tally_ledger_name'] and row['status'] != 'NEEDS_MAPPING':
                    mappings['taxes'][int(row['odoo_tax_id'])] = row['tally_ledger_name']

    if PARTNER_MAPPING_FILE.exists():
        with open(PARTNER_MAPPING_FILE, 'r', encoding='utf-8') as f:
            for row in csv.DictReader(f):
                mappings['partners'][int(row['odoo_partner_id'])] = {
                    'tally_ledger_name': row['tally_ledger_name'],
                    'tally_group': row['tally_group'],
                }

    print(f"✓ Loaded mappings: {len(mappings['accounts'])} accounts, "
          f"{len(mappings['taxes'])} taxes, {len(mappings['partners'])} partners")
    return mappings


def apply_mapping(extract_file, output_file=None):
    """Apply ledger mapping to extracted Odoo data."""
    mappings = load_mapping()

    with open(extract_file, 'r', encoding='utf-8') as f:
        vouchers = json.load(f)

    mapped = []
    unmapped_items = []

    for v in vouchers:
        mapped_v = v.copy()
        mapped_v['mapped_lines'] = []
        has_unmapped = False

        # Map partner
        partner_id = v.get('partner_id')
        if partner_id and partner_id in mappings['partners']:
            mapped_v['tally_party_ledger'] = mappings['partners'][partner_id]['tally_ledger_name']
            mapped_v['tally_party_group'] = mappings['partners'][partner_id]['tally_group']
        elif partner_id:
            # Use partner name directly as fallback
            mapped_v['tally_party_ledger'] = v.get('partner_name', f'Partner_{partner_id}')
            mapped_v['tally_party_group'] = 'Sundry Creditors' if v.get('move_type') in ('in_invoice', 'in_refund') else 'Sundry Debtors'

        # Map line items
        for line in v.get('lines', []):
            mapped_line = line.copy()
            account_id = line.get('account_id')

            if account_id and account_id in mappings['accounts']:
                mapped_line['tally_ledger'] = mappings['accounts'][account_id]
            else:
                mapped_line['tally_ledger'] = line.get('account_name', '')
                if not mapped_line['tally_ledger']:
                    has_unmapped = True
                    unmapped_items.append({
                        'voucher': v.get('name', ''),
                        'type': 'account',
                        'odoo_id': account_id,
                        'odoo_name': line.get('account_name', ''),
                    })

            # Map taxes on this line
            mapped_line['tally_tax_ledgers'] = []
            for tax in line.get('taxes', []):
                tax_id = tax.get('tax_id')
                if tax_id and tax_id in mappings['taxes']:
                    mapped_line['tally_tax_ledgers'].append({
                        'ledger': mappings['taxes'][tax_id],
                        'amount': tax.get('amount', 0),
                    })
                else:
                    mapped_line['tally_tax_ledgers'].append({
                        'ledger': tax.get('tax_name', ''),
                        'amount': tax.get('amount', 0),
                    })
                    if not tax.get('tax_name'):
                        has_unmapped = True
                        unmapped_items.append({
                            'voucher': v.get('name', ''),
                            'type': 'tax',
                            'odoo_id': tax_id,
                            'odoo_name': tax.get('tax_name', ''),
                        })

            mapped_v['mapped_lines'].append(mapped_line)

        mapped_v['mapping_complete'] = not has_unmapped
        mapped.append(mapped_v)

    # Save mapped data
    if output_file is None:
        output_file = TMP_DIR / "mapped_vouchers.json"

    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(mapped, f, indent=2, default=str)

    complete = sum(1 for m in mapped if m['mapping_complete'])
    print(f"\n✓ Mapping applied: {len(mapped)} vouchers")
    print(f"  Fully mapped: {complete}")
    print(f"  Has unmapped items: {len(mapped) - complete}")

    if unmapped_items:
        print(f"\n⚠ Unmapped items ({len(unmapped_items)}):")
        for item in unmapped_items[:10]:
            print(f"  - [{item['type']}] {item['odoo_name']} (ID: {item['odoo_id']}) in voucher {item['voucher']}")
        if len(unmapped_items) > 10:
            print(f"  ... and {len(unmapped_items) - 10} more")

    return mapped


def check_mapping():
    """Check mapping files for completeness."""
    issues = []

    if not MAPPING_FILE.exists():
        issues.append(f"Account mapping file not found: {MAPPING_FILE}")
    else:
        with open(MAPPING_FILE, 'r', encoding='utf-8') as f:
            rows = list(csv.DictReader(f))
            needs = [r for r in rows if r['status'] == 'NEEDS_MAPPING']
            if needs:
                issues.append(f"Account mapping: {len(needs)} entries need manual mapping")
                for r in needs[:5]:
                    issues.append(f"  - {r['odoo_account_code']} {r['odoo_account_name']}")

    if not TAX_MAPPING_FILE.exists():
        issues.append(f"Tax mapping file not found: {TAX_MAPPING_FILE}")
    else:
        with open(TAX_MAPPING_FILE, 'r', encoding='utf-8') as f:
            rows = list(csv.DictReader(f))
            needs = [r for r in rows if r['status'] == 'NEEDS_MAPPING']
            if needs:
                issues.append(f"Tax mapping: {len(needs)} entries need manual mapping")

    if not PARTNER_MAPPING_FILE.exists():
        issues.append(f"Partner mapping file not found: {PARTNER_MAPPING_FILE}")

    if issues:
        print("⚠ Mapping issues found:")
        for issue in issues:
            print(f"  {issue}")
    else:
        print("✓ All mappings are complete")

    return len(issues) == 0


def main():
    parser = argparse.ArgumentParser(description="Odoo → Tally Ledger Mapper")
    parser.add_argument('--init', action='store_true', help='Initialize mapping CSVs from Odoo')
    parser.add_argument('--map', type=str, help='Apply mapping to extracted data file')
    parser.add_argument('--check', action='store_true', help='Check mapping completeness')
    args = parser.parse_args()

    if args.init:
        models, uid = connect_odoo()
        print("\n--- Generating Account Mapping ---")
        init_account_mapping(models, uid)
        print("\n--- Generating Tax Mapping ---")
        init_tax_mapping(models, uid)
        print("\n--- Generating Partner Mapping ---")
        init_partner_mapping(models, uid)
        print("\n✓ All mapping files generated. Please review them in .tmp/ folder.")

    elif args.map:
        apply_mapping(args.map)

    elif args.check:
        check_mapping()

    else:
        parser.print_help()


if __name__ == "__main__":
    main()
