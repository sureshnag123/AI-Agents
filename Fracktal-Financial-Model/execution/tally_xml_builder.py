"""
Tally XML Builder: Converts mapped Odoo voucher data into Tally Prime XML import format.

Usage:
    python execution/tally_xml_builder.py                            # Build from default mapped file
    python execution/tally_xml_builder.py --input mapped_vouchers.json
    python execution/tally_xml_builder.py --single-file              # All vouchers in one XML
    python execution/tally_xml_builder.py --create-masters           # Also create ledger masters
"""

import os
import sys
import json
import argparse
from datetime import datetime
from pathlib import Path
from xml.etree.ElementTree import Element, SubElement, tostring, ElementTree
from xml.dom.minidom import parseString
from dotenv import load_dotenv

load_dotenv()

# --- Configuration ---
TALLY_COMPANY = os.getenv("TALLY_COMPANY_NAME", "")
TMP_DIR = Path(".tmp")
XML_DIR = TMP_DIR / "tally_xml"
MAPPED_FILE = TMP_DIR / "mapped_vouchers.json"


def format_date(date_str):
    """Convert YYYY-MM-DD to YYYYMMDD (Tally format)."""
    if not date_str:
        return datetime.now().strftime('%Y%m%d')
    return date_str.replace('-', '')


def build_envelope(company_name, report_name="Vouchers"):
    """Create the base Tally XML envelope structure."""
    envelope = Element('ENVELOPE')

    header = SubElement(envelope, 'HEADER')
    SubElement(header, 'TALLYREQUEST').text = 'Import Data'

    body = SubElement(envelope, 'BODY')
    import_data = SubElement(body, 'IMPORTDATA')

    request_desc = SubElement(import_data, 'REQUESTDESC')
    SubElement(request_desc, 'REPORTNAME').text = report_name

    static_vars = SubElement(request_desc, 'STATICVARIABLES')
    SubElement(static_vars, 'SVCURRENTCOMPANY').text = company_name

    request_data = SubElement(import_data, 'REQUESTDATA')

    return envelope, request_data


def build_purchase_voucher(voucher, company_name):
    """Build Tally XML for a Purchase voucher (Vendor Bill)."""
    tally_msg = Element('TALLYMESSAGE')
    tally_msg.set('xmlns:UDF', 'TallyUDF')

    vch = SubElement(tally_msg, 'VOUCHER')
    vch.set('VCHTYPE', 'Purchase')
    vch.set('ACTION', 'Create')

    SubElement(vch, 'DATE').text = format_date(voucher.get('invoice_date'))
    SubElement(vch, 'VOUCHERTYPENAME').text = 'Purchase'
    SubElement(vch, 'VOUCHERNUMBER').text = voucher.get('name', '')
    SubElement(vch, 'REFERENCE').text = voucher.get('ref', '') or voucher.get('name', '')
    SubElement(vch, 'PARTYLEDGERNAME').text = voucher.get('tally_party_ledger', voucher.get('partner_name', ''))
    SubElement(vch, 'NARRATION').text = build_narration(voucher)
    SubElement(vch, 'ISINVOICE').text = 'Yes'
    SubElement(vch, 'PERSISTEDVIEW').text = 'Invoice Voucher View'

    if voucher.get('due_date'):
        SubElement(vch, 'EFFECTIVEDATE').text = format_date(voucher['due_date'])

    # --- Expense/Purchase ledger entries (Debit side) ---
    for line in voucher.get('mapped_lines', []):
        if line.get('price_subtotal', 0) == 0 and line.get('debit', 0) == 0:
            continue
        # Skip tax lines and receivable/payable lines
        account_name = line.get('account_name', '').lower()
        if any(kw in account_name for kw in ['receivable', 'payable', 'tax', 'gst', 'cgst', 'sgst', 'igst']):
            continue

        entry = SubElement(vch, 'ALLLEDGERENTRIES.LIST')
        SubElement(entry, 'LEDGERNAME').text = line.get('tally_ledger', line.get('account_name', ''))
        SubElement(entry, 'ISDEEMEDPOSITIVE').text = 'Yes'
        # In Tally XML: negative = debit for purchase
        amount = abs(line.get('price_subtotal', 0) or line.get('debit', 0))
        SubElement(entry, 'AMOUNT').text = f"-{amount:.2f}"

    # --- Tax ledger entries (Debit side) ---
    tax_totals = aggregate_taxes(voucher)
    for tax_ledger, tax_amount in tax_totals.items():
        if tax_amount == 0:
            continue
        entry = SubElement(vch, 'ALLLEDGERENTRIES.LIST')
        SubElement(entry, 'LEDGERNAME').text = tax_ledger
        SubElement(entry, 'ISDEEMEDPOSITIVE').text = 'Yes'
        SubElement(entry, 'AMOUNT').text = f"-{abs(tax_amount):.2f}"

    # --- Party ledger entry (Credit side) ---
    party_entry = SubElement(vch, 'ALLLEDGERENTRIES.LIST')
    SubElement(party_entry, 'LEDGERNAME').text = voucher.get('tally_party_ledger', voucher.get('partner_name', ''))
    SubElement(party_entry, 'ISDEEMEDPOSITIVE').text = 'No'
    SubElement(party_entry, 'AMOUNT').text = f"{abs(voucher.get('amount_total', 0)):.2f}"

    # Bill allocation for the party entry
    bill_alloc = SubElement(party_entry, 'BILLALLOCATIONS.LIST')
    SubElement(bill_alloc, 'NAME').text = voucher.get('name', '')
    SubElement(bill_alloc, 'BILLTYPE').text = 'New Ref'
    SubElement(bill_alloc, 'AMOUNT').text = f"{abs(voucher.get('amount_total', 0)):.2f}"

    return tally_msg


def build_sales_voucher(voucher, company_name):
    """Build Tally XML for a Sales voucher (Customer Invoice)."""
    tally_msg = Element('TALLYMESSAGE')
    tally_msg.set('xmlns:UDF', 'TallyUDF')

    vch = SubElement(tally_msg, 'VOUCHER')
    vch.set('VCHTYPE', 'Sales')
    vch.set('ACTION', 'Create')

    SubElement(vch, 'DATE').text = format_date(voucher.get('invoice_date'))
    SubElement(vch, 'VOUCHERTYPENAME').text = 'Sales'
    SubElement(vch, 'VOUCHERNUMBER').text = voucher.get('name', '')
    SubElement(vch, 'REFERENCE').text = voucher.get('ref', '') or voucher.get('name', '')
    SubElement(vch, 'PARTYLEDGERNAME').text = voucher.get('tally_party_ledger', voucher.get('partner_name', ''))
    SubElement(vch, 'NARRATION').text = build_narration(voucher)
    SubElement(vch, 'ISINVOICE').text = 'Yes'
    SubElement(vch, 'PERSISTEDVIEW').text = 'Invoice Voucher View'

    if voucher.get('due_date'):
        SubElement(vch, 'EFFECTIVEDATE').text = format_date(voucher['due_date'])

    # --- Party ledger entry (Debit side) ---
    party_entry = SubElement(vch, 'ALLLEDGERENTRIES.LIST')
    SubElement(party_entry, 'LEDGERNAME').text = voucher.get('tally_party_ledger', voucher.get('partner_name', ''))
    SubElement(party_entry, 'ISDEEMEDPOSITIVE').text = 'Yes'
    SubElement(party_entry, 'AMOUNT').text = f"-{abs(voucher.get('amount_total', 0)):.2f}"

    # Bill allocation
    bill_alloc = SubElement(party_entry, 'BILLALLOCATIONS.LIST')
    SubElement(bill_alloc, 'NAME').text = voucher.get('name', '')
    SubElement(bill_alloc, 'BILLTYPE').text = 'New Ref'
    SubElement(bill_alloc, 'AMOUNT').text = f"-{abs(voucher.get('amount_total', 0)):.2f}"

    # --- Sales/Income ledger entries (Credit side) ---
    for line in voucher.get('mapped_lines', []):
        if line.get('price_subtotal', 0) == 0 and line.get('credit', 0) == 0:
            continue
        account_name = line.get('account_name', '').lower()
        if any(kw in account_name for kw in ['receivable', 'payable', 'tax', 'gst', 'cgst', 'sgst', 'igst']):
            continue

        entry = SubElement(vch, 'ALLLEDGERENTRIES.LIST')
        SubElement(entry, 'LEDGERNAME').text = line.get('tally_ledger', line.get('account_name', ''))
        SubElement(entry, 'ISDEEMEDPOSITIVE').text = 'No'
        amount = abs(line.get('price_subtotal', 0) or line.get('credit', 0))
        SubElement(entry, 'AMOUNT').text = f"{amount:.2f}"

    # --- Tax ledger entries (Credit side) ---
    tax_totals = aggregate_taxes(voucher)
    for tax_ledger, tax_amount in tax_totals.items():
        if tax_amount == 0:
            continue
        entry = SubElement(vch, 'ALLLEDGERENTRIES.LIST')
        SubElement(entry, 'LEDGERNAME').text = tax_ledger
        SubElement(entry, 'ISDEEMEDPOSITIVE').text = 'No'
        SubElement(entry, 'AMOUNT').text = f"{abs(tax_amount):.2f}"

    return tally_msg


def build_credit_note_voucher(voucher, company_name):
    """Build Tally XML for a Credit Note (Customer Refund)."""
    tally_msg = Element('TALLYMESSAGE')
    tally_msg.set('xmlns:UDF', 'TallyUDF')

    vch = SubElement(tally_msg, 'VOUCHER')
    vch.set('VCHTYPE', 'Credit Note')
    vch.set('ACTION', 'Create')

    SubElement(vch, 'DATE').text = format_date(voucher.get('invoice_date'))
    SubElement(vch, 'VOUCHERTYPENAME').text = 'Credit Note'
    SubElement(vch, 'VOUCHERNUMBER').text = voucher.get('name', '')
    SubElement(vch, 'REFERENCE').text = voucher.get('ref', '') or voucher.get('name', '')
    SubElement(vch, 'PARTYLEDGERNAME').text = voucher.get('tally_party_ledger', voucher.get('partner_name', ''))
    SubElement(vch, 'NARRATION').text = build_narration(voucher)
    SubElement(vch, 'ISINVOICE').text = 'Yes'

    # Party (Credit side - reducing receivable)
    party_entry = SubElement(vch, 'ALLLEDGERENTRIES.LIST')
    SubElement(party_entry, 'LEDGERNAME').text = voucher.get('tally_party_ledger', voucher.get('partner_name', ''))
    SubElement(party_entry, 'ISDEEMEDPOSITIVE').text = 'No'
    SubElement(party_entry, 'AMOUNT').text = f"{abs(voucher.get('amount_total', 0)):.2f}"

    bill_alloc = SubElement(party_entry, 'BILLALLOCATIONS.LIST')
    SubElement(bill_alloc, 'NAME').text = voucher.get('ref', '') or voucher.get('name', '')
    SubElement(bill_alloc, 'BILLTYPE').text = 'Agst Ref'
    SubElement(bill_alloc, 'AMOUNT').text = f"{abs(voucher.get('amount_total', 0)):.2f}"

    # Sales return entries (Debit side)
    for line in voucher.get('mapped_lines', []):
        if line.get('price_subtotal', 0) == 0:
            continue
        account_name = line.get('account_name', '').lower()
        if any(kw in account_name for kw in ['receivable', 'payable', 'tax', 'gst', 'cgst', 'sgst', 'igst']):
            continue

        entry = SubElement(vch, 'ALLLEDGERENTRIES.LIST')
        SubElement(entry, 'LEDGERNAME').text = line.get('tally_ledger', line.get('account_name', ''))
        SubElement(entry, 'ISDEEMEDPOSITIVE').text = 'Yes'
        SubElement(entry, 'AMOUNT').text = f"-{abs(line.get('price_subtotal', 0)):.2f}"

    # Tax entries
    tax_totals = aggregate_taxes(voucher)
    for tax_ledger, tax_amount in tax_totals.items():
        if tax_amount == 0:
            continue
        entry = SubElement(vch, 'ALLLEDGERENTRIES.LIST')
        SubElement(entry, 'LEDGERNAME').text = tax_ledger
        SubElement(entry, 'ISDEEMEDPOSITIVE').text = 'Yes'
        SubElement(entry, 'AMOUNT').text = f"-{abs(tax_amount):.2f}"

    return tally_msg


def build_debit_note_voucher(voucher, company_name):
    """Build Tally XML for a Debit Note (Vendor Refund)."""
    tally_msg = Element('TALLYMESSAGE')
    tally_msg.set('xmlns:UDF', 'TallyUDF')

    vch = SubElement(tally_msg, 'VOUCHER')
    vch.set('VCHTYPE', 'Debit Note')
    vch.set('ACTION', 'Create')

    SubElement(vch, 'DATE').text = format_date(voucher.get('invoice_date'))
    SubElement(vch, 'VOUCHERTYPENAME').text = 'Debit Note'
    SubElement(vch, 'VOUCHERNUMBER').text = voucher.get('name', '')
    SubElement(vch, 'REFERENCE').text = voucher.get('ref', '') or voucher.get('name', '')
    SubElement(vch, 'PARTYLEDGERNAME').text = voucher.get('tally_party_ledger', voucher.get('partner_name', ''))
    SubElement(vch, 'NARRATION').text = build_narration(voucher)
    SubElement(vch, 'ISINVOICE').text = 'Yes'

    # Purchase return entries (Credit side)
    for line in voucher.get('mapped_lines', []):
        if line.get('price_subtotal', 0) == 0:
            continue
        account_name = line.get('account_name', '').lower()
        if any(kw in account_name for kw in ['receivable', 'payable', 'tax', 'gst', 'cgst', 'sgst', 'igst']):
            continue

        entry = SubElement(vch, 'ALLLEDGERENTRIES.LIST')
        SubElement(entry, 'LEDGERNAME').text = line.get('tally_ledger', line.get('account_name', ''))
        SubElement(entry, 'ISDEEMEDPOSITIVE').text = 'No'
        SubElement(entry, 'AMOUNT').text = f"{abs(line.get('price_subtotal', 0)):.2f}"

    # Tax entries
    tax_totals = aggregate_taxes(voucher)
    for tax_ledger, tax_amount in tax_totals.items():
        if tax_amount == 0:
            continue
        entry = SubElement(vch, 'ALLLEDGERENTRIES.LIST')
        SubElement(entry, 'LEDGERNAME').text = tax_ledger
        SubElement(entry, 'ISDEEMEDPOSITIVE').text = 'No'
        SubElement(entry, 'AMOUNT').text = f"{abs(tax_amount):.2f}"

    # Party (Debit side - reducing payable)
    party_entry = SubElement(vch, 'ALLLEDGERENTRIES.LIST')
    SubElement(party_entry, 'LEDGERNAME').text = voucher.get('tally_party_ledger', voucher.get('partner_name', ''))
    SubElement(party_entry, 'ISDEEMEDPOSITIVE').text = 'Yes'
    SubElement(party_entry, 'AMOUNT').text = f"-{abs(voucher.get('amount_total', 0)):.2f}"

    bill_alloc = SubElement(party_entry, 'BILLALLOCATIONS.LIST')
    SubElement(bill_alloc, 'NAME').text = voucher.get('ref', '') or voucher.get('name', '')
    SubElement(bill_alloc, 'BILLTYPE').text = 'Agst Ref'
    SubElement(bill_alloc, 'AMOUNT').text = f"-{abs(voucher.get('amount_total', 0)):.2f}"

    return tally_msg


def build_ledger_master(ledger_name, group_name, company_name):
    """Build Tally XML to create a ledger master (if it doesn't exist in Tally)."""
    tally_msg = Element('TALLYMESSAGE')
    tally_msg.set('xmlns:UDF', 'TallyUDF')

    ledger = SubElement(tally_msg, 'LEDGER')
    ledger.set('NAME', ledger_name)
    ledger.set('ACTION', 'Create')

    SubElement(ledger, 'NAME.LIST').text = ledger_name
    SubElement(ledger, 'PARENT').text = group_name
    SubElement(ledger, 'ISBILLWISEON').text = 'Yes' if group_name in ('Sundry Debtors', 'Sundry Creditors') else 'No'

    return tally_msg


def aggregate_taxes(voucher):
    """Aggregate tax amounts by Tally tax ledger across all lines."""
    tax_totals = {}
    for line in voucher.get('mapped_lines', []):
        for tax in line.get('tally_tax_ledgers', []):
            ledger = tax.get('ledger', '')
            if ledger:
                tax_totals[ledger] = tax_totals.get(ledger, 0) + tax.get('amount', 0)
    return tax_totals


def build_narration(voucher):
    """Build narration text for the Tally voucher."""
    parts = []
    parts.append(f"Odoo {voucher.get('move_type', '')}: {voucher.get('name', '')}")
    if voucher.get('origin'):
        parts.append(f"Origin: {voucher['origin']}")
    if voucher.get('ref'):
        parts.append(f"Ref: {voucher['ref']}")
    if voucher.get('narration'):
        parts.append(voucher['narration'])
    return ' | '.join(parts)


def prettify_xml(elem):
    """Return pretty-printed XML string."""
    rough_string = tostring(elem, encoding='unicode')
    try:
        parsed = parseString(rough_string)
        return parsed.toprettyxml(indent="  ", encoding=None)
    except Exception:
        return rough_string


def build_xml_files(mapped_file=None, single_file=False, create_masters=False):
    """Main function to build Tally XML files from mapped voucher data."""
    if mapped_file is None:
        mapped_file = MAPPED_FILE

    if not Path(mapped_file).exists():
        print(f"✗ Mapped data file not found: {mapped_file}")
        print("  Run: python execution/ledger_mapper.py --map .tmp/odoo_extract.json")
        return []

    with open(mapped_file, 'r', encoding='utf-8') as f:
        vouchers = json.load(f)

    if not vouchers:
        print("⚠ No vouchers to process")
        return []

    if not TALLY_COMPANY:
        print("✗ TALLY_COMPANY_NAME not set in .env")
        return []

    XML_DIR.mkdir(parents=True, exist_ok=True)

    # Build voucher type handlers
    builders = {
        'Purchase': build_purchase_voucher,
        'Sales': build_sales_voucher,
        'Credit Note': build_credit_note_voucher,
        'Debit Note': build_debit_note_voucher,
    }

    xml_files = []
    success = 0
    errors = 0

    if single_file:
        # All vouchers in one XML file
        envelope, request_data = build_envelope(TALLY_COMPANY)

        # Optionally create ledger masters first
        if create_masters:
            masters = collect_ledger_masters(vouchers)
            master_envelope, master_data = build_envelope(TALLY_COMPANY, "Ledgers")
            for master_msg in masters:
                master_data.append(master_msg)
            master_file = XML_DIR / "ledger_masters.xml"
            with open(master_file, 'w', encoding='utf-8') as f:
                f.write(prettify_xml(master_envelope))
            xml_files.append(str(master_file))
            print(f"✓ Ledger masters XML: {master_file} ({len(masters)} ledgers)")

        for v in vouchers:
            vch_type = v.get('tally_voucher_type', 'Journal')
            builder = builders.get(vch_type)
            if builder:
                try:
                    tally_msg = builder(v, TALLY_COMPANY)
                    request_data.append(tally_msg)
                    success += 1
                except Exception as e:
                    print(f"  ✗ Error building {v.get('name', '?')}: {e}")
                    errors += 1
            else:
                print(f"  ⚠ Unsupported voucher type: {vch_type} for {v.get('name', '?')}")
                errors += 1

        out_file = XML_DIR / f"vouchers_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xml"
        with open(out_file, 'w', encoding='utf-8') as f:
            f.write(prettify_xml(envelope))
        xml_files.append(str(out_file))

    else:
        # Separate XML file per voucher type
        grouped = {}
        for v in vouchers:
            vch_type = v.get('tally_voucher_type', 'Journal')
            grouped.setdefault(vch_type, []).append(v)

        if create_masters:
            masters = collect_ledger_masters(vouchers)
            master_envelope, master_data = build_envelope(TALLY_COMPANY, "Ledgers")
            for master_msg in masters:
                master_data.append(master_msg)
            master_file = XML_DIR / "ledger_masters.xml"
            with open(master_file, 'w', encoding='utf-8') as f:
                f.write(prettify_xml(master_envelope))
            xml_files.append(str(master_file))
            print(f"✓ Ledger masters XML: {master_file} ({len(masters)} ledgers)")

        for vch_type, type_vouchers in grouped.items():
            envelope, request_data = build_envelope(TALLY_COMPANY)
            type_success = 0

            builder = builders.get(vch_type)
            if not builder:
                print(f"  ⚠ Skipping unsupported type: {vch_type} ({len(type_vouchers)} vouchers)")
                errors += len(type_vouchers)
                continue

            for v in type_vouchers:
                try:
                    tally_msg = builder(v, TALLY_COMPANY)
                    request_data.append(tally_msg)
                    type_success += 1
                    success += 1
                except Exception as e:
                    print(f"  ✗ Error building {v.get('name', '?')}: {e}")
                    errors += 1

            safe_type = vch_type.replace(' ', '_').lower()
            out_file = XML_DIR / f"{safe_type}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xml"
            with open(out_file, 'w', encoding='utf-8') as f:
                f.write(prettify_xml(envelope))
            xml_files.append(str(out_file))
            print(f"✓ {vch_type}: {type_success} vouchers → {out_file}")

    print(f"\n--- XML Build Summary ---")
    print(f"  Success: {success}")
    print(f"  Errors: {errors}")
    print(f"  XML files: {len(xml_files)}")
    for xf in xml_files:
        print(f"    → {xf}")

    return xml_files


def collect_ledger_masters(vouchers):
    """Collect unique ledger names from vouchers and create master entries."""
    ledgers = {}  # name → group

    for v in vouchers:
        # Party ledger
        party = v.get('tally_party_ledger', '')
        party_group = v.get('tally_party_group', 'Sundry Debtors')
        if party:
            ledgers[party] = party_group

        # Line item ledgers
        for line in v.get('mapped_lines', []):
            ledger = line.get('tally_ledger', '')
            if ledger:
                # Guess group from context
                account_name = line.get('account_name', '').lower()
                if 'purchase' in account_name:
                    ledgers[ledger] = 'Purchase Accounts'
                elif 'sale' in account_name or 'revenue' in account_name:
                    ledgers[ledger] = 'Sales Accounts'
                elif 'expense' in account_name:
                    ledgers[ledger] = 'Indirect Expenses'
                else:
                    ledgers.setdefault(ledger, 'Suspense A/c')

            # Tax ledgers
            for tax in line.get('tally_tax_ledgers', []):
                tax_ledger = tax.get('ledger', '')
                if tax_ledger:
                    ledgers[tax_ledger] = 'Duties & Taxes'

    masters = []
    for name, group in ledgers.items():
        masters.append(build_ledger_master(name, group, TALLY_COMPANY))

    return masters


def main():
    parser = argparse.ArgumentParser(description="Tally XML Builder")
    parser.add_argument('--input', type=str, default=str(MAPPED_FILE),
                        help='Path to mapped vouchers JSON file')
    parser.add_argument('--single-file', action='store_true',
                        help='Output all vouchers in a single XML file')
    parser.add_argument('--create-masters', action='store_true',
                        help='Also generate ledger master creation XML')
    args = parser.parse_args()

    build_xml_files(
        mapped_file=args.input,
        single_file=args.single_file,
        create_masters=args.create_masters,
    )


if __name__ == "__main__":
    main()
