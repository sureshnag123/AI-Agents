"""
Tally XML Pusher: Sends generated Tally XML files to Tally Prime's HTTP XML Server.

Usage:
    python execution/tally_xml_pusher.py                          # Push all XML files in .tmp/tally_xml/
    python execution/tally_xml_pusher.py --file path/to/file.xml  # Push specific file
    python execution/tally_xml_pusher.py --dry-run                # Validate without pushing
    python execution/tally_xml_pusher.py --masters-first          # Push ledger masters before vouchers
"""

import os
import sys
import csv
import json
import requests
import argparse
from datetime import datetime
from pathlib import Path
from xml.etree.ElementTree import parse as parse_xml
from dotenv import load_dotenv

load_dotenv()

# --- Configuration ---
TALLY_HOST = os.getenv("TALLY_HOST", "localhost")
TALLY_PORT = os.getenv("TALLY_PORT", "9000")
TALLY_COMPANY = os.getenv("TALLY_COMPANY_NAME", "")
TALLY_URL = f"http://{TALLY_HOST}:{TALLY_PORT}"

TMP_DIR = Path(".tmp")
XML_DIR = TMP_DIR / "tally_xml"
SYNC_LOG_FILE = TMP_DIR / "sync_log.csv"
ERROR_LOG_FILE = TMP_DIR / "sync_errors.log"


def check_tally_connection():
    """Check if Tally Prime XML Server is running and accessible."""
    try:
        # Send a simple request to check if Tally is responding
        test_xml = """<ENVELOPE>
  <HEADER>
    <TALLYREQUEST>Export Data</TALLYREQUEST>
  </HEADER>
  <BODY>
    <EXPORTDATA>
      <REQUESTDESC>
        <REPORTNAME>List of Companies</REPORTNAME>
      </REQUESTDESC>
    </EXPORTDATA>
  </BODY>
</ENVELOPE>"""

        response = requests.post(
            TALLY_URL,
            data=test_xml,
            headers={'Content-Type': 'application/xml'},
            timeout=10
        )

        if response.status_code == 200:
            # Check if response contains company info
            if 'ENVELOPE' in response.text or 'COMPANY' in response.text:
                print(f"✓ Tally Prime is running at {TALLY_URL}")
                if TALLY_COMPANY:
                    if TALLY_COMPANY in response.text:
                        print(f"  ✓ Company '{TALLY_COMPANY}' found in Tally")
                    else:
                        print(f"  ⚠ Company '{TALLY_COMPANY}' not found in response. Make sure it's open in Tally.")
                return True
            else:
                print(f"⚠ Got response from {TALLY_URL} but it doesn't look like Tally")
                return False
        else:
            print(f"✗ Tally returned status {response.status_code}")
            return False

    except requests.ConnectionError:
        print(f"✗ Cannot connect to Tally at {TALLY_URL}")
        print(f"  Make sure Tally Prime is running and XML Server is enabled:")
        print(f"  Gateway of Tally → F12 → Connectivity → Tally.NET → Enable XML Server = Yes")
        return False
    except requests.Timeout:
        print(f"✗ Connection to Tally at {TALLY_URL} timed out")
        return False


def push_xml_to_tally(xml_content, filename=""):
    """
    Send XML content to Tally Prime and return the response.

    Returns:
        dict with keys: success (bool), response_text (str), errors (list)
    """
    result = {
        'success': False,
        'response_text': '',
        'errors': [],
        'created': 0,
        'altered': 0,
        'failed': 0,
    }

    try:
        response = requests.post(
            TALLY_URL,
            data=xml_content.encode('utf-8'),
            headers={'Content-Type': 'application/xml; charset=utf-8'},
            timeout=60
        )

        result['response_text'] = response.text

        if response.status_code == 200:
            # Parse Tally's response for success/failure indicators
            resp_text = response.text.upper()

            if 'CREATED' in resp_text:
                # Try to extract count
                import re
                created_match = re.search(r'CREATED\s*=\s*"?(\d+)"?', response.text, re.IGNORECASE)
                if created_match:
                    result['created'] = int(created_match.group(1))

                altered_match = re.search(r'ALTERED\s*=\s*"?(\d+)"?', response.text, re.IGNORECASE)
                if altered_match:
                    result['altered'] = int(altered_match.group(1))

                result['success'] = True

            elif 'ERROR' in resp_text or 'LINEERROR' in resp_text:
                # Extract error messages
                import re
                error_matches = re.findall(r'<LINEERROR>(.*?)</LINEERROR>', response.text, re.IGNORECASE)
                result['errors'] = error_matches
                result['success'] = len(error_matches) == 0

            else:
                # No clear success/error indicators - assume success if 200
                result['success'] = True

        else:
            result['errors'].append(f"HTTP {response.status_code}: {response.text[:200]}")

    except requests.ConnectionError as e:
        result['errors'].append(f"Connection error: {e}")
    except requests.Timeout:
        result['errors'].append("Request timed out (60s)")
    except Exception as e:
        result['errors'].append(f"Unexpected error: {e}")

    return result


def push_xml_files(xml_files=None, dry_run=False, masters_first=False):
    """
    Push XML files to Tally Prime.

    Args:
        xml_files: List of XML file paths. If None, uses all files in XML_DIR.
        dry_run: If True, validate files without pushing.
        masters_first: If True, push ledger_masters.xml before voucher files.
    """
    if xml_files is None:
        if not XML_DIR.exists():
            print(f"✗ XML directory not found: {XML_DIR}")
            print("  Run: python execution/tally_xml_builder.py first")
            return
        xml_files = sorted(XML_DIR.glob("*.xml"))
        if not xml_files:
            print(f"⚠ No XML files found in {XML_DIR}")
            return

    # Sort: masters first if requested
    if masters_first:
        master_files = [f for f in xml_files if 'master' in str(f).lower()]
        other_files = [f for f in xml_files if 'master' not in str(f).lower()]
        xml_files = master_files + other_files

    print(f"\n{'=== DRY RUN ===' if dry_run else '=== Pushing to Tally ==='}")
    print(f"  Target: {TALLY_URL}")
    print(f"  Company: {TALLY_COMPANY}")
    print(f"  Files: {len(xml_files)}")
    print()

    if not dry_run:
        if not check_tally_connection():
            print("\n✗ Aborting: Cannot connect to Tally")
            return

    # Initialize sync log
    sync_results = []
    total_created = 0
    total_errors = 0

    for xml_file in xml_files:
        xml_path = Path(xml_file)
        print(f"\n--- {xml_path.name} ---")

        try:
            with open(xml_path, 'r', encoding='utf-8') as f:
                xml_content = f.read()
        except Exception as e:
            print(f"  ✗ Cannot read file: {e}")
            total_errors += 1
            continue

        # Basic validation
        if '<ENVELOPE>' not in xml_content:
            print(f"  ✗ Invalid XML: missing <ENVELOPE> tag")
            total_errors += 1
            continue

        # Count vouchers in file
        voucher_count = xml_content.count('<VOUCHER ')
        ledger_count = xml_content.count('<LEDGER ')
        print(f"  Contains: {voucher_count} vouchers, {ledger_count} ledger masters")

        if dry_run:
            print(f"  ✓ Validation passed (dry run, not pushing)")
            # Extract voucher numbers for the log
            import re
            vch_numbers = re.findall(r'<VOUCHERNUMBER>(.*?)</VOUCHERNUMBER>', xml_content)
            for vn in vch_numbers:
                sync_results.append({
                    'timestamp': datetime.now().isoformat(),
                    'file': xml_path.name,
                    'odoo_invoice_name': vn,
                    'status': 'dry_run',
                    'message': 'Validated, not pushed',
                })
            continue

        # Push to Tally
        result = push_xml_to_tally(xml_content, xml_path.name)

        if result['success']:
            print(f"  ✓ Success: {result['created']} created, {result['altered']} altered")
            total_created += result['created'] + result['altered']

            # Log synced vouchers
            import re
            vch_numbers = re.findall(r'<VOUCHERNUMBER>(.*?)</VOUCHERNUMBER>', xml_content)
            for vn in vch_numbers:
                sync_results.append({
                    'timestamp': datetime.now().isoformat(),
                    'file': xml_path.name,
                    'odoo_invoice_name': vn,
                    'status': 'synced',
                    'message': f'Created: {result["created"]}, Altered: {result["altered"]}',
                })
        else:
            print(f"  ✗ Failed:")
            for err in result['errors']:
                print(f"    - {err}")
            total_errors += 1

            # Log errors
            import re
            vch_numbers = re.findall(r'<VOUCHERNUMBER>(.*?)</VOUCHERNUMBER>', xml_content)
            for vn in vch_numbers:
                sync_results.append({
                    'timestamp': datetime.now().isoformat(),
                    'file': xml_path.name,
                    'odoo_invoice_name': vn,
                    'status': 'error',
                    'message': '; '.join(result['errors'][:3]),
                })

            # Write to error log
            with open(ERROR_LOG_FILE, 'a', encoding='utf-8') as f:
                f.write(f"\n{'='*60}\n")
                f.write(f"Timestamp: {datetime.now().isoformat()}\n")
                f.write(f"File: {xml_path.name}\n")
                f.write(f"Errors: {result['errors']}\n")
                f.write(f"Response:\n{result['response_text'][:500]}\n")

    # Update sync log
    update_sync_log(sync_results)

    # Summary
    print(f"\n{'='*40}")
    print(f"{'DRY RUN ' if dry_run else ''}SYNC SUMMARY")
    print(f"{'='*40}")
    print(f"  Files processed: {len(xml_files)}")
    print(f"  Vouchers logged: {len(sync_results)}")
    if not dry_run:
        print(f"  Successfully created/altered: {total_created}")
        print(f"  Files with errors: {total_errors}")

    return sync_results


def update_sync_log(results):
    """Append sync results to the sync log CSV."""
    TMP_DIR.mkdir(exist_ok=True)

    file_exists = SYNC_LOG_FILE.exists()
    fieldnames = ['timestamp', 'file', 'odoo_invoice_name', 'status', 'message']

    with open(SYNC_LOG_FILE, 'a', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        if not file_exists:
            writer.writeheader()
        writer.writerows(results)

    print(f"  Sync log updated: {SYNC_LOG_FILE}")


def view_sync_log(last_n=20):
    """Display recent sync log entries."""
    if not SYNC_LOG_FILE.exists():
        print("No sync log found yet.")
        return

    with open(SYNC_LOG_FILE, 'r', encoding='utf-8') as f:
        rows = list(csv.DictReader(f))

    if not rows:
        print("Sync log is empty.")
        return

    recent = rows[-last_n:]
    print(f"\n--- Last {len(recent)} sync entries ---")
    for row in recent:
        status_icon = '✓' if row['status'] == 'synced' else '✗' if row['status'] == 'error' else '○'
        print(f"  {status_icon} [{row['timestamp'][:19]}] {row['odoo_invoice_name']} → {row['status']}")

    # Stats
    synced = sum(1 for r in rows if r['status'] == 'synced')
    errors = sum(1 for r in rows if r['status'] == 'error')
    print(f"\n  Total synced: {synced} | Total errors: {errors}")


def main():
    parser = argparse.ArgumentParser(description="Tally XML Pusher")
    parser.add_argument('--file', type=str, help='Push a specific XML file')
    parser.add_argument('--dry-run', action='store_true', help='Validate without pushing')
    parser.add_argument('--masters-first', action='store_true',
                        help='Push ledger masters before vouchers')
    parser.add_argument('--check', action='store_true', help='Check Tally connection only')
    parser.add_argument('--log', action='store_true', help='View sync log')
    parser.add_argument('--log-last', type=int, default=20, help='Number of log entries to show')
    args = parser.parse_args()

    if args.check:
        check_tally_connection()
    elif args.log:
        view_sync_log(args.log_last)
    elif args.file:
        push_xml_files(xml_files=[args.file], dry_run=args.dry_run, masters_first=args.masters_first)
    else:
        push_xml_files(dry_run=args.dry_run, masters_first=args.masters_first)


if __name__ == "__main__":
    main()
