"""
GSTR2B Reconciliation Agent - Configuration
"""
import os
from pathlib import Path

# Base directory
BASE_DIR = Path(os.path.dirname(os.path.abspath(__file__)))

# Output directory
OUTPUT_DIR = BASE_DIR / "output"
LOG_DIR = BASE_DIR / "logs"
SAMPLE_DIR = BASE_DIR / "sample_data"

# Matching tolerances
FUZZY_MATCH_THRESHOLD = 80          # Minimum score for fuzzy vendor name match (0-100)
DATE_TOLERANCE_DAYS = 3             # ± days tolerance for invoice date matching
AMOUNT_TOLERANCE_PERCENT = 1.0      # % tolerance for amount matching
AMOUNT_TOLERANCE_ABSOLUTE = 1.0     # ₹ absolute tolerance for rounding differences

# GSTIN format regex
GSTIN_REGEX = r"^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[A-Z0-9]{1}Z[A-Z0-9]{1}$"

# Invoice number cleaning - characters to strip
INVOICE_STRIP_CHARS = [" ", "-", "/", "\\", ".", "#", "'", '"']

# Tally expected columns (flexible mapping)
TALLY_COLUMN_MAP = {
    "voucher_type": ["Voucher Type", "Vch Type", "Type"],
    "voucher_date": ["Voucher Date", "Date", "Vch Date"],
    "voucher_number": ["Voucher Number", "Vch No", "Voucher No"],
    "party_name": ["Party Name", "Vendor Name", "Ledger Name", "Party"],
    "gstin": ["GSTIN", "GSTIN/UIN", "GST No", "Supplier GSTIN"],
    "bill_number": ["Vendor Bill Number", "Reference Number", "Ref No", "Bill No", "Invoice No", "Invoice Number"],
    "taxable_value": ["Taxable Value", "Taxable Amount", "Assessable Value"],
    "igst": ["IGST Amount", "IGST", "Integrated Tax"],
    "cgst": ["CGST Amount", "CGST", "Central Tax"],
    "sgst": ["SGST Amount", "SGST", "State Tax"],
    "total_value": ["Total Invoice Value", "Total Value", "Invoice Value", "Grand Total"],
    "narration": ["Narration", "Remarks", "Description"],
}

# GSTR2B expected columns
# Includes both standard names and actual GST portal column names (with ₹ suffix)
GSTR2B_COLUMN_MAP = {
    "gstin": ["Supplier GSTIN", "GSTIN of supplier", "GSTIN"],
    "supplier_name": ["Supplier Name", "Trade/Legal name", "Trade Name", "Trade/Legal name of the Supplier"],
    "invoice_number": ["Invoice Number", "Invoice No", "Inv No", "Document Number", "Invoice number"],
    "invoice_date": ["Invoice Date", "Inv Date", "Document Date"],
    "taxable_value": ["Taxable Value", "Taxable Amount", "Taxable Value (₹)", "Taxable Value(₹)"],
    "igst": ["IGST", "Integrated Tax", "IGST Amount", "Integrated Tax(₹)", "Integrated Tax (₹)"],
    "cgst": ["CGST", "Central Tax", "CGST Amount", "Central Tax(₹)", "Central Tax (₹)"],
    "sgst": ["SGST", "State Tax", "SGST Amount", "SGST/UTGST", "State/UT Tax", "State/UT Tax(₹)", "State/UT Tax (₹)"],
    "total_tax": ["Total Tax", "Tax Amount", "Total Tax(₹)", "Total Tax (₹)"],
}

# Voucher types to include from Tally / Odoo
TALLY_VOUCHER_TYPES = ["Purchase", "Journal", "Debit Note"]

# ---------------------------------------------------------------------------
# Odoo Connection Defaults (override in UI — do not store passwords here)
# ---------------------------------------------------------------------------
ODOO_URL = ""           # e.g. "https://mycompany.odoo.com"
ODOO_DB = ""            # e.g. "mycompany"
ODOO_USERNAME = ""      # e.g. "admin@mycompany.com"

# Report sheet names
REPORT_SHEETS = {
    "summary": "Reconciliation Summary",
    "in_tally_not_gstr2b": "In Tally NOT in GSTR2B",
    "in_gstr2b_not_tally": "In GSTR2B NOT in Tally",
    "no_bill_no_payment": "GSTR2B No Bill-Payment",
    "matched": "Matched Entries",
    "amount_mismatch": "Amount Mismatch",
    "duplicates": "Duplicate Entries",
}
