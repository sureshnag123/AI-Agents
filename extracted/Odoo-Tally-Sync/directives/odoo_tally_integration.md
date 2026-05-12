# Odoo 19 ↔ Tally Prime Integration

## Goal
Eliminate double-entry of accounting transactions by automatically syncing vouchers from Odoo 19 to Tally Prime. Odoo remains the operational ERP (Purchase Orders, Sales, Inventory, Manufacturing), while Tally Prime remains the statutory books of accounts preferred by auditors.

## Problem Statement

| System | Current Usage | Pain Point |
|--------|--------------|------------|
| **Odoo 19** | PO → Tax Invoice, Quotation → Sales Invoice, Inventory, Manufacturing | Full operational ERP but accounting not fully implemented |
| **Tally Prime** | Purchase & Sales (duplicated from Odoo), Cash, Bank, Journal, all statutory reporting | Auditor-preferred; all vouchers entered manually |
| **Double Entry** | Purchase invoices & Sales invoices entered in BOTH systems | ~2× time spent on data entry, risk of mismatches |

### What Gets Synced (Odoo → Tally)

| Voucher Type | Source in Odoo | Tally Voucher Type | Frequency |
|-------------|----------------|-------------------|-----------|
| Purchase Invoice | `account.move` (in_invoice) | **Purchase** | Daily/Weekly |
| Purchase Return (Debit Note) | `account.move` (in_refund) | **Debit Note** | Daily/Weekly |
| Sales Invoice | `account.move` (out_invoice) | **Sales** | Daily/Weekly |
| Sales Return (Credit Note) | `account.move` (out_refund) | **Credit Note** | Daily/Weekly |
| Payments Made | `account.payment` (outbound) | **Payment** | Daily/Weekly |
| Payments Received | `account.payment` (inbound) | **Receipt** | Daily/Weekly |
| Inventory Movements | `stock.move` (done) | **Stock Journal** | Weekly |

### What Stays Only in Tally (No Sync Needed)

| Entry Type | Reason |
|-----------|--------|
| Cash vouchers | Only entered in Tally |
| Bank vouchers | Only entered in Tally |
| Journal entries | Only entered in Tally |
| Contra entries | Only entered in Tally |
| Statutory adjustments | Auditor-managed |

## Architecture

```
┌─────────────────────────────────────────────────────────────────┐
│                    Orchestration Layer                           │
│                (You / Copilot / Scheduler)                       │
│                                                                 │
│   1. Read this directive                                        │
│   2. Decide: full sync? incremental? specific voucher type?     │
│   3. Call scripts in order                                      │
│   4. Review sync report, handle errors                          │
└──────────┬──────────────────┬──────────────────┬────────────────┘
           │                  │                  │
 ┌─────────▼────────┐ ┌──────▼───────────┐ ┌────▼──────────────┐
 │ odoo_connector.py │ │ tally_xml_builder│ │ tally_connector.py│
 │ (XML-RPC to Odoo) │ │ (Odoo→Tally XML) │ │ (HTTP to Tally)   │
 └──────────────────┘ └──────────────────┘ └───────────────────┘
           │                  │                  │
 ┌─────────▼──────────────────▼──────────────────▼────────────────┐
 │                    odoo_tally_sync.py                           │
 │         (Main pipeline: fetch → transform → push)              │
 └──────────────────────┬─────────────────┬───────────────────────┘
                        │                 │
              ┌─────────▼──────┐  ┌───────▼──────────┐
              │ sync_report.py │  │ sync_scheduler.py │
              │ (Status logs)  │  │ (Daily/weekly)    │
              └────────────────┘  └──────────────────┘
```

## Inputs

| Input | Required | Description |
|-------|----------|-------------|
| Odoo URL | Yes | Odoo 19 instance URL |
| Odoo DB | Yes | Database name |
| Odoo Username | Yes | Login email |
| Odoo Password/API Key | Yes | API key recommended |
| Tally Host | Yes | Tally Prime server IP (default: `localhost`) |
| Tally Port | Yes | Tally XML server port (default: `9000`) |
| Company Name (Tally) | Yes | Exact company name as in Tally Prime |
| Sync Mode | No | `full` or `incremental` (default: incremental) |
| Voucher Types | No | Comma-separated list (default: all) |
| Date Range | No | `--from YYYY-MM-DD --to YYYY-MM-DD` |

## Tally Prime Setup (One-Time)

### Enable XML Server in Tally Prime
1. Open Tally Prime
2. Press **F12** (Configure) → **Connectivity**
3. Set **Tally Prime Server** → **Yes**
4. Set **Port Number** → `9000` (default) or your preferred port
5. Set **Allow Remote Access** → **Yes** (if Tally runs on a different machine)
6. Accept and restart Tally

### Verify Tally XML Server
```bash
python execution/tally_connector.py --test
```
This sends a test XML request to Tally and confirms connectivity.

## Ledger & Account Mapping

### Why Mapping Matters
Odoo and Tally may use different names for the same ledger/account. The sync tool uses a mapping file (`.tmp/ledger_mapping.json`) to translate.

### Auto-Discovery
```bash
# Fetch all ledgers from Tally
python execution/tally_connector.py --list-ledgers

# Fetch all accounts from Odoo  
python execution/odoo_connector.py --list-accounts

# Generate initial mapping (fuzzy match)
python execution/odoo_tally_sync.py --generate-mapping
```

### Manual Mapping Override
Edit `.tmp/ledger_mapping.json`:
```json
{
  "odoo_to_tally_ledgers": {
    "Trade Payables": "Sundry Creditors",
    "Trade Receivables": "Sundry Debtors",
    "GST Input CGST": "Input CGST",
    "GST Input SGST": "Input SGST",
    "GST Input IGST": "Input IGST",
    "GST Output CGST": "Output CGST",
    "GST Output SGST": "Output SGST", 
    "GST Output IGST": "Output IGST",
    "Purchase of Goods": "Purchase Accounts",
    "Sales Revenue": "Sales Accounts",
    "Bank": "Bank Accounts",
    "Cash": "Cash-in-Hand"
  },
  "odoo_to_tally_partners": {
    "Vendor ABC Pvt Ltd": "ABC Pvt Ltd",
    "Customer XYZ Ltd": "XYZ Ltd"
  },
  "auto_create_missing_ledgers": true,
  "default_group_for_new_creditors": "Sundry Creditors",
  "default_group_for_new_debtors": "Sundry Debtors"
}
```

## Execution

### Step 1: Test Both Connections
```bash
# Test Odoo
python execution/odoo_connector.py --test

# Test Tally
python execution/tally_connector.py --test
```

### Step 2: Generate Ledger Mapping
```bash
python execution/odoo_tally_sync.py --generate-mapping
```
Review the generated mapping file at `.tmp/ledger_mapping.json` and fix any mismatches.

### Step 3: Dry Run (Preview What Will Sync)
```bash
# Preview all voucher types for last 7 days
python execution/odoo_tally_sync.py --dry-run

# Preview specific types
python execution/odoo_tally_sync.py --dry-run --types "purchase,sales"

# Preview specific date range
python execution/odoo_tally_sync.py --dry-run --from 2026-02-01 --to 2026-02-26
```

### Step 4: Run Actual Sync
```bash
# Sync all voucher types (incremental — only unsynced entries)
python execution/odoo_tally_sync.py

# Sync specific types
python execution/odoo_tally_sync.py --types "purchase,sales"

# Full sync (re-sync everything in date range)
python execution/odoo_tally_sync.py --mode full --from 2026-02-01 --to 2026-02-26

# Daily sync (last 24 hours)
python execution/odoo_tally_sync.py --mode incremental
```

### Step 5: Review Sync Report
```bash
python execution/sync_report.py                    # Console summary
python execution/sync_report.py --detailed          # Line-by-line details
python execution/sync_report.py --csv               # Export to CSV
python execution/sync_report.py --failed-only       # Show only failures
```

### Step 6: Schedule Automated Sync
```bash
# Daily at 8 PM (after business hours)
python execution/sync_scheduler.py --hour 20 --minute 0

# Weekly on Saturday at 10 AM
python execution/sync_scheduler.py --weekly --day saturday --hour 10

# One-shot (for Windows Task Scheduler / cron)
python execution/sync_scheduler.py --once
```

## GST Handling

Indian GST is critical for both systems. The sync handles:

| GST Component | Odoo Field | Tally Ledger |
|--------------|-----------|-------------|
| CGST | Tax line with CGST tag | Input CGST / Output CGST |
| SGST | Tax line with SGST tag | Input SGST / Output SGST |
| IGST | Tax line with IGST tag | Input IGST / Output IGST |
| Cess | Tax line with Cess tag | Cess on GST |

### Tax Rate Mapping
```json
{
  "gst_5": {"cgst": 2.5, "sgst": 2.5, "igst": 5},
  "gst_12": {"cgst": 6, "sgst": 6, "igst": 12},
  "gst_18": {"cgst": 9, "sgst": 9, "igst": 18},
  "gst_28": {"cgst": 14, "sgst": 14, "igst": 28}
}
```

## Sync State Management

The system tracks what has been synced to prevent duplicates:

- **Sync log**: `.tmp/sync_log.json` — records every synced voucher with Odoo ID, Tally master ID, timestamp, status
- **Incremental mode**: Only fetches Odoo records created/modified after the last sync timestamp
- **Duplicate detection**: Checks Tally for existing voucher with same reference number before importing
- **Conflict resolution**: If a voucher exists in Tally with different amounts, the system flags it for manual review

## Edge Cases & Learnings

- **Partner not in Tally**: If `auto_create_missing_ledgers` is true, the script creates the ledger in Tally under the appropriate group (Sundry Creditors/Debtors). Otherwise, it skips with a warning.
- **Tax mismatch**: If Odoo has a tax configuration the mapper doesn't recognize, it flags the voucher and skips it.
- **Multi-currency**: Tally handles forex differently. For foreign currency invoices, ensure the currency and exchange rate are mapped.
- **Tally not running**: The sync will fail gracefully with a clear error. Tally Prime must be open with the target company loaded.
- **Odoo draft invoices**: Only `posted` (confirmed) invoices are synced. Drafts are ignored.
- **Partial payments**: Payment vouchers sync the actual payment amount, not the invoice total.
- **Credit notes / Debit notes**: Mapped to Tally's Credit Note and Debit Note voucher types respectively.
- **Date format**: Odoo uses `YYYY-MM-DD`, Tally expects `YYYYMMDD` — the builder handles conversion.
- **Large batches**: For initial full sync, the script processes in batches of 100 vouchers to avoid timeouts.
- **Voucher numbering**: Tally auto-generates voucher numbers. Odoo's invoice number is stored in the Narration/Reference field for cross-reference.

## Rollback

If a sync batch goes wrong:
```bash
# Show last sync batch
python execution/sync_report.py --last-batch

# Delete last batch from Tally (requires Tally to be open)
python execution/odoo_tally_sync.py --rollback --batch-id BATCH_2026_02_26_001
```

## Security Notes

- Use Odoo **API keys** instead of passwords (Settings → Users → API Keys)
- Tally XML server has **no authentication** by default — restrict access via firewall rules
- Consider running Tally on `localhost` only and using SSH tunnels for remote access
- Never commit `.env` to version control
