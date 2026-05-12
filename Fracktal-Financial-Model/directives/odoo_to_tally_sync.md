# Directive: Odoo 19 → Tally Prime Sync

## Problem Statement

Fracktal is running **Odoo 19** for operational workflows (Purchase Orders → Tax Invoices, Quotations → Sales Invoices, Inventory, Manufacturing Orders) but auditors require **Tally Prime** for accounting & finance. This results in:

- **Double data entry**: Purchase & Sales vouchers entered in both systems
- **All other vouchers** (Cash, Bank, Journal) entered manually in Tally only
- **Significant time waste** managing both systems simultaneously

## Solution Overview

Build an automated **Odoo → Tally Prime sync pipeline** that:

1. Extracts validated invoices/bills from Odoo 19 via XML-RPC API
2. Maps Odoo accounts, partners, taxes → Tally Ledgers & Groups
3. Generates Tally-compatible XML (Tally's native import format)
4. Pushes XML to Tally Prime's HTTP XML Server (port 9000) or saves as importable files
5. Runs on a schedule (daily or weekly) with deduplication

## Architecture

```
┌──────────────┐     XML-RPC      ┌──────────────────┐
│   Odoo 19    │ ──────────────► │  odoo_extractor   │
│  (Source)    │                  │  (Python Script)  │
└──────────────┘                  └────────┬─────────┘
                                           │ JSON data
                                           ▼
                                  ┌──────────────────┐
                                  │  ledger_mapper    │
                                  │  (CSV + Python)   │
                                  └────────┬─────────┘
                                           │ Mapped data
                                           ▼
                                  ┌──────────────────┐
                                  │ tally_xml_builder │
                                  │  (Python Script)  │
                                  └────────┬─────────┘
                                           │ Tally XML
                                           ▼
                                  ┌──────────────────┐     HTTP POST
                                  │ tally_xml_pusher  │ ──────────────►  Tally Prime
                                  │  (Python Script)  │                  (Port 9000)
                                  └──────────────────┘
```

## Voucher Types to Sync (from Odoo → Tally)

| # | Odoo Source             | Tally Voucher Type | Priority |
|---|-------------------------|--------------------|----------|
| 1 | Vendor Bills (Posted)   | Purchase           | HIGH     |
| 2 | Customer Invoices (Posted) | Sales           | HIGH     |
| 3 | Credit Notes (Customer) | Credit Note        | MEDIUM   |
| 4 | Debit Notes (Vendor)    | Debit Note         | MEDIUM   |
| 5 | Payments (Outgoing)     | Payment            | LOW      |
| 6 | Receipts (Incoming)     | Receipt            | LOW      |

> **Phase 1**: Focus on Purchase & Sales vouchers (the biggest time-savers)
> **Phase 2**: Add Credit/Debit Notes
> **Phase 3**: Add Payment/Receipt vouchers

## Inputs

| Input | Source | Details |
|-------|--------|---------|
| Odoo credentials | `.env` | `ODOO_URL`, `ODOO_DB`, `ODOO_USERNAME`, `ODOO_PASSWORD` |
| Tally connection | `.env` | `TALLY_HOST` (default: localhost), `TALLY_PORT` (default: 9000) |
| Ledger mapping | `.tmp/ledger_mapping.csv` | Maps Odoo account names → Tally ledger names |
| Company name | `.env` | `TALLY_COMPANY_NAME` - must match exactly in Tally |
| Sync config | `.env` | `SYNC_MODE` (daily/weekly), `SYNC_DAYS_BACK` (default: 7) |

## Outputs

| Output | Location | Details |
|--------|----------|---------|
| Tally XML files | `.tmp/tally_xml/` | One XML file per batch |
| Sync log | `.tmp/sync_log.csv` | Tracks what was synced (deduplication) |
| Error log | `.tmp/sync_errors.log` | Failed vouchers with reasons |

## Execution Scripts

| Script | Purpose |
|--------|---------|
| `execution/odoo_extractor.py` | Connect to Odoo, extract invoices/bills |
| `execution/ledger_mapper.py` | Map Odoo accounts to Tally ledgers, generate/update mapping CSV |
| `execution/tally_xml_builder.py` | Convert mapped data → Tally XML format |
| `execution/tally_xml_pusher.py` | Push XML to Tally Prime HTTP server |
| `execution/odoo_tally_sync.py` | Master orchestrator: runs full pipeline |

## Step-by-Step Workflow

### Step 1: Initial Setup (One-time)

1. Configure `.env` with Odoo & Tally connection details
2. Run `ledger_mapper.py --init` to auto-generate mapping CSV from Odoo Chart of Accounts
3. **Manually review** `.tmp/ledger_mapping.csv` and fill in Tally ledger names
4. Ensure Tally Prime is running with XML Server enabled (Gateway of Tally → F12 → Connectivity → Tally.NET → Enable XML Server = Yes)

### Step 2: Extract from Odoo

1. Connect to Odoo 19 via XML-RPC
2. Query `account.move` for posted invoices/bills within date range
3. For each invoice, fetch: line items, taxes, partner details, payment terms
4. Filter out already-synced entries (check sync_log.csv)
5. Save extracted data as JSON in `.tmp/odoo_extract.json`

### Step 3: Map Ledgers

1. Load ledger mapping from `.tmp/ledger_mapping.csv`
2. For each extracted entry, map:
   - Partner name → Party Ledger in Tally
   - Account name → Ledger in Tally
   - Tax name → Tax Ledger in Tally
3. Flag any unmapped entries for manual review
4. Save mapped data to `.tmp/mapped_vouchers.json`

### Step 4: Build Tally XML

1. Read mapped vouchers
2. For each voucher, generate Tally XML envelope:
   - `<ENVELOPE>` → `<HEADER>` → `<BODY>` → `<IMPORTDATA>` → `<TALLYMESSAGE>`
   - Create `<VOUCHER>` elements with proper debit/credit allocations
   - Include bill-wise details, GST details, narration
3. Save XML files to `.tmp/tally_xml/`

### Step 5: Push to Tally

1. Read XML files from `.tmp/tally_xml/`
2. POST each to Tally's HTTP server: `http://{TALLY_HOST}:{TALLY_PORT}`
3. Parse response for success/failure
4. Log results to sync_log.csv
5. Report summary: X synced, Y failed, Z skipped

## Tally XML Format Reference

### Purchase Voucher Example
```xml
<ENVELOPE>
  <HEADER>
    <TALLYREQUEST>Import Data</TALLYREQUEST>
  </HEADER>
  <BODY>
    <IMPORTDATA>
      <REQUESTDESC>
        <REPORTNAME>Vouchers</REPORTNAME>
        <STATICVARIABLES>
          <SVCURRENTCOMPANY>Company Name</SVCURRENTCOMPANY>
        </STATICVARIABLES>
      </REQUESTDESC>
      <REQUESTDATA>
        <TALLYMESSAGE xmlns:UDF="TallyUDF">
          <VOUCHER VCHTYPE="Purchase" ACTION="Create">
            <DATE>20260226</DATE>
            <VOUCHERTYPENAME>Purchase</VOUCHERTYPENAME>
            <PARTYLEDGERNAME>Vendor Name</PARTYLEDGERNAME>
            <VOUCHERNUMBER>BILL/001</VOUCHERNUMBER>
            <NARRATION>Odoo Bill: BILL/2026/001</NARRATION>
            <ALLLEDGERENTRIES.LIST>
              <LEDGERNAME>Purchase Account</LEDGERNAME>
              <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>
              <AMOUNT>-10000.00</AMOUNT>
            </ALLLEDGERENTRIES.LIST>
            <ALLLEDGERENTRIES.LIST>
              <LEDGERNAME>CGST Input</LEDGERNAME>
              <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>
              <AMOUNT>-900.00</AMOUNT>
            </ALLLEDGERENTRIES.LIST>
            <ALLLEDGERENTRIES.LIST>
              <LEDGERNAME>SGST Input</LEDGERNAME>
              <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>
              <AMOUNT>-900.00</AMOUNT>
            </ALLLEDGERENTRIES.LIST>
            <ALLLEDGERENTRIES.LIST>
              <LEDGERNAME>Vendor Name</LEDGERNAME>
              <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>
              <AMOUNT>11800.00</AMOUNT>
            </ALLLEDGERENTRIES.LIST>
          </VOUCHER>
        </TALLYMESSAGE>
      </REQUESTDATA>
    </IMPORTDATA>
  </BODY>
</ENVELOPE>
```

### Sales Voucher follows same structure with `VCHTYPE="Sales"` and reversed debit/credit.

## GST Handling

| Odoo Tax | Tally Ledger (Suggested) |
|----------|--------------------------|
| CGST @ 9% | CGST Input / CGST Output |
| SGST @ 9% | SGST Input / SGST Output |
| IGST @ 18% | IGST Input / IGST Output |
| CGST @ 2.5% | CGST Input @ 2.5% / CGST Output @ 2.5% |
| SGST @ 2.5% | SGST Input @ 2.5% / SGST Output @ 2.5% |
| IGST @ 5% | IGST Input @ 5% / IGST Output @ 5% |

> These must match your actual Tally ledger names exactly.

## Edge Cases & Learnings

- **Tally company name must match exactly** (case-sensitive) or import silently fails
- **Tally XML Server must be enabled** in Tally Prime settings before pushing
- **Duplicate prevention**: Always check sync_log.csv before re-syncing; use Odoo invoice number as unique key
- **Multi-currency**: If Odoo has foreign currency invoices, convert to INR before Tally import (Tally handles forex differently)
- **Rounding differences**: Odoo and Tally may round GST differently by a few paise—use Tally's rounding method
- **Date format**: Tally expects `YYYYMMDD` format (no separators)
- **Amount sign convention**: In Tally XML, negative = debit, positive = credit (counterintuitive)
- **Bill-wise details**: For payment reconciliation, include `<BILLALLOCATIONS.LIST>` with invoice references
- **Inventory vouchers**: If items have stock tracking, consider using Tally's inventory vouchers instead of accounting vouchers (Phase 3+)
- **Odoo 19 API**: Uses XML-RPC at `/xmlrpc/2/common` and `/xmlrpc/2/object`

## Sync Schedule Options

| Mode | When to Use |
|------|-------------|
| **Daily** | High volume of transactions, need real-time books |
| **Weekly (Recommended)** | Moderate volume, review before import |
| **On-demand** | Run manually when needed |

## Safety & Rollback

- Always generate XML files first (`--dry-run` mode) before pushing to Tally
- Review XML in `.tmp/tally_xml/` before confirming push
- Tally has an "Alter" mode for vouchers—if wrong data is pushed, it can be deleted from Tally manually
- Keep sync_log.csv as audit trail

## Prerequisites Checklist

- [ ] Odoo 19 running with XML-RPC enabled (default)
- [ ] Tally Prime installed with XML Server enabled
- [ ] Tally and Odoo on same network (or VPN if remote)
- [ ] Python environment with `xmlrpc.client`, `requests`, `lxml`
- [ ] `.env` file configured with all credentials
- [ ] Ledger mapping CSV reviewed and completed
- [ ] Test with 2-3 vouchers before full sync
