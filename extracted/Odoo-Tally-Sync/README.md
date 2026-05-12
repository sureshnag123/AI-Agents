# Odoo 19 ↔ Tally Prime Sync Agent

Eliminates double-entry of accounting transactions by automatically syncing vouchers from **Odoo 19** to **Tally Prime**.

## The Problem

| Pain Point | Impact |
|-----------|--------|
| Purchase & Sales entered in both Odoo and Tally | 2× data entry time |
| Cash, Bank, Journal vouchers only in Tally | Auditors need everything in Tally |
| Manual process, error-prone | Mismatches between systems |

## The Solution

This agent automatically exports invoices, payments, and returns from Odoo 19 and imports them into Tally Prime as vouchers — on a daily or weekly schedule.

```
Odoo 19 (XML-RPC) ──→ Transform ──→ Tally Prime (XML Server)
   │                                      │
   ├─ Purchase Invoices ──→ Purchase Vouchers
   ├─ Sales Invoices ──────→ Sales Vouchers
   ├─ Vendor Payments ─────→ Payment Vouchers
   ├─ Customer Receipts ───→ Receipt Vouchers
   ├─ Debit Notes ─────────→ Debit Note Vouchers
   └─ Credit Notes ────────→ Credit Note Vouchers
```

## Quick Start

### 1. Setup
```powershell
# Run the setup script (creates venv, installs deps, copies .env)
.\setup.ps1
```

### 2. Configure
Edit `.env` with your connection details:
```
ODOO_URL=https://your-company.odoo.com
ODOO_DB=your-database
ODOO_USERNAME=admin@yourcompany.com
ODOO_PASSWORD=your-api-key

TALLY_HOST=localhost
TALLY_PORT=9000
TALLY_COMPANY_NAME=Your Company Name As In Tally
```

### 3. Enable Tally XML Server
In Tally Prime: **F12 → Connectivity → Tally Prime Server = Yes → Port 9000**

### 4. Test Connections
```bash
python execution/odoo_connector.py --test
python execution/tally_connector.py --test
```

### 5. Generate Ledger Mapping
```bash
python execution/odoo_tally_sync.py --generate-mapping
# Review and edit .tmp/ledger_mapping.json
```

### 6. Dry Run
```bash
python execution/odoo_tally_sync.py --dry-run
```

### 7. Sync!
```bash
# Incremental sync (only new entries)
python execution/odoo_tally_sync.py

# Specific types only
python execution/odoo_tally_sync.py --types "purchase,sales"

# Full sync for a date range
python execution/odoo_tally_sync.py --mode full --from 2026-02-01 --to 2026-02-28
```

### 8. Automate
```bash
# Daily at 8 PM
python execution/sync_scheduler.py --hour 20

# Weekly on Saturday
python execution/sync_scheduler.py --weekly --day saturday --hour 10

# Or use Windows Task Scheduler with one-shot mode
python execution/sync_scheduler.py --once
```

## Scripts

| Script | Purpose |
|--------|---------|
| `odoo_connector.py` | Fetches invoices, payments, accounts from Odoo 19 |
| `tally_connector.py` | Communicates with Tally Prime XML Server |
| `tally_xml_builder.py` | Transforms Odoo data into Tally XML voucher format |
| `odoo_tally_sync.py` | Main sync pipeline (fetch → transform → push) |
| `sync_report.py` | Generates sync status reports |
| `sync_scheduler.py` | Schedules daily/weekly automated sync |

## Architecture

Built on the **DOE Framework** (Directives → Orchestration → Execution):
- **Directive**: `directives/odoo_tally_integration.md` — complete SOP
- **Orchestration**: AI agent or manual CLI invocation
- **Execution**: Deterministic Python scripts

## Requirements

- Python 3.10+
- Odoo 19 with API access
- Tally Prime with XML Server enabled (F12 → Connectivity)
- Both systems on the same network (or VPN)
