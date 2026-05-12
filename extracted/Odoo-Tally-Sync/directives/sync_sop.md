# Odoo → Tally Sync: Standard Operating Procedure (SOP)

> **Company**: Fracktal Works Pvt. Ltd.  
> **Version**: 1.0  
> **Last Updated**: 27-Feb-2026  
> **Owner**: Accounts Team  
> **Agent Location**: `D:\Fracktal Documents\AGENTS\odoo-tally-sync`

---

## Table of Contents

1. [Purpose](#1-purpose)
2. [Scope](#2-scope)
3. [Prerequisites Checklist](#3-prerequisites-checklist)
4. [Daily Sync Procedure](#4-daily-sync-procedure)
5. [Weekly Reconciliation](#5-weekly-reconciliation)
6. [Monthly Close Procedure](#6-monthly-close-procedure)
7. [Safety & Precautions](#7-safety--precautions)
8. [Error Handling Playbook](#8-error-handling-playbook)
9. [Rollback Procedure](#9-rollback-procedure)
10. [Audit Trail & Compliance](#10-audit-trail--compliance)

---

## 1. Purpose

This SOP defines the process for syncing accounting entries from **Odoo 19** (operational ERP) to **Tally Prime** (statutory books of accounts). The sync eliminates manual double-entry of Purchase Invoices, Sales Invoices, Debit/Credit Notes, and Payments.

**Key Principle**: Odoo is the **source of truth** for transactions. Tally is the **statutory reporting destination**. Data flows **one-way** only (Odoo → Tally).

---

## 2. Scope

### What Gets Synced

| Voucher Type | Odoo Source | Tally Destination | Recommended Frequency |
|---|---|---|---|
| Purchase Invoices | Vendor Bills (`in_invoice`) | Purchase Voucher | Daily |
| Sales Invoices | Customer Invoices (`out_invoice`) | Sales Voucher | Daily |
| Debit Notes | Vendor Refunds (`in_refund`) | Debit Note | Daily |
| Credit Notes | Customer Refunds (`out_refund`) | Credit Note | Daily |
| Vendor Payments | Outbound Payments | Payment Voucher | Weekly |
| Customer Receipts | Inbound Payments | Receipt Voucher | Weekly |

### What Stays Manually in Tally (NOT Synced)

| Entry Type | Reason |
|---|---|
| Cash Vouchers | Entered only in Tally |
| Bank Vouchers (direct) | Entered only in Tally |
| Journal Entries | Auditor-managed |
| Contra Entries | Bank-to-cash, only in Tally |
| Statutory Adjustments | Year-end, auditor-managed |

---

## 3. Prerequisites Checklist

Run this checklist **before your first production sync** and whenever the setup changes.

### One-Time Setup

- [ ] Tally Prime installed and company data loaded
- [ ] Tally XML Server enabled: `F12 → Connectivity → Tally.NET Server → Yes → Port 9000`
- [ ] Python 3.12+ installed on the sync machine
- [ ] Agent folder set up at `D:\Fracktal Documents\AGENTS\odoo-tally-sync`
- [ ] Virtual environment created: `.venv` with all dependencies
- [ ] `.env` file configured with correct credentials:
  ```
  ODOO_URL=https://fracktal.odoo.com
  ODOO_DB=fracktal
  ODOO_USERNAME=suresh@fracktal.in
  ODOO_PASSWORD=<api_key>
  TALLY_HOST=192.168.0.93
  TALLY_PORT=9000
  TALLY_COMPANY_NAME=Fracktal Works Pvt. Ltd. - (from 1-Apr-2025)
  ```
- [ ] Ledger mapping generated and reviewed (`.tmp/ledger_mapping.json`)
- [ ] `auto_create_missing_ledgers` set to `true` in mapping
- [ ] TEST prefix dry-run completed and verified in Tally

### Before Every Sync (Pre-flight)

- [ ] Tally Prime is **open** on the target machine with the correct company loaded
- [ ] Port 9000 is listening (test: open `http://192.168.0.93:9000` in browser — should show XML error, not timeout)
- [ ] Internet is connected (Odoo is cloud-hosted)
- [ ] No other sync process is running

---

## 4. Daily Sync Procedure

**When**: Every working day, **after business hours** (recommended: 6 PM - 8 PM)  
**Who**: Accounts Executive or Automated Scheduler  
**Duration**: ~2-5 minutes for typical daily volume

### Step-by-Step

Open PowerShell and navigate to the agent folder:
```powershell
cd "D:\Fracktal Documents\AGENTS\odoo-tally-sync"
```

#### Step 1: Activate Virtual Environment
```powershell
& ".\.venv\Scripts\Activate.ps1"
```

#### Step 2: Test Connections
```powershell
python execution/tally_connector.py --test
python execution/odoo_connector.py --test
```
**Expected**: Both show "Connected" with company/UID details.  
**If either fails**: See [Error Handling Playbook](#8-error-handling-playbook).

#### Step 3: Dry Run
```powershell
python execution/odoo_tally_sync.py --dry-run --types "purchase,sales,purchase_return,sales_return"
```
**Check**:
- Number of records matches what was entered in Odoo today
- No unexpected ledger names in the output
- No errors in the dry-run summary

#### Step 4: Run Live Sync
```powershell
python execution/odoo_tally_sync.py --mode incremental --types "purchase,sales,purchase_return,sales_return"
```

#### Step 5: Verify Results
```
✅ Expected output: "Total synced: X, Total errors: 0"
```

If errors > 0:
```powershell
python execution/sync_report.py --failed-only
```
Then follow the [Error Handling Playbook](#8-error-handling-playbook).

#### Step 6: Spot-Check in Tally
Open Tally Prime → **Display** → **Day Book** → Select today's date.  
Verify that the synced vouchers appear with correct:
- Party name
- Amount
- Voucher type
- Narration (should contain Odoo bill number)

---

## 5. Weekly Reconciliation

**When**: Every **Saturday** (or last working day of the week)  
**Who**: Senior Accountant  
**Duration**: 15-30 minutes

### Step-by-Step

#### Step 1: Generate Weekly Report
```powershell
python execution/sync_report.py --detailed --from <monday-date> --to <saturday-date>
```

#### Step 2: Cross-Check Register Totals

| Register | Where to Check | What to Compare |
|---|---|---|
| Purchase Register | Odoo: Accounting → Vendors → Bills → Filter by date | Total Amount |
| | Tally: Display → Account Books → Purchase Register | Should match |
| Sales Register | Odoo: Accounting → Customers → Invoices → Filter by date | Total Amount |
| | Tally: Display → Account Books → Sales Register | Should match |

#### Step 3: Identify Discrepancies

If totals don't match:
```powershell
# Check for failed syncs this week
python execution/sync_report.py --failed-only --from <monday-date> --to <saturday-date>

# Re-sync any missed entries
python execution/odoo_tally_sync.py --mode full --from <date> --to <date> --types "purchase"
```

#### Step 4: Sign-Off
Record the weekly reconciliation result:
- Week ending date
- Total vouchers synced (Purchase + Sales + Notes + Payments)
- Discrepancies found and resolved
- Pending items (if any)

---

## 6. Monthly Close Procedure

**When**: **1st-3rd of the following month** (e.g., April data closed by May 3rd)  
**Who**: Senior Accountant + Auditor review  
**Duration**: 30-60 minutes

### Step-by-Step

#### Step 1: Full Month Sync (Catch Any Missed Entries)
```powershell
python execution/odoo_tally_sync.py --mode full --from 2026-04-01 --to 2026-04-30 --types "purchase,sales,purchase_return,sales_return,payment,receipt"
```

#### Step 2: Generate Monthly CSV Report
```powershell
python execution/sync_report.py --csv --from 2026-04-01 --to 2026-04-30
```
This creates a CSV file in `.tmp/` for accountant review.

#### Step 3: Tally Register Comparison

| Check | Expected |
|---|---|
| Purchase Register total | Odoo total ± ₹1 (rounding) |
| Sales Register total | Odoo total ± ₹1 (rounding) |
| Debit Notes total | Odoo total ± ₹1 (rounding) |
| Credit Notes total | Odoo total ± ₹1 (rounding) |
| GST: Input CGST + SGST + IGST | Match Odoo tax summary |
| TDS Deducted total | Match Odoo TDS ledger |

#### Step 4: Archive Batch Logs
```powershell
# Move month's batch files to archive
New-Item -ItemType Directory -Path ".tmp/archive/2026-04" -Force
Move-Item ".tmp/BATCH_2026_04_*.json" ".tmp/archive/2026-04/"
```

#### Step 5: Monthly Sign-Off
- [ ] All voucher types fully synced
- [ ] Register totals matched
- [ ] Error count: 0 (or all resolved)
- [ ] GST figures reconciled
- [ ] TDS figures reconciled
- [ ] Approved by: _________________ Date: _________

---

## 7. Safety & Precautions

### 🔴 CRITICAL RULES (Never Break These)

| # | Rule | Why |
|---|------|-----|
| 1 | **ALWAYS dry-run before live sync** | Prevents pushing wrong data to Tally. A dry-run costs nothing. |
| 2 | **Never run two syncs simultaneously** | Will cause duplicate vouchers in Tally. The sync has no locking mechanism. |
| 3 | **Never delete `sync_log.json`** | This tracks what's been synced. Deleting it will cause ALL entries to re-sync as duplicates. |
| 4 | **Never delete `ledger_mapping.json`** | Contains all your custom mappings. Losing it means re-mapping everything. |
| 5 | **Backup Tally data before first full sync** | Tally → F12 → Backup. If something goes wrong, you can restore. |
| 6 | **Sync AFTER business hours** | Tally users may be editing data. Syncing during work hours risks conflicts. |

### 🟡 IMPORTANT PRECAUTIONS

| Area | Precaution | Action |
|------|------------|--------|
| **Odoo API Key** | Keys can expire or be revoked | Check `.env` if authentication fails. Generate new key in Odoo → Settings → Users → API Keys |
| **Tally Restart** | XML Server resets when Tally restarts | After every Tally restart: `F12 → Connectivity → Tally.NET Server → Yes` |
| **Company Name** | Must be EXACT match including special characters | The company in Tally is `Fracktal Works Pvt. Ltd. - (from 1-Apr-2025)` — not the formal name |
| **New Vendors/Customers** | First sync may fail if partner doesn't exist in Tally | Keep `auto_create_missing_ledgers: true` in mapping, OR create the party in Tally first |
| **New Account Heads** | Odoo may add new GL accounts over time | Periodically run `--generate-mapping` to discover new accounts and map them |
| **Tally Financial Year** | Tally has FY boundaries | If syncing across FY (e.g., March vs April), ensure both years are loaded in Tally |
| **Network** | Tally runs on local network only | Ensure the sync machine and Tally machine are on the same LAN, or use VPN |

### 🔐 SECURITY CHECKLIST

| # | Security Measure | Status |
|---|-----------------|--------|
| 1 | Odoo API key stored in `.env` only (never in code) | [ ] |
| 2 | `.env` is in `.gitignore` (never version-controlled) | [ ] |
| 3 | Tally port 9000 restricted to local network via Windows Firewall | [ ] |
| 4 | Odoo API key rotated every 90 days | [ ] |
| 5 | Agent folder access limited to authorized personnel | [ ] |
| 6 | Sync machine has screen lock when unattended | [ ] |
| 7 | Tally running on `localhost` or trusted IP only | [ ] |

#### How to Restrict Tally Port (Windows Firewall)
```powershell
# Block port 9000 from external networks (run as Administrator)
New-NetFirewallRule -DisplayName "Tally XML - Block External" `
  -Direction Inbound -LocalPort 9000 -Protocol TCP `
  -RemoteAddress "!192.168.0.0/24" -Action Block
```

#### How to Rotate Odoo API Key
1. Log into Odoo → Settings → Users → Select user
2. Go to **API Keys** tab → Create new key
3. Copy the key
4. Update `.env`: `ODOO_PASSWORD=<new_key>`
5. Delete the old key in Odoo
6. Test: `python execution/odoo_connector.py --test`

---

## 8. Error Handling Playbook

### Common Errors & Fixes

| Error | Cause | Fix |
|-------|-------|-----|
| `ConnectionRefused: port 9000` | Tally XML Server not enabled | Open Tally → F12 → Connectivity → Tally.NET Server → Yes |
| `TimeoutError: Tally unreachable` | Tally not open, or wrong IP | Open Tally, load company, verify IP in `.env` |
| `Odoo authentication failed` | API key expired/revoked | Generate new API key in Odoo, update `.env` |
| `Ledger 'XYZ' does not exist` | Partner/account not in Tally | Add mapping in `ledger_mapping.json` OR set `auto_create_missing_ledgers: true` |
| `EXCEPTIONS=1, CREATED=0` (silent) | Voucher is unbalanced | Check Odoo invoice for unusual lines (TDS, rounding, write-offs). All journal lines must be included. |
| `EXCEPTIONS=1` with no error text | Duplicate voucher number | Check if voucher already exists in Tally Day Book |
| `ValueError: Invalid field` | Odoo API changed | Update field names in `odoo_connector.py` |
| `UnicodeEncodeError` | Special characters in names | Already handled by XML escaping. If new case, add to `_esc()` in builder |

### Escalation Path

```
Level 1: Accounts Executive
  → Re-run with --dry-run, check error log
  → Fix ledger mapping if needed
  → Re-sync failed entries

Level 2: Senior Accountant  
  → Review unbalanced vouchers
  → Verify GST / TDS calculations in Odoo
  → Manual correction in Tally if needed

Level 3: IT / Developer
  → Script bugs, Odoo API changes
  → Network / firewall issues
  → Data corruption recovery
```

---

## 9. Rollback Procedure

If a sync batch pushed incorrect data to Tally:

### Option A: Delete Specific Vouchers in Tally
1. Open Tally → Display → Day Book → Go to the date
2. Find the incorrect voucher (look for Narration starting with "Odoo")
3. Press `Alt+D` to delete
4. Remove the ID from `.tmp/sync_log.json` under `synced_ids`
5. Fix the issue (mapping, data, etc.)
6. Re-sync with `--mode full` for that date range

### Option B: Rollback Entire Batch
```powershell
# Identify the batch
python execution/sync_report.py --last-batch

# Rollback (deletes all vouchers from that batch in Tally)
python execution/odoo_tally_sync.py --rollback --batch-id BATCH_2026_02_27_151610
```

### Option C: Restore Tally Backup
1. Close Tally Prime
2. Go to Tally data folder (`C:\Users\<User>\TallyPrime\Data`)
3. Replace with the backup taken before sync
4. Open Tally and verify
5. Clear sync state: delete `.tmp/sync_log.json` (will treat everything as unsynced)

> **WARNING**: Option C should be the absolute last resort. Always prefer Option A or B.

---

## 10. Audit Trail & Compliance

### What's Logged

| File | Contents | Retention |
|------|----------|-----------|
| `.tmp/sync_log.json` | All synced IDs by voucher type | Permanent (do not delete) |
| `.tmp/BATCH_*.json` | Per-run details: what was synced, skipped, failed | 1 year minimum |
| `.tmp/logs/sync_*.log` | Detailed timestamped log of each run | 6 months |
| `.tmp/ledger_mapping.json` | Current mapping state | Permanent (do not delete) |

### For Auditors

When auditors need proof that Odoo and Tally are in sync:

1. **Monthly CSV reports** from `sync_report.py --csv` show every voucher synced
2. **Batch JSON files** contain timestamps, Odoo IDs, and Tally responses
3. **Tally Narration field** on every synced voucher contains the Odoo bill/invoice number for cross-reference
4. **REFERENCE field** in Tally maps to the vendor/customer invoice number from Odoo

### Compliance Notes

- All synced vouchers are created via Tally's official XML Import API
- No direct database modification — all entries go through Tally's validation
- Tally rejects unbalanced vouchers (debit ≠ credit) — impossible to post incorrect totals
- GST components (CGST/SGST/IGST) maintain separate ledger entries for GSTR filing

---

## Quick Reference Card

```
┌─────────────────────────────────────────────────────────────┐
│                    DAILY SYNC CHEAT SHEET                    │
├─────────────────────────────────────────────────────────────┤
│                                                             │
│  cd "D:\Fracktal Documents\AGENTS\odoo-tally-sync"          │
│  & ".\.venv\Scripts\Activate.ps1"                           │
│                                                             │
│  1. TEST:  python execution/tally_connector.py --test       │
│  2. DRY:   python execution/odoo_tally_sync.py --dry-run    │
│  3. SYNC:  python execution/odoo_tally_sync.py              │
│  4. CHECK: python execution/sync_report.py                  │
│                                                             │
│  If errors → python execution/sync_report.py --failed-only  │
│  Emergency → python execution/odoo_tally_sync.py --rollback │
│                                                             │
└─────────────────────────────────────────────────────────────┘
```

---

## Revision History

| Version | Date | Author | Changes |
|---------|------|--------|---------|
| 1.0 | 27-Feb-2026 | AI Agent (Copilot) | Initial SOP created after successful test sync of 12 Feb-2026 purchase invoices |
