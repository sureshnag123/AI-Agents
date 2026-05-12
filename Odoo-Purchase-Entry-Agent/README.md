# Odoo Purchase Entry Agent

Automates vendor bill creation in Odoo 19 from purchase PDF invoices stored on Google Drive.

You maintain a simple Excel sheet with the expense ledger name and a Google Drive link per bill.
The agent downloads each PDF, reads it using Claude Vision (AI OCR), and creates a draft vendor
bill in Odoo — with the original PDF attached.

---

## How It Works

```
Excel File
 ├── Column A: Expense Ledger   ← which Odoo account to debit
 ├── Column B: Notes            ← optional (goes into bill narration)
 └── Column C: PDF Drive Link   ← Google Drive public share URL

         ↓ Agent runs

PDF downloaded from Drive
         ↓
Claude Vision extracts:
  vendor name, invoice number, date, line items, GST, total
         ↓
Odoo vendor bill created (Draft)
  + original PDF attached
         ↓
Excel updated:
 ├── Column D: Odoo Bill ID  (e.g. BILL/2026/0042)
 └── Column E: Status        (SUCCESS / DUPLICATE / ERROR)
```

---

## Project Structure

```
Odoo-Purchase-Entry-Agent/
│
├── README.md                        ← You are here
├── AGENTS.md                        ← Agent system prompt (AI instructions)
├── CLAUDE.md                        ← Claude Code entrypoint (same as AGENTS.md)
├── .env                             ← Connection credentials (never commit)
├── requirements.txt                 ← Python dependencies
│
├── directives/
│   └── pdf_purchase_upload.md       ← Full SOP: inputs, steps, error handling
│
├── execution/
│   ├── odoo_connector.py            ← Odoo 19 XML-RPC connector (reusable)
│   ├── pdf_ocr_extractor.py         ← Download PDF + Claude Vision → JSON
│   ├── odoo_bill_creator.py         ← JSON + ledger → Odoo vendor bill
│   └── pdf_upload_agent.py          ← Main runner: reads Excel, orchestrates all
│
└── .tmp/                            ← Auto-generated: debug JSON, logs (git-ignored)
```

---

## Quick Start

### 1. Install dependencies
```bash
pip install -r requirements.txt
```

### 2. Fill in `.env`
```
ODOO_URL=https://yourcompany.odoo.com
ODOO_DB=yourdb
ODOO_USERNAME=you@company.com
ODOO_PASSWORD=your_api_key_or_password
ANTHROPIC_API_KEY=sk-ant-...
```
Get your Claude API key at: [console.anthropic.com](https://console.anthropic.com)

### 3. Test Odoo connection
```bash
python execution/odoo_connector.py --test
```

### 4. Find your ledger names
```bash
python execution/odoo_connector.py --list-accounts
# Full list saved to .tmp/odoo_accounts.json
```

### 5. Check tax rates available
```bash
python execution/odoo_bill_creator.py --list-taxes
```

### 6. Prepare your Excel file

| A | B | C |
|---|---|---|
| **Expense Ledger** | **Notes** | **PDF Drive Link** |
| Office Supplies | March 2026 | https://drive.google.com/file/d/... |
| Travel Expenses | | https://drive.google.com/file/d/... |

Make sure all PDFs are shared as **"Anyone with the link"** → **Viewer**.

### 7. Dry run (recommended first)
```bash
python execution/pdf_upload_agent.py bills.xlsx --dry-run
```
Check `.tmp/row_N_extracted.json` to verify extracted data before going live.

### 8. Upload
```bash
python execution/pdf_upload_agent.py bills.xlsx
```

### 9. Re-run safely
Already-uploaded rows (column D has a BILL/... value) are automatically skipped.

---

## CLI Reference

```bash
# Full batch upload
python execution/pdf_upload_agent.py bills.xlsx

# Dry run (no Odoo changes)
python execution/pdf_upload_agent.py bills.xlsx --dry-run

# Process a single row (e.g. row 5 for testing)
python execution/pdf_upload_agent.py bills.xlsx --row 5

# Start from a specific row
python execution/pdf_upload_agent.py bills.xlsx --start-row 3

# Test a specific ledger name
python execution/odoo_bill_creator.py --test-ledger "Office Supplies"

# List all purchase taxes
python execution/odoo_bill_creator.py --list-taxes

# List all Odoo accounts
python execution/odoo_connector.py --list-accounts

# Test OCR on a single Drive URL
python execution/pdf_ocr_extractor.py "https://drive.google.com/file/d/..."
```

---

## What Gets Created in Odoo

| Field | Value |
|---|---|
| Type | Vendor Bill (`in_invoice`) |
| Status | **Draft** (not posted — you review before confirming) |
| Vendor | Matched by name from PDF, or auto-created |
| Expense Account | From your Excel Column A |
| Invoice Date | From PDF |
| Line Items | From PDF with GST/tax |
| Attachment | Original PDF file |
| Narration | Source Drive URL + your Notes |

---

## Troubleshooting

See [directives/pdf_purchase_upload.md](directives/pdf_purchase_upload.md) for detailed
error explanations and fixes.
