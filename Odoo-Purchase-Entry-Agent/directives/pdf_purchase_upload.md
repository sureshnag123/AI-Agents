# Directive: PDF Purchase Bill Upload Agent

## Purpose
Read an Excel file containing expense ledger names and Google Drive PDF links,
extract invoice data from each PDF using Claude Vision (OCR),
and create vendor bills (draft) in Odoo 19.

---

## Input: Excel File Format

| Column A | Column B | Column C | Column D | Column E |
|---|---|---|---|---|
| Expense Ledger *(required)* | Notes *(optional)* | PDF Drive Link *(required)* | Odoo Bill ID *(written back)* | Status *(written back)* |
| Office Supplies | April invoice | https://drive.google.com/file/d/... | BILL/2026/0042 | SUCCESS |

- **Column A** — Expense Ledger: must match an account name in Odoo's Chart of Accounts exactly (case-insensitive)
- **Column B** — Notes: optional text added to bill narration
- **Column C** — PDF Drive Link: Google Drive "Anyone with link can view" share URL
- **Column D, E** — written back by the agent after processing; do not fill manually

---

## Scripts

| Script | Purpose |
|---|---|
| `execution/pdf_ocr_extractor.py` | Download PDF from Drive, send to Claude Vision, return structured JSON |
| `execution/odoo_bill_creator.py` | Take JSON + ledger name, create Odoo vendor bill, attach PDF |
| `execution/pdf_upload_agent.py` | Orchestrator — reads Excel, calls both above, writes results back |

---

## How to Run

### 1. First-time setup
```bash
# Install dependencies
pip install -r requirements.txt

# Verify Odoo connection
python execution/odoo_connector.py --test

# Check that your expense ledger names exist in Odoo
python execution/odoo_connector.py --list-accounts
# Full list saved to .tmp/odoo_accounts.json — use exact names in Excel column A

# Check available purchase tax rates
python execution/odoo_bill_creator.py --list-taxes
```

### 2. Prepare Excel file
- Create `bills.xlsx` with headers in row 1
- Fill Column A (Expense Ledger) and Column C (PDF Drive Link) for each bill
- Make sure all PDFs are shared as "Anyone with the link can view"

### 3. Dry run first (recommended)
```bash
python execution/pdf_upload_agent.py bills.xlsx --dry-run
```
Dry run: downloads + OCRs each PDF, prints extracted data, does NOT create anything in Odoo.
Check `.tmp/row_N_extracted.json` to verify what was extracted.

### 4. Upload
```bash
python execution/pdf_upload_agent.py bills.xlsx
```

### 5. Process a single row (for testing or re-runs)
```bash
python execution/pdf_upload_agent.py bills.xlsx --row 3
```

---

## What the Agent Does Per Row

```
Row N
 ├─ Read Expense Ledger (Col A) + Drive URL (Col C)
 ├─ Skip if Col D already has a BILL/... reference
 ├─ Download PDF from Google Drive
 ├─ Convert PDF pages → images (PyMuPDF, no poppler needed)
 ├─ Send images to Claude Vision (claude-sonnet-4-6)
 │    Extracts: vendor_name, invoice_number, invoice_date, due_date,
 │              gstin, line_items[], subtotal, tax_rate, tax_amount, total
 ├─ Save extracted JSON → .tmp/row_N_extracted.json (debug)
 ├─ Find or create vendor in Odoo (res.partner)
 ├─ Look up expense account by ledger name (account.account)
 ├─ Find matching purchase tax by rate (account.tax)
 ├─ Check for duplicate (same vendor + invoice number)
 ├─ Create vendor bill as DRAFT (account.move, move_type=in_invoice)
 ├─ Attach original PDF to the bill (ir.attachment)
 └─ Write back: Col D = BILL/YYYY/NNNN, Col E = SUCCESS / DUPLICATE / ERROR
```

---

## Error Handling

| Error | What the agent does |
|---|---|
| Expense Ledger not found in Odoo | Writes error to Col E, skips row, continues |
| PDF download fails (private/expired link) | Writes error to Col E, continues |
| Claude Vision returns malformed JSON | Writes error, saves raw response to .tmp |
| Duplicate invoice (same vendor + invoice#) | Writes "DUPLICATE — BILL/..." to Col E |
| Any other exception | Full error message written to Col E |

The agent saves Excel after every row — safe to interrupt and re-run.

---

## Odoo Bill Fields Set

| Odoo Field | Value |
|---|---|
| `move_type` | `in_invoice` (Vendor Bill) |
| `state` | `draft` (not posted) |
| `partner_id` | Matched or created from PDF vendor name |
| `journal_id` | First purchase journal |
| `ref` | Invoice number from PDF |
| `invoice_date` | Date from PDF |
| `invoice_date_due` | Due date from PDF (if present) |
| `invoice_line_ids` | Line items from PDF with tax |
| `narration` | "Uploaded via PDF Purchase Agent. Source PDF: <url>" |
| Attachment | Original PDF file |

---

## Environment Variables Required (.env)

```
ODOO_URL=https://yourcompany.odoo.com
ODOO_DB=yourdb
ODOO_USERNAME=user@company.com
ODOO_PASSWORD=your_api_key

ANTHROPIC_API_KEY=sk-ant-...
```

---

## Common Issues

**"Expense ledger not found"**
→ Run `python execution/odoo_connector.py --list-accounts`
→ Use the exact name from `code — Name` column in Odoo

**"PDF download failed / HTML returned"**
→ Open the Drive link in a browser in private mode — if it asks for login, the file is not publicly shared
→ In Drive: right-click → Share → "Anyone with the link" → Viewer

**"No active purchase tax found for 18%"**
→ Run `python execution/odoo_bill_creator.py --list-taxes`
→ Verify the tax exists and is active in Odoo Settings → Accounting → Taxes

**Claude Vision returns wrong data**
→ Check `.tmp/row_N_extracted.json`
→ If PDF is a scanned image at low resolution, increase DPI in `pdf_ocr_extractor.py` (default 200)
