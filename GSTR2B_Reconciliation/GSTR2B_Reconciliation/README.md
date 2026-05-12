# GSTR2B Reconciliation Agent

Automated monthly reconciliation between **Tally Prime** data and **GSTR2B** portal downloads. Eliminates manual ITC (Input Tax Credit) tracking and produces accurate, Excel-based reconciliation reports.

---

## Features

- **Smart Matching** — Primary match on GSTIN + Invoice Number, secondary fuzzy match on vendor name, date, and amounts
- **Invoice Number Normalization** — Strips spaces, hyphens, special characters for reliable matching
- **Narration Parsing** — Extracts invoice references from Tally narrations when bill numbers are missing
- **Duplicate Detection** — Flags potential duplicate ITC claims in both sources
- **Amount Tolerance** — Configurable percentage and absolute tolerance for rounding differences
- **Formatted Excel Reports** — Color-coded, filterable reports with summary dashboard
- **Streamlit Web UI** — Upload files, run reconciliation, and download reports from a browser
- **CLI Mode** — Run via command line for automation/scheduling

---

## Folder Structure

```
GSTR2B_Reconciliation/
├── app.py                      # Streamlit web UI
├── run.py                      # CLI runner
├── config.py                   # Configuration & settings
├── requirements.txt            # Python dependencies
├── generate_sample_data.py     # Creates sample test files
├── core/
│   ├── __init__.py
│   ├── data_loader.py          # Tally & GSTR2B file parsers
│   ├── reconciler.py           # Matching engine
│   ├── report_generator.py     # Excel report builder
│   └── utils.py                # Cleaning, validation, fuzzy matching
├── output/                     # Generated reports (by month)
│   └── 2026-03/
│       ├── Reconciliation_Summary.xlsx
│       └── Detailed_Reco_Report.xlsx
├── sample_data/                # Sample input files
│   ├── Sample_Tally_Export.xlsx
│   └── Sample_GSTR2B.xlsx
└── logs/                       # Execution logs
```

---

## Setup Instructions

### 1. Prerequisites

- Python 3.10 or higher
- pip (Python package manager)

### 2. Install Dependencies

```bash
cd GSTR2B_Reconciliation
pip install -r requirements.txt
```

### 3. Generate Sample Data (Optional)

```bash
python generate_sample_data.py
```

This creates `sample_data/Sample_Tally_Export.xlsx` and `sample_data/Sample_GSTR2B.xlsx` for testing.

### 4. Copy to Target Directory (Optional)

To place the project at the specified path:
```bash
xcopy /E /I GSTR2B_Reconciliation D:\Suresh_AGENTS\GSTR2B_Reconciliation
```

---

## Usage

### Option A: Web UI (Recommended)

```bash
streamlit run app.py
```

This opens a browser interface where you can:
1. Upload Tally export file
2. Upload GSTR2B Excel file
3. Select the month
4. Adjust matching thresholds (optional)
5. Click **Run Reconciliation**
6. View results in tabs and download Excel reports

### Option B: Command Line

```bash
python run.py --tally "path/to/tally.xlsx" --gstr2b "path/to/gstr2b.xlsx" --month 2026-03
```

**Arguments:**

| Flag | Required | Description |
|------|----------|-------------|
| `--tally` | Yes | Path to Tally export (xlsx/csv) |
| `--gstr2b` | Yes | Path to GSTR2B file (xlsx/csv) |
| `--month` | No | Period in YYYY-MM format (default: current month) |
| `--output` | No | Custom output directory |

---

## Input File Formats

### Tally Export

| Column | Description | Required |
|--------|-------------|----------|
| Voucher Type | Purchase / Journal / Debit Note | Yes |
| Voucher Date | Invoice date | Yes |
| Voucher Number | Tally voucher number | No |
| Party Name | Vendor / supplier name | Yes |
| GSTIN | 15-char GST number | Recommended |
| Vendor Bill Number | Supplier invoice reference | Recommended |
| Taxable Value | Base amount before tax | Yes |
| IGST Amount | Integrated GST | Yes |
| CGST Amount | Central GST | Yes |
| SGST Amount | State GST | Yes |
| Total Invoice Value | Gross total | No (computed) |
| Narration | Free text (used for missing bill references) | No |

> Column names are flexible — the system maps common variations automatically.

### GSTR2B File

Standard Excel download from the GST Portal. The system auto-detects the B2B sheet and header row.

| Column | Description |
|--------|-------------|
| Supplier GSTIN | 15-char GST number |
| Trade/Legal name | Supplier name |
| Invoice Number | Supplier invoice reference |
| Invoice Date | Date of invoice |
| Taxable Value | Base amount |
| IGST / CGST / SGST | Tax components |
| Total Tax | Sum of all taxes |

---

## Output Reports

### Reconciliation_Summary.xlsx

Single-sheet dashboard with key metrics: record counts, ITC totals, differences, and processing time.

### Detailed_Reco_Report.xlsx

Multi-sheet workbook:

| Sheet | Description |
|-------|-------------|
| Reconciliation Summary | Dashboard with all metrics |
| Matched Entries | Successfully reconciled records with match type and score |
| In Tally NOT in GSTR2B | ITC booked in Tally but not appearing in GSTR2B |
| In GSTR2B NOT in Tally | ITC available in GSTR2B but not recorded in Tally |
| GSTR2B No Bill-Payment | GSTR2B entries with no corresponding bill/payment in Tally |
| Amount Mismatch | Matched entries where tax amounts differ |
| Duplicate Entries | Potential duplicate ITC claims |

---

## Matching Logic

### Phase 1 — Primary Match (Exact)
- **Key:** GSTIN + Normalized Invoice Number
- Amounts are verified; mismatches flagged separately

### Phase 2 — Secondary Match (Fuzzy)
- **Scoring:** GSTIN (30pts) + Name similarity (20pts) + Invoice number (20pts) + Date proximity (15pts) + Amount match (15pts)
- Minimum score of 50 required
- Vendor names are cleaned (strip Pvt/Ltd/LLP etc.) and compared using token-sort and partial ratios

### Normalization
- Invoice numbers: uppercase, strip spaces/hyphens/slashes/dots
- GSTIN: uppercase, trimmed
- Vendor names: remove legal suffixes, special characters

---

## Configuration

Edit `config.py` to adjust:

| Setting | Default | Description |
|---------|---------|-------------|
| `FUZZY_MATCH_THRESHOLD` | 80 | Minimum fuzzy name match score (0-100) |
| `DATE_TOLERANCE_DAYS` | 3 | ± days for date matching |
| `AMOUNT_TOLERANCE_PERCENT` | 1.0 | % tolerance for amounts |
| `AMOUNT_TOLERANCE_ABSOLUTE` | 1.0 | ₹ absolute tolerance |
| `TALLY_VOUCHER_TYPES` | Purchase, Journal, Debit Note | Which voucher types to include |

---

## Troubleshooting

| Issue | Solution |
|-------|----------|
| Column not found | Check that your file has at least one of the column name variants listed in `config.py` |
| Low match rate | Lower the `FUZZY_MATCH_THRESHOLD` or increase `DATE_TOLERANCE_DAYS` |
| GSTR2B header not detected | Ensure the Excel has column headers in one of the first 10 rows |
| Slow performance | For 10K+ invoices, the CLI mode is faster than the UI |

---

## License

Internal tool — Fracktal Works Private Limited.
