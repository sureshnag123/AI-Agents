# Odoo Manufacturing Implementation Agent Directive

> Standard Operating Procedure for implementing and monitoring the end-to-end manufacturing workflow at Fracktal Works Pvt Ltd using Odoo as the primary system.

## Goal
Automate the implementation tracking, configuration validation, KPI reporting, and operational monitoring of the Fracktal Works manufacturing workflow across Sales → Purchase → Inventory → Production → QC → Dispatch. Deliver visibility via Google Sheets dashboards and actionable status reports.

## Inputs
| Input | Required | Description |
|-------|----------|-------------|
| Odoo URL | Yes | Base URL of Odoo instance (e.g., https://fracktal.odoo.com) |
| Odoo DB | Yes | Odoo database name |
| Odoo Username | Yes | Login email for Odoo |
| Odoo API Key | Yes | API key from Odoo Settings → Technical → API Keys |
| Google Sheets ID | No | Dashboard spreadsheet ID for KPI output |
| Report Mode | No | `setup`, `kpi`, `inventory`, `production`, `dispatch`, `full` |
| Date Range | No | Start and end date for reports (default: last 30 days) |

## Tools/Scripts

### Odoo Connection
- `odoo_connect.py` - Base Odoo XML-RPC client. Provides authenticated connection and CRUD helpers used by all other scripts.
  - Requires: `ODOO_URL`, `ODOO_DB`, `ODOO_USERNAME`, `ODOO_API_KEY`

### Setup Validation
- `odoo_setup_checker.py` - Validates that all required modules, inventory locations, approval rules, quality points, and BOMs are configured correctly.
  - Modes: `modules`, `locations`, `approvals`, `quality`, `bom`, `full`
  - Outputs: Checklist with pass/fail status per item

### KPI Reporting
- `odoo_kpi_report.py` - Pulls live KPI metrics from Odoo across all departments.
  - Modes: `procurement`, `inventory`, `production`, `qc`, `dispatch`, `full`
  - Outputs: JSON summary + Google Sheets push

### SKU Management
- `odoo_sku_manager.py` - Generate, assign, and export SKUs for all inventory items
  - Modes: `list-missing`, `list-all`, `preview`, `assign`, `assign-bulk`, `export`
  - SKU Format: `FW-[CATEGORY]-[0001]` (permanent, never changes once assigned)
  - Requires: `ODOO_URL`, `ODOO_DB`, `ODOO_USERNAME`, `ODOO_API_KEY`

### Barcode Operations
- `odoo_barcode.py` - Barcode-driven inventory transactions (GRN, Material Issue, Cycle Count, Dispatch)
  - Modes: `lookup`, `grn`, `issue`, `count`, `dispatch`
  - Accepts: barcode number, SKU (FW-XX-NNNN), or Odoo internal reference
  - Works with: USB scanner (keyboard wedge), Bluetooth scanner, or manual entry
  - Requires: `ODOO_URL`, `ODOO_DB`, `ODOO_USERNAME`, `ODOO_API_KEY`

### Google Sheets
- `read_sheet.py` - Read data from sheets
- `update_sheet.py` - Create or update sheets
- `append_to_sheet.py` - Append rows to sheets
  - Requires: `credentials.json` and OAuth flow

## Workflow

### Mode: barcode — Barcode-Based Inventory Operations

Use this for all physical transactions at the warehouse/stores floor.

#### SKU Numbering System

Every item in inventory gets a permanent SKU in the format `FW-[CATEGORY]-[0001]`.

| Code | Category | Examples |
|------|----------|---------|
| `RM` | Raw Material | Steel bar, Aluminium sheet, Copper wire |
| `BT` | Bought-Out Part | Bearings, Motors, Fasteners, Switches |
| `SF` | Semi-Finished | Machined bracket, Welded frame |
| `AS` | Assembly | Sub-assembly, PCB module |
| `FG` | Finished Goods | Fracktal Move, complete machine |
| `SP` | Spare Part | Replacement nozzle, belt |
| `CS` | Consumable | Welding rods, cutting fluid, sandpaper |
| `PK` | Packaging | Box, foam insert, strap |

**SKU Rules:**
- SKU = Odoo Internal Reference field (`default_code`) on product
- Once assigned, a SKU is PERMANENT — never change or reuse
- SKU is printed on every product label alongside the barcode
- Barcode scanner can scan the barcode OR the operator can type the SKU manually

**Step 1: Find products without SKU**
```bash
python execution/odoo_sku_manager.py --mode list-missing
```

**Step 2: Preview auto-assignment for a category (dry run)**
```bash
python execution/odoo_sku_manager.py --mode preview --category RM
python execution/odoo_sku_manager.py --mode preview --category BT
```

**Step 3: Assign in bulk per category**
```bash
python execution/odoo_sku_manager.py --mode assign-bulk --category RM
python execution/odoo_sku_manager.py --mode assign-bulk --category BT
python execution/odoo_sku_manager.py --mode assign-bulk --category FG
```

**Step 4: Assign individual product (manual)**
```bash
python execution/odoo_sku_manager.py --mode assign --product-id 42 --sku FW-FG-0001
```

**Step 5: Export master list for label printing**
```bash
python execution/odoo_sku_manager.py --mode export --output .tmp/sku_master.csv
```
Import `sku_master.csv` into your label printer software (Zebra, Dymo, Bartender, etc.)
Label must show: **SKU + Product Name + Barcode (if available)**

---

**Hardware setup:**
- USB barcode scanner → plug into PC/laptop → acts as keyboard (no driver needed)
- Bluetooth scanner → pair with device → same keyboard wedge behaviour
- Mobile camera scanner → use Odoo mobile app (native barcode support)
- Products must have barcode field populated in Odoo: Inventory → Products → [Product] → Barcode

**Barcode on product labels:** Print labels from Odoo → Inventory → Products → Print → Barcode PDF

---

#### Barcode: Product Lookup (no changes — info only)
```bash
python execution/odoo_barcode.py --mode lookup --barcode 8901234567890
```
Shows: product name, internal ref, on-hand qty, location-wise stock breakdown.

---

#### Barcode: GRN — Goods Receipt
```bash
python execution/odoo_barcode.py --mode grn --ref PO/2024/0042
```
Or start interactive session (scanner prompts PO ref too):
```bash
python execution/odoo_barcode.py --mode grn --interactive
```
Flow:
1. Enter PO reference → system loads pending items
2. Scan each arriving item barcode → system matches to PO line
3. Enter received qty (or press Enter for 1)
4. Repeat for all items → type `done`
5. Review summary → confirm → GRN validated in Odoo → IQC triggered

**Rule:** Only scan items that PASS IQC. Hold IQC-failed items — do not confirm GRN.

---

#### Barcode: Material Issue (Indent)
```bash
python execution/odoo_barcode.py --mode issue --ref MO/2024/0017
```
Flow:
1. Enter MO reference → system shows BOM-required components
2. Stores person physically picks items → scans each barcode
3. Enter qty being issued
4. System warns if item is NOT in BOM (extra issue = supervisor approval required)
5. Confirm → material consumption recorded against MO → stock deducted

**Rule:** Stores ONLY issues BOM items. Any extra triggers warning and requires sign-off.

---

#### Barcode: Cycle Count
```bash
python execution/odoo_barcode.py --mode count --location "Approved Stock"
```
Flow:
1. Select location to count
2. Scan item → system shows current Odoo qty
3. Enter physical count qty
4. Repeat for all items → type `done`
5. System highlights variances → confirm → inventory adjusted in Odoo
6. Count log saved to `.tmp/cycle_count_YYYYMMDD.json`

**Recommended schedule:** 20–30 items per week (rotating). Every item covered in 6 months.

---

#### Barcode: Dispatch Scan
```bash
python execution/odoo_barcode.py --mode dispatch --ref WH/OUT/00123
```
Flow:
1. Enter Delivery Order reference → system shows items to ship
2. Scan each item being packed → system matches and counts
3. System flags if item is NOT on this delivery order
4. Review: all items matched? → confirm → Delivery Order validated in Odoo → stock reduced
5. Invoice can now be generated (one click from SO)

**Rule:** Never dispatch without completed FQC. Dispatch scan is the final physical gate.

---

### Mode: setup — Validate Odoo Configuration

Run this when first implementing or verifying the system is correctly set up.

**Step 1: Check required modules are installed**
```bash
python execution/odoo_setup_checker.py --mode modules --output .tmp/setup_modules.json
```

**Step 2: Check inventory locations exist**
```bash
python execution/odoo_setup_checker.py --mode locations --output .tmp/setup_locations.json
```

**Step 3: Check approval rules are configured**
```bash
python execution/odoo_setup_checker.py --mode approvals --output .tmp/setup_approvals.json
```

**Step 4: Check quality control points**
```bash
python execution/odoo_setup_checker.py --mode quality --output .tmp/setup_quality.json
```

**Step 5: Check BOM coverage for products**
```bash
python execution/odoo_setup_checker.py --mode bom --output .tmp/setup_bom.json
```

**Step 6: Full setup report**
```bash
python execution/odoo_setup_checker.py --mode full --output .tmp/setup_full.json
```

Interpret the output and tell the user:
- Which items are correctly configured (PASS)
- Which items are missing or misconfigured (FAIL) with exact steps to fix in Odoo
- Priority order to fix FAIL items (highest-impact first)

---

### Mode: kpi — Pull and Report KPIs

Run this weekly or monthly to track operational performance.

**Step 1: Pull full KPI report**
```bash
python execution/odoo_kpi_report.py --mode full --days 30 --output .tmp/kpi_report.json
```

**Step 2: Push to Google Sheets dashboard**
```bash
python execution/update_sheet.py --sheet-id [GOOGLE_SHEETS_ID] --tab "KPI Dashboard" --data-file .tmp/kpi_report.json
```

Interpret the KPI data:
- Flag any KPI below target threshold (see Targets section below)
- Identify top 3 problem areas
- Suggest specific corrective actions

---

### Mode: inventory — Inventory Health Check

Run this daily or weekly to catch issues before they hit production.

```bash
python execution/odoo_kpi_report.py --mode inventory --days 7 --output .tmp/inventory_status.json
```

Report includes:
- Items below reorder point (stockout risk)
- Items in "Under IQC" location > 24 hours (IQC bottleneck)
- Non-PO GRNs in last 7 days (procurement compliance)
- Scrap logged in last 7 days with reasons
- Inventory accuracy from last cycle count

---

### Mode: production — Production Status

Run this daily to track active manufacturing orders.

```bash
python execution/odoo_kpi_report.py --mode production --days 7 --output .tmp/production_status.json
```

Report includes:
- Open MOs by status (Confirmed / In Progress / Done / Blocked)
- MOs overdue (planned end date < today)
- MOs awaiting material (availability check failed)
- Rework alerts (quality failures sent back to production)
- FQC queue (MOs done, awaiting final inspection)

---

### Mode: dispatch — Dispatch Readiness

Run this daily to prepare dispatch schedule.

```bash
python execution/odoo_kpi_report.py --mode dispatch --days 7 --output .tmp/dispatch_status.json
```

Report includes:
- Orders with FQC approved (ready to dispatch today)
- Orders with committed delivery date within 3 days (urgency view)
- Overdue deliveries (committed date already passed)
- Pending invoices (delivered but not invoiced)
- Overdue customer payments

---

### Mode: full — Complete Operational Report

Run this weekly for the management review meeting.

```bash
python execution/odoo_kpi_report.py --mode full --days 7 --output .tmp/full_report.json
python execution/update_sheet.py --sheet-id [GOOGLE_SHEETS_ID] --tab "Weekly Report" --data-file .tmp/full_report.json
```

---

## KPI Targets Reference

| KPI | Target | Alert Threshold |
|-----|--------|-----------------|
| PO Compliance Rate | > 95% | < 85% = CRITICAL |
| Vendor On-Time Delivery | > 80% | < 60% = WARNING |
| Inventory Accuracy | > 95% | < 90% = WARNING |
| IQC Turnaround | < 4 hours | > 8 hours = WARNING |
| FQC First Pass Rate | > 90% | < 80% = WARNING |
| On-Time Dispatch | > 90% | < 75% = CRITICAL |
| BOM Coverage | 100% | < 90% = WARNING |
| Non-PO GRN Rate | < 5% | > 10% = CRITICAL |

---

## Outputs

| Output | Location | Description |
|--------|----------|-------------|
| Setup Checklist | `.tmp/setup_full.json` + printed | Module/config validation with pass/fail |
| KPI Dashboard | Google Sheets | Live metrics across all departments |
| Inventory Alert | `.tmp/inventory_status.json` | Stockout risks, IQC backlogs, scrap |
| Production Status | `.tmp/production_status.json` | Open MOs, overdue, blocked |
| Dispatch Plan | `.tmp/dispatch_status.json` | Ready, urgent, overdue orders |
| Weekly Report | Google Sheets | Full operational summary |

---

## Edge Cases and Error Handling

### Odoo Authentication Failures
- Verify `ODOO_API_KEY` is active in Odoo → Settings → Technical → API Keys
- XML-RPC uses port 443 for HTTPS — ensure firewall allows
- If using Odoo.sh or SaaS: API key auth is the correct method (not password)

### Missing Odoo Modules
- If `stock_landed_costs`, `quality_control`, or `mrp` modules not found:
  → Report as FAIL in setup checker
  → Provide exact Odoo menu path to install: Settings → Apps → search → Install
- Do not attempt to install modules via API — guide user to do it manually

### Empty Data / New Instance
- If no SOs, MOs, or POs exist yet: report counts as 0, do not error
- Note in report: "System newly configured — no historical data yet"

### Google Sheets Auth
- First run requires browser OAuth flow
- Token auto-refreshes after initial auth
- If token expires, delete `token.json` and re-authenticate

### Large Odoo Datasets
- Limit all queries to last 90 days by default
- Use Odoo domain filters to reduce API payload
- Paginate requests if record count > 500

---

## Implementation Phase Tracking

When asked to track implementation progress, use this phase map:

### Phase 1 (Days 1–30): Foundation
- [ ] Odoo modules activated (Quality, Manufacturing, Purchase Approval)
- [ ] Inventory locations created (Under IQC, Approved Stock, Rejection Hold, Scrap)
- [ ] Purchase approval thresholds set
- [ ] Sales approval thresholds set
- [ ] BOM created for top 10 products
- [ ] IQC checklist created in Odoo Quality
- [ ] FQC checklist created in Odoo Quality
- [ ] Team briefed: No PO = No purchase
- [ ] Physical stock count done and Odoo reconciled

### Phase 2 (Days 31–60): Stabilization
- [ ] Work Orders enabled in Manufacturing
- [ ] IPQC checkpoints added to top 5 products
- [ ] Scrap tracking enforced
- [ ] Reordering rules for top 20 consumables
- [ ] 2-vendor RFQ rule in practice
- [ ] Vendor OTD tracking started
- [ ] Automated reminders configured (vendor PO due, customer payment)
- [ ] CEO/Ops dashboard live in Odoo

### Phase 3 (Days 61–90): Optimization
- [ ] First full monthly KPI review completed
- [ ] Lot/Serial traceability for critical items
- [ ] 100% BOM coverage achieved
- [ ] Customer Portal activated
- [ ] Vendor scorecards formalized
- [ ] SOPs written (1 page each for GRN, Indent, IQC, FQC, Dispatch)
- [ ] Backup person trained for every critical role

---

## Best Practices

1. **Run setup checker before any KPI report** — garbage config = garbage metrics
2. **Always use `--output` flag** — save intermediates to `.tmp/` for debugging
3. **Review alerts before pushing to Sheets** — validate numbers make sense
4. **Never modify Odoo data via script** — this agent is READ-ONLY for Odoo
4. **Update this directive** when you discover Odoo API quirks or field name changes
5. **Prefer Odoo domain filters** over fetching all records and filtering in Python

---

## Required Environment Variables
```env
ODOO_URL=https://yourcompany.odoo.com
ODOO_DB=your_database_name
ODOO_USERNAME=admin@yourcompany.com
ODOO_API_KEY=your_odoo_api_key
GOOGLE_SHEETS_CREDENTIALS_FILE=credentials.json
GOOGLE_SHEETS_DASHBOARD_ID=your_google_sheet_id
```

## Dependencies
```
google-api-python-client
google-auth-httplib2
google-auth-oauthlib
requests
python-dotenv
anthropic
```
