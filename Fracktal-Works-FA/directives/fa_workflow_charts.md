# Finance & Accounting Workflow Charts — Fracktal Works Private Limited

## Purpose

Generate comprehensive workflow charts for all Finance & Accounting department processes. Output is an Excel workbook with swim-lane process maps, RACI matrices, and step-by-step SOPs for each workflow.

## When to Use

- Documenting F&A department processes
- Onboarding new team members
- Audit preparation — show process controls
- Identifying bottlenecks or gaps in current processes
- Setting up or revising Payment Approval Policy
- Process improvement initiatives

---

## Covered Workflows

| # | Workflow | Description |
|---|----------|-------------|
| 1 | **Accounts Payable (AP)** | Invoice receipt → Verification → Approval → Payment → Reconciliation |
| 2 | **Accounts Receivable (AR)** | Billing → Follow-up → Collection → Receipt → Reconciliation |
| 3 | **Month-End Close** | Trial Balance → Adjustments → Report Generation → Review → Sign-off |
| 4 | **Expense Reimbursement** | Submission → Manager Approval → Finance Review → Payment |
| 5 | **Payment Approval Policy** | Threshold-based approval matrix → Authorization levels → Compliance |
| 6 | **Financial Reporting** | Data Collection → Report Generation → Review → Distribution |
| 7 | **Payroll Processing** | Attendance → Calculation → Approval → Disbursement → Statutory Filing |
| 8 | **Bank Reconciliation** | Statement Download → Matching → Exception Handling → Sign-off |
| 9 | **GST Compliance** | Invoice Capture → GSTR Filing → Payment → Reconciliation |
| 10 | **Budget Planning** | Department Inputs → Consolidation → Review → Approval → Monitoring |

---

## Inputs

| Input | Required | Description |
|-------|----------|-------------|
| Workflows to generate | No | Comma-separated list (e.g., `AP,AR,Month-End`). Defaults to ALL |
| Company name | No | Defaults to "Fracktal Works Private Limited" |
| Output path | No | Defaults to `.tmp/FA_Workflow_Charts_<timestamp>.xlsx` |
| Approval thresholds | No | JSON with payment approval limits (see below) |
| Department heads | No | JSON mapping of roles to names for RACI matrix |

### Default Approval Thresholds

```json
{
  "petty_cash": {"limit": 5000, "approver": "Department Head"},
  "operational": {"limit": 50000, "approver": "Finance Manager"},
  "capital": {"limit": 200000, "approver": "CFO"},
  "strategic": {"limit": 500000, "approver": "CEO"},
  "board_level": {"limit": null, "approver": "Board of Directors"}
}
```

---

## Execution

### Step 1: Generate All Workflow Charts

```bash
python execution/fa_workflow_generator.py --company "Fracktal Works Private Limited"
```

### Step 2: Generate Specific Workflows Only

```bash
python execution/fa_workflow_generator.py --workflows "AP,AR,Payment Approval" --company "Fracktal Works Private Limited"
```

### Step 3: Customize Approval Policy

```bash
python execution/fa_workflow_generator.py --workflows "Payment Approval" --thresholds '{"petty_cash": {"limit": 5000, "approver": "Team Lead"}, "operational": {"limit": 100000, "approver": "Finance Manager"}}'
```

### Step 4: Export to Google Sheets (optional)

```bash
python execution/update_sheet.py --spreadsheet-id YOUR_SHEET_ID --source ".tmp/FA_Workflow_Charts_*.xlsx"
```

---

## Output Structure

The generated Excel workbook contains the following sheets:

### Per Workflow:
| Sheet | Content |
|-------|---------|
| `{Workflow}_Flowchart` | Step-by-step process with swim lanes (Role → Step → Decision → Next) |
| `{Workflow}_RACI` | RACI matrix (Responsible, Accountable, Consulted, Informed) |

### Summary Sheets:
| Sheet | Content |
|-------|---------|
| `Index` | Table of contents with hyperlinks to each workflow |
| `Approval_Matrix` | Complete payment approval policy with thresholds |
| `Controls_Summary` | Key controls across all processes for audit |
| `KPI_Metrics` | Suggested KPIs for each workflow |

---

## Flowchart Format (Excel Swim Lane)

Each flowchart sheet uses this column structure:

| Column | Description |
|--------|-------------|
| A | **Step #** — Sequential step number |
| B | **Phase** — Process phase (Initiation, Processing, Approval, Completion) |
| C | **Responsible Role** — Who performs this step |
| D | **Action/Task** — What happens at this step |
| E | **Decision Point?** — Yes/No — is this a decision gate? |
| F | **If Yes** — What happens on Yes path |
| G | **If No** — What happens on No path |
| H | **Document/System** — Supporting document or system used |
| I | **Control Point** — Internal control description |
| J | **SLA/Timeline** — Expected time to complete |
| K | **Escalation** — Who to escalate to if SLA breached |

---

## RACI Format

| Column | Description |
|--------|-------------|
| A | **Process Step** |
| B | **Requestor/Initiator** |
| C | **Department Head** |
| D | **Finance Executive** |
| E | **Finance Manager** |
| F | **CFO** |
| G | **CEO** |
| H | **Auditor** |

Values: **R** (Responsible), **A** (Accountable), **C** (Consulted), **I** (Informed)

---

## Edge Cases & Notes

- If approval thresholds are not provided, use the defaults above
- All amounts are in INR (₹)
- GST compliance workflow follows Indian GST regulations
- Payroll workflow assumes monthly payroll cycle
- Bank reconciliation assumes daily bank statement downloads
- Month-end close follows a T+5 working days timeline
- Keep formulas relative so users can extend workflows easily

---

## Learnings Log

| Date | Learning | Applied To |
|------|----------|------------|
| _Initial_ | Use conditional formatting for decision points (green=approved, red=rejected) | All flowcharts |
| _Initial_ | Freeze panes on Row 1 for header visibility | All sheets |
| _Initial_ | Add data validation dropdowns for RACI values (R/A/C/I) | RACI sheets |
