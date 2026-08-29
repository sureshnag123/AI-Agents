# HR Workflow Charts — Fracktal Works Private Limited

## Purpose

Generate comprehensive workflow charts for all Human Resources department processes. Output is an Excel workbook with swim-lane process maps, RACI matrices, and step-by-step SOPs for each workflow.

## When to Use

- Documenting HR department processes
- Onboarding new HR team members
- Audit / compliance preparation — show process controls
- Identifying bottlenecks or gaps in current HR processes
- Setting up or revising leave, exit, or approval policies
- Process improvement initiatives

---

## Covered Workflows

| # | Workflow | Description |
|---|----------|--------------|
| 1 | **Recruitment & Selection** | Requisition → Sourcing → Interviews → Offer → Background Check |
| 2 | **Onboarding & Induction** | Offer Acceptance → Documentation → Induction → Probation Tracking |
| 3 | **Attendance & Leave Management** | Attendance Capture → Leave Application → Approval → LOP Calculation |
| 4 | **Payroll Input Processing** | Attendance/Leave Compilation → Master Update → Handover to Finance |
| 5 | **Statutory Compliance (PF/ESI/PT)** | Registration → Monthly Contribution → Filing → Renewals |
| 6 | **Performance Management** | Goal Setting → Mid-Year Review → Annual Appraisal → Increment/Promotion |
| 7 | **Exit & Full-and-Final Settlement** | Resignation → Notice Period → Exit Interview → Clearance → F&F Settlement |
| 8 | **Employee Grievance & POSH** | Complaint → Acknowledgement → Investigation → ICC Resolution |
| 9 | **Training & Development** | Training Need Identification → Planning → Delivery → Evaluation |
| 10 | **HR Policy & Documentation** | Drafting → Legal Review → Approval → Rollout → Periodic Audit |

---

## Inputs

| Input | Required | Description |
|-------|----------|-------------|
| Workflows to generate | No | Comma-separated list (e.g., `Recruitment,Onboarding,Exit`). Defaults to ALL |
| Company name | No | Defaults to "Fracktal Works Private Limited" |
| Output path | No | Defaults to `.tmp/HR_Workflow_Charts_<timestamp>.xlsx` |
| Leave policy | No | JSON with leave type entitlements (see below) |
| Approval matrix | No | JSON mapping HR decision types to approver roles |

### Default Leave Policy (days/year)

```json
{
  "casual_leave": {"days": 12, "approver": "Reporting Manager"},
  "sick_leave": {"days": 12, "approver": "Reporting Manager"},
  "earned_leave": {"days": 15, "approver": "Reporting Manager + HR"},
  "maternity_leave": {"days": 182, "approver": "HR Head"},
  "paternity_leave": {"days": 7, "approver": "Reporting Manager"},
  "loss_of_pay": {"days": null, "approver": "HR Head"}
}
```

---

## Execution

### Step 1: Generate All Workflow Charts

```bash
python execution/hr_workflow_generator.py --company "Fracktal Works Private Limited"
```

### Step 2: Generate Specific Workflows Only

```bash
python execution/hr_workflow_generator.py --workflows "Recruitment,Onboarding,Exit" --company "Fracktal Works Private Limited"
```

### Step 3: Customize Leave Policy

```bash
python execution/hr_workflow_generator.py --workflows "Attendance" --leave-policy '{"casual_leave": {"days": 10, "approver": "Reporting Manager"}}'
```

### Step 4: Generate the (currently empty) HR/Payroll Master Workbook

```bash
python execution/generate_hr_master_workbook.py --company "Fracktal Works Private Limited" --fy "2026-27"
```

Populate `Master_Employees` with real employee data, then use `execution/generate_ecr_file.py` for monthly PF ECR filing once statutory codes (Establishment Code, ESI Code) are available.

### Step 5: Export to Google Sheets (optional)

```bash
python execution/update_sheet.py --spreadsheet-id YOUR_SHEET_ID --source ".tmp/HR_Workflow_Charts_*.xlsx"
```

---

## Output Structure

The generated Excel workbook contains the following sheets:

### Per Workflow:
| Sheet | Content |
|-------|---------|
| `{Workflow}_Flow` | Step-by-step process with swim lanes (Role → Step → Decision → Next) |
| `{Workflow}_RACI` | RACI matrix (Responsible, Accountable, Consulted, Informed) |

### Summary Sheets:
| Sheet | Content |
|-------|---------|
| `Index` | Table of contents with hyperlinks to each workflow |
| `Leave_Policy` | Complete leave entitlement and approval policy |
| `Controls_Summary` | Key controls across all HR processes for audit |
| `KPI_Metrics` | Suggested HR KPIs for each workflow |

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
| B | **Employee/Candidate** |
| C | **Reporting Manager** |
| D | **HR Executive** |
| E | **HR Manager** |
| F | **HR Head** |
| G | **Finance/CFO** |
| H | **Auditor** |

Values: **R** (Responsible), **A** (Accountable), **C** (Consulted), **I** (Informed)

---

## Edge Cases & Notes

- All statutory rules follow Indian labour law (PF, ESI, PT, Gratuity, Maternity Benefit Act, POSH Act)
- PF applicability: mandatory once monthly wages ≤ ₹15,000 at PF-wage definition, or establishment already covered
- ESI applicability: mandatory for employees with gross wages ≤ ₹21,000/month (₹25,000 for persons with disability)
- Payroll Input Processing hands off to the Finance team's Payroll Processing workflow (see Fracktal-Works-FA agent) — HR owns attendance/leave/master data, Finance owns calculation & disbursement
- Exit workflow assumes standard 30/60/90-day notice period per employment contract
- Keep formulas relative so users can extend workflows easily
- No real employee data is stored in this workspace until the user populates `Master_Employees` — treat any generated workbook as a template until then

---

## Learnings Log

| Date | Learning | Applied To |
|------|----------|------------|
| _Initial_ | Use conditional formatting for decision points (green=approved, red=rejected) | All flowcharts |
| _Initial_ | Freeze panes on Row 1 for header visibility | All sheets |
| _Initial_ | Add data validation dropdowns for RACI values (R/A/C/I) | RACI sheets |
| _Initial_ | Statutory establishment codes (PF/ESI) are company-specific and unknown until registration is confirmed — leave as placeholders in generated workbooks | Master Workbook, ECR generator |
