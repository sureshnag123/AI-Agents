# Fracktal Works HR

> Generated from the DOE Framework, following the same pattern as `Fracktal-Works-FA`

**Type:** Human Resources Workflow & Payroll Agent

Generates comprehensive workflow charts for all HR department processes at Fracktal Works Private Limited — Recruitment, Onboarding, Attendance & Leave, Payroll Input, Statutory Compliance (PF/ESI/PT), Performance Management, Exit & F&F Settlement, Grievance & POSH, Training & Development, and HR Policy. Also maintains an empty, formula-driven HR/Payroll master workbook (employee master + monthly salary statements) and generates EPFO ECR 2.0 filing files once populated with real data.

**No employee data or statutory registration numbers are included.** Every generated workbook is a template — fill in `Statutory_Compliance` and `Master_Employees` before relying on any calculation or filing output.

## 🚀 Instant Start

**Just double-click:** `fracktal-works-hr.code-workspace`

VS Code will automatically:
1. Open the workspace
2. Prompt to trust the folder (click **Yes**)
3. Run setup (creates venv, installs dependencies)
4. Prompt to install recommended extensions

Then select **"Fracktal Works HR"** from the Copilot Chat agent dropdown and start working!

## Alternative: Manual Setup

**Windows (PowerShell):**
```powershell
.\setup.ps1
```

**macOS/Linux:**
```bash
chmod +x setup.sh && ./setup.sh
```

This automatically:
- ✅ Creates Python virtual environment
- ✅ Installs all dependencies
- ✅ Copies `.env.example` to `.env`
- ✅ Creates `.tmp/` directory

## Manual Setup (Alternative)

1. **Create and activate virtual environment:**
   ```powershell
   python -m venv .venv
   .\.venv\Scripts\Activate.ps1
   ```

2. **Install dependencies:**
   ```powershell
   pip install -r requirements.txt
   ```

3. **Set up environment:**
   ```powershell
   cp .env.example .env
   # Edit .env only if you plan to use the optional Google Sheets export
   ```

## Using the Agent

1. **Open in VS Code** - The custom agent is pre-configured
2. **Open Copilot Chat** - Press `Ctrl+Shift+I`
3. **Select your agent** - Choose "Fracktal Works HR" from the agent dropdown
4. **Start working** - The agent knows the directives and scripts available

## VS Code Tasks (Optional)

Press `Ctrl+Shift+B` to run the default build task (Setup Agent Environment).

Other tasks available:
- **Setup Agent Environment** - Full one-command setup
- **Activate Virtual Environment** - Activate .venv
- **Install Requirements** - Install/update dependencies

## Structure

```
Fracktal Works HR/
├── fracktal-works-hr.code-workspace  # ← Double-click to open VS Code!
├── setup.ps1 / setup.sh              # One-command setup scripts
├── AGENTS.md                         # System prompt for AI agents
├── .env.example                      # Template for optional API keys
├── requirements.txt                  # Python dependencies
├── .github/agents/                   # VS Code custom agent config
├── .vscode/                          # VS Code settings & tasks
├── directives/                       # What to do (SOPs)
├── execution/                        # How to do it (scripts)
└── webapp/                           # Payroll & Compliance web app (see webapp/README.md)
```

## Payroll & Compliance Web App

`webapp/` is a Flask app for the monthly payroll cycle: upload the payroll Excel, download per-employee PDF
payslips, and track PF/ESI/PT filing status through a browser. It reuses the same extraction/reconciliation
logic as `execution/generate_payslip.py` (imported directly, not re-implemented) so the CLI and the web app
always agree. See `webapp/README.md` for how to run it locally or host it for other HR staff to use.

## Available Directives

| Directive | Description |
|-----------|--------------|
| `directives/hr_workflow_charts.md` | HR workflow chart generation — 10 workflows covering Recruitment, Onboarding, Attendance, Payroll Input, Statutory Compliance, Performance Management, Exit, Grievance & POSH, Training, HR Policy |

## Quick Usage

```bash
# Generate ALL HR workflow charts (10 workflows)
python execution/hr_workflow_generator.py

# Generate specific workflows only
python execution/hr_workflow_generator.py --workflows "Recruitment,Onboarding,Exit"

# Custom company name
python execution/hr_workflow_generator.py --company "Your Company Name"

# List available workflows
python execution/hr_workflow_generator.py --list-workflows

# Generate the empty HR/Payroll master workbook
python execution/generate_hr_master_workbook.py --fy "2026-27"

# Generate the monthly PF ECR file (after Master_Employees is populated)
python execution/generate_ecr_file.py --workbook ".tmp/Fracktal_HR_Master_FY2627_<timestamp>.xlsx" --month "APR26"
```

Outputs land in `.tmp/`.

## Required API Keys

- **None required** for workflow charts, master workbook, or ECR generation (fully local)
- Optional: Google Sheets credentials for exporting workflow charts to Sheets

## Google Credentials (if using Google Sheets)

If your agent uses Google Sheets scripts:
1. Place `credentials.json` in this folder
2. Run any sheet script once to generate `token.json`

## Relationship to Fracktal-Works-FA

This agent (HR) and the sibling `Fracktal-Works-FA` agent (Finance) split the payroll cycle: HR owns attendance/leave/employee master and statutory PF/ESI/PT registration & filing; Finance owns salary calculation review, disbursement, and TDS. The **Payroll Input Processing** workflow here is the documented handoff between the two.

---

*This workspace follows the [DOE Framework](https://github.com/vjvarada/DOE-Framework-Agentic-AI) pattern used by Fracktal-Works-FA.*
