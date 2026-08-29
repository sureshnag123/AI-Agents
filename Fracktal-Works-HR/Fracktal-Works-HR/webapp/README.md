# Fracktal Works HR — Payroll & Compliance Web App

Upload the monthly payroll Excel, generate per-employee PDF payslips, and track
PF/ESI/PT filing status through a browser.

## First run (creates the first login)

No default credentials ship with this app. Set these two environment
variables the first time you start it — they seed exactly one admin account:

```powershell
$env:INITIAL_ADMIN_USER = "hr_admin"
$env:INITIAL_ADMIN_PASSWORD = "choose-a-strong-password"
$env:SECRET_KEY = "generate-a-random-string-for-session-signing"
python webapp/app.py
```

Then open http://localhost:5000 and log in. Once logged in, add further
users from the **Settings** tab — the env vars are only read once (when the
`users` table is empty) so they can be unset afterward.

## Day-to-day use

```powershell
python webapp/app.py
```

Open the site, and use the tabs:
- **Upload Payroll** — upload the month's Excel (needs a `Salary Sheet_<Month> <Year>` tab and an `Employee_Details` tab).
- **Payslips** — download each employee's PDF, or all of them as one ZIP.
- **Compliance (PF/ESI/PT)** — see totals due, due dates, and mark each as filed once submitted on the government portal.
- **Settings** — company name/address/CIN/TAN/statutory codes, logo, and user accounts.

## Running it hosted (reachable by other HR staff)

This app reads `PORT` and `SECRET_KEY` from the environment and has no
hardcoded paths outside `webapp/`, so it can run on any host. Use
**waitress** (pure-Python WSGI server — included in `requirements.txt`) for
production instead of Flask's dev server:

```bash
waitress-serve --host=0.0.0.0 --port=$PORT app:app
```

Deploy this however you normally host small internal tools — e.g. a small
VPS (systemd service + reverse proxy for HTTPS), or a PaaS like Render/
Railway (set `INITIAL_ADMIN_USER`, `INITIAL_ADMIN_PASSWORD`, `SECRET_KEY`,
`PORT` as environment variables there, and point the start command at the
`waitress-serve` line above). Persist the `webapp/data/` and
`webapp/static/uploads/` folders across deploys — they hold the SQLite
database, uploaded payroll files, and the logo.

## Where the payroll logic comes from

All payroll extraction/reconciliation rules (CTC-inclusive gross, intern/
contractor flat-pay handling, Employee_Details join, Earned Leave fuzzy
name-matching) live in `../execution/generate_payslip.py` and are imported
directly by `payslip_data.py` — the CLI (`generate_payslip.py --mode full`)
and this web app always compute identical numbers because they share the
same code, not a re-implementation of it.
