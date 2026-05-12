# Odoo Payment Reminder Agent

Automated payment reminder system for **Odoo 19** that sends email reminders to customers/debtors based on invoice payment terms.

## Features

- **Overdue reminders** — automatically email customers with past-due invoices
- **Pre-due reminders** — courtesy emails before payment deadlines
- **Escalation system** — gentle → firm → urgent escalation based on days overdue
- **Aging reports** — full AR aging breakdown (Current, 1-30, 31-60, 61-90, 90+ days)
- **Dual send method** — send via Odoo mail system or direct SMTP
- **Invoice chatter logging** — every reminder is logged on the invoice in Odoo
- **Scheduling** — built-in scheduler or integrate with Task Scheduler / cron

## Quick Start

### 1. Install Dependencies
```bash
pip install -r requirements.txt
```

### 2. Configure
```bash
cp .env.example .env
# Edit .env with your Odoo connection details
```

### 3. Test Connection
```bash
python execution/odoo_connector.py
```

### 4. Preview Reminders (Dry Run)
```bash
python execution/send_payment_reminders.py --mode overdue --dry-run
```

### 5. Send Reminders
```bash
# Overdue invoices
python execution/send_payment_reminders.py --mode overdue

# Upcoming due (next 7 days)
python execution/send_payment_reminders.py --mode upcoming --days 7
```

### 6. Generate Aging Report
```bash
python execution/reminder_report.py
python execution/reminder_report.py --csv   # Export to CSV
```

### 7. Schedule Daily Reminders
```bash
# Option A: Built-in scheduler (runs daily at 9 AM)
python execution/schedule_reminders.py --hour 9

# Option B: One-shot for Windows Task Scheduler / cron
python execution/schedule_reminders.py --once
```

## Odoo Setup Requirements

1. **Odoo 19** with Invoicing module enabled
2. **API Key** (recommended over password): Settings → Users → API Keys tab
3. **Outgoing mail server** configured in Odoo (Settings → Technical → Outgoing Mail Servers)
4. **Payment terms** configured on invoices (Invoicing → Configuration → Payment Terms)

## File Structure

```
odoo-payment-reminder/
├── AGENTS.md                   # AI agent system prompt
├── README.md                   # This file
├── requirements.txt            # Python dependencies
├── .env.example                # Environment variable template
├── .gitignore
├── directives/
│   └── odoo_payment_reminders.md   # SOP for the reminder workflow
├── execution/
│   ├── odoo_connector.py           # Odoo XML-RPC API connector
│   ├── send_payment_reminders.py   # Email builder + sender
│   ├── schedule_reminders.py       # Daily scheduler
│   └── reminder_report.py          # AR aging report generator
├── .github/agents/
│   └── odoo-payment-reminder.agent.md  # VS Code Copilot agent config
├── .vscode/
│   └── settings.json
└── .tmp/                       # Temporary files (auto-created)
```

## How It Works

1. **Connects to Odoo 19** via XML-RPC (`/xmlrpc/2/common` + `/xmlrpc/2/object`)
2. **Fetches invoices** from `account.move` where `move_type=out_invoice`, `state=posted`, and `payment_state` is `not_paid` or `partial`
3. **Groups invoices by customer** (partner)
4. **Builds HTML + plain-text emails** with invoice table, amounts, and payment terms
5. **Sends emails** via Odoo's `mail.mail` model or direct SMTP
6. **Logs a note** on each invoice's chatter in Odoo
7. **Saves results** to `.tmp/` for audit trail

## License

MIT
