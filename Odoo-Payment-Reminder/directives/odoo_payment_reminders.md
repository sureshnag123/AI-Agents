# Odoo Payment Reminder Automation

## Goal
Automatically send payment reminder emails to customers / debtors based on their invoice payment terms in Odoo 19. Supports pre-due reminders, overdue reminders, and escalation workflows.

## When to Use
- Daily automated reminders for overdue invoices
- Pre-due courtesy reminders before payment deadlines
- Generating accounts receivable aging reports
- Escalated reminder campaigns (gentle → firm → urgent)

## Inputs
| Input | Required | Description |
|-------|----------|-------------|
| Odoo URL | Yes | Your Odoo 19 instance URL (e.g., `https://mycompany.odoo.com`) |
| Odoo Database | Yes | Database name |
| Odoo Username | Yes | Login email or username |
| Odoo Password/API Key | Yes | Password or API key (API key recommended) |
| Reminder Mode | No | `overdue` (default) or `upcoming` |
| Days parameter | No | Min days overdue, or days ahead for upcoming |
| SMTP config | No | Only if sending via SMTP instead of Odoo mail |

## Architecture

```
┌──────────────────────────────────────────────────────────┐
│                   Orchestration Layer                     │
│              (You / Copilot / Scheduler)                 │
│                                                          │
│   1. Read this directive                                 │
│   2. Decide: overdue? upcoming? escalation?              │
│   3. Call the appropriate script                         │
│   4. Review results, handle errors                       │
└──────────────┬───────────────────────────┬───────────────┘
               │                           │
    ┌──────────▼──────────┐     ┌──────────▼──────────────┐
    │  odoo_connector.py  │     │ send_payment_reminders.py│
    │  (Odoo XML-RPC API) │     │ (Email builder + sender) │
    └─────────────────────┘     └──────────────────────────┘
               │                           │
    ┌──────────▼──────────┐     ┌──────────▼──────────────┐
    │ reminder_report.py  │     │  schedule_reminders.py   │
    │ (Aging reports)     │     │  (Daily scheduler)       │
    └─────────────────────┘     └──────────────────────────┘
```

## Scripts

### 1. `odoo_connector.py` — Odoo API Connection
Reusable XML-RPC connector for Odoo 19. Handles authentication and provides helpers for:
- Fetching overdue invoices
- Fetching upcoming-due invoices
- Reading partner (customer) email/phone
- Reading payment term details
- Logging reminder notes on invoice chatter

### 2. `send_payment_reminders.py` — Email Sender
Fetches invoices, groups by customer, builds reminder emails (HTML + plain text), and sends via:
- **Odoo Mail** (default) — uses Odoo's built-in `mail.mail` model
- **SMTP** (optional) — direct SMTP relay (Gmail, Outlook, etc.)

```bash
# Send overdue reminders
python execution/send_payment_reminders.py --mode overdue

# Pre-due reminders (next 7 days)
python execution/send_payment_reminders.py --mode upcoming --days 7

# Preview without sending
python execution/send_payment_reminders.py --mode overdue --dry-run

# Use SMTP instead of Odoo mail
python execution/send_payment_reminders.py --mode overdue --smtp
```

### 3. `schedule_reminders.py` — Scheduler
Runs reminders on a daily schedule. Two options:
- **APScheduler daemon** (keeps running in background)
- **One-shot mode** (`--once`) for use with Windows Task Scheduler or cron

```bash
# Background scheduler (daily at 9:00 AM)
python execution/schedule_reminders.py --hour 9 --minute 0

# One-shot for Task Scheduler
python execution/schedule_reminders.py --once

# Escalation mode (multiple severity levels)
python execution/schedule_reminders.py --once --escalate
```

### 4. `reminder_report.py` — Aging Report
Generates an accounts receivable aging report with buckets:
- Current (not yet due)
- 1-30 days overdue
- 31-60 days overdue
- 61-90 days overdue
- 90+ days overdue

```bash
python execution/reminder_report.py          # Console output
python execution/reminder_report.py --csv    # Export to CSV
python execution/reminder_report.py --json   # Export to JSON
```

## Setup Procedure

### Step 1: Install Python Dependencies
```bash
pip install -r requirements.txt
```

### Step 2: Configure Odoo Connection
Copy `.env.example` to `.env` and fill in:
```
ODOO_URL=https://your-company.odoo.com
ODOO_DB=your-database-name
ODOO_USERNAME=admin@yourcompany.com
ODOO_PASSWORD=your-api-key-here
```

**Recommended: Use an API key instead of password.**
In Odoo 19: Settings → Users → Select user → "API Keys" tab → Generate.

### Step 3: Configure Reminder Settings
```
COMPANY_NAME=Your Company Name
REMINDER_FROM_EMAIL=accounts@yourcompany.com
REMINDER_SENDER_NAME=Accounts Receivable
REMINDER_HOUR=9
REMINDER_MINUTE=0
```

### Step 4: Test Connection
```bash
python execution/odoo_connector.py
```
This prints your Odoo version, UID, and a list of overdue invoices.

### Step 5: Dry Run
```bash
python execution/send_payment_reminders.py --mode overdue --dry-run
```
Preview emails without actually sending them.

### Step 6: Send for Real
```bash
python execution/send_payment_reminders.py --mode overdue
```

### Step 7: Set Up Automation (Choose One)

#### Option A: APScheduler Daemon
```bash
python execution/schedule_reminders.py --hour 9 --minute 0
```
Runs continuously; sends reminders daily at the configured time.

#### Option B: Windows Task Scheduler
1. Open Task Scheduler
2. Create Basic Task → "Odoo Payment Reminders"
3. Trigger: Daily at 9:00 AM
4. Action: Start a program
   - Program: `python`
   - Arguments: `execution/schedule_reminders.py --once`
   - Start in: `<path-to-this-workspace>`

#### Option C: Linux Cron
```bash
0 9 * * * cd /path/to/workspace && python execution/schedule_reminders.py --once
```

## Odoo Payment Terms — How They Work

Odoo 19 payment terms (`account.payment.term`) define when invoices become due:
- **Immediate Payment** — Due on invoice date
- **30 Days** — Due 30 days after invoice date
- **2/10 Net 30** — 2% discount if paid in 10 days, full amount in 30
- **End of Following Month** — Due at end of next month
- **Custom** — Any combination of percentage + days/months

The script reads the payment term from each invoice and includes it in the reminder email for context.

## Escalation Strategy

When using `--escalate`, the scheduler runs 4 phases:

| Phase | Target | Tone | Frequency |
|-------|--------|------|-----------|
| 1 | Due in 3 days | Courtesy / gentle | Pre-due |
| 2 | 1-7 days overdue | Friendly reminder | Weekly |
| 3 | 14+ days overdue | Firm reminder | Bi-weekly |
| 4 | 30+ days overdue | Urgent / final notice | Monthly |

## Edge Cases & Learnings

- **No email on partner**: Invoice is skipped with a warning. Fix in Odoo: Contacts → edit partner → add email.
- **Partial payments**: `amount_residual` shows remaining balance. The reminder shows only what's still owed.
- **Multi-currency**: Each invoice uses its own currency. Emails group by customer (may mix currencies).
- **API rate limits**: Odoo XML-RPC has no hard rate limit, but avoid hitting it thousands of times per minute. The script batches requests.
- **API key vs password**: API keys are recommended. Generate in Odoo: Settings → Users → API Keys tab.
- **Odoo 19 compatibility**: Uses `account.move` model (unified invoice model since Odoo 13+). The `move_type='out_invoice'` filter ensures only customer invoices.

## Outputs
- Reminder emails sent via Odoo mail or SMTP
- Chatter notes logged on each invoice in Odoo
- Results JSON saved to `.tmp/reminder_results_YYYY-MM-DD.json`
- Aging report (console, CSV, JSON) via `reminder_report.py`
- Scheduler log at `.tmp/scheduler.log`
