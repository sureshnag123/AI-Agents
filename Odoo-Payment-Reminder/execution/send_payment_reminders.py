#!/usr/bin/env python3
"""
Send Payment Reminder Emails via Odoo

Fetches overdue and upcoming-due invoices from Odoo 19, groups them by
customer, builds reminder emails from templates, and sends them through
Odoo's mail system (or optionally SMTP directly).

Usage:
    # Send reminders for all overdue invoices
    python send_payment_reminders.py --mode overdue

    # Send pre-due reminders (invoices due in next 7 days)
    python send_payment_reminders.py --mode upcoming --days 7

    # Dry-run (preview emails without sending)
    python send_payment_reminders.py --mode overdue --dry-run

    # Send via SMTP instead of Odoo mail
    python send_payment_reminders.py --mode overdue --smtp
"""

import os
import sys
import json
import argparse
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import date, timedelta
from collections import defaultdict
from typing import Optional

from dotenv import load_dotenv

load_dotenv()

# Add execution directory to path so we can import odoo_connector
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from odoo_connector import OdooConnector


# ── Email Templates ─────────────────────────────────────────────────

OVERDUE_SUBJECT = "Payment Reminder: Invoice {invoice_number} is overdue"
OVERDUE_BODY = """Dear {customer_name},

We hope this message finds you well.

This is a friendly reminder that the following invoice(s) are past due:

{invoice_table}

Total Outstanding: {currency} {total_outstanding}

Original Payment Terms: {payment_terms}

We kindly request that you arrange payment at your earliest convenience. If you have already made the payment, please disregard this reminder and accept our apologies.

If you have any questions regarding these invoices or need to discuss payment arrangements, please don't hesitate to contact us.

Thank you for your prompt attention to this matter.

Best regards,
{company_name}
{sender_name}
{sender_email}
"""

UPCOMING_SUBJECT = "Upcoming Payment Due: Invoice {invoice_number}"
UPCOMING_BODY = """Dear {customer_name},

We hope this message finds you well.

This is a courtesy reminder that the following invoice(s) will be due soon:

{invoice_table}

Total Due: {currency} {total_outstanding}

Payment Terms: {payment_terms}

Please ensure payment is arranged by the due date to avoid any late fees. If you have already scheduled the payment, thank you!

If you have any questions, please feel free to reach out.

Best regards,
{company_name}
{sender_name}
{sender_email}
"""

# HTML version for richer emails
OVERDUE_HTML = """
<div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
    <h2 style="color: #c0392b;">Payment Reminder</h2>
    <p>Dear {customer_name},</p>
    <p>This is a friendly reminder that the following invoice(s) are <strong>past due</strong>:</p>

    <table style="width: 100%; border-collapse: collapse; margin: 16px 0;">
        <thead>
            <tr style="background: #ecf0f1;">
                <th style="padding: 8px; border: 1px solid #ddd; text-align: left;">Invoice</th>
                <th style="padding: 8px; border: 1px solid #ddd; text-align: left;">Invoice Date</th>
                <th style="padding: 8px; border: 1px solid #ddd; text-align: left;">Due Date</th>
                <th style="padding: 8px; border: 1px solid #ddd; text-align: right;">Amount Due</th>
                <th style="padding: 8px; border: 1px solid #ddd; text-align: left;">Days Overdue</th>
            </tr>
        </thead>
        <tbody>
            {invoice_rows}
        </tbody>
    </table>

    <p><strong>Total Outstanding: {currency} {total_outstanding}</strong></p>
    <p><em>Payment Terms: {payment_terms}</em></p>

    <p>We kindly request that you arrange payment at your earliest convenience.
       If you have already made the payment, please disregard this reminder.</p>

    <p>Best regards,<br/>
    <strong>{sender_name}</strong><br/>
    {company_name}<br/>
    {sender_email}</p>
</div>
"""

UPCOMING_HTML = """
<div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
    <h2 style="color: #2980b9;">Upcoming Payment Reminder</h2>
    <p>Dear {customer_name},</p>
    <p>This is a courtesy reminder that the following invoice(s) will be <strong>due soon</strong>:</p>

    <table style="width: 100%; border-collapse: collapse; margin: 16px 0;">
        <thead>
            <tr style="background: #ecf0f1;">
                <th style="padding: 8px; border: 1px solid #ddd; text-align: left;">Invoice</th>
                <th style="padding: 8px; border: 1px solid #ddd; text-align: left;">Invoice Date</th>
                <th style="padding: 8px; border: 1px solid #ddd; text-align: left;">Due Date</th>
                <th style="padding: 8px; border: 1px solid #ddd; text-align: right;">Amount Due</th>
                <th style="padding: 8px; border: 1px solid #ddd; text-align: left;">Days Until Due</th>
            </tr>
        </thead>
        <tbody>
            {invoice_rows}
        </tbody>
    </table>

    <p><strong>Total Due: {currency} {total_outstanding}</strong></p>
    <p><em>Payment Terms: {payment_terms}</em></p>

    <p>Please ensure payment is arranged by the due date. Thank you!</p>

    <p>Best regards,<br/>
    <strong>{sender_name}</strong><br/>
    {company_name}<br/>
    {sender_email}</p>
</div>
"""


# ── Core Logic ──────────────────────────────────────────────────────

def group_invoices_by_partner(invoices: list[dict]) -> dict:
    """Group invoices by partner (customer) ID."""
    grouped = defaultdict(list)
    for inv in invoices:
        partner_id = inv["partner_id"][0]
        grouped[partner_id].append(inv)
    return dict(grouped)


def build_invoice_table_text(invoices: list[dict], mode: str) -> str:
    """Build a plain-text table of invoices."""
    lines = []
    header = f"{'Invoice':<20} {'Date':<12} {'Due Date':<12} {'Amount Due':>12} {'Status':<15}"
    lines.append(header)
    lines.append("-" * len(header))

    for inv in invoices:
        due_date = inv["invoice_date_due"]
        if mode == "overdue":
            days = (date.today() - date.fromisoformat(str(due_date))).days
            status = f"{days} days overdue"
        else:
            days = (date.fromisoformat(str(due_date)) - date.today()).days
            status = f"Due in {days} days"

        currency = inv["currency_id"][1] if inv.get("currency_id") else ""
        lines.append(
            f"{inv['name']:<20} {str(inv['invoice_date']):<12} "
            f"{str(due_date):<12} {currency} {inv['amount_residual']:>8.2f} {status:<15}"
        )

    return "\n".join(lines)


def build_invoice_rows_html(invoices: list[dict], mode: str) -> str:
    """Build HTML table rows for invoices."""
    rows = []
    for inv in invoices:
        due_date = inv["invoice_date_due"]
        currency = inv["currency_id"][1] if inv.get("currency_id") else ""

        if mode == "overdue":
            days = (date.today() - date.fromisoformat(str(due_date))).days
            days_text = f'<span style="color: #c0392b; font-weight: bold;">{days} days</span>'
        else:
            days = (date.fromisoformat(str(due_date)) - date.today()).days
            days_text = f"{days} days"

        rows.append(f"""
            <tr>
                <td style="padding: 8px; border: 1px solid #ddd;">{inv['name']}</td>
                <td style="padding: 8px; border: 1px solid #ddd;">{inv['invoice_date']}</td>
                <td style="padding: 8px; border: 1px solid #ddd;">{due_date}</td>
                <td style="padding: 8px; border: 1px solid #ddd; text-align: right;">{currency} {inv['amount_residual']:.2f}</td>
                <td style="padding: 8px; border: 1px solid #ddd;">{days_text}</td>
            </tr>
        """)
    return "\n".join(rows)


def send_via_odoo(odoo: OdooConnector, partner: dict, subject: str, body_html: str, invoice_ids: list[int]) -> bool:
    """Send email through Odoo's built-in mail system and log on invoice."""
    try:
        # Create and send mail.mail record
        mail_id = odoo.create("mail.mail", {
            "subject": subject,
            "body_html": body_html,
            "email_from": os.getenv("REMINDER_FROM_EMAIL", odoo.username),
            "email_to": partner.get("email", ""),
            "auto_delete": True,
        })
        # Send immediately
        odoo._execute("mail.mail", "send", [mail_id])

        # Log a note on each invoice
        for inv_id in invoice_ids:
            odoo.log_reminder_on_invoice(
                inv_id,
                f"📧 Payment reminder email sent to {partner.get('email', 'N/A')}"
            )

        return True
    except Exception as e:
        print(f"  ❌ Odoo mail error: {e}")
        return False


def send_via_smtp(partner: dict, subject: str, body_text: str, body_html: str) -> bool:
    """Send email via direct SMTP."""
    smtp_host = os.getenv("SMTP_HOST", "")
    smtp_port = int(os.getenv("SMTP_PORT", "587"))
    smtp_user = os.getenv("SMTP_USER", "")
    smtp_pass = os.getenv("SMTP_PASSWORD", "")
    from_email = os.getenv("REMINDER_FROM_EMAIL", smtp_user)

    if not all([smtp_host, smtp_user, smtp_pass]):
        print("  ❌ SMTP not configured. Set SMTP_HOST, SMTP_USER, SMTP_PASSWORD in .env")
        return False

    to_email = partner.get("email", "")
    if not to_email:
        print(f"  ⚠ No email for partner {partner.get('name', 'Unknown')}")
        return False

    try:
        msg = MIMEMultipart("alternative")
        msg["Subject"] = subject
        msg["From"] = from_email
        msg["To"] = to_email

        msg.attach(MIMEText(body_text, "plain"))
        msg.attach(MIMEText(body_html, "html"))

        with smtplib.SMTP(smtp_host, smtp_port) as server:
            server.ehlo()
            if smtp_port == 587:
                server.starttls()
            server.login(smtp_user, smtp_pass)
            server.sendmail(from_email, to_email, msg.as_string())

        return True
    except Exception as e:
        print(f"  ❌ SMTP error: {e}")
        return False


def process_reminders(
    mode: str = "overdue",
    days: int = 0,
    dry_run: bool = False,
    use_smtp: bool = False,
    escalation_days: Optional[list[int]] = None,
):
    """
    Main entry point: fetch invoices, build emails, send reminders.

    Args:
        mode: 'overdue' or 'upcoming'
        days: For overdue = minimum days overdue; for upcoming = days ahead to look
        dry_run: If True, print emails instead of sending
        use_smtp: If True, use SMTP instead of Odoo mail
        escalation_days: List of day thresholds for escalated reminders (e.g. [7, 14, 30, 60])
    """
    odoo = OdooConnector()
    print(f"✓ Connected to Odoo {odoo.version().get('server_version', '?')}")

    company_name = os.getenv("COMPANY_NAME", "Our Company")
    sender_name = os.getenv("REMINDER_SENDER_NAME", "Accounts Receivable")
    sender_email = os.getenv("REMINDER_FROM_EMAIL", odoo.username)

    # Fetch invoices
    if mode == "overdue":
        invoices = odoo.get_overdue_invoices(days_overdue=days)
        subject_tpl = OVERDUE_SUBJECT
        body_tpl = OVERDUE_BODY
        html_tpl = OVERDUE_HTML
        print(f"Found {len(invoices)} overdue invoice(s) (>{days} days past due)")
    else:
        days = days or 7
        invoices = odoo.get_upcoming_due_invoices(days_ahead=days)
        subject_tpl = UPCOMING_SUBJECT
        body_tpl = UPCOMING_BODY
        html_tpl = UPCOMING_HTML
        print(f"Found {len(invoices)} invoice(s) due in next {days} days")

    if not invoices:
        print("No invoices to process. Done.")
        return

    # Group by partner
    grouped = group_invoices_by_partner(invoices)
    print(f"Grouped into {len(grouped)} customer(s)\n")

    sent_count = 0
    failed_count = 0
    skipped_count = 0

    for partner_id, partner_invoices in grouped.items():
        # Get partner contact
        partner = odoo.get_partner_email(partner_id)
        customer_name = partner.get("name", "Customer")
        email = partner.get("email", "")

        if not email:
            print(f"⚠ SKIP: {customer_name} — no email address on file")
            skipped_count += 1
            continue

        # Get payment terms from first invoice
        payment_terms = "Standard"
        first_inv = partner_invoices[0]
        if first_inv.get("invoice_payment_term_id"):
            term = odoo.get_payment_term_details(first_inv["invoice_payment_term_id"][0])
            payment_terms = term.get("name", "Standard")

        # Calculate totals
        currency = partner_invoices[0]["currency_id"][1] if partner_invoices[0].get("currency_id") else ""
        total_outstanding = sum(inv["amount_residual"] for inv in partner_invoices)
        invoice_ids = [inv["id"] for inv in partner_invoices]

        # Build email content
        invoice_table = build_invoice_table_text(partner_invoices, mode)
        invoice_rows = build_invoice_rows_html(partner_invoices, mode)
        invoice_number = partner_invoices[0]["name"] if len(partner_invoices) == 1 else f"{len(partner_invoices)} invoices"

        fmt = {
            "customer_name": customer_name,
            "invoice_number": invoice_number,
            "invoice_table": invoice_table,
            "invoice_rows": invoice_rows,
            "currency": currency,
            "total_outstanding": f"{total_outstanding:,.2f}",
            "payment_terms": payment_terms,
            "company_name": company_name,
            "sender_name": sender_name,
            "sender_email": sender_email,
        }

        subject = subject_tpl.format(**fmt)
        body_text = body_tpl.format(**fmt)
        body_html = html_tpl.format(**fmt)

        print(f"{'[DRY-RUN] ' if dry_run else ''}📧 {customer_name} <{email}>")
        print(f"   Subject: {subject}")
        print(f"   Invoices: {len(partner_invoices)} | Total: {currency} {total_outstanding:,.2f}")

        if dry_run:
            print(f"   --- Preview ---")
            print(f"   {body_text[:200]}...")
            print()
            sent_count += 1
            continue

        # Send
        if use_smtp:
            success = send_via_smtp(partner, subject, body_text, body_html)
        else:
            success = send_via_odoo(odoo, partner, subject, body_html, invoice_ids)

        if success:
            print(f"   ✓ Sent successfully")
            sent_count += 1
        else:
            print(f"   ✗ Failed to send")
            failed_count += 1

        print()

    # Summary
    print("=" * 50)
    print(f"SUMMARY — {mode.upper()} REMINDERS")
    print(f"  Sent:    {sent_count}")
    print(f"  Failed:  {failed_count}")
    print(f"  Skipped: {skipped_count} (no email)")
    print(f"  Total:   {len(grouped)} customers, {len(invoices)} invoices")
    print("=" * 50)

    # Save results to .tmp
    results = {
        "mode": mode,
        "date": str(date.today()),
        "sent": sent_count,
        "failed": failed_count,
        "skipped": skipped_count,
        "customers": len(grouped),
        "invoices": len(invoices),
        "dry_run": dry_run,
    }
    tmp_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), ".tmp")
    os.makedirs(tmp_dir, exist_ok=True)
    results_file = os.path.join(tmp_dir, f"reminder_results_{date.today()}.json")
    with open(results_file, "w") as f:
        json.dump(results, f, indent=2)
    print(f"\nResults saved to {results_file}")


# ── CLI ─────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Send payment reminder emails via Odoo 19"
    )
    parser.add_argument(
        "--mode",
        choices=["overdue", "upcoming"],
        default="overdue",
        help="Type of reminders to send (default: overdue)",
    )
    parser.add_argument(
        "--days",
        type=int,
        default=0,
        help="For overdue: min days overdue (0=all). For upcoming: days ahead (default 7).",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Preview emails without sending",
    )
    parser.add_argument(
        "--smtp",
        action="store_true",
        help="Send via SMTP instead of Odoo mail system",
    )
    args = parser.parse_args()

    process_reminders(
        mode=args.mode,
        days=args.days,
        dry_run=args.dry_run,
        use_smtp=args.smtp,
    )


if __name__ == "__main__":
    main()
