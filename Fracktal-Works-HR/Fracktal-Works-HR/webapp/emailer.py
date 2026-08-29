"""Sends payslip PDFs by email via plain SMTP (stdlib smtplib + email —
no extra dependency). Works with Gmail Workspace, Office365, or any other
SMTP provider given host/port/username/password."""

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication


def send_payslip_email(smtp_settings, to_email, subject, body, pdf_bytes, pdf_filename):
    """Raises on failure (caller decides how to record/report it)."""
    msg = MIMEMultipart()
    from_name = smtp_settings.get("smtp_from_name") or smtp_settings["smtp_username"]
    msg["From"] = f"{from_name} <{smtp_settings['smtp_username']}>"
    msg["To"] = to_email
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain"))

    attachment = MIMEApplication(pdf_bytes, _subtype="pdf")
    attachment.add_header("Content-Disposition", "attachment", filename=pdf_filename)
    msg.attach(attachment)

    host = smtp_settings["smtp_host"]
    port = int(smtp_settings["smtp_port"])
    with smtplib.SMTP(host, port, timeout=20) as server:
        if smtp_settings.get("use_tls"):
            server.starttls()
        server.login(smtp_settings["smtp_username"], smtp_settings["smtp_password"])
        server.sendmail(smtp_settings["smtp_username"], [to_email], msg.as_string())


def send_password_reset_email(smtp_settings, to_email, reset_link, ttl_minutes):
    msg = MIMEMultipart()
    from_name = smtp_settings.get("smtp_from_name") or smtp_settings["smtp_username"]
    msg["From"] = f"{from_name} <{smtp_settings['smtp_username']}>"
    msg["To"] = to_email
    msg["Subject"] = "Password reset — Fracktal Works HR Payroll"
    body = (
        "A password reset was requested for this account on the Fracktal Works "
        "HR Payroll portal.\n\n"
        f"Reset your password using this link (valid for {ttl_minutes} minutes):\n{reset_link}\n\n"
        "If you didn't request this, you can ignore this email — your password "
        "will stay unchanged."
    )
    msg.attach(MIMEText(body, "plain"))

    host = smtp_settings["smtp_host"]
    port = int(smtp_settings["smtp_port"])
    with smtplib.SMTP(host, port, timeout=20) as server:
        if smtp_settings.get("use_tls"):
            server.starttls()
        server.login(smtp_settings["smtp_username"], smtp_settings["smtp_password"])
        server.sendmail(smtp_settings["smtp_username"], [to_email], msg.as_string())


def send_test_email(smtp_settings, to_email):
    body = "This is a test email from the Fracktal Works HR Payroll app to confirm your SMTP settings work."
    msg = MIMEMultipart()
    from_name = smtp_settings.get("smtp_from_name") or smtp_settings["smtp_username"]
    msg["From"] = f"{from_name} <{smtp_settings['smtp_username']}>"
    msg["To"] = to_email
    msg["Subject"] = "Test email — Fracktal Works HR Payroll"
    msg.attach(MIMEText(body, "plain"))

    host = smtp_settings["smtp_host"]
    port = int(smtp_settings["smtp_port"])
    with smtplib.SMTP(host, port, timeout=20) as server:
        if smtp_settings.get("use_tls"):
            server.starttls()
        server.login(smtp_settings["smtp_username"], smtp_settings["smtp_password"])
        server.sendmail(smtp_settings["smtp_username"], [to_email], msg.as_string())
