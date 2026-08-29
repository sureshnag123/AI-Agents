#!/usr/bin/env python3
"""
Fracktal Works HR — Payroll & Compliance Web App

Upload the monthly payroll Excel, generate per-employee PDF payslips, and
track PF/ESI/PT filing status, all through a browser.

Run locally:
    python app.py
Run in production (any host — reads PORT/SECRET_KEY from env):
    waitress-serve --port=$PORT app:app
"""

import io
import os
import re
import time
import zipfile
import calendar
from pathlib import Path
from datetime import date

from dotenv import load_dotenv
from flask import Flask, render_template, request, redirect, url_for, send_file, flash, abort
from flask_login import login_required, current_user

import db
import auth
import payslip_data
import payslip_pdf
import emailer

WEBAPP_DIR = Path(__file__).parent
load_dotenv(WEBAPP_DIR.parent / ".env")
PROJECT_ROOT = WEBAPP_DIR.parent
UPLOAD_DIR = WEBAPP_DIR / "static" / "uploads" / "payroll"
LOGO_DIR = WEBAPP_DIR / "static" / "uploads" / "logo"

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dev-only-insecure-key-change-in-production")

auth.login_manager.init_app(app)
app.register_blueprint(auth.auth_bp)

db.init_db()

MONTH_NUM = {name.lower(): i for i, name in enumerate(calendar.month_name) if name}


@app.context_processor
def inject_company():
    return {"company": db.get_company_settings()}


def _parse_month_label(label):
    """Extract (month_number, year) from a label like 'June - 2026'."""
    year_match = re.search(r"(20\d{2})", label or "")
    year = int(year_match.group(1)) if year_match else None
    month_num = None
    for name, num in MONTH_NUM.items():
        if name in (label or "").lower():
            month_num = num
            break
    return month_num, year


def _due_dates(month_label):
    month_num, year = _parse_month_label(month_label)
    if not month_num or not year:
        return {}
    if month_num == 12:
        due_month, due_year = 1, year + 1
    else:
        due_month, due_year = month_num + 1, year
    return {item: date(due_year, due_month, day) for item, day in db.DUE_DAY.items()}


@app.route("/")
@login_required
def dashboard():
    months = db.list_months_with_summary()
    # Backfill months uploaded before month_summary existed.
    for m in months:
        if m["total_ctc"] is None:
            excel_full_path = WEBAPP_DIR / m["excel_path"]
            _, rows, _, _ = payslip_data.load_rows_for_excel(excel_full_path, m["salary_sheet_name"])
            summary = payslip_data.month_summary(rows)
            db.save_month_summary(m["id"], summary)
            m.update(summary)
    latest = months[0] if months else None
    return render_template("dashboard.html", months=months, latest=latest)


@app.route("/months/<int:month_id>/delete", methods=["POST"])
@login_required
def delete_month(month_id):
    month = db.get_month(month_id)
    if not month:
        abort(404)
    excel_path = db.delete_month(month_id)
    if excel_path:
        file_path = WEBAPP_DIR / excel_path
        if file_path.exists():
            file_path.unlink()
    flash(f"Deleted {month['month_label']} and its payslips/compliance records.", "success")
    return redirect(url_for("dashboard"))


@app.route("/upload", methods=["GET", "POST"])
@login_required
def upload():
    if request.method == "POST":
        file = request.files.get("payroll_file")
        if not file or file.filename == "":
            flash("Please choose an Excel file to upload.", "error")
            return redirect(url_for("upload"))

        UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
        safe_name = re.sub(r"[^\w.\-]", "_", file.filename)
        saved_path = UPLOAD_DIR / f"{int(time.time())}_{safe_name}"
        file.save(saved_path)

        salary_sheet_name = request.form.get("salary_sheet") or payslip_data.discover_salary_sheet(saved_path)
        if not salary_sheet_name:
            flash("Could not find a 'Salary Sheet_<Month> <Year>' tab in this workbook. "
                  "Please re-upload and specify the sheet name.", "error")
            return redirect(url_for("upload"))

        try:
            month_label, rows, unmatched, warnings = payslip_data.load_rows_for_excel(saved_path, salary_sheet_name)
        except ValueError as e:
            flash(str(e), "error")
            return redirect(url_for("upload"))

        if not rows:
            flash(f"No employee rows found in '{salary_sheet_name}'.", "error")
            return redirect(url_for("upload"))

        rel_path = str(saved_path.relative_to(WEBAPP_DIR))
        month_id = db.upsert_month(salary_sheet_name, month_label, rel_path, current_user.username)
        db.save_month_summary(month_id, payslip_data.month_summary(rows))

        flash(f"Processed {len(rows)} employee(s) for {month_label}.", "success")
        if unmatched:
            flash(f"{len(unmatched)} employee(s) had no match in Employee_Details "
                  f"(Employee ID / Bank Name left blank): {', '.join(unmatched)}", "warning")
        for w in warnings:
            flash(w, "warning")

        archive_folder = db.get_company_settings().get("archive_folder")
        if archive_folder:
            archived_count, archive_error = payslip_pdf.archive_payslips(
                rows, month_label, archive_folder, payslip_data.logo_path())
            if archive_error:
                flash(archive_error, "error")
            else:
                flash(f"Archived {archived_count} payslip PDF(s) to {archive_folder}\\{month_label}", "success")

        return redirect(url_for("payslips", month_id=month_id))

    return render_template("upload.html")


@app.route("/payslips")
@login_required
def payslips():
    months = db.list_months()
    if not months:
        return render_template("payslips.html", months=[], selected=None, rows=None)

    month_id = request.args.get("month_id", type=int) or months[0]["id"]
    month = db.get_month(month_id)
    if not month:
        abort(404)

    excel_full_path = WEBAPP_DIR / month["excel_path"]
    _, rows, unmatched, warnings = payslip_data.load_rows_for_excel(excel_full_path, month["salary_sheet_name"])
    email_status = db.get_email_status(month_id)

    return render_template("payslips.html", months=months, selected=month, rows=rows,
                            unmatched=unmatched, warnings=warnings, email_status=email_status)


@app.route("/payslips/<int:month_id>/pdf/<path:row_id>")
@login_required
def payslip_pdf_download(month_id, row_id):
    month = db.get_month(month_id)
    if not month:
        abort(404)
    excel_full_path = WEBAPP_DIR / month["excel_path"]
    _, rows, _, _ = payslip_data.load_rows_for_excel(excel_full_path, month["salary_sheet_name"])
    match = next((data for rid, data in rows if rid == row_id), None)
    if not match:
        abort(404)
    pdf_bytes = payslip_pdf.render_payslip_pdf(match, payslip_data.logo_path())
    filename = re.sub(r"[^\w.\-]", "_", f"{row_id}_{month['month_label']}.pdf")
    return send_file(io.BytesIO(pdf_bytes), mimetype="application/pdf",
                      as_attachment=True, download_name=filename)


@app.route("/payslips/<int:month_id>/zip")
@login_required
def payslip_zip_download(month_id):
    month = db.get_month(month_id)
    if not month:
        abort(404)
    excel_full_path = WEBAPP_DIR / month["excel_path"]
    _, rows, _, _ = payslip_data.load_rows_for_excel(excel_full_path, month["salary_sheet_name"])

    logo = payslip_data.logo_path()
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for row_id, data in rows:
            pdf_bytes = payslip_pdf.render_payslip_pdf(data, logo)
            zf.writestr(re.sub(r"[^\w.\-]", "_", f"{row_id}.pdf"), pdf_bytes)
    zip_buffer.seek(0)

    filename = re.sub(r"[^\w.\-]", "_", f"Payslips_{month['month_label']}.zip")
    return send_file(zip_buffer, mimetype="application/zip", as_attachment=True, download_name=filename)


def _send_one_payslip(month, data, row_id):
    """Sends one employee's payslip email and records the result. Returns
    (status, error) where status is 'sent', 'failed', or 'skipped'."""
    email = (data.get("email") or "").strip()
    if not email:
        return "skipped", "no email address on file"

    smtp_settings = db.get_email_settings()
    if not smtp_settings or not smtp_settings.get("smtp_username"):
        return "failed", "SMTP not configured — set it up in Settings first"

    subject = f"Payslip for {month['month_label']} — {data['company_name']}"
    body = (f"Dear {data['name']},\n\n"
            f"Please find attached your payslip for {month['month_label']}.\n\n"
            f"Regards,\nHR Team\n{data['company_name']}")
    pdf_bytes = payslip_pdf.render_payslip_pdf(data, payslip_data.logo_path())
    pdf_filename = re.sub(r"[^\w.\-]", "_", f"{row_id}_{month['month_label']}.pdf")

    try:
        emailer.send_payslip_email(smtp_settings, email, subject, body, pdf_bytes, pdf_filename)
    except Exception as e:
        db.record_email_result(month["id"], row_id, email, "failed", str(e))
        return "failed", str(e)

    db.record_email_result(month["id"], row_id, email, "sent")
    return "sent", None


@app.route("/payslips/<int:month_id>/email-all", methods=["POST"])
@login_required
def email_all_payslips(month_id):
    month = db.get_month(month_id)
    if not month:
        abort(404)
    excel_full_path = WEBAPP_DIR / month["excel_path"]
    _, rows, _, _ = payslip_data.load_rows_for_excel(excel_full_path, month["salary_sheet_name"])

    sent, failed, skipped = 0, 0, 0
    failures = []
    for row_id, data in rows:
        status, error = _send_one_payslip(month, data, row_id)
        if status == "sent":
            sent += 1
        elif status == "skipped":
            skipped += 1
        else:
            failed += 1
            failures.append(f"{data['name']}: {error}")

    flash(f"Emails: {sent} sent, {failed} failed, {skipped} skipped (no email on file).",
          "success" if failed == 0 else "warning")
    for f in failures[:10]:
        flash(f, "error")

    return redirect(url_for("payslips", month_id=month_id))


@app.route("/payslips/<int:month_id>/email/<path:row_id>", methods=["POST"])
@login_required
def email_one_payslip(month_id, row_id):
    month = db.get_month(month_id)
    if not month:
        abort(404)
    excel_full_path = WEBAPP_DIR / month["excel_path"]
    _, rows, _, _ = payslip_data.load_rows_for_excel(excel_full_path, month["salary_sheet_name"])
    match = next((data for rid, data in rows if rid == row_id), None)
    if not match:
        abort(404)

    status, error = _send_one_payslip(month, match, row_id)
    if status == "sent":
        flash(f"Payslip emailed to {match['name']} ({match['email']}).", "success")
    elif status == "skipped":
        flash(f"Could not email {match['name']}: {error}.", "warning")
    else:
        flash(f"Failed to email {match['name']}: {error}", "error")

    return redirect(url_for("payslips", month_id=month_id))


@app.route("/compliance", methods=["GET", "POST"])
@login_required
def compliance():
    months = db.list_months()
    if not months:
        return render_template("compliance.html", months=[], selected=None)

    month_id = request.args.get("month_id", type=int) or months[0]["id"]
    month = db.get_month(month_id)
    if not month:
        abort(404)

    if request.method == "POST":
        item = request.form.get("item")
        filed = request.form.get("filed") == "1"
        db.set_compliance_status(month_id, item, filed, current_user.username)
        return redirect(url_for("compliance", month_id=month_id))

    excel_full_path = WEBAPP_DIR / month["excel_path"]
    _, rows, _, _ = payslip_data.load_rows_for_excel(excel_full_path, month["salary_sheet_name"])
    totals = payslip_data.compliance_totals(rows)
    status = db.get_compliance_status(month_id)
    due_dates = _due_dates(month["month_label"])
    today = date.today()
    overdue = {
        item: bool(due_dates.get(item)) and not status[item]["filed"] and today > due_dates[item]
        for item in db.COMPLIANCE_ITEMS
    }

    ecr_text, ecr_included, ecr_skipped, ecr_warnings = payslip_data.build_ecr_text(rows)

    return render_template("compliance.html", months=months, selected=month, totals=totals,
                            status=status, due_dates=due_dates, overdue=overdue, today=today,
                            ecr_included=ecr_included, ecr_skipped=ecr_skipped, ecr_warnings=ecr_warnings)


@app.route("/compliance/<int:month_id>/ecr")
@login_required
def download_ecr(month_id):
    month = db.get_month(month_id)
    if not month:
        abort(404)
    excel_full_path = WEBAPP_DIR / month["excel_path"]
    _, rows, _, _ = payslip_data.load_rows_for_excel(excel_full_path, month["salary_sheet_name"])
    text, included, _skipped, _warnings = payslip_data.build_ecr_text(rows)

    if included == 0:
        flash("No PF-enrolled employees with valid 12-digit UANs found — nothing to generate.", "error")
        return redirect(url_for("compliance", month_id=month_id))

    filename = re.sub(r"[^\w.\-]", "_", f"ECR_{month['month_label']}.txt")
    return send_file(io.BytesIO(text.encode("utf-8")), mimetype="text/plain",
                      as_attachment=True, download_name=filename)


@app.route("/settings", methods=["GET", "POST"])
@login_required
def settings():
    if request.method == "POST":
        form_type = request.form.get("form_type")

        if form_type == "company":
            current_settings = db.get_company_settings()
            fields = {
                "name": request.form.get("name", "").strip(),
                "address": request.form.get("address", "").strip(),
                "cin": request.form.get("cin", "").strip(),
                "tan": request.form.get("tan", "").strip(),
                "pf_code": request.form.get("pf_code", "").strip(),
                "esi_code": request.form.get("esi_code", "").strip(),
                "pt_reg": request.form.get("pt_reg", "").strip(),
                "archive_folder": request.form.get("archive_folder", "").strip(),
                "eps_excluded_uans": request.form.get("eps_excluded_uans", "").strip(),
                "logo_path": current_settings["logo_path"],
            }
            logo_file = request.files.get("logo")
            if logo_file and logo_file.filename:
                LOGO_DIR.mkdir(parents=True, exist_ok=True)
                ext = Path(logo_file.filename).suffix or ".png"
                logo_saved_path = LOGO_DIR / f"logo{ext}"
                logo_file.save(logo_saved_path)
                fields["logo_path"] = str(logo_saved_path.relative_to(WEBAPP_DIR / "static"))

            db.update_company_settings(fields)
            flash("Company settings updated.", "success")

        elif form_type == "user":
            username = request.form.get("new_username", "").strip()
            password = request.form.get("new_password", "")
            if username and password:
                if db.get_user_by_username(username):
                    flash("That username already exists.", "error")
                else:
                    db.create_user(username, password)
                    flash(f"User '{username}' created.", "success")

        elif form_type == "email":
            current_email_settings = db.get_email_settings() or {}
            fields = {
                "smtp_host": request.form.get("smtp_host", "").strip(),
                "smtp_port": request.form.get("smtp_port", "587").strip() or "587",
                "smtp_username": request.form.get("smtp_username", "").strip(),
                "smtp_from_name": request.form.get("smtp_from_name", "").strip(),
                "use_tls": 1 if request.form.get("use_tls") == "on" else 0,
            }
            new_password = request.form.get("smtp_password", "")
            fields["smtp_password"] = new_password if new_password else current_email_settings.get("smtp_password", "")
            db.update_email_settings(fields)
            flash("Email (SMTP) settings updated.", "success")

        elif form_type == "test_email":
            test_to = request.form.get("test_email_to", "").strip()
            smtp_settings = db.get_email_settings()
            if not test_to:
                flash("Enter an email address to send the test to.", "error")
            elif not smtp_settings or not smtp_settings.get("smtp_username"):
                flash("Save your SMTP settings first.", "error")
            else:
                try:
                    emailer.send_test_email(smtp_settings, test_to)
                    flash(f"Test email sent to {test_to}.", "success")
                except Exception as e:
                    flash(f"Test email failed: {e}", "error")

        return redirect(url_for("settings"))

    return render_template("settings.html", settings=db.get_company_settings(),
                            email_settings=db.get_email_settings(), users=db.list_users())


@app.route("/settings/backup")
@login_required
def download_backup():
    """Download the raw SQLite database — the hosting free tier's storage is
    not permanent, so download this before any planned redeploy."""
    return send_file(
        db.DB_PATH,
        as_attachment=True,
        download_name=f"fracktal_hr_backup_{date.today().isoformat()}.db",
    )


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=True, use_reloader=False)
