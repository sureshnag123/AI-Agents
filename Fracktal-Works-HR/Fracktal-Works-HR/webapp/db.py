"""SQLite persistence for the payroll & compliance web app.

Tables:
  users              — login accounts (Flask-Login)
  months             — one row per uploaded payroll month
  month_summary      — aggregate totals for a month, snapshotted at upload time
                        (so the dashboard doesn't need to re-parse every Excel
                        file on every page load, and a month's numbers stay
                        stable even if the underlying file is later replaced)
  compliance_status  — PF/ESI/PT filed/not-filed checklist, per month
  company_settings   — single-row company profile (name, address, statutory codes, logo)
"""

import os
import sqlite3
from datetime import datetime
from pathlib import Path

from werkzeug.security import generate_password_hash

DATA_DIR = Path(__file__).parent / "data"
DB_PATH = DATA_DIR / "fracktal_hr.db"

# Defaults confirmed against the company's registration documents this session.
DEFAULT_COMPANY = {
    "name": "Fracktal Works Private Limited",
    "address": "No. 3, 50ft Laggere Main Road, Chowdeshwari Nagar, Bengaluru 560058, Karnataka, India",
    "cin": "U30009KA2013PTC070124",
    "tan": "BLRF03155F",
    "pf_code": "PYKRP1426103000",
    "esi_code": "49000552030001099",
    "pt_reg": "356662288",
    "logo_path": "uploads/logo/logo.png",
    "archive_folder": "",
    # Confirmed via a real EPFO portal ECR upload rejection (RFE-21, June 2026):
    # these two UANs are recorded by EPFO as EPS-excluded (joined EPF after
    # 1 Sept 2014 on wages above the then-ceiling under a prior employer) —
    # Guruprasad CD and Dheeraj Kumar M N.
    "eps_excluded_uans": "102205396062,102248036176",
}

COMPLIANCE_ITEMS = ["PF", "ESI", "PT"]
# Due date (day-of-month, following month) per compliance item — Karnataka PT.
DUE_DAY = {"PF": 15, "ESI": 15, "PT": 20}


def get_db():
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def init_db():
    conn = get_db()
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL,
            created_at TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS months (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            salary_sheet_name TEXT UNIQUE NOT NULL,
            month_label TEXT NOT NULL,
            excel_path TEXT NOT NULL,
            uploaded_by TEXT,
            uploaded_at TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS month_summary (
            month_id INTEGER PRIMARY KEY REFERENCES months(id) ON DELETE CASCADE,
            employee_count INTEGER NOT NULL,
            total_ctc REAL NOT NULL,
            total_deductions REAL NOT NULL,
            total_net_pay REAL NOT NULL,
            employee_pf REAL NOT NULL,
            employer_pf REAL NOT NULL,
            employee_esi REAL NOT NULL,
            employer_esi REAL NOT NULL,
            pt REAL NOT NULL
        );

        CREATE TABLE IF NOT EXISTS compliance_status (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            month_id INTEGER NOT NULL REFERENCES months(id) ON DELETE CASCADE,
            item TEXT NOT NULL CHECK (item IN ('PF', 'ESI', 'PT')),
            filed INTEGER NOT NULL DEFAULT 0,
            filed_by TEXT,
            filed_at TEXT,
            UNIQUE(month_id, item)
        );

        CREATE TABLE IF NOT EXISTS company_settings (
            id INTEGER PRIMARY KEY CHECK (id = 1),
            name TEXT NOT NULL,
            address TEXT NOT NULL,
            cin TEXT,
            tan TEXT,
            pf_code TEXT,
            esi_code TEXT,
            pt_reg TEXT,
            logo_path TEXT,
            archive_folder TEXT,
            eps_excluded_uans TEXT
        );

        CREATE TABLE IF NOT EXISTS email_settings (
            id INTEGER PRIMARY KEY CHECK (id = 1),
            smtp_host TEXT,
            smtp_port INTEGER,
            smtp_username TEXT,
            smtp_password TEXT,
            smtp_from_name TEXT,
            use_tls INTEGER NOT NULL DEFAULT 1
        );

        CREATE TABLE IF NOT EXISTS email_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            month_id INTEGER NOT NULL REFERENCES months(id) ON DELETE CASCADE,
            row_id TEXT NOT NULL,
            email_address TEXT,
            status TEXT NOT NULL CHECK (status IN ('sent', 'failed')),
            sent_at TEXT NOT NULL,
            error TEXT,
            UNIQUE(month_id, row_id)
        );

        CREATE TABLE IF NOT EXISTS password_resets (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
            token TEXT UNIQUE NOT NULL,
            created_at TEXT NOT NULL,
            expires_at TEXT NOT NULL,
            used INTEGER NOT NULL DEFAULT 0
        );
    """)
    conn.commit()

    # Migrate: add columns to a company_settings table created before they existed.
    existing_cols = {row["name"] for row in conn.execute("PRAGMA table_info(company_settings)").fetchall()}
    if "archive_folder" not in existing_cols:
        conn.execute("ALTER TABLE company_settings ADD COLUMN archive_folder TEXT DEFAULT ''")
        conn.commit()
    if "eps_excluded_uans" not in existing_cols:
        conn.execute("ALTER TABLE company_settings ADD COLUMN eps_excluded_uans TEXT DEFAULT ''")
        conn.commit()
        # Backfill the two UANs confirmed EPS-excluded via a real EPFO portal
        # rejection this session, onto whatever company_settings row already exists.
        conn.execute(
            "UPDATE company_settings SET eps_excluded_uans = ? WHERE id = 1 AND (eps_excluded_uans IS NULL OR eps_excluded_uans = '')",
            (DEFAULT_COMPANY["eps_excluded_uans"],),
        )
        conn.commit()

    # Seed company settings once.
    row = conn.execute("SELECT id FROM company_settings WHERE id = 1").fetchone()
    if row is None:
        conn.execute(
            """INSERT INTO company_settings
                   (id, name, address, cin, tan, pf_code, esi_code, pt_reg, logo_path, archive_folder, eps_excluded_uans)
               VALUES (1, :name, :address, :cin, :tan, :pf_code, :esi_code, :pt_reg, :logo_path, :archive_folder, :eps_excluded_uans)""",
            DEFAULT_COMPANY,
        )
        conn.commit()

    # Seed email settings once (Gmail-style defaults, fully editable).
    row = conn.execute("SELECT id FROM email_settings WHERE id = 1").fetchone()
    if row is None:
        conn.execute(
            """INSERT INTO email_settings (id, smtp_host, smtp_port, smtp_username, smtp_password,
                                            smtp_from_name, use_tls)
               VALUES (1, 'smtp.gmail.com', 587, '', '', :from_name, 1)""",
            {"from_name": DEFAULT_COMPANY["name"]},
        )
        conn.commit()

    # Seed the first admin user from env vars, only if no users exist yet.
    user_count = conn.execute("SELECT COUNT(*) AS n FROM users").fetchone()["n"]
    if user_count == 0:
        admin_user = os.environ.get("INITIAL_ADMIN_USER")
        admin_password = os.environ.get("INITIAL_ADMIN_PASSWORD")
        if admin_user and admin_password:
            conn.execute(
                "INSERT INTO users (username, password_hash, created_at) VALUES (?, ?, ?)",
                (admin_user, generate_password_hash(admin_password), datetime.utcnow().isoformat()),
            )
            conn.commit()
    conn.close()


def get_company_settings():
    conn = get_db()
    row = conn.execute("SELECT * FROM company_settings WHERE id = 1").fetchone()
    conn.close()
    return dict(row) if row else dict(DEFAULT_COMPANY)


def update_company_settings(fields):
    conn = get_db()
    conn.execute(
        """UPDATE company_settings SET name=:name, address=:address, cin=:cin, tan=:tan,
           pf_code=:pf_code, esi_code=:esi_code, pt_reg=:pt_reg, logo_path=:logo_path,
           archive_folder=:archive_folder, eps_excluded_uans=:eps_excluded_uans WHERE id = 1""",
        fields,
    )
    conn.commit()
    conn.close()


def get_email_settings():
    conn = get_db()
    row = conn.execute("SELECT * FROM email_settings WHERE id = 1").fetchone()
    conn.close()
    return dict(row) if row else None


def update_email_settings(fields):
    conn = get_db()
    conn.execute(
        """UPDATE email_settings SET smtp_host=:smtp_host, smtp_port=:smtp_port,
           smtp_username=:smtp_username, smtp_password=:smtp_password,
           smtp_from_name=:smtp_from_name, use_tls=:use_tls WHERE id = 1""",
        fields,
    )
    conn.commit()
    conn.close()


def get_email_status(month_id):
    """row_id -> {status, sent_at, error} for every email attempt recorded for this month."""
    conn = get_db()
    rows = conn.execute("SELECT * FROM email_log WHERE month_id = ?", (month_id,)).fetchall()
    conn.close()
    return {r["row_id"]: dict(r) for r in rows}


def record_email_result(month_id, row_id, email_address, status, error=None):
    conn = get_db()
    conn.execute(
        """INSERT INTO email_log (month_id, row_id, email_address, status, sent_at, error)
           VALUES (?, ?, ?, ?, ?, ?)
           ON CONFLICT(month_id, row_id) DO UPDATE SET
               email_address=excluded.email_address, status=excluded.status,
               sent_at=excluded.sent_at, error=excluded.error""",
        (month_id, row_id, email_address, status, datetime.utcnow().isoformat(), error),
    )
    conn.commit()
    conn.close()


def list_months():
    conn = get_db()
    rows = conn.execute("SELECT * FROM months ORDER BY uploaded_at DESC").fetchall()
    conn.close()
    return [dict(r) for r in rows]


def get_month(month_id):
    conn = get_db()
    row = conn.execute("SELECT * FROM months WHERE id = ?", (month_id,)).fetchone()
    conn.close()
    return dict(row) if row else None


def get_month_by_sheet(salary_sheet_name):
    conn = get_db()
    row = conn.execute("SELECT * FROM months WHERE salary_sheet_name = ?", (salary_sheet_name,)).fetchone()
    conn.close()
    return dict(row) if row else None


def delete_month(month_id):
    """Deletes the month row (cascades to month_summary, compliance_status,
    email_log via ON DELETE CASCADE). Returns the excel_path that was on
    record, so the caller can also remove the stored upload file from disk.
    Does NOT touch anything in a configured archive folder — that's treated
    as a deliberate external copy, not app-managed storage."""
    conn = get_db()
    row = conn.execute("SELECT excel_path FROM months WHERE id = ?", (month_id,)).fetchone()
    if row is None:
        conn.close()
        return None
    excel_path = row["excel_path"]
    conn.execute("DELETE FROM months WHERE id = ?", (month_id,))
    conn.commit()
    conn.close()
    return excel_path


def upsert_month(salary_sheet_name, month_label, excel_path, uploaded_by):
    conn = get_db()
    existing = conn.execute("SELECT id FROM months WHERE salary_sheet_name = ?", (salary_sheet_name,)).fetchone()
    now = datetime.utcnow().isoformat()
    if existing:
        conn.execute(
            "UPDATE months SET month_label=?, excel_path=?, uploaded_by=?, uploaded_at=? WHERE id=?",
            (month_label, excel_path, uploaded_by, now, existing["id"]),
        )
        month_id = existing["id"]
    else:
        cur = conn.execute(
            "INSERT INTO months (salary_sheet_name, month_label, excel_path, uploaded_by, uploaded_at) "
            "VALUES (?, ?, ?, ?, ?)",
            (salary_sheet_name, month_label, excel_path, uploaded_by, now),
        )
        month_id = cur.lastrowid
        for item in COMPLIANCE_ITEMS:
            conn.execute(
                "INSERT INTO compliance_status (month_id, item, filed) VALUES (?, ?, 0)",
                (month_id, item),
            )
    conn.commit()
    conn.close()
    return month_id


def save_month_summary(month_id, summary):
    conn = get_db()
    conn.execute(
        """INSERT INTO month_summary
               (month_id, employee_count, total_ctc, total_deductions, total_net_pay,
                employee_pf, employer_pf, employee_esi, employer_esi, pt)
           VALUES (:month_id, :employee_count, :total_ctc, :total_deductions, :total_net_pay,
                   :employee_pf, :employer_pf, :employee_esi, :employer_esi, :pt)
           ON CONFLICT(month_id) DO UPDATE SET
               employee_count=excluded.employee_count, total_ctc=excluded.total_ctc,
               total_deductions=excluded.total_deductions, total_net_pay=excluded.total_net_pay,
               employee_pf=excluded.employee_pf, employer_pf=excluded.employer_pf,
               employee_esi=excluded.employee_esi, employer_esi=excluded.employer_esi, pt=excluded.pt""",
        {"month_id": month_id, **summary},
    )
    conn.commit()
    conn.close()


def list_months_with_summary():
    """Months joined with their saved summary, newest first — for the dashboard."""
    conn = get_db()
    rows = conn.execute(
        """SELECT m.*, s.employee_count, s.total_ctc, s.total_deductions, s.total_net_pay,
                  s.employee_pf, s.employer_pf, s.employee_esi, s.employer_esi, s.pt
           FROM months m LEFT JOIN month_summary s ON s.month_id = m.id
           ORDER BY m.uploaded_at DESC"""
    ).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def get_compliance_status(month_id):
    conn = get_db()
    rows = conn.execute("SELECT * FROM compliance_status WHERE month_id = ?", (month_id,)).fetchall()
    conn.close()
    by_item = {r["item"]: dict(r) for r in rows}
    for item in COMPLIANCE_ITEMS:
        by_item.setdefault(item, {"item": item, "filed": 0, "filed_by": None, "filed_at": None})
    return by_item


def set_compliance_status(month_id, item, filed, filed_by):
    conn = get_db()
    now = datetime.utcnow().isoformat() if filed else None
    conn.execute(
        """INSERT INTO compliance_status (month_id, item, filed, filed_by, filed_at)
           VALUES (?, ?, ?, ?, ?)
           ON CONFLICT(month_id, item) DO UPDATE SET filed=excluded.filed, filed_by=excluded.filed_by,
                                                      filed_at=excluded.filed_at""",
        (month_id, item, 1 if filed else 0, filed_by if filed else None, now),
    )
    conn.commit()
    conn.close()


def get_user_by_username(username):
    conn = get_db()
    row = conn.execute("SELECT * FROM users WHERE username = ?", (username,)).fetchone()
    conn.close()
    return dict(row) if row else None


def get_user_by_id(user_id):
    conn = get_db()
    row = conn.execute("SELECT * FROM users WHERE id = ?", (user_id,)).fetchone()
    conn.close()
    return dict(row) if row else None


def create_user(username, password):
    conn = get_db()
    conn.execute(
        "INSERT INTO users (username, password_hash, created_at) VALUES (?, ?, ?)",
        (username, generate_password_hash(password), datetime.utcnow().isoformat()),
    )
    conn.commit()
    conn.close()


def list_users():
    conn = get_db()
    rows = conn.execute("SELECT id, username, created_at FROM users ORDER BY created_at").fetchall()
    conn.close()
    return [dict(r) for r in rows]


def update_user_password(user_id, password):
    conn = get_db()
    conn.execute(
        "UPDATE users SET password_hash = ? WHERE id = ?",
        (generate_password_hash(password), user_id),
    )
    conn.commit()
    conn.close()


def create_password_reset(user_id, token, expires_at):
    conn = get_db()
    conn.execute(
        "INSERT INTO password_resets (user_id, token, created_at, expires_at, used) VALUES (?, ?, ?, ?, 0)",
        (user_id, token, datetime.utcnow().isoformat(), expires_at),
    )
    conn.commit()
    conn.close()


def get_valid_password_reset(token):
    """Returns the reset row if token exists, is unused, and hasn't expired — else None."""
    conn = get_db()
    row = conn.execute("SELECT * FROM password_resets WHERE token = ? AND used = 0", (token,)).fetchone()
    conn.close()
    if not row or row["expires_at"] < datetime.utcnow().isoformat():
        return None
    return dict(row)


def mark_password_reset_used(token):
    conn = get_db()
    conn.execute("UPDATE password_resets SET used = 1 WHERE token = ?", (token,))
    conn.commit()
    conn.close()
