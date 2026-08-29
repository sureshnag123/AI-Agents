"""Flask-Login setup: session-based auth, no default shipped credentials.

The first account is seeded by db.init_db() from INITIAL_ADMIN_USER /
INITIAL_ADMIN_PASSWORD env vars if the users table is empty. Additional
users are added from the Settings tab once logged in.
"""

import secrets
from datetime import datetime, timedelta

from flask import Blueprint, render_template, redirect, url_for, request, flash
from flask_login import LoginManager, UserMixin, login_user, logout_user, login_required, current_user
from werkzeug.security import check_password_hash

import db
import emailer

login_manager = LoginManager()
login_manager.login_view = "auth.login"

auth_bp = Blueprint("auth", __name__)

RESET_TOKEN_TTL_MINUTES = 60


class User(UserMixin):
    def __init__(self, row):
        self.id = str(row["id"])
        self.username = row["username"]


@login_manager.user_loader
def load_user(user_id):
    row = db.get_user_by_id(int(user_id))
    return User(row) if row else None


@auth_bp.route("/login", methods=["GET", "POST"])
def login():
    if current_user.is_authenticated:
        return redirect(url_for("dashboard"))
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        row = db.get_user_by_username(username)
        if row and check_password_hash(row["password_hash"], password):
            login_user(User(row))
            return redirect(url_for("dashboard"))
        flash("Invalid username or password.", "error")
    return render_template("login.html")


@auth_bp.route("/forgot-password", methods=["GET", "POST"])
def forgot_password():
    if current_user.is_authenticated:
        return redirect(url_for("dashboard"))
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        row = db.get_user_by_username(username)
        if row:
            token = secrets.token_urlsafe(32)
            expires_at = (datetime.utcnow() + timedelta(minutes=RESET_TOKEN_TTL_MINUTES)).isoformat()
            db.create_password_reset(row["id"], token, expires_at)
            reset_link = url_for("auth.reset_password", token=token, _external=True)

            smtp_settings = db.get_email_settings()
            if smtp_settings and smtp_settings.get("smtp_username"):
                try:
                    emailer.send_password_reset_email(smtp_settings, username, reset_link, RESET_TOKEN_TTL_MINUTES)
                except Exception as e:
                    # Don't leak delivery failures to the requester (avoids confirming
                    # the account exists); fall back to the server console.
                    print(f"[forgot-password] email send failed for {username}: {e}")
                    print(f"[forgot-password] reset link for {username}: {reset_link}")
            else:
                # No SMTP configured yet — this is a small internal tool, so the
                # person running the server locally can hand the link over directly.
                print(f"[forgot-password] SMTP not configured. Reset link for {username}: {reset_link}")

        # Same message whether or not the account exists, so this can't be used to
        # discover valid usernames.
        flash("If that account exists, a password reset link has been sent. "
              "(If SMTP email isn't set up yet, ask whoever is running the server "
              "to check its console output for the link.)", "success")
        return redirect(url_for("auth.login"))
    return render_template("forgot_password.html")


@auth_bp.route("/reset-password/<token>", methods=["GET", "POST"])
def reset_password(token):
    if current_user.is_authenticated:
        return redirect(url_for("dashboard"))
    reset = db.get_valid_password_reset(token)
    if not reset:
        flash("This reset link is invalid or has expired. Request a new one.", "error")
        return redirect(url_for("auth.forgot_password"))

    if request.method == "POST":
        password = request.form.get("password", "")
        confirm = request.form.get("confirm_password", "")
        if len(password) < 8:
            flash("Password must be at least 8 characters.", "error")
        elif password != confirm:
            flash("Passwords do not match.", "error")
        else:
            db.update_user_password(reset["user_id"], password)
            db.mark_password_reset_used(token)
            flash("Password updated. Please sign in.", "success")
            return redirect(url_for("auth.login"))

    return render_template("reset_password.html", token=token)


@auth_bp.route("/logout")
@login_required
def logout():
    logout_user()
    return redirect(url_for("auth.login"))
