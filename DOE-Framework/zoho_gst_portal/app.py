"""
Zoho Billing → GST Portal (GSTR-1) Monthly Upload Tool
Run this file and open http://localhost:5000 in your browser.
"""

import os
import json
import logging
import calendar
from datetime import date
from pathlib import Path

import requests
from flask import Flask, render_template, request, redirect, url_for, flash, session, jsonify, Response

app = Flask(__name__)
app.secret_key = "zoho-gst-secret-2026"

SETTINGS_FILE = Path(__file__).parent / "settings.json"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)s  %(message)s",
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler(Path(__file__).parent / "app.log"),
    ]
)
log = logging.getLogger(__name__)


# ─── Settings helpers ────────────────────────────────────────────────────────

def load_settings():
    if SETTINGS_FILE.exists():
        with open(SETTINGS_FILE) as f:
            return json.load(f)
    return {}

def save_settings(data: dict):
    with open(SETTINGS_FILE, "w") as f:
        json.dump(data, f, indent=2)


# ─── Zoho helpers ────────────────────────────────────────────────────────────

def _refresh_zoho_token(cfg: dict) -> str:
    """Use refresh_token to get a new access_token. Saves and returns it."""
    accounts_server = cfg.get("zoho_accounts_server", "https://accounts.zoho.in")
    resp = requests.post(f"{accounts_server}/oauth/v2/token", data={
        "grant_type":    "refresh_token",
        "client_id":     cfg.get("zoho_client_id", ""),
        "client_secret": cfg.get("zoho_client_secret", ""),
        "refresh_token": cfg.get("zoho_refresh_token", ""),
    }, timeout=30)
    if resp.status_code == 200:
        tokens = resp.json()
        new_token = tokens.get("access_token", "")
        if new_token:
            cfg["zoho_token"] = new_token
            save_settings(cfg)
            log.info("Zoho token refreshed successfully.")
            return new_token
    log.error("Token refresh failed: %s", resp.text)
    return ""


def _zoho_headers(token: str, cfg: dict) -> dict:
    """Build Zoho API request headers including org ID if available."""
    headers = {"Authorization": f"Zoho-oauthtoken {token}"}
    org_id = cfg.get("zoho_org_id", "")
    if org_id:
        headers["X-com-zoho-billing-organizationid"] = org_id
    return headers


def fetch_invoices(zoho_token: str, date_start: str, date_end: str) -> list:
    """Fetch sales invoices from Zoho Billing for the given date range."""
    cfg = load_settings()
    api_domain = cfg.get("zoho_api_domain", "https://www.zohoapis.in")
    org_id = cfg.get("zoho_org_id", "")
    url = f"{api_domain}/billing/v1/invoices"
    log.info("Zoho API call -> %s  (org_id=%s)", url, org_id or "NOT SET - may cause 401!")
    headers = _zoho_headers(zoho_token, cfg)
    params = {
        "date_start": date_start,
        "date_end": date_end,
        # No status filter — fetch all invoices in the date range
        # (sent, paid, overdue, etc.) so we don't miss any for GSTR-1
    }
    resp = requests.get(url, headers=headers, params=params, timeout=30)
    log.info("Zoho response: HTTP %s from %s", resp.status_code, resp.url)

    # Auto-refresh token if expired (401) then retry once
    if resp.status_code == 401:
        log.info("Zoho 401 - attempting token refresh...")
        zoho_token = _refresh_zoho_token(cfg)
        if zoho_token:
            headers = _zoho_headers(zoho_token, cfg)
            resp = requests.get(url, headers=headers, params=params, timeout=30)
            log.info("Zoho retry response: HTTP %s", resp.status_code)

    if not resp.ok:
        log.error("Zoho error body: %s", resp.text[:500])
    resp.raise_for_status()
    data = resp.json()
    invoices = data.get("invoices", [])
    log.info("Zoho returned %d invoices. Keys in response: %s", len(invoices), list(data.keys()))
    if invoices:
        log.info("Sample invoice fields: %s", list(invoices[0].keys()))
        log.info("Sample invoice data: %s", str(invoices[0])[:600])
    else:
        log.info("Full Zoho response (truncated): %s", str(data)[:500])
    return invoices


def map_to_gstr1(invoices: list, cfg: dict, year: int, month: int) -> dict:
    """
    Generate GSTR-1 JSON in the exact NIC format accepted by the GST portal.
    Spec: https://developer.gst.gov.in/apiportal/taxpayer/returns
    """
    gstin      = cfg.get("gstin", "")
    state_code = gstin[:2] if gstin else "29"   # e.g. "29" = Karnataka
    fp         = f"{month:02d}{year}"            # e.g. "032026"

    grand_total = 0.0
    b2b   = {}   # keyed by customer GSTIN
    b2cs  = {}   # keyed by (tax_rate, sply_tp, pos)
    b2cl  = []   # B2C Large: inter-state, no GSTIN, > 2.5 lakh
    nil_b2c  = 0.0
    expt_b2c = 0.0

    for inv in invoices:
        total   = float(inv.get("total", 0))
        tax     = float(inv.get("tax_total", 0) or inv.get("tax_amount", 0) or 0)
        taxable = round(total - tax, 2)
        grand_total += total

        cust_gstin = (inv.get("gst_no") or inv.get("customer_gstin")
                      or inv.get("billing_gst_no") or "").strip()
        inv_no   = inv.get("invoice_number") or inv.get("number", "")
        inv_date = inv.get("invoice_date") or inv.get("date", "")
        # Convert YYYY-MM-DD → DD-MM-YYYY for GST portal
        if inv_date and "-" in str(inv_date):
            parts = str(inv_date).split("-")
            if len(parts) == 3 and len(parts[0]) == 4:
                inv_date = f"{parts[2]}-{parts[1]}-{parts[0]}"

        # ── B2B: customer has a valid GSTIN ──────────────────────────────────
        if cust_gstin and len(cust_gstin) == 15 and cust_gstin != "URP":
            cust_state = cust_gstin[:2]
            sply_tp = "INTRA" if cust_state == state_code else "INTER"
            igst = round(tax, 2) if sply_tp == "INTER" else 0.0
            cgst = round(tax / 2, 2) if sply_tp == "INTRA" else 0.0
            sgst = round(tax / 2, 2) if sply_tp == "INTRA" else 0.0
            if cust_gstin not in b2b:
                b2b[cust_gstin] = {"ctin": cust_gstin, "inv": []}
            b2b[cust_gstin]["inv"].append({
                "inum": inv_no, "idt": inv_date, "val": round(total, 2),
                "pos": cust_state, "rchrg": "N", "inv_typ": "R",
                "itms": [{"num": 1, "itm_det": {
                    "rt": round(tax / taxable * 100, 2) if taxable > 0 else 0,
                    "txval": taxable, "iamt": igst, "camt": cgst,
                    "samt": sgst, "csamt": 0.0
                }}]
            })
            continue

        # ── B2C: no GSTIN (individual / URP) ─────────────────────────────────
        if tax == 0:
            # Nil-rated or exempt supply
            nil_b2c += taxable
            continue

        rt = round(tax / taxable * 100) if taxable > 0 else 0
        # Inter-state B2C > 2.5 lakh → B2CL (individual invoice required)
        if total > 250000:
            igst = round(tax, 2)
            b2cl.append({
                "pos": "96",   # place of supply — update if known
                "inv": [{"inum": inv_no, "idt": inv_date, "val": round(total, 2),
                          "itms": [{"num": 1, "itm_det": {
                              "rt": rt, "txval": taxable,
                              "iamt": igst, "csamt": 0.0
                          }}]}]
            })
            continue

        # Intra-state B2C small — aggregate by (rate, state)
        key = (rt, state_code)
        if key not in b2cs:
            b2cs[key] = {"txval": 0.0, "iamt": 0.0, "camt": 0.0, "samt": 0.0}
        b2cs[key]["txval"] += taxable
        b2cs[key]["camt"]  += tax / 2
        b2cs[key]["samt"]  += tax / 2

    # Build b2cs list
    b2cs_list = [
        {
            "sply_tp": "INTRA", "typ": "OE", "pos": pos,
            "rt": rt,
            "txval": round(d["txval"], 2),
            "iamt": 0.0,
            "camt": round(d["camt"], 2),
            "samt": round(d["samt"], 2),
            "csamt": 0.0,
        }
        for (rt, pos), d in b2cs.items()
    ]

    # Nil/exempt section
    exemp = {}
    if nil_b2c > 0 or expt_b2c > 0:
        exemp = {"nil": [
            {"sply_tp": "INTRB2B",  "nil_amt": 0.0, "expt_amt": 0.0, "ngsup_amt": 0.0},
            {"sply_tp": "INTRB2C",  "nil_amt": round(nil_b2c, 2),
             "expt_amt": round(expt_b2c, 2), "ngsup_amt": 0.0},
            {"sply_tp": "INTERB2B", "nil_amt": 0.0, "expt_amt": 0.0, "ngsup_amt": 0.0},
            {"sply_tp": "INTERB2C", "nil_amt": 0.0, "expt_amt": 0.0, "ngsup_amt": 0.0},
        ]}

    return {
        "gstin":   gstin,
        "fp":      fp,
        "gt":      round(grand_total, 2),
        "cur_gt":  round(grand_total, 2),
        "b2b":     list(b2b.values()),
        "b2cl":    b2cl,
        "b2cs":    b2cs_list,
        "cdnr":    [],
        "cdnur":   [],
        "exp":     [],
        "at":      [],
        "atadj":   [],
        "exemp":   exemp,
        "hsnsum":  {},
    }


# ─── GST Portal helpers ───────────────────────────────────────────────────────

def authenticate_gst(cfg: dict) -> str:
    """Authenticate with GST Portal and return access token."""
    url = "https://api.gst.gov.in/authenticate"
    payload = {
        "username":      cfg["gst_user_id"],
        "password":      cfg["gst_password"],
    }
    if cfg.get("gst_client_id"):
        payload["client_id"]     = cfg["gst_client_id"]
        payload["client_secret"] = cfg["gst_client_secret"]

    resp = requests.post(url, json=payload, timeout=30)
    resp.raise_for_status()
    data = resp.json()
    return data.get("access_token") or data.get("token") or ""


def push_gstr1(payload: dict, token: str) -> dict:
    """Upload GSTR-1 data to GST Portal."""
    url = "https://api.gst.gov.in/gstr1/upload"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type":  "application/json",
    }
    resp = requests.post(url, headers=headers, json=payload, timeout=60)
    resp.raise_for_status()
    return resp.json()


# ─── Routes ──────────────────────────────────────────────────────────────────

@app.route("/")
def dashboard():
    cfg = load_settings()
    today = date.today()
    # Default: previous month
    month = int(request.args.get("month", today.month - 1 or 12))
    year  = int(request.args.get("year",  today.year if today.month > 1 else today.year - 1))

    history = session.get("history", [])
    return render_template(
        "dashboard.html",
        cfg=cfg,
        month=month,
        year=year,
        months=list(enumerate(calendar.month_name))[1:],
        years=list(range(2024, today.year + 2)),
        history=history,
        has_settings=bool(cfg.get("gst_user_id")),
        zoho_missing=not bool(cfg.get("zoho_token")),
    )


@app.route("/download", methods=["POST"])
def download():
    """Fetch invoices and return a GSTR-1 JSON file for download."""
    cfg = load_settings()
    month = int(request.form["month"])
    year  = int(request.form["year"])

    last_day   = calendar.monthrange(year, month)[1]
    date_start = f"{year}-{month:02d}-01"
    date_end   = f"{year}-{month:02d}-{last_day:02d}"

    try:
        invoices = fetch_invoices(cfg["zoho_token"], date_start, date_end)
        if not invoices:
            flash(f"No invoices found for {calendar.month_name[month]} {year}.", "warning")
            return redirect(url_for("dashboard", month=month, year=year), 303)

        payload  = map_to_gstr1(invoices, cfg, year, month)
        filename = f"GSTR1_{cfg.get('gstin','GSTIN')}_{month:02d}{year}.json"
        log.info("GSTR-1 generated: %d B2B, %d B2CS, %d B2CL, nil=%.2f",
                 len(payload.get("b2b", [])), len(payload.get("b2cs", [])),
                 len(payload.get("b2cl", [])),
                 payload.get("exemp", {}).get("nil", [{}]*2)[1].get("nil_amt", 0)
                 if payload.get("exemp") else 0)
        return Response(
            json.dumps(payload, indent=2),
            mimetype="application/json",
            headers={"Content-Disposition": f"attachment; filename={filename}"}
        )
    except Exception as e:
        log.error("Download error: %s", e)
        flash(f"Error: {e}", "danger")
        return redirect(url_for("dashboard", month=month, year=year), 303)


@app.route("/push", methods=["POST"])
def push():
    cfg = load_settings()
    month = int(request.form["month"])
    year  = int(request.form["year"])

    last_day   = calendar.monthrange(year, month)[1]
    date_start = f"{year}-{month:02d}-01"
    date_end   = f"{year}-{month:02d}-{last_day:02d}"

    try:
        log.info("Fetching invoices %s → %s", date_start, date_end)
        invoices = fetch_invoices(cfg["zoho_token"], date_start, date_end)

        if not invoices:
            flash(f"No invoices found for {calendar.month_name[month]} {year}.", "warning")
            return redirect(url_for("dashboard", month=month, year=year), 303)

        payload      = map_to_gstr1(invoices, cfg, year, month)
        gst_token    = authenticate_gst(cfg)
        result       = push_gstr1(payload, gst_token)

        summary = {
            "month":    f"{calendar.month_name[month]} {year}",
            "count":    len(invoices),
            "status":   "Success",
            "message":  str(result),
            "invoices": payload["invoices"],
        }
        history = session.get("history", [])
        history.insert(0, summary)
        session["history"] = history[:10]   # keep last 10 runs

        log.info("Pushed %d invoices. Result: %s", len(invoices), result)
        flash(f"Successfully pushed {len(invoices)} invoice(s) to GST Portal!", "success")

    except requests.HTTPError as e:
        status = e.response.status_code if e.response else "?"
        body   = e.response.text[:300] if e.response else str(e)
        log.error("HTTP %s: %s", status, body)
        flash(f"API error (HTTP {status}): {body}", "danger")

    except Exception as e:
        log.error("Unexpected error: %s", e)
        flash(f"Error: {e}", "danger")

    return redirect(url_for("dashboard", month=month, year=year), 303)


@app.route("/settings", methods=["GET", "POST"])
def settings():
    cfg = load_settings()
    if request.method == "POST":
        cfg.update({
            "zoho_token":        request.form["zoho_token"].strip(),
            "zoho_client_id":    request.form.get("zoho_client_id", "").strip(),
            "zoho_client_secret":request.form.get("zoho_client_secret", "").strip(),
            "zoho_org_id":       request.form.get("zoho_org_id", "").strip(),
            "gst_user_id":       request.form["gst_user_id"].strip(),
            "gst_password":      request.form["gst_password"].strip(),
            "gst_client_id":     request.form.get("gst_client_id", "").strip(),
            "gst_client_secret": request.form.get("gst_client_secret", "").strip(),
            "gstin":             request.form.get("gstin", "").strip(),
            "company_name":      request.form.get("company_name", "").strip(),
        })
        save_settings(cfg)
        flash("Settings saved successfully!", "success")
        return redirect(url_for("dashboard"))
    return render_template("settings.html", cfg=cfg)


@app.route("/zoho/fetch-org")
def zoho_fetch_org():
    """Auto-detect the Zoho organization ID and save it."""
    cfg = load_settings()
    token = cfg.get("zoho_token", "")
    if not token:
        flash("Connect Zoho first before fetching organization.", "warning")
        return redirect(url_for("settings"))

    api_domain = cfg.get("zoho_api_domain", "https://www.zohoapis.in")
    url = f"{api_domain}/billing/v1/organizations"
    try:
        headers = _zoho_headers(token, cfg)
        resp = requests.get(url, headers=headers, timeout=30)
        if resp.status_code == 401:
            token = _refresh_zoho_token(cfg)
            if token:
                headers = _zoho_headers(token, cfg)
                resp = requests.get(url, headers=headers, timeout=30)
        log.info("Fetch-org response: HTTP %s -- %s", resp.status_code, resp.text[:300])
        resp.raise_for_status()
        data = resp.json()
        orgs = data.get("organizations", [])
        if orgs:
            org = orgs[0]
            org_id   = str(org.get("organization_id", ""))
            org_name = org.get("name", "")
            cfg["zoho_org_id"] = org_id
            save_settings(cfg)
            flash(f"Organization detected: {org_name} (ID: {org_id}) - saved!", "success")
        else:
            flash(f"No organizations found. Zoho response: {data}", "warning")
    except requests.HTTPError as e:
        body = e.response.text[:300] if e.response else str(e)
        log.error("Fetch-org HTTP error: %s", body)
        flash(f"Zoho API error: {body}", "danger")
    except Exception as e:
        log.error("Fetch-org error: %s", e)
        flash(f"Could not fetch organization: {e}", "danger")
    return redirect(url_for("settings"))


@app.route("/zoho/authorize")
def zoho_authorize():
    cfg = load_settings()
    client_id    = cfg.get("zoho_client_id", "")
    redirect_uri = "http://localhost:5000/callback"
    if not client_id:
        flash("Enter your Zoho Client ID in Settings first.", "warning")
        return redirect(url_for("settings"))
    auth_url = (
        "https://accounts.zoho.in/oauth/v2/auth"
        f"?response_type=code&client_id={client_id}"
        "&scope=ZohoSubscriptions.fullaccess.all"
        f"&redirect_uri={redirect_uri}&access_type=offline"
    )
    return redirect(auth_url)


@app.route("/callback")
def zoho_callback():
    code  = request.args.get("code", "")
    error = request.args.get("error", "")
    if error or not code:
        flash(f"Zoho authorization failed: {error or 'no code'}", "danger")
        return redirect(url_for("settings"))

    # Zoho passes the correct accounts server in the callback (India = accounts.zoho.in)
    accounts_server = request.args.get("accounts-server", "https://accounts.zoho.in")

    cfg = load_settings()
    try:
        resp = requests.post(f"{accounts_server}/oauth/v2/token", data={
            "grant_type":    "authorization_code",
            "client_id":     cfg.get("zoho_client_id", ""),
            "client_secret": cfg.get("zoho_client_secret", ""),
            "redirect_uri":  "http://localhost:5000/callback",
            "code":          code,
        }, timeout=30)
        resp.raise_for_status()
        tokens = resp.json()
        log.info("Zoho token response: %s", tokens)

        if "error" in tokens:
            flash(f"Zoho error: {tokens['error']} — {tokens.get('error_description', '')}", "danger")
            return redirect(url_for("dashboard"))

        access_token  = tokens.get("access_token", "")
        refresh_token = tokens.get("refresh_token", "")

        if not access_token:
            flash(f"Zoho did not return a token. Full response: {tokens}", "danger")
            return redirect(url_for("dashboard"))

        cfg["zoho_token"]        = access_token
        cfg["zoho_refresh_token"]= refresh_token
        cfg["zoho_api_domain"]   = tokens.get("api_domain", "https://www.zohoapis.in")
        cfg["zoho_accounts_server"] = accounts_server
        save_settings(cfg)
        flash("Zoho connected successfully! Your token has been saved.", "success")
    except Exception as e:
        log.error("Token exchange error: %s", e)
        flash(f"Token exchange failed: {e}", "danger")

    return redirect(url_for("dashboard"))


if __name__ == "__main__":
    print("\n" + "="*55)
    print("  Zoho Billing -> GST Portal  |  Monthly Upload Tool")
    print("="*55)
    print("  Open your browser and go to:  http://localhost:5000")
    print("  Press Ctrl+C to stop the app.")
    print("="*55 + "\n")
    app.run(debug=False, port=5000)
