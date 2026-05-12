"""
Fetch all posted vendor bills from Odoo for FY 2025-26 (Apr-2025 to Mar-2026)
and save to odoo_bills_FY2025-26.xlsx
"""
import sys, io, json
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

import xmlrpc.client, ssl
from calendar import monthrange
from collections import defaultdict
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

URL     = 'https://fracktal.odoo.com'
DB      = 'fracktal'
USER    = 'suresh@fracktal.in'
API_KEY = '73b3a2650cf37c597f82ca3d0e02b8bf18882303'

MONTHS = [
    ("Apr-2025", "2025-04"), ("May-2025", "2025-05"), ("Jun-2025", "2025-06"),
    ("Jul-2025", "2025-07"), ("Aug-2025", "2025-08"), ("Sep-2025", "2025-09"),
    ("Oct-2025", "2025-10"), ("Nov-2025", "2025-11"), ("Dec-2025", "2025-12"),
    ("Jan-2026", "2026-01"), ("Feb-2026", "2026-02"), ("Mar-2026", "2026-03"),
]

ctx     = ssl._create_unverified_context()
common  = xmlrpc.client.ServerProxy(f'{URL}/xmlrpc/2/common', context=ctx, allow_none=True)
models  = xmlrpc.client.ServerProxy(f'{URL}/xmlrpc/2/object', context=ctx, allow_none=True)
uid     = common.authenticate(DB, USER, API_KEY, {})
print(f"Connected: UID={uid}")

def exe(model, method, args, kwargs=None):
    return models.execute_kw(DB, uid, API_KEY, model, method, args, kwargs or {})

# ── Fetch all FY bills in one call ─────────────────────────────────────────
print("Fetching vendor bills Apr-2025 → Mar-2026 ...")
domain = [
    ["move_type", "in", ["in_invoice", "in_refund"]],
    ["invoice_date", ">=", "2025-04-01"],
    ["invoice_date", "<=", "2026-03-31"],
    ["state", "=", "posted"],
]
inv_fields = [
    "id", "name", "ref", "move_type", "invoice_date",
    "partner_id", "amount_untaxed", "amount_tax", "amount_total",
    "invoice_line_ids",
]
invoices = exe("account.move", "search_read", [domain],
               {"fields": inv_fields, "order": "invoice_date asc"})
print(f"  Found {len(invoices)} vendor bills/refunds")

# ── Partner map ────────────────────────────────────────────────────────────
partner_ids = list({inv["partner_id"][0] for inv in invoices if inv.get("partner_id")})
raw_partners = exe("res.partner", "read", [partner_ids], {"fields": ["id","name","vat"]})
partner_map = {p["id"]: p for p in raw_partners}

# ── Tax lines per invoice ─────────────────────────────────────────────────
move_ids = [inv["id"] for inv in invoices]

# Fetch in chunks of 200 to avoid timeout
def chunked(lst, n):
    for i in range(0, len(lst), n):
        yield lst[i:i+n]

raw_tax_lines = []
for chunk in chunked(move_ids, 200):
    tl = exe("account.move.line", "search_read",
             [[["move_id","in",chunk], ["tax_line_id","!=",False]]],
             {"fields": ["move_id","tax_line_id","name","balance","debit","credit"]})
    raw_tax_lines.extend(tl)
    print(f"  Tax lines fetched so far: {len(raw_tax_lines)}")

tax_by_move = defaultdict(list)
for tl in raw_tax_lines:
    tax_by_move[tl["move_id"][0]].append(tl)

# ── Build rows ──────────────────────────────────────────────────────────────
def get_month_label(date_str):
    if not date_str:
        return ""
    y, m, d = date_str.split("-")
    month_names = {
        "04":"Apr","05":"May","06":"Jun","07":"Jul","08":"Aug","09":"Sep",
        "10":"Oct","11":"Nov","12":"Dec","01":"Jan","02":"Feb","03":"Mar"
    }
    return f"{month_names.get(m,'?')}-{y}"

rows = []
for inv in invoices:
    pid     = inv["partner_id"][0] if inv.get("partner_id") else None
    partner = partner_map.get(pid, {})

    vendor_ref  = (inv.get("ref") or "").strip()
    odoo_name   = (inv.get("name") or "").strip()
    bill_no_raw = vendor_ref or odoo_name

    igst = cgst = sgst = 0.0
    for tl in tax_by_move.get(inv["id"], []):
        tax_name = ""
        if tl.get("name"):
            tax_name = str(tl["name"]).upper()
        elif isinstance(tl.get("tax_line_id"), (list,tuple)) and len(tl["tax_line_id"]) > 1:
            tax_name = str(tl["tax_line_id"][1]).upper()
        amt = abs(float(tl.get("balance") or 0))
        if "IGST" in tax_name:
            igst += amt
        elif "CGST" in tax_name:
            cgst += amt
        elif "SGST" in tax_name or "UTGST" in tax_name:
            sgst += amt

    taxable    = float(inv.get("amount_untaxed") or 0)
    amount_tax = float(inv.get("amount_tax")     or 0)
    amount_tot = float(inv.get("amount_total")   or 0)

    computed = round(igst + cgst + sgst, 2)
    if computed == 0 and amount_tax > 0:
        igst = amount_tax   # fallback

    is_refund = inv.get("move_type") == "in_refund"
    sign = -1 if is_refund else 1

    rows.append({
        "month":      get_month_label(inv.get("invoice_date","")),
        "odoo_ref":   odoo_name,
        "vendor_ref": vendor_ref,
        "inv_date":   inv.get("invoice_date",""),
        "partner":    (partner.get("name") or "").strip(),
        "gstin":      (partner.get("vat")  or "").strip().upper(),
        "doc_type":   "Vendor Refund" if is_refund else "Vendor Bill",
        "taxable":    round(taxable * sign, 2),
        "igst":       round(igst    * sign, 2),
        "cgst":       round(cgst    * sign, 2),
        "sgst":       round(sgst    * sign, 2),
        "total_gst":  round((igst+cgst+sgst) * sign, 2),
        "total":      round(amount_tot * sign, 2),
    })

print(f"Total rows built: {len(rows)}")

# ── Save to JSON for next step ──────────────────────────────────────────────
with open("odoo_bills_raw.json", "w", encoding="utf-8") as f:
    json.dump(rows, f, ensure_ascii=False, indent=2)
print("Saved: odoo_bills_raw.json")

# ── Quick month-on-month summary ────────────────────────────────────────────
month_labels = [m[0] for m in MONTHS]
summary = {m: {"count":0,"taxable":0,"igst":0,"cgst":0,"sgst":0,"total_gst":0}
           for m in month_labels}
for r in rows:
    m = r["month"]
    if m in summary:
        summary[m]["count"]     += 1
        summary[m]["taxable"]   += r["taxable"]
        summary[m]["igst"]      += r["igst"]
        summary[m]["cgst"]      += r["cgst"]
        summary[m]["sgst"]      += r["sgst"]
        summary[m]["total_gst"] += r["total_gst"]

print(f"\n{'Month':<12} {'Bills':>6} {'Taxable':>15} {'IGST':>12} {'CGST':>12} {'SGST':>12} {'Total GST':>14}")
print("-"*85)
gt = {"count":0,"taxable":0,"igst":0,"cgst":0,"sgst":0,"total_gst":0}
for m in month_labels:
    s = summary[m]
    print(f"{m:<12} {s['count']:>6} {s['taxable']:>15,.2f} {s['igst']:>12,.2f} {s['cgst']:>12,.2f} {s['sgst']:>12,.2f} {s['total_gst']:>14,.2f}")
    for k in gt: gt[k] += s[k]
print("="*85)
print(f"{'TOTAL FY':<12} {gt['count']:>6} {gt['taxable']:>15,.2f} {gt['igst']:>12,.2f} {gt['cgst']:>12,.2f} {gt['sgst']:>12,.2f} {gt['total_gst']:>14,.2f}")
