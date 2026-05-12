"""
Odoo ERP API Connector for GST Compliance Automation
=====================================================
Connects to Odoo via XML-RPC to fetch sales/purchase invoices.

SETUP:
  1. Set your Odoo credentials in the config dict below
  2. Ensure XML-RPC is enabled on your Odoo instance
  3. Run: python odoo_connector.py

This module exports two functions:
  - fetch_sales_invoices(config, date_from, date_to) → DataFrame
  - fetch_purchase_invoices(config, date_from, date_to) → DataFrame
"""

import xmlrpc.client
import pandas as pd
from datetime import datetime

# ── CONFIGURATION ──────────────────────────────────────────────
ODOO_CONFIG = {
    "url":      "https://your-odoo-instance.com",   # ← Replace
    "db":       "your_database_name",                # ← Replace
    "username": "admin@yourcompany.com",             # ← Replace
    "password": "your_api_key_or_password",           # ← Replace
}

STATE_CODE_MAP = {
    "Andhra Pradesh": "37", "Arunachal Pradesh": "12", "Assam": "18",
    "Bihar": "10", "Chhattisgarh": "22", "Goa": "30", "Gujarat": "24",
    "Haryana": "06", "Himachal Pradesh": "02", "Jharkhand": "20",
    "Karnataka": "29", "Kerala": "32", "Madhya Pradesh": "23",
    "Maharashtra": "27", "Manipur": "14", "Meghalaya": "17",
    "Mizoram": "15", "Nagaland": "13", "Odisha": "21", "Punjab": "03",
    "Rajasthan": "08", "Sikkim": "11", "Tamil Nadu": "33",
    "Telangana": "36", "Tripura": "16", "Uttar Pradesh": "09",
    "Uttarakhand": "05", "West Bengal": "19", "Delhi": "07",
    "Jammu and Kashmir": "01", "Ladakh": "02", "Puducherry": "34",
    "Chandigarh": "04", "Dadra and Nagar Haveli": "26", "Daman and Diu": "25",
    "Lakshadweep": "31", "Andaman and Nicobar Islands": "35",
}


# ── AUTHENTICATION ─────────────────────────────────────────────
def authenticate(config):
    common = xmlrpc.client.ServerProxy(f"{config['url']}/xmlrpc/2/common")
    uid = common.authenticate(config["db"], config["username"], config["password"], {})
    if not uid:
        raise ConnectionError("Odoo authentication failed. Check credentials.")
    models = xmlrpc.client.ServerProxy(f"{config['url']}/xmlrpc/2/object")
    return uid, models


# ── SALES INVOICES ─────────────────────────────────────────────
def fetch_sales_invoices(config, date_from, date_to):
    """
    Fetch customer invoices and credit notes from Odoo.
    Returns a DataFrame matching the GST Sales Register format.
    """
    uid, models = authenticate(config)

    domain = [
        ("move_type", "in", ["out_invoice", "out_refund"]),
        ("state", "=", "posted"),
        ("invoice_date", ">=", date_from),
        ("invoice_date", "<=", date_to),
    ]

    fields = [
        "name", "invoice_date", "partner_id", "amount_total",
        "amount_untaxed", "move_type", "l10n_in_gstin",
        "fiscal_position_id", "invoice_line_ids",
    ]

    invoices = models.execute_kw(
        config["db"], uid, config["password"],
        "account.move", "search_read", [domain], {"fields": fields}
    )

    rows = []
    for inv in invoices:
        partner_id = inv["partner_id"][0] if inv["partner_id"] else None
        partner_data = {}
        if partner_id:
            partner_data = models.execute_kw(
                config["db"], uid, config["password"],
                "res.partner", "read", [partner_id],
                {"fields": ["name", "vat", "state_id", "l10n_in_gst_treatment"]}
            )
            if isinstance(partner_data, list):
                partner_data = partner_data[0] if partner_data else {}

        gstin = partner_data.get("vat", "") or ""
        state = partner_data.get("state_id", [None, ""])[1] if partner_data.get("state_id") else ""
        pos = STATE_CODE_MAP.get(state, "")

        line_ids = inv.get("invoice_line_ids", [])
        if line_ids:
            lines = models.execute_kw(
                config["db"], uid, config["password"],
                "account.move.line", "read", [line_ids],
                {"fields": ["product_id", "price_subtotal", "tax_ids",
                            "l10n_in_hsn_code", "quantity"]}
            )
        else:
            lines = []

        taxable = sum(l.get("price_subtotal", 0) for l in lines if l.get("price_subtotal", 0) > 0)
        igst = cgst = sgst = 0.0

        for line in lines:
            tax_ids = line.get("tax_ids", [])
            if tax_ids:
                taxes = models.execute_kw(
                    config["db"], uid, config["password"],
                    "account.tax", "read", [tax_ids],
                    {"fields": ["name", "amount", "tax_group_id"]}
                )
                for t in taxes:
                    tname = (t.get("name") or "").upper()
                    amt = line["price_subtotal"] * t["amount"] / 100
                    if "IGST" in tname:
                        igst += amt
                    elif "CGST" in tname:
                        cgst += amt
                    elif "SGST" in tname:
                        sgst += amt

        hsn = ""
        for line in lines:
            h = line.get("l10n_in_hsn_code", "")
            if h:
                hsn = str(h)
                break

        doc_type = "Credit Note" if inv["move_type"] == "out_refund" else "Invoice"

        rows.append({
            "Invoice Number": inv["name"],
            "Invoice Date": inv["invoice_date"],
            "Customer GSTIN": gstin,
            "Customer Name": partner_data.get("name", ""),
            "Invoice Value": round(inv["amount_total"], 2),
            "Taxable Value": round(taxable, 2),
            "IGST": round(igst, 2),
            "CGST": round(cgst, 2),
            "SGST": round(sgst, 2),
            "Place of Supply": pos,
            "HSN Code": hsn,
            "Document Type": doc_type,
        })

    df = pd.DataFrame(rows)
    df["Invoice Date"] = pd.to_datetime(df["Invoice Date"]).dt.strftime("%d-%m-%Y")
    return df


# ── PURCHASE INVOICES ──────────────────────────────────────────
def fetch_purchase_invoices(config, date_from, date_to):
    """
    Fetch vendor bills and refunds from Odoo.
    Returns a DataFrame matching the GST Purchase Register format.
    """
    uid, models = authenticate(config)

    domain = [
        ("move_type", "in", ["in_invoice", "in_refund"]),
        ("state", "=", "posted"),
        ("invoice_date", ">=", date_from),
        ("invoice_date", "<=", date_to),
    ]

    fields = [
        "name", "ref", "invoice_date", "partner_id",
        "amount_total", "amount_untaxed", "move_type",
        "invoice_line_ids",
    ]

    bills = models.execute_kw(
        config["db"], uid, config["password"],
        "account.move", "search_read", [domain], {"fields": fields}
    )

    rows = []
    for bill in bills:
        partner_id = bill["partner_id"][0] if bill["partner_id"] else None
        partner_data = {}
        if partner_id:
            partner_data = models.execute_kw(
                config["db"], uid, config["password"],
                "res.partner", "read", [partner_id],
                {"fields": ["name", "vat", "state_id"]}
            )
            if isinstance(partner_data, list):
                partner_data = partner_data[0] if partner_data else {}

        gstin = partner_data.get("vat", "") or ""

        line_ids = bill.get("invoice_line_ids", [])
        if line_ids:
            lines = models.execute_kw(
                config["db"], uid, config["password"],
                "account.move.line", "read", [line_ids],
                {"fields": ["price_subtotal", "tax_ids", "l10n_in_hsn_code"]}
            )
        else:
            lines = []

        taxable = sum(l.get("price_subtotal", 0) for l in lines if l.get("price_subtotal", 0) > 0)
        igst = cgst = sgst = 0.0

        for line in lines:
            tax_ids = line.get("tax_ids", [])
            if tax_ids:
                taxes = models.execute_kw(
                    config["db"], uid, config["password"],
                    "account.tax", "read", [tax_ids],
                    {"fields": ["name", "amount"]}
                )
                for t in taxes:
                    tname = (t.get("name") or "").upper()
                    amt = line["price_subtotal"] * t["amount"] / 100
                    if "IGST" in tname:
                        igst += amt
                    elif "CGST" in tname:
                        cgst += amt
                    elif "SGST" in tname:
                        sgst += amt

        hsn = ""
        for line in lines:
            h = line.get("l10n_in_hsn_code", "")
            if h:
                hsn = str(h)
                break

        inv_num = bill.get("ref") or bill["name"]

        rows.append({
            "Vendor GSTIN": gstin,
            "Vendor Name": partner_data.get("name", ""),
            "Invoice Number": inv_num,
            "Invoice Date": bill["invoice_date"],
            "Taxable Value": round(taxable, 2),
            "IGST": round(igst, 2),
            "CGST": round(cgst, 2),
            "SGST": round(sgst, 2),
            "Invoice Value": round(bill["amount_total"], 2),
            "HSN Code": hsn,
            "Reverse Charge": "N",
            "Blocked ITC": "N",
        })

    df = pd.DataFrame(rows)
    df["Invoice Date"] = pd.to_datetime(df["Invoice Date"]).dt.strftime("%d-%m-%Y")
    return df


# ── GSTR-2B PARSER ─────────────────────────────────────────────
def parse_gstr2b_json(filepath):
    """Parse GSTR-2B JSON downloaded from the GST portal."""
    import json
    with open(filepath, "r") as f:
        data = json.load(f)

    rows = []
    docdata = data.get("data", {}).get("docdata", {})

    # B2B section
    for supplier in docdata.get("b2b", []):
        gstin = supplier.get("ctin", "")
        name = supplier.get("trdnm", "")
        filing = supplier.get("supprd", "Filed")
        for inv in supplier.get("inv", []):
            taxable = igst = cgst = sgst = 0.0
            for item in inv.get("items", []):
                det = item.get("itm_det", {})
                taxable += det.get("txval", 0)
                igst += det.get("iamt", 0)
                cgst += det.get("camt", 0)
                sgst += det.get("samt", 0)
            rows.append({
                "Supplier GSTIN": gstin,
                "Supplier Name": name,
                "Invoice Number": inv.get("inum", ""),
                "Invoice Date": inv.get("idt", ""),
                "Taxable Value": round(taxable, 2),
                "IGST": round(igst, 2),
                "CGST": round(cgst, 2),
                "SGST": round(sgst, 2),
                "Invoice Value": round(inv.get("val", taxable + igst + cgst + sgst), 2),
                "ITC Available": "Yes" if inv.get("itcavl", "Y") == "Y" else "No",
                "Reason": inv.get("rsn", ""),
                "Filing Status": filing,
            })

    return pd.DataFrame(rows)


def parse_gstr2b_excel(filepath):
    """Parse GSTR-2B Excel downloaded from the GST portal."""
    df = pd.read_excel(filepath)
    col_map = {}
    for col in df.columns:
        cl = col.lower().strip()
        if "gstin" in cl and "supplier" in cl:
            col_map["Supplier GSTIN"] = col
        elif "trade" in cl or ("supplier" in cl and "name" in cl):
            col_map["Supplier Name"] = col
        elif "invoice" in cl and "number" in cl:
            col_map["Invoice Number"] = col
        elif "invoice" in cl and "date" in cl:
            col_map["Invoice Date"] = col
        elif "taxable" in cl:
            col_map["Taxable Value"] = col
        elif "igst" in cl:
            col_map["IGST"] = col
        elif "cgst" in cl:
            col_map["CGST"] = col
        elif "sgst" in cl:
            col_map["SGST"] = col
        elif "invoice" in cl and "value" in cl:
            col_map["Invoice Value"] = col
    df = df.rename(columns={v: k for k, v in col_map.items()})
    return df


# ── CLI ENTRY POINT ────────────────────────────────────────────
if __name__ == "__main__":
    print("Odoo GST Connector — Test Mode")
    print("=" * 50)
    print("To use with live Odoo, update ODOO_CONFIG above and run:")
    print("  sales_df = fetch_sales_invoices(ODOO_CONFIG, '2025-01-01', '2025-03-31')")
    print("  purch_df = fetch_purchase_invoices(ODOO_CONFIG, '2025-01-01', '2025-03-31')")
    print()
    print("To parse GSTR-2B from GST portal:")
    print("  gstr2b_df = parse_gstr2b_json('path/to/gstr2b.json')")
    print("  gstr2b_df = parse_gstr2b_excel('path/to/gstr2b.xlsx')")
