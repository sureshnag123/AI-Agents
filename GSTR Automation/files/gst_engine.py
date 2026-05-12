"""
GST Compliance Engine — Core Processing Module
================================================
Handles: GSTR-1 preparation, GSTR-2B reconciliation, ITC analysis,
         alert generation, and monthly automation scheduling.
"""

import pandas as pd
import numpy as np
import re
import json
from datetime import datetime, timedelta
from collections import defaultdict
import random

# ── CONSTANTS ──────────────────────────────────────────────────

GSTIN_PATTERN = re.compile(
    r"^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[0-9A-Z]{1}[Z]{1}[0-9A-Z]{1}$"
)

STATES = {
    "01": "Jammu & Kashmir", "02": "Himachal Pradesh", "03": "Punjab",
    "04": "Chandigarh", "05": "Uttarakhand", "06": "Haryana",
    "07": "Delhi", "08": "Rajasthan", "09": "Uttar Pradesh",
    "10": "Bihar", "11": "Sikkim", "12": "Arunachal Pradesh",
    "13": "Nagaland", "14": "Manipur", "15": "Mizoram",
    "16": "Tripura", "17": "Meghalaya", "18": "Assam",
    "19": "West Bengal", "20": "Jharkhand", "21": "Odisha",
    "22": "Chhattisgarh", "23": "Madhya Pradesh", "24": "Gujarat",
    "26": "Dadra & Nagar Haveli", "27": "Maharashtra", "29": "Karnataka",
    "30": "Goa", "31": "Lakshadweep", "32": "Kerala",
    "33": "Tamil Nadu", "34": "Puducherry", "35": "Andaman & Nicobar",
    "36": "Telangana", "37": "Andhra Pradesh", "38": "Ladakh",
    "96": "Foreign Country", "97": "Other Territory",
}

OWN_STATE = "29"  # Karnataka — change to your state

BLOCKED_ITC_SECTIONS = [
    "Motor vehicles (Sec 17(5)(a))",
    "Food & beverages (Sec 17(5)(b))",
    "Club membership (Sec 17(5)(b))",
    "Personal consumption (Sec 17(5)(g))",
    "Free samples (Sec 17(5)(h))",
]


# ── VALIDATION UTILITIES ──────────────────────────────────────

def validate_gstin(gstin):
    if not gstin or pd.isna(gstin) or str(gstin).strip() == "":
        return "Missing"
    gstin = str(gstin).strip().upper()
    if GSTIN_PATTERN.match(gstin):
        state_code = gstin[:2]
        if state_code in STATES:
            return "Valid"
        return "Invalid (bad state code)"
    if len(gstin) != 15:
        return f"Invalid (length={len(gstin)})"
    return "Invalid (format error)"

def detect_ocr_errors(gstin):
    """Flag common OCR misreads in GSTIN."""
    if not gstin or pd.isna(gstin):
        return []
    issues = []
    s = str(gstin).strip()
    ocr_swaps = {"0": "O", "O": "0", "1": "I", "I": "1", "5": "S", "S": "5", "8": "B", "B": "8"}
    if len(s) == 15:
        for i, ch in enumerate(s):
            if i < 2 and ch.isalpha():
                issues.append(f"Pos {i+1}: '{ch}' should be digit (OCR: {ocr_swaps.get(ch, '?')})")
            if 2 <= i <= 6 and ch.isdigit():
                issues.append(f"Pos {i+1}: '{ch}' should be letter (OCR: {ocr_swaps.get(ch, '?')})")
    return issues

def validate_tax_calc(row, tolerance=1.0):
    taxable = row.get("Taxable Value", 0)
    rate = row.get("Tax Rate", 0)
    igst = row.get("IGST", 0)
    cgst = row.get("CGST", 0)
    sgst = row.get("SGST", 0)
    expected = round(taxable * rate / 100, 2)
    actual = round(igst + cgst + sgst, 2)
    if abs(expected - actual) > tolerance:
        return False, expected, actual
    return True, expected, actual


# ── SAMPLE DATA GENERATION ────────────────────────────────────

def generate_sample_data(n_sales=500, n_purchases=400, seed=42):
    """Generate realistic sample Odoo data for demo/testing."""
    random.seed(seed)
    np.random.seed(seed)

    HSN_CODES = [
        ("8471", 18), ("6109", 5), ("8517", 18), ("3004", 12),
        ("9403", 18), ("8528", 28), ("7308", 18), ("3926", 18),
        ("4820", 12), ("8544", 18), ("8443", 18), ("3808", 18),
        ("7318", 18), ("8504", 18), ("9001", 12),
    ]

    def _gstin(state_code):
        pan = ''.join(random.choices("ABCDEFGHIJKLMNOPQRSTUVWXYZ", k=5)) + \
              ''.join(random.choices("0123456789", k=4)) + \
              random.choice("ABCDEFGHIJKLMNOPQRSTUVWXYZ")
        return f"{state_code}{pan}{random.choice('12345')}Z{random.choice('ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789')}"

    state_codes = list(STATES.keys())[:20]

    # Vendors
    vendors = []
    for i in range(50):
        sc = random.choice(state_codes)
        vendors.append({"gstin": _gstin(sc), "name": f"Vendor_{i+1:03d} Enterprises", "state": sc})

    # Customers
    customers = []
    for i in range(70):
        sc = random.choice(state_codes)
        has_gstin = random.random() < 0.65
        customers.append({
            "gstin": _gstin(sc) if has_gstin else "",
            "name": f"Customer_{i+1:03d} Pvt Ltd" if has_gstin else f"Walk-in {i+1:03d}",
            "state": sc
        })

    # ─── Sales Register ───
    base = datetime(2025, 1, 1)
    sales_rows = []
    for i in range(n_sales):
        c = random.choice(customers)
        hsn, rate = random.choice(HSN_CODES)
        dt = base + timedelta(days=random.randint(0, 89))
        taxable = round(random.uniform(5000, 600000), 2)
        pos = c["state"]
        interstate = pos != OWN_STATE
        igst = round(taxable * rate / 100, 2) if interstate else 0
        cgst = 0 if interstate else round(taxable * rate / 200, 2)
        sgst = 0 if interstate else round(taxable * rate / 200, 2)
        inv_val = round(taxable + igst + cgst + sgst, 2)
        is_cn = random.random() < 0.05
        is_export = (pos in ["96", "97"]) or (random.random() < 0.02)
        prefix = "CN" if is_cn else ("EXP" if is_export else "INV")
        sales_rows.append({
            "Invoice Number": f"{prefix}/2025/{i+1:05d}",
            "Invoice Date": dt.strftime("%d-%m-%Y"),
            "Customer GSTIN": c["gstin"],
            "Customer Name": c["name"],
            "Invoice Value": inv_val,
            "Taxable Value": taxable,
            "IGST": igst, "CGST": cgst, "SGST": sgst,
            "Place of Supply": pos,
            "HSN Code": hsn,
            "Tax Rate": rate,
            "Document Type": "Credit Note" if is_cn else ("Export" if is_export else "Invoice"),
        })
    # Inject anomalies
    for _ in range(4):
        sales_rows.append(sales_rows[random.randint(0, len(sales_rows)-1)].copy())
    sales_rows[8]["Customer GSTIN"] = "INVALID123"
    sales_rows[22]["Customer GSTIN"] = "29ABCDE1234Z"
    sales_rows[45]["Customer GSTIN"] = "O7AAAAA0000A1Z5"  # OCR error
    sales_rows[100]["HSN Code"] = ""  # Missing HSN

    # ─── Purchase Register ───
    purch_rows = []
    for i in range(n_purchases):
        v = random.choice(vendors)
        hsn, rate = random.choice(HSN_CODES)
        dt = base + timedelta(days=random.randint(0, 89))
        taxable = round(random.uniform(8000, 400000), 2)
        interstate = v["state"] != OWN_STATE
        igst = round(taxable * rate / 100, 2) if interstate else 0
        cgst = 0 if interstate else round(taxable * rate / 200, 2)
        sgst = 0 if interstate else round(taxable * rate / 200, 2)
        inv_val = round(taxable + igst + cgst + sgst, 2)
        purch_rows.append({
            "Vendor GSTIN": v["gstin"],
            "Vendor Name": v["name"],
            "Invoice Number": f"BILL/{i+1:05d}",
            "Invoice Date": dt.strftime("%d-%m-%Y"),
            "Taxable Value": taxable,
            "IGST": igst, "CGST": cgst, "SGST": sgst,
            "Invoice Value": inv_val,
            "HSN Code": hsn,
            "Tax Rate": rate,
            "Reverse Charge": "Y" if random.random() < 0.03 else "N",
            "Blocked ITC": "Y" if random.random() < 0.06 else "N",
            "Block Reason": random.choice(BLOCKED_ITC_SECTIONS) if random.random() < 0.06 else "",
        })
    for _ in range(3):
        purch_rows.append(purch_rows[random.randint(0, len(purch_rows)-1)].copy())

    # ─── Simulated GSTR-2B ───
    n_p = len(purch_rows)
    keep = sorted(random.sample(range(n_p), int(n_p * 0.83)))
    gstr2b_rows = []
    for idx in keep:
        r = purch_rows[idx]
        rec = {
            "Supplier GSTIN": r["Vendor GSTIN"],
            "Supplier Name": r["Vendor Name"],
            "Invoice Number": r["Invoice Number"],
            "Invoice Date": r["Invoice Date"],
            "Taxable Value": r["Taxable Value"],
            "IGST": r["IGST"], "CGST": r["CGST"], "SGST": r["SGST"],
            "Invoice Value": r["Invoice Value"],
            "ITC Available": "Yes" if random.random() < 0.90 else "No",
            "Reason": "" if random.random() < 0.90 else random.choice(
                ["Return not filed", "Cancelled", "Amendment pending", "IMS action"]),
            "Filing Status": "Filed" if random.random() < 0.85 else "Not Filed",
        }
        if random.random() < 0.09:
            diff = round(random.uniform(-800, 800), 2)
            if rec["IGST"] > 0:
                rec["IGST"] = round(max(0, rec["IGST"] + diff), 2)
            else:
                rec["CGST"] = round(max(0, rec["CGST"] + diff/2), 2)
                rec["SGST"] = round(max(0, rec["SGST"] + diff/2), 2)
            rec["Taxable Value"] = round(max(0, rec["Taxable Value"] + diff * 2), 2)
            rec["Invoice Value"] = round(rec["Taxable Value"] + rec["IGST"] + rec["CGST"] + rec["SGST"], 2)
        gstr2b_rows.append(rec)
    for i in range(25):
        v = random.choice(vendors)
        hsn, rate = random.choice(HSN_CODES)
        taxable = round(random.uniform(10000, 250000), 2)
        interstate = v["state"] != OWN_STATE
        igst = round(taxable * rate / 100, 2) if interstate else 0
        cgst = 0 if interstate else round(taxable * rate / 200, 2)
        sgst = 0 if interstate else round(taxable * rate / 200, 2)
        gstr2b_rows.append({
            "Supplier GSTIN": v["gstin"], "Supplier Name": v["name"],
            "Invoice Number": f"EXT2B/{i+1:04d}",
            "Invoice Date": (base + timedelta(days=random.randint(0, 89))).strftime("%d-%m-%Y"),
            "Taxable Value": taxable, "IGST": igst, "CGST": cgst, "SGST": sgst,
            "Invoice Value": round(taxable + igst + cgst + sgst, 2),
            "ITC Available": "Yes", "Reason": "", "Filing Status": "Filed",
        })

    return pd.DataFrame(sales_rows), pd.DataFrame(purch_rows), pd.DataFrame(gstr2b_rows)


# ── GSTR-1 PREPARATION ───────────────────────────────────────

def prepare_gstr1(sales_df, own_state=OWN_STATE):
    df = sales_df.copy()
    df["GSTIN_Status"] = df["Customer GSTIN"].apply(validate_gstin)
    df["OCR_Issues"] = df["Customer GSTIN"].apply(lambda g: "; ".join(detect_ocr_errors(g)) if detect_ocr_errors(g) else "")
    df["_date"] = pd.to_datetime(df["Invoice Date"], format="%d-%m-%Y", errors="coerce")

    # Duplicates
    dup_cols = ["Invoice Number", "Invoice Date", "Customer GSTIN", "Invoice Value"]
    df["Is_Duplicate"] = df.duplicated(subset=dup_cols, keep="first")

    # Tax validation
    df["_expected_tax"] = (df["Taxable Value"] * df["Tax Rate"] / 100).round(2)
    df["_actual_tax"] = (df["IGST"] + df["CGST"] + df["SGST"]).round(2)
    df["Tax_Mismatch"] = (df["_expected_tax"] - df["_actual_tax"]).abs() > 1.0

    # Missing HSN
    df["Missing_HSN"] = df["HSN Code"].isna() | (df["HSN Code"].astype(str).str.strip() == "")

    # Categorize
    def _cat(row):
        if row["Document Type"] == "Credit Note":
            return "Credit/Debit Note"
        if row["Document Type"] == "Export":
            return "Export"
        if row["GSTIN_Status"] == "Valid":
            return "B2B"
        if row["Place of Supply"] != own_state and row["Invoice Value"] > 250000:
            return "B2C Large"
        return "B2C Small"

    df["Category"] = df.apply(_cat, axis=1)

    b2b = df[df["Category"] == "B2B"]
    b2c_large = df[df["Category"] == "B2C Large"]
    b2c_small = df[df["Category"] == "B2C Small"]
    cdn = df[df["Category"] == "Credit/Debit Note"]
    exports = df[df["Category"] == "Export"]

    hsn_summary = df.groupby(["HSN Code", "Tax Rate"]).agg(
        Count=("Invoice Number", "count"),
        Taxable_Value=("Taxable Value", "sum"),
        IGST=("IGST", "sum"), CGST=("CGST", "sum"), SGST=("SGST", "sum"),
        Total_Value=("Invoice Value", "sum")
    ).reset_index().round(2)

    # Monthly breakdown
    df["Month"] = df["_date"].dt.to_period("M").astype(str)
    monthly = df.groupby("Month").agg(
        Invoices=("Invoice Number", "count"),
        Taxable=("Taxable Value", "sum"),
        IGST=("IGST", "sum"), CGST=("CGST", "sum"), SGST=("SGST", "sum"),
        Total=("Invoice Value", "sum")
    ).reset_index().round(2)

    # Errors
    errors = []
    for _, r in df[df["GSTIN_Status"].str.startswith("Invalid")].iterrows():
        errors.append({"Invoice": r["Invoice Number"], "Type": "Invalid GSTIN",
                        "Detail": f"{r['Customer GSTIN']} → {r['GSTIN_Status']}", "Severity": "HIGH"})
    for _, r in df[df["Is_Duplicate"]].iterrows():
        errors.append({"Invoice": r["Invoice Number"], "Type": "Duplicate",
                        "Detail": f"Duplicate of existing invoice", "Severity": "CRITICAL"})
    for _, r in df[df["Tax_Mismatch"]].iterrows():
        errors.append({"Invoice": r["Invoice Number"], "Type": "Tax Mismatch",
                        "Detail": f"Expected ₹{r['_expected_tax']:,.2f}, Got ₹{r['_actual_tax']:,.2f}", "Severity": "HIGH"})
    for _, r in df[df["Missing_HSN"]].iterrows():
        errors.append({"Invoice": r["Invoice Number"], "Type": "Missing HSN",
                        "Detail": "HSN code is blank", "Severity": "MEDIUM"})
    for _, r in df[df["OCR_Issues"] != ""].iterrows():
        errors.append({"Invoice": r["Invoice Number"], "Type": "Possible OCR Error",
                        "Detail": r["OCR_Issues"], "Severity": "MEDIUM"})

    summary = {
        "Total Invoices": len(df),
        "B2B Invoices": len(b2b),
        "B2C Large": len(b2c_large),
        "B2C Small": len(b2c_small),
        "Credit/Debit Notes": len(cdn),
        "Exports": len(exports),
        "Duplicates Detected": int(df["Is_Duplicate"].sum()),
        "Invalid GSTINs": int(df["GSTIN_Status"].str.startswith("Invalid").sum()),
        "Missing HSN Codes": int(df["Missing_HSN"].sum()),
        "Tax Mismatches": int(df["Tax_Mismatch"].sum()),
        "OCR Errors Flagged": int((df["OCR_Issues"] != "").sum()),
        "Total Taxable Value": round(df["Taxable Value"].sum(), 2),
        "Total IGST": round(df["IGST"].sum(), 2),
        "Total CGST": round(df["CGST"].sum(), 2),
        "Total SGST": round(df["SGST"].sum(), 2),
        "Total Invoice Value": round(df["Invoice Value"].sum(), 2),
        "GST Liability (IGST)": round(df["IGST"].sum(), 2),
        "GST Liability (CGST+SGST)": round(df["CGST"].sum() + df["SGST"].sum(), 2),
    }

    return {
        "full_data": df, "b2b": b2b, "b2c_large": b2c_large,
        "b2c_small": b2c_small, "cdn": cdn, "exports": exports,
        "hsn_summary": hsn_summary, "monthly": monthly,
        "errors": pd.DataFrame(errors) if errors else pd.DataFrame(columns=["Invoice","Type","Detail","Severity"]),
        "summary": summary,
    }


# ── GSTR-2B RECONCILIATION ───────────────────────────────────

def reconcile_2b(purchase_df, gstr2b_df):
    books = purchase_df.copy()
    portal = gstr2b_df.copy()

    books["_key"] = books["Vendor GSTIN"].str.strip().str.upper() + "|" + books["Invoice Number"].str.strip().str.upper()
    portal["_key"] = portal["Supplier GSTIN"].str.strip().str.upper() + "|" + portal["Invoice Number"].str.strip().str.upper()

    bk = set(books["_key"])
    pk = set(portal["_key"])
    matched_keys = bk & pk
    in_books_only = bk - pk
    in_2b_only = pk - bk

    matched_records, tax_diffs = [], []
    for key in matched_keys:
        b = books[books["_key"] == key].iloc[0]
        p = portal[portal["_key"] == key].iloc[0]
        tax_b = round(b["IGST"] + b["CGST"] + b["SGST"], 2)
        tax_p = round(p["IGST"] + p["CGST"] + p["SGST"], 2)
        tv_diff = round(abs(b["Taxable Value"] - p["Taxable Value"]), 2)
        t_diff = round(abs(tax_b - tax_p), 2)
        rec = {
            "Vendor GSTIN": b["Vendor GSTIN"], "Vendor Name": b["Vendor Name"],
            "Invoice Number": b["Invoice Number"], "Invoice Date": b["Invoice Date"],
            "Books Taxable": b["Taxable Value"], "2B Taxable": p["Taxable Value"],
            "Taxable Diff": tv_diff,
            "Books Tax": tax_b, "2B Tax": tax_p, "Tax Diff": t_diff,
            "Status": "Tax Mismatch" if t_diff > 1 else ("Value Mismatch" if tv_diff > 1 else "Matched"),
            "Filing Status": p.get("Filing Status", ""),
            "ITC Available": p.get("ITC Available", ""),
        }
        matched_records.append(rec)
        if t_diff > 1 or tv_diff > 1:
            tax_diffs.append(rec)

    matched_df = pd.DataFrame(matched_records)
    tax_diff_df = pd.DataFrame(tax_diffs)

    missing_in_2b = books[books["_key"].isin(in_books_only)][[
        "Vendor GSTIN", "Vendor Name", "Invoice Number", "Invoice Date",
        "Taxable Value", "IGST", "CGST", "SGST", "Invoice Value"
    ]].copy()
    missing_in_2b["Action"] = "Follow up with vendor — ITC at risk"

    missing_in_books = portal[portal["_key"].isin(in_2b_only)][[
        "Supplier GSTIN", "Supplier Name", "Invoice Number", "Invoice Date",
        "Taxable Value", "IGST", "CGST", "SGST", "Invoice Value"
    ]].copy()
    missing_in_books.columns = ["Vendor GSTIN", "Vendor Name", "Invoice Number", "Invoice Date",
                                 "Taxable Value", "IGST", "CGST", "SGST", "Invoice Value"]
    missing_in_books["Action"] = "Verify with accounts team — book if valid"

    # Duplicate detection
    dup_mask = books.duplicated(subset=["Vendor GSTIN", "Invoice Number", "Invoice Date"], keep=False)
    dup_invoices = books[dup_mask][[
        "Vendor GSTIN", "Vendor Name", "Invoice Number", "Invoice Date", "Invoice Value"
    ]].copy()

    perfect = int((matched_df["Status"] == "Matched").sum()) if len(matched_df) > 0 else 0
    summary = {
        "Total in Books": len(books),
        "Total in GSTR-2B": len(portal),
        "Perfectly Matched": perfect,
        "Value/Tax Mismatches": len(tax_diff_df),
        "In Books, Missing in 2B": len(missing_in_2b),
        "In 2B, Missing in Books": len(missing_in_books),
        "Duplicate Invoices": len(dup_invoices),
        "Match Rate (%)": round(perfect / max(len(matched_keys), 1) * 100, 1),
    }

    return {
        "matched": matched_df, "tax_diffs": tax_diff_df,
        "missing_in_2b": missing_in_2b, "missing_in_books": missing_in_books,
        "duplicates": dup_invoices, "summary": summary,
    }


# ── ITC ELIGIBILITY ANALYSIS ─────────────────────────────────

def analyze_itc(purchase_df, gstr2b_df, recon):
    books = purchase_df.copy()
    portal = gstr2b_df.copy()

    books["Total_Tax"] = books["IGST"] + books["CGST"] + books["SGST"]
    portal["Total_Tax"] = portal["IGST"] + portal["CGST"] + portal["SGST"]

    itc_books = round(books["Total_Tax"].sum(), 2)
    itc_2b = round(portal["Total_Tax"].sum(), 2)

    eligible_mask = portal["ITC Available"] == "Yes" if "ITC Available" in portal.columns else pd.Series([True]*len(portal))
    itc_eligible = round(portal.loc[eligible_mask, "Total_Tax"].sum(), 2)
    itc_ineligible = round(itc_2b - itc_eligible, 2)

    blocked_mask = books["Blocked ITC"] == "Y" if "Blocked ITC" in books.columns else pd.Series([False]*len(books))
    blocked_itc = round(books.loc[blocked_mask, "Total_Tax"].sum(), 2)

    rc_mask = books["Reverse Charge"] == "Y" if "Reverse Charge" in books.columns else pd.Series([False]*len(books))
    rc_itc = round(books.loc[rc_mask, "Total_Tax"].sum(), 2)

    missing_2b = recon["missing_in_2b"]
    itc_not_in_2b = round((missing_2b["IGST"] + missing_2b["CGST"] + missing_2b["SGST"]).sum(), 2) if len(missing_2b) > 0 else 0

    mismatch_impact = round(recon["tax_diffs"]["Tax Diff"].sum(), 2) if len(recon["tax_diffs"]) > 0 else 0

    # Vendor-wise analysis
    bv = books.groupby("Vendor GSTIN").agg(
        Vendor_Name=("Vendor Name", "first"),
        Invoice_Count=("Invoice Number", "count"),
        Books_Taxable=("Taxable Value", "sum"),
        Books_ITC=("Total_Tax", "sum"),
    ).reset_index()

    pv = portal.groupby("Supplier GSTIN").agg(
        Portal_Taxable=("Taxable Value", "sum"),
        Portal_ITC=("Total_Tax", "sum"),
        Filing_Status=("Filing Status", lambda x: "Not Filed" if "Not Filed" in x.values else "Filed"),
        ITC_Status=("ITC Available", lambda x: "Partially Blocked" if "No" in x.values else "All Available"),
    ).reset_index().rename(columns={"Supplier GSTIN": "Vendor GSTIN"})

    vendor_itc = bv.merge(pv, on="Vendor GSTIN", how="outer").fillna(0)
    vendor_itc["Books_ITC"] = vendor_itc["Books_ITC"].astype(float).round(2)
    vendor_itc["Portal_ITC"] = vendor_itc["Portal_ITC"].astype(float).round(2)
    vendor_itc["ITC_Difference"] = (vendor_itc["Books_ITC"] - vendor_itc["Portal_ITC"]).round(2)
    vendor_itc["Risk"] = vendor_itc.apply(
        lambda r: "CRITICAL" if abs(r["ITC_Difference"]) > 50000 else
                  ("HIGH" if abs(r["ITC_Difference"]) > 10000 else
                   ("MEDIUM" if abs(r["ITC_Difference"]) > 1000 else "LOW")), axis=1
    )
    vendor_itc = vendor_itc.sort_values("ITC_Difference", key=abs, ascending=False)

    not_filed = portal[portal.get("Filing Status", pd.Series()) == "Not Filed"]["Supplier GSTIN"].unique().tolist() if "Filing Status" in portal.columns else []

    # Blocked ITC breakdown
    blocked_detail = pd.DataFrame()
    if "Block Reason" in books.columns:
        blocked_detail = books[blocked_mask].groupby("Block Reason").agg(
            Count=("Invoice Number", "count"),
            Blocked_Amount=("Total_Tax", "sum")
        ).reset_index().round(2)

    summary = {
        "Total ITC as per Books": itc_books,
        "Total ITC as per GSTR-2B": itc_2b,
        "Eligible ITC": itc_eligible,
        "Ineligible ITC (2B)": itc_ineligible,
        "Blocked ITC (Sec 17(5))": blocked_itc,
        "Reverse Charge ITC": rc_itc,
        "ITC Not in 2B": itc_not_in_2b,
        "Tax Mismatch Impact": mismatch_impact,
        "Net ITC Difference": round(itc_books - itc_eligible, 2),
        "Claimable ITC": round(itc_eligible - blocked_itc, 2),
        "Vendors Not Filed Count": len(not_filed),
    }

    return {
        "summary": summary,
        "vendor_itc": vendor_itc,
        "not_filed_vendors": not_filed,
        "blocked_detail": blocked_detail,
    }


# ── ALERT GENERATION ──────────────────────────────────────────

def generate_alerts(gstr1, recon, itc):
    alerts = []

    # GSTR-1 alerts
    s = gstr1["summary"]
    if s["Duplicates Detected"] > 0:
        alerts.append({"Severity": "CRITICAL", "Category": "GSTR-1",
            "Alert": f"{s['Duplicates Detected']} duplicate invoices detected",
            "Action": "Remove duplicates before filing GSTR-1", "Impact": "Filing rejection risk"})
    if s["Invalid GSTINs"] > 0:
        alerts.append({"Severity": "HIGH", "Category": "GSTR-1",
            "Alert": f"{s['Invalid GSTINs']} invalid GSTIN(s) in sales register",
            "Action": "Correct GSTINs in Odoo before filing", "Impact": "Invoice rejection on portal"})
    if s["Missing HSN Codes"] > 0:
        alerts.append({"Severity": "MEDIUM", "Category": "GSTR-1",
            "Alert": f"{s['Missing HSN Codes']} invoices missing HSN code",
            "Action": "Update HSN codes in Odoo product master", "Impact": "HSN summary will be incomplete"})
    if s["OCR Errors Flagged"] > 0:
        alerts.append({"Severity": "MEDIUM", "Category": "GSTR-1",
            "Alert": f"{s['OCR Errors Flagged']} possible OCR errors in GSTIN",
            "Action": "Verify flagged GSTINs against original documents", "Impact": "Incorrect B2B reporting"})

    # Reconciliation alerts
    r = recon["summary"]
    if r["In Books, Missing in 2B"] > 0:
        alerts.append({"Severity": "CRITICAL", "Category": "Reconciliation",
            "Alert": f"{r['In Books, Missing in 2B']} purchase invoices missing from GSTR-2B",
            "Action": "Contact vendors to file their GSTR-1", "Impact": f"₹{itc['summary']['ITC Not in 2B']:,.0f} ITC at risk"})
    if r["Value/Tax Mismatches"] > 0:
        alerts.append({"Severity": "HIGH", "Category": "Reconciliation",
            "Alert": f"{r['Value/Tax Mismatches']} invoices have value/tax differences",
            "Action": "Reconcile with vendor invoices and amend", "Impact": f"₹{itc['summary']['Tax Mismatch Impact']:,.0f} tax difference"})
    if r["In 2B, Missing in Books"] > 0:
        alerts.append({"Severity": "HIGH", "Category": "Reconciliation",
            "Alert": f"{r['In 2B, Missing in Books']} invoices in 2B not recorded in books",
            "Action": "Verify and book valid invoices in Odoo", "Impact": "Potential unclaimed ITC"})
    if r["Duplicate Invoices"] > 0:
        alerts.append({"Severity": "HIGH", "Category": "Reconciliation",
            "Alert": f"{r['Duplicate Invoices']} duplicate purchase invoices detected",
            "Action": "Remove duplicate entries from purchase register", "Impact": "ITC overclaim risk"})

    # ITC alerts
    i = itc["summary"]
    if i["Vendors Not Filed Count"] > 0:
        alerts.append({"Severity": "CRITICAL", "Category": "ITC",
            "Alert": f"{i['Vendors Not Filed Count']} vendors have not filed GSTR-1",
            "Action": "Send reminder to non-compliant vendors", "Impact": "ITC will be blocked/reversed"})
    if i["Blocked ITC (Sec 17(5))"] > 0:
        alerts.append({"Severity": "MEDIUM", "Category": "ITC",
            "Alert": f"₹{i['Blocked ITC (Sec 17(5))']:,.0f} blocked under Section 17(5)",
            "Action": "Review blocked items — ensure correct classification", "Impact": "Non-claimable ITC"})
    if abs(i["Net ITC Difference"]) > 50000:
        alerts.append({"Severity": "CRITICAL", "Category": "ITC",
            "Alert": f"Net ITC difference of ₹{i['Net ITC Difference']:,.0f}",
            "Action": "Resolve all mismatches before filing GSTR-3B", "Impact": "Incorrect ITC claim in 3B"})

    return pd.DataFrame(alerts)


# ── GSTR-1 JSON EXPORT ────────────────────────────────────────

def export_gstr1_json(gstr1_result):
    b2b = gstr1_result["b2b"]
    out = {"b2b": [], "b2cl": [], "b2cs": [], "cdnr": [], "hsn": {"data": []}}

    for gstin, grp in b2b.groupby("Customer GSTIN"):
        invs = []
        for _, r in grp.iterrows():
            invs.append({
                "inum": r["Invoice Number"], "idt": r["Invoice Date"],
                "val": float(r["Invoice Value"]), "pos": str(r["Place of Supply"]),
                "rchrg": "N", "inv_typ": "R",
                "itms": [{"num": 1, "itm_det": {
                    "txval": float(r["Taxable Value"]), "rt": float(r["Tax Rate"]),
                    "iamt": float(r["IGST"]), "camt": float(r["CGST"]),
                    "samt": float(r["SGST"]), "csamt": 0
                }}]
            })
        out["b2b"].append({"ctin": gstin, "inv": invs})

    for _, r in gstr1_result["hsn_summary"].iterrows():
        out["hsn"]["data"].append({
            "hsn_sc": str(r["HSN Code"]), "qty": int(r["Count"]),
            "txval": float(r["Taxable_Value"]), "rt": float(r["Tax Rate"]),
            "iamt": float(r["IGST"]), "camt": float(r["CGST"]),
            "samt": float(r["SGST"]), "csamt": 0,
        })

    return out


# ── MONTHLY SCHEDULE ──────────────────────────────────────────

def get_monthly_schedule(year, month):
    from calendar import monthrange
    last_day = monthrange(year, month)[1]
    base = datetime(year, month, 1)
    return {
        "Period": f"{base.strftime('%B %Y')}",
        "Day 1-5": f"Fetch Odoo data ({base.strftime('%d-%b')} to {datetime(year,month,5).strftime('%d-%b')})",
        "Day 6": f"Download GSTR-2B from GST Portal ({datetime(year,month,6).strftime('%d-%b')})",
        "Day 7": f"Run reconciliation engine ({datetime(year,month,7).strftime('%d-%b')})",
        "Day 8": f"Generate compliance reports ({datetime(year,month,8).strftime('%d-%b')})",
        "Day 10": f"GSTR-1 filing deadline (11th of next month)",
        "Day 13": f"GSTR-3B preparation deadline",
        "Day 20": f"GSTR-3B filing deadline (20th of next month)",
    }
