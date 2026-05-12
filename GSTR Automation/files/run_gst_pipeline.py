"""
GST Compliance Automation — Main Orchestrator
==============================================
Runs the complete pipeline:
  1. Generate/Fetch data
  2. Process GSTR-1
  3. Reconcile GSTR-2B
  4. Analyze ITC
  5. Generate alerts
  6. Build Excel report
  7. Export JSON
  8. Generate dashboard data
"""

import json
import sys
import os
sys.path.insert(0, "/home/claude")

from gst_engine import (
    generate_sample_data, prepare_gstr1, reconcile_2b,
    analyze_itc, generate_alerts, export_gstr1_json, get_monthly_schedule
)
from report_builder import build_workbook

BANNER = """
╔══════════════════════════════════════════════════════════════╗
║          GST COMPLIANCE AUTOMATION ENGINE v2.0              ║
║    Odoo ERP → GSTR-1 / GSTR-2B / ITC / Alerts              ║
║    Automated Monthly GST Return Preparation                 ║
╚══════════════════════════════════════════════════════════════╝
"""

def main():
    print(BANNER)

    # ── STEP 1: DATA ──
    print("━" * 60)
    print("  STEP 1 │ Generating sample Odoo data")
    print("━" * 60)
    sales_df, purchase_df, gstr2b_df = generate_sample_data(n_sales=500, n_purchases=400)
    print(f"  ✓ Sales Register:    {len(sales_df):,} invoices")
    print(f"  ✓ Purchase Register: {len(purchase_df):,} invoices")
    print(f"  ✓ GSTR-2B Records:   {len(gstr2b_df):,} entries")

    # ── STEP 2: GSTR-1 ──
    print(f"\n{'━' * 60}")
    print("  STEP 2 │ Preparing GSTR-1")
    print("━" * 60)
    gstr1 = prepare_gstr1(sales_df)
    s = gstr1["summary"]
    print(f"  ✓ B2B Invoices:      {s['B2B Invoices']:,}")
    print(f"  ✓ B2C Large:         {s['B2C Large']:,}")
    print(f"  ✓ B2C Small:         {s['B2C Small']:,}")
    print(f"  ✓ Credit/Debit Notes:{s['Credit/Debit Notes']:,}")
    print(f"  ✓ Exports:           {s['Exports']:,}")
    print(f"  ⚠ Duplicates:        {s['Duplicates Detected']}")
    print(f"  ⚠ Invalid GSTINs:    {s['Invalid GSTINs']}")
    print(f"  ⚠ Missing HSN:       {s['Missing HSN Codes']}")
    print(f"  ⚠ OCR Errors:        {s['OCR Errors Flagged']}")
    print(f"  ─────────────────────────────────")
    print(f"  Total Taxable:       ₹{s['Total Taxable Value']:>14,.2f}")
    print(f"  IGST Liability:      ₹{s['GST Liability (IGST)']:>14,.2f}")
    print(f"  CGST+SGST Liability: ₹{s['GST Liability (CGST+SGST)']:>14,.2f}")

    # ── STEP 3: RECONCILIATION ──
    print(f"\n{'━' * 60}")
    print("  STEP 3 │ Reconciling GSTR-2B")
    print("━" * 60)
    recon = reconcile_2b(purchase_df, gstr2b_df)
    r = recon["summary"]
    print(f"  ✓ Perfectly Matched: {r['Perfectly Matched']:,}")
    print(f"  ⚠ Mismatches:        {r['Value/Tax Mismatches']:,}")
    print(f"  ✗ Missing in 2B:     {r['In Books, Missing in 2B']:,}")
    print(f"  ✗ Missing in Books:  {r['In 2B, Missing in Books']:,}")
    print(f"  ⚠ Duplicates:        {r['Duplicate Invoices']:,}")
    print(f"  Match Rate:          {r['Match Rate (%)']:.1f}%")

    # ── STEP 4: ITC ANALYSIS ──
    print(f"\n{'━' * 60}")
    print("  STEP 4 │ Analyzing ITC Eligibility")
    print("━" * 60)
    itc = analyze_itc(purchase_df, gstr2b_df, recon)
    i = itc["summary"]
    print(f"  ITC as per Books:    ₹{i['Total ITC as per Books']:>14,.2f}")
    print(f"  ITC as per 2B:       ₹{i['Total ITC as per GSTR-2B']:>14,.2f}")
    print(f"  Eligible ITC:        ₹{i['Eligible ITC']:>14,.2f}")
    print(f"  Blocked ITC:         ₹{i['Blocked ITC (Sec 17(5))']:>14,.2f}")
    print(f"  ITC Not in 2B:       ₹{i['ITC Not in 2B']:>14,.2f}")
    print(f"  ─────────────────────────────────")
    print(f"  NET ITC DIFFERENCE:  ₹{i['Net ITC Difference']:>14,.2f}")
    print(f"  CLAIMABLE ITC:       ₹{i['Claimable ITC']:>14,.2f}")
    print(f"  Vendors Not Filed:   {i['Vendors Not Filed Count']}")

    # ── STEP 5: ALERTS ──
    print(f"\n{'━' * 60}")
    print("  STEP 5 │ Generating Compliance Alerts")
    print("━" * 60)
    alerts_df = generate_alerts(gstr1, recon, itc)
    for _, a in alerts_df.iterrows():
        icon = {"CRITICAL": "🔴", "HIGH": "🟠", "MEDIUM": "🟡", "LOW": "🟢"}.get(a["Severity"], "⚪")
        print(f"  {icon} [{a['Severity']}] {a['Alert']}")
        print(f"      → {a['Action']}")

    # ── STEP 6: EXCEL REPORT ──
    print(f"\n{'━' * 60}")
    print("  STEP 6 │ Building Excel Report (18 sheets)")
    print("━" * 60)
    wb = build_workbook(gstr1, recon, itc, alerts_df, sales_df, purchase_df, gstr2b_df)
    xlsx_path = "/mnt/user-data/outputs/GST_Compliance_Report_v2.xlsx"
    wb.save(xlsx_path)
    print(f"  ✓ Saved: {xlsx_path}")

    # ── STEP 7: JSON EXPORT ──
    print(f"\n{'━' * 60}")
    print("  STEP 7 │ Exporting GSTR-1 JSON for Portal Upload")
    print("━" * 60)
    gstr1_json = export_gstr1_json(gstr1)
    json_path = "/mnt/user-data/outputs/GSTR1_Upload_Ready.json"
    with open(json_path, "w") as f:
        json.dump(gstr1_json, f, indent=2, default=str)
    print(f"  ✓ B2B entries: {len(gstr1_json['b2b'])}")
    print(f"  ✓ HSN entries: {len(gstr1_json['hsn']['data'])}")
    print(f"  ✓ Saved: {json_path}")

    # ── STEP 8: DASHBOARD DATA ──
    print(f"\n{'━' * 60}")
    print("  STEP 8 │ Preparing Dashboard Data")
    print("━" * 60)
    dashboard_data = {
        "generated_at": __import__("datetime").datetime.now().isoformat(),
        "period": "Q1 FY 2025-26 (Jan–Mar 2025)",
        "gstr1_summary": gstr1["summary"],
        "recon_summary": recon["summary"],
        "itc_summary": itc["summary"],
        "alerts": alerts_df.to_dict("records"),
        "monthly": gstr1["monthly"].to_dict("records"),
        "hsn_summary": gstr1["hsn_summary"].to_dict("records"),
        "top_vendors_at_risk": itc["vendor_itc"].head(10)[[
            c for c in ["Vendor GSTIN","Vendor_Name","Books_ITC","Portal_ITC","ITC_Difference","Risk"]
            if c in itc["vendor_itc"].columns
        ]].to_dict("records"),
        "schedule": get_monthly_schedule(2025, 4),
    }
    dash_path = "/mnt/user-data/outputs/dashboard_data.json"
    with open(dash_path, "w") as f:
        json.dump(dashboard_data, f, indent=2, default=str)
    print(f"  ✓ Saved: {dash_path}")

    # ── EXECUTIVE SUMMARY ──
    print(f"\n{'━' * 60}")
    print("  EXECUTIVE SUMMARY")
    print("━" * 60)
    print(f"""
  GSTR-1 READINESS
  ├─ {s['Total Invoices']} invoices categorized across B2B/B2C/CDN/Export
  ├─ {s['Duplicates Detected']} duplicates + {s['Invalid GSTINs']} bad GSTINs need fixing
  └─ Total GST liability: ₹{s['GST Liability (IGST)'] + s['GST Liability (CGST+SGST)']:,.0f}

  GSTR-2B RECONCILIATION
  ├─ {r['Match Rate (%)']:.1f}% invoices matched successfully
  ├─ {r['In Books, Missing in 2B']} invoices missing from portal
  └─ {r['Value/Tax Mismatches']} invoices have value discrepancies

  ITC POSITION
  ├─ Claimable ITC: ₹{i['Claimable ITC']:,.0f}
  ├─ ITC at risk (not in 2B): ₹{i['ITC Not in 2B']:,.0f}
  ├─ Blocked under Sec 17(5): ₹{i['Blocked ITC (Sec 17(5))']:,.0f}
  └─ {i['Vendors Not Filed Count']} vendors have NOT filed returns

  IMMEDIATE ACTIONS REQUIRED
  ├─ Fix {s['Duplicates Detected']} duplicate + {s['Invalid GSTINs']} invalid GSTIN entries
  ├─ Follow up with {i['Vendors Not Filed Count']} non-compliant vendors
  ├─ Reconcile {r['Value/Tax Mismatches']} mismatched invoices
  └─ Book {r['In 2B, Missing in Books']} valid invoices from 2B
""")

    print("═" * 60)
    print("  ALL FILES GENERATED SUCCESSFULLY")
    print("═" * 60)

    return dashboard_data

if __name__ == "__main__":
    dashboard_data = main()
