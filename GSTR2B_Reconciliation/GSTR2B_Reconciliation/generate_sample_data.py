"""
Generate sample Tally and GSTR2B data for testing.
"""
import pandas as pd
from pathlib import Path

SAMPLE_DIR = Path(__file__).parent / "sample_data"
SAMPLE_DIR.mkdir(exist_ok=True)

# ---- Sample Tally Data ----
tally_data = {
    "Voucher Type": ["Purchase"] * 12 + ["Journal"] * 2,
    "Voucher Date": [
        "01-03-2026", "03-03-2026", "05-03-2026", "07-03-2026", "10-03-2026",
        "12-03-2026", "14-03-2026", "15-03-2026", "17-03-2026", "18-03-2026",
        "20-03-2026", "22-03-2026", "25-03-2026", "28-03-2026",
    ],
    "Voucher Number": [f"PUR/{i:04d}" for i in range(1, 15)],
    "Party Name": [
        "ABC Traders Pvt Ltd", "XYZ Electronics", "PQR Plastics LLP",
        "Sunrise Components Pvt Ltd", "National Filaments Co",
        "Global Resins India Pvt Ltd", "Metro Hardware", "DEF Electricals",
        "Sunrise Components Pvt Ltd", "JKL Industries",
        "MNO Enterprises", "ABC Traders Pvt Ltd", "QRS Technologies", "National Filaments Co",
    ],
    "GSTIN": [
        "29AABCA1234R1ZM", "27BBCDE5678S1ZN", "29CDEFG9012T1ZO",
        "29HIJKL3456U1ZP", "33MNOPQ7890V1ZQ",
        "29RSTUV1234W1ZR", "29WXYZ5678A1ZS", "27ABCDE9012B1ZT",
        "29HIJKL3456U1ZP", "29FGHIJ3456C1ZU",
        "29KLMNO7890D1ZV", "29AABCA1234R1ZM", "29PQRST1234E1ZW", "33MNOPQ7890V1ZQ",
    ],
    "Vendor Bill Number": [
        "INV/2026/001", "TAX-INV-445", "PL/MAR/101", "SC-2026-0088",
        "NF-INV-2026-032", "GRI-0567", "MH-2026-99", "DEF/26/112",
        "SC-2026-0099", "JKL/INV/556",
        "", "INV/2026/010", "QRS-T-2026-015", "NF-INV-2026-040",
    ],
    "Taxable Value": [
        50000, 125000, 32000, 78000, 15600,
        210000, 8500, 44000, 65000, 92000,
        18000, 55000, 140000, 22000,
    ],
    "IGST Amount": [
        0, 22500, 0, 0, 2808,
        0, 0, 7920, 0, 0,
        0, 0, 0, 3960,
    ],
    "CGST Amount": [
        4500, 0, 2880, 7020, 0,
        18900, 765, 0, 5850, 8280,
        1620, 4950, 12600, 0,
    ],
    "SGST Amount": [
        4500, 0, 2880, 7020, 0,
        18900, 765, 0, 5850, 8280,
        1620, 4950, 12600, 0,
    ],
    "Total Invoice Value": [
        59000, 147500, 37760, 92040, 18408,
        248800, 10030, 51920, 76700, 108560,
        21240, 64900, 165200, 25960,
    ],
    "Narration": [
        "Being purchase of raw materials vide Inv INV/2026/001",
        "Purchase of electronic components",
        "Plastic granules for production",
        "Component purchase as per PO-445",
        "Filament purchase for 3D printers",
        "Resin supply for SLA printing",
        "Hardware and fasteners",
        "Electrical components for assembly",
        "Spare parts purchase ref SC-2026-0099",
        "Industrial supplies",
        "Stationery and office supplies Invoice REF: MNO-22",
        "Follow-up order raw materials",
        "IT equipment and accessories",
        "Journal entry for ITC reversal ref NF-INV-2026-040",
    ],
}

tally_df = pd.DataFrame(tally_data)
tally_path = SAMPLE_DIR / "Sample_Tally_Export.xlsx"
tally_df.to_excel(tally_path, index=False)
print(f"Created: {tally_path}")

# ---- Sample GSTR2B Data ----
gstr2b_data = {
    "Supplier GSTIN": [
        "29AABCA1234R1ZM", "27BBCDE5678S1ZN", "29CDEFG9012T1ZO",
        "29HIJKL3456U1ZP", "33MNOPQ7890V1ZQ",
        "29RSTUV1234W1ZR", "29WXYZ5678A1ZS", "27ABCDE9012B1ZT",
        "29HIJKL3456U1ZP", "29FGHIJ3456C1ZU",
        "29UVWXY5678F1ZX",  # Extra - not in Tally
        "29AABCA1234R1ZM",
        "36ZZZAA1234G1ZY",  # Extra - not in Tally
    ],
    "Trade/Legal name": [
        "ABC TRADERS PVT. LTD.", "XYZ ELECTRONICS", "PQR PLASTICS LLP",
        "SUNRISE COMPONENTS PVT. LTD.", "NATIONAL FILAMENTS CO",
        "GLOBAL RESINS INDIA PVT. LTD.", "METRO HARDWARE", "DEF ELECTRICALS",
        "SUNRISE COMPONENTS PVT. LTD.", "JKL INDUSTRIES",
        "RST POLYMERS PVT. LTD.",
        "ABC TRADERS PVT. LTD.",
        "NEW VENDOR SOLUTIONS",
    ],
    "Invoice Number": [
        "INV/2026/001", "TAX-INV-445", "PL/MAR/101",
        "SC-2026-0088", "NF-INV-2026-032",
        "GRI-0567", "MH-2026-99", "DEF/26/112",
        "SC-2026-0099", "JKL/INV/556",
        "RST-2026-078",
        "INV/2026/010",
        "NVS-2026-001",
    ],
    "Invoice Date": [
        "01-03-2026", "03-03-2026", "05-03-2026",
        "07-03-2026", "10-03-2026",
        "12-03-2026", "14-03-2026", "15-03-2026",
        "17-03-2026", "18-03-2026",
        "19-03-2026",
        "22-03-2026",
        "25-03-2026",
    ],
    "Taxable Value": [
        50000, 125000, 32000,
        78000, 15600,
        210000, 8500, 44500,  # DEF has slight mismatch: 44500 vs 44000
        65000, 92000,
        35000,
        55000,
        72000,
    ],
    "Integrated Tax": [
        0, 22500, 0,
        0, 2808,
        0, 0, 8010,  # Mismatch
        0, 0,
        0,
        0,
        0,
    ],
    "Central Tax": [
        4500, 0, 2880,
        7020, 0,
        18900, 765, 0,
        5850, 8280,
        3150,
        4950,
        6480,
    ],
    "State Tax": [
        4500, 0, 2880,
        7020, 0,
        18900, 765, 0,
        5850, 8280,
        3150,
        4950,
        6480,
    ],
    "Total Tax": [
        9000, 22500, 5760,
        14040, 2808,
        37800, 1530, 8010,
        11700, 16560,
        6300,
        9900,
        12960,
    ],
}

gstr2b_df = pd.DataFrame(gstr2b_data)
gstr2b_path = SAMPLE_DIR / "Sample_GSTR2B.xlsx"
gstr2b_df.to_excel(gstr2b_path, index=False)
print(f"Created: {gstr2b_path}")

print("\nSample files created successfully!")
print(f"\nTo test, run:\n  python run.py --tally \"{tally_path}\" --gstr2b \"{gstr2b_path}\" --month 2026-03")
