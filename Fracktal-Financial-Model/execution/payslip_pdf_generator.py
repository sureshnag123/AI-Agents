"""
payslip_pdf_generator.py — Generates professional PDF payslips for Thota Hospitality LLP.

Usage:
    from payslip_pdf_generator import generate_payslip

Requires: reportlab
"""

import os
import re
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm, cm
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, HRFlowable, Image,
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# ---------------------------------------------------------------------------
# Register Calibri fonts (supports ₹ symbol)
# ---------------------------------------------------------------------------
_FONT_DIR = r"C:\Windows\Fonts"
try:
    pdfmetrics.registerFont(TTFont("Calibri", os.path.join(_FONT_DIR, "calibri.ttf")))
    pdfmetrics.registerFont(TTFont("Calibri-Bold", os.path.join(_FONT_DIR, "calibrib.ttf")))
    pdfmetrics.registerFont(TTFont("Calibri-Italic", os.path.join(_FONT_DIR, "calibrii.ttf")))
    pdfmetrics.registerFont(TTFont("Calibri-BoldItalic", os.path.join(_FONT_DIR, "calibriz.ttf")))
    pdfmetrics.registerFontFamily(
        "Calibri",
        normal="Calibri",
        bold="Calibri-Bold",
        italic="Calibri-Italic",
        boldItalic="Calibri-BoldItalic",
    )
    FONT_NAME = "Calibri"
    FONT_BOLD = "Calibri-Bold"
except Exception:
    FONT_NAME = "Helvetica"
    FONT_BOLD = "Helvetica-Bold"

# ---------------------------------------------------------------------------
# Default logo path (can be overridden via parameter)
# ---------------------------------------------------------------------------
DEFAULT_LOGO_PATH = r"C:\Users\Lenovo\OneDrive\Pictures\Screenshots 1\logo.jpg"


# ---------------------------------------------------------------------------
# Colour palette
# ---------------------------------------------------------------------------
BRAND_DARK   = colors.HexColor("#1a237e")   # dark navy
BRAND_ACCENT = colors.HexColor("#283593")   # medium navy
HEADER_BG    = colors.HexColor("#e8eaf6")   # light lavender
WHITE        = colors.white
BLACK        = colors.black
GREEN_DARK   = colors.HexColor("#2e7d32")


def _fmt(val, prefix="₹") -> str:
    """Format a number as Indian currency string."""
    if val is None or val == 0:
        return f"{prefix}0.00" if prefix else "0.00"
    try:
        v = float(val)
        if v == int(v):
            return f"{prefix}{int(v):,}"
        return f"{prefix}{v:,.2f}"
    except (ValueError, TypeError):
        return str(val)


def _fmt_int(val) -> str:
    if val is None:
        return "0"
    try:
        return str(int(float(val)))
    except (ValueError, TypeError):
        return str(val)


def _amount_in_words(amount: float) -> str:
    """Convert amount to words (simplified Indian numbering)."""
    try:
        from num2words import num2words
        integer_part = int(amount)
        decimal_part = int(round((amount - integer_part) * 100))
        words = num2words(integer_part, lang='en_IN').title()
        if decimal_part > 0:
            words += f" and {num2words(decimal_part, lang='en_IN').title()} Paise"
        return f"Rupees {words} Only"
    except ImportError:
        # Fallback without num2words
        return f"Rupees {_fmt(amount)} Only"


def generate_payslip(
    emp: dict,
    company: str,
    month: str,
    output_dir: str,
    logo_path: str | None = None,
) -> str:
    """
    Generate a single PDF payslip and return the file path.

    Parameters
    ----------
    emp        : dict with canonical employee keys (from payroll_reader)
    company    : company name
    month      : display month label (e.g. "JAN-2026")
    output_dir : directory to save the PDF
    logo_path  : path to company logo image (jpg/png). Uses DEFAULT_LOGO_PATH if None.
    """
    # Resolve logo
    if logo_path is None:
        logo_path = DEFAULT_LOGO_PATH
    has_logo = logo_path and os.path.isfile(logo_path)
    os.makedirs(output_dir, exist_ok=True)

    safe_name = re.sub(r"[^\w\-]", "_", emp["emp_name"].strip())
    safe_month = re.sub(r"[^\w\-]", "_", month.strip())
    filename = f"Payslip_{safe_name}_{safe_month}.pdf"
    filepath = os.path.join(output_dir, filename)

    doc = SimpleDocTemplate(
        filepath,
        pagesize=A4,
        topMargin=15 * mm,
        bottomMargin=15 * mm,
        leftMargin=15 * mm,
        rightMargin=15 * mm,
    )

    styles = getSampleStyleSheet()
    elements = []

    # Custom styles
    style_company = ParagraphStyle(
        "Company", parent=styles["Title"],
        fontName=FONT_BOLD, fontSize=16, textColor=BRAND_DARK,
        alignment=TA_CENTER, spaceAfter=2,
    )
    style_subtitle = ParagraphStyle(
        "Subtitle", parent=styles["Normal"],
        fontName=FONT_NAME, fontSize=10, textColor=BRAND_ACCENT,
        alignment=TA_CENTER, spaceAfter=4,
    )
    style_section = ParagraphStyle(
        "Section", parent=styles["Heading3"],
        fontName=FONT_BOLD, fontSize=10, textColor=WHITE,
        backColor=BRAND_DARK, spaceAfter=2, spaceBefore=6,
        leftIndent=4, rightIndent=4,
    )
    style_label = ParagraphStyle(
        "Label", parent=styles["Normal"],
        fontName=FONT_NAME, fontSize=8.5,
        textColor=colors.HexColor("#424242"),
    )
    style_value = ParagraphStyle(
        "Value", parent=styles["Normal"],
        fontName=FONT_BOLD, fontSize=8.5, textColor=BLACK,
    )
    style_net = ParagraphStyle(
        "NetPay", parent=styles["Normal"],
        fontName=FONT_BOLD, fontSize=12, textColor=GREEN_DARK,
        alignment=TA_RIGHT,
    )
    style_footer = ParagraphStyle(
        "Footer", parent=styles["Normal"],
        fontName=FONT_NAME, fontSize=7.5,
        textColor=colors.HexColor("#757575"),
        alignment=TA_CENTER,
    )

    # ---- Header with Logo ----
    if has_logo:
        logo = Image(logo_path, width=50, height=50, hAlign='LEFT')
        header_data = [[
            logo,
            Paragraph(f"{company}<br/><font size=9 color='#283593'>PAYSLIP FOR THE MONTH OF {month}</font>", style_company),
        ]]
        header_table = Table(header_data, colWidths=[60, 380])
        header_table.setStyle(TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("ALIGN", (0, 0), (0, 0), "LEFT"),
            ("ALIGN", (1, 0), (1, 0), "CENTER"),
            ("LEFTPADDING", (0, 0), (-1, -1), 2),
            ("RIGHTPADDING", (0, 0), (-1, -1), 2),
            ("TOPPADDING", (0, 0), (-1, -1), 0),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
        ]))
        elements.append(header_table)
    else:
        elements.append(Paragraph(company, style_company))
        elements.append(Paragraph(f"PAYSLIP FOR THE MONTH OF {month}", style_subtitle))
    elements.append(HRFlowable(width="100%", thickness=1.5, color=BRAND_DARK, spaceAfter=6))

    # ---- Employee Info Table ----
    doj_display = str(emp.get("doj", "N/A") or "N/A")
    if hasattr(emp.get("doj"), "strftime"):
        try:
            doj_display = emp["doj"].strftime("%d-%b-%Y")
        except Exception:
            pass

    # UAN and ESI display
    uan_display = str(emp.get("uan_no", "N.A.")).strip()
    esi_display = str(emp.get("esi_no", "N.A.")).strip()
    if not uan_display or uan_display in ("N/A", "None", "0", "0.0"):
        uan_display = "N.A."
    if not esi_display or esi_display in ("N/A", "None", "0", "0.0"):
        esi_display = "N.A."

    info_data = [
        [
            Paragraph("<b>Employee Name</b>", style_label),
            Paragraph(emp["emp_name"], style_value),
            Paragraph("<b>Employee ID</b>", style_label),
            Paragraph(str(emp["sl_no"]), style_value),
        ],
        [
            Paragraph("<b>Position</b>", style_label),
            Paragraph(emp["position"], style_value),
            Paragraph("<b>Department</b>", style_label),
            Paragraph(emp["department"], style_value),
        ],
        [
            Paragraph("<b>Date of Joining</b>", style_label),
            Paragraph(doj_display, style_value),
            Paragraph("<b>Mode of Pay</b>", style_label),
            Paragraph(emp.get("mode_of_pay", "NEFT"), style_value),
        ],
        [
            Paragraph("<b>UAN No</b>", style_label),
            Paragraph(uan_display, style_value),
            Paragraph("<b>ESI IP No</b>", style_label),
            Paragraph(esi_display, style_value),
        ],
        [
            Paragraph("<b>PF No</b>", style_label),
            Paragraph(emp.get("pf_no", "N/A"), style_value),
            Paragraph("<b>Bank A/C No</b>", style_label),
            Paragraph(emp.get("bank_account_no", "N/A"), style_value),
        ],
    ]

    col_widths = [90, 130, 90, 130]
    info_table = Table(info_data, colWidths=col_widths)
    info_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), WHITE),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#bdbdbd")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
    ]))
    elements.append(info_table)
    elements.append(Spacer(1, 4 * mm))

    # ---- Attendance Table ----
    elements.append(Paragraph("ATTENDANCE SUMMARY", style_section))
    attend_data = [
        [
            Paragraph("<b>Calendar Days</b>", style_label),
            Paragraph(_fmt_int(emp["calendar_days"]), style_value),
            Paragraph("<b>Present Days</b>", style_label),
            Paragraph(_fmt_int(emp["present_days"]), style_value),
        ],
        [
            Paragraph("<b>Holidays</b>", style_label),
            Paragraph(_fmt_int(emp["holidays"]), style_value),
            Paragraph("<b>Leaves (SL/EL/CL)</b>", style_label),
            Paragraph(_fmt_int(emp["leaves"]), style_value),
        ],
        [
            Paragraph("<b>LOP / Absent</b>", style_label),
            Paragraph(_fmt_int(emp["lop_absent"]), style_value),
            Paragraph("<b>Paid Days</b>", style_label),
            Paragraph(_fmt_int(emp["paid_days"]), style_value),
        ],
    ]
    attend_table = Table(attend_data, colWidths=col_widths)
    attend_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), WHITE),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#bdbdbd")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
    ]))
    elements.append(attend_table)
    elements.append(Spacer(1, 4 * mm))

    # ---- Earnings & Deductions side by side ----
    elements.append(Paragraph("EARNINGS & DEDUCTIONS", style_section))

    # Build earnings rows
    earnings_items = [
        ("Basic", emp["basic"]),
        ("HRA", emp["hra"]),
        ("Conveyance Allowance", emp["conveyance"]),
    ]
    # Include special / medical allowance depending on which has value
    if emp.get("special_allowance", 0) > 0:
        earnings_items.append(("Special Allowance", emp["special_allowance"]))
    if emp.get("medical_allowance", 0) > 0:
        earnings_items.append(("Medical Allowance", emp["medical_allowance"]))
    if emp.get("incentive", 0) > 0:
        earnings_items.append(("Incentive", emp["incentive"]))
    if emp.get("overtime", 0) > 0:
        earnings_items.append(("Over Time", emp["overtime"]))
    if emp.get("arrears", 0) > 0:
        earnings_items.append(("Arrears", emp["arrears"]))
    if emp.get("bonus", 0) > 0:
        earnings_items.append(("Bonus", emp["bonus"]))

    total_earnings = emp.get("total_eff_gross", 0) or sum(v for _, v in earnings_items)

    # Build deductions rows
    deductions_items = []
    if emp.get("employee_pf", 0) > 0:
        deductions_items.append(("Employee PF", emp["employee_pf"]))
    if emp.get("employee_esi", 0) > 0:
        deductions_items.append(("Employee ESI", emp["employee_esi"]))
    if emp.get("pt", 0) > 0:
        deductions_items.append(("Professional Tax", emp["pt"]))
    if emp.get("tds", 0) > 0:
        deductions_items.append(("TDS", emp["tds"]))
    if emp.get("salary_advance", 0) > 0:
        deductions_items.append(("Salary Advance", emp["salary_advance"]))
    if emp.get("lop_amount", 0) > 0:
        deductions_items.append(("LOP Deduction", emp["lop_amount"]))

    total_deductions = emp.get("total_deductions", 0) or sum(v for _, v in deductions_items)

    # Make both lists equal length
    max_rows = max(len(earnings_items), len(deductions_items), 1)
    while len(earnings_items) < max_rows:
        earnings_items.append(("", 0))
    while len(deductions_items) < max_rows:
        deductions_items.append(("", 0))

    # Header row
    ed_data = [[
        Paragraph("<b>Earnings</b>", style_label),
        Paragraph("<b>Amount (₹)</b>", style_label),
        Paragraph("<b>Deductions</b>", style_label),
        Paragraph("<b>Amount (₹)</b>", style_label),
    ]]

    for i in range(max_rows):
        e_label, e_val = earnings_items[i]
        d_label, d_val = deductions_items[i]
        ed_data.append([
            Paragraph(e_label, style_label),
            Paragraph(_fmt(e_val) if e_label else "", style_value),
            Paragraph(d_label, style_label),
            Paragraph(_fmt(d_val) if d_label else "", style_value),
        ])

    # Totals row
    ed_data.append([
        Paragraph("<b>Total Earnings</b>", style_label),
        Paragraph(f"<b>{_fmt(total_earnings)}</b>", style_value),
        Paragraph("<b>Total Deductions</b>", style_label),
        Paragraph(f"<b>{_fmt(total_deductions)}</b>", style_value),
    ])

    ed_col_widths = [120, 100, 120, 100]
    ed_table = Table(ed_data, colWidths=ed_col_widths)
    ed_style = [
        ("BACKGROUND", (0, 1), (-1, -2), WHITE),     # plain white for data rows
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#bdbdbd")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        # Header row (colored)
        ("BACKGROUND", (0, 0), (-1, 0), BRAND_ACCENT),
        ("TEXTCOLOR", (0, 0), (-1, 0), WHITE),
        # Totals row (colored)
        ("BACKGROUND", (0, -1), (-1, -1), HEADER_BG),
        ("LINEABOVE", (0, -1), (-1, -1), 1.5, BRAND_DARK),
        # Vertical separator between earnings & deductions
        ("LINEAFTER", (1, 0), (1, -1), 1.5, BRAND_DARK),
    ]

    ed_table.setStyle(TableStyle(ed_style))
    elements.append(ed_table)
    elements.append(Spacer(1, 4 * mm))

    # ---- Employer Contributions (info only) ----
    if emp.get("employer_pf", 0) > 0 or emp.get("employer_esi", 0) > 0:
        elements.append(Paragraph("EMPLOYER CONTRIBUTIONS (For Information)", style_section))
        contrib_data = [[
            Paragraph("<b>Employer PF</b>", style_label),
            Paragraph(_fmt(emp.get("employer_pf", 0)), style_value),
            Paragraph("<b>Employer ESI</b>", style_label),
            Paragraph(_fmt(emp.get("employer_esi", 0)), style_value),
        ]]
        contrib_table = Table(contrib_data, colWidths=col_widths)
        contrib_table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), WHITE),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#bdbdbd")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ]))
        elements.append(contrib_table)
        elements.append(Spacer(1, 4 * mm))

    # ---- Net Pay Box ----
    net_salary = emp.get("net_salary", 0)
    net_words = _amount_in_words(net_salary)

    net_data = [
        [
            Paragraph("<b>NET PAY</b>", ParagraphStyle(
                "NetLabel", parent=style_label, fontSize=12, textColor=BRAND_DARK,
                fontName=FONT_BOLD,
            )),
            Paragraph(_fmt(net_salary), style_net),
        ],
        [
            Paragraph(f"<i>{net_words}</i>", ParagraphStyle(
                "NetWords", parent=style_footer, fontSize=8, alignment=TA_LEFT,
                textColor=colors.HexColor("#424242"),
            )),
            Paragraph("", style_label),
        ],
    ]
    net_table = Table(net_data, colWidths=[280, 160])
    net_table.setStyle(TableStyle([
        ("BOX", (0, 0), (-1, -1), 2, GREEN_DARK),
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#e8f5e9")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("SPAN", (0, 1), (-1, 1)),
    ]))
    elements.append(net_table)
    elements.append(Spacer(1, 8 * mm))

    # ---- CTC Info ----
    if emp.get("ctc", 0) > 0:
        elements.append(
            Paragraph(
                f"<b>CTC (Cost to Company):</b> {_fmt(emp['ctc'])}",
                ParagraphStyle("CTC", parent=style_label, fontSize=9),
            )
        )
        elements.append(Spacer(1, 4 * mm))

    # ---- Footer ----
    elements.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor("#bdbdbd"), spaceAfter=4))
    elements.append(Paragraph(
        "This is a system-generated payslip and does not require a signature. "
        "For any queries, please contact the HR/Accounts department.",
        style_footer,
    ))
    elements.append(Paragraph(
        f"Generated for {company} | {month}",
        style_footer,
    ))

    doc.build(elements)
    return filepath
