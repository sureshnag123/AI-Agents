"""Renders one payslip PDF from the same data dict Excel generation uses.

Uses xhtml2pdf (pure-Python, no system/native dependencies) so this works
identically on the developer's Windows machine and on whatever Linux host
this app is eventually deployed to.
"""

import base64
import re
import zipfile
from io import BytesIO
from pathlib import Path

from flask import render_template
from xhtml2pdf import pisa

import generate_payslip  # exposes earnings_deduction_lines(data)


def _logo_data_uri(path):
    if not path:
        return None
    p = Path(path)
    if not p.exists():
        return None
    ext = p.suffix.lstrip(".").lower()
    mime = "png" if ext == "png" else "jpeg"
    b64 = base64.b64encode(p.read_bytes()).decode("ascii")
    return f"data:image/{mime};base64,{b64}"


def render_payslip_pdf(data, logo_file_path=None):
    """Returns PDF bytes for one employee's payslip."""
    earnings, deductions = generate_payslip.earnings_deduction_lines(data)
    html = render_template(
        "payslip_pdf.html",
        data=data,
        earnings=earnings,
        deductions=deductions,
        logo_data_uri=_logo_data_uri(logo_file_path),
    )
    buffer = BytesIO()
    pisa.CreatePDF(src=html, dest=buffer, encoding="utf-8")
    return buffer.getvalue()


def archive_payslips(rows, month_label, archive_folder, logo_file_path=None):
    """Write every employee's PDF (plus one ZIP) into
    <archive_folder>/<month_label>/ on the server's filesystem. Returns
    (archived_count, error) — error is None on success, or a message to
    surface to the user if the folder couldn't be written to."""
    if not archive_folder:
        return 0, None

    safe_month = re.sub(r'[<>:"/\\|?*]', "_", month_label)
    month_dir = Path(archive_folder) / safe_month
    try:
        month_dir.mkdir(parents=True, exist_ok=True)
    except OSError as e:
        return 0, f"Could not write to archive folder '{archive_folder}': {e}"

    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for row_id, data in rows:
            pdf_bytes = render_payslip_pdf(data, logo_file_path)
            filename = re.sub(r'[<>:"/\\|?*]', "_", f"{row_id}.pdf")
            (month_dir / filename).write_bytes(pdf_bytes)
            zf.writestr(filename, pdf_bytes)

    zip_name = re.sub(r'[<>:"/\\|?*]', "_", f"Payslips_{month_label}.zip")
    (month_dir / zip_name).write_bytes(zip_buffer.getvalue())

    return len(rows), None
