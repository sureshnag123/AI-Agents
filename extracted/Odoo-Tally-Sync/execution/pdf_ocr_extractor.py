#!/usr/bin/env python3
"""
PDF OCR Extractor using Claude Vision API

Downloads a PDF from a Google Drive public share link, converts pages to images,
and sends to Claude Vision for structured invoice data extraction.

Usage:
    from pdf_ocr_extractor import extract_invoice_data
    data = extract_invoice_data("https://drive.google.com/file/d/XXX/view")

    # CLI test
    python pdf_ocr_extractor.py "https://drive.google.com/file/d/XXX/view"
"""

import os
import re
import json
import base64
import tempfile
import sys
from pathlib import Path

import requests
import fitz  # PyMuPDF — no external poppler needed
import anthropic
from dotenv import load_dotenv

SCRIPT_DIR = Path(__file__).parent.resolve()
PROJECT_ROOT = SCRIPT_DIR.parent
load_dotenv(PROJECT_ROOT / ".env")

ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY", "")


# ── Google Drive helpers ─────────────────────────────────────────────────────

def extract_drive_file_id(url: str) -> str:
    """Extract file ID from various Google Drive URL formats."""
    # Format: https://drive.google.com/file/d/FILE_ID/view
    m = re.search(r"/file/d/([a-zA-Z0-9_-]+)", url)
    if m:
        return m.group(1)
    # Format: https://drive.google.com/open?id=FILE_ID
    m = re.search(r"[?&]id=([a-zA-Z0-9_-]+)", url)
    if m:
        return m.group(1)
    raise ValueError(f"Cannot extract file ID from Drive URL: {url}")


def download_pdf_from_drive(drive_url: str, dest_path: Path) -> Path:
    """Download a publicly shared Google Drive PDF to dest_path."""
    file_id = extract_drive_file_id(drive_url)
    download_url = f"https://drive.google.com/uc?export=download&id={file_id}"

    session = requests.Session()
    headers = {"User-Agent": "Mozilla/5.0"}
    response = session.get(download_url, headers=headers, stream=True, timeout=60)

    # Google shows a confirmation page for large files — extract confirm token
    content_type = response.headers.get("Content-Type", "")
    if "text/html" in content_type:
        token_match = re.search(r'confirm=([0-9A-Za-z_-]+)', response.text)
        if token_match:
            confirm_url = f"{download_url}&confirm={token_match.group(1)}"
            response = session.get(confirm_url, headers=headers, stream=True, timeout=60)
        else:
            # Newer Google Drive uses a different confirmation mechanism
            uuid_match = re.search(r'uuid=([0-9A-Za-z_-]+)', response.text)
            if uuid_match:
                confirm_url = (
                    f"https://drive.usercontent.google.com/download"
                    f"?id={file_id}&export=download&confirm=t&uuid={uuid_match.group(1)}"
                )
                response = session.get(confirm_url, headers=headers, stream=True, timeout=60)

    response.raise_for_status()

    with open(dest_path, "wb") as f:
        for chunk in response.iter_content(chunk_size=32768):
            if chunk:
                f.write(chunk)

    size_kb = dest_path.stat().st_size / 1024
    print(f"  Downloaded: {size_kb:.1f} KB → {dest_path.name}")
    return dest_path


# ── PDF → images ─────────────────────────────────────────────────────────────

def pdf_to_images_base64(pdf_path: Path, dpi: int = 200) -> list:
    """Convert PDF pages to base64-encoded PNG images for Claude Vision."""
    doc = fitz.open(str(pdf_path))
    images = []
    for page_num in range(len(doc)):
        page = doc[page_num]
        mat = fitz.Matrix(dpi / 72, dpi / 72)
        pix = page.get_pixmap(matrix=mat)
        img_bytes = pix.tobytes("png")
        b64 = base64.standard_b64encode(img_bytes).decode("utf-8")
        images.append({
            "type": "image",
            "source": {
                "type": "base64",
                "media_type": "image/png",
                "data": b64,
            },
        })
    doc.close()
    print(f"  Converted {len(images)} page(s) to images")
    return images


# ── Claude Vision extraction ─────────────────────────────────────────────────

EXTRACTION_PROMPT = """You are an invoice data extraction assistant. Extract the following fields from this purchase invoice/bill and return ONLY a valid JSON object — no explanation, no markdown fences.

{
  "vendor_name": "exact vendor/supplier company name as printed",
  "invoice_number": "invoice or bill number",
  "invoice_date": "YYYY-MM-DD or null",
  "due_date": "YYYY-MM-DD or null",
  "gstin": "vendor GSTIN if present, else null",
  "line_items": [
    {
      "description": "item or service description",
      "quantity": 1.0,
      "unit_price": 0.0,
      "amount": 0.0,
      "tax_rate": 0.0
    }
  ],
  "subtotal": 0.0,
  "tax_rate": 0.0,
  "tax_amount": 0.0,
  "total": 0.0,
  "currency": "INR"
}

Rules:
- All amount fields must be numbers (float), never strings
- invoice_date / due_date must be YYYY-MM-DD or null
- tax_rate is the percentage number (e.g. 18 for 18% GST, 5 for 5% GST)
- If multiple tax rates exist, use the dominant rate for the summary tax_rate
- If a field is not found, use null
- Return ONLY the JSON object"""


def extract_invoice_data(drive_url: str, notes: str = "") -> dict:
    """
    Download PDF from Drive URL and extract invoice data via Claude Vision.

    Returns a dict with all extracted fields plus:
        _pdf_bytes  — raw PDF bytes (for attaching to Odoo bill)
        _drive_url  — original Drive URL
        _notes      — notes from Excel column B
    """
    if not ANTHROPIC_API_KEY:
        raise ValueError(
            "ANTHROPIC_API_KEY not set. Add it to your .env file: ANTHROPIC_API_KEY=sk-ant-..."
        )

    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

    with tempfile.TemporaryDirectory() as tmpdir:
        pdf_path = Path(tmpdir) / "invoice.pdf"

        print(f"  Downloading PDF from Drive...")
        download_pdf_from_drive(drive_url, pdf_path)

        print(f"  Converting PDF to images...")
        images = pdf_to_images_base64(pdf_path)

        if not images:
            raise ValueError("PDF has no pages or could not be rendered")

        pdf_bytes = pdf_path.read_bytes()

    print(f"  Sending to Claude Vision ({len(images)} page(s))...")
    content = images + [{"type": "text", "text": EXTRACTION_PROMPT}]

    message = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=2048,
        messages=[{"role": "user", "content": content}],
    )

    raw = message.content[0].text.strip()

    # Strip markdown code fences if Claude adds them
    raw = re.sub(r"^```(?:json)?\s*", "", raw)
    raw = re.sub(r"\s*```$", "", raw)

    data = json.loads(raw)

    # Ensure numeric fields are floats not strings
    for field in ("subtotal", "tax_rate", "tax_amount", "total"):
        if data.get(field) is not None:
            try:
                data[field] = float(data[field])
            except (ValueError, TypeError):
                data[field] = 0.0

    data["_pdf_bytes"] = pdf_bytes
    data["_drive_url"] = drive_url
    data["_notes"] = notes

    return data


# ── CLI ───────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python pdf_ocr_extractor.py <google_drive_url>")
        sys.exit(1)

    url = sys.argv[1]
    print(f"Extracting invoice data from: {url}\n")
    result = extract_invoice_data(url)

    # Don't print raw PDF bytes
    display = {k: v for k, v in result.items() if not k.startswith("_")}
    print("\nExtracted Data:")
    print(json.dumps(display, indent=2, default=str))
