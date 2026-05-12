#!/usr/bin/env python3
"""
PDF OCR Extractor — Free (Google Gemini 1.5 Flash)

Auto-detect strategy per PDF:
  Digital PDF  → PyMuPDF text extract  → Gemini text API  (fast, cheap tokens)
  Scanned PDF  → PyMuPDF render pages  → Gemini Vision API (image understanding)

Gemini 1.5 Flash free tier (no credit card required):
  • 15 requests / minute
  • 1,500 requests / day
  Get your free key at: https://aistudio.google.com/app/apikey

Usage:
    from pdf_ocr_extractor import extract_invoice_data
    data = extract_invoice_data("https://drive.google.com/file/d/XXX/view")

    # CLI test
    python pdf_ocr_extractor.py "https://drive.google.com/file/d/XXX/view"
"""

import os
import re
import json
import tempfile
import sys
from pathlib import Path

import requests
import fitz  # PyMuPDF — no external poppler needed
import google.generativeai as genai
from dotenv import load_dotenv

SCRIPT_DIR   = Path(__file__).parent.resolve()
PROJECT_ROOT = SCRIPT_DIR.parent
TMP_DIR      = PROJECT_ROOT / ".tmp"
load_dotenv(PROJECT_ROOT / ".env")

GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "")

# PDFs with fewer embedded chars than this are treated as scanned
DIGITAL_TEXT_THRESHOLD = 150


# ── Google Drive helpers ──────────────────────────────────────────────────────

def extract_drive_file_id(url: str) -> str:
    """Extract file ID from various Google Drive URL formats."""
    m = re.search(r"/file/d/([a-zA-Z0-9_-]+)", url)
    if m:
        return m.group(1)
    m = re.search(r"[?&]id=([a-zA-Z0-9_-]+)", url)
    if m:
        return m.group(1)
    raise ValueError(f"Cannot extract file ID from Drive URL: {url}")


def download_pdf_from_drive(drive_url: str, dest_path: Path) -> Path:
    """Download a publicly shared Google Drive PDF to dest_path."""
    file_id     = extract_drive_file_id(drive_url)
    download_url = f"https://drive.google.com/uc?export=download&id={file_id}"

    session  = requests.Session()
    headers  = {"User-Agent": "Mozilla/5.0"}
    response = session.get(download_url, headers=headers, stream=True, timeout=60)

    content_type = response.headers.get("Content-Type", "")
    if "text/html" in content_type:
        # Old confirm token
        token_match = re.search(r'confirm=([0-9A-Za-z_-]+)', response.text)
        if token_match:
            confirm_url = f"{download_url}&confirm={token_match.group(1)}"
            response = session.get(confirm_url, headers=headers, stream=True, timeout=60)
        else:
            # Newer uuid-based confirmation
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


# ── PDF helpers ───────────────────────────────────────────────────────────────

def extract_pdf_text(pdf_path: Path) -> str:
    """Extract embedded text from a digital PDF using PyMuPDF."""
    doc  = fitz.open(str(pdf_path))
    text = "\n".join(page.get_text() for page in doc)
    doc.close()
    return text.strip()


def pdf_to_page_images(pdf_path: Path, dpi: int = 200) -> list:
    """Render each PDF page to raw PNG bytes (for scanned PDFs)."""
    doc    = fitz.open(str(pdf_path))
    images = []
    mat    = fitz.Matrix(dpi / 72, dpi / 72)
    for page in doc:
        pix = page.get_pixmap(matrix=mat)
        images.append(pix.tobytes("png"))
    doc.close()
    print(f"  Rendered {len(images)} page(s) to images ({dpi} dpi)")
    return images


# ── Gemini helpers ────────────────────────────────────────────────────────────

EXTRACTION_PROMPT = """Extract all invoice fields from the content and return ONLY a valid JSON object — no explanation, no markdown fences.

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
- All amount fields must be numbers (float), not strings
- Dates must be YYYY-MM-DD or null
- tax_rate is the percentage number (18 for 18% GST, 5 for 5%)
- If multiple tax rates exist, use the dominant one for the summary tax_rate
- Use null for any field not found
- Return ONLY the JSON object"""


def _get_model() -> genai.GenerativeModel:
    if not GEMINI_API_KEY:
        raise ValueError(
            "GEMINI_API_KEY not set.\n"
            "Get a FREE key (no credit card) at: https://aistudio.google.com/app/apikey\n"
            "Then add to .env:  GEMINI_API_KEY=AIza..."
        )
    genai.configure(api_key=GEMINI_API_KEY)
    return genai.GenerativeModel("gemini-1.5-flash")


def _parse_response(raw: str) -> dict:
    """Strip markdown fences and parse Gemini JSON response."""
    raw = raw.strip()
    raw = re.sub(r"^```(?:json)?\s*", "", raw)
    raw = re.sub(r"\s*```$", "", raw)
    try:
        return json.loads(raw)
    except json.JSONDecodeError as e:
        TMP_DIR.mkdir(parents=True, exist_ok=True)
        raw_file = TMP_DIR / "last_ocr_raw_response.txt"
        raw_file.write_text(raw, encoding="utf-8")
        raise ValueError(
            f"Gemini returned invalid JSON: {e}\n"
            f"Raw response saved to: {raw_file}"
        ) from e


def _extract_digital(text: str) -> dict:
    """Send extracted PDF text to Gemini for structured parsing."""
    model    = _get_model()
    prompt   = EXTRACTION_PROMPT + f"\n\nINVOICE TEXT:\n{text}"
    response = model.generate_content(prompt)
    return _parse_response(response.text)


def _extract_scanned(images: list) -> dict:
    """Send rendered page images to Gemini Vision for OCR + structured parsing."""
    model = _get_model()
    # Build content list: one inline image part per page, then the prompt
    parts = [{"mime_type": "image/png", "data": img_bytes} for img_bytes in images]
    parts.append(EXTRACTION_PROMPT)
    response = model.generate_content(parts)
    return _parse_response(response.text)


# ── Main entry point ──────────────────────────────────────────────────────────

def extract_invoice_data(drive_url: str, notes: str = "") -> dict:
    """
    Download PDF from Drive URL, auto-detect type, extract invoice data.

    Returns extracted fields plus private keys:
        _pdf_bytes  — raw PDF bytes (for attaching to Odoo bill)
        _drive_url  — original Drive URL
        _notes      — notes from Excel column B
    """
    with tempfile.TemporaryDirectory() as tmpdir:
        pdf_path = Path(tmpdir) / "invoice.pdf"

        print("  Downloading PDF from Drive...")
        download_pdf_from_drive(drive_url, pdf_path)

        # Auto-detect: digital or scanned?
        text = extract_pdf_text(pdf_path)

        if len(text) >= DIGITAL_TEXT_THRESHOLD:
            print(f"  Type: digital PDF ({len(text)} chars extracted)")
            print("  Sending text to Gemini...")
            data = _extract_digital(text)
        else:
            print(f"  Type: scanned PDF ({len(text)} chars — too few, using Vision)")
            images = pdf_to_page_images(pdf_path)
            print(f"  Sending {len(images)} page image(s) to Gemini Vision...")
            data = _extract_scanned(images)

        pdf_bytes = pdf_path.read_bytes()

    # Normalize numeric fields
    for field in ("subtotal", "tax_rate", "tax_amount", "total"):
        if data.get(field) is not None:
            try:
                data[field] = float(data[field])
            except (ValueError, TypeError):
                data[field] = 0.0

    data["_pdf_bytes"] = pdf_bytes
    data["_drive_url"] = drive_url
    data["_notes"]     = notes

    return data


# ── CLI ───────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python pdf_ocr_extractor.py <google_drive_url>")
        sys.exit(1)

    url = sys.argv[1]
    print(f"Extracting invoice data from: {url}\n")
    result = extract_invoice_data(url)

    display = {k: v for k, v in result.items() if not k.startswith("_")}
    print("\nExtracted Data:")
    print(json.dumps(display, indent=2, default=str))
