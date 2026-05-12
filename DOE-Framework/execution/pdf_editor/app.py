"""
PDF Editor & Signature Attestation Web Application  (v2 – bug fixes)
=====================================================================
Flask-based web app: edit PDFs, add text/highlights/images, attest signatures.

Usage:
    python app.py                # http://127.0.0.1:5500
    python app.py --port 8080
    python app.py --debug
"""

import os, io, json, uuid, base64, argparse
from datetime import datetime
from pathlib import Path

from flask import (
    Flask, render_template, request, jsonify, send_file,
    redirect, url_for,
)
from werkzeug.utils import secure_filename
import fitz  # PyMuPDF
from PIL import Image

# ---------------------------------------------------------------------------
BASE_DIR  = Path(__file__).resolve().parent
UPLOAD_FOLDER    = BASE_DIR / "uploads"
SIGNATURE_FOLDER = BASE_DIR / "signatures"
IMAGE_FOLDER     = BASE_DIR / "images"
ALLOWED_PDF_EXT  = {"pdf"}
ALLOWED_IMG_EXT  = {"png", "jpg", "jpeg", "bmp", "gif"}

app = Flask(__name__, static_folder="static", template_folder="templates")
app.secret_key = os.environ.get("FLASK_SECRET", "pdf-editor-secret-key-change-me")
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB

for d in [UPLOAD_FOLDER, SIGNATURE_FOLDER, IMAGE_FOLDER]:
    d.mkdir(exist_ok=True)

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def allowed_file(fname, exts):
    return "." in fname and fname.rsplit(".", 1)[1].lower() in exts

def get_pdf_path(file_id):
    return UPLOAD_FOLDER / f"{file_id}.pdf"

def render_page_image(pdf_path, page_num, zoom=2.0):
    doc = fitz.open(str(pdf_path))
    if page_num < 0 or page_num >= len(doc):
        doc.close()
        return None
    page = doc[page_num]
    pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
    img_bytes = pix.tobytes("png")
    doc.close()
    return img_bytes

def _mime_for(ext):
    return {"png": "image/png", "jpg": "image/jpeg", "jpeg": "image/jpeg",
            "bmp": "image/bmp", "gif": "image/gif"}.get(ext, "application/octet-stream")

# ---------------------------------------------------------------------------
# Page routes
# ---------------------------------------------------------------------------
@app.route("/")
def index():
    return render_template("index.html")

@app.route("/editor/<file_id>")
def editor(file_id):
    pdf_path = get_pdf_path(file_id)
    if not pdf_path.exists():
        return redirect(url_for("index"))
    doc = fitz.open(str(pdf_path))
    total = len(doc); doc.close()
    return render_template("editor.html", file_id=file_id, total_pages=total)

# ---------------------------------------------------------------------------
# API – PDF Upload
# ---------------------------------------------------------------------------
@app.route("/api/upload", methods=["POST"])
def api_upload():
    if "file" not in request.files:
        return jsonify({"error": "No file provided"}), 400
    f = request.files["file"]
    if not f.filename or not allowed_file(f.filename, ALLOWED_PDF_EXT):
        return jsonify({"error": "Only PDF files are allowed"}), 400
    file_id = str(uuid.uuid4())[:12]
    dest = get_pdf_path(file_id)
    f.save(str(dest))
    doc = fitz.open(str(dest))
    total = len(doc); doc.close()
    return jsonify({"file_id": file_id, "filename": secure_filename(f.filename), "total_pages": total})

# ---------------------------------------------------------------------------
# API – Page rendering
# ---------------------------------------------------------------------------
@app.route("/api/page/<file_id>/<int:page_num>")
def api_page_image(file_id, page_num):
    pdf_path = get_pdf_path(file_id)
    if not pdf_path.exists():
        return jsonify({"error": "File not found"}), 404
    zoom = float(request.args.get("zoom", 2.0))
    img = render_page_image(pdf_path, page_num, zoom)
    if img is None:
        return jsonify({"error": "Invalid page"}), 400
    return send_file(io.BytesIO(img), mimetype="image/png")

@app.route("/api/page-info/<file_id>/<int:page_num>")
def api_page_info(file_id, page_num):
    pdf_path = get_pdf_path(file_id)
    if not pdf_path.exists():
        return jsonify({"error": "File not found"}), 404
    doc = fitz.open(str(pdf_path))
    if page_num < 0 or page_num >= len(doc):
        doc.close(); return jsonify({"error": "Invalid page"}), 400
    p = doc[page_num]
    info = {"width": p.rect.width, "height": p.rect.height}
    doc.close()
    return jsonify(info)

# ---------------------------------------------------------------------------
# API – Signatures
# ---------------------------------------------------------------------------
@app.route("/api/save-signature", methods=["POST"])
def api_save_signature():
    sig_id = str(uuid.uuid4())[:10]
    if "image" in request.files:
        f = request.files["image"]
        ext = f.filename.rsplit(".", 1)[1].lower() if "." in f.filename else "png"
        if ext not in ALLOWED_IMG_EXT:
            ext = "png"
        sig_path = SIGNATURE_FOLDER / f"{sig_id}.{ext}"
        f.save(str(sig_path))
    elif request.is_json and "dataUrl" in request.json:
        data_url = request.json["dataUrl"]
        header, encoded = data_url.split(",", 1)
        img_bytes = base64.b64decode(encoded)
        sig_path = SIGNATURE_FOLDER / f"{sig_id}.png"
        sig_path.write_bytes(img_bytes)
    else:
        return jsonify({"error": "No signature data provided"}), 400
    return jsonify({"signature_id": sig_id, "path": sig_path.name})

@app.route("/api/signatures")
def api_list_signatures():
    sigs = []
    for f in sorted(SIGNATURE_FOLDER.iterdir(), key=lambda p: p.stat().st_mtime, reverse=True):
        if f.suffix.lower().lstrip(".") in ALLOWED_IMG_EXT:
            sigs.append({"id": f.stem, "filename": f.name,
                         "url": f"/api/signature-image/{f.name}"})
    return jsonify(sigs)

@app.route("/api/signature-image/<filename>")
def api_signature_image(filename):
    sig_path = SIGNATURE_FOLDER / secure_filename(filename)
    if not sig_path.exists():
        return jsonify({"error": "Not found"}), 404
    ext = sig_path.suffix.lower().lstrip(".")
    return send_file(str(sig_path), mimetype=_mime_for(ext))

@app.route("/api/delete-signature/<sig_id>", methods=["DELETE"])
def api_delete_signature(sig_id):
    for f in SIGNATURE_FOLDER.iterdir():
        if f.stem == sig_id:
            f.unlink()
            return jsonify({"status": "deleted"})
    return jsonify({"error": "Not found"}), 404

# ---------------------------------------------------------------------------
# API – Images (new: upload/list/serve/delete overlay images)
# ---------------------------------------------------------------------------
@app.route("/api/upload-image", methods=["POST"])
def api_upload_image():
    if "image" not in request.files:
        return jsonify({"error": "No image provided"}), 400
    f = request.files["image"]
    if not f.filename:
        return jsonify({"error": "Empty filename"}), 400
    ext = f.filename.rsplit(".", 1)[1].lower() if "." in f.filename else ""
    if ext not in ALLOWED_IMG_EXT:
        return jsonify({"error": f"Unsupported format. Allowed: {', '.join(ALLOWED_IMG_EXT)}"}), 400
    img_id = str(uuid.uuid4())[:10]
    img_path = IMAGE_FOLDER / f"{img_id}.{ext}"
    f.save(str(img_path))
    return jsonify({
        "image_id": img_id,
        "filename": secure_filename(f.filename),
        "url": f"/api/image-file/{img_path.name}",
    })

@app.route("/api/images")
def api_list_images():
    imgs = []
    for f in sorted(IMAGE_FOLDER.iterdir(), key=lambda p: p.stat().st_mtime, reverse=True):
        if f.suffix.lower().lstrip(".") in ALLOWED_IMG_EXT:
            imgs.append({"id": f.stem, "filename": f.name,
                         "url": f"/api/image-file/{f.name}"})
    return jsonify(imgs)

@app.route("/api/image-file/<filename>")
def api_image_file(filename):
    img_path = IMAGE_FOLDER / secure_filename(filename)
    if not img_path.exists():
        return jsonify({"error": "Not found"}), 404
    ext = img_path.suffix.lower().lstrip(".")
    return send_file(str(img_path), mimetype=_mime_for(ext))

@app.route("/api/delete-image/<img_id>", methods=["DELETE"])
def api_delete_image(img_id):
    for f in IMAGE_FOLDER.iterdir():
        if f.stem == img_id:
            f.unlink()
            return jsonify({"status": "deleted"})
    return jsonify({"error": "Not found"}), 404

# ---------------------------------------------------------------------------
# API – Apply edits & download
# ---------------------------------------------------------------------------
@app.route("/api/apply-edits", methods=["POST"])
def api_apply_edits():
    data = request.get_json()
    if not data or "file_id" not in data:
        return jsonify({"error": "Missing file_id"}), 400
    file_id = data["file_id"]
    edits   = data.get("edits", [])
    pdf_path = get_pdf_path(file_id)
    if not pdf_path.exists():
        return jsonify({"error": "File not found"}), 404

    doc = fitz.open(str(pdf_path))

    for edit in edits:
        page_num = edit.get("page", 0)
        if page_num < 0 or page_num >= len(doc):
            continue
        page = doc[page_num]
        etype = edit.get("type", "")

        if etype in ("text", "date_stamp"):
            x = float(edit.get("x", 0))
            y = float(edit.get("y", 0))
            text = edit.get("text", "")
            fs   = float(edit.get("fontSize", 12))
            color = edit.get("color", [0, 0, 0])
            color = tuple(c if c <= 1 else c / 255.0 for c in color)
            tw = fitz.TextWriter(page.rect)
            font = fitz.Font("helv")
            tw.append((x, y + fs), text, font=font, fontsize=fs)
            tw.write_text(page, color=color)

        elif etype == "signature":
            x, y = float(edit.get("x", 0)), float(edit.get("y", 0))
            w, h = float(edit.get("width", 150)), float(edit.get("height", 50))
            sig_data = edit.get("signatureData", "")
            img_bytes = _resolve_image_bytes(sig_data, SIGNATURE_FOLDER)
            if img_bytes:
                page.insert_image(fitz.Rect(x, y, x + w, y + h), stream=img_bytes)

        elif etype == "image":
            x, y = float(edit.get("x", 0)), float(edit.get("y", 0))
            w, h = float(edit.get("width", 150)), float(edit.get("height", 100))
            img_data = edit.get("imageData", "")
            img_bytes = _resolve_image_bytes(img_data, IMAGE_FOLDER)
            if img_bytes:
                page.insert_image(fitz.Rect(x, y, x + w, y + h), stream=img_bytes)

        elif etype == "highlight":
            x, y = float(edit.get("x", 0)), float(edit.get("y", 0))
            w, h = float(edit.get("width", 200)), float(edit.get("height", 20))
            color = edit.get("color", [1, 1, 0])
            color = tuple(c if c <= 1 else c / 255.0 for c in color)
            opacity = float(edit.get("opacity", 0.35))
            annot = page.add_rect_annot(fitz.Rect(x, y, x + w, y + h))
            annot.set_colors(fill=color)
            annot.set_opacity(opacity)
            annot.update()

        elif etype == "rectangle":
            x, y = float(edit.get("x", 0)), float(edit.get("y", 0))
            w, h = float(edit.get("width", 100)), float(edit.get("height", 50))
            color = edit.get("color", [1, 0, 0])
            color = tuple(c if c <= 1 else c / 255.0 for c in color)
            lw = float(edit.get("lineWidth", 1.5))
            shape = page.new_shape()
            shape.draw_rect(fitz.Rect(x, y, x + w, y + h))
            shape.finish(color=color, width=lw)
            shape.commit()

    output_id = str(uuid.uuid4())[:12]
    out_path  = UPLOAD_FOLDER / f"{output_id}_edited.pdf"
    doc.save(str(out_path))
    doc.close()
    return jsonify({
        "download_url": f"/api/download/{output_id}_edited",
        "output_id": f"{output_id}_edited",
    })


def _resolve_image_bytes(ref, folder):
    """Resolve an image reference (base64 data-URL or file-stem ID) to bytes."""
    if not ref:
        return None
    if ref.startswith("data:"):
        _, encoded = ref.split(",", 1)
        return base64.b64decode(encoded)
    # Treat as file-stem ID
    for f in folder.iterdir():
        if f.stem == ref:
            return f.read_bytes()
    return None


@app.route("/api/download/<output_id>")
def api_download(output_id):
    safe = secure_filename(output_id)
    path = UPLOAD_FOLDER / f"{safe}.pdf"
    if not path.exists():
        return jsonify({"error": "File not found"}), 404
    return send_file(str(path), as_attachment=True,
                     download_name=f"edited_{safe}.pdf", mimetype="application/pdf")

@app.route("/api/extract-text/<file_id>/<int:page_num>")
def api_extract_text(file_id, page_num):
    pdf_path = get_pdf_path(file_id)
    if not pdf_path.exists():
        return jsonify({"error": "File not found"}), 404
    doc = fitz.open(str(pdf_path))
    if page_num < 0 or page_num >= len(doc):
        doc.close(); return jsonify({"error": "Invalid page"}), 400
    text = doc[page_num].get_text()
    doc.close()
    return jsonify({"text": text})

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="PDF Editor & Signature Attestation")
    parser.add_argument("--port",  type=int, default=5500)
    parser.add_argument("--host",  type=str, default="127.0.0.1")
    parser.add_argument("--debug", action="store_true")
    args = parser.parse_args()
    print(f"\n{'='*60}")
    print(f"  PDF Editor & Signature Attestation  v2")
    print(f"  http://{args.host}:{args.port}")
    print(f"{'='*60}\n")
    app.run(host=args.host, port=args.port, debug=args.debug)
