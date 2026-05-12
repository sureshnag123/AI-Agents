# PDF Editor & Signature Attestation

## Purpose
A local web-based application for editing PDF files and attesting signatures. All processing happens on the user's machine—no data leaves the local network.

## Quick Start
```bash
cd execution/pdf_editor
pip install flask PyMuPDF Pillow reportlab
python app.py                # → http://127.0.0.1:5500
python app.py --port 8080    # custom port
python app.py --debug        # hot-reload
```

## Features

### PDF Viewing
- Upload any PDF (up to 50 MB)
- Page-by-page rendering with zoom (100%–300%)
- Thumbnail navigation sidebar
- Text extraction per page

### Annotations
| Tool | Description |
|------|-------------|
| **Text** | Click anywhere to add editable text. Customize font size, color, weight. |
| **Highlight** | Click to place a translucent highlight box. Drag to reposition, resize with handle. |
| **Rectangle** | Draw bordered rectangles around content. |
| **Signature** | Place a saved signature image at any location. Drag + resize. |
| **Attest** | Composite stamp: signature + name + designation + date + attestation label. |

### Signature Management
- **Draw**: Free-hand drawing pad in the right panel
- **Upload**: Import a PNG/JPG signature image
- **Library**: All signatures are saved to `signatures/` folder and persist across sessions
- **Select → Place**: Click a library signature, then click on the page to stamp it

### Attestation Flow
1. Draw or upload your signature and save to library
2. Select the signature from the library (click it)
3. Fill in **Signatory Name**, **Title/Designation**, **Date**, and **Attestation Text** in the Attestation panel
4. Select the **Attest** tool from the toolbar
5. Click on the PDF page where you want the attestation stamp
6. Reposition by dragging, then **Save & Download**

### Save & Download
- Click **Save & Download** to burn all annotations/signatures into the PDF
- A new PDF is generated server-side using PyMuPDF and downloaded automatically
- Original file is never modified

## Architecture
```
execution/pdf_editor/
├── app.py              # Flask backend (API + page rendering)
├── templates/
│   ├── index.html      # Upload / landing page
│   └── editor.html     # Main editor interface
├── static/
│   └── css/
│       └── style.css   # Full stylesheet
├── uploads/            # Temporary uploaded PDFs
├── signatures/         # Persistent signature library
└── requirements.txt    # Python dependencies
```

## API Endpoints

| Method | Endpoint | Description |
|--------|----------|-------------|
| POST | `/api/upload` | Upload a PDF file |
| GET | `/api/page/<id>/<page>` | Render page as PNG |
| GET | `/api/page-info/<id>/<page>` | Page dimensions |
| POST | `/api/save-signature` | Save signature (base64 or file) |
| GET | `/api/signatures` | List saved signatures |
| GET | `/api/signature-image/<name>` | Serve signature image |
| DELETE | `/api/delete-signature/<id>` | Delete a signature |
| POST | `/api/apply-edits` | Apply all edits, return edited PDF |
| GET | `/api/download/<id>` | Download edited PDF |
| GET | `/api/extract-text/<id>/<page>` | Extract page text |

## Dependencies
- **Flask** – Web framework
- **PyMuPDF (fitz)** – PDF rendering, text writing, image insertion
- **Pillow** – Image processing
- **ReportLab** – PDF generation utilities

## Edge Cases & Notes
- Large PDFs (100+ pages) may be slow to thumbnail—lazy loading mitigates this
- Signature pad uses mouse events + touch events for tablet support
- All overlay positions are stored in PDF coordinate space (72 DPI) and scaled by zoom
- The `uploads/` folder can be periodically cleaned; files are temporary
- `signatures/` folder persists across sessions—back it up if needed

## Deployment Modes
- **Local (Default)**: `python app.py` — runs on localhost
- **Network**: `python app.py --host 0.0.0.0` — accessible on LAN
- **Production**: Use `gunicorn` or `waitress` behind a reverse proxy

## Troubleshooting
- **"fitz not found"**: Install with `pip install PyMuPDF` (not `pip install fitz`)
- **Signature not appearing**: Ensure you clicked "Save to Library" and then selected the signature from the library list before clicking on the page
- **Large file timeout**: Increase `MAX_CONTENT_LENGTH` in `app.py`
