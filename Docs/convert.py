# convert.py
# PDF -> Markdown with OCR fallback (PyMuPDF + pytesseract)
# Usage:
#   python convert.py <file.pdf>
#   python convert.py <folder_with_pdfs>
# Options:
#   --lang en+de           # OCR languages (tesseract codes joined by '+')
#   --dpi 200              # OCR render DPI (higher = slower, more accurate)
#   --tesseract "C:\Path\To\tesseract.exe"  # if tesseract not on PATH
#   --min-text 120         # min chars to consider a page "has text" (else OCR)
#   --no-ocr               # disable OCR fallback

import argparse
import os
import pathlib
import re
import sys
from typing import Optional

import fitz  # PyMuPDF
from markdownify import markdownify as to_md

# OCR (optional)
try:
    import pytesseract
    from PIL import Image
    _HAS_TESSERACT = True
except Exception:
    _HAS_TESSERACT = False

# --- Config defaults ---
DEFAULT_LANG = "eng"        # tesseract language code
DEFAULT_DPI = 200           # render DPI for OCR
DEFAULT_MIN_TEXT = 120      # chars threshold to skip OCR

IMG_MD_REGEX = re.compile(r'!\[[^\]]*\]\(data:image/[^)]+\)', re.IGNORECASE)

def clean_md(s: str) -> str:
    # strip base64 images and collapse whitespace
    s = IMG_MD_REGEX.sub("", s)
    s = re.sub(r'[ \t]+', ' ', s)
    s = re.sub(r'\n{3,}', '\n\n', s)
    return s.strip()

def page_to_markdown_or_text(page: fitz.Page) -> str:
    """Try markdown -> xhtml->markdown (strip <img>) -> plain text."""
    # 1) native markdown (some builds)
    try:
        md = page.get_text("markdown")
        if md and md.strip():
            return md
    except Exception:
        pass
    # 2) xhtml -> markdown
    try:
        html = page.get_text("xhtml")
        if html:
            html = re.sub(r'<img[^>]*>', '', html, flags=re.IGNORECASE)
            md = to_md(html, heading_style="ATX")
            if md and md.strip():
                return md
    except Exception:
        pass
    # 3) plain text
    try:
        txt = page.get_text("text")
        return txt or ""
    except Exception:
        return ""

def ocr_page_to_text(page: fitz.Page, dpi: int, lang: str) -> str:
    """Render a page to an image and OCR it with pytesseract."""
    # Render to pixmap at requested DPI
    mat = fitz.Matrix(dpi / 72.0, dpi / 72.0)
    pix = page.get_pixmap(matrix=mat, alpha=False)
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    txt = pytesseract.image_to_string(img, lang=lang)
    return txt or ""

def pdf_to_md(src: pathlib.Path, dst: pathlib.Path, *,
              tesseract_cmd: Optional[str],
              ocr_enabled: bool,
              lang: str,
              dpi: int,
              min_text_chars: int) -> int:
    """
    Convert one PDF to Markdown.
    Returns total extracted character count (rough signal for success).
    """
    # Configure tesseract path if provided / available
    if ocr_enabled and _HAS_TESSERACT:
        if tesseract_cmd:
            pytesseract.pytesseract.tesseract_cmd = tesseract_cmd
        # If user set TESSDATA_PREFIX, respect it; otherwise leave default.

    with fitz.open(str(src)) as doc:
        parts = []
        total_chars = 0

        for i, page in enumerate(doc, start=1):
            md = page_to_markdown_or_text(page)
            # If little/no text and OCR is available, try OCR
            if ocr_enabled and _HAS_TESSERACT and len(md.strip()) < min_text_chars:
                try:
                    ocr_txt = ocr_page_to_text(page, dpi=dpi, lang=lang)
                    # Prefer OCR if it meaningfully adds text
                    if len(ocr_txt.strip()) > len(md.strip()):
                        md = ocr_txt
                except Exception as e:
                    # If OCR fails, keep whatever we had
                    print(f"  [warn] OCR failed on page {i}: {e}", file=sys.stderr)

            total_chars += len(md)
            parts.append(f"<!-- page:{i} -->\n\n{md}")

    out = clean_md("\n\n".join(parts))
    dst.write_text(out, encoding="utf-8")

    # Heuristic warning
    if total_chars < min_text_chars:
        print("⚠️  Very little text extracted — this PDF may be mostly images. "
              "If OCR is disabled or missing, enable it and/or check Tesseract path.",
              file=sys.stderr)

    return total_chars

def parse_args():
    ap = argparse.ArgumentParser(description="PDF -> Markdown with OCR fallback.")
    ap.add_argument("path", help="PDF file or folder containing PDFs")
    ap.add_argument("--lang", default=DEFAULT_LANG,
                    help="Tesseract languages (e.g., 'eng', or 'eng+de'). Default: eng")
    ap.add_argument("--dpi", type=int, default=DEFAULT_DPI, help="OCR render DPI (default: 200)")
    ap.add_argument("--tesseract", dest="tesseract_cmd", default=None,
                    help="Path to tesseract.exe if not on PATH")
    ap.add_argument("--min-text", type=int, default=DEFAULT_MIN_TEXT,
                    help="Min chars on a page to skip OCR (default: 120)")
    ap.add_argument("--no-ocr", action="store_true",
                    help="Disable OCR fallback (text extraction only)")
    return ap.parse_args()

def main():
    args = parse_args()
    p = pathlib.Path(args.path)

    ocr_enabled = not args.no_ocr
    if ocr_enabled and not _HAS_TESSERACT:
        print("ℹ️  pytesseract / PIL not installed; OCR fallback unavailable. "
              "Install with: pip install pillow pytesseract", file=sys.stderr)
        ocr_enabled = False

    # If user provided a tesseract path, ensure it exists
    if args.tesseract_cmd:
        tpath = pathlib.Path(args.tesseract_cmd)
        if not tpath.exists():
            print(f"⚠️  Provided tesseract path does not exist: {tpath}", file=sys.stderr)
            print("    Continuing without OCR fallback.", file=sys.stderr)
            ocr_enabled = False

    if p.is_file() and p.suffix.lower() == ".pdf":
        out = p.with_suffix(".md")
        chars = pdf_to_md(
            p, out,
            tesseract_cmd=args.tesseract_cmd,
            ocr_enabled=ocr_enabled,
            lang=args.lang,
            dpi=args.dpi,
            min_text_chars=args.min_text,
        )
        print(f"→ {out}  ({chars} chars)")
    else:
        # Batch convert all PDFs under this folder (recursive)
        pdfs = list(p.rglob("*.pdf"))
        if not pdfs:
            print("No PDFs found.", file=sys.stderr)
            sys.exit(1)
        for pdf in pdfs:
            out = pdf.with_suffix(".md")
            chars = pdf_to_md(
                pdf, out,
                tesseract_cmd=args.tesseract_cmd,
                ocr_enabled=ocr_enabled,
                lang=args.lang,
                dpi=args.dpi,
                min_text_chars=args.min_text,
            )
            print(f"→ {out}  ({chars} chars)")

if __name__ == "__main__":
    main()
