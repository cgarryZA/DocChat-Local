#!/usr/bin/env python3
"""
Walks ./Docs, converts documents to Markdown, and writes to ./data/raw
preserving relative paths.

- Non-PDFs (docx, doc, txt, html, htm, rtf, odt, md): Pandoc (or copy for .md)
- PDFs: PyMuPDF text extraction with OCR fallback via Tesseract (optional)

Usage examples:
  python run_total_convert.py
  python run_total_convert.py --lang eng+de --dpi 250 --force
  python run_total_convert.py --no-ocr

Requirements:
  - Pandoc installed and on PATH (for non-PDF conversions)
  - pip install: pymupdf markdownify pillow pytesseract (for PDF OCR path)
  - (Optional) Tesseract binary installed; if not on PATH, use --tesseract PATH

Project structure (assumed):
  ./Docs/         <-- source files (input)
  ./data/raw/     <-- destination (output)
"""

import argparse
import os
import shutil
import subprocess
import sys
from pathlib import Path
from typing import Optional
import re

# --- Project paths (relative to this script) ---
ROOT = Path(__file__).resolve().parent
DOCS_DIR = ROOT / "Docs"
RAW_DIR = ROOT / "data" / "raw"

# --- Pandoc config ---
PANDOC_OUTPUT_FORMAT = "gfm"  # GitHub-flavored Markdown

# --- PDF conversion dependencies ---
IMG_MD_REGEX = re.compile(r'!\[[^\]]*\]\(data:image/[^)]+\)', re.IGNORECASE)

try:
    import fitz  # PyMuPDF
    from markdownify import markdownify as to_md
    HAVE_PYMUPDF = True
except Exception:
    HAVE_PYMUPDF = False
    to_md = None  # type: ignore

# OCR (optional)
try:
    import pytesseract
    from PIL import Image
    HAVE_TESSERACT_PY = True
except Exception:
    HAVE_TESSERACT_PY = False


def strip_inline_images(md: str) -> str:
    # Remove base64 inline images from markdownified XHTML
    md = IMG_MD_REGEX.sub("", md)
    md = re.sub(r'[ \t]+', ' ', md)
    md = re.sub(r'\n{3,}', '\n\n', md)
    return md.strip()


def check_pandoc_available() -> bool:
    try:
        subprocess.run(["pandoc", "--version"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=True)
        return True
    except Exception:
        return False


def convert_with_pandoc(src: Path, dst: Path, *, force: bool) -> bool:
    """Convert non-PDF into Markdown using pandoc. Returns True if written."""
    if dst.exists() and not force:
        print(f"  [skip] exists: {dst}")
        return False

    dst.parent.mkdir(parents=True, exist_ok=True)
    cmd = ["pandoc", str(src), "-t", PANDOC_OUTPUT_FORMAT, "-o", str(dst)]
    try:
        subprocess.run(cmd, check=True)
        print(f"  [ok] pandoc: {src} -> {dst}")
        return True
    except subprocess.CalledProcessError as e:
        print(f"  [err] pandoc failed for {src}: {e}")
        return False


def copy_markdown(src: Path, dst: Path, *, force: bool) -> bool:
    if dst.exists() and not force:
        print(f"  [skip] exists: {dst}")
        return False
    dst.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(src, dst)
    print(f"  [ok] copy: {src} -> {dst}")
    return True


def page_to_markdown_or_text(page) -> str:
    """Try PyMuPDF native markdown -> xhtml->markdown (strip <img>) -> plain text."""
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
            md = to_md(html, heading_style="ATX") if to_md else ""
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


def ocr_page_to_text(page, dpi: int, lang: str) -> str:
    """Render a page to an image and OCR it with pytesseract."""
    mat = fitz.Matrix(dpi / 72.0, dpi / 72.0)
    pix = page.get_pixmap(matrix=mat, alpha=False)
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    txt = pytesseract.image_to_string(img, lang=lang)
    return txt or ""


def convert_pdf_to_markdown(src: Path, dst: Path, *, lang: str, dpi: int, min_text_chars: int,
                            ocr_enabled: bool, tesseract_cmd: Optional[str], force: bool) -> bool:
    """PDF -> Markdown with OCR fallback. Returns True if written."""
    if not HAVE_PYMUPDF:
        print("  [err] PyMuPDF / markdownify not installed. pip install pymupdf markdownify", file=sys.stderr)
        return False

    if dst.exists() and not force:
        print(f"  [skip] exists: {dst}")
        return False

    if ocr_enabled and HAVE_TESSERACT_PY and tesseract_cmd:
        if not Path(tesseract_cmd).exists():
            print(f"  [warn] tesseract path not found: {tesseract_cmd} -> continuing without OCR", file=sys.stderr)
            ocr_enabled = False
        else:
            pytesseract.pytesseract.tesseract_cmd = tesseract_cmd  # type: ignore

    total_chars = 0
    parts = []

    try:
        doc = fitz.open(str(src))
    except Exception as e:
        print(f"  [err] cannot open PDF: {src} ({e})")
        return False

    with doc:
        for i, page in enumerate(doc, start=1):
            md = page_to_markdown_or_text(page)
            if ocr_enabled and HAVE_TESSERACT_PY and len(md.strip()) < min_text_chars:
                try:
                    ocr_txt = ocr_page_to_text(page, dpi=dpi, lang=lang)
                    if len(ocr_txt.strip()) > len(md.strip()):
                        md = ocr_txt
                except Exception as e:
                    print(f"  [warn] OCR failed on page {i}: {e}", file=sys.stderr)

            total_chars += len(md)
            parts.append(f"<!-- page:{i} -->\n\n{md}")

    out = strip_inline_images("\n\n".join(parts))
    dst.parent.mkdir(parents=True, exist_ok=True)
    dst.write_text(out, encoding="utf-8")

    if total_chars < min_text_chars:
        print("  [warn] Very little text extracted; PDF may be mostly images. Consider OCR / higher DPI.")

    print(f"  [ok] pdf: {src} -> {dst}  ({total_chars} chars)")
    return True


def build_argparser() -> argparse.ArgumentParser:
    ap = argparse.ArgumentParser(description="Convert Docs/ -> data/raw/ preserving relative paths.")
    ap.add_argument("--lang", default="eng", help="Tesseract languages (e.g., 'eng', or 'eng+de'). Default: eng")
    ap.add_argument("--dpi", type=int, default=200, help="OCR render DPI (default: 200)")
    ap.add_argument("--min-text", type=int, default=120, help="Min chars per page to skip OCR (default: 120)")
    ap.add_argument("--tesseract", dest="tesseract_cmd", default=None, help="Path to tesseract.exe if not on PATH")
    ap.add_argument("--no-ocr", action="store_true", help="Disable OCR fallback for PDFs")
    ap.add_argument("--force", action="store_true", default=False, help=argparse.SUPPRESS)  # backward compat
    ap.add_argument("--overwrite", action="store_true", help="Overwrite existing outputs")
    ap.add_argument("--include-md", action="store_true", help="Also copy .md files from Docs to data/raw")
    return ap


def main():
    args = build_argparser().parse_args()
    ocr_enabled = not args.no_ocr

    RAW_DIR.mkdir(parents=True, exist_ok=True)

    # Supported by pandoc (input side) that we’ll route through pandoc:
    pandoc_exts = {".docx", ".doc", ".txt", ".html", ".htm", ".rtf", ".odt", ".pptx", ".ppt"}
    # We’ll optionally copy .md as-is (if --include-md)
    md_ext = ".md"

    have_pandoc = check_pandoc_available()
    if not have_pandoc:
        print("ℹ️  pandoc not found on PATH. Non-PDFs will be skipped.", file=sys.stderr)

    # Walk Docs/
    if not DOCS_DIR.exists():
        print(f"[err] Docs directory not found: {DOCS_DIR}", file=sys.stderr)
        sys.exit(1)

    converted = 0
    skipped = 0
    errors = 0

    print(f"Scanning: {DOCS_DIR}")
    for src in DOCS_DIR.rglob("*"):
        if not src.is_file():
            continue
        if src.name.startswith("~"):  # skip temp files
            continue

        rel = src.relative_to(DOCS_DIR)  # path under Docs/
        out_rel = rel.with_suffix(".md")
        dst = RAW_DIR / out_rel

        ext = src.suffix.lower()

        try:
            if ext == ".pdf":
                ok = convert_pdf_to_markdown(
                    src, dst,
                    lang=args.lang,
                    dpi=args.dpi,
                    min_text_chars=args.min_text,
                    ocr_enabled=ocr_enabled,
                    tesseract_cmd=args.tesseract_cmd,
                    force=args.overwrite or args.force,
                )
                if ok: converted += 1
                else: skipped += 1

            elif ext in pandoc_exts:
                if not have_pandoc:
                    print(f"  [skip] pandoc missing -> {src}")
                    skipped += 1
                else:
                    ok = convert_with_pandoc(src, dst, force=args.overwrite or args.force)
                    if ok: converted += 1
                    else: skipped += 1

            elif ext == md_ext and args.include_md:
                ok = copy_markdown(src, dst, force=args.overwrite or args.force)
                if ok: converted += 1
                else: skipped += 1

            else:
                # Ignore other file types silently
                continue

        except Exception as e:
            errors += 1
            print(f"  [err] {src}: {e}", file=sys.stderr)

    print("\nSummary:")
    print(f"  Converted: {converted}")
    print(f"  Skipped:   {skipped}")
    print(f"  Errors:    {errors}")
    print(f"Output root: {RAW_DIR}")


if __name__ == "__main__":
    main()
