#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Total converter: scans ./Docs and writes Markdown to ./data/raw
- Non-PDFs: prefer Pandoc for DOCX/RTF/HTML. MD/TXT copied as-is.
- PDFs: PyMuPDF text; OCR fallback via Tesseract for image pages.
- .doc handled by: Word COM (pywin32) -> .docx -> Pandoc, else LibreOffice, else instruct user.

Usage examples (from repo root):
  python run_total_convert.py --include-md
  python run_total_convert.py --include-md --no-ocr
  python run_total_convert.py --include-md --tesseract "C:\\Users\\YOU\\AppData\\Local\\Programs\\Tesseract-OCR\\tesseract.exe" --lang eng+de --dpi 250
"""
import argparse, os, sys, shutil, subprocess, tempfile
from pathlib import Path
from typing import Optional, Tuple

# Paths
ROOT = Path(__file__).resolve().parent
DOCS_DIR = ROOT / "Docs"
OUT_DIR = ROOT / "data" / "raw"
OUT_DIR.mkdir(parents=True, exist_ok=True)

# ---------- helpers ----------

def which(cmd: str) -> Optional[str]:
    return shutil.which(cmd)

def run(cmd: list[str]) -> Tuple[int, str, str]:
    p = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    out, err = p.communicate()
    return p.returncode, out, err

def ensure_parent(p: Path):
    p.parent.mkdir(parents=True, exist_ok=True)

def log(msg: str): print(msg, flush=True)
def warn(msg: str): print(f"  [warn] {msg}", flush=True)
def err(msg: str):  print(f"  [err]  {msg}", flush=True)

# ---------- converters ----------

def convert_with_pandoc(src: Path, dst: Path) -> bool:
    """Use Pandoc to convert to GitHub-Flavored Markdown."""
    pandoc = which("pandoc")
    if not pandoc:
        warn("Pandoc not found; skipping non-PDF rich formats. Install from https://pandoc.org/installing")
        return False
    ensure_parent(dst)
    code, out, e = run([pandoc, str(src), "-t", "gfm", "-o", str(dst)])
    if code != 0:
        err(f"pandoc failed for {src}: {e.strip() or out.strip()}")
        return False
    return True

def convert_doc_to_docx_with_word(src: Path) -> Optional[Path]:
    """
    Try Microsoft Word automation (COM) to convert .doc -> .docx.
    Requires: pip install pywin32   and Word installed.
    """
    try:
        import pythoncom  # type: ignore
        import win32com.client  # type: ignore
    except Exception:
        return None

    try:
        pythoncom.CoInitialize()
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(str(src))
        out = src.with_suffix(".docx")
        # 16 = wdFormatDocumentDefault (.docx)
        doc.SaveAs(str(out), FileFormat=16)
        doc.Close(False)
        word.Quit()
        log(f"  [ok] Converted {src.name} -> {out.name} via Microsoft Word")
        return out
    except Exception as e:
        warn(f"Microsoft Word COM failed for {src.name}: {e}")
        try:
            word.Quit()
        except Exception:
            pass
        return None
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

def convert_doc_to_docx_with_soffice(src: Path) -> Optional[Path]:
    """Try LibreOffice headless to turn .doc -> .docx. Returns new path or None."""
    soffice = which("soffice") or which("soffice.exe")
    if not soffice:
        return None
    with tempfile.TemporaryDirectory() as td:
        code, out, e = run([soffice, "--headless", "--convert-to", "docx", "--outdir", td, str(src)])
        if code != 0:
            warn(f"LibreOffice failed to convert {src.name} -> docx: {e.strip() or out.strip()}")
            return None
        out_path = Path(td) / (src.stem + ".docx")
        if out_path.exists():
            final = src.with_suffix(".docx")
            shutil.move(str(out_path), str(final))
            log(f"  [ok] Converted {src.name} -> {final.name} via LibreOffice")
            return final
    return None

def html_to_md_fallback(html: str) -> str:
    """Prefer markdownify; else fallback to plain text using BeautifulSoup."""
    try:
        from markdownify import markdownify as to_md
        return to_md(html)
    except Exception:
        try:
            from bs4 import BeautifulSoup
            return BeautifulSoup(html, "lxml").get_text("\n")
        except Exception:
            return html  # last resort: raw HTML

def convert_html(src: Path, dst: Path) -> bool:
    try:
        html = src.read_text(encoding="utf-8", errors="ignore")
    except Exception as e:
        err(f"Failed to read HTML {src}: {e}")
        return False
    md = html_to_md_fallback(html)
    ensure_parent(dst)
    dst.write_text(md, encoding="utf-8")
    return True

def pdf_to_md_pymupdf(src: Path) -> str:
    """Extract text/markdown per page with PyMuPDF; no OCR here."""
    try:
        import fitz  # PyMuPDF
    except Exception as e:
        raise RuntimeError("PyMuPDF (pymupdf) is required for PDF text parsing.") from e
    text_parts = []
    with fitz.open(str(src)) as doc:
        for page in doc:
            try:
                text_parts.append(page.get_text("text"))
            except Exception:
                text_parts.append(page.get_text())
    return "\n\n".join(text_parts).strip()

def ocr_pdf_with_tesseract(src: Path, lang: str, dpi: int, tesseract_path: Optional[str]) -> str:
    """OCR each page image then concatenate text."""
    try:
        import fitz
        from PIL import Image
        import pytesseract
    except Exception as e:
        raise RuntimeError("OCR requires: pymupdf, pillow, pytesseract") from e
    if tesseract_path:
        pytesseract.pytesseract.tesseract_cmd = tesseract_path

    lines: list[str] = []
    with fitz.open(str(src)) as doc:
        for i, page in enumerate(doc, start=1):
            pix = page.get_pixmap(dpi=dpi)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            try:
                txt = pytesseract.image_to_string(img, lang=lang)
            except Exception as e:
                warn(f"OCR fail on page {i}: {e}")
                txt = ""
            lines.append(txt.strip())
    return "\n\n".join(lines).strip()

def convert_pdf(src: Path, dst: Path, use_ocr: bool, lang: str, dpi: int, tesseract_path: Optional[str]) -> bool:
    """Try text extraction; if text looks empty and OCR allowed, OCR it."""
    try:
        text = pdf_to_md_pymupdf(src)
    except Exception as e:
        warn(f"PDF text pass failed ({e}); trying OCR {'ENABLED' if use_ocr else 'DISABLED'}")
        text = ""

    if use_ocr and len(text.strip()) < 40:
        try:
            text = ocr_pdf_with_tesseract(src, lang, dpi, tesseract_path)
        except Exception as e:
            warn(f"OCR pass failed: {e}")

    text = text.strip()
    if not text:
        warn(f"No extractable text for {src.name}.")
        return False

    ensure_parent(dst)
    dst.write_text(text, encoding="utf-8")
    return True

# ---------- main ----------

def main():
    ap = argparse.ArgumentParser(description="Convert Docs/ -> data/raw as Markdown/text.")
    ap.add_argument("--include-md", action="store_true", help="Also copy .md files to data/raw.")
    ap.add_argument("--no-ocr", action="store_true", help="Disable OCR fallback for PDFs.")
    ap.add_argument("--tesseract", default="", help="Full path to tesseract.exe (optional).")
    ap.add_argument("--lang", default="eng", help="OCR languages, e.g. 'eng+de'.")
    ap.add_argument("--dpi", type=int, default=200, help="OCR rendering DPI (default 200).")
    args = ap.parse_args()

    if not DOCS_DIR.exists():
        log(f"Docs folder not found: {DOCS_DIR}")
        sys.exit(0)

    log(f"Scanning: {DOCS_DIR}")
    converted = 0
    skipped = 0
    errs = 0

    for src in DOCS_DIR.rglob("*"):
        if src.is_dir():
            continue
        rel = src.relative_to(DOCS_DIR)
        out = (OUT_DIR / rel).with_suffix(".md")

        ext = src.suffix.lower()
        try:
            if ext == ".pdf":
                ok = convert_pdf(src, out, use_ocr=not args.no_ocr, lang=args.lang, dpi=args.dpi, tesseract_path=args.tesseract or None)
                if ok:
                    log(f"  [ok] PDF -> {out.relative_to(OUT_DIR)}")
                    converted += 1
                else:
                    skipped += 1

            elif ext in {".md"}:
                if args.include_md:
                    if out.exists() and out.read_text(encoding="utf-8", errors="ignore") == src.read_text(encoding="utf-8", errors="ignore"):
                        log(f"  [skip] exists: {out}")
                        skipped += 1
                    else:
                        ensure_parent(out)
                        shutil.copy2(src, out)
                        log(f"  [ok] Copied MD -> {out.relative_to(OUT_DIR)}")
                        converted += 1
                else:
                    skipped += 1

            elif ext in {".txt"}:
                ensure_parent(out)
                shutil.copy2(src, out)
                log(f"  [ok] Copied TXT -> {out.relative_to(OUT_DIR)}")
                converted += 1

            elif ext in {".html", ".htm"}:
                if convert_html(src, out):
                    log(f"  [ok] HTML -> {out.relative_to(OUT_DIR)}")
                    converted += 1
                else:
                    skipped += 1

            elif ext in {".docx", ".rtf"}:
                if convert_with_pandoc(src, out):
                    log(f"  [ok] Pandoc -> {out.relative_to(OUT_DIR)}")
                    converted += 1
                else:
                    skipped += 1

            elif ext == ".doc":
                # Try Word COM first, then LibreOffice, then instruct user
                new_src = convert_doc_to_docx_with_word(src) or convert_doc_to_docx_with_soffice(src)
                if new_src and convert_with_pandoc(new_src, out):
                    log(f"  [ok] DOC->DOCX->MD -> {out.relative_to(OUT_DIR)}")
                    converted += 1
                else:
                    err("Cannot convert .doc directly. Use Word: Save As .docx, then rerun.")
                    skipped += 1

            else:
                log(f"  [skip] unsupported: {src.name}")
                skipped += 1

        except Exception as e:
            err(f"{src.name}: {e}")
            errs += 1

    log("\nSummary:")
    log(f"  Converted: {converted}")
    log(f"  Skipped:   {skipped}")
    log(f"  Errors:    {errs}")
    log(f"Output root: {OUT_DIR}")

if __name__ == "__main__":
    main()
