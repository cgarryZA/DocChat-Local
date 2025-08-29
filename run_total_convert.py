#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Total converter: scans ./Docs and writes Markdown to ./data/raw

- Non-PDFs: prefer Pandoc for DOCX/RTF/HTML. MD/TXT copied as-is.
- PDFs: PyMuPDF text; OCR fallback via Tesseract for image pages.
- .doc handled by:
    1) Microsoft Word COM (pywin32) in a helper process (timeout, no dialogs)
    2) LibreOffice headless (soffice) fallback
    3) If both fail, instruct user (but DO NOT delete .doc even with --keep-unsupported)
- Spreadsheets:
    - CSV/TSV -> Markdown table
    - XLSX (and XLS if engine available) -> Markdown with one section per sheet

Flags (defaults are DESTRUCTIVE on success/unsupported):
  --keep-source-on-success       KEEP original when conversion succeeds (default is delete)
  --keep-unsupported             KEEP files with unsupported extensions (default is delete; never deletes .doc)
  --keep-source-if-up-to-date    KEEP source if skipped as up-to-date (default is delete)
  --max-rows-per-sheet N         Limit rows emitted per sheet (CSV/XLSX/XLS). Default 2000.
  --soffice PATH                 Explicit path to soffice(.exe) (overrides auto-detect)
  --no-word                      Skip Word COM for .doc (use only LibreOffice if present)
  --no-prune-empty-dirs          Disable pruning of empty directories (enabled by default)

Usage examples (from repo root):
  python run_total_convert.py --include-md
  python run_total_convert.py --include-md --keep-source-if-up-to-date
  python run_total_convert.py --include-md --keep-source-on-success --keep-unsupported
  python run_total_convert.py --include-md --soffice "C:\\Program Files\\LibreOffice\\program\\soffice.exe"
"""

import argparse, os, sys, shutil, subprocess, tempfile, platform
from pathlib import Path
from typing import Optional, Tuple

# Optional progress bar
try:
    from tqdm import tqdm
except Exception:
    tqdm = None  # fallback to plain loop

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

def is_windows() -> bool:
    return platform.system().lower().startswith("win")

def powershell_unblock(path: Path) -> None:
    """Remove Mark-of-the-Web (Zone.Identifier) without launching PowerShell."""
    if not is_windows():
        return
    try:
        import ctypes
        from ctypes import wintypes
        DeleteFileW = ctypes.windll.kernel32.DeleteFileW
        DeleteFileW.argtypes = [wintypes.LPCWSTR]
        DeleteFileW.restype = wintypes.BOOL
        stream = str(path) + ":Zone.Identifier"
        DeleteFileW(stream)
        long_stream = r"\\?\{}".format(stream)
        DeleteFileW(long_stream)
    except Exception:
        pass

def file_up_to_date(src: Path, dst: Path) -> bool:
    """dst exists and mtime >= src mtime -> up-to-date."""
    if not dst.exists():
        return False
    try:
        return dst.stat().st_mtime >= src.stat().st_mtime
    except Exception:
        return False

# ---- empty folder pruning ----

def _rmdir_if_empty(p: Path, root: Path, enable_log: bool = True) -> bool:
    if not p.exists() or not p.is_dir():
        return False
    try:
        next(p.iterdir())
        return False  # not empty
    except StopIteration:
        pass  # empty
    except Exception:
        return False

    try:
        p.rmdir()
        if enable_log:
            try:
                rel = p.relative_to(root)
            except Exception:
                rel = p
            log(f"  [rmdir] removed empty folder: {rel}")
        return True
    except Exception as e:
        warn(f"Could not remove folder {p}: {e}")
        return False

def prune_empty_dirs_upwards(start: Path, root: Path):
    """Remove empty dirs from 'start' up to (but not including) 'root'."""
    if not start.exists():
        return
    curr = start
    while True:
        if curr == root or not curr.exists():
            break
        if not _rmdir_if_empty(curr, root):
            break
        curr = curr.parent

def prune_all_empty_dirs(root: Path):
    """Bottom-up sweep to remove all empty directories under root."""
    if not root.exists():
        return
    # Walk deepest-first
    for p in sorted([d for d in root.rglob("*") if d.is_dir()], key=lambda x: len(x.parts), reverse=True):
        _rmdir_if_empty(p, root, enable_log=True)

def delete_path_and_prune(path: Path, root: Path):
    """Delete a file then prune empty parent dirs."""
    try:
        path.unlink()
        try:
            rel = path.relative_to(root)
        except Exception:
            rel = path
        log(f"  [del ] {rel}")
    except Exception as e:
        warn(f"Could not delete {path}: {e}")
    prune_empty_dirs_upwards(path.parent, root)

def delete_unsupported(path: Path, do_delete: bool):
    """Delete unsupported/unconvertible *by extension* (NEVER deletes .doc) if requested."""
    if not do_delete:
        return
    delete_path_and_prune(path, DOCS_DIR)

def delete_on_success(path: Path, enabled: bool):
    """Delete a source file after successful conversion (if enabled)."""
    if not enabled:
        return
    delete_path_and_prune(path, DOCS_DIR)

def delete_on_up_to_date(path: Path, enabled: bool):
    """Delete a source file when skipped due to up-to-date output (if enabled)."""
    if not enabled:
        return
    delete_path_and_prune(path, DOCS_DIR)

# ---------- global soffice handling ----------

def find_soffice(explicit: Optional[str]) -> Optional[str]:
    """Resolve soffice path (explicit > PATH > common Windows locations)."""
    if explicit:
        p = Path(explicit)
        return str(p) if p.exists() else None

    w = which("soffice") or which("soffice.exe")
    if w:
        return w

    if is_windows():
        candidates = [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ]
        for c in candidates:
            if Path(c).exists():
                return c

    return None

_SOFFICE_HINT_SHOWN = False
def maybe_show_soffice_hint():
    """Print a one-time hint for adding LibreOffice to PATH."""
    global _SOFFICE_HINT_SHOWN
    if _SOFFICE_HINT_SHOWN or not is_windows():
        return
    _SOFFICE_HINT_SHOWN = True
    log("  [hint] LibreOffice not found. Options:")
    log(r"         • Pass --soffice ""C:\Program Files\LibreOffice\program\soffice.exe""")
    log( "         • Or add to PATH (current session):")
    log(r"             $env:Path += ';C:\Program Files\LibreOffice\program'")
    log( "         • To persist PATH (PowerShell):")
    log(r"             setx PATH ""$($env:Path);C:\Program Files\LibreOffice\program""")

# ---------- converters (general) ----------

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
    # touch dst to src mtime
    try:
        os.utime(dst, (src.stat().st_atime, src.stat().st_mtime))
    except Exception:
        pass
    return code == 0

# ---------- .doc helpers ----------

def convert_doc_to_docx_with_word(src: Path, timeout_sec: int = 60) -> Optional[Path]:
    """Convert .doc -> .docx using Microsoft Word in a helper process (timeout-safe)."""
    if not is_windows():
        return None

    powershell_unblock(src)

    worker_code = r"""
import sys
try:
    import pythoncom
    import win32com.client
except Exception:
    sys.exit(3)

path = sys.argv[1]
out  = path[:-4] + ".docx" if path.lower().endswith(".doc") else path + ".docx"

try:
    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    try: word.DisplayAlerts = 0
    except Exception: pass
    try: word.AutomationSecurity = 3
    except Exception: pass

    doc = word.Documents.Open(
        path,
        ReadOnly=True,
        AddToRecentFiles=False,
        ConfirmConversions=False,
        OpenAndRepair=True
    )
    doc.SaveAs(out, FileFormat=16)  # .docx
    doc.Close(False)
    word.Quit()
    pythoncom.CoUninitialize()
    print(out)
    sys.exit(0)
except Exception:
    try:
        word.Quit()
    except Exception:
        pass
    try:
        pythoncom.CoUninitialize()
    except Exception:
        pass
    sys.exit(2)
"""
    try:
        proc = subprocess.Popen(
            [sys.executable, "-c", worker_code, str(src)],
            stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True
        )
        try:
            out, _ = proc.communicate(timeout=timeout_sec)
        except subprocess.TimeoutExpired:
            proc.kill()
            try:
                subprocess.run(["taskkill", "/IM", "WINWORD.EXE", "/F"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            except Exception:
                pass
            warn(f"Word COM timed out for {src.name} after {timeout_sec}s")
            return None

        if proc.returncode == 0:
            out_path = Path(out.strip())
            if out_path.exists():
                log(f"  [ok] Converted {src.name} -> {out_path.name} via Microsoft Word")
                return out_path
        else:
            return None
    except Exception:
        return None
    return None

def convert_doc_to_docx_with_soffice(src: Path, soffice_path: Optional[str]) -> Optional[Path]:
    """Try LibreOffice headless to turn .doc -> .docx. Returns new path or None."""
    soffice = soffice_path or which("soffice") or which("soffice.exe")
    if not soffice:
        maybe_show_soffice_hint()
        return None
    with tempfile.TemporaryDirectory() as td:
        code, out, e = run([soffice, "--headless", "--convert-to", "docx", "--outdir", td, str(src)])
        if code != 0:
            warn(f"LibreOffice failed to convert {src.name} -> docx: {e.strip() or out.strip()}")
            return None
        out_path = Path(td) / (src.stem + ".docx")
        if out_path.exists():
            final = src.with_suffix(".docx")
            try:
                shutil.move(str(out_path), str(final))
            except Exception:
                shutil.copy2(str(out_path), str(final))
            log(f"  [ok] Converted {src.name} -> {final.name} via LibreOffice")
            return final
    return None

# ---------- HTML ----------

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
            return html  # last resort

def convert_html(src: Path, dst: Path) -> bool:
    try:
        html = src.read_text(encoding="utf-8", errors="ignore")
    except Exception as e:
        err(f"Failed to read HTML {src}: {e}")
        return False
    md = html_to_md_fallback(html)
    ensure_parent(dst)
    dst.write_text(md, encoding="utf-8")
    try:
        os.utime(dst, (src.stat().st_atime, src.stat().st_mtime))
    except Exception:
        pass
    return True

# ---------- PDFs ----------

def pdf_to_md_pymupdf(src: Path) -> str:
    """Extract text per page with PyMuPDF; no OCR here."""
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
    try:
        os.utime(dst, (src.stat().st_atime, src.stat().st_mtime))
    except Exception:
        pass
    return True

# ---------- CSV / TSV ----------

def convert_csv_like(src: Path, dst: Path, max_rows: int) -> bool:
    """Convert CSV/TSV to Markdown table (first row as header if present)."""
    try:
        import pandas as pd
    except Exception:
        warn(f"pandas not installed; cannot convert {src.name}. pip install pandas openpyxl")
        return False

    encodings = ["utf-8", "utf-16", "latin1"]
    df = None
    for enc in encodings:
        try:
            df = pd.read_csv(src, sep=None, engine="python", encoding=enc, on_bad_lines="skip")
            break
        except Exception:
            continue
    if df is None:
        err(f"Failed to read CSV/TSV: {src}")
        return False

    extra_note = ""
    if max_rows and len(df) > max_rows:
        df = df.head(max_rows)
        extra_note = f"\n\n> _Truncated to first {max_rows} rows._\n"

    md = df.to_markdown(index=False)
    content = f"# {src.name}\n\n{md}{extra_note}\n"
    ensure_parent(dst)
    dst.write_text(content, encoding="utf-8")
    try:
        os.utime(dst, (src.stat().st_atime, src.stat().st_mtime))
    except Exception:
        pass
    return True

# ---------- Excel (XLSX / XLS) ----------

def convert_xls_to_xlsx_with_soffice(src: Path, soffice_path: Optional[str]) -> Optional[Path]:
    """LibreOffice headless to turn .xls -> .xlsx. Returns new path or None."""
    soffice = soffice_path or which("soffice") or which("soffice.exe")
    if not soffice:
        maybe_show_soffice_hint()
        return None
    with tempfile.TemporaryDirectory() as td:
        code, out, e = run([soffice, "--headless", "--convert-to", "xlsx", "--outdir", td, str(src)])
        if code != 0:
            warn(f"LibreOffice failed to convert {src.name} -> xlsx: {e.strip() or out.strip()}")
            return None
        out_path = Path(td) / (src.stem + ".xlsx")
        if out_path.exists():
            final = src.with_suffix(".xlsx")
            try:
                shutil.move(str(out_path), str(final))
            except Exception:
                shutil.copy2(str(out_path), str(final))
            log(f"  [ok] Converted {src.name} -> {final.name} via LibreOffice")
            return final
    return None

def convert_excel(src: Path, dst: Path, max_rows: int, soffice_path: Optional[str]) -> bool:
    """Convert Excel workbook to one Markdown file (section per sheet)."""
    try:
        import pandas as pd
    except Exception:
        warn(f"pandas not installed; cannot convert {src.name}. pip install pandas openpyxl xlrd")
        return False

    try:
        xl = pd.ExcelFile(src)  # engine auto
    except Exception as e:
        new_src = convert_xls_to_xlsx_with_soffice(src, soffice_path)
        if new_src is not None:
            try:
                xl = pd.ExcelFile(new_src)
            except Exception as e2:
                err(f"Failed to open Excel via pandas after soffice: {src} / {e2}")
                return False
        else:
            err(f"Failed to open Excel: {src} / {e}")
            return False

    parts = [f"# {src.name}\n"]
    for sheet_name in xl.sheet_names:
        try:
            df = xl.parse(sheet_name=sheet_name, dtype=object)
        except Exception as e:
            warn(f"Sheet read fail '{sheet_name}' in {src.name}: {e}")
            continue

        if df is None or df.empty:
            continue

        extra_note = ""
        if max_rows and len(df) > max_rows:
            df = df.head(max_rows)
            extra_note = f"\n\n> _Truncated to first {max_rows} rows._\n"

        try:
            md = df.to_markdown(index=False)
        except Exception:
            md = df.fillna("").to_markdown(index=False)

        parts.append(f"## {sheet_name}\n\n{md}{extra_note}\n")

    if len(parts) == 1:
        warn(f"No readable sheets in {src.name}")
        return False

    ensure_parent(dst)
    dst.write_text("\n".join(parts), encoding="utf-8")
    try:
        os.utime(dst, (src.stat().st_atime, src.stat().st_mtime))
    except Exception:
        pass
    return True

# ---------- main ----------

def main():
    ap = argparse.ArgumentParser(description="Convert Docs/ -> data/raw as Markdown/text.")
    ap.add_argument("--include-md", action="store_true", help="Also copy .md files to data/raw.")
    ap.add_argument("--no-ocr", action="store_true", help="Disable OCR fallback for PDFs.")
    ap.add_argument("--tesseract", default="", help="Full path to tesseract.exe (optional).")
    ap.add_argument("--lang", default="eng", help="OCR languages, e.g. 'eng+de'.")
    ap.add_argument("--dpi", type=int, default=200, help="OCR rendering DPI (default 200).")
    # Inverted deletion flags: KEEP instead of DELETE
    ap.add_argument("--keep-unsupported", action="store_true", help="Keep unsupported source files by type (default deletes; excludes .doc).")
    ap.add_argument("--keep-source-on-success", action="store_true", help="Keep original file when conversion succeeds (default deletes).")
    ap.add_argument("--keep-source-if-up-to-date", action="store_true", help="Keep original file when skipped as up-to-date (default deletes).")
    ap.add_argument("--doc-timeout", type=int, default=60, help="Timeout (seconds) for Word COM .doc -> .docx conversion.")
    ap.add_argument("--no-word", action="store_true", help="Do not use Microsoft Word COM for .doc (use LibreOffice only if present).")
    ap.add_argument("--max-rows-per-sheet", type=int, default=2000, help="Row limit per sheet for CSV/Excel output (0 = no limit).")
    ap.add_argument("--soffice", default="", help="Explicit path to soffice(.exe). Overrides auto-detect.")
    ap.add_argument("--no-prune-empty-dirs", action="store_true", help="Disable pruning empty directories.")
    args = ap.parse_args()

    if not DOCS_DIR.exists():
        log(f"Docs folder not found: {DOCS_DIR}")
        sys.exit(0)

    # Effective deletion booleans (defaults are destructive)
    delete_unsupported_default = not args.keep_unsupported
    delete_on_success_default = not args.keep_source_on_success
    delete_if_up_to_date_default = not args.keep_source_if_up_to_date

    # Resolve soffice once
    soffice_path = find_soffice(args.soffice.strip() or None)
    if soffice_path:
        log(f"[soffice] Using: {soffice_path}")

    # Pre-sweep: remove any empty dirs before we start
    if not args.no_prune_empty_dirs:
        log("Pre-sweep: pruning empty directories...")
        prune_all_empty_dirs(DOCS_DIR)
        prune_all_empty_dirs(OUT_DIR)

    # Gather files (skip folders)
    all_files = [p for p in DOCS_DIR.rglob("*") if p.is_file()]
    log(f"Scanning: {DOCS_DIR}  (files: {len(all_files)})")

    converted = 0
    skipped = 0
    errs = 0

    iterator = tqdm(all_files, unit="file") if tqdm else all_files
    try:
        for src in iterator:
            rel = src.relative_to(DOCS_DIR)
            out = (OUT_DIR / rel).with_suffix(".md")
            ext = src.suffix.lower()

            try:
                # Global fast path: if output is already up-to-date, skip and (optionally) delete source
                if ext not in {".md"} and file_up_to_date(src, out):
                    log(f"  [skip] up-to-date: {rel}")
                    delete_on_up_to_date(src, delete_if_up_to_date_default)
                    if not args.no_prune_empty_dirs:
                        prune_empty_dirs_upwards((DOCS_DIR / rel).parent, DOCS_DIR)
                    skipped += 1
                    continue

                if ext == ".pdf":
                    ok = convert_pdf(
                        src, out,
                        use_ocr=not args.no_ocr,
                        lang=args.lang, dpi=args.dpi,
                        tesseract_path=args.tesseract or None
                    )
                    if ok:
                        log(f"  [ok] PDF -> {out.relative_to(OUT_DIR)}")
                        converted += 1
                        delete_on_success(src, delete_on_success_default)
                        if not args.no_prune_empty_dirs and delete_on_success_default:
                            prune_empty_dirs_upwards(src.parent, DOCS_DIR)
                    else:
                        skipped += 1

                elif ext == ".md":
                    if args.include_md:
                        try:
                            if out.exists() and out.read_text(encoding="utf-8", errors="ignore") == src.read_text(encoding="utf-8", errors="ignore"):
                                log(f"  [skip] up-to-date: {rel}")
                                delete_on_up_to_date(src, delete_if_up_to_date_default)
                                if not args.no_prune_empty_dirs:
                                    prune_empty_dirs_upwards(src.parent, DOCS_DIR)
                                skipped += 1
                            else:
                                ensure_parent(out)
                                shutil.copy2(src, out)
                                log(f"  [ok] Copied MD -> {out.relative_to(OUT_DIR)}")
                                converted += 1
                                delete_on_success(src, delete_on_success_default)
                                if not args.no_prune_empty_dirs and delete_on_success_default:
                                    prune_empty_dirs_upwards(src.parent, DOCS_DIR)
                        except Exception as e:
                            err(f"{src.name}: {e}")
                            errs += 1
                    else:
                        skipped += 1

                elif ext == ".txt":
                    if file_up_to_date(src, out):
                        log(f"  [skip] up-to-date: {rel}")
                        delete_on_up_to_date(src, delete_if_up_to_date_default)
                        if not args.no_prune_empty_dirs:
                            prune_empty_dirs_upwards(src.parent, DOCS_DIR)
                        skipped += 1
                    else:
                        ensure_parent(out)
                        shutil.copy2(src, out)
                        log(f"  [ok] Copied TXT -> {out.relative_to(OUT_DIR)}")
                        converted += 1
                        delete_on_success(src, delete_on_success_default)
                        if not args.no_prune_empty_dirs and delete_on_success_default:
                            prune_empty_dirs_upwards(src.parent, DOCS_DIR)

                elif ext in {".html", ".htm"}:
                    if convert_html(src, out):
                        log(f"  [ok] HTML -> {out.relative_to(OUT_DIR)}")
                        converted += 1
                        delete_on_success(src, delete_on_success_default)
                        if not args.no_prune_empty_dirs and delete_on_success_default:
                            prune_empty_dirs_upwards(src.parent, DOCS_DIR)
                    else:
                        skipped += 1

                elif ext in {".docx", ".rtf"}:
                    if convert_with_pandoc(src, out):
                        log(f"  [ok] Pandoc -> {out.relative_to(OUT_DIR)}")
                        converted += 1
                        delete_on_success(src, delete_on_success_default)
                        if not args.no_prune_empty_dirs and delete_on_success_default:
                            prune_empty_dirs_upwards(src.parent, DOCS_DIR)
                    else:
                        skipped += 1

                elif ext == ".doc":
                    # If a sibling .docx exists, prefer that and skip the .doc
                    docx_path = src.with_suffix(".docx")
                    if docx_path.exists():
                        log(f"  [skip] .docx present; skipping {rel}")
                        delete_on_up_to_date(src, delete_if_up_to_date_default)
                        if not args.no_prune_empty_dirs:
                            prune_empty_dirs_upwards(src.parent, DOCS_DIR)
                        skipped += 1
                        continue

                    if out.exists() and file_up_to_date(src, out):
                        log(f"  [skip] up-to-date: {rel}")
                        delete_on_up_to_date(src, delete_if_up_to_date_default)
                        if not args.no_prune_empty_dirs:
                            prune_empty_dirs_upwards(src.parent, DOCS_DIR)
                        skipped += 1
                        continue

                    new_src = None
                    if not args.no_word:
                        new_src = convert_doc_to_docx_with_word(src, timeout_sec=args.doc_timeout)
                    if not new_src:
                        new_src = convert_doc_to_docx_with_soffice(src, soffice_path)

                    if new_src and convert_with_pandoc(new_src, out):
                        log(f"  [ok] DOC->DOCX->MD -> {out.relative_to(OUT_DIR)}")
                        converted += 1
                        delete_on_success(src, delete_on_success_default)
                        if delete_on_success_default and new_src.exists():
                            try:
                                new_src.unlink()
                                try:
                                    rel_new = new_src.relative_to(DOCS_DIR)
                                except Exception:
                                    rel_new = new_src
                                log(f"  [del ] intermediate removed: {rel_new}")
                            except Exception:
                                pass
                        if not args.no_prune_empty_dirs and delete_on_success_default:
                            prune_empty_dirs_upwards(src.parent, DOCS_DIR)
                    else:
                        # Do NOT delete .doc here (not an up-to-date case)
                        warn("Cannot convert .doc (blocked/corrupt/missing converters). Try opening in Word and Save As .docx.")
                        if not soffice_path:
                            maybe_show_soffice_hint()
                        skipped += 1

                elif ext in {".csv", ".tsv"}:
                    if convert_csv_like(src, out, max_rows=args.max_rows_per_sheet):
                        log(f"  [ok] {ext.upper().lstrip('.')} -> {out.relative_to(OUT_DIR)}")
                        converted += 1
                        delete_on_success(src, delete_on_success_default)
                        if not args.no_prune_empty_dirs and delete_on_success_default:
                            prune_empty_dirs_upwards(src.parent, DOCS_DIR)
                    else:
                        skipped += 1

                elif ext in {".xlsx", ".xls"}:
                    if convert_excel(src, out, max_rows=args.max_rows_per_sheet, soffice_path=soffice_path):
                        log(f"  [ok] EXCEL -> {out.relative_to(OUT_DIR)}")
                        converted += 1
                        delete_on_success(src, delete_on_success_default)
                        if not args.no_prune_empty_dirs and delete_on_success_default:
                            prune_empty_dirs_upwards(src.parent, DOCS_DIR)
                    else:
                        skipped += 1

                else:
                    log(f"  [skip] unsupported: {rel}")
                    delete_unsupported(src, delete_unsupported_default)
                    if not args.no_prune_empty_dirs and delete_unsupported_default:
                        prune_empty_dirs_upwards(src.parent, DOCS_DIR)
                    skipped += 1

            except KeyboardInterrupt:
                raise
            except Exception as e:
                err(f"{src.name}: {e}")
                errs += 1
            finally:
                if tqdm:
                    iterator.update(0)

    except KeyboardInterrupt:
        warn("Interrupted by user.")

    # Final bottom-up sweep for any empty dirs we may have walked past
    if not args.no_prune_empty_dirs:
        prune_all_empty_dirs(DOCS_DIR)
        prune_all_empty_dirs(OUT_DIR)

    log("\nSummary:")
    log(f"  Converted: {converted}")
    log(f"  Skipped:   {skipped}")
    log(f"  Errors:    {errs}")
    log(f"Output root: {OUT_DIR}")

if __name__ == "__main__":
    main()
