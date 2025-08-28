import os, json, sqlite3, re
from pathlib import Path
from typing import List, Dict
from markdownify import markdownify as to_md
from sentence_transformers import SentenceTransformer
import numpy as np
import faiss

from .config import RAW_DIR, INDEX_DIR, EMBED_MODEL, CHUNK_TOKENS, CHUNK_OVERLAP

DB_PATH = INDEX_DIR / "chunks.sqlite"
FAISS_PATH = INDEX_DIR / "faiss.index"
META_PATH = INDEX_DIR / "meta.json"

SUPPORTED_EXTS = {".md", ".txt", ".html", ".htm"}

# ---------- helpers ----------

def load_embedder():
    return SentenceTransformer(EMBED_MODEL)

def html_to_markdown(html: str) -> str:
    return to_md(html)

def read_any(path: Path) -> str:
    ext = path.suffix.lower()
    if ext in (".md", ".txt"):
        return path.read_text(encoding="utf-8", errors="ignore")
    if ext in (".html", ".htm"):
        html = path.read_text(encoding="utf-8", errors="ignore")
        return html_to_markdown(html)
    # PDFs are handled upstream (converted to .md before ingest)
    raise ValueError(f"Unsupported file type for now: {ext}")

def simple_tokenize(s: str) -> List[str]:
    # whitespace tokens (good enough for chunk sizing)
    return s.split()

def slugify_heading(s: str) -> str:
    """Generate URL-friendly anchor ids that match common markdown renderers."""
    s = s.strip().lower()
    s = re.sub(r"[^\w\s-]", "", s)     # drop punctuation
    s = re.sub(r"[\s_]+", "-", s)      # spaces/underscores -> dashes
    s = re.sub(r"-{2,}", "-", s).strip("-")
    return s

def chunk_markdown(md: str, source: str) -> List[Dict]:
    """Split a markdown doc into heading-aware, token-limited chunks.
    Each chunk carries its section title path and a stable anchor derived from the deepest heading.
    """
    lines = md.splitlines()
    sections = []
    path_stack: List[str] = []
    buf: List[str] = []

    def flush_section():
        if buf:
            sections.append(("\n".join(buf), path_stack.copy()))
            buf.clear()

    for line in lines:
        if line.startswith("#"):
            flush_section()
            # update heading path
            level = len(line) - len(line.lstrip("#"))
            title = line.lstrip("#").strip()
            # ensure path stack length == level
            path_stack = path_stack[: max(0, level - 1)] + [title]
            buf.append(line)
        else:
            buf.append(line)
    flush_section()

    chunks = []
    from .utils import sliding_window, md_title_path
    for text, path in sections:
        toks = simple_tokenize(text)
        anchor = slugify_heading(path[-1]) if path else ""
        for s, e in sliding_window(toks, CHUNK_TOKENS, CHUNK_OVERLAP):
            piece = " ".join(toks[s:e])
            chunks.append(
                {
                    "source": source,
                    "section_title": md_title_path(path),
                    "anchor": anchor,
                    "text": piece,
                }
            )
    return chunks

def build_sqlite():
    con = sqlite3.connect(DB_PATH)
    cur = con.cursor()
    # Recreate table to ensure we have the 'anchor' column
    cur.execute("DROP TABLE IF EXISTS chunks")
    cur.execute(
        """
        CREATE TABLE chunks (
          id INTEGER PRIMARY KEY,
          source TEXT,
          section_title TEXT,
          anchor TEXT,
          text TEXT
        )
        """
    )
    con.commit()
    return con

# ---------- main pipeline ----------

def main():
    embedder = load_embedder()
    con = build_sqlite()
    cur = con.cursor()

    all_chunks = []
    # Walk RAW_DIR, skipping noisy dirs
    for root, dirs, files in os.walk(RAW_DIR):
        skip_dirs = {".git", ".venv", "__pycache__", "node_modules", "site-packages"}
        dirs[:] = [d for d in dirs if d not in skip_dirs]
        for f in files:
            p = Path(root) / f
            if p.suffix.lower() not in SUPPORTED_EXTS:
                # keep log minimal, but surface truly odd cases
                if p.suffix.lower() not in {".png", ".jpg", ".jpeg", ".gif", ".pdf"}:
                    print(f"Skip (unsupported ext): {p}")
                continue
            rel_source = str(p.relative_to(RAW_DIR))
            try:
                md = read_any(p)
            except ValueError:
                print(f"Skip (unsupported): {p}")
                continue
            chunks = chunk_markdown(md, source=rel_source)
            all_chunks.extend(chunks)

    if not all_chunks:
        print("No chunks produced. Put manuals into data/raw and retry.")
        return

    # Insert rows and collect texts for embedding
    texts = []
    for ch in all_chunks:
        cur.execute(
            "INSERT INTO chunks(source, section_title, anchor, text) VALUES (?,?,?,?)",
            (ch["source"], ch["section_title"], ch["anchor"], ch["text"]),
        )
        texts.append(ch["text"])
    con.commit()

    # Embed
    embs = embedder.encode(
        texts, batch_size=64, show_progress_bar=True, normalize_embeddings=True
    )
    embs = np.asarray(embs, dtype="float32")
    dim = embs.shape[1]
    index = faiss.IndexFlatIP(dim)  # cosine if normalized
    index.add(embs)
    faiss.write_index(index, str(FAISS_PATH))

    # Save meta
    meta = {
        "count": len(all_chunks),
        "embed_model": EMBED_MODEL,
        "schema": {"chunks": ["id", "source", "section_title", "anchor", "text"]},
    }
    META_PATH.write_text(json.dumps(meta, indent=2), encoding="utf-8")
    print(f"Indexed {meta['count']} chunks.")

if __name__ == "__main__":
    main()
