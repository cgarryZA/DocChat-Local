import sqlite3
import numpy as np
import httpx
from typing import List, Dict
from sentence_transformers import SentenceTransformer
import faiss

from .config import INDEX_DIR, EMBED_MODEL, OLLAMA_HOST, OLLAMA_MODEL, TOP_K

DB_PATH = INDEX_DIR / "chunks.sqlite"
FAISS_PATH = str(INDEX_DIR / "faiss.index")

class RAG:
    def __init__(self):
        # keep the SQLite connection open; row_factory not needed since we select fixed columns
        self.con = sqlite3.connect(DB_PATH)
        self.cur = self.con.cursor()
        self.embedder = SentenceTransformer(EMBED_MODEL)
        self.index = faiss.read_index(FAISS_PATH)

    def retrieve(self, query: str, k: int = TOP_K) -> List[Dict]:
        q = self.embedder.encode([query], normalize_embeddings=True).astype("float32")
        scores, idx = self.index.search(q, k)
        out: List[Dict] = []
        # faiss returns 0-based indexes aligned with insertion order; our SQLite ids start at 1
        for rank, i in enumerate(idx[0].tolist()):
            self.cur.execute(
                "SELECT id, source, section_title, anchor, text FROM chunks WHERE id=?",
                (i + 1,),
            )
            row = self.cur.fetchone()
            if row:
                out.append(
                    {
                        "rank": rank + 1,
                        "id": row[0],
                        "source": row[1],
                        "section_title": row[2],
                        "anchor": row[3],
                        "text": row[4],
                        "score": float(scores[0][rank]),
                    }
                )
        return out

    def build_prompt(self, question: str, ctx: List[Dict]) -> str:
        # Compose a compact, citation-aware prompt
        parts = [
            "You answer only from the provided context. If info is missing, say so and suggest the closest section.",
            "Cite sources like [n] where n is the context item number.",
            "",
            f"Question:\n{question}\n",
            "Context:",
        ]
        for i, c in enumerate(ctx, start=1):
            header = f"[{i}] {c['section_title']} — {c['source']}"
            parts.append(header)
            parts.append(c["text"])
            parts.append("")
        parts.append("Answer:")
        return "\n".join(parts)

    async def generate(self, prompt: str) -> str:
        # Ollama /api/generate streaming=false
        payload = {"model": OLLAMA_MODEL, "prompt": prompt, "stream": False}
        async with httpx.AsyncClient(timeout=120) as client:
            r = await client.post(f"{OLLAMA_HOST}/api/generate", json=payload)
            r.raise_for_status()
            data = r.json()
            return data.get("response", "")
