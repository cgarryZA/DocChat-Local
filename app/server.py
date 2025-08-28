from fastapi import FastAPI, HTTPException, Response
from fastapi.staticfiles import StaticFiles
from fastapi.responses import HTMLResponse
from pydantic import BaseModel
from pathlib import Path
import httpx, markdown as md

from .rag import RAG
from .config import OLLAMA_HOST, OLLAMA_MODEL, RAW_DIR

app = FastAPI(title="Manuals RAG (CPU)")
rag = RAG()

# Serve the static chat UI
app.mount("/ui", StaticFiles(directory="static", html=True), name="ui")

class AskReq(BaseModel):
    question: str
    k: int | None = None

@app.post("/ask")
async def ask(req: AskReq):
    k = req.k or 8
    ctx = rag.retrieve(req.question, k=k)
    if not ctx:
        raise HTTPException(404, "No context found. Have you ingested any manuals?")
    prompt = rag.build_prompt(req.question, ctx)
    answer = await rag.generate(prompt)

    # Build citations with clickable links to /view/<doc>#<anchor>
    cites = []
    for i, c in enumerate(ctx):
        link = f"/view/{c['source']}"
        if c.get("anchor"):
            link += f"#{c['anchor']}"
        cites.append({
            "n": i + 1,
            "source": c["source"],
            "section": c["section_title"],
            "anchor": c.get("anchor", ""),
            "link": link,
        })
    return {"answer": answer, "citations": cites, "used": k}

@app.get("/view/{doc_path:path}", response_class=HTMLResponse)
def view_doc(doc_path: str):
    """Render a Markdown manual from data/raw as HTML with heading anchors."""
    src = (RAW_DIR / doc_path).resolve()
    if RAW_DIR.resolve() not in src.parents or not src.exists():
        raise HTTPException(404, "Document not found")
    text = src.read_text(encoding="utf-8", errors="ignore")
    html_body = md.markdown(
        text,
        extensions=["toc", "fenced_code", "tables", "codehilite"],
    )
    html = f"""
    <!doctype html><html><head>
    <meta charset="utf-8"/>
    <title>{src.name}</title>
    <style>
      body {{ font-family: ui-sans-serif,system-ui,-apple-system,Segoe UI,Roboto,Arial;
             margin: 24px; line-height:1.6; max-width:960px; }}
      pre {{ background:#0b1220; color:#e5e7eb; padding:12px; border-radius:8px; overflow:auto; }}
      code {{ font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, monospace; }}
      h1,h2,h3,h4 {{ scroll-margin-top: 80px; }}
      a {{ color:#0ea5e9; text-decoration:none; }}
      a:hover {{ text-decoration:underline; }}
      .topbar {{
        position:sticky;top:0;background:#fff;border-bottom:1px solid #eee;
        padding:10px 0;margin:-24px -24px 16px -24px;
      }}
      .topbar-inner {{ max-width:960px;margin:0 auto;padding:0 24px; }}
    </style>
    </head><body>
      <div class="topbar"><div class="topbar-inner">
        <a href="/ui/">← Back to QA</a> &nbsp; <strong>{doc_path}</strong>
      </div></div>
      {html_body}
    </body></html>
    """
    return HTMLResponse(html)

@app.get("/health")
async def health():
    try:
        async with httpx.AsyncClient(timeout=5) as client:
            r = await client.get(f"{OLLAMA_HOST}/api/tags")
            r.raise_for_status()
        return {"ok": True, "ollama": "up", "model": OLLAMA_MODEL}
    except Exception as e:
        return Response(
            content=f'{{"ok": false, "ollama": "down", "error": "{str(e)}"}}',
            media_type="application/json",
            status_code=503,
        )
