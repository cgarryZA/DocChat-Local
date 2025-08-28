📘 Manuals RAG (CPU-only, Local)

Ask questions about your manuals — everything runs on your machine:

CPU models via Ollama

Your docs stay local

Clickable citations to exact sections

Table of Contents

Prerequisites (install once)

Install on Windows

Clone & Add Manuals

Run (One Command)

Ask Questions

Update Manuals Later

Configuration (optional)

Troubleshooting

Prerequisites (install once)
Tool	Why you need it	Required?
Python 3.10+ (3.11 recommended)	Run the app & scripts	✅
Ollama	Local LLM runtime (CPU)	✅
Pandoc	Best conversion of DOCX/DOC/RTF/HTML → Markdown	⭐ Recommended
Tesseract OCR	Extract text from scanned PDFs (images only)	➖ Optional
Git	Clone this repository	✅

If you skip Pandoc, we still index PDF / MD / TXT.
If you skip Tesseract, scanned PDFs (image-only) won’t produce text.

Install on Windows

Open PowerShell and follow these steps.

1) Python

Install from the Microsoft Store or python.org
.
Verify:

py --version

2) Ollama

Download & install: https://ollama.com/download

Start it in a separate PowerShell window and leave it running:

ollama serve


(First time? You can pre-pull the model: ollama pull qwen2.5:3b-instruct)

3) Pandoc (recommended)

Download: https://pandoc.org/installing.html

Verify:

pandoc --version

4) Tesseract OCR (optional; for scanned PDFs)

Install (UB Mannheim Windows build):
https://github.com/UB-Mannheim/tesseract/wiki

Option A — Add to PATH (so tesseract works everywhere):

# Replace <youruser> with your Windows username if you used a per-user install
$dir = "C:\Users\<youruser>\AppData\Local\Programs\Tesseract-OCR"
$env:Path = "$env:Path;$dir"
setx PATH "$env:Path;$dir" > $null
setx TESSDATA_PREFIX "$dir\tessdata" > $null
# Close & reopen PowerShell, then:
tesseract --version


Option B — Skip PATH changes: you can pass the full path to tesseract.exe to the bootstrap script later.

5) Git

Install: https://git-scm.com/

Verify:

git --version

Clone & Add Manuals
git clone <YOUR_REPO_URL>
cd DOCS-LLM


Place your manuals under Docs/ (subfolders OK).

Supported:

PDF (text-based out of the box; scanned needs Tesseract)

DOCX / DOC / RTF / HTML / TXT / MD (best with Pandoc)

Run (One Command)

Keep Ollama running in another window (ollama serve), then:

powershell -ExecutionPolicy Bypass -File scripts\bootstrap.ps1


If Tesseract isn’t on PATH, pass its full path:

powershell -ExecutionPolicy Bypass -File scripts\bootstrap.ps1 `
  -Tesseract "C:\Users\<youruser>\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"


To skip OCR entirely (fastest, text-only PDFs):

powershell -ExecutionPolicy Bypass -File scripts\bootstrap.ps1 -NoOCR


What the script does:

Creates a virtual environment and installs pinned dependencies

If Ollama is reachable, pulls the model (qwen2.5:3b-instruct)

Converts Docs/ → Markdown in data/raw/

Pandoc for non-PDFs (if installed)

PyMuPDF for PDFs, Tesseract OCR fallback if available

Builds the FAISS index in data/index/

Serves the UI at http://localhost:8000/ui/

Ask Questions

Open http://localhost:8000/ui/

Ask something like “How do I reset the device to factory settings?”
Click citation chips to jump to the exact section of your manual.

Update Manuals Later

After adding/modifying files in Docs/, either re-run the bootstrap script or:

# Convert Docs -> data/raw
python .\run_total_convert.py --include-md
# Rebuild the index
python -m app.ingest
# Serve (if not already running)
uvicorn app.server:app --host 0.0.0.0 --port 8000 --reload

Configuration (optional)

Create a .env (or copy .env.example) to override defaults:

OLLAMA_HOST=http://localhost:11434
OLLAMA_MODEL=qwen2.5:3b-instruct
EMBED_MODEL=sentence-transformers/all-MiniLM-L6-v2
CHUNK_TOKENS=900
CHUNK_OVERLAP=120
TOP_K=8

Troubleshooting
“500 on /ask”

Ensure Ollama is running and the model exists:

ollama serve
ollama pull qwen2.5:3b-instruct
curl http://localhost:11434/api/tags

“No context found”

Check converted files exist in data/raw/ and re-ingest:

python .\run_total_convert.py --include-md
python -m app.ingest

FAISS / NumPy error

This repo pins compatible versions; if you upgraded, roll back:

pip install --upgrade "numpy<2" "faiss-cpu==1.7.4"

Scanned PDFs show little/no text

Install Tesseract, then re-convert and re-ingest:

python .\run_total_convert.py --include-md
python -m app.ingest

DOCX/DOC/RTF/HTML not converting

Install Pandoc, then re-run the bootstrap or conversion.

Healthcheck: http://localhost:8000/health

API docs: http://localhost:8000/docs

That’s it.
Install prereqs → put manuals in Docs/ → run the bootstrap → open /ui.
Everything stays local; no cloud services required.