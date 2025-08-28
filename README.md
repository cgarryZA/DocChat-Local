📘 Manuals RAG (CPU-only, local)

Ask questions about your manuals. Everything runs on your machine:

CPU models via Ollama

Your docs stay local

Clickable citations to exact sections

0) Prereqs you need (once)

You only need to install these one time:

Python 3.10+ (recommended 3.11)

Ollama (local LLM runtime) 

Pandoc (recommended for DOCX/DOC/RTF/HTML → Markdown)

Tesseract OCR (optional; for scanned PDFs only)

Git (to clone the repo)

If you skip Pandoc, we’ll still ingest PDF/MD/TXT.
If you skip Tesseract, scanned PDFs (images only) won’t be readable.

1) Install the prerequisites
Windows (PowerShell)

Python

Install from Microsoft Store or python.org.

After install:

py --version


Ollama

Download & install from https://ollama.com/download

Start it:

ollama serve


Leave this window open.

Pandoc (recommended)

Download installer from https://pandoc.org/installing.html

Verify:

pandoc --version


Tesseract OCR (optional; for scanned PDFs)

Install the Windows build from https://github.com/UB-Mannheim/tesseract/wiki?utm_source=chatgpt.com

Add Tesseract to Path

$dir = "C:\Users\<youruser>\AppData\Local\Programs\Tesseract-OCR"
$env:Path = "$env:Path;$dir"
setx PATH "$env:Path;$dir" > $null
setx TESSDATA_PREFIX "$dir\tessdata" > $null

Verify:

tesseract --version


Git

Install from https://git-scm.com/

Verify:

git --version


No admin / locked-down PCs?
You can still run everything: install Ollama & Python normally, and skip Pandoc/Tesseract (or use their user installers). The app will still index PDF/MD/TXT.

macOS (Terminal)
# 1) Homebrew (if you don't have it)
# /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"

# 2) Python
brew install python@3.11

# 3) Ollama (then run `ollama serve` in a separate terminal)
brew install --cask ollama

# 4) Pandoc (recommended)
brew install pandoc

# 5) Tesseract OCR (optional; for scanned PDFs)
brew install tesseract

# 6) Git
brew install git


Open a new terminal and start Ollama:

ollama serve

Ubuntu/Debian (Terminal)
# Python + Git
sudo apt-get update
sudo apt-get install -y python3 python3-venv python3-pip git

# Ollama (follow their site for the latest install method; then:)
ollama serve   # in a separate terminal

# Pandoc (recommended)
sudo apt-get install -y pandoc

# Tesseract OCR (optional; for scanned PDFs)
sudo apt-get install -y tesseract-ocr

2) Clone the repo & add your manuals
git clone <YOUR_REPO_URL>
cd DOCS-LLM


Put your files under Docs/ (subfolders OK). Supported:

PDF (text-based works out of the box; scanned needs Tesseract to OCR)

DOCX / DOC / RTF / HTML / TXT / MD (Pandoc recommended for best results)

3) Run the project (one command)
Windows (PowerShell)
# Make sure Ollama is running in a separate window:  ollama serve
powershell -ExecutionPolicy Bypass -File scripts\bootstrap.ps1

macOS / Linux
# Make sure Ollama is running in a separate terminal:  ollama serve
./scripts/bootstrap.sh


This command will:

Create a virtualenv and install pinned dependencies

(If Ollama is reachable) pull the model (default: qwen2.5:3b-instruct)

Convert everything in Docs/ → Markdown in data/raw/

Pandoc for non-PDFs (if available)

PyMuPDF for PDFs; OCR fallback if Tesseract is installed

Build the FAISS index in data/index/

Serve the app at: http://localhost:8000/ui/

4) Ask questions

Open: http://localhost:8000/ui/

Type a question (e.g., “How do I reset the device to factory settings?”).
Click citation chips to jump to the exact section.

5) Updating docs later

Add/modify files in Docs/, then either re-run the bootstrap script or:

python run_total_convert.py --include-md
python -m app.ingest
uvicorn app.server:app --host 0.0.0.0 --port 8000 --reload

6) Configuration (optional)

Create a .env (or copy .env.example) to override defaults:

OLLAMA_HOST=http://localhost:11434
OLLAMA_MODEL=qwen2.5:3b-instruct
EMBED_MODEL=sentence-transformers/all-MiniLM-L6-v2
CHUNK_TOKENS=900
CHUNK_OVERLAP=120
TOP_K=8

7) Troubleshooting

“500 on /ask”
Make sure Ollama is running and the model is present:

ollama serve                  # in a separate terminal
ollama pull qwen2.5:3b-instruct
curl http://localhost:11434/api/tags


“No context found”
Check converted files exist in data/raw/ and re-ingest:

python run_total_convert.py --include-md
python -m app.ingest


FAISS / NumPy error
This repo pins compatible versions. If you upgraded, roll back:

pip install --upgrade "numpy<2" "faiss-cpu==1.7.4"


Scanned PDFs show little/no text
Install Tesseract, then re-convert:

# Windows/macOS: install Tesseract (see step 1)
# Linux (Debian/Ubuntu):
sudo apt-get install -y tesseract-ocr

# Re-run conversion + ingest
python run_total_convert.py --include-md
python -m app.ingest


DOCX/DOC/RTF/HTML not converting
Install Pandoc and re-run the bootstrap/conversion.

That’s it ✅

Install prereqs → put manuals in Docs/ → run the bootstrap → open /ui.

Everything stays local; no cloud services required.