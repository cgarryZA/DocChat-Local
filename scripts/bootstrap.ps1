# Bootstrap script for Windows (PowerShell)
# Usage: powershell -ExecutionPolicy Bypass -File scripts\bootstrap.ps1
# Optional params:
#   -Tesseract "C:\Path\to\tesseract.exe"   (for OCR on scanned PDFs)
#   -NoOCR                                 (disable OCR fallback entirely)

param(
  [string]$Tesseract = "",   # e.g. "C:\Users\YOU\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"
  [switch]$NoOCR
)

# One-command local setup & run (Windows)
$ErrorActionPreference = "Stop"

# Go to repo root
$repo = Split-Path $MyInvocation.MyCommand.Path -Parent | Split-Path
Set-Location $repo

# 1) Venv + deps
if (-not (Test-Path ".venv")) { py -3 -m venv .venv }
& .\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt

# 2) Ensure folders
New-Item -ItemType Directory -Force -Path data\raw | Out-Null
New-Item -ItemType Directory -Force -Path data\index | Out-Null

# 3) Check Ollama; try to pull model if API is reachable
if ([string]::IsNullOrWhiteSpace($env:OLLAMA_HOST)) { $env:OLLAMA_HOST = "http://localhost:11434" }
if ([string]::IsNullOrWhiteSpace($env:OLLAMA_MODEL)) { $env:OLLAMA_MODEL = "qwen2.5:3b-instruct" }

try {
  $tags = Invoke-WebRequest -UseBasicParsing -Uri "$env:OLLAMA_HOST/api/tags" -TimeoutSec 3
  if ($tags.StatusCode -eq 200) {
    Write-Host "✅ Ollama reachable. Ensuring model: $env:OLLAMA_MODEL"
    Invoke-WebRequest -UseBasicParsing -Method POST -Uri "$env:OLLAMA_HOST/api/pull" `
      -Body (@{ name = $env:OLLAMA_MODEL } | ConvertTo-Json) -ContentType "application/json" | Out-Null
  } else {
    Write-Host "⚠️  Ollama API not responding. Start Ollama (ollama serve) to enable answers." -ForegroundColor Yellow
  }
} catch {
  Write-Host "⚠️  Ollama not found or not running. Install from https://ollama.com/download and run 'ollama serve'." -ForegroundColor Yellow
}

# 3b) Optional: wire up Tesseract for this session if a path was provided
$convertArgs = @("--include-md")
if ($NoOCR) { $convertArgs += "--no-ocr" }

if (-not [string]::IsNullOrWhiteSpace($Tesseract)) {
  if (-not (Test-Path $Tesseract)) {
    Write-Host "⚠️  Tesseract path not found: $Tesseract  (continuing without OCR)" -ForegroundColor Yellow
    $convertArgs += "--no-ocr"
  } else {
    $convertArgs += @("--tesseract", $Tesseract)
    $tessDir = Split-Path $Tesseract -Parent
    # make sure tessdata is discoverable for this shell
    $env:Path = "$env:Path;$tessDir"
    $env:TESSDATA_PREFIX = Join-Path $tessDir "tessdata"
  }
}

# 3c) Nice-to-have: warn if Pandoc missing (non-PDF conversions)
if (-not (Get-Command pandoc -ErrorAction SilentlyContinue)) {
  Write-Host "ℹ️  Pandoc not found. DOCX/RTF/HTML will be skipped. Install from https://pandoc.org/installing" -ForegroundColor Yellow
}

# 4) Convert Docs -> data/raw (uses pandoc when available, PyMuPDF for PDFs)
python .\run_total_convert.py @convertArgs

# 5) Ingest
python -m app.ingest

# 6) Run server
Start-Process "http://localhost:8000/ui/"
uvicorn app.server:app --host 0.0.0.0 --port 8000 --reload
