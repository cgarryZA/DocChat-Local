# One-command local setup & run (Windows)
$ErrorActionPreference = "Stop"
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
$env:OLLAMA_HOST = $env:OLLAMA_HOST -ne $null ? $env:OLLAMA_HOST : "http://localhost:11434"
$env:OLLAMA_MODEL = $env:OLLAMA_MODEL -ne $null ? $env:OLLAMA_MODEL : "qwen2.5:3b-instruct"

try {
  $tags = Invoke-WebRequest -UseBasicParsing -Uri "$env:OLLAMA_HOST/api/tags" -TimeoutSec 3
  if ($tags.StatusCode -eq 200) {
    Write-Host "✅ Ollama reachable. Ensuring model: $env:OLLAMA_MODEL"
    Invoke-WebRequest -UseBasicParsing -Method POST -Uri "$env:OLLAMA_HOST/api/pull" `
      -Body (@{ name=$env:OLLAMA_MODEL } | ConvertTo-Json) -ContentType "application/json" | Out-Null
  } else {
    Write-Host "⚠️  Ollama API not responding. Start Ollama (ollama serve) to enable answers." -ForegroundColor Yellow
  }
} catch {
  Write-Host "⚠️  Ollama not found or not running. Install from https://ollama.com/download and run 'ollama serve'." -ForegroundColor Yellow
}

# 4) Convert Docs -> data/raw (uses pandoc when available, PyMuPDF for PDFs)
python .\run_total_convert.py --include-md

# 5) Ingest
python -m app.ingest

# 6) Run server
Start-Process "http://localhost:8000/ui/"
uvicorn app.server:app --host 0.0.0.0 --port 8000 --reload
