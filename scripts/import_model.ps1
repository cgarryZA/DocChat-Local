# Importer: restores data/index + data/raw from a ZIP in Models/ (or a given path)
# Usage:
#   powershell -ExecutionPolicy Bypass -File scripts\import_model.ps1               # picks newest Models\*.zip
#   powershell -ExecutionPolicy Bypass -File scripts\import_model.ps1 -Zip Models\manuals-rag-bundle-20250101-120000.zip
#   powershell -ExecutionPolicy Bypass -File scripts\import_model.ps1 -Run          # also launches server

param(
  [string]$Zip = "",
  [switch]$Run
)

$ErrorActionPreference = "Stop"

# repo root (scripts\..)
$repo = Split-Path $MyInvocation.MyCommand.Path -Parent | Split-Path
Set-Location $repo

$modelsDir = Join-Path $repo "Models"
$idxDir    = Join-Path $repo "data\index"
$rawDir    = Join-Path $repo "data\raw"

# pick latest ZIP if not provided
if (-not $Zip -or $Zip.Trim() -eq "") {
  if (-not (Test-Path $modelsDir)) { throw "Models folder not found: $modelsDir" }
  $latest = Get-ChildItem $modelsDir -Filter "*.zip" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
  if (-not $latest) { throw "No ZIPs found in $modelsDir" }
  $Zip = $latest.FullName
} else {
  if (-not (Test-Path $Zip)) { throw "ZIP not found: $Zip" }
}

Write-Host "Using bundle: $Zip"

# clean current data
Remove-Item -Recurse -Force $idxDir 2>$null
Remove-Item -Recurse -Force $rawDir 2>$null
New-Item -ItemType Directory -Force -Path $idxDir | Out-Null
New-Item -ItemType Directory -Force -Path $rawDir | Out-Null

# extract
Expand-Archive -Path $Zip -DestinationPath $repo -Force

# validate
$faiss = Join-Path $idxDir "faiss.index"
$sqlite = Join-Path $idxDir "chunks.sqlite"
if (-not (Test-Path $faiss) -or -not (Test-Path $sqlite)) {
  throw "Bundle missing required files under data\index (faiss.index / chunks.sqlite)."
}

# show manifest if present
$manifest = Join-Path $repo "bundle.json"
if (Test-Path $manifest) {
  try {
    $m = Get-Content $manifest -Raw | ConvertFrom-Json
    Write-Host "Bundle info:"
    Write-Host ("  built_at     : {0}" -f $m.built_at)
    Write-Host ("  embed_model  : {0}" -f $m.embed_model)
    Write-Host ("  chunk_count  : {0}" -f $m.chunk_count)
    Write-Host ("  ollama_model : {0}" -f $m.ollama_model)
  } catch { }
  Remove-Item -Force $manifest 2>$null
}

Write-Host "✅ Imported bundle into:"
Write-Host "  $idxDir"
Write-Host "  $rawDir"

if ($Run) {
  Write-Host "Starting server at http://localhost:8000/ui/ ..."
  Start-Process "http://localhost:8000/ui/"
  uvicorn app.server:app --host 0.0.0.0 --port 8000 --reload
}
