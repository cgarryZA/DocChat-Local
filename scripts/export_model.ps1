# One-click exporter: zips data/index + data/raw into Models/<name>.zip
# Usage:
#   powershell -ExecutionPolicy Bypass -File scripts\export_model.ps1
#   powershell -ExecutionPolicy Bypass -File scripts\export_model.ps1 -Name "bundle-prod"

param(
  [string]$Name = ""
)

$ErrorActionPreference = "Stop"

# repo root (scripts\..)
$repo = Split-Path $MyInvocation.MyCommand.Path -Parent | Split-Path
Set-Location $repo

$modelsDir = Join-Path $repo "Models"
$idxDir    = Join-Path $repo "data\index"
$rawDir    = Join-Path $repo "data\raw"

New-Item -ItemType Directory -Force -Path $modelsDir | Out-Null

# sanity checks
if (-not (Test-Path (Join-Path $idxDir "faiss.index")) -or -not (Test-Path (Join-Path $idxDir "chunks.sqlite"))) {
  throw "Missing data/index artifacts. Run ingestion first:  python -m app.ingest"
}
if (-not (Test-Path $rawDir)) {
  throw "Missing data/raw. Run conversion first:  python run_total_convert.py --include-md"
}

# read meta if present
$metaPath = Join-Path $idxDir "meta.json"
$embedModel = ""
$chunkCount = $null
if (Test-Path $metaPath) {
  try {
    $meta = Get-Content $metaPath -Raw | ConvertFrom-Json
    $embedModel = $meta.embed_model
    $chunkCount = $meta.count
  } catch { }
}

# manifest
$stage = New-Item -ItemType Directory -Force -Path (Join-Path $env:TEMP ("mr_export_" + [guid]::NewGuid().ToString())) | Select-Object -ExpandProperty FullName
$manifest = @{
  app          = "Manuals RAG"
  schema       = "bundle-v1"
  built_at     = (Get-Date).ToString("s")
  ollama_model = $env:OLLAMA_MODEL
  embed_model  = $embedModel
  chunk_count  = $chunkCount
}
$bundleJson = Join-Path $stage "bundle.json"
$manifest | ConvertTo-Json | Out-File -Encoding utf8 $bundleJson

# name
$ts   = Get-Date -Format "yyyyMMdd-HHmmss"
$base = if ($Name) { "manuals-rag-$Name-$ts" } else { "manuals-rag-bundle-$ts" }
$zip  = Join-Path $modelsDir ($base + ".zip")

# zip (keep folder structure)
if (Test-Path $zip) { Remove-Item -Force $zip }
Compress-Archive -Path "data\index","data\raw",$bundleJson -DestinationPath $zip -Force

# checksum
$hash = (Get-FileHash $zip -Algorithm SHA256).Hash
$hashPath = $zip + ".sha256"
$hash | Out-File -Encoding ascii $hashPath

Write-Host "✅ Exported bundle:"
Write-Host "  ZIP : $zip"
Write-Host "  SHA256: $hash"
