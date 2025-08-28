# One-click exporter: zips data/index + data/raw into Models/<name>.zip
# If chunks.sqlite is locked, uses Python's sqlite3.backup() as a fallback.
# Usage:
#   powershell -ExecutionPolicy Bypass -File scripts\export_model.ps1
#   powershell -ExecutionPolicy Bypass -File scripts\export_model.ps1 -Name "prod"

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

if (-not (Test-Path (Join-Path $idxDir "faiss.index"))) {
  throw "Missing data\index\faiss.index. Run ingestion first:  python -m app.ingest"
}
if (-not (Test-Path (Join-Path $idxDir "chunks.sqlite"))) {
  throw "Missing data\index\chunks.sqlite. Run ingestion first:  python -m app.ingest"
}
if (-not (Test-Path $rawDir)) {
  throw "Missing data\raw. Run conversion first:  python run_total_convert.py --include-md"
}

function Test-Locked {
  param([string]$Path)
  try {
    $fs = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::None)
    $fs.Close()
    return $false
  } catch {
    return $true
  }
}

# Staging dirs
$stage = New-Item -ItemType Directory -Force -Path (Join-Path $env:TEMP ("mr_export_" + [guid]::NewGuid().ToString())) | Select-Object -ExpandProperty FullName
$stageData = Join-Path $stage "data"
$stageIdx  = Join-Path $stage "data\index"
$stageRaw  = Join-Path $stage "data\raw"
New-Item -ItemType Directory -Force -Path $stageIdx, $stageRaw | Out-Null

# Copy raw (mirror)
& robocopy $rawDir $stageRaw /MIR /NFL /NDL /NJH /NJS /NP | Out-Null

# Copy FAISS + meta (if present)
Copy-Item (Join-Path $idxDir "faiss.index") $stageIdx -Force
Copy-Item (Join-Path $idxDir "meta.json") $stageIdx -Force -ErrorAction SilentlyContinue

# Copy sqlite or safe-backup if locked
$sqliteSrc = Join-Path $idxDir "chunks.sqlite"
$sqliteDst = Join-Path $stageIdx "chunks.sqlite"

if (-not (Test-Locked $sqliteSrc)) {
  Copy-Item $sqliteSrc $sqliteDst -Force
  Copy-Item (Join-Path $idxDir "chunks.sqlite-wal") $stageIdx -Force -ErrorAction SilentlyContinue
  Copy-Item (Join-Path $idxDir "chunks.sqlite-shm") $stageIdx -Force -ErrorAction SilentlyContinue
} else {
  Write-Host "chunks.sqlite appears locked; attempting safe backup via Python sqlite3.backup() ..."
  $venvPy = Join-Path $repo ".venv\Scripts\python.exe"
  if (-not (Test-Path $venvPy)) { $venvPy = "python" }

  # Use env vars so we can keep the Python here-string literal
  $env:SQLITE_SRC = $sqliteSrc
  $env:SQLITE_DST = $sqliteDst

  $py = @'
import os, sqlite3
src = os.environ["SQLITE_SRC"]
dst = os.environ["SQLITE_DST"]
con = sqlite3.connect(src, timeout=15)
bk  = sqlite3.connect(dst)
with bk:
    con.backup(bk)
bk.close()
con.close()
print("ok")
'@

  $tmpPy = Join-Path $stage "backup_sqlite.py"
  $py | Out-File -Encoding ascii $tmpPy
  & $venvPy $tmpPy | Out-Null
  Remove-Item $tmpPy -Force -ErrorAction SilentlyContinue
  # clean env
  Remove-Item Env:\SQLITE_SRC -ErrorAction SilentlyContinue
  Remove-Item Env:\SQLITE_DST -ErrorAction SilentlyContinue
}

# Build manifest
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

# Name + zip
$ts   = Get-Date -Format "yyyyMMdd-HHmmss"
$base = if ($Name) { "manuals-rag-$Name-$ts" } else { "manuals-rag-bundle-$ts" }
$zip  = Join-Path $modelsDir ($base + ".zip")
if (Test-Path $zip) { Remove-Item -Force $zip }

Compress-Archive -Path (Join-Path $stage "data"), $bundleJson -DestinationPath $zip -Force

# checksum
$hash = (Get-FileHash $zip -Algorithm SHA256).Hash
$hashPath = $zip + ".sha256"
$hash | Out-File -Encoding ascii $hashPath

Write-Host "Exported bundle:"
Write-Host "  ZIP   : $zip"
Write-Host "  SHA256: $hash"
