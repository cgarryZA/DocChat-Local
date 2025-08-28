# Importer: restores data/index + data/raw from a ZIP in Models/ (or a given path)
# Stops the running server (incl. uvicorn reloader) on port 8000, imports, then restarts if it was running (or if -Run is passed).
# Usage:
#   powershell -ExecutionPolicy Bypass -File scripts\import_model.ps1
#   powershell -ExecutionPolicy Bypass -File scripts\import_model.ps1 -Zip Models\bundle.zip
#   powershell -ExecutionPolicy Bypass -File scripts\import_model.ps1 -Run

param(
  [string]$Zip = "",
  [switch]$Run
)

$ErrorActionPreference = "Stop"

# --- helpers ---
function Get-ServerPids {
  param([int]$Port = 8000)
  try {
    $conns = Get-NetTCPConnection -LocalPort $Port -State Listen, Established -ErrorAction SilentlyContinue
    if ($null -eq $conns) { return @() }
    return ($conns | Select-Object -ExpandProperty OwningProcess | Sort-Object -Unique)
  } catch { return @() }
}

function Stop-Server {
  param([int]$Port = 8000)
  $pids = Get-ServerPids -Port $Port
  $stopped = $false

  if ($pids.Count -gt 0) {
    Write-Host "Stopping server on port $Port (PIDs: $($pids -join ', ')) ..."
    foreach ($procId in $pids) {
      try { Stop-Process -Id $procId -Force -ErrorAction SilentlyContinue } catch {}
    }
    $stopped = $true
  }

  # Kill any uvicorn/python reloader children still hanging around
  try {
    $procs = Get-CimInstance Win32_Process | Where-Object {
      ($_.Name -match 'python' -or $_.Name -match 'uvicorn' -or $_.Name -match 'watchfiles') -and
      ($_.CommandLine -match 'uvicorn' -or $_.CommandLine -match 'app.server:app' -or $_.CommandLine -match 'DOCS-LLM')
    }
    foreach ($p in $procs) {
      try { Stop-Process -Id $p.ProcessId -Force -ErrorAction SilentlyContinue } catch {}
      $stopped = $true
    }
  } catch {}

  # wait until port is free (up to 10s)
  $deadline = (Get-Date).AddSeconds(10)
  while ((Get-ServerPids -Port $Port).Count -gt 0 -and (Get-Date) -lt $deadline) {
    Start-Sleep -Milliseconds 200
  }
  return $stopped
}

function Wait-Unlocked {
  param([string]$Path, [int]$TimeoutSec = 30)
  $deadline = (Get-Date).AddSeconds($TimeoutSec)
  while ((Get-Date) -lt $deadline) {
    try {
      $fs = [System.IO.File]::Open($Path, [System.IO.FileMode]::OpenOrCreate, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
      $fs.Close()
      return $true
    } catch {
      Start-Sleep -Milliseconds 300
    }
  }
  return $false
}

# --- repo paths ---
$repo = Split-Path $MyInvocation.MyCommand.Path -Parent | Split-Path
Set-Location $repo

$modelsDir = Join-Path $repo "Models"
$idxDir    = Join-Path $repo "data\index"
$rawDir    = Join-Path $repo "data\raw"

# --- pick zip ---
if (-not $Zip -or $Zip.Trim() -eq "") {
  if (-not (Test-Path $modelsDir)) { throw "Models folder not found: $modelsDir" }
  $latest = Get-ChildItem $modelsDir -Filter "*.zip" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
  if (-not $latest) { throw "No ZIPs found in $modelsDir" }
  $Zip = $latest.FullName
} else {
  if (-not (Test-Path $Zip)) { throw "ZIP not found: $Zip" }
}
Write-Host "Using bundle: $Zip"

# --- stop server if running ---
$wasRunning = (Get-ServerPids -Port 8000).Count -gt 0
if ($wasRunning) { [void](Stop-Server -Port 8000) }

# --- ensure DB not locked (best-effort) ---
$dbPath = Join-Path $idxDir "chunks.sqlite"
if (Test-Path $dbPath) {
  if (-not (Wait-Unlocked -Path $dbPath -TimeoutSec 30)) {
    Write-Host "Warning: database still appears locked after stopping server."
    Write-Host "If you have it open in another tool (e.g. VS Code SQLite extension), close it and rerun."
    throw "chunks.sqlite is locked by another process; import aborted to keep data consistent."
  }
}

# --- clean current data ---
if (Test-Path $idxDir) { Remove-Item -Recurse -Force $idxDir -ErrorAction SilentlyContinue }
if (Test-Path $rawDir) { Remove-Item -Recurse -Force $rawDir -ErrorAction SilentlyContinue }
New-Item -ItemType Directory -Force -Path $idxDir | Out-Null
New-Item -ItemType Directory -Force -Path $rawDir | Out-Null

# --- extract bundle to repo root (contains 'data/' + optional bundle.json) ---
Expand-Archive -Path $Zip -DestinationPath $repo -Force

# --- validate ---
$faiss  = Join-Path $idxDir "faiss.index"
$sqlite = Join-Path $idxDir "chunks.sqlite"
if (-not (Test-Path $faiss) -or -not (Test-Path $sqlite)) {
  throw "Bundle missing required files under data\index (faiss.index / chunks.sqlite)."
}

# --- show manifest if present ---
$manifest = Join-Path $repo "bundle.json"
if (Test-Path $manifest) {
  try {
    $m = Get-Content $manifest -Raw | ConvertFrom-Json
    Write-Host "Bundle info:"
    if ($m.built_at)     { Write-Host ("  built_at     : {0}" -f $m.built_at) }
    if ($m.embed_model)  { Write-Host ("  embed_model  : {0}" -f $m.embed_model) }
    if ($m.chunk_count)  { Write-Host ("  chunk_count  : {0}" -f $m.chunk_count) }
    if ($m.ollama_model) { Write-Host ("  ollama_model : {0}" -f $m.ollama_model) }
  } catch {}
  Remove-Item -Force $manifest -ErrorAction SilentlyContinue
}

Write-Host "Imported bundle into:"
Write-Host "  $idxDir"
Write-Host "  $rawDir"

# --- restart server if it was running OR -Run passed ---
if ($wasRunning -or $Run) {
  Write-Host "Starting server at http://localhost:8000/ui/ ..."
  Start-Process "http://localhost:8000/ui/"
  $cmd = 'uvicorn app.server:app --host 0.0.0.0 --port 8000 --reload'
  Start-Process -FilePath "powershell" -ArgumentList "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", $cmd | Out-Null
}
