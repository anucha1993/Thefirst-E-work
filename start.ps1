# One-click launcher for e-WorkPermit Report Tool
# - First run: creates venv + installs libraries + Playwright Chromium
# - Next runs: opens GUI immediately
#
# Notes are in English only to avoid PS 5.1 ANSI parsing issues.
# Console output uses UTF-8 so Thai messages display correctly.

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8
$Host.UI.RawUI.WindowTitle = "e-WorkPermit Report"

Set-Location -Path $PSScriptRoot

function Write-Step($msg, $color = "Cyan") {
    Write-Host $msg -ForegroundColor $color
}

Write-Host ""
Write-Step "=================================================="
Write-Step "  e-WorkPermit Report Tool"
Write-Step "=================================================="
Write-Host ""

# --- 1) Check Python ---
$pythonCmd = $null
foreach ($cmd in @("python", "py")) {
    try {
        $null = & $cmd --version 2>&1
        if ($LASTEXITCODE -eq 0) { $pythonCmd = $cmd; break }
    } catch {}
}
if (-not $pythonCmd) {
    Write-Host "[X] Python not found." -ForegroundColor Red
    Write-Host "    Please install Python 3.10+ from https://www.python.org/downloads/"
    Write-Host "    Tick 'Add Python to PATH' during installation."
    Write-Host ""
    Read-Host "Press Enter to close"
    exit 1
}
$pyVer = & $pythonCmd --version 2>&1
Write-Host "[OK] Found Python: $pyVer" -ForegroundColor Green

# --- 2) Create venv if missing ---
$venvDir = Join-Path $PSScriptRoot ".venv"
$venvPy  = Join-Path $venvDir "Scripts\python.exe"
$venvPyw = Join-Path $venvDir "Scripts\pythonw.exe"

if (-not (Test-Path $venvPy)) {
    Write-Host ""
    Write-Step "[1/3] Creating virtual environment (.venv) ..." "Yellow"
    & $pythonCmd -m venv .venv
    if (-not (Test-Path $venvPy)) {
        Write-Host "[X] Failed to create venv." -ForegroundColor Red
        Read-Host "Press Enter to close"; exit 1
    }
    Write-Host "      venv created" -ForegroundColor Green

    Write-Host ""
    Write-Step "[2/3] Installing libraries (playwright, dotenv, openpyxl) ..." "Yellow"
    & $venvPy -m pip install --upgrade pip --quiet
    & $venvPy -m pip install -r requirements.txt
    if ($LASTEXITCODE -ne 0) {
        Write-Host "[X] Library install failed." -ForegroundColor Red
        Read-Host "Press Enter to close"; exit 1
    }
    Write-Host "      libraries installed" -ForegroundColor Green

    Write-Host ""
    Write-Step "[3/3] Downloading Chromium browser (~150 MB, one time) ..." "Yellow"
    & $venvPy -m playwright install chromium
    if ($LASTEXITCODE -ne 0) {
        Write-Host "[X] Chromium install failed." -ForegroundColor Red
        Read-Host "Press Enter to close"; exit 1
    }
    Write-Host "      Chromium installed" -ForegroundColor Green

    Write-Host ""
    Write-Step "==================================================" "Green"
    Write-Step "  Setup complete. Next launches will be instant." "Green"
    Write-Step "==================================================" "Green"
}
else {
    Write-Host "[OK] venv ready - skipping install" -ForegroundColor Green
}

# --- 3) Launch GUI ---
Write-Host ""
Write-Step "Launching application..." "Cyan"
if (Test-Path $venvPyw) {
    Start-Process -FilePath $venvPyw -ArgumentList "gui_app.py" -WorkingDirectory $PSScriptRoot | Out-Null
}
else {
    & $venvPy gui_app.py
}
Start-Sleep -Seconds 2
