# start_dv.ps1 - Fabric DV Image Embedder watcher launcher
# Launched by: RUN_AI_SHIPPING_SYSTEM.bat

$Host.UI.RawUI.WindowTitle = "DV BOT  |  Fabric DV Image Embedder  |  Tropic Knits"

# -- Cloud dashboard -------------------------------------------------------
$DashboardUrl = "https://fabric-dv-bot.onrender.com"
$PushKey      = "DVp8mK2xNqLc9ZhWjReYsF5aTbGvUoE3"

$env:DV_RENDER_URL           = $DashboardUrl
$env:DV_PUSH_API_KEY         = $PushKey
$env:PYTHONDONTWRITEBYTECODE = "1"

Set-Location "C:\IT\DV"

Write-Host ""
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   DV BOT  -  Fabric DV Image Embedder  -  Tropic Knits" -ForegroundColor Cyan
Write-Host "  ================================================" -ForegroundColor Cyan
Write-Host "   Dashboard : $DashboardUrl" -ForegroundColor DarkGray
Write-Host "   Bot dir   : C:\IT\DV" -ForegroundColor DarkGray
Write-Host ""

# -- Python check ----------------------------------------------------------
if (-not (Get-Command python -ErrorAction SilentlyContinue)) {
    Write-Host "  [ERROR] Python not found on PATH." -ForegroundColor Red
    Read-Host "  Press Enter to close"
    exit 1
}

# -- Dependency check ------------------------------------------------------
Write-Host "  Checking dependencies..." -ForegroundColor Cyan
python -c "import playwright, openpyxl, PIL" 2>&1 | Out-Null
if ($LASTEXITCODE -ne 0) {
    Write-Host "  Installing missing packages..." -ForegroundColor Yellow
    python -m pip install playwright openpyxl Pillow --quiet
    python -m playwright install chromium
}

Write-Host "  Starting DV watcher..." -ForegroundColor Cyan
Write-Host ""

python dv_watcher.py

Write-Host ""
Write-Host "  DV watcher stopped." -ForegroundColor Yellow
Read-Host "  Press Enter to close"
