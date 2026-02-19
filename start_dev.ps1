# start_dev.ps1
# Starts the FastAPI development server with hot-reload
# Usage: .\start_dev.ps1

param(
    [int]$Port = 8000,
    [switch]$NoBrowser
)

$Host.UI.RawUI.WindowTitle = "Invoice Generator - Dev Server"

Write-Host ""
Write-Host "  ╔══════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "  ║     Invoice Generator - Development Server   ║" -ForegroundColor Cyan
Write-Host "  ╚══════════════════════════════════════════════╝" -ForegroundColor Cyan
Write-Host ""

# Check if Python is available
$pythonVersion = python --version 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-Host "  [ERROR] Python is not installed or not in PATH" -ForegroundColor Red
    Read-Host "Press Enter to exit..."
    exit 1
}
Write-Host "  [OK] $pythonVersion" -ForegroundColor Green

# Check if uvicorn is installed
python -c "import uvicorn" 2>&1 | Out-Null
if ($LASTEXITCODE -ne 0) {
    Write-Host "  [INFO] Installing uvicorn..." -ForegroundColor Yellow
    pip install uvicorn[standard] --quiet
}
Write-Host "  [OK] Uvicorn installed" -ForegroundColor Green

Write-Host ""
Write-Host "  🚀 Starting server on http://localhost:$Port" -ForegroundColor Yellow
Write-Host "  📁 Frontend: http://localhost:$Port/frontend/" -ForegroundColor Yellow
Write-Host ""
Write-Host "  Press Ctrl+C to stop the server" -ForegroundColor DarkGray
Write-Host "  ─────────────────────────────────────────────────" -ForegroundColor DarkGray
Write-Host ""

# Clean Python bytecode cache to force fresh imports
Write-Host "  🧹 Clearing Python cache (__pycache__)..." -ForegroundColor DarkGray
Get-ChildItem -Path . -Recurse -Directory -Filter "__pycache__" -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
Write-Host "  [OK] Python cache cleared" -ForegroundColor Green

# Generate a cache-bust token for the browser URL
$cacheBust = Get-Date -Format "HHmmss"
Write-Host "  🧹 Browser cache-bust token: $cacheBust (hard-refresh with Ctrl+Shift+R)" -ForegroundColor DarkGray

# Open browser after a short delay (unless disabled)
if (-not $NoBrowser) {
    Start-Job -ScriptBlock {
        Start-Sleep -Seconds 2
        Start-Process "http://localhost:$using:Port/frontend/?v=$using:cacheBust"
    } | Out-Null
}

# Start uvicorn with hot-reload
try {
    uvicorn api.main:app --reload --host 0.0.0.0 --port $Port
}
catch {
    Write-Host ""
    Write-Host "  [ERROR] Server crashed: $_" -ForegroundColor Red
}

Write-Host ""
Write-Host "  Server stopped." -ForegroundColor Yellow
Read-Host "Press Enter to exit..."
