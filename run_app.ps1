# run_app.ps1

$ScriptDir = $PSScriptRoot
Set-Location $ScriptDir

Write-Host "==================================" -ForegroundColor Cyan
Write-Host "   INVOICE GENERATOR LAUNCHER     " -ForegroundColor Cyan
Write-Host "==================================" -ForegroundColor Cyan

# 1. Try to activate virtual environment if it exists
$VenvScript = Join-Path $ScriptDir ".venv\Scripts\Activate.ps1"
if (Test-Path $VenvScript) {
    Write-Host "[INFO] Activating virtual environment..." -ForegroundColor Green
    . $VenvScript
} else {
    Write-Host "[WARN] No .venv found in root. Using global Python interpreter." -ForegroundColor Yellow
}

# 2. Start the Backend Server (FastAPI)
Write-Host "[INFO] Starting Backend Server (Uvicorn)..." -ForegroundColor Cyan
$BackendTitle = "InvoiceGenerator_Backend_Log"

# We use Start-Process to run it in a separate window so the main script isn't blocked 
# or so we can keep the console output visible in its own window.
try {
    Start-Process -FilePath "python" `
                  -ArgumentList "-m uvicorn api.main:app --reload --host 127.0.0.1 --port 8000" `
                  -WorkingDirectory $ScriptDir `
                  -WindowStyle Normal
} catch {
    Write-Error "Failed to start backend. Is Python installed and added to PATH?"
    Read-Host "Press Enter to exit..."
    exit 1
}

# 3. Wait a moment for server to warm up
Write-Host "[INFO] Waiting 3 seconds for server to initialize..." -ForegroundColor Cyan
Start-Sleep -Seconds 3

# 4. Open the Frontend in Default Browser
$Url = "http://127.0.0.1:8000/frontend/index.html"
Write-Host "[INFO] Opening Browser: $Url" -ForegroundColor Green
try {
    Start-Process $Url
} catch {
    Write-Host "[WARN] Could not open default browser. Please navigate to $Url manually." -ForegroundColor Yellow
}

Write-Host "[SUCCESS] Application launched." -ForegroundColor Green
Write-Host "You can close this window, but keep the Backend window open." -ForegroundColor Gray
Start-Sleep -Seconds 5
