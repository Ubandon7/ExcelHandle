$ErrorActionPreference = "Stop"

$projectRoot = $PSScriptRoot
$venvPython = Join-Path $projectRoot ".venv\Scripts\python.exe"
$requirements = Join-Path $projectRoot "backend\requirements.txt"
$frontendDist = Join-Path $projectRoot "frontend\dist\index.html"
$errorLog = Join-Path $projectRoot "start-error.log"

Set-Location $projectRoot

$serviceUrl = "http://127.0.0.1:8000"
$existingListener = Get-NetTCPConnection -LocalPort 8000 -State Listen -ErrorAction SilentlyContinue
if ($existingListener) {
    try {
        $health = Invoke-RestMethod -Uri "$serviceUrl/api/health" -TimeoutSec 3
        if ($health.status -eq "ok") {
            Write-Host "ExcelHandle is already running: $serviceUrl" -ForegroundColor Green
            if ($env:EXCEL_HANDLE_NO_BROWSER -ne "1") { Start-Process $serviceUrl }
            exit 0
        }
        throw "Health check failed"
    }
    catch {
        throw "Port 8000 is in use by another program. Close it and try again."
    }
}

try {
    Remove-Item -LiteralPath $errorLog -ErrorAction SilentlyContinue

    if (-not (Test-Path $venvPython)) {
        $systemPython = (Get-Command python -ErrorAction Stop).Source
        & $systemPython -m venv (Join-Path $projectRoot ".venv")
        if ($LASTEXITCODE -ne 0) { throw "Failed to create the Python virtual environment." }
    }

    & $venvPython -c "import fastapi, uvicorn, openpyxl, xlrd, pyxlsb" 2>$null
    if ($LASTEXITCODE -ne 0) {
        & $venvPython -m pip install -r $requirements
        if ($LASTEXITCODE -ne 0) { throw "Failed to install Python dependencies." }
    }

    if (-not (Test-Path $frontendDist)) {
        if (-not (Get-Command npm -ErrorAction SilentlyContinue)) {
            throw "Node.js/npm was not found and the frontend has not been built. Install Node.js 20+ and try again."
        }
        Push-Location (Join-Path $projectRoot "frontend")
        try {
            npm install
            if ($LASTEXITCODE -ne 0) { throw "Failed to install frontend dependencies." }
            npm run build
            if ($LASTEXITCODE -ne 0) { throw "Failed to build the frontend." }
        }
        finally {
            Pop-Location
        }
    }

    Write-Host "ExcelHandle started: $serviceUrl" -ForegroundColor Green
    Write-Host "Press Ctrl+C to stop the service." -ForegroundColor DarkGray
    if ($env:EXCEL_HANDLE_NO_BROWSER -ne "1") {
        Start-Job -ScriptBlock {
        param($url)
        for ($attempt = 0; $attempt -lt 30; $attempt++) {
            Start-Sleep -Milliseconds 300
            try {
                $response = Invoke-RestMethod -Uri "$url/api/health" -TimeoutSec 1
                if ($response.status -eq "ok") {
                    Start-Process $url
                    break
                }
            }
            catch {}
        }
        } -ArgumentList $serviceUrl | Out-Null
    }

    & $venvPython -m uvicorn app.main:app --app-dir (Join-Path $projectRoot "backend") --host 127.0.0.1 --port 8000
    if ($LASTEXITCODE -ne 0) { throw "The web service stopped with exit code $LASTEXITCODE." }
}
catch {
    $message = $_.Exception.Message
    $details = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $message`r`n$($_ | Out-String)"
    Set-Content -LiteralPath $errorLog -Value $details -Encoding UTF8
    Write-Host "Startup failed: $message" -ForegroundColor Red
    Write-Host "Error log: $errorLog" -ForegroundColor Yellow
    exit 1
}
