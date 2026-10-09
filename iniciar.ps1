# iniciar.ps1 - Script PowerShell para iniciar o Sistema do 46º BPM
$OutputEncoding = [System.Text.Encoding]::UTF8

Write-Host "======================================================================" -ForegroundColor Cyan
Write-Host "   46º BATALHÃO DE POLÍCIA MILITAR - BAGÉ / RS" -ForegroundColor Yellow
Write-Host "   GESTÃO DE FROTA DE VIATURAS E EQUIPAMENTOS DO 46º BPM" -ForegroundColor Cyan
Write-Host "======================================================================" -ForegroundColor Cyan
Write-Host ""

$pythonExe = $null

# 1. Verifica se existe ambiente virtual venv
$venvPy = Join-Path $PSScriptRoot "venv\Scripts\python.exe"
if (Test-Path $venvPy) {
    $pythonExe = $venvPy
} elseif (Get-Command python -ErrorAction SilentlyContinue) {
    $pythonExe = "python"
} elseif (Test-Path "$env:LOCALAPPDATA\Programs\Python\Python312\python.exe") {
    $pythonExe = "$env:LOCALAPPDATA\Programs\Python\Python312\python.exe"
}

if (-not $pythonExe) {
    Write-Host "[ERRO] Python não foi encontrado no sistema." -ForegroundColor Red
    pause
    exit 1
}

Write-Host "[+] Python detectado: $pythonExe" -ForegroundColor Green

# Verifica se dependências estão instaladas
& $pythonExe -c "import reportlab, flask, pandas, openpyxl" 2>$null
if ($LASTEXITCODE -ne 0) {
    Write-Host "[!] Instalando dependencias necessarias (requirements.txt)..." -ForegroundColor Yellow
    $reqPath = Join-Path $PSScriptRoot "requirements.txt"
    & $pythonExe -m pip install -r $reqPath
}

# Verifica se já está rodando na porta 5000
$portListening = Get-NetTCPConnection -LocalPort 5000 -State Listen -ErrorAction SilentlyContinue
if ($portListening) {
    Write-Host "[+] O sistema ja esta rodando na porta 5000!" -ForegroundColor Green
    Write-Host "[+] Abrindo navegador em http://127.0.0.1:5000..." -ForegroundColor Green
    Start-Process "http://127.0.0.1:5000"
    exit 0
}

Write-Host "[+] Inicializando banco SQLite..." -ForegroundColor Green
& $pythonExe database.py

# Iniciar navegador
Start-Job -ScriptBlock {
    Start-Sleep -Seconds 2
    Start-Process "http://127.0.0.1:5000"
} | Out-Null

Write-Host "[+] Abrindo navegador em http://127.0.0.1:5000..." -ForegroundColor Green
Write-Host "[+] Iniciando servidor Flask..." -ForegroundColor Green
Write-Host ""

& $pythonExe app.py
