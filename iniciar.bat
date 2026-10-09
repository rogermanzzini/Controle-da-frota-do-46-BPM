@echo off
chcp 65001 > nul
title 46º BPM - Gestão de Frota e Equipamentos

echo ======================================================================
echo    46º BATALHÃO DE POLÍCIA MILITAR - BAGÉ / RS
echo    GESTÃO DE FROTA DE VIATURAS E EQUIPAMENTOS DO 46º BPM
echo ======================================================================
echo.

REM 1. Verifica se existe o ambiente virtual venv
if exist "%~dp0venv\Scripts\python.exe" (
    set "PYCMD=%~dp0venv\Scripts\python.exe"
    goto check_exec
)

REM 2. Verifica Python no PATH
where python >nul 2>nul
if %errorlevel% equ 0 (
    set "PYCMD=python"
    goto check_exec
)

REM 3. Fallback para o Python 3.12 instalado
if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" (
    set "PYCMD=%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
    goto check_exec
)

echo [ERRO] Python não foi localizado no sistema.
pause
exit /b 1

:check_exec
echo [+] Utilizando Python: %PYCMD%

REM Verifica se os módulos obrigatórios estão instalados
"%PYCMD%" -c "import reportlab, flask, pandas, openpyxl, pdfplumber" >nul 2>nul
if %errorlevel% neq 0 (
    echo [!] Instalando dependencias necessarias (requirements.txt)...
    "%PYCMD%" -m pip install -r "%~dp0requirements.txt" --user
)

REM Verifica se o servidor já está ativo na porta 5000
netstat -ano | findstr /R /C:":5000 *LISTENING" >nul 2>nul
if %errorlevel% equ 0 (
    echo [+] O sistema ja esta ativo e rodando na porta 5000!
    echo [+] Abrindo o navegador em http://127.0.0.1:5000 ...
    start http://127.0.0.1:5000
    timeout /t 3 >nul
    exit /b 0
)

echo [+] Inicializando banco de dados SQLite e servidor...
echo.

REM Inicializar banco
"%PYCMD%" database.py

REM Abrir navegador após 2 segundos em segundo plano
start "" cmd /c "timeout /t 2 >nul & start http://127.0.0.1:5000"

REM Iniciar Flask
"%PYCMD%" app.py

pause
