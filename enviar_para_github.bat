@echo off
title 46 BPM - Publicar no GitHub
color 0A
cd /d "%~dp0"

echo =================================================================
echo   46 BATALHAO DE POLICIA MILITAR - BAGE / RS
echo   Enviar Projeto para o GitHub (Para Deploy no Render / Nuvem)
echo =================================================================
echo.
echo Se voce ainda nao criou o repositorio no GitHub:
echo   1. Acesse: https://github.com/new
echo   2. Escolha o nome (ex: bpm46-frota) e marque como Privado ou Publico
echo   3. Copie o link HTTPS gerado (ex: https://github.com/seu-usuario/bpm46-frota.git)
echo.
set /p REPO_URL="Cole o link do seu repositorio do GitHub aqui: "

if "%REPO_URL%"=="" (
    echo Nenhum link fornecido. Operacao cancelada.
    pause
    exit /b
)

echo.
echo Configurando repositorio remoto...
tools\git\cmd\git.exe remote remove origin 2>nul
tools\git\cmd\git.exe remote add origin %REPO_URL%
tools\git\cmd\git.exe branch -M main

echo Enviando arquivos para o GitHub...
tools\git\cmd\git.exe push -u origin main

echo.
if %ERRORLEVEL% EQU 0 (
    echo =================================================================
    echo   SUCESSO! O projeto foi enviado para o GitHub!
    echo   Agora acesse https://render.com para conectar e publicar online.
    echo =================================================================
) else (
    echo =================================================================
    echo   Ocorreu um erro no envio. Verifique o link e seu login no GitHub.
    echo =================================================================
)
pause
