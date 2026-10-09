@echo off
title 46 BPM - Enviar Projeto para o GitHub
color 0A
cd /d "%~dp0"

echo =================================================================
echo   46 BATALHAO DE POLICIA MILITAR - BAGE / RS
echo   Enviando Projeto para o GitHub Oficial do 46 BPM
echo =================================================================
echo.
echo Repositorio destino:
echo   https://github.com/rogermanzzini/Controle-da-frota-do-46-BPM.git
echo.

set REPO_URL=https://github.com/rogermanzzini/Controle-da-frota-do-46-BPM.git

echo 1. Configurando repositorio remoto...
tools\git\cmd\git.exe remote remove origin 2>nul
tools\git\cmd\git.exe remote add origin %REPO_URL%
tools\git\cmd\git.exe branch -M main

echo.
echo 2. Enviando arquivos para o GitHub...
echo (Se o GitHub pedir login, confirme na janela do seu navegador)
echo.
tools\git\cmd\git.exe push -u origin main

echo.
if %ERRORLEVEL% EQU 0 (
    echo =================================================================
    echo   SUCESSO! Todos os arquivos foram enviados para o seu GitHub!
    echo   Agora acesse https://render.com para publicar na nuvem.
    echo =================================================================
) else (
    echo =================================================================
    echo   Ocorreu um erro no envio. Verifique suas credenciais do GitHub.
    echo =================================================================
)
pause
