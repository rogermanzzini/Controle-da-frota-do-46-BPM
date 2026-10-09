@echo off
title 46 BPM - Publicar Aplicativo na Internet (Waitress + Cloudflare)
color 0B
cd /d "%~dp0"

echo =================================================================
echo   46 BATALHAO DE POLICIA MILITAR - BAGE / RS
echo   Sistema de Controle de Frota e Abastecimentos
echo =================================================================
echo.
echo 1. Iniciando Servidor WSGI de Producao (Waitress)...
start "Servidor WSGI - 46 BPM" python wsgi.py

timeout /t 3 /nobreak >nul

echo 2. Iniciando Tunel Seguro HTTPS (Cloudflare)...
echo.
echo =================================================================
echo   Copie o endereco HTTPS que termina com ".trycloudflare.com"
echo   Esse sera o link publico para acessar o aplicativo de qualquer
echo   celular, tablet ou computador na internet!
echo =================================================================
echo.

cloudflared.exe tunnel --url http://127.0.0.1:5000
pause
