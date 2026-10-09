@echo off
title 46 BPM - Servidor de Producao (Waitress)
color 0A
cd /d "%~dp0"

echo =================================================================
echo   46 BATALHAO DE POLICIA MILITAR - BAGE / RS
echo   Iniciando Servidor Web em MODO PRODUCAO (Waitress WSGI)...
echo =================================================================
echo.

python wsgi.py
pause
