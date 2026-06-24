@echo off
title Criar PAINEL CARGA TCI v2 — Chrome
cd /d "%~dp0"

echo.
echo  PAINEL CARGA TCI v2 — tudo no Chrome
echo  ------------------------------------
echo  1. Planilha nova no Google
echo  2. Setup (passo a passo + codigo para copiar)
echo.
pause

call "%~dp0_abrir-chrome.bat" "https://docs.google.com/spreadsheets/create"
timeout /t 2 /nobreak >nul
call "%~dp0_abrir-chrome-arquivo.bat" "%~dp0SETUP-PAINEL-CARGA.html"

echo.
echo  Na planilha nova: PAINEL CARGA TCI v2
echo  Cole o codigo do setup no Apps Script
echo.
