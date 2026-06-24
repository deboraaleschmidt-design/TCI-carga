@echo off
title Copias fonte TCI — Chrome
cd /d "%~dp0"
call "%~dp0_abrir-chrome.bat" "https://docs.google.com/spreadsheets/d/1dxK-p-blCB8TBRAfCBRWhjXHNiYBpGExkFAQE9PQa4M/edit"
timeout /t 1 /nobreak >nul
call "%~dp0_abrir-chrome.bat" "https://docs.google.com/spreadsheets/d/1K1SzwTSSCri57o0U1A58EWT6cA4s0-XFjC5f5PyYR-A/edit"
timeout /t 1 /nobreak >nul
call "%~dp0_abrir-chrome.bat" "https://docs.google.com/spreadsheets/d/1ixrY1sw9XovlNcLIcSX2c6-Ij-Q4qyDbHhUOeH6SWII/edit"
