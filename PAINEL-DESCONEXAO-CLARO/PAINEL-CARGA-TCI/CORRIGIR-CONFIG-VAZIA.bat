@echo off
title CONFIG vazia — ajuda Chrome
cd /d "%~dp0"
call "%~dp0_abrir-chrome.bat" "https://docs.google.com/spreadsheets/d/1YAYbViLyvQoIEkeBizwOHSKMc9sxkE0Rl8L8bUGro9M/edit#gid=1399571813"
timeout /t 2 /nobreak >nul
call "%~dp0_abrir-chrome-arquivo.bat" "%~dp0CONFIG-VAZIA-LEIA-ME.txt"
timeout /t 1 /nobreak >nul
call "%~dp0_abrir-chrome-arquivo.bat" "%~dp0CONFIG-COLAR.csv"
