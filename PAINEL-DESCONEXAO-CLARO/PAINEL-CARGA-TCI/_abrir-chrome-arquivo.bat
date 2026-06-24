@echo off
REM Abre arquivo local no Google Chrome (caminho com espacos — sem file://)
set "ARQ=%~1"
if not exist "%ARQ%" (
  echo Arquivo nao encontrado:
  echo   %ARQ%
  pause
  exit /b 1
)
set "CHROME=%ProgramFiles%\Google\Chrome\Application\chrome.exe"
if not exist "%CHROME%" set "CHROME=%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe"
if not exist "%CHROME%" set "CHROME=%LocalAppData%\Google\Chrome\Application\chrome.exe"
if exist "%CHROME%" (
  start "" "%CHROME%" "%ARQ%"
) else (
  start chrome "%ARQ%"
)
exit /b 0
