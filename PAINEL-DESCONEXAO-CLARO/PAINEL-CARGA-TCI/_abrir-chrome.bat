@echo off
REM Helper — abre URL no Google Chrome
set "URL=%~1"
if "%URL%"=="" exit /b 1
set "CHROME=%ProgramFiles%\Google\Chrome\Application\chrome.exe"
if not exist "%CHROME%" set "CHROME=%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe"
if not exist "%CHROME%" set "CHROME=%LocalAppData%\Google\Chrome\Application\chrome.exe"
if exist "%CHROME%" (start "" "%CHROME%" "%URL%") else (start chrome "%URL%")
exit /b 0
