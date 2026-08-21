@echo off
title ZUGFeRD Schnellerfassung

set URL=http://127.0.0.1:3737/quick-invoice.html

echo Pruefe ob der ZUGFeRD-Server laeuft...
curl -s -o nul -w "%%{http_code}" http://127.0.0.1:3737/ping > "%TEMP%\zugferd-ping.txt" 2>nul
set /p PING_CODE=<"%TEMP%\zugferd-ping.txt"
del "%TEMP%\zugferd-ping.txt" >nul 2>&1

if "%PING_CODE%"=="200" (
  echo Server laeuft bereits.
) else (
  echo Server nicht erreichbar - starte start-server.bat...
  start "ZUGFeRD Server" "%~dp0start-server.bat"
  echo Warte auf Serverstart...
  timeout /t 4 /nobreak > nul
)

echo Oeffne Schnellerfassung im Browser...
start "" "%URL%"
