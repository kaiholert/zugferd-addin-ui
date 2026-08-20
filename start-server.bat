@echo off
chcp 65001 > nul
title ZUGFeRD Server

echo ==================================================
echo   ZUGFeRD E-Rechnungs-Server
echo ==================================================
echo.

cd /d "%~dp0zugferd-server"

where node > nul 2>&1
if errorlevel 1 (
    echo [FEHLER] Node.js nicht gefunden!
    echo Bitte Node.js von https://nodejs.org herunterladen.
    pause
    exit /b 1
)

if not exist "node_modules\" (
    echo [INFO] Installiere Abhaengigkeiten...
    npm install
    if errorlevel 1 (
        echo [FEHLER] npm install fehlgeschlagen.
        pause
        exit /b 1
    )
    echo.
)

if not exist "config.json" (
    echo [FEHLER] config.json nicht gefunden!
    echo Bitte config.json mit Firmendaten anlegen.
    pause
    exit /b 1
)

echo [OK] Starte Server auf http://127.0.0.1:3737 ...
echo      Fenster offen lassen - Beenden mit Strg+C
echo.
node src/server.js

echo.
echo [INFO] Server beendet.
pause
