@echo off
chcp 65001 > nul
title ZUGFeRD - HTTPS Zertifikat installieren

echo ==================================================
echo   HTTPS Entwicklungszertifikat installieren
echo ==================================================
echo.
echo Dieses Skript installiert ein selbstsigniertes
echo Zertifikat fuer localhost, damit Word die Add-in
echo Task Pane ohne Sicherheitswarnung laden kann.
echo.
echo HINWEIS: Windows fragt gleich, ob du das Zertifikat
echo in den Vertrauensspeicher aufnehmen moechtest.
echo Bitte mit "Ja" bestaetigen.
echo.
pause

echo.
echo [1/4] Pruefe Node.js / npx...
where npx > nul 2>&1
if errorlevel 1 (
    echo       [FEHLER] npx nicht gefunden.
    echo       Bitte Node.js von https://nodejs.org installieren.
    pause
    exit /b 1
)
echo       OK

echo.
echo [2/4] Installiere office-addin-dev-certs und Zertifikat...
echo       (Erster Aufruf laedt das Tool herunter - bitte warten)
echo.
call npx office-addin-dev-certs install --machine
if errorlevel 1 (
    echo.
    echo       [FEHLER] Versuch 1 fehlgeschlagen - versuche ohne --machine...
    call npx office-addin-dev-certs install
    if errorlevel 1 (
        echo.
        echo       [FEHLER] Beide Versuche fehlgeschlagen.
        echo       Moegliche Ursachen:
        echo       - Keine Internetverbindung fuer Download
        echo       - "Ja"-Bestaetigung im Windows-Dialog verpasst
        echo       - npm-Registry durch Firmen-Proxy blockiert
        pause
        exit /b 1
    )
)

echo.
echo [3/4] Pruefe Zertifikat-Dateien...
set "CERT=%USERPROFILE%\.office-addin-dev-certs\localhost.crt"
set "KEY=%USERPROFILE%\.office-addin-dev-certs\localhost.key"

if not exist "%CERT%" (
    echo       [FEHLER] Zertifikat fehlt: %CERT%
    pause
    exit /b 1
)
echo       OK: %CERT%
if not exist "%KEY%" (
    echo       [FEHLER] Key fehlt: %KEY%
    pause
    exit /b 1
)
echo       OK: %KEY%

echo.
echo [4/4] Verifiziere im Windows-Vertrauensspeicher...
certutil -store Root localhost > nul 2>&1
if errorlevel 1 (
    certutil -store CA localhost > nul 2>&1
    if errorlevel 1 (
        echo       [WARNUNG] Zertifikat nicht im Vertrauensspeicher.
        echo       Bitte Skript neu starten und Windows-Dialog
        echo       mit "Ja" bestaetigen.
    ) else (
        echo       OK: Im CA-Store gefunden
    )
) else (
    echo       OK: Im Root-Store gefunden
)

echo.
echo ==================================================
echo   Fertig! Naechste Schritte:
echo.
echo   1. start-taskpane-server.bat neu starten
echo   2. Word neu starten
echo   3. Add-in erneut oeffnen
echo.
echo   Falls Fehler bleibt: Im Browser https://localhost:3000
echo   oeffnen, Zertifikat dort einmalig akzeptieren,
echo   dann Word Add-in erneut versuchen.
echo ==================================================
echo.
pause
