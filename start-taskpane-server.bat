@echo off
chcp 65001 > nul
title ZUGFeRD Task Pane Server (HTTPS :3000)
setlocal enabledelayedexpansion

echo ==================================================
echo   ZUGFeRD Task Pane - HTTPS Server
echo   Diagnose-Modus: ausfuehrlich
echo ==================================================
echo.

set "ADDIN_DIR=%~dp0word-addin"
set "CERT=%USERPROFILE%\.office-addin-dev-certs\localhost.crt"
set "KEY=%USERPROFILE%\.office-addin-dev-certs\localhost.key"
set "PORT=3000"
set "LOGFILE=%~dp0taskpane-server.log"

echo [1/6] Pruefe Node.js / npx...
where node > nul 2>&1
if errorlevel 1 (
    echo       [FEHLER] Node.js wurde nicht gefunden.
    echo       Bitte von https://nodejs.org installieren und PC neu starten.
    pause
    exit /b 1
)
for /f "tokens=*" %%v in ('node --version') do set NODEVER=%%v
echo       OK: Node.js gefunden ^(Version !NODEVER!^)

where npx > nul 2>&1
if errorlevel 1 (
    echo       [FEHLER] npx wurde nicht gefunden ^(gehoert eigentlich zu Node.js^).
    echo       Node.js-Installation evtl. beschaedigt - bitte neu installieren.
    pause
    exit /b 1
)
echo       OK: npx gefunden

echo.
echo [2/6] Pruefe Add-in-Ordner...
if not exist "%ADDIN_DIR%" (
    echo       [FEHLER] Ordner nicht gefunden: %ADDIN_DIR%
    echo       Pruefe, ob dieses Skript im richtigen Hauptordner liegt.
    pause
    exit /b 1
)
if not exist "%ADDIN_DIR%\taskpane.html" (
    echo       [FEHLER] taskpane.html fehlt in: %ADDIN_DIR%
    echo       Der Ordner ist vorhanden, aber unvollstaendig.
    pause
    exit /b 1
)
echo       OK: %ADDIN_DIR% ^(taskpane.html gefunden^)

echo.
echo [3/6] Pruefe ob Port %PORT% bereits belegt ist...
set "PORT_BUSY="
for /f "tokens=5" %%p in ('netstat -ano ^| findstr ":%PORT% " ^| findstr "LISTENING"') do (
    set "PORT_BUSY=%%p"
)
if defined PORT_BUSY (
    echo       [WARNUNG] Port %PORT% wird bereits von Prozess-ID !PORT_BUSY! verwendet.
    echo       Moeglich: Ein anderer http-server laeuft schon - das ist OK,
    echo       wenn es ein vorheriger Start dieses Skripts war.
    echo       Falls Probleme auftreten: Taskmanager oeffnen, Prozess mit
    echo       PID !PORT_BUSY! beenden, dann dieses Skript erneut starten.
    echo.
) else (
    echo       OK: Port %PORT% ist frei
)

echo.
echo [4/6] Pruefe HTTPS-Entwicklungszertifikat...
if not exist "%CERT%" (
    echo       [INFO] Zertifikat nicht gefunden unter:
    echo              %CERT%
    echo       [INFO] Installiere office-addin-dev-certs jetzt...
    echo              ^(Erster Aufruf laedt das Tool herunter - kann
    echo               1-2 Minuten dauern, bitte warten^)
    echo.
    call npx office-addin-dev-certs install
    if errorlevel 1 (
        echo.
        echo       [FEHLER] Zertifikat-Installation fehlgeschlagen.
        echo       Moegliche Ursachen:
        echo       - Keine Internetverbindung beim ersten Download
        echo       - Administratorrechte verweigert ^(Zertifikat-Store^)
        echo       - npm-Registry ^(registry.npmjs.org^) blockiert
        pause
        exit /b 1
    )
    echo.
    if not exist "%CERT%" (
        echo       [FEHLER] Installation lief durch, aber Zertifikat
        echo       trotzdem nicht gefunden unter:
        echo       %CERT%
        echo       Bitte manuell pruefen: dir "%USERPROFILE%\.office-addin-dev-certs"
        pause
        exit /b 1
    )
    echo       OK: Zertifikat wurde installiert
) else (
    echo       OK: Zertifikat vorhanden ^(%CERT%^)
)

if not exist "%KEY%" (
    echo       [FEHLER] Zertifikat-Datei da, aber KEY-Datei fehlt:
    echo       %KEY%
    pause
    exit /b 1
)
echo       OK: Key-Datei vorhanden

echo.
echo [5/6] Pruefe http-server Verfuegbarkeit...
call npx --no-install http-server --version > nul 2>&1
if errorlevel 1 (
    echo       [INFO] http-server ist noch nicht im npx-Cache.
    echo       Wird beim Start automatisch heruntergeladen ^(einmalig^).
) else (
    echo       OK: http-server bereits verfuegbar
)

echo.
echo [6/6] Starte HTTPS-Server...
echo       URL:      https://localhost:%PORT%/taskpane.html
echo       Ordner:   %ADDIN_DIR%
echo       Log:      %LOGFILE%
echo.
echo       Alle eingehenden Anfragen werden unten angezeigt.
echo       Fenster offen lassen - Beenden mit Strg+C
echo ==================================================
echo.

:: -c-1 deaktiviert Caching, --cors erlaubt Cross-Origin fuer Office.js
:: Ausgabe gleichzeitig auf Bildschirm UND in Log-Datei (ueber PowerShell Tee)
powershell -NoProfile -Command ^
  "npx http-server '%ADDIN_DIR%' -p %PORT% -S -C '%CERT%' -K '%KEY%' --cors -c-1 2>&1 | Tee-Object -FilePath '%LOGFILE%'"

echo.
echo ==================================================
echo   Server wurde beendet.
echo   Falls das unerwartet war, pruefe die Log-Datei:
echo   %LOGFILE%
echo ==================================================
pause
