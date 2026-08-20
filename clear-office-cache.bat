@echo off
chcp 65001 > nul
title Office Add-in Cache leeren (vollstaendig)

echo ==================================================
echo   Office Add-in Cache leeren (vollstaendig)
echo ==================================================
echo.
echo Bitte alle offenen Word-Dokumente speichern.
echo Word, Excel und PowerPoint werden jetzt geschlossen.
echo.
pause

echo.
echo [1/5] Schliesse Office-Anwendungen...
taskkill /IM WINWORD.EXE  /F > nul 2>&1
taskkill /IM EXCEL.EXE    /F > nul 2>&1
taskkill /IM POWERPNT.EXE /F > nul 2>&1
taskkill /IM MSEDGE.EXE   /F > nul 2>&1
taskkill /IM msedgewebview2.exe /F > nul 2>&1
timeout /t 3 > nul
echo       OK

echo.
echo [2/5] Leere WEF Add-in Cache...
set "WEF=%LOCALAPPDATA%\Microsoft\Office\16.0\Wef"
if exist "%WEF%" (
    del /s /f /q "%WEF%\*.*" > nul 2>&1
    for /d %%d in ("%WEF%\*") do rd /s /q "%%d" > nul 2>&1
    echo       OK: %WEF%
) else (
    echo       INFO: Kein WEF-Ordner gefunden
)

echo.
echo [3/5] Leere WebView2 User Data Cache...
set "WV2=%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\webview2"
if exist "%WV2%" (
    rd /s /q "%WV2%" > nul 2>&1
    echo       OK: %WV2%
) else (
    echo       INFO: Kein WebView2-Ordner gefunden
)

echo.
echo [4/5] Leere temporaere Internet-Dateien fuer Office...
set "IETMP=%LOCALAPPDATA%\Microsoft\Windows\INetCache\Content.IE5"
if exist "%IETMP%" (
    del /s /f /q "%IETMP%\*.*" > nul 2>&1
    echo       OK: IE-Cache geleert
) else (
    echo       INFO: Kein IE-Cache gefunden
)

:: WebView2-EBWebView Cache (wird von Office fuer Add-ins genutzt)
set "EBWV=%LOCALAPPDATA%\Microsoft\Edge\User Data\Default\Cache"
if exist "%EBWV%" (
    del /s /f /q "%EBWV%\*.*" > nul 2>&1
    echo       OK: Edge WebView Cache geleert
)

echo.
echo [5/5] Setze Add-in Versions-Flag zurueck...
reg delete "HKCU\Software\Microsoft\Office\16.0\WEF" /f > nul 2>&1
echo       OK

echo.
echo ==================================================
echo   Cache vollstaendig geleert!
echo ==================================================
echo.
echo   Naechste Schritte:
echo   1. Word neu starten
echo   2. Rechnungsvorlage_ZUGFeRD_ContentControls.docx oeffnen
echo   3. Start ^> ZUGFeRD Export
echo   4. Task Pane sollte jetzt aktuell sein
echo.
echo   TIPP: Zur Verifikation in der Task Pane die URL
echo   pruefen - rechtsklick in die Task Pane und
echo   "Inspect" oder F12 waehlen (falls verfuegbar)
echo ==================================================
echo.
pause
