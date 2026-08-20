@echo off
chcp 65001 > nul
title ZUGFeRD Add-in deinstallieren

echo ==================================================
echo   ZUGFeRD Add-in - Deinstallation
echo ==================================================
echo.

net session > nul 2>&1
if errorlevel 1 (
    echo [FEHLER] Bitte als Administrator ausfuehren.
    pause
    exit /b 1
)

echo Entferne Netzwerkfreigabe "ZUGFeRDCatalog"...
net share ZUGFeRDCatalog /delete > nul 2>&1

echo.
echo Fertig. Bitte zusaetzlich in Word den Katalog manuell
echo entfernen unter:
echo Datei ^> Optionen ^> Trust Center ^> Einstellungen
echo ^> Vertrauenswuerdige Add-In-Kataloge
echo.
pause
