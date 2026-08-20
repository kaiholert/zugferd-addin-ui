@echo off
chcp 65001 > nul
title ZUGFeRD Add-in installieren

echo ==================================================
echo   ZUGFeRD Add-in - Installation
echo ==================================================
echo.

net session > nul 2>&1
if errorlevel 1 (
    echo [FEHLER] Bitte als Administrator ausfuehren.
    echo Rechtsklick auf diese Datei ^> "Als Administrator ausfuehren"
    pause
    exit /b 1
)

set "CATALOG_DIR=%~dp0word-addin-catalog"

echo [1/2] Pruefe Katalog-Ordner...
if not exist "%CATALOG_DIR%" mkdir "%CATALOG_DIR%"
echo       OK: %CATALOG_DIR%

echo.
echo [2/2] Kopiere Manifest...
copy /Y "%~dp0word-addin\manifest.xml" "%CATALOG_DIR%\ZUGFeRD-Export.xml" > nul
echo       OK: ZUGFeRD-Export.xml

:: Laufwerksbuchstabe und Pfad ohne Doppelpunkt fuer administrative Freigabe ermitteln
:: Beispiel: C:\Users\Name\zugferd-addin\word-addin-catalog
::        -> \\localhost\C$\Users\Name\zugferd-addin\word-addin-catalog
set "DRIVE_LETTER=%CATALOG_DIR:~0,1%"
set "PATH_NO_DRIVE=%CATALOG_DIR:~2%"
set "ADMIN_UNC=\\localhost\%DRIVE_LETTER%$%PATH_NO_DRIVE%"

echo.
echo ==================================================
echo   Installation abgeschlossen!
echo ==================================================
echo.
echo   Dieser PC hat KEINE eigene Netzwerkfreigabe noetig.
echo   Wir nutzen die in Windows bereits vorhandene
echo   administrative Freigabe ^(%DRIVE_LETTER%$^).
echo.
echo   Naechster Schritt - in Word eintragen:
echo   Datei ^> Optionen ^> Trust Center
echo   ^> Einstellungen fuer das Trust Center
echo   ^> Vertrauenswuerdige Add-In-Kataloge
echo.
echo   Katalog-URL ^(genau so eintragen^):
echo.
echo   %ADMIN_UNC%
echo.
echo   Dann: Katalog hinzufuegen ^> Haekchen "Im Menue anzeigen"
echo   ^> OK ^> OK ^> Word KOMPLETT schliessen und neu starten
echo.
echo   Danach: Start ^> Weitere Add-Ins ^> Freigegebener Ordner
echo           ^> ZUGFeRD Export
echo.
echo   Falls das nicht funktioniert ^(z.B. weil die IT-Abteilung
echo   den Zugriff auf administrative Freigaben per Richtlinie
echo   gesperrt hat^), wende dich bitte an deine IT-Abteilung -
echo   in diesem Fall ist eine Konfiguration auf Domaenenebene
echo   noetig, die nicht per Skript umgangen werden sollte.
echo ==================================================
echo %ADMIN_UNC% > "%~dp0unc-pfad.txt"
echo.
pause
