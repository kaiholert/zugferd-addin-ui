@echo off
chcp 65001 > nul
title ZUGFeRD - GitHub Pages Setup
setlocal enabledelayedexpansion

echo ==================================================
echo   ZUGFeRD Add-in - GitHub Pages Einrichtung
echo ==================================================
echo.
echo Dieses Skript richtet GitHub Pages als HTTPS-Host
echo fuer die Add-in Task Pane ein.
echo.
echo Voraussetzungen:
echo   - GitHub-Konto vorhanden
echo   - Git installiert (https://git-scm.com)
echo   - Internetverbindung
echo.

echo [1/5] Pruefe Git...
where git > nul 2>&1
if errorlevel 1 (
    echo [FEHLER] Git nicht gefunden.
    echo Bitte von https://git-scm.com/download/win installieren.
    pause
    exit /b 1
)
for /f "tokens=*" %%v in ('git --version') do set GITVER=%%v
echo       OK: !GITVER!

echo.
echo [2/5] GitHub-Daten eingeben...
echo.
set /p GITHUB_USER="Dein GitHub-Benutzername: "
set /p REPO_NAME="Repository-Name (z.B. zugferd-addin-ui): "

set "GITHUB_URL=https://github.com/%GITHUB_USER%/%REPO_NAME%.git"
set "PAGES_URL=https://%GITHUB_USER%.github.io/%REPO_NAME%"

echo.
echo       Repository-URL: %GITHUB_URL%
echo       GitHub Pages:   %PAGES_URL%
echo.

echo [3/5] Bereite word-addin Ordner als Git-Repository vor...
cd /d "%~dp0word-addin"

:: Manifest URL eintragen (GITHUB_PAGES_URL ersetzen)
set "MANIFEST=manifest.xml"
set "MANIFEST_TMP=manifest_tmp.xml"

powershell -NoProfile -Command ^
  "(Get-Content '%MANIFEST%') -replace 'GITHUB_PAGES_URL', '%PAGES_URL%' | Set-Content '%MANIFEST_TMP%'"
move /Y "%MANIFEST_TMP%" "%MANIFEST%" > nul
echo       OK: manifest.xml aktualisiert mit %PAGES_URL%

:: Git initialisieren falls noch nicht vorhanden
if not exist ".git" (
    git init
    git branch -M main
)

git add -A
git commit -m "ZUGFeRD Add-in Task Pane - Initial setup"
if errorlevel 1 (
    echo.
    echo       [INFO] Nichts zu committen oder Git-User nicht konfiguriert.
    echo       Falls Git-User fehlt, einmalig ausfuehren:
    echo         git config --global user.email "deine@email.de"
    echo         git config --global user.name "Dein Name"
    echo       Dann dieses Skript erneut starten.
    pause
    exit /b 1
)
echo       OK: Commit erstellt

echo.
echo [4/5] Push zu GitHub...
echo.
echo       GitHub fragt jetzt nach deinen Zugangsdaten.
echo       Tipp: Nutze ein Personal Access Token statt Passwort:
echo       GitHub.com ^> Settings ^> Developer settings
echo       ^> Personal access tokens ^> Token (classic) ^> repo-Rechte
echo.

git remote remove origin > nul 2>&1
git remote add origin "%GITHUB_URL%"
git push -u origin main
if errorlevel 1 (
    echo.
    echo [FEHLER] Push fehlgeschlagen. Moegliche Ursachen:
    echo   - Repository existiert noch nicht auf GitHub
    echo     ^> Bitte unter github.com/new anlegen ^(ohne README^)
    echo   - Zugangsdaten falsch / Token abgelaufen
    echo   - Proxy blockiert den Zugriff
    echo.
    echo Dann dieses Skript erneut ausfuehren.
    pause
    exit /b 1
)
echo       OK: Code auf GitHub

echo.
echo [5/5] GitHub Pages aktivieren...
echo.
echo       GitHub Pages muss einmalig manuell aktiviert werden:
echo.
echo       1. Browser oeffnen:
echo          %GITHUB_URL%
echo          ^(ersetze .git am Ende durch nichts^)
echo       2. Settings ^> Pages
echo       3. Source: "Deploy from a branch"
echo       4. Branch: main / (root)
echo       5. Save klicken
echo.
echo       Pages ist nach ca. 1-2 Minuten erreichbar unter:
echo          %PAGES_URL%
echo.

echo ==================================================
echo   Setup abgeschlossen!
echo ==================================================
echo.
echo   Sobald GitHub Pages aktiv ist:
echo.
echo   1. Browser: %PAGES_URL%/taskpane.html
echo      aufrufen ^(muss laden, sonst noch warten^)
echo.
echo   2. word-addin-catalog\ZUGFeRD-Export.xml
echo      aus dem Katalog-Ordner in Word neu laden:
echo      clear-office-cache.bat ausfuehren
echo      Word neu starten ^> Add-in erneut einfuegen
echo.
echo   3. start-server.bat starten ^(bleibt lokal^)
echo.
echo   Pages-URL: %PAGES_URL%
echo %PAGES_URL% > "%~dp0github-pages-url.txt"
echo ==================================================
echo.
pause
