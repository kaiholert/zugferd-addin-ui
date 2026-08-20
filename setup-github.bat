@echo off
chcp 65001 > nul
title ZUGFeRD - GitHub Pages Setup (ueberholt)

echo ==================================================
echo   ZUGFeRD Add-in - GitHub Pages Einrichtung
echo ==================================================
echo.
echo Dieses Skript wird NICHT MEHR benoetigt und sollte
echo NICHT ausgefuehrt werden.
echo.
echo Grund: Es wurde urspruenglich fuer eine alte Projekt-
echo struktur geschrieben, in der word-addin\ sein EIGENES
echo Git-Repository war (word-addin\.git). Seit der Repo-
echo Erweiterung liegt .git jetzt im Projekt-Root und deckt
echo den gesamten Ordner zugferd-addin\ ab. Wuerde man dieses
echo Skript trotzdem starten, wuerde es faelschlicherweise
echo ein zweites, verschachteltes Git-Repository in word-addin\
echo anlegen und die aktuelle Struktur beschaedigen.
echo.
echo Setup ist bereits erledigt:
echo   - Repository: https://github.com/kaiholert/zugferd-addin-ui
echo   - Pages-URL:  https://kaiholert.github.io/zugferd-addin-ui
echo   - manifest.xml enthaelt bereits die finale Pages-URL
echo     (kein Platzhalter mehr).
echo.
echo Deploy laeuft jetzt automatisch ueber GitHub Actions:
echo   .github\workflows\deploy-pages.yml
echo   (deployt word-addin\ bei jedem Push nach main)
echo.
echo Falls GitHub Pages jemals neu eingerichtet werden muss:
echo   1. Auf GitHub: Settings ^> Pages
echo   2. "Build and deployment" ^> Source: "GitHub Actions"
echo   3. Push nach main loest den Workflow automatisch aus
echo      (oder manuell: Actions-Tab ^> Deploy GitHub Pages ^>
echo      Run workflow)
echo.
pause
