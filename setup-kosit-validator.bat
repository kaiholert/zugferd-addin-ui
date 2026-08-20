@echo off
setlocal

echo ===================================================
echo  ZUGFeRD - KoSIT Validator Setup
echo ===================================================
echo.
echo Dieses Skript laedt den KoSIT Validator (Java-Tool)
echo und die ZUGFeRD-Validierungsregeln herunter, damit
echo der Server erzeugte XML-Rechnungen lokal pruefen kann.
echo.
echo Benoetigt wird eine installierte Java-Laufzeit (JRE) 11+.
echo.

java -version >nul 2>&1
if errorlevel 1 (
  echo [WARNUNG] Java wurde nicht gefunden.
  echo Bitte zuerst ein JRE installieren, z.B. https://adoptium.net
  echo Danach dieses Skript erneut ausfuehren.
  pause
  exit /b 1
)

set TOOLS_DIR=%~dp0zugferd-server\tools\kosit
if not exist "%TOOLS_DIR%" mkdir "%TOOLS_DIR%"

echo.
echo [1/3] Lade KoSIT Validator (validator-1.6.2-standalone.jar)...
curl -L -o "%TOOLS_DIR%\validator-1.6.2-standalone.jar" "https://github.com/itplr-kosit/validator/releases/download/v1.6.2/validator-1.6.2-standalone.jar"
if errorlevel 1 (
  echo [FEHLER] Download des Validators fehlgeschlagen.
  pause
  exit /b 1
)

echo.
echo [2/3] Lade ZUGFeRD-Validierungsregeln (scenarios.xml + Schema/Schematron)...
curl -L -o "%TOOLS_DIR%\zugferd-config.zip" "https://github.com/LandrixSoftware/validator-configuration-zugferd/archive/refs/heads/main.zip"
if errorlevel 1 (
  echo [FEHLER] Download der Validierungsregeln fehlgeschlagen.
  pause
  exit /b 1
)

echo.
echo [3/3] Entpacke Validierungsregeln...
powershell -NoProfile -Command "Expand-Archive -Path '%TOOLS_DIR%\zugferd-config.zip' -DestinationPath '%TOOLS_DIR%' -Force"
del "%TOOLS_DIR%\zugferd-config.zip"

set CONFIG_DIR=
for /d %%D in ("%TOOLS_DIR%\validator-configuration-zugferd-*") do set CONFIG_DIR=%%D

echo.
echo ===================================================
echo  Fertig.
echo ===================================================
echo.
echo Bitte in zugferd-server\config.json den Block "kositValidation"
echo wie folgt anpassen (Pfade ggf. mit doppelten Backslashes):
echo.
echo   "kositValidation": {
echo     "enabled": true,
echo     "javaPath": "java",
echo     "validatorJar": "%TOOLS_DIR%\validator-1.6.2-standalone.jar",
echo     "scenariosXml": "%CONFIG_DIR%\scenarios.xml",
echo     "timeoutMs": 30000,
echo     "blockOnInvalid": false
echo   }
echo.
echo Danach start-server.bat neu starten (bzw. den laufenden
echo Server beenden und neu starten), damit die Aenderung greift.
echo.
pause
