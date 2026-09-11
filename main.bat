@echo off
chcp 65001 > nul
setlocal enabledelayedexpansion

REM --- 1. Definición de Variables ---
SET "PYTHON_EXE=py.exe"
SET "PROYECTO_DIR=C:\Users\eliseo_lopezp\proyecto1"
SET "SCRIPT=main.py"
SET "LOG_DIR=%PROYECTO_DIR%\logs"

REM --- 2. Preparación ---
cd /d "%PROYECTO_DIR%"
if not exist "%LOG_DIR%" mkdir "%LOG_DIR%"

REM --- 3. Obtención de fecha robusta usando PowerShell ---
REM Esto asegura el formato YYYYMMDD sin importar la configuración regional.
for /f "delims=" %%a in ('powershell -NoProfile -Command "Get-Date -Format 'yyyyMMdd'"') do set "DATE_STR=%%a"
set "LOG_FILE=%LOG_DIR%\main_%DATE_STR%.txt"

REM --- 4. Ejecución ---
echo ========================
echo Inicio: %date% %time%
echo Ejecutando script...
echo Logs guardados en: %LOG_FILE%

REM Ejecuta el script, muestra la salida en pantalla y la guarda en el log.
REM Se usa Tee-Object de PowerShell para duplicar la salida.
powershell -NoProfile -Command "& '%PYTHON_EXE%' -u '%SCRIPT%' 'AUTO' 2>&1 | Tee-Object -FilePath '%LOG_FILE%' -Append"

echo Fin: %date% %time%
echo ========================
echo Proceso finalizado.
echo El log completo se encuentra en: %LOG_FILE%

exit /b