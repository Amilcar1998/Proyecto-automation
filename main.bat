@echo off
chcp 65001 > nul
setlocal enabledelayedexpansion

REM =========================================================
REM  EJECUCIÓN AUTOMÁTICA TRANSFERENCIAS WMS-INFOR
REM =========================================================

SET "PYTHON_EXE=C:\Users\eliseo_lopezp\AppData\Local\Programs\Python\Python313\python.exe"
SET "PROYECTO_DIR=C:\Users\eliseo_lopezp\proyecto1"
SET "SCRIPT=main.py"

cd /d "%PROYECTO_DIR%"

REM Ejecutar SOLO mostrando logs en pantalla (sin archivo)
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
"$ErrorActionPreference='SilentlyContinue'; ^
Write-Host '========================' -ForegroundColor Green; ^
Write-Host 'Inicio: %DATE% %TIME%'; ^
& '%PYTHON_EXE%' '%SCRIPT%' AUTO 2>&1; ^
Write-Host 'Fin: %DATE% %TIME%'; ^
Write-Host '========================' -ForegroundColor Green"

exit
