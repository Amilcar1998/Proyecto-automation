@echo off
echo ===================================================
echo Instalando dependencias del proyecto Automation
echo ===================================================

echo.
echo Verificando instalacion de Python...
python --version >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo [ERROR] Python no esta instalado o no esta en las variables de entorno (PATH).
    echo Por favor instala Python (agrega al PATH) antes de continuar.
    pause
    exit /b
)

echo.
echo Actualizando pip a la ultima version...
python -m pip install --upgrade pip

echo.
echo Instalando paquetes necesarios...
pip install pandas>=1.3.0
pip install pyodbc>=4.0.30
pip install python-docx>=0.8.11
pip install openpyxl>=3.0.9
pip install pywin32>=300
pip install requests>=2.25.0
pip install xlwt>=1.3.0
pip install paramiko
pip install selenium

echo.
echo ===================================================
echo Instalacion completada exitosamente.
echo ===================================================
pause
