@echo off
setlocal EnableExtensions
cd /d "%~dp0"
title Nomina de Docentes - Configurador inicial

rem ============================================================
rem  Prepara todo en una PC nueva:
rem   1) Verifica si hay Python; si no, lo instala (sin admin)
rem   2) Crea el entorno virtual .venv
rem   3) Instala las dependencias (pdfplumber, openpyxl)
rem   4) Abre el programa para generar la nomina
rem ============================================================

set "PY="

rem --- 1. Buscar un Python ya instalado ---
where python >nul 2>nul && set "PY=python"
if not defined PY (
    where py >nul 2>nul && set "PY=py -3"
)

rem --- 2. Si no hay Python, instalarlo ---
if not defined PY call :instalar_python

if not defined PY (
    echo.
    echo  [ERROR] No se pudo instalar Python automaticamente.
    echo          Descargalo manualmente desde https://www.python.org/downloads/
    echo          y marca la casilla "Add Python to PATH" al instalarlo.
    echo.
    pause
    exit /b 1
)

echo  [OK] Python encontrado: %PY%

rem --- 3. Crear el entorno virtual la primera vez ---
if not exist ".venv\Scripts\python.exe" (
    echo  Creando entorno virtual .venv ...
    %PY% -m venv .venv
    if errorlevel 1 (
        echo  [ERROR] No se pudo crear el entorno virtual.
        pause
        exit /b 1
    )
)

set "VPY=.venv\Scripts\python.exe"

rem --- 4. Instalar dependencias ---
echo  Verificando dependencias - esto puede tardar unos minutos la primera vez...
"%VPY%" -m pip install --upgrade pip --quiet
"%VPY%" -m pip install -r requirements.txt --quiet
if errorlevel 1 (
    echo  [ERROR] Fallo la instalacion de dependencias. Revisa tu conexion a internet.
    pause
    exit /b 1
)

echo  [OK] Todo listo. Abriendo el programa...
echo.
"%VPY%" generar_nomina.py

endlocal
exit /b 0

rem ============================================================
rem  Subrutina: instalar Python sin permisos de administrador
rem ============================================================
:instalar_python
echo  Python no esta instalado. Intentando instalarlo automaticamente...

rem  Opcion A: winget (Windows 10/11 actualizados)
where winget >nul 2>nul
if %errorlevel%==0 (
    echo  Descargando Python con winget...
    winget install -e --id Python.Python.3.12 --silent --accept-package-agreements --accept-source-agreements
    if not errorlevel 1 goto :buscar_python_instalado
)

rem  Opcion B: instalador oficial en modo silencioso (solo usuario)
echo  Descargando instalador oficial de Python...
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "try { [Net.ServicePointManager]::SecurityProtocol='Tls12'; Invoke-WebRequest -Uri 'https://www.python.org/ftp/python/3.12.8/python-3.12.8-amd64.exe' -OutFile \"$env:TEMP\python-setup.exe\" -UseBasicParsing; 'OK' } catch { $_.Exception.Message }" | findstr /C:"OK" >nul
if errorlevel 1 (
    echo  [ERROR] No se pudo descargar el instalador. Verifica tu conexion a internet.
    pause
    exit /b 1
)
echo  Instalando Python (tarda unos minutos)...
"%TEMP%\python-setup.exe" /quiet InstallAllUsers=0 PrependPath=1 Include_test=0
del "%TEMP%\python-setup.exe" >nul 2>nul

:buscar_python_instalado
rem  Buscar en las rutas tipicas de instalacion por usuario
for /d %%D in ("%LOCALAPPDATA%\Programs\Python\Python3*") do set "PY=%%D\python.exe"
for /d %%D in ("%ProgramFiles%\Python3*") do if not defined PY set "PY=%%D\python.exe"
if defined PY goto :eof

echo  [ERROR] Python se instalo pero no se encontro. Cierra esta ventana,
echo          abre una nueva y vuelve a ejecutar generar.bat
pause
exit /b 1
