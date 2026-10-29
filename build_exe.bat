@echo off
setlocal EnableExtensions
cd /d "%~dp0"
title Nomina de Docentes - Compilar EXE

rem ============================================================
rem  Compila "NominaDocentes.exe" para usuarios SIN Python.
rem  Requiere haber ejecutado generar.bat una vez (crea .venv).
rem  Resultado: carpeta dist\NominaDocentes.exe
rem ============================================================

set "NOMBRE_EXE=NominaDocentes"

if not exist ".venv\Scripts\python.exe" (
    echo  [ERROR] No existe el entorno .venv.
    echo          Ejecuta primero generar.bat y vuelve a intentar.
    pause
    exit /b 1
)

set "VPY=.venv\Scripts\python.exe"

echo  Instalando PyInstaller...
"%VPY%" -m pip install pyinstaller --quiet

echo  Generando metadatos de version...
"%VPY%" generar_version_info.py
if errorlevel 1 (
    echo  [ERROR] No se pudieron generar los metadatos del EXE.
    pause
    exit /b 1
)

echo.
echo  Compilando EXE - esto tarda unos minutos...
"%VPY%" -m PyInstaller --onefile --noconsole --clean ^
    --name "%NOMBRE_EXE%" ^
    --version-file "version_info.txt" ^
    --icon "favicon.ico" ^
    --add-data "favicon.ico;." ^
    --collect-all pypdfium2 ^
    generar_nomina.py
if errorlevel 1 (
    echo  [ERROR] Fallo la compilacion.
    pause
    exit /b 1
)

echo.
echo  [OK] EXE generado: dist\%NOMBRE_EXE%.exe
echo       Copia ese archivo a cualquier PC (sin Python) y ejecutalo con doble clic.
echo.
pause
