@echo off
setlocal EnableExtensions
cd /d "%~dp0"
title Nomina de Docentes - Compilar EXE

rem ============================================================
rem  Compila "NOMINA DOCENTES.exe" para usuarios SIN Python.
rem  Requiere haber ejecutado generar.bat una vez (crea .venv).
rem  Resultado: carpeta dist\NOMINA DOCENTES.exe
rem ============================================================

if not exist ".venv\Scripts\python.exe" (
    echo  [ERROR] No existe el entorno .venv.
    echo          Ejecuta primero generar.bat y vuelve a intentar.
    pause
    exit /b 1
)

set "VPY=.venv\Scripts\python.exe"

echo  Instalando PyInstaller...
"%VPY%" -m pip install pyinstaller --quiet

echo.
echo  Compilando EXE - esto tarda unos minutos...
"%VPY%" -m PyInstaller --onefile --noconsole --clean ^
    --name "NOMINA DOCENTES" ^
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
echo  [OK] EXE generado: dist\NOMINA DOCENTES.exe
echo       Copia ese archivo a cualquier PC (sin Python) y ejecutalo con doble clic.
echo.
pause
