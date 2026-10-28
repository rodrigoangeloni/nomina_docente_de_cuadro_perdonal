@echo off
setlocal EnableExtensions
cd /d "%~dp0"
title Nomina de Docentes - Publicar version

rem ============================================================
rem  Publica una nueva version:
rem   1) actualiza VERSION en generar_nomina.py
rem   2) compila dist\NOMINA DOCENTES.exe
rem   3) crea el tag vX.Y.Z, lo sube y crea el Release con el EXE
rem  Uso: publicar.bat 1.0.1
rem ============================================================

if "%~1"=="" (
    echo  Uso: publicar.bat VERSION    ej: publicar.bat 1.0.1
    pause
    exit /b 1
)
set "V=%~1"

if not exist ".venv\Scripts\python.exe" (
    echo  [ERROR] Falta el entorno .venv. Ejecuta generar.bat una vez primero.
    pause
    exit /b 1
)
set "VPY=.venv\Scripts\python.exe"

echo  [1/4] Actualizando VERSION a %V% ...
"%VPY%" bump_version.py %V%
if errorlevel 1 (
    pause
    exit /b 1
)

echo  [2/4] Compilando EXE ...
"%VPY%" -m pip install pyinstaller --quiet
"%VPY%" -m PyInstaller --onefile --noconsole --clean --name "NOMINA DOCENTES" --collect-all pypdfium2 generar_nomina.py
if errorlevel 1 (
    echo  [ERROR] Fallo la compilacion.
    pause
    exit /b 1
)

echo  [3/4] Versionando con git ...
git add generar_nomina.py
git commit -m "Version %V%"
git tag -f v%V%
git push origin HEAD --tags
if errorlevel 1 (
    echo  [ERROR] No se pudo hacer push. Verifica tu conexion y tu sesion de GitHub.
    pause
    exit /b 1
)

echo  [4/4] Creando el Release en GitHub ...
where gh >nul 2>nul
if %errorlevel%==0 (
    gh release create v%V% "dist\NOMINA DOCENTES.exe" --title "NOMINA DOCENTES v%V%" --notes "Version %V% - nomina de docentes generada desde cuadro de personal"
    if errorlevel 1 goto :release_fallo
    goto :fin_ok
) else (
    goto :sin_gh
)

:release_fallo
echo  [AVISO] No se pudo crear el release con gh. Subelo manualmente.
goto :manual

:sin_gh
echo  [INFO] GitHub CLI no esta instalado. Instalalo desde https://cli.github.com/
echo         o sube el EXE manualmente:
:manual
echo         1) Abre https://github.com/rodrigoangeloni/nomina_docente_de_cuadro_perdonal/releases
echo         2) Crea/edita el release del tag v%V%
echo         3) Adjunta el archivo dist\NOMINA DOCENTES.exe

:fin_ok
echo.
echo  [OK] Version v%V% publicada. Los usuarios con el EXE recibiran
echo       el aviso de actualizacion al ejecutarlo.
echo.
pause
