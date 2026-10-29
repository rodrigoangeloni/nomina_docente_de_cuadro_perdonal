@echo off
setlocal EnableExtensions EnableDelayedExpansion
cd /d "%~dp0"
title Nomina de Docentes - Publicar version

rem ============================================================
rem  Publica una nueva version:
rem   1) actualiza VERSION en generar_nomina.py (SemVer)
rem   2) genera los metadatos de Windows y compila el EXE
rem   3) crea el tag vX.Y.Z y lo sube
rem   4) crea el Release en GitHub con el EXE
rem  Uso: publicar.bat 1.1.1
rem        o doble clic para elegir patch/minor/major
rem ============================================================

set "NOMBRE_EXE=NominaDocentes"

if not exist ".venv\Scripts\python.exe" (
    echo  [ERROR] Falta el entorno .venv. Ejecuta generar.bat una vez primero.
    pause
    exit /b 1
)
set "VPY=.venv\Scripts\python.exe"

for /f "usebackq delims=" %%V in (`"%VPY%" bump_version.py --show`) do set "ACTUAL=%%V"
if not defined ACTUAL (
    echo  [ERROR] No se pudo leer la version actual de generar_nomina.py.
    pause
    exit /b 1
)
for /f "usebackq delims=" %%V in (`"%VPY%" bump_version.py --next patch`) do set "PATCH=%%V"
for /f "usebackq delims=" %%V in (`"%VPY%" bump_version.py --next minor`) do set "MINOR=%%V"
for /f "usebackq delims=" %%V in (`"%VPY%" bump_version.py --next major`) do set "MAJOR=%%V"

echo.
echo  Version actual: v%ACTUAL%
echo.

set "INTERACTIVO=0"
if not "%~1"=="" (
    set "V=%~1"
    goto :version_lista
)

set "INTERACTIVO=1"
echo  Que tipo de version vas a publicar?
echo.
echo    1. Patch  - correcciones            [%PATCH%]
echo    2. Minor  - nueva funcionalidad     [%MINOR%]
echo    3. Major  - cambio incompatible     [%MAJOR%]
echo    4. Escribir otra version
echo.

:menu
set "OPC=1"
set /p "OPC=Opcion [1-4], Enter para Patch: "
if "%OPC%"=="1" set "V=%PATCH%"
if "%OPC%"=="1" goto :version_lista
if "%OPC%"=="2" set "V=%MINOR%"
if "%OPC%"=="2" goto :version_lista
if "%OPC%"=="3" set "V=%MAJOR%"
if "%OPC%"=="3" goto :version_lista
if "%OPC%"=="4" goto :version_manual
echo  [ERROR] Opcion invalida.
goto :menu

:version_manual
set /p "V=Version nueva: "

:version_lista
set "V=%V: =%"
if /I "%V:~0,1%"=="v" set "V=%V:~1%"
if not defined V (
    echo  [ERROR] No se ingreso ninguna version.
    pause
    exit /b 1
)

"%VPY%" bump_version.py --validate "%V%"
if errorlevel 1 (
    pause
    exit /b 1
)

git rev-parse --verify --quiet v%V% >nul 2>nul
if not errorlevel 1 (
    echo  [ERROR] El tag v%V% ya existe. Elegi otra version; los tags publicados no se sobrescriben.
    pause
    exit /b 1
)

findstr /C:"## [%V%]" CHANGELOG.md >nul 2>nul
if errorlevel 1 (
    echo  [AVISO] No se encontro la seccion "## [%V%]" en CHANGELOG.md.
    if "%INTERACTIVO%"=="1" (
        set "SEGUIR="
        set /p "SEGUIR=Continuar de todos modos? S/N: "
        if /I not "!SEGUIR!"=="S" (
            echo  Publicacion cancelada. Agrega la seccion al CHANGELOG y volve a intentar.
            pause
            exit /b 0
        )
    )
)

if "%INTERACTIVO%"=="1" (
    echo.
    echo  Se va a publicar v%V% en github.com/rodrigoangeloni/nomina_docente_de_cuadro_perdonal
    set "CONFIRMAR="
    set /p "CONFIRMAR=Confirmar publicacion? S/N: "
    if /I not "!CONFIRMAR!"=="S" (
        echo  Publicacion cancelada.
        pause
        exit /b 0
    )
)

echo  [1/4] Actualizando VERSION a %V% ...
"%VPY%" bump_version.py %V%
if errorlevel 1 (
    pause
    exit /b 1
)

echo  [2/4] Generando metadatos y compilando EXE ...
"%VPY%" generar_version_info.py
if errorlevel 1 (
    echo  [ERROR] No se pudieron generar los metadatos del EXE.
    pause
    exit /b 1
)
"%VPY%" -m pip install pyinstaller --quiet
"%VPY%" -m PyInstaller --onefile --noconsole --clean --name "%NOMBRE_EXE%" --version-file "version_info.txt" --icon "favicon.ico" --add-data "favicon.ico;." --collect-all pypdfium2 generar_nomina.py
if errorlevel 1 (
    echo  [ERROR] Fallo la compilacion.
    pause
    exit /b 1
)

echo  [3/4] Versionando con git ...
git add .gitignore bump_version.py generar_nomina.py generar_version_info.py build_exe.bat publicar.bat README.md AGENTS.md CHANGELOG.md
git commit -m "Release v%V%"
if errorlevel 1 (
    echo  [ERROR] No se pudo crear el commit. Revisa los cambios pendientes.
    pause
    exit /b 1
)
git tag v%V%
if errorlevel 1 (
    echo  [ERROR] El tag v%V% ya existe. Elegi otra version; los tags publicados no se sobrescriben.
    pause
    exit /b 1
)
git push origin HEAD
if errorlevel 1 (
    echo  [ERROR] No se pudo hacer push del commit. Verifica tu conexion y tu sesion de GitHub.
    pause
    exit /b 1
)
git push origin v%V%
if errorlevel 1 (
    echo  [ERROR] No se pudo subir el tag v%V%.
    pause
    exit /b 1
)

echo  [4/4] Creando el Release en GitHub ...
set "GH="
where gh >nul 2>nul && set "GH=gh"
if not defined GH if exist "%ProgramFiles%\GitHub CLI\gh.exe" set "GH=%ProgramFiles%\GitHub CLI\gh.exe"
if not defined GH (
    echo  [AVISO] GitHub CLI no esta instalado. Instalalo desde https://cli.github.com/
    goto :manual
)
"%GH%" release create v%V% "dist\%NOMBRE_EXE%.exe" --title "Nomina de Docentes v%V%" --notes "Version %V% - generador de nomina docente desde el Cuadro de Personal"
if errorlevel 1 (
    echo  [AVISO] No se pudo crear el release con gh. Subilo manualmente.
    goto :manual
)
echo.
echo  [OK] Version v%V% publicada. Los usuarios con el EXE recibiran
echo       el aviso de actualizacion al ejecutarlo.
echo.
pause
exit /b 0

:manual
echo         1) Abre https://github.com/rodrigoangeloni/nomina_docente_de_cuadro_perdonal/releases
echo         2) Crea/edita el release del tag v%V%
echo         3) Adjunta el archivo dist\%NOMBRE_EXE%.exe
echo.
echo  [AVISO] El commit y el tag ya se subieron, pero el Release no se creo.
echo.
pause
exit /b 1
