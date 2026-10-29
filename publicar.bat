@echo off
setlocal EnableExtensions
cd /d "%~dp0"
title Nomina de Docentes - Publicar version

rem ============================================================
rem  Publica una nueva version:
rem   1) actualiza VERSION en generar_nomina.py (SemVer)
rem   2) genera los metadatos de Windows y compila el EXE
rem   3) crea el tag vX.Y.Z y lo sube
rem   4) crea el Release en GitHub con el EXE
rem  Uso: publicar.bat 1.1.0
rem ============================================================

if "%~1"=="" (
    echo  Uso: publicar.bat VERSION    ej: publicar.bat 1.1.0
    pause
    exit /b 1
)
set "V=%~1"
set "NOMBRE_EXE=NominaDocentes"

if not exist ".venv\Scripts\python.exe" (
    echo  [ERROR] Falta el entorno .venv. Ejecuta generar.bat una vez primero.
    pause
    exit /b 1
)
set "VPY=.venv\Scripts\python.exe"

rem  No reutilizar una version ya publicada
git rev-parse --verify --quiet v%V% >nul 2>nul
if not errorlevel 1 (
    echo  [ERROR] El tag v%V% ya existe. Elegi otra version; los tags publicados no se sobrescriben.
    pause
    exit /b 1
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
git commit -m "Release v%V%: GUI guiada, metadatos del EXE y cabecera desde PDF"
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
where gh >nul 2>nul
if %errorlevel%==0 (
    gh release create v%V% "dist\%NOMBRE_EXE%.exe" --title "Nomina de Docentes v%V%" --notes "Version %V% - generador de nomina docente desde el Cuadro de Personal"
    if errorlevel 1 goto :release_fallo
    goto :fin_ok
) else (
    goto :sin_gh
)

:release_fallo
echo  [AVISO] No se pudo crear el release con gh. Subilo manualmente.
goto :manual

:sin_gh
echo  [INFO] GitHub CLI no esta instalado. Instalalo desde https://cli.github.com/
echo         o sube el EXE manualmente:
:manual
echo         1) Abre https://github.com/rodrigoangeloni/nomina_docente_de_cuadro_perdonal/releases
echo         2) Crea/edita el release del tag v%V%
echo         3) Adjunta el archivo dist\%NOMBRE_EXE%.exe
echo.
echo  [AVISO] El commit y el tag ya se subieron, pero el Release no se creo.
echo.
pause
exit /b 0

:fin_ok
echo.
echo  [OK] Version v%V% publicada. Los usuarios con el EXE recibiran
echo       el aviso de actualizacion al ejecutarlo.
echo.
pause
