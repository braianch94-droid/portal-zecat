@echo off
chcp 65001 >nul
title Actualizando Dashboard de Productividad...

echo.
echo ================================================
echo   ZECAT — Dashboard de Productividad
echo ================================================
echo.

set PROD_FOLDER=%~dp0
set PROD_FOLDER=%PROD_FOLDER:~0,-1%
set REPO_FOLDER=C:\Users\bchevasco\OneDrive - Articulos Promocionales SA\Escritorio\GIT huB
set GH_BIN=%USERPROFILE%\gh_cli\bin

:: ── PASO 1: Generar el HTML desde Excel ──────────────────
echo [1/3] Generando dashboard desde productividad.xlsx...
powershell -ExecutionPolicy Bypass -WindowStyle Normal -File "%PROD_FOLDER%\Update-Dashboard.ps1"
if errorlevel 1 (
    echo.
    echo  ERROR: No se pudo generar el dashboard.
    echo  Verificar que productividad.xlsx este cerrado y actualizado.
    pause
    exit /b 1
)
echo  OK - Dashboard generado.
echo.

:: ── PASO 2: Copiar HTML al repositorio ───────────────────
echo [2/3] Copiando al portal...
copy /Y "%PROD_FOLDER%\Dashboard_Productividad.html" "%REPO_FOLDER%\dashboard-productividad-zecat\index.html" >nul
if errorlevel 1 (
    echo.
    echo  ERROR: No se pudo copiar el archivo al repositorio.
    pause
    exit /b 1
)
echo  OK - Archivo copiado al repo.
echo.

:: ── PASO 3: Subir a GitHub ────────────────────────────────
echo [3/3] Subiendo a GitHub...
cd /d "%REPO_FOLDER%"

:: Actualizar token de autenticación desde gh CLI
set PATH=%PATH%;%GH_BIN%
for /f "delims=" %%T in ('gh auth token 2^>nul') do set GH_TOKEN=%%T
if "%GH_TOKEN%"=="" (
    echo  AVISO: No se encontro token de GitHub. Intentando push igual...
) else (
    git remote set-url origin "https://braianch94-droid:%GH_TOKEN%@github.com/braianch94-droid/portal-zecat.git" >nul 2>&1
)

git add dashboard-productividad-zecat/index.html >nul 2>&1
git diff --cached --quiet
if not errorlevel 1 (
    echo  Sin cambios nuevos para subir.
    goto :fin
)

git commit -m "update: productividad %DATE%" >nul 2>&1
git push >nul 2>&1
if errorlevel 1 (
    echo.
    echo  ERROR: No se pudo hacer push a GitHub.
    echo  Verificar conexion a internet y autenticacion de gh.
    pause
    exit /b 1
)
echo  OK - Subido a GitHub.
echo.

:fin
echo ================================================
echo   Listo. El portal se actualiza en 2-3 min.
echo   https://braianch94-droid.github.io/portal-zecat/
echo ================================================
echo.
echo Actualizado: %date% %time% >> "%PROD_FOLDER%\actualizaciones.log"
timeout /t 5 >nul
