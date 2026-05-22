@echo off
setlocal
set REPO=%~dp0
set REPO=%REPO:~0,-1%

echo ==============================================
echo  ZECAT - Actualizar y publicar dashboard
echo ==============================================
echo.

echo [1/3] Generando dashboard...
powershell -ExecutionPolicy Bypass -File "%REPO%\Update-Dashboard.ps1"
if %errorlevel% neq 0 (
    echo ERROR: El script fallo. Revisa los datos fuente.
    pause
    exit /b 1
)
echo    OK
echo.

echo [2/3] Preparando para publicar...
cd /d "%REPO%"
copy /Y "%REPO%\Dashboard_Productividad.html" "%REPO%\index.html" >nul
git add index.html
git diff --cached --quiet
if %errorlevel% equ 0 (
    echo    Sin cambios nuevos. El dashboard ya estaba actualizado.
    goto :done
)
for /f "tokens=2 delims==" %%I in ('wmic os get localdatetime /value') do set dt=%%I
set FECHA=%dt:~6,2%/%dt:~4,2%/%dt:~0,4% %dt:~8,2%:%dt:~10,2%
git commit -m "Dashboard actualizado %FECHA%"
echo    OK
echo.

echo [3/3] Publicando en GitHub...
git push
if %errorlevel% neq 0 (
    echo ERROR: No se pudo hacer push. Revisa tu conexion o el token.
    pause
    exit /b 1
)
echo    OK
echo.

:done
echo ==============================================
echo  Listo. Ver dashboard online en:
echo  https://braianch94-droid.github.io/dashboard-productividad-zecat/
echo ==============================================
echo.
echo Actualizado: %date% %time% >> "%REPO%\actualizaciones.log"
pause
