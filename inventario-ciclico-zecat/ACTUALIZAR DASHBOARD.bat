@echo off
cd /d "%~dp0"
powershell.exe -ExecutionPolicy Bypass -File "Actualizar_Dashboard.ps1"
if %errorlevel% == 0 (
    echo.
    echo Dashboard actualizado correctamente.
    start "" "Dashboard_Ciclico.html"
) else (
    echo.
    echo Ocurrio un error al actualizar el dashboard.
    pause
)
