@echo off
cd /d "%~dp0"
echo Generando lista de conteo diario y enviando email...
echo.
powershell.exe -ExecutionPolicy Bypass -File "Generar_Conteo_Diario.ps1"
if %errorlevel% == 0 (
    echo.
    echo Listo! Email enviado y dashboard actualizado.
) else (
    echo.
    echo Ocurrio un error. Revisa la consola.
    pause
)
