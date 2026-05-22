@echo off
echo ================================================
echo  Actualizando Reporte de Personal...
echo ================================================
echo.

echo Cerrando Excel si esta abierto...
taskkill /f /im EXCEL.EXE 2>nul
timeout /t 2 /nobreak >nul

echo NOTA: Este bat genera desde la carpeta fuente (Inteligencia Artificial\Asistencia)
echo Ejecuta ese Actualizar.bat en su lugar para actualizar datos.
echo.
echo Si queres actualizar solo el HTML (sin regenerar datos), ejecuta este bat:
echo.

echo Actualizando GitHub Pages...
copy /Y "%~dp0Reporte_Personal.html" "%~dp0index.html" >nul
cd /d "%~dp0"
git add Reporte_Personal.html index.html
git commit -m "sync: actualizar asistencia %DATE% %TIME%"
git push origin master

echo.
echo Listo! El portal fue actualizado.
timeout /t 3 /nobreak >nul
