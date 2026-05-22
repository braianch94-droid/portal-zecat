@echo off
:: Ejecutar con "clic derecho > Ejecutar como administrador" UNA SOLA VEZ

set "BASE=C:\Users\bchevasco\OneDrive - Articulos Promocionales SA\Escritorio\Inteligencia Artificial\Ciclico"

echo Eliminando tareas anteriores si existen...
schtasks /Delete /TN "Zecat - Actualizar Dashboard Ciclico" /F 2>nul
schtasks /Delete /TN "Zecat - Conteo Ciclico Diario y Email" /F 2>nul

echo.
echo Instalando tarea: Conteo Diario y Email (7:00 AM, Lunes a Viernes)...
powershell.exe -ExecutionPolicy Bypass -Command ^
  "$BASE = 'C:\Users\bchevasco\OneDrive - Articulos Promocionales SA\Escritorio\Inteligencia Artificial\Ciclico';" ^
  "$action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument ('-ExecutionPolicy Bypass -WindowStyle Hidden -File \"' + $BASE + '\Generar_Conteo_Diario.ps1\"') -WorkingDirectory $BASE;" ^
  "$trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday,Tuesday,Wednesday,Thursday,Friday -At '07:00AM';" ^
  "$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -RunOnlyIfNetworkAvailable -ExecutionTimeLimit (New-TimeSpan -Minutes 30) -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 5) -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries;" ^
  "Register-ScheduledTask -TaskName 'Zecat - Conteo Ciclico Diario y Email' -Action $action -Trigger $trigger -Settings $settings -RunLevel Highest -Force;"

echo.
echo Instalando tarea: Actualizar Dashboard (8:00 AM, Lunes a Viernes)...
powershell.exe -ExecutionPolicy Bypass -Command ^
  "$BASE = 'C:\Users\bchevasco\OneDrive - Articulos Promocionales SA\Escritorio\Inteligencia Artificial\Ciclico';" ^
  "$action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument ('-ExecutionPolicy Bypass -WindowStyle Hidden -File \"' + $BASE + '\Actualizar_Dashboard.ps1\"') -WorkingDirectory $BASE;" ^
  "$trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday,Tuesday,Wednesday,Thursday,Friday -At '08:00AM';" ^
  "$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -RunOnlyIfNetworkAvailable -ExecutionTimeLimit (New-TimeSpan -Minutes 30) -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 5) -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries;" ^
  "Register-ScheduledTask -TaskName 'Zecat - Actualizar Dashboard Ciclico' -Action $action -Trigger $trigger -Settings $settings -RunLevel Highest -Force;"

echo.
echo ============================================================
echo Tareas instaladas:
echo  - 07:00 AM Lun-Vie: Analiza stock, elige 20 items, envia email
echo  - 08:00 AM Lun-Vie: Actualiza dashboard HTML
echo  - BATERIA: No detiene la tarea si cambia a bateria
echo  - INICIO TARDIO: Corre igual si la PC estaba apagada a las 7AM
echo ============================================================
echo.
pause
