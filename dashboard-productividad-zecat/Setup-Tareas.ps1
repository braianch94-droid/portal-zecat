#Requires -RunAsAdministrator
<#
    Setup-Tareas.ps1
    Registra una tarea programada que ejecuta el ciclo completo:
      1. Genera el dashboard desde productividad.xlsx
      2. Copia el HTML al repositorio portal-zecat
      3. Hace git push → el portal online se actualiza en 2-3 min

    Ejecutar UNA SOLA VEZ como Administrador.
    Para cambiar el horario, editar $Horario y volver a ejecutar.
#>

# ── Configuración ─────────────────────────────────────────────────────────────
$Horario   = "07:00"   # Hora de ejecución diaria (formato HH:mm)
$NomTarea  = "DashboardProductividad_AutoUpdate"
$BatPath   = "C:\Users\bchevasco\OneDrive - Articulos Promocionales SA\Escritorio\Productividad\Actualizar-Dashboard.bat"
# ──────────────────────────────────────────────────────────────────────────────

if (-not (Test-Path $BatPath)) {
    Write-Error "No se encontro el archivo: $BatPath"
    Write-Host  "Asegurate de que Actualizar-Dashboard.bat este en la carpeta Productividad."
    exit 1
}

$action = New-ScheduledTaskAction `
    -Execute  "cmd.exe" `
    -Argument "/c `"$BatPath`""

$trigger = New-ScheduledTaskTrigger -Daily -At $Horario

$settings = New-ScheduledTaskSettingsSet `
    -ExecutionTimeLimit  (New-TimeSpan -Hours 1) `
    -StartWhenAvailable `
    -MultipleInstances   IgnoreNew `
    -Hidden

# Eliminar tarea anterior si existe
$existing = Get-ScheduledTask -TaskName $NomTarea -ErrorAction SilentlyContinue
if ($existing) {
    Unregister-ScheduledTask -TaskName $NomTarea -Confirm:$false
    Write-Host "Tarea anterior eliminada."
}

Register-ScheduledTask `
    -TaskName   $NomTarea `
    -Action     $action `
    -Trigger    $trigger `
    -Settings   $settings `
    -RunLevel   Highest `
    -Force | Out-Null

Write-Host ""
Write-Host "============================================="
Write-Host "  Tarea programada registrada correctamente"
Write-Host "============================================="
Write-Host ""
Write-Host "  Nombre  : $NomTarea"
Write-Host "  Horario : todos los dias a las $Horario hs"
Write-Host "  Accion  : genera dashboard + sube a GitHub"
Write-Host ""
Write-Host "  El portal online se actualizara automaticamente"
Write-Host "  cada dia ~$Horario hs (Argentina)."
Write-Host ""
Write-Host "  Podas verificar en:"
Write-Host "  Programador de tareas > Biblioteca del Programador de tareas"
Write-Host "  Tarea: $NomTarea"
Write-Host ""
