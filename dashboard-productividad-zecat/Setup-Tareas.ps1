#Requires -RunAsAdministrator
<#
    Setup-Tareas.ps1
    Registra 3 tareas programadas para actualizar el Dashboard de Productividad:
      - 08:00 hs
      - 12:00 hs
      - 16:00 hs
    Ejecutar UNA SOLA VEZ como Administrador.
#>

$scriptPath = Join-Path $PSScriptRoot "Update-Dashboard.ps1"

if (-not (Test-Path $scriptPath)) {
    Write-Error "No se encontro el script: $scriptPath"
    exit 1
}

$exe  = "powershell.exe"
$args = "-ExecutionPolicy Bypass -WindowStyle Hidden -NonInteractive -File `"$scriptPath`""

$action   = New-ScheduledTaskAction -Execute $exe -Argument $args
$settings = New-ScheduledTaskSettingsSet `
    -ExecutionTimeLimit (New-TimeSpan -Hours 1) `
    -StartWhenAvailable `
    -MultipleInstances IgnoreNew

$tasks = @(
    @{ Name = "DashboardProductividad_08"; Hour = "08:00" }
    @{ Name = "DashboardProductividad_12"; Hour = "12:00" }
    @{ Name = "DashboardProductividad_16"; Hour = "16:00" }
)

foreach ($t in $tasks) {
    $trigger = New-ScheduledTaskTrigger -Daily -At $t.Hour

    $existing = Get-ScheduledTask -TaskName $t.Name -ErrorAction SilentlyContinue
    if ($existing) {
        Unregister-ScheduledTask -TaskName $t.Name -Confirm:$false
        Write-Host "Tarea existente eliminada: $($t.Name)"
    }

    Register-ScheduledTask `
        -TaskName  $t.Name `
        -Action    $action `
        -Trigger   $trigger `
        -Settings  $settings `
        -RunLevel  Highest `
        -Force | Out-Null

    Write-Host "OK  Tarea registrada: $($t.Name) a las $($t.Hour)"
}

Write-Host ""
Write-Host "Listo. El dashboard se actualizara automaticamente a las 08:00, 12:00 y 16:00."
Write-Host "Podes verificar las tareas en: Programador de tareas > Biblioteca del Programador de tareas"
