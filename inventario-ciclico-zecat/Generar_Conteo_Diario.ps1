# ============================================================
# Generar_Conteo_Diario.ps1
# Analiza Stock a HOY.xlsx, elige 10 articulos a contar,
# compara con lo contado ayer en Ciclico y envia email.
# ============================================================

$ScriptDir     = Split-Path -Parent $MyInvocation.MyCommand.Path
$StockFile     = Join-Path $ScriptDir "Stock a HOY.xlsx"
$CiclicoFile   = Join-Path $ScriptDir "Ciclico ARG 2025-26.xlsx"
$HistorialFile = Join-Path $ScriptDir "historial_recomendaciones.json"
$CoberturaFile = Join-Path $ScriptDir "cobertura_familias.json"
$CicloFile     = Join-Path $ScriptDir "ciclo_actual.json"

$DashUrl = ("file:///" + ($ScriptDir -replace "\\","/") + "/Dashboard_Ciclico.html") -replace " ","%20"

$Destinatarios = "bchevasco@zecat.com; deposito@zecat.com; logistica2@zecat.com"
$HoyFecha      = Get-Date -Format "dd/MM/yyyy"
$AyerFecha     = (Get-Date).AddDays(-1).ToString("dd/MM/yyyy")
$HoyLabel      = (Get-Culture).TextInfo.ToTitleCase((Get-Date -Format "dddd dd 'de' MMMM yyyy").ToLower())

function HE($s) { if ($null -eq $s) { return "" }; [System.Net.WebUtility]::HtmlEncode([string]$s) }

Write-Host "=== Conteo Diario $HoyFecha ==="

# ── FINES DE SEMANA Y FERIADOS NACIONALES ────────────────────────
$hoy        = [datetime]::Today
$diaSemana  = [int]$hoy.DayOfWeek   # 0=Dom 1=Lun ... 6=Sab
if ($diaSemana -eq 0 -or $diaSemana -eq 6) {
    Write-Host "Fin de semana ($HoyFecha). No se ejecuta."
    exit 0
}
$feriadosARG = @(
    # 2025 ──────────────────────────────────────────────────────
    "01/01/2025","03/03/2025","04/03/2025","24/03/2025","02/04/2025",
    "18/04/2025","01/05/2025","25/05/2025","16/06/2025","20/06/2025",
    "09/07/2025","18/08/2025","13/10/2025","24/11/2025","08/12/2025","25/12/2025",
    # 2026 ──────────────────────────────────────────────────────
    "01/01/2026","16/02/2026","17/02/2026","24/03/2026","02/04/2026",
    "03/04/2026","01/05/2026","25/05/2026","15/06/2026","20/06/2026",
    "09/07/2026","17/08/2026","12/10/2026","23/11/2026","08/12/2026","25/12/2026"
    # Agregar puentes turisticos cuando el gobierno los anuncie
)
if ($feriadosARG -contains $HoyFecha) {
    Write-Host "Feriado nacional: $HoyFecha. No se ejecuta."
    exit 0
}
# ─────────────────────────────────────────────────────────────────

# ── PAUSA VACACIONES ─────────────────────────────────────────────
$pausaDesde = [datetime]"2026-05-18"
$pausaHasta = [datetime]"2026-05-25"
if ($hoy -ge $pausaDesde -and $hoy -le $pausaHasta) {
    Write-Host "PAUSA VACACIONES: $HoyFecha esta dentro del periodo $($pausaDesde.ToString('dd/MM/yyyy')) - $($pausaHasta.ToString('dd/MM/yyyy')). No se envia conteo."
    exit 0
}
# ─────────────────────────────────────────────────────────────────

# ── 1. Cargar historial (para seguimiento de ayer) ───────────────
$historial = @()
if (Test-Path $HistorialFile) {
    try {
        $raw = Get-Content $HistorialFile -Raw -Encoding UTF8 | ConvertFrom-Json
        $historial = if ($raw -is [array]) { $raw } else { @($raw) }
    } catch { $historial = @() }
}

# ── 1b. Cargar ciclo actual ───────────────────────────────────
# yaRecomendados = todos los articulos recomendados en el ciclo en curso
# Un articulo no vuelve a aparecer hasta que se completen TODOS los elegibles
$cicloNum      = 1
$cicloFecha    = $HoyFecha
$yaRecomendados = @{}   # hashtable articulo.Trim() -> true (lookup O(1))

if (Test-Path $CicloFile) {
    try {
        $cicloRaw = Get-Content $CicloFile -Raw -Encoding UTF8 | ConvertFrom-Json
        $cicloNum   = [int]$cicloRaw.cicloNumero
        $cicloFecha = $cicloRaw.fechaInicio
        foreach ($a in $cicloRaw.yaRecomendados) { $yaRecomendados[$a.Trim()] = $true }
        Write-Host "Ciclo $cicloNum iniciado $cicloFecha | Ya recomendados: $($yaRecomendados.Count)"
    } catch { Write-Host "Ciclo nuevo (archivo no leible)" }
} else {
    Write-Host "Primer ciclo - creando ciclo_actual.json"
}

# ── 2. Leer Ciclico ───────────────────────────────────────────
Write-Host "Leyendo Ciclico..."
$tmpC = "$env:TEMP\cic_tmp.xlsx"
Copy-Item $CiclicoFile $tmpC -Force

$contadosAyer     = @{}   # art -> bool
$contadosReciente = @{}   # art -> datetime ultimo conteo
$ayerTotal = 0; $ayerSinDif = 0; $ayerPos = 0; $ayerNeg = 0
$ayerFamilias = @{}

$xl = New-Object -ComObject Excel.Application; $xl.Visible=$false; $xl.DisplayAlerts=$false
try {
    $wb = $xl.Workbooks.Open($tmpC); $ws = $wb.Sheets.Item(1); $nr = $ws.UsedRange.Rows.Count
    for ($r = 2; $r -le $nr; $r++) {
        $art = $ws.Cells.Item($r, 2).Text
        $dia = $ws.Cells.Item($r,14).Text
        $dif = $ws.Cells.Item($r,10).Text
        $fam = $ws.Cells.Item($r, 3).Text
        if ($art -eq "" -or $dia -eq "") { continue }
        if ($dia -eq $AyerFecha) {
            $contadosAyer[$art] = $true
            $ayerTotal++
            if ($dif -eq "Sin Dif")       { $ayerSinDif++ }
            elseif ($dif -eq "Dif Positiva") { $ayerPos++ }
            elseif ($dif -eq "Dif Negativa") { $ayerNeg++ }
            if ($fam -ne "") {
                if ($ayerFamilias.ContainsKey($fam)) { $ayerFamilias[$fam]++ } else { $ayerFamilias[$fam] = 1 }
            }
        }
        try {
            $dp  = [datetime]::ParseExact($dia,"d/MM/yyyy",$null)
            $key = $art.Trim()
            if (-not $contadosReciente.ContainsKey($key) -or $contadosReciente[$key] -lt $dp) {
                $contadosReciente[$key] = $dp
            }
        } catch {}
    }
    $wb.Close($false)
} finally { $xl.Quit(); [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl)|Out-Null }
Remove-Item $tmpC -Force

$ayerAcc = if ($ayerTotal -gt 0) { [math]::Round($ayerSinDif/$ayerTotal*100,1) } else { 0 }
Write-Host "Ayer: $ayerTotal conteos (exactitud $ayerAcc%)"

# ── 3. Leer Stock a HOY ───────────────────────────────────────
Write-Host "Leyendo Stock a HOY..."
$tmpS = "$env:TEMP\stk_tmp.xlsx"
Copy-Item $StockFile $tmpS -Force

$stockItems = @()
$xl2 = New-Object -ComObject Excel.Application; $xl2.Visible=$false; $xl2.DisplayAlerts=$false
try {
    $wb2 = $xl2.Workbooks.Open($tmpS); $ws2 = $wb2.Sheets.Item(1); $nr2 = $ws2.UsedRange.Rows.Count
    for ($r = 2; $r -le $nr2; $r++) {
        $situ = $ws2.Cells.Item($r,48).Text.ToLower().Trim()
        if ($situ -notlike "*activo en web*" -or $situ -like "*inactivo*") { continue }
        $art   = $ws2.Cells.Item($r, 1).Text.Trim()
        $sw    = $ws2.Cells.Item($r, 7).Value2
        $ddpU  = $ws2.Cells.Item($r,13).Value2
        $ddpV  = $ws2.Cells.Item($r,14).Value2
        $sku   = $ws2.Cells.Item($r,39).Text.Trim()
        $fam   = $ws2.Cells.Item($r,42).Text.Trim()
        $sOrig = $ws2.Cells.Item($r,48).Text.Trim()
        if ($art -eq "" -or $sw -eq $null) { continue }
        $swN   = if ($sw -le 0) { 0.1 } else { $sw }
        $dvN   = if ($null -eq $ddpV) { 0 } else { $ddpV }
        $boost = if ($swN -lt 200) { 3 } else { 1 }
        $score = ($dvN / $swN) * $boost
        $stockItems += [PSCustomObject]@{
            Articulo  = $art; SKU=$sku; Familia=$fam
            StockWeb  = [math]::Round($swN); DdpUnit=[math]::Round($ddpU,2); DdpVal=[math]::Round($dvN,2)
            Situacion = $sOrig; Critico=($swN -lt 200); Score=$score
        }
    }
    $wb2.Close($false)
} finally { $xl2.Quit(); [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl2)|Out-Null }
Remove-Item $tmpS -Force
Write-Host "Elegibles: $($stockItems.Count)"

# ── 4. Exportar cobertura por familia ─────────────────────────
$famTotal   = @{}
$famContados= @{}
foreach ($it in $stockItems) {
    $f = $it.Familia; if ($f -eq "") { $f = "Sin Familia" }
    if ($famTotal.ContainsKey($f)) { $famTotal[$f]++ } else { $famTotal[$f] = 1 }
    $artKey = $it.Articulo.Trim().ToLower()
    $contado = $contadosReciente.ContainsKey($it.Articulo) -or
               ($contadosReciente.Keys | Where-Object { $_.Trim().ToLower() -eq $artKey }).Count -gt 0
    if ($contado) {
        if ($famContados.ContainsKey($f)) { $famContados[$f]++ } else { $famContados[$f] = 1 }
    }
}
$totalElegibles = $stockItems.Count
$totalContados  = ($famContados.Values | Measure-Object -Sum).Sum
$coberturaGlobal= if ($totalElegibles -gt 0) { [math]::Round($totalContados/$totalElegibles*100,1) } else { 0 }

$famCovArr = @()
foreach ($f in $famTotal.Keys | Sort-Object) {
    $tot = $famTotal[$f]
    $cnt = if ($famContados.ContainsKey($f)) { $famContados[$f] } else { 0 }
    $pend= $tot - $cnt
    $pct = if ($tot -gt 0) { [math]::Round($cnt/$tot*100,1) } else { 0 }
    $famCovArr += [PSCustomObject]@{ familia=$f; total=$tot; contados=$cnt; pendientes=$pend; pct=$pct }
}
$cobJson = [PSCustomObject]@{
    fechaGenerado=$HoyFecha; totalElegibles=$totalElegibles
    totalContados=[int]$totalContados; coberturaGlobal=$coberturaGlobal
    familias=$famCovArr
}
$utf8NoBom = New-Object System.Text.UTF8Encoding $false
[System.IO.File]::WriteAllText($CoberturaFile, ($cobJson|ConvertTo-Json -Depth 5), $utf8NoBom)
Write-Host "Cobertura exportada: $coberturaGlobal% ($totalContados/$totalElegibles)"

# ── 5. Seleccionar items del ciclo ───────────────────────────
# Regla: cada item se recomienda UNA sola vez por ciclo.
# Cuando se agotan todos los elegibles, se reinicia el ciclo automaticamente.
$totalTarget = 20

# Filtrar los que ya NO necesitan recomendarse:
#   - ya fueron contados en el Ciclico ($contadosReciente)
#   - ya fueron recomendados en este ciclo ($yaRecomendados)
$itemsPendientes = @($stockItems | Where-Object {
    $a = $_.Articulo.Trim()
    (-not $contadosReciente.ContainsKey($a)) -and (-not $yaRecomendados.ContainsKey($a))
})

$yaContadosTotal = ($stockItems | Where-Object { $contadosReciente.ContainsKey($_.Articulo.Trim()) }).Count

# Si el ciclo se completo (no quedan pendientes), reiniciar
$cicloCompleto = $false
if ($itemsPendientes.Count -eq 0) {
    $cicloCompleto = $true
    $cicloNum++
    $cicloFecha     = $HoyFecha
    $yaRecomendados = @{}
    # Al reiniciar el ciclo, el nuevo pool excluye solo los contados (yaRecomendados quedo vacio)
    $itemsPendientes = @($stockItems | Where-Object { -not $contadosReciente.ContainsKey($_.Articulo.Trim()) })
    Write-Host "*** CICLO COMPLETADO - Iniciando Ciclo $cicloNum ***"
}

$pendientesTotal = $itemsPendientes.Count
Write-Host "Ciclo $cicloNum | Contados en Ciclico: $yaContadosTotal | Ya recomendados: $($yaRecomendados.Count) | Pendientes hoy: $pendientesTotal de $($stockItems.Count)"

# Seleccion en cascada por tramo de stock (prioridad: menor stock primero)
# Dentro del tramo, ordenar por Score (stock critico + mayor valor DDP)
function Get-TierCandidatos($items, $minSW, $maxSW, $n) {
    if ($n -le 0) { return @() }
    $pool = @($items | Where-Object { $_.StockWeb -ge $minSW -and $_.StockWeb -lt $maxSW } | Sort-Object Score -Descending)
    return @($pool | Select-Object -First $n)
}

$selT1 = Get-TierCandidatos $itemsPendientes 0    300     $totalTarget
$rem2  = $totalTarget - $selT1.Count
$selT2 = Get-TierCandidatos $itemsPendientes 300  500     $rem2
$rem3  = $rem2 - $selT2.Count
$selT3 = Get-TierCandidatos $itemsPendientes 500  1000    $rem3
$rem4  = $rem3 - $selT3.Count
$selT4 = Get-TierCandidatos $itemsPendientes 1000 9999999 $rem4

# Lista unificada ordenada por stock ascendente (mas critico primero)
$todosItems = @()
if ($selT1.Count -gt 0) { $todosItems += @($selT1) }
if ($selT2.Count -gt 0) { $todosItems += @($selT2) }
if ($selT3.Count -gt 0) { $todosItems += @($selT3) }
if ($selT4.Count -gt 0) { $todosItems += @($selT4) }
$todosItems = @($todosItems | Sort-Object StockWeb)

$cntT1 = $selT1.Count; $cntT2 = $selT2.Count
$cntT3 = $selT3.Count; $cntT4 = $selT4.Count
$totalItems = $todosItems.Count
Write-Host "Seleccionados: T1=$cntT1 | T2=$cntT2 | T3=$cntT3 | T4=$cntT4 | Total=$totalItems"

# ── 6. Analizar recomendados ayer vs contados ─────────────────
$ayerEntry = $historial | Where-Object { $_.fecha -eq $AyerFecha } | Select-Object -Last 1
$recAyer = @(); $pendientesAyer = @()
if ($null -ne $ayerEntry -and $ayerEntry.items.Count -gt 0) {
    foreach ($i in $ayerEntry.items) {
        if ($contadosAyer.ContainsKey($i.articulo)) { $recAyer += $i.articulo }
        else { $pendientesAyer += $i.articulo }
    }
}

# ── 7. Guardar historial y actualizar ciclo ───────────────────
$nuevaEntrada = [PSCustomObject]@{
    fecha = $HoyFecha
    items = @($todosItems | ForEach-Object { [PSCustomObject]@{ articulo=$_.Articulo.Trim(); sku=$_.SKU; stockWeb=$_.StockWeb; ddpVal=$_.DdpVal } })
}
$historial = @($historial | Where-Object { $_.fecha -ne $HoyFecha }) + $nuevaEntrada
[System.IO.File]::WriteAllText($HistorialFile, ($historial|ConvertTo-Json -Depth 5), $utf8NoBom)

# Agregar items de hoy al ciclo y guardar
foreach ($it in $todosItems) { $yaRecomendados[$it.Articulo.Trim()] = $true }
$pendientesPostSelect = ($stockItems | Where-Object {
    $a = $_.Articulo.Trim()
    (-not $contadosReciente.ContainsKey($a)) -and (-not $yaRecomendados.ContainsKey($a))
}).Count
$cicloObj = [PSCustomObject]@{
    cicloNumero    = $cicloNum
    fechaInicio    = $cicloFecha
    totalElegibles = $stockItems.Count
    contadosEnCiclico = $yaContadosTotal
    recomendadosEsteCiclo = $yaRecomendados.Count
    pendientes     = $pendientesPostSelect
    yaRecomendados = @($yaRecomendados.Keys | Sort-Object)
}
[System.IO.File]::WriteAllText($CicloFile, ($cicloObj | ConvertTo-Json -Depth 3), $utf8NoBom)
Write-Host "Ciclo guardado | Contados: $yaContadosTotal | Recomendados ciclo: $($yaRecomendados.Count) | Pendientes: $pendientesPostSelect"

# ── 8. Construir tabla HOY unificada (lista unica ordenada por stock) ─
$tablaHoy = ""
$pos = 1
foreach ($it in $todosItems) {
    $aNom  = HE $it.Articulo
    $fNom  = HE $it.Familia
    $sw    = [string]$it.StockWeb
    $ddpU  = '$' + $it.DdpUnit.ToString('F2')
    $ddpV  = '$' + $it.DdpVal.ToString('N0')
    # Color segun urgencia (tramo de stock)
    $swN = [double]$it.StockWeb
    if     ($swN -lt 300)  { $swCol = "#dc2626"; $rowBg = "#fff8f8" }
    elseif ($swN -lt 500)  { $swCol = "#d97706"; $rowBg = "#fffdf7" }
    elseif ($swN -lt 1000) { $swCol = "#a16207"; $rowBg = "#fefef5" }
    else                   { $swCol = "#0369a1"; $rowBg = "#f8fbff" }
    $sitB = if ($it.Situacion -like "stock < 500*") {
        "<span style='background:#fef3c7;color:#92400e;border-radius:10px;padding:2px 8px;font-size:10px'>stock&lt;500</span>"
    } else {
        "<span style='background:#dcfce7;color:#166534;border-radius:10px;padding:2px 8px;font-size:10px'>con stock</span>"
    }
    $tablaHoy += "<tr style='border-bottom:1px solid #f3f4f6;background:$rowBg'>"
    $tablaHoy += "<td style='padding:8px 12px;color:#9ca3af;font-weight:700'>$pos</td>"
    $tablaHoy += "<td style='padding:8px 12px'>$aNom</td>"
    $tablaHoy += "<td style='padding:8px 12px;color:#6b7280;font-size:12px'>$fNom</td>"
    $tablaHoy += "<td style='padding:8px 12px;text-align:right;font-weight:700;color:$swCol'>$sw</td>"
    $tablaHoy += "<td style='padding:8px 12px;text-align:right;color:#374151'>$ddpU</td>"
    $tablaHoy += "<td style='padding:8px 12px;text-align:right;font-weight:700;color:#d97706'>$ddpV</td>"
    $tablaHoy += "<td style='padding:8px 12px'>$sitB</td>"
    $tablaHoy += "</tr>"
    $pos++
}

# ── 9. Construir resumen ayer ─────────────────────────────────
$seccionAyer = ""

# Bloque 1: estadisticas del dia
if ($ayerTotal -gt 0) {
    $accColor = if ($ayerAcc -ge 60) { "#16a34a" } elseif ($ayerAcc -ge 40) { "#d97706" } else { "#dc2626" }
    $famCubiertas = ($ayerFamilias.Keys | Sort-Object) -join ", "
    $seccionAyer += @"
<h2 style='font-size:17px;color:#1a1a2e;margin:0 0 14px'>&#128202; Resumen del conteo &mdash; $AyerFecha</h2>
<div style='display:flex;gap:12px;margin-bottom:16px;flex-wrap:wrap'>
  <div style='background:#f8fafc;border-radius:10px;padding:12px 18px;border-left:3px solid #2563eb;min-width:120px'>
    <div style='font-size:10px;text-transform:uppercase;color:#6b7280;font-weight:600'>Conteos realizados</div>
    <div style='font-size:24px;font-weight:700;color:#1a1a2e'>$ayerTotal</div>
  </div>
  <div style='background:#f8fafc;border-radius:10px;padding:12px 18px;border-left:3px solid $accColor;min-width:120px'>
    <div style='font-size:10px;text-transform:uppercase;color:#6b7280;font-weight:600'>Exactitud</div>
    <div style='font-size:24px;font-weight:700;color:$accColor'>$ayerAcc%</div>
  </div>
  <div style='background:#f8fafc;border-radius:10px;padding:12px 18px;border-left:3px solid #16a34a;min-width:90px'>
    <div style='font-size:10px;text-transform:uppercase;color:#6b7280;font-weight:600'>Sin Dif.</div>
    <div style='font-size:24px;font-weight:700;color:#16a34a'>$ayerSinDif</div>
  </div>
  <div style='background:#f8fafc;border-radius:10px;padding:12px 18px;border-left:3px solid #f59e0b;min-width:90px'>
    <div style='font-size:10px;text-transform:uppercase;color:#6b7280;font-weight:600'>Dif. Positiva</div>
    <div style='font-size:24px;font-weight:700;color:#d97706'>$ayerPos</div>
  </div>
  <div style='background:#f8fafc;border-radius:10px;padding:12px 18px;border-left:3px solid #dc2626;min-width:90px'>
    <div style='font-size:10px;text-transform:uppercase;color:#6b7280;font-weight:600'>Dif. Negativa</div>
    <div style='font-size:24px;font-weight:700;color:#dc2626'>$ayerNeg</div>
  </div>
</div>
<p style='font-size:12px;color:#6b7280;margin-bottom:20px'>Familias cubiertas ayer: <strong>$famCubiertas</strong></p>
"@
} else {
    $seccionAyer += "<h2 style='font-size:17px;color:#1a1a2e;margin:0 0 12px'>&#128202; Resumen del conteo &mdash; $AyerFecha</h2>"
    $seccionAyer += "<p style='font-size:13px;color:#9ca3af;margin-bottom:16px'>No se registraron conteos en el C&iacute;clico para el $AyerFecha.</p>"
}

# Bloque 2: seguimiento de recomendados
if ($null -ne $ayerEntry -and $ayerEntry.items.Count -gt 0) {
    $cntOK  = $recAyer.Count
    $cntPen = $pendientesAyer.Count
    $seccionAyer += "<h3 style='font-size:14px;color:#1a1a2e;margin:16px 0 10px'>Seguimiento de los art&iacute;culos recomendados ayer</h3>"
    $seccionAyer += "<p style='font-size:12px;color:#6b7280;margin-bottom:12px'><strong style='color:#16a34a'>$cntOK contados</strong> &nbsp;|&nbsp; <strong style='color:#dc2626'>$cntPen pendientes</strong></p>"
    $seccionAyer += "<table style='width:100%;border-collapse:collapse;font-size:13px'>"
    $seccionAyer += "<thead><tr style='background:#f8fafc'><th style='padding:8px 12px;text-align:left;font-size:11px;text-transform:uppercase;color:#6b7280;border-bottom:2px solid #e5e7eb;width:40px'></th><th style='padding:8px 12px;text-align:left;font-size:11px;text-transform:uppercase;color:#6b7280;border-bottom:2px solid #e5e7eb'>Articulo solicitado</th><th style='padding:8px 12px;text-align:left;font-size:11px;text-transform:uppercase;color:#6b7280;border-bottom:2px solid #e5e7eb'>Estado</th></tr></thead><tbody>"
    foreach ($i in $ayerEntry.items) {
        $wasC = $contadosAyer.ContainsKey($i.articulo)
        $icon  = if ($wasC) { "&#9989;" } else { "&#10060;" }
        $col   = if ($wasC) { "#16a34a" } else { "#dc2626" }
        $est   = if ($wasC) { "Contado" } else { "Pendiente" }
        $aE    = HE $i.articulo
        $seccionAyer += "<tr style='border-bottom:1px solid #f3f4f6'><td style='padding:8px 12px;font-size:18px'>$icon</td><td style='padding:8px 12px'>$aE</td><td style='padding:8px 12px;color:$col;font-weight:600'>$est</td></tr>"
    }
    $seccionAyer += "</tbody></table>"

    # Alerta si hay pendientes
    if ($cntPen -gt 0) {
        $pendList = ($pendientesAyer | ForEach-Object { "<li style='margin:4px 0'>$(HE $_)</li>" }) -join ""
        $seccionAyer += "<div style='background:#fff7ed;border:1px solid #fed7aa;border-radius:8px;padding:12px 16px;margin-top:12px'>"
        $seccionAyer += "<strong style='color:#9a3412;font-size:13px'>&#9888; Articulos que quedaron pendientes de ayer:</strong>"
        $seccionAyer += "<ul style='margin:8px 0 0 16px;color:#92400e;font-size:12px'>$pendList</ul></div>"
    }
} else {
    $seccionAyer += "<p style='font-size:12px;color:#9ca3af;margin-top:8px'>No hay recomendaciones registradas para $AyerFecha.</p>"
}

# ── 10. Armar HTML del email ──────────────────────────────────
$htmlEmail = "<!DOCTYPE html><html><head><meta charset='UTF-8'></head><body style='font-family:Arial,sans-serif;background:#f0f2f5;margin:0;padding:20px'>"
$htmlEmail += "<div style='max-width:720px;margin:0 auto;background:white;border-radius:12px;overflow:hidden;box-shadow:0 4px 20px rgba(0,0,0,.1)'>"

# Header
$htmlEmail += "<div style='background:linear-gradient(135deg,#1a1a2e,#0f3460);color:white;padding:24px 32px'>"
$htmlEmail += "<div style='font-size:12px;opacity:.7;margin-bottom:4px'>ZECAT &mdash; Art&iacute;culos Promocionales SA</div>"
$htmlEmail += "<h1 style='font-size:20px;font-weight:700;margin:0'>Conteo C&iacute;clico Diario</h1>"
$htmlEmail += "<div style='font-size:13px;opacity:.85;margin-top:6px'>$HoyLabel</div></div>"

# Body
$htmlEmail += "<div style='padding:28px 32px'>"

# ── BLOQUE 1: Resumen del dia anterior ──────────────────────────
$htmlEmail += $seccionAyer

# Separador
$htmlEmail += "<hr style='border:none;border-top:2px solid #e5e7eb;margin:28px 0'>"

# ── BLOQUE 2: Lista para contar hoy ─────────────────────────────
$cicloRecorridos  = $yaContadosTotal + $yaRecomendados.Count
$cicloRestantes   = $pendientesPostSelect
$cicloProgresoPct = if ($stockItems.Count -gt 0) { [math]::Round($cicloRecorridos / $stockItems.Count * 100, 1) } else { 0 }
$htmlEmail += "<h2 style='font-size:17px;color:#1a1a2e;margin:0 0 8px'>&#128203; Lista para contar hoy &mdash; $HoyLabel</h2>"
$htmlEmail += "<div style='background:#f0f9ff;border:1px solid #bae6fd;border-radius:10px;padding:12px 16px;margin-bottom:14px;display:flex;align-items:center;gap:20px;flex-wrap:wrap'>"
$htmlEmail += "<div><span style='font-size:11px;text-transform:uppercase;color:#0369a1;font-weight:600'>Ciclo $cicloNum</span><br>"
$htmlEmail += "<span style='font-size:20px;font-weight:700;color:#0369a1'>$cicloRecorridos</span> <span style='font-size:12px;color:#6b7280'>de $($stockItems.Count) SKUs recorridos ($cicloProgresoPct%)</span></div>"
$htmlEmail += "<div style='flex:1;min-width:160px'><div style='background:#e0f2fe;border-radius:999px;height:10px'><div style='background:#0369a1;border-radius:999px;height:10px;width:$cicloProgresoPct%'></div></div></div>"
$htmlEmail += "<div><span style='font-size:16px;font-weight:700;color:#dc2626'>$cicloRestantes</span> <span style='font-size:12px;color:#6b7280'>SKUs restantes en el ciclo</span></div>"
$htmlEmail += "</div>"
$htmlEmail += "<p style='font-size:13px;color:#6b7280;margin:0 0 14px'>Criterio: <strong>menor stock web primero</strong> (m&aacute;s cr&iacute;tico arriba) + mayor valor DDP. Ning&uacute;n &iacute;tem se repite hasta completar el ciclo.</p>"
# Badges de composicion (solo los tramos que aportaron articulos)
if ($cntT1 -gt 0 -or $cntT2 -gt 0 -or $cntT3 -gt 0 -or $cntT4 -gt 0) {
    $htmlEmail += "<div style='display:flex;gap:8px;margin-bottom:18px;flex-wrap:wrap;font-size:12px'>"
    if ($cntT1 -gt 0) { $htmlEmail += "<span style='background:#fee2e2;color:#dc2626;border-radius:8px;padding:4px 10px;font-weight:600'>&#9632; Cr&iacute;tico (&lt;300 uds): $cntT1</span>" }
    if ($cntT2 -gt 0) { $htmlEmail += "<span style='background:#fff7ed;color:#d97706;border-radius:8px;padding:4px 10px;font-weight:600'>&#9632; Medio (300-499): $cntT2</span>" }
    if ($cntT3 -gt 0) { $htmlEmail += "<span style='background:#fefce8;color:#a16207;border-radius:8px;padding:4px 10px;font-weight:600'>&#9632; Normal (500-999): $cntT3</span>" }
    if ($cntT4 -gt 0) { $htmlEmail += "<span style='background:#f0f9ff;color:#0369a1;border-radius:8px;padding:4px 10px;font-weight:600'>&#9632; Alto (&ge;1000): $cntT4</span>" }
    $htmlEmail += "</div>"
}
$htmlEmail += "<table style='width:100%;border-collapse:collapse;font-size:13px'>"
$htmlEmail += "<thead><tr style='background:#f8fafc'>"
$htmlEmail += "<th style='padding:10px 12px;text-align:left;font-size:11px;text-transform:uppercase;color:#6b7280;border-bottom:2px solid #e5e7eb'>#</th>"
$htmlEmail += "<th style='padding:10px 12px;text-align:left;font-size:11px;text-transform:uppercase;color:#6b7280;border-bottom:2px solid #e5e7eb'>Art&iacute;culo</th>"
$htmlEmail += "<th style='padding:10px 12px;text-align:left;font-size:11px;text-transform:uppercase;color:#6b7280;border-bottom:2px solid #e5e7eb'>Familia</th>"
$htmlEmail += "<th style='padding:10px 12px;text-align:right;font-size:11px;text-transform:uppercase;color:#6b7280;border-bottom:2px solid #e5e7eb'>Stock Web</th>"
$htmlEmail += "<th style='padding:10px 12px;text-align:right;font-size:11px;text-transform:uppercase;color:#6b7280;border-bottom:2px solid #e5e7eb'>DDP Unit.</th>"
$htmlEmail += "<th style='padding:10px 12px;text-align:right;font-size:11px;text-transform:uppercase;color:#6b7280;border-bottom:2px solid #e5e7eb'>Val. DDP</th>"
$htmlEmail += "<th style='padding:10px 12px;text-align:left;font-size:11px;text-transform:uppercase;color:#6b7280;border-bottom:2px solid #e5e7eb'>Situaci&oacute;n</th>"
$htmlEmail += "</tr></thead><tbody>$tablaHoy</tbody></table>"

# Dashboard adjunto
$htmlEmail += "<div style='margin-top:28px;background:#f0f9ff;border:1px solid #bae6fd;border-radius:10px;padding:16px 20px'>"
$htmlEmail += "<div style='font-size:13px;color:#0369a1;font-weight:600;margin-bottom:8px'>&#128206; Archivos adjuntos</div>"
$htmlEmail += "<ul style='margin:0;padding-left:18px;font-size:12px;color:#374151;line-height:1.8'>"
$htmlEmail += "<li><strong>Conteo_Diario_$fechaTag.csv</strong> &mdash; Lista de SKUs a contar hoy (doble clic para abrir en Excel)</li>"
$htmlEmail += "<li><strong>Dashboard_Ciclico.html</strong> &mdash; Dashboard completo de indicadores (abrir con Chrome / Edge)</li>"
$htmlEmail += "</ul>"
$htmlEmail += "</div>"

$htmlEmail += "</div>"
$htmlEmail += "<div style='background:#f8fafc;padding:14px 32px;font-size:11px;color:#9ca3af;text-align:center;border-top:1px solid #e5e7eb'>Generado autom&aacute;ticamente &mdash; Sistema de Inventario C&iacute;clico Zecat ARG</div>"
$htmlEmail += "</div></body></html>"

# ── 11. Generar CSV con lista de hoy (abre directo en Excel) ─
$fechaTag = $HoyFecha -replace "/",""          # ej. 12052026
$csvFile  = Join-Path $ScriptDir "Conteo_Diario_$fechaTag.csv"
$csvOK    = $false
Write-Host "Generando lista CSV..."
try {
    $csvLines = @()
    $csvLines += "N°;SKU;Articulo;Familia;Stock Web;DDP Unitario;Valor DDP;Situacion"
    $pos = 1
    foreach ($it in $todosItems) {
        $skuS  = [string]$it.SKU
        $artS  = [string]$it.Articulo
        $famS  = [string]$it.Familia
        $swS   = [string]$it.StockWeb
        $ddpUS = $it.DdpUnit.ToString("F2")
        $ddpVS = $it.DdpVal.ToString("F2")
        $sitS  = [string]$it.Situacion
        $csvLines += "$pos;$skuS;$artS;$famS;$swS;$ddpUS;$ddpVS;$sitS"
        $pos++
    }
    $csvContent = $csvLines -join "`r`n"
    # UTF-8 con BOM para que Excel lo abra con tildes correctas
    $utf8Bom = New-Object System.Text.UTF8Encoding $true
    [System.IO.File]::WriteAllText($csvFile, $csvContent, $utf8Bom)
    $csvOK = $true
    Write-Host "Lista generada: $csvFile"
} catch {
    Write-Host "AVISO: No se pudo generar CSV ($_)"
}

# ── 12. Enviar via Outlook ────────────────────────────────────
Write-Host "Enviando email..."
try {
    $ol   = New-Object -ComObject Outlook.Application
    $mail = $ol.CreateItem(0)
    $mail.To       = $Destinatarios
    $mail.Subject  = "Conteo Ciclico Diario - $HoyFecha | Zecat ARG ($totalItems articulos)"
    $mail.HTMLBody = $htmlEmail
    # Adjuntar CSV con la lista del dia (se abre en Excel)
    if ($csvOK -and (Test-Path $csvFile)) { $mail.Attachments.Add($csvFile) | Out-Null }
    # Adjuntar el dashboard HTML
    $dashFile = Join-Path $ScriptDir "Dashboard_Ciclico.html"
    if (Test-Path $dashFile) { $mail.Attachments.Add($dashFile) | Out-Null }
    $mail.Send()
    Write-Host "Email enviado a: $Destinatarios"
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ol)|Out-Null
} catch {
    Write-Host "AVISO: No se pudo enviar via Outlook ($_). Guardando preview..."
    $prev = Join-Path $ScriptDir ("email_preview_" + ($fechaTag) + ".html")
    [System.IO.File]::WriteAllText($prev, $htmlEmail, $utf8NoBom)
    Write-Host "Preview guardado: $prev"
}

# ── 12. Actualizar dashboard ──────────────────────────────────
$dashScript = Join-Path $ScriptDir "Actualizar_Dashboard.ps1"
if (Test-Path $dashScript) {
    Write-Host "Actualizando dashboard..."
    & powershell.exe -ExecutionPolicy Bypass -File $dashScript
}

Write-Host "`n=== Items recomendados para hoy ==="
foreach ($it in $top10) {
    $sw  = $it.StockWeb
    $ddv = $it.DdpVal.ToString("N0")
    $cri = if ($it.Critico) { " [CRITICO]" } else { "" }
    Write-Host "  $sw uds | `$$ddv DDP$cri | $($it.Articulo)"
}
