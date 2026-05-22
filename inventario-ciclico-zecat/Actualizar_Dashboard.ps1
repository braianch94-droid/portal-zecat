# Actualizar_Dashboard.ps1 - Lee Ciclico ARG 2025-26.xlsx y regenera Dashboard_Ciclico.html

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ExcelFile = Join-Path $ScriptDir "Ciclico ARG 2025-26.xlsx"
$HtmlFile  = Join-Path $ScriptDir "Dashboard_Ciclico.html"

if (-not (Test-Path $ExcelFile)) {
    Write-Host "ERROR: No se encontro el archivo $ExcelFile"
    exit 1
}

function HEnc($s) {
    if ($s -eq $null) { return "" }
    return [System.Net.WebUtility]::HtmlEncode([string]$s)
}

# ── FINES DE SEMANA Y FERIADOS NACIONALES ────────────────────────
$hoyDT     = [datetime]::Today
$hoyStr    = $hoyDT.ToString("dd/MM/yyyy")
$diaSemana = [int]$hoyDT.DayOfWeek   # 0=Dom 1=Lun ... 6=Sab
if ($diaSemana -eq 0 -or $diaSemana -eq 6) {
    Write-Host "Fin de semana ($hoyStr). No se actualiza dashboard."
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
if ($feriadosARG -contains $hoyStr) {
    Write-Host "Feriado nacional: $hoyStr. No se actualiza dashboard."
    exit 0
}
# ─────────────────────────────────────────────────────────────────

# ── PAUSA VACACIONES ─────────────────────────────────────────────
$pausaDesde = [datetime]"2026-05-18"
$pausaHasta = [datetime]"2026-05-25"
if ($hoyDT -ge $pausaDesde -and $hoyDT -le $pausaHasta) {
    Write-Host "PAUSA VACACIONES: $hoyStr dentro del periodo de pausa. No se actualiza dashboard."
    exit 0
}
# ─────────────────────────────────────────────────────────────────

Write-Host "Leyendo datos de Excel..."

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$tmpFile = "$env:TEMP\ciclico_dash_tmp.xlsx"
Copy-Item $ExcelFile $tmpFile -Force

try {
    $wb = $excel.Workbooks.Open($tmpFile)
    $ws = $wb.Sheets.Item(1)
    $rows = $ws.UsedRange.Rows.Count

    $data = @()
    for ($r = 2; $r -le $rows; $r++) {
        $diaText = $ws.Cells.Item($r, 14).Text
        $ano = if ($diaText -match "(\d{4})") { $matches[1] } else { "" }
        $mesText = $ws.Cells.Item($r, 13).Text
        $mesAno = if ($mesText -ne "" -and $ano -ne "") { "$mesText $ano" } else { $mesText }
        $data += [PSCustomObject]@{
            SKU        = $ws.Cells.Item($r,  1).Text
            Articulo   = $ws.Cells.Item($r,  2).Text
            Family     = $ws.Cells.Item($r,  3).Text
            Sistema    = $ws.Cells.Item($r,  4).Value2
            Total      = $ws.Cells.Item($r,  7).Value2
            DifCont    = $ws.Cells.Item($r,  8).Value2
            Diferencia = $ws.Cells.Item($r, 10).Text
            Mes        = $mesText
            Ano        = $ano
            MesAno     = $mesAno
            Dia        = $diaText
            DDP        = $ws.Cells.Item($r, 15).Value2
            StockDDP   = $ws.Cells.Item($r, 16).Value2
            DifDDP     = $ws.Cells.Item($r, 17).Value2
        }
    }
    $wb.Close($false)
} finally {
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}

Write-Host "Procesando $($data.Count) registros..."

$total      = $data.Count
$uniqueSKUs = ($data | Where-Object { $_.SKU -ne "" } | Group-Object SKU).Count
$sinDif  = ($data | Where-Object { $_.Diferencia -eq "Sin Dif" }).Count
$difPos  = ($data | Where-Object { $_.Diferencia -eq "Dif Positiva" }).Count
$difNeg  = ($data | Where-Object { $_.Diferencia -eq "Dif Negativa" }).Count
$accGlobal     = if ($total -gt 0) { [math]::Round($sinDif / $total * 100, 1) } else { 0 }
$totalDifDDP   = [math]::Round(($data | Where-Object { $_.DifDDP -ne $null } | Measure-Object -Property DifDDP -Sum).Sum, 2)
$totalStockDDP = [math]::Round(($data | Where-Object { $_.StockDDP -ne $null } | Measure-Object -Property StockDDP -Sum).Sum, 2)

$ordenMeses = @("enero","febrero","marzo","abril","mayo","junio","julio","agosto","septiembre","octubre","noviembre","diciembre")

# Group by MesAno, sort chronologically by año then mes
$mesesPresentes = $data | Where-Object { $_.MesAno -ne "" } | Group-Object MesAno | Sort-Object {
    $parts = $_.Name -split " "
    $m = $parts[0].ToLower(); $y = if ($parts.Count -gt 1) { [int]$parts[1] } else { 0 }
    $y * 100 + $ordenMeses.IndexOf($m)
}

$jsMeses = @(); $jsConteos = @(); $jsSinDif = @(); $jsDifPos = @(); $jsDifNeg = @(); $jsAcc = @()

foreach ($g in $mesesPresentes) {
    $mSin = ($g.Group | Where-Object { $_.Diferencia -eq "Sin Dif" }).Count
    $mPos = ($g.Group | Where-Object { $_.Diferencia -eq "Dif Positiva" }).Count
    $mNeg = ($g.Group | Where-Object { $_.Diferencia -eq "Dif Negativa" }).Count
    $mAcc = if ($g.Count -gt 0) { [math]::Round($mSin / $g.Count * 100, 1) } else { 0 }
    $cap  = (Get-Culture).TextInfo.ToTitleCase($g.Name.ToLower())
    $jsMeses   += "`"$cap`""
    $jsConteos += $g.Count
    $jsSinDif  += $mSin
    $jsDifPos  += $mPos
    $jsDifNeg  += $mNeg
    $jsAcc     += $mAcc
}

$famGroups = $data | Where-Object { $_.Family -ne "" -and $_.Family -ne "0" } |
    Group-Object Family | Sort-Object Count -Descending

$jsFamLabels = @(); $jsFamAcc = @(); $jsFamImpacto = @(); $famRows = @()

foreach ($g in $famGroups) {
    $fSin  = ($g.Group | Where-Object { $_.Diferencia -eq "Sin Dif" }).Count
    $fAcc  = if ($g.Count -gt 0) { [math]::Round($fSin / $g.Count * 100, 1) } else { 0 }
    $fImp  = [math]::Round(($g.Group | Where-Object { $_.DifDDP -ne $null } | Measure-Object -Property DifDDP -Sum).Sum, 2)
    $short = $g.Name -replace "^[A-Z]-\d+ - ", ""
    $jsFamLabels  += "`"$(HEnc $short)`""
    $jsFamAcc     += $fAcc
    $jsFamImpacto += $fImp
    $famRows      += [PSCustomObject]@{ Fam=$g.Name; Cnt=$g.Count; Sin=$fSin; Acc=$fAcc; Imp=$fImp }
}

$top10 = $data | Where-Object { $_.DifDDP -ne $null -and [math]::Abs($_.DifDDP) -gt 0 } |
    Sort-Object { [math]::Abs($_.DifDDP) } -Descending | Select-Object -First 10

# ── Conteos por dia ───────────────────────────────────────────
$byDay = @{}
foreach ($row in $data) {
    if ($row.Dia -ne "") {
        if ($byDay.ContainsKey($row.Dia)) { $byDay[$row.Dia]++ } else { $byDay[$row.Dia] = 1 }
    }
}
$diasOrdenados = $byDay.GetEnumerator() | Sort-Object {
    try { [datetime]::ParseExact($_.Key, "d/MM/yyyy", $null) } catch { [datetime]::MinValue }
}
$jsDiaLabels = @(); $jsDiaCounts = @()
foreach ($d in $diasOrdenados) {
    $jsDiaLabels  += "`"$($d.Key)`""
    $jsDiaCounts  += $d.Value
}
$totalDias = $byDay.Count
$maxDia    = ($byDay.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 1)
$maxDiaLabel = $maxDia.Key; $maxDiaCount = $maxDia.Value
$promDia   = if ($totalDias -gt 0) { [math]::Round($total / $totalDias, 1) } else { 0 }

# ── Ranking articulos mas contados ───────────────────────────
$byArticulo = @{}
foreach ($row in $data) {
    if ($row.Articulo -ne "") {
        if ($byArticulo.ContainsKey($row.Articulo)) { $byArticulo[$row.Articulo]++ } else { $byArticulo[$row.Articulo] = 1 }
    }
}
$topArticulos = $byArticulo.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 20
$artTableRows = ""
$artRank = 1
foreach ($a in $topArticulos) {
    $artName = HEnc $a.Key
    $artCount = $a.Value
    $barPct = [math]::Round($artCount / $maxDia.Value * 100)
    $artTableRows += "<tr><td style='color:#9ca3af;font-weight:600;width:32px'>$artRank</td><td>$artName</td><td style='text-align:right;font-weight:700;color:#2563eb;width:60px'>$artCount</td><td style='width:160px'><div style='background:#e5e7eb;border-radius:4px;height:8px'><div style='background:#2563eb;border-radius:4px;height:8px;width:$barPct%'></div></div></td></tr>`n"
    $artRank++
}

# ── Cobertura por familia (desde JSON generado por conteo diario) ──
$coberturaData = $null
$cobFechaGen   = ""
$cobGlobal     = 0
$cobTotalElig  = 0
$cobTotalCnt   = 0
$cobFamRows    = ""
$jsCobLabels   = @(); $jsCobContados = @(); $jsCobPendientes = @()

$cobFile = Join-Path $ScriptDir "cobertura_familias.json"
if (Test-Path $cobFile) {
    try {
        $cobJson = Get-Content $cobFile -Raw -Encoding UTF8 | ConvertFrom-Json
        $cobFechaGen  = $cobJson.fechaGenerado
        $cobGlobal    = $cobJson.coberturaGlobal
        $cobTotalElig = $cobJson.totalElegibles
        $cobTotalCnt  = $cobJson.totalContados
        $diasRestantes = [math]::Round(($cobTotalElig - $cobTotalCnt) / 13)

        foreach ($fam in $cobJson.familias | Sort-Object pct) {
            $fn  = HEnc $fam.familia
            $tot = $fam.total; $cnt = $fam.contados; $pen = $fam.pendientes; $pct = $fam.pct
            $barColor  = if ($pct -ge 70) { "#16a34a" } elseif ($pct -ge 40) { "#d97706" } else { "#ef4444" }
            $pctColor  = if ($pct -ge 70) { "#16a34a" } elseif ($pct -ge 40) { "#d97706" } else { "#dc2626" }
            $cobFamRows += "<tr style='border-bottom:1px solid #f3f4f6'>"
            $cobFamRows += "<td style='padding:10px 12px;font-weight:600'>$fn</td>"
            $cobFamRows += "<td style='padding:10px 12px;text-align:right'>$tot</td>"
            $cobFamRows += "<td style='padding:10px 12px;text-align:right;color:#16a34a;font-weight:600'>$cnt</td>"
            $cobFamRows += "<td style='padding:10px 12px;text-align:right;color:#dc2626'>$pen</td>"
            $cobFamRows += "<td style='padding:10px 12px;text-align:right;font-weight:700;color:$pctColor'>$pct%</td>"
            $cobFamRows += "<td style='padding:10px 14px;width:160px'><div style='background:#e5e7eb;border-radius:4px;height:10px'><div style='background:$barColor;border-radius:4px;height:10px;width:$pct%'></div></div></td>"
            $cobFamRows += "</tr>`n"
            $shortFam = $fam.familia -replace "^[A-Z]-\d+ - ",""
            $jsCobLabels    += "`"$(HEnc $shortFam)`""
            $jsCobContados  += $cnt
            $jsCobPendientes+= $pen
        }
    } catch { Write-Host "AVISO: No se pudo leer cobertura_familias.json" }
}

$jsCobLblArr  = "[" + ($jsCobLabels    -join ",") + "]"
$jsCobCntArr  = "[" + ($jsCobContados  -join ",") + "]"
$jsCobPenArr  = "[" + ($jsCobPendientes-join ",") + "]"

$primerMes    = if ($jsMeses.Count -gt 0) { $jsMeses[0].Trim('"') } else { "" }
$ultimoMes    = if ($mesesPresentes.Count -gt 0) { $mesesPresentes[-1].Name } else { "" }
$accUltimo    = if ($jsAcc.Count -gt 0) { $jsAcc[-1] } else { 0 }
$fechaUpdate  = Get-Date -Format "dd/MM/yyyy HH:mm"

function ToJsArray($arr) { "[" + ($arr -join ",") + "]" }

$jsMesesArr   = ToJsArray $jsMeses
$jsConteosArr = ToJsArray $jsConteos
$jsSinDifArr  = ToJsArray $jsSinDif
$jsDifPosArr  = ToJsArray $jsDifPos
$jsDifNegArr  = ToJsArray $jsDifNeg
$jsAccArr     = ToJsArray $jsAcc
$jsFamLblArr  = ToJsArray $jsFamLabels
$jsFamAccArr  = ToJsArray $jsFamAcc
$jsFamImpArr  = ToJsArray $jsFamImpacto
$jsDiaLblArr  = ToJsArray $jsDiaLabels
$jsDiaCntArr  = ToJsArray $jsDiaCounts

# ── Bloques HTML de cobertura ─────────────────────────────────
$coberturaKpis   = ""
$seccionCobertura= ""
if ($cobTotalElig -gt 0) {
    $diasR    = [math]::Round(($cobTotalElig - $cobTotalCnt) / 13)
    $cobFalt  = $cobTotalElig - $cobTotalCnt
    $cobColor = if ($cobGlobal -ge 70) { "#16a34a" } elseif ($cobGlobal -ge 40) { "#d97706" } else { "#dc2626" }
    $progBar  = [math]::Min(100, $cobGlobal)

    $coberturaKpis = @"
<!-- Barra de progreso total -->
<div style="background:white;border-radius:12px;padding:22px 28px;box-shadow:0 2px 8px rgba(0,0,0,.06);margin-bottom:20px">
  <div style="display:flex;justify-content:space-between;align-items:baseline;margin-bottom:10px">
    <div>
      <span style="font-size:14px;font-weight:700;color:#1a1a2e">Progreso Total del Inventario C&iacute;clico</span>
      <span style="font-size:12px;color:#9ca3af;margin-left:12px">SKUs contados al menos 1 vez sobre el universo elegible</span>
    </div>
    <div style="font-size:28px;font-weight:700;color:$cobColor">$cobGlobal%</div>
  </div>
  <div style="background:#e5e7eb;border-radius:999px;height:18px;overflow:hidden;margin-bottom:14px">
    <div style="background:linear-gradient(90deg,$cobColor,$(if($cobGlobal -ge 70){'#4ade80'} elseif($cobGlobal -ge 40){'#fbbf24'} else {'#f87171'}));height:18px;width:$progBar%;border-radius:999px;transition:width .5s"></div>
  </div>
  <div style="display:flex;gap:32px;flex-wrap:wrap">
    <div><span style="font-size:22px;font-weight:700;color:#16a34a">$cobTotalCnt</span> <span style="font-size:12px;color:#6b7280">contados</span></div>
    <div><span style="font-size:22px;font-weight:700;color:#dc2626">$cobFalt</span> <span style="font-size:12px;color:#6b7280">pendientes</span></div>
    <div><span style="font-size:22px;font-weight:700;color:#1a1a2e">$cobTotalElig</span> <span style="font-size:12px;color:#6b7280">universo total</span></div>
    <div style="margin-left:auto;text-align:right"><span style="font-size:18px;font-weight:700;color:#7c3aed">$diasR d&iacute;as</span><br><span style="font-size:11px;color:#9ca3af">para completar a 13 SKUs/d&iacute;a</span></div>
  </div>
</div>

<!-- KPIs de cobertura -->
<div class="kpi-grid" style="grid-template-columns:repeat(4,1fr);margin-bottom:24px">
  <div class="kpi-card green">
    <div class="kpi-label">Cobertura Global</div>
    <div class="kpi-value" style="color:$cobColor">$cobGlobal%</div>
    <div class="kpi-sub">de elegibles contados al menos 1 vez</div>
  </div>
  <div class="kpi-card blue">
    <div class="kpi-label">Total Elegibles</div>
    <div class="kpi-value">$cobTotalElig</div>
    <div class="kpi-sub">activos en web (con y sin stock)</div>
  </div>
  <div class="kpi-card amber">
    <div class="kpi-label">Pendientes de Contar</div>
    <div class="kpi-value amber">$cobFalt</div>
    <div class="kpi-sub">no contados en los &uacute;ltimos 30 d&iacute;as</div>
  </div>
  <div class="kpi-card purple">
    <div class="kpi-label">D&iacute;as para Completar</div>
    <div class="kpi-value" style="font-size:22px">$diasR</div>
    <div class="kpi-sub">a 13 SKUs/d&iacute;a (promedio objetivo)</div>
  </div>
</div>
"@

    $seccionCobertura = @"
<div class="charts-row" style="grid-template-columns:1fr">
  <div class="chart-card">
    <div class="chart-title">Progreso de Cobertura por Familia</div>
    <div class="chart-subtitle">Art&iacute;culos contados vs. pendientes por familia &mdash; actualizado $cobFechaGen</div>
    <div style="position:relative;height:300px"><canvas id="chartCobertura"></canvas></div>
  </div>
</div>

<div class="table-card">
  <div class="chart-title" style="margin-bottom:4px">Detalle de Cobertura por Familia</div>
  <div class="chart-subtitle" style="margin-bottom:16px">Progreso acumulado del per&iacute;odo &mdash; elegibles activos en web</div>
  <table>
    <thead><tr>
      <th>Familia</th>
      <th style="text-align:right">Total Elegibles</th>
      <th style="text-align:right;color:#16a34a">Contados</th>
      <th style="text-align:right;color:#dc2626">Pendientes</th>
      <th style="text-align:right">Cobertura</th>
      <th style="width:160px">Progreso</th>
    </tr></thead>
    <tbody>$cobFamRows</tbody>
  </table>
</div>
"@
}

# Top 10 table rows
$topTableRows = ""
$rank = 1
foreach ($row in $top10) {
    $cls = if ($row.Diferencia -eq "Dif Positiva") { "badge-pos" } else { "badge-neg" }
    $uni = if ($row.DifCont -ne $null) { [math]::Abs([int]$row.DifCont) } else { 0 }
    $ddp = if ($row.DDP -ne $null) { '$' + ([math]::Round($row.DDP, 2)).ToString("N2") } else { "-" }
    $imp = '$' + ([math]::Abs($row.DifDDP)).ToString("N2")
    $topTableRows += "<tr><td style='color:#9ca3af;font-weight:600'>$rank</td><td class='sku-text'>$(HEnc $row.SKU)</td><td style='max-width:260px'>$(HEnc $row.Articulo)</td><td><span class='$cls'>$(HEnc $row.Diferencia)</span></td><td style='text-align:right;font-weight:600'>$uni</td><td style='text-align:right'>$ddp</td><td style='text-align:right;font-weight:700;color:#d97706'>$imp</td><td><span class='badge-sin'>$(HEnc $row.Mes)</span></td></tr>`n"
    $rank++
}

# Family table rows
$famTableRows = ""
foreach ($row in $famRows) {
    $accColor = if ($row.Acc -lt 25) { "#dc2626" } elseif ($row.Acc -lt 40) { "#d97706" } else { "#16a34a" }
    $impColor = if ($row.Imp -ge 0) { "#d97706" } else { "#dc2626" }
    $famTableRows += "<tr><td style='font-weight:600'>$(HEnc $row.Fam)</td><td style='text-align:right'>$($row.Cnt)</td><td style='text-align:right'>$($row.Sin)</td><td style='text-align:right;font-weight:700;color:$accColor'>$($row.Acc.ToString('0.0'))%</td><td style='text-align:right;font-weight:700;color:$impColor'>`$$($row.Imp.ToString('N2'))</td></tr>`n"
}

$pctPos = [math]::Round($difPos / $total * 100, 1)
$pctNeg = [math]::Round($difNeg / $total * 100, 1)

Write-Host "Generando HTML..."

$html = @"
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Dashboard Inventario C&iacute;clico | Zecat ARG 2025-26</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<style>
:root{--bg:#f0f2f5;--card:#fff;--text:#1a1a2e;--muted:#6b7280;--faint:#9ca3af;--border:#e5e7eb;--border-lt:#f3f4f6;--thead:#f8fafc;--hover:#f8fafc;--shadow:rgba(0,0,0,.06);--shadow2:rgba(0,0,0,.1);}
html[data-theme=dark]{--bg:#080B15;--card:#0d1117;--text:#e6edf3;--muted:#8b949e;--faint:#6e7681;--border:#1f2937;--border-lt:#161b22;--thead:#0d1117;--hover:#111827;--shadow:rgba(0,0,0,.4);--shadow2:rgba(0,0,0,.6);}
*{margin:0;padding:0;box-sizing:border-box;}
body{font-family:'Inter',Arial,sans-serif;background:var(--bg);color:var(--text);transition:background .25s,color .25s;}
header{background:var(--card);color:var(--text);padding:18px 32px;display:flex;align-items:center;justify-content:space-between;border-bottom:1px solid var(--border);position:relative;}
header::before{content:'';position:absolute;top:0;left:0;right:0;height:2px;background:linear-gradient(90deg,transparent 0%,#F97171 50%,transparent 100%);}
header h1{font-size:20px;font-weight:700;letter-spacing:.2px;}
header .subtitle{font-size:12px;color:var(--muted);margin-top:2px;}
header .badge{background:rgba(249,113,113,.07);border:1px solid rgba(249,113,113,.18);border-radius:8px;padding:6px 14px;font-size:11.5px;text-align:right;line-height:1.7;color:var(--muted);}
.theme-btn{background:rgba(249,113,113,.07);border:1px solid rgba(249,113,113,.18);border-radius:50%;width:38px;height:38px;font-size:18px;cursor:pointer;margin-left:12px;flex-shrink:0;transition:background .2s,transform .2s;line-height:1;}
.theme-btn:hover{background:rgba(249,113,113,.18);transform:scale(1.08);}
.container{max-width:1400px;margin:0 auto;padding:24px;}
.kpi-grid{display:grid;grid-template-columns:repeat(5,1fr);gap:16px;margin-bottom:24px;}
.kpi-card{background:var(--card);border-radius:12px;padding:18px 20px;box-shadow:0 2px 8px var(--shadow);border-top:3px solid var(--border);transition:transform .2s,box-shadow .2s;}
.kpi-card:hover{transform:translateY(-2px);box-shadow:0 6px 16px var(--shadow2);}
.kpi-card.blue  {border-top-color:#F97171;}
.kpi-card.green {border-top-color:#16a34a;}
.kpi-card.red   {border-top-color:#dc2626;}
.kpi-card.amber {border-top-color:#d97706;}
.kpi-card.purple{border-top-color:#7c3aed;}
.kpi-label{font-size:11px;text-transform:uppercase;letter-spacing:.8px;color:var(--muted);font-weight:600;margin-bottom:8px;}
.kpi-value{font-size:28px;font-weight:700;color:var(--text);line-height:1;}
.kpi-value.amber{color:#d97706;}
.kpi-value.red  {color:#dc2626;}
.kpi-sub{font-size:12px;color:var(--faint);margin-top:6px;}
.charts-row{display:grid;gap:20px;margin-bottom:20px;}
.charts-row-2{grid-template-columns:1fr 1fr;}
.charts-row-3{grid-template-columns:1.2fr 1fr 1fr;}
.chart-card{background:var(--card);border-radius:12px;padding:22px;box-shadow:0 2px 8px var(--shadow);}
.chart-title{font-size:14px;font-weight:700;color:var(--text);margin-bottom:4px;}
.chart-subtitle{font-size:11px;color:var(--faint);margin-bottom:18px;}
.table-card{background:var(--card);border-radius:12px;padding:22px;box-shadow:0 2px 8px var(--shadow);margin-bottom:20px;}
table{width:100%;border-collapse:collapse;font-size:13px;}
thead tr{background:var(--thead);}
thead th{padding:10px 12px;text-align:left;font-size:11px;text-transform:uppercase;letter-spacing:.5px;color:var(--muted);font-weight:600;border-bottom:2px solid var(--border);}
tbody tr{border-bottom:1px solid var(--border-lt);transition:background .15s;}
tbody tr:hover{background:var(--hover);}
tbody td{padding:10px 12px;color:var(--text);}
.badge-pos{background:#dcfce7;color:#16a34a;border-radius:12px;padding:2px 10px;font-size:11px;font-weight:600;white-space:nowrap;}
.badge-neg{background:#fee2e2;color:#dc2626;border-radius:12px;padding:2px 10px;font-size:11px;font-weight:600;white-space:nowrap;}
.badge-sin{background:#e0f2fe;color:#0284c7;border-radius:12px;padding:2px 10px;font-size:11px;font-weight:600;white-space:nowrap;}
html[data-theme=dark] .badge-pos{background:#14532d;color:#4ade80;}
html[data-theme=dark] .badge-neg{background:#450a0a;color:#f87171;}
html[data-theme=dark] .badge-sin{background:#0c2a4a;color:#38bdf8;}
.sku-text{font-family:monospace;font-size:12px;color:var(--muted);}
.alert-bar{background:linear-gradient(90deg,#fee2e2,#fff7ed);border:1px solid #fca5a5;border-radius:10px;padding:14px 20px;margin-bottom:20px;display:flex;align-items:center;gap:12px;}
.alert-text{font-size:13px;color:#7f1d1d;}
html[data-theme=dark] .alert-bar{background:linear-gradient(90deg,#450a0a,#3d1f00);border-color:#7f1d1d;}
html[data-theme=dark] .alert-text{color:#fca5a5;}
html[data-theme=dark] div[style*="background:white"]{background:#0d1117 !important;}
html[data-theme=dark] div[style*="background:#e5e7eb"]{background:#1f2937 !important;}
html[data-theme=dark] span[style*="color:#9ca3af"],html[data-theme=dark] div[style*="color:#9ca3af"]{color:#6e7681 !important;}
html[data-theme=dark] span[style*="color:#6b7280"],html[data-theme=dark] div[style*="color:#6b7280"]{color:#8b949e !important;}
html[data-theme=dark] span[style*="color:#1a1a2e"],html[data-theme=dark] div[style*="color:#1a1a2e"]{color:#e6edf3 !important;}
footer{text-align:center;padding:20px;font-size:11px;color:var(--faint);}
@media(max-width:1100px){.kpi-grid{grid-template-columns:repeat(3,1fr)}.charts-row-3{grid-template-columns:1fr 1fr}}
@media(max-width:700px){.kpi-grid{grid-template-columns:1fr 1fr}.charts-row-2,.charts-row-3{grid-template-columns:1fr}}
</style>
<script>(function(){var t=localStorage.getItem('theme')||'light';document.documentElement.setAttribute('data-theme',t);})();</script>
</head>
<body>
<header>
  <div>
    <div class="subtitle">ZECAT &mdash; Art&iacute;culos Promocionales SA</div>
    <h1>Dashboard Inventario C&iacute;clico ARG 2025-26</h1>
  </div>
  <div style="display:flex;align-items:center;gap:10px">
    <div class="badge">Actualizado: $fechaUpdate<br>Datos acumulados: $primerMes &mdash; $ultimoMes</div>
    <button class="theme-btn" id="themeBtn" onclick="toggleTheme()" title="Cambiar tema claro / oscuro">&#x1F319;</button>
  </div>
</header>

<div class="container">

<div class="alert-bar">
  <span style="font-size:20px">&#9888;&#65039;</span>
  <div class="alert-text">
    Tendencia cr&iacute;tica: La exactitud pas&oacute; de <strong>$($jsAcc[0])% en $primerMes</strong>
    a <strong>$accUltimo% en $ultimoMes</strong>.
    El impacto econ&oacute;mico acumulado es de <strong>`$$($totalDifDDP.ToString("N0"))</strong>.
  </div>
</div>

<div class="kpi-grid">
  <div class="kpi-card blue">
    <div class="kpi-label">SKUs &Uacute;nicos Contados</div>
    <div class="kpi-value">$($uniqueSKUs.ToString("N0"))</div>
    <div class="kpi-sub">$($total.ToString("N0")) conteos totales (con repetidos)</div>
  </div>
  <div class="kpi-card green">
    <div class="kpi-label">Exactitud Total</div>
    <div class="kpi-value amber">$accGlobal%</div>
    <div class="kpi-sub">$sinDif &iacute;tems sin diferencia</div>
  </div>
  <div class="kpi-card amber">
    <div class="kpi-label">Diferencia Positiva</div>
    <div class="kpi-value amber">$difPos</div>
    <div class="kpi-sub">$pctPos% &mdash; M&aacute;s f&iacute;sico que sistema</div>
  </div>
  <div class="kpi-card red">
    <div class="kpi-label">Diferencia Negativa</div>
    <div class="kpi-value red">$difNeg</div>
    <div class="kpi-sub">$pctNeg% &mdash; Menos f&iacute;sico que sistema</div>
  </div>
  <div class="kpi-card purple">
    <div class="kpi-label">Impacto Econ&oacute;mico</div>
    <div class="kpi-value" style="font-size:22px">`$$($totalDifDDP.ToString("N0"))</div>
    <div class="kpi-sub">Stock total `$$($totalStockDDP.ToString("N0"))</div>
  </div>
</div>

<div class="kpi-grid" style="grid-template-columns:repeat(3,1fr);margin-bottom:24px">
  <div class="kpi-card blue">
    <div class="kpi-label">D&iacute;as con Conteo</div>
    <div class="kpi-value">$totalDias</div>
    <div class="kpi-sub">Jornadas activas en el per&iacute;odo</div>
  </div>
  <div class="kpi-card amber">
    <div class="kpi-label">Promedio por D&iacute;a</div>
    <div class="kpi-value">$promDia</div>
    <div class="kpi-sub">SKUs contados por jornada</div>
  </div>
  <div class="kpi-card green">
    <div class="kpi-label">D&iacute;a M&aacute;s Activo</div>
    <div class="kpi-value" style="font-size:20px">$maxDiaLabel</div>
    <div class="kpi-sub">$maxDiaCount SKUs contados</div>
  </div>
</div>

$coberturaKpis

<div class="charts-row charts-row-3">
  <div class="chart-card">
    <div class="chart-title">Evoluci&oacute;n de Exactitud por Mes</div>
    <div class="chart-subtitle">% de SKUs sin diferencia vs. total contado por mes</div>
    <div style="position:relative;height:220px"><canvas id="chartTrend"></canvas></div>
  </div>
  <div class="chart-card">
    <div class="chart-title">Distribuci&oacute;n de Diferencias</div>
    <div class="chart-subtitle">$total conteos totales &mdash; $uniqueSKUs SKUs &uacute;nicos</div>
    <div style="position:relative;height:220px"><canvas id="chartDonut"></canvas></div>
  </div>
  <div class="chart-card">
    <div class="chart-title">Conteos por Mes</div>
    <div class="chart-subtitle">Cantidad de SKUs contados por mes</div>
    <div style="position:relative;height:220px"><canvas id="chartConteos"></canvas></div>
  </div>
</div>

<div class="charts-row charts-row-2">
  <div class="chart-card">
    <div class="chart-title">Exactitud por Familia</div>
    <div class="chart-subtitle">% de SKUs sin diferencia &mdash; l&iacute;nea verde = 70% objetivo</div>
    <div style="position:relative;height:260px"><canvas id="chartFamily"></canvas></div>
  </div>
  <div class="chart-card">
    <div class="chart-title">Impacto Econ&oacute;mico por Familia (DDP)</div>
    <div class="chart-subtitle">Suma de diferencias en valor DDP &mdash; naranja=sobrante, rojo=faltante</div>
    <div style="position:relative;height:260px"><canvas id="chartImpact"></canvas></div>
  </div>
</div>

<div class="charts-row" style="grid-template-columns:1fr">
  <div class="chart-card">
    <div class="chart-title">Detalle por Mes &mdash; Composici&oacute;n de Conteos</div>
    <div class="chart-subtitle">Barras apiladas: Sin Diferencia / Positiva / Negativa</div>
    <div style="position:relative;height:220px"><canvas id="chartStacked"></canvas></div>
  </div>
</div>

<div class="table-card">
  <div class="chart-title" style="margin-bottom:4px">Top 10 SKUs con Mayor Impacto Econ&oacute;mico</div>
  <div class="chart-subtitle" style="margin-bottom:16px">Ordenados por valor absoluto de diferencia DDP</div>
  <table>
    <thead><tr><th>#</th><th>SKU WMS</th><th>Art&iacute;culo</th><th>Diferencia</th><th style="text-align:right">Unidades</th><th style="text-align:right">Costo DDP</th><th style="text-align:right">Impacto `$</th><th>Mes</th></tr></thead>
    <tbody>$topTableRows</tbody>
  </table>
</div>

<div class="table-card">
  <div class="chart-title" style="margin-bottom:4px">Resumen por Familia</div>
  <div class="chart-subtitle" style="margin-bottom:16px">Exactitud y exposici&oacute;n econ&oacute;mica por l&iacute;nea de producto</div>
  <table>
    <thead><tr><th>Familia</th><th style="text-align:right">SKUs</th><th style="text-align:right">Sin Dif.</th><th style="text-align:right">Exactitud</th><th style="text-align:right">Impacto `$</th></tr></thead>
    <tbody>$famTableRows</tbody>
  </table>
</div>

$seccionCobertura

<div class="charts-row" style="grid-template-columns:1fr">
  <div class="chart-card">
    <div class="chart-title">Conteos por D&iacute;a &mdash; Actividad Diaria</div>
    <div class="chart-subtitle">Cantidad de SKUs contados en cada jornada de inventario ($totalDias d&iacute;as activos)</div>
    <div style="position:relative;height:220px"><canvas id="chartDia"></canvas></div>
  </div>
</div>

<div class="charts-row charts-row-2">
  <div class="table-card" style="margin-bottom:0">
    <div class="chart-title" style="margin-bottom:4px">Ranking de Art&iacute;culos M&aacute;s Contados</div>
    <div class="chart-subtitle" style="margin-bottom:16px">Top 20 art&iacute;culos por frecuencia de conteo</div>
    <table>
      <thead><tr><th>#</th><th>Art&iacute;culo</th><th style="text-align:right">Veces</th><th>Frecuencia</th></tr></thead>
      <tbody>$artTableRows</tbody>
    </table>
  </div>
  <div class="chart-card" style="display:flex;flex-direction:column">
    <div class="chart-title" style="margin-bottom:4px">Conteos por Mes &mdash; Detalle</div>
    <div class="chart-subtitle" style="margin-bottom:18px">Distribuci&oacute;n de conteos y exactitud mensual</div>
    <div style="position:relative;flex:1;min-height:300px"><canvas id="chartMesDetalle"></canvas></div>
  </div>
</div>

</div>
<footer>Dashboard generado autom&aacute;ticamente desde Ciclico ARG 2025-26.xlsx &nbsp;|&nbsp; Zecat &mdash; Art&iacute;culos Promocionales SA &nbsp;|&nbsp; $fechaUpdate</footer>

<script>
function _isDark(){return document.documentElement.getAttribute('data-theme')==='dark';}
function _gc(){return _isDark()?'#1f2937':'#f3f4f6';}
Chart.defaults.font.family='Inter,Arial,sans-serif';
Chart.defaults.font.size=12;
Chart.defaults.color=_isDark()?'#8b949e':'#555';
const meses=$jsMesesArr, conteos=$jsConteosArr, sinDif=$jsSinDifArr,
      difPos=$jsDifPosArr, difNeg=$jsDifNegArr, acc=$jsAccArr,
      famL=$jsFamLblArr, famAcc=$jsFamAccArr, famImp=$jsFamImpArr;
const _charts=[];

_charts.push(new Chart('chartTrend',{type:'line',data:{labels:meses,datasets:[
  {label:'% Exactitud',data:acc,borderColor:'#F97171',backgroundColor:'rgba(249,113,113,.08)',borderWidth:2.5,tension:.3,fill:true,
   pointBackgroundColor:acc.map(v=>v<30?'#dc2626':v<50?'#d97706':'#16a34a'),pointRadius:5},
  {label:'Objetivo 70%',data:Array(meses.length).fill(70),borderColor:'#16a34a',borderWidth:1.5,borderDash:[6,4],pointRadius:0,fill:false}
]},options:{responsive:true,maintainAspectRatio:false,
  plugins:{legend:{position:'bottom',labels:{boxWidth:12,font:{size:11}}}},
  scales:{y:{min:0,max:100,ticks:{callback:v=>v+'%'},grid:{color:_gc()}},x:{grid:{display:false}}}}}));

_charts.push(new Chart('chartDonut',{type:'doughnut',data:{
  labels:['Sin Diferencia','Dif. Positiva','Dif. Negativa'],
  datasets:[{data:[$sinDif,$difPos,$difNeg],backgroundColor:['#F97171','#f59e0b','#ef4444'],borderWidth:2,borderColor:'#fff'}]
},options:{responsive:true,maintainAspectRatio:false,cutout:'62%',
  plugins:{legend:{position:'bottom',labels:{boxWidth:12,font:{size:11},padding:12}},
  tooltip:{callbacks:{label:ctx=>' '+ctx.label+': '+ctx.raw+' SKUs'}}}}}));

_charts.push(new Chart('chartConteos',{type:'bar',data:{labels:meses,datasets:[
  {label:'SKUs contados',data:conteos,backgroundColor:conteos.map((_,i)=>'hsla('+(220+i*10)+',70%,55%,.85)'),borderRadius:5}
]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},
  scales:{y:{grid:{color:_gc()}},x:{grid:{display:false}}}}}));

_charts.push(new Chart('chartFamily',{type:'bar',data:{labels:famL,datasets:[
  {label:'% Exactitud',data:famAcc,backgroundColor:famAcc.map(v=>v<25?'#ef4444':v<40?'#f59e0b':v<55?'#3b82f6':'#16a34a'),borderRadius:5},
  {label:'Objetivo 70%',data:Array(famL.length).fill(70),type:'line',borderColor:'#16a34a',borderWidth:2,borderDash:[5,4],pointRadius:0,fill:false}
]},options:{responsive:true,maintainAspectRatio:false,
  plugins:{legend:{position:'bottom',labels:{boxWidth:12,font:{size:11}}}},
  scales:{y:{min:0,max:100,ticks:{callback:v=>v+'%'},grid:{color:_gc()}},x:{grid:{display:false},ticks:{font:{size:10}}}}}}));

_charts.push(new Chart('chartImpact',{type:'bar',data:{labels:famL,datasets:[
  {label:'Impacto DDP',data:famImp,backgroundColor:famImp.map(v=>v>=0?'rgba(245,158,11,.8)':'rgba(239,68,68,.8)'),borderRadius:4}
]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},
  tooltip:{callbacks:{label:ctx=>' \$'+ctx.raw.toLocaleString('es-AR',{maximumFractionDigits:0})}}},
  scales:{y:{grid:{color:_gc()},ticks:{callback:v=>'\$'+v.toLocaleString('es-AR',{maximumFractionDigits:0})}},x:{grid:{display:false},ticks:{font:{size:10}}}}}}));

_charts.push(new Chart('chartStacked',{type:'bar',data:{labels:meses,datasets:[
  {label:'Sin Diferencia',data:sinDif,backgroundColor:'#F97171'},
  {label:'Dif. Positiva', data:difPos,backgroundColor:'#f59e0b'},
  {label:'Dif. Negativa', data:difNeg,backgroundColor:'#ef4444'}
]},options:{responsive:true,maintainAspectRatio:false,
  plugins:{legend:{position:'bottom',labels:{boxWidth:12,font:{size:11}}}},
  scales:{x:{stacked:true,grid:{display:false}},y:{stacked:true,grid:{color:_gc()}}}}}));

const diaL=$jsDiaLblArr, diaC=$jsDiaCntArr;
_charts.push(new Chart('chartDia',{type:'bar',data:{labels:diaL,datasets:[
  {label:'SKUs contados',data:diaC,
   backgroundColor:diaC.map(v=>v>150?'#dc2626':v>80?'#f59e0b':'#F97171'),
   borderRadius:3}
]},options:{responsive:true,maintainAspectRatio:false,
  plugins:{legend:{display:false},tooltip:{callbacks:{label:ctx=>' '+ctx.raw+' SKUs'}}},
  scales:{y:{grid:{color:_gc()},ticks:{font:{size:11}}},
    x:{grid:{display:false},ticks:{font:{size:9},maxRotation:45,minRotation:30}}}}}));

if (document.getElementById('chartCobertura')) {
  _charts.push(new Chart('chartCobertura',{type:'bar',data:{labels:$jsCobLblArr,datasets:[
    {label:'Contados',   data:$jsCobCntArr, backgroundColor:'#16a34a',borderRadius:4},
    {label:'Pendientes', data:$jsCobPenArr, backgroundColor:'#e5e7eb',borderRadius:4}
  ]},options:{responsive:true,maintainAspectRatio:false,
    plugins:{legend:{position:'bottom',labels:{boxWidth:12,font:{size:11}}}},
    scales:{x:{stacked:true,grid:{display:false},ticks:{font:{size:10}}},
      y:{stacked:true,grid:{color:_gc()},ticks:{font:{size:11}}}}}}));
}

_charts.push(new Chart('chartMesDetalle',{type:'bar',data:{labels:meses,datasets:[
  {label:'Total contados',data:conteos,backgroundColor:'rgba(249,113,113,.15)',borderColor:'#F97171',borderWidth:1.5,borderRadius:4,yAxisID:'y'},
  {label:'% Exactitud',data:acc,type:'line',borderColor:'#16a34a',backgroundColor:'rgba(22,163,74,.1)',
   borderWidth:2,tension:.3,fill:true,pointRadius:4,pointBackgroundColor:acc.map(v=>v<30?'#dc2626':v<50?'#f59e0b':'#16a34a'),yAxisID:'y1'}
]},options:{responsive:true,maintainAspectRatio:false,
  plugins:{legend:{position:'bottom',labels:{boxWidth:12,font:{size:11}}}},
  scales:{
    y:{position:'left',grid:{color:_gc()},title:{display:true,text:'SKUs',font:{size:10}}},
    y1:{position:'right',min:0,max:100,grid:{display:false},ticks:{callback:v=>v+'%'},title:{display:true,text:'Exactitud',font:{size:10}}},
    x:{grid:{display:false}}}}}));

function toggleTheme(){
  var h=document.documentElement;
  var dark=h.getAttribute('data-theme')==='dark';
  var next=dark?'light':'dark';
  h.setAttribute('data-theme',next);
  localStorage.setItem('theme',next);
  var nowDark=next==='dark';
  var gc=nowDark?'#1f2937':'#f3f4f6';
  Chart.defaults.color=nowDark?'#8b949e':'#555';
  _charts.forEach(function(c){
    if(!c||!c.options)return;
    var sc=c.options.scales||{};
    Object.values(sc).forEach(function(ax){if(ax.grid)ax.grid.color=gc;});
    c.update('none');
  });
  var b=document.getElementById('themeBtn');
  if(b)b.textContent=nowDark?'\u2600\uFE0F':'\u{1F319}';
}
(function(){
  var t=localStorage.getItem('theme')||'light';
  var b=document.getElementById('themeBtn');
  if(b) b.textContent=t==='dark'?'\u2600\uFE0F':'\u{1F319}';
})();
</script>
</body></html>
"@

$utf8NoBom = New-Object System.Text.UTF8Encoding $false
[System.IO.File]::WriteAllText($HtmlFile, $html, $utf8NoBom)
Write-Host "Dashboard guardado en: $HtmlFile"

# ── PUBLICAR EN GITHUB ────────────────────────────────────────────
Write-Host "Publicando dashboard en GitHub..."
Push-Location $ScriptDir
try {
    $indexFile = Join-Path $ScriptDir "index.html"
    Copy-Item $HtmlFile $indexFile -Force
    git add "index.html" 2>$null
    git diff --cached --quiet 2>$null
    if ($LASTEXITCODE -ne 0) {
        $fecha = (Get-Date).ToString("dd/MM/yyyy HH:mm")
        git commit -m "dashboard: actualizar $fecha" 2>$null
        git push 2>$null
        Write-Host "Publicado en GitHub Pages."
    } else {
        Write-Host "Sin cambios nuevos para publicar."
    }
} catch {
    Write-Host "AVISO: No se pudo publicar en GitHub. El dashboard local fue generado igual."
} finally {
    Pop-Location
}
# ─────────────────────────────────────────────────────────────────

Write-Host "Listo!"
