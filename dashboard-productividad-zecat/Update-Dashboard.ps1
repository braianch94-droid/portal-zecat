# ============================================================
# Update-Dashboard.ps1  v5
# GRUPOS: SIN LOGO/CON LOGO = picking regular (target 84)
#         PIE DE MAQUINA    = hoja propia, MERMA1/2, sin target
#         PEDIDO MERMA      = hoja propia, MERMA1/2, sin target
#         MUESTRA SIMPLE    = LEZCANO AGUSTIN, hoja propia
# STAFFING: solo lineas de picking regular (SIN/CON LOGO)
# ============================================================
param(
    [string]$Source  = "$PSScriptRoot\productividad.xlsx",
    [string]$Output  = "$PSScriptRoot\Dashboard_Productividad.xlsx",
    [string]$LezFilter = "Lezcano"
)

$TARGET      = 84   # Líneas/operario/día para verde
$TARGET_WARN = 70   # Umbral mínimo aceptable (naranja); ajustar junto con TARGET si cambia
$NOW     = Get-Date

# Operarios excluidos del cálculo de PRODUCTIVIDAD del equipo (extras, no-pickers regulares)
# Usar fragmentos del nombre, sin distinguir mayúsculas. Ej: "AIRALA" matchea "AIRALA CESAR".
# Agregar más separados por coma. "Falta definir" se excluye siempre.
$ExcludeFromProd = @("AIRALA","Falta definir","Pie de Maquina","MAQUINA","Muestra Simple","LEZCANO")

# Normalizacion de nombres de picker con variantes en el Excel (APELLIDO NOMBRE canónico)
$PickerNameMap = @{
    "FAVIO CAÑETE" = "CAÑETE FAVIO"
}

function IsExcludedFromProd($name){
    foreach($ex in $ExcludeFromProd){ if($name -like "*$ex*"){return $true} }
    return $false
}
$MES_NOM = @{1="Enero";2="Febrero";3="Marzo";4="Abril";5="Mayo";6="Junio";
             7="Julio";8="Agosto";9="Septiembre";10="Octubre";11="Noviembre";12="Diciembre"}
$ARR_UP   = [char]9650
$ARR_DOWN = [char]9660
$ARR_FLAT = [char]8594

$MaqEmailMap = @{
    "emilianodemarchiszecat@gmail.com" = "DEMARCHIS EMILIANO"
    "tomassandovalzecat@gmail.com"     = "SANDOVAL TOMAS"
}

$CtrlEmailMap = @{
    "danabarrioszecat@gmail.com"    = "BARRIOS DANA"
    "josepenidazecat@gmail.com"     = "PENIDA JOSE"
    "romeromelaniezecat@gmail.com"  = "ROMERO MELANIE"
    "tamaraarmoazecat@gmail.com"    = "ARMOA TAMARA"
    "tomaspittalazecat@gmail.com"   = "PITTALA TOMAS"
    "bejermanwmszecat@gmail.com"    = "BEJERMAN"
    "tomasbrunzecat@gmail.com"      = "BRUN TOMAS"
    "adrianromerozecat@gmail.com"   = "ADRIAN ROMERO"
    "braianarielocampo@gmail.com"   = "OCAMPO BRIAN"
    "logistica2@zecat.com"          = "AIRALA CESAR"
    "mcamerano@zecat.com"           = "PRODUCCION"
    "magwmszecat@gmail.com"         = "SISTEMA WMS"
}

# Emails a excluir de la seccion Control (no pertenecen al deposito AR)
$CtrlExclude = @("admdeposito.ch@zecat.cl")

function rgb($r,$g,$b){ [long]$r + [long]$g*256 + [long]$b*65536 }
$C = @{
    NavyHdr    = rgb 31  56  100;  BlueHdr   = rgb 46  117 182
    SkyBlue    = rgb 189 215 238;  PaleBlue  = rgb 222 235 247
    White      = rgb 255 255 255;  GreenBg   = rgb 0   176 80
    GreenFg    = rgb 255 255 255;  YellowBg  = rgb 255 192 0
    YellowFg   = rgb 50  30  0;    RedBg     = rgb 192 0   0
    RedFg      = rgb 255 255 255;  GreenTxt  = rgb 0   97  0
    YellowTxt  = rgb 156 87  0;    RedTxt    = rgb 192 0   6
    LightGray  = rgb 242 242 242;  MidGray   = rgb 217 217 217
    Gray       = rgb 166 166 166;  TextDark  = rgb 31  31  31
    TextWhite  = rgb 255 255 255;  TealHdr   = rgb 0   112 192
    TealLight  = rgb 189 224 237;  LezTab    = rgb 112 48  160
    LezHdr     = rgb 112 48  160;  LezLight  = rgb 237 225 252
    LezMid     = rgb 200 175 235;  OrangeHdr = rgb 197 90  17
    OrangeLight= rgb 252 228 214;  MermaHdr  = rgb 0   70  127
    MermaLight = rgb 189 215 238;  MermaMid  = rgb 141 180 226
    Merma2Hdr  = rgb 84  130 53;   Merma2Lt  = rgb 198 224 180
    StaffHdr   = rgb 68  114 196;  StaffLight= rgb 180 198 231
    PieHdr     = rgb 0   80  100;  PieLight  = rgb 155 210 225
    PieMid     = rgb 100 170 195
}
function W($cell,$val,$bg,$fg,$bold=$false,$sz=10,$ha=-4131,$fmt=$null){
    if($null -ne $val){ try{$cell.Value2=$val}catch{$cell.Value2="$val"} }
    if($null -ne $bg){ $cell.Interior.Color=[long]$bg }
    if($null -ne $fg){ $cell.Font.Color=[long]$fg }
    $cell.Font.Bold=[bool]$bold; $cell.Font.Size=[double]$sz
    $cell.HorizontalAlignment=[int]$ha; $cell.VerticalAlignment=-4108
    if($null -ne $fmt -and $fmt -ne ""){ $cell.NumberFormat="$fmt" }
}
function MR($ws,$r1,$c1,$r2,$c2){ $ws.Range($ws.Cells.Item($r1,$c1),$ws.Cells.Item($r2,$c2)).Merge()|Out-Null }
function RowH($ws,$r,$h){ $ws.Rows.Item($r).RowHeight=$h }
function ColW($ws,$c,$w){ $ws.Columns.Item($c).ColumnWidth=$w }
function PerfColors($ld){
    $ld=[double]$ld
    if($ld -ge $TARGET){ return ,@($C.GreenBg,$C.GreenFg) }
    if($ld -ge 70)     { return ,@($C.YellowBg,$C.YellowFg) }
    return ,@($C.RedBg,$C.RedFg)
}
function RecColors($n){
    $n=[int]$n
    if($n -eq 0){ return ,@($C.GreenBg,$C.GreenFg) }
    if($n -le 3){ return ,@($C.YellowBg,$C.YellowFg) }
    return ,@($C.RedBg,$C.RedFg)
}
function RateFg($r){ $r=[double]$r; if($r -le 1){return $C.GreenTxt}; if($r -le 3){return $C.YellowTxt}; return $C.RedTxt }
function KPI-Card($ws,$r,$c1,$c2,$label,$val,$bgVal,$fgVal,$fmt=""){
    RowH $ws $r 20; MR $ws $r $c1 $r $c2
    W ($ws.Cells.Item($r,$c1)) $label $C.PaleBlue $C.BlueHdr $true 9 -4108
    $r++; RowH $ws $r 32; MR $ws $r $c1 $r $c2
    $cell=$ws.Cells.Item($r,$c1)
    if($fmt){W $cell $val $bgVal $fgVal $true 22 -4108 $fmt}else{W $cell "$val" $bgVal $fgVal $true 22 -4108}
    return $r+1
}
function LezKPI($ws,$r,$c1,$c2,$label,$val,$bg,$fg,$fmt=""){
    RowH $ws $r 18; MR $ws $r $c1 $r $c2
    W ($ws.Cells.Item($r,$c1)) $label $C.LezLight $C.LezHdr $true 9 -4108
    $r++; RowH $ws $r 30; MR $ws $r $c1 $r $c2
    $cell=$ws.Cells.Item($r,$c1)
    if($fmt){W $cell $val $bg $fg $true 20 -4108 $fmt}else{W $cell "$val" $bg $fg $true 20 -4108}
    return $r+1
}
function SecHdr($ws,$r,$cols,$txt){
    MR $ws $r 1 $r $cols
    W ($ws.Cells.Item($r,1)) "  $txt" $C.BlueHdr $C.TextWhite $true 12 -4131
    RowH $ws $r 22
}
function TblHdr($ws,$r,$hdrs){
    RowH $ws $r 20
    for($c=0;$c -lt $hdrs.Count;$c++){ W ($ws.Cells.Item($r,$c+1)) $hdrs[$c] $C.MidGray $C.NavyHdr $true 9 -4108 }
}
function AF($ws,$r,$c1,$c2){
    try{ $ws.Range($ws.Cells.Item($r,$c1),$ws.Cells.Item($r,$c2)).AutoFilter()|Out-Null }catch{}
}

Write-Host "[$($NOW.ToString('HH:mm:ss'))] Leyendo datos desde Picking (fuente primaria)..."
$xl = New-Object -ComObject Excel.Application
$xl.Visible=$false; $xl.DisplayAlerts=$false

try {
    $srcWb = $xl.Workbooks.Open($Source,$false,$true)

    # ===========================================================
    # LEER PICKING (FUENTE PRIMARIA v5)
    # Col1=NUMERO_OLEADA Col5=TOTAL_LINEAS Col6=CANTIDAD_UNIDADES
    # Col7=PRIMERA_CREACION(texto) Col9=GRUPO_ORDEN Col10=Pickeador
    # Grupos:
    #   PEDIDO MERMA    -> MERMA 1/2 por horario, hoja MERMA
    #   PIE DE MAQUINA  -> MERMA 1/2 por horario, hoja PIE_MAQUINA
    #   MUESTRA SIMPLE  -> LEZCANO AGUSTIN, hoja LEZCANO
    #   SIN LOGO/CON LOGO -> pickeador col10, picking regular (target 84)
    # ===========================================================
    $wsPk  = $srcWb.Sheets.Item("Picking")
    $pkArr = $wsPk.UsedRange.Value2
    $pkRows= $pkArr.GetUpperBound(0)

    $pkMes      = @{}   # picking regular por operario+mes
    $pkDay      = @{}   # lineas por operario+fecha (todos los grupos)
    $staffByMes = @{}   # solo SIN LOGO / CON LOGO (con y sin pickeador)
    $lezcanoMes = @{}
    $mermaMes   = @{}
    $pieMes     = @{}
    $allMon     = @{}
    $allResp    = @{}   # pickeadores regulares unicos
    $dayGroupData = @{}  # date → @{CL;SL;CLops;SLops} para filtros Web
    $allOperators = @{}
    $opsByDay   = @{}   # operarios por dia (todos los grupos reales)

    for($r=2; $r -le $pkRows; $r++){
        $L     = [double]($pkArr[$r,5] -as [double])
        $U     = [double]($pkArr[$r,6] -as [double])
        $grp   = "$($pkArr[$r,9])".Trim().ToUpper()
        $pkRaw = $pkArr[$r,10]
        $pk    = if(($pkRaw -is [double] -or $pkRaw -is [int] -or $pkRaw -is [long]) -and $pkRaw -lt 0){"Falta definir pickeador"}else{"$pkRaw".Trim()}  # #N/A / errores int → etiqueta
        $dtStr = "$($pkArr[$r,7])"
        $dt    = $null
        if($dtStr){ try{ $dt=[datetime]::Parse($dtStr) }catch{} }
        if(-not $dt){ continue }
        $ym  = "$($dt.Year)-$('{0:00}' -f $dt.Month)"
        $ymd = $dt.ToString("yyyy-MM-dd")
        $dow = [int]$dt.DayOfWeek
        $allMon[$ym]=$true

        # Clasificacion de grupo
        $picker=$null; $isMerma=$false; $isPie=$false; $isLezcano=$false; $isRegular=$false

        if($grp -like "*MERMA*"){
            $picker  = if($dt.Hour -ge 6 -and $dt.Hour -lt 15){ "OPERARIO MERMA 1" }else{ "OPERARIO MERMA 2" }
            $isMerma = $true

        } elseif($grp -like "*PIE*"){
            $picker = if($dt.Hour -ge 6 -and $dt.Hour -lt 15){ "OPERARIO MERMA 1" }else{ "OPERARIO MERMA 2" }
            $isPie  = $true

        } elseif($grp -like "*MUESTRA*"){
            $picker    = "LEZCANO AGUSTIN"
            $isLezcano = $true

        } else {
            # SIN LOGO / CON LOGO: picking regular
            # Staffing cuenta TODAS las lineas regulares (incluye Falta Pickeador = demanda sin cubrir)
            if(-not $staffByMes[$ym]){ $staffByMes[$ym]=@{TotalL=0;WorkDays=@{}} }
            $staffByMes[$ym].TotalL += $L
            if($dow -ge 1 -and $dow -le 5){ $staffByMes[$ym].WorkDays[$ymd]=$true }
            if(-not $pk -or $pk -eq "Falta Pickeador"){ continue }
            $picker    = if($PickerNameMap[$pk]){$PickerNameMap[$pk]}else{$pk}
            $isRegular = $true
            $allResp[$picker] = $true
        }

        $allOperators[$picker] = $true
        if(-not $opsByDay[$ymd]){ $opsByDay[$ymd]=@{} }
        $opsByDay[$ymd][$picker]=$true
        $dk="$picker|$ymd"
        if(-not $pkDay[$dk]){$pkDay[$dk]=0}; $pkDay[$dk]+=$L

        if($isRegular){
            $key = "$picker|$ym"
            if(-not $pkMes[$key]){ $pkMes[$key]=@{L=0;U=0;Olas=0;Days=@{};CL=0;SL=0} }
            $pkMes[$key].L+=$L; $pkMes[$key].U+=$U; $pkMes[$key].Olas++
            $pkMes[$key].Days[$ymd]=$true
            if($grp -like "*CON LOGO*"){ $pkMes[$key].CL++ } else { $pkMes[$key].SL++ }
            # Acumular por grupo por dia (para filtros Web)
            if(-not $dayGroupData[$ymd]){ $dayGroupData[$ymd]=@{CL=0;SL=0;CLops=@{};SLops=@{}} }
            if($grp -like "*CON LOGO*"){
                $dayGroupData[$ymd].CL += $L
                $dayGroupData[$ymd].CLops[$picker] = $true
            } else {
                $dayGroupData[$ymd].SL += $L
                $dayGroupData[$ymd].SLops[$picker] = $true
            }
        }
        if($isLezcano){
            if(-not $lezcanoMes[$ym]){ $lezcanoMes[$ym]=@{L=0;U=0;Olas=0;Days=@{}} }
            $lezcanoMes[$ym].L+=$L; $lezcanoMes[$ym].U+=$U; $lezcanoMes[$ym].Olas++
            $lezcanoMes[$ym].Days[$ymd]=$true
        }
        if($isMerma){
            if(-not $mermaMes[$ym]){ $mermaMes[$ym]=@{M1L=0;M1U=0;M1Days=@{};M2L=0;M2U=0;M2Days=@{}} }
            if($picker -eq "OPERARIO MERMA 1"){
                $mermaMes[$ym].M1L+=$L; $mermaMes[$ym].M1U+=$U; $mermaMes[$ym].M1Days[$ymd]=$true
            } else {
                $mermaMes[$ym].M2L+=$L; $mermaMes[$ym].M2U+=$U; $mermaMes[$ym].M2Days[$ymd]=$true
            }
        }
        if($isPie){
            if(-not $pieMes[$ym]){ $pieMes[$ym]=@{M1L=0;M1U=0;M1Days=@{};M2L=0;M2U=0;M2Days=@{}} }
            if($picker -eq "OPERARIO MERMA 1"){
                $pieMes[$ym].M1L+=$L; $pieMes[$ym].M1U+=$U; $pieMes[$ym].M1Days[$ymd]=$true
            } else {
                $pieMes[$ym].M2L+=$L; $pieMes[$ym].M2U+=$U; $pieMes[$ym].M2Days[$ymd]=$true
            }
        }
    }

    # ===========================================================
    # LEER RECLAMOS
    # Cols: 12=categoria 13=mes 14=ano 15=Pickeador
    # ===========================================================
    $wsRec = $srcWb.Sheets.Item("Reclamos")
    $rArr  = $wsRec.UsedRange.Value2
    $rRows = $rArr.GetUpperBound(0)
    $recOpMes=@{}; $recByCat=@{}; $recByFam=@{}; $recByMes=@{}; $recByPik=@{}; $lezcanoRecMes=@{}
    $recDataByMonth=@{}   # key=YM → {Cnt,ByOp={op:cnt},ByCat={cat:cnt}}

    for($r=2; $r -le $rRows; $r++){
        # Nueva estructura Reclamos: col4=fecha reclamo(OA), col9=qty reclam., col12=familia, col13=categoría, col15=Pickeador
        $fechaRaw=$rArr[$r,4]
        if(-not $fechaRaw){ continue }
        $dtRec=$null; try{$dtRec=[datetime]::FromOADate([double]$fechaRaw)}catch{continue}
        $ym="$($dtRec.Year)-$('{0:00}' -f $dtRec.Month)"
        $rpkRaw=$rArr[$r,15]
        $rpk=if(-not $rpkRaw -or ($rpkRaw -is [double] -and $rpkRaw -lt 0)){"Falta definir pickeador"}else{"$rpkRaw".Trim()}  # #N/A → etiqueta
        $qty=[double]($rArr[$r,9] -as [double])
        $cat=$rArr[$r,13]; $fam=$rArr[$r,12]
        $isLez=($rpk -like "*$LezFilter*")
        if(-not $recByMes[$ym]){$recByMes[$ym]=@{Cnt=0;Qty=0}}
        $recByMes[$ym].Cnt++; $recByMes[$ym].Qty+=$qty
        if($cat){if(-not $recByCat[$cat]){$recByCat[$cat]=@{Cnt=0;Qty=0}};$recByCat[$cat].Cnt++;$recByCat[$cat].Qty+=$qty}
        if($fam){if(-not $recByFam[$fam]){$recByFam[$fam]=0};$recByFam[$fam]++}
        # Acumular en recDataByMonth (incluye Lezcano)
        if(-not $recDataByMonth[$ym]){$recDataByMonth[$ym]=@{Cnt=0;ByOp=@{};ByCat=@{};ByOpCat=@{}}}
        $recDataByMonth[$ym].Cnt++
        $opKey=if($isLez){"LEZCANO AGUSTIN"}else{$rpk}
        if(-not $recDataByMonth[$ym].ByOp[$opKey]){$recDataByMonth[$ym].ByOp[$opKey]=0}
        $recDataByMonth[$ym].ByOp[$opKey]++
        if($cat){
            if(-not $recDataByMonth[$ym].ByCat[$cat]){$recDataByMonth[$ym].ByCat[$cat]=0}
            $recDataByMonth[$ym].ByCat[$cat]++
            if(-not $recDataByMonth[$ym].ByOpCat[$opKey]){$recDataByMonth[$ym].ByOpCat[$opKey]=@{}}
            if(-not $recDataByMonth[$ym].ByOpCat[$opKey][$cat]){$recDataByMonth[$ym].ByOpCat[$opKey][$cat]=0}
            $recDataByMonth[$ym].ByOpCat[$opKey][$cat]++
        }
        if($isLez){
            if(-not $lezcanoRecMes[$ym]){$lezcanoRecMes[$ym]=@{Cnt=0;Qty=0}}
            $lezcanoRecMes[$ym].Cnt++; $lezcanoRecMes[$ym].Qty+=$qty
        } else {
            $key="$rpk|$ym"
            if(-not $recOpMes[$key]){$recOpMes[$key]=@{Cnt=0;Qty=0}}
            $recOpMes[$key].Cnt++; $recOpMes[$key].Qty+=$qty
            if(-not $recByPik[$rpk]){$recByPik[$rpk]=@{Cnt=0;Qty=0;Cats=@{}}}
            $recByPik[$rpk].Cnt++; $recByPik[$rpk].Qty+=$qty
            if($cat){if(-not $recByPik[$rpk].Cats[$cat]){$recByPik[$rpk].Cats[$cat]=0};$recByPik[$rpk].Cats[$cat]++}
        }
    }

    # ===========================================================
    # LEER MAQUINISTA
    # ===========================================================
    $wsMaq=$srcWb.Sheets.Item("Maquinista"); $mArr=$wsMaq.UsedRange.Value2; $mRows=$mArr.GetUpperBound(0)
    $maqMes=@{}; $maqDay=@{}; $maqNames=@{}
    $ctrlHasData=($srcWb.Sheets.Item("Control").UsedRange.Rows.Count -gt 1)
    $ctrlMes=@{}

    for($r=2; $r -le $mRows; $r++){
        $email="$($mArr[$r,2])".Trim(); $dv=$mArr[$r,1]
        if(-not $email -or -not $dv){continue}
        $dt=$null; try{$dt=[datetime]::FromOADate([double]$dv)}catch{continue}
        $nombre=if($MaqEmailMap[$email]){$MaqEmailMap[$email]}else{$email}
        $ym="$($dt.Year)-$('{0:00}' -f $dt.Month)"; $ymd=$dt.ToString("yyyy-MM-dd")
        $pa02=[double]($mArr[$r,3] -as [double]); $rec=[double]($mArr[$r,5] -as [double])
        $rl01=[double]($mArr[$r,7] -as [double]); $pallet=[double]($mArr[$r,9] -as [double])
        $flow=[double]($mArr[$r,11] -as [double]); $totMov=[double]($mArr[$r,13] -as [double])
        $totUnd=[double]($mArr[$r,14] -as [double])
        $key="$nombre|$ym"
        if(-not $maqMes[$key]){$maqMes[$key]=@{Mov=0;Und=0;Pa02=0;Rec=0;Rl01=0;Pallet=0;Flow=0;Days=@{}}}
        $maqMes[$key].Mov+=$totMov; $maqMes[$key].Und+=$totUnd
        $maqMes[$key].Pa02+=$pa02; $maqMes[$key].Rec+=$rec
        $maqMes[$key].Rl01+=$rl01; $maqMes[$key].Pallet+=$pallet
        $maqMes[$key].Flow+=$flow; $maqMes[$key].Days[$ymd]=$true
        $dk="$nombre|$ymd"; if(-not $maqDay[$dk]){$maqDay[$dk]=0}; $maqDay[$dk]+=$totMov
        $maqNames[$nombre]=$true
    }
    if($ctrlHasData){
        $wsCtrl=$srcWb.Sheets.Item("Control"); $cArr=$wsCtrl.UsedRange.Value2; $cRows=$cArr.GetUpperBound(0)
        for($r=2; $r -le $cRows; $r++){
            $dv=$cArr[$r,1]; $email="$($cArr[$r,2])".Trim()
            $ord=[double]($cArr[$r,3] -as [double]); $und=[double]($cArr[$r,4] -as [double])
            if(-not $dv -or -not $email){continue}
            if($CtrlExclude -contains $email){continue}
            $dt=$null; try{$dt=[datetime]::FromOADate([double]$dv)}catch{continue}
            $nombre=if($CtrlEmailMap[$email]){$CtrlEmailMap[$email]}else{$email}
            $ym="$($dt.Year)-$('{0:00}' -f $dt.Month)"; $ymd=$dt.ToString("yyyy-MM-dd")
            $key="$nombre|$ym"
            if(-not $ctrlMes[$key]){$ctrlMes[$key]=@{Ord=0;Und=0;Days=@{}}}
            $ctrlMes[$key].Ord+=$ord; $ctrlMes[$key].Und+=$und; $ctrlMes[$key].Days[$ymd]=$true
        }
    }

    $srcWb.Close($false)
    Write-Host "[$($NOW.ToString('HH:mm:ss'))] Picking: $pkRows filas | Reclamos: $rRows filas | Maquinista: $mRows filas"

    # ===========================================================
    # AGREGACIONES
    # ===========================================================
    $sortedMon  = $allMon.Keys  | Sort-Object
    $sortedResp = @($allResp.Keys | Where-Object{-not (IsExcludedFromProd $_)} | Sort-Object)
    $latestYM   = $sortedMon[-1]
    $lyParts    = $latestYM.Split("-")
    $latestNom  = "$($MES_NOM[[int]$lyParts[1]]) $($lyParts[0])"

    # sumRows: pickeadores regulares por mes
    $sumRows=[System.Collections.Generic.List[PSObject]]::new()
    foreach($ym in $sortedMon){
        $ymp=$ym.Split("-"); $ymN="$($MES_NOM[[int]$ymp[1]]) $($ymp[0])"
        foreach($resp in $sortedResp){
            $k="$resp|$ym"; if(-not $pkMes[$k]){continue}
            $pd=$pkMes[$k]; $days=$pd.Days.Count
            $ld=if($days){[Math]::Round($pd.L/$days,1)}else{0}
            $ul=if($pd.L-gt 0){[Math]::Round($pd.U/$pd.L,2)}else{0}
            $comp=[Math]::Round($ld/$TARGET*100,1)
            $tgt=$days*$TARGET; $diff=[int]$pd.L-$tgt
            $pctCL=if($pd.Olas){[Math]::Round($pd.CL/$pd.Olas*100,0)}else{0}
            $rec=if($recOpMes[$k]){$recOpMes[$k]}else{@{Cnt=0;Qty=0}}
            $rate=if($pd.L){[Math]::Round($rec.Cnt/($pd.L/1000),2)}else{0}
            $sumRows.Add([PSCustomObject]@{
                YM=$ym;MesNom=$ymN;Resp=$resp;Dias=$days;Olas=$pd.Olas;Lineas=[int]$pd.L
                LineasDia=$ld;Target=$tgt;Diff=$diff;Cumplim=$comp
                Unidades=[int]$pd.U;ULinea=$ul;ConLogo=$pd.CL;SinLogo=$pd.SL;PctConLogo=$pctCL
                RecCnt=$rec.Cnt;RecQty=$rec.Qty;RecRate=$rate
            })
        }
    }

    $latestRows=@($sumRows|Where-Object{$_.YM -eq $latestYM}|Sort-Object LineasDia -Descending)
    $totL=0;$totU=0;$totRC=0
    foreach($op in $latestRows){$totL+=$op.Lineas;$totU+=$op.Unidades;$totRC+=$op.RecCnt}
    # Excluir extras del cálculo de productividad del equipo
    $prodRows=@($latestRows|Where-Object{-not (IsExcludedFromProd $_.Resp)})
    $totOps=$prodRows.Count
    $avgLD=if($totOps){[Math]::Round(($prodRows|Measure-Object LineasDia -Average).Average,1)}else{0}
    $cumAv=[Math]::Round($avgLD/$TARGET*100,1)
    $rateG=if($totL){[Math]::Round($totRC/($totL/1000),2)}else{0}
    $totOlas=($latestRows|Measure-Object Olas -Sum).Sum

    # Tendencia
    $trendByResp=@{}
    if($sortedMon.Count -ge 2){
        $mC=$sortedMon[-1]; $mP=$sortedMon[-2]
        foreach($resp in $allResp.Keys){
            $pdC=$pkMes["$resp|$mC"]; $pdP=$pkMes["$resp|$mP"]
            $ldC=if($pdC -and $pdC.Days.Count){$pdC.L/$pdC.Days.Count}else{$null}
            $ldP=if($pdP -and $pdP.Days.Count){$pdP.L/$pdP.Days.Count}else{$null}
            if($null -ne $ldC -and $null -ne $ldP){
                $dv=[Math]::Round($ldC-$ldP,1)
                if($dv-gt 3){$trendByResp[$resp]=@{Txt="$ARR_UP +$dv";Fg=$C.GreenTxt}}
                elseif($dv-lt -3){$trendByResp[$resp]=@{Txt="$ARR_DOWN $dv";Fg=$C.RedTxt}}
                else{$trendByResp[$resp]=@{Txt="$ARR_FLAT $dv";Fg=$C.Gray}}
            }else{$trendByResp[$resp]=@{Txt="-";Fg=$C.Gray}}
        }
    }

    # Ranking por mes
    $rankByPkMon=@{}
    foreach($ym in $sortedMon){
        $mOps=@($sumRows|Where-Object{$_.YM -eq $ym}|Sort-Object LineasDia -Descending)
        for($i=0;$i -lt $mOps.Count;$i++){ $rankByPkMon["$($mOps[$i].Resp)|$ym"]=$i+1 }
    }

    # Promedio equipo por mes (solo picking regular)
    $teamAvgByMonth=@{}
    foreach($ym in $sortedMon){
        $mR=@($sumRows|Where-Object{$_.YM -eq $ym})
        if($mR.Count){$teamAvgByMonth[$ym]=[Math]::Round(($mR|Measure-Object LineasDia -Average).Average,1)}
    }

    # staffRows - solo picking regular
    $staffRows=[System.Collections.Generic.List[PSObject]]::new()
    foreach($ym in $sortedMon){
        $sd=$staffByMes[$ym]; if(-not $sd){continue}
        $dh=$sd.WorkDays.Count; $tL=$sd.TotalL
        $pn=if($dh){[Math]::Round($tL/$dh/$TARGET,1)}else{0}
        $pa=0
        if($dh){
            $sumOps=0
            foreach($day in $sd.WorkDays.Keys){
                if($opsByDay[$day]){ $sumOps += $opsByDay[$day].Keys.Count }
            }
            $pa=[Math]::Round($sumOps/$dh,1)
        }
        $ymp=$ym.Split("-")
        $staffRows.Add([PSCustomObject]@{
            YM=$ym;MesNom="$($MES_NOM[[int]$ymp[1]]) $($ymp[0])"
            TotalLineas=[int]$tL;DiasHabiles=$dh;PersonasNecesarias=$pn;PersonasActuales=$pa
        })
    }
    $latestStaff=($staffRows|Where-Object{$_.YM -eq $latestYM}|Select-Object -First 1)

    # Ultimos 7 dias de picking (solo equipo regular SIN/CON LOGO)
    $reg7Dates=@($pkDay.Keys|Where-Object{$allResp[$_.Split("|")[0]]}|
        ForEach-Object{$_.Split("|")[1]}|Sort-Object -Unique|Select-Object -Last 7)
    $day7Rows=[System.Collections.Generic.List[PSObject]]::new()
    $prevD7L=$null
    foreach($d7ymd in $reg7Dates){
        $d7dt=[datetime]::Parse($d7ymd)
        $d7L=0;$d7Ops=@{}
        foreach($resp in $sortedResp){
            $d7k="$resp|$d7ymd"
            if($pkDay[$d7k]){$d7L+=$pkDay[$d7k];$d7Ops[$resp]=$true}
        }
        $d7n=$d7Ops.Count
        $d7LD=if($d7n){[Math]::Round($d7L/$d7n,1)}else{0}
        $d7T=$d7n*$TARGET
        $d7C=if($d7T){[Math]::Round($d7L/$d7T*100,1)}else{0}
        $d7Dlt=if($null -ne $prevD7L){[int]($d7L-$prevD7L)}else{$null}
        $d7gd=$dayGroupData[$d7ymd]
        $d7CL=if($d7gd){[int]$d7gd.CL}else{0}
        $d7SL=if($d7gd){[int]$d7gd.SL}else{0}
        $d7CLn=if($d7gd){$d7gd.CLops.Count}else{0}
        $d7SLn=if($d7gd){$d7gd.SLops.Count}else{0}
        $day7Rows.Add([PSCustomObject]@{
            Fecha=$d7dt.ToString("dd/MM")
            DiaSem=@('Do','Lu','Ma','Mi','Ju','Vi','Sa')[[int]$d7dt.DayOfWeek]
            Ops=$d7n;Lines=[int]$d7L;LinesPerOp=$d7LD;Target=$d7T;Cumpl=$d7C;Delta=$d7Dlt
            CLLines=$d7CL;CLOps=$d7CLn;SLLines=$d7SL;SLOps=$d7SLn
        })
        $prevD7L=$d7L
    }

    # mermaRows
    $mermaRows=[System.Collections.Generic.List[PSObject]]::new()
    foreach($ym in ($mermaMes.Keys|Sort-Object)){
        $md=$mermaMes[$ym]
        $d1=$md.M1Days.Count; $d2=$md.M2Days.Count
        $ld1=if($d1){[Math]::Round($md.M1L/$d1,1)}else{0}
        $ld2=if($d2){[Math]::Round($md.M2L/$d2,1)}else{0}
        $ymp=$ym.Split("-")
        $mermaRows.Add([PSCustomObject]@{
            YM=$ym;MesNom="$($MES_NOM[[int]$ymp[1]]) $($ymp[0])"
            M1Dias=$d1;M1Lineas=[int]$md.M1L;M1LD=$ld1;M1Unidades=[int]$md.M1U
            M2Dias=$d2;M2Lineas=[int]$md.M2L;M2LD=$ld2;M2Unidades=[int]$md.M2U
            TotalLineas=[int]($md.M1L+$md.M2L)
        })
    }

    # pieRows
    $pieRows=[System.Collections.Generic.List[PSObject]]::new()
    foreach($ym in ($pieMes.Keys|Sort-Object)){
        $pd=$pieMes[$ym]
        $d1=$pd.M1Days.Count; $d2=$pd.M2Days.Count
        $ld1=if($d1){[Math]::Round($pd.M1L/$d1,1)}else{0}
        $ld2=if($d2){[Math]::Round($pd.M2L/$d2,1)}else{0}
        $ymp=$ym.Split("-")
        $pieRows.Add([PSCustomObject]@{
            YM=$ym;MesNom="$($MES_NOM[[int]$ymp[1]]) $($ymp[0])"
            M1Dias=$d1;M1Lineas=[int]$pd.M1L;M1LD=$ld1;M1Unidades=[int]$pd.M1U
            M2Dias=$d2;M2Lineas=[int]$pd.M2L;M2LD=$ld2;M2Unidades=[int]$pd.M2U
            TotalLineas=[int]($pd.M1L+$pd.M2L)
        })
    }

    # lezRows
    $lezSortedMon=$lezcanoMes.Keys|Sort-Object
    $lezRows=[System.Collections.Generic.List[PSObject]]::new()
    $prevLezLD=$null
    foreach($ym in $lezSortedMon){
        $pd=$lezcanoMes[$ym]; $days=$pd.Days.Count
        $ld=if($days){[Math]::Round($pd.L/$days,1)}else{0}
        $ul=if($pd.L-gt 0){[Math]::Round($pd.U/$pd.L,2)}else{0}
        $rec=if($lezcanoRecMes[$ym]){$lezcanoRecMes[$ym]}else{@{Cnt=0;Qty=0}}
        $ymp=$ym.Split("-"); $tmAvg=if($teamAvgByMonth[$ym]){$teamAvgByMonth[$ym]}else{0}
        $vsTeam=[Math]::Round($ld-$tmAvg,1)
        $trendStr=if($null -ne $prevLezLD){
            $dv2=[Math]::Round($ld-$prevLezLD,1)
            if($dv2-gt 3){"$ARR_UP +$dv2"}elseif($dv2-lt -3){"$ARR_DOWN $dv2"}else{"$ARR_FLAT $dv2"}
        }else{"-"}
        $lezRows.Add([PSCustomObject]@{
            YM=$ym;MesNom="$($MES_NOM[[int]$ymp[1]]) $($ymp[0])";Dias=$days
            Olas=$pd.Olas;Lineas=[int]$pd.L;LineasDia=$ld
            Unidades=[int]$pd.U;ULinea=$ul
            RecCnt=$rec.Cnt;TeamAvg=$tmAvg;VsTeam=$vsTeam;TrendStr=$trendStr
        })
        $prevLezLD=$ld
    }

    # ===========================================================
    # CREAR WORKBOOK (11 HOJAS)
    # RESUMEN / STAFFING / RANKING / EVOLUCION_MENSUAL /
    # PRODUCCION_DIARIA / LEZCANO / PIE_MAQUINA / MERMA /
    # RECLAMOS / MAQUINISTA / CONTROL
    # ===========================================================
    if(Test-Path $Output){try{Remove-Item $Output -Force -ErrorAction Stop}catch{}}
    $xl.SheetsInNewWorkbook=11
    $dashWb=$xl.Workbooks.Add()
    $xl.SheetsInNewWorkbook=3
    $ws1  =$dashWb.Sheets.Item(1);  $ws1.Name  ="RESUMEN"
    $ws2  =$dashWb.Sheets.Item(2);  $ws2.Name  ="STAFFING"
    $ws3  =$dashWb.Sheets.Item(3);  $ws3.Name  ="RANKING"
    $ws4  =$dashWb.Sheets.Item(4);  $ws4.Name  ="EVOLUCION_MENSUAL"
    $ws5  =$dashWb.Sheets.Item(5);  $ws5.Name  ="PRODUCCION_DIARIA"
    $ws6  =$dashWb.Sheets.Item(6);  $ws6.Name  ="LEZCANO"
    $ws7  =$dashWb.Sheets.Item(7);  $ws7.Name  ="PIE_MAQUINA"
    $ws8  =$dashWb.Sheets.Item(8);  $ws8.Name  ="MERMA"
    $ws9  =$dashWb.Sheets.Item(9);  $ws9.Name  ="RECLAMOS"
    $ws10 =$dashWb.Sheets.Item(10); $ws10.Name ="MAQUINISTA"
    $ws11 =$dashWb.Sheets.Item(11); $ws11.Name ="CONTROL"

    $ws1.Tab.Color  =rgb 31 56 100;  $ws2.Tab.Color  =rgb 68 114 196
    $ws3.Tab.Color  =rgb 46 117 182; $ws4.Tab.Color  =rgb 0  128 0
    $ws5.Tab.Color  =rgb 17 34  71;  $ws6.Tab.Color  =rgb 112 48 160
    $ws7.Tab.Color  =rgb 0  80  100; $ws8.Tab.Color  =rgb 0  70  127
    $ws9.Tab.Color  =rgb 192 0   0;  $ws10.Tab.Color =rgb 197 90  17
    $ws11.Tab.Color =rgb 128 128 128

    # ===========================================================
    # HOJA 1: RESUMEN
    # ===========================================================
    ColW $ws1 1 2;ColW $ws1 2 24;ColW $ws1 3 16;ColW $ws1 4 16
    ColW $ws1 5 16;ColW $ws1 6 16;ColW $ws1 7 16
    RowH $ws1 1 6;RowH $ws1 2 44;RowH $ws1 3 24;RowH $ws1 4 6
    MR $ws1 2 2 2 7
    W ($ws1.Cells.Item(2,2)) "  DASHBOARD DE PRODUCTIVIDAD - PICKING" $C.NavyHdr $C.TextWhite $true 20 -4131
    MR $ws1 3 2 3 5
    W ($ws1.Cells.Item(3,2)) "  Articulos Promocionales SA | Picking regular: SIN/CON LOGO" $C.BlueHdr $C.TextWhite $true 12 -4131
    MR $ws1 3 6 3 7
    W ($ws1.Cells.Item(3,6)) "Actualizado: $($NOW.ToString('dd/MM/yyyy HH:mm'))" $C.BlueHdr $C.SkyBlue $false 9 -4152
    $row=6
    MR $ws1 $row 2 $row 7
    W ($ws1.Cells.Item($row,2)) "  MES EN CURSO: $($latestNom.ToUpper())" $C.SkyBlue $C.NavyHdr $true 13 -4131
    $row++; RowH $ws1 $row 4; $row++

    $cumBg=if($cumAv-ge 100){$C.GreenBg}elseif($cumAv-ge 83){$C.YellowBg}else{$C.RedBg}
    $cumFg=if($cumAv-ge 100){$C.GreenFg}elseif($cumAv-ge 83){$C.YellowFg}else{$C.RedFg}
    $ldPc =PerfColors $avgLD
    $recBg=if($totRC-eq 0){$C.GreenBg}elseif($totRC-le 5){$C.YellowBg}else{$C.RedBg}
    $recFg=if($totRC-eq 0){$C.GreenFg}elseif($totRC-le 5){$C.YellowFg}else{$C.RedFg}
    $trBg =if($rateG-le 1){$C.GreenBg}elseif($rateG-le 3){$C.YellowBg}else{$C.RedBg}
    $trFg =if($rateG-le 1){$C.GreenFg}elseif($rateG-le 3){$C.YellowFg}else{$C.RedFg}

    $r1=KPI-Card $ws1 $row 2 3 "OPERARIOS ACTIVOS"  $totOps        $C.White  $C.NavyHdr
    $r2=KPI-Card $ws1 $row 4 5 "TOTAL LINEAS MES"   $totL          $C.White  $C.NavyHdr "#,##0"
    $r3=KPI-Card $ws1 $row 6 7 "CUMPLIMIENTO PROM"  "$cumAv%"      $cumBg    $cumFg
    $row=[Math]::Max([Math]::Max($r1,$r2),$r3); RowH $ws1 $row 4; $row++
    $r1=KPI-Card $ws1 $row 2 3 "LINEAS/DIA PROM"    "$avgLD  (target $TARGET)" $ldPc[0] $ldPc[1]
    $r2=KPI-Card $ws1 $row 4 5 "RECLAMOS DEL MES"   $totRC         $recBg    $recFg
    $r3=KPI-Card $ws1 $row 6 7 "TASA REC/1000 LIN"  $rateG         $trBg     $trFg "0.00"
    $row=[Math]::Max([Math]::Max($r1,$r2),$r3); RowH $ws1 $row 4; $row++

    # KPI staffing
    $pnBg=if($latestStaff -and $latestStaff.PersonasActuales -ge $latestStaff.PersonasNecesarias){$C.GreenBg}else{$C.YellowBg}
    $pnFg=if($latestStaff -and $latestStaff.PersonasActuales -ge $latestStaff.PersonasNecesarias){$C.GreenFg}else{$C.YellowFg}
    $pnVal=if($latestStaff){"$($latestStaff.PersonasNecesarias)"}else{"-"}
    $paVal=if($latestStaff){"$($latestStaff.PersonasActuales)"}else{"-"}
    $r1=KPI-Card $ws1 $row 2 3 "PERSONAS NECESARIAS (PICKING)" $pnVal $pnBg $pnFg
    $r2=KPI-Card $ws1 $row 4 5 "PERSONAS ACTIVAS (PROM/DIA)"   $paVal $C.White $C.NavyHdr
    $r3=KPI-Card $ws1 $row 6 7 "OLAS PROCESADAS"               $totOlas $C.White $C.NavyHdr "#,##0"
    $row=[Math]::Max([Math]::Max($r1,$r2),$r3); RowH $ws1 $row 6; $row++

    # ── GRAFICO EVOLUCION DIARIA (flotante, arriba derecha) ──
    $cdDates=@($pkDay.Keys|Where-Object{$allResp[$_.Split("|")[0]]}|
        ForEach-Object{$_.Split("|")[1]}|Sort-Object -Unique|Select-Object -Last 30)
    $cdR=2;$cdC=16
    foreach($cdd in $cdDates){
        $cddT=0
        foreach($rsp in $sortedResp){$cddk="$rsp|$cdd";if($pkDay[$cddk]){$cddT+=$pkDay[$cddk]}}
        $ws1.Cells.Item($cdR,$cdC).Value2=[datetime]::Parse($cdd).ToString("dd/MM")
        $ws1.Cells.Item($cdR,$cdC+1).Value2=[int]$cddT
        $cdR++
    }
    try{
        $valRange=$ws1.Range($ws1.Cells.Item(2,$cdC+1),$ws1.Cells.Item($cdR-1,$cdC+1))
        $lblRange=$ws1.Range($ws1.Cells.Item(2,$cdC),  $ws1.Cells.Item($cdR-1,$cdC))
        $xl.ScreenUpdating=$false
        $xl.Visible=$true
        $dashWb.Activate(); $ws1.Activate()
        $cObj=$ws1.ChartObjects().Add([double]650,[double]44,[double]510,[double]200)
        $xl.Visible=$false
        $cht=$cObj.Chart
        $cht.ChartType=4  # xlLine
        $cht.SetSourceData($valRange)
        try{
            $cSer=$cht.SeriesCollection().Item(1)
            try{$cSer.XValues=$lblRange}catch{}
            $cSer.Name="Total lineas equipo"
            try{$cSer.Border.Color=[long](rgb 46 117 182);$cSer.Border.Weight=2}catch{}
            try{$cSer.MarkerStyle=-4142}catch{}
        }catch{Write-Host "  [WARN] Series: $_"}
        $cht.HasTitle=$true
        $cht.ChartTitle.Text="Evolucion diaria del equipo (ultimos 30 dias)"
        $cht.ChartTitle.Font.Size=10;$cht.ChartTitle.Font.Bold=$true
        $cht.HasLegend=$false
        try{$cht.PlotArea.Interior.Color=[long]$C.LightGray}catch{}
        try{$cht.ChartArea.Interior.Color=[long]$C.White}catch{}
        try{$cht.ChartArea.Border.LineStyle=-4142}catch{}
        try{$cht.Axes(1).TickLabels.Font.Size=7;$cht.Axes(1).TickLabelSpacing=5}catch{}
        try{$cht.Axes(2).TickLabels.Font.Size=8}catch{}
        # Ocultar columnas de datos al final (no antes de XValues)
        $ws1.Columns.Item($cdC).Hidden=$true
        $ws1.Columns.Item($cdC+1).Hidden=$true
        $xl.ScreenUpdating=$true
        Write-Host "[$($NOW.ToString('HH:mm:ss'))] GRAFICO OK"
    }catch{Write-Host "  [WARN] Grafico($($_.InvocationInfo.ScriptLineNumber)): $($_.Exception.Message)"}

    MR $ws1 $row 2 $row 7
    W ($ws1.Cells.Item($row,2)) "  DESTACADOS DEL MES" $C.SkyBlue $C.NavyHdr $true 11 -4131; $row++
    if($latestRows.Count-ge 1){
        $top=$latestRows[0]; RowH $ws1 $row 22
        W ($ws1.Cells.Item($row,2)) "$([char]9733) MEJOR OPERARIO" $C.GreenBg $C.GreenFg $true 10 -4108
        W ($ws1.Cells.Item($row,3)) $top.Resp $C.White $C.NavyHdr $true 10 -4131
        MR $ws1 $row 4 $row 5
        W ($ws1.Cells.Item($row,4)) "$($top.LineasDia) lin/dia  ($($top.Cumplim)%)" $C.White $C.GreenTxt $true 10 -4131; $row++
    }
    $bottom=$latestRows|Sort-Object LineasDia|Select-Object -First 1
    if($bottom -and $bottom.LineasDia-lt $TARGET){
        RowH $ws1 $row 22
        W ($ws1.Cells.Item($row,2)) "$([char]9888) BAJO TARGET" $C.RedBg $C.RedFg $true 10 -4108
        W ($ws1.Cells.Item($row,3)) $bottom.Resp $C.White $C.NavyHdr $true 10 -4131
        MR $ws1 $row 4 $row 5
        W ($ws1.Cells.Item($row,4)) "$($bottom.LineasDia) lin/dia  ($($bottom.Cumplim)%)" $C.White $C.RedTxt $true 10 -4131; $row++
    }
    RowH $ws1 $row 6; $row++
    # ── ULTIMOS 7 DIAS ──
    MR $ws1 $row 2 $row 9
    W ($ws1.Cells.Item($row,2)) "  ULTIMOS 7 DIAS DE PICKING (EQUIPO)" $C.SkyBlue $C.NavyHdr $true 11 -4131; $row++
    RowH $ws1 $row 20
    $d7hdrs=@("FECHA","DIA","OPERARIOS","TOTAL LINEAS","LINEAS/OP","TARGET DIA","CUMPL%","VS AYER")
    for($c7=0;$c7 -lt $d7hdrs.Count;$c7++){W ($ws1.Cells.Item($row,$c7+2)) $d7hdrs[$c7] $C.NavyHdr $C.TextWhite $true 9 -4108}
    $row++
    $ridx7=0
    foreach($d7 in $day7Rows){
        RowH $ws1 $row 18
        $rb7=if($ridx7%2){$C.White}else{$C.LightGray}
        $pc7=PerfColors $d7.LinesPerOp
        $cBg7=if($d7.Cumpl-ge 100){$C.GreenBg}elseif($d7.Cumpl-ge 83){$C.YellowBg}else{$C.RedBg}
        $cFg7=if($d7.Cumpl-ge 100){$C.GreenFg}elseif($d7.Cumpl-ge 83){$C.YellowFg}else{$C.RedFg}
        $dStr7=if($null -eq $d7.Delta){"-"}elseif($d7.Delta-ge 0){"+$($d7.Delta)"}else{"$($d7.Delta)"}
        $dFg7=if($null -eq $d7.Delta){$C.Gray}elseif($d7.Delta-gt 50){$C.GreenTxt}elseif($d7.Delta-lt -50){$C.RedTxt}else{$C.Gray}
        W ($ws1.Cells.Item($row,2)) $d7.Fecha       $rb7    $C.TextDark $false 9 -4108
        W ($ws1.Cells.Item($row,3)) $d7.DiaSem      $rb7    $C.TextDark $false 9 -4108
        W ($ws1.Cells.Item($row,4)) $d7.Ops         $rb7    $C.TextDark $false 9 -4108
        W ($ws1.Cells.Item($row,5)) $d7.Lines       $rb7    $C.TextDark $false 9 -4108 "#,##0"
        W ($ws1.Cells.Item($row,6)) $d7.LinesPerOp  $pc7[0] $pc7[1]    $true  9 -4108 "0.0"
        W ($ws1.Cells.Item($row,7)) $d7.Target      $rb7    $C.TextDark $false 9 -4108 "#,##0"
        W ($ws1.Cells.Item($row,8)) "$($d7.Cumpl)%" $cBg7   $cFg7      $true  9 -4108
        W ($ws1.Cells.Item($row,9)) $dStr7          $rb7    $dFg7       $true  9 -4108
        $ridx7++;$row++
    }
    if($day7Rows.Count){
        RowH $ws1 $row 20
        $avg7L =[Math]::Round(($day7Rows|Measure-Object Lines      -Average).Average,0)
        $avg7LD=[Math]::Round(($day7Rows|Measure-Object LinesPerOp -Average).Average,1)
        $avg7C =[Math]::Round(($day7Rows|Measure-Object Cumpl      -Average).Average,1)
        $ldt7=PerfColors $avg7LD
        $cBgA=if($avg7C-ge 100){$C.GreenBg}elseif($avg7C-ge 83){$C.YellowBg}else{$C.RedBg}
        $cFgA=if($avg7C-ge 100){$C.GreenFg}elseif($avg7C-ge 83){$C.YellowFg}else{$C.RedFg}
        W ($ws1.Cells.Item($row,2)) "PROM 7 DIAS"  $C.SkyBlue $C.NavyHdr $true 9 -4131
        W ($ws1.Cells.Item($row,5)) $avg7L         $C.SkyBlue $C.NavyHdr $true 9 -4108 "#,##0"
        W ($ws1.Cells.Item($row,6)) $avg7LD        $ldt7[0]   $ldt7[1]   $true 9 -4108 "0.0"
        W ($ws1.Cells.Item($row,8)) "$avg7C%"      $cBgA      $cFgA      $true 9 -4108
        $row++
    }
    RowH $ws1 $row 6; $row++
    MR $ws1 $row 2 $row 7
    W ($ws1.Cells.Item($row,2)) "  RANKING $($latestNom.ToUpper())" $C.SkyBlue $C.NavyHdr $true 11 -4131; $row++
    ColW $ws1 8 12;ColW $ws1 9 14;ColW $ws1 10 14;ColW $ws1 11 10;ColW $ws1 12 10;ColW $ws1 13 10
    RowH $ws1 $row 20
    @("POS","OPERARIO","DIAS","OLAS","LINEAS","LIN/DIA","CUMPL%","UNIDADES","% C/LOGO","RECLAMOS","TASA","TENDENCIA")|
        ForEach-Object -Begin{$c=2} -Process{W ($ws1.Cells.Item($row,$c)) $_ $C.NavyHdr $C.TextWhite $true 9 -4108;$c++}
    $row++
    $pos=1
    foreach($op in $latestRows){
        RowH $ws1 $row 18
        $rb=if($pos%2){$C.White}else{$C.LightGray}
        $pc=PerfColors $op.LineasDia; $rc=RecColors $op.RecCnt
        $cBg=if($op.Cumplim-ge 100){$C.GreenBg}elseif($op.Cumplim-ge 83){$C.YellowBg}else{$C.RedBg}
        $cFg=if($op.Cumplim-ge 100){$C.GreenFg}elseif($op.Cumplim-ge 83){$C.YellowFg}else{$C.RedFg}
        $tr=if($trendByResp[$op.Resp]){$trendByResp[$op.Resp]}else{@{Txt="-";Fg=$C.Gray}}
        $clBg=if($op.PctConLogo-ge 50){$C.TealLight}else{$rb}
        W ($ws1.Cells.Item($row,2))  $pos           $rb    $C.TextDark $true  9 -4108
        W ($ws1.Cells.Item($row,3))  $op.Resp       $rb    $C.TextDark $false 9 -4131
        W ($ws1.Cells.Item($row,4))  $op.Dias       $rb    $C.TextDark $false 9 -4108
        W ($ws1.Cells.Item($row,5))  $op.Olas       $rb    $C.TextDark $false 9 -4108
        W ($ws1.Cells.Item($row,6))  $op.Lineas     $rb    $C.TextDark $false 9 -4108 "#,##0"
        W ($ws1.Cells.Item($row,7))  $op.LineasDia  $pc[0] $pc[1]      $true  9 -4108 "0.0"
        W ($ws1.Cells.Item($row,8))  "$($op.Cumplim)%" $cBg $cFg       $true  9 -4108
        W ($ws1.Cells.Item($row,9))  $op.Unidades   $rb    $C.TextDark $false 9 -4108 "#,##0"
        W ($ws1.Cells.Item($row,10)) "$($op.PctConLogo)%" $clBg $C.TextDark $false 9 -4108
        W ($ws1.Cells.Item($row,11)) $op.RecCnt     $rc[0] $rc[1]      ($op.RecCnt-gt 0) 9 -4108
        W ($ws1.Cells.Item($row,12)) $op.RecRate    $rb    (RateFg $op.RecRate) $false 9 -4108 "0.00"
        W ($ws1.Cells.Item($row,13)) $tr.Txt        $rb    $tr.Fg      $true  9 -4108
        $pos++;$row++
    }
    RowH $ws1 $row 20
    $ldPcT=PerfColors $avgLD; $rcT=RecColors $totRC
    W ($ws1.Cells.Item($row,2)) "TOTAL / PROMEDIO" $C.SkyBlue $C.NavyHdr $true 9 -4131
    W ($ws1.Cells.Item($row,6))  $totL   $C.SkyBlue $C.NavyHdr $true 9 -4108 "#,##0"
    W ($ws1.Cells.Item($row,7))  $avgLD  $ldPcT[0]  $ldPcT[1]  $true 9 -4108 "0.0"
    W ($ws1.Cells.Item($row,8))  "$cumAv%" $cumBg $cumFg $true 9 -4108
    W ($ws1.Cells.Item($row,9))  $totU   $C.SkyBlue $C.NavyHdr $true 9 -4108 "#,##0"
    W ($ws1.Cells.Item($row,11)) $totRC  $rcT[0] $rcT[1] $true 9 -4108
    Write-Host "[$($NOW.ToString('HH:mm:ss'))] RESUMEN OK"

    # ===========================================================
    # HOJA 2: STAFFING
    # ===========================================================
    ColW $ws2 1 3;ColW $ws2 2 22;ColW $ws2 3 14;ColW $ws2 4 14;ColW $ws2 5 18;ColW $ws2 6 18;ColW $ws2 7 16
    RowH $ws2 1 6;RowH $ws2 2 44;RowH $ws2 3 22;RowH $ws2 4 6
    MR $ws2 2 2 2 7
    W ($ws2.Cells.Item(2,2)) "  STAFFING - PERSONAS NECESARIAS (PICKING REGULAR)" $C.StaffHdr $C.TextWhite $true 18 -4131
    MR $ws2 3 2 3 7
    W ($ws2.Cells.Item(3,2)) "  Solo SIN LOGO + CON LOGO | Formula: Total Lineas / Dias Habiles / $TARGET  |  $($NOW.ToString('dd/MM/yyyy HH:mm'))" $C.StaffLight $C.StaffHdr $false 10 -4131
    $row=6

    if($latestStaff){
        RowH $ws2 $row 4;$row++
        $pnBg2=if($latestStaff.PersonasActuales -ge $latestStaff.PersonasNecesarias){$C.GreenBg}else{$C.YellowBg}
        $pnFg2=if($latestStaff.PersonasActuales -ge $latestStaff.PersonasNecesarias){$C.GreenFg}else{$C.YellowFg}
        $colK=2
        foreach($kd in @(
            @{L="TOTAL LINEAS PICKING ($latestNom)";V=$latestStaff.TotalLineas;Bg=$C.White;Fg=$C.StaffHdr;Fmt="#,##0"},
            @{L="DIAS HABILES";V=$latestStaff.DiasHabiles;Bg=$C.White;Fg=$C.StaffHdr;Fmt=""},
            @{L="PERSONAS NECESARIAS";V=$latestStaff.PersonasNecesarias;Bg=$pnBg2;Fg=$pnFg2;Fmt="0.0"},
            @{L="PERSONAS ACTIVAS (PROM/DIA)";V=$latestStaff.PersonasActuales;Bg=$C.White;Fg=$C.StaffHdr;Fmt="0.0"}
        )){
            RowH $ws2 $row 18; MR $ws2 $row $colK $row ($colK+1)
            W ($ws2.Cells.Item($row,$colK)) $kd.L $C.StaffLight $C.StaffHdr $true 9 -4108
            $row++; RowH $ws2 $row 34; MR $ws2 $row $colK $row ($colK+1)
            $cell=$ws2.Cells.Item($row,$colK)
            if($kd.Fmt){W $cell $kd.V $kd.Bg $kd.Fg $true 22 -4108 $kd.Fmt}else{W $cell "$($kd.V)" $kd.Bg $kd.Fg $true 22 -4108}
            $row-=1; $colK+=2
        }
        $row+=2; RowH $ws2 $row 6; $row++
    }

    MR $ws2 $row 2 $row 7
    W ($ws2.Cells.Item($row,2)) "  HISTORICO POR MES" $C.StaffHdr $C.TextWhite $true 12 -4131; $row++
    $staffHdrRow=$row; RowH $ws2 $row 20
    @("MES","TOTAL LINEAS PICKING","DIAS HABILES","PERSONAS NECES.","PERSONAS ACTIVAS","BRECHA")|
        ForEach-Object -Begin{$c=2} -Process{W ($ws2.Cells.Item($row,$c)) $_ $C.MidGray $C.NavyHdr $true 9 -4108;$c++}
    $row++
    $pos=0
    foreach($sr in ($staffRows|Sort-Object YM -Descending)){
        RowH $ws2 $row 17
        $bg=if($pos%2){$C.StaffLight}else{$C.White}
        $brecha=$sr.PersonasActuales-[Math]::Ceiling($sr.PersonasNecesarias)
        $bBg=if($brecha-ge 0){$C.GreenBg}else{$C.YellowBg}
        $bFg=if($brecha-ge 0){$C.GreenFg}else{$C.YellowFg}
        $pnBg3=if($sr.PersonasActuales-ge $sr.PersonasNecesarias){$C.GreenBg}else{$C.RedBg}
        $pnFg3=if($sr.PersonasActuales-ge $sr.PersonasNecesarias){$C.GreenFg}else{$C.RedFg}
        W ($ws2.Cells.Item($row,2)) $sr.MesNom            $bg      $C.TextDark $false 9 -4131
        W ($ws2.Cells.Item($row,3)) $sr.TotalLineas        $bg      $C.TextDark $false 9 -4108 "#,##0"
        W ($ws2.Cells.Item($row,4)) $sr.DiasHabiles        $bg      $C.TextDark $false 9 -4108
        W ($ws2.Cells.Item($row,5)) $sr.PersonasNecesarias $pnBg3   $pnFg3      $true  9 -4108 "0.0"
        W ($ws2.Cells.Item($row,6)) $sr.PersonasActuales   $bg      $C.TextDark $false 9 -4108 "0.0"
        $brechaStr=if($brecha-ge 0){"+$brecha"}else{"$brecha"}
        W ($ws2.Cells.Item($row,7)) $brechaStr $bBg $bFg $true 9 -4108
        $pos++;$row++
    }
    AF $ws2 $staffHdrRow 2 7
    Write-Host "[$($NOW.ToString('HH:mm:ss'))] STAFFING OK"

    # ===========================================================
    # HOJA 3: RANKING (solo picking regular SIN/CON LOGO)
    # ===========================================================
    RowH $ws3 1 6;RowH $ws3 2 40;RowH $ws3 3 22;RowH $ws3 4 6
    MR $ws3 2 1 2 14
    W ($ws3.Cells.Item(2,1)) "RANKING DE PRODUCTIVIDAD - PICKING REGULAR (SIN LOGO / CON LOGO)" $C.NavyHdr $C.TextWhite $true 18 -4108
    MR $ws3 3 1 3 14
    W ($ws3.Cells.Item(3,1)) "Target: $TARGET lin/dia | Verde=sobre target | Amarillo=cerca | Rojo=bajo | PIE DE MAQUINA y MUESTRA SIMPLE se miden en sus hojas  |  $($NOW.ToString('dd/MM/yyyy HH:mm'))" $C.SkyBlue $C.NavyHdr $false 10 -4108
    $hdrsRk=@("POS","OPERARIO","DIAS","OLAS","LINEAS","LIN/DIA","TARGET","CUMPL%","DIFF","UNIDADES","U/LINEA","% C/LOGO","RECLAMOS","TENDENCIA")
    $row=5
    foreach($ym in ($sortedMon|Sort-Object -Descending)){
        $ymp=$ym.Split("-"); $ymNom="$($MES_NOM[[int]$ymp[1]]) $($ymp[0])"
        $mOps=@($sumRows|Where-Object{$_.YM -eq $ym}|Sort-Object LineasDia -Descending)
        if(-not $mOps.Count){continue}
        RowH $ws3 $row 22; MR $ws3 $row 1 $row 14
        W ($ws3.Cells.Item($row,1)) "  $($ymNom.ToUpper())" $C.BlueHdr $C.TextWhite $true 12 -4131; $row++
        $rkHdrRow=$row; RowH $ws3 $row 20
        for($c=0;$c -lt $hdrsRk.Count;$c++){W ($ws3.Cells.Item($row,$c+1)) $hdrsRk[$c] $C.MidGray $C.NavyHdr $true 9 -4108}
        $row++
        $pos=1
        foreach($op in $mOps){
            RowH $ws3 $row 17
            $rb=if($pos%2){$C.White}else{$C.LightGray}
            $pc=PerfColors $op.LineasDia; $rc=RecColors $op.RecCnt
            $dTx=if($op.Diff-ge 0){"+"+"$($op.Diff)"}else{"$($op.Diff)"}
            $dBg=if($op.Diff-ge 0){$C.GreenBg}else{$C.RedBg}; $dFg=if($op.Diff-ge 0){$C.GreenFg}else{$C.RedFg}
            $cBg=if($op.Cumplim-ge 100){$C.GreenBg}elseif($op.Cumplim-ge 83){$C.YellowBg}else{$C.RedBg}
            $cFg=if($op.Cumplim-ge 100){$C.GreenFg}elseif($op.Cumplim-ge 83){$C.YellowFg}else{$C.RedFg}
            $tr=if($trendByResp[$op.Resp]){$trendByResp[$op.Resp]}else{@{Txt="-";Fg=$C.Gray}}
            $clBg=if($op.PctConLogo-ge 50){$C.TealLight}else{$rb}
            W ($ws3.Cells.Item($row,1))  $pos          $rb    $C.TextDark $true  9 -4108
            W ($ws3.Cells.Item($row,2))  $op.Resp      $rb    $C.TextDark $false 9 -4131
            W ($ws3.Cells.Item($row,3))  $op.Dias      $rb    $C.TextDark $false 9 -4108
            W ($ws3.Cells.Item($row,4))  $op.Olas      $rb    $C.TextDark $false 9 -4108
            W ($ws3.Cells.Item($row,5))  $op.Lineas    $rb    $C.TextDark $false 9 -4108 "#,##0"
            W ($ws3.Cells.Item($row,6))  $op.LineasDia $pc[0] $pc[1]      $true  9 -4108 "0.0"
            W ($ws3.Cells.Item($row,7))  $op.Target    $rb    $C.TextDark $false 9 -4108 "#,##0"
            W ($ws3.Cells.Item($row,8))  "$($op.Cumplim)%" $cBg $cFg $true 9 -4108
            W ($ws3.Cells.Item($row,9))  $dTx          $dBg   $dFg        $false 9 -4108
            W ($ws3.Cells.Item($row,10)) $op.Unidades  $rb    $C.TextDark $false 9 -4108 "#,##0"
            W ($ws3.Cells.Item($row,11)) $op.ULinea    $rb    $C.TextDark $false 9 -4108 "0.00"
            W ($ws3.Cells.Item($row,12)) "$($op.PctConLogo)%" $clBg $C.TextDark $false 9 -4108
            W ($ws3.Cells.Item($row,13)) $op.RecCnt    $rc[0] $rc[1]      ($op.RecCnt-gt 0) 9 -4108
            W ($ws3.Cells.Item($row,14)) $tr.Txt       $rb    $tr.Fg      $true  9 -4108
            $pos++;$row++
        }
        RowH $ws3 $row 18
        $mL=0;$mU=0;$mRC=0
        foreach($op in $mOps){$mL+=$op.Lineas;$mU+=$op.Unidades;$mRC+=$op.RecCnt}
        $mLD=if($mOps.Count){[Math]::Round(($mOps|Measure-Object LineasDia -Average).Average,1)}else{0}
        $ldT=PerfColors $mLD; $rcM=RecColors $mRC
        W ($ws3.Cells.Item($row,2))  "TOTAL / PROMEDIO" $C.SkyBlue $C.NavyHdr $true 9 -4131
        W ($ws3.Cells.Item($row,5))  $mL  $C.SkyBlue $C.NavyHdr $true 9 -4108 "#,##0"
        W ($ws3.Cells.Item($row,6))  $mLD $ldT[0] $ldT[1] $true 9 -4108 "0.0"
        W ($ws3.Cells.Item($row,10)) $mU  $C.SkyBlue $C.NavyHdr $true 9 -4108 "#,##0"
        W ($ws3.Cells.Item($row,13)) $mRC $rcM[0] $rcM[1] $true 9 -4108
        AF $ws3 $rkHdrRow 1 14
        RowH $ws3 ($row+1) 8; $row+=2
    }
    $ws3.UsedRange.EntireColumn.AutoFit()|Out-Null; ColW $ws3 2 24
    Write-Host "[$($NOW.ToString('HH:mm:ss'))] RANKING OK"

    # ===========================================================
    # HOJA 4: EVOLUCION_MENSUAL
    # ===========================================================
    $nMon=$sortedMon.Count
    RowH $ws4 1 6;RowH $ws4 2 36;RowH $ws4 3 22;RowH $ws4 4 6
    MR $ws4 2 1 2 ($nMon+5)
    W ($ws4.Cells.Item(2,1)) "EVOLUCION MENSUAL DE LINEAS/DIA POR OPERARIO (PICKING REGULAR)" $C.NavyHdr $C.TextWhite $true 16 -4108
    MR $ws4 3 1 3 ($nMon+5)
    W ($ws4.Cells.Item(3,1)) "| Verde >= $TARGET  | Amarillo 70-$($TARGET-1)  | Rojo < 70  | Solo SIN LOGO / CON LOGO  |  $($NOW.ToString('dd/MM/yyyy HH:mm'))" $C.SkyBlue $C.NavyHdr $false 10 -4108
    $row=5; $evolHdrRow=$row; RowH $ws4 $row 30
    W ($ws4.Cells.Item($row,1)) "OPERARIO" $C.NavyHdr $C.TextWhite $true 10 -4131
    for($m=0;$m -lt $nMon;$m++){
        $ymp=$sortedMon[$m].Split("-")
        $mLbl="$($MES_NOM[[int]$ymp[1]])`n$($ymp[0])"
        $cell=$ws4.Cells.Item($row,$m+2)
        $cell.Value2=$mLbl;$cell.Interior.Color=$C.NavyHdr;$cell.Font.Color=$C.TextWhite
        $cell.Font.Bold=$true;$cell.Font.Size=9;$cell.HorizontalAlignment=-4108
        $cell.WrapText=$true;$cell.VerticalAlignment=-4108
    }
    W ($ws4.Cells.Item($row,$nMon+2)) "PROM"    $C.NavyHdr $C.TextWhite $true 9 -4108
    W ($ws4.Cells.Item($row,$nMon+3)) "TEND"    $C.NavyHdr $C.TextWhite $true 9 -4108
    W ($ws4.Cells.Item($row,$nMon+4)) "RANK"    $C.NavyHdr $C.TextWhite $true 9 -4108
    W ($ws4.Cells.Item($row,$nMon+5)) "%C/LOGO" $C.NavyHdr $C.TextWhite $true 9 -4108
    $row++
    $idx=0
    foreach($resp in $sortedResp){
        RowH $ws4 $row 17
        $rb=if($idx%2){$C.LightGray}else{$C.White}
        W ($ws4.Cells.Item($row,1)) $resp $rb $C.TextDark $false 9 -4131
        $allLD=@();$totLinR=0;$totOlasR=0;$totCLR=0
        for($m=0;$m -lt $nMon;$m++){
            $k="$resp|$($sortedMon[$m])"
            if($pkMes[$k]){
                $pd=$pkMes[$k];$days=$pd.Days.Count
                $ld=if($days){[Math]::Round($pd.L/$days,1)}else{0}
                $allLD+=$ld;$totLinR+=$pd.L;$totOlasR+=$pd.Olas;$totCLR+=$pd.CL
                $pc=PerfColors $ld; W ($ws4.Cells.Item($row,$m+2)) $ld $pc[0] $pc[1] $true 9 -4108 "0.0"
            }else{
                $cell=$ws4.Cells.Item($row,$m+2)
                $cell.Interior.Color=$C.MidGray;$cell.Value2="-"
                $cell.Font.Color=$C.Gray;$cell.Font.Size=9;$cell.HorizontalAlignment=-4108
            }
        }
        if($allLD.Count){
            $oAvg=[Math]::Round(($allLD|Measure-Object -Average).Average,1)
            $pc=PerfColors $oAvg; W ($ws4.Cells.Item($row,$nMon+2)) $oAvg $pc[0] $pc[1] $true 10 -4108 "0.0"
        }
        $tr=if($trendByResp[$resp]){$trendByResp[$resp]}else{@{Txt="-";Fg=$C.Gray}}
        W ($ws4.Cells.Item($row,$nMon+3)) $tr.Txt $rb $tr.Fg $true 9 -4108
        $rk=$rankByPkMon["$resp|$latestYM"]
        $rkTxt=if($rk){"#$rk"}else{"-"}
        W ($ws4.Cells.Item($row,$nMon+4)) $rkTxt $rb $C.TextDark $true 9 -4108
        $pctCLR=if($totOlasR){[Math]::Round($totCLR/$totOlasR*100,0)}else{0}
        $clBg=if($pctCLR-ge 50){$C.TealLight}else{$rb}
        W ($ws4.Cells.Item($row,$nMon+5)) "$pctCLR%" $clBg $C.TextDark $false 9 -4108
        $idx++;$row++
    }
    RowH $ws4 $row 20
    W ($ws4.Cells.Item($row,1)) "TARGET DIARIO" $C.SkyBlue $C.NavyHdr $true 9 -4131
    for($m=0;$m -lt $nMon;$m++){ W ($ws4.Cells.Item($row,$m+2)) $TARGET $C.SkyBlue $C.NavyHdr $true 9 -4108 }
    W ($ws4.Cells.Item($row,$nMon+2)) $TARGET $C.SkyBlue $C.NavyHdr $true 9 -4108
    $ws4.UsedRange.EntireColumn.AutoFit()|Out-Null; ColW $ws4 1 24
    for($m=2;$m -le ($nMon+1);$m++){$w=$ws4.Columns.Item($m).ColumnWidth;if($w-gt 12){ColW $ws4 $m 10}}
    Write-Host "[$($NOW.ToString('HH:mm:ss'))] EVOLUCION_MENSUAL OK"

    # ===========================================================
    # HOJA 5: PRODUCCION_DIARIA (heatmap todos los operarios)
    # ===========================================================
    $allDates=($pkDay.Keys|ForEach-Object{$_.Split("|")[1]}|Sort-Object -Unique|Select-Object -Last 60)
    $nDays=$allDates.Count
    $allOps2=@($sortedResp)+@("LEZCANO AGUSTIN","OPERARIO MERMA 1","OPERARIO MERMA 2")|
        Where-Object{$allOperators[$_]}|Select-Object -Unique
    RowH $ws5 1 6;RowH $ws5 2 36;RowH $ws5 3 22;RowH $ws5 4 6
    MR $ws5 2 1 2 ($nDays+2)
    W ($ws5.Cells.Item(2,1)) "PRODUCCION DIARIA - LINEAS POR OPERARIO (ULTIMOS $nDays DIAS)" $C.NavyHdr $C.TextWhite $true 16 -4108
    MR $ws5 3 1 3 ($nDays+2)
    W ($ws5.Cells.Item(3,1)) "MERMA 1/2 incluye PEDIDO MERMA + PIE DE MAQUINA combinados | Verde >= $TARGET | Amarillo 70-$($TARGET-1) | Rojo < 70" $C.SkyBlue $C.NavyHdr $false 10 -4108
    $row=5; RowH $ws5 $row 45
    W ($ws5.Cells.Item($row,1)) "OPERARIO" $C.NavyHdr $C.TextWhite $true 10 -4131
    for($d=0;$d -lt $nDays;$d++){
        $dt=[datetime]::Parse($allDates[$d])
        $lbl="$($dt.Day.ToString('00'))/$($dt.Month.ToString('00'))`n$(@('Do','Lu','Ma','Mi','Ju','Vi','Sa')[[int]$dt.DayOfWeek])"
        $cell=$ws5.Cells.Item($row,$d+2)
        $cell.Value2=$lbl;$cell.Interior.Color=$C.NavyHdr;$cell.Font.Color=$C.TextWhite
        $cell.Font.Bold=$true;$cell.Font.Size=8;$cell.HorizontalAlignment=-4108;$cell.WrapText=$true
    }
    W ($ws5.Cells.Item($row,$nDays+2)) "PROM/DIA" $C.NavyHdr $C.TextWhite $true 9 -4108
    W ($ws5.Cells.Item($row,$nDays+3)) "PROM 7D"  $C.NavyHdr $C.TextWhite $true 9 -4108
    $last7DAll=@($allDates|Select-Object -Last 7)
    $row++
    $idx=0
    foreach($resp in $allOps2){
        RowH $ws5 $row 17
        $rb=if($idx%2){$C.LightGray}else{$C.White}
        W ($ws5.Cells.Item($row,1)) $resp $rb $C.TextDark $false 9 -4131
        $dayVals=@()
        for($d=0;$d -lt $nDays;$d++){
            $dk="$resp|$($allDates[$d])"
            $cell=$ws5.Cells.Item($row,$d+2)
            if($pkDay[$dk]){
                $ld=[Math]::Round($pkDay[$dk],0); $dayVals+=$ld
                $pc=PerfColors $ld; W $cell $ld $pc[0] $pc[1] $true 8 -4108
            }else{
                $cell.Interior.Color=$C.MidGray;$cell.Value2="-"
                $cell.Font.Color=$C.Gray;$cell.Font.Size=8;$cell.HorizontalAlignment=-4108
            }
        }
        if($dayVals.Count){
            $avg=[Math]::Round(($dayVals|Measure-Object -Average).Average,1)
            $pc=PerfColors $avg
            W ($ws5.Cells.Item($row,$nDays+2)) $avg $pc[0] $pc[1] $true 9 -4108 "0.0"
        }
        # Promedio ultimos 7 dias
        $dv7p=@()
        foreach($d7y2 in $last7DAll){
            $dk7p2="$resp|$d7y2"
            if($pkDay[$dk7p2]){$dv7p+=[Math]::Round($pkDay[$dk7p2],0)}
        }
        if($dv7p.Count){
            $a7p=[Math]::Round(($dv7p|Measure-Object -Average).Average,1)
            $pc7p=PerfColors $a7p
            W ($ws5.Cells.Item($row,$nDays+3)) $a7p $pc7p[0] $pc7p[1] $true 9 -4108 "0.0"
        }
        $idx++;$row++
    }
    $nDaysInt=[int]$nDays; ColW $ws5 1 24
    for($d=2;$d -le ($nDaysInt+1);$d++){try{ColW $ws5 $d 5.5}catch{}}
    try{ColW $ws5 ($nDaysInt+2) 9}catch{}
    try{ColW $ws5 ($nDaysInt+3) 9}catch{}
    Write-Host "[$($NOW.ToString('HH:mm:ss'))] PRODUCCION_DIARIA OK"

    # ===========================================================
    # HOJA 6: LEZCANO (MUESTRA SIMPLE)
    # ===========================================================
    RowH $ws6 1 6;RowH $ws6 2 44;RowH $ws6 3 24;RowH $ws6 4 6
    MR $ws6 2 1 2 11
    W ($ws6.Cells.Item(2,1)) "  LEZCANO AGUSTIN -- MUESTRA SIMPLE" $C.LezHdr $C.TextWhite $true 18 -4131
    MR $ws6 3 1 3 8
    W ($ws6.Cells.Item(3,1)) "  Fuente: Picking (MUESTRA SIMPLE) | Target: $TARGET lin/dia" $C.LezMid $C.LezHdr $true 11 -4131
    MR $ws6 3 9 3 11
    W ($ws6.Cells.Item(3,9)) "Actualizado: $($NOW.ToString('dd/MM/yyyy HH:mm'))" $C.LezMid $C.LezHdr $false 9 -4152
    ColW $ws6 1 3;ColW $ws6 2 18;ColW $ws6 3 8;ColW $ws6 4 8;ColW $ws6 5 14;ColW $ws6 6 12
    ColW $ws6 7 14;ColW $ws6 8 10;ColW $ws6 9 10;ColW $ws6 10 10;ColW $ws6 11 14
    if($lezRows.Count -eq 0){
        MR $ws6 6 1 6 11
        W ($ws6.Cells.Item(6,1)) "Sin datos de LEZCANO AGUSTIN (MUESTRA SIMPLE) en Picking." $C.LightGray $C.Gray $false 11 -4108
    }else{
        $lezAvgLD=[Math]::Round(($lezRows|Measure-Object LineasDia -Average).Average,1)
        $lezTotL=($lezRows|Measure-Object Lineas -Sum).Sum
        $lezTotRec=($lezRows|Measure-Object RecCnt -Sum).Sum
        $row=6; RowH $ws6 $row 4;$row++
        $lezPc=PerfColors $lezAvgLD
        $lezRcBg=if($lezTotRec-eq 0){$C.GreenBg}elseif($lezTotRec-le 5){$C.YellowBg}else{$C.RedBg}
        $lezRcFg=if($lezTotRec-eq 0){$C.GreenFg}elseif($lezTotRec-le 5){$C.YellowFg}else{$C.RedFg}
        $r1=LezKPI $ws6 $row 2 4  "LIN/DIA PROMEDIO"    "$lezAvgLD (target $TARGET)" $lezPc[0] $lezPc[1]
        $r2=LezKPI $ws6 $row 5 7  "TOTAL LINEAS HIST."   $lezTotL   $C.White $C.LezHdr "#,##0"
        $r3=LezKPI $ws6 $row 8 11 "RECLAMOS HISTORICO"   $lezTotRec $lezRcBg $lezRcFg
        $row=[Math]::Max([Math]::Max($r1,$r2),$r3)
        RowH $ws6 $row 6;$row++
        MR $ws6 $row 2 $row 11
        W ($ws6.Cells.Item($row,2)) "  EVOLUCION MES A MES" $C.LezHdr $C.TextWhite $true 12 -4131;$row++
        $lezHdrRow=$row; RowH $ws6 $row 20
        @("MES","DIAS","OLAS","LINEAS","LIN/DIA","vs TARGET","vs EQUIPO PICK","UNIDADES","U/LINEA","RECLAMOS","TENDENCIA")|
            ForEach-Object -Begin{$c=2} -Process{W ($ws6.Cells.Item($row,$c)) $_ $C.LezHdr $C.TextWhite $true 9 -4108;$c++}
        $row++
        $idx=0
        foreach($lr in $lezRows){
            RowH $ws6 $row 18
            $rb=if($idx%2){$C.LezLight}else{$C.White}
            $pc=PerfColors $lr.LineasDia
            $vsT=[Math]::Round($lr.LineasDia-$TARGET,1)
            $vsTTxt=if($vsT-ge 0){"+$vsT"}else{"$vsT"}
            $vsTBg=if($vsT-ge 0){$C.GreenBg}else{$C.RedBg}; $vsTFg=if($vsT-ge 0){$C.GreenFg}else{$C.RedFg}
            $vsEqFg=if($lr.VsTeam-ge 5){$C.GreenTxt}elseif($lr.VsTeam-ge -5){$C.YellowTxt}else{$C.RedTxt}
            $vsEqTx=if($lr.VsTeam-ge 0){"+$($lr.VsTeam)"}else{"$($lr.VsTeam)"}
            $rcBg=if($lr.RecCnt-eq 0){$C.GreenBg}elseif($lr.RecCnt-le 3){$C.YellowBg}else{$C.RedBg}
            $rcFg=if($lr.RecCnt-eq 0){$C.GreenFg}elseif($lr.RecCnt-le 3){$C.YellowFg}else{$C.RedFg}
            $trFgLez=$C.Gray
            if($lr.TrendStr -like "*$ARR_UP*"){$trFgLez=$C.GreenTxt}
            if($lr.TrendStr -like "*$ARR_DOWN*"){$trFgLez=$C.RedTxt}
            W ($ws6.Cells.Item($row,2))  $lr.MesNom    $rb     $C.TextDark $true  9 -4131
            W ($ws6.Cells.Item($row,3))  $lr.Dias      $rb     $C.TextDark $false 9 -4108
            W ($ws6.Cells.Item($row,4))  $lr.Olas      $rb     $C.TextDark $false 9 -4108
            W ($ws6.Cells.Item($row,5))  $lr.Lineas    $rb     $C.TextDark $false 9 -4108 "#,##0"
            W ($ws6.Cells.Item($row,6))  $lr.LineasDia $pc[0]  $pc[1]      $true  9 -4108 "0.0"
            W ($ws6.Cells.Item($row,7))  $vsTTxt       $vsTBg  $vsTFg      $true  9 -4108
            W ($ws6.Cells.Item($row,8))  $vsEqTx       $rb     $vsEqFg     $true  9 -4108
            W ($ws6.Cells.Item($row,9))  $lr.Unidades  $rb     $C.TextDark $false 9 -4108 "#,##0"
            W ($ws6.Cells.Item($row,10)) $lr.ULinea    $rb     $C.TextDark $false 9 -4108 "0.00"
            W ($ws6.Cells.Item($row,11)) $lr.RecCnt    $rcBg   $rcFg       ($lr.RecCnt-gt 0) 9 -4108
            W ($ws6.Cells.Item($row,12)) $lr.TrendStr  $rb     $trFgLez    $true  9 -4108
            $idx++;$row++
        }
        RowH $ws6 $row 20
        $lezPcAvg=PerfColors $lezAvgLD; $rcTLez=RecColors $lezTotRec
        W ($ws6.Cells.Item($row,2)) "PROMEDIO / TOTAL" $C.LezMid $C.LezHdr $true 9 -4131
        W ($ws6.Cells.Item($row,5)) $lezTotL $C.LezMid $C.LezHdr $true 9 -4108 "#,##0"
        W ($ws6.Cells.Item($row,6)) $lezAvgLD $lezPcAvg[0] $lezPcAvg[1] $true 9 -4108 "0.0"
        W ($ws6.Cells.Item($row,11)) $lezTotRec $rcTLez[0] $rcTLez[1] $true 9 -4108
        AF $ws6 $lezHdrRow 2 12
    }
    Write-Host "[$($NOW.ToString('HH:mm:ss'))] LEZCANO OK"

    # ===========================================================
    # HOJA 7: PIE_MAQUINA (medicion separada, sin target)
    # ===========================================================
    ColW $ws7 1 3;ColW $ws7 2 18;ColW $ws7 3 10;ColW $ws7 4 12;ColW $ws7 5 12;ColW $ws7 6 12
    ColW $ws7 7 10;ColW $ws7 8 12;ColW $ws7 9 12;ColW $ws7 10 12;ColW $ws7 11 14
    RowH $ws7 1 6;RowH $ws7 2 44;RowH $ws7 3 22;RowH $ws7 4 6
    MR $ws7 2 2 2 11
    W ($ws7.Cells.Item(2,2)) "  PIE DE MAQUINA - MERMA 1 vs MERMA 2" $C.PieHdr $C.TextWhite $true 18 -4131
    MR $ws7 3 2 3 11
    W ($ws7.Cells.Item(3,2)) "  MERMA 1 = turno 06:00-14:59 | MERMA 2 = resto del dia | Fuente: Picking PIE DE MAQUINA | Sin target (depende de produccion)  |  $($NOW.ToString('dd/MM/yyyy HH:mm'))" $C.PieLight $C.PieHdr $false 10 -4131
    $row=6

    if($pieRows.Count -eq 0){
        MR $ws7 $row 2 $row 11
        W ($ws7.Cells.Item($row,2)) "Sin datos de PIE DE MAQUINA en Picking." $C.LightGray $C.Gray $false 11 -4108
    } else {
        $totP1L=($pieRows|Measure-Object M1Lineas -Sum).Sum
        $totP2L=($pieRows|Measure-Object M2Lineas -Sum).Sum
        $totP1D=0; foreach($pr in $pieRows){$totP1D+=$pr.M1Dias}
        $totP2D=0; foreach($pr in $pieRows){$totP2D+=$pr.M2Dias}
        $avgP1LD=if($totP1D){[Math]::Round($totP1L/$totP1D,1)}else{0}
        $avgP2LD=if($totP2D){[Math]::Round($totP2L/$totP2D,1)}else{0}

        RowH $ws7 $row 4;$row++
        $colK=2
        foreach($kd in @(
            @{L="MERMA 1 - LIN/DIA PROM";V=$avgP1LD;Bg=$C.PieHdr;Fg=$C.TextWhite;Fmt="0.0"},
            @{L="MERMA 1 - TOTAL LINEAS";V=$totP1L;Bg=$C.PieLight;Fg=$C.PieHdr;Fmt="#,##0"},
            @{L="MERMA 2 - LIN/DIA PROM";V=$avgP2LD;Bg=$C.PieHdr;Fg=$C.TextWhite;Fmt="0.0"},
            @{L="MERMA 2 - TOTAL LINEAS";V=$totP2L;Bg=$C.PieLight;Fg=$C.PieHdr;Fmt="#,##0"}
        )){
            RowH $ws7 $row 18; MR $ws7 $row $colK $row ($colK+1)
            W ($ws7.Cells.Item($row,$colK)) $kd.L $C.PieLight $C.PieHdr $true 9 -4108
            $row++; RowH $ws7 $row 32; MR $ws7 $row $colK $row ($colK+1)
            $cell=$ws7.Cells.Item($row,$colK)
            if($kd.Fmt){W $cell $kd.V $kd.Bg $kd.Fg $true 22 -4108 $kd.Fmt}else{W $cell "$($kd.V)" $kd.Bg $kd.Fg $true 22 -4108}
            $row-=1; $colK+=2
        }
        $row+=2; RowH $ws7 $row 6;$row++

        MR $ws7 $row 2 $row 11
        W ($ws7.Cells.Item($row,2)) "  DETALLE POR MES" $C.PieHdr $C.TextWhite $true 12 -4131;$row++
        $pieHdrRow=$row; RowH $ws7 $row 20
        @("MES","M1 DIAS","M1 LINEAS","M1 LIN/DIA","M1 UNIDADES","M2 DIAS","M2 LINEAS","M2 LIN/DIA","M2 UNIDADES","TOTAL LINEAS")|
            ForEach-Object -Begin{$c=2} -Process{W ($ws7.Cells.Item($row,$c)) $_ $C.MidGray $C.NavyHdr $true 9 -4108;$c++}
        $row++
        $pos=0
        foreach($pr in ($pieRows|Sort-Object YM -Descending)){
            RowH $ws7 $row 17
            $bg=if($pos%2){$C.LightGray}else{$C.White}
            W ($ws7.Cells.Item($row,2))  $pr.MesNom     $bg         $C.TextDark $false 9 -4131
            W ($ws7.Cells.Item($row,3))  $pr.M1Dias     $bg         $C.TextDark $false 9 -4108
            W ($ws7.Cells.Item($row,4))  $pr.M1Lineas   $C.PieLight $C.PieHdr  $false 9 -4108 "#,##0"
            W ($ws7.Cells.Item($row,5))  $pr.M1LD       $C.PieLight $C.PieHdr  $true  9 -4108 "0.0"
            W ($ws7.Cells.Item($row,6))  $pr.M1Unidades $C.PieLight $C.PieHdr  $false 9 -4108 "#,##0"
            W ($ws7.Cells.Item($row,7))  $pr.M2Dias     $bg         $C.TextDark $false 9 -4108
            W ($ws7.Cells.Item($row,8))  $pr.M2Lineas   $C.MidGray  $C.NavyHdr  $false 9 -4108 "#,##0"
            W ($ws7.Cells.Item($row,9))  $pr.M2LD       $C.MidGray  $C.NavyHdr  $true  9 -4108 "0.0"
            W ($ws7.Cells.Item($row,10)) $pr.M2Unidades $C.MidGray  $C.NavyHdr  $false 9 -4108 "#,##0"
            W ($ws7.Cells.Item($row,11)) $pr.TotalLineas $bg        $C.TextDark $true  9 -4108 "#,##0"
            $pos++;$row++
        }
        RowH $ws7 $row 18
        W ($ws7.Cells.Item($row,2))  "TOTAL"  $C.PieMid  $C.PieHdr   $true 9 -4131
        W ($ws7.Cells.Item($row,4))  $totP1L  $C.PieHdr  $C.TextWhite $true 9 -4108 "#,##0"
        W ($ws7.Cells.Item($row,5))  $avgP1LD $C.PieLight $C.PieHdr  $true 9 -4108 "0.0"
        W ($ws7.Cells.Item($row,8))  $totP2L  $C.PieHdr  $C.TextWhite $true 9 -4108 "#,##0"
        W ($ws7.Cells.Item($row,9))  $avgP2LD $C.PieLight $C.PieHdr  $true 9 -4108 "0.0"
        W ($ws7.Cells.Item($row,11)) ($totP1L+$totP2L) $C.PieMid $C.PieHdr $true 9 -4108 "#,##0"
        AF $ws7 $pieHdrRow 2 11
    }
    Write-Host "[$($NOW.ToString('HH:mm:ss'))] PIE_MAQUINA OK"

    # ===========================================================
    # HOJA 8: MERMA (PEDIDO MERMA, sin target)
    # ===========================================================
    ColW $ws8 1 3;ColW $ws8 2 18;ColW $ws8 3 10;ColW $ws8 4 12;ColW $ws8 5 12;ColW $ws8 6 12
    ColW $ws8 7 10;ColW $ws8 8 12;ColW $ws8 9 12;ColW $ws8 10 12;ColW $ws8 11 14
    RowH $ws8 1 6;RowH $ws8 2 44;RowH $ws8 3 22;RowH $ws8 4 6
    MR $ws8 2 2 2 11
    W ($ws8.Cells.Item(2,2)) "  PEDIDO MERMA - MERMA 1 vs MERMA 2" $C.MermaHdr $C.TextWhite $true 18 -4131
    MR $ws8 3 2 3 11
    W ($ws8.Cells.Item(3,2)) "  MERMA 1 = turno 06:00-14:59 | MERMA 2 = resto del dia | Fuente: Picking PEDIDO MERMA | Sin target (sale de lo que merma produccion)  |  $($NOW.ToString('dd/MM/yyyy HH:mm'))" $C.MermaLight $C.MermaHdr $false 10 -4131
    $row=6

    if($mermaRows.Count -eq 0){
        MR $ws8 $row 2 $row 11
        W ($ws8.Cells.Item($row,2)) "Sin datos de PEDIDO MERMA en Picking." $C.LightGray $C.Gray $false 11 -4108
    } else {
        $totM1L=($mermaRows|Measure-Object M1Lineas -Sum).Sum
        $totM2L=($mermaRows|Measure-Object M2Lineas -Sum).Sum
        $totM1D=0; foreach($mr in $mermaRows){$totM1D+=$mr.M1Dias}
        $totM2D=0; foreach($mr in $mermaRows){$totM2D+=$mr.M2Dias}
        $avgM1LD=if($totM1D){[Math]::Round($totM1L/$totM1D,1)}else{0}
        $avgM2LD=if($totM2D){[Math]::Round($totM2L/$totM2D,1)}else{0}

        RowH $ws8 $row 4;$row++
        $colK=2
        foreach($kd in @(
            @{L="MERMA 1 - LIN/DIA PROM";V=$avgM1LD;Bg=$C.MermaHdr;Fg=$C.TextWhite;Fmt="0.0"},
            @{L="MERMA 1 - TOTAL LINEAS";V=$totM1L;Bg=$C.MermaLight;Fg=$C.MermaHdr;Fmt="#,##0"},
            @{L="MERMA 2 - LIN/DIA PROM";V=$avgM2LD;Bg=$C.Merma2Hdr;Fg=$C.TextWhite;Fmt="0.0"},
            @{L="MERMA 2 - TOTAL LINEAS";V=$totM2L;Bg=$C.Merma2Lt;Fg=$C.Merma2Hdr;Fmt="#,##0"}
        )){
            RowH $ws8 $row 18; MR $ws8 $row $colK $row ($colK+1)
            W ($ws8.Cells.Item($row,$colK)) $kd.L $C.MermaLight $C.MermaHdr $true 9 -4108
            $row++; RowH $ws8 $row 32; MR $ws8 $row $colK $row ($colK+1)
            $cell=$ws8.Cells.Item($row,$colK)
            if($kd.Fmt){W $cell $kd.V $kd.Bg $kd.Fg $true 22 -4108 $kd.Fmt}else{W $cell "$($kd.V)" $kd.Bg $kd.Fg $true 22 -4108}
            $row-=1; $colK+=2
        }
        $row+=2; RowH $ws8 $row 6;$row++

        MR $ws8 $row 2 $row 11
        W ($ws8.Cells.Item($row,2)) "  DETALLE POR MES" $C.MermaHdr $C.TextWhite $true 12 -4131;$row++
        $mermaHdrRow=$row; RowH $ws8 $row 20
        @("MES","M1 DIAS","M1 LINEAS","M1 LIN/DIA","M1 UNIDADES","M2 DIAS","M2 LINEAS","M2 LIN/DIA","M2 UNIDADES","TOTAL LINEAS")|
            ForEach-Object -Begin{$c=2} -Process{W ($ws8.Cells.Item($row,$c)) $_ $C.MidGray $C.NavyHdr $true 9 -4108;$c++}
        $row++
        $pos=0
        foreach($mr in ($mermaRows|Sort-Object YM -Descending)){
            RowH $ws8 $row 17
            $bg=if($pos%2){$C.LightGray}else{$C.White}
            # Sin colores de performance (sin target): colores fijos por operario
            W ($ws8.Cells.Item($row,2))  $mr.MesNom     $bg            $C.TextDark  $false 9 -4131
            W ($ws8.Cells.Item($row,3))  $mr.M1Dias     $bg            $C.TextDark  $false 9 -4108
            W ($ws8.Cells.Item($row,4))  $mr.M1Lineas   $C.MermaLight  $C.MermaHdr  $false 9 -4108 "#,##0"
            W ($ws8.Cells.Item($row,5))  $mr.M1LD       $C.MermaHdr    $C.TextWhite $true  9 -4108 "0.0"
            W ($ws8.Cells.Item($row,6))  $mr.M1Unidades $C.MermaLight  $C.MermaHdr  $false 9 -4108 "#,##0"
            W ($ws8.Cells.Item($row,7))  $mr.M2Dias     $bg            $C.TextDark  $false 9 -4108
            W ($ws8.Cells.Item($row,8))  $mr.M2Lineas   $C.Merma2Lt    $C.Merma2Hdr $false 9 -4108 "#,##0"
            W ($ws8.Cells.Item($row,9))  $mr.M2LD       $C.Merma2Hdr   $C.TextWhite $true  9 -4108 "0.0"
            W ($ws8.Cells.Item($row,10)) $mr.M2Unidades $C.Merma2Lt    $C.Merma2Hdr $false 9 -4108 "#,##0"
            W ($ws8.Cells.Item($row,11)) $mr.TotalLineas $bg           $C.TextDark  $true  9 -4108 "#,##0"
            $pos++;$row++
        }
        RowH $ws8 $row 18
        W ($ws8.Cells.Item($row,2))  "TOTAL"   $C.MermaMid  $C.MermaHdr  $true 9 -4131
        W ($ws8.Cells.Item($row,4))  $totM1L   $C.MermaHdr  $C.TextWhite $true 9 -4108 "#,##0"
        W ($ws8.Cells.Item($row,5))  $avgM1LD  $C.MermaLight $C.MermaHdr $true 9 -4108 "0.0"
        W ($ws8.Cells.Item($row,8))  $totM2L   $C.Merma2Hdr $C.TextWhite $true 9 -4108 "#,##0"
        W ($ws8.Cells.Item($row,9))  $avgM2LD  $C.Merma2Lt  $C.Merma2Hdr $true 9 -4108 "0.0"
        W ($ws8.Cells.Item($row,11)) ($totM1L+$totM2L) $C.MermaMid $C.MermaHdr $true 9 -4108 "#,##0"
        AF $ws8 $mermaHdrRow 2 11
    }
    Write-Host "[$($NOW.ToString('HH:mm:ss'))] MERMA OK"

    # ===========================================================
    # HOJA 9: RECLAMOS
    # ===========================================================
    RowH $ws9 1 6;RowH $ws9 2 36;RowH $ws9 3 22;RowH $ws9 4 6
    MR $ws9 2 1 2 8
    W ($ws9.Cells.Item(2,1)) "ANALISIS DE RECLAMOS" $C.NavyHdr $C.TextWhite $true 18 -4108
    MR $ws9 3 1 3 8
    W ($ws9.Cells.Item(3,1)) "Actualizado: $($NOW.ToString('dd/MM/yyyy HH:mm'))" $C.SkyBlue $C.NavyHdr $false 10 -4108
    ColW $ws9 1 26;ColW $ws9 2 14;ColW $ws9 3 14;ColW $ws9 4 14;ColW $ws9 5 14;ColW $ws9 6 14;ColW $ws9 7 14;ColW $ws9 8 30
    $row=5
    SecHdr $ws9 $row 4 "RECLAMOS POR MES";$row++
    $recMesHdrRow=$row
    TblHdr $ws9 $row @("MES","RECLAMOS","UNIDADES RECLAMADAS","% DEL TOTAL");$row++
    $totRec=0;$totRecQty=0
    foreach($v in $recByMes.Values){$totRec+=$v.Cnt;$totRecQty+=$v.Qty}
    $pos=0
    foreach($ym in ($recByMes.Keys|Sort-Object -Descending)){
        RowH $ws9 $row 16
        $ymp=$ym.Split("-");$ymN="$($MES_NOM[[int]$ymp[1]]) $($ymp[0])"
        $bg=if($pos%2){$C.LightGray}else{$C.White}
        $pct=if($totRec){[Math]::Round($recByMes[$ym].Cnt/$totRec*100,1)}else{0}
        W ($ws9.Cells.Item($row,1)) $ymN                  $bg $C.TextDark $false 9 -4131
        W ($ws9.Cells.Item($row,2)) $recByMes[$ym].Cnt    $bg $C.TextDark $false 9 -4108
        W ($ws9.Cells.Item($row,3)) $recByMes[$ym].Qty    $bg $C.TextDark $false 9 -4108 "#,##0"
        W ($ws9.Cells.Item($row,4)) "$pct%"               $bg $C.TextDark $false 9 -4108
        $pos++;$row++
    }
    RowH $ws9 $row 18
    W ($ws9.Cells.Item($row,1)) "TOTAL" $C.SkyBlue $C.NavyHdr $true 9 -4131
    W ($ws9.Cells.Item($row,2)) $totRec    $C.SkyBlue $C.NavyHdr $true 9 -4108
    W ($ws9.Cells.Item($row,3)) $totRecQty $C.SkyBlue $C.NavyHdr $true 9 -4108 "#,##0"
    AF $ws9 $recMesHdrRow 1 4
    $row+=2
    SecHdr $ws9 $row 4 "RECLAMOS POR CATEGORIA";$row++
    $recCatHdrRow=$row
    TblHdr $ws9 $row @("CATEGORIA","RECLAMOS","UNID RECLAMADAS","% DEL TOTAL");$row++
    $pos=0
    foreach($kv in ($recByCat.GetEnumerator()|Sort-Object{$_.Value.Cnt} -Descending)){
        RowH $ws9 $row 22
        $bg=if($pos%2){$C.LightGray}else{$C.White}
        $pct=if($totRec){[Math]::Round($kv.Value.Cnt/$totRec*100,1)}else{0}
        $cell=$ws9.Cells.Item($row,1);$cell.Value2=$kv.Key;$cell.Interior.Color=$bg
        $cell.Font.Size=9;$cell.HorizontalAlignment=-4131;$cell.WrapText=$true
        W ($ws9.Cells.Item($row,2)) $kv.Value.Cnt $bg $C.TextDark $false 9 -4108
        W ($ws9.Cells.Item($row,3)) $kv.Value.Qty $bg $C.TextDark $false 9 -4108 "#,##0"
        W ($ws9.Cells.Item($row,4)) "$pct%"       $bg $C.TextDark $false 9 -4108
        $pos++;$row++
    }
    AF $ws9 $recCatHdrRow 1 4
    $row++
    SecHdr $ws9 $row 8 "RECLAMOS POR OPERARIO";$row++
    $recPikHdrRow=$row
    TblHdr $ws9 $row @("OPERARIO","RECLAMOS","UNID RECL","LINEAS TOT","TASA/1000","% TIPO FREC","TIPO MAS FRECUENTE");$row++
    $linByResp=@{}
    foreach($k in $pkMes.Keys){
        $rp=$k.Split("|")[0]
        if($allResp[$rp]){if(-not $linByResp[$rp]){$linByResp[$rp]=0};$linByResp[$rp]+=$pkMes[$k].L}
    }
    $pos=0
    foreach($kv in ($recByPik.GetEnumerator()|Sort-Object{$_.Value.Cnt} -Descending)){
        RowH $ws9 $row 16
        $bg=if($pos%2){$C.LightGray}else{$C.White}
        $totLin=if($linByResp[$kv.Key]){$linByResp[$kv.Key]}else{0}
        $rate=if($totLin){[Math]::Round($kv.Value.Cnt/($totLin/1000),2)}else{0}
        $topCat=if($kv.Value.Cats.Count){($kv.Value.Cats.GetEnumerator()|Sort-Object Value -Desc|Select-Object -First 1)}else{$null}
        $topCatN=if($topCat){$topCat.Key}else{"-"}
        $topCatP=if($topCat -and $kv.Value.Cnt){"$([Math]::Round($topCat.Value/$kv.Value.Cnt*100,0))%"}else{"-"}
        $rFg=if($kv.Value.Cnt-le 5){$C.YellowTxt}else{$C.RedTxt}
        W ($ws9.Cells.Item($row,1)) $kv.Key       $bg $C.TextDark $false 9 -4131
        W ($ws9.Cells.Item($row,2)) $kv.Value.Cnt $bg $rFg        $true  9 -4108
        W ($ws9.Cells.Item($row,3)) $kv.Value.Qty $bg $C.TextDark $false 9 -4108 "#,##0"
        W ($ws9.Cells.Item($row,4)) ([int]$totLin) $bg $C.TextDark $false 9 -4108 "#,##0"
        W ($ws9.Cells.Item($row,5)) $rate         $bg (RateFg $rate) $false 9 -4108 "0.00"
        W ($ws9.Cells.Item($row,6)) $topCatP      $bg $C.TextDark $false 9 -4108
        $cell=$ws9.Cells.Item($row,7);$cell.Value2=$topCatN;$cell.Interior.Color=$bg
        $cell.Font.Size=8;$cell.HorizontalAlignment=-4131;$cell.WrapText=$true
        $pos++;$row++
    }
    AF $ws9 $recPikHdrRow 1 7
    Write-Host "[$($NOW.ToString('HH:mm:ss'))] RECLAMOS OK"

    # ===========================================================
    # HOJA 10: MAQUINISTA
    # ===========================================================
    RowH $ws10 1 6;RowH $ws10 2 44;RowH $ws10 3 22;RowH $ws10 4 6
    MR $ws10 2 1 2 8
    W ($ws10.Cells.Item(2,1)) "  MAQUINISTAS - REABASTECIMIENTO DE DEPOSITO" $C.OrangeHdr $C.TextWhite $true 18 -4131
    MR $ws10 3 1 3 8
    W ($ws10.Cells.Item(3,1)) "  Movimientos de reabastecimiento por operario  |  $($NOW.ToString('dd/MM/yyyy HH:mm'))" $C.OrangeLight $C.OrangeHdr $false 10 -4131
    ColW $ws10 1 3;ColW $ws10 2 26;ColW $ws10 3 14;ColW $ws10 4 14;ColW $ws10 5 14;ColW $ws10 6 14;ColW $ws10 7 14;ColW $ws10 8 14
    $row=6
    if($maqNames.Count -eq 0){
        MR $ws10 $row 2 $row 8
        W ($ws10.Cells.Item($row,2)) "Sin datos de Maquinista en el archivo fuente." $C.LightGray $C.Gray $false 11 -4108
    }else{
        $maqTotByName=@{}
        foreach($key in $maqMes.Keys){ $nm=$key.Split("|")[0]; if(-not $maqTotByName[$nm]){$maqTotByName[$nm]=0}; $maqTotByName[$nm]+=$maqMes[$key].Mov }
        RowH $ws10 $row 4;$row++
        $colKPI=2
        foreach($nm in ($maqNames.Keys|Sort-Object)){
            $r1=$row
            RowH $ws10 $row 18;MR $ws10 $row $colKPI $row ($colKPI+2)
            W ($ws10.Cells.Item($row,$colKPI)) $nm $C.OrangeLight $C.OrangeHdr $true 10 -4108
            $row++;RowH $ws10 $row 32;MR $ws10 $row $colKPI $row ($colKPI+2)
            W ($ws10.Cells.Item($row,$colKPI)) $maqTotByName[$nm] $C.OrangeLight $C.OrangeHdr $true 20 -4108 "#,##0"
            $row=$r1; $colKPI+=3
        }
        RowH $ws10 ($row+1) 34; $row+=2;RowH $ws10 $row 6;$row++
        MR $ws10 $row 2 $row 8
        W ($ws10.Cells.Item($row,2)) "  MOVIMIENTOS POR MES Y OPERARIO" $C.OrangeHdr $C.TextWhite $true 12 -4131;$row++
        $maqHdrRow=$row; RowH $ws10 $row 20
        @("MES","DIAS ACTIVO")+($maqNames.Keys|Sort-Object)|
            ForEach-Object -Begin{$c=2} -Process{W ($ws10.Cells.Item($row,$c)) $_ $C.MidGray $C.NavyHdr $true 9 -4108;$c++}
        $row++
        $maqAllMons=@{}
        foreach($k in $maqMes.Keys){$maqAllMons[$k.Split("|")[1]]=$true}
        $pos=0
        foreach($ym in ($maqAllMons.Keys|Sort-Object -Descending)){
            RowH $ws10 $row 17
            $ymp=$ym.Split("-");$ymN="$($MES_NOM[[int]$ymp[1]]) $($ymp[0])"
            $bg=if($pos%2){$C.OrangeLight}else{$C.White}
            W ($ws10.Cells.Item($row,2)) $ymN $bg $C.TextDark $false 9 -4131
            $cIdx=4
            foreach($nm in ($maqNames.Keys|Sort-Object)){
                $mk="$nm|$ym"
                $mov=if($maqMes[$mk]){$maqMes[$mk].Mov}else{0}
                $dias=if($maqMes[$mk]){$maqMes[$mk].Days.Count}else{0}
                $movBg=if($mov-gt 100){$C.GreenBg}elseif($mov-gt 30){$C.YellowBg}elseif($mov-gt 0){$C.RedBg}else{$bg}
                $movFg=if($mov-gt 100){$C.GreenFg}elseif($mov-gt 30){$C.YellowFg}elseif($mov-gt 0){$C.RedFg}else{$C.Gray}
                if($cIdx -eq 4){W ($ws10.Cells.Item($row,3)) $dias $bg $C.TextDark $false 9 -4108}
                W ($ws10.Cells.Item($row,$cIdx)) $mov $movBg $movFg ($mov-gt 0) 9 -4108 "#,##0"
                $cIdx++
            }
            $pos++;$row++
        }
        AF $ws10 $maqHdrRow 2 ($maqNames.Count+3)
    }
    Write-Host "[$($NOW.ToString('HH:mm:ss'))] MAQUINISTA OK"

    # ===========================================================
    # HOJA 11: CONTROL
    # ===========================================================
    MR $ws11 2 1 2 5
    if($ctrlHasData){
        W ($ws11.Cells.Item(2,1)) "CONTROL DE MERCADERIA" $C.NavyHdr $C.TextWhite $true 18 -4108
        W ($ws11.Cells.Item(4,1)) "Los datos de Control se procesaran en la proxima actualizacion." $C.PaleBlue $C.NavyHdr $false 11 -4131
    }else{
        W ($ws11.Cells.Item(2,1)) "CONTROL - SIN DATOS AUN" $C.MidGray $C.Gray $true 18 -4108
        MR $ws11 4 1 4 5
        W ($ws11.Cells.Item(4,1)) "Cuando la pestana Control tenga datos, esta hoja se actualizara automaticamente." $C.LightGray $C.Gray $false 11 -4131
    }

    # ===========================================================
    # GENERAR HTML DASHBOARD - Escala de grises, todo dinamico
    # ===========================================================
    $htmlOut = [IO.Path]::ChangeExtension($Output,".html")
    $IC = [System.Globalization.CultureInfo]::InvariantCulture

    # ---- Construir datos para charts y monthlyData JS ----
    $allYears = @{}
    # Leer pickeadores oficiales por mes desde estructura ARG - CHI.xlsx (col N)
    $pickeadoresByMon = @{}
    $estructuraPath = "C:\Users\bchevasco\OneDrive - Articulos Promocionales SA\Escritorio\Inteligencia Artificial\Personal\estructura ARG - CHI.xlsx"
    if(Test-Path $estructuraPath){
        try{
            $xtWb = $xl.Workbooks.Open($estructuraPath)
            # Buscar hoja "ARG" (case-insensitive); si no existe, usar la primera hoja
            $xtWs = $null
            foreach($sh in $xtWb.Sheets){ if("$($sh.Name)".Trim() -ieq "ARG"){ $xtWs=$sh; break } }
            if(-not $xtWs){ $xtWs = $xtWb.Sheets.Item(1) }
            $xtArr = $xtWs.UsedRange.Value2
            $xtRows = $xtArr.GetUpperBound(0)
            $xtCols = $xtArr.GetUpperBound(1)
            $pkCol=$null; $mesCol=$null
            for($c=1;$c -le $xtCols;$c++){
                $h="$($xtArr[1,$c])".Trim()
                if($h -like "*Pickeadores*"){ $pkCol=$c }
                if($h -ieq "Mes"){ $mesCol=$c }
            }
            # Diccionario de abreviaturas de meses en español para parsear texto tipo "ene-25"
            $MESES_ABR=@{"ene"=1;"feb"=2;"mar"=3;"abr"=4;"may"=5;"jun"=6;"jul"=7;"ago"=8;"sep"=9;"oct"=10;"nov"=11;"dic"=12}
            if($pkCol -and $mesCol){
                for($r=2;$r -le $xtRows;$r++){
                    $mesVal=$xtArr[$r,$mesCol]; $pkVal=$xtArr[$r,$pkCol]
                    if(-not $mesVal -or -not $pkVal){ continue }
                    $ymx=$null
                    # Opción 1: valor numérico OA date (fecha Excel normal)
                    if($mesVal -is [double] -or $mesVal -is [int]){
                        try{
                            $d=[datetime]::FromOADate([double]$mesVal)
                            $ymx="$($d.Year)-$('{0:00}' -f $d.Month)"
                        }catch{}
                    }
                    # Opción 2: texto "ene-25" / "ene-2025"
                    if(-not $ymx){
                        $ms="$mesVal".Trim().ToLower()
                        if($ms -match "^([a-z]{3})-(\d{2,4})$"){
                            $mn=$MESES_ABR[$matches[1]]
                            if($mn){
                                $fy=if($matches[2].Length -eq 2){"20$($matches[2])"}else{$matches[2]}
                                $ymx="$fy-$('{0:00}' -f $mn)"
                            }
                        }
                    }
                    if(-not $ymx){ continue }
                    try{
                        $pkInt=[int]$pkVal
                        if($pkInt -gt 0 -and (-not $pickeadoresByMon[$ymx] -or $pickeadoresByMon[$ymx] -lt $pkInt)){
                            $pickeadoresByMon[$ymx]=$pkInt
                        }
                    }catch{}
                }
            } else {
                Write-Host "  [WARN] Columna no encontrada — Mes:$mesCol Pickeadores:$pkCol (hoja '$($xtWs.Name)')"
            }
            $xtWb.Close($false)
            $pkLog=($pickeadoresByMon.GetEnumerator()|Sort-Object Key|ForEach-Object{"$($_.Key)=$($_.Value)"}) -join ", "
            Write-Host "[$($NOW.ToString('HH:mm:ss'))] Pickeadores OK: $($pickeadoresByMon.Count) meses -> $pkLog"
        }catch{ Write-Host "  [WARN] Error leyendo estructura ARG-CHI: $_" }
    } else { Write-Host "  [WARN] No se encontró: $estructuraPath" }

    $jsMonthlyDataParts = [System.Collections.Generic.List[string]]::new()
    $jsAllDataList = [System.Collections.Generic.List[string]]::new()
    $jsMonLabelsList=[System.Collections.Generic.List[string]]::new()
    $jsTeamLDList   =[System.Collections.Generic.List[string]]::new()
    $jsTeamLinList  =[System.Collections.Generic.List[string]]::new()
    $jsTeamColors   =[System.Collections.Generic.List[string]]::new()
    $jsStaffNecList =[System.Collections.Generic.List[string]]::new()
    $jsStaffActList =[System.Collections.Generic.List[string]]::new()
    $jsMerma1List   =[System.Collections.Generic.List[string]]::new()
    $jsMerma2List   =[System.Collections.Generic.List[string]]::new()
    $jsMerma1LList  =[System.Collections.Generic.List[string]]::new()
    $jsMerma2LList  =[System.Collections.Generic.List[string]]::new()
    $jsPie1List     =[System.Collections.Generic.List[string]]::new()
    $jsPie2List     =[System.Collections.Generic.List[string]]::new()

    foreach($mon in $sortedMon){
        $mp=$mon.Split("-")
        $allYears[$mp[0]]=$true
        $jsMonLabelsList.Add("'$($MES_NOM[[int]$mp[1]]) $($mp[0])'")
        $mR=@($sumRows|Where-Object{$_.YM -eq $mon})
        if($mR.Count){
            $mRProd=@($mR|Where-Object{-not (IsExcludedFromProd $_.Resp)})
            $mWkDays=if($staffByMes[$mon]){[int]$staffByMes[$mon].WorkDays.Count}else{1}
            $mWkDays=if($mWkDays -gt 0){$mWkDays}else{1}
            $mProdTotL=($mRProd|Measure-Object Lineas -Sum).Sum
            $mPick=if($pickeadoresByMon[$mon]){$pickeadoresByMon[$mon]}else{[math]::Max($mRProd.Count,1)}
            $tld=[Math]::Round($mProdTotL/$mWkDays/$mPick,1)
            $jsTeamLDList.Add($tld.ToString($IC))
            $jsTeamLinList.Add(($mR|Measure-Object Lineas -Sum).Sum.ToString($IC))
            if($tld -ge $TARGET){$jsTeamColors.Add("'#16a34a'")}
            elseif($tld -ge 70){$jsTeamColors.Add("'#d97706'")}
            else{$jsTeamColors.Add("'#dc2626'")}
        }else{$jsTeamLDList.Add("0");$jsTeamLinList.Add("0");$jsTeamColors.Add("'#dc2626'")}
        $sr2=$staffRows|Where-Object{$_.YM -eq $mon}|Select-Object -First 1
        $sr2Pick=if($pickeadoresByMon[$mon]){$pickeadoresByMon[$mon]}else{if($sr2){$sr2.PersonasActuales}else{0}}
        if($sr2){$jsStaffNecList.Add($sr2.PersonasNecesarias.ToString($IC));$jsStaffActList.Add($sr2Pick.ToString($IC))}
        else{$jsStaffNecList.Add("0");$jsStaffActList.Add("0")}
        $mr2=$mermaRows|Where-Object{$_.YM -eq $mon}|Select-Object -First 1
        if($mr2){
            $jsMerma1List.Add($mr2.M1LD.ToString($IC));$jsMerma2List.Add($mr2.M2LD.ToString($IC))
            $jsMerma1LList.Add($mr2.M1Lineas.ToString($IC));$jsMerma2LList.Add($mr2.M2Lineas.ToString($IC))
        } else {
            $jsMerma1List.Add("null");$jsMerma2List.Add("null")
            $jsMerma1LList.Add("null");$jsMerma2LList.Add("null")
        }
        $pr2=$pieRows|Where-Object{$_.YM -eq $mon}|Select-Object -First 1
        if($pr2){$jsPie1List.Add($pr2.M1LD.ToString($IC));$jsPie2List.Add($pr2.M2LD.ToString($IC))}
        else{$jsPie1List.Add("null");$jsPie2List.Add("null")}
        $jsAllDataList.Add("'$mon'")
    }

    # ---- monthlyData JS: datos completos por mes para KPIs y ranking dinamico ----
    foreach($mon in $sortedMon){
        $mRows=@($sumRows|Where-Object{$_.YM -eq $mon}|Sort-Object LineasDia -Descending)
        $mTotL=0;$mTotU=0;$mTotRC=0;$mTotOlas=0
        foreach($op in $mRows){$mTotL+=$op.Lineas;$mTotU+=$op.Unidades;$mTotRC+=$op.RecCnt;$mTotOlas+=$op.Olas}
        # Excluir extras del cálculo de productividad del equipo
        $mProdRows=@($mRows|Where-Object{-not (IsExcludedFromProd $_.Resp)})
        $mTotOps=$mProdRows.Count
        # Usar pickeadores oficiales del archivo de estructura como denominador real
        $mPickOficial=if($pickeadoresByMon[$mon]){$pickeadoresByMon[$mon]}else{[math]::Max($mTotOps,1)}
        $mWDays=if($staffByMes[$mon]){[int]$staffByMes[$mon].WorkDays.Count}else{1}
        $mWDays=if($mWDays -gt 0){$mWDays}else{1}
        $mProdTotalL=($mProdRows|Measure-Object Lineas -Sum).Sum
        $mAvgLD=[Math]::Round($mProdTotalL/$mWDays/$mPickOficial,1)
        $mCumAv=[Math]::Round($mAvgLD/$TARGET*100,1)
        $mRateG=if($mTotL){[Math]::Round($mTotRC/($mTotL/1000),2)}else{0}
        $sr=$staffRows|Where-Object{$_.YM -eq $mon}|Select-Object -First 1
        $mStaffNec=if($sr){$sr.PersonasNecesarias}else{0}
        $mStaffAct=$mPickOficial
        $pickerParts=[System.Collections.Generic.List[string]]::new()
        foreach($op in $mProdRows){
            $rn=$op.Resp -replace "'",""
            $tr=if($trendByResp[$op.Resp]){($trendByResp[$op.Resp].Txt -replace "'","")}else{"-"}
            $pickerParts.Add("{resp:'$rn',dias:$($op.Dias),olas:$($op.Olas),lineas:$($op.Lineas),ld:$($op.LineasDia.ToString($IC)),cumpl:$($op.Cumplim.ToString($IC)),unidades:$($op.Unidades),recCnt:$($op.RecCnt),trend:'$tr'}")
        }
        $pickersArr="["+($pickerParts -join ",")+"]"
        $jsMonthlyDataParts.Add("'$mon':{totOps:$mTotOps,pickeadores:$mPickOficial,avgLD:$($mAvgLD.ToString($IC)),cumAv:$($mCumAv.ToString($IC)),totL:$mTotL,totU:$mTotU,totRC:$mTotRC,rateG:$($mRateG.ToString($IC)),totOlas:$mTotOlas,staffNec:$($mStaffNec.ToString($IC)),staffAct:$($mStaffAct.ToString($IC)),pickers:$pickersArr}")
    }
    $jsMonthlyData = "{" + ($jsMonthlyDataParts -join ",") + "}"

    $jsMonLabels  = $jsMonLabelsList -join ","
    $jsTeamLD     = $jsTeamLDList    -join ","
    $jsTeamPtClrs = $jsTeamColors    -join ","
    $jsStaffNec   = $jsStaffNecList  -join ","
    $jsStaffAct   = $jsStaffActList  -join ","
    $jsMerma1     = $jsMerma1List    -join ","
    $jsMerma2     = $jsMerma2List    -join ","
    $jsMerma1L    = $jsMerma1LList   -join ","
    $jsMerma2L    = $jsMerma2LList   -join ","
    $jsPie1       = $jsPie1List      -join ","
    $jsPie2       = $jsPie2List      -join ","
    $nMons        = $sortedMon.Count
    $jsAllMons    = $jsAllDataList   -join ","

    # Anos disponibles para el filtro
    $jsYearsList=[System.Collections.Generic.List[string]]::new()
    foreach($y in ($allYears.Keys|Sort-Object -Descending)){ $jsYearsList.Add("'$y'") }
    $jsYears = $jsYearsList -join ","

    # Mes inicial del filtro (ultimo mes disponible)
    $latestYMP = $latestYM.Split("-")
    $jsInitYear = $latestYMP[0]
    $jsInitMon  = [int]$latestYMP[1]

    # JSON para Control
    $ctrlYMParts=[System.Collections.Generic.List[string]]::new()
    $ctrlByYM=@{}
    foreach($key in $ctrlMes.Keys){
        $pts=$key.Split("|"); $nm=$pts[0]; $ym=$pts[1]
        if(-not $ctrlByYM[$ym]){$ctrlByYM[$ym]=[System.Collections.Generic.List[string]]::new()}
        $mc=$ctrlMes[$key]; $dias=$mc.Days.Count
        $ctrlByYM[$ym].Add("{nm:'$nm',und:$([int]$mc.Und),ord:$([int]$mc.Ord),dias:$dias}")
    }
    foreach($ym in ($ctrlByYM.Keys|Sort-Object)){
        $arr="["+($ctrlByYM[$ym] -join ",")+"]"
        $ctrlYMParts.Add("'$ym':$arr")
    }
    $jsCtrlData="{"+($ctrlYMParts -join ",")+"}"

    # JSON para Maquinistas
    $maqYMParts=[System.Collections.Generic.List[string]]::new()
    $maqByYM=@{}
    foreach($key in $maqMes.Keys){
        $pts=$key.Split("|"); $nm=$pts[0]; $ym=$pts[1]
        if(-not $maqByYM[$ym]){$maqByYM[$ym]=[System.Collections.Generic.List[string]]::new()}
        $mm=$maqMes[$key]; $dias=$mm.Days.Count
        $maqByYM[$ym].Add("{nm:'$nm',mov:$([int]$mm.Mov),und:$([int]$mm.Und),dias:$dias,pa02:$([int]$mm.Pa02),rec:$([int]$mm.Rec),rl01:$([int]$mm.Rl01),pallet:$([int]$mm.Pallet),flow:$([int]$mm.Flow)}")
    }
    foreach($ym in ($maqByYM.Keys|Sort-Object)){
        $ymP=$ym.Split("-"); $lbl="$($MES_NOM[[int]$ymP[1]]) $($ymP[0])"
        $ops="["+($maqByYM[$ym] -join ",")+"]"
        $maqYMParts.Add("'$ym':{lbl:'$lbl',ops:$ops}")
    }
    $jsMaqDetailData="{"+($maqYMParts -join ",")+"}"

    # JS evolucion individual por picker
    $pickersJs=""
    $pickerColors=@('#2563eb','#16a34a','#d97706','#dc2626','#7c3aed','#0891b2','#be185d','#ca8a04','#15803d','#b45309')
    $ci=0
    foreach($resp in $sortedResp){
        $valList=[System.Collections.Generic.List[string]]::new()
        foreach($mon in $sortedMon){
            $k="$resp|$mon"
            if($pkMes[$k]){
                $dd=$pkMes[$k].Days.Count
                if($dd){$valList.Add([Math]::Round($pkMes[$k].L/$dd,1).ToString($IC))}else{$valList.Add("0")}
            }else{$valList.Add("null")}
        }
        $vals=$valList -join ","
        $col=$pickerColors[$ci % $pickerColors.Count]
        $pickersJs+="  {label:'$resp',data:[$vals],borderColor:'$col',backgroundColor:'${col}22',borderWidth:2,tension:.3,pointRadius:3,fill:false,spanGaps:true},`n"
        $ci++
    }

    # Muestra Simple (Lezcano) JS monthly data
    $jsLezLDList  =[System.Collections.Generic.List[string]]::new()
    $jsLezMonParts=[System.Collections.Generic.List[string]]::new()
    foreach($mon in $sortedMon){
        $lr=$lezRows|Where-Object{$_.YM -eq $mon}|Select-Object -First 1
        if($lr){
            $jsLezLDList.Add($lr.LineasDia.ToString($IC))
            $vsT=[Math]::Round($lr.LineasDia - $lr.TeamAvg, 1)
            $jsLezMonParts.Add("'$mon':{ld:$($lr.LineasDia.ToString($IC)),lineas:$($lr.Lineas),dias:$($lr.Dias),olas:$($lr.Olas),unidades:$($lr.Unidades),recCnt:$($lr.RecCnt),vsTeam:$($vsT.ToString($IC))}")
        } else {
            $jsLezLDList.Add("null")
            $jsLezMonParts.Add("'$mon':null")
        }
    }
    $jsLezLD      = $jsLezLDList -join ","
    $jsLezMonthly = "{"+($jsLezMonParts -join ",")+"}"

    # Pie de Maquina: lineas totales M1/M2 por mes
    $jsPieM1LList=[System.Collections.Generic.List[string]]::new()
    $jsPieM2LList=[System.Collections.Generic.List[string]]::new()
    foreach($mon in $sortedMon){
        $pr=$pieRows|Where-Object{$_.YM -eq $mon}|Select-Object -First 1
        if($pr){ $jsPieM1LList.Add($pr.M1Lineas.ToString($IC));$jsPieM2LList.Add($pr.M2Lineas.ToString($IC)) }
        else{ $jsPieM1LList.Add("null");$jsPieM2LList.Add("null") }
    }
    $jsPieM1L = $jsPieM1LList -join ","
    $jsPieM2L = $jsPieM2LList -join ","

    # Operarios para filtro
    $jsOpList=[System.Collections.Generic.List[string]]::new()
    foreach($resp in $sortedResp){ $jsOpList.Add("'" + ($resp -replace "'","") + "'") }
    $jsOperarios = $jsOpList -join ","

    # Reclamos: objeto mensual con ByOp y ByCat para dashboard dinamico
    $jsRecMonParts=[System.Collections.Generic.List[string]]::new()
    foreach($mon in $sortedMon){
        $d=$recDataByMonth[$mon]
        if($d){
            $opParts=[System.Collections.Generic.List[string]]::new()
            foreach($op in ($d.ByOp.Keys|Sort-Object)){
                $opEsc=$op -replace "'","" -replace '"',""
                $opParts.Add("'$opEsc':$($d.ByOp[$op])")
            }
            $catParts=[System.Collections.Generic.List[string]]::new()
            foreach($cat in ($d.ByCat.Keys|Sort-Object)){
                $catEsc=$cat -replace "'","" -replace '"',""
                $catParts.Add("'$catEsc':$($d.ByCat[$cat])")
            }
            $byOpStr  ="{"+($opParts  -join ",")+"}"
            $byCatStr ="{"+($catParts -join ",")+"}"
            $byOpCatParts=[System.Collections.Generic.List[string]]::new()
            foreach($op2 in ($d.ByOpCat.Keys|Sort-Object)){
                $op2Esc=$op2 -replace "'","" -replace '"',""
                $catSubParts=[System.Collections.Generic.List[string]]::new()
                foreach($cat2 in ($d.ByOpCat[$op2].Keys|Sort-Object)){
                    $cat2Esc=$cat2 -replace "'","" -replace '"',""
                    $catSubParts.Add("'$cat2Esc':$($d.ByOpCat[$op2][$cat2])")
                }
                $byOpCatParts.Add("'$op2Esc':{"+($catSubParts -join ",")+"}")
            }
            $byOpCatStr="{"+($byOpCatParts -join ",")+"}"
            $jsRecMonParts.Add("'$mon':{cnt:$($d.Cnt),byOp:$byOpStr,byCat:$byCatStr,byOpCat:$byOpCatStr}")
        } else {
            $jsRecMonParts.Add("'$mon':null")
        }
    }
    $jsRecMonthly="{"+($jsRecMonParts -join ",")+"}"

    # --- Serializar datos para HTML: tabla 7 dias y grafico evolucion 30 dias ---
    $jsDay7Parts=[System.Collections.Generic.List[string]]::new()
    foreach($r7 in $day7Rows){
        $dlt7=if($null -eq $r7.Delta){"null"}else{[string][int]$r7.Delta}
        $jsDay7Parts.Add("{f:'$($r7.Fecha)',d:'$($r7.DiaSem)',ops:$($r7.Ops),lines:$($r7.Lines),lpo:$($r7.LinesPerOp),tgt:$($r7.Target),cum:$($r7.Cumpl),dlt:$dlt7,cl:$($r7.CLLines),clops:$($r7.CLOps),sl:$($r7.SLLines),slops:$($r7.SLOps)}")
    }
    $jsDay7Rows="["+($jsDay7Parts -join ",")+"]"

    $evol30Dates=@($pkDay.Keys|Where-Object{$allResp[$_.Split("|")[0]]}|ForEach-Object{$_.Split("|")[1]}|Sort-Object -Unique|Select-Object -Last 30)
    $evol30LblParts=[System.Collections.Generic.List[string]]::new()
    $evol30DataParts=[System.Collections.Generic.List[string]]::new()
    $evol30TgtParts=[System.Collections.Generic.List[string]]::new()
    foreach($e30d in $evol30Dates){
        $e30T=0
        foreach($rsp in $sortedResp){$e30k="$rsp|$e30d";if($pkDay[$e30k]){$e30T+=$pkDay[$e30k]}}
        $evol30LblParts.Add("'$([datetime]::Parse($e30d).ToString('dd/MM'))'")
        $evol30DataParts.Add("$([int]$e30T)")
        $e30ym=$e30d.Substring(0,7)
        $e30pick=if($pickeadoresByMon[$e30ym]){$pickeadoresByMon[$e30ym]}else{4}
        $evol30TgtParts.Add("$($e30pick*$TARGET)")
    }
    $jsEvol30Labels="["+($evol30LblParts -join ",")+"]"
    $jsEvol30Data="["+($evol30DataParts -join ",")+"]"
    # Datos por grupo (Con Logo / Sin Logo) para filtros
    $evol30CLParts=[System.Collections.Generic.List[string]]::new()
    $evol30SLParts=[System.Collections.Generic.List[string]]::new()
    foreach($e30d in $evol30Dates){
        $e30gd=$dayGroupData[$e30d]
        $evol30CLParts.Add($(if($e30gd){"$([int]$e30gd.CL)"}else{"0"}))
        $evol30SLParts.Add($(if($e30gd){"$([int]$e30gd.SL)"}else{"0"}))
    }
    $jsEvol30CL="["+($evol30CLParts -join ",")+"]"
    $jsEvol30SL="["+($evol30SLParts -join ",")+"]"
    $jsEvol30Target="["+($evol30TgtParts -join ",")+"]"

    $html = @"
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Dashboard Productividad | Zecat ARG</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<style>
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:Arial,sans-serif;background:#f0f0f0;color:#111111}
header{background:#111111;color:white;padding:18px 32px;display:flex;align-items:center;justify-content:space-between}
header .header-left{display:flex;align-items:center;gap:16px}
header h1{font-size:22px;font-weight:700;letter-spacing:.5px}
header .subtitle{font-size:13px;opacity:.55;margin-top:3px}
header .badge{background:rgba(255,255,255,.1);border-radius:20px;padding:6px 16px;font-size:12px;text-align:right;line-height:1.6}
.zecat-logo{width:46px;height:46px;flex-shrink:0}
.tab-nav{background:#1a1a1a;display:flex;gap:0;padding:0 24px}
.tab-btn{background:transparent;color:#888888;border:none;padding:13px 22px;font-size:13px;font-weight:600;cursor:pointer;border-bottom:3px solid transparent;transition:all .2s;white-space:nowrap}
.tab-btn:hover{color:#cccccc;background:rgba(255,255,255,.05)}
.tab-btn.active{color:white;border-bottom-color:#2563eb}
.sec{display:none}.sec.active{display:block}
.container{max-width:1440px;margin:0 auto;padding:20px 24px}
.filter-bar{background:white;border-radius:10px;padding:12px 20px;box-shadow:0 1px 4px rgba(0,0,0,.08);margin-bottom:18px;display:flex;align-items:center;gap:20px;flex-wrap:wrap;border:1px solid #e0e0e0}
.filter-bar label{font-size:11px;font-weight:600;color:#555555;text-transform:uppercase;letter-spacing:.5px}
.filter-bar select{border:1.5px solid #cccccc;border-radius:8px;padding:5px 10px;font-size:13px;color:#111111;background:white;cursor:pointer}
.filter-bar select:focus{outline:none;border-color:#444444}
.filter-badge{background:#f5f5f5;border:1px solid #cccccc;border-radius:20px;padding:4px 14px;font-size:12px;color:#333333;font-weight:600}
.kpi-grid{display:grid;gap:14px;margin-bottom:20px}
.g5{grid-template-columns:repeat(5,1fr)}.g4{grid-template-columns:repeat(4,1fr)}.g3{grid-template-columns:repeat(3,1fr)}.g2{grid-template-columns:repeat(2,1fr)}
.kpi-card{background:white;border-radius:12px;padding:18px 20px;box-shadow:0 1px 3px rgba(0,0,0,.05);border:1px solid #ebebeb;border-top:3px solid #ccc;transition:transform .15s;position:relative}
.kpi-card:hover{transform:translateY(-2px);box-shadow:0 4px 12px rgba(0,0,0,.09)}
.kpi-card.blue{border-top-color:#2563eb;background:#f0f7ff}.kpi-card.green{border-top-color:#16a34a;background:#f0fdf4}
.kpi-card.red{border-top-color:#dc2626;background:#fff5f5}.kpi-card.amber{border-top-color:#d97706;background:#fffcf0}
.kpi-card.purple{border-top-color:#7c3aed;background:#f5f3ff}.kpi-card.teal{border-top-color:#0891b2;background:#f0fdff}
.kpi-card.slate{border-top-color:#64748b}.kpi-card.gray{border-top-color:#9ca3af}
.kpi-label{font-size:11px;text-transform:uppercase;letter-spacing:.7px;color:#6b7280;font-weight:600;margin-bottom:10px}
.kpi-value{font-size:30px;font-weight:700;color:#0f172a;line-height:1}
.kpi-sub{font-size:12px;color:#9ca3af;margin-top:7px}
.charts-row{display:grid;gap:18px;margin-bottom:18px}
.c2{grid-template-columns:1fr 1fr}.c3{grid-template-columns:1.2fr 1fr 1fr}.c1{grid-template-columns:1fr}
.chart-card{background:white;border-radius:10px;padding:20px;box-shadow:0 1px 4px rgba(0,0,0,.08);border:1px solid #e8e8e8}
.chart-title{font-size:14px;font-weight:700;color:#111111;margin-bottom:3px}
.chart-subtitle{font-size:11px;color:#999999;margin-bottom:16px}
.table-card{background:white;border-radius:10px;padding:20px;box-shadow:0 1px 4px rgba(0,0,0,.08);margin-bottom:18px;border:1px solid #e8e8e8}
.rank-title{font-size:14px;font-weight:700;color:#111111;margin-bottom:2px}
.rank-sub{font-size:11px;color:#999999;margin-bottom:14px}
table{width:100%;border-collapse:collapse;font-size:13px}
thead tr{background:#f5f5f5}
thead th{padding:9px 11px;text-align:left;font-size:11px;text-transform:uppercase;letter-spacing:.5px;color:#777777;font-weight:600;border-bottom:2px solid #e0e0e0}
tbody tr{border-bottom:1px solid #f0f0f0;transition:background .15s}
tbody tr:hover{background:#f8f8f8}
tbody td{padding:9px 11px;color:#333333}
.info-box{background:white;border:1px solid #e0e0e0;border-radius:8px;padding:10px 16px;font-size:13px;color:#555555;margin-bottom:12px}
.nodata{background:white;border-radius:10px;padding:48px;text-align:center;color:#aaaaaa;font-size:15px;border:1px solid #e8e8e8}
footer{text-align:center;padding:16px;font-size:11px;color:#aaaaaa}
@media(max-width:1100px){.g5{grid-template-columns:repeat(3,1fr)}.c3{grid-template-columns:1fr 1fr}}
@media(max-width:700px){.g5,.g4,.g3{grid-template-columns:1fr 1fr}.c2,.c3,.g2{grid-template-columns:1fr}}
/* ===== DARK MODE ===== */
body.dark{background:#0f172a;color:#f1f5f9}
body.dark .tab-nav{background:#020617}
body.dark .tab-btn{color:#64748b}body.dark .tab-btn:hover{color:#94a3b8;background:rgba(255,255,255,.04)}body.dark .tab-btn.active{color:#f1f5f9;border-bottom-color:#3b82f6}
body.dark .filter-bar{background:#1e293b;border-color:#334155}
body.dark .filter-bar label{color:#94a3b8}
body.dark .filter-bar select{background:#0f172a;color:#e2e8f0;border-color:#475569}
body.dark .filter-badge{background:#0f172a;border-color:#334155;color:#cbd5e1}
body.dark .kpi-card{background:#1e293b;border-color:#334155;border-top-color:inherit;box-shadow:none}
body.dark .kpi-card.blue{background:#0f1e35}body.dark .kpi-card.green{background:#0a1f15}body.dark .kpi-card.amber{background:#1f1505}body.dark .kpi-card.teal{background:#071820}body.dark .kpi-card.red{background:#1f0a0a}body.dark .kpi-card.purple{background:#130d28}
body.dark .kpi-label{color:#94a3b8}body.dark .kpi-value{color:#f1f5f9}body.dark .kpi-sub{color:#64748b}
body.dark .chart-card{background:#1e293b;border-color:#334155;box-shadow:none}
body.dark .table-card{background:#1e293b;border-color:#334155;box-shadow:none}
body.dark .chart-title{color:#f1f5f9}body.dark .chart-subtitle{color:#64748b}
body.dark .rank-title{color:#f1f5f9}body.dark .rank-sub{color:#64748b}
body.dark thead tr{background:#0f172a}
body.dark thead th{color:#94a3b8;border-bottom-color:#334155}
body.dark tbody tr{border-bottom-color:#1e293b}
body.dark tbody tr:hover{background:#0f172a}
body.dark tbody td{color:#e2e8f0}
body.dark .nodata{background:#1e293b;border-color:#334155;color:#64748b}
body.dark footer{color:#475569;background:#020617}
body.dark .info-box{background:#1e293b;border-color:#334155;color:#94a3b8}
tfoot td{border-top:2px solid #e0e0e0;font-weight:700}
body.dark tfoot tr{background:#0f172a}
body.dark tfoot td{color:#e2e8f0;border-top-color:#334155}
.grp-btns{display:flex;gap:8px;margin-bottom:14px;flex-wrap:wrap}
.grp-btn{border:1.5px solid #cccccc;border-radius:20px;padding:5px 16px;font-size:12px;font-weight:600;cursor:pointer;background:white;color:#555555;transition:all .18s}
.grp-btn:hover{border-color:#2563eb;color:#2563eb}
.grp-btn.active{background:#2563eb;border-color:#2563eb;color:white}
body.dark .grp-btn{background:#1e293b;border-color:#334155;color:#94a3b8}
body.dark .grp-btn:hover{border-color:#3b82f6;color:#60a5fa}
body.dark .grp-btn.active{background:#2563eb;border-color:#2563eb;color:white}
#themeToggle{background:rgba(255,255,255,.12);border:1px solid rgba(255,255,255,.18);border-radius:50px;width:42px;height:42px;color:white;cursor:pointer;font-size:18px;display:flex;align-items:center;justify-content:center;transition:all .2s;flex-shrink:0}
#themeToggle:hover{background:rgba(255,255,255,.22);transform:scale(1.08)}
</style>
</head>
<body>
<header>
  <div class="header-left">
    <svg class="zecat-logo" viewBox="0 0 100 100" xmlns="http://www.w3.org/2000/svg">
      <circle cx="50" cy="50" r="50" fill="#F05A47"/>
      <text x="50" y="68" font-family="Arial,sans-serif" font-size="58" font-weight="900" fill="white" text-anchor="middle" letter-spacing="-2">Z</text>
    </svg>
    <div>
      <div class="subtitle">ZECAT &mdash; Art&iacute;culos Promocionales SA</div>
      <h1>Dashboard de Productividad</h1>
    </div>
  </div>
  <div class="badge">Actualizado: $($NOW.ToString('dd/MM/yyyy HH:mm'))<br>Target picking: $TARGET lin/d&iacute;a</div>
  <button id="themeToggle" onclick="toggleTheme()" title="Cambiar tema claro / oscuro" aria-label="Cambiar tema"><span id="themeIcon">&#9790;</span></button>
</header>

<!-- NAVEGACION POR TABS -->
<nav class="tab-nav">
  <button class="tab-btn active" onclick="switchTab('picking')" id="btn-picking">&#128230; Productividad Picking</button>
  <button class="tab-btn" onclick="switchTab('pie')" id="btn-pie">&#9881;&#65039; Pie de M&aacute;quina</button>
  <button class="tab-btn" onclick="switchTab('muestra')" id="btn-muestra">&#128203; Muestra Simple</button>
  <button class="tab-btn" onclick="switchTab('reclamos')" id="btn-reclamos">&#128683; Reclamos</button>
  <button class="tab-btn" onclick="switchTab('control')" id="btn-control">&#128269; Control</button>
  <button class="tab-btn" onclick="switchTab('maquinistas')" id="btn-maquinistas">&#128295; Maquinistas</button>
  <button class="tab-btn" onclick="switchTab('resumen')" id="btn-resumen">&#128202; Resumen General</button>
  <button class="tab-btn" onclick="switchTab('eficiencia')" id="btn-eficiencia">&#127942; Bonos</button>
</nav>

<div class="container">

<!-- FILTROS COMPARTIDOS -->
<div class="filter-bar">
  <label>Filtrar:</label>
  <div><label for="selAnio">A&ntilde;o</label>&nbsp;
    <select id="selAnio" onchange="applyFilter()">
      <option value="all">Todos</option>
    </select>
  </div>
  <div><label for="selMes">Mes</label>&nbsp;
    <select id="selMes" onchange="applyFilter()">
      <option value="all">Todos</option>
      <option value="1">Enero</option><option value="2">Febrero</option><option value="3">Marzo</option>
      <option value="4">Abril</option><option value="5">Mayo</option><option value="6">Junio</option>
      <option value="7">Julio</option><option value="8">Agosto</option><option value="9">Septiembre</option>
      <option value="10">Octubre</option><option value="11">Noviembre</option><option value="12">Diciembre</option>
    </select>
  </div>
  <div id="opFilterWrap"><label for="selOp">Operario</label>&nbsp;
    <select id="selOp" onchange="applyFilter()">
      <option value="all">Todos</option>
    </select>
  </div>
  <span class="filter-badge" id="filterBadge">Cargando...</span>
  <button onclick="resetFilter()" style="border:1px solid #cccccc;border-radius:8px;padding:5px 12px;font-size:12px;cursor:pointer;background:white;color:#333">&#10006; Resetear</button>
</div>

<!-- ===== SECCION: PICKING ===== -->
<div id="sec-picking" class="sec active">

<div class="grp-btns">
  <button class="grp-btn active" id="grpBtnAll" onclick="setGrpFilter('all')">Todos (Con + Sin Logo)</button>
  <button class="grp-btn" id="grpBtnCL" onclick="setGrpFilter('cl')">Con Logo</button>
  <button class="grp-btn" id="grpBtnSL" onclick="setGrpFilter('sl')">Sin Logo</button>
</div>
<div class="charts-row c1">
  <div class="chart-card" style="border-top:3px solid #2563eb;margin-bottom:0">
    <div class="chart-title" id="evol30Title">&#128200; Evoluci&oacute;n diaria del equipo &mdash; &uacute;ltimos 30 d&iacute;as h&aacute;biles</div>
    <div class="chart-subtitle" id="evol30Sub">Total l&iacute;neas picking regular (Con Logo + Sin Logo) &mdash; equipo completo</div>
    <div style="position:relative;height:200px"><canvas id="chartEvol30"></canvas></div>
  </div>
</div>

<div class="kpi-grid g5" style="margin-top:18px">
  <div class="kpi-card blue"><div class="kpi-label">Operarios Activos</div><div class="kpi-value" id="kpiOps">-</div><div class="kpi-sub" id="kpiOpsSub">Picking regular</div></div>
  <div class="kpi-card green"><div class="kpi-label">Lineas / D&iacute;a</div><div class="kpi-value" id="kpiLD">-</div><div class="kpi-sub">Target: $TARGET lin/d&iacute;a</div></div>
  <div class="kpi-card amber"><div class="kpi-label">Cumplimiento</div><div class="kpi-value" id="kpiCum">-</div><div class="kpi-sub" id="kpiCumSub">Total lineas: -</div></div>
  <div class="kpi-card red"><div class="kpi-label">Reclamos</div><div class="kpi-value" id="kpiRC">-</div><div class="kpi-sub">Tasa: <span id="kpiRate">-</span>/1000 lin</div></div>
  <div class="kpi-card purple"><div class="kpi-label">Unidades</div><div class="kpi-value" style="font-size:22px" id="kpiU">-</div><div class="kpi-sub"><span id="kpiOlas2">-</span> olas</div></div>
</div>
<div class="kpi-grid g3">
  <div class="kpi-card teal"><div class="kpi-label">Personas Necesarias</div><div class="kpi-value" id="kpiStaffNec">-</div><div class="kpi-sub">Total lin / d&iacute;as hab / $TARGET</div></div>
  <div class="kpi-card blue"><div class="kpi-label">Personas Activas (prom/d&iacute;a)</div><div class="kpi-value" id="kpiStaffAct">-</div><div class="kpi-sub">Promedio diario del per&iacute;odo</div></div>
  <div class="kpi-card green"><div class="kpi-label">Total L&iacute;neas</div><div class="kpi-value" style="font-size:22px" id="kpiL">-</div><div class="kpi-sub">Picking regular acumulado</div></div>
</div>

<div class="info-box">&#127919; Target: <strong>$TARGET lineas/d&iacute;a</strong> &mdash; Verde &ge; $TARGET &mdash; Naranja &ge; $TARGET_WARN &mdash; Rojo &lt; $TARGET_WARN</div>

<div class="table-card" style="margin-bottom:18px">
  <div class="rank-title">&#128197; Resumen &mdash; &Uacute;ltimos 7 d&iacute;as de picking regular</div>
  <div class="rank-sub">Total equipo (SIN LOGO + CON LOGO) &mdash; target $TARGET lin/operario/d&iacute;a &mdash; <em>datos actuales, no filtran por per&iacute;odo</em></div>
  <table>
    <thead><tr>
      <th>Fecha</th><th>D&iacute;a</th>
      <th style="text-align:right">Operarios</th>
      <th style="text-align:right">Total L&iacute;neas</th>
      <th style="text-align:right">Lin/Op</th>
      <th style="text-align:right">Target d&iacute;a</th>
      <th style="text-align:center">Cumpl%</th>
      <th style="text-align:center">vs Ayer</th>
    </tr></thead>
    <tbody id="body7d"></tbody>
    <tfoot id="foot7d"></tfoot>
  </table>
</div>

<div class="charts-row c3">
  <div class="chart-card">
    <div class="chart-title" id="rankChartTitle">Lineas/D&iacute;a por Operario</div>
    <div class="chart-subtitle">Comparaci&oacute;n vs. target $TARGET</div>
    <div style="position:relative;height:260px"><canvas id="chartRanking"></canvas></div>
  </div>
  <div class="chart-card">
    <div class="chart-title">Evoluci&oacute;n Equipo</div>
    <div class="chart-subtitle">Promedio mensual lineas/d&iacute;a (picking regular)</div>
    <div style="position:relative;height:260px"><canvas id="chartTeamTrend"></canvas></div>
  </div>
  <div class="chart-card">
    <div class="chart-title">Staffing &mdash; Necesarias vs Activas</div>
    <div class="chart-subtitle">Lineas regulares / d&iacute;as h&aacute;biles / $TARGET</div>
    <div style="position:relative;height:260px"><canvas id="chartStaff"></canvas></div>
  </div>
</div>

<div class="table-card">
  <div class="rank-title" id="rankTitle">Ranking &mdash; Picking Regular</div>
  <div class="rank-sub">Ordenado por lineas/d&iacute;a &nbsp;|&nbsp; MERMA y PIE DE MAQUINA en sus tabs propios</div>
  <table>
    <thead><tr><th>#</th><th>Operario</th><th style="text-align:right">D&iacute;as</th><th style="text-align:right">Olas</th><th style="text-align:right">Lineas</th><th style="text-align:center">Lin/D&iacute;a</th><th style="text-align:center">Cumpl%</th><th style="text-align:right">Unidades</th><th style="text-align:center">Reclamos</th><th style="text-align:center">Tasa Rec/1000</th><th style="text-align:center">Tendencia</th></tr></thead>
    <tbody id="rankBody"></tbody>
  </table>
</div>

<div class="table-card">
  <div class="rank-title">Totales por Operario &mdash; Per&iacute;odo seleccionado</div>
  <div class="rank-sub">Acumulado del per&iacute;odo &mdash; ordenado por total l&iacute;neas</div>
  <table>
    <thead><tr><th>Operario</th><th style="text-align:right">Total D&iacute;as</th><th style="text-align:right">Total Olas</th><th style="text-align:right">Total L&iacute;neas</th><th style="text-align:center">Prom Lin/D&iacute;a</th><th style="text-align:center">Cumpl%</th><th style="text-align:right">Total Unidades</th><th style="text-align:center">Total Reclamos</th></tr></thead>
    <tbody id="totalesBody"></tbody>
  </table>
</div>

<div class="charts-row c1">
  <div class="chart-card">
    <div class="chart-title">Evoluci&oacute;n Individual &mdash; Lineas/D&iacute;a</div>
    <div class="chart-subtitle">Cada l&iacute;nea = un operario &mdash; Operario seleccionado resaltado</div>
    <div style="position:relative;height:300px"><canvas id="chartEvol"></canvas></div>
  </div>
</div>

</div><!-- /sec-picking -->

<!-- ===== SECCION: PIE DE MAQUINA ===== -->
<div id="sec-pie" class="sec">

<div class="kpi-grid g4">
  <div class="kpi-card blue"><div class="kpi-label">Turno 1 &mdash; Lin/D&iacute;a Prom</div><div class="kpi-value" id="pieM1LD">-</div><div class="kpi-sub">Turno 06:00&ndash;14:59</div></div>
  <div class="kpi-card green"><div class="kpi-label">Turno 2 &mdash; Lin/D&iacute;a Prom</div><div class="kpi-value" id="pieM2LD">-</div><div class="kpi-sub">Resto del d&iacute;a</div></div>
  <div class="kpi-card amber"><div class="kpi-label">Total L&iacute;neas T1</div><div class="kpi-value" style="font-size:22px" id="pieM1L">-</div><div class="kpi-sub">Acumulado per&iacute;odo</div></div>
  <div class="kpi-card purple"><div class="kpi-label">Total L&iacute;neas T2</div><div class="kpi-value" style="font-size:22px" id="pieM2L">-</div><div class="kpi-sub">Acumulado per&iacute;odo</div></div>
</div>

<div class="charts-row c1">
  <div class="chart-card">
    <div class="chart-title">PIE DE M&Aacute;QUINA &mdash; Turno 1 vs Turno 2</div>
    <div class="chart-subtitle">06-15h vs resto del d&iacute;a &mdash; Sin target (volumen seg&uacute;n producci&oacute;n)</div>
    <div style="position:relative;height:280px"><canvas id="chartPie"></canvas></div>
  </div>
</div>

<div id="pieNoData" class="nodata" style="display:none">&#9888;&#65039; No hay datos de PIE DE M&Aacute;QUINA en el Excel fuente todav&iacute;a. El grupo aparecer&aacute; autom&aacute;ticamente cuando se carguen filas con ese grupo de &oacute;rdenes.</div>

</div><!-- /sec-pie -->

<!-- ===== SECCION: MUESTRA SIMPLE ===== -->
<div id="sec-muestra" class="sec">

<div class="kpi-grid g5">
  <div class="kpi-card purple"><div class="kpi-label">Lin/D&iacute;a Promedio</div><div class="kpi-value" id="lezLD">-</div><div class="kpi-sub">Muestra Simple</div></div>
  <div class="kpi-card blue"><div class="kpi-label">Total L&iacute;neas</div><div class="kpi-value" style="font-size:22px" id="lezL">-</div><div class="kpi-sub">Per&iacute;odo seleccionado</div></div>
  <div class="kpi-card green"><div class="kpi-label">D&iacute;as Trabajados</div><div class="kpi-value" id="lezDias">-</div><div class="kpi-sub">Con producci&oacute;n de muestras</div></div>
  <div class="kpi-card amber"><div class="kpi-label">Unidades</div><div class="kpi-value" style="font-size:22px" id="lezU">-</div><div class="kpi-sub">Unidades procesadas</div></div>
  <div class="kpi-card slate"><div class="kpi-label">vs Equipo Picking</div><div class="kpi-value" id="lezVsTeam">-</div><div class="kpi-sub">Diferencia lin/d&iacute;a</div></div>
</div>
<div class="kpi-grid g2">
  <div class="kpi-card red"><div class="kpi-label">Reclamos</div><div class="kpi-value" id="lezRC">-</div><div class="kpi-sub">Muestra Simple</div></div>
  <div class="kpi-card teal"><div class="kpi-label">Olas / Pedidos</div><div class="kpi-value" id="lezOlas">-</div><div class="kpi-sub">Total olas del per&iacute;odo</div></div>
</div>

<div class="charts-row c1">
  <div class="chart-card">
    <div class="chart-title">LEZCANO AGUST&Iacute;N &mdash; Evoluci&oacute;n Mensual</div>
    <div class="chart-subtitle">Lineas/d&iacute;a por mes &mdash; Muestra Simple</div>
    <div style="position:relative;height:280px"><canvas id="chartLez"></canvas></div>
  </div>
</div>

</div><!-- /sec-muestra -->

<!-- ===== SECCION: RECLAMOS ===== -->
<div id="sec-reclamos" class="sec">

<div class="kpi-grid g5">
  <div class="kpi-card red"><div class="kpi-label">Total Reclamos</div><div class="kpi-value" id="recTotal">-</div><div class="kpi-sub">Per&iacute;odo seleccionado</div></div>
  <div class="kpi-card amber"><div class="kpi-label">Tasa Reclamos</div><div class="kpi-value" id="recTasa">-</div><div class="kpi-sub">Reclamos / 1000 l&iacute;neas picking</div></div>
  <div class="kpi-card gray"><div class="kpi-label">Sin Identificar</div><div class="kpi-value" id="recSinId">-</div><div class="kpi-sub">Sin operario asignado</div></div>
  <div class="kpi-card purple"><div class="kpi-label">Categor&iacute;a M&aacute;s Frecuente</div><div class="kpi-value" style="font-size:16px" id="recTopCat">-</div><div class="kpi-sub" id="recTopCatCnt">-</div></div>
  <div class="kpi-card blue"><div class="kpi-label">Operario con M&aacute;s Reclamos</div><div class="kpi-value" style="font-size:16px" id="recTopOp">-</div><div class="kpi-sub" id="recTopOpCnt">-</div></div>
</div>

<div class="charts-row c2">
  <div class="chart-card">
    <div class="chart-title">Reclamos por Mes</div>
    <div class="chart-subtitle">Evoluci&oacute;n mensual de reclamos</div>
    <div style="position:relative;height:280px"><canvas id="chartRecMon"></canvas></div>
  </div>
  <div class="chart-card">
    <div class="chart-title">Reclamos por Categor&iacute;a</div>
    <div class="chart-subtitle">Top categor&iacute;as del per&iacute;odo &mdash; ordenado por frecuencia</div>
    <div style="position:relative;height:280px"><canvas id="chartRecCat"></canvas></div>
  </div>
</div>

<div class="charts-row c1">
  <div class="chart-card">
    <div class="chart-title" id="recOpChartTitle">Reclamos por Operario</div>
    <div class="chart-subtitle">Ordenado de mayor a menor &mdash; top 15</div>
    <div style="position:relative;height:260px"><canvas id="chartRecOp"></canvas></div>
  </div>
</div>

<div class="charts-row c2">
  <div class="table-card" style="margin-bottom:0">
    <div class="rank-title">Ranking por Operario</div>
    <div class="rank-sub">Reclamos, tasa y categor&iacute;a frecuente &mdash; per&iacute;odo seleccionado</div>
    <table>
      <thead><tr><th>#</th><th>Operario</th><th style="text-align:right">Reclamos</th><th style="text-align:right">Tasa/1000</th><th>Categor&iacute;a Frecuente</th></tr></thead>
      <tbody id="recRankBody"></tbody>
    </table>
  </div>
  <div class="table-card" style="margin-bottom:0">
    <div class="rank-title">Ranking por Categor&iacute;a</div>
    <div class="rank-sub">Frecuencia de cada tipo de reclamo &mdash; per&iacute;odo seleccionado</div>
    <table>
      <thead><tr><th>#</th><th>Categor&iacute;a</th><th style="text-align:right">Reclamos</th><th style="text-align:right">% del Total</th></tr></thead>
      <tbody id="recCatBody"></tbody>
    </table>
  </div>
</div>

</div><!-- /sec-reclamos -->

<!-- ===== SECCION: CONTROL ===== -->
<div id="sec-control" class="sec">
  <div class="kpi-grid g4" style="margin-bottom:18px">
    <div class="kpi-card blue"><div class="kpi-label">UND Procesadas</div><div class="kpi-value" id="ctrlKpiUnd">&#8212;</div><div class="kpi-sub" id="ctrlKpiUndSub"></div></div>
    <div class="kpi-card amber"><div class="kpi-label">ORD Procesadas</div><div class="kpi-value" id="ctrlKpiOrd">&#8212;</div><div class="kpi-sub" id="ctrlKpiOrdSub"></div></div>
    <div class="kpi-card green"><div class="kpi-label">UND/d&iacute;a promedio</div><div class="kpi-value" id="ctrlKpiUdpd">&#8212;</div><div class="kpi-sub" id="ctrlKpiUdpdSub"></div></div>
    <div class="kpi-card teal"><div class="kpi-label">Operarios</div><div class="kpi-value" id="ctrlKpiOps">&#8212;</div><div class="kpi-sub" id="ctrlKpiOpsSub"></div></div>
  </div>
  <div class="charts-grid" style="margin-bottom:18px">
    <div class="chart-card"><div class="chart-title">Evoluci&oacute;n mensual &mdash; UND/d&iacute;a</div><canvas id="chartCtrlEvol" height="220"></canvas></div>
    <div class="chart-card"><div class="chart-title">Ranking operarios &mdash; UND totales</div><canvas id="chartCtrlRank" height="220"></canvas></div>
  </div>
  <div class="chart-card" style="padding:14px"><div class="rank-title">Detalle mensual</div><div id="ctrlTable"></div></div>
</div><!-- /sec-control -->

<!-- ===== SECCION: MAQUINISTAS ===== -->
<div id="sec-maquinistas" class="sec">
  <div class="kpi-grid g4" style="margin-bottom:18px">
    <div class="kpi-card blue"><div class="kpi-label">Total MOV</div><div class="kpi-value" id="maqKpiMov">&#8212;</div><div class="kpi-sub" id="maqKpiMovSub"></div></div>
    <div class="kpi-card teal"><div class="kpi-label">Total UND</div><div class="kpi-value" id="maqKpiUnd">&#8212;</div><div class="kpi-sub" id="maqKpiUndSub"></div></div>
    <div class="kpi-card green"><div class="kpi-label">MOV/d&iacute;a promedio</div><div class="kpi-value" id="maqKpiMpd">&#8212;</div><div class="kpi-sub" id="maqKpiMpdSub"></div></div>
    <div class="kpi-card amber"><div class="kpi-label">Operarios</div><div class="kpi-value" id="maqKpiOps">&#8212;</div><div class="kpi-sub" id="maqKpiOpsSub"></div></div>
  </div>
  <div class="charts-grid" style="margin-bottom:18px">
    <div class="chart-card"><div class="chart-title">Evoluci&oacute;n MOV/d&iacute;a &mdash; por operario</div><canvas id="chartMaqEvol" height="220"></canvas></div>
    <div class="chart-card"><div id="maqBreakSub" class="chart-title">Desglose por tarea</div><canvas id="chartMaqBreak" height="220"></canvas></div>
  </div>
  <div class="chart-card" style="padding:14px"><div class="rank-title">Detalle mensual completo</div><div id="maqDetailTable"></div></div>
</div><!-- /sec-maquinistas -->

<!-- ===== SECCION: RESUMEN GENERAL ===== -->
<div id="sec-resumen" class="sec">
<div id="resumenContent"></div>
</div><!-- /sec-resumen -->

<!-- ===== SECCION: BONOS ===== -->
<div id="sec-eficiencia" class="sec">
<div class="grp-btns" style="margin-bottom:16px">
  <button class="grp-btn active" id="bonBtn-pick" onclick="bonSetSub('pick')">Pickeadores</button>
  <button class="grp-btn" id="bonBtn-maq" onclick="bonSetSub('maq')">Maquinistas</button>
  <button class="grp-btn" id="bonBtn-ctrl" onclick="bonSetSub('ctrl')">Control</button>
  <button class="grp-btn" id="bonBtn-pie" onclick="bonSetSub('pie')">Pie de M&aacute;quina</button>
</div>
<div id="bonGrupalWrap" style="margin-bottom:12px"></div>
<div id="bonContent"></div>
</div><!-- /sec-eficiencia -->

</div><!-- /container -->
<footer>Dashboard v5 &mdash; Zecat Art&iacute;culos Promocionales SA &nbsp;|&nbsp; $($NOW.ToString('dd/MM/yyyy HH:mm'))</footer>

<script>
Chart.defaults.font.family='Arial,sans-serif';
Chart.defaults.font.size=12;

const TARGET=$TARGET;
const TARGET_WARN=$TARGET_WARN;
const allMonKeys=[$jsAllMons];
const monLabels=[$jsMonLabels];
const teamLD=[$jsTeamLD];
const teamColors=[$jsTeamPtClrs];
const staffNec=[$jsStaffNec];
const staffAct=[$jsStaffAct];
const merma1=[$jsMerma1];
const merma2=[$jsMerma2];
const pie1=[$jsPie1];
const pie2=[$jsPie2];
const monthlyData=$jsMonthlyData;
const lezMonthly=$jsLezMonthly;
const merma1L=[$jsMerma1L];
const merma2L=[$jsMerma2L];
const recMonthly=$jsRecMonthly;
const MES=['','Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
const day7Rows=$jsDay7Rows;
const evol30Labels=$jsEvol30Labels;
const evol30Data=$jsEvol30Data;
const evol30CL=$jsEvol30CL;
const evol30SL=$jsEvol30SL;
const evol30Target=$jsEvol30Target;
var _grpFilter='all';
const ctrlExcelData=$jsCtrlData;
const maqDetailData=$jsMaqDetailData;
var _chartCtrlEvol=null,_chartCtrlRank=null,_chartMaqEvol=null,_chartMaqBreak=null;

function chartColors(){var dark=document.body.classList.contains('dark');return{grid:dark?'rgba(255,255,255,0.08)':'rgba(0,0,0,0.08)',tick:dark?'#94a3b8':'#555',leg:dark?'#cbd5e1':'#333'};}
function buildControl(){
  if(!ctrlExcelData||!Object.keys(ctrlExcelData).length)return;
  var anio=document.getElementById('selAnio')?document.getElementById('selAnio').value:'all';
  var mes=document.getElementById('selMes')?document.getElementById('selMes').value:'all';
  var months=Object.keys(ctrlExcelData).sort();
  var sel=months.filter(function(ym){var p=ym.split('-');if(anio!=='all'&&p[0]!==anio)return false;if(mes!=='all'&&parseInt(p[1])!==parseInt(mes))return false;return true;});
  if(!sel.length)sel=[months[months.length-1]];
  var cc=chartColors();
  var opsMap={};sel.forEach(function(ym){(ctrlExcelData[ym]||[]).forEach(function(op){if(!opsMap[op.nm])opsMap[op.nm]={nm:op.nm,und:0,ord:0,dias:0};opsMap[op.nm].und+=op.und;opsMap[op.nm].ord+=op.ord;opsMap[op.nm].dias+=op.dias;});});
  var ops=Object.values(opsMap).sort(function(a,b){return b.und-a.und;});
  var totUnd=ops.reduce(function(s,o){return s+o.und;},0),totOrd=ops.reduce(function(s,o){return s+o.ord;},0),totDias=ops.reduce(function(s,o){return s+o.dias;},0);
  var avgUpd=totDias?Math.round(totUnd/totDias):0;
  document.getElementById('ctrlKpiUnd').textContent=totUnd.toLocaleString('es-AR');
  document.getElementById('ctrlKpiUndSub').textContent=avgUpd+' und/dia prom';
  document.getElementById('ctrlKpiOrd').textContent=totOrd.toLocaleString('es-AR');
  document.getElementById('ctrlKpiOrdSub').textContent=totOrd.toLocaleString('es-AR')+' ordenes';
  document.getElementById('ctrlKpiUdpd').textContent=avgUpd.toLocaleString('es-AR');
  document.getElementById('ctrlKpiUdpdSub').textContent=sel.length+' mes'+(sel.length>1?'es':'');
  document.getElementById('ctrlKpiOps').textContent=ops.length;
  document.getElementById('ctrlKpiOpsSub').textContent='operarios';
  var CLRS=['#2563eb','#16a34a','#d97706','#7c3aed','#0891b2'];
  var MON_S=['','Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];
  var monLbl=months.map(function(ym){var p=ym.split('-');return MON_S[parseInt(p[1])]+' '+p[0].slice(2);});
  var allNms=[];months.forEach(function(ym){(ctrlExcelData[ym]||[]).forEach(function(op){if(allNms.indexOf(op.nm)<0)allNms.push(op.nm);});});
  var evolDs=allNms.map(function(nm,i){return{label:nm.split(' ')[0],data:months.map(function(ym){var op=(ctrlExcelData[ym]||[]).find(function(o){return o.nm===nm;});return(op&&op.dias)?Math.round(op.und/op.dias):null;}),borderColor:CLRS[i%CLRS.length],backgroundColor:CLRS[i%CLRS.length]+'33',tension:0.35,spanGaps:true,pointRadius:4,fill:false};});
  var isDarkC=document.body.classList.contains('dark');
  var tooltipBase={backgroundColor:isDarkC?'#1e293b':'#fff',titleColor:isDarkC?'#f1f5f9':'#111',bodyColor:isDarkC?'#94a3b8':'#555',borderColor:isDarkC?'#334155':'#e5e7eb',borderWidth:1,padding:10,cornerRadius:8};
  var el=document.getElementById('chartCtrlEvol');
  var evolOpts={responsive:true,maintainAspectRatio:true,interaction:{mode:'index',intersect:false},plugins:{legend:{labels:{color:cc.leg,font:{size:11},usePointStyle:true,pointStyleWidth:8,boxHeight:8}},tooltip:{...tooltipBase,callbacks:{label:function(ctx){return ' '+ctx.dataset.label+': '+Math.round(ctx.parsed.y).toLocaleString('es-AR')+' und/dia';}}}},scales:{x:{grid:{color:cc.grid,lineWidth:0.5},border:{dash:[4,4]},ticks:{color:cc.tick,font:{size:10}}},y:{grid:{color:cc.grid,lineWidth:0.5},ticks:{color:cc.tick,font:{size:10},callback:function(v){return v.toLocaleString('es-AR');}},title:{display:true,text:'UND/dia',color:cc.tick,font:{size:10}}}}};
  if(_chartCtrlEvol){_chartCtrlEvol.data.labels=monLbl;_chartCtrlEvol.data.datasets=evolDs;_chartCtrlEvol.options.scales.x.grid.color=cc.grid;_chartCtrlEvol.options.scales.y.grid.color=cc.grid;_chartCtrlEvol.options.scales.x.ticks.color=cc.tick;_chartCtrlEvol.options.scales.y.ticks.color=cc.tick;_chartCtrlEvol.options.plugins.legend.labels.color=cc.leg;_chartCtrlEvol.update();}
  else if(el){_chartCtrlEvol=new Chart(el,{type:'line',data:{labels:monLbl,datasets:evolDs},options:evolOpts});}
  var rkDs=[{label:'UND',data:ops.map(function(o){return o.und;}),backgroundColor:ops.map(function(o,i){return CLRS[i%CLRS.length];}),borderRadius:6,borderSkipped:false}];
  var rk=document.getElementById('chartCtrlRank');
  var rkOpts={responsive:true,maintainAspectRatio:true,indexAxis:'y',plugins:{legend:{display:false},tooltip:{...tooltipBase,callbacks:{label:function(ctx){return ' '+ctx.parsed.x.toLocaleString('es-AR')+' und';}}}},scales:{x:{grid:{color:cc.grid,lineWidth:0.5},ticks:{color:cc.tick,font:{size:10},callback:function(v){return v>=1000?(v/1000).toFixed(0)+'k':v;}},border:{dash:[4,4]}},y:{grid:{display:false},ticks:{color:cc.tick,font:{size:11,weight:'600'}}}}};
  if(_chartCtrlRank){_chartCtrlRank.data.labels=ops.map(function(o){return o.nm.split(' ')[0];});_chartCtrlRank.data.datasets=rkDs;_chartCtrlRank.update();}
  else if(rk){_chartCtrlRank=new Chart(rk,{type:'bar',data:{labels:ops.map(function(o){return o.nm.split(' ')[0];}),datasets:rkDs},options:rkOpts});}
  var isDark=document.body.classList.contains('dark'),thBg=isDark?'#1e293b':'#f5f5f5';
  var MON=['','Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];
  var h='<table style="width:100%;border-collapse:collapse;font-size:12.5px"><thead><tr style="background:'+thBg+'"><th style="padding:8px 12px;text-align:left;border-bottom:2px solid #e0e0e0">Mes</th><th style="padding:8px 12px;text-align:left;border-bottom:2px solid #e0e0e0">Operario</th><th style="padding:8px 12px;text-align:right;border-bottom:2px solid #e0e0e0">Dias</th><th style="padding:8px 12px;text-align:right;border-bottom:2px solid #e0e0e0">UND</th><th style="padding:8px 12px;text-align:right;border-bottom:2px solid #e0e0e0">UND/dia</th><th style="padding:8px 12px;text-align:right;border-bottom:2px solid #e0e0e0">ORD</th></tr></thead><tbody>';
  var alt=false;months.forEach(function(ym){(ctrlExcelData[ym]||[]).forEach(function(op){var upd=op.dias?Math.round(op.und/op.dias):0,clr=upd>=150?'#16a34a':upd>=80?'#2563eb':upd>=40?'#d97706':'#dc2626',bg=alt?(isDark?'#1e293b':'#f9f9f9'):(isDark?'#0f172a':'#fff'),p=ym.split('-');alt=!alt;h+='<tr style="background:'+bg+'"><td style="padding:7px 12px;color:#888">'+MON[parseInt(p[1])]+' '+p[0]+'</td><td style="padding:7px 12px;font-weight:600">'+op.nm+'</td><td style="padding:7px 12px;text-align:right;color:#888">'+op.dias+'</td><td style="padding:7px 12px;text-align:right;font-weight:700">'+op.und.toLocaleString('es-AR')+'</td><td style="padding:7px 12px;text-align:right;color:'+clr+';font-weight:700">'+upd+'</td><td style="padding:7px 12px;text-align:right;color:#888">'+op.ord+'</td></tr>';});});
  h+='</tbody></table>';document.getElementById('ctrlTable').innerHTML=h;
}
function buildMaquinistas(){
  if(!maqDetailData||!Object.keys(maqDetailData).length)return;
  var anio=document.getElementById('selAnio')?document.getElementById('selAnio').value:'all';
  var mes=document.getElementById('selMes')?document.getElementById('selMes').value:'all';
  var months=Object.keys(maqDetailData).sort();
  var sel=months.filter(function(ym){var p=ym.split('-');if(anio!=='all'&&p[0]!==anio)return false;if(mes!=='all'&&parseInt(p[1])!==parseInt(mes))return false;return true;});
  if(!sel.length)sel=[months[months.length-1]];
  var cc=chartColors();
  var opsMap={};sel.forEach(function(ym){(maqDetailData[ym]||{ops:[]}).ops.forEach(function(op){if(!opsMap[op.nm])opsMap[op.nm]={nm:op.nm,mov:0,und:0,dias:0};opsMap[op.nm].mov+=op.mov;opsMap[op.nm].und+=op.und;opsMap[op.nm].dias+=op.dias;});});
  var ops=Object.values(opsMap).sort(function(a,b){return b.mov-a.mov;});
  var totMov=ops.reduce(function(s,o){return s+o.mov;},0),totUnd=ops.reduce(function(s,o){return s+o.und;},0),totDias=ops.reduce(function(s,o){return s+o.dias;},0);
  var avgMpd=totDias?Math.round(totMov/totDias):0;
  document.getElementById('maqKpiMov').textContent=totMov.toLocaleString('es-AR');
  document.getElementById('maqKpiMovSub').textContent=avgMpd+' mov/dia prom';
  document.getElementById('maqKpiUnd').textContent=totUnd>0?totUnd.toLocaleString('es-AR'):'N/D';
  document.getElementById('maqKpiMpd').textContent=avgMpd.toLocaleString('es-AR');
  document.getElementById('maqKpiMpdSub').textContent=totMov.toLocaleString('es-AR')+' mov totales';
  document.getElementById('maqKpiOps').textContent=ops.length;
  document.getElementById('maqKpiOpsSub').textContent=sel.length+' mes'+(sel.length>1?'es':'');
  var CLRS=['#2563eb','#16a34a','#d97706','#7c3aed'];
  var allNms=[];months.forEach(function(ym){(maqDetailData[ym]||{ops:[]}).ops.forEach(function(op){if(allNms.indexOf(op.nm)<0)allNms.push(op.nm);});});
  var MON_S=['','Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];
  var monLbl=months.map(function(ym){var p=ym.split('-');return MON_S[parseInt(p[1])]+' '+p[0].slice(2);});
  var evolDs=allNms.map(function(nm,i){return{label:nm.split(' ')[0]+' '+(nm.split(' ')[1]||'').charAt(0)+'.',data:months.map(function(ym){var op=(maqDetailData[ym]||{ops:[]}).ops.find(function(o){return o.nm===nm;});return(op&&op.dias)?Math.round(op.mov/op.dias):null;}),borderColor:CLRS[i%CLRS.length],backgroundColor:CLRS[i%CLRS.length]+'33',tension:0.35,spanGaps:true,pointRadius:4,fill:false};});
  var isDarkM=document.body.classList.contains('dark');
  var ttBase={backgroundColor:isDarkM?'#1e293b':'#fff',titleColor:isDarkM?'#f1f5f9':'#111',bodyColor:isDarkM?'#94a3b8':'#555',borderColor:isDarkM?'#334155':'#e5e7eb',borderWidth:1,padding:10,cornerRadius:8};
  var maqEvolEl=document.getElementById('chartMaqEvol');
  var maqEvolOpts={responsive:true,maintainAspectRatio:true,interaction:{mode:'index',intersect:false},plugins:{legend:{labels:{color:cc.leg,font:{size:11},usePointStyle:true,pointStyleWidth:8,boxHeight:8}},tooltip:{...ttBase,callbacks:{label:function(ctx){return ' '+ctx.dataset.label+': '+Math.round(ctx.parsed.y).toLocaleString('es-AR')+' mov/dia';}}}},scales:{x:{grid:{color:cc.grid,lineWidth:0.5},border:{dash:[4,4]},ticks:{color:cc.tick,font:{size:10}}},y:{grid:{color:cc.grid,lineWidth:0.5},ticks:{color:cc.tick,font:{size:10},callback:function(v){return v.toLocaleString('es-AR');}},title:{display:true,text:'MOV/dia',color:cc.tick,font:{size:10}}}}};
  if(_chartMaqEvol){_chartMaqEvol.data.labels=monLbl;_chartMaqEvol.data.datasets=evolDs;_chartMaqEvol.options.scales.x.grid.color=cc.grid;_chartMaqEvol.options.scales.y.grid.color=cc.grid;_chartMaqEvol.options.scales.x.ticks.color=cc.tick;_chartMaqEvol.options.scales.y.ticks.color=cc.tick;_chartMaqEvol.options.plugins.legend.labels.color=cc.leg;_chartMaqEvol.update();}
  else if(maqEvolEl){_chartMaqEvol=new Chart(maqEvolEl,{type:'line',data:{labels:monLbl,datasets:evolDs},options:maqEvolOpts});}
  var lastYm=sel[sel.length-1],lastOps=(maqDetailData[lastYm]||{ops:[]}).ops;
  var TASK_LABELS=['PA02','Recep.','RL01','Pallet','Flow'],TASK_KEYS=['pa02','rec','rl01','pallet','flow'];
  var breakDs=lastOps.map(function(op,i){return{label:op.nm.split(' ')[0],data:TASK_KEYS.map(function(k){return op[k]||0;}),backgroundColor:CLRS[i%CLRS.length],borderRadius:6,borderSkipped:false};});
  var maqBrkEl=document.getElementById('chartMaqBreak');
  document.getElementById('maqBreakSub').textContent='Desglose - '+(maqDetailData[lastYm]||{lbl:lastYm}).lbl;
  var breakOpts={responsive:true,maintainAspectRatio:true,plugins:{legend:{labels:{color:cc.leg,font:{size:11},usePointStyle:true,pointStyleWidth:8,boxHeight:8}},tooltip:{...ttBase,callbacks:{label:function(ctx){return ' '+ctx.dataset.label+': '+ctx.parsed.y.toLocaleString('es-AR')+' mov';}}}},scales:{x:{grid:{display:false},ticks:{color:cc.tick,font:{size:11}}},y:{grid:{color:cc.grid,lineWidth:0.5},ticks:{color:cc.tick,font:{size:10},callback:function(v){return v>=1000?(v/1000).toFixed(0)+'k':v;}}}}};
  if(_chartMaqBreak){_chartMaqBreak.data.datasets=breakDs;_chartMaqBreak.update();}
  else if(maqBrkEl){_chartMaqBreak=new Chart(maqBrkEl,{type:'bar',data:{labels:TASK_LABELS,datasets:breakDs},options:breakOpts});}
  var isDark=document.body.classList.contains('dark'),thBg=isDark?'#1e293b':'#f5f5f5';
  var h='<table style="width:100%;border-collapse:collapse;font-size:12.5px"><thead><tr style="background:'+thBg+'"><th style="padding:8px 12px;text-align:left;border-bottom:2px solid #e0e0e0">Mes</th><th style="padding:8px 12px;text-align:left;border-bottom:2px solid #e0e0e0">Operario</th><th style="padding:8px 12px;text-align:right;border-bottom:2px solid #e0e0e0">Dias</th><th style="padding:8px 12px;text-align:right;border-bottom:2px solid #e0e0e0">MOV</th><th style="padding:8px 12px;text-align:right;border-bottom:2px solid #e0e0e0">MOV/dia</th><th style="padding:8px 12px;text-align:right;border-bottom:2px solid #e0e0e0">PA02</th><th style="padding:8px 12px;text-align:right;border-bottom:2px solid #e0e0e0">Recep.</th><th style="padding:8px 12px;text-align:right;border-bottom:2px solid #e0e0e0">RL01</th></tr></thead><tbody>';
  var alt=false;months.forEach(function(ym){var md=maqDetailData[ym]||{lbl:ym,ops:[]};md.ops.forEach(function(op){var mpd=op.dias?Math.round(op.mov/op.dias):0,clr=mpd>=60?'#16a34a':mpd>=30?'#2563eb':mpd>=10?'#d97706':'#dc2626',bg=alt?(isDark?'#1e293b':'#f9f9f9'):(isDark?'#0f172a':'#fff');alt=!alt;h+='<tr style="background:'+bg+'"><td style="padding:7px 12px;color:#888">'+md.lbl+'</td><td style="padding:7px 12px;font-weight:600">'+op.nm+'</td><td style="padding:7px 12px;text-align:right;color:#888">'+op.dias+'</td><td style="padding:7px 12px;text-align:right;font-weight:700">'+op.mov.toLocaleString('es-AR')+'</td><td style="padding:7px 12px;text-align:right;color:'+clr+';font-weight:700">'+mpd+'</td><td style="padding:7px 12px;text-align:right;color:#888">'+op.pa02+'</td><td style="padding:7px 12px;text-align:right;color:#888">'+op.rec+'</td><td style="padding:7px 12px;text-align:right;color:#888">'+op.rl01+'</td></tr>';});});
  h+='</tbody></table>';document.getElementById('maqDetailTable').innerHTML=h;
}

// Helper: filtrar meses ctrl/maq segun selectores activos
function _bonFiltMon(months){
  var anio=document.getElementById('selAnio').value;
  var mes=document.getElementById('selMes').value;
  var s=months.filter(function(ym){var p=ym.split('-');if(anio!=='all'&&p[0]!==anio)return false;if(mes!=='all'&&parseInt(p[1])!==parseInt(mes))return false;return true;});
  return s.length?s:[months[months.length-1]];
}

// ===========================================================
// RESUMEN GENERAL
// ===========================================================
function buildResumen(){
  var idx=getFilteredIndices();
  var selKeys=idx.map(function(i){return allMonKeys[i];});
  var isDark=document.body.classList.contains('dark');
  var thBg=isDark?'#1e293b':'#f8fafc';
  var brd=isDark?'#334155':'#e5e7eb';
  var rowEven=isDark?'#1e293b':'#f9f9f9';
  var rowOdd=isDark?'#0f172a':'#fff';
  var pkL=0,pkU=0,pkRC=0,pkOlas=0,pkLDSum=0,pkValidM=0;
  selKeys.forEach(function(k){var d=monthlyData[k];if(!d)return;pkValidM++;pkL+=d.totL;pkU+=d.totU;pkRC+=d.totRC;pkOlas+=d.totOlas;pkLDSum+=d.avgLD;});
  var pkAvgLD=pkValidM?Math.round(pkLDSum/pkValidM*10)/10:0;
  var pkCumpl=Math.round(pkAvgLD/TARGET*1000)/10;
  var lezL=0,lezLDSum=0,lezValidM=0;
  selKeys.forEach(function(k){var d=lezMonthly[k];if(!d)return;lezValidM++;lezL+=d.lineas;lezLDSum+=d.ld;});
  var lezAvgLD=lezValidM?Math.round(lezLDSum/lezValidM*10)/10:0;
  var recTotal=0;
  selKeys.forEach(function(k){var d=recMonthly[k];if(d)recTotal+=d.cnt;});
  var recTasa=pkL?Math.round(recTotal/(pkL/1000)*100)/100:0;
  var ctrlMs=_bonFiltMon(Object.keys(ctrlExcelData).sort());
  var ctrlTotUnd=0,ctrlTotDias=0,ctrlOps={};
  ctrlMs.forEach(function(ym){(ctrlExcelData[ym]||[]).forEach(function(op){ctrlTotUnd+=op.und;ctrlTotDias+=op.dias;ctrlOps[op.nm]=true;});});
  var ctrlAvgUpd=ctrlTotDias?Math.round(ctrlTotUnd/ctrlTotDias):0;
  var maqMs=_bonFiltMon(Object.keys(maqDetailData).sort());
  var maqTotMov=0,maqTotDias=0,maqOps={};
  maqMs.forEach(function(ym){(maqDetailData[ym]||{ops:[]}).ops.forEach(function(op){maqTotMov+=op.mov;maqTotDias+=op.dias;maqOps[op.nm]=true;});});
  var maqAvgMpd=maqTotDias?Math.round(maqTotMov/maqTotDias):0;
  var pieM1Vals=idx.map(function(i){return merma1[i];}).filter(function(v){return v!==null&&v!==undefined;});
  var pieM2Vals=idx.map(function(i){return merma2[i];}).filter(function(v){return v!==null&&v!==undefined;});
  var pieM1Avg=pieM1Vals.length?Math.round(pieM1Vals.reduce(function(s,v){return s+v;},0)/pieM1Vals.length*10)/10:0;
  var pieM2Avg=pieM2Vals.length?Math.round(pieM2Vals.reduce(function(s,v){return s+v;},0)/pieM2Vals.length*10)/10:0;
  var cumplC=pkCumpl>=100?'#16a34a':pkCumpl>=83?'#d97706':'#dc2626';
  var recC=recTasa===0?'#16a34a':recTasa<=3?'#d97706':'#dc2626';
  var MON_S=['','Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];
  var periodLabel=selKeys.length===1?(function(){var p=selKeys[0].split('-');return MON_S[parseInt(p[1])]+' '+p[0];}()):(idx.length+' meses');
  function kCard(label,val,sub,color){return '<div style="padding:14px 18px;border-radius:10px;border:1px solid '+brd+';border-left:4px solid '+color+';background:'+thBg+'"><div style="font-size:11px;text-transform:uppercase;letter-spacing:.5px;color:#888;margin-bottom:4px">'+label+'</div><div style="font-size:26px;font-weight:800;color:'+color+'">'+val+'</div><div style="font-size:11px;color:#888;margin-top:2px">'+sub+'</div></div>';}
  function stBadge(ok){return ok?'<span style="background:#dcfce7;color:#16a34a;padding:2px 8px;border-radius:999px;font-size:11px;font-weight:700">&#10003; OK</span>':'<span style="background:#fee2e2;color:#dc2626;padding:2px 8px;border-radius:999px;font-size:11px;font-weight:700">&#9888; Revisar</span>';}
  var html='<div class="rank-title" style="margin-bottom:14px">&#128202; Resumen General &mdash; '+periodLabel+'</div>';
  html+='<div style="display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin-bottom:20px">';
  html+=kCard('Picking &mdash; Lin/D&iacute;a',pkAvgLD,pkCumpl+'% cumpl &mdash; '+pkL.toLocaleString('es-AR')+' lin &mdash; '+pkOlas+' olas',cumplC);
  html+=kCard('Muestra Simple &mdash; Lin/D&iacute;a',lezAvgLD?lezAvgLD:'-',lezL.toLocaleString('es-AR')+' lineas totales','#7c3aed');
  html+=kCard('Reclamos',recTotal,'Tasa: '+recTasa+'/1000 lin &mdash; '+pkU.toLocaleString('es-AR')+' und picking',recC);
  html+=kCard('Control &mdash; UND/D&iacute;a',ctrlAvgUpd?ctrlAvgUpd:'-',ctrlTotUnd.toLocaleString('es-AR')+' und &mdash; '+Object.keys(ctrlOps).length+' ops','#0891b2');
  html+=kCard('Maquinistas &mdash; MOV/D&iacute;a',maqAvgMpd?maqAvgMpd:'-',maqTotMov.toLocaleString('es-AR')+' mov &mdash; '+Object.keys(maqOps).length+' ops','#2563eb');
  html+=kCard('Pie de M&aacute;quina &mdash; T1',pieM1Avg?pieM1Avg:'-','Lin/d&iacute;a prom &mdash; T2: '+(pieM2Avg?pieM2Avg:'-'),'#0891b2');
  html+='</div>';
  html+='<div class="chart-card" style="padding:16px"><table style="width:100%;border-collapse:collapse;font-size:13px"><thead><tr style="background:'+thBg+'"><th style="padding:8px 14px;text-align:left;border-bottom:2px solid '+brd+'">Secci&oacute;n</th><th style="padding:8px 14px;text-align:right;border-bottom:2px solid '+brd+'">&Iacute;ndice</th><th style="padding:8px 14px;text-align:right;border-bottom:2px solid '+brd+'">Acumulado</th><th style="padding:8px 14px;text-align:right;border-bottom:2px solid '+brd+'">Estado</th></tr></thead><tbody>';
  html+='<tr style="background:'+rowEven+'"><td style="padding:8px 14px;font-weight:600">&#128230; Picking Regular</td><td style="text-align:right;padding:8px 14px;font-weight:700;color:'+cumplC+'">'+pkAvgLD+' lin/d&iacute;a ('+pkCumpl+'%)</td><td style="text-align:right;padding:8px 14px">'+pkL.toLocaleString('es-AR')+' lin / '+pkU.toLocaleString('es-AR')+' und</td><td style="text-align:right;padding:8px 14px">'+stBadge(pkCumpl>=83)+'</td></tr>';
  html+='<tr style="background:'+rowOdd+'"><td style="padding:8px 14px;font-weight:600">&#128203; Muestra Simple</td><td style="text-align:right;padding:8px 14px;font-weight:700;color:#7c3aed">'+lezAvgLD+' lin/d&iacute;a</td><td style="text-align:right;padding:8px 14px">'+lezL.toLocaleString('es-AR')+' lin</td><td style="text-align:right;padding:8px 14px">'+stBadge(lezAvgLD>0)+'</td></tr>';
  html+='<tr style="background:'+rowEven+'"><td style="padding:8px 14px;font-weight:600">&#128683; Reclamos</td><td style="text-align:right;padding:8px 14px;font-weight:700;color:'+recC+'">'+recTasa+'/1000 lin</td><td style="text-align:right;padding:8px 14px">'+recTotal+' reclamos &mdash; '+pkOlas+' olas</td><td style="text-align:right;padding:8px 14px">'+stBadge(recTasa<=3)+'</td></tr>';
  html+='<tr style="background:'+rowOdd+'"><td style="padding:8px 14px;font-weight:600">&#128269; Control</td><td style="text-align:right;padding:8px 14px;font-weight:700;color:#0891b2">'+ctrlAvgUpd+' und/d&iacute;a</td><td style="text-align:right;padding:8px 14px">'+ctrlTotUnd.toLocaleString('es-AR')+' und &mdash; '+Object.keys(ctrlOps).length+' ops</td><td style="text-align:right;padding:8px 14px">'+stBadge(ctrlAvgUpd>3000)+'</td></tr>';
  html+='<tr style="background:'+rowEven+'"><td style="padding:8px 14px;font-weight:600">&#128295; Maquinistas</td><td style="text-align:right;padding:8px 14px;font-weight:700;color:#2563eb">'+maqAvgMpd+' mov/d&iacute;a</td><td style="text-align:right;padding:8px 14px">'+maqTotMov.toLocaleString('es-AR')+' mov &mdash; '+Object.keys(maqOps).length+' ops</td><td style="text-align:right;padding:8px 14px">'+stBadge(maqAvgMpd>=100)+'</td></tr>';
  html+='<tr style="background:'+rowOdd+'"><td style="padding:8px 14px;font-weight:600">&#9881; Pie de M&aacute;quina</td><td style="text-align:right;padding:8px 14px;font-weight:700;color:#0891b2">T1: '+pieM1Avg+' &mdash; T2: '+pieM2Avg+'</td><td style="text-align:right;padding:8px 14px">lin/d&iacute;a prom</td><td style="text-align:right;padding:8px 14px">'+stBadge(pieM1Avg>0||pieM2Avg>0)+'</td></tr>';
  html+='</tbody></table></div>';
  document.getElementById('resumenContent').innerHTML=html;
}

// ===========================================================
// BONOS
// ===========================================================
var _bonSub='pick';
var _bonGrupal=true;
var _bonPres={};
var _bonCtrlRec={};
var _bonPieErr=0;
try{var _bg=localStorage.getItem('zecat-bon-grupal');if(_bg!==null)_bonGrupal=_bg==='1';}catch(e){}
try{var _bp=localStorage.getItem('zecat-bon-pres');if(_bp)_bonPres=JSON.parse(_bp);}catch(e){}
try{var _bc=localStorage.getItem('zecat-bon-ctrl');if(_bc)_bonCtrlRec=JSON.parse(_bc);}catch(e){}

function bonSetSub(sub){
  _bonSub=sub;
  ['pick','maq','ctrl','pie'].forEach(function(s){var b=document.getElementById('bonBtn-'+s);if(b)b.classList.toggle('active',s===sub);});
  _bonRenderContent();
}
function bonToggleGrupal(){
  var el=document.getElementById('bonGrupalChk');
  if(el)_bonGrupal=el.checked;
  try{localStorage.setItem('zecat-bon-grupal',_bonGrupal?'1':'0');}catch(e){}
  _bonRenderContent();
}
function bonTogglePres(key){
  var el=document.getElementById('bonPres-'+key);
  if(el)_bonPres[key]=el.checked;
  try{localStorage.setItem('zecat-bon-pres',JSON.stringify(_bonPres));}catch(e){}
  _bonRenderContent();
}
function _bonCalc(pres,n1,n2,grupal){
  var t=0;if(pres)t+=5;if(n1)t+=2.5;if(n2)t+=4.5;if(grupal)t+=3;return Math.min(t,15);
}
function _bonBadge(pct){
  var c=pct>=12?'#16a34a':pct>=7.5?'#2563eb':pct>=5?'#d97706':'#888';
  return '<span style="font-weight:800;font-size:15px;color:'+c+'">'+pct+'%</span>';
}
function buildEficiencia(){
  var isDark=document.body.classList.contains('dark');
  var thBg=isDark?'#1e293b':'#f8fafc';var brd=isDark?'#334155':'#e5e7eb';
  var gw=document.getElementById('bonGrupalWrap');
  if(gw){gw.innerHTML='<div style="display:flex;align-items:center;gap:12px;padding:12px 16px;background:'+thBg+';border-radius:10px;border:1px solid '+brd+';border-left:4px solid #16a34a"><label style="display:flex;align-items:center;gap:8px;font-weight:600;cursor:pointer"><input type="checkbox" id="bonGrupalChk" '+(_bonGrupal?'checked':'')+' onchange="bonToggleGrupal()" style="width:16px;height:16px;cursor:pointer"> Bono Grupal Dep&oacute;sito <span style="color:#16a34a;font-weight:700">+3%</span></label><span style="color:#888;font-size:12px">Aplica a todos cuando el dep&oacute;sito cumple el objetivo grupal</span></div>';}
  _bonRenderContent();
}
function _bonRenderContent(){
  var isDark=document.body.classList.contains('dark');
  var thBg=isDark?'#1e293b':'#f8fafc';var brd=isDark?'#334155':'#e5e7eb';
  var rowEven=isDark?'#1e293b':'#f9f9f9';var rowOdd=isDark?'#0f172a':'#fff';
  var html='';
  function boolIcon(v){return v?'<span style="color:#16a34a;font-weight:700">&#10003;</span>':'<span style="color:#dc2626">&#8212;</span>';}
  function presCheck(key){var chk=(_bonPres[key]!==undefined)?_bonPres[key]:true;return '<input type="checkbox" data-key="'+key+'" '+(chk?'checked':'')+' onchange="bonTogglePres(this.dataset.key)" id="bonPres-'+key+'" style="width:15px;height:15px;cursor:pointer">';}
  function tblHdr(extra){return '<table style="width:100%;border-collapse:collapse;font-size:13px"><thead><tr style="background:'+thBg+'"><th style="padding:8px 12px;text-align:left;border-bottom:2px solid '+brd+'">Operario</th>'+extra+'<th style="padding:8px 12px;text-align:center;border-bottom:2px solid '+brd+'">Presentismo<br><span style="font-size:10px;font-weight:400">5%</span></th><th style="padding:8px 12px;text-align:center;border-bottom:2px solid '+brd+'">N1<br><span style="font-size:10px;font-weight:400">+2.5%</span></th><th style="padding:8px 12px;text-align:center;border-bottom:2px solid '+brd+'">N2<br><span style="font-size:10px;font-weight:400">+4.5%</span></th><th style="padding:8px 12px;text-align:center;border-bottom:2px solid '+brd+'">Grupal<br><span style="font-size:10px;font-weight:400">+3%</span></th><th style="padding:8px 12px;text-align:center;border-bottom:2px solid '+brd+'">TOTAL</th></tr></thead><tbody>';}
  if(_bonSub==='pick'){
    var idx=getFilteredIndices();
    var selKeys=idx.map(function(i){return allMonKeys[i];});
    var opsMap={};
    selKeys.forEach(function(k){var d=monthlyData[k];if(!d)return;d.pickers.forEach(function(pk){if(!opsMap[pk.resp])opsMap[pk.resp]={dias:0,lineas:0};opsMap[pk.resp].dias+=pk.dias;opsMap[pk.resp].lineas+=pk.lineas;});});
    var pickers=Object.entries(opsMap).map(function(e){var resp=e[0],p=e[1];var ld=p.dias?Math.round(p.lineas/p.dias*10)/10:0;var cumpl=Math.round(ld/TARGET*1000)/10;return{nm:resp,ld:ld,cumpl:cumpl,n1:cumpl>=100,n2:cumpl>=120};}).sort(function(a,b){return b.ld-a.ld;});
    html+=tblHdr('<th style="padding:8px 12px;text-align:right;border-bottom:2px solid '+brd+'">Lin/D&iacute;a</th><th style="padding:8px 12px;text-align:right;border-bottom:2px solid '+brd+'">Cumpl%</th>');
    var alt=false;
    pickers.forEach(function(p){
      var key='pk'+p.nm.replace(/[^A-Za-z0-9]/g,'_');
      var bg=alt?rowEven:rowOdd;alt=!alt;
      var cumplC=p.cumpl>=100?'#16a34a':p.cumpl>=83?'#d97706':'#dc2626';
      var t=_bonCalc((_bonPres[key]!==undefined?_bonPres[key]:true),p.n1,p.n2,_bonGrupal);
      html+='<tr style="background:'+bg+'"><td style="padding:8px 12px;font-weight:600">'+p.nm+'</td><td style="text-align:right;padding:8px 12px;font-weight:700;color:'+cumplC+'">'+p.ld+'</td><td style="text-align:right;padding:8px 12px;color:'+cumplC+'">'+p.cumpl+'%</td><td style="text-align:center;padding:8px 12px">'+presCheck(key)+'</td><td style="text-align:center;padding:8px 12px">'+boolIcon(p.n1)+'</td><td style="text-align:center;padding:8px 12px">'+boolIcon(p.n2)+'</td><td style="text-align:center;padding:8px 12px">'+boolIcon(_bonGrupal)+'</td><td style="text-align:center;padding:8px 12px">'+_bonBadge(t)+'</td></tr>';
    });
    if(!pickers.length)html+='<tr><td colspan="8" style="text-align:center;padding:20px;color:#888">Sin datos para el per&iacute;odo seleccionado</td></tr>';
    html+='</tbody></table>';
  }
  else if(_bonSub==='maq'){
    var maqMs=_bonFiltMon(Object.keys(maqDetailData).sort());
    var opsMapM={};
    maqMs.forEach(function(ym){(maqDetailData[ym]||{ops:[]}).ops.forEach(function(op){if(!opsMapM[op.nm])opsMapM[op.nm]={mov:0,dias:0};opsMapM[op.nm].mov+=op.mov;opsMapM[op.nm].dias+=op.dias;});});
    var maqOps=Object.entries(opsMapM).map(function(e){var nm=e[0],o=e[1];var mpd=o.dias?Math.round(o.mov/o.dias):0;return{nm:nm,mpd:mpd,n1:mpd>=100,n2:mpd>=120};}).sort(function(a,b){return b.mpd-a.mpd;});
    html+=tblHdr('<th style="padding:8px 12px;text-align:right;border-bottom:2px solid '+brd+'">MOV/D&iacute;a</th>');
    var alt=false;
    maqOps.forEach(function(p){
      var key='mq'+p.nm.replace(/[^A-Za-z0-9]/g,'_');
      var bg=alt?rowEven:rowOdd;alt=!alt;
      var mpdC=p.mpd>=100?'#16a34a':p.mpd>=50?'#d97706':'#dc2626';
      var t=_bonCalc((_bonPres[key]!==undefined?_bonPres[key]:true),p.n1,p.n2,_bonGrupal);
      html+='<tr style="background:'+bg+'"><td style="padding:8px 12px;font-weight:600">'+p.nm+'</td><td style="text-align:right;padding:8px 12px;font-weight:700;color:'+mpdC+'">'+p.mpd+'</td><td style="text-align:center;padding:8px 12px">'+presCheck(key)+'</td><td style="text-align:center;padding:8px 12px">'+boolIcon(p.n1)+'</td><td style="text-align:center;padding:8px 12px">'+boolIcon(p.n2)+'</td><td style="text-align:center;padding:8px 12px">'+boolIcon(_bonGrupal)+'</td><td style="text-align:center;padding:8px 12px">'+_bonBadge(t)+'</td></tr>';
    });
    if(!maqOps.length)html+='<tr><td colspan="7" style="text-align:center;padding:20px;color:#888">Sin datos para el per&iacute;odo seleccionado</td></tr>';
    html+='</tbody></table>';
  }
  else if(_bonSub==='ctrl'){
    var ctrlMs=_bonFiltMon(Object.keys(ctrlExcelData).sort());
    var opsMapC={};
    ctrlMs.forEach(function(ym){(ctrlExcelData[ym]||[]).forEach(function(op){if(!opsMapC[op.nm])opsMapC[op.nm]={und:0,dias:0};opsMapC[op.nm].und+=op.und;opsMapC[op.nm].dias+=op.dias;});});
    var ctrlOps=Object.entries(opsMapC).map(function(e){var nm=e[0],o=e[1];var upd=o.dias?Math.round(o.und/o.dias):0;return{nm:nm,upd:upd};}).sort(function(a,b){return b.upd-a.upd;});
    html+='<div style="background:#fef9c3;border:1px solid #fde047;border-radius:8px;padding:10px 14px;font-size:12px;color:#713f12;margin-bottom:12px">&#9888; Ingresá los reclamos por operario para calcular N1 y N2.</div>';
    html+=tblHdr('<th style="padding:8px 12px;text-align:right;border-bottom:2px solid '+brd+'">UND/D&iacute;a</th><th style="padding:8px 12px;text-align:center;border-bottom:2px solid '+brd+'">Reclamos</th>');
    var alt=false;
    ctrlOps.forEach(function(p){
      var key='ct'+p.nm.replace(/[^A-Za-z0-9]/g,'_');
      var rec=(_bonCtrlRec[key]!==undefined)?_bonCtrlRec[key]:0;
      var n1=p.upd>3000&&rec<=1;
      var n2=p.upd>3500&&rec===0;
      var bg=alt?rowEven:rowOdd;alt=!alt;
      var updC=p.upd>3500?'#16a34a':p.upd>3000?'#2563eb':p.upd>1500?'#d97706':'#dc2626';
      var t=_bonCalc((_bonPres[key]!==undefined?_bonPres[key]:true),n1,n2,_bonGrupal);
      html+='<tr style="background:'+bg+'"><td style="padding:8px 12px;font-weight:600">'+p.nm+'</td><td style="text-align:right;padding:8px 12px;font-weight:700;color:'+updC+'">'+p.upd.toLocaleString('es-AR')+'</td><td style="text-align:center;padding:8px 12px"><input type="number" min="0" max="99" value="'+rec+'" data-key="'+key+'" onchange="_bonCtrlRec[this.dataset.key]=parseInt(this.value)||0;try{localStorage.setItem('zecat-bon-ctrl',JSON.stringify(_bonCtrlRec));}catch(e){};_bonRenderContent();" style="width:60px;text-align:center;border:1px solid '+brd+';border-radius:4px;padding:2px 4px;font-size:13px;background:'+rowOdd+'"></td><td style="text-align:center;padding:8px 12px">'+presCheck(key)+'</td><td style="text-align:center;padding:8px 12px">'+boolIcon(n1)+'</td><td style="text-align:center;padding:8px 12px">'+boolIcon(n2)+'</td><td style="text-align:center;padding:8px 12px">'+boolIcon(_bonGrupal)+'</td><td style="text-align:center;padding:8px 12px">'+_bonBadge(t)+'</td></tr>';
    });
    if(!ctrlOps.length)html+='<tr><td colspan="8" style="text-align:center;padding:20px;color:#888">Sin datos para el per&iacute;odo seleccionado</td></tr>';
    html+='</tbody></table>';
  }
  else if(_bonSub==='pie'){
    var n1Pie=_bonPieErr<=1;var n2Pie=_bonPieErr===0;
    var tPie=_bonCalc(true,n1Pie,n2Pie,_bonGrupal);
    var tC=tPie>=12?'#16a34a':tPie>=7.5?'#2563eb':tPie>=5?'#d97706':'#888';
    html+='<div class="chart-card" style="padding:20px;max-width:520px"><div class="rank-title" style="margin-bottom:16px">&#9881;&#65039; Pie de M&aacute;quina &mdash; Datos Manuales</div>';
    html+='<div style="margin-bottom:20px"><label style="display:block;font-weight:600;margin-bottom:8px">N&uacute;mero de errores en el per&iacute;odo:</label><input type="number" min="0" max="99" value="'+_bonPieErr+'" onchange="_bonPieErr=parseInt(this.value)||0;_bonRenderContent();" style="font-size:20px;padding:8px 16px;border:2px solid '+brd+';border-radius:8px;width:130px;text-align:center"></div>';
    html+='<div style="display:grid;grid-template-columns:repeat(3,1fr);gap:12px">';
    html+='<div style="padding:14px;border-radius:8px;border:1px solid '+brd+';text-align:center;border-top:3px solid '+(n2Pie?'#16a34a':'#e5e7eb')+'"><div style="font-size:11px;color:#888;text-transform:uppercase;margin-bottom:6px">N2 &mdash; 0 errores</div><div style="font-size:24px;font-weight:800;color:'+(n2Pie?'#16a34a':'#888')+'">'+(n2Pie?'&#10003;':'&#8212;')+'</div><div style="font-size:11px;color:#888">+4.5%</div></div>';
    html+='<div style="padding:14px;border-radius:8px;border:1px solid '+brd+';text-align:center;border-top:3px solid '+(n1Pie?'#2563eb':'#e5e7eb')+'"><div style="font-size:11px;color:#888;text-transform:uppercase;margin-bottom:6px">N1 &mdash; &le;1 error</div><div style="font-size:24px;font-weight:800;color:'+(n1Pie?'#2563eb':'#888')+'">'+(n1Pie?'&#10003;':'&#8212;')+'</div><div style="font-size:11px;color:#888">+2.5%</div></div>';
    html+='<div style="padding:14px;border-radius:8px;border:1px solid '+brd+';text-align:center;border-top:3px solid '+tC+'"><div style="font-size:11px;color:#888;text-transform:uppercase;margin-bottom:6px">Total Bono</div><div style="font-size:28px;font-weight:900;color:'+tC+'">'+tPie+'%</div><div style="font-size:11px;color:#888">m&aacute;x 15%</div></div>';
    html+='</div><div style="margin-top:14px;font-size:12px;color:#888;padding:10px;background:'+rowEven+';border-radius:6px">Presentismo (5%) incluido. Grupal '+(_bonGrupal?'activado &#10003;':'desactivado &#8212;')+'.</div></div>';
  }
  document.getElementById('bonContent').innerHTML=html;
}

// Tab navigation
function switchTab(name){
  document.querySelectorAll('.sec').forEach(function(s){s.classList.remove('active');});
  document.querySelectorAll('.tab-btn').forEach(function(b){b.classList.remove('active');});
  document.getElementById('sec-'+name).classList.add('active');
  document.getElementById('btn-'+name).classList.add('active');
  document.getElementById('opFilterWrap').style.display=(name==='picking'||name==='reclamos'||name==='control'||name==='maquinistas')?'':'none';
  if(name==='control') buildControl();
  if(name==='maquinistas') buildMaquinistas();
  if(name==='resumen') buildResumen();
  if(name==='eficiencia') buildEficiencia();
}

// Poblar selector de anios
const years=[$jsYears];
const selAnioEl=document.getElementById('selAnio');
years.forEach(function(y){var o=document.createElement('option');o.value=y;o.textContent=y;selAnioEl.appendChild(o);});

// Poblar selector de operarios
const operarios=[$jsOperarios];
const selOpEl=document.getElementById('selOp');
operarios.forEach(function(op){var o=document.createElement('option');o.value=op;o.textContent=op;selOpEl.appendChild(o);});

// Ocultar meses sin datos del selector
(function(){var availMons=new Set(allMonKeys.map(function(k){return parseInt(k.split('-')[1]);}));
Array.from(document.getElementById('selMes').options).forEach(function(opt){if(opt.value!=='all'&&!availMons.has(parseInt(opt.value)))opt.remove();});})();

// Colores de rendimiento (verde/naranja/rojo)
function perfColor(v){return v>=TARGET?'#16a34a':v>=TARGET_WARN?'#d97706':'#dc2626';}
function perfColorRec(v){return v===0?'#16a34a':v<=3?'#d97706':'#dc2626';}

function getFilteredIndices(){
  var anio=document.getElementById('selAnio').value;
  var mes=document.getElementById('selMes').value;
  var idx=[];
  for(var i=0;i<allMonKeys.length;i++){
    var p=allMonKeys[i].split('-');
    if(anio!=='all'&&p[0]!==anio) continue;
    if(mes!=='all'&&parseInt(p[1])!==parseInt(mes)) continue;
    idx.push(i);
  }
  return idx;
}

function applyFilter(){
  var _s=document.querySelector('.sec.active');var _id=_s?_s.id:'';
  if(_id==='sec-control'){buildControl();return;}
  if(_id==='sec-maquinistas'){buildMaquinistas();return;}
  if(_id==='sec-resumen'){buildResumen();return;}
  if(_id==='sec-eficiencia'){buildEficiencia();return;}
  var idx=getFilteredIndices();
  var labels=idx.map(function(i){return monLabels[i];});
  var selKeys=idx.map(function(i){return allMonKeys[i];});
  var anio=document.getElementById('selAnio').value;
  var mes=document.getElementById('selMes').value;
  var selOp=document.getElementById('selOp').value;

  // Badge
  var badge='Mostrando: todos los datos';
  if(anio!=='all'||mes!=='all'||selOp!=='all'){
    var bp=[];
    if(anio!=='all') bp.push(anio);
    if(mes!=='all') bp.push(MES[parseInt(mes)]);
    if(selOp!=='all') bp.push(selOp);
    badge='Filtro: '+bp.join(' / ')+' ('+idx.length+' mes'+(idx.length!==1?'es':'')+')';
  }
  document.getElementById('filterBadge').textContent=badge;

  // ===== PICKING: agregar datos segun filtro de operario =====
  var aggL=0,aggU=0,aggRC=0,aggOlas=0,aggSN=0,aggSA=0;
  var aggOpsSum=0,aggLDSum=0,teamValidM=0;
  var opLDSum=0,opValidM=0;
  var rankMap={};

  selKeys.forEach(function(k){
    var d=monthlyData[k];
    if(!d) return;
    teamValidM++;
    aggOpsSum+=(d.pickeadores||d.totOps); aggLDSum+=d.avgLD; aggSN+=d.staffNec; aggSA+=(d.pickeadores||d.staffAct);
    if(selOp==='all'){aggL+=d.totL;aggU+=d.totU;aggRC+=d.totRC;aggOlas+=d.totOlas;}
    d.pickers.forEach(function(pk){
      if(selOp!=='all'&&pk.resp!==selOp) return;
      if(!rankMap[pk.resp]) rankMap[pk.resp]={dias:0,olas:0,lineas:0,unidades:0,recCnt:0,trend:pk.trend};
      rankMap[pk.resp].dias+=pk.dias; rankMap[pk.resp].olas+=pk.olas;
      rankMap[pk.resp].lineas+=pk.lineas; rankMap[pk.resp].unidades+=pk.unidades;
      rankMap[pk.resp].recCnt+=pk.recCnt; rankMap[pk.resp].trend=pk.trend;
      if(selOp!=='all'){aggL+=pk.lineas;aggU+=pk.unidades;aggRC+=pk.recCnt;aggOlas+=pk.olas;opLDSum+=pk.ld;opValidM++;}
    });
  });

  // Tendencia dinámica: si se filtra por un solo mes, calcular vs el mes anterior en allMonKeys
  if(selKeys.length===1){
    var curMonIdx=allMonKeys.indexOf(selKeys[0]);
    if(curMonIdx>0){
      var prevMonData=monthlyData[allMonKeys[curMonIdx-1]];
      if(prevMonData){
        var prevLdMap={};
        prevMonData.pickers.forEach(function(pk){prevLdMap[pk.resp]=pk.ld;});
        Object.keys(rankMap).forEach(function(resp){
          var curLd=rankMap[resp].dias?Math.round(rankMap[resp].lineas/rankMap[resp].dias*10)/10:0;
          var pLd=prevLdMap[resp];
          if(pLd!==undefined&&pLd>0){
            var delta=Math.round((curLd-pLd)*10)/10;
            rankMap[resp].trend=delta>3?'▲ +'+delta:delta<-3?'▼ '+delta:'→ '+delta;
          }else{rankMap[resp].trend='-';}
        });
      }
    }else{Object.keys(rankMap).forEach(function(resp){rankMap[resp].trend='-';});}
  }

  var dv=teamValidM||1;
  var mAvgLD=selOp==='all'?(teamValidM?Math.round(aggLDSum/dv*10)/10:0):(opValidM?Math.round(opLDSum/opValidM*10)/10:0);
  var mCumAv=Math.round(mAvgLD/TARGET*1000)/10;
  var mOps=selOp==='all'?(teamValidM?Math.round(aggOpsSum/dv*10)/10:0):1;
  var mRateG=aggL?Math.round(aggRC/(aggL/1000)*100)/100:0;
  var mSN=selOp==='all'?Math.round(aggSN/dv*10)/10:'-';
  var mSA=selOp==='all'?Math.round(aggSA/dv*10)/10:'-';

  // KPI cards picking
  document.getElementById('kpiOps').textContent=mOps;
  document.getElementById('kpiLD').textContent=mAvgLD;
  document.getElementById('kpiLD').style.color=perfColor(mAvgLD);
  document.getElementById('kpiCum').textContent=mCumAv+'%';
  document.getElementById('kpiCum').style.color=mCumAv>=100?'#16a34a':mCumAv>=83?'#d97706':'#dc2626';
  document.getElementById('kpiCumSub').textContent='Total lineas: '+aggL.toLocaleString('es-AR');
  document.getElementById('kpiRC').textContent=aggRC;
  document.getElementById('kpiRC').style.color=perfColorRec(aggRC);
  document.getElementById('kpiRate').textContent=mRateG;
  document.getElementById('kpiU').textContent=aggU.toLocaleString('es-AR');
  document.getElementById('kpiOlas2').textContent=aggOlas;
  document.getElementById('kpiStaffNec').textContent=mSN;
  if(selOp==='all') document.getElementById('kpiStaffNec').style.color=(mSA>=mSN)?'#16a34a':'#d97706';
  document.getElementById('kpiStaffAct').textContent=mSA;
  document.getElementById('kpiL').textContent=aggL.toLocaleString('es-AR');

  // Titulo del periodo
  var periodoLabel='';
  if(anio==='all'&&mes==='all') periodoLabel='Todos los periodos';
  else if(anio!=='all'&&mes==='all') periodoLabel=anio;
  else if(anio==='all'&&mes!=='all') periodoLabel=MES[parseInt(mes)];
  else periodoLabel=MES[parseInt(mes)]+' '+anio;
  document.getElementById('rankTitle').textContent='Ranking '+periodoLabel+' — Picking Regular (SIN/CON LOGO)';
  document.getElementById('rankChartTitle').textContent='Lineas/Dia por Operario — '+periodoLabel;

  // Ranking table
  var pickers=Object.entries(rankMap).map(function(e){
    var resp=e[0],p=e[1];
    var ld=p.dias?Math.round(p.lineas/p.dias*10)/10:0;
    var cumpl=Math.round(ld/TARGET*1000)/10;
    var tasa=p.lineas>0?Math.round(p.recCnt/(p.lineas/1000)*100)/100:0;
    return {resp:resp,dias:p.dias,olas:p.olas,lineas:p.lineas,unidades:p.unidades,recCnt:p.recCnt,trend:p.trend,ld:ld,cumpl:cumpl,tasa:tasa};
  }).sort(function(a,b){return b.ld-a.ld;});

  var tbody=document.getElementById('rankBody');
  tbody.innerHTML='';
  pickers.forEach(function(p,i){
    var pos=i+1;
    var medal=pos===1?'&#127945;':pos===2?'&#129352;':pos===3?'&#129353;':String(pos);
    var ldC=perfColor(p.ld);
    var cumC=p.cumpl>=100?'#16a34a':p.cumpl>=83?'#d97706':'#dc2626';
    var rcC=perfColorRec(p.recCnt);
    var tasaC=p.tasa===0?'#16a34a':p.tasa<=3?'#d97706':'#dc2626';
    var tr='<tr>'
      +'<td style="text-align:center;font-weight:700">'+medal+'</td>'
      +'<td style="font-weight:600">'+p.resp+'</td>'
      +'<td style="text-align:right">'+p.dias+'</td>'
      +'<td style="text-align:right">'+p.olas+'</td>'
      +'<td style="text-align:right">'+p.lineas.toLocaleString("es-AR")+'</td>'
      +'<td style="text-align:center;font-weight:700;color:'+ldC+'">'+p.ld+'</td>'
      +'<td style="text-align:center;font-weight:700;color:'+cumC+'">'+p.cumpl+'%</td>'
      +'<td style="text-align:right">'+p.unidades.toLocaleString("es-AR")+'</td>'
      +'<td style="text-align:center;font-weight:700;color:'+rcC+'">'+p.recCnt+'</td>'
      +'<td style="text-align:center;font-weight:700;color:'+tasaC+'">'+p.tasa+'</td>'
      +'<td style="text-align:center">'+p.trend+'</td>'
      +'</tr>';
    tbody.innerHTML+=tr;
  });

  // Totales table (acumulado del periodo, ordenado por total lineas)
  var totalesArr=pickers.slice().sort(function(a,b){return b.lineas-a.lineas;});
  var totBody=document.getElementById('totalesBody');
  totBody.innerHTML='';
  totalesArr.forEach(function(p){
    var ldC=perfColor(p.ld);
    var cumC=p.cumpl>=100?'#16a34a':p.cumpl>=83?'#d97706':'#dc2626';
    var rcC=perfColorRec(p.recCnt);
    var tr='<tr>'
      +'<td style="font-weight:600">'+p.resp+'</td>'
      +'<td style="text-align:right">'+p.dias+'</td>'
      +'<td style="text-align:right">'+p.olas+'</td>'
      +'<td style="text-align:right">'+p.lineas.toLocaleString("es-AR")+'</td>'
      +'<td style="text-align:center;font-weight:700;color:'+ldC+'">'+p.ld+'</td>'
      +'<td style="text-align:center;font-weight:700;color:'+cumC+'">'+p.cumpl+'%</td>'
      +'<td style="text-align:right">'+p.unidades.toLocaleString("es-AR")+'</td>'
      +'<td style="text-align:center;font-weight:700;color:'+rcC+'">'+p.recCnt+'</td>'
      +'</tr>';
    totBody.innerHTML+=tr;
  });

  // chartRanking
  chartRanking.data.labels=pickers.map(function(p){return p.resp;});
  chartRanking.data.datasets[0].data=pickers.map(function(p){return p.ld;});
  chartRanking.data.datasets[0].backgroundColor=pickers.map(function(p){return perfColor(p.ld);});
  chartRanking.data.datasets[1].data=Array(pickers.length).fill(TARGET);
  chartRanking.update();

  // chartTeamTrend
  chartTeamTrend.data.labels=labels;
  chartTeamTrend.data.datasets[0].data=idx.map(function(i){return teamLD[i];});
  chartTeamTrend.data.datasets[0].pointBackgroundColor=idx.map(function(i){return teamColors[i];});
  chartTeamTrend.data.datasets[1].data=Array(labels.length).fill(TARGET);
  chartTeamTrend.update();

  // chartStaff
  chartStaff.data.labels=labels;
  chartStaff.data.datasets[0].data=idx.map(function(i){return staffNec[i];});
  chartStaff.data.datasets[1].data=idx.map(function(i){return staffAct[i];});
  chartStaff.update();

  // chartEvol: resaltar operario seleccionado, atenuar los demas
  chartEvol.data.labels=labels;
  var numDS=chartEvol.data.datasets.length-1;
  for(var di=0;di<numDS;di++){
    (function(d){
      chartEvol.data.datasets[d].data=idx.map(function(i){return pickerDataFull[d][i];});
      if(selOp==='all'){
        chartEvol.data.datasets[d].borderColor=pickerOrigColors[d];
        chartEvol.data.datasets[d].borderWidth=2;
        chartEvol.data.datasets[d].borderDash=[];
      } else if(chartEvol.data.datasets[d].label===selOp){
        chartEvol.data.datasets[d].borderColor=pickerOrigColors[d];
        chartEvol.data.datasets[d].borderWidth=3.5;
        chartEvol.data.datasets[d].borderDash=[];
      } else {
        chartEvol.data.datasets[d].borderColor='rgba(180,180,180,0.3)';
        chartEvol.data.datasets[d].borderWidth=1;
        chartEvol.data.datasets[d].borderDash=[4,4];
      }
    })(di);
  }
  chartEvol.data.datasets[numDS].data=Array(labels.length).fill(TARGET);
  chartEvol.update();

  // chartMerma eliminado

  // ===== PIE DE MAQUINA (= PEDIDO MERMA) SECTION =====
  var pieAggM1L=0,pieAggM2L=0,pieHasData=false;
  idx.forEach(function(i){
    if(merma1[i]||merma2[i]) pieHasData=true;
    if(merma1L[i]) pieAggM1L+=merma1L[i];
    if(merma2L[i]) pieAggM2L+=merma2L[i];
  });
  var pieM1LDVals=idx.map(function(i){return merma1[i];}).filter(function(v){return v!==null&&v!==undefined;});
  var pieM2LDVals=idx.map(function(i){return merma2[i];}).filter(function(v){return v!==null&&v!==undefined;});
  var pieAvgM1LD=pieM1LDVals.length?Math.round(pieM1LDVals.reduce(function(s,v){return s+v;},0)/pieM1LDVals.length*10)/10:0;
  var pieAvgM2LD=pieM2LDVals.length?Math.round(pieM2LDVals.reduce(function(s,v){return s+v;},0)/pieM2LDVals.length*10)/10:0;
  if(pieHasData){
    document.getElementById('pieNoData').style.display='none';
    document.getElementById('pieM1LD').textContent=pieAvgM1LD||'-';
    document.getElementById('pieM2LD').textContent=pieAvgM2LD||'-';
    document.getElementById('pieM1L').textContent=pieAggM1L.toLocaleString('es-AR');
    document.getElementById('pieM2L').textContent=pieAggM2L.toLocaleString('es-AR');
  } else {
    document.getElementById('pieNoData').style.display='block';
    document.getElementById('pieM1LD').textContent='-';
    document.getElementById('pieM2LD').textContent='-';
    document.getElementById('pieM1L').textContent='-';
    document.getElementById('pieM2L').textContent='-';
  }
  chartPie.data.labels=labels;
  chartPie.data.datasets[0].data=idx.map(function(i){return merma1[i];});
  chartPie.data.datasets[1].data=idx.map(function(i){return merma2[i];});
  chartPie.update();

  // ===== MUESTRA SIMPLE SECTION =====
  var lezAggL=0,lezAggU=0,lezAggDias=0,lezAggOlas=0,lezAggRC=0,lezLDSum=0,lezVsSum=0,lezValidM=0;
  selKeys.forEach(function(k){
    var d=lezMonthly[k];
    if(!d) return;
    lezValidM++;
    lezAggL+=d.lineas; lezAggU+=d.unidades; lezAggDias+=d.dias; lezAggOlas+=d.olas; lezAggRC+=d.recCnt;
    lezLDSum+=d.ld; lezVsSum+=d.vsTeam;
  });
  var lezAvgLD=lezValidM?Math.round(lezLDSum/lezValidM*10)/10:0;
  var lezVsTeam=lezValidM?Math.round(lezVsSum/lezValidM*10)/10:0;
  var vsSign=lezVsTeam>0?'+':'';
  document.getElementById('lezLD').textContent=lezValidM?lezAvgLD:'-';
  document.getElementById('lezL').textContent=lezValidM?lezAggL.toLocaleString('es-AR'):'-';
  document.getElementById('lezDias').textContent=lezValidM?lezAggDias:'-';
  document.getElementById('lezU').textContent=lezValidM?lezAggU.toLocaleString('es-AR'):'-';
  document.getElementById('lezRC').textContent=lezValidM?lezAggRC:'-';
  document.getElementById('lezOlas').textContent=lezValidM?lezAggOlas:'-';
  document.getElementById('lezVsTeam').textContent=lezValidM?(vsSign+lezVsTeam+' lin/dia'):'-';
  document.getElementById('lezVsTeam').style.color=lezVsTeam>=0?'#16a34a':'#dc2626';
  chartLez.data.labels=labels;
  chartLez.data.datasets[0].data=idx.map(function(i){var k=allMonKeys[i];return lezMonthly[k]?lezMonthly[k].ld:null;});
  chartLez.update();

  // ===== RECLAMOS SECTION =====
  var recAggTotal=0, recSinIdTotal=0;
  var recOpMap={};
  var recCatMap={};
  selKeys.forEach(function(k){
    var d=recMonthly[k];
    if(!d) return;
    recSinIdTotal+=(d.byOp['Sin identificar']||0);
    if(selOp!=='all'){
      var opCnt=d.byOp[selOp]||0;
      recAggTotal+=opCnt;
      recOpMap[selOp]=(recOpMap[selOp]||0)+opCnt;
      // Use byOpCat for accurate per-operator categories
      if(d.byOpCat&&d.byOpCat[selOp]){
        Object.keys(d.byOpCat[selOp]).forEach(function(cat){recCatMap[cat]=(recCatMap[cat]||0)+d.byOpCat[selOp][cat];});
      }
    } else {
      recAggTotal+=d.cnt;
      Object.keys(d.byOp).forEach(function(op){recOpMap[op]=(recOpMap[op]||0)+d.byOp[op];});
      Object.keys(d.byCat).forEach(function(cat){recCatMap[cat]=(recCatMap[cat]||0)+d.byCat[cat];});
    }
  });

  // Pre-compute top category per operator from byOpCat (for ranking table)
  var opTopCatCounts={};
  selKeys.forEach(function(k){
    var d=recMonthly[k];
    if(!d||!d.byOpCat) return;
    Object.keys(d.byOpCat).forEach(function(op){
      if(!opTopCatCounts[op]) opTopCatCounts[op]={};
      Object.keys(d.byOpCat[op]).forEach(function(cat){
        opTopCatCounts[op][cat]=(opTopCatCounts[op][cat]||0)+d.byOpCat[op][cat];
      });
    });
  });
  var opBestCat={};
  Object.keys(opTopCatCounts).forEach(function(op){
    var best='-',bestN=0;
    Object.keys(opTopCatCounts[op]).forEach(function(cat){if(opTopCatCounts[op][cat]>bestN){bestN=opTopCatCounts[op][cat];best=cat;}});
    opBestCat[op]=best;
  });

  // Top categoría y top operario
  var topCat='-',topCatCnt=0;
  Object.keys(recCatMap).forEach(function(c){if(recCatMap[c]>topCatCnt){topCatCnt=recCatMap[c];topCat=c;}});
  var topOp='-',topOpCnt=0;
  Object.keys(recOpMap).forEach(function(o){if(o!=='Sin identificar'&&recOpMap[o]>topOpCnt){topOpCnt=recOpMap[o];topOp=o;}});

  // Tasa = reclamos / lineas picking * 1000
  var recTasa=aggL>0?Math.round(recAggTotal/(aggL/1000)*100)/100:0;

  document.getElementById('recTotal').textContent=recAggTotal;
  document.getElementById('recTotal').style.color=recAggTotal===0?'#16a34a':recAggTotal<=5?'#d97706':'#dc2626';
  document.getElementById('recTasa').textContent=recTasa.toFixed(2);
  document.getElementById('recTasa').style.color=recTasa===0?'#16a34a':recTasa<=3?'#d97706':'#dc2626';
  document.getElementById('recSinId').textContent=recSinIdTotal;
  document.getElementById('recSinId').style.color=recSinIdTotal===0?'#16a34a':'#d97706';
  document.getElementById('recTopCat').textContent=topCat;
  document.getElementById('recTopCatCnt').textContent=topCatCnt>0?(topCatCnt+' reclamos'):'Sin datos';
  document.getElementById('recTopOp').textContent=topOp;
  document.getElementById('recTopOpCnt').textContent=topOpCnt>0?(topOpCnt+' reclamos'):'Sin datos';

  // chartRecMon
  chartRecMon.data.labels=labels;
  chartRecMon.data.datasets[0].data=idx.map(function(i){
    var k=allMonKeys[i];
    var d=recMonthly[k];
    if(!d) return 0;
    if(selOp!=='all') return d.byOp[selOp]||0;
    return d.cnt;
  });
  chartRecMon.update();

  // chartRecCat (horizontal bar)
  var catLabels=Object.keys(recCatMap).sort(function(a,b){return recCatMap[b]-recCatMap[a];}).slice(0,10);
  var catData=catLabels.map(function(c){return recCatMap[c];});
  var catColors=['#dc2626','#d97706','#2563eb','#7c3aed','#0891b2','#16a34a','#ea580c','#db2777','#65a30d','#0369a1'];
  chartRecCat.data.labels=catLabels;
  chartRecCat.data.datasets[0].data=catData;
  chartRecCat.data.datasets[0].backgroundColor=catColors.slice(0,catLabels.length);
  chartRecCat.update();

  // chartRecOp
  var opArr=Object.entries(recOpMap).sort(function(a,b){return b[1]-a[1];}).slice(0,15);
  var opLabels=opArr.map(function(e){return e[0];});
  var opData=opArr.map(function(e){return e[1];});
  var opColors=opData.map(function(v){return v===0?'#16a34a':v<=3?'#d97706':'#dc2626';});
  chartRecOp.data.labels=opLabels;
  chartRecOp.data.datasets[0].data=opData;
  chartRecOp.data.datasets[0].backgroundColor=opColors;
  document.getElementById('recOpChartTitle').textContent='Reclamos por Operario'+(selOp!=='all'?' — '+selOp:'');
  chartRecOp.update();

  // Operator ranking table
  var recRankArr=Object.entries(recOpMap).sort(function(a,b){return b[1]-a[1];});
  var recBody=document.getElementById('recRankBody');
  recBody.innerHTML='';
  recRankArr.forEach(function(e,i){
    var op=e[0],cnt=e[1];
    var opL=rankMap[op]?rankMap[op].lineas:0;
    var opTasa=opL>0?Math.round(cnt/(opL/1000)*100)/100:0;
    var rcC=cnt===0?'#16a34a':cnt<=3?'#d97706':'#dc2626';
    var taC=opTasa===0?'#16a34a':opTasa<=3?'#d97706':'#dc2626';
    var bestCat=opBestCat[op]||'-';
    var medal=i===0?'&#127945;':i===1?'&#129352;':i===2?'&#129353;':(i+1)+'';
    var tr='<tr>'
      +'<td style="text-align:center;font-weight:700">'+medal+'</td>'
      +'<td style="font-weight:600">'+op+'</td>'
      +'<td style="text-align:right;font-weight:700;color:'+rcC+'">'+cnt+'</td>'
      +'<td style="text-align:right;color:'+taC+'">'+opTasa.toFixed(2)+'</td>'
      +'<td style="color:#555">'+bestCat+'</td>'
      +'</tr>';
    recBody.innerHTML+=tr;
  });

  // Category ranking table
  var recCatArr=Object.entries(recCatMap).sort(function(a,b){return b[1]-a[1];});
  var recCatBodyEl=document.getElementById('recCatBody');
  recCatBodyEl.innerHTML='';
  recCatArr.forEach(function(e,i){
    var cat=e[0],cnt=e[1];
    var pct=recAggTotal>0?Math.round(cnt/recAggTotal*1000)/10:0;
    var cC=i===0?'#dc2626':i<=2?'#d97706':'#555555';
    var tr='<tr>'
      +'<td style="text-align:center;font-weight:700;color:'+cC+'">'+(i+1)+'</td>'
      +'<td style="font-weight:600">'+cat+'</td>'
      +'<td style="text-align:right;font-weight:700;color:'+cC+'">'+cnt+'</td>'
      +'<td style="text-align:right">'+pct+'%</td>'
      +'</tr>';
    recCatBodyEl.innerHTML+=tr;
  });
}

function resetFilter(){
  document.getElementById('selAnio').value='all';
  document.getElementById('selMes').value='all';
  document.getElementById('selOp').value='all';
  applyFilter();
}

// Datos de pickers para chartEvol dinamico
const pickerDataFull=[];
const pickerOrigColors=[];

// Inicializar graficos con datos vacios (applyFilter() los llenara)
const chartRanking=new Chart('chartRanking',{type:'bar',data:{
  labels:[],datasets:[
    {label:'Lineas/Dia',data:[],backgroundColor:[],borderRadius:5},
    {label:'Target $TARGET',data:[],type:'line',borderColor:'#1a1a2e',borderWidth:1.5,borderDash:[6,4],pointRadius:0,fill:false}
  ]},options:{responsive:true,maintainAspectRatio:false,
  plugins:{legend:{position:'bottom',labels:{boxWidth:12,font:{size:11},color:'#444444'}}},
  scales:{y:{min:0,grid:{color:'#eeeeee'},ticks:{color:'#666666'}},x:{grid:{display:false},ticks:{color:'#666666',font:{size:9},maxRotation:35}}}}});

const chartTeamTrend=new Chart('chartTeamTrend',{type:'line',data:{
  labels:[$jsMonLabels],datasets:[
    {label:'Lin/Dia equipo',data:[$jsTeamLD],borderColor:'#2563eb',backgroundColor:'rgba(37,99,235,.1)',borderWidth:2.5,tension:.3,fill:true,pointRadius:4,pointBackgroundColor:[$jsTeamPtClrs]},
    {label:'Target $TARGET',data:Array($nMons).fill($TARGET),borderColor:'#16a34a',borderWidth:1.5,borderDash:[6,4],pointRadius:0,fill:false}
  ]},options:{responsive:true,maintainAspectRatio:false,
  plugins:{legend:{position:'bottom',labels:{boxWidth:12,font:{size:11},color:'#444444'}}},
  scales:{y:{min:0,grid:{color:'#eeeeee'},ticks:{color:'#666666'}},x:{grid:{display:false},ticks:{color:'#666666'}}}}});

const chartStaff=new Chart('chartStaff',{type:'bar',data:{
  labels:[$jsMonLabels],datasets:[
    {label:'Necesarias',data:[$jsStaffNec],backgroundColor:'rgba(68,114,196,.8)',borderRadius:4},
    {label:'Activas (prom/dia)',data:[$jsStaffAct],backgroundColor:'rgba(22,163,74,.7)',borderRadius:4}
  ]},options:{responsive:true,maintainAspectRatio:false,
  plugins:{legend:{position:'bottom',labels:{boxWidth:12,font:{size:11},color:'#444444'}}},
  scales:{y:{min:0,grid:{color:'#eeeeee'},ticks:{stepSize:1,color:'#666666'}},x:{grid:{display:false},ticks:{color:'#666666'}}}}});

const chartEvol=new Chart('chartEvol',{type:'line',data:{
  labels:[$jsMonLabels],datasets:[
$pickersJs
    {label:'Target $TARGET',data:Array($nMons).fill($TARGET),borderColor:'#1a1a2e',borderWidth:2,borderDash:[8,4],pointRadius:0,fill:false}
  ]},options:{responsive:true,maintainAspectRatio:false,
  plugins:{legend:{position:'bottom',labels:{boxWidth:10,font:{size:10},padding:8,color:'#444444'}}},
  scales:{y:{min:0,grid:{color:'#eeeeee'},ticks:{color:'#666666'}},x:{grid:{display:false},ticks:{color:'#666666'}}}}});

chartEvol.data.datasets.slice(0,-1).forEach(function(ds){pickerDataFull.push(ds.data.slice());pickerOrigColors.push(ds.borderColor);});

const chartEvol30=new Chart('chartEvol30',{type:'line',data:{
  labels:evol30Labels,datasets:[
    {label:'Total lineas equipo',data:evol30Data,borderColor:'#2563eb',backgroundColor:'rgba(37,99,235,.08)',borderWidth:2.5,tension:.3,fill:true,pointRadius:3,pointHoverRadius:5},
    {label:'Target',data:evol30Target,borderColor:'rgba(220,38,38,.65)',borderWidth:1.5,borderDash:[7,4],pointRadius:0,fill:false,tension:0}
  ]},options:{responsive:true,maintainAspectRatio:false,
  plugins:{legend:{display:true,position:'bottom',labels:{boxWidth:12,font:{size:10},color:'#444444'}},tooltip:{callbacks:{label:function(ctx){return ctx.dataset.label+': '+ctx.parsed.y.toLocaleString('es-AR')+(ctx.datasetIndex===0?' lineas':'');}}}},
  scales:{y:{min:0,grid:{color:'#eeeeee'},ticks:{color:'#666666'}},x:{grid:{display:false},ticks:{color:'#666666',font:{size:9},maxTicksLimit:15}}}}});

function setGrpFilter(g){
  _grpFilter=g;
  // Botones
  document.getElementById('grpBtnAll').classList.toggle('active',g==='all');
  document.getElementById('grpBtnCL').classList.toggle('active',g==='cl');
  document.getElementById('grpBtnSL').classList.toggle('active',g==='sl');
  // Datos del chart
  var data=g==='cl'?evol30CL:g==='sl'?evol30SL:evol30Data;
  var col=g==='cl'?'#2563eb':g==='sl'?'#16a34a':'#2563eb';
  var bg=g==='cl'?'rgba(37,99,235,.08)':g==='sl'?'rgba(22,163,74,.08)':'rgba(37,99,235,.08)';
  chartEvol30.data.datasets[0].data=data;
  chartEvol30.data.datasets[0].borderColor=col;
  chartEvol30.data.datasets[0].backgroundColor=bg;
  // dataset[1] = target line, no se toca
  chartEvol30.update();
  // Titulo
  var lbl=g==='cl'?'Con Logo':g==='sl'?'Sin Logo':'Con Logo + Sin Logo';
  document.getElementById('evol30Title').textContent='📈 Evolución diaria del equipo — últimos 30 días hábiles';
  document.getElementById('evol30Sub').textContent='Total líneas picking regular ('+lbl+') — equipo completo';
  // Rebuild tabla 7 dias con filtro
  _build7dTable();
}

function _build7dTable(){
  var b=document.getElementById('body7d');
  var f=document.getElementById('foot7d');
  if(!day7Rows||!day7Rows.length){
    b.innerHTML='<tr><td colspan="8" style="text-align:center;color:#aaa;padding:20px">Sin datos disponibles</td></tr>';
    return;
  }
  b.innerHTML='';
  var g=_grpFilter;
  day7Rows.forEach(function(r){
    var lines=g==='cl'?r.cl:g==='sl'?r.sl:r.lines;
    var ops=g==='cl'?r.clops:g==='sl'?r.slops:r.ops;
    var lpo=ops?Math.round(lines/ops*10)/10:0;
    var tgt=ops*$TARGET;
    var cum=tgt?Math.round(lines/tgt*1000)/10:0;
    var cumC=cum>=100?'#16a34a':cum>=TARGET_WARN?'#d97706':'#dc2626';
    var dltStr=g!=='all'?'&mdash;':(r.dlt===null?'&mdash;':r.dlt>0?'<span style="color:#16a34a;font-weight:700">+'+r.dlt+'</span>':r.dlt<0?'<span style="color:#dc2626;font-weight:700">'+r.dlt+'</span>':'<span style="color:#999">0</span>');
    b.innerHTML+='<tr>'
      +'<td><strong>'+r.f+'</strong></td>'
      +'<td>'+r.d+'</td>'
      +'<td style="text-align:right">'+ops+'</td>'
      +'<td style="text-align:right"><strong>'+lines.toLocaleString('es-AR')+'</strong></td>'
      +'<td style="text-align:right">'+lpo+'</td>'
      +'<td style="text-align:right">'+(tgt||'&mdash;')+'</td>'
      +'<td style="text-align:center;font-weight:700;color:'+cumC+'">'+cum+'%</td>'
      +'<td style="text-align:center">'+dltStr+'</td>'
      +'</tr>';
  });
  var aL=Math.round(day7Rows.reduce(function(s,r){return s+(g==='cl'?r.cl:g==='sl'?r.sl:r.lines);},0)/day7Rows.length);
  var aO=(day7Rows.reduce(function(s,r){return s+(g==='cl'?r.clops:g==='sl'?r.slops:r.ops);},0)/day7Rows.length).toFixed(1);
  var aP=parseFloat(aO)?Math.round(aL/parseFloat(aO)*10)/10:0;
  var aC=(day7Rows.reduce(function(s,r){
    var li=g==='cl'?r.cl:g==='sl'?r.sl:r.lines;
    var op=g==='cl'?r.clops:g==='sl'?r.slops:r.ops;
    var t=op*$TARGET;
    return s+(t?li/t*100:0);
  },0)/day7Rows.length).toFixed(1);
  var acC=parseFloat(aC)>=100?'#16a34a':parseFloat(aC)>=TARGET_WARN?'#d97706':'#dc2626';
  f.innerHTML='<tr>'
    +'<td colspan="2" style="padding:9px 11px">PROM 7 D&Iacute;AS</td>'
    +'<td style="text-align:right;padding:9px 11px">'+aO+'</td>'
    +'<td style="text-align:right;padding:9px 11px">'+aL.toLocaleString('es-AR')+'</td>'
    +'<td style="text-align:right;padding:9px 11px">'+aP+'</td>'
    +'<td style="text-align:right;padding:9px 11px">&mdash;</td>'
    +'<td style="text-align:center;padding:9px 11px;color:'+acC+'">'+aC+'%</td>'
    +'<td style="text-align:center;padding:9px 11px">&mdash;</td>'
    +'</tr>';
}

const chartMerma=null; // canvas eliminado — datos de Pedido Merma no se muestran

const chartPie=new Chart('chartPie',{type:'line',data:{
  labels:[$jsMonLabels],datasets:[
    {label:'Turno 1 (06-15h)',data:[$jsMerma1],borderColor:'#2563eb',backgroundColor:'rgba(37,99,235,.12)',borderWidth:2.5,tension:.3,pointRadius:5,fill:true,spanGaps:true},
    {label:'Turno 2 (resto)',data:[$jsMerma2],borderColor:'#ea580c',backgroundColor:'rgba(234,88,12,.12)',borderWidth:2.5,tension:.3,pointRadius:5,fill:true,spanGaps:true}
  ]},options:{responsive:true,maintainAspectRatio:false,
  plugins:{legend:{position:'bottom',labels:{boxWidth:12,font:{size:11},color:'#444444'}}},
  scales:{y:{min:0,grid:{color:'#eeeeee'},ticks:{color:'#666666'}},x:{grid:{display:false},ticks:{color:'#666666'}}}}});

const chartLez=new Chart('chartLez',{type:'line',data:{
  labels:[$jsMonLabels],datasets:[
    {label:'Lin/Dia Lezcano',data:[$jsLezLD],borderColor:'#7c3aed',backgroundColor:'rgba(124,58,237,.1)',borderWidth:2.5,tension:.3,pointRadius:5,fill:true,spanGaps:true}
  ]},options:{responsive:true,maintainAspectRatio:false,
  plugins:{legend:{position:'bottom',labels:{boxWidth:12,font:{size:11},color:'#444444'}}},
  scales:{y:{min:0,grid:{color:'#eeeeee'},ticks:{color:'#666666'}},x:{grid:{display:false},ticks:{color:'#666666'}}}}});

const chartRecMon=new Chart('chartRecMon',{type:'bar',data:{
  labels:[],datasets:[
    {label:'Reclamos',data:[],backgroundColor:'rgba(220,38,38,.7)',borderColor:'#dc2626',borderWidth:1,borderRadius:4}
  ]},options:{responsive:true,maintainAspectRatio:false,
  plugins:{legend:{display:false}},
  scales:{y:{min:0,grid:{color:'#eeeeee'},ticks:{stepSize:1,color:'#666666'}},x:{grid:{display:false},ticks:{color:'#666666'}}}}});

const chartRecCat=new Chart('chartRecCat',{type:'bar',data:{
  labels:[],datasets:[{data:[],backgroundColor:[],borderRadius:4,borderSkipped:false}]},
  options:{indexAxis:'y',responsive:true,maintainAspectRatio:false,
  plugins:{legend:{display:false}},
  scales:{x:{min:0,grid:{color:'#eeeeee'},ticks:{stepSize:1,color:'#666666'}},y:{grid:{display:false},ticks:{color:'#444444',font:{size:11}}}}}});

const chartRecOp=new Chart('chartRecOp',{type:'bar',data:{
  labels:[],datasets:[
    {label:'Reclamos',data:[],backgroundColor:[],borderRadius:4}
  ]},options:{responsive:true,maintainAspectRatio:false,
  plugins:{legend:{display:false}},
  scales:{y:{min:0,grid:{color:'#eeeeee'},ticks:{stepSize:1,color:'#666666'}},x:{grid:{display:false},ticks:{color:'#666666',font:{size:10},maxRotation:35}}}}});

// ===== DARK / LIGHT THEME =====
var _allCharts=[chartRanking,chartTeamTrend,chartStaff,chartEvol,chartMerma,chartPie,chartLez,chartRecMon,chartRecCat,chartRecOp,chartEvol30];
function updateChartColors(dark){
  var grid=dark?'#334155':'#eeeeee';
  var tick=dark?'#94a3b8':'#666666';
  var leg=dark?'#94a3b8':'#444444';
  Chart.defaults.color=dark?'#94a3b8':'#666666';
  _allCharts.forEach(function(c){
    if(!c||!c.options) return;
    if(c.options.scales){
      Object.keys(c.options.scales).forEach(function(k){
        var ax=c.options.scales[k];
        if(ax.grid) ax.grid.color=grid;
        if(ax.ticks) ax.ticks.color=tick;
      });
    }
    if(c.options.plugins&&c.options.plugins.legend&&c.options.plugins.legend.labels)
      c.options.plugins.legend.labels.color=leg;
    c.update('none');
  });
}
function toggleTheme(){
  var dark=document.body.classList.toggle('dark');
  localStorage.setItem('zecat-theme',dark?'dark':'light');
  document.getElementById('themeIcon').innerHTML=dark?'&#9728;':'&#9790;';
  updateChartColors(dark);
}
// Aplicar tema guardado al cargar (antes del try de init)
(function(){
  if(localStorage.getItem('zecat-theme')==='dark'){
    document.body.classList.add('dark');
    document.getElementById('themeIcon').innerHTML='&#9728;';
    // Los charts se actualizan después del init
  }
})();

// Inicializar tab y filtro al ultimo mes disponible
try {
  switchTab('picking');
  document.getElementById('selAnio').value='$jsInitYear';
  document.getElementById('selMes').value='$jsInitMon';
  applyFilter();
  // Poblar tabla 7 dias (datos estaticos, filtrable por grupo)
  _build7dTable();
  // Si ya hay dark mode guardado, aplicar colores a los charts recién creados
  if(document.body.classList.contains('dark')) updateChartColors(true);
} catch(err) {
  var errDiv=document.createElement('div');
  errDiv.style.cssText='position:fixed;top:0;left:0;right:0;background:#dc2626;color:white;padding:12px 20px;font-size:13px;font-family:monospace;z-index:9999;white-space:pre-wrap;';
  errDiv.textContent='ERROR JS: '+err.message+'\n'+err.stack;
  document.body.prepend(errDiv);
}
</script>
</body></html>
"@

    # Incrustar Chart.js inline para que el HTML sea 100% self-contained
    $chartJsFile = Join-Path (Split-Path $PSCommandPath) "chartjs.min.js"
    if(Test-Path $chartJsFile){
        $chartJsInline = [System.IO.File]::ReadAllText($chartJsFile, [System.Text.Encoding]::UTF8)
        $html = $html.Replace(
            '<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>',
            "<script>$chartJsInline</script>"
        )
    }
    $html | Out-File -FilePath $htmlOut -Encoding utf8 -NoNewline
    Write-Host "[$($NOW.ToString('HH:mm:ss'))] HTML guardado -> $htmlOut"

    # ===========================================================
    # GUARDAR EXCEL
    # ===========================================================
    $dashWb.Sheets.Item("RESUMEN").Activate()
    $savedPath=$Output
    try {
        if(Test-Path $Output){ Remove-Item $Output -Force -ErrorAction Stop }
        $dashWb.SaveAs($Output,51)
    } catch {
        $savedPath=$Output -replace '\.xlsx$',"_$(Get-Date -Format 'HHmmss').xlsx"
        try{$dashWb.SaveAs($savedPath,51)}catch{Write-Warning "No se pudo guardar Excel: $_"}
    }
    $dashWb.Close($false)
    Write-Host "[$($NOW.ToString('HH:mm:ss'))] Dashboard guardado -> $savedPath"
    Write-Host "[$($NOW.ToString('HH:mm:ss'))] COMPLETADO OK - v5"

} catch {
    Write-Error "ERROR: $_"
    Write-Error $_.ScriptStackTrace
} finally {
    try{$xl.Quit()}catch{}
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)|Out-Null
    [System.GC]::Collect()
}
