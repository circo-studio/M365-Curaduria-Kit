#Requires -Modules PnP.PowerShell
<#
===============================================================================
  Script_A_VaciarPapelera.ps1  |  Circo Studio  |  Kit de Curaduría M365
===============================================================================

  *** SCRIPT DESTRUCTIVO — BORRA CONTENIDO DE FORMA PERMANENTE ***

  A diferencia del resto de los scripts A, este MODIFICA el tenant.
  Lo borrado no se puede recuperar: no hay tercera etapa de papelera.
  Requiere validación explícita del owner del sitio antes de usarse.

PROPÓSITO:
  Utilitario interactivo para medir y vaciar la papelera de reciclaje de un
  sitio. Pensado para papeleras grandes (cientos de GB), donde la UI de
  SharePoint tira timeout y no llega a completar la operación.

RENDIMIENTO:
  El borrado usa "Clear-PnPRecycleBinItem -RowLimit N", que borra hasta N
  ítems por llamada del lado del servidor. Es lo más rápido disponible:
  una llamada cada 1000 ítems en lugar de una por ítem.

  Limitaciones del cmdlet, que definen cuándo se puede usar el modo rápido:
    - "-All" borra ambas etapas; "-SecondStageOnly" borra solo la segunda.
    - NO existe un modo "solo primera etapa": vaciar únicamente la 1ª etapa
      obliga a borrado individual.
    - El borrado por lotes no permite filtrar por fecha ni por tipo, así que
      "-BorradosAntesDe" también cae a borrado individual.

  En modo rápido el LOG registra una fila por lote (el detalle ítem por ítem
  queda en el CSV de inventario previo). En modo individual registra una fila
  por ítem.

USO:
  # Interactivo: pide la URL y qué etapa vaciar
  .\Script_A_VaciarPapelera.ps1 -ConfigPath "..\..\Config\cliente_config.json"

  # Solo medir, sin ofrecer vaciado
  .\Script_A_VaciarPapelera.ps1 -ConfigPath ".\cliente_config.json" -SoloMedir

  # Vaciado directo de ambas etapas, lotes de 2000
  .\Script_A_VaciarPapelera.ps1 -ConfigPath ".\cliente_config.json" `
      -Url "https://tenant.sharepoint.com/sites/EJEMPLO" -Etapa Ambas -TamanoLote 2000

  # Solo lo borrado antes de una fecha (fuerza modo individual)
  .\Script_A_VaciarPapelera.ps1 -ConfigPath ".\cliente_config.json" -BorradosAntesDe "2024-01-01"

ARCHIVOS GENERADOS:
  [Sitio]_Papelera_[timestamp]_INVENTARIO.csv  — manifiesto previo al borrado
  [Sitio]_Papelera_[timestamp]_LOG.csv         — resultado del vaciado

NOTAS:
  - Vaciar la 1ª etapa con este script borra permanentemente: NO manda los
    ítems a la 2ª etapa. Para liberar cuota hay que vaciar ambas.
  - Los ítems alcanzados por políticas de retención se copian a la
    Preservation Hold Library y NO liberan cuota al borrarse.
  - La métrica de storage del tenant se actualiza diferido: el sitio puede
    seguir mostrando la cuota vieja por varias horas después del vaciado.
  - Interrumpir con Ctrl+C es seguro: lo ya borrado queda borrado, el resto
    sigue en la papelera y el script se puede volver a correr.

===============================================================================
#>

param(
    [Parameter(Mandatory, HelpMessage="Path al archivo cliente_config.json")]
    [string]$ConfigPath,

    # URL del sitio. Si se omite, se pide por consola
    [string]$Url,

    # Qué vaciar. Si se omite, se pregunta después de mostrar los volúmenes
    [ValidateSet("Primera","Segunda","Ambas")]
    [string]$Etapa,

    # Solo mide, no ofrece vaciar
    [switch]$SoloMedir,

    # Simula el vaciado: recorre los ítems pero no borra nada
    [switch]$Simular,

    # Borra únicamente los ítems eliminados antes de esta fecha (ej: "2024-01-01")
    # Fuerza el modo individual: el borrado por lotes no admite filtros
    [datetime]$BorradosAntesDe,

    # Ítems por llamada en modo rápido. Más alto = menos llamadas pero más
    # riesgo de timeout. Por encima de 5000 puede chocar con el list view threshold
    [ValidateRange(1,5000)]
    [int]$TamanoLote = 1000,

    # Fuerza borrado ítem por ítem (más lento, log con detalle completo)
    [switch]$BorradoIndividual,

    # Tope de ítems a leer al medir (protege contra papeleras gigantes)
    [int]$MaxItemsAnalisis = 200000,

    # Timeout de cada llamada al servidor, en segundos
    [int]$TimeoutSegundos = 600,

    # Salta la medición previa y va directo al vaciado (sin ETA ni inventario)
    [switch]$SinMedicion,

    # No genera el CSV de inventario previo (no recomendado)
    [switch]$SinInventario,

    # Saltea las confirmaciones interactivas
    [switch]$SinConfirmacion,

    [switch]$ForceAuth
)

$moduloPath = Join-Path $PSScriptRoot "M365Curaduria.psm1"
if (-not (Test-Path $moduloPath)) {
    Write-Host "[ERROR] No se encontró M365Curaduria.psm1 en: $moduloPath" -ForegroundColor Red
    exit 1
}
Import-Module $moduloPath -Force

# ============================================================
# CONFIGURACIÓN
# ============================================================
$config = $null
try { $config = Read-ClientConfig -ConfigPath $ConfigPath }
catch { Write-Host "[ERROR] $_" -ForegroundColor Red; exit 1 }

if (-not (Test-Path $config.CarpetaLocal)) {
    New-Item -ItemType Directory -Path $config.CarpetaLocal -Force | Out-Null
}

function ConvertTo-Gb {
    param([double]$Bytes)
    return [math]::Round($Bytes / 1GB, 2)
}

function Format-Duracion {
    param([double]$Segundos)
    if ($Segundos -le 0 -or [double]::IsInfinity($Segundos) -or [double]::IsNaN($Segundos)) { return "--" }
    $ts = [timespan]::FromSeconds($Segundos)
    if ($ts.TotalHours -ge 1)   { return "{0}h {1}m" -f [int]$ts.TotalHours, $ts.Minutes }
    if ($ts.TotalMinutes -ge 1) { return "{0}m {1}s" -f [int]$ts.TotalMinutes, $ts.Seconds }
    return "{0}s" -f [int]$ts.TotalSeconds
}

Write-Log "======================================================" -Level HEADER
Write-Log "  Kit M365 Curaduría — Vaciar Papelera"                 -Level HEADER
Write-Log "  Cliente : $($config.NombreCliente)"                   -Level HEADER
Write-Log "  *** SCRIPT DESTRUCTIVO — EL BORRADO ES PERMANENTE ***" -Level WARN
Write-Log "======================================================" -Level HEADER

# ============================================================
# URL DEL SITIO
# ============================================================
if (-not $Url) {
    Write-Log "" -Level INFO
    $Url = Read-Host "URL del sitio (ej: $($config.SharePointUrl)/sites/EJEMPLO)"
}

$Url = ([string]$Url).Trim().Trim('"').TrimEnd('/')

if (-not $Url) {
    Write-Log "No se ingresó una URL. Cancelado." -Level ERROR
    exit 1
}
if ($Url -notmatch '^https?://') {
    Write-Log "La URL debe empezar con https:// — recibido: $Url" -Level ERROR
    exit 1
}

$nombreSitio   = ($Url -split '/')[-1]
$nombreLimpio  = ($nombreSitio -replace "[^a-zA-Z0-9]", "_")
$timestamp     = Get-Date -Format "yyyyMMdd_HHmmss"
$csvInventario = "$($config.CarpetaLocal)\${nombreLimpio}_Papelera_${timestamp}_INVENTARIO.csv"
$csvLog        = "$($config.CarpetaLocal)\${nombreLimpio}_Papelera_${timestamp}_LOG.csv"

# ============================================================
# CONEXIÓN
# ============================================================
Write-Log "" -Level INFO
Write-Log "Conectando a $Url ..." -Level INFO

$conn = $null
try {
    $connectParams = @{
        Url         = $Url
        ClientId    = $config.ClientId
        Interactive = $true
    }
    if ($ForceAuth) { $connectParams["ForceAuthentication"] = $true }
    else            { $connectParams["PersistLogin"] = $true }

    $conn = Connect-PnPOnline @connectParams -ReturnConnection -ErrorAction Stop
    $web  = Get-PnPWeb -Connection $conn -ErrorAction Stop

    # Timeout acotado: si una llamada se cuelga conviene que falle y reintente
    try {
        $ctx = Get-PnPContext -Connection $conn
        if ($ctx) { $ctx.RequestTimeout = $TimeoutSegundos * 1000 }
    }
    catch { }

    Write-Log "Conexión OK: $($web.Title)" -Level OK
}
catch {
    Write-Log "No se pudo conectar: $($_.Exception.Message)" -Level ERROR
    Write-Log "Verificar que la cuenta sea Site Collection Admin del sitio." -Level ERROR
    exit 1
}

# --- ¿El cmdlet soporta borrado por lotes? ---
$soportaRowLimit = $false
try {
    $soportaRowLimit = (Get-Command Clear-PnPRecycleBinItem -ErrorAction Stop).Parameters.ContainsKey('RowLimit')
}
catch { $soportaRowLimit = $false }

if (-not $soportaRowLimit) {
    Write-Log "Esta versión de PnP no soporta -RowLimit en Clear-PnPRecycleBinItem." -Level WARN
    Write-Log "  Se usará borrado individual. Para acelerar: Update-Module PnP.PowerShell" -Level WARN
}

# ============================================================
# MEDICIÓN
# ============================================================
$todos    = @()
$truncado = $false
$gb1 = 0; $gb2 = 0; $items1 = 0; $items2 = 0

if (-not $SinMedicion) {

    Write-Log "" -Level INFO
    Write-Log "Midiendo la papelera (una sola enumeración, puede tardar)..." -Level INFO

    $swMedicion = [System.Diagnostics.Stopwatch]::StartNew()

    foreach ($et in @("Primera","Segunda")) {
        try {
            $items = if ($et -eq "Primera") {
                @(Get-PnPRecycleBinItem -FirstStage -RowLimit $MaxItemsAnalisis -Connection $conn -ErrorAction Stop)
            } else {
                @(Get-PnPRecycleBinItem -SecondStage -RowLimit $MaxItemsAnalisis -Connection $conn -ErrorAction Stop)
            }

            $suma = ($items | Measure-Object -Property Size -Sum).Sum
            if (-not $suma) { $suma = 0 }

            if ($et -eq "Primera") { $items1 = $items.Count; $gb1 = ConvertTo-Gb ([double]$suma) }
            else                   { $items2 = $items.Count; $gb2 = ConvertTo-Gb ([double]$suma) }

            if ($items.Count -ge $MaxItemsAnalisis) { $truncado = $true }

            $todos += $items
        }
        catch {
            Write-Log "  [WARN] No se pudo leer la $et etapa: $($_.Exception.Message -replace "`r`n|`n", ' ')" -Level WARN
        }
    }

    $swMedicion.Stop()

    Write-Log "" -Level INFO
    Write-Log "======================================================" -Level HEADER
    Write-Log "  PAPELERA DE $($web.Title)"                            -Level HEADER
    Write-Log "  1ª etapa (usuarios) : $items1 ítems — $gb1 GB"        -Level $(if ($gb1 -gt 0) {"WARN"} else {"OK"})
    Write-Log "  2ª etapa (admin)    : $items2 ítems — $gb2 GB"        -Level $(if ($gb2 -gt 0) {"WARN"} else {"OK"})
    Write-Log "  TOTAL               : $($items1 + $items2) ítems — $([math]::Round($gb1 + $gb2, 2)) GB" -Level HEADER
    Write-Log "  Medición en $(Format-Duracion $swMedicion.Elapsed.TotalSeconds)" -Level INFO
    Write-Log "======================================================" -Level HEADER

    if ($truncado) {
        Write-Log "[WARN] Enumeración truncada en $MaxItemsAnalisis ítems por etapa — hay más de lo informado." -Level WARN
    }

    if ($todos.Count -gt 0) {
        $fechas = @($todos | Where-Object { $_.DeletedDate } | Sort-Object DeletedDate)
        if ($fechas.Count -gt 0) {
            Write-Log "  Borrado más antiguo : $($fechas[0].DeletedDate.ToString('yyyy-MM-dd'))" -Level INFO
            Write-Log "  Borrado más reciente: $($fechas[-1].DeletedDate.ToString('yyyy-MM-dd'))" -Level INFO
        }

        $topBorradores = $todos | Where-Object { $_.DeletedByEmail } |
                         Group-Object DeletedByEmail |
                         Sort-Object { ($_.Group | Measure-Object Size -Sum).Sum } -Descending |
                         Select-Object -First 3

        if ($topBorradores) {
            Write-Log "  Quién borró más volumen:" -Level INFO
            foreach ($b in $topBorradores) {
                $gbB = ConvertTo-Gb ([double](($b.Group | Measure-Object Size -Sum).Sum))
                Write-Log "    $($b.Name) — $($b.Count) ítems, $gbB GB" -Level INFO
            }
        }
    }

    if (($items1 + $items2) -eq 0) {
        Write-Log "" -Level INFO
        Write-Log "La papelera está vacía. Nada que hacer." -Level OK
        exit 0
    }
}
else {
    Write-Log "" -Level INFO
    Write-Log "Medición omitida (-SinMedicion): sin ETA ni inventario previo." -Level WARN
}

if ($SoloMedir) {
    Write-Log "" -Level INFO
    Write-Log "Modo -SoloMedir: no se vacía nada." -Level OK
    exit 0
}

# ============================================================
# ELECCIÓN DE ETAPA
# ============================================================
if (-not $Etapa) {
    Write-Log "" -Level INFO
    Write-Host "  ¿Qué querés vaciar?" -ForegroundColor Cyan
    Write-Host "    1) Solo 1ª etapa  ($items1 ítems, $gb1 GB)  [lento: borrado individual]"
    Write-Host "    2) Solo 2ª etapa  ($items2 ítems, $gb2 GB)  [rápido]"
    Write-Host "    3) Ambas etapas   ($([math]::Round($gb1 + $gb2, 2)) GB)  [rápido]"
    Write-Host "    0) Cancelar"
    $opcion = Read-Host "Opción"

    switch ($opcion.Trim()) {
        "1" { $Etapa = "Primera" }
        "2" { $Etapa = "Segunda" }
        "3" { $Etapa = "Ambas" }
        default { Write-Log "Cancelado por el usuario." -Level WARN; exit 0 }
    }
}

$etapasAProcesar = switch ($Etapa) {
    "Primera" { @("Primera") }
    "Segunda" { @("Segunda") }
    "Ambas"   { @("Primera","Segunda") }
}

$gbElegido = 0
if ($etapasAProcesar -contains "Primera") { $gbElegido += $gb1 }
if ($etapasAProcesar -contains "Segunda") { $gbElegido += $gb2 }

# ============================================================
# MODO DE BORRADO
# ============================================================
$motivoIndividual = $null

if ($Simular)                { $motivoIndividual = "modo simulación" }
elseif ($BorradoIndividual)  { $motivoIndividual = "-BorradoIndividual" }
elseif (-not $soportaRowLimit) { $motivoIndividual = "la versión de PnP no soporta -RowLimit" }
elseif ($BorradosAntesDe)    { $motivoIndividual = "el filtro -BorradosAntesDe no existe en el borrado por lotes" }
elseif ($Etapa -eq "Primera"){ $motivoIndividual = "el cmdlet no tiene modo 'solo primera etapa'" }

$modoRapido = ($null -eq $motivoIndividual)

Write-Log "" -Level INFO
if ($modoRapido) {
    Write-Log "Modo: RÁPIDO — lotes de hasta $TamanoLote ítems por llamada" -Level OK
}
else {
    Write-Log "Modo: INDIVIDUAL (una llamada por ítem) — motivo: $motivoIndividual" -Level WARN
    if ($Etapa -eq "Primera" -and -not $BorradoIndividual -and -not $Simular) {
        Write-Log "  Sugerencia: vaciar 'Ambas' etapas usa el modo rápido y libera más cuota." -Level INFO
    }
}

if ($BorradosAntesDe) {
    Write-Log "Filtro activo: solo ítems borrados antes de $($BorradosAntesDe.ToString('yyyy-MM-dd'))" -Level WARN
}

# ============================================================
# INVENTARIO PREVIO
# ============================================================
if (-not $SinInventario -and $todos.Count -gt 0) {
    try {
        $swInv = [System.Diagnostics.Stopwatch]::StartNew()

        $todos | Select-Object `
            @{n='Etapa';         e={ if ($_.ItemState -like "*Second*") { "Segunda" } else { "Primera" } }},
            @{n='Titulo';        e={ $_.Title }},
            @{n='Tipo';          e={ $_.ItemType }},
            @{n='Tamano_MB';     e={ [math]::Round(([double]$_.Size)/1MB, 2) }},
            @{n='Ubicacion';     e={ $_.DirName }},
            @{n='Borrado_Por';   e={ $_.DeletedByEmail }},
            @{n='Fecha_Borrado'; e={ if ($_.DeletedDate) { $_.DeletedDate.ToString('yyyy-MM-dd HH:mm') } else { "" } }},
            @{n='Id';            e={ $_.Id }} |
            Export-Csv -Path $csvInventario -NoTypeInformation -Encoding UTF8

        $swInv.Stop()
        Write-Log "" -Level INFO
        Write-Log "Inventario previo guardado en $(Format-Duracion $swInv.Elapsed.TotalSeconds): $csvInventario" -Level OK
        Write-Log "  Es el único registro de lo que había. Guardarlo antes de continuar." -Level WARN
    }
    catch {
        Write-Log "No se pudo generar el inventario previo: $($_.Exception.Message)" -Level ERROR
        if (-not $SinConfirmacion) {
            $seguir = Read-Host "Continuar igual, sin inventario? (SI/NO)"
            if ($seguir.Trim().ToUpper() -ne "SI") { Write-Log "Cancelado." -Level WARN; exit 0 }
        }
    }
}

# ============================================================
# CONFIRMACIÓN
# ============================================================
if (-not $Simular -and -not $SinConfirmacion) {
    Write-Log "" -Level INFO
    Write-Log "======================================================" -Level WARN
    Write-Log "  Se va a BORRAR DE FORMA PERMANENTE:"                -Level WARN
    Write-Log "  Sitio  : $($web.Title)"                             -Level WARN
    Write-Log "  URL    : $Url"                                      -Level WARN
    Write-Log "  Etapas : $Etapa"                                    -Level WARN
    Write-Log "  Volumen: ~$gbElegido GB"                            -Level WARN
    Write-Log "  Esta acción NO se puede deshacer."                  -Level WARN
    Write-Log "======================================================" -Level WARN

    $confirma = Read-Host "Para confirmar, escribir el nombre del sitio: $nombreSitio"
    if ($confirma.Trim() -ne $nombreSitio) {
        Write-Log "El nombre no coincide. Cancelado." -Level WARN
        exit 0
    }
}

# ============================================================
# LOG CON BUFFER (se escribe por lote, no por ítem)
# ============================================================
"Etapa,Titulo,Tipo,Tamano_MB,Resultado,Detalle,Timestamp" |
    Out-File -FilePath $csvLog -Encoding UTF8

$script:bufferLog = [System.Collections.Generic.List[string]]::new()

function Add-FilaLog {
    param([string]$Etapa, [string]$Titulo, [string]$Tipo, [string]$MB, [string]$Resultado, [string]$Detalle)
    $t = (($Titulo  -replace '"', "'") -replace "`r`n|`n", " ")
    $d = (($Detalle -replace '"', "'") -replace "`r`n|`n", " ")
    $script:bufferLog.Add("`"$Etapa`",`"$t`",`"$Tipo`",`"$MB`",`"$Resultado`",`"$d`",`"$(Get-Date -Format 'HH:mm:ss')`"")
}

function Flush-Log {
    if ($script:bufferLog.Count -eq 0) { return }
    # Reintento: si el CSV está abierto en Excel, el buffer espera al próximo flush
    for ($i = 1; $i -le 3; $i++) {
        try {
            Add-Content -Path $script:csvLog -Value $script:bufferLog -Encoding UTF8 -ErrorAction Stop
            $script:bufferLog.Clear()
            return
        }
        catch { Start-Sleep -Milliseconds 400 }
    }
    Write-Log "  [WARN] No se pudo escribir el log (¿está abierto en Excel?) — se reintenta en el próximo lote" -Level WARN
}

# ============================================================
# LISTA DE TRABAJO
# ============================================================
$script:borrados       = 0
$script:errores        = 0
$script:saltados       = 0
$script:bytesLiberados = [double]0
$script:sw             = [System.Diagnostics.Stopwatch]::StartNew()
$script:ultimoReporte  = [datetime]::Now.AddSeconds(-60)
$script:ultimaBarra    = [datetime]::Now.AddSeconds(-60)

$trabajo = @()

if ($todos.Count -gt 0) {
    $trabajo = @($todos | Where-Object {
        $et = if ($_.ItemState -like "*Second*") { "Segunda" } else { "Primera" }
        $etapasAProcesar -contains $et
    })
}
elseif (-not $modoRapido) {
    # El modo individual necesita la lista; el rápido no
    Write-Log "" -Level INFO
    Write-Log "Enumerando la papelera..." -Level INFO
    foreach ($et in $etapasAProcesar) {
        try {
            $trabajo += if ($et -eq "Primera") {
                @(Get-PnPRecycleBinItem -FirstStage -RowLimit $MaxItemsAnalisis -Connection $conn -ErrorAction Stop)
            } else {
                @(Get-PnPRecycleBinItem -SecondStage -RowLimit $MaxItemsAnalisis -Connection $conn -ErrorAction Stop)
            }
        }
        catch {
            Write-Log "  [WARN] No se pudo leer la $et etapa: $($_.Exception.Message -replace "`r`n|`n", ' ')" -Level WARN
        }
    }
}

if ($BorradosAntesDe -and $trabajo.Count -gt 0) {
    $antes = $trabajo.Count
    $trabajo = @($trabajo | Where-Object { -not $_.DeletedDate -or $_.DeletedDate -lt $BorradosAntesDe })
    $script:saltados += ($antes - $trabajo.Count)
    Write-Log "  Filtrados por fecha: $($antes - $trabajo.Count) ítems quedan fuera" -Level INFO
}

$script:totalTrabajo = if ($trabajo.Count -gt 0) { $trabajo.Count } else { $items1 + $items2 }

if ($script:totalTrabajo -eq 0) {
    Write-Log "" -Level INFO
    Write-Log "No hay ítems que cumplan los criterios. Nada que borrar." -Level OK
    exit 0
}

Write-Log "" -Level INFO
Write-Log "Ítems a borrar: $($script:totalTrabajo)" -Level HEADER

# ============================================================
# PROGRESO
# ============================================================
function Update-Progreso {
    param([int]$Procesados, [string]$Operacion, [switch]$Forzar)

    $seg       = [math]::Max(0.1, $script:sw.Elapsed.TotalSeconds)
    $ritmo     = $Procesados / $seg
    $gbLib     = ConvertTo-Gb $script:bytesLiberados
    $gbPorMin  = [math]::Round(($script:bytesLiberados / 1GB) / ($seg / 60), 2)
    $restantes = [math]::Max(0, $script:totalTrabajo - $Procesados)
    $eta       = if ($ritmo -gt 0) { $restantes / $ritmo } else { 0 }
    $pct       = [math]::Min(100, [math]::Round(($Procesados / [math]::Max(1, $script:totalTrabajo)) * 100, 1))

    # La barra se redibuja como mucho 2 veces por segundo
    if ($Forzar -or ([datetime]::Now - $script:ultimaBarra).TotalMilliseconds -ge 500) {
        Write-Progress -Activity "Vaciando papelera — $($script:web.Title)" `
                       -Status "$Procesados de $($script:totalTrabajo) ítems | $gbLib GB | $([math]::Round($ritmo,1)) it/s | $gbPorMin GB/min" `
                       -CurrentOperation "$Operacion — restan $restantes, ETA $(Format-Duracion $eta)" `
                       -PercentComplete $pct
        $script:ultimaBarra = [datetime]::Now
    }

    if (([datetime]::Now - $script:ultimoReporte).TotalSeconds -ge 10 -or $Procesados -ge $script:totalTrabajo) {
        Write-Log ("  {0,6:N1}% — {1}/{2} ítems | {3} GB | {4:N1} it/s | ETA {5}" -f `
                    $pct, $Procesados, $script:totalTrabajo, $gbLib, $ritmo, (Format-Duracion $eta)) -Level INFO
        $script:ultimoReporte = [datetime]::Now
    }
}

$script:web = $web

# ============================================================
# VACIADO
# ============================================================

if ($modoRapido) {

    # ----- MODO RÁPIDO: Clear-PnPRecycleBinItem -RowLimit -----
    $lotesEstimados = [math]::Ceiling($script:totalTrabajo / $TamanoLote)
    Write-Log "Iniciando borrado por lotes — ~$lotesEstimados llamadas al servidor" -Level HEADER

    Update-Progreso -Procesados 0 -Operacion "Enviando primer lote" -Forzar

    $nroLote     = 0
    $maxLotes    = $lotesEstimados + 10   # margen por si aparecen ítems nuevos
    $seguir      = $true
    $fallos      = 0
    $sinProgreso = 0
    $idAnterior  = ""

    while ($seguir -and $nroLote -lt $maxLotes) {

        $nroLote++
        Update-Progreso -Procesados $script:borrados -Operacion "Enviando lote $nroLote (hasta $TamanoLote ítems)" -Forzar

        $swLote = [System.Diagnostics.Stopwatch]::StartNew()
        $ok     = $false

        try {
            if ($Etapa -eq "Segunda") {
                Clear-PnPRecycleBinItem -SecondStageOnly -RowLimit $TamanoLote -Force -Connection $conn -ErrorAction Stop
            } else {
                Clear-PnPRecycleBinItem -All -RowLimit $TamanoLote -Force -Connection $conn -ErrorAction Stop
            }
            $ok = $true
        }
        catch {
            $msg = $_.Exception.Message -replace "`r`n|`n", " "
            $fallos++

            if ($msg -match "429|503|throttl") {
                $espera = 10 * $fallos
                Write-Log "  [WARN] Throttling del servidor — esperando ${espera}s" -Level WARN
                Update-Progreso -Procesados $script:borrados -Operacion "Throttling, esperando ${espera}s" -Forzar
                Start-Sleep -Seconds $espera
                $nroLote--   # el lote no se consumió
            }
            else {
                Write-Log "  [ERROR] Lote $nroLote falló: $msg" -Level ERROR
                Add-FilaLog -Etapa $Etapa -Titulo "Lote $nroLote" -Tipo "-" -MB "-" -Resultado "ERROR" -Detalle $msg
            }

            if ($fallos -ge 3) {
                Write-Log "  Tres fallos seguidos. Se corta el modo rápido." -Level ERROR
                break
            }
        }

        $swLote.Stop()

        if ($ok) {
            $fallos = 0

            # El cmdlet no informa cuántos borró: se estima por el tamaño del lote
            $desde = $script:borrados
            $hasta = [math]::Min($script:totalTrabajo - 1, $desde + $TamanoLote - 1)
            $cuantos = [math]::Max(0, $hasta - $desde + 1)

            $bytesLote = [double]0
            if ($trabajo.Count -gt 0 -and $desde -lt $trabajo.Count) {
                $finReal = [math]::Min($trabajo.Count - 1, $hasta)
                $bytesLote = [double](($trabajo[$desde..$finReal] | Measure-Object -Property Size -Sum).Sum)
            }

            $script:borrados       += $cuantos
            $script:bytesLiberados += $bytesLote

            Add-FilaLog -Etapa $Etapa -Titulo "Lote $nroLote" -Tipo "LOTE" `
                        -MB ([math]::Round($bytesLote/1MB, 2)) -Resultado "OK" `
                        -Detalle "~$cuantos ítems en $([math]::Round($swLote.Elapsed.TotalSeconds,1))s"

            if ($nroLote -eq 1) {
                Write-Log "  Primer lote OK en $([math]::Round($swLote.Elapsed.TotalSeconds,1))s" -Level OK
            }

            Update-Progreso -Procesados $script:borrados -Operacion "Lote $nroLote completo"
        }

        Flush-Log

        # ¿Queda algo? Consulta barata de 1 ítem.
        # Si el mismo ítem sigue apareciendo, el servidor no está borrando nada:
        # casi siempre es retención o bloqueo legal, y conviene cortar ahí.
        try {
            $resto = if ($Etapa -eq "Segunda") {
                @(Get-PnPRecycleBinItem -SecondStage -RowLimit 1 -Connection $conn -ErrorAction Stop)
            } else {
                @(Get-PnPRecycleBinItem -RowLimit 1 -Connection $conn -ErrorAction Stop)
            }

            $seguir = ($resto.Count -gt 0)

            if ($seguir) {
                $idActual = [string]$resto[0].Id

                if ($idActual -eq $idAnterior) {
                    $sinProgreso++
                    if ($sinProgreso -ge 2) {
                        Write-Log "" -Level INFO
                        Write-Log "  [ERROR] Dos lotes seguidos sin borrar nada — se corta." -Level ERROR
                        Write-Log "  El ítem '$($resto[0].Title)' sigue en la papelera después de pedir su borrado." -Level WARN
                        Write-Log "  Causa habitual: política de retención o bloqueo legal (eDiscovery)." -Level WARN
                        Write-Log "  Esos ítems no se pueden eliminar hasta que se libere la política." -Level WARN
                        $seguir = $false
                    }
                }
                else {
                    $sinProgreso = 0
                    $idAnterior  = $idActual
                }
            }
        }
        catch {
            Write-Log "  [WARN] No se pudo verificar si queda contenido: $($_.Exception.Message -replace "`r`n|`n", ' ')" -Level WARN
            $seguir = $false
        }
    }

    if ($nroLote -ge $maxLotes) {
        Write-Log "  [WARN] Se alcanzó el máximo de lotes. Volver a correr el script para continuar." -Level WARN
    }

    # El conteo por lote es una estimación: se ajusta al final con lo que quedó
    if (-not $seguir -and $script:borrados -gt $script:totalTrabajo) {
        $script:borrados = $script:totalTrabajo
    }
}
else {

    # ----- MODO INDIVIDUAL -----
    Write-Log "Iniciando borrado individual — $($script:totalTrabajo) llamadas" -Level HEADER
    Update-Progreso -Procesados 0 -Operacion "Preparando" -Forzar

    $procesados = 0

    foreach ($item in $trabajo) {

        $et = if ($item.ItemState -like "*Second*") { "Segunda" } else { "Primera" }
        $mb = [math]::Round(([double]$item.Size)/1MB, 2)

        if ($Simular) {
            Add-FilaLog -Etapa $et -Titulo $item.Title -Tipo $item.ItemType -MB $mb -Resultado "SIMULADO" -Detalle "Se borraría"
            $script:saltados++
        }
        else {
            try {
                Clear-PnPRecycleBinItem -Identity $item.Id -Force -Connection $conn -ErrorAction Stop
                $script:borrados++
                $script:bytesLiberados += [double]$item.Size
                Add-FilaLog -Etapa $et -Titulo $item.Title -Tipo $item.ItemType -MB $mb -Resultado "OK" -Detalle ""
            }
            catch {
                $script:errores++
                Add-FilaLog -Etapa $et -Titulo $item.Title -Tipo $item.ItemType -MB $mb `
                            -Resultado "ERROR" -Detalle ($_.Exception.Message -replace "`r`n|`n", " ")
            }
        }

        $procesados++
        Update-Progreso -Procesados $procesados -Operacion "Ítem $procesados de $($script:totalTrabajo)"

        if ($procesados % 100 -eq 0) { Flush-Log }
    }
}

Flush-Log
Write-Progress -Activity "Vaciando papelera" -Completed
$script:sw.Stop()

# ============================================================
# VERIFICACIÓN FINAL
# ============================================================
$quedanItems = -1
try {
    $verif = if ($Etapa -eq "Segunda") {
        @(Get-PnPRecycleBinItem -SecondStage -RowLimit 5000 -Connection $conn -ErrorAction Stop)
    } else {
        @(Get-PnPRecycleBinItem -RowLimit 5000 -Connection $conn -ErrorAction Stop)
    }
    $quedanItems = $verif.Count
}
catch { }

# ============================================================
# RESUMEN FINAL
# ============================================================
$seg     = [math]::Max(0.1, $script:sw.Elapsed.TotalSeconds)
$gbTotal = ConvertTo-Gb $script:bytesLiberados

Write-Log "" -Level INFO
Write-Log "======================================================" -Level HEADER
Write-Log "  RESUMEN FINAL"                                         -Level HEADER
Write-Log "  Sitio            : $($web.Title)"                      -Level INFO
Write-Log "  Etapas           : $Etapa"                              -Level INFO
Write-Log "  Método           : $(if ($modoRapido) { "lotes de $TamanoLote" } else { 'individual' })" -Level INFO
Write-Log "  Ítems borrados   : $($script:borrados) de $($script:totalTrabajo)$(if ($modoRapido) { ' (estimado)' })" -Level OK
Write-Log "  Espacio liberado : $gbTotal GB$(if ($modoRapido) { ' (estimado)' })" -Level OK
Write-Log "  Saltados         : $($script:saltados)"                 -Level INFO
Write-Log "  Errores          : $($script:errores)"                  -Level $(if ($script:errores -gt 0) {"ERROR"} else {"OK"})
Write-Log "  Duración         : $(Format-Duracion $seg)"             -Level INFO
Write-Log "  Ritmo            : $([math]::Round($script:borrados / $seg, 1)) ítems/s" -Level INFO
if ($quedanItems -ge 0) {
    Write-Log "  Quedan en papelera: $quedanItems$(if ($quedanItems -ge 5000) { '+' })" -Level $(if ($quedanItems -gt 0) {"WARN"} else {"OK"})
}
Write-Log "  Log              : $csvLog"                             -Level INFO
if (-not $SinInventario -and (Test-Path $csvInventario)) {
    Write-Log "  Inventario previo: $csvInventario"                  -Level INFO
}
Write-Log "======================================================" -Level HEADER

if ($quedanItems -gt 0) {
    Write-Log "Todavía queda contenido en la papelera." -Level WARN
    Write-Log "  Si son ítems bajo retención o bloqueo legal, no se pueden borrar." -Level WARN
    Write-Log "  Si no, volver a correr el script para la tanda siguiente." -Level INFO
}

if ($script:errores -gt 0) {
    Write-Log "Ítems con error — revisar la columna Detalle del log." -Level WARN
}

if ($script:borrados -gt 0) {
    Write-Log "La cuota del sitio se actualiza diferido: puede tardar horas en reflejarse." -Level WARN
    Write-Log "Para verificar el efecto: Script_A_DiagnosticoStorage.ps1 -SoloSitios `"$Url`"" -Level INFO
}

Write-Log "Vaciado de papelera finalizado." -Level OK