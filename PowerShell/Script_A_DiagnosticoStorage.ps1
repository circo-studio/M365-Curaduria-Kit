#Requires -Modules PnP.PowerShell
<#
===============================================================================
  Script_A_DiagnosticoStorage.ps1  |  Circo Studio  |  Kit de Curaduría M365
===============================================================================

PROPÓSITO:
  Explica el desvío entre la cuota que reporta el inventario (StorageUsageCurrent,
  todo lo que consume storage) y el contenido efectivamente relevado por
  Script_B_Masivo (File_x0020_Size, solo la versión vigente de cada archivo).

  Corre SOLO sobre los sitios con desvío grande, no sobre todo el tenant.
  Es de solo lectura: no modifica nada en el tenant.

  Desglosa la cuota en:
    - Contenido actual   (versión vigente de bibliotecas visibles)
    - Historial versiones (TotalSize - TotalFileStreamSize por biblioteca)
    - Papelera 1ª etapa   (borrado por usuarios, 93 días)
    - Papelera 2ª etapa   (purgado de la 1ª etapa, cuota adicional)
    - Preservation Hold   (copias que generan las políticas de retención)
    - Bibliotecas ocultas (sistema, apps, assets)
    - No explicado        (residuo: adjuntos de listas, OneNote, apps o lag
                           de la métrica de cuota, que se actualiza diferido)

USO:
  # Diagnóstico con umbrales default (desvío >= 5 GB y cuota >= 2x lo relevado)
  .\Script_A_DiagnosticoStorage.ps1 -ConfigPath "..\..\Config\cliente_config.json"

  # Más sensible, y sin recorrer la papelera (mucho más rápido)
  .\Script_A_DiagnosticoStorage.ps1 -ConfigPath ".\cliente_config.json" `
      -UmbralGB 1 -UmbralFactor 1.2 -SinPapelera

  # Solo un sitio puntual, ignorando umbrales
  .\Script_A_DiagnosticoStorage.ps1 -ConfigPath ".\cliente_config.json" `
      -SoloSitios "https://tenant.sharepoint.com/sites/EJEMPLO"

ENTRADAS:
  [Cliente]_A_Inventario_Sitios_[timestamp].csv   — cuota por sitio (Script A)
  [Cliente]_B_Docs_[Sitio]_[timestamp].csv        — contenido relevado (Script B)

ARCHIVOS GENERADOS:
  [Cliente]_A_DiagnosticoStorage_[timestamp]_RESUMEN.csv  — una fila por sitio
  [Cliente]_A_DiagnosticoStorage_[timestamp]_DETALLE.csv  — una fila por biblioteca

===============================================================================
#>

param(
    [Parameter(Mandatory, HelpMessage="Path al archivo cliente_config.json")]
    [string]$ConfigPath,

    # Desvío absoluto mínimo (GB) para diagnosticar un sitio
    [double]$UmbralGB = 5.0,

    # Relación mínima cuota / relevado para diagnosticar un sitio
    # 1.0 = diagnosticar todo lo que supere -UmbralGB, sin importar la proporción
    [double]$UmbralFactor = 2.0,

    # URLs a diagnosticar sí o sí, ignorando umbrales
    [string[]]$SoloSitios = @(),

    # Incluye sitios con cuota pero sin CSV de relevamiento (relevado = 0)
    [switch]$IncluirSinRelevar,

    # Saltea el recorrido de la papelera (es la parte más lenta)
    [switch]$SinPapelera,

    # Tope de ítems a leer por etapa de papelera (protege contra papeleras enormes)
    [int]$MaxItemsPapelera = 20000,

    # Limita la cantidad de sitios a diagnosticar (0 = todos los que superen el umbral)
    [int]$Top = 0,

    [switch]$ForceAuth
)

$moduloPath = Join-Path $PSScriptRoot "M365Curaduria.psm1"
if (-not (Test-Path $moduloPath)) {
    Write-Host "[ERROR] No se encontró M365Curaduria.psm1 en: $moduloPath" -ForegroundColor Red
    exit 1
}
Import-Module $moduloPath -Force

# ============================================================
# CONFIGURACIÓN Y SETUP
# ============================================================
$config = $null
try { $config = Read-ClientConfig -ConfigPath $ConfigPath }
catch { Write-Host "[ERROR] $_" -ForegroundColor Red; exit 1 }

if (-not (Test-Path $config.CarpetaLocal)) {
    New-Item -ItemType Directory -Path $config.CarpetaLocal -Force | Out-Null
}

$timestamp  = Get-Date -Format "yyyyMMdd_HHmmss"
$prefijo    = ($config.NombreCliente -replace "[^a-zA-Z0-9]", "_")
$csvResumen = "$($config.CarpetaLocal)\${prefijo}_A_DiagnosticoStorage_${timestamp}_RESUMEN.csv"
$csvDetalle = "$($config.CarpetaLocal)\${prefijo}_A_DiagnosticoStorage_${timestamp}_DETALLE.csv"

Write-Log "======================================================" -Level HEADER
Write-Log "  Kit M365 Curaduría — Script A DiagnosticoStorage"      -Level HEADER
Write-Log "  Cliente       : $($config.NombreCliente)"             -Level HEADER
Write-Log "  Umbral desvío : >= $UmbralGB GB  y  factor >= ${UmbralFactor}x" -Level HEADER
Write-Log "  Papelera      : $(if ($SinPapelera) { 'NO se recorre' } else { "hasta $MaxItemsPapelera ítems por etapa" })" -Level HEADER
Write-Log "  Resumen       : $csvResumen"                          -Level HEADER
Write-Log "======================================================" -Level HEADER

# Cast invariante (no depende del locale, a diferencia de TryParse)
function ConvertTo-Numero {
    param([string]$Valor)
    if (-not $Valor) { return 0.0 }
    try   { return [double]($Valor -replace ',', '.') }
    catch { return 0.0 }
}

function ConvertTo-Gb {
    param([double]$Bytes)
    return [math]::Round($Bytes / 1GB, 2)
}

# Títulos que identifican la Preservation Hold Library (varía por idioma del tenant)
$patronesPHL = @(
    "Preservation Hold Library",
    "Preservation*",
    "*conservaci*n*",
    "*retenci*n*"
)

function Test-EsPHL {
    param([string]$Titulo)
    foreach ($p in $patronesPHL) { if ($Titulo -like $p) { return $true } }
    return $false
}

# ============================================================
# CUOTA POR SITIO (inventario del Script A)
# ============================================================
$csvInventario = Get-ChildItem $config.CarpetaLocal -Filter "${prefijo}_A_Inventario_Sitios_*.csv" -ErrorAction SilentlyContinue |
                 Sort-Object LastWriteTime -Descending | Select-Object -First 1

if (-not $csvInventario) {
    Write-Log "No se encontró el CSV de inventario. Correr primero Script_A_Inventario.ps1" -Level ERROR
    exit 1
}

Write-Log "" -Level INFO
Write-Log "CSV de inventario   : $($csvInventario.Name)" -Level OK

$inventario = @(Import-Csv -Path $csvInventario.FullName -Encoding UTF8 | Where-Object { $_.URL })
Write-Log "  Sitios en el inventario: $($inventario.Count)" -Level INFO

# ============================================================
# CONTENIDO RELEVADO POR SITIO (CSVs del Script B)
# ============================================================
$csvsDocs = @(Get-ChildItem $config.CarpetaLocal -Filter "${prefijo}_B_Docs_*.csv" -ErrorAction SilentlyContinue)
if ($csvsDocs.Count -eq 0) {
    $csvsDocs = @(Get-ChildItem $config.CarpetaLocal -Filter "*_B_Docs_*.csv" -ErrorAction SilentlyContinue)
}

if ($csvsDocs.Count -eq 0) {
    Write-Log "No se encontraron CSVs de relevamiento (*_B_Docs_*.csv)." -Level WARN
    Write-Log "Sin relevamiento no hay desvío que medir. Correr Script_B_Masivo, o usar -SoloSitios." -Level WARN
    if ($SoloSitios.Count -eq 0 -and -not $IncluirSinRelevar) { exit 1 }
}
else {
    Write-Log "CSVs de relevamiento: $($csvsDocs.Count)" -Level OK
}

# Acumular bytes y archivos por URL de sitio (normalizada)
$relevado = @{}

foreach ($csv in $csvsDocs) {
    try {
        $filas = @(Import-Csv -Path $csv.FullName -Encoding UTF8)
    }
    catch {
        Write-Log "  No se pudo leer $($csv.Name): $_" -Level WARN
        continue
    }

    foreach ($f in $filas) {
        if (-not $f.URL_Sitio) { continue }
        $k = ([string]$f.URL_Sitio).Trim().TrimEnd('/').ToLower()

        if (-not $relevado.ContainsKey($k)) {
            $relevado[$k] = [PSCustomObject]@{ Bytes = [double]0; Archivos = 0 }
        }
        $relevado[$k].Bytes    += (ConvertTo-Numero $f.'Tamaño_Bytes')
        $relevado[$k].Archivos += 1
    }
}

Write-Log "  Sitios con relevamiento: $($relevado.Keys.Count)" -Level INFO

# ============================================================
# SELECCIÓN DE SITIOS CON DESVÍO
# ============================================================
$forzadas = @($SoloSitios | ForEach-Object { ([string]$_).Trim().TrimEnd('/').ToLower() })
$candidatos = [System.Collections.Generic.List[object]]::new()

foreach ($fila in $inventario) {

    $url = ([string]$fila.URL).Trim().TrimEnd('/')
    $key = $url.ToLower()

    $cuotaGb = ConvertTo-Numero $fila.Storage_GB

    $relGb   = 0.0
    $relArch = 0
    $tieneRelevamiento = $relevado.ContainsKey($key)
    if ($tieneRelevamiento) {
        $relGb   = [math]::Round($relevado[$key].Bytes / 1GB, 2)
        $relArch = $relevado[$key].Archivos
    }

    $desvio = [math]::Round($cuotaGb - $relGb, 2)
    $factor = if ($relGb -gt 0) { [math]::Round($cuotaGb / $relGb, 2) } else { 0 }

    $forzado = $forzadas -contains $key

    if (-not $forzado) {
        if (-not $tieneRelevamiento -and -not $IncluirSinRelevar) { continue }
        if ($cuotaGb -le 0) { continue }
        if ($desvio -lt $UmbralGB) { continue }
        # factor 0 = sin relevamiento; se evalúa solo por desvío absoluto
        if ($factor -gt 0 -and $factor -lt $UmbralFactor) { continue }
    }

    $candidatos.Add([PSCustomObject]@{
        Nombre     = $fila.Nombre
        URL        = $url
        Tipo       = if ($fila.Tipo_Sitio -eq "Canal de Teams") { "CANAL" } else { "TeamSite" }
        Cuota_GB   = $cuotaGb
        Relevado_GB = $relGb
        Archivos_Relevados = $relArch
        Desvio_GB  = $desvio
        Factor     = $factor
        Forzado    = $forzado
    })
}

if ($candidatos.Count -eq 0) {
    Write-Log "" -Level INFO
    Write-Log "Ningún sitio supera el umbral de desvío. Nada que diagnosticar." -Level OK
    Write-Log "Para ser más sensible: -UmbralGB 1 -UmbralFactor 1.2" -Level INFO
    exit 0
}

$aDiagnosticar = @($candidatos | Sort-Object Desvio_GB -Descending)
if ($Top -gt 0 -and $aDiagnosticar.Count -gt $Top) {
    $aDiagnosticar = @($aDiagnosticar | Select-Object -First $Top)
    Write-Log "  Limitado a los $Top sitios de mayor desvío (-Top)" -Level WARN
}

Write-Log "" -Level INFO
Write-Log "Sitios con desvío   : $($aDiagnosticar.Count)" -Level OK
Write-Log "Desvío total        : $([math]::Round(($aDiagnosticar | Measure-Object -Property Desvio_GB -Sum).Sum, 1)) GB" -Level INFO

# ============================================================
# LOGIN ÚNICO
# ============================================================
Write-Log "" -Level INFO
try {
    Connect-TenantOnce -Config $config -ForceAuth:$ForceAuth
}
catch {
    Write-Log "No se pudo iniciar sesión: $_" -Level ERROR
    exit 1
}

# ============================================================
# DIAGNÓSTICO POR SITIO
# ============================================================
$resumen = [System.Collections.Generic.List[PSObject]]::new()
$detalle = [System.Collections.Generic.List[PSObject]]::new()

$total  = $aDiagnosticar.Count
$orden  = 0
$conError = 0

foreach ($sitio in $aDiagnosticar) {
    $orden++

    Write-Log "" -Level INFO
    Write-Log "[$orden/$total] $($sitio.Nombre)" -Level HEADER
    Write-Log "  Cuota: $($sitio.Cuota_GB) GB  |  Relevado: $($sitio.Relevado_GB) GB  |  Desvío: $($sitio.Desvio_GB) GB ($($sitio.Factor)x)" -Level INFO
    Write-Log "  URL  : $($sitio.URL)" -Level INFO

    # Valores por defecto de la fila de resumen
    $filaResumen = [PSCustomObject]@{
        Nombre               = $sitio.Nombre
        URL                  = $sitio.URL
        Tipo                 = $sitio.Tipo
        Cuota_GB             = $sitio.Cuota_GB
        Relevado_GB          = $sitio.Relevado_GB
        Archivos_Relevados   = $sitio.Archivos_Relevados
        Desvio_GB            = $sitio.Desvio_GB
        Factor               = $sitio.Factor
        Actual_Visibles_GB   = 0
        Versiones_GB         = 0
        Papelera_1ra_GB      = 0
        Papelera_2da_GB      = 0
        PreservationHold_GB  = 0
        Ocultas_GB           = 0
        No_Explicado_GB      = 0
        Bibliotecas_Visibles = 0
        Bibliotecas_Ocultas  = 0
        Items_Papelera       = 0
        Papelera_Truncada    = "No"
        Causa_Principal      = ""
        Accion_Sugerida      = ""
        Estado               = "OK"
        Detalle              = ""
    }

    $conn = $null
    try {
        $conn = Connect-PnPOnline -Url $sitio.URL -ClientId $config.ClientId `
                                  -PersistLogin -ReturnConnection -ErrorAction Stop
    }
    catch {
        $msg = $_.Exception.Message -replace "`r`n|`n", " "
        Write-Log "  [ERROR] Sin acceso al sitio: $msg" -Level ERROR
        $filaResumen.Estado  = "SIN_ACCESO"
        $filaResumen.Detalle = $msg
        $resumen.Add($filaResumen)
        $conError++
        continue
    }

    # --- Bibliotecas y métricas de storage ---
    $bytesActualVisibles = [double]0
    $bytesVersiones      = [double]0
    $bytesPHL            = [double]0
    $bytesOcultas        = [double]0
    $cantVisibles        = 0
    $cantOcultas         = 0

    try {
        $listas = @(Get-PnPList -Connection $conn -ErrorAction Stop |
                    Where-Object { $_.BaseType -eq "DocumentLibrary" })

        Write-Log "  Bibliotecas (incluye ocultas): $($listas.Count)" -Level INFO

        foreach ($lista in $listas) {

            # Cada biblioteca se aísla: un error puntual no debe abortar el sitio
            try {
                $esPHL    = Test-EsPHL -Titulo $lista.Title
                $esOculta = [bool]$lista.Hidden

                $totalBytes  = $null
                $actualBytes = $null

                # Get-PnPFolderStorageMetric: TotalSize incluye versiones y metadata,
                # TotalFileStreamSize es solo la versión vigente de cada archivo
                try {
                    $metrica = Get-PnPFolderStorageMetric -List $lista.Title -Connection $conn -ErrorAction Stop
                    $totalBytes  = [double]$metrica.TotalSize
                    $actualBytes = [double]$metrica.TotalFileStreamSize
                }
                catch {
                    try {
                        $rootUrl = (Get-PnPList -Identity $lista.Title -Includes RootFolder -Connection $conn -ErrorAction Stop).RootFolder.ServerRelativeUrl
                        $metrica = Get-PnPFolderStorageMetric -Folder $rootUrl -Connection $conn -ErrorAction Stop
                        $totalBytes  = [double]$metrica.TotalSize
                        $actualBytes = [double]$metrica.TotalFileStreamSize
                    }
                    catch {
                        Write-Log "    [WARN] Sin métricas para '$($lista.Title)'" -Level WARN
                    }
                }

                # Tipado explícito a double: los tamaños desbordan Int32 (> 2 GB)
                [double]$tot = if ($null -ne $totalBytes)  { $totalBytes }  else { 0 }
                [double]$act = if ($null -ne $actualBytes) { $actualBytes } else { 0 }
                [double]$ver = $tot - $act
                if ($ver -lt 0) { $ver = [double]0 }

                if ($esPHL) {
                    $bytesPHL += $tot
                }
                elseif ($esOculta) {
                    $bytesOcultas += $tot
                    $cantOcultas++
                }
                else {
                    $bytesActualVisibles += $act
                    $bytesVersiones      += $ver
                    $cantVisibles++
                }

                $detalle.Add([PSCustomObject]@{
                    Sitio            = $sitio.Nombre
                    URL_Sitio        = $sitio.URL
                    Biblioteca       = $lista.Title
                    Oculta           = if ($esOculta) { "Sí" } else { "No" }
                    Es_PreservHold   = if ($esPHL)    { "Sí" } else { "No" }
                    ItemCount        = $lista.ItemCount
                    Total_GB         = ConvertTo-Gb $tot
                    Actual_GB        = ConvertTo-Gb $act
                    Versiones_GB     = ConvertTo-Gb $ver
                    Ratio_Versiones  = if ($act -gt 0) { [math]::Round($tot / $act, 2) } else { 0 }
                })
            }
            catch {
                $msgBib = $_.Exception.Message -replace "`r`n|`n", " "
                Write-Log "    [WARN] Error en biblioteca '$($lista.Title)': $msgBib" -Level WARN
            }
        }
    }
    catch {
        $msg = $_.Exception.Message -replace "`r`n|`n", " "
        Write-Log "  [ERROR] Al leer bibliotecas: $msg" -Level ERROR
        $filaResumen.Estado  = "ERROR_BIBLIOTECAS"
        $filaResumen.Detalle = $msg
        $conError++
    }

    # --- Papelera de reciclaje (1ª y 2ª etapa) ---
    $bytesPap1 = [double]0
    $bytesPap2 = [double]0
    $itemsPap  = 0
    $truncada  = $false

    if (-not $SinPapelera) {
        foreach ($etapa in @("Primera","Segunda")) {
            try {
                $items = if ($etapa -eq "Primera") {
                    @(Get-PnPRecycleBinItem -FirstStage -RowLimit $MaxItemsPapelera -Connection $conn -ErrorAction Stop)
                } else {
                    @(Get-PnPRecycleBinItem -SecondStage -RowLimit $MaxItemsPapelera -Connection $conn -ErrorAction Stop)
                }

                $suma = ($items | Measure-Object -Property Size -Sum).Sum
                if (-not $suma) { $suma = 0 }

                if ($etapa -eq "Primera") { $bytesPap1 = [double]$suma } else { $bytesPap2 = [double]$suma }
                $itemsPap += $items.Count

                if ($items.Count -ge $MaxItemsPapelera) { $truncada = $true }
            }
            catch {
                Write-Log "    [WARN] No se pudo leer la papelera ($etapa etapa): $($_.Exception.Message -replace "`r`n|`n", ' ')" -Level WARN
            }
        }

        if ($truncada) {
            Write-Log "    [WARN] Papelera truncada en $MaxItemsPapelera ítems — el total está subestimado" -Level WARN
        }
    }

    # --- Consolidar ---
    $filaResumen.Actual_Visibles_GB   = ConvertTo-Gb $bytesActualVisibles
    $filaResumen.Versiones_GB         = ConvertTo-Gb $bytesVersiones
    $filaResumen.Papelera_1ra_GB      = ConvertTo-Gb $bytesPap1
    $filaResumen.Papelera_2da_GB      = ConvertTo-Gb $bytesPap2
    $filaResumen.PreservationHold_GB  = ConvertTo-Gb $bytesPHL
    $filaResumen.Ocultas_GB           = ConvertTo-Gb $bytesOcultas
    $filaResumen.Bibliotecas_Visibles = $cantVisibles
    $filaResumen.Bibliotecas_Ocultas  = $cantOcultas
    $filaResumen.Items_Papelera       = $itemsPap
    $filaResumen.Papelera_Truncada    = if ($truncada) { "Sí" } else { "No" }

    $explicado = $filaResumen.Actual_Visibles_GB + $filaResumen.Versiones_GB +
                 $filaResumen.Papelera_1ra_GB + $filaResumen.Papelera_2da_GB +
                 $filaResumen.PreservationHold_GB + $filaResumen.Ocultas_GB

    $filaResumen.No_Explicado_GB = [math]::Round($sitio.Cuota_GB - $explicado, 2)

    # --- Causa principal y acción sugerida ---
    $componentes = @(
        @{ Nombre = "Historial de versiones"; Gb = $filaResumen.Versiones_GB;
           Accion = "Limitar versiones en la biblioteca y hacer trimming antes de migrar" }
        @{ Nombre = "Papelera de reciclaje";  Gb = ($filaResumen.Papelera_1ra_GB + $filaResumen.Papelera_2da_GB);
           Accion = "Vaciar ambas etapas de la papelera (revisar antes con el cliente)" }
        @{ Nombre = "Preservation Hold";      Gb = $filaResumen.PreservationHold_GB;
           Accion = "Revisar políticas de retención con el cliente — no se puede borrar directo" }
        @{ Nombre = "Bibliotecas ocultas";    Gb = $filaResumen.Ocultas_GB;
           Accion = "Revisar el DETALLE: assets, apps o bibliotecas de sistema" }
        @{ Nombre = "Contenido actual";       Gb = $filaResumen.Actual_Visibles_GB;
           Accion = "Storage real de contenido vigente — comparar con lo relevado" }
        @{ Nombre = "No explicado";           Gb = $filaResumen.No_Explicado_GB;
           Accion = "Adjuntos de listas, OneNote o lag de la métrica de cuota (se actualiza diferido)" }
    )

    $mayor = $componentes | Sort-Object { $_.Gb } -Descending | Select-Object -First 1
    $filaResumen.Causa_Principal = "$($mayor.Nombre) ($($mayor.Gb) GB)"
    $filaResumen.Accion_Sugerida = $mayor.Accion

    # Aviso si el relevamiento no cubrió todo el contenido vigente
    $gapRelevamiento = [math]::Round($filaResumen.Actual_Visibles_GB - $sitio.Relevado_GB, 2)
    if ($gapRelevamiento -gt 1) {
        $filaResumen.Detalle = "Relevamiento por debajo del contenido vigente en $gapRelevamiento GB — revisar bibliotecas excluidas por Script B"
    }

    $resumen.Add($filaResumen)

    # --- Salida en consola ---
    Write-Log "  ── Desglose de la cuota ──" -Level INFO
    Write-Log "    Contenido actual    : $($filaResumen.Actual_Visibles_GB) GB" -Level INFO
    Write-Log "    Historial versiones : $($filaResumen.Versiones_GB) GB" -Level $(if ($filaResumen.Versiones_GB -gt 0) {"WARN"} else {"INFO"})
    Write-Log "    Papelera 1ª etapa   : $($filaResumen.Papelera_1ra_GB) GB" -Level $(if ($filaResumen.Papelera_1ra_GB -gt 0) {"WARN"} else {"INFO"})
    Write-Log "    Papelera 2ª etapa   : $($filaResumen.Papelera_2da_GB) GB" -Level $(if ($filaResumen.Papelera_2da_GB -gt 0) {"WARN"} else {"INFO"})
    Write-Log "    Preservation Hold   : $($filaResumen.PreservationHold_GB) GB" -Level $(if ($filaResumen.PreservationHold_GB -gt 0) {"WARN"} else {"INFO"})
    Write-Log "    Bibliotecas ocultas : $($filaResumen.Ocultas_GB) GB" -Level INFO
    Write-Log "    No explicado        : $($filaResumen.No_Explicado_GB) GB" -Level INFO
    Write-Log "    → Causa principal   : $($filaResumen.Causa_Principal)" -Level OK
    if ($filaResumen.Detalle) { Write-Log "    [WARN] $($filaResumen.Detalle)" -Level WARN }

    Start-Sleep -Milliseconds 300
}

# ============================================================
# EXPORTAR
# ============================================================
Write-Log "" -Level INFO

if ($resumen.Count -gt 0) {
    $resumen | Export-Csv -Path $csvResumen -NoTypeInformation -Encoding UTF8
    Write-Log "Resumen exportado : $csvResumen" -Level OK
}

if ($detalle.Count -gt 0) {
    $detalle | Sort-Object Sitio, @{Expression={$_.Total_GB}; Descending=$true} |
        Export-Csv -Path $csvDetalle -NoTypeInformation -Encoding UTF8
    Write-Log "Detalle exportado : $csvDetalle" -Level OK
}

# ============================================================
# RESUMEN FINAL
# ============================================================
$sumaCuota      = ($resumen | Measure-Object -Property Cuota_GB -Sum).Sum
$sumaRelevado   = ($resumen | Measure-Object -Property Relevado_GB -Sum).Sum
$sumaVersiones  = ($resumen | Measure-Object -Property Versiones_GB -Sum).Sum
$sumaPapelera   = (($resumen | Measure-Object -Property Papelera_1ra_GB -Sum).Sum +
                   ($resumen | Measure-Object -Property Papelera_2da_GB -Sum).Sum)
$sumaPHL        = ($resumen | Measure-Object -Property PreservationHold_GB -Sum).Sum
$sumaOcultas    = ($resumen | Measure-Object -Property Ocultas_GB -Sum).Sum
$sumaNoExplic   = ($resumen | Measure-Object -Property No_Explicado_GB -Sum).Sum

Write-Log "" -Level INFO
Write-Log "======================================================" -Level HEADER
Write-Log "  RESUMEN FINAL"                                         -Level HEADER
Write-Log "  Sitios diagnosticados : $($resumen.Count)"            -Level INFO
Write-Log "  Con error / sin acceso: $conError"                    -Level $(if ($conError -gt 0) {"WARN"} else {"OK"})
Write-Log "  ── Totales sobre los sitios diagnosticados ──"        -Level INFO
Write-Log "  Cuota                 : $([math]::Round($sumaCuota, 1)) GB"     -Level INFO
Write-Log "  Relevado (Script B)   : $([math]::Round($sumaRelevado, 1)) GB"  -Level INFO
Write-Log "  Historial versiones   : $([math]::Round($sumaVersiones, 1)) GB" -Level WARN
Write-Log "  Papelera (1ª + 2ª)    : $([math]::Round($sumaPapelera, 1)) GB"  -Level WARN
Write-Log "  Preservation Hold     : $([math]::Round($sumaPHL, 1)) GB"       -Level WARN
Write-Log "  Bibliotecas ocultas   : $([math]::Round($sumaOcultas, 1)) GB"   -Level INFO
Write-Log "  No explicado          : $([math]::Round($sumaNoExplic, 1)) GB"  -Level INFO
Write-Log "======================================================" -Level HEADER

if ($SinPapelera) {
    Write-Log "La papelera no se recorrió (-SinPapelera): su volumen cae en 'No explicado'." -Level WARN
}

$topCausas = $resumen | Where-Object { $_.Estado -eq "OK" } |
             Group-Object Causa_Principal | Sort-Object Count -Descending

if ($topCausas) {
    Write-Log "Causas principales por cantidad de sitios:" -Level INFO
    foreach ($c in $topCausas) {
        Write-Log "  $($c.Count) sitio(s) — $($c.Name)" -Level INFO
    }
}

Write-Log "Script A DiagnosticoStorage finalizado." -Level OK