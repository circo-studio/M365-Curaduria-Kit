#Requires -Modules PnP.PowerShell
<#
===============================================================================
  Script_C_Refresh.ps1  |  Circo Studio  |  Kit de Curaduría M365
===============================================================================

PROPÓSITO:
  Re-relevamiento incremental. Compara la fecha del último CSV de cada sitio
  contra la "Ultima_Actividad" del inventario. Solo reprocesa los sitios que
  tuvieron actividad después del CSV previo.

  Ideal para actualización mensual del ISI sin volver a procesar todo el tenant.

USO:
  .\Script_C_Refresh.ps1 -ConfigPath "..\..\Config\cliente_config.json"

  # Forzar re-proceso de absolutamente todos los sitios:
  .\Script_C_Refresh.ps1 -ConfigPath "..." -DiasUmbral 999

PARÁMETROS:
  -ConfigPath   (obligatorio) Path al cliente_config.json
  -DiasUmbral   (opcional)    Override del valor en el config. 999 = forzar todos.
  -ForceAuth    (switch)      Re-autenticación forzada

LÓGICA DE DECISIÓN POR SITIO:
  ¿Existe CSV previo?
    NO  → procesar (nunca fue relevado)
    SI  → ¿Actividad > fecha del CSV? → re-procesar
          ¿Actividad reciente (< diasUmbral días)? → re-procesar
          Resto → saltear (sin cambios)

===============================================================================
#>

param(
    [Parameter(Mandatory, HelpMessage="Path al archivo cliente_config.json")]
    [string]$ConfigPath,

    # Override del diasUmbral del config (usar 999 para forzar todos)
    [int]$DiasUmbral = -1,

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

# DiasUmbral: parámetro > config > default 7
$umbral = if ($DiasUmbral -ge 0) { $DiasUmbral } else { $config.DiasUmbral }

if (-not (Test-Path $config.CarpetaLocal)) {
    New-Item -ItemType Directory -Path $config.CarpetaLocal -Force | Out-Null
}

$timestamp   = Get-Date -Format "yyyyMMdd_HHmmss"
$fechaHoy    = Get-Date
$fechaUmbral = $fechaHoy.AddDays(-$umbral)
$prefijo     = ($config.NombreCliente -replace "[^a-zA-Z0-9]", "_")
$logFile     = "$($config.CarpetaLocal)\${prefijo}_C_Refresh_${timestamp}_LOG.txt"
New-Item -ItemType File -Path $logFile -Force | Out-Null

Write-Log "======================================================" -Level HEADER -LogFile $logFile
Write-Log "  Kit M365 Curaduría — Script C Refresh"                -Level HEADER -LogFile $logFile
Write-Log "  Cliente  : $($config.NombreCliente)"                  -Level HEADER -LogFile $logFile
Write-Log "  Umbral   : últimos $umbral días"                      -Level HEADER -LogFile $logFile
Write-Log "  Referencia: $($fechaHoy.ToString('yyyy-MM-dd'))"      -Level HEADER -LogFile $logFile
Write-Log "======================================================" -Level HEADER -LogFile $logFile

# ============================================================
# BUSCAR CSV DE INVENTARIO
# ============================================================
$csvInventario = Get-ChildItem $config.CarpetaLocal -Filter "${prefijo}_A_Inventario_Sitios_*.csv" |
                 Sort-Object LastWriteTime -Descending |
                 Select-Object -First 1

if (-not $csvInventario) {
    Write-Log "No se encontró el CSV de inventario. Correr primero Script_A_Inventario.ps1" `
              -Level ERROR -LogFile $logFile
    exit 1
}

Write-Log "CSV de inventario: $($csvInventario.Name)" -Level OK -LogFile $logFile

# ============================================================
# OBTENER TODOS LOS SITIOS LEGACY (TEAMS) CON CONTENIDO
# (el refresh incluye los que ya fueron analizados -- puede haber tenido actividad nueva)
# ============================================================
$inventario  = Import-Csv -Path $csvInventario.FullName -Encoding UTF8
$todosSitios = [System.Collections.Generic.List[object]]::new()

foreach ($fila in $inventario) {
    # Filtrar por label legacy -- el Script A Inventario lo asigna segun template o prefijo
    $labelLegacy = if ($config.LabelLegacy) { $config.LabelLegacy } else { "Legacy (Teams)" }
    if ($fila.Clasificacion -ne $labelLegacy) { continue }
    if ($fila.Bloqueado     -eq "ReadOnly")            { continue }

    # Excluir canales de padres ya cubiertos
    if ($fila.Tipo_Sitio -eq "Canal de Teams") {
        $canalExcluido = $false
        foreach ($prefijoCfg in $config.CanalesExcluidos) {
            if ($fila.Nombre -like "$prefijoCfg*") { $canalExcluido = $true; break }
        }
        if ($canalExcluido) { continue }
    }

    $gb = [double]($fila.Storage_GB -replace ',', '.')
    if ($gb -le 0) { continue }

    $fila | Add-Member -NotePropertyName "_StorageNum"  -NotePropertyValue $gb   -Force
    $todosSitios.Add($fila)
}

$todosSitios = $todosSitios | Sort-Object { $_._StorageNum } -Descending
Write-Log "Total sitios $labelLegacy con contenido: $($todosSitios.Count)" -Level OK -LogFile $logFile

# ============================================================
# PRE-ANÁLISIS: decidir qué procesar y qué saltear
# ============================================================

function Get-CsvPrevio {
    param([string]$NombreLimpio)
    $candidatos = @()
    $candidatos += Get-ChildItem $config.CarpetaLocal -Filter "${prefijo}_B_Docs_${NombreLimpio}_*.csv" -ErrorAction SilentlyContinue
    # Compatibilidad con CSVs de ejecuciones anteriores al kit
    $candidatos += Get-ChildItem $config.CarpetaLocal -Filter "*Relevamiento_${NombreLimpio}_*.csv" -ErrorAction SilentlyContinue
    if ($candidatos.Count -eq 0) { return $null }
    return $candidatos | Sort-Object LastWriteTime -Descending | Select-Object -First 1
}

$aProcesar = [System.Collections.Generic.List[object]]::new()
$aSaltear  = [System.Collections.Generic.List[object]]::new()

foreach ($sitio in $todosSitios) {
    $nombreLimpio = $sitio.Nombre -replace "[^a-zA-Z0-9_-]", "_" -replace "__+", "_"
    $nombreLimpio = $nombreLimpio.Substring(0, [Math]::Min($nombreLimpio.Length, 50))
    $sitio | Add-Member -NotePropertyName "_NombreLimpio" -NotePropertyValue $nombreLimpio -Force

    $fechaActividad = $null
    if ($sitio.Ultima_Actividad -and $sitio.Ultima_Actividad -ne "") {
        try { $fechaActividad = [datetime]::Parse($sitio.Ultima_Actividad) } catch {}
    }

    $csvPrevio = Get-CsvPrevio -NombreLimpio $nombreLimpio

    $sitio | Add-Member -NotePropertyName "_CsvPrevio"      -NotePropertyValue $csvPrevio      -Force
    $sitio | Add-Member -NotePropertyName "_FechaActividad" -NotePropertyValue $fechaActividad -Force

    $debeCorrer = $false
    $motivo     = ""

    if (-not $csvPrevio) {
        $debeCorrer = $true
        $motivo = "Sin CSV previo — nunca relevado"
    }
    elseif ($umbral -ge 999) {
        $debeCorrer = $true
        $motivo = "Forzado (DiasUmbral=999)"
    }
    elseif ($fechaActividad -and $fechaActividad -gt $csvPrevio.LastWriteTime) {
        $debeCorrer = $true
        $motivo = "Actividad $($fechaActividad.ToString('yyyy-MM-dd')) > CSV $($csvPrevio.LastWriteTime.ToString('yyyy-MM-dd HH:mm'))"
    }
    elseif ($fechaActividad -and $fechaActividad -ge $fechaUmbral) {
        $debeCorrer = $true
        $motivo = "Actividad reciente ($($fechaActividad.ToString('yyyy-MM-dd')))"
    }
    else {
        $debeCorrer = $false
        $motivo = if ($csvPrevio) {
            "CSV vigente ($($csvPrevio.LastWriteTime.ToString('yyyy-MM-dd HH:mm')))"
        } else {
            "Sin actividad reciente"
        }
    }

    $sitio | Add-Member -NotePropertyName "_DebeCorrer" -NotePropertyValue $debeCorrer -Force
    $sitio | Add-Member -NotePropertyName "_Motivo"     -NotePropertyValue $motivo     -Force

    if ($debeCorrer) { $aProcesar.Add($sitio) } else { $aSaltear.Add($sitio) }
}

# Mostrar plan antes de ejecutar
Write-Log "" -Level INFO
Write-Log "PLAN DE EJECUCIÓN:" -Level HEADER -LogFile $logFile
Write-Log "  A procesar : $($aProcesar.Count) sitios" -Level OK   -LogFile $logFile
Write-Log "  A saltear  : $($aSaltear.Count) sitios"  -Level WARN -LogFile $logFile
Write-Log "" -Level INFO

Write-Log "Sitios a PROCESAR:" -Level HEADER -LogFile $logFile
foreach ($s in ($aProcesar | Sort-Object { $_._StorageNum } -Descending)) {
    Write-Log ("  {0,8} GB  {1}" -f $s.Storage_GB, $s.Nombre) -Level INFO -LogFile $logFile
    Write-Log ("             → {0}" -f $s._Motivo) -Level INFO -LogFile $logFile
}

Write-Log "" -Level INFO
Write-Log "Sitios SALTEADOS:" -Level HEADER -LogFile $logFile
foreach ($s in $aSaltear) {
    Write-Log ("  {0,8} GB  {1}  [{2}]" -f $s.Storage_GB, $s.Nombre, $s._Motivo) -Level WARN -LogFile $logFile
}

if ($aProcesar.Count -eq 0) {
    Write-Log "" -Level INFO
    Write-Log "Todos los sitios están vigentes. Nada que procesar." -Level OK -LogFile $logFile
    exit 0
}

# ============================================================
# LOGIN ÚNICO
# ============================================================
try { Connect-TenantOnce -Config $config -LogFile $logFile -ForceAuth:$ForceAuth }
catch { Write-Log "No se pudo iniciar sesión." -Level ERROR -LogFile $logFile; exit 1 }

# ============================================================
# LOOP PRINCIPAL — solo los sitios que deben correr
# ============================================================
$resumen = [System.Collections.Generic.List[PSObject]]::new()
$total   = $aProcesar.Count
$orden   = 0

$aProcesar = $aProcesar | Sort-Object { $_._StorageNum } -Descending

foreach ($sitio in $aProcesar) {
    $orden++
    $tipo = if ($sitio.Tipo_Sitio -eq "Canal de Teams") { "CANAL" } else { "TeamSite" }

    Write-Log "" -Level INFO
    Write-Log "[$orden/$total] $($sitio.Nombre) — $($sitio.Storage_GB) GB [$tipo]" -Level HEADER -LogFile $logFile
    Write-Log "  Motivo: $($sitio._Motivo)" -Level INFO -LogFile $logFile

    $inicio    = Get-Date
    $resultado = Invoke-RelevamientoSitio `
        -SiteUrl       $sitio.URL             `
        -SiteName      $sitio.Nombre          `
        -CarpetaSalida $config.CarpetaLocal   `
        -Config        $config                `
        -LogFile       $logFile               `
        -Timestamp     $timestamp

    $duracion = [math]::Round(((Get-Date) - $inicio).TotalSeconds)

    $filaResumen = [PSCustomObject]@{
        "Orden"            = $orden
        "Nombre"           = $sitio.Nombre
        "URL"              = $sitio.URL
        "Tipo"             = $tipo
        "Storage_GB"       = $sitio.Storage_GB
        "Ultima_Actividad" = $sitio.Ultima_Actividad
        "Motivo_Refresh"   = $sitio._Motivo
        "Bibliotecas"      = $resultado.Bibliotecas
        "Archivos"         = $resultado.Archivos
        "Errores"          = $resultado.Errores
        "Duracion_Seg"     = $duracion
        "Estado"           = $resultado.Estado
        "CSV_Salida"       = $resultado.CsvPath
    }
    $resumen.Add($filaResumen)

    # Guardar resumen parcial
    $resumenPath = "$($config.CarpetaLocal)\${prefijo}_C_Refresh_Resumen_${timestamp}.csv"
    $resumen | Export-Csv -Path $resumenPath -NoTypeInformation -Encoding UTF8
}

# ============================================================
# RESUMEN FINAL
# ============================================================
$ok           = @($resumen | Where-Object { $_.Estado -eq "OK" })
$errores      = @($resumen | Where-Object { $_.Estado -like "Error*" })
$sinContenido = @($resumen | Where-Object { $_.Estado -in @("SinArchivos","SinBibliotecas") })

Write-Log "" -Level INFO
Write-Log "======================================================" -Level HEADER -LogFile $logFile
Write-Log "  RESUMEN FINAL"                                         -Level HEADER -LogFile $logFile
Write-Log "======================================================" -Level HEADER -LogFile $logFile
Write-Log "  Procesados OK   : $($ok.Count)"           -Level OK   -LogFile $logFile
Write-Log "  Sin contenido   : $($sinContenido.Count)" -Level WARN -LogFile $logFile
Write-Log "  Con error       : $($errores.Count)"      -Level $(if ($errores.Count -gt 0) {"ERROR"} else {"OK"}) -LogFile $logFile
Write-Log "  Salteados       : $($aSaltear.Count)"     -Level WARN -LogFile $logFile
Write-Log "  Total archivos  : $(($ok | Measure-Object -Property Archivos -Sum).Sum)" -Level OK -LogFile $logFile

if ($errores.Count -gt 0) {
    Write-Log "" -Level INFO
    Write-Log "Sitios con error:" -Level ERROR -LogFile $logFile
    $errores | ForEach-Object {
        Write-Log "  [$($_.Orden)] $($_.Nombre) — $($_.Estado)" -Level ERROR -LogFile $logFile
    }
}

Write-Log "  Resumen: $resumenPath" -Level OK -LogFile $logFile
Write-Log "  Log    : $logFile"     -Level OK -LogFile $logFile
Write-Log "======================================================" -Level HEADER -LogFile $logFile
Write-Log "Script C Refresh finalizado." -Level OK -LogFile $logFile
