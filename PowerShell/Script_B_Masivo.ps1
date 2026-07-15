#Requires -Modules PnP.PowerShell
<#
===============================================================================
  Script_B_Masivo.ps1  |  Circo Studio  |  Kit de Curaduría Informacional M365
===============================================================================

PROPÓSITO:
  Relevamiento masivo de documentos sobre todos los sitios Legacy (Teams) pendientes
  del tenant. Lee la configuración del cliente desde un JSON externo y aplica
  los filtros de exclusión definidos en él.

  Genera un CSV por sitio + un CSV de resumen de la ejecución.
  Es idempotente: si el CSV de un sitio ya existe en la carpeta de salida,
  lo saltea sin reprocesar.

USO:
  .\Script_B_Masivo.ps1 -ConfigPath "..\..\Config\cliente_config.json"

PARÁMETROS:
  -ConfigPath   (obligatorio) Path al cliente_config.json
  -ForceAuth    (switch)      Fuerza re-autenticación aunque haya token cacheado
  -SoloSitio    (opcional)    URL de un único sitio — útil para debug o reintento puntual

ARCHIVOS GENERADOS en la carpeta configurada en el JSON:
  [Cliente]_B_Docs_[NombreSitio]_[timestamp].csv   — uno por sitio procesado
  [Cliente]_B_Masivo_Resumen_[timestamp].csv        — estado de la ejecución completa
  [Cliente]_B_Masivo_[timestamp]_LOG.txt            — log detallado

NOTAS:
  - El script puede tardar varias horas en tenants grandes. No cerrar la sesión.
  - Si se interrumpe, volver a ejecutar con el mismo -ConfigPath.
    Los sitios con CSV previo se saltean automáticamente (checkpoint).
  - Para reprocesar un sitio que ya fue relevado, eliminar su CSV de la
    carpeta de salida y volver a ejecutar.

===============================================================================
#>

param(
    [Parameter(Mandatory, HelpMessage="Path al archivo cliente_config.json")]
    [string]$ConfigPath,

    [switch]$ForceAuth,

    [string]$SoloSitio = ""
)

# Importar el módulo de funciones compartidas
$moduloPath = Join-Path $PSScriptRoot "M365Curaduria.psm1"
if (-not (Test-Path $moduloPath)) {
    Write-Host "[ERROR] No se encontró el módulo M365Curaduria.psm1 en: $moduloPath" -ForegroundColor Red
    Write-Host "        Asegurarse de que Script_B_Masivo.ps1 y M365Curaduria.psm1 estén en la misma carpeta." -ForegroundColor Red
    exit 1
}
Import-Module $moduloPath -Force

# ============================================================
# CARGA DE CONFIGURACIÓN
# ============================================================
$config = $null
try {
    $config = Read-ClientConfig -ConfigPath $ConfigPath
}
catch {
    Write-Host "[ERROR] $_" -ForegroundColor Red
    exit 1
}

# ============================================================
# SETUP DE CARPETAS Y LOG
# ============================================================
if (-not (Test-Path $config.CarpetaLocal)) {
    New-Item -ItemType Directory -Path $config.CarpetaLocal -Force | Out-Null
}

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$prefijo   = ($config.NombreCliente -replace "[^a-zA-Z0-9]", "_")
$logFile   = "$($config.CarpetaLocal)\${prefijo}_B_Masivo_${timestamp}_LOG.txt"
New-Item -ItemType File -Path $logFile -Force | Out-Null

Write-Log "======================================================" -Level HEADER -LogFile $logFile
Write-Log "  Kit M365 Curaduría — Script B Masivo"                 -Level HEADER -LogFile $logFile
Write-Log "  Cliente  : $($config.NombreCliente)"                  -Level HEADER -LogFile $logFile
Write-Log "  Tenant   : $($config.SharePointUrl)"                  -Level HEADER -LogFile $logFile
Write-Log "  Config   : $ConfigPath"                               -Level HEADER -LogFile $logFile
Write-Log "  Salida   : $($config.CarpetaLocal)"                   -Level HEADER -LogFile $logFile
Write-Log "======================================================" -Level HEADER -LogFile $logFile

# ============================================================
# BUSCAR CSV DE INVENTARIO (Script A)
# ============================================================
$csvInventario = Get-ChildItem $config.CarpetaLocal -Filter "${prefijo}_A_Inventario_Sitios_*.csv" |
                 Sort-Object LastWriteTime -Descending |
                 Select-Object -First 1

if (-not $csvInventario) {
    Write-Log "No se encontró el CSV de inventario (Script A) en: $($config.CarpetaLocal)" -Level ERROR -LogFile $logFile
    Write-Log "Correr primero Script_A_Inventario.ps1" -Level ERROR -LogFile $logFile
    exit 1
}

Write-Log "CSV de inventario: $($csvInventario.Name)" -Level OK -LogFile $logFile

# ============================================================
# FILTRAR SITIOS A PROCESAR
# ============================================================
$sitiosPendientes = $null

# Si se pasó -SoloSitio, construir lista de uno
if ($SoloSitio) {
    Write-Log "Modo sitio único: $SoloSitio" -Level WARN -LogFile $logFile
    $filaFake = [PSCustomObject]@{
        Nombre       = $SoloSitio
        URL          = $SoloSitio
        Tipo_Sitio   = "TeamSite"
        Storage_GB   = "0"
        Clasificacion = if ($config.LabelLegacy) { $config.LabelLegacy } else { "Legacy (Teams)" }
        Bloqueado    = "No"
        _StorageNum  = 0
    }
    $sitiosPendientes = @($filaFake)
}
else {
    try {
        $sitiosPendientes = Get-SitiosFiltrados `
            -CsvInventario $csvInventario.FullName `
            -Config        $config `
            -LogFile       $logFile
    }
    catch {
        Write-Log "Error al filtrar sitios: $_" -Level ERROR -LogFile $logFile
        exit 1
    }
}

if (-not $sitiosPendientes -or @($sitiosPendientes).Count -eq 0) {
    Write-Log "No hay sitios pendientes de relevamiento. Verificar el CSV de inventario." -Level WARN -LogFile $logFile
    exit 0
}

# ============================================================
# LOGIN ÚNICO
# ============================================================
try {
    Connect-TenantOnce -Config $config -LogFile $logFile -ForceAuth:$ForceAuth
}
catch {
    Write-Log "No se pudo iniciar sesión. Abortando." -Level ERROR -LogFile $logFile
    exit 1
}

# ============================================================
# LOOP PRINCIPAL
# ============================================================
$resumen = [System.Collections.Generic.List[PSObject]]::new()
$total   = @($sitiosPendientes).Count
$orden   = 0

foreach ($sitio in $sitiosPendientes) {
    $orden++
    $tipo = if ($sitio.Tipo_Sitio -eq "Canal de Teams") { "CANAL" } else { "TeamSite" }

    Write-Log "" -Level INFO
    Write-Log "[$orden/$total] $($sitio.Nombre) — $($sitio.Storage_GB) GB [$tipo]" -Level HEADER -LogFile $logFile

    $inicio    = Get-Date
    $resultado = Invoke-RelevamientoSitio `
        -SiteUrl      $sitio.URL       `
        -SiteName     $sitio.Nombre    `
        -CarpetaSalida $config.CarpetaLocal `
        -Config        $config         `
        -LogFile       $logFile        `
        -Timestamp     $timestamp

    $duracion = [math]::Round(((Get-Date) - $inicio).TotalSeconds)

    $filaResumen = [PSCustomObject]@{
        "Orden"        = $orden
        "Nombre"       = $sitio.Nombre
        "URL"          = $sitio.URL
        "Tipo"         = $tipo
        "Storage_GB"   = $sitio.Storage_GB
        "Bibliotecas"  = $resultado.Bibliotecas
        "Archivos"     = $resultado.Archivos
        "Errores"      = $resultado.Errores
        "Duracion_Seg" = $duracion
        "Estado"       = $resultado.Estado
        "CSV_Salida"   = $resultado.CsvPath
    }
    $resumen.Add($filaResumen)

    # Guardar resumen parcial después de cada sitio (por si se interrumpe)
    $resumenPath = "$($config.CarpetaLocal)\${prefijo}_B_Masivo_Resumen_${timestamp}.csv"
    $resumen | Export-Csv -Path $resumenPath -NoTypeInformation -Encoding UTF8
}

# ============================================================
# RESUMEN FINAL
# ============================================================
$ok           = @($resumen | Where-Object { $_.Estado -eq "OK" })
$salteados    = @($resumen | Where-Object { $_.Estado -eq "Salteado" })
$sinContenido = @($resumen | Where-Object { $_.Estado -in @("SinArchivos","SinBibliotecas") })
$errores      = @($resumen | Where-Object { $_.Estado -like "Error*" })

Write-Log "" -Level INFO
Write-Log "======================================================" -Level HEADER -LogFile $logFile
Write-Log "  RESUMEN FINAL"                                         -Level HEADER -LogFile $logFile
Write-Log "======================================================" -Level HEADER -LogFile $logFile
Write-Log "  Procesados OK             : $($ok.Count)"             -Level OK     -LogFile $logFile
Write-Log "  Salteados (CSV previo)    : $($salteados.Count)"      -Level WARN   -LogFile $logFile
Write-Log "  Sin contenido             : $($sinContenido.Count)"   -Level WARN   -LogFile $logFile
Write-Log "  Con error                 : $($errores.Count)"        -Level $(if ($errores.Count -gt 0) { "ERROR" } else { "OK" }) -LogFile $logFile
Write-Log "  Total archivos relevados  : $(($ok | Measure-Object -Property Archivos -Sum).Sum)" -Level OK -LogFile $logFile
Write-Log "------------------------------------------------------" -Level HEADER -LogFile $logFile

if ($errores.Count -gt 0) {
    Write-Log "Sitios con error — revisar manualmente:" -Level ERROR -LogFile $logFile
    $errores | ForEach-Object {
        Write-Log "  [$($_.Orden)] $($_.Nombre) — $($_.Estado)" -Level ERROR -LogFile $logFile
    }
}

Write-Log "  Resumen: $resumenPath" -Level OK   -LogFile $logFile
Write-Log "  Log    : $logFile"     -Level OK   -LogFile $logFile
Write-Log "======================================================" -Level HEADER -LogFile $logFile
Write-Log "Script B Masivo finalizado." -Level OK -LogFile $logFile
