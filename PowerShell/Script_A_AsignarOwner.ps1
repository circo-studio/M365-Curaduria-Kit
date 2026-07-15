#Requires -Modules PnP.PowerShell
<#
===============================================================================
  Script_A_AsignarOwner.ps1  |  Circo Studio  |  Kit de Curaduría M365
===============================================================================

PROPÓSITO:
  Asigna la cuenta B2B de Circo Studio como Site Collection Administrator en
  los sitios Legacy (Teams) pendientes de relevamiento, para que Script_B_Masivo pueda
  acceder a su contenido.

  Aplica los mismos filtros de exclusión que Script_B_Masivo (sitios ya
  analizados y canales de padres cubiertos), ordenados por storage descendente.

USO:
  .\Script_A_AsignarOwner.ps1 -ConfigPath "..\..\Config\cliente_config.json"

ARCHIVOS GENERADOS:
  [Cliente]_A_AsignarOwner_[timestamp]_LOG.csv   — resultado por sitio

===============================================================================
#>

param(
    [Parameter(Mandatory, HelpMessage="Path al archivo cliente_config.json")]
    [string]$ConfigPath,

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

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$prefijo   = ($config.NombreCliente -replace "[^a-zA-Z0-9]", "_")
$logCsv    = "$($config.CarpetaLocal)\${prefijo}_A_AsignarOwner_${timestamp}_LOG.csv"

# Header del log CSV
"Orden,Nombre,URL,Tipo,Storage_GB,Resultado,Detalle,Timestamp" |
    Out-File -FilePath $logCsv -Encoding UTF8

Write-Log "======================================================" -Level HEADER
Write-Log "  Kit M365 Curaduría — Script A AsignarOwner"           -Level HEADER
Write-Log "  Cliente       : $($config.NombreCliente)"             -Level HEADER
Write-Log "  Cuenta Circo  : $($config.UPNCirco)"                  -Level HEADER
Write-Log "  Log de salida : $logCsv"                              -Level HEADER
Write-Log "======================================================" -Level HEADER

# ============================================================
# BUSCAR CSV DE INVENTARIO
# ============================================================
$csvInventario = Get-ChildItem $config.CarpetaLocal -Filter "${prefijo}_A_Inventario_Sitios_*.csv" |
                 Sort-Object LastWriteTime -Descending |
                 Select-Object -First 1

if (-not $csvInventario) {
    Write-Log "No se encontró el CSV de inventario. Correr primero Script_A_Inventario.ps1" -Level ERROR
    exit 1
}

Write-Log "CSV de inventario: $($csvInventario.Name)" -Level OK

# ============================================================
# FILTRAR SITIOS
# ============================================================
$sitiosPendientes = $null
try {
    $sitiosPendientes = Get-SitiosFiltrados `
        -CsvInventario $csvInventario.FullName `
        -Config        $config
}
catch {
    Write-Log "Error al filtrar sitios: $_" -Level ERROR
    exit 1
}

if (-not $sitiosPendientes -or @($sitiosPendientes).Count -eq 0) {
    Write-Log "No hay sitios pendientes de asignación de owner." -Level WARN
    exit 0
}

Write-Log "Sitios a procesar: $(@($sitiosPendientes).Count)" -Level OK
Write-Log "Storage total: $([math]::Round(($sitiosPendientes | Measure-Object -Property _StorageNum -Sum).Sum, 1)) GB" -Level INFO

# ============================================================
# CONECTAR AL ADMIN CENTER (una sola vez)
# ============================================================
Write-Log "" -Level INFO
Write-Log "Conectando al Admin Center..." -Level INFO

try {
    $connectParams = @{
        Url      = $config.AdminUrl
        ClientId = $config.ClientId
        Interactive = $true
    }
    if ($ForceAuth) { $connectParams["ForceAuthentication"] = $true }
    else            { $connectParams["PersistLogin"] = $true }

    Connect-PnPOnline @connectParams -ErrorAction Stop
    Write-Log "Conexión al Admin Center OK" -Level OK
}
catch {
    Write-Log "Error al conectar al Admin Center: $_" -Level ERROR
    exit 1
}

# ============================================================
# LOOP PRINCIPAL
# ============================================================
$total    = @($sitiosPendientes).Count
$exitosos = 0
$fallidos = 0
$orden    = 0

foreach ($sitio in $sitiosPendientes) {
    $orden++
    $tipo   = if ($sitio.Tipo_Sitio -eq "Canal de Teams") { "CANAL" } else { "TeamSite" }
    $ts_item = Get-Date -Format "HH:mm:ss"

    Write-Log "" -Level INFO
    Write-Log "[$orden/$total] $($sitio.Nombre)" -Level INFO
    Write-Log "  Tipo    : $tipo  |  Storage: $($sitio.Storage_GB) GB" -Level INFO
    Write-Log "  URL     : $($sitio.URL)" -Level INFO

    try {
        # Set-PnPTenantSite -Owners AGREGA sin reemplazar owners existentes
        Set-PnPTenantSite -Url $sitio.URL -Owners $config.UPNCirco -ErrorAction Stop
        Write-Log "  [OK] Owner asignado" -Level OK

        "`"$orden`",`"$($sitio.Nombre)`",`"$($sitio.URL)`",`"$tipo`",`"$($sitio.Storage_GB)`",`"OK`",`"Owner asignado`",`"$ts_item`"" |
            Out-File -FilePath $logCsv -Encoding UTF8 -Append

        $exitosos++
    }
    catch {
        $errorMsg = $_.Exception.Message -replace '"', "'" -replace "`r`n|`n", " "
        Write-Log "  [ERROR] $errorMsg" -Level ERROR

        "`"$orden`",`"$($sitio.Nombre)`",`"$($sitio.URL)`",`"$tipo`",`"$($sitio.Storage_GB)`",`"ERROR`",`"$errorMsg`",`"$ts_item`"" |
            Out-File -FilePath $logCsv -Encoding UTF8 -Append

        $fallidos++
    }

    # Pausa anti-throttling
    Start-Sleep -Milliseconds 300
}

# ============================================================
# RESUMEN FINAL
# ============================================================
Write-Log "" -Level INFO
Write-Log "======================================================" -Level HEADER
Write-Log "  RESUMEN FINAL"                                         -Level HEADER
Write-Log "  Total procesados  : $total"                           -Level INFO
Write-Log "  Exitosos          : $exitosos"                        -Level OK
Write-Log "  Con error         : $fallidos"                        -Level $(if ($fallidos -gt 0) {"ERROR"} else {"OK"})
Write-Log "  Log guardado en   : $logCsv"                          -Level INFO
Write-Log "======================================================" -Level HEADER

if ($fallidos -gt 0) {
    Write-Log "Sitios con error — revisar el log antes de correr Script_B_Masivo:" -Level WARN
    Import-Csv -Path $logCsv | Where-Object { $_.Resultado -eq "ERROR" } | ForEach-Object {
        Write-Log "  [$($_.Orden)] $($_.Nombre)" -Level ERROR
        Write-Log "       $($_.Detalle)"          -Level ERROR
    }
}

Write-Log "Script A AsignarOwner finalizado." -Level OK
