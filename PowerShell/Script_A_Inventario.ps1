#Requires -Modules PnP.PowerShell
<#
===============================================================================
  Script_A_Inventario.ps1  |  Circo Studio  |  Kit de Curaduría M365
===============================================================================

PROPÓSITO:
  Genera el inventario completo de sitios y subsitios del tenant SharePoint.
  Es el punto de partida de todo engagement. Su CSV es el input del Script B.

USO:
  .\Script_A_Inventario.ps1 -ConfigPath "..\..\Config\cliente_config.json"

ARCHIVOS GENERADOS:
  [Cliente]_A_Inventario_Sitios_[timestamp].csv
  [Cliente]_A_Inventario_[timestamp]_LOG.txt

PRÓXIMO PASO:
  Abrir el CSV, completar la columna "Relevar_Documentos" con "SI" en los
  sitios que se quieren relevar, y ejecutar Script_B_Masivo.ps1.

===============================================================================
#>

param(
    [Parameter(Mandatory, HelpMessage="Path al archivo cliente_config.json")]
    [string]$ConfigPath,

    [switch]$ForceAuth,

    # Incluir subsitios en el inventario (default: sí)
    [bool]$IncludeSubsites = $true
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
$logFile   = "$($config.CarpetaLocal)\${prefijo}_A_Inventario_${timestamp}_LOG.txt"
$outputCsv = "$($config.CarpetaLocal)\${prefijo}_A_Inventario_Sitios_${timestamp}.csv"
New-Item -ItemType File -Path $logFile -Force | Out-Null

Write-Log "======================================================" -Level HEADER -LogFile $logFile
Write-Log "  Kit M365 Curaduría — Script A Inventario"             -Level HEADER -LogFile $logFile
Write-Log "  Cliente : $($config.NombreCliente)"                   -Level HEADER -LogFile $logFile
Write-Log "  Tenant  : $($config.AdminUrl)"                        -Level HEADER -LogFile $logFile
Write-Log "======================================================" -Level HEADER -LogFile $logFile

# ============================================================
# LOGIN
# ============================================================
try { Connect-TenantOnce -Config $config -LogFile $logFile -ForceAuth:$ForceAuth }
catch { Write-Log "No se pudo iniciar sesión." -Level ERROR -LogFile $logFile; exit 1 }

# ============================================================
# OBTENER SITIOS
# ============================================================
Write-Log "Obteniendo sitios del tenant..." -Level INFO -LogFile $logFile

$allSites = $null
try {
    Connect-PnPOnline -Url $config.AdminUrl -ClientId $config.ClientId -PersistLogin
    $allSites = Get-PnPTenantSite -IncludeOneDriveSites:$false -Detailed
    Write-Log "Sitios encontrados: $($allSites.Count)" -Level OK -LogFile $logFile
}
catch {
    Write-Log "Error al obtener sitios: $_" -Level ERROR -LogFile $logFile
    exit 1
}

$results = [System.Collections.Generic.List[PSObject]]::new()
$counter = 0

foreach ($site in $allSites) {
    $counter++
    if ($counter % 10 -eq 0) {
        Write-Log "Procesados $counter de $($allSites.Count)..." -Level INFO -LogFile $logFile
    }

    # Tipo de sitio
    $siteType = switch ($site.Template) {
        { $_ -like "GROUP*" }              { "Team Site (M365 Group)" }
        { $_ -like "SITEPAGEPUBLISHING*" } { "Communication Site" }
        { $_ -like "STS*" }               { "Team Site (clásico)" }
        { $_ -like "TEAMCHANNEL*" }        { "Canal de Teams" }
        default                            { $site.Template }
    }

    # Clasificación — GR- vs. gobernado
    $rootUrl = $config.SharePointUrl
    # Clasificacion -- basada en template de SharePoint (universal)
    # No depende del naming convention del cliente
    $templateBase = $site.Template -replace "#.*$", ""

    $esLegacyTemplate = $templateBase -in @("GROUP", "TEAMCHANNEL")

    # Prefijos de nombre como criterio adicional opcional (si el cliente los configuro)
    $esLegacyPrefijo = $false
    if (-not $esLegacyTemplate -and $config.PrefijosLegacy -and $config.PrefijosLegacy.Count -gt 0) {
        foreach ($prefijo in $config.PrefijosLegacy) {
            if ($site.Title -like "$prefijo*") { $esLegacyPrefijo = $true; break }
        }
    }

    $labelLegacy = if ($config.LabelLegacy) { $config.LabelLegacy } else { "Legacy (Teams)" }

    $clasificacion = if ($esLegacyTemplate -or $esLegacyPrefijo) {
        $labelLegacy
    } elseif ($site.Url -eq $rootUrl -or $site.Url -eq "$rootUrl/") {
        "Hub Site (raiz)"
    } elseif ($site.IsHubSite) {
        "Hub Site"
    } elseif ($site.HubSiteId -ne "00000000-0000-0000-0000-000000000000") {
        "Spoke Site (Hub asociado)"
    } else {
        "Sitio independiente"
    }

    $storageGB = [math]::Round($site.StorageUsageCurrent / 1024, 2)
    $owner     = if ($site.Owner) { $site.Owner } else { "Sin owner" }

    $row = [PSCustomObject]@{
        "Nombre"             = $site.Title
        "URL"                = $site.Url
        "Clasificacion"      = $clasificacion
        "Tipo_Sitio"         = $siteType
        "Template"           = $site.Template
        "Tiene_Teams"        = if ($site.GroupId -ne "00000000-0000-0000-0000-000000000000") { "Sí" } else { "No" }
        "Es_HubSite"         = if ($site.IsHubSite) { "Sí" } else { "No" }
        "HubSite_ID"         = if ($site.HubSiteId -ne "00000000-0000-0000-0000-000000000000") { $site.HubSiteId } else { "" }
        "Storage_GB"         = $storageGB
        "Storage_MB"         = [math]::Round($site.StorageUsageCurrent, 0)
        "Owner"              = $owner
        "Ultima_Actividad"   = if ($site.LastContentModifiedDate) { $site.LastContentModifiedDate.ToString("yyyy-MM-dd") } else { "" }
        "Fecha_Creacion"     = if ($site.CreatedDate) { $site.CreatedDate.ToString("yyyy-MM-dd") } else { "" }
        "Sharing_Externo"    = $site.SharingCapability
        "Bloqueado"          = if ($site.LockState -eq "Unlock") { "No" } else { $site.LockState }
        "Relevar_Documentos" = ""  # Completar manualmente con SI antes de correr Script B
    }
    $results.Add($row)
}

# ============================================================
# SUBSITIOS (OPCIONAL)
# ============================================================
$subsiteResults = [System.Collections.Generic.List[PSObject]]::new()
$subCounter     = 0

if ($IncludeSubsites) {
    Write-Log "Buscando subsitios..." -Level INFO -LogFile $logFile

    foreach ($site in $allSites) {
        try {
            Connect-PnPOnline -Url $site.Url -ClientId $config.ClientId -PersistLogin
            $subsites = Get-PnPSubWeb -Recurse -ErrorAction SilentlyContinue

            if ($subsites -and @($subsites).Count -gt 0) {
                foreach ($sub in $subsites) {
                    $subCounter++
                    $subRow = [PSCustomObject]@{
                        "Nombre"             = $sub.Title
                        "URL"                = $sub.Url
                        "Clasificacion"      = "Subsitio"
                        "Tipo_Sitio"         = "Subsitio de $($site.Title)"
                        "Template"           = $sub.WebTemplate
                        "Tiene_Teams"        = "No"
                        "Es_HubSite"         = "No"
                        "HubSite_ID"         = ""
                        "Storage_GB"         = ""
                        "Storage_MB"         = ""
                        "Owner"              = ""
                        "Ultima_Actividad"   = if ($sub.LastItemModifiedDate) { $sub.LastItemModifiedDate.ToString("yyyy-MM-dd") } else { "" }
                        "Fecha_Creacion"     = if ($sub.Created) { $sub.Created.ToString("yyyy-MM-dd") } else { "" }
                        "Sharing_Externo"    = ""
                        "Bloqueado"          = "No"
                        "Relevar_Documentos" = ""
                    }
                    $subsiteResults.Add($subRow)
                }
            }
        }
        catch {
            Write-Log "No se pudieron obtener subsitios de $($site.Url): $_" -Level WARN -LogFile $logFile
        }
    }
    Write-Log "Subsitios encontrados: $subCounter" -Level OK -LogFile $logFile
}

# ============================================================
# EXPORTAR CSV
# ============================================================
$allResults = [System.Collections.Generic.List[PSObject]]::new()
foreach ($r in $results)        { $allResults.Add($r) }
foreach ($r in $subsiteResults) { $allResults.Add($r) }

$sorted = $allResults | Sort-Object @{Expression="Clasificacion"}, @{Expression="Storage_GB"; Descending=$true}
$sorted | Export-Csv -Path $outputCsv -NoTypeInformation -Encoding UTF8
Write-Log "CSV exportado: $outputCsv" -Level OK -LogFile $logFile

# ============================================================
# RESUMEN
# ============================================================
Write-Log "======================================================" -Level HEADER -LogFile $logFile
Write-Log "  RESUMEN"                                               -Level HEADER -LogFile $logFile
Write-Log "======================================================" -Level HEADER -LogFile $logFile

$allResults | Group-Object Clasificacion | Sort-Object Count -Descending | ForEach-Object {
    Write-Log ("  {0,-32} {1,4} sitios" -f $_.Name, $_.Count) -Level INFO -LogFile $logFile
}

$totalGB = [math]::Round(($results | Measure-Object -Property Storage_GB -Sum).Sum, 2)
Write-Log "------------------------------------------------------" -Level HEADER -LogFile $logFile
Write-Log "  Total sitios    : $($results.Count)"                  -Level OK     -LogFile $logFile
Write-Log "  Total subsitios : $subCounter"                        -Level OK     -LogFile $logFile
Write-Log "  Storage total   : $totalGB GB"                        -Level OK     -LogFile $logFile
Write-Log "------------------------------------------------------" -Level HEADER -LogFile $logFile
Write-Log "  PRÓXIMO PASO:"                                         -Level HEADER -LogFile $logFile
Write-Log "  Abrir el CSV y completar 'Relevar_Documentos' = SI"   -Level INFO   -LogFile $logFile
Write-Log "  en los sitios a relevar, luego correr Script_B_Masivo.ps1" -Level INFO -LogFile $logFile
Write-Log "======================================================" -Level HEADER -LogFile $logFile
Write-Log "Script A finalizado." -Level OK -LogFile $logFile
