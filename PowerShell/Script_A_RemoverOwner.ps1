#Requires -Modules PnP.PowerShell
<#
===============================================================================
  Script_A_RemoverOwner.ps1  |  Circo Studio  |  Kit de Curaduría M365
===============================================================================

PROPÓSITO:
  Remueve la cuenta B2B de Circo Studio como Site Collection Administrator de
  los sitios donde la asignó Script_A_AsignarOwner.ps1. Es la operación de
  cierre / limpieza posterior al relevamiento.

  Se alimenta del LOG CSV que genera Script_A_AsignarOwner.ps1 y procesa
  ÚNICAMENTE las filas con Remover=SI. Los sitios donde la cuenta ya figuraba
  como owner antes del relevamiento (Resultado=YA_ERA_OWNER, Remover=NO) no se
  tocan: ese acceso es preexistente del cliente.

  Además soporta un CSV de exclusión de URLs para dejar afuera sitios puntuales.

USO:
  # Usa el LOG más reciente de Script_A_AsignarOwner en CarpetaLocal
  .\Script_A_RemoverOwner.ps1 -ConfigPath "..\Config\cliente_config.json"

  # Apuntando a un LOG específico
  .\Script_A_RemoverOwner.ps1 -ConfigPath ".\cliente_config.json" `
      -CsvAsignacion ".\Cliente_A_AsignarOwner_20260115_101500_LOG.csv"

  # Simulación (no aplica cambios)
  .\Script_A_RemoverOwner.ps1 -ConfigPath ".\cliente_config.json" -Simular

  # Fallback: recorrer el inventario completo en lugar del LOG
  .\Script_A_RemoverOwner.ps1 -ConfigPath ".\cliente_config.json" -Fuente Inventario

  # Generar la plantilla del CSV de exclusión y salir
  .\Script_A_RemoverOwner.ps1 -ConfigPath ".\cliente_config.json" -GenerarPlantillaExclusiones

CSV DE ENTRADA (-Fuente LogAsignacion, default):
  Salida de Script_A_AsignarOwner.ps1:
    [Cliente]_A_AsignarOwner_[timestamp]_LOG.csv
  Se toman las filas con Remover=SI. Si el CSV es de una versión anterior del
  kit (sin columna Remover), se usan las filas con Resultado=OK.

CSV DE EXCLUSIÓN (opcional, -ExclusionesCsv):
  - Si no se pasa, busca el más reciente "[Cliente]_Exclusiones_URLs*.csv"
    en CarpetaLocal. Si no existe, no excluye nada.
  - Acepta CSV con header (columna URL / Url / SiteUrl / Sitio_URL / Link) con
    delimitador "," o ";", o un TXT/CSV plano con una URL por línea.
  - Líneas que empiezan con # se ignoran (comentarios).
  - Matcheo case-insensitive, ignora la barra final.
  - Soporta wildcards: https://tenant.sharepoint.com/sites/GR-Finanzas*
  - Sin wildcard, también excluye los subsitios de esa URL.
  - Con -InvertirFiltro el CSV pasa a ser lista de INCLUSIÓN (solo esas URLs).

ARCHIVOS GENERADOS:
  [Cliente]_A_RemoverOwner_[timestamp]_LOG.csv   — resultado por sitio
    Resultado: OK | ERROR | OMITIDO | SIMULADO | NO_ERA_ADMIN

NOTA:
  Remueve el rol de Site Collection Administrator (el inverso de
  Set-PnPTenantSite -Owners). Si además se agregó la cuenta como owner del
  grupo de Microsoft 365 del Team, eso se remueve por separado con
  Remove-PnPMicrosoft365GroupOwner.

===============================================================================
#>

param(
    [Parameter(Mandatory, HelpMessage="Path al archivo cliente_config.json")]
    [string]$ConfigPath,

    # Origen de la lista de sitios
    #   LogAsignacion — salida de Script_A_AsignarOwner (default, recomendado)
    #   Inventario    — CSV de Script_A_Inventario (fallback / limpieza total)
    [ValidateSet("LogAsignacion","Inventario")]
    [string]$Fuente = "LogAsignacion",

    # LOG CSV de Script_A_AsignarOwner a usar (si se omite, toma el más reciente)
    [string]$CsvAsignacion,

    # CSV con URLs a excluir (si se omite, se autodetecta en CarpetaLocal)
    [string]$ExclusionesCsv,

    # Nombre de la columna de URL en el CSV de exclusión (si se omite, se autodetecta)
    [string]$ColumnaUrl,

    # Invierte la lógica del CSV de exclusión: pasa a ser lista de inclusión
    [switch]$InvertirFiltro,

    # Solo con -Fuente Inventario: restringe a sitios Legacy (Teams)
    [switch]$SoloLegacy,

    # No ejecuta cambios: solo reporta qué haría
    [switch]$Simular,

    # Saltea la confirmación interactiva
    [switch]$SinConfirmacion,

    # Crea la plantilla del CSV de exclusión y termina
    [switch]$GenerarPlantillaExclusiones,

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
$logCsv    = "$($config.CarpetaLocal)\${prefijo}_A_RemoverOwner_${timestamp}_LOG.csv"

# Si se pasó un CSV de asignación explícito, la fuente queda determinada
if ($CsvAsignacion) { $Fuente = "LogAsignacion" }

# ------------------------------------------------------------
# Plantilla del CSV de exclusión (y salida temprana)
# ------------------------------------------------------------
if ($GenerarPlantillaExclusiones) {
    $plantilla = "$($config.CarpetaLocal)\${prefijo}_Exclusiones_URLs.csv"
    if (Test-Path $plantilla) {
        Write-Host "[WARN] Ya existe: $plantilla — no se sobrescribe." -ForegroundColor Yellow
    }
    else {
        @(
            "# Kit de Curaduria M365 - URLs excluidas de Script_A_RemoverOwner"
            "# Una URL por fila. Se admiten wildcards con *. Las lineas con # se ignoran."
            "URL,Motivo"
            "$($config.SharePointUrl)/sites/EJEMPLO,Relevamiento pendiente"
        ) | Out-File -FilePath $plantilla -Encoding UTF8
        Write-Host "[OK] Plantilla creada: $plantilla" -ForegroundColor Green
    }
    exit 0
}

# Header del log CSV
"Orden,Nombre,URL,Tipo,Storage_GB,Resultado,Detalle,Timestamp" |
    Out-File -FilePath $logCsv -Encoding UTF8

function Add-FilaLog {
    param(
        [int]$Orden, [string]$Nombre, [string]$URL, [string]$Tipo,
        [string]$Storage, [string]$Resultado, [string]$Detalle
    )
    $n = (($Nombre  -replace '"', "'") -replace "`r`n|`n", " ")
    $d = (($Detalle -replace '"', "'") -replace "`r`n|`n", " ")
    $ts = Get-Date -Format "HH:mm:ss"
    "`"$Orden`",`"$n`",`"$URL`",`"$Tipo`",`"$Storage`",`"$Resultado`",`"$d`",`"$ts`"" |
        Out-File -FilePath $script:logCsv -Encoding UTF8 -Append
}

$modoTexto = if ($Simular) { "SIMULACIÓN (no aplica cambios)" } else { "EJECUCIÓN REAL" }

Write-Log "======================================================" -Level HEADER
Write-Log "  Kit M365 Curaduría — Script A RemoverOwner"            -Level HEADER
Write-Log "  Cliente       : $($config.NombreCliente)"             -Level HEADER
Write-Log "  Cuenta Circo  : $($config.UPNCirco)"                  -Level HEADER
Write-Log "  Modo          : $modoTexto"                           -Level $(if ($Simular) {"WARN"} else {"HEADER"})
Write-Log "  Fuente        : $Fuente"                              -Level HEADER
Write-Log "  Log de salida : $logCsv"                              -Level HEADER
Write-Log "======================================================" -Level HEADER

# ============================================================
# CARGAR FILTRO DE URLs DESDE CSV
# ============================================================
function Get-UrlsDesdeCsv {
    param(
        [Parameter(Mandatory)][string]$Path,
        [string]$Columna = ""
    )

    $urls = [System.Collections.Generic.List[string]]::new()

    $lineas = @(Get-Content -Path $Path -Encoding UTF8 |
                Where-Object { $_.Trim() -ne "" -and -not $_.Trim().StartsWith("#") })

    if ($lineas.Count -eq 0) { return $urls }

    # Detectar delimitador por la primera línea útil
    $primera = $lineas[0]
    $delim = if ((($primera -split ';').Count) -gt (($primera -split ',').Count)) { ';' } else { ',' }

    # ¿Tiene header? Si la primera línea ya es una URL, no lo tiene
    $tieneHeader = $primera -notmatch '^\s*"?\s*https?://'

    if ($tieneHeader) {
        $datos = @($lineas | ConvertFrom-Csv -Delimiter $delim)
        if ($datos.Count -eq 0) { return $urls }

        $propiedades = @($datos[0].PSObject.Properties.Name)
        $candidatas  = if ($Columna) { @($Columna) }
                       else { @("URL","Url","url","SiteUrl","Site_URL","Sitio_URL","Sitio","Link","Direccion") }

        $col = $null
        foreach ($c in $candidatas) {
            $match = $propiedades | Where-Object { $_ -eq $c } | Select-Object -First 1
            if ($match) { $col = $match; break }
        }

        # Fallback: primera columna cuyo valor parezca una URL
        if (-not $col) {
            foreach ($p in $propiedades) {
                if ([string]$datos[0].$p -match '^\s*https?://') { $col = $p; break }
            }
        }

        if (-not $col) {
            throw "No se encontró una columna de URL en $Path (columnas: $($propiedades -join ', ')). Usar -ColumnaUrl."
        }

        foreach ($fila in $datos) {
            if ($fila.$col) { $urls.Add([string]$fila.$col) }
        }
    }
    else {
        # Lista plana: una URL por línea (se toma la primera columna si hay más)
        foreach ($l in $lineas) { $urls.Add((($l -split $delim)[0])) }
    }

    # Normalizar: sin comillas, sin espacios, sin barra final, minúsculas
    $limpias = [System.Collections.Generic.List[string]]::new()
    foreach ($u in $urls) {
        $n = ([string]$u).Trim().Trim('"').Trim().TrimEnd('/').ToLower()
        if ($n -and $n -notin $limpias) { $limpias.Add($n) }
    }

    return $limpias
}

$patronesCsv = @()
$csvFiltro   = $null

if ($ExclusionesCsv) {
    if (-not (Test-Path $ExclusionesCsv)) {
        Write-Log "No se encontró el CSV de filtro: $ExclusionesCsv" -Level ERROR
        exit 1
    }
    $csvFiltro = (Resolve-Path $ExclusionesCsv).Path
}
else {
    $auto = Get-ChildItem $config.CarpetaLocal -Filter "${prefijo}_Exclusiones_URLs*.csv" -ErrorAction SilentlyContinue |
            Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($auto) { $csvFiltro = $auto.FullName }
}

if ($csvFiltro) {
    try { $patronesCsv = @(Get-UrlsDesdeCsv -Path $csvFiltro -Columna $ColumnaUrl) }
    catch { Write-Log "$_" -Level ERROR; exit 1 }

    $rol = if ($InvertirFiltro) { "INCLUSIÓN (procesa solo estas URLs)" } else { "EXCLUSIÓN (no se tocan)" }
    Write-Log "CSV de filtro     : $(Split-Path $csvFiltro -Leaf)" -Level OK
    Write-Log "  URLs cargadas   : $($patronesCsv.Count)  —  rol: $rol" -Level INFO
}
else {
    Write-Log "Sin CSV de filtro de URLs (no se excluye ninguna)." -Level WARN
    if ($InvertirFiltro) {
        Write-Log "-InvertirFiltro requiere un CSV de URLs. Abortando." -Level ERROR
        exit 1
    }
}

# Devuelve el patrón que matcheó, o $null
function Get-PatronMatch {
    param([string]$Url, [string[]]$Patrones)

    if (-not $Patrones -or $Patrones.Count -eq 0) { return $null }

    $u = ([string]$Url).Trim().TrimEnd('/').ToLower()

    foreach ($p in $Patrones) {
        if ($p -match '[\*\?]') {
            if ($u -like $p) { return $p }
        }
        else {
            # Igualdad exacta o subsitio de la URL listada
            if ($u -eq $p -or $u.StartsWith("$p/")) { return $p }
        }
    }
    return $null
}

# ============================================================
# ARMAR LISTA DE SITIOS
# ============================================================

# Cast invariante (el cast de PowerShell no depende del locale, a diferencia de TryParse)
function ConvertTo-Gb {
    param([string]$Valor)
    if (-not $Valor) { return 0.0 }
    try   { return [double]($Valor -replace ',', '.') }
    catch { return 0.0 }
}

$sitios = [System.Collections.Generic.List[object]]::new()

if ($Fuente -eq "LogAsignacion") {

    # --- Ubicar el LOG de asignación ---
    $pathLog = $null

    if ($CsvAsignacion) {
        if (-not (Test-Path $CsvAsignacion)) {
            Write-Log "No se encontró el CSV de asignación: $CsvAsignacion" -Level ERROR
            exit 1
        }
        $pathLog = (Resolve-Path $CsvAsignacion).Path
    }
    else {
        $ultimo = Get-ChildItem $config.CarpetaLocal -Filter "${prefijo}_A_AsignarOwner_*_LOG.csv" -ErrorAction SilentlyContinue |
                  Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if (-not $ultimo) {
            Write-Log "No se encontró un LOG de Script_A_AsignarOwner en $($config.CarpetaLocal)." -Level ERROR
            Write-Log "Correr primero Script_A_AsignarOwner.ps1, pasar -CsvAsignacion, o usar -Fuente Inventario." -Level ERROR
            exit 1
        }
        $pathLog = $ultimo.FullName
    }

    Write-Log "CSV de asignación : $(Split-Path $pathLog -Leaf)" -Level OK

    $filas = @(Import-Csv -Path $pathLog -Encoding UTF8)
    if ($filas.Count -eq 0) {
        Write-Log "El CSV de asignación está vacío." -Level WARN
        exit 0
    }

    Write-Log "  Filas leídas    : $($filas.Count)" -Level INFO

    # --- Selección de filas a remover ---
    # Formato actual: columna Remover (SI/NO) que ya distingue accesos preexistentes.
    # Formato anterior del kit: sin columna Remover → se usan las filas con Resultado=OK.
    $columnas     = @($filas[0].PSObject.Properties.Name)
    $tieneRemover = $columnas -contains "Remover"

    if ($tieneRemover) {
        $elegibles = @($filas | Where-Object {
            $_.URL -and $_.Remover -eq "SI" -and $_.Resultado -ne "YA_ERA_OWNER"
        })
        $preexistentes = @($filas | Where-Object { $_.Resultado -eq "YA_ERA_OWNER" }).Count

        Write-Log "  Marcadas Remover=SI : $($elegibles.Count)" -Level OK
        if ($preexistentes -gt 0) {
            Write-Log "  Owner preexistente  : $preexistentes (no se tocan)" -Level WARN
        }
    }
    else {
        Write-Log "  El CSV no tiene columna 'Remover' (formato anterior del kit)." -Level WARN
        Write-Log "  Se usan las filas con Resultado=OK — verificar que no incluyan accesos preexistentes." -Level WARN
        $elegibles = @($filas | Where-Object { $_.URL -and $_.Resultado -eq "OK" })
        Write-Log "  Filas Resultado=OK  : $($elegibles.Count)" -Level OK
    }

    $descartadas = $filas.Count - $elegibles.Count
    if ($descartadas -gt 0) {
        Write-Log "  Descartadas         : $descartadas (Remover=NO, errores o sin URL)" -Level INFO
    }

    foreach ($f in $elegibles) {
        $sitios.Add([PSCustomObject]@{
            Nombre      = $f.Nombre
            URL         = $f.URL
            Tipo        = if ($f.Tipo) { $f.Tipo } else { "" }
            Storage_GB  = $f.Storage_GB
            Bloqueado   = "No"
            _StorageNum = ConvertTo-Gb $f.Storage_GB
        })
    }
}
else {

    $csvInventario = Get-ChildItem $config.CarpetaLocal -Filter "${prefijo}_A_Inventario_Sitios_*.csv" -ErrorAction SilentlyContinue |
                     Sort-Object LastWriteTime -Descending | Select-Object -First 1

    if (-not $csvInventario) {
        Write-Log "No se encontró el CSV de inventario. Correr primero Script_A_Inventario.ps1" -Level ERROR
        exit 1
    }

    Write-Log "CSV de inventario : $($csvInventario.Name)" -Level OK
    Write-Log "AVISO: el inventario no distingue accesos preexistentes del cliente." -Level WARN
    Write-Log "       Preferir -Fuente LogAsignacion, o cargar las URLs a preservar en el CSV de exclusión." -Level WARN

    $inventario = @(Import-Csv -Path $csvInventario.FullName -Encoding UTF8)
    Write-Log "  Filas leídas    : $($inventario.Count)" -Level INFO

    foreach ($fila in $inventario) {

        if (-not $fila.URL) { continue }

        # Filtro opcional de sitios Legacy (Teams) — mismo criterio que Get-SitiosFiltrados
        if ($SoloLegacy) {
            $esLegacy = $false

            if ($config.TemplatesLegacy -and $config.TemplatesLegacy.Count -gt 0) {
                foreach ($tmpl in $config.TemplatesLegacy) {
                    if ($fila.Template -like "$tmpl*") { $esLegacy = $true; break }
                }
            }
            if (-not $esLegacy -and $config.PrefijosLegacy -and $config.PrefijosLegacy.Count -gt 0) {
                foreach ($pfx in $config.PrefijosLegacy) {
                    if ($fila.Nombre -like "$pfx*") { $esLegacy = $true; break }
                }
            }
            if (-not $esLegacy) { continue }
        }

        $sitios.Add([PSCustomObject]@{
            Nombre      = $fila.Nombre
            URL         = $fila.URL
            Tipo        = if ($fila.Tipo_Sitio -eq "Canal de Teams") { "CANAL" } else { "TeamSite" }
            Storage_GB  = $fila.Storage_GB
            Bloqueado   = $fila.Bloqueado
            _StorageNum = ConvertTo-Gb $fila.Storage_GB
        })
    }
}

if ($sitios.Count -eq 0) {
    Write-Log "No hay sitios para procesar." -Level WARN
    exit 0
}

# Ordenar por storage descendente (mismo criterio que AsignarOwner)
$sitiosOrdenados = @($sitios | Sort-Object { $_._StorageNum } -Descending)

# Conteo previo de lo que se va a tocar (para el resumen de confirmación)
$aProcesar = 0
foreach ($s in $sitiosOrdenados) {
    $m = Get-PatronMatch -Url $s.URL -Patrones $patronesCsv
    $salta = if ($InvertirFiltro) { -not $m } else { [bool]$m }
    if (-not $salta) { $aProcesar++ }
}

Write-Log "" -Level INFO
Write-Log "Sitios en la lista  : $($sitiosOrdenados.Count)" -Level OK
Write-Log "Sitios a procesar   : $aProcesar" -Level OK
Write-Log "Filtrados por URL   : $($sitiosOrdenados.Count - $aProcesar)" -Level INFO

if ($aProcesar -eq 0) {
    Write-Log "El filtro dejó afuera todos los sitios. Nada que hacer." -Level WARN
    exit 0
}

# ============================================================
# CONFIRMACIÓN
# ============================================================
if (-not $Simular -and -not $SinConfirmacion) {
    Write-Log "" -Level INFO
    Write-Log "Se va a REMOVER $($config.UPNCirco) como Site Collection Admin de $aProcesar sitio(s)." -Level WARN
    Write-Log "Esta acción quita el acceso de Circo Studio a esos sitios." -Level WARN
    $rta = Read-Host "Escribir SI para continuar (cualquier otra cosa cancela)"
    if ($rta.Trim().ToUpper() -ne "SI") {
        Write-Log "Cancelado por el usuario." -Level WARN
        exit 0
    }
}

# ============================================================
# CONECTAR AL TENANT (una sola vez — token cacheado)
# ============================================================
Write-Log "" -Level INFO
Write-Log "Conectando al tenant..." -Level INFO

try {
    $connectParams = @{
        Url         = $config.AdminUrl
        ClientId    = $config.ClientId
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
$total      = $sitiosOrdenados.Count
$exitosos   = 0
$fallidos   = 0
$omitidos   = 0
$noEraAdmin = 0
$simulados  = 0
$orden      = 0

foreach ($sitio in $sitiosOrdenados) {
    $orden++

    Write-Log "" -Level INFO
    Write-Log "[$orden/$total] $($sitio.Nombre)" -Level INFO
    Write-Log "  Tipo    : $($sitio.Tipo)  |  Storage: $($sitio.Storage_GB) GB" -Level INFO
    Write-Log "  URL     : $($sitio.URL)" -Level INFO

    # --- Filtro por CSV de URLs ---
    $patron = Get-PatronMatch -Url $sitio.URL -Patrones $patronesCsv
    $saltar = if ($InvertirFiltro) { -not $patron } else { [bool]$patron }

    if ($saltar) {
        $detalle = if ($InvertirFiltro) { "No está en el CSV de inclusión" }
                   else { "Excluido por CSV — patrón: $patron" }
        Write-Log "  [OMITIDO] $detalle" -Level WARN
        Add-FilaLog -Orden $orden -Nombre $sitio.Nombre -URL $sitio.URL -Tipo $sitio.Tipo `
                    -Storage $sitio.Storage_GB -Resultado "OMITIDO" -Detalle $detalle
        $omitidos++
        continue
    }

    # --- Sitios bloqueados (ReadOnly) ---
    if ($sitio.Bloqueado -eq "ReadOnly") {
        Write-Log "  [OMITIDO] Sitio en estado ReadOnly" -Level WARN
        Add-FilaLog -Orden $orden -Nombre $sitio.Nombre -URL $sitio.URL -Tipo $sitio.Tipo `
                    -Storage $sitio.Storage_GB -Resultado "OMITIDO" -Detalle "Sitio bloqueado (ReadOnly)"
        $omitidos++
        continue
    }

    try {
        # Conectar al sitio (reutiliza token cacheado — sin -Interactive)
        $connSitio = Connect-PnPOnline -Url $sitio.URL -ClientId $config.ClientId `
                                       -PersistLogin -ReturnConnection -ErrorAction Stop

        # Verificar si la cuenta realmente es Site Collection Admin
        $admins = @(Get-PnPSiteCollectionAdmin -Connection $connSitio -ErrorAction Stop)
        $esAdmin = @($admins | Where-Object {
            $_.Email -eq $config.UPNCirco -or
            $_.LoginName -like "*$($config.UPNCirco)*" -or
            $_.UserPrincipalName -eq $config.UPNCirco
        }).Count -gt 0

        if (-not $esAdmin) {
            Write-Log "  [SKIP] La cuenta no figura como Site Collection Admin" -Level WARN
            Add-FilaLog -Orden $orden -Nombre $sitio.Nombre -URL $sitio.URL -Tipo $sitio.Tipo `
                        -Storage $sitio.Storage_GB -Resultado "NO_ERA_ADMIN" -Detalle "Nada que remover"
            $noEraAdmin++
        }
        elseif ($Simular) {
            Write-Log "  [SIMULADO] Se removería el owner (admins actuales: $($admins.Count))" -Level WARN
            Add-FilaLog -Orden $orden -Nombre $sitio.Nombre -URL $sitio.URL -Tipo $sitio.Tipo `
                        -Storage $sitio.Storage_GB -Resultado "SIMULADO" -Detalle "Owner presente — se removería"
            $simulados++
        }
        else {
            # Inverso de Set-PnPTenantSite -Owners: quita el rol sin tocar a los demás admins
            Remove-PnPSiteCollectionAdmin -Owners $config.UPNCirco `
                                          -Connection $connSitio -ErrorAction Stop
            Write-Log "  [OK] Owner removido" -Level OK
            Add-FilaLog -Orden $orden -Nombre $sitio.Nombre -URL $sitio.URL -Tipo $sitio.Tipo `
                        -Storage $sitio.Storage_GB -Resultado "OK" -Detalle "Owner removido"
            $exitosos++
        }
    }
    catch {
        $errorMsg = $_.Exception.Message -replace '"', "'" -replace "`r`n|`n", " "
        Write-Log "  [ERROR] $errorMsg" -Level ERROR
        Add-FilaLog -Orden $orden -Nombre $sitio.Nombre -URL $sitio.URL -Tipo $sitio.Tipo `
                    -Storage $sitio.Storage_GB -Resultado "ERROR" -Detalle $errorMsg
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
Write-Log "  Modo               : $modoTexto"                      -Level INFO
Write-Log "  Total en la lista  : $total"                           -Level INFO
Write-Log "  Removidos OK       : $exitosos"                        -Level OK
Write-Log "  Simulados          : $simulados"                       -Level INFO
Write-Log "  No era admin       : $noEraAdmin"                      -Level INFO
Write-Log "  Omitidos por filtro: $omitidos"                        -Level INFO
Write-Log "  Con error          : $fallidos"                        -Level $(if ($fallidos -gt 0) {"ERROR"} else {"OK"})
Write-Log "  Log guardado en    : $logCsv"                          -Level INFO
Write-Log "======================================================" -Level HEADER

if ($fallidos -gt 0) {
    Write-Log "Sitios con error — la cuenta puede seguir teniendo acceso:" -Level WARN
    Import-Csv -Path $logCsv | Where-Object { $_.Resultado -eq "ERROR" } | ForEach-Object {
        Write-Log "  [$($_.Orden)] $($_.Nombre)" -Level ERROR
        Write-Log "       $($_.Detalle)"          -Level ERROR
    }
}

Write-Log "Script A RemoverOwner finalizado." -Level OK
