# =============================================================================
# M365Curaduria.psm1  |  Circo Studio  |  Kit de Curaduría Informacional M365
# =============================================================================
# Módulo de funciones compartidas para todos los scripts del kit.
# Importar desde cada script con:
#   Import-Module "$PSScriptRoot\M365Curaduria.psm1" -Force
#
# Funciones exportadas:
#   Read-ClientConfig       — carga y valida cliente_config.json
#   Write-Log               — logging con niveles y colores
#   Connect-TenantOnce      — login único con token cacheado
#   Get-FileCategory        — clasifica archivos por extensión
#   Get-SitiosFiltrados     -- filtra sitios Legacy (Teams) segun config
#   Invoke-RelevamientoSitio — relevamiento completo de un sitio
# =============================================================================


# =============================================================================
# Read-ClientConfig
# Carga el JSON de configuración del cliente y valida los campos obligatorios.
# Devuelve un PSCustomObject con toda la config o lanza un error descriptivo.
# =============================================================================
function Read-ClientConfig {
    param(
        [Parameter(Mandatory)]
        [string]$ConfigPath
    )

    # Verificar que el archivo existe
    if (-not (Test-Path $ConfigPath)) {
        throw "No se encontró el archivo de configuración: $ConfigPath`n" +
              "Generarlo con el editor HTML del kit: Config/editor_config.html"
    }

    # Leer y parsear JSON
    $raw = $null
    try {
        $raw = Get-Content -Path $ConfigPath -Encoding UTF8 -Raw | ConvertFrom-Json
    }
    catch {
        throw "Error al parsear el JSON de configuración: $_`n" +
              "Verificar que el archivo sea JSON válido."
    }

    # Validar campos obligatorios
    $obligatorios = @(
        "cliente.tenant_prefix",
        "cliente.sharepoint_url",
        "cliente.admin_url",
        "cliente.client_id",
        "cliente.upn_circo",
        "cliente.carpeta_local"
    )

    foreach ($campo in $obligatorios) {
        $partes = $campo -split "\."
        $val = $raw
        foreach ($p in $partes) { $val = $val.$p }
        if (-not $val) {
            throw "Campo obligatorio faltante en la configuración: $campo`n" +
                  "Completar en el editor HTML antes de ejecutar."
        }
    }

    # Construir objeto con valores y defaults
    $config = [PSCustomObject]@{

        # Identidad del cliente
        NombreCliente   = $raw.cliente.nombre
        TenantPrefix    = $raw.cliente.tenant_prefix
        SharePointUrl   = $raw.cliente.sharepoint_url
        AdminUrl        = $raw.cliente.admin_url
        ClientId        = $raw.cliente.client_id
        UPNCirco        = $raw.cliente.upn_circo
        CarpetaLocal    = $raw.cliente.carpeta_local

        # Fechas de referencia
        FechaMigracion  = [datetime]::Parse($raw.fechas.migracion)
        InactividadAnios = if ($raw.fechas.inactividad_anios) { [int]$raw.fechas.inactividad_anios } else { 2 }

        # Parámetros de refresh
        DiasUmbral      = if ($raw.refresh.dias_umbral) { [int]$raw.refresh.dias_umbral } else { 7 }

        # Clasificación de sitios
        GBPrincipal     = if ($raw.clasificacion.gb_principal) { [double]$raw.clasificacion.gb_principal } else { 10.0 }
        GBMediano       = if ($raw.clasificacion.gb_mediano)   { [double]$raw.clasificacion.gb_mediano }   else { 1.0 }
        MBArchivoGrande = if ($raw.clasificacion.mb_archivo_grande) { [int]$raw.clasificacion.mb_archivo_grande } else { 100 }

        # Listas de exclusión
        SitiosExcluidos   = if ($raw.sitios_excluidos)  { @($raw.sitios_excluidos)  } else { @() }
        CanalesExcluidos  = if ($raw.canales_excluidos) { @($raw.canales_excluidos) } else { @() }

        # Informes prioritarios
        InformesPrioritarios = if ($raw.informes_prioritarios) { @($raw.informes_prioritarios) } else { @() }

        # Sitios legacy (Teams) -- identificacion por template de SharePoint, no por prefijo de nombre
        # TemplatesLegacy: los templates que genera Teams automaticamente -- universales, sin depender del naming del cliente
        # PrefijosLegacy:  prefijos de nombre opcionales del cliente (ej: "GR-", "EQ-") -- solo para label visual en logs
        # LabelLegacy:     texto que aparece en logs y etiquetas -- default generico si no hay prefijo configurado
        TemplatesLegacy  = if ($raw.sitios_legacy -and $raw.sitios_legacy.templates) {
                               @($raw.sitios_legacy.templates)
                           } else {
                               @("GROUP#0","TEAMCHANNEL#0","GROUP#1","TEAMCHANNEL#1")
                           }
        PrefijosLegacy   = if ($raw.sitios_legacy -and $raw.sitios_legacy.prefijos) { @($raw.sitios_legacy.prefijos) } else { @() }
        LabelLegacy      = if ($raw.sitios_legacy -and $raw.sitios_legacy.label)    { $raw.sitios_legacy.label }     else { "Legacy (Teams)" }
    }

    return $config
}


# =============================================================================
# Write-Log
# Escribe una línea de log en consola (con color) y en un archivo de texto.
# Niveles: INFO | OK | WARN | ERROR | HEADER
# =============================================================================
function Write-Log {
    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [ValidateSet("INFO","OK","WARN","ERROR","HEADER")]
        [string]$Level = "INFO",

        # Path al archivo de log — si no se pasa, solo escribe en consola
        [string]$LogFile = ""
    )

    $timestamp = Get-Date -Format "HH:mm:ss"
    $entrada   = "[$timestamp] [$Level] $Message"

    # Escribir en archivo si se proporcionó
    if ($LogFile) {
        Add-Content -Path $LogFile -Value $entrada -Encoding UTF8
    }

    # Escribir en consola con color según nivel
    switch ($Level) {
        "INFO"   { Write-Host $entrada -ForegroundColor Gray    }
        "OK"     { Write-Host $entrada -ForegroundColor Green   }
        "WARN"   { Write-Host $entrada -ForegroundColor Yellow  }
        "ERROR"  { Write-Host $entrada -ForegroundColor Red     }
        "HEADER" { Write-Host $entrada -ForegroundColor Cyan    }
    }
}


# =============================================================================
# Connect-TenantOnce
# Realiza el login interactivo una sola vez y cachea el token con PersistLogin.
# Los scripts llaman a esta función al inicio; las reconexiones por sitio
# usan Connect-PnPOnline sin -Interactive (reutiliza el token cacheado).
# =============================================================================
function Connect-TenantOnce {
    param(
        [Parameter(Mandatory)]
        [PSCustomObject]$Config,

        [string]$LogFile = "",

        # Si se pasa, fuerza re-autenticación aunque haya token cacheado
        [switch]$ForceAuth
    )

    Write-Log "Iniciando sesión en el tenant (única vez — completar la ventana de login)..." `
              -Level WARN -LogFile $LogFile

    $connectParams = @{
        Url      = $Config.SharePointUrl
        ClientId = $Config.ClientId
        Interactive = $true
    }

    if ($ForceAuth) {
        $connectParams["ForceAuthentication"] = $true
    } else {
        $connectParams["PersistLogin"] = $true
    }

    try {
        Connect-PnPOnline @connectParams -ErrorAction Stop
        Write-Log "Sesión iniciada. Token cacheado para todos los sitios." -Level OK -LogFile $LogFile
    }
    catch {
        Write-Log "Error al iniciar sesión: $_" -Level ERROR -LogFile $LogFile
        throw
    }
}


# =============================================================================
# Get-FileCategory
# Clasifica un archivo en una categoría legible según su extensión.
# =============================================================================
function Get-FileCategory {
    param(
        [string]$Extension
    )

    switch ($Extension.ToLower()) {
        { $_ -in @("pdf") }                                       { return "PDF" }
        { $_ -in @("doc","docx","odt") }                          { return "Word" }
        { $_ -in @("xls","xlsx","xlsm","xlsb","csv","ods") }      { return "Excel / Datos" }
        { $_ -in @("ppt","pptx","odp") }                          { return "PowerPoint" }
        { $_ -in @("jpg","jpeg","png","gif","bmp","tif","tiff","heic","webp") } { return "Imagen" }
        { $_ -in @("mp4","avi","mov","wmv","mkv","m4v") }         { return "Video" }
        { $_ -in @("zip","rar","7z","tar","gz","bz2") }           { return "Comprimido" }
        { $_ -in @("dwg","dxf","dgn","nwd","rvt","ifc") }         { return "CAD / Planos" }
        { $_ -in @("msg","eml") }                                 { return "Email" }
        { $_ -in @("mpp","mpt") }                                 { return "Proyecto" }
        { $_ -in @("txt","log","xml","json","html","htm","md") }   { return "Texto / Código" }
        { $_ -in @("xer","p6xml") }                               { return "Primavera" }
        default                                                    { return "Otro" }
    }
}


# =============================================================================
# Get-SitiosFiltrados
# Aplica los filtros de exclusion del config sobre un CSV de inventario
# y devuelve la lista de sitios a procesar, ordenados por storage descendente.
#
# Criterios (todos configurables desde cliente_config.json):
#   - Solo sitios cuyo Template esta en config.TemplatesLegacy (sitios de Teams)
#     O cuyo nombre empieza con alguno de config.PrefijosLegacy si estan configurados
#   - Excluye ReadOnly
#   - Excluye IDs en config.SitiosExcluidos
#   - Excluye canales cuyos prefijos estan en config.CanalesExcluidos
#   - Solo sitios con Storage > 0 GB
#
# NOTA: el filtro es por template de SharePoint (universal) -- no por prefijo de nombre.
# El prefijo de nombre (ej: "GR-", "EQ-") es solo una convencion del cliente
# y se usa unicamente como label visual, no como criterio de filtrado.
# =============================================================================
function Get-SitiosFiltrados {
    param(
        [Parameter(Mandatory)]
        [string]$CsvInventario,

        [Parameter(Mandatory)]
        [PSCustomObject]$Config,

        [string]$LogFile = ""
    )

    if (-not (Test-Path $CsvInventario)) {
        throw "No se encontró el CSV de inventario: $CsvInventario"
    }

    $inventario = Import-Csv -Path $CsvInventario -Encoding UTF8
    Write-Log "CSV de inventario cargado: $($inventario.Count) filas" -Level INFO -LogFile $LogFile

    # Extraer IDs de los sitios ya analizados
    $idsExcluidos = @($Config.SitiosExcluidos | ForEach-Object { $_.id })

    $pendientes = [System.Collections.Generic.List[object]]::new()

    foreach ($fila in $inventario) {

        # Solo sitios Legacy (Teams) -- filtrar por template de SharePoint
        # El template es universal y no depende del naming convention del cliente
        # Si el cliente tiene PrefijosLegacy configurados, se usa como filtro adicional opcional
        $esLegacy = $false

        # Criterio principal: template de SharePoint
        if ($Config.TemplatesLegacy -and $Config.TemplatesLegacy.Count -gt 0) {
            foreach ($tmpl in $Config.TemplatesLegacy) {
                if ($fila.Template -like "$tmpl*") { $esLegacy = $true; break }
            }
        }

        # Criterio alternativo: prefijo de nombre (si el cliente lo tiene configurado y el template no matcheo)
        if (-not $esLegacy -and $Config.PrefijosLegacy -and $Config.PrefijosLegacy.Count -gt 0) {
            foreach ($prefijo in $Config.PrefijosLegacy) {
                if ($fila.Nombre -like "$prefijo*") { $esLegacy = $true; break }
            }
        }

        if (-not $esLegacy) { continue }

        # Excluir bloqueados (ReadOnly)
        if ($fila.Bloqueado -eq "ReadOnly") { continue }

        # Excluir sitios ya analizados (por ID en la URL)
        $esExcluido = $false
        foreach ($id in $idsExcluidos) {
            if ($fila.URL -like "*$id*") { $esExcluido = $true; break }
        }
        if ($esExcluido) { continue }

        # Excluir canales de padres ya analizados (por prefijo de nombre)
        if ($fila.Tipo_Sitio -eq "Canal de Teams") {
            $canalExcluido = $false
            foreach ($prefijo in $Config.CanalesExcluidos) {
                if ($fila.Nombre -like "$prefijo*") { $canalExcluido = $true; break }
            }
            if ($canalExcluido) { continue }
        }

        # Solo sitios con storage real
        $gb = [double]($fila.Storage_GB -replace ',', '.')
        if ($gb -le 0) { continue }

        $fila | Add-Member -NotePropertyName "_StorageNum" -NotePropertyValue $gb -Force
        $pendientes.Add($fila)
    }

    # Ordenar por storage descendente (los más grandes primero)
    $resultado = $pendientes | Sort-Object { $_._StorageNum } -Descending

    Write-Log "Sitios a procesar: $(@($resultado).Count)" -Level OK -LogFile $LogFile

    $storageTotal = ($resultado | Measure-Object -Property _StorageNum -Sum).Sum
    Write-Log "Storage total estimado: $([math]::Round($storageTotal, 1)) GB" -Level INFO -LogFile $LogFile

    return $resultado
}


# =============================================================================
# Invoke-RelevamientoSitio
# Conecta a un sitio SharePoint, extrae el inventario completo de archivos
# y los escribe en un CSV por lotes (sin acumular en memoria).
#
# Devuelve un hashtable con:
#   Estado      — OK | SinArchivos | SinBibliotecas | ErrorConexion | ErrorBibliotecas
#   CsvPath     — path al CSV generado (vacío si no hubo archivos)
#   Archivos    — cantidad de archivos procesados
#   Errores     — cantidad de ítems con error
#   Bibliotecas — cantidad de bibliotecas encontradas
# =============================================================================
function Invoke-RelevamientoSitio {
    param(
        [Parameter(Mandatory)]
        [string]$SiteUrl,

        [Parameter(Mandatory)]
        [string]$SiteName,

        [Parameter(Mandatory)]
        [string]$CarpetaSalida,

        [Parameter(Mandatory)]
        [PSCustomObject]$Config,

        [string]$LogFile     = "",
        [string]$Timestamp   = "",

        # Si se pasa, usa este path en lugar de construir uno automático
        [string]$CsvPathOverride = ""
    )

    # Construir nombre limpio para el archivo CSV
    $nombreLimpio = $SiteName -replace "[^a-zA-Z0-9_-]", "_" -replace "__+", "_"
    $nombreLimpio = $nombreLimpio.Substring(0, [Math]::Min($nombreLimpio.Length, 50))

    $ts       = if ($Timestamp) { $Timestamp } else { Get-Date -Format "yyyyMMdd_HHmmss" }
    $prefijo  = if ($Config.NombreCliente) { ($Config.NombreCliente -replace "[^a-zA-Z0-9]","_") + "_" } else { "" }
    $csvSitio = if ($CsvPathOverride) { $CsvPathOverride } else {
        "$CarpetaSalida\${prefijo}B_Docs_${nombreLimpio}_${ts}.csv"
    }

    # Checkpoint: si el CSV ya existe para este sitio, saltear
    $existente = Get-ChildItem $CarpetaSalida -Filter "*B_Docs_${nombreLimpio}_*.csv" -ErrorAction SilentlyContinue |
                 Select-Object -First 1
    if ($existente) {
        Write-Log "  SALTEADO — CSV previo existe: $($existente.Name)" -Level WARN -LogFile $LogFile
        return @{ Estado = "Salteado"; CsvPath = $existente.FullName; Archivos = 0; Errores = 0; Bibliotecas = 0 }
    }

    # Conectar al sitio (reutiliza token cacheado — sin -Interactive)
    try {
        Connect-PnPOnline -Url $SiteUrl -ClientId $Config.ClientId -PersistLogin -ErrorAction Stop
        Write-Log "  Conexión OK: $SiteName" -Level OK -LogFile $LogFile
    }
    catch {
        Write-Log "  Error de conexión en $SiteName : $_" -Level ERROR -LogFile $LogFile
        return @{ Estado = "ErrorConexion"; CsvPath = ""; Archivos = 0; Errores = 1; Bibliotecas = 0 }
    }

    # Obtener bibliotecas de documentos (excluir sistema)
    $bibliotecasSistema = @(
        "Form Templates", "Site Assets", "Site Pages",
        "Style Library", "Preservation Hold Library",
        "Páginas del sitio", "Recursos del sitio", "Activos del sitio"
    )

    $bibliotecas = $null
    try {
        $bibliotecas = Get-PnPList | Where-Object {
            $_.BaseType -eq "DocumentLibrary" -and
            $_.Hidden   -eq $false            -and
            $_.Title    -notin $bibliotecasSistema
        }
    }
    catch {
        Write-Log "  Error al obtener bibliotecas en $SiteName : $_" -Level ERROR -LogFile $LogFile
        Disconnect-PnPOnline
        return @{ Estado = "ErrorBibliotecas"; CsvPath = ""; Archivos = 0; Errores = 1; Bibliotecas = 0 }
    }

    if (-not $bibliotecas -or @($bibliotecas).Count -eq 0) {
        Write-Log "  Sin bibliotecas de documentos en $SiteName" -Level WARN -LogFile $LogFile
        Disconnect-PnPOnline
        return @{ Estado = "SinBibliotecas"; CsvPath = ""; Archivos = 0; Errores = 0; Bibliotecas = 0 }
    }

    $cantBibliotecas = @($bibliotecas).Count
    Write-Log "  Bibliotecas encontradas: $cantBibliotecas" -Level INFO -LogFile $LogFile
    $bibliotecas | ForEach-Object {
        Write-Log "    → $($_.Title) ($($_.ItemCount) ítems)" -Level INFO -LogFile $LogFile
    }

    # Campos a recuperar de cada ítem
    $fields = @(
        "FileRef", "FileLeafRef", "File_x0020_Type", "File_x0020_Size",
        "Created", "Modified", "Author", "Editor",
        "_UIVersionString", "CheckoutUser"
    )

    # Variables en scope script para que el ScriptBlock pueda accederlas
    $script:csvSitioRef   = $csvSitio
    $script:siteNameRef   = $SiteName
    $script:siteUrlRef    = $SiteUrl
    $script:sitePath      = ([System.Uri]$SiteUrl).AbsolutePath.TrimEnd("/")
    $script:totalArchivos = 0
    $script:totalErrores  = 0
    $script:primerLote    = $true
    $script:bibActual     = ""
    $script:logFileRef    = $LogFile

    # Iterar cada biblioteca con ScriptBlock — escribe por lotes sin acumular en memoria
    foreach ($bib in $bibliotecas) {

        $script:bibActual = $bib.Title
        Write-Log "  Extrayendo: $($bib.Title) ($($bib.ItemCount) ítems)..." -Level INFO -LogFile $LogFile

        try {
            Get-PnPListItem -List $bib.Title -PageSize 100 -Fields $fields -ScriptBlock {
                Param($items)

                $lote = [System.Collections.Generic.List[PSObject]]::new()

                foreach ($item in $items) {
                    if ($item.FileSystemObjectType -eq "File") {
                        $script:totalArchivos++
                        try {
                            $filePath   = $item["FileRef"]
                            $fileName   = $item["FileLeafRef"]
                            $extension  = if ($item["File_x0020_Type"]) { $item["File_x0020_Type"].ToLower() } else { "" }
                            $sizeBytes  = [long]($item["File_x0020_Size"])
                            $created    = $item["Created"]
                            $modified   = $item["Modified"]
                            $createdBy  = $item["Author"].LookupValue
                            $modifiedBy = $item["Editor"].LookupValue
                            $version    = $item["_UIVersionString"]
                            $checkedOut = if ($item["CheckoutUser"]) { "Sí" } else { "No" }

                            # Calcular ruta relativa y profundidad
                            $relPath   = $filePath -replace "^$([regex]::Escape($script:sitePath))/", ""
                            $folderRel = $relPath  -replace "/$([regex]::Escape($fileName))$", ""
                            $depth     = ($relPath -split "/").Count - 1

                            $row = [PSCustomObject]@{
                                "Sitio"               = $script:siteNameRef
                                "URL_Sitio"           = $script:siteUrlRef
                                "Biblioteca"          = $script:bibActual
                                "Ruta_Completa"       = $filePath
                                "Carpeta"             = $folderRel
                                "Nombre_Archivo"      = $fileName
                                "Extension"           = $extension
                                "Categoria"           = Get-FileCategory $extension
                                "Tamaño_Bytes"        = $sizeBytes
                                "Tamaño_KB"           = [math]::Round($sizeBytes / 1KB, 1)
                                "Tamaño_MB"           = [math]::Round($sizeBytes / 1MB, 2)
                                "Fecha_Creacion"      = if ($created)  { $created.ToString("yyyy-MM-dd HH:mm") }  else { "" }
                                "Fecha_Modificacion"  = if ($modified) { $modified.ToString("yyyy-MM-dd HH:mm") } else { "" }
                                "Creado_Por"          = $createdBy
                                "Modificado_Por"      = $modifiedBy
                                "Version"             = $version
                                "Archivo_Bloqueado"   = $checkedOut
                                "Profundidad_Carpeta" = $depth
                                "Destino_Migracion"   = ""
                                "Accion"              = ""
                                "Notas"               = ""
                            }
                            $lote.Add($row)

                        }
                        catch {
                            $script:totalErrores++
                        }
                    }
                }

                # Escribir lote al CSV inmediatamente — sin acumular en memoria
                if ($lote.Count -gt 0) {
                    if ($script:primerLote) {
                        $lote | Export-Csv -Path $script:csvSitioRef -NoTypeInformation -Encoding UTF8
                        $script:primerLote = $false
                    }
                    else {
                        $lote | Export-Csv -Path $script:csvSitioRef -NoTypeInformation -Encoding UTF8 -Append
                    }
                }

                if ($script:totalArchivos % 500 -eq 0 -and $script:totalArchivos -gt 0) {
                    Write-Log "    Archivos procesados: $($script:totalArchivos)..." -Level INFO -LogFile $script:logFileRef
                }

            } | Out-Null

            Write-Log "    $($bib.Title) — OK" -Level OK -LogFile $LogFile

        }
        catch {
            Write-Log "    Error en biblioteca $($bib.Title): $_" -Level ERROR -LogFile $LogFile
            $script:totalErrores++
        }
    }

    Disconnect-PnPOnline

    $estado = if ($script:totalArchivos -gt 0) { "OK" } else { "SinArchivos" }

    Write-Log "  $SiteName completado — $($script:totalArchivos) archivos | $($script:totalErrores) errores" `
              -Level OK -LogFile $LogFile

    if ($script:totalArchivos -gt 0) {
        Write-Log "  CSV: $csvSitio" -Level OK -LogFile $LogFile
    }

    return @{
        Estado      = $estado
        CsvPath     = if ($script:totalArchivos -gt 0) { $csvSitio } else { "" }
        Archivos    = $script:totalArchivos
        Errores     = $script:totalErrores
        Bibliotecas = $cantBibliotecas
    }
}


# =============================================================================
# Funciones exportadas del módulo
# =============================================================================
Export-ModuleMember -Function @(
    "Read-ClientConfig",
    "Write-Log",
    "Connect-TenantOnce",
    "Get-FileCategory",
    "Get-SitiosFiltrados",
    "Invoke-RelevamientoSitio"
)
