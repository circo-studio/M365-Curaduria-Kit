#Requires -Modules PnP.PowerShell
<#
===============================================================================
  Script_C_DescargarBiblioteca.ps1  |  Circo Studio  |  Kit de Curaduría M365
===============================================================================

PROPÓSITO:
  Descarga el contenido completo de una biblioteca de documentos y lo comprime
  en ZIP, preservando la estructura de carpetas. Pensado para bibliotecas
  grandes (decenas de GB y decenas de miles de archivos).

DISEÑO PARA VOLÚMENES GRANDES:
  - Los archivos se bajan de a uno a un temporal y se agregan al ZIP con
    System.IO.Compression. Nunca se carga el contenido en memoria, así que
    el consumo de RAM no depende del tamaño de la biblioteca.
  - El ZIP se parte en volúmenes (default 4 GB). Cada volumen se cierra y se
    valida apenas se completa, así un corte no invalida lo ya descargado y
    los archivos resultantes son transferibles.
  - Archivo de estado: cada archivo descargado se registra apenas termina.
    Con -Reanudar el script saltea lo ya bajado y sigue donde quedó.
  - Los formatos ya comprimidos (pdf, docx, jpg, zip, mp4...) se guardan sin
    recomprimir: mismo tamaño final y mucho menos CPU.

USO:
  # Interactivo: pide URL y muestra las bibliotecas del sitio
  .\Script_C_DescargarBiblioteca.ps1 -ConfigPath "..\..\Config\cliente_config.json"

  # Directo
  .\Script_C_DescargarBiblioteca.ps1 -ConfigPath ".\cliente_config.json" `
      -Url "https://tenant.sharepoint.com/sites/EJEMPLO" -Biblioteca "Documents"

  # Retomar una descarga interrumpida
  .\Script_C_DescargarBiblioteca.ps1 -ConfigPath ".\cliente_config.json" `
      -Url "https://..." -Biblioteca "Documents" -Reanudar

  # Solo una carpeta, solo PDFs, volúmenes de 2 GB
  .\Script_C_DescargarBiblioteca.ps1 -ConfigPath ".\cliente_config.json" `
      -Url "https://..." -Biblioteca "Documents" `
      -SoloCarpeta "Contratos/2024" -Extensiones pdf -TamanoVolumenGB 2

  # Ver qué bajaría, sin bajar nada
  .\Script_C_DescargarBiblioteca.ps1 -ConfigPath ".\cliente_config.json" -Url "..." -Simular

ARCHIVOS GENERADOS (en [Destino]\[Sitio]_[Biblioteca]\):
  [Sitio]_[Biblioteca]_vol01.zip, _vol02.zip, ...   — los volúmenes
  _LOG.csv        — una fila por archivo: OK | ERROR | OMITIDO | SALTADO
  _FALTANTES.csv  — solo los que no se pudieron bajar (vacío si salió todo)
  _ESTADO.csv     — control interno de reanudación (no borrar hasta terminar)

===============================================================================
#>

param(
    [Parameter(Mandatory, HelpMessage="Path al archivo cliente_config.json")]
    [string]$ConfigPath,

    # URL del sitio. Si se omite, se pide por consola
    [string]$Url,

    # Nombre de la biblioteca. Si se omite, se listan las del sitio
    [string]$Biblioteca,

    # Carpeta donde se arman los ZIP. Default: [CarpetaLocal]\Descargas
    [string]$CarpetaDestino,

    # Tamaño máximo de cada volumen ZIP
    [ValidateRange(0.1, 50)]
    [double]$TamanoVolumenGB = 4,

    # Solo una subcarpeta de la biblioteca (ruta relativa, ej: "Contratos/2024")
    [string]$SoloCarpeta,

    # Extensiones a incluir (sin punto). Vacío = todas
    [string[]]$Extensiones = @(),

    # Excluye archivos por encima de este tamaño
    [int]$TamanoMaxMB = 0,

    # Solo archivos modificados desde esta fecha
    [datetime]$ModificadoDesde,

    # Retoma una descarga previa salteando lo ya bajado
    [switch]$Reanudar,

    # Deja los archivos sueltos en carpetas, sin comprimir
    [switch]$SinComprimir,

    # Solo enumera y reporta, no descarga
    [switch]$Simular,

    # Pausa entre archivos, en milisegundos (anti-throttling)
    [int]$PausaMs = 50,

    # Reintentos por archivo
    [int]$MaxReintentos = 3,

    [switch]$SinConfirmacion,
    [switch]$ForceAuth
)

$moduloPath = Join-Path $PSScriptRoot "M365Curaduria.psm1"
if (-not (Test-Path $moduloPath)) {
    Write-Host "[ERROR] No se encontró M365Curaduria.psm1 en: $moduloPath" -ForegroundColor Red
    exit 1
}
Import-Module $moduloPath -Force

Add-Type -AssemblyName System.IO.Compression | Out-Null
Add-Type -AssemblyName System.IO.Compression.FileSystem | Out-Null

# ============================================================
# HELPERS
# ============================================================
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

# Formatos que ya vienen comprimidos: recomprimirlos gasta CPU sin ganar espacio
$YA_COMPRIMIDOS = @(
    "zip","7z","rar","gz","bz2","xz","cab",
    "jpg","jpeg","png","gif","webp","heic",
    "mp3","mp4","avi","mkv","mov","wmv","m4a","aac",
    "pdf","docx","xlsx","pptx","docm","xlsm","pptm","odt","ods","odp"
)

function Get-NivelCompresion {
    param([string]$NombreArchivo)
    $ext = ([System.IO.Path]::GetExtension($NombreArchivo)).TrimStart('.').ToLower()
    if ($YA_COMPRIMIDOS -contains $ext) {
        return [System.IO.Compression.CompressionLevel]::NoCompression
    }
    return [System.IO.Compression.CompressionLevel]::Optimal
}

function Get-NombreSeguro {
    param([string]$Texto)
    return ($Texto -replace '[^a-zA-Z0-9\-_]', '_')
}

# ============================================================
# CONFIGURACIÓN
# ============================================================
$config = $null
try { $config = Read-ClientConfig -ConfigPath $ConfigPath }
catch { Write-Host "[ERROR] $_" -ForegroundColor Red; exit 1 }

Write-Log "======================================================" -Level HEADER
Write-Log "  Kit M365 Curaduría — Descarga de biblioteca"          -Level HEADER
Write-Log "  Cliente : $($config.NombreCliente)"                   -Level HEADER
Write-Log "======================================================" -Level HEADER

# --- URL del sitio ---
if (-not $Url) {
    Write-Log "" -Level INFO
    $Url = Read-Host "URL del sitio (ej: $($config.SharePointUrl)/sites/EJEMPLO)"
}
$Url = ([string]$Url).Trim().Trim('"').TrimEnd('/')

if (-not $Url -or $Url -notmatch '^https?://') {
    Write-Log "URL inválida. Debe empezar con https://" -Level ERROR
    exit 1
}

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
    Write-Log "Conexión OK: $($web.Title)" -Level OK
}
catch {
    Write-Log "No se pudo conectar: $($_.Exception.Message)" -Level ERROR
    Write-Log "Verificar que la cuenta tenga acceso al sitio." -Level ERROR
    exit 1
}

# ============================================================
# ELECCIÓN DE BIBLIOTECA
# ============================================================
$listas = @()
try {
    $listas = @(Get-PnPList -Connection $conn -ErrorAction Stop |
                Where-Object { $_.BaseType -eq "DocumentLibrary" -and -not $_.Hidden } |
                Sort-Object -Property ItemCount -Descending)
}
catch {
    Write-Log "No se pudieron listar las bibliotecas: $($_.Exception.Message)" -Level ERROR
    exit 1
}

if ($listas.Count -eq 0) {
    Write-Log "El sitio no tiene bibliotecas de documentos visibles." -Level ERROR
    exit 1
}

if (-not $Biblioteca) {
    Write-Log "" -Level INFO
    Write-Host "  Bibliotecas del sitio:" -ForegroundColor Cyan
    for ($i = 0; $i -lt $listas.Count; $i++) {
        Write-Host ("    {0}) {1}  ({2} ítems)" -f ($i + 1), $listas[$i].Title, $listas[$i].ItemCount)
    }
    Write-Host "    0) Cancelar"
    $op = Read-Host "Opción"

    $n = 0
    if (-not [int]::TryParse($op.Trim(), [ref]$n) -or $n -lt 1 -or $n -gt $listas.Count) {
        Write-Log "Cancelado." -Level WARN
        exit 0
    }
    $Biblioteca = $listas[$n - 1].Title
}

$lista = $listas | Where-Object { $_.Title -eq $Biblioteca } | Select-Object -First 1
if (-not $lista) {
    try { $lista = Get-PnPList -Identity $Biblioteca -Connection $conn -ErrorAction Stop }
    catch {
        Write-Log "No se encontró la biblioteca '$Biblioteca' en el sitio." -Level ERROR
        exit 1
    }
}

Write-Log "" -Level INFO
Write-Log "Biblioteca: $($lista.Title) — $($lista.ItemCount) ítems (incluye carpetas)" -Level OK

# ============================================================
# CARPETAS DE TRABAJO
# ============================================================
if (-not $CarpetaDestino) {
    $CarpetaDestino = Join-Path $config.CarpetaLocal "Descargas"
}

$nombreSitio = Get-NombreSeguro (($Url -split '/')[-1])
$nombreBib   = Get-NombreSeguro $lista.Title
$carpetaJob  = Join-Path $CarpetaDestino "${nombreSitio}_${nombreBib}"
$carpetaTmp  = Join-Path $carpetaJob "_tmp"
$carpetaOut  = Join-Path $carpetaJob "Archivos"   # solo con -SinComprimir

New-Item -ItemType Directory -Path $carpetaJob -Force | Out-Null
New-Item -ItemType Directory -Path $carpetaTmp -Force | Out-Null
if ($SinComprimir) { New-Item -ItemType Directory -Path $carpetaOut -Force | Out-Null }

$pathLog       = Join-Path $carpetaJob "_LOG.csv"
$pathFaltantes = Join-Path $carpetaJob "_FALTANTES.csv"
$pathEstado    = Join-Path $carpetaJob "_ESTADO.csv"

Write-Log "Destino   : $carpetaJob" -Level INFO

# ============================================================
# ENUMERACIÓN
# ============================================================
Write-Log "" -Level INFO
Write-Log "Enumerando archivos (puede tardar en bibliotecas grandes)..." -Level INFO

$swEnum = [System.Diagnostics.Stopwatch]::StartNew()
$items = @()

try {
    $items = @(Get-PnPListItem -List $lista.Title -PageSize 1000 `
                               -Fields "FileRef","FileLeafRef","File_x0020_Size","Modified","FSObjType" `
                               -Connection $conn -ErrorAction Stop)
}
catch {
    Write-Log "Error al enumerar: $($_.Exception.Message)" -Level ERROR
    exit 1
}

$swEnum.Stop()
Write-Log "  $($items.Count) ítems leídos en $(Format-Duracion $swEnum.Elapsed.TotalSeconds)" -Level OK

# Ruta raíz de la biblioteca, para calcular rutas relativas
$rootUrl = $lista.RootFolder.ServerRelativeUrl
if (-not $rootUrl) {
    $rootUrl = (Get-PnPList -Identity $lista.Title -Includes RootFolder -Connection $conn).RootFolder.ServerRelativeUrl
}
$rootUrl = $rootUrl.TrimEnd('/')

# --- Filtrado ---
$archivos = [System.Collections.Generic.List[object]]::new()
$omitidosCarpeta = 0; $omitidosExt = 0; $omitidosTam = 0; $omitidosFecha = 0

foreach ($item in $items) {

    if ($item.FileSystemObjectType -ne "File") { continue }

    $fileRef  = [string]$item.FieldValues["FileRef"]
    $fileName = [string]$item.FieldValues["FileLeafRef"]
    if (-not $fileRef) { continue }

    # Ruta relativa dentro de la biblioteca
    $rel = $fileRef
    if ($rel.StartsWith($rootUrl, [System.StringComparison]::OrdinalIgnoreCase)) {
        $rel = $rel.Substring($rootUrl.Length).TrimStart('/')
    }

    if ($SoloCarpeta) {
        $prefijo = $SoloCarpeta.Trim('/').Replace('\', '/')
        if (-not $rel.StartsWith("$prefijo/", [System.StringComparison]::OrdinalIgnoreCase)) {
            $omitidosCarpeta++; continue
        }
    }

    if ($Extensiones.Count -gt 0) {
        $ext = ([System.IO.Path]::GetExtension($fileName)).TrimStart('.').ToLower()
        if ($Extensiones -notcontains $ext) { $omitidosExt++; continue }
    }

    [double]$tam = 0
    $valorTam = $item.FieldValues["File_x0020_Size"]
    if ($valorTam) { try { $tam = [double]$valorTam } catch { $tam = 0 } }

    if ($TamanoMaxMB -gt 0 -and $tam -gt ($TamanoMaxMB * 1MB)) { $omitidosTam++; continue }

    $mod = $item.FieldValues["Modified"]
    if ($ModificadoDesde -and $mod -and ([datetime]$mod) -lt $ModificadoDesde) {
        $omitidosFecha++; continue
    }

    $archivos.Add([PSCustomObject]@{
        FileRef      = $fileRef
        Nombre       = $fileName
        RutaRelativa = $rel
        Bytes        = $tam
        Modificado   = if ($mod) { ([datetime]$mod).ToString("yyyy-MM-dd HH:mm") } else { "" }
    })
}

$totalBytes = ($archivos | Measure-Object -Property Bytes -Sum).Sum
if (-not $totalBytes) { $totalBytes = 0 }

Write-Log "" -Level INFO
Write-Log "======================================================" -Level HEADER
Write-Log "  A DESCARGAR"                                          -Level HEADER
Write-Log "  Archivos : $($archivos.Count)"                        -Level OK
Write-Log "  Tamaño   : $(ConvertTo-Gb $totalBytes) GB"            -Level OK
Write-Log "  Volúmenes: ~$([math]::Max(1, [math]::Ceiling((ConvertTo-Gb $totalBytes) / $TamanoVolumenGB))) de hasta $TamanoVolumenGB GB" -Level INFO
if ($omitidosCarpeta) { Write-Log "  Omitidos por carpeta   : $omitidosCarpeta" -Level INFO }
if ($omitidosExt)     { Write-Log "  Omitidos por extensión : $omitidosExt"     -Level INFO }
if ($omitidosTam)     { Write-Log "  Omitidos por tamaño    : $omitidosTam"     -Level INFO }
if ($omitidosFecha)   { Write-Log "  Omitidos por fecha     : $omitidosFecha"   -Level INFO }
Write-Log "======================================================" -Level HEADER

if ($archivos.Count -eq 0) {
    Write-Log "No hay archivos que cumplan los criterios." -Level WARN
    exit 0
}

# --- Espacio en disco ---
try {
    $unidad = (Get-Item $carpetaJob).PSDrive
    if ($unidad -and $unidad.Free) {
        $libreGb = ConvertTo-Gb ([double]$unidad.Free)
        $necesarioGb = [math]::Round((ConvertTo-Gb $totalBytes) * 1.1, 1)
        Write-Log "Espacio libre en $($unidad.Name): $libreGb GB — necesario aprox: $necesarioGb GB" -Level $(if ($libreGb -lt $necesarioGb) {"ERROR"} else {"INFO"})
        if ($libreGb -lt $necesarioGb) {
            Write-Log "Espacio insuficiente en disco." -Level ERROR
            if (-not $SinConfirmacion) {
                $sigue = Read-Host "Continuar igual? (SI/NO)"
                if ($sigue.Trim().ToUpper() -ne "SI") { exit 1 }
            }
        }
    }
}
catch { }

if ($Simular) {
    Write-Log "" -Level INFO
    Write-Log "Modo -Simular: no se descarga nada." -Level OK
    $archivos | Select-Object RutaRelativa, @{n='MB';e={[math]::Round($_.Bytes/1MB,2)}}, Modificado |
        Export-Csv -Path (Join-Path $carpetaJob "_SIMULACION.csv") -NoTypeInformation -Encoding UTF8
    Write-Log "Listado guardado en: $(Join-Path $carpetaJob '_SIMULACION.csv')" -Level OK
    exit 0
}

# ============================================================
# REANUDACIÓN
# ============================================================
$yaDescargados = New-Object System.Collections.Generic.HashSet[string]

if ($Reanudar -and (Test-Path $pathEstado)) {
    try {
        $previos = @(Import-Csv -Path $pathEstado -Encoding UTF8)
        foreach ($p in $previos) {
            if ($p.FileRef) { [void]$yaDescargados.Add($p.FileRef.ToLower()) }
        }
        Write-Log "" -Level INFO
        Write-Log "Reanudando: $($yaDescargados.Count) archivo(s) ya descargados se saltean" -Level OK
    }
    catch {
        Write-Log "No se pudo leer el estado previo: $($_.Exception.Message)" -Level WARN
    }
}
elseif (-not $Reanudar) {
    # Corrida nueva: se reinician los archivos de control
    "FileRef,Volumen,Timestamp" | Out-File -FilePath $pathEstado -Encoding UTF8
    "Ruta,Nombre,Tamano_MB,Modificado,Resultado,Volumen,Detalle,Timestamp" | Out-File -FilePath $pathLog -Encoding UTF8

    $viejos = Get-ChildItem $carpetaJob -Filter "*_vol*.zip" -ErrorAction SilentlyContinue
    if ($viejos -and -not $SinConfirmacion) {
        Write-Log "" -Level WARN
        Write-Log "Ya existen $($viejos.Count) volumen(es) de una corrida anterior." -Level WARN
        Write-Log "Sin -Reanudar se van a sobrescribir." -Level WARN
        $rta = Read-Host "Continuar y sobrescribir? (SI/NO)"
        if ($rta.Trim().ToUpper() -ne "SI") { Write-Log "Cancelado." -Level WARN; exit 0 }
    }
}

if (-not (Test-Path $pathEstado)) { "FileRef,Volumen,Timestamp" | Out-File -FilePath $pathEstado -Encoding UTF8 }
if (-not (Test-Path $pathLog))    { "Ruta,Nombre,Tamano_MB,Modificado,Resultado,Volumen,Detalle,Timestamp" | Out-File -FilePath $pathLog -Encoding UTF8 }

# ============================================================
# BUFFERS DE LOG
# ============================================================
$script:pathLog    = $pathLog
$script:pathEstado = $pathEstado
$script:bufLog     = [System.Collections.Generic.List[string]]::new()
$script:bufEstado  = [System.Collections.Generic.List[string]]::new()

function Add-Log {
    param([string]$Ruta, [string]$Nombre, [double]$Bytes, [string]$Modificado,
          [string]$Resultado, [string]$Volumen, [string]$Detalle)
    $r = (($Ruta    -replace '"', "'") -replace "`r`n|`n", " ")
    $n = (($Nombre  -replace '"', "'") -replace "`r`n|`n", " ")
    $d = (($Detalle -replace '"', "'") -replace "`r`n|`n", " ")
    $mb = [math]::Round($Bytes / 1MB, 3)
    $script:bufLog.Add("`"$r`",`"$n`",`"$mb`",`"$Modificado`",`"$Resultado`",`"$Volumen`",`"$d`",`"$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`"")
}

function Add-Estado {
    param([string]$FileRef, [string]$Volumen)
    $script:bufEstado.Add("`"$FileRef`",`"$Volumen`",`"$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`"")
}

function Write-Buffer {
    param(
        [Parameter(Mandatory)][System.Collections.Generic.List[string]]$Buffer,
        [Parameter(Mandatory)][string]$Path
    )
    if ($Buffer.Count -eq 0) { return }
    # Reintento: si el CSV está abierto en Excel, el buffer espera al próximo flush
    for ($i = 1; $i -le 3; $i++) {
        try {
            Add-Content -Path $Path -Value $Buffer -Encoding UTF8 -ErrorAction Stop
            $Buffer.Clear()
            return
        }
        catch { Start-Sleep -Milliseconds 300 }
    }
}

function Flush-Buffers {
    Write-Buffer -Buffer $script:bufLog    -Path $script:pathLog
    Write-Buffer -Buffer $script:bufEstado -Path $script:pathEstado
}

# ============================================================
# MANEJO DE VOLÚMENES ZIP
# ============================================================
$script:volNum   = 0
$script:zip      = $null
$script:zipFs    = $null
$script:zipPath  = ""
$limiteBytes     = [double]($TamanoVolumenGB * 1GB)

function Open-Volumen {
    $script:volNum++
    $nombre = "{0}_{1}_vol{2:d2}.zip" -f $nombreSitio, $nombreBib, $script:volNum
    $script:zipPath = Join-Path $carpetaJob $nombre
    $script:zipFs = [System.IO.File]::Open($script:zipPath, [System.IO.FileMode]::Create)
    $script:zip = New-Object System.IO.Compression.ZipArchive(
        $script:zipFs, [System.IO.Compression.ZipArchiveMode]::Create, $true)
    Write-Log "  → Nuevo volumen: $nombre" -Level OK
}

function Close-Volumen {
    if (-not $script:zip) { return }

    $nombre = Split-Path $script:zipPath -Leaf
    $script:zip.Dispose(); $script:zip = $null
    $script:zipFs.Dispose(); $script:zipFs = $null

    # Validación: se reabre y se cuentan las entradas. Un ZIP que no se puede
    # abrir acá tampoco lo va a poder abrir el cliente.
    try {
        $verif = [System.IO.Compression.ZipFile]::OpenRead($script:zipPath)
        $n = $verif.Entries.Count
        $verif.Dispose()
        $gb = ConvertTo-Gb ((Get-Item $script:zipPath).Length)
        Write-Log "  ✓ $nombre cerrado y validado — $n archivos, $gb GB" -Level OK
    }
    catch {
        Write-Log "  [ERROR] $nombre no se pudo validar: $($_.Exception.Message)" -Level ERROR
    }
}

# ============================================================
# DESCARGA
# ============================================================
$total       = $archivos.Count
$ok          = 0
$errores     = 0
$salteados   = 0
$bytesOk     = [double]0
$procesados  = 0
$sw          = [System.Diagnostics.Stopwatch]::StartNew()
$ultimaBarra = [datetime]::Now.AddSeconds(-60)
$ultimoTexto = [datetime]::Now.AddSeconds(-60)

Write-Log "" -Level INFO
Write-Log "Iniciando descarga de $total archivo(s) — $(ConvertTo-Gb $totalBytes) GB" -Level HEADER

if (-not $SinComprimir) { Open-Volumen }

foreach ($arch in $archivos) {

    $procesados++

    # --- Ya descargado en una corrida anterior ---
    if ($yaDescargados.Contains($arch.FileRef.ToLower())) {
        $salteados++
        Add-Log -Ruta $arch.RutaRelativa -Nombre $arch.Nombre -Bytes $arch.Bytes `
                -Modificado $arch.Modificado -Resultado "SALTADO" -Volumen "" `
                -Detalle "Ya descargado en una corrida previa"
        continue
    }

    # Nombre temporal corto: evita el límite de 260 caracteres de Windows
    $tmpFile = Join-Path $carpetaTmp ([System.Guid]::NewGuid().ToString("N") + ".tmp")
    $bajado  = $false
    $ultimoError = ""

    for ($intento = 1; $intento -le $MaxReintentos -and -not $bajado; $intento++) {
        try {
            Get-PnPFile -Url $arch.FileRef -Path $carpetaTmp `
                        -Filename (Split-Path $tmpFile -Leaf) `
                        -AsFile -Force -Connection $conn -ErrorAction Stop
            $bajado = $true
        }
        catch {
            $ultimoError = $_.Exception.Message -replace "`r`n|`n", " "
            if ($ultimoError -match "429|503|throttl") {
                $espera = 5 * $intento
                Write-Log "    [WARN] Throttling — esperando ${espera}s" -Level WARN
                Start-Sleep -Seconds $espera
            }
            elseif ($intento -lt $MaxReintentos) {
                Start-Sleep -Seconds 2
            }
        }
    }

    if (-not $bajado) {
        $errores++
        Add-Log -Ruta $arch.RutaRelativa -Nombre $arch.Nombre -Bytes $arch.Bytes `
                -Modificado $arch.Modificado -Resultado "ERROR" -Volumen "" -Detalle $ultimoError
        if (Test-Path $tmpFile) { Remove-Item $tmpFile -Force -ErrorAction SilentlyContinue }
        continue
    }

    # --- Guardar: al ZIP o a carpeta suelta ---
    try {
        if ($SinComprimir) {
            $destino = Join-Path $carpetaOut ($arch.RutaRelativa -replace '/', '\')
            $dirDestino = Split-Path $destino -Parent
            if (-not (Test-Path $dirDestino)) { New-Item -ItemType Directory -Path $dirDestino -Force | Out-Null }
            Move-Item -Path $tmpFile -Destination $destino -Force
            $volActual = "(sin comprimir)"
        }
        else {
            # ¿Entra en el volumen actual? Si no, se cierra y se abre el siguiente
            $tamArchivo = (Get-Item $tmpFile).Length
            if ($script:zipFs.Length -gt 0 -and ($script:zipFs.Length + $tamArchivo) -gt $limiteBytes) {
                Flush-Buffers
                Close-Volumen
                Open-Volumen
            }

            [void][System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile(
                $script:zip, $tmpFile, $arch.RutaRelativa, (Get-NivelCompresion $arch.Nombre))

            $volActual = Split-Path $script:zipPath -Leaf
            Remove-Item $tmpFile -Force -ErrorAction SilentlyContinue
        }

        $ok++
        $bytesOk += $arch.Bytes
        Add-Log -Ruta $arch.RutaRelativa -Nombre $arch.Nombre -Bytes $arch.Bytes `
                -Modificado $arch.Modificado -Resultado "OK" -Volumen $volActual -Detalle ""
        Add-Estado -FileRef $arch.FileRef -Volumen $volActual
    }
    catch {
        $errores++
        $msg = $_.Exception.Message -replace "`r`n|`n", " "
        Add-Log -Ruta $arch.RutaRelativa -Nombre $arch.Nombre -Bytes $arch.Bytes `
                -Modificado $arch.Modificado -Resultado "ERROR" -Volumen "" -Detalle "Al comprimir: $msg"
        if (Test-Path $tmpFile) { Remove-Item $tmpFile -Force -ErrorAction SilentlyContinue }
    }

    # --- Progreso ---
    $seg   = [math]::Max(0.1, $sw.Elapsed.TotalSeconds)
    $mbps  = ($bytesOk / 1MB) / $seg
    $pct   = [math]::Round(($procesados / $total) * 100, 1)
    $rest  = $total - $procesados
    $eta   = if ($procesados -gt 0) { ($seg / $procesados) * $rest } else { 0 }

    if (([datetime]::Now - $ultimaBarra).TotalMilliseconds -ge 500) {
        Write-Progress -Activity "Descargando $($lista.Title) — $($web.Title)" `
                       -Status "$procesados/$total | $(ConvertTo-Gb $bytesOk) GB | $([math]::Round($mbps,1)) MB/s" `
                       -CurrentOperation "$($arch.RutaRelativa) — ETA $(Format-Duracion $eta)" `
                       -PercentComplete ([math]::Min(100, $pct))
        $ultimaBarra = [datetime]::Now
    }

    if (([datetime]::Now - $ultimoTexto).TotalSeconds -ge 15 -or $procesados -eq $total) {
        Write-Log ("  {0,6:N1}% — {1}/{2} | {3} GB | {4:N1} MB/s | errores: {5} | ETA {6}" -f `
                    $pct, $procesados, $total, (ConvertTo-Gb $bytesOk), $mbps, $errores, (Format-Duracion $eta)) -Level INFO
        $ultimoTexto = [datetime]::Now
    }

    if ($procesados % 50 -eq 0) { Flush-Buffers }
    if ($PausaMs -gt 0) { Start-Sleep -Milliseconds $PausaMs }
}

Flush-Buffers
if (-not $SinComprimir) { Close-Volumen }
Write-Progress -Activity "Descargando" -Completed
$sw.Stop()

# Limpieza de temporales
try { Remove-Item $carpetaTmp -Recurse -Force -ErrorAction SilentlyContinue } catch { }

# ============================================================
# FALTANTES
# ============================================================
$faltantes = @()
try {
    $faltantes = @(Import-Csv -Path $pathLog -Encoding UTF8 | Where-Object { $_.Resultado -eq "ERROR" })
}
catch { }

if ($faltantes.Count -gt 0) {
    $faltantes | Export-Csv -Path $pathFaltantes -NoTypeInformation -Encoding UTF8
}
else {
    "Ruta,Nombre,Tamano_MB,Modificado,Resultado,Volumen,Detalle,Timestamp" |
        Out-File -FilePath $pathFaltantes -Encoding UTF8
}

# ============================================================
# RESUMEN FINAL
# ============================================================
$seg = [math]::Max(0.1, $sw.Elapsed.TotalSeconds)
$volumenes = @(Get-ChildItem $carpetaJob -Filter "*_vol*.zip" -ErrorAction SilentlyContinue)
$tamZips = if ($volumenes) { ($volumenes | Measure-Object -Property Length -Sum).Sum } else { 0 }

Write-Log "" -Level INFO
Write-Log "======================================================" -Level HEADER
Write-Log "  RESUMEN FINAL"                                         -Level HEADER
Write-Log "  Sitio            : $($web.Title)"                      -Level INFO
Write-Log "  Biblioteca       : $($lista.Title)"                     -Level INFO
Write-Log "  Archivos OK      : $ok de $total"                       -Level OK
Write-Log "  Salteados        : $salteados"                          -Level INFO
Write-Log "  Con error        : $errores"                            -Level $(if ($errores -gt 0) {"ERROR"} else {"OK"})
Write-Log "  Descargado       : $(ConvertTo-Gb $bytesOk) GB de $(ConvertTo-Gb $totalBytes) GB" -Level OK
if (-not $SinComprimir) {
    Write-Log "  Volúmenes ZIP    : $($volumenes.Count) — $(ConvertTo-Gb $tamZips) GB en disco" -Level OK
}
Write-Log "  Duración         : $(Format-Duracion $seg)"             -Level INFO
Write-Log "  Velocidad media  : $([math]::Round(($bytesOk/1MB)/$seg, 1)) MB/s" -Level INFO
Write-Log "  Carpeta          : $carpetaJob"                         -Level INFO
Write-Log "  Log              : $(Split-Path $pathLog -Leaf)"        -Level INFO
Write-Log "======================================================" -Level HEADER

if ($errores -gt 0) {
    Write-Log "FALTARON $errores archivo(s) — detalle en $(Split-Path $pathFaltantes -Leaf)" -Level ERROR
    Write-Log "Para reintentar solo los faltantes:" -Level INFO
    Write-Log "  .\Script_C_DescargarBiblioteca.ps1 -ConfigPath `"$ConfigPath`" -Url `"$Url`" -Biblioteca `"$($lista.Title)`" -Reanudar" -Level INFO

    Write-Log "" -Level WARN
    Write-Log "Primeros faltantes:" -Level WARN
    foreach ($f in ($faltantes | Select-Object -First 5)) {
        Write-Log "  $($f.Ruta)" -Level ERROR
        Write-Log "      $($f.Detalle)" -Level ERROR
    }
}
else {
    Write-Log "Se descargaron todos los archivos. No faltó ninguno." -Level OK
    Write-Log "El archivo _ESTADO.csv se puede borrar." -Level INFO
}

Write-Log "Descarga finalizada." -Level OK
