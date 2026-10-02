<#
===============================================================================
  Script_D_Excel.ps1  |  Circo Studio  |  Kit de Curaduría Informacional M365
===============================================================================

PROPÓSITO:
  Genera un Excel (.xlsx) con los sitios que pasan el filtro de relevamiento
  (la misma lógica que usa Script_B_Masivo.ps1, vía Get-SitiosFiltrados) --
  los sitios descartados por los filtros (no-Legacy, bloqueados, ya
  excluidos, sin storage) NO aparecen en el archivo. En la columna "Nombre"
  se agrega un vínculo al CSV de documentos que generó el Script B para ese
  sitio (relacionado por URL), y las filas cuyo relevamiento terminó en
  error quedan resaltadas con fondo rojo claro.

  100% LOCAL: este script no se conecta a ningún tenant ni a SharePoint
  Online -- no importa PnP.PowerShell, no llama a Connect-PnPOnline, no hace
  ningún request a la nube. Solo lee los CSV que ya están en la carpeta
  configurada y arma el Excel a partir de ellos.

  El resultado es una Tabla real de Excel (la misma que se obtiene con
  Ctrl+T: con nombre "InventarioSitios", filtros nativos en cada columna y
  encabezado con estilo), no un simple rango con autofiltro manual.

USO:
  .\Script_D_Excel.ps1 -ConfigPath "..\..\Config\cliente_config.json"

PARÁMETROS:
  -ConfigPath   (obligatorio) Path al cliente_config.json

REQUIERE HABER CORRIDO ANTES:
  - Script_A_Inventario.ps1  (CSV de inventario de sitios)
  - Script_B_Masivo.ps1      (CSV de resumen + un CSV de documentos por sitio)

MÓDULOS:
  - ImportExcel: si no está instalado, este script lo instala solo la primera
    vez (PSGallery, alcance CurrentUser). Es lo único que toca internet -- y
    ni así toca el tenant del cliente.
  - A propósito NO se usa "#Requires -Modules ImportExcel" al inicio del
    archivo: ese chequeo lo hace PowerShell ANTES de ejecutar una sola línea
    del script, así que si el módulo no está instalado, corta la ejecución
    de entrada y nunca llega al bloque que lo instala solo. Por eso la
    verificación se hace en runtime, más abajo.

ARCHIVOS GENERADOS en la carpeta configurada en el JSON:
  [Cliente]_D_Excel_[timestamp].xlsx
  [Cliente]_D_Excel_[timestamp]_LOG.txt

NOTAS:
  - Solo aparecen los sitios que pasan Get-SitiosFiltrados -- los mismos
    criterios (Legacy, no bloqueado, no excluido, con storage) que usa
    Script_B_Masivo.ps1 para decidir qué relevar.
  - Se agrega una columna "Estado_Relevamiento" al final (OK, ErrorBibliotecas,
    Salteado, o vacío si el sitio pasó el filtro pero Script B todavía no lo
    procesó), para que se entienda por qué una fila está resaltada en rojo.
  - Los vínculos de la columna "Nombre" apuntan al path local absoluto del CSV
    que generó Script B en ESTA MISMA PC (tal como quedó registrado en el CSV
    de Resumen). Si el Excel se comparte con otra persona o PC, esos vínculos
    van a apuntar a una ruta que no existe ahí -- están pensados para uso
    local del analista, no para compartir "tal cual" con el cliente final.
  - Los vínculos se generan como fórmulas nativas de Excel (=HYPERLINK(...)),
    no como relaciones OOXML de "vínculo externo" -- por eso no dispara la
    advertencia de "este libro contiene vínculos a otros orígenes de datos"
    ni el cartel de reparación de contenido al abrir el archivo.
  - Los sitios sin CSV de documentos generado (saltados, con error, sin
    contenido, o no procesados en Script B) quedan sin vínculo, en texto
    plano, sin romper el resto del archivo.

===============================================================================
#>

param(
    [Parameter(Mandatory, HelpMessage="Path al archivo cliente_config.json")]
    [string]$ConfigPath
)

# Oculta las barras de progreso de Install-Module/Import-Module (no son
# "advertencias" pero ensucian la consola sin aportar nada al usuario final)
$ProgressPreference = 'SilentlyContinue'

# ============================================================
# MÓDULO COMPARTIDO DEL KIT
# ============================================================
$moduloPath = Join-Path $PSScriptRoot "M365Curaduria.psm1"
if (-not (Test-Path $moduloPath)) {
    Write-Host "[ERROR] No se encontró el módulo M365Curaduria.psm1 en: $moduloPath" -ForegroundColor Red
    exit 1
}
Import-Module $moduloPath -Force -WarningAction SilentlyContinue

# ============================================================
# MÓDULO ImportExcel -- se instala solo la primera vez, sin tocar el tenant
# ============================================================
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "Módulo ImportExcel no encontrado. Instalando (una sola vez)..." -ForegroundColor Cyan

    try {
        # Evita el prompt interactivo de "repositorio no confiable" y fuerza TLS 1.2
        # (necesario en algunos Windows/PowerShell 5.1 para llegar a PSGallery)
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

        if (-not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)) {
            Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser -ErrorAction Stop | Out-Null
        }

        $repo = Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue
        if ($repo -and $repo.InstallationPolicy -ne 'Trusted') {
            Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction SilentlyContinue
        }

        Install-Module ImportExcel -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop -WarningAction SilentlyContinue
    }
    catch {
        Write-Host "[ERROR] No se pudo instalar el módulo ImportExcel: $_" -ForegroundColor Red
        Write-Host "        Instalalo manualmente con: Install-Module ImportExcel -Scope CurrentUser" -ForegroundColor Red
        exit 1
    }
}
Import-Module ImportExcel -Force -WarningAction SilentlyContinue

# ============================================================
# CONFIGURACIÓN
# ============================================================
$config = $null
try { $config = Read-ClientConfig -ConfigPath $ConfigPath }
catch { Write-Host "[ERROR] $_" -ForegroundColor Red; exit 1 }

if (-not (Test-Path $config.CarpetaLocal)) {
    Write-Host "[ERROR] No existe la carpeta configurada: $($config.CarpetaLocal)" -ForegroundColor Red
    exit 1
}

$timestamp  = Get-Date -Format "yyyyMMdd_HHmmss"
$prefijo    = ($config.NombreCliente -replace "[^a-zA-Z0-9]", "_")
$logFile    = "$($config.CarpetaLocal)\${prefijo}_D_Excel_${timestamp}_LOG.txt"
$outputXlsx = "$($config.CarpetaLocal)\${prefijo}_D_Excel_${timestamp}.xlsx"
New-Item -ItemType File -Path $logFile -Force | Out-Null

Write-Log "======================================================" -Level HEADER -LogFile $logFile
Write-Log "  Kit M365 Curaduría — Script D Excel"                    -Level HEADER -LogFile $logFile
Write-Log "  Cliente : $($config.NombreCliente)"                     -Level HEADER -LogFile $logFile
Write-Log "  (Este script no se conecta a ningún tenant)"            -Level HEADER -LogFile $logFile
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
# BUSCAR CSV DE RESUMEN (Script B) -- mapea URL -> CSV de documentos
# ============================================================
$csvResumen = Get-ChildItem $config.CarpetaLocal -Filter "${prefijo}_B_Masivo_Resumen_*.csv" |
              Sort-Object LastWriteTime -Descending |
              Select-Object -First 1

if (-not $csvResumen) {
    Write-Log "No se encontró el CSV de resumen (Script B) en: $($config.CarpetaLocal)" -Level ERROR -LogFile $logFile
    Write-Log "Correr primero Script_B_Masivo.ps1" -Level ERROR -LogFile $logFile
    exit 1
}
Write-Log "CSV de resumen (Script B): $($csvResumen.Name)" -Level OK -LogFile $logFile

# ============================================================
# CARGAR DATOS
# ============================================================
# Se usa la MISMA funcion que usa Script_B_Masivo.ps1 para decidir que
# sitios se relevan (Get-SitiosFiltrados), asi el Excel muestra exactamente
# los sitios candidatos -- los que quedan afuera por los filtros (no-Legacy,
# bloqueados, ya excluidos, sin storage) directamente no aparecen.
$resumen = @(Import-Csv -Path $csvResumen.FullName -Encoding UTF8)

$sitiosFiltrados = $null
try {
    $sitiosFiltrados = Get-SitiosFiltrados -CsvInventario $csvInventario.FullName -Config $config -LogFile $logFile
}
catch {
    Write-Log "Error al aplicar el filtro de sitios: $_" -Level ERROR -LogFile $logFile
    exit 1
}

if (-not $sitiosFiltrados -or @($sitiosFiltrados).Count -eq 0) {
    Write-Log "Ningún sitio pasó el filtro (Get-SitiosFiltrados). No hay nada para volcar al Excel." -Level WARN -LogFile $logFile
    exit 0
}

# Get-SitiosFiltrados agrega una propiedad interna "_StorageNum" para poder
# ordenar -- la sacamos para que el Excel quede con las mismas columnas que
# el CSV original de Script A, sin campos internos de uso interno del kit.
$inventario = @($sitiosFiltrados | Select-Object -ExcludeProperty _StorageNum)

if ($inventario.Count -eq 0) {
    Write-Log "El CSV de inventario está vacío." -Level ERROR -LogFile $logFile
    exit 1
}

$columnas = @($inventario[0].PSObject.Properties.Name)
if ($columnas -notcontains "Nombre" -or $columnas -notcontains "URL") {
    Write-Log "El CSV de inventario no tiene las columnas 'Nombre' y/o 'URL'. Abortando." -Level ERROR -LogFile $logFile
    exit 1
}

Write-Log "Sitios que pasan el filtro (candidatos a relevamiento): $($inventario.Count)" -Level INFO -LogFile $logFile

# Mapa URL -> Estado del relevamiento (OK, ErrorBibliotecas, Salteado, etc.),
# para poder resaltar en rojo claro las filas que terminaron en error
$urlToEstado = @{}
foreach ($fila in $resumen) {
    if ($fila.URL) { $urlToEstado[$fila.URL] = $fila.Estado }
}

# Agregar la columna "Estado_Relevamiento" al final, para que se entienda
# a simple vista por qué una fila quedó marcada en rojo
foreach ($fila in $inventario) {
    $estado = if ($urlToEstado.ContainsKey($fila.URL)) { $urlToEstado[$fila.URL] } else { "" }
    $fila | Add-Member -NotePropertyName "Estado_Relevamiento" -NotePropertyValue $estado -Force
}
$columnas = @($inventario[0].PSObject.Properties.Name)

# Mapa URL -> path local del CSV de documentos, solo si el archivo existe
# realmente en esta PC (evita vínculos rotos en el Excel)
$urlToCsv     = @{}
$csvFaltantes = 0
foreach ($fila in $resumen) {
    $ruta = $fila.CSV_Salida
    if ($ruta -and $ruta.Trim() -ne "") {
        if (Test-Path $ruta) {
            $urlToCsv[$fila.URL] = $ruta
        }
        else {
            $csvFaltantes++
            Write-Log "CSV de documentos referenciado en el resumen pero no encontrado en disco: $ruta" -Level WARN -LogFile $logFile
        }
    }
}
if ($csvFaltantes -gt 0) {
    Write-Log "$csvFaltantes CSV de documentos figuran en el resumen pero no están en disco (quedan sin vínculo)." -Level WARN -LogFile $logFile
}

# Normalizar Storage_GB / Storage_MB a numero real, si existen, para que
# Excel los reconozca como numeros (ordenables/sumables) y no como texto
foreach ($fila in $inventario) {
    foreach ($campoStorage in @("Storage_GB", "Storage_MB")) {
        if ($columnas -contains $campoStorage) {
            $valorNum = $null
            $textoLimpio = ($fila.$campoStorage -replace ',', '.')
            if ([double]::TryParse($textoLimpio, [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$valorNum)) {
                $fila.$campoStorage = $valorNum
            }
        }
    }
}

# ============================================================
# GENERAR EL EXCEL
# ============================================================
Write-Log "Generando Excel..." -Level INFO -LogFile $logFile

# Colorea un rango detectando EN RUNTIME que firma de SetColor() acepta la
# version de EPPlus instalada -- no hay una firma unica que funcione en
# todas las versiones (algunas piden (Color), otras (Color, double), etc.),
# asi que se prueban las variantes conocidas por reflexion en vez de asumir
# una fija. Si ninguna funciona, devuelve $false y quien llama decide que
# hacer (en este script: loguear un WARN y seguir sin ese color puntual).
function Set-ColorEnRango {
    param(
        [Parameter(Mandatory)] $ColorObj,   # ej: $Range.Style.Fill.BackgroundColor
        [Parameter(Mandatory)] [string]$HexColor   # ej: "085675" (sin "#", sin alfa)
    )

    $color = [System.Drawing.ColorTranslator]::FromHtml("#$HexColor")
    $metodos = $ColorObj.GetType().GetMethods() | Where-Object { $_.Name -eq 'SetColor' }

    foreach ($m in $metodos) {
        $parametros = $m.GetParameters()
        try {
            if ($parametros.Count -eq 1 -and $parametros[0].ParameterType.Name -match 'Color') {
                $m.Invoke($ColorObj, @($color)) | Out-Null
                return $true
            }
            if ($parametros.Count -eq 2 -and $parametros[0].ParameterType.Name -match 'Color') {
                # variante (Color, double) usada en algunas versiones para tint/shade
                $m.Invoke($ColorObj, @($color, [double]0)) | Out-Null
                return $true
            }
        }
        catch { continue }
    }

    # Ultimo intento: la propiedad .Rgb, por si en esta version SI es escribible
    try {
        $ColorObj.Rgb = "FF" + $HexColor
        return $true
    }
    catch { }

    return $false
}

function Set-RangoColor {
    param(
        [Parameter(Mandatory)] $Range,
        [string]$FondoHex = $null,    # sin "#", ej "085675"
        [switch]$Bold
    )
    $Range.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
    if ($FondoHex) {
        $ok = Set-ColorEnRango -ColorObj $Range.Style.Fill.BackgroundColor -HexColor $FondoHex
        if (-not $ok) { throw "No se encontró una firma compatible de SetColor() en esta versión de EPPlus." }
    }
    if ($Bold) { $Range.Style.Font.Bold = $true }
}

$colNombre = [array]::IndexOf($columnas, "Nombre") + 1   # EPPlus usa índices base 1
$vinculados = 0

try {
    # -TableName + -TableStyle arma una Tabla real de Excel (la misma que se
    # obtiene con Ctrl+T): con nombre, filtros nativos, encabezado con
    # estilo y bandas de color automaticas -- no un simple rango con
    # autofiltro. Esto tambien resuelve el color del encabezado sin
    # necesidad de tocar el objeto de estilo a mano (que es lo que venia
    # fallando por la version de EPPlus instalada).
    $excel = $inventario | Export-Excel -Path $outputXlsx `
        -WorksheetName "Inventario de Sitios" `
        -TableName "InventarioSitios" -TableStyle Medium2 `
        -AutoSize -FreezeTopRow `
        -PassThru

    $ws = $excel.Workbook.Worksheets["Inventario de Sitios"]

    # Formato numerico para Storage_GB / Storage_MB, en su propio try/catch
    # (independiente del color, para que un fallo no tape al otro). Se
    # intenta primero via el cmdlet, y si esa via tambien falla en esta
    # version de EPPlus, se prueba asignando la propiedad directamente.
    try {
        foreach ($par in @{ "Storage_GB" = "#,##0.00"; "Storage_MB" = "#,##0" }.GetEnumerator()) {
            if ($columnas -contains $par.Key) {
                $colIdx   = [array]::IndexOf($columnas, $par.Key) + 1
                $celdaIni = $ws.Cells[2, $colIdx]
                $rangoCol = $ws.Cells[2, $colIdx, $inventario.Count + 1, $colIdx]
                try {
                    Set-ExcelRange -Worksheet $ws -Range $rangoCol.Address -NumberFormat $par.Value
                }
                catch {
                    $rangoCol.Style.Numberformat.Format = $par.Value
                }
            }
        }
    }
    catch {
        Write-Log "No se pudo aplicar el formato numérico de Storage_GB/MB -- se sigue igual, es solo estético. Detalle: $_" -Level WARN -LogFile $logFile
    }

    # Vínculo en "Nombre" -> CSV de documentos del sitio (por URL), como
    # fórmula HYPERLINK nativa -- no como relación externa OOXML.
    # Además, fondo rojo claro en toda la fila si el relevamiento terminó
    # en error (cualquier Estado que empiece con "Error").
    $coloreoErrorFallido = $false
    $rangoFilaCompleta = $null
    $erroresColoreados = 0

    for ($i = 0; $i -lt $inventario.Count; $i++) {
        $excelRow = $i + 2   # la fila 1 es el encabezado
        $urlSitio = $inventario[$i].URL
        $estado   = $inventario[$i].Estado_Relevamiento

        if ($urlSitio -and $urlToCsv.ContainsKey($urlSitio)) {
            $nombreEscapado = ($inventario[$i].Nombre) -replace '"', '""'
            $rutaEscapada   = ($urlToCsv[$urlSitio]) -replace '"', '""'
            $ws.Cells[$excelRow, $colNombre].Formula = "HYPERLINK(""$rutaEscapada"",""$nombreEscapado"")"
            $vinculados++
        }

        if ($estado -match '(?i)^error' -and -not $coloreoErrorFallido) {
            try {
                $rangoFilaCompleta = $ws.Cells[$excelRow, 1, $excelRow, $columnas.Count]
                Set-RangoColor -Range $rangoFilaCompleta -FondoHex "FADBD8"
                $erroresColoreados++
            }
            catch {
                # Mismo criterio que el estilo del encabezado: si falla por
                # algun choque de version de EPPlus, no aborta el script --
                # se avisa una sola vez y se sigue sin colorear el resto.
                $coloreoErrorFallido = $true
                Write-Log "No se pudo aplicar el resaltado en rojo de filas con error -- se sigue igual, es solo estético. Detalle: $_" -Level WARN -LogFile $logFile
            }
        }
    }

    Close-ExcelPackage $excel
}
catch {
    Write-Log "Error al generar el Excel: $_" -Level ERROR -LogFile $logFile
    exit 1
}

# ============================================================
# RESUMEN
# ============================================================
Write-Log "------------------------------------------------------" -Level HEADER -LogFile $logFile
Write-Log "Excel generado : $outputXlsx"                             -Level OK     -LogFile $logFile
Write-Log "Sitios con vínculo a su CSV de documentos: $vinculados de $($inventario.Count)" -Level OK -LogFile $logFile
Write-Log "Sitios resaltados en rojo (error en el relevamiento): $erroresColoreados" -Level OK -LogFile $logFile
Write-Log "======================================================" -Level HEADER -LogFile $logFile
Write-Log "Script D finalizado." -Level OK -LogFile $logFile