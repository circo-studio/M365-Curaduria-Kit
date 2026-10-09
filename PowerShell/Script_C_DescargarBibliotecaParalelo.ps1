#Requires -Version 7.2
#Requires -Modules PnP.PowerShell
<#
===============================================================================
  Script_C_DescargarBiblioteca.ps1  |  Circo Studio  |  Kit de Curaduría M365
  Versión 2.0 — descarga en paralelo
===============================================================================

PROPÓSITO:
  Descarga el contenido completo de una biblioteca de documentos y lo guarda
  en volúmenes ZIP (o en carpetas sueltas), preservando la estructura.
  Pensado para bibliotecas grandes (cientos de GB, cientos de miles de archivos).

CÓMO DESCARGA (cambios respecto de la v1):
  - PnP se usa solo para conectar, listar y enumerar. Los archivos se bajan
    por la API REST de SharePoint con un HttpClient compartido, con -Hilos
    descargas simultáneas (default 8). La v1 bajaba de a un archivo por vez.
  - El token se obtiene en el hilo principal (Get-PnPAccessToken) y se
    renueva solo antes de vencer: sirve para corridas de muchas horas.
  - Throttling (429/503): se respeta el header Retry-After y se pausan TODOS
    los hilos, como pide Microsoft. Esas esperas no consumen reintentos.
  - Todo el tráfico va identificado con el User-Agent de ISV de Circo Studio,
    que SharePoint prioriza frente al tráfico sin identificar.
  - Se detectan conexiones colgadas (sin datos durante -SegundosSinDatos) y
    descargas truncadas (bytes recibidos distintos del Content-Length).
  - Mientras el hilo principal comprime un archivo, las demás descargas
    siguen corriendo en segundo plano.
  - Respeta el método de autenticación del cliente_config.json
    (Interactivo o Certificado app-only).

INTEGRIDAD Y REANUDACIÓN:
  - Un archivo se registra en _ESTADO.csv recién cuando el volumen ZIP que lo
    contiene se cerró y validó. Si la PC se apaga a mitad de un volumen, ese
    volumen se descarta y sus archivos se vuelven a bajar al reanudar: nunca
    queda un archivo marcado como bajado sin estar realmente guardado.
  - Ctrl+C cierra y valida el volumen en curso antes de salir.
  - Con -Reanudar se validan los volúmenes existentes, se apartan los rotos
    (*.zip.incompleto) y la numeración continúa (la v1 sobrescribía vol01).
  - Cada archivo conserva su fecha de modificación de SharePoint.

USO:
  # Interactivo: pide URL y muestra las bibliotecas del sitio
  .\Script_C_DescargarBiblioteca.ps1 -ConfigPath "..\..\Config\cliente_config.json"

  # Directo, 12 descargas simultáneas
  .\Script_C_DescargarBiblioteca.ps1 -ConfigPath ".\cliente_config.json" `
      -Url "https://tenant.sharepoint.com/sites/EJEMPLO" -Biblioteca "Documents" -Hilos 12

  # Lo más rápido para un backup: archivos sueltos, sin ZIP
  .\Script_C_DescargarBiblioteca.ps1 -ConfigPath ".\cliente_config.json" `
      -Url "https://..." -Biblioteca "Documents" -SinComprimir

  # Retomar una descarga interrumpida
  .\Script_C_DescargarBiblioteca.ps1 -ConfigPath ".\cliente_config.json" `
      -Url "https://..." -Biblioteca "Documents" -Reanudar

  # Solo una carpeta, solo PDFs, volúmenes de 2 GB
  .\Script_C_DescargarBiblioteca.ps1 -ConfigPath ".\cliente_config.json" `
      -Url "https://..." -Biblioteca "Documents" `
      -SoloCarpeta "Contratos/2024" -Extensiones pdf -TamanoVolumenGB 2

  # Ver qué bajaría, sin bajar nada
  .\Script_C_DescargarBiblioteca.ps1 -ConfigPath ".\cliente_config.json" -Url "..." -Simular

AJUSTE DE -Hilos:
  Empezar con 8. Si el resumen final muestra 0 esperas por throttling y la
  velocidad está lejos del ancho de banda, probar 12 o 16. Si aparecen
  muchas esperas, bajar a 4-6: más hilos con throttling es más lento.

ARCHIVOS GENERADOS (en [Destino]\[Sitio]_[Biblioteca]\):
  [Sitio]_[Biblioteca]_vol01.zip, _vol02.zip, ...   — los volúmenes
  Archivos\                — solo con -SinComprimir
  _LOG.csv                 — una fila por archivo: OK | ERROR | SALTADO
  _FALTANTES.csv           — los que no quedaron guardados (vacío si salió todo)
  _ESTADO.csv              — control interno de reanudación (no borrar hasta terminar)
  _EJECUCION_LOG.txt       — log de la ejecución (lo mismo que se ve en consola)
  _SIMULACION.csv          — solo con -Simular

REQUISITOS:
  - PowerShell 7.2 o superior + PnP.PowerShell.
  - Para rutas de más de 260 caracteres con -SinComprimir: LongPathsEnabled=1
    en Windows (el script avisa si hace falta).

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

    # Descargas simultáneas
    [ValidateRange(1, 32)]
    [int]$Hilos = 8,

    # Compresión para formatos comprimibles (txt, csv, xml, dwg...).
    # Los ya comprimidos (pdf, docx, jpg, zip, mp4...) se guardan siempre sin recomprimir.
    [ValidateSet("Rapida", "Optima", "Ninguna")]
    [string]$Compresion = "Rapida",

    # Reintentos por archivo ante errores de red (el throttling no cuenta)
    [ValidateRange(1, 20)]
    [int]$MaxReintentos = 5,

    # Corta y reintenta una descarga que no recibe datos durante este tiempo
    [ValidateRange(10, 3600)]
    [int]$SegundosSinDatos = 120,

    # Identificación del tráfico ante SharePoint (formato ISV|Empresa|App/Versión)
    [string]$UserAgent = "ISV|CircoStudio|KitCuraduria/2.0",

    # Pausa entre el inicio de cada descarga, en milisegundos. Normalmente 0:
    # el throttling ya se maneja con Retry-After. Se mantiene por compatibilidad.
    [int]$PausaMs = 0,

    [switch]$SinConfirmacion,
    [switch]$ForceAuth
)

$moduloPath = Join-Path $PSScriptRoot "M365Curaduria.psm1"
if (-not (Test-Path -LiteralPath $moduloPath)) {
    Write-Host "[ERROR] No se encontró M365Curaduria.psm1 en: $moduloPath" -ForegroundColor Red
    exit 1
}
Import-Module $moduloPath -Force

Add-Type -AssemblyName System.IO.Compression | Out-Null
Add-Type -AssemblyName System.IO.Compression.FileSystem | Out-Null

# ============================================================
# DESCARGADOR (C#)
# PowerShell no puede usar callbacks async de .NET de forma segura, así que
# la parte concurrente vive en una clase compilada al vuelo. El hilo
# principal (PowerShell) solo lanza tareas, recoge resultados, escribe logs
# y arma los ZIP: un único escritor, sin locks sobre archivos.
# ============================================================
if (-not ('CircoKit.Descargador' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading;
using System.Threading.Tasks;

namespace CircoKit
{
    public sealed class ResultadoDescarga
    {
        public int Id;
        public bool Ok;
        public bool AuthVencida;
        public int Status;
        public long Bytes;
        public long BytesEsperados = -1;
        public int Intentos;
        public int Throttles;
        public string Error = "";
    }

    // Descargas HTTP en paralelo contra la API REST de SharePoint.
    // - Un solo HttpClient compartido (thread-safe) con pool de conexiones.
    // - 429/503: respeta Retry-After y pausa a TODOS los hilos (lo que pide Microsoft).
    // - Detecta conexiones colgadas (sin datos durante N segundos) y descargas truncadas.
    // - El token se cambia en caliente con SetToken() desde el hilo principal.
    public sealed class Descargador : IDisposable
    {
        private readonly HttpClient _http;
        private readonly CancellationTokenSource _global = new CancellationTokenSource();
        private volatile string _token = "";
        private long _pausaHastaTicks;
        private long _bytesEnVivo;
        private long _throttles;

        public int MaxReintentos { get; set; } = 5;
        public int MaxEsperasThrottling { get; set; } = 30;
        public int SegundosSinDatos { get; set; } = 120;

        public long BytesEnVivo { get { return Interlocked.Read(ref _bytesEnVivo); } }
        public long ThrottlesTotal { get { return Interlocked.Read(ref _throttles); } }
        public DateTime PausaHastaUtc { get { return new DateTime(Interlocked.Read(ref _pausaHastaTicks), DateTimeKind.Utc); } }

        public Descargador(string userAgent, int maxConexiones)
        {
            var handler = new SocketsHttpHandler
            {
                MaxConnectionsPerServer = Math.Max(1, maxConexiones),
                AutomaticDecompression = DecompressionMethods.None,
                PooledConnectionLifetime = TimeSpan.FromMinutes(15),
                ConnectTimeout = TimeSpan.FromSeconds(30)
            };
            _http = new HttpClient(handler) { Timeout = Timeout.InfiniteTimeSpan };
            if (!string.IsNullOrWhiteSpace(userAgent))
                _http.DefaultRequestHeaders.TryAddWithoutValidation("User-Agent", userAgent);
        }

        public void SetToken(string token) { _token = token ?? ""; }

        public void Cancelar() { try { _global.Cancel(); } catch { } }

        // URL de descarga por ruta server-relative. decodedurl + alias @a soporta
        // nombres con #, %, comillas y espacios sin problemas de codificación.
        public static string ArmarUrl(string webUrl, string serverRelativeUrl)
        {
            string literal = "'" + serverRelativeUrl.Replace("'", "''") + "'";
            return webUrl.TrimEnd('/') +
                   "/_api/web/GetFileByServerRelativePath(decodedurl=@a)/$value?@a=" +
                   Uri.EscapeDataString(literal);
        }

        public Task<ResultadoDescarga> DescargarAsync(int id, string url, string destino, DateTime modificadoUtc)
        {
            return Task.Run(() => Descargar(id, url, destino, modificadoUtc));
        }

        private async Task<ResultadoDescarga> Descargar(int id, string url, string destino, DateTime modificadoUtc)
        {
            var r = new ResultadoDescarga { Id = id };
            int fallos = 0;

            while (true)
            {
                if (_global.IsCancellationRequested) { r.Error = "Cancelado"; return r; }
                long escritos = 0;
                try
                {
                    await EsperarPausaGlobal().ConfigureAwait(false);
                    r.Intentos++;

                    using (var req = new HttpRequestMessage(HttpMethod.Get, url))
                    {
                        req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", _token);

                        using (var resp = await _http.SendAsync(req, HttpCompletionOption.ResponseHeadersRead, _global.Token).ConfigureAwait(false))
                        {
                            int st = (int)resp.StatusCode;
                            r.Status = st;

                            if (st == 429 || st == 503)
                            {
                                r.Throttles++;
                                Interlocked.Increment(ref _throttles);
                                FijarPausaGlobal(LeerRetryAfter(resp, r.Throttles));
                                if (r.Throttles > MaxEsperasThrottling)
                                {
                                    r.Error = "HTTP " + st + " (throttling persistente)";
                                    return r;
                                }
                                continue; // el throttling no consume reintentos
                            }
                            if (st == 401)
                            {
                                r.AuthVencida = true;
                                r.Error = "HTTP 401 (token vencido o sin permiso)";
                                return r;
                            }
                            if (st == 403 || st == 404)
                            {
                                r.Error = "HTTP " + st + " " + resp.ReasonPhrase;
                                return r;
                            }
                            if (!resp.IsSuccessStatusCode)
                                throw new HttpRequestException("HTTP " + st + " " + resp.ReasonPhrase);

                            long? largo = resp.Content.Headers.ContentLength;
                            r.BytesEsperados = largo ?? -1;

                            using (var origen = await resp.Content.ReadAsStreamAsync().ConfigureAwait(false))
                            using (var fs = new FileStream(destino, FileMode.Create, FileAccess.Write, FileShare.None, 1 << 16, FileOptions.Asynchronous))
                            {
                                var buffer = new byte[1 << 20];
                                while (true)
                                {
                                    int n;
                                    using (var cts = CancellationTokenSource.CreateLinkedTokenSource(_global.Token))
                                    {
                                        cts.CancelAfter(TimeSpan.FromSeconds(SegundosSinDatos));
                                        n = await origen.ReadAsync(buffer, 0, buffer.Length, cts.Token).ConfigureAwait(false);
                                    }
                                    if (n == 0) break;
                                    await fs.WriteAsync(buffer, 0, n).ConfigureAwait(false);
                                    escritos += n;
                                    Interlocked.Add(ref _bytesEnVivo, n);
                                }
                            }

                            if (largo.HasValue && largo.Value != escritos)
                                throw new IOException("Descarga incompleta: " + escritos + " de " + largo.Value + " bytes");
                        }
                    }

                    try { File.SetLastWriteTimeUtc(destino, modificadoUtc); } catch { }
                    r.Ok = true;
                    r.Bytes = escritos;
                    r.Error = "";
                    return r;
                }
                catch (Exception ex)
                {
                    Interlocked.Add(ref _bytesEnVivo, -escritos);
                    BorrarSilencioso(destino);
                    if (_global.IsCancellationRequested) { r.Error = "Cancelado"; return r; }

                    r.Error = (ex is OperationCanceledException)
                        ? "Sin datos durante " + SegundosSinDatos + " s (conexión colgada)"
                        : ex.GetBaseException().Message;

                    fallos++;
                    if (fallos >= MaxReintentos) return r;
                    try
                    {
                        await Task.Delay(TimeSpan.FromSeconds(Math.Min(60, 2 * Math.Pow(2, fallos - 1))), _global.Token).ConfigureAwait(false);
                    }
                    catch { r.Error = "Cancelado"; return r; }
                }
            }
        }

        private static TimeSpan LeerRetryAfter(HttpResponseMessage resp, int n)
        {
            var ra = resp.Headers.RetryAfter;
            if (ra != null)
            {
                if (ra.Delta.HasValue) return Acotar(ra.Delta.Value);
                if (ra.Date.HasValue) return Acotar(ra.Date.Value - DateTimeOffset.UtcNow);
            }
            return Acotar(TimeSpan.FromSeconds(Math.Min(120, 5 * Math.Pow(2, Math.Min(n, 6) - 1))));
        }

        private static TimeSpan Acotar(TimeSpan t)
        {
            if (t < TimeSpan.FromSeconds(1)) return TimeSpan.FromSeconds(1);
            if (t > TimeSpan.FromMinutes(10)) return TimeSpan.FromMinutes(10);
            return t;
        }

        private void FijarPausaGlobal(TimeSpan espera)
        {
            long hasta = DateTime.UtcNow.Add(espera).Ticks;
            while (true)
            {
                long actual = Interlocked.Read(ref _pausaHastaTicks);
                if (hasta <= actual) return;
                if (Interlocked.CompareExchange(ref _pausaHastaTicks, hasta, actual) == actual) return;
            }
        }

        private async Task EsperarPausaGlobal()
        {
            while (true)
            {
                TimeSpan resta = new DateTime(Interlocked.Read(ref _pausaHastaTicks), DateTimeKind.Utc) - DateTime.UtcNow;
                if (resta <= TimeSpan.Zero) return;
                await Task.Delay(resta, _global.Token).ConfigureAwait(false);
            }
        }

        private static void BorrarSilencioso(string path)
        {
            try { if (File.Exists(path)) File.Delete(path); } catch { }
        }

        public void Dispose()
        {
            Cancelar();
            _http.Dispose();
            _global.Dispose();
        }
    }
}
'@
}

# ============================================================
# HELPERS
# ============================================================
$script:logTxt = ""

function Write-L {
    param([AllowEmptyString()][string]$Mensaje = "", [string]$Nivel = "INFO")
    # Nunca debe cortar el flujo (ej: log abierto por otro proceso)
    try { Write-Log -Message $Mensaje -Level $Nivel -LogFile $script:logTxt } catch { }
}

function ConvertTo-Gb {
    param([double]$Bytes)
    return [math]::Round($Bytes / 1GB, 2)
}

function Format-Duracion {
    param([double]$Segundos)
    if ($Segundos -le 0 -or [double]::IsInfinity($Segundos) -or [double]::IsNaN($Segundos)) { return "--" }
    $ts = [timespan]::FromSeconds($Segundos)
    if ($ts.TotalHours -ge 1)   { return "{0}h {1}m" -f [int][math]::Floor($ts.TotalHours), $ts.Minutes }
    if ($ts.TotalMinutes -ge 1) { return "{0}m {1}s" -f [int][math]::Floor($ts.TotalMinutes), $ts.Seconds }
    return "{0}s" -f [int]$ts.TotalSeconds
}

# Formatos que ya vienen comprimidos: recomprimirlos gasta CPU sin ganar espacio
$YA_COMPRIMIDOS = [System.Collections.Generic.HashSet[string]]::new(
    [string[]]@(
        "zip","7z","rar","gz","bz2","xz","cab",
        "jpg","jpeg","png","gif","webp","heic",
        "mp3","mp4","avi","mkv","mov","wmv","m4a","aac","m4v",
        "pdf","docx","xlsx","pptx","docm","xlsm","pptm","odt","ods","odp","xlsb"
    ),
    [System.StringComparer]::OrdinalIgnoreCase)

$NIVEL_COMPRIMIBLES = switch ($Compresion) {
    "Optima"  { [System.IO.Compression.CompressionLevel]::Optimal }
    "Ninguna" { [System.IO.Compression.CompressionLevel]::NoCompression }
    default   { [System.IO.Compression.CompressionLevel]::Fastest }
}

function Get-NivelCompresion {
    param([string]$NombreArchivo)
    $ext = ([System.IO.Path]::GetExtension($NombreArchivo)).TrimStart('.')
    if ($YA_COMPRIMIDOS.Contains($ext)) {
        return [System.IO.Compression.CompressionLevel]::NoCompression
    }
    return $NIVEL_COMPRIMIBLES
}

function Get-NombreSeguro {
    param([string]$Texto)
    return ($Texto -replace '[^a-zA-Z0-9\-_]', '_')
}

function Get-NumeroVolumen {
    param([string]$Nombre)
    if ($Nombre -match '_vol(\d+)\.zip') { return [int]$Matches[1] }
    return 0
}

# Vencimiento (UTC) de un token JWT, leyendo el claim "exp"
function Get-JwtVencimientoUtc {
    param([string]$Jwt)
    try {
        $p = $Jwt.Split('.')[1].Replace('-', '+').Replace('_', '/')
        switch ($p.Length % 4) { 2 { $p += '==' } 3 { $p += '=' } }
        $claims = [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($p)) | ConvertFrom-Json
        return [DateTimeOffset]::FromUnixTimeSeconds([long]$claims.exp).UtcDateTime
    }
    catch {
        # Si no se puede leer, se asume vida corta y se renueva seguido
        return [datetime]::UtcNow.AddMinutes(20)
    }
}

# ============================================================
# CONFIGURACIÓN
# ============================================================
$config = $null
try { $config = Read-ClientConfig -ConfigPath $ConfigPath }
catch { Write-Host "[ERROR] $_" -ForegroundColor Red; exit 1 }

# Read-ClientConfig no expone el método de autenticación: se lee del JSON.
# Soporta el formato del editor (cliente.auth_method / thumbprint / tenantID)
# y el formato con sección "auth" (auth.modo / thumbprint / tenant_id).
$rawConfig = $null
try { $rawConfig = Get-Content -LiteralPath $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json } catch { }

$authMetodo = [string]$rawConfig.auth.modo
if (-not $authMetodo) { $authMetodo = [string]$rawConfig.cliente.auth_method }
$authThumb  = [string]$rawConfig.auth.thumbprint
if (-not $authThumb)  { $authThumb  = [string]$rawConfig.cliente.thumbprint }
$authTenant = [string]$rawConfig.auth.tenant_id
if (-not $authTenant) { $authTenant = [string]$rawConfig.auth.tenantID }
if (-not $authTenant) { $authTenant = [string]$rawConfig.cliente.tenantID }

$authMetodo = switch -Regex ($authMetodo.Trim().ToLower()) {
    '^(cert|app)' { "Certificado"; break }
    '^inter'      { "Interactivo"; break }
    default       { "" }
}

Write-L "======================================================" HEADER
Write-L "  Kit M365 Curaduría — Descarga de biblioteca (v2)"     HEADER
Write-L "  Cliente : $($config.NombreCliente)"                   HEADER
Write-L "======================================================" HEADER

# --- URL del sitio ---
if (-not $Url) {
    Write-L ""
    $Url = Read-Host "URL del sitio (ej: $($config.SharePointUrl)/sites/EJEMPLO)"
}
$Url = ([string]$Url).Trim().Trim('"').TrimEnd('/')

if (-not $Url -or $Url -notmatch '^https?://') {
    Write-L "URL inválida. Debe empezar con https://" ERROR
    exit 1
}

# --- Método de autenticación ---
if (-not $authMetodo) {
    if ($SinConfirmacion) {
        $authMetodo = "Interactivo"
    }
    else {
        Write-Host ""
        Write-Host "  Método de autenticación:" -ForegroundColor Cyan
        Write-Host "    1) Interactivo — login de usuario (ventana / MFA)"
        Write-Host "    2) Certificado — app-only"
        $op = (Read-Host "Opción [1]").Trim()
        $authMetodo = if ($op -eq "2") { "Certificado" } else { "Interactivo" }
    }
}

if ($authMetodo -eq "Certificado" -and (-not $authThumb -or -not $authTenant)) {
    Write-L "Autenticación por certificado sin thumbprint o tenant ID en el config." ERROR
    Write-L "Completarlos en el editor HTML del kit o usar el método Interactivo." ERROR
    exit 1
}

# ============================================================
# CONEXIÓN
# ============================================================
Write-L ""
Write-L "Conectando a $Url ($authMetodo)..."

$conn = $null
try {
    $connectParams = @{
        Url              = $Url
        ClientId         = $config.ClientId
        ReturnConnection = $true
        ErrorAction      = "Stop"
    }
    if ($authMetodo -eq "Certificado") {
        $connectParams["Thumbprint"] = $authThumb
        $connectParams["Tenant"]     = $authTenant
    }
    else {
        $connectParams["Interactive"] = $true
        if ($ForceAuth) { $connectParams["ForceAuthentication"] = $true }
        else            { $connectParams["PersistLogin"] = $true }
    }

    $conn = Connect-PnPOnline @connectParams
    $web  = Get-PnPWeb -Connection $conn -ErrorAction Stop
    Write-L "Conexión OK: $($web.Title)" OK
}
catch {
    Write-L "No se pudo conectar: $($_.Exception.Message)" ERROR
    Write-L "Verificar que la cuenta (o la app) tenga acceso al sitio." ERROR
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
    Write-L "No se pudieron listar las bibliotecas: $($_.Exception.Message)" ERROR
    exit 1
}

if ($listas.Count -eq 0) {
    Write-L "El sitio no tiene bibliotecas de documentos visibles." ERROR
    exit 1
}

if (-not $Biblioteca) {
    Write-L ""
    Write-Host "  Bibliotecas del sitio:" -ForegroundColor Cyan
    for ($i = 0; $i -lt $listas.Count; $i++) {
        Write-Host ("    {0}) {1}  ({2} ítems)" -f ($i + 1), $listas[$i].Title, $listas[$i].ItemCount)
    }
    Write-Host "    0) Cancelar"
    $op = Read-Host "Opción"

    $n = 0
    if (-not [int]::TryParse($op.Trim(), [ref]$n) -or $n -lt 1 -or $n -gt $listas.Count) {
        Write-L "Cancelado." WARN
        exit 0
    }
    $Biblioteca = $listas[$n - 1].Title
}

$lista = $listas | Where-Object { $_.Title -eq $Biblioteca } | Select-Object -First 1
if (-not $lista) {
    try { $lista = Get-PnPList -Identity $Biblioteca -Connection $conn -ErrorAction Stop }
    catch {
        Write-L "No se encontró la biblioteca '$Biblioteca' en el sitio." ERROR
        exit 1
    }
}

Write-L ""
Write-L "Biblioteca: $($lista.Title) — $($lista.ItemCount) ítems (incluye carpetas)" OK

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

[void][System.IO.Directory]::CreateDirectory($carpetaJob)

# Temporales de una corrida anterior interrumpida: se descartan
if ([System.IO.Directory]::Exists($carpetaTmp)) {
    try { [System.IO.Directory]::Delete($carpetaTmp, $true) } catch { }
}
[void][System.IO.Directory]::CreateDirectory($carpetaTmp)
if ($SinComprimir) { [void][System.IO.Directory]::CreateDirectory($carpetaOut) }

$pathLog       = Join-Path $carpetaJob "_LOG.csv"
$pathFaltantes = Join-Path $carpetaJob "_FALTANTES.csv"
$pathEstado    = Join-Path $carpetaJob "_ESTADO.csv"
$script:logTxt = Join-Path $carpetaJob "_EJECUCION_LOG.txt"

Write-L "Destino   : $carpetaJob"
Write-L "Auth      : $authMetodo | Hilos: $Hilos | Compresión: $(if ($SinComprimir) {'sin ZIP'} else {$Compresion})"
Write-L "User-Agent: $UserAgent"

# ============================================================
# ENUMERACIÓN
# ============================================================
Write-L ""
Write-L "Enumerando archivos (puede tardar en bibliotecas grandes)..."

$swEnum = [System.Diagnostics.Stopwatch]::StartNew()
$items = @()

try {
    $items = @(Get-PnPListItem -List $lista.Title -PageSize 1000 `
                               -Fields "FileRef","FileLeafRef","File_x0020_Size","Modified","FSObjType" `
                               -Connection $conn -ErrorAction Stop)
}
catch {
    Write-L "Error al enumerar: $($_.Exception.Message)" ERROR
    exit 1
}

$swEnum.Stop()
Write-L "  $($items.Count) ítems leídos en $(Format-Duracion $swEnum.Elapsed.TotalSeconds)" OK

# Ruta raíz de la biblioteca, para calcular rutas relativas
$rootUrl = $lista.RootFolder.ServerRelativeUrl
if (-not $rootUrl) {
    $rootUrl = (Get-PnPList -Identity $lista.Title -Includes RootFolder -Connection $conn).RootFolder.ServerRelativeUrl
}
$rootUrl = $rootUrl.TrimEnd('/')

# --- Filtrado ---
$archivos = [System.Collections.Generic.List[object]]::new()
$omitidosCarpeta = 0; $omitidosExt = 0; $omitidosTam = 0; $omitidosFecha = 0
$prefijoCarpeta = if ($SoloCarpeta) { $SoloCarpeta.Trim('/').Replace('\', '/') } else { "" }
$extFiltro = [System.Collections.Generic.HashSet[string]]::new(
    [string[]]@($Extensiones | ForEach-Object { ([string]$_).Trim().TrimStart('.') } | Where-Object { $_ }),
    [System.StringComparer]::OrdinalIgnoreCase)
$limiteTam = [long]$TamanoMaxMB * 1MB

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

    if ($prefijoCarpeta -and -not $rel.StartsWith("$prefijoCarpeta/", [System.StringComparison]::OrdinalIgnoreCase)) {
        $omitidosCarpeta++; continue
    }

    if ($extFiltro.Count -gt 0) {
        $ext = ([System.IO.Path]::GetExtension($fileName)).TrimStart('.')
        if (-not $extFiltro.Contains($ext)) { $omitidosExt++; continue }
    }

    [long]$tam = 0
    [void][long]::TryParse([string]$item.FieldValues["File_x0020_Size"], [ref]$tam)

    if ($limiteTam -gt 0 -and $tam -gt $limiteTam) { $omitidosTam++; continue }

    $mod = $item.FieldValues["Modified"]
    $modUtc = [datetime]::UtcNow
    if ($mod) {
        $modDt  = [datetime]$mod
        $modUtc = if ($modDt.Kind -eq [System.DateTimeKind]::Local) { $modDt.ToUniversalTime() }
                  else { [datetime]::SpecifyKind($modDt, [System.DateTimeKind]::Utc) }
        if ($ModificadoDesde -and $modDt -lt $ModificadoDesde) { $omitidosFecha++; continue }
    }

    $archivos.Add([PSCustomObject]@{
        FileRef        = $fileRef
        Clave          = $fileRef.ToLowerInvariant()
        Nombre         = $fileName
        RutaRelativa   = $rel
        Bytes          = $tam
        Modificado     = if ($mod) { $modUtc.ToLocalTime().ToString("yyyy-MM-dd HH:mm") } else { "" }
        ModificadoUtc  = $modUtc
        ReintentosAuth = 0
    })
}
$items = $null   # liberar memoria: en bibliotecas grandes son cientos de MB

$totalBytes = [long]0
foreach ($a in $archivos) { $totalBytes += $a.Bytes }

Write-L ""
Write-L "======================================================" HEADER
Write-L "  A DESCARGAR"                                          HEADER
Write-L "  Archivos : $($archivos.Count)"                        OK
Write-L "  Tamaño   : $(ConvertTo-Gb $totalBytes) GB"            OK
if (-not $SinComprimir) {
    Write-L "  Volúmenes: ~$([math]::Max(1, [math]::Ceiling((ConvertTo-Gb $totalBytes) / $TamanoVolumenGB))) de hasta $TamanoVolumenGB GB"
}
if ($omitidosCarpeta) { Write-L "  Omitidos por carpeta   : $omitidosCarpeta" }
if ($omitidosExt)     { Write-L "  Omitidos por extensión : $omitidosExt" }
if ($omitidosTam)     { Write-L "  Omitidos por tamaño    : $omitidosTam" }
if ($omitidosFecha)   { Write-L "  Omitidos por fecha     : $omitidosFecha" }
Write-L "======================================================" HEADER

if ($archivos.Count -eq 0) {
    Write-L "No hay archivos que cumplan los criterios." WARN
    exit 0
}

if ($Simular) {
    Write-L ""
    Write-L "Modo -Simular: no se descarga nada." OK
    $pathSim = Join-Path $carpetaJob "_SIMULACION.csv"
    $archivos | Select-Object RutaRelativa, @{n='MB';e={[math]::Round($_.Bytes/1MB,2)}}, Modificado |
        Export-Csv -LiteralPath $pathSim -NoTypeInformation -Encoding utf8BOM
    Write-L "Listado guardado en: $pathSim" OK
    exit 0
}

# ============================================================
# REANUDACIÓN
# ============================================================
$encabezadoLog    = "Ruta,Nombre,Tamano_MB,Modificado,Resultado,Volumen,Detalle,Timestamp"
$encabezadoEstado = "FileRef,Volumen,Timestamp"
$script:volNum    = 0
$script:completados = [System.Collections.Generic.HashSet[string]]::new()

if ($Reanudar) {
    # 1) Volúmenes existentes: los que no abren se apartan, la numeración sigue
    $volInvalidos = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $previosVol = @(Get-ChildItem -LiteralPath $carpetaJob -File -ErrorAction SilentlyContinue |
                    Where-Object { $_.Name -match '_vol\d+\.zip' })
    foreach ($v in $previosVol) {
        $script:volNum = [math]::Max($script:volNum, (Get-NumeroVolumen $v.Name))
        if ($v.Extension -ne ".zip") { continue }
        try {
            $z = [System.IO.Compression.ZipFile]::OpenRead($v.FullName)
            $z.Dispose()
        }
        catch {
            [void]$volInvalidos.Add($v.Name)
            Move-Item -LiteralPath $v.FullName -Destination "$($v.FullName).incompleto" -Force -ErrorAction SilentlyContinue
            Write-L "Volumen dañado (corte durante la escritura): $($v.Name) → se aparta como .incompleto y sus archivos se vuelven a bajar" WARN
        }
    }

    # 2) Estado: lo registrado en volúmenes válidos (o en carpeta) se saltea
    if (Test-Path -LiteralPath $pathEstado) {
        try {
            foreach ($p in (Import-Csv -LiteralPath $pathEstado -Encoding UTF8)) {
                if ($p.FileRef -and -not $volInvalidos.Contains([string]$p.Volumen)) {
                    [void]$script:completados.Add($p.FileRef.ToLowerInvariant())
                }
            }
            Write-L ""
            Write-L "Reanudando: $($script:completados.Count) archivo(s) ya guardados se saltean" OK
        }
        catch {
            Write-L "No se pudo leer el estado previo: $($_.Exception.Message)" WARN
        }
    }

    # Si hubo volúmenes dañados, el estado se reescribe sin sus entradas
    if ($volInvalidos.Count -gt 0 -and (Test-Path -LiteralPath $pathEstado)) {
        $validas = @(Import-Csv -LiteralPath $pathEstado -Encoding UTF8 |
                     Where-Object { -not $volInvalidos.Contains([string]$_.Volumen) })
        $tmpEstado = "$pathEstado.tmp"
        if ($validas.Count -gt 0) { $validas | Export-Csv -LiteralPath $tmpEstado -NoTypeInformation -Encoding utf8BOM }
        else { $encabezadoEstado | Out-File -LiteralPath $tmpEstado -Encoding utf8BOM }
        Move-Item -LiteralPath $tmpEstado -Destination $pathEstado -Force
    }
}
else {
    # Corrida nueva: se reinician los archivos de control y se borran volúmenes viejos
    $viejos = @(Get-ChildItem -LiteralPath $carpetaJob -File -ErrorAction SilentlyContinue |
                Where-Object { $_.Name -match '_vol\d+\.zip' })
    if ($viejos.Count -gt 0) {
        if (-not $SinConfirmacion) {
            Write-L "" WARN
            Write-L "Ya existen $($viejos.Count) volumen(es) de una corrida anterior." WARN
            Write-L "Sin -Reanudar se van a BORRAR y se descarga todo de nuevo." WARN
            $rta = Read-Host "Continuar y borrar? (SI/NO)"
            if ($rta.Trim().ToUpper() -ne "SI") { Write-L "Cancelado." WARN; exit 0 }
        }
        $viejos | Remove-Item -Force -ErrorAction SilentlyContinue
    }
    $encabezadoEstado | Out-File -LiteralPath $pathEstado -Encoding utf8BOM
    $encabezadoLog    | Out-File -LiteralPath $pathLog    -Encoding utf8BOM
}

if (-not (Test-Path -LiteralPath $pathEstado)) { $encabezadoEstado | Out-File -LiteralPath $pathEstado -Encoding utf8BOM }
if (-not (Test-Path -LiteralPath $pathLog))    { $encabezadoLog    | Out-File -LiteralPath $pathLog    -Encoding utf8BOM }

# ============================================================
# BUFFERS DE LOG
# ============================================================
$script:pathLog    = $pathLog
$script:pathEstado = $pathEstado
$script:bufLog     = [System.Collections.Generic.List[string]]::new()
$script:bufEstado  = [System.Collections.Generic.List[string]]::new()
$script:ultimoFlush = [datetime]::Now

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
    $f = $FileRef -replace '"', '""'
    $script:bufEstado.Add("`"$f`",`"$Volumen`",`"$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`"")
}

function Write-Buffer {
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][System.Collections.Generic.List[string]]$Buffer,
        [Parameter(Mandatory)][string]$Path
    )
    if ($Buffer.Count -eq 0) { return }
    # Reintento: si el CSV está abierto en Excel, el buffer espera al próximo flush
    for ($i = 1; $i -le 3; $i++) {
        try {
            Add-Content -LiteralPath $Path -Value $Buffer -Encoding utf8BOM -ErrorAction Stop
            $Buffer.Clear()
            return
        }
        catch { Start-Sleep -Milliseconds 300 }
    }
}

function Flush-Buffers {
    Write-Buffer -Buffer $script:bufLog    -Path $script:pathLog
    Write-Buffer -Buffer $script:bufEstado -Path $script:pathEstado
    $script:ultimoFlush = [datetime]::Now
}

# ============================================================
# CONTADORES
# ============================================================
$script:ok          = 0
$script:errores     = 0
$script:salteados   = 0
$script:bytesOk     = [long]0
$script:difTamano   = 0
$script:ultimoError = [System.Collections.Generic.Dictionary[string,string]]::new()

# ============================================================
# MANEJO DE VOLÚMENES ZIP
# Los archivos de un volumen se confirman en _ESTADO.csv recién cuando el
# volumen se cerró y validó: un corte nunca deja archivos "fantasma".
# ============================================================
$script:zip       = $null
$script:zipFs     = $null
$script:zipPath   = ""
$script:pendVol   = [System.Collections.Generic.List[object]]::new()
$limiteBytes      = [long]($TamanoVolumenGB * 1GB)

function Open-Volumen {
    $script:volNum++
    $nombre = "{0}_{1}_vol{2:d2}.zip" -f $nombreSitio, $nombreBib, $script:volNum
    $script:zipPath = Join-Path $carpetaJob $nombre
    $script:zipFs = [System.IO.File]::Open($script:zipPath, [System.IO.FileMode]::Create)
    $script:zip = [System.IO.Compression.ZipArchive]::new(
        $script:zipFs, [System.IO.Compression.ZipArchiveMode]::Create, $true)
    Write-L "  → Nuevo volumen: $nombre" OK
}

function Close-Volumen {
    if (-not $script:zip) { return }

    $nombre = Split-Path $script:zipPath -Leaf
    $valido = $false
    try {
        $script:zip.Dispose()
        $script:zipFs.Dispose()

        # Validación: se reabre y se cuentan las entradas. Un ZIP que no se
        # puede abrir acá tampoco lo va a poder abrir el cliente.
        $verif = [System.IO.Compression.ZipFile]::OpenRead($script:zipPath)
        $n = $verif.Entries.Count
        $verif.Dispose()
        $valido = ($n -eq $script:pendVol.Count)
        $gb = ConvertTo-Gb ([System.IO.FileInfo]::new($script:zipPath).Length)
        if ($valido -and $n -eq 0) { }
        elseif ($valido) { Write-L "  ✓ $nombre cerrado y validado — $n archivos, $gb GB" OK }
        else { Write-L "  [ERROR] $nombre tiene $n entradas y se esperaban $($script:pendVol.Count)" ERROR }
    }
    catch {
        Write-L "  [ERROR] $nombre no se pudo validar: $($_.Exception.Message)" ERROR
    }
    finally {
        $script:zip = $null
        $script:zipFs = $null
    }

    if ($valido -and $script:pendVol.Count -eq 0) {
        # Volumen vacío (todo lo de este tramo falló): no se deja un ZIP sin contenido
        Remove-Item -LiteralPath $script:zipPath -Force -ErrorAction SilentlyContinue
        $script:volNum--
    }
    elseif ($valido) {
        foreach ($a in $script:pendVol) {
            Add-Estado -FileRef $a.FileRef -Volumen $nombre
            [void]$script:completados.Add($a.Clave)
        }
    }
    else {
        Move-Item -LiteralPath $script:zipPath -Destination "$($script:zipPath).incompleto" -Force -ErrorAction SilentlyContinue
        foreach ($a in $script:pendVol) {
            $script:ok--
            $script:bytesOk -= $a.Bytes
            $script:errores++
            $script:ultimoError[$a.Clave] = "Volumen $nombre inválido"
            Add-Log -Ruta $a.RutaRelativa -Nombre $a.Nombre -Bytes $a.Bytes -Modificado $a.Modificado `
                    -Resultado "ERROR" -Volumen $nombre -Detalle "Volumen inválido: se vuelve a bajar con -Reanudar"
        }
    }
    $script:pendVol.Clear()
    Flush-Buffers
}

# ============================================================
# COLA DE DESCARGA
# ============================================================
$script:cola = [System.Collections.Generic.Queue[object]]::new()
$bytesPendientes = [long]0
$rutaMasLarga = 0

# De mayor a menor: los archivos grandes arrancan primero y no quedan
# solos al final (la "cola larga" que alarga la corrida)
foreach ($a in ($archivos | Sort-Object -Property Bytes -Descending)) {
    if ($script:completados.Contains($a.Clave)) {
        $script:salteados++
        Add-Log -Ruta $a.RutaRelativa -Nombre $a.Nombre -Bytes $a.Bytes -Modificado $a.Modificado `
                -Resultado "SALTADO" -Volumen "" -Detalle "Ya descargado en una corrida previa"
        continue
    }
    $script:cola.Enqueue($a)
    $bytesPendientes += $a.Bytes
    if ($SinComprimir) {
        $rutaMasLarga = [math]::Max($rutaMasLarga, ($carpetaOut.Length + 1 + $a.RutaRelativa.Length))
    }
}
Flush-Buffers

if ($script:salteados -gt 0) { Write-L "Pendientes: $($script:cola.Count) archivo(s) — $(ConvertTo-Gb $bytesPendientes) GB" OK }

# --- Rutas largas (solo archivos sueltos en Windows) ---
if ($SinComprimir -and $IsWindows -and $rutaMasLarga -ge 260) {
    $lp = 0
    try { $lp = (Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\FileSystem' -ErrorAction Stop).LongPathsEnabled } catch { }
    if ($lp -ne 1) {
        Write-L "Hay rutas de hasta $rutaMasLarga caracteres y Windows no tiene habilitadas las rutas largas." WARN
        Write-L "Esos archivos van a fallar. Opciones: usar un -CarpetaDestino más corto (ej: D:\B), quitar -SinComprimir" WARN
        Write-L "o habilitar LongPathsEnabled=1 (requiere admin)." WARN
        if (-not $SinConfirmacion) {
            $sigue = Read-Host "Continuar igual? (SI/NO)"
            if ($sigue.Trim().ToUpper() -ne "SI") { exit 0 }
        }
    }
}

# --- Espacio en disco ---
try {
    $unidad = [System.IO.DriveInfo]::new([System.IO.Path]::GetPathRoot([System.IO.Path]::GetFullPath($carpetaJob)))
    $libreGb = ConvertTo-Gb ([double]$unidad.AvailableFreeSpace)
    # +10% de margen, más los temporales en curso (hasta -Hilos archivos a la vez)
    $necesarioGb = [math]::Round((ConvertTo-Gb $bytesPendientes) * 1.1, 1)
    Write-L "Espacio libre en $($unidad.Name): $libreGb GB — necesario aprox: $necesarioGb GB" $(if ($libreGb -lt $necesarioGb) {"ERROR"} else {"INFO"})
    if ($libreGb -lt $necesarioGb) {
        Write-L "Espacio insuficiente en disco." ERROR
        if (-not $SinConfirmacion) {
            $sigue = Read-Host "Continuar igual? (SI/NO)"
            if ($sigue.Trim().ToUpper() -ne "SI") { exit 1 }
        }
    }
}
catch { }

# ============================================================
# TOKEN Y DESCARGADOR
# ============================================================
$script:descargador = [CircoKit.Descargador]::new($UserAgent, $Hilos)
$script:descargador.MaxReintentos    = $MaxReintentos
$script:descargador.SegundosSinDatos = $SegundosSinDatos

$script:tokenVence    = [datetime]::MinValue
$script:ultimoIntento = [datetime]::MinValue

function Update-Token {
    param([switch]$Forzar)
    $ahora = [datetime]::UtcNow
    if (-not $Forzar -and $ahora -lt $script:tokenVence.AddMinutes(-10)) { return $true }
    # Cerca del vencimiento se reintenta como mucho una vez por minuto
    if (-not $Forzar -and ($ahora - $script:ultimoIntento).TotalSeconds -lt 60) { return $true }
    $script:ultimoIntento = $ahora
    try {
        $t = [string](Get-PnPAccessToken -ResourceTypeName SharePoint -Connection $conn -ErrorAction Stop)
        if (-not $t) { throw "Token vacío" }
        $script:descargador.SetToken($t)
        $nuevoVence = Get-JwtVencimientoUtc $t
        if ($nuevoVence -gt $script:tokenVence -and $script:tokenVence -ne [datetime]::MinValue) {
            Write-L "  Token renovado — vence $($nuevoVence.ToLocalTime().ToString('HH:mm'))"
        }
        $script:tokenVence = $nuevoVence
        return $true
    }
    catch {
        Write-L "  No se pudo obtener/renovar el token: $($_.Exception.Message)" WARN
        return $false
    }
}

if (-not (Update-Token -Forzar)) {
    Write-L "Sin token no se puede descargar. Probar con -ForceAuth." ERROR
    $script:descargador.Dispose()
    exit 1
}

# ============================================================
# GUARDADO (hilo principal: ZIP o carpeta)
# ============================================================
function Save-Descargado {
    param($Arch, [string]$TmpFile)

    if ($SinComprimir) {
        $destino = [System.IO.Path]::Combine([string[]](@($carpetaOut) + $Arch.RutaRelativa.Split('/')))
        [void][System.IO.Directory]::CreateDirectory([System.IO.Path]::GetDirectoryName($destino))
        [System.IO.File]::Move($TmpFile, $destino, $true)
        Add-Log -Ruta $Arch.RutaRelativa -Nombre $Arch.Nombre -Bytes $Arch.Bytes -Modificado $Arch.Modificado `
                -Resultado "OK" -Volumen "(sin comprimir)" -Detalle $Arch.Nota
        Add-Estado -FileRef $Arch.FileRef -Volumen "(sin comprimir)"
        [void]$script:completados.Add($Arch.Clave)
        return
    }

    # ¿Entra en el volumen actual? Si no, se cierra y se abre el siguiente
    $tamArchivo = [System.IO.FileInfo]::new($TmpFile).Length
    if ($script:zipFs.Length -gt 0 -and ($script:zipFs.Length + $tamArchivo) -gt $limiteBytes) {
        Close-Volumen
        Open-Volumen
    }

    [void][System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile(
        $script:zip, $TmpFile, $Arch.RutaRelativa, (Get-NivelCompresion $Arch.Nombre))
    [System.IO.File]::Delete($TmpFile)

    $script:pendVol.Add($Arch)
    Add-Log -Ruta $Arch.RutaRelativa -Nombre $Arch.Nombre -Bytes $Arch.Bytes -Modificado $Arch.Modificado `
            -Resultado "OK" -Volumen (Split-Path $script:zipPath -Leaf) -Detalle $Arch.Nota
}

# ============================================================
# DESCARGA EN PARALELO
# ============================================================
$script:enCurso = [System.Collections.Generic.Dictionary[int,object]]::new()
$script:nextId  = 0

function Start-Descargas {
    while ($script:enCurso.Count -lt $Hilos -and $script:cola.Count -gt 0) {
        $a = $script:cola.Dequeue()
        $script:nextId++
        $tmp = Join-Path $carpetaTmp ([guid]::NewGuid().ToString("N") + ".tmp")
        $urlDescarga = [CircoKit.Descargador]::ArmarUrl($Url, $a.FileRef)
        $tarea = $script:descargador.DescargarAsync($script:nextId, $urlDescarga, $tmp, $a.ModificadoUtc)
        $script:enCurso[$script:nextId] = [PSCustomObject]@{ Tarea = $tarea; Arch = $a; Tmp = $tmp }
        if ($PausaMs -gt 0) { Start-Sleep -Milliseconds $PausaMs }
    }
}

function Complete-Descarga {
    param($Entrada)
    $a = $Entrada.Arch
    $r = $null
    $msg = ""
    try { $r = $Entrada.Tarea.Result } catch { $msg = $_.Exception.GetBaseException().Message }

    if ($r -and $r.Ok) {
        $nota = ""
        if ($a.Bytes -gt 0 -and $r.Bytes -ne $a.Bytes) {
            # Pasa con algunos Office (SharePoint reescribe metadatos): se guarda igual y se avisa
            $nota = "Tamaño distinto al informado por SharePoint ($($a.Bytes) vs $($r.Bytes) bytes)"
            $script:difTamano++
        }
        $a | Add-Member -NotePropertyName Nota -NotePropertyValue $nota -Force
        try {
            Save-Descargado -Arch $a -TmpFile $Entrada.Tmp
            $script:ok++
            $script:bytesOk += $r.Bytes
        }
        catch {
            $script:errores++
            $m = $_.Exception.Message -replace "`r`n|`n", " "
            $script:ultimoError[$a.Clave] = "Al guardar: $m"
            Add-Log -Ruta $a.RutaRelativa -Nombre $a.Nombre -Bytes $a.Bytes -Modificado $a.Modificado `
                    -Resultado "ERROR" -Volumen "" -Detalle "Al guardar: $m"
            try { [System.IO.File]::Delete($Entrada.Tmp) } catch { }
        }
        return
    }

    # 401: se renueva el token y el archivo vuelve a la cola (hasta 2 veces)
    if ($r -and $r.AuthVencida -and $a.ReintentosAuth -lt 2) {
        $a.ReintentosAuth++
        [void](Update-Token -Forzar)
        $script:cola.Enqueue($a)
        return
    }

    if ($r) { $msg = $r.Error }
    $script:errores++
    $script:ultimoError[$a.Clave] = $msg
    Add-Log -Ruta $a.RutaRelativa -Nombre $a.Nombre -Bytes $a.Bytes -Modificado $a.Modificado `
            -Resultado "ERROR" -Volumen "" -Detalle $msg
    try { [System.IO.File]::Delete($Entrada.Tmp) } catch { }
}

$sw          = [System.Diagnostics.Stopwatch]::StartNew()
$ultimaBarra = [datetime]::Now.AddSeconds(-60)
$ultimoTexto = [datetime]::Now
$totalCola   = $script:cola.Count
$terminoBien = $false
$script:detener = $false

# Ctrl+C controlado: en vez de cortar el script en cualquier punto (lo que puede
# dejar un ZIP a medio escribir), se lee como tecla y se hace un cierre ordenado.
$ctrlCPrevio = $null
$leerTeclado = $false
try {
    if (-not [Console]::IsInputRedirected) {
        $ctrlCPrevio = [Console]::TreatControlCAsInput
        [Console]::TreatControlCAsInput = $true
        $leerTeclado = $true
    }
} catch { $leerTeclado = $false }

function Test-PedidoDetener {
    if (-not $leerTeclado) { return }
    try {
        while ([Console]::KeyAvailable) {
            $k = [Console]::ReadKey($true)
            if ($k.Key -eq [ConsoleKey]::C -and ($k.Modifiers -band [ConsoleModifiers]::Control)) {
                $script:detener = $true
            }
        }
    } catch { }
}

function Show-Progreso {
    param([switch]$Forzar)
    $ahora = [datetime]::Now
    $seg   = [math]::Max(0.1, $sw.Elapsed.TotalSeconds)
    $vivos = [double]$script:descargador.BytesEnVivo
    $mbps  = ($vivos / 1MB) / $seg
    $hechos = $script:ok + $script:errores
    $pct   = if ($bytesPendientes -gt 0) { [math]::Min(100, [math]::Round(($vivos / $bytesPendientes) * 100, 1)) }
             else { [math]::Round(($hechos / [math]::Max(1, $totalCola)) * 100, 1) }
    $eta   = if ($mbps -gt 0) { (([double]$bytesPendientes - $vivos) / 1MB) / $mbps } else { 0 }

    $pausa = ($script:descargador.PausaHastaUtc - [datetime]::UtcNow).TotalSeconds
    $txtPausa = if ($pausa -gt 0) { " | THROTTLING: pausa $([int]$pausa)s" } else { "" }

    if ($Forzar -or ($ahora - $script:ultimaBarra).TotalMilliseconds -ge 700) {
        Write-Progress -Activity "Descargando $($lista.Title) — $($web.Title)" `
                       -Status ("{0}/{1} | {2} GB | {3:N1} MB/s | en curso: {4}{5}" -f $hechos, $totalCola, (ConvertTo-Gb $vivos), $mbps, $script:enCurso.Count, $txtPausa) `
                       -CurrentOperation "ETA $(Format-Duracion $eta) | errores: $($script:errores) | esperas por throttling: $($script:descargador.ThrottlesTotal)" `
                       -PercentComplete $pct
        $script:ultimaBarra = $ahora
    }

    if ($Forzar -or ($ahora - $script:ultimoTexto).TotalSeconds -ge 15) {
        Write-L ("  {0,6:N1}% — {1}/{2} | {3} GB | {4:N1} MB/s | errores: {5} | throttling: {6} | ETA {7}{8}" -f `
                 $pct, $hechos, $totalCola, (ConvertTo-Gb $vivos), $mbps, $script:errores,
                 $script:descargador.ThrottlesTotal, (Format-Duracion $eta), $txtPausa)
        $script:ultimoTexto = $ahora
    }
}

Write-L ""
Write-L "Iniciando descarga de $totalCola archivo(s) — $(ConvertTo-Gb $bytesPendientes) GB — $Hilos en paralelo" HEADER
if ($leerTeclado) { Write-L "Ctrl+C detiene la descarga de forma ordenada (cierra y valida el volumen en curso)." }

try {
    if (-not $SinComprimir -and $script:cola.Count -gt 0) { Open-Volumen }

    while ($script:cola.Count -gt 0 -or $script:enCurso.Count -gt 0) {

        Test-PedidoDetener
        if ($script:detener) {
            Write-L "" WARN
            Write-L "Ctrl+C recibido — se cancelan las descargas en curso y se cierra el volumen..." WARN
            $script:descargador.Cancelar()
            break
        }

        [void](Update-Token)
        Start-Descargas
        if ($script:enCurso.Count -eq 0) { break }

        $tareas = [System.Threading.Tasks.Task[]]@($script:enCurso.Values | ForEach-Object { $_.Tarea })
        [void][System.Threading.Tasks.Task]::WaitAny($tareas, 500)

        $listos = @($script:enCurso.Keys | Where-Object { $script:enCurso[$_].Tarea.IsCompleted })
        foreach ($id in $listos) {
            $e = $script:enCurso[$id]
            [void]$script:enCurso.Remove($id)
            Start-Descargas            # se rellena la ventana ANTES de comprimir
            Complete-Descarga -Entrada $e
        }

        if (([datetime]::Now - $script:ultimoFlush).TotalSeconds -ge 5) { Flush-Buffers }
        Show-Progreso
    }
    $terminoBien = -not $script:detener
}
finally {
    # Primero lo crítico (cerrar ZIP, volcar estado); los mensajes, después.
    if (-not $terminoBien) { $script:descargador.Cancelar() }
    if (-not $SinComprimir) { Close-Volumen }
    Flush-Buffers
    $script:descargador.Dispose()
    if ($leerTeclado) { try { [Console]::TreatControlCAsInput = $ctrlCPrevio } catch { } }
    try { Write-Progress -Activity "Descargando" -Completed } catch { }
    if (-not $terminoBien) {
        Write-L "Interrumpido. Para seguir: mismo comando con -Reanudar" WARN
    }
}

$sw.Stop()

# Limpieza de temporales
try { [System.IO.Directory]::Delete($carpetaTmp, $true) } catch { }

# ============================================================
# FALTANTES
# Se calculan contra lo efectivamente guardado, no contra el historial del
# log: un archivo que falló en una corrida y se bajó en otra no figura.
# ============================================================
$faltantes = @(
    foreach ($a in $archivos) {
        if ($script:completados.Contains($a.Clave)) { continue }
        $det = if ($script:ultimoError.ContainsKey($a.Clave)) { $script:ultimoError[$a.Clave] } else { "No procesado" }
        [PSCustomObject]@{
            Ruta       = $a.RutaRelativa
            Nombre     = $a.Nombre
            Tamano_MB  = [math]::Round($a.Bytes / 1MB, 3)
            Modificado = $a.Modificado
            Resultado  = "ERROR"
            Volumen    = ""
            Detalle    = $det
            Timestamp  = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        }
    }
)

$tmpFalt = "$pathFaltantes.tmp"
if ($faltantes.Count -gt 0) { $faltantes | Export-Csv -LiteralPath $tmpFalt -NoTypeInformation -Encoding utf8BOM }
else { $encabezadoLog | Out-File -LiteralPath $tmpFalt -Encoding utf8BOM }
Move-Item -LiteralPath $tmpFalt -Destination $pathFaltantes -Force

# ============================================================
# RESUMEN FINAL
# ============================================================
$seg = [math]::Max(0.1, $sw.Elapsed.TotalSeconds)
$volumenes = @(Get-ChildItem -LiteralPath $carpetaJob -File -ErrorAction SilentlyContinue | Where-Object { $_.Name -match '_vol\d+\.zip$' })
$tamZips = [long]0
foreach ($v in $volumenes) { $tamZips += $v.Length }
$mbpsMedia = [math]::Round(($script:bytesOk / 1MB) / $seg, 1)
$throttles = $script:descargador.ThrottlesTotal

Write-L ""
Write-L "======================================================" HEADER
Write-L "  RESUMEN FINAL"                                         HEADER
Write-L "  Sitio            : $($web.Title)"
Write-L "  Biblioteca       : $($lista.Title)"
Write-L "  Guardados (total): $($archivos.Count - $faltantes.Count) de $($archivos.Count)" OK
Write-L "  Esta corrida     : $($script:ok) OK | $($script:salteados) salteados"
Write-L "  Faltan           : $($faltantes.Count)" $(if ($faltantes.Count -gt 0) {"ERROR"} else {"OK"})
Write-L "  Descargado       : $(ConvertTo-Gb $script:bytesOk) GB en esta corrida" OK
if (-not $SinComprimir) {
    Write-L "  Volúmenes ZIP    : $($volumenes.Count) — $(ConvertTo-Gb $tamZips) GB en disco" OK
}
Write-L "  Duración         : $(Format-Duracion $seg)"
Write-L "  Velocidad media  : $mbpsMedia MB/s con $Hilos hilos"
Write-L "  Esperas throttling: $throttles"
if ($script:difTamano -gt 0) {
    Write-L "  Tamaño distinto  : $($script:difTamano) archivo(s) (detalle en _LOG.csv, columna Detalle)" WARN
}
Write-L "  Carpeta          : $carpetaJob"
Write-L "======================================================" HEADER

# Sugerencia de ajuste de hilos
if ($terminoBien -and $script:bytesOk -gt 1GB) {
    if ($throttles -eq 0 -and $Hilos -lt 16) {
        Write-L "Sin throttling: se puede probar con -Hilos $([math]::Min(32, $Hilos * 2)) si la velocidad está lejos del ancho de banda." INFO
    }
    elseif ($throttles -gt ($script:ok / 20) -and $Hilos -gt 2) {
        Write-L "Mucho throttling: probar con -Hilos $([math]::Max(2, [int]($Hilos / 2)))." WARN
    }
}

if ($faltantes.Count -gt 0) {
    Write-L "FALTARON $($faltantes.Count) archivo(s) — detalle en $(Split-Path $pathFaltantes -Leaf)" ERROR
    Write-L "Para reintentar solo los faltantes:"
    Write-L "  .\Script_C_DescargarBiblioteca.ps1 -ConfigPath `"$ConfigPath`" -Url `"$Url`" -Biblioteca `"$($lista.Title)`" -Reanudar"

    Write-L "" WARN
    Write-L "Primeros faltantes:" WARN
    foreach ($f in ($faltantes | Select-Object -First 5)) {
        Write-L "  $($f.Ruta)" ERROR
        Write-L "      $($f.Detalle)" ERROR
    }
}
elseif ($terminoBien) {
    Write-L "Se descargaron todos los archivos. No faltó ninguno." OK
    Write-L "El archivo _ESTADO.csv se puede borrar." INFO
}

Write-L "Descarga finalizada." OK