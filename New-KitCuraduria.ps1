#Requires -Version 5.1
<#
===============================================================================
  New-KitCuraduria.ps1  |  Circo Studio  |  Kit de Curaduria Informacional M365
===============================================================================

PROPOSITO:
  Crea la estructura de carpetas del Kit de Curaduria M365 en la ruta indicada
  y opcionalmente copia los archivos desde una carpeta fuente.

USO:
  # Solo estructura vacia:
  .\New-KitCuraduria.ps1 -Destino "C:\CuraduriaM365\Kit"

  # Estructura + copiar archivos desde fuente:
  .\New-KitCuraduria.ps1 -Destino "C:\CuraduriaM365\Kit" -Fuente "D:\Descargas\Kit"

  # Con carpeta de trabajo para un cliente especifico:
  .\New-KitCuraduria.ps1 -Destino "C:\CuraduriaM365\Kit" -Cliente "ACME"

  # Todo junto:
  .\New-KitCuraduria.ps1 -Destino "C:\CuraduriaM365\Kit" -Fuente "D:\Kit" -Cliente "ACME" -Force

ESTRUCTURA GENERADA:
  [Destino]\
  |- Config\
  |    |- kit_config_editor.html
  |    `- cliente_config.json         <- vacio, listo para completar
  |- PowerShell\
  |    |- M365Curaduria.psm1
  |    |- Script_A_Inventario.ps1
  |    |- Script_A_AsignarOwner.ps1
  |    |- Script_B_Masivo.ps1
  |    `- Script_C_Refresh.ps1
  |- Runbooks\
  |    `- Runbook_IdentidadVisual_M365.ps1
  |- Python\
  |    |- m365_curaduria.py
  |    |- analizar_contenido.py
  |    |- generar_informes_sitio.py
  |    `- generar_indice.py
  `- Posicionamiento\
       |- CS_Curaduria_Index.html
       |- CS_Curaduria_Por_Que.html
       |- CS_Curaduria_Como.html
       |- CS_Curaduria_Kit.html
       `- CS_Script_B_Documentacion_IT.html

  Si se pasa -Cliente "Nombre":
  [Destino]\Clientes\[Nombre]\
  |- Relevamiento\
  |    `- Informes\
  `- Config\
       `- cliente_config.json         <- preconfigurado con carpeta_local

===============================================================================
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory, HelpMessage="Ruta donde crear el kit")]
    [string]$Destino,

    [Parameter(HelpMessage="Ruta de origen de los archivos del kit")]
    [string]$Fuente = "",

    [Parameter(HelpMessage="Nombre del cliente, crea carpeta de trabajo adicional")]
    [string]$Cliente = "",

    [switch]$Force
)

# ============================================================
# HELPERS
# ============================================================
function Write-OK   { param([string]$msg) Write-Host "  OK  $msg" -ForegroundColor Green }
function Write-Skip { param([string]$msg) Write-Host "  --  $msg" -ForegroundColor DarkGray }
function Write-Warn { param([string]$msg) Write-Host "  !!  $msg" -ForegroundColor Yellow }
function Write-Err  { param([string]$msg) Write-Host "  XX  $msg" -ForegroundColor Red }

function New-KitFolder {
    param([string]$Path)
    if (Test-Path $Path) {
        Write-Skip "Ya existe: $Path"
    } else {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
        Write-OK "Carpeta creada: $Path"
    }
}

function Copy-KitFile {
    param([string]$Src, [string]$Dst)
    if (-not (Test-Path $Src)) {
        Write-Warn "No encontrado en fuente: $(Split-Path $Src -Leaf)"
        return $false
    }
    if ((Test-Path $Dst) -and (-not $Force)) {
        Write-Skip "Ya existe (usar -Force para sobreescribir): $(Split-Path $Dst -Leaf)"
        return $true
    }
    Copy-Item -Path $Src -Destination $Dst -Force
    Write-OK "Copiado: $(Split-Path $Dst -Leaf)"
    return $true
}

function New-ConfigJson {
    param(
        [string]$Path,
        [string]$NombreCliente = "",
        [string]$CarpetaLocal  = ""
    )
    if ((Test-Path $Path) -and (-not $Force)) {
        Write-Skip "Config ya existe: $(Split-Path $Path -Leaf)"
        return
    }
    $cfg = [ordered]@{
        cliente = [ordered]@{
            nombre         = $NombreCliente
            tenant_prefix  = ""
            sharepoint_url = ""
            admin_url      = ""
            client_id      = ""
            upn_circo      = ""
            carpeta_local  = $CarpetaLocal
        }
        fechas = [ordered]@{
            migracion         = (Get-Date -Format "yyyy-MM-dd")
            inactividad_anios = 2
        }
        refresh = [ordered]@{
            dias_umbral = 7
        }
        clasificacion = [ordered]@{
            gb_principal      = 10
            gb_mediano        = 1
            mb_archivo_grande = 100
        }
        sitios_excluidos      = @()
        canales_excluidos     = @()
        informes_prioritarios = @()
        sitios_legacy         = [ordered]@{
            label     = "Legacy (Teams)"
            prefijos  = @()
            templates = @("GROUP#0","TEAMCHANNEL#0","GROUP#1","TEAMCHANNEL#1")
        }
    }
    $cfg | ConvertTo-Json -Depth 5 | Out-File -FilePath $Path -Encoding UTF8
    Write-OK "Config JSON creado: $(Split-Path $Path -Leaf)"
}

function Show-Tree {
    param([string]$Root)
    $items = Get-ChildItem -Path $Root -Recurse | Sort-Object FullName
    foreach ($item in $items) {
        $rel   = $item.FullName.Replace($Root, "").TrimStart("\")
        $depth = ($rel -split "\\").Count - 1
        $pad   = "  " + ("  " * $depth)
        if ($item.PSIsContainer) {
            Write-Host "${pad}$($item.Name)\" -ForegroundColor Cyan
        } else {
            Write-Host "${pad}$($item.Name)" -ForegroundColor DarkGray
        }
    }
}


# ============================================================
# MAPA DE ARCHIVOS DEL KIT
# ============================================================
$archivosKit = @(
    [pscustomobject]@{ Carpeta = "Config";          Archivo = "kit_config_editor.html" }
    [pscustomobject]@{ Carpeta = "PowerShell";      Archivo = "M365Curaduria.psm1" }
    [pscustomobject]@{ Carpeta = "PowerShell";      Archivo = "Script_A_Inventario.ps1" }
    [pscustomobject]@{ Carpeta = "PowerShell";      Archivo = "Script_A_AsignarOwner.ps1" }
    [pscustomobject]@{ Carpeta = "PowerShell";      Archivo = "Script_B_Masivo.ps1" }
    [pscustomobject]@{ Carpeta = "PowerShell";      Archivo = "Script_C_Refresh.ps1" }
    [pscustomobject]@{ Carpeta = "Runbooks";        Archivo = "Runbook_IdentidadVisual_M365.ps1" }
    [pscustomobject]@{ Carpeta = "Python";          Archivo = "m365_curaduria.py" }
    [pscustomobject]@{ Carpeta = "Python";          Archivo = "analizar_contenido.py" }
    [pscustomobject]@{ Carpeta = "Python";          Archivo = "generar_informes_sitio.py" }
    [pscustomobject]@{ Carpeta = "Python";          Archivo = "generar_indice.py" }
    [pscustomobject]@{ Carpeta = "Posicionamiento"; Archivo = "CS_Curaduria_Index.html" }
    [pscustomobject]@{ Carpeta = "Posicionamiento"; Archivo = "CS_Curaduria_Por_Que.html" }
    [pscustomobject]@{ Carpeta = "Posicionamiento"; Archivo = "CS_Curaduria_Como.html" }
    [pscustomobject]@{ Carpeta = "Posicionamiento"; Archivo = "CS_Curaduria_Kit.html" }
    [pscustomobject]@{ Carpeta = "Posicionamiento"; Archivo = "CS_Script_B_Documentacion_IT.html" }
    [pscustomobject]@{ Carpeta = "Posicionamiento"; Archivo = "CS_AppRegistration_Guia.html" }
)


# ============================================================
# INICIO
# ============================================================
Write-Host ""
Write-Host "=================================================" -ForegroundColor DarkGray
Write-Host "  Circo Studio  |  Kit de Curaduria M365"        -ForegroundColor White
Write-Host "  Estructura de carpetas"                         -ForegroundColor DarkGray
Write-Host "=================================================" -ForegroundColor DarkGray
Write-Host "  Destino : $Destino"
if ($Fuente)  { Write-Host "  Fuente  : $Fuente" }
if ($Cliente) { Write-Host "  Cliente : $Cliente" }
Write-Host ""


# ============================================================
# PASO 1 -- CARPETAS DEL KIT
# ============================================================
Write-Host "[ 1/3 ]  Carpetas del kit" -ForegroundColor White

$carpetasKit = @(
    $Destino
    "$Destino\Config"
    "$Destino\PowerShell"
    "$Destino\Runbooks"
    "$Destino\Python"
    "$Destino\Posicionamiento"
)
foreach ($carpeta in $carpetasKit) {
    New-KitFolder $carpeta
}


# ============================================================
# PASO 2 -- ARCHIVOS DEL KIT
# ============================================================
Write-Host ""
Write-Host "[ 2/3 ]  Archivos del kit" -ForegroundColor White

$copiados = 0
$faltantes = 0

if ($Fuente) {
    if (-not (Test-Path $Fuente)) {
        Write-Err "La carpeta fuente no existe: $Fuente"
        Write-Warn "Las carpetas fueron creadas. Copiar los archivos manualmente."
    } else {
        foreach ($item in $archivosKit) {
            $srcConCarpeta = Join-Path $Fuente "$($item.Carpeta)\$($item.Archivo)"
            $srcRaiz       = Join-Path $Fuente $item.Archivo
            $dst           = Join-Path $Destino "$($item.Carpeta)\$($item.Archivo)"

            if (Test-Path $srcConCarpeta) {
                $src = $srcConCarpeta
            } elseif (Test-Path $srcRaiz) {
                $src = $srcRaiz
            } else {
                $src = $srcConCarpeta
            }

            $ok = Copy-KitFile -Src $src -Dst $dst
            if ($ok) { $copiados++ } else { $faltantes++ }
        }
    }
} else {
    Write-Warn "No se especifico -Fuente. Carpetas creadas vacias."
    Write-Host ""
    Write-Host "  Archivos a colocar por carpeta:" -ForegroundColor DarkGray
    $archivosKit | Group-Object Carpeta | ForEach-Object {
        Write-Host ""
        Write-Host "  $($_.Name)\" -ForegroundColor Cyan
        $_.Group | ForEach-Object {
            Write-Host "    $($_.Archivo)" -ForegroundColor DarkGray
        }
    }
}

New-ConfigJson -Path "$Destino\Config\cliente_config.json"


# ============================================================
# PASO 3 -- CARPETA DE TRABAJO DEL CLIENTE
# ============================================================
Write-Host ""
Write-Host "[ 3/3 ]  Carpeta de trabajo del cliente" -ForegroundColor White

if ($Cliente) {
    $nombreLimpio = $Cliente -replace "[^a-zA-Z0-9_\-]", "_" -replace "__+", "_"
    $raizCliente  = "$Destino\Clientes\$nombreLimpio"

    $carpetasCliente = @(
        "$Destino\Clientes"
        $raizCliente
        "$raizCliente\Relevamiento"
        "$raizCliente\Relevamiento\Informes"
        "$raizCliente\Config"
    )
    foreach ($carpeta in $carpetasCliente) {
        New-KitFolder $carpeta
    }

    New-ConfigJson `
        -Path          "$raizCliente\Config\cliente_config.json" `
        -NombreCliente $Cliente `
        -CarpetaLocal  "$raizCliente\Relevamiento"

    Write-Host ""
    Write-Host "  Proximo paso: abrir el editor HTML y completar la config" -ForegroundColor DarkGray
    Write-Host "  $Destino\Config\kit_config_editor.html" -ForegroundColor Cyan
    Write-Host "  Guardar el JSON resultante en:" -ForegroundColor DarkGray
    Write-Host "  $raizCliente\Config\cliente_config.json" -ForegroundColor Cyan

} else {
    Write-Skip "No se especifico -Cliente."
    Write-Host "  Para crear la carpeta de un cliente:" -ForegroundColor DarkGray
    Write-Host "  .\New-KitCuraduria.ps1 -Destino '$Destino' -Cliente 'NombreCliente'" -ForegroundColor DarkGray
}


# ============================================================
# RESUMEN FINAL
# ============================================================
Write-Host ""
Write-Host "=================================================" -ForegroundColor DarkGray
Write-Host "  Estructura final"                               -ForegroundColor White
Write-Host "=================================================" -ForegroundColor DarkGray
Write-Host "  $Destino\" -ForegroundColor White
Show-Tree -Root $Destino
Write-Host ""

if ($copiados -gt 0)  { Write-OK   "$copiados archivos copiados" }
if ($faltantes -gt 0) { Write-Warn "$faltantes archivos no encontrados en la fuente" }

Write-Host ""
Write-Host "  Kit listo en: $Destino" -ForegroundColor White
Write-Host "=================================================" -ForegroundColor DarkGray
Write-Host ""
