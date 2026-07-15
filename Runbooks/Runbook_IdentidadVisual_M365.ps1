# =============================================================================
# Runbook_IdentidadVisual_M365.ps1  |  Circo Studio  |  Kit de Curaduría M365
# =============================================================================
#
# PROPÓSITO:
#   Configura la identidad visual corporativa del tenant SharePoint Online:
#   theme de color, logo, imagen hero y navegación global del Hub Site.
#
# MODO DE USO:
#   Este script NO se ejecuta completo de una sola vez.
#   Ejecutar cada sección por separado seleccionando el bloque y presionando F8.
#   Verificar la salida en consola antes de pasar a la siguiente sección.
#
#   Secuencia recomendada:
#     1. Completar todas las variables de la Sección 0
#     2. Ejecutar Sección 1 (verificar módulos) — solo la primera vez por máquina
#     3. Ejecutar Sección 2 (registrar el theme)
#     4. Ejecutar Sección 3 (subir branding y aplicar al Hub)
#     5. Validar visualmente en el browser antes de continuar
#     6. Ejecutar Sección 4 (navegación global)
#     7. Ejecutar Sección 5 una vez validado el Hub (spokes — descomentear)
#     8. Ejecutar Sección 6 (verificación final)
#
# PREREQUISITOS:
#   - Cuenta con rol SharePoint Administrator en el tenant
#   - PnP.PowerShell instalado (Sección 1 lo verifica e instala si falta)
#   - Microsoft.Online.SharePoint.PowerShell instalado (ídem)
#   - Archivos de logo e imagen hero disponibles localmente
#
# PASOS MANUALES (no automatizables por API):
#   - Configurar el Header del Hub Site (layout Extended + imagen hero):
#     Hub Site → Settings → Change the look → Header
#     Layout: Extended | Background image: la URL del hero subido en Sección 3
#
# =============================================================================


# =============================================================================
# SECCIÓN 0 — VARIABLES GLOBALES
# Completar todos estos valores antes de ejecutar cualquier sección.
# =============================================================================

# ── Tenant ───────────────────────────────────────────────────────────────────
# URL del Admin Center del tenant
$adminUrl     = "https://[TENANT]-admin.sharepoint.com"   # ← reemplazar [TENANT]

# URL raíz del Hub Site principal
# Generalmente es la raíz del tenant o un sitio dedicado (/sites/[NombreHub])
$hubUrl       = "https://[TENANT].sharepoint.com"          # ← reemplazar

# ── Theme ────────────────────────────────────────────────────────────────────
# Nombre del theme tal como va a aparecer en el tenant
# Usar un nombre que identifique al cliente, ej: "Acme Corporativo"
$themeName    = "[CLIENTE] Corporativo"                    # ← reemplazar

# ── Archivos de branding ─────────────────────────────────────────────────────
# Paths locales en la máquina donde se ejecuta el script
# El logo debe ser PNG con fondo transparente, preferentemente 200x50px o similar
# El hero debe ser JPG/PNG, mínimo 1920x300px para que se vea bien en header Extended
$logoPath     = "C:\Branding\logo_blanco.png"              # ← ajustar path
$heroPath     = "C:\Branding\imagen_hero.jpg"              # ← ajustar path

# URLs donde van a quedar los archivos dentro de SharePoint
# No modificar a menos que el Hub Site esté en /sites/[OtroNombre]
$logoUrl      = "/sites/[HUB-SITE-PATH]/SiteAssets/logo.png"   # ← ajustar path SPO
$heroUrl      = "/sites/[HUB-SITE-PATH]/SiteAssets/hero.jpg"   # ← ajustar path SPO

# ── Sitios spoke ─────────────────────────────────────────────────────────────
# Lista de sitios que van a recibir el theme en la Sección 5
# Agregar o quitar según la arquitectura del cliente
$spokeSites   = @(
    "https://[TENANT].sharepoint.com/sites/[Spoke1]"      # ← completar
    "https://[TENANT].sharepoint.com/sites/[Spoke2]"      # ← completar
    # Agregar más sitios según sea necesario
)

# ── Navegación del Hub ───────────────────────────────────────────────────────
# Ítems de la barra de navegación global del Hub Site
# Title: texto que aparece en el menú | Url: URL de destino (relativa o absoluta)
# Mantener entre 5 y 8 ítems para que el menú no se sature
$navItems     = @(
    @{ Title = "Inicio";      Url = "/" }                  # ← ajustar URLs
    @{ Title = "[Sección 1]"; Url = "/sites/[Spoke1]" }   # ← completar
    @{ Title = "[Sección 2]"; Url = "/sites/[Spoke2]" }   # ← completar
    @{ Title = "[Sección 3]"; Url = "/sites/[Spoke3]" }   # ← completar
    # Agregar o quitar ítems según la arquitectura del cliente
)


# =============================================================================
# SECCIÓN 0b — PALETA DE COLORES
# Completar los valores hex del cliente.
# Todos los campos son obligatorios para que el theme quede completo.
#
# Guía rápida:
#   themePrimary     → color principal de la marca (botones, links, highlights)
#   themeSecondary   → variante más oscura del primario
#   accent           → color de acento para llamadas a la acción
#   Los rangos "Lighter/Light/Dark/Darker" se calculan a partir del primario.
#   Los neutrales definen la escala de grises de la UI.
#
# Herramienta recomendada para generar la paleta completa a partir de un hex:
#   https://fabricweb.z5.web.core.windows.net/fabric-website/tools/theme-designer
# =============================================================================

$themePalette = @{

    # Primarios — definir a partir del color principal de la marca del cliente
    "themePrimary"         = "#[HEX]"    # ← color principal (ej: azul navy, verde corporativo)
    "themeLighterAlt"      = "#[HEX]"    # ← tinte muy claro (fondo hover ligero)
    "themeLighter"         = "#[HEX]"    # ← tinte claro
    "themeLight"           = "#[HEX]"    # ← tinte medio-claro
    "themeTertiary"        = "#[HEX]"    # ← tono medio
    "themeSecondary"       = "#[HEX]"    # ← tono estándar (links, íconos)
    "themeDarkAlt"         = "#[HEX]"    # ← variante oscura (hover de botones)
    "themeDark"            = "#[HEX]"    # ← oscuro (texto sobre fondo claro)
    "themeDarker"          = "#[HEX]"    # ← muy oscuro

    # Acento — color secundario de la marca para CTAs y highlights
    "accent"               = "#[HEX]"    # ← color de acento (ej: cyan, naranja, dorado)

    # Neutrales — escala de grises de la interfaz
    # En la mayoría de los casos estos valores estándar funcionan bien.
    # Solo modificar si el cliente tiene un sistema de grises warm/cool específico.
    "neutralLighterAlt"    = "#F8F8F8"
    "neutralLighter"       = "#F4F4F4"
    "neutralLight"         = "#EAEAEA"
    "neutralQuaternaryAlt" = "#DADADA"
    "neutralQuaternary"    = "#D0D0D0"
    "neutralTertiaryAlt"   = "#C8C8C8"
    "neutralTertiary"      = "#A6A6A6"
    "neutralSecondary"     = "#666666"
    "neutralPrimaryAlt"    = "#3C3C3C"
    "neutralPrimary"       = "#2D2D2D"
    "neutralDark"          = "#212121"
    "black"                = "#1A1A1A"
    "white"                = "#FFFFFF"
}


# =============================================================================
# SECCIÓN 1 — VERIFICAR E INSTALAR MÓDULOS
# Ejecutar una sola vez por máquina.
# =============================================================================

Write-Host "`n[1/6] Verificando módulos..." -ForegroundColor Cyan

$modulosSPO = Get-InstalledModule -Name "Microsoft.Online.SharePoint.PowerShell" -ErrorAction SilentlyContinue
$modulosPnP = Get-InstalledModule -Name "PnP.PowerShell" -ErrorAction SilentlyContinue

if (-not $modulosSPO) {
    Write-Host "  Instalando Microsoft.Online.SharePoint.PowerShell..." -ForegroundColor Yellow
    Install-Module -Name "Microsoft.Online.SharePoint.PowerShell" -Force -AllowClobber
} else {
    Write-Host "  ✓ SPO Management Shell: v$($modulosSPO.Version)" -ForegroundColor Green
}

if (-not $modulosPnP) {
    Write-Host "  Instalando PnP.PowerShell..." -ForegroundColor Yellow
    Install-Module -Name "PnP.PowerShell" -Force -AllowClobber -Scope CurrentUser
} else {
    Write-Host "  ✓ PnP.PowerShell: v$($modulosPnP.Version)" -ForegroundColor Green
}

Write-Host "  Módulos listos.`n" -ForegroundColor Green


# =============================================================================
# SECCIÓN 2 — REGISTRAR EL THEME EN EL TENANT
# Requiere: Connect-SPOService (rol SharePoint Administrator)
# =============================================================================

Write-Host "`n[2/6] Registrando theme '$themeName' en el tenant..." -ForegroundColor Cyan

# Verificar que la paleta está completa antes de proceder
$camposFaltantes = $themePalette.GetEnumerator() | Where-Object { $_.Value -like "*[HEX]*" }
if ($camposFaltantes) {
    Write-Host "  [ERROR] Hay colores sin completar en la paleta:" -ForegroundColor Red
    $camposFaltantes | ForEach-Object {
        Write-Host "    → $($_.Key): $($_.Value)" -ForegroundColor Red
    }
    Write-Host "  Completar todos los colores en la Sección 0b antes de continuar." -ForegroundColor Yellow
    return
}

Import-Module "Microsoft.Online.SharePoint.PowerShell" -WarningAction SilentlyContinue

Write-Host "  Conectando al Admin Center: $adminUrl ..." -ForegroundColor Gray
Connect-SPOService -Url $adminUrl -ModernAuth $true

# -Overwrite permite re-ejecutar sin error si el theme ya existe
Add-SPOTheme -Name $themeName -Palette $themePalette -IsInverted $false -Overwrite

# Verificar que quedó registrado
$themeVerif = Get-SPOTheme -Name $themeName -ErrorAction SilentlyContinue
if ($themeVerif) {
    Write-Host "  ✓ Theme '$themeName' registrado correctamente en el tenant.`n" -ForegroundColor Green
} else {
    Write-Host "  ✗ No se pudo verificar el registro del theme. Revisar permisos de SharePoint Administrator.`n" -ForegroundColor Red
}


# =============================================================================
# SECCIÓN 3 — APLICAR THEME AL HUB Y SUBIR BRANDING
# Requiere: Connect-PnPOnline al Hub Site
# =============================================================================

Write-Host "`n[3/6] Aplicando theme y subiendo branding al Hub Site..." -ForegroundColor Cyan

Import-Module "PnP.PowerShell" -WarningAction SilentlyContinue

Write-Host "  Conectando al Hub Site: $hubUrl ..." -ForegroundColor Gray
Connect-PnPOnline -Url $hubUrl -Interactive

# Aplicar el theme registrado en la Sección 2
Set-PnPWebTheme -Theme $themeName
Write-Host "  ✓ Theme '$themeName' aplicado al Hub Site." -ForegroundColor Green

# Subir logo
if (Test-Path $logoPath) {
    Add-PnPFile -Path $logoPath -Folder "SiteAssets"
    Set-PnPWeb -SiteLogoUrl $logoUrl
    Write-Host "  ✓ Logo subido y aplicado: $logoUrl" -ForegroundColor Green
} else {
    Write-Host "  ⚠ Logo no encontrado en: $logoPath" -ForegroundColor Yellow
    Write-Host "    Ajustar la variable `$logoPath en la Sección 0 y volver a ejecutar esta sección." -ForegroundColor Yellow
}

# Subir imagen hero
if (Test-Path $heroPath) {
    Add-PnPFile -Path $heroPath -Folder "SiteAssets"
    Write-Host "  ✓ Imagen hero subida: $heroUrl" -ForegroundColor Green
} else {
    Write-Host "  ⚠ Imagen hero no encontrada en: $heroPath" -ForegroundColor Yellow
    Write-Host "    Ajustar la variable `$heroPath en la Sección 0 y volver a ejecutar esta sección." -ForegroundColor Yellow
}

Write-Host ""
Write-Host "  ┌─────────────────────────────────────────────────────────┐" -ForegroundColor Yellow
Write-Host "  │  PASO MANUAL REQUERIDO — no automatizable por API       │" -ForegroundColor Yellow
Write-Host "  │                                                          │" -ForegroundColor Yellow
Write-Host "  │  Ir al Hub Site en el browser:                          │" -ForegroundColor Yellow
Write-Host "  │  Settings → Change the look → Header                   │" -ForegroundColor Yellow
Write-Host "  │  Layout: Extended                                        │" -ForegroundColor Yellow
Write-Host "  │  Background image: $heroUrl" -ForegroundColor Yellow
Write-Host "  └─────────────────────────────────────────────────────────┘" -ForegroundColor Yellow
Write-Host ""


# =============================================================================
# SECCIÓN 4 — CONFIGURAR NAVEGACIÓN GLOBAL DEL HUB
# Requiere: Connect-PnPOnline al Hub Site (ya conectado desde Sección 3)
# PREREQUISITO: Validar visualmente el Hub en el browser antes de ejecutar.
# =============================================================================

Write-Host "`n[4/6] Configurando navegación global del Hub..." -ForegroundColor Cyan

# Verificar que los ítems de navegación están completos
$navFaltantes = $navItems | Where-Object { $_.Title -like "*[*]*" -or $_.Url -like "*[*]*" }
if ($navFaltantes) {
    Write-Host "  [AVISO] Hay ítems de navegación con valores placeholder:" -ForegroundColor Yellow
    $navFaltantes | ForEach-Object {
        Write-Host "    → '$($_.Title)' → $($_.Url)" -ForegroundColor Yellow
    }
    Write-Host "  Completar en la Sección 0 antes de ejecutar esta sección." -ForegroundColor Yellow
    return
}

# Limpiar navegación existente
Write-Host "  Limpiando navegación anterior..." -ForegroundColor Gray
Get-PnPNavigationNode -Location TopNavigationBar |
    ForEach-Object { Remove-PnPNavigationNode -Identity $_.Id -Force }

# Agregar nuevos ítems
foreach ($item in $navItems) {
    Add-PnPNavigationNode -Title $item.Title -Url $item.Url -Location TopNavigationBar
    Write-Host "  ✓ $($item.Title) → $($item.Url)" -ForegroundColor Green
}

Write-Host "  ✓ Navegación global configurada ($($navItems.Count) ítems).`n" -ForegroundColor Green


# =============================================================================
# SECCIÓN 5 — APLICAR THEME A SITIOS SPOKE
# Descomentear el bloque cuando el Hub esté validado visualmente.
# PREREQUISITO: Theme registrado (Sección 2) y Hub validado en browser.
# =============================================================================

Write-Host "`n[5/6] Aplicar theme a sitios spoke — DESCOMENTEAR para ejecutar" -ForegroundColor Cyan
Write-Host "  (Sección comentada por defecto — ejecutar solo después de validar el Hub)" -ForegroundColor Gray

<#
$erroresSPO = @()

foreach ($url in $spokeSites) {
    try {
        Connect-PnPOnline -Url $url -Interactive
        Set-PnPWebTheme -Theme $themeName
        Write-Host "  ✓ $url" -ForegroundColor Green
    }
    catch {
        Write-Host "  ✗ Error en $url : $($_.Exception.Message)" -ForegroundColor Red
        $erroresSPO += $url
    }
}

if ($erroresSPO.Count -gt 0) {
    Write-Host "`n  Sitios con error ($($erroresSPO.Count)):" -ForegroundColor Yellow
    $erroresSPO | ForEach-Object { Write-Host "    - $_" -ForegroundColor Red }
} else {
    Write-Host "`n  ✓ Theme aplicado a todos los sitios spoke sin errores." -ForegroundColor Green
}
#>


# =============================================================================
# SECCIÓN 6 — VERIFICACIÓN FINAL
# Lee el estado actual desde el tenant para confirmar que todo quedó aplicado.
# =============================================================================

Write-Host "`n[6/6] Verificación final..." -ForegroundColor Cyan

# Verificar theme en el tenant
Import-Module "Microsoft.Online.SharePoint.PowerShell" -WarningAction SilentlyContinue
Connect-SPOService -Url $adminUrl -ModernAuth $true

$themeCheck = Get-SPOTheme -Name $themeName -ErrorAction SilentlyContinue
if ($themeCheck) {
    Write-Host "  ✓ Theme '$themeName' presente en el tenant." -ForegroundColor Green
} else {
    Write-Host "  ✗ Theme '$themeName' NO encontrado en el tenant." -ForegroundColor Red
}

# Verificar logo del Hub
Import-Module "PnP.PowerShell" -WarningAction SilentlyContinue
Connect-PnPOnline -Url $hubUrl -Interactive
$web = Get-PnPWeb -Includes SiteLogoUrl
if ($web.SiteLogoUrl) {
    Write-Host "  ✓ Logo del Hub: $($web.SiteLogoUrl)" -ForegroundColor Green
} else {
    Write-Host "  ⚠ El Hub no tiene logo asignado." -ForegroundColor Yellow
}

# Verificar navegación
$navNodes = Get-PnPNavigationNode -Location TopNavigationBar
Write-Host "  ✓ Navegación global: $(@($navNodes).Count) ítems" -ForegroundColor Green
$navNodes | ForEach-Object {
    Write-Host "    → $($_.Title) ($($_.Url))" -ForegroundColor Gray
}

# Resumen final
Write-Host ""
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkGray
Write-Host "  Identidad visual M365 — Estado final" -ForegroundColor White
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor DarkGray
Write-Host "  Theme registrado : $themeName"      -ForegroundColor Gray
Write-Host "  Hub Site         : $hubUrl"          -ForegroundColor Gray
Write-Host "  Spokes           : $($spokeSites.Count) sitios configurados" -ForegroundColor Gray
Write-Host ""
Write-Host "  Pendiente manual :" -ForegroundColor Yellow
Write-Host "  → Configurar Header del Hub: Settings → Change the look → Header" -ForegroundColor Yellow
Write-Host "    Layout: Extended | Background image: $heroUrl" -ForegroundColor Yellow
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`n" -ForegroundColor DarkGray
