"""
===============================================================================
m365_curaduria.py  |  Circo Studio  |  Kit de Curaduría Informacional M365
===============================================================================
Módulo compartido para los tres scripts Python del kit.

Importar desde cada script con:
    from m365_curaduria import load_config, load_csvs, clean_df, CSS, ...

Funciones exportadas:
    load_config(config_path)         — carga y valida cliente_config.json
    load_csvs(config)                — carga todos los CSVs de relevamiento
    clean_df(df, config)             — limpia y enriquece el DataFrame
    get_carpeta_informes(config)     — path de la carpeta de salida de informes
    rec_info(post_mig, pct_pm)       — recomendación por sitio
    h(s)                             — escapa HTML
    pill(text, style)                — badge/pill HTML
    fmt_gb(gb)                       — formatea GB/MB legible
    fmt_num(n)                       — formatea número con separadores
    bar(pct, color)                  — barra de progreso inline
    sec_header(num, title, badge, badge_color, anchor)
    build_topbar(label, title, badge_text, meta)
    build_sidebar(sections, stats)
    build_footer(cliente_nombre, doc_label)
    css(config)                      — CSS completo con acento del cliente
===============================================================================
"""

import os
import glob
import json
import argparse
import pandas as pd
from datetime import datetime


# ===========================================================================
# CONSTANTES
# ===========================================================================

BIBLIOTECAS_SISTEMA = {
    "Site Assets", "Activos del sitio", "SiteAssets",
    "Form Templates", "Style Library", "Preservation Hold Library",
    "Pages", "Páginas del sitio", "Recursos del sitio"
}

ARCHIVOS_SISTEMA = {"__siteIcon__.png", "__siteIcon__.jpg"}

FORMATOS_ESPECIALES = {
    "dwg":  ("CAD / Planos",      "Preview limitado en SPO — visor integrado disponible"),
    "dxf":  ("CAD / Planos",      "Similar a DWG — intercambio AutoCAD"),
    "heic": ("Imagen",            "Sin preview nativo en SPO — considerar convertir a JPG"),
    "xlsm": ("Excel con macros",  "Macros no ejecutan en SPO — solo descarga local"),
    "xlsb": ("Excel binario",     "Formato binario — sin co-autoría en SPO"),
    "msg":  ("Email",             "Adjuntos de Outlook — sin preview en SPO"),
    "mp4":  ("Video",             "Videos pesados — evaluar Azure Blob Storage"),
    "mov":  ("Video",             "Videos pesados — evaluar Azure Blob Storage"),
    "avi":  ("Video",             "Videos pesados — evaluar Azure Blob Storage"),
    "zip":  ("Comprimido",        "ZIPs grandes pueden superar límite SPMT (250 MB)"),
    "rar":  ("Comprimido",        "RAR no tiene preview en SPO"),
    "7z":   ("Comprimido",        "7z no tiene preview en SPO"),
    "xer":  ("Primavera",         "Formato propietario — no indexable por SPO"),
    "nwd":  ("CAD / Planos",      "Navisworks — requiere visor dedicado"),
}

# Colores de acento disponibles para clientes
ACCENT_PRESETS = {
    "teal":   ("006B6B", "E6F5F5", "99D6D6"),
    "blue":   ("0050C8", "EBF2FF", "C8DCFF"),
    "orange": ("C84B00", "FFF2EB", "FFD5B8"),
    "green":  ("006E3A", "EBFAF2", "B3EDD0"),
    "purple": ("5B3FA8", "F0EEFF", "C8B0FF"),
    "red":    ("8B1A1A", "FFF0F0", "FFCCCC"),
}


# ===========================================================================
# CONFIGURACIÓN
# ===========================================================================

def load_config(config_path: str) -> dict:
    """
    Carga y valida el cliente_config.json.
    Devuelve un dict con todos los parámetros, con defaults para los opcionales.
    Lanza ValueError si faltan campos obligatorios.
    """
    if not os.path.exists(config_path):
        raise FileNotFoundError(
            f"No se encontró el archivo de configuración: {config_path}\n"
            "Generarlo con Config/kit_config_editor.html"
        )

    with open(config_path, encoding="utf-8") as f:
        raw = json.load(f)

    # Validar campos obligatorios
    obligatorios = [
        ("cliente", "nombre"),
        ("cliente", "tenant_prefix"),
        ("cliente", "carpeta_local"),
    ]
    for seccion, campo in obligatorios:
        if not raw.get(seccion, {}).get(campo):
            raise ValueError(
                f"Campo obligatorio faltante en la configuración: {seccion}.{campo}\n"
                "Completar en el editor HTML antes de ejecutar."
            )

    # Construir config con defaults
    c = raw.get("cliente", {})
    f = raw.get("fechas", {})
    r = raw.get("refresh", {})
    cl = raw.get("clasificacion", {})

    tenant = c.get("tenant_prefix", "")
    fecha_mig_str = f.get("migracion", "2000-01-01")

    try:
        fecha_migracion = datetime.strptime(fecha_mig_str, "%Y-%m-%d")
    except ValueError:
        raise ValueError(
            f"Formato de fecha incorrecto: {fecha_mig_str}. Usar YYYY-MM-DD."
        )

    inactividad_anios = int(f.get("inactividad_anios", 2))
    fecha_inactividad = datetime(
        fecha_migracion.year - inactividad_anios,
        fecha_migracion.month,
        fecha_migracion.day
    )

    # Color de acento — desde config o default teal
    accent_key = c.get("color_acento", "teal")
    accent = ACCENT_PRESETS.get(accent_key, ACCENT_PRESETS["teal"])

    return {
        # Identidad
        "nombre":          c.get("nombre", ""),
        "tenant_prefix":   tenant,
        "sharepoint_url":  c.get("sharepoint_url", f"https://{tenant}.sharepoint.com"),
        "admin_url":       c.get("admin_url", f"https://{tenant}-admin.sharepoint.com"),
        "client_id":       c.get("client_id", ""),
        "upn_circo":       c.get("upn_circo", ""),
        "carpeta_local":   c.get("carpeta_local", ""),

        # Fechas
        "fecha_migracion":  fecha_migracion,
        "fecha_inactividad": fecha_inactividad,
        "inactividad_anios": inactividad_anios,

        # Refresh
        "dias_umbral": int(r.get("dias_umbral", 7)),

        # Clasificación
        "gb_principal":      float(cl.get("gb_principal", 10.0)),
        "gb_mediano":        float(cl.get("gb_mediano", 1.0)),
        "mb_archivo_grande": int(cl.get("mb_archivo_grande", 100)),

        # Listas
        "sitios_excluidos":   raw.get("sitios_excluidos", []),
        "canales_excluidos":  raw.get("canales_excluidos", []),
        "informes_prioritarios": raw.get("informes_prioritarios", []),

        # Acento visual
        "accent_key":  accent_key,
        "accent":      accent[0],   # hex sin #
        "accent_bg":   accent[1],   # hex sin #
        "accent_mid":  accent[2],   # hex sin #
    }


def parse_args(description: str) -> argparse.Namespace:
    """
    Parser de argumentos estándar para los tres scripts.
    Acepta --config y --sitio (este último solo lo usa generar_informes_sitio.py).
    """
    parser = argparse.ArgumentParser(description=description)
    parser.add_argument(
        "--config",
        required=True,
        help="Path al cliente_config.json"
    )
    parser.add_argument(
        "--sitio",
        default="",
        help="(opcional) Nombre exacto de un sitio — genera solo ese informe"
    )
    return parser.parse_args()


# ===========================================================================
# CARGA Y LIMPIEZA DE CSVs
# ===========================================================================

def load_csvs(config: dict) -> pd.DataFrame:
    """
    Carga todos los CSVs de relevamiento de la carpeta configurada.
    Busca los patrones del kit generalizado y también patrones de naming legacy.
    para mantener compatibilidad con ejecuciones anteriores.
    """
    carpeta = config["carpeta_local"]
    prefijo = config["nombre"].replace(" ", "_").replace(".", "_")

    patrones = [
        # Kit generalizado
        os.path.join(carpeta, f"{prefijo}_B_Docs_*.csv"),
        # Compatibilidad con relevamientos anteriores al kit (naming legacy)
        os.path.join(carpeta, "*_B_Docs_*.csv"),
        os.path.join(carpeta, "*_Relevamiento_*[!LOG].csv"),
    ]

    excluir = {
        "CONSOLIDADO", "_LOG", "Resumen", "AsignarOwner",
        "Inventario", "Refresh", "Analisis", "Indice"
    }

    archivos = []
    vistos = set()
    for p in patrones:
        for f in glob.glob(p):
            bn = os.path.basename(f)
            if f in vistos:
                continue
            if any(x in bn for x in excluir):
                continue
            vistos.add(f)
            archivos.append(f)

    if not archivos:
        raise FileNotFoundError(
            f"No se encontraron CSVs de relevamiento en: {carpeta}\n"
            "Ejecutar primero Script_B_Masivo.ps1"
        )

    print(f"CSVs encontrados: {len(archivos)}")
    dfs = []
    for f in sorted(archivos):
        try:
            df = pd.read_csv(f, encoding="utf-8-sig", low_memory=False)
            dfs.append(df)
            print(f"  ok  {os.path.basename(f)} — {len(df):,} filas")
        except Exception as e:
            print(f"  err {os.path.basename(f)}: {e}")

    if not dfs:
        raise ValueError("Ningún CSV se pudo leer correctamente.")

    return pd.concat(dfs, ignore_index=True)


def clean_df(df: pd.DataFrame, config: dict) -> pd.DataFrame:
    """
    Limpia el DataFrame y agrega columnas derivadas:
        Post_Migracion   — bool: modificado después de fecha_migracion
        Sin_Actividad_Ny — bool: modificado antes de fecha_inactividad
        Fecha_Modificacion_dt / Fecha_Creacion_dt — datetime parseado
    """
    df = df[~df["Biblioteca"].isin(BIBLIOTECAS_SISTEMA)].copy()
    df = df[~df["Nombre_Archivo"].isin(ARCHIVOS_SISTEMA)].copy()

    df["Fecha_Modificacion_dt"] = pd.to_datetime(
        df["Fecha_Modificacion"], errors="coerce"
    )
    df["Fecha_Creacion_dt"] = pd.to_datetime(
        df["Fecha_Creacion"], errors="coerce"
    )
    df["Post_Migracion"] = df["Fecha_Modificacion_dt"] > config["fecha_migracion"]
    df["Sin_Actividad_Ny"] = df["Fecha_Modificacion_dt"] < config["fecha_inactividad"]

    df["Tamaño_Bytes"] = pd.to_numeric(df["Tamaño_Bytes"], errors="coerce").fillna(0)
    df["Tamaño_MB"]    = pd.to_numeric(df["Tamaño_MB"],    errors="coerce").fillna(0)

    if "Profundidad_Carpeta" in df.columns:
        df["Profundidad_Carpeta"] = pd.to_numeric(
            df["Profundidad_Carpeta"], errors="coerce"
        ).fillna(0)

    return df


# ===========================================================================
# RUTAS
# ===========================================================================

def get_carpeta_informes(config: dict) -> str:
    """Devuelve (y crea si no existe) la carpeta de salida de informes HTML."""
    path = os.path.join(config["carpeta_local"], "Informes")
    os.makedirs(path, exist_ok=True)
    return path


def get_output_path(config: dict, sufijo: str, ts: str, carpeta: str = None) -> str:
    """Construye el path de salida para un archivo generado."""
    prefijo = config["nombre"].replace(" ", "_").replace(".", "_")
    nombre  = f"{prefijo}_{sufijo}_{ts}.html"
    base    = carpeta if carpeta else config["carpeta_local"]
    return os.path.join(base, nombre)


# ===========================================================================
# RECOMENDACIÓN
# ===========================================================================

def rec_info(post_mig: int, pct_pm: float) -> tuple:
    """
    Devuelve (texto_recomendacion, clase_pill, color_css) según actividad post-migración.
    """
    if post_mig == 0:
        return "Archivar",                  "pill-red",   "red"
    if pct_pm >= 30:
        return "Activo — revisar destino",  "pill-green", "green"
    return     "Revisar",                   "pill-gold",  "gold"


# ===========================================================================
# HELPERS HTML
# ===========================================================================

def h(s) -> str:
    """Escapa caracteres HTML."""
    return (str(s)
            .replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
            .replace('"', "&quot;"))


def pill(text: str, style: str = "ghost") -> str:
    return f'<span class="pill pill-{style}">{h(text)}</span>'


def fmt_gb(gb: float) -> str:
    if gb >= 1:
        return f"{gb:.1f} GB"
    if gb >= 0.001:
        return f"{gb * 1024:.0f} MB"
    return "< 1 MB"


def fmt_num(n) -> str:
    return f"{int(n):,}".replace(",", ".")


def bar(pct: float, color: str = "teal") -> str:
    w = min(pct, 100)
    return (
        f'<div class="sbar-wrap">'
        f'<div class="sbar-outer">'
        f'<div class="sbar-inner bar-{color}" style="width:{w:.0f}%"></div>'
        f'</div>'
        f'<span class="sbar-val">{pct:.1f}%</span>'
        f'</div>'
    )


def sec_header(num: str, title: str, badge: str, badge_color: str, anchor: str) -> str:
    return (
        f'<div class="section-header" id="{anchor}">'
        f'<span class="section-num">{num}</span>'
        f'<span class="section-title">{title}</span>'
        f'<span class="section-badge badge-{badge_color}">{badge}</span>'
        f'</div>'
    )


def build_topbar(label: str, title: str, badge_text: str, meta: str) -> str:
    return (
        f'<div class="topbar">'
        f'<div class="topbar-left">'
        f'<span class="topbar-doc">{h(label)}</span>'
        f'<span class="topbar-title">{h(title)}</span>'
        f'<span class="topbar-badge">{h(badge_text)}</span>'
        f'</div>'
        f'<span class="topbar-meta">{h(meta)}</span>'
        f'</div>'
    )


def build_sidebar(sections: list, stats: list = None) -> str:
    """
    sections: lista de dicts con keys:
        label (str), items: lista de (num, text, anchor)
    stats: lista de (text,) — se muestran sin anchor
    """
    html = '<nav class="sidebar">'
    for sec in sections:
        html += f'<div class="sidebar-section"><div class="sidebar-label">{h(sec["label"])}</div>'
        for num, text, anchor in sec["items"]:
            html += (
                f'<a href="#{anchor}">'
                f'<span class="sidebar-num">{h(str(num))}</span>'
                f'{h(text)}'
                f'</a>'
            )
        html += '</div>'

    if stats:
        html += '<div class="sidebar-section"><div class="sidebar-label">Stats</div>'
        for text in stats:
            html += (
                f'<a href="#">'
                f'<span class="sidebar-num"></span>'
                f'{h(text)}'
                f'</a>'
            )
        html += '</div>'

    html += '</nav>'
    return html


def build_footer(cliente_nombre: str, doc_label: str) -> str:
    ts = datetime.now().strftime("%d/%m/%Y %H:%M")
    return (
        f'<div class="footer">'
        f'<div class="footer-meta">'
        f'{h(cliente_nombre)} · {h(doc_label)}<br>'
        f'Generado: {ts} · Circo Studio'
        f'</div>'
        f'<div class="footer-logo">CIRCO STUDIO · MICROSOFT GOLD PARTNER</div>'
        f'</div>'
    )


# ===========================================================================
# CSS DEL DESIGN SYSTEM
# Con --accent parametrizable desde el config del cliente.
# ===========================================================================

def css(config: dict) -> str:
    """
    Devuelve el CSS completo del design system.
    Los tres colores de acento se inyectan desde el config del cliente.
    """
    accent     = "#" + config.get("accent",     "006B6B")
    accent_bg  = "#" + config.get("accent_bg",  "E6F5F5")
    accent_mid = "#" + config.get("accent_mid", "99D6D6")

    return f"""
:root{{
  --bg:#F4F2ED;--bg-2:#EAE7DF;--bg-3:#DDD9CE;--white:#FFFFFF;
  --ink:#1A1A1A;--ink-mid:#3D3D3D;--ink-light:#6B6B6B;--ink-faint:#A0A0A0;
  --blue:#0050C8;--blue-bg:#EBF2FF;--blue-mid:#C8DCFF;
  --orange:#C84B00;--orange-bg:#FFF2EB;--orange-mid:#FFD5B8;
  --green:#006E3A;--green-bg:#EBFAF2;--green-mid:#B3EDD0;
  --red:#8B1A1A;--red-bg:#FFF0F0;--red-mid:#FFCCCC;
  --gold:#8A6300;--gold-bg:#FFF8E6;--gold-mid:#FFE9A0;
  --accent:{accent};--accent-bg:{accent_bg};--accent-mid:{accent_mid};
  --border:#D0CAC0;--border-2:#B8B0A2;
  --font-sans:'IBM Plex Sans',sans-serif;
  --font-mono:'IBM Plex Mono',monospace;
  --font-serif:'IBM Plex Serif',serif;
}}
*,*::before,*::after{{box-sizing:border-box;margin:0;padding:0;}}
html{{scroll-behavior:smooth;}}
body{{background:var(--bg);color:var(--ink);font-family:var(--font-sans);font-size:14px;line-height:1.65;}}

/* TOPBAR */
.topbar{{background:var(--ink);padding:0 48px;height:52px;display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;z-index:200;}}
.topbar-left{{display:flex;align-items:center;gap:16px;}}
.topbar-doc{{font-family:var(--font-mono);font-size:11px;color:rgba(255,255,255,0.45);letter-spacing:.04em;}}
.topbar-title{{font-size:13px;font-weight:500;color:rgba(255,255,255,0.9);}}
.topbar-badge{{font-family:var(--font-mono);font-size:10px;padding:2px 8px;border-radius:2px;background:var(--accent);color:white;letter-spacing:.06em;}}
.topbar-meta{{font-family:var(--font-mono);font-size:10px;color:rgba(255,255,255,0.35);}}

/* LAYOUT */
.layout{{display:flex;min-height:calc(100vh - 52px);}}
.sidebar{{width:240px;flex-shrink:0;background:var(--bg-2);border-right:1px solid var(--border);padding:32px 0;position:sticky;top:52px;height:calc(100vh - 52px);overflow-y:auto;}}
.sidebar-label{{font-family:var(--font-mono);font-size:9px;font-weight:500;color:var(--ink-faint);letter-spacing:.1em;text-transform:uppercase;padding:0 24px 8px;}}
.sidebar-section{{margin-bottom:8px;}}
.sidebar a{{display:flex;align-items:center;gap:8px;padding:7px 24px;font-size:12.5px;color:var(--ink-light);text-decoration:none;border-left:2px solid transparent;transition:all .15s;}}
.sidebar a:hover{{color:var(--ink);background:var(--bg-3);border-left-color:var(--border-2);}}
.sidebar a.active{{color:var(--accent);font-weight:500;border-left-color:var(--accent);background:var(--accent-bg);}}
.sidebar-num{{font-family:var(--font-mono);font-size:10px;color:var(--ink-faint);min-width:18px;}}
.main{{flex:1;padding:48px 56px;max-width:1200px;}}

/* HERO */
.hero{{border:1px solid var(--border);background:var(--white);padding:40px 48px;margin-bottom:40px;position:relative;overflow:hidden;}}
.hero::before{{content:'';position:absolute;top:0;left:0;right:0;height:3px;background:linear-gradient(90deg,var(--accent) 0%,var(--blue) 40%,var(--orange) 80%,var(--gold) 100%);}}
.hero-eyebrow{{font-family:var(--font-mono);font-size:10px;color:var(--ink-faint);letter-spacing:.1em;text-transform:uppercase;margin-bottom:12px;}}
.hero h1{{font-family:var(--font-serif);font-size:28px;font-weight:600;color:var(--ink);margin-bottom:12px;line-height:1.25;}}
.hero-sub{{font-size:14px;color:var(--ink-light);max-width:700px;line-height:1.7;margin-bottom:28px;}}
.hero-stats{{display:flex;gap:28px;padding-top:24px;border-top:1px solid var(--border);flex-wrap:wrap;}}
.stat{{display:flex;flex-direction:column;gap:4px;}}
.stat-value{{font-family:var(--font-mono);font-size:22px;font-weight:500;}}
.stat-value.accent{{color:var(--accent);}}
.stat-value.orange{{color:var(--orange);}}
.stat-value.blue{{color:var(--blue);}}
.stat-value.green{{color:var(--green);}}
.stat-value.gold{{color:var(--gold);}}
.stat-value.red{{color:var(--red);}}
.stat-value.ink{{color:var(--ink);}}
.stat-label{{font-size:11px;color:var(--ink-faint);text-transform:uppercase;letter-spacing:.06em;}}

/* SECCIONES */
.section{{margin-bottom:48px;}}
.section-header{{display:flex;align-items:baseline;gap:12px;margin-bottom:20px;padding-bottom:12px;border-bottom:1px solid var(--border);}}
.section-num{{font-family:var(--font-mono);font-size:11px;color:var(--ink-faint);font-weight:500;min-width:24px;}}
.section-title{{font-family:var(--font-serif);font-size:20px;font-weight:600;color:var(--ink);}}
.section-badge{{font-family:var(--font-mono);font-size:9px;padding:3px 8px;border-radius:2px;letter-spacing:.08em;margin-left:auto;white-space:nowrap;}}
.badge-red{{background:var(--red-bg);color:var(--red);border:1px solid var(--red-mid);}}
.badge-orange{{background:var(--orange-bg);color:var(--orange);border:1px solid var(--orange-mid);}}
.badge-gold{{background:var(--gold-bg);color:var(--gold);border:1px solid var(--gold-mid);}}
.badge-blue{{background:var(--blue-bg);color:var(--blue);border:1px solid var(--blue-mid);}}
.badge-green{{background:var(--green-bg);color:var(--green);border:1px solid var(--green-mid);}}
.badge-accent{{background:var(--accent-bg);color:var(--accent);border:1px solid var(--accent-mid);}}

/* CALLOUTS */
.callout{{padding:14px 18px;border-radius:2px;margin:20px 0;border-left:3px solid;font-size:13px;line-height:1.6;}}
.callout-label{{display:block;font-family:var(--font-mono);font-size:10px;letter-spacing:.08em;text-transform:uppercase;margin-bottom:6px;font-weight:600;}}
.callout-red{{background:var(--red-bg);border-color:var(--red);color:var(--ink-mid);}}.callout-red .callout-label{{color:var(--red);}}
.callout-orange{{background:var(--orange-bg);border-color:var(--orange);color:var(--ink-mid);}}.callout-orange .callout-label{{color:var(--orange);}}
.callout-gold{{background:var(--gold-bg);border-color:var(--gold);color:var(--ink-mid);}}.callout-gold .callout-label{{color:var(--gold);}}
.callout-green{{background:var(--green-bg);border-color:var(--green);color:var(--ink-mid);}}.callout-green .callout-label{{color:var(--green);}}
.callout-accent{{background:var(--accent-bg);border-color:var(--accent);color:var(--ink-mid);}}.callout-accent .callout-label{{color:var(--accent);}}
.callout-blue{{background:var(--blue-bg);border-color:var(--blue);color:var(--ink-mid);}}.callout-blue .callout-label{{color:var(--blue);}}

/* TABLAS */
.table-wrap{{overflow-x:auto;margin:16px 0;border:1px solid var(--border);border-radius:2px;}}
table{{width:100%;border-collapse:collapse;font-size:12.5px;}}
thead tr{{background:var(--bg-2);}}
th{{padding:9px 12px;text-align:left;font-size:10px;font-weight:600;color:var(--ink-light);text-transform:uppercase;letter-spacing:.05em;border-bottom:1px solid var(--border);white-space:nowrap;}}
td{{padding:8px 12px;border-bottom:1px solid var(--border);color:var(--ink-mid);vertical-align:middle;}}
tbody tr:last-child td{{border-bottom:none;}}
tbody tr:hover{{background:var(--bg);}}
.mono{{font-family:var(--font-mono);font-size:11px;}}
.bold{{font-weight:600;color:var(--ink);}}
.td-right{{text-align:right;}}
.td-center{{text-align:center;}}
.row-critical{{background:rgba(139,26,26,0.04);}}
.row-warning{{background:rgba(138,99,0,0.04);}}
.row-ok{{background:rgba(0,110,58,0.04);}}

/* PILLS */
.pill{{display:inline-flex;align-items:center;padding:1px 7px;border-radius:9px;font-size:10px;font-weight:500;font-family:var(--font-mono);white-space:nowrap;}}
.pill-red{{background:var(--red-bg);color:var(--red);border:1px solid var(--red-mid);}}
.pill-orange{{background:var(--orange-bg);color:var(--orange);border:1px solid var(--orange-mid);}}
.pill-gold{{background:var(--gold-bg);color:var(--gold);border:1px solid var(--gold-mid);}}
.pill-blue{{background:var(--blue-bg);color:var(--blue);border:1px solid var(--blue-mid);}}
.pill-green{{background:var(--green-bg);color:var(--green);border:1px solid var(--green-mid);}}
.pill-accent{{background:var(--accent-bg);color:var(--accent);border:1px solid var(--accent-mid);}}
.pill-ghost{{background:var(--bg-2);color:var(--ink-light);border:1px solid var(--border);}}

/* BARRAS */
.sbar-wrap{{display:flex;align-items:center;gap:6px;}}
.sbar-outer{{width:80px;height:5px;background:var(--bg-3);border-radius:1px;flex-shrink:0;}}
.sbar-inner{{height:5px;border-radius:1px;}}
.bar-red{{background:var(--red);}}.bar-orange{{background:var(--orange);}}
.bar-gold{{background:var(--gold);}}.bar-blue{{background:var(--blue);}}
.bar-green{{background:var(--green);}}.bar-accent{{background:var(--accent);}}
.sbar-val{{font-family:var(--font-mono);font-size:11px;color:var(--ink-mid);white-space:nowrap;}}

/* CARDS RESUMEN */
.summary-grid{{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin:20px 0;}}
.sum-card{{background:var(--white);border:1px solid var(--border);padding:18px 20px;position:relative;}}
.sum-card::before{{content:'';position:absolute;left:0;top:0;bottom:0;width:3px;}}
.sum-card.orange::before{{background:var(--orange);}}.sum-card.red::before{{background:var(--red);}}
.sum-card.blue::before{{background:var(--blue);}}.sum-card.accent::before{{background:var(--accent);}}
.sum-card.green::before{{background:var(--green);}}.sum-card.gold::before{{background:var(--gold);}}
.sc-val{{font-family:var(--font-mono);font-size:26px;font-weight:500;margin-bottom:4px;}}
.sum-card.orange .sc-val{{color:var(--orange);}}.sum-card.red .sc-val{{color:var(--red);}}
.sum-card.blue .sc-val{{color:var(--blue);}}.sum-card.accent .sc-val{{color:var(--accent);}}
.sum-card.green .sc-val{{color:var(--green);}}.sum-card.gold .sc-val{{color:var(--gold);}}
.sc-label{{font-size:12px;color:var(--ink-light);line-height:1.4;}}

/* STORAGE BARS */
.storage-bars{{background:var(--white);border:1px solid var(--border);padding:24px;}}
.storage-bar-item{{margin-bottom:14px;}}.storage-bar-item:last-child{{margin-bottom:0;}}
.sbi-header{{display:flex;justify-content:space-between;align-items:baseline;margin-bottom:5px;}}
.sbi-name{{font-size:12.5px;font-weight:500;color:var(--ink);white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:440px;}}
.sbi-val{{font-family:var(--font-mono);font-size:11px;color:var(--ink-light);white-space:nowrap;margin-left:8px;}}
.sbi-bar-outer{{height:6px;background:var(--bg-3);border-radius:1px;}}
.sbi-bar-inner{{height:6px;border-radius:1px;}}
.sbi-meta{{font-size:11px;color:var(--ink-faint);margin-top:3px;}}

/* MISC */
.placeholder{{background:var(--bg-2);border:2px dashed var(--border-2);padding:20px 24px;border-radius:2px;color:var(--ink-faint);font-size:13px;font-style:italic;margin:16px 0;}}
.placeholder-label{{display:block;font-family:var(--font-mono);font-size:10px;letter-spacing:.08em;text-transform:uppercase;color:var(--ink-faint);margin-bottom:8px;font-weight:600;}}
hr.divider{{border:none;border-top:1px solid var(--border);margin:40px 0;}}
p{{margin-bottom:12px;color:var(--ink-mid);}}
strong{{color:var(--ink);}}
h3{{font-size:14px;font-weight:600;color:var(--ink);margin:20px 0 12px;}}
code{{font-family:var(--font-mono);font-size:12px;background:var(--bg-2);padding:1px 5px;}}
.footer{{margin-top:64px;padding-top:20px;border-top:1px solid var(--border);display:flex;justify-content:space-between;align-items:center;}}
.footer-meta{{font-family:var(--font-mono);font-size:10px;color:var(--ink-faint);line-height:1.8;}}
.footer-logo{{font-family:var(--font-mono);font-size:10px;color:var(--ink-faint);letter-spacing:.08em;}}
""".strip()


# ===========================================================================
# BOILERPLATE HTML
# ===========================================================================

GOOGLE_FONTS = (
    '<link rel="preconnect" href="https://fonts.googleapis.com">'
    '<link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@300;400;500;600'
    '&family=IBM+Plex+Mono:wght@400;500&family=IBM+Plex+Serif:wght@400;600&display=swap" rel="stylesheet">'
)

SIDEBAR_JS = """
<script>
const _secs = document.querySelectorAll('.section[id],.hero[id]');
const _links = document.querySelectorAll('.sidebar a[href^="#"]');
_secs.forEach(s => {
  new IntersectionObserver(entries => {
    entries.forEach(e => {
      if (e.isIntersecting) {
        _links.forEach(l => l.classList.remove('active'));
        const a = document.querySelector('.sidebar a[href="#'+e.target.id+'"]');
        if (a) a.classList.add('active');
      }
    });
  }, {rootMargin: '-20% 0px -70% 0px'}).observe(s);
});
</script>
"""


def html_doc(title: str, config: dict, body: str) -> str:
    """Envuelve el cuerpo en un documento HTML completo."""
    return (
        f'<!DOCTYPE html>\n<html lang="es">\n<head>\n'
        f'<meta charset="UTF-8">\n'
        f'<meta name="viewport" content="width=device-width, initial-scale=1.0">\n'
        f'<title>{h(title)}</title>\n'
        f'{GOOGLE_FONTS}\n'
        f'<style>{css(config)}</style>\n'
        f'</head>\n<body>\n'
        f'{body}\n'
        f'{SIDEBAR_JS}\n'
        f'</body>\n</html>'
    )
