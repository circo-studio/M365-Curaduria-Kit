"""
===============================================================================
generar_indice.py  |  Circo Studio  |  Kit de Curaduría Informacional M365
-------------------------------------------------------------------------------
Genera el índice maestro con links a todos los informes de sitio.

Diferencia dos tipos de informes:
  - Prioritarios: análisis profundo manual, definidos en cliente_config.json
  - Automáticos:  generados por generar_informes_sitio.py, detectados por naming

Secciones:
  00 Totales del tenant (hero)
  01 Informes prioritarios
  02 Sitios principales ≥ GB_PRINCIPAL
  03 Sitios medianos  GB_MEDIANO – GB_PRINCIPAL
  04 Sitios menores   < GB_MEDIANO
  05 Candidatos a archivo

Uso:
    python generar_indice.py --config "...\\Config\\cliente_config.json"

Genera:
    [CarpetaLocal]\\Informes\\[Cliente]_Indice_Informes.html
===============================================================================
"""

import os
import sys
import re
import glob
import pandas as pd
from datetime import datetime

sys.path.insert(0, os.path.dirname(__file__))
from m365_curaduria import (
    load_config, parse_args, load_csvs, clean_df,
    get_carpeta_informes, rec_info,
    h, pill, fmt_gb, fmt_num,
    build_topbar, build_sidebar, build_footer, html_doc, sec_header,
)


# ===========================================================================
# HELPERS ESPECÍFICOS DEL ÍNDICE
# ===========================================================================

def encontrar_informe_auto(sitio_nombre: str, carpeta_informes: str,
                            prefijo_cliente: str) -> str | None:
    """Busca el HTML automático más reciente para un sitio dado."""
    nombre_limpio = re.sub(r"[^a-zA-Z0-9_-]", "_", sitio_nombre)
    nombre_limpio = re.sub(r"_+", "_", nombre_limpio)[:60]
    # Patrón del kit generalizado
    matches = glob.glob(os.path.join(carpeta_informes,
                                     f"{prefijo_cliente}_Informe_{nombre_limpio}_*.html"))
    # Compatibilidad con informes generados antes del kit (naming legacy)
    if not matches:
        matches = glob.glob(os.path.join(carpeta_informes,
                                         f"*_Informe_{nombre_limpio}_*.html"))
    return os.path.basename(max(matches, key=os.path.getmtime)) if matches else None


def tipo_informe(sitio_nombre: str, carpeta_informes: str,
                 config: dict, prefijo_cliente: str) -> tuple:
    """
    Devuelve (filename, tipo) donde tipo es:
      'prioritario' | 'prioritario_faltante' | 'automatico' | None
    """
    # Buscar en la lista de prioritarios del config
    for p in config["informes_prioritarios"]:
        if p.get("sitio") == sitio_nombre:
            fname  = p.get("archivo_html", "")
            existe = fname and os.path.exists(os.path.join(carpeta_informes, fname))
            return fname, ("prioritario" if existe else "prioritario_faltante")
    # Si no es prioritario, buscar automático
    fname = encontrar_informe_auto(sitio_nombre, carpeta_informes, prefijo_cliente)
    return (fname, "automatico") if fname else (None, None)


def link_prio(fname: str) -> str:
    return f'<a href="{h(fname)}" target="_blank" class="link-prio">★ abrir informe</a>'

def link_auto(fname: str) -> str:
    return f'<a href="{h(fname)}" target="_blank" class="link-auto">↗ informe</a>'

def link_sin() -> str:
    return '<span class="link-sin">sin informe</span>'

def link_pendiente() -> str:
    return '<span class="link-sin">★ pendiente</span>'


# ===========================================================================
# CSS ADICIONAL PARA EL ÍNDICE (complementa el CSS base del módulo)
# ===========================================================================

CSS_INDICE = """
.total-bar-wrap{background:var(--white);border:1px solid var(--border);padding:22px 28px;margin-bottom:28px;}
.total-bar-label{font-family:var(--font-mono);font-size:10px;color:var(--ink-faint);letter-spacing:.08em;text-transform:uppercase;margin-bottom:12px;}
.total-bar-track{height:14px;background:var(--bg-3);border-radius:2px;display:flex;overflow:hidden;margin-bottom:12px;}
.total-bar-seg{height:14px;}
.total-bar-legend{display:flex;gap:20px;flex-wrap:wrap;}
.tbl-item{display:flex;align-items:center;gap:7px;font-size:12px;}
.tbl-dot{width:10px;height:10px;border-radius:2px;flex-shrink:0;}
.tbl-label{color:var(--ink-mid);}
.tbl-val{font-family:var(--font-mono);color:var(--ink-light);font-size:11px;}
.prio-cards{display:flex;flex-direction:column;gap:8px;margin:16px 0;}
.prio-card{background:var(--white);border:1px solid var(--border);border-left:3px solid var(--gold);padding:14px 18px;display:grid;grid-template-columns:32px 1fr auto;gap:12px;align-items:start;}
.prio-card.green{border-left-color:var(--green);}
.prio-card.gold{border-left-color:var(--gold);}
.prio-card.red{border-left-color:var(--red);}
.prio-num{font-family:var(--font-mono);font-size:16px;font-weight:500;color:var(--ink-faint);padding-top:2px;}
.prio-name{font-weight:600;font-size:14px;color:var(--ink);margin-bottom:5px;}
.prio-meta{display:flex;gap:14px;flex-wrap:wrap;align-items:center;margin-bottom:5px;}
.prio-gb{font-family:var(--font-mono);font-size:12px;font-weight:500;}
.prio-gb.green{color:var(--green);}.prio-gb.gold{color:var(--gold);}.prio-gb.red{color:var(--red);}
.prio-destino{font-size:12px;color:var(--ink-light);}
.prio-hallazgo{font-size:12px;color:var(--ink-faint);line-height:1.5;}
.prio-right{display:flex;flex-direction:column;align-items:flex-end;gap:6px;white-space:nowrap;}
.site-card{background:var(--white);border:1px solid var(--border);padding:14px 18px;display:grid;grid-template-columns:1fr auto;gap:12px;align-items:center;border-left:3px solid var(--border);margin-bottom:6px;}
.site-card.red{border-left-color:var(--red);}
.site-card.green{border-left-color:var(--green);}
.site-card.gold{border-left-color:var(--gold);}
.sc-name{font-weight:600;font-size:13px;color:var(--ink);margin-bottom:5px;}
.sc-meta{display:flex;gap:14px;align-items:center;flex-wrap:wrap;}
.sc-gb{font-family:var(--font-mono);font-size:12px;font-weight:500;}
.sc-gb.red{color:var(--red);}.sc-gb.green{color:var(--green);}.sc-gb.gold{color:var(--gold);}
.link-prio{display:inline-flex;align-items:center;gap:4px;font-family:var(--font-mono);font-size:10px;color:var(--orange);text-decoration:none;padding:3px 10px;border:1px solid var(--orange-mid);border-radius:2px;background:var(--orange-bg);white-space:nowrap;}
.link-prio:hover{background:var(--orange);color:white;}
.link-auto{display:inline-flex;align-items:center;gap:4px;font-family:var(--font-mono);font-size:10px;color:var(--accent);text-decoration:none;padding:3px 10px;border:1px solid var(--accent-mid);border-radius:2px;background:var(--accent-bg);white-space:nowrap;}
.link-auto:hover{background:var(--accent);color:white;}
.link-sin{display:inline-flex;align-items:center;gap:4px;font-family:var(--font-mono);font-size:10px;color:var(--ink-faint);padding:3px 10px;border:1px solid var(--border);border-radius:2px;background:var(--bg-2);pointer-events:none;}
"""


# ===========================================================================
# CONSTRUCCIÓN HTML
# ===========================================================================

def build_html(df: pd.DataFrame, config: dict, carpeta_informes: str, ts_label: str) -> str:

    prefijo_cliente = config["nombre"].replace(" ", "_").replace(".", "_")

    # ── Métricas globales ────────────────────────────────────
    total_arch = len(df)
    total_gb   = df["Tamaño_Bytes"].sum() / 1e9
    total_sit  = df["Sitio"].nunique()
    usuarios   = df["Modificado_Por"].nunique()

    # ── Datos por sitio ──────────────────────────────────────
    sitios = []
    for (sitio, url), g in df.groupby(["Sitio", "URL_Sitio"]):
        n   = len(g)
        gb  = g["Tamaño_Bytes"].sum() / 1e9
        pm  = int(g["Post_Migracion"].sum())
        pp  = pm / n * 100 if n else 0
        ul  = g["Fecha_Modificacion_dt"].max()
        rec, rec_cls, rec_col = rec_info(pm, pp)
        fname, tinf = tipo_informe(sitio, carpeta_informes, config, prefijo_cliente)
        sitios.append({
            "sitio": sitio, "url": url, "archivos": n, "gb": gb,
            "post_mig": pm, "pct_pm": pp,
            "ultima": ul.strftime("%Y-%m-%d") if pd.notna(ul) else "—",
            "rec": rec, "rec_cls": rec_cls, "rec_col": rec_col,
            "informe": fname, "tipo_inf": tinf
        })
    sitios.sort(key=lambda x: x["gb"], reverse=True)

    gb_principal = config["gb_principal"]
    gb_mediano   = config["gb_mediano"]

    grandes  = [s for s in sitios if s["gb"] >= gb_principal]
    medianos = [s for s in sitios if gb_mediano <= s["gb"] < gb_principal]
    menores  = [s for s in sitios if s["gb"] < gb_mediano]

    gb_grandes  = sum(s["gb"] for s in grandes)
    gb_medianos = sum(s["gb"] for s in medianos)
    gb_menores  = sum(s["gb"] for s in menores)

    n_archivar = sum(1 for s in sitios if s["rec"] == "Archivar")
    n_activos  = sum(1 for s in sitios if s["rec"] == "Activo — revisar destino")
    n_prio     = sum(1 for s in sitios if s.get("tipo_inf") in ("prioritario", "prioritario_faltante"))
    n_auto     = sum(1 for s in sitios if s.get("tipo_inf") == "automatico")
    gb_arch    = sum(s["gb"] for s in sitios if s["rec"] == "Archivar")

    # ── Topbar ───────────────────────────────────────────────
    topbar = build_topbar(
        label=f"{h(config['nombre'])} · ÍNDICE",
        title=f"Relevamiento Legacy Teams — {total_sit} sitios · {fmt_gb(total_gb)}",
        badge_text="ANÁLISIS",
        meta=f"{config['sharepoint_url']} · {ts_label[:10]}"
    )

    # ── Sidebar ──────────────────────────────────────────────
    sidebar_links_prio = ""
    for s in [x for x in sitios if x.get("tipo_inf") == "prioritario"]:
        sidebar_links_prio += (
            f'<a href="{h(s["informe"])}" target="_blank" style="font-size:11px;padding:4px 24px">'
            f'<span class="sidebar-num" style="color:var(--orange)">★</span>'
            f'{h(s["sitio"][:28])}</a>'
        )

    sidebar = (
        '<nav class="sidebar">'
        '<div class="sidebar-section"><div class="sidebar-label">Navegación</div>'
        '<a href="#totales"><span class="sidebar-num">00</span>Totales del tenant</a>'
        '<a href="#prioritarios"><span class="sidebar-num">01</span>Informes prioritarios</a>'
        f'<a href="#principales"><span class="sidebar-num">02</span>Sitios principales ≥{gb_principal:.0f} GB</a>'
        f'<a href="#medianos"><span class="sidebar-num">03</span>Sitios medianos</a>'
        '<a href="#menores"><span class="sidebar-num">04</span>Sitios menores</a>'
        '<a href="#archivar"><span class="sidebar-num">05</span>Candidatos a archivo</a>'
        '</div>'
        f'<div class="sidebar-section"><div class="sidebar-label">Prioritarios ({n_prio})</div>'
        f'{sidebar_links_prio}'
        '</div></nav>'
    )

    # ── Hero ─────────────────────────────────────────────────
    hero = f"""<div class="hero" id="totales">
  <div class="hero-eyebrow">{h(config['nombre'])} · ÍNDICE DE RELEVAMIENTO LEGACY</div>
  <h1>Índice de informes de contenido<br>{total_sit} sitios relevados · {fmt_gb(total_gb)}</h1>
  <div class="hero-sub">
    Inventario de documentos de los sitios <strong>Legacy Teams</strong> del tenant
    <code>{h(config['sharepoint_url'])}</code>.
    Referencia de migración: <strong>{config['fecha_migracion'].strftime('%d/%m/%Y')}</strong>.
    Los informes <span style="color:var(--orange);font-weight:600">★ prioritarios</span> son análisis profundos con decisiones de destino.
    Los informes <span style="color:var(--accent);font-weight:600">automáticos</span> cubren el resto del corpus.
  </div>
  <div class="hero-stats">
    <div class="stat"><span class="stat-value ink">{total_sit}</span><span class="stat-label">Sitios relevados</span></div>
    <div class="stat"><span class="stat-value blue">{fmt_gb(total_gb)}</span><span class="stat-label">Storage total</span></div>
    <div class="stat"><span class="stat-value orange">{fmt_num(total_arch)}</span><span class="stat-label">Archivos totales</span></div>
    <div class="stat"><span class="stat-value orange">{n_prio}</span><span class="stat-label">Informes prioritarios</span></div>
    <div class="stat"><span class="stat-value accent">{n_auto}</span><span class="stat-label">Informes automáticos</span></div>
    <div class="stat"><span class="stat-value red">{n_archivar}</span><span class="stat-label">Candidatos a archivo</span></div>
  </div>
</div>"""

    # Barra de storage
    seg_colors = [
        (f"Principales ≥{gb_principal:.0f} GB", gb_grandes,  "#0050C8"),
        (f"Medianos {gb_mediano:.0f}–{gb_principal:.0f} GB", gb_medianos, config["accent"]),
        ("Menores <1 GB",     gb_menores,  "#A0A0A0"),
    ]
    segs   = "".join(f'<div class="total-bar-seg" style="width:{v/total_gb*100 if total_gb else 0:.1f}%;background:#{c}"></div>'
                     for _, v, c in seg_colors)
    legend = "".join(f'<div class="tbl-item"><div class="tbl-dot" style="background:#{c}"></div>'
                     f'<span class="tbl-label">{n}</span>'
                     f'<span class="tbl-val">{fmt_gb(v)} ({v/total_gb*100 if total_gb else 0:.0f}%)</span></div>'
                     for n, v, c in seg_colors)

    total_bar = f"""<div class="total-bar-wrap">
  <div class="total-bar-label">Distribución de storage por grupo</div>
  <div class="total-bar-track">{segs}</div>
  <div class="total-bar-legend">{legend}</div>
</div>"""

    sum_cards = f"""<div class="summary-grid">
  <div class="sum-card orange"><div class="sc-val">{n_prio}</div><div class="sc-label">Informes prioritarios — análisis profundo con decisiones de destino</div></div>
  <div class="sum-card accent"><div class="sc-val">{n_auto}</div><div class="sc-label">Informes automáticos — cobertura completa del resto</div></div>
  <div class="sum-card red"><div class="sc-val">{n_archivar}</div><div class="sc-label">Candidatos a archivo — sin actividad post-migración</div></div>
  <div class="sum-card green"><div class="sc-val">{n_activos}</div><div class="sc-label">Sitios activos post-migración — revisar destino</div></div>
</div>"""

    # ── 01 Informes prioritarios ─────────────────────────────
    prio_cards = ""
    for i, p in enumerate(config["informes_prioritarios"], 1):
        sitio_key = p.get("sitio", "")
        fname     = p.get("archivo_html", "")
        col       = p.get("estado", "gold")
        existe    = fname and os.path.exists(os.path.join(carpeta_informes, fname))
        lnk       = link_prio(fname) if existe else link_pendiente()

        prio_cards += f"""<div class="prio-card {col}">
  <div class="prio-num">0{i}</div>
  <div class="prio-body">
    <div class="prio-name">{h(sitio_key)}</div>
    <div class="prio-meta">
      <span class="prio-gb {col}">{h(p.get('storage','—'))}</span>
      <span class="prio-destino">→ {h(p.get('destino','—'))}</span>
    </div>
  </div>
  <div class="prio-right">{lnk}</div>
</div>"""

    if not prio_cards:
        prio_cards = '<p style="color:var(--ink-faint);font-style:italic">No hay informes prioritarios configurados en el cliente_config.json.</p>'

    sec_01 = f"""<div class="section" id="prioritarios">
  {sec_header("01", "Informes prioritarios", f"{len(config['informes_prioritarios'])} sitios · análisis profundo", "orange", "prioritarios")}
  <p>Estos sitios concentran el grueso del storage y tienen decisiones de destino no triviales.
  Cada informe incluye análisis de contenido, hallazgos críticos y mapa de destino.</p>
  <div class="prio-cards">{prio_cards}</div>
</div>"""

    # ── 02 Sitios principales (automáticos) ──────────────────
    site_cards = ""
    for s in grandes:
        if s.get("tipo_inf") in ("prioritario", "prioritario_faltante"):
            continue
        pm_c = "green" if s["pct_pm"] >= 30 else ("red" if s["pct_pm"] == 0 else "gold")
        pm_w = min(s["pct_pm"], 100)
        lnk  = link_auto(s["informe"]) if s["informe"] else link_sin()
        site_cards += f"""<div class="site-card {s['rec_col']}">
  <div>
    <div class="sc-name">{h(s['sitio'])}</div>
    <div class="sc-meta">
      <span class="sc-gb {s['rec_col']}">{fmt_gb(s['gb'])}</span>
      <span style="font-family:var(--font-mono);font-size:11px;color:var(--ink-faint)">{fmt_num(s['archivos'])} arch.</span>
      <div class="sbar-wrap"><div class="sbar-outer"><div class="sbar-inner bar-{pm_c}" style="width:{pm_w:.0f}%"></div></div><span class="sbar-val">{s['pct_pm']:.0f}% post-mig</span></div>
      <span style="font-family:var(--font-mono);font-size:11px;color:var(--ink-faint)">{s['ultima']}</span>
      {pill(s['rec'], s['rec_cls'])}
    </div>
  </div>
  <div style="display:flex;flex-direction:column;align-items:flex-end;gap:6px">{lnk}{pill('automático','accent')}</div>
</div>"""

    n_grandes_auto = len([s for s in grandes if s.get("tipo_inf") not in ("prioritario","prioritario_faltante")])
    gb_grandes_auto = sum(s["gb"] for s in grandes if s.get("tipo_inf") not in ("prioritario","prioritario_faltante"))

    sec_02 = f"""<div class="section" id="principales">
  {sec_header("02", f"Sitios principales — ≥ {gb_principal:.0f} GB (automáticos)", f"{n_grandes_auto} sitios · {fmt_gb(gb_grandes_auto)}", "accent", "principales")}
  <div style="display:flex;flex-direction:column;gap:0">{site_cards or '<p style="color:var(--ink-faint);font-style:italic">Todos los sitios principales tienen informe prioritario.</p>'}</div>
</div>"""

    # ── 03 Medianos ──────────────────────────────────────────
    rows_med = ""
    for s in medianos:
        pm_c = "green" if s["pct_pm"] >= 30 else ("red" if s["pct_pm"] == 0 else "gold")
        pm_w = min(s["pct_pm"], 100)
        lnk  = link_auto(s["informe"]) if s.get("tipo_inf") == "automatico" and s["informe"] else \
               (link_prio(s["informe"]) if s.get("tipo_inf") == "prioritario" and s["informe"] else link_sin())
        rows_med += f"""<tr>
  <td class="bold">{h(s['sitio'])}</td>
  <td class="mono td-right">{fmt_gb(s['gb'])}</td>
  <td class="mono td-right">{fmt_num(s['archivos'])}</td>
  <td><div class="sbar-wrap"><div class="sbar-outer"><div class="sbar-inner bar-{pm_c}" style="width:{pm_w:.0f}%"></div></div><span class="sbar-val">{s['pct_pm']:.0f}%</span></div></td>
  <td class="mono">{s['ultima']}</td>
  <td>{pill(s['rec'], s['rec_cls'])}</td>
  <td>{lnk}</td>
</tr>"""

    sec_03 = f"""<div class="section" id="medianos">
  {sec_header("03", f"Sitios medianos — {gb_mediano:.0f} GB a {gb_principal:.0f} GB", f"{len(medianos)} sitios · {fmt_gb(gb_medianos)}", "accent", "medianos")}
  <div class="table-wrap"><table>
    <thead><tr><th>Sitio</th><th class="td-right">Storage</th><th class="td-right">Archivos</th><th>% Post-Mig</th><th>Última Activ.</th><th>Estado</th><th>Informe</th></tr></thead>
    <tbody>{rows_med or '<tr><td colspan="7" style="color:var(--ink-faint);font-style:italic;text-align:center">Sin sitios en este rango</td></tr>'}</tbody>
  </table></div>
</div>"""

    # ── 04 Menores ───────────────────────────────────────────
    rows_men = ""
    for s in menores:
        pm_c = "green" if s["pct_pm"] >= 30 else ("red" if s["pct_pm"] == 0 else "gold")
        lnk  = link_auto(s["informe"]) if s["informe"] else link_sin()
        rows_men += f"""<tr>
  <td class="bold">{h(s['sitio'])}</td>
  <td class="mono td-right">{fmt_gb(s['gb'])}</td>
  <td class="mono td-right">{fmt_num(s['archivos'])}</td>
  <td class="mono td-center" style="color:var(--{pm_c})">{s['pct_pm']:.0f}%</td>
  <td class="mono">{s['ultima']}</td>
  <td>{pill(s['rec'], s['rec_cls'])}</td>
  <td>{lnk}</td>
</tr>"""

    sec_04 = f"""<div class="section" id="menores">
  {sec_header("04", f"Sitios menores — menos de {gb_mediano:.0f} GB", f"{len(menores)} sitios · {fmt_gb(gb_menores)}", "gold", "menores")}
  <div class="callout callout-gold"><span class="callout-label">Acción sugerida</span>La mayoría son candidatos directos a archivo por su bajo volumen. Revisar informe solo si el nombre sugiere contenido estratégico o regulatorio.</div>
  <div class="table-wrap"><table>
    <thead><tr><th>Sitio</th><th class="td-right">Storage</th><th class="td-right">Archivos</th><th class="td-center">% Post-Mig</th><th>Última Activ.</th><th>Estado</th><th>Informe</th></tr></thead>
    <tbody>{rows_men or '<tr><td colspan="7" style="color:var(--ink-faint);font-style:italic;text-align:center">Sin sitios en este rango</td></tr>'}</tbody>
  </table></div>
</div>"""

    # ── 05 Candidatos a archivo ──────────────────────────────
    rows_arch = ""
    for s in [x for x in sitios if x["rec"] == "Archivar"]:
        tipo_b = pill("★ prioritario","orange") if s.get("tipo_inf")=="prioritario" else pill("automático","accent")
        lnk    = (link_prio(s["informe"]) if s.get("tipo_inf")=="prioritario" else link_auto(s["informe"])) \
                 if s["informe"] else link_sin()
        grupo  = f"≥{gb_principal:.0f} GB" if s["gb"] >= gb_principal else \
                 (f"{gb_mediano:.0f}–{gb_principal:.0f} GB" if s["gb"] >= gb_mediano else f"<{gb_mediano:.0f} GB")
        rows_arch += f"""<tr>
  <td class="bold">{h(s['sitio'])}</td>
  <td class="mono td-right">{fmt_gb(s['gb'])}</td>
  <td class="mono td-right">{fmt_num(s['archivos'])}</td>
  <td class="mono">{s['ultima']}</td>
  <td>{pill(grupo,'ghost')}</td>
  <td>{tipo_b}</td>
  <td>{lnk}</td>
</tr>"""

    sec_05 = f"""<div class="section" id="archivar">
  {sec_header("05", "Candidatos a archivo", f"{n_archivar} sitios · {fmt_gb(gb_arch)}", "red", "archivar")}
  <div class="callout callout-red">
    <span class="callout-label">Proceso de archivo</span>
    Ningún archivo en estos sitios fue modificado después del {config['fecha_migracion'].strftime('%d/%m/%Y')}.
    Pasos: (1) revisar el informe y confirmar contenido migrado, (2) desasociar el equipo de Teams,
    (3) archivar desde el Admin Center. <strong>No eliminar</strong> — conservar mínimo 1 año como referencia histórica.
  </div>
  <div class="table-wrap"><table>
    <thead><tr><th>Sitio</th><th class="td-right">Storage</th><th class="td-right">Archivos</th><th>Última Actividad</th><th>Grupo</th><th>Tipo informe</th><th>Informe</th></tr></thead>
    <tbody>{rows_arch or '<tr><td colspan="7" style="color:var(--ink-faint);font-style:italic;text-align:center">Sin candidatos a archivo — todos los sitios tienen actividad post-migración.</td></tr>'}</tbody>
  </table></div>
</div>"""

    # ── Ensamblar ────────────────────────────────────────────
    body = (
        topbar +
        '<div class="layout">' + sidebar +
        '<div class="main">' +
        hero + total_bar + sum_cards +
        '<hr class="divider">' +
        sec_01 + '<hr class="divider">' +
        sec_02 + '<hr class="divider">' +
        sec_03 + '<hr class="divider">' +
        sec_04 + '<hr class="divider">' +
        sec_05 +
        build_footer(config["nombre"], "Índice de Relevamiento Legacy") +
        '</div></div>'
    )

    from m365_curaduria import css as base_css, GOOGLE_FONTS, SIDEBAR_JS
    return (
        f'<!DOCTYPE html>\n<html lang="es">\n<head>\n'
        f'<meta charset="UTF-8">\n'
        f'<meta name="viewport" content="width=device-width, initial-scale=1.0">\n'
        f'<title>{h(config["nombre"])} · Índice de Relevamiento</title>\n'
        f'{GOOGLE_FONTS}\n'
        f'<style>{base_css(config)}\n{CSS_INDICE}</style>\n'
        f'</head>\n<body>\n{body}\n{SIDEBAR_JS}\n</body>\n</html>'
    )


# ===========================================================================
# MAIN
# ===========================================================================

def main():
    args = parse_args("Genera el índice maestro con links a todos los informes de sitio.")

    print("=" * 60)
    print("  Kit M365 Curaduría — Índice maestro")
    print("=" * 60)

    config = load_config(args.config)
    print(f"  Cliente : {config['nombre']}")
    print(f"  Carpeta : {config['carpeta_local']}")

    df_raw = load_csvs(config)
    df     = clean_df(df_raw, config)

    carpeta_informes = get_carpeta_informes(config)
    prefijo_cliente  = config["nombre"].replace(" ", "_").replace(".", "_")

    # Verificar informes prioritarios
    print(f"\nInformes prioritarios configurados: {len(config['informes_prioritarios'])}")
    for p in config["informes_prioritarios"]:
        fname  = p.get("archivo_html", "")
        existe = fname and os.path.exists(os.path.join(carpeta_informes, fname))
        estado = "OK" if existe else "XX falta"
        print(f"  {estado}  {fname or '(sin nombre)'}")

    n_auto = len(glob.glob(os.path.join(carpeta_informes, f"{prefijo_cliente}_Informe_*.html")))
    print(f"\nInformes automáticos en carpeta: {n_auto}")
    print(f"Sitios en CSVs: {df['Sitio'].nunique()}")

    ts       = datetime.now()
    ts_label = ts.strftime("%d/%m/%Y %H:%M")

    print("\nGenerando índice...")
    html = build_html(df, config, carpeta_informes, ts_label)

    out_path = os.path.join(carpeta_informes, f"{prefijo_cliente}_Indice_Informes.html")
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)

    kb = os.path.getsize(out_path) / 1024
    print(f"\nÍndice generado: {out_path}")
    print(f"Tamaño: {kb:.0f} KB")


if __name__ == "__main__":
    main()
