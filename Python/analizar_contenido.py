"""
===============================================================================
analizar_contenido.py  |  Circo Studio  |  Kit de Curaduría Informacional M365
-------------------------------------------------------------------------------
Análisis consolidado de actividad post-migración.
Lee todos los CSVs de relevamiento y genera un HTML con 8 secciones:
  01 Por sitio          — estado, storage, % post-migración, recomendación
  02 Post-migración     — archivos modificados después de la fecha de referencia
  03 Candidatos archivo — sitios sin actividad post-migración
  04 Sin actividad Ny   — archivos no modificados en N años
  05 Por categoría      — distribución por tipo de archivo
  06 Top usuarios       — owners reales del contenido
  07 Archivos bloqueados — checked out, bloquean migración SPMT
  08 Matriz destino      — tabla para completar con el cliente

Uso:
    python analizar_contenido.py --config "..\\..\\Config\\cliente_config.json"

Genera:
    [CarpetaLocal]\\[Cliente]_Analisis_Contenido_YYYYMMDD_HHMMSS.html
===============================================================================
"""

import os
import sys
import pandas as pd
from datetime import datetime

# Módulo compartido — debe estar en la misma carpeta
sys.path.insert(0, os.path.dirname(__file__))
from m365_curaduria import (
    load_config, parse_args, load_csvs, clean_df,
    get_output_path, rec_info,
    h, pill, fmt_gb, fmt_num, bar, sec_header,
    build_topbar, build_sidebar, build_footer, html_doc,
)


# ===========================================================================
# CONSTRUCCIÓN HTML
# ===========================================================================

def build_html(df: pd.DataFrame, config: dict, ts_label: str) -> str:

    fecha_mig = config["fecha_migracion"]
    fecha_str = fecha_mig.strftime("%d/%m/%Y")

    # ── Métricas globales ────────────────────────────────────
    total_archivos  = len(df)
    total_gb        = df["Tamaño_Bytes"].sum() / 1e9
    total_sitios    = df["Sitio"].nunique()
    post_mig        = int(df["Post_Migracion"].sum())
    pct_post        = post_mig / total_archivos * 100 if total_archivos else 0
    sin_act         = int(df["Sin_Actividad_Ny"].sum())
    bloqueados      = int((df.get("Archivo_Bloqueado", pd.Series()) == "Sí").sum())
    usuarios_unicos = df["Modificado_Por"].nunique()

    # ── Datos por sitio ──────────────────────────────────────
    sitios_data = []
    for (sitio, url), g in df.groupby(["Sitio", "URL_Sitio"]):
        n  = len(g)
        gb = g["Tamaño_Bytes"].sum() / 1e9
        pm = int(g["Post_Migracion"].sum())
        pp = pm / n * 100 if n else 0
        ul = g["Fecha_Modificacion_dt"].max()
        rec, rec_cls, rec_col = rec_info(pm, pp)
        sitios_data.append({
            "sitio": sitio, "url": url, "archivos": n, "gb": gb,
            "post_mig": pm, "pct_pm": pp,
            "ultima": ul.strftime("%Y-%m-%d") if pd.notna(ul) else "—",
            "usuarios": g["Modificado_Por"].nunique(),
            "cat_top": g["Categoria"].value_counts().index[0] if n > 0 else "",
            "bibs": g["Biblioteca"].nunique(),
            "rec": rec, "rec_cls": rec_cls, "rec_col": rec_col
        })
    sitios_data.sort(key=lambda x: x["gb"], reverse=True)

    sitios_sin = [s for s in sitios_data if s["rec"] == "Archivar"]
    gb_arch    = sum(s["gb"] for s in sitios_sin)

    cat_stats = df.groupby("Categoria").agg(
        archivos=("Nombre_Archivo", "count"),
        gb=("Tamaño_Bytes", lambda x: x.sum() / 1e9)
    ).sort_values("gb", ascending=False)

    usr_stats = df.groupby("Modificado_Por").agg(
        archivos=("Nombre_Archivo", "count"),
        gb=("Tamaño_Bytes", lambda x: round(x.sum() / 1e9, 2)),
        sitios=("Sitio", "nunique"),
        post_mig=("Post_Migracion", "sum"),
        ultimo=("Fecha_Modificacion_dt", "max")
    ).sort_values("archivos", ascending=False).head(30).reset_index()

    df_bloq  = df[df.get("Archivo_Bloqueado", pd.Series(dtype=str)) == "Sí"].sort_values("Sitio")
    df_inact = df[df["Sin_Actividad_Ny"]].sort_values("Fecha_Modificacion_dt")
    df_post  = (df[df["Post_Migracion"]]
                .sort_values(["Sitio", "Fecha_Modificacion_dt"], ascending=[True, False])
                .head(300))

    # ── Topbar ───────────────────────────────────────────────
    topbar = build_topbar(
        label=f"{h(config['nombre'])} · ANÁLISIS DE CONTENIDO",
        title="Actividad post-migración y candidatos a archivo",
        badge_text="ANÁLISIS",
        meta=f"{config['sharepoint_url']} · {ts_label[:10]}"
    )

    # ── Sidebar ──────────────────────────────────────────────
    sidebar = build_sidebar(
        sections=[
            {"label": "Documento", "items": [
                ("00", "Resumen ejecutivo", "resumen"),
            ]},
            {"label": "Análisis", "items": [
                ("01", "Por sitio",           "por-sitio"),
                ("02", "Post-migración",      "post-migracion"),
                ("03", "Candidatos a archivo","candidatos-archivo"),
                ("04", f"Sin actividad +{config['inactividad_anios']}a", "sin-actividad"),
                ("05", "Por categoría",       "categorias"),
                ("06", "Top usuarios",        "usuarios"),
                ("07", "Archivos bloqueados", "bloqueados"),
                ("08", "Matriz destino",      "matriz-destino"),
            ]},
        ],
        stats=[
            f"{total_sitios} sitios",
            f"{fmt_num(total_archivos)} archivos",
            fmt_gb(total_gb),
        ]
    )

    # ── Hero ─────────────────────────────────────────────────
    hero = f"""<div class="hero" id="resumen">
  <div class="hero-eyebrow">{h(config['nombre'])} · ANÁLISIS DE CONTENIDO LEGACY</div>
  <h1>Relevamiento completo — actividad post-migración</h1>
  <div class="hero-sub">
    Inventario consolidado. Fecha de referencia de migración:
    <strong>{fecha_str}</strong>.
    Los sitios sin actividad posterior son candidatos directos a archivo.
  </div>
  <div class="hero-stats">
    <div class="stat"><span class="stat-value ink">{total_sitios}</span><span class="stat-label">Sitios analizados</span></div>
    <div class="stat"><span class="stat-value orange">{fmt_num(total_archivos)}</span><span class="stat-label">Archivos totales</span></div>
    <div class="stat"><span class="stat-value blue">{fmt_gb(total_gb)}</span><span class="stat-label">Volumen total</span></div>
    <div class="stat"><span class="stat-value green">{fmt_num(post_mig)}</span><span class="stat-label">Modif. post-migración</span></div>
    <div class="stat"><span class="stat-value red">{len(sitios_sin)}</span><span class="stat-label">Sitios a archivar</span></div>
    <div class="stat"><span class="stat-value gold">{fmt_num(sin_act)}</span><span class="stat-label">Archivos +{config['inactividad_anios']}a sin uso</span></div>
    <div class="stat"><span class="stat-value accent">{usuarios_unicos}</span><span class="stat-label">Usuarios activos</span></div>
    <div class="stat"><span class="stat-value ink">{bloqueados}</span><span class="stat-label">Archivos bloqueados</span></div>
  </div>
</div>
<div class="summary-grid">
  <div class="sum-card orange"><div class="sc-val">{len([s for s in sitios_data if s['rec']=='Activo — revisar destino'])}</div><div class="sc-label">Sitios activos post-migración — revisar destino</div></div>
  <div class="sum-card red"><div class="sc-val">{len(sitios_sin)}</div><div class="sc-label">Sin actividad — candidatos directos a archivo</div></div>
  <div class="sum-card gold"><div class="sc-val">{pct_post:.0f}%</div><div class="sc-label">Del total modificado post-migración</div></div>
  <div class="sum-card accent"><div class="sc-val">{fmt_gb(gb_arch)}</div><div class="sc-label">Recuperables archivando sitios inactivos</div></div>
</div>"""

    # ── 01 Por sitio ─────────────────────────────────────────
    rows_sitio = ""
    for s in sitios_data:
        pw = min(s["pct_pm"], 100)
        bc = "green" if s["pct_pm"] >= 30 else ("red" if s["pct_pm"] == 0 else "gold")
        rc_row = "row-critical" if s["rec"] == "Archivar" else ("row-ok" if s["pct_pm"] >= 30 else "row-warning")
        rows_sitio += f"""<tr class="{rc_row}">
  <td class="bold">{h(s['sitio'])}</td>
  <td class="mono td-right">{fmt_gb(s['gb'])}</td>
  <td class="mono td-right">{fmt_num(s['archivos'])}</td>
  <td class="mono td-center">{s['bibs']}</td>
  <td><div class="sbar-wrap"><div class="sbar-outer"><div class="sbar-inner bar-{bc}" style="width:{pw:.0f}%"></div></div><span class="sbar-val">{s['pct_pm']:.0f}%</span></div></td>
  <td class="mono">{h(s['ultima'])}</td>
  <td class="mono td-center">{s['usuarios']}</td>
  <td>{pill(s['cat_top'], 'ghost')}</td>
  <td>{pill(s['rec'], s['rec_cls'])}</td>
</tr>"""

    sec_01 = f"""<div class="section" id="por-sitio">
  {sec_header("01", "Análisis por sitio", f"{total_sitios} sitios", "accent", "por-sitio")}
  <div class="callout callout-accent">
    <span class="callout-label">Criterio de recomendación</span>
    <strong>Archivar</strong> — 0 archivos modificados después del {fecha_str} &nbsp;·&nbsp;
    <strong>Revisar</strong> — actividad post-migración menor al 30% &nbsp;·&nbsp;
    <strong>Activo</strong> — actividad igual o mayor al 30%
  </div>
  <div class="table-wrap"><table>
    <thead><tr>
      <th>Sitio</th><th class="td-right">Storage</th><th class="td-right">Archivos</th>
      <th class="td-center">Bibls.</th><th>% Post-Mig</th><th>Última Modif.</th>
      <th class="td-center">Usuarios</th><th>Cat. Principal</th><th>Recomendación</th>
    </tr></thead>
    <tbody>{rows_sitio}</tbody>
  </table></div>
</div>"""

    # ── 02 Post-migración ────────────────────────────────────
    nota_post = f" (mostrando {len(df_post)} de {fmt_num(int(df['Post_Migracion'].sum()))})" \
                if int(df["Post_Migracion"].sum()) > 300 else ""
    rows_post = ""
    for _, r in df_post.iterrows():
        rows_post += f"""<tr>
  <td class="mono" style="font-size:11px">{h(r['Sitio'])}</td>
  <td class="mono" style="font-size:11px">{h(str(r['Biblioteca']))}</td>
  <td style="font-size:12px">{h(str(r['Nombre_Archivo']))}</td>
  <td>{pill(str(r['Categoria']), 'ghost')}</td>
  <td class="mono td-right">{float(r['Tamaño_MB']):.2f} MB</td>
  <td class="mono">{h(str(r['Fecha_Modificacion']))}</td>
  <td class="mono" style="font-size:11px">{h(str(r['Modificado_Por']))}</td>
</tr>"""

    sec_02 = f"""<div class="section" id="post-migracion">
  {sec_header("02", "Archivos modificados post-migración", f"{fmt_num(post_mig)} archivos{nota_post}", "green", "post-migracion")}
  <div class="callout callout-green">
    <span class="callout-label">Implicancia</span>
    Archivos creados o editados después del {fecha_str}.
    Confirmar que estos archivos ya se encuentran en el hub destino antes de archivar el sitio.
  </div>
  <div class="table-wrap"><table>
    <thead><tr>
      <th>Sitio</th><th>Biblioteca</th><th>Archivo</th><th>Categoría</th>
      <th class="td-right">Tamaño</th><th>Fecha Modif.</th><th>Modificado Por</th>
    </tr></thead>
    <tbody>{rows_post}</tbody>
  </table></div>
</div>"""

    # ── 03 Candidatos a archivo ──────────────────────────────
    bars_arch = ""
    max_gb_arch = max((s["gb"] for s in sitios_sin), default=1)
    for s in sorted(sitios_sin, key=lambda x: x["gb"], reverse=True):
        pct = min(s["gb"] / max_gb_arch * 100, 100)
        bars_arch += f"""<div class="storage-bar-item">
  <div class="sbi-header">
    <span class="sbi-name">{h(s['sitio'])}</span>
    <span class="sbi-val">{fmt_gb(s['gb'])}</span>
  </div>
  <div class="sbi-bar-outer"><div class="sbi-bar-inner bar-red" style="width:{pct:.1f}%"></div></div>
  <div class="sbi-meta">Última actividad: {s['ultima']} · {fmt_num(s['archivos'])} archivos · {s['usuarios']} usuarios históricos</div>
</div>"""

    contenido_03 = (
        f'<div class="storage-bars">{bars_arch}</div>'
        if bars_arch else
        '<div class="callout callout-green"><span class="callout-label">Sin sitios inactivos</span>Todos los sitios tienen actividad post-migración.</div>'
    )

    sec_03 = f"""<div class="section" id="candidatos-archivo">
  {sec_header("03", "Candidatos a archivo — sin actividad post-migración", f"{len(sitios_sin)} sitios · {fmt_gb(gb_arch)}", "red", "candidatos-archivo")}
  <div class="callout callout-red">
    <span class="callout-label">Acción requerida</span>
    Ningún archivo en estos sitios fue modificado después del {fecha_str}.
    Confirmar migración al hub, desasociar el equipo de Teams y archivar desde el Admin Center.
  </div>
  {contenido_03}
</div>"""

    # ── 04 Sin actividad Ny ──────────────────────────────────
    nota_inact = f" (mostrando 150 de {fmt_num(len(df_inact))})" if len(df_inact) > 150 else ""
    rows_inact = ""
    for _, r in df_inact.head(150).iterrows():
        rows_inact += f"""<tr>
  <td class="mono" style="font-size:11px">{h(r['Sitio'])}</td>
  <td style="font-size:12px">{h(str(r['Nombre_Archivo']))}</td>
  <td>{pill(str(r['Categoria']), 'ghost')}</td>
  <td class="mono td-right">{float(r['Tamaño_MB']):.2f} MB</td>
  <td class="mono">{h(str(r['Fecha_Modificacion']))}</td>
  <td class="mono" style="font-size:11px">{h(str(r['Creado_Por']))}</td>
</tr>"""

    sec_04 = f"""<div class="section" id="sin-actividad">
  {sec_header("04", f"Sin actividad desde antes del {config['fecha_inactividad'].strftime('%d/%m/%Y')} (+{config['inactividad_anios']} años)", f"{fmt_num(sin_act)} archivos{nota_inact}", "gold", "sin-actividad")}
  <div class="callout callout-gold">
    <span class="callout-label">Candidatos a eliminación</span>
    Revisar con los owners de negocio antes de eliminar — pueden ser documentos de referencia que no se editan pero sí se consultan.
  </div>
  <div class="table-wrap"><table>
    <thead><tr>
      <th>Sitio</th><th>Archivo</th><th>Categoría</th><th class="td-right">Tamaño</th>
      <th>Última Modif.</th><th>Creado Por</th>
    </tr></thead>
    <tbody>{rows_inact}</tbody>
  </table></div>
</div>"""

    # ── 05 Categorías ────────────────────────────────────────
    rows_cat = ""
    for cat_name, cr in cat_stats.iterrows():
        pa = cr["archivos"] / total_archivos * 100 if total_archivos else 0
        pg = cr["gb"] / total_gb * 100 if total_gb else 0
        rows_cat += f"""<tr>
  <td class="bold">{h(cat_name)}</td>
  <td class="mono td-right">{fmt_num(int(cr['archivos']))}</td>
  <td class="mono td-right">{pa:.1f}%</td>
  <td><div class="sbar-wrap"><div class="sbar-outer"><div class="sbar-inner bar-blue" style="width:{min(pg,100):.0f}%"></div></div><span class="sbar-val">{fmt_gb(cr['gb'])}</span></div></td>
  <td class="mono td-right">{pg:.1f}%</td>
</tr>"""

    sec_05 = f"""<div class="section" id="categorias">
  {sec_header("05", "Volumen por categoría de archivo", f"{cat_stats.shape[0]} categorías", "blue", "categorias")}
  <div class="table-wrap"><table>
    <thead><tr>
      <th>Categoría</th><th class="td-right">Archivos</th><th class="td-right">% Archivos</th>
      <th>Volumen</th><th class="td-right">% Volumen</th>
    </tr></thead>
    <tbody>{rows_cat}</tbody>
  </table></div>
</div>"""

    # ── 06 Top usuarios ──────────────────────────────────────
    rows_usr = ""
    for _, ur in usr_stats.iterrows():
        ultimo = ur["ultimo"].strftime("%Y-%m-%d") if pd.notna(ur["ultimo"]) else "—"
        rows_usr += f"""<tr>
  <td class="bold">{h(str(ur['Modificado_Por']))}</td>
  <td class="mono td-right">{fmt_num(int(ur['archivos']))}</td>
  <td class="mono td-right">{fmt_gb(float(ur['gb']))}</td>
  <td class="mono td-center">{int(ur['sitios'])}</td>
  <td class="mono td-right">{fmt_num(int(ur['post_mig']))}</td>
  <td class="mono">{h(ultimo)}</td>
</tr>"""

    sec_06 = f"""<div class="section" id="usuarios">
  {sec_header("06", "Top usuarios por volumen de contenido", f"{usuarios_unicos} usuarios únicos", "accent", "usuarios")}
  <div class="callout callout-accent">
    <span class="callout-label">Para qué sirve esto</span>
    Los usuarios con mayor volumen modificado son los <strong>owners reales</strong> de cada sitio.
    Usarlos para asignar responsables en la Matriz de Destino.
  </div>
  <div class="table-wrap"><table>
    <thead><tr>
      <th>Usuario</th><th class="td-right">Archivos Modif.</th><th class="td-right">GB</th>
      <th class="td-center">Sitios</th><th class="td-right">Post-Mig</th><th>Último Archivo</th>
    </tr></thead>
    <tbody>{rows_usr}</tbody>
  </table></div>
</div>"""

    # ── 07 Bloqueados ────────────────────────────────────────
    if len(df_bloq) == 0:
        contenido_bloq = '<div class="callout callout-green"><span class="callout-label">Sin bloqueos</span>No se encontraron archivos con checkout activo.</div>'
    else:
        rows_bloq = ""
        for _, r in df_bloq.iterrows():
            rows_bloq += f"""<tr>
  <td class="mono" style="font-size:11px">{h(r['Sitio'])}</td>
  <td style="font-size:12px">{h(str(r['Nombre_Archivo']))}</td>
  <td>{pill(str(r['Categoria']), 'ghost')}</td>
  <td class="mono td-right">{float(r['Tamaño_MB']):.2f} MB</td>
  <td class="mono">{h(str(r['Fecha_Modificacion']))}</td>
  <td class="mono" style="font-size:11px">{h(str(r['Modificado_Por']))}</td>
</tr>"""
        contenido_bloq = f"""<div class="callout callout-orange">
  <span class="callout-label">Acción requerida antes de migrar</span>
  Archivos con checkout activo no pueden ser migrados con SPMT.
  Contactar a los responsables para que realicen el check-in.
</div>
<div class="table-wrap"><table>
  <thead><tr>
    <th>Sitio</th><th>Archivo</th><th>Categoría</th><th class="td-right">Tamaño</th>
    <th>Última Modif.</th><th>Responsable</th>
  </tr></thead>
  <tbody>{rows_bloq}</tbody>
</table></div>"""

    sec_07 = f"""<div class="section" id="bloqueados">
  {sec_header("07", "Archivos bloqueados (checked out)", f"{bloqueados} archivos", "orange" if bloqueados > 0 else "green", "bloqueados")}
  {contenido_bloq}
</div>"""

    # ── 08 Matriz destino ────────────────────────────────────
    rows_mtz = ""
    for s in sitios_data:
        rows_mtz += f"""<tr>
  <td class="bold" style="font-size:12px">{h(s['sitio'])}</td>
  <td class="mono td-right">{fmt_gb(s['gb'])}</td>
  <td class="mono td-right">{fmt_num(s['archivos'])}</td>
  <td>{pill(s['rec'], s['rec_cls'])}</td>
  <td style="background:rgba(255,255,180,0.5);color:var(--ink-faint);font-size:11px;font-style:italic">— completar —</td>
  <td style="background:rgba(255,255,180,0.5);color:var(--ink-faint);font-size:11px;font-style:italic">— completar —</td>
  <td style="background:rgba(255,255,180,0.5);color:var(--ink-faint);font-size:11px;font-style:italic">— completar —</td>
</tr>"""

    sec_08 = f"""<div class="section" id="matriz-destino">
  {sec_header("08", "Matriz de destino de migración", "Para completar con el cliente", "gold", "matriz-destino")}
  <div class="callout callout-gold">
    <span class="callout-label">Instrucciones</span>
    Las columnas en amarillo se completan junto con los responsables del cliente.
    <strong>Destino Hub</strong>: sitio destino en la arquitectura gobernada.
    <strong>Subcarpeta</strong>: carpeta dentro del sitio destino.
    <strong>Responsable</strong>: owner del cliente que valida y aprueba la migración.
  </div>
  <div class="table-wrap"><table>
    <thead><tr>
      <th>Sitio</th><th class="td-right">Storage</th><th class="td-right">Archivos</th>
      <th>Estado</th><th>Destino Hub</th><th>Subcarpeta</th><th>Responsable</th>
    </tr></thead>
    <tbody>{rows_mtz}</tbody>
  </table></div>
</div>"""

    # ── Ensamblar ────────────────────────────────────────────
    body = (
        topbar +
        '<div class="layout">' +
        sidebar +
        '<div class="main">' +
        hero +
        '<hr class="divider">' +
        sec_01 + sec_02 + sec_03 + sec_04 +
        sec_05 + sec_06 + sec_07 + sec_08 +
        build_footer(config["nombre"], "Análisis de Contenido") +
        '</div></div>'
    )

    return html_doc(
        title=f"{config['nombre']} · Análisis de Contenido",
        config=config,
        body=body
    )


# ===========================================================================
# MAIN
# ===========================================================================

def main():
    args = parse_args("Genera el análisis consolidado de contenido post-migración.")

    print("=" * 60)
    print("  Kit M365 Curaduría — Análisis de Contenido")
    print("=" * 60)

    config = load_config(args.config)
    print(f"  Cliente  : {config['nombre']}")
    print(f"  Migración: {config['fecha_migracion'].strftime('%d/%m/%Y')}")
    print(f"  Carpeta  : {config['carpeta_local']}")
    print()

    df_raw = load_csvs(config)
    df     = clean_df(df_raw, config)
    print(f"\nFilas tras limpieza: {len(df):,}")

    ts       = datetime.now()
    ts_str   = ts.strftime("%Y%m%d_%H%M%S")
    ts_label = ts.strftime("%d/%m/%Y %H:%M")

    print("\nGenerando HTML...")
    html = build_html(df, config, ts_label)

    out_path = get_output_path(config, "Analisis_Contenido", ts_str)
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)

    kb = os.path.getsize(out_path) / 1024
    print(f"\nInforme generado: {out_path}")
    print(f"Tamaño: {kb:.0f} KB")


if __name__ == "__main__":
    main()
