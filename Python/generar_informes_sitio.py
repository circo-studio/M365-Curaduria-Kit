"""
===============================================================================
generar_informes_sitio.py  |  Circo Studio  |  Kit de Curaduría M365
-------------------------------------------------------------------------------
Genera un HTML de análisis profundo por cada sitio del corpus.

Secciones por informe:
  01 Estructura de carpetas    07 Post-migración
  02 Categorías y formatos     08 Archivos grandes
  03 Formatos especiales       09 Destino de migración (placeholder)
  04 Duplicados                10 Decisiones pendientes (automáticas)
  05 Actividad histórica
  06 Usuarios

Uso:
    python generar_informes_sitio.py --config "...\\Config\\cliente_config.json"
    python generar_informes_sitio.py --config "..." --sitio "GR - Nombre del Sitio"

Genera en [CarpetaLocal]\\Informes\\ :
    [Cliente]_Informe_[NombreLimpio]_[timestamp].html
===============================================================================
"""

import os
import sys
import re
import pandas as pd
from datetime import datetime
from collections import Counter

sys.path.insert(0, os.path.dirname(__file__))
from m365_curaduria import (
    load_config, parse_args, load_csvs, clean_df,
    get_carpeta_informes, rec_info, FORMATOS_ESPECIALES,
    h, pill, fmt_gb, fmt_num, bar, sec_header,
    build_topbar, build_sidebar, build_footer, html_doc,
)


# ===========================================================================
# ANÁLISIS POR SITIO
# ===========================================================================

def analizar(df: pd.DataFrame, config: dict) -> dict:
    a = {}
    a["total_archivos"] = len(df)
    a["total_bytes"]    = df["Tamaño_Bytes"].sum()
    a["total_gb"]       = a["total_bytes"] / 1e9
    a["bibliotecas"]    = df["Biblioteca"].nunique()
    a["usuarios"]       = df["Modificado_Por"].nunique()
    a["prof_max"]       = int(df["Profundidad_Carpeta"].max()) if "Profundidad_Carpeta" in df.columns else 0
    a["prof_media"]     = round(df["Profundidad_Carpeta"].mean(), 1) if "Profundidad_Carpeta" in df.columns else 0
    a["fecha_min"]      = df["Fecha_Modificacion_dt"].min()
    a["fecha_max"]      = df["Fecha_Modificacion_dt"].max()
    a["post_mig"]       = int(df["Post_Migracion"].sum())
    a["pct_post"]       = a["post_mig"] / a["total_archivos"] * 100 if a["total_archivos"] else 0
    a["sin_act"]        = int(df["Sin_Actividad_Ny"].sum())

    a["cat"] = df.groupby("Categoria").agg(
        n=("Nombre_Archivo", "count"),
        gb=("Tamaño_Bytes", lambda x: x.sum() / 1e9)
    ).sort_values("gb", ascending=False)

    a["ext"] = df.groupby("Extension").agg(
        n=("Nombre_Archivo", "count"),
        gb=("Tamaño_Bytes", lambda x: x.sum() / 1e9)
    ).sort_values("gb", ascending=False)

    a["especiales"] = {
        ext: (info, a["ext"].loc[ext]["n"], a["ext"].loc[ext]["gb"])
        for ext, info in FORMATOS_ESPECIALES.items()
        if ext in a["ext"].index
    }

    a["grandes"] = df[df["Tamaño_Bytes"] > config["mb_archivo_grande"] * 1024 * 1024].sort_values(
        "Tamaño_Bytes", ascending=False).head(20)

    a["zips_grandes"] = df[
        (df["Extension"] == "zip") & (df["Tamaño_Bytes"] > 250 * 1024 * 1024)
    ].sort_values("Tamaño_Bytes", ascending=False)

    # Duplicados por nombre + tamaño
    df_dup = df[df["Tamaño_Bytes"] > 0].copy()
    dup    = df_dup.groupby(["Nombre_Archivo", "Tamaño_Bytes"]).size().reset_index(name="copies")
    dup    = dup[dup["copies"] > 1]
    dup["bytes_extra"] = dup["Tamaño_Bytes"] * (dup["copies"] - 1)
    a["dup_archivos"]    = int(dup["copies"].sum() - len(dup))
    a["dup_gb"]          = dup["bytes_extra"].sum() / 1e9
    a["dup_top"]         = dup.sort_values("bytes_extra", ascending=False).head(15)

    df["ano"] = df["Fecha_Modificacion_dt"].dt.year
    a["por_ano"] = df.groupby("ano").agg(
        n=("Nombre_Archivo", "count"),
        gb=("Tamaño_Bytes", lambda x: x.sum() / 1e9)
    ).sort_index()

    a["top_usuarios"] = df.groupby("Modificado_Por").agg(
        n=("Nombre_Archivo", "count"),
        gb=("Tamaño_Bytes", lambda x: round(x.sum() / 1e9, 2)),
        post_mig=("Post_Migracion", "sum"),
        ultimo=("Fecha_Modificacion_dt", "max")
    ).sort_values("n", ascending=False).head(20).reset_index()

    def nivel1(carpeta):
        partes = [p for p in str(carpeta).split("/") if p not in ("Shared Documents", "Documents", "")]
        return partes[0] if partes else "(raíz)"

    df["carpeta_n1"] = df["Carpeta"].apply(nivel1)
    a["top_carpetas"] = df.groupby("carpeta_n1").agg(
        n=("Nombre_Archivo", "count"),
        gb=("Tamaño_Bytes", lambda x: x.sum() / 1e9)
    ).sort_values("gb", ascending=False).head(20)

    a["bloqueados"] = df[df.get("Archivo_Bloqueado", pd.Series(dtype=str)) == "Sí"]
    a["post_files"] = (df[df["Post_Migracion"]]
                       .sort_values("Fecha_Modificacion_dt", ascending=False).head(50))

    return a


# ===========================================================================
# CONSTRUCCIÓN HTML POR SITIO
# ===========================================================================

def build_html_sitio(df: pd.DataFrame, sitio_nombre: str, sitio_url: str,
                     config: dict, ts_label: str) -> str:

    a = analizar(df, config)
    fecha_mig = config["fecha_migracion"]
    fecha_str = fecha_mig.strftime("%d/%m/%Y")

    # Estado del sitio
    if a["post_mig"] == 0:
        estado_badge   = "CANDIDATO A ARCHIVO"
        estado_color   = "red"
        estado_callout = "callout-red"
        estado_nota    = f"Ningún archivo fue modificado después del {fecha_str}. Puede archivarse una vez confirmada la migración."
    elif a["pct_post"] >= 30:
        estado_badge   = "ACTIVO POST-MIGRACIÓN"
        estado_color   = "green"
        estado_callout = "callout-green"
        estado_nota    = f"{a['post_mig']:,} archivos ({a['pct_post']:.0f}%) modificados post-migración. Confirmar destino antes de archivar."
    else:
        estado_badge   = "ACTIVIDAD BAJA"
        estado_color   = "gold"
        estado_callout = "callout-gold"
        estado_nota    = f"{a['post_mig']:,} archivos ({a['pct_post']:.0f}%) modificados post-migración. Revisar con el área."

    topbar = build_topbar(
        label=f"{h(config['nombre'])} · RELEVAMIENTO",
        title=f"{h(sitio_nombre)} — Análisis previo a migración",
        badge_text="BORRADOR INTERNO",
        meta=f"{fmt_num(a['total_archivos'])} archivos · {fmt_gb(a['total_gb'])} · {ts_label[:10]}"
    )

    sidebar = build_sidebar(
        sections=[
            {"label": "Documento", "items": [("00", "Resumen ejecutivo", "resumen")]},
            {"label": "Contenido", "items": [
                ("01", "Estructura de carpetas", "estructura"),
                ("02", "Categorías y formatos",  "categorias"),
                ("03", "Formatos especiales",    "especiales"),
                ("04", "Duplicados",             "duplicados"),
                ("05", "Actividad histórica",    "actividad"),
                ("06", "Usuarios",               "usuarios"),
                ("07", "Post-migración",         "post-migracion"),
            ]},
            {"label": "Plan", "items": [
                ("08", "Archivos grandes",     "archivos-grandes"),
                ("09", "Destino migración",    "destino"),
                ("10", "Decisiones pendientes","decisiones"),
            ]},
        ],
        stats=[
            f"{fmt_num(a['total_archivos'])} archivos",
            fmt_gb(a["total_gb"]),
            f"{a['usuarios']} usuarios",
        ]
    )

    fecha_rango = ""
    if pd.notna(a["fecha_min"]) and pd.notna(a["fecha_max"]):
        fecha_rango = f"{a['fecha_min'].strftime('%Y-%m-%d')} → {a['fecha_max'].strftime('%Y-%m-%d')}"

    cat_dom     = a["cat"].index[0] if len(a["cat"]) > 0 else "—"
    cat_dom_pct = a["cat"].iloc[0]["n"] / a["total_archivos"] * 100 if a["total_archivos"] else 0

    hero = f"""<div class="hero" id="resumen">
  <div class="hero-eyebrow">Relevamiento · {h(sitio_nombre)}</div>
  <h1>Análisis previo a migración<br>{h(sitio_nombre)}</h1>
  <p class="hero-sub">
    {fmt_num(a["total_archivos"])} archivos · {fmt_gb(a["total_gb"])} en {a["bibliotecas"]} {'biblioteca' if a["bibliotecas"]==1 else 'bibliotecas'}.
    Profundidad máxima: <strong>{a["prof_max"]} niveles</strong>.
    Categoría dominante: <strong>{h(cat_dom)} ({cat_dom_pct:.0f}%)</strong>.
    Período: <strong>{fecha_rango}</strong>.
    {f'<strong>{a["pct_post"]:.0f}% del contenido fue modificado post-migración.</strong>' if a["post_mig"]>0 else '<strong>Sin actividad post-migración.</strong>'}
    {f'Duplicados estimados: <strong>{fmt_gb(a["dup_gb"])}</strong> ({a["dup_archivos"]:,} archivos).' if a["dup_gb"] > 0.5 else ''}
  </p>
  <div class="hero-stats">
    <div class="stat"><span class="stat-value accent">{fmt_num(a['total_archivos'])}</span><span class="stat-label">Archivos</span></div>
    <div class="stat"><span class="stat-value accent">{fmt_gb(a['total_gb'])}</span><span class="stat-label">Tamaño total</span></div>
    <div class="stat"><span class="stat-value {'green' if a['pct_post']>=30 else 'red' if a['pct_post']==0 else 'gold'}">{a['pct_post']:.0f}%</span><span class="stat-label">Post-migración</span></div>
    {f'<div class="stat"><span class="stat-value orange">{fmt_gb(a["dup_gb"])}</span><span class="stat-label">Duplicados est.</span></div>' if a["dup_gb"] > 0.1 else ''}
    {f'<div class="stat"><span class="stat-value gold">{fmt_num(a["sin_act"])}</span><span class="stat-label">Sin actividad +{config["inactividad_anios"]}a</span></div>' if a["sin_act"] > 0 else ''}
    <div class="stat"><span class="stat-value blue">{a['prof_max']}</span><span class="stat-label">Niveles profundidad</span></div>
    <div class="stat"><span class="stat-value ink">{a['usuarios']}</span><span class="stat-label">Usuarios</span></div>
    {f'<div class="stat"><span class="stat-value red">{len(a["bloqueados"])}</span><span class="stat-label">Bloqueados</span></div>' if len(a["bloqueados"])>0 else ''}
  </div>
</div>
<div class="{estado_callout} callout">
  <span class="callout-label">Estado del sitio — {estado_badge}</span>
  {estado_nota}
</div>"""

    # ── 01 Estructura ────────────────────────────────────────
    max_gb_carp = a["top_carpetas"]["gb"].max() if len(a["top_carpetas"]) > 0 else 1
    rows_carp = ""
    for carp, cr in a["top_carpetas"].iterrows():
        pct = cr["n"] / a["total_archivos"] * 100
        pw  = min(cr["gb"] / max_gb_carp * 100, 100)
        rows_carp += f"""<tr>
  <td class="bold mono">{h(str(carp))}</td>
  <td class="mono td-right">{fmt_num(int(cr['n']))}</td>
  <td class="td-right"><div class="sbar-wrap" style="justify-content:flex-end">
    <div class="sbar-outer"><div class="sbar-inner bar-accent" style="width:{pw:.0f}%"></div></div>
    <span class="sbar-val">{fmt_gb(cr['gb'])}</span>
  </div></td>
  <td class="mono td-right">{pct:.1f}%</td>
</tr>"""

    sec_01 = f"""<div class="section" id="estructura">
  {sec_header("01","Estructura de carpetas — nivel superior",f"Top {len(a['top_carpetas'])} carpetas","accent","estructura")}
  <div class="table-wrap"><table>
    <thead><tr><th>Carpeta (nivel 1)</th><th class="td-right">Archivos</th><th class="td-right">Volumen</th><th class="td-right">%</th></tr></thead>
    <tbody>{rows_carp}</tbody>
  </table></div>
</div>"""

    # ── 02 Categorías ────────────────────────────────────────
    rows_cat = ""
    for cat_n, cr in a["cat"].iterrows():
        pa = cr["n"] / a["total_archivos"] * 100
        pg = cr["gb"] / a["total_gb"] * 100 if a["total_gb"] else 0
        rows_cat += f"""<tr>
  <td class="bold">{h(cat_n)}</td>
  <td class="mono td-right">{fmt_num(int(cr['n']))}</td>
  <td>{bar(pa,'blue')}</td>
  <td class="mono td-right">{fmt_gb(cr['gb'])}</td>
  <td class="mono td-right">{pg:.1f}%</td>
</tr>"""

    rows_ext = ""
    for ext_n, er in a["ext"].head(15).iterrows():
        pg = er["gb"] / a["total_gb"] * 100 if a["total_gb"] else 0
        rows_ext += f"""<tr>
  <td class="mono bold">.{h(ext_n)}</td>
  <td class="mono td-right">{fmt_num(int(er['n']))}</td>
  <td class="mono td-right">{fmt_gb(er['gb'])}</td>
  <td class="mono td-right">{pg:.1f}%</td>
  <td>{''+pill("formato especial","gold") if ext_n in FORMATOS_ESPECIALES else ''}</td>
</tr>"""

    sec_02 = f"""<div class="section" id="categorias">
  {sec_header("02","Categorías y formatos",f"{a['cat'].shape[0]} categorías","blue","categorias")}
  <div class="table-wrap"><table>
    <thead><tr><th>Categoría</th><th class="td-right">Archivos</th><th>% Archivos</th><th class="td-right">Volumen</th><th class="td-right">% Volumen</th></tr></thead>
    <tbody>{rows_cat}</tbody>
  </table></div>
  <h3>Top extensiones</h3>
  <div class="table-wrap"><table>
    <thead><tr><th>Extensión</th><th class="td-right">Archivos</th><th class="td-right">Volumen</th><th class="td-right">% Total</th><th>Nota</th></tr></thead>
    <tbody>{rows_ext}</tbody>
  </table></div>
</div>"""

    # ── 03 Formatos especiales ───────────────────────────────
    if a["especiales"]:
        rows_esp = ""
        for ext, (info_tuple, n_ext, gb_ext) in a["especiales"].items():
            cat_label, nota = info_tuple
            pg = gb_ext / a["total_gb"] * 100 if a["total_gb"] else 0
            if ext in ("heic", "xlsm", "xlsb"):         alerta = pill("Atención requerida", "gold")
            elif ext in ("mp4", "mov", "avi"):           alerta = pill("Evaluar Azure Blob",  "gold")
            elif ext in ("zip","rar","7z") and gb_ext > 0.25: alerta = pill("Verificar límite SPMT", "red")
            else:                                        alerta = pill("Revisar", "blue")

            detalle_zips = ""
            if ext == "zip" and len(a["zips_grandes"]) > 0:
                detalle_zips = "<br><br><strong>ZIPs &gt;250 MB (límite SPMT):</strong><br>"
                for _, zr in a["zips_grandes"].iterrows():
                    detalle_zips += f'<code>{h(str(zr["Nombre_Archivo"]))}</code> — {fmt_gb(zr["Tamaño_Bytes"]/1e9)}<br>'

            rows_esp += f"""<tr>
  <td class="mono bold">.{h(ext)}</td>
  <td class="mono td-right">{fmt_num(int(n_ext))}</td>
  <td class="mono td-right">{fmt_gb(gb_ext)}</td>
  <td class="mono td-right">{pg:.1f}%</td>
  <td>{alerta}</td>
  <td style="font-size:12px">{h(nota)}{detalle_zips}</td>
</tr>"""

        sec_03 = f"""<div class="section" id="especiales">
  {sec_header("03","Formatos especiales",f"{len(a['especiales'])} formatos detectados","orange","especiales")}
  <div class="callout callout-orange">
    <span class="callout-label">Por qué importa</span>
    Archivos &gt;250 MB superan el límite de SPMT · Videos → evaluar Azure Blob · XLSM no ejecutan macros en SPO.
  </div>
  <div class="table-wrap"><table>
    <thead><tr><th>Formato</th><th class="td-right">Archivos</th><th class="td-right">Volumen</th><th class="td-right">% Total</th><th>Alerta</th><th>Nota</th></tr></thead>
    <tbody>{rows_esp}</tbody>
  </table></div>
</div>"""
    else:
        sec_03 = f"""<div class="section" id="especiales">
  {sec_header("03","Formatos especiales","Sin alertas","green","especiales")}
  <div class="callout callout-green"><span class="callout-label">Sin formatos problemáticos</span>No se detectaron formatos que requieran acción especial antes de migrar.</div>
</div>"""

    # ── 04 Duplicados ────────────────────────────────────────
    if a["dup_gb"] > 0.01:
        rows_dup = ""
        for _, dr in a["dup_top"].iterrows():
            rows_dup += f"""<tr>
  <td class="mono">{h(str(dr['Nombre_Archivo']))}</td>
  <td class="mono td-center">{int(dr['copies'])}</td>
  <td class="mono td-right">{fmt_gb(dr['Tamaño_Bytes']/1e9)}</td>
  <td class="mono td-right bold" style="color:var(--orange)">{fmt_gb(dr['bytes_extra']/1e9)}</td>
</tr>"""
        sec_04 = f"""<div class="section" id="duplicados">
  {sec_header("04","Duplicados estimados",f"{fmt_gb(a['dup_gb'])} recuperables","orange","duplicados")}
  <div class="callout callout-gold">
    <span class="callout-label">Metodología</span>
    Duplicados detectados por <strong>nombre + tamaño idénticos</strong>.
    {a['dup_archivos']:,} archivos extra → {fmt_gb(a['dup_gb'])} recuperables.
    Recomendación: migrar todo y deduplicar post-migración con validación del área.
  </div>
  <div class="table-wrap"><table>
    <thead><tr><th>Nombre del archivo</th><th class="td-center">Copias</th><th class="td-right">Tamaño c/u</th><th class="td-right">Espacio extra</th></tr></thead>
    <tbody>{rows_dup}</tbody>
  </table></div>
</div>"""
    else:
        sec_04 = f"""<div class="section" id="duplicados">
  {sec_header("04","Duplicados","Sin duplicados significativos","green","duplicados")}
  <div class="callout callout-green"><span class="callout-label">Sin duplicados detectados</span>No se encontraron archivos con nombre y tamaño idénticos en este sitio.</div>
</div>"""

    # ── 05 Actividad histórica ───────────────────────────────
    max_n_ano = a["por_ano"]["n"].max() if len(a["por_ano"]) > 0 else 1
    rows_ano = ""
    for ano, ar in a["por_ano"].iterrows():
        if pd.isna(ano): continue
        pw   = min(ar["n"] / max_n_ano * 100, 100)
        post = " ← post-migración" if int(ano) >= fecha_mig.year else ""
        bg   = 'style="background:var(--green-bg)"' if int(ano) >= fecha_mig.year and ar["n"] > 0 else ""
        bc   = "green" if int(ano) >= fecha_mig.year else "accent"
        rows_ano += f"""<tr {bg}>
  <td class="mono bold">{int(ano)}</td>
  <td><div class="sbar-wrap"><div class="sbar-outer" style="width:120px"><div class="sbar-inner bar-{bc}" style="width:{pw:.0f}%"></div></div><span class="sbar-val">{fmt_num(int(ar['n']))}</span></div></td>
  <td class="mono td-right">{fmt_gb(ar['gb'])}</td>
  <td style="font-size:11px;color:var(--ink-faint)">{post}</td>
</tr>"""

    sec_05 = f"""<div class="section" id="actividad">
  {sec_header("05","Actividad histórica",f"Rango: {a['fecha_min'].strftime('%Y') if pd.notna(a['fecha_min']) else '—'} – {a['fecha_max'].strftime('%Y') if pd.notna(a['fecha_max']) else '—'}","blue","actividad")}
  <div class="table-wrap"><table>
    <thead><tr><th>Año</th><th>Archivos modificados</th><th class="td-right">Volumen</th><th>Nota</th></tr></thead>
    <tbody>{rows_ano}</tbody>
  </table></div>
</div>"""

    # ── 06 Usuarios ──────────────────────────────────────────
    rows_usr = ""
    for _, ur in a["top_usuarios"].iterrows():
        ultimo = ur["ultimo"].strftime("%Y-%m-%d") if pd.notna(ur["ultimo"]) else "—"
        pct    = ur["n"] / a["total_archivos"] * 100
        rows_usr += f"""<tr>
  <td class="bold">{h(str(ur['Modificado_Por']))}</td>
  <td class="mono td-right">{fmt_num(int(ur['n']))}</td>
  <td class="mono td-right">{pct:.1f}%</td>
  <td class="mono td-right">{fmt_gb(float(ur['gb']))}</td>
  <td class="mono td-right">{fmt_num(int(ur['post_mig']))}</td>
  <td class="mono">{h(ultimo)}</td>
</tr>"""

    sec_06 = f"""<div class="section" id="usuarios">
  {sec_header("06","Usuarios",f"{a['usuarios']} usuarios únicos","accent","usuarios")}
  <div class="callout callout-accent">
    <span class="callout-label">Para la migración</span>
    El usuario con más archivos es el <strong>owner real</strong> del contenido.
    Verificar en Entra ID si algún usuario con gran volumen ya no está activo — SPMT puede tener problemas con autores dados de baja.
  </div>
  <div class="table-wrap"><table>
    <thead><tr><th>Usuario</th><th class="td-right">Archivos</th><th class="td-right">%</th><th class="td-right">GB</th><th class="td-right">Post-Mig</th><th>Último archivo</th></tr></thead>
    <tbody>{rows_usr}</tbody>
  </table></div>
</div>"""

    # ── 07 Post-migración ────────────────────────────────────
    if a["post_mig"] > 0:
        nota_pm  = f" (mostrando 50 de {a['post_mig']:,})" if a["post_mig"] > 50 else ""
        rows_pm  = ""
        for _, r in a["post_files"].iterrows():
            rows_pm += f"""<tr>
  <td style="font-size:12px">{h(str(r['Nombre_Archivo']))}</td>
  <td class="mono" style="font-size:11px">{h(str(r['Carpeta']))}</td>
  <td class="mono td-right">{float(r['Tamaño_MB']):.2f} MB</td>
  <td class="mono">{h(str(r['Fecha_Modificacion']))}</td>
  <td style="font-size:11px">{h(str(r['Modificado_Por']))}</td>
</tr>"""
        sec_07 = f"""<div class="section" id="post-migracion">
  {sec_header("07",f"Archivos modificados post-migración",f"{fmt_num(a['post_mig'])} archivos{nota_pm}","green","post-migracion")}
  <div class="callout callout-green">
    <span class="callout-label">Implicancia</span>
    Confirmar que estos archivos ya se encuentran en el hub destino antes de archivar el sitio.
  </div>
  <div class="table-wrap"><table>
    <thead><tr><th>Archivo</th><th>Carpeta</th><th class="td-right">Tamaño</th><th>Fecha Modif.</th><th>Modificado Por</th></tr></thead>
    <tbody>{rows_pm}</tbody>
  </table></div>
</div>"""
    else:
        sec_07 = f"""<div class="section" id="post-migracion">
  {sec_header("07","Archivos modificados post-migración","0 archivos","red","post-migracion")}
  <div class="callout callout-red"><span class="callout-label">Sin actividad post-migración</span>Ningún archivo fue modificado después del {fecha_str}. Puede archivarse directamente.</div>
</div>"""

    # ── 08 Archivos grandes ──────────────────────────────────
    if len(a["grandes"]) > 0:
        rows_gr = ""
        for _, r in a["grandes"].iterrows():
            spmt = pill("supera límite SPMT","red") if r["Tamaño_Bytes"] > 250*1024*1024 else pill("OK para SPMT","green")
            rows_gr += f"""<tr>
  <td style="font-size:12px">{h(str(r['Nombre_Archivo']))}</td>
  <td class="mono bold td-right">{fmt_gb(r['Tamaño_Bytes']/1e9)}</td>
  <td class="mono" style="font-size:11px">{h(str(r['Carpeta']))}</td>
  <td class="mono">{h(str(r['Fecha_Modificacion']))}</td>
  <td>{spmt}</td>
</tr>"""
        sec_08 = f"""<div class="section" id="archivos-grandes">
  {sec_header("08",f"Archivos grandes (>{config['mb_archivo_grande']} MB)",f"{len(a['grandes'])} archivos","gold","archivos-grandes")}
  <div class="callout callout-gold"><span class="callout-label">Límite SPMT — 250 MB por archivo</span>Los archivos que superan 250 MB requieren migración manual o por script. SPMT puede ignorarlos silenciosamente.</div>
  <div class="table-wrap"><table>
    <thead><tr><th>Archivo</th><th class="td-right">Tamaño</th><th>Carpeta</th><th>Última Modif.</th><th>SPMT</th></tr></thead>
    <tbody>{rows_gr}</tbody>
  </table></div>
</div>"""
    else:
        sec_08 = f"""<div class="section" id="archivos-grandes">
  {sec_header("08",f"Archivos grandes (>{config['mb_archivo_grande']} MB)","Sin archivos grandes","green","archivos-grandes")}
  <div class="callout callout-green"><span class="callout-label">Sin restricciones para SPMT</span>Todos los archivos están por debajo del umbral configurado.</div>
</div>"""

    # ── 09 Destino (placeholder) ─────────────────────────────
    sec_09 = f"""<div class="section" id="destino">
  {sec_header("09","Destino de migración","Por definir","gold","destino")}
  <div class="placeholder">
    <span class="placeholder-label">Pendiente de completar</span>
    Definir el sitio destino en la arquitectura hub-and-spoke, la estructura de carpetas destino
    y la tabla de mapeo origen → destino por carpeta de primer nivel. Completar junto con el área responsable.
  </div>
  <div style="display:grid;grid-template-columns:2fr .2fr 2fr;gap:10px;align-items:center;margin:16px 0">
    <div style="background:var(--bg-2);border:1px solid var(--border);padding:12px 14px;border-radius:2px">
      <div style="font-family:var(--font-mono);font-size:10px;color:var(--ink-faint);margin-bottom:4px;text-transform:uppercase;letter-spacing:.05em">Origen</div>
      <div style="font-size:13px;font-weight:600">{h(sitio_nombre)}</div>
      <div style="font-family:var(--font-mono);font-size:10px;color:var(--ink-faint);margin-top:4px">{fmt_num(a['total_archivos'])} archivos · {fmt_gb(a['total_gb'])}</div>
    </div>
    <div style="text-align:center;color:var(--ink-faint);font-size:18px">→</div>
    <div style="background:var(--bg-2);border:1px solid var(--border);padding:12px 14px;border-radius:2px">
      <div style="font-family:var(--font-mono);font-size:10px;color:var(--ink-faint);margin-bottom:4px;text-transform:uppercase;letter-spacing:.05em">Destino</div>
      <div style="font-size:13px;font-weight:600;color:var(--ink-faint);font-style:italic">— por definir —</div>
    </div>
  </div>
</div>"""

    # ── 10 Decisiones automáticas ────────────────────────────
    decisiones = []
    d = 1
    if len(a["zips_grandes"]) > 0:
        decisiones.append((d, "red", f"Migración manual de {len(a['zips_grandes'])} ZIP(s) que superan 250 MB",
            "SPMT no puede migrar estos archivos. Requieren carga manual o script PowerShell.", pill("Bloqueante SPMT","red")))
        d += 1
    if "heic" in a["especiales"]:
        n_heic = int(a["especiales"]["heic"][1])
        decisiones.append((d, "gold", f"Conversión de {fmt_num(n_heic)} archivos HEIC a JPG",
            "Sin preview nativo en SPO. Decidir si se convierten antes o después de migrar.", pill("Experiencia de usuario","gold")))
        d += 1
    if "xlsm" in a["especiales"] or "xlsb" in a["especiales"]:
        ext_m = "xlsm" if "xlsm" in a["especiales"] else "xlsb"
        n_m   = int(a["especiales"][ext_m][1])
        decisiones.append((d, "gold", f"{fmt_num(n_m)} archivos {ext_m.upper()} — macros no ejecutan en SPO",
            "Si son herramientas activas, evaluar conversión a Power Apps.", pill("Post-migración","blue")))
        d += 1
    if any(ext in a["especiales"] for ext in ("mp4","mov","avi")):
        gb_video = sum(a["especiales"][e][2] for e in ("mp4","mov","avi") if e in a["especiales"])
        decisiones.append((d, "orange", f"{fmt_gb(gb_video)} de videos — evaluar Azure Blob Storage",
            "Los videos de gran volumen son candidatos a Azure Blob tier Cool.", pill("Destino alternativo","gold")))
        d += 1
    if a["dup_gb"] > 1:
        decisiones.append((d, "gold", f"{fmt_gb(a['dup_gb'])} estimado en duplicados",
            "Migrar todo y deduplicar post-migración con revisión del área.", pill("Optimización","blue")))
        d += 1

    items_dec = ""
    for num, col, titulo, desc, tag in decisiones:
        items_dec += f"""<div style="display:flex;gap:16px;align-items:flex-start;background:var(--white);border:1px solid var(--border);padding:16px 18px;margin-bottom:8px">
  <div style="font-family:var(--font-mono);font-size:20px;font-weight:500;min-width:32px;color:var(--{col})">D-{num:02d}</div>
  <div>
    <div style="font-weight:600;font-size:14px;margin-bottom:4px">{titulo}</div>
    <div style="font-size:13px;color:var(--ink-light);margin-bottom:6px">{desc}</div>
    {tag}
  </div>
</div>"""

    if not items_dec:
        items_dec = f"""<div style="display:flex;gap:16px;align-items:flex-start;background:var(--white);border:1px solid var(--border);padding:16px 18px">
  <div style="font-family:var(--font-mono);font-size:20px;font-weight:500;min-width:32px;color:var(--green)">✓</div>
  <div>
    <div style="font-weight:600;font-size:14px;margin-bottom:4px">Sin decisiones técnicas bloqueantes detectadas</div>
    <div style="font-size:13px;color:var(--ink-light)">{pill("Listo para migrar","green")}</div>
  </div>
</div>"""

    sec_10 = f"""<div class="section" id="decisiones">
  {sec_header("10","Decisiones pendientes",f"{len(decisiones)} ítems detectados","red" if decisiones else "green","decisiones")}
  {items_dec}
  <div class="placeholder" style="margin-top:12px">
    <span class="placeholder-label">Decisiones adicionales — completar manualmente</span>
    Agregar aquí cualquier decisión de negocio o arquitectural específica de este sitio que no sea detectable automáticamente.
  </div>
</div>"""

    body = (
        topbar +
        '<div class="layout">' + sidebar +
        '<main class="main">' +
        hero + '<hr class="divider">' +
        sec_01 + '<hr class="divider">' + sec_02 + '<hr class="divider">' +
        sec_03 + '<hr class="divider">' + sec_04 + '<hr class="divider">' +
        sec_05 + '<hr class="divider">' + sec_06 + '<hr class="divider">' +
        sec_07 + '<hr class="divider">' + sec_08 + '<hr class="divider">' +
        sec_09 + '<hr class="divider">' + sec_10 +
        build_footer(config["nombre"], f"Relevamiento {sitio_nombre}") +
        '</main></div>'
    )

    return html_doc(
        title=f"{config['nombre']} · {sitio_nombre} — Análisis",
        config=config,
        body=body
    )


# ===========================================================================
# MAIN
# ===========================================================================

def main():
    args = parse_args("Genera un HTML de análisis profundo por cada sitio del corpus.")

    print("=" * 60)
    print("  Kit M365 Curaduría — Informes por sitio")
    print("=" * 60)

    config = load_config(args.config)
    print(f"  Cliente : {config['nombre']}")
    print(f"  Carpeta : {config['carpeta_local']}")

    df_raw = load_csvs(config)
    df     = clean_df(df_raw, config)
    print(f"\nFilas cargadas: {len(df):,}")

    carpeta_out = get_carpeta_informes(config)
    ts          = datetime.now()
    ts_str      = ts.strftime("%Y%m%d_%H%M%S")
    ts_label    = ts.strftime("%d/%m/%Y %H:%M")

    sitios    = df.groupby(["Sitio", "URL_Sitio"])
    generados = 0
    errores   = 0

    for (sitio_nombre, sitio_url), df_sitio in sitios:
        if args.sitio and args.sitio.lower() not in sitio_nombre.lower():
            continue

        gb = df_sitio["Tamaño_Bytes"].sum() / 1e9
        print(f"\n  {sitio_nombre} ({len(df_sitio):,} archivos · {gb:.1f} GB)...")

        try:
            html = build_html_sitio(df_sitio, sitio_nombre, sitio_url, config, ts_label)

            nombre_limpio = re.sub(r"[^a-zA-Z0-9_-]", "_", sitio_nombre)
            nombre_limpio = re.sub(r"_+", "_", nombre_limpio)[:60]
            prefijo       = config["nombre"].replace(" ", "_").replace(".", "_")
            out_path      = os.path.join(carpeta_out, f"{prefijo}_Informe_{nombre_limpio}_{ts_str}.html")

            with open(out_path, "w", encoding="utf-8") as f:
                f.write(html)

            kb = os.path.getsize(out_path) / 1024
            print(f"    OK → {os.path.basename(out_path)} ({kb:.0f} KB)")
            generados += 1
        except Exception as e:
            print(f"    ERROR: {e}")
            errores += 1

    print()
    print("=" * 60)
    print(f"  Informes generados : {generados}")
    print(f"  Errores            : {errores}")
    print(f"  Carpeta de salida  : {carpeta_out}")
    print("=" * 60)


if __name__ == "__main__":
    main()
