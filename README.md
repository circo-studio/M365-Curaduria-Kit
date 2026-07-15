# M365 Curaduría Kit

Toolkit para diagnóstico, análisis y mejora del estado informacional de organizaciones en Microsoft 365.

---

## Qué es esto

Conjunto de scripts PowerShell y Python, plantillas de entregables HTML y documentación metodológica para ejecutar engagements de **curaduría informacional** sobre tenants de M365 (SharePoint Online, Teams, OneDrive).

La premisa central: los datos que una organización genera en su día a día tienen valor estratégico que la mayoría no percibe. Este kit lo hace visible, lo cuantifica y lo convierte en decisiones accionables — y en la condición técnica para que la IA sobre conocimiento interno funcione.

---

## Para quién es

Para cualquier persona u organización que necesite diagnosticar y mejorar el estado informacional de su entorno M365. Los scripts son un punto de inicio — están diseñados para ser modificados y adaptados a cada realidad. No hay dos tenants iguales.

---

## La metodología en una página

**ISI — Índice de Salud Informacional** (0–100): indicador sintético calculado sobre seis dimensiones — Ownership (25%), Arquitectura (20%), Sharing externo (20%), Calidad del corpus (15%), Ciclo de vida (10%), Actividad (10%).

**Cinco dominios de análisis:** Estructura y arquitectura · Custodia y ownership · Calidad del corpus · Acceso y seguridad · Preparación para IA.

**Cuatro fases del engagement:**
- **Fase 0 — Exploración:** inventario de sitios, línea de base ISI
- **Fase 1 — Diagnóstico:** análisis completo, mapa de prioridades
- **Fase 2 — Curación activa:** corrección de sharing, ownership, deduplicación, migración
- **Fase 3 — Madurez:** retención Purview, etiquetado, POC de IA sobre corpus curado

---

## Estructura del repo

```
M365-Curaduria-Kit/
│
├── PowerShell/
│   ├── M365Curaduria.psm1              # Módulo compartido (funciones base)
│   ├── Script_A_Inventario.ps1         # Inventario de sitios del tenant
│   ├── Script_A_AsignarOwner.ps1       # Asignación de owner temporal
│   ├── Script_B_Masivo.ps1             # Relevamiento masivo de documentos
│   └── Script_C_Refresh.ps1            # Actualización incremental
│
├── Python/
│   ├── m365_curaduria.py               # Módulo compartido
│   ├── analizar_contenido.py           # Detección de duplicados, obsoletos, categorías
│   ├── generar_informes_sitio.py       # Métricas por sitio
│   └── generar_indice.py               # Índice consolidado del tenant
│
├── Config/
│   ├── cliente_config.json             # Plantilla de configuración por engagement
│   └── kit_config_editor.html          # Editor visual de configuración + snippets
│
├── Posicionamiento/
│   ├── CS_Curaduria_Index.html         # Índice de la serie
│   ├── CS_Curaduria_Por_Que.html       # La gobernanza como problema no-tecnológico
│   ├── CS_Curaduria_Como.html          # Modelo de engagement en cuatro fases
│   ├── CS_Curaduria_Kit.html           # Introducción al kit
│   ├── CS_Script_B_Documentacion_IT.html  # Documentación técnica: Script B Masivo
│   └── CS_AppRegistration_Guia.html    # Guía de App Registration para el tenant
│
├── Runbooks/
│   └── Runbook_IdentidadVisual_M365.ps1   # Runbook: identidad visual M365
│
└── New-KitCuraduria.ps1                # Scaffolding de nuevo engagement
```

---

## Cómo usar en un nuevo engagement

1. Ejecutar `New-KitCuraduria.ps1` — genera la estructura de carpetas locales
2. Editar `Config/cliente_config.json` (usar `kit_config_editor.html` como editor visual)
3. Ejecutar **Script_A_Inventario** para obtener el inventario de sitios del tenant
4. Ejecutar **Script_A_AsignarOwner** para asignar acceso en los sitios a relevar
5. Ejecutar **Script_B_Masivo** sobre la lista de sitios priorizados
6. Correr los tres scripts Python sobre los CSVs generados
7. Revisar los entregables HTML producidos con los datos del tenant

---

## Requisitos técnicos

- **PowerShell 7+** con módulo [PnP.PowerShell](https://pnp.github.io/powershell/) instalado
- **Python 3.10+**
- App Registration en el tenant con permisos `Sites.FullControl.All`, o acceso delegado con una cuenta con rol de Site Collection Administrator

---

## Licencia

MIT — ver [LICENSE](./LICENSE)

© 2025 Circo Studio · [circostudio.io](https://circostudio.io)
