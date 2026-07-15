# M365 Curaduría Kit

Toolkit de Circo Studio para diagnóstico, análisis y mejora del estado informacional de organizaciones en Microsoft 365.

---

## Qué es esto

Conjunto de scripts PowerShell y Python, plantillas de entregables HTML y documentación metodológica para ejecutar engagements de **curaduría informacional** sobre tenants de M365 (SharePoint Online, Teams, OneDrive).

La premisa central: los datos que una organización genera en su día a día tienen valor estratégico que la mayoría no percibe. Este kit lo hace visible, lo cuantifica y lo convierte en decisiones accionables — y en la condición técnica para que la IA sobre conocimiento interno funcione.

---

## Para quién es

Consultores de Circo Studio que ejecutan engagements de gobernanza informacional. El kit es cliente-agnóstico: cada engagement instancia su propia configuración sobre la estructura común.

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
├── powershell/
│   ├── M365Curaduria.psm1          # Módulo compartido (6 funciones base)
│   ├── ScriptA-Inventario.ps1      # Inventario de sitios del tenant
│   ├── ScriptB-Relevamiento.ps1    # Relevamiento de documentos por sitio
│   ├── ScriptC-AsignarOwner.ps1    # Asignación de owner temporal
│   ├── ScriptD-Masivo.ps1          # Wrapper de ejecución masiva
│   └── New-KitCuraduria.ps1        # Scaffolding de nuevo engagement
│
├── python/
│   ├── m365_curaduria.py           # Módulo compartido
│   ├── analisis_contenido.py       # Detección de duplicados, obsoletos, categorías
│   ├── reporte_sitios.py           # Métricas por sitio
│   └── generar_indice.py           # Índice consolidado del tenant
│
├── config/
│   ├── cliente_config.json         # Plantilla de configuración por engagement
│   └── kit_config_editor.html      # Editor visual de configuración
│
├── entregables/
│   ├── analisis-sitio.html         # Plantilla: análisis de sitio individual
│   ├── comparativo-carpetas.html   # Plantilla: origen vs. destino
│   ├── dashboard-isi.html          # Plantilla: dashboard ISI
│   └── reporte-ejecutivo.html      # Plantilla: presentación a dirección
│
├── posicionamiento/
│   ├── index.html                  # Índice de la serie
│   ├── por-que.html                # La gobernanza como problema no-tecnológico
│   ├── metodologia.html            # Modelo de engagement en cuatro fases
│   ├── kit.html                    # Introducción al kit
│   └── script-b-masivo.html        # Documentación técnica IT: Script B Masivo
│
└── runbook/
    └── configuracion-visual-m365.html  # Runbook: identidad visual M365
```

---

## Cómo usar en un nuevo engagement

1. Ejecutar `New-KitCuraduria.ps1` — genera la estructura de carpetas locales para el cliente
2. Editar `cliente_config.json` con los datos del tenant (usar `kit_config_editor.html`)
3. Ejecutar **Script A** para obtener el inventario de sitios
4. Priorizar sitios y ejecutar **Script D** (masivo) o **Script B** sitio por sitio
5. Correr los scripts Python sobre los CSVs generados
6. Producir los entregables HTML con los datos del cliente

---

## Requisitos técnicos

- **PowerShell 7+** con módulo [PnP.PowerShell](https://pnp.github.io/powershell/) instalado
- **Python 3.10+** con dependencias en `python/requirements.txt`
- App Registration en el tenant del cliente con permisos `Sites.FullControl.All` (o acceso delegado con cuenta de Circo Studio como SCA)

---

## Licencia

MIT — ver [LICENSE](./LICENSE)

© 2025 Circo Studio
