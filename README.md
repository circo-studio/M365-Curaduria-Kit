# M365 Curaduría Kit
### Preparación informacional para IA en Microsoft 365

> La IA amplifica lo que encuentra. No inventa conocimiento — refleja el estado del corpus sobre el que opera.
> Este kit es la herramienta para medir y gobernar ese estado antes de encender motores como Copilot.

---

## El problema

Microsoft 365 acumula deuda informacional de forma silenciosa: sitios sin owner, permisos rotos, documentos con versiones superadas mezcladas con vigentes, contenido huérfano de empleados que ya no están. El gobierno tradicional de datos falla porque se aborda como un problema de IT — y el área de IT no tiene el contexto de negocio para resolverlo sola.

El resultado: cuando se activa Copilot o cualquier sistema RAG sobre ese corpus, la IA no falla silenciosamente. Responde con confianza usando información obsoleta, contradictoria o simplemente equivocada.

**La secuencia correcta es siempre: curar primero, indexar después.**

---

## La solución

Este kit provee los scripts, las métricas y la metodología para diagnosticar el estado informacional de un tenant M365 y construir un plan de mejora accionable.

El resultado central es el **ISI — Índice de Salud Informacional**: un número entre 0 y 100 que sintetiza el estado del corpus en seis dimensiones.

| Dimensión | Peso | Qué mide |
|---|---|---|
| Ownership | 25% | % de contenido bajo custodia activa |
| Arquitectura | 20% | % del corpus en estructura gobernada |
| Sharing externo | 20% | % de sitios con política correcta |
| Calidad del corpus | 15% | Reducción de duplicados, obsoletos, formatos problemáticos |
| Ciclo de vida | 10% | % con política de retención definida |
| Actividad | 10% | % de contenido modificado en últimos 90 días |

Un ISI de 25–35 es esperable en organizaciones sin gobernanza previa. Cuando el ISI supera 70, el corpus tiene las condiciones para alimentar sistemas de IA con resultados confiables.

---

## Qué incluye el kit

### PowerShell — extracción de metadatos

| Script | Qué hace |
|---|---|
| `M365Curaduria.psm1` | Módulo compartido: autenticación, logging, clasificación de archivos |
| `Script_A_Inventario.ps1` | Inventario completo del tenant: sitios, storage, owners, sharing externo |
| `Script_A_AsignarOwner.ps1` | Asignación de acceso temporal en sitios sin owner para poder analizarlos |
| `Script_B_Masivo.ps1` | Relevamiento masivo de documentos: metadatos, rutas, fechas, autores |
| `Script_C_Refresh.ps1` | Actualización incremental — re-escanea sitios modificados desde el último relevamiento |

### Python — análisis y entregables

| Script | Qué hace |
|---|---|
| `m365_curaduria.py` | Módulo compartido: lectura de config, clasificación, utilidades |
| `analizar_contenido.py` | Detección de duplicados, obsoletos, formatos problemáticos, distribución por categoría |
| `generar_informes_sitio.py` | HTMLs de análisis por sitio con métricas, hallazgos y matriz de destino |
| `generar_indice.py` | Índice consolidado del tenant con el ISI y el estado de cada sitio |

### Config y herramientas

| Archivo | Qué hace |
|---|---|
| `Config/cliente_config.json` | Configuración del engagement: tenant URL, ClientId, rutas, sitios excluidos |
| `Config/kit_config_editor.html` | Editor visual del config — genera el JSON y los snippets de ejecución listos para pegar |
| `New-KitCuraduria.ps1` | Scaffolding: crea la estructura de carpetas locales para un nuevo engagement |

### Documentación

La carpeta `Posicionamiento/` contiene una serie de documentos HTML que explican la metodología, el modelo de engagement y las instrucciones técnicas para el área de IT del cliente (incluyendo guía de App Registration y documentación del Script B).

---

## Seguridad y privacidad

> **Para equipos de InfoSec y administradores de tenant**

Los scripts operan en **modo solo lectura**. No descargan, modifican ni eliminan ningún archivo del tenant.

Lo que hacen:
- Leer metadatos de sitios y documentos vía PnP PowerShell y Microsoft Graph API
- Escribir CSVs localmente en la máquina que ejecuta el script
- Generar HTMLs de análisis a partir de esos CSVs, también en local

Lo que **no** hacen:
- Acceder al contenido de los archivos (solo metadatos: nombre, tamaño, fechas, autor, ruta)
- Enviar datos fuera del entorno local
- Modificar permisos, estructuras ni configuraciones del tenant

Toda la información generada queda en el entorno local de quien ejecuta el kit. No hay dependencias de servicios externos ni telemetría.

---

## Quick Start

### Requisitos previos

- PowerShell 7+ con [PnP.PowerShell](https://pnp.github.io/powershell/) instalado
- Python 3.10+
- Una App Registration en el tenant con permisos `Sites.FullControl.All` (o una cuenta con rol de Site Collection Administrator). Ver `Posicionamiento/CS_AppRegistration_Guia.html` para instrucciones paso a paso.

### Primera ejecución

**1. Configurar el engagement**

Abrir `Config/kit_config_editor.html` en el browser, completar los datos del tenant y descargar el `cliente_config.json` generado a la carpeta `Config/`.

**2. Obtener el inventario del tenant**

```powershell
.\PowerShell\Script_A_Inventario.ps1 -ConfigPath ".\Config\cliente_config.json"
```

Genera un CSV con todos los sitios del tenant: storage, owners, sharing externo, tipo de sitio. Este es el insumo para decidir qué sitios relevar.

**3. Relevar documentos**

```powershell
.\PowerShell\Script_B_Masivo.ps1 -ConfigPath ".\Config\cliente_config.json"
```

Itera sobre los sitios priorizados y extrae los metadatos de todos los documentos. Puede tardar horas en tenants grandes — es idempotente, se puede interrumpir y retomar.

**4. Analizar y generar entregables**

```bash
python Python/analizar_contenido.py --config Config/cliente_config.json
python Python/generar_informes_sitio.py --config Config/cliente_config.json
python Python/generar_indice.py --config Config/cliente_config.json
```

Produce los HTMLs de análisis por sitio y el índice consolidado con el ISI de línea de base.

---

## Adaptar el kit

Los scripts son un punto de inicio. Cada tenant tiene su propia realidad — naming conventions distintas, estructuras heredadas, modelos de acceso particulares. El kit está construido para ser modificado: los parámetros de clasificación, los umbrales del ISI, los criterios de filtrado y los formatos de salida son configurables o directamente editables en el código.

---

## Laboratorio

Si este kit te resultó útil, en [Circo Studio Laboratorio](https://circostudio.io/laboratorio.html) publicamos ensayos sobre tecnología, gobernanza informacional e inteligencia artificial en organizaciones. El mismo marco conceptual que da origen a este kit, desarrollado en forma de argumento.

---
## Contribución y licencia

El kit se publica bajo licencia **MIT**. Se puede usar, modificar y distribuir libremente, incluyendo para uso corporativo. Ver [LICENSE](./LICENSE).

Los aportes son bienvenidos — issues, PRs y forks.

© 2025 Circo Studio · [circostudio.io](https://circostudio.io)
