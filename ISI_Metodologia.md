# El Índice de Salud Informacional (ISI) — Modelo de Referencia

Este documento describe el modelo matemático y las reglas lógicas que definen el ISI como estándar de medición de la salud informacional de un tenant Microsoft 365.

> **Nota sobre la implementación actual del kit**
>
> Los scripts incluidos en este repositorio (`analizar_contenido.py`, `generar_informes_sitio.py`) implementan una versión funcional del ISI adaptada a lo que es extraíble con PnP PowerShell sin acceso al contenido de los archivos. Las diferencias concretas respecto al modelo de referencia son:
>
> - **Calidad del corpus:** el kit detecta duplicados por nombre + tamaño exacto. El modelo de referencia propone además considerar rutas similares y patrones de distribución, que requieren lógica adicional.
> - **Ciclo de vida:** el kit usa la presencia de políticas de Purview como proxy binario (sí/no). El modelo contempla granularidad por biblioteca y tipo de contenido.
> - **Ownership:** el kit evalúa la presencia de SCA activo. La validación contra Entra ID (cuenta deshabilitada vs. activa) requiere permisos adicionales de Graph API que no siempre están disponibles en todos los engagements.
>
> El modelo de referencia es el objetivo. Los scripts son un punto de inicio. Toda organización debería ajustar los umbrales, los pesos y las reglas de cálculo a su propia realidad.

---

## El modelo matemático general

El ISI de un tenant se calcula como el promedio ponderado de seis dimensiones evaluadas por sitio, agregado por volumen de almacenamiento.

El puntaje de un sitio `j`:

```
ISI_j = Σ (D_i,j × W_i)   para i = 1..6
```

Donde `D_i,j` es el puntaje (0–100) en la dimensión `i` para el sitio `j`, y `W_i` es el peso de esa dimensión.

El ISI global del tenant es el promedio de los `ISI_j` ponderado por el storage de cada sitio — los sitios más grandes impactan proporcionalmente más en la métrica final.

### Pesos por dimensión

| Dimensión | Peso | Qué mide |
|---|---|---|
| Ownership | 25% | Responsable activo y validado |
| Arquitectura | 20% | Estructura gobernada (naming / hubs) |
| Sharing externo | 20% | Políticas de acceso y exposición |
| Calidad del corpus | 15% | Deduplicación y formatos procesables |
| Ciclo de vida | 10% | Retención y archivado explícito |
| Actividad | 10% | Modificación reciente (últimos 90 días) |

El ownership tiene el mayor peso porque sin custodia humana las demás dimensiones no son sostenibles — el corpus más limpio técnicamente pierde valor si nadie lo mantiene.

---

## Reglas de cálculo por dimensión

### 1. Ownership (25%)

Verifica el estado del Site Collection Administrator o de los propietarios del grupo M365 contra Entra ID.

- **100 puntos:** Al menos un Knowledge Owner explícito asignado con cuenta activa en Entra ID.
- **50 puntos:** Dueños presentes pero genéricos (cuenta de servicio, cuenta de IT compartida) o solo miembros sin propietarios explícitos.
- **0 puntos:** Sitio huérfano, o todos los propietarios registrados tienen cuenta deshabilitada en Entra ID.

### 2. Arquitectura (20%)

Evalúa si el sitio vive en una estructura gobernada con criterios de acceso definidos.

- **100 puntos:** El sitio está asociado a un Hub Site de negocio, o cumple estrictamente con las convenciones de naming configuradas.
- **0 puntos:** Sitio Teams auto-generado sin clasificación, spoke suelto, o estructura legacy sin modificación estructural.

### 3. Sharing externo (20%)

Audita la política de `SharingCapability` a nivel de sitio.

- **100 puntos:** `Disabled` o `ExistingExternalUserSharingOnly` — compartición externa bloqueada o estrictamente controlada.
- **0 puntos:** `ExternalUserAndGuestSharing` o `ExternalUserSharingOnly` sin controles adicionales.

Un sitio con sharing externo irrestricto es un riesgo doble: de seguridad hoy, y de oversharing cuando la IA use ese sitio como fuente.

### 4. Calidad del corpus (15%)

Mide la proporción del corpus libre de ruido informacional.

```
D_calidad = 100 − ((V_duplicados + V_legacy + V_basura) / V_total × 100)
```

Donde `V` representa volumen en bytes:
- `V_duplicados`: archivos con mismo nombre, tamaño y extensión en rutas similares.
- `V_legacy`: archivos no procesables eficientemente por motores de búsqueda semántica (`.doc`, `.xls` sin conversión, `.zip` grandes, binarios sin OCR).
- `V_basura`: archivos en ramas de carpetas ignoradas (carpetas nombradas `SUPERADO`, `BORRADORES`, `TMP`, etc.).

### 5. Ciclo de vida (10%)

Determina si el corpus crece sin control o tiene política de auto-regulación.

- **100 puntos:** El sitio tiene políticas de Microsoft Purview aplicadas (retención o archivado automático).
- **0 puntos:** Sin políticas aplicadas a nivel de sitio ni de biblioteca.

### 6. Actividad (10%)

Evalúa la vigencia del conocimiento como proxy de relevancia.

```
D_actividad = (V_modificado_≤90días / V_total) × 100
```

Se extrae el campo `Modified` de cada archivo desde el CSV del relevamiento. Un sitio donde más del 50% del contenido no ha sido modificado en años indica corpus potencialmente obsoleto — la IA lo ingeriría igual, con confianza.

---

## Escala de madurez para IA

| Rango | Estado | Interpretación |
|---|---|---|
| 0 – 35 | Línea de base | Rango típico inicial. Ruido informacional alto. Activar Copilot o RAG en este estado produce respuestas con confianza sobre contenido obsoleto o contradictorio. |
| 36 – 69 | Fase de curación | Progreso medible. El ruido se reduce, el ownership se consolida. El corpus mejora su confiabilidad progresivamente. |
| 70+ | Preparación IA | Umbral técnico superado. Corpus confiable, owners validados. Las pruebas de concepto de IA pueden ejecutarse de forma segura y auditable. |

---

## Cómo adaptar el modelo

Los pesos y umbrales propuestos reflejan una práctica de gobernanza informacional en organizaciones medianas con M365 como plataforma central. Son un punto de inicio, no una verdad universal.

Algunas organizaciones pueden necesitar:
- Mayor peso en Sharing externo si operan en sectores regulados (salud, finanzas, gobierno).
- Menor peso en Actividad si gestionan corpus de archivo con valor histórico pero baja modificación esperada.
- Dimensiones adicionales: cumplimiento normativo, etiquetado de confidencialidad, cobertura de Content Types.

El modelo es abierto. El kit es MIT.
