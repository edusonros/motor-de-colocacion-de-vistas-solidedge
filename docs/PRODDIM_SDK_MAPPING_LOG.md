# Registro: mapeo PRODDIM ↔ documentación SDK Solid Edge (Draft / Dimensions)

**Fecha de elaboración:** 2026-05-11  
**Alcance:** cruce entre `ProductionDvRefCleanDimensionEngine.vb` (motor [PRODDIM]) y los temas HTML del SDK en `docs\SDK_HTML` analizados en la investigación previa.

**Propósito:** dejar constancia auditada de cómo cada fase del código encaja con la API oficial (colección `Dimensions`, geometría de vista `DVLine2d`, `ViewToSheet`, `Dimension.*`, etc.).

---

## 1. Convenciones

- Las **referencias de líneas** son al archivo `Services\Dimensioning\ProductionDvRefCleanDimensionEngine.vb` salvo indicación contraria.
- Los **temas SDK** citan los nombres de archivo HTML del CHM/HTML export (`SolidEdgeDraft~…`, `SolidEdgeFrameworkSupport~…`, `SolidEdgeConstants~…`).
- Para comparación rápida con el otro motor de cotas del repo se incluye una nota sobre `DimensionPlacementEngine.vb`.

---

## 2. Arranque, hoja y colección Dimensions

| Líneas (aprox.) | Código / comportamiento | Tema SDK / API |
|-----------------|-------------------------|----------------|
| 113–131 | `Run`, resolución de `Sheet`, activación | `DraftDocument`, `Sheet`; ejemplo en `SolidEdgeFrameworkSupport~Dimensions.html` usa `objSheet = objDraftDocument.ActiveSheet` |
| 124–129, 157–165 | `sh.Dimensions` → `Dimensions` | `SolidEdgeDraft~Sheet~Dimensions.html`: *Returns the Dimensions collection object for the referenced object.* |
| 167–170, 826–868 | `DraftDocument.DimensionStyles`, `DimensionStyles.Item` | Colección de estilos de documento Draft (errores COM documentados por el proyecto en TabletPC/`0x80040223`) |
| 870–888 | `d.Style = styleTyped` | `SolidEdgeFrameworkSupport~Dimension~Style` / `SolidEdgeFrameworkSupport~DimStyle.html` |

---

## 3. DrawingView, DVLines2d, DVLine2d (geometría proyectada)

| Líneas (aprox.) | Código / comportamiento | Tema SDK / API |
|-----------------|-------------------------|----------------|
| 412–505 | `BuildViewFrames`: `sh.DrawingViews`, `DrawingView`, `DVLines2d.Count`, `DrawingView.Range` | `SolidEdgeDraft~DrawingView.html`, `SolidEdgeDraft~DrawingView~DVLines2d.html` |
| 444 | `dv.DVLines2d.Count` | Propiedad DVLines2d: colección `DVLine2d` |
| 507–598 | `CollectDvLines`: `dv.DVLines2d.Item(i)`, `GetStartPoint`, `GetEndPoint`, `Range`, `KeyPointCount` | `SolidEdgeDraft~DVLine2d~GetStartPoint.html`, `GetEndPoint`, `SolidEdgeDraft~DVLine2d~Range.html` — coordenadas en espacio vista |
| 560–561 | `dv.ViewToSheet(midX, midY, …)` (punto medio de línea solo para log) | `SolidEdgeDraft~DrawingView~ViewToSheet.html` |
| 89 | `DvLineInfo.Obj As DVLine2d` — referencias COM pasadas luego a `AddDistance*` | Objeto válido como *Object1/Object2* en `AddDistanceBetweenObjects` |

---

## 4. ViewToSheet + AddDistanceBetweenObjects (proximidad y keyPoint)

| Líneas (aprox.) | Código / comportamiento | Tema SDK / API |
|-----------------|-------------------------|----------------|
| 209–216, 263–266 | `ViewToSheet` de puntos de anclaje elegidos (H_MAX: vértices superiores transformados a hoja; V_MAX: X medio + MidY inferior/superior) | Conversión vista→hoja alineada con guía práctica Draft: cotas sobre `Sheet` con picks coherentes |
| 1091–1129 | `TryAddDistanceControlled`: loop `kp1`,`kp2` ∈ {False,True}; `AddDistanceBetweenObjects` con `z=0`; fallback `AddDistanceBetweenObjectsEX` con `bTangent=False` | `SolidEdgeFrameworkSupport~Dimensions~AddDistanceBetweenObjects.html`: *locate point* = proximidad para calcular key point; `keyPoint` True/False; `SolidEdgeFrameworkSupport~Dimensions~AddDistanceBetweenObjectsEX.html` |

**Nota de diseño (PRODDIM vs `DimensionPlacementEngine`):**

- PRODDIM transforma picks a **hoja** antes de llamar (`ViewToSheet` en 210–211, 265–266) y registra vista+hoja (`[ADDDIST_INPUT]`).
- `DimensionPlacementEngine.TryAddDistanceBetweenObjectsGeneric` (aprox. 677–694) usa por defecto **coordenadas de hoja** y opcionalmente un respaldo con `SheetToView` (`UseViewSpaceProximityForAddDistance`). Ambas rutas siguen la misma semántica del SDK: objeto + punto de proximidad + flags keyPoint.

---

## 5. Post-creación: Constraint, Reattach, Update, Value

| Líneas (aprox.) | Código / comportamiento | Tema SDK / API |
|-----------------|-------------------------|----------------|
| 1132–1137 | `d.Constraint = False` | `SolidEdgeFrameworkSupport~Dimensions~Constraint` / propiedad en `Dimension` |
| 1140–1149 | `draft.UpdateAll`, `app.DoIdle` | Patrón COM/OleMessageFilter (no es un único tema; estabiliza estado tras mutación) |
| 1155–1184 | `ReattachToDrawingView(dv)` + log `DimReattachStatusConstants` | `SolidEdgeFrameworkSupport~Dimension~ReattachToDrawingView.html`; `SolidEdgeConstants~DimReattachStatusConstants.html` (0=succeeded, 1=failed) |
| 1202–1208 | `d.Value` | Valor numérico de la cota |
| 1187–1199 | `MeasurementAxisEx` / `MeasurementAxisDirection` vía `CallByName` | API extendida no detallada en el HTML mínimo revisado; uso pragmático para V_MAX |

---

## 6. TrackDistance, Range, GetKeyPoint, colocación y validación

| Líneas (aprox.) | Código / comportamiento | Tema SDK / API |
|-----------------|-------------------------|----------------|
| 1275–1337 | `PlaceCleanDimension`: barrido `d.TrackDistance` | `SolidEdgeFrameworkSupport~Dimension~TrackDistance.html` — *distance that separates the dimension from the object* (lineales) |
| 1246–1273, 1351–1374 | `TryGetDimensionRangeSheetBBox` + `TryCornersViewToSheetBBox` si `Range` parece espacio vista | Compensa ambigüedad documentada en comentarios internos; `ViewToSheet` es el ancla SDK |
| 1376–1431 | `CenterTextByDimensionKeypointSweep`: `KeyPointCount`, `GetKeyPoint` | `Dimension.GetKeyPoint` en API; `SetKeyPoint` documentado en `SolidEdgeFrameworkSupport~Dimension~SetKeyPoint.html` (PRODDIM aquí inspecciona; ajuste de texto vía `CoordinateTextPosition`/`SetTextOffsets` por reflexión) |
| 1446–1504 | `FinalValidateKeepsDimension`: `Value`, `Visible`, `GetRelatedCount`, `GetDisplayData`, `Range` | Validación de aplicación; no sustituye contrato SDK individual |

---

## 7. Eliminación

| Líneas (aprox.) | Código / comportamiento | Tema SDK / API |
|-----------------|-------------------------|----------------|
| 1569–1576 | `d.Delete()` | Patrón habitual de objetos `Dimension` en Automation |

---

## 8. Resumen de correspondencia “flujo SDK → PRODDIM”

1. **Obtener `Sheet.Dimensions`** → líneas 157–165.  
2. **Enumerar `DrawingView` y `DVLines2d`, leer `DVLine2d` con `GetStartPoint`/`GetEndPoint`/`Range`** → 412–598.  
3. **Elegir dos `DVLine2d` (lógica de negocio H_MAX/V_MAX)** → 172–295, 902–1088.  
4. **Convertir picks con `ViewToSheet`** → 209–216, 265–266.  
5. **`AddDistanceBetweenObjects` / EX con proximidad y keyPoint** → 1091–1129.  
6. **`ReattachToDrawingView`** → 223, 273, 1155–1184.  
7. **`TrackDistance` + scoring de colocación** → 1275–1337.  
8. **`Dimension.Style` (DimStyle)** → 870–888.  

---

## 9. Archivo complementario (mismo dominio, otra estrategia de coordenadas)

`Services\Dimensioning\DimensionPlacementEngine.vb` expone en comentarios de cabecera (aprox. 10–16) la misma API `AddDistanceBetweenObjects` con **puntos de proximidad en hoja** y rutas `SheetToView` opcionales (685–699), más bucles keyPoint (769–861). Sirve de **cotejo** con PRODDIM: mismo contrato SDK, distinta política de espacio de coordenadas y saneamiento (`IsInsertedDimensionSpatiallySane`, crop de vista).

---

## 10. Cierre del registro

- **Trabajo realizado:** lectura de `ProductionDvRefCleanDimensionEngine.vb` y `DimensionPlacementEngine.vb`; correlación con temas `docs\SDK_HTML` citados en la investigación de dimensionado en Draft.  
- **No modificado:** ningún archivo de código; solo este registro en `docs\PRODDIM_SDK_MAPPING_LOG.md`.
