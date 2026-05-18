# Investigación técnica: acotación automática sobre `DrawingView` (SDK HTML local)

Ámbito: documentación en `docs\SDK_HTML\`. Todo lo que no aparezca ahí queda marcado como **NO CONFIRMADO EN SDK_HTML**.

---

## Convenciones de tablas del informe

Cada fila resume: página consultada, objeto COM, miembro, firma VB documentada, coordenadas (si el HTML lo aclara), contenedor conceptual, uso hipotético para DV, riesgos, prueba mínima.

---

## 1. `Sheet` ↔ colección `Dimensions`

| HTML | Objeto | Miembro | Firma documentada | Coordenadas / unidades | Contenedor | Uso para DV | Riesgos | Prueba mínima |
|------|--------|---------|-------------------|-------------------------|------------|-------------|---------|----------------|
| `SolidEdgeDraft~Sheet~Dimensions.html` | `Sheet` | `Dimensions` (propiedad, solo lectura) | `Public Property Dimensions As Object` | NO CONFIRMADO EN SDK_HTML (el HTML no especifica sistema de coords. de los métodos posteriores) | `Sheet` | Punto de creación habitual de cotas en la hoja vía la colección `Dimensions` devuelta. | Propiedad tipada como `Object` en el HTML; enlazado en tiempo de ejecución a tipo de `SolidEdgeFrameworkSupport.Dimensions`. | Abrir DFT; `dims = ActiveSheet.Dimensions`; comprobar `Count` estable. |

| HTML | Objeto | Miembro | Firma documentada | Coordenadas | Contenedor | Uso para DV | Riesgos | Prueba mínima |
|------|--------|---------|-------------------|-------------|------------|-------------|---------|----------------|
| `SolidEdgeDraft~Sheet~ArrangeDimensionsInSelectSet.html` | `Sheet` | `ArrangeDimensionsInSelectSet` | `Public Sub ArrangeDimensionsInSelectSet(ByVal iStackPitchMultiplier As Double, ByVal bAssociative As Boolean)` | NO CONFIRMADO | `Sheet` | Reordenar cotas ya en el conjunto de selección; el HTML relaciona explícitamente con `DrawingView.AddConnectedDimensionsToSelectSet`. | No sustituye a la creación de cotas; requiere preselección. | Prefijar selección con `AddConnectedDimensionsToSelectSet`; llamar `ArrangeDimensionsInSelectSet` y observar resultado. |

---

## 2. `DrawingView`: geometría DV y utilidades de coordenadas

| HTML | Objeto | Miembro | Firma documentada | Coordenadas | Contenedor | Uso para DV | Riesgos | Prueba mínima |
|------|--------|---------|-------------------|-------------|------------|-------------|---------|----------------|
| `SolidEdgeDraft~DrawingView_members.html` (lista) | `DrawingView` | `DVLines2d` | (propiedad, HTML: devuelve colección `DVLine2d`) | NO CONFIRMADO EN SDK_HTML para valores devueltos por `DVLine2d.GetStartPoint/GetEndPoint` | `DrawingView` | Enumerar aristas proyectadas como líneas. | Confiar solo en combinación documentada vista↔hoja (`ViewToSheet`) para ubicación final. | `Count` después de vista actualizada. |
| `SolidEdgeDraft~DrawingView_members.html` | `DrawingView` | `DVArcs2d` | (colección) | igual que líneas para puntos de arco | `DrawingView` | Arcos proyectados (`GetStartPoint/GetEndPoint` en `SolidEdgeDraft~DVArc2d_*.html`). | Ídem. | Contar ítems. |
| idem | `DrawingView` | `DVCircles2d` | (colección) | Centro vía `DVCircle2d.GetCenterPoint` (`SolidEdgeDraft~DVCircle2d~GetCenterPoint.html`: `Sub GetCenterPoint(ByRef x,y)`) — **sin afirmar espacio en coordenadas** | `DrawingView` | radios/centros en vista | igual | leer centro |
| idem | `DrawingView` | `DVPoints2d` | (colección) | NO CONFIRMADO | `DrawingView` | vértices aislados | puede estar vacío | `Count` |
| `SolidEdgeDraft~DrawingView~ViewToSheet.html` | `DrawingView` | `ViewToSheet` | `Public Sub ViewToSheet(ByVal xView As Double, ByVal yView As Double, ByRef xSheet As Double, ByRef ySheet As Double)` | **Entrada vista → salida hoja** explícito en descripción HTML | `DrawingView` | Convertir puntos de geometría (si están en coords. de vista) a hoja antes de usar como puntos de proximidad si la API lo exige así. **NO CONFIRMADO EN SDK_HTML** que `AddDistanceBetweenObjects` exija coords. de hoja en plano. | Asumir sin prueba ⇒ fallos COM o cotas falsas. | Log de par (vista,hoja) y provocar llamada única a cotación. |
| `SolidEdgeDraft~DrawingView~SheetToView.html` | `DrawingView` | `SheetToView` | `Public Sub SheetToView(ByVal xSheet, ByVal ySheet, ByRef xView, ByRef yView)` | **Hoja → vista** según HTML | `DrawingView` | Inversa de la anterior cuando un dato llega en hoja. | igual | igual |
| `SolidEdgeDraft~DrawingView~ModelToView.html` | `DrawingView` | `ModelToView` | `Public Sub ModelToView(ByVal xModel As Double, ByVal yModel As Double, ByVal zModel As Double, ByRef xView As Double, ByRef yView As Double)` | Modelo→vista (`zModel` aparece como “X coordinate of model.” en texto del parámetro: **ambigüedad en la doc**, no usar sin verificar empírico) | `DrawingView` | Puente modelo→vista | error en documentación parámetro `zModel` | prueba puntual conocida |

### `DVLine2d`: extremos de línea

| HTML | Objeto | Miembro | Firma documentada | Coordenadas | Contenedor | Uso | Riesgos | Prueba |
|------|--------|---------|-------------------|-------------|------------|-----|---------|--------|
| `SolidEdgeDraft~DVLine2d~GetStartPoint.html` | `DVLine2d` | `GetStartPoint` | `Public Sub GetStartPoint(ByRef x As Double, ByRef y As Double)` | **NO CONFIRMADO EN SDK_HTML** si son coords. vista u otras | `DrawingView` (objeto DV) | Puntos para bbox y candidatos | No mezclar con coords. de hoja sin conversión aplicada donde proceda | log `x,y` y comparar con `ViewToSheet` |
| `SolidEdgeDraft~DVLine2d~GetEndPoint.html` | `DVLine2d` | `GetEndPoint` | `Public Sub GetEndPoint(ByRef x As Double, ByRef y As Double)` | idem | idem | idem | idem | idem |
| `SolidEdgeDraft~DVLine2d.html` (+ members) | `DVLine2d` | `Length`, `Angle` | propiedades (ver HTML) | NO CONFIRMADO espacio angular | idem | Clasificar H/V/u oblicuo | tol. numérica manual | clasificar líneas |

### `RetrieveDimensions`, `GraphicMember`

| HTML | Objeto | Miembro | Firma documentada | Coordenadas | Contenedor | Uso | Riesgos | Prueba |
|------|--------|---------|-------------------|-------------|------------|-----|---------|--------|
| `SolidEdgeDraft~DrawingView~RetrieveDimensions.html` | `DrawingView` | `RetrieveDimensions` | `Public Sub RetrieveDimensions(Optional ByVal IsRetrieve As Boolean = True, Optional ByVal DimensionStyleName As String, Optional ByVal TypeLinear As Boolean = True, … hasta `GetCenterMarkToArc`)` | N/A | `DrawingView` | Traer/eliminar cotas **desde modelo** según flags; línea diferente que cotar geometría DV manualmente | No crea cotas nueva API “entre dos DVLine”; solo recuperación | ejecutar en vista con PMI y revisar lista `Sheet.Dimensions` |
| `SolidEdgeDraft~DrawingView~GetReferenceToGraphicMember.html` | `DrawingView` | `GetReferenceToGraphicMember` | (ver ese HTML para firma completa — presente en `DrawingView_members`) | NO CONFIRMADO | `DrawingView` | Posible puente modelo↔grafismo de vista **[DIMENSIÓN: investigar en segunda iteración leyendo ese HTML]** | Alcance por tipo de vista | llamada acotada y log |

---

## 3. `SolidEdgeFrameworkSupport.Dimensions`: creación

| HTML | Objeto | Miembro | Firma documentada | Coordenadas (x,y,z) | Contenedor | Uso para DV | Riesgos | Prueba |
|------|--------|---------|-------------------|---------------------|------------|-------------|---------|--------|
| `SolidEdgeFrameworkSupport~Dimensions~AddDistanceBetweenObjects.html` | `Dimensions` | `AddDistanceBetweenObjects` | `Public Function AddDistanceBetweenObjects(ByVal Object1 As Object, ByVal x1 As Double, ByVal y1 As Double, ByVal z1 As Double, ByVal keyPoint1 As Boolean, ByVal Object2 As Object, ByVal x2 As Double, ByVal y2 As Double, ByVal z2 As Double, ByVal keyPoint2 As Boolean) As Dimension` | HTML describe `x*,y*,z*` como “locate point” para **punto de proximidad** al calcular keypoint; **NO CONFIRMADO EN SDK_HTML** sistema de refs. (¿hoja? ¿vista?) para planos `.dft` | `Dimensions` (`Parent` ⇒ típicamente `Sheet` cuando se usa `sheet.Dimensions`) | Cotar distancia entre **dos objetos COM** — candidatos naturales `DVLine2d`, `DVArc2d`, etc., si COM los acepta | Objetos `DV*` no listados literalmente como `Object*` en ese HTML (`Object` genérico) | crear con dos `DVLine2d`; log + ver en UI |
| `SolidEdgeFrameworkSupport~Dimensions~AddDistanceBetweenObjectsEX.html` | `Dimensions` | `AddDistanceBetweenObjectsEX` | añade `bTangent1 As Boolean`, `bTangent2 As Boolean` al patrón anterior | igual | igual | Opción tangentes | más parámetros | solo si primera falla tangencia |
| `SolidEdgeFrameworkSupport~Dimensions~AddLength.html` | `Dimensions` | `AddLength` | `Public Function AddLength(ByVal Object As Object) As Dimension`; *Remarks*: “Valid objects are **Line**, **Arc** or **Curve**.” | N/A entrada | igual | El HTML **no** nombra `DVLine2d`. **NO CONFIRMADO EN SDK_HTML** compatibilidad con `DVLine2d`. | Usar solo como fallback experimental después de intentar distancia entre objetos | llamada tardía opcional DV vs error COM |
| `SolidEdgeFrameworkSupport~Dimensions_members.html` | `Dimensions` | `DimInitData` | `Public Property DimInitData As DimInitData` (solo lectura) | N/A propiedad colección | `Dimensions` | Inicialización alternativa ligada a `AddDimension` (ver ese método) | flujo paralelo mayor | segunda fase tras distancia |

---

## 4. `Dimension`: trackers, estado, reapilar

| HTML | Objeto | Miembro | Firma / tipo documentado | Unidades TrackDistance | Contenedor | Uso | Riesgos | Prueba |
|------|--------|---------|-------------------------|--------------------------|------------|-----|---------|--------|
| `SolidEdgeFrameworkSupport~Dimension~TrackDistance.html` | `Dimension` | `TrackDistance` | `Public Property TrackDistance As Double` ( lectura/escritura) | HTML (*Remarks*): distancia geométrica que separa cota del objeto origen → **NO CONFIRMADO EN SDK_HTML** igualdad con metros de documento; correlación típica con unidades internas doc. | `Dimension` | Alejar texto/líneas ~0.1 según objetivo práctico | Valor equivocado ⇒ solapamiento o cota invisible | antes/después + captura pantalla |
| `SolidEdgeFrameworkSupport~Dimension~TrackAngle.html` | `Dimension` | `TrackAngle` | Propiedad ángulo cota objeto medido (*members*) | NO CONFIRMADO | `Dimension` | Ajustar orientación de cotas lineales proyectadas | interacciona con herramienta interactiva | lectura inicial + escritura menor |
| `SolidEdgeFrameworkSupport~Dimension_members.html` | `Dimension` | `GetKeyPoint` / `SetKeyPoint` | métodos públicos | devuelven coords. (**ver HTML de firma**) | `Dimension` | Ajustar anclajes tras crear | Índices keypoint ⇒ **NO CONFIRMADO** mapeo sin leer página detalle | después de crear cota |

### Estado y reapilar a vista

| HTML | Objeto | Miembro | Firma | Contenedor | Uso | Riesgos | Prueba |
|------|--------|---------|-------|------------|-----|---------|--------|
| `SolidEdgeFrameworkSupport~Dimension~StatusOfDimension.html` | `Dimension` | `StatusOfDimension` | `Public Property StatusOfDimension As DimStatusConstants` (solo lectura) | Descr.: estado Detached, Error, Driving, Driven… | clasificar cotas experimentales | seDimStatusDetached (1), seDimStatusError (2), Driving (3), Driven (4), seOneEndDetached (5), … (`SolidEdgeFrameworkSupport~DimStatusConstants.html`) | mapear a `connected` / flotante / inválido con criterio explícito en código |
| `SolidEdgeFrameworkSupport~Dimension~ReattachToDrawingView.html` | `Dimension` | `ReattachToDrawingView` | `Public Function ReattachToDrawingView(ByVal DrawingView As Object) As DimReattachStatusConstants` | parámetro: vista donde “dimensions are **not attached**” según texto | recuperar vínculos | convocado solo cuando el estado sugiera reapilar | resultado enum + status posterior |

---

## 5. Seleccionar DV conectadas a cotas (flujo paralelo útil)

| HTML | Objeto | Miembro | Firma | Contenedor | Uso | Nota |
|------|--------|---------|-------|------------|-----|------|
| `SolidEdgeDraft~DrawingView~AddConnectedDimensionsToSelectSet.html` | `DrawingView` | `AddConnectedDimensionsToSelectSet` | `Public Sub AddConnectedDimensionsToSelectSet()` sin args | Vista | Meter en selección cotas ya conectadas a DV | Remark: **no borra selección previa** |

---

## 6. `DrawingView`: tipos útiles (`DrawingViewTypeConstants`)

Origen `SolidEdgeDraft~DrawingViewTypeConstants.html`:

- `igPrincipleView` = 1 (vista principal, típicamente ortogonal frontal/planta/perfil según workflow).
- `igIsometricView` = 2 (excluir para experimento ortogonal estable).
- Otros especializados (`igAuxiliaryView`, `igDetailView`, secciones, etc.) según proyecto.

**Importante:** la propiedad **`DrawingViewType`** (`SolidEdgeDraft~DrawingView~DrawingViewType.html`, `Public Property DrawingViewType As DrawingViewTypeConstants`) es la que devuelve el tipo de vista (p. ej. isométrica). La propiedad **`Type`** (`SolidEdgeDraft~DrawingView~Type.html`, `Public Property Type As Long`) describe “the type of the object being referenced” de forma distinta; **no** deben confundirse al filtrar isométricas.

---

## 7. Materiales marcados como **NO CONFIRMADO EN SDK_HTML** frente al plan

| Aspecto del plan | ¿Doc local clara? | Notas |
|-----------------|-------------------|-------|
| Unidades físicas (`0.10 m`) en cotas `.dft` | NO CONFIRMADO | `GetValueEx` menciona “document units or database units” en otro miembro pero no garantiza igualdad para `TrackDistance`. |
| Sistema coords. esperado por `AddDistanceBetweenObjects` en vistas | PARCIALMENTE CONFIRMADO (descripción punto de proximidad) sin espacio nominado para DFT |
| ¿`DVLine2d` aceptado como `Line` en `AddLength`? | NO CONFIRMADO | Texto sólo lista “Line, Arc or Curve”. |
| `Dimension.Update` después de trackers | Sin entrada en lista rápida de `Dimension_members` | NO CONFIRMADO EN SDK_HTML en esta muestra (exploración no exhaustiva por miles de HTML). |
| Lista completa index keypoint línea DV | existe `DVLine2d.KeyPointCount`/`GetKeyPoint` pero mapping semántico requiere leer parametrización índices | trabajo futuro |

---

## 8. Hallazgos adicionales (búsqueda por términos)

`- [DVDIM][SDK_FINDING]` method=`DrawingView.RetrieveDimensions` html=`SolidEdgeDraft~DrawingView~RetrieveDimensions.html` possibleUse=Poblar vistas con cotas de modelo PMI en una sola llamada parametrizada, como alternativa a dimensionar aristas DV desde cero. risk=Depende existencia PMI y flags; puede generar alto número de elementos y no equivaler a dimensiones tipo “overall” personalizadas.

`- [DVDIM][SDK_FINDING]` method=`Sheet.ArrangeDimensionsInSelectSet` + `DrawingView.AddConnectedDimensionsToSelectSet` html=véase §1 y §5 possibleUse=reorganizar trabajo manual post extracción. risk=Selección debe gestionarse bien.

`- [DVDIM][SDK_FINDING]` method=`DrawingView.GetReferenceToGraphicMember` html=`SolidEdgeDraft~DrawingView~GetReferenceToGraphicMember.html` possibleUse=enlazar entidad modelo con miembro grafico proyectado antes de crear constraints adicionales. risk=tipo de objeto devuelto y casos válidos ⇒ leer ese HTML antes de llamar desde producción.

---

## 9. Pruebas mínimas ordenadas recomendadas (laboratorio)

1. Inventario `DV*` por vista + `ViewToSheet` en esquinas y puntos medio de hasta 3 aristas línea.
2. `Dimensions.AddDistanceBetweenObjects` entre dos `DVLine2d` usando proximidades en **coords. hoja tras `ViewToSheet`** con `keyPoint=True/False` A/B experiment.
3. Ajustar `TrackDistance`.
4. Leer `StatusOfDimension`; si `Detached` repetir opcionalmente `ReattachToDrawingView` y releer estado.
5. Registrar delta `dimsAfter - dimsBefore` en `Sheet.Dimensions.Count`.

Esta orden respeta prohibición explícita de declarar delta=0 sin logs de `[DVDIM][CREATE][TRY]` reales documentados previos.

---

## 10. Invocación del laboratorio `AutoDimDrawingViewLab` (código del repo)

- Clase: `Services\AutoDimDrawingViewLab.vb` (compilada en el ejecutable).
- Entrada principal: `AutoDimDrawingViewLab.RunDrawingViewDimensionLab(activeDraft, logger)` donde `activeDraft` es el COM `DraftDocument` abierto (u objeto con `ActiveSheet`).
- Alternativa: `AutoDimDrawingViewLab.RunOnSheet(activeSheet, logger)` si ya se tiene la `Sheet`.
- Debe ejecutarse en **hilo STA** con Solid Edge disponible (misma regla que el resto de COM del proyecto).
- **No** está enlazado al motor de generación por defecto; llamar solo desde depuración, botón experimental o consola interna cuando se quiera auditar un DFT abierto.
