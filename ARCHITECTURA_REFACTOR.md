# Arquitectura del sistema de generación automática de Drafts

## Resumen

El sistema genera automáticamente un `.dft` desde un `.par` o `.psm`, eligiendo la mejor vista base, aplicando rotación 0° o 90° si conviene, calculando la escala óptima y colocando:

- Vista base
- Vista proyectada derecha (AddByFold)
- Vista proyectada inferior (AddByFold)
- Vista isométrica
- Vista flat (si PSM)

---

## 1. Arquitectura modular

### Módulos principales

| Módulo | Función |
|--------|---------|
| `DraftLayoutTypes.vb` | Estructuras de datos: ViewRect, ViewSize, BaseViewCandidate, FoldLayoutOption, ResolvedLayout, ViewSizesAt1 |
| `DraftGenerator.vb`   | Motor principal: orquestación, medición, candidatos, resolución de layout, inserción |
| `LayoutEngine.vb`     | GetUsableAreaForTemplate() para el área útil del template |
| `CojonudoBestFit_Bueno_02.vb` | CreateFlatViewForDraft (flat pattern), delegación de CreateDraftAlzadoPrimerDiedro |

### Estructuras de datos (`DraftLayoutTypes.vb`)

- **ViewRect**: Rectángulo min/max, Width, Height, Center, Area
- **ViewSize**: Width, Height, Area, Create(w,h)
- **Point2D**: X, Y
- **ViewSizesAt1**: Diccionario por orientación → (Width, Height)
- **BaseViewCandidate**: BaseOri, BaseOriName, Rotated90, BaseSize, RightSize, DownSize
- **FoldLayoutOption**: Candidato + escala + BlockWidth/Height + RejectReason + LayoutScore
- **ResolvedLayout**: TemplatePath, Scale, posiciones top-left de base, right, down, iso, flat

---

## 2. Flujo principal

```
CreateDraftAlzadoPrimerDiedro()
    └── DraftGenerator.CreateAutomaticDraftFromModel()
            ├── MeasureAllViewSizes()      → ViewSizesAt1 (Front, Top, Left, Right, Bottom @ escala 1)
            ├── GenerateBaseViewCandidates() → 6 candidatos (Front/Top/Left × 0°/90°)
            ├── ResolveBestLayout()        → ResolvedLayout (mejor candidato + posiciones)
            ├── InsertBaseView()
            ├── InsertFoldedView(igFoldRight)
            ├── InsertFoldedView(igFoldDown)
            ├── InsertIsoView()
            └── CreateFlatViewForDraft()   [si PSM]
```

---

## 3. Motor de layout

1. **Candidatos base**: Front 0°, Front 90°, Top 0°, Top 90°, Left 0°, Left 90°  
   Para cada uno se calculan dimensiones a escala 1 de base, derecha e inferior (considerando rotación).

2. **Bloque ortogonal**:
   - `blockW = baseW + GAP(20mm) + rightW`
   - `blockH = baseH + GAP(20mm) + downH`

3. **Escala**: Se toma la mayor escala estándar que permite que el bloque quepa en el área útil.

4. **Score**: Se puntúa por aprovechamiento del área y bonus de escala. El mejor candidato define el layout final.

5. **Posiciones**: top-left de base, right, down; ISO encima del cajetín; flat centrado en hueco inferior.

---

## 4. Reglas de layout

- Offset desde borde útil superior-izquierdo: **30 mm**
- Gap entre base y derecha: **20 mm**
- Gap entre base e inferior: **20 mm**
- Evitar solapamiento con cajetín
- Base visualmente dominante

---

## 5. Sustitución del flujo anterior

| Antes | Ahora |
|------|-------|
| `CreateDraftAlzadoPrimerDiedro` con lógica compleja (LayoutByFold, FixedComposition, clásico) | `CreateDraftAlzadoPrimerDiedro` delega a `DraftGenerator.CreateAutomaticDraftFromModel` |
| Varias ramas: SelectBaseViewForLayoutByFold, FixedCompositionLayout, PickFormatAndScale, etc. | Un único flujo en DraftGenerator |

El código antiguo permanece en `CreateDraftAlzadoPrimerDiedro_Legacy` (privado) como referencia.

---

## 6. Helpers antiguos que se pueden eliminar (cuando se confirme estabilidad)

- `SelectBaseViewForLayoutByFold`
- `InsertThreeViewsByFold` (si ya no se usa en ningún flujo activo)
- `LayoutByFoldBaseResult` y lógica asociada
- Código de `CreateDraftAlzadoPrimerDiedro_Legacy` cuando no se requiera fallback

---

## 7. Partes reutilizadas del código viejo

- **GetUsableAreaForTemplate** (LayoutEngine): Área útil por template (A2, A3, etc.)
- **EnsureFlatPatternReady** y **TryCreateFlatView_Safe** (CojonudoBestFit): Lógica de flat pattern
- **CreateFlatViewForDraft**: Nuevo wrapper público que une EnsureFlatPatternReady + TryCreateFlatView_Safe
- Constantes: escala estándar, gaps, factor ISO

---

## 8. Logs técnicos

El sistema escribe en consola:

- Orientación candidata y rotación
- Dimensiones a escala 1
- Dimensiones del bloque total
- Escala calculada
- Template y área útil
- Posiciones finales (base, right, down, iso, flat)
- Motivo de descarte de candidatos (si se añade logging detallado)

---

## 9. Archivos del proyecto

```
DraftLayoutTypes.vb   ← estructuras
DraftGenerator.vb     ← motor principal
LayoutEngine.vb       ← área útil
CojonudoBestFit_Bueno_02.vb  ← CreateFlatViewForDraft, delegación CreateDraftAlzadoPrimerDiedro
```
