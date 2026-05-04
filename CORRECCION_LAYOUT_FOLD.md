# Corrección Layout por Fold – Resumen

## 1. Qué estaba mal

### Error 1: Fórmulas de Y con scaleFactor
Se eliminaron las fórmulas que mezclaban dimensiones sin escalar, escaladas y coordenadas de hoja:
- `Y = (topEdge - baseH1) * scaleFactor`
- `Y = topEdge * scaleFactor - baseH`

**Causa**: La posición final debe depender del **Range real** de la vista insertada, no de fórmulas teóricas.

### Error 2: Doble layout
El flujo actual (`CreateDraftAlzadoPrimerDiedro` → `CreateAutomaticDraftFromModel`) **no** llamaba a `ApplyLayout_IsoTop_FlatBottom` ni a `CenterAllViewsOnUsableArea`. Esas funciones solo existen en `CreateDraftAlzadoPrimerDiedro_Legacy`, que no se utiliza.

El único post-proceso era `FitDraftView`, que solo hace **zoom** (`sw.Fit()`), no recolocación de vistas. Se ha documentado esto en los logs.

---

## 2. Cambios implementados

### InsertBaseView (DraftGenerator.vb)
- **Antes**: Inserción con centro calculado desde `layout.BaseWidth/Height`; luego `MoveViewTopLeft`.
- **Ahora**:
  1. Inserción provisional en posición fija `(0.15, 0.2)`.
  2. Se lee el **Range real** de la vista.
  3. `MoveViewTopLeft(baseView, leftEdge, topEdge)` para fijar la esquina superior izquierda.
  4. Se vuelve a leer el Range para comprobar la posición final.

### Parámetros
- `leftEdge`, `topEdge`: objetivos en coordenadas reales de hoja (ej. 40 mm, 260 mm).
- Ya no se usan fórmulas con `scaleFactor` para la posición de la base.

### Right y Down
- Se insertan con `AddByFold` y se posicionan con `MoveViewTopLeft`.
- Se registran sus posiciones finales con `LogViewTopLeft`.

---

## 3. Flujo condicionado / desactivado

- **No se aplica**: `ApplyLayout_IsoTop_FlatBottom`, `CenterAllViewsOnUsableArea` en el flujo principal.
- **FitDraftView**: Solo zoom; no recoloca vistas.
- Los logs indican explícitamente: `[FOLD] skipping global relayout for folded main block`.

---

## 4. Logs nuevos [FOLD]

```
[FOLD] usable area = MinX=... MaxX=... MinY=... MaxY=...mm
[FOLD] target top-left for base = (40,260)mm
[FOLD] base range before move = left=... top=...mm
[FOLD] base range after move = left=... top=...mm
[FOLD] base final top-left = (40,260)mm (target was 40,260)
[FOLD] right final top-left = (248,260)mm
[FOLD] down final top-left = (40,184)mm
[FOLD] skipping global relayout for folded main block
[FOLD] apply global layout only to iso/flat (already at fixed positions)
```

---

## 5. Comportamiento esperado

- La base queda con la esquina superior izquierda en `(40, 260)` mm.
- Right queda a la derecha de la base (ej. 248, 260 mm).
- Down queda debajo de la base (ej. 40, 184 mm).
- El bloque base/derecha/abajo no se recoloca tras su colocación.
- ISO y Flat conservan posiciones fijas (320, 140) y (35, 100) mm.
