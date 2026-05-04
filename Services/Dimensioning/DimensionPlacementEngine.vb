Option Strict Off

Imports System.Globalization
Imports System.Runtime.InteropServices
Imports Microsoft.VisualBasic
Imports SolidEdgeDraft
Imports SolidEdgeFrameworkSupport
Imports FrameworkDimension = SolidEdgeFrameworkSupport.Dimension

''' <summary>
''' Cotas con <see cref="Dimensions.AddDistanceBetweenObjects"/> entre dos objetos 2D de vista
''' (<see cref="DVLine2d"/>, arcos, círculos, …) con puntos de proximidad en hoja.
''' Las rutas <c>*ExtremePoints</c> usan <see cref="DimensionExtremePoint"/> (extremos min/max X/Y reales) en lugar de
''' “línea vertical izquierda vs derecha” o anclajes verticales ambiguos.
''' La API COM no expone cota pura punto-a-punto: siempre se pasan dos referencias de entidad + (x,y) de proximidad.
''' </summary>
Friend NotInheritable Class DimensionPlacementEngine

    ' --- BLOQUE TEMPORAL DIAGNÓSTICO (quitar o poner False al terminar pruebas) ---
    ''' <summary>Poner True solo para probar si el fallo es la colocación automática. No usar en producción.</summary>
    Friend Shared DebugForceDimensionPlacement As Boolean = False

    ''' <summary>Coordenadas de hoja (m) conocidas del plano de prueba — editar según necesidad.</summary>
    Private Const DebugForceOverallWidthY As Double = 0.252558R
    Private Const DebugForceOverallHeightX As Double = 0.04905R
    Private Const DebugForceThicknessHorizontalY As Double = 0.252558R
    Private Const DebugForceThicknessVerticalX As Double = 0.04905R
    ' --- fin bloque temporal ---

    ''' <summary>Fracción histórica (legacy) usada antes para separar la cota de la geometría (hoja, m).</summary>
    Private Const PlacementOffsetFraction As Double = 0.1R
    Private Const PlacementOffsetFractionNew As Double = 0.025R
    Private Const PlacementOffsetMin As Double = 0.0008R
    Private Const PlacementOffsetMax As Double = 0.006R

    ''' <summary>Depuración: separación fija en hoja (m); sustituye offsets proporcionales minúsculos en features locales.</summary>
    Private Const OffsetFixedVisibleWidthM As Double = 0.01R
    Private Const OffsetFixedVisibleHeightM As Double = 0.01R
    Private Const OffsetFixedVisibleThicknessM As Double = 0.006R

    Private Sub New()
    End Sub

    ''' <param name="frame">Marco único de la vista base (origen = Range.MinX, Range.MinY).</param>
    Public Shared Function TryInsertHorizontalBetweenLines(
        dims As Dimensions,
        leftLine As DvLineSheetInfo,
        rightLine As DvLineSheetInfo,
        frame As ViewPlacementFrame,
        log As DimensionLogger,
        Optional dv As DrawingView = Nothing,
        Optional forcedYDimSheet As Nullable(Of Double) = Nothing) As Boolean

        If dims Is Nothing OrElse leftLine Is Nothing OrElse rightLine Is Nothing OrElse frame Is Nothing Then Return False
        If leftLine.Line Is Nothing OrElse rightLine.Line Is Nothing Then Return False
        If ReferenceEquals(leftLine.Line, rightLine.Line) Then
            log?.Err("Cota horizontal: la misma DVLine2d para ambos lados.")
            Return False
        End If

        Dim viewBox As ViewSheetBoundingBox = frame.GetSheetBoundingBox()
        Dim oldOffsetY As Double = viewBox.Height * PlacementOffsetFraction
        Dim offsetY As Double = ComputePlacementOffset(frame.Height, "LEGACY_HORIZONTAL_LINES", "horizontal")
        Dim localTopY As Double = frame.FromSheetY(viewBox.MaxY)
        Dim localDimY As Double = localTopY + offsetY
        Dim yDim As Double = If(forcedYDimSheet.HasValue, forcedYDimSheet.Value, frame.ToSheetY(localDimY))
        log?.Frame(String.Format(CultureInfo.InvariantCulture, "local->sheet Y: local={0:0.######} => sheet={1:0.######} (cota horizontal, borde superior vista + offset)", localDimY, yDim))
        log?.Info(String.Format(CultureInfo.InvariantCulture, "[DIM][PLACE][OLD] offset anterior={0:0.######}", oldOffsetY))
        log?.Info(String.Format(CultureInfo.InvariantCulture, "[DIM][PLACE][NEW] offset nuevo={0:0.######}", offsetY))
        log?.Info(String.Format(CultureInfo.InvariantCulture, "[DIM][PLACE][NEW] Y colocación final={0:0.######}", yDim))

        Dim lx1 As Double = frame.FromSheetX(leftLine.MidX)
        Dim lx2 As Double = frame.FromSheetX(rightLine.MidX)
        Dim x1 As Double = frame.ToSheetX(lx1)
        Dim x2 As Double = frame.ToSheetX(lx2)
        Dim ly1 As Double = frame.FromSheetY(yDim)
        log?.DimPlaceLine(String.Format(CultureInfo.InvariantCulture,
            "feature=HORIZONTAL_EXTERIOR_LINES local=({0:0.######},{1:0.######})-({2:0.######},{3:0.######}) sheet=({4:0.######},{5:0.######})-({6:0.######},{7:0.######})",
            lx1, ly1, lx2, ly1, x1, yDim, x2, yDim))
        Dim y1 As Double = yDim
        Dim y2 As Double = yDim

        log?.Info(String.Format(CultureInfo.InvariantCulture,
            "Cota horizontal: marco base alto_local={0:0.######} → offsetY={1:0.######} → Y_hoja={2:0.######}",
            frame.Height, offsetY, yDim))
        log?.Info(String.Format(CultureInfo.InvariantCulture,
            "Cota horizontal: puntos hoja solicitados ({0:0.######},{1:0.######})-({2:0.######},{3:0.######})",
            x1, y1, x2, y2))

        Dim ok As Boolean
        Dim methodUsed As String = Nothing
        ok = TryAddDistanceBetweenObjectsGeneric(dims, leftLine.Line, rightLine.Line, x1, y1, x2, y2, log, "horizontal", methodUsed, "HORIZONTAL_EXTERIOR_LINES", frame, dv)
        If ok Then
            log?.DimAssert("placement calculado con marco único=True")
            log?.Info("Método de inserción horizontal usado: " & methodUsed)
            log?.Info("Cota horizontal insertada OK (línea de cota fuera por encima del bbox de la vista base).")
        End If
        Return ok
    End Function

    ''' <summary>Cota vertical exterior: envolvente Y real (líneas + arcos + círculos), con fallback a horizontales extremas.</summary>
    Public Shared Function TryInsertVerticalExterior(
        dims As Dimensions,
        view As DrawingView,
        frame As ViewPlacementFrame,
        extreme As ExtremeDvLinesResult,
        log As DimensionLogger,
        Optional forcedXDimSheet As Nullable(Of Double) = Nothing) As Boolean

        If dims Is Nothing OrElse view Is Nothing OrElse log Is Nothing OrElse frame Is Nothing Then Return False

        Dim viewBox As ViewSheetBoundingBox = frame.GetSheetBoundingBox()
        Dim r As VerticalExteriorAnchors.ResolveResult = Nothing
        If Not VerticalExteriorAnchors.TryResolve(view, viewBox, extreme, log, r) OrElse Not r.Success Then
            log?.VertWarn("No se pudo resolver anclajes verticales; cota vertical omitida.")
            Return False
        End If

        Dim oldOffsetX As Double = viewBox.Width * PlacementOffsetFraction
        Dim offsetX As Double = ComputePlacementOffset(frame.Width, "LEGACY_VERTICAL_EXTERIOR", "vertical")
        Dim localRightX As Double = frame.FromSheetX(viewBox.MaxX)
        Dim localDimX As Double = localRightX + offsetX
        Dim xDim As Double = If(forcedXDimSheet.HasValue, forcedXDimSheet.Value, frame.ToSheetX(localDimX))
        log?.Frame(String.Format(CultureInfo.InvariantCulture, "local->sheet X: local={0:0.######} => sheet={1:0.######} (cota vertical exterior, borde derecho vista base + offset)", localDimX, xDim))
        log?.Vert(String.Format(CultureInfo.InvariantCulture, "[DIM][PLACE][OLD] offset anterior={0:0.######}", oldOffsetX))
        log?.Vert(String.Format(CultureInfo.InvariantCulture, "[DIM][PLACE][NEW] offset nuevo={0:0.######}", offsetX))
        log?.Vert(String.Format(CultureInfo.InvariantCulture, "[DIM][PLACE][NEW] X colocación final={0:0.######}", xDim))

        Dim ly1 As Double = frame.FromSheetY(r.Y1Sheet)
        Dim ly2 As Double = frame.FromSheetY(r.Y2Sheet)
        Dim lxDim As Double = frame.FromSheetX(xDim)
        log?.DimPlaceLine(String.Format(CultureInfo.InvariantCulture,
            "feature=VERTICAL_EXTERIOR local=({0:0.######},{1:0.######})-({2:0.######},{3:0.######}) sheet=({4:0.######},{5:0.######})-({6:0.######},{7:0.######})",
            lxDim, ly1, lxDim, ly2, xDim, r.Y1Sheet, xDim, r.Y2Sheet))

        Dim x1 As Double = xDim
        Dim y1 As Double = r.Y1Sheet
        Dim x2 As Double = xDim
        Dim y2 As Double = r.Y2Sheet

        log?.Vert(String.Format(CultureInfo.InvariantCulture,
            "offset X local = {0:0.######}m → X_hoja = ToSheetX(derecha_vista_local + offset) = {1:0.######}m",
            offsetX, xDim))
        log?.Vert(String.Format(CultureInfo.InvariantCulture,
            "puntos hoja solicitados ({0:0.######},{1:0.######})-({2:0.######},{3:0.######})",
            x1, y1, x2, y2))

        Dim methodUsed As String = Nothing
        Dim ok As Boolean = TryAddDistanceBetweenObjectsGeneric(
            dims, r.BottomObject, r.TopObject, x1, y1, x2, y2, log, "vertical", methodUsed, "VERTICAL_EXTERIOR", frame, view)

        If ok Then
            log?.DimAssert("placement calculado con marco único=True")
            log?.Vert("Método inserción: " & If(methodUsed, "(n/d)"))
            log?.Vert("cota vertical insertada OK")
            Return True
        End If

        If Not r.UsedFallback AndAlso extreme IsNot Nothing AndAlso
            extreme.BottomHorizontal IsNot Nothing AndAlso extreme.TopHorizontal IsNot Nothing Then
            log?.VertFallback("COM rechazó la selección geométrica; reintento con líneas horizontales extremas.")
            Return TryInsertVerticalBetweenLines(dims, extreme.BottomHorizontal, extreme.TopHorizontal, frame, log, view, forcedXDimSheet)
        End If

        log?.VertWarn("Inserción vertical fallida.")
        Return False
    End Function

    ''' <param name="frame">Marco único vista base.</param>
    ''' <remarks>Reservado; la ruta activa es <see cref="TryInsertVerticalExterior"/>.</remarks>
    Public Shared Function TryInsertVerticalBetweenLines(
        dims As Dimensions,
        bottomLine As DvLineSheetInfo,
        topLine As DvLineSheetInfo,
        frame As ViewPlacementFrame,
        log As DimensionLogger,
        Optional dv As DrawingView = Nothing,
        Optional forcedXDimSheet As Nullable(Of Double) = Nothing) As Boolean

        If dims Is Nothing OrElse bottomLine Is Nothing OrElse topLine Is Nothing OrElse frame Is Nothing Then Return False
        If bottomLine.Line Is Nothing OrElse topLine.Line Is Nothing Then Return False
        If ReferenceEquals(bottomLine.Line, topLine.Line) Then
            log?.Err("Cota vertical: la misma DVLine2d para ambos lados.")
            Return False
        End If

        Dim viewBox As ViewSheetBoundingBox = frame.GetSheetBoundingBox()
        Dim oldOffsetX As Double = viewBox.Width * PlacementOffsetFraction
        Dim offsetX As Double = ComputePlacementOffset(frame.Width, "LEGACY_VERTICAL_LINES", "vertical")
        Dim localRightX As Double = frame.FromSheetX(viewBox.MaxX)
        Dim localDimX As Double = localRightX + offsetX
        Dim xDim As Double = If(forcedXDimSheet.HasValue, forcedXDimSheet.Value, frame.ToSheetX(localDimX))
        log?.Frame(String.Format(CultureInfo.InvariantCulture, "local->sheet X: local={0:0.######} => sheet={1:0.######} (cota vertical líneas, borde derecho vista base + offset)", localDimX, xDim))
        log?.Info(String.Format(CultureInfo.InvariantCulture, "[DIM][PLACE][OLD] offset anterior={0:0.######}", oldOffsetX))
        log?.Info(String.Format(CultureInfo.InvariantCulture, "[DIM][PLACE][NEW] offset nuevo={0:0.######}", offsetX))
        log?.Info(String.Format(CultureInfo.InvariantCulture, "[DIM][PLACE][NEW] X colocación final={0:0.######}", xDim))

        Dim ly1 As Double = frame.FromSheetY(bottomLine.MidY)
        Dim ly2 As Double = frame.FromSheetY(topLine.MidY)
        Dim lxDim As Double = frame.FromSheetX(xDim)
        log?.DimPlaceLine(String.Format(CultureInfo.InvariantCulture,
            "feature=VERTICAL_FALLBACK_LINES local=({0:0.######},{1:0.######})-({2:0.######},{3:0.######}) sheet=({4:0.######},{5:0.######})-({6:0.######},{7:0.######})",
            lxDim, ly1, lxDim, ly2, xDim, bottomLine.MidY, xDim, topLine.MidY))

        Dim x1 As Double = xDim
        Dim y1 As Double = bottomLine.MidY
        Dim x2 As Double = xDim
        Dim y2 As Double = topLine.MidY

        log?.Info(String.Format(CultureInfo.InvariantCulture,
            "Cota vertical: ancho_local_vista={0:0.######}m → offsetX={1:0.######}m → X_hoja={2:0.######}m",
            frame.Width, offsetX, xDim))
        log?.Info(String.Format(CultureInfo.InvariantCulture,
            "Cota vertical: puntos hoja solicitados ({0:0.######},{1:0.######})-({2:0.######},{3:0.######})",
            x1, y1, x2, y2))

        Dim ok As Boolean
        Dim methodUsed As String = Nothing
        ok = TryAddDistanceBetweenObjectsGeneric(dims, bottomLine.Line, topLine.Line, x1, y1, x2, y2, log, "vertical", methodUsed, "VERTICAL_FALLBACK_LINES", frame, dv)
        If ok Then
            log?.DimAssert("placement calculado con marco único=True")
            log?.Info("Método de inserción vertical usado: " & methodUsed)
            log?.Info("Cota vertical insertada OK (línea de cota fuera a la derecha del bbox de la vista base).")
        End If
        Return ok
    End Function

    ''' <summary>
    ''' Inserta cota horizontal usando dos entidades COM cualesquiera (línea/arco/círculo/...) y puntos de proximidad en hoja.
    ''' Mantiene el backend AddDistanceBetweenObjects.
    ''' </summary>
    Public Shared Function TryInsertHorizontalBetweenObjects(
        dims As Dimensions,
        leftObj As Object,
        rightObj As Object,
        frame As ViewPlacementFrame,
        leftPickX As Double,
        rightPickX As Double,
        log As DimensionLogger,
        Optional dv As DrawingView = Nothing,
        Optional featureName As String = "UNSPEC") As Boolean

        If dims Is Nothing OrElse leftObj Is Nothing OrElse rightObj Is Nothing OrElse frame Is Nothing Then Return False
        If ReferenceEquals(leftObj, rightObj) Then
            log?.Err("Cota horizontal: la misma entidad para ambos lados.")
            Return False
        End If

        Dim viewBox As ViewSheetBoundingBox = frame.GetSheetBoundingBox()
        Dim localBox As New ViewSheetBoundingBox()
        Dim hasLocal As Boolean = TryComputeLocalPairBBoxSheet(dv, leftObj, rightObj, localBox, featureName, leftPickX, rightPickX, log)
        Dim localMandatory As Boolean = IsLocalPlacementMandatory(featureName)
        If localMandatory AndAlso Not hasLocal Then
            log?.Info("[DIM][PLACE][MODE] " & featureName & "=LOCAL")
            log?.Err("[DIM][PLACE][LOCAL][ERR] no se pudo calcular bbox local para feature obligatoria.")
            If String.Equals(featureName, "OVERALL_WIDTH", StringComparison.OrdinalIgnoreCase) Then
                log?.Err("[DIM][WIDTH][LOCAL] fallback=True motivo=bbox_local_invalido")
            End If
            Return False
        End If
        Dim refBox As ViewSheetBoundingBox = If(hasLocal, localBox, viewBox)

        Dim oldOffsetY As Double = refBox.Height * PlacementOffsetFraction
        Dim yDim As Double
        Dim offsetY As Double
        Dim refTopLocalY As Double = frame.FromSheetY(refBox.MaxY)
        If String.Equals(featureName, "OVERALL_WIDTH", StringComparison.OrdinalIgnoreCase) OrElse
           String.Equals(featureName, "THICKNESS", StringComparison.OrdinalIgnoreCase) Then
            LogFixedVisibleOffsetBanner(log)
            If String.Equals(featureName, "OVERALL_WIDTH", StringComparison.OrdinalIgnoreCase) Then
                offsetY = OffsetFixedVisibleWidthM
            Else
                offsetY = OffsetFixedVisibleThicknessM
            End If
            yDim = frame.ToSheetY(refTopLocalY + offsetY)
        Else
            Dim spanLocalY As Double = Math.Max(refBox.Height, 1.0E-12)
            offsetY = ComputePlacementOffset(spanLocalY, featureName, "horizontal")
            yDim = frame.ToSheetY(refTopLocalY + offsetY)
        End If

        If DebugForceDimensionPlacement Then
            If String.Equals(featureName, "OVERALL_WIDTH", StringComparison.OrdinalIgnoreCase) Then
                log?.Info("[DIM][FORCE] DebugForceDimensionPlacement=True")
                log?.Info(String.Format(CultureInfo.InvariantCulture, "[DIM][FORCE] OVERALL_WIDTH Y fija={0:0.######}", DebugForceOverallWidthY))
                yDim = DebugForceOverallWidthY
                offsetY = yDim - refBox.MaxY
            ElseIf String.Equals(featureName, "THICKNESS", StringComparison.OrdinalIgnoreCase) Then
                log?.Info("[DIM][FORCE] DebugForceDimensionPlacement=True")
                log?.Info(String.Format(CultureInfo.InvariantCulture, "[DIM][FORCE] THICKNESS posición fija=(Y_cota_horizontal={0:0.######})", DebugForceThicknessHorizontalY))
                yDim = DebugForceThicknessHorizontalY
                offsetY = yDim - refBox.MaxY
            End If
        End If

        log?.Frame(String.Format(CultureInfo.InvariantCulture, "local->sheet Y: local={0:0.######} => sheet={1:0.######} (feature={2}, borde superior ref + offset)",
            frame.FromSheetY(yDim), yDim, featureName))
        log?.Info(String.Format(CultureInfo.InvariantCulture, "[DIM][PLACE][OLD] offset anterior={0:0.######}", oldOffsetY))
        log?.Info(String.Format(CultureInfo.InvariantCulture, "[DIM][PLACE][NEW] offset nuevo={0:0.######}", offsetY))
        log?.Info(String.Format(CultureInfo.InvariantCulture, "[DIM][PLACE][NEW] Y colocación final={0:0.######}", yDim))
        If hasLocal Then
            log?.Info("[DIM][PLACE][MODE] " & featureName & "=LOCAL")
            log?.Info(String.Format(CultureInfo.InvariantCulture,
                "[DIM][PLACE][LOCAL] feature={0} bbox par (hoja)=({1:0.######},{2:0.######})-({3:0.######},{4:0.######}) local_ref=({5:0.######},{6:0.######})-({7:0.######},{8:0.######})",
                featureName, localBox.MinX, localBox.MinY, localBox.MaxX, localBox.MaxY,
                frame.FromSheetX(localBox.MinX), frame.FromSheetY(localBox.MinY),
                frame.FromSheetX(localBox.MaxX), frame.FromSheetY(localBox.MaxY)))
        End If
        log?.Info(String.Format(CultureInfo.InvariantCulture,
            "[DIM][PLACE][FINAL] feature={0} Y colocación final={1:0.######}",
            featureName, yDim))

        Dim lxPick1 As Double = frame.FromSheetX(leftPickX)
        Dim lxPick2 As Double = frame.FromSheetX(rightPickX)
        Dim x1 As Double = frame.ToSheetX(lxPick1)
        Dim x2 As Double = frame.ToSheetX(lxPick2)
        Dim lyDim As Double = frame.FromSheetY(yDim)
        log?.Frame(String.Format(CultureInfo.InvariantCulture, "local->sheet X: local={0:0.######} => sheet={1:0.######} (pick izq)", lxPick1, x1))
        log?.Frame(String.Format(CultureInfo.InvariantCulture, "local->sheet X: local={0:0.######} => sheet={1:0.######} (pick der)", lxPick2, x2))
        log?.DimPlaceLine(String.Format(CultureInfo.InvariantCulture,
            "feature={0} local=({1:0.######},{2:0.######})-({3:0.######},{4:0.######}) sheet=({5:0.######},{6:0.######})-({7:0.######},{8:0.######})",
            featureName, lxPick1, lyDim, lxPick2, lyDim, x1, yDim, x2, yDim))
        Dim y1 As Double = yDim
        Dim y2 As Double = yDim

        LogStrictPairPlacementSource(log, featureName, leftObj, rightObj, localBox, hasLocal, frame,
                                     String.Format(CultureInfo.InvariantCulture, "pick_sig_X1={0:0.######} pick_sig_X2={1:0.######}", leftPickX, rightPickX),
                                     x1, y1, x2, y2, horizontal:=True)

        log?.Info(String.Format(CultureInfo.InvariantCulture,
            "Cota horizontal (GEOM): ref_top_localY={0:0.######} alto_ref={1:0.######}m → Y_hoja={2:0.######}m",
            refTopLocalY, refBox.Height, yDim))
        log?.Info("Cota horizontal (GEOM): Object1=" & GetObjectTypeName(leftObj) & " Object2=" & GetObjectTypeName(rightObj))
        log?.Info(String.Format(CultureInfo.InvariantCulture,
            "Cota horizontal (GEOM): pick1=({0:0.######},{1:0.######}) pick2=({2:0.######},{3:0.######})",
            x1, y1, x2, y2))

        Dim methodUsed As String = Nothing
        Dim ok As Boolean = TryAddDistanceBetweenObjectsGeneric(dims, leftObj, rightObj, x1, y1, x2, y2, log, "horizontal", methodUsed, featureName, frame, dv)

        If String.Equals(featureName, "OVERALL_WIDTH", StringComparison.OrdinalIgnoreCase) Then
            If hasLocal Then
                log?.Info(String.Format(CultureInfo.InvariantCulture,
                    "[DIM][WIDTH][LOCAL] bbox local aceptado=({0:0.######},{1:0.######})-({2:0.######},{3:0.######})",
                    localBox.MinX, localBox.MinY, localBox.MaxX, localBox.MaxY))
                log?.Info(String.Format(CultureInfo.InvariantCulture, "[DIM][WIDTH][LOCAL] Y colocación final={0:0.######}", yDim))
            End If
            If ok Then
                log?.Info("[DIM][WIDTH][LOCAL] fallback=False")
            Else
                log?.Warn("[DIM][WIDTH][LOCAL] fallback=True motivo=insercion_COM_fallida")
            End If
        End If

        If ok Then
            log?.DimAssert("placement calculado con marco único=True")
            log?.Info("Método inserción horizontal (GEOM): " & methodUsed)
        End If
        Return ok
    End Function

    ''' <summary>
    ''' Inserta cota vertical usando dos entidades COM cualesquiera (línea/arco/círculo/...) y puntos de proximidad en hoja.
    ''' </summary>
    Public Shared Function TryInsertVerticalBetweenObjects(
        dims As Dimensions,
        bottomObj As Object,
        topObj As Object,
        frame As ViewPlacementFrame,
        bottomPickY As Double,
        topPickY As Double,
        log As DimensionLogger,
        Optional dv As DrawingView = Nothing,
        Optional featureName As String = "UNSPEC") As Boolean

        If dims Is Nothing OrElse bottomObj Is Nothing OrElse topObj Is Nothing OrElse frame Is Nothing Then Return False
        If ReferenceEquals(bottomObj, topObj) Then
            log?.Err("Cota vertical: la misma entidad para ambos lados.")
            Return False
        End If

        Dim viewBox As ViewSheetBoundingBox = frame.GetSheetBoundingBox()
        Dim localBox As New ViewSheetBoundingBox()
        Dim hasLocal As Boolean = TryComputeLocalPairBBoxSheet(dv, bottomObj, topObj, localBox, featureName, bottomPickY, topPickY, log)
        Dim localMandatory As Boolean = IsLocalPlacementMandatory(featureName)
        If localMandatory AndAlso Not hasLocal Then
            log?.Vert("[DIM][PLACE][MODE] " & featureName & "=LOCAL")
            log?.Err("[DIM][PLACE][LOCAL][ERR] no se pudo calcular bbox local para feature obligatoria.")
            Return False
        End If
        Dim refBox As ViewSheetBoundingBox = If(hasLocal, localBox, viewBox)

        Dim oldOffsetX As Double = refBox.Width * PlacementOffsetFraction
        Dim xDim As Double
        Dim offsetX As Double
        Dim refLeftLocalX As Double = frame.FromSheetX(refBox.MinX)
        Dim refRightLocalX As Double = frame.FromSheetX(refBox.MaxX)
        If String.Equals(featureName, "OVERALL_HEIGHT", StringComparison.OrdinalIgnoreCase) OrElse
           String.Equals(featureName, "THICKNESS", StringComparison.OrdinalIgnoreCase) Then
            LogFixedVisibleOffsetBanner(log)
            If String.Equals(featureName, "OVERALL_HEIGHT", StringComparison.OrdinalIgnoreCase) Then
                offsetX = OffsetFixedVisibleHeightM
                xDim = frame.ToSheetX(refLeftLocalX - offsetX)
            Else
                offsetX = OffsetFixedVisibleThicknessM
                xDim = frame.ToSheetX(refRightLocalX + offsetX)
            End If
        Else
            Dim spanLocalX As Double = Math.Max(refBox.Width, 1.0E-12)
            offsetX = ComputePlacementOffset(spanLocalX, featureName, "vertical")
            xDim = frame.ToSheetX(refRightLocalX + offsetX)
        End If

        If DebugForceDimensionPlacement Then
            If String.Equals(featureName, "OVERALL_HEIGHT", StringComparison.OrdinalIgnoreCase) Then
                log?.Vert("[DIM][FORCE] DebugForceDimensionPlacement=True")
                log?.Vert(String.Format(CultureInfo.InvariantCulture, "[DIM][FORCE] OVERALL_HEIGHT X fija={0:0.######}", DebugForceOverallHeightX))
                xDim = DebugForceOverallHeightX
                offsetX = Math.Abs(refBox.MinX - xDim)
            ElseIf String.Equals(featureName, "THICKNESS", StringComparison.OrdinalIgnoreCase) Then
                log?.Vert("[DIM][FORCE] DebugForceDimensionPlacement=True")
                log?.Vert(String.Format(CultureInfo.InvariantCulture, "[DIM][FORCE] THICKNESS posición fija=(X_cota_vertical={0:0.######})", DebugForceThicknessVerticalX))
                xDim = DebugForceThicknessVerticalX
                offsetX = xDim - refBox.MaxX
            End If
        End If

        log?.Frame(String.Format(CultureInfo.InvariantCulture, "local->sheet X: local={0:0.######} => sheet={1:0.######} (feature={2}, colocación vertical)",
            frame.FromSheetX(xDim), xDim, featureName))
        log?.Vert(String.Format(CultureInfo.InvariantCulture, "[DIM][PLACE][OLD] offset anterior={0:0.######}", oldOffsetX))
        log?.Vert(String.Format(CultureInfo.InvariantCulture, "[DIM][PLACE][NEW] offset nuevo={0:0.######}", offsetX))
        log?.Vert(String.Format(CultureInfo.InvariantCulture, "[DIM][PLACE][NEW] X colocación final={0:0.######}", xDim))
        If hasLocal Then
            log?.Vert("[DIM][PLACE][MODE] " & featureName & "=LOCAL")
            log?.Vert(String.Format(CultureInfo.InvariantCulture,
                "[DIM][PLACE][LOCAL] feature={0} bbox par (hoja)=({1:0.######},{2:0.######})-({3:0.######},{4:0.######}) local_ref=({5:0.######},{6:0.######})-({7:0.######},{8:0.######})",
                featureName, localBox.MinX, localBox.MinY, localBox.MaxX, localBox.MaxY,
                frame.FromSheetX(localBox.MinX), frame.FromSheetY(localBox.MinY),
                frame.FromSheetX(localBox.MaxX), frame.FromSheetY(localBox.MaxY)))
        End If
        log?.Vert(String.Format(CultureInfo.InvariantCulture,
            "[DIM][PLACE][FINAL] feature={0} X colocación final={1:0.######}",
            featureName, xDim))

        Dim ly1 As Double = frame.FromSheetY(bottomPickY)
        Dim ly2 As Double = frame.FromSheetY(topPickY)
        Dim lxDim As Double = frame.FromSheetX(xDim)
        log?.Frame(String.Format(CultureInfo.InvariantCulture, "local->sheet Y: local={0:0.######} => sheet={1:0.######} (pick inferior)", ly1, bottomPickY))
        log?.Frame(String.Format(CultureInfo.InvariantCulture, "local->sheet Y: local={0:0.######} => sheet={1:0.######} (pick superior)", ly2, topPickY))
        log?.DimPlaceLine(String.Format(CultureInfo.InvariantCulture,
            "feature={0} local=({1:0.######},{2:0.######})-({3:0.######},{4:0.######}) sheet=({5:0.######},{6:0.######})-({7:0.######},{8:0.######})",
            featureName, lxDim, ly1, lxDim, ly2, xDim, bottomPickY, xDim, topPickY))

        Dim x1 As Double = xDim
        Dim y1 As Double = bottomPickY
        Dim x2 As Double = xDim
        Dim y2 As Double = topPickY

        LogStrictPairPlacementSource(log, featureName, bottomObj, topObj, localBox, hasLocal, frame,
                                     String.Format(CultureInfo.InvariantCulture, "pick_sig_Y1={0:0.######} pick_sig_Y2={1:0.######}", bottomPickY, topPickY),
                                     x1, y1, x2, y2, horizontal:=False)

        log?.Vert(String.Format(CultureInfo.InvariantCulture,
            "Cota vertical (GEOM): ref_right_localX={0:0.######} ancho_ref={1:0.######}m → X_hoja={2:0.######}m",
            refRightLocalX, refBox.Width, xDim))
        log?.Vert("Cota vertical (GEOM): Object1=" & GetObjectTypeName(bottomObj) & " Object2=" & GetObjectTypeName(topObj))
        log?.Vert(String.Format(CultureInfo.InvariantCulture,
            "Cota vertical (GEOM): pick1=({0:0.######},{1:0.######}) pick2=({2:0.######},{3:0.######})",
            x1, y1, x2, y2))

        Dim methodUsed As String = Nothing
        Dim ok As Boolean = TryAddDistanceBetweenObjectsGeneric(dims, bottomObj, topObj, x1, y1, x2, y2, log, "vertical", methodUsed, featureName, frame, dv)
        If ok Then
            log?.DimAssert("placement calculado con marco único=True")
            log?.Vert("Método inserción vertical (GEOM): " & methodUsed)
        End If
        Return ok
    End Function

    ''' <summary>
    ''' Cota horizontal exterior entre extremos X reales (izquierda/derecha) con línea de cota por encima del bbox de la vista base.
    ''' Object1/Object2 = entidades que aportaron esos extremos; proximidad en X = X del extremo, en Y = banda superior (marco único).
    ''' </summary>
    Public Shared Function TryInsertHorizontalExteriorFromExtremePoints(
        dims As Dimensions,
        leftPt As DimensionExtremePoint,
        rightPt As DimensionExtremePoint,
        topPt As DimensionExtremePoint,
        dv As DrawingView,
        frame As ViewPlacementFrame,
        log As DimensionLogger,
        Optional forcedYDimSheet As Nullable(Of Double) = Nothing) As Boolean

        If dims Is Nothing OrElse leftPt Is Nothing OrElse rightPt Is Nothing OrElse frame Is Nothing Then Return False
        If leftPt.SourceObject Is Nothing OrElse rightPt.SourceObject Is Nothing Then
            log?.Err("Cota horizontal extremos: SourceObject nulo en left/right.")
            Return False
        End If

        log?.Info("[DIM][API] AddDistanceBetweenObjects requiere dos objetos COM; los extremos fijan la proximidad (X real de izq/der). No hay dimensión solo entre puntos libres sin entidad.")
        log?.LogLine("[DIM][COORD] exterior_horizontal strategy=" & DimensionInsertionConfigFormat.StrategyName(DimensionInsertionConfig.ActiveExteriorPickStrategy))

        Dim viewBox As ViewSheetBoundingBox = frame.GetSheetBoundingBox()
        Dim strat As ExteriorPickStrategy = DimensionInsertionConfig.ActiveExteriorPickStrategy
        Dim yDim As Double = If(forcedYDimSheet.HasValue, forcedYDimSheet.Value, ComputeHorizontalExteriorYDim(strat, frame, viewBox, topPt, log))
        Dim offRng As Double = ComputeDimensionClearance(frame.Height, "horizontal")

        Dim x1 As Double = leftPt.XSheet
        Dim x2 As Double = rightPt.XSheet
        Dim y1 As Double = yDim
        Dim y2 As Double = yDim

        log?.Frame(String.Format(CultureInfo.InvariantCulture,
            "horizontal exterior: Range centro hoja=({0:0.######},{1:0.######}) MaxY={2:0.######} + offset_marco={3:0.######} => Y línea cota={4:0.######}",
            frame.CenterX, frame.CenterY, frame.MaxY, offRng, yDim))
        log?.DimPlaceLine(String.Format(CultureInfo.InvariantCulture,
            "feature=HORIZONTAL_EXTREME_POINTS local=({0:0.######},{1:0.######})-({2:0.######},{3:0.######}) sheet=({4:0.######},{5:0.######})-({6:0.######},{7:0.######})",
            frame.FromSheetX(x1), frame.FromSheetY(y1), frame.FromSheetX(x2), frame.FromSheetY(y2), x1, y1, x2, y2))
        log?.Info(String.Format(CultureInfo.InvariantCulture,
            "Cota horizontal extremos: left={0} right={1} | picks hoja ({2:0.######},{3:0.######})-({4:0.######},{5:0.######})",
            leftPt.FormatOneLine(), rightPt.FormatOneLine(), x1, y1, x2, y2))
        If forcedYDimSheet.HasValue Then
            log?.LogLine("[DIM][COORD] HORIZONTAL_EXTREME_POINTS yDim=FORCED_GAP_BETWEEN_VIEWS=" & yDim.ToString("0.######", CultureInfo.InvariantCulture))
        Else
            log?.LogLine("[DIM][COORD] HORIZONTAL_EXTREME_POINTS yDim=Range.MaxY+ExteriorOffset estrategia=" & DimensionInsertionConfigFormat.StrategyName(strat))
        End If

        Dim methodUsed As String = Nothing
        Return TryAddDistanceBetweenObjectsGeneric(dims, leftPt.SourceObject, rightPt.SourceObject, x1, y1, x2, y2, log, "horizontal", methodUsed, "HORIZONTAL_EXTREME_POINTS", frame, dv)
    End Function

    Private Shared Function ComputeHorizontalExteriorYDim(
        strategy As ExteriorPickStrategy,
        frame As ViewPlacementFrame,
        viewBox As ViewSheetBoundingBox,
        topPt As DimensionExtremePoint,
        log As DimensionLogger) As Double

        ' Colocación respecto al marco oficial de la vista en HOJA (DrawingView.Range): borde superior + clearance (ajustado).
        Dim off As Double = ComputeDimensionClearance(frame.Height, "horizontal")
        Dim y As Double = frame.MaxY + off
        If strategy = ExteriorPickStrategy.SheetOffsetBandA Then
            ' Modo legado: offset proporcional (solo si se fuerza estrategia A).
            Dim offsetY As Double = ComputeDimensionClearance(frame.Height, "horizontal")
            Dim localTopY As Double = frame.FromSheetY(viewBox.MaxY)
            y = frame.ToSheetY(localTopY + offsetY)
            log?.Info(String.Format(CultureInfo.InvariantCulture, "[DIM][PLACE] horizontal exterior estrategia A: Y={0:0.######} (proporcional)", y))
        Else
            log?.Info(String.Format(CultureInfo.InvariantCulture, "[DIM][PLACE] horizontal exterior: Y=MaxY+{0:0.######}m = {1:0.######}", off, y))
            If topPt IsNot Nothing Then
                log?.Info(String.Format(CultureInfo.InvariantCulture, "[DIM][PLACE] tope geométrico real Y={0:0.######} (solo referencia vs MaxY={1:0.######})", topPt.YSheet, frame.MaxY))
            End If
        End If
        Return y
    End Function

    ''' <summary>
    ''' Cota vertical exterior entre extremos Y reales (inferior/superior); línea de cota a la derecha del bbox (marco único).
    ''' </summary>
    Public Shared Function TryInsertVerticalExteriorFromExtremePoints(
        dims As Dimensions,
        bottomPt As DimensionExtremePoint,
        topPt As DimensionExtremePoint,
        rightPt As DimensionExtremePoint,
        dv As DrawingView,
        frame As ViewPlacementFrame,
        log As DimensionLogger,
        Optional forcedXDimSheet As Nullable(Of Double) = Nothing) As Boolean

        If dims Is Nothing OrElse bottomPt Is Nothing OrElse topPt Is Nothing OrElse frame Is Nothing Then Return False
        If bottomPt.SourceObject Is Nothing OrElse topPt.SourceObject Is Nothing Then
            log?.Err("Cota vertical extremos: SourceObject nulo en bottom/top.")
            Return False
        End If

        log?.Info("[DIM][API] AddDistanceBetweenObjects: proximidad Y real de inferior/superior; X = cota a la derecha (offset o borde derecho real + inset).")
        log?.LogLine("[DIM][COORD] exterior_vertical strategy=" & DimensionInsertionConfigFormat.StrategyName(DimensionInsertionConfig.ActiveExteriorPickStrategy))

        Dim viewBox As ViewSheetBoundingBox = frame.GetSheetBoundingBox()
        Dim strat As ExteriorPickStrategy = DimensionInsertionConfig.ActiveExteriorPickStrategy
        Dim xDim As Double = If(forcedXDimSheet.HasValue, forcedXDimSheet.Value, ComputeVerticalExteriorXDim(strat, frame, viewBox, rightPt, log))
        Dim offRng As Double = ComputeDimensionClearance(frame.Width, "vertical")

        Dim x1 As Double = xDim
        Dim y1 As Double = bottomPt.YSheet
        Dim x2 As Double = xDim
        Dim y2 As Double = topPt.YSheet

        log?.Frame(String.Format(CultureInfo.InvariantCulture,
            "vertical exterior: Range centro hoja=({0:0.######},{1:0.######}) MaxX={2:0.######} + offset_marco={3:0.######} => X línea cota={4:0.######}",
            frame.CenterX, frame.CenterY, frame.MaxX, offRng, xDim))
        log?.DimPlaceLine(String.Format(CultureInfo.InvariantCulture,
            "feature=VERTICAL_EXTREME_POINTS local=({0:0.######},{1:0.######})-({2:0.######},{3:0.######}) sheet=({4:0.######},{5:0.######})-({6:0.######},{7:0.######})",
            frame.FromSheetX(x1), frame.FromSheetY(y1), frame.FromSheetX(x2), frame.FromSheetY(y2), x1, y1, x2, y2))
        log?.Vert(String.Format(CultureInfo.InvariantCulture,
            "Cota vertical extremos: bottom={0} top={1} | picks hoja ({2:0.######},{3:0.######})-({4:0.######},{5:0.######})",
            bottomPt.FormatOneLine(), topPt.FormatOneLine(), x1, y1, x2, y2))
        If forcedXDimSheet.HasValue Then
            log?.LogLine("[DIM][COORD] VERTICAL_EXTREME_POINTS xDim=FORCED_BASE_FRAME=" & xDim.ToString("0.######", CultureInfo.InvariantCulture))
        Else
            log?.LogLine("[DIM][COORD] VERTICAL_EXTREME_POINTS xDim=Range.MaxX+ExteriorOffset estrategia=" & DimensionInsertionConfigFormat.StrategyName(strat))
        End If

        Dim methodUsed As String = Nothing
        Return TryAddDistanceBetweenObjectsGeneric(dims, bottomPt.SourceObject, topPt.SourceObject, x1, y1, x2, y2, log, "vertical", methodUsed, "VERTICAL_EXTREME_POINTS", frame, dv)
    End Function

    Private Shared Function ComputeVerticalExteriorXDim(
        strategy As ExteriorPickStrategy,
        frame As ViewPlacementFrame,
        viewBox As ViewSheetBoundingBox,
        rightPt As DimensionExtremePoint,
        log As DimensionLogger) As Double

        Dim off As Double = ComputeDimensionClearance(frame.Width, "vertical")
        Dim x As Double = frame.MaxX + off
        If strategy = ExteriorPickStrategy.SheetOffsetBandA Then
            Dim offsetX As Double = ComputeDimensionClearance(frame.Width, "vertical")
            Dim localRightX As Double = frame.FromSheetX(viewBox.MaxX)
            x = frame.ToSheetX(localRightX + offsetX)
            log?.Vert(String.Format(CultureInfo.InvariantCulture, "[DIM][PLACE] vertical exterior estrategia A: X={0:0.######} (proporcional)", x))
        Else
            log?.Vert(String.Format(CultureInfo.InvariantCulture, "[DIM][PLACE] vertical exterior: X=MaxX+{0:0.######}m = {1:0.######}", off, x))
            If rightPt IsNot Nothing Then
                log?.Vert(String.Format(CultureInfo.InvariantCulture, "[DIM][PLACE] borde geométrico real X={0:0.######} (referencia vs MaxX={1:0.######})", rightPt.XSheet, frame.MaxX))
            End If
        End If
        Return x
    End Function

    ''' <summary>
    ''' API: AddDistanceBetweenObjects — intenta primero coordenadas de hoja; opcionalmente respaldo en espacio vista (SheetToView).
    ''' Ruta compartida con Iso129DimensioningService (ISO 129-1 pragmática).
    ''' </summary>
    Friend Shared Function TryAddDistanceBetweenObjectsGeneric(
        dims As Dimensions,
        o1 As Object,
        o2 As Object,
        x1 As Double, y1 As Double,
        x2 As Double, y2 As Double,
        log As DimensionLogger,
        axisLabel As String,
        ByRef methodUsedOut As String,
        Optional contextLabel As String = Nothing,
        Optional frame As ViewPlacementFrame = Nothing,
        Optional dv As DrawingView = Nothing,
        Optional skipInsertedDimensionSpatialSanity As Boolean = False,
        Optional ByRef createdDimensionOut As FrameworkDimension = Nothing) As Boolean

        createdDimensionOut = Nothing
        If o1 Is Nothing OrElse o2 Is Nothing Then Return False
        If DimensionInsertionConfig.EnableDimensionInsertionDiagnostics AndAlso dv IsNot Nothing AndAlso log IsNot Nothing Then
            Dim tag As String = If(String.IsNullOrEmpty(contextLabel), axisLabel, contextLabel)
            DimensionCoordinateDiagnostics.LogViewAndSheetContext(dv, frame, log, tag)
            DimensionCoordinateDiagnostics.LogPicksSheetAndView(dv, x1, y1, x2, y2, log, tag)
            DimensionCoordinateDiagnostics.LogRoundTrip(dv, x1, y1, log, tag & "_P1")
            DimensionCoordinateDiagnostics.LogRoundTrip(dv, x2, y2, log, tag & "_P2")
        End If
        log?.Info("[" & axisLabel.ToUpperInvariant() & "][COM] Probar AddDistanceBetweenObjects con Object1=" & GetObjectTypeName(o1) & " Object2=" & GetObjectTypeName(o2))

        Dim lastEx As Exception = Nothing

        ' 1) Camino principal: coordenadas de HOJA (AddDistanceBetweenObjects en pliego).
        log?.LogLine("[DIM][API] AddDistance* proximidad en HOJA (coordenadas absolutas de pliego).")
        If TryAddDistanceBetweenObjectsKpLoops(
            dims, o1, o2, x1, y1, x2, y2, log, axisLabel, methodUsedOut, contextLabel, frame, dv,
            x1, y1, x2, y2, "(sheet)", lastEx, skipInsertedDimensionSpatialSanity, createdDimensionOut) Then
            Return True
        End If

        ' 2) Respaldo: proximidad en espacio vista (SheetToView) — puede crear cotas con Range “raro”; solo si falla hoja.
        If dv IsNot Nothing AndAlso DimensionInsertionConfig.UseViewSpaceProximityForAddDistance Then
            Dim vx1 As Double = x1, vy1 As Double = y1, vx2 As Double = x2, vy2 As Double = y2
            Try
                dv.SheetToView(x1, y1, vx1, vy1)
                dv.SheetToView(x2, y2, vx2, vy2)
                log?.LogLine("[DIM][API] AddDistance* respaldo: proximidad en VISTA (SheetToView).")
                If TryAddDistanceBetweenObjectsKpLoops(
                    dims, o1, o2, vx1, vy1, vx2, vy2, log, axisLabel, methodUsedOut, contextLabel, frame, dv,
                    x1, y1, x2, y2, "(view)", lastEx, skipInsertedDimensionSpatialSanity, createdDimensionOut) Then
                    Return True
                End If
            Catch ex As Exception
                log?.Warn("[DIM][API] camino view-space omitido: " & ex.Message)
            End Try
        End If

        If lastEx IsNot Nothing Then
            log?.ComFail("AddDistanceBetweenObjects / EX (" & axisLabel & "), Object+Object", "Dimensions", lastEx)
        Else
            log?.Err("Cota " & axisLabel & ": API devolvió Nothing sin excepción.")
        End If
        methodUsedOut = Nothing
        Return False
    End Function

    ''' <summary>
    ''' Descarta cotas cuyo <c>Range</c> cae fuera de una banda esperada alrededor del marco de la vista.
    ''' Importante: por tus logs, <c>Dimension.Range</c> parece devolverse en view-space (donde SheetToView centra alrededor de 0,0 en X/Y),
    ''' por lo que el "marco esperado" se calcula en el mismo sistema usando <c>dv.SheetToView</c>.
    ''' </summary>
    Private Shared Function IsInsertedDimensionSpatiallySane(d As FrameworkDimension, frame As ViewPlacementFrame, dv As DrawingView, log As DimensionLogger) As Boolean
        If d Is Nothing OrElse frame Is Nothing Then Return True
        Try
            Dim rx1 As Double = 0, ry1 As Double = 0, rx2 As Double = 0, ry2 As Double = 0
            d.Range(rx1, ry1, rx2, ry2)
            Dim minX As Double = Math.Min(rx1, rx2)
            Dim maxX As Double = Math.Max(rx1, rx2)
            Dim minY As Double = Math.Min(ry1, ry2)
            Dim maxY As Double = Math.Max(ry1, ry2)

            ' Calculamos el bbox esperado en el mismo sistema de coordenadas que devuelve d.Range.
            Dim expMinX As Double = frame.MinX, expMaxX As Double = frame.MaxX
            Dim expMinY As Double = frame.MinY, expMaxY As Double = frame.MaxY
            Dim vspan As Double = Math.Max(frame.Width, frame.Height)

            If dv IsNot Nothing Then
                Dim v00x As Double = 0, v00y As Double = 0
                Dim v10x As Double = 0, v10y As Double = 0
                Dim v01x As Double = 0, v01y As Double = 0
                Dim v11x As Double = 0, v11y As Double = 0
                dv.SheetToView(frame.MinX, frame.MinY, v00x, v00y)
                dv.SheetToView(frame.MaxX, frame.MinY, v10x, v10y)
                dv.SheetToView(frame.MinX, frame.MaxY, v01x, v01y)
                dv.SheetToView(frame.MaxX, frame.MaxY, v11x, v11y)

                expMinX = Math.Min(Math.Min(v00x, v10x), Math.Min(v01x, v11x))
                expMaxX = Math.Max(Math.Max(v00x, v10x), Math.Max(v01x, v11x))
                expMinY = Math.Min(Math.Min(v00y, v10y), Math.Min(v01y, v11y))
                expMaxY = Math.Max(Math.Max(v00y, v10y), Math.Max(v01y, v11y))

                vspan = Math.Max(expMaxX - expMinX, expMaxY - expMinY)
            End If

            Dim margin As Double = Math.Max(vspan * 4.0R, 0.05R) + DimensionInsertionConfig.ExteriorDimensionOffsetFromRangeM * 10.0R
            Dim overlap As Boolean = Not (maxX < expMinX OrElse minX > expMaxX OrElse maxY < expMinY OrElse minY > expMaxY)
            If Not overlap Then
                log?.Warn("[DIM][SANITY] Range de cota fuera de banda alrededor del Range de vista; se descarta.")
                Return False
            End If
            If vspan > 1.0E-12 Then
                Dim dw As Double = maxX - minX
                Dim dh As Double = maxY - minY
                If dw > 120.0R * vspan OrElse dh > 120.0R * vspan Then
                    log?.Warn("[DIM][SANITY] Range de cota desmesurado; se descarta.")
                    Return False
                End If
            End If
            Return True
        Catch
            Return True
        End Try
    End Function

    ''' <summary>Prueba todas las combinaciones keyPoint y AddDistance / AddDistanceEX con los (x,y) ya en el espacio esperado por la API.</summary>
    Private Shared Function TryAddDistanceBetweenObjectsKpLoops(
        dims As Dimensions,
        o1 As Object,
        o2 As Object,
        px1 As Double,
        py1 As Double,
        px2 As Double,
        py2 As Double,
        log As DimensionLogger,
        axisLabel As String,
        ByRef methodUsedOut As String,
        contextLabel As String,
        frame As ViewPlacementFrame,
        dv As DrawingView,
        sheetX1 As Double,
        sheetY1 As Double,
        sheetX2 As Double,
        sheetY2 As Double,
        coordTag As String,
        ByRef lastEx As Exception,
        Optional skipInsertedDimensionSpatialSanity As Boolean = False,
        Optional ByRef createdDimensionOut As FrameworkDimension = Nothing) As Boolean

        Dim kps As Boolean() = {False, True}
        For Each kp1 In kps
            For Each kp2 In kps
                Try
                    Dim d As FrameworkDimension = dims.AddDistanceBetweenObjects(
                        o1, px1, py1, 0.0, kp1,
                        o2, px2, py2, 0.0, kp2)
                    TrySetConstraintFalse(d)
                    If d IsNot Nothing Then
                        If Not skipInsertedDimensionSpatialSanity AndAlso frame IsNot Nothing AndAlso Not IsInsertedDimensionSpatiallySane(d, frame, dv, log) Then
                            Try
                                d.Delete()
                            Catch
                            End Try
                            Continue For
                        End If
                        methodUsedOut = "AddDistanceBetweenObjects" & coordTag & "(kp1=" & kp1.ToString() & ",kp2=" & kp2.ToString() & ")"
                        log?.Info("[" & axisLabel.ToUpperInvariant() & "][COM] OK " & methodUsedOut)
                        If DimensionInsertionConfig.EnableDimensionInsertionDiagnostics AndAlso dv IsNot Nothing Then
                            DimensionCoordinateDiagnostics.LogCreatedDimensionState(d, log, If(contextLabel, axisLabel), sheetX1, sheetY1, sheetX2, sheetY2)
                        Else
                            TryLogInsertedDimensionState(d, frame, log, If(contextLabel, axisLabel))
                        End If
                        EnsureDrawingViewCropContainsDimensions(dv, d, axisLabel, frame, log)
                        createdDimensionOut = d
                        Return True
                    End If
                Catch ex As Exception
                    lastEx = ex
                    log?.Err("[" & axisLabel.ToUpperInvariant() & "][COM] EX AddDistanceBetweenObjects" & coordTag & "(kp1=" & kp1.ToString() & ",kp2=" & kp2.ToString() & "): " & FormatExceptionWithHresult(ex))
                End Try
            Next
        Next

        For Each kp1 In kps
            For Each kp2 In kps
                Try
                    Dim dEx As FrameworkDimension = dims.AddDistanceBetweenObjectsEX(
                        o1, px1, py1, 0.0, kp1, False,
                        o2, px2, py2, 0.0, kp2, False)
                    TrySetConstraintFalse(dEx)
                    If dEx IsNot Nothing Then
                        If Not skipInsertedDimensionSpatialSanity AndAlso frame IsNot Nothing AndAlso Not IsInsertedDimensionSpatiallySane(dEx, frame, dv, log) Then
                            Try
                                dEx.Delete()
                            Catch
                            End Try
                            Continue For
                        End If
                        methodUsedOut = "AddDistanceBetweenObjectsEX" & coordTag & "(kp1=" & kp1.ToString() & ",kp2=" & kp2.ToString() & ",tan=F)"
                        log?.Info("[" & axisLabel.ToUpperInvariant() & "][COM] OK " & methodUsedOut)
                        If DimensionInsertionConfig.EnableDimensionInsertionDiagnostics AndAlso dv IsNot Nothing Then
                            DimensionCoordinateDiagnostics.LogCreatedDimensionState(dEx, log, If(contextLabel, axisLabel), sheetX1, sheetY1, sheetX2, sheetY2)
                        Else
                            TryLogInsertedDimensionState(dEx, frame, log, If(contextLabel, axisLabel))
                        End If
                        EnsureDrawingViewCropContainsDimensions(dv, dEx, axisLabel, frame, log)
                        createdDimensionOut = dEx
                        Return True
                    End If
                Catch ex As Exception
                    lastEx = ex
                    log?.Err("[" & axisLabel.ToUpperInvariant() & "][COM] EX AddDistanceBetweenObjectsEX" & coordTag & "(kp1=" & kp1.ToString() & ",kp2=" & kp2.ToString() & "): " & FormatExceptionWithHresult(ex))
                End Try
            Next
        Next

        Return False
    End Function

    ''' <summary>
    ''' Solución actual: si las dimensiones se crean correctamente pero no se ven hasta que amplías manualmente,
    ''' el problema está en el recorte/navegación de la <see cref="DrawingView"/>.
    ''' Por eso SOLO ajustamos CropLeft/CropRight/CropTop/CropBottom (sin tocar dv.Range ni usar SetUserRange).
    ''' </summary>
    Private Shared Sub EnsureDrawingViewCropContainsDimensions(
        dv As DrawingView,
        d As FrameworkDimension,
        axisLabel As String,
        frame As ViewPlacementFrame,
        log As DimensionLogger)

        If dv Is Nothing OrElse d Is Nothing OrElse log Is Nothing Then Return

        Try
            '--- 1) Rangos en el mismo sistema de coordenadas usado por la vista ---
            ' Hipótesis (verificada con logs): CropLeft/CropRight/CropTop/CropBottom esperan
            ' coordenadas en el mismo sistema que DrawingView.Range y Dimension.Range (el 2D "view-space"
            ' de la DrawingView).
            Dim baseMinX As Double = 0, baseMinY As Double = 0, baseMaxX As Double = 0, baseMaxY As Double = 0
            dv.Range(baseMinX, baseMinY, baseMaxX, baseMaxY)

            Dim rx1 As Double = 0, ry1 As Double = 0, rx2 As Double = 0, ry2 As Double = 0
            d.Range(rx1, ry1, rx2, ry2)
            Dim dimMinX As Double = Math.Min(rx1, rx2)
            Dim dimMaxX As Double = Math.Max(rx1, rx2)
            Dim dimMinY As Double = Math.Min(ry1, ry2)
            Dim dimMaxY As Double = Math.Max(ry1, ry2)

            Dim combinedMinX As Double = Math.Min(baseMinX, dimMinX)
            Dim combinedMinY As Double = Math.Min(baseMinY, dimMinY)
            Dim combinedMaxX As Double = Math.Max(baseMaxX, dimMaxX)
            Dim combinedMaxY As Double = Math.Max(baseMaxY, dimMaxY)

            '--- 2) Padding pequeño (suficiente para flechas/texto/extensiones) ---
            Dim padX As Double = ComputeCropPaddingM(dv, axisLabel, frame, log, isHorizontal:=String.Equals(axisLabel, "horizontal", StringComparison.OrdinalIgnoreCase))
            Dim padY As Double = padX

            '--- 3) Crop actual ---
            Dim beforeLeft As Double = GetCropValue(dv, "CropLeft")
            Dim beforeRight As Double = GetCropValue(dv, "CropRight")
            Dim beforeTop As Double = GetCropValue(dv, "CropTop")
            Dim beforeBottom As Double = GetCropValue(dv, "CropBottom")

            log.Frame(String.Format(CultureInfo.InvariantCulture,
                "[DIM][CROP] baseView.Range before=({0:0.######},{1:0.######})-({2:0.######},{3:0.######}) dim.Range=({4:0.######},{5:0.######})-({6:0.######},{7:0.######})",
                baseMinX, baseMinY, baseMaxX, baseMaxY, dimMinX, dimMinY, dimMaxX, dimMaxY))

            log.Frame(String.Format(CultureInfo.InvariantCulture,
                "[DIM][CROP] combinedRange=({0:0.######},{1:0.######})-({2:0.######},{3:0.######}) padding=({4:0.######})",
                combinedMinX, combinedMinY, combinedMaxX, combinedMaxY, padX))

            log.Frame(String.Format(CultureInfo.InvariantCulture,
                "[DIM][CROP] before CropLeft={0:0.######} CropRight={1:0.######} CropTop={2:0.######} CropBottom={3:0.######}",
                beforeLeft, beforeRight, beforeTop, beforeBottom))

            ' Requisito del usuario: NO tocar la esquina inferior izquierda del view.
            ' Por tanto: CropLeft y CropBottom se mantienen (solo expandimos hacia +X y +Y).
            Dim afterLeft As Double = beforeLeft
            Dim afterBottom As Double = beforeBottom
            Dim afterRight As Double = Math.Max(beforeRight, combinedMaxX + padX)
            Dim afterTop As Double = Math.Max(beforeTop, combinedMaxY + padY)

            ' Evitar cambios nulos.
            Dim eps As Double = 1.0E-8R
            Dim changed As Boolean =
                Math.Abs(afterLeft - beforeLeft) > eps OrElse
                Math.Abs(afterRight - beforeRight) > eps OrElse
                Math.Abs(afterTop - beforeTop) > eps OrElse
                Math.Abs(afterBottom - beforeBottom) > eps

            If Not changed Then
                log.Frame("[DIM][CROP] updated=False (sin cambios significativos)")
                Return
            End If

            SetCropValue(dv, "CropLeft", afterLeft)
            SetCropValue(dv, "CropRight", afterRight)
            SetCropValue(dv, "CropTop", afterTop)
            SetCropValue(dv, "CropBottom", afterBottom)

            '--- 4) Actualización segura y logs ---
            Try : dv.Update() : Catch : End Try

            Dim afterLeft2 As Double = GetCropValue(dv, "CropLeft")
            Dim afterRight2 As Double = GetCropValue(dv, "CropRight")
            Dim afterTop2 As Double = GetCropValue(dv, "CropTop")
            Dim afterBottom2 As Double = GetCropValue(dv, "CropBottom")

            log.Frame(String.Format(CultureInfo.InvariantCulture,
                "[DIM][CROP] after CropLeft={0:0.######} CropRight={1:0.######} CropTop={2:0.######} CropBottom={3:0.######}",
                afterLeft2, afterRight2, afterTop2, afterBottom2))
            Dim eps2 As Double = 1.0E-8R
            Dim updated2 As Boolean =
                Math.Abs(afterLeft2 - beforeLeft) > eps2 OrElse
                Math.Abs(afterRight2 - beforeRight) > eps2 OrElse
                Math.Abs(afterTop2 - beforeTop) > eps2 OrElse
                Math.Abs(afterBottom2 - beforeBottom) > eps2

            If updated2 Then
                log.Frame("[DIM][CROP] updated=True")
            Else
                log.Frame("[DIM][CROP] updated=False (valores no cambiaron tras set)")
            End If

            Dim hClr As Double = If(frame IsNot Nothing, ComputeDimensionClearance(frame.Height, "horizontal"), padX)
            Dim vClr As Double = If(frame IsNot Nothing, ComputeDimensionClearance(frame.Width, "vertical"), padX)
            log.Frame(String.Format(CultureInfo.InvariantCulture,
                "[DIM][CLR] horizontal clearance={0:0.######} vertical clearance={1:0.######}", hClr, vClr))

        Catch ex As Exception
            log.Frame("[DIM][CROP] Error: " & ex.Message)
        End Try

    End Sub

    Private Shared Function GetCropValue(dv As DrawingView, propName As String) As Double
        Try
            Dim obj As Object = CallByName(dv, propName, CallType.Get)
            If obj Is Nothing Then Return 0
            Return CDbl(obj)
        Catch
            Return 0
        End Try
    End Function

    Private Shared Sub SetCropValue(dv As DrawingView, propName As String, value As Double)
        Try
            CallByName(dv, propName, CallType.Let, value)
        Catch
            ' No romper el flujo si Crop* no está disponible en esa versión/objeto.
        End Try
    End Sub

    Private Shared Function ComputeCropPaddingM(
        dv As DrawingView,
        axisLabel As String,
        frame As ViewPlacementFrame,
        log As DimensionLogger,
        isHorizontal As Boolean) As Double

        ' Padding típico para que flechas/textos/extensiones entren en el recorte.
        ' Se mantiene pequeño (no exagerado) y se acota.
        Dim basePad As Double = 0.003R

        If frame IsNot Nothing Then
            If isHorizontal Then
                Dim proposed As Double = ComputeDimensionClearance(frame.Height, "horizontal") * 0.40R
                basePad = Math.Max(basePad, proposed)
            Else
                Dim proposed As Double = ComputeDimensionClearance(frame.Width, "vertical") * 0.40R
                basePad = Math.Max(basePad, proposed)
            End If
        End If

        ' Asegura un mínimo y máximo.
        basePad = Math.Min(0.008R, Math.Max(0.002R, basePad))
        Return basePad
    End Function

    Private Shared Function ComputeDimensionClearance(viewSize As Double, orientation As String) As Double
        ' Clearance (offset) exterior más ajustado a tu caso:
        ' - horizontal: depende del alto de la vista (frame.Height)
        ' - vertical: depende del ancho de la vista (frame.Width)
        Dim minOff As Double = 0.004R
        Dim maxOff As Double = 0.009R
        Dim proposed As Double

        If String.Equals(orientation, "horizontal", StringComparison.OrdinalIgnoreCase) Then
            proposed = viewSize * 0.09R
        Else
            proposed = viewSize * 0.04R
        End If

        If proposed < minOff Then proposed = minOff
        If proposed > maxOff Then proposed = maxOff
        Return proposed
    End Function

    Private Shared Sub TryLogInsertedDimensionState(d As FrameworkDimension, frame As ViewPlacementFrame, log As DimensionLogger, context As String)
        If d Is Nothing OrElse log Is Nothing Then Return
        Dim ctx As String = If(String.IsNullOrWhiteSpace(context), "dimension", context)
        Try
            Dim xMn As Double = 0, yMn As Double = 0, xMx As Double = 0, yMx As Double = 0
            d.Range(xMn, yMn, xMx, yMx)
            Dim minX As Double = Math.Min(xMn, xMx)
            Dim maxX As Double = Math.Max(xMn, xMx)
            Dim minY As Double = Math.Min(yMn, yMx)
            Dim maxY As Double = Math.Max(yMn, yMx)
            If frame IsNot Nothing Then
                log.DimPost(String.Format(CultureInfo.InvariantCulture,
                    "{0} Range hoja=({1:0.######},{2:0.######})-({3:0.######},{4:0.######}) local=({5:0.######},{6:0.######})-({7:0.######},{8:0.######})",
                    ctx, minX, minY, maxX, maxY,
                    frame.FromSheetX(minX), frame.FromSheetY(minY),
                    frame.FromSheetX(maxX), frame.FromSheetY(maxY)))
            Else
                log.DimPost(String.Format(CultureInfo.InvariantCulture,
                    "{0} Range hoja=({1:0.######},{2:0.######})-({3:0.######},{4:0.######})",
                    ctx, minX, minY, maxX, maxY))
            End If
        Catch ex As Exception
            log.DimPost(ctx & " Range: no legible (" & ex.Message & ")")
        End Try

        Try
            Dim nKp As Integer = CInt(d.KeyPointCount)
            If nKp <= 0 Then Return
            Dim lim As Integer = Math.Min(nKp, 12)
            For i As Integer = 0 To lim - 1
                Dim px As Double = 0, py As Double = 0, pz As Double = 0
                Dim kpt As SolidEdgeConstants.KeyPointType
                Dim hdl As SolidEdgeConstants.HandleType
                d.GetKeyPoint(i, px, py, pz, kpt, hdl)
                If frame IsNot Nothing Then
                    log.DimPost(String.Format(CultureInfo.InvariantCulture,
                        "{0} KeyPoint[{1}] hoja=({2:0.######},{3:0.######}) local=({4:0.######},{5:0.######}) kpt={6} hdl={7}",
                        ctx, i, px, py, frame.FromSheetX(px), frame.FromSheetY(py), CInt(kpt), CInt(hdl)))
                Else
                    log.DimPost(String.Format(CultureInfo.InvariantCulture,
                        "{0} KeyPoint[{1}] hoja=({2:0.######},{3:0.######}) kpt={4} hdl={5}",
                        ctx, i, px, py, CInt(kpt), CInt(hdl)))
                End If
            Next
        Catch ex As Exception
            log.DimPost(ctx & " KeyPoints: no legible (" & ex.Message & ")")
        End Try

        Try
            Dim tox As Double = 0, toy As Double = 0
            d.GetTextOffsets(tox, toy)
            log.DimPost(String.Format(CultureInfo.InvariantCulture,
                "{0} GetTextOffsets (relativo a posición calculada del texto) dx={1:0.######} dy={2:0.######}",
                ctx, tox, toy))
        Catch ex As Exception
            log.DimPost(ctx & " GetTextOffsets: no legible (" & ex.Message & ")")
        End Try
    End Sub

    Private Shared Function GetObjectTypeName(obj As Object) As String
        If obj Is Nothing Then Return "(Nothing)"
        Try
            Return obj.GetType().Name
        Catch
            Return "(tipo no legible)"
        End Try
    End Function

    Private Shared Function FormatExceptionWithHresult(ex As Exception) As String
        If ex Is Nothing Then Return "(sin excepción)"
        Dim msg As String = ex.GetType().Name & ": " & ex.Message
        Dim cex As COMException = TryCast(ex, COMException)
        If cex IsNot Nothing Then
            msg &= " | HRESULT=0x" & cex.ErrorCode.ToString("X8", CultureInfo.InvariantCulture)
        End If
        Return msg
    End Function

    Private Shared Sub TrySetConstraintFalse(d As FrameworkDimension)
        If d Is Nothing Then Return
        Try
            d.Constraint = False
        Catch
        End Try
    End Sub

    Private Shared Sub LogFixedVisibleOffsetBanner(log As DimensionLogger)
        If log Is Nothing Then Return
        log.Info("[DIM][OFFSET][MODE] FIXED_VISIBLE")
        log.Info(String.Format(CultureInfo.InvariantCulture, "[DIM][OFFSET][WIDTH] {0:0.######}", OffsetFixedVisibleWidthM))
        log.Info(String.Format(CultureInfo.InvariantCulture, "[DIM][OFFSET][HEIGHT] {0:0.######}", OffsetFixedVisibleHeightM))
        log.Info(String.Format(CultureInfo.InvariantCulture, "[DIM][OFFSET][THICKNESS] {0:0.######}", OffsetFixedVisibleThicknessM))
    End Sub

    Private Shared Function ComputePlacementOffset(viewSpan As Double, featureName As String, axis As String) As Double
        Dim frac As Double = PlacementOffsetFractionNew
        Dim minOff As Double = PlacementOffsetMin
        Dim maxOff As Double = PlacementOffsetMax

        If String.Equals(featureName, "OVERALL_HEIGHT", StringComparison.OrdinalIgnoreCase) AndAlso String.Equals(axis, "vertical", StringComparison.OrdinalIgnoreCase) Then
            frac = 0.015R : minOff = 0.00045R : maxOff = 0.0018R
        ElseIf String.Equals(featureName, "THICKNESS", StringComparison.OrdinalIgnoreCase) Then
            frac = 0.01R : minOff = 0.00035R : maxOff = 0.0012R
        ElseIf String.Equals(featureName, "OVERALL_WIDTH", StringComparison.OrdinalIgnoreCase) AndAlso String.Equals(axis, "horizontal", StringComparison.OrdinalIgnoreCase) Then
            frac = 0.02R : minOff = 0.0006R : maxOff = 0.0025R
        End If

        Dim proposed As Double = viewSpan * frac
        If proposed < minOff Then Return minOff
        If proposed > maxOff Then Return maxOff
        Return proposed
    End Function

    Private Shared Function IsLocalPlacementMandatory(featureName As String) As Boolean
        If String.IsNullOrWhiteSpace(featureName) Then Return False
        Return String.Equals(featureName, "OVERALL_WIDTH", StringComparison.OrdinalIgnoreCase) OrElse
               String.Equals(featureName, "OVERALL_HEIGHT", StringComparison.OrdinalIgnoreCase) OrElse
               String.Equals(featureName, "THICKNESS", StringComparison.OrdinalIgnoreCase)
    End Function

    ''' <summary>
    ''' Traza [DIM][PLACE][SRC] / ASSERT / FIX: colocación estricta solo desde entity1+entity2, picks y bbox local (unión en hoja).
    ''' </summary>
    Private Shared Sub LogStrictPairPlacementSource(
        log As DimensionLogger,
        featureName As String,
        obj1 As Object,
        obj2 As Object,
        localBox As ViewSheetBoundingBox,
        hasLocal As Boolean,
        frame As ViewPlacementFrame,
        pickSigSummary As String,
        x1 As Double,
        y1 As Double,
        x2 As Double,
        y2 As Double,
        horizontal As Boolean)

        If log Is Nothing OrElse Not IsLocalPlacementMandatory(featureName) Then Return
        log.PlaceSrc("feature=" & featureName)
        log.PlaceSrc("entity1=" & DescribeEntityBrief(obj1))
        log.PlaceSrc("entity2=" & DescribeEntityBrief(obj2))
        log.PlaceSrc("point1=" & FormatPt(x1, y1) & If(horizontal, " (X pick→marco vista base + Y desde ref local)", " (X desde ref local + Y pick→marco vista base)"))
        log.PlaceSrc("point2=" & FormatPt(x2, y2))
        log.PlaceSrc("pick_significativos=" & pickSigSummary)
        If hasLocal Then
            log.PlaceSrc(String.Format(CultureInfo.InvariantCulture,
                "bbox par entidades (hoja)=({0:0.######},{1:0.######})-({2:0.######},{3:0.######})",
                localBox.MinX, localBox.MinY, localBox.MaxX, localBox.MaxY))
            If frame IsNot Nothing Then
                log.PlaceSrc(String.Format(CultureInfo.InvariantCulture,
                    "mismo bbox en marco vista base=({0:0.######},{1:0.######})-({2:0.######},{3:0.######})",
                    frame.FromSheetX(localBox.MinX), frame.FromSheetY(localBox.MinY),
                    frame.FromSheetX(localBox.MaxX), frame.FromSheetY(localBox.MaxY)))
            End If
        Else
            log.PlaceSrc("bbox par entidades=(n/d — feature no obligatoria local o fallo)")
        End If

        If hasLocal Then
            log.PlaceAssertMsg("placement calculado SOLO desde entidades de la feature = True")
            log.PlaceFixMsg("se eliminó referencia externa a la pareja geométrica")
        End If
    End Sub

    Private Shared Function FormatPt(x As Double, y As Double) As String
        Return String.Format(CultureInfo.InvariantCulture, "({0:0.######},{1:0.######})", x, y)
    End Function

    Private Shared Function DescribeEntityBrief(obj As Object) As String
        If obj Is Nothing Then Return "(Nothing)"
        Dim t As String = GetObjectTypeName(obj)
        Try
            If TypeOf obj Is DVLine2d Then
                Dim ln As DVLine2d = CType(obj, DVLine2d)
                Dim x1 As Double = 0, y1 As Double = 0, x2 As Double = 0, y2 As Double = 0
                ln.GetStartPoint(x1, y1)
                ln.GetEndPoint(x2, y2)
                Return t & " v=(" & FormatPt(x1, y1) & ")-(" & FormatPt(x2, y2) & ")"
            End If
        Catch
        End Try
        Return t
    End Function

    Private Shared Function TryComputeLocalPairBBoxSheet(
        dv As DrawingView,
        obj1 As Object,
        obj2 As Object,
        ByRef outBox As ViewSheetBoundingBox,
        featureName As String,
        pick1 As Double,
        pick2 As Double,
        log As DimensionLogger) As Boolean

        If dv Is Nothing OrElse obj1 Is Nothing OrElse obj2 Is Nothing Then Return False

        Dim has1 As Boolean = False, has2 As Boolean = False
        Dim b1 As ViewSheetBoundingBox, b2 As ViewSheetBoundingBox
        has1 = TryGetObjectBBoxSheet(dv, obj1, b1)
        has2 = TryGetObjectBBoxSheet(dv, obj2, b2)
        If Not has1 OrElse Not has2 Then Return False

        outBox.MinX = Math.Min(b1.MinX, b2.MinX)
        outBox.MaxX = Math.Max(b1.MaxX, b2.MaxX)
        outBox.MinY = Math.Min(b1.MinY, b2.MinY)
        outBox.MaxY = Math.Max(b1.MaxY, b2.MaxY)
        ' La unión de dos líneas horizontales al mismo Y (o dos verticales al mismo X) tiene altura 0 o ancho 0;
        ' sigue siendo un bbox local válido para colocación de cota horizontal/vertical.
        Const Eps As Double = 1.0E-12
        Dim ok As Boolean = (outBox.Width > Eps OrElse outBox.Height > Eps) AndAlso Not (outBox.Width <= Eps AndAlso outBox.Height <= Eps)
        log?.Info("[DIM][PLACE][LOCAL][DEBUG] feature=" & featureName)
        log?.Info("[DIM][PLACE][LOCAL][DEBUG] entity1=" & GetObjectTypeName(obj1))
        log?.Info("[DIM][PLACE][LOCAL][DEBUG] entity2=" & GetObjectTypeName(obj2))
        log?.Info(String.Format(CultureInfo.InvariantCulture, "[DIM][PLACE][LOCAL][DEBUG] p1=({0:0.######}) p2=({1:0.######})", pick1, pick2))
        log?.Info(String.Format(CultureInfo.InvariantCulture,
                                "[DIM][PLACE][LOCAL][DEBUG] bbox local calculado=({0:0.######},{1:0.######})-({2:0.######},{3:0.######})",
                                outBox.MinX, outBox.MinY, outBox.MaxX, outBox.MaxY))
        If ok AndAlso (outBox.Width <= Eps OrElse outBox.Height <= Eps) Then
            log?.Info("[DIM][PLACE][LOCAL][DEBUG] bbox degenerado 1D aceptado (ancho o alto nulo en unión)")
        End If
        Return ok
    End Function

    Private Shared Function TryGetObjectBBoxSheet(dv As DrawingView, obj As Object, ByRef b As ViewSheetBoundingBox) As Boolean
        b = Nothing
        Try
            If TypeOf obj Is DVLine2d Then
                Dim ln As DVLine2d = CType(obj, DVLine2d)
                Dim vx1 As Double = 0, vy1 As Double = 0, vx2 As Double = 0, vy2 As Double = 0
                ln.GetStartPoint(vx1, vy1)
                ln.GetEndPoint(vx2, vy2)
                Dim sx1 As Double = 0, sy1 As Double = 0, sx2 As Double = 0, sy2 As Double = 0
                dv.ViewToSheet(vx1, vy1, sx1, sy1)
                dv.ViewToSheet(vx2, vy2, sx2, sy2)
                b.MinX = Math.Min(sx1, sx2) : b.MaxX = Math.Max(sx1, sx2)
                b.MinY = Math.Min(sy1, sy2) : b.MaxY = Math.Max(sy1, sy2)
                Return True
            End If

            If TypeOf obj Is DVArc2d OrElse TypeOf obj Is DVCircle2d Then
                Dim x1 As Double = 0, y1 As Double = 0, x2 As Double = 0, y2 As Double = 0
                If TypeOf obj Is DVArc2d Then
                    CType(obj, DVArc2d).Range(x1, y1, x2, y2)
                Else
                    CType(obj, DVCircle2d).Range(x1, y1, x2, y2)
                End If
                Dim sx1 As Double = 0, sy1 As Double = 0, sx2 As Double = 0, sy2 As Double = 0, sx3 As Double = 0, sy3 As Double = 0, sx4 As Double = 0, sy4 As Double = 0
                dv.ViewToSheet(x1, y1, sx1, sy1)
                dv.ViewToSheet(x2, y1, sx2, sy2)
                dv.ViewToSheet(x1, y2, sx3, sy3)
                dv.ViewToSheet(x2, y2, sx4, sy4)
                b.MinX = Math.Min(Math.Min(sx1, sx2), Math.Min(sx3, sx4))
                b.MaxX = Math.Max(Math.Max(sx1, sx2), Math.Max(sx3, sx4))
                b.MinY = Math.Min(Math.Min(sy1, sy2), Math.Min(sy3, sy4))
                b.MaxY = Math.Max(Math.Max(sy1, sy2), Math.Max(sy3, sy4))
                Return True
            End If

            If TypeOf obj Is DVEllipse2d Then
                Dim x1 As Double = 0, y1 As Double = 0, x2 As Double = 0, y2 As Double = 0
                CType(obj, DVEllipse2d).Range(x1, y1, x2, y2)
                Dim sx1 As Double = 0, sy1 As Double = 0, sx2 As Double = 0, sy2 As Double = 0, sx3 As Double = 0, sy3 As Double = 0, sx4 As Double = 0, sy4 As Double = 0
                dv.ViewToSheet(x1, y1, sx1, sy1)
                dv.ViewToSheet(x2, y1, sx2, sy2)
                dv.ViewToSheet(x1, y2, sx3, sy3)
                dv.ViewToSheet(x2, y2, sx4, sy4)
                b.MinX = Math.Min(Math.Min(sx1, sx2), Math.Min(sx3, sx4))
                b.MaxX = Math.Max(Math.Max(sx1, sx2), Math.Max(sx3, sx4))
                b.MinY = Math.Min(Math.Min(sy1, sy2), Math.Min(sy3, sy4))
                b.MaxY = Math.Max(Math.Max(sy1, sy2), Math.Max(sy3, sy4))
                Return True
            End If

            ' Fallback robusto para casos interop no tipados (p.ej. __ComObject):
            Dim rx1 As Double = 0, ry1 As Double = 0, rx2 As Double = 0, ry2 As Double = 0
            If TryGetRangeDynamic(obj, rx1, ry1, rx2, ry2) Then
                Dim sx1 As Double = 0, sy1 As Double = 0, sx2 As Double = 0, sy2 As Double = 0, sx3 As Double = 0, sy3 As Double = 0, sx4 As Double = 0, sy4 As Double = 0
                dv.ViewToSheet(rx1, ry1, sx1, sy1)
                dv.ViewToSheet(rx2, ry1, sx2, sy2)
                dv.ViewToSheet(rx1, ry2, sx3, sy3)
                dv.ViewToSheet(rx2, ry2, sx4, sy4)
                b.MinX = Math.Min(Math.Min(sx1, sx2), Math.Min(sx3, sx4))
                b.MaxX = Math.Max(Math.Max(sx1, sx2), Math.Max(sx3, sx4))
                b.MinY = Math.Min(Math.Min(sy1, sy2), Math.Min(sy3, sy4))
                b.MaxY = Math.Max(Math.Max(sy1, sy2), Math.Max(sy3, sy4))
                Return True
            End If
        Catch
            Return False
        End Try
        Return False
    End Function

    Private Shared Function TryGetRangeDynamic(obj As Object, ByRef x1 As Double, ByRef y1 As Double, ByRef x2 As Double, ByRef y2 As Double) As Boolean
        x1 = 0 : y1 = 0 : x2 = 0 : y2 = 0
        Try
            CallByName(obj, "Range", CallType.Method, x1, y1, x2, y2)
            Return True
        Catch
            Return False
        End Try
    End Function

End Class
