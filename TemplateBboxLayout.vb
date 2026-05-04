Option Strict Off

Imports System.Globalization
Imports SolidEdgeDraft
Imports SolidEdgeFramework

''' <summary>
''' Reparto de vistas por slots (bbox) en coordenadas de hoja de referencia, mapeadas al área útil
''' de cada plantilla (A4/A3/A2/A1, horizontal o vertical). Las vistas se pueden centrar en cada slot.
''' </summary>
Public Module TemplateBboxLayout

    ''' <summary>False = comportamiento anterior (solo FixedComposition clásico + escala×2).</summary>
    Public UseSlotBasedLayout As Boolean = True

    Private Const GAP_H As Double = 0.012
    Private Const GAP_V As Double = 0.012
    Private Const ISO_FACTOR As Double = 0.45

    ' --- Área libre de referencia (AABB del polígono indicado, m, sistema SE: Y hacia arriba) ---
    Private Const RefFreeMinX As Double = 0.01
    Private Const RefFreeMaxX As Double = 0.4098
    Private Const RefFreeMinY As Double = 0.01
    Private Const RefFreeMaxY As Double = 0.2762

    ' Slots en coordenadas REF (m)
    Private Const MainMinX As Double = 0.0264
    Private Const MainMinY As Double = 0.1723
    Private Const MainMaxX As Double = 0.2334
    Private Const MainMaxY As Double = 0.2719

    Private Const RightMinX As Double = 0.2454
    Private Const RightMinY As Double = 0.1717
    Private Const RightMaxX As Double = 0.3965
    Private Const RightMaxY As Double = 0.2719

    Private Const BelowMinX As Double = 0.0264
    Private Const BelowMinY As Double = 0.0883
    Private Const BelowMaxX As Double = 0.2334
    Private Const BelowMaxY As Double = 0.1723

    Private Const FlatMinX As Double = 0.0156
    Private Const FlatMinY As Double = 0.01
    Private Const FlatMaxX As Double = 0.2198
    Private Const FlatMaxY As Double = 0.0869

    Private Const IsoMinX As Double = 0.3163
    Private Const IsoMinY As Double = 0.04
    Private Const IsoMaxX As Double = 0.4098
    Private Const IsoMaxY As Double = 0.1013

    ' Sobre la REF, envolvente del bloque principal+derecha (X) y principal+debajo (Y)
    Private Const BlockRowMinX As Double = 0.0264
    Private Const BlockRowMaxX As Double = 0.3965
    Private Const BlockColMinY As Double = 0.0883
    Private Const BlockColMaxY As Double = 0.2719

    ''' <summary>Escalas de dibujo 1:1 … 1:200 (factor modelo→hoja, orden descendente).</summary>
    Public ReadOnly UneDrawingScaleFactors As Double() = {
        1.0R, 0.5R, 0.2R, 0.1R, 0.05R, 0.04R, 1.0R / 30.0R, 0.025R, 0.02R, 1.0R / 75.0R, 0.01R,
        1.0R / 150.0R, 1.0R / 200.0R
    }

    Public Function SnapUneScale(maxScale As Double) As Double
        If maxScale <= 0 Then Return UneDrawingScaleFactors(UneDrawingScaleFactors.Length - 1)
        For Each s In UneDrawingScaleFactors
            If s <= maxScale + 1.0E-9 Then Return s
        Next
        Return UneDrawingScaleFactors(UneDrawingScaleFactors.Length - 1)
    End Function

    ''' <summary>Tamaño modelo 1:1 (m) usado como envolvente conservadora para la isométrica.</summary>
    Public Function IsoModelEnvelopeAtScale1(mainW1 As Double, mainH1 As Double) As Double
        Return Math.Max(mainW1, mainH1) * 1.15R
    End Function

    ''' <summary>Escala UNE para la vista isométrica (independiente de la de las 3 principales; limitada por el slot ISO).</summary>
    Public Function ComputeIsoUneScale(principalScale As Double, mainW1 As Double, mainH1 As Double, u As LayoutEngine.UsableArea) As Double
        If u Is Nothing OrElse mainW1 <= 0 OrElse mainH1 <= 0 Then
            Return principalScale * ISO_FACTOR
        End If
        Dim isoR = MapRefRectToUsable(IsoMinX, IsoMinY, IsoMaxX, IsoMaxY, u)
        Dim d1 = IsoModelEnvelopeAtScale1(mainW1, mainH1)
        If d1 <= 1.0E-12 Then Return principalScale * ISO_FACTOR
        Dim sMax = Math.Min(isoR.Width / d1, isoR.Height / d1)
        Return SnapUneScale(sMax)
    End Function

    ''' <summary>Tras crear la flat a escala actual, reduce a escala UNE si no cabe en el slot chapa (mapeado REF→útil).</summary>
    Public Function ResolveFlatScaleForBboxSlot(
        dvs As DrawingViews,
        flat As DrawingView,
        templatePath As String,
        currentScale As Double,
        log As Action(Of String)) As Double

        If flat Is Nothing OrElse dvs Is Nothing OrElse Not UseSlotBasedLayout OrElse String.IsNullOrWhiteSpace(templatePath) Then
            Return currentScale
        End If
        If currentScale <= 0 Then Return currentScale
        Dim u = LayoutEngine.GetUsableAreaForTemplate(templatePath)
        If u Is Nothing Then Return currentScale

        Dim xmin As Double, ymin As Double, xmax As Double, ymax As Double
        If Not CojonudoBestFit_Bueno.TryGetViewRangePublic(flat, xmin, ymin, xmax, ymax) Then Return currentScale
        Dim fw = Math.Abs(xmax - xmin)
        Dim fh = Math.Abs(ymax - ymin)
        If fw <= 1.0E-12 OrElse fh <= 1.0E-12 Then Return currentScale

        Dim flatR = MapRefRectToUsable(FlatMinX, FlatMinY, FlatMaxX, FlatMaxY, u)
        Const margin As Double = 1.02R
        Dim rx = flatR.Width / (fw * margin)
        Dim ry = flatR.Height / (fh * margin)
        Dim ratio = Math.Min(rx, ry)
        If ratio >= 1.0R Then
            log?.Invoke("[BBOX-LAYOUT] Flat ya cabe en slot; escala=" & currentScale.ToString("G9", CultureInfo.InvariantCulture))
            Return currentScale
        End If
        Dim raw = currentScale * ratio
        Dim snapped = SnapUneScale(raw)
        log?.Invoke(String.Format(CultureInfo.InvariantCulture, "[BBOX-LAYOUT] Flat ajuste slot ratio={0:0.####} raw={1:0.######} UNE={2:0.######}",
                                  ratio, raw, snapped))
        Return snapped
    End Function

    Private Function RefSpanX() As Double
        Return RefFreeMaxX - RefFreeMinX
    End Function

    Private Function RefSpanY() As Double
        Return RefFreeMaxY - RefFreeMinY
    End Function

    ''' <summary>Mapea X en ref (m) a X en área útil de la plantilla actual.</summary>
    Public Function MapRefXToUsable(xRef As Double, u As LayoutEngine.UsableArea) As Double
        If u Is Nothing Then Return xRef
        Dim rx = RefSpanX()
        If rx <= 1.0E-12 Then Return u.MinX
        Return u.MinX + (xRef - RefFreeMinX) / rx * u.Width
    End Function

    Public Function MapRefYToUsable(yRef As Double, u As LayoutEngine.UsableArea) As Double
        If u Is Nothing Then Return yRef
        Dim ry = RefSpanY()
        If ry <= 1.0E-12 Then Return u.MinY
        Return u.MinY + (yRef - RefFreeMinY) / ry * u.Height
    End Function

    Public Function MapRefRectToUsable(minXr As Double, minYr As Double, maxXr As Double, maxYr As Double, u As LayoutEngine.UsableArea) As LayoutEngine.UsableArea
        Dim o As New LayoutEngine.UsableArea With {
            .MinX = MapRefXToUsable(minXr, u),
            .MaxX = MapRefXToUsable(maxXr, u),
            .MinY = MapRefYToUsable(minYr, u),
            .MaxY = MapRefYToUsable(maxYr, u)
        }
        If o.MaxX < o.MinX Then
            Dim t = o.MinX : o.MinX = o.MaxX : o.MaxX = t
        End If
        If o.MaxY < o.MinY Then
            Dim t = o.MinY : o.MinY = o.MaxY : o.MaxY = t
        End If
        Return o
    End Function

    ''' <summary>Escala UNE máxima para que el bloque 3 vistas quepa en la envolvente mapeada.</summary>
    Public Function ComputeUneScaleForThreeViewBlock(plan As FixedCompositionLayout.FinalLayoutPlan, u As LayoutEngine.UsableArea) As Double
        If plan Is Nothing OrElse u Is Nothing Then Return 0.05
        Dim rowW = MapRefXToUsable(BlockRowMaxX, u) - MapRefXToUsable(BlockRowMinX, u)
        Dim colH = MapRefYToUsable(BlockColMaxY, u) - MapRefYToUsable(BlockColMinY, u)
        Dim bw1 = plan.MainWidthAt1 + GAP_H + plan.RightWidthAt1
        Dim bh1 = plan.MainHeightAt1 + GAP_V + plan.BottomHeightAt1
        If bw1 <= 0 OrElse bh1 <= 0 Then Return 0.05
        Dim sMax = Math.Min(rowW / bw1, colH / bh1)
        Return SnapUneScale(sMax)
    End Function

    ''' <summary>Recalcula <see cref="FixedCompositionLayout.LayoutZones"/> con top-left coherentes con InsertStandard3 (misma base Y fila principal).</summary>
    Public Function ComputeZonesFromBboxSlots(plan As FixedCompositionLayout.FinalLayoutPlan, u As LayoutEngine.UsableArea) As FixedCompositionLayout.LayoutZones
        Dim z As New FixedCompositionLayout.LayoutZones
        If plan Is Nothing OrElse u Is Nothing Then Return z

        Dim mw = plan.MainWidthAt1 * plan.Scale
        Dim mh = plan.MainHeightAt1 * plan.Scale
        Dim rw = plan.RightWidthAt1 * plan.Scale
        Dim rh = plan.RightHeightAt1 * plan.Scale
        Dim bw = plan.BottomWidthAt1 * plan.Scale
        Dim bh = plan.BottomHeightAt1 * plan.Scale

        Dim mainR = MapRefRectToUsable(MainMinX, MainMinY, MainMaxX, MainMaxY, u)
        Dim rightR = MapRefRectToUsable(RightMinX, RightMinY, RightMaxX, RightMaxY, u)
        Dim belowR = MapRefRectToUsable(BelowMinX, BelowMinY, BelowMaxX, BelowMaxY, u)

        Dim cxM = (mainR.MinX + mainR.MaxX) / 2.0R
        Dim cyM = (mainR.MinY + mainR.MaxY) / 2.0R
        Dim cxR = (rightR.MinX + rightR.MaxX) / 2.0R
        Dim cxB = (belowR.MinX + belowR.MaxX) / 2.0R
        Dim cyB = (belowR.MinY + belowR.MaxY) / 2.0R

        ' Fila principal: misma base Y (InsertStandard3 usa el mismo y0 para base y derecha)
        Dim baseY As Double = cyM - mh / 2.0R

        z.MainTopLeftX = cxM - mw / 2.0R
        z.MainTopLeftY = baseY + mh

        z.RightTopLeftX = cxR - rw / 2.0R
        z.RightTopLeftY = baseY + rh

        z.GapRight = z.RightTopLeftX - z.MainTopLeftX
        z.GapBelow = mh + GAP_V

        ' Debajo: centrar en slot; alinear X con la vista principal
        z.BottomTopLeftX = cxM - bw / 2.0R
        z.BottomTopLeftY = cyB + bh / 2.0R

        ' ISO / Flat: top-left aproximado para primera colocación (luego se centra fino)
        If plan.IncludeIso Then
            Dim isoR = MapRefRectToUsable(IsoMinX, IsoMinY, IsoMaxX, IsoMaxY, u)
            Dim isoSc = ComputeIsoUneScale(plan.Scale, plan.MainWidthAt1, plan.MainHeightAt1, u)
            Dim isoDisp = IsoModelEnvelopeAtScale1(plan.MainWidthAt1, plan.MainHeightAt1) * isoSc
            Dim cxI = (isoR.MinX + isoR.MaxX) / 2.0R
            Dim cyI = (isoR.MinY + isoR.MaxY) / 2.0R
            z.IsoTopLeftX = cxI - isoDisp / 2.0R
            z.IsoTopLeftY = cyI + isoDisp / 2.0R
        End If

        If plan.IncludeFlat Then
            Dim flatR = MapRefRectToUsable(FlatMinX, FlatMinY, FlatMaxX, FlatMaxY, u)
            Dim estW = 0.12R
            Dim estH = 0.08R
            Dim cxF = (flatR.MinX + flatR.MaxX) / 2.0R
            Dim cyF = (flatR.MinY + flatR.MaxY) / 2.0R
            z.FlatTopLeftX = cxF - estW / 2.0R
            z.FlatTopLeftY = cyF + estH / 2.0R
        End If

        Return z
    End Function

    Public Function IsoFitsInSlot(plan As FixedCompositionLayout.FinalLayoutPlan, u As LayoutEngine.UsableArea) As Boolean
        If plan Is Nothing OrElse u Is Nothing Then Return False
        Dim isoR = MapRefRectToUsable(IsoMinX, IsoMinY, IsoMaxX, IsoMaxY, u)
        Dim isoSc = ComputeIsoUneScale(plan.Scale, plan.MainWidthAt1, plan.MainHeightAt1, u)
        Dim isoDisp = IsoModelEnvelopeAtScale1(plan.MainWidthAt1, plan.MainHeightAt1) * isoSc
        Return isoDisp <= isoR.Width + 1.0E-6 AndAlso isoDisp <= isoR.Height + 1.0E-6
    End Function

    ''' <summary>Centra la vista en el slot mapeado (post-Update, Range fiable).</summary>
    Public Sub MoveViewCenterToSheetPoint(app As SolidEdgeFramework.Application, dv As DrawingView, cx As Double, cy As Double, log As Action(Of String), tag As String)
        If dv Is Nothing OrElse app Is Nothing Then Return
        Dim xmin As Double, ymin As Double, xmax As Double, ymax As Double
        If Not CojonudoBestFit_Bueno.TryGetViewRangePublic(dv, xmin, ymin, xmax, ymax) Then
            log?.Invoke($"[BBOX-LAYOUT] MoveViewCenter[{tag}]: sin Range")
            Return
        End If
        Dim w = xmax - xmin
        Dim h = ymax - ymin
        If w <= 0 OrElse h <= 0 Then Return
        Dim targetLeft = cx - w / 2.0R
        Dim targetTop = cy + h / 2.0R
        Dim ox As Double = 0, oy As Double = 0
        Try
            dv.GetOrigin(ox, oy)
        Catch
        End Try
        Dim dx = targetLeft - xmin
        Dim dy = targetTop - ymax
        Try
            dv.SetOrigin(ox + dx, oy + dy)
            log?.Invoke(String.Format(CultureInfo.InvariantCulture, "[BBOX-LAYOUT] {0} center=({1:0.######},{2:0.######})", tag, cx, cy))
        Catch ex As Exception
            log?.Invoke("[BBOX-LAYOUT] SetOrigin " & tag & ": " & ex.Message)
        End Try
        Try
            app.DoIdle()
        Catch
        End Try
    End Sub

    Public Sub ApplySlotCentersAfterInsert(
        app As SolidEdgeFramework.Application,
        templatePath As String,
        plan As FixedCompositionLayout.FinalLayoutPlan,
        vMain As DrawingView,
        vRight As DrawingView,
        vBelow As DrawingView,
        vIso As DrawingView,
        vFlat As DrawingView,
        flatInserted As Boolean,
        log As Action(Of String))

        If Not UseSlotBasedLayout OrElse plan Is Nothing Then Return
        Dim u = LayoutEngine.GetUsableAreaForTemplate(templatePath)
        If u Is Nothing Then Return

        Dim mainR = MapRefRectToUsable(MainMinX, MainMinY, MainMaxX, MainMaxY, u)
        Dim rightR = MapRefRectToUsable(RightMinX, RightMinY, RightMaxX, RightMaxY, u)
        Dim belowR = MapRefRectToUsable(BelowMinX, BelowMinY, BelowMaxX, BelowMaxY, u)

        If vMain IsNot Nothing Then
            MoveViewCenterToSheetPoint(app, vMain, (mainR.MinX + mainR.MaxX) / 2.0R, (mainR.MinY + mainR.MaxY) / 2.0R, log, "Main")
        End If
        If vRight IsNot Nothing Then
            MoveViewCenterToSheetPoint(app, vRight, (rightR.MinX + rightR.MaxX) / 2.0R, (rightR.MinY + rightR.MaxY) / 2.0R, log, "Right")
        End If
        If vBelow IsNot Nothing Then
            MoveViewCenterToSheetPoint(app, vBelow, (belowR.MinX + belowR.MaxX) / 2.0R, (belowR.MinY + belowR.MaxY) / 2.0R, log, "Below")
        End If

        If plan.IncludeIso AndAlso vIso IsNot Nothing Then
            Dim isoR = MapRefRectToUsable(IsoMinX, IsoMinY, IsoMaxX, IsoMaxY, u)
            MoveViewCenterToSheetPoint(app, vIso, (isoR.MinX + isoR.MaxX) / 2.0R, (isoR.MinY + isoR.MaxY) / 2.0R, log, "Iso")
        End If

        If flatInserted AndAlso plan.IncludeFlat AndAlso vFlat IsNot Nothing Then
            Dim flatR = MapRefRectToUsable(FlatMinX, FlatMinY, FlatMaxX, FlatMaxY, u)
            MoveViewCenterToSheetPoint(app, vFlat, (flatR.MinX + flatR.MaxX) / 2.0R, (flatR.MinY + flatR.MaxY) / 2.0R, log, "Flat")
        End If
    End Sub
End Module
