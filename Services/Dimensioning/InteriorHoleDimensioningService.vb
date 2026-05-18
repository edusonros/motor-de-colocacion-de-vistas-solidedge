Option Strict Off

Imports System.Collections.Generic
Imports System.Globalization
Imports System.Linq
Imports SolidEdgeDraft
Imports SolidEdgeFrameworkSupport

''' <summary>
''' Agujeros interiores: <c>AddCircularDiameter</c> en cada círculo; posición con línea auxiliar de hoja + <c>AddLength</c>.
''' </summary>
Friend NotInheritable Class InteriorHoleDimensioningService

    Private Const MinHoleRadiusM As Double = 0.0008R
    Private Const InteriorMarginFraction As Double = 0.04R
    Private Const MaxHoleRadiusFraction As Double = 0.42R
    Private Const ValueTolMm As Double = 0.05R
    Private Const ValueTolM As Double = 0.002R
    Private Const MinCenterSepM As Double = 0.0005R
    Private Const ClusterTolFraction As Double = 0.08R

    Private NotInheritable Class HoleInfo
        Public Circle As Object
        Public CxView As Double
        Public CyView As Double
        Public Radius As Double
        Public CxSheet As Double
        Public CySheet As Double
    End Class

    Private Sub New()
    End Sub

    Public Shared Function CreateForViews(
        draft As DraftDocument,
        sheet As Sheet,
        workList As IList(Of DrawingViewGeometryInfo),
        dims As Dimensions,
        styleObj As Object,
        norm As DimensioningNormConfig,
        log As DimensionLogger) As Integer

        If norm IsNot Nothing AndAlso Not norm.EnableInteriorHoleCenterDimensions Then Return 0
        If draft Is Nothing OrElse sheet Is Nothing OrElse dims Is Nothing OrElse workList Is Nothing Then Return 0

        Dim created As Integer = 0
        For Each info As DrawingViewGeometryInfo In workList
            If info Is Nothing OrElse info.View Is Nothing OrElse info.CountCircles <= 0 Then Continue For
            created += CreateForView(draft, sheet, info, dims, styleObj, norm, log)
        Next
        Return LogHoleSummary(created, log)
    End Function

    Friend Shared Function CreateForLegacyViews(
        draft As DraftDocument,
        sheet As Sheet,
        views As IList(Of DrawingView),
        dims As Dimensions,
        styleObj As Object,
        norm As DimensioningNormConfig,
        appLog As Extraer_dft_dxf_flatdxf.LegacyV02Dimensioning.DimensionLogger) As Integer

        If norm IsNot Nothing AndAlso Not norm.EnableInteriorHoleCenterDimensions Then Return 0
        If draft Is Nothing OrElse sheet Is Nothing OrElse dims Is Nothing OrElse views Is Nothing Then Return 0
        Dim log As DimensionLogger
        If appLog IsNot Nothing Then
            Dim fwd As New Logger(Sub(m) appLog.LogLine(m))
            log = New DimensionLogger(fwd)
        Else
            log = New DimensionLogger(Nothing)
        End If

        Dim created As Integer = 0
        Dim idx As Integer = 0
        For Each dv As DrawingView In views
            idx += 1
            If dv Is Nothing Then Continue For
            Dim info As DrawingViewGeometryInfo = DrawingViewGeometryReader.Read(dv, idx, log)
            If info Is Nothing OrElse info.CountCircles <= 0 Then Continue For
            created += CreateForView(draft, sheet, info, dims, styleObj, norm, log)
        Next
        Return LogHoleSummary(created, log)
    End Function

    Private Shared Function LogHoleSummary(created As Integer, log As DimensionLogger) As Integer
        If created > 0 Then
            log?.LogLine("[DIM][HOLE][SUMMARY] created=" & created.ToString(CultureInfo.InvariantCulture))
        End If
        Return created
    End Function

    Private Shared Function CreateForView(
        draft As DraftDocument,
        sheet As Sheet,
        info As DrawingViewGeometryInfo,
        dims As Dimensions,
        styleObj As Object,
        norm As DimensioningNormConfig,
        log As DimensionLogger) As Integer

        Dim holes As List(Of HoleInfo) = CollectInteriorHoles(info, log)
        If holes Is Nothing OrElse holes.Count = 0 Then Return 0

        Dim ex As ExtremeDvLinesResult = info.Extreme
        If ex Is Nothing Then
            log?.LogLine("[DIM][HOLE][SKIP] view=" & info.ViewIndex.ToString(CultureInfo.InvariantCulture) & " reason=no_extreme_lines")
            Return 0
        End If

        Dim effStyle As Object = DrawingViewDimensionCreator.ResolveForcedStyleObject(draft, sheet, styleObj, log)
        Dim viewIdx As Integer = info.ViewIndex
        Dim usedNominalSide As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Dim created As Integer = 0

        log?.LogLine("[DIM][HOLE][VIEW] idx=" & viewIdx.ToString(CultureInfo.InvariantCulture) &
                     " name=" & info.ViewName &
                     " interior_circles=" & holes.Count.ToString(CultureInfo.InvariantCulture))

        created += CreateCircularDiametersForHoles(dims, holes, viewIdx, effStyle, norm, log)

        Dim refHole As HoleInfo = SelectReferenceHole(holes)
        If refHole IsNot Nothing Then
            created += CreateReferenceHoleToSides(sheet, dims, refHole, ex, viewIdx, usedNominalSide, effStyle, log)
        End If

        If holes.Count > 1 Then
            created += CreateCenterSpacingInRowsAndColumns(sheet, dims, holes, viewIdx, usedNominalSide, effStyle, log)
        End If

        Return created
    End Function

    Private Shared Function CreateCircularDiametersForHoles(
        dims As Dimensions,
        holes As List(Of HoleInfo),
        viewIdx As Integer,
        styleObj As Object,
        norm As DimensioningNormConfig,
        log As DimensionLogger) As Integer

        Dim created As Integer = 0
        For Each h As HoleInfo In holes
            If h Is Nothing OrElse h.Circle Is Nothing Then Continue For
            Dim rMm As Double = h.Radius * 1000.0R
            If norm IsNot Nothing AndAlso rMm < norm.MinRadiusToDimensionMm Then
                log?.LogLine("[DIM][HOLE][DIA][SKIP] view=" & viewIdx.ToString(CultureInfo.InvariantCulture) &
                             " r_mm=" & rMm.ToString("0.###", CultureInfo.InvariantCulture) & " reason=min_radius")
                Continue For
            End If

            Dim dObj As Dimension = DrawingViewDimensionCreator.TryCreateCircularDiameterOnReference(dims, h.Circle, log)
            If dObj Is Nothing Then
                log?.LogLine("[DIM][HOLE][DIA][FAIL] view=" & viewIdx.ToString(CultureInfo.InvariantCulture) &
                             " r_mm=" & rMm.ToString("0.###", CultureInfo.InvariantCulture))
                Continue For
            End If

            DrawingViewDimensionCreator.TryApplyStyleToDimension(dObj, styleObj, log)
            TryApplyVisibleTrackDistance(dObj, True, log)
            created += 1
            log?.LogLine("[DIM][HOLE][DIA][OK] view=" & viewIdx.ToString(CultureInfo.InvariantCulture) &
                         " method=AddCircularDiameter dia_mm=" & (rMm * 2.0R).ToString("0.###", CultureInfo.InvariantCulture))
        Next
        Return created
    End Function

    Private Shared Function SelectReferenceHole(holes As List(Of HoleInfo)) As HoleInfo
        Return holes.OrderBy(Function(h) h.CySheet).ThenBy(Function(h) h.CxSheet).FirstOrDefault()
    End Function

    Private Shared Function CreateReferenceHoleToSides(
        sheet As Sheet,
        dims As Dimensions,
        hole As HoleInfo,
        ex As ExtremeDvLinesResult,
        viewIdx As Integer,
        usedNominalSide As HashSet(Of String),
        styleObj As Object,
        log As DimensionLogger) As Integer

        Dim created As Integer = 0
        Dim refLeft As DvLineSheetInfo = ex.LeftVertical
        Dim refBottom As DvLineSheetInfo = ex.BottomHorizontal

        If refLeft IsNot Nothing Then
            Dim distH As Double = Math.Abs(hole.CxSheet - refLeft.MidX)
            If TryCreateAuxLengthBetweenPoints(
                sheet, dims, refLeft.MidX, hole.CySheet, hole.CxSheet, hole.CySheet,
                distH, True, DimensionSide.Bottom, viewIdx, usedNominalSide, styleObj, log, "center_to_left") Then
                created += 1
            End If
        End If

        If refBottom IsNot Nothing Then
            Dim distV As Double = Math.Abs(hole.CySheet - refBottom.MidY)
            If TryCreateAuxLengthBetweenPoints(
                sheet, dims, hole.CxSheet, refBottom.MidY, hole.CxSheet, hole.CySheet,
                distV, False, DimensionSide.Right, viewIdx, usedNominalSide, styleObj, log, "center_to_bottom") Then
                created += 1
            End If
        End If

        Return created
    End Function

    Private Shared Function CreateCenterSpacingInRowsAndColumns(
        sheet As Sheet,
        dims As Dimensions,
        holes As List(Of HoleInfo),
        viewIdx As Integer,
        usedNominalSide As HashSet(Of String),
        styleObj As Object,
        log As DimensionLogger) As Integer

        Dim created As Integer = 0
        Dim minX As Double = holes.Min(Function(h) h.CxSheet)
        Dim maxX As Double = holes.Max(Function(h) h.CxSheet)
        Dim minY As Double = holes.Min(Function(h) h.CySheet)
        Dim maxY As Double = holes.Max(Function(h) h.CySheet)
        Dim clusterTol As Double = Math.Max(Math.Max((maxX - minX) * ClusterTolFraction, (maxY - minY) * ClusterTolFraction), 0.004R)

        Dim rows = ClusterByCoordinate(holes, Function(h) h.CySheet, clusterTol)
        For Each row As List(Of HoleInfo) In rows
            If row.Count < 2 Then Continue For
            Dim ordered = row.OrderBy(Function(h) h.CxSheet).ToList()
            For i As Integer = 0 To ordered.Count - 2
                Dim a = ordered(i)
                Dim b = ordered(i + 1)
                Dim dist As Double = Math.Abs(b.CxSheet - a.CxSheet)
                If dist < MinCenterSepM Then Continue For
                If TryCreateAuxLengthBetweenPoints(
                    sheet, dims, a.CxSheet, a.CySheet, b.CxSheet, b.CySheet,
                    dist, True, DimensionSide.Top, viewIdx, usedNominalSide, styleObj, log, "center_spacing_h") Then
                    created += 1
                End If
            Next
        Next

        Dim cols = ClusterByCoordinate(holes, Function(h) h.CxSheet, clusterTol)
        For Each col As List(Of HoleInfo) In cols
            If col.Count < 2 Then Continue For
            Dim ordered = col.OrderBy(Function(h) h.CySheet).ToList()
            For i As Integer = 0 To ordered.Count - 2
                Dim a = ordered(i)
                Dim b = ordered(i + 1)
                Dim dist As Double = Math.Abs(b.CySheet - a.CySheet)
                If dist < MinCenterSepM Then Continue For
                If TryCreateAuxLengthBetweenPoints(
                    sheet, dims, a.CxSheet, a.CySheet, b.CxSheet, b.CySheet,
                    dist, False, DimensionSide.Right, viewIdx, usedNominalSide, styleObj, log, "center_spacing_v") Then
                    created += 1
                End If
            Next
        Next

        Return created
    End Function

    Private Shared Function ClusterByCoordinate(
        holes As List(Of HoleInfo),
        coord As Func(Of HoleInfo, Double),
        tol As Double) As List(Of List(Of HoleInfo))

        Dim sorted = holes.OrderBy(coord).ToList()
        Dim groups As New List(Of List(Of HoleInfo))()
        Dim current As List(Of HoleInfo) = Nothing
        Dim anchor As Double = 0
        For Each h As HoleInfo In sorted
            Dim c As Double = coord(h)
            If current Is Nothing OrElse Math.Abs(c - anchor) > tol Then
                current = New List(Of HoleInfo)()
                groups.Add(current)
                anchor = c
            End If
            current.Add(h)
        Next
        Return groups
    End Function

    Private Shared Function CollectInteriorHoles(info As DrawingViewGeometryInfo, log As DimensionLogger) As List(Of HoleInfo)
        Dim out As New List(Of HoleInfo)()
        If info Is Nothing OrElse info.View Is Nothing Then Return out

        Dim col As Object = Nothing
        Try
            col = CallByName(info.View, "DVCircles2d", CallType.Get)
        Catch
            Return out
        End Try
        Dim n As Integer = SafeCount(col)
        If n <= 0 Then Return out

        Dim box As ViewSheetBoundingBox = info.Box
        Dim marginX As Double = Math.Max(box.Width * InteriorMarginFraction, 0.001R)
        Dim marginY As Double = Math.Max(box.Height * InteriorMarginFraction, 0.001R)
        Dim innerMinX As Double = box.MinX + marginX
        Dim innerMaxX As Double = box.MaxX - marginX
        Dim innerMinY As Double = box.MinY + marginY
        Dim innerMaxY As Double = box.MaxY - marginY
        Dim maxR As Double = Math.Min(box.Width, box.Height) * MaxHoleRadiusFraction

        For i As Integer = 1 To n
            Dim c As Object = Nothing
            Try
                c = CallByName(col, "Item", CallType.Method, i)
            Catch
                c = Nothing
            End Try
            If c Is Nothing Then Continue For

            Dim cxV As Double = 0, cyV As Double = 0, r As Double = 0
            If Not TryCircleCenterRadiusView(c, cxV, cyV, r) Then Continue For
            If r < MinHoleRadiusM OrElse r > maxR Then Continue For

            Dim sx As Double = 0, sy As Double = 0
            Try
                info.View.ViewToSheet(cxV, cyV, sx, sy)
            Catch
                Continue For
            End Try

            If sx < innerMinX OrElse sx > innerMaxX OrElse sy < innerMinY OrElse sy > innerMaxY Then Continue For

            out.Add(New HoleInfo With {
                .Circle = c,
                .CxView = cxV,
                .CyView = cyV,
                .Radius = r,
                .CxSheet = sx,
                .CySheet = sy
            })
        Next

        If out.Count > 0 Then
            log?.LogLine("[DIM][HOLE][COLLECT] view=" & info.ViewIndex.ToString(CultureInfo.InvariantCulture) &
                         " interior=" & out.Count.ToString(CultureInfo.InvariantCulture) &
                         " total_circles=" & n.ToString(CultureInfo.InvariantCulture))
        End If
        Return out
    End Function

    ''' <summary>Línea auxiliar oculta en hoja + <c>AddLength</c> (equivalente fiable a distancia centro–punto).</summary>
    Private Shared Function TryCreateAuxLengthBetweenPoints(
        sheet As Sheet,
        dims As Dimensions,
        x1 As Double,
        y1 As Double,
        x2 As Double,
        y2 As Double,
        expectedM As Double,
        isHorizontal As Boolean,
        placementSide As DimensionSide,
        viewIdx As Integer,
        usedNominalSide As HashSet(Of String),
        styleObj As Object,
        log As DimensionLogger,
        kindTag As String) As Boolean

        If sheet Is Nothing OrElse dims Is Nothing Then Return False
        If expectedM < MinCenterSepM Then Return False
        If Not TryRegisterNominalSide(viewIdx, isHorizontal, placementSide, expectedM, usedNominalSide) Then
            log?.LogLine("[DIM][HOLE][SKIP] view=" & viewIdx.ToString(CultureInfo.InvariantCulture) &
                         " kind=" & kindTag &
                         " reason=duplicate_nominal_side mm=" & (expectedM * 1000.0R).ToString("0.###", CultureInfo.InvariantCulture))
            Return False
        End If

        Dim aux As Object = CreateSheetAuxLine(sheet, x1, y1, x2, y2)
        If aux Is Nothing Then
            usedNominalSide.Remove(BuildNominalSideKey(viewIdx, isHorizontal, placementSide, expectedM))
            log?.LogLine("[DIM][HOLE][FAIL] view=" & viewIdx.ToString(CultureInfo.InvariantCulture) &
                         " kind=" & kindTag & " reason=aux_line_create_failed")
            Return False
        End If

        Dim dimObj As Dimension = Nothing
        Try
            Dim refObj As Object = aux
            Try
                refObj = CallByName(aux, "Reference", CallType.Get)
            Catch
            End Try
            dimObj = TryCast(CallByName(dims, "AddLength", CallType.Method, refObj), Dimension)
        Catch ex As Exception
            log?.LogLine("[DIM][HOLE][FAIL] view=" & viewIdx.ToString(CultureInfo.InvariantCulture) &
                         " kind=" & kindTag & " method=AddLength ex=" & ex.Message)
        End Try

        TryHideAuxLine(aux)

        If dimObj Is Nothing Then
            usedNominalSide.Remove(BuildNominalSideKey(viewIdx, isHorizontal, placementSide, expectedM))
            Return False
        End If

        Dim valM As Double = ReadDimensionValueMeters(dimObj)
        If Not Double.IsNaN(valM) AndAlso Math.Abs(valM - expectedM) > ValueTolM Then
            Try
                dimObj.Delete()
            Catch
            End Try
            usedNominalSide.Remove(BuildNominalSideKey(viewIdx, isHorizontal, placementSide, expectedM))
            log?.LogLine("[DIM][HOLE][REJECT] view=" & viewIdx.ToString(CultureInfo.InvariantCulture) &
                         " kind=" & kindTag &
                         " reason=value_mismatch expected_mm=" & (expectedM * 1000.0R).ToString("0.###", CultureInfo.InvariantCulture) &
                         " actual_mm=" & (valM * 1000.0R).ToString("0.###", CultureInfo.InvariantCulture))
            Return False
        End If

        DrawingViewDimensionCreator.TryApplyStyleToDimension(dimObj, styleObj, log)
        TryApplyVisibleTrackDistance(dimObj, isHorizontal, log)
        log?.LogLine("[DIM][HOLE][OK] view=" & viewIdx.ToString(CultureInfo.InvariantCulture) &
                     " kind=" & kindTag &
                     " mm=" & (expectedM * 1000.0R).ToString("0.###", CultureInfo.InvariantCulture) &
                     " method=AddLength(aux_line)")
        Return True
    End Function

    Private Shared Function CreateSheetAuxLine(sheet As Sheet, x1 As Double, y1 As Double, x2 As Double, y2 As Double) As Object
        If sheet Is Nothing Then Return Nothing
        Try
            Return CallByName(sheet.Lines2d, "AddBy2Points", CallType.Method, x1, y1, x2, y2)
        Catch
        End Try
        Try
            Return CallByName(sheet.Lines2d, "AddLine", CallType.Method, x1, y1, x2, y2)
        Catch
            Return Nothing
        End Try
    End Function

    Private Shared Sub TryHideAuxLine(aux As Object)
        If aux Is Nothing Then Return
        Try
            CallByName(aux, "Visible", CallType.Let, False)
        Catch
        End Try
    End Sub

    Private Shared Function ReadDimensionValueMeters(d As Dimension) As Double
        If d Is Nothing Then Return Double.NaN
        Try
            Return CDbl(d.Value)
        Catch
        End Try
        Try
            Return CDbl(CallByName(d, "DisplayValue", CallType.Get))
        Catch
            Return Double.NaN
        End Try
    End Function

    Private Shared Sub TryApplyVisibleTrackDistance(dimObj As Dimension, isHorizontal As Boolean, log As DimensionLogger)
        If dimObj Is Nothing Then Return
        Try
            Dim td As Double = If(isHorizontal, 0.012R, 0.01R)
            CallByName(dimObj, "TrackDistance", CallType.Let, td)
            log?.LogLine("[DIM][HOLE][TRACK] TrackDistance=" & td.ToString("0.######", CultureInfo.InvariantCulture))
        Catch
        End Try
    End Sub

    Private Shared Function TryRegisterNominalSide(
        viewIdx As Integer,
        isHorizontal As Boolean,
        side As DimensionSide,
        nominalM As Double,
        used As HashSet(Of String)) As Boolean

        Dim key As String = BuildNominalSideKey(viewIdx, isHorizontal, side, nominalM)
        If used.Contains(key) Then Return False
        used.Add(key)
        Return True
    End Function

    Private Shared Function BuildNominalSideKey(
        viewIdx As Integer,
        isHorizontal As Boolean,
        side As DimensionSide,
        nominalM As Double) As String

        Dim mmBucket As Integer = CInt(Math.Round(nominalM * 1000.0R / ValueTolMm))
        Dim orient As String = If(isHorizontal, "H", "V")
        Dim sideName As String = side.ToString().ToUpperInvariant()
        Return "V" & viewIdx.ToString(CultureInfo.InvariantCulture) & "|" & orient & "|" & sideName & "|" & mmBucket.ToString(CultureInfo.InvariantCulture)
    End Function

    Private Shared Function TryCircleCenterRadiusView(c As Object, ByRef cx As Double, ByRef cy As Double, ByRef r As Double) As Boolean
        cx = 0 : cy = 0 : r = 0
        If c Is Nothing Then Return False
        Try
            If TypeOf c Is DVCircle2d Then
                Dim circ As DVCircle2d = CType(c, DVCircle2d)
                circ.GetCenterPoint(cx, cy)
                Try
                    r = CDbl(circ.Radius)
                Catch
                    r = CDbl(circ.Diameter) / 2.0R
                End Try
                Return r > MinHoleRadiusM
            End If
            CallByName(c, "GetCenterPoint", CallType.Method, cx, cy)
            r = Convert.ToDouble(CallByName(c, "Radius", CallType.Get), CultureInfo.InvariantCulture)
            Return r > MinHoleRadiusM
        Catch
            Return False
        End Try
    End Function

    Private Shared Function SafeCount(col As Object) As Integer
        If col Is Nothing Then Return 0
        Try
            Return CInt(CallByName(col, "Count", CallType.Get))
        Catch
            Return 0
        End Try
    End Function

End Class
