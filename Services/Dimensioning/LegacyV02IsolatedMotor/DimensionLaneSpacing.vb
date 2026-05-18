Option Strict Off

Imports System.Collections.Generic
Imports System.Globalization
Imports System.Linq
Imports SolidEdgeDraft
Imports SolidEdgeFrameworkSupport

Namespace Extraer_dft_dxf_flatdxf.LegacyV02Dimensioning

''' <summary>
''' Espaciado en carriles: primera cota a MinGapFromView (12 mm) y sucesivas +GapBetweenDimensionRows (10 mm) en la misma cara.
''' En la vista principal, empuja cotas cuyo centro cae dentro del bbox hacia el exterior.
''' </summary>
Friend NotInheritable Class DimensionLaneSpacing

    Private Const InsideViewMarginM As Double = 0.0015R

    Private NotInheritable Class LaneDim
        Public D As Dimension
        Public ViewIndex As Integer
        Public IsHorizontal As Boolean
        Public Side As String
        Public DistFromViewEdge As Double
        Public Track0 As Double
    End Class

    Private Sub New()
    End Sub

    Public Shared Function Apply(
        sheet As Sheet,
        viewInfos As IList(Of DrawingViewGeometryInfo),
        norm As DimensioningNormConfig,
        log As DimensionLogger) As Integer

        If sheet Is Nothing OrElse norm Is Nothing Then Return 0
        If norm.SuppressDimensionTrackDistanceSpacing Then Return 0

        Dim slots As List(Of DimensionDuplicateCleanup.ViewSlot) = DimensionDuplicateCleanup.BuildViewSlotsPublic(viewInfos)
        If slots.Count = 0 Then Return 0

        Dim primaryIdx As Integer = ResolvePrimaryViewIndex(viewInfos)
        Dim firstGap As Double = Math.Max(0.008R, norm.MinGapFromView)
        Dim rowGap As Double = Math.Max(0.006R, norm.GapBetweenDimensionRows)

        Dim dims As Dimensions = Nothing
        Try
            dims = sheet.Dimensions
        Catch
            Return 0
        End Try
        If dims Is Nothing OrElse dims.Count = 0 Then Return 0

        Dim items As New List(Of LaneDim)()
        For i As Integer = 1 To dims.Count
            Dim d As Dimension = Nothing
            Try
                d = CType(dims.Item(i), Dimension)
            Catch
                d = Nothing
            End Try
            If d Is Nothing Then Continue For
            If DimensionDuplicateCleanup.TryIsRadialOrAngularDimensionPublic(d) Then Continue For

            Dim valMm As Double
            Dim isH As Boolean
            Dim midX As Double
            Dim midY As Double
            If Not DimensionDuplicateCleanup.TryReadLinearLayoutPublic(d, valMm, isH, midX, midY) Then Continue For

            Dim viewIdx As Integer = DimensionDuplicateCleanup.ResolveViewIndexPublic(midX, midY, slots)
            If viewIdx < 0 Then Continue For

            Dim slot = slots.FirstOrDefault(Function(s) s.ViewIndex = viewIdx)
            If slot Is Nothing Then Continue For

            Dim side As String = InferSide(isH, midX, midY, slot.Box)
            Dim distEdge As Double = DistFromViewEdgeM(isH, side, midX, midY, slot.Box)
            Dim td0 As Double = 0R
            Try
                td0 = CDbl(CallByName(d, "TrackDistance", CallType.Get))
            Catch
                td0 = 0R
            End Try

            items.Add(New LaneDim With {
                .D = d,
                .ViewIndex = viewIdx,
                .IsHorizontal = isH,
                .Side = side,
                .DistFromViewEdge = distEdge,
                .Track0 = td0
            })
        Next

        If items.Count = 0 Then Return 0

        Dim adjusted As Integer = 0
        Dim groups = items.GroupBy(Function(x) x.ViewIndex.ToString(CultureInfo.InvariantCulture) & "|" &
                                              If(x.IsHorizontal, "H", "V") & "|" & x.Side)

        For Each g In groups
            Dim ordered = g.OrderBy(Function(x) x.DistFromViewEdge).ThenBy(Function(x) x.Track0).ToList()
            For laneIdx As Integer = 0 To ordered.Count - 1
                Dim it = ordered(laneIdx)
                Dim tdTarget As Double = firstGap + laneIdx * rowGap
                If TrySetTrackDistance(it.D, tdTarget) Then
                    adjusted += 1
                    log?.LogLine("[DIM][LANE][SET] view=" & it.ViewIndex.ToString(CultureInfo.InvariantCulture) &
                                 " side=" & it.Side &
                                 " orient=" & If(it.IsHorizontal, "H", "V") &
                                 " laneIdx=" & laneIdx.ToString(CultureInfo.InvariantCulture) &
                                 " track_m=" & tdTarget.ToString("0.######", CultureInfo.InvariantCulture))
                End If
            Next
        Next

        If norm.PreferOutsideDimensions AndAlso primaryIdx >= 0 Then
            Dim primarySlot = slots.FirstOrDefault(Function(s) s.ViewIndex = primaryIdx)
            If primarySlot IsNot Nothing Then
                adjusted += PushInsideDimensionsOutside(primarySlot, items.Where(Function(x) x.ViewIndex = primaryIdx).ToList(), firstGap, norm, log)
            End If
        End If

        If adjusted > 0 Then
            log?.LogLine("[DIM][LANE][SUMMARY] adjusted=" & adjusted.ToString(CultureInfo.InvariantCulture) &
                         " first_mm=" & (firstGap * 1000.0R).ToString("0.#", CultureInfo.InvariantCulture) &
                         " step_mm=" & (rowGap * 1000.0R).ToString("0.#", CultureInfo.InvariantCulture))
        End If
        Return adjusted
    End Function

    Private Shared Function ResolvePrimaryViewIndex(viewInfos As IList(Of DrawingViewGeometryInfo)) As Integer
        If viewInfos Is Nothing OrElse viewInfos.Count = 0 Then Return -1
        Dim best As DrawingViewGeometryInfo = Nothing
        Dim bestArea As Double = -1R
        For Each info In viewInfos
            If info Is Nothing Then Continue For
            Dim a As Double = info.Box.Width * info.Box.Height
            If a > bestArea Then
                bestArea = a
                best = info
            End If
        Next
        Return If(best Is Nothing, -1, best.ViewIndex)
    End Function

    Private Shared Function InferSide(isHorizontal As Boolean, midX As Double, midY As Double, box As ViewSheetBoundingBox) As String
        Dim cx As Double = (box.MinX + box.MaxX) * 0.5R
        Dim cy As Double = (box.MinY + box.MaxY) * 0.5R
        If isHorizontal Then
            Return If(midY >= cy, "TOP", "BOTTOM")
        End If
        Return If(midX >= cx, "RIGHT", "LEFT")
    End Function

    Private Shared Function DistFromViewEdgeM(isHorizontal As Boolean, side As String, midX As Double, midY As Double, box As ViewSheetBoundingBox) As Double
        Select Case side
            Case "TOP"
                Return Math.Max(0R, midY - box.MaxY)
            Case "BOTTOM"
                Return Math.Max(0R, box.MinY - midY)
            Case "RIGHT"
                Return Math.Max(0R, midX - box.MaxX)
            Case "LEFT"
                Return Math.Max(0R, box.MinX - midX)
            Case Else
                Return 0R
        End Select
    End Function

    Private Shared Function PushInsideDimensionsOutside(
        slot As DimensionDuplicateCleanup.ViewSlot,
        dims As List(Of LaneDim),
        minGap As Double,
        norm As DimensioningNormConfig,
        log As DimensionLogger) As Integer

        If slot Is Nothing OrElse dims Is Nothing OrElse Not norm.AvoidInsideContour Then Return 0
        Dim n As Integer = 0
        Dim b = slot.Box
        Dim shrinkX As Double = b.Width * 0.04R
        Dim shrinkY As Double = b.Height * 0.04R
        Dim innerMinX = b.MinX + shrinkX
        Dim innerMaxX = b.MaxX - shrinkX
        Dim innerMinY = b.MinY + shrinkY
        Dim innerMaxY = b.MaxY - shrinkY

        For Each it In dims
            If it Is Nothing OrElse it.D Is Nothing Then Continue For
            Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
            Try
                it.D.Range(x1, y1, x2, y2)
            Catch
                Continue For
            End Try
            Dim cx As Double = (Math.Min(x1, x2) + Math.Max(x1, x2)) * 0.5R
            Dim cy As Double = (Math.Min(y1, y2) + Math.Max(y1, y2)) * 0.5R
            If cx <= innerMinX OrElse cx >= innerMaxX OrElse cy <= innerMinY OrElse cy >= innerMaxY Then Continue For

            Dim depthIn As Double = Math.Min(Math.Min(cx - innerMinX, innerMaxX - cx),
                                             Math.Min(cy - innerMinY, innerMaxY - cy))
            Dim tdTarget As Double = Math.Max(minGap * 2.5R, minGap + depthIn + minGap)
            If TrySetTrackDistance(it.D, tdTarget) Then
                n += 1
                log?.LogLine("[DIM][OUTSIDE][PUSH] view=" & slot.ViewIndex.ToString(CultureInfo.InvariantCulture) &
                             " side=" & it.Side &
                             " track_m=" & tdTarget.ToString("0.######", CultureInfo.InvariantCulture))
            End If
        Next
        Return n
    End Function

    Private Shared Function TrySetTrackDistance(d As Dimension, td As Double) As Boolean
        If d Is Nothing OrElse td < 0R Then Return False
        Try
            Dim cur As Double = CDbl(CallByName(d, "TrackDistance", CallType.Get))
            If Math.Abs(cur - td) < 0.00005R Then Return False
            CallByName(d, "TrackDistance", CallType.Let, td)
            Return True
        Catch
            Return False
        End Try
    End Function

End Class

End Namespace
