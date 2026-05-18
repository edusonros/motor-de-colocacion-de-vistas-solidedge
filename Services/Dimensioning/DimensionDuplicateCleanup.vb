Option Strict Off

Imports System.Collections.Generic
Imports System.Globalization
Imports System.Linq
Imports SolidEdgeDraft
Imports SolidEdgeConstants
Imports SolidEdgeFrameworkSupport

''' <summary>
''' Elimina cotas duplicadas tras el barrido: mismos keypoints y, por vista, mismo valor + orientación + lado (H/V).
''' </summary>
Friend NotInheritable Class DimensionDuplicateCleanup

    Private Const ValueToleranceMm As Double = 0.05R
    Private Const KeyPointToleranceM As Double = 0.001R
    Private Const ViewBoxMarginM As Double = 0.002R

    Friend NotInheritable Class ViewSlot
        Public ViewIndex As Integer
        Public Box As ViewSheetBoundingBox
    End Class

    Friend Shared Function BuildViewSlotsPublic(viewInfos As IList(Of DrawingViewGeometryInfo)) As List(Of ViewSlot)
        Return BuildViewSlots(viewInfos)
    End Function

    Friend Shared Function ResolveViewIndexPublic(midX As Double, midY As Double, slots As List(Of ViewSlot)) As Integer
        Return ResolveViewIndex(midX, midY, slots)
    End Function

    Friend Shared Function TryReadLinearLayoutPublic(
        d As Dimension,
        ByRef valMm As Double,
        ByRef isHorizontal As Boolean,
        ByRef midX As Double,
        ByRef midY As Double) As Boolean
        Return TryReadLinearLayout(d, valMm, isHorizontal, midX, midY)
    End Function

    Friend Shared Function TryIsRadialOrAngularDimensionPublic(d As Dimension) As Boolean
        Return TryIsRadialOrAngularDimension(d)
    End Function

    Private Sub New()
    End Sub

    Public Shared Function RunPostCreationCleanup(
        sheet As Sheet,
        viewInfos As IList(Of DrawingViewGeometryInfo),
        log As DimensionLogger,
        Optional valueToleranceMm As Double = ValueToleranceMm) As Integer

        Dim removed As Integer = RemoveDuplicateDimensionsByValueAndKeypoints(sheet, log, valueToleranceMm)
        removed += RemovePerViewDuplicatesByNominalAndOrientation(sheet, viewInfos, log, valueToleranceMm)
        Return removed
    End Function

    Public Shared Function RemoveDuplicateDimensionsByValueAndKeypoints(
        sheet As Sheet,
        log As DimensionLogger,
        Optional valueToleranceMm As Double = ValueToleranceMm) As Integer
        If sheet Is Nothing Then Return 0
        Dim dims As Dimensions = Nothing
        Try
            dims = sheet.Dimensions
        Catch
            Return 0
        End Try
        If dims Is Nothing OrElse dims.Count < 2 Then Return 0

        Dim seen As New Dictionary(Of String, Integer)(StringComparer.Ordinal)
        Dim removed As Integer = 0
        Dim tolMm As Double = Math.Max(0.001R, valueToleranceMm)

        For i As Integer = dims.Count To 1 Step -1
            Dim d As Dimension = Nothing
            Try
                d = CType(dims.Item(i), Dimension)
            Catch
                d = Nothing
            End Try
            If d Is Nothing Then Continue For

            Dim sig As String = TryBuildDimensionSignature(d, tolMm)
            If String.IsNullOrWhiteSpace(sig) Then Continue For

            If seen.ContainsKey(sig) Then
                Try
                    d.Delete()
                    removed += 1
                    log?.LogLine("[DIM][DEDUP][KP][DELETE] idx=" & i.ToString(CultureInfo.InvariantCulture) &
                                 " sig=" & sig)
                Catch ex As Exception
                    log?.LogLine("[DIM][DEDUP][KP][DELETE][WARN] idx=" & i.ToString(CultureInfo.InvariantCulture) & " " & ex.Message)
                End Try
            Else
                seen(sig) = i
            End If
        Next

        If removed > 0 AndAlso log IsNot Nothing Then
            log.LogLine("[DIM][DEDUP][KP][SUMMARY] removed=" & removed.ToString(CultureInfo.InvariantCulture))
        End If
        Return removed
    End Function

    ''' <summary>Una cota lineal por valor nominal, orientación (H/V) y lado dentro de cada vista.</summary>
    Public Shared Function RemovePerViewDuplicatesByNominalAndOrientation(
        sheet As Sheet,
        viewInfos As IList(Of DrawingViewGeometryInfo),
        log As DimensionLogger,
        Optional valueToleranceMm As Double = ValueToleranceMm) As Integer

        If sheet Is Nothing Then Return 0
        Dim slots As List(Of ViewSlot) = BuildViewSlots(viewInfos)
        If slots.Count = 0 Then Return 0

        Dim dims As Dimensions = Nothing
        Try
            dims = sheet.Dimensions
        Catch
            Return 0
        End Try
        If dims Is Nothing OrElse dims.Count < 2 Then Return 0

        Dim tolMm As Double = Math.Max(0.001R, valueToleranceMm)
        Dim seen As New Dictionary(Of String, Integer)(StringComparer.Ordinal)
        Dim removed As Integer = 0

        For i As Integer = dims.Count To 1 Step -1
            Dim d As Dimension = Nothing
            Try
                d = CType(dims.Item(i), Dimension)
            Catch
                d = Nothing
            End Try
            If d Is Nothing Then Continue For
            If TryIsRadialOrAngularDimension(d) Then Continue For

            Dim valMm As Double
            Dim isHorizontal As Boolean
            Dim midX As Double
            Dim midY As Double
            If Not TryReadLinearLayout(d, valMm, isHorizontal, midX, midY) Then Continue For

            Dim viewIdx As Integer = ResolveViewIndex(midX, midY, slots)
            If viewIdx < 0 Then Continue For

            Dim valBucket As Integer = CInt(Math.Round(valMm / tolMm))
            Dim orient As String = If(isHorizontal, "H", "V")
            Dim side As String = InferLaneSide(isHorizontal, midX, midY, slots, viewIdx)
            Dim sig As String = "DV" & viewIdx.ToString(CultureInfo.InvariantCulture) &
                                "|" & orient & "|" & side & "|" & valBucket.ToString(CultureInfo.InvariantCulture)

            If seen.ContainsKey(sig) Then
                Try
                    d.Delete()
                    removed += 1
                    log?.LogLine("[DIM][DEDUP][VIEW][DELETE] idx=" & i.ToString(CultureInfo.InvariantCulture) &
                                 " view=" & viewIdx.ToString(CultureInfo.InvariantCulture) &
                                 " orient=" & orient &
                                 " mm=" & valMm.ToString("0.###", CultureInfo.InvariantCulture) &
                                 " sig=" & sig)
                Catch ex As Exception
                    log?.LogLine("[DIM][DEDUP][VIEW][DELETE][WARN] idx=" & i.ToString(CultureInfo.InvariantCulture) & " " & ex.Message)
                End Try
            Else
                seen(sig) = i
            End If
        Next

        If removed > 0 AndAlso log IsNot Nothing Then
            log.LogLine("[DIM][DEDUP][VIEW][SUMMARY] removed=" & removed.ToString(CultureInfo.InvariantCulture) &
                        " unique=" & seen.Count.ToString(CultureInfo.InvariantCulture))
        End If
        Return removed
    End Function

    Private Shared Function BuildViewSlots(viewInfos As IList(Of DrawingViewGeometryInfo)) As List(Of ViewSlot)
        Dim slots As New List(Of ViewSlot)()
        If viewInfos Is Nothing Then Return slots
        For Each info As DrawingViewGeometryInfo In viewInfos
            If info Is Nothing OrElse info.Box.Width < 1.0E-9R OrElse info.Box.Height < 1.0E-9R Then Continue For
            slots.Add(New ViewSlot With {.ViewIndex = info.ViewIndex, .Box = info.Box})
        Next
        Return slots
    End Function

    Private Shared Function ResolveViewIndex(midX As Double, midY As Double, slots As List(Of ViewSlot)) As Integer
        Dim bestIdx As Integer = -1
        Dim bestDist As Double = Double.MaxValue
        Dim bestContainedIdx As Integer = -1
        Dim bestContainedDist As Double = Double.MaxValue
        For Each slot As ViewSlot In slots
            If slot Is Nothing Then Continue For
            Dim b = slot.Box
            Dim cx As Double = (b.MinX + b.MaxX) * 0.5R
            Dim cy As Double = (b.MinY + b.MaxY) * 0.5R
            Dim dist As Double = (midX - cx) * (midX - cx) + (midY - cy) * (midY - cy)
            If dist < bestDist Then
                bestDist = dist
                bestIdx = slot.ViewIndex
            End If
            If midX >= b.MinX - ViewBoxMarginM AndAlso midX <= b.MaxX + ViewBoxMarginM AndAlso
               midY >= b.MinY - ViewBoxMarginM AndAlso midY <= b.MaxY + ViewBoxMarginM Then
                If dist < bestContainedDist Then
                    bestContainedDist = dist
                    bestContainedIdx = slot.ViewIndex
                End If
            End If
        Next
        If bestContainedIdx >= 0 Then Return bestContainedIdx
        Return bestIdx
    End Function

    Private Shared Function InferLaneSide(
        isHorizontal As Boolean,
        midX As Double,
        midY As Double,
        slots As List(Of ViewSlot),
        viewIdx As Integer) As String

        Dim slot = slots.FirstOrDefault(Function(s) s IsNot Nothing AndAlso s.ViewIndex = viewIdx)
        If slot Is Nothing Then Return "GEN"
        Dim b = slot.Box
        Dim cx As Double = (b.MinX + b.MaxX) * 0.5R
        Dim cy As Double = (b.MinY + b.MaxY) * 0.5R
        If isHorizontal Then
            Return If(midY >= cy, "TOP", "BOTTOM")
        End If
        Return If(midX >= cx, "RIGHT", "LEFT")
    End Function

    Private Shared Function TryReadLinearLayout(
        d As Dimension,
        ByRef valMm As Double,
        ByRef isHorizontal As Boolean,
        ByRef midX As Double,
        ByRef midY As Double) As Boolean

        valMm = 0R
        isHorizontal = False
        midX = 0R
        midY = 0R
        If d Is Nothing Then Return False

        Try
            valMm = CDbl(d.Value) * 1000.0R
            If Double.IsNaN(valMm) OrElse Double.IsInfinity(valMm) OrElse valMm <= 0.001R Then Return False
        Catch
            Return False
        End Try

        Dim x1 As Double = 0, y1 As Double = 0, x2 As Double = 0, y2 As Double = 0
        Try
            d.Range(x1, y1, x2, y2)
        Catch
            Return False
        End Try
        Dim minX = Math.Min(x1, x2)
        Dim maxX = Math.Max(x1, x2)
        Dim minY = Math.Min(y1, y2)
        Dim maxY = Math.Max(y1, y2)
        Dim sx = maxX - minX
        Dim sy = maxY - minY
        If sx + sy < 1.0E-12R Then Return False
        isHorizontal = sx >= sy
        midX = (minX + maxX) * 0.5R
        midY = (minY + maxY) * 0.5R
        Return True
    End Function

    Private Shared Function TryIsRadialOrAngularDimension(d As Dimension) As Boolean
        If d Is Nothing Then Return False
        Try
            Dim display As String = Convert.ToString(CallByName(d, "DisplayString", CallType.Get), CultureInfo.InvariantCulture)
            If Not String.IsNullOrWhiteSpace(display) Then
                Dim t = display.TrimStart()
                If t.Length > 0 AndAlso (t(0) = "R"c OrElse t(0) = "r"c OrElse t.StartsWith("Ø", StringComparison.Ordinal) OrElse t.StartsWith("ø", StringComparison.Ordinal)) Then
                    Return True
                End If
            End If
        Catch
        End Try
        Try
            Dim typeName As String = Convert.ToString(CallByName(d, "DimensionType", CallType.Get), CultureInfo.InvariantCulture)
            If String.IsNullOrWhiteSpace(typeName) Then Return False
            Dim u = typeName.ToUpperInvariant()
            If u.Contains("RADIAL") OrElse u.Contains("RADIUS") OrElse u.Contains("ANGULAR") OrElse u.Contains("ARC") Then
                Return True
            End If
        Catch
        End Try
        Return False
    End Function

    Friend Shared Function TryBuildDimensionSignature(d As Dimension, valueToleranceMm As Double) As String
        If d Is Nothing Then Return ""
        Dim valMm As Double
        Try
            valMm = CDbl(d.Value) * 1000.0R
        Catch
            Return ""
        End Try
        Dim valBucket As Integer = CInt(Math.Round(valMm / Math.Max(valueToleranceMm, 0.001R)))

        Dim kpParts As New List(Of String)()
        Dim nKp As Integer = 0
        Try
            nKp = CInt(d.KeyPointCount)
        Catch
            nKp = 0
        End Try
        For ki As Integer = 0 To Math.Min(nKp, 8) - 1
            Try
                Dim px As Double = 0, py As Double = 0, pz As Double = 0
                Dim kpt As SolidEdgeConstants.KeyPointType
                Dim hdl As SolidEdgeConstants.HandleType
                d.GetKeyPoint(ki, px, py, pz, kpt, hdl)
                Dim bx As Integer = CInt(Math.Round(px / KeyPointToleranceM))
                Dim by As Integer = CInt(Math.Round(py / KeyPointToleranceM))
                kpParts.Add(bx.ToString(CultureInfo.InvariantCulture) & "," & by.ToString(CultureInfo.InvariantCulture))
            Catch
            End Try
        Next
        kpParts.Sort(StringComparer.Ordinal)
        Return "V" & valBucket.ToString(CultureInfo.InvariantCulture) & "|KP|" & String.Join(";", kpParts)
    End Function

End Class
