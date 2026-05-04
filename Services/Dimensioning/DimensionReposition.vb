Option Strict Off

Imports System.Collections.Generic
Imports System.Globalization
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports SolidEdgeFrameworkSupport
Imports FrameworkDimension = SolidEdgeFrameworkSupport.Dimension

''' <summary>
''' Tras <see cref="Dimensions.AddDistanceBetweenObjects"/>, Solid Edge puede recolocar la cota.
''' Por defecto solo se registra diagnóstico <c>[DIM][REPOS]</c> sin modificar el objeto COM.
''' La modificación activa (<see cref="Dimension.TrackDistance"/>, <see cref="Dimension.SetKeyPoint"/>) solo si
''' <see cref="EnableExperimentalDimensionRepositioning"/> es <c>True</c> (pruebas controladas).
''' </summary>
Friend NotInheritable Class DimensionReposition

    ''' <summary>
    ''' <c>False</c> (predeterminado): estabilidad visual; no se altera la cota tras insertarla.
    ''' <c>True</c>: reactiva la ruta experimental (iteración TrackDistance + SetKeyPoint uniforme) — usar solo en entorno de prueba.
    ''' </summary>
    Friend Shared EnableExperimentalDimensionRepositioning As Boolean = False

    Private Const TolM As Double = 0.00008R
    Private Const MaxTrackIterations As Integer = 12
    Private Const SameLineTolM As Double = 0.0004R

    Private Sub New()
    End Sub

    ''' <summary>
    ''' Diagnóstico post-inserción: target, Range, propiedades de lectura. Opcionalmente reposición experimental.
    ''' </summary>
    ''' <returns>Si experimental está desactivado, <c>True</c> (no se ha fallado nada). Si experimental activo, éxito del intento de alinear.</returns>
    Friend Shared Function TryRepositionInsertedDimension(
        d As FrameworkDimension,
        targetSheetX As Double,
        targetSheetY As Double,
        axisLabel As String,
        frame As ViewPlacementFrame,
        log As DimensionLogger) As Boolean

        If d Is Nothing OrElse log Is Nothing Then Return False

        Dim horizontal As Boolean = String.Equals(axisLabel, "horizontal", StringComparison.OrdinalIgnoreCase)
        log.Repos(String.Format(CultureInfo.InvariantCulture,
            "target=({0:0.######},{1:0.######}) orient={2}",
            targetSheetX, targetSheetY, If(horizontal, "horizontal", "vertical")))
        log.Repos("EnableExperimentalDimensionRepositioning=" & EnableExperimentalDimensionRepositioning.ToString())

        LogRangeLine(d, frame, log, "before")
        TryReadTrackDistances(d, log)

        InspectDimensionReadOnly(d, frame, log)

        If Not EnableExperimentalDimensionRepositioning Then
            log.Repos("modo=READONLY (sin TrackDistance iterado, sin SetKeyPoint, sin traslaciones)")
            LogRangeLine(d, frame, log, "after")
            log.Repos("property used = (ninguna — reposición experimental desactivada)")
            log.Repos("moved=n/a (solo diagnóstico)")
            Return True
        End If

        ' --- Ruta experimental (solo si EnableExperimentalDimensionRepositioning=True) ---
        log.Repos("modo=EXPERIMENTAL — aplicando TrackDistance/SetKeyPoint (riesgo de deformación)")
        Dim propUsed As String = Nothing
        Dim moved As Boolean = False
        Try
            If horizontal Then
                moved = TryAlignHorizontal(d, targetSheetY, log, propUsed)
            Else
                moved = TryAlignVertical(d, targetSheetX, log, propUsed)
            End If
        Catch ex As Exception
            log.Repos("EX durante reposición experimental: " & FormatEx(ex))
            LogWritableLocationPropertyDiagnostics(d, log)
        End Try

        LogRangeLine(d, frame, log, "after")
        TryReadTrackDistances(d, log)

        If Not String.IsNullOrEmpty(propUsed) Then
            log.Repos("property used = " & propUsed)
        Else
            log.Repos("property used = (ninguna aplicada)")
        End If

        Dim okFinal As Boolean = False
        If horizontal Then
            Dim yLine As Double = EstimateHorizontalDimensionLineY(d)
            okFinal = Not Double.IsNaN(yLine) AndAlso Math.Abs(yLine - targetSheetY) < TolM * 5
        Else
            Dim xLine As Double = EstimateVerticalDimensionLineX(d)
            okFinal = Not Double.IsNaN(xLine) AndAlso Math.Abs(xLine - targetSheetX) < TolM * 5
        End If

        log.Repos("moved=" & If(okFinal, "True", "False"))

        If Not okFinal Then
            log.Repos("diagnóstico experimental: ver TrackDistance en log anterior")
            LogWritableLocationPropertyDiagnostics(d, log)
        End If

        Return okFinal
    End Function

    ''' <summary>
    ''' Inspección solo lectura del objeto <see cref="Dimension"/> para investigar qué expone la API
    ''' (keypoints, offsets de texto, ejes, tipos) sin modificar geometría. Pensado para identificar
    ''' futuras propiedades que muevan solo línea/texto de cota.
    ''' </summary>
    Friend Shared Sub InspectDimensionReadOnly(d As FrameworkDimension, frame As ViewPlacementFrame, log As DimensionLogger)
        If d Is Nothing OrElse log Is Nothing Then Return
        log.Repos("--- inspección segura (solo lectura, sin mutar) ---")
        Try
            log.Repos("CLR type=" & d.GetType().FullName)
        Catch
        End Try

        TryReadTrackDistances(d, log)

        Try
            log.Repos("DimensionType=" & CInt(d.DimensionType).ToString(CultureInfo.InvariantCulture))
        Catch ex As Exception
            log.Repos("DimensionType: no legible — " & ex.Message)
        End Try

        Try
            log.Repos("MeasurementAxisDirection=" & CInt(d.MeasurementAxisDirection).ToString(CultureInfo.InvariantCulture))
        Catch ex As Exception
            log.Repos("MeasurementAxisDirection: no legible — " & ex.Message)
        End Try

        Try
            log.Repos("MeasurementAxisEx=" & CInt(d.MeasurementAxisEx).ToString(CultureInfo.InvariantCulture))
        Catch ex As Exception
            log.Repos("MeasurementAxisEx: no legible — " & ex.Message)
        End Try

        Try
            Dim tox As Double = 0, toy As Double = 0
            d.GetTextOffsets(tox, toy)
            log.Repos(String.Format(CultureInfo.InvariantCulture,
                "GetTextOffsets dx={0:0.######} dy={1:0.######} (relativo a posición calculada del texto)", tox, toy))
        Catch ex As Exception
            log.Repos("GetTextOffsets: no legible — " & ex.Message)
        End Try

        Try
            Dim nKp As Integer = CInt(d.KeyPointCount)
            log.Repos("KeyPointCount=" & nKp.ToString(CultureInfo.InvariantCulture))
            Dim lim As Integer = Math.Min(nKp, 16)
            For i As Integer = 0 To lim - 1
                Dim px As Double = 0, py As Double = 0, pz As Double = 0
                Dim kpt As SolidEdgeConstants.KeyPointType
                Dim hdl As SolidEdgeConstants.HandleType
                d.GetKeyPoint(i, px, py, pz, kpt, hdl)
                If frame IsNot Nothing Then
                    log.Repos(String.Format(CultureInfo.InvariantCulture,
                        "KeyPoint[{0}] hoja=({1:0.######},{2:0.######}) local=({3:0.######},{4:0.######}) kpt={5} hdl={6}",
                        i, px, py, frame.FromSheetX(px), frame.FromSheetY(py), CInt(kpt), CInt(hdl)))
                Else
                    log.Repos(String.Format(CultureInfo.InvariantCulture,
                        "KeyPoint[{0}] hoja=({1:0.######},{2:0.######}) kpt={3} hdl={4}",
                        i, px, py, CInt(kpt), CInt(hdl)))
                End If
            Next
        Catch ex As Exception
            log.Repos("KeyPoints inspección: " & ex.Message)
        End Try

        LogWritableLocationPropertyDiagnostics(d, log)
        log.Repos("--- fin inspección segura ---")
    End Sub

    Private Shared Function TryAlignHorizontal(d As FrameworkDimension, targetY As Double, log As DimensionLogger, ByRef propUsed As String) As Boolean
        Dim y0 As Double = EstimateHorizontalDimensionLineY(d)
        If Double.IsNaN(y0) Then
            log.Repos("no se pudo estimar Y de línea de cota (keypoints); se prueba TrackDistance directo.")
        Else
            log.Repos(String.Format(CultureInfo.InvariantCulture,
                "estimado Y línea cota (hoja)={0:0.######} objetivo={1:0.######} Δ={2:0.######}",
                y0, targetY, targetY - y0))
        End If

        Dim iter As Integer = 0
        While iter < MaxTrackIterations
            Dim yLine As Double = EstimateHorizontalDimensionLineY(d)
            If Double.IsNaN(yLine) Then yLine = TryGetRangeCenterY(d)
            If Double.IsNaN(yLine) Then Exit While
            Dim errY As Double = targetY - yLine
            If Math.Abs(errY) < TolM Then
                propUsed = "TrackDistance (iterado)"
                Return True
            End If
            Dim td As Double = d.TrackDistance
            d.TrackDistance = td + errY
            log.Repos(String.Format(CultureInfo.InvariantCulture,
                "TrackDistance: {0:0.######} + ΔY {1:0.######} => {2:0.######}", td, errY, d.TrackDistance))
            iter += 1
        End While

        If TryUniformShiftKeypointsY(d, targetY, log) Then
            propUsed = "SetKeyPoint (traslación uniforme Y)"
            Dim yCheck As Double = EstimateHorizontalDimensionLineY(d)
            Return Not Double.IsNaN(yCheck) AndAlso Math.Abs(yCheck - targetY) < TolM * 6
        End If

        Return False
    End Function

    Private Shared Function TryAlignVertical(d As FrameworkDimension, targetX As Double, log As DimensionLogger, ByRef propUsed As String) As Boolean
        Dim x0 As Double = EstimateVerticalDimensionLineX(d)
        If Double.IsNaN(x0) Then
            log.Repos("no se pudo estimar X de línea de cota (keypoints); se prueba TrackDistance directo.")
        Else
            log.Repos(String.Format(CultureInfo.InvariantCulture,
                "estimado X línea cota (hoja)={0:0.######} objetivo={1:0.######} Δ={2:0.######}",
                x0, targetX, targetX - x0))
        End If

        Dim iter As Integer = 0
        While iter < MaxTrackIterations
            Dim xLine As Double = EstimateVerticalDimensionLineX(d)
            If Double.IsNaN(xLine) Then xLine = TryGetRangeCenterX(d)
            If Double.IsNaN(xLine) Then Exit While
            Dim errX As Double = targetX - xLine
            If Math.Abs(errX) < TolM Then
                propUsed = "TrackDistance (iterado)"
                Return True
            End If
            Dim td As Double = d.TrackDistance
            d.TrackDistance = td + errX
            log.Repos(String.Format(CultureInfo.InvariantCulture,
                "TrackDistance: {0:0.######} + ΔX {1:0.######} => {2:0.######}", td, errX, d.TrackDistance))
            iter += 1
        End While

        If TryUniformShiftKeypointsX(d, targetX, log) Then
            propUsed = "SetKeyPoint (traslación uniforme X)"
            Dim xCheck As Double = EstimateVerticalDimensionLineX(d)
            Return Not Double.IsNaN(xCheck) AndAlso Math.Abs(xCheck - targetX) < TolM * 6
        End If

        Return False
    End Function

    Private Shared Function EstimateHorizontalDimensionLineY(d As FrameworkDimension) As Double
        Dim pts As List(Of Tuple(Of Double, Double)) = CollectKeyPointsXY(d)
        If pts Is Nothing OrElse pts.Count < 2 Then Return Double.NaN
        Dim bestDx As Double = -1
        Dim bestY As Double = Double.NaN
        For i As Integer = 0 To pts.Count - 2
            For j As Integer = i + 1 To pts.Count - 1
                Dim dy As Double = Math.Abs(pts(i).Item2 - pts(j).Item2)
                If dy > SameLineTolM Then Continue For
                Dim dx As Double = Math.Abs(pts(i).Item1 - pts(j).Item1)
                If dx > bestDx Then
                    bestDx = dx
                    bestY = (pts(i).Item2 + pts(j).Item2) / 2.0R
                End If
            Next
        Next
        If bestDx > 1.0E-9 Then Return bestY
        Return Double.NaN
    End Function

    Private Shared Function EstimateVerticalDimensionLineX(d As FrameworkDimension) As Double
        Dim pts As List(Of Tuple(Of Double, Double)) = CollectKeyPointsXY(d)
        If pts Is Nothing OrElse pts.Count < 2 Then Return Double.NaN
        Dim bestDy As Double = -1
        Dim bestX As Double = Double.NaN
        For i As Integer = 0 To pts.Count - 2
            For j As Integer = i + 1 To pts.Count - 1
                Dim dx As Double = Math.Abs(pts(i).Item1 - pts(j).Item1)
                If dx > SameLineTolM Then Continue For
                Dim dy As Double = Math.Abs(pts(i).Item2 - pts(j).Item2)
                If dy > bestDy Then
                    bestDy = dy
                    bestX = (pts(i).Item1 + pts(j).Item1) / 2.0R
                End If
            Next
        Next
        If bestDy > 1.0E-9 Then Return bestX
        Return Double.NaN
    End Function

    Private Shared Function TryGetRangeCenterX(d As FrameworkDimension) As Double
        Try
            Dim xMn As Double = 0, yMn As Double = 0, xMx As Double = 0, yMx As Double = 0
            d.Range(xMn, yMn, xMx, yMx)
            Return (Math.Min(xMn, xMx) + Math.Max(xMn, xMx)) / 2.0R
        Catch
            Return Double.NaN
        End Try
    End Function

    Private Shared Function TryGetRangeCenterY(d As FrameworkDimension) As Double
        Try
            Dim xMn As Double = 0, yMn As Double = 0, xMx As Double = 0, yMx As Double = 0
            d.Range(xMn, yMn, xMx, yMx)
            Return (Math.Min(yMn, yMx) + Math.Max(yMn, yMx)) / 2.0R
        Catch
            Return Double.NaN
        End Try
    End Function

    Private Shared Function CollectKeyPointsXY(d As FrameworkDimension) As List(Of Tuple(Of Double, Double))
        Dim list As New List(Of Tuple(Of Double, Double))
        Try
            Dim n As Integer = CInt(d.KeyPointCount)
            If n <= 0 Then Return list
            For i As Integer = 0 To n - 1
                Dim px As Double = 0, py As Double = 0, pz As Double = 0
                Dim kpt As SolidEdgeConstants.KeyPointType
                Dim hdl As SolidEdgeConstants.HandleType
                d.GetKeyPoint(i, px, py, pz, kpt, hdl)
                list.Add(New Tuple(Of Double, Double)(px, py))
            Next
        Catch
            Return Nothing
        End Try
        Return list
    End Function

    Private Shared Function TryUniformShiftKeypointsY(d As FrameworkDimension, targetY As Double, log As DimensionLogger) As Boolean
        Dim yLine As Double = EstimateHorizontalDimensionLineY(d)
        If Double.IsNaN(yLine) Then yLine = TryGetRangeCenterY(d)
        If Double.IsNaN(yLine) Then Return False
        Dim delta As Double = targetY - yLine
        If Math.Abs(delta) < 1.0E-12 Then Return True
        Try
            Dim n As Integer = CInt(d.KeyPointCount)
            For i As Integer = 0 To n - 1
                Dim px As Double = 0, py As Double = 0, pz As Double = 0
                Dim kpt As SolidEdgeConstants.KeyPointType
                Dim hdl As SolidEdgeConstants.HandleType
                d.GetKeyPoint(i, px, py, pz, kpt, hdl)
                d.SetKeyPoint(i, px, py + delta, 0.0)
            Next
            log.Repos(String.Format(CultureInfo.InvariantCulture, "SetKeyPoint: ΔY uniforme={0:0.######} (keypoints={1})", delta, n))
            Return True
        Catch ex As Exception
            log.Repos("SetKeyPoint ΔY uniforme EX: " & FormatEx(ex))
            Return False
        End Try
    End Function

    Private Shared Function TryUniformShiftKeypointsX(d As FrameworkDimension, targetX As Double, log As DimensionLogger) As Boolean
        Dim xLine As Double = EstimateVerticalDimensionLineX(d)
        If Double.IsNaN(xLine) Then xLine = TryGetRangeCenterX(d)
        If Double.IsNaN(xLine) Then Return False
        Dim delta As Double = targetX - xLine
        If Math.Abs(delta) < 1.0E-12 Then Return True
        Try
            Dim n As Integer = CInt(d.KeyPointCount)
            For i As Integer = 0 To n - 1
                Dim px As Double = 0, py As Double = 0, pz As Double = 0
                Dim kpt As SolidEdgeConstants.KeyPointType
                Dim hdl As SolidEdgeConstants.HandleType
                d.GetKeyPoint(i, px, py, pz, kpt, hdl)
                d.SetKeyPoint(i, px + delta, py, 0.0)
            Next
            log.Repos(String.Format(CultureInfo.InvariantCulture, "SetKeyPoint: ΔX uniforme={0:0.######} (keypoints={1})", delta, n))
            Return True
        Catch ex As Exception
            log.Repos("SetKeyPoint ΔX uniforme EX: " & FormatEx(ex))
            Return False
        End Try
    End Function

    Private Shared Sub LogRangeLine(d As FrameworkDimension, frame As ViewPlacementFrame, log As DimensionLogger, tag As String)
        If d Is Nothing OrElse log Is Nothing Then Return
        Try
            Dim xMn As Double = 0, yMn As Double = 0, xMx As Double = 0, yMx As Double = 0
            d.Range(xMn, yMn, xMx, yMx)
            Dim minX As Double = Math.Min(xMn, xMx)
            Dim maxX As Double = Math.Max(xMn, xMx)
            Dim minY As Double = Math.Min(yMn, yMx)
            Dim maxY As Double = Math.Max(yMn, yMx)
            If frame IsNot Nothing Then
                log.Repos(String.Format(CultureInfo.InvariantCulture,
                    "{0} Range hoja=({1:0.######},{2:0.######})-({3:0.######},{4:0.######}) local=({5:0.######},{6:0.######})-({7:0.######},{8:0.######})",
                    tag, minX, minY, maxX, maxY,
                    frame.FromSheetX(minX), frame.FromSheetY(minY),
                    frame.FromSheetX(maxX), frame.FromSheetY(maxY)))
            Else
                log.Repos(String.Format(CultureInfo.InvariantCulture,
                    "{0} Range hoja=({1:0.######},{2:0.######})-({3:0.######},{4:0.######})",
                    tag, minX, minY, maxX, maxY))
            End If
        Catch ex As Exception
            log.Repos(tag & " Range: no legible (" & ex.Message & ")")
        End Try
    End Sub

    Private Shared Sub TryReadTrackDistances(d As FrameworkDimension, log As DimensionLogger)
        If d Is Nothing OrElse log Is Nothing Then Return
        Try
            Dim td As Double = d.TrackDistance
            log.Repos("TrackDistance (lectura)=" & FormatDouble(td))
        Catch ex As Exception
            log.Repos("TrackDistance: no legible — " & FormatEx(ex))
        End Try
        Try
            Dim absTd As Double = d.AbsoluteTrackDistance
            log.Repos("AbsoluteTrackDistance (lectura)=" & FormatDouble(absTd))
        Catch ex As Exception
            log.Repos("AbsoluteTrackDistance: no legible — " & ex.Message)
        End Try
    End Sub

    Private Shared Sub LogWritableLocationPropertyDiagnostics(d As FrameworkDimension, log As DimensionLogger)
        If d Is Nothing OrElse log Is Nothing Then Return
        log.Repos("propiedades escritura candidatas (reflexión, para investigación futura):")
        Try
            Dim t As Type = d.GetType()
            For Each pi As PropertyInfo In t.GetProperties(BindingFlags.Public Or BindingFlags.Instance)
                If Not pi.CanWrite Then Continue For
                Dim n As String = pi.Name
                If n.IndexOf("Text", StringComparison.OrdinalIgnoreCase) < 0 AndAlso
                   n.IndexOf("Track", StringComparison.OrdinalIgnoreCase) < 0 AndAlso
                   n.IndexOf("Key", StringComparison.OrdinalIgnoreCase) < 0 AndAlso
                   n.IndexOf("Position", StringComparison.OrdinalIgnoreCase) < 0 AndAlso
                   n.IndexOf("Offset", StringComparison.OrdinalIgnoreCase) < 0 AndAlso
                   n.IndexOf("Distance", StringComparison.OrdinalIgnoreCase) < 0 AndAlso
                   n.IndexOf("Axis", StringComparison.OrdinalIgnoreCase) < 0 AndAlso
                   n.IndexOf("Projection", StringComparison.OrdinalIgnoreCase) < 0 AndAlso
                   n.IndexOf("Leader", StringComparison.OrdinalIgnoreCase) < 0 AndAlso
                   n.IndexOf("Terminator", StringComparison.OrdinalIgnoreCase) < 0 Then
                    Continue For
                End If
                log.Repos("  writable: " & n & " As " & pi.PropertyType.Name)
            Next
        Catch ex As Exception
            log.Repos("reflexión EX: " & ex.Message)
        End Try
    End Sub

    Private Shared Function FormatEx(ex As Exception) As String
        If ex Is Nothing Then Return ""
        Dim s As String = ex.GetType().Name & ": " & ex.Message
        Dim cex As COMException = TryCast(ex, COMException)
        If cex IsNot Nothing Then
            s &= " | HRESULT=0x" & cex.ErrorCode.ToString("X8", CultureInfo.InvariantCulture)
        End If
        Return s
    End Function

    Private Shared Function FormatDouble(v As Double) As String
        If Double.IsNaN(v) Then Return "NaN"
        Return v.ToString("0.######", CultureInfo.InvariantCulture)
    End Function

End Class
