Option Strict Off

Imports System.Collections.Generic
Imports System.Globalization
Imports SolidEdgeDraft
Imports SolidEdgeConstants
Imports SolidEdgeFrameworkSupport

Namespace Extraer_dft_dxf_flatdxf.LegacyV02Dimensioning

''' <summary>Elimina cotas con el mismo valor nominal y los mismos keypoints (tolerancia en m).</summary>
Friend NotInheritable Class DimensionDuplicateCleanup

    Private Const ValueToleranceMm As Double = 0.05R
    Private Const KeyPointToleranceM As Double = 0.001R

    Private Sub New()
    End Sub

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
                    log?.LogLine("[DIM][DEDUP][DELETE] idx=" & i.ToString(CultureInfo.InvariantCulture) &
                                 " sig=" & sig & " keep_idx=" & seen(sig).ToString(CultureInfo.InvariantCulture))
                Catch ex As Exception
                    log?.LogLine("[DIM][DEDUP][DELETE][WARN] idx=" & i.ToString(CultureInfo.InvariantCulture) & " " & ex.Message)
                End Try
            Else
                seen(sig) = i
            End If
        Next

        If removed > 0 AndAlso log IsNot Nothing Then
            log.LogLine("[DIM][DEDUP][SUMMARY] removed=" & removed.ToString(CultureInfo.InvariantCulture) &
                        " kept=" & seen.Count.ToString(CultureInfo.InvariantCulture))
        End If
        Return removed
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

End Namespace
