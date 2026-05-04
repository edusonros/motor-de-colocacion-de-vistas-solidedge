Option Strict Off

Imports System.Globalization

Friend Enum PlannedDimensionKind
    HorizontalTotal
    VerticalTotal
    Radial
    Diameter
    LineLength
End Enum

Friend NotInheritable Class PlannedDimension
    Public Property Kind As PlannedDimensionKind
    Public Property TypeTag As String
    Public Property Obj1 As Object
    Public Property Obj2 As Object
    Public Property P1X As Double
    Public Property P1Y As Double
    Public Property P2X As Double
    Public Property P2Y As Double
    Public Property Signature As String
End Class

Friend NotInheritable Class DrawingViewDimensionPlanner
    Private Sub New()
    End Sub

    Public Shared Function Plan(info As DrawingViewGeometryInfo, log As DimensionLogger) As List(Of PlannedDimension)
        Dim out As New List(Of PlannedDimension)()
        If info Is Nothing Then Return out
        Dim minUsefulLen As Double = Math.Max(Math.Max(info.Box.Width, info.Box.Height) * 0.1R, 0.01R)

        Dim ex As ExtremeDvLinesResult = info.Extreme
        If ex IsNot Nothing AndAlso ex.LeftVertical IsNot Nothing AndAlso ex.RightVertical IsNot Nothing Then
            Dim yRef As Double = Math.Max(ex.LeftVertical.MidY, ex.RightVertical.MidY) + Math.Max(info.Box.Height * 0.08R, 0.004R)
            Dim h As New PlannedDimension With {
                .Kind = PlannedDimensionKind.HorizontalTotal,
                .TypeTag = "H_TOTAL",
                .Obj1 = ex.LeftVertical.Line,
                .Obj2 = ex.RightVertical.Line,
                .P1X = ex.LeftVertical.MidX,
                .P1Y = yRef,
                .P2X = ex.RightVertical.MidX,
                .P2Y = yRef,
                .Signature = "H_TOTAL|" & info.ViewIndex.ToString(CultureInfo.InvariantCulture)
            }
            out.Add(h)
            log?.LogLine("[DIM][PLAN][H_TOTAL] view=" & info.ViewIndex.ToString(CultureInfo.InvariantCulture))
        End If

        If ex IsNot Nothing AndAlso ex.BottomHorizontal IsNot Nothing AndAlso ex.TopHorizontal IsNot Nothing Then
            Dim xRef As Double = Math.Max(ex.BottomHorizontal.MidX, ex.TopHorizontal.MidX) + Math.Max(info.Box.Width * 0.08R, 0.004R)
            Dim v As New PlannedDimension With {
                .Kind = PlannedDimensionKind.VerticalTotal,
                .TypeTag = "V_TOTAL",
                .Obj1 = ex.BottomHorizontal.Line,
                .Obj2 = ex.TopHorizontal.Line,
                .P1X = xRef,
                .P1Y = ex.BottomHorizontal.MidY,
                .P2X = xRef,
                .P2Y = ex.TopHorizontal.MidY,
                .Signature = "V_TOTAL|" & info.ViewIndex.ToString(CultureInfo.InvariantCulture)
            }
            out.Add(v)
            log?.LogLine("[DIM][PLAN][V_TOTAL] view=" & info.ViewIndex.ToString(CultureInfo.InvariantCulture))
        End If

        ' Cotas extra interiores sobre líneas reales de la vista (AddLength sobre DVLine2d).
        If ex IsNot Nothing AndAlso ex.TopHorizontal IsNot Nothing AndAlso ex.TopHorizontal.Line IsNot Nothing AndAlso ex.TopHorizontal.Length >= minUsefulLen Then
            Dim l1 As New PlannedDimension With {
                .Kind = PlannedDimensionKind.LineLength,
                .TypeTag = "LENGTH_TOP",
                .Obj1 = ex.TopHorizontal.Line,
                .Signature = "LENGTH_TOP|" & info.ViewIndex.ToString(CultureInfo.InvariantCulture)
            }
            out.Add(l1)
            log?.LogLine("[DIM][PLAN][LENGTH] view=" & info.ViewIndex.ToString(CultureInfo.InvariantCulture) & " target=top_horizontal")
        End If
        If ex IsNot Nothing AndAlso ex.LeftVertical IsNot Nothing AndAlso ex.LeftVertical.Line IsNot Nothing AndAlso ex.LeftVertical.Length >= minUsefulLen Then
            Dim l2 As New PlannedDimension With {
                .Kind = PlannedDimensionKind.LineLength,
                .TypeTag = "LENGTH_LEFT",
                .Obj1 = ex.LeftVertical.Line,
                .Signature = "LENGTH_LEFT|" & info.ViewIndex.ToString(CultureInfo.InvariantCulture)
            }
            out.Add(l2)
            log?.LogLine("[DIM][PLAN][LENGTH] view=" & info.ViewIndex.ToString(CultureInfo.InvariantCulture) & " target=left_vertical")
        End If

        ' Temporal: radial/diámetro desactivado hasta cerrar firma API estable en esta versión COM.
        If info.FirstCircle IsNot Nothing OrElse info.FirstArc IsNot Nothing Then
            log?.LogLine("[DIM][PLAN][RADIAL] view=" & info.ViewIndex.ToString(CultureInfo.InvariantCulture) & " skip=api_signature_pending")
        End If

        Return out
    End Function
End Class
