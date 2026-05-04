Option Strict Off

Imports System.Globalization
Imports SolidEdgeDraft

Friend NotInheritable Class DrawingViewGeometryInfo
    Public Property View As DrawingView
    Public Property ViewIndex As Integer
    Public Property ViewName As String
    Public Property Box As ViewSheetBoundingBox
    Public Property Extreme As ExtremeDvLinesResult
    Public Property CountLines As Integer
    Public Property CountArcs As Integer
    Public Property CountCircles As Integer
    Public Property CountLineStrings As Integer
    Public Property CountSplines As Integer
    Public Property CountPoints As Integer
    Public Property FirstArc As Object
    Public Property FirstCircle As Object
End Class

Friend NotInheritable Class DrawingViewGeometryReader
    Private Sub New()
    End Sub

    Public Shared Function Read(view As DrawingView, viewIndex As Integer, log As DimensionLogger) As DrawingViewGeometryInfo
        If view Is Nothing Then Return Nothing

        Dim box As New ViewSheetBoundingBox()
        If Not ViewGeometryReader.TryReadBoundingBox(view, log, box) Then
            log?.LogLine("[DIM][VIEW][SKIP] idx=" & viewIndex.ToString(CultureInfo.InvariantCulture) & " reason=no_bbox")
            Return Nothing
        End If

        Dim exLines As ExtremeDvLinesResult = Nothing
        ViewGeometryReader.TryBuildExtremeLines(view, box, log, exLines)

        Dim info As New DrawingViewGeometryInfo With {
            .View = view,
            .ViewIndex = viewIndex,
            .ViewName = SafeToString(CallByNameSafe(view, "Name")),
            .Box = box,
            .Extreme = exLines,
            .CountLines = SafeCount(CallByNameSafe(view, "DVLines2d")),
            .CountArcs = SafeCount(CallByNameSafe(view, "DVArcs2d")),
            .CountCircles = SafeCount(CallByNameSafe(view, "DVCircles2d")),
            .CountLineStrings = SafeCount(CallByNameSafe(view, "DVLineStrings2d")),
            .CountSplines = SafeCount(CallByNameSafe(view, "DVBSplineCurves2d")),
            .CountPoints = SafeCount(CallByNameSafe(view, "DVPoints2d")),
            .FirstArc = SafeFirst(CallByNameSafe(view, "DVArcs2d")),
            .FirstCircle = SafeFirst(CallByNameSafe(view, "DVCircles2d"))
        }

        log?.LogLine(String.Format(CultureInfo.InvariantCulture,
            "[DIM][GEOM][COUNT] view={0} lines={1} arcs={2} circles={3} linestrings={4} splines={5} points={6}",
            info.ViewIndex, info.CountLines, info.CountArcs, info.CountCircles, info.CountLineStrings, info.CountSplines, info.CountPoints))

        Return info
    End Function

    Private Shared Function CallByNameSafe(obj As Object, member As String) As Object
        Try
            Return CallByName(obj, member, CallType.Get)
        Catch
            Return Nothing
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

    Private Shared Function SafeFirst(col As Object) As Object
        If col Is Nothing Then Return Nothing
        Try
            If CInt(CallByName(col, "Count", CallType.Get)) <= 0 Then Return Nothing
            Return CallByName(col, "Item", CallType.Method, 1)
        Catch
            Return Nothing
        End Try
    End Function

    Private Shared Function SafeToString(v As Object) As String
        If v Is Nothing Then Return ""
        Try
            Return Convert.ToString(v, CultureInfo.InvariantCulture)
        Catch
            Return ""
        End Try
    End Function
End Class
