Option Strict Off

Imports System.Collections.Generic
Imports System.Globalization
Imports System.Linq
Imports SolidEdgeDraft

Public Enum DvLineOrientationKind
    Horizontal
    Vertical
    Oblique
End Enum

''' <summary>Segmento DVLine2d en hoja (sin GetKeyPoint; solo GetStartPoint/EndPoint + ViewToSheet).</summary>
Public Class DvLineGeomInfo
    Public Property Line As DVLine2d
    Public Property SourceIndex As Integer
    Public Property Sx1 As Double
    Public Property Sy1 As Double
    Public Property Sx2 As Double
    Public Property Sy2 As Double
    Public Property Length As Double
    Public Property Orientation As DvLineOrientationKind
    Public Property IsHiddenOrNonModel As Boolean

    Public ReadOnly Property MidX As Double
        Get
            Return (Sx1 + Sx2) / 2.0
        End Get
    End Property

    Public ReadOnly Property MidY As Double
        Get
            Return (Sy1 + Sy2) / 2.0
        End Get
    End Property

    Public ReadOnly Property MinXs As Double
        Get
            Return Math.Min(Sx1, Sx2)
        End Get
    End Property

    Public ReadOnly Property MaxXs As Double
        Get
            Return Math.Max(Sx1, Sx2)
        End Get
    End Property

    Public ReadOnly Property MinYs As Double
        Get
            Return Math.Min(Sy1, Sy2)
        End Get
    End Property

    Public ReadOnly Property MaxYs As Double
        Get
            Return Math.Max(Sy1, Sy2)
        End Get
    End Property
End Class

Public Class ViewGeometryInfo
    Public Property View As DrawingView
    Public Property Lines As New List(Of DvLineGeomInfo)()
    Public Property CountArcs As Integer
    Public Property CountCircles As Integer
    Public Property CountLineStrings As Integer
    Public Property CountSplines As Integer
    Public Property CountPoints As Integer
End Class

''' <summary>Lectura de geometría 2D de vista para ISO 129 (DV*, sin 2D Model / Drop / GetKeyPoint).</summary>
Public NotInheritable Class DimensionGeometryReader
    Private Const RelevanceWidthFraction As Double = 0.05R
    Private Const AxisToleranceWidthFraction As Double = 0.001R

    Private Sub New()
    End Sub

    Public Shared Function LeerGeometriaDV(view As DrawingView, config As DimensioningNormConfig, log As Action(Of String)) As ViewGeometryInfo
        Dim info As New ViewGeometryInfo With {.View = view}
        If view Is Nothing Then Return info

        Try
            view.Update()
        Catch ex As Exception
            log?.Invoke("[DIM][ISO129][GEOM] view.Update ex=" & ex.Message)
        End Try

        Dim boxW As Double = 0.01
        Try
            Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
            view.Range(x1, y1, x2, y2)
            boxW = Math.Max(Math.Abs(x2 - x1), Math.Abs(y2 - y1))
        Catch
            boxW = 0.01
        End Try

        Dim relevanceThreshold As Double = Math.Max(boxW * RelevanceWidthFraction, 1.0E-9)
        Dim axisTol As Double = Math.Max(boxW * AxisToleranceWidthFraction, 1.0E-7)

        CollectLines(view, info.Lines, relevanceThreshold, axisTol, config, strictModel:=True, log)
        If info.Lines.Count < 4 Then
            Dim loose As New List(Of DvLineGeomInfo)()
            CollectLines(view, loose, relevanceThreshold, axisTol, config, strictModel:=False, log)
            If loose.Count > info.Lines.Count Then
                log?.Invoke("[DIM][ISO129][GEOM] Pocas aristas modelo; segunda pasada incluye no-modelo.")
                info.Lines = loose
            End If
        End If

        info.CountArcs = SafeCountCollection(view, "DVArcs2d", log)
        info.CountCircles = SafeCountCollection(view, "DVCircles2d", log)
        info.CountLineStrings = SafeCountCollection(view, "DVLineStrings2d", log)
        info.CountSplines = SafeCountCollection(view, "DVBSplineCurves2d", log)
        info.CountPoints = SafeCountCollection(view, "DVPoints2d", log)

        Dim vname As String = SafeViewName(view)
        log?.Invoke(String.Format(CultureInfo.InvariantCulture,
            "[DIM][ISO129][GEOM] view={0} lines={1} arcs={2} circles={3} linestrings={4} splines={5} points={6}",
            vname, info.Lines.Count, info.CountArcs, info.CountCircles, info.CountLineStrings, info.CountSplines, info.CountPoints))

        Return info
    End Function

    Public Shared Function CalcularBoundingBoxGeometria(geom As ViewGeometryInfo) As BoundingBox2D
        Dim b As New BoundingBox2D With {
            .MinX = Double.PositiveInfinity,
            .MinY = Double.PositiveInfinity,
            .MaxX = Double.NegativeInfinity,
            .MaxY = Double.NegativeInfinity
        }
        If geom Is Nothing OrElse geom.Lines Is Nothing OrElse geom.Lines.Count = 0 Then
            b.MinX = 0 : b.MinY = 0 : b.MaxX = 0 : b.MaxY = 0
            Return b
        End If
        For Each L In geom.Lines
            If L Is Nothing Then Continue For
            b.MinX = Math.Min(b.MinX, Math.Min(L.Sx1, L.Sx2))
            b.MaxX = Math.Max(b.MaxX, Math.Max(L.Sx1, L.Sx2))
            b.MinY = Math.Min(b.MinY, Math.Min(L.Sy1, L.Sy2))
            b.MaxY = Math.Max(b.MaxY, Math.Max(L.Sy1, L.Sy2))
        Next
        If Double.IsInfinity(b.MinX) Then b.MinX = 0 : b.MaxX = 0 : b.MinY = 0 : b.MaxY = 0
        Return b
    End Function

    Private Shared Sub CollectLines(
        view As DrawingView,
        outList As List(Of DvLineGeomInfo),
        relevanceThreshold As Double,
        axisTol As Double,
        config As DimensioningNormConfig,
        strictModel As Boolean,
        log As Action(Of String))

        outList.Clear()
        Dim linesCol As DVLines2d = Nothing
        Try
            linesCol = view.DVLines2d
        Catch ex As Exception
            log?.Invoke("[DIM][ISO129][GEOM] DVLines2d no disponible: " & ex.Message)
            Return
        End Try
        If linesCol Is Nothing Then Return

        Dim n As Integer = 0
        Try
            n = linesCol.Count
        Catch ex As Exception
            log?.Invoke("[DIM][ISO129][GEOM] DVLines2d.Count: " & ex.Message)
            Return
        End Try

        Dim modelEdge As Integer = CInt(SolidEdgeConstants.GraphicMemberEdgeTypeConstants.seModelEdgeType)

        For i As Integer = 1 To n
            Dim ln As DVLine2d = Nothing
            Try
                ln = CType(linesCol.Item(i), DVLine2d)
            Catch
                Continue For
            End Try
            If ln Is Nothing Then Continue For

            Dim et As Integer = -1
            Try
                et = CInt(ln.EdgeType)
            Catch
                et = -1
            End Try

            Dim isHidden As Boolean = False
            If strictModel Then
                If et >= 0 AndAlso et <> modelEdge Then Continue For
            Else
                isHidden = (et >= 0 AndAlso et <> modelEdge)
            End If

            If config IsNot Nothing AndAlso config.AvoidHiddenGeometry AndAlso isHidden Then Continue For

            Dim vx1 As Double, vy1 As Double, vx2 As Double, vy2 As Double
            Try
                ln.GetStartPoint(vx1, vy1)
                ln.GetEndPoint(vx2, vy2)
            Catch
                Continue For
            End Try

            Dim sx1 As Double, sy1 As Double, sx2 As Double, sy2 As Double
            Try
                view.ViewToSheet(vx1, vy1, sx1, sy1)
                view.ViewToSheet(vx2, vy2, sx2, sy2)
            Catch
                Continue For
            End Try

            If Not AreFinite(sx1, sy1, sx2, sy2) Then Continue For

            Dim dx As Double = sx2 - sx1
            Dim dy As Double = sy2 - sy1
            Dim len As Double = Math.Sqrt(dx * dx + dy * dy)
            If len < relevanceThreshold Then Continue For

            Dim ori As DvLineOrientationKind = DvLineOrientationKind.Oblique
            If Math.Abs(dy) <= axisTol Then ori = DvLineOrientationKind.Horizontal
            If Math.Abs(dx) <= axisTol Then ori = DvLineOrientationKind.Vertical

            outList.Add(New DvLineGeomInfo With {
                .Line = ln,
                .SourceIndex = i,
                .Sx1 = sx1, .Sy1 = sy1, .Sx2 = sx2, .Sy2 = sy2,
                .Length = len,
                .Orientation = ori,
                .IsHiddenOrNonModel = isHidden
            })
        Next
    End Sub

    Private Shared Function SafeCountCollection(view As DrawingView, member As String, log As Action(Of String)) As Integer
        If view Is Nothing Then Return 0
        Try
            Dim col As Object = CallByName(view, member, CallType.Get)
            If col Is Nothing Then Return 0
            Return CInt(CallByName(col, "Count", CallType.Get))
        Catch ex As Exception
            log?.Invoke("[DIM][ISO129][GEOM] " & member & " omitido: " & ex.Message)
            Return 0
        End Try
    End Function

    Private Shared Function AreFinite(x1 As Double, y1 As Double, x2 As Double, y2 As Double) As Boolean
        Return Not (Double.IsNaN(x1) OrElse Double.IsNaN(y1) OrElse Double.IsNaN(x2) OrElse Double.IsNaN(y2) OrElse
                    Double.IsInfinity(x1) OrElse Double.IsInfinity(y1) OrElse Double.IsInfinity(x2) OrElse Double.IsInfinity(y2))
    End Function

    Private Shared Function SafeViewName(view As DrawingView) As String
        If view Is Nothing Then Return "?"
        Try
            Return Convert.ToString(CallByName(view, "Name", CallType.Get), CultureInfo.InvariantCulture)
        Catch
            Return "?"
        End Try
    End Function
End Class
