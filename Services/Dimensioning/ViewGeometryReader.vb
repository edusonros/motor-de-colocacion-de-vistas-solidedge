Option Strict Off

Imports System.Collections.Generic
Imports System.Globalization
Imports System.Linq
Imports SolidEdgeDraft

''' <summary>Límite de la vista en coordenadas de hoja (metros), desde DrawingView.Range.</summary>
Public Structure ViewSheetBoundingBox
    Public MinX As Double
    Public MinY As Double
    Public MaxX As Double
    Public MaxY As Double

    Public ReadOnly Property Width As Double
        Get
            Return Math.Abs(MaxX - MinX)
        End Get
    End Property

    Public ReadOnly Property Height As Double
        Get
            Return Math.Abs(MaxY - MinY)
        End Get
    End Property
End Structure

''' <summary>Una arista 2D de vista lista para acotar (referencia COM + puntos en hoja).</summary>
Friend Class DvLineSheetInfo
    Public Line As DVLine2d
    ''' <summary>Índice 1-based en DVLines2d.Item para el log.</summary>
    Public SourceIndex As Integer
    Public Sx1 As Double
    Public Sy1 As Double
    Public Sx2 As Double
    Public Sy2 As Double
    Public Length As Double
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

''' <summary>Contadores de la cosecha de DVLines2d para el log de acotado.</summary>
Friend Class ExtremeLineBuildStats
    ''' <summary>Ítems en <see cref="DVLines2d"/> (1..Count).</summary>
    Public TotalDvLines2d As Integer
    ''' <summary>Longitud en hoja &lt; umbral (5% del ancho de vista).</summary>
    Public DiscardedTooShort As Integer
    ''' <summary>COM inválido, arista no modelo en modo estricto, etc.</summary>
    Public DiscardedOther As Integer
    Public RelevanceThresholdM As Double
    ''' <summary>Tolerancia |dx|/|dy| en m (hoja) para clasificar vertical/horizontal.</summary>
    Public AxisToleranceM As Double
    Public UsedLooseEdgePass As Boolean
End Class

''' <summary>Líneas extremas elegidas para cotas total ancho/alto.</summary>
Friend Class ExtremeDvLinesResult
    Public AllLines As List(Of DvLineSheetInfo)
    Public LeftVertical As DvLineSheetInfo
    Public RightVertical As DvLineSheetInfo
    Public BottomHorizontal As DvLineSheetInfo
    Public TopHorizontal As DvLineSheetInfo
    Public Stats As ExtremeLineBuildStats
End Class

''' <summary>Lectura de Range y de DVLines2d para referencias COM válidas en cotas.</summary>
Friend NotInheritable Class ViewGeometryReader

    ''' <summary>Longitud mínima = 5% del ancho de la vista (hoja), para ignorar detalles locales al elegir extremos.</summary>
    Private Const RelevanceWidthFraction As Double = 0.05R
    ''' <summary>Tolerancia de alineación a ejes: |Δx| o |Δy| en hoja (metros).</summary>
    Private Const AxisToleranceWidthFraction As Double = 0.001R

    Private Sub New()
    End Sub

    Public Shared Function TryReadBoundingBox(view As DrawingView, log As DimensionLogger, ByRef box As ViewSheetBoundingBox) As Boolean
        box = New ViewSheetBoundingBox()
        If view Is Nothing Then
            log?.Warn("TryReadBoundingBox: vista Nothing.")
            Return False
        End If

        Try
            view.Update()
        Catch ex As Exception
            log?.ComFail("DrawingView.Update", "DrawingView", ex)
        End Try

        Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
        Try
            view.Range(x1, y1, x2, y2)
        Catch ex As Exception
            log?.ComFail("DrawingView.Range", "DrawingView", ex)
            Return False
        End Try

        box.MinX = Math.Min(x1, x2)
        box.MinY = Math.Min(y1, y2)
        box.MaxX = Math.Max(x1, x2)
        box.MaxY = Math.Max(y1, y2)

        If box.Width <= 1.0E-9 OrElse box.Height <= 1.0E-9 Then
            log?.Warn("Rango de vista degenerado (ancho o alto ~0).")
            Return False
        End If

        Return True
    End Function

    ''' <summary>Lee DVLines2d y elige líneas extremas para cotas horizontal/vertical totales.</summary>
    Public Shared Function TryBuildExtremeLines(view As DrawingView, box As ViewSheetBoundingBox, log As DimensionLogger, ByRef result As ExtremeDvLinesResult) As Boolean
        result = Nothing
        If view Is Nothing OrElse Not box.Width > 0 Then Return False

        Dim relevanceThreshold As Double = Math.Max(box.Width * RelevanceWidthFraction, 1.0E-9)
        Dim axisTol As Double = Math.Max(box.Width * AxisToleranceWidthFraction, 1.0E-7)

        Dim strictList As New List(Of DvLineSheetInfo)()
        Dim stStrict As New ExtremeLineBuildStats With {
            .RelevanceThresholdM = relevanceThreshold,
            .AxisToleranceM = axisTol,
            .UsedLooseEdgePass = False
        }
        CollectInner(view, box, log, strictList, relevanceThreshold, modelEdgesOnly:=True, stStrict)

        Dim finalList As List(Of DvLineSheetInfo) = strictList
        Dim finalStats As ExtremeLineBuildStats = stStrict

        If strictList.Count < 4 Then
            Dim looseList As New List(Of DvLineSheetInfo)()
            Dim stLoose As New ExtremeLineBuildStats With {
                .RelevanceThresholdM = relevanceThreshold,
                .AxisToleranceM = axisTol,
                .UsedLooseEdgePass = True
            }
            CollectInner(view, box, log, looseList, relevanceThreshold, modelEdgesOnly:=False, stLoose)

            If looseList.Count > strictList.Count Then
                If strictList.Count < 4 Then
                    log?.Warn("Pocas aristas de modelo; se usan todas las DVLine2d visibles para acotar.")
                End If
                finalList = looseList
                finalStats = stLoose
            End If
        End If

        If finalList Is Nothing OrElse finalList.Count = 0 Then
            log?.Err("DVLines2d: no se obtuvo ninguna línea utilizable (todas bajo umbral o no válidas).")
            Return False
        End If

        Dim res As New ExtremeDvLinesResult With {.AllLines = finalList, .Stats = finalStats}
        PickExtremeLines(finalList, res, finalStats.AxisToleranceM)

        Dim nH As Integer = finalList.Where(Function(L) IsHorizontalForExtremes(L, finalStats.AxisToleranceM)).Count()
        Dim nV As Integer = finalList.Where(Function(L) IsVerticalForExtremes(L, finalStats.AxisToleranceM)).Count()
        log?.Info(String.Format(CultureInfo.InvariantCulture,
            "Filtrado DVLines2d: totales={0}, umbral longitud={1:0.######}m (5% ancho={2:0.######}m), descartadas cortas={3}, otras descartadas={4}, válidas={5}, tolerancia eje={6:0.######}m, horiz={7}, vert={8}, paso suelto={9}",
            finalStats.TotalDvLines2d,
            finalStats.RelevanceThresholdM,
            box.Width,
            finalStats.DiscardedTooShort,
            finalStats.DiscardedOther,
            finalList.Count,
            finalStats.AxisToleranceM,
            nH,
            nV,
            finalStats.UsedLooseEdgePass.ToString()))

        log?.Info("Extremos elegidos — izq: " & FormatLineLog(res.LeftVertical) &
                  " | der: " & FormatLineLog(res.RightVertical) &
                  " | inf: " & FormatLineLog(res.BottomHorizontal) &
                  " | sup: " & FormatLineLog(res.TopHorizontal))

        result = res
        Return True
    End Function

    Private Shared Sub CollectInner(view As DrawingView, box As ViewSheetBoundingBox, log As DimensionLogger,
                                    outList As List(Of DvLineSheetInfo), relevanceThreshold As Double, modelEdgesOnly As Boolean,
                                    stats As ExtremeLineBuildStats)
        Dim linesCol As DVLines2d = Nothing
        Try
            linesCol = view.DVLines2d
        Catch ex As Exception
            log?.ComFail("DrawingView.DVLines2d", "DrawingView", ex)
            Return
        End Try
        If linesCol Is Nothing Then Return

        Dim n As Integer = 0
        Try
            n = linesCol.Count
        Catch ex As Exception
            log?.ComFail("DVLines2d.Count", "DVLines2d", ex)
            Return
        End Try

        If stats IsNot Nothing Then stats.TotalDvLines2d = n

        Dim modelEdge As Integer = CInt(SolidEdgeConstants.GraphicMemberEdgeTypeConstants.seModelEdgeType)

        For i As Integer = 1 To n
            Dim ln As DVLine2d = Nothing
            Try
                ln = CType(linesCol.Item(i), DVLine2d)
            Catch ex As Exception
                log?.Warn("Línea descartada por no ser válida para acotar (Item " & i.ToString(CultureInfo.InvariantCulture) & "): " & ex.Message)
                If stats IsNot Nothing Then stats.DiscardedOther += 1
                Continue For
            End Try
            If ln Is Nothing Then
                log?.Warn("Línea descartada por no ser válida para acotar (Item " & i.ToString(CultureInfo.InvariantCulture) & "): referencia Nothing")
                If stats IsNot Nothing Then stats.DiscardedOther += 1
                Continue For
            End If

            If modelEdgesOnly Then
                Dim et As Integer = -1
                Try
                    et = CInt(ln.EdgeType)
                Catch
                    et = -1
                End Try
                If et >= 0 AndAlso et <> modelEdge Then
                    If stats IsNot Nothing Then stats.DiscardedOther += 1
                    Continue For
                End If
            End If

            Dim vx1 As Double, vy1 As Double, vx2 As Double, vy2 As Double
            Try
                ln.GetStartPoint(vx1, vy1)
                ln.GetEndPoint(vx2, vy2)
            Catch ex As Exception
                log?.Warn("Línea descartada por no ser válida para acotar (Item " & i.ToString(CultureInfo.InvariantCulture) & " GetStart/EndPoint): " & ex.Message)
                If stats IsNot Nothing Then stats.DiscardedOther += 1
                Continue For
            End Try

            Dim sx1 As Double, sy1 As Double, sx2 As Double, sy2 As Double
            Try
                view.ViewToSheet(vx1, vy1, sx1, sy1)
                view.ViewToSheet(vx2, vy2, sx2, sy2)
            Catch ex As Exception
                log?.Warn("Línea descartada por no ser válida para acotar (Item " & i.ToString(CultureInfo.InvariantCulture) & " ViewToSheet): " & ex.Message)
                If stats IsNot Nothing Then stats.DiscardedOther += 1
                Continue For
            End Try

            If Not AreFinite(sx1, sy1, sx2, sy2) Then
                log?.Warn("Línea descartada por no ser válida para acotar (Item " & i.ToString(CultureInfo.InvariantCulture) & "): coordenadas no finitas")
                If stats IsNot Nothing Then stats.DiscardedOther += 1
                Continue For
            End If

            Dim dx As Double = sx2 - sx1
            Dim dy As Double = sy2 - sy1
            Dim len As Double = Math.Sqrt(dx * dx + dy * dy)
            If len < relevanceThreshold Then
                If stats IsNot Nothing Then stats.DiscardedTooShort += 1
                Continue For
            End If

            outList.Add(New DvLineSheetInfo With {
                .Line = ln,
                .SourceIndex = i,
                .Sx1 = sx1, .Sy1 = sy1, .Sx2 = sx2, .Sy2 = sy2,
                .Length = len
            })
        Next
    End Sub

    Private Shared Sub PickExtremeLines(lines As List(Of DvLineSheetInfo), res As ExtremeDvLinesResult, axisTol As Double)
        Dim vert = lines.Where(Function(L) IsVerticalForExtremes(L, axisTol)).ToList()
        Dim horiz = lines.Where(Function(L) IsHorizontalForExtremes(L, axisTol)).ToList()

        If vert.Count >= 2 Then
            res.LeftVertical = vert.OrderBy(Function(L) L.MinXs).First()
            res.RightVertical = vert.OrderByDescending(Function(L) L.MaxXs).First()
            If ReferenceEquals(res.LeftVertical, res.RightVertical) Then
                res.RightVertical = vert.Where(Function(L) Not ReferenceEquals(L, res.LeftVertical)).OrderByDescending(Function(L) L.MaxXs).FirstOrDefault()
            End If
        End If

        If horiz.Count >= 2 Then
            res.BottomHorizontal = horiz.OrderBy(Function(L) L.MidY).First()
            res.TopHorizontal = horiz.OrderByDescending(Function(L) L.MidY).First()
            If ReferenceEquals(res.BottomHorizontal, res.TopHorizontal) Then
                res.TopHorizontal = horiz.Where(Function(L) Not ReferenceEquals(L, res.BottomHorizontal)).OrderByDescending(Function(L) L.MidY).FirstOrDefault()
            End If
        End If

        If res.LeftVertical Is Nothing OrElse res.RightVertical Is Nothing Then
            Dim byX = lines.OrderBy(Function(L) L.MidX).ToList()
            If byX.Count >= 2 Then
                res.LeftVertical = byX.First()
                res.RightVertical = byX.Last()
                If ReferenceEquals(res.LeftVertical, res.RightVertical) AndAlso byX.Count > 2 Then
                    res.RightVertical = byX(byX.Count - 2)
                End If
            End If
        End If

        If res.BottomHorizontal Is Nothing OrElse res.TopHorizontal Is Nothing Then
            Dim byY = lines.OrderBy(Function(L) L.MidY).ToList()
            If byY.Count >= 2 Then
                res.BottomHorizontal = byY.First()
                res.TopHorizontal = byY.Last()
                If ReferenceEquals(res.BottomHorizontal, res.TopHorizontal) AndAlso byY.Count > 2 Then
                    res.TopHorizontal = byY(byY.Count - 2)
                End If
            End If
        End If
    End Sub

    ''' <summary>Vertical si |Δx| en hoja ≤ tolerancia (metros).</summary>
    Private Shared Function IsVerticalForExtremes(L As DvLineSheetInfo, axisTol As Double) As Boolean
        If L Is Nothing OrElse L.Length < 1.0E-12 Then Return False
        Return Math.Abs(L.Sx2 - L.Sx1) <= axisTol
    End Function

    ''' <summary>Horizontal si |Δy| en hoja ≤ tolerancia (metros).</summary>
    Private Shared Function IsHorizontalForExtremes(L As DvLineSheetInfo, axisTol As Double) As Boolean
        If L Is Nothing OrElse L.Length < 1.0E-12 Then Return False
        Return Math.Abs(L.Sy2 - L.Sy1) <= axisTol
    End Function

    Private Shared Function AreFinite(x1 As Double, y1 As Double, x2 As Double, y2 As Double) As Boolean
        Return Not (Double.IsNaN(x1) OrElse Double.IsNaN(y1) OrElse Double.IsNaN(x2) OrElse Double.IsNaN(y2) OrElse
                    Double.IsInfinity(x1) OrElse Double.IsInfinity(y1) OrElse Double.IsInfinity(x2) OrElse Double.IsInfinity(y2))
    End Function

    Public Shared Function FormatLineLog(L As DvLineSheetInfo) As String
        If L Is Nothing Then Return "(ninguna)"
        Return String.Format(CultureInfo.InvariantCulture,
            "DVLine2d índice={0} L={1:0.####}m hoja ({2:0.####},{3:0.####})-({4:0.####},{5:0.####})",
            L.SourceIndex, L.Length, L.Sx1, L.Sy1, L.Sx2, L.Sy2)
    End Function

End Class
