Option Strict Off

Imports System.Collections.Generic
Imports System.Globalization
Imports System.Linq
Imports SolidEdgeDraft

''' <summary>
''' Recorre la geometría 2D visible de la vista base (líneas, arcos, círculos, elipses) en coordenadas de hoja,
''' genera puntos candidatos reales (extremos de segmentos, arco, muestras) y selecciona
''' izquierda/derecha/abajo/arriba por min/max X/Y. No sustituye el bbox global por geometría: los extremos provienen de puntos concretos.
''' </summary>
Friend NotInheritable Class RealExtremePointsResolver

    Private Const RelevanceWidthFraction As Double = 0.05R

    Private Sub New()
    End Sub

    ''' <summary>
    ''' Resuelve los cuatro extremos geométricos reales a partir de candidatos de vista base.
    ''' </summary>
    Friend Shared Function TryResolveRealExtremePoints(
        baseView As DrawingView,
        frame As ViewPlacementFrame,
        box As ViewSheetBoundingBox,
        log As DimensionLogger,
        ByRef leftPt As DimensionExtremePoint,
        ByRef rightPt As DimensionExtremePoint,
        ByRef bottomPt As DimensionExtremePoint,
        ByRef topPt As DimensionExtremePoint) As Boolean

        leftPt = Nothing
        rightPt = Nothing
        bottomPt = Nothing
        topPt = Nothing

        If baseView Is Nothing OrElse frame Is Nothing OrElse log Is Nothing OrElse box.Width <= 0 Then Return False

        Dim relevance As Double = Math.Max(box.Width * RelevanceWidthFraction, 1.0E-9)
        Dim candidates As New List(Of DimensionExtremePoint)()

        CollectFromLines(baseView, box, frame, log, candidates, relevance)
        CollectFromArcs(baseView, box, frame, log, candidates)
        CollectFromCircles(baseView, box, frame, log, candidates)
        CollectFromEllipses(baseView, box, frame, log, candidates)

        For Each c In candidates
            log.ExtPt("candidate " & c.FormatOneLine())
        Next

        If candidates.Count < 2 Then
            log.Err("[DIM][EXTPT] candidatos insuficientes (<2); no se pueden resolver extremos.")
            Return False
        End If

        Dim minX As Double = candidates.Min(Function(p) p.XSheet)
        Dim maxX As Double = candidates.Max(Function(p) p.XSheet)
        Dim minY As Double = candidates.Min(Function(p) p.YSheet)
        Dim maxY As Double = candidates.Max(Function(p) p.YSheet)

        Const EpsSpan As Double = 1.0E-6
        Dim spanX As Double = maxX - minX
        Dim spanY As Double = maxY - minY
        If spanX <= EpsSpan Then
            log.Err(String.Format(CultureInfo.InvariantCulture, "[DIM][EXTPT] envolvente X degenerada (ΔX={0:E3}).", spanX))
            Return False
        End If
        If spanY <= EpsSpan Then
            log.Err(String.Format(CultureInfo.InvariantCulture, "[DIM][EXTPT] envolvente Y degenerada (ΔY={0:E3}).", spanY))
            Return False
        End If

        Dim tolX As Double = Math.Max(1.0E-5R, spanX * 1.0E-8R)
        Dim tolY As Double = Math.Max(1.0E-5R, spanY * 1.0E-8R)
        leftPt = SelectAtX(candidates, minX, tolX, preferLowerY:=True)
        rightPt = SelectAtX(candidates, maxX, tolX, preferLowerY:=False)
        bottomPt = SelectAtY(candidates, minY, tolY, preferLeftX:=True)
        topPt = SelectAtY(candidates, maxY, tolY, preferLeftX:=False)

        If leftPt Is Nothing OrElse rightPt Is Nothing OrElse bottomPt Is Nothing OrElse topPt Is Nothing Then
            log.Err("[DIM][EXTPT] fallo al seleccionar uno de los cuatro extremos.")
            Return False
        End If

        log.ExtPt("left = " & leftPt.FormatOneLine())
        log.ExtPt("right = " & rightPt.FormatOneLine())
        log.ExtPt("bottom = " & bottomPt.FormatOneLine())
        log.ExtPt("top = " & topPt.FormatOneLine())

        log.ExtPt(String.Format(CultureInfo.InvariantCulture,
            "span hoja: ΔX={0:0.######} ΔY={1:0.######} (desde puntos reales)",
            maxX - minX, maxY - minY))

        Return True
    End Function

    Private Shared Function SelectAtX(cands As List(Of DimensionExtremePoint), targetX As Double, tol As Double, preferLowerY As Boolean) As DimensionExtremePoint
        Dim near = cands.Where(Function(p) Math.Abs(p.XSheet - targetX) <= tol).ToList()
        If near.Count = 0 Then near = cands.Where(Function(p) Math.Abs(p.XSheet - targetX) <= tol * 100).ToList()
        If near.Count = 0 Then Return cands.OrderBy(Function(p) Math.Abs(p.XSheet - targetX)).First()
        If preferLowerY Then Return near.OrderBy(Function(p) p.YSheet).First()
        Return near.OrderByDescending(Function(p) p.YSheet).First()
    End Function

    Private Shared Function SelectAtY(cands As List(Of DimensionExtremePoint), targetY As Double, tol As Double, preferLeftX As Boolean) As DimensionExtremePoint
        Dim near = cands.Where(Function(p) Math.Abs(p.YSheet - targetY) <= tol).ToList()
        If near.Count = 0 Then near = cands.Where(Function(p) Math.Abs(p.YSheet - targetY) <= tol * 100).ToList()
        If near.Count = 0 Then Return cands.OrderBy(Function(p) Math.Abs(p.YSheet - targetY)).First()
        If preferLeftX Then Return near.OrderBy(Function(p) p.XSheet).First()
        Return near.OrderByDescending(Function(p) p.XSheet).First()
    End Function

    Private Shared Sub AddPoint(
        list As List(Of DimensionExtremePoint),
        view As DrawingView,
        frame As ViewPlacementFrame,
        sx As Double,
        sy As Double,
        sourceObj As Object,
        entityType As String,
        entityIndex As Integer,
        desc As String,
        isLineEp As Boolean,
        isArcEp As Boolean,
        isArcSamp As Boolean,
        isCirc As Boolean,
        isEll As Boolean)

        If Not (AreFinite(sx) AndAlso AreFinite(sy)) Then Return
        Dim p As New DimensionExtremePoint With {
            .XSheet = sx,
            .YSheet = sy,
            .XLocal = frame.FromSheetX(sx),
            .YLocal = frame.FromSheetY(sy),
            .SourceObject = sourceObj,
            .SourceEntityType = entityType,
            .SourceEntityIndex = entityIndex,
            .Description = desc,
            .IsFromLineEndpoint = isLineEp,
            .IsFromArcEndpoint = isArcEp,
            .IsFromArcSample = isArcSamp,
            .IsFromCircleExtreme = isCirc,
            .IsFromEllipseSample = isEll
        }
        list.Add(p)
    End Sub

    Private Shared Function AreFinite(x As Double) As Boolean
        Return Not (Double.IsNaN(x) OrElse Double.IsInfinity(x))
    End Function

    Private Shared Sub CollectFromLines(view As DrawingView, box As ViewSheetBoundingBox, frame As ViewPlacementFrame, log As DimensionLogger,
                                        list As List(Of DimensionExtremePoint), relevance As Double)
        Dim linesCol As DVLines2d = Nothing
        Try
            linesCol = view.DVLines2d
        Catch
            Return
        End Try
        If linesCol Is Nothing Then Return
        Dim n As Integer = 0
        Try
            n = linesCol.Count
        Catch
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

            Try
                Dim et As Integer = CInt(ln.EdgeType)
                If et >= 0 AndAlso et <> modelEdge Then Continue For
            Catch
            End Try

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

            Dim dx As Double = sx2 - sx1
            Dim dy As Double = sy2 - sy1
            If Math.Sqrt(dx * dx + dy * dy) < relevance Then Continue For

            AddPoint(list, view, frame, sx1, sy1, ln, "DVLine2d", i, "endpoint start", True, False, False, False, False)
            AddPoint(list, view, frame, sx2, sy2, ln, "DVLine2d", i, "endpoint end", True, False, False, False, False)
        Next
    End Sub

    Private Shared Function PointInsideLoose(sx As Double, sy As Double, box As ViewSheetBoundingBox) As Boolean
        Dim slack As Double = Math.Max(box.Width, box.Height) * 0.02R
        Return sx >= box.MinX - slack AndAlso sx <= box.MaxX + slack AndAlso sy >= box.MinY - slack AndAlso sy <= box.MaxY + slack
    End Function

    Private Shared Sub CollectFromArcs(view As DrawingView, box As ViewSheetBoundingBox, frame As ViewPlacementFrame, log As DimensionLogger, list As List(Of DimensionExtremePoint))
        Dim col As Object = Nothing
        Try
            col = view.DVArcs2d
        Catch
            Return
        End Try
        If col Is Nothing Then Return
        Dim n As Integer = 0
        Try : n = CInt(CallByName(col, "Count", CallType.Get)) : Catch : Return : End Try

        For i As Integer = 1 To n
            Dim a As Object = GetComItem(col, i)
            If a Is Nothing Then Continue For

            Try
                Dim vx1 As Double, vy1 As Double, vx2 As Double, vy2 As Double
                a.Range(vx1, vy1, vx2, vy2)
                For Each px In New Double() {vx1, vx2}
                    For Each py In New Double() {vy1, vy2}
                        Dim sx As Double, sy As Double
                        view.ViewToSheet(px, py, sx, sy)
                        If PointInsideLoose(sx, sy, box) Then
                            AddPoint(list, view, frame, sx, sy, a, "DVArc2d", i, "Range corner", False, False, True, False, False)
                        End If
                    Next
                Next
            Catch
            End Try

            Try
                Dim vx As Double, vy As Double
                a.GetStartPoint(vx, vy)
                Dim ssx As Double, ssy As Double
                view.ViewToSheet(vx, vy, ssx, ssy)
                If PointInsideLoose(ssx, ssy, box) Then AddPoint(list, view, frame, ssx, ssy, a, "DVArc2d", i, "arc start", False, True, False, False, False)
            Catch
            End Try
            Try
                Dim vx2 As Double, vy2 As Double
                a.GetEndPoint(vx2, vy2)
                Dim esx As Double, esy As Double
                view.ViewToSheet(vx2, vy2, esx, esy)
                If PointInsideLoose(esx, esy, box) Then AddPoint(list, view, frame, esx, esy, a, "DVArc2d", i, "arc end", False, True, False, False, False)
            Catch
            End Try

            Try
                Dim cx As Double, cy As Double, r As Double
                a.GetCenterPoint(cx, cy)
                r = CDbl(a.Radius)
                If r > 1.0E-12 Then
                    For k As Integer = 0 To 7
                        Dim ang As Double = k * Math.PI / 4.0R
                        Dim ax As Double = cx + r * Math.Cos(ang)
                        Dim ay As Double = cy + r * Math.Sin(ang)
                        Dim gsx As Double, gsy As Double
                        view.ViewToSheet(ax, ay, gsx, gsy)
                        If PointInsideLoose(gsx, gsy, box) Then
                            AddPoint(list, view, frame, gsx, gsy, a, "DVArc2d", i, "arc sample k=" & k.ToString(CultureInfo.InvariantCulture), False, False, True, False, False)
                        End If
                    Next
                End If
            Catch
            End Try
        Next
    End Sub

    Private Shared Sub CollectFromCircles(view As DrawingView, box As ViewSheetBoundingBox, frame As ViewPlacementFrame, log As DimensionLogger, list As List(Of DimensionExtremePoint))
        Dim col As Object = Nothing
        Try
            col = view.DVCircles2d
        Catch
            Return
        End Try
        If col Is Nothing Then Return
        Dim n As Integer = 0
        Try : n = CInt(CallByName(col, "Count", CallType.Get)) : Catch : Return : End Try

        For i As Integer = 1 To n
            Dim c As Object = GetComItem(col, i)
            If c Is Nothing Then Continue For

            Try
                Dim cx As Double, cy As Double, r As Double
                c.GetCenterPoint(cx, cy)
                r = CDbl(c.Radius)
                If r <= 1.0E-12 Then Continue For
                Dim dirs As Double()() = {
                    New Double() {1.0R, 0.0R},
                    New Double() {-1.0R, 0.0R},
                    New Double() {0.0R, 1.0R},
                    New Double() {0.0R, -1.0R}
                }
                For d As Integer = 0 To dirs.Length - 1
                    Dim px As Double = cx + r * dirs(d)(0)
                    Dim py As Double = cy + r * dirs(d)(1)
                    Dim sx As Double, sy As Double
                    view.ViewToSheet(px, py, sx, sy)
                    If PointInsideLoose(sx, sy, box) Then
                        AddPoint(list, view, frame, sx, sy, c, "DVCircle2d", i, "circle extreme dir=" & d.ToString(CultureInfo.InvariantCulture), False, False, False, True, False)
                    End If
                Next
            Catch
            End Try
        Next
    End Sub

    Private Shared Sub CollectFromEllipses(view As DrawingView, box As ViewSheetBoundingBox, frame As ViewPlacementFrame, log As DimensionLogger, list As List(Of DimensionExtremePoint))
        Dim col As Object = Nothing
        Try
            col = view.DVEllipses2d
        Catch
            Return
        End Try
        If col Is Nothing Then Return
        Dim n As Integer = 0
        Try : n = CInt(CallByName(col, "Count", CallType.Get)) : Catch : Return : End Try

        For i As Integer = 1 To n
            Dim e As Object = GetComItem(col, i)
            If e Is Nothing Then Continue For
            Try
                Dim vx1 As Double, vy1 As Double, vx2 As Double, vy2 As Double
                e.Range(vx1, vy1, vx2, vy2)
                For Each px In New Double() {vx1, vx2}
                    For Each py In New Double() {vy1, vy2}
                        Dim sx As Double, sy As Double
                        view.ViewToSheet(px, py, sx, sy)
                        If PointInsideLoose(sx, sy, box) Then
                            AddPoint(list, view, frame, sx, sy, e, "DVEllipse2d", i, "ellipse Range corner", False, False, False, False, True)
                        End If
                    Next
                Next
            Catch
            End Try
        Next
    End Sub

    Private Shared Function GetComItem(coll As Object, index As Integer) As Object
        Try
            Return CallByName(coll, "Item", CallType.Get, index)
        Catch
            Try
                Return CallByName(coll, "Item", CallType.Method, index)
            Catch
                Return Nothing
            End Try
        End Try
    End Function

End Class
