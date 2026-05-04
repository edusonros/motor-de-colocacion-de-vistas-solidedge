Option Strict Off

Imports System.Collections.Generic
Imports System.Globalization
Imports System.Linq
Imports SolidEdgeDraft

''' <summary>
''' Selección de objetos 2D (líneas, arcos, círculos) para la cota vertical exterior según envolvente Y real en hoja.
''' </summary>
Friend NotInheritable Class VerticalExteriorAnchors

    Private Const PriorityArcCircle As Integer = 1
    Private Const PriorityLine As Integer = 2

    Private Class YExtentCandidate
        Public Entity As Object
        Public Kind As String
        Public Index As Integer
        Public Priority As Integer
        Public YminSheet As Double
        Public YmaxSheet As Double
        Public XMidSheet As Double
    End Class

    Public Class ResolveResult
        Public Success As Boolean
        Public UsedFallback As Boolean
        Public BottomObject As Object
        Public TopObject As Object
        Public Y1Sheet As Double
        Public Y2Sheet As Double
        ''' <summary>Extremo Y inferior de las entidades candidatas en coordenadas de hoja (no bbox de hoja completo).</summary>
        Public ExtentYMinSheet As Double
        ''' <summary>Extremo Y superior de las entidades candidatas en coordenadas de hoja.</summary>
        Public ExtentYMaxSheet As Double
        Public BottomDescription As String
        Public TopDescription As String
        Public CandidateHeight As Double
    End Class

    Private Sub New()
    End Sub

    Public Shared Function TryResolve(
        view As DrawingView,
        box As ViewSheetBoundingBox,
        extreme As ExtremeDvLinesResult,
        log As DimensionLogger,
        ByRef outResult As ResolveResult) As Boolean

        outResult = New ResolveResult With {.Success = False, .UsedFallback = False}
        If view Is Nothing OrElse log Is Nothing Then Return False

        Dim slack As Double = Math.Max(box.Height * 0.15R, 1.0E-6)
        Dim eps As Double = Math.Max(box.Height * 1.0E-6R, 1.0E-9)

        log.Vert(String.Format(CultureInfo.InvariantCulture,
            "BoundingBox alto total = {0:0.######}m (ymin={1:0.######} ymax={2:0.######})",
            box.Height, box.MinY, box.MaxY))

        Dim candidates As New List(Of YExtentCandidate)()
        CollectFromDvLines2d(view, box, candidates, slack)
        CollectFromArcs(view, box, candidates, slack)
        CollectFromCircles(view, box, candidates, slack)

        If candidates.Count = 0 Then
            log.VertWarn("No se encontraron extremos verticales válidos (sin entidades 2D en envolvente).")
            Return TryFallback(extreme, box, log, outResult)
        End If

        Dim globalYmin As Double = candidates.Min(Function(c) c.YminSheet)
        Dim globalYmax As Double = candidates.Max(Function(c) c.YmaxSheet)

        log.Vert(String.Format(CultureInfo.InvariantCulture,
            "ymax entidades (hoja, recorte vista base) = {0:0.######}m", globalYmax))
        log.Vert(String.Format(CultureInfo.InvariantCulture,
            "ymin entidades (hoja, recorte vista base) = {0:0.######}m", globalYmin))

        outResult.ExtentYMinSheet = globalYmin
        outResult.ExtentYMaxSheet = globalYmax
        outResult.CandidateHeight = globalYmax - globalYmin

        If outResult.CandidateHeight <= 1.0E-9 Then
            log.VertWarn("Altura candidata ~0; fallback.")
            Return TryFallback(extreme, box, log, outResult)
        End If

        If globalYmin < box.MinY - slack OrElse globalYmax > box.MaxY + slack Then
            log.VertWarn(String.Format(CultureInfo.InvariantCulture,
                "Envolvente Y fuera del bbox ampliado (slack={0:0.######}m); se mantiene selección geométrica.",
                slack))
        End If

        Dim bottomC As YExtentCandidate = PickBottomCandidate(candidates, globalYmin, eps)
        Dim topC As YExtentCandidate = PickTopCandidate(candidates, globalYmax, eps)

        If bottomC Is Nothing OrElse topC Is Nothing Then
            log.VertWarn("No se encontraron extremos verticales válidos (selección vacía).")
            Return TryFallback(extreme, box, log, outResult)
        End If

        If ReferenceEquals(bottomC.Entity, topC.Entity) Then
            log.VertWarn("Misma entidad arriba y abajo; se busca alternativa o fallback.")
            Dim othersTop = candidates.Where(Function(c) Not ReferenceEquals(c.Entity, bottomC.Entity)).ToList()
            Dim altTop = PickTopCandidate(othersTop, globalYmax, eps)
            If altTop IsNot Nothing Then
                topC = altTop
            Else
                Dim othersBot = candidates.Where(Function(c) Not ReferenceEquals(c.Entity, topC.Entity)).ToList()
                Dim altBot = PickBottomCandidate(othersBot, globalYmin, eps)
                If altBot IsNot Nothing Then
                    bottomC = altBot
                Else
                    Return TryFallback(extreme, box, log, outResult)
                End If
            End If
        End If

        outResult.BottomObject = bottomC.Entity
        outResult.TopObject = topC.Entity
        outResult.Y1Sheet = globalYmin
        outResult.Y2Sheet = globalYmax
        outResult.BottomDescription = bottomC.Kind & " idx=" & bottomC.Index.ToString(CultureInfo.InvariantCulture) &
            " Y[" & bottomC.YminSheet.ToString("0.######", CultureInfo.InvariantCulture) & ".." &
            bottomC.YmaxSheet.ToString("0.######", CultureInfo.InvariantCulture) & "]"
        outResult.TopDescription = topC.Kind & " idx=" & topC.Index.ToString(CultureInfo.InvariantCulture) &
            " Y[" & topC.YminSheet.ToString("0.######", CultureInfo.InvariantCulture) & ".." &
            topC.YmaxSheet.ToString("0.######", CultureInfo.InvariantCulture) & "]"

        log.Vert("entidad inferior elegida = " & outResult.BottomDescription)
        log.Vert("entidad superior elegida = " & outResult.TopDescription)
        log.Vert(String.Format(CultureInfo.InvariantCulture,
            "altura candidata = {0:0.######}m (bbox alto referencia = {1:0.######}m)",
            outResult.CandidateHeight, box.Height))

        outResult.Success = True
        Return True
    End Function

    Private Shared Function PickBottomCandidate(cands As List(Of YExtentCandidate), globalYmin As Double, eps As Double) As YExtentCandidate
        If cands Is Nothing OrElse cands.Count = 0 Then Return Nothing
        Dim near = cands.Where(Function(c) c.YminSheet <= globalYmin + eps).ToList()
        If near.Count = 0 Then near = cands
        Return near.OrderBy(Function(c) c.Priority).ThenBy(Function(c) c.YminSheet).FirstOrDefault()
    End Function

    Private Shared Function PickTopCandidate(cands As List(Of YExtentCandidate), globalYmax As Double, eps As Double) As YExtentCandidate
        If cands Is Nothing OrElse cands.Count = 0 Then Return Nothing
        Dim near = cands.Where(Function(c) c.YmaxSheet >= globalYmax - eps).ToList()
        If near.Count = 0 Then near = cands
        Return near.OrderBy(Function(c) c.Priority).ThenByDescending(Function(c) c.YmaxSheet).FirstOrDefault()
    End Function

    Private Shared Function TryFallback(extreme As ExtremeDvLinesResult, box As ViewSheetBoundingBox, log As DimensionLogger, ByRef outResult As ResolveResult) As Boolean
        log.VertFallback("Se usa lógica antigua (líneas horizontales extremas inferiores/superiores).")
        If extreme Is Nothing OrElse extreme.BottomHorizontal Is Nothing OrElse extreme.TopHorizontal Is Nothing Then
            log.VertWarn("Fallback imposible: faltan BottomHorizontal/TopHorizontal.")
            Return False
        End If
        If extreme.BottomHorizontal.Line Is Nothing OrElse extreme.TopHorizontal.Line Is Nothing Then Return False

        outResult.UsedFallback = True
        outResult.BottomObject = extreme.BottomHorizontal.Line
        outResult.TopObject = extreme.TopHorizontal.Line
        outResult.Y1Sheet = extreme.BottomHorizontal.MidY
        outResult.Y2Sheet = extreme.TopHorizontal.MidY
        outResult.ExtentYMinSheet = Math.Min(extreme.BottomHorizontal.MinYs, extreme.BottomHorizontal.MaxYs)
        outResult.ExtentYMaxSheet = Math.Max(extreme.TopHorizontal.MinYs, extreme.TopHorizontal.MaxYs)
        outResult.CandidateHeight = outResult.ExtentYMaxSheet - outResult.ExtentYMinSheet
        outResult.BottomDescription = "FALLBACK DVLine2d índice=" & extreme.BottomHorizontal.SourceIndex.ToString(CultureInfo.InvariantCulture)
        outResult.TopDescription = "FALLBACK DVLine2d índice=" & extreme.TopHorizontal.SourceIndex.ToString(CultureInfo.InvariantCulture)
        outResult.Success = True
        log.Vert("entidad inferior elegida = " & outResult.BottomDescription)
        log.Vert("entidad superior elegida = " & outResult.TopDescription)
        log.Vert(String.Format(CultureInfo.InvariantCulture,
            "altura candidata = {0:0.######}m", outResult.CandidateHeight))
        Return True
    End Function

    Private Shared Sub CollectFromDvLines2d(view As DrawingView, box As ViewSheetBoundingBox, candidates As List(Of YExtentCandidate), slack As Double)
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

        For i As Integer = 1 To n
            Dim ln As DVLine2d = Nothing
            Try
                ln = CType(linesCol.Item(i), DVLine2d)
            Catch
                Continue For
            End Try
            If ln Is Nothing Then Continue For

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

            Dim ya As Double = Math.Min(sy1, sy2)
            Dim yb As Double = Math.Max(sy1, sy2)
            If yb < box.MinY - slack OrElse ya > box.MaxY + slack Then Continue For

            candidates.Add(New YExtentCandidate With {
                .Entity = ln,
                .Kind = "DVLine2d",
                .Index = i,
                .Priority = PriorityLine,
                .YminSheet = ya,
                .YmaxSheet = yb,
                .XMidSheet = (sx1 + sx2) / 2.0R
            })
        Next
    End Sub

    Private Shared Sub CollectFromArcs(view As DrawingView, box As ViewSheetBoundingBox, candidates As List(Of YExtentCandidate), slack As Double)
        Dim col As Object = Nothing
        Try
            col = view.DVArcs2d
        Catch
            Return
        End Try
        If col Is Nothing Then Return
        Dim n As Integer = 0
        Try : n = CInt(col.Count) : Catch : Return : End Try

        For i As Integer = 1 To n
            Dim a As Object = Nothing
            Try : a = GetComItem(col, i) : Catch : Continue For : End Try
            If a Is Nothing Then Continue For

            Dim ymin As Double, ymax As Double, xMid As Double
            If Not TryArcSheetYExtents(view, a, ymin, ymax, xMid) Then Continue For
            If ymax < box.MinY - slack OrElse ymin > box.MaxY + slack Then Continue For

            candidates.Add(New YExtentCandidate With {
                .Entity = a,
                .Kind = "DVArc2d",
                .Index = i,
                .Priority = PriorityArcCircle,
                .YminSheet = ymin,
                .YmaxSheet = ymax,
                .XMidSheet = xMid
            })
        Next
    End Sub

    Private Shared Sub CollectFromCircles(view As DrawingView, box As ViewSheetBoundingBox, candidates As List(Of YExtentCandidate), slack As Double)
        Dim col As Object = Nothing
        Try
            col = view.DVCircles2d
        Catch
            Return
        End Try
        If col Is Nothing Then Return
        Dim n As Integer = 0
        Try : n = CInt(col.Count) : Catch : Return : End Try

        For i As Integer = 1 To n
            Dim c As Object = Nothing
            Try : c = GetComItem(col, i) : Catch : Continue For : End Try
            If c Is Nothing Then Continue For

            Dim ymin As Double, ymax As Double, xMid As Double
            If Not TryCircleSheetYExtents(view, c, ymin, ymax, xMid) Then Continue For
            If ymax < box.MinY - slack OrElse ymin > box.MaxY + slack Then Continue For

            candidates.Add(New YExtentCandidate With {
                .Entity = c,
                .Kind = "DVCircle2d",
                .Index = i,
                .Priority = PriorityArcCircle,
                .YminSheet = ymin,
                .YmaxSheet = ymax,
                .XMidSheet = xMid
            })
        Next
    End Sub

    Private Shared Sub AccumSheetY(view As DrawingView, vx As Double, vy As Double,
                                   ByRef has As Boolean, ByRef ymin As Double, ByRef ymax As Double,
                                   ByRef xSum As Double, ByRef nPts As Integer)
        Dim sx As Double, sy As Double
        view.ViewToSheet(vx, vy, sx, sy)
        If Not has Then
            ymin = sy
            ymax = sy
            has = True
        Else
            ymin = Math.Min(ymin, sy)
            ymax = Math.Max(ymax, sy)
        End If
        xSum += sx
        nPts += 1
    End Sub

    Private Shared Function TryArcSheetYExtents(view As DrawingView, a As Object, ByRef ymin As Double, ByRef ymax As Double, ByRef xMid As Double) As Boolean
        Dim has As Boolean = False
        Dim xSum As Double = 0R
        Dim nPts As Integer = 0
        ymin = 0R
        ymax = 0R
        xMid = 0R

        Try
            Dim vx1 As Double, vy1 As Double, vx2 As Double, vy2 As Double
            a.Range(vx1, vy1, vx2, vy2)
            AccumSheetY(view, vx1, vy1, has, ymin, ymax, xSum, nPts)
            AccumSheetY(view, vx2, vy1, has, ymin, ymax, xSum, nPts)
            AccumSheetY(view, vx1, vy2, has, ymin, ymax, xSum, nPts)
            AccumSheetY(view, vx2, vy2, has, ymin, ymax, xSum, nPts)
        Catch
        End Try

        Try
            Dim cx As Double, cy As Double, r As Double
            a.GetCenterPoint(cx, cy)
            r = CDbl(a.Radius)
            If r > 1.0E-12 Then
                For k As Integer = 0 To 7
                    Dim ang As Double = k * Math.PI / 4.0R
                    AccumSheetY(view, cx + r * Math.Cos(ang), cy + r * Math.Sin(ang), has, ymin, ymax, xSum, nPts)
                Next
            End If
        Catch
        End Try

        Try
            Dim sx As Double, sy As Double
            a.GetStartPoint(sx, sy)
            AccumSheetY(view, sx, sy, has, ymin, ymax, xSum, nPts)
        Catch
        End Try
        Try
            Dim ex As Double, ey As Double
            a.GetEndPoint(ex, ey)
            AccumSheetY(view, ex, ey, has, ymin, ymax, xSum, nPts)
        Catch
        End Try

        If nPts = 0 Then Return False
        xMid = xSum / nPts
        Return True
    End Function

    Private Shared Function TryCircleSheetYExtents(view As DrawingView, c As Object, ByRef ymin As Double, ByRef ymax As Double, ByRef xMid As Double) As Boolean
        Dim has As Boolean = False
        Dim xSum As Double = 0R
        Dim nPts As Integer = 0
        ymin = 0R
        ymax = 0R
        xMid = 0R

        Try
            Dim vx1 As Double, vy1 As Double, vx2 As Double, vy2 As Double
            c.Range(vx1, vy1, vx2, vy2)
            AccumSheetY(view, vx1, vy1, has, ymin, ymax, xSum, nPts)
            AccumSheetY(view, vx2, vy1, has, ymin, ymax, xSum, nPts)
            AccumSheetY(view, vx1, vy2, has, ymin, ymax, xSum, nPts)
            AccumSheetY(view, vx2, vy2, has, ymin, ymax, xSum, nPts)
        Catch
        End Try

        Try
            Dim cx As Double, cy As Double, r As Double
            c.GetCenterPoint(cx, cy)
            r = CDbl(c.Radius)
            If r > 1.0E-12 Then
                For k As Integer = 0 To 7
                    Dim ang As Double = k * Math.PI / 4.0R
                    AccumSheetY(view, cx + r * Math.Cos(ang), cy + r * Math.Sin(ang), has, ymin, ymax, xSum, nPts)
                Next
            End If
        Catch
        End Try

        If nPts = 0 Then Return False
        xMid = xSum / nPts
        Return True
    End Function

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
