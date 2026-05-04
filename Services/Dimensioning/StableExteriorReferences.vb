Option Strict Off

Imports System.Globalization
Imports SolidEdgeDraft

''' <summary>
''' Prioriza referencias de contorno desde endpoints de <see cref="DVLine2d"/> (aristas modelo) para cotas exteriores.
''' Arcos/círculos no se evalúan aquí; si no hay datos suficientes, el llamador usa <see cref="RealExtremePointsResolver"/>.
''' </summary>
Friend NotInheritable Class StableExteriorReferences

    Friend Shared Function TryResolveLineFirstExtremes(
        view As DrawingView,
        frame As ViewPlacementFrame,
        log As DimensionLogger,
        ByRef leftPt As DimensionExtremePoint,
        ByRef rightPt As DimensionExtremePoint,
        ByRef bottomPt As DimensionExtremePoint,
        ByRef topPt As DimensionExtremePoint,
        ByRef reason As String) As Boolean

        leftPt = Nothing
        rightPt = Nothing
        bottomPt = Nothing
        topPt = Nothing
        reason = "init"
        If view Is Nothing OrElse frame Is Nothing Then
            reason = "null_view_or_frame"
            Return False
        End If

        Dim lines As DVLines2d = Nothing
        Try
            lines = view.DVLines2d
        Catch
            lines = Nothing
        End Try
        If lines Is Nothing Then
            reason = "no_DVLines2d"
            log?.LogLine("[DIM][EXT][LINE] no DVLines2d; se usará resolvedor completo (líneas+arcos+…).")
            Return False
        End If

        Dim lc As Integer = 0
        Try
            lc = CInt(lines.Count)
        Catch
            lc = 0
        End Try
        If lc <= 0 Then
            reason = "lines_count_0"
            log?.LogLine("[DIM][EXT][LINE] DVLines2d.Count=0; fallback.")
            Return False
        End If

        Dim modelEdge As Integer = CInt(SolidEdgeConstants.GraphicMemberEdgeTypeConstants.seModelEdgeType)

        Dim minX As Double = Double.PositiveInfinity
        Dim maxX As Double = Double.NegativeInfinity
        Dim minY As Double = Double.PositiveInfinity
        Dim maxY As Double = Double.NegativeInfinity

        Dim bestMinX As ExtremeRecord = Nothing
        Dim bestMaxX As ExtremeRecord = Nothing
        Dim bestMinY As ExtremeRecord = Nothing
        Dim bestMaxY As ExtremeRecord = Nothing

        For i As Integer = 1 To lc
            Dim ln As DVLine2d = Nothing
            Try
                ln = CType(lines.Item(i), DVLine2d)
            Catch
                ln = Nothing
            End Try
            If ln Is Nothing Then Continue For

            Try
                Dim et As Integer = CInt(ln.EdgeType)
                If et >= 0 AndAlso et <> modelEdge Then Continue For
            Catch
            End Try

            Dim vx1 As Double = 0, vy1 As Double = 0, vx2 As Double = 0, vy2 As Double = 0
            Try
                ln.GetStartPoint(vx1, vy1)
                ln.GetEndPoint(vx2, vy2)
            Catch
                Continue For
            End Try

            ConsiderEndpoint(view, frame, ln, i, vx1, vy1, "start", minX, maxX, minY, maxY, bestMinX, bestMaxX, bestMinY, bestMaxY)
            ConsiderEndpoint(view, frame, ln, i, vx2, vy2, "end", minX, maxX, minY, maxY, bestMinX, bestMaxX, bestMinY, bestMaxY)
        Next

        If bestMinX Is Nothing OrElse bestMaxX Is Nothing OrElse bestMinY Is Nothing OrElse bestMaxY Is Nothing Then
            reason = "incomplete_extremes"
            log?.LogLine("[DIM][EXT][LINE] extremos incompletos desde solo líneas; fallback a geometría mixta.")
            Return False
        End If

        leftPt = BuildPoint(bestMinX, frame)
        rightPt = BuildPoint(bestMaxX, frame)
        bottomPt = BuildPoint(bestMinY, frame)
        topPt = BuildPoint(bestMaxY, frame)

        reason = "line_endpoints_" & lc.ToString(CultureInfo.InvariantCulture) & "_lines_scanned"
        If log IsNot Nothing Then
            log.LogLine("[DIM][EXT][LINE] elegidos extremos solo-línea: " & reason)
            log.LogLine("[DIM][EXT][LINE] descartado en esta pasada: DVArc2d/DVCircle2d/DVEllipse2d (último recurso en resolvedor completo).")
        End If
        Return True
    End Function

    Private NotInheritable Class ExtremeRecord
        Friend Line As DVLine2d
        Friend Index As Integer
        Friend Vx As Double
        Friend Vy As Double
        Friend Sx As Double
        Friend Sy As Double
        Friend Ep As String
    End Class

    Private Shared Sub ConsiderEndpoint(
        view As DrawingView,
        frame As ViewPlacementFrame,
        ln As DVLine2d,
        idx As Integer,
        vx As Double,
        vy As Double,
        ep As String,
        ByRef minX As Double,
        ByRef maxX As Double,
        ByRef minY As Double,
        ByRef maxY As Double,
        ByRef bestMinX As ExtremeRecord,
        ByRef bestMaxX As ExtremeRecord,
        ByRef bestMinY As ExtremeRecord,
        ByRef bestMaxY As ExtremeRecord)

        Dim sx As Double = 0, sy As Double = 0
        Try
            view.ViewToSheet(vx, vy, sx, sy)
        Catch
            Return
        End Try

        Dim rec As New ExtremeRecord With {
            .Line = ln,
            .Index = idx,
            .Vx = vx,
            .Vy = vy,
            .Sx = sx,
            .Sy = sy,
            .Ep = ep
        }

        If sx < minX Then
            minX = sx
            bestMinX = rec
        End If
        If sx > maxX Then
            maxX = sx
            bestMaxX = rec
        End If
        If sy < minY Then
            minY = sy
            bestMinY = rec
        End If
        If sy > maxY Then
            maxY = sy
            bestMaxY = rec
        End If
    End Sub

    Private Shared Function BuildPoint(r As ExtremeRecord, frame As ViewPlacementFrame) As DimensionExtremePoint
        Return New DimensionExtremePoint With {
            .XSheet = r.Sx,
            .YSheet = r.Sy,
            .XLocal = frame.FromSheetX(r.Sx),
            .YLocal = frame.FromSheetY(r.Sy),
            .SourceObject = r.Line,
            .SourceEntityType = "DVLine2d",
            .SourceEntityIndex = r.Index,
            .Description = "line ep " & r.Ep,
            .IsFromLineEndpoint = True,
            .IsFromArcEndpoint = False,
            .IsFromArcSample = False,
            .IsFromCircleExtreme = False,
            .IsFromEllipseSample = False
        }
    End Function

End Class
