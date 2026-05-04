Option Strict Off

Imports System.Globalization
Imports SolidEdgeDraft
Imports SolidEdgeFrameworkSupport
Imports FrameworkDimension = SolidEdgeFrameworkSupport.Dimension

Public NotInheritable Class DimensionValidationService
    Private Sub New()
    End Sub

    Public Shared Function CumpleNormaISO129(cand As DimensionCandidate, config As DimensioningNormConfig) As Boolean
        If cand Is Nothing OrElse config Is Nothing Then Return False
        If Not config.EnableISO129Rules Then Return True
        If cand.IsRedundant Then Return False
        If config.AvoidDuplicateDimensions AndAlso cand.NominalValue <= config.MinFeatureSeparation Then Return False
        If config.AvoidHiddenGeometry AndAlso cand.UsesHiddenGeometry Then Return False
        If Not cand.ViewShowsFeatureClearly Then Return False
        Return True
    End Function

    Public Shared Function ValidarYRecolocarCota(
        sheet As Sheet,
        dimObj As Object,
        cand As DimensionCandidate,
        bbox As BoundingBox2D,
        config As DimensioningNormConfig,
        log As Action(Of String)) As Boolean

        If dimObj Is Nothing Then Return False
        If config Is Nothing Then config = DimensioningNormConfig.DefaultConfig()

        Dim d As FrameworkDimension = TryCast(dimObj, FrameworkDimension)
        If d Is Nothing Then
            log?.Invoke("[DIM][ISO129][VALIDATE] tipo no FrameworkDimension; keep pragmatic=" & config.AutoDimPragmaticVisibleFirst.ToString())
            Return config.AutoDimPragmaticVisibleFirst
        End If

        Dim rangeOk As Boolean = False
        Dim rx1 As Double, ry1 As Double, rx2 As Double, ry2 As Double
        Try
            d.Range(rx1, ry1, rx2, ry2)
            rangeOk = AreFinite(rx1, ry1, rx2, ry2)
        Catch ex As Exception
            log?.Invoke("[DIM][ISO129][VALIDATE] range_unavailable_keep_by_pragmatic_mode=" & config.AutoDimPragmaticVisibleFirst.ToString() & " ex=" & ex.Message)
            Return config.AutoDimPragmaticVisibleFirst
        End Try

        If Not rangeOk Then
            log?.Invoke("[DIM][ISO129][VALIDATE] range_unavailable_keep_by_pragmatic_mode=True")
            Return True
        End If

        Dim sheetArea = DimensionPlacementService.TryGetSheetWorkArea(sheet, log)
        Dim minX As Double = Math.Min(rx1, rx2)
        Dim maxX As Double = Math.Max(rx1, rx2)
        Dim minY As Double = Math.Min(ry1, ry2)
        Dim maxY As Double = Math.Max(ry1, ry2)

        If config.AvoidBorder Then
            Dim m As Double = 0.002R
            If minX < sheetArea.MinX - m OrElse maxX > sheetArea.MaxX + m OrElse minY < sheetArea.MinY - m OrElse maxY > sheetArea.MaxY + m Then
                log?.Invoke(String.Format(CultureInfo.InvariantCulture,
                    "[DIM][ISO129][VALIDATE] fuera_hoja margin m={0:0.######} pragmatic_keep={1}",
                    m, config.AutoDimPragmaticVisibleFirst))
                If Not config.AutoDimPragmaticVisibleFirst Then Return False
            End If
        End If

        If config.AvoidTitleBlock Then
            Dim tb = DimensionPlacementService.GetTitleBlockAvoidanceBox(sheetArea, config)
            If tb IsNot Nothing AndAlso BoxesOverlap(minX, minY, maxX, maxY, tb.MinX, tb.MinY, tb.MaxX, tb.MaxY) Then
                log?.Invoke("[DIM][ISO129][VALIDATE] cruza_cajetin_heuristic")
                If Not config.AutoDimPragmaticVisibleFirst Then Return False
            End If
        End If

        If cand IsNot Nothing AndAlso cand.NominalValue > 1.0E-9 Then
            Dim span As Double = Math.Max(Math.Abs(rx2 - rx1), Math.Abs(ry2 - ry1))
            If span > cand.NominalValue * 50.0R Then
                log?.Invoke("[DIM][ISO129][VALIDATE] valor_absurdo_range_vs_nominal")
                Return False
            End If
        End If

        If bbox IsNot Nothing AndAlso config.AvoidInsideContour Then
            Dim rcx As Double = (minX + maxX) / 2.0R
            Dim rcy As Double = (minY + maxY) / 2.0R
            If rcx > bbox.MinX AndAlso rcx < bbox.MaxX AndAlso rcy > bbox.MinY AndAlso rcy < bbox.MaxY Then
                log?.Invoke("[DIM][ISO129][VALIDATE] centro_range_dentro_bbox_geometrico (no borrar en modo pragmático)")
            End If
        End If

        log?.Invoke("[DIM][ISO129][VALIDATE] OK range=(" & minX.ToString("G9", CultureInfo.InvariantCulture) & "," & minY.ToString("G9", CultureInfo.InvariantCulture) & ")-(" &
                    maxX.ToString("G9", CultureInfo.InvariantCulture) & "," & maxY.ToString("G9", CultureInfo.InvariantCulture) & ")")
        Return True
    End Function

    Private Shared Function BoxesOverlap(ax1 As Double, ay1 As Double, ax2 As Double, ay2 As Double,
                                         bx1 As Double, by1 As Double, bx2 As Double, by2 As Double) As Boolean
        Return Not (ax2 < bx1 OrElse ax1 > bx2 OrElse ay2 < by1 OrElse ay1 > by2)
    End Function

    Private Shared Function AreFinite(x1 As Double, y1 As Double, x2 As Double, y2 As Double) As Boolean
        Return Not (Double.IsNaN(x1) OrElse Double.IsNaN(y1) OrElse Double.IsNaN(x2) OrElse Double.IsNaN(y2))
    End Function
End Class
