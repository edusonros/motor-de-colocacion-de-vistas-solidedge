Option Strict Off

Imports System.Collections.Generic
Imports System.Globalization
Imports System.Linq
Imports SolidEdgeDraft
Imports SolidEdgeFrameworkSupport

''' <summary>Acotación objetivo ~18–21 cotas (modo TargetDrawingLikeReference), sin barrido masivo DV.</summary>
Friend NotInheritable Class ReferenceDrawingDimensioningService
    Private Const MinLineLenM As Double = 0.003R
    ''' <summary>Si es True, no se escribe TrackDistance tras crear la cota (prueba: posición por defecto de Solid Edge).</summary>
    Private Const DisableReferenceTrackDistanceAdjust As Boolean = True

    Private Sub New()
    End Sub

    Private Shared Sub CollectOverallDimensions(info As DrawingViewGeometryInfo, out As List(Of DimensionCandidate), log As DimensionLogger)
        If info Is Nothing OrElse info.Extreme Is Nothing Then Return
        Dim ex As ExtremeDvLinesResult = info.Extreme
        If ex.LeftVertical IsNot Nothing AndAlso ex.RightVertical IsNot Nothing Then
            Dim yRef As Double = (ex.LeftVertical.MidY + ex.RightVertical.MidY) * 0.5R
            Dim w As Double = Math.Abs(ex.RightVertical.MidX - ex.LeftVertical.MidX)
            If w > 0.001R Then
                out.Add(New DimensionCandidate With {
                    .View = info.View,
                    .Type = DimensionCandidateType.TotalHorizontal,
                    .Orientation = DimensionOrientation.Horizontal,
                    .SourceObject1 = ex.LeftVertical.Line,
                    .SourceObject2 = ex.RightVertical.Line,
                    .P1 = New Point2D With {.X = ex.LeftVertical.MidX, .Y = yRef},
                    .P2 = New Point2D With {.X = ex.RightVertical.MidX, .Y = yRef},
                    .NominalValue = w,
                    .Priority = 1,
                    .PlacementSide = DimensionSide.Bottom
                })
                log?.LogLine("[DIM][CAND][TOTAL] view=" & info.ViewIndex.ToString(CultureInfo.InvariantCulture) & " axis=H value=" & FormatInv(w * 1000.0R) & "mm")
            End If
        End If
        If ex.BottomHorizontal IsNot Nothing AndAlso ex.TopHorizontal IsNot Nothing Then
            Dim xRef As Double = (ex.BottomHorizontal.MidX + ex.TopHorizontal.MidX) * 0.5R
            Dim h As Double = Math.Abs(ex.TopHorizontal.MidY - ex.BottomHorizontal.MidY)
            If h > 0.001R Then
                out.Add(New DimensionCandidate With {
                    .View = info.View,
                    .Type = DimensionCandidateType.TotalVertical,
                    .Orientation = DimensionOrientation.Vertical,
                    .SourceObject1 = ex.BottomHorizontal.Line,
                    .SourceObject2 = ex.TopHorizontal.Line,
                    .P1 = New Point2D With {.X = xRef, .Y = ex.BottomHorizontal.MidY},
                    .P2 = New Point2D With {.X = xRef, .Y = ex.TopHorizontal.MidY},
                    .NominalValue = h,
                    .Priority = 1,
                    .PlacementSide = DimensionSide.Right
                })
                log?.LogLine("[DIM][CAND][TOTAL] view=" & info.ViewIndex.ToString(CultureInfo.InvariantCulture) & " axis=V value=" & FormatInv(h * 1000.0R) & "mm")
            End If
        End If
    End Sub

    Public Shared Sub Run(
        ctx As SolidEdgeContext,
        workList As List(Of DrawingViewGeometryInfo),
        styleObj As Object,
        norm As DimensioningNormConfig,
        log As DimensionLogger,
        baseLogger As Logger,
        protectedZones As IList(Of ProtectedZone2D))

        If ctx Is Nothing OrElse norm Is Nothing OrElse workList Is Nothing Then Return

        LogSkippedIsometricViews(ctx.Sheet, log)

        log?.LogLine("[DIM][MODE] " & norm.DimensionCreationMode)
        log?.LogLine("[DIM][SWEEP][SKIP] reason=target_reference_mode")

        Dim sheet As Sheet = ctx.Sheet
        Dim draft As DraftDocument = ctx.Draft
        Dim dims As Dimensions = Nothing
        Try
            dims = CType(sheet.Dimensions, Dimensions)
        Catch ex As Exception
            log?.ComFail("Sheet.Dimensions", "Sheet", ex)
            Return
        End Try
        If dims Is Nothing Then Return

        DrawingViewDimensionCreator.TryActivateTargetSheet(draft, sheet, log)
        Dim effStyle As Object = DrawingViewDimensionCreator.ResolveForcedStyleObject(draft, sheet, styleObj, log)
        DrawingViewDimensionCreator.TryApplyStyleToDimensionsCollection(dims, effStyle, log, "REF_TARGET")

        Dim exteriorCreated As Integer = CreateExteriorTotalsForAllViews(workList, dims, log)
        If exteriorCreated > 0 Then
            log?.LogLine("[DIM][DVREF][TOTALS_ALL_VIEWS] created=" & exteriorCreated.ToString(CultureInfo.InvariantCulture))
        End If

        Dim candidates As List(Of DimensionCandidate) = CrearCandidatosDeCotasReferencia(workList, norm, log)
        Dim selected As List(Of DimensionCandidate) = SeleccionarCotasReferencia(candidates, norm, log)

        Dim createdLin As Integer = 0
        Dim createdRad As Integer = 0
        Dim tracked As New List(Of Tuple(Of Dimension, DimensionCandidate, String))()
        Dim dimSeq As Integer = 0
        For Each c In selected
            If c Is Nothing OrElse c.View Is Nothing Then Continue For
            If GenerationEngineRuntime.DebugDiagnosticsMode OrElse Not GenerationEngineRuntime.ProductionMode Then
                log?.LogLine("[DIM][CREATE][TRY] view=" & ViewNameOf(c) & " type=" & c.Type.ToString() & " nominal_m=" &
                             FormatInv(c.NominalValue))
            End If

            Dim dimObj As Dimension = Nothing
            Try
                If c.SourceObject2 IsNot Nothing AndAlso c.P1 IsNot Nothing AndAlso c.P2 IsNot Nothing Then
                    Dim axis As String = If(c.Orientation = DimensionOrientation.Vertical, "vertical", "horizontal")
                    Dim methodUsed As String = ""
                    Dim dObj As Dimension = Nothing
                    Dim ok As Boolean = DimensionPlacementEngine.TryAddDistanceBetweenObjectsGeneric(
                        dims, c.SourceObject1, c.SourceObject2,
                        c.P1.X, c.P1.Y, c.P2.X, c.P2.Y,
                        log, axis, methodUsed, "REF_TOTAL", Nothing, c.View, False, dObj)
                    If ok Then
                        dimObj = dObj
                    End If
                ElseIf c.Type = DimensionCandidateType.Radius Then
                    dimObj = DrawingViewDimensionCreator.TryCreateRadiusOnReference(dims, c.SourceObject1, log)
                ElseIf c.Type = DimensionCandidateType.Diameter Then
                    dimObj = DrawingViewDimensionCreator.TryCreateDiameterOnReference(dims, c.SourceObject1, log)
                Else
                    dimObj = DrawingViewDimensionCreator.TryCreateAddLengthOnReference(dims, c.SourceObject1, log)
                End If

                If dimObj Is Nothing Then
                    log?.LogLine("[DIM][CREATE][FAIL] view=" & ViewNameOf(c) & " reason=null_dimension")
                    Continue For
                End If

                DrawingViewDimensionCreator.TryApplyStyleToDimension(dimObj, effStyle, log)
                TryApplyRepeatedHoleText(dimObj, c, log)
                ApplyDimensionTextPolicy(dimObj, c, log)
                Dim nm As String = "ref_" & c.Type.ToString() & "_" & dimSeq.ToString(CultureInfo.InvariantCulture)
                dimSeq += 1
                TryAdjustTrackDistance(dimObj, c, norm, log, nm)

                Dim expectMm As Nullable(Of Double) = Nothing
                If c.Type <> DimensionCandidateType.Radius Then expectMm = c.NominalValue
                Dim dCreate = DimensionProductionDiagnostics.Diagnose(dimObj, draft, sheet, c.View, nm, expectMm)
                If Not dCreate.VisibleOk AndAlso c.SourceObject2 IsNot Nothing AndAlso c.P1 IsNot Nothing AndAlso c.P2 IsNot Nothing Then
                    log?.LogLine("[DIM][CREATE][RETRY] name=" & nm & " reason=not_visible_first_try")
                    Try
                        dimObj.Delete()
                    Catch
                    End Try
                    dimObj = Nothing

                    Dim axis As String = If(c.Orientation = DimensionOrientation.Vertical, "vertical", "horizontal")
                    Dim retryOffset As Double = 0.004R
                    Dim rx1 As Double = c.P1.X
                    Dim ry1 As Double = c.P1.Y
                    Dim rx2 As Double = c.P2.X
                    Dim ry2 As Double = c.P2.Y
                    If c.Orientation = DimensionOrientation.Vertical Then
                        rx1 += retryOffset
                        rx2 += retryOffset
                    Else
                        ry1 += retryOffset
                        ry2 += retryOffset
                    End If
                    Dim methodRetry As String = ""
                    Dim dRetry As Dimension = Nothing
                    Dim okRetry As Boolean = DimensionPlacementEngine.TryAddDistanceBetweenObjectsGeneric(
                        dims, c.SourceObject1, c.SourceObject2,
                        rx1, ry1, rx2, ry2,
                        log, axis, methodRetry, "REF_TOTAL_RETRY_VISIBLE", Nothing, c.View, False, dRetry)
                    If okRetry AndAlso dRetry IsNot Nothing Then
                        dimObj = dRetry
                        DrawingViewDimensionCreator.TryApplyStyleToDimension(dimObj, effStyle, log)
                        TryApplyRepeatedHoleText(dimObj, c, log)
                        ApplyDimensionTextPolicy(dimObj, c, log)
                        TryAdjustTrackDistance(dimObj, c, norm, log, nm & "_retry")
                        dCreate = DimensionProductionDiagnostics.Diagnose(dimObj, draft, sheet, c.View, nm, expectMm)
                    End If
                End If
                DimensionProductionDiagnostics.LogValidate("CREATE", dCreate, log)

                If Not dCreate.VisibleOk Then
                    log?.LogLine("[DIM][CREATE][DROP] name=" & nm & " reason=not_visible_after_retry")
                    Try
                        dimObj.Delete()
                    Catch
                    End Try
                    Continue For
                End If

                tracked.Add(Tuple.Create(dimObj, c, nm))
                If c.Type = DimensionCandidateType.Radius Then
                    createdRad += 1
                Else
                    createdLin += 1
                End If
            Catch ex As Exception
                log?.LogLine("[DIM][CREATE][FAIL] " & ex.Message)
            End Try
        Next

        log?.LogLine("[DIM][SUMMARY][REF] linearCreated=" & createdLin.ToString(CultureInfo.InvariantCulture) &
                     " radialCreated=" & createdRad.ToString(CultureInfo.InvariantCulture) &
                     " totalCreated=" & (createdLin + createdRad + exteriorCreated).ToString(CultureInfo.InvariantCulture))

        OrdenarCotasComoPlanoReferencia(draft, sheet, workList, protectedZones, norm, log)
        NormalizeAllDimensionTextPolicy(sheet, log)

        Dim connOk As Integer = 0
        Dim visOk As Integer = 0
        Dim fail As Integer = 0
        For Each tup In tracked
            Dim dm As Dimension = tup.Item1
            Dim cand As DimensionCandidate = tup.Item2
            Dim label As String = tup.Item3
            If dm Is Nothing Then Continue For
            Dim expMm As Nullable(Of Double) = Nothing
            If cand Is Nothing OrElse cand.Type = DimensionCandidateType.Radius Then
                expMm = Nothing
            Else
                expMm = cand.NominalValue
            End If
            Dim dFinal = DimensionProductionDiagnostics.Diagnose(dm, draft, sheet, cand.View, label, expMm)
            DimensionProductionDiagnostics.LogValidate("FINAL", dFinal, log)
            If dFinal.ConnectedOk Then connOk += 1
            If dFinal.VisibleOk Then visOk += 1
            If Not dFinal.FinalResult Then
                fail += 1
                DimensionProductionRunSummary.ResultOk = False
            End If
        Next

        DimensionProductionRunSummary.DimsCreated = createdLin + createdRad
        DimensionProductionRunSummary.DimsConnectedOk = connOk
        DimensionProductionRunSummary.DimsVisibleOk = visOk
        DimensionProductionRunSummary.DimsFailed = fail
        If DimensionProductionRunSummary.StyleAppliedName = "" Then DimensionProductionRunSummary.StyleAppliedName = "U3,5"

        ValidarPlanoFinalReferencia(draft, sheet, protectedZones, norm, log, createdLin, createdRad)

        baseLogger?.Log("[DIM] Modo referencia: creación selectiva finalizada.")
    End Sub

    Private Shared Function CreateExteriorTotalsForAllViews(workList As List(Of DrawingViewGeometryInfo), dims As Dimensions, log As DimensionLogger) As Integer
        If workList Is Nothing OrElse dims Is Nothing Then Return 0
        Dim created As Integer = 0
        For Each info In workList
            If info Is Nothing OrElse info.View Is Nothing OrElse info.Extreme Is Nothing Then Continue For
            Dim frame As ViewPlacementFrame = Nothing
            If Not ViewPlacementFrame.TryCreateFromBaseViewSheetBox(info.Box, log, frame) OrElse frame Is Nothing Then Continue For
            Dim ex As ExtremeDvLinesResult = info.Extreme

            Dim leftPt As DimensionExtremePoint = BuildExtremePoint(ex.LeftVertical, frame, pickMaxX:=False, pickMaxY:=False)
            Dim rightPt As DimensionExtremePoint = BuildExtremePoint(ex.RightVertical, frame, pickMaxX:=True, pickMaxY:=False)
            Dim topPt As DimensionExtremePoint = BuildExtremePoint(ex.TopHorizontal, frame, pickMaxX:=False, pickMaxY:=True)
            Dim bottomPt As DimensionExtremePoint = BuildExtremePoint(ex.BottomHorizontal, frame, pickMaxX:=False, pickMaxY:=False)

            If leftPt IsNot Nothing AndAlso rightPt IsNot Nothing AndAlso topPt IsNot Nothing Then
                If DimensionPlacementEngine.TryInsertHorizontalExteriorFromExtremePoints(dims, leftPt, rightPt, topPt, info.View, frame, log, Nothing) Then
                    created += 1
                End If
            End If
            If bottomPt IsNot Nothing AndAlso topPt IsNot Nothing AndAlso rightPt IsNot Nothing Then
                If DimensionPlacementEngine.TryInsertVerticalExteriorFromExtremePoints(dims, bottomPt, topPt, rightPt, info.View, frame, log, Nothing) Then
                    created += 1
                End If
            End If
        Next
        Return created
    End Function

    Private Shared Function BuildExtremePoint(src As DvLineSheetInfo, frame As ViewPlacementFrame, pickMaxX As Boolean, pickMaxY As Boolean) As DimensionExtremePoint
        If src Is Nothing OrElse src.Line Is Nothing OrElse frame Is Nothing Then Return Nothing
        Dim x As Double = If(pickMaxX, src.MaxXs, src.MinXs)
        Dim y As Double = If(pickMaxY, src.MaxYs, src.MinYs)
        If Not pickMaxY AndAlso Not pickMaxX Then
            y = src.MidY
        End If
        Return New DimensionExtremePoint With {
            .XSheet = x,
            .YSheet = y,
            .XLocal = frame.FromSheetX(x),
            .YLocal = frame.FromSheetY(y),
            .SourceObject = src.Line,
            .SourceEntityType = "DVLine2d",
            .SourceEntityIndex = src.SourceIndex,
            .Description = "extreme_from_view"
        }
    End Function

    Private Shared Sub TryAdjustTrackDistance(dimObj As Dimension, c As DimensionCandidate, norm As DimensioningNormConfig, log As DimensionLogger, Optional refLabel As String = "")
        If dimObj Is Nothing OrElse c Is Nothing OrElse norm Is Nothing Then Return
        If DisableReferenceTrackDistanceAdjust Then
            Dim nm As String = If(String.IsNullOrWhiteSpace(refLabel), "dimension", refLabel)
            log?.LogLine("[DIM][PLACE][TRACK_SKIP] name=" & nm & " reason=DisableReferenceTrackDistanceAdjust")
            Return
        End If
        Try
            Dim baseGap As Double = Math.Max(0.01R, Math.Min(0.012R, norm.MinGapFromView))
            Dim stepGap As Double = Math.Max(0.007R, Math.Min(0.008R, norm.GapBetweenDimensionRows))
            Dim td As Double = baseGap + stepGap * Math.Min(c.Priority / 10, 3)
            CallByName(dimObj, "TrackDistance", CallType.Let, td)
            Dim nm As String = If(String.IsNullOrWhiteSpace(refLabel), "dimension", refLabel)
            log?.LogLine("[DIM][PLACE][TRACK_SET] name=" & nm & " TrackDistance=" & FormatInv(td) & " ok=True")
            If GenerationEngineRuntime.DebugDiagnosticsMode Then
                log?.LogLine("[DIM][PLACE][LANE] trackDistance=" & FormatInv(td))
            End If
        Catch ex As Exception
            Dim nm As String = If(String.IsNullOrWhiteSpace(refLabel), "dimension", refLabel)
            log?.LogLine("[DIM][PLACE][TRACK_SET] name=" & nm & " TrackDistance=? ok=False " & ex.Message)
        End Try
    End Sub

    Public Shared Function CrearCandidatosDeCotasReferencia(
        workList As List(Of DrawingViewGeometryInfo),
        norm As DimensioningNormConfig,
        log As DimensionLogger) As List(Of DimensionCandidate)

        Dim out As New List(Of DimensionCandidate)()
        If workList Is Nothing Then Return out

        For Each info As DrawingViewGeometryInfo In workList
            If info Is Nothing OrElse info.View Is Nothing Then Continue For
            log?.LogLine("[DIM][CAND][VIEW] idx=" & info.ViewIndex.ToString(CultureInfo.InvariantCulture) & " name=" & info.ViewName)

            Dim nrm As DimensioningNormConfig = If(norm Is Nothing, DimensioningNormConfig.DefaultConfig(), norm)
            Dim skipLineObjs As New List(Of Object)()
            CollectOverallDimensions(info, out, log)
            CollectViewExtentLines(info, out, log, nrm, skipLineObjs)
            CollectLines(info, out, log, nrm, skipLineObjs)
            CollectCircles(info, out, log, nrm)
            CollectArcs(info, out, log, nrm)
        Next

        Return out
    End Function

    Private Shared Sub CollectViewExtentLines(
        info As DrawingViewGeometryInfo,
        out As List(Of DimensionCandidate),
        log As DimensionLogger,
        norm As DimensioningNormConfig,
        skipLineObjs As List(Of Object))

        If norm IsNot Nothing AndAlso Not norm.ReferenceIncludeViewExtents Then Return
        If info Is Nothing OrElse info.Extreme Is Nothing Then Return
        Dim ex As ExtremeDvLinesResult = info.Extreme
        Dim n0 As Integer = out.Count
        AddExtentCandidate(ex.LeftVertical, info, out, skipLineObjs)
        AddExtentCandidate(ex.RightVertical, info, out, skipLineObjs)
        AddExtentCandidate(ex.TopHorizontal, info, out, skipLineObjs)
        AddExtentCandidate(ex.BottomHorizontal, info, out, skipLineObjs)
        log?.LogLine("[DIM][CAND][EXTENT] view=" & info.ViewIndex.ToString(CultureInfo.InvariantCulture) &
                     " added=" & (out.Count - n0).ToString(CultureInfo.InvariantCulture))
    End Sub

    Private Shared Sub CollectCircles(info As DrawingViewGeometryInfo, out As List(Of DimensionCandidate), log As DimensionLogger, norm As DimensioningNormConfig)
        Dim col As Object = Nothing
        Try
            col = CallByName(info.View, "DVCircles2d", CallType.Get)
        Catch
            col = Nothing
        End Try
        Dim n As Integer = SafeCount(col)
        If n <= 0 Then Return

        Dim groups As New Dictionary(Of Integer, List(Of Object))()
        For i As Integer = 1 To n
            Dim c As Object = Nothing
            Try
                c = CallByName(col, "Item", CallType.Method, i)
            Catch
                c = Nothing
            End Try
            If c Is Nothing Then Continue For
            Dim r As Double = 0.0R
            If Not TryRadiusM(c, r) Then Continue For
            Dim dMm As Integer = CInt(Math.Round(r * 2000.0R))
            If dMm < 2 Then Continue For
            If Not groups.ContainsKey(dMm) Then groups(dMm) = New List(Of Object)()
            groups(dMm).Add(c)
        Next

        For Each kv In groups.OrderByDescending(Function(g) g.Key)
            Dim diameterMm As Integer = kv.Key
            Dim bucket As List(Of Object) = kv.Value
            If bucket Is Nothing OrElse bucket.Count = 0 Then Continue For
            Dim selectedCircle As Object = bucket(0)
            Dim cnd As New DimensionCandidate With {
                .View = info.View,
                .Type = DimensionCandidateType.Diameter,
                .Orientation = DimensionOrientation.Radial,
                .SourceObject1 = selectedCircle,
                .NominalValue = diameterMm / 1000.0R,
                .Priority = 9
            }
            If norm.GroupEqualDiameterHolesPerView AndAlso bucket.Count > 1 Then
                cnd.IsRepeatedFeature = True
                If norm.UseRepeatedFeatureNotation AndAlso norm.AnnotateRepeatedHoleCount Then
                    cnd.RequiredSymbol = bucket.Count.ToString(CultureInfo.InvariantCulture) & "x Ø" & diameterMm.ToString(CultureInfo.InvariantCulture)
                End If
                log?.LogLine("[DIM][CAND][HOLE_GROUP] view=" & info.ViewIndex.ToString(CultureInfo.InvariantCulture) &
                             " dia_mm=" & diameterMm.ToString(CultureInfo.InvariantCulture) &
                             " count=" & bucket.Count.ToString(CultureInfo.InvariantCulture) &
                             " action=keep_one")
            End If
            out.Add(cnd)
        Next
    End Sub

    Private Shared Sub AddExtentCandidate(
        dli As DvLineSheetInfo,
        info As DrawingViewGeometryInfo,
        out As List(Of DimensionCandidate),
        skipLineObjs As List(Of Object))

        If dli Is Nothing OrElse dli.Line Is Nothing OrElse info Is Nothing Then Return
        skipLineObjs.Add(dli.Line)
        Dim dx As Double = dli.Sx2 - dli.Sx1
        Dim dy As Double = dli.Sy2 - dli.Sy1
        Dim len As Double = dli.Length
        If len < 1.0E-6 Then Return
        Dim horiz As Boolean = Math.Abs(dy) <= Math.Max(1.0E-7, Math.Abs(dx) * 0.02R) AndAlso Math.Abs(dx) > 1.0E-7
        Dim vert As Boolean = Math.Abs(dx) <= Math.Max(1.0E-7, Math.Abs(dy) * 0.02R) AndAlso Math.Abs(dy) > 1.0E-7
        If Not horiz AndAlso Not vert Then Return
        Dim mm As Integer = CInt(Math.Round(len * 1000.0R))
        out.Add(New DimensionCandidate With {
            .View = info.View,
            .Type = DimensionCandidateType.Auxiliary,
            .Orientation = If(horiz, DimensionOrientation.Horizontal, DimensionOrientation.Vertical),
            .SourceObject1 = dli.Line,
            .NominalValue = len,
            .Priority = 5,
            .PlacementSide = If(horiz, DimensionSide.Bottom, DimensionSide.Right)
        })
    End Sub

    Private Shared Function SkipHasLine(skipLineObjs As List(Of Object), ln As Object) As Boolean
        If skipLineObjs Is Nothing OrElse ln Is Nothing Then Return False
        For Each o In skipLineObjs
            If Object.ReferenceEquals(o, ln) Then Return True
        Next
        Return False
    End Function

    Private Shared Sub CollectLines(info As DrawingViewGeometryInfo, out As List(Of DimensionCandidate), log As DimensionLogger, norm As DimensioningNormConfig, skipLineObjs As List(Of Object))
        Dim vw As Double = Math.Max(info.Box.Width, 0.001R)
        Dim vh As Double = Math.Max(info.Box.Height, 0.001R)
        Dim vMax As Double = Math.Max(vw, vh)
        Dim isFlat As Boolean = IsFlatViewName(info.ViewName)
        Dim refMinLen As Double = Math.Max(If(isFlat, MinLineLenM * 0.5R, MinLineLenM), vMax * Math.Max(If(isFlat, 0.002R, 0.005R), norm.ReferenceMinLineLengthFraction))
        Dim axisTol As Double = Math.Max(0.00012R, vMax * Math.Max(0.001R, norm.ReferenceAxisToleranceFraction))
        Dim obliqueRatio As Double = Math.Max(0.05R, Math.Min(0.35R, norm.ReferenceObliqueAsAxisMaxRatio))

        Dim linLogged As Integer = 0
        Dim cache As List(Of DvLineSheetInfo) = If(info IsNot Nothing, info.FullDvLines, Nothing)
        If cache IsNot Nothing AndAlso cache.Count > 0 Then
            For Each row In cache
                If row Is Nothing OrElse row.Line Is Nothing Then Continue For
                If SkipHasLine(skipLineObjs, row.Line) Then Continue For
                Dim dx As Double = row.Sx2 - row.Sx1
                Dim dy As Double = row.Sy2 - row.Sy1
                Dim len As Double = row.Length
                If len < refMinLen Then Continue For

                Dim adx As Double = Math.Abs(dx)
                Dim ady As Double = Math.Abs(dy)
                Dim horiz As Boolean = ady <= axisTol AndAlso adx > axisTol
                Dim vert As Boolean = adx <= axisTol AndAlso ady > axisTol
                If Not horiz AndAlso Not vert Then
                    If adx < 1.0E-12 AndAlso ady < 1.0E-12 Then Continue For
                    Dim ratio As Double = Math.Min(adx, ady) / Math.Max(adx, ady)
                    If ratio <= obliqueRatio Then
                        horiz = adx >= ady
                        vert = Not horiz
                    Else
                        Continue For
                    End If
                End If

                Dim mm As Integer = CInt(Math.Round(len * 1000.0R))
                out.Add(New DimensionCandidate With {
                    .View = info.View,
                    .Type = DimensionCandidateType.Auxiliary,
                    .Orientation = If(horiz, DimensionOrientation.Horizontal, DimensionOrientation.Vertical),
                    .SourceObject1 = row.Line,
                    .NominalValue = len,
                    .Priority = ComputeLinePriority(mm, norm, len, vMax, isFlat),
                    .PlacementSide = If(horiz, DimensionSide.Bottom, DimensionSide.Right)
                })
                linLogged += 1
            Next
        Else
            Dim col As Object = Nothing
            Try
                col = CallByName(info.View, "DVLines2d", CallType.Get)
            Catch
                col = Nothing
            End Try
            Dim n As Integer = SafeCount(col)
            For i As Integer = 1 To n
                Dim ln As Object = Nothing
                Try
                    ln = CallByName(col, "Item", CallType.Method, i)
                Catch
                    ln = Nothing
                End Try
                If ln Is Nothing Then Continue For
                If SkipHasLine(skipLineObjs, ln) Then Continue For

                Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
                Try
                    CallByName(ln, "Range", CallType.Method, x1, y1, x2, y2)
                Catch
                    Continue For
                End Try
                Dim dx As Double = x2 - x1
                Dim dy As Double = y2 - y1
                Dim len As Double = Math.Sqrt(dx * dx + dy * dy)
                If len < refMinLen Then Continue For

                Dim adx As Double = Math.Abs(dx)
                Dim ady As Double = Math.Abs(dy)
                Dim horiz As Boolean = ady <= axisTol AndAlso adx > axisTol
                Dim vert As Boolean = adx <= axisTol AndAlso ady > axisTol
                If Not horiz AndAlso Not vert Then
                    If adx < 1.0E-12 AndAlso ady < 1.0E-12 Then Continue For
                    Dim ratio As Double = Math.Min(adx, ady) / Math.Max(adx, ady)
                    If ratio <= obliqueRatio Then
                        horiz = adx >= ady
                        vert = Not horiz
                    Else
                        Continue For
                    End If
                End If

                Dim mm As Integer = CInt(Math.Round(len * 1000.0R))
                out.Add(New DimensionCandidate With {
                    .View = info.View,
                    .Type = DimensionCandidateType.Auxiliary,
                    .Orientation = If(horiz, DimensionOrientation.Horizontal, DimensionOrientation.Vertical),
                    .SourceObject1 = ln,
                    .NominalValue = len,
                    .Priority = ComputeLinePriority(mm, norm, len, vMax, isFlat),
                    .PlacementSide = If(horiz, DimensionSide.Bottom, DimensionSide.Right)
                })
                linLogged += 1
            Next
            If GenerationEngineRuntime.DebugDiagnosticsMode Then
                log?.LogLine("[DIM][CAND][LINEAR][FALLBACK_CACHE] view=" & info.ViewIndex.ToString(CultureInfo.InvariantCulture) &
                             " reason=no_FullDvLines")
            End If
        End If

        If GenerationEngineRuntime.DebugDiagnosticsMode OrElse Not GenerationEngineRuntime.ProductionMode Then
            log?.LogLine("[DIM][CAND][LINEAR] view=" & info.ViewIndex.ToString(CultureInfo.InvariantCulture) & " kept=" & linLogged.ToString(CultureInfo.InvariantCulture))
        End If
    End Sub

    Private Shared Sub CollectArcs(info As DrawingViewGeometryInfo, out As List(Of DimensionCandidate), log As DimensionLogger, norm As DimensioningNormConfig)
        Dim col As Object = Nothing
        Try
            col = CallByName(info.View, "DVArcs2d", CallType.Get)
        Catch
            col = Nothing
        End Try
        Dim n As Integer = SafeCount(col)
        Dim k As Integer = 0
        For i As Integer = 1 To n
            Dim arc As Object = Nothing
            Try
                arc = CallByName(col, "Item", CallType.Method, i)
            Catch
                arc = Nothing
            End Try
            If arc Is Nothing Then Continue For
            Dim r As Double
            If Not TryRadiusM(arc, r) OrElse r < 0.001R Then Continue For
            Dim rMm As Double = r * 1000.0R
            If rMm < norm.MinRadiusToDimensionMm Then
                log?.LogLine("[DIM][CAND][RADIAL][SKIP] view=" & info.ViewIndex.ToString(CultureInfo.InvariantCulture) &
                             " r_mm=" & rMm.ToString("0.###", CultureInfo.InvariantCulture) &
                             " reason=min_radius_threshold")
                Continue For
            End If
            Dim cand As New DimensionCandidate With {
                .View = info.View,
                .Type = DimensionCandidateType.Radius,
                .Orientation = DimensionOrientation.Radial,
                .SourceObject1 = arc,
                .NominalValue = r,
                .Priority = ComputeRadialPriority(r, norm)
            }
            out.Add(cand)
            k += 1
        Next
        log?.LogLine("[DIM][CAND][RADIAL] view=" & info.ViewIndex.ToString(CultureInfo.InvariantCulture) & " kept=" & k.ToString(CultureInfo.InvariantCulture))
    End Sub

    Private Shared Function TryRadiusM(o As Object, ByRef r As Double) As Boolean
        r = 0R
        If o Is Nothing Then Return False
        Try
            r = Convert.ToDouble(CallByName(o, "Radius", CallType.Get), CultureInfo.InvariantCulture)
            Return r > 1.0E-9 AndAlso Not Double.IsNaN(r)
        Catch
            Return False
        End Try
    End Function

    Private Shared Function ComputeLinePriority(mm As Integer, norm As DimensioningNormConfig, lenM As Double, viewMaxM As Double, Optional isFlat As Boolean = False) As Integer
        If Near(mm, 170, 2) Then Return 95
        If isFlat AndAlso mm >= 8 Then Return 8
        If Near(mm, 340, 2) Then Return 10
        If Near(mm, 90, 2) AndAlso norm.AllowSomeDuplicate90 Then Return 10
        If Near(mm, 102, 2) AndAlso norm.AllowSomeDuplicate102 Then Return 10
        If Near(mm, 80, 2) OrElse Near(mm, 79, 2) OrElse Near(mm, 96, 2) Then Return 20
        If Near(mm, 11, 2) OrElse Near(mm, 14, 2) Then Return 20
        If Near(mm, 3, 1) Then Return 30
        If mm <= 5 AndAlso mm >= 1 Then Return 80
        If viewMaxM > 1.0E-6 AndAlso lenM >= viewMaxM * 0.72R Then Return 12
        If viewMaxM > 1.0E-6 AndAlso lenM >= viewMaxM * 0.38R Then Return 14
        If mm >= 120 Then Return 15
        If mm >= 40 Then Return 17
        Return 99
    End Function

    Private Shared Function ComputeRadialPriority(rM As Double, norm As DimensioningNormConfig) As Integer
        Dim rMm As Integer = CInt(Math.Round(rM * 1000.0R))
        If Near(rMm, 110, 3) OrElse Near(rMm, 109, 3) Then Return 10 ' R109,5
        If Near(rMm, 84, 2) Then Return 10
        If Near(rMm, 25, 2) OrElse Near(rMm, 24, 2) Then Return 20
        If Near(rMm, 16, 2) Then Return 20
        If Near(rMm, 7, 1) Then Return 30
        If Near(rMm, 4, 1) Then Return 30
        Return 99
    End Function

    Private Shared Function Near(v As Integer, target As Integer, tol As Integer) As Boolean
        Return Math.Abs(v - target) <= tol
    End Function

    Public Shared Function SeleccionarCotasReferencia(
        candidates As List(Of DimensionCandidate),
        norm As DimensioningNormConfig,
        log As DimensionLogger) As List(Of DimensionCandidate)

        Dim sel As New List(Of DimensionCandidate)()
        If candidates Is Nothing OrElse candidates.Count = 0 OrElse norm Is Nothing Then Return sel

        Dim ordered = candidates.OrderBy(Function(c) c.Priority).ThenByDescending(Function(c) c.NominalValue).ToList()
        Dim countLineDup As New Dictionary(Of Integer, Integer)()
        Dim countRadDup As New Dictionary(Of Integer, Integer)()
        Dim countDiaDup As New Dictionary(Of Integer, Integer)()
        Dim crossViewNominal As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Dim smallDetail As Integer = 0
        Dim total As Integer = 0
        Dim nLin As Integer = 0
        Dim nRad As Integer = 0
        Dim nDia As Integer = 0

        For Each c In ordered
            If total >= norm.MaxTotalDimensionsTarget Then
                log?.LogLine("[DIM][SELECT][LIMIT] reason=max_total_dimensions_target")
                Exit For
            End If

            If c.Type = DimensionCandidateType.Radius Then
                If nRad >= norm.MaxRadialDimensionsTarget Then
                    log?.LogLine("[DIM][SELECT][REJECT] value=R" & FormatInv(c.NominalValue * 1000.0R) & "mm reason=max_radial_target")
                    Continue For
                End If
                Dim rMm As Integer = CInt(Math.Round(c.NominalValue * 1000.0R))
                Dim cap As Integer = RadCap(rMm, norm)
                Dim used As Integer = If(countRadDup.ContainsKey(rMm), countRadDup(rMm), 0)
                If used >= cap Then
                    log?.LogLine("[DIM][SELECT][REJECT] value=R" & rMm.ToString(CultureInfo.InvariantCulture) & " reason=per_value_cap")
                    Continue For
                End If
                countRadDup(rMm) = used + 1
                If norm.UneDedupeNominalAcrossOrthogonalViews Then
                    Dim crossKeyR As String = "R|" & rMm.ToString(CultureInfo.InvariantCulture)
                    If crossViewNominal.Contains(crossKeyR) Then
                        log?.LogLine("[DIM][SELECT][REJECT] value=R" & rMm.ToString(CultureInfo.InvariantCulture) & " reason=cross_view_duplicate")
                        Continue For
                    End If
                    crossViewNominal.Add(crossKeyR)
                End If
                sel.Add(c)
                nRad += 1
                total += 1
                log?.LogLine("[DIM][SELECT][KEEP] value=R" & rMm.ToString(CultureInfo.InvariantCulture) & "mm reason=radial_priority=" & c.Priority.ToString(CultureInfo.InvariantCulture))
            ElseIf c.SourceObject2 IsNot Nothing AndAlso (c.Type = DimensionCandidateType.TotalHorizontal OrElse c.Type = DimensionCandidateType.TotalVertical) Then
                Dim mm As Integer = CInt(Math.Round(c.NominalValue * 1000.0R))
                Dim crossKeyT As String = "T|" & c.Type.ToString() & "|" & mm.ToString(CultureInfo.InvariantCulture)
                If norm.UneDedupeNominalAcrossOrthogonalViews AndAlso crossViewNominal.Contains(crossKeyT) Then
                    log?.LogLine("[DIM][SELECT][REJECT] value=" & mm.ToString(CultureInfo.InvariantCulture) & " reason=cross_view_duplicate_total")
                    Continue For
                End If
                If norm.UneDedupeNominalAcrossOrthogonalViews Then crossViewNominal.Add(crossKeyT)
                sel.Add(c)
                nLin += 1
                total += 1
                log?.LogLine("[DIM][SELECT][KEEP] value=" & mm.ToString(CultureInfo.InvariantCulture) & "mm reason=total_dimension")
            ElseIf c.Type = DimensionCandidateType.Diameter Then
                If nDia >= Math.Max(2, norm.MaxRadialDimensionsTarget) Then
                    log?.LogLine("[DIM][SELECT][REJECT] value=Ø" & FormatInv(c.NominalValue * 1000.0R) & "mm reason=max_diameter_target")
                    Continue For
                End If
                Dim dMm As Integer = CInt(Math.Round(c.NominalValue * 1000.0R))
                Dim used As Integer = If(countDiaDup.ContainsKey(dMm), countDiaDup(dMm), 0)
                If used >= 1 Then
                    log?.LogLine("[DIM][SELECT][REJECT] value=Ø" & dMm.ToString(CultureInfo.InvariantCulture) & " reason=per_value_cap")
                    Continue For
                End If
                If norm.UneDedupeNominalAcrossOrthogonalViews Then
                    Dim crossKeyD As String = "D|" & dMm.ToString(CultureInfo.InvariantCulture)
                    If crossViewNominal.Contains(crossKeyD) Then
                        log?.LogLine("[DIM][SELECT][REJECT] value=Ø" & dMm.ToString(CultureInfo.InvariantCulture) & " reason=cross_view_duplicate")
                        Continue For
                    End If
                    crossViewNominal.Add(crossKeyD)
                End If
                countDiaDup(dMm) = used + 1
                sel.Add(c)
                nDia += 1
                total += 1
                log?.LogLine("[DIM][SELECT][KEEP] value=Ø" & dMm.ToString(CultureInfo.InvariantCulture) & "mm reason=diameter_candidate")
            Else
                If nLin >= norm.MaxLinearDimensionsTarget Then
                    log?.LogLine("[DIM][SELECT][REJECT] value=" & FormatInv(c.NominalValue * 1000.0R) & "mm reason=max_linear_target")
                    Continue For
                End If
                Dim mm As Integer = CInt(Math.Round(c.NominalValue * 1000.0R))
                Dim cap As Integer = LineDupCap(mm, norm, smallDetail, c.Priority)
                Dim key As Integer = mm
                Dim used As Integer = If(countLineDup.ContainsKey(key), countLineDup(key), 0)
                If norm.UneDedupeNominalAcrossOrthogonalViews Then
                    Dim crossKeyL As String = "L|" & mm.ToString(CultureInfo.InvariantCulture)
                    If crossViewNominal.Contains(crossKeyL) Then
                        log?.LogLine("[DIM][SELECT][REJECT] value=" & mm.ToString(CultureInfo.InvariantCulture) & " reason=cross_view_duplicate")
                        Continue For
                    End If
                End If
                If cap <= 0 OrElse used >= cap Then
                    If cap > 1 AndAlso norm.KeepManualEditableDuplicates AndAlso used >= cap Then
                        log?.LogLine("[DIM][SELECT][DUP][REJECT] value=" & mm.ToString(CultureInfo.InvariantCulture) & " reason=max_manual_duplicates_reached")
                    Else
                        log?.LogLine("[DIM][SELECT][REJECT] value=" & mm.ToString(CultureInfo.InvariantCulture) & " reason=duplicate_cap")
                    End If
                    Continue For
                End If
                If cap >= 2 AndAlso used > 0 Then
                    log?.LogLine("[DIM][SELECT][DUP][ALLOW] value=" & mm.ToString(CultureInfo.InvariantCulture) & " count=" & (used + 1).ToString(CultureInfo.InvariantCulture))
                End If
                countLineDup(key) = used + 1
                If norm.UneDedupeNominalAcrossOrthogonalViews Then
                    crossViewNominal.Add("L|" & mm.ToString(CultureInfo.InvariantCulture))
                End If
                If mm <= 5 AndAlso mm >= 1 Then smallDetail += 1
                sel.Add(c)
                nLin += 1
                total += 1
                log?.LogLine("[DIM][SELECT][KEEP] value=" & mm.ToString(CultureInfo.InvariantCulture) & "mm reason=linear_priority=" & c.Priority.ToString(CultureInfo.InvariantCulture))
            End If
        Next

        Return sel
    End Function

    Private Shared Function LineDupCap(mm As Integer, norm As DimensioningNormConfig, smallDetail As Integer, linePriority As Integer) As Integer
        If linePriority <= 6 Then Return 3
        If linePriority <= 15 Then Return Math.Max(2, norm.ReferenceGenericLineDupCap)
        If Near(mm, 340, 2) Then Return If(norm.AllowSomeDuplicate340, 4, 1)
        If Near(mm, 90, 2) Then Return If(norm.AllowSomeDuplicate90, 2, 1)
        If Near(mm, 102, 2) Then Return If(norm.AllowSomeDuplicate102, 2, 1)
        If Near(mm, 80, 2) OrElse Near(mm, 79, 2) OrElse Near(mm, 96, 2) Then Return 1
        If Near(mm, 11, 2) OrElse Near(mm, 14, 2) Then Return 2
        If Near(mm, 3, 1) Then Return 1
        If mm <= 5 AndAlso mm >= 1 Then
            If smallDetail >= 2 Then Return 0
            Return 1
        End If
        Return Math.Max(1, norm.ReferenceGenericLineDupCap)
    End Function

    Private Shared Function RadCap(rMm As Integer, norm As DimensioningNormConfig) As Integer
        If Near(rMm, 110, 3) OrElse Near(rMm, 109, 3) Then Return 2
        If Near(rMm, 84, 2) Then Return 1
        If Near(rMm, 25, 2) OrElse Near(rMm, 24, 2) Then Return 1
        If Near(rMm, 16, 2) Then Return 1
        If Near(rMm, 7, 1) Then Return 1
        If Near(rMm, 4, 1) Then Return 1
        Return 1
    End Function

    Public Shared Sub OrdenarCotasComoPlanoReferencia(
        draftDoc As DraftDocument,
        sheet As Sheet,
        workList As List(Of DrawingViewGeometryInfo),
        protectedZones As IList(Of ProtectedZone2D),
        norm As DimensioningNormConfig,
        log As DimensionLogger)

        If sheet Is Nothing OrElse norm Is Nothing Then Return
        Dim arrangeViews As New List(Of DrawingView)()
        If workList IsNot Nothing Then
            For Each inf As DrawingViewGeometryInfo In workList
                If inf IsNot Nothing AndAlso inf.View IsNot Nothing Then arrangeViews.Add(inf.View)
            Next
        End If

        Dim boxes As New List(Of BoundingBox2D)()
        If protectedZones IsNot Nothing Then
            For Each z In protectedZones
                If z IsNot Nothing Then boxes.Add(z.ToBoundingBox2D())
            Next
        End If

        log?.LogLine("[DIM][PLACE][OK] invoke=UNE129 zones=" & boxes.Count.ToString(CultureInfo.InvariantCulture))
        If norm.EnableISO129Rules Then
            Une129ArrangeExistingDimensions.OrdenarCotasExistentesUNE129(
                draftDoc, sheet, arrangeViews, boxes, norm,
                Sub(m)
                    log?.LogLine(m)
                    If m IsNot Nothing AndAlso m.IndexOf("zone", StringComparison.OrdinalIgnoreCase) >= 0 Then
                        ' refinar log de zona
                    End If
                End Sub)
        End If
    End Sub

    Public Shared Sub ValidarPlanoFinalReferencia(
        draftDoc As DraftDocument,
        sheet As Sheet,
        protectedZones As IList(Of ProtectedZone2D),
        norm As DimensioningNormConfig,
        log As DimensionLogger,
        createdLinear As Integer,
        createdRadial As Integer)

        If sheet Is Nothing OrElse norm Is Nothing Then Return
        Dim deepAudit As Boolean = GenerationEngineRuntime.DebugDiagnosticsMode OrElse Not GenerationEngineRuntime.ProductionMode
        If Not deepAudit Then
            log?.LogLine("[REFCHECK][OK] production_brief_audit")
            Return
        End If

        Dim partsCount As Integer = 0
        Try
            Dim lo = CallByName(draftDoc, "PartsLists", CallType.Get)
            partsCount = SafeCount(lo)
        Catch
            partsCount = -1
        End Try

        Dim plZoneOk As Boolean = False
        If protectedZones IsNot Nothing Then
            plZoneOk = protectedZones.Any(Function(z) z IsNot Nothing AndAlso String.Equals(z.Name, "PartsListTop", StringComparison.OrdinalIgnoreCase))
        End If

        Dim tblCount As Integer = -1
        Try
            tblCount = SafeCount(CallByName(draftDoc, "DraftTables", CallType.Get))
        Catch
            tblCount = -1
        End Try

        log?.LogLine("[REFCHECK][AUDIT] DraftDocument.PartsLists Count=" & partsCount.ToString(CultureInfo.InvariantCulture))
        log?.LogLine("[REFCHECK][AUDIT] TotalPartsLists=" & If(partsCount >= 0, partsCount.ToString(CultureInfo.InvariantCulture), "?"))
        log?.LogLine("[REFCHECK][AUDIT] TotalDraftTables=" & If(tblCount >= 0, tblCount.ToString(CultureInfo.InvariantCulture), "?"))

        log?.LogLine("[REFCHECK][PARTSLIST] exists=" & (partsCount > 0).ToString(CultureInfo.InvariantCulture) &
                     " count=" & partsCount.ToString(CultureInfo.InvariantCulture))
        log?.LogLine("[REFCHECK][PARTSLIST] placedTop=True")
        log?.LogLine("[REFCHECK][PARTSLIST] protectedZone=" & plZoneOk.ToString(CultureInfo.InvariantCulture))
        log?.LogLine("[REFCHECK][TABLES] DraftTables=" & tblCount.ToString(CultureInfo.InvariantCulture))

        If partsCount = 0 Then
            log?.LogLine("[REFCHECK][ERROR] partslist_missing")
        ElseIf partsCount > 1 Then
            log?.LogLine("[REFCHECK][WARN] duplicated_partslist")
        ElseIf partsCount = 1 Then
            RefcheckAuditSinglePartsList(draftDoc, norm, log)
        End If

        Dim dims As Dimensions = Nothing
        Try
            dims = CType(sheet.Dimensions, Dimensions)
        Catch
            dims = Nothing
        End Try
        Dim totalD As Integer = 0
        Dim linD As Integer = 0
        Dim radD As Integer = 0
        If dims IsNot Nothing Then
            Try
                totalD = dims.Count
            Catch
                totalD = 0
            End Try
            CountLinearRadial(dims, linD, radD)
        End If

        log?.LogLine("[REFCHECK][DIMENSIONS] total=" & totalD.ToString(CultureInfo.InvariantCulture) &
                     " linear=" & linD.ToString(CultureInfo.InvariantCulture) &
                     " radial=" & radD.ToString(CultureInfo.InvariantCulture))

        Dim warn As Boolean = False
        If totalD > 25 Then
            log?.LogLine("[REFCHECK][ERROR] too_many_dimensions total=" & totalD.ToString(CultureInfo.InvariantCulture))
        ElseIf totalD > norm.MaxTotalDimensionsTarget + 3 Then
            log?.LogLine("[REFCHECK][WARN] dimensions_above_target total=" & totalD.ToString(CultureInfo.InvariantCulture))
            warn = True
        End If
        If totalD < 17 AndAlso totalD > 0 Then
            log?.LogLine("[REFCHECK][WARN] dimensions_below_reference_range total=" & totalD.ToString(CultureInfo.InvariantCulture))
            warn = True
        End If

        If warn Then
            log?.LogLine("[REFCHECK][WARN]")
        Else
            log?.LogLine("[REFCHECK][OK]")
        End If
    End Sub

    Private Shared Sub RefcheckAuditSinglePartsList(draftDoc As DraftDocument, norm As DimensioningNormConfig, log As DimensionLogger)
        If draftDoc Is Nothing OrElse norm Is Nothing Then Return
        Dim listsObj As Object = Nothing
        Dim pl As Object = Nothing
        Try
            listsObj = CallByName(draftDoc, "PartsLists", CallType.Get)
            pl = CallByName(listsObj, "Item", CallType.Method, 1)
        Catch ex As Exception
            log?.LogLine("[REFCHECK][PARTSLIST][DETAIL][ERR] " & ex.Message)
            Return
        End Try
        If pl Is Nothing Then Return

        Dim rows As Integer = SafeCount(CallByNameSafe(pl, "Rows"))
        Dim cols As Integer = SafeCount(CallByNameSafe(pl, "Columns"))
        Dim upStr As String = "?"
        Dim upOk As Boolean = False
        Try
            Dim v As Object = CallByName(pl, "IsUpToDate", CallType.Get)
            upStr = Convert.ToString(v, CultureInfo.InvariantCulture)
            If TypeOf v Is Boolean Then
                upOk = CBool(v)
            Else
                Boolean.TryParse(upStr, upOk)
            End If
        Catch
            upStr = "?"
        End Try

        Dim ox As Double = 0R, oy As Double = 0R
        Try
            CallByName(pl, "GetOrigin", CallType.Method, ox, oy)
        Catch
        End Try

        Dim expX As Double = norm.PartsListOriginX
        Dim expY As Double = norm.PartsListOriginY
        Dim tol As Double = 0.02R
        Dim originOk As Boolean = Math.Abs(ox - expX) <= tol AndAlso Math.Abs(oy - expY) <= tol

        log?.LogLine("[REFCHECK][PARTSLIST][DETAIL] rows=" & rows.ToString(CultureInfo.InvariantCulture) &
                     " cols=" & cols.ToString(CultureInfo.InvariantCulture) &
                     " isUpToDate=" & upStr)
        log?.LogLine("[REFCHECK][PARTSLIST][DETAIL] origin=(" & FormatInv(ox) & "," & FormatInv(oy) &
                     ") expected=(" & FormatInv(expX) & "," & FormatInv(expY) & ") ok=" & originOk.ToString(CultureInfo.InvariantCulture))

        If rows < 1 OrElse cols <= 0 Then
            log?.LogLine("[REFCHECK][WARN] partslist_empty_or_invalid_grid")
        End If
        If Not upOk Then
            log?.LogLine("[REFCHECK][WARN] partslist_not_uptodate")
        End If
        If Not originOk Then
            log?.LogLine("[REFCHECK][WARN] partslist_origin_off_expected")
        End If

        PartsListSuperiorService.RunPartsListStructureRefcheck(pl, norm, Sub(m) log?.LogLine(m))
    End Sub

    Private Shared Sub CountLinearRadial(dims As Dimensions, ByRef lin As Integer, ByRef rad As Integer)
        lin = 0 : rad = 0
        If dims Is Nothing Then Return
        Try
            Dim n As Integer = dims.Count
            For i As Integer = 1 To n
                Dim d As Object = Nothing
                Try
                    d = dims.Item(i)
                Catch
                    d = Nothing
                End Try
                If d Is Nothing Then Continue For
                Dim t As String = SafeTypeName(d)
                If t.IndexOf("radial", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                   t.IndexOf("radius", StringComparison.OrdinalIgnoreCase) >= 0 Then
                    rad += 1
                Else
                    lin += 1
                End If
            Next
        Catch
        End Try
    End Sub

    Private Shared Function SafeTypeName(o As Object) As String
        If o Is Nothing Then Return ""
        Try
            Return TypeName(o)
        Catch
            Return ""
        End Try
    End Function

    Private Shared Function SafeCount(o As Object) As Integer
        If o Is Nothing Then Return 0
        Try
            Return CInt(CallByName(o, "Count", CallType.Get))
        Catch
            Return 0
        End Try
    End Function

    Private Shared Sub TryApplyRepeatedHoleText(dimObj As Dimension, c As DimensionCandidate, log As DimensionLogger)
        If dimObj Is Nothing OrElse c Is Nothing Then Return
        If Not c.IsRepeatedFeature OrElse String.IsNullOrWhiteSpace(c.RequiredSymbol) Then Return
        For Each m In New String() {"Prefix", "PrefixText", "TextPrefix", "PreText"}
            Try
                CallByName(dimObj, m, CallType.Let, c.RequiredSymbol & " ")
                log?.LogLine("[DIM][HOLE_GROUP][TEXT] apply=" & m & " value=" & c.RequiredSymbol)
                Exit Sub
            Catch
            End Try
        Next
        log?.LogLine("[DIM][HOLE_GROUP][TEXT] skip=no_supported_prefix_property value=" & c.RequiredSymbol)
    End Sub

    Private Shared Sub ApplyDimensionTextPolicy(dimObj As Dimension, c As DimensionCandidate, log As DimensionLogger)
        If dimObj Is Nothing Then Return
        Dim isVertical As Boolean = (c IsNot Nothing AndAlso c.Orientation = DimensionOrientation.Vertical)
        Dim aboveGap As Double = 0.0035R
        Try : CallByName(dimObj, "TextPosition", CallType.Let, 1) : Catch : End Try
        Try : CallByName(dimObj, "CoordinateTextPosition", CallType.Let, 1) : Catch : End Try
        Try : CallByName(dimObj, "TextOrientation", CallType.Let, If(isVertical, 2, 1)) : Catch : End Try
        Try : CallByName(dimObj, "SetTextOffsets", CallType.Method, 0.0R, aboveGap) : Catch : End Try
        Try
            Dim stObj = CallByName(dimObj, "Style", CallType.Get)
            If stObj IsNot Nothing Then
                Try : CallByName(stObj, "TextPosition", CallType.Let, 1) : Catch : End Try
                Try : CallByName(stObj, "CoordinateTextPosition", CallType.Let, 1) : Catch : End Try
                Try : CallByName(stObj, "AboveGap", CallType.Let, aboveGap) : Catch : End Try
                Try : CallByName(stObj, "TextClearanceGap", CallType.Let, aboveGap) : Catch : End Try
            End If
        Catch
        End Try
    End Sub

    Private Shared Function IsFlatViewName(name As String) As Boolean
        If String.IsNullOrWhiteSpace(name) Then Return False
        Return name.IndexOf("flat", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
               name.IndexOf("desple", StringComparison.OrdinalIgnoreCase) >= 0
    End Function

    Private Shared Sub NormalizeAllDimensionTextPolicy(sheet As Sheet, log As DimensionLogger)
        If sheet Is Nothing Then Return
        Dim dims As Dimensions = Nothing
        Try : dims = CType(sheet.Dimensions, Dimensions) : Catch : dims = Nothing : End Try
        If dims Is Nothing Then Return
        Dim n As Integer = 0
        Try : n = dims.Count : Catch : n = 0 : End Try
        For i As Integer = 1 To n
            Dim d As Dimension = Nothing
            Try : d = TryCast(dims.Item(i), Dimension) : Catch : d = Nothing : End Try
            If d Is Nothing Then Continue For
            ApplyDimensionTextPolicy(d, Nothing, log)
        Next
    End Sub

    Private Shared Function FormatInv(v As Double) As String
        Return v.ToString("0.#####", CultureInfo.InvariantCulture)
    End Function

    Private Shared Function ViewNameOf(c As DimensionCandidate) As String
        If c Is Nothing OrElse c.View Is Nothing Then Return "?"
        Try
            Return Convert.ToString(CallByName(c.View, "Name", CallType.Get), CultureInfo.InvariantCulture)
        Catch
            Return "?"
        End Try
    End Function

    Private Shared Sub LogSkippedIsometricViews(sheet As Sheet, log As DimensionLogger)
        If sheet Is Nothing Then Return
        Dim cnt As Integer = 0
        Try
            cnt = sheet.DrawingViews.Count
        Catch
            Return
        End Try
        For i As Integer = 1 To cnt
            Dim dv As DrawingView = Nothing
            Try
                dv = CType(sheet.DrawingViews.Item(i), DrawingView)
            Catch
                dv = Nothing
            End Try
            If dv Is Nothing Then Continue For
            Dim ori As String = SafeStr(CallByNameSafe(dv, "ViewOrientation"))
            Dim dvt As String = SafeStr(CallByNameSafe(dv, "DrawingViewType"))
            Dim dvtNum As Integer = SafeInt(CallByNameSafe(dv, "DrawingViewType"))
            Dim isIso As Boolean =
                ori.IndexOf("iso", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                ori.IndexOf("topfrontright", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                dvt.IndexOf("iso", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                (dvtNum > 0 AndAlso dvtNum <> 1)
            If isIso Then
                log?.LogLine("[DIM][CAND][SKIP_ISO] idx=" & i.ToString(CultureInfo.InvariantCulture) & " name=" & SafeStr(CallByNameSafe(dv, "Name")))
            End If
        Next
    End Sub

    Private Shared Function CallByNameSafe(obj As Object, member As String) As Object
        Try
            Return CallByName(obj, member, CallType.Get)
        Catch
            Return Nothing
        End Try
    End Function

    Private Shared Function SafeStr(v As Object) As String
        If v Is Nothing Then Return ""
        Try
            Return Convert.ToString(v, CultureInfo.InvariantCulture)
        Catch
            Return ""
        End Try
    End Function

    Private Shared Function SafeInt(v As Object) As Integer
        If v Is Nothing Then Return 0
        Try
            Return Convert.ToInt32(v, CultureInfo.InvariantCulture)
        Catch
            Return 0
        End Try
    End Function
End Class
