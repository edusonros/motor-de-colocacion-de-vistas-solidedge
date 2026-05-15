Option Strict Off

Imports System
Imports System.Collections.Generic
Imports System.Globalization
Imports System.Linq
Imports System.Runtime.InteropServices
Imports Microsoft.VisualBasic
Imports SolidEdgeDraft
Imports SolidEdgeFramework
Imports SolidEdgeFrameworkSupport
Imports FrameworkDimension = SolidEdgeFrameworkSupport.Dimension

''' <summary>
''' Motor aislado [PRODDIM]: DVLine2d fuertemente tipado + AddDistanceBetweenObjects sin expandir crop/rango.
''' Primer hito: H_MAX / V_MAX ancladas por nombre (p. ej. "DrawingView 4503"); fallback geométrico.
''' </summary>
Friend NotInheritable Class ProductionDvRefCleanDimensionEngine

    Private Const TolValue As Double = 0.0005R
    Private Const TrackHistoric As Double = 0.024R
    ''' <summary>Peso al elegir TrackDistance: favorece |track| pequeño (línea de cota cerca de las referencias) si el solape con la hoja es similar.</summary>
    Private Const TrackProximityWeight As Double = 6.0R
    Private Const GeomEps As Double = 1.0E-9R
    ''' <summary>Fracción mínima de la altura de vista [RangeMinY, RangeMaxY] que debe cubrir una vertical para considerarse silueta (H_MAX).</summary>
    Private Const HMaxVerticalSpanStrong As Double = 0.36R
    Private Const HMaxVerticalSpanRelaxed As Double = 0.22R
    Private Const PreferredNameFragmentH As String = "4503"
    Private Const PreferredNameFragmentV As String = "4411"
    Private NotInheritable Class ViewAnalysis
        Public ReadOnly Frame As ViewFrame
        Public Name As String
        Public Index As Integer
        Public DrawingViewType As Integer
        Public IsIso As Boolean
        Public IsOrthogonalCandidate As Boolean
        Public LineCount As Integer
        Public BoundsWidth As Double
        Public BoundsHeight As Double
        Public Sub New(vf As ViewFrame)
            Frame = vf
            If vf Is Nothing Then Return
            Name = If(vf.Name, "")
            Index = vf.Index
            DrawingViewType = vf.DrawingViewType
            LineCount = vf.LineCount
            IsIso = (DrawingViewType = 2)
            Dim lw As Double = vf.Width
            Dim lh As Double = vf.Height
            If lw <= GeomEps Then lw = Math.Max(lw, vf.WidthSheet)
            If lh <= GeomEps Then lh = Math.Max(lh, vf.HeightSheet)
            BoundsWidth = lw
            BoundsHeight = lh
            IsOrthogonalCandidate = Not IsIso AndAlso LineCount > 0 AndAlso BoundsWidth > GeomEps AndAlso BoundsHeight > GeomEps
        End Sub
    End Class

    Private NotInheritable Class ViewFrame
        Public View As DrawingView
        Public Name As String
        Public Index As Integer
        ''' <summary>CInt(<see cref="DrawingView.DrawingViewType"/>).</summary>
        Public DrawingViewType As Integer
        ''' <summary>Cotas visibles DVLine2d (solo conteo rápido; sin API de orientación para filtrado).</summary>
        Public LineCount As Integer
        Public ScaleFactor As Double
        ''' <summary>DrawingView.Range en hoja (instantánea inicial).</summary>
        Public InitialSheetMinX As Double
        Public InitialSheetMinY As Double
        Public InitialSheetMaxX As Double
        Public InitialSheetMaxY As Double
        Public CenterSheetX As Double
        Public CenterSheetY As Double
        Public WidthSheet As Double
        Public HeightSheet As Double
        ''' <summary>Agregado de DVLine2d en espacio vista.</summary>
        Public RangeMinX As Double
        Public RangeMinY As Double
        Public RangeMaxX As Double
        Public RangeMaxY As Double
        Public CenterX As Double
        Public CenterY As Double
        Public Width As Double
        Public Height As Double
    End Class

    Private NotInheritable Class DvLineInfo
        Public Index As Integer
        Public Obj As DVLine2d
        Public StartViewX As Double
        Public StartViewY As Double
        Public EndViewX As Double
        Public EndViewY As Double
        Public MinViewX As Double
        Public MaxViewX As Double
        Public MinViewY As Double
        Public MaxViewY As Double
        Public LengthView As Double
        Public IsHorizontal As Boolean
        Public IsVertical As Boolean
        Public IsInclined As Boolean
        Public RefIsNothing As Boolean
        Public ModelMemberIsNothing As Boolean
        Public EdgeType As Integer
        Public KeyPointCount As Integer
        Public MidViewX As Double
        Public MidViewY As Double
    End Class

    Private Sub New()
    End Sub

    Friend Shared Sub Run(app As Application, draft As DraftDocument, log As DimensionLogger)
        If draft Is Nothing Then Return
        P(log, "[START]")
        Dim created As Integer = 0
        Dim kept As Integer = 0
        Dim deleted As Integer = 0
        Dim hOk As Boolean = False
        Dim vOk As Boolean = False
        Dim hVal As Double = Double.NaN
        Dim vVal As Double = Double.NaN

        Dim sh As Sheet = ResolveWorkingSheet(draft, log)
        If sh Is Nothing Then
            P(log, "[SUMMARY] created=0 kept=0 deleted=0 H_MAX=FAIL V_MAX=FAIL finalResult=FAIL reason=no_sheet")
            DimensionProductionRunSummary.RecordProddimRun(0, 0, 0, False, Nothing)
            Return
        End If

        DrawingViewDimensionCreator.TryActivateTargetSheet(draft, sh, log)

        Dim frames As New List(Of ViewFrame)()
        BuildViewFrames(sh, frames, log)
        If frames.Count = 0 Then
            P(log, "[SUMMARY] created=0 kept=0 deleted=0 H_MAX=FAIL V_MAX=FAIL finalResult=FAIL reason=no_views")
            DimensionProductionRunSummary.RecordProddimRun(0, 0, 0, False, Nothing)
            Return
        End If

        Dim viewList As New List(Of ViewAnalysis)()
        For Each vf In frames
            If vf Is Nothing Then Continue For
            viewList.Add(New ViewAnalysis(vf))
        Next
        Dim pinLog As Action(Of String) = Sub(m As String) P(log, m)
        Dim vaH As ViewAnalysis = ResolvePinnedView(viewList, "H_MAX", PreferredNameFragmentH, "main_view", pinLog)
        Dim vaV As ViewAnalysis = ResolvePinnedView(viewList, "V_MAX", PreferredNameFragmentV, "best_vertical_view", pinLog)
        Dim frameH As ViewFrame = If(vaH IsNot Nothing, vaH.Frame, Nothing)
        Dim frameV As ViewFrame = If(vaV IsNot Nothing, vaV.Frame, Nothing)
        If frameH Is Nothing OrElse frameH.View Is Nothing Then
            P(log, "[SUMMARY] created=0 kept=0 deleted=0 H_MAX=FAIL V_MAX=FAIL finalResult=FAIL reason=no_valid_view_H_MAX")
            DimensionProductionRunSummary.RecordProddimRun(0, 0, 0, False, Nothing)
            Return
        End If

        Dim dims As Dimensions = Nothing
        Try
            dims = CType(sh.Dimensions, Dimensions)
        Catch ex As Exception
            LogApiError(log, "Sheet", "Dimensions", ex)
            P(log, "[SUMMARY] created=0 kept=0 deleted=0 H_MAX=FAIL V_MAX=FAIL finalResult=FAIL reason=no_dimensions_collection")
            DimensionProductionRunSummary.RecordProddimRun(0, 0, 0, False, Nothing)
            Return
        End Try

        Dim dimStyleSdk As DimStyle = Nothing
        Dim resolvedStyleNm As String = ""
        Dim sdkStyleResolved As Boolean = TryResolveDimStyleDraftTyped(draft, log, resolvedStyleNm, dimStyleSdk)
        P(log, "[STYLE_RESOLVE] requested=U3,5(or_U2,5_fallback) resolved=" & sdkStyleResolved.ToString(CultureInfo.InvariantCulture) & If(String.IsNullOrEmpty(resolvedStyleNm), "", " name=" & resolvedStyleNm))

        ' --- H_MAX ---
        Dim linesH As List(Of DvLineInfo) = CollectDvLines(frameH.View, frameH.Name, log, logEachLine:=True)
        Dim expH As Double = 0R
        Dim leftL As DvLineInfo = Nothing
        Dim rightL As DvLineInfo = Nothing
        If TrySelectHorizontalExtents(linesH, frameH, leftL, rightL, expH, log) Then
            P(log, String.Format(CultureInfo.InvariantCulture, "[CANDIDATE] tag=H_MAX view={0} leftLine={1} rightLine={2} expected={3:0.######} reason=exterior_vertical_minmax_or_band_fallback",
                                 frameH.Name, leftL.Index.ToString(CultureInfo.InvariantCulture), rightL.Index.ToString(CultureInfo.InvariantCulture), expH))

            Dim yTopL As Double = Math.Max(leftL.StartViewY, leftL.EndViewY)
            Dim yTopR As Double = Math.Max(rightL.StartViewY, rightL.EndViewY)
            Dim yBotL As Double = Math.Min(leftL.StartViewY, leftL.EndViewY)
            Dim yBotR As Double = Math.Min(rightL.StartViewY, rightL.EndViewY)

            Dim vxl As Double, vyl As Double, vxr As Double, vyr As Double
            Dim topLeftIsStart As Boolean, topRightIsStart As Boolean
            TryGetVerticalLineTopByMaxViewY(leftL, vxl, vyl, topLeftIsStart)
            TryGetVerticalLineTopByMaxViewY(rightL, vxr, vyr, topRightIsStart)

            P(log, String.Format(CultureInfo.InvariantCulture,
                "[H_MAX][TOP_RULE] criterion=max(viewY); left_item={0} startView=({1},{2}) endView=({3},{4}) superiorVertex={5} superiorView=({6},{7}) ySpan={8:0.######}",
                leftL.Index.ToString(CultureInfo.InvariantCulture),
                Fmt6(leftL.StartViewX), Fmt6(leftL.StartViewY), Fmt6(leftL.EndViewX), Fmt6(leftL.EndViewY),
                If(topLeftIsStart, "START", "END"),
                Fmt6(vxl), Fmt6(vyl), Fmt6(yTopL - yBotL)))
            P(log, String.Format(CultureInfo.InvariantCulture,
                "[H_MAX][TOP_RULE] criterion=max(viewY); right_item={0} startView=({1},{2}) endView=({3},{4}) superiorVertex={5} superiorView=({6},{7}) ySpan={8:0.######}",
                rightL.Index.ToString(CultureInfo.InvariantCulture),
                Fmt6(rightL.StartViewX), Fmt6(rightL.StartViewY), Fmt6(rightL.EndViewX), Fmt6(rightL.EndViewY),
                If(topRightIsStart, "START", "END"),
                Fmt6(vxr), Fmt6(vyr), Fmt6(yTopR - yBotR)))

            If Math.Abs(yTopL - yTopR) > frameH.Height * 0.02R AndAlso frameH.Height > GeomEps Then
                P(log, String.Format(CultureInfo.InvariantCulture, "[WARN] tag=H_MAX view={0} reason=superior_y_mismatch_between_poles_dy={1:0.######}",
                                     frameH.Name, Math.Abs(yTopL - yTopR)))
            End If

            Dim sx1 As Double, sy1 As Double, sx2 As Double, sy2 As Double
            frameH.View.ViewToSheet(vxl, vyl, sx1, sy1)
            frameH.View.ViewToSheet(vxr, vyr, sx2, sy2)
            P(log, String.Format(CultureInfo.InvariantCulture,
                "[H_MAX][ADDDIST_INPUT] left_item={0} right_item={1} viewTopLeft=({2},{3}) viewTopRight=({4},{5}) sheet1=({6},{7}) sheet2=({8},{9})",
                leftL.Index.ToString(CultureInfo.InvariantCulture), rightL.Index.ToString(CultureInfo.InvariantCulture),
                Fmt6(vxl), Fmt6(vyl), Fmt6(vxr), Fmt6(vyr),
                Fmt6(sx1), Fmt6(sy1), Fmt6(sx2), Fmt6(sy2)))

            Dim dH As FrameworkDimension = Nothing
            P(log, "[CREATE_TRY] tag=H_MAX method=AddDistanceBetweenObjects_DVRefControlled top_endpoints_view_to_sheet")
            If TryAddDistanceControlled(dims, frameH.View, leftL.Obj, sx1, sy1, rightL.Obj, sx2, sy2, dH, log) AndAlso dH IsNot Nothing Then
                created += 1
                TryApplyConstraintFalse(dH)
                TryReattachDimensionToDrawingView(dH, frameH.View, "H_MAX", log)
                TryDraftUpdateAfterDimensionMutation(draft, app)
                Dim valRaw As Double = ReadDimValue(dH)
                P(log, String.Format(CultureInfo.InvariantCulture, "[CREATE_OK] tag=H_MAX value={0:0.######} expected={1:0.######} delta={2:0.######}",
                                     valRaw, expH, Math.Abs(valRaw - expH)))
                PlaceCleanDimension(app, draft, dH, frameH, "H_MAX", "horizontal", log)
                CenterTextByDimensionKeypointSweep(app, draft, dH, frameH, "H_MAX", log)
                ApplyProductionStyleTypedBestEffort(dH, dimStyleSdk, "H_MAX", log)
                If FinalValidateKeepsDimension(draft, frameH, dH, "H_MAX", expH, sh, frameH.View, log) Then
                    P(log, "[KEEP_REASON] tag=H_MAX reason=value_visible_connected_materialized")
                    kept += 1
                    hOk = True
                    hVal = valRaw
                Else
                    SafeDeleteDimension(dH, log)
                    deleted += 1
                End If
            Else
                P(log, "[CREATE_FAIL] tag=H_MAX reason=add_distance_failed")
            End If
        Else
            P(log, "[CANDIDATE] tag=H_MAX reason=no_valid_vertical_pair")
        End If

        ' --- V_MAX (vista fija ObjectID=4411, primer hito) ---
        Dim bottomL As DvLineInfo = Nothing
        Dim topL As DvLineInfo = Nothing
        Dim expV As Double = 0R

        If frameV IsNot Nothing AndAlso frameV.View IsNot Nothing Then
            Dim linesV As List(Of DvLineInfo) = CollectDvLines(frameV.View, frameV.Name, log, logEachLine:=False)
            If TrySelectVerticalExtents(linesV, frameV, bottomL, topL, expV, log) Then
                Dim dx As Double = Math.Abs(bottomL.MidViewX - topL.MidViewX)
                P(log, String.Format(CultureInfo.InvariantCulture, "[VERT_VIEW_SCORE] view={0} score=0 expected={1:0.######} bottomLine={2} topLine={3} reason=horiz_extremes_deltaX={4:0.######} preferred={5}",
                                     frameV.Name, expV, bottomL.Index.ToString(CultureInfo.InvariantCulture), topL.Index.ToString(CultureInfo.InvariantCulture), dx, PreferredNameFragmentV))
            End If
        End If

        If frameV IsNot Nothing AndAlso bottomL IsNot Nothing AndAlso topL IsNot Nothing Then
            P(log, String.Format(CultureInfo.InvariantCulture, "[VERT_VIEW_SELECTED] view={0} score=0 reason=name_pin_{1}", frameV.Name, PreferredNameFragmentV))
            Dim xRef As Double = (bottomL.MidViewX + topL.MidViewX) * 0.5R
            Dim sx1 As Double, sy1 As Double, sx2 As Double, sy2 As Double
            frameV.View.ViewToSheet(xRef, bottomL.MidViewY, sx1, sy1)
            frameV.View.ViewToSheet(xRef, topL.MidViewY, sx2, sy2)
            Dim dV As FrameworkDimension = Nothing
            P(log, "[CREATE_TRY] tag=V_MAX method=AddDistanceBetweenObjects_DVRefControlled")
            If TryAddDistanceControlled(dims, frameV.View, bottomL.Obj, sx1, sy1, topL.Obj, sx2, sy2, dV, log) AndAlso dV IsNot Nothing Then
                created += 1
                TryApplyConstraintFalse(dV)
                TryApplyVerticalMeasurementAxis(dV, log)
                TryReattachDimensionToDrawingView(dV, frameV.View, "V_MAX", log)
                TryDraftUpdateAfterDimensionMutation(draft, app)
                Dim valRaw As Double = ReadDimValue(dV)
                P(log, String.Format(CultureInfo.InvariantCulture, "[CREATE_OK] tag=V_MAX value={0:0.######} expected={1:0.######} delta={2:0.######}",
                                     valRaw, expV, Math.Abs(valRaw - expV)))
                PlaceCleanDimension(app, draft, dV, frameV, "V_MAX", "vertical", log)
                CenterTextByDimensionKeypointSweep(app, draft, dV, frameV, "V_MAX", log)
                ApplyProductionStyleTypedBestEffort(dV, dimStyleSdk, "V_MAX", log)
                If FinalValidateKeepsDimension(draft, frameV, dV, "V_MAX", expV, sh, frameV.View, log) Then
                    P(log, "[KEEP_REASON] tag=V_MAX reason=value_visible_connected_materialized")
                    kept += 1
                    vOk = True
                    vVal = valRaw
                Else
                    SafeDeleteDimension(dV, log)
                    deleted += 1
                End If
            Else
                P(log, "[CREATE_FAIL] tag=V_MAX reason=add_distance_failed")
            End If
        Else
            P(log, "[VERT_VIEW_SELECTED] view=(none) score=-inf reason=no_candidate")
        End If

        Try
            If draft IsNot Nothing Then draft.UpdateAll(True)
        Catch
        End Try
        Try
            If app IsNot Nothing Then app.DoIdle()
        Catch
        End Try

        Dim fr As String = If(hOk AndAlso vOk, "SUCCESS", "FAIL")
        P(log, String.Format(CultureInfo.InvariantCulture,
            "[SUMMARY] created={0} kept={1} deleted={2} H_MAX={3} H_MAX_VALUE={4:0.######} V_MAX={5} V_MAX_VALUE={6:0.######} finalResult={7}",
            created, kept, deleted,
            If(hOk, "SUCCESS", "FAIL"), If(hOk, hVal, Double.NaN),
            If(vOk, "SUCCESS", "FAIL"), If(vOk, vVal, Double.NaN), fr))
        Dim styleForSummary As String = If(sdkStyleResolved AndAlso Not String.IsNullOrWhiteSpace(resolvedStyleNm), resolvedStyleNm, Nothing)
        DimensionProductionRunSummary.RecordProddimRun(created, kept, deleted, hOk AndAlso vOk, styleForSummary)
    End Sub

    Private Shared Sub P(log As DimensionLogger, msg As String)
        log?.LogLine("[PRODDIM]" & msg)
    End Sub

    Private Shared Function Fmt6(x As Double) As String
        Return x.ToString("0.######", CultureInfo.InvariantCulture)
    End Function

    Private Shared Sub LogApiError(log As DimensionLogger, obj As String, method As String, ex As Exception)
        If log Is Nothing OrElse ex Is Nothing Then Return
        Dim hr As String = ""
        Dim cex As COMException = TryCast(ex, COMException)
        If cex IsNot Nothing Then hr = "0x" & cex.ErrorCode.ToString("X8", CultureInfo.InvariantCulture)
        log.LogLine("[PRODDIM][API_ERROR] object=" & obj & " method=" & method & " hresult=" & hr & " message=" & ex.Message.Replace(ChrW(10), " ").Replace(ChrW(13), " ") & " action=REQUEST_SDK_DOC")
    End Sub

    Private Shared Function ResolveWorkingSheet(draft As DraftDocument, log As DimensionLogger) As Sheet
        Dim preferred As Sheet = Nothing
        Dim fallback As Sheet = Nothing
        Dim visit = Sub(sh As Sheet, shi As Integer)
                          If sh Is Nothing Then Return
                          Dim nm As String = SafeSheetName(sh)
                          Dim nDv As Integer = SafeDrawingViewsCount(sh)
                          P(log, String.Format(CultureInfo.InvariantCulture, "[SHEET_SCAN] idx={0} name={1} views={2}", shi, nm, nDv.ToString()))
                          If IsTwoDModelSheet(nm) Then Return
                          If nDv <= 0 Then Return
                          If fallback Is Nothing Then fallback = sh
                          If String.Equals(nm, "Hoja1", StringComparison.OrdinalIgnoreCase) AndAlso preferred Is Nothing Then
                              preferred = sh
                          End If
                      End Sub

        Dim anySections As Boolean = False
        Try
            Dim sections = draft.Sections
            Dim nSec As Integer = CInt(sections.Count)
            For s As Integer = 1 To nSec
                anySections = True
                Dim sec As Section = Nothing
                Try : sec = CType(sections.Item(s), Section) : Catch : sec = Nothing : End Try
                If sec Is Nothing Then Continue For
                Dim sheetsCol = sec.Sheets
                Dim nSh As Integer = CInt(sheetsCol.Count)
                For i As Integer = 1 To nSh
                    Dim sh As Sheet = Nothing
                    Try : sh = CType(sheetsCol.Item(i), Sheet) : Catch : sh = Nothing : End Try
                    visit(sh, i)
                Next
            Next
        Catch ex As Exception
            P(log, "[SHEET_SCAN][ERR] " & ex.Message)
        End Try

        If Not anySections OrElse (preferred Is Nothing AndAlso fallback Is Nothing) Then
            Try
                Dim n As Integer = CInt(draft.Sheets.Count)
                For i As Integer = 1 To n
                    Dim sh As Sheet = Nothing
                    Try : sh = CType(draft.Sheets.Item(i), Sheet) : Catch : sh = Nothing : End Try
                    visit(sh, i)
                Next
            Catch ex2 As Exception
                P(log, "[SHEET_SCAN][FALLBACK_SHEETS_ERR] " & ex2.Message)
            End Try
        End If

        Dim chosen As Sheet = If(preferred IsNot Nothing, preferred, fallback)
        If chosen IsNot Nothing Then
            P(log, String.Format(CultureInfo.InvariantCulture, "[SHEET_SELECTED] name={0} views={1}",
                                 SafeSheetName(chosen), SafeDrawingViewsCount(chosen).ToString()))
        End If
        Return chosen
    End Function

    Private Shared Function SafeSheetName(sh As Sheet) As String
        If sh Is Nothing Then Return ""
        Try
            Return Convert.ToString(sh.Name, CultureInfo.InvariantCulture).Trim()
        Catch
            Return ""
        End Try
    End Function

    Private Shared Function SafeDrawingViewsCount(sh As Sheet) As Integer
        If sh Is Nothing Then Return 0
        Try
            Return CInt(sh.DrawingViews.Count)
        Catch
            Return 0
        End Try
    End Function

    Private Shared Function IsTwoDModelSheet(name As String) As Boolean
        Return String.Equals(If(name, "").Trim(), "2D Model", StringComparison.OrdinalIgnoreCase)
    End Function

    Private Shared Sub BuildViewFrames(sh As Sheet, frames As List(Of ViewFrame), log As DimensionLogger)
        Dim n As Integer = 0
        Try
            n = CInt(sh.DrawingViews.Count)
        Catch
            n = 0
        End Try
        For i As Integer = 1 To n
            Dim dv As DrawingView = Nothing
            Try
                dv = CType(sh.DrawingViews.Item(i), DrawingView)
            Catch
                dv = Nothing
            End Try
            If dv Is Nothing Then Continue For
            Dim name As String = ""
            Dim typ As Integer = -1
            Dim sc As Double = 1R
            Try : name = Convert.ToString(dv.Name, CultureInfo.InvariantCulture) : Catch : name = "" : End Try
            Try : typ = CInt(dv.DrawingViewType) : Catch : typ = -1 : End Try
            Try : sc = CDbl(dv.ScaleFactor) : Catch : sc = 1R : End Try

            Dim rx1 As Double = 0, ry1 As Double = 0, rx2 As Double = 0, ry2 As Double = 0
            Try
                dv.Range(rx1, ry1, rx2, ry2)
            Catch
                rx1 = 0 : ry1 = 0 : rx2 = 0 : ry2 = 0
            End Try

            Dim nLines As Integer = 0
            Dim nArcs As Integer = 0
            Dim nPts As Integer = 0
            Try : nLines = CInt(dv.DVLines2d.Count) : Catch : nLines = 0 : End Try
            Try : nArcs = CInt(dv.DVArcs2d.Count) : Catch : nArcs = 0 : End Try
            Try : nPts = CInt(dv.DVPoints2d.Count) : Catch : nPts = 0 : End Try

            P(log, String.Format(CultureInfo.InvariantCulture,
                "[VIEW_SCAN] idx={0} name={1} type={2} scale={3:0.######} range=({4},{5})-({6},{7}) lines={8} arcs={9} points={10}",
                i, name, typ.ToString(CultureInfo.InvariantCulture), sc, Fmt6(rx1), Fmt6(ry1), Fmt6(rx2), Fmt6(ry2),
                nLines.ToString(CultureInfo.InvariantCulture), nArcs.ToString(CultureInfo.InvariantCulture), nPts.ToString(CultureInfo.InvariantCulture)))

            Dim lines As List(Of DvLineInfo) = CollectDvLines(dv, name, log, logEachLine:=True)
            Dim gminX As Double = Double.PositiveInfinity
            Dim gmaxX As Double = Double.NegativeInfinity
            Dim gminY As Double = Double.PositiveInfinity
            Dim gmaxY As Double = Double.NegativeInfinity
            For Each ln In lines
                gminX = Math.Min(gminX, ln.MinViewX)
                gmaxX = Math.Max(gmaxX, ln.MaxViewX)
                gminY = Math.Min(gminY, ln.MinViewY)
                gmaxY = Math.Max(gmaxY, ln.MaxViewY)
            Next
            If Double.IsInfinity(gminX) Then
                gminX = 0R
                gmaxX = 0R
                gminY = 0R
                gmaxY = 0R
            End If
            Dim gw As Double = Math.Max(0R, gmaxX - gminX)
            Dim gh As Double = Math.Max(0R, gmaxY - gminY)
            P(log, String.Format(CultureInfo.InvariantCulture, "[BOUNDS] view={0} minX={1} maxX={2} minY={3} maxY={4} width={5} height={6}",
                                 name, Fmt6(gminX), Fmt6(gmaxX), Fmt6(gminY), Fmt6(gmaxY), Fmt6(gw), Fmt6(gh)))

            Dim smx1 As Double = Math.Min(rx1, rx2)
            Dim smy1 As Double = Math.Min(ry1, ry2)
            Dim smx2 As Double = Math.Max(rx1, rx2)
            Dim smy2 As Double = Math.Max(ry1, ry2)
            Dim vf As New ViewFrame With {
                .View = dv,
                .Name = name,
                .Index = i,
                .DrawingViewType = typ,
                .LineCount = nLines,
                .ScaleFactor = sc,
                .InitialSheetMinX = smx1,
                .InitialSheetMinY = smy1,
                .InitialSheetMaxX = smx2,
                .InitialSheetMaxY = smy2,
                .WidthSheet = Math.Max(0R, smx2 - smx1),
                .HeightSheet = Math.Max(0R, smy2 - smy1),
                .CenterSheetX = (smx1 + smx2) * 0.5R,
                .CenterSheetY = (smy1 + smy2) * 0.5R,
                .RangeMinX = gminX,
                .RangeMaxX = gmaxX,
                .RangeMinY = gminY,
                .RangeMaxY = gmaxY,
                .Width = gw,
                .Height = gh,
                .CenterX = (gminX + gmaxX) * 0.5R,
                .CenterY = (gminY + gmaxY) * 0.5R
            }
            frames.Add(vf)
        Next
    End Sub

    Private Shared Function CollectDvLines(dv As DrawingView, viewName As String, log As DimensionLogger, logEachLine As Boolean) As List(Of DvLineInfo)
        Dim list As New List(Of DvLineInfo)()
        If dv Is Nothing Then Return list
        Dim n As Integer = 0
        Try
            n = CInt(dv.DVLines2d.Count)
        Catch
            Return list
        End Try
        For i As Integer = 1 To n
            Dim ln As DVLine2d = Nothing
            Try
                ln = CType(dv.DVLines2d.Item(i), DVLine2d)
            Catch
                ln = Nothing
            End Try
            If ln Is Nothing Then Continue For
            Dim x1 As Double = 0, y1 As Double = 0, x2 As Double = 0, y2 As Double = 0
            Dim rminX As Double = 0, rminY As Double = 0, rmaxX As Double = 0, rmaxY As Double = 0
            Try
                ln.GetStartPoint(x1, y1)
                ln.GetEndPoint(x2, y2)
                ln.Range(rminX, rminY, rmaxX, rmaxY)
            Catch ex As Exception
                P(log, "[DVLINE][READ_ERR] view=" & viewName & " idx=" & i.ToString(CultureInfo.InvariantCulture) & " " & ex.Message)
                Continue For
            End Try
            Dim dx As Double = Math.Abs(x2 - x1)
            Dim dy As Double = Math.Abs(y2 - y1)
            Dim len As Double = Math.Sqrt(dx * dx + dy * dy)
            Dim tol As Double = Math.Max(len * 1.0E-8R, 1.0E-10R)
            Dim isH As Boolean = (dy <= tol AndAlso dx > tol)
            Dim isV As Boolean = (dx <= tol AndAlso dy > tol)
            Dim isI As Boolean = Not isH AndAlso Not isV
            Dim refN As Boolean = True
            Dim mmN As Boolean = True
            Try
                refN = (ln.Reference Is Nothing)
            Catch
                refN = True
            End Try
            Try
                mmN = (ln.ModelMember Is Nothing)
            Catch
                mmN = True
            End Try
            Dim edgeT As Integer = -1
            Try : edgeT = CInt(ln.EdgeType) : Catch : edgeT = -1 : End Try
            Dim kpc As Integer = 0
            Try : kpc = CInt(ln.KeyPointCount) : Catch : kpc = 0 : End Try
            Dim midX As Double = (x1 + x2) * 0.5R
            Dim midY As Double = (y1 + y2) * 0.5R
            Dim msx As Double = 0, msy As Double = 0
            Try
                dv.ViewToSheet(midX, midY, msx, msy)
            Catch
                msx = 0
                msy = 0
            End Try
            If logEachLine Then
                P(log, String.Format(CultureInfo.InvariantCulture,
                    "[DVLINE] view={0} idx={1} startView=({2},{3}) endView=({4},{5}) rangeView=({6},{7})-({8},{9}) length={10} h={11} v={12} midSheet=({13},{14}) refNothing={15} modelMemberNothing={16} edgeType={17}",
                    viewName, i.ToString(CultureInfo.InvariantCulture), Fmt6(x1), Fmt6(y1), Fmt6(x2), Fmt6(y2),
                    Fmt6(rminX), Fmt6(rminY), Fmt6(rmaxX), Fmt6(rmaxY), Fmt6(len),
                    isH.ToString(CultureInfo.InvariantCulture), isV.ToString(CultureInfo.InvariantCulture),
                    Fmt6(msx), Fmt6(msy), refN.ToString(CultureInfo.InvariantCulture), mmN.ToString(CultureInfo.InvariantCulture), edgeT.ToString(CultureInfo.InvariantCulture)))
            End If
            list.Add(New DvLineInfo With {
                .Index = i,
                .Obj = ln,
                .StartViewX = x1,
                .StartViewY = y1,
                .EndViewX = x2,
                .EndViewY = y2,
                .MinViewX = rminX,
                .MaxViewX = rmaxX,
                .MinViewY = rminY,
                .MaxViewY = rmaxY,
                .LengthView = len,
                .IsHorizontal = isH,
                .IsVertical = isV,
                .IsInclined = isI,
                .RefIsNothing = refN,
                .ModelMemberIsNothing = mmN,
                .EdgeType = edgeT,
                .KeyPointCount = kpc,
                .MidViewX = midX,
                .MidViewY = midY
            })
        Next
        Return list
    End Function

    ''' <summary>
    ''' Para una DVLine2d casi-vertical, selecciona en espacio vista el extremo de <see cref="DvLineInfo.StartViewY"/> / <see cref="DvLineInfo.EndViewY"/>
    ''' con Y mayor como candidato “superior”. Loguear ambos valores para confirmar que coincide con el borde visual de la chapa.
    ''' </summary>
    Private Shared Sub TryGetVerticalLineTopByMaxViewY(ln As DvLineInfo,
                                                     ByRef topViewX As Double,
                                                     ByRef topViewY As Double,
                                                     ByRef topFromStartVertex As Boolean)
        topViewX = 0R
        topViewY = 0R
        topFromStartVertex = True
        If ln Is Nothing Then Return
        If ln.StartViewY >= ln.EndViewY Then
            topViewX = ln.StartViewX
            topViewY = ln.StartViewY
            topFromStartVertex = True
        Else
            topViewX = ln.EndViewX
            topViewY = ln.EndViewY
            topFromStartVertex = False
        End If
    End Sub

    Private Shared Function ViewRejectReason(va As ViewAnalysis) As String
        If va Is Nothing Then Return "null_entry"
        If va.IsIso Then Return "drawingViewType=2_isometric"
        If va.LineCount <= 0 Then Return "zero_lines"
        If va.BoundsWidth <= GeomEps Then Return "zero_bounds_width"
        If va.BoundsHeight <= GeomEps Then Return "zero_bounds_height"
        Return ""
    End Function

    ''' <summary>Primer hito: excluir vistas cuyo nombre sugiere desplegado Flat (no tocar geometría Flat en profundidad todavía).</summary>
    Private Shared Function ViewNameIndicatesFlat(nm As String) As Boolean
        Return Not String.IsNullOrEmpty(nm) AndAlso nm.IndexOf("Flat", StringComparison.OrdinalIgnoreCase) >= 0
    End Function

    Private Shared Function ResolvePinnedView(views As List(Of ViewAnalysis), tag As String, preferredNameFragment As String,
                                              fallbackMode As String, log As Action(Of String)) As ViewAnalysis
        If views Is Nothing OrElse views.Count = 0 Then
            If log IsNot Nothing Then log("[VIEW_PIN][FAIL] tag=" & tag & " reason=no_sheet_views_analyzed")
            Return Nothing
        End If

        Dim validViews As List(Of ViewAnalysis) =
            views.Where(Function(v) v IsNot Nothing AndAlso v.IsOrthogonalCandidate).ToList()
        Dim usedRelaxedGate As Boolean = False

        If log IsNot Nothing Then
            log(String.Format(CultureInfo.InvariantCulture, "[VIEW_LIST][COUNT] all={0} valid={1}",
                views.Count.ToString(CultureInfo.InvariantCulture), validViews.Count.ToString(CultureInfo.InvariantCulture)))
            For Each va In views
                If va Is Nothing Then Continue For
                Dim rr As String = ViewRejectReason(va)
                If rr = "" Then rr = "(orthogonal_candidate_ok)"
                log(String.Format(CultureInfo.InvariantCulture,
                    "[VIEW_LIST][ITEM] idx={0} name={1} type={2} isIso={3} isOrthogonalCandidate={4} lines={5} width={6} height={7} rejectReason={8}",
                    va.Index.ToString(CultureInfo.InvariantCulture), va.Name, va.DrawingViewType.ToString(CultureInfo.InvariantCulture),
                    va.IsIso.ToString(CultureInfo.InvariantCulture), va.IsOrthogonalCandidate.ToString(CultureInfo.InvariantCulture),
                    va.LineCount.ToString(CultureInfo.InvariantCulture), Fmt6(va.BoundsWidth), Fmt6(va.BoundsHeight), rr))
            Next
        End If

        If validViews.Count = 0 AndAlso log IsNot Nothing Then
            For Each va In views
                If va Is Nothing Then Continue For
                Dim rj As String = ViewRejectReason(va)
                log(String.Format(CultureInfo.InvariantCulture,
                    "[VIEW_REJECT] name={0} type={1} lines={2} width={3} height={4} reason={5}",
                    va.Name, va.DrawingViewType.ToString(CultureInfo.InvariantCulture),
                    va.LineCount.ToString(CultureInfo.InvariantCulture), Fmt6(va.BoundsWidth), Fmt6(va.BoundsHeight), rj))
            Next
        End If

        If validViews.Count = 0 Then
            Dim relaxed As List(Of ViewAnalysis) =
                views.Where(Function(v) v IsNot Nothing AndAlso Not v.IsIso AndAlso v.LineCount > 0).ToList()
            If relaxed.Count > 0 Then
                If log IsNot Nothing Then log("[VIEW_LIST][WARN] reason=relaxing_gate_to_non_iso_with_lines_keeps_bounds_unchecked")
                validViews = relaxed
                usedRelaxedGate = True
            Else
                Dim loose As List(Of ViewAnalysis) =
                    views.Where(Function(v) v IsNot Nothing AndAlso Not v.IsIso).ToList()
                If loose.Count > 0 Then
                    If log IsNot Nothing Then log("[VIEW_LIST][WARN] reason=relaxing_gate_to_any_non_iso")
                    validViews = loose
                    usedRelaxedGate = True
                End If
            End If
        End If

        If validViews.Count = 0 Then
            If log IsNot Nothing Then log("[VIEW_PIN][FAIL] tag=" & tag & " reason=no_valid_orthogonal_views")
            Return Nothing
        End If

        If log IsNot Nothing Then
            log("[VIEW_PIN][TRY] tag=" & tag & " preferred=" & If(preferredNameFragment, ""))
        End If

        Dim fragTrim As String = If(preferredNameFragment, "").Trim()
        Dim exactTarget As String = "DrawingView " & fragTrim

        For Each va In validViews
            If va Is Nothing OrElse va.Frame Is Nothing Then Continue For
            Dim nm As String = va.Name
            Dim matchExact As Boolean = (Not String.IsNullOrEmpty(fragTrim)) AndAlso String.Equals(nm, exactTarget, StringComparison.OrdinalIgnoreCase)
            Dim matchContains As Boolean = (Not String.IsNullOrEmpty(fragTrim)) AndAlso nm.IndexOf(fragTrim, StringComparison.OrdinalIgnoreCase) >= 0
            Dim matchEnds As Boolean = (Not String.IsNullOrEmpty(fragTrim)) AndAlso nm.Trim().EndsWith(fragTrim, StringComparison.OrdinalIgnoreCase)
            If log IsNot Nothing Then
                Dim orthoShown As Boolean = va.IsOrthogonalCandidate OrElse usedRelaxedGate
                log(String.Format(CultureInfo.InvariantCulture,
                    "[VIEW_PIN][CAND] tag={0} idx={1} name={2} matchExact={3} matchContains={4} matchEndsWith={5} isOrthogonalCandidate={6}",
                    tag, va.Index.ToString(CultureInfo.InvariantCulture), nm,
                    matchExact.ToString(CultureInfo.InvariantCulture), matchContains.ToString(CultureInfo.InvariantCulture), matchEnds.ToString(CultureInfo.InvariantCulture),
                    orthoShown.ToString(CultureInfo.InvariantCulture)))
            End If
        Next

        If Not String.IsNullOrEmpty(fragTrim) Then
            For Each va In validViews
                If va Is Nothing Then Continue For
                Dim nm As String = va.Name
                If String.Equals(nm, exactTarget, StringComparison.OrdinalIgnoreCase) Then
                    If log IsNot Nothing Then log("[VIEW_PIN][OK] tag=" & tag & " name=" & nm & " reason=exact")
                    Return va
                End If
            Next
            For Each va In validViews
                If va Is Nothing Then Continue For
                Dim nm As String = va.Name
                If nm.IndexOf(fragTrim, StringComparison.OrdinalIgnoreCase) >= 0 Then
                    If log IsNot Nothing Then log("[VIEW_PIN][OK] tag=" & tag & " name=" & nm & " reason=contains")
                    Return va
                End If
            Next
            For Each va In validViews
                If va Is Nothing Then Continue For
                Dim nm As String = va.Name
                If nm.Trim().EndsWith(fragTrim, StringComparison.OrdinalIgnoreCase) Then
                    If log IsNot Nothing Then log("[VIEW_PIN][OK] tag=" & tag & " name=" & nm & " reason=endswith")
                    Return va
                End If
            Next
        End If

        Dim picked As ViewAnalysis = Nothing
        If String.Equals(fallbackMode, "best_vertical_view", StringComparison.OrdinalIgnoreCase) Then
            picked = PickVMaxFallback(validViews)
        Else
            picked = PickHMaxAreaExcludingFlat(validViews)
        End If

        If picked Is Nothing OrElse picked.Frame Is Nothing Then
            If log IsNot Nothing Then log("[VIEW_PIN][FAIL] tag=" & tag & " reason=no_fallback_selection")
            Return Nothing
        End If
        If log IsNot Nothing Then
            log("[VIEW_PIN][FALLBACK] tag=" & tag & " mode=" & fallbackMode & " selected=" & picked.Frame.Name)
            log("[VIEW_PIN][OK] tag=" & tag & " name=" & picked.Frame.Name & " reason=fallback")
        End If
        Return picked
    End Function

    ''' <summary>Válida ortogonal mayor área de proyección; excluye nombres "Flat" en primer hito.</summary>
    Private Shared Function PickHMaxAreaExcludingFlat(validViews As List(Of ViewAnalysis)) As ViewAnalysis
        Dim nonFlat As List(Of ViewAnalysis) =
            validViews.Where(Function(v) v IsNot Nothing AndAlso Not ViewNameIndicatesFlat(v.Name)).ToList()
        Dim scan As List(Of ViewAnalysis) = If(nonFlat.Count > 0, nonFlat, validViews)
        Dim best As ViewAnalysis = Nothing
        Dim bestArea As Double = Double.NegativeInfinity
        For Each va In scan
            If va Is Nothing OrElse va.Frame Is Nothing Then Continue For
            Dim a As Double = va.BoundsWidth * va.BoundsHeight
            If best Is Nothing Then
                best = va
                bestArea = a
            ElseIf a > bestArea + GeomEps Then
                best = va
                bestArea = a
            ElseIf Math.Abs(a - bestArea) <= GeomEps AndAlso va.Index < best.Index Then
                best = va
            End If
        Next
        Return best
    End Function

    Private Shared Function PickVMaxFallback(validViews As List(Of ViewAnalysis)) As ViewAnalysis
        Const targetExp As Double = 0.09R
        Dim bestVa As ViewAnalysis = Nothing
        Dim bestScore As Double = Double.NegativeInfinity
        For Each va In validViews
            If va Is Nothing OrElse va.Frame Is Nothing OrElse va.Frame.View Is Nothing Then Continue For
            Dim lines As List(Of DvLineInfo) = CollectDvLines(va.Frame.View, va.Frame.Name, Nothing, logEachLine:=False)
            Dim b As DvLineInfo = Nothing
            Dim t As DvLineInfo = Nothing
            Dim expV As Double = 0R
            Dim okExtents As Boolean = TrySelectVerticalExtents(lines, va.Frame, b, t, expV, Nothing)
            Dim nm As String = va.Name
            Dim bon4411 As Double = 0R
            If Not String.IsNullOrEmpty(nm) AndAlso nm.IndexOf("4411", StringComparison.OrdinalIgnoreCase) >= 0 Then bon4411 = 55R
            Dim heightNear As Double = -Math.Abs(va.BoundsHeight - targetExp) * 180R
            Dim score As Double
            If okExtents Then
                Dim dx As Double = Math.Abs(b.MidViewX - t.MidViewX)
                score = bon4411 + 25R - Math.Abs(expV - targetExp) * 240R - dx * 5R + heightNear
            Else
                score = bon4411 + heightNear - 120R
            End If
            If score > bestScore Then
                bestScore = score
                bestVa = va
            End If
        Next
        If bestVa IsNot Nothing Then Return bestVa
        For Each va In validViews
            If va IsNot Nothing AndAlso va.Frame IsNot Nothing Then Return va
        Next
        Return Nothing
    End Function

    ''' <summary>
    ''' SDK (<c>SolidEdgeFrameworkSupport~Dimension~Style.html</c>): <c>Dimension.Style As DimStyle</c>, read-write.
    ''' Sin <c>StyleName</c> en miembros públicos de Dimension. Se intenta sólo una asignación fuertemente tipada.
    ''' </summary>
    Private Shared Function TryResolveDimStyleDraftTyped(draft As DraftDocument, log As DimensionLogger, ByRef resolvedNameOut As String, ByRef styleOut As DimStyle) As Boolean
        resolvedNameOut = ""
        styleOut = Nothing
        If draft Is Nothing Then Return False
        Dim styles As DimensionStyles = Nothing
        Try
            styles = CType(draft.DimensionStyles, DimensionStyles)
        Catch ex As Exception
            LogApiError(log, "DraftDocument", "DimensionStyles", ex)
            Return False
        End Try
        If styles Is Nothing Then Return False
        Dim candNames As String() = {"U3,5", "U2,5"}
        For Each cand In candNames
            Dim named As DimensionStyle = Nothing
            Try
                named = CType(styles.Item(cand), DimensionStyle)
            Catch ex As Exception
                LogApiError(log, "DimensionStyles", "Item(" & cand & ")", ex)
                Continue For
            End Try
            If named Is Nothing Then
                P(log, "[STYLE_RESOLVE][WARN] requested=" & cand & " reason=Item_returned_nothing")
                Continue For
            End If
            Dim dimSt As DimStyle = TryCast(named, DimStyle)
            If dimSt Is Nothing Then
                Try
                    dimSt = CType(named, DimStyle)
                Catch ex As Exception
                    LogApiError(log, "DimensionStyle", "CType(DimensionStyle,DimStyle)", ex)
                    P(log, "[STYLE_RESOLVE][WARN] requested=" & cand & " reason=DimensionStyle_not_castable_to_DimStyle")
                    Continue For
                End Try
            End If
            If dimSt IsNot Nothing Then
                styleOut = dimSt
                resolvedNameOut = cand
                Return True
            End If
        Next
        Return False
    End Function

    Private Shared Sub ApplyProductionStyleTypedBestEffort(d As FrameworkDimension, styleTyped As DimStyle, tag As String, log As DimensionLogger)
        If d Is Nothing Then Return
        If styleTyped Is Nothing Then
            P(log, "[STYLE_APPLY][WARN] tag=" & tag & " reason=no_DimStyle_resolved skipping_assign")
            Return
        End If
        Try
            d.Style = styleTyped
            Dim nm As String = ""
            Try
                nm = Convert.ToString(styleTyped.Name, CultureInfo.InvariantCulture)
            Catch
                nm = ""
            End Try
            P(log, "[STYLE_APPLY] tag=" & tag & " route=Dimension.Style=DimStyle(typed) result=OK final=" & nm.Trim())
        Catch ex As Exception
            LogApiError(log, "Dimension", "Style (DimStyle)", ex)
            P(log, "[STYLE_APPLY][WARN] tag=" & tag & " reason=assignment_failed conserving_dimension")
        End Try
    End Sub

    ''' <summary>Parte de la altura de la envolvente de líneas en vista que cubre esta vertical (0–1).</summary>
    Private Shared Function VerticalSpanCoverage(ln As DvLineInfo, vf As ViewFrame) As Double
        If ln Is Nothing OrElse vf Is Nothing Then Return 0R
        Dim rh As Double = vf.RangeMaxY - vf.RangeMinY
        If rh <= GeomEps Then Return 0R
        Dim olMin As Double = Math.Max(ln.MinViewY, vf.RangeMinY)
        Dim olMax As Double = Math.Min(ln.MaxViewY, vf.RangeMaxY)
        Dim overlap As Double = Math.Max(0R, olMax - olMin)
        Return overlap / rh
    End Function

    ''' <summary>Pareja exterior: vertical con <see cref="DvLineInfo.MinViewX"/> mínimo y vertical con <see cref="DvLineInfo.MaxViewX"/> máximo (desempate: longitud y arista con referencia de modelo).</summary>
    Private Shared Function TryPickExteriorVerticalExtents(verts As List(Of DvLineInfo), vf As ViewFrame,
                                                          ByRef pickL As DvLineInfo, ByRef pickR As DvLineInfo) As Boolean
        pickL = Nothing
        pickR = Nothing
        If verts Is Nothing OrElse verts.Count < 2 OrElse vf Is Nothing Then Return False

        Dim wallBand As Double = Math.Max(vf.Width * 3.0E-5R, 1.0E-6R)
        Dim minX As Double = Double.PositiveInfinity
        For Each c In verts
            If c.MinViewX < minX Then minX = c.MinViewX
        Next
        Dim maxX As Double = Double.NegativeInfinity
        For Each c In verts
            If c.MaxViewX > maxX Then maxX = c.MaxViewX
        Next

        Dim bestL As DvLineInfo = Nothing
        Dim bestLScore As Double = Double.NegativeInfinity
        For Each c In verts
            If c.MinViewX > minX + wallBand Then Continue For
            Dim sc As Double = c.LengthView
            If Not c.RefIsNothing Then sc *= 1.06R
            If sc > bestLScore + GeomEps Then
                bestLScore = sc
                bestL = c
            End If
        Next

        Dim bestR As DvLineInfo = Nothing
        Dim bestRScore As Double = Double.NegativeInfinity
        For Each c In verts
            If c.MaxViewX < maxX - wallBand Then Continue For
            Dim sc As Double = c.LengthView
            If Not c.RefIsNothing Then sc *= 1.06R
            If sc > bestRScore + GeomEps Then
                bestRScore = sc
                bestR = c
            End If
        Next

        If bestL Is Nothing OrElse bestR Is Nothing Then Return False
        If ReferenceEquals(bestL, bestR) Then Return False
        Dim sep As Double = Math.Abs(bestR.MidViewX - bestL.MidViewX)
        If sep <= wallBand * 4R Then Return False

        pickL = bestL
        pickR = bestR
        Return True
    End Function

    Private Shared Function TrySelectHorizontalExtents(lines As List(Of DvLineInfo), vf As ViewFrame,
                                                      ByRef leftOut As DvLineInfo, ByRef rightOut As DvLineInfo,
                                                      ByRef expected As Double, log As DimensionLogger) As Boolean
        leftOut = Nothing
        rightOut = Nothing
        expected = 0R
        If lines Is Nothing OrElse vf Is Nothing Then Return False
        Dim verts As New List(Of DvLineInfo)()
        Dim minLen As Double = Math.Max(vf.Height * 0.35R, 0.015R)
        For Each ln In lines
            If ln Is Nothing OrElse Not ln.IsVertical Then Continue For
            If ln.LengthView < minLen Then Continue For
            verts.Add(ln)
        Next
        If verts.Count < 2 Then Return False

        Dim spanLevels As Double() = {HMaxVerticalSpanStrong, HMaxVerticalSpanRelaxed, 0R}
        For si As Integer = 0 To spanLevels.Length - 1
            Dim spanMin As Double = spanLevels(si)
            Dim useList As List(Of DvLineInfo)
            If spanMin > GeomEps Then
                useList = New List(Of DvLineInfo)()
                For Each ln In verts
                    If VerticalSpanCoverage(ln, vf) >= spanMin Then useList.Add(ln)
                Next
            Else
                useList = verts
            End If
            If useList.Count < 2 Then Continue For

            Dim pickL As DvLineInfo = Nothing
            Dim pickR As DvLineInfo = Nothing
            If Not TryPickExteriorVerticalExtents(useList, vf, pickL, pickR) Then Continue For

            Dim expTry As Double = Math.Abs(pickR.MidViewX - pickL.MidViewX)
            If expTry < 0.02R OrElse expTry > 3.0R Then Continue For

            leftOut = pickL
            rightOut = pickR
            expected = expTry
            If log IsNot Nothing Then
                Dim spCovL As Double = VerticalSpanCoverage(pickL, vf)
                Dim spCovR As Double = VerticalSpanCoverage(pickR, vf)
                P(log, String.Format(CultureInfo.InvariantCulture,
                    "[EXTENTS_H] route=exterior_minmax_x spanMin={0:0.###} covL={1:0.###} covR={2:0.###}",
                    spanMin, spCovL, spCovR))
            End If
            Return True
        Next

        ' --- Retroceso: extremos por MidViewX + bandas al envolvente (comportamiento anterior) ---
        verts.Sort(Function(a, b) a.MidViewX.CompareTo(b.MidViewX))
        Dim fallbackL As DvLineInfo = verts(0)
        Dim fallbackR As DvLineInfo = verts(verts.Count - 1)

        Dim nearBand As Double = Math.Max(vf.Width * 0.12R, 0.008R)
        Dim altL As DvLineInfo = Nothing
        Dim bestScore As Double = Double.NegativeInfinity
        For Each c In verts
            If c.MidViewX <= vf.RangeMinX + nearBand Then
                Dim score As Double = c.LengthView
                If Not c.RefIsNothing Then score *= 1.05R
                If score > bestScore Then bestScore = score : altL = c
            End If
        Next
        Dim altR As DvLineInfo = Nothing
        bestScore = Double.NegativeInfinity
        For Each c In verts
            If c.MidViewX >= vf.RangeMaxX - nearBand Then
                Dim score As Double = c.LengthView
                If Not c.RefIsNothing Then score *= 1.05R
                If score > bestScore Then bestScore = score : altR = c
            End If
        Next
        If altL IsNot Nothing Then fallbackL = altL
        If altR IsNot Nothing Then fallbackR = altR
        If ReferenceEquals(fallbackL, fallbackR) Then Return False

        expected = Math.Abs(fallbackR.MidViewX - fallbackL.MidViewX)
        If expected < 0.02R OrElse expected > 3.0R Then Return False

        leftOut = fallbackL
        rightOut = fallbackR
        If log IsNot Nothing Then P(log, "[EXTENTS_H] route=band_midX_fallback reason=exterior_minmax_unusable")
        Return True
    End Function

    Private Shared Function TrySelectVerticalExtents(lines As List(Of DvLineInfo), vf As ViewFrame,
                                                     ByRef bottomOut As DvLineInfo, ByRef topOut As DvLineInfo,
                                                     ByRef expected As Double, log As DimensionLogger) As Boolean
        bottomOut = Nothing
        topOut = Nothing
        expected = 0R
        If lines Is Nothing OrElse vf Is Nothing Then Return False
        Dim hors As New List(Of DvLineInfo)()
        Dim minLen As Double = Math.Max(vf.Width * 0.28R, 0.01R)
        For Each ln In lines
            If ln Is Nothing OrElse Not ln.IsHorizontal Then Continue For
            If ln.LengthView < minLen Then Continue For
            hors.Add(ln)
        Next
        If hors.Count < 2 Then Return False

        Dim band As Double = Math.Max(vf.Height * 0.08R, 0.004R)
        Dim bottom As DvLineInfo = Nothing
        Dim bestBotScore As Double = Double.NegativeInfinity
        For Each h In hors
            If h.MidViewY <= vf.RangeMinY + band Then
                Dim sc As Double = h.LengthView - Math.Abs(h.MidViewY - vf.RangeMinY) * 2R
                If Not h.RefIsNothing Then sc *= 1.03R
                If sc > bestBotScore Then bestBotScore = sc : bottom = h
            End If
        Next
        Dim top As DvLineInfo = Nothing
        Dim bestTopScore As Double = Double.NegativeInfinity
        For Each h In hors
            If h.MidViewY >= vf.RangeMaxY - band Then
                Dim sc As Double = h.LengthView - Math.Abs(h.MidViewY - vf.RangeMaxY) * 2R
                If Not h.RefIsNothing Then sc *= 1.03R
                If sc > bestTopScore Then bestTopScore = sc : top = h
            End If
        Next
        If bottom Is Nothing Then
            hors.Sort(Function(a, b) a.MidViewY.CompareTo(b.MidViewY))
            bottom = hors(0)
        End If
        If top Is Nothing Then
            hors.Sort(Function(a, b) b.MidViewY.CompareTo(a.MidViewY))
            top = hors(0)
        End If
        If bottom Is Nothing OrElse top Is Nothing OrElse ReferenceEquals(bottom, top) Then Return False
        expected = Math.Abs(top.MidViewY - bottom.MidViewY)
        If expected < 0.005R OrElse expected > 1.0R Then Return False
        bottomOut = bottom
        topOut = top
        Return True
    End Function

    Private Shared Function TryAddDistanceControlled(dims As Dimensions, dv As DrawingView,
                                                     o1 As Object, x1 As Double, y1 As Double,
                                                     o2 As Object, x2 As Double, y2 As Double,
                                                     ByRef created As FrameworkDimension,
                                                     log As DimensionLogger) As Boolean
        created = Nothing
        If dims Is Nothing OrElse o1 Is Nothing OrElse o2 Is Nothing Then Return False
        Dim kpb As Boolean() = {False, True}
        For Each kp1 In kpb
            For Each kp2 In kpb
                Try
                    Dim d As FrameworkDimension = dims.AddDistanceBetweenObjects(o1, x1, y1, 0.0R, kp1, o2, x2, y2, 0.0R, kp2)
                    If d IsNot Nothing Then
                        created = d
                        Return True
                    End If
                Catch ex As Exception
                    LogApiError(log, "Dimensions", "AddDistanceBetweenObjects", ex)
                    Return False
                End Try
            Next
        Next
        For Each kp1 In kpb
            For Each kp2 In kpb
                Try
                    Dim dEx As FrameworkDimension = dims.AddDistanceBetweenObjectsEX(
                        o1, x1, y1, 0.0R, kp1, False,
                        o2, x2, y2, 0.0R, kp2, False)
                    If dEx IsNot Nothing Then
                        created = dEx
                        Return True
                    End If
                Catch ex As Exception
                    LogApiError(log, "Dimensions", "AddDistanceBetweenObjectsEX", ex)
                    Return False
                End Try
            Next
        Next
        Return False
    End Function

    Private Shared Sub TryApplyConstraintFalse(d As FrameworkDimension)
        If d Is Nothing Then Return
        Try
            d.Constraint = False
        Catch
        End Try
    End Sub

    Private Shared Sub TryDraftUpdateAfterDimensionMutation(draft As DraftDocument, app As SolidEdgeFramework.Application)
        Try
            If draft IsNot Nothing Then draft.UpdateAll(True)
        Catch
        End Try
        Try
            If app IsNot Nothing Then app.DoIdle()
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' <see cref="Dimension.ReattachToDrawingView"/> (SDK): reengancha la cota a la vista cuando la referencia está desasociada o mal correlacionada.
    ''' Documentación indica vistas «para las que la cota no está adjunta». Devuelve <see cref="SolidEdgeConstants.DimReattachStatusConstants"/>.
    ''' </summary>
    Private Shared Function TryReattachDimensionToDrawingView(d As FrameworkDimension, dv As DrawingView, tag As String, log As DimensionLogger) As Boolean
        If d Is Nothing OrElse dv Is Nothing Then Return False
        Dim viewName As String = ""
        Try
            viewName = Convert.ToString(dv.Name, CultureInfo.InvariantCulture)
        Catch
            viewName = ""
        End Try
        Try
            Dim st As SolidEdgeConstants.DimReattachStatusConstants = d.ReattachToDrawingView(dv)
            Dim ok As Boolean = (st = SolidEdgeConstants.DimReattachStatusConstants.igDimReattachSucceeded)
            P(log, String.Format(CultureInfo.InvariantCulture,
                "[REATtACH_DVIEW] tag={0} view={1} DimReattachStatus={2}(0=succeeded_1=failed) succeeded={3}",
                tag, viewName, CInt(st).ToString(CultureInfo.InvariantCulture), ok.ToString(CultureInfo.InvariantCulture)))
            Return ok
        Catch ex As Exception
            LogApiError(log, "Dimension", "ReattachToDrawingView", ex)
        End Try
        Try
            Dim retLate As Object = CallByName(d, "ReattachToDrawingView", CallType.Method, CObj(dv))
            Dim code As Integer = Convert.ToInt32(retLate, CultureInfo.InvariantCulture)
            Dim okLb As Boolean = (code = 0)
            P(log, String.Format(CultureInfo.InvariantCulture,
                "[REATtACH_DVIEW] tag={0} view={1} statusInt={2} succeeded={3} route=CallByName",
                tag, viewName, code.ToString(CultureInfo.InvariantCulture), okLb.ToString(CultureInfo.InvariantCulture)))
            Return okLb
        Catch ex2 As Exception
            LogApiError(log, "Dimension", "ReattachToDrawingView(CallByName)", ex2)
            Return False
        End Try
    End Function

    Private Shared Sub TryApplyVerticalMeasurementAxis(d As FrameworkDimension, log As DimensionLogger)
        If d Is Nothing Then Return
        Try
            CallByName(d, "MeasurementAxisEx", CallType.Let, CInt(SolidEdgeConstants.DimAxisModeConstants.igDimAxisModeImplied))
        Catch ex As Exception
            LogApiError(log, "Dimension", "MeasurementAxisEx", ex)
            Return
        End Try
        Try
            CallByName(d, "MeasurementAxisDirection", CallType.Let, True)
        Catch ex As Exception
            LogApiError(log, "Dimension", "MeasurementAxisDirection", ex)
        End Try
    End Sub

    Private Shared Function ReadDimValue(d As FrameworkDimension) As Double
        If d Is Nothing Then Return Double.NaN
        Try
            Return CDbl(d.Value)
        Catch
            Return Double.NaN
        End Try
    End Function

    ''' <summary>
    ''' <see cref="FrameworkDimension.Range"/> a veces devuelve caja en espacio vista (p. ej. anchos de orden modelo)
    ''' mientras el marco de vista en <see cref="ViewFrame"/> está en hoja. Sin normalizar, el score de TrackDistance elige mal.
    ''' </summary>
    Private Shared Function LikelyViewSpaceDimRange(vf As ViewFrame, x1 As Double, y1 As Double, x2 As Double, y2 As Double) As Boolean
        If vf Is Nothing Then Return False
        Dim vw As Double = Math.Max(vf.InitialSheetMaxX - vf.InitialSheetMinX, GeomEps)
        Dim vh As Double = Math.Max(vf.InitialSheetMaxY - vf.InitialSheetMinY, GeomEps)
        Dim dw As Double = Math.Abs(x2 - x1)
        Dim dh As Double = Math.Abs(y2 - y1)
        Const spanFactor As Double = 3.5R
        Return dw > vw * spanFactor OrElse dh > vh * spanFactor
    End Function

    Private Shared Function TryCornersViewToSheetBBox(dv As DrawingView, x1 As Double, y1 As Double, x2 As Double, y2 As Double,
                                                      ByRef minX As Double, ByRef minY As Double, ByRef maxX As Double, ByRef maxY As Double) As Boolean
        minX = 0R : minY = 0R : maxX = 0R : maxY = 0R
        If dv Is Nothing Then Return False
        Dim xs As New List(Of Double)(4)
        Dim ys As New List(Of Double)(4)
        Dim pairs As (Double, Double)() = {(x1, y1), (x2, y1), (x1, y2), (x2, y2)}
        For Each pr In pairs
            Try
                Dim sx As Double, sy As Double
                dv.ViewToSheet(pr.Item1, pr.Item2, sx, sy)
                xs.Add(sx)
                ys.Add(sy)
            Catch
            End Try
        Next
        If xs.Count = 0 Then Return False
        minX = xs.Min() : maxX = xs.Max() : minY = ys.Min() : maxY = ys.Max()
        Return True
    End Function

    ''' <summary>Obtiene bbox de la cota en coordenadas de hoja; si Range parece vista, usa envolvente de esquinas tras ViewToSheet.</summary>
    Private Shared Function TryGetDimensionRangeSheetBBox(d As FrameworkDimension, dv As DrawingView, vf As ViewFrame,
                                                          ByRef minX As Double, ByRef minY As Double, ByRef maxX As Double, ByRef maxY As Double) As Boolean
        minX = 0R : minY = 0R : maxX = 0R : maxY = 0R
        If d Is Nothing Then Return False
        Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
        Try
            d.Range(x1, y1, x2, y2)
        Catch
            Return False
        End Try
        minX = Math.Min(x1, x2)
        maxX = Math.Max(x1, x2)
        minY = Math.Min(y1, y2)
        maxY = Math.Max(y1, y2)
        If dv Is Nothing OrElse vf Is Nothing Then Return True
        If Not LikelyViewSpaceDimRange(vf, x1, y1, x2, y2) Then Return True
        Dim tMinX As Double, tMinY As Double, tMaxX As Double, tMaxY As Double
        If Not TryCornersViewToSheetBBox(dv, x1, y1, x2, y2, tMinX, tMinY, tMaxX, tMaxY) Then Return True
        Dim vw As Double = Math.Max(vf.InitialSheetMaxX - vf.InitialSheetMinX, GeomEps)
        Dim vh As Double = Math.Max(vf.InitialSheetMaxY - vf.InitialSheetMinY, GeomEps)
        Dim tw As Double = tMaxX - tMinX
        Dim th As Double = tMaxY - tMinY
        If tw <= GeomEps OrElse th <= GeomEps Then Return True
        If tw > vw * 6R OrElse th > vh * 6R Then Return True
        minX = tMinX : maxX = tMaxX : minY = tMinY : maxY = tMaxY
        Return True
    End Function

    Private Shared Sub PlaceCleanDimension(app As Application, draft As DraftDocument, d As FrameworkDimension,
                                          vf As ViewFrame, tag As String, axis As String,
                                          log As DimensionLogger)
        If d Is Nothing OrElse vf Is Nothing Then Return
        Dim isHorizDim As Boolean = String.Equals(axis, "horizontal", StringComparison.OrdinalIgnoreCase)
        Dim before As String = ReadRangeStr(d)
        P(log, "[PLACE_START] tag=" & tag)
        P(log, "[PLACE_BEFORE] tag=" & tag & " dimRange=" & before)

        Dim candidatesRaw As Double() = {
            0R,
            0.003R, -0.003R, 0.005R, -0.005R, 0.008R, -0.008R,
            0.012R, -0.012R, 0.015R, -0.015R, 0.018R, -0.018R,
            TrackHistoric, -TrackHistoric
        }
        Dim candList As New List(Of Double)()
        Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        For Each td0 In candidatesRaw
            Dim k As String = td0.ToString(CultureInfo.InvariantCulture)
            If seen.Add(k) Then candList.Add(td0)
        Next
        Dim bestTd As Double = 0R
        Dim bestSc As Double = Double.NegativeInfinity
        Dim bestCombined As Double = Double.NegativeInfinity
        Dim bestAfter As String = before
        For Each td In candList
            P(log, String.Format(CultureInfo.InvariantCulture, "[TRACK_TRY] tag={0} track={1:0.######}", tag, td))
            Try
                d.TrackDistance = td
            Catch ex As Exception
                LogApiError(log, "Dimension", "TrackDistance", ex)
                Continue For
            End Try
            Try
                If draft IsNot Nothing Then draft.UpdateAll(True)
            Catch
            End Try
            Try
                If app IsNot Nothing Then app.DoIdle()
            Catch
            End Try
            Dim rAfter As String = ReadRangeStr(d)
            P(log, String.Format(CultureInfo.InvariantCulture, "[PLACE_AFTER] tag={0} track={1:0.######} dimRange={2}", tag, td, rAfter))
            Dim sc As Double = ScorePlacementOverlap(d, vf, vf.View, isHorizDim)
            Dim combined As Double = sc - TrackProximityWeight * Math.Abs(td)
            Dim reason As String = "overlap_frame_minus_track_proximity"
            P(log, String.Format(CultureInfo.InvariantCulture, "[PLACE_SCORE] tag={0} overlap={1:0.###} combined={2:0.###} reason={3}", tag, sc, combined, reason))
            If combined > bestCombined + GeomEps OrElse
               (Math.Abs(combined - bestCombined) <= 1.0E-6R AndAlso Math.Abs(td) + GeomEps < Math.Abs(bestTd)) Then
                bestCombined = combined
                bestSc = sc
                bestTd = td
                bestAfter = rAfter
            End If
        Next
        If bestSc < 0R Then
            P(log, String.Format(CultureInfo.InvariantCulture, "[PLACE][WARN] tag={0} overlap_score={1:0.###} reason=overlap_frame_sheet", tag, bestSc))
        End If
        Try
            d.TrackDistance = bestTd
        Catch
        End Try
        P(log, String.Format(CultureInfo.InvariantCulture, "[PLACE_SELECTED] tag={0} track={1:0.######} dimRange={2}", tag, bestTd, bestAfter))
    End Sub

    Private Shared Function ReadRangeStr(d As FrameworkDimension) As String
        If d Is Nothing Then Return "(null)"
        Try
            Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
            d.Range(x1, y1, x2, y2)
            Return "(" & Fmt6(x1) & "," & Fmt6(y1) & ")-(" & Fmt6(x2) & "," & Fmt6(y2) & ")"
        Catch ex As Exception
            Return "(unreadable: " & ex.Message & ")"
        End Try
    End Function

    Private Shared Function ScorePlacementOverlap(d As FrameworkDimension, vf As ViewFrame, dv As DrawingView, isHorizontalDimension As Boolean) As Double
        ' Mayor score si el bbox (en hoja) queda cerca del borde de la vista en hoja; Range puede venir en espacio vista.
        Dim minX As Double, minY As Double, maxX As Double, maxY As Double
        If Not TryGetDimensionRangeSheetBBox(d, dv, vf, minX, minY, maxX, maxY) Then
            Return Double.NegativeInfinity
        End If
        Dim vW As Double = Math.Max(vf.InitialSheetMaxX - vf.InitialSheetMinX, 1.0E-12R)
        Dim vH As Double = Math.Max(vf.InitialSheetMaxY - vf.InitialSheetMinY, 1.0E-12R)
        Dim dW As Double = maxX - minX
        Dim dH As Double = maxY - minY
        Dim spanPenalty As Double = 0R
        If dW > 40R * vW OrElse dH > 40R * vH Then spanPenalty -= 1000R
        Dim distOut As Double
        If isHorizontalDimension Then
            Dim above As Double = vf.InitialSheetMaxY - minY
            Dim below As Double = maxY - vf.InitialSheetMinY
            distOut = Math.Min(Math.Abs(above), Math.Abs(below))
        Else
            Dim rightMost As Double = vf.InitialSheetMaxX - minX
            Dim leftMost As Double = maxX - vf.InitialSheetMinX
            distOut = Math.Min(Math.Abs(rightMost), Math.Abs(leftMost))
        End If
        Return -distOut * 10R + spanPenalty
    End Function

    Private Shared Sub CenterTextByDimensionKeypointSweep(app As Application, draft As DraftDocument,
                                                        d As FrameworkDimension, vf As ViewFrame, tag As String, log As DimensionLogger)
        If d Is Nothing Then Return
        Dim n As Integer = 0
        Try
            n = CInt(d.KeyPointCount)
        Catch
            n = 0
        End Try
        P(log, "[TEXT_KP_START] tag=" & tag & " keyPointCount=" & n.ToString(CultureInfo.InvariantCulture))
        Dim rcx As Double = 0R, rcy As Double = 0R
        Dim hasC As Boolean = TryDimRangeCenter(d, If(vf IsNot Nothing, vf.View, Nothing), vf, rcx, rcy)
        Dim bestI As Integer = -1
        Dim bestD As Double = Double.MaxValue
        Dim lim As Integer = Math.Max(0, Math.Min(n, 40))
        For i As Integer = 0 To lim - 1
            Try
                Dim px As Double, py As Double, pz As Double
                Dim kpt As SolidEdgeConstants.KeyPointType
                Dim hdl As SolidEdgeConstants.HandleType
                d.GetKeyPoint(i, px, py, pz, kpt, hdl)
                P(log, String.Format(CultureInfo.InvariantCulture, "[TEXT_KP] tag={0} idx={1} point=({2},{3})", tag, i.ToString(CultureInfo.InvariantCulture), Fmt6(px), Fmt6(py)))
                If hasC Then
                    Dim dd As Double = (px - rcx) * (px - rcx) + (py - rcy) * (py - rcy)
                    If dd < bestD Then
                        bestD = dd
                        bestI = i
                    End If
                End If
            Catch
            End Try
        Next
        If bestI >= 0 Then
            P(log, String.Format(CultureInfo.InvariantCulture, "[TEXT_KP_BEST] tag={0} idx={1} score={2:0.######}", tag, bestI.ToString(CultureInfo.InvariantCulture), bestD))
        End If
        Dim ok As Boolean = False
        Try
            CallByName(d, "CoordinateTextPosition", CallType.Let, 1)
            ok = True
        Catch
        End Try
        Try
            CallByName(d, "SetTextOffsets", CallType.Method, 0.0R, 0.0005R)
            ok = True
        Catch
        End Try
        Try
            If draft IsNot Nothing Then draft.UpdateAll(True)
        Catch
        End Try
        Try
            If app IsNot Nothing Then app.DoIdle()
        Catch
        End Try
        P(log, "[TEXT_CENTER_RESULT] tag=" & tag & " ok=" & ok.ToString(CultureInfo.InvariantCulture))
    End Sub

    Private Shared Function TryDimRangeCenter(d As FrameworkDimension, dv As DrawingView, vf As ViewFrame, ByRef cx As Double, ByRef cy As Double) As Boolean
        cx = 0R
        cy = 0R
        If d Is Nothing Then Return False
        Dim minX As Double, minY As Double, maxX As Double, maxY As Double
        If TryGetDimensionRangeSheetBBox(d, dv, vf, minX, minY, maxX, maxY) Then
            cx = (minX + maxX) * 0.5R
            cy = (minY + maxY) * 0.5R
            Return True
        End If
        Return False
    End Function

    Private Shared Function FinalValidateKeepsDimension(draft As DraftDocument, vf As ViewFrame, d As FrameworkDimension,
                                                        tag As String, expected As Double, sh As Sheet, dv As DrawingView,
                                                        log As DimensionLogger) As Boolean
        If d Is Nothing OrElse vf Is Nothing Then Return False
        Dim val As Double = ReadDimValue(d)
        Dim delta As Double = Math.Abs(val - expected)
        Dim valOk As Boolean = (Not Double.IsNaN(val)) AndAlso delta <= TolValue
        Dim mat As Boolean = False
        Try
            mat = (d.GetDisplayData() IsNot Nothing)
        Catch
            mat = False
        End Try
        Dim vis As Boolean = True
        Try
            vis = CBool(CallByName(d, "Visible", CallType.Get))
        Catch
            vis = True
        End Try
        Dim conn As String = "unknown"
        Dim rc As Integer = 0
        Try
            d.GetRelatedCount(rc)
            conn = If(rc > 0, "True", "False")
        Catch
            conn = "unknown"
        End Try
        Dim connOk As Boolean = String.Equals(conn, "True", StringComparison.OrdinalIgnoreCase)
        Dim rangeReadable As Boolean = True
        Try
            Dim z1, z2, z3, z4 As Double
            d.Range(z1, z2, z3, z4)
            rangeReadable = True
        Catch
            rangeReadable = False
            P(log, String.Format(CultureInfo.InvariantCulture, "[VALIDATE][WARN] tag={0} reason=Dimension.Range_unreadable", tag))
        End Try
        Dim viewRangeOk As Boolean = Not IsViewRangeExpanded(dv, vf)
        Dim outside As Boolean = IsDimOutsideSheetLoose(d, sh, dv, vf)
        Dim absurd As Boolean = IsAbsurdSpan(d, vf, dv)
        Dim coreKeep As Boolean = valOk AndAlso vis AndAlso connOk AndAlso mat
        If Not viewRangeOk Then
            P(log, String.Format(CultureInfo.InvariantCulture, "[VALIDATE][WARN] tag={0} reason=DrawingView.Range_changed_vs_snapshot", tag))
        End If
        If outside Then
            P(log, String.Format(CultureInfo.InvariantCulture, "[VALIDATE][WARN] tag={0} reason=outside_sheet_loose_bbox", tag))
        End If
        If absurd Then
            P(log, String.Format(CultureInfo.InvariantCulture, "[VALIDATE][WARN] tag={0} reason=absurd_span_vs_view", tag))
        End If
        If Not rangeReadable Then
            P(log, String.Format(CultureInfo.InvariantCulture, "[VALIDATE][WARN] tag={0} reason=range_not_checked_for_placement_soft_checks", tag))
        End If

        P(log, String.Format(CultureInfo.InvariantCulture,
            "[VALIDATE] tag={0} value={1:0.######} expected={2:0.######} delta={3:0.######} visible={4} connected={5} materialized={6} rangeReadable={7} ok={8}",
            tag, val, expected, delta, vis.ToString(CultureInfo.InvariantCulture), conn, mat.ToString(CultureInfo.InvariantCulture), rangeReadable.ToString(CultureInfo.InvariantCulture), coreKeep.ToString(CultureInfo.InvariantCulture)))

        Return coreKeep
    End Function

    Private Shared Function IsViewRangeExpanded(dv As DrawingView, vf As ViewFrame) As Boolean
        If dv Is Nothing OrElse vf Is Nothing Then Return False
        Try
            Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
            dv.Range(x1, y1, x2, y2)
            Dim tol As Double = 0.000002R
            Dim mx1 As Double = Math.Min(x1, x2)
            Dim my1 As Double = Math.Min(y1, y2)
            Dim mx2 As Double = Math.Max(x1, x2)
            Dim my2 As Double = Math.Max(y1, y2)
            If mx1 + tol < vf.InitialSheetMinX Then Return True
            If my1 + tol < vf.InitialSheetMinY Then Return True
            If mx2 - tol > vf.InitialSheetMaxX Then Return True
            If my2 - tol > vf.InitialSheetMaxY Then Return True
        Catch
            Return False
        End Try
        Return False
    End Function

    Private Shared Function IsAbsurdSpan(d As FrameworkDimension, vf As ViewFrame, dv As DrawingView) As Boolean
        If vf Is Nothing Then Return False
        Dim minX As Double, minY As Double, maxX As Double, maxY As Double
        If Not TryGetDimensionRangeSheetBBox(d, dv, vf, minX, minY, maxX, maxY) Then Return False
        Dim dW As Double = maxX - minX
        Dim dH As Double = maxY - minY
        Dim vw As Double = Math.Max(vf.InitialSheetMaxX - vf.InitialSheetMinX, GeomEps)
        Dim vh As Double = Math.Max(vf.InitialSheetMaxY - vf.InitialSheetMinY, GeomEps)
        Return (dW > vw * 25R OrElse dH > vh * 25R)
    End Function

    Private Shared Function IsDimOutsideSheetLoose(d As FrameworkDimension, sh As Sheet, dv As DrawingView, vf As ViewFrame) As Boolean
        If d Is Nothing OrElse sh Is Nothing Then Return False
        Dim w As Double = 0, hgt As Double = 0
        If Not TryReadSheetSizeProd(sh, w, hgt) Then Return False
        Dim mnX As Double, mnY As Double, mxX As Double, mxY As Double
        If Not TryGetDimensionRangeSheetBBox(d, dv, vf, mnX, mnY, mxX, mxY) Then Return False
        Return (mnX < -0.001R OrElse mnY < -0.001R OrElse mxX > w + 0.001R OrElse mxY > hgt + 0.001R)
    End Function

    Private Shared Function TryReadSheetSizeProd(sh As Sheet, ByRef width As Double, ByRef height As Double) As Boolean
        width = 0R
        height = 0R
        If sh Is Nothing Then Return False
        Try
            Dim ss As Object = CallByName(sh, "SheetSetup", CallType.Get)
            If ss IsNot Nothing Then
                width = Convert.ToDouble(CallByName(ss, "SheetWidth", CallType.Get), CultureInfo.InvariantCulture)
                height = Convert.ToDouble(CallByName(ss, "SheetHeight", CallType.Get), CultureInfo.InvariantCulture)
                If width > 0R AndAlso height > 0R Then Return True
            End If
        Catch
        End Try
        Try
            width = Convert.ToDouble(CallByName(sh, "Width", CallType.Get), CultureInfo.InvariantCulture)
            height = Convert.ToDouble(CallByName(sh, "Height", CallType.Get), CultureInfo.InvariantCulture)
            If width > 0R AndAlso height > 0R Then Return True
        Catch
        End Try
        Return False
    End Function

    Private Shared Sub SafeDeleteDimension(d As FrameworkDimension, log As DimensionLogger)
        If d Is Nothing Then Return
        Try
            d.Delete()
        Catch ex As Exception
            P(log, "[DELETE_DIM][WARN] " & ex.Message)
        End Try
    End Sub

End Class
