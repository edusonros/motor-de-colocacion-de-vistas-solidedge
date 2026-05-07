Option Strict Off

Imports System.Collections.Generic
Imports System.Globalization
Imports System.Linq
Imports System.Reflection
Imports SolidEdgeDraft
Imports SolidEdgeFramework
Imports SolidEdgeFrameworkSupport
Imports FrameworkDimension = SolidEdgeFrameworkSupport.Dimension

''' <summary>
''' Laboratorio exclusivo de acotación sobre DVGeometry (sin Drop, sin 2D Model).
''' </summary>
Friend NotInheritable Class DVGeometryDimensionPlacementLab

    Private Const DimStyleName As String = "U3,5"
    Private Const AuxLayerName As String = "DV_DIMLAB_AUX"
    Private Const PlacementOffset As Double = 0.02R
    Private Const TolAxis As Double = 0.0002R

    Private Shared ReadOnly ReflectInvoke As BindingFlags =
        BindingFlags.InvokeMethod Or BindingFlags.Public Or BindingFlags.Instance

    Private NotInheritable Class DvLineInfo
        Public Id As Integer
        Public Obj As Object
        Public X1 As Double
        Public Y1 As Double
        Public X2 As Double
        Public Y2 As Double
        Public Length As Double
        Public AngleDeg As Double
        Public Orientation As String
    End Class

    Private NotInheritable Class InterestPoint
        Public Id As Integer
        Public PointType As String
        Public SourceType As String
        Public SourceEntity As String
        Public XView As Double
        Public YView As Double
        Public XSheet As Double
        Public YSheet As Double
    End Class

    Private NotInheritable Class CandidateDim
        Public Kind As String
        Public O1 As Object
        Public O2 As Object
        Public X1 As Double
        Public Y1 As Double
        Public X2 As Double
        Public Y2 As Double
        Public Expected As Double
    End Class

    Public Shared Sub Run(app As Application, draft As DraftDocument, logSink As Action(Of String), Optional debugSave As Boolean = False)
        Dim L = Sub(m As String)
                    Try
                        logSink?.Invoke(m)
                    Catch
                    End Try
                End Sub

        L("[DV_DIMLAB][START]")
        If draft Is Nothing Then
            L("[DV_DIMLAB][ABORT] reason=no_draft")
            Return
        End If
        Dim p As String = SafeStr(CallByNameSafe(draft, "FullName"))
        L("[DV_DIMLAB][DFT] path=" & p)
        L("[DV_DIMLAB][NO_DROP] true")
        L("[DV_DIMLAB][NO_2DMODEL] true")

        Dim src As Sheet = ResolveSourceSheet(draft)
        If src Is Nothing Then
            L("[DV_DIMLAB][ABORT] reason=no_source_sheet")
            Return
        End If

        Dim views As List(Of DrawingView) = GetSheetViews(src)
        L("[DV_DIMLAB][SOURCE_SHEET] name=" & SafeStr(CallByNameSafe(src, "Name")) & " views=" & views.Count.ToString(CultureInfo.InvariantCulture))

        Dim viewsProcessed As Integer = 0
        Dim viewsSelected As Integer = 0
        Dim dvLinesFound As Integer = 0
        Dim dvArcsFound As Integer = 0
        Dim dvPointsFound As Integer = 0
        Dim candH As Integer = 0
        Dim candV As Integer = 0
        Dim candPartial As Integer = 0

        Dim tryA As Integer = 0, okA As Integer = 0, failA As Integer = 0
        Dim tryB As Integer = 0, okB As Integer = 0, failB As Integer = 0
        Dim tryC As Integer = 0, okC As Integer = 0, failC As Integer = 0
        Dim tryD As Integer = 0, okD As Integer = 0, failD As Integer = 0
        Dim tryE As Integer = 0, okE As Integer = 0, failE As Integer = 0
        Dim dimsConnected As Integer = 0
        Dim dimsFloatingVisible As Integer = 0
        Dim dimsAuxFallback As Integer = 0
        Dim bestMethod As String = "none"

        Dim dims As Dimensions = Nothing
        Try : dims = CType(CallByName(src, "Dimensions", CallType.Get), Dimensions) : Catch : dims = Nothing : End Try
        If dims Is Nothing Then
            L("[DV_DIMLAB][ABORT] reason=no_dimensions_collection")
            Return
        End If

        Dim bridge As New Logger(Sub(msg) L(msg))
        Dim dimLog As New DimensionLogger(bridge)
        Dim styleObj As Object = ResolveStyleObject(draft, DimStyleName, L)
        If styleObj Is Nothing Then
            L("[DV_DIMLAB][STYLE][FAIL] reason=not_found/type_mismatch")
        End If
        Dim auxLayer As Object = EnsureLayer(src, AuxLayerName)

        Dim globalPoints As New List(Of InterestPoint)()

        For i As Integer = 0 To views.Count - 1
            Dim dv As DrawingView = views(i)
            viewsProcessed += 1

            Dim nm As String = SafeStr(CallByNameSafe(dv, "Name"))
            Dim typ As String = SafeStr(CallByNameSafe(dv, "DrawingViewType"))
            Dim sc As String = SafeStr(CallByNameSafe(dv, "ScaleFactor"))
            If String.IsNullOrWhiteSpace(sc) Then sc = SafeStr(CallByNameSafe(dv, "Scale"))
            L("[DV_DIMLAB][VIEW][INFO] idx=" & (i + 1).ToString(CultureInfo.InvariantCulture) & " name=" & nm & " type=" & typ & " scale=" & sc)

            Dim rx1 As Double = 0, ry1 As Double = 0, rx2 As Double = 0, ry2 As Double = 0
            Try
                dv.Range(rx1, ry1, rx2, ry2)
                L("[DV_DIMLAB][VIEW][RANGE] xmin=" & F(Math.Min(rx1, rx2)) & " ymin=" & F(Math.Min(ry1, ry2)) &
                    " xmax=" & F(Math.Max(rx1, rx2)) & " ymax=" & F(Math.Max(ry1, ry2)))
            Catch ex As Exception
                L("[DV_DIMLAB][VIEW][RANGE] FAIL NO_CONFIRMADO " & ex.Message)
            End Try

            Dim nL As Integer = SafeCount(CallByNameSafe(dv, "DVLines2d"))
            Dim nA As Integer = SafeCount(CallByNameSafe(dv, "DVArcs2d"))
            Dim nC As Integer = SafeCount(CallByNameSafe(dv, "DVCircles2d"))
            Dim nS As Integer = SafeCount(CallByNameSafe(dv, "DVBSplineCurves2d"))
            If nS = 0 Then nS = SafeCount(CallByNameSafe(dv, "DVBSplines2d"))
            Dim nP As Integer = SafeCount(CallByNameSafe(dv, "DVPoints2d"))
            L("[DV_DIMLAB][VIEW][DVCOUNT] lines=" & nL.ToString(CultureInfo.InvariantCulture) &
              " arcs=" & nA.ToString(CultureInfo.InvariantCulture) &
              " circles=" & nC.ToString(CultureInfo.InvariantCulture) &
              " splines=" & nS.ToString(CultureInfo.InvariantCulture) &
              " points=" & nP.ToString(CultureInfo.InvariantCulture))

            dvLinesFound += nL
            dvArcsFound += nA
            dvPointsFound += nP
        Next

        Dim targetView As DrawingView = SelectPreferredView(views)
        If targetView Is Nothing Then
            L("[DV_DIMLAB][ABORT] reason=no_target_view")
            Return
        End If
        viewsSelected = 1
        L("[DV_DIMLAB][VIEW][SELECTED] name=" & SafeStr(CallByNameSafe(targetView, "Name")) & " reason=prefer_4411_or_max_HV_non_iso")

        Dim targetLines As List(Of DvLineInfo) = InspectDvLines(targetView, L)
        If targetLines.Count = 0 Then
            L("[DV_DIMLAB][ABORT] reason=no_valid_dvline_coords")
            L("[DV_DIMLAB][SUMMARY]")
            L("views_processed=" & viewsProcessed.ToString(CultureInfo.InvariantCulture))
            L("views_selected=" & viewsSelected.ToString(CultureInfo.InvariantCulture))
            L("dvlines_found=" & dvLinesFound.ToString(CultureInfo.InvariantCulture))
            L("dvarcs_found=" & dvArcsFound.ToString(CultureInfo.InvariantCulture))
            L("points_found=0")
            L("candidates_h_total=0")
            L("candidates_v_total=0")
            L("candidates_partial=0")
            L("dims_try_A=0")
            L("dims_ok_A=0")
            L("dims_fail_A=0")
            L("dims_try_B=0")
            L("dims_ok_B=0")
            L("dims_fail_B=0")
            L("dims_try_C=0")
            L("dims_ok_C=0")
            L("dims_fail_C=0")
            L("dims_try_D=0")
            L("dims_ok_D=0")
            L("dims_fail_D=0")
            L("dims_connected=0")
            L("dims_floating_visible=0")
            L("dims_aux_fallback=0")
            L("best_method=none")
            L("recommended_next_step=resolver_extraccion_dvline_coords")
            Return
        End If
        Dim targetPoints As List(Of InterestPoint) = BuildInterestPoints(targetView, targetLines, L)
        globalPoints.AddRange(targetPoints)
        Dim targetCandidates As List(Of CandidateDim) = BuildCandidates(targetView, targetLines, L, candH, candV, candPartial)

        ' A/B/C sobre candidatos de la vista objetivo
        For Each c In targetCandidates
                If Math.Abs(c.Expected) <= 1.0E-9R Then
                    L("[DV_DIMLAB][DVLINE][REJECT_ZERO_COORDS] reason=expected_zero view=" & SafeStr(CallByNameSafe(targetView, "Name")) & " kind=" & c.Kind)
                    Continue For
                End If

                Dim d As FrameworkDimension = Nothing
                tryA += 1
                L("[DV_DIMLAB][DIM][TRY_A] method=AddDistanceBetweenObjects objects=DVLine2d,DVLine2d view=" & SafeStr(CallByNameSafe(targetView, "Name")) & " kind=" & c.Kind)
                Dim mUsed As String = ""
                Dim ok = DimensionPlacementEngine.TryAddDistanceBetweenObjectsGeneric(
                    dims, c.O1, c.O2, c.X1, c.Y1, c.X2, c.Y2, dimLog, If(c.Kind.StartsWith("H"), "horizontal", "vertical"),
                    mUsed, "DV_DIMLAB_" & SafeStr(CallByNameSafe(targetView, "Name")) & "_" & c.Kind, Nothing, targetView, True, d)
                If ok AndAlso d IsNot Nothing Then
                    okA += 1
                    ValidateDim(d, c.Expected, L, "A", dimsConnected, dimsFloatingVisible, bestMethod)
                    ApplyStyle(d, styleObj, L)
                Else
                    failA += 1
                    L("[DV_DIMLAB][DIM][FAIL_A] view=" & SafeStr(CallByNameSafe(targetView, "Name")) & " kind=" & c.Kind & " error=no_dimension_returned")
                End If

                tryB += 1
                L("[DV_DIMLAB][DIM][TRY_B] method=AddDistanceBetweenObjects objects=Reference/GraphicMember view=" & SafeStr(CallByNameSafe(targetView, "Name")) & " kind=" & c.Kind)
                Dim rb1 As Object = ResolveRefOrGraphic(c.O1, L, "o1")
                Dim rb2 As Object = ResolveRefOrGraphic(c.O2, L, "o2")
                d = Nothing
                ok = False
                If rb1 IsNot Nothing AndAlso rb2 IsNot Nothing Then
                    ok = DimensionPlacementEngine.TryAddDistanceBetweenObjectsGeneric(
                        dims, rb1, rb2, c.X1, c.Y1, c.X2, c.Y2, dimLog, If(c.Kind.StartsWith("H"), "horizontal", "vertical"),
                        mUsed, "DV_DIMLAB_" & SafeStr(CallByNameSafe(targetView, "Name")) & "_" & c.Kind & "_B", Nothing, targetView, True, d)
                End If
                If ok AndAlso d IsNot Nothing Then
                    okB += 1
                    ValidateDim(d, c.Expected, L, "B", dimsConnected, dimsFloatingVisible, bestMethod)
                    ApplyStyle(d, styleObj, L)
                Else
                    failB += 1
                    L("[DV_DIMLAB][DIM][FAIL_B] view=" & SafeStr(CallByNameSafe(targetView, "Name")) & " kind=" & c.Kind & " error=no_dimension_returned")
                End If

                tryC += 1
                L("[DV_DIMLAB][DIM][TRY_C] method=ADBO_DVLine_with_calculated_proximity_points view=" & SafeStr(CallByNameSafe(targetView, "Name")) & " kind=" & c.Kind)
                Dim px1 As Double = c.X1, py1 As Double = c.Y1, px2 As Double = c.X2, py2 As Double = c.Y2
                If c.Kind.StartsWith("H") Then
                    py1 = py1 + PlacementOffset
                    py2 = py2 + PlacementOffset
                Else
                    px1 = px1 - PlacementOffset
                    px2 = px2 - PlacementOffset
                End If
                d = Nothing
                ok = DimensionPlacementEngine.TryAddDistanceBetweenObjectsGeneric(
                    dims, c.O1, c.O2, px1, py1, px2, py2, dimLog, If(c.Kind.StartsWith("H"), "horizontal", "vertical"),
                    mUsed, "DV_DIMLAB_" & SafeStr(CallByNameSafe(targetView, "Name")) & "_" & c.Kind & "_C", Nothing, targetView, True, d)
                If ok AndAlso d IsNot Nothing Then
                    okC += 1
                    ValidateDim(d, c.Expected, L, "C", dimsConnected, dimsFloatingVisible, bestMethod)
                    ApplyStyle(d, styleObj, L)
                Else
                    failC += 1
                    L("[DV_DIMLAB][DIM][FAIL_C] view=" & SafeStr(CallByNameSafe(targetView, "Name")) & " kind=" & c.Kind & " error=no_dimension_returned")
                End If
        Next

        ' D: fallback con línea auxiliar en hoja desde extremos de puntos.
        tryD += 1
        L("[DV_DIMLAB][DIM][TRY_D] method=AddLength_AuxLine_From_DVPoints")
        Dim auxDim As FrameworkDimension = TryCreateAuxDimFromPoints(src, dims, globalPoints, auxLayer, L)
        If auxDim IsNot Nothing Then
            okD += 1
            dimsAuxFallback += 1
            ValidateDim(auxDim, Double.NaN, L, "D", dimsConnected, dimsFloatingVisible, bestMethod)
            ApplyStyle(auxDim, styleObj, L)
        Else
            failD += 1
            L("[DV_DIMLAB][DIM][FAIL_D] error=no_dimension_returned")
        End If

        ' E: intentos con DVPoints2d / keypoints si existen.
        tryE += 1
        L("[DV_DIMLAB][DIM][TRY_E] method=KeyPoint/DVPoint based")
        Dim eAny As Boolean = TryDvPointsBased(targetView, dims, dimLog, L, dimsConnected, dimsFloatingVisible, bestMethod, styleObj)
        If eAny Then
            okE += 1
        Else
            failE += 1
            L("[DV_DIMLAB][DIM][FAIL_E] error=no_dimension_returned")
        End If

        L("[DV_DIMLAB][SUMMARY]")
        L("views_processed=" & viewsProcessed.ToString(CultureInfo.InvariantCulture))
        L("views_selected=" & viewsSelected.ToString(CultureInfo.InvariantCulture))
        L("dvlines_found=" & dvLinesFound.ToString(CultureInfo.InvariantCulture))
        L("dvarcs_found=" & dvArcsFound.ToString(CultureInfo.InvariantCulture))
        L("points_found=" & globalPoints.Count.ToString(CultureInfo.InvariantCulture))
        L("candidates_h_total=" & candH.ToString(CultureInfo.InvariantCulture))
        L("candidates_v_total=" & candV.ToString(CultureInfo.InvariantCulture))
        L("candidates_partial=" & candPartial.ToString(CultureInfo.InvariantCulture))
        L("dims_try_A=" & tryA.ToString(CultureInfo.InvariantCulture))
        L("dims_ok_A=" & okA.ToString(CultureInfo.InvariantCulture))
        L("dims_fail_A=" & failA.ToString(CultureInfo.InvariantCulture))
        L("dims_try_B=" & tryB.ToString(CultureInfo.InvariantCulture))
        L("dims_ok_B=" & okB.ToString(CultureInfo.InvariantCulture))
        L("dims_fail_B=" & failB.ToString(CultureInfo.InvariantCulture))
        L("dims_try_C=" & tryC.ToString(CultureInfo.InvariantCulture))
        L("dims_ok_C=" & okC.ToString(CultureInfo.InvariantCulture))
        L("dims_fail_C=" & failC.ToString(CultureInfo.InvariantCulture))
        L("dims_try_D=" & tryD.ToString(CultureInfo.InvariantCulture))
        L("dims_ok_D=" & okD.ToString(CultureInfo.InvariantCulture))
        L("dims_fail_D=" & failD.ToString(CultureInfo.InvariantCulture))
        L("dims_connected=" & dimsConnected.ToString(CultureInfo.InvariantCulture))
        L("dims_floating_visible=" & dimsFloatingVisible.ToString(CultureInfo.InvariantCulture))
        L("dims_aux_fallback=" & dimsAuxFallback.ToString(CultureInfo.InvariantCulture))
        L("best_method=" & bestMethod)
        L("recommended_next_step=consolidar_método_con_mayor_OK_para_motor_definitivo")

        If debugSave Then
            Try : draft.Save() : Catch : End Try
        End If
    End Sub

    Private Shared Function SelectPreferredView(views As List(Of DrawingView)) As DrawingView
        If views Is Nothing OrElse views.Count = 0 Then Return Nothing

        For Each dv In views
            Dim nm As String = SafeStr(CallByNameSafe(dv, "Name"))
            If nm.IndexOf("4411", StringComparison.OrdinalIgnoreCase) >= 0 Then Return dv
        Next

        Dim best As DrawingView = Nothing
        Dim bestScore As Integer = Integer.MinValue
        For Each dv In views
            Dim typ As String = SafeStr(CallByNameSafe(dv, "DrawingViewType"))
            Dim isIso As Boolean = IsLikelyIso(dv, typ)
            If isIso Then Continue For
            Dim hv As Integer = CountHvLines(dv)
            Dim nL As Integer = SafeCount(CallByNameSafe(dv, "DVLines2d"))
            Dim score As Integer = hv * 1000 + nL
            If score > bestScore Then
                bestScore = score
                best = dv
            End If
        Next

        If best IsNot Nothing Then Return best
        Return views(0)
    End Function

    Private Shared Function ResolveSourceSheet(draft As DraftDocument) As Sheet
        Dim n As Integer = SafeCount(CallByNameSafe(draft, "Sheets"))
        Dim hoja1 As Sheet = Nothing
        Dim best As Sheet = Nothing
        Dim bestCount As Integer = -1
        For i As Integer = 1 To n
            Dim sh As Sheet = Nothing
            Try : sh = CType(CallByName(draft.Sheets, "Item", CallType.Method, i), Sheet) : Catch : sh = Nothing : End Try
            If sh Is Nothing Then Continue For
            Dim nm As String = SafeStr(CallByNameSafe(sh, "Name"))
            Dim vc As Integer = SafeCount(CallByNameSafe(sh, "DrawingViews"))
            If String.Equals(nm, "Hoja1", StringComparison.OrdinalIgnoreCase) AndAlso vc > 0 Then hoja1 = sh
            If vc > bestCount Then
                best = sh
                bestCount = vc
            End If
        Next
        If hoja1 IsNot Nothing Then Return hoja1
        Return best
    End Function

    Private Shared Function GetSheetViews(sh As Sheet) As List(Of DrawingView)
        Dim list As New List(Of DrawingView)()
        If sh Is Nothing Then Return list
        Dim dvs As Object = CallByNameSafe(sh, "DrawingViews")
        Dim n As Integer = SafeCount(dvs)
        For i As Integer = 1 To n
            Dim dv As DrawingView = Nothing
            Try : dv = CType(CallByName(dvs, "Item", CallType.Method, i), DrawingView) : Catch : End Try
            If dv IsNot Nothing Then list.Add(dv)
        Next
        Return list
    End Function

    Private Shared Function CountHvLines(dv As DrawingView) As Integer
        Dim countHv As Integer = 0
        Dim lines As Object = CallByNameSafe(dv, "DVLines2d")
        For i As Integer = 1 To SafeCount(lines)
            Dim ln As Object = Nothing
            Try : ln = CallByName(lines, "Item", CallType.Method, i) : Catch : ln = Nothing : End Try
            If ln Is Nothing Then Continue For
            Dim x1 As Double = 0, y1 As Double = 0, x2 As Double = 0, y2 As Double = 0
            If Not TryReadDvLineEndpoints(ln, x1, y1, x2, y2) Then Continue For
            Dim dx As Double = Math.Abs(x2 - x1)
            Dim dy As Double = Math.Abs(y2 - y1)
            If dx <= TolAxis OrElse dy <= TolAxis Then countHv += 1
        Next
        Return countHv
    End Function

    Private Shared Function InspectDvLines(dv As DrawingView, log As Action(Of String)) As List(Of DvLineInfo)
        Dim out As New List(Of DvLineInfo)()
        Dim lines As Object = CallByNameSafe(dv, "DVLines2d")
        For i As Integer = 1 To SafeCount(lines)
            Dim ln As Object = Nothing
            Try : ln = CallByName(lines, "Item", CallType.Method, i) : Catch : ln = Nothing : End Try
            If ln Is Nothing Then Continue For

            Dim x1 As Double = 0, y1 As Double = 0, x2 As Double = 0, y2 As Double = 0
            If Not TryReadDvLineEndpointsDetailed(ln, i, x1, y1, x2, y2, log) Then Continue For
            If IsZeroCoordSet(x1, y1, x2, y2) Then
                log("[DV_DIMLAB][DVLINE][REJECT_ZERO_COORDS] id=" & i.ToString(CultureInfo.InvariantCulture) &
                    " x1=" & F(x1) & " y1=" & F(y1) & " x2=" & F(x2) & " y2=" & F(y2))
                Continue For
            End If

            Dim dx As Double = x2 - x1
            Dim dy As Double = y2 - y1
            Dim len As Double = Math.Sqrt(dx * dx + dy * dy)
            Dim angle As Double = Math.Atan2(dy, dx) * 180.0R / Math.PI
            Dim orient As String = "inclined"
            If Math.Abs(dy) <= TolAxis Then orient = "horizontal"
            If Math.Abs(dx) <= TolAxis Then orient = "vertical"

            log("[DV_DIMLAB][DVLINE] id=" & i.ToString(CultureInfo.InvariantCulture) &
                " x1=" & F(x1) & " y1=" & F(y1) & " x2=" & F(x2) & " y2=" & F(y2) &
                " length=" & F(len) & " angle=" & F(angle) & " orientation=" & orient)
            log("[DV_DIMLAB][DVLINE][COM] id=" & i.ToString(CultureInfo.InvariantCulture) & " type=" & ln.GetType().FullName)

            ProbeProperty(ln, i, "Parent", log)
            ProbeProperty(ln, i, "Layer", log)
            ProbeProperty(ln, i, "Style", log)
            ProbeProperty(ln, i, "ModelMember", log)
            ProbeProperty(ln, i, "ReferenceKey", log)
            ProbeProperty(ln, i, "RefKey", log)
            ProbeProperty(ln, i, "GraphicMember", log)
            ProbeMethod(ln, i, "GetReferenceToGraphicMember", log)
            ProbeMethod(ln, i, "GetKeyPoint", log)
            ProbeMethod(ln, i, "KeyPoints", log)
            ProbeMethod(ln, i, "Connect", log)
            ProbeMethod(ln, i, "Relations", log)

            Dim refObj As Object = Nothing
            Try : refObj = CallByName(ln, "Reference", CallType.Get) : Catch : refObj = Nothing : End Try
            log("[DV_DIMLAB][DVLINE][REFERENCE] id=" & i.ToString(CultureInfo.InvariantCulture) &
                " ok=" & (refObj IsNot Nothing).ToString(CultureInfo.InvariantCulture) &
                " type=" & If(refObj Is Nothing, "", refObj.GetType().FullName))
            Dim gmObj As Object = Nothing
            Try : gmObj = CallByName(ln, "GraphicMember", CallType.Get) : Catch : gmObj = Nothing : End Try
            log("[DV_DIMLAB][DVLINE][GRAPHIC_MEMBER] id=" & i.ToString(CultureInfo.InvariantCulture) &
                " ok=" & (gmObj IsNot Nothing).ToString(CultureInfo.InvariantCulture) &
                " type=" & If(gmObj Is Nothing, "", gmObj.GetType().FullName))

            out.Add(New DvLineInfo With {.Id = i, .Obj = ln, .X1 = x1, .Y1 = y1, .X2 = x2, .Y2 = y2, .Length = len, .AngleDeg = angle, .Orientation = orient})
        Next
        Return out
    End Function

    Private Shared Function BuildInterestPoints(dv As DrawingView, lines As List(Of DvLineInfo), log As Action(Of String)) As List(Of InterestPoint)
        Dim pts As New List(Of InterestPoint)()
        Dim pid As Integer = 0
        Dim mode As String = "Unknown"
        If lines.Count > 0 Then
            mode = DetectCoordinateMode(dv, lines(0))
        End If
        log("[DV_DIMLAB][TRANSFORM][MODE] detected=" & mode)

        For Each ln In lines
            pid += 1
            AddPoint(pts, pid, "start", ln, ln.X1, ln.Y1, dv, log)
            pid += 1
            AddPoint(pts, pid, "end", ln, ln.X2, ln.Y2, dv, log)
            pid += 1
            AddPoint(pts, pid, "midpoint", ln, (ln.X1 + ln.X2) * 0.5R, (ln.Y1 + ln.Y2) * 0.5R, dv, log)
        Next

        Dim minX As Double = Double.PositiveInfinity, maxX As Double = Double.NegativeInfinity
        Dim minY As Double = Double.PositiveInfinity, maxY As Double = Double.NegativeInfinity
        For Each ln In lines
            minX = Math.Min(minX, Math.Min(ln.X1, ln.X2))
            maxX = Math.Max(maxX, Math.Max(ln.X1, ln.X2))
            minY = Math.Min(minY, Math.Min(ln.Y1, ln.Y2))
            maxY = Math.Max(maxY, Math.Max(ln.Y1, ln.Y2))
        Next
        If Not Double.IsInfinity(minX) Then
            pid += 1 : AddPoint(pts, pid, "bbox_left", Nothing, minX, (minY + maxY) * 0.5R, dv, log)
            pid += 1 : AddPoint(pts, pid, "bbox_right", Nothing, maxX, (minY + maxY) * 0.5R, dv, log)
            pid += 1 : AddPoint(pts, pid, "bbox_bottom", Nothing, (minX + maxX) * 0.5R, minY, dv, log)
            pid += 1 : AddPoint(pts, pid, "bbox_top", Nothing, (minX + maxX) * 0.5R, maxY, dv, log)
        End If
        Return pts
    End Function

    Private Shared Function BuildCandidates(dv As DrawingView, lines As List(Of DvLineInfo), log As Action(Of String),
                                            ByRef cH As Integer, ByRef cV As Integer, ByRef cPartial As Integer) As List(Of CandidateDim)
        Dim out As New List(Of CandidateDim)()
        If lines.Count = 0 Then Return out
        Dim rx1 As Double = 0, ry1 As Double = 0, rx2 As Double = 0, ry2 As Double = 0
        Try : dv.Range(rx1, ry1, rx2, ry2) : Catch : End Try
        Dim minX As Double = Math.Min(rx1, rx2), maxX As Double = Math.Max(rx1, rx2)
        Dim minY As Double = Math.Min(ry1, ry2), maxY As Double = Math.Max(ry1, ry2)

        Dim verticals = lines.Where(Function(l) l.Orientation = "vertical").OrderBy(Function(l) (l.X1 + l.X2) * 0.5R).ToList()
        Dim horizontals = lines.Where(Function(l) l.Orientation = "horizontal").OrderBy(Function(l) (l.Y1 + l.Y2) * 0.5R).ToList()
        If verticals.Count >= 2 Then
            Dim left = verticals.First()
            Dim right = verticals.Last()
            Dim yRef = (Math.Max(Math.Min(left.Y1, left.Y2), Math.Min(right.Y1, right.Y2)) +
                        Math.Min(Math.Max(left.Y1, left.Y2), Math.Max(right.Y1, right.Y2))) * 0.5R
            If Double.IsNaN(yRef) OrElse Double.IsInfinity(yRef) Then yRef = (minY + maxY) * 0.5R
            Dim yTrack As Double = minY - PlacementOffset
            out.Add(New CandidateDim With {
                .Kind = "H_TOTAL",
                .O1 = left.Obj, .O2 = right.Obj,
                .X1 = (left.X1 + left.X2) * 0.5R, .Y1 = yTrack,
                .X2 = (right.X1 + right.X2) * 0.5R, .Y2 = yTrack,
                .Expected = Math.Abs(((right.X1 + right.X2) * 0.5R) - ((left.X1 + left.X2) * 0.5R))
            })
            cH += 1
            log("[DV_DIMLAB][CAND][H_TOTAL] view=" & SafeStr(CallByNameSafe(dv, "Name")) &
                " leftEntity=DVLine2d rightEntity=DVLine2d expected=" & F(out.Last().Expected))
            For i As Integer = 0 To verticals.Count - 2
                cPartial += 1
                log("[DV_DIMLAB][CAND][H_PARTIAL] view=" & SafeStr(CallByNameSafe(dv, "Name")) &
                    " e1=DVLine2d e2=DVLine2d expected=" & F(Math.Abs(verticals(i + 1).X1 - verticals(i).X1)))
            Next
        End If
        If horizontals.Count >= 2 Then
            Dim bot = horizontals.First()
            Dim top = horizontals.Last()
            Dim xTrack As Double = minX - PlacementOffset
            out.Add(New CandidateDim With {
                .Kind = "V_TOTAL",
                .O1 = bot.Obj, .O2 = top.Obj,
                .X1 = xTrack, .Y1 = (bot.Y1 + bot.Y2) * 0.5R,
                .X2 = xTrack, .Y2 = (top.Y1 + top.Y2) * 0.5R,
                .Expected = Math.Abs(((top.Y1 + top.Y2) * 0.5R) - ((bot.Y1 + bot.Y2) * 0.5R))
            })
            cV += 1
            log("[DV_DIMLAB][CAND][V_TOTAL] view=" & SafeStr(CallByNameSafe(dv, "Name")) &
                " bottomEntity=DVLine2d topEntity=DVLine2d expected=" & F(out.Last().Expected))
            For i As Integer = 0 To horizontals.Count - 2
                cPartial += 1
                log("[DV_DIMLAB][CAND][V_PARTIAL] view=" & SafeStr(CallByNameSafe(dv, "Name")) &
                    " e1=DVLine2d e2=DVLine2d expected=" & F(Math.Abs(horizontals(i + 1).Y1 - horizontals(i).Y1)))
            Next
        End If
        Return out
    End Function

    Private Shared Function TryCreateAuxDimFromPoints(sh As Sheet, dims As Dimensions, points As List(Of InterestPoint),
                                                      auxLayer As Object, log As Action(Of String)) As FrameworkDimension
        If sh Is Nothing OrElse dims Is Nothing OrElse points Is Nothing OrElse points.Count = 0 Then Return Nothing
        Dim pLeft = points.OrderBy(Function(p) p.XSheet).FirstOrDefault()
        Dim pRight = points.OrderByDescending(Function(p) p.XSheet).FirstOrDefault()
        If pLeft Is Nothing OrElse pRight Is Nothing Then Return Nothing
        Dim y As Double = Math.Min(pLeft.YSheet, pRight.YSheet) - PlacementOffset
        Dim aux As Object = Nothing
        Try : aux = CallByName(sh.Lines2d, "AddBy2Points", CallType.Method, pLeft.XSheet, y, pRight.XSheet, y) : Catch : aux = Nothing : End Try
        If aux Is Nothing Then Return Nothing
        Try : CallByName(aux, "Layer", CallType.Let, auxLayer) : Catch : End Try
        log("[DV_DIMLAB][AUX_LINE][CREATE] x1=" & F(pLeft.XSheet) & " y1=" & F(y) & " x2=" & F(pRight.XSheet) & " y2=" & F(y))
        log("[DV_DIMLAB][AUX_LINE][KEEP] reason=dimension_depends_on_aux_geometry")
        Dim d As FrameworkDimension = Nothing
        Try : d = TryCast(CallByName(dims, "AddLength", CallType.Method, aux), FrameworkDimension) : Catch : d = Nothing : End Try
        If d IsNot Nothing Then
            log("[DV_DIMLAB][DIM][OK_D] value=" & SafeStr(CallByNameSafe(d, "Value")))
        End If
        Return d
    End Function

    Private Shared Function TryDvPointsBased(dv As DrawingView, dims As Dimensions, dimLog As DimensionLogger, log As Action(Of String),
                                             ByRef dimsConnected As Integer, ByRef dimsFloatingVisible As Integer,
                                             ByRef bestMethod As String, styleObj As Object) As Boolean
        Dim pts As Object = CallByNameSafe(dv, "DVPoints2d")
        If SafeCount(pts) >= 2 Then
            Dim p1 As Object = Nothing, p2 As Object = Nothing
            Try : p1 = CallByName(pts, "Item", CallType.Method, 1) : Catch : p1 = Nothing : End Try
            Try : p2 = CallByName(pts, "Item", CallType.Method, 2) : Catch : p2 = Nothing : End Try
            If p1 IsNot Nothing AndAlso p2 IsNot Nothing Then
                Dim d As FrameworkDimension = Nothing
                Dim methodUsed As String = ""
                Dim x1 As Double = SafeD(CallByNameSafe(p1, "X")), y1 As Double = SafeD(CallByNameSafe(p1, "Y"))
                Dim x2 As Double = SafeD(CallByNameSafe(p2, "X")), y2 As Double = SafeD(CallByNameSafe(p2, "Y"))
                Dim ok = DimensionPlacementEngine.TryAddDistanceBetweenObjectsGeneric(
                    dims, p1, p2, x1, y1, x2, y2, dimLog, "horizontal", methodUsed, "DV_DIMLAB_E", Nothing, dv, True, d)
                If ok AndAlso d IsNot Nothing Then
                    ValidateDim(d, Double.NaN, log, "E", dimsConnected, dimsFloatingVisible, bestMethod)
                    ApplyStyle(d, styleObj, log)
                    log("[DV_DIMLAB][DIM][OK_E] value=" & SafeStr(CallByNameSafe(d, "Value")))
                    Return True
                End If
            End If
        End If
        Return False
    End Function

    Private Shared Sub ValidateDim(d As FrameworkDimension, expected As Double, log As Action(Of String), tag As String,
                                   ByRef dimsConnected As Integer, ByRef dimsFloatingVisible As Integer, ByRef bestMethod As String)
        If d Is Nothing Then Return
        Dim id As String = SafeStr(CallByNameSafe(d, "Name"))
        Dim val As Double = SafeD(CallByNameSafe(d, "Value"))
        Dim delta As Double = If(Double.IsNaN(expected), Double.NaN, Math.Abs(val - expected))
        log("[DV_DIMLAB][DIM][VALIDATE] id=" & id & " value=" & F(val) & " expected=" & F(expected) & " delta=" & F(delta))

        Dim status As String = SafeStr(CallByNameSafe(d, "StatusOfDimension"))
        If String.IsNullOrWhiteSpace(status) Then status = SafeStr(CallByNameSafe(d, "Status"))
        log("[DV_DIMLAB][DIM][STATUS] id=" & id & " status=" & status)

        Dim relatedObj As Object = Nothing
        Try : relatedObj = CallByName(d, "GetRelatedObjects", CallType.Method) : Catch : relatedObj = Nothing : End Try
        Dim relCount As Integer = SafeCount(relatedObj)
        log("[DV_DIMLAB][DIM][RELATED] id=" & id & " count=" & relCount.ToString(CultureInfo.InvariantCulture) & " sig=NO_CONFIRMADO")

        Dim disp As Object = Nothing
        Dim dispLines As Integer = 0
        Dim dispArcs As Integer = 0
        Try : disp = CallByName(d, "GetDisplayData", CallType.Method) : Catch : disp = Nothing : End Try
        If disp IsNot Nothing Then
            dispLines = SafeCount(CallByNameSafe(disp, "Lines2d"))
            dispArcs = SafeCount(CallByNameSafe(disp, "Arcs2d"))
        End If
        log("[DV_DIMLAB][DIM][DISPLAY] id=" & id & " lines=" & dispLines.ToString(CultureInfo.InvariantCulture) & " arcs=" & dispArcs.ToString(CultureInfo.InvariantCulture))

        If relCount > 0 Then
            dimsConnected += 1
            If bestMethod = "none" Then bestMethod = "connected_" & tag
            log("[DV_DIMLAB][DIM][KEEP] id=" & id & " reason=connected")
        Else
            dimsFloatingVisible += 1
            If bestMethod = "none" Then bestMethod = "floating_visible_" & tag
            log("[DV_DIMLAB][DIM][KEEP] id=" & id & " reason=floating_visible")
        End If
    End Sub

    Private Shared Function ResolveStyleObject(draft As DraftDocument, target As String, log As Action(Of String)) As Object
        log("[DV_DIMLAB][STYLE][FIND] target=" & target)
        Dim styles As Object = CallByNameSafe(draft, "DimensionStyles")
        If styles Is Nothing Then
            log("[DV_DIMLAB][STYLE][FAIL] reason=not_found/type_mismatch")
            Return Nothing
        End If
        For i As Integer = 1 To SafeCount(styles)
            Dim st As Object = Nothing
            Try : st = CallByName(styles, "Item", CallType.Method, i) : Catch : st = Nothing : End Try
            If st Is Nothing Then Continue For
            Dim nm As String = SafeStr(CallByNameSafe(st, "Name"))
            log("[DV_DIMLAB][STYLE][CAND] name=" & nm)
            If String.Equals(nm, target, StringComparison.OrdinalIgnoreCase) Then
                log("[DV_DIMLAB][STYLE][OK] name=" & nm)
                Return st
            End If
        Next
        log("[DV_DIMLAB][STYLE][FAIL] reason=not_found/type_mismatch")
        Return Nothing
    End Function

    Private Shared Sub ApplyStyle(d As FrameworkDimension, styleObj As Object, log As Action(Of String))
        If d Is Nothing OrElse styleObj Is Nothing Then Return
        Try
            CallByName(d, "Style", CallType.Let, styleObj)
        Catch ex As Exception
            log("[DV_DIMLAB][STYLE][FAIL] reason=not_found/type_mismatch msg=" & ex.Message)
        End Try
    End Sub

    Private Shared Function ResolveRefOrGraphic(o As Object, log As Action(Of String), label As String) As Object
        If o Is Nothing Then Return Nothing
        Dim r As Object = Nothing
        Try : r = CallByName(o, "Reference", CallType.Get) : Catch : r = Nothing : End Try
        If r IsNot Nothing Then Return r
        Try : r = CallByName(o, "GraphicMember", CallType.Get) : Catch : r = Nothing : End Try
        If r IsNot Nothing Then Return r
        log("[DV_DIMLAB][DIM][TRY_B] no_ref_or_graphic label=" & label & " NO_CONFIRMADO")
        Return Nothing
    End Function

    Private Shared Sub AddPoint(out As List(Of InterestPoint), id As Integer, pointType As String, srcLine As DvLineInfo,
                                xView As Double, yView As Double, dv As DrawingView, log As Action(Of String))
        Dim xs As Double = xView, ys As Double = yView
        Try
            dv.ViewToSheet(xView, yView, xs, ys)
            log("[DV_DIMLAB][TRANSFORM][VIEW_TO_SHEET] view=" & SafeStr(CallByNameSafe(dv, "Name")) &
                " in=(" & F(xView) & "," & F(yView) & ") out=(" & F(xs) & "," & F(ys) & ")")
        Catch
            ' NO_CONFIRMADO: algunas APIs ya devuelven coordenada de hoja.
        End Try
        out.Add(New InterestPoint With {
            .Id = id,
            .PointType = pointType,
            .SourceType = If(srcLine Is Nothing, "bbox", "DVLine2d"),
            .SourceEntity = If(srcLine Is Nothing, "", srcLine.Id.ToString(CultureInfo.InvariantCulture)),
            .XView = xView, .YView = yView, .XSheet = xs, .YSheet = ys
        })
        log("[DV_DIMLAB][POINT] id=" & id.ToString(CultureInfo.InvariantCulture) & " type=" & pointType &
            " entity=" & If(srcLine Is Nothing, "bbox", "DVLine2d#" & srcLine.Id.ToString(CultureInfo.InvariantCulture)) &
            " view=(" & F(xView) & "," & F(yView) & ") sheet=(" & F(xs) & "," & F(ys) & ")")
    End Sub

    Private Shared Function DetectCoordinateMode(dv As DrawingView, ln As DvLineInfo) As String
        If dv Is Nothing OrElse ln Is Nothing Then Return "Unknown"
        Dim sx As Double = 0, sy As Double = 0
        Try
            dv.ViewToSheet(ln.X1, ln.Y1, sx, sy)
            Dim dx As Double = Math.Abs(sx - ln.X1) + Math.Abs(sy - ln.Y1)
            If dx > 0.000001R Then Return "ViewCoordinates"
            Return "SheetCoordinates"
        Catch
            Return "Unknown"
        End Try
    End Function

    Private Shared Sub ProbeProperty(obj As Object, id As Integer, name As String, log As Action(Of String))
        Try
            Dim v As Object = CallByName(obj, name, CallType.Get)
            log("[DV_DIMLAB][DVLINE][PROP] id=" & id.ToString(CultureInfo.InvariantCulture) & " name=" & name & " value=" & SafeStr(v))
        Catch
            log("[DV_DIMLAB][DVLINE][PROP_MISSING] id=" & id.ToString(CultureInfo.InvariantCulture) & " name=" & name)
        End Try
    End Sub

    Private Shared Sub ProbeMethod(obj As Object, id As Integer, methodName As String, log As Action(Of String))
        Dim exists As Boolean = False
        Try
            exists = obj.GetType().GetMethod(methodName) IsNot Nothing
        Catch
            exists = False
        End Try
        log("[DV_DIMLAB][DVLINE][METHOD_EXISTS] id=" & id.ToString(CultureInfo.InvariantCulture) &
            " name=" & methodName & " exists=" & exists.ToString(CultureInfo.InvariantCulture))
    End Sub

    Private Shared Function TryReadDvLineEndpoints(ln As Object, ByRef x1 As Double, ByRef y1 As Double, ByRef x2 As Double, ByRef y2 As Double) As Boolean
        ' Versión silenciosa para scoring de vistas.
        x1 = 0 : y1 = 0 : x2 = 0 : y2 = 0
        If ln Is Nothing Then Return False
        Try
            x1 = SafeD(CallByName(ln, "StartPointX", CallType.Get))
            y1 = SafeD(CallByName(ln, "StartPointY", CallType.Get))
            x2 = SafeD(CallByName(ln, "EndPointX", CallType.Get))
            y2 = SafeD(CallByName(ln, "EndPointY", CallType.Get))
            If Not IsZeroCoordSet(x1, y1, x2, y2) Then Return True
        Catch
        End Try
        Try
            Dim a1 As Object() = {0.0R, 0.0R}
            ln.GetType().InvokeMember("GetStartPoint", ReflectInvoke, Nothing, ln, a1)
            x1 = CDbl(a1(0)) : y1 = CDbl(a1(1))
            Dim a2 As Object() = {0.0R, 0.0R}
            ln.GetType().InvokeMember("GetEndPoint", ReflectInvoke, Nothing, ln, a2)
            x2 = CDbl(a2(0)) : y2 = CDbl(a2(1))
            Return Not IsZeroCoordSet(x1, y1, x2, y2)
        Catch
        End Try
        Return False
    End Function

    Private Shared Function TryReadDvLineEndpointsDetailed(ln As Object, id As Integer,
                                                           ByRef x1 As Double, ByRef y1 As Double, ByRef x2 As Double, ByRef y2 As Double,
                                                           log As Action(Of String)) As Boolean
        x1 = 0 : y1 = 0 : x2 = 0 : y2 = 0
        If ln Is Nothing Then Return False

        Dim source As String = ""
        If TryCoordsFromProps(ln, id, source, x1, y1, x2, y2, log, "StartPointX/StartPointY/EndPointX/EndPointY",
                              Function() Tuple.Create(SafeD(CallByName(ln, "StartPointX", CallType.Get)),
                                                      SafeD(CallByName(ln, "StartPointY", CallType.Get)),
                                                      SafeD(CallByName(ln, "EndPointX", CallType.Get)),
                                                      SafeD(CallByName(ln, "EndPointY", CallType.Get)))) Then Return True

        If TryCoordsFromProps(ln, id, source, x1, y1, x2, y2, log, "X1/Y1/X2/Y2",
                              Function() Tuple.Create(SafeD(CallByName(ln, "X1", CallType.Get)),
                                                      SafeD(CallByName(ln, "Y1", CallType.Get)),
                                                      SafeD(CallByName(ln, "X2", CallType.Get)),
                                                      SafeD(CallByName(ln, "Y2", CallType.Get)))) Then Return True

        log("[DV_DIMLAB][DVLINE][RAW_PROP_TRY] id=" & id.ToString(CultureInfo.InvariantCulture) & " name=GetStartPoint/GetEndPoint")
        Try
            Dim a1 As Object() = {0.0R, 0.0R}
            ln.GetType().InvokeMember("GetStartPoint", ReflectInvoke, Nothing, ln, a1)
            Dim a2 As Object() = {0.0R, 0.0R}
            ln.GetType().InvokeMember("GetEndPoint", ReflectInvoke, Nothing, ln, a2)
            x1 = SafeD(a1(0)) : y1 = SafeD(a1(1))
            x2 = SafeD(a2(0)) : y2 = SafeD(a2(1))
            If Not IsZeroCoordSet(x1, y1, x2, y2) Then
                log("[DV_DIMLAB][DVLINE][RAW_PROP_OK] id=" & id.ToString(CultureInfo.InvariantCulture) & " name=GetStartPoint/GetEndPoint")
                log("[DV_DIMLAB][DVLINE][COORD_SOURCE] id=" & id.ToString(CultureInfo.InvariantCulture) & " source=GetStartPoint/GetEndPoint")
                log("[DV_DIMLAB][DVLINE][VALID_COORDS] id=" & id.ToString(CultureInfo.InvariantCulture) &
                    " x1=" & F(x1) & " y1=" & F(y1) & " x2=" & F(x2) & " y2=" & F(y2))
                Return True
            End If
            log("[DV_DIMLAB][DVLINE][REJECT_ZERO_COORDS] id=" & id.ToString(CultureInfo.InvariantCulture) & " source=GetStartPoint/GetEndPoint")
        Catch ex As Exception
            log("[DV_DIMLAB][DVLINE][RAW_PROP_TRY] id=" & id.ToString(CultureInfo.InvariantCulture) & " name=GetStartPoint/GetEndPoint fail=" & ex.Message)
        End Try

        log("[DV_DIMLAB][DVLINE][RAW_PROP_TRY] id=" & id.ToString(CultureInfo.InvariantCulture) & " name=StartPoint/EndPoint")
        Try
            Dim sp As Object = CallByName(ln, "StartPoint", CallType.Get)
            Dim ep As Object = CallByName(ln, "EndPoint", CallType.Get)
            x1 = SafeD(CallByName(sp, "X", CallType.Get))
            y1 = SafeD(CallByName(sp, "Y", CallType.Get))
            x2 = SafeD(CallByName(ep, "X", CallType.Get))
            y2 = SafeD(CallByName(ep, "Y", CallType.Get))
            If Not IsZeroCoordSet(x1, y1, x2, y2) Then
                log("[DV_DIMLAB][DVLINE][RAW_PROP_OK] id=" & id.ToString(CultureInfo.InvariantCulture) & " name=StartPoint/EndPoint")
                log("[DV_DIMLAB][DVLINE][COORD_SOURCE] id=" & id.ToString(CultureInfo.InvariantCulture) & " source=StartPoint/EndPoint")
                log("[DV_DIMLAB][DVLINE][VALID_COORDS] id=" & id.ToString(CultureInfo.InvariantCulture) &
                    " x1=" & F(x1) & " y1=" & F(y1) & " x2=" & F(x2) & " y2=" & F(y2))
                Return True
            End If
            log("[DV_DIMLAB][DVLINE][REJECT_ZERO_COORDS] id=" & id.ToString(CultureInfo.InvariantCulture) & " source=StartPoint/EndPoint")
        Catch ex As Exception
            log("[DV_DIMLAB][DVLINE][RAW_PROP_TRY] id=" & id.ToString(CultureInfo.InvariantCulture) & " name=StartPoint/EndPoint fail=" & ex.Message)
        End Try

        log("[DV_DIMLAB][DVLINE][RAW_PROP_TRY] id=" & id.ToString(CultureInfo.InvariantCulture) & " name=GetEndPoints")
        Try
            Dim a As Object() = {0.0R, 0.0R, 0.0R, 0.0R}
            ln.GetType().InvokeMember("GetEndPoints", ReflectInvoke, Nothing, ln, a)
            x1 = SafeD(a(0)) : y1 = SafeD(a(1))
            x2 = SafeD(a(2)) : y2 = SafeD(a(3))
            If Not IsZeroCoordSet(x1, y1, x2, y2) Then
                log("[DV_DIMLAB][DVLINE][RAW_PROP_OK] id=" & id.ToString(CultureInfo.InvariantCulture) & " name=GetEndPoints")
                log("[DV_DIMLAB][DVLINE][COORD_SOURCE] id=" & id.ToString(CultureInfo.InvariantCulture) & " source=GetEndPoints")
                log("[DV_DIMLAB][DVLINE][VALID_COORDS] id=" & id.ToString(CultureInfo.InvariantCulture) &
                    " x1=" & F(x1) & " y1=" & F(y1) & " x2=" & F(x2) & " y2=" & F(y2))
                Return True
            End If
            log("[DV_DIMLAB][DVLINE][REJECT_ZERO_COORDS] id=" & id.ToString(CultureInfo.InvariantCulture) & " source=GetEndPoints")
        Catch ex As Exception
            log("[DV_DIMLAB][DVLINE][RAW_PROP_TRY] id=" & id.ToString(CultureInfo.InvariantCulture) & " name=GetEndPoints fail=" & ex.Message)
        End Try

        log("[DV_DIMLAB][DVLINE][RAW_PROP_TRY] id=" & id.ToString(CultureInfo.InvariantCulture) & " name=GetPointAtIndex")
        Try
            Dim p0 As Object() = {0, 0.0R, 0.0R}
            ln.GetType().InvokeMember("GetPointAtIndex", ReflectInvoke, Nothing, ln, p0)
            Dim p1 As Object() = {1, 0.0R, 0.0R}
            ln.GetType().InvokeMember("GetPointAtIndex", ReflectInvoke, Nothing, ln, p1)
            x1 = SafeD(p0(1)) : y1 = SafeD(p0(2))
            x2 = SafeD(p1(1)) : y2 = SafeD(p1(2))
            If Not IsZeroCoordSet(x1, y1, x2, y2) Then
                log("[DV_DIMLAB][DVLINE][RAW_PROP_OK] id=" & id.ToString(CultureInfo.InvariantCulture) & " name=GetPointAtIndex")
                log("[DV_DIMLAB][DVLINE][COORD_SOURCE] id=" & id.ToString(CultureInfo.InvariantCulture) & " source=GetPointAtIndex")
                log("[DV_DIMLAB][DVLINE][VALID_COORDS] id=" & id.ToString(CultureInfo.InvariantCulture) &
                    " x1=" & F(x1) & " y1=" & F(y1) & " x2=" & F(x2) & " y2=" & F(y2))
                Return True
            End If
            log("[DV_DIMLAB][DVLINE][REJECT_ZERO_COORDS] id=" & id.ToString(CultureInfo.InvariantCulture) & " source=GetPointAtIndex")
        Catch ex As Exception
            log("[DV_DIMLAB][DVLINE][RAW_PROP_TRY] id=" & id.ToString(CultureInfo.InvariantCulture) & " name=GetPointAtIndex fail=" & ex.Message)
        End Try

        ' Exploración defensiva adicional pedida por usuario.
        ProbeUnsupportedCoordSource(ln, id, "GetLineData", log)
        ProbeUnsupportedCoordSource(ln, id, "GetDisplayData", log)
        ProbeUnsupportedCoordSource(ln, id, "DisplayData", log)
        ProbeUnsupportedCoordSource(ln, id, "ModelMember", log)

        Return False
    End Function

    Private Shared Function TryCoordsFromProps(ln As Object, id As Integer, ByRef source As String,
                                               ByRef x1 As Double, ByRef y1 As Double, ByRef x2 As Double, ByRef y2 As Double,
                                               log As Action(Of String), label As String,
                                               getter As Func(Of Tuple(Of Double, Double, Double, Double))) As Boolean
        log("[DV_DIMLAB][DVLINE][RAW_PROP_TRY] id=" & id.ToString(CultureInfo.InvariantCulture) & " name=" & label)
        Try
            Dim t = getter.Invoke()
            x1 = t.Item1 : y1 = t.Item2 : x2 = t.Item3 : y2 = t.Item4
            If Not IsZeroCoordSet(x1, y1, x2, y2) Then
                source = label
                log("[DV_DIMLAB][DVLINE][RAW_PROP_OK] id=" & id.ToString(CultureInfo.InvariantCulture) & " name=" & label)
                log("[DV_DIMLAB][DVLINE][COORD_SOURCE] id=" & id.ToString(CultureInfo.InvariantCulture) & " source=" & label)
                log("[DV_DIMLAB][DVLINE][VALID_COORDS] id=" & id.ToString(CultureInfo.InvariantCulture) &
                    " x1=" & F(x1) & " y1=" & F(y1) & " x2=" & F(x2) & " y2=" & F(y2))
                Return True
            End If
            log("[DV_DIMLAB][DVLINE][REJECT_ZERO_COORDS] id=" & id.ToString(CultureInfo.InvariantCulture) & " source=" & label)
            Return False
        Catch ex As Exception
            log("[DV_DIMLAB][DVLINE][RAW_PROP_TRY] id=" & id.ToString(CultureInfo.InvariantCulture) & " name=" & label & " fail=" & ex.Message)
            Return False
        End Try
    End Function

    Private Shared Sub ProbeUnsupportedCoordSource(ln As Object, id As Integer, label As String, log As Action(Of String))
        log("[DV_DIMLAB][DVLINE][RAW_PROP_TRY] id=" & id.ToString(CultureInfo.InvariantCulture) & " name=" & label)
        Try
            Dim v As Object = Nothing
            If label = "DisplayData" OrElse label = "ModelMember" Then
                v = CallByName(ln, label, CallType.Get)
            Else
                v = ln.GetType().InvokeMember(label, ReflectInvoke, Nothing, ln, New Object() {})
            End If
            log("[DV_DIMLAB][DVLINE][RAW_PROP_OK] id=" & id.ToString(CultureInfo.InvariantCulture) & " name=" & label & " value=NO_CONFIRMADO")
        Catch ex As Exception
            log("[DV_DIMLAB][DVLINE][RAW_PROP_TRY] id=" & id.ToString(CultureInfo.InvariantCulture) & " name=" & label & " fail=" & ex.Message & " NO_CONFIRMADO")
        End Try
    End Sub

    Private Shared Function IsZeroCoordSet(x1 As Double, y1 As Double, x2 As Double, y2 As Double) As Boolean
        Return Math.Abs(x1) <= 1.0E-12R AndAlso Math.Abs(y1) <= 1.0E-12R AndAlso
               Math.Abs(x2) <= 1.0E-12R AndAlso Math.Abs(y2) <= 1.0E-12R
    End Function

    Private Shared Function IsLikelyIso(dv As DrawingView, typ As String) As Boolean
        Dim ori As String = SafeStr(CallByNameSafe(dv, "ViewOrientation"))
        Dim s As String = (ori & "|" & typ).ToLowerInvariant()
        Return s.Contains("iso")
    End Function

    Private Shared Function EnsureLayer(sh As Sheet, name As String) As Object
        If sh Is Nothing Then Return Nothing
        Dim layers As Object = CallByNameSafe(sh, "Layers")
        If layers Is Nothing Then Return Nothing
        For i As Integer = 1 To SafeCount(layers)
            Dim ly As Object = Nothing
            Try : ly = CallByName(layers, "Item", CallType.Method, i) : Catch : ly = Nothing : End Try
            If ly Is Nothing Then Continue For
            If String.Equals(SafeStr(CallByNameSafe(ly, "Name")), name, StringComparison.OrdinalIgnoreCase) Then Return ly
        Next
        Try : Return CallByName(layers, "Add", CallType.Method, name) : Catch : Return Nothing : End Try
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
            Return Convert.ToInt32(CallByName(col, "Count", CallType.Get), CultureInfo.InvariantCulture)
        Catch
            Return 0
        End Try
    End Function

    Private Shared Function SafeD(v As Object) As Double
        If v Is Nothing Then Return 0
        Try
            Return Convert.ToDouble(v, CultureInfo.InvariantCulture)
        Catch
            Return 0
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

    Private Shared Function F(v As Double) As String
        If Double.IsNaN(v) Then Return "NaN"
        Return v.ToString("0.######", CultureInfo.InvariantCulture)
    End Function
End Class

