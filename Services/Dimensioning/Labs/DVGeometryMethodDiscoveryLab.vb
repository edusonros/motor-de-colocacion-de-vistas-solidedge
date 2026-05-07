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
''' Laboratorio de descubrimiento sistemático de métodos SESDK sobre DVGeometry.
''' No modifica motor de vistas/placement/partslist/export.
''' </summary>
Friend NotInheritable Class DVGeometryMethodDiscoveryLab

    Private Const LogPrefix As String = "[DV_METHODLAB]"
    Private Const DimStyleName As String = "U3,5"
    Private Const AuxLayerName As String = "DV_METHODLAB_AUX"
    Private Const PlacementOffset As Double = 0.02R
    Private Const Tol As Double = 0.0002R

    Private Shared ReadOnly InvokeFlags As BindingFlags =
        BindingFlags.InvokeMethod Or BindingFlags.Public Or BindingFlags.Instance

    Private NotInheritable Class ViewInfo
        Public Dv As DrawingView
        Public Name As String
        Public IsIso As Boolean
        Public LineCount As Integer
        Public ArcCount As Integer
        Public CircleCount As Integer
        Public PointCount As Integer
        Public SplineCount As Integer
        Public HvScore As Integer
        Public MinX As Double
        Public MinY As Double
        Public MaxX As Double
        Public MaxY As Double
    End Class

    Private NotInheritable Class DvLineSample
        Public Id As Integer
        Public Obj As Object
        Public Length As Double
        Public X1 As Double
        Public Y1 As Double
        Public X2 As Double
        Public Y2 As Double
        Public HasCoords As Boolean
        Public CoordSource As String
    End Class

    Private NotInheritable Class Candidate
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

        L(LogPrefix & "[START]")
        If draft Is Nothing Then
            L(LogPrefix & "[ABORT] reason=no_draft")
            Return
        End If

        L(LogPrefix & "[DFT] path=" & SafeStr(CallByNameSafe(draft, "FullName")))
        L(LogPrefix & "[NO_DROP] true")
        L(LogPrefix & "[NO_2DMODEL] true")
        L(LogPrefix & "[VIEW_ENGINE_UNTOUCHED] true")

        Dim src As Sheet = ResolveSourceSheet(draft)
        If src Is Nothing Then
            L(LogPrefix & "[ABORT] reason=no_source_sheet")
            Return
        End If

        Dim views As List(Of ViewInfo) = CollectViews(src, L)
        L(LogPrefix & "[SOURCE_SHEET] name=" & SafeStr(CallByNameSafe(src, "Name")) & " views=" & views.Count.ToString(CultureInfo.InvariantCulture))
        If views.Count = 0 Then
            L(LogPrefix & "[ABORT] reason=no_views")
            Return
        End If

        Dim selected As ViewInfo = views.Where(Function(v) Not v.IsIso).OrderByDescending(Function(v) v.LineCount).ThenByDescending(Function(v) v.HvScore).FirstOrDefault()
        If selected Is Nothing Then selected = views.OrderByDescending(Function(v) v.LineCount).First()
        L(LogPrefix & "[VIEW][SELECTED] name=" & selected.Name & " reason=max_dvlines_non_iso")

        Dim samples As List(Of DvLineSample) = ProbeDvLines(selected, L)
        If samples.Where(Function(s) s.HasCoords).Count() = 0 Then
            For Each alt In views
                If Object.ReferenceEquals(alt, selected) Then Continue For
                L(LogPrefix & "[VIEW][FALLBACK_PROBE] name=" & alt.Name & " reason=selected_has_no_valid_coords")
                Dim altSamples As List(Of DvLineSample) = ProbeDvLines(alt, L)
                If altSamples.Where(Function(s) s.HasCoords).Count() > 0 Then
                    selected = alt
                    samples = altSamples
                    L(LogPrefix & "[VIEW][SELECTED] name=" & selected.Name & " reason=fallback_first_with_valid_coords")
                    Exit For
                End If
            Next
        End If
        Dim rangeOkCount As Integer = ProbeLineRanges(samples, L)
        Dim keypointOkCount As Integer = ProbeLineKeypoints(samples, L)
        Dim referenceOkCount As Integer = ProbeLineReferences(samples, L)

        Dim gmCount As Integer = ProbeGraphicMembers(selected, L)
        Dim getRefToGmOk As Integer = ProbeGetReferenceToGraphicMember(selected, samples, L)

        Dim arcCenterOk As Integer = ProbeArcs(selected, L)
        Dim circleCenterOk As Integer = ProbeCircles(selected, L)
        Dim pointXyOk As Integer = ProbeDvPoints(selected, L)

        ProbeTransforms(selected, samples, L)

        Dim cH As Integer = 0, cV As Integer = 0, cP As Integer = 0
        Dim candidates As List(Of Candidate) = BuildCandidates(selected, samples, L, cH, cV, cP)

        Dim dims As Dimensions = Nothing
        Try : dims = CType(CallByName(src, "Dimensions", CallType.Get), Dimensions) : Catch : dims = Nothing : End Try
        If dims Is Nothing Then
            L(LogPrefix & "[ABORT] reason=no_dimensions_collection")
            Return
        End If

        Dim styleObj As Object = ResolveStyleObject(draft, L)
        TrySetCollectionStyle(dims, styleObj, L)
        Dim auxLayer As Object = EnsureLayer(src, AuxLayerName)

        ' FASE EXTRA: ingeniería inversa de dimensiones existentes válidas antes de crear nuevas.
        ProbeExistingDimensions(src, selected, L)

        Dim adboDvOk As Integer = 0
        Dim adboExDvOk As Integer = 0
        Dim adboRefOk As Integer = 0
        Dim adboExRefOk As Integer = 0
        Dim dimInitOk As Integer = 0
        Dim auxOk As Integer = 0
        Dim dimsRetrievedOk As Integer = 0
        Dim dimsConnected As Integer = 0
        Dim dimsFloating As Integer = 0
        Dim bestCoordMethod As String = If(samples.Any(Function(s) s.HasCoords), samples.First(Function(s) s.HasCoords).CoordSource, "none")
        Dim bestDimMethod As String = "none"

        For Each c In candidates
            If c.Expected <= 0 Then Continue For
            ' A: AddDistanceBetweenObjects DVLine2d
            L(LogPrefix & "[DIM][TRY] method=ADBO objType=DVLine2d source=" & c.Kind)
            Dim dA As FrameworkDimension = Nothing
            If TryAddDistanceGeneric(dims, c.O1, c.O2, c.X1, c.Y1, c.X2, c.Y2, If(c.Kind.StartsWith("H"), "horizontal", "vertical"), selected.Dv, dA) Then
                If AcceptDimension(dA, c.Expected, L, "ADBO_DVLINE", dimsConnected, dimsFloating) Then
                    adboDvOk += 1
                    If bestDimMethod = "none" Then bestDimMethod = "ADBO_DVLINE"
                    ApplyStyle(dA, styleObj, L)
                End If
            End If

            ' B: AddDistanceBetweenObjectsEX DVLine2d (firma no confirmada)
            L(LogPrefix & "[DIM][TRY] method=ADBOEX objType=DVLine2d source=" & c.Kind)
            Dim dB As FrameworkDimension = Nothing
            If TryAddDistanceEx(dims, c.O1, c.O2, c.X1, c.Y1, c.X2, c.Y2, dB, L) Then
                If AcceptDimension(dB, c.Expected, L, "ADBOEX_DVLINE", dimsConnected, dimsFloating) Then
                    adboExDvOk += 1
                    If bestDimMethod = "none" Then bestDimMethod = "ADBOEX_DVLINE"
                    ApplyStyle(dB, styleObj, L)
                End If
            End If

            ' C: AddDistanceBetweenObjects con Reference
            Dim r1 As Object = SafeReference(c.O1)
            Dim r2 As Object = SafeReference(c.O2)
            If r1 IsNot Nothing AndAlso r2 IsNot Nothing Then
                L(LogPrefix & "[DIM][TRY] method=ADBO objType=Reference source=" & c.Kind)
                Dim dC As FrameworkDimension = Nothing
                If TryAddDistanceGeneric(dims, r1, r2, c.X1, c.Y1, c.X2, c.Y2, If(c.Kind.StartsWith("H"), "horizontal", "vertical"), selected.Dv, dC) Then
                    If AcceptDimension(dC, c.Expected, L, "ADBO_REFERENCE", dimsConnected, dimsFloating) Then
                        adboRefOk += 1
                        If bestDimMethod = "none" Then bestDimMethod = "ADBO_REFERENCE"
                        ApplyStyle(dC, styleObj, L)
                    End If
                End If

                ' D: AddDistanceBetweenObjectsEX con Reference
                L(LogPrefix & "[DIM][TRY] method=ADBOEX objType=Reference source=" & c.Kind)
                Dim dD As FrameworkDimension = Nothing
                If TryAddDistanceEx(dims, r1, r2, c.X1, c.Y1, c.X2, c.Y2, dD, L) Then
                    If AcceptDimension(dD, c.Expected, L, "ADBOEX_REFERENCE", dimsConnected, dimsFloating) Then
                        adboExRefOk += 1
                        If bestDimMethod = "none" Then bestDimMethod = "ADBOEX_REFERENCE"
                        ApplyStyle(dD, styleObj, L)
                    End If
                End If
            End If

            ' F: AddDimension con DimInitData (NO_CONFIRMADO firma)
            L(LogPrefix & "[DIM][TRY] method=DimInitData source=" & c.Kind)
            Dim dF As FrameworkDimension = Nothing
            If TryAddByDimInitData(dims, c, dF, L) Then
                If AcceptDimension(dF, c.Expected, L, "DIMINITDATA", dimsConnected, dimsFloating) Then
                    dimInitOk += 1
                    If bestDimMethod = "none" Then bestDimMethod = "DIMINITDATA"
                    ApplyStyle(dF, styleObj, L)
                End If
            End If
        Next

        ' G: AddLength auxiliar desde bbox seleccionado (fallback)
        L(LogPrefix & "[DIM][TRY] method=AddLength_AuxLine source=bbox")
        Dim dAux As FrameworkDimension = Nothing
        If TryAddAuxLength(src, dims, selected, auxLayer, dAux, L) Then
            If AcceptDimension(dAux, -1, L, "ADDLENGTH_AUX", dimsConnected, dimsFloating) Then
                auxOk += 1
                If bestDimMethod = "none" Then bestDimMethod = "ADDLENGTH_AUX"
                ApplyStyle(dAux, styleObj, L)
            End If
        End If

        ' Retrieve/connected dimensions
        If TryRetrieveDimensions(selected.Dv, src, L) Then dimsRetrievedOk += 1

        L(LogPrefix & "[SUMMARY]")
        L("views_processed=" & views.Count.ToString(CultureInfo.InvariantCulture))
        L("selected_view=" & selected.Name)
        L("dvlines_total=" & selected.LineCount.ToString(CultureInfo.InvariantCulture))
        L("dvlines_length_ok=" & samples.Where(Function(s) s.Length > 0).Count().ToString(CultureInfo.InvariantCulture))
        L("dvlines_start_end_ok=" & samples.Where(Function(s) s.HasCoords).Count().ToString(CultureInfo.InvariantCulture))
        L("dvlines_range_ok=" & rangeOkCount.ToString(CultureInfo.InvariantCulture))
        L("dvlines_keypoint_ok=" & keypointOkCount.ToString(CultureInfo.InvariantCulture))
        L("dvlines_reference_ok=" & referenceOkCount.ToString(CultureInfo.InvariantCulture))
        L("dvlines_referencekey_ok=" & samples.Count.ToString(CultureInfo.InvariantCulture)) ' se intentó en todos
        L("graphicmembers_count=" & gmCount.ToString(CultureInfo.InvariantCulture))
        L("get_ref_to_gm_ok=" & getRefToGmOk.ToString(CultureInfo.InvariantCulture))
        L("dvarcs_center_ok=" & arcCenterOk.ToString(CultureInfo.InvariantCulture))
        L("dvcircles_center_ok=" & circleCenterOk.ToString(CultureInfo.InvariantCulture))
        L("dvpoints_xy_ok=" & pointXyOk.ToString(CultureInfo.InvariantCulture))
        L("candidates_h_total=" & cH.ToString(CultureInfo.InvariantCulture))
        L("candidates_v_total=" & cV.ToString(CultureInfo.InvariantCulture))
        L("dims_adbo_dvline_ok=" & adboDvOk.ToString(CultureInfo.InvariantCulture))
        L("dims_adboex_dvline_ok=" & adboExDvOk.ToString(CultureInfo.InvariantCulture))
        L("dims_adbo_reference_ok=" & adboRefOk.ToString(CultureInfo.InvariantCulture))
        L("dims_adboex_reference_ok=" & adboExRefOk.ToString(CultureInfo.InvariantCulture))
        L("dims_diminitdata_ok=" & dimInitOk.ToString(CultureInfo.InvariantCulture))
        L("dims_addlength_aux_ok=" & auxOk.ToString(CultureInfo.InvariantCulture))
        L("dims_retrieved_ok=" & dimsRetrievedOk.ToString(CultureInfo.InvariantCulture))
        L("dims_connected=" & dimsConnected.ToString(CultureInfo.InvariantCulture))
        L("dims_floating=" & dimsFloating.ToString(CultureInfo.InvariantCulture))
        L("best_coordinate_method=" & bestCoordMethod)
        L("best_dimension_method=" & bestDimMethod)
        L("recommended_next_step=promocionar_metodo_con_mas_OK_y_menos_FAIL")

        If debugSave Then
            Try : draft.Save() : Catch : End Try
        End If
    End Sub

    Private Shared Sub ProbeExistingDimensions(sh As Sheet, selected As ViewInfo, log As Action(Of String))
        If sh Is Nothing Then Return
        Dim dims As Object = CallByNameSafe(sh, "Dimensions")
        Dim n As Integer = SafeCount(dims)
        For i As Integer = 1 To n
            Dim d As Object = Nothing
            Try : d = CallByName(dims, "Item", CallType.Method, i) : Catch : d = Nothing : End Try
            If d Is Nothing Then Continue For

            Dim dimType As String = SafeStr(CallByNameSafe(d, "Type"))
            Dim val As Double = SafeD(CallByNameSafe(d, "Value"))
            Dim sty As String = SafeStr(CallByNameSafe(CallByNameSafe(d, "Style"), "Name"))
            Dim parentType As String = TypeNameSafe(CallByNameSafe(d, "Parent"))
            Dim status As String = SafeStr(CallByNameSafe(d, "StatusOfDimension"))
            Dim trackDist As String = SafeStr(CallByNameSafe(d, "TrackDistance"))
            Dim trackAng As String = SafeStr(CallByNameSafe(d, "TrackAngle"))
            Dim layerName As String = SafeStr(CallByNameSafe(CallByNameSafe(d, "Layer"), "Name"))
            Dim color As String = SafeStr(CallByNameSafe(d, "Color"))
            Dim prefix As String = SafeStr(CallByNameSafe(d, "Prefix"))
            Dim suffix As String = SafeStr(CallByNameSafe(d, "Suffix"))
            Dim axisEx As String = SafeStr(CallByNameSafe(d, "MeasurementAxisEx"))
            Dim orient As String = SafeStr(CallByNameSafe(d, "Orientation"))
            Dim term As String = SafeStr(CallByNameSafe(d, "TerminatorType"))

            log(LogPrefix & "[EXISTING_DIM][INFO] idx=" & i.ToString(CultureInfo.InvariantCulture) &
                " type=" & dimType &
                " value=" & F(val) &
                " style=" & sty &
                " parent=" & parentType &
                " status=" & status &
                " trackDistance=" & trackDist &
                " trackAngle=" & trackAng &
                " layer=" & layerName &
                " color=" & color &
                " prefix=" & prefix &
                " suffix=" & suffix &
                " measurementAxisEx=" & axisEx &
                " orientation=" & orient &
                " terminatorType=" & term)

            ' Related / parent objects / relationships
            Dim relCount As Integer = 0
            Dim parentObjCount As Integer = 0
            Dim relationshipsCount As Integer = 0
            Try
                Dim rel As Object = CallByName(d, "GetRelatedObjects", CallType.Method)
                relCount = SafeCount(rel)
            Catch
            End Try
            parentObjCount = SafeCount(CallByNameSafe(d, "ParentObjects"))
            relationshipsCount = SafeCount(CallByNameSafe(d, "Relationships"))
            log(LogPrefix & "[EXISTING_DIM][RELATED] idx=" & i.ToString(CultureInfo.InvariantCulture) &
                " getRelatedCount=" & relCount.ToString(CultureInfo.InvariantCulture) &
                " parentObjectsCount=" & parentObjCount.ToString(CultureInfo.InvariantCulture) &
                " relationshipsCount=" & relationshipsCount.ToString(CultureInfo.InvariantCulture))

            ' Display data
            Dim dd As Object = Nothing
            Dim dLines As Integer = 0, dArcs As Integer = 0, dPoints As Integer = 0, dCircles As Integer = 0
            Try
                dd = CallByName(d, "GetDisplayData", CallType.Method)
                dLines = SafeCount(CallByNameSafe(dd, "Lines2d"))
                dArcs = SafeCount(CallByNameSafe(dd, "Arcs2d"))
                dPoints = SafeCount(CallByNameSafe(dd, "Points2d"))
                dCircles = SafeCount(CallByNameSafe(dd, "Circles2d"))
            Catch
            End Try
            log(LogPrefix & "[EXISTING_DIM][DISPLAYDATA] idx=" & i.ToString(CultureInfo.InvariantCulture) &
                " lines=" & dLines.ToString(CultureInfo.InvariantCulture) &
                " arcs=" & dArcs.ToString(CultureInfo.InvariantCulture) &
                " points=" & dPoints.ToString(CultureInfo.InvariantCulture) &
                " circles=" & dCircles.ToString(CultureInfo.InvariantCulture))

            ' Reference-like probes: GetReferenceKey / GetKeyPoint
            Dim refBytes As Integer = 0
            Dim keypointHits As Integer = 0
            Try
                Dim keyObj As Object = Nothing
                CallByName(d, "GetReferenceKey", CallType.Method, keyObj)
                refBytes = ByteLength(keyObj)
            Catch
            End Try
            For kp As Integer = 0 To 2
                Try
                    Dim x As Double = 0, y As Double = 0
                    CallByName(d, "GetKeyPoint", CallType.Method, kp, x, y)
                    keypointHits += 1
                Catch
                End Try
            Next
            log(LogPrefix & "[EXISTING_DIM][REFERENCE] idx=" & i.ToString(CultureInfo.InvariantCulture) &
                " referenceKeyBytes=" & refBytes.ToString(CultureInfo.InvariantCulture) &
                " getKeyPointHits=" & keypointHits.ToString(CultureInfo.InvariantCulture))

            ' Intento de vínculo con vista y miembros gráficos
            Dim viewName As String = ""
            Dim via As String = "none"
            Dim gmHits As Integer = 0

            Dim pObj As Object = CallByNameSafe(d, "Parent")
            If pObj IsNot Nothing Then
                Dim pType As String = TypeNameSafe(pObj)
                If pType.IndexOf("DrawingView", StringComparison.OrdinalIgnoreCase) >= 0 Then
                    viewName = SafeStr(CallByNameSafe(pObj, "Name"))
                    via = "parent"
                End If
            End If

            Dim relObjs As Object = Nothing
            Try : relObjs = CallByName(d, "GetRelatedObjects", CallType.Method) : Catch : relObjs = Nothing : End Try
            If relObjs IsNot Nothing Then
                For r As Integer = 1 To SafeCount(relObjs)
                    Dim o As Object = Nothing
                    Try : o = CallByName(relObjs, "Item", CallType.Method, r) : Catch : o = Nothing : End Try
                    If o Is Nothing Then Continue For
                    Dim t As String = TypeNameSafe(o)
                    If t.IndexOf("GraphicMember", StringComparison.OrdinalIgnoreCase) >= 0 Then gmHits += 1
                    If String.IsNullOrWhiteSpace(viewName) Then
                        Dim dv As Object = CallByNameSafe(o, "DrawingView")
                        If dv IsNot Nothing Then
                            viewName = SafeStr(CallByNameSafe(dv, "Name"))
                            via = "related.DrawingView"
                        End If
                    End If
                Next
            End If

            If String.IsNullOrWhiteSpace(viewName) AndAlso selected IsNot Nothing Then
                viewName = selected.Name
                via = "fallback_selected_view"
            End If

            log(LogPrefix & "[EXISTING_DIM][VIEW_LINK] idx=" & i.ToString(CultureInfo.InvariantCulture) &
                " view=" & viewName &
                " via=" & via &
                " graphicMembersRelated=" & gmHits.ToString(CultureInfo.InvariantCulture) &
                " floating=" & ((relCount = 0 AndAlso parentObjCount = 0).ToString(CultureInfo.InvariantCulture)))
        Next
    End Sub

    Private Shared Function ResolveSourceSheet(draft As DraftDocument) As Sheet
        Dim n As Integer = SafeCount(CallByNameSafe(draft, "Sheets"))
        Dim hoja1 As Sheet = Nothing
        Dim best As Sheet = Nothing
        Dim bestViews As Integer = -1
        For i As Integer = 1 To n
            Dim sh As Sheet = Nothing
            Try : sh = CType(CallByName(draft.Sheets, "Item", CallType.Method, i), Sheet) : Catch : sh = Nothing : End Try
            If sh Is Nothing Then Continue For
            Dim vc As Integer = SafeCount(CallByNameSafe(sh, "DrawingViews"))
            If String.Equals(SafeStr(CallByNameSafe(sh, "Name")), "Hoja1", StringComparison.OrdinalIgnoreCase) AndAlso vc > 0 Then hoja1 = sh
            If vc > bestViews Then
                bestViews = vc
                best = sh
            End If
        Next
        If hoja1 IsNot Nothing Then Return hoja1
        Return best
    End Function

    Private Shared Function CollectViews(sh As Sheet, log As Action(Of String)) As List(Of ViewInfo)
        Dim out As New List(Of ViewInfo)()
        Dim dvs As Object = CallByNameSafe(sh, "DrawingViews")
        For i As Integer = 1 To SafeCount(dvs)
            Dim dv As DrawingView = Nothing
            Try : dv = CType(CallByName(dvs, "Item", CallType.Method, i), DrawingView) : Catch : dv = Nothing : End Try
            If dv Is Nothing Then Continue For

            Dim vi As New ViewInfo With {.Dv = dv, .Name = SafeStr(CallByNameSafe(dv, "Name"))}
            Dim typ As String = SafeStr(CallByNameSafe(dv, "Type"))
            Dim viewType As String = SafeStr(CallByNameSafe(dv, "ViewType"))
            Dim drawingViewType As String = SafeStr(CallByNameSafe(dv, "DrawingViewType"))
            Dim scale As String = SafeStr(CallByNameSafe(dv, "ScaleFactor"))
            If String.IsNullOrWhiteSpace(scale) Then scale = SafeStr(CallByNameSafe(dv, "Scale"))

            Dim minX As Double = 0, minY As Double = 0, maxX As Double = 0, maxY As Double = 0
            Try : dv.Range(minX, minY, maxX, maxY) : Catch : End Try
            vi.MinX = Math.Min(minX, maxX)
            vi.MinY = Math.Min(minY, maxY)
            vi.MaxX = Math.Max(minX, maxX)
            vi.MaxY = Math.Max(minY, maxY)

            Dim originX As Double = 0, originY As Double = 0
            Dim originLog As String = ""
            Try
                Dim a As Object() = {0.0R, 0.0R}
                dv.GetType().InvokeMember("GetOrigin", InvokeFlags, Nothing, dv, a)
                originX = SafeD(a(0))
                originY = SafeD(a(1))
                originLog = " origin=(" & F(originX) & "," & F(originY) & ")"
            Catch
                originLog = " origin=NO_CONFIRMADO"
            End Try

            vi.LineCount = SafeCount(CallByNameSafe(dv, "DVLines2d"))
            vi.ArcCount = SafeCount(CallByNameSafe(dv, "DVArcs2d"))
            vi.CircleCount = SafeCount(CallByNameSafe(dv, "DVCircles2d"))
            vi.PointCount = SafeCount(CallByNameSafe(dv, "DVPoints2d"))
            vi.SplineCount = SafeCount(CallByNameSafe(dv, "DVBSplineCurves2d"))
            If vi.SplineCount = 0 Then vi.SplineCount = SafeCount(CallByNameSafe(dv, "DVBSplines2d"))
            Dim modelMembers As Integer = SafeCount(CallByNameSafe(dv, "ModelMembers"))
            Dim graphicMembers As Integer = SafeCount(CallByNameSafe(dv, "GraphicMembers"))
            vi.IsIso = (typ & "|" & viewType & "|" & drawingViewType).ToLowerInvariant().Contains("iso")

            vi.HvScore = CountHvLines(dv)

            log(LogPrefix & "[VIEW][INFO] idx=" & i.ToString(CultureInfo.InvariantCulture) & " name=" & vi.Name &
                " type=" & typ & " viewType=" & If(String.IsNullOrWhiteSpace(viewType), drawingViewType, viewType) & " scale=" & scale)
            log(LogPrefix & "[VIEW][RANGE] xmin=" & F(vi.MinX) & " ymin=" & F(vi.MinY) & " xmax=" & F(vi.MaxX) & " ymax=" & F(vi.MaxY) & originLog)
            log(LogPrefix & "[VIEW][COUNTS] lines=" & vi.LineCount.ToString(CultureInfo.InvariantCulture) &
                " arcs=" & vi.ArcCount.ToString(CultureInfo.InvariantCulture) &
                " circles=" & vi.CircleCount.ToString(CultureInfo.InvariantCulture) &
                " points=" & vi.PointCount.ToString(CultureInfo.InvariantCulture) &
                " splines=" & vi.SplineCount.ToString(CultureInfo.InvariantCulture) &
                " modelMembers=" & modelMembers.ToString(CultureInfo.InvariantCulture) &
                " graphicMembers=" & graphicMembers.ToString(CultureInfo.InvariantCulture))
            out.Add(vi)
        Next
        Return out
    End Function

    Private Shared Function CountHvLines(dv As DrawingView) As Integer
        Dim n As Integer = 0
        Dim i As Integer = 0
        For Each ln In EnumerateDvLines(dv)
            i += 1
            If ln Is Nothing Then Continue For
            Dim x1 As Double = 0, y1 As Double = 0, x2 As Double = 0, y2 As Double = 0
            If TryGetStartEnd(ln, x1, y1, x2, y2, Nothing, i, False) Then
                If Math.Abs(x2 - x1) <= Tol OrElse Math.Abs(y2 - y1) <= Tol Then n += 1
            End If
        Next
        Return n
    End Function

    Private Shared Function ProbeDvLines(v As ViewInfo, log As Action(Of String)) As List(Of DvLineSample)
        Dim out As New List(Of DvLineSample)()
        Dim i As Integer = 0
        For Each ln In EnumerateDvLines(v.Dv)
            i += 1
            If ln Is Nothing Then Continue For
            Dim s As New DvLineSample With {.Id = i, .Obj = ln}

            s.Length = SafeD(CallByNameSafe(ln, "Length"))
            Dim ang As Double = SafeD(CallByNameSafe(ln, "Angle"))
            Dim kpc As Integer = CInt(SafeD(CallByNameSafe(ln, "KeyPointCount")))
            Dim edgeType As String = SafeStr(CallByNameSafe(ln, "EdgeType"))
            Dim typ As String = SafeStr(CallByNameSafe(ln, "Type"))
            log(LogPrefix & "[DVLINE][BASIC] id=" & i.ToString(CultureInfo.InvariantCulture) &
                " length=" & F(s.Length) & " angle=" & F(ang) & " keyPointCount=" & kpc.ToString(CultureInfo.InvariantCulture) &
                " edgeType=" & edgeType & " type=" & typ)

            Dim r As Object = CallByNameSafe(ln, "Reference")
            log(LogPrefix & "[DVLINE][REF_PROP] id=" & i.ToString(CultureInfo.InvariantCulture) &
                " reference_ok=" & (r IsNot Nothing).ToString(CultureInfo.InvariantCulture) &
                " reference_type=" & If(r Is Nothing, "", r.GetType().FullName))
            Dim mm As Object = CallByNameSafe(ln, "ModelMember")
            log(LogPrefix & "[DVLINE][MODELMEMBER] id=" & i.ToString(CultureInfo.InvariantCulture) &
                " ok=" & (mm IsNot Nothing).ToString(CultureInfo.InvariantCulture) &
                " type=" & If(mm Is Nothing, "", mm.GetType().FullName))
            Dim rel As Object = CallByNameSafe(ln, "Relationships")
            log(LogPrefix & "[DVLINE][RELATIONSHIPS] id=" & i.ToString(CultureInfo.InvariantCulture) &
                " ok=" & (rel IsNot Nothing).ToString(CultureInfo.InvariantCulture) &
                " count=" & SafeCount(rel).ToString(CultureInfo.InvariantCulture))

            Dim okCoords As Boolean = TryGetStartEnd(ln, s.X1, s.Y1, s.X2, s.Y2, log, i, True)
            If okCoords Then
                s.HasCoords = True
                s.CoordSource = "GetStartPoint/GetEndPoint"
            Else
                s.CoordSource = "none"
            End If
            out.Add(s)
        Next
        Return out
    End Function

    Private Shared Function EnumerateDvLines(dv As DrawingView) As List(Of Object)
        Dim result As New List(Of Object)()
        If dv Is Nothing Then Return result
        Dim col As Object = CallByNameSafe(dv, "DVLines2d")
        If col Is Nothing Then Return result

        Dim count As Integer = 0
        Try
            ' API oficial de colección DVLines2d: estabiliza snapshot de ítems para ItemEx.
            count = Convert.ToInt32(CallByName(col, "ResetCollectionAndGetCount", CallType.Method), CultureInfo.InvariantCulture)
        Catch
            count = SafeCount(col)
        End Try

        If count <= 0 Then Return result
        For i As Integer = 1 To count
            Dim ln As Object = Nothing
            Try : ln = CallByName(col, "ItemEx", CallType.Method, i) : Catch : ln = Nothing : End Try
            If ln Is Nothing Then
                Try : ln = CallByName(col, "Item", CallType.Method, i) : Catch : ln = Nothing : End Try
            End If
            If ln IsNot Nothing Then result.Add(ln)
        Next
        Return result
    End Function

    Private Shared Function TryGetStartEnd(ln As Object, ByRef x1 As Double, ByRef y1 As Double, ByRef x2 As Double, ByRef y2 As Double,
                                           log As Action(Of String), lineId As Integer, writeLogs As Boolean) As Boolean
        x1 = 0 : y1 = 0 : x2 = 0 : y2 = 0

        If writeLogs Then log(LogPrefix & "[DVLINE][GET_START_END][TRY] id=" & lineId.ToString(CultureInfo.InvariantCulture) & " signature=R_direct")
        Try
            Dim sx As Double = 0.0R, sy As Double = 0.0R
            Dim ex As Double = 0.0R, ey As Double = 0.0R
            ' Mismo patrón que ya devuelve coordenadas reales en DimensionKeypointRelinkLabV2.
            ln.GetStartPoint(sx, sy)
            ln.GetEndPoint(ex, ey)
            x1 = sx : y1 = sy : x2 = ex : y2 = ey
            If ValidateCoords(ln, x1, y1, x2, y2, lineId, "R_direct", log, writeLogs) Then Return True
        Catch ex As Exception
            If writeLogs Then log(LogPrefix & "[DVLINE][GET_START_END][FAIL] id=" & lineId.ToString(CultureInfo.InvariantCulture) & " signature=R_direct error=" & ex.Message)
        End Try

        If writeLogs Then log(LogPrefix & "[DVLINE][GET_START_END][TRY] id=" & lineId.ToString(CultureInfo.InvariantCulture) & " signature=A")
        Try
            Dim sx As Double = 0, sy As Double = 0, ex As Double = 0, ey As Double = 0
            CallByName(ln, "GetStartPoint", CallType.Method, sx, sy)
            CallByName(ln, "GetEndPoint", CallType.Method, ex, ey)
            x1 = sx : y1 = sy : x2 = ex : y2 = ey
            If ValidateCoords(ln, x1, y1, x2, y2, lineId, "A", log, writeLogs) Then Return True
        Catch ex As Exception
            If writeLogs Then log(LogPrefix & "[DVLINE][GET_START_END][FAIL] id=" & lineId.ToString(CultureInfo.InvariantCulture) & " signature=A error=" & ex.Message)
        End Try

        If writeLogs Then log(LogPrefix & "[DVLINE][GET_START_END][TRY] id=" & lineId.ToString(CultureInfo.InvariantCulture) & " signature=B")
        Try
            Dim ox As Object = 0.0R, oy As Object = 0.0R, exo As Object = 0.0R, eyo As Object = 0.0R
            CallByName(ln, "GetStartPoint", CallType.Method, ox, oy)
            CallByName(ln, "GetEndPoint", CallType.Method, exo, eyo)
            x1 = SafeD(ox) : y1 = SafeD(oy) : x2 = SafeD(exo) : y2 = SafeD(eyo)
            If ValidateCoords(ln, x1, y1, x2, y2, lineId, "B", log, writeLogs) Then Return True
        Catch ex As Exception
            If writeLogs Then log(LogPrefix & "[DVLINE][GET_START_END][FAIL] id=" & lineId.ToString(CultureInfo.InvariantCulture) & " signature=B error=" & ex.Message)
        End Try

        If writeLogs Then log(LogPrefix & "[DVLINE][GET_START_END][TRY] id=" & lineId.ToString(CultureInfo.InvariantCulture) & " signature=E")
        Try
            Dim a1 As Object() = {0.0R, 0.0R}
            ln.GetType().InvokeMember("GetStartPoint", InvokeFlags, Nothing, ln, a1)
            Dim a2 As Object() = {0.0R, 0.0R}
            ln.GetType().InvokeMember("GetEndPoint", InvokeFlags, Nothing, ln, a2)
            x1 = SafeD(a1(0)) : y1 = SafeD(a1(1)) : x2 = SafeD(a2(0)) : y2 = SafeD(a2(1))
            If ValidateCoords(ln, x1, y1, x2, y2, lineId, "E", log, writeLogs) Then Return True
        Catch ex As Exception
            If writeLogs Then log(LogPrefix & "[DVLINE][GET_START_END][FAIL] id=" & lineId.ToString(CultureInfo.InvariantCulture) & " signature=E error=" & ex.Message)
        End Try

        Return False
    End Function

    Private Shared Function ValidateCoords(ln As Object, x1 As Double, y1 As Double, x2 As Double, y2 As Double, id As Integer, sig As String, log As Action(Of String), writeLogs As Boolean) As Boolean
        Dim len As Double = SafeD(CallByNameSafe(ln, "Length"))
        Dim allZero As Boolean = IsZero(x1) AndAlso IsZero(y1) AndAlso IsZero(x2) AndAlso IsZero(y2)
        If allZero AndAlso len > 0 Then
            If writeLogs Then
                log(LogPrefix & "[DVLINE][GET_START_END][REJECT_ZERO] id=" & id.ToString(CultureInfo.InvariantCulture) & " signature=" & sig & " reason=both_points_zero_and_length_gt_zero")
                log(LogPrefix & "[DVLINE][COORD_CONTRADICTION] id=" & id.ToString(CultureInfo.InvariantCulture) & " length=" & F(len) & " start=(0,0) end=(0,0)")
            End If
            Return False
        End If
        If writeLogs Then
            log(LogPrefix & "[DVLINE][GET_START_END][OK] id=" & id.ToString(CultureInfo.InvariantCulture) &
                " signature=" & sig & " start=(" & F(x1) & "," & F(y1) & ") end=(" & F(x2) & "," & F(y2) & ")")
        End If
        Return True
    End Function

    Private Shared Function ProbeLineRanges(samples As List(Of DvLineSample), log As Action(Of String)) As Integer
        Dim okCount As Integer = 0
        For Each s In samples
            log(LogPrefix & "[DVLINE][RANGE][TRY] id=" & s.Id.ToString(CultureInfo.InvariantCulture) & " signature=A")
            Dim x1 As Double = 0, y1 As Double = 0, x2 As Double = 0, y2 As Double = 0
            Dim ok As Boolean = False
            Try
                CallByName(s.Obj, "Range", CallType.Method, x1, y1, x2, y2)
                ok = True
            Catch ex As Exception
                log(LogPrefix & "[DVLINE][RANGE][FAIL] id=" & s.Id.ToString(CultureInfo.InvariantCulture) & " signature=A error=" & ex.Message)
            End Try
            If Not ok Then
                log(LogPrefix & "[DVLINE][RANGE][TRY] id=" & s.Id.ToString(CultureInfo.InvariantCulture) & " signature=C")
                Try
                    Dim a As Object() = {0.0R, 0.0R, 0.0R, 0.0R}
                    s.Obj.GetType().InvokeMember("Range", InvokeFlags, Nothing, s.Obj, a)
                    x1 = SafeD(a(0)) : y1 = SafeD(a(1)) : x2 = SafeD(a(2)) : y2 = SafeD(a(3))
                    ok = True
                Catch ex As Exception
                    log(LogPrefix & "[DVLINE][RANGE][FAIL] id=" & s.Id.ToString(CultureInfo.InvariantCulture) & " signature=C error=" & ex.Message)
                End Try
            End If
            If ok Then
                okCount += 1
                log(LogPrefix & "[DVLINE][RANGE][OK] id=" & s.Id.ToString(CultureInfo.InvariantCulture) &
                    " xmin=" & F(Math.Min(x1, x2)) & " ymin=" & F(Math.Min(y1, y2)) &
                    " xmax=" & F(Math.Max(x1, x2)) & " ymax=" & F(Math.Max(y1, y2)))
                Dim dx As Double = Math.Abs(x2 - x1)
                Dim dy As Double = Math.Abs(y2 - y1)
                Dim orientation As String = If(dx >= dy, "horizontal", "vertical")
                log(LogPrefix & "[DVLINE][RANGE_DERIVED] id=" & s.Id.ToString(CultureInfo.InvariantCulture) &
                    " orientation=" & orientation & " mid=(" & F((x1 + x2) * 0.5R) & "," & F((y1 + y2) * 0.5R) & ") dx=" & F(dx) & " dy=" & F(dy))
            End If
        Next
        Return okCount
    End Function

    Private Shared Function ProbeLineKeypoints(samples As List(Of DvLineSample), log As Action(Of String)) As Integer
        Dim okCount As Integer = 0
        For Each s In samples
            Dim kpc As Integer = CInt(SafeD(CallByNameSafe(s.Obj, "KeyPointCount")))
            Dim localOk As Integer = 0
            For Each kpIdx In New Integer() {0, 1, 2, 3}
                Dim x As Double = 0, y As Double = 0
                Dim sig As String = ""
                log(LogPrefix & "[DVLINE][KEYPOINT][TRY] id=" & s.Id.ToString(CultureInfo.InvariantCulture) & " kp=" & kpIdx.ToString(CultureInfo.InvariantCulture) & " signature=MULTI")
                If TryGetKeyPointAny(s.Obj, kpIdx, x, y, sig) Then
                    log(LogPrefix & "[DVLINE][KEYPOINT][OK] id=" & s.Id.ToString(CultureInfo.InvariantCulture) & " kp=" & kpIdx.ToString(CultureInfo.InvariantCulture) &
                        " x=" & F(x) & " y=" & F(y) & " signature=" & sig)
                    localOk += 1
                Else
                    log(LogPrefix & "[DVLINE][KEYPOINT][FAIL] id=" & s.Id.ToString(CultureInfo.InvariantCulture) & " kp=" & kpIdx.ToString(CultureInfo.InvariantCulture) & " signature=MULTI error=no_matching_signature")
                End If
            Next
            If kpc > 0 AndAlso localOk = 0 Then
                log(LogPrefix & "[DVLINE][KEYPOINT][CONTRADICTION] id=" & s.Id.ToString(CultureInfo.InvariantCulture) & " keyPointCount=" & kpc.ToString(CultureInfo.InvariantCulture))
            End If
            If localOk > 0 Then okCount += 1
        Next
        Return okCount
    End Function

    Private Shared Function ProbeLineReferences(samples As List(Of DvLineSample), log As Action(Of String)) As Integer
        Dim okCount As Integer = 0
        For Each s In samples
            Dim r As Object = CallByNameSafe(s.Obj, "Reference")
            If r IsNot Nothing Then
                okCount += 1
                log(LogPrefix & "[DVLINE][REFERENCE][OK] id=" & s.Id.ToString(CultureInfo.InvariantCulture) &
                    " type=" & r.GetType().FullName & " objectType=" & SafeStr(CallByNameSafe(r, "Type")))
            Else
                log(LogPrefix & "[DVLINE][REFERENCE][FAIL] id=" & s.Id.ToString(CultureInfo.InvariantCulture) & " error=reference_is_nothing")
            End If

            Dim keyBytes As Integer = 0
            Dim keySig As String = ""
            If TryGetReferenceKeyAny(s.Obj, keyBytes, keySig) Then
                log(LogPrefix & "[DVLINE][REFERENCEKEY][OK] id=" & s.Id.ToString(CultureInfo.InvariantCulture) & " bytes=" & keyBytes.ToString(CultureInfo.InvariantCulture) & " signature=" & keySig)
            Else
                log(LogPrefix & "[DVLINE][REFERENCEKEY][FAIL] id=" & s.Id.ToString(CultureInfo.InvariantCulture) & " error=no_matching_signature")
            End If
        Next
        Return okCount
    End Function

    Private Shared Function ProbeGraphicMembers(v As ViewInfo, log As Action(Of String)) As Integer
        Dim gm As Object = CallByNameSafe(v.Dv, "GraphicMembers")
        Dim c As Integer = SafeCount(gm)
        log(LogPrefix & "[GRAPHICMEMBERS][COUNT] view=" & v.Name & " count=" & c.ToString(CultureInfo.InvariantCulture))
        For i As Integer = 1 To Math.Min(c, 30)
            Dim g As Object = Nothing
            Try : g = CallByName(gm, "Item", CallType.Method, i) : Catch : g = Nothing : End Try
            If g Is Nothing Then Continue For
            log(LogPrefix & "[GRAPHICMEMBER][INFO] idx=" & i.ToString(CultureInfo.InvariantCulture) &
                " type=" & SafeStr(CallByNameSafe(g, "Type")) &
                " edgeType=" & SafeStr(CallByNameSafe(g, "EdgeType")) &
                " modelMemberOk=" & (CallByNameSafe(g, "ModelMember") IsNot Nothing).ToString(CultureInfo.InvariantCulture))
        Next
        Return c
    End Function

    Private Shared Function ProbeGetReferenceToGraphicMember(v As ViewInfo, samples As List(Of DvLineSample), log As Action(Of String)) As Integer
        Dim okCount As Integer = 0
        If samples.Count = 0 Then Return 0
        Dim ln As Object = samples(0).Obj

        okCount += TryGetRefToGm(v.Dv, ln, "DVLine2d", log)
        okCount += TryGetRefToGm(v.Dv, CallByNameSafe(ln, "Reference"), "DVLine2d.Reference", log)
        okCount += TryGetRefToGm(v.Dv, CallByNameSafe(ln, "ModelMember"), "DVLine2d.ModelMember", log)

        Dim gmCol As Object = CallByNameSafe(v.Dv, "GraphicMembers")
        If SafeCount(gmCol) > 0 Then
            Dim gm As Object = Nothing
            Try : gm = CallByName(gmCol, "Item", CallType.Method, 1) : Catch : gm = Nothing : End Try
            okCount += TryGetRefToGm(v.Dv, gm, "GraphicMember", log)
        End If
        Return okCount
    End Function

    Private Shared Function TryGetRefToGm(dv As DrawingView, inputObj As Object, tag As String, log As Action(Of String)) As Integer
        If inputObj Is Nothing Then
            log(LogPrefix & "[GET_REF_TO_GM][FAIL] input=" & tag & " error=input_nothing")
            Return 0
        End If
        log(LogPrefix & "[GET_REF_TO_GM][TRY] input=" & tag)
        Dim outRef As Object = Nothing
        Dim usedSig As String = ""
        If TryGetReferenceToGraphicMemberAny(dv, inputObj, outRef, usedSig) Then
            log(LogPrefix & "[GET_REF_TO_GM][OK] input=" & tag & " refType=" & outRef.GetType().FullName & " signature=" & usedSig)
            Return 1
        End If
        log(LogPrefix & "[GET_REF_TO_GM][FAIL] input=" & tag & " error=no_matching_signature")
        Return 0
    End Function

    Private Shared Function TryGetKeyPointAny(obj As Object, kpIdx As Integer, ByRef x As Double, ByRef y As Double, ByRef signature As String) As Boolean
        x = 0 : y = 0 : signature = ""
        If obj Is Nothing Then Return False

        Try
            Dim dx As Double = 0, dy As Double = 0
            CallByName(obj, "GetKeyPoint", CallType.Method, kpIdx, dx, dy)
            x = dx : y = dy : signature = "A_idx_double_double"
            Return True
        Catch
        End Try

        Try
            Dim ox As Object = 0.0R, oy As Object = 0.0R
            CallByName(obj, "GetKeyPoint", CallType.Method, kpIdx, ox, oy)
            x = SafeD(ox) : y = SafeD(oy) : signature = "B_idx_obj_obj"
            Return True
        Catch
        End Try

        Try
            Dim ox As Object = 0.0R, oy As Object = 0.0R, oz As Object = 0.0R
            CallByName(obj, "GetKeyPoint", CallType.Method, kpIdx, ox, oy, oz)
            x = SafeD(ox) : y = SafeD(oy) : signature = "C_idx_obj_obj_obj"
            Return True
        Catch
        End Try

        Try
            Dim arr As Object() = {kpIdx, 0.0R, 0.0R}
            obj.GetType().InvokeMember("GetKeyPoint", InvokeFlags, Nothing, obj, arr)
            x = SafeD(arr(1)) : y = SafeD(arr(2)) : signature = "D_invoke_idx_byrefxy"
            Return True
        Catch
        End Try

        Try
            Dim arr As Object() = {kpIdx, 0.0R, 0.0R, 0.0R}
            obj.GetType().InvokeMember("GetKeyPoint", InvokeFlags, Nothing, obj, arr)
            x = SafeD(arr(1)) : y = SafeD(arr(2)) : signature = "E_invoke_idx_byrefxyz"
            Return True
        Catch
        End Try

        Try
            Dim p As Object = CallByName(obj, "GetKeyPoint", CallType.Method, kpIdx)
            x = SafeD(CallByNameSafe(p, "X")) : y = SafeD(CallByNameSafe(p, "Y"))
            signature = "F_idx_returns_point"
            Return True
        Catch
        End Try

        Return False
    End Function

    Private Shared Function TryGetReferenceKeyAny(obj As Object, ByRef bytes As Integer, ByRef signature As String) As Boolean
        bytes = 0 : signature = ""
        If obj Is Nothing Then Return False

        Try
            Dim k As Object = Nothing
            CallByName(obj, "GetReferenceKey", CallType.Method, k)
            bytes = ByteLength(k) : signature = "A_byref_object"
            Return bytes > 0
        Catch
        End Try

        Try
            Dim k As Object = CallByName(obj, "GetReferenceKey", CallType.Method)
            bytes = ByteLength(k) : signature = "B_returns_value"
            Return bytes > 0
        Catch
        End Try

        Try
            Dim arr As Object() = {Nothing}
            obj.GetType().InvokeMember("GetReferenceKey", InvokeFlags, Nothing, obj, arr)
            bytes = ByteLength(arr(0)) : signature = "C_invoke_byref_object"
            Return bytes > 0
        Catch
        End Try

        Try
            Dim k As Object = obj.GetType().InvokeMember("GetReferenceKey", InvokeFlags, Nothing, obj, Nothing)
            bytes = ByteLength(k) : signature = "D_invoke_returns_value"
            Return bytes > 0
        Catch
        End Try

        Return False
    End Function

    Private Shared Function TryGetReferenceToGraphicMemberAny(dv As DrawingView, inputObj As Object, ByRef outRef As Object, ByRef signature As String) As Boolean
        outRef = Nothing : signature = ""
        If dv Is Nothing OrElse inputObj Is Nothing Then Return False

        Dim attempts As New List(Of Tuple(Of String, Func(Of Object))) From {
            Tuple.Create("M1_obj", CType(Function() CallByName(dv, "GetReferenceToGraphicMember", CallType.Method, inputObj), Func(Of Object))),
            Tuple.Create("M1_obj_int1", CType(Function() CallByName(dv, "GetReferenceToGraphicMember", CallType.Method, inputObj, 1), Func(Of Object))),
            Tuple.Create("M1_obj_true", CType(Function() CallByName(dv, "GetReferenceToGraphicMember", CallType.Method, inputObj, True), Func(Of Object))),
            Tuple.Create("M1_invoke_obj", CType(Function() dv.GetType().InvokeMember("GetReferenceToGraphicMember", InvokeFlags, Nothing, dv, New Object() {inputObj}), Func(Of Object))),
            Tuple.Create("M1_invoke_obj_int1", CType(Function() dv.GetType().InvokeMember("GetReferenceToGraphicMember", InvokeFlags, Nothing, dv, New Object() {inputObj, 1}), Func(Of Object))),
            Tuple.Create("M2_obj", CType(Function() CallByName(dv, "GetReferenceToGraphicMember2", CallType.Method, inputObj), Func(Of Object))),
            Tuple.Create("M2_obj_int1", CType(Function() CallByName(dv, "GetReferenceToGraphicMember2", CallType.Method, inputObj, 1), Func(Of Object))),
            Tuple.Create("M2_invoke_obj", CType(Function() dv.GetType().InvokeMember("GetReferenceToGraphicMember2", InvokeFlags, Nothing, dv, New Object() {inputObj}), Func(Of Object)))
        }

        For Each at In attempts
            Try
                Dim r As Object = at.Item2.Invoke()
                If r IsNot Nothing Then
                    outRef = r
                    signature = at.Item1
                    Return True
                End If
            Catch
            End Try
        Next
        Return False
    End Function

    Private Shared Function ProbeArcs(v As ViewInfo, log As Action(Of String)) As Integer
        Dim okCenter As Integer = 0
        Dim arcs As Object = CallByNameSafe(v.Dv, "DVArcs2d")
        For i As Integer = 1 To SafeCount(arcs)
            Dim a As Object = Nothing
            Try : a = CallByName(arcs, "Item", CallType.Method, i) : Catch : a = Nothing : End Try
            If a Is Nothing Then Continue For
            log(LogPrefix & "[DVARC][BASIC] id=" & i.ToString(CultureInfo.InvariantCulture) &
                " radius=" & F(SafeD(CallByNameSafe(a, "Radius"))) &
                " startAngle=" & F(SafeD(CallByNameSafe(a, "StartAngle"))) &
                " sweepAngle=" & F(SafeD(CallByNameSafe(a, "SweepAngle"))))
            Dim cx As Double = 0, cy As Double = 0
            Try
                CallByName(a, "GetCenterPoint", CallType.Method, cx, cy)
                okCenter += 1
                log(LogPrefix & "[DVARC][CENTER][OK] id=" & i.ToString(CultureInfo.InvariantCulture) & " x=" & F(cx) & " y=" & F(cy))
            Catch ex As Exception
                log(LogPrefix & "[DVARC][CENTER][FAIL] id=" & i.ToString(CultureInfo.InvariantCulture) & " error=" & ex.Message)
            End Try
        Next
        Return okCenter
    End Function

    Private Shared Function ProbeCircles(v As ViewInfo, log As Action(Of String)) As Integer
        Dim okCenter As Integer = 0
        Dim circles As Object = CallByNameSafe(v.Dv, "DVCircles2d")
        For i As Integer = 1 To SafeCount(circles)
            Dim c As Object = Nothing
            Try : c = CallByName(circles, "Item", CallType.Method, i) : Catch : c = Nothing : End Try
            If c Is Nothing Then Continue For
            Dim cx As Double = 0, cy As Double = 0
            Try
                CallByName(c, "GetCenterPoint", CallType.Method, cx, cy)
                okCenter += 1
                log(LogPrefix & "[DVCIRCLE][CENTER][OK] id=" & i.ToString(CultureInfo.InvariantCulture) & " x=" & F(cx) & " y=" & F(cy))
            Catch ex As Exception
                log(LogPrefix & "[DVCIRCLE][CENTER][FAIL] id=" & i.ToString(CultureInfo.InvariantCulture) & " error=" & ex.Message)
            End Try
        Next
        Return okCenter
    End Function

    Private Shared Function ProbeDvPoints(v As ViewInfo, log As Action(Of String)) As Integer
        Dim okxy As Integer = 0
        Dim pts As Object = CallByNameSafe(v.Dv, "DVPoints2d")
        For i As Integer = 1 To SafeCount(pts)
            Dim p As Object = Nothing
            Try : p = CallByName(pts, "Item", CallType.Method, i) : Catch : p = Nothing : End Try
            If p Is Nothing Then Continue For
            Dim x As Double = SafeD(CallByNameSafe(p, "x"))
            Dim y As Double = SafeD(CallByNameSafe(p, "y"))
            If Not (IsZero(x) AndAlso IsZero(y)) Then
                okxy += 1
                log(LogPrefix & "[DVPOINT][XY][OK] id=" & i.ToString(CultureInfo.InvariantCulture) & " x=" & F(x) & " y=" & F(y))
            End If
        Next
        Return okxy
    End Function

    Private Shared Sub ProbeTransforms(v As ViewInfo, samples As List(Of DvLineSample), log As Action(Of String))
        Dim s As DvLineSample = samples.FirstOrDefault(Function(x) x.HasCoords)
        If s Is Nothing Then Return
        Dim sx As Double = 0, sy As Double = 0
        Try
            v.Dv.ViewToSheet(s.X1, s.Y1, sx, sy)
            log(LogPrefix & "[TRANSFORM][VIEW_TO_SHEET] view=" & v.Name & " in=(" & F(s.X1) & "," & F(s.Y1) & ") out=(" & F(sx) & "," & F(sy) & ")")
        Catch ex As Exception
            log(LogPrefix & "[TRANSFORM][VIEW_TO_SHEET][FAIL] view=" & v.Name & " error=" & ex.Message)
        End Try

        Dim vx As Double = 0, vy As Double = 0
        Try
            v.Dv.SheetToView(sx, sy, vx, vy)
            log(LogPrefix & "[TRANSFORM][SHEET_TO_VIEW] view=" & v.Name & " in=(" & F(sx) & "," & F(sy) & ") out=(" & F(vx) & "," & F(vy) & ")")
            Dim delta As Double = Math.Sqrt((vx - s.X1) * (vx - s.X1) + (vy - s.Y1) * (vy - s.Y1))
            log(LogPrefix & "[TRANSFORM][ROUNDTRIP] view=" & v.Name &
                " original=(" & F(s.X1) & "," & F(s.Y1) & ") sheet=(" & F(sx) & "," & F(sy) &
                ") back=(" & F(vx) & "," & F(vy) & ") delta=" & F(delta))
        Catch ex As Exception
            log(LogPrefix & "[TRANSFORM][SHEET_TO_VIEW][FAIL] view=" & v.Name & " error=" & ex.Message)
        End Try

        If sx < v.MinX OrElse sx > v.MaxX OrElse sy < v.MinY OrElse sy > v.MaxY Then
            log(LogPrefix & "[TRANSFORM][WARN_OUTSIDE_RANGE] point=(" & F(sx) & "," & F(sy) & ") range=(" &
                F(v.MinX) & "," & F(v.MinY) & ")-(" & F(v.MaxX) & "," & F(v.MaxY) & ")")
        End If
    End Sub

    Private Shared Function BuildCandidates(v As ViewInfo, samples As List(Of DvLineSample), log As Action(Of String),
                                            ByRef cH As Integer, ByRef cV As Integer, ByRef cP As Integer) As List(Of Candidate)
        Dim out As New List(Of Candidate)()
        Dim valid = samples.Where(Function(s) s.HasCoords).ToList()
        If valid.Count = 0 Then Return out

        Dim minX As Double = valid.Min(Function(s) Math.Min(s.X1, s.X2))
        Dim maxX As Double = valid.Max(Function(s) Math.Max(s.X1, s.X2))
        Dim minY As Double = valid.Min(Function(s) Math.Min(s.Y1, s.Y2))
        Dim maxY As Double = valid.Max(Function(s) Math.Max(s.Y1, s.Y2))

        Dim verticals = valid.Where(Function(s) Math.Abs(s.X2 - s.X1) <= Tol).OrderBy(Function(s) (s.X1 + s.X2) * 0.5R).ToList()
        Dim horizontals = valid.Where(Function(s) Math.Abs(s.Y2 - s.Y1) <= Tol).OrderBy(Function(s) (s.Y1 + s.Y2) * 0.5R).ToList()

        If verticals.Count >= 2 Then
            Dim l = verticals.First()
            Dim r = verticals.Last()
            Dim exp As Double = Math.Abs(((r.X1 + r.X2) * 0.5R) - ((l.X1 + l.X2) * 0.5R))
            If exp > 0 Then
                out.Add(New Candidate With {
                    .Kind = "H_TOTAL", .O1 = l.Obj, .O2 = r.Obj,
                    .X1 = (l.X1 + l.X2) * 0.5R, .Y1 = minY - PlacementOffset,
                    .X2 = (r.X1 + r.X2) * 0.5R, .Y2 = minY - PlacementOffset,
                    .Expected = exp
                })
                cH += 1
                log(LogPrefix & "[CAND][H_TOTAL] source=DVLine_coords left=DVLine#" & l.Id.ToString(CultureInfo.InvariantCulture) &
                    " right=DVLine#" & r.Id.ToString(CultureInfo.InvariantCulture) & " expected=" & F(exp))
            End If
        End If

        If horizontals.Count >= 2 Then
            Dim b = horizontals.First()
            Dim t = horizontals.Last()
            Dim exp As Double = Math.Abs(((t.Y1 + t.Y2) * 0.5R) - ((b.Y1 + b.Y2) * 0.5R))
            If exp > 0 Then
                out.Add(New Candidate With {
                    .Kind = "V_TOTAL", .O1 = b.Obj, .O2 = t.Obj,
                    .X1 = minX - PlacementOffset, .Y1 = (b.Y1 + b.Y2) * 0.5R,
                    .X2 = minX - PlacementOffset, .Y2 = (t.Y1 + t.Y2) * 0.5R,
                    .Expected = exp
                })
                cV += 1
                log(LogPrefix & "[CAND][V_TOTAL] source=DVLine_coords bottom=DVLine#" & b.Id.ToString(CultureInfo.InvariantCulture) &
                    " top=DVLine#" & t.Id.ToString(CultureInfo.InvariantCulture) & " expected=" & F(exp))
            End If
        End If

        Dim ordered = valid.OrderBy(Function(s) (s.X1 + s.X2) * 0.5R).ToList()
        For i As Integer = 0 To ordered.Count - 2
            Dim exp As Double = Math.Abs(((ordered(i + 1).X1 + ordered(i + 1).X2) * 0.5R) - ((ordered(i).X1 + ordered(i).X2) * 0.5R))
            If exp <= 0 Then Continue For
            cP += 1
            log(LogPrefix & "[CAND][PARTIAL] source=DVLine_coords e1=DVLine#" & ordered(i).Id.ToString(CultureInfo.InvariantCulture) &
                " e2=DVLine#" & ordered(i + 1).Id.ToString(CultureInfo.InvariantCulture) & " expected=" & F(exp))
        Next
        Return out
    End Function

    Private Shared Function TryAddDistanceGeneric(dims As Dimensions, o1 As Object, o2 As Object, x1 As Double, y1 As Double, x2 As Double, y2 As Double, axis As String, dv As DrawingView, ByRef dOut As FrameworkDimension) As Boolean
        dOut = Nothing
        Dim bridge As New Logger()
        Dim dimLog As New DimensionLogger(bridge)
        Dim m As String = ""
        Return DimensionPlacementEngine.TryAddDistanceBetweenObjectsGeneric(dims, o1, o2, x1, y1, x2, y2, dimLog, axis, m, "DV_METHODLAB_" & axis, Nothing, dv, True, dOut)
    End Function

    Private Shared Function TryAddDistanceEx(dims As Dimensions, o1 As Object, o2 As Object, x1 As Double, y1 As Double, x2 As Double, y2 As Double, ByRef dOut As FrameworkDimension, log As Action(Of String)) As Boolean
        dOut = Nothing
        If dims Is Nothing OrElse o1 Is Nothing OrElse o2 Is Nothing Then Return False
        Try
            Dim d As Object = CallByName(dims, "AddDistanceBetweenObjectsEX", CallType.Method, o1, o2, x1, y1, x2, y2)
            dOut = TryCast(d, FrameworkDimension)
            Return dOut IsNot Nothing
        Catch ex As Exception
            log(LogPrefix & "[DIM][FAIL] method=ADBOEX error=" & ex.Message)
            Return False
        End Try
    End Function

    Private Shared Function TryAddByDimInitData(dims As Dimensions, c As Candidate, ByRef dOut As FrameworkDimension, log As Action(Of String)) As Boolean
        dOut = Nothing
        If dims Is Nothing OrElse c Is Nothing Then Return False
        Try
            Dim initType = Type.GetTypeFromProgID("SolidEdgeFrameworkSupport.DimInitData", False)
            If initType Is Nothing Then
                log(LogPrefix & "[DIM][FAIL] method=DimInitData error=progid_not_found")
                Return False
            End If
            Dim initObj As Object = Activator.CreateInstance(initType)
            CallByName(initObj, "SetNumberOfParents", CallType.Method, 2)
            CallByName(initObj, "SetParentByIndex", CallType.Method, 1, c.O1)
            CallByName(initObj, "SetParentByIndex", CallType.Method, 2, c.O2)
            CallByName(initObj, "SetDimPosition", CallType.Method, (c.X1 + c.X2) * 0.5R, (c.Y1 + c.Y2) * 0.5R, 0.0R)
            Dim d As Object = CallByName(dims, "AddDimension", CallType.Method, initObj)
            dOut = TryCast(d, FrameworkDimension)
            Return dOut IsNot Nothing
        Catch ex As Exception
            log(LogPrefix & "[DIM][FAIL] method=DimInitData error=" & ex.Message & " NO_CONFIRMADO")
            Return False
        End Try
    End Function

    Private Shared Function TryAddAuxLength(sh As Sheet, dims As Dimensions, v As ViewInfo, auxLayer As Object, ByRef dOut As FrameworkDimension, log As Action(Of String)) As Boolean
        dOut = Nothing
        If sh Is Nothing OrElse dims Is Nothing Then Return False
        Dim y As Double = v.MinY - PlacementOffset
        Dim aux As Object = Nothing
        Try : aux = CallByName(sh.Lines2d, "AddBy2Points", CallType.Method, v.MinX, y, v.MaxX, y) : Catch : aux = Nothing : End Try
        If aux Is Nothing Then Return False
        Try : CallByName(aux, "Layer", CallType.Let, auxLayer) : Catch : End Try
        Try : dOut = TryCast(CallByName(dims, "AddLength", CallType.Method, aux), FrameworkDimension) : Catch : dOut = Nothing : End Try
        Return dOut IsNot Nothing
    End Function

    Private Shared Function AcceptDimension(d As FrameworkDimension, expected As Double, log As Action(Of String), method As String, ByRef connected As Integer, ByRef floating As Integer) As Boolean
        If d Is Nothing Then
            log(LogPrefix & "[DIM][FAIL] method=" & method & " error=no_dimension_returned")
            Return False
        End If
        Dim val As Double = SafeD(CallByNameSafe(d, "Value"))
        If val = 0 Then
            log(LogPrefix & "[DIM][REJECT_ZERO] method=" & method & " value=0 expected=" & F(expected))
            Try : CallByName(d, "Delete", CallType.Method) : Catch : End Try
            Return False
        End If
        Dim delta As Double = If(expected > 0, Math.Abs(val - expected), 0)
        log(LogPrefix & "[DIM][OK] method=" & method & " value=" & F(val) & " expected=" & F(expected) & " delta=" & F(delta))

        Dim relObj As Object = Nothing
        Try : relObj = CallByName(d, "GetRelatedObjects", CallType.Method) : Catch : relObj = Nothing : End Try
        Dim relCount As Integer = SafeCount(relObj)
        log(LogPrefix & "[DIM][RELATED] method=" & method & " count=" & relCount.ToString(CultureInfo.InvariantCulture) & " sig=NO_CONFIRMADO")
        log(LogPrefix & "[DIM][STATUS] method=" & method & " status=" & SafeStr(CallByNameSafe(d, "Status")))

        Dim dd As Object = Nothing
        Try : dd = CallByName(d, "GetDisplayData", CallType.Method) : Catch : dd = Nothing : End Try
        Dim dl As Integer = SafeCount(CallByNameSafe(dd, "Lines2d"))
        Dim da As Integer = SafeCount(CallByNameSafe(dd, "Arcs2d"))
        log(LogPrefix & "[DIM][DISPLAY] method=" & method & " displayLines=" & dl.ToString(CultureInfo.InvariantCulture) & " displayArcs=" & da.ToString(CultureInfo.InvariantCulture))

        If relCount > 0 Then connected += 1 Else floating += 1
        Return True
    End Function

    Private Shared Function TryRetrieveDimensions(dv As DrawingView, sh As Sheet, log As Action(Of String)) As Boolean
        If dv Is Nothing Then Return False
        log(LogPrefix & "[RETRIEVE_DIMS][TRY] view=" & SafeStr(CallByNameSafe(dv, "Name")))
        Dim before As Integer = SafeCount(CallByNameSafe(sh, "Dimensions"))
        Try
            CallByName(dv, "RetrieveDimensions", CallType.Method)
            log(LogPrefix & "[RETRIEVE_DIMS][OK] createdOrFound=NO_CONFIRMADO")
        Catch ex As Exception
            log(LogPrefix & "[RETRIEVE_DIMS][FAIL] error=" & ex.Message)
        End Try

        log(LogPrefix & "[CONNECTED_DIMS_SELECTSET][TRY] view=" & SafeStr(CallByNameSafe(dv, "Name")))
        Try
            Dim ss As Object = CallByNameSafe(sh, "SelectSet")
            Dim cBefore As Integer = SafeCount(ss)
            CallByName(dv, "AddConnectedDimensionsToSelectSet", CallType.Method)
            Dim cAfter As Integer = SafeCount(ss)
            log(LogPrefix & "[CONNECTED_DIMS_SELECTSET][OK] countBefore=" & cBefore.ToString(CultureInfo.InvariantCulture) & " countAfter=" & cAfter.ToString(CultureInfo.InvariantCulture))
        Catch ex As Exception
            log(LogPrefix & "[CONNECTED_DIMS_SELECTSET][FAIL] error=" & ex.Message)
        End Try

        log(LogPrefix & "[ARRANGE_DIMS][TRY]")
        Dim arranged As Boolean = False
        Try
            Dim ss As Object = CallByNameSafe(sh, "SelectSet")
            CallByName(sh, "ArrangeDimensionsInSelectSet", CallType.Method, ss)
            arranged = True
            log(LogPrefix & "[ARRANGE_DIMS][OK] signature=with_selectset")
        Catch ex As Exception
            log(LogPrefix & "[ARRANGE_DIMS][FAIL] signature=with_selectset error=" & ex.Message)
        End Try
        If Not arranged Then
            Try
                CallByName(sh, "ArrangeDimensionsInSelectSet", CallType.Method)
                arranged = True
                log(LogPrefix & "[ARRANGE_DIMS][OK] signature=no_args")
            Catch ex As Exception
                log(LogPrefix & "[ARRANGE_DIMS][FAIL] signature=no_args error=" & ex.Message)
            End Try
        End If

        Dim after As Integer = SafeCount(CallByNameSafe(sh, "Dimensions"))
        Return after >= before
    End Function

    Private Shared Function ResolveStyleObject(draft As DraftDocument, log As Action(Of String)) As Object
        log(LogPrefix & "[STYLE][FIND] target=" & DimStyleName)
        Dim styles As Object = CallByNameSafe(draft, "DimensionStyles")
        If styles Is Nothing Then
            log(LogPrefix & "[STYLE][FAIL] error=dimensionstyles_collection_missing")
            Return Nothing
        End If
        For i As Integer = 1 To SafeCount(styles)
            Dim st As Object = Nothing
            Try : st = CallByName(styles, "Item", CallType.Method, i) : Catch : st = Nothing : End Try
            If st Is Nothing Then Continue For
            Dim nm As String = SafeStr(CallByNameSafe(st, "Name"))
            If String.Equals(nm, DimStyleName, StringComparison.OrdinalIgnoreCase) Then
                log(LogPrefix & "[STYLE][OK] name=U3,5")
                Return st
            End If
        Next
        log(LogPrefix & "[STYLE][FAIL] error=not_found")
        Return Nothing
    End Function

    Private Shared Sub TrySetCollectionStyle(dims As Dimensions, styleObj As Object, log As Action(Of String))
        log(LogPrefix & "[STYLE][SET_COLLECTION_TRY]")
        If dims Is Nothing OrElse styleObj Is Nothing Then
            log(LogPrefix & "[STYLE][SET_COLLECTION_FAIL] reason=no_dims_or_style")
            Return
        End If
        Try
            CallByName(dims, "Style", CallType.Let, styleObj)
            log(LogPrefix & "[STYLE][SET_COLLECTION_OK]")
        Catch ex As Exception
            log(LogPrefix & "[STYLE][SET_COLLECTION_FAIL] " & ex.Message)
        End Try
    End Sub

    Private Shared Sub ApplyStyle(d As FrameworkDimension, styleObj As Object, log As Action(Of String))
        If d Is Nothing OrElse styleObj Is Nothing Then Return
        Try
            CallByName(d, "Style", CallType.Let, styleObj)
        Catch ex As Exception
            log(LogPrefix & "[STYLE][FAIL] " & ex.Message & " -> trying StyleName")
            Try
                CallByName(d, "StyleName", CallType.Let, DimStyleName)
                log(LogPrefix & "[STYLE][OK] via=StyleName name=" & DimStyleName)
            Catch ex2 As Exception
                log(LogPrefix & "[STYLE][FAIL] via=StyleName error=" & ex2.Message)
            End Try
        End Try
    End Sub

    Private Shared Function EnsureLayer(sh As Sheet, layerName As String) As Object
        Dim layers As Object = CallByNameSafe(sh, "Layers")
        If layers Is Nothing Then Return Nothing
        For i As Integer = 1 To SafeCount(layers)
            Dim ly As Object = Nothing
            Try : ly = CallByName(layers, "Item", CallType.Method, i) : Catch : ly = Nothing : End Try
            If ly Is Nothing Then Continue For
            If String.Equals(SafeStr(CallByNameSafe(ly, "Name")), layerName, StringComparison.OrdinalIgnoreCase) Then Return ly
        Next
        Try : Return CallByName(layers, "Add", CallType.Method, layerName) : Catch : Return Nothing : End Try
    End Function

    Private Shared Function SafeReference(o As Object) As Object
        If o Is Nothing Then Return Nothing
        Try
            Return CallByName(o, "Reference", CallType.Get)
        Catch
            Return Nothing
        End Try
    End Function

    Private Shared Function TypeNameSafe(o As Object) As String
        If o Is Nothing Then Return ""
        Try
            Return o.GetType().FullName
        Catch
            Return ""
        End Try
    End Function

    Private Shared Function ByteLength(obj As Object) As Integer
        If obj Is Nothing Then Return 0
        Try
            Dim arr As Byte() = TryCast(obj, Byte())
            If arr IsNot Nothing Then Return arr.Length
        Catch
        End Try
        Return 0
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
        Return v.ToString("0.######", CultureInfo.InvariantCulture)
    End Function

    Private Shared Function IsZero(v As Double) As Boolean
        Return Math.Abs(v) <= 1.0E-12R
    End Function

End Class

