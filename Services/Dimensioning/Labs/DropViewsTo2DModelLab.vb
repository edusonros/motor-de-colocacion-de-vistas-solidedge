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
''' Laboratorio experimental independiente: prueba de Drop de DrawingViews y acotación sobre geometría 2D real.
''' Se ejecuta solo por flag JobConfiguration.RunDropViewsTo2DModelLab.
''' </summary>
Friend NotInheritable Class DropViewsTo2DModelLab
    Private Const LayerGeometry As String = "DROP_LAB_GEOMETRY"
    Private Const LayerDimensions As String = "DROP_LAB_DIMENSIONS"

    Private Shared ReadOnly DvInvokeFlags As BindingFlags =
        BindingFlags.InvokeMethod Or BindingFlags.Public Or BindingFlags.Instance

    Private Sub New()
    End Sub

    Private NotInheritable Class Group2D
        Public ViewIndex As Integer
        Public ViewName As String
        Public Lines As New List(Of Line2d)()
        Public LineMetas As New List(Of Tuple(Of Line2d, Double, Double, Double, Double))()
        Public Arcs As New List(Of Arc2d)()
        Public Circles As New List(Of Circle2d)()
        Public MinX As Double = Double.MaxValue
        Public MinY As Double = Double.MaxValue
        Public MaxX As Double = Double.MinValue
        Public MaxY As Double = Double.MinValue
    End Class

    Private NotInheritable Class Snapshot2D
        Public Lines As New List(Of Tuple(Of Double, Double, Double, Double))()
        Public Arcs As New List(Of Tuple(Of Double, Double, Double, Double, Double, Double))() ' cx,cy,r,start,end,sweep
        Public Circles As New List(Of Tuple(Of Double, Double, Double))() ' cx,cy,r
    End Class

    Public Shared Sub Run(app As Application, draft As DraftDocument, log As Action(Of String), Optional debugSave As Boolean = False)
        Dim L = Sub(m As String)
                    Try
                        log?.Invoke(m)
                    Catch
                    End Try
                End Sub

        L("[DROP2D][START]")
        If draft Is Nothing Then
            L("[DROP2D][ABORT] reason=no_draft")
            Return
        End If

        Dim draftPath As String = ""
        Try : draftPath = Convert.ToString(draft.FullName, CultureInfo.InvariantCulture) : Catch : End Try
        L("[DROP2D][DFT] path=" & draftPath)

        Dim sourceSheet As Sheet = Nothing
        Dim target2D As Sheet = Nothing
        ResolveSheetsForLab(draft, sourceSheet, target2D, L)
        If sourceSheet Is Nothing OrElse target2D Is Nothing Then
            L("[DROP2D][ABORT] reason=sheet_resolution_failed")
            Return
        End If

        L("[DROP2D][SOURCE_SHEET] name=" & SafeName(sourceSheet))
        L("[DROP2D][TARGET_SHEET] name=2D Model")

        Dim layerGeom As Object = EnsureLayer(target2D, LayerGeometry, L)
        Dim layerDim As Object = EnsureLayer(target2D, LayerDimensions, L)
        Dim deleted As Integer = CleanupPreviousLabGeometry(target2D, L)
        L("[DROP2D][CLEANUP] deleted_previous_lab_geometry=" & deleted.ToString(CultureInfo.InvariantCulture))

        Dim views As New List(Of DrawingView)()
        Dim nViews As Integer = 0
        Try : nViews = sourceSheet.DrawingViews.Count : Catch : nViews = 0 : End Try
        For i As Integer = 1 To nViews
            Dim dv As DrawingView = Nothing
            Try : dv = CType(sourceSheet.DrawingViews.Item(i), DrawingView) : Catch : dv = Nothing : End Try
            If dv IsNot Nothing Then views.Add(dv)
        Next

        Dim groups As New List(Of Group2D)()
        Dim droppedOk As Integer = 0
        Dim droppedFail As Integer = 0
        Dim keypointsFound As Integer = 0
        Dim connectPointsFound As Integer = 0
        Dim dimCreated As Integer = 0
        Dim dimConnected As Integer = 0
        Dim dimFloating As Integer = 0
        Dim bestMethod As String = "none"

        For i As Integer = 0 To views.Count - 1
            Dim dv = views(i)
            Dim idx = i + 1
            Dim name As String = SafeStr(GetCom(dv, "Name"))
            Dim typ As String = SafeStr(GetCom(dv, "DrawingViewType"))
            Dim scale As String = SafeStr(GetCom(dv, "Scale"))
            Dim ori As String = SafeStr(GetCom(dv, "ViewOrientation"))
            Dim isIso As Boolean = ori.IndexOf("iso", StringComparison.OrdinalIgnoreCase) >= 0 OrElse typ.IndexOf("iso", StringComparison.OrdinalIgnoreCase) >= 0

            L("[DROP2D][VIEW][INFO] idx=" & idx.ToString(CultureInfo.InvariantCulture) & " name=" & name & " type=" & typ & " scale=" & scale & " isIso=" & isIso.ToString(CultureInfo.InvariantCulture))

            Dim cntLines As Integer = SafeCount(GetCom(dv, "DVLines2d"))
            Dim cntArcs As Integer = SafeCount(GetCom(dv, "DVArcs2d"))
            Dim cntCircles As Integer = SafeCount(GetCom(dv, "DVCircles2d"))
            Dim cntPoints As Integer = SafeCount(GetCom(dv, "DVPoints2d"))
            L("[DROP2D][VIEW][DVCOUNT] lines=" & cntLines.ToString(CultureInfo.InvariantCulture) & " arcs=" & cntArcs.ToString(CultureInfo.InvariantCulture) & " circles=" & cntCircles.ToString(CultureInfo.InvariantCulture) & " points=" & cntPoints.ToString(CultureInfo.InvariantCulture))

            Dim bx1 As Double = 0, by1 As Double = 0, bx2 As Double = 0, by2 As Double = 0
            Try
                dv.Range(bx1, by1, bx2, by2)
                L("[DROP2D][VIEW][RANGE] xmin=" & FormatInv(Math.Min(bx1, bx2)) & " ymin=" & FormatInv(Math.Min(by1, by2)) & " xmax=" & FormatInv(Math.Max(bx1, bx2)) & " ymax=" & FormatInv(Math.Max(by1, by2)))
            Catch ex As Exception
                L("[DROP2D][VIEW][RANGE] FAIL " & ex.Message)
            End Try

            Dim snap As Snapshot2D = CaptureViewSnapshotBeforeDrop(dv, L)

            Dim dropMethod As String = ""
            If TryDropView(dv, dropMethod, L) Then
                droppedOk += 1
                L("[DROP2D][VIEW][DROP_OK] generated_objects=NO_CONFIRMADO method=" & dropMethod)
            Else
                droppedFail += 1
                L("[DROP2D][VIEW][DROP_FAIL] reason=no_supported_method")
            End If

            Dim g As Group2D = CopySnapshotTo2DModel(snap, target2D, idx, name, layerGeom, L)
            If g IsNot Nothing Then
                groups.Add(g)
                AnalyzeGroupGeometry(g, keypointsFound, connectPointsFound, L)
                Dim createdByGroup As Integer = TryCreateDimensionsOnGroup(target2D, g, layerDim, bestMethod, dimConnected, dimFloating, L)
                dimCreated += createdByGroup
            End If
        Next

        Dim linesTotal As Integer = groups.Sum(Function(g) g.Lines.Count)
        Dim arcsTotal As Integer = groups.Sum(Function(g) g.Arcs.Count)
        Dim circlesTotal As Integer = groups.Sum(Function(g) g.Circles.Count)

        L("[DROP2D][SUMMARY]")
        L("views_processed=" & views.Count.ToString(CultureInfo.InvariantCulture))
        L("views_dropped_ok=" & droppedOk.ToString(CultureInfo.InvariantCulture))
        L("views_dropped_fail=" & droppedFail.ToString(CultureInfo.InvariantCulture))
        L("entities_2d_created_lines=" & linesTotal.ToString(CultureInfo.InvariantCulture))
        L("entities_2d_created_arcs=" & arcsTotal.ToString(CultureInfo.InvariantCulture))
        L("entities_2d_created_circles=" & circlesTotal.ToString(CultureInfo.InvariantCulture))
        L("keypoints_found=" & keypointsFound.ToString(CultureInfo.InvariantCulture))
        L("connectpoints_found=" & connectPointsFound.ToString(CultureInfo.InvariantCulture))
        L("dimensions_created=" & dimCreated.ToString(CultureInfo.InvariantCulture))
        L("dimensions_connected=" & dimConnected.ToString(CultureInfo.InvariantCulture))
        L("dimensions_floating=" & dimFloating.ToString(CultureInfo.InvariantCulture))
        L("best_method=" & If(String.IsNullOrWhiteSpace(bestMethod), "none", bestMethod))
        L("recommended_next_step=validar_si_AddDistanceBetweenObjects_sobre_Lines2d_reales_es_estable")

        If debugSave Then
            Try : draft.Save() : Catch : End Try
        End If
    End Sub

    Private Shared Function TryCreateDimensionsOnGroup(targetSheet As Sheet, g As Group2D, layerDim As Object, ByRef bestMethod As String, ByRef connected As Integer, ByRef floating As Integer, log As Action(Of String)) As Integer
        If targetSheet Is Nothing OrElse g Is Nothing Then Return 0
        Dim dims As Dimensions = Nothing
        Try : dims = CType(targetSheet.Dimensions, Dimensions) : Catch : dims = Nothing : End Try
        If dims Is Nothing Then Return 0

        Try : CallByName(targetSheet, "Activate", CallType.Method) : Catch : End Try

        If g.Lines.Count = 0 AndAlso g.Arcs.Count = 0 AndAlso g.Circles.Count = 0 Then
            log("[DROP2D][GROUP][EMPTY] view=" & g.ViewName & " reason=no_entities_after_copy")
            Return 0
        End If

        Dim created As Integer = 0
        Dim leftLine As Line2d = Nothing
        Dim rightLine As Line2d = Nothing
        Dim bottomLine As Line2d = Nothing
        Dim topLine As Line2d = Nothing
        Dim leftMidX As Double = Double.MaxValue
        Dim rightMidX As Double = Double.MinValue
        Dim bottomMidY As Double = Double.MaxValue
        Dim topMidY As Double = Double.MinValue

        For Each lm In g.LineMetas
            If lm Is Nothing OrElse lm.Item1 Is Nothing Then Continue For
            Dim ln As Line2d = lm.Item1
            Dim x1 As Double = lm.Item2, y1 As Double = lm.Item3, x2 As Double = lm.Item4, y2 As Double = lm.Item5
            Dim dx As Double = Math.Abs(x2 - x1)
            Dim dy As Double = Math.Abs(y2 - y1)
            Dim midX As Double = (x1 + x2) * 0.5R
            Dim midY As Double = (y1 + y2) * 0.5R
            If dx >= dy Then
                If bottomLine Is Nothing OrElse midY < bottomMidY Then
                    bottomLine = ln
                    bottomMidY = midY
                End If
                If topLine Is Nothing OrElse midY > topMidY Then
                    topLine = ln
                    topMidY = midY
                End If
            Else
                If leftLine Is Nothing OrElse midX < leftMidX Then
                    leftLine = ln
                    leftMidX = midX
                End If
                If rightLine Is Nothing OrElse midX > rightMidX Then
                    rightLine = ln
                    rightMidX = midX
                End If
            End If
        Next

        If leftLine IsNot Nothing AndAlso rightLine IsNot Nothing Then
            Dim yRef As Double = (g.MinY + g.MaxY) * 0.5R
            Dim methodUsed As String = ""
            Dim d As FrameworkDimension = Nothing
            log("[DROP2D][DIM][TRY] method=AddDistanceBetweenObjects entity1=Line2d entity2=Line2d")
            Dim ok = DimensionPlacementEngine.TryAddDistanceBetweenObjectsGeneric(dims, leftLine, rightLine, g.MinX, yRef, g.MaxX, yRef, Nothing, "horizontal", methodUsed, "DROP2D_H_TOTAL", Nothing, Nothing, True, d)
            If ok AndAlso d IsNot Nothing Then
                created += 1
                bestMethod = methodUsed
                TrySetLayer(d, layerDim)
                Dim rel As Integer = SafeCount(CallByNameSafe(d, "GetRelatedObjects"))
                Dim st As String = SafeStr(CallByNameSafe(d, "Status"))
                log("[DROP2D][DIM][OK] method=" & methodUsed & " value=" & SafeStr(CallByNameSafe(d, "Value")) & " style=" & SafeStr(CallByNameSafe(d, "StyleName")))
                log("[DROP2D][DIM][RELATED] count=" & rel.ToString(CultureInfo.InvariantCulture) & " sig=NO_CONFIRMADO")
                log("[DROP2D][DIM][STATUS] status=" & st)
                If rel > 0 Then connected += 1 Else floating += 1
            Else
                log("[DROP2D][DIM][FAIL] method=AddDistanceBetweenObjects error=no_dimension_returned")
            End If
        End If

        If bottomLine IsNot Nothing AndAlso topLine IsNot Nothing Then
            Dim xRef As Double = (g.MinX + g.MaxX) * 0.5R
            Dim methodUsed As String = ""
            Dim d As FrameworkDimension = Nothing
            log("[DROP2D][DIM][TRY] method=AddDistanceBetweenObjects entity1=Line2d entity2=Line2d")
            Dim ok = DimensionPlacementEngine.TryAddDistanceBetweenObjectsGeneric(dims, bottomLine, topLine, xRef, g.MinY, xRef, g.MaxY, Nothing, "vertical", methodUsed, "DROP2D_V_TOTAL", Nothing, Nothing, True, d)
            If ok AndAlso d IsNot Nothing Then
                created += 1
                If String.IsNullOrWhiteSpace(bestMethod) Then bestMethod = methodUsed
                TrySetLayer(d, layerDim)
                Dim rel As Integer = SafeCount(CallByNameSafe(d, "GetRelatedObjects"))
                Dim st As String = SafeStr(CallByNameSafe(d, "Status"))
                log("[DROP2D][DIM][OK] method=" & methodUsed & " value=" & SafeStr(CallByNameSafe(d, "Value")) & " style=" & SafeStr(CallByNameSafe(d, "StyleName")))
                log("[DROP2D][DIM][RELATED] count=" & rel.ToString(CultureInfo.InvariantCulture) & " sig=NO_CONFIRMADO")
                log("[DROP2D][DIM][STATUS] status=" & st)
                If rel > 0 Then connected += 1 Else floating += 1
            Else
                log("[DROP2D][DIM][FAIL] method=AddDistanceBetweenObjects error=no_dimension_returned")
            End If
        End If

        ' Sin Line2d válidas (p. ej. snapshot filtró degeneradas) pero sí Arc2d → al menos una prueba real de cota.
        If created = 0 AndAlso g.Arcs IsNot Nothing Then
            For Each ac As Arc2d In g.Arcs
                If ac Is Nothing Then Continue For
                log("[DROP2D][DIM][TRY] method=AddRadius entity=Arc2d reason=no_lines_for_distance")
                Dim dRad As FrameworkDimension = DrawingViewDimensionCreator.TryCreateRadiusOnReference(dims, ac, Nothing)
                If dRad IsNot Nothing Then
                    created += 1
                    If String.IsNullOrWhiteSpace(bestMethod) Then bestMethod = "AddRadius"
                    TrySetLayer(dRad, layerDim)
                    Dim rel As Integer = SafeCount(CallByNameSafe(dRad, "GetRelatedObjects"))
                    Dim st As String = SafeStr(CallByNameSafe(dRad, "Status"))
                    log("[DROP2D][DIM][OK] method=AddRadius value=" & SafeStr(CallByNameSafe(dRad, "Value")) & " style=" & SafeStr(CallByNameSafe(dRad, "StyleName")))
                    log("[DROP2D][DIM][RELATED] count=" & rel.ToString(CultureInfo.InvariantCulture))
                    log("[DROP2D][DIM][STATUS] status=" & st)
                    If rel > 0 Then connected += 1 Else floating += 1
                    Exit For
                End If
            Next
            If created = 0 Then
                log("[DROP2D][DIM][FAIL] method=AddRadius error=no_dimension_returned arcs=" & g.Arcs.Count.ToString(CultureInfo.InvariantCulture))
            End If
        End If

        Return created
    End Function

    Private Shared Sub AnalyzeGroupGeometry(g As Group2D, ByRef keypointsFound As Integer, ByRef connectPointsFound As Integer, log As Action(Of String))
        If g Is Nothing Then Return
        If g.Lines.Count = 0 AndAlso g.Arcs.Count = 0 AndAlso g.Circles.Count = 0 Then
            log("[DROP2D][GROUP][EMPTY] view=" & g.ViewName & " reason=no_entities_for_analysis")
            Return
        End If
        log("[DROP2D][EXTREMES] view=" & g.ViewName & " xmin=" & FormatInv(g.MinX) & " xmax=" & FormatInv(g.MaxX) & " ymin=" & FormatInv(g.MinY) & " ymax=" & FormatInv(g.MaxY))
        Dim entLabel As String = "bbox"
        If g.LineMetas.Count > 0 Then
            entLabel = "Line2d"
        ElseIf g.Arcs.Count > 0 OrElse g.Circles.Count > 0 Then
            entLabel = "Arc/Circle"
        End If
        log("[DROP2D][CAND][H_TOTAL] basis=" & entLabel & " spanX=" & FormatInv((g.MaxX - g.MinX)))
        log("[DROP2D][CAND][V_TOTAL] basis=" & entLabel & " spanY=" & FormatInv((g.MaxY - g.MinY)))

        Dim id As Integer = 0
        For Each lm In g.LineMetas
            If lm Is Nothing Then Continue For
            id += 1
            Dim ln As Line2d = lm.Item1
            Dim x1 As Double = lm.Item2, y1 As Double = lm.Item3, x2 As Double = lm.Item4, y2 As Double = lm.Item5
            Dim L As Double = Math.Sqrt((x2 - x1) * (x2 - x1) + (y2 - y1) * (y2 - y1))
            Dim ang As Double = Math.Atan2(y2 - y1, x2 - x1) * 180.0R / Math.PI
            log("[DROP2D][GEOM][LINE] id=" & id.ToString(CultureInfo.InvariantCulture) & " x1=" & FormatInv(x1) & " y1=" & FormatInv(y1) & " x2=" & FormatInv(x2) & " y2=" & FormatInv(y2) & " length=" & FormatInv(L) & " angle=" & FormatInv(ang))
            If ln IsNot Nothing Then
                keypointsFound += LogKeypoints("Line2d", ln, log)
                connectPointsFound += LogConnectPoints("Line2d", ln, log)
            End If
        Next
        id = 0
        For Each ac In g.Arcs
            id += 1
            Dim cx As Double = 0, cy As Double = 0, r As Double = 0, sa As Double = 0, ea As Double = 0
            Try : ac.GetCenterPoint(cx, cy) : Catch : log("[DROP2D][GEOM][PROP_MISSING] entity=Arc2d prop=GetCenterPoint") : End Try
            Try : r = Convert.ToDouble(CallByName(ac, "Radius", CallType.Get), CultureInfo.InvariantCulture) : Catch : log("[DROP2D][GEOM][PROP_MISSING] entity=Arc2d prop=Radius") : End Try
            Try : sa = Convert.ToDouble(CallByName(ac, "StartAngle", CallType.Get), CultureInfo.InvariantCulture) : Catch : log("[DROP2D][GEOM][PROP_MISSING] entity=Arc2d prop=StartAngle") : End Try
            Try : ea = Convert.ToDouble(CallByName(ac, "EndAngle", CallType.Get), CultureInfo.InvariantCulture) : Catch : log("[DROP2D][GEOM][PROP_MISSING] entity=Arc2d prop=EndAngle") : End Try
            log("[DROP2D][GEOM][ARC] id=" & id.ToString(CultureInfo.InvariantCulture) & " cx=" & FormatInv(cx) & " cy=" & FormatInv(cy) & " r=" & FormatInv(r) & " start=" & FormatInv(sa) & " end=" & FormatInv(ea))
            keypointsFound += LogKeypoints("Arc2d", ac, log)
            connectPointsFound += LogConnectPoints("Arc2d", ac, log)
        Next
        id = 0
        For Each cc In g.Circles
            id += 1
            Dim cx As Double = 0, cy As Double = 0, r As Double = 0
            Try : cc.GetCenterPoint(cx, cy) : Catch : log("[DROP2D][GEOM][PROP_MISSING] entity=Circle2d prop=GetCenterPoint") : End Try
            Try : r = Convert.ToDouble(CallByName(cc, "Radius", CallType.Get), CultureInfo.InvariantCulture) : Catch : log("[DROP2D][GEOM][PROP_MISSING] entity=Circle2d prop=Radius") : End Try
            log("[DROP2D][GEOM][CIRCLE] id=" & id.ToString(CultureInfo.InvariantCulture) & " cx=" & FormatInv(cx) & " cy=" & FormatInv(cy) & " r=" & FormatInv(r))
            keypointsFound += LogKeypoints("Circle2d", cc, log)
            connectPointsFound += LogConnectPoints("Circle2d", cc, log)
        Next
    End Sub

    Private Shared Function LogKeypoints(entityName As String, entityObj As Object, log As Action(Of String)) As Integer
        If entityObj Is Nothing Then Return 0
        Dim cnt As Integer = 0
        Try
            cnt = Convert.ToInt32(CallByName(entityObj, "KeyPointCount", CallType.Get), CultureInfo.InvariantCulture)
            log("[DROP2D][GEOM][KEYPOINTS] entity=" & entityName & " count=" & cnt.ToString(CultureInfo.InvariantCulture) & " data=NO_CONFIRMADO")
            Return cnt
        Catch
            log("[DROP2D][GEOM][PROP_MISSING] entity=" & entityName & " prop=KeyPointCount")
            Return 0
        End Try
    End Function

    Private Shared Function LogConnectPoints(entityName As String, entityObj As Object, log As Action(Of String)) As Integer
        If entityObj Is Nothing Then Return 0
        Dim cp As Object = Nothing
        Try
            cp = CallByName(entityObj, "ConnectPoints", CallType.Get)
        Catch
            log("[DROP2D][GEOM][PROP_MISSING] entity=" & entityName & " prop=ConnectPoints")
            Return 0
        End Try
        Dim cnt As Integer = SafeCount(cp)
        log("[DROP2D][GEOM][CONNECTPOINTS] entity=" & entityName & " count=" & cnt.ToString(CultureInfo.InvariantCulture) & " data=NO_CONFIRMADO")
        Return cnt
    End Function

    Private Shared Function CaptureViewSnapshotBeforeDrop(dv As DrawingView, log As Action(Of String)) As Snapshot2D
        Dim snap As New Snapshot2D()
        If dv Is Nothing Then Return snap
        Try
            Dim lines As Object = GetCom(dv, "DVLines2d")
            Dim nLines As Integer = SafeCount(lines)
            For i As Integer = 1 To nLines
                Dim ln As Object = Nothing
                Try : ln = CallByName(lines, "Item", CallType.Method, i) : Catch : ln = Nothing : End Try
                If ln Is Nothing Then Continue For
                Dim vx1 As Double = 0, vy1 As Double = 0, vx2 As Double = 0, vy2 As Double = 0
                If Not TryExtractDvLineEndpoints(ln, vx1, vy1, vx2, vy2) Then Continue For
                Dim sx1 As Double = 0, sy1 As Double = 0, sx2 As Double = 0, sy2 As Double = 0
                Try
                    dv.ViewToSheet(vx1, vy1, sx1, sy1)
                    dv.ViewToSheet(vx2, vy2, sx2, sy2)
                Catch
                    Continue For
                End Try
                If IsDegenerateSegment(sx1, sy1, sx2, sy2) Then
                    Continue For
                End If
                snap.Lines.Add(Tuple.Create(sx1, sy1, sx2, sy2))
            Next

            Dim arcs As Object = GetCom(dv, "DVArcs2d")
            Dim nArcs As Integer = SafeCount(arcs)
            For i As Integer = 1 To nArcs
                Dim a As Object = Nothing
                Try : a = CallByName(arcs, "Item", CallType.Method, i) : Catch : a = Nothing : End Try
                If a Is Nothing Then Continue For
                Dim cx As Double = 0, cy As Double = 0, r As Double = 0, sa As Double = 0, ea As Double = 0, sw As Double = 0
                Try : CallByName(a, "GetCenterPoint", CallType.Method, cx, cy) : Catch : Continue For : End Try
                Try : r = Convert.ToDouble(CallByName(a, "Radius", CallType.Get), CultureInfo.InvariantCulture) : Catch : Continue For : End Try
                Try : sa = Convert.ToDouble(CallByName(a, "StartAngle", CallType.Get), CultureInfo.InvariantCulture) : Catch : sa = 0 : End Try
                Try : ea = Convert.ToDouble(CallByName(a, "EndAngle", CallType.Get), CultureInfo.InvariantCulture) : Catch : ea = sa + 180.0R : End Try
                Try : sw = Convert.ToDouble(CallByName(a, "SweepAngle", CallType.Get), CultureInfo.InvariantCulture) : Catch : sw = (ea - sa) : End Try
                Dim scx As Double = 0, scy As Double = 0
                Try : dv.ViewToSheet(cx, cy, scx, scy) : Catch : Continue For : End Try
                snap.Arcs.Add(Tuple.Create(scx, scy, r, sa, ea, sw))
            Next

            Dim circles As Object = GetCom(dv, "DVCircles2d")
            Dim nCircles As Integer = SafeCount(circles)
            For i As Integer = 1 To nCircles
                Dim c As Object = Nothing
                Try : c = CallByName(circles, "Item", CallType.Method, i) : Catch : c = Nothing : End Try
                If c Is Nothing Then Continue For
                Dim cx As Double = 0, cy As Double = 0, r As Double = 0
                Try : CallByName(c, "GetCenterPoint", CallType.Method, cx, cy) : Catch : Continue For : End Try
                Try : r = Convert.ToDouble(CallByName(c, "Radius", CallType.Get), CultureInfo.InvariantCulture) : Catch : Continue For : End Try
                Dim scx As Double = 0, scy As Double = 0
                Try : dv.ViewToSheet(cx, cy, scx, scy) : Catch : Continue For : End Try
                snap.Circles.Add(Tuple.Create(scx, scy, r))
            Next
        Catch ex As Exception
            log("[DROP2D][SNAPSHOT][FAIL] reason=" & ex.Message)
        End Try
        log("[DROP2D][SNAPSHOT][OK] lines=" & snap.Lines.Count.ToString(CultureInfo.InvariantCulture) &
            " arcs=" & snap.Arcs.Count.ToString(CultureInfo.InvariantCulture) &
            " circles=" & snap.Circles.Count.ToString(CultureInfo.InvariantCulture))
        Return snap
    End Function

    Private Shared Function CopySnapshotTo2DModel(snap As Snapshot2D, targetSheet As Sheet, viewIndex As Integer, viewName As String, layerGeom As Object, log As Action(Of String)) As Group2D
        If snap Is Nothing OrElse targetSheet Is Nothing Then Return Nothing
        log("[DROP2D][COPY_TO_2DMODEL][TRY] view=" & viewName)
        Dim g As New Group2D With {.ViewIndex = viewIndex, .ViewName = viewName}
        Dim xOffset As Double = 0.02R + ((viewIndex - 1) Mod 3) * 0.18R
        Dim yOffset As Double = 0.24R - ((viewIndex - 1) \ 3) * 0.12R

        Try
            For Each it In snap.Lines
                Dim sx1 As Double = it.Item1
                Dim sy1 As Double = it.Item2
                Dim sx2 As Double = it.Item3
                Dim sy2 As Double = it.Item4
                Dim l2 As Object = Nothing
                Try : l2 = CallByName(targetSheet.Lines2d, "AddBy2Points", CallType.Method, sx1 + xOffset, sy1 - yOffset, sx2 + xOffset, sy2 - yOffset) : Catch : l2 = Nothing : End Try
                If l2 Is Nothing Then Continue For
                TrySetLayer(l2, layerGeom)
                Dim lineObj As Line2d = Nothing
                Try
                    lineObj = CType(l2, Line2d)
                    g.Lines.Add(lineObj)
                Catch
                End Try
                g.LineMetas.Add(Tuple.Create(lineObj, sx1 + xOffset, sy1 - yOffset, sx2 + xOffset, sy2 - yOffset))
                UpdateGroupBounds(g, sx1 + xOffset, sy1 - yOffset)
                UpdateGroupBounds(g, sx2 + xOffset, sy2 - yOffset)
            Next

            For Each it In snap.Arcs
                Dim scx As Double = it.Item1
                Dim scy As Double = it.Item2
                Dim r As Double = it.Item3
                Dim sa As Double = it.Item4
                Dim ea As Double = it.Item5
                Dim sw As Double = it.Item6
                Dim arc2 As Object = TryAddArc2d(targetSheet, scx + xOffset, scy - yOffset, r, sa, ea, sw)
                If arc2 Is Nothing Then Continue For
                TrySetLayer(arc2, layerGeom)
                Try : g.Arcs.Add(CType(arc2, Arc2d)) : Catch : End Try
                UpdateGroupBounds(g, scx + xOffset - r, scy - yOffset - r)
                UpdateGroupBounds(g, scx + xOffset + r, scy - yOffset + r)
            Next

            For Each it In snap.Circles
                Dim scx As Double = it.Item1
                Dim scy As Double = it.Item2
                Dim r As Double = it.Item3
                Dim c2 As Object = Nothing
                Try : c2 = CallByName(targetSheet.Circles2d, "AddByCenterRadius", CallType.Method, scx + xOffset, scy - yOffset, r) : Catch : c2 = Nothing : End Try
                If c2 Is Nothing Then Continue For
                TrySetLayer(c2, layerGeom)
                Try : g.Circles.Add(CType(c2, Circle2d)) : Catch : End Try
                UpdateGroupBounds(g, scx + xOffset - r, scy - yOffset - r)
                UpdateGroupBounds(g, scx + xOffset + r, scy - yOffset + r)
            Next

            log("[DROP2D][COPY_TO_2DMODEL][OK] copied_lines=" & g.Lines.Count.ToString(CultureInfo.InvariantCulture) &
                " arcs=" & g.Arcs.Count.ToString(CultureInfo.InvariantCulture) &
                " circles=" & g.Circles.Count.ToString(CultureInfo.InvariantCulture))
            Return g
        Catch ex As Exception
            log("[DROP2D][COPY_TO_2DMODEL][FAIL] reason=" & ex.Message)
            Return Nothing
        End Try
    End Function

    Private Shared Sub UpdateGroupBounds(g As Group2D, x As Double, y As Double)
        If g Is Nothing Then Return
        If x < g.MinX Then g.MinX = x
        If x > g.MaxX Then g.MaxX = x
        If y < g.MinY Then g.MinY = y
        If y > g.MaxY Then g.MaxY = y
    End Sub

    Private Shared Function TryAddArc2d(targetSheet As Sheet, cx As Double, cy As Double, r As Double, startA As Double, endA As Double, sweepA As Double) As Object
        If targetSheet Is Nothing Then Return Nothing
        Dim arc2 As Object = Nothing
        Dim saRad As Double = ToRadiansIfNeeded(startA)
        Dim eaRad As Double = ToRadiansIfNeeded(endA)
        Dim sx As Double = cx + r * Math.Cos(saRad)
        Dim sy As Double = cy + r * Math.Sin(saRad)
        Dim ex As Double = cx + r * Math.Cos(eaRad)
        Dim ey As Double = cy + r * Math.Sin(eaRad)

        ' Firma más habitual en Arc2d: centro + punto inicio + punto fin.
        Try : arc2 = CallByName(targetSheet.Arcs2d, "AddByCenterStartEnd", CallType.Method, cx, cy, sx, sy, ex, ey) : Catch : arc2 = Nothing : End Try
        If arc2 IsNot Nothing Then Return arc2

        Dim sweep As Double = If(Math.Abs(sweepA) > 1.0E-9, sweepA, endA - startA)
        Dim sweepRad As Double = ToRadiansIfNeeded(sweep)
        ' Fallback alternativo observado en algunas versiones COM.
        Try : arc2 = CallByName(targetSheet.Arcs2d, "AddByCenterStartSweep", CallType.Method, cx, cy, sx, sy, sweepRad) : Catch : arc2 = Nothing : End Try
        If arc2 IsNot Nothing Then Return arc2

        ' Último fallback por si la variante espera radio + ángulos.
        Try : arc2 = CallByName(targetSheet.Arcs2d, "AddByCenterStartSweep", CallType.Method, cx, cy, r, startA, sweep) : Catch : arc2 = Nothing : End Try
        Return arc2
    End Function

    Private Shared Function TryExtractDvLineEndpoints(dvLine As Object, ByRef x1 As Double, ByRef y1 As Double, ByRef x2 As Double, ByRef y2 As Double) As Boolean
        x1 = 0 : y1 = 0 : x2 = 0 : y2 = 0
        If dvLine Is Nothing Then Return False

        If TryExtractDvLineEndpointsReflect(dvLine, x1, y1, x2, y2) Then
            Return True
        End If

        Try
            CallByName(dvLine, "GetEndPoints", CallType.Method, x1, y1, x2, y2)
            If Not IsDegenerateSegment(x1, y1, x2, y2) Then Return True
        Catch
        End Try

        Try
            x1 = Convert.ToDouble(CallByName(dvLine, "StartPointX", CallType.Get), CultureInfo.InvariantCulture)
            y1 = Convert.ToDouble(CallByName(dvLine, "StartPointY", CallType.Get), CultureInfo.InvariantCulture)
            x2 = Convert.ToDouble(CallByName(dvLine, "EndPointX", CallType.Get), CultureInfo.InvariantCulture)
            y2 = Convert.ToDouble(CallByName(dvLine, "EndPointY", CallType.Get), CultureInfo.InvariantCulture)
            If Not IsDegenerateSegment(x1, y1, x2, y2) Then Return True
        Catch
        End Try

        Try
            Dim p1 As Object = CallByName(dvLine, "StartPoint", CallType.Get)
            Dim p2 As Object = CallByName(dvLine, "EndPoint", CallType.Get)
            x1 = Convert.ToDouble(CallByName(p1, "X", CallType.Get), CultureInfo.InvariantCulture)
            y1 = Convert.ToDouble(CallByName(p1, "Y", CallType.Get), CultureInfo.InvariantCulture)
            x2 = Convert.ToDouble(CallByName(p2, "X", CallType.Get), CultureInfo.InvariantCulture)
            y2 = Convert.ToDouble(CallByName(p2, "Y", CallType.Get), CultureInfo.InvariantCulture)
            If Not IsDegenerateSegment(x1, y1, x2, y2) Then Return True
        Catch
        End Try

        Try
            Dim sX As Double = 0, sY As Double = 0, eX As Double = 0, eY As Double = 0
            CallByName(dvLine, "GetStartPoint", CallType.Method, sX, sY)
            CallByName(dvLine, "GetEndPoint", CallType.Method, eX, eY)
            x1 = sX : y1 = sY : x2 = eX : y2 = eY
            If Not IsDegenerateSegment(x1, y1, x2, y2) Then Return True
        Catch
        End Try

        Try
            CallByName(dvLine, "Range", CallType.Method, x1, y1, x2, y2)
            If Not IsDegenerateSegment(x1, y1, x2, y2) Then Return True
        Catch
        End Try

        Return False
    End Function

    ''' <summary>
    ''' GetStartPoint/GetEndPoint de DVLine2d usan ByRef; CallByName no pobló bien esos valores (quedaban 0,0).
    ''' Misma estrategia que DrawingViewDimensioningLab (InvokeMember).
    ''' </summary>
    Private Shared Function TryExtractDvLineEndpointsReflect(dvLine As Object, ByRef x1 As Double, ByRef y1 As Double, ByRef x2 As Double, ByRef y2 As Double) As Boolean
        x1 = 0 : y1 = 0 : x2 = 0 : y2 = 0
        Try
            Dim argsStart As Object() = New Object() {0.0R, 0.0R}
            dvLine.GetType().InvokeMember("GetStartPoint", DvInvokeFlags, Nothing, dvLine, argsStart)
            x1 = CDbl(argsStart(0))
            y1 = CDbl(argsStart(1))

            Dim argsEnd As Object() = New Object() {0.0R, 0.0R}
            dvLine.GetType().InvokeMember("GetEndPoint", DvInvokeFlags, Nothing, dvLine, argsEnd)
            x2 = CDbl(argsEnd(0))
            y2 = CDbl(argsEnd(1))

            Dim dx As Double = x2 - x1
            Dim dy As Double = y2 - y1
            Dim endpointLen As Double = Math.Sqrt(dx * dx + dy * dy)

            Dim dvLen As Double = 0
            Try : dvLen = CDbl(CallByName(dvLine, "Length", CallType.Get)) : Catch : dvLen = 0 : End Try

            If dvLen > 0.000001R AndAlso endpointLen < 0.000001R Then
                Return False
            End If

            Return Not IsDegenerateSegment(x1, y1, x2, y2)
        Catch
            Return False
        End Try
    End Function

    Private Shared Function IsDegenerateSegment(x1 As Double, y1 As Double, x2 As Double, y2 As Double) As Boolean
        Return Math.Abs(x2 - x1) <= 1.0E-9 AndAlso Math.Abs(y2 - y1) <= 1.0E-9
    End Function

    Private Shared Function ToRadiansIfNeeded(angle As Double) As Double
        If Math.Abs(angle) > (Math.PI * 2.0R + 0.000001R) Then
            Return angle * (Math.PI / 180.0R)
        End If
        Return angle
    End Function

    Private Shared Function TryDropView(dv As DrawingView, ByRef methodUsed As String, log As Action(Of String)) As Boolean
        methodUsed = ""
        If dv Is Nothing Then Return False
        Dim names As String() = {"Drop", "DropView", "Break", "ConvertTo2D", "ConvertToGeometry"}
        For Each m In names
            Try
                CallByName(dv, m, CallType.Method)
                methodUsed = m
                Return True
            Catch ex As Exception
                log("[DROP2D][VIEW][DROP_FAIL] reason=method_" & m & "_not_available_or_failed msg=" & ex.Message & " note=NO_CONFIRMADO")
            End Try
        Next
        Return False
    End Function

    Private Shared Sub ResolveSheetsForLab(draft As DraftDocument, ByRef sourceSheet As Sheet, ByRef target2D As Sheet, log As Action(Of String))
        sourceSheet = Nothing
        target2D = Nothing
        If draft Is Nothing Then Return

        Dim n As Integer = 0
        Try : n = draft.Sheets.Count : Catch : n = 0 : End Try
        Dim hoja1Candidate As Sheet = Nothing
        Dim bestByViews As Sheet = Nothing
        Dim bestByViewsCount As Integer = -1
        For i As Integer = 1 To n
            Dim sh As Sheet = Nothing
            Try : sh = CType(draft.Sheets.Item(i), Sheet) : Catch : sh = Nothing : End Try
            If sh Is Nothing Then Continue For
            Dim nm As String = SafeName(sh)
            If String.Equals(nm, "2D Model", StringComparison.OrdinalIgnoreCase) Then
                target2D = sh
            Else
                Dim vCount As Integer = 0
                Try : vCount = sh.DrawingViews.Count : Catch : vCount = 0 : End Try
                If String.Equals(nm, "Hoja1", StringComparison.OrdinalIgnoreCase) Then
                    hoja1Candidate = sh
                End If
                If vCount > bestByViewsCount Then
                    bestByViews = sh
                    bestByViewsCount = vCount
                End If
            End If
        Next

        If hoja1Candidate IsNot Nothing Then
            Dim hoja1Views As Integer = 0
            Try : hoja1Views = hoja1Candidate.DrawingViews.Count : Catch : hoja1Views = 0 : End Try
            If hoja1Views > 0 Then
                sourceSheet = hoja1Candidate
            End If
        End If

        If sourceSheet Is Nothing AndAlso bestByViews IsNot Nothing AndAlso bestByViewsCount > 0 Then
            sourceSheet = bestByViews
        End If

        If target2D Is Nothing Then
            Try
                target2D = CType(CallByName(draft.Sheets, "Add", CallType.Method, "2D Model"), Sheet)
            Catch
                Try
                    target2D = CType(draft.ActiveSheet, Sheet)
                Catch
                    target2D = Nothing
                End Try
            End Try
        End If

        If sourceSheet Is Nothing Then
            Try : sourceSheet = CType(draft.ActiveSheet, Sheet) : Catch : sourceSheet = Nothing : End Try
        End If

        If sourceSheet IsNot Nothing Then
            Dim srcViews As Integer = 0
            Try : srcViews = sourceSheet.DrawingViews.Count : Catch : srcViews = 0 : End Try
            log("[DROP2D][SOURCE_SHEET][RESOLVE] name=" & SafeName(sourceSheet) & " views=" & srcViews.ToString(CultureInfo.InvariantCulture))
        End If
    End Sub

    Private Shared Function CleanupPreviousLabGeometry(sh As Sheet, log As Action(Of String)) As Integer
        If sh Is Nothing Then Return 0
        Dim deleted As Integer = 0
        deleted += DeleteByLayer(sh, "Lines2d", LayerGeometry)
        deleted += DeleteByLayer(sh, "Arcs2d", LayerGeometry)
        deleted += DeleteByLayer(sh, "Circles2d", LayerGeometry)
        deleted += DeleteByLayer(sh, "LineStrings2d", LayerGeometry)
        deleted += DeleteByLayer(sh, "BSplineCurves2d", LayerGeometry)
        deleted += DeleteByLayer(sh, "Points2d", LayerGeometry)
        deleted += DeleteByLayer(sh, "Dimensions", LayerDimensions)
        Return deleted
    End Function

    Private Shared Function DeleteByLayer(sh As Sheet, colName As String, layerName As String) As Integer
        Dim col As Object = GetCom(sh, colName)
        Dim n As Integer = SafeCount(col)
        Dim del As Integer = 0
        For i As Integer = n To 1 Step -1
            Dim it As Object = Nothing
            Try : it = CallByName(col, "Item", CallType.Method, i) : Catch : it = Nothing : End Try
            If it Is Nothing Then Continue For
            Dim lname As String = ""
            Try : lname = SafeStr(CallByName(CallByName(it, "Layer", CallType.Get), "Name", CallType.Get)) : Catch : lname = "" : End Try
            If Not String.Equals(lname, layerName, StringComparison.OrdinalIgnoreCase) Then Continue For
            Try : CallByName(it, "Delete", CallType.Method) : del += 1 : Catch : End Try
        Next
        Return del
    End Function

    Private Shared Function EnsureLayer(sheet As Sheet, layerName As String, log As Action(Of String)) As Object
        If sheet Is Nothing Then Return Nothing
        Dim layers As Object = GetCom(sheet, "Layers")
        If layers Is Nothing Then Return Nothing
        Dim n As Integer = SafeCount(layers)
        For i As Integer = 1 To n
            Dim ly As Object = Nothing
            Try : ly = CallByName(layers, "Item", CallType.Method, i) : Catch : ly = Nothing : End Try
            If ly Is Nothing Then Continue For
            Dim nm As String = SafeStr(CallByNameSafe(ly, "Name"))
            If String.Equals(nm, layerName, StringComparison.OrdinalIgnoreCase) Then Return ly
        Next
        Try
            Dim created = CallByName(layers, "Add", CallType.Method, layerName)
            Return created
        Catch ex As Exception
            log("[DROP2D][LAYER][WARN] create_fail name=" & layerName & " msg=" & ex.Message)
            Return Nothing
        End Try
    End Function

    Private Shared Sub TrySetLayer(obj As Object, layerObj As Object)
        If obj Is Nothing OrElse layerObj Is Nothing Then Return
        Try : CallByName(obj, "Layer", CallType.Let, layerObj) : Catch : End Try
    End Sub

    Private Shared Function GetCom(obj As Object, member As String) As Object
        Try
            Return CallByName(obj, member, CallType.Get)
        Catch
            Return Nothing
        End Try
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

    Private Shared Function SafeName(sh As Sheet) As String
        If sh Is Nothing Then Return ""
        Try
            Return Convert.ToString(sh.Name, CultureInfo.InvariantCulture)
        Catch
            Return ""
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

    Private Shared Function FormatInv(v As Double) As String
        Return v.ToString("0.######", CultureInfo.InvariantCulture)
    End Function
End Class

