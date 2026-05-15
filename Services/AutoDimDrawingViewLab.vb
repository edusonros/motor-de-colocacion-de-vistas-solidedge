Option Strict Off

Imports System.Globalization
Imports System.Math
Imports System.Linq

''' <summary>
''' Laboratorio aislado: auditoría geometría DV sobre <c>DrawingView</c> y creación experimental de cotas
''' con <c>Sheet.Dimensions.AddDistanceBetweenObjects</c> conforme a la investigación SDK en
''' <c>docs\investigacion_acotacion_drawingview.md</c>.
''' </summary>
Public NotInheritable Class AutoDimDrawingViewLab

    ''' <remarks>
    ''' Valor <c>2</c> = <c>igIsometricView</c> según SDK HTML local
    ''' (<c>SolidEdgeDraft~DrawingViewTypeConstants.html</c>,
    ''' <c>SolidEdgeDraft~DrawingView~DrawingViewType.html</c>).
    ''' </remarks>
    Private Const IgIsometricView As Integer = 2

    Private Shared ReadOnly AngleTolDeg As Double = 2.0R

    Private Enum CreateDimOutcome As Integer
        Failed = 0
        Connected = 1
        FloatingVisible = 2
        BadKept = 3
    End Enum

    Private Structure ViewRunStats
        Public DimsCreated As Integer
        Public Connected As Integer
        Public Floating As Integer
        Public Failed As Integer
    End Structure

    Public Class InterestingPoint
        Public Property SourceType As String
        Public Property SourceIndex As Integer
        Public Property Role As String
        Public Property XView As Double
        Public Property YView As Double
        Public Property XSheet As Double
        Public Property YSheet As Double
        Public Property Confidence As Double
    End Class

    Private Class LineSample
        Public Idx As Integer
        Public LineObj As Object
        Public Sx As Double, Sy As Double
        Public Ex As Double, Ey As Double
        Public Len As Double
        Public AngleDeg As Double
        Public Orient As String
        Public Mx As Double
        Public My As Double
    End Class

    ''' <summary>Entrada recomendada: documento borrador abierto (.dft COM).</summary>
    Public Shared Sub RunDrawingViewDimensionLab(activeDraft As Object, Optional logger As Logger = Nothing)
        If activeDraft Is Nothing Then
            Lg(logger, "[DVDIM][FAIL] activeDraft null")
            Return
        End If
        Dim sheet As Object = Nothing
        Try
            sheet = CallByName(activeDraft, "ActiveSheet", CallType.Get)
        Catch ex As Exception
            Lg(logger, "[DVDIM][FAIL] ActiveSheet " & ex.Message)
            Return
        End Try
        RunOnSheet(sheet, logger)
    End Sub

    ''' <summary>Hoja ya resuelta (p. ej. <c>DraftDocument.ActiveSheet</c>).</summary>
    Public Shared Sub RunOnSheet(activeSheet As Object, Optional logger As Logger = Nothing)
        Lg(logger, "[DVDIM][ENTER]")
        If activeSheet Is Nothing Then
            Lg(logger, "[DVDIM][FAIL] sheet null")
            Return
        End If

        Dim sheetName As String = SafeStr(TryLateGet(activeSheet, "Name"))
        Lg(logger, "[DVDIM][SHEET] name=" & sheetName)

        Dim dimsColl As Object = Nothing
        Dim dimsBefore As Integer = 0
        Try
            dimsColl = CallByName(activeSheet, "Dimensions", CallType.Get)
            dimsBefore = SafeInt(CallByName(dimsColl, "Count", CallType.Get), 0)
        Catch ex As Exception
            Lg(logger, "[DVDIM][FAIL] Dimensions " & ex.Message)
            Return
        End Try

        Dim views As Object = Nothing
        Dim nViews As Integer = 0
        Try
            views = CallByName(activeSheet, "DrawingViews", CallType.Get)
            nViews = SafeInt(CallByName(views, "Count", CallType.Get), 0)
        Catch ex As Exception
            Lg(logger, "[DVDIM][FAIL] DrawingViews " & ex.Message)
            Return
        End Try

        Dim viewsProcessed As Integer = 0
        Dim totalCreated As Integer = 0
        Dim docConnected As Integer = 0
        Dim docFloating As Integer = 0
        Dim docFailed As Integer = 0

        For vi As Integer = 1 To nViews
            Dim vObj As Object = GetCollItem(views, vi)
            If vObj Is Nothing Then Continue For
            Dim vs As ViewRunStats = ProcessOneDrawingView(vObj, dimsColl, vi, logger)
            totalCreated += vs.DimsCreated
            docConnected += vs.Connected
            docFloating += vs.Floating
            docFailed += vs.Failed
            viewsProcessed += 1
        Next

        Dim dimsAfter As Integer = SafeInt(CallByName(dimsColl, "Count", CallType.Get), 0)
        Lg(logger, "[DVDIM][SUMMARY][DOC] viewsProcessed=" & viewsProcessed.ToString(CultureInfo.InvariantCulture) &
            " dimsBefore=" & dimsBefore.ToString(CultureInfo.InvariantCulture) &
            " dimsAfter=" & dimsAfter.ToString(CultureInfo.InvariantCulture) &
            " delta=" & (dimsAfter - dimsBefore).ToString(CultureInfo.InvariantCulture) &
            " dimsCreatedExperiment=" & totalCreated.ToString(CultureInfo.InvariantCulture) &
            " connected=" & docConnected.ToString(CultureInfo.InvariantCulture) &
            " floating=" & docFloating.ToString(CultureInfo.InvariantCulture) &
            " failed=" & docFailed.ToString(CultureInfo.InvariantCulture))
        Lg(logger, "[DVDIM][DONE]")
    End Sub

    Private Shared Function ProcessOneDrawingView(dv As Object, dimsColl As Object,
                                                  viewOrdinal As Integer, logger As Logger) As ViewRunStats
        Dim stats As New ViewRunStats()

        Dim vName As String = SafeStr(TryLateGet(dv, "Name"))
        Dim typeLong As Integer = SafeInt(TryLateGetNumber(dv, "Type"), -1)
        Dim dvt As Integer = SafeInt(TryLateGetNumber(dv, "DrawingViewType"), -1)
        Dim scaleRaw As Object = Nothing
        Try : scaleRaw = CallByName(dv, "ScaleFactor", CallType.Get) : Catch : End Try
        Dim ox As Double = 0, oy As Double = 0
        Try : CallByName(dv, "GetOrigin", CallType.Method, ox, oy) : Catch : ox = Double.NaN : oy = Double.NaN : End Try

        Lg(logger, "[DVDIM][VIEW] idx=" & viewOrdinal.ToString(CultureInfo.InvariantCulture) &
            " name=" & vName &
            " DrawingViewType=" & dvt.ToString(CultureInfo.InvariantCulture) &
            " Type_long=" & typeLong.ToString(CultureInfo.InvariantCulture) &
            " scale=" & F(scaleRaw) &
            " origin_x=" & F(ox) & " origin_y=" & F(oy))

        If dvt = IgIsometricView Then
            Lg(logger, "[DVDIM][VIEW][SKIP_ISO] DrawingViewType=igIsometricView (2) según SolidEdgeDraft~DrawingViewTypeConstants.html")
            SummarizeSkippedView(viewOrdinal, logger)
            Return stats
        End If

        Dim samples As List(Of LineSample) = CollectDvLines(dv, viewOrdinal, logger)
        Dim nArcs As Integer = CountDvCollection(dv, "DVArcs2d")
        Dim nCirc As Integer = CountDvCollection(dv, "DVCircles2d")
        Dim nPt As Integer = CountDvCollection(dv, "DVPoints2d")

        Dim bbox As Tuple(Of Double, Double, Double, Double) = CalcBBox(samples)
        If bbox IsNot Nothing Then
            Lg(logger, "[DVDIM][VIEW][RANGE] minX=" & F(bbox.Item1) & " minY=" & F(bbox.Item2) &
                " maxX=" & F(bbox.Item3) & " maxY=" & F(bbox.Item4))
        Else
            Lg(logger, "[DVDIM][VIEW][RANGE] vacío")
        End If

        Lg(logger, "[DVDIM][GEOM][DVLINES] count=" & samples.Count.ToString(CultureInfo.InvariantCulture))
        Lg(logger, "[DVDIM][GEOM][DVARCS] count=" & nArcs.ToString(CultureInfo.InvariantCulture))
        Lg(logger, "[DVDIM][GEOM][DVCIRCLES] count=" & nCirc.ToString(CultureInfo.InvariantCulture))
        Lg(logger, "[DVDIM][GEOM][DVPOINTS] count=" & nPt.ToString(CultureInfo.InvariantCulture))
        LogDvCirclesBrief(dv, viewOrdinal, logger)
        LogDvPointsBrief(dv, viewOrdinal, logger)

        Dim verts As List(Of LineSample) = samples.Where(Function(x) String.Equals(x.Orient, "V", StringComparison.OrdinalIgnoreCase)).ToList()
        Dim hors As List(Of LineSample) = samples.Where(Function(x) String.Equals(x.Orient, "H", StringComparison.OrdinalIgnoreCase)).ToList()

        Dim interesting As New List(Of InterestingPoint)()
        EmitInterestingPoints(logger, dv, viewOrdinal, samples, verts, hors, bbox, interesting)

        Dim hCand As Tuple(Of LineSample, LineSample, Double, Double) = PickVerticalSpan(verts, samples, bbox, dv, logger, viewOrdinal, "H_TOTAL")
        Dim vCand As Tuple(Of LineSample, LineSample, Double, Double) = PickHorizontalSpan(hors, samples, bbox, dv, logger, viewOrdinal, "V_TOTAL")

        If hCand.Item1 IsNot Nothing AndAlso hCand.Item2 IsNot Nothing Then
            AccumulateOutcome(stats, TryCreateOuterDistance(dimsColl, dv, hCand.Item1, hCand.Item2, logger, viewOrdinal,
                                     "H_TOTAL", hCand.Item3))
        Else
            Lg(logger, "[DVDIM][CAND][H_TOTAL][FAIL] view=" & viewOrdinal.ToString(CultureInfo.InvariantCulture) &
                " Falta par de objetos VERT con extremos suficientemente separados.")
            stats.Failed += 1
        End If

        If vCand.Item1 IsNot Nothing AndAlso vCand.Item2 IsNot Nothing Then
            AccumulateOutcome(stats, TryCreateOuterDistance(dimsColl, dv, vCand.Item1, vCand.Item2, logger, viewOrdinal,
                                     "V_TOTAL", vCand.Item3))
        Else
            Lg(logger, "[DVDIM][CAND][V_TOTAL][FAIL] view=" & viewOrdinal.ToString(CultureInfo.InvariantCulture) &
                " Falta par de objetos HORIZ con extremos suficientemente separados.")
            stats.Failed += 1
        End If

        Lg(logger, "[DVDIM][SUMMARY][VIEW] view=" & viewOrdinal.ToString(CultureInfo.InvariantCulture) &
            " lines=" & samples.Count.ToString(CultureInfo.InvariantCulture) &
            " points=" & interesting.Count.ToString(CultureInfo.InvariantCulture) &
            " hCandidates=" & If(hCand.Item1 IsNot Nothing AndAlso hCand.Item2 IsNot Nothing, "1", "0") &
            " vCandidates=" & If(vCand.Item1 IsNot Nothing AndAlso vCand.Item2 IsNot Nothing, "1", "0") &
            " dimsCreated=" & stats.DimsCreated.ToString(CultureInfo.InvariantCulture) &
            " connected=" & stats.Connected.ToString(CultureInfo.InvariantCulture) &
            " floating=" & stats.Floating.ToString(CultureInfo.InvariantCulture) &
            " failed=" & stats.Failed.ToString(CultureInfo.InvariantCulture))

        Return stats
    End Function

    Private Shared Sub AccumulateOutcome(ByRef stats As ViewRunStats, outcome As CreateDimOutcome)
        Select Case outcome
            Case CreateDimOutcome.Connected
                stats.DimsCreated += 1
                stats.Connected += 1
            Case CreateDimOutcome.FloatingVisible, CreateDimOutcome.BadKept
                stats.DimsCreated += 1
                stats.Floating += 1
            Case CreateDimOutcome.Failed
                stats.Failed += 1
        End Select
    End Sub

    Private Shared Function TryLateGetNumber(o As Object, prop As String) As Object
        If o Is Nothing Then Return Nothing
        Try
            Return CallByName(o, prop, CallType.Get)
        Catch
            Return Nothing
        End Try
    End Function

    Private Shared Sub LogDvCirclesBrief(dv As Object, viewOrdinal As Integer, logger As Logger)
        Dim coll As Object = Nothing
        Try
            coll = CallByName(dv, "DVCircles2d", CallType.Get)
        Catch
            Return
        End Try
        Dim n As Integer = SafeInt(CallByName(coll, "Count", CallType.Get), 0)
        Const maxLog As Integer = 12
        For i As Integer = 1 To Math.Min(n, maxLog)
            Dim c As Object = GetCollItem(coll, i)
            If c Is Nothing Then Continue For
            Dim cx As Double = 0, cy As Double = 0
            Dim rad As Object = Nothing
            Try
                CallByName(c, "GetCenterPoint", CallType.Method, cx, cy)
            Catch ex As Exception
                Lg(logger, "[DVDIM][CIRCLE][ERR] view=" & viewOrdinal.ToString(CultureInfo.InvariantCulture) &
                    " idx=" & i.ToString(CultureInfo.InvariantCulture) & " " & ex.Message)
                Continue For
            End Try
            Try : rad = CallByName(c, "Radius", CallType.Get) : Catch : End Try
            Lg(logger, "[DVDIM][CIRCLE] view=" & viewOrdinal.ToString(CultureInfo.InvariantCulture) &
                " idx=" & i.ToString(CultureInfo.InvariantCulture) &
                " cx_view=" & F(cx) & " cy_view=" & F(cy) & " radius=" & F(rad))
        Next
        If n > maxLog Then
            Lg(logger, "[DVDIM][CIRCLE] view=" & viewOrdinal.ToString(CultureInfo.InvariantCulture) &
                " ... truncated logged=" & maxLog.ToString(CultureInfo.InvariantCulture) & " of " & n.ToString(CultureInfo.InvariantCulture))
        End If
    End Sub

    Private Shared Sub LogDvPointsBrief(dv As Object, viewOrdinal As Integer, logger As Logger)
        Dim coll As Object = Nothing
        Try
            coll = CallByName(dv, "DVPoints2d", CallType.Get)
        Catch
            Return
        End Try
        Dim n As Integer = SafeInt(CallByName(coll, "Count", CallType.Get), 0)
        Const maxLog As Integer = 12
        For i As Integer = 1 To Math.Min(n, maxLog)
            Dim p As Object = GetCollItem(coll, i)
            If p Is Nothing Then Continue For
            Dim px As Double = 0, py As Double = 0
            Try
                CallByName(p, "GetPoint", CallType.Method, px, py)
            Catch
                Try
                    px = SafeDbl(CallByName(p, "x", CallType.Get), 0)
                    py = SafeDbl(CallByName(p, "y", CallType.Get), 0)
                Catch ex As Exception
                    Lg(logger, "[DVDIM][POINT][ERR] view=" & viewOrdinal.ToString(CultureInfo.InvariantCulture) &
                        " idx=" & i.ToString(CultureInfo.InvariantCulture) & " " & ex.Message)
                    Continue For
                End Try
            End Try
            Lg(logger, "[DVDIM][POINT] view=" & viewOrdinal.ToString(CultureInfo.InvariantCulture) &
                " idx=" & i.ToString(CultureInfo.InvariantCulture) &
                " x_view=" & F(px) & " y_view=" & F(py))
        Next
    End Sub

    Private Shared Sub SummarizeSkippedView(viewOrdinal As Integer, logger As Logger)
        Lg(logger, "[DVDIM][SUMMARY][VIEW] view=" & viewOrdinal.ToString(CultureInfo.InvariantCulture) &
            " skipped=isometric dimsCreated=0")
    End Sub

    Private Shared Function CollectDvLines(dv As Object, viewOrdinal As Integer, logger As Logger) As List(Of LineSample)
        Dim out As New List(Of LineSample)()
        Dim coll As Object = Nothing
        Try
            coll = CallByName(dv, "DVLines2d", CallType.Get)
        Catch
            Return out
        End Try
        Dim n As Integer = SafeInt(CallByName(coll, "Count", CallType.Get), 0)

        For i As Integer = 1 To n
            Dim ln As Object = GetCollItem(coll, i)
            If ln Is Nothing Then Continue For
            Dim sx As Double = 0, sy As Double = 0, exCoord As Double = 0, eyCoord As Double = 0
            Try
                CallByName(ln, "GetStartPoint", CallType.Method, sx, sy)
                CallByName(ln, "GetEndPoint", CallType.Method, exCoord, eyCoord)
            Catch exPts As Exception
                Lg(logger, "[DVDIM][LINE][ERR] view=" & viewOrdinal.ToString(CultureInfo.InvariantCulture) &
                    " idx=" & i.ToString(CultureInfo.InvariantCulture) & " " & exPts.Message)
                Continue For
            End Try

            Dim dx As Double = exCoord - sx
            Dim dy As Double = eyCoord - sy
            Dim len As Double = Sqrt(dx * dx + dy * dy)
            Dim angle As Double = 0
            Const MinLenGeom As Double = 1.0E-09
            If len >= MinLenGeom Then angle = Atan2(dy, dx) * 180.0R / PI
            Dim orient As String = ClassifyOrient(dx, dy, len, angle)

            Dim sNew As New LineSample With {
                .Idx = i,
                .LineObj = ln,
                .Sx = sx, .Sy = sy, .Ex = exCoord, .Ey = eyCoord,
                .Len = len,
                .AngleDeg = angle,
                .Orient = orient,
                .Mx = (sx + exCoord) / 2.0R,
                .My = (sy + eyCoord) / 2.0R}

            Dim xs1 As Double = 0, ys1 As Double = 0, xs2 As Double = 0, ys2 As Double = 0
            Dim okSheet As Boolean = TryViewToSheetBatch(dv, logger, sx, sy, xs1, ys1) AndAlso
                                   TryViewToSheetBatch(dv, logger, exCoord, eyCoord, xs2, ys2)

            Lg(logger, "[DVDIM][LINE] view=" & viewOrdinal.ToString(CultureInfo.InvariantCulture) &
                " idx=" & i.ToString(CultureInfo.InvariantCulture) &
                " sx_view=" & F(sx) & " sy_view=" & F(sy) & " ex_view=" & F(exCoord) & " ey_view=" & F(eyCoord) &
                " len=" & F(len) & " angle=" & F(angle) & " orient=" & orient)
            If okSheet Then
                Lg(logger, "[DVDIM][LINE][SHEET] view=" & viewOrdinal.ToString(CultureInfo.InvariantCulture) &
                    " idx=" & i.ToString(CultureInfo.InvariantCulture) &
                    " sx_sheet=" & F(xs1) & " sy_sheet=" & F(ys1) &
                    " ex_sheet=" & F(xs2) & " ey_sheet=" & F(ys2))
            End If

            out.Add(sNew)
        Next
        Return out
    End Function

    Private Shared Function TryViewToSheetBatch(dv As Object, logger As Logger,
                                                xv As Double, yv As Double,
                                                ByRef xs As Double, ByRef ys As Double) As Boolean
        Try
            CallByName(dv, "ViewToSheet", CallType.Method, xv, yv, xs, ys)
            Return True
        Catch ex As Exception
            Lg(logger, "[DVDIM][COORD][FAIL] method=ViewToSheet line=combinación(xView,yView) error=" & ex.Message)
            Return False
        End Try
    End Function

    Private Shared Function CalcBBox(samples As IEnumerable(Of LineSample)) As Tuple(Of Double, Double, Double, Double)
        Dim first As Boolean = True
        Dim minX As Double = 0, minY As Double = 0, maxX As Double = 0, maxY As Double = 0
        For Each s In samples
            Dim candidates As Double()() = {
                New Double() {s.Sx, s.Sy},
                New Double() {s.Ex, s.Ey}}
            For Each p In candidates
                If first Then
                    minX = p(0) : maxX = p(0) : minY = p(1) : maxY = p(1) : first = False
                Else
                    minX = Min(minX, p(0)) : maxX = Max(maxX, p(0)) : minY = Min(minY, p(1)) : maxY = Max(maxY, p(1))
                End If
            Next
        Next
        If first Then Return Nothing
        Return Tuple.Create(minX, minY, maxX, maxY)
    End Function

    Private Shared Function ClassifyOrient(dx As Double, dy As Double, len As Double, angleDeg As Double) As String
        Const MinLenGeom As Double = 1.0E-09
        If len < MinLenGeom Then Return "OTHER"
        Dim aa As Double = Abs(angleDeg)
        Dim dh As Double = Min(aa, Abs(aa - 180.0R))
        If dh <= AngleTolDeg Then Return "H"
        Dim dv As Double = Min(Abs(aa - 90.0R), Abs(aa + 90.0R))
        If dv <= AngleTolDeg Then Return "V"
        Return "OTHER"
    End Function

    Private Shared Function CountDvCollection(dv As Object, propName As String) As Integer
        Try
            Dim c As Object = CallByName(dv, propName, CallType.Get)
            If c Is Nothing Then Return 0
            Return SafeInt(CallByName(c, "Count", CallType.Get), 0)
        Catch
            Return 0
        End Try
    End Function

    Private Shared Sub EmitInterestingPoints(logger As Logger, dv As Object, viewIdx As Integer,
                                             allLines As List(Of LineSample),
                                             vertLines As List(Of LineSample),
                                             horzLines As List(Of LineSample),
                                             bbox As Tuple(Of Double, Double, Double, Double),
                                             sink As List(Of InterestingPoint))
        Dim add As Action(Of String, Double, Double, Double) =
            Sub(role As String, xv As Double, yv As Double, conf As Double)
                Dim ip As New InterestingPoint With {
                    .SourceType = "Synthetic",
                    .SourceIndex = 0,
                    .Role = role,
                    .XView = xv, .YView = yv, .Confidence = conf,
                    .XSheet = xv, .YSheet = yv}
                Dim xs As Double = 0, ys As Double = 0
                If TryViewToSheetBatch(dv, logger, xv, yv, xs, ys) Then
                    ip.XSheet = xs : ip.YSheet = ys
                End If
                sink.Add(ip)
                Lg(logger, "[DVDIM][POINT][INTERESTING] view=" & viewIdx.ToString(CultureInfo.InvariantCulture) &
                    " role=" & role & " source=BBox idx=-1 xView=" & F(xv) & " yView=" & F(yv) &
                    " xSheet=" & F(ip.XSheet) & " ySheet=" & F(ip.YSheet))
            End Sub

        If bbox IsNot Nothing Then
            Dim minX As Double = bbox.Item1, minY As Double = bbox.Item2, maxX As Double = bbox.Item3, maxY As Double = bbox.Item4
            add("LeftExtreme", minX, (minY + maxY) / 2.0R, 0.7)
            add("RightExtreme", maxX, (minY + maxY) / 2.0R, 0.7)
            add("BottomExtreme", (minX + maxX) / 2.0R, minY, 0.7)
            add("TopExtreme", (minX + maxX) / 2.0R, maxY, 0.7)
        End If

        For Each s In allLines
            Dim mids As InterestingPoint = New InterestingPoint With {
                .SourceType = "DVLine2d",
                .SourceIndex = s.Idx,
                .Role = "Mid",
                .XView = s.Mx, .YView = s.My,
                .XSheet = s.Mx, .YSheet = s.My,
                .Confidence = 0.85}
            Dim xsMid As Double = 0, ysMid As Double = 0
            If TryViewToSheetBatch(dv, logger, s.Mx, s.My, xsMid, ysMid) Then mids.XSheet = xsMid : mids.YSheet = ysMid
            sink.Add(mids)
            Lg(logger, "[DVDIM][POINT][INTERESTING] view=" & viewIdx.ToString(CultureInfo.InvariantCulture) &
                " role=" & mids.Role & " source=" & mids.SourceType & " idx=" & s.Idx.ToString(CultureInfo.InvariantCulture) &
                " xView=" & F(mids.XView) & " yView=" & F(mids.YView) &
                " xSheet=" & F(mids.XSheet) & " ySheet=" & F(mids.YSheet))
        Next
    End Sub

    ''' <summary>Distancia euclídea en hoja entre puntos medios de dos líneas (ViewToSheet en ambos); NO CONFIRMADO EN SDK_HTML que represente la cota final.</summary>
    Private Shared Function SheetDistanceMidpoints(dv As Object, logger As Logger, a As LineSample, b As LineSample) As Double
        Dim x1s As Double, y1s As Double, x2s As Double, y2s As Double
        If Not TryMidSheet(dv, a, logger, x1s, y1s) OrElse Not TryMidSheet(dv, b, logger, x2s, y2s) Then Return Double.NaN
        Dim dx As Double = x2s - x1s
        Dim dy As Double = y2s - y1s
        Return Sqrt(dx * dx + dy * dy)
    End Function

    Private Shared Function PickVerticalSpan(vertLines As List(Of LineSample), allSamples As List(Of LineSample),
                                           bbox As Tuple(Of Double, Double, Double, Double),
                                           dv As Object, logger As Logger, viewIdx As Integer, tag As String) As Tuple(Of LineSample, LineSample, Double, Double)
        If vertLines.Count >= 2 Then
            Dim ordered = vertLines.OrderBy(Function(z) Min(z.Sx, z.Ex)).ToList()
            Dim leftMost As LineSample = ordered.First()
            Dim rightMost As LineSample = ordered.Last()
            If Not ReferenceEquals(leftMost, rightMost) Then
                Dim dist As Double = Abs(Min(rightMost.Sx, rightMost.Ex) - Max(leftMost.Sx, leftMost.Ex))
                Dim dSheet As Double = SheetDistanceMidpoints(dv, logger, leftMost, rightMost)
                Lg(logger, "[DVDIM][CAND][" & tag & "] view=" & viewIdx.ToString(CultureInfo.InvariantCulture) &
                    " leftSource=DVLine2d:" & leftMost.Idx.ToString(CultureInfo.InvariantCulture) &
                    " rightSource=DVLine2d:" & rightMost.Idx.ToString(CultureInfo.InvariantCulture) &
                    " spanViewX=" & F(dist) &
                    " distanceSheet=" & F(dSheet))
                Return Tuple.Create(leftMost, rightMost, dist, dSheet)
            End If
        End If
        Return FallbackLinePairByMidX(allSamples, dv, logger, viewIdx, tag)
    End Function

    Private Shared Function PickHorizontalSpan(horLines As List(Of LineSample), allSamples As List(Of LineSample),
                                               bbox As Tuple(Of Double, Double, Double, Double),
                                               dv As Object, logger As Logger, viewIdx As Integer, tag As String) As Tuple(Of LineSample, LineSample, Double, Double)
        If horLines.Count >= 2 Then
            Dim ordered = horLines.OrderBy(Function(z) Min(z.Sy, z.Ey)).ToList()
            Dim bot As LineSample = ordered.First()
            Dim topMost As LineSample = ordered.Last()
            If Not ReferenceEquals(bot, topMost) Then
                Dim dist As Double = Abs(Max(topMost.Sy, topMost.Ey) - Min(bot.Sy, bot.Ey))
                Dim dSheet As Double = SheetDistanceMidpoints(dv, logger, bot, topMost)
                Lg(logger, "[DVDIM][CAND][" & tag & "] view=" & viewIdx.ToString(CultureInfo.InvariantCulture) &
                    " bottomSource=DVLine2d:" & bot.Idx.ToString(CultureInfo.InvariantCulture) &
                    " topSource=DVLine2d:" & topMost.Idx.ToString(CultureInfo.InvariantCulture) &
                    " spanViewY=" & F(dist) &
                    " distanceSheet=" & F(dSheet))
                Return Tuple.Create(bot, topMost, dist, dSheet)
            End If
        End If
        Return FallbackLinePairByMidY(allSamples, dv, logger, viewIdx, tag)
    End Function

    Private Shared Function FallbackLinePairByMidX(lines As List(Of LineSample), dv As Object, logger As Logger, viewIdx As Integer, tag As String) As Tuple(Of LineSample, LineSample, Double, Double)
        If lines.Count = 0 Then Return Tuple.Create(CType(Nothing, LineSample), CType(Nothing, LineSample), 0.0R, Double.NaN)
        Dim bot = lines.Aggregate(Function(acc, ln) If(acc.Mx <= ln.Mx, acc, ln))
        Dim top = lines.Aggregate(Function(acc, ln) If(acc.Mx >= ln.Mx, acc, ln))
        If ReferenceEquals(bot, top) Then Return Tuple.Create(CType(Nothing, LineSample), CType(Nothing, LineSample), 0.0R, Double.NaN)
        Dim dist As Double = Abs(top.Mx - bot.Mx)
        Dim dSheet As Double = SheetDistanceMidpoints(dv, logger, bot, top)
        Lg(logger, "[DVDIM][CAND][" & tag & "_FALLBACK_MID_X] view=" & viewIdx.ToString(CultureInfo.InvariantCulture) &
            " DVLine2d_A=" & bot.Idx.ToString(CultureInfo.InvariantCulture) & " DVLine2d_B=" & top.Idx.ToString(CultureInfo.InvariantCulture) &
            " span_est_midXview=" & F(dist) & " distanceSheet=" & F(dSheet))
        Return Tuple.Create(bot, top, dist, dSheet)
    End Function

    Private Shared Function FallbackLinePairByMidY(lines As List(Of LineSample), dv As Object, logger As Logger, viewIdx As Integer, tag As String) As Tuple(Of LineSample, LineSample, Double, Double)
        If lines.Count = 0 Then Return Tuple.Create(CType(Nothing, LineSample), CType(Nothing, LineSample), 0.0R, Double.NaN)
        Dim low = lines.Aggregate(Function(acc, ln) If(acc.My <= ln.My, acc, ln))
        Dim high = lines.Aggregate(Function(acc, ln) If(acc.My >= ln.My, acc, ln))
        If ReferenceEquals(low, high) Then Return Tuple.Create(CType(Nothing, LineSample), CType(Nothing, LineSample), 0.0R, Double.NaN)
        Dim dist As Double = Abs(high.My - low.My)
        Dim dSheet As Double = SheetDistanceMidpoints(dv, logger, low, high)
        Lg(logger, "[DVDIM][CAND][" & tag & "_FALLBACK_MID_Y] view=" & viewIdx.ToString(CultureInfo.InvariantCulture) &
            " DVLine2d_A=" & low.Idx.ToString(CultureInfo.InvariantCulture) & " DVLine2d_B=" & high.Idx.ToString(CultureInfo.InvariantCulture) &
            " span_est_midYview=" & F(dist) & " distanceSheet=" & F(dSheet))
        Return Tuple.Create(low, high, dist, dSheet)
    End Function

    Private Shared Function TryCreateOuterDistance(dimsColl As Object, dv As Object,
                                                   a As LineSample, b As LineSample,
                                                   logger As Logger,
                                                   viewOrdinal As Integer, kind As String,
                                                   spanHint As Double) As CreateDimOutcome

        Dim x1s As Double, y1s As Double, z1 As Double = 0.0R
        Dim x2s As Double, y2s As Double, z2 As Double = 0.0R
        If Not TryMidSheet(dv, a, logger, x1s, y1s) OrElse Not TryMidSheet(dv, b, logger, x2s, y2s) Then
            Lg(logger, "[DVDIM][CREATE][FAIL] view=" & viewOrdinal.ToString(CultureInfo.InvariantCulture) &
                " kind=" & kind & " error=ViewToSheet_midpoint")
            Return CreateDimOutcome.Failed
        End If

        Dim kp1 As Boolean = False
        Dim kp2 As Boolean = False

        Lg(logger, "[DVDIM][CREATE][TRY] view=" & viewOrdinal.ToString(CultureInfo.InvariantCulture) &
            " kind=" & kind &
            " method=AddDistanceBetweenObjects obj1=DVLine2d:" & a.Idx.ToString(CultureInfo.InvariantCulture) &
            " obj2=DVLine2d:" & b.Idx.ToString(CultureInfo.InvariantCulture) &
            " pSheet1=(" & F(x1s) & "," & F(y1s) & ",0)" &
            " pSheet2=(" & F(x2s) & "," & F(y2s) & ",0)" &
            " kp1=" & kp1.ToString() & " kp2=" & kp2.ToString() &
            " NOTA=z=0 NO_CONFIRMADO_SDK espacio esperado punto proximidad DFT.")

        Dim dimObj As Object = Nothing
        Try
            dimObj = CallByName(dimsColl, "AddDistanceBetweenObjects", CallType.Method,
                                a.LineObj, x1s, y1s, z1, kp1,
                                b.LineObj, x2s, y2s, z2, kp2)
        Catch ex As Exception
            If logger IsNot Nothing Then logger.LogComException($"[DVDIM][CREATE] kind={kind}", ex)
            Lg(logger, "[DVDIM][CREATE][FAIL] view=" & viewOrdinal.ToString(CultureInfo.InvariantCulture) &
                " kind=" & kind & " error=COM primera variante (ver LogComException)")
            Return TrySecondVariantWithViewCoords(dimsColl, dv, a, b, logger, viewOrdinal, kind, spanHint)
        End Try

        If dimObj Is Nothing Then
            Lg(logger, "[DVDIM][CREATE][FAIL] view=" & viewOrdinal.ToString(CultureInfo.InvariantCulture) &
                " kind=" & kind & " error=Nothing retornado")
            Return TrySecondVariantWithViewCoords(dimsColl, dv, a, b, logger, viewOrdinal, kind, spanHint)
        End If

        Lg(logger, "[DVDIM][CREATE][OK] view=" & viewOrdinal.ToString(CultureInfo.InvariantCulture) &
            " kind=" & kind &
            " dimId=NO_CONFIRMADO_SDK value=" & F(TryLateGet(dimObj, "Value")))

        Return FinalizeCreatedDimension(dimObj, dv, logger, spanHint)
    End Function

    Private Shared Function FinalizeCreatedDimension(dimObj As Object, dv As Object, logger As Logger, spanHint As Double) As CreateDimOutcome
        TryTrack(logger, dimObj)
        TryReattachIfNeeded(logger, dimObj, dv)
        Dim st As ValidationOutcome = ValidateAndMaybeDelete(dimObj, logger, spanHint)
        Select Case st
            Case ValidationOutcome.DeletedInvalid
                Lg(logger, "[DVDIM][DELETE] dim=com_object reason=invalid_deleted (seDimStatusError)")
                Return CreateDimOutcome.Failed
            Case ValidationOutcome.KeptConnected
                Return CreateDimOutcome.Connected
            Case ValidationOutcome.KeptFloating
                Return CreateDimOutcome.FloatingVisible
            Case Else
                Return CreateDimOutcome.BadKept
        End Select
    End Function

    Private Shared Function TrySecondVariantWithViewCoords(dimsColl As Object, dv As Object, a As LineSample, b As LineSample,
                                                           logger As Logger, viewOrdinal As Integer, kind As String,
                                                           spanHint As Double) As CreateDimOutcome
        Dim z As Double = 0.0R
        Lg(logger, "[DVDIM][CREATE][TRY2_VIEW_COORDS] view=" & viewOrdinal.ToString(CultureInfo.InvariantCulture) &
            " kind=" & kind &
            " method=AddDistanceBetweenObjects p1=(" & F(a.Mx) & "," & F(a.My) & ",0) p2=(" & F(b.Mx) & "," & F(b.My) & ",0) kp1=False kp2=False")
        Try
            Dim d As Object = CallByName(dimsColl, "AddDistanceBetweenObjects", CallType.Method,
                                        a.LineObj, a.Mx, a.My, z, False,
                                        b.LineObj, b.Mx, b.My, z, False)
            If d IsNot Nothing Then
                Lg(logger, "[DVDIM][CREATE][OK] view=" & viewOrdinal.ToString(CultureInfo.InvariantCulture) &
                    " kind=" & kind & " dimId=NO_CONFIRMADO_SDK value=" & F(TryLateGet(d, "Value")) & " route=VIEW_COORDS")
                Return FinalizeCreatedDimension(d, dv, logger, spanHint)
            End If
            Lg(logger, "[DVDIM][CREATE][FAIL2_VIEW] kind=" & kind & " error=Nothing retornado")
        Catch ex As Exception
            Lg(logger, "[DVDIM][CREATE][FAIL2_VIEW] kind=" & kind & " error=" & ex.Message)
        End Try
        Return CreateDimOutcome.Failed
    End Function

    Private Shared Function TryMidSheet(dv As Object, ln As LineSample, logger As Logger,
                                        ByRef xs As Double, ByRef ys As Double) As Boolean
        Dim mx As Double = ln.Mx, my As Double = ln.My
        Return TryViewToSheetBatch(dv, logger, mx, my, xs, ys)
    End Function

    Private Enum ValidationOutcome As Integer
        KeptConnected = 0
        KeptFloating = 1
        DeletedInvalid = 2
    End Enum

    Private Shared Function ValidateAndMaybeDelete(dimObj As Object, logger As Logger, bboxHint As Double) As ValidationOutcome
        Try
            Dim st As Integer = SafeInt(CallByName(dimObj, "StatusOfDimension", CallType.Get), -99)
            Dim statusTxt As String = st.ToString(CultureInfo.InvariantCulture)
            ' SolidEdgeFrameworkSupport~DimStatusConstants.html
            Dim seErr As Integer = 2 ' seDimStatusError
            Dim seDetached As Integer = 1
            Dim seOneEndDetached As Integer = 5

            Dim valStr As String = F(TryLateGet(dimObj, "Value"))
            Dim cat As String
            Select Case st
                Case 3, 4 ' Driving / Driven
                    cat = "connected"
                Case seDetached, seOneEndDetached
                    cat = "floating_visible"
                Case Else
                    cat = "created_but_bad_position"
            End Select

            Lg(logger, "[DVDIM][VALIDATE] dim=com_object status=" & cat &
                " COM_StatusOfDimension_raw=" & statusTxt &
                " visible=UNKNOWN_SDK relatedCount=N/A value=" & valStr &
                " bboxHint=" & F(bboxHint))

            If cat = "connected" Then
                Lg(logger, "[DVDIM][KEEP] dim=com_object reason=connected")
                Return ValidationOutcome.KeptConnected
            End If
            If cat = "floating_visible" Then
                Lg(logger, "[DVDIM][KEEP] dim=com_object reason=floating_visible")
                Return ValidationOutcome.KeptFloating
            End If

            If st = seErr Then
                Try
                    CallByName(dimObj, "Delete", CallType.Method)
                    Lg(logger, "[DVDIM][DELETE] dim=com_object reason=invalid_deleted (seDimStatusError)")
                    Return ValidationOutcome.DeletedInvalid
                Catch exDel As Exception
                    Lg(logger, "[DVDIM][DELETE][FAIL] " & exDel.Message)
                    Return ValidationOutcome.KeptFloating
                End Try
            End If

            Lg(logger, "[DVDIM][KEEP] dim=com_object reason=created_but_bad_position")
            Return ValidationOutcome.KeptFloating
        Catch ex As Exception
            If logger IsNot Nothing Then logger.LogComException("[DVDIM][VALIDATE]", ex)
            Return ValidationOutcome.KeptFloating
        End Try
    End Function

    Private Shared Sub TryTrack(logger As Logger, dimObj As Object)
        Try
            Dim oldV As Double = SafeDbl(CallByName(dimObj, "TrackDistance", CallType.Get), Double.NaN)
            Lg(logger, "[DVDIM][TRACK][TRY] dim=com_object old=" & F(oldV) & " new=0.10")
            CallByName(dimObj, "TrackDistance", CallType.Set, 0.1R)
            Lg(logger, "[DVDIM][TRACK][OK] dim=com_object TrackDistance establecido 0.1 (unidades: NO_CONFIRMADO_SDK = doc units típicamente)")
        Catch ex As Exception
            Lg(logger, "[DVDIM][TRACK][FAIL] dim=com_object error=" & ex.Message &
                " Dimension.Update:NO_CONFIRMADO_EN_ESTA_MUESTRA_SDK_HTML")
        End Try
    End Sub

    Private Shared Sub TryReattachIfNeeded(logger As Logger, dimObj As Object, dv As Object)
        Try
            Dim st As Integer = SafeInt(CallByName(dimObj, "StatusOfDimension", CallType.Get), 0)
            If st <> 1 AndAlso st <> 5 Then Return
            Lg(logger, "[DVDIM][REATTACH][TRY] dim=com_object view=attached_DrawingView")
            Dim rv As Object = Nothing
            Try
                rv = CallByName(dimObj, "ReattachToDrawingView", CallType.Method, dv)
            Catch ex As Exception
                Lg(logger, "[DVDIM][REATTACH][RESULT] dim=com_object err=" & ex.Message)
                Return
            End Try
            Dim rs As String = If(rv Is Nothing, "(null)", rv.ToString())
            Lg(logger, "[DVDIM][REATTACH][RESULT] dim=com_object result=" & rs)
        Catch ex As Exception
            Lg(logger, "[DVDIM][REATTACH][SKIP] err=" & ex.Message)
        End Try
    End Sub

    Private Shared Function GetCollItem(coll As Object, idx As Integer) As Object
        If coll Is Nothing Then Return Nothing
        Try : Return CallByName(coll, "Item", CallType.Get, idx) : Catch : End Try
        Try : Return CallByName(coll, "Item", CallType.Method, idx) : Catch : End Try
        Return Nothing
    End Function

    Private Shared Function SafeInt(o As Object, defVal As Integer) As Integer
        Try
            Return CInt(o)
        Catch
            Return defVal
        End Try
    End Function

    Private Shared Function SafeDbl(o As Object, defVal As Double) As Double
        Try
            Return CDbl(o)
        Catch
            Return defVal
        End Try
    End Function

    Private Shared Function SafeStr(o As Object) As String
        If o Is Nothing Then Return ""
        Try
            Return Convert.ToString(o).Trim()
        Catch
            Return ""
        End Try
    End Function

    Private Shared Function TryLateGet(o As Object, prop As String) As Object
        If o Is Nothing Then Return Nothing
        Try : Return CallByName(o, prop, CallType.Get) : Catch : End Try
        Return Nothing
    End Function

    Private Shared Function F(x As Object) As String
        If x Is Nothing Then Return "null"
        If TypeOf x Is Double Then
            Dim d As Double = CDbl(x)
            If Double.IsNaN(d) OrElse Double.IsInfinity(d) Then Return "NaN"
        End If
        Try
            Return Convert.ToString(x, CultureInfo.InvariantCulture)
        Catch
            Return "?"
        End Try
    End Function

    Private Shared Sub Lg(logger As Logger, msg As String)
        If logger IsNot Nothing Then
            Try
                logger.Log(msg)
            Catch
            End Try
        Else
            System.Diagnostics.Debug.WriteLine(msg)
        End If
    End Sub

End Class
