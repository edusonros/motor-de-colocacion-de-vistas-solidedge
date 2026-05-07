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
''' Laboratorio: Drop de DrawingViews, detección de sheets nuevas por referencia antes/después,
''' auditoría de Lines2d reales en esas sheets e intentos de cotación sin DVLine2d.
''' </summary>
Friend NotInheritable Class DropCreatedSheetsDimensionLab

    Private Const DesiredDimStyle As String = "U3,5"
    Private Const AuxLayer As String = "DROP_DIM_AUX"
    Private Const TolAxisRatio As Double = 0.0005R

    Private Class SheetEntityFingerprint
        Public SheetRef As Sheet
        Public Lines As Integer
        Public Arcs As Integer
        Public Circles As Integer
        Public BSplines As Integer
        Public LineStrings As Integer
        Public Points As Integer
    End Class

    Private Class LineGeom
        Public Idx As Integer
        Public Obj As Line2d
        Public X1 As Double
        Public Y1 As Double
        Public X2 As Double
        Public Y2 As Double
        Public Orientation As String

        Friend Function Midx() As Double
            Return (X1 + X2) * 0.5R
        End Function

        Friend Function Midy() As Double
            Return (Y1 + Y2) * 0.5R
        End Function
    End Class

    Private Shared ReadOnly ReflectInvoke As BindingFlags =
        BindingFlags.InvokeMethod Or BindingFlags.Public Or BindingFlags.Instance

    Public Shared Sub Run(app As Application,
                          draft As DraftDocument,
                          logSink As Action(Of String),
                          Optional debugSave As Boolean = False)
        Dim L = Sub(m As String)
                    Try
                        logSink?.Invoke(m)
                    Catch
                    End Try
                End Sub

        L("[DROP_SHEETS][START]")
        If draft Is Nothing Then
            L("[DROP_SHEETS][ABORT] reason=no_draft")
            Return
        End If

        Dim path As String = ""
        Try : path = Convert.ToString(draft.FullName, CultureInfo.InvariantCulture) : Catch : path = "" : End Try
        L("[DROP_SHEETS][DFT] path=" & path)

        Dim bridge As New Logger(
            Sub(line As String)
                Try : logSink?.Invoke(line) : Catch : End Try
            End Sub)
        Dim dimLogger As New DimensionLogger(bridge)

        Dim summaryViews As Integer = 0
        Dim summaryDropOk As Integer = 0
        Dim summaryDropFail As Integer = 0
        Dim summarySheetsCreated As Integer = 0
        Dim summarySheetsAnalyzed As Integer = 0
        Dim summaryLines As Long = 0
        Dim summaryArcs As Long = 0
        Dim summaryCircles As Long = 0
        Dim summaryBsplines As Long = 0
        Dim summaryKeys As Long = 0
        Dim summaryCp As Long = 0
        Dim summaryHcand As Integer = 0
        Dim summaryVcand As Integer = 0
        Dim summaryDims As Integer = 0
        Dim summaryConn As Integer = 0
        Dim summaryFloat As Integer = 0
        Dim adboOk As Integer = 0
        Dim adboFail As Integer = 0
        Dim addLenAuxOk As Integer = 0
        Dim bestMethod As String = "none"

        Dim sourceSheet As Sheet = ResolveSourceSheetDraft(draft, L)
        If sourceSheet Is Nothing Then
            L("[DROP_SHEETS][ABORT] reason=no_source_sheet_with_views")
            Return
        End If

        Dim srcName As String = SafeSheetName(sourceSheet)
        Dim nViews As Integer = SafeCount(CallByNameSafeGet(sourceSheet, "DrawingViews"))
        L("[DROP_SHEETS][SOURCE_SHEET] name=" & srcName & " views=" & nViews.ToString(CultureInfo.InvariantCulture))

        Dim initialSnapshots As List(Of SheetSnapshotOrder) = EnumerateOrderedSheets(draft)
        L("[DROP_SHEETS][SHEETS_BEFORE] count=" & initialSnapshots.Count.ToString(CultureInfo.InvariantCulture))
        For Each s As SheetSnapshotOrder In initialSnapshots
            L("[DROP_SHEETS][SHEET_BEFORE] idx=" & s.Ordinal.ToString(CultureInfo.InvariantCulture) & " name=" & s.Name)
        Next

        Dim viewsOrder As New List(Of DrawingView)()
        For i As Integer = 1 To nViews
            Dim dv As DrawingView = Nothing
            Try : dv = CType(CallByName(CallByNameSafeGet(sourceSheet, "DrawingViews"), "Item", CallType.Method, i), DrawingView) : Catch : End Try
            If dv IsNot Nothing Then viewsOrder.Add(dv)
        Next

        Dim sheetsToAnalyze As New List(Of Sheet)()

        Dim viewIdx As Integer = 0
        For Each dv In viewsOrder
            viewIdx += 1
            summaryViews += 1
            Dim vname As String = SafeStr(CallByNameSafe(dv, "Name"))

            Dim typ As String = SafeStr(CallByNameSafe(dv, "DrawingViewType"))
            Dim scaleFac As String = SafeStr(CallByNameSafe(dv, "Scale"))
            If String.IsNullOrWhiteSpace(scaleFac) Then scaleFac = SafeStr(CallByNameSafe(dv, "ScaleFactor"))
            L("[DROP_SHEETS][VIEW][BEFORE] idx=" & viewIdx.ToString(CultureInfo.InvariantCulture) & " name=" & vname & " type=" & typ & " scale=" & scaleFac)

            Dim cntL As Integer = SafeCount(CallByNameSafe(dv, "DVLines2d"))
            Dim cntA As Integer = SafeCount(CallByNameSafe(dv, "DVArcs2d"))
            Dim cntC As Integer = SafeCount(CallByNameSafe(dv, "DVCircles2d"))
            Dim cntSp As Integer = SafeCount(CallByNameSafe(dv, "DVBSplineCurves2d"))
            If cntSp = 0 Then cntSp = SafeCount(CallByNameSafe(dv, "DVBSplines2d"))
            Dim cntP As Integer = SafeCount(CallByNameSafe(dv, "DVPoints2d"))
            L("[DROP_SHEETS][VIEW][DVCOUNT] lines=" & cntL.ToString(CultureInfo.InvariantCulture) &
              " arcs=" & cntA.ToString(CultureInfo.InvariantCulture) &
              " circles=" & cntC.ToString(CultureInfo.InvariantCulture) &
              " splines=" & cntSp.ToString(CultureInfo.InvariantCulture) &
              " points=" & cntP.ToString(CultureInfo.InvariantCulture))

            Dim rx1 As Double = 0, ry1 As Double = 0, rx2 As Double = 0, ry2 As Double = 0
            Try
                dv.Range(rx1, ry1, rx2, ry2)
                L("[DROP_SHEETS][VIEW][RANGE] xmin=" & FormatInv(Math.Min(rx1, rx2)) & " ymin=" & FormatInv(Math.Min(ry1, ry2)) &
                  " xmax=" & FormatInv(Math.Max(rx1, rx2)) & " ymax=" & FormatInv(Math.Max(ry1, ry2)))
            Catch ex As Exception
                L("[DROP_SHEETS][VIEW][RANGE] FAIL NO_CONFIRMADO " & ex.Message)
            End Try

            Dim beforeRefs As List(Of Sheet) = EnumerateSheetsRefList(draft)
            Dim fingerBefore As Dictionary(Of Sheet, SheetEntityFingerprint) = CaptureAllFingerprints(draft)

            L("[DROP_SHEETS][VIEW][DROP_TRY] method=Drop")
            Dim dropMethod As String = ""
            Dim dropOk As Boolean = TryDropView(dv, dropMethod, L)
            If Not dropOk Then
                summaryDropFail += 1
                L("[DROP_SHEETS][VIEW][DROP_FAIL] view=" & vname & " error=no_supported_method")
                LogGeometryDeltaWhenNoNewSheet(draft, fingerBefore, sourceSheet, vname, L)
                Continue For
            End If

            summaryDropOk += 1
            L("[DROP_SHEETS][VIEW][DROP_OK] view=" & vname & " method=" & dropMethod)

            Dim afterRefs As List(Of Sheet) = EnumerateSheetsRefList(draft)
            Dim newSheets As List(Of Sheet) = FindNewSheets(beforeRefs, afterRefs)
            If newSheets.Count = 0 Then
                L("[DROP_SHEETS][SHEET_CREATED][NONE] view=" & vname)
                LogGeometryDeltaWhenNoNewSheet(draft, fingerBefore, sourceSheet, vname, L)
            Else
                If newSheets.Count > 1 Then
                    L("[DROP_SHEETS][SHEET_CREATED][MULTI] view=" & vname & " count=" & newSheets.Count.ToString(CultureInfo.InvariantCulture) & " note=NO_CONFIRMADO_assoc_first")
                End If
                Dim created As Sheet = newSheets(0)
                sheetsToAnalyze.Add(created)
                summarySheetsCreated += newSheets.Count
                Dim ord As Integer = SheetOrdinal(draft, created)
                L("[DROP_SHEETS][SHEET_CREATED] view=" & vname & " sheet=" & SafeSheetName(created) & " idx=" & ord.ToString(CultureInfo.InvariantCulture))
            End If
        Next

        Dim processed As New HashSet(Of Sheet)()
        For Each sh In sheetsToAnalyze
            If sh Is Nothing OrElse Not processed.Add(sh) Then Continue For
            summarySheetsAnalyzed += 1
            RunPhasesOnCreatedSheet(draft, sh, dimLogger, L,
                summaryLines, summaryArcs, summaryCircles, summaryBsplines,
                summaryKeys, summaryCp, summaryHcand, summaryVcand,
                summaryDims, summaryConn, summaryFloat, adboOk, adboFail, addLenAuxOk, bestMethod)
        Next

        L("[DROP_SHEETS][SUMMARY]")
        L("views_processed=" & summaryViews.ToString(CultureInfo.InvariantCulture))
        L("drop_ok=" & summaryDropOk.ToString(CultureInfo.InvariantCulture))
        L("drop_fail=" & summaryDropFail.ToString(CultureInfo.InvariantCulture))
        L("created_sheets=" & summarySheetsCreated.ToString(CultureInfo.InvariantCulture))
        L("sheets_analyzed=" & summarySheetsAnalyzed.ToString(CultureInfo.InvariantCulture))
        L("lines2d_found=" & summaryLines.ToString(CultureInfo.InvariantCulture))
        L("arcs2d_found=" & summaryArcs.ToString(CultureInfo.InvariantCulture))
        L("circles2d_found=" & summaryCircles.ToString(CultureInfo.InvariantCulture))
        L("bsplines2d_found=" & summaryBsplines.ToString(CultureInfo.InvariantCulture))
        L("keypoints_found=" & summaryKeys.ToString(CultureInfo.InvariantCulture))
        L("connectpoints_found=" & summaryCp.ToString(CultureInfo.InvariantCulture))
        L("h_total_candidates=" & summaryHcand.ToString(CultureInfo.InvariantCulture))
        L("v_total_candidates=" & summaryVcand.ToString(CultureInfo.InvariantCulture))
        L("dimensions_created=" & summaryDims.ToString(CultureInfo.InvariantCulture))
        L("dimensions_connected=" & summaryConn.ToString(CultureInfo.InvariantCulture))
        L("dimensions_floating=" & summaryFloat.ToString(CultureInfo.InvariantCulture))
        L("adbo_on_lines2d_ok=" & adboOk.ToString(CultureInfo.InvariantCulture))
        L("adbo_on_lines2d_fail=" & adboFail.ToString(CultureInfo.InvariantCulture))
        L("addlength_aux_ok=" & addLenAuxOk.ToString(CultureInfo.InvariantCulture))
        L("best_method=" & bestMethod)
        L("recommended_next_step=valorar_ADBO_vs_AUX_y_refinar_candidates_paralelos")

        If debugSave Then
            Try : draft.Save() : L("[DROP_SHEETS][SAVE] invoked=True") : Catch ex As Exception : L("[DROP_SHEETS][SAVE][FAIL] " & ex.Message) : End Try
        End If
    End Sub

    Private Shared Sub RunPhasesOnCreatedSheet(
        draft As DraftDocument,
        sh As Sheet,
        dimLogger As DimensionLogger,
        log As Action(Of String),
        ByRef totalLines As Long,
        ByRef totalArcs As Long,
        ByRef totalCircles As Long,
        ByRef totalBsplines As Long,
        ByRef totalKp As Long,
        ByRef totalCp As Long,
        ByRef hCandCount As Integer,
        ByRef vCandCount As Integer,
        ByRef dims As Integer,
        ByRef conn As Integer,
        ByRef floats As Integer,
        ByRef adboOk As Integer,
        ByRef adboFail As Integer,
        ByRef addLenAuxOk As Integer,
        ByRef bestMethod As String)

        Dim sname As String = SafeSheetName(sh)
        log("[DROP_SHEETS][CREATED_SHEET][AUDIT] sheet=" & sname)
        Try : CallByName(sh, "Activate", CallType.Method) : Catch : log("[DROP_SHEETS][CREATED_SHEET][ACTIVATE][WARN] sheet=" & sname) : End Try

        Dim lc As Integer = SafeCount(CallByNameSafe(sh, "Lines2d"))
        Dim ac As Integer = SafeCount(CallByNameSafe(sh, "Arcs2d"))
        Dim cc As Integer = SafeCount(CallByNameSafe(sh, "Circles2d"))
        Dim bc As Integer = SafeCount(CallByNameSafe(sh, "BSplineCurves2d"))
        Dim lst As Integer = SafeCount(CallByNameSafe(sh, "LineStrings2d"))
        Dim pt As Integer = SafeCount(CallByNameSafe(sh, "Points2d"))
        Dim dimc As Integer = SafeCount(CallByNameSafe(sh, "Dimensions"))

        totalLines += lc
        totalArcs += ac
        totalCircles += cc
        totalBsplines += bc

        log("[DROP_SHEETS][CREATED_SHEET][COUNT] lines=" & lc.ToString(CultureInfo.InvariantCulture) &
            " arcs=" & ac.ToString(CultureInfo.InvariantCulture) &
            " circles=" & cc.ToString(CultureInfo.InvariantCulture) &
            " bsplines=" & bc.ToString(CultureInfo.InvariantCulture) &
            " linestrings=" & lst.ToString(CultureInfo.InvariantCulture) &
            " points=" & pt.ToString(CultureInfo.InvariantCulture) &
            " dims=" & dimc.ToString(CultureInfo.InvariantCulture))

        LogLayersOutline(sh, log, sname)

        Dim lineGeoms As List(Of LineGeom) = ReadAllLines2d(sh, log, sname, totalKp, totalCp)
        ReadArcsAndCirclesDetailed(sh, log, sname, totalKp, totalCp)

        Dim bboxMinX As Double = Double.PositiveInfinity
        Dim bboxMaxX As Double = Double.NegativeInfinity
        Dim bboxMinY As Double = Double.PositiveInfinity
        Dim bboxMaxY As Double = Double.NegativeInfinity
        AccumulateBBoxFromLines(lineGeoms, bboxMinX, bboxMaxX, bboxMinY, bboxMaxY)
        AccumulateBBoxFromArcCirclesDraft(sh, bboxMinX, bboxMaxX, bboxMinY, bboxMaxY, log, sname)

        If Double.IsInfinity(bboxMinX) Then
            log("[DROP_SHEETS][BBOX] sheet=" & sname & " xmin=EMPTY xmax=EMPTY ymin=EMPTY ymax=EMPTY width=EMPTY height=EMPTY note=NO_CONFIRMADO")
        Else
            Dim w As Double = bboxMaxX - bboxMinX
            Dim hh As Double = bboxMaxY - bboxMinY
            log("[DROP_SHEETS][BBOX] sheet=" & sname & " xmin=" & FormatInv(bboxMinX) & " xmax=" & FormatInv(bboxMaxX) &
                " ymin=" & FormatInv(bboxMinY) & " ymax=" & FormatInv(bboxMaxY) & " width=" & FormatInv(w) & " height=" & FormatInv(hh))
        End If

        Dim leftV As LineGeom = Nothing
        Dim rightV As LineGeom = Nothing
        Dim bottomH As LineGeom = Nothing
        Dim topH As LineGeom = Nothing
        ClassifyExtents(lineGeoms, leftV, rightV, bottomH, topH, Math.Max(Math.Max((bboxMaxX - bboxMinX), (bboxMaxY - bboxMinY)), 0.001R))

        If leftV IsNot Nothing AndAlso rightV IsNot Nothing Then
            hCandCount += 1
            log("[DROP_SHEETS][CAND][H_TOTAL] sheet=" & sname & " leftEntity=Line2d rightEntity=Line2d expected=" &
                FormatInv(Math.Abs(rightV.Midx() - leftV.Midx())))
        Else
            log("[DROP_SHEETS][CAND][H_TOTAL] sheet=" & sname & " leftEntity=NONE rightEntity=NONE expected=NO_CONFIRMADO")
        End If
        If bottomH IsNot Nothing AndAlso topH IsNot Nothing Then
            vCandCount += 1
            log("[DROP_SHEETS][CAND][V_TOTAL] sheet=" & sname & " bottomEntity=Line2d topEntity=Line2d expected=" &
                FormatInv(Math.Abs(topH.Midy() - bottomH.Midy())))
        Else
            log("[DROP_SHEETS][CAND][V_TOTAL] sheet=" & sname & " bottomEntity=NONE topEntity=NONE expected=NO_CONFIRMADO")
        End If

        LogPartialCandidates(lineGeoms, log, sname)

        Dim dimsObj As Dimensions = Nothing
        Try : dimsObj = CType(CallByName(sh, "Dimensions", CallType.Get), Dimensions) : Catch : End Try
        If dimsObj Is Nothing Then
            log("[DROP_SHEETS][DIM][FAIL] sheet=" & sname & " method=(none) error=no_dimensions_collection")
            Return
        End If

        Dim layerAux As Object = EnsureLayerSafe(sh, AuxLayer, log)

        Dim yMid As Double = (bboxMinY + bboxMaxY) * 0.5R
        Dim xMid As Double = (bboxMinX + bboxMaxX) * 0.5R
        Dim pad As Double = Math.Max(Math.Max((bboxMaxX - bboxMinX), (bboxMaxY - bboxMinY)) * 0.05R, 0.004R)

        Dim methodUsed As String = ""

        If leftV IsNot Nothing AndAlso rightV IsNot Nothing Then
            log("[DROP_SHEETS][DIM][TRY] sheet=" & sname & " method=AddDistanceBetweenObjects type=H_TOTAL entity1=Line2d entity2=Line2d")
            Dim dObj As FrameworkDimension = Nothing
            Dim okDist = DimensionPlacementEngine.TryAddDistanceBetweenObjectsGeneric(
                dimsObj, leftV.Obj, rightV.Obj,
                bboxMinX, yMid, bboxMaxX, yMid,
                dimLogger, "horizontal", methodUsed, "DROP_SHEETS_H_TOTAL", Nothing, Nothing, True, dObj)
            If okDist AndAlso dObj IsNot Nothing AndAlso Not IsDegenerateDimensionValue(dObj) Then
                adboOk += 1
                dims += 1
                bestMethod = If(String.IsNullOrWhiteSpace(methodUsed), "AddDistanceBetweenObjects", methodUsed)
                ApplyU35OrLog(draft, dObj, dimLogger, log)
                LogDimOutcome(dObj, log, sname, methodUsed, conn, floats)
            Else
                adboFail += 1
                log("[DROP_SHEETS][DIM][FAIL] sheet=" & sname & " method=AddDistanceBetweenObjects error=" &
                    If(dObj Is Nothing, "no_dimension_returned", "degenerate_or_rejected"))
                TryAuxHorizontalLength(draft, sh, dimsObj, layerAux, bboxMinX, bboxMaxX, yMid - pad, dimLogger, log, sname, addLenAuxOk, dims, conn, floats, bestMethod)
            End If
        End If

        If bottomH IsNot Nothing AndAlso topH IsNot Nothing Then
            log("[DROP_SHEETS][DIM][TRY] sheet=" & sname & " method=AddDistanceBetweenObjects type=V_TOTAL entity1=Line2d entity2=Line2d")
            Dim dObj2 As FrameworkDimension = Nothing
            Dim okV = DimensionPlacementEngine.TryAddDistanceBetweenObjectsGeneric(
                dimsObj, bottomH.Obj, topH.Obj,
                xMid, bboxMinY, xMid, bboxMaxY,
                dimLogger, "vertical", methodUsed, "DROP_SHEETS_V_TOTAL", Nothing, Nothing, True, dObj2)
            If okV AndAlso dObj2 IsNot Nothing AndAlso Not IsDegenerateDimensionValue(dObj2) Then
                adboOk += 1
                dims += 1
                If String.IsNullOrWhiteSpace(bestMethod) Then bestMethod = If(String.IsNullOrWhiteSpace(methodUsed), "AddDistanceBetweenObjects", methodUsed)
                ApplyU35OrLog(draft, dObj2, dimLogger, log)
                LogDimOutcome(dObj2, log, sname, methodUsed, conn, floats)
            Else
                adboFail += 1
                log("[DROP_SHEETS][DIM][FAIL] sheet=" & sname & " method=AddDistanceBetweenObjects_v error=" &
                    If(dObj2 Is Nothing, "no_dimension_returned", "degenerate_or_rejected"))
                TryAuxVerticalDistanceOrLength(draft, sh, dimsObj, layerAux, bboxMinY, bboxMaxY, xMid + pad,
                    dimLogger, log, sname, addLenAuxOk, dims, conn, floats, bestMethod)
            End If
        End If

        If ac > 0 Then
            Dim ar As Object = Nothing
            Try : ar = CallByName(CallByNameSafeGet(sh, "Arcs2d"), "Item", CallType.Method, 1) : Catch : End Try
            If ar IsNot Nothing Then
                log("[DROP_SHEETS][DIM][TRY] sheet=" & sname & " method=AddRadius probe=Arc2d ")
                Dim dRad As FrameworkDimension = DrawingViewDimensionCreator.TryCreateRadiusOnReference(dimsObj, ar, dimLogger)
                If dRad IsNot Nothing AndAlso Not IsDegenerateDimensionValue(dRad) Then
                    dims += 1
                    If String.IsNullOrWhiteSpace(bestMethod) Then bestMethod = "AddRadius"
                    ApplyU35OrLog(draft, dRad, dimLogger, log)
                    LogDimOutcome(dRad, log, sname, "AddRadius", conn, floats)
                Else
                    log("[DROP_SHEETS][DIM][FAIL] sheet=" & sname & " method=AddRadius error=no_dimension_returned")
                End If
            End If
        End If
    End Sub

    Private Shared Sub TryAuxHorizontalLength(draft As DraftDocument, sh As Sheet, dimsObj As Dimensions, layerAux As Object,
                                               bx1 As Double, bx2 As Double, auxY As Double,
                                               dimLogger As DimensionLogger, log As Action(Of String),
                                               sheetName As String,
                                               ByRef addLenAuxOk As Integer,
                                               ByRef dims As Integer, ByRef conn As Integer,
                                               ByRef floats As Integer, ByRef bestMethod As String)

        Dim l2 As Object = Nothing
        Try : l2 = CallByName(sh.Lines2d, "AddBy2Points", CallType.Method, bx1, auxY, bx2, auxY) : Catch : End Try
        If l2 Is Nothing Then
            log("[DROP_SHEETS][DIM][FALLBACK] sheet=" & sheetName & " method=AddLengthAuxLine fail=AddBy2Points")
            Return
        End If
        TrySetLayerLate(l2, layerAux)
        log("[DROP_SHEETS][DIM][FALLBACK] sheet=" & sheetName & " method=AddLengthAuxLine")
        log("[DROP_SHEETS][DIM][AUX_LINE_CREATE] sheet=" & sheetName & " x1=" & FormatInv(bx1) & " y1=" & FormatInv(auxY) & " x2=" & FormatInv(bx2) & " y2=" & FormatInv(auxY))
        log("[DROP_SHEETS][DIM][AUX_LINE_KEEP] reason=dimension_depends_on_aux_geometry")

        Dim d As FrameworkDimension = Nothing
        Try : d = TryCast(CallByName(dimsObj, "AddLength", CallType.Method, l2), FrameworkDimension) : Catch : End Try
        If d IsNot Nothing AndAlso Not IsDegenerateDimensionValue(d) Then
            addLenAuxOk += 1
            dims += 1
            If String.IsNullOrWhiteSpace(bestMethod) Then bestMethod = "AddLength"
            ApplyU35OrLog(draft, d, dimLogger, log)
            LogDimOutcome(d, log, sheetName, "AddLength_auxH", conn, floats)
        Else
            log("[DROP_SHEETS][DIM][FAIL] sheet=" & sheetName & " method=AddLength_aux error=no_dimension_returned")
        End If
    End Sub

    Private Shared Sub TryAuxVerticalDistanceOrLength(draft As DraftDocument, sh As Sheet, dimsObj As Dimensions, layerAux As Object,
                                                       by1 As Double, by2 As Double, auxX As Double,
                                                       dimLogger As DimensionLogger, log As Action(Of String),
                                                       sheetName As String,
                                                       ByRef addLenAuxOk As Integer,
                                                       ByRef dims As Integer, ByRef conn As Integer,
                                                       ByRef floats As Integer, ByRef bestMethod As String)

        Dim l2 As Object = Nothing
        Try : l2 = CallByName(sh.Lines2d, "AddBy2Points", CallType.Method, auxX, by1, auxX, by2) : Catch : End Try
        If l2 Is Nothing Then
            Return
        End If
        TrySetLayerLate(l2, layerAux)
        log("[DROP_SHEETS][DIM][FALLBACK] sheet=" & sheetName & " method=AddLengthAuxLine_vertical")
        log("[DROP_SHEETS][DIM][AUX_LINE_CREATE] sheet=" & sheetName & " x1=" & FormatInv(auxX) & " y1=" & FormatInv(by1) & " x2=" & FormatInv(auxX) & " y2=" & FormatInv(by2))

        Dim d As FrameworkDimension = Nothing
        Try : d = TryCast(CallByName(dimsObj, "AddLength", CallType.Method, l2), FrameworkDimension) : Catch : End Try
        If d IsNot Nothing AndAlso Not IsDegenerateDimensionValue(d) Then
            addLenAuxOk += 1
            dims += 1
            If String.IsNullOrWhiteSpace(bestMethod) Then bestMethod = "AddLength"
            ApplyU35OrLog(draft, d, dimLogger, log)
            LogDimOutcome(d, log, sheetName, "AddLength_auxV", conn, floats)
        Else
            log("[DROP_SHEETS][DIM][FAIL] sheet=" & sheetName & " method=AddLength_auxV error=no_dimension_returned")
        End If
    End Sub

    Private Shared Sub LogDimOutcome(d As FrameworkDimension, log As Action(Of String), sheetName As String, methodTag As String,
                                     ByRef conn As Integer, ByRef floats As Integer)
        Dim rel As Integer = 0
        Dim relObj As Object = Nothing
        Try : relObj = CallByName(d, "GetRelatedObjects", CallType.Method) : Catch : End Try
        If relObj Is Nothing Then Try : relObj = CallByNameSafeGet(d, "RelatedObjects") : Catch : End Try
        rel = SafeCount(relObj)
        Dim st As String = SafeStr(CallByNameSafe(d, "Status"))
        Dim val As String = SafeStr(CallByNameSafe(d, "Value"))
        Dim sty As String = SafeStr(CallByNameSafe(d, "StyleName"))
        log("[DROP_SHEETS][DIM][OK] sheet=" & sheetName & " method=" & methodTag & " value=" & val & " style=" & sty)
        log("[DROP_SHEETS][DIM][RELATED] count=" & rel.ToString(CultureInfo.InvariantCulture) & " sig=NO_CONFIRMADO")
        log("[DROP_SHEETS][DIM][STATUS] status=" & st)
        If rel > 0 Then conn += 1 Else floats += 1
    End Sub

    Private Shared Sub ApplyU35OrLog(draft As DraftDocument, d As FrameworkDimension, dimLogger As DimensionLogger, log As Action(Of String))
        If d Is Nothing OrElse draft Is Nothing Then Return
        Dim stObj As Object = DimensionStyleResolver.ResolveDimensionStyle(draft, DesiredDimStyle, dimLogger)
        If stObj Is Nothing Then
            log("[DROP_SHEETS][DIM][STYLE][FAIL] requested=" & DesiredDimStyle & " error=not_found")
            Return
        End If
        If DimensionStyleResolver.ApplyDimensionStyle(d, stObj, dimLogger) Then
            log("[DROP_SHEETS][DIM][STYLE][OK] " & DesiredDimStyle)
        Else
            log("[DROP_SHEETS][DIM][STYLE][FAIL] requested=" & DesiredDimStyle & " error=apply_rejected")
        End If
    End Sub

    Private Shared Function IsDegenerateDimensionValue(d As FrameworkDimension) As Boolean
        If d Is Nothing Then Return True
        Try
            Dim v As Double = Convert.ToDouble(CallByName(d, "Value", CallType.Get), CultureInfo.InvariantCulture)
            If Math.Abs(v) < 1.0E-12R Then Return True
        Catch
        End Try
        Return False
    End Function

    Private Shared Sub LogGeometryDeltaWhenNoNewSheet(draft As DraftDocument,
                                                      fingerBefore As Dictionary(Of Sheet, SheetEntityFingerprint),
                                                      sourceSheet As Sheet,
                                                      viewName As String,
                                                      log As Action(Of String))
        Dim after = CaptureAllFingerprints(draft)
        For Each kvp In after
            Dim sh As Sheet = kvp.Key
            If sh Is Nothing Then Continue For
            Dim nm As String = SafeSheetName(sh)
            Dim beforeFp As SheetEntityFingerprint = Nothing
            If Not fingerBefore.TryGetValue(sh, beforeFp) Then beforeFp = Nothing
            Dim bL As Integer = If(beforeFp IsNot Nothing, beforeFp.Lines, 0)
            Dim bA As Integer = If(beforeFp IsNot Nothing, beforeFp.Arcs, 0)
            If kvp.Value.Lines <> bL OrElse kvp.Value.Arcs <> bA Then
                log("[DROP_SHEETS][GEOM_DELTA] view=" & viewName & " sheet=" & nm &
                    " lines_before=" & bL.ToString(CultureInfo.InvariantCulture) & " lines_after=" & kvp.Value.Lines.ToString(CultureInfo.InvariantCulture) &
                    " arcs_before=" & bA.ToString(CultureInfo.InvariantCulture) & " arcs_after=" & kvp.Value.Arcs.ToString(CultureInfo.InvariantCulture))
            End If
        Next
    End Sub

    Private Class SheetSnapshotOrder
        Public Ordinal As Integer
        Public Name As String
    End Class

    Private Shared Function EnumerateOrderedSheets(draft As DraftDocument) As List(Of SheetSnapshotOrder)
        Dim list As New List(Of SheetSnapshotOrder)()
        Dim n As Integer = SafeCount(CallByNameSafe(draft, "Sheets"))
        For i As Integer = 1 To n
            Dim sh As Sheet = Nothing
            Try : sh = CType(CallByName(draft.Sheets, "Item", CallType.Method, i), Sheet) : Catch : End Try
            If sh Is Nothing Then Continue For
            list.Add(New SheetSnapshotOrder With {.Ordinal = i, .Name = SafeSheetName(sh)})
        Next
        Return list
    End Function

    Private Shared Function EnumerateSheetsRefList(draft As DraftDocument) As List(Of Sheet)
        Dim list As New List(Of Sheet)()
        Dim n As Integer = SafeCount(CallByNameSafe(draft, "Sheets"))
        For i As Integer = 1 To n
            Dim sh As Sheet = Nothing
            Try : sh = CType(CallByName(draft.Sheets, "Item", CallType.Method, i), Sheet) : Catch : End Try
            If sh IsNot Nothing Then list.Add(sh)
        Next
        Return list
    End Function

    Private Shared Function FindNewSheets(before As List(Of Sheet), after As List(Of Sheet)) As List(Of Sheet)
        Dim novel As New List(Of Sheet)()
        For Each sh In after
            If Not before.Contains(sh) Then novel.Add(sh)
        Next
        Return novel
    End Function

    Private Shared Function SheetOrdinal(draft As DraftDocument, target As Sheet) As Integer
        Dim n As Integer = SafeCount(CallByNameSafe(draft, "Sheets"))
        For i As Integer = 1 To n
            Dim sh As Sheet = Nothing
            Try : sh = CType(CallByName(draft.Sheets, "Item", CallType.Method, i), Sheet) : Catch : End Try
            If sh IsNot Nothing AndAlso ReferenceEquals(sh, target) Then Return i
        Next
        Return -1
    End Function

    Private Shared Function CaptureAllFingerprints(draft As DraftDocument) As Dictionary(Of Sheet, SheetEntityFingerprint)
        Dim d As New Dictionary(Of Sheet, SheetEntityFingerprint)()
        Dim n As Integer = SafeCount(CallByNameSafe(draft, "Sheets"))
        For i As Integer = 1 To n
            Dim sh As Sheet = Nothing
            Try : sh = CType(CallByName(draft.Sheets, "Item", CallType.Method, i), Sheet) : Catch : End Try
            If sh Is Nothing Then Continue For
            d(sh) = New SheetEntityFingerprint With {
                .SheetRef = sh,
                .Lines = SafeCount(CallByNameSafe(sh, "Lines2d")),
                .Arcs = SafeCount(CallByNameSafe(sh, "Arcs2d")),
                .Circles = SafeCount(CallByNameSafe(sh, "Circles2d")),
                .BSplines = SafeCount(CallByNameSafe(sh, "BSplineCurves2d")),
                .LineStrings = SafeCount(CallByNameSafe(sh, "LineStrings2d")),
                .Points = SafeCount(CallByNameSafe(sh, "Points2d"))
            }
        Next
        Return d
    End Function

    Private Shared Function ResolveSourceSheetDraft(draft As DraftDocument, log As Action(Of String)) As Sheet
        Dim n As Integer = SafeCount(CallByNameSafe(draft, "Sheets"))
        Dim hoja1 As Sheet = Nothing
        Dim best As Sheet = Nothing
        Dim bestC As Integer = -1
        For i As Integer = 1 To n
            Dim sh As Sheet = Nothing
            Try : sh = CType(CallByName(draft.Sheets, "Item", CallType.Method, i), Sheet) : Catch : End Try
            If sh Is Nothing Then Continue For
            Dim nm As String = SafeSheetName(sh)
            Dim vc As Integer = SafeCount(CallByNameSafe(sh, "DrawingViews"))
            If String.Equals(nm, "Hoja1", StringComparison.OrdinalIgnoreCase) AndAlso vc > 0 Then
                hoja1 = sh
            End If
            If vc > bestC Then
                bestC = vc
                best = sh
            End If
        Next
        If hoja1 IsNot Nothing Then Return hoja1
        Return best
    End Function

    Private Shared Function TryDropView(dv As DrawingView, ByRef methodUsed As String, log As Action(Of String)) As Boolean
        methodUsed = ""
        If dv Is Nothing Then Return False
        Dim names As String() = {"Drop", "DropView", "Break"}
        For Each m In names
            Try
                CallByName(dv, m, CallType.Method)
                methodUsed = m
                Return True
            Catch ex As Exception
                log("[DROP_SHEETS][VIEW][DROP_TRY][FAIL] method=" & m & " msg=" & ex.Message)
            End Try
        Next
        Return False
    End Function

    Private Shared Sub LogLayersOutline(sh As Sheet, log As Action(Of String), sheetName As String)
        Dim layers As Object = CallByNameSafeGet(sh, "Layers")
        Dim n As Integer = SafeCount(layers)
        If n <= 0 Then
            log("[DROP_SHEETS][CREATED_SHEET][LAYERS] sheet=" & sheetName & " count=0 NO_CONFIRMADO")
            Return
        End If
        Dim parts As New List(Of String)()
        For i As Integer = 1 To Math.Min(n, 40)
            Dim ly As Object = Nothing
            Try : ly = CallByName(layers, "Item", CallType.Method, i) : Catch : End Try
            If ly Is Nothing Then Continue For
            parts.Add(SafeStr(CallByNameSafe(ly, "Name")))
        Next
        log("[DROP_SHEETS][CREATED_SHEET][LAYERS] sheet=" & sheetName & " sample=" & String.Join("|", parts))
    End Sub

    Private Shared Function ReadAllLines2d(sh As Sheet, log As Action(Of String), sheetName As String,
                                          ByRef totalKp As Long, ByRef totalCp As Long) As List(Of LineGeom)
        Dim list As New List(Of LineGeom)()
        Dim col As Object = CallByNameSafeGet(sh, "Lines2d")
        Dim n As Integer = SafeCount(col)
        For i As Integer = 1 To n
            Dim ln As Line2d = Nothing
            Try : ln = CType(CallByName(col, "Item", CallType.Method, i), Line2d) : Catch : End Try
            If ln Is Nothing Then Continue For
            Dim x1 As Double = 0, y1 As Double = 0, x2 As Double = 0, y2 As Double = 0
            If Not TryLine2dEndpoints(ln, x1, y1, x2, y2) Then
                log("[DROP_SHEETS][GEOM][LINE][PROP_MISSING] id=" & i.ToString(CultureInfo.InvariantCulture) & " prop=endpoints")
                Continue For
            End If
            Dim dx As Double = Math.Abs(x2 - x1)
            Dim dy As Double = Math.Abs(y2 - y1)
            Dim Llen As Double = Math.Sqrt((x2 - x1) * (x2 - x1) + (y2 - y1) * (y2 - y1))
            Dim ang As Double = Math.Atan2(y2 - y1, x2 - x1) * 180.0R / Math.PI
            Dim tol As Double = Math.Max(TolAxisRatio * Math.Max(Llen, 0.005R), 1.0E-9R)
            Dim orient As String = "inclined"
            If dy <= tol Then orient = "horizontal"
            If dx <= tol Then orient = "vertical"
            log("[DROP_SHEETS][GEOM][LINE] sheet=" & sheetName & " id=" & i.ToString(CultureInfo.InvariantCulture) &
                " x1=" & FormatInv(x1) & " y1=" & FormatInv(y1) & " x2=" & FormatInv(x2) & " y2=" & FormatInv(y2) &
                " length=" & FormatInv(Llen) & " angle=" & FormatInv(ang) & " orientation=" & orient)
            LogLineKeypointsConnect(ln, log, sheetName, i, totalKp, totalCp)
            LogLineMeta(ln, log, sheetName, i)
            list.Add(New LineGeom With {.Idx = i, .Obj = ln, .X1 = x1, .Y1 = y1, .X2 = x2, .Y2 = y2, .Orientation = orient})
        Next
        Return list
    End Function

    Private Shared Sub LogLineKeypointsConnect(ln As Line2d, log As Action(Of String), sheetName As String, lineId As Integer,
                                               ByRef totalKp As Long, ByRef totalCp As Long)
        Dim kc As Integer = 0
        Try
            kc = CInt(CallByName(ln, "KeyPointCount", CallType.Get))
            totalKp += kc
            log("[DROP_SHEETS][GEOM][LINE][KEYPOINTS] sheet=" & sheetName & " id=" & lineId.ToString(CultureInfo.InvariantCulture) &
                " count=" & kc.ToString(CultureInfo.InvariantCulture) & " data=NO_CONFIRMADO")
        Catch
            log("[DROP_SHEETS][GEOM][LINE][PROP_MISSING] id=" & lineId.ToString(CultureInfo.InvariantCulture) & " prop=KeyPointCount")
        End Try
        Dim cp As Object = Nothing
        Try
            cp = CallByName(ln, "ConnectPoints", CallType.Get)
        Catch
            log("[DROP_SHEETS][GEOM][LINE][PROP_MISSING] id=" & lineId.ToString(CultureInfo.InvariantCulture) & " prop=ConnectPoints")
            Return
        End Try
        Dim cc As Integer = SafeCount(cp)
        totalCp += cc
        log("[DROP_SHEETS][GEOM][LINE][CONNECTPOINTS] sheet=" & sheetName & " id=" & lineId.ToString(CultureInfo.InvariantCulture) &
            " count=" & cc.ToString(CultureInfo.InvariantCulture) & " data=NO_CONFIRMADO")
    End Sub

    Private Shared Sub LogLineMeta(ln As Line2d, log As Action(Of String), sheetName As String, lineId As Integer)
        Dim ly As String = ""
        Dim st As String = ""
        Dim par As String = ""
        Try
            Dim lo As Object = CallByName(ln, "Layer", CallType.Get)
            ly = SafeStr(CallByNameSafe(lo, "Name"))
        Catch
            log("[DROP_SHEETS][GEOM][LINE][PROP_MISSING] id=" & lineId.ToString(CultureInfo.InvariantCulture) & " prop=Layer")
        End Try
        Try
            Dim so As Object = CallByName(ln, "Style", CallType.Get)
            st = SafeStr(CallByNameSafe(so, "Name"))
        Catch
            log("[DROP_SHEETS][GEOM][LINE][PROP_MISSING] id=" & lineId.ToString(CultureInfo.InvariantCulture) & " prop=Style")
        End Try
        Try : par = TypeName(CallByName(ln, "Parent", CallType.Get)) : Catch : par = "" : End Try
        log("[DROP_SHEETS][GEOM][LINE][META] sheet=" & sheetName & " id=" & lineId.ToString(CultureInfo.InvariantCulture) &
            " layer=" & ly & " style=" & st & " parent=" & If(String.IsNullOrEmpty(par), "NO_CONFIRMADO", par))
    End Sub

    Private Shared Sub ReadArcsAndCirclesDetailed(sh As Sheet, log As Action(Of String), sheetName As String,
                                                  ByRef totalKp As Long, ByRef totalCp As Long)
        Dim arcs As Object = CallByNameSafeGet(sh, "Arcs2d")
        Dim na As Integer = SafeCount(arcs)
        For i As Integer = 1 To na
            Dim a As Object = Nothing
            Try : a = CallByName(arcs, "Item", CallType.Method, i) : Catch : End Try
            If a Is Nothing Then Continue For
            Dim cx As Double = 0, cy As Double = 0, r As Double = 0, sa As Double = 0, ea As Double = 0
            Try : CallByName(a, "GetCenterPoint", CallType.Method, cx, cy) : Catch : End Try
            Try : r = CDbl(CallByName(a, "Radius", CallType.Get)) : Catch : End Try
            Try : sa = CDbl(CallByName(a, "StartAngle", CallType.Get)) : Catch : sa = Double.NaN : End Try
            Try : ea = CDbl(CallByName(a, "EndAngle", CallType.Get)) : Catch : ea = Double.NaN : End Try
            log("[DROP_SHEETS][GEOM][ARC] sheet=" & sheetName & " id=" & i.ToString(CultureInfo.InvariantCulture) &
                " cx=" & FormatInv(cx) & " cy=" & FormatInv(cy) & " r=" & FormatInv(r) & " startAngle=" & FormatInv(sa) & " endAngle=" & FormatInv(ea))

            Dim sx As Double = 0, sy As Double = 0, ex As Double = 0, ey As Double = 0
            Dim gotS As Boolean = TryArcEndpointReflect(a, "GetStartPoint", sx, sy)
            Dim gotE As Boolean = TryArcEndpointReflect(a, "GetEndPoint", ex, ey)
            If gotS AndAlso gotE Then
                log("[DROP_SHEETS][GEOM][ARC][ENDS] sheet=" & sheetName & " id=" & i.ToString(CultureInfo.InvariantCulture) &
                    " sx=" & FormatInv(sx) & " sy=" & FormatInv(sy) & " ex=" & FormatInv(ex) & " ey=" & FormatInv(ey))
            Else
                log("[DROP_SHEETS][GEOM][ARC][ENDS] sheet=" & sheetName & " id=" & i.ToString(CultureInfo.InvariantCulture) & " NO_CONFIRMADO")
            End If

            Dim ax1 As Double = cx - r, ax2 As Double = cx + r, ay1 As Double = cy - r, ay2 As Double = cy + r
            If gotS AndAlso gotE Then
                ax1 = Math.Min(ax1, Math.Min(sx, ex))
                ax2 = Math.Max(ax2, Math.Max(sx, ex))
                ay1 = Math.Min(ay1, Math.Min(sy, ey))
                ay2 = Math.Max(ay2, Math.Max(sy, ey))
            End If
            log("[DROP_SHEETS][GEOM][ARC][BBOX] sheet=" & sheetName & " id=" & i.ToString(CultureInfo.InvariantCulture) &
                " xmin=" & FormatInv(ax1) & " xmax=" & FormatInv(ax2) & " ymin=" & FormatInv(ay1) & " ymax=" & FormatInv(ay2))

            LogEntityKpCp(a, log, "[DROP_SHEETS][GEOM][ARC][KEYPOINTS]", "[DROP_SHEETS][GEOM][ARC][CONNECTPOINTS]", sheetName, i, totalKp, totalCp)
        Next

        Dim circ As Object = CallByNameSafeGet(sh, "Circles2d")
        Dim nc As Integer = SafeCount(circ)
        For i As Integer = 1 To nc
            Dim c As Object = Nothing
            Try : c = CallByName(circ, "Item", CallType.Method, i) : Catch : End Try
            If c Is Nothing Then Continue For
            Dim cx As Double = 0, cy As Double = 0, r As Double = 0
            Try : CallByName(c, "GetCenterPoint", CallType.Method, cx, cy) : Catch : End Try
            Try : r = CDbl(CallByName(c, "Radius", CallType.Get)) : Catch : End Try
            log("[DROP_SHEETS][GEOM][CIRCLE] sheet=" & sheetName & " id=" & i.ToString(CultureInfo.InvariantCulture) &
                " cx=" & FormatInv(cx) & " cy=" & FormatInv(cy) & " r=" & FormatInv(r))
            Dim qx1 As Double = cx - r, qx2 As Double = cx + r, qy1 As Double = cy - r, qy2 As Double = cy + r
            log("[DROP_SHEETS][GEOM][CIRCLE][QUAD] sheet=" & sheetName & " id=" & i.ToString(CultureInfo.InvariantCulture) &
                " left=(" & FormatInv(qx1) & "," & FormatInv(cy) & ") right=(" & FormatInv(qx2) & "," & FormatInv(cy) &
                ") bottom=(" & FormatInv(cx) & "," & FormatInv(qy1) & ") top=(" & FormatInv(cx) & "," & FormatInv(qy2) & ")")
            LogEntityKpCp(c, log, "[DROP_SHEETS][GEOM][CIRCLE][KEYPOINTS]", "[DROP_SHEETS][GEOM][CIRCLE][CONNECTPOINTS]", sheetName, i, totalKp, totalCp)
        Next

        Dim bsp As Object = CallByNameSafeGet(sh, "BSplineCurves2d")
        Dim nb As Integer = SafeCount(bsp)
        For i As Integer = 1 To nb
            Dim b As Object = Nothing
            Try : b = CallByName(bsp, "Item", CallType.Method, i) : Catch : End Try
            If b Is Nothing Then Continue For
            Dim rx1 As Double = 0, ry1 As Double = 0, rx2 As Double = 0, ry2 As Double = 0
            Dim gotRange As Boolean = False
            Try : CallByName(b, "Range", CallType.Method, rx1, ry1, rx2, ry2) : gotRange = True : Catch : End Try
            If gotRange Then
                log("[DROP_SHEETS][GEOM][BSPLINE][BBOX] sheet=" & sheetName & " id=" & i.ToString(CultureInfo.InvariantCulture) &
                    " xmin=" & FormatInv(Math.Min(rx1, rx2)) & " xmax=" & FormatInv(Math.Max(rx1, rx2)) &
                    " ymin=" & FormatInv(Math.Min(ry1, ry2)) & " ymax=" & FormatInv(Math.Max(ry1, ry2)))
            Else
                log("[DROP_SHEETS][GEOM][BSPLINE] sheet=" & sheetName & " id=" & i.ToString(CultureInfo.InvariantCulture) & " bbox=NO_CONFIRMADO")
            End If
        Next

        Dim lstCol As Object = CallByNameSafeGet(sh, "LineStrings2d")
        Dim nls As Integer = SafeCount(lstCol)
        If nls > 0 Then
            log("[DROP_SHEETS][GEOM][LINESTRING] sheet=" & sheetName & " count=" & nls.ToString(CultureInfo.InvariantCulture) & " detail=NO_CONFIRMADO")
        End If

        Dim pts As Object = CallByNameSafeGet(sh, "Points2d")
        Dim np As Integer = SafeCount(pts)
        If np > 0 Then
            log("[DROP_SHEETS][GEOM][POINTS2d] sheet=" & sheetName & " count=" & np.ToString(CultureInfo.InvariantCulture) & " detail=NO_CONFIRMADO")
        End If
    End Sub

    Private Shared Sub LogEntityKpCp(ent As Object, log As Action(Of String),
                                     kpTag As String, cpTag As String,
                                     sheetName As String, idx As Integer,
                                     ByRef totalKp As Long, ByRef totalCp As Long)
        Try
            Dim k As Integer = CInt(CallByName(ent, "KeyPointCount", CallType.Get))
            totalKp += k
            log(kpTag & " sheet=" & sheetName & " id=" & idx.ToString(CultureInfo.InvariantCulture) & " count=" & k.ToString(CultureInfo.InvariantCulture) & " data=NO_CONFIRMADO")
        Catch
            log("[DROP_SHEETS][GEOM][PROP_MISSING] sheet=" & sheetName & " id=" & idx.ToString(CultureInfo.InvariantCulture) & " prop=KeyPointCount context=" & kpTag)
        End Try
        Dim cpObj As Object = Nothing
        Try : cpObj = CallByName(ent, "ConnectPoints", CallType.Get)
        Catch
            Return
        End Try
        Dim cnt As Integer = SafeCount(cpObj)
        totalCp += cnt
        log(cpTag & " sheet=" & sheetName & " id=" & idx.ToString(CultureInfo.InvariantCulture) & " count=" & cnt.ToString(CultureInfo.InvariantCulture) & " data=NO_CONFIRMADO")
    End Sub

    Private Shared Sub AccumulateBBoxFromLines(lines As List(Of LineGeom), ByRef minX As Double, ByRef maxX As Double, ByRef minY As Double, ByRef maxY As Double)
        For Each lg As LineGeom In lines
            ExpandBBox(lg.X1, lg.Y1, minX, maxX, minY, maxY)
            ExpandBBox(lg.X2, lg.Y2, minX, maxX, minY, maxY)
        Next
    End Sub

    Private Shared Sub AccumulateBBoxFromArcCirclesDraft(sh As Sheet,
                                                        ByRef minX As Double, ByRef maxX As Double, ByRef minY As Double, ByRef maxY As Double,
                                                        log As Action(Of String), sheetName As String)
        Dim arcs As Object = CallByNameSafeGet(sh, "Arcs2d")
        For i As Integer = 1 To SafeCount(arcs)
            Dim a As Object = Nothing
            Try : a = CallByName(arcs, "Item", CallType.Method, i) : Catch : End Try
            If a Is Nothing Then Continue For
            Dim cx As Double = 0, cy As Double = 0, r As Double = 0
            Try : CallByName(a, "GetCenterPoint", CallType.Method, cx, cy) : Catch : End Try
            Try : r = CDbl(CallByName(a, "Radius", CallType.Get)) : Catch : End Try
            ExpandBBox(cx - r, cy - r, minX, maxX, minY, maxY)
            ExpandBBox(cx + r, cy + r, minX, maxX, minY, maxY)
        Next
        Dim circ As Object = CallByNameSafeGet(sh, "Circles2d")
        For i As Integer = 1 To SafeCount(circ)
            Dim c As Object = Nothing
            Try : c = CallByName(circ, "Item", CallType.Method, i) : Catch : End Try
            If c Is Nothing Then Continue For
            Dim cx As Double = 0, cy As Double = 0, r As Double = 0
            Try : CallByName(c, "GetCenterPoint", CallType.Method, cx, cy) : Catch : End Try
            Try : r = CDbl(CallByName(c, "Radius", CallType.Get)) : Catch : End Try
            ExpandBBox(cx - r, cy - r, minX, maxX, minY, maxY)
            ExpandBBox(cx + r, cy + r, minX, maxX, minY, maxY)
        Next
        Dim bsp As Object = CallByNameSafeGet(sh, "BSplineCurves2d")
        For i As Integer = 1 To SafeCount(bsp)
            Dim b As Object = Nothing
            Try : b = CallByName(bsp, "Item", CallType.Method, i) : Catch : End Try
            If b Is Nothing Then Continue For
            Dim rx1 As Double = 0, ry1 As Double = 0, rx2 As Double = 0, ry2 As Double = 0
            Try : CallByName(b, "Range", CallType.Method, rx1, ry1, rx2, ry2)
                ExpandBBox(rx1, ry1, minX, maxX, minY, maxY)
                ExpandBBox(rx2, ry2, minX, maxX, minY, maxY)
            Catch
            End Try
        Next
    End Sub

    Private Shared Sub ExpandBBox(x As Double, y As Double, ByRef minX As Double, ByRef maxX As Double, ByRef minY As Double, ByRef maxY As Double)
        If x < minX Then minX = x
        If x > maxX Then maxX = x
        If y < minY Then minY = y
        If y > maxY Then maxY = y
    End Sub

    Private Shared Sub ClassifyExtents(lines As List(Of LineGeom),
                                       ByRef leftV As LineGeom, ByRef rightV As LineGeom,
                                       ByRef bottomH As LineGeom, ByRef topH As LineGeom,
                                       spanGuess As Double)
        Dim verts As New List(Of LineGeom)()
        Dim hors As New List(Of LineGeom)()
        For Each lg As LineGeom In lines
            If lg Is Nothing Then Continue For
            If lg.Orientation = "vertical" Then verts.Add(lg)
            If lg.Orientation = "horizontal" Then hors.Add(lg)
        Next
        Dim bestLeft As LineGeom = Nothing
        Dim bestRight As LineGeom = Nothing
        Dim minMidX As Double = Double.PositiveInfinity
        Dim maxMidX As Double = Double.NegativeInfinity
        For Each v As LineGeom In verts
            Dim mx As Double = v.Midx()
            If mx < minMidX Then
                minMidX = mx
                bestLeft = v
            End If
            If mx > maxMidX Then
                maxMidX = mx
                bestRight = v
            End If
        Next
        If bestLeft IsNot Nothing AndAlso bestRight IsNot Nothing AndAlso Not ReferenceEquals(bestLeft.Obj, bestRight.Obj) Then
            leftV = bestLeft
            rightV = bestRight
        End If

        Dim bestBot As LineGeom = Nothing
        Dim bestTop As LineGeom = Nothing
        Dim minMidY As Double = Double.PositiveInfinity
        Dim maxMidY As Double = Double.NegativeInfinity
        For Each h As LineGeom In hors
            Dim my As Double = h.Midy()
            If my < minMidY Then
                minMidY = my
                bestBot = h
            End If
            If my > maxMidY Then
                maxMidY = my
                bestTop = h
            End If
        Next
        If bestBot IsNot Nothing AndAlso bestTop IsNot Nothing AndAlso Not ReferenceEquals(bestBot.Obj, bestTop.Obj) Then
            bottomH = bestBot
            topH = bestTop
        End If
    End Sub

    Private Shared Sub LogPartialCandidates(lines As List(Of LineGeom), log As Action(Of String), sheetName As String)
        Dim vertLines As List(Of LineGeom) = lines.Where(Function(l) l.Orientation = "vertical").OrderBy(Function(l) l.Midx()).ToList()
        For i As Integer = 0 To vertLines.Count - 2
            Dim a As LineGeom = vertLines(i), b As LineGeom = vertLines(i + 1)
            If ReferenceEquals(a.Obj, b.Obj) Then Continue For
            log("[DROP_SHEETS][CAND][H_PARTIAL] sheet=" & sheetName & " e1=Line2d e2=Line2d expected=" & FormatInv(Math.Abs(b.Midx() - a.Midx())))
        Next
        Dim hzLines As List(Of LineGeom) = lines.Where(Function(l) l.Orientation = "horizontal").OrderBy(Function(l) l.Midy()).ToList()
        For i As Integer = 0 To hzLines.Count - 2
            Dim a2 As LineGeom = hzLines(i), b2 As LineGeom = hzLines(i + 1)
            log("[DROP_SHEETS][CAND][V_PARTIAL] sheet=" & sheetName & " e1=Line2d e2=Line2d expected=" & FormatInv(Math.Abs(b2.Midy() - a2.Midy())))
        Next
    End Sub

    Private Shared Function TryLine2dEndpoints(ln As Line2d, ByRef x1 As Double, ByRef y1 As Double, ByRef x2 As Double, ByRef y2 As Double) As Boolean
        Try
            Dim args As Object() = New Object() {0.0R, 0.0R, 0.0R, 0.0R}
            ln.GetType().InvokeMember("GetEndPoints", ReflectInvoke, Nothing, ln, args)
            x1 = CDbl(args(0))
            y1 = CDbl(args(1))
            x2 = CDbl(args(2))
            y2 = CDbl(args(3))
            Return True
        Catch
        End Try
        Try
            ln.GetEndPoints(x1, y1, x2, y2)
            Return True
        Catch
        End Try
        Try
            Dim a1 As Object() = {0.0R, 0.0R}
            ln.GetType().InvokeMember("GetStartPoint", ReflectInvoke, Nothing, ln, a1)
            x1 = CDbl(a1(0)) : y1 = CDbl(a1(1))
            Dim a2 As Object() = {0.0R, 0.0R}
            ln.GetType().InvokeMember("GetEndPoint", ReflectInvoke, Nothing, ln, a2)
            x2 = CDbl(a2(0)) : y2 = CDbl(a2(1))
            Return True
        Catch
        End Try
        Return False
    End Function

    Private Shared Function TryArcEndpointReflect(arcObj As Object, methodName As String, ByRef x As Double, ByRef y As Double) As Boolean
        Try
            Dim args As Object() = {0.0R, 0.0R}
            arcObj.GetType().InvokeMember(methodName, ReflectInvoke, Nothing, arcObj, args)
            x = CDbl(args(0))
            y = CDbl(args(1))
            Return True
        Catch
            Return False
        End Try
    End Function

    Private Shared Function EnsureLayerSafe(sh As Sheet, layerName As String, log As Action(Of String)) As Object
        If sh Is Nothing Then Return Nothing
        Dim layers As Object = CallByNameSafeGet(sh, "Layers")
        If layers Is Nothing Then Return Nothing
        Dim n As Integer = SafeCount(layers)
        For i As Integer = 1 To n
            Dim ly As Object = Nothing
            Try : ly = CallByName(layers, "Item", CallType.Method, i) : Catch : End Try
            If ly Is Nothing Then Continue For
            If String.Equals(SafeStr(CallByNameSafe(ly, "Name")), layerName, StringComparison.OrdinalIgnoreCase) Then Return ly
        Next
        Try
            Return CallByName(layers, "Add", CallType.Method, layerName)
        Catch ex As Exception
            log("[DROP_SHEETS][LAYER][WARN] name=" & layerName & " " & ex.Message)
            Return Nothing
        End Try
    End Function

    Private Shared Sub TrySetLayerLate(obj As Object, layerObj As Object)
        If obj Is Nothing OrElse layerObj Is Nothing Then Return
        Try : CallByName(obj, "Layer", CallType.Let, layerObj) : Catch : End Try
    End Sub

    Private Shared Function SafeCount(o As Object) As Integer
        If o Is Nothing Then Return 0
        Try : Return CInt(CallByName(o, "Count", CallType.Get))
        Catch : Return 0
        End Try
    End Function

    Private Shared Function SafeSheetName(sh As Sheet) As String
        If sh Is Nothing Then Return ""
        Try : Return Convert.ToString(sh.Name, CultureInfo.InvariantCulture) : Catch : Return ""
        End Try
    End Function

    Private Shared Function SafeStr(o As Object) As String
        If o Is Nothing Then Return ""
        Try : Return Convert.ToString(o, CultureInfo.InvariantCulture) : Catch : Return ""
        End Try
    End Function

    Private Shared Function FormatInv(v As Double) As String
        If Double.IsNaN(v) Then Return "NaN"
        Return v.ToString("0.######", CultureInfo.InvariantCulture)
    End Function

    Private Shared Function CallByNameSafeGet(obj As Object, prop As String) As Object
        Try : Return CallByName(obj, prop, CallType.Get)
        Catch : Return Nothing
        End Try
    End Function

    Private Shared Function CallByNameSafe(obj As Object, prop As String) As Object
        Return CallByNameSafeGet(obj, prop)
    End Function

End Class
