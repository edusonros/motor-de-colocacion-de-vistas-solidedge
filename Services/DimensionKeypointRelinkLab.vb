Option Strict Off
Option Explicit On

Imports System.Globalization
Imports System.Linq
Imports SolidEdgeDraft
Imports SolidEdgeFrameworkSupport

Public Module DimensionKeypointRelinkLab

    Private Const TargetValueM As Double = 0.34R
    Private Const TargetToleranceM As Double = 0.001R
    Private Const AxisTol As Double = 0.000001R

    Private Enum TriStateLab
        Unknown = 0
        Yes = 1
        No = 2
    End Enum

    Private Structure Pt2
        Public X As Double
        Public Y As Double
    End Structure

    Public Sub Run(
        ByVal draftDoc As DraftDocument,
        ByVal log As Action(Of String),
        Optional ByVal DebugSave As Boolean = False,
        Optional ByVal DebugCleanup As Boolean = True
    )
        If draftDoc Is Nothing Then Return
        If log Is Nothing Then log = Sub(_m As String)
                                     End Sub

        Dim createdForCleanup As New List(Of Object)()
        Dim relinkPossible As TriStateLab = TriStateLab.Unknown
        Dim adboConnected As TriStateLab = TriStateLab.Unknown
        Dim auxDeleteKeeps As TriStateLab = TriStateLab.Unknown
        Dim auxHiddenViable As TriStateLab = TriStateLab.Unknown
        Dim recommended As String = "insufficient_data"

        Try
            log("[DIMRELINK][START] debugSave=" & DebugSave.ToString(CultureInfo.InvariantCulture) &
                " debugCleanup=" & DebugCleanup.ToString(CultureInfo.InvariantCulture))

            Dim sheet As Sheet = ResolveSheetHoja1OrActive(draftDoc, log)
            If sheet Is Nothing Then
                log("[DIMRELINK][ERROR] No se encontró hoja de trabajo.")
                Exit Sub
            End If

            Try
                CallByName(sheet, "Activate", CallType.Method)
            Catch
            End Try

            Dim dims As Dimensions = Nothing
            Try
                dims = CType(sheet.Dimensions, Dimensions)
            Catch ex As Exception
                log("[DIMRELINK][ERROR] Sheet.Dimensions: " & ex.Message)
                Exit Sub
            End Try
            If dims Is Nothing Then
                log("[DIMRELINK][ERROR] Dimensions=Nothing")
                Exit Sub
            End If

            Dim targetDim As Object = FindTargetDimension(dims, log)
            If targetDim Is Nothing Then
                log("[DIMRELINK][TARGET_DIM][MISS] value~0.34m tol=0.001")
                Exit Sub
            End If

            IntrospectDimension(targetDim, log)

            Dim targetView As DrawingView = FindPrimaryOrthogonalView(sheet, log)
            If targetView Is Nothing Then
                log("[DIMRELINK][DV][ERROR] No se encontró DrawingView ortogonal principal.")
                Exit Sub
            End If

            Dim leftLine As Object = Nothing
            Dim rightLine As Object = Nothing
            Dim leftMid As Pt2
            Dim rightMid As Pt2
            Dim leftStart As Pt2
            Dim leftEnd As Pt2
            Dim rightStart As Pt2
            Dim rightEnd As Pt2

            If Not SelectExtremeVerticalLines(targetView, leftLine, rightLine, leftMid, rightMid, leftStart, leftEnd, rightStart, rightEnd, log) Then
                log("[DIMRELINK][DV][ERROR] No se pudieron obtener verticales extremas para test.")
                Exit Sub
            End If

            Dim testAOk As Boolean = RunTestA_RelinkExistingDimension(targetDim, leftLine, rightLine, leftStart, leftEnd, rightStart, rightEnd, leftMid, rightMid, log)
            relinkPossible = If(testAOk, TriStateLab.Yes, TriStateLab.No)

            Dim testBConnected As TriStateLab = RunTestB_AddDistanceBetweenObjects(
                dims, targetView, leftLine, rightLine, leftMid, rightMid, createdForCleanup, log)
            adboConnected = testBConnected

            Dim testCDeleteKeeps As TriStateLab = RunTestC_AuxLineAddLengthDelete(
                draftDoc, sheet, dims, leftMid, rightMid, createdForCleanup, log)
            auxDeleteKeeps = testCDeleteKeeps

            Dim testDHidden As TriStateLab = RunTestD_AuxLineAddLengthHide(
                draftDoc, sheet, dims, leftMid, rightMid, createdForCleanup, log)
            auxHiddenViable = testDHidden

            If relinkPossible = TriStateLab.Yes Then
                recommended = "relink_existing_dimension"
            ElseIf adboConnected = TriStateLab.Yes Then
                recommended = "create_new_dimension_with_AddDistanceBetweenObjects_to_DV"
            ElseIf adboConnected = TriStateLab.No AndAlso auxHiddenViable = TriStateLab.Yes Then
                recommended = "aux_hidden_line_plus_AddLength"
            ElseIf auxDeleteKeeps = TriStateLab.Yes OrElse auxHiddenViable = TriStateLab.Yes Then
                recommended = "manual_editing_aid_only_not_true_associative"
            Else
                recommended = "unknown_need_manual_probe"
            End If

            If DebugSave Then
                Try
                    CallByName(draftDoc, "Save", CallType.Method)
                    log("[DIMRELINK][SAVE] OK")
                Catch ex As Exception
                    log("[DIMRELINK][SAVE][FAIL] " & ex.Message)
                End Try
            Else
                log("[DIMRELINK][SAVE] SKIP debugSave=False")
            End If

            log("[DIMRELINK][SUMMARY] " &
                "existing_dimension_relink_possible=" & ToSummary(relinkPossible) & " " &
                "adbo_to_dv_connected=" & ToSummary(adboConnected) & " " &
                "aux_line_delete_keeps_dimension=" & ToSummary(auxDeleteKeeps) & " " &
                "aux_line_hidden_viable=" & ToSummary(auxHiddenViable) & " " &
                "recommended_strategy=" & recommended)

        Finally
            If DebugCleanup Then
                For i As Integer = createdForCleanup.Count - 1 To 0 Step -1
                    Try
                        CallByName(createdForCleanup(i), "Delete", CallType.Method)
                    Catch
                    End Try
                Next
                Try
                    CallByName(draftDoc, "UpdateAll", CallType.Method, True)
                Catch
                End Try
                log("[DIMRELINK][CLEANUP] done=True")
            Else
                log("[DIMRELINK][CLEANUP] skip=True")
            End If
            log("[DIMRELINK][END]")
        End Try
    End Sub

    Private Function FindTargetDimension(ByVal dims As Dimensions, ByVal log As Action(Of String)) As Object
        Dim n As Integer = 0
        Try
            n = dims.Count
        Catch
            n = 0
        End Try

        For i As Integer = 1 To n
            Dim d As Object = Nothing
            Try
                d = dims.Item(i)
            Catch
                d = Nothing
            End Try
            If d Is Nothing Then Continue For

            Dim val As Double = ReadDimensionValueSafe(d)
            Dim t As String = SafeTypeName(d)
            Dim style As String = SafeStyleName(d)
            log("[DIMRELINK][SCAN_DIM] idx=" & i.ToString(CultureInfo.InvariantCulture) &
                " value=" & val.ToString("G17", CultureInfo.InvariantCulture) &
                " type=" & t & " style=" & style)

            If Not Double.IsNaN(val) AndAlso Math.Abs(val - TargetValueM) <= TargetToleranceM Then
                log("[DIMRELINK][TARGET_DIM] idx=" & i.ToString(CultureInfo.InvariantCulture) &
                    " value=" & val.ToString("G17", CultureInfo.InvariantCulture) &
                    " type=" & t & " style=" & style)
                Return d
            End If
        Next

        Return Nothing
    End Function

    Private Sub IntrospectDimension(ByVal dimObj As Object, ByVal log As Action(Of String))
        Dim methodNames As String() = {
            "GetValueEx", "UpdateStatus", "Range", "RangeBox", "GetRelatedCount", "GetRelated",
            "GetDisplayData", "GetKeyPoint", "SetKeyPoint", "GetTextOffsets", "SetTextOffsets", "SetConnect"
        }
        For Each m In methodNames
            Dim exists As Boolean = ProbeMethodAvailability(dimObj, m)
            log("[DIMRELINK][DIM][METHOD_EXISTS] name=" & m & " exists=" & exists.ToString(CultureInfo.InvariantCulture))
        Next

        Dim props As String() = {"Type", "Value", "StatusOfDimension", "Parent", "Style", "TrackDistance", "AbsoluteTrackDistance"}
        For Each p In props
            Try
                Dim v As Object = CallByName(dimObj, p, CallType.Get)
                log("[DIMRELINK][DIM][PROP] name=" & p & " value=" & FormatOne(v))
            Catch ex As Exception
                log("[DIMRELINK][DIM][PROP] name=" & p & " error=" & ex.Message)
            End Try
        Next

        Try
            Dim status As Object = CallByName(dimObj, "UpdateStatus", CallType.Method)
            log("[DIMRELINK][DIM][STATUS] UpdateStatus=" & FormatOne(status))
        Catch ex As Exception
            log("[DIMRELINK][DIM][STATUS] UpdateStatus_error=" & ex.Message)
        End Try

        Dim relatedCount As Integer = ReadRelatedCount(dimObj)
        log("[DIMRELINK][DIM][RELATED] count=" & relatedCount.ToString(CultureInfo.InvariantCulture))
        If relatedCount > 0 Then
            For i As Integer = 0 To relatedCount + 1
                Try
                    Dim relObj As Object = CallByName(dimObj, "GetRelated", CallType.Method, i)
                    If relObj IsNot Nothing Then
                        log("[DIMRELINK][DIM][RELATED] index=" & i.ToString(CultureInfo.InvariantCulture) &
                            " type=" & SafeTypeName(relObj))
                    End If
                Catch
                End Try
            Next
        End If

        Dim displayData As Object = Nothing
        Try
            displayData = CallByName(dimObj, "GetDisplayData", CallType.Method)
        Catch ex As Exception
            log("[DIMRELINK][DIM][DISPLAYDATA] GetDisplayData_error=" & ex.Message)
        End Try

        If displayData IsNot Nothing Then
            log("[DIMRELINK][DIM][DISPLAYDATA] type=" & SafeTypeName(displayData))
            Dim lc As Integer = ReadCountMethod(displayData, "GetLineCount")
            log("[DIMRELINK][DIM][DISPLAYDATA] lineCount=" & lc.ToString(CultureInfo.InvariantCulture))
            For i As Integer = 1 To Math.Max(0, lc)
                LogDisplayDataAtIndex(displayData, "GetLineAtIndex", "GetLineAtIndex", i, log)
            Next

            Dim ac As Integer = ReadCountMethod(displayData, "GetArcCount")
            log("[DIMRELINK][DIM][DISPLAYDATA] arcCount=" & ac.ToString(CultureInfo.InvariantCulture))
            For i As Integer = 1 To Math.Max(0, ac)
                LogDisplayDataAtIndex(displayData, "GetArcAtIndex", "GetArcAtIndex", i, log)
            Next
        End If
    End Sub

    Private Function FindPrimaryOrthogonalView(ByVal sheet As Sheet, ByVal log As Action(Of String)) As DrawingView
        Dim best As DrawingView = Nothing
        Dim bestScore As Double = Double.MinValue
        For Each v As DrawingView In sheet.DrawingViews
            If IsIsometric(v) Then Continue For
            Dim linesCount As Integer = 0
            Try
                linesCount = v.DVLines2d.Count
            Catch
                linesCount = 0
            End Try
            If linesCount <= 0 Then Continue For
            Dim score As Double = ComputeViewScore(v)
            If score > bestScore Then
                bestScore = score
                best = v
            End If
        Next
        If best IsNot Nothing Then
            log("[DIMRELINK][DV][SELECTED] name=" & SafeViewName(best) & " score=" & bestScore.ToString("G17", CultureInfo.InvariantCulture))
        End If
        Return best
    End Function

    Private Function SelectExtremeVerticalLines(
        ByVal dv As DrawingView,
        ByRef leftLine As Object,
        ByRef rightLine As Object,
        ByRef leftMid As Pt2,
        ByRef rightMid As Pt2,
        ByRef leftStart As Pt2,
        ByRef leftEnd As Pt2,
        ByRef rightStart As Pt2,
        ByRef rightEnd As Pt2,
        ByVal log As Action(Of String)
    ) As Boolean
        Dim minX As Double = Double.PositiveInfinity
        Dim maxX As Double = Double.NegativeInfinity

        For Each ln As Object In dv.DVLines2d
            Dim p1 As Pt2 = GetLineStart(ln)
            Dim p2 As Pt2 = GetLineEnd(ln)
            Dim dx As Double = Math.Abs(p2.X - p1.X)
            Dim dy As Double = Math.Abs(p2.Y - p1.Y)
            Dim kind As String = "oblique"
            If dx <= AxisTol AndAlso dy > AxisTol Then kind = "vertical"
            If dy <= AxisTol AndAlso dx > AxisTol Then kind = "horizontal"
            log("[DIMRELINK][DV][LINE] type=" & kind &
                " x1=" & p1.X.ToString("G17", CultureInfo.InvariantCulture) &
                " y1=" & p1.Y.ToString("G17", CultureInfo.InvariantCulture) &
                " x2=" & p2.X.ToString("G17", CultureInfo.InvariantCulture) &
                " y2=" & p2.Y.ToString("G17", CultureInfo.InvariantCulture))

            If kind <> "vertical" Then Continue For

            Dim xm As Double = (p1.X + p2.X) / 2.0R
            If xm < minX Then
                minX = xm
                leftLine = ln
                leftStart = p1
                leftEnd = p2
            End If
            If xm > maxX Then
                maxX = xm
                rightLine = ln
                rightStart = p1
                rightEnd = p2
            End If
        Next

        If leftLine Is Nothing OrElse rightLine Is Nothing Then Return False
        leftMid = New Pt2 With {.X = (leftStart.X + leftEnd.X) / 2.0R, .Y = (leftStart.Y + leftEnd.Y) / 2.0R}
        rightMid = New Pt2 With {.X = (rightStart.X + rightEnd.X) / 2.0R, .Y = (rightStart.Y + rightEnd.Y) / 2.0R}

        log("[DIMRELINK][DV][EXTREME_LEFT] type=" & SafeTypeName(leftLine) &
            " midX=" & leftMid.X.ToString("G17", CultureInfo.InvariantCulture) &
            " midY=" & leftMid.Y.ToString("G17", CultureInfo.InvariantCulture))
        log("[DIMRELINK][DV][EXTREME_RIGHT] type=" & SafeTypeName(rightLine) &
            " midX=" & rightMid.X.ToString("G17", CultureInfo.InvariantCulture) &
            " midY=" & rightMid.Y.ToString("G17", CultureInfo.InvariantCulture))
        Return True
    End Function

    Private Function RunTestA_RelinkExistingDimension(
        ByVal dimObj As Object,
        ByVal leftLine As Object,
        ByVal rightLine As Object,
        ByVal leftStart As Pt2,
        ByVal leftEnd As Pt2,
        ByVal rightStart As Pt2,
        ByVal rightEnd As Pt2,
        ByVal leftMid As Pt2,
        ByVal rightMid As Pt2,
        ByVal log As Action(Of String)
    ) As Boolean
        Dim beforeSignature As String = BuildRelatedSignature(dimObj)
        Dim anySuccess As Boolean = False
        Dim combos As New List(Of Pt2) From {leftStart, leftEnd, rightStart, rightEnd, leftMid, rightMid}

        For kpIndex As Integer = 0 To 2
            For Each p In combos
                log("[DIMRELINK][TEST_A][TRY] api=SetKeyPoint kpIndex=" & kpIndex.ToString(CultureInfo.InvariantCulture) &
                    " x=" & p.X.ToString("G17", CultureInfo.InvariantCulture) &
                    " y=" & p.Y.ToString("G17", CultureInfo.InvariantCulture))
                If TryInvokeWithVariants(dimObj, "SetKeyPoint",
                    New Object() {kpIndex, p.X, p.Y, 0.0R},
                    New Object() {kpIndex, p.X, p.Y},
                    New Object() {kpIndex, leftLine, p.X, p.Y, 0.0R},
                    New Object() {kpIndex, rightLine, p.X, p.Y, 0.0R}) Then
                    anySuccess = True
                    log("[DIMRELINK][TEST_A][OK] api=SetKeyPoint kpIndex=" & kpIndex.ToString(CultureInfo.InvariantCulture))
                Else
                    log("[DIMRELINK][TEST_A][FAIL] api=SetKeyPoint kpIndex=" & kpIndex.ToString(CultureInfo.InvariantCulture))
                End If
            Next
        Next

        If ProbeMethodAvailability(dimObj, "SetConnect") Then
            log("[DIMRELINK][TEST_A][TRY] api=SetConnect")
            If TryInvokeWithVariants(dimObj, "SetConnect",
                New Object() {0, leftLine},
                New Object() {1, rightLine},
                New Object() {leftLine, rightLine}) Then
                anySuccess = True
                log("[DIMRELINK][TEST_A][OK] api=SetConnect")
            Else
                log("[DIMRELINK][TEST_A][FAIL] api=SetConnect")
            End If
        End If

        Dim statusAfter As String = ReadStatusText(dimObj)
        Dim relatedAfter As String = BuildRelatedSignature(dimObj)
        log("[DIMRELINK][TEST_A][STATUS_AFTER] " & statusAfter)
        log("[DIMRELINK][TEST_A][RELATED_AFTER] " & relatedAfter)

        Dim changed As Boolean = (beforeSignature <> relatedAfter) AndAlso (relatedAfter.Length > 0)
        If changed Then
            log("[DIMRELINK][RESULT] relink_possible=True")
        End If
        Return anySuccess AndAlso changed
    End Function

    Private Function RunTestB_AddDistanceBetweenObjects(
        ByVal dims As Dimensions,
        ByVal dv As DrawingView,
        ByVal leftLine As Object,
        ByVal rightLine As Object,
        ByVal leftMid As Pt2,
        ByVal rightMid As Pt2,
        ByVal createdForCleanup As List(Of Object),
        ByVal log As Action(Of String)
    ) As TriStateLab
        log("[DIMRELINK][TEST_B][ADBO][TRY]")
        Dim placement As Pt2 = New Pt2 With {
            .X = (leftMid.X + rightMid.X) / 2.0R,
            .Y = Math.Max(leftMid.Y, rightMid.Y) + Math.Abs(rightMid.X - leftMid.X) * 0.15R + 0.01R
        }

        Dim d As Object = Nothing
        Try
            d = dims.AddDistanceBetweenObjects(
                leftLine, leftMid.X, leftMid.Y, 0.0R, False,
                rightLine, rightMid.X, rightMid.Y, 0.0R, False
            )
        Catch ex As Exception
            log("[DIMRELINK][TEST_B][ADBO][FAIL] variant=basic msg=" & ex.Message)
        End Try

        If d Is Nothing Then
            Try
                d = CallByName(dims, "AddDistanceBetweenObjects", CallType.Method,
                    leftLine, leftMid.X, leftMid.Y, 0.0R, True,
                    rightLine, rightMid.X, rightMid.Y, 0.0R, True)
            Catch ex As Exception
                log("[DIMRELINK][TEST_B][ADBO][FAIL] variant=keypoint_true msg=" & ex.Message)
            End Try
        End If

        If d Is Nothing Then
            Return TriStateLab.No
        End If

        createdForCleanup.Add(d)
        log("[DIMRELINK][TEST_B][ADBO][OK] type=" & SafeTypeName(d))

        Try
            CallByName(d, "SetKeyPoint", CallType.Method, 0, placement.X, placement.Y, 0.0R)
        Catch
        End Try

        Dim value As Double = ReadDimensionValueSafe(d)
        log("[DIMRELINK][TEST_B][ADBO][VALUE] value=" & value.ToString("G17", CultureInfo.InvariantCulture))
        log("[DIMRELINK][TEST_B][ADBO][STATUS] " & ReadStatusText(d))

        Dim rc As Integer = ReadRelatedCount(d)
        log("[DIMRELINK][TEST_B][ADBO][RELATED] count=" & rc.ToString(CultureInfo.InvariantCulture))
        Dim sig As String = BuildRelatedSignature(d)
        log("[DIMRELINK][TEST_B][ADBO][RELATED] sig=" & sig)

        ' Variante inspirada en muestras SE: AddLength(DVLine2d.Reference)
        Dim dLeftLen As Object = TryAddLengthOnObject(dims, leftLine, log, "TEST_B_ADDLENGTH_LEFT_REF")
        If dLeftLen IsNot Nothing Then
            createdForCleanup.Add(dLeftLen)
            log("[DIMRELINK][TEST_B][ADDLENGTH_REF][LEFT] value=" & ReadDimensionValueSafe(dLeftLen).ToString("G17", CultureInfo.InvariantCulture) &
                " relatedCount=" & ReadRelatedCount(dLeftLen).ToString(CultureInfo.InvariantCulture))
        End If
        Dim dRightLen As Object = TryAddLengthOnObject(dims, rightLine, log, "TEST_B_ADDLENGTH_RIGHT_REF")
        If dRightLen IsNot Nothing Then
            createdForCleanup.Add(dRightLen)
            log("[DIMRELINK][TEST_B][ADDLENGTH_REF][RIGHT] value=" & ReadDimensionValueSafe(dRightLen).ToString("G17", CultureInfo.InvariantCulture) &
                " relatedCount=" & ReadRelatedCount(dRightLen).ToString(CultureInfo.InvariantCulture))
        End If

        If rc >= 2 AndAlso IsDimensionConnectedToView(dv, d) Then
            Return TriStateLab.Yes
        End If
        If rc > 0 Then Return TriStateLab.Unknown
        Return TriStateLab.No
    End Function

    Private Function RunTestC_AuxLineAddLengthDelete(
        ByVal draftDoc As DraftDocument,
        ByVal sheet As Sheet,
        ByVal dims As Dimensions,
        ByVal p1 As Pt2,
        ByVal p2 As Pt2,
        ByVal createdForCleanup As List(Of Object),
        ByVal log As Action(Of String)
    ) As TriStateLab
        Dim auxLine As Object = CreateAuxLine(sheet, p1, p2, log, "TEST_C")
        If auxLine Is Nothing Then Return TriStateLab.No
        createdForCleanup.Add(auxLine)
        log("[DIMRELINK][TEST_C][AUX_LINE_CREATE] ok=True type=" & SafeTypeName(auxLine))

        Dim d As Object = TryAddLengthOnObject(dims, auxLine, log, "TEST_C")
        If d Is Nothing Then Return TriStateLab.No
        createdForCleanup.Add(d)
        log("[DIMRELINK][TEST_C][ADDLENGTH_OK] value=" & ReadDimensionValueSafe(d).ToString("G17", CultureInfo.InvariantCulture))

        log("[DIMRELINK][TEST_C][DELETE_AUX_LINE] try=True")
        Try
            CallByName(auxLine, "Delete", CallType.Method)
        Catch ex As Exception
            log("[DIMRELINK][TEST_C][DELETE_AUX_LINE] error=" & ex.Message)
            Return TriStateLab.No
        End Try

        Try
            CallByName(draftDoc, "UpdateAll", CallType.Method, True)
        Catch
        End Try

        Dim statusAfter As String = ReadStatusText(d)
        Dim relatedAfter As String = BuildRelatedSignature(d)
        Dim valueAfter As Double = ReadDimensionValueSafe(d)
        Dim stillAlive As Boolean = IsDimensionComAlive(d)
        If statusAfter.IndexOf("RPC_E_DISCONNECTED", StringComparison.OrdinalIgnoreCase) >= 0 Then stillAlive = False

        log("[DIMRELINK][TEST_C][STATUS_AFTER_DELETE] status=" & statusAfter)
        log("[DIMRELINK][TEST_C][RELATED_AFTER_DELETE] sig=" & relatedAfter)
        log("[DIMRELINK][TEST_C][VALUE_AFTER_DELETE] value=" & valueAfter.ToString("G17", CultureInfo.InvariantCulture) &
            " comAlive=" & stillAlive.ToString(CultureInfo.InvariantCulture))

        If stillAlive AndAlso Not Double.IsNaN(valueAfter) Then
            Return TriStateLab.Yes
        End If
        Return TriStateLab.No
    End Function

    Private Function RunTestD_AuxLineAddLengthHide(
        ByVal draftDoc As DraftDocument,
        ByVal sheet As Sheet,
        ByVal dims As Dimensions,
        ByVal p1 As Pt2,
        ByVal p2 As Pt2,
        ByVal createdForCleanup As List(Of Object),
        ByVal log As Action(Of String)
    ) As TriStateLab
        Dim auxLine As Object = CreateAuxLine(sheet, p1, p2, log, "TEST_D")
        If auxLine Is Nothing Then Return TriStateLab.No
        createdForCleanup.Add(auxLine)

        Dim d As Object = TryAddLengthOnObject(dims, auxLine, log, "TEST_D")
        If d Is Nothing Then Return TriStateLab.No
        createdForCleanup.Add(d)

        log("[DIMRELINK][TEST_D][HIDE_AUX_LINE] try=True")
        Dim hidden As Boolean = False
        If TrySetProperty(auxLine, "Visible", False) Then hidden = True
        If TrySetProperty(auxLine, "Display", False) Then hidden = True
        If TrySetProperty(auxLine, "Show", False) Then hidden = True
        If TryMoveToHiddenLayer(sheet, auxLine) Then hidden = True

        Try
            CallByName(draftDoc, "UpdateAll", CallType.Method, True)
        Catch
        End Try

        Dim rc As Integer = ReadRelatedCount(d)
        Dim stable As Boolean = (rc > 0)
        log("[DIMRELINK][TEST_D][STATUS] hiddenApplied=" & hidden.ToString(CultureInfo.InvariantCulture) &
            " relatedCount=" & rc.ToString(CultureInfo.InvariantCulture) &
            " status=" & ReadStatusText(d))

        If hidden AndAlso stable Then Return TriStateLab.Yes
        If hidden Then Return TriStateLab.Unknown
        Return TriStateLab.No
    End Function

    Private Function CreateAuxLine(ByVal sheet As Sheet, ByVal p1 As Pt2, ByVal p2 As Pt2, ByVal log As Action(Of String), ByVal tag As String) As Object
        Dim lines As Object = Nothing
        Try
            lines = CallByName(sheet, "Lines2d", CallType.Get)
        Catch ex As Exception
            log("[DIMRELINK][" & tag & "][AUX_LINE_CREATE][FAIL] Lines2d: " & ex.Message)
            Return Nothing
        End Try
        If lines Is Nothing Then Return Nothing

        Try
            Return CallByName(lines, "AddBy2Points", CallType.Method, p1.X, p1.Y, p2.X, p2.Y)
        Catch
        End Try
        Try
            Return CallByName(lines, "AddLine", CallType.Method, p1.X, p1.Y, p2.X, p2.Y)
        Catch ex As Exception
            log("[DIMRELINK][" & tag & "][AUX_LINE_CREATE][FAIL] Add method: " & ex.Message)
            Return Nothing
        End Try
    End Function

    Private Function TryAddLengthOnObject(ByVal dims As Dimensions, ByVal srcObj As Object, ByVal log As Action(Of String), ByVal tag As String) As Object
        Dim refObj As Object = Nothing
        Try
            refObj = CallByName(srcObj, "Reference", CallType.Get)
        Catch
            refObj = srcObj
        End Try

        Try
            Return CallByName(dims, "AddLength", CallType.Method, refObj)
        Catch ex As Exception
            log("[DIMRELINK][" & tag & "][ADDLENGTH_FAIL] " & ex.Message)
            Return Nothing
        End Try
    End Function

    Private Function ResolveSheetHoja1OrActive(ByVal draftDoc As DraftDocument, ByVal log As Action(Of String)) As Sheet
        Try
            Dim sheets As Sheets = draftDoc.Sheets
            For i As Integer = 1 To sheets.Count
                Dim sh As Sheet = CType(sheets.Item(i), Sheet)
                Dim n As String = SafeSheetName(sh)
                If String.Equals(n, "Hoja1", StringComparison.OrdinalIgnoreCase) Then
                    log("[DIMRELINK][SHEET] using=Hoja1")
                    Return sh
                End If
            Next
        Catch
        End Try
        Try
            Dim active As Sheet = draftDoc.ActiveSheet
            log("[DIMRELINK][SHEET] using=ActiveSheet name=" & SafeSheetName(active))
            Return active
        Catch ex As Exception
            log("[DIMRELINK][SHEET][ERROR] " & ex.Message)
            Return Nothing
        End Try
    End Function

    Private Function ComputeViewScore(ByVal v As DrawingView) As Double
        Dim minX As Double = Double.PositiveInfinity
        Dim maxX As Double = Double.NegativeInfinity
        Dim minY As Double = Double.PositiveInfinity
        Dim maxY As Double = Double.NegativeInfinity
        For Each ln As Object In v.DVLines2d
            Dim p1 As Pt2 = GetLineStart(ln)
            Dim p2 As Pt2 = GetLineEnd(ln)
            minX = Math.Min(minX, Math.Min(p1.X, p2.X))
            maxX = Math.Max(maxX, Math.Max(p1.X, p2.X))
            minY = Math.Min(minY, Math.Min(p1.Y, p2.Y))
            maxY = Math.Max(maxY, Math.Max(p1.Y, p2.Y))
        Next
        If Double.IsInfinity(minX) Then Return Double.MinValue
        Return (maxX - minX) * (maxY - minY)
    End Function

    Private Function IsIsometric(ByVal v As DrawingView) As Boolean
        Try
            Dim ori As SolidEdgeConstants.ViewOrientationConstants = SolidEdgeConstants.ViewOrientationConstants.igFrontView
            Dim vx As Double = 0, vy As Double = 0, vz As Double = 0, lx As Double = 0, ly As Double = 0, lz As Double = 0
            v.ViewOrientation(vx, vy, vz, lx, ly, lz, ori)
            Return ori = SolidEdgeConstants.ViewOrientationConstants.igTopFrontRightView
        Catch
            Return False
        End Try
    End Function

    Private Function IsDimensionConnectedToView(ByVal dv As DrawingView, ByVal dimObj As Object) As Boolean
        If dv Is Nothing OrElse dimObj Is Nothing Then Return False
        Dim ss As Object = Nothing
        Try
            ss = CallByName(CallByName(dv, "Application", CallType.Get), "ActiveSelectSet", CallType.Get)
        Catch
            ss = Nothing
        End Try
        If ss Is Nothing Then Return False

        Try
            CallByName(ss, "RemoveAll", CallType.Method)
        Catch
        End Try
        Try
            CallByName(dv, "AddConnectedDimensionsToSelectSet", CallType.Method)
        Catch
            Return False
        End Try

        Dim n As Integer = 0
        Try
            n = Convert.ToInt32(CallByName(ss, "Count", CallType.Get), CultureInfo.InvariantCulture)
        Catch
            n = 0
        End Try
        For i As Integer = 1 To n
            Try
                Dim it As Object = CallByName(ss, "Item", CallType.Method, i)
                If Object.ReferenceEquals(it, dimObj) Then Return True
            Catch
            End Try
        Next
        Return False
    End Function

    Private Function TryMoveToHiddenLayer(ByVal sheet As Sheet, ByVal obj As Object) As Boolean
        Dim layers As Object = Nothing
        Try
            layers = CallByName(sheet, "Layers", CallType.Get)
        Catch
            layers = Nothing
        End Try
        If layers Is Nothing Then Return False

        Dim n As Integer = 0
        Try
            n = Convert.ToInt32(CallByName(layers, "Count", CallType.Get), CultureInfo.InvariantCulture)
        Catch
            n = 0
        End Try
        For i As Integer = 1 To n
            Dim lyr As Object = Nothing
            Try
                lyr = CallByName(layers, "Item", CallType.Method, i)
            Catch
                lyr = Nothing
            End Try
            If lyr Is Nothing Then Continue For
            Dim name As String = ""
            Try
                name = Convert.ToString(CallByName(lyr, "Name", CallType.Get), CultureInfo.InvariantCulture)
            Catch
                name = ""
            End Try
            Dim low As String = name.ToLowerInvariant()
            If low.Contains("hidden") OrElse low.Contains("ocult") OrElse low.Contains("no print") OrElse low.Contains("noprint") Then
                Try
                    CallByName(obj, "Layer", CallType.Let, lyr)
                    Return True
                Catch
                End Try
            End If
        Next
        Return False
    End Function

    Private Function TrySetProperty(ByVal obj As Object, ByVal propName As String, ByVal value As Object) As Boolean
        Try
            CallByName(obj, propName, CallType.Let, value)
            Return True
        Catch
            Return False
        End Try
    End Function

    Private Function ReadRelatedCount(ByVal dimObj As Object) As Integer
        Try
            Return Convert.ToInt32(CallByName(dimObj, "GetRelatedCount", CallType.Method), CultureInfo.InvariantCulture)
        Catch
            Return 0
        End Try
    End Function

    Private Function BuildRelatedSignature(ByVal dimObj As Object) As String
        Dim count As Integer = ReadRelatedCount(dimObj)
        If count <= 0 Then Return ""
        Dim chunks As New List(Of String)()
        For i As Integer = 0 To count + 1
            Try
                Dim o As Object = CallByName(dimObj, "GetRelated", CallType.Method, i)
                If o IsNot Nothing Then
                    chunks.Add(i.ToString(CultureInfo.InvariantCulture) & ":" & SafeTypeName(o))
                End If
            Catch
            End Try
        Next
        Return String.Join("|", chunks)
    End Function

    Private Function ReadStatusText(ByVal dimObj As Object) As String
        Try
            Return "status=" & FormatOne(CallByName(dimObj, "StatusOfDimension", CallType.Get))
        Catch
            Try
                Return "status=" & FormatOne(CallByName(dimObj, "UpdateStatus", CallType.Method))
            Catch ex As Exception
                Return "status_error=" & ex.Message
            End Try
        End Try
    End Function

    Private Function ReadCountMethod(ByVal obj As Object, ByVal methodName As String) As Integer
        Try
            Dim v As Object = CallByName(obj, methodName, CallType.Method)
            If v Is Nothing Then Return 0
            Return Convert.ToInt32(v, CultureInfo.InvariantCulture)
        Catch
            Return 0
        End Try
    End Function

    Private Function ReadDimensionValueSafe(ByVal dimObj As Object) As Double
        If dimObj Is Nothing Then Return Double.NaN
        Try
            Dim v As Object = CallByName(dimObj, "Value", CallType.Get)
            Return Convert.ToDouble(v, CultureInfo.InvariantCulture)
        Catch
        End Try
        Try
            Dim v1 As Object = CallByName(dimObj, "GetValueEx", CallType.Method)
            Return Convert.ToDouble(v1, CultureInfo.InvariantCulture)
        Catch
        End Try
        Try
            Dim v2 As Object = CallByName(dimObj, "GetValueEx", CallType.Method, 0)
            Return Convert.ToDouble(v2, CultureInfo.InvariantCulture)
        Catch
        End Try
        Return Double.NaN
    End Function

    Private Function IsDimensionComAlive(ByVal dimObj As Object) As Boolean
        If dimObj Is Nothing Then Return False
        Try
            Dim _status As Object = CallByName(dimObj, "StatusOfDimension", CallType.Get)
            Return True
        Catch ex As Exception
            If ex.Message.IndexOf("RPC_E_DISCONNECTED", StringComparison.OrdinalIgnoreCase) >= 0 Then Return False
        End Try
        Try
            Dim _u As Object = CallByName(dimObj, "UpdateStatus", CallType.Method)
            Return True
        Catch ex As Exception
            If ex.Message.IndexOf("RPC_E_DISCONNECTED", StringComparison.OrdinalIgnoreCase) >= 0 Then Return False
        End Try
        Return False
    End Function

    Private Function ProbeMethodAvailability(ByVal target As Object, ByVal methodName As String) As Boolean
        If target Is Nothing Then Return False
        If HasMethod(target, methodName) Then Return True

        Select Case methodName
            Case "GetValueEx"
                Return TryInvokeWithVariants(target, methodName, New Object() {}, New Object() {0})
            Case "UpdateStatus", "GetDisplayData", "GetRelatedCount"
                Return TryInvokeWithVariants(target, methodName, New Object() {})
            Case "GetRelated"
                Return TryInvokeWithVariants(target, methodName, New Object() {0}, New Object() {1})
            Case "GetKeyPoint"
                Return TryInvokeWithVariants(target, methodName, New Object() {0, 0.0R, 0.0R, 0.0R, 0, 0})
            Case "SetKeyPoint"
                Return TryInvokeWithVariants(target, methodName, New Object() {0, 0.0R, 0.0R, 0.0R}, New Object() {0, 0.0R, 0.0R})
            Case "GetTextOffsets"
                Return TryInvokeWithVariants(target, methodName, New Object() {0.0R, 0.0R, 0.0R})
            Case "SetTextOffsets"
                Return TryInvokeWithVariants(target, methodName, New Object() {0.0R, 0.0R, 0.0R}, New Object() {0.0R, 0.0R})
            Case "SetConnect"
                Return TryInvokeWithVariants(target, methodName, New Object() {0, Nothing}, New Object() {Nothing, Nothing})
            Case "Range", "RangeBox"
                Return TryInvokeWithVariants(target, methodName, New Object() {0.0R, 0.0R, 0.0R, 0.0R})
            Case Else
                Return False
        End Select
    End Function

    Private Sub LogDisplayDataAtIndex(
        ByVal displayData As Object,
        ByVal methodName As String,
        ByVal label As String,
        ByVal idx As Integer,
        ByVal log As Action(Of String)
    )
        Dim outValues As Object() = Nothing
        Dim signature As String = ""
        Dim returnValue As Object = Nothing
        Dim err As String = ""
        If TryInvokeDisplayDataMethod(displayData, methodName, idx, outValues, signature, returnValue, err) Then
            Dim merged As Object() = MergeOutputs(outValues, returnValue)
            log("[DIMRELINK][DIM][DISPLAYDATA] " & label & " idx=" & idx.ToString(CultureInfo.InvariantCulture) &
                " sig=" & signature & " raw=" & FormatRawValues(merged))
        Else
            log("[DIMRELINK][DIM][DISPLAYDATA] " & label & " idx=" & idx.ToString(CultureInfo.InvariantCulture) &
                " error=" & err)
        End If
    End Sub

    Private Function TryInvokeDisplayDataMethod(
        ByVal dd As Object,
        ByVal methodName As String,
        ByVal atIndex As Integer?,
        ByRef outValues As Object(),
        ByRef signature As String,
        ByRef returnValue As Object,
        ByRef errorMessage As String
    ) As Boolean
        outValues = Array.Empty(Of Object)()
        signature = ""
        returnValue = Nothing
        errorMessage = "No overload found."

        Try
            Dim methods = dd.GetType().GetMethods().Where(Function(m) m.Name = methodName).ToArray()
            If methods Is Nothing OrElse methods.Length = 0 Then
                errorMessage = "Método no encontrado."
                Return False
            End If

            For Each mi In methods
                Try
                    Dim ps = mi.GetParameters()
                    Dim args(ps.Length - 1) As Object
                    For i As Integer = 0 To ps.Length - 1
                        Dim pType As Type = ps(i).ParameterType
                        Dim baseType As Type = If(pType.IsByRef, pType.GetElementType(), pType)
                        If i = 0 AndAlso atIndex.HasValue Then
                            args(i) = Convert.ChangeType(atIndex.Value, baseType, CultureInfo.InvariantCulture)
                        Else
                            args(i) = CreateDefaultForType(baseType)
                        End If
                    Next

                    returnValue = mi.Invoke(dd, args)
                    signature = mi.ToString()
                    If atIndex.HasValue Then
                        outValues = args.Skip(1).ToArray()
                    Else
                        outValues = args
                    End If
                    Return True
                Catch exOne As Exception
                    errorMessage = exOne.Message
                End Try
            Next
        Catch ex As Exception
            errorMessage = ex.Message
        End Try

        Return False
    End Function

    Private Function CreateDefaultForType(ByVal t As Type) As Object
        If t Is GetType(String) Then Return ""
        If t Is GetType(Boolean) Then Return False
        If t.IsValueType Then Return Activator.CreateInstance(t)
        Return Nothing
    End Function

    Private Function MergeOutputs(ByVal outValues As Object(), ByVal returnValue As Object) As Object()
        If returnValue Is Nothing Then
            Return If(outValues, Array.Empty(Of Object)())
        End If
        Dim baseVals As Object() = If(outValues, Array.Empty(Of Object)())
        Dim result(baseVals.Length) As Object
        For i As Integer = 0 To baseVals.Length - 1
            result(i) = baseVals(i)
        Next
        result(baseVals.Length) = returnValue
        Return result
    End Function

    Private Function FormatRawValues(ByVal vals As Object()) As String
        If vals Is Nothing OrElse vals.Length = 0 Then Return "<empty>"
        Dim parts As New List(Of String)
        For i As Integer = 0 To vals.Length - 1
            parts.Add("v" & (i + 1).ToString(CultureInfo.InvariantCulture) & "=" & FormatOne(vals(i)))
        Next
        Return String.Join(" ", parts)
    End Function

    Private Function TryInvokeWithVariants(ByVal target As Object, ByVal methodName As String, ParamArray variants As Object()) As Boolean
        For Each rawArgs As Object In variants
            Dim args As Object() = TryCast(rawArgs, Object())
            If args Is Nothing Then
                args = New Object() {rawArgs}
            End If
            Try
                CallByName(target, methodName, CallType.Method, args)
                Return True
            Catch
            End Try
        Next
        Return False
    End Function

    Private Function HasMethod(ByVal obj As Object, ByVal methodName As String) As Boolean
        If obj Is Nothing Then Return False
        Try
            Return obj.GetType().GetMethods().Any(Function(m) String.Equals(m.Name, methodName, StringComparison.OrdinalIgnoreCase))
        Catch
            Return False
        End Try
    End Function

    Private Function GetLineStart(ByVal lineObj As Object) As Pt2
        Dim x As Double = 0.0R
        Dim y As Double = 0.0R
        Try
            lineObj.GetStartPoint(x, y)
        Catch
        End Try
        Return New Pt2 With {.X = x, .Y = y}
    End Function

    Private Function GetLineEnd(ByVal lineObj As Object) As Pt2
        Dim x As Double = 0.0R
        Dim y As Double = 0.0R
        Try
            lineObj.GetEndPoint(x, y)
        Catch
        End Try
        Return New Pt2 With {.X = x, .Y = y}
    End Function

    Private Function SafeTypeName(ByVal o As Object) As String
        If o Is Nothing Then Return "Nothing"
        Try
            Return o.GetType().Name
        Catch
            Return "?"
        End Try
    End Function

    Private Function SafeViewName(ByVal v As DrawingView) As String
        If v Is Nothing Then Return "?"
        Try
            Return v.Name
        Catch
            Return "?"
        End Try
    End Function

    Private Function SafeSheetName(ByVal s As Sheet) As String
        If s Is Nothing Then Return "?"
        Try
            Return s.Name
        Catch
            Return "?"
        End Try
    End Function

    Private Function SafeStyleName(ByVal dimObj As Object) As String
        Try
            Dim st As Object = CallByName(dimObj, "Style", CallType.Get)
            If st Is Nothing Then Return "<none>"
            Try
                Return Convert.ToString(CallByName(st, "Name", CallType.Get), CultureInfo.InvariantCulture)
            Catch
                Return Convert.ToString(st, CultureInfo.InvariantCulture)
            End Try
        Catch
            Return "<err>"
        End Try
    End Function

    Private Function FormatOne(ByVal value As Object) As String
        If value Is Nothing Then Return "Nothing"
        If TypeOf value Is Double Then Return DirectCast(value, Double).ToString("G17", CultureInfo.InvariantCulture)
        If TypeOf value Is Single Then Return DirectCast(value, Single).ToString("G9", CultureInfo.InvariantCulture)
        If TypeOf value Is IFormattable Then Return DirectCast(value, IFormattable).ToString(Nothing, CultureInfo.InvariantCulture)
        Return Convert.ToString(value, CultureInfo.InvariantCulture)
    End Function

    Private Function ToSummary(ByVal v As TriStateLab) As String
        Select Case v
            Case TriStateLab.Yes
                Return "True"
            Case TriStateLab.No
                Return "False"
            Case Else
                Return "Unknown"
        End Select
    End Function
End Module
