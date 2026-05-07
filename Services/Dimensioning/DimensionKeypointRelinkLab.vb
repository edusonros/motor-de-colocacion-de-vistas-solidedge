Option Strict Off
Option Explicit On

Imports System.Globalization
Imports System.Linq
Imports SolidEdgeDraft
Imports SolidEdgeFrameworkSupport

Public Module DimensionKeypointRelinkLabV2
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

    Public Sub Run(ByVal draftDoc As DraftDocument, ByVal log As Action(Of String), Optional ByVal DebugSave As Boolean = False, Optional ByVal DebugCleanup As Boolean = True)
        If draftDoc Is Nothing Then Return
        If log Is Nothing Then log = Sub(_m As String)
                                     End Sub

        Dim createdForCleanup As New List(Of Object)()
        Dim existingRelink As TriStateLab = TriStateLab.Unknown
        Dim adboConnected As TriStateLab = TriStateLab.Unknown
        Dim adboBehaviorConnected As TriStateLab = TriStateLab.Unknown
        Dim auxDeleteKeeps As TriStateLab = TriStateLab.Unknown
        Dim auxHiddenViable As TriStateLab = TriStateLab.Unknown

        log("[DIMLAB][START] debugSave=" & DebugSave.ToString(CultureInfo.InvariantCulture) & " debugCleanup=" & DebugCleanup.ToString(CultureInfo.InvariantCulture))
        Try
            Dim sheet As Sheet = ResolveSheetHoja1OrActive(draftDoc)
            If sheet Is Nothing Then
                log("[DIMLAB][ERROR] no_sheet")
                Exit Sub
            End If
            Try : CallByName(sheet, "Activate", CallType.Method) : Catch : End Try

            Dim dims As Dimensions = CType(sheet.Dimensions, Dimensions)
            ForensicAnalyzeAllDimensions(dims, sheet, log)
            Dim targetDim As Object = FindTargetDimension(dims, log)
            If targetDim Is Nothing Then
                log("[DIMLAB][TARGET_DIM] exact_match=False -> fallback=nearest")
                targetDim = FindNearestDimensionByValue(dims, TargetValueM, log)
            End If
            If targetDim IsNot Nothing Then
                IntrospectDimension(targetDim, log)
            Else
                log("[DIMLAB][TARGET_DIM] fallback_not_found=True (se omite TEST_A y se continúa)")
            End If

            Dim dv As DrawingView = FindPrimaryOrthogonalView(sheet, log)
            If dv Is Nothing Then
                log("[DIMLAB][DV][VIEW] missing=True")
                Exit Sub
            End If

            Dim leftLine As Object = Nothing, rightLine As Object = Nothing, topLine As Object = Nothing, bottomLine As Object = Nothing
            Dim leftMid As Pt2, rightMid As Pt2, leftStart As Pt2, leftEnd As Pt2, rightStart As Pt2, rightEnd As Pt2
            If Not SelectExtremes(dv, leftLine, rightLine, topLine, bottomLine, leftMid, rightMid, leftStart, leftEnd, rightStart, rightEnd, log) Then
                log("[DIMLAB][DV][VIEW] no_extremes=True")
                Exit Sub
            End If

            If targetDim IsNot Nothing Then
                existingRelink = If(RunTestA(targetDim, leftLine, rightLine, leftStart, leftEnd, rightStart, rightEnd, leftMid, rightMid, draftDoc, log), TriStateLab.Yes, TriStateLab.No)
            Else
                existingRelink = TriStateLab.Unknown
                log("[DIMLAB][TEST_A][SKIP] reason=no_target_dimension")
            End If
            adboConnected = RunTestB(dims, dv, leftLine, rightLine, leftMid, rightMid, draftDoc, createdForCleanup, log)
            adboBehaviorConnected = RunTestB_AssociativityByBehavior(dims, dv, leftLine, rightLine, leftMid, rightMid, draftDoc, createdForCleanup, log)
            auxDeleteKeeps = RunTestC(dims, sheet, draftDoc, leftMid, rightMid, createdForCleanup, log)
            auxHiddenViable = RunTestD(dims, sheet, draftDoc, leftMid, rightMid, createdForCleanup, log)

            If DebugSave Then
                Try : CallByName(draftDoc, "Save", CallType.Method) : Catch : End Try
            End If

            Dim recommended As String = "none_associative_manual_only"
            If existingRelink = TriStateLab.Yes Then
                recommended = "relink_existing_dimension"
            ElseIf adboConnected = TriStateLab.Yes Then
                recommended = "create_new_dimension_with_AddDistanceBetweenObjects"
            ElseIf auxHiddenViable = TriStateLab.Yes Then
                recommended = "hidden_aux_geometry_support"
            End If

            log("[DIMLAB][SUMMARY]")
            log("existing_dimension_relink_possible=" & ToSummary(existingRelink))
            log("adbo_to_dv_connected=" & ToSummary(adboConnected))
            log("adbo_behavior_connected=" & ToSummary(adboBehaviorConnected))
            log("aux_line_delete_keeps_dimension=" & ToSummary(auxDeleteKeeps))
            log("aux_line_hidden_viable=" & ToSummary(auxHiddenViable))
            log("recommended_strategy=" & recommended)

        Finally
            If DebugCleanup Then
                For i As Integer = createdForCleanup.Count - 1 To 0 Step -1
                    Try : CallByName(createdForCleanup(i), "Delete", CallType.Method) : Catch : End Try
                Next
                Try : CallByName(draftDoc, "UpdateAll", CallType.Method, True) : Catch : End Try
            End If
            log("[DIMLAB][END]")
        End Try
    End Sub

    Private Function FindTargetDimension(ByVal dims As Dimensions, ByVal log As Action(Of String)) As Object
        For i As Integer = 1 To dims.Count
            Dim d As Object = Nothing
            Try : d = dims.Item(i) : Catch : d = Nothing : End Try
            If d Is Nothing Then Continue For
            Dim v As Double = ReadDimensionValueSafe(d)
            log("[DIMLAB][DIM][PROP] idx=" & i.ToString(CultureInfo.InvariantCulture) & " name=Value value=" & v.ToString("G17", CultureInfo.InvariantCulture))
            If Not Double.IsNaN(v) AndAlso Math.Abs(v - TargetValueM) <= TargetToleranceM Then
                log("[DIMLAB][TARGET_DIM] idx=" & i.ToString(CultureInfo.InvariantCulture) & " value=" & v.ToString("G17", CultureInfo.InvariantCulture) &
                    " type=" & SafeTypeName(d) & " style=" & SafeStyleName(d))
                Return d
            End If
        Next
        Return Nothing
    End Function

    Private Function FindNearestDimensionByValue(ByVal dims As Dimensions, ByVal target As Double, ByVal log As Action(Of String)) As Object
        Dim best As Object = Nothing
        Dim bestDelta As Double = Double.PositiveInfinity
        For i As Integer = 1 To dims.Count
            Dim d As Object = Nothing
            Try : d = dims.Item(i) : Catch : d = Nothing : End Try
            If d Is Nothing Then Continue For
            Dim v As Double = ReadDimensionValueSafe(d)
            If Double.IsNaN(v) Then Continue For
            Dim delta As Double = Math.Abs(v - target)
            If delta < bestDelta Then
                bestDelta = delta
                best = d
            End If
        Next
        If best IsNot Nothing Then
            log("[DIMLAB][TARGET_DIM][FALLBACK] value=" & ReadDimensionValueSafe(best).ToString("G17", CultureInfo.InvariantCulture) &
                " delta=" & bestDelta.ToString("G17", CultureInfo.InvariantCulture) &
                " type=" & SafeTypeName(best))
        End If
        Return best
    End Function

    Private Sub ForensicAnalyzeAllDimensions(ByVal dims As Dimensions, ByVal sheet As Sheet, ByVal log As Action(Of String))
        If dims Is Nothing Then Return
        log("[DIMLAB][FORENSIC][START] dims=" & dims.Count.ToString(CultureInfo.InvariantCulture))
        For i As Integer = 1 To dims.Count
            Dim d As Object = Nothing
            Try : d = dims.Item(i) : Catch : d = Nothing : End Try
            If d Is Nothing Then Continue For

            log("[DIMLAB][FORENSIC][DIM] idx=" & i.ToString(CultureInfo.InvariantCulture) &
                " type=" & FormatOne(SafeGet(d, "Type")) &
                " value=" & FormatOne(SafeGet(d, "Value")) &
                " style=" & SafeStyleName(d) &
                " status=" & FormatOne(SafeGet(d, "StatusOfDimension")) &
                " trackDistance=" & FormatOne(SafeGet(d, "TrackDistance")) &
                " trackAngle=" & FormatOne(SafeGet(d, "TrackAngle")) &
                " parentType=" & SafeTypeName(SafeGet(d, "Parent")) &
                " layer=" & FormatOne(SafeGet(SafeGet(d, "Layer"), "Name")))

            Dim relCount As Integer = ReadRelatedCount(d)
            Dim gmHits As Integer = 0
            Dim dvHits As Integer = 0
            Dim refHits As Integer = 0
            Dim dvName As String = ""
            log("[DIMLAB][FORENSIC][RELATED_COUNT] idx=" & i.ToString(CultureInfo.InvariantCulture) & " count=" & relCount.ToString(CultureInfo.InvariantCulture))

            For r As Integer = 0 To relCount + 2
                Try
                    Dim o As Object = CallByName(d, "GetRelated", CallType.Method, r)
                    If o Is Nothing Then Continue For
                    Dim tn As String = SafeTypeName(o)
                    If tn.IndexOf("GraphicMember", StringComparison.OrdinalIgnoreCase) >= 0 Then gmHits += 1
                    If tn.IndexOf("Reference", StringComparison.OrdinalIgnoreCase) >= 0 Then refHits += 1
                    If tn.IndexOf("DrawingView", StringComparison.OrdinalIgnoreCase) >= 0 Then
                        dvHits += 1
                        If String.IsNullOrWhiteSpace(dvName) Then dvName = SafeText(SafeGet(o, "Name"))
                    End If
                    If String.IsNullOrWhiteSpace(dvName) Then
                        Dim odv As Object = SafeGet(o, "DrawingView")
                        If odv IsNot Nothing Then
                            dvName = SafeText(SafeGet(odv, "Name"))
                            dvHits += 1
                        End If
                    End If
                    log("[DIMLAB][FORENSIC][RELATED] idx=" & i.ToString(CultureInfo.InvariantCulture) &
                        " relIdx=" & r.ToString(CultureInfo.InvariantCulture) &
                        " type=" & tn)
                Catch
                End Try
            Next

            Dim relSig As String = BuildRelatedSignature(d)
            Dim floating As Boolean = (relCount = 0)
            log("[DIMLAB][FORENSIC][LINK] idx=" & i.ToString(CultureInfo.InvariantCulture) &
                " drawingView=" & If(String.IsNullOrWhiteSpace(dvName), "?", dvName) &
                " dvHits=" & dvHits.ToString(CultureInfo.InvariantCulture) &
                " graphicMemberHits=" & gmHits.ToString(CultureInfo.InvariantCulture) &
                " referenceHits=" & refHits.ToString(CultureInfo.InvariantCulture) &
                " floating=" & floating.ToString(CultureInfo.InvariantCulture) &
                " relSig=" & relSig)

            Try
                Dim dd As Object = CallByName(d, "GetDisplayData", CallType.Method)
                log("[DIMLAB][FORENSIC][DISPLAY] idx=" & i.ToString(CultureInfo.InvariantCulture) &
                    " lineCount=" & ReadCountMethod(dd, "GetLineCount").ToString(CultureInfo.InvariantCulture) &
                    " arcCount=" & ReadCountMethod(dd, "GetArcCount").ToString(CultureInfo.InvariantCulture))
            Catch ex As Exception
                log("[DIMLAB][FORENSIC][DISPLAY] idx=" & i.ToString(CultureInfo.InvariantCulture) & " error=" & ex.Message)
            End Try
        Next
        log("[DIMLAB][FORENSIC][END]")
    End Sub

    Private Sub IntrospectDimension(ByVal d As Object, ByVal log As Action(Of String))
        Dim methods As String() = {"GetValueEx", "UpdateStatus", "Range", "RangeBox", "GetRelatedCount", "GetRelated", "GetDisplayData", "GetTextOffsets", "SetTextOffsets", "GetKeyPoint", "SetKeyPoint", "SetConnect"}
        For Each m In methods
            log("[DIMLAB][DIM][METHOD_EXISTS] name=" & m & " exists=" & ProbeMethodAvailability(d, m).ToString(CultureInfo.InvariantCulture))
        Next

        For Each p In New String() {"Type", "Value", "StatusOfDimension", "Parent", "Style", "TrackDistance", "AbsoluteTrackDistance"}
            Try
                log("[DIMLAB][DIM][PROP] name=" & p & " value=" & FormatOne(CallByName(d, p, CallType.Get)))
            Catch ex As Exception
                log("[DIMLAB][DIM][PROP] name=" & p & " error=" & ex.Message)
            End Try
        Next

        log("[DIMLAB][DIM][STATUS] " & ReadStatusText(d))
        Dim rc As Integer = ReadRelatedCount(d)
        log("[DIMLAB][DIM][RELATED_COUNT] " & rc.ToString(CultureInfo.InvariantCulture))
        For i As Integer = 0 To rc + 1
            Try
                Dim r As Object = CallByName(d, "GetRelated", CallType.Method, i)
                If r IsNot Nothing Then log("[DIMLAB][DIM][RELATED] idx=" & i.ToString(CultureInfo.InvariantCulture) & " type=" & SafeTypeName(r))
            Catch
            End Try
        Next

        Try
            Dim dd As Object = CallByName(d, "GetDisplayData", CallType.Method)
            log("[DIMLAB][DIM][DISPLAYDATA] type=" & SafeTypeName(dd))
            log("[DIMLAB][DIM][DISPLAYDATA] lineCount=" & ReadCountMethod(dd, "GetLineCount").ToString(CultureInfo.InvariantCulture))
            log("[DIMLAB][DIM][DISPLAYDATA] arcCount=" & ReadCountMethod(dd, "GetArcCount").ToString(CultureInfo.InvariantCulture))
        Catch ex As Exception
            log("[DIMLAB][DIM][DISPLAYDATA] error=" & ex.Message)
        End Try
    End Sub

    Private Function FindPrimaryOrthogonalView(ByVal sheet As Sheet, ByVal log As Action(Of String)) As DrawingView
        Dim best As DrawingView = Nothing
        Dim bestScore As Double = Double.MinValue
        For Each v As DrawingView In sheet.DrawingViews
            If IsIsometric(v) Then Continue For
            Dim score As Double = ComputeViewScore(v)
            If score > bestScore Then
                best = v
                bestScore = score
            End If
        Next
        If best IsNot Nothing Then log("[DIMLAB][DV][VIEW] name=" & best.Name & " score=" & bestScore.ToString("G17", CultureInfo.InvariantCulture))
        Return best
    End Function

    Private Function SelectExtremes(
        ByVal dv As DrawingView,
        ByRef leftLine As Object,
        ByRef rightLine As Object,
        ByRef topLine As Object,
        ByRef bottomLine As Object,
        ByRef leftMid As Pt2,
        ByRef rightMid As Pt2,
        ByRef leftStart As Pt2,
        ByRef leftEnd As Pt2,
        ByRef rightStart As Pt2,
        ByRef rightEnd As Pt2,
        ByVal log As Action(Of String)) As Boolean

        Dim minX As Double = Double.PositiveInfinity, maxX As Double = Double.NegativeInfinity
        Dim minY As Double = Double.PositiveInfinity, maxY As Double = Double.NegativeInfinity

        For Each ln As Object In dv.DVLines2d
            Dim p1 = GetLineStart(ln)
            Dim p2 = GetLineEnd(ln)
            Dim dx As Double = Math.Abs(p2.X - p1.X)
            Dim dy As Double = Math.Abs(p2.Y - p1.Y)
            Dim o As String = "oblicua"
            If dx <= AxisTol AndAlso dy > AxisTol Then o = "vertical"
            If dy <= AxisTol AndAlso dx > AxisTol Then o = "horizontal"
            Dim len As Double = Math.Sqrt((p2.X - p1.X) * (p2.X - p1.X) + (p2.Y - p1.Y) * (p2.Y - p1.Y))
            log("[DIMLAB][DV][LINE] type=" & o & " length=" & len.ToString("G17", CultureInfo.InvariantCulture) &
                " x1=" & p1.X.ToString("G17", CultureInfo.InvariantCulture) & " y1=" & p1.Y.ToString("G17", CultureInfo.InvariantCulture) &
                " x2=" & p2.X.ToString("G17", CultureInfo.InvariantCulture) & " y2=" & p2.Y.ToString("G17", CultureInfo.InvariantCulture))

            If o = "vertical" Then
                Dim xm As Double = (p1.X + p2.X) / 2.0R
                If xm < minX Then minX = xm : leftLine = ln : leftStart = p1 : leftEnd = p2
                If xm > maxX Then maxX = xm : rightLine = ln : rightStart = p1 : rightEnd = p2
            ElseIf o = "horizontal" Then
                Dim ym As Double = (p1.Y + p2.Y) / 2.0R
                If ym < minY Then minY = ym : bottomLine = ln
                If ym > maxY Then maxY = ym : topLine = ln
            End If
        Next

        If leftLine Is Nothing OrElse rightLine Is Nothing Then Return False
        leftMid = New Pt2 With {.X = (leftStart.X + leftEnd.X) / 2.0R, .Y = (leftStart.Y + leftEnd.Y) / 2.0R}
        rightMid = New Pt2 With {.X = (rightStart.X + rightEnd.X) / 2.0R, .Y = (rightStart.Y + rightEnd.Y) / 2.0R}
        log("[DIMLAB][DV][EXTREME_LEFT] x=" & leftMid.X.ToString("G17", CultureInfo.InvariantCulture))
        log("[DIMLAB][DV][EXTREME_RIGHT] x=" & rightMid.X.ToString("G17", CultureInfo.InvariantCulture))
        If topLine IsNot Nothing Then log("[DIMLAB][DV][EXTREME_TOP] y=" & maxY.ToString("G17", CultureInfo.InvariantCulture))
        If bottomLine IsNot Nothing Then log("[DIMLAB][DV][EXTREME_BOTTOM] y=" & minY.ToString("G17", CultureInfo.InvariantCulture))
        Return True
    End Function

    Private Function RunTestA(ByVal dimObj As Object, ByVal leftLine As Object, ByVal rightLine As Object, ByVal leftStart As Pt2, ByVal leftEnd As Pt2, ByVal rightStart As Pt2, ByVal rightEnd As Pt2, ByVal leftMid As Pt2, ByVal rightMid As Pt2, ByVal draftDoc As DraftDocument, ByVal log As Action(Of String)) As Boolean
        Dim beforeRel As String = BuildRelatedSignature(dimObj)
        For kp As Integer = 0 To 2
            For Each p In New Pt2() {leftStart, leftEnd, rightStart, rightEnd, leftMid, rightMid}
                log("[DIMLAB][TEST_A][TRY] api=SetKeyPoint kp=" & kp.ToString(CultureInfo.InvariantCulture))
                If TryInvokeWithVariants(dimObj, "SetKeyPoint", New Object() {kp, p.X, p.Y, 0.0R}, New Object() {kp, p.X, p.Y}) Then
                    log("[DIMLAB][TEST_A][OK] api=SetKeyPoint kp=" & kp.ToString(CultureInfo.InvariantCulture))
                Else
                    log("[DIMLAB][TEST_A][FAIL] api=SetKeyPoint kp=" & kp.ToString(CultureInfo.InvariantCulture))
                End If
                Try : CallByName(draftDoc, "UpdateAll", CallType.Method, True) : Catch : End Try
            Next
        Next
        Dim afterRel As String = BuildRelatedSignature(dimObj)
        log("[DIMLAB][TEST_A][STATUS_AFTER] " & ReadStatusText(dimObj))
        log("[DIMLAB][TEST_A][RELATED_AFTER] " & afterRel)
        log("[DIMLAB][TEST_A][VALUE_AFTER] " & ReadDimensionValueSafe(dimObj).ToString("G17", CultureInfo.InvariantCulture))
        Return (beforeRel <> afterRel AndAlso afterRel <> "")
    End Function

    Private Function RunTestB(ByVal dims As Dimensions, ByVal dv As DrawingView, ByVal leftLine As Object, ByVal rightLine As Object, ByVal leftMid As Pt2, ByVal rightMid As Pt2, ByVal draftDoc As DraftDocument, ByVal created As List(Of Object), ByVal log As Action(Of String)) As TriStateLab
        log("[DIMLAB][TEST_B][ADBO][TRY]")
        Dim d As Object = Nothing
        Try
            d = dims.AddDistanceBetweenObjects(leftLine, leftMid.X, leftMid.Y, 0.0R, False, rightLine, rightMid.X, rightMid.Y, 0.0R, False)
        Catch ex As Exception
            log("[DIMLAB][TEST_B][ADBO][FAIL] " & ex.Message)
        End Try
        If d Is Nothing Then Return TriStateLab.No
        created.Add(d)
        Try : CallByName(draftDoc, "UpdateAll", CallType.Method, True) : Catch : End Try
        log("[DIMLAB][TEST_B][ADBO][OK] type=" & SafeTypeName(d))
        log("[DIMLAB][TEST_B][ADBO][VALUE] " & ReadDimensionValueSafe(d).ToString("G17", CultureInfo.InvariantCulture))
        log("[DIMLAB][TEST_B][ADBO][STATUS] " & ReadStatusText(d))
        log("[DIMLAB][TEST_B][ADBO][RELATED] count=" & ReadRelatedCount(d).ToString(CultureInfo.InvariantCulture) & " sig=" & BuildRelatedSignature(d))
        Dim rc As Integer = ReadRelatedCount(d)
        If rc >= 2 AndAlso IsDimensionConnectedToView(dv, d) Then Return TriStateLab.Yes
        If rc > 0 Then Return TriStateLab.Unknown
        Return TriStateLab.No
    End Function

    Private Function RunTestB_AssociativityByBehavior(
        ByVal dims As Dimensions,
        ByVal dv As DrawingView,
        ByVal leftLine As Object,
        ByVal rightLine As Object,
        ByVal leftMid As Pt2,
        ByVal rightMid As Pt2,
        ByVal draftDoc As DraftDocument,
        ByVal created As List(Of Object),
        ByVal log As Action(Of String)) As TriStateLab

        Dim best As TriStateLab = TriStateLab.No

        ' Intento 1: ADBO sobre DVLine2d.
        Dim d As Object = Nothing
        log("[DIMLAB][ASSOC_BEHAVIOR][TRY] method=ADBO source=DVLINE flags=False")
        Try
            d = dims.AddDistanceBetweenObjects(leftLine, leftMid.X, leftMid.Y, 0.0R, False, rightLine, rightMid.X, rightMid.Y, 0.0R, False)
            If d IsNot Nothing Then
                created.Add(d)
                Dim r = EvaluateAssociativityBehavior(d, dv, draftDoc, log, "ADBO_DVLINE_FFalse")
                If r = TriStateLab.Yes Then Return TriStateLab.Yes
                If r = TriStateLab.Unknown Then best = TriStateLab.Unknown
            End If
        Catch ex As Exception
            log("[DIMLAB][ASSOC_BEHAVIOR][FAIL] method=ADBO source=DVLINE flags=False error=" & ex.Message)
        End Try

        ' Intento 2: ADBO sobre DVLine2d con flags=True y puntos de extremo.
        Dim dFlags As Object = Nothing
        log("[DIMLAB][ASSOC_BEHAVIOR][TRY] method=ADBO source=DVLINE flags=True")
        Try
            dFlags = dims.AddDistanceBetweenObjects(leftLine, leftMid.X, leftMid.Y, 0.0R, True, rightLine, rightMid.X, rightMid.Y, 0.0R, True)
            If dFlags IsNot Nothing Then
                created.Add(dFlags)
                Dim r = EvaluateAssociativityBehavior(dFlags, dv, draftDoc, log, "ADBO_DVLINE_FTrue")
                If r = TriStateLab.Yes Then Return TriStateLab.Yes
                If r = TriStateLab.Unknown Then best = TriStateLab.Unknown
            End If
        Catch ex As Exception
            log("[DIMLAB][ASSOC_BEHAVIOR][FAIL] method=ADBO source=DVLINE flags=True error=" & ex.Message)
        End Try

        ' Intento 2: ADBO sobre Reference.
        Dim lref As Object = SafeGet(leftLine, "Reference")
        Dim rref As Object = SafeGet(rightLine, "Reference")
        If lref IsNot Nothing AndAlso rref IsNot Nothing Then
            Dim dRef As Object = Nothing
            log("[DIMLAB][ASSOC_BEHAVIOR][TRY] method=ADBO source=REFERENCE flags=False")
            Try
                dRef = dims.AddDistanceBetweenObjects(lref, leftMid.X, leftMid.Y, 0.0R, False, rref, rightMid.X, rightMid.Y, 0.0R, False)
                If dRef IsNot Nothing Then
                    created.Add(dRef)
                    Dim r = EvaluateAssociativityBehavior(dRef, dv, draftDoc, log, "ADBO_REFERENCE_FFalse")
                    If r = TriStateLab.Yes Then Return TriStateLab.Yes
                    If r = TriStateLab.Unknown Then best = TriStateLab.Unknown
                End If
            Catch ex As Exception
                log("[DIMLAB][ASSOC_BEHAVIOR][FAIL] method=ADBO source=REFERENCE flags=False error=" & ex.Message)
            End Try

            Dim dRefFlags As Object = Nothing
            log("[DIMLAB][ASSOC_BEHAVIOR][TRY] method=ADBO source=REFERENCE flags=True")
            Try
                dRefFlags = dims.AddDistanceBetweenObjects(lref, leftMid.X, leftMid.Y, 0.0R, True, rref, rightMid.X, rightMid.Y, 0.0R, True)
                If dRefFlags IsNot Nothing Then
                    created.Add(dRefFlags)
                    Dim r = EvaluateAssociativityBehavior(dRefFlags, dv, draftDoc, log, "ADBO_REFERENCE_FTrue")
                    If r = TriStateLab.Yes Then Return TriStateLab.Yes
                    If r = TriStateLab.Unknown Then best = TriStateLab.Unknown
                End If
            Catch ex As Exception
                log("[DIMLAB][ASSOC_BEHAVIOR][FAIL] method=ADBO source=REFERENCE flags=True error=" & ex.Message)
            End Try

            ' Intento 3: ADBOEX sobre Reference.
            Dim dRefEx As Object = Nothing
            log("[DIMLAB][ASSOC_BEHAVIOR][TRY] method=ADBOEX source=REFERENCE")
            Try
                dRefEx = CallByName(dims, "AddDistanceBetweenObjectsEX", CallType.Method, lref, rref, leftMid.X, leftMid.Y, rightMid.X, rightMid.Y)
                If dRefEx IsNot Nothing Then
                    created.Add(dRefEx)
                    Dim r = EvaluateAssociativityBehavior(dRefEx, dv, draftDoc, log, "ADBOEX_REFERENCE")
                    If r = TriStateLab.Yes Then Return TriStateLab.Yes
                    If r = TriStateLab.Unknown Then best = TriStateLab.Unknown
                End If
            Catch ex As Exception
                log("[DIMLAB][ASSOC_BEHAVIOR][FAIL] method=ADBOEX source=REFERENCE error=" & ex.Message)
            End Try
        Else
            log("[DIMLAB][ASSOC_BEHAVIOR][SKIP] source=REFERENCE reason=missing_reference")
        End If

        Return best
    End Function

    Private Function EvaluateAssociativityBehavior(ByVal d As Object, ByVal dv As DrawingView, ByVal draftDoc As DraftDocument, ByVal log As Action(Of String), ByVal tag As String) As TriStateLab
        Dim beforeVal As Double = ReadDimensionValueSafe(d)
        Dim beforeTrack As Double = ReadTrackDistanceSafe(d)
        Dim beforeRange As String = ReadDimensionRangeSignature(d)
        log("[DIMLAB][ASSOC_BEHAVIOR][BEFORE] tag=" & tag & " value=" & beforeVal.ToString("G17", CultureInfo.InvariantCulture) &
            " trackDistance=" & beforeTrack.ToString("G17", CultureInfo.InvariantCulture) &
            " range=" & beforeRange)

        Dim ox As Double = 0.0R, oy As Double = 0.0R
        Dim originReadOk As Boolean = TryReadViewOrigin(dv, ox, oy)
        Dim moved As Boolean = False
        Dim dx As Double = 0.003R
        Dim dy As Double = 0.0R
        moved = TryMoveViewOrigin(dv, ox + dx, oy + dy)
        If moved Then
            log("[DIMLAB][ASSOC_BEHAVIOR][MOVE_VIEW] tag=" & tag & " applied=True dx=" & dx.ToString("G17", CultureInfo.InvariantCulture) & " dy=" & dy.ToString("G17", CultureInfo.InvariantCulture))
        Else
            log("[DIMLAB][ASSOC_BEHAVIOR][MOVE_VIEW] tag=" & tag & " applied=False error=no_supported_origin_signature")
        End If

        Try : CallByName(draftDoc, "UpdateAll", CallType.Method, True) : Catch : End Try

        Dim afterVal As Double = ReadDimensionValueSafe(d)
        Dim afterTrack As Double = ReadTrackDistanceSafe(d)
        Dim afterRange As String = ReadDimensionRangeSignature(d)
        log("[DIMLAB][ASSOC_BEHAVIOR][AFTER] tag=" & tag & " value=" & afterVal.ToString("G17", CultureInfo.InvariantCulture) &
            " trackDistance=" & afterTrack.ToString("G17", CultureInfo.InvariantCulture) &
            " range=" & afterRange)

        If moved AndAlso originReadOk Then
            If TryMoveViewOrigin(dv, ox, oy) Then
                Try : CallByName(draftDoc, "UpdateAll", CallType.Method, True) : Catch : End Try
                log("[DIMLAB][ASSOC_BEHAVIOR][MOVE_VIEW_RESTORE] tag=" & tag & " ok=True")
            Else
                log("[DIMLAB][ASSOC_BEHAVIOR][MOVE_VIEW_RESTORE] tag=" & tag & " ok=False error=no_supported_origin_signature")
            End If
        End If

        Dim changedGeom As Boolean =
            (Not String.Equals(beforeRange, afterRange, StringComparison.Ordinal)) OrElse
            Math.Abs(beforeTrack - afterTrack) > 0.0000001R
        Dim valueStable As Boolean = (Not Double.IsNaN(beforeVal) AndAlso Not Double.IsNaN(afterVal) AndAlso Math.Abs(beforeVal - afterVal) < 0.00001R)
        Dim connectedFunctional As Boolean = moved AndAlso changedGeom AndAlso valueStable

        log("[DIMLAB][ASSOC_BEHAVIOR][RESULT] tag=" & tag & " connected_functional=" & connectedFunctional.ToString(CultureInfo.InvariantCulture) &
            " changedGeom=" & changedGeom.ToString(CultureInfo.InvariantCulture) &
            " valueStable=" & valueStable.ToString(CultureInfo.InvariantCulture))

        If connectedFunctional Then Return TriStateLab.Yes
        If moved Then Return TriStateLab.Unknown
        Return TriStateLab.No
    End Function

    Private Function RunTestC(ByVal dims As Dimensions, ByVal sheet As Sheet, ByVal draftDoc As DraftDocument, ByVal p1 As Pt2, ByVal p2 As Pt2, ByVal created As List(Of Object), ByVal log As Action(Of String)) As TriStateLab
        Dim aux As Object = CreateAuxLine(sheet, p1, p2)
        If aux Is Nothing Then Return TriStateLab.No
        created.Add(aux)
        log("[DIMLAB][TEST_C][AUX_LINE_CREATE] ok=True")
        Dim d As Object = TryAddLengthOnObject(dims, aux)
        If d Is Nothing Then Return TriStateLab.No
        created.Add(d)
        log("[DIMLAB][TEST_C][ADDLENGTH_OK] value=" & ReadDimensionValueSafe(d).ToString("G17", CultureInfo.InvariantCulture))
        log("[DIMLAB][TEST_C][RELATED_BEFORE_DELETE] count=" & ReadRelatedCount(d).ToString(CultureInfo.InvariantCulture) & " sig=" & BuildRelatedSignature(d))
        log("[DIMLAB][TEST_C][DELETE_AUX_LINE]")
        Try : CallByName(aux, "Delete", CallType.Method) : Catch : Return TriStateLab.No : End Try
        Try : CallByName(draftDoc, "UpdateAll", CallType.Method, True) : Catch : End Try
        log("[DIMLAB][TEST_C][STATUS_AFTER_DELETE] " & ReadStatusText(d))
        log("[DIMLAB][TEST_C][RELATED_AFTER_DELETE] count=" & ReadRelatedCount(d).ToString(CultureInfo.InvariantCulture) & " sig=" & BuildRelatedSignature(d))
        If IsDimensionComAlive(d) AndAlso Not Double.IsNaN(ReadDimensionValueSafe(d)) Then Return TriStateLab.Yes
        Return TriStateLab.No
    End Function

    Private Function RunTestD(ByVal dims As Dimensions, ByVal sheet As Sheet, ByVal draftDoc As DraftDocument, ByVal p1 As Pt2, ByVal p2 As Pt2, ByVal created As List(Of Object), ByVal log As Action(Of String)) As TriStateLab
        Dim aux As Object = CreateAuxLine(sheet, p1, p2)
        If aux Is Nothing Then Return TriStateLab.No
        created.Add(aux)
        log("[DIMLAB][TEST_D][AUX_LAYER_CREATE] try=True")
        Dim d As Object = TryAddLengthOnObject(dims, aux)
        If d Is Nothing Then Return TriStateLab.No
        created.Add(d)
        Dim hidden As Boolean = TrySetProperty(aux, "Visible", False) OrElse TryMoveToHiddenLayer(sheet, aux)
        log("[DIMLAB][TEST_D][AUX_LINE_HIDE] hiddenApplied=" & hidden.ToString(CultureInfo.InvariantCulture))
        Try : CallByName(draftDoc, "UpdateAll", CallType.Method, True) : Catch : End Try
        log("[DIMLAB][TEST_D][STATUS] " & ReadStatusText(d))
        log("[DIMLAB][TEST_D][RELATED] count=" & ReadRelatedCount(d).ToString(CultureInfo.InvariantCulture) & " sig=" & BuildRelatedSignature(d))
        If hidden AndAlso ReadRelatedCount(d) > 0 Then Return TriStateLab.Yes
        If hidden Then Return TriStateLab.Unknown
        Return TriStateLab.No
    End Function

    Private Function ResolveSheetHoja1OrActive(ByVal draftDoc As DraftDocument) As Sheet
        Try
            For i As Integer = 1 To draftDoc.Sheets.Count
                Dim sh As Sheet = CType(draftDoc.Sheets.Item(i), Sheet)
                If String.Equals(sh.Name, "Hoja1", StringComparison.OrdinalIgnoreCase) Then Return sh
            Next
        Catch
        End Try
        Try : Return draftDoc.ActiveSheet : Catch : Return Nothing : End Try
    End Function

    Private Function ComputeViewScore(ByVal v As DrawingView) As Double
        Dim minX As Double = Double.PositiveInfinity, maxX As Double = Double.NegativeInfinity
        Dim minY As Double = Double.PositiveInfinity, maxY As Double = Double.NegativeInfinity
        Try
            For Each ln As Object In v.DVLines2d
                Dim p1 = GetLineStart(ln)
                Dim p2 = GetLineEnd(ln)
                minX = Math.Min(minX, Math.Min(p1.X, p2.X))
                maxX = Math.Max(maxX, Math.Max(p1.X, p2.X))
                minY = Math.Min(minY, Math.Min(p1.Y, p2.Y))
                maxY = Math.Max(maxY, Math.Max(p1.Y, p2.Y))
            Next
        Catch
            Return Double.MinValue
        End Try
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

    Private Function CreateAuxLine(ByVal sheet As Sheet, ByVal p1 As Pt2, ByVal p2 As Pt2) As Object
        Try : Return CallByName(sheet.Lines2d, "AddBy2Points", CallType.Method, p1.X, p1.Y, p2.X, p2.Y) : Catch : End Try
        Try : Return CallByName(sheet.Lines2d, "AddLine", CallType.Method, p1.X, p1.Y, p2.X, p2.Y) : Catch : End Try
        Return Nothing
    End Function

    Private Function TryAddLengthOnObject(ByVal dims As Dimensions, ByVal srcObj As Object) As Object
        Dim r As Object = srcObj
        Try : r = CallByName(srcObj, "Reference", CallType.Get) : Catch : End Try
        Try : Return CallByName(dims, "AddLength", CallType.Method, r) : Catch : Return Nothing : End Try
    End Function

    Private Function TrySetProperty(ByVal obj As Object, ByVal propName As String, ByVal value As Object) As Boolean
        Try : CallByName(obj, propName, CallType.Let, value) : Return True : Catch : Return False : End Try
    End Function

    Private Function TryMoveToHiddenLayer(ByVal sheet As Sheet, ByVal obj As Object) As Boolean
        Try
            Dim layers As Object = CallByName(sheet, "Layers", CallType.Get)
            For i As Integer = 1 To Convert.ToInt32(CallByName(layers, "Count", CallType.Get), CultureInfo.InvariantCulture)
                Dim lyr As Object = CallByName(layers, "Item", CallType.Method, i)
                Dim n As String = Convert.ToString(CallByName(lyr, "Name", CallType.Get), CultureInfo.InvariantCulture).ToLowerInvariant()
                If n.Contains("aux_dim_ref") OrElse n.Contains("no_plot") OrElse n.Contains("dim_support") OrElse n.Contains("hidden") Then
                    CallByName(obj, "Layer", CallType.Let, lyr)
                    Return True
                End If
            Next
        Catch
        End Try
        Return False
    End Function

    Private Function ReadRelatedCount(ByVal dimObj As Object) As Integer
        Try : Return Convert.ToInt32(CallByName(dimObj, "GetRelatedCount", CallType.Method), CultureInfo.InvariantCulture) : Catch : Return 0 : End Try
    End Function

    Private Function BuildRelatedSignature(ByVal dimObj As Object) As String
        Dim rc As Integer = ReadRelatedCount(dimObj)
        If rc <= 0 Then Return ""
        Dim items As New List(Of String)
        For i As Integer = 0 To rc + 1
            Try
                Dim r As Object = CallByName(dimObj, "GetRelated", CallType.Method, i)
                If r IsNot Nothing Then items.Add(i.ToString(CultureInfo.InvariantCulture) & ":" & SafeTypeName(r))
            Catch
            End Try
        Next
        Return String.Join("|", items)
    End Function

    Private Function ReadStatusText(ByVal dimObj As Object) As String
        Try : Return "status=" & FormatOne(CallByName(dimObj, "StatusOfDimension", CallType.Get)) : Catch : End Try
        Try : Return "status=" & FormatOne(CallByName(dimObj, "UpdateStatus", CallType.Method)) : Catch ex As Exception : Return "status_error=" & ex.Message : End Try
    End Function

    Private Function ReadCountMethod(ByVal obj As Object, ByVal methodName As String) As Integer
        Try
            Dim v As Object = CallByName(obj, methodName, CallType.Method)
            Return Convert.ToInt32(v, CultureInfo.InvariantCulture)
        Catch
            Return 0
        End Try
    End Function

    Private Function IsDimensionConnectedToView(ByVal dv As DrawingView, ByVal dimObj As Object) As Boolean
        Try
            Dim ss As Object = CallByName(CallByName(dv, "Application", CallType.Get), "ActiveSelectSet", CallType.Get)
            CallByName(ss, "RemoveAll", CallType.Method)
            CallByName(dv, "AddConnectedDimensionsToSelectSet", CallType.Method)
            Dim n As Integer = Convert.ToInt32(CallByName(ss, "Count", CallType.Get), CultureInfo.InvariantCulture)
            For i As Integer = 1 To n
                If Object.ReferenceEquals(CallByName(ss, "Item", CallType.Method, i), dimObj) Then Return True
            Next
        Catch
        End Try
        Return False
    End Function

    Private Function ReadDimensionValueSafe(ByVal dimObj As Object) As Double
        Try : Return Convert.ToDouble(CallByName(dimObj, "Value", CallType.Get), CultureInfo.InvariantCulture) : Catch : End Try
        Try : Return Convert.ToDouble(CallByName(dimObj, "GetValueEx", CallType.Method), CultureInfo.InvariantCulture) : Catch : End Try
        Try : Return Convert.ToDouble(CallByName(dimObj, "GetValueEx", CallType.Method, 0), CultureInfo.InvariantCulture) : Catch : End Try
        Return Double.NaN
    End Function

    Private Function ReadTrackDistanceSafe(ByVal dimObj As Object) As Double
        Try : Return Convert.ToDouble(CallByName(dimObj, "TrackDistance", CallType.Get), CultureInfo.InvariantCulture) : Catch : End Try
        Try : Return Convert.ToDouble(CallByName(dimObj, "AbsoluteTrackDistance", CallType.Get), CultureInfo.InvariantCulture) : Catch : End Try
        Return Double.NaN
    End Function

    Private Function ReadDimensionRangeSignature(ByVal dimObj As Object) As String
        Try
            Dim x1 As Double = 0, y1 As Double = 0, x2 As Double = 0, y2 As Double = 0
            CallByName(dimObj, "Range", CallType.Method, x1, y1, x2, y2)
            Return x1.ToString("G17", CultureInfo.InvariantCulture) & "," &
                   y1.ToString("G17", CultureInfo.InvariantCulture) & "," &
                   x2.ToString("G17", CultureInfo.InvariantCulture) & "," &
                   y2.ToString("G17", CultureInfo.InvariantCulture)
        Catch
            Return "NA"
        End Try
    End Function

    Private Function IsDimensionComAlive(ByVal dimObj As Object) As Boolean
        Try : Dim _x = CallByName(dimObj, "StatusOfDimension", CallType.Get) : Return True : Catch : End Try
        Try : Dim _y = CallByName(dimObj, "UpdateStatus", CallType.Method) : Return True : Catch : End Try
        Return False
    End Function

    Private Function ProbeMethodAvailability(ByVal target As Object, ByVal methodName As String) As Boolean
        If target Is Nothing Then Return False
        If HasMethod(target, methodName) Then Return True
        Select Case methodName
            Case "SetKeyPoint"
                Return TryInvokeWithVariants(target, methodName, New Object() {0, 0.0R, 0.0R, 0.0R}, New Object() {0, 0.0R, 0.0R})
            Case "SetTextOffsets"
                Return TryInvokeWithVariants(target, methodName, New Object() {0.0R, 0.0R}, New Object() {0.0R, 0.0R, 0.0R})
            Case Else
                Return False
        End Select
    End Function

    Private Function TryInvokeWithVariants(ByVal target As Object, ByVal methodName As String, ParamArray variants As Object()) As Boolean
        For Each rawArgs As Object In variants
            Dim args As Object() = TryCast(rawArgs, Object())
            If args Is Nothing Then args = New Object() {rawArgs}
            Try
                CallByName(target, methodName, CallType.Method, args)
                Return True
            Catch
            End Try
        Next
        Return False
    End Function

    Private Function HasMethod(ByVal obj As Object, ByVal methodName As String) As Boolean
        Try : Return obj.GetType().GetMethods().Any(Function(m) String.Equals(m.Name, methodName, StringComparison.OrdinalIgnoreCase)) : Catch : Return False : End Try
    End Function

    Private Function GetLineStart(ByVal lineObj As Object) As Pt2
        Dim x As Double = 0.0R, y As Double = 0.0R
        Try : lineObj.GetStartPoint(x, y) : Catch : End Try
        Return New Pt2 With {.X = x, .Y = y}
    End Function

    Private Function GetLineEnd(ByVal lineObj As Object) As Pt2
        Dim x As Double = 0.0R, y As Double = 0.0R
        Try : lineObj.GetEndPoint(x, y) : Catch : End Try
        Return New Pt2 With {.X = x, .Y = y}
    End Function

    Private Function SafeTypeName(ByVal o As Object) As String
        Try : Return If(o Is Nothing, "Nothing", o.GetType().Name) : Catch : Return "?" : End Try
    End Function

    Private Function SafeStyleName(ByVal dimObj As Object) As String
        Try
            Dim st As Object = CallByName(dimObj, "Style", CallType.Get)
            If st Is Nothing Then Return "<none>"
            Return Convert.ToString(CallByName(st, "Name", CallType.Get), CultureInfo.InvariantCulture)
        Catch
            Return "<err>"
        End Try
    End Function

    Private Function FormatOne(ByVal value As Object) As String
        If value Is Nothing Then Return "Nothing"
        If TypeOf value Is Double Then Return DirectCast(value, Double).ToString("G17", CultureInfo.InvariantCulture)
        If TypeOf value Is IFormattable Then Return DirectCast(value, IFormattable).ToString(Nothing, CultureInfo.InvariantCulture)
        Return Convert.ToString(value, CultureInfo.InvariantCulture)
    End Function

    Private Function SafeGet(ByVal obj As Object, ByVal member As String) As Object
        If obj Is Nothing Then Return Nothing
        Try
            Return CallByName(obj, member, CallType.Get)
        Catch
            Return Nothing
        End Try
    End Function

    Private Function SafeText(ByVal value As Object) As String
        If value Is Nothing Then Return ""
        Try
            Return Convert.ToString(value, CultureInfo.InvariantCulture)
        Catch
            Return ""
        End Try
    End Function

    Private Function SafeToDouble(ByVal value As Object) As Double
        If value Is Nothing Then Return 0.0R
        Try
            Return Convert.ToDouble(value, CultureInfo.InvariantCulture)
        Catch
            Return 0.0R
        End Try
    End Function

    Private Function TryReadViewOrigin(ByVal dv As DrawingView, ByRef x As Double, ByRef y As Double) As Boolean
        x = 0.0R : y = 0.0R
        If dv Is Nothing Then Return False
        Try
            x = Convert.ToDouble(CallByName(dv, "OriginX", CallType.Get), CultureInfo.InvariantCulture)
            y = Convert.ToDouble(CallByName(dv, "OriginY", CallType.Get), CultureInfo.InvariantCulture)
            Return True
        Catch
        End Try
        Try
            Dim p As Object = CallByName(dv, "Origin", CallType.Get)
            x = SafeToDouble(SafeGet(p, "X"))
            y = SafeToDouble(SafeGet(p, "Y"))
            Return True
        Catch
        End Try
        Try
            Dim args As Object() = {0.0R, 0.0R}
            dv.GetType().InvokeMember("GetOrigin", Reflection.BindingFlags.InvokeMethod Or Reflection.BindingFlags.Public Or Reflection.BindingFlags.Instance, Nothing, dv, args)
            x = SafeToDouble(args(0))
            y = SafeToDouble(args(1))
            Return True
        Catch
        End Try
        Return False
    End Function

    Private Function TryMoveViewOrigin(ByVal dv As DrawingView, ByVal x As Double, ByVal y As Double) As Boolean
        If dv Is Nothing Then Return False
        Try
            CallByName(dv, "OriginX", CallType.Let, x)
            CallByName(dv, "OriginY", CallType.Let, y)
            Return True
        Catch
        End Try
        Try
            Dim p As Object = CallByName(dv, "Origin", CallType.Get)
            If p IsNot Nothing Then
                CallByName(p, "X", CallType.Let, x)
                CallByName(p, "Y", CallType.Let, y)
                Return True
            End If
        Catch
        End Try
        Try
            CallByName(dv, "SetOrigin", CallType.Method, x, y)
            Return True
        Catch
        End Try
        Try
            Dim args As Object() = {x, y}
            dv.GetType().InvokeMember("SetOrigin", Reflection.BindingFlags.InvokeMethod Or Reflection.BindingFlags.Public Or Reflection.BindingFlags.Instance, Nothing, dv, args)
            Return True
        Catch
        End Try
        Return False
    End Function

    Private Function ToSummary(ByVal v As TriStateLab) As String
        If v = TriStateLab.Yes Then Return "True"
        If v = TriStateLab.No Then Return "False"
        Return "Unknown"
    End Function
End Module
