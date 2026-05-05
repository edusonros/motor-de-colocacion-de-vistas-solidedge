Option Strict Off

Imports System.Collections.Generic
Imports System.Globalization
Imports System.IO
Imports System.Linq
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports Microsoft.VisualBasic
Imports SolidEdgeDraft
Imports SolidEdgeFramework
Imports SolidEdgeFrameworkSupport
Imports Dimension = SolidEdgeFrameworkSupport.Dimension

Namespace Services.Dimensioning.Labs

    ''' <summary>Laboratorio aislado DIMLAB: DVLine2d.Reference + AddDistanceBetweenObjects (no sustituye el motor principal).</summary>
    Public NotInheritable Class DrawingViewDimensioningLab
        Private Const DimLabHorizontalGap As Double = 0.024R
        Private Const DimLabVerticalGap As Double = 0.03R
        Private Const DimLabMinGap As Double = 0.024R
        Private Const DimLabMaxGap As Double = 0.03R

        Private Sub New()
        End Sub

        Private Class DVEndpointCandidate
            Public LineIndex As Integer
            Public Line As Object
            Public Ref As Object
            Public X As Double
            Public Y As Double
            Public IsStart As Boolean
            Public Length As Double
        End Class

        Private Class LineBounds
            Public MinX As Double
            Public MaxX As Double
            Public MinY As Double
            Public MaxY As Double
            Public ExpectedWidth As Double
            Public ExpectedHeight As Double
        End Class

        Private Shared Function ToDVBounds(b As LineBounds) As DVBounds
            Return New DVBounds With {
                .MinX = b.MinX, .MaxX = b.MaxX, .MinY = b.MinY, .MaxY = b.MaxY,
                .ExpectedWidth = b.ExpectedWidth, .ExpectedHeight = b.ExpectedHeight
            }
        End Function

        Private Class HorizSeg
            Public Ym As Double
            Public Line As Object
            Public X1 As Double
            Public Y1 As Double
            Public X2 As Double
            Public Y2 As Double
            Public Length As Double
        End Class

        Private Class HorizLinePick
            Public Line As Object
            Public X1 As Double
            Public Y1 As Double
            Public X2 As Double
            Public Y2 As Double
        End Class

        Private NotInheritable Class LabLineReadSession
            Public SelectionLogged As Boolean
        End Class

        Private NotInheritable Class LabVisibilitySummary
            Public HorizTopCreate As String = "N/A"
            Public HorizTopVisible As Boolean = False
            Public HorizTopRange As String = "N/A"
            Public HorizTopTrack As String = "N/A"
            Public HorizTopReason As String = "none"
            Public HorizTopDone As Boolean = False
            Public VertBestCreate As String = "N/A"
            Public VertBestVisible As Boolean = False
            Public VertBestRange As String = "N/A"
            Public VertBestTrack As String = "N/A"
            Public VertBestReason As String = "none"
            Public VertRecorded As Boolean = False
            Public GapCreate As String = "N/A"
            Public GapVisible As Boolean = False
            Public GapRange As String = "N/A"
            Public GapTrack As String = "N/A"
            Public GapReason As String = "none"
            Public GapDone As Boolean = False

            Public Sub Record(testName As String, createRes As String, visible As Boolean, rangeStr As String, trackStr As String, reason As String)
                If String.Equals(testName, "HorizontalTotal_TopPair", StringComparison.OrdinalIgnoreCase) Then
                    HorizTopCreate = createRes
                    HorizTopVisible = visible
                    HorizTopRange = rangeStr
                    HorizTopTrack = trackStr
                    HorizTopReason = reason
                    HorizTopDone = True
                ElseIf (testName.StartsWith("VerticalTotal", StringComparison.OrdinalIgnoreCase) OrElse
                        testName.StartsWith("Vertical_", StringComparison.OrdinalIgnoreCase)) AndAlso Not VertRecorded Then
                    VertBestCreate = createRes
                    VertBestVisible = visible
                    VertBestRange = rangeStr
                    VertBestTrack = trackStr
                    VertBestReason = reason
                    VertRecorded = True
                ElseIf String.Equals(testName, "SmallGap_BetweenParallelHorizontals", StringComparison.OrdinalIgnoreCase) Then
                    GapCreate = createRes
                    GapVisible = visible
                    GapRange = rangeStr
                    GapTrack = trackStr
                    GapReason = reason
                    GapDone = True
                End If
            End Sub
        End Class

        Private NotInheritable Class LabRunContext
            Public App As Application
            Public Draft As DraftDocument
            Public Sheet As Sheet
            Public Vis As New LabVisibilitySummary()
            Public ReadOnly PreserveDiagnostics As New List(Of Object)()
            Public ForensicInteractive As Boolean
            Public EnableAltPlacementLog As Boolean
            Public AbortReason As String
            Public SummaryHorizCreate As String = "SKIPPED"
            Public SummaryHorizValue As String = ""
            Public SummaryHorizVisible As Boolean = False
            Public SummaryHorizConnected As Boolean = False
            Public SummaryVertCreate As String = "SKIPPED"
            Public SummaryVertValue As String = ""
            Public SummaryVertVisible As Boolean = False
            Public SummaryVertConnected As Boolean = False
            Public SummaryVertValueClass As String = ""
            Public Vis0TextObj As Object
            Public Vis0LineObj As Object
            Public Vis0PlainDim As Dimension
            Public DvRefHorizontalDim As Dimension
            Public PlainTextProbe As String = "N/A"
            Public PlainLineProbe As String = "N/A"
            Public PlainAddLengthCreate As String = "N/A"
            Public PlainGraphicsMaterialized As Boolean = False
            Public DvRefHorizGraphicsMaterialized As Boolean = False
            Public SummaryHorizontalViewLabel As String = ""
            Public SummaryVerticalViewLabel As String = ""
            Public SummaryVerticalExpectedHeight As String = ""
            Public SummaryVerticalReason As String = ""
            Public KeepFailedDimensions As Boolean = False
            Public SeenConnectedBySelectSet As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            Public ResolvedStyleObj As Object
            Public ResolvedStyleName As String = "N/A"
            Public TextCenterHorizontalResult As String = "SKIP"
            Public TextCenterVerticalResult As String = "SKIP"
            Public AxisModeDefaultResult As String = "N/A"
            Public AxisModeImpliedResult As String = "N/A"
            Public AxisModeExplicitResult As String = "N/A"
            Public BestVerticalAxisMode As String = "N/A"
            Public IsCleanFullStrict As Boolean = False
            Public RequestedStyleName As String = "U3,5"
            Public SummaryHorizontalStyleFinal As String = "N/A"
            Public SummaryVerticalStyleFinal As String = "N/A"
            Public SummaryStyleResult As String = "N/A"
        End Class

        Private Enum ConnectivityState
            Unknown = 0
            FalseState = 1
            TrueState = 2
        End Enum

        Private NotInheritable Class LabEvalResult
            Public ValueOk As Boolean
            Public VisibleOk As Boolean
            Public Materialized As Boolean
            Public Connected As ConnectivityState
            Public Result As String
        End Class

        Private NotInheritable Class VerticalViewScanResult
            Public View As DrawingView
            Public Idx As Integer
            Public Name As String
            Public Score As Integer
            Public ScoreReason As String
            Public Bounds As LineBounds
            Public DvBounds As DVBounds
            Public TolB As Double
            Public BottomL As List(Of DVEndpointCandidate)
            Public TopL As List(Of DVEndpointCandidate)
            Public LeftL As List(Of DVEndpointCandidate)
            Public RightL As List(Of DVEndpointCandidate)
            Public BestSameXDelta As Double
            Public BestPb As DVEndpointCandidate
            Public BestPt As DVEndpointCandidate
            Public HasHorizTopBottomOverlap As Boolean
            Public LineCount As Integer
        End Class

        Private Shared ReadOnly DvInvokeFlags As BindingFlags =
            BindingFlags.InvokeMethod Or BindingFlags.Public Or BindingFlags.Instance

        ''' <summary>Lectura tipada de extremos DVLine2d: solo variables Double locales ByRef en GetStartPoint/GetEndPoint.</summary>
        Private Shared Function TryReadDVLineEndpoints(dvLine As DVLine2d,
                                                       ByRef x1 As Double,
                                                       ByRef y1 As Double,
                                                       ByRef x2 As Double,
                                                       ByRef y2 As Double,
                                                       lab As DimLabLogger,
                                                       Optional idx As Integer = -1,
                                                       Optional doLog As Boolean = True) As Boolean
            x1 = 0.0R
            y1 = 0.0R
            x2 = 0.0R
            y2 = 0.0R
            Try
                Dim sx As Double = 0.0R
                Dim sy As Double = 0.0R
                Dim ex As Double = 0.0R
                Dim ey As Double = 0.0R
                dvLine.GetStartPoint(sx, sy)
                dvLine.GetEndPoint(ex, ey)
                x1 = sx
                y1 = sy
                x2 = ex
                y2 = ey

                Dim dx = x2 - x1
                Dim dy = y2 - y1
                Dim endpointLength = Math.Sqrt(dx * dx + dy * dy)

                Dim dvLength As Double = 0.0R
                Try
                    dvLength = CDbl(dvLine.Length)
                Catch
                End Try

                If doLog Then
                    lab.Log("DVLINE", "READ_TYPED", "idx=" & idx.ToString(CultureInfo.InvariantCulture) &
                              " start=(" & x1.ToString("G17", CultureInfo.InvariantCulture) & "," & y1.ToString("G17", CultureInfo.InvariantCulture) & ")" &
                              " end=(" & x2.ToString("G17", CultureInfo.InvariantCulture) & "," & y2.ToString("G17", CultureInfo.InvariantCulture) & ")" &
                              " endpointLength=" & endpointLength.ToString("G17", CultureInfo.InvariantCulture) &
                              " dvLength=" & dvLength.ToString("G17", CultureInfo.InvariantCulture))
                End If

                If dvLength > 0.000001R AndAlso endpointLength < 0.000001R Then
                    If doLog Then
                        lab.Log("DVLINE", "READ_TYPED_INVALID", "idx=" & idx.ToString(CultureInfo.InvariantCulture) &
                                  " reason=endpoints_zero_but_length_nonzero")
                    End If
                    Return False
                End If

                Return True
            Catch ex As Exception
                If doLog Then
                    lab.Log("DVLINE", "READ_TYPED_FAIL", "idx=" & idx.ToString(CultureInfo.InvariantCulture) & " " & ex.Message)
                End If
            End Try
            Return False
        End Function

        Private Shared Function TryReadDVLineEndpointsLateBinding(dvLine As Object,
                                                                  ByRef x1 As Double,
                                                                  ByRef y1 As Double,
                                                                  ByRef x2 As Double,
                                                                  ByRef y2 As Double,
                                                                  lab As DimLabLogger,
                                                                  Optional idx As Integer = -1,
                                                                  Optional doLog As Boolean = True) As Boolean
            x1 = 0.0R
            y1 = 0.0R
            x2 = 0.0R
            y2 = 0.0R
            If dvLine Is Nothing Then Return False
            Try
                Dim argsStart As Object() = New Object() {0.0R, 0.0R}
                dvLine.GetType().InvokeMember("GetStartPoint", DvInvokeFlags, Nothing, dvLine, argsStart)
                x1 = CDbl(argsStart(0))
                y1 = CDbl(argsStart(1))

                Dim argsEnd As Object() = New Object() {0.0R, 0.0R}
                dvLine.GetType().InvokeMember("GetEndPoint", DvInvokeFlags, Nothing, dvLine, argsEnd)
                x2 = CDbl(argsEnd(0))
                y2 = CDbl(argsEnd(1))

                Dim dx = x2 - x1
                Dim dy = y2 - y1
                Dim endpointLength = Math.Sqrt(dx * dx + dy * dy)
                Dim dvLength As Double = 0.0R
                Try
                    dvLength = CDbl(CallByName(dvLine, "Length", CallType.Get))
                Catch
                End Try

                If doLog Then
                    lab.Log("DVLINE", "READ_LATE", "idx=" & idx.ToString(CultureInfo.InvariantCulture) &
                              " start=(" & x1.ToString("G17", CultureInfo.InvariantCulture) & "," & y1.ToString("G17", CultureInfo.InvariantCulture) & ")" &
                              " end=(" & x2.ToString("G17", CultureInfo.InvariantCulture) & "," & y2.ToString("G17", CultureInfo.InvariantCulture) & ")" &
                              " endpointLength=" & endpointLength.ToString("G17", CultureInfo.InvariantCulture) &
                              " dvLength=" & dvLength.ToString("G17", CultureInfo.InvariantCulture))
                End If

                If dvLength > 0.000001R AndAlso endpointLength < 0.000001R Then
                    If doLog Then
                        lab.Log("DVLINE", "READ_LATE_INVALID", "idx=" & idx.ToString(CultureInfo.InvariantCulture) &
                                  " reason=endpoints_zero_but_length_nonzero")
                    End If
                    Return False
                End If
                Return True
            Catch ex As Exception
                If doLog Then
                    lab.Log("DVLINE", "READ_LATE_FAIL", "idx=" & idx.ToString(CultureInfo.InvariantCulture) & " " & ex.Message)
                End If
            End Try
            Return False
        End Function

        Private Shared Function TryReadDVLineEndpointsBest(lineObj As Object,
                                                           ByRef x1 As Double,
                                                           ByRef y1 As Double,
                                                           ByRef x2 As Double,
                                                           ByRef y2 As Double,
                                                           lab As DimLabLogger,
                                                           idx As Integer,
                                                           session As LabLineReadSession,
                                                           Optional doLog As Boolean = True) As Boolean
            If lineObj Is Nothing Then Return False
            Dim typedLine As DVLine2d = Nothing
            Try
                typedLine = CType(lineObj, DVLine2d)
            Catch
            End Try

            Dim okTyped As Boolean = False
            If typedLine IsNot Nothing Then
                okTyped = TryReadDVLineEndpoints(typedLine, x1, y1, x2, y2, lab, idx, doLog)
            End If

            If okTyped Then
                If session IsNot Nothing AndAlso Not session.SelectionLogged Then
                    lab.Log("DVLINE", "READ_METHOD_SELECTED", "typed")
                    session.SelectionLogged = True
                End If
                Return True
            End If

            If TryReadDVLineEndpointsLateBinding(lineObj, x1, y1, x2, y2, lab, idx, doLog) Then
                If session IsNot Nothing AndAlso Not session.SelectionLogged Then
                    lab.Log("DVLINE", "READ_METHOD_SELECTED", "late_binding")
                    session.SelectionLogged = True
                End If
                Return True
            End If

            Return False
        End Function

        Public Shared Sub Run(app As Application,
                              draft As DraftDocument,
                              Optional sheet As Sheet = Nothing,
                              Optional effectiveRunMainDimensioning As Boolean = False,
                              Optional appLog As Action(Of String) = Nothing,
                              Optional outDimLabRoot As String = Nothing,
                              Optional forensicInteractive As Boolean = False,
                              Optional mode As DimLabMode = DimLabMode.Full,
                              Optional enableVisibleProbe As Boolean = False,
                              Optional enableAltPlacementLog As Boolean = False,
                              Optional runHorizontalControlInVerticalOnly As Boolean = True,
                              Optional keepFailedDimensions As Boolean = False,
                              Optional cleanPreviousLabDimensions As Boolean = True)

            Dim logAct = If(appLog, Sub(s As String) Console.WriteLine(s))
            Dim outRoot = If(String.IsNullOrWhiteSpace(outDimLabRoot),
                Path.Combine(System.Environment.CurrentDirectory, "OUT_DIMLAB"), outDimLabRoot.Trim())
            Dim lab As New DimLabLogger(outRoot, logAct)

            lab.Log("BOOT", "START", "outRoot=" & outRoot & " archivoTexto=" & lab.TextFilePath)
            lab.Log("RUN", "exclusive", "True")
            lab.Log("MODE", mode.ToString(), "")
            lab.Log("PRECHECK", "effectiveRunMainDimensioning", effectiveRunMainDimensioning.ToString(CultureInfo.InvariantCulture))
            lab.Log("PRECHECK", "autoDimensioningDisabled", "True")
            lab.Log("PRECHECK", "forensicInteractive", forensicInteractive.ToString(CultureInfo.InvariantCulture))
            lab.Log("PRECHECK", "enableVisibleProbe", enableVisibleProbe.ToString(CultureInfo.InvariantCulture))
            lab.Log("PRECHECK", "enableAltPlacementLog", enableAltPlacementLog.ToString(CultureInfo.InvariantCulture))

            If draft Is Nothing Then
                lab.Log("RUN", "ABORT", "draft=Nothing")
                Return
            End If

            Try
                draft.Activate()
            Catch exA As Exception
                lab.Log("DRAFT", "ACTIVATE_WARN", exA.Message)
            End Try

            Dim activeNm0 As String = SafeSheetNameFromDraft(draft)
            If IsTwoDModelSheetName(activeNm0) Then
                lab.Log("2DMODEL", "DETECTED", "activeSheet=" & activeNm0)
                lab.Log("2DMODEL", "SWITCH_TO_WORKING_SHEET", "")
            End If

            Dim sh As Sheet = ResolveWorkingDrawingSheet(draft, lab)
            If sh Is Nothing Then
                lab.Log("ABORT", "reason", "no_working_sheet_with_drawingviews")
                Return
            End If

            If Not TryActivateWorkingSheet(app, draft, sh, lab) Then
                lab.Log("ABORT", "reason", "cannot_activate_working_sheet_still_in_2d_model")
                Return
            End If

            If forensicInteractive Then
                ApplyInteractiveActivation(app, draft, sh, lab)
            End If

            Dim nDim0 As Integer = SafeDimensionsCount(sh)
            lab.Log("PRECHECK", "existingDimensionsBeforeLab", nDim0.ToString(CultureInfo.InvariantCulture))

            Dim created As New List(Of Object)()

            Dim sumRetrieve As String = "NO_ATTEMPT"
            Dim sumHoriz As String = "FAIL"
            Dim sumVert As String = "FAIL"
            Dim sumGap As String = "FAIL"
            Dim sumAux As String = "FAIL"
            Dim labCtx As LabRunContext = Nothing
            Dim abortedVis0 As Boolean = False

            Try
                Dim dims As Dimensions = Nothing
                Try
                    dims = CType(sh.Dimensions, Dimensions)
                Catch ex As Exception
                    lab.Log("RUN", "ABORT", "Dimensions: " & ex.Message)
                    GoTo Summary
                End Try

                labCtx = New LabRunContext With {
                    .App = app, .Draft = draft, .Sheet = sh,
                    .ForensicInteractive = forensicInteractive,
                    .EnableAltPlacementLog = enableAltPlacementLog,
                    .SummaryHorizontalViewLabel = "",
                    .SummaryVerticalViewLabel = "",
                    .SummaryVerticalExpectedHeight = "",
                    .SummaryVerticalReason = "",
                    .KeepFailedDimensions = keepFailedDimensions,
                    .IsCleanFullStrict = (mode = DimLabMode.CleanFullStrict)
                }
                lab.Log("PLACE", "GAP_CONFIG", "horizontalGap=" & Gd(DimLabHorizontalGap) &
                        " verticalGap=" & Gd(DimLabVerticalGap) &
                        " minGap=" & Gd(DimLabMinGap) &
                        " maxGap=" & Gd(DimLabMaxGap))
                ResolveDimLabDimensionStyle(draft, lab, labCtx)

                lab.Log("MODE", mode.ToString(), "")
                lab.Log("PRECHECK", "DimLabKeepFailedDimensions", keepFailedDimensions.ToString(CultureInfo.InvariantCulture))
                lab.Log("PRECHECK", "DimLabCleanPreviousLabDimensions", cleanPreviousLabDimensions.ToString(CultureInfo.InvariantCulture))

                If cleanPreviousLabDimensions Then
                    CleanStartDimensions(sh, lab)
                End If

                If enableVisibleProbe AndAlso mode <> DimLabMode.CleanFull AndAlso mode <> DimLabMode.CleanFullStrict Then
                    If Not RunVisibleZeroProbe(sh, draft, app, dims, created, lab, labCtx) Then
                        abortedVis0 = True
                        If forensicInteractive OrElse mode = DimLabMode.ForensicHorizontal Then
                            lab.Log("ABORT", "reason", "plain_sheet_dimension_not_visible")
                            GoTo Summary
                        End If
                    ElseIf forensicInteractive Then
                        TrySelectAddForensic(draft, labCtx, lab, "AFTER_VIS0")
                    End If
                Else
                    lab.Log("VIS0", "SKIP", "EnableDimLabVisibleProbe=False")
                End If

                Dim dv As DrawingView = FindDrawingViewForLab(sh, draft, lab)
                If dv Is Nothing Then
                    lab.Log("RUN", "ABORT", "no_suitable_drawing_view")
                    GoTo Summary
                End If

                Dim horizIdx = GetDrawingViewSheetIndex(sh, dv)
                lab.Log("VIEW_SELECT", "HORIZONTAL", "idx=" & horizIdx.ToString(CultureInfo.InvariantCulture) & " name=" & SafeDrawingViewName(dv))
                labCtx.SummaryHorizontalViewLabel = "DrawingView " & horizIdx.ToString(CultureInfo.InvariantCulture)

                Dim lineReadSession As New LabLineReadSession()
                DumpDVLines(dv, lab, lineReadSession)

                Dim runFullRetrieve As Boolean = (mode = DimLabMode.Full) AndAlso Not forensicInteractive
                If runFullRetrieve Then
                    sumRetrieve = RunRetrieveDimensionsTest(dv, draft, lab)
                Else
                    sumRetrieve = If(forensicInteractive, "SKIPPED_FORENSIC", "SKIPPED_MODE")
                    lab.Log("RUN", "SKIP", "RetrieveDimensions mode=" & mode.ToString())
                End If

                Dim bounds As LineBounds = BuildDVLineEndpointBounds(dv, lab, lineReadSession)
                If bounds Is Nothing Then
                    GoTo Summary
                End If

                Dim dvBounds = ToDVBounds(bounds)
                lab.Log("DVLINE", "VALIDATION", "pre_dimension_bounds_ok width=" & bounds.ExpectedWidth.ToString("G17", CultureInfo.InvariantCulture) &
                        " height=" & bounds.ExpectedHeight.ToString("G17", CultureInfo.InvariantCulture))

                Dim tolB = Math.Max(Math.Max(bounds.ExpectedWidth, bounds.ExpectedHeight) * 0.002R, 0.000001R)
                lab.Log("DVREF", "BOUNDARY_TOL", "tol=" & tolB.ToString("G17", CultureInfo.InvariantCulture))

                Dim candidates As New List(Of DVEndpointCandidate)()
                BuildEndpointCandidates(dv, candidates, lab, lineReadSession)
                If candidates.Count < 2 Then
                    lab.Log("DVREF", "CANDIDATES", "count_insufficient=" & candidates.Count.ToString(CultureInfo.InvariantCulture))
                    GoTo Summary
                End If

                Dim leftL As New List(Of DVEndpointCandidate)()
                Dim rightL As New List(Of DVEndpointCandidate)()
                Dim bottomL As New List(Of DVEndpointCandidate)()
                Dim topL As New List(Of DVEndpointCandidate)()
                ClassifyBoundaries(candidates, bounds, tolB, leftL, rightL, bottomL, topL, lab)

                Dim runHorizontal =
                    mode = DimLabMode.HorizontalOnly OrElse
                    mode = DimLabMode.ForensicHorizontal OrElse
                    mode = DimLabMode.CleanFull OrElse
                    mode = DimLabMode.CleanFullStrict OrElse
                    mode = DimLabMode.Full OrElse
                    (mode = DimLabMode.VerticalOnly AndAlso runHorizontalControlInVerticalOnly)

                If runHorizontal Then
                    sumHoriz = RunHorizontalExclusiveForLab(dims, dv, draft, dvBounds, bounds, leftL, rightL, created, lab, labCtx)
                    If labCtx.DvRefHorizontalDim IsNot Nothing Then
                        Dim vf = ValidateVisibleFinal(labCtx.DvRefHorizontalDim, sh, draft, app, "HorizontalTotal_TopPair_DVRef", lab)
                        labCtx.DvRefHorizGraphicsMaterialized = vf.Materialized
                    End If
                Else
                    sumHoriz = "SKIPPED_MODE"
                    labCtx.SummaryHorizCreate = "SKIPPED"
                    lab.Log("HORIZONTAL", "SKIP", "reason=VerticalOnly_horizontal_control_disabled")
                End If
                If mode = DimLabMode.CleanFull OrElse mode = DimLabMode.CleanFullStrict Then
                    lab.Log("CLEANFULL", "HORIZONTAL", "result=" & sumHoriz)
                End If

                Dim runVertical As Boolean = (mode = DimLabMode.VerticalOnly OrElse mode = DimLabMode.Full OrElse mode = DimLabMode.CleanFull OrElse mode = DimLabMode.CleanFullStrict)
                If runVertical Then
                    sumVert = RunVerticalTotalMultiViewLab(dims, draft, sh, dv, created, lab, labCtx)
                Else
                    sumVert = "SKIPPED_MODE"
                    labCtx.SummaryVertCreate = "SKIPPED"
                    lab.Log("VERTICAL", "SKIP", "reason=" & mode.ToString())
                End If

                If mode = DimLabMode.CleanFull OrElse mode = DimLabMode.CleanFullStrict Then
                    lab.Log("CLEANFULL", "VERTICAL", "result=" & sumVert)
                    If sumHoriz.IndexOf("SUCCESS", StringComparison.OrdinalIgnoreCase) >= 0 AndAlso sumVert.IndexOf("SUCCESS", StringComparison.OrdinalIgnoreCase) >= 0 Then
                        lab.Log("CLEANFULL", "STOP", "reason=horizontal_and_vertical_success")
                    End If
                End If

                Dim runGapAux As Boolean = (mode = DimLabMode.Full AndAlso Not forensicInteractive)
                If runGapAux Then
                    sumGap = RunSmallGapTest(dims, dv, draft, created, lab, lineReadSession, labCtx)
                    sumAux = RunAuxiliaryLine2dFallbackTest(dims, sh, draft, dv, created, lab, lineReadSession)
                Else
                    sumGap = "SKIPPED_MODE" : sumAux = "SKIPPED_MODE"
                    lab.Log("RUN", "SKIP", "gap_aux mode=" & mode.ToString())
                End If

                If forensicInteractive Then
                    TrySelectAddForensic(draft, labCtx, lab, "AFTER_DVREF_HORIZONTAL")
                End If

            Catch ex As COMException
                lab.Log("RUN", "EXCEPTION_COM", "hr=0x" & ex.ErrorCode.ToString("X8", CultureInfo.InvariantCulture) & " " & ex.ToString())
            Catch ex As Exception
                lab.Log("RUN", "EXCEPTION", ex.ToString())
            End Try

Summary:
            Try
                If draft IsNot Nothing Then
                    draft.UpdateAll(True)
                    lab.Log("UPDATE", "DRAFT_UPDATEALL_AFTER_DIMENSIONS", "OK")
                End If
            Catch exU As Exception
                lab.Log("UPDATE", "DRAFT_UPDATEALL_AFTER_DIMENSIONS", exU.Message)
            End Try
            Try
                If app IsNot Nothing Then
                    app.DoIdle()
                    lab.Log("UPDATE", "DOIDLE_AFTER_DIMENSIONS", "OK")
                End If
            Catch exI As Exception
                lab.Log("UPDATE", "DOIDLE_AFTER_DIMENSIONS", exI.Message)
            End Try

            If labCtx IsNot Nothing Then
                Dim stylePass = ApplyResolvedStyleToAllSheetDimensions(sh, labCtx, lab)
                If labCtx.IsCleanFullStrict AndAlso Not stylePass Then
                    sumHoriz = "STYLE_FAIL"
                    sumVert = "STYLE_FAIL"
                    labCtx.SummaryHorizCreate = "STYLE_FAIL"
                    labCtx.SummaryVertCreate = "STYLE_FAIL"
                End If
            End If

            Dim rec As String = BuildRecommendedNextStepV2(labCtx, sumHoriz, sumVert)
            lab.Log("SUMMARY", "RetrieveDimensions", sumRetrieve)
            lab.Log("SUMMARY", "DVRef_HorizontalTotal", sumHoriz)
            lab.Log("SUMMARY", "DVRef_VerticalTotal", sumVert)
            lab.Log("SUMMARY", "DVRef_SmallGap", sumGap)
            lab.Log("SUMMARY", "AuxLine2dFallback", sumAux)

            lab.Log("SUMMARY", "Mode", mode.ToString())
            If labCtx IsNot Nothing Then
                lab.Log("SUMMARY", "Horizontal_Create", labCtx.SummaryHorizCreate)
                lab.Log("SUMMARY", "Horizontal_Value", If(String.IsNullOrEmpty(labCtx.SummaryHorizValue), "N/A", labCtx.SummaryHorizValue))
                lab.Log("SUMMARY", "Horizontal_Visible", labCtx.SummaryHorizVisible.ToString(CultureInfo.InvariantCulture))
                lab.Log("SUMMARY", "Horizontal_Connected", labCtx.SummaryHorizConnected.ToString(CultureInfo.InvariantCulture))
                lab.Log("SUMMARY", "Vertical_Create", labCtx.SummaryVertCreate)
                lab.Log("SUMMARY", "Vertical_Value", If(String.IsNullOrEmpty(labCtx.SummaryVertValue), "N/A", labCtx.SummaryVertValue))
                lab.Log("SUMMARY", "Vertical_Visible", labCtx.SummaryVertVisible.ToString(CultureInfo.InvariantCulture))
                lab.Log("SUMMARY", "Vertical_Connected", labCtx.SummaryVertConnected.ToString(CultureInfo.InvariantCulture))
                lab.Log("SUMMARY", "Vertical_ValueClass", If(String.IsNullOrEmpty(labCtx.SummaryVertValueClass), "N/A", labCtx.SummaryVertValueClass))

                Dim hMv = If(String.IsNullOrEmpty(labCtx.SummaryHorizontalViewLabel), "N/A", labCtx.SummaryHorizontalViewLabel)
                Dim vMv = If(String.IsNullOrEmpty(labCtx.SummaryVerticalViewLabel), "N/A", labCtx.SummaryVerticalViewLabel)
                Dim hResLv = SummarizeHorizontalLabOutcome(sumHoriz)
                Dim vResLv = SummarizeVerticalLabOutcome(sumVert, labCtx)
                Dim vExpect = If(String.IsNullOrEmpty(labCtx.SummaryVerticalExpectedHeight), "N/A", labCtx.SummaryVerticalExpectedHeight)
                Dim vReason = If(String.IsNullOrEmpty(labCtx.SummaryVerticalReason), "N/A", labCtx.SummaryVerticalReason)

                lab.Log("SUMMARY", "Horizontal_View", hMv)
                lab.Log("SUMMARY", "Horizontal_Result", hResLv)
                lab.Log("SUMMARY", "Vertical_View", vMv)
                lab.Log("SUMMARY", "Vertical_Expected", vExpect)
                lab.Log("SUMMARY", "Vertical_Result", vResLv)
                lab.Log("SUMMARY", "Vertical_Reason", vReason)
                lab.Log("SUMMARY", "DimensionStyle", If(String.IsNullOrWhiteSpace(labCtx.ResolvedStyleName), "N/A", labCtx.ResolvedStyleName))
                lab.Log("SUMMARY", "GapHorizontal", Gd(DimLabHorizontalGap))
                lab.Log("SUMMARY", "GapVertical", Gd(DimLabVerticalGap))
                lab.Log("SUMMARY", "Horizontal_Gap", Gd(DimLabHorizontalGap))
                lab.Log("SUMMARY", "Vertical_Gap", Gd(DimLabVerticalGap))
                lab.Log("SUMMARY", "TextCenterHorizontal", labCtx.TextCenterHorizontalResult)
                lab.Log("SUMMARY", "TextCenterVertical", labCtx.TextCenterVerticalResult)
                lab.Log("SUMMARY", "Horizontal_TextCenter", labCtx.TextCenterHorizontalResult)
                lab.Log("SUMMARY", "Vertical_TextCenter", labCtx.TextCenterVerticalResult)
                lab.Log("SUMMARY", "AxisMode_Default_Result", labCtx.AxisModeDefaultResult)
                lab.Log("SUMMARY", "AxisMode_Implied_Result", labCtx.AxisModeImpliedResult)
                lab.Log("SUMMARY", "AxisMode_Explicit_Result", labCtx.AxisModeExplicitResult)
                lab.Log("SUMMARY", "BestVerticalAxisMode", labCtx.BestVerticalAxisMode)
                lab.Log("SUMMARY", "DimensionStyleRequested", labCtx.RequestedStyleName)
                lab.Log("SUMMARY", "Horizontal_StyleFinal", labCtx.SummaryHorizontalStyleFinal)
                lab.Log("SUMMARY", "Vertical_StyleFinal", labCtx.SummaryVerticalStyleFinal)
                Dim styleOkH = String.Equals(NormalizeStyleNameForCompare(labCtx.SummaryHorizontalStyleFinal), NormalizeStyleNameForCompare(labCtx.RequestedStyleName), StringComparison.OrdinalIgnoreCase)
                Dim styleOkV = String.Equals(NormalizeStyleNameForCompare(labCtx.SummaryVerticalStyleFinal), NormalizeStyleNameForCompare(labCtx.RequestedStyleName), StringComparison.OrdinalIgnoreCase)
                labCtx.SummaryStyleResult = If(styleOkH AndAlso styleOkV, "SUCCESS", "FAIL")
                lab.Log("SUMMARY", "StyleResult", labCtx.SummaryStyleResult)
                Dim finalResult = "SUCCESS"
                If labCtx.SummaryStyleResult <> "SUCCESS" Then
                    finalResult = "FAIL_STYLE"
                ElseIf sumHoriz.IndexOf("SUCCESS", StringComparison.OrdinalIgnoreCase) < 0 OrElse sumVert.IndexOf("SUCCESS", StringComparison.OrdinalIgnoreCase) < 0 Then
                    finalResult = "FAIL"
                End If
                lab.Log("SUMMARY", "FinalResult", finalResult)
            Else
                lab.Log("SUMMARY", "Horizontal_Create", "N/A")
                lab.Log("SUMMARY", "Vertical_Create", "N/A")
                lab.Log("SUMMARY", "FinalResult", "FAIL")
            End If
            lab.Log("SUMMARY", "RecommendedNextStep", rec)

            If labCtx IsNot Nothing AndAlso labCtx.ForensicInteractive Then
                lab.Log("SUMMARY", "PlainSheet_TextVisibleProbe", labCtx.PlainTextProbe)
                lab.Log("SUMMARY", "PlainSheet_LineVisibleProbe", labCtx.PlainLineProbe)
                lab.Log("SUMMARY", "PlainSheet_AddLength_Create", labCtx.PlainAddLengthCreate)
                lab.Log("SUMMARY", "PlainSheet_AddLength_GraphicsMaterialized", labCtx.PlainGraphicsMaterialized.ToString(CultureInfo.InvariantCulture))
                lab.Log("SUMMARY", "DVRef_HorizontalTotal_GraphicsMaterialized", labCtx.DvRefHorizGraphicsMaterialized.ToString(CultureInfo.InvariantCulture))
                lab.Log("REOPEN", "SKIP", "reason=interactive_forensic_prefer_manual_review")
            End If

            If labCtx IsNot Nothing Then
                If labCtx.Vis.HorizTopDone Then
                    lab.Log("SUMMARY", "DVRef_HorizontalTotal_Create", labCtx.Vis.HorizTopCreate)
                    lab.Log("SUMMARY", "DVRef_HorizontalTotal_Visible", labCtx.Vis.HorizTopVisible.ToString(CultureInfo.InvariantCulture))
                    lab.Log("SUMMARY", "DVRef_HorizontalTotal_Range", labCtx.Vis.HorizTopRange)
                    lab.Log("SUMMARY", "DVRef_HorizontalTotal_TrackDistance", labCtx.Vis.HorizTopTrack)
                    If Not labCtx.Vis.HorizTopVisible Then
                        lab.Log("SUMMARY", "DVRef_HorizontalTotal_VisibleReason", If(String.IsNullOrEmpty(labCtx.Vis.HorizTopReason), "range_zero_or_outside_sheet", labCtx.Vis.HorizTopReason))
                    End If
                Else
                    lab.Log("SUMMARY", "DVRef_HorizontalTotal_Create", sumHoriz)
                    lab.Log("SUMMARY", "DVRef_HorizontalTotal_Visible", "N/A")
                    lab.Log("SUMMARY", "DVRef_HorizontalTotal_Range", "N/A")
                    lab.Log("SUMMARY", "DVRef_HorizontalTotal_TrackDistance", "N/A")
                End If
                If labCtx.Vis.VertRecorded Then
                    lab.Log("SUMMARY", "DVRef_VerticalTotal_Create", labCtx.Vis.VertBestCreate)
                    lab.Log("SUMMARY", "DVRef_VerticalTotal_Visible", labCtx.Vis.VertBestVisible.ToString(CultureInfo.InvariantCulture))
                    lab.Log("SUMMARY", "DVRef_VerticalTotal_Range", labCtx.Vis.VertBestRange)
                    lab.Log("SUMMARY", "DVRef_VerticalTotal_TrackDistance", labCtx.Vis.VertBestTrack)
                End If
                If labCtx.Vis.GapDone Then
                    lab.Log("SUMMARY", "DVRef_SmallGap_Create", labCtx.Vis.GapCreate)
                    lab.Log("SUMMARY", "DVRef_SmallGap_Visible", labCtx.Vis.GapVisible.ToString(CultureInfo.InvariantCulture))
                    lab.Log("SUMMARY", "DVRef_SmallGap_Range", labCtx.Vis.GapRange)
                End If
            End If

            If (mode = DimLabMode.CleanFull OrElse mode = DimLabMode.CleanFullStrict) AndAlso sh IsNot Nothing AndAlso labCtx IsNot Nothing Then
                CleanupIntermediateDimensionsKeepingPrimary(sh, labCtx, lab)
            End If

            If sh IsNot Nothing Then
                lab.Log("SUMMARY", "SheetDimensionsFinalCount", SafeDimensionsCount(sh).ToString(CultureInfo.InvariantCulture))
                lab.Log("SUMMARY", "CountInterpretation", "SheetDimensionsFinalCount=collection_only_not_visual_proof")
                LogFinalDimensionPositionsCompact(sh, lab)
            End If

            If forensicInteractive AndAlso abortedVis0 Then
                lab.Log("SUMMARY", "AbortNote", "ABORT_after_VIS0_plain_sheet_not_materialized")
            End If

            If forensicInteractive Then
                Try
                    System.Windows.Forms.MessageBox.Show(
                        "DIMLAB terminado. Revisa visualmente si aparecen: texto, línea y cota auxiliar. No cierres Solid Edge todavía.",
                        "DIMLAB forense",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Information)
                Catch
                End Try
                lab.Log("INTERACTIVE", "PAUSE", "MsgBox_DIMLAB_forensic_ok")
            End If

            ' Pase final obligatorio: reaplicar estilo justo antes de terminar el laboratorio.
            If labCtx IsNot Nothing Then
                lab.Log("STYLE", "LIST_AVAILABLE_NEAR_END", "start")
                LogAvailableStylesNearEnd(draft, sh, lab)
                lab.Log("STYLE", "FINAL_PASS_BEFORE_END", "start requested=" & labCtx.RequestedStyleName)
                Dim finalPassOk = ApplyResolvedStyleToAllSheetDimensions(sh, labCtx, lab)
                lab.Log("STYLE", "FINAL_PASS_BEFORE_END", "result=" & finalPassOk.ToString(CultureInfo.InvariantCulture))
                If labCtx.IsCleanFullStrict AndAlso Not finalPassOk Then
                    lab.Log("STYLE", "FINAL_PASS_BEFORE_END", "strict_fail=True")
                End If
            End If

            ' No borrar lo creado al terminar: antes CleanupCreated() eliminaba todas las cotas de prueba en modo normal.
            If forensicInteractive Then
                lab.Log("CLEANUP", "SKIP", "forensic_keep_all_objects")
            Else
                lab.Log("CLEANUP", "SKIP", "keep_all_after_dimlab=true")
            End If
        End Sub

        Private Shared Function BuildRecommended(h As String, v As String, g As String, r As String) As String
            If h.IndexOf("SUCCESS", StringComparison.OrdinalIgnoreCase) >= 0 AndAlso
               v.IndexOf("SUCCESS", StringComparison.OrdinalIgnoreCase) >= 0 Then
                Return "crear_motor_experimental_DVLine2d_Reference_para_totales_exteriores_UNE"
            End If
            If h.IndexOf("SUCCESS", StringComparison.OrdinalIgnoreCase) >= 0 AndAlso
               v.IndexOf("SUCCESS", StringComparison.OrdinalIgnoreCase) < 0 Then
                Return "DVLine2d_Reference_validado_horizontalmente_seguir_investigando_vertical"
            End If
            If g.IndexOf("SUCCESS", StringComparison.OrdinalIgnoreCase) >= 0 Then
                Return "DVLine2d_Reference_funciona_como_tecnologia_base"
            End If
            If r.IndexOf("E_ABORT", StringComparison.OrdinalIgnoreCase) >= 0 Then
                Return "aparcar_RetrieveDimensions_para_modelos_PMI_recuperables"
            End If
            Return "revisar_logs_DIMLAB"
        End Function

        Private NotInheritable Class VisibleFinalResult
            Public Materialized As Boolean
            Public Visible As Boolean
            Public Reason As String
            Public Uncertain As Boolean
        End Class

        Private Shared Function SafeSheetName(sh As Sheet) As String
            If sh Is Nothing Then Return ""
            Try
                Return Convert.ToString(sh.Name, CultureInfo.InvariantCulture).Trim()
            Catch
                Return ""
            End Try
        End Function

        Private Shared Function SafeSheetNameFromDraft(draft As DraftDocument) As String
            Try
                Return SafeSheetName(CType(draft.ActiveSheet, Sheet))
            Catch
                Return ""
            End Try
        End Function

        Private Shared Function IsTwoDModelSheetName(name As String) As Boolean
            If String.IsNullOrWhiteSpace(name) Then Return False
            Return String.Equals(name.Trim(), "2D Model", StringComparison.OrdinalIgnoreCase)
        End Function

        Private Shared Function SafeDrawingViewsCountSheet(sh As Sheet) As Integer
            If sh Is Nothing Then Return -1
            Try
                Return CInt(sh.DrawingViews.Count)
            Catch
                Return -1
            End Try
        End Function

        Private Shared Function SafeLines2dCountSheet(sh As Sheet) As Integer
            If sh Is Nothing Then Return -1
            Try
                Return CInt(sh.Lines2d.Count)
            Catch
                Return -1
            End Try
        End Function

        ''' <summary>Pliego de trabajo con vistas: excluye pestaña 2D Model; prioriza Hoja1 si tiene DrawingViews.</summary>
        Private Shared Function ResolveWorkingDrawingSheet(draft As DraftDocument, lab As DimLabLogger) As Sheet
            Dim candidates As New List(Of Sheet)()
            Dim preferred As Sheet = Nothing

            Dim visit = Sub(si As Integer, shi As Integer, sh As Sheet)
                            Dim nm = SafeSheetName(sh)
                            Dim nDv = SafeDrawingViewsCountSheet(sh)
                            Dim nLn = SafeLines2dCountSheet(sh)
                            Dim nDim = SafeDimensionsCount(sh)
                            lab.Log("SHEET_SCAN", "ROW", "sectionIndex=" & si.ToString(CultureInfo.InvariantCulture) &
                                    " sheetIndex=" & shi.ToString(CultureInfo.InvariantCulture) &
                                    " sheetName=" & nm &
                                    " drawingViews=" & nDv.ToString(CultureInfo.InvariantCulture) &
                                    " lines2d=" & nLn.ToString(CultureInfo.InvariantCulture) &
                                    " dimensions=" & nDim.ToString(CultureInfo.InvariantCulture))
                            If IsTwoDModelSheetName(nm) Then
                                lab.Log("SHEET_SCAN", "REJECT", "sheetName=" & nm & " reason=two_d_model_sheet")
                                Return
                            End If
                            If nDv > 0 Then
                                candidates.Add(sh)
                                If String.Equals(nm, "Hoja1", StringComparison.OrdinalIgnoreCase) AndAlso preferred Is Nothing Then
                                    preferred = sh
                                    lab.Log("SHEET_SCAN", "CANDIDATE", "sheetName=Hoja1 drawingViews=" & nDv.ToString(CultureInfo.InvariantCulture))
                                End If
                            End If
                        End Sub

            Dim anySections As Boolean = False
            Try
                Dim sections = draft.Sections
                Dim nSec = CInt(sections.Count)
                For si As Integer = 1 To nSec
                    anySections = True
                    Dim sec As SolidEdgeDraft.Section = Nothing
                    Try
                        sec = CType(sections.Item(si), SolidEdgeDraft.Section)
                    Catch
                        Continue For
                    End Try
                    If sec Is Nothing Then Continue For
                    Dim sheetsCol = sec.Sheets
                    Dim nSh = CInt(sheetsCol.Count)
                    For shi As Integer = 1 To nSh
                        Dim sh As Sheet = Nothing
                        Try
                            sh = CType(sheetsCol.Item(shi), Sheet)
                        Catch
                            Continue For
                        End Try
                        If sh Is Nothing Then Continue For
                        visit(si, shi, sh)
                    Next
                Next
            Catch ex As Exception
                lab.Log("SHEET_SCAN", "SECTIONS_ERR", ex.Message)
            End Try

            If Not anySections OrElse candidates.Count = 0 Then
                Try
                    Dim n = CInt(draft.Sheets.Count)
                    For i As Integer = 1 To n
                        Dim sh As Sheet = Nothing
                        Try : sh = CType(draft.Sheets.Item(i), Sheet) : Catch : Continue For : End Try
                        If sh Is Nothing Then Continue For
                        visit(0, i, sh)
                    Next
                Catch ex2 As Exception
                    lab.Log("SHEET_SCAN", "FALLBACK_SHEETS_ERR", ex2.Message)
                End Try
            End If

            Dim chosen As Sheet = If(preferred, If(candidates.Count > 0, candidates(0), Nothing))
            If chosen Is Nothing Then
                Return Nothing
            End If
            Dim finalName = SafeSheetName(chosen)
            Dim finalDv = SafeDrawingViewsCountSheet(chosen)
            lab.Log("SHEET", "RESOLVED_WORKING", "name=" & finalName & " drawingViews=" & finalDv.ToString(CultureInfo.InvariantCulture))
            Return chosen
        End Function

        Private Shared Function TryActivateWorkingSheet(app As Application, draft As DraftDocument, workingSh As Sheet, lab As DimLabLogger) As Boolean
            Dim targetName = SafeSheetName(workingSh)
            lab.Log("SHEET", "ACTIVATE_TRY", "name=" & targetName)
            Try
                draft.Activate()
            Catch exD As Exception
                lab.Log("SHEET", "ACTIVATE_WARN", "draft " & exD.Message)
            End Try
            Try
                workingSh.Activate()
            Catch exS As Exception
                lab.Log("SHEET", "ACTIVATE_FAIL", exS.Message)
                Return False
            End Try
            Try
                If app IsNot Nothing Then app.DoIdle()
            Catch
            End Try
            lab.Log("SHEET", "ACTIVATE_OK", "name=" & targetName)

            Dim an = SafeSheetNameFromDraft(draft)
            Dim dvAct = -1
            Try
                dvAct = SafeDrawingViewsCountSheet(CType(draft.ActiveSheet, Sheet))
            Catch
            End Try
            lab.Log("SHEET", "ACTIVE_AFTER_ACTIVATE", "name=" & an)
            lab.Log("SHEET", "DRAWINGVIEWS_AFTER_ACTIVATE", "count=" & dvAct.ToString(CultureInfo.InvariantCulture))
            If IsTwoDModelSheetName(an) Then
                Return False
            End If
            Return True
        End Function

        Private Shared Sub ApplyInteractiveActivation(app As Application, draft As DraftDocument, sh As Sheet, lab As DimLabLogger)
            Try
                draft.Activate()
            Catch ex As Exception
                lab.Log("INTERACTIVE", "DRAFT_ACTIVATE", "FAIL " & ex.Message)
            End Try
            Try
                If sh IsNot Nothing Then sh.Activate()
            Catch ex As Exception
                lab.Log("INTERACTIVE", "SHEET_ACTIVATE", "FAIL " & ex.Message)
            End Try
            Try
                If app IsNot Nothing Then app.DoIdle()
            Catch
            End Try
        End Sub

        Private Shared Function TryReadDisplayDataCounts(d As Object,
                                                         ByRef lineCount As Integer,
                                                         ByRef arcCount As Integer,
                                                         ByRef textCount As Integer,
                                                         ByRef pointCount As Integer) As Boolean
            lineCount = 0 : arcCount = 0 : textCount = 0 : pointCount = 0
            If d Is Nothing Then Return False
            Dim dd As Object = Nothing
            Try
                dd = CallByName(d, "GetDisplayData", CallType.Method)
            Catch
                Return False
            End Try
            If dd Is Nothing Then Return False
            Try : lineCount = CInt(CallByName(dd, "GetLineCount", CallType.Get)) : Catch : End Try
            If lineCount = 0 Then
                Try : lineCount = CInt(CallByName(dd, "GetLineCount", CallType.Method)) : Catch : End Try
            End If
            Try : arcCount = CInt(CallByName(dd, "GetArcCount", CallType.Get)) : Catch : End Try
            If arcCount = 0 Then
                Try : arcCount = CInt(CallByName(dd, "GetArcCount", CallType.Method)) : Catch : End Try
            End If
            Try : textCount = CInt(CallByName(dd, "GetTextCount", CallType.Get)) : Catch : End Try
            If textCount = 0 Then
                Try : textCount = CInt(CallByName(dd, "GetTextCount", CallType.Method)) : Catch : End Try
            End If
            Try : pointCount = CInt(CallByName(dd, "GetPointCount", CallType.Get)) : Catch : End Try
            Return True
        End Function

        Private Shared Sub LogDimDisplayDataInspect(d As Object, lab As DimLabLogger, name As String)
            Dim lc As Integer, ac As Integer, tc As Integer, pc As Integer
            If Not TryReadDisplayDataCounts(d, lc, ac, tc, pc) Then
                lab.Log("DIM", "DISPLAYDATA_FAIL", "name=" & name & " error=GetDisplayData_or_counts")
                Return
            End If
            lab.Log("DIM", "DISPLAYDATA", "name=" & name & " lineCount=" & lc.ToString(CultureInfo.InvariantCulture) &
                    " arcCount=" & ac.ToString(CultureInfo.InvariantCulture) &
                    " textCount=" & tc.ToString(CultureInfo.InvariantCulture) &
                    " pointCount=" & pc.ToString(CultureInfo.InvariantCulture))
        End Sub

        Private Shared Function ValidateVisibleFinal(dimObj As Dimension,
                                                     sheet As Sheet,
                                                     draft As DraftDocument,
                                                     app As Application,
                                                     testName As String,
                                                     lab As DimLabLogger) As VisibleFinalResult
            Dim r As New VisibleFinalResult With {.Reason = "unknown", .Materialized = False, .Visible = False, .Uncertain = False}
            Try
                If draft IsNot Nothing Then draft.UpdateAll(True)
            Catch ex As Exception
                lab.Log("VALIDATE_VISIBLE", "UPDATE_FAIL", testName & " " & ex.Message)
            End Try
            Try
                If app IsNot Nothing Then app.DoIdle()
            Catch
            End Try

            Dim rx1 As Double, ry1 As Double, rx2 As Double, ry2 As Double
            Dim rangeOk = TryReadDimensionRange(dimObj, rx1, ry1, rx2, ry2)
            Dim rngStr = If(rangeOk, FormatRangeStr(rx1, ry1, rx2, ry2), "unreadable")
            lab.Log("VALIDATE_VISIBLE", "RANGE_FINAL", "name=" & testName & " range=" & rngStr)

            Dim lc As Integer, ac As Integer, tc As Integer, pc As Integer
            Dim ddOk = TryReadDisplayDataCounts(dimObj, lc, ac, tc, pc)
            If ddOk Then
                lab.Log("VALIDATE_VISIBLE", "DISPLAYDATA_FINAL", "name=" & testName &
                        " lineCount=" & lc.ToString(CultureInfo.InvariantCulture) &
                        " arcCount=" & ac.ToString(CultureInfo.InvariantCulture) &
                        " textCount=" & tc.ToString(CultureInfo.InvariantCulture) &
                        " pointCount=" & pc.ToString(CultureInfo.InvariantCulture))
            Else
                lab.Log("VALIDATE_VISIBLE", "DISPLAYDATA_FINAL", "name=" & testName & " lineCount=0 arcCount=0 (read_fail)")
                lc = 0 : ac = 0
            End If

            Dim rangeDeg = Not rangeOk OrElse DimensionRangeIsInvalid(rx1, ry1, rx2, ry2)
            Dim hasGfx = (lc + ac) > 0
            If rangeDeg AndAlso hasGfx Then
                r.Uncertain = True
                r.Reason = "range_zero_but_displaydata_nonzero"
                r.Visible = False
                r.Materialized = True
                lab.Log("VALIDATE_VISIBLE", "result", "name=" & testName & " visible=uncertain materialized=True reason=" & r.Reason)
                Return r
            End If
            If rangeDeg AndAlso Not hasGfx Then
                r.Reason = "no_graphics_materialized"
                r.Materialized = False
                r.Visible = False
                lab.Log("VALIDATE_VISIBLE", "result", "name=" & testName & " visible=False reason=" & r.Reason)
                Return r
            End If

            Dim rsn = GetVisibilityReasonQuiet(dimObj, sheet)
            If String.Equals(rsn, "in_sheet", StringComparison.OrdinalIgnoreCase) Then
                r.Visible = True
                r.Materialized = True
                r.Reason = "in_sheet"
            ElseIf rangeDeg Then
                r.Reason = "dimension_range_zero"
                r.Visible = False
                r.Materialized = hasGfx
            Else
                r.Reason = rsn
                r.Visible = False
                r.Materialized = hasGfx OrElse Not rangeDeg
            End If
            lab.Log("VALIDATE_VISIBLE", "result", "name=" & testName & " visible=" & r.Visible.ToString(CultureInfo.InvariantCulture) &
                    " materialized=" & r.Materialized.ToString(CultureInfo.InvariantCulture) & " reason=" & r.Reason)
            Return r
        End Function

        Private Shared Function TryEnsureDimLabLayer(draft As DraftDocument, sheet As Sheet, lab As DimLabLogger) As Object
            Const layerName = "DIMLAB_VISIBLE_TEST"
            Dim layers As Object = Nothing
            Try
                layers = sheet.Layers
            Catch
            End Try
            If layers Is Nothing Then
                Try : layers = draft.Layers : Catch : End Try
            End If
            If layers Is Nothing Then
                lab.Log("VIS0", "LAYER", "FAIL no_Layers_collection")
                Return Nothing
            End If
            Try
                Return CallByName(layers, "Add", CallType.Method, layerName)
            Catch ex As Exception
                lab.Log("VIS0", "LAYER", "FAIL " & ex.Message)
                Return Nothing
            End Try
        End Function

        Private Shared Sub TrySetObjectLayer(target As Object, layerObj As Object, lab As DimLabLogger, tag As String)
            If target Is Nothing OrElse layerObj Is Nothing Then Return
            Try
                CallByName(target, "Layer", CallType.Let, layerObj)
            Catch ex As Exception
                lab.Log("VIS0", "LAYER_APPLY", "FAIL " & tag & " " & ex.Message)
            End Try
        End Sub

        Private Shared Sub ResolveDimLabDimensionStyle(draft As DraftDocument, lab As DimLabLogger, ctx As LabRunContext)
            Dim styleObj As Object = Nothing
            Dim styleName As String = ""
            Dim found = ResolveDimensionStyleObject(draft, ctx.Sheet, ctx.RequestedStyleName, lab, styleObj, styleName)
            If found Then
                ctx.ResolvedStyleObj = styleObj
                ctx.ResolvedStyleName = styleName
            Else
                ctx.ResolvedStyleName = "N/A"
            End If
        End Sub

        Private Shared Function ResolveDimensionStyleObject(draft As DraftDocument, sheet As Sheet, preferredName As String, lab As DimLabLogger, ByRef styleObj As Object, ByRef styleName As String) As Boolean
            styleObj = Nothing
            styleName = ""
            If draft Is Nothing AndAlso sheet Is Nothing Then Return False
            Return ResolveDimStyleForDimension(draft, sheet, preferredName, lab, styleObj, styleName)
        End Function

        Private Shared Function ResolveDimStyleForDimension(draft As DraftDocument, sheet As Sheet, preferredName As String, lab As DimLabLogger, ByRef styleObj As Object, ByRef styleName As String) As Boolean
            styleObj = Nothing
            styleName = ""
            Dim found As Boolean = False
            Dim foundObj As Object = Nothing
            Dim foundName As String = ""
            Dim foundSource As String = ""

            Dim collectionNames = New String() {"DimensionStyles", "DimStyles", "LinearStyles", "Styles"}

            If draft IsNot Nothing Then
                For Each coll In collectionNames
                    ScanStyleCollectionForMatch(draft, "draft", coll, preferredName, lab, found, foundObj, foundName, foundSource)
                Next
            End If
            If sheet IsNot Nothing Then
                For Each coll In collectionNames
                    ScanStyleCollectionForMatch(sheet, "sheet", coll, preferredName, lab, found, foundObj, foundName, foundSource)
                Next
            End If

            If found AndAlso foundObj IsNot Nothing Then
                styleObj = foundObj
                styleName = foundName
                lab.Log("STYLE_RESOLVE_MATCH", "FOUND", "requested=" & preferredName & " resolved=" & styleName & " source=" & foundSource)
                lab.Log("STYLE", "RESOLVE", "requested=" & preferredName & " found=True resolvedName=" & styleName)
                Return True
            End If

            lab.Log("STYLE_RESOLVE_FAIL", "NOT_FOUND", "requested=" & preferredName)
            lab.Log("STYLE", "RESOLVE", "requested=" & preferredName & " found=False resolvedName=")
            Return False
        End Function

        Private Shared Sub ScanStyleCollectionForMatch(owner As Object,
                                                        ownerLabel As String,
                                                        collectionName As String,
                                                        preferredName As String,
                                                        lab As DimLabLogger,
                                                        ByRef found As Boolean,
                                                        ByRef foundObj As Object,
                                                        ByRef foundName As String,
                                                        ByRef foundSource As String)
            Dim styles As Object = Nothing
            Dim count As Integer = -1
            Try
                styles = CallByName(owner, collectionName, CallType.Get)
                If styles IsNot Nothing Then
                    count = CInt(CallByName(styles, "Count", CallType.Get))
                    lab.Log("STYLE_SCAN_COLLECTION", "OPEN", "name=" & collectionName & " ok=True count=" & count.ToString(CultureInfo.InvariantCulture) & " source=" & ownerLabel)
                    lab.Log("SCREEN", "STYLE_SEARCH", "source=" & ownerLabel & " collection=" & collectionName & " ok=True count=" & count.ToString(CultureInfo.InvariantCulture))
                Else
                    lab.Log("STYLE_SCAN_COLLECTION", "OPEN", "name=" & collectionName & " ok=False count=0 source=" & ownerLabel)
                    lab.Log("SCREEN", "STYLE_SEARCH", "source=" & ownerLabel & " collection=" & collectionName & " ok=False")
                    Return
                End If
            Catch ex As Exception
                lab.Log("STYLE_SCAN_COLLECTION", "OPEN", "name=" & collectionName & " ok=False source=" & ownerLabel & " error=" & ex.Message)
                lab.Log("SCREEN", "STYLE_SEARCH", "source=" & ownerLabel & " collection=" & collectionName & " ok=False")
                Return
            End Try

            For i As Integer = 1 To count
                Dim it As Object = Nothing
                Try
                    it = CallByName(styles, "Item", CallType.Method, i)
                Catch
                    Continue For
                End Try
                If it Is Nothing Then Continue For

                Dim nm As String = ""
                Try
                    nm = Convert.ToString(CallByName(it, "Name", CallType.Get), CultureInfo.InvariantCulture)
                Catch
                    nm = ""
                End Try
                lab.Log("STYLE_SCAN", "ITEM", "collection=" & collectionName & " idx=" & i.ToString(CultureInfo.InvariantCulture) & " name=" & nm & " source=" & ownerLabel)
                lab.Log("SCREEN", "STYLE_SEARCH", "source=" & ownerLabel & " collection=" & collectionName & " idx=" & i.ToString(CultureInfo.InvariantCulture) & " name=" & nm)
                lab.Log("STYLE_RESOLVE_TRY", "MATCH", "requested=" & preferredName & " collection=" & collectionName & " candidate=" & nm)

                If Not found AndAlso String.Equals(NormalizeStyleNameForCompare(nm), NormalizeStyleNameForCompare(preferredName), StringComparison.OrdinalIgnoreCase) Then
                    found = True
                    foundObj = it
                    foundName = nm
                    foundSource = ownerLabel & "." & collectionName
                    lab.Log("STYLE_RESOLVE_MATCH", "CANDIDATE", "requested=" & preferredName & " resolved=" & nm & " source=" & foundSource)
                End If
            Next
        End Sub

        Private Shared Function ApplyDimensionStyleStrict(dimObj As Object, styleObj As Object, styleName As String, lab As DimLabLogger, dimLabel As String, cleanFullStrict As Boolean) As Boolean
            If dimObj Is Nothing Then Return False
            Dim finalName = ""

            Try
                lab.Log("STYLE_APPLY_TRY", "ROUTE", "dim=" & dimLabel & " route=Style_method_string value=" & styleName)
                CallByName(dimObj, "Style", CallType.Method, styleName)
                finalName = ReadDimensionStyleName(dimObj, lab)
                If String.Equals(NormalizeStyleNameForCompare(finalName), NormalizeStyleNameForCompare(styleName), StringComparison.OrdinalIgnoreCase) Then
                    lab.Log("STYLE_APPLY_OK", "ROUTE", "dim=" & dimLabel & " route=Style_method_string final=" & finalName)
                    lab.Log("STYLE_FINAL", "DIM", "dim=" & dimLabel & " final=" & finalName)
                    Return True
                End If
            Catch ex As Exception
                lab.Log("STYLE_APPLY_FAIL", "ROUTE", "dim=" & dimLabel & " route=Style_method_string error=" & ex.Message)
            End Try

            Dim routes = New String() {"StyleName", "DimensionStyleName", "DimStyleName", "Style"}
            For Each route In routes
                Try
                    lab.Log("STYLE_APPLY_TRY", "ROUTE", "dim=" & dimLabel & " route=" & route & "_string value=" & styleName)
                    CallByName(dimObj, route, CallType.Let, styleName)
                    finalName = ReadDimensionStyleName(dimObj, lab)
                    If String.Equals(NormalizeStyleNameForCompare(finalName), NormalizeStyleNameForCompare(styleName), StringComparison.OrdinalIgnoreCase) Then
                        lab.Log("STYLE_APPLY_OK", "ROUTE", "dim=" & dimLabel & " route=" & route & "_string final=" & finalName)
                        lab.Log("STYLE_FINAL", "DIM", "dim=" & dimLabel & " final=" & finalName)
                        Return True
                    End If
                Catch ex As Exception
                    lab.Log("STYLE_APPLY_FAIL", "ROUTE", "dim=" & dimLabel & " route=" & route & "_string error=" & ex.Message)
                End Try
            Next

            Try
                lab.Log("STYLE_APPLY_TRY", "ROUTE", "dim=" & dimLabel & " route=SetStyle_method value=" & styleName)
                CallByName(dimObj, "SetStyle", CallType.Method, styleName)
                finalName = ReadDimensionStyleName(dimObj, lab)
                If String.Equals(NormalizeStyleNameForCompare(finalName), NormalizeStyleNameForCompare(styleName), StringComparison.OrdinalIgnoreCase) Then
                    lab.Log("STYLE_APPLY_OK", "ROUTE", "dim=" & dimLabel & " route=SetStyle_method final=" & finalName)
                    lab.Log("STYLE_FINAL", "DIM", "dim=" & dimLabel & " final=" & finalName)
                    Return True
                End If
            Catch ex As Exception
                lab.Log("STYLE_APPLY_FAIL", "ROUTE", "dim=" & dimLabel & " route=SetStyle_method error=" & ex.Message)
            End Try

            If styleObj IsNot Nothing Then
                lab.Log("STYLE_APPLY_TRY", "ROUTE", "dim=" & dimLabel & " route=Style_object")
                Try
                    CallByName(dimObj, "Style", CallType.Let, styleObj)
                Catch ex As Exception
                    lab.Log("STYLE_APPLY_FAIL", "ROUTE", "dim=" & dimLabel & " route=Style_object error=" & ex.Message)
                End Try
                Try
                    CallByName(dimObj, "Style", CallType.Set, styleObj)
                Catch ex As Exception
                    lab.Log("STYLE_APPLY_FAIL", "ROUTE", "dim=" & dimLabel & " route=Style_object_set error=" & ex.Message)
                End Try
            End If

            finalName = ReadDimensionStyleName(dimObj, lab)
            If String.Equals(NormalizeStyleNameForCompare(finalName), NormalizeStyleNameForCompare(styleName), StringComparison.OrdinalIgnoreCase) Then
                lab.Log("STYLE_APPLY_OK", "ROUTE", "dim=" & dimLabel & " route=Style_object final=" & finalName)
                lab.Log("STYLE_FINAL", "DIM", "dim=" & dimLabel & " final=" & finalName)
                Return True
            End If

            lab.Log("STYLE", "FINAL_FAIL", "name=" & dimLabel & " requested=" & styleName & " final=" & finalName)
            If cleanFullStrict Then
                lab.Log("STYLE", "FATAL", "CleanFullStrict requiere U3,5. No aceptar fallback.")
            End If
            lab.Log("STYLE_FINAL", "DIM", "dim=" & dimLabel & " final=" & finalName)
            Return False
        End Function

        Private Shared Sub TryApplyStyleToDimensionsCollectionLab(dims As Dimensions, styleName As String, styleObj As Object, lab As DimLabLogger, sourceTag As String)
            If dims Is Nothing Then Return
            Try
                lab.Log("STYLE_PRECREATE_TRY", "ROUTE", "route=Dimensions.Style(Method) value=" & styleName & " source=" & sourceTag)
                CallByName(dims, "Style", CallType.Method, styleName)
                lab.Log("STYLE_PRECREATE_RESULT", "ROUTE", "route=Dimensions.Style(Method) ok=True source=" & sourceTag)
                Return
            Catch ex As Exception
                lab.Log("STYLE_PRECREATE_RESULT", "ROUTE", "route=Dimensions.Style(Method) ok=False source=" & sourceTag & " error=" & ex.Message)
            End Try
            Try
                lab.Log("STYLE_PRECREATE_TRY", "ROUTE", "route=Dimensions.Style(Let) value=" & styleName & " source=" & sourceTag)
                CallByName(dims, "Style", CallType.Let, styleName)
                lab.Log("STYLE_PRECREATE_RESULT", "ROUTE", "route=Dimensions.Style(Let) ok=True source=" & sourceTag)
                Return
            Catch ex As Exception
                lab.Log("STYLE_PRECREATE_RESULT", "ROUTE", "route=Dimensions.Style(Let) ok=False source=" & sourceTag & " error=" & ex.Message)
            End Try
            If styleObj IsNot Nothing Then
                Try
                    lab.Log("STYLE_PRECREATE_TRY", "ROUTE", "route=Dimensions.Style(object) source=" & sourceTag)
                    CallByName(dims, "Style", CallType.Let, styleObj)
                    lab.Log("STYLE_PRECREATE_RESULT", "ROUTE", "route=Dimensions.Style(object) ok=True source=" & sourceTag)
                Catch ex As Exception
                    lab.Log("STYLE_PRECREATE_RESULT", "ROUTE", "route=Dimensions.Style(object) ok=False source=" & sourceTag & " error=" & ex.Message)
                End Try
            End If
        End Sub

        Private Shared Function ApplyResolvedStyleToAllSheetDimensions(sh As Sheet, ctx As LabRunContext, lab As DimLabLogger) As Boolean
            If sh Is Nothing OrElse ctx Is Nothing Then Return False
            If ctx.ResolvedStyleObj Is Nothing Then
                lab.Log("STYLE", "FINAL_FAIL", "name=ALL requested=" & ctx.RequestedStyleName & " final=no_style_object")
                Return False
            End If
            TryApplyStyleToDimensionsCollectionLab(TryCast(sh.Dimensions, Dimensions), ctx.RequestedStyleName, ctx.ResolvedStyleObj, lab, "final_pass_collection")
            Dim n As Integer = SafeDimensionsCount(sh)
            If n <= 0 Then Return True

            Dim allOk As Boolean = True
            Dim firstFinal As String = ""
            Dim lastFinal As String = ""

            For i As Integer = 1 To n
                Dim d As Dimension = Nothing
                Try
                    d = TryCast(sh.Dimensions.Item(i), Dimension)
                Catch
                    Continue For
                End Try
                If d Is Nothing Then Continue For
                Dim label = "SheetDim_" & i.ToString(CultureInfo.InvariantCulture)
                Dim ok = ApplyDimensionStyleStrict(d, ctx.ResolvedStyleObj, ctx.RequestedStyleName, lab, label, ctx.IsCleanFullStrict)
                Dim finalName = ReadDimensionStyleName(d, lab)
                If String.IsNullOrWhiteSpace(firstFinal) Then firstFinal = finalName
                lastFinal = finalName
                If Not ok Then allOk = False
            Next

            If String.IsNullOrWhiteSpace(ctx.SummaryHorizontalStyleFinal) OrElse String.Equals(ctx.SummaryHorizontalStyleFinal, "N/A", StringComparison.OrdinalIgnoreCase) Then
                ctx.SummaryHorizontalStyleFinal = If(String.IsNullOrWhiteSpace(firstFinal), "FAIL", firstFinal)
            End If
            If String.IsNullOrWhiteSpace(ctx.SummaryVerticalStyleFinal) OrElse String.Equals(ctx.SummaryVerticalStyleFinal, "N/A", StringComparison.OrdinalIgnoreCase) Then
                ctx.SummaryVerticalStyleFinal = If(String.IsNullOrWhiteSpace(lastFinal), ctx.SummaryHorizontalStyleFinal, lastFinal)
            End If
            ctx.SummaryStyleResult = If(allOk, "SUCCESS", "FAIL")
            Return allOk
        End Function

        Private Shared Sub LogAvailableStylesNearEnd(draft As DraftDocument, sheet As Sheet, lab As DimLabLogger)
            Dim collectionNames = New String() {"DimensionStyles", "DimStyles", "LinearStyles", "Styles"}
            If draft IsNot Nothing Then
                For Each coll In collectionNames
                    LogStyleCollection(draft, "draft_near_end", coll, lab)
                Next
            End If
            If sheet IsNot Nothing Then
                For Each coll In collectionNames
                    LogStyleCollection(sheet, "sheet_near_end", coll, lab)
                Next
            End If
        End Sub

        Private Shared Sub LogStyleCollection(owner As Object, ownerLabel As String, collectionName As String, lab As DimLabLogger)
            If owner Is Nothing Then Return
            Dim styles As Object = Nothing
            Dim count As Integer = -1
            Try
                styles = CallByName(owner, collectionName, CallType.Get)
                If styles Is Nothing Then
                    lab.Log("STYLE_SCAN_COLLECTION", "OPEN", "name=" & collectionName & " ok=False count=0 source=" & ownerLabel)
                    Return
                End If
                count = CInt(CallByName(styles, "Count", CallType.Get))
                lab.Log("STYLE_SCAN_COLLECTION", "OPEN", "name=" & collectionName & " ok=True count=" & count.ToString(CultureInfo.InvariantCulture) & " source=" & ownerLabel)
            Catch ex As Exception
                lab.Log("STYLE_SCAN_COLLECTION", "OPEN", "name=" & collectionName & " ok=False source=" & ownerLabel & " error=" & ex.Message)
                Return
            End Try

            For i As Integer = 1 To count
                Try
                    Dim it = CallByName(styles, "Item", CallType.Method, i)
                    Dim nm As String = ""
                    Try
                        nm = Convert.ToString(CallByName(it, "Name", CallType.Get), CultureInfo.InvariantCulture)
                    Catch
                        nm = ""
                    End Try
                    lab.Log("STYLE_SCAN", "ITEM", "collection=" & collectionName & " idx=" & i.ToString(CultureInfo.InvariantCulture) & " name=" & nm & " source=" & ownerLabel)
                Catch
                End Try
            Next
        End Sub

        Private Shared Sub LogDimensionStyleDiagnostics(dimObj As Object, dimLabel As String, lab As DimLabLogger)
            If dimObj Is Nothing Then Return
            Try
                Dim stObj = CallByName(dimObj, "Style", CallType.Get)
                Dim t As String = ""
                Try : t = stObj.GetType().FullName : Catch : t = "" : End Try
                lab.Log("STYLE_DIAG", "DIM_STYLE_OBJECT", "name=" & dimLabel & " type=" & t & " typename=" & TypeName(stObj))
                Dim stName As String = ""
                Try : stName = Convert.ToString(CallByName(stObj, "Name", CallType.Get), CultureInfo.InvariantCulture) : Catch : stName = "" : End Try
                lab.Log("STYLE_DIAG", "DIM_STYLE_NAME", "name=" & dimLabel & " style=" & stName)
            Catch ex As Exception
                lab.Log("STYLE_DIAG", "DIM_STYLE_OBJECT", "name=" & dimLabel & " type=unreadable error=" & ex.Message)
            End Try

            For Each p In New String() {"Style", "StyleName", "DimensionStyle", "DimensionStyleName", "DimStyle", "DimStyleName"}
                Dim ok As Boolean = False
                Dim val As String = ""
                Try
                    Dim v = CallByName(dimObj, p, CallType.Get)
                    ok = True
                    val = Convert.ToString(v, CultureInfo.InvariantCulture)
                Catch ex As Exception
                    ok = False
                    val = ex.Message
                End Try
                lab.Log("STYLE_DIAG", "PROPERTY_EXISTS", "property=" & p & " ok=" & ok.ToString(CultureInfo.InvariantCulture) & " value=" & val)
            Next
        End Sub

        Private Shared Function ReadDimensionStyleName(dimObj As Object, lab As DimLabLogger) As String
            If dimObj Is Nothing Then Return ""
            Dim nm As String = ""

            lab.Log("STYLE", "READ_TRY", "route=Style.Name")
            Try
                Dim st = CallByName(dimObj, "Style", CallType.Get)
                nm = Convert.ToString(CallByName(st, "Name", CallType.Get), CultureInfo.InvariantCulture)
                lab.Log("STYLE", "READ_OK", "route=Style.Name name=" & nm)
                Return nm
            Catch ex As Exception
                lab.Log("STYLE", "READ_FAIL", "route=Style.Name error=" & ex.Message)
            End Try

            lab.Log("STYLE", "READ_TRY", "route=DimensionStyle.Name")
            Try
                Dim st = CallByName(dimObj, "DimensionStyle", CallType.Get)
                nm = Convert.ToString(CallByName(st, "Name", CallType.Get), CultureInfo.InvariantCulture)
                lab.Log("STYLE", "READ_OK", "route=DimensionStyle.Name name=" & nm)
                Return nm
            Catch ex As Exception
                lab.Log("STYLE", "READ_FAIL", "route=DimensionStyle.Name error=" & ex.Message)
            End Try

            lab.Log("STYLE", "READ_TRY", "route=StyleName")
            Try
                nm = Convert.ToString(CallByName(dimObj, "StyleName", CallType.Get), CultureInfo.InvariantCulture)
                lab.Log("STYLE", "READ_OK", "route=StyleName name=" & nm)
                Return nm
            Catch ex As Exception
                lab.Log("STYLE", "READ_FAIL", "route=StyleName error=" & ex.Message)
            End Try

            lab.Log("STYLE", "READ_TRY", "route=Style")
            Try
                nm = Convert.ToString(CallByName(dimObj, "Style", CallType.Get), CultureInfo.InvariantCulture)
                lab.Log("STYLE", "READ_OK", "route=Style name=" & nm)
                Return nm
            Catch ex As Exception
                lab.Log("STYLE", "READ_FAIL", "route=Style error=" & ex.Message)
            End Try

            Return nm
        End Function

        Private Shared Function NormalizeStyleNameForCompare(s As String) As String
            If String.IsNullOrWhiteSpace(s) Then Return ""
            Dim t = s.Trim().Replace(ChrW(&HA0), " ").Replace(".", ",")
            t = t.Replace(" ", "")
            Return t
        End Function

        Private Shared Function IsStyleNameU35(name As String) As Boolean
            If String.IsNullOrWhiteSpace(name) Then Return False
            Dim n = NormalizeStyleNameForCompare(name)
            Return String.Equals(n, "U3,5", StringComparison.OrdinalIgnoreCase)
        End Function

        Private Shared Function TryCenterTextInsideDimension(d As Dimension, draft As DraftDocument, app As Application, lab As DimLabLogger, name As String) As String
            If d Is Nothing Then Return "SKIP"
            Dim centered As Boolean = False
            Dim profileApplied As String = "none"
            For Each mn In New String() {"CenterText", "CenterDimensionText", "Center"}
                Try
                    CallByName(d, mn, CallType.Method)
                    centered = True
                    profileApplied = "center_method:" & mn
                    Exit For
                Catch
                End Try
            Next
            Try
                Dim td = SafeGetTrackDistance(d)
                If Not Double.IsNaN(td) AndAlso td > 0.006R Then
                    Dim tdIn = 0.006R
                    d.TrackDistance = tdIn
                    Try : CallByName(d, "AbsoluteTrackDistance", CallType.Let, tdIn) : Catch : End Try
                    centered = True
                End If
            Catch
            End Try
            Try
                Dim offsets = CallByName(d, "GetTextOffsets", CallType.Method)
                lab.Log("TEXT_CENTER", "OFFSETS_READ", "name=" & name & " value=" & Convert.ToString(offsets, CultureInfo.InvariantCulture))
            Catch
            End Try
            Try
                CallByName(d, "SetTextOffsets", CallType.Method, 0.0R, 0.0R)
                centered = True
                If profileApplied = "none" Then profileApplied = "offsets_zero"
            Catch
            End Try
            ' Barrido de keypoints y posiciones objetivo para identificar qué handle controla el texto.
            Try
                Dim rx1 As Double = 0.0R, ry1 As Double = 0.0R, rx2 As Double = 0.0R, ry2 As Double = 0.0R
                If TryReadDimensionRange(d, rx1, ry1, rx2, ry2) Then
                    Dim minX0 As Double = Math.Min(rx1, rx2)
                    Dim maxX0 As Double = Math.Max(rx1, rx2)
                    Dim minY0 As Double = Math.Min(ry1, ry2)
                    Dim maxY0 As Double = Math.Max(ry1, ry2)
                    Dim cx As Double = (minX0 + maxX0) * 0.5R
                    Dim cy As Double = (minY0 + maxY0) * 0.5R
                    Dim w0 As Double = Math.Abs(maxX0 - minX0)
                    Dim h0 As Double = Math.Abs(maxY0 - minY0)
                    Dim off As Double = Math.Max(0.0015R, Math.Min(0.006R, Math.Max(w0, h0) * 0.12R))
                    Dim nKp As Integer = 0
                    Try : nKp = CInt(CallByName(d, "KeyPointCount", CallType.Get)) : Catch : nKp = 0 : End Try
                    If nKp > 0 Then
                        lab.Log("TEXT_CENTER", "SWEEP_START", "name=" & name & " keyPoints=" & nKp.ToString(CultureInfo.InvariantCulture) &
                                " baseCenter=(" & Gd(cx) & "," & Gd(cy) & ") offset=" & Gd(off))

                        Dim targets = New List(Of Tuple(Of String, Double, Double)) From {
                            New Tuple(Of String, Double, Double)("CENTER", cx, cy)
                        }

                        Dim bestScore As Double = Double.MaxValue
                        Dim bestIdx As Integer = -1
                        Dim bestTx As Double = cx
                        Dim bestTy As Double = cy
                        Dim bestTag As String = "CENTER"
                        Dim maxKp As Integer = Math.Min(nKp - 1, 8)

                        For kpIdx As Integer = 0 To maxKp
                            Dim px As Double = 0.0R, py As Double = 0.0R, pz As Double = 0.0R
                            Dim kpt As SolidEdgeConstants.KeyPointType
                            Dim hdl As SolidEdgeConstants.HandleType
                            Try
                                d.GetKeyPoint(kpIdx, px, py, pz, kpt, hdl)
                                lab.Log("TEXT_CENTER", "KP", "name=" & name & " idx=" & kpIdx.ToString(CultureInfo.InvariantCulture) &
                                        " p=(" & Gd(px) & "," & Gd(py) & ") kpt=" & CInt(kpt).ToString(CultureInfo.InvariantCulture) &
                                        " hdl=" & CInt(hdl).ToString(CultureInfo.InvariantCulture))
                            Catch
                            End Try

                            For Each t In targets
                                Try
                                    d.SetKeyPoint(kpIdx, t.Item2, t.Item3, 0.0R)
                                    Try
                                        If draft IsNot Nothing Then
                                            draft.UpdateAll(True)
                                        End If
                                    Catch
                                    End Try
                                    Try
                                        If app IsNot Nothing Then
                                            app.DoIdle()
                                        End If
                                    Catch
                                    End Try

                                    Dim ax1 As Double = 0.0R, ay1 As Double = 0.0R, ax2 As Double = 0.0R, ay2 As Double = 0.0R
                                    Dim score As Double = Double.MaxValue
                                    If TryReadDimensionRange(d, ax1, ay1, ax2, ay2) Then
                                        Dim acx As Double = (Math.Min(ax1, ax2) + Math.Max(ax1, ax2)) * 0.5R
                                        Dim acy As Double = (Math.Min(ay1, ay2) + Math.Max(ay1, ay2)) * 0.5R
                                        score = Math.Sqrt((acx - cx) * (acx - cx) + (acy - cy) * (acy - cy))
                                        lab.Log("TEXT_CENTER", "SWEEP_TRY", "name=" & name &
                                                " kp=" & kpIdx.ToString(CultureInfo.InvariantCulture) &
                                                " target=" & t.Item1 &
                                                " targetP=(" & Gd(t.Item2) & "," & Gd(t.Item3) & ")" &
                                                " centerAfter=(" & Gd(acx) & "," & Gd(acy) & ")" &
                                                " score=" & Gd(score))
                                    Else
                                        lab.Log("TEXT_CENTER", "SWEEP_TRY", "name=" & name &
                                                " kp=" & kpIdx.ToString(CultureInfo.InvariantCulture) &
                                                " target=" & t.Item1 & " centerAfter=unreadable score=NaN")
                                    End If

                                    If score < bestScore Then
                                        bestScore = score
                                        bestIdx = kpIdx
                                        bestTx = t.Item2
                                        bestTy = t.Item3
                                        bestTag = t.Item1
                                    End If
                                Catch ex As Exception
                                    lab.Log("TEXT_CENTER", "SWEEP_FAIL", "name=" & name &
                                            " kp=" & kpIdx.ToString(CultureInfo.InvariantCulture) &
                                            " target=" & t.Item1 &
                                            " error=" & ex.Message)
                                End Try
                            Next
                        Next

                        If bestIdx >= 0 Then
                            Try
                                d.SetKeyPoint(bestIdx, bestTx, bestTy, 0.0R)
                                centered = True
                                profileApplied = If(profileApplied = "none",
                                                    "setkeypoint_sweep_best_" & bestTag,
                                                    profileApplied & "+setkeypoint_sweep_best_" & bestTag)
                                lab.Log("TEXT_CENTER", "SWEEP_BEST", "name=" & name &
                                        " kp=" & bestIdx.ToString(CultureInfo.InvariantCulture) &
                                        " target=" & bestTag &
                                        " p=(" & Gd(bestTx) & "," & Gd(bestTy) & ")" &
                                        " score=" & Gd(bestScore))
                            Catch ex As Exception
                                lab.Log("TEXT_CENTER", "SWEEP_BEST_FAIL", "name=" & name & " error=" & ex.Message)
                            End Try
                        End If
                    End If
                End If
            Catch
            End Try
            ' Intentos directos sobre propiedades del objeto Dimension.
            Try
                CallByName(d, "CoordinateTextPosition", CallType.Let, 0) ' solicitado por usuario
                centered = True
                profileApplied = If(profileApplied = "none", "coord_text_above_dim_0", profileApplied & "+coord_text_above_dim_0")
                lab.Log("TEXT_CENTER", "COORD_TEXT_POSITION", "name=" & name & " value=0 target=Dimension")
            Catch
                Try
                    CallByName(d, "CoordinateTextPosition", CallType.Let, 1) ' fallback típico
                    centered = True
                    profileApplied = If(profileApplied = "none", "coord_text_above_dim_1", profileApplied & "+coord_text_above_dim_1")
                    lab.Log("TEXT_CENTER", "COORD_TEXT_POSITION", "name=" & name & " value=1(igDimStyleCoordTextAbove) target=Dimension")
                Catch
                End Try
            Catch
            End Try
            Try
                CallByName(d, "CoordinateTextOrientation", CallType.Let, 1)
                centered = True
                profileApplied = If(profileApplied = "none", "coord_text_orientation_dim", profileApplied & "+coord_text_orientation_dim")
                lab.Log("TEXT_CENTER", "COORD_TEXT_ORIENTATION", "name=" & name & " value=1 target=Dimension")
            Catch
            End Try
            ' Perfil "internal centered": evitar texto "pulled out" + posición/orientación internas.
            Try
                CallByName(d, "OverridePulledOutText", CallType.Let, False)
                centered = True
                If profileApplied = "none" Then profileApplied = "override_pulled_out_false"
            Catch
            End Try
            Try
                CallByName(d, "OverridePulledOutText2", CallType.Let, False)
                centered = True
            Catch
            End Try
            Try
                CallByName(d, "TextPosition", CallType.Let, 1) ' igDimStyleTextAbove
                centered = True
                lab.Log("TEXT_CENTER", "TEXT_POSITION", "name=" & name & " value=1(igDimStyleTextAbove)")
                profileApplied = If(profileApplied = "none", "text_position_above_1", profileApplied & "+text_position_above_1")
            Catch
                Try
                    CallByName(d, "TextPosition", CallType.Let, 0) ' fallback
                    centered = True
                    lab.Log("TEXT_CENTER", "TEXT_POSITION", "name=" & name & " value=0")
                    profileApplied = If(profileApplied = "none", "text_position_above_0", profileApplied & "+text_position_above_0")
                Catch
                End Try
            End Try
            Try
                If name.IndexOf("Vertical", StringComparison.OrdinalIgnoreCase) >= 0 Then
                    CallByName(d, "TextOrientation", CallType.Let, 2) ' vertical solicitado
                    lab.Log("TEXT_CENTER", "TEXT_ORIENTATION", "name=" & name & " value=2(igDimStyleTextVertical)")
                    profileApplied = If(profileApplied = "none", "text_orientation_vertical_2", profileApplied & "+text_orientation_vertical_2")
                Else
                    CallByName(d, "TextOrientation", CallType.Let, 1) ' horizontal para cota horizontal
                    lab.Log("TEXT_CENTER", "TEXT_ORIENTATION", "name=" & name & " value=1(igDimStyleTextHorizontal)")
                    profileApplied = If(profileApplied = "none", "text_orientation_horizontal_1", profileApplied & "+text_orientation_horizontal_1")
                End If
                centered = True
            Catch
                Try
                    CallByName(d, "TextOrientation", CallType.Let, 0)
                    centered = True
                    lab.Log("TEXT_CENTER", "TEXT_ORIENTATION", "name=" & name & " value=0(fallback)")
                Catch
                End Try
            End Try
            ' Intentos sobre el DimStyle concreto de la cota (override por objeto).
            Try
                Dim stObj = CallByName(d, "Style", CallType.Get)
                If stObj IsNot Nothing Then
                    Try
                        CallByName(stObj, "TextPosition", CallType.Let, 1)
                        centered = True
                        profileApplied = If(profileApplied = "none", "style_text_position_above_1", profileApplied & "+style_text_position_above_1")
                        lab.Log("TEXT_CENTER", "STYLE_TEXT_POSITION", "name=" & name & " value=1(igDimStyleTextAbove)")
                    Catch
                        Try
                            CallByName(stObj, "TextPosition", CallType.Let, 0)
                            centered = True
                            profileApplied = If(profileApplied = "none", "style_text_position_above_0", profileApplied & "+style_text_position_above_0")
                            lab.Log("TEXT_CENTER", "STYLE_TEXT_POSITION", "name=" & name & " value=0")
                        Catch
                        End Try
                    Catch
                    End Try
                    Try
                        If name.IndexOf("Vertical", StringComparison.OrdinalIgnoreCase) >= 0 Then
                            CallByName(stObj, "TextOrientation", CallType.Let, 2)
                            profileApplied = If(profileApplied = "none", "style_text_orientation_vertical_2", profileApplied & "+style_text_orientation_vertical_2")
                            lab.Log("TEXT_CENTER", "STYLE_TEXT_ORIENTATION", "name=" & name & " value=2(igDimStyleTextVertical)")
                        Else
                            CallByName(stObj, "TextOrientation", CallType.Let, 1)
                            profileApplied = If(profileApplied = "none", "style_text_orientation_horizontal_1", profileApplied & "+style_text_orientation_horizontal_1")
                            lab.Log("TEXT_CENTER", "STYLE_TEXT_ORIENTATION", "name=" & name & " value=1(igDimStyleTextHorizontal)")
                        End If
                        centered = True
                    Catch
                    End Try
                    Try
                        CallByName(stObj, "CoordinateTextPosition", CallType.Let, 0)
                        centered = True
                        profileApplied = If(profileApplied = "none", "style_coord_text_above_0", profileApplied & "+style_coord_text_above_0")
                        lab.Log("TEXT_CENTER", "STYLE_COORD_TEXT_POSITION", "name=" & name & " value=0")
                    Catch
                        Try
                            CallByName(stObj, "CoordinateTextPosition", CallType.Let, 1)
                            centered = True
                            profileApplied = If(profileApplied = "none", "style_coord_text_above_1", profileApplied & "+style_coord_text_above_1")
                            lab.Log("TEXT_CENTER", "STYLE_COORD_TEXT_POSITION", "name=" & name & " value=1(igDimStyleCoordTextAbove)")
                        Catch
                        End Try
                    Catch
                    End Try
                    Try
                        CallByName(stObj, "CoordinateTextOrientation", CallType.Let, 1)
                        centered = True
                        profileApplied = If(profileApplied = "none", "style_coord_text_orientation", profileApplied & "+style_coord_text_orientation")
                        lab.Log("TEXT_CENTER", "STYLE_COORD_TEXT_ORIENTATION", "name=" & name & " value=1")
                    Catch
                    End Try
                    Try
                        CallByName(stObj, "TextClearanceGap", CallType.Let, 0.0R)
                        centered = True
                        profileApplied = If(profileApplied = "none", "style_text_clearance_0", profileApplied & "+style_text_clearance_0")
                        lab.Log("TEXT_CENTER", "STYLE_TEXT_CLEARANCE_GAP", "name=" & name & " value=0")
                    Catch
                    End Try
                    Try
                        CallByName(stObj, "AboveGap", CallType.Let, 0.0R)
                        centered = True
                        profileApplied = If(profileApplied = "none", "style_above_gap_0", profileApplied & "+style_above_gap_0")
                        lab.Log("TEXT_CENTER", "STYLE_ABOVE_GAP", "name=" & name & " value=0")
                    Catch
                    End Try
                End If
            Catch
            End Try
            Try
                If draft IsNot Nothing Then
                    draft.UpdateAll(True)
                End If
            Catch
            End Try
            Try
                If app IsNot Nothing Then
                    app.DoIdle()
                End If
            Catch
            End Try
            If centered Then
                lab.Log("TEXT_CENTER", "RESULT", "name=" & name & " OK profile=" & profileApplied)
                Return "OK"
            Else
                lab.Log("TEXT_CENTER", "SKIP", "reason=no_compilable_api_found name=" & name)
                Return "SKIP"
            End If
        End Function

        Private Shared Function RunVisibleZeroProbe(sh As Sheet,
                                                    draft As DraftDocument,
                                                    app As Application,
                                                    dims As Dimensions,
                                                    created As List(Of Object),
                                                    lab As DimLabLogger,
                                                    ctx As LabRunContext) As Boolean
            ctx.PlainTextProbe = "FAIL"
            ctx.PlainLineProbe = "FAIL"
            ctx.PlainAddLengthCreate = "FAIL"
            ctx.PlainGraphicsMaterialized = False

            Dim shName = SafeSheetName(sh)
            lab.Log("VIS0", "TARGET_SHEET", shName)
            If IsTwoDModelSheetName(shName) Then
                lab.Log("2DMODEL", "BLOCK_VIS0", "refuse_probe_on_2d_model_sheet")
                Return False
            End If

            Dim layerObj = TryEnsureDimLabLayer(draft, sh, lab)

            lab.Log("VIS0", "TEXT_CREATE_TRY", "")
            Dim tb As Object = Nothing
            Try
                Dim tbc = sh.TextBoxes
                Try : tb = tbc.Add(0.02R, 0.265R, "DIMLAB_MARKER_VISIBLE") : Catch : End Try
                If tb Is Nothing Then
                    Try : tb = CallByName(tbc, "Add", CallType.Method, 0.02R, 0.265R, 0.0R, "DIMLAB_MARKER_VISIBLE") : Catch : End Try
                End If
            Catch ex As Exception
                lab.Log("VIS0", "TEXT_CREATE_FAIL", ex.Message)
            End Try
            If tb Is Nothing Then
                lab.Log("VIS0", "TEXT_CREATE", "FAIL")
            Else
                created.Add(tb)
                ctx.Vis0TextObj = tb
                TrySetObjectLayer(tb, layerObj, lab, "TextBox")
                ctx.PlainTextProbe = "OK"
                lab.Log("VIS0", "TEXT_CREATE", "OK")
            End If

            lab.Log("VIS0", "LINE_CREATE_TRY", "")
            Dim ln As Object = Nothing
            Try
                ln = sh.Lines2d.AddBy2Points(0.025R, 0.265R, 0.085R, 0.265R)
            Catch ex As Exception
                lab.Log("VIS0", "LINE_CREATE_FAIL", ex.Message)
            End Try
            If ln Is Nothing Then
                lab.Log("VIS0", "LINE_CREATE", "FAIL")
            Else
                created.Add(ln)
                ctx.Vis0LineObj = ln
                TrySetObjectLayer(ln, layerObj, lab, "Line2d")
                ctx.PlainLineProbe = "OK"
                lab.Log("VIS0", "LINE_CREATE_OK", "sheet=" & shName & " start=(0.025,0.265) end=(0.085,0.265)")
            End If

            Dim cntBefore = SafeDimensionsCount(sh)
            lab.Log("VIS0", "DIMCOUNT", "before=" & cntBefore.ToString(CultureInfo.InvariantCulture))

            lab.Log("VIS0", "ADDLENGTH_TRY", "")
            Dim dPlain As Dimension = Nothing
            Try
                Dim refObj As Object = ln
                If ln IsNot Nothing Then
                    Try : refObj = CallByName(ln, "Reference", CallType.Get) : Catch : End Try
                End If
                dPlain = TryCast(dims.AddLength(refObj), Dimension)
            Catch ex As Exception
                lab.Log("VIS0", "ADDLENGTH_FAIL", ex.Message)
            End Try
            Dim cntAfter = SafeDimensionsCount(sh)
            lab.Log("VIS0", "DIMCOUNT", "after=" & cntAfter.ToString(CultureInfo.InvariantCulture))

            If dPlain Is Nothing Then
                lab.Log("VIS0", "ADDLENGTH", "FAIL value=n/a")
                Return False
            End If
            created.Add(dPlain)
            ctx.Vis0PlainDim = dPlain
            ctx.PlainAddLengthCreate = "OK"

            Try
                Dim stObj = draft.DimensionStyles.Item("U3,5")
                If stObj IsNot Nothing Then dPlain.DimensionStyle = stObj
            Catch exSt As Exception
                lab.Log("VIS0", "STYLE_U35_SKIP", exSt.Message)
            End Try

            lab.Log("VIS0", "TRACK_SET", "value=0.008")
            Try
                dPlain.TrackDistance = 0.008R
                lab.Log("VIS0", "TRACK_SET", "OK")
            Catch ex As Exception
                lab.Log("VIS0", "TRACK_SET", "FAIL " & ex.Message)
            End Try

            Try
                If draft IsNot Nothing Then draft.UpdateAll(True)
                If app IsNot Nothing Then app.DoIdle()
            Catch
            End Try

            Dim rv = ReadDimValue(dPlain)
            lab.Log("VIS0", "ADDLENGTH", "OK value=" & rv.ToString("G17", CultureInfo.InvariantCulture))

            Dim rx1 As Double, ry1 As Double, rx2 As Double, ry2 As Double
            If TryReadDimensionRange(dPlain, rx1, ry1, rx2, ry2) Then
                lab.Log("VIS0", "RANGE", FormatRangeStr(rx1, ry1, rx2, ry2))
            Else
                lab.Log("VIS0", "RANGE", "unreadable")
            End If

            Dim lc As Integer, ac As Integer, tc As Integer, pc As Integer
            If TryReadDisplayDataCounts(dPlain, lc, ac, tc, pc) Then
                lab.Log("VIS0", "DISPLAYDATA", "lineCount=" & lc.ToString(CultureInfo.InvariantCulture) &
                        " arcCount=" & ac.ToString(CultureInfo.InvariantCulture) &
                        " textCount=" & tc.ToString(CultureInfo.InvariantCulture) &
                        " pointCount=" & pc.ToString(CultureInfo.InvariantCulture))
            Else
                lab.Log("VIS0", "DISPLAYDATA", "FAIL")
            End If

            LogDimDisplayDataInspect(dPlain, lab, "VIS0_PlainAddLength")

            Dim vf = ValidateVisibleFinal(dPlain, sh, draft, app, "VIS0_PlainAddLength", lab)
            ctx.PlainGraphicsMaterialized = vf.Materialized
            If Not vf.Materialized Then
                Return False
            End If
            Return True
        End Function

        Private Shared Sub TrySelectAddForensic(draft As DraftDocument, ctx As LabRunContext, lab As DimLabLogger, phase As String)
            Dim ss As Object = Nothing
            Try : ss = draft.SelectSet : Catch : End Try
            If ss Is Nothing Then
                lab.Log("SELECT", "FAIL", phase & " no_SelectSet")
                Return
            End If
            Try : CallByName(ss, "RemoveAll", CallType.Method) : Catch : End Try

            Dim addFn = Sub(tag As String, o As Object)
                            Try
                                If o Is Nothing Then
                                    lab.Log("SELECT", "ADD", "phase=" & phase & " object=" & tag & " FAIL null")
                                    Return
                                End If
                                CallByName(ss, "Add", CallType.Method, o)
                                lab.Log("SELECT", "ADD", "phase=" & phase & " object=" & tag & " OK")
                            Catch ex As Exception
                                lab.Log("SELECT", "ADD", "phase=" & phase & " object=" & tag & " FAIL " & ex.Message)
                            End Try
                        End Sub

            addFn("TextBox", ctx.Vis0TextObj)
            addFn("Line2d", ctx.Vis0LineObj)
            addFn("PlainAddLengthDimension", ctx.Vis0PlainDim)
            addFn("DVRefDimension", ctx.DvRefHorizontalDim)

            Dim c = SafeSsCount(ss)
            lab.Log("SELECT", "COUNT", "phase=" & phase & " count=" & c.ToString(CultureInfo.InvariantCulture))
        End Sub

        Private Shared Function ClassifyWrongDimensionValue(actual As Double, bounds As DVBounds, test As String, lab As DimLabLogger) As String
            If Double.IsNaN(actual) OrElse Double.IsInfinity(actual) Then
                lab.Log("VALUE_CLASSIFY", "INFO", "test=" & test & " actual=NaN class=UNKNOWN")
                Return "UNKNOWN"
            End If
            Dim tolW = Math.Max(0.001R, bounds.ExpectedWidth * 0.005R)
            Dim tolH = Math.Max(0.001R, bounds.ExpectedHeight * 0.005R)
            Dim tS = 0.003R
            Dim cls As String = "UNKNOWN_MISMATCH"
            If Math.Abs(actual - bounds.ExpectedWidth) <= tolW Then
                cls = "WIDTH_INSTEAD_OF_HEIGHT"
            ElseIf Math.Abs(actual - bounds.ExpectedHeight) <= tolH Then
                cls = "OK"
            ElseIf Math.Abs(actual - 0.102R) <= tS Then
                cls = "PARTIAL_VERTICAL_SIDE"
            ElseIf Math.Abs(actual - 0.007R) <= tS Then
                cls = "LOCAL_GAP"
            ElseIf Math.Abs(actual - 0.17R) <= tS Then
                cls = "HALF_WIDTH_VISIBLE_OR_VIEW_SCALE"
            ElseIf Math.Abs(actual - 0.34R) <= tS Then
                cls = "INNER_WIDTH"
            ElseIf Math.Abs(actual - 0.354R) <= tS Then
                cls = "FULL_WIDTH_INSTEAD_OF_HEIGHT"
            End If
            lab.Log("VALUE_CLASSIFY", "INFO", "test=" & test & " actual=" & Gd(actual) & " class=" & cls)
            Return cls
        End Function

        Private Shared Function LabVisibleOk(d As Dimension, sh As Sheet, draft As DraftDocument, app As Application, lab As DimLabLogger, tag As String) As Boolean
            Dim vf = ValidateVisibleFinal(d, sh, draft, app, tag, lab)
            If vf.Visible Then Return True
            Dim lc As Integer, ac As Integer, tc As Integer, pc As Integer
            If TryReadDisplayDataCounts(d, lc, ac, tc, pc) AndAlso lc + tc > 0 Then Return True
            Return vf.Materialized
        End Function

        Private Shared Function EvaluateDimensionForLab(d As Dimension,
                                                        expected As Double,
                                                        draft As DraftDocument,
                                                        dv As DrawingView,
                                                        sh As Sheet,
                                                        app As Application,
                                                        lab As DimLabLogger,
                                                        name As String,
                                                        ctx As LabRunContext) As LabEvalResult
            Dim actual = ReadDimValue(d)
            Dim tolOk = Math.Max(0.001R, Math.Abs(expected) * 0.005R)
            Dim valueOk = Math.Abs(actual - expected) <= tolOk
            Dim vf = ValidateVisibleFinal(d, sh, draft, app, name, lab)
            Dim visibleOk = vf.Visible
            Dim materialized = vf.Materialized
            Dim connState = CheckConnectedBySelectSet(draft, dv, d, lab, name, ctx)

            Dim result As String
            If valueOk AndAlso visibleOk AndAlso materialized Then
                If connState = ConnectivityState.TrueState Then
                    result = "SUCCESS_CONNECTED"
                Else
                    result = "SUCCESS_VALUE_VISIBLE_CONNECTION_UNCERTAIN"
                End If
            ElseIf valueOk AndAlso visibleOk Then
                result = "SUCCESS_VALUE_VISIBLE_BUT_DISPLAYDATA_UNCERTAIN"
            ElseIf Not valueOk Then
                result = "WRONG_VALUE"
            ElseIf Not visibleOk Then
                result = "NOT_VISIBLE"
            Else
                result = "FAIL"
            End If

            lab.Log("EVAL", "STATE", "name=" & name &
                    " valueOk=" & valueOk.ToString(CultureInfo.InvariantCulture) &
                    " visibleOk=" & visibleOk.ToString(CultureInfo.InvariantCulture) &
                    " materialized=" & materialized.ToString(CultureInfo.InvariantCulture) &
                    " connected=" & connState.ToString())
            lab.Log("EVAL", "RESULT", "name=" & name & " result=" & result)

            Return New LabEvalResult With {
                .ValueOk = valueOk,
                .VisibleOk = visibleOk,
                .Materialized = materialized,
                .Connected = connState,
                .Result = result
            }
        End Function

        Private Shared Sub LogHorizontalDvRefOutcome(d As Dimension, resRaw As String, expected As Double, dvBounds As DVBounds,
                                                     draft As DraftDocument, dv As DrawingView, lab As DimLabLogger, ctx As LabRunContext)
            Dim sh As Sheet = ctx.Sheet
            Dim actual = ReadDimValue(d)
            lab.Log("HORIZONTAL", "OK", "value=" & Gd(actual))
            lab.Log("HORIZONTAL", "DELTA", "expected=" & Gd(expected) & " actual=" & Gd(actual) & " delta=" & Gd(Math.Abs(actual - expected)))
            Dim ev = EvaluateDimensionForLab(d, expected, draft, dv, sh, ctx.App, lab, "HorizontalTotal_lab", ctx)
            If ev.Result = "WRONG_VALUE" Then
                ClassifyWrongDimensionValue(actual, dvBounds, "horizontal", lab)
            End If
            lab.Log("HORIZONTAL", "RESULT", ev.Result)
        End Sub

        Private Shared Function IsHorizontalLabSuccess(d As Dimension, resRaw As String, expected As Double, draft As DraftDocument, dv As DrawingView, lab As DimLabLogger, ctx As LabRunContext) As Boolean
            If d Is Nothing Then Return False
            Dim ev = EvaluateDimensionForLab(d, expected, draft, dv, ctx.Sheet, ctx.App, lab, "HorizontalLab_final", ctx)
            Return ev.Result.StartsWith("SUCCESS", StringComparison.OrdinalIgnoreCase)
        End Function

        Private Shared Function TryHorizontalExclusiveStep(dims As Dimensions, dv As DrawingView, draft As DraftDocument,
                                                         dvBounds As DVBounds,
                                                         p1 As DVEndpointCandidate, p2 As DVEndpointCandidate,
                                                         testName As String,
                                                         created As List(Of Object), lab As DimLabLogger, ctx As LabRunContext,
                                                         ByRef outD As Dimension) As String
            outD = Nothing
            Dim r = TryOneDistance(dims, dv, draft, testName, p1, p2, dvBounds.ExpectedWidth, True, True, created, lab, ctx, quietLog:=True, outCreatedDimension:=outD)
            If outD IsNot Nothing Then LogHorizontalDvRefOutcome(outD, r, dvBounds.ExpectedWidth, dvBounds, draft, dv, lab, ctx)
            Return r
        End Function

        Private Shared Function RunHorizontalExclusiveForLab(dims As Dimensions, dv As DrawingView, draft As DraftDocument,
                                                             dvBounds As DVBounds,
                                                             boundsLb As LineBounds,
                                                             leftL As List(Of DVEndpointCandidate),
                                                             rightL As List(Of DVEndpointCandidate),
                                                             created As List(Of Object),
                                                             lab As DimLabLogger,
                                                             ctx As LabRunContext) As String
            lab.Log("HORIZONTAL", "START", "placeMode=auto")
            ctx.SummaryHorizCreate = "FAIL"
            If leftL.Count = 0 OrElse rightL.Count = 0 Then
                lab.Log("HORIZONTAL", "RESULT", "FAIL no_endpoints")
                Return "FAIL"
            End If

            Dim outD As Dimension = Nothing
            Dim p1 = PickClosestY(leftL, boundsLb.MaxY)
            Dim p2 = PickClosestY(rightL, p1.Y)
            lab.Log("HORIZONTAL", "PAIR", "leftLine=" & p1.LineIndex.ToString(CultureInfo.InvariantCulture) &
                    " pLeft=(" & Gd(p1.X) & "," & Gd(p1.Y) & ") rightLine=" & p2.LineIndex.ToString(CultureInfo.InvariantCulture) &
                    " pRight=(" & Gd(p2.X) & "," & Gd(p2.Y) & ") sameYDelta=" & Gd(Math.Abs(p1.Y - p2.Y)))
            lab.Log("HORIZONTAL", "TRY", "expected=" & Gd(dvBounds.ExpectedWidth))
            Dim rTop = TryHorizontalExclusiveStep(dims, dv, draft, dvBounds, p1, p2, "HorizontalTotal_TopPair", created, lab, ctx, outD)
            If IsHorizontalLabSuccess(outD, rTop, dvBounds.ExpectedWidth, draft, dv, lab, ctx) Then
                lab.Log("HORIZONTAL", "RESULT", "SUCCESS_VALUE_VISIBLE_CONNECTION_UNCERTAIN")
                lab.Log("HORIZONTAL", "FALLBACK_SKIP", "reason=TopPair_success")
                ctx.DvRefHorizontalDim = outD
                ctx.SummaryHorizCreate = "SUCCESS"
                ctx.SummaryHorizValue = Gd(ReadDimValue(outD))
                ctx.SummaryHorizConnected = CheckConnected(draft, dv, outD, lab, "HorizontalTotal_TopPair", ctx)
                ctx.SummaryHorizVisible = LabVisibleOk(outD, ctx.Sheet, draft, ctx.App, lab, "HorizontalTotal_TopPair")
                Return "SUCCESS_VALUE_VISIBLE_CONNECTION_UNCERTAIN"
            End If

            p1 = PickClosestY(leftL, boundsLb.MinY)
            p2 = PickClosestY(rightL, p1.Y)
            lab.Log("HORIZONTAL", "PAIR", "bottomPair leftLine=" & p1.LineIndex.ToString(CultureInfo.InvariantCulture) &
                    " pLeft=(" & Gd(p1.X) & "," & Gd(p1.Y) & ") rightLine=" & p2.LineIndex.ToString(CultureInfo.InvariantCulture) &
                    " pRight=(" & Gd(p2.X) & "," & Gd(p2.Y) & ") sameYDelta=" & Gd(Math.Abs(p1.Y - p2.Y)))
            lab.Log("HORIZONTAL", "TRY", "expected=" & Gd(dvBounds.ExpectedWidth) & " strategy=BottomPair")
            Dim rBot = TryHorizontalExclusiveStep(dims, dv, draft, dvBounds, p1, p2, "HorizontalTotal_BottomPair", created, lab, ctx, outD)
            If IsHorizontalLabSuccess(outD, rBot, dvBounds.ExpectedWidth, draft, dv, lab, ctx) Then
                lab.Log("HORIZONTAL", "RESULT", "SUCCESS_VALUE_VISIBLE_CONNECTION_UNCERTAIN")
                ctx.DvRefHorizontalDim = outD
                ctx.SummaryHorizCreate = "SUCCESS"
                ctx.SummaryHorizValue = Gd(ReadDimValue(outD))
                ctx.SummaryHorizConnected = CheckConnected(draft, dv, outD, lab, "HorizontalTotal_BottomPair", ctx)
                ctx.SummaryHorizVisible = LabVisibleOk(outD, ctx.Sheet, draft, ctx.App, lab, "HorizontalTotal_BottomPair")
                Return "SUCCESS_VALUE_VISIBLE_CONNECTION_UNCERTAIN"
            End If

            Dim bestA As DVEndpointCandidate = Nothing
            Dim bestC As DVEndpointCandidate = Nothing
            Dim bestDy = Double.MaxValue
            For Each a In leftL
                For Each c In rightL
                    Dim dy = Math.Abs(a.Y - c.Y)
                    If dy < bestDy Then bestDy = dy : bestA = a : bestC = c
                Next
            Next
            If bestA Is Nothing OrElse bestC Is Nothing Then
                lab.Log("HORIZONTAL", "RESULT", "FAIL")
                Return "FAIL"
            End If
            lab.Log("HORIZONTAL", "PAIR", "BestSameY leftLine=" & bestA.LineIndex.ToString(CultureInfo.InvariantCulture) &
                    " pLeft=(" & Gd(bestA.X) & "," & Gd(bestA.Y) & ") rightLine=" & bestC.LineIndex.ToString(CultureInfo.InvariantCulture) &
                    " pRight=(" & Gd(bestC.X) & "," & Gd(bestC.Y) & ") sameYDelta=" & Gd(bestDy))
            lab.Log("HORIZONTAL", "TRY", "expected=" & Gd(dvBounds.ExpectedWidth) & " strategy=BestSameY")
            Dim rBest = TryHorizontalExclusiveStep(dims, dv, draft, dvBounds, bestA, bestC, "HorizontalTotal_BestSameY", created, lab, ctx, outD)
            If IsHorizontalLabSuccess(outD, rBest, dvBounds.ExpectedWidth, draft, dv, lab, ctx) Then
                lab.Log("HORIZONTAL", "RESULT", "SUCCESS_VALUE_VISIBLE_CONNECTION_UNCERTAIN")
                ctx.DvRefHorizontalDim = outD
                ctx.SummaryHorizCreate = "SUCCESS"
                ctx.SummaryHorizValue = Gd(ReadDimValue(outD))
                ctx.SummaryHorizConnected = CheckConnected(draft, dv, outD, lab, "HorizontalTotal_BestSameY", ctx)
                ctx.SummaryHorizVisible = LabVisibleOk(outD, ctx.Sheet, draft, ctx.App, lab, "HorizontalTotal_BestSameY")
                Return "SUCCESS_VALUE_VISIBLE_CONNECTION_UNCERTAIN"
            End If
            ctx.SummaryHorizCreate = "FAIL"
            Return "FAIL"
        End Function

        Private Shared Function IsVerticalLineObj(ln As Object, span As Double, lab As DimLabLogger, session As LabLineReadSession, lineIdx As Integer) As Boolean
            Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
            If Not TryReadDVLineEndpointsBest(ln, x1, y1, x2, y2, lab, lineIdx, session, doLog:=False) Then Return False
            Dim dy = Math.Abs(y2 - y1)
            Dim dx = Math.Abs(x2 - x1)
            Dim tol = Math.Max(1.0E-12R, span * 1.0E-9R)
            Return dx <= tol OrElse (dy > 1.0E-12R AndAlso dx / dy < 0.01R)
        End Function

        Private Shared Sub LogVerticalOutcome(d As Dimension, resRaw As String, expected As Double, dvBounds As DVBounds, draft As DraftDocument, dv As DrawingView, strat As String, lab As DimLabLogger, ctx As LabRunContext)
            If d Is Nothing Then Return
            Dim sh As Sheet = ctx.Sheet
            Dim actual = ReadDimValue(d)
            lab.Log("VERTICAL", "OK", "value=" & Gd(actual))
            lab.Log("VERTICAL", "DELTA", "expected=" & Gd(expected) & " actual=" & Gd(actual) & " delta=" & Gd(Math.Abs(actual - expected)))
            Dim ev = EvaluateDimensionForLab(d, expected, draft, dv, sh, ctx.App, lab, strat, ctx)
            Dim verdict = ev.Result
            If verdict = "WRONG_VALUE" Then
                Dim hint = ClassifyWrongDimensionValue(actual, dvBounds, strat, lab)
                lab.Log("VERTICAL", "WRONG_VALUE_HINT", "value=" & Gd(actual) & " likely=" & hint)
            End If
            lab.Log("VERTICAL", "RESULT", verdict)
            Dim tdAfter = SafeGetTrackDistance(d)
            Dim rxa As Double, rya As Double, rxb As Double, ryb As Double
            Dim rangeStr = "unreadable"
            If TryReadDimensionRange(d, rxa, rya, rxb, ryb) Then rangeStr = FormatRangeStr(rxa, rya, rxb, ryb)
            ctx.Vis.Record(strat, verdict, ev.VisibleOk, rangeStr, Gd(tdAfter), GetVisibilityReasonQuiet(d, sh))
        End Sub

        Private Shared Function IsVerticalLabSuccess(d As Dimension, resRaw As String, expected As Double, draft As DraftDocument, dv As DrawingView, lab As DimLabLogger, ctx As LabRunContext) As Boolean
            If d Is Nothing Then Return False
            Dim ev = EvaluateDimensionForLab(d, expected, draft, dv, ctx.Sheet, ctx.App, lab, "VerticalLab_final", ctx)
            Return ev.Result.StartsWith("SUCCESS", StringComparison.OrdinalIgnoreCase)
        End Function

        Private Shared Function TryVerticalLabStep(dims As Dimensions, dv As DrawingView, draft As DraftDocument,
                                                   dvBounds As DVBounds, pBottom As DVEndpointCandidate, pTop As DVEndpointCandidate,
                                                   stratName As String, created As List(Of Object), lab As DimLabLogger, ctx As LabRunContext,
                                                   ByRef outD As Dimension) As String
            outD = Nothing
            lab.Log("VERTICAL", "TRY", "name=" & stratName & " expected=" & Gd(dvBounds.ExpectedHeight))
            lab.Log("VERTICAL", "POINTS", "pBottom=(" & Gd(pBottom.X) & "," & Gd(pBottom.Y) & ") pTop=(" & Gd(pTop.X) & "," & Gd(pTop.Y) & ") sameXDelta=" & Gd(Math.Abs(pBottom.X - pTop.X)))
            Dim r = TryOneDistance(dims, dv, draft, stratName, pBottom, pTop, dvBounds.ExpectedHeight, True, True, created, lab, ctx, quietLog:=True, outCreatedDimension:=outD)
            If outD IsNot Nothing Then LogVerticalOutcome(outD, r, dvBounds.ExpectedHeight, dvBounds, draft, dv, stratName, lab, ctx)
            Return r
        End Function

        Private Shared Function RunVerticalTotalMultiViewLab(dims As Dimensions,
                                                             draft As DraftDocument,
                                                             sh As Sheet,
                                                             mainHorizontalView As DrawingView,
                                                             created As List(Of Object),
                                                             lab As DimLabLogger,
                                                             ctx As LabRunContext) As String
            lab.Log("VERTICAL", "LAB", "START RunVerticalTotalMultiViewLab")
            ctx.SummaryVertCreate = "FAIL"
            ctx.SummaryVerticalReason = "no_vertical_view_success"

            Dim sw As Double, shh As Double
            If Not TrySheetSizeM(sh, sw, shh) Then
                sw = 0.3R
                shh = 0.3R
            End If

            Dim mainCx As Double, mainCy As Double
            If Not GetViewCenterOnSheet(mainHorizontalView, mainCx, mainCy) Then
                mainCx = 0.0R
                mainCy = 0.0R
            End If

            Dim allScans As List(Of VerticalViewScanResult) = Nothing
            Dim tryList = FindBestViewForVerticalTotalDimension(sh, lab, mainHorizontalView, mainCx, mainCy, sw, shh, allScans)
            If tryList.Count = 0 Then
                lab.Log("VERTICAL", "RESULT", "FAIL no_orthogonal_candidates")
                ctx.SummaryVerticalReason = "no_candidates"
                Return "FAIL"
            End If

            Dim bestOverall As String = "FAIL"
            For Each scan In tryList
                ctx.SummaryVerticalViewLabel = "DrawingView " & scan.Idx.ToString(CultureInfo.InvariantCulture)
                ctx.SummaryVerticalExpectedHeight = Gd(scan.DvBounds.ExpectedHeight)
                Dim r = CreateVerticalTotalDimensionFromDVRef_Lab(dims, scan.View, draft, scan, created, lab, ctx)
                bestOverall = MaxResult(bestOverall, r)
                If r.IndexOf("SUCCESS", StringComparison.OrdinalIgnoreCase) >= 0 Then
                    lab.Log("VERTICAL", "STOP_VIEW_LOOP", "reason=valid_vertical_found view=DrawingView " & scan.Idx.ToString(CultureInfo.InvariantCulture))
                    ctx.SummaryVerticalReason = "ok viewIdx=" & scan.Idx.ToString(CultureInfo.InvariantCulture)
                    Return r
                End If
            Next

            ctx.SummaryVerticalReason = "all_views_strategies_failed"
            Return bestOverall
        End Function

        Private Shared Function GetDrawingViewSheetIndex(sh As Sheet, target As DrawingView) As Integer
            If sh Is Nothing OrElse target Is Nothing Then Return -1
            Try
                Dim cnt = CInt(sh.DrawingViews.Count)
                For i As Integer = 1 To cnt
                    Dim dv As DrawingView = Nothing
                    Try
                        dv = CType(sh.DrawingViews.Item(i), DrawingView)
                    Catch
                        Continue For
                    End Try
                    If ReferenceEquals(dv, target) Then Return i
                Next
            Catch
            End Try
            Return -1
        End Function

        Private Shared Function SafeDrawingViewName(dv As DrawingView) As String
            If dv Is Nothing Then Return ""
            Try
                Return Convert.ToString(dv.Name, CultureInfo.InvariantCulture).Trim()
            Catch
                Return ""
            End Try
        End Function

        Private Shared Function GetViewCenterOnSheet(v As DrawingView, ByRef cx As Double, ByRef cy As Double) As Boolean
            cx = 0.0R
            cy = 0.0R
            Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
            If Not TryViewRangeSheet(v, x1, y1, x2, y2) Then Return False
            cx = (x1 + x2) / 2.0R
            cy = (y1 + y2) / 2.0R
            Return True
        End Function

        Private Shared Function DrawingViewLabelLooksFlat(name As String) As Boolean
            If String.IsNullOrWhiteSpace(name) Then Return False
            Return name.IndexOf("flat", StringComparison.OrdinalIgnoreCase) >= 0
        End Function

        Private Shared Function HorizontalSegmentsOverlapX(ax1 As Double, ax2 As Double, bx1 As Double, bx2 As Double) As Boolean
            Dim aMin = Math.Min(ax1, ax2), aMax = Math.Max(ax1, ax2)
            Dim bMin = Math.Min(bx1, bx2), bMax = Math.Max(bx1, bx2)
            Return Not (aMax < bMin OrElse aMin > bMax)
        End Function

        Private Shared Function EnumerateAndScoreVerticalViewCandidates(sh As Sheet,
                                                                       lab As DimLabLogger,
                                                                       mainHorizontalView As DrawingView,
                                                                       mainCx As Double,
                                                                       mainCy As Double,
                                                                       sheetW As Double,
                                                                       sheetH As Double) As List(Of VerticalViewScanResult)
            Dim all As New List(Of VerticalViewScanResult)()
            Dim cnt As Integer = 0
            Try : cnt = CInt(sh.DrawingViews.Count) : Catch : Return all : End Try
            For i As Integer = 1 To cnt
                Dim dv As DrawingView = Nothing
                Try
                    dv = CType(sh.DrawingViews.Item(i), DrawingView)
                Catch
                    Continue For
                End Try
                If dv Is Nothing OrElse IsIsoView(dv) Then Continue For
                Dim nL As Integer = 0
                Try : nL = dv.DVLines2d.Count : Catch : End Try
                If nL <= 0 Then Continue For
                Dim scan = ScanOrthogonalViewForVerticalDimension(dv, i, lab, mainHorizontalView, mainCx, mainCy, sheetW, sheetH)
                all.Add(scan)
            Next

            For Each s In all.OrderByDescending(Function(x) x.Score)
                lab.Log("VIEW_SELECT", "VERTICAL_CANDIDATE",
                        "idx=" & s.Idx.ToString(CultureInfo.InvariantCulture) &
                        " name=" & s.Name & " score=" & s.Score.ToString(CultureInfo.InvariantCulture) &
                        " reason=" & s.ScoreReason)
            Next
            Return all
        End Function

        Private Shared Function ScanOrthogonalViewForVerticalDimension(dv As DrawingView,
                                                                       idx As Integer,
                                                                       lab As DimLabLogger,
                                                                       mainHorizontalView As DrawingView,
                                                                       mainCx As Double,
                                                                       mainCy As Double,
                                                                       sheetW As Double,
                                                                       sheetH As Double) As VerticalViewScanResult
            Dim nm = SafeDrawingViewName(dv)
            lab.Log("VERTICAL_VIEW_SCAN", "VIEW", "idx=" & idx.ToString(CultureInfo.InvariantCulture) & " name=" & nm)

            Dim session As New LabLineReadSession()
            Dim bounds = BuildDVLineEndpointBounds(dv, lab, session, quiet:=True)
            Dim nL As Integer = 0
            Try : nL = dv.DVLines2d.Count : Catch : End Try

            If bounds Is Nothing Then
                lab.Log("VERTICAL_VIEW_SCAN", "BOUNDS", "width=0 height=0 (abort)")
                lab.Log("VERTICAL_VIEW_SCAN", "BOUNDARIES", "top=0 bottom=0 left=0 right=0")
                lab.Log("VERTICAL_VIEW_SCAN", "BEST_SAMEX", "pBottom=n/a pTop=n/a sameXDelta=n/a")
                lab.Log("VERTICAL_VIEW_SCAN", "SCORE", "score=-1000 reason=no_bounds")
                Return New VerticalViewScanResult With {
                    .View = dv, .Idx = idx, .Name = nm, .Score = -1000, .ScoreReason = "no_bounds",
                    .Bounds = Nothing, .DvBounds = New DVBounds(), .TolB = 0.0R,
                    .BottomL = New List(Of DVEndpointCandidate)(),
                    .TopL = New List(Of DVEndpointCandidate)(),
                    .LeftL = New List(Of DVEndpointCandidate)(),
                    .RightL = New List(Of DVEndpointCandidate)(),
                    .BestSameXDelta = Double.NaN,
                    .BestPb = Nothing, .BestPt = Nothing,
                    .HasHorizTopBottomOverlap = False,
                    .LineCount = nL
                }
            End If

            Dim dvB = ToDVBounds(bounds)
            lab.Log("VERTICAL_VIEW_SCAN", "BOUNDS", "width=" & Gd(bounds.ExpectedWidth) & " height=" & Gd(bounds.ExpectedHeight))

            Dim tolB = Math.Max(Math.Max(bounds.ExpectedWidth, bounds.ExpectedHeight) * 0.002R, 0.000001R)
            Dim candidates As New List(Of DVEndpointCandidate)()
            BuildEndpointCandidates(dv, candidates, lab, session)

            Dim leftL As New List(Of DVEndpointCandidate)()
            Dim rightL As New List(Of DVEndpointCandidate)()
            Dim bottomL As New List(Of DVEndpointCandidate)()
            Dim topL As New List(Of DVEndpointCandidate)()
            ClassifyBoundaries(candidates, bounds, tolB, leftL, rightL, bottomL, topL, lab, quiet:=True)

            lab.Log("VERTICAL_VIEW_SCAN", "BOUNDARIES",
                    "top=" & topL.Count.ToString(CultureInfo.InvariantCulture) &
                    " bottom=" & bottomL.Count.ToString(CultureInfo.InvariantCulture) &
                    " left=" & leftL.Count.ToString(CultureInfo.InvariantCulture) &
                    " right=" & rightL.Count.ToString(CultureInfo.InvariantCulture))

            Dim bestA As DVEndpointCandidate = Nothing
            Dim bestC As DVEndpointCandidate = Nothing
            Dim bestDx = Double.MaxValue
            For Each a In bottomL
                For Each c In topL
                    Dim dx = Math.Abs(a.X - c.X)
                    If dx < bestDx Then bestDx = dx : bestA = a : bestC = c
                Next
            Next
            Dim sameXDelta = If(bestA Is Nothing OrElse bestC Is Nothing, Double.PositiveInfinity, bestDx)

            Dim pbStr = If(bestA Is Nothing, "n/a", "(" & Gd(bestA.X) & "," & Gd(bestA.Y) & ")")
            Dim ptStr = If(bestC Is Nothing, "n/a", "(" & Gd(bestC.X) & "," & Gd(bestC.Y) & ")")
            lab.Log("VERTICAL_VIEW_SCAN", "BEST_SAMEX",
                    "pBottom=" & pbStr & " pTop=" & ptStr & " sameXDelta=" & If(Double.IsInfinity(sameXDelta) OrElse Double.IsNaN(sameXDelta), "n/a", Gd(sameXDelta)))

            Dim spanV = Math.Max(bounds.ExpectedWidth, bounds.ExpectedHeight)
            Dim botH = FindLongestHorizontalNearY(dv, bounds.MinY, True, spanV, lab, session)
            Dim topH = FindLongestHorizontalNearY(dv, bounds.MaxY, False, spanV, lab, session)
            Dim hasOv = botH.Line IsNot Nothing AndAlso topH.Line IsNot Nothing AndAlso
                    HorizontalSegmentsOverlapX(botH.X1, botH.X2, topH.X1, topH.X2)

            Dim score As Integer = 0
            Dim reasons As New List(Of String)()

            If nL > 0 Then score += 50 : reasons.Add("lines")
            If bottomL.Count > 0 AndAlso topL.Count > 0 Then
                score += 50
                reasons.Add("top_bottom")
            Else
                score -= 50
                reasons.Add("no_top_bottom")
            End If

            If Not Double.IsInfinity(sameXDelta) AndAlso Not Double.IsNaN(sameXDelta) Then
                If sameXDelta < 0.0005R Then
                    score += 100
                    reasons.Add("sameX_lt_0p0005")
                ElseIf sameXDelta < 0.001R Then
                    score += 50
                    reasons.Add("sameX_lt_0p001")
                End If
                If sameXDelta > 0.004R Then
                    score -= 50
                    reasons.Add("sameX_gt_0p004")
                End If
            Else
                score -= 50
                reasons.Add("no_sameX_pair")
            End If

            Dim hExp = bounds.ExpectedHeight
            If hExp >= 0.02R AndAlso hExp <= 0.2R Then score += 20 : reasons.Add("height_band")

            score += 20
            reasons.Add("not_iso")

            Dim candCx As Double, candCy As Double
            If GetViewCenterOnSheet(dv, candCx, candCy) Then
                Dim xTol = Math.Max(0.005R, sheetW * 0.02R)
                Dim yTol = Math.Max(0.001R, sheetH * 0.002R)
                If candCy < mainCy - yTol AndAlso Math.Abs(candCx - mainCx) <= xTol Then
                    score += 30
                    reasons.Add("likely_bottom_view")
                End If
            End If

            If DrawingViewLabelLooksFlat(nm) Then
                score -= 40
                reasons.Add("flat_name_deprioritized")
            End If

            If ReferenceEquals(dv, mainHorizontalView) Then
                score -= 80
                reasons.Add("is_main_horizontal_view_vertical_penalty")
            End If

            If hasOv Then score += 10 : reasons.Add("horiz_overlap")

            Dim reasonJoined = String.Join(",", reasons.ToArray())
            lab.Log("VERTICAL_VIEW_SCAN", "SCORE", "score=" & score.ToString(CultureInfo.InvariantCulture) & " reason=" & reasonJoined)

            Return New VerticalViewScanResult With {
                .View = dv, .Idx = idx, .Name = nm, .Score = score, .ScoreReason = reasonJoined,
                .Bounds = bounds, .DvBounds = dvB, .TolB = tolB,
                .BottomL = bottomL, .TopL = topL, .LeftL = leftL, .RightL = rightL,
                .BestSameXDelta = sameXDelta, .BestPb = bestA, .BestPt = bestC,
                .HasHorizTopBottomOverlap = hasOv,
                .LineCount = nL
            }
        End Function

        Private Shared Function BuildVerticalTryList(all As List(Of VerticalViewScanResult), main As DrawingView, lab As DimLabLogger) As List(Of VerticalViewScanResult)
            Dim others = all.Where(Function(s) Not ReferenceEquals(s.View, main)).OrderByDescending(Function(s) s.Score).Take(3).ToList()
            Dim picked As List(Of VerticalViewScanResult)
            If others.Count >= 1 Then
                picked = others
            Else
                picked = all.OrderByDescending(Function(s) s.Score).Take(3).ToList()
            End If
            If picked.Count > 0 Then
                Dim w = picked(0)
                lab.Log("VERTICAL_VIEW_SELECT", "SELECT",
                        "selectedIdx=" & w.Idx.ToString(CultureInfo.InvariantCulture) &
                        " name=" & w.Name & " score=" & w.Score.ToString(CultureInfo.InvariantCulture) &
                        " reason=" & w.ScoreReason)
            End If
            Return picked
        End Function

        ''' <summary>Analiza todas las vistas ortogonales con DVLines2d en la hoja y devuelve hasta 3 vistas a probar para la cota vertical total (sin forzar la vista principal si hay alternativas).</summary>
        Private Shared Function FindBestViewForVerticalTotalDimension(sh As Sheet,
                                                                      lab As DimLabLogger,
                                                                      mainHorizontalView As DrawingView,
                                                                      mainCx As Double,
                                                                      mainCy As Double,
                                                                      sheetW As Double,
                                                                      sheetH As Double,
                                                                      ByRef allScanned As List(Of VerticalViewScanResult)) As List(Of VerticalViewScanResult)
            allScanned = EnumerateAndScoreVerticalViewCandidates(sh, lab, mainHorizontalView, mainCx, mainCy, sheetW, sheetH)
            Dim mainScan = allScanned.FirstOrDefault(Function(s) ReferenceEquals(s.View, mainHorizontalView))
            If mainScan IsNot Nothing AndAlso
                    (mainScan.BottomL.Count = 0 OrElse mainScan.TopL.Count = 0 OrElse mainScan.BestSameXDelta > 0.004R) Then
                lab.Log("VERTICAL", "MAIN_VIEW_SKIP", "reason=no_clean_sameX_pair")
            End If
            Return BuildVerticalTryList(allScanned, mainHorizontalView, lab)
        End Function

        Private Shared Function TryVerticalObjectBetweenHorizontalRefsLab(dims As Dimensions,
                                                                        dv As DrawingView,
                                                                        draft As DraftDocument,
                                                                        b As LineBounds,
                                                                        lab As DimLabLogger,
                                                                        readSession As LabLineReadSession,
                                                                        created As List(Of Object),
                                                                        ctx As LabRunContext,
                                                                        ByRef outD As Dimension) As String
            outD = Nothing
            Dim spanV = Math.Max(b.ExpectedWidth, b.ExpectedHeight)
            Dim bot = FindLongestHorizontalNearY(dv, b.MinY, True, spanV, lab, readSession)
            Dim top = FindLongestHorizontalNearY(dv, b.MaxY, False, spanV, lab, readSession)
            If bot.Line Is Nothing OrElse top.Line Is Nothing Then Return "FAIL"
            If Not HorizontalSegmentsOverlapX(bot.X1, bot.X2, top.X1, top.X2) Then Return "FAIL"
            Dim refB As Object = Nothing, refT As Object = Nothing
            Try : refB = CallByName(bot.Line, "Reference", CallType.Get) : Catch : End Try
            Try : refT = CallByName(top.Line, "Reference", CallType.Get) : Catch : End Try
            If refB Is Nothing OrElse refT Is Nothing Then Return "FAIL"
            Dim xMid = (b.MinX + b.MaxX) / 2.0R
            Dim by = (bot.Y1 + bot.Y2) / 2.0R
            Dim ty = (top.Y1 + top.Y2) / 2.0R
            Dim bx = PickXNearAnchor(bot.X1, bot.X2, xMid)
            Dim tx = PickXNearAnchor(top.X1, top.X2, xMid)
            Dim db = ToDVBounds(b)
            Dim axisX = xMid
            Dim axisY = by
            Dim axisLine As Object = Nothing
            Try
                axisLine = ctx.Sheet.Lines2d.AddBy2Points(axisX, axisY, axisX, axisY + 0.05R)
                created.Add(axisLine)
                lab.Log("AXISMODE", "AXIS_LINE", "start=(" & Gd(axisX) & "," & Gd(axisY) & ") end=(" & Gd(axisX) & "," & Gd(axisY + 0.05R) & ")")
            Catch ex As Exception
                lab.Log("AXISMODE", "AXIS_LINE", "fail=" & ex.Message)
            End Try

            Dim rDefault = TryVerticalAxisModeCandidate(dims, dv, draft, created, lab, ctx, "Vertical_Default_HorizontalVertical", "igDimAxisModeDefault", Nothing, refB, bx, by, refT, tx, ty, b.ExpectedHeight, outD, db)
            ctx.AxisModeDefaultResult = rDefault
            If IsVerticalLabSuccess(outD, rDefault, b.ExpectedHeight, draft, dv, lab, ctx) Then
                ctx.BestVerticalAxisMode = "Default"
                Return rDefault
            End If

            Dim rImplied = TryVerticalAxisModeCandidate(dims, dv, draft, created, lab, ctx, "Vertical_Implied_ByTwoPoints", "igDimAxisModeImplied", Nothing, refB, bx, by, refT, tx, ty, b.ExpectedHeight, outD, db)
            ctx.AxisModeImpliedResult = rImplied
            If IsVerticalLabSuccess(outD, rImplied, b.ExpectedHeight, draft, dv, lab, ctx) Then
                ctx.BestVerticalAxisMode = "Implied"
                Return rImplied
            End If

            Dim rExplicit = TryVerticalAxisModeCandidate(dims, dv, draft, created, lab, ctx, "Vertical_Explicit_VerticalAxis", "igDimAxisModeExplicit", axisLine, refB, bx, by, refT, tx, ty, b.ExpectedHeight, outD, db)
            ctx.AxisModeExplicitResult = rExplicit
            If IsVerticalLabSuccess(outD, rExplicit, b.ExpectedHeight, draft, dv, lab, ctx) Then
                ctx.BestVerticalAxisMode = "Explicit"
                Return rExplicit
            End If

            Return MaxResult(MaxResult(rDefault, rImplied), rExplicit)
        End Function

        Private Shared Function TryVerticalAxisModeCandidate(dims As Dimensions,
                                                             dv As DrawingView,
                                                             draft As DraftDocument,
                                                             created As List(Of Object),
                                                             lab As DimLabLogger,
                                                             ctx As LabRunContext,
                                                             testName As String,
                                                             modeName As String,
                                                             axisLine As Object,
                                                             refB As Object, bx As Double, by As Double,
                                                             refT As Object, tx As Double, ty As Double,
                                                             expected As Double,
                                                             ByRef outD As Dimension,
                                                             db As DVBounds) As String
            outD = Nothing
            SetDimensionsAxisMode(dims, modeName, axisLine, lab, testName)
            lab.Log("AXISMODE", "TRY", "name=" & testName & " mode=" & modeName)
            Dim r = TryOneDistanceRaw(dims, dv, draft, testName, refB, bx, by, refT, tx, ty, expected, False, False, created, lab, ctx, quietLog:=True, outCreatedDimension:=outD)
            If outD IsNot Nothing Then
                TrySetDimensionAxisPostCreate(outD, modeName, axisLine, lab)
                LogVerticalOutcome(outD, r, expected, db, draft, dv, testName, lab, ctx)
                lab.Log("AXISMODE", "VISIBLE", "name=" & testName & " " & LabVisibleOk(outD, ctx.Sheet, draft, ctx.App, lab, testName).ToString(CultureInfo.InvariantCulture))
            End If
            lab.Log("AXISMODE", "RESULT", "name=" & testName & " result=" & r)
            Return r
        End Function

        Private Shared Sub SetDimensionsAxisMode(dims As Dimensions, modeName As String, axisLine As Object, lab As DimLabLogger, testName As String)
            Try
                Dim modeVal As Integer = 0
                If String.Equals(modeName, "igDimAxisModeImplied", StringComparison.OrdinalIgnoreCase) Then
                    modeVal = CInt(SolidEdgeConstants.DimAxisModeConstants.igDimAxisModeImplied)
                ElseIf String.Equals(modeName, "igDimAxisModeExplicit", StringComparison.OrdinalIgnoreCase) Then
                    modeVal = CInt(SolidEdgeConstants.DimAxisModeConstants.igDimAxisModeExplicit)
                Else
                    modeVal = CInt(SolidEdgeConstants.DimAxisModeConstants.igDimAxisModeDefault)
                End If
                CallByName(dims, "AxisMode", CallType.Let, modeVal)
            Catch ex As Exception
                lab.Log("AXISMODE", "SET_FAIL", "name=" & testName & " prop=AxisMode error=" & ex.Message)
            End Try
            If axisLine IsNot Nothing Then
                Try
                    CallByName(dims, "Axis", CallType.Let, axisLine)
                Catch ex As Exception
                    lab.Log("AXISMODE", "SET_FAIL", "name=" & testName & " prop=Axis error=" & ex.Message)
                End Try
            End If
        End Sub

        Private Shared Sub TrySetDimensionAxisPostCreate(d As Dimension, modeName As String, axisLine As Object, lab As DimLabLogger)
            Dim modeVal As Integer = CInt(SolidEdgeConstants.DimAxisModeConstants.igDimAxisModeDefault)
            If String.Equals(modeName, "igDimAxisModeImplied", StringComparison.OrdinalIgnoreCase) Then modeVal = CInt(SolidEdgeConstants.DimAxisModeConstants.igDimAxisModeImplied)
            If String.Equals(modeName, "igDimAxisModeExplicit", StringComparison.OrdinalIgnoreCase) Then modeVal = CInt(SolidEdgeConstants.DimAxisModeConstants.igDimAxisModeExplicit)
            Try
                CallByName(d, "MeasurementAxisEx", CallType.Let, modeVal)
                lab.Log("DIM_AXIS", "SET", "MeasurementAxisEx=" & modeName & " OK")
            Catch ex As Exception
                lab.Log("DIM_AXIS", "SET", "MeasurementAxisEx=" & modeName & " FAIL " & ex.Message)
            End Try
            Try
                CallByName(d, "MeasurementAxisDirection", CallType.Let, True)
                lab.Log("DIM_AXIS", "SET", "MeasurementAxisDirection=True OK")
            Catch ex As Exception
                lab.Log("DIM_AXIS", "SET", "MeasurementAxisDirection=True FAIL " & ex.Message)
            End Try
            If axisLine IsNot Nothing Then
                Try
                    CallByName(d, "Axis", CallType.Let, axisLine)
                    lab.Log("DIM_AXIS", "SET", "Axis=axisLine OK")
                Catch ex As Exception
                    lab.Log("DIM_AXIS", "SET", "Axis=axisLine FAIL " & ex.Message)
                End Try
            End If
        End Sub

        Private Shared Function CreateVerticalTotalDimensionFromDVRef_Lab(dims As Dimensions,
                                                                          verticalView As DrawingView,
                                                                          draft As DraftDocument,
                                                                          scan As VerticalViewScanResult,
                                                                          created As List(Of Object),
                                                                          lab As DimLabLogger,
                                                                          ctx As LabRunContext) As String
            Dim dvBounds = scan.DvBounds
            Dim boundsLb = scan.Bounds
            Dim bottomL = scan.BottomL
            Dim topL = scan.TopL

            lab.Log("VERTICAL", "VIEW", "idx=" & scan.Idx.ToString(CultureInfo.InvariantCulture) & " name=" & scan.Name)
            lab.Log("VERTICAL", "EXPECTED", "height=" & Gd(dvBounds.ExpectedHeight))

            Dim readSession As New LabLineReadSession()
            lab.Log("VERTICAL", "LAB", "START CreateVerticalTotalDimensionFromDVRef_Lab scanned_view")
            ctx.SummaryVertCreate = "FAIL"
            If boundsLb Is Nothing OrElse bottomL.Count = 0 OrElse topL.Count = 0 Then
                lab.Log("VERTICAL", "RESULT", "FAIL no_boundary_endpoints_bottom_top")
                Return "FAIL"
            End If

            Dim outD As Dimension = Nothing
            Const MaxStrategies As Integer = 3
            Dim used As Integer = 0
            Const SkipIntermediateVerticalStrategies As Boolean = True

            If Not SkipIntermediateVerticalStrategies AndAlso scan.BestPb IsNot Nothing AndAlso scan.BestPt IsNot Nothing AndAlso used < MaxStrategies Then
                used += 1
                Dim r0 = TryVerticalLabStep(dims, verticalView, draft, dvBounds, scan.BestPb, scan.BestPt, "Vertical_BestSameX_Keypoints", created, lab, ctx, outD)
                If IsVerticalLabSuccess(outD, r0, dvBounds.ExpectedHeight, draft, verticalView, lab, ctx) Then
                    lab.Log("VERTICAL", "RESULT", "SUCCESS_VALUE_VISIBLE_CONNECTION_UNCERTAIN")
                    lab.Log("VERTICAL", "STOP", "reason=valid_value_visible_found")
                    lab.Log("VERTICAL", "SELECTED_VIEW", "DrawingView " & scan.Idx.ToString(CultureInfo.InvariantCulture))
                    ctx.SummaryVertCreate = "SUCCESS"
                    ctx.SummaryVertValue = Gd(ReadDimValue(outD))
                    ctx.SummaryVertConnected = CheckConnected(draft, verticalView, outD, lab, "Vertical_BestSameX_Keypoints", ctx)
                    ctx.SummaryVertVisible = LabVisibleOk(outD, ctx.Sheet, draft, ctx.App, lab, "Vertical_BestSameX_Keypoints")
                    ctx.SummaryVertValueClass = "OK"
                    Return "SUCCESS_VERTICAL_BESTSAMEX"
                End If
            End If

            If Not SkipIntermediateVerticalStrategies AndAlso used < MaxStrategies Then
                used += 1
                Dim pb = PickClosestX(bottomL, boundsLb.MaxX)
                Dim pt = PickClosestX(topL, pb.X)
                Dim rA = TryVerticalLabStep(dims, verticalView, draft, dvBounds, pb, pt, "Vertical_RightSide_Keypoints", created, lab, ctx, outD)
                If IsVerticalLabSuccess(outD, rA, dvBounds.ExpectedHeight, draft, verticalView, lab, ctx) Then
                    lab.Log("VERTICAL", "RESULT", "SUCCESS_VALUE_VISIBLE_CONNECTION_UNCERTAIN")
                    lab.Log("VERTICAL", "STOP", "reason=valid_value_visible_found")
                    lab.Log("VERTICAL", "SELECTED_VIEW", "DrawingView " & scan.Idx.ToString(CultureInfo.InvariantCulture))
                    ctx.SummaryVertCreate = "SUCCESS"
                    ctx.SummaryVertValue = Gd(ReadDimValue(outD))
                    ctx.SummaryVertConnected = CheckConnected(draft, verticalView, outD, lab, "Vertical_RightSide_Keypoints", ctx)
                    ctx.SummaryVertVisible = LabVisibleOk(outD, ctx.Sheet, draft, ctx.App, lab, "Vertical_RightSide_Keypoints")
                    ctx.SummaryVertValueClass = "OK"
                    Return "SUCCESS_VERTICAL_RIGHT"
                End If
            End If

            If used < MaxStrategies Then
                used += 1
                If scan.HasHorizTopBottomOverlap Then
                    Dim rD = TryVerticalObjectBetweenHorizontalRefsLab(dims, verticalView, draft, boundsLb, lab, readSession, created, ctx, outD)
                    If IsVerticalLabSuccess(outD, rD, dvBounds.ExpectedHeight, draft, verticalView, lab, ctx) Then
                        lab.Log("VERTICAL", "RESULT", "SUCCESS_VALUE_VISIBLE_CONNECTION_UNCERTAIN")
                        lab.Log("VERTICAL", "STOP", "reason=valid_value_visible_found")
                        lab.Log("VERTICAL", "SELECTED_VIEW", "DrawingView " & scan.Idx.ToString(CultureInfo.InvariantCulture))
                        ctx.SummaryVertCreate = "SUCCESS"
                        ctx.SummaryVertValue = Gd(ReadDimValue(outD))
                        ctx.SummaryVertConnected = CheckConnected(draft, verticalView, outD, lab, "Vertical_ObjectBetweenHorizontalRefs", ctx)
                        ctx.SummaryVertVisible = LabVisibleOk(outD, ctx.Sheet, draft, ctx.App, lab, "Vertical_ObjectBetweenHorizontalRefs")
                        ctx.SummaryVertValueClass = "OK"
                        Return "SUCCESS_VERTICAL_HORIZ_REFS"
                    End If
                Else
                    Dim pb2 = PickClosestX(bottomL, boundsLb.MinX)
                    Dim pt2 = PickClosestX(topL, pb2.X)
                    Dim rB = TryVerticalLabStep(dims, verticalView, draft, dvBounds, pb2, pt2, "Vertical_LeftSide_Keypoints", created, lab, ctx, outD)
                    If IsVerticalLabSuccess(outD, rB, dvBounds.ExpectedHeight, draft, verticalView, lab, ctx) Then
                        lab.Log("VERTICAL", "RESULT", "SUCCESS_VALUE_VISIBLE_CONNECTION_UNCERTAIN")
                        lab.Log("VERTICAL", "STOP", "reason=valid_value_visible_found")
                        lab.Log("VERTICAL", "SELECTED_VIEW", "DrawingView " & scan.Idx.ToString(CultureInfo.InvariantCulture))
                        ctx.SummaryVertCreate = "SUCCESS"
                        ctx.SummaryVertValue = Gd(ReadDimValue(outD))
                        ctx.SummaryVertConnected = CheckConnected(draft, verticalView, outD, lab, "Vertical_LeftSide_Keypoints", ctx)
                        ctx.SummaryVertVisible = LabVisibleOk(outD, ctx.Sheet, draft, ctx.App, lab, "Vertical_LeftSide_Keypoints")
                        ctx.SummaryVertValueClass = "OK"
                        Return "SUCCESS_VERTICAL_LEFT"
                    End If
                End If
            End If

            If outD IsNot Nothing Then
                ctx.SummaryVertValue = Gd(ReadDimValue(outD))
                ctx.SummaryVertValueClass = ClassifyWrongDimensionValue(ReadDimValue(outD), dvBounds, "vertical_last", lab)
                ctx.SummaryVertCreate = "WRONG_VALUE"
            End If

            Return "FAIL"
        End Function

        Private Shared Function SummarizeHorizontalLabOutcome(sumHoriz As String) As String
            If sumHoriz Is Nothing Then Return "FAIL"
            If sumHoriz.IndexOf("SUCCESS", StringComparison.OrdinalIgnoreCase) >= 0 Then Return "SUCCESS"
            If sumHoriz.IndexOf("SKIPPED", StringComparison.OrdinalIgnoreCase) >= 0 OrElse sumHoriz.IndexOf("SKIP", StringComparison.OrdinalIgnoreCase) >= 0 Then Return "SKIPPED"
            Return "FAIL"
        End Function

        Private Shared Function SummarizeVerticalLabOutcome(sumVert As String, ctx As LabRunContext) As String
            If sumVert IsNot Nothing AndAlso sumVert.IndexOf("SKIPPED", StringComparison.OrdinalIgnoreCase) >= 0 Then Return "SKIPPED"
            If sumVert IsNot Nothing AndAlso sumVert.IndexOf("SUCCESS", StringComparison.OrdinalIgnoreCase) >= 0 Then Return "SUCCESS"
            If ctx IsNot Nothing AndAlso String.Equals(ctx.SummaryVertCreate, "WRONG_VALUE", StringComparison.OrdinalIgnoreCase) Then Return "WRONG_VALUE"
            If sumVert IsNot Nothing AndAlso sumVert.IndexOf("WRONG", StringComparison.OrdinalIgnoreCase) >= 0 Then Return "WRONG_VALUE"
            Return "FAIL"
        End Function

        Private Shared Function BuildRecommendedNextStepV2(ctx As LabRunContext, sumHoriz As String, sumVert As String) As String
            If ctx Is Nothing Then Return "revisar logs DIMLAB"
            Dim hOk = sumHoriz.IndexOf("SUCCESS", StringComparison.OrdinalIgnoreCase) >= 0
            Dim vOk = sumVert.IndexOf("SUCCESS", StringComparison.OrdinalIgnoreCase) >= 0
            Dim vWrong = sumVert.IndexOf("WRONG", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                (ctx IsNot Nothing AndAlso ctx.SummaryVertCreate = "WRONG_VALUE")

            Dim sameView = Not String.IsNullOrEmpty(ctx.SummaryHorizontalViewLabel) AndAlso
                String.Equals(ctx.SummaryHorizontalViewLabel, ctx.SummaryVerticalViewLabel, StringComparison.Ordinal)

            If hOk AndAlso vOk AndAlso Not sameView Then Return "convertir CleanFull en base de motor experimental"
            If hOk AndAlso vOk Then Return "crear motor experimental DVRef total dimensions"
            If hOk AndAlso sumVert.IndexOf("SKIPPED", StringComparison.OrdinalIgnoreCase) >= 0 Then Return "ejecutar DimLabMode=VerticalOnly"
            If hOk AndAlso vWrong Then Return "probar DimInitData/orientación explícita para vertical"
            If Not hOk Then Return "no tocar vertical; corregir horizontal"
            If hOk AndAlso Not vOk Then Return "probar DimInitData/orientación explícita para vertical"
            Return "revisar logs DIMLAB"
        End Function

        Private Shared Function ResolveSheet(draft As DraftDocument, sheet As Sheet, lab As DimLabLogger) As Sheet
            Dim sh = sheet
            If sh Is Nothing Then
                Try : sh = draft.ActiveSheet : Catch : End Try
            End If
            If sh Is Nothing Then Return Nothing
            Dim nm As String = ""
            Try : nm = sh.Name : Catch : End Try
            lab.Log("SHEET", "RESOLVED", "name=" & nm)
            Return sh
        End Function

        Private Shared Function SafeDimensionsCount(sh As Sheet) As Integer
            Try
                Return CInt(sh.Dimensions.Count)
            Catch
                Return -1
            End Try
        End Function

        Private Shared Sub CleanStartDimensions(sh As Sheet, lab As DimLabLogger)
            If sh Is Nothing Then Return
            Dim beforeCnt = SafeDimensionsCount(sh)
            lab.Log("CLEAN_START", "DIMENSIONS_BEFORE", beforeCnt.ToString(CultureInfo.InvariantCulture))
            Dim deleted As Integer = 0
            Dim n As Integer = beforeCnt
            For i As Integer = n To 1 Step -1
                Dim d As Object = Nothing
                Try
                    d = sh.Dimensions.Item(i)
                Catch
                    Continue For
                End Try
                If d Is Nothing Then Continue For
                If TryDeleteDimension(d) Then deleted += 1
            Next
            lab.Log("CLEAN_START", "DELETED", deleted.ToString(CultureInfo.InvariantCulture))
            lab.Log("CLEAN_START", "DIMENSIONS_AFTER", SafeDimensionsCount(sh).ToString(CultureInfo.InvariantCulture))
        End Sub

        Private Shared Function TryDeleteDimension(d As Object) As Boolean
            If d Is Nothing Then Return False
            Try
                CallByName(d, "Delete", CallType.Method)
                Return True
            Catch
                Return False
            End Try
        End Function

        Private Shared Function Gd(v As Double) As String
            If Double.IsNaN(v) OrElse Double.IsInfinity(v) Then Return "NaN"
            Return v.ToString("G17", CultureInfo.InvariantCulture)
        End Function

        Private Shared Function NearlyEqual(a As Double, b As Double, tol As Double) As Boolean
            Return Math.Abs(a - b) <= tol
        End Function

        Private Shared Sub CleanupIntermediateDimensionsKeepingPrimary(sh As Sheet, ctx As LabRunContext, lab As DimLabLogger)
            Dim targetH As Double = Double.NaN
            Dim targetV As Double = Double.NaN
            Try
                targetH = Double.Parse(ctx.SummaryHorizValue, CultureInfo.InvariantCulture)
            Catch
            End Try
            Try
                targetV = Double.Parse(ctx.SummaryVertValue, CultureInfo.InvariantCulture)
            Catch
            End Try
            If Double.IsNaN(targetH) AndAlso Double.IsNaN(targetV) Then
                lab.Log("CLEAN_FINAL", "SKIP", "reason=missing_target_values")
                Return
            End If

            Dim keepH As Integer = 0
            Dim keepV As Integer = 0
            Dim n As Integer = SafeDimensionsCount(sh)
            For i As Integer = n To 1 Step -1
                Dim d As Dimension = Nothing
                Try
                    d = TryCast(sh.Dimensions.Item(i), Dimension)
                Catch
                    d = Nothing
                End Try
                If d Is Nothing Then Continue For

                Dim v As Double = ReadDimValue(d)
                Dim keep As Boolean = False
                If Not Double.IsNaN(targetH) AndAlso NearlyEqual(v, targetH, 0.002R) AndAlso keepH = 0 Then
                    keep = True
                    keepH += 1
                ElseIf Not Double.IsNaN(targetV) AndAlso NearlyEqual(v, targetV, 0.002R) AndAlso keepV = 0 Then
                    keep = True
                    keepV += 1
                End If

                If Not keep Then
                    lab.Log("CLEAN_FINAL", "DELETE_INTERMEDIATE", "idx=" & i.ToString(CultureInfo.InvariantCulture) & " value=" & Gd(v))
                    TryDeleteDimension(d)
                End If
            Next

            lab.Log("CLEAN_FINAL", "DONE", "keptH=" & keepH.ToString(CultureInfo.InvariantCulture) &
                    " keptV=" & keepV.ToString(CultureInfo.InvariantCulture) &
                    " finalCount=" & SafeDimensionsCount(sh).ToString(CultureInfo.InvariantCulture))
        End Sub

        Private Shared Function TryViewRangeSheet(view As DrawingView, ByRef vxMin As Double, ByRef vyMin As Double, ByRef vxMax As Double, ByRef vyMax As Double) As Boolean
            vxMin = 0.0R : vyMin = 0.0R : vxMax = 0.0R : vyMax = 0.0R
            Try
                view.Range(vxMin, vyMin, vxMax, vyMax)
                Return True
            Catch
                Return False
            End Try
        End Function

        Private Shared Function TrySheetSizeM(sheet As Sheet, ByRef w As Double, ByRef h As Double) As Boolean
            w = 0.0R : h = 0.0R
            Try
                w = CDbl(sheet.SheetSetup.SheetWidth)
                h = CDbl(sheet.SheetSetup.SheetHeight)
                Return w > 0.0001R AndAlso h > 0.0001R
            Catch
                Return False
            End Try
        End Function

        Private Shared Function ComputeFreeTopY(sheet As Sheet) As Double
            Dim w As Double, h As Double
            If TrySheetSizeM(sheet, w, h) Then
                Dim margin As Double = 0.02R
                Return Math.Max(0.0R, h - margin)
            End If
            Return 0.276R
        End Function

        Private Shared Function ComputeFreeRightX(sheet As Sheet) As Double
            Dim w As Double, h As Double
            If TrySheetSizeM(sheet, w, h) Then
                Dim margin As Double = 0.02R
                Return Math.Max(0.0R, w - margin)
            End If
            Return 0.4R
        End Function

        Private Shared Function ComputeDesiredTrackHorizontal(availableTop As Double) As Double
            Return DimLabHorizontalGap
        End Function

        Private Shared Function ComputeDesiredTrackVertical(availableRight As Double) As Double
            Return DimLabVerticalGap
        End Function

        Private Shared Function TryReadDimensionRange(d As Dimension, ByRef rx1 As Double, ByRef ry1 As Double, ByRef rx2 As Double, ByRef ry2 As Double) As Boolean
            rx1 = 0.0R : ry1 = 0.0R : rx2 = 0.0R : ry2 = 0.0R
            Try
                d.Range(rx1, ry1, rx2, ry2)
                Return True
            Catch
                Return False
            End Try
        End Function

        Private Shared Function DimensionRangeIsInvalid(rx1 As Double, ry1 As Double, rx2 As Double, ry2 As Double) As Boolean
            Dim w = Math.Abs(rx2 - rx1)
            Dim h = Math.Abs(ry2 - ry1)
            If w < 0.000000000001R AndAlso h < 0.000000000001R Then Return True
            If Math.Abs(rx1) < 0.000000000000001R AndAlso Math.Abs(ry1) < 0.000000000000001R AndAlso Math.Abs(rx2) < 0.000000000000001R AndAlso Math.Abs(ry2) < 0.000000000000001R Then Return True
            Return False
        End Function

        Private Shared Function FormatRangeStr(rx1 As Double, ry1 As Double, rx2 As Double, ry2 As Double) As String
            Return Gd(rx1) & "," & Gd(ry1) & ".." & Gd(rx2) & "," & Gd(ry2)
        End Function

        Private Shared Sub LogDimensionPosition(tag As String, d As Dimension, lab As DimLabLogger)
            If d Is Nothing Then Return
            Dim rx1 As Double, ry1 As Double, rx2 As Double, ry2 As Double
            If TryReadDimensionRange(d, rx1, ry1, rx2, ry2) Then
                Dim minX = Math.Min(rx1, rx2)
                Dim maxX = Math.Max(rx1, rx2)
                Dim minY = Math.Min(ry1, ry2)
                Dim maxY = Math.Max(ry1, ry2)
                Dim cx = (minX + maxX) * 0.5R
                Dim cy = (minY + maxY) * 0.5R
                lab.Log("PLACE", "POSITION", "name=" & tag &
                        " center=(" & Gd(cx) & "," & Gd(cy) & ")" &
                        " bbox=(" & Gd(minX) & "," & Gd(minY) & ")->(" & Gd(maxX) & "," & Gd(maxY) & ")")
            Else
                lab.Log("PLACE", "POSITION", "name=" & tag & " unreadable_range")
            End If

            Try
                Dim tdx As Double = 0.0R
                Dim tdy As Double = 0.0R
                d.GetTextOffsets(tdx, tdy)
                lab.Log("PLACE", "TEXT_OFFSETS", "name=" & tag & " dx=" & Gd(tdx) & " dy=" & Gd(tdy))
            Catch
                lab.Log("PLACE", "TEXT_OFFSETS", "name=" & tag & " unreadable")
            End Try
        End Sub

        Private Shared Sub LogFinalDimensionPositionsCompact(sh As Sheet, lab As DimLabLogger)
            If sh Is Nothing Then Return
            Dim n As Integer = SafeDimensionsCount(sh)
            lab.Log("DIM_POS", "COUNT", n.ToString(CultureInfo.InvariantCulture))
            For i As Integer = 1 To n
                Dim d As Dimension = Nothing
                Try
                    d = TryCast(sh.Dimensions.Item(i), Dimension)
                Catch
                    d = Nothing
                End Try
                If d Is Nothing Then Continue For

                Dim v As Double = Double.NaN
                Try : v = CDbl(CallByName(d, "Value", CallType.Get)) : Catch : End Try
                Dim st As String = ReadDimensionStyleName(d, lab)

                Dim rx1 As Double, ry1 As Double, rx2 As Double, ry2 As Double
                Dim centerTxt As String = "unreadable"
                Dim bboxTxt As String = "unreadable"
                If TryReadDimensionRange(d, rx1, ry1, rx2, ry2) Then
                    Dim minX = Math.Min(rx1, rx2)
                    Dim maxX = Math.Max(rx1, rx2)
                    Dim minY = Math.Min(ry1, ry2)
                    Dim maxY = Math.Max(ry1, ry2)
                    centerTxt = "(" & Gd((minX + maxX) * 0.5R) & "," & Gd((minY + maxY) * 0.5R) & ")"
                    bboxTxt = "(" & Gd(minX) & "," & Gd(minY) & ")->(" & Gd(maxX) & "," & Gd(maxY) & ")"
                End If

                Dim tdx As Double = Double.NaN, tdy As Double = Double.NaN
                Try : d.GetTextOffsets(tdx, tdy) : Catch : End Try

                lab.Log("DIM_POS", "ITEM",
                        "idx=" & i.ToString(CultureInfo.InvariantCulture) &
                        " value=" & Gd(v) &
                        " style=" & st &
                        " center=" & centerTxt &
                        " bbox=" & bboxTxt &
                        " textOffsets=(" & Gd(tdx) & "," & Gd(tdy) & ")")
            Next
        End Sub

        Private Shared Function RectanglesOverlap2D(ax1 As Double, ay1 As Double, ax2 As Double, ay2 As Double,
                                                    bx1 As Double, by1 As Double, bx2 As Double, by2 As Double) As Boolean
            Dim aMinX = Math.Min(ax1, ax2), aMaxX = Math.Max(ax1, ax2)
            Dim aMinY = Math.Min(ay1, ay2), aMaxY = Math.Max(ay1, ay2)
            Dim bMinX = Math.Min(bx1, bx2), bMaxX = Math.Max(bx1, bx2)
            Dim bMinY = Math.Min(by1, by2), bMaxY = Math.Max(by1, by2)
            Return Not (aMaxX < bMinX OrElse aMinX > bMaxX OrElse aMaxY < bMinY OrElse aMinY > bMaxY)
        End Function

        Private Shared Function GetVisibilityReasonQuiet(dimObj As Dimension, sheet As Sheet) As String
            Dim rx1 As Double, ry1 As Double, rx2 As Double, ry2 As Double
            If Not TryReadDimensionRange(dimObj, rx1, ry1, rx2, ry2) Then Return "range_read_fail"
            If DimensionRangeIsInvalid(rx1, ry1, rx2, ry2) Then Return "range_zero"
            Dim sw As Double, sh As Double
            If Not TrySheetSizeM(sheet, sw, sh) Then Return "in_sheet"
            If Not RectanglesOverlap2D(rx1, ry1, rx2, ry2, 0.0R, 0.0R, sw, sh) Then Return "range_outside_sheet"
            Return "in_sheet"
        End Function

        Private Shared Function ValidateDimensionVisibleOnSheet(dimObj As Dimension,
                                                               sheet As Sheet,
                                                               draft As DraftDocument,
                                                               app As Application,
                                                               testName As String,
                                                               lab As DimLabLogger) As Boolean
            Dim vf = ValidateVisibleFinal(dimObj, sheet, draft, app, testName, lab)
            Return vf.Materialized
        End Function

        Private Shared Function SafeGetTrackDistance(d As Dimension) As Double
            Try
                Return CDbl(d.TrackDistance)
            Catch
                Return Double.NaN
            End Try
        End Function

        Private Shared Function ForceDimensionVisibleInSheet(dimObj As Dimension,
                                                             view As DrawingView,
                                                             sheet As Sheet,
                                                             draft As DraftDocument,
                                                             app As Application,
                                                             testName As String,
                                                             axisIntent As String,
                                                             lab As DimLabLogger,
                                                             Optional logAltPlacement As Boolean = False) As Boolean
            Try
                Dim td0 = SafeGetTrackDistance(dimObj)
                lab.Log("PLACE", "BEFORE", "name=" & testName & " track=" & Gd(td0))

                Dim vxMin As Double, vyMin As Double, vxMax As Double, vyMax As Double
                If Not TryViewRangeSheet(view, vxMin, vyMin, vxMax, vyMax) Then
                    lab.Log("PLACE", "VIEW_RANGE_FAIL", "name=" & testName)
                End If

                Dim freeTop = ComputeFreeTopY(sheet)
                Dim freeRight = ComputeFreeRightX(sheet)
                Dim availableTop = freeTop - vyMax
                Dim availableRight = freeRight - vxMax
                If logAltPlacement AndAlso String.Equals(axisIntent, "H", StringComparison.OrdinalIgnoreCase) AndAlso availableTop < 0.018R Then
                    lab.Log("PLACE", "NOTE", "horizontal_top_close_to_partlist consider_bottom_lane=True")
                End If
                Dim desiredTrack As Double
                Dim side As String
                If String.Equals(axisIntent, "V", StringComparison.OrdinalIgnoreCase) Then
                    desiredTrack = ComputeDesiredTrackVertical(availableRight)
                    If availableRight >= DimLabVerticalGap Then
                        side = "RIGHT"
                        desiredTrack = Math.Abs(desiredTrack)
                    Else
                        side = "LEFT"
                        desiredTrack = -Math.Abs(desiredTrack)
                    End If
                    lab.Log("PLACE", "LANE_CHOICE", "name=" & testName & " axis=V side=" & side & " gap=" & Gd(Math.Abs(desiredTrack)) & " reason=user_requested_2cm_clearance")
                Else
                    desiredTrack = -Math.Abs(ComputeDesiredTrackHorizontal(availableTop))
                    side = "BOTTOM"
                    lab.Log("PLACE", "LANE_CHOICE", "name=" & testName & " axis=H side=BOTTOM gap=" & Gd(Math.Abs(desiredTrack)) & " reason=user_requested_2cm_clearance")
                End If

                lab.Log("PLACE", "DESIRED", "name=" & testName & " desiredTrack=" & Gd(desiredTrack) &
                        " freeTop=" & Gd(freeTop) & " viewMaxY=" & Gd(vyMax) & " availableTop=" & Gd(availableTop) &
                        " freeRight=" & Gd(freeRight) & " viewMaxX=" & Gd(vxMax) & " availableRight=" & Gd(availableRight))

                ' No llamar BreakAlignmentSet: en cotas sin conjunto de alineación SE devuelve E_POINTER (NullReferenceException COM).
                lab.Log("PLACE", "BREAK_ALIGNMENT", "skipped name=" & testName)

                Try
                    dimObj.TrackDistance = desiredTrack
                    lab.Log("PLACE", "TRACK_SET", "name=" & testName & " TrackDistance=" & Gd(desiredTrack) & " OK")
                Catch ex As Exception
                    lab.Log("PLACE", "TRACK_SET_FAIL", "name=" & testName & " " & ex.Message)
                End Try
                Try
                    CallByName(dimObj, "AbsoluteTrackDistance", CallType.Let, desiredTrack)
                    lab.Log("PLACE", "ABS_TRACK_SET", "name=" & testName & " AbsoluteTrackDistance=" & Gd(desiredTrack) & " OK")
                Catch ex As Exception
                    lab.Log("PLACE", "ABS_TRACK_SET_FAIL", "name=" & testName & " " & ex.Message)
                End Try

                Try
                    draft.UpdateAll(True)
                Catch ex As Exception
                    lab.Log("PLACE", "UPDATE_FAIL", ex.Message)
                End Try
                Try
                    If app IsNot Nothing Then app.DoIdle()
                Catch exD As Exception
                    lab.Log("PLACE", "DOIDLE_SKIP", exD.Message)
                End Try

                Dim td1 = SafeGetTrackDistance(dimObj)
                Dim rx1 As Double, ry1 As Double, rx2 As Double, ry2 As Double
                Dim rangeOk = TryReadDimensionRange(dimObj, rx1, ry1, rx2, ry2)
                Dim rangeStr = If(rangeOk, FormatRangeStr(rx1, ry1, rx2, ry2), "unreadable")
                lab.Log("PLACE", "AFTER", "name=" & testName & " track=" & Gd(td1) & " range=" & rangeStr)

                Dim visOk = String.Equals(GetVisibilityReasonQuiet(dimObj, sheet), "in_sheet", StringComparison.OrdinalIgnoreCase)
                lab.Log("PLACE", "VISIBLE_PROBE", "name=" & testName & " " & visOk.ToString(CultureInfo.InvariantCulture) & " reason=" & GetVisibilityReasonQuiet(dimObj, sheet))

                If Not visOk Then
                    lab.Log("PLACE", "VISIBLE_PROBE_RETRY", "name=" & testName & " skipped_track_scaling=True reason=fixed_gap_0.02")
                End If

                Return ValidateDimensionVisibleOnSheet(dimObj, sheet, draft, app, testName, lab)
            Catch ex As COMException
                lab.Log("PLACE", "FORCE_FAIL_COM", "name=" & testName & " hr=0x" & ex.ErrorCode.ToString("X8", CultureInfo.InvariantCulture) & " " & ex.Message)
                Return False
            Catch ex As Exception
                lab.Log("PLACE", "FORCE_FAIL", "name=" & testName & " " & ex.Message)
                Return False
            End Try
        End Function

        Private Shared Function TryCreateAuxVisibleDiagHorizontal(dims As Dimensions,
                                                                sheet As Sheet,
                                                                draft As DraftDocument,
                                                                view As DrawingView,
                                                                app As Application,
                                                                vx1 As Double,
                                                                vy1 As Double,
                                                                vx2 As Double,
                                                                vy2 As Double,
                                                                testName As String,
                                                                lab As DimLabLogger,
                                                                created As List(Of Object),
                                                                preserve As List(Of Object)) As Boolean
            lab.Log("AUX_VISIBLE_DIAG", "CREATE", "forTest=" & testName & " viewSeg=(" & Gd(vx1) & "," & Gd(vy1) & ")-(" & Gd(vx2) & "," & Gd(vy2) & ")")
            Dim sx1 As Double, sy1 As Double, sx2 As Double, sy2 As Double
            Try
                view.ViewToSheet(vx1, vy1, sx1, sy1)
                view.ViewToSheet(vx2, vy2, sx2, sy2)
            Catch ex As Exception
                lab.Log("AUX_VISIBLE_DIAG", "FAIL", "ViewToSheet " & ex.Message)
                Return False
            End Try
            Dim aux As Object = Nothing
            Try
                aux = sheet.Lines2d.AddBy2Points(sx1, sy1, sx2, sy2)
            Catch ex As Exception
                lab.Log("AUX_VISIBLE_DIAG", "FAIL", "AddBy2Points " & ex.Message)
                Return False
            End Try
            If aux Is Nothing Then Return False
            created.Add(aux)
            preserve.Add(aux)
            Dim d As Dimension = Nothing
            Try
                Dim r As Object = aux
                Try : r = CallByName(aux, "Reference", CallType.Get) : Catch : End Try
                d = TryCast(CallByName(dims, "AddLength", CallType.Method, r), Dimension)
            Catch ex As Exception
                lab.Log("AUX_VISIBLE_DIAG", "FAIL", "AddLength " & ex.Message)
                Return False
            End Try
            If d Is Nothing Then Return False
            created.Add(d)
            preserve.Add(d)
            Try
                d.TrackDistance = 0.012R
            Catch ex As Exception
                lab.Log("AUX_VISIBLE_DIAG", "TRACK_SKIP", ex.Message)
            End Try
            Try
                CallByName(d, "AbsoluteTrackDistance", CallType.Let, 0.012R)
            Catch
            End Try
            Try : draft.UpdateAll(True) : Catch : End Try
            Dim v = ReadDimValue(d)
            lab.Log("AUX_VISIBLE_DIAG", "OK", "value=" & Gd(v))
            Dim vis = ValidateDimensionVisibleOnSheet(d, sheet, draft, app, testName & "_AuxVisibleDiag", lab)
            lab.Log("AUX_VISIBLE_DIAG", "VISIBLE", vis.ToString(CultureInfo.InvariantCulture))
            Return vis
        End Function

        Private Shared Function IsIsoView(dv As DrawingView) As Boolean
            Try
                Dim ori As SolidEdgeConstants.ViewOrientationConstants = SolidEdgeConstants.ViewOrientationConstants.igFrontView
                Dim vx As Double, vy As Double, vz As Double, lx As Double, ly As Double, lz As Double
                dv.ViewOrientation(vx, vy, vz, lx, ly, lz, ori)
                If ori = SolidEdgeConstants.ViewOrientationConstants.igTopFrontRightView Then Return True
            Catch
            End Try
            Try
                Dim dvt As String = Convert.ToString(dv.DrawingViewType, CultureInfo.InvariantCulture).ToLowerInvariant()
                If dvt.Contains("iso") Then Return True
            Catch
            End Try
            Return False
        End Function

        Private Shared Function FindBestDrawingView(sh As Sheet, lab As DimLabLogger) As DrawingView
            Dim cnt As Integer = 0
            Try : cnt = sh.DrawingViews.Count : Catch : Return Nothing : End Try
            For i As Integer = 1 To cnt
                Dim dv As DrawingView = Nothing
                Try
                    dv = CType(sh.DrawingViews.Item(i), DrawingView)
                Catch
                    Continue For
                End Try
                If dv Is Nothing Then Continue For
                Dim nm As String = "", sf As String = "?"
                Dim nL As Integer = 0, nA As Integer = 0, nC As Integer = 0
                Dim rx1 As Double, ry1 As Double, rx2 As Double, ry2 As Double
                Try : nm = dv.Name : Catch : End Try
                Try : sf = dv.ScaleFactor.ToString(CultureInfo.InvariantCulture) : Catch : End Try
                Try : nL = dv.DVLines2d.Count : Catch : End Try
                Try : nA = dv.DVArcs2d.Count : Catch : End Try
                Try : nC = dv.DVCircles2d.Count : Catch : End Try
                Dim rng As String = "n/a"
                Try
                    dv.Range(rx1, ry1, rx2, ry2)
                    rng = rx1.ToString("G6", CultureInfo.InvariantCulture) & "," & ry1.ToString("G6", CultureInfo.InvariantCulture) &
                          "->" & rx2.ToString("G6", CultureInfo.InvariantCulture) & "," & ry2.ToString("G6", CultureInfo.InvariantCulture)
                Catch
                End Try
                lab.Log("VIEW", "SCAN", "idx=" & i.ToString(CultureInfo.InvariantCulture) & " Name=" & nm & " ScaleFactor=" & sf &
                        " Range=" & rng & " DVLines2d.Count=" & nL.ToString(CultureInfo.InvariantCulture) &
                        " DVArcs2d.Count=" & nA.ToString(CultureInfo.InvariantCulture) &
                        " DVCircles2d.Count=" & nC.ToString(CultureInfo.InvariantCulture))
            Next

            For i As Integer = 1 To cnt
                Dim dv As DrawingView = Nothing
                Try
                    dv = CType(sh.DrawingViews.Item(i), DrawingView)
                Catch
                    Continue For
                End Try
                If dv Is Nothing Then Continue For
                If IsIsoView(dv) Then
                    lab.Log("VIEW", "SKIP", "idx=" & i.ToString(CultureInfo.InvariantCulture) & " reason=isometric")
                    Continue For
                End If
                Dim nL As Integer = 0
                Try : nL = dv.DVLines2d.Count : Catch : End Try
                If nL > 0 Then
                    lab.Log("VIEW", "PICK", "idx=" & i.ToString(CultureInfo.InvariantCulture) & " first_orthogonal_with_DVLines")
                    Return dv
                End If
            Next

            For i As Integer = 1 To cnt
                Dim dv As DrawingView = Nothing
                Try
                    dv = CType(sh.DrawingViews.Item(i), DrawingView)
                Catch
                    Continue For
                End Try
                If dv Is Nothing Then Continue For
                Dim nL2 As Integer = 0
                Try : nL2 = dv.DVLines2d.Count : Catch : End Try
                If nL2 > 0 Then
                    lab.Log("VIEW", "PICK", "idx=" & i.ToString(CultureInfo.InvariantCulture) & " fallback_first_with_DVLines_no_isometric_filter")
                    Return dv
                End If
            Next
            Return Nothing
        End Function

        Private Shared Sub LogViewAbortContext(draft As DraftDocument, resolvedSh As Sheet, lab As DimLabLogger)
            Dim rName = SafeSheetName(resolvedSh)
            Dim rCnt = SafeDrawingViewsCountSheet(resolvedSh)
            lab.Log("VIEW", "COUNT_ON_RESOLVED_SHEET", "count=" & rCnt.ToString(CultureInfo.InvariantCulture))
            Dim aName = SafeSheetNameFromDraft(draft)
            Dim aCnt = -1
            Try
                aCnt = SafeDrawingViewsCountSheet(CType(draft.ActiveSheet, Sheet))
            Catch
            End Try
            lab.Log("VIEW", "COUNT_ON_ACTIVE_SHEET", "count=" & aCnt.ToString(CultureInfo.InvariantCulture))
            lab.Log("ACTIVE_CONTEXT", "resolvedSheetName", rName)
            lab.Log("ACTIVE_CONTEXT", "activeSheetName", aName)
        End Sub

        Private Shared Function FindDrawingViewForLab(sh As Sheet, draft As DraftDocument, lab As DimLabLogger) As DrawingView
            Dim dv = FindBestDrawingView(sh, lab)
            If dv Is Nothing Then
                LogViewAbortContext(draft, sh, lab)
            End If
            Return dv
        End Function

        Private Shared Sub DumpDVLines(view As DrawingView, lab As DimLabLogger, session As LabLineReadSession)
            Dim n As Integer = 0
            Try : n = view.DVLines2d.Count : Catch : Return : End Try
            For i As Integer = 1 To n
                Dim ln As Object = Nothing
                Try : ln = view.DVLines2d.Item(i) : Catch : Continue For : End Try
                If ln Is Nothing Then Continue For
                Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
                Dim readOk = TryReadDVLineEndpointsBest(ln, x1, y1, x2, y2, lab, i, session, doLog:=True)
                Dim len As Double = -1, kpc As Integer = -1
                Dim refN As Boolean = True, mmN As Boolean = True
                Dim edgeT As String = "?"
                Try : len = CDbl(CallByName(ln, "Length", CallType.Get)) : Catch : End Try
                Try : kpc = CInt(CallByName(ln, "KeyPointCount", CallType.Get)) : Catch : End Try
                Try : refN = (CallByName(ln, "Reference", CallType.Get) Is Nothing) : Catch : refN = True : End Try
                Try : mmN = (CallByName(ln, "ModelMember", CallType.Get) Is Nothing) : Catch : mmN = True : End Try
                Try : edgeT = Convert.ToString(CallByName(ln, "EdgeType", CallType.Get), CultureInfo.InvariantCulture) : Catch : edgeT = "?" : End Try
                lab.Log("DVLINE", "DUMP", "index=" & i.ToString(CultureInfo.InvariantCulture) &
                        " readOk=" & readOk.ToString(CultureInfo.InvariantCulture) &
                        " start=(" & x1.ToString("G17", CultureInfo.InvariantCulture) & "," & y1.ToString("G17", CultureInfo.InvariantCulture) & ")" &
                        " end=(" & x2.ToString("G17", CultureInfo.InvariantCulture) & "," & y2.ToString("G17", CultureInfo.InvariantCulture) & ")" &
                        " length=" & len.ToString("G17", CultureInfo.InvariantCulture) &
                        " KeyPointCount=" & kpc.ToString(CultureInfo.InvariantCulture) &
                        " ReferenceIsNothing=" & refN.ToString(CultureInfo.InvariantCulture) &
                        " ModelMemberIsNothing=" & mmN.ToString(CultureInfo.InvariantCulture) &
                        " EdgeType=" & edgeT)
            Next
        End Sub

        Private Shared Function RunRetrieveDimensionsTest(dv As DrawingView, draft As DraftDocument, lab As DimLabLogger) As String
            Dim styles As String() = {"U3,5", "U2,5", "ISO (mm)", ""}
            For Each st In styles
                Try
                    lab.Log("RETRIEVE", "TRY", "style=" & If(String.IsNullOrEmpty(st), "(active)", st))
                    CallByName(dv, "RetrieveDimensions", CallType.Method, st)
                    Try : draft.UpdateAll(True) : Catch : End Try
                    lab.Log("RETRIEVE", "RESULT", "style=" & If(String.IsNullOrEmpty(st), "(active)", st) & " OK")
                    Return "SUCCESS"
                Catch ex As COMException
                    If (ex.ErrorCode = &H80004004) OrElse ex.Message.IndexOf("abort", StringComparison.OrdinalIgnoreCase) >= 0 Then
                        lab.Log("RETRIEVE", "RESULT", "E_ABORT style=" & st & " hr=0x" & ex.ErrorCode.ToString("X8", CultureInfo.InvariantCulture))
                    Else
                        lab.Log("RETRIEVE", "RESULT", "FAIL style=" & st & " " & ex.Message)
                    End If
                Catch ex As Exception
                    lab.Log("RETRIEVE", "RESULT", "FAIL style=" & st & " " & ex.Message)
                End Try
            Next
            Return "E_ABORT_OR_FAIL"
        End Function

        Private Shared Function BuildDVLineEndpointBounds(view As DrawingView, lab As DimLabLogger, session As LabLineReadSession,
                                                          Optional quiet As Boolean = False) As LineBounds
            Dim minX = Double.PositiveInfinity, maxX = Double.NegativeInfinity
            Dim minY = Double.PositiveInfinity, maxY = Double.NegativeInfinity
            Dim n As Integer = 0
            Try : n = view.DVLines2d.Count : Catch : Return Nothing : End Try
            Dim any As Boolean = False
            Dim readerFault As Boolean = False
            For i As Integer = 1 To n
                Dim ln As Object = Nothing
                Try : ln = view.DVLines2d.Item(i) : Catch : Continue For : End Try
                If ln Is Nothing Then Continue For
                Dim lnLen As Double = -1
                Try : lnLen = CDbl(CallByName(ln, "Length", CallType.Get)) : Catch : End Try
                Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
                If Not TryReadDVLineEndpointsBest(ln, x1, y1, x2, y2, lab, i, session, doLog:=False) Then
                    If lnLen > 0.000001R Then readerFault = True
                    Continue For
                End If
                any = True
                minX = Math.Min(minX, Math.Min(x1, x2))
                maxX = Math.Max(maxX, Math.Max(x1, x2))
                minY = Math.Min(minY, Math.Min(y1, y2))
                maxY = Math.Max(maxY, Math.Max(y1, y2))
                Dim refN As Boolean = True
                Try : refN = (CallByName(ln, "Reference", CallType.Get) Is Nothing) : Catch : refN = True : End Try
                Dim et As String = "?"
                Try : et = Convert.ToString(CallByName(ln, "EdgeType", CallType.Get), CultureInfo.InvariantCulture) : Catch : End Try
                Dim kpc As Integer = -1
                Try : kpc = CInt(CallByName(ln, "KeyPointCount", CallType.Get)) : Catch : End Try
                If Not quiet Then
                    lab.Log("DVREF", "BOUND_SCAN", "idx=" & i.ToString(CultureInfo.InvariantCulture) &
                            " start=(" & x1.ToString("G17", CultureInfo.InvariantCulture) & "," & y1.ToString("G17", CultureInfo.InvariantCulture) & ")" &
                            " end=(" & x2.ToString("G17", CultureInfo.InvariantCulture) & "," & y2.ToString("G17", CultureInfo.InvariantCulture) & ")" &
                            " length=" & lnLen.ToString("G17", CultureInfo.InvariantCulture) &
                            " refNothing=" & refN.ToString(CultureInfo.InvariantCulture) &
                            " edgeType=" & et & " keyPointCount=" & kpc.ToString(CultureInfo.InvariantCulture))
                End If
            Next
            If Not any Then
                If Not quiet Then lab.Log("DVREF", "ABORT", "reason=endpoint_reader_failed no_valid_endpoint_reads")
                Return Nothing
            End If
            Dim b As New LineBounds With {
                .MinX = minX, .MaxX = maxX, .MinY = minY, .MaxY = maxY,
                .ExpectedWidth = maxX - minX,
                .ExpectedHeight = maxY - minY
            }
            If Not quiet Then
                lab.Log("DVREF", "BOUNDS", "minX=" & b.MinX.ToString("G17", CultureInfo.InvariantCulture) &
                        " maxX=" & b.MaxX.ToString("G17", CultureInfo.InvariantCulture) &
                        " minY=" & b.MinY.ToString("G17", CultureInfo.InvariantCulture) &
                        " maxY=" & b.MaxY.ToString("G17", CultureInfo.InvariantCulture))
                lab.Log("DVREF", "EXPECTED_TOTALS", "width=" & b.ExpectedWidth.ToString("G17", CultureInfo.InvariantCulture) &
                        " height=" & b.ExpectedHeight.ToString("G17", CultureInfo.InvariantCulture))
            End If
            If b.ExpectedWidth < 0.000001R OrElse b.ExpectedHeight < 0.000001R Then
                If Not quiet Then
                    If readerFault Then
                        lab.Log("DVREF", "ABORT", "reason=endpoint_reader_failed width=" & b.ExpectedWidth.ToString("G17", CultureInfo.InvariantCulture) &
                                " height=" & b.ExpectedHeight.ToString("G17", CultureInfo.InvariantCulture))
                    Else
                        lab.Log("DVREF", "ABORT", "reason=degenerate_bounds_check_endpoint_reader width=" & b.ExpectedWidth.ToString("G17", CultureInfo.InvariantCulture) &
                                " height=" & b.ExpectedHeight.ToString("G17", CultureInfo.InvariantCulture))
                    End If
                End If
                Return Nothing
            End If
            Return b
        End Function

        Private Shared Sub BuildEndpointCandidates(view As DrawingView, candidates As List(Of DVEndpointCandidate), lab As DimLabLogger, session As LabLineReadSession)
            Dim n As Integer = 0
            Try : n = view.DVLines2d.Count : Catch : Exit Sub : End Try
            For i As Integer = 1 To n
                Dim ln As Object = Nothing
                Try : ln = view.DVLines2d.Item(i) : Catch : Continue For : End Try
                If ln Is Nothing Then Continue For
                Dim refObj As Object = Nothing
                Try : refObj = CallByName(ln, "Reference", CallType.Get) : Catch : End Try
                If refObj Is Nothing Then Continue For
                Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
                If Not TryReadDVLineEndpointsBest(ln, x1, y1, x2, y2, lab, i, session, doLog:=False) Then Continue For
                Dim lnLen As Double = 0
                Try : lnLen = CDbl(CallByName(ln, "Length", CallType.Get)) : Catch : End Try
                candidates.Add(New DVEndpointCandidate With {.LineIndex = i, .Line = ln, .Ref = refObj, .X = x1, .Y = y1, .IsStart = True, .Length = lnLen})
                candidates.Add(New DVEndpointCandidate With {.LineIndex = i, .Line = ln, .Ref = refObj, .X = x2, .Y = y2, .IsStart = False, .Length = lnLen})
            Next
        End Sub

        Private Shared Sub ClassifyBoundaries(cands As List(Of DVEndpointCandidate), b As LineBounds, tol As Double,
                                              leftL As List(Of DVEndpointCandidate), rightL As List(Of DVEndpointCandidate),
                                              bottomL As List(Of DVEndpointCandidate), topL As List(Of DVEndpointCandidate),
                                              lab As DimLabLogger,
                                              Optional quiet As Boolean = False)
            For Each c In cands
                If Math.Abs(c.X - b.MinX) <= tol Then
                    leftL.Add(c)
                    If Not quiet Then LogCand(lab, "LEFT", c)
                End If
                If Math.Abs(c.X - b.MaxX) <= tol Then
                    rightL.Add(c)
                    If Not quiet Then LogCand(lab, "RIGHT", c)
                End If
                If Math.Abs(c.Y - b.MinY) <= tol Then
                    bottomL.Add(c)
                    If Not quiet Then LogCand(lab, "BOTTOM", c)
                End If
                If Math.Abs(c.Y - b.MaxY) <= tol Then
                    topL.Add(c)
                    If Not quiet Then LogCand(lab, "TOP", c)
                End If
            Next
            If Not quiet Then
                lab.Log("DVREF", "BOUNDARY_LEFT", "count=" & leftL.Count.ToString(CultureInfo.InvariantCulture))
                lab.Log("DVREF", "BOUNDARY_RIGHT", "count=" & rightL.Count.ToString(CultureInfo.InvariantCulture))
                lab.Log("DVREF", "BOUNDARY_BOTTOM", "count=" & bottomL.Count.ToString(CultureInfo.InvariantCulture))
                lab.Log("DVREF", "BOUNDARY_TOP", "count=" & topL.Count.ToString(CultureInfo.InvariantCulture))
            End If
        End Sub

        Private Shared Sub LogCand(lab As DimLabLogger, side As String, c As DVEndpointCandidate)
            lab.Log("DVREF", "BOUNDARY_CAND", "side=" & side & " idx=" & c.LineIndex.ToString(CultureInfo.InvariantCulture) &
                    " point=(" & c.X.ToString("G17", CultureInfo.InvariantCulture) & "," & c.Y.ToString("G17", CultureInfo.InvariantCulture) & ")" &
                    " length=" & c.Length.ToString("G17", CultureInfo.InvariantCulture) & " refNothing=False")
        End Sub

        Private Shared Function RunHorizontalTotalTests(dims As Dimensions, dv As DrawingView, draft As DraftDocument,
                                                        b As LineBounds, leftL As List(Of DVEndpointCandidate), rightL As List(Of DVEndpointCandidate),
                                                        created As List(Of Object), lab As DimLabLogger, ctx As LabRunContext) As String
            Dim best As String = "FAIL"
            If leftL.Count = 0 OrElse rightL.Count = 0 Then Return best
            Dim p1 = PickClosestY(leftL, b.MaxY)
            Dim p2 = PickClosestY(rightL, p1.Y)
            best = MaxResult(best, TryOneDistance(dims, dv, draft, "HorizontalTotal_TopPair", p1, p2, b.ExpectedWidth, True, True, created, lab, ctx))
            p1 = PickClosestY(leftL, b.MinY)
            p2 = PickClosestY(rightL, p1.Y)
            best = MaxResult(best, TryOneDistance(dims, dv, draft, "HorizontalTotal_BottomPair", p1, p2, b.ExpectedWidth, True, True, created, lab, ctx))
            best = MaxResult(best, TryHorizontalBestSameY(dims, dv, draft, b, leftL, rightL, created, lab, ctx))
            Return best
        End Function

        Private Shared Function PickClosestY(lst As List(Of DVEndpointCandidate), y As Double) As DVEndpointCandidate
            Dim best = lst(0)
            Dim bestD = Double.MaxValue
            For Each c In lst
                Dim d = Math.Abs(c.Y - y)
                If d < bestD Then bestD = d : best = c
            Next
            Return best
        End Function

        Private Shared Function PickClosestX(lst As List(Of DVEndpointCandidate), x As Double) As DVEndpointCandidate
            Dim best = lst(0)
            Dim bestD = Double.MaxValue
            For Each c In lst
                Dim d = Math.Abs(c.X - x)
                If d < bestD Then bestD = d : best = c
            Next
            Return best
        End Function

        Private Shared Function TryHorizontalBestSameY(dims As Dimensions, dv As DrawingView, draft As DraftDocument,
                                                       b As LineBounds, leftL As List(Of DVEndpointCandidate), rightL As List(Of DVEndpointCandidate),
                                                       created As List(Of Object), lab As DimLabLogger, ctx As LabRunContext) As String
            Dim bestA As DVEndpointCandidate = Nothing
            Dim bestC As DVEndpointCandidate = Nothing
            Dim bestDy = Double.MaxValue
            For Each a In leftL
                For Each c In rightL
                    Dim dy = Math.Abs(a.Y - c.Y)
                    If dy < bestDy Then bestDy = dy : bestA = a : bestC = c
                Next
            Next
            If bestA Is Nothing OrElse bestC Is Nothing Then Return "FAIL"
            Return TryOneDistance(dims, dv, draft, "HorizontalTotal_BestSameY", bestA, bestC, b.ExpectedWidth, True, True, created, lab, ctx)
        End Function

        Private Shared Function ResultScore(s As String) As Integer
            If s Is Nothing Then Return 0
            If s.IndexOf("SUCCESS_CONNECTED", StringComparison.OrdinalIgnoreCase) >= 0 Then Return 5
            If s.IndexOf("SUCCESS_VISIBLE", StringComparison.OrdinalIgnoreCase) >= 0 Then Return 4
            If s.IndexOf("SUCCESS", StringComparison.OrdinalIgnoreCase) >= 0 Then Return 3
            If s.IndexOf("WRONG_VALUE_DIAGONAL", StringComparison.OrdinalIgnoreCase) >= 0 Then Return 2
            If s.IndexOf("WRONG_VALUE", StringComparison.OrdinalIgnoreCase) >= 0 Then Return 1
            Return 0
        End Function

        Private Shared Function MaxResult(a As String, b As String) As String
            Return If(ResultScore(b) > ResultScore(a), b, a)
        End Function

        Private Shared Function RunVerticalTotalTests(dims As Dimensions, dv As DrawingView, draft As DraftDocument,
                                                      b As LineBounds, bottomL As List(Of DVEndpointCandidate), topL As List(Of DVEndpointCandidate),
                                                      view As DrawingView, created As List(Of Object), lab As DimLabLogger, readSession As LabLineReadSession, ctx As LabRunContext) As String
            Dim best As String = "FAIL"
            If bottomL.Count = 0 OrElse topL.Count = 0 Then Return best
            Dim pb = PickClosestX(bottomL, b.MinX)
            Dim pt = PickClosestX(topL, pb.X)
            best = MaxResult(best, TryOneDistance(dims, dv, draft, "VerticalTotal_LeftPair_Keypoints", pb, pt, b.ExpectedHeight, True, True, created, lab, ctx))
            pb = PickClosestX(bottomL, b.MaxX)
            pt = PickClosestX(topL, pb.X)
            best = MaxResult(best, TryOneDistance(dims, dv, draft, "VerticalTotal_RightPair_Keypoints", pb, pt, b.ExpectedHeight, True, True, created, lab, ctx))
            best = MaxResult(best, TryVerticalBestSameX(dims, dv, draft, b, bottomL, topL, created, lab, ctx))
            best = MaxResult(best, TryVerticalParallelObject(dims, dv, draft, b, view, "VerticalTotal_LeftObject_ParallelHorizontals", True, created, lab, readSession, ctx))
            best = MaxResult(best, TryVerticalParallelObject(dims, dv, draft, b, view, "VerticalTotal_RightObject_ParallelHorizontals", False, created, lab, readSession, ctx))
            best = MaxResult(best, TryVerticalParallelObjectBest(dims, dv, draft, b, view, created, lab, readSession, ctx))
            Return best
        End Function

        Private Shared Function TryVerticalBestSameX(dims As Dimensions, dv As DrawingView, draft As DraftDocument,
                                                     b As LineBounds, bottomL As List(Of DVEndpointCandidate), topL As List(Of DVEndpointCandidate),
                                                     created As List(Of Object), lab As DimLabLogger, ctx As LabRunContext) As String
            Dim bestA As DVEndpointCandidate = Nothing
            Dim bestC As DVEndpointCandidate = Nothing
            Dim bestDx = Double.MaxValue
            For Each a In bottomL
                For Each c In topL
                    Dim dx = Math.Abs(a.X - c.X)
                    If dx < bestDx Then bestDx = dx : bestA = a : bestC = c
                Next
            Next
            If bestA Is Nothing OrElse bestC Is Nothing Then Return "FAIL"
            Return TryOneDistance(dims, dv, draft, "VerticalTotal_BestSameX_Keypoints", bestA, bestC, b.ExpectedHeight, True, True, created, lab, ctx)
        End Function

        Private Shared Function IsHorizontalLine(ln As Object, span As Double, lab As DimLabLogger, session As LabLineReadSession, lineIdx As Integer) As Boolean
            Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
            If Not TryReadDVLineEndpointsBest(ln, x1, y1, x2, y2, lab, lineIdx, session, doLog:=False) Then Return False
            Dim dy = Math.Abs(y2 - y1)
            Dim dx = Math.Abs(x2 - x1)
            Dim tol = Math.Max(1.0E-12R, span * 1.0E-9R)
            Return dy <= tol OrElse (dx > 1.0E-12R AndAlso dy / dx < 0.01R)
        End Function

        Private Shared Function FindLongestHorizontalNearY(view As DrawingView, yTarget As Double, useMinY As Boolean, span As Double, lab As DimLabLogger, session As LabLineReadSession) As HorizLinePick
            Dim r As New HorizLinePick()
            Dim bestLen As Double = -1
            Dim n As Integer = 0
            Try : n = view.DVLines2d.Count : Catch : Return r : End Try
            For i As Integer = 1 To n
                Dim ln As Object = Nothing
                Try : ln = view.DVLines2d.Item(i) : Catch : Continue For : End Try
                If ln Is Nothing OrElse Not IsHorizontalLine(ln, span, lab, session, i) Then Continue For
                Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
                If Not TryReadDVLineEndpointsBest(ln, x1, y1, x2, y2, lab, i, session, doLog:=False) Then Continue For
                Dim ym = (y1 + y2) / 2.0R
                Dim dist = Math.Abs(ym - yTarget)
                Dim L = Math.Abs(x2 - x1)
                If L > bestLen AndAlso dist < span * 0.15R Then
                    bestLen = L
                    r.Line = ln
                    r.X1 = x1 : r.Y1 = y1 : r.X2 = x2 : r.Y2 = y2
                End If
            Next
            Return r
        End Function

        Private Shared Function TryVerticalParallelObject(dims As Dimensions, dv As DrawingView, draft As DraftDocument,
                                                          b As LineBounds, view As DrawingView, name As String, useLeft As Boolean,
                                                          created As List(Of Object), lab As DimLabLogger, readSession As LabLineReadSession, ctx As LabRunContext) As String
            Dim xAnchor = If(useLeft, b.MinX, b.MaxX)
            Dim spanV = Math.Max(b.ExpectedWidth, b.ExpectedHeight)
            Dim bot = FindLongestHorizontalNearY(view, b.MinY, True, spanV, lab, readSession)
            Dim top = FindLongestHorizontalNearY(view, b.MaxY, False, spanV, lab, readSession)
            If bot.Line Is Nothing OrElse top.Line Is Nothing Then Return "FAIL"
            Dim refB As Object = Nothing, refT As Object = Nothing
            Try : refB = CallByName(bot.Line, "Reference", CallType.Get) : Catch : End Try
            Try : refT = CallByName(top.Line, "Reference", CallType.Get) : Catch : End Try
            If refB Is Nothing OrElse refT Is Nothing Then Return "FAIL"
            Dim by = (bot.Y1 + bot.Y2) / 2.0R
            Dim ty = (top.Y1 + top.Y2) / 2.0R
            Dim bx = PickXNearAnchor(bot.X1, bot.X2, xAnchor)
            Dim tx = PickXNearAnchor(top.X1, top.X2, xAnchor)
            lab.Log("DVREF", "PAIR_SELECT", "name=" & name & " bottomMid=(" & bx.ToString("G17", CultureInfo.InvariantCulture) & "," & by.ToString("G17", CultureInfo.InvariantCulture) & ") topMid=(" & tx.ToString("G17", CultureInfo.InvariantCulture) & "," & ty.ToString("G17", CultureInfo.InvariantCulture) & ")")
            Return TryOneDistanceRaw(dims, dv, draft, name, refB, bx, by, refT, tx, ty, b.ExpectedHeight, False, False, created, lab, ctx)
        End Function

        Private Shared Function PickXNearAnchor(x1 As Double, x2 As Double, anchor As Double) As Double
            If Math.Abs(x1 - anchor) <= Math.Abs(x2 - anchor) Then Return x1
            Return x2
        End Function

        Private Shared Function TryVerticalParallelObjectBest(dims As Dimensions, dv As DrawingView, draft As DraftDocument,
                                                              b As LineBounds, view As DrawingView, created As List(Of Object), lab As DimLabLogger, readSession As LabLineReadSession, ctx As LabRunContext) As String
            Dim a = TryVerticalParallelObject(dims, dv, draft, b, view, "VerticalTotal_BestObject_ParallelHorizontals_L", True, created, lab, readSession, ctx)
            Dim b2 = TryVerticalParallelObject(dims, dv, draft, b, view, "VerticalTotal_BestObject_ParallelHorizontals_R", False, created, lab, readSession, ctx)
            Return MaxResult(a, b2)
        End Function

        Private Shared Function TryOneDistance(dims As Dimensions, dv As DrawingView, draft As DraftDocument, testName As String,
                                               p1 As DVEndpointCandidate, p2 As DVEndpointCandidate, expected As Double,
                                               keyPoint1 As Boolean, keyPoint2 As Boolean,
                                               created As List(Of Object), lab As DimLabLogger, ctx As LabRunContext,
                                               Optional quietLog As Boolean = False,
                                               Optional ByRef outCreatedDimension As Dimension = Nothing) As String
            If Not quietLog Then
                lab.Log("DVREF", "PAIR_SELECT", "name=" & testName & " p1LineIdx=" & p1.LineIndex.ToString(CultureInfo.InvariantCulture) &
                        " p1=(" & p1.X.ToString("G17", CultureInfo.InvariantCulture) & "," & p1.Y.ToString("G17", CultureInfo.InvariantCulture) & ")" &
                        " p2LineIdx=" & p2.LineIndex.ToString(CultureInfo.InvariantCulture) &
                        " p2=(" & p2.X.ToString("G17", CultureInfo.InvariantCulture) & "," & p2.Y.ToString("G17", CultureInfo.InvariantCulture) & ")" &
                        " sameYDelta=" & Math.Abs(p1.Y - p2.Y).ToString("G17", CultureInfo.InvariantCulture))
            End If
            Return TryOneDistanceRaw(dims, dv, draft, testName, p1.Ref, p1.X, p1.Y, p2.Ref, p2.X, p2.Y, expected, keyPoint1, keyPoint2, created, lab, ctx, quietLog, outCreatedDimension)
        End Function

        Private Shared Function TryOneDistanceRaw(dims As Dimensions, dv As DrawingView, draft As DraftDocument, testName As String,
                                                  r1 As Object, x1 As Double, y1 As Double, r2 As Object, x2 As Double, y2 As Double,
                                                  expected As Double, kp1 As Boolean, kp2 As Boolean,
                                                  created As List(Of Object), lab As DimLabLogger, ctx As LabRunContext,
                                                  Optional quietLog As Boolean = False,
                                                  Optional ByRef outCreatedDimension As Dimension = Nothing) As String
            outCreatedDimension = Nothing
            If ctx IsNot Nothing Then
                TryApplyStyleToDimensionsCollectionLab(dims, ctx.RequestedStyleName, ctx.ResolvedStyleObj, lab, "precreate_" & testName)
            End If
            Dim axis = If(Math.Abs(y2 - y1) >= Math.Abs(x2 - x1), "V", "H")
            If Not quietLog Then
                lab.Log("DVREF", "TRY", "name=" & testName & " axis=" & axis & " keyPoint=" & kp1.ToString(CultureInfo.InvariantCulture) & "/" & kp2.ToString(CultureInfo.InvariantCulture) & " expected=" & expected.ToString("G17", CultureInfo.InvariantCulture))
                lab.Log("DVREF", "TRY_POINTS", "p1=(" & x1.ToString("G17", CultureInfo.InvariantCulture) & "," & y1.ToString("G17", CultureInfo.InvariantCulture) & ") p2=(" & x2.ToString("G17", CultureInfo.InvariantCulture) & "," & y2.ToString("G17", CultureInfo.InvariantCulture) & ")")
                lab.Log("DVREF", "TRY_REFS", "ref1Nothing=" & (r1 Is Nothing).ToString(CultureInfo.InvariantCulture) & " ref2Nothing=" & (r2 Is Nothing).ToString(CultureInfo.InvariantCulture))
            End If

            Dim sh As Sheet = ctx.Sheet
            Dim cntBefore = SafeDimensionsCount(sh)
            If Not quietLog Then
                lab.Log("DIMCOUNT", "before", cntBefore.ToString(CultureInfo.InvariantCulture))
            End If

            Dim d As Dimension = Nothing
            Try
                d = dims.AddDistanceBetweenObjects(r1, x1, y1, 0.0R, kp1, r2, x2, y2, 0.0R, kp2)
            Catch ex As Exception
                lab.Log("DVREF", "FAIL", "name=" & testName & " " & ex.Message)
                Dim cntGhost = SafeDimensionsCount(sh)
                If cntGhost > cntBefore Then
                    lab.Log("GHOST_DIM", "DETECTED", "test=" & testName &
                            " before=" & cntBefore.ToString(CultureInfo.InvariantCulture) &
                            " after=" & cntGhost.ToString(CultureInfo.InvariantCulture))
                    Try
                        Dim lastIdx As Integer = cntGhost
                        If lastIdx >= 1 Then
                            Dim ghostD = TryCast(sh.Dimensions.Item(lastIdx), Dimension)
                            If ghostD IsNot Nothing Then
                                InspectDimension(ghostD, lab, testName & "_ghost_last")
                                If ctx IsNot Nothing AndAlso Not ctx.KeepFailedDimensions Then
                                    lab.Log("DELETE_FAILED_DIM", "TRY", "name=" & testName & "_ghost_last")
                                    If TryDeleteDimension(ghostD) Then lab.Log("DELETE_FAILED_DIM", "OK", "name=" & testName & "_ghost_last")
                                End If
                            End If
                        End If
                    Catch
                    End Try
                End If
                Return "FAIL"
            End Try
            If d Is Nothing Then
                lab.Log("DVREF", "FAIL", "name=" & testName & " null_dimension")
                Return "FAIL"
            End If
            created.Add(d)
            outCreatedDimension = d

            Dim cntAfter = SafeDimensionsCount(sh)
            If Not quietLog Then
                lab.Log("DIMCOUNT", "after", cntAfter.ToString(CultureInfo.InvariantCulture))
                lab.Log("DIMCOUNT", "delta", (cntAfter - cntBefore).ToString(CultureInfo.InvariantCulture))
            End If

            Dim rx0 As Double, ry0 As Double, rx3 As Double, ry3 As Double
            If TryReadDimensionRange(d, rx0, ry0, rx3, ry3) AndAlso DimensionRangeIsInvalid(rx0, ry0, rx3, ry3) Then
                If Not quietLog Then
                    lab.Log("DVREF", "CREATED_BUT_RANGE_INVALID", "name=" & testName & " range=" & FormatRangeStr(rx0, ry0, rx3, ry3))
                End If
            End If

            Dim altPlacement = ctx IsNot Nothing AndAlso ctx.EnableAltPlacementLog
            Dim visible = ForceDimensionVisibleInSheet(d, dv, sh, draft, ctx.App, testName, axis, lab, altPlacement)
            Dim visReason As String = If(visible, "in_sheet", GetVisibilityReasonQuiet(d, sh))
            LogDimensionPosition(testName & "_after_place", d, lab)
            Dim textCenterRes = TryCenterTextInsideDimension(d, draft, ctx.App, lab, testName)
            LogDimensionPosition(testName & "_after_text_center", d, lab)
            Dim styleOk = ApplyDimensionStyleStrict(d, ctx.ResolvedStyleObj, ctx.RequestedStyleName, lab, testName & "_postplace", ctx.IsCleanFullStrict)
            LogDimensionStyleDiagnostics(d, testName, lab)
            Dim finalStyle = ReadDimensionStyleName(d, lab)
            lab.Log("STYLE", "FINAL", "name=" & testName & " final=" & finalStyle)
            If testName.IndexOf("Horizontal", StringComparison.OrdinalIgnoreCase) >= 0 Then
                ctx.SummaryHorizontalStyleFinal = If(String.IsNullOrWhiteSpace(finalStyle), "FAIL", finalStyle)
                ctx.TextCenterHorizontalResult = textCenterRes
                lab.Log("TEXT_CENTER", "HORIZONTAL", "result=" & textCenterRes)
            ElseIf testName.IndexOf("Vertical", StringComparison.OrdinalIgnoreCase) >= 0 Then
                ctx.SummaryVerticalStyleFinal = If(String.IsNullOrWhiteSpace(finalStyle), "FAIL", finalStyle)
                ctx.TextCenterVerticalResult = textCenterRes
                lab.Log("TEXT_CENTER", "VERTICAL", "result=" & textCenterRes)
            End If
            If ctx.IsCleanFullStrict AndAlso Not styleOk Then
                Return "STYLE_FAIL"
            End If

            Try : draft.UpdateAll(True) : Catch : End Try
            Try
                If ctx.App IsNot Nothing Then ctx.App.DoIdle()
            Catch
            End Try

            Dim actual = ReadDimValue(d)
            Dim euc = Math.Sqrt((x2 - x1) * (x2 - x1) + (y2 - y1) * (y2 - y1))
            Dim tolOk = Math.Max(0.001R, Math.Abs(expected) * 0.005R)
            If Not quietLog Then
                lab.Log("DVREF", "OK", "value=" & actual.ToString("G17", CultureInfo.InvariantCulture))
                lab.Log("DVREF", "DELTA", "expected=" & expected.ToString("G17", CultureInfo.InvariantCulture) & " actual=" & actual.ToString("G17", CultureInfo.InvariantCulture) & " delta=" & Math.Abs(actual - expected).ToString("G17", CultureInfo.InvariantCulture))
            End If
            Dim ev = EvaluateDimensionForLab(d, expected, draft, dv, sh, ctx.App, lab, testName, ctx)
            Dim res As String = ev.Result
            If Math.Abs(actual - euc) <= tolOk AndAlso Math.Abs(actual - expected) > tolOk Then
                If Not quietLog Then
                    lab.Log("DVREF", "DIAGONAL_DETECTED", "true euclidean=" & euc.ToString("G17", CultureInfo.InvariantCulture) & " expected=" & expected.ToString("G17", CultureInfo.InvariantCulture) & " actual=" & actual.ToString("G17", CultureInfo.InvariantCulture))
                End If
                res = "WRONG_VALUE_DIAGONAL"
            End If
            If Not quietLog Then
                lab.Log("DVREF", "RESULT", res & " name=" & testName)
            End If

            Dim tdAfter = SafeGetTrackDistance(d)
            Dim rxa As Double, rya As Double, rxb As Double, ryb As Double
            Dim rangeStr = "unreadable"
            If TryReadDimensionRange(d, rxa, rya, rxb, ryb) Then rangeStr = FormatRangeStr(rxa, rya, rxb, ryb)
            If Not quietLog AndAlso res.IndexOf("SUCCESS", StringComparison.OrdinalIgnoreCase) >= 0 AndAlso Not visible Then
                lab.Log("DVREF", "RESULT_VISIBILITY", "CREATED_BUT_NOT_VISIBLE name=" & testName & " reason=" & visReason)
            End If

            If Not quietLog Then
                ctx.Vis.Record(testName, res, ev.VisibleOk, rangeStr, Gd(tdAfter), visReason)
            End If

            If Not quietLog AndAlso Not visible AndAlso res.IndexOf("SUCCESS", StringComparison.OrdinalIgnoreCase) >= 0 Then
                Dim doAux = testName.IndexOf("HorizontalTotal", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                    String.Equals(testName, "SmallGap_BetweenParallelHorizontals", StringComparison.OrdinalIgnoreCase)
                If doAux AndAlso String.Equals(axis, "H", StringComparison.OrdinalIgnoreCase) Then
                    TryCreateAuxVisibleDiagHorizontal(dims, sh, draft, dv, ctx.App, x1, y1, x2, y2, testName, lab, created, ctx.PreserveDiagnostics)
                End If
            End If

            If Not quietLog Then
                Try
                    InspectDimension(d, lab, testName)
                Catch ex As COMException
                    lab.Log("DIM", "INSPECT_ABORT_COM", testName & " hr=0x" & ex.ErrorCode.ToString("X8", CultureInfo.InvariantCulture) & " " & ex.Message)
                Catch ex As Exception
                    lab.Log("DIM", "INSPECT_ABORT", testName & " " & ex.Message)
                End Try
            End If

            Dim isFailed = res.IndexOf("SUCCESS", StringComparison.OrdinalIgnoreCase) < 0
            If isFailed AndAlso ctx IsNot Nothing AndAlso Not ctx.KeepFailedDimensions Then
                lab.Log("DELETE_FAILED_DIM", "TRY", "name=" & testName)
                If TryDeleteDimension(d) Then
                    lab.Log("DELETE_FAILED_DIM", "OK", "name=" & testName)
                End If
            End If
            Return res
        End Function

        Private Shared Function ReadDimValue(d As Dimension) As Double
            Try
                Return CDbl(d.Value)
            Catch
                Return Double.NaN
            End Try
        End Function

        Private Shared Function CheckConnectedBySelectSet(draft As DraftDocument, dv As DrawingView, d As Object, lab As DimLabLogger, tag As String, ctx As LabRunContext) As ConnectivityState
            Dim ss As Object = Nothing
            Try : ss = draft.SelectSet : Catch : Return ConnectivityState.Unknown : End Try
            Dim c0 = SafeSsCount(ss)
            Try
                dv.AddConnectedDimensionsToSelectSet()
            Catch ex As COMException
                lab.Log("SELECTSET", "ERR_COM", tag & " hr=0x" & ex.ErrorCode.ToString("X8", CultureInfo.InvariantCulture) & " " & ex.Message)
                lab.Log("CONNECTED_CHECK", "STATE", "result=Unknown reason=selectset_exception")
                Return ConnectivityState.Unknown
            Catch ex As Exception
                lab.Log("SELECTSET", "ERR", tag & " " & ex.Message)
                lab.Log("CONNECTED_CHECK", "STATE", "result=Unknown reason=selectset_exception")
                Return ConnectivityState.Unknown
            End Try
            Dim c1 = SafeSsCount(ss)
            lab.Log("SELECTSET", "DELTA", tag & " before=" & c0.ToString(CultureInfo.InvariantCulture) & " after=" & c1.ToString(CultureInfo.InvariantCulture))
            If c1 > c0 Then
                If ctx IsNot Nothing Then ctx.SeenConnectedBySelectSet.Add(tag)
                lab.Log("CONNECTED_CHECK", "STATE", "result=True reason=selectset_delta_positive")
                Return ConnectivityState.TrueState
            End If
            If ctx IsNot Nothing AndAlso ctx.SeenConnectedBySelectSet.Contains(tag) Then
                lab.Log("CONNECTED_CHECK", "STATE", "result=True reason=previously_detected")
                Return ConnectivityState.TrueState
            End If
            lab.Log("CONNECTED_CHECK", "STATE", "result=Unknown reason=delta_nonpositive")
            Return ConnectivityState.Unknown
        End Function

        Private Shared Function CheckConnected(draft As DraftDocument, dv As DrawingView, d As Object, lab As DimLabLogger, tag As String, ctx As LabRunContext) As Boolean
            Return CheckConnectedBySelectSet(draft, dv, d, lab, tag, ctx) = ConnectivityState.TrueState
        End Function

        Private Shared Function SafeSsCount(ss As Object) As Integer
            If ss Is Nothing Then Return -1
            Try
                Return CInt(CallByName(ss, "Count", CallType.Get))
            Catch
                Return -1
            End Try
        End Function

        Private Shared Sub InspectDimension(d As Object, lab As DimLabLogger, tag As String)
            If d Is Nothing Then Exit Sub
            Try : lab.Log("DIM", "INSPECT", tag & " Value=" & Convert.ToString(CallByName(d, "Value", CallType.Get), CultureInfo.InvariantCulture)) : Catch : End Try
            Try : lab.Log("DIM", "INSPECT", tag & " StatusOfDimension=" & Convert.ToString(CallByName(d, "StatusOfDimension", CallType.Get), CultureInfo.InvariantCulture)) : Catch : End Try
            Try : lab.Log("DIM", "INSPECT", tag & " DimensionType=" & Convert.ToString(CallByName(d, "DimensionType", CallType.Get), CultureInfo.InvariantCulture)) : Catch : End Try
            Try : lab.Log("DIM", "INSPECT", tag & " TrackDistance=" & Convert.ToString(CallByName(d, "TrackDistance", CallType.Get), CultureInfo.InvariantCulture)) : Catch : End Try
            Try : lab.Log("DIM", "INSPECT", tag & " AbsoluteTrackDistance=" & Convert.ToString(CallByName(d, "AbsoluteTrackDistance", CallType.Get), CultureInfo.InvariantCulture)) : Catch : End Try
            Try : lab.Log("DIM", "INSPECT", tag & " Layer=" & Convert.ToString(CallByName(d, "Layer", CallType.Get), CultureInfo.InvariantCulture)) : Catch : End Try
            Try
                Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
                CallByName(d, "Range", CallType.Method, x1, y1, x2, y2)
                lab.Log("DIM", "INSPECT", tag & " Range=" & x1.ToString("G9", CultureInfo.InvariantCulture) & "," & y1.ToString("G9", CultureInfo.InvariantCulture) & ".." & x2.ToString("G9", CultureInfo.InvariantCulture) & "," & y2.ToString("G9", CultureInfo.InvariantCulture))
            Catch
            End Try
            LogDimDisplayDataInspect(d, lab, tag)
        End Sub

        Private Shared Function RunSmallGapTest(dims As Dimensions, dv As DrawingView, draft As DraftDocument, created As List(Of Object), lab As DimLabLogger, readSession As LabLineReadSession, ctx As LabRunContext) As String
            Dim span = 1.0R
            Try
                Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
                dv.Range(x1, y1, x2, y2)
                span = Math.Max(Math.Abs(x2 - x1), Math.Abs(y2 - y1))
            Catch
            End Try
            Dim horiz As New List(Of HorizSeg)()
            Dim n As Integer = 0
            Try : n = dv.DVLines2d.Count : Catch : Return "FAIL" : End Try
            For i As Integer = 1 To n
                Dim ln As Object = Nothing
                Try : ln = dv.DVLines2d.Item(i) : Catch : Continue For : End Try
                If ln Is Nothing OrElse Not IsHorizontalLine(ln, span, lab, readSession, i) Then Continue For
                Dim a1 As Double, b1 As Double, a2 As Double, b2 As Double
                If Not TryReadDVLineEndpointsBest(ln, a1, b1, a2, b2, lab, i, readSession, doLog:=False) Then Continue For
                Dim L = Math.Abs(a2 - a1)
                If L < span * 0.1R Then Continue For
                Dim hs As New HorizSeg With {
                    .Ym = (b1 + b2) / 2.0R,
                    .Line = ln,
                    .X1 = a1, .Y1 = b1, .X2 = a2, .Y2 = b2,
                    .Length = L
                }
                horiz.Add(hs)
            Next
            If horiz.Count < 2 Then Return "FAIL"
            horiz.Sort(Function(u, v) u.Ym.CompareTo(v.Ym))
            Dim bestGap = Double.MaxValue
            Dim lo As HorizSeg = Nothing
            Dim hi As HorizSeg = Nothing
            For i = 0 To horiz.Count - 2
                Dim g = horiz(i + 1).Ym - horiz(i).Ym
                If g > 0.0002R AndAlso g < 0.02R AndAlso g < bestGap Then
                    bestGap = g
                    lo = horiz(i)
                    hi = horiz(i + 1)
                End If
            Next
            If lo Is Nothing OrElse hi Is Nothing Then Return "FAIL"
            Dim r1 As Object = Nothing, r2 As Object = Nothing
            Try : r1 = CallByName(lo.Line, "Reference", CallType.Get) : Catch : End Try
            Try : r2 = CallByName(hi.Line, "Reference", CallType.Get) : Catch : End Try
            Dim mx1 = (lo.X1 + lo.X2) / 2.0R
            Dim mx2 = (hi.X1 + hi.X2) / 2.0R
            lab.Log("DVREF", "PAIR_SELECT", "name=SmallGap_BetweenParallelHorizontals gapApprox=" & bestGap.ToString("G17", CultureInfo.InvariantCulture))
            Return TryOneDistanceRaw(dims, dv, draft, "SmallGap_BetweenParallelHorizontals", r1, mx1, lo.Ym, r2, mx2, hi.Ym, bestGap, False, True, created, lab, ctx)
        End Function

        Private Shared Function RunAuxiliaryLine2dFallbackTest(dims As Dimensions, sh As Sheet, draft As DraftDocument, view As DrawingView, created As List(Of Object), lab As DimLabLogger, readSession As LabLineReadSession) As String
            Dim span = 1.0R
            Dim n As Integer = 0
            Try : n = view.DVLines2d.Count : Catch : Return "FAIL" : End Try
            Dim pick As Object = Nothing
            Dim pickIdx As Integer = -1
            Dim vx1 As Double, vy1 As Double, vx2 As Double, vy2 As Double
            For i As Integer = 1 To n
                Dim ln As Object = Nothing
                Try : ln = view.DVLines2d.Item(i) : Catch : Continue For : End Try
                If ln Is Nothing OrElse Not IsHorizontalLine(ln, span, lab, readSession, i) Then Continue For
                If Not TryReadDVLineEndpointsBest(ln, vx1, vy1, vx2, vy2, lab, i, readSession, doLog:=False) Then Continue For
                If Math.Abs(vx2 - vx1) > 0.05R Then pick = ln : pickIdx = i : Exit For
            Next
            If pick Is Nothing OrElse pickIdx < 0 Then Return "FAIL"
            If Not TryReadDVLineEndpointsBest(pick, vx1, vy1, vx2, vy2, lab, pickIdx, readSession, doLog:=True) Then Return "FAIL"
            Dim sx1 As Double, sy1 As Double, sx2 As Double, sy2 As Double
            Try
                view.ViewToSheet(vx1, vy1, sx1, sy1)
                view.ViewToSheet(vx2, vy2, sx2, sy2)
            Catch ex As Exception
                lab.Log("AUX", "FAIL", "ViewToSheet " & ex.Message)
                Return "FAIL"
            End Try
            Dim aux As Object = Nothing
            Try
                aux = CallByName(sh.Lines2d, "AddBy2Points", CallType.Method, sx1, sy1, sx2, sy2)
            Catch ex As Exception
                lab.Log("AUX", "FAIL", "AddBy2Points " & ex.Message)
                Return "FAIL"
            End Try
            If aux Is Nothing Then Return "FAIL"
            created.Add(aux)
            Dim d As Object = Nothing
            Try
                Dim r As Object = aux
                Try : r = CallByName(aux, "Reference", CallType.Get) : Catch : End Try
                d = CallByName(dims, "AddLength", CallType.Method, r)
            Catch ex As Exception
                lab.Log("AUX", "FAIL", "AddLength " & ex.Message)
                Return "FAIL"
            End Try
            If d Is Nothing Then Return "FAIL"
            created.Add(d)
            Try : draft.UpdateAll(True) : Catch : End Try
            Dim vis = ReadDimValue(TryCast(d, Dimension))
            Dim sf As Double = 1.0R
            Try : sf = view.ScaleFactor : Catch : End Try
            Dim realFromVisible = If(sf > 1.0E-12R, vis / sf, vis)
            Dim dvLen As Double = 0
            Try : dvLen = CDbl(CallByName(pick, "Length", CallType.Get)) : Catch : End Try
            lab.Log("AUX", "VALUES", "visibleLength=" & vis.ToString("G17", CultureInfo.InvariantCulture) & " realFromVisible=" & realFromVisible.ToString("G17", CultureInfo.InvariantCulture) & " realFromDV=" & dvLen.ToString("G17", CultureInfo.InvariantCulture))
            lab.Log("AUX", "RESULT", "SUCCESS_VISIBLE_ONLY")
            Return "SUCCESS_VISIBLE_ONLY"
        End Function

    End Class

End Namespace
