Option Strict Off

Imports System
Imports System.Collections.Generic
Imports System.Globalization
Imports System.Linq
Imports SolidEdgeDraft
Imports SolidEdgeFrameworkSupport

Friend NotInheritable Class DrawingViewDimensionCreator
    Private Const ForcedStyleName As String = "U3,5"
    Private Const UneStrictMode As Boolean = True

    Private Sub New()
    End Sub

    Public Shared Function CreatePlannedDimensions(draft As DraftDocument,
                                                   sheet As Sheet,
                                                   viewInfo As DrawingViewGeometryInfo,
                                                   plans As List(Of PlannedDimension),
                                                   styleObj As Object,
                                                   signatures As HashSet(Of String),
                                                   log As DimensionLogger,
                                                   ByRef createdLinear As Integer,
                                                   ByRef createdRadial As Integer,
                                                   ByRef discarded As Integer,
                                                   ByRef errors As Integer) As Integer
        Dim created As Integer = 0
        If sheet Is Nothing OrElse plans Is Nothing OrElse plans.Count = 0 Then Return created

        Dim dims As Dimensions = Nothing
        Try
            dims = CType(sheet.Dimensions, Dimensions)
        Catch ex As Exception
            log?.ComFail("Sheet.Dimensions", "Sheet", ex)
            Return 0
        End Try
        If dims Is Nothing Then Return 0

        Dim effectiveStyleObj As Object = ResolveForcedStyleObject(draft, sheet, styleObj)

        For Each p In plans
            If signatures.Contains(p.Signature) Then
                discarded += 1
                Continue For
            End If
            signatures.Add(p.Signature)
            log?.LogLine("[DIM][CREATE][TRY] view=" & viewInfo.ViewIndex.ToString(CultureInfo.InvariantCulture) & " type=" & p.TypeTag)
            TryActivateTargetSheet(draft, sheet, log)

            Dim dimObj As Dimension = Nothing
            Try
                If p.Kind = PlannedDimensionKind.HorizontalTotal OrElse p.Kind = PlannedDimensionKind.VerticalTotal Then
                    dimObj = TryCreateLinearDimension(dims, viewInfo.View, p, effectiveStyleObj, log)
                ElseIf p.Kind = PlannedDimensionKind.LineLength Then
                    dimObj = TryCreateLineLengthDimension(dims, p, log)
                ElseIf p.Kind = PlannedDimensionKind.Diameter Then
                    dimObj = TryCast(CallByName(dims, "AddDiameter", CallType.Method, p.Obj1, p.P1X, p.P1Y, 0.0R, False), Dimension)
                ElseIf p.Kind = PlannedDimensionKind.Radial Then
                    dimObj = TryCast(CallByName(dims, "AddRadialDiameter", CallType.Method, p.Obj1, p.P1X, p.P1Y, 0.0R, False), Dimension)
                End If

                If dimObj Is Nothing Then
                    errors += 1
                    log?.LogLine("[DIM][CREATE][FAIL] view=" & viewInfo.ViewIndex.ToString(CultureInfo.InvariantCulture) & " type=" & p.TypeTag & " reason=null_dimension")
                    Continue For
                End If

                LogDimensionResidence(dimObj, sheet, viewInfo, p, log, "after_create")

                Try
                    TryApplyStyleToDimension(dimObj, effectiveStyleObj, log)
                Catch exStyle As Exception
                    log?.ComFail("Dimension.Style", "Dimension", exStyle)
                End Try

                RepositionCreatedDimension(dimObj, viewInfo, p, log)

                LogDimensionResidence(dimObj, sheet, viewInfo, p, log, "after_style")

                If IsDimensionOutsideSheetBounds(dimObj, sheet, log) Then
                    Try
                        dimObj.Delete()
                    Catch
                    End Try
                    discarded += 1
                    log?.LogLine("[DIM][CREATE][DROP] view=" & viewInfo.ViewIndex.ToString(CultureInfo.InvariantCulture) &
                                 " type=" & p.TypeTag & " reason=outside_template_bounds")
                    Continue For
                End If

                created += 1
                If p.Kind = PlannedDimensionKind.HorizontalTotal OrElse p.Kind = PlannedDimensionKind.VerticalTotal OrElse p.Kind = PlannedDimensionKind.LineLength Then
                    createdLinear += 1
                Else
                    createdRadial += 1
                End If

                ' Modo actual: mantener cota dentro de DrawingView para asegurar asociatividad.
                log?.LogLine("[DIM][DETACH][SKIP] mode=keep_in_view")
                log?.LogLine("[DIM][CREATE][OK] view=" & viewInfo.ViewIndex.ToString(CultureInfo.InvariantCulture) & " type=" & p.TypeTag)
            Catch ex As Exception
                errors += 1
                log?.ComFail("CreatePlannedDimensions", "Dimensions", ex)
            End Try
        Next

        Return created
    End Function

    Public Shared Function CreateDimensionsByViewSweep(
        draft As DraftDocument,
        sheet As Sheet,
        viewInfo As DrawingViewGeometryInfo,
        styleObj As Object,
        log As DimensionLogger,
        ByRef createdLinear As Integer,
        ByRef createdRadial As Integer,
        ByRef discarded As Integer,
        ByRef errors As Integer,
        Optional crossViewNominalKeys As HashSet(Of String) = Nothing,
        Optional norm As DimensioningNormConfig = Nothing
    ) As Integer
        Dim created As Integer = 0
        If sheet Is Nothing OrElse viewInfo Is Nothing OrElse viewInfo.View Is Nothing Then Return created

        Dim uneNorm As DimensioningNormConfig = If(norm, DimensioningNormConfig.DefaultConfig())
        Dim useCrossRegistry As Boolean = crossViewNominalKeys IsNot Nothing

        Dim dims As Dimensions = Nothing
        Try
            dims = CType(sheet.Dimensions, Dimensions)
        Catch ex As Exception
            log?.ComFail("Sheet.Dimensions", "Sheet", ex)
            Return 0
        End Try
        If dims Is Nothing Then Return 0

        Dim effectiveStyleObj As Object = ResolveForcedStyleObject(draft, sheet, styleObj)
        TryActivateTargetSheet(draft, sheet, log)
        TryApplyStyleToDimensionsCollection(dims, effectiveStyleObj, log, "VIEW_SWEEP")

        Dim centerX As Double = (viewInfo.Box.MinX + viewInfo.Box.MaxX) * 0.5R
        Dim centerY As Double = (viewInfo.Box.MinY + viewInfo.Box.MaxY) * 0.5R

        Dim entities As New List(Of SweepEntity)()
        AddSweepEntities(entities, viewInfo.View, "DVLines2d", "LINE", centerX, centerY)
        AddSweepEntities(entities, viewInfo.View, "DVArcs2d", "ARC", centerX, centerY)
        AddSweepEntities(entities, viewInfo.View, "DVCircles2d", "CIRCLE", centerX, centerY)
        AddSweepEntities(entities, viewInfo.View, "DVEllipses2d", "ELLIPSE", centerX, centerY)

        Dim ordered = entities.OrderBy(Function(e) e.DistanceToCenter).ToList()
        If UneStrictMode Then
            ordered = ordered.OrderBy(Function(e) e.Priority).ThenBy(Function(e) e.Band).ThenBy(Function(e) e.DistanceToCenter).ToList()
        End If
        log?.LogLine("[DIM][SWEEP][ORDER] view=" & viewInfo.ViewIndex.ToString(CultureInfo.InvariantCulture) &
                     " total=" & ordered.Count.ToString(CultureInfo.InvariantCulture) &
                     " strategy=inside_to_outside_by_band gap=small une_strict=" & UneStrictMode.ToString())

        Dim layoutScale As Double = 1.0R
        If uneNorm.EnableISO129Rules Then
            layoutScale = Math.Max(1.06R, Math.Min(1.42R, uneNorm.MinGapFromView / 0.012R))
            layoutScale = Math.Max(layoutScale, Math.Min(1.35R, 0.92R + uneNorm.GapBetweenDimensionRows / 0.008R * 0.05R))
        End If

        Dim baseGap As Double = If(UneStrictMode, 0.0012R, 0.0015R) * layoutScale
        Dim gapStep As Double = If(UneStrictMode, 0.00035R, 0.00045R) * layoutScale
        Dim dedupe As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Dim bandCounters As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
        Dim hLim As Integer = If(uneNorm.KeepIntentionalDuplicateDimensions, 80, 3)
        Dim vLim As Integer = If(uneNorm.KeepIntentionalDuplicateDimensions, 80, 3)
        Dim rLim As Integer = If(uneNorm.KeepIntentionalDuplicateDimensions, 40, 2)
        Dim bandMax As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase) From {
            {"H", hLim},
            {"V", vLim},
            {"R", rLim},
            {"GEN", 0},
            {"LINE", If(uneNorm.KeepIntentionalDuplicateDimensions, 40, 2)}
        }

        For i As Integer = 0 To ordered.Count - 1
            Dim it = ordered(i)
            Dim dimObj As Dimension = Nothing
            Try
                Dim nominalRegistryKey As String = Nothing
                If useCrossRegistry Then
                    nominalRegistryKey = TryComputeUneNominalRegistryKey(it, uneNorm)
                    If Not String.IsNullOrEmpty(nominalRegistryKey) AndAlso crossViewNominalKeys.Contains(nominalRegistryKey) Then
                        discarded += 1
                        log?.LogLine("[DIM][UNE129][CROSSVIEW][SKIP] view=" & viewInfo.ViewIndex.ToString(CultureInfo.InvariantCulture) &
                                     " kind=" & it.Kind & " key=" & nominalRegistryKey & " (solo si EnableDuplicateDimensionCleanup)")
                        Continue For
                    End If
                End If

                If UneStrictMode Then
                    Dim sig As String = BuildSweepEntitySignature(it.Obj, it.Kind)
                    If Not String.IsNullOrWhiteSpace(sig) Then
                        If dedupe.Contains(sig) Then
                            discarded += 1
                            log?.LogLine("[DIM][UNE][DEDUP][SKIP] view=" & viewInfo.ViewIndex.ToString(CultureInfo.InvariantCulture) &
                                         " kind=" & it.Kind & " sig=" & sig)
                            Continue For
                        End If
                        dedupe.Add(sig)
                    End If

                    If bandMax.ContainsKey(it.Band) Then
                        Dim current As Integer = If(bandCounters.ContainsKey(it.Band), bandCounters(it.Band), 0)
                        If current >= bandMax(it.Band) Then
                            discarded += 1
                            log?.LogLine("[DIM][UNE][BAND_LIMIT][SKIP] view=" & viewInfo.ViewIndex.ToString(CultureInfo.InvariantCulture) &
                                         " kind=" & it.Kind & " band=" & it.Band &
                                         " limit=" & bandMax(it.Band).ToString(CultureInfo.InvariantCulture))
                            Continue For
                        End If
                    End If
                End If

                Select Case it.Kind
                    Case "LINE"
                        dimObj = TryCreateByReferenceMethods(dims, it.Obj, New String() {"AddLength"}, log, "LINE")
                    Case "ARC", "CIRCLE"
                        dimObj = TryCreateByReferenceMethods(dims, it.Obj, New String() {"AddRadius", "AddRadialDiameter", "AddCircularDiameter", "AddLength"}, log, it.Kind)
                    Case "ELLIPSE"
                        dimObj = TryCreateByReferenceMethods(dims, it.Obj, New String() {"AddLength"}, log, it.Kind)
                End Select

                If dimObj Is Nothing Then
                    discarded += 1
                    Continue For
                End If

                TryApplyStyleToDimension(dimObj, effectiveStyleObj, log)

                If Not bandCounters.ContainsKey(it.Band) Then bandCounters(it.Band) = 0
                Dim bandIndex As Integer = bandCounters(it.Band)
                Dim td As Double = ComputeTrackDistanceForBand(viewInfo, it.Band, bandIndex, baseGap, gapStep, layoutScale)
                bandCounters(it.Band) = bandIndex + 1
                Try
                    CallByName(dimObj, "TrackDistance", CallType.Let, td)
                Catch
                End Try
                If String.Equals(it.Kind, "LINE", StringComparison.OrdinalIgnoreCase) Then
                    Try
                        CallByName(dimObj, "ProjectionLineDirection", CallType.Let, True)
                    Catch
                    End Try
                End If

                If UneStrictMode Then
                    log?.LogLine("[DIM][UNE][APPLY] view=" & viewInfo.ViewIndex.ToString(CultureInfo.InvariantCulture) &
                                 " kind=" & it.Kind &
                                 " band=" & it.Band &
                                 " orderInBand=" & bandIndex.ToString(CultureInfo.InvariantCulture) &
                                 " trackDistance=" & td.ToString("0.######", CultureInfo.InvariantCulture))
                End If

                If IsDimensionOutsideSheetBounds(dimObj, sheet, log) Then
                    Try : dimObj.Delete() : Catch : End Try
                    discarded += 1
                    log?.LogLine("[DIM][SWEEP][DROP] view=" & viewInfo.ViewIndex.ToString(CultureInfo.InvariantCulture) &
                                 " kind=" & it.Kind & " reason=outside_template_bounds")
                    Continue For
                End If

                If useCrossRegistry AndAlso Not String.IsNullOrEmpty(nominalRegistryKey) Then
                    crossViewNominalKeys.Add(nominalRegistryKey)
                End If

                created += 1
                If String.Equals(it.Kind, "LINE", StringComparison.OrdinalIgnoreCase) Then
                    createdLinear += 1
                Else
                    createdRadial += 1
                End If
            Catch ex As Exception
                errors += 1
                log?.LogLine("[DIM][SWEEP][ERR] view=" & viewInfo.ViewIndex.ToString(CultureInfo.InvariantCulture) &
                             " kind=" & it.Kind & " msg=" & ex.Message)
            End Try
        Next

        Return created
    End Function

    Private Shared Sub AddSweepEntities(
        ByVal list As List(Of SweepEntity),
        ByVal view As DrawingView,
        ByVal collectionName As String,
        ByVal kind As String,
        ByVal cx As Double,
        ByVal cy As Double
    )
        Dim col As Object = Nothing
        Try
            col = CallByName(view, collectionName, CallType.Get)
        Catch
            col = Nothing
        End Try
        If col Is Nothing Then Return

        For Each obj As Object In col
            Dim px As Double = 0.0R, py As Double = 0.0R
            Try
                Dim x1 As Double = 0, y1 As Double = 0, x2 As Double = 0, y2 As Double = 0
                CallByName(obj, "Range", CallType.Method, x1, y1, x2, y2)
                px = (Math.Min(x1, x2) + Math.Max(x1, x2)) * 0.5R
                py = (Math.Min(y1, y2) + Math.Max(y1, y2)) * 0.5R
            Catch
                Try
                    px = Convert.ToDouble(CallByName(obj, "X", CallType.Get), CultureInfo.InvariantCulture)
                    py = Convert.ToDouble(CallByName(obj, "Y", CallType.Get), CultureInfo.InvariantCulture)
                Catch
                    px = 0.0R : py = 0.0R
                End Try
            End Try

            list.Add(New SweepEntity With {
                .Obj = obj,
                .Kind = kind,
                .Priority = GetSweepPriority(kind),
                .Band = InferSweepBand(kind, obj),
                .DistanceToCenter = Math.Sqrt((px - cx) * (px - cx) + (py - cy) * (py - cy))
            })
        Next
    End Sub

    Private Shared Function GetSweepPriority(ByVal kind As String) As Integer
        If String.Equals(kind, "LINE", StringComparison.OrdinalIgnoreCase) Then Return 1
        If String.Equals(kind, "ARC", StringComparison.OrdinalIgnoreCase) Then Return 2
        If String.Equals(kind, "CIRCLE", StringComparison.OrdinalIgnoreCase) Then Return 2
        If String.Equals(kind, "ELLIPSE", StringComparison.OrdinalIgnoreCase) Then Return 3
        Return 9
    End Function

    Private Shared Function BuildSweepEntitySignature(ByVal obj As Object, ByVal kind As String) As String
        If obj Is Nothing Then Return ""
        Try
            Dim key As Object = CallByName(obj, "Key", CallType.Get)
            Dim ks As String = Convert.ToString(key, CultureInfo.InvariantCulture)
            If Not String.IsNullOrWhiteSpace(ks) Then Return kind & "|KEY|" & ks
        Catch
        End Try
        Try
            Dim rk As Object = CallByName(obj, "GetReferenceKey", CallType.Method)
            If rk IsNot Nothing Then
                If TypeOf rk Is Array Then
                    Dim arr = CType(rk, Array)
                    Return kind & "|REFKEY_LEN|" & arr.Length.ToString(CultureInfo.InvariantCulture)
                End If
                Return kind & "|REFKEY|" & Convert.ToString(rk, CultureInfo.InvariantCulture)
            End If
        Catch
        End Try
        Return ""
    End Function

    Private Shared Function InferSweepBand(ByVal kind As String, ByVal obj As Object) As String
        If obj Is Nothing Then Return "GEN"
        If String.Equals(kind, "LINE", StringComparison.OrdinalIgnoreCase) Then
            Try
                Dim x1 As Double = 0, y1 As Double = 0, x2 As Double = 0, y2 As Double = 0
                CallByName(obj, "Range", CallType.Method, x1, y1, x2, y2)
                Dim dx As Double = Math.Abs(x2 - x1)
                Dim dy As Double = Math.Abs(y2 - y1)
                If dx >= dy Then Return "H"
                Return "V"
            Catch
                Return "LINE"
            End Try
        End If
        If String.Equals(kind, "ARC", StringComparison.OrdinalIgnoreCase) OrElse String.Equals(kind, "CIRCLE", StringComparison.OrdinalIgnoreCase) Then
            Return "R"
        End If
        Return "GEN"
    End Function

    Private Shared Function TryComputeUneNominalRegistryKey(it As SweepEntity, norm As DimensioningNormConfig) As String
        If it Is Nothing OrElse norm Is Nothing Then Return Nothing
        Dim sep As Double = Math.Max(norm.MinFeatureSeparation, 1.0E-6R)
        If String.Equals(it.Kind, "LINE", StringComparison.OrdinalIgnoreCase) Then
            Dim lenM As Double
            If Not TryDvLineLengthMeters(it.Obj, lenM) Then Return Nothing
            If lenM < sep * 1.5R Then Return Nothing
            Dim bucket As Integer = CInt(Math.Round(lenM / sep))
            Return "L|" & it.Band & "|" & bucket.ToString(CultureInfo.InvariantCulture)
        End If
        If String.Equals(it.Kind, "ARC", StringComparison.OrdinalIgnoreCase) OrElse
           String.Equals(it.Kind, "CIRCLE", StringComparison.OrdinalIgnoreCase) Then
            Dim rM As Double
            If Not TryDvArcOrCircleRadiusMeters(it.Obj, rM) Then Return Nothing
            If rM < sep * 0.75R Then Return Nothing
            Dim bucket As Integer = CInt(Math.Round(rM / sep))
            Return "R|" & bucket.ToString(CultureInfo.InvariantCulture)
        End If
        Return Nothing
    End Function

    Private Shared Function TryDvLineLengthMeters(lineObj As Object, ByRef lenM As Double) As Boolean
        lenM = 0R
        If lineObj Is Nothing Then Return False
        Try
            Dim x1 As Double = 0, y1 As Double = 0, x2 As Double = 0, y2 As Double = 0
            CallByName(lineObj, "Range", CallType.Method, x1, y1, x2, y2)
            Dim dx As Double = x2 - x1
            Dim dy As Double = y2 - y1
            lenM = Math.Sqrt(dx * dx + dy * dy)
            Return lenM > 1.0E-9 AndAlso Not Double.IsNaN(lenM)
        Catch
            Return False
        End Try
    End Function

    Private Shared Function TryDvArcOrCircleRadiusMeters(obj As Object, ByRef rM As Double) As Boolean
        rM = 0R
        If obj Is Nothing Then Return False
        Try
            rM = Convert.ToDouble(CallByName(obj, "Radius", CallType.Get), CultureInfo.InvariantCulture)
            Return rM > 1.0E-9 AndAlso Not Double.IsNaN(rM)
        Catch
        End Try
        Return False
    End Function

    Private Shared Function ComputeTrackDistanceForBand(
        ByVal viewInfo As DrawingViewGeometryInfo,
        ByVal band As String,
        ByVal orderInBand As Integer,
        ByVal baseGap As Double,
        ByVal gapStep As Double,
        Optional ByVal layoutScale As Double = 1.0R
    ) As Double
        Dim scale As Double = 0.03R
        If viewInfo IsNot Nothing Then
            Dim refSize As Double = Math.Max(viewInfo.Box.Width, viewInfo.Box.Height)
            If refSize > 0.0R Then
                scale = refSize
            End If
        End If

        ' Distancia mínima respecto a pieza y separación entre cotas por banda.
        Dim bandBaseFactor As Double
        Dim bandStepFactor As Double
        Dim ls As Double = Math.Max(0.85R, Math.Min(1.6R, layoutScale))
        Select Case band
            Case "H"
                bandBaseFactor = 0.12R * ls
                bandStepFactor = 0.078R * ls
            Case "V"
                bandBaseFactor = 0.12R * ls
                bandStepFactor = 0.078R * ls
            Case "R"
                bandBaseFactor = 0.1R * ls
                bandStepFactor = 0.068R * ls
            Case Else
                bandBaseFactor = 0.085R * ls
                bandStepFactor = 0.055R * ls
        End Select

        Dim td As Double = (scale * bandBaseFactor) + (scale * bandStepFactor * orderInBand)

        ' Clamps para evitar extremos en vistas muy pequeñas o muy grandes.
        Dim minTd As Double = Math.Max(baseGap * 2.2R, 0.0038R * ls)
        Dim maxTd As Double = Math.Max(minTd, 0.034R * ls)
        If td < minTd Then td = minTd
        If td > maxTd Then td = maxTd
        Return td
    End Function

    Friend Shared Function TryCreateAddLengthOnReference(dims As Dimensions, dvLine2d As Object, log As DimensionLogger) As Dimension
        If dims Is Nothing OrElse dvLine2d Is Nothing Then Return Nothing
        Return TryCreateByReferenceMethods(dims, dvLine2d, New String() {"AddLength"}, log, "REF_LINE")
    End Function

    Friend Shared Function TryCreateRadiusOnReference(dims As Dimensions, arcOrCircle As Object, log As DimensionLogger) As Dimension
        If dims Is Nothing OrElse arcOrCircle Is Nothing Then Return Nothing
        Return TryCreateByReferenceMethods(dims, arcOrCircle, New String() {"AddRadius", "AddRadialDiameter", "AddCircularDiameter"}, log, "REF_RAD")
    End Function

    Private Shared Function TryCreateByReferenceMethods(
        ByVal dims As Dimensions,
        ByVal srcObj As Object,
        ByVal methodPriority As String(),
        ByVal log As DimensionLogger,
        ByVal kind As String
    ) As Dimension
        If dims Is Nothing OrElse srcObj Is Nothing Then Return Nothing

        Dim refObj As Object = Nothing
        Try
            refObj = CallByName(srcObj, "Reference", CallType.Get)
        Catch ex As Exception
            log?.LogLine("[DIM][SWEEP][REF][SKIP] kind=" & kind & " msg=" & ex.Message)
            Return Nothing
        End Try
        If refObj Is Nothing Then Return Nothing

        For Each methodName In methodPriority
            Try
                Dim d As Dimension = TryCast(CallByName(dims, methodName, CallType.Method, refObj), Dimension)
                If d IsNot Nothing Then
                    If kind.StartsWith("REF_", StringComparison.OrdinalIgnoreCase) Then
                        log?.LogLine("[DIM][CREATE][OK] kind=" & kind & " method=" & methodName)
                    Else
                        log?.LogLine("[DIM][SWEEP][CREATE][OK] kind=" & kind & " method=" & methodName)
                    End If
                    Return d
                End If
            Catch ex As Exception
                If kind.StartsWith("REF_", StringComparison.OrdinalIgnoreCase) Then
                    log?.LogLine("[DIM][CREATE][FAIL] kind=" & kind & " method=" & methodName & " msg=" & ex.Message)
                Else
                    log?.LogLine("[DIM][SWEEP][CREATE][FAIL] kind=" & kind & " method=" & methodName & " msg=" & ex.Message)
                End If
            End Try
        Next

        Return Nothing
    End Function

    Private Class SweepEntity
        Public Property Obj As Object
        Public Property Kind As String
        Public Property Priority As Integer
        Public Property Band As String
        Public Property DistanceToCenter As Double
    End Class

    Private Shared Function TryCreateLineLengthDimension(dims As Dimensions, p As PlannedDimension, log As DimensionLogger) As Dimension
        If dims Is Nothing OrElse p Is Nothing OrElse p.Obj1 Is Nothing Then Return Nothing
        Try
            Dim d As Dimension = TryCast(CallByName(dims, "AddLength", CallType.Method, p.Obj1), Dimension)
            If d IsNot Nothing Then
                log?.LogLine("[DIM][VIEW_CREATE][OK] mode=ADD_LENGTH type=" & p.TypeTag)
            End If
            Return d
        Catch ex As Exception
            log?.ComFail("TryCreateLineLengthDimension", "AddLength", ex)
            Return Nothing
        End Try
    End Function

    Private Shared Sub RepositionCreatedDimension(dimObj As Dimension, viewInfo As DrawingViewGeometryInfo, p As PlannedDimension, log As DimensionLogger)
        If dimObj Is Nothing OrElse viewInfo Is Nothing OrElse p Is Nothing Then Return
        Dim dx As Double = Math.Max(viewInfo.Box.Width * 0.04R, 0.003R)
        Dim dy As Double = Math.Max(viewInfo.Box.Height * 0.04R, 0.003R)

        Try
            Select Case p.Kind
                Case PlannedDimensionKind.HorizontalTotal
                    Dim xMid As Double = (p.P1X + p.P2X) * 0.5R
                    Dim yOut As Double = viewInfo.Box.MaxY + dy
                    If Not TrySetDimensionKeyPoint(dimObj, xMid, yOut) Then
                        Dim td As Double = Math.Max(viewInfo.Box.Height * 0.03R, 0.002R)
                        CallByName(dimObj, "TrackDistance", CallType.Let, td)
                    End If
                    log?.LogLine("[DIM][RELOCATE] type=H_TOTAL")

                Case PlannedDimensionKind.VerticalTotal
                    Dim yMid As Double = (p.P1Y + p.P2Y) * 0.5R
                    Dim xOut As Double = viewInfo.Box.MaxX + dx
                    If Not TrySetDimensionKeyPoint(dimObj, xOut, yMid) Then
                        Dim td As Double = Math.Max(viewInfo.Box.Width * 0.03R, 0.002R)
                        CallByName(dimObj, "TrackDistance", CallType.Let, td)
                    End If
                    log?.LogLine("[DIM][RELOCATE] type=V_TOTAL")

                Case PlannedDimensionKind.LineLength
                    ' Longitudes interiores: empuje mínimo para mantenerlas cerca de su línea.
                    Dim td As Double = Math.Max(Math.Min(viewInfo.Box.Height, viewInfo.Box.Width) * 0.015R, 0.0015R)
                    CallByName(dimObj, "TrackDistance", CallType.Let, td)
                    log?.LogLine("[DIM][RELOCATE] type=" & p.TypeTag & " mode=trackdistance_small")
            End Select
        Catch ex As Exception
            log?.LogLine("[DIM][RELOCATE][WARN] type=" & p.TypeTag & " msg=" & ex.Message)
        End Try
    End Sub

    Private Shared Function TrySetDimensionKeyPoint(dimObj As Dimension, x As Double, y As Double) As Boolean
        If dimObj Is Nothing Then Return False
        For Each kp As Integer In New Integer() {0, 1, 2}
            Try
                CallByName(dimObj, "SetKeyPoint", CallType.Method, kp, x, y, 0.0R)
                Return True
            Catch
            End Try
        Next
        Return False
    End Function

    Friend Shared Sub TryActivateTargetSheet(draft As DraftDocument, targetSheet As Sheet, log As DimensionLogger)
        If targetSheet Is Nothing Then Return
        Try
            CallByName(targetSheet, "Activate", CallType.Method)
        Catch
            Try
                If draft IsNot Nothing Then
                    Dim active As Object = CallByName(draft, "ActiveSheet", CallType.Get)
                    If active IsNot Nothing AndAlso (Not Object.ReferenceEquals(active, targetSheet)) Then
                        CallByName(targetSheet, "Activate", CallType.Method)
                    End If
                End If
            Catch
            End Try
        End Try
        Dim activeName As String = ""
        Try
            If draft IsNot Nothing Then activeName = SafeToString(CallByName(CallByName(draft, "ActiveSheet", CallType.Get), "Name", CallType.Get))
        Catch
            activeName = ""
        End Try
        log?.LogLine("[DIM][SHEET][ACTIVE] name=" & activeName)
    End Sub

    Private Shared Sub LogDimensionResidence(dimObj As Dimension,
                                             expectedSheet As Sheet,
                                             viewInfo As DrawingViewGeometryInfo,
                                             p As PlannedDimension,
                                             log As DimensionLogger,
                                             stage As String)
        If dimObj Is Nothing Then Return
        Dim parentType As String = ""
        Dim parentName As String = ""
        Dim expectedName As String = ""
        Try
            If expectedSheet IsNot Nothing Then expectedName = SafeToString(CallByName(expectedSheet, "Name", CallType.Get))
        Catch
            expectedName = ""
        End Try
        Try
            Dim parentObj As Object = CallByName(dimObj, "Parent", CallType.Get)
            If parentObj IsNot Nothing Then
                parentType = TypeName(parentObj)
                parentName = SafeToString(CallByName(parentObj, "Name", CallType.Get))
            End If
        Catch
        End Try
        log?.LogLine("[DIM][RESIDENCE] stage=" & stage &
                     " view=" & viewInfo.ViewIndex.ToString(CultureInfo.InvariantCulture) &
                     " type=" & p.TypeTag &
                     " expectedSheet=" & expectedName &
                     " parentType=" & parentType &
                     " parentName=" & parentName)
    End Sub

    Private Shared Function TryCreateLinearDimension(dims As Dimensions, dv As DrawingView, p As PlannedDimension, styleObj As Object, log As DimensionLogger) As Dimension
        TryApplyStyleToDimensionsCollection(dims, styleObj, log, "SHEET_DIMENSIONS")

        Dim d1 As Dimension = TryCreateDimensionFromObjectReference(dims, p.Obj1, log, "OBJ1")
        If d1 IsNot Nothing Then Return d1

        Dim d2 As Dimension = TryCreateDimensionFromObjectReference(dims, p.Obj2, log, "OBJ2")
        If d2 IsNot Nothing Then Return d2

        log?.LogLine("[DIM][VIEW_CREATE][FAIL] AddDistanceBetweenObjects deshabilitado y AddLength/AddRadius/AddRadialDiameter/AddCircularDiameter no disponibles.")
        Return Nothing
    End Function

    Private Shared Function TryCreateDimensionFromObjectReference(
        ByVal dims As Dimensions,
        ByVal srcObj As Object,
        ByVal log As DimensionLogger,
        ByVal srcTag As String
    ) As Dimension
        If dims Is Nothing OrElse srcObj Is Nothing Then Return Nothing

        Dim refObj As Object = Nothing
        Try
            refObj = CallByName(srcObj, "Reference", CallType.Get)
        Catch ex As Exception
            log?.LogLine("[DIM][VIEW_CREATE][REF][SKIP] src=" & srcTag & " no Reference: " & ex.Message)
            Return Nothing
        End Try

        If refObj Is Nothing Then
            log?.LogLine("[DIM][VIEW_CREATE][REF][SKIP] src=" & srcTag & " Reference=Nothing")
            Return Nothing
        End If

        Dim methods As String() = {"AddLength", "AddRadius", "AddRadialDiameter", "AddCircularDiameter"}
        For Each m In methods
            Try
                Dim d As Dimension = TryCast(CallByName(dims, m, CallType.Method, refObj), Dimension)
                If d IsNot Nothing Then
                    log?.LogLine("[DIM][VIEW_CREATE][OK] mode=" & m & "(Reference) src=" & srcTag)
                    Return d
                End If
            Catch ex As Exception
                log?.LogLine("[DIM][VIEW_CREATE][" & m & "][FAIL] src=" & srcTag & " msg=" & ex.Message)
            End Try
        Next

        Return Nothing
    End Function

    ''' <summary>Prueba aislada: aplica U3,5 reutilizando <see cref="ResolveForcedStyleObject"/> y <see cref="TryApplyStyleToDimension"/> (DimStyle, no String).</summary>
    Public Shared Sub ApplyU35ForIsolatedTest(
        draft As DraftDocument,
        sheet As Sheet,
        dimObj As Dimension,
        unitLog As Action(Of String)
    )
        If dimObj Is Nothing OrElse unitLog Is Nothing Then Return
        unitLog("[DIM][UNIT][STYLE][REUSE_EXISTING]")
        unitLog("[DIM][UNIT][STYLE][TRY] name=U3,5")
        Dim styleObj As Object = ResolveForcedStyleObject(draft, sheet, Nothing)
        Dim logger As New Logger(Sub(m As String) unitLog(m))
        Dim dimLog As New DimensionLogger(logger)
        Try
            TryApplyStyleToDimension(dimObj, styleObj, dimLog)
        Catch ex As Exception
            unitLog("[DIM][UNIT][STYLE][EX] " & ex.Message)
        End Try
        Dim applied As Boolean = IsDimensionStyleApplied(dimObj, ForcedStyleName)
        unitLog("[DIM][UNIT][STYLE][OK] applied=" & applied.ToString())
        unitLog("[DIM][UNIT][STYLE][READBACK] name=" & ReadDimensionStyleName(dimObj))
    End Sub

    Friend Shared Function ResolveForcedStyleObject(draft As DraftDocument, sheet As Sheet, currentStyleObj As Object) As Object
        If currentStyleObj IsNot Nothing Then Return currentStyleObj
        Dim styles As Object = Nothing
        Try
            If draft IsNot Nothing Then styles = CallByName(draft, "DimensionStyles", CallType.Get)
        Catch
            styles = Nothing
        End Try
        If styles Is Nothing Then
            Try
                If sheet IsNot Nothing Then styles = CallByName(sheet, "DimensionStyles", CallType.Get)
            Catch
                styles = Nothing
            End Try
        End If
        If styles Is Nothing Then Return Nothing
        Try
            Dim n As Integer = CInt(CallByName(styles, "Count", CallType.Get))
            For i As Integer = 1 To n
                Dim it As Object = CallByName(styles, "Item", CallType.Method, i)
                Dim nm As String = NormalizeStyleName(Convert.ToString(CallByName(it, "Name", CallType.Get), CultureInfo.InvariantCulture))
                If String.Equals(nm, NormalizeStyleName(ForcedStyleName), StringComparison.OrdinalIgnoreCase) Then
                    Return it
                End If
            Next
        Catch
        End Try
        Return Nothing
    End Function

    Friend Shared Sub TryApplyStyleToDimensionsCollection(dims As Dimensions, styleObj As Object, log As DimensionLogger, sourceTag As String)
        If dims Is Nothing Then Return
        Try
            CallByName(dims, "Style", CallType.Let, ForcedStyleName)
            log?.LogLine("[DIM][STYLE][COLLECTION] source=" & sourceTag & " mode=forced_name value=" & ForcedStyleName)
            Return
        Catch
        End Try

        Try
            ' Siguiendo patrón SDK: configurar estilo en la colección antes de crear.
            CallByName(dims, "Style", CallType.Let, styleObj)
            log?.LogLine("[DIM][STYLE][COLLECTION] source=" & sourceTag & " mode=object")
            Return
        Catch
        End Try

        Dim preferredName As String = ""
        Try
            If styleObj IsNot Nothing Then
                preferredName = Convert.ToString(CallByName(styleObj, "Name", CallType.Get), CultureInfo.InvariantCulture)
            End If
        Catch
            preferredName = ""
        End Try
        If String.IsNullOrWhiteSpace(preferredName) Then preferredName = ForcedStyleName

        For Each nameCandidate In BuildStyleNameCandidates(preferredName)
            Try
                CallByName(dims, "Style", CallType.Let, nameCandidate)
                log?.LogLine("[DIM][STYLE][COLLECTION] source=" & sourceTag & " mode=name value=" & nameCandidate)
                Exit For
            Catch
            End Try
        Next
    End Sub

    Private Shared Function TryGetReferenceToGraphicMember2(dv As DrawingView, graphicMember As Object, log As DimensionLogger) As Object
        If dv Is Nothing OrElse graphicMember Is Nothing Then Return Nothing
        Try
            Return CallByName(dv, "GetReferenceToGraphicMember2", CallType.Get, graphicMember)
        Catch ex As Exception
            log?.LogLine("[DIM][REF2][FAIL] " & ex.Message)
            Return Nothing
        End Try
    End Function

    Private Shared Function TryDetachFromDrawingViewIfNeeded(draft As DraftDocument,
                                                             sheet As Sheet,
                                                             viewInfo As DrawingViewGeometryInfo,
                                                             dimObj As Dimension,
                                                             p As PlannedDimension,
                                                             log As DimensionLogger) As Dimension
        Dim dv As DrawingView = If(viewInfo Is Nothing, Nothing, viewInfo.View)
        If draft Is Nothing OrElse sheet Is Nothing OrElse dv Is Nothing OrElse dimObj Is Nothing Then Return dimObj

        Dim connected As Boolean = IsDimensionConnectedToView(draft, sheet, dv, dimObj)
        log?.LogLine("[DIM][DETACH][TRY] connected=" & connected.ToString(CultureInfo.InvariantCulture))

        If connected Then
            Try
                ' Fase 2.1: forzar desplazamiento para sacar cota de la vista.
                Dim td0 As Double = SafeToDouble(CallByName(dimObj, "TrackDistance", CallType.Get))
                Dim delta As Double = Math.Max(viewInfo.Box.Width, viewInfo.Box.Height) * 0.6R
                If delta <= 0.0R Then delta = 0.02R
                CallByName(dimObj, "TrackDistance", CallType.Let, td0 + delta)
            Catch
            End Try

            Try : CallByName(dimObj, "BreakAlignmentSet", CallType.Method) : Catch : End Try
            Try : CallByName(dimObj, "RemoveFromAlignmentSet", CallType.Method) : Catch : End Try
            Try : CallByName(dimObj, "UpdateStatus", CallType.Method) : Catch : End Try
            Try : CallByName(draft, "UpdateAll", CallType.Method, True) : Catch : End Try
        End If

        Dim statusTxt As String = SafeToString(CallByNameSafe(dimObj, "StatusOfDimension", False))
        Dim relCount As Integer = SafeToInt(CallByNameSafe(dimObj, "GetRelatedCount", True))
        log?.LogLine("[DIM][DETACH][STATUS] status=" & statusTxt)
        log?.LogLine("[DIM][DETACH][REL] count=" & relCount.ToString(CultureInfo.InvariantCulture))

        connected = IsDimensionConnectedToView(draft, sheet, dv, dimObj)
        If (Not connected) AndAlso relCount > 0 Then
            RepositionOutsideView(dimObj, viewInfo, p, log)
            log?.LogLine("[DIM][DETACH][OK] mode=native")
            log?.LogLine("[DIM][FINAL][KEEP] mode=native")
            Return dimObj
        End If

        ' Fallback robusto: copiar/pegar en hoja y eliminar original.
        Try
            Dim dims As Dimensions = CType(sheet.Dimensions, Dimensions)
            Dim beforeCount As Integer = 0
            Try : beforeCount = dims.Count : Catch : beforeCount = 0 : End Try

            CallByName(dimObj, "Copy", CallType.Method)
            CallByName(sheet, "Paste", CallType.Method)

            Dim afterCount As Integer = 0
            Try : afterCount = dims.Count : Catch : afterCount = beforeCount : End Try
            If afterCount > 0 AndAlso afterCount >= beforeCount Then
                Dim pasted As Dimension = TryCast(CallByName(dims, "Item", CallType.Method, afterCount), Dimension)
                If pasted IsNot Nothing Then
                    Try : dimObj.Delete() : Catch : End Try
                    Dim pastedConnected As Boolean = IsDimensionConnectedToView(draft, sheet, dv, pasted)
                    Dim pastedRel As Integer = SafeToInt(CallByNameSafe(pasted, "GetRelatedCount", True))
                    Dim pastedStatus As String = SafeToString(CallByNameSafe(pasted, "StatusOfDimension", False))
                    log?.LogLine("[DIM][DETACH][STATUS] status=" & pastedStatus & " mode=copy_paste")
                    log?.LogLine("[DIM][DETACH][REL] count=" & pastedRel.ToString(CultureInfo.InvariantCulture) & " mode=copy_paste")
                    If (Not pastedConnected) AndAlso pastedRel > 0 Then
                        RepositionOutsideView(pasted, viewInfo, p, log)
                        log?.LogLine("[DIM][DETACH][OK] mode=copy_paste")
                        log?.LogLine("[DIM][FINAL][KEEP] mode=copy_paste")
                        Return pasted
                    End If
                    log?.LogLine("[DIM][FINAL][DELETE] mode=copy_paste_invalid")
                    Try : pasted.Delete() : Catch : End Try
                End If
            End If
        Catch ex As Exception
            log?.LogLine("[DIM][DETACH][FAIL] copy_paste msg=" & ex.Message)
        End Try

        log?.LogLine("[DIM][DETACH][FAIL] still_connected=True")
        log?.LogLine("[DIM][FINAL][KEEP] mode=original_connected")
        Return dimObj
    End Function

    Private Shared Sub RepositionOutsideView(dimObj As Dimension, info As DrawingViewGeometryInfo, p As PlannedDimension, log As DimensionLogger)
        If dimObj Is Nothing OrElse info Is Nothing Then Return
        Try
            Dim delta As Double = Math.Max(info.Box.Width, info.Box.Height) * 0.18R
            If delta <= 0.0R Then delta = 0.01R
            If p IsNot Nothing AndAlso p.Kind = PlannedDimensionKind.VerticalTotal Then
                Dim xOut As Double = info.Box.MaxX + delta
                Dim yMid As Double = (p.P1Y + p.P2Y) * 0.5R
                CallByName(dimObj, "SetKeyPoint", CallType.Method, 0, xOut, yMid, 0.0R)
            Else
                Dim yOut As Double = info.Box.MaxY + delta
                Dim xMid As Double = If(p Is Nothing, info.Box.MinX + (info.Box.Width * 0.5R), (p.P1X + p.P2X) * 0.5R)
                CallByName(dimObj, "SetKeyPoint", CallType.Method, 0, xMid, yOut, 0.0R)
            End If
            log?.LogLine("[DIM][RELOCATE] ok=True")
        Catch
            Try
                Dim td As Double = SafeToDouble(CallByName(dimObj, "TrackDistance", CallType.Get))
                CallByName(dimObj, "TrackDistance", CallType.Let, td + Math.Max(info.Box.Width, info.Box.Height) * 0.2R)
                log?.LogLine("[DIM][RELOCATE] ok=True mode=trackdistance")
            Catch ex As Exception
                log?.LogLine("[DIM][RELOCATE] ok=False msg=" & ex.Message)
            End Try
        End Try
    End Sub

    Private Shared Function IsDimensionConnectedToView(draft As DraftDocument, sheet As Sheet, dv As DrawingView, dimObj As Dimension) As Boolean
        If draft Is Nothing OrElse sheet Is Nothing OrElse dv Is Nothing OrElse dimObj Is Nothing Then Return False
        Dim ss As Object = Nothing
        Try
            ss = CallByName(draft, "SelectSet", CallType.Get)
            If ss Is Nothing Then ss = CallByName(sheet, "SelectSet", CallType.Get)
            If ss Is Nothing Then Return False

            Try : CallByName(ss, "RemoveAll", CallType.Method) : Catch : End Try
            Try : CallByName(dv, "AddConnectedDimensionsToSelectSet", CallType.Method) : Catch : End Try
            Dim n As Integer = 0
            Try : n = CInt(CallByName(ss, "Count", CallType.Get)) : Catch : n = 0 : End Try
            For i As Integer = 1 To n
                Dim it As Object = Nothing
                Try : it = CallByName(ss, "Item", CallType.Method, i) : Catch : it = Nothing : End Try
                If Object.ReferenceEquals(it, dimObj) Then Return True
            Next
        Catch
        End Try
        Return False
    End Function

    Friend Shared Sub TryApplyStyleToDimension(dimObj As Dimension, styleObj As Object, log As DimensionLogger)
        If dimObj Is Nothing Then Return

        Dim preferredName As String = ""
        Try
            If styleObj IsNot Nothing Then
                preferredName = Convert.ToString(CallByName(styleObj, "Name", CallType.Get), CultureInfo.InvariantCulture)
            End If
        Catch
            preferredName = ""
        End Try
        If String.IsNullOrWhiteSpace(preferredName) Then preferredName = ForcedStyleName
        Dim targetName As String = ForcedStyleName

        Dim applied As Boolean = False

        ' -2) Asignación directa COM (más fiable en algunos builds que CallByName).
        If styleObj IsNot Nothing Then
            Try
                dimObj.Style = CType(styleObj, DimStyle)
                If IsDimensionStyleApplied(dimObj, ForcedStyleName) OrElse IsDimensionStyleApplied(dimObj, preferredName) Then
                    applied = True
                End If
            Catch
            End Try
        End If

        ' -1) Forzado explícito solicitado por usuario.
        If Not applied Then
            For Each nameCandidate In BuildStyleNameCandidates(targetName)
                Try
                    CallByName(dimObj, "Style", CallType.Method, nameCandidate)
                    If IsDimensionStyleApplied(dimObj, nameCandidate) Then
                        applied = True
                        Exit For
                    End If
                Catch
                End Try
            Next
        End If

        ' 0) Camino COM específico observado: Dimension.Style("U3,5")
        If (Not applied) AndAlso (Not String.IsNullOrWhiteSpace(preferredName)) Then
            For Each nameCandidate In BuildStyleNameCandidates(preferredName)
                Try
                    CallByName(dimObj, "Style", CallType.Method, nameCandidate)
                    If IsDimensionStyleApplied(dimObj, preferredName) OrElse IsDimensionStyleApplied(dimObj, nameCandidate) Then
                        applied = True
                        Exit For
                    End If
                Catch
                End Try
            Next
        End If

        ' 1) Intento por objeto de estilo.
        If Not applied Then
            Try
                CallByName(dimObj, "Style", CallType.Let, styleObj)
                applied = IsDimensionStyleApplied(dimObj, preferredName)
            Catch
            End Try
        End If

        ' 2) Intentos por nombre (incluye variantes coma/punto).
        If Not applied AndAlso Not String.IsNullOrWhiteSpace(preferredName) Then
            For Each nameCandidate In BuildStyleNameCandidates(preferredName)
                Try
                    CallByName(dimObj, "Style", CallType.Let, nameCandidate)
                    If IsDimensionStyleApplied(dimObj, preferredName) OrElse IsDimensionStyleApplied(dimObj, nameCandidate) Then
                        applied = True
                        Exit For
                    End If
                Catch
                End Try
            Next
        End If

        ' 3) Refuerzo final: repetir con valor exacto "U3,5" y luego por objeto.
        If Not IsDimensionStyleApplied(dimObj, targetName) Then
            Try
                CallByName(dimObj, "Style", CallType.Method, targetName)
            Catch
            End Try
            Try
                CallByName(dimObj, "Style", CallType.Let, targetName)
            Catch
            End Try
            If styleObj IsNot Nothing Then
                Try
                    CallByName(dimObj, "Style", CallType.Let, styleObj)
                Catch
                End Try
            End If
        End If

        Dim finalStyleName As String = ReadDimensionStyleName(dimObj)
        If IsDimensionStyleApplied(dimObj, targetName) Then
            log?.LogLine("[DIM][STYLE] requested=" & targetName & " final=" & finalStyleName)
        Else
            log?.LogLine("[DIM][STYLE][WARN] requested=" & targetName & " final=" & finalStyleName)
        End If
    End Sub

    Private Shared Function ReadDimensionStyleName(dimObj As Dimension) As String
        If dimObj Is Nothing Then Return ""
        Try
            Dim st As Object = CallByName(dimObj, "Style", CallType.Get)
            If st Is Nothing Then Return ""
            Try
                Return Convert.ToString(CallByName(st, "Name", CallType.Get), CultureInfo.InvariantCulture)
            Catch
                Return Convert.ToString(st, CultureInfo.InvariantCulture)
            End Try
        Catch
            Return ""
        End Try
    End Function

    Private Shared Function IsDimensionStyleApplied(dimObj As Dimension, expectedName As String) As Boolean
        Dim current As String = NormalizeStyleName(ReadDimensionStyleName(dimObj))
        Dim expected As String = NormalizeStyleName(expectedName)
        If String.IsNullOrWhiteSpace(current) OrElse String.IsNullOrWhiteSpace(expected) Then Return False
        Return String.Equals(current, expected, StringComparison.OrdinalIgnoreCase)
    End Function

    Private Shared Function BuildStyleNameCandidates(name As String) As List(Of String)
        Dim out As New List(Of String)()
        If String.IsNullOrWhiteSpace(name) Then Return out
        Dim n As String = name.Trim()
        out.Add(n)
        out.Add(n.Replace(",", "."))
        out.Add(n.Replace(".", ","))
        Return out.Distinct(StringComparer.OrdinalIgnoreCase).ToList()
    End Function

    Private Shared Function NormalizeStyleName(s As String) As String
        If String.IsNullOrWhiteSpace(s) Then Return ""
        Dim t As String = s.Trim().Replace(ChrW(&HA0), " ")
        Do While t.Contains("  ")
            t = t.Replace("  ", " ")
        Loop
        Return t
    End Function

    Private Shared Function CallByNameSafe(obj As Object, member As String, isMethod As Boolean) As Object
        If obj Is Nothing Then Return Nothing
        Try
            If isMethod Then
                Return CallByName(obj, member, CallType.Method)
            End If
            Return CallByName(obj, member, CallType.Get)
        Catch
            Return Nothing
        End Try
    End Function

    Private Shared Function SafeToInt(v As Object) As Integer
        Try
            If v Is Nothing Then Return 0
            Return Convert.ToInt32(v, CultureInfo.InvariantCulture)
        Catch
            Return 0
        End Try
    End Function

    Private Shared Function SafeToDouble(v As Object) As Double
        Try
            If v Is Nothing Then Return 0.0R
            Return Convert.ToDouble(v, CultureInfo.InvariantCulture)
        Catch
            Return 0.0R
        End Try
    End Function

    Private Shared Function IsDimensionOutsideSheetBounds(
        ByVal dimObj As Dimension,
        ByVal sheet As Sheet,
        ByVal log As DimensionLogger
    ) As Boolean
        If dimObj Is Nothing OrElse sheet Is Nothing Then Return False
        Try
            Dim width As Double = 0.0R
            Dim height As Double = 0.0R
            If Not TryReadSheetSize(sheet, width, height) Then
                log?.LogLine("[DIM][BOUNDS][SKIP] No se pudieron leer dimensiones de hoja (Width/Height o SheetWidth/SheetHeight).")
                Return False
            End If

            Dim x1 As Double = 0.0R, y1 As Double = 0.0R, x2 As Double = 0.0R, y2 As Double = 0.0R
            dimObj.Range(x1, y1, x2, y2)

            Dim minX As Double = Math.Min(x1, x2)
            Dim maxX As Double = Math.Max(x1, x2)
            Dim minY As Double = Math.Min(y1, y2)
            Dim maxY As Double = Math.Max(y1, y2)

            Dim outside As Boolean = (minX < 0.0R OrElse minY < 0.0R OrElse maxX > width OrElse maxY > height)
            If outside Then
                log?.LogLine("[DIM][BOUNDS][OUTSIDE] range=(" &
                             minX.ToString("0.######", CultureInfo.InvariantCulture) & "," &
                             minY.ToString("0.######", CultureInfo.InvariantCulture) & ")-(" &
                             maxX.ToString("0.######", CultureInfo.InvariantCulture) & "," &
                             maxY.ToString("0.######", CultureInfo.InvariantCulture) & ")" &
                             " sheet=(" & width.ToString("0.######", CultureInfo.InvariantCulture) & "," &
                             height.ToString("0.######", CultureInfo.InvariantCulture) & ")")
            End If
            Return outside
        Catch ex As Exception
            log?.LogLine("[DIM][BOUNDS][SKIP] no se pudo validar Range: " & ex.Message)
            Return False
        End Try
    End Function

    Private Shared Function TryReadSheetSize(ByVal sheet As Sheet, ByRef width As Double, ByRef height As Double) As Boolean
        width = 0.0R
        height = 0.0R
        If sheet Is Nothing Then Return False

        ' Prioridad 1 (SDK Draft): Sheet.SheetSetup.SheetWidth / SheetHeight.
        Try
            Dim ss As Object = CallByName(sheet, "SheetSetup", CallType.Get)
            If ss IsNot Nothing Then
                width = Convert.ToDouble(CallByName(ss, "SheetWidth", CallType.Get), CultureInfo.InvariantCulture)
                height = Convert.ToDouble(CallByName(ss, "SheetHeight", CallType.Get), CultureInfo.InvariantCulture)
                If width > 0.0R AndAlso height > 0.0R Then Return True
            End If
        Catch
        End Try

        ' Prioridad 2: variantes abreviadas.
        Try
            width = Convert.ToDouble(CallByName(sheet, "Width", CallType.Get), CultureInfo.InvariantCulture)
            height = Convert.ToDouble(CallByName(sheet, "Height", CallType.Get), CultureInfo.InvariantCulture)
            If width > 0.0R AndAlso height > 0.0R Then Return True
        Catch
        End Try

        ' Prioridad 3: fallback para variantes de interop.
        Try
            width = Convert.ToDouble(CallByName(sheet, "SheetWidth", CallType.Get), CultureInfo.InvariantCulture)
            height = Convert.ToDouble(CallByName(sheet, "SheetHeight", CallType.Get), CultureInfo.InvariantCulture)
            If width > 0.0R AndAlso height > 0.0R Then Return True
        Catch
        End Try

        Return False
    End Function

    Private Shared Function SafeToString(v As Object) As String
        Try
            If v Is Nothing Then Return ""
            Return Convert.ToString(v, CultureInfo.InvariantCulture)
        Catch
            Return ""
        End Try
    End Function
End Class
