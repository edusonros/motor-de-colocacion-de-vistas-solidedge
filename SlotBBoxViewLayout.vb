Option Strict Off
Imports System.Collections.Generic
Imports System.Globalization
Imports SolidEdgeDraft
Imports SolidEdgeFramework

''' <summary>Motor opcional de colocación por slots (bbox en hoja real) usando centro de <see cref="DrawingView.Range"/>.</summary>
Friend Class ViewSlot
    Public Property Name As String
    Public Property MinX As Double
    Public Property MinY As Double
    Public Property MaxX As Double
    Public Property MaxY As Double

    Public ReadOnly Property CenterX As Double
        Get
            Return (MinX + MaxX) / 2.0R
        End Get
    End Property

    Public ReadOnly Property CenterY As Double
        Get
            Return (MinY + MaxY) / 2.0R
        End Get
    End Property

    Public ReadOnly Property Width As Double
        Get
            Return MaxX - MinX
        End Get
    End Property

    Public ReadOnly Property Height As Double
        Get
            Return MaxY - MinY
        End Get
    End Property

    Public Overrides Function ToString() As String
        Return String.Format(CultureInfo.InvariantCulture, "{0} [{1:0.######},{2:0.######}]-[{3:0.######},{4:0.######}]",
                             Name, MinX, MinY, MaxX, MaxY)
    End Function
End Class

Public Module SlotBBoxViewLayout

    Private Const RefSheetW As Double = 0.42R
    Private Const RefSheetH As Double = 0.297R

    ''' <summary>Polígono de área útil referencia A3 horizontal (metros, sistema hoja SE).</summary>
    Private ReadOnly RefFreePoly As Double()() = {
        New Double() {0.01R, 0.01R},
        New Double() {0.2198R, 0.01R},
        New Double() {0.2198R, 0.04R},
        New Double() {0.4098R, 0.04R},
        New Double() {0.4098R, 0.2762R},
        New Double() {0.01R, 0.2762R}
    }

    Private Const TallRatioThreshold As Double = 2.0R
    Private Const ValidateTol As Double = 0.001R
    Private Const MaxValidateIterations As Integer = 14

    Private Sub Lg(log As Action(Of String), msg As String)
        log?.Invoke(msg)
    End Sub

    Private Sub DoIdleSafe(app As Application)
        If app Is Nothing Then Return
        Try : app.DoIdle() : Catch : End Try
    End Sub

    Private Sub SafeUpdateView(dv As DrawingView)
        If dv Is Nothing Then Return
        Try : dv.Update() : Catch : End Try
    End Sub

    Private Sub DraftUpdateAll(dft As DraftDocument)
        If dft Is Nothing Then Return
        Try : dft.UpdateAll(True) : Catch : End Try
    End Sub

    Private Sub MoveViewRangeCenterTo(dv As DrawingView,
                                     targetCenterX As Double,
                                     targetCenterY As Double,
                                     logPrefix As String,
                                     log As Action(Of String))
        If dv Is Nothing Then Return
        Dim viewName As String = SafeViewName(dv)

        Dim minX As Double, minY As Double, maxX As Double, maxY As Double
        If Not CojonudoBestFit_Bueno.TryGetViewRangePublic(dv, minX, minY, maxX, maxY) Then
            Lg(log, $"{logPrefix}[VIEWLAYOUT][CENTER][BEFORE] name={viewName} sin Range legible.")
            Return
        End If
        NormalizeRange(minX, minY, maxX, maxY)

        Dim currentCenterX = (minX + maxX) / 2.0R
        Dim currentCenterY = (minY + maxY) / 2.0R

        Dim originX As Double = 0, originY As Double = 0
        Try : dv.GetOrigin(originX, originY) : Catch : End Try

        Lg(log, $"{logPrefix}[VIEWLAYOUT][CENTER][BEFORE] name={viewName} Range=({minX:0.######},{minY:0.######})-({maxX:0.######},{maxY:0.######}) " &
                  $"cenR=({currentCenterX:0.######},{currentCenterY:0.######}) ori=({originX:0.######},{originY:0.######}) tgt=({targetCenterX:0.######},{targetCenterY:0.######})")

        Dim newOriginX = originX + targetCenterX - currentCenterX
        Dim newOriginY = originY + targetCenterY - currentCenterY

        Lg(log, $"{logPrefix}[VIEWLAYOUT][CENTER][MOVE] name={viewName} newOrigin=({newOriginX:0.######},{newOriginY:0.######})")

        Try
            dv.SetOrigin(newOriginX, newOriginY)
        Catch ex As Exception
            Lg(log, $"{logPrefix}[VIEWLAYOUT][CENTER][AFTER][ERR] name={viewName} {ex.Message}")
            Return
        End Try

        If Not CojonudoBestFit_Bueno.TryGetViewRangePublic(dv, minX, minY, maxX, maxY) Then
            Lg(log, $"{logPrefix}[VIEWLAYOUT][CENTER][AFTER] name={viewName} sin Range tras SetOrigin.")
            Return
        End If
        NormalizeRange(minX, minY, maxX, maxY)
        Lg(log, $"{logPrefix}[VIEWLAYOUT][CENTER][AFTER] name={viewName} Range=({minX:0.######},{minY:0.######})-({maxX:0.######},{maxY:0.######})")
    End Sub

    Private Function SafeViewName(dv As DrawingView) As String
        Try
            Dim s = TryCast(dv.Name, String)
            If Not String.IsNullOrWhiteSpace(s) Then Return s
        Catch
        End Try
        Return "DrawingView"
    End Function

    Private Sub NormalizeRange(ByRef xmin As Double, ByRef ymin As Double, ByRef xmax As Double, ByRef ymax As Double)
        Dim xl = xmin : Dim yl = ymin : Dim xh = xmax : Dim yh = ymax
        xmin = Math.Min(xl, xh) : xmax = Math.Max(xl, xh)
        ymin = Math.Min(yl, yh) : ymax = Math.Max(yl, yh)
    End Sub

    Friend Function ScaleSlotFromReference(refSlot As ViewSlot, actualSheetWidth As Double, actualSheetHeight As Double) As ViewSlot
        Dim a As New ViewSlot With {.Name = refSlot.Name}
        a.MinX = refSlot.MinX / RefSheetW * actualSheetWidth
        a.MinY = refSlot.MinY / RefSheetH * actualSheetHeight
        a.MaxX = refSlot.MaxX / RefSheetW * actualSheetWidth
        a.MaxY = refSlot.MaxY / RefSheetH * actualSheetHeight
        Return a
    End Function

    ''' <summary>Escala UNE más grande menor o igual que <paramref name="maxAllowedScale"/>.</summary>
    Friend Function ChooseRecommendedScale(maxAllowedScale As Double, log As Action(Of String), logPrefix As String) As Double
        Lg(log, $"{logPrefix}[VIEWLAYOUT][SCALE][FIT] maxAllowed={maxAllowedScale.ToString("G9", CultureInfo.InvariantCulture)}")
        Dim smallest = TemplateBboxLayout.UneDrawingScaleFactors(TemplateBboxLayout.UneDrawingScaleFactors.Length - 1)
        If maxAllowedScale < smallest - 1.0E-12 Then
            Lg(log, $"{logPrefix}[VIEWLAYOUT][SCALE][WARN] maxAllowed menor que escala UNE mínima; se fuerza 1:200 (= {smallest.ToString("G9", CultureInfo.InvariantCulture)})")
        End If
        Dim snapped = TemplateBboxLayout.SnapUneScale(Math.Min(maxAllowedScale, 1.0R))
        Lg(log, $"{logPrefix}[VIEWLAYOUT][SCALE][UNE] chosen={snapped.ToString("G9", CultureInfo.InvariantCulture)}")
        Return snapped
    End Function

    Private Function GetScaleFactorSafe(dv As DrawingView) As Double
        If dv Is Nothing Then Return 1.0R
        Try
            Return CDbl(dv.ScaleFactor)
        Catch
            Return 1.0R
        End Try
    End Function

    Private Function TrySetScaleFactor(dv As DrawingView, sf As Double, log As Action(Of String), tag As String) As Boolean
        If dv Is Nothing Then Return False
        Try
            dv.ScaleFactor = sf
            Return True
        Catch ex As Exception
            Lg(log, $"[VIEWLAYOUT][SCALE][WARN] SetScaleFactor {tag}: {ex.Message}")
            Return False
        End Try
    End Function

    Private Function FitScaleMultiply(slot As ViewSlot, rw As Double, rh As Double) As Double
        If rw <= 1.0E-15 OrElse rh <= 1.0E-15 Then Return 0
        Return Math.Min(slot.Width / rw, slot.Height / rh)
    End Function

    Private Function UneIndexForScale(sf As Double) As Integer
        Dim arr = TemplateBboxLayout.UneDrawingScaleFactors
        For i = 0 To arr.Length - 1
            If Math.Abs(arr(i) - sf) < 1.0E-8 Then Return i
        Next
        Dim snapped = TemplateBboxLayout.SnapUneScale(sf)
        For i = 0 To arr.Length - 1
            If Math.Abs(arr(i) - snapped) < 1.0E-8 Then Return i
        Next
        Return Math.Max(0, arr.Length - 1)
    End Function

    ''' <summary>Escalón UNE inmediatamente menor (dibujo más reducido) que <paramref name="currentUne"/>.</summary>
    Private Function NextUneSmaller(currentUne As Double) As Double?
        Dim arr = TemplateBboxLayout.UneDrawingScaleFactors
        Dim idx = UneIndexForScale(currentUne)
        If idx < 0 OrElse idx >= arr.Length - 1 Then Return Nothing
        Return arr(idx + 1)
    End Function

    Private Function ViewRangePair(dv As DrawingView,
                                   ByRef rw As Double, ByRef rh As Double) As Boolean
        rw = 0 : rh = 0
        Dim xmin As Double, ymin As Double, xmax As Double, ymax As Double
        If Not CojonudoBestFit_Bueno.TryGetViewRangePublic(dv, xmin, ymin, xmax, ymax) Then Return False
        NormalizeRange(xmin, ymin, xmax, ymax)
        rw = xmax - xmin : rh = ymax - ymin
        Return rw > 1.0E-15 AndAlso rh > 1.0E-15
    End Function

    Private Sub LogOneViewRange(dv As DrawingView, lineTag As String, pf As String, log As Action(Of String))
        If dv Is Nothing Then Return
        Dim rx As Double, ry As Double, rxx As Double, ryy As Double
        If Not CojonudoBestFit_Bueno.TryGetViewRangePublic(dv, rx, ry, rxx, ryy) Then
            Lg(log, $"{pf}{lineTag} name={SafeViewName(dv)} Range=(n/a)")
            Return
        End If
        NormalizeRange(rx, ry, rxx, ryy)
        Lg(log, $"{pf}{lineTag} name={SafeViewName(dv)} Range=({rx:0.######},{ry:0.######})-({rxx:0.######},{ryy:0.######}) SF={GetScaleFactorSafe(dv):0.######}")
    End Sub

    Private Sub LogAllViewRanges(layout As ResolvedLayout,
                                 vMain As DrawingView, vRight As DrawingView, vBelow As DrawingView,
                                 vIso As DrawingView, vFlat As DrawingView,
                                 flatInserted As Boolean, pf As String, log As Action(Of String))
        Dim tag = "[VIEWLAYOUT][VIEW][RANGE]"
        LogOneViewRange(vMain, $"{tag} Main", pf, log)
        LogOneViewRange(vRight, $"{tag} Right", pf, log)
        LogOneViewRange(vBelow, $"{tag} Below", pf, log)
        If layout.IncludeIso AndAlso vIso IsNot Nothing Then LogOneViewRange(vIso, $"{tag} Iso", pf, log)
        If flatInserted AndAlso layout.IncludeFlat AndAlso vFlat IsNot Nothing Then LogOneViewRange(vFlat, $"{tag} Flat", pf, log)
    End Sub

    Private Function ValidateViewInSlot(dv As DrawingView, slot As ViewSlot,
                                        ByRef failReason As String) As Boolean
        failReason = ""
        Dim xmin As Double, ymin As Double, xmax As Double, ymax As Double
        If slot Is Nothing OrElse Not CojonudoBestFit_Bueno.TryGetViewRangePublic(dv, xmin, ymin, xmax, ymax) Then
            failReason = "sin Range "
            Return False
        End If
        NormalizeRange(xmin, ymin, xmax, ymax)
        If xmin < slot.MinX - ValidateTol Then failReason &= $"MinX<{slot.MinX} "
        If xmax > slot.MaxX + ValidateTol Then failReason &= $"MaxX>{slot.MaxX} "
        If ymin < slot.MinY - ValidateTol Then failReason &= $"MinY<{slot.MinY} "
        If ymax > slot.MaxY + ValidateTol Then failReason &= $"MaxY>{slot.MaxY} "
        Return failReason.Length = 0
    End Function

    Private Function DetectFormatLabel(w As Double, h As Double) As String
        Dim tol As Double = 0.004R
        Dim wNorm = Math.Min(w, h)
        Dim hNorm = Math.Max(w, h)
        If Math.Abs(wNorm - 0.21R) < tol AndAlso Math.Abs(hNorm - 0.297R) < tol Then Return "A4"
        If Math.Abs(wNorm - 0.297R) < tol AndAlso Math.Abs(hNorm - 0.42R) < tol Then Return "A3"
        If Math.Abs(wNorm - 0.42R) < tol AndAlso Math.Abs(hNorm - 0.594R) < tol Then Return "A2"
        If Math.Abs(wNorm - 0.594R) < tol AndAlso Math.Abs(hNorm - 0.841R) < tol Then Return "A1"
        Return "CUSTOM"
    End Function

    Private Function BuildReferenceSlots() As ViewSlot()
        Return New ViewSlot() {
            New ViewSlot With {.Name = "Main", .MinX = 0.0264R, .MinY = 0.1723R, .MaxX = 0.2334R, .MaxY = 0.2719R},
            New ViewSlot With {.Name = "Right", .MinX = 0.2454R, .MinY = 0.1717R, .MaxX = 0.3965R, .MaxY = 0.2719R},
            New ViewSlot With {.Name = "Below", .MinX = 0.0264R, .MinY = 0.0883R, .MaxX = 0.2334R, .MaxY = 0.1723R},
            New ViewSlot With {.Name = "Flat", .MinX = 0.0156R, .MinY = 0.01R, .MaxX = 0.2198R, .MaxY = 0.0869R},
            New ViewSlot With {.Name = "Iso", .MinX = 0.3163R, .MinY = 0.04R, .MaxX = 0.4098R, .MaxY = 0.1013R}
        }
    End Function

    Friend Function TryApplySlotBBoxLayout(app As Application,
                                           dft As DraftDocument,
                                           sheet As Sheet,
                                           layout As ResolvedLayout,
                                           vMain As DrawingView,
                                           vRight As DrawingView,
                                           vBelow As DrawingView,
                                           vIso As DrawingView,
                                           vFlat As DrawingView,
                                           flatInserted As Boolean,
                                           log As Action(Of String)) As Boolean

        Dim pf As String = ""
        Try
            Lg(log, $"{pf}[VIEWLAYOUT][ENTER] Slot+Bbox placement (Range center)")

            If app Is Nothing OrElse dft Is Nothing OrElse sheet Is Nothing OrElse layout Is Nothing OrElse log Is Nothing Then
                Return False
            End If
            If vMain Is Nothing OrElse vRight Is Nothing OrElse vBelow Is Nothing Then
                Lg(log, $"{pf}[VIEWLAYOUT][FALLBACK_TO_LEGACY] vistas principales incompletas")
                Return False
            End If

            Lg(log, $"{pf}[VIEWLAYOUT][CONFIG] EnableSlotBBoxViewLayout ejecutando motor extendido tras inserción AddByFold clásica")

            Dim sw As Double = 0, sh As Double = 0
            Try
                sw = sheet.SheetSetup.SheetWidth
                sh = sheet.SheetSetup.SheetHeight
            Catch ex As Exception
                Lg(log, $"{pf}[VIEWLAYOUT][FALLBACK_TO_LEGACY] SheetSetup: {ex.Message}")
                Return False
            End Try

            Dim landscape = sw >= sh
            Lg(log, $"{pf}[VIEWLAYOUT][SHEET] width={sw.ToString("G9", CultureInfo.InvariantCulture)} height={sh.ToString("G9", CultureInfo.InvariantCulture)} orientation={If(landscape, "landscape", "portrait")}")

            Dim refSlots = BuildReferenceSlots()
            For Each rs In refSlots
                Lg(log, $"{pf}[VIEWLAYOUT][SLOT][REF] {rs}")
            Next

            Dim slots As New Dictionary(Of String, ViewSlot)(StringComparer.OrdinalIgnoreCase)
            For Each rs In refSlots
                Dim act = ScaleSlotFromReference(rs, sw, sh)
                Lg(log, $"{pf}[VIEWLAYOUT][SLOT][ACTUAL] {act}")
                slots(rs.Name) = act
            Next

            For i = 0 To RefFreePoly.Length - 1
                Dim p = RefFreePoly(i)
                Dim ax = p(0) / RefSheetW * sw
                Dim ay = p(1) / RefSheetH * sh
                Lg(log, $"{pf}[VIEWLAYOUT][FREE_AREA] vtx{i} Sheet=({ax.ToString("G9", CultureInfo.InvariantCulture)},{ay.ToString("G9", CultureInfo.InvariantCulture)})")
            Next

            DraftUpdateAll(dft)
            DoIdleSafe(app)
            LogAllViewRanges(layout, vMain, vRight, vBelow, vIso, vFlat, flatInserted, pf, log)

            ' ----- Rotación piezas muy altas (bloque ortográficas, -90º) -----
            Dim rwMn As Double, rhMn As Double
            If ViewRangePair(vMain, rwMn, rhMn) Then
                Dim ratio As Double = If(rwMn <= 1.0E-15, Double.MaxValue, rhMn / rwMn)
                Lg(log, $"{pf}[VIEWLAYOUT][ROTATE_CHECK] Main h/w ratio={ratio:0.######} umb={TallRatioThreshold:0.######}")
                If ratio >= TallRatioThreshold - 1.0E-12 Then
                    Dim okRot = CojonudoBestFit_Bueno.RotateDrawingViewsBlockByAngle(app, vMain, vRight, vBelow, -Math.PI / 2.0R)
                    Lg(log, $"{pf}[VIEWLAYOUT][ROTATE_APPLY] block -90deg success={okRot}")
                    DraftUpdateAll(dft)
                    DoIdleSafe(app)
                    If flatInserted AndAlso layout.IncludeFlat AndAlso vFlat IsNot Nothing Then
                        Try
                            Dim cur As Double = 0
                            vFlat.GetRotationAngle(cur)
                            vFlat.SetRotationAngle(cur - Math.PI / 2.0R)
                            SafeUpdateView(vFlat)
                            DoIdleSafe(app)
                        Catch exFx As Exception
                            Lg(log, $"{pf}[VIEWLAYOUT][ROTATE_APPLY][FLAT_WARN] {exFx.Message}")
                        End Try
                    End If
                Else
                    Lg(log, $"{pf}[VIEWLAYOUT][ROTATE_SKIP] ratio bajo umbral")
                End If
            Else
                Lg(log, $"{pf}[VIEWLAYOUT][ROTATE_SKIP] sin medida principal")
            End If

            DraftUpdateAll(dft)
            DoIdleSafe(app)

            Dim rwM As Double, rhM As Double, rwR As Double, rhR As Double, rwB As Double, rhB As Double
            If Not ViewRangePair(vMain, rwM, rhM) OrElse Not ViewRangePair(vRight, rwR, rhR) OrElse Not ViewRangePair(vBelow, rwB, rhB) Then
                Lg(log, $"{pf}[VIEWLAYOUT][FALLBACK_TO_LEGACY] no se puede medir Rangos para escala grupo principal")
                Return False
            End If

            Dim sBase = GetScaleFactorSafe(vMain)
            Dim fitMain = FitScaleMultiply(slots("Main"), rwM, rhM)
            Dim fitRight = FitScaleMultiply(slots("Right"), rwR, rhR)
            Dim fitBelow = FitScaleMultiply(slots("Below"), rwB, rhB)

            Dim mainGrpMult = Math.Min(fitMain, Math.Min(fitRight, fitBelow))
            Dim maxSfMainGrp = sBase * mainGrpMult
            Lg(log, $"{pf}[VIEWLAYOUT][SCALE][MAIN_GROUP] fits mult min={mainGrpMult.ToString("G9", CultureInfo.InvariantCulture)} ceilingSf={maxSfMainGrp.ToString("G9", CultureInfo.InvariantCulture)}")
            Dim uneMain = ChooseRecommendedScale(maxSfMainGrp, log, pf)

            Dim uneIso As Double = uneMain
            If layout.IncludeIso AndAlso vIso IsNot Nothing Then
                Dim viw As Double, vih As Double
                If ViewRangePair(vIso, viw, vih) Then
                    Dim sIso0 = GetScaleFactorSafe(vIso)
                    Dim isoFitMult = FitScaleMultiply(slots("Iso"), viw, vih)
                    Dim maxIso = sIso0 * isoFitMult
                    Lg(log, $"{pf}[VIEWLAYOUT][SCALE][ISO] ceilingSf={maxIso.ToString("G9", CultureInfo.InvariantCulture)}")
                    uneIso = ChooseRecommendedScale(maxIso, log, pf)
                Else
                    uneIso = uneMain
                End If
            End If

            Dim uneFlat As Double? = Nothing
            If flatInserted AndAlso layout.IncludeFlat AndAlso vFlat IsNot Nothing Then
                Dim vfW As Double, vfH As Double
                If ViewRangePair(vFlat, vfW, vfH) Then
                    Dim sF0 = GetScaleFactorSafe(vFlat)
                    Dim flatFitMult = FitScaleMultiply(slots("Flat"), vfW, vfH)
                    Dim maxFlatSf = sF0 * flatFitMult
                    Lg(log, $"{pf}[VIEWLAYOUT][SCALE][FLAT] ceilingSf={maxFlatSf.ToString("G9", CultureInfo.InvariantCulture)}")
                    uneFlat = ChooseRecommendedScale(maxFlatSf, log, pf)
                End If
            End If

            Dim iterMain = uneMain
            Dim iterIso = uneIso
            Dim iterFlat As Double = If(uneFlat.HasValue, uneFlat.Value, GetScaleFactorSafe(vFlat))

            Dim fmt = DetectFormatLabel(sw, sh)

            For it = 0 To MaxValidateIterations - 1
                If Not TrySetScaleFactor(vMain, iterMain, log, "Main") Then Exit For
                If Not TrySetScaleFactor(vRight, iterMain, log, "Right") Then Exit For
                If Not TrySetScaleFactor(vBelow, iterMain, log, "Below") Then Exit For
                If layout.IncludeIso AndAlso vIso IsNot Nothing Then
                    TrySetScaleFactor(vIso, iterIso, log, "Iso")
                End If
                If uneFlat.HasValue AndAlso vFlat IsNot Nothing Then
                    TrySetScaleFactor(vFlat, iterFlat, log, "Flat")
                End If

                DraftUpdateAll(dft)
                DoIdleSafe(app)
                SafeUpdateView(vMain)
                SafeUpdateView(vRight)
                SafeUpdateView(vBelow)
                If vIso IsNot Nothing Then SafeUpdateView(vIso)
                If vFlat IsNot Nothing Then SafeUpdateView(vFlat)
                DraftUpdateAll(dft)
                DoIdleSafe(app)

                MoveViewRangeCenterTo(vMain, slots("Main").CenterX, slots("Main").CenterY, pf, log)
                MoveViewRangeCenterTo(vRight, slots("Right").CenterX, slots("Right").CenterY, pf, log)
                MoveViewRangeCenterTo(vBelow, slots("Below").CenterX, slots("Below").CenterY, pf, log)
                If layout.IncludeIso AndAlso vIso IsNot Nothing Then
                    MoveViewRangeCenterTo(vIso, slots("Iso").CenterX, slots("Iso").CenterY, pf, log)
                End If
                If flatInserted AndAlso layout.IncludeFlat AndAlso vFlat IsNot Nothing AndAlso uneFlat.HasValue Then
                    MoveViewRangeCenterTo(vFlat, slots("Flat").CenterX, slots("Flat").CenterY, pf, log)
                End If

                DraftUpdateAll(dft)
                DoIdleSafe(app)

                Dim failMain As String = "", failR As String = "", failB As String = ""
                Dim okM = ValidateViewInSlot(vMain, slots("Main"), failMain)
                Dim okR = ValidateViewInSlot(vRight, slots("Right"), failR)
                Dim okB = ValidateViewInSlot(vBelow, slots("Below"), failB)

                Dim frIso As String = ""
                Dim okI As Boolean = True
                If layout.IncludeIso AndAlso vIso IsNot Nothing Then okI = ValidateViewInSlot(vIso, slots("Iso"), frIso)

                Dim flatFail As String = ""
                Dim okF As Boolean = True
                If flatInserted AndAlso layout.IncludeFlat AndAlso vFlat IsNot Nothing AndAlso uneFlat.HasValue Then okF = ValidateViewInSlot(vFlat, slots("Flat"), flatFail)

                Dim allOk As Boolean = okM AndAlso okR AndAlso okB AndAlso okI AndAlso okF

                If okM Then Lg(log, $"{pf}[VIEWLAYOUT][VALIDATE][OK] Main within slot.") Else Lg(log, $"{pf}[VIEWLAYOUT][VALIDATE][FAIL] Main {failMain}")
                If okR Then Lg(log, $"{pf}[VIEWLAYOUT][VALIDATE][OK] Right within slot.") Else Lg(log, $"{pf}[VIEWLAYOUT][VALIDATE][FAIL] Right {failR}")
                If okB Then Lg(log, $"{pf}[VIEWLAYOUT][VALIDATE][OK] Below within slot.") Else Lg(log, $"{pf}[VIEWLAYOUT][VALIDATE][FAIL] Below {failB}")
                If layout.IncludeIso AndAlso vIso IsNot Nothing Then
                    If okI Then Lg(log, $"{pf}[VIEWLAYOUT][VALIDATE][OK] Iso within slot.") Else Lg(log, $"{pf}[VIEWLAYOUT][VALIDATE][FAIL] Iso {frIso}")
                End If
                If flatInserted AndAlso layout.IncludeFlat AndAlso uneFlat.HasValue Then
                    If okF Then Lg(log, $"{pf}[VIEWLAYOUT][VALIDATE][OK] Flat within slot.") Else Lg(log, $"{pf}[VIEWLAYOUT][VALIDATE][FAIL] Flat {flatFail}")
                End If

                If allOk Then Exit For

                Dim anyStep As Boolean = False
                If Not okM OrElse Not okR OrElse Not okB Then
                    Dim nxt As Double? = NextUneSmaller(iterMain)
                    If nxt.HasValue Then
                        Lg(log, $"{pf}[VIEWLAYOUT][VALIDATE][RETRY_SCALE_DOWN] Main grupo -> {nxt.Value.ToString("G9", CultureInfo.InvariantCulture)}")
                        iterMain = nxt.Value : anyStep = True
                    End If
                End If
                If Not okI Then
                    Dim nxtI As Double? = NextUneSmaller(iterIso)
                    If nxtI.HasValue Then
                        Lg(log, $"{pf}[VIEWLAYOUT][VALIDATE][RETRY_SCALE_DOWN] Iso -> {nxtI.Value.ToString("G9", CultureInfo.InvariantCulture)}")
                        iterIso = nxtI.Value : anyStep = True
                    End If
                End If
                If Not okF AndAlso uneFlat.HasValue Then
                    Dim nxf As Double? = NextUneSmaller(iterFlat)
                    If nxf.HasValue Then
                        Lg(log, $"{pf}[VIEWLAYOUT][VALIDATE][RETRY_SCALE_DOWN] Flat -> {nxf.Value.ToString("G9", CultureInfo.InvariantCulture)}")
                        iterFlat = nxf.Value : anyStep = True
                    End If
                End If

                If Not anyStep Then
                    Lg(log, $"{pf}[VIEWLAYOUT][VALIDATE][FINAL_WARN] No hay más escalones UNE; se conserva mejor esfuerzo.")
                    Exit For
                End If
            Next

            Lg(log, $"{pf}[VIEWLAYOUT][SUMMARY] format={fmt} orient={If(landscape, "landscape", "portrait")} sheet_m={sw.ToString("G6", CultureInfo.InvariantCulture)}x{sh.ToString("G6", CultureInfo.InvariantCulture)} uneMain={iterMain.ToString("G9", CultureInfo.InvariantCulture)} uneIso={iterIso.ToString("G9", CultureInfo.InvariantCulture)}")
            If uneFlat.HasValue Then Lg(log, $"{pf}[VIEWLAYOUT][SUMMARY] uneFlat={iterFlat.ToString("G9", CultureInfo.InvariantCulture)}")
            For Each kv In slots
                Lg(log, $"{pf}[VIEWLAYOUT][SUMMARY] assignedSlot {kv.Key}={kv.Value}")
            Next

            DraftUpdateAll(dft)
            DoIdleSafe(app)
            LogAllViewRanges(layout, vMain, vRight, vBelow, vIso, vFlat, flatInserted, pf, log)

            Dim rsnMain As String = "", rsnRight As String = "", rsnBelow As String = ""
            Dim inMain = ValidateViewInSlot(vMain, slots("Main"), rsnMain)
            Dim inR = ValidateViewInSlot(vRight, slots("Right"), rsnRight)
            Dim inB = ValidateViewInSlot(vBelow, slots("Below"), rsnBelow)
            Lg(log, $"{pf}[VIEWLAYOUT][SUMMARY] dentro_slot Main={inMain} Right={inR} Below={inB}")

            Dim fiIso As String = ""
            If layout.IncludeIso AndAlso vIso IsNot Nothing Then
                Dim okIz = ValidateViewInSlot(vIso, slots("Iso"), fiIso)
                Lg(log, $"{pf}[VIEWLAYOUT][SUMMARY] dentro_slot Iso={okIz}")
            End If

            Dim rsnFlat As String = ""
            If flatInserted AndAlso layout.IncludeFlat AndAlso uneFlat.HasValue AndAlso vFlat IsNot Nothing Then
                Dim okFl = ValidateViewInSlot(vFlat, slots("Flat"), rsnFlat)
                Lg(log, $"{pf}[VIEWLAYOUT][SUMMARY] dentro_slot Flat={okFl}")
            End If

            Lg(log, $"{pf}[VIEWLAYOUT][SUMMARY] fallback_a_layout_anterior=false")
            Lg(log, $"{pf}[VIEWLAYOUT][ENTER] DONE")
            Return True
        Catch ex As Exception
            Lg(log, $"[VIEWLAYOUT][FALLBACK_TO_LEGACY] Excepción: {ex.GetType().Name} {ex.Message}")
            Return False
        End Try
    End Function

End Module
