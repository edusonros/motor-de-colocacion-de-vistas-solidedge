Option Strict Off

Imports System
Imports System.Collections.Generic
Imports System.Globalization
Imports System.Linq
Imports System.Runtime.InteropServices
Imports Microsoft.VisualBasic
Imports SolidEdgeDraft
' No Imports SolidEdgeConstants: ViewOrientationConstants y SheetMetalDrawingViewTypeConstants
' también existen en SolidEdgeDraft y provocan BC30561 (ambiguo).

''' <summary>Separación opcional para futuras fases (desplazamiento manual de cotas).</summary>
Public Module DimensioningConstants
    Public Const OFFSET_DIM As Double = 0.012
    ''' <summary>Reservado para fase diámetros; no usado en la acotación mínima actual.</summary>
    Public Const MIN_DIAMETER As Double = 0.003
End Module

''' <summary>
''' Acotación automática solo sobre la <strong>vista base</strong> del pliego (<c>IsPrimary</c> o <c>DrawingViews.Item(1)</c>).
''' Geometría 2D: <see cref="DVLine2d"/>, <see cref="DVArc2d"/>, <see cref="DVCircle2d"/> de esa vista; el placement usa el marco único <see cref="ViewPlacementFrame"/> derivado de <see cref="DrawingView.Range"/> (origen MinX, MinY de la vista base).
''' Las cotas van en <see cref="SolidEdgeFrameworkSupport.Dimensions"/> de la hoja; no se mezclan referencias con otras vistas ni con el modelo 3D directo.
''' Cotas exteriores generales: primero <see cref="RealExtremePointsResolver.TryResolveRealExtremePoints"/>; si falla, fallback a líneas extremas DVLines2d + <see cref="VerticalExteriorAnchors"/>.
''' </summary>
Public NotInheritable Class DimensioningEngine

    Private Sub New()
    End Sub

    ''' <summary>
    ''' Punto de entrada recomendado: acota usando la <see cref="DrawingView"/> indicada y el borrador que la contiene.
    ''' Tras insertar o actualizar vistas, puede llamarse una vez la vista principal ya construida (evita cotas duplicadas si se repite en cada vista intermedia).
    ''' </summary>
    Public Shared Sub ApplyAutoDimensioning(view As DrawingView, Optional appLogger As Logger = Nothing)
        If view Is Nothing Then
            appLogger?.Log("[DIM][WARN] ApplyAutoDimensioning: DrawingView Nothing.")
            Return
        End If
        Dim draft As DraftDocument = TryGetDraftDocumentFromView(view)
        If draft Is Nothing Then
            appLogger?.Log("[DIM][ERR] ApplyAutoDimensioning: no se pudo obtener DraftDocument desde la vista (Parent/Document).")
            Return
        End If
        RunAutoDimensioning(draft, view, appLogger, Nothing)
    End Sub

    ''' <summary>Resuelve el documento de dibujo que posee la hoja de la vista (interop COM).</summary>
    Public Shared Function TryGetDraftDocumentFromView(view As DrawingView) As DraftDocument
        If view Is Nothing Then Return Nothing
        Try
            Dim parentObj As Object = Nothing
            Try
                parentObj = view.Parent
            Catch
            End Try
            If TypeOf parentObj Is Sheet Then
                Return TryGetDraftFromSheet(CType(parentObj, Sheet))
            End If
            If TypeOf parentObj Is DrawingViews Then
                Dim dvs As DrawingViews = CType(parentObj, DrawingViews)
                Dim shObj As Object = Nothing
                Try
                    shObj = dvs.Parent
                Catch
                End Try
                If TypeOf shObj Is Sheet Then
                    Return TryGetDraftFromSheet(CType(shObj, Sheet))
                End If
            End If
        Catch
        End Try
        Return Nothing
    End Function

    Private Shared Function TryGetDraftFromSheet(sh As Sheet) As DraftDocument
        If sh Is Nothing Then Return Nothing
        Dim docObj As Object = Nothing
        Try
            docObj = sh.Parent
        Catch
        End Try
        If TypeOf docObj Is DraftDocument Then Return CType(docObj, DraftDocument)
        Try
            docObj = CallByName(sh, "Document", CallType.Get)
        Catch
        End Try
        If TypeOf docObj Is DraftDocument Then Return CType(docObj, DraftDocument)
        Return Nothing
    End Function

    Public Shared Sub RunAutoDimensioning(draft As DraftDocument, mainView As DrawingView)
        RunAutoDimensioning(draft, mainView, Nothing, Nothing)
    End Sub

    ''' <param name="mainView">Vista base del pipeline (vBase); si es Nothing se intenta resolver en la hoja.</param>
    Public Shared Sub RunAutoDimensioning(draft As DraftDocument, mainView As DrawingView, appLogger As Logger)
        RunAutoDimensioning(draft, mainView, appLogger, Nothing)
    End Sub

    ''' <param name="norm">UNE-EN ISO 129-1 / separación y deduplicación entre vistas; Nothing = <see cref="DimensioningNormConfig.DefaultConfig"/>.</param>
    Public Shared Sub RunAutoDimensioning(draft As DraftDocument, mainView As DrawingView, appLogger As Logger, norm As DimensioningNormConfig)
        RunAutoDimensioning(draft, mainView, appLogger, norm, Nothing)
    End Sub

    ''' <param name="protectedZones">Zonas prohibidas (PartsList superior, cajetín, etc.) para UNE129 y layout.</param>
    Public Shared Sub RunAutoDimensioning(draft As DraftDocument, mainView As DrawingView, appLogger As Logger, norm As DimensioningNormConfig, protectedZones As IList(Of ProtectedZone2D))
        Dim log As New DimensionLogger(appLogger)
        Try
            RunInternal(draft, mainView, log, appLogger, norm, protectedZones)
        Catch ex As Exception
            log.ComFail("RunAutoDimensioning (global)", "DimensioningEngine", ex)
        Finally
            log.Info("Fin acotado")
        End Try
    End Sub

    Private Shared Sub RunInternal(draft As DraftDocument, mainView As DrawingView, log As DimensionLogger, baseLogger As Logger, norm As DimensioningNormConfig, Optional protectedZones As IList(Of ProtectedZone2D) = Nothing)
        ' Motor único de acotado: se elimina cualquier flujo alternativo/experimental.
        UniqueDvAutoDimensioningEngine.Run(draft, log, baseLogger, norm, protectedZones)
        Return
#If False Then
        ' Colocación respecto al marco de la vista que se acota (GeoFrame): evita cotas de una vista
        ' baja colocadas encima de las superiores (líneas de extensión atravesando piezas).
        ' Referencia global en [DIM][FRAME] sigue siendo la vista base; OFFSET/STEP usan Height/Width de cada vista.

        Dim horizontalOkCount As Integer = 0
        Dim globalHIdx As Integer = 0
        ' Agrupar por bbox de hoja del marco (no por DrawingView): igualdad COM/RCW no es fiable.
        For Each hGrp In hJobs.GroupBy(Function(j) SheetBoxKey(j.GeoFrame))
            Dim firstJob As HorizontalPlacementJob = hGrp.FirstOrDefault()
            If firstJob Is Nothing OrElse firstJob.GeoFrame Is Nothing Then Continue For
            Dim gf As ViewPlacementFrame = firstJob.GeoFrame
            Dim offsetH_Y As Double = gf.Height * 0.05R
            Dim stepH_Y As Double = gf.Height * 0.04R
            Dim listH As List(Of HorizontalPlacementJob) = hGrp.OrderBy(Function(j) j.SortKeyX).ToList()
            For k As Integer = 0 To listH.Count - 1
                Dim job As HorizontalPlacementJob = listH(k)
                Dim yRaw As Double = gf.MaxY + offsetH_Y + k * stepH_Y
                Dim yDim As Double = CapHorizontalDimensionLineY(yRaw, gf, targetSheetBoxes, log)
                Dim xLog As Double = job.SortKeyX
                If job.Kind = 0 AndAlso job.LeftE IsNot Nothing AndAlso job.RightE IsNot Nothing Then
                    xLog = (job.LeftE.XSheet + job.RightE.XSheet) / 2.0R
                ElseIf job.Kind = 1 AndAlso job.LeftLine IsNot Nothing AndAlso job.RightLine IsNot Nothing Then
                    xLog = (job.LeftLine.MidX + job.RightLine.MidX) / 2.0R
                End If
                log.LogLine(String.Format(CultureInfo.InvariantCulture, "[DIM][PLACE][H] index={0} X={1:0.######} Y={2:0.######}", globalHIdx, xLog, yDim))
                globalHIdx += 1

                Dim yN As Nullable(Of Double) = yDim
                If job.Kind = 0 Then
                    If DimensionPlacementEngine.TryInsertHorizontalExteriorFromExtremePoints(
                        job.Dims, job.LeftE, job.RightE, job.TopE, job.Dv, job.GeoFrame, log, yN) Then
                        horizontalOkCount += 1
                    End If
                ElseIf job.Kind = 1 Then
                    If DimensionPlacementEngine.TryInsertHorizontalBetweenLines(
                        job.Dims, job.LeftLine, job.RightLine, job.GeoFrame, log, job.Dv, yN) Then
                        horizontalOkCount += 1
                    End If
                End If
            Next
        Next

        Dim verticalOkCount As Integer = 0
        Dim globalVIdx As Integer = 0
        For Each vGrp In vJobs.GroupBy(Function(j) SheetBoxKey(j.GeoFrame))
            Dim firstV As VerticalPlacementJob = vGrp.FirstOrDefault()
            If firstV Is Nothing OrElse firstV.GeoFrame Is Nothing Then Continue For
            Dim gfv As ViewPlacementFrame = firstV.GeoFrame
            Dim offsetV_X As Double = gfv.Width * 0.05R
            Dim stepV_X As Double = gfv.Width * 0.04R
            Dim listV As List(Of VerticalPlacementJob) = vGrp.OrderBy(Function(j) j.SortKeyY).ToList()
            For k As Integer = 0 To listV.Count - 1
                Dim vjob As VerticalPlacementJob = listV(k)
                Dim xRaw As Double = gfv.MaxX + offsetV_X + k * stepV_X
                Dim xDim As Double = CapVerticalDimensionLineX(xRaw, gfv, targetSheetBoxes, log)
                Dim yLog As Double = vjob.SortKeyY
                If vjob.Kind = 0 AndAlso vjob.BottomE IsNot Nothing AndAlso vjob.TopE IsNot Nothing Then
                    yLog = (vjob.BottomE.YSheet + vjob.TopE.YSheet) / 2.0R
                ElseIf vjob.Kind = 1 AndAlso vjob.Extreme IsNot Nothing AndAlso
                    vjob.Extreme.BottomHorizontal IsNot Nothing AndAlso vjob.Extreme.TopHorizontal IsNot Nothing Then
                    yLog = (vjob.Extreme.BottomHorizontal.MidY + vjob.Extreme.TopHorizontal.MidY) / 2.0R
                End If
                log.LogLine(String.Format(CultureInfo.InvariantCulture, "[DIM][PLACE][V] index={0} X={1:0.######} Y={2:0.######}", globalVIdx, xDim, yLog))
                globalVIdx += 1

                Dim xN As Nullable(Of Double) = xDim
                If vjob.Kind = 0 Then
                    If DimensionPlacementEngine.TryInsertVerticalExteriorFromExtremePoints(
                        vjob.Dims, vjob.BottomE, vjob.TopE, vjob.RightE, vjob.Dv, vjob.GeoFrame, log, xN) Then
                        verticalOkCount += 1
                    End If
                ElseIf vjob.Kind = 1 Then
                    If DimensionPlacementEngine.TryInsertVerticalExterior(vjob.Dims, vjob.Dv, vjob.GeoFrame, vjob.Extreme, log, xN) Then
                        verticalOkCount += 1
                    End If
                End If
            Next
        Next

        For ci As Integer = 0 To targets.Count - 1
            Dim tv As DrawingView = targets(ci)
            If tv Is Nothing Then Continue For
            Dim cbox As New ViewSheetBoundingBox()
            If Not ViewGeometryReader.TryReadBoundingBox(tv, log, cbox) Then Continue For
            Dim cframe As ViewPlacementFrame = Nothing
            If Not ViewPlacementFrame.TryCreateFromBaseViewSheetBox(cbox, log, cframe) OrElse cframe Is Nothing Then Continue For
            Dim csh As SolidEdgeDraft.Sheet = Nothing
            Try
                csh = CType(tv.Sheet, SolidEdgeDraft.Sheet)
            Catch
            End Try
            If csh Is Nothing Then Continue For
            Dim cdims As SolidEdgeFrameworkSupport.Dimensions = Nothing
            Try
                cdims = CType(csh.Dimensions, SolidEdgeFrameworkSupport.Dimensions)
            Catch
            End Try
            If cdims Is Nothing Then Continue For
            Dim cc As Integer = TryInsertCircularDimensionsForView(tv, cdims, cframe, log)
            log.Info("[DIM][RESULT][VIEW] horizontal_ok=(batched) vertical_ok=(batched) circular_count=" & cc.ToString(CultureInfo.InvariantCulture))
        Next

        log.Info("[DIM][RESULT] horizontal_placed=" & horizontalOkCount.ToString(CultureInfo.InvariantCulture) &
                 " vertical_placed=" & verticalOkCount.ToString(CultureInfo.InvariantCulture))
#End If
    End Sub

    Private NotInheritable Class HorizontalPlacementJob
        Public Kind As Integer
        Public Dims As SolidEdgeFrameworkSupport.Dimensions
        Public Dv As DrawingView
        Public GeoFrame As ViewPlacementFrame
        Public SortKeyX As Double
        Public LeftE As DimensionExtremePoint
        Public RightE As DimensionExtremePoint
        Public TopE As DimensionExtremePoint
        Public LeftLine As DvLineSheetInfo
        Public RightLine As DvLineSheetInfo
    End Class

    Private NotInheritable Class VerticalPlacementJob
        Public Kind As Integer
        Public Dims As SolidEdgeFrameworkSupport.Dimensions
        Public Dv As DrawingView
        Public GeoFrame As ViewPlacementFrame
        Public SortKeyY As Double
        Public BottomE As DimensionExtremePoint
        Public TopE As DimensionExtremePoint
        Public RightE As DimensionExtremePoint
        Public Extreme As ExtremeDvLinesResult
    End Class

    Private Shared Function SheetBoxKey(frame As ViewPlacementFrame) As String
        If frame Is Nothing Then Return vbNullString
        Return String.Format(CultureInfo.InvariantCulture,
            "{0:0.##########}|{1:0.##########}|{2:0.##########}|{3:0.##########}",
            frame.MinX, frame.MinY, frame.MaxX, frame.MaxY)
    End Function

    ''' <summary>
    ''' Limita Y de la línea de cota horizontal para que no invada el rectángulo de una vista situada más arriba en la hoja
    ''' (coordenadas de hoja con Y creciente hacia arriba), reduciendo extensiones que atraviesan vistas intermedias.
    ''' </summary>
    Private Shared Function CapHorizontalDimensionLineY(
        yDraft As Double,
        gf As ViewPlacementFrame,
        allTargetBoxes As IList(Of ViewSheetBoundingBox),
        log As DimensionLogger) As Double

        If gf Is Nothing OrElse allTargetBoxes Is Nothing OrElse allTargetBoxes.Count = 0 Then Return yDraft
        Dim clearance As Double = Math.Max(gf.Height * 0.02R, 0.0015R)
        Dim capTop As Double = Double.PositiveInfinity
        For Each b In allTargetBoxes
            If b.MinY > gf.MaxY + 1.0E-7R Then
                capTop = Math.Min(capTop, b.MinY - clearance)
            End If
        Next
        If Double.IsPositiveInfinity(capTop) Then Return yDraft
        If capTop <= gf.MaxY Then Return yDraft
        Dim yC As Double = Math.Min(yDraft, capTop)
        Dim yFloor As Double = gf.MaxY + Math.Max(gf.Height * 0.015R, 0.001R)
        If yC < yFloor Then Return yDraft
        If Math.Abs(yC - yDraft) > 1.0E-6R AndAlso log IsNot Nothing Then
            log.LogLine(String.Format(CultureInfo.InvariantCulture,
                "[DIM][PLACE][H][CAP] Y_draft={0:0.######} Y_capped={1:0.######} (vista encima)",
                yDraft, yC))
        End If
        Return yC
    End Function

    ''' <summary>Análogo en X para cotas verticales y una vista inmediatamente a la derecha.</summary>
    Private Shared Function CapVerticalDimensionLineX(
        xDraft As Double,
        gf As ViewPlacementFrame,
        allTargetBoxes As IList(Of ViewSheetBoundingBox),
        log As DimensionLogger) As Double

        If gf Is Nothing OrElse allTargetBoxes Is Nothing OrElse allTargetBoxes.Count = 0 Then Return xDraft
        Dim clearance As Double = Math.Max(gf.Width * 0.02R, 0.0015R)
        Dim capLeft As Double = Double.PositiveInfinity
        For Each b In allTargetBoxes
            If b.MinX > gf.MaxX + 1.0E-7R Then
                capLeft = Math.Min(capLeft, b.MinX - clearance)
            End If
        Next
        If Double.IsPositiveInfinity(capLeft) Then Return xDraft
        If capLeft <= gf.MaxX Then Return xDraft
        Dim xC As Double = Math.Min(xDraft, capLeft)
        Dim xFloor As Double = gf.MaxX + Math.Max(gf.Width * 0.015R, 0.001R)
        If xC < xFloor Then Return xDraft
        If Math.Abs(xC - xDraft) > 1.0E-6R AndAlso log IsNot Nothing Then
            log.LogLine(String.Format(CultureInfo.InvariantCulture,
                "[DIM][PLACE][V][CAP] X_draft={0:0.######} X_capped={1:0.######} (vista derecha)",
                xDraft, xC))
        End If
        Return xC
    End Function

    Private Shared Function ComputeVerticalFallbackSortY(extreme As ExtremeDvLinesResult, box As ViewSheetBoundingBox) As Double
        If extreme IsNot Nothing AndAlso extreme.BottomHorizontal IsNot Nothing AndAlso extreme.TopHorizontal IsNot Nothing Then
            Return Math.Min(extreme.BottomHorizontal.MidY, extreme.TopHorizontal.MidY)
        End If
        Return (box.MinY + box.MaxY) / 2.0R
    End Function

    Private Shared Sub CollectPlacementJobsForView(
        targetView As DrawingView,
        baseView As DrawingView,
        log As DimensionLogger,
        baseLogger As Logger,
        hJobs As List(Of HorizontalPlacementJob),
        vJobs As List(Of VerticalPlacementJob))

        If targetView Is Nothing Then Return

        Dim isBase As Boolean = baseView IsNot Nothing AndAlso Object.ReferenceEquals(targetView, baseView)

        log.Info("Vista para acotado: " & DescribeDrawingView(targetView, log))

        Dim box As New ViewSheetBoundingBox()
        If Not ViewGeometryReader.TryReadBoundingBox(targetView, log, box) Then
            log.Err("No hay bounding box de vista (DrawingView.Range); se omiten cotas para esta vista.")
            Return
        End If

        Dim placementFrame As ViewPlacementFrame = Nothing
        If Not ViewPlacementFrame.TryCreateFromBaseViewSheetBox(box, log, placementFrame) OrElse placementFrame Is Nothing Then
            log.Err("No se pudo crear ViewPlacementFrame para la vista; cotas omitidas.")
            Return
        End If

        baseLogger?.Log("[DIM] " & If(isBase, "Vista base", "Vista secundaria") &
                        " bbox: MinX=" & box.MinX.ToString("0.######", CultureInfo.InvariantCulture) &
                        " MinY=" & box.MinY.ToString("0.######", CultureInfo.InvariantCulture) &
                        " MaxX=" & box.MaxX.ToString("0.######", CultureInfo.InvariantCulture) &
                        " MaxY=" & box.MaxY.ToString("0.######", CultureInfo.InvariantCulture))

        Dim viewSheet As SolidEdgeDraft.Sheet = Nothing
        Try
            viewSheet = CType(targetView.Sheet, SolidEdgeDraft.Sheet)
        Catch
        End Try
        If viewSheet Is Nothing Then
            log.Err("No se pudo resolver DrawingView.Sheet para la vista; cotas omitidas.")
            Return
        End If

        Dim dims As SolidEdgeFrameworkSupport.Dimensions = Nothing
        Try
            dims = CType(viewSheet.Dimensions, SolidEdgeFrameworkSupport.Dimensions)
        Catch ex As Exception
            log.ComFail("CType(viewSheet.Dimensions, Dimensions)", "DrawingView.Sheet.Dimensions", ex)
            Return
        End Try
        If dims Is Nothing Then
            log.Err("DrawingView.Sheet.Dimensions es Nothing.")
            Return
        End If

        log.Info("[DIM][CTX] dimsParent=DrawingView.Sheet (usado)")
        log.Info("[DIM][CTX] ProximityCoordSpace=SHEET primero; fallback_VISTA=" &
                 DimensionInsertionConfig.UseViewSpaceProximityForAddDistance.ToString(CultureInfo.InvariantCulture))
        log.Info("[DIM][CROP][STRATEGY] dv.Range y SetUserRange descartados: la visibilidad se ajusta SOLO con CropLeft/CropRight/CropTop/CropBottom.")

        Dim leftE As DimensionExtremePoint = Nothing
        Dim rightE As DimensionExtremePoint = Nothing
        Dim bottomE As DimensionExtremePoint = Nothing
        Dim topE As DimensionExtremePoint = Nothing

        Dim haveExtremes As Boolean = False
        Dim lineReason As String = Nothing
        If DimensionInsertionConfig.PreferLineOnlyExteriorReferences OrElse
           DimensionInsertionConfig.ActiveExteriorPickStrategy = ExteriorPickStrategy.LineOnlyReferencesD Then
            If StableExteriorReferences.TryResolveLineFirstExtremes(targetView, placementFrame, log, leftE, rightE, bottomE, topE, lineReason) Then
                haveExtremes = True
                log.Info("[DIM][EXT][LINE] extremos priorizando solo DVLine2d: OK (" & lineReason & ").")
            Else
                log.Info("[DIM][EXT][LINE] solo-línea no aplicable (" & lineReason & "); resolución con geometría completa (líneas+arcos+…).")
            End If
        End If

        If Not haveExtremes Then
            haveExtremes = RealExtremePointsResolver.TryResolveRealExtremePoints(targetView, placementFrame, box, log, leftE, rightE, bottomE, topE)
        End If

        If haveExtremes Then
            baseLogger?.Log("[DIM][PLACE][VIEW] H+V desde EXTPT (encolado)")
            hJobs.Add(New HorizontalPlacementJob With {
                .Kind = 0,
                .Dims = dims,
                .Dv = targetView,
                .GeoFrame = placementFrame,
                .SortKeyX = Math.Min(leftE.XSheet, rightE.XSheet),
                .LeftE = leftE,
                .RightE = rightE,
                .TopE = topE
            })
            vJobs.Add(New VerticalPlacementJob With {
                .Kind = 0,
                .Dims = dims,
                .Dv = targetView,
                .GeoFrame = placementFrame,
                .SortKeyY = Math.Min(bottomE.YSheet, topE.YSheet),
                .BottomE = bottomE,
                .TopE = topE,
                .RightE = rightE
            })
        Else
            log.Warn("[DIM][EXTPT] resolución de extremos fallida; fallback a líneas extremas.")
            Dim extreme As ExtremeDvLinesResult = Nothing
            If Not ViewGeometryReader.TryBuildExtremeLines(targetView, box, log, extreme) OrElse extreme Is Nothing OrElse extreme.AllLines Is Nothing Then
                log.Err("Fallback: no se pudieron leer DVLines2d de la vista; no se insertan cotas.")
                Return
            End If

            If extreme.LeftVertical IsNot Nothing AndAlso extreme.RightVertical IsNot Nothing Then
                hJobs.Add(New HorizontalPlacementJob With {
                    .Kind = 1,
                    .Dims = dims,
                    .Dv = targetView,
                    .GeoFrame = placementFrame,
                    .SortKeyX = Math.Min(extreme.LeftVertical.MidX, extreme.RightVertical.MidX),
                    .LeftLine = extreme.LeftVertical,
                    .RightLine = extreme.RightVertical
                })
            Else
                log.Err("Cota horizontal omitida (fallback): faltan verticales extremas.")
            End If

            vJobs.Add(New VerticalPlacementJob With {
                .Kind = 1,
                .Dims = dims,
                .Dv = targetView,
                .GeoFrame = placementFrame,
                .SortKeyY = ComputeVerticalFallbackSortY(extreme, box),
                .Extreme = extreme
            })
        End If
    End Sub

    Private Shared Function UnusedViewBoxForStrictPairPlacement() As ViewSheetBoundingBox
        Return New ViewSheetBoundingBox()
    End Function

    Private Shared Function BuildDimensionTargetViews(sheet As Sheet, baseView As DrawingView, log As DimensionLogger, appLogger As Logger) As List(Of DrawingView)
        Dim result As New List(Of DrawingView)()
        If sheet Is Nothing OrElse baseView Is Nothing Then Return result

        ' Base primero.
        result.Add(baseView)

        Dim views As DrawingViews = Nothing
        Try
            views = sheet.DrawingViews
        Catch ex As Exception
            log?.ComFail("Sheet.DrawingViews", "DrawingViews", ex)
            Return result
        End Try
        If views Is Nothing Then Return result

        Dim n As Integer = 0
        Try
            n = views.Count
        Catch ex As Exception
            log?.ComFail("DrawingViews.Count", "DrawingViews", ex)
            Return result
        End Try

        For i As Integer = 1 To n
            Dim dv As DrawingView = Nothing
            Try
                dv = CType(views.Item(i), DrawingView)
            Catch
                Continue For
            End Try
            If dv Is Nothing Then Continue For
            If Object.ReferenceEquals(dv, baseView) Then Continue For

            If IsIsometricDrawingView(dv, log) Then
                appLogger?.Log("[DIM] Vista omitida (isométrica): índice " & i.ToString(CultureInfo.InvariantCulture))
                Continue For
            End If

            result.Add(dv)
        Next

        Return result
    End Function

    Private Shared Function IsIsometricDrawingView(dv As DrawingView, log As DimensionLogger) As Boolean
        If dv Is Nothing Then Return False
        Try
            Dim vx As Double, vy As Double, vz As Double, lx As Double, ly As Double, lz As Double
            Dim ori As SolidEdgeConstants.ViewOrientationConstants = SolidEdgeConstants.ViewOrientationConstants.igFrontView
            dv.ViewOrientation(vx, vy, vz, lx, ly, lz, ori)
            Dim s As String = ori.ToString()
            If s.IndexOf("TopFrontRight", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
               s.IndexOf("TopFrontLeft", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
               s.IndexOf("TopBackRight", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
               s.IndexOf("TopBackLeft", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
               s.IndexOf("BottomFrontRight", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
               s.IndexOf("BottomFrontLeft", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
               s.IndexOf("BottomBackRight", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
               s.IndexOf("BottomBackLeft", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
               s.IndexOf("Iso", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
               s.IndexOf("Isometric", StringComparison.OrdinalIgnoreCase) >= 0 Then
                Return True
            End If
        Catch
        End Try
        Return False
    End Function

    Private Shared Function TryInsertCircularDimensionsForView(targetView As DrawingView, dims As SolidEdgeFrameworkSupport.Dimensions, frame As ViewPlacementFrame, log As DimensionLogger) As Integer
        If targetView Is Nothing OrElse dims Is Nothing OrElse frame Is Nothing Then Return 0

        Dim inserted As Integer = 0
        Dim clearance As Double = Math.Max(0.003R, Math.Min(0.008R, Math.Max(frame.Width, frame.Height) * 0.06R))

        ' Círculos -> diámetro
        Try
            Dim circles As DVCircles2d = targetView.DVCircles2d
            If circles IsNot Nothing Then
                Dim n As Integer = 0
                Try : n = circles.Count : Catch : n = 0 : End Try
                For i As Integer = 1 To n
                    Dim c As DVCircle2d = Nothing
                    Try : c = CType(circles.Item(i), DVCircle2d) : Catch : c = Nothing : End Try
                    If c Is Nothing Then Continue For

                    Dim cx As Double = 0, cy As Double = 0, r As Double = 0
                    If Not TryGetCircleCenterRadiusLocal(c, cx, cy, r) Then Continue For
                    If r < 0.0008R Then Continue For

                    Dim px As Double = cx + r + clearance
                    Dim py As Double = cy + r * 0.35R
                    Dim sx As Double = 0, sy As Double = 0
                    Try
                        targetView.ViewToSheet(px, py, sx, sy)
                    Catch
                        Continue For
                    End Try

                    If TryCallCircularDimensionMethod(dims, c, sx, sy, True, log) Then
                        inserted += 1
                    End If
                Next
            End If
        Catch ex As Exception
            log?.ComFail("DrawingView.DVCircles2d", "DrawingView", ex)
        End Try

        ' Arcos -> radio
        Try
            Dim arcs As DVArcs2d = targetView.DVArcs2d
            If arcs IsNot Nothing Then
                Dim n As Integer = 0
                Try : n = arcs.Count : Catch : n = 0 : End Try
                For i As Integer = 1 To n
                    Dim a As DVArc2d = Nothing
                    Try : a = CType(arcs.Item(i), DVArc2d) : Catch : a = Nothing : End Try
                    If a Is Nothing Then Continue For

                    Dim cx As Double = 0, cy As Double = 0, r As Double = 0
                    If Not TryGetArcCenterRadiusLocal(a, cx, cy, r) Then Continue For
                    If r < 0.0008R Then Continue For

                    Dim px As Double = cx + r + clearance
                    Dim py As Double = cy + r * 0.2R
                    Dim sx As Double = 0, sy As Double = 0
                    Try
                        targetView.ViewToSheet(px, py, sx, sy)
                    Catch
                        Continue For
                    End Try

                    If TryCallCircularDimensionMethod(dims, a, sx, sy, False, log) Then
                        inserted += 1
                    End If
                Next
            End If
        Catch ex As Exception
            log?.ComFail("DrawingView.DVArcs2d", "DrawingView", ex)
        End Try

        log?.Info("[DIM][CIRC] cotas circulares insertadas=" & inserted.ToString(CultureInfo.InvariantCulture))
        Return inserted
    End Function

    Private Shared Function TryCallCircularDimensionMethod(dims As SolidEdgeFrameworkSupport.Dimensions, entity As Object, xSheet As Double, ySheet As Double, preferDiameter As Boolean, log As DimensionLogger) As Boolean
        If dims Is Nothing OrElse entity Is Nothing Then Return False

        Dim names As String() = If(preferDiameter,
            New String() {"AddDiameter", "AddDiameterByCircle", "AddDiameterDimension", "AddDiameter2"},
            New String() {"AddRadius", "AddRadiusByArc", "AddRadial", "AddRadialDimension", "AddRadius2"})

        For Each m In names
            If TryInvokeCircularAdd(dims, m, entity, xSheet, ySheet, log) Then
                log?.Info("[DIM][CIRC] método=" & m & " ok")
                Return True
            End If
        Next

        log?.Warn("[DIM][CIRC] no se pudo insertar " & If(preferDiameter, "diámetro", "radio") & " para entidad=" & SafeTypeName(entity))
        Return False
    End Function

    Private Shared Function TryInvokeCircularAdd(dims As Object, methodName As String, entity As Object, xSheet As Double, ySheet As Double, log As DimensionLogger) As Boolean
        Try
            CallByName(dims, methodName, CallType.Method, entity, xSheet, ySheet)
            Return True
        Catch
        End Try
        Try
            CallByName(dims, methodName, CallType.Method, entity, xSheet, ySheet, 0)
            Return True
        Catch
        End Try
        Try
            CallByName(dims, methodName, CallType.Method, entity, xSheet, ySheet, 0, False)
            Return True
        Catch ex As Exception
            log?.Info("[DIM][CIRC] " & methodName & " falló: " & ex.Message)
        End Try
        Return False
    End Function
    ''' <summary>Líneas horizontales estructurales (hoja): descarta microsegmentos y remates cortos.</summary>
    Private Shared Function CollectStructuralHorizontalLines(
        geom As ViewGeometrySummary,
        box As ViewSheetBoundingBox,
        logger As Logger,
        rejectTag As String) As List(Of LineGeometryInfo)

        Dim lst As New List(Of LineGeometryInfo)()
        If geom?.Lines Is Nothing Then Return lst
        ' Sin filtro de seguridad por tamaño: permitir también segmentos pequeños.
        Dim minLen As Double = 0.0R
        For Each ln In geom.Lines
            If Not String.Equals(ln.Orientation, "horizontal", StringComparison.OrdinalIgnoreCase) Then Continue For
            If ln.Length < minLen Then
                logger?.Log("[DIM][" & rejectTag & "][REJECT] DVLine2d#" & ln.Index.ToString(CultureInfo.InvariantCulture) &
                            " motivo=microsegmento_o_remate_corto L=" & ln.Length.ToString("0.######", CultureInfo.InvariantCulture) &
                            " min_requerido=" & minLen.ToString("0.######", CultureInfo.InvariantCulture))
                Continue For
            End If
            lst.Add(ln)
        Next
        Return lst.OrderByDescending(Function(l) l.Length).ToList()
    End Function

    ''' <summary>Selección explícita OVERALL_WIDTH: horizontales dominantes del contorno (no xmin/xmax globales).</summary>
    Private Shared Function SelectEntitiesForOverallWidth(
        dv As DrawingView,
        geom As ViewGeometrySummary,
        box As ViewSheetBoundingBox,
        logger As Logger,
        ByRef leftObj As Object,
        ByRef rightObj As Object,
        ByRef leftPickX As Double,
        ByRef rightPickX As Double,
        ByRef widthValue As Double) As Boolean

        leftObj = Nothing : rightObj = Nothing
        leftPickX = 0 : rightPickX = 0 : widthValue = 0
        If dv Is Nothing OrElse geom Is Nothing OrElse logger Is Nothing Then Return False

        Dim pool = CollectStructuralHorizontalLines(geom, box, logger, "WIDTH")
        If pool.Count = 0 Then
            logger.Log("[DIM][WIDTH][REJECT] motivo=sin_horizontales_estructurales")
            Return False
        End If

        Dim maxLen As Double = pool(0).Length
        Dim frac As Double = 0.72R
        Dim dominant = pool.Where(Function(l) l.Length >= maxLen * frac).ToList()
        If dominant.Count < 2 Then
            dominant = pool.Take(Math.Min(8, pool.Count)).ToList()
        End If

        Dim nShow As Integer = Math.Min(12, dominant.Count)
        For i As Integer = 0 To nShow - 1
            Dim ln = dominant(i)
            logger.Log("[DIM][WIDTH][CANDIDATE] DVLine2d#" & ln.Index.ToString(CultureInfo.InvariantCulture) &
                       " L=" & ln.Length.ToString("0.######", CultureInfo.InvariantCulture) &
                       " midY=" & ln.MidY.ToString("0.######", CultureInfo.InvariantCulture) &
                       " x=[" & ln.BboxMinX.ToString("0.######", CultureInfo.InvariantCulture) & ".." & ln.BboxMaxX.ToString("0.######", CultureInfo.InvariantCulture) & "]")
        Next

        Dim unionMinX As Double = dominant.Min(Function(l) l.BboxMinX)
        Dim unionMaxX As Double = dominant.Max(Function(l) l.BboxMaxX)

        Dim leftOrdered = dominant.OrderBy(Function(l) l.BboxMinX).ThenByDescending(Function(l) l.Length).ToList()
        Dim rightOrdered = dominant.OrderByDescending(Function(l) l.BboxMaxX).ThenByDescending(Function(l) l.Length).ToList()

        If leftOrdered.Count = 0 OrElse rightOrdered.Count = 0 Then
            logger.Log("[DIM][WIDTH][REJECT] motivo=no_hay_extremos_izq_der_en_cluster")
            Return False
        End If

        Dim leftInfo As LineGeometryInfo = leftOrdered(0)
        Dim rightInfo As LineGeometryInfo = rightOrdered(0)
        If leftInfo.Index = rightInfo.Index Then
            If rightOrdered.Count > 1 AndAlso rightOrdered(1).Index <> leftInfo.Index Then
                rightInfo = rightOrdered(1)
            ElseIf leftOrdered.Count > 1 AndAlso leftOrdered(1).Index <> rightInfo.Index Then
                leftInfo = leftOrdered(1)
            Else
                logger.Log("[DIM][WIDTH][REJECT] motivo=una_sola_linea_domina_ambos_extremos")
                Return False
            End If
        End If

        Dim linesCol As DVLines2d = Nothing
        Try
            linesCol = dv.DVLines2d
        Catch
        End Try
        If linesCol Is Nothing Then Return False
        leftObj = TryGetLineItemSafe(linesCol, leftInfo.Index)
        rightObj = TryGetLineItemSafe(linesCol, rightInfo.Index)
        If leftObj Is Nothing OrElse rightObj Is Nothing Then
            logger.Log("[DIM][WIDTH][REJECT] motivo=COM_DVLine2d_no_resuelto")
            Return False
        End If

        leftPickX = unionMinX
        rightPickX = unionMaxX
        widthValue = Math.Abs(unionMaxX - unionMinX)
        If widthValue <= 0.0005R Then
            logger.Log("[DIM][WIDTH][REJECT] motivo=ancho_degenerado")
            Return False
        End If

        logger.Log("[DIM][WIDTH][ACCEPT] entidad principal izq=DVLine2d#" & leftInfo.Index.ToString(CultureInfo.InvariantCulture) &
                   " der=DVLine2d#" & rightInfo.Index.ToString(CultureInfo.InvariantCulture))
        logger.Log("[DIM][WIDTH][RESULT] usando horizontales dominantes del contorno principal")
        Return True
    End Function

    ''' <summary>Selección explícita OVERALL_HEIGHT: horizontal inferior dominante vs horizontal superior dominante (no dos verticales del mismo lateral).</summary>
    Private Shared Function SelectEntitiesForOverallHeight(
        dv As DrawingView,
        geom As ViewGeometrySummary,
        box As ViewSheetBoundingBox,
        logger As Logger,
        ByRef bottomObj As Object,
        ByRef topObj As Object,
        ByRef bottomPickY As Double,
        ByRef topPickY As Double,
        ByRef heightValue As Double) As Boolean

        bottomObj = Nothing : topObj = Nothing
        bottomPickY = 0 : topPickY = 0 : heightValue = 0
        If dv Is Nothing OrElse geom Is Nothing OrElse logger Is Nothing Then Return False

        Dim pool = CollectStructuralHorizontalLines(geom, box, logger, "HEIGHT")
        If pool.Count < 2 Then
            logger.Log("[DIM][HEIGHT][REJECT] motivo=menos_de_dos_horizontales_estructurales")
            Return False
        End If

        Dim poolSpan = pool.Take(Math.Min(16, pool.Count)).ToList()
        For Each ln In poolSpan
            logger.Log("[DIM][HEIGHT][CANDIDATE] DVLine2d#" & ln.Index.ToString(CultureInfo.InvariantCulture) &
                       " L=" & ln.Length.ToString("0.######", CultureInfo.InvariantCulture) &
                       " midY=" & ln.MidY.ToString("0.######", CultureInfo.InvariantCulture))
        Next

        Dim bottomInfo As LineGeometryInfo = pool.OrderBy(Function(l) l.MidY).First()
        Dim topInfo As LineGeometryInfo = pool.OrderByDescending(Function(l) l.MidY).First()
        If bottomInfo.Index = topInfo.Index Then
            logger.Log("[DIM][HEIGHT][REJECT] motivo=una_sola_horizontal_no_define_altura")
            Return False
        End If

        ' Sin separación mínima de seguridad: permitir alturas pequeñas.
        Dim minSep As Double = 0.0R
        If Math.Abs(topInfo.MidY - bottomInfo.MidY) < minSep Then
            logger.Log("[DIM][HEIGHT][REJECT] motivo=separación_vertical_insuficiente")
            Return False
        End If

        Dim linesCol As DVLines2d = Nothing
        Try
            linesCol = dv.DVLines2d
        Catch
        End Try
        If linesCol Is Nothing Then Return False
        bottomObj = TryGetLineItemSafe(linesCol, bottomInfo.Index)
        topObj = TryGetLineItemSafe(linesCol, topInfo.Index)
        If bottomObj Is Nothing OrElse topObj Is Nothing Then
            logger.Log("[DIM][HEIGHT][REJECT] motivo=COM_DVLine2d_no_resuelto")
            Return False
        End If

        bottomPickY = bottomInfo.MidY
        topPickY = topInfo.MidY
        heightValue = Math.Abs(topPickY - bottomPickY)

        logger.Log("[DIM][HEIGHT][ACCEPT] inferior=DVLine2d#" & bottomInfo.Index.ToString(CultureInfo.InvariantCulture) &
                   " superior=DVLine2d#" & topInfo.Index.ToString(CultureInfo.InvariantCulture))
        logger.Log("[DIM][HEIGHT][RESULT] usando horizontales dominante inferior/superior")
        Return True
    End Function

    ''' <summary>Delegación explícita THICKNESS (pares paralelos cercanos, lógica separada de OVERALL_HEIGHT).</summary>
    Private Shared Function SelectEntitiesForThickness(
        dv As DrawingView,
        geom As ViewGeometrySummary,
        box As ViewSheetBoundingBox,
        logger As Logger,
        ByRef outFeature As FeatureDimensionCandidate) As Boolean
        Return GeometricFeatureAnalyzer.TryDetectThicknessFeature(dv, geom, box, logger, outFeature)
    End Function

    Private Shared Function SafeTypeName(obj As Object) As String
        If obj Is Nothing Then Return "(Nothing)"
        Try
            Return obj.GetType().Name
        Catch
            Return "(tipo no legible)"
        End Try
    End Function

    Private Class FeatureDimensionCandidate
        Public FeatureType As String
        Public Object1 As Object
        Public Object2 As Object
        Public Axis As String ' "horizontal" / "vertical"
        Public Pick1X As Double
        Public Pick1Y As Double
        Public Pick2X As Double
        Public Pick2Y As Double
        Public Value As Double
        Public Confidence As Double
        Public RealPoint1 As SignificantPoint
        Public RealPoint2 As SignificantPoint
    End Class

    Private Enum SignificantPointKind
        StartPoint = 1
        EndPoint = 2
        MidPoint = 3
        CenterPoint = 4
        QuadrantTop = 5
        QuadrantBottom = 6
        QuadrantLeft = 7
        QuadrantRight = 8
        TangencyCandidate = 9
    End Enum

    Private Enum SignificantPointCoordSystem
        LocalEntity = 1
        View = 2
        Sheet = 3
    End Enum

    Private Class SignificantPoint
        Public Kind As SignificantPointKind
        Public X As Double
        Public Y As Double
        Public Owner As Object
        Public OwnerType As String
        Public Tag As String
        Public CoordSystem As SignificantPointCoordSystem
    End Class

    Private Class VerticalFunctionalCandidate
        Public Kind As String
        Public Index1Based As Integer
        Public ObjType As String
        Public Label As String
        Public PickX As Double
        Public PickY As Double
        Public MinX As Double
        Public MaxX As Double
        Public MinY As Double
        Public MaxY As Double
        Public Orientation As String
        Public Length As Double
        Public Score As Integer
        Public DiscardReason As String
    End Class

    Private Shared Function TryResolveFunctionalVerticalPair(
        dv As DrawingView,
        geom As ViewGeometrySummary,
        box As ViewSheetBoundingBox,
        logger As Logger,
        ByRef bottomObj As Object,
        ByRef topObj As Object,
        ByRef bottomPickY As Double,
        ByRef topPickY As Double,
        ByRef candidatesOut As List(Of VerticalFunctionalCandidate)) As Boolean

        bottomObj = Nothing : topObj = Nothing
        bottomPickY = 0 : topPickY = 0
        candidatesOut = New List(Of VerticalFunctionalCandidate)()
        If dv Is Nothing OrElse geom Is Nothing OrElse logger Is Nothing Then Return False

        ' Reducir first-chance COM/E_POINTER: actualizar y cachear colecciones una sola vez.
        Try
            dv.Update()
        Catch
        End Try

        Dim linesCol As DVLines2d = Nothing
        Dim arcsCol As DVArcs2d = Nothing
        Try
            linesCol = dv.DVLines2d
        Catch ex As Exception
            logger.Log("[DIM][VERT][FILTER] DVLines2d no accesible: " & ex.Message)
        End Try
        Try
            arcsCol = dv.DVArcs2d
        Catch ex As Exception
            logger.Log("[DIM][VERT][FILTER] DVArcs2d no accesible: " & ex.Message)
        End Try

        ' Quitar filtros de seguridad para permitir también rasgos pequeños.
        Dim minLenMicro As Double = 0.0R
        Dim minLenStructV As Double = 0.0R
        Dim minLenMainH As Double = 0.0R
        Dim minArcRadius As Double = 0.0R
        Dim minArcLen As Double = 0.0R
        Dim tolConnY As Double = Math.Max(0.0015R, box.Height * 0.05R)

        Dim kept As New List(Of VerticalFunctionalCandidate)()
        Dim discarded As New List(Of VerticalFunctionalCandidate)()

        ' 1) Líneas: prioridad principal para altura funcional = caras horizontales superior/inferior.
        For Each ln In geom.Lines
            Dim cand As New VerticalFunctionalCandidate With {
                .Kind = "line",
                .Index1Based = ln.Index,
                .ObjType = "DVLine2d",
                .Label = "DVLine2d#" & ln.Index.ToString(CultureInfo.InvariantCulture),
                .PickX = ln.MidX,
                .PickY = ln.MidY,
                .MinX = ln.BboxMinX,
                .MaxX = ln.BboxMaxX,
                .MinY = ln.BboxMinY,
                .MaxY = ln.BboxMaxY,
                .Orientation = ln.Orientation,
                .Length = ln.Length,
                .Score = 99
            }
            If ln.Length < minLenMicro Then
                cand.DiscardReason = "micro segmento/remate"
                discarded.Add(cand)
                Continue For
            End If

            If String.Equals(ln.Orientation, "horizontal", StringComparison.OrdinalIgnoreCase) AndAlso ln.Length >= minLenMainH Then
                cand.Score = 1
                kept.Add(cand)
            ElseIf String.Equals(ln.Orientation, "vertical", StringComparison.OrdinalIgnoreCase) AndAlso ln.Length >= minLenStructV Then
                ' No descartar automáticamente por X extremo: puede ser contorno funcional.
                Dim touchesMinY As Boolean = Math.Abs(ln.BboxMinY - geom.EntityUnionMinY) <= tolConnY
                Dim touchesMaxY As Boolean = Math.Abs(ln.BboxMaxY - geom.EntityUnionMaxY) <= tolConnY
                Dim spanRatio As Double = If(box.Height > 0.000001R, (ln.BboxMaxY - ln.BboxMinY) / box.Height, 0R)
                If touchesMinY OrElse touchesMaxY OrElse spanRatio >= 0.55R Then
                    cand.Score = 2
                    kept.Add(cand)
                Else
                    cand.Score = 3
                    kept.Add(cand)
                End If
            Else
                cand.DiscardReason = "línea no principal para altura funcional"
                discarded.Add(cand)
            End If
        Next

        ' 2) Arcos: solo si no son radios de transición pequeños.
        For Each ar In geom.Arcs
            Dim cand As New VerticalFunctionalCandidate With {
                .Kind = "arc",
                .Index1Based = ar.Index,
                .ObjType = "DVArc2d",
                .Label = "DVArc2d#" & ar.Index.ToString(CultureInfo.InvariantCulture),
                .PickX = (ar.BboxMinX + ar.BboxMaxX) / 2.0R,
                .PickY = (ar.BboxMinY + ar.BboxMaxY) / 2.0R,
                .MinX = ar.BboxMinX,
                .MaxX = ar.BboxMaxX,
                .MinY = ar.BboxMinY,
                .MaxY = ar.BboxMaxY,
                .Orientation = "curve",
                .Length = ar.ArcLengthApprox,
                .Score = 4
            }
            If ar.Radius < minArcRadius OrElse ar.ArcLengthApprox < minArcLen Then
                cand.DiscardReason = "arco/radio de transición pequeño"
                discarded.Add(cand)
                Continue For
            End If
            kept.Add(cand)
        Next

        For Each d In discarded
            logger.Log("[DIM][VERT][FILTER] descartada " & d.Label &
                       " tipo=" & d.ObjType &
                       " ymin=" & d.MinY.ToString("0.######", CultureInfo.InvariantCulture) &
                       " ymax=" & d.MaxY.ToString("0.######", CultureInfo.InvariantCulture) &
                       " motivo=" & d.DiscardReason)
        Next

        Dim ordered = kept.OrderBy(Function(c) c.Score).ThenByDescending(Function(c) c.MaxY - c.MinY).ToList()
        candidatesOut = ordered
        For Each c In ordered
            logger.Log("[DIM][VERT][CANDIDATE] " & c.Label &
                       " tipo=" & c.ObjType &
                       " prioridad=" & c.Score.ToString(CultureInfo.InvariantCulture) &
                       " orient=" & If(c.Orientation, "") &
                       " ymin=" & c.MinY.ToString("0.######", CultureInfo.InvariantCulture) &
                       " ymax=" & c.MaxY.ToString("0.######", CultureInfo.InvariantCulture))
        Next

        If ordered.Count = 0 Then
            logger.Log("[DIM][VERT][FILTER] sin candidatas funcionales tras el filtro.")
            Return False
        End If

        Dim bottom As VerticalFunctionalCandidate = Nothing
        Dim top As VerticalFunctionalCandidate = Nothing
        Dim detectedHeight As Double = Math.Abs(geom.EntityUnionMaxY - geom.EntityUnionMinY)
        If Not SelectValidEntityPairForVertical(dv, ordered, detectedHeight, logger, bottom, top) Then
            logger.Log("[DIM][VERT][FILTER] no se pudo seleccionar pareja válida para OVERALL_HEIGHT.")
            Return False
        End If

        bottomObj = ResolveCandidateObject(bottom, linesCol, arcsCol)
        topObj = ResolveCandidateObject(top, linesCol, arcsCol)
        If bottomObj Is Nothing OrElse topObj Is Nothing Then
            logger.Log("[DIM][VERT][FILTER] selección funcional sin referencia COM válida (bottom/top).")
            Return False
        End If

        bottomPickY = bottom.MinY
        topPickY = top.MaxY

        logger.Log("[DIM][VERT][SELECT] entidad inferior final=" & bottom.Label &
                   " tipo=" & bottom.ObjType &
                   " pick=(" & bottom.PickX.ToString("0.######", CultureInfo.InvariantCulture) & "," & bottomPickY.ToString("0.######", CultureInfo.InvariantCulture) & ")")
        logger.Log("[DIM][VERT][SELECT] entidad superior final=" & top.Label &
                   " tipo=" & top.ObjType &
                   " pick=(" & top.PickX.ToString("0.######", CultureInfo.InvariantCulture) & "," & topPickY.ToString("0.######", CultureInfo.InvariantCulture) & ")")
        logger.Log("[DIM][VERT][RESULT] altura funcional candidata=" &
                   (topPickY - bottomPickY).ToString("0.######", CultureInfo.InvariantCulture) &
                   " fuente=selección_priorizada")

        Return (bottomObj IsNot Nothing AndAlso topObj IsNot Nothing AndAlso topPickY > bottomPickY)
    End Function

    Private Shared Function TryInsertVerticalGeometryWithStrategies(
        dims As SolidEdgeFrameworkSupport.Dimensions,
        dv As DrawingView,
        frame As ViewPlacementFrame,
        log As DimensionLogger,
        baseLogger As Logger,
        candidates As List(Of VerticalFunctionalCandidate),
        preferredBottomObj As Object,
        preferredTopObj As Object,
        preferredBottomPickY As Double,
        preferredTopPickY As Double,
        expectedHeight As Double) As Boolean

        If dims Is Nothing OrElse dv Is Nothing OrElse frame Is Nothing Then Return False
        If candidates Is Nothing Then candidates = New List(Of VerticalFunctionalCandidate)()

        ' Estrategia 1: par seleccionado por filtro funcional.
        If TryInsertVerticalWithTryLog("1_funcional_inicial", dims, dv, frame, log, baseLogger, preferredBottomObj, preferredTopObj, preferredBottomPickY, preferredTopPickY, expectedHeight) Then
            Return True
        End If

        baseLogger?.Log("[DIM][VERT][TRY] fallback estrategia 2")
        ' Estrategia 2: forzar dos entidades distintas (inferior y superior) de mayor prioridad.
        Dim bottomCand = candidates.OrderBy(Function(c) c.Score).ThenBy(Function(c) c.MinY).FirstOrDefault()
        Dim topCand = candidates.OrderBy(Function(c) c.Score).ThenByDescending(Function(c) c.MaxY).
            FirstOrDefault(Function(c) Not (c.Kind = bottomCand?.Kind AndAlso c.Index1Based = bottomCand?.Index1Based))

        If bottomCand IsNot Nothing AndAlso topCand IsNot Nothing Then
            Dim o1 As Object = ResolveCandidateObjectFromView(dv, bottomCand)
            Dim o2 As Object = ResolveCandidateObjectFromView(dv, topCand)
            If TryInsertVerticalWithTryLog("2_distintas_prioridad", dims, dv, frame, log, baseLogger, o1, o2, bottomCand.MinY, topCand.MaxY, expectedHeight) Then
                Return True
            End If
        End If

        baseLogger?.Log("[DIM][VERT][TRY] fallback estrategia 3")
        ' Estrategia 3: mejor combinación compatible/distinta entre candidatas (top-N).
        Dim topN = candidates.OrderBy(Function(c) c.Score).ThenByDescending(Function(c) c.MaxY - c.MinY).Take(8).ToList()
        Dim bestA As VerticalFunctionalCandidate = Nothing
        Dim bestB As VerticalFunctionalCandidate = Nothing
        Dim bestSpan As Double = Double.MinValue
        For Each a In topN
            For Each b In topN
                If a Is Nothing OrElse b Is Nothing Then Continue For
                If a.Kind = b.Kind AndAlso a.Index1Based = b.Index1Based Then Continue For
                Dim span As Double = b.MaxY - a.MinY
                If span > bestSpan Then
                    bestSpan = span
                    bestA = a
                    bestB = b
                End If
            Next
        Next

        If bestA IsNot Nothing AndAlso bestB IsNot Nothing Then
            Dim o1 As Object = ResolveCandidateObjectFromView(dv, bestA)
            Dim o2 As Object = ResolveCandidateObjectFromView(dv, bestB)
            If TryInsertVerticalWithTryLog("3_mejor_combinacion", dims, dv, frame, log, baseLogger, o1, o2, bestA.MinY, bestB.MaxY, expectedHeight) Then
                Return True
            End If
        End If

        Return False
    End Function

    Private Shared Function TryInsertVerticalWithTryLog(
        strategyName As String,
        dims As SolidEdgeFrameworkSupport.Dimensions,
        dv As DrawingView,
        frame As ViewPlacementFrame,
        log As DimensionLogger,
        baseLogger As Logger,
        obj1 As Object,
        obj2 As Object,
        pickY1 As Double,
        pickY2 As Double,
        expectedHeight As Double) As Boolean

        Dim sameObj As Boolean = (obj1 IsNot Nothing AndAlso obj2 IsNot Nothing AndAlso Object.ReferenceEquals(obj1, obj2))
        baseLogger?.Log("[DIM][VERT][TRY] estrategia=" & strategyName)
        baseLogger?.Log("[DIM][VERT][TRY] Object1 tipo real=" & SafeTypeName(obj1))
        baseLogger?.Log("[DIM][VERT][TRY] Object2 tipo real=" & SafeTypeName(obj2))
        baseLogger?.Log("[DIM][VERT][TRY] sameObject=" & sameObj.ToString())
        baseLogger?.Log("[DIM][VERT][TRY] picks y1=" & pickY1.ToString("0.######", CultureInfo.InvariantCulture) &
                        " y2=" & pickY2.ToString("0.######", CultureInfo.InvariantCulture))

        Dim reason As String = ""
        Dim y1 As Double = pickY1
        Dim y2 As Double = pickY2
        If Not ValidatePairBeforeCom("OVERALL_HEIGHT", strategyName, obj1, obj2, 0, y1, 0, y2, expectedHeight, "vertical", baseLogger, reason) Then
            baseLogger?.Log("[DIM][PAIR][REJECT] motivo=" & reason)
            Return False
        End If
        baseLogger?.Log("[DIM][PAIR][ACCEPT] estrategia=" & strategyName)

        Return DimensionPlacementEngine.TryInsertVerticalBetweenObjects(dims, obj1, obj2, frame, pickY1, pickY2, log, dv, "OVERALL_HEIGHT")
    End Function

    Private Shared Function ResolveCandidateObjectFromView(dv As DrawingView, c As VerticalFunctionalCandidate) As Object
        If dv Is Nothing OrElse c Is Nothing Then Return Nothing
        Try
            Select Case c.Kind
                Case "line"
                    Dim lc As DVLines2d = dv.DVLines2d
                    If lc Is Nothing Then Return Nothing
                    Return CType(lc.Item(c.Index1Based), DVLine2d)
                Case "arc"
                    Dim ac As DVArcs2d = dv.DVArcs2d
                    If ac Is Nothing Then Return Nothing
                    Return CType(ac.Item(c.Index1Based), DVArc2d)
                Case Else
                    Return Nothing
            End Select
        Catch
            Return Nothing
        End Try
    End Function

    Private Shared Function IsDuplicateFeatureValue(existing As List(Of Double), value As Double) As Boolean
        If existing Is Nothing Then Return False
        For Each v In existing
            Dim tol As Double = Math.Max(0.0005R, Math.Abs(v) * 0.01R)
            If Math.Abs(v - value) <= tol Then Return True
        Next
        Return False
    End Function

    Private Shared Function ValidatePairBeforeCom(
        featureName As String,
        strategyName As String,
        obj1 As Object,
        obj2 As Object,
        x1 As Double,
        y1 As Double,
        x2 As Double,
        y2 As Double,
        expectedValue As Double,
        axis As String,
        logger As Logger,
        ByRef rejectReason As String) As Boolean

        rejectReason = ""
        Dim t1 As String = SafeTypeName(obj1)
        Dim t2 As String = SafeTypeName(obj2)
        Dim sameObj As Boolean = (obj1 IsNot Nothing AndAlso obj2 IsNot Nothing AndAlso Object.ReferenceEquals(obj1, obj2))
        Dim idA As String = BuildGeometryId(obj1)
        Dim idB As String = BuildGeometryId(obj2)
        Dim sameGeom As Boolean = IsSameGeometricEntity(obj1, obj2)
        Dim candidateDistance As Double = If(String.Equals(axis, "vertical", StringComparison.OrdinalIgnoreCase), Math.Abs(y2 - y1), Math.Abs(x2 - x1))

        logger?.Log("[DIM][PAIR][CHECK] feature=" & featureName & " estrategia=" & strategyName)
        logger?.Log("[DIM][PAIR][CHECK] obj1 tipo=" & t1)
        logger?.Log("[DIM][PAIR][CHECK] obj2 tipo=" & t2)
        logger?.Log("[DIM][PAIR][CHECK] idA=" & idA & " idB=" & idB)
        logger?.Log("[DIM][PAIR][CHECK] misma_entidad=" & sameObj.ToString())
        logger?.Log("[DIM][PAIR][CHECK] misma_geometria=" & sameGeom.ToString())
        logger?.Log("[DIM][PAIR][CHECK] distancia_candidata=" & candidateDistance.ToString("0.######", CultureInfo.InvariantCulture))

        If obj1 Is Nothing OrElse obj2 Is Nothing Then
            rejectReason = "objeto COM nulo"
            logger?.Log("[DIM][PAIR][CHECK] validación previa=REJECT motivo=" & rejectReason)
            Return False
        End If
        If sameObj Then
            rejectReason = "misma entidad COM para ambos extremos"
            logger?.Log("[DIM][PAIR][CHECK] validación previa=REJECT motivo=" & rejectReason)
            Return False
        End If
        If sameGeom Then
            rejectReason = "misma_geometria"
            logger?.Log("[DIM][PAIR][CHECK] validación previa=REJECT motivo=" & rejectReason)
            Return False
        End If
        If candidateDistance <= 0.0005R Then
            rejectReason = "distancia degenerada o micro"
            logger?.Log("[DIM][PAIR][CHECK] validación previa=REJECT motivo=" & rejectReason)
            Return False
        End If

        ' Reglas por feature.
        If String.Equals(featureName, "THICKNESS", StringComparison.OrdinalIgnoreCase) Then
            If Not (TypeOf obj1 Is DVLine2d AndAlso TypeOf obj2 Is DVLine2d) Then
                rejectReason = "THICKNESS requiere dos DVLine2d"
                logger?.Log("[DIM][PAIR][CHECK] validación previa=REJECT motivo=" & rejectReason)
                Return False
            End If
            Dim tolT As Double = Math.Max(0.0005R, Math.Abs(expectedValue) * 0.25R)
            If expectedValue > 0 AndAlso Math.Abs(candidateDistance - expectedValue) > tolT Then
                rejectReason = "distancia no encaja con THICKNESS esperado"
                logger?.Log("[DIM][PAIR][CHECK] validación previa=REJECT motivo=" & rejectReason)
                Return False
            End If
        ElseIf String.Equals(featureName, "OVERALL_HEIGHT", StringComparison.OrdinalIgnoreCase) Then
            Dim allowed As Boolean = (TypeOf obj1 Is DVLine2d OrElse TypeOf obj1 Is DVArc2d) AndAlso (TypeOf obj2 Is DVLine2d OrElse TypeOf obj2 Is DVArc2d)
            If Not allowed Then
                rejectReason = "tipos no compatibles para OVERALL_HEIGHT"
                logger?.Log("[DIM][PAIR][CHECK] validación previa=REJECT motivo=" & rejectReason)
                Return False
            End If
            If strategyName = "1_funcional_inicial" AndAlso (Not TypeOf obj1 Is DVLine2d OrElse Not TypeOf obj2 Is DVLine2d) Then
                rejectReason = "estrategia 1 exige dos líneas estructurales"
                logger?.Log("[DIM][PAIR][CHECK] validación previa=REJECT motivo=" & rejectReason)
                Return False
            End If
            Dim tolH As Double = Math.Max(0.001R, Math.Abs(expectedValue) * 0.35R)
            If expectedValue > 0 AndAlso Math.Abs(candidateDistance - expectedValue) > tolH Then
                rejectReason = "distancia candidata se desvía de OVERALL_HEIGHT"
                logger?.Log("[DIM][PAIR][CHECK] validación previa=REJECT motivo=" & rejectReason)
                Return False
            End If
        End If

        logger?.Log("[DIM][PAIR][CHECK] validación previa=OK")
        Return True
    End Function

    Private Shared Function SelectValidEntityPairForVertical(
        dv As DrawingView,
        candidates As List(Of VerticalFunctionalCandidate),
        expectedHeight As Double,
        logger As Logger,
        ByRef outBottom As VerticalFunctionalCandidate,
        ByRef outTop As VerticalFunctionalCandidate) As Boolean

        outBottom = Nothing : outTop = Nothing
        If dv Is Nothing OrElse candidates Is Nothing OrElse candidates.Count < 2 Then Return False

        Dim globalMinY As Double = candidates.Min(Function(c) c.MinY)
        Dim globalMaxY As Double = candidates.Max(Function(c) c.MaxY)
        Dim globalMinX As Double = candidates.Min(Function(c) c.MinX)
        Dim globalMaxX As Double = candidates.Max(Function(c) c.MaxX)
        Dim tolYExtreme As Double = Math.Max(0.0015R, Math.Abs(expectedHeight) * 0.2R)
        Dim tolXEdge As Double = Math.Max(0.0015R, (globalMaxX - globalMinX) * 0.05R)

        Dim bestA As VerticalFunctionalCandidate = Nothing
        Dim bestB As VerticalFunctionalCandidate = Nothing
        Dim bestErr As Double = Double.MaxValue

        ' Prioridad 1: dos entidades estructurales que representen minY/maxY funcional (mezcla permitida).
        Dim structural = candidates.
            Where(Function(c) c IsNot Nothing AndAlso c.Score <= 2 AndAlso (c.Kind = "line" OrElse c.Kind = "arc")).
            ToList()
        Dim nearBottomS = structural.Where(Function(c) Math.Abs(c.MinY - globalMinY) <= tolYExtreme).ToList()
        Dim nearTopS = structural.Where(Function(c) Math.Abs(c.MaxY - globalMaxY) <= tolYExtreme).ToList()
        If TryPickBestPairCrossSets(dv, nearBottomS, nearTopS, expectedHeight, globalMinY, globalMaxY, globalMinX, globalMaxX, tolYExtreme, tolXEdge, logger, bestA, bestB, bestErr, "P1_estructural_funcional") Then
            outBottom = bestA : outTop = bestB : Return True
        End If

        ' Prioridad 2: línea principal + arco funcional (si aporta realmente el extremo).
        Dim linesMain = candidates.Where(Function(c) c IsNot Nothing AndAlso c.Kind = "line" AndAlso c.Score <= 2).ToList()
        Dim arcsMain = candidates.Where(Function(c) c IsNot Nothing AndAlso c.Kind = "arc").ToList()
        If TryPickBestPairCrossSets(dv, linesMain, arcsMain, expectedHeight, globalMinY, globalMaxY, globalMinX, globalMaxX, tolYExtreme, tolXEdge, logger, bestA, bestB, bestErr, "P2_linea_arco_funcional") _
            OrElse TryPickBestPairCrossSets(dv, arcsMain, linesMain, expectedHeight, globalMinY, globalMaxY, globalMinX, globalMaxX, tolYExtreme, tolXEdge, logger, bestA, bestB, bestErr, "P2_arco_linea_funcional") Then
            outBottom = bestA : outTop = bestB : Return True
        End If

        ' Prioridad 3: fallback actual, incluyendo verticales.
        Dim verticalLines = candidates.
            Where(Function(c) c IsNot Nothing AndAlso c.Kind = "line" AndAlso
                               String.Equals(c.Orientation, "vertical", StringComparison.OrdinalIgnoreCase)).
            ToList()
        If TryPickBestPairForHeight(dv, verticalLines, expectedHeight, globalMinY, globalMaxY, globalMinX, globalMaxX, tolYExtreme, tolXEdge, logger, bestA, bestB, bestErr, "P3_verticales_respaldo") Then
            outBottom = bestA : outTop = bestB : Return True
        End If

        Return False
    End Function

    Private Shared Function TryPickBestPairForHeight(
        dv As DrawingView,
        pool As List(Of VerticalFunctionalCandidate),
        expectedHeight As Double,
        globalMinY As Double,
        globalMaxY As Double,
        globalMinX As Double,
        globalMaxX As Double,
        tolYExtreme As Double,
        tolXEdge As Double,
        logger As Logger,
        ByRef outBottom As VerticalFunctionalCandidate,
        ByRef outTop As VerticalFunctionalCandidate,
        ByRef outErr As Double,
        sourceTag As String) As Boolean

        If pool Is Nothing OrElse pool.Count < 2 Then Return False
        Dim bestA As VerticalFunctionalCandidate = Nothing
        Dim bestB As VerticalFunctionalCandidate = Nothing
        Dim bestErr As Double = Double.MaxValue

        For i As Integer = 0 To pool.Count - 2
            For j As Integer = i + 1 To pool.Count - 1
                Dim a = pool(i)
                Dim b = pool(j)
                Dim bottom = If(a.MinY <= b.MinY, a, b)
                Dim top = If(a.MaxY >= b.MaxY, a, b)
                If Not EvaluateHeightPair(dv, bottom, top, expectedHeight, globalMinY, globalMaxY, globalMinX, globalMaxX, tolYExtreme, tolXEdge, logger, sourceTag) Then Continue For
                Dim h As Double = Math.Abs(top.MaxY - bottom.MinY)
                Dim err As Double = Math.Abs(h - expectedHeight)
                If err < bestErr Then
                    bestErr = err
                    bestA = bottom
                    bestB = top
                End If
            Next
        Next

        If bestA IsNot Nothing AndAlso bestB IsNot Nothing Then
            outBottom = bestA
            outTop = bestB
            outErr = bestErr
            logger?.Log("[DIM][VERT][SELECT] fuente=" & sourceTag &
                        " bottom=" & bestA.Label & " top=" & bestB.Label &
                        " error_altura=" & bestErr.ToString("0.######", CultureInfo.InvariantCulture))
            Return True
        End If
        Return False
    End Function

    Private Shared Function TryPickBestPairCrossSets(
        dv As DrawingView,
        bottomSet As List(Of VerticalFunctionalCandidate),
        topSet As List(Of VerticalFunctionalCandidate),
        expectedHeight As Double,
        globalMinY As Double,
        globalMaxY As Double,
        globalMinX As Double,
        globalMaxX As Double,
        tolYExtreme As Double,
        tolXEdge As Double,
        logger As Logger,
        ByRef outBottom As VerticalFunctionalCandidate,
        ByRef outTop As VerticalFunctionalCandidate,
        ByRef outErr As Double,
        sourceTag As String) As Boolean

        If bottomSet Is Nothing OrElse topSet Is Nothing OrElse bottomSet.Count = 0 OrElse topSet.Count = 0 Then Return False
        Dim bestA As VerticalFunctionalCandidate = Nothing
        Dim bestB As VerticalFunctionalCandidate = Nothing
        Dim bestErr As Double = Double.MaxValue

        For Each btm In bottomSet
            For Each tp In topSet
                If btm Is Nothing OrElse tp Is Nothing Then Continue For
                If btm.Kind = tp.Kind AndAlso btm.Index1Based = tp.Index1Based Then
                    logger?.Log("[DIM][PAIR][REJECT] motivo=misma_entidad_indice")
                    Continue For
                End If
                If Not EvaluateHeightPair(dv, btm, tp, expectedHeight, globalMinY, globalMaxY, globalMinX, globalMaxX, tolYExtreme, tolXEdge, logger, sourceTag) Then Continue For
                Dim h As Double = Math.Abs(tp.MaxY - btm.MinY)
                Dim err As Double = Math.Abs(h - expectedHeight)
                If err < bestErr Then
                    bestErr = err
                    bestA = btm
                    bestB = tp
                End If
            Next
        Next

        If bestA IsNot Nothing AndAlso bestB IsNot Nothing Then
            outBottom = bestA
            outTop = bestB
            outErr = bestErr
            logger?.Log("[DIM][VERT][SELECT] fuente=" & sourceTag &
                        " bottom=" & bestA.Label & " top=" & bestB.Label &
                        " error_altura=" & bestErr.ToString("0.######", CultureInfo.InvariantCulture))
            Return True
        End If
        Return False
    End Function

    Private Shared Function EvaluateHeightPair(
        dv As DrawingView,
        bottom As VerticalFunctionalCandidate,
        top As VerticalFunctionalCandidate,
        expectedHeight As Double,
        globalMinY As Double,
        globalMaxY As Double,
        globalMinX As Double,
        globalMaxX As Double,
        tolYExtreme As Double,
        tolXEdge As Double,
        logger As Logger,
        sourceTag As String) As Boolean

        If bottom Is Nothing OrElse top Is Nothing Then Return False
        logger?.Log("[DIM][HEIGHT][CHECK] candidato A=" & bottom.Label & " (" & bottom.ObjType & ")")
        logger?.Log("[DIM][HEIGHT][CHECK] candidato B=" & top.Label & " (" & top.ObjType & ")")
        logger?.Log("[DIM][HEIGHT][CHECK] A=" & bottom.Label)
        logger?.Log("[DIM][HEIGHT][CHECK] B=" & top.Label)
        logger?.Log("[DIM][HEIGHT][CHECK] orientaciones=" & If(bottom.Orientation, "?") & "/" & If(top.Orientation, "?"))
        Dim objB = ResolveCandidateObjectFromView(dv, bottom)
        Dim objT = ResolveCandidateObjectFromView(dv, top)
        If objB Is Nothing OrElse objT Is Nothing Then
            logger?.Log("[DIM][HEIGHT][CHECK] razón de rechazo=objeto_no_resuelto fuente=" & sourceTag)
            Return False
        End If

        logger?.Log("[DIM][PAIR][CHECK] idA=" & BuildGeometryId(objB) & " idB=" & BuildGeometryId(objT))
        Dim sameGeom As Boolean = IsSameGeometricEntity(objB, objT)
        logger?.Log("[DIM][PAIR][CHECK] misma_geometria=" & sameGeom.ToString())
        If sameGeom Then
            logger?.Log("[DIM][HEIGHT][CHECK] razón de rechazo=misma_geometria fuente=" & sourceTag)
            Return False
        End If

        If top.MaxY <= bottom.MinY Then
            logger?.Log("[DIM][HEIGHT][CHECK] razón de rechazo=orden_y_no_valido fuente=" & sourceTag)
            Return False
        End If

        If Not IsValidOverallHeightPair(bottom, top, globalMinY, globalMaxY, globalMinX, globalMaxX, tolYExtreme, tolXEdge, logger) Then
            logger?.Log("[DIM][HEIGHT][REJECT] motivo=pareja_bbox_no_funcional")
            Return False
        End If

        Dim touchesBottomExtreme As Boolean = Math.Abs(bottom.MinY - globalMinY) <= tolYExtreme
        Dim touchesTopExtreme As Boolean = Math.Abs(top.MaxY - globalMaxY) <= tolYExtreme
        If Not (touchesBottomExtreme AndAlso touchesTopExtreme) Then
            logger?.Log("[DIM][HEIGHT][CHECK] razón de rechazo=no_representa_extremos_reales fuente=" & sourceTag)
            Return False
        End If

        Dim h As Double = Math.Abs(top.MaxY - bottom.MinY)
        Dim tolH As Double = Math.Max(0.001R, Math.Abs(expectedHeight) * 0.35R)
        If expectedHeight > 0 AndAlso Math.Abs(h - expectedHeight) > tolH Then
            logger?.Log("[DIM][HEIGHT][CHECK] razón de rechazo=distancia_fuera_tolerancia fuente=" & sourceTag)
            Return False
        End If

        logger?.Log("[DIM][HEIGHT][ACCEPT] pareja final seleccionada fuente=" & sourceTag &
                    " A=" & bottom.Label & " B=" & top.Label &
                    " h=" & h.ToString("0.######", CultureInfo.InvariantCulture))
        logger?.Log("[DIM][HEIGHT][ACCEPT] pareja funcional válida")
        Return True
    End Function

    Private Shared Function IsValidOverallHeightPair(
        a As VerticalFunctionalCandidate,
        b As VerticalFunctionalCandidate,
        globalMinY As Double,
        globalMaxY As Double,
        globalMinX As Double,
        globalMaxX As Double,
        tolYExtreme As Double,
        tolXEdge As Double,
        logger As Logger) As Boolean

        If a Is Nothing OrElse b Is Nothing Then Return False

        Dim aIsHorizontalLower As Boolean =
            String.Equals(a.Orientation, "horizontal", StringComparison.OrdinalIgnoreCase) AndAlso
            Math.Abs(a.MinY - globalMinY) <= tolYExtreme
        Dim bIsHorizontalLower As Boolean =
            String.Equals(b.Orientation, "horizontal", StringComparison.OrdinalIgnoreCase) AndAlso
            Math.Abs(b.MinY - globalMinY) <= tolYExtreme

        Dim aMidX As Double = (a.MinX + a.MaxX) / 2.0R
        Dim bMidX As Double = (b.MinX + b.MaxX) / 2.0R
        Dim aIsVerticalLateral As Boolean =
            String.Equals(a.Orientation, "vertical", StringComparison.OrdinalIgnoreCase) AndAlso
            (Math.Abs(aMidX - globalMinX) <= tolXEdge OrElse Math.Abs(aMidX - globalMaxX) <= tolXEdge)
        Dim bIsVerticalLateral As Boolean =
            String.Equals(b.Orientation, "vertical", StringComparison.OrdinalIgnoreCase) AndAlso
            (Math.Abs(bMidX - globalMinX) <= tolXEdge OrElse Math.Abs(bMidX - globalMaxX) <= tolXEdge)

        logger?.Log("[DIM][HEIGHT][CHECK] A_es_horizontal_inferior=" & aIsHorizontalLower.ToString())
        logger?.Log("[DIM][HEIGHT][CHECK] B_es_vertical_lateral=" & bIsVerticalLateral.ToString())
        logger?.Log("[DIM][HEIGHT][CHECK] B_es_horizontal_inferior=" & bIsHorizontalLower.ToString())
        logger?.Log("[DIM][HEIGHT][CHECK] A_es_vertical_lateral=" & aIsVerticalLateral.ToString())

        If (aIsHorizontalLower AndAlso bIsVerticalLateral) OrElse (bIsHorizontalLower AndAlso aIsVerticalLateral) Then
            Return False
        End If

        Return True
    End Function

    Private Shared Function BuildGeometryId(obj As Object) As String
        If obj Is Nothing Then Return "(null)"
        Try
            If TypeOf obj Is DVLine2d Then
                Dim ln As DVLine2d = CType(obj, DVLine2d)
                Dim x1 As Double = 0, y1 As Double = 0, x2 As Double = 0, y2 As Double = 0
                ln.GetStartPoint(x1, y1)
                ln.GetEndPoint(x2, y2)
                Return "LINE:" & NormalizeLineId(x1, y1, x2, y2)
            End If
        Catch
        End Try
        Return SafeTypeName(obj)
    End Function

    Private Shared Function IsSameGeometricEntity(obj1 As Object, obj2 As Object) As Boolean
        If obj1 Is Nothing OrElse obj2 Is Nothing Then Return False
        If Object.ReferenceEquals(obj1, obj2) Then Return True

        ' Regla fuerte para líneas.
        If TypeOf obj1 Is DVLine2d AndAlso TypeOf obj2 Is DVLine2d Then
            Try
                Dim a As DVLine2d = CType(obj1, DVLine2d)
                Dim b As DVLine2d = CType(obj2, DVLine2d)
                Dim a1x As Double = 0, a1y As Double = 0, a2x As Double = 0, a2y As Double = 0
                Dim b1x As Double = 0, b1y As Double = 0, b2x As Double = 0, b2y As Double = 0
                a.GetStartPoint(a1x, a1y) : a.GetEndPoint(a2x, a2y)
                b.GetStartPoint(b1x, b1y) : b.GetEndPoint(b2x, b2y)

                Dim tol As Double = 0.000001R
                Dim sameDir As Boolean =
                    Math.Abs(a1x - b1x) <= tol AndAlso Math.Abs(a1y - b1y) <= tol AndAlso
                    Math.Abs(a2x - b2x) <= tol AndAlso Math.Abs(a2y - b2y) <= tol
                Dim oppDir As Boolean =
                    Math.Abs(a1x - b2x) <= tol AndAlso Math.Abs(a1y - b2y) <= tol AndAlso
                    Math.Abs(a2x - b1x) <= tol AndAlso Math.Abs(a2y - b1y) <= tol
                If sameDir OrElse oppDir Then
                    Dim la As Double = Math.Sqrt((a2x - a1x) * (a2x - a1x) + (a2y - a1y) * (a2y - a1y))
                    Dim lb As Double = Math.Sqrt((b2x - b1x) * (b2x - b1x) + (b2y - b1y) * (b2y - b1y))
                    If Math.Abs(la - lb) <= tol Then
                        Return True
                    End If
                End If
            Catch
            End Try
        End If

        Return String.Equals(BuildGeometryId(obj1), BuildGeometryId(obj2), StringComparison.Ordinal)
    End Function

    Private Shared Function NormalizeLineId(x1 As Double, y1 As Double, x2 As Double, y2 As Double) As String
        ' id independiente del sentido de la línea
        Dim a As String = Round6(x1) & "," & Round6(y1)
        Dim b As String = Round6(x2) & "," & Round6(y2)
        If String.CompareOrdinal(a, b) <= 0 Then
            Return a & "->" & b
        End If
        Return b & "->" & a
    End Function

    Private Shared Function Round6(v As Double) As String
        Return Math.Round(v, 6, MidpointRounding.AwayFromZero).ToString("0.######", CultureInfo.InvariantCulture)
    End Function

    Private Shared Function ExtractSignificantPoints(dv As DrawingView, owner As Object, ownerTag As String, logger As Logger) As List(Of SignificantPoint)
        Dim outList As New List(Of SignificantPoint)()
        If dv Is Nothing OrElse owner Is Nothing Then Return outList

        Dim ownerType As String = SafeTypeName(owner)
        logger?.Log("[DIM][SIGPOINT] entidad=" & ownerTag & " tipo=" & ownerType)

        If TypeOf owner Is DVLine2d Then
            Dim ln As DVLine2d = CType(owner, DVLine2d)
            Dim x1 As Double = 0, y1 As Double = 0, x2 As Double = 0, y2 As Double = 0
            Try
                ln.GetStartPoint(x1, y1)
                ln.GetEndPoint(x2, y2)
                AddSigPoint(outList, SignificantPointKind.StartPoint, x1, y1, owner, ownerType, "line_start", SignificantPointCoordSystem.View, logger)
                AddSigPoint(outList, SignificantPointKind.EndPoint, x2, y2, owner, ownerType, "line_end", SignificantPointCoordSystem.View, logger)
                AddSigPoint(outList, SignificantPointKind.MidPoint, (x1 + x2) / 2.0R, (y1 + y2) / 2.0R, owner, ownerType, "line_mid", SignificantPointCoordSystem.View, logger)
                AddSigPoint(outList, SignificantPointKind.TangencyCandidate, x1, y1, owner, ownerType, "line_tan_start", SignificantPointCoordSystem.View, logger)
                AddSigPoint(outList, SignificantPointKind.TangencyCandidate, x2, y2, owner, ownerType, "line_tan_end", SignificantPointCoordSystem.View, logger)
            Catch ex As Exception
                logger?.Log("[DIM][SIGPOINT][WARN] DVLine2d sin Start/End: " & ex.Message)
            End Try
            Return outList
        End If

        If TypeOf owner Is DVArc2d Then
            Dim a As DVArc2d = CType(owner, DVArc2d)
            Dim cx As Double = 0, cy As Double = 0, r As Double = 0
            Dim sAng As Double = 0, sw As Double = 0
            Dim hasCenter As Boolean = TryGetArcCenterRadiusLocal(a, cx, cy, r)
            Dim hasAngles As Boolean = TryGetArcAnglesLocal(a, sAng, sw)
            Dim x1 As Double = 0, y1 As Double = 0, x2 As Double = 0, y2 As Double = 0
            Dim hasSE As Boolean = False
            Try
                a.GetStartPoint(x1, y1)
                a.GetEndPoint(x2, y2)
                hasSE = True
            Catch ex As Exception
                logger?.Log("[DIM][SIGPOINT][WARN] DVArc2d sin Start/End: " & ex.Message)
            End Try
            If hasCenter Then
                AddSigPoint(outList, SignificantPointKind.CenterPoint, cx, cy, owner, ownerType, "arc_center", SignificantPointCoordSystem.View, logger)
            Else
                logger?.Log("[DIM][SIGPOINT][WARN] DVArc2d sin centro/radio legibles.")
            End If
            If hasSE Then
                AddSigPoint(outList, SignificantPointKind.StartPoint, x1, y1, owner, ownerType, "arc_start", SignificantPointCoordSystem.View, logger)
                AddSigPoint(outList, SignificantPointKind.EndPoint, x2, y2, owner, ownerType, "arc_end", SignificantPointCoordSystem.View, logger)
                AddSigPoint(outList, SignificantPointKind.TangencyCandidate, x1, y1, owner, ownerType, "arc_tan_start", SignificantPointCoordSystem.View, logger)
                AddSigPoint(outList, SignificantPointKind.TangencyCandidate, x2, y2, owner, ownerType, "arc_tan_end", SignificantPointCoordSystem.View, logger)
            End If

            If hasCenter AndAlso hasAngles AndAlso r > 0 Then
                AddArcQuadrantIfVisible(outList, owner, ownerType, cx, cy, r, sAng, sw, 0.0R, SignificantPointKind.QuadrantRight, "arc_q_right", logger)
                AddArcQuadrantIfVisible(outList, owner, ownerType, cx, cy, r, sAng, sw, Math.PI / 2.0R, SignificantPointKind.QuadrantTop, "arc_q_top", logger)
                AddArcQuadrantIfVisible(outList, owner, ownerType, cx, cy, r, sAng, sw, Math.PI, SignificantPointKind.QuadrantLeft, "arc_q_left", logger)
                AddArcQuadrantIfVisible(outList, owner, ownerType, cx, cy, r, sAng, sw, 3.0R * Math.PI / 2.0R, SignificantPointKind.QuadrantBottom, "arc_q_bottom", logger)
            End If
            Return outList
        End If

        If TypeOf owner Is DVCircle2d Then
            Dim c As DVCircle2d = CType(owner, DVCircle2d)
            Dim cx As Double = 0, cy As Double = 0, r As Double = 0
            If TryGetCircleCenterRadiusLocal(c, cx, cy, r) AndAlso r > 0 Then
                AddSigPoint(outList, SignificantPointKind.CenterPoint, cx, cy, owner, ownerType, "circle_center", SignificantPointCoordSystem.View, logger)
                AddSigPoint(outList, SignificantPointKind.QuadrantTop, cx, cy + r, owner, ownerType, "circle_q_top", SignificantPointCoordSystem.View, logger)
                AddSigPoint(outList, SignificantPointKind.QuadrantBottom, cx, cy - r, owner, ownerType, "circle_q_bottom", SignificantPointCoordSystem.View, logger)
                AddSigPoint(outList, SignificantPointKind.QuadrantLeft, cx - r, cy, owner, ownerType, "circle_q_left", SignificantPointCoordSystem.View, logger)
                AddSigPoint(outList, SignificantPointKind.QuadrantRight, cx + r, cy, owner, ownerType, "circle_q_right", SignificantPointCoordSystem.View, logger)
            Else
                logger?.Log("[DIM][SIGPOINT][WARN] DVCircle2d sin centro/radio legibles.")
            End If
        End If

        Return outList
    End Function

    Private Shared Sub AddSigPoint(
        list As List(Of SignificantPoint),
        kind As SignificantPointKind,
        x As Double,
        y As Double,
        owner As Object,
        ownerType As String,
        tag As String,
        coordSystem As SignificantPointCoordSystem,
        logger As Logger)

        list.Add(New SignificantPoint With {
            .Kind = kind, .X = x, .Y = y, .Owner = owner, .OwnerType = ownerType, .Tag = tag, .CoordSystem = coordSystem
        })
    End Sub

    Private Shared Function TryGetArcCenterRadiusLocal(a As DVArc2d, ByRef cx As Double, ByRef cy As Double, ByRef rad As Double) As Boolean
        cx = 0 : cy = 0 : rad = 0
        Try
            Dim obj As Object = a
            obj.GetCenterPoint(cx, cy)
            Try
                rad = CDbl(obj.Radius)
            Catch
                Try
                    rad = CDbl(obj.Diameter) / 2.0R
                Catch
                End Try
            End Try
            Return rad > 0
        Catch
            Return False
        End Try
    End Function

    Private Shared Function TryGetCircleCenterRadiusLocal(c As DVCircle2d, ByRef cx As Double, ByRef cy As Double, ByRef rad As Double) As Boolean
        cx = 0 : cy = 0 : rad = 0
        Try
            Dim obj As Object = c
            obj.GetCenterPoint(cx, cy)
            Try
                rad = CDbl(obj.Radius)
            Catch
                rad = CDbl(obj.Diameter) / 2.0R
            End Try
            Return rad > 0
        Catch
            Return False
        End Try
    End Function

    Private Shared Function TryGetArcAnglesLocal(a As DVArc2d, ByRef startRad As Double, ByRef sweepRad As Double) As Boolean
        startRad = 0 : sweepRad = 0
        Try
            Dim obj As Object = a
            Try
                startRad = CDbl(obj.StartAngle)
                sweepRad = CDbl(obj.SweepAngle)
                If Math.Abs(sweepRad) > Math.PI * 4.0R Then
                    startRad *= Math.PI / 180.0R
                    sweepRad *= Math.PI / 180.0R
                End If
                Return True
            Catch
            End Try

            Dim cx As Double = 0, cy As Double = 0, r As Double = 0
            If Not TryGetArcCenterRadiusLocal(a, cx, cy, r) Then Return False
            Dim sx As Double = 0, sy As Double = 0, ex As Double = 0, ey As Double = 0
            a.GetStartPoint(sx, sy)
            a.GetEndPoint(ex, ey)
            startRad = Math.Atan2(sy - cy, sx - cx)
            Dim endRad As Double = Math.Atan2(ey - cy, ex - cx)
            sweepRad = endRad - startRad
            Return True
        Catch
            Return False
        End Try
    End Function

    Private Shared Sub AddArcQuadrantIfVisible(list As List(Of SignificantPoint), owner As Object, ownerType As String, cx As Double, cy As Double, r As Double, startRad As Double, sweepRad As Double, qAng As Double, kind As SignificantPointKind, tag As String, logger As Logger)
        If IsAngleInsideArcSweep(qAng, startRad, sweepRad) Then
            AddSigPoint(list, kind, cx + r * Math.Cos(qAng), cy + r * Math.Sin(qAng), owner, ownerType, tag, SignificantPointCoordSystem.View, logger)
        End If
    End Sub

    Private Shared Function IsAngleInsideArcSweep(testAng As Double, startRad As Double, sweepRad As Double) As Boolean
        Dim twoPi As Double = 2.0R * Math.PI
        Dim s As Double = NormalizeAngle(startRad)
        Dim t As Double = NormalizeAngle(testAng)
        Dim e As Double = NormalizeAngle(startRad + sweepRad)
        Dim w As Double = sweepRad
        If Math.Abs(w) >= twoPi - 0.001R Then Return True
        If w >= 0 Then
            If e >= s Then
                Return t >= s AndAlso t <= e
            End If
            Return t >= s OrElse t <= e
        End If
        If e <= s Then
            Return t <= s AndAlso t >= e
        End If
        Return t <= s OrElse t >= e
    End Function

    Private Shared Function NormalizeAngle(a As Double) As Double
        Dim twoPi As Double = 2.0R * Math.PI
        Dim n As Double = a Mod twoPi
        If n < 0 Then n += twoPi
        Return n
    End Function

    Private Shared Function ConvertSignificantPointToSheet(dv As DrawingView, p As SignificantPoint, logger As Logger) As SignificantPoint
        If p Is Nothing Then Return Nothing
        If dv Is Nothing Then Return p

        Dim outP As New SignificantPoint With {
            .Kind = p.Kind,
            .X = p.X,
            .Y = p.Y,
            .Owner = p.Owner,
            .OwnerType = p.OwnerType,
            .Tag = p.Tag,
            .CoordSystem = p.CoordSystem
        }

        If p.CoordSystem = SignificantPointCoordSystem.Sheet Then
            Return outP
        End If

        Try
            Dim sx As Double = 0, sy As Double = 0
            dv.ViewToSheet(p.X, p.Y, sx, sy)
            outP.X = sx
            outP.Y = sy
            outP.CoordSystem = SignificantPointCoordSystem.Sheet
        Catch ex As Exception
            logger?.Log("[DIM][COORD][WARN] fallo al convertir a hoja: " & ex.Message)
        End Try
        Return outP
    End Function

    Private Shared Function SelectPointForAxisExtreme(
        dv As DrawingView,
        entityObj As Object,
        entityTag As String,
        axis As String,
        pickMin As Boolean,
        logger As Logger,
        Optional box As Nullable(Of ViewSheetBoundingBox) = Nothing) As SignificantPoint

        Dim pts = ExtractSignificantPoints(dv, entityObj, entityTag, logger)
        If pts Is Nothing OrElse pts.Count = 0 Then Return Nothing
        Dim best As SignificantPoint = Nothing
        For Each p In pts
            Dim pSheet As SignificantPoint = ConvertSignificantPointToSheet(dv, p, logger)
            If pSheet Is Nothing Then Continue For
            If best Is Nothing Then
                best = pSheet
            ElseIf String.Equals(axis, "x", StringComparison.OrdinalIgnoreCase) Then
                If pickMin AndAlso pSheet.X < best.X Then best = pSheet
                If (Not pickMin) AndAlso pSheet.X > best.X Then best = pSheet
            Else
                If pickMin AndAlso pSheet.Y < best.Y Then best = pSheet
                If (Not pickMin) AndAlso pSheet.Y > best.Y Then best = pSheet
            End If
        Next
        If best IsNot Nothing Then
            logger?.Log("[DIM][SIGSELECT] entidad=" & entityTag &
                        " punto elegido=" & best.Kind.ToString() &
                        " (" & Round6(best.X) & "," & Round6(best.Y) & ")")
            If box.HasValue Then
                Dim b = box.Value
                Dim tol As Double = Math.Max(0.002R, Math.Max(b.Width, b.Height) * 0.1R)
                If best.X < b.MinX - tol OrElse best.X > b.MaxX + tol OrElse best.Y < b.MinY - tol OrElse best.Y > b.MaxY + tol Then
                    logger?.Log("[DIM][COORD][WARN] punto convertido fuera del bbox esperado")
                End If
            End If
        End If
        Return best
    End Function

    Private NotInheritable Class GeometricFeatureAnalyzer
        Private Sub New()
        End Sub

        Public Shared Function TryDetectThicknessFeature(
            dv As DrawingView,
            geom As ViewGeometrySummary,
            box As ViewSheetBoundingBox,
            logger As Logger,
            ByRef outFeature As FeatureDimensionCandidate) As Boolean

            outFeature = Nothing
            If dv Is Nothing OrElse geom Is Nothing Then Return False

            Dim lines = If(geom.Lines, New List(Of LineGeometryInfo)())
            If lines.Count < 2 Then Return False

            Dim minStructLen As Double = Math.Max(0.008R, Math.Max(box.Width, box.Height) * 0.12R)
            Dim maxThickness As Double = Math.Max(0.03R, Math.Min(box.Width, box.Height) * 0.35R)

            Dim structuralH = lines.Where(Function(l) String.Equals(l.Orientation, "horizontal", StringComparison.OrdinalIgnoreCase) AndAlso l.Length >= minStructLen).ToList()
            Dim structuralV = lines.Where(Function(l) String.Equals(l.Orientation, "vertical", StringComparison.OrdinalIgnoreCase) AndAlso l.Length >= minStructLen).ToList()

            logger?.Log("[DIM][CLASSIFY] líneas estructurales horizontales=" & structuralH.Count.ToString(CultureInfo.InvariantCulture) &
                        " verticales=" & structuralV.Count.ToString(CultureInfo.InvariantCulture))

            Dim minOverlapY As Double = Math.Max(0.006R, box.Height * 0.038R)
            Dim minOverlapX As Double = Math.Max(0.006R, box.Width * 0.038R)

            Dim best As FeatureDimensionCandidate = Nothing
            Dim bestScore As Double = Double.MinValue

            ' Espesor por dos líneas verticales paralelas cercanas (cota horizontal).
            For i As Integer = 0 To structuralV.Count - 2
                For j As Integer = i + 1 To structuralV.Count - 1
                    Dim a = structuralV(i)
                    Dim b = structuralV(j)
                    Dim t As Double = Math.Abs(a.MidX - b.MidX)
                    If t < 0.0005R OrElse t > maxThickness Then Continue For

                    Dim overlapY As Double = Math.Min(a.BboxMaxY, b.BboxMaxY) - Math.Max(a.BboxMinY, b.BboxMinY)
                    If overlapY <= 0 Then Continue For
                    If overlapY < minOverlapY Then Continue For

                    Dim score As Double = overlapY / (1.0R + t)
                    If score > bestScore Then
                        logger?.Log("[DIM][THICKNESS][CANDIDATE] par_V DVLine2d#" & a.Index.ToString(CultureInfo.InvariantCulture) & "/#" & b.Index.ToString(CultureInfo.InvariantCulture) &
                                    " sepX=" & t.ToString("0.######", CultureInfo.InvariantCulture) & " overlapY=" & overlapY.ToString("0.######", CultureInfo.InvariantCulture) &
                                    " score=" & score.ToString("0.######", CultureInfo.InvariantCulture))
                        bestScore = score
                        best = BuildThicknessCandidate(dv, a.Index, b.Index, "horizontal", t)
                    End If
                Next
            Next

            ' Espesor por dos líneas horizontales paralelas cercanas (cota vertical), si no hay mejor vertical-pair.
            For i As Integer = 0 To structuralH.Count - 2
                For j As Integer = i + 1 To structuralH.Count - 1
                    Dim a = structuralH(i)
                    Dim b = structuralH(j)
                    Dim t As Double = Math.Abs(a.MidY - b.MidY)
                    If t < 0.0005R OrElse t > maxThickness Then Continue For

                    Dim overlapX As Double = Math.Min(a.BboxMaxX, b.BboxMaxX) - Math.Max(a.BboxMinX, b.BboxMinX)
                    If overlapX <= -0.0005 Then Continue For
                    If overlapX < minOverlapX Then Continue For

                    Dim score As Double = overlapX / (1.0R + t)
                    If score > bestScore Then
                        logger?.Log("[DIM][THICKNESS][CANDIDATE] par_H DVLine2d#" & a.Index.ToString(CultureInfo.InvariantCulture) & "/#" & b.Index.ToString(CultureInfo.InvariantCulture) &
                                    " sepY=" & t.ToString("0.######", CultureInfo.InvariantCulture) & " overlapX=" & overlapX.ToString("0.######", CultureInfo.InvariantCulture) &
                                    " score=" & score.ToString("0.######", CultureInfo.InvariantCulture))
                        bestScore = score
                        best = BuildThicknessCandidate(dv, a.Index, b.Index, "vertical", t)
                    End If
                Next
            Next

            If best Is Nothing OrElse best.Object1 Is Nothing OrElse best.Object2 Is Nothing Then
                logger?.Log("[DIM][THICKNESS][RESULT] sin candidato válido tras filtros de espesor")
                Return False
            End If
            If Object.ReferenceEquals(best.Object1, best.Object2) Then Return False

            best.FeatureType = "THICKNESS"
            best.Confidence = Math.Min(0.99R, Math.Max(0.25R, bestScore / 100.0R))
            outFeature = best
            logger?.Log("[DIM][THICKNESS][ACCEPT] eje=" & best.Axis & " valor=" & best.Value.ToString("0.######", CultureInfo.InvariantCulture) &
                        " score=" & bestScore.ToString("0.######", CultureInfo.InvariantCulture))
            logger?.Log("[DIM][THICKNESS][RESULT] par paralelo cercano con solape funcional suficiente")
            Return True
        End Function

        Private Shared Function BuildThicknessCandidate(dv As DrawingView, idx1 As Integer, idx2 As Integer, axis As String, value As Double) As FeatureDimensionCandidate
            Dim c As DVLines2d = Nothing
            Dim l1 As DVLine2d = Nothing
            Dim l2 As DVLine2d = Nothing
            Try
                c = dv.DVLines2d
                If c Is Nothing Then Return Nothing
                l1 = CType(c.Item(idx1), DVLine2d)
                l2 = CType(c.Item(idx2), DVLine2d)
            Catch
                Return Nothing
            End Try
            If l1 Is Nothing OrElse l2 Is Nothing Then Return Nothing

            Dim sig1 As SignificantPoint = DimensioningEngine.SelectPointForAxisExtreme(dv, l1, "DVLine2d#" & idx1.ToString(CultureInfo.InvariantCulture), If(String.Equals(axis, "horizontal", StringComparison.OrdinalIgnoreCase), "x", "y"), True, Nothing)
            Dim sig2 As SignificantPoint = DimensioningEngine.SelectPointForAxisExtreme(dv, l2, "DVLine2d#" & idx2.ToString(CultureInfo.InvariantCulture), If(String.Equals(axis, "horizontal", StringComparison.OrdinalIgnoreCase), "x", "y"), False, Nothing)
            If sig1 Is Nothing OrElse sig2 Is Nothing Then
                Return Nothing
            End If

            Return New FeatureDimensionCandidate With {
                .Object1 = l1,
                .Object2 = l2,
                .Axis = axis,
                .Pick1X = sig1.X,
                .Pick1Y = sig1.Y,
                .Pick2X = sig2.X,
                .Pick2Y = sig2.Y,
                .Value = value,
                .RealPoint1 = sig1,
                .RealPoint2 = sig2
            }
        End Function
    End Class

    Private Shared Function ResolveCandidateObject(c As VerticalFunctionalCandidate, linesCol As DVLines2d, arcsCol As DVArcs2d) As Object
        If c Is Nothing Then Return Nothing
        Select Case c.Kind
            Case "line"
                Return TryGetLineItemSafe(linesCol, c.Index1Based)
            Case "arc"
                Return TryGetArcItemSafe(arcsCol, c.Index1Based)
            Case Else
                Return Nothing
        End Select
    End Function

    Private Shared Function TryGetLineItemSafe(linesCol As DVLines2d, index1Based As Integer) As Object
        If linesCol Is Nothing Then Return Nothing
        Try
            Return CType(linesCol.Item(index1Based), DVLine2d)
        Catch
            Return Nothing
        End Try
    End Function

    Private Shared Function TryGetArcItemSafe(arcsCol As DVArcs2d, index1Based As Integer) As Object
        If arcsCol Is Nothing Then Return Nothing
        Try
            Return CType(arcsCol.Item(index1Based), DVArc2d)
        Catch
            Return Nothing
        End Try
    End Function

    Private Shared Function DescribeDrawingView(dv As DrawingView, log As DimensionLogger) As String
        If dv Is Nothing Then Return "(Nothing)"
        Dim parts As New List(Of String)()

        Try
            parts.Add("DrawingViewType=" & CInt(dv.DrawingViewType).ToString(CultureInfo.InvariantCulture))
        Catch ex As Exception
            log?.ComFail("DrawingView.DrawingViewType", "DrawingView", ex)
        End Try

        Try
            Dim vx As Double, vy As Double, vz As Double, lx As Double, ly As Double, lz As Double
            Dim ori As SolidEdgeConstants.ViewOrientationConstants = SolidEdgeConstants.ViewOrientationConstants.igFrontView
            dv.ViewOrientation(vx, vy, vz, lx, ly, lz, ori)
            parts.Add("ViewOrientation=" & ori.ToString())
        Catch ex As Exception
            log?.ComFail("DrawingView.ViewOrientation", "DrawingView", ex)
        End Try

        If parts.Count = 0 Then Return "(sin metadatos COM)"
        Return String.Join(", ", parts)
    End Function

    Private Shared Function TryGetDrawingViewIsPrimary(dv As DrawingView) As Boolean
        If dv Is Nothing Then Return False
        Try
            Return CBool(CallByName(dv, "IsPrimary", CallType.Get))
        Catch
            Return False
        End Try
    End Function

    ''' <summary>Vista acotada: solo la base del pliego (<c>IsPrimary</c> vía interop si existe; si no, <c>DrawingViews.Item(1)</c>). No se recorren otras vistas para elegir candidatos.</summary>
    Private Shared Function ResolveBaseDrawingViewOnly(sheet As Sheet, log As DimensionLogger, appLogger As Logger) As DrawingView
        Dim views As DrawingViews = Nothing
        Try
            views = sheet.DrawingViews
        Catch ex As Exception
            log.ComFail("Sheet.DrawingViews", "DrawingViews", ex)
            Return Nothing
        End Try
        If views Is Nothing Then Return Nothing

        Dim n As Integer = 0
        Try
            n = views.Count
        Catch ex As Exception
            log.ComFail("DrawingViews.Count", "DrawingViews", ex)
            Return Nothing
        End Try
        If n <= 0 Then Return Nothing

        For i As Integer = 1 To n
            Dim dv As DrawingView = Nothing
            Try
                dv = CType(views.Item(i), DrawingView)
            Catch ex As Exception
                log.ComFail("DrawingViews.Item(" & i.ToString(CultureInfo.InvariantCulture) & ")", "DrawingViews", ex)
                Continue For
            End Try
            If dv Is Nothing Then Continue For
            If TryGetDrawingViewIsPrimary(dv) Then
                log.Info("[DIM] Vista base resuelta por IsPrimary (índice " & i.ToString(CultureInfo.InvariantCulture) & ").")
                appLogger?.Log("[DIM] Vista base resuelta por IsPrimary índice=" & i.ToString(CultureInfo.InvariantCulture))
                Return dv
            End If
        Next

        Try
            Dim first As DrawingView = CType(views.Item(1), DrawingView)
            log.Info("[DIM] Vista base resuelta por DrawingViews.Item(1) (ninguna vista con IsPrimary=True).")
            appLogger?.Log("[DIM] Vista base resuelta por Item(1) (sin IsPrimary=True en interop).")
            Return first
        Catch ex As Exception
            log.ComFail("DrawingViews.Item(1) vista base", "DrawingViews", ex)
            Return Nothing
        End Try
    End Function

End Class
