Option Strict Off

Imports System.Collections.Generic
Imports System.Globalization
Imports System.Linq
Imports System.Reflection
Imports SolidEdgeDraft
Imports SolidEdgeFrameworkSupport
Imports FrameworkDimension = SolidEdgeFrameworkSupport.Dimension

''' <summary>Primera iteración de acotación normalizada (ISO 129-1 inspirada): arquitectura extensible y trazas [DIM][ISO129].</summary>
Public NotInheritable Class Iso129DimensioningService
    Private Sub New()
    End Sub

    Public Shared Sub AplicarNormaAcotacionISO129(
        draftDoc As DraftDocument,
        sheet As Sheet,
        views As List(Of DrawingView),
        Optional config As DimensioningNormConfig = Nothing,
        Optional appLogger As Logger = Nothing)

        If config Is Nothing Then config = DimensioningNormConfig.DefaultConfig()
        Dim Lg = Sub(m As String) appLogger?.Log(m)
        Lg("[DIM][ISO129][ENTER] draft=" & SafeName(draftDoc) & " views_in=" & If(views?.Count, 0).ToString())

        If draftDoc Is Nothing OrElse sheet Is Nothing Then
            Lg("[DIM][ISO129][ERR] draft o sheet Nothing")
            Return
        End If

        Try
            draftDoc.UpdateAll(True)
        Catch ex As Exception
            Lg("[DIM][ISO129][WARN] UpdateAll: " & ex.Message)
        End Try
        DoIdleSafeFromDraft(draftDoc, Lg)

        Dim selected = SeleccionarVistasAcotables(views, config, Lg, maxViews:=3)
        Lg("[DIM][ISO129][VIEW_SELECT] count=" & selected.Count.ToString())

        Dim sheetArea = DimensionPlacementService.TryGetSheetWorkArea(sheet, Lg)
        Dim totalCreated As Integer = 0
        Dim totalKept As Integer = 0
        Dim totalDeleted As Integer = 0

        For Each dv In selected
            If dv Is Nothing Then Continue For
            Dim vname As String = SafeViewName(dv)
            Lg("[DIM][ISO129][VIEW] name=" & vname)

            Dim geom = DimensionGeometryReader.LeerGeometriaDV(dv, config, Lg)
            Dim bbox = DimensionGeometryReader.CalcularBoundingBoxGeometria(geom)
            EnsureNonDegenerateBbox(bbox, Lg)
            Lg(String.Format(CultureInfo.InvariantCulture,
                "[DIM][ISO129][BBOX] view={0} minX={1:0.######} minY={2:0.######} maxX={3:0.######} maxY={4:0.######} w={5:0.######} h={6:0.######}",
                vname, bbox.MinX, bbox.MinY, bbox.MaxX, bbox.MaxY, bbox.Width, bbox.Height))

            Dim rawList = ConstruirCandidatosDeCota(dv, geom, bbox, config, Lg)
            Dim filtered = FiltrarCandidatosISO129(rawList, config, Lg)
            filtered = filtered.OrderBy(Function(c) c.Priority).ToList()

            Dim placementCtx As New Iso129PlacementContext()
            Dim createdRanges As New List(Of String)()
            Dim nThisView As Integer = 0
            Dim viewKept As Integer = 0
            Dim viewDeleted As Integer = 0

            For Each cand In filtered
                If nThisView >= config.MaxDimensionsPerViewInitial Then Exit For
                If Not DimensionValidationService.CumpleNormaISO129(cand, config) Then
                    Lg("[DIM][ISO129][FILTER] skip candidato no cumple norma previa")
                    Continue For
                End If

                DimensionPlacementService.CalcularPosicionExteriorCota(cand, bbox, placementCtx, config, sheetArea, Lg)

                Dim createRes = CrearCotaNormalizada(sheet, dv, cand, bbox, config, appLogger, Lg)
                If Not createRes.Ok OrElse createRes.Dimension Is Nothing Then
                    Lg("[DIM][ISO129][CREATE][FAIL] error=" & If(createRes.ErrorMessage, ""))
                    Continue For
                End If

                totalCreated += 1
                nThisView += 1
                Lg(String.Format(CultureInfo.InvariantCulture,
                    "[DIM][ISO129][CREATE][OK] method={0} connected={1} floatingVisible={2}",
                    If(createRes.MethodUsed, ""), createRes.IsConnected, createRes.FloatingButVisible))

                Dim keep As Boolean = DimensionValidationService.ValidarYRecolocarCota(sheet, createRes.Dimension, cand, bbox, config, Lg)
                If keep Then
                    totalKept += 1
                    viewKept += 1
                    Lg("[DIM][ISO129][KEEP] dim conservada")
                    TryRecordRangeFingerprint(createRes.Dimension, createdRanges, Lg)
                Else
                    totalDeleted += 1
                    viewDeleted += 1
                    Lg("[DIM][ISO129][DELETE] dim eliminada tras validación")
                    TryDeleteDim(createRes.Dimension)
                End If
            Next

            Lg(String.Format(CultureInfo.InvariantCulture,
                "[DIM][ISO129][SUMMARY][VIEW] view={0} attempts_ok={1} kept={2} deleted={3}",
                vname, nThisView, viewKept, viewDeleted))
        Next

        Try
            draftDoc.UpdateAll(True)
        Catch ex As Exception
            Lg("[DIM][ISO129][WARN] UpdateAll final: " & ex.Message)
        End Try
        DoIdleSafeFromDraft(draftDoc, Lg)

        TryArrangeHelpers(draftDoc, selected, Lg)

        Lg(String.Format(CultureInfo.InvariantCulture,
            "[DIM][ISO129][SUMMARY][DOC] created_total={0} kept_total={1} deleted_total={2}",
            totalCreated, totalKept, totalDeleted))
        Lg("[DIM][ISO129][EXIT]")
    End Sub

    Private Shared Sub TryRecordRangeFingerprint(dimObj As Object, ranges As List(Of String), Lg As Action(Of String))
        Dim d As FrameworkDimension = TryCast(dimObj, FrameworkDimension)
        If d Is Nothing Then Return
        Try
            Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
            d.Range(x1, y1, x2, y2)
            ranges.Add(String.Format(CultureInfo.InvariantCulture, "{0:G6}|{1:G6}|{2:G6}|{3:G6}", x1, y1, x2, y2))
        Catch
        End Try
    End Sub

    Private Shared Sub TryArrangeHelpers(draftDoc As DraftDocument, views As List(Of DrawingView), Lg As Action(Of String))
        Lg("[DIM][ISO129][ARRANGE][TRY]")
        Try
            Dim t = draftDoc.GetType()
            Dim m = t.GetMethod("ArrangeDimensionsInSelectSet")
            If m IsNot Nothing Then
                m.Invoke(draftDoc, Nothing)
                Lg("[DIM][ISO129][ARRANGE][OK] ArrangeDimensionsInSelectSet reflection")
                Return
            End If
        Catch ex As Exception
            Lg("[DIM][ISO129][ARRANGE][FAIL] " & ex.Message)
        End Try
        If views Is Nothing Then
            Lg("[DIM][ISO129][ARRANGE][FAIL] no views")
            Return
        End If
        For Each dv In views
            If dv Is Nothing Then Continue For
            Try
                dv.AddConnectedDimensionsToSelectSet()
                Lg("[DIM][ISO129][ARRANGE][OK] AddConnectedDimensionsToSelectSet view=" & SafeViewName(dv))
            Catch ex As Exception
                Lg("[DIM][ISO129][ARRANGE][FAIL] " & ex.Message)
            End Try
        Next
    End Sub

    Private Shared Sub TryDeleteDim(dimObj As Object)
        Dim d As FrameworkDimension = TryCast(dimObj, FrameworkDimension)
        If d Is Nothing Then Return
        Try
            d.Delete()
        Catch
        End Try
    End Sub

    Private Shared Function SeleccionarVistasAcotables(
        views As List(Of DrawingView),
        config As DimensioningNormConfig,
        Lg As Action(Of String),
        maxViews As Integer) As List(Of DrawingView)

        Dim out As New List(Of DrawingView)()
        If views Is Nothing Then Return out

        Dim scored As New List(Of Tuple(Of Integer, DrawingView))()
        For Each dv In views
            If dv Is Nothing Then Continue For
            If IsIsometricDrawingView(dv) Then
                Lg("[DIM][ISO129][VIEW_SELECT] skip iso name=" & SafeViewName(dv))
                Continue For
            End If
            Dim score As Integer = 0
            Try
                score = SafeCount(CallByName(dv, "DVLines2d", CallType.Get))
            Catch
                score = 0
            End Try
            scored.Add(Tuple.Create(-score, dv))
        Next

        scored = scored.OrderBy(Function(t) t.Item1).ToList()
        For i As Integer = 0 To Math.Min(maxViews, scored.Count) - 1
            out.Add(scored(i).Item2)
        Next
        Return out
    End Function

    Private Shared Function SafeCount(col As Object) As Integer
        If col Is Nothing Then Return 0
        Try
            Return CInt(CallByName(col, "Count", CallType.Get))
        Catch
            Return 0
        End Try
    End Function

    Private Shared Function IsIsometricDrawingView(dv As DrawingView) As Boolean
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

    Private Shared Function PickAxisTolerance(bbox As BoundingBox2D) As Double
        Return Math.Max(bbox.Width * 0.001R, 1.0E-7)
    End Function

    Private Class Iso129ExtremeLines
        Public LeftV As DvLineGeomInfo
        Public RightV As DvLineGeomInfo
        Public BottomH As DvLineGeomInfo
        Public TopH As DvLineGeomInfo
    End Class

    Private Shared Function PickExtremeLines(lines As List(Of DvLineGeomInfo), axisTol As Double) As Iso129ExtremeLines
        Dim ex As New Iso129ExtremeLines()
        If lines Is Nothing OrElse lines.Count = 0 Then Return ex

        Dim vert = lines.Where(Function(L) L.Orientation = DvLineOrientationKind.Vertical OrElse Math.Abs(L.Sx2 - L.Sx1) <= axisTol).ToList()
        Dim horiz = lines.Where(Function(L) L.Orientation = DvLineOrientationKind.Horizontal OrElse Math.Abs(L.Sy2 - L.Sy1) <= axisTol).ToList()

        If vert.Count >= 2 Then
            ex.LeftV = vert.OrderBy(Function(L) L.MinXs).First()
            ex.RightV = vert.OrderByDescending(Function(L) L.MaxXs).First()
            If ReferenceEquals(ex.LeftV, ex.RightV) Then
                ex.RightV = vert.Where(Function(L) Not ReferenceEquals(L, ex.LeftV)).OrderByDescending(Function(L) L.MaxXs).FirstOrDefault()
            End If
        End If

        If horiz.Count >= 2 Then
            ex.BottomH = horiz.OrderBy(Function(L) L.MidY).First()
            ex.TopH = horiz.OrderByDescending(Function(L) L.MidY).First()
            If ReferenceEquals(ex.BottomH, ex.TopH) Then
                ex.TopH = horiz.Where(Function(L) Not ReferenceEquals(L, ex.BottomH)).OrderByDescending(Function(L) L.MidY).FirstOrDefault()
            End If
        End If

        If ex.LeftV Is Nothing OrElse ex.RightV Is Nothing Then
            Dim byX = lines.OrderBy(Function(L) L.MidX).ToList()
            If byX.Count >= 2 Then
                ex.LeftV = byX.First()
                ex.RightV = byX.Last()
            End If
        End If
        If ex.BottomH Is Nothing OrElse ex.TopH Is Nothing Then
            Dim byY = lines.OrderBy(Function(L) L.MidY).ToList()
            If byY.Count >= 2 Then
                ex.BottomH = byY.First()
                ex.TopH = byY.Last()
            End If
        End If
        Return ex
    End Function

    Public Shared Function ConstruirCandidatosDeCota(
        view As DrawingView,
        geom As ViewGeometryInfo,
        bbox As BoundingBox2D,
        config As DimensioningNormConfig,
        Lg As Action(Of String)) As List(Of DimensionCandidate)

        Dim list As New List(Of DimensionCandidate)()
        If view Is Nothing OrElse geom Is Nothing OrElse geom.Lines Is Nothing OrElse geom.Lines.Count < 2 Then Return list

        Dim axisTol = PickAxisTolerance(bbox)
        Dim exL = PickExtremeLines(geom.Lines, axisTol)
        Dim leftV = exL.LeftV, rightV = exL.RightV, bottomH = exL.BottomH, topH = exL.TopH

        If leftV IsNot Nothing AndAlso rightV IsNot Nothing AndAlso leftV.Line IsNot Nothing AndAlso rightV.Line IsNot Nothing Then
            Dim spanX As Double = Math.Abs(rightV.MidX - leftV.MidX)
            If spanX > config.MinFeatureSeparation Then
                Dim c As New DimensionCandidate With {
                    .View = view,
                    .Type = DimensionCandidateType.TotalHorizontal,
                    .Orientation = DimensionOrientation.Horizontal,
                    .SourceObject1 = leftV.Line,
                    .SourceObject2 = rightV.Line,
                    .P1 = New Point2D With {.X = leftV.MidX, .Y = leftV.MidY},
                    .P2 = New Point2D With {.X = rightV.MidX, .Y = rightV.MidY},
                    .NominalValue = spanX,
                    .Priority = 10,
                    .IsTotalDimension = True,
                    .UsesHiddenGeometry = leftV.IsHiddenOrNonModel OrElse rightV.IsHiddenOrNonModel,
                    .IsoAuxLine1 = leftV,
                    .IsoAuxLine2 = rightV
                }
                list.Add(c)
                Lg(String.Format(CultureInfo.InvariantCulture,
                    "[DIM][ISO129][CAND][H_TOTAL] spanX={0:0.######} leftIdx={1} rightIdx={2}",
                    spanX, leftV.SourceIndex, rightV.SourceIndex))
            End If
        End If

        If bottomH IsNot Nothing AndAlso topH IsNot Nothing AndAlso bottomH.Line IsNot Nothing AndAlso topH.Line IsNot Nothing Then
            Dim spanY As Double = Math.Abs(topH.MidY - bottomH.MidY)
            If spanY > config.MinFeatureSeparation Then
                Dim c As New DimensionCandidate With {
                    .View = view,
                    .Type = DimensionCandidateType.TotalVertical,
                    .Orientation = DimensionOrientation.Vertical,
                    .SourceObject1 = bottomH.Line,
                    .SourceObject2 = topH.Line,
                    .P1 = New Point2D With {.X = bottomH.MidX, .Y = bottomH.MidY},
                    .P2 = New Point2D With {.X = topH.MidX, .Y = topH.MidY},
                    .NominalValue = spanY,
                    .Priority = 20,
                    .IsTotalDimension = True,
                    .UsesHiddenGeometry = bottomH.IsHiddenOrNonModel OrElse topH.IsHiddenOrNonModel,
                    .IsoAuxLine1 = bottomH,
                    .IsoAuxLine2 = topH
                }
                list.Add(c)
                Lg(String.Format(CultureInfo.InvariantCulture,
                    "[DIM][ISO129][CAND][V_TOTAL] spanY={0:0.######} botIdx={1} topIdx={2}",
                    spanY, bottomH.SourceIndex, topH.SourceIndex))
            End If
        End If

        ' Parciales: hasta 2 pares de verticales interiores
        Dim verts = geom.Lines.Where(Function(L) L.Orientation = DvLineOrientationKind.Vertical OrElse Math.Abs(L.Sx2 - L.Sx1) <= axisTol).OrderBy(Function(L) L.MidX).ToList()
        Dim partialCount As Integer = 0
        For i As Integer = 0 To verts.Count - 2
            If partialCount >= 2 Then Exit For
            Dim a = verts(i)
            Dim b = verts(i + 1)
            If a Is Nothing OrElse b Is Nothing OrElse ReferenceEquals(a, b) Then Continue For
            If ReferenceEquals(a, leftV) AndAlso ReferenceEquals(b, rightV) Then Continue For
            Dim sep As Double = Math.Abs(b.MidX - a.MidX)
            If sep <= config.MinFeatureSeparation Then Continue For
            Dim totalSpan As Double = 0
            If leftV IsNot Nothing AndAlso rightV IsNot Nothing Then totalSpan = Math.Abs(rightV.MidX - leftV.MidX)
            If Math.Abs(sep - totalSpan) < config.MinFeatureSeparation * 3 Then Continue For

            list.Add(New DimensionCandidate With {
                .View = view,
                .Type = DimensionCandidateType.PartialHorizontal,
                .Orientation = DimensionOrientation.Horizontal,
                .SourceObject1 = a.Line,
                .SourceObject2 = b.Line,
                .P1 = New Point2D With {.X = a.MidX, .Y = a.MidY},
                .P2 = New Point2D With {.X = b.MidX, .Y = b.MidY},
                .NominalValue = sep,
                .Priority = 50 + partialCount,
                .IsTotalDimension = False,
                .UsesHiddenGeometry = a.IsHiddenOrNonModel OrElse b.IsHiddenOrNonModel,
                .IsoAuxLine1 = a,
                .IsoAuxLine2 = b
            })
            Lg(String.Format(CultureInfo.InvariantCulture, "[DIM][ISO129][CAND][H_PARTIAL] sep={0:0.######} i={1} j={2}", sep, a.SourceIndex, b.SourceIndex))
            partialCount += 1
        Next

        partialCount = 0
        Dim hors = geom.Lines.Where(Function(L) L.Orientation = DvLineOrientationKind.Horizontal OrElse Math.Abs(L.Sy2 - L.Sy1) <= axisTol).OrderBy(Function(L) L.MidY).ToList()
        For i As Integer = 0 To hors.Count - 2
            If partialCount >= 2 Then Exit For
            Dim a = hors(i)
            Dim b = hors(i + 1)
            If a Is Nothing OrElse b Is Nothing Then Continue For
            If ReferenceEquals(a, bottomH) AndAlso ReferenceEquals(b, topH) Then Continue For
            Dim sep As Double = Math.Abs(b.MidY - a.MidY)
            If sep <= config.MinFeatureSeparation Then Continue For

            list.Add(New DimensionCandidate With {
                .View = view,
                .Type = DimensionCandidateType.PartialVertical,
                .Orientation = DimensionOrientation.Vertical,
                .SourceObject1 = a.Line,
                .SourceObject2 = b.Line,
                .P1 = New Point2D With {.X = a.MidX, .Y = a.MidY},
                .P2 = New Point2D With {.X = b.MidX, .Y = b.MidY},
                .NominalValue = sep,
                .Priority = 60 + partialCount,
                .IsTotalDimension = False,
                .UsesHiddenGeometry = a.IsHiddenOrNonModel OrElse b.IsHiddenOrNonModel,
                .IsoAuxLine1 = a,
                .IsoAuxLine2 = b
            })
            Lg(String.Format(CultureInfo.InvariantCulture, "[DIM][ISO129][CAND][V_PARTIAL] sep={0:0.######} i={1} j={2}", sep, a.SourceIndex, b.SourceIndex))
            partialCount += 1
        Next

        If config.UseRepeatedFeatureNotation Then
            Lg("[DIM][ISO129][CAND] n× repetidos: TODO controlado (sin inventario agujeros aún)")
        End If

        Return list
    End Function

    Public Shared Function FiltrarCandidatosISO129(candidatos As List(Of DimensionCandidate), config As DimensioningNormConfig, Lg As Action(Of String)) As List(Of DimensionCandidate)
        If candidatos Is Nothing Then Return New List(Of DimensionCandidate)()
        Dim out As New List(Of DimensionCandidate)()
        Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Dim tol As Double = Math.Max(config.MinFeatureSeparation, 1.0E-6)

        For Each c In candidatos.OrderBy(Function(x) x.Priority)
            If c Is Nothing Then Continue For
            If c.NominalValue <= config.MinFeatureSeparation Then
                Lg("[DIM][ISO129][FILTER] descartado nominal<=MinFeatureSeparation")
                Continue For
            End If
            If config.AvoidHiddenGeometry AndAlso c.UsesHiddenGeometry Then
                Lg("[DIM][ISO129][FILTER] descartado geometría oculta/no-modelo")
                Continue For
            End If
            If c.IsRedundant Then
                Lg("[DIM][ISO129][FILTER] descartado redundante")
                Continue For
            End If

            Dim key As String = c.Type.ToString() & "|" & Math.Round(c.NominalValue / tol).ToString(CultureInfo.InvariantCulture)
            If config.AvoidDuplicateDimensions AndAlso seen.Contains(key) Then
                Lg("[DIM][ISO129][FILTER] duplicado key=" & key)
                Continue For
            End If
            seen.Add(key)
            out.Add(c)
        Next

        Lg("[DIM][ISO129][FILTER] in=" & candidatos.Count.ToString() & " out=" & out.Count.ToString())
        Return out
    End Function

    Public Shared Function CrearCotaNormalizada(
        sheet As Sheet,
        view As DrawingView,
        cand As DimensionCandidate,
        bbox As BoundingBox2D,
        config As DimensioningNormConfig,
        appLogger As Logger,
        Lg As Action(Of String)) As DimensionCreationResult

        Dim res As New DimensionCreationResult()
        If sheet Is Nothing OrElse view Is Nothing OrElse cand Is Nothing Then
            res.ErrorMessage = "Nothing"
            Return res
        End If

        Dim dimLog As New DimensionLogger(appLogger)

        Dim dimsObj As Object = Nothing
        Try
            dimsObj = sheet.Dimensions
        Catch ex As Exception
            res.ErrorMessage = ex.Message
            Return res
        End Try
        Dim dims As Dimensions = TryCast(dimsObj, Dimensions)
        If dims Is Nothing Then
            res.ErrorMessage = "Dimensions Nothing"
            Return res
        End If

        Dim vsb As New ViewSheetBoundingBox With {
            .MinX = bbox.MinX,
            .MinY = bbox.MinY,
            .MaxX = bbox.MaxX,
            .MaxY = bbox.MaxY
        }
        Dim frame As ViewPlacementFrame = Nothing
        If Not ViewPlacementFrame.TryCreateFromBaseViewSheetBox(vsb, dimLog, frame) Then
            res.ErrorMessage = "frame"
            Return res
        End If

        Dim place = cand.PlacementPoint
        If place Is Nothing Then place = New Point2D()

        Dim L1 = TryCast(cand.IsoAuxLine1, DvLineGeomInfo)
        Dim L2 = TryCast(cand.IsoAuxLine2, DvLineGeomInfo)

        Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double

        If cand.Orientation = DimensionOrientation.Horizontal Then
            Dim yLine As Double = place.Y
            If L1 IsNot Nothing AndAlso L2 IsNot Nothing Then
                yLine = DimensionPlacementService.ClampPickYToVerticalLine(place.Y, L1)
                yLine = (yLine + DimensionPlacementService.ClampPickYToVerticalLine(place.Y, L2)) / 2.0R
            End If
            x1 = If(L1?.MidX, cand.P1.X)
            x2 = If(L2?.MidX, cand.P2.X)
            y1 = yLine
            y2 = yLine
        Else
            Dim xLine As Double = place.X
            If L1 IsNot Nothing AndAlso L2 IsNot Nothing Then
                xLine = DimensionPlacementService.ClampPickXToHorizontalLine(place.X, L1)
                xLine = (xLine + DimensionPlacementService.ClampPickXToHorizontalLine(place.X, L2)) / 2.0R
            End If
            y1 = If(L1?.MidY, cand.P1.Y)
            y2 = If(L2?.MidY, cand.P2.Y)
            x1 = xLine
            x2 = xLine
        End If

        Lg(String.Format(CultureInfo.InvariantCulture,
            "[DIM][ISO129][CREATE][TRY] type={0} method=AddDistanceBetweenObjects picks=({1:0.######},{2:0.######})-({3:0.######},{4:0.######})",
            cand.Type, x1, y1, x2, y2))

        Dim axis As String = If(cand.Orientation = DimensionOrientation.Horizontal, "horizontal", "vertical")
        Dim methodUsed As String = Nothing
        Dim created As FrameworkDimension = Nothing
        Dim wasFallback As Boolean = False

        Try
            Dim ok1 As Boolean = DimensionPlacementEngine.TryAddDistanceBetweenObjectsGeneric(
                dims, cand.SourceObject1, cand.SourceObject2, x1, y1, x2, y2, dimLog, axis, methodUsed,
                "ISO129_" & cand.Type.ToString(), frame, view,
                skipInsertedDimensionSpatialSanity:=True, createdDimensionOut:=created)
            If ok1 AndAlso created IsNot Nothing Then
                res.Ok = True
                res.Dimension = created
                res.MethodUsed = methodUsed
                If methodUsed IsNot Nothing AndAlso methodUsed.IndexOf("(view)", StringComparison.OrdinalIgnoreCase) >= 0 Then
                    res.WasFallback = True
                    Lg("[DIM][ISO129][CREATE][FALLBACK] reason=view_space_after_sheet_failed")
                Else
                    res.WasFallback = wasFallback
                End If
                res.FloatingButVisible = config.AutoDimPragmaticVisibleFirst
                res.IsConnected = ProbeConnected(created, Lg)
                Lg(String.Format(CultureInfo.InvariantCulture,
                    "[DIM][ISO129][CREATE][OK] connected={0} floatingVisible={1}",
                    res.IsConnected, res.FloatingButVisible))
                Return res
            End If
        Catch ex As Exception
            res.ErrorMessage = ex.Message
            Lg("[DIM][ISO129][CREATE][FAIL] " & ex.Message)
            Return res
        End Try

        res.ErrorMessage = "AddDistanceBetweenObjects falló"
        Lg("[DIM][ISO129][CREATE][FAIL] " & res.ErrorMessage)
        Return res
    End Function

    Private Shared Function ProbeConnected(d As FrameworkDimension, Lg As Action(Of String)) As Boolean
        If d Is Nothing Then Return False
        Try
            Dim st = d.UpdateStatus()
            Lg("[DIM][ISO129][CREATE] UpdateStatus=" & st.ToString())
            Return True
        Catch
            Return False
        End Try
    End Function

    Private Shared Sub DoIdleSafeFromDraft(draft As DraftDocument, Lg As Action(Of String))
        Dim app As Object = Nothing
        Try
            app = draft.Application
        Catch
        End Try
        If app Is Nothing Then Return
        Try
            CallByName(app, "DoIdle", CallType.Method)
            Lg("[DIM][ISO129] DoIdle OK")
        Catch ex As Exception
            Lg("[DIM][ISO129] DoIdle: " & ex.Message)
        End Try
    End Sub

    Private Shared Function SafeName(doc As DraftDocument) As String
        If doc Is Nothing Then Return ""
        Try
            Return Convert.ToString(CallByName(doc, "Name", CallType.Get), CultureInfo.InvariantCulture)
        Catch
            Return ""
        End Try
    End Function

    Private Shared Function SafeViewName(dv As DrawingView) As String
        If dv Is Nothing Then Return ""
        Try
            Return Convert.ToString(CallByName(dv, "Name", CallType.Get), CultureInfo.InvariantCulture)
        Catch
            Return ""
        End Try
    End Function

    Private Shared Sub EnsureNonDegenerateBbox(bbox As BoundingBox2D, Lg As Action(Of String))
        If bbox Is Nothing Then Return
        Dim eps As Double = 0.001R
        If bbox.Width <= 1.0E-9 OrElse bbox.Height <= 1.0E-9 Then
            bbox.MaxX = bbox.MinX + eps
            bbox.MaxY = bbox.MinY + eps
            Lg("[DIM][ISO129][BBOX] expand degenerate geometry bbox (+1mm)")
        End If
    End Sub
End Class
