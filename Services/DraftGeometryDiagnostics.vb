Option Strict Off

Imports System.Collections.Generic
Imports System.IO
Imports System.Globalization
Imports System.Linq
Imports System.Text
Imports SolidEdgeDraft
Imports SolidEdgeFramework

''' <summary>
''' Diagnóstico geométrico del Draft: solo lectura tipada (interop SolidEdgeDraft), prefijos [GEOM].
''' No usa reflexión ni CallByName para colecciones 2D. No inserta cotas.
''' </summary>
Public NotInheritable Class DraftGeometryDiagnostics

    Public Shared MaxEntitiesDetailedPerType As Integer = 250

    Private Shared ReadOnly TwoPi As Double = 2.0R * Math.PI

    Private Sub New()
    End Sub

    Public Shared Function Run(draft As DraftDocument, logger As Logger, Optional config As JobConfiguration = Nothing) As List(Of ViewGeometrySummary)
        Dim outList As New List(Of ViewGeometrySummary)()
        If draft Is Nothing OrElse logger Is Nothing Then Return outList

        Dim opt = ResolveOptions(config)

        LogGeomDocHeader(draft, logger)
        LogActiveSheetViewsGeometry(draft, logger, opt, outList)
        If config IsNot Nothing Then
            Try
                TryWriteGeomReportFile(logger, config, outList)
            Catch exFile As Exception
                logger.Log("[GEOM][FILE][ERR] No se pudo escribir fichero de diagnóstico: " & exFile.Message)
            End Try
        End If
        logger.Log("[GEOM][DOC] Fin diagnóstico geométrico.")
        Return outList
    End Function

    Private Class DiagOptions
        Public TolAxisWidthFraction As Double = 0.001R
        Public TolAxisMinM As Double = 0.0000001R
        Public ArcSmallRadiusM As Double = 0.008R
        Public ArcLargeRadiusM As Double = 0.03R
        Public HoleCandidateDiamM As Double = 0.025R
    End Class

    Private Shared Function ResolveOptions(cfg As JobConfiguration) As DiagOptions
        Dim o As New DiagOptions()
        If cfg Is Nothing Then Return o
        o.TolAxisWidthFraction = cfg.GeometryDiagnosticsTolAxisWidthFraction
        o.TolAxisMinM = cfg.GeometryDiagnosticsTolAxisMinM
        o.ArcSmallRadiusM = cfg.GeometryDiagnosticsArcSmallRadiusM
        o.ArcLargeRadiusM = cfg.GeometryDiagnosticsArcLargeRadiusM
        o.HoleCandidateDiamM = cfg.GeometryDiagnosticsHoleCandidateDiamM
        Return o
    End Function

    Private Shared Sub LogGeomDocHeader(draft As DraftDocument, logger As Logger)
        Dim doc As Object = CObj(draft)
        Dim full As String = ""
        Dim name As String = ""
        Try : full = CStr(doc.FullName) : Catch : End Try
        Try : name = CStr(doc.Name) : Catch : End Try
        logger.Log("[GEOM][DOC] Inicio (acceso tipado a DV*2d; sin reflexión en geometría).")
        logger.Log("[GEOM][DOC] FullName=" & full & " Name=" & name)
        Dim activeName As String = ""
        Try
            Dim sh As Sheet = draft.ActiveSheet
            If sh IsNot Nothing Then activeName = CStr(sh.Name)
        Catch ex As Exception
            activeName = "(error: " & ex.Message & ")"
        End Try
        logger.Log("[GEOM][DOC] ActiveSheet.Name=" & activeName)
    End Sub

    Private Shared Sub LogActiveSheetViewsGeometry(draft As DraftDocument, logger As Logger, opt As DiagOptions, summaries As List(Of ViewGeometrySummary))
        Dim sheet As Sheet = Nothing
        Try
            sheet = draft.ActiveSheet
        Catch ex As Exception
            logger.Log("[GEOM][VIEW] No se pudo leer ActiveSheet: " & ex.Message)
            Return
        End Try
        If sheet Is Nothing Then
            logger.Log("[GEOM][VIEW] ActiveSheet Nothing.")
            Return
        End If

        Dim dvs As DrawingViews = Nothing
        Try
            dvs = sheet.DrawingViews
        Catch ex As Exception
            logger.Log("[GEOM][VIEW] DrawingViews: " & ex.Message)
            Return
        End Try
        If dvs Is Nothing Then
            logger.Log("[GEOM][VIEW] DrawingViews Nothing.")
            Return
        End If

        Dim vc As Integer = 0
        Try : vc = dvs.Count : Catch : End Try
        logger.Log("[GEOM][VIEW] Hoja activa DrawingViews.Count=" & vc.ToString(CultureInfo.InvariantCulture))

        For vi As Integer = 1 To vc
            Dim dv As DrawingView = Nothing
            Try
                dv = CType(dvs.Item(vi), DrawingView)
            Catch ex As Exception
                logger.Log("[GEOM][VIEW] Item(" & vi.ToString(CultureInfo.InvariantCulture) & ") no legible: " & ex.Message)
                Continue For
            End Try
            If dv Is Nothing Then Continue For

            Try
                dv.Update()
            Catch
            End Try

            Dim sum As New ViewGeometrySummary()
            FillViewHeader(vi, dv, sum, logger)
            Dim tolAxis As Double = Math.Max(sum.ViewWidth * opt.TolAxisWidthFraction, opt.TolAxisMinM)
            sum.TolAxisM = tolAxis
            logger.Log("[GEOM][TOL] Vista#" & vi.ToString(CultureInfo.InvariantCulture) &
                      " tolAxisM=" & FormatInv(tolAxis) &
                      " (max(anchoVista*" & FormatInv(opt.TolAxisWidthFraction) & "," & FormatInv(opt.TolAxisMinM) & "))")

            Dim boxes As New List(Of LabeledBBox)()
            FillCollectionCountsTyped(vi, dv, sum, logger)
            ProcessLinesTyped(vi, dv, sum, tolAxis, opt, logger, boxes)
            ProcessArcsTyped(vi, dv, sum, opt, logger, boxes)
            ProcessCirclesTyped(vi, dv, sum, opt, logger, boxes)
            ProcessEllipsesTyped(vi, dv, sum, logger, boxes)
            ProcessPointsTyped(vi, dv, sum, logger, boxes)
            ProcessLineStringsTyped(vi, dv, sum, logger, boxes)
            ProcessBSplinesTyped(vi, dv, sum, logger, boxes)

            FinalizeExtentsAndClassifyLog(vi, sum, boxes, tolAxis, logger)
            summaries.Add(sum)
        Next
    End Sub

    Private Class LabeledBBox
        Public Label As String
        Public MinX As Double
        Public MaxX As Double
        Public MinY As Double
        Public MaxY As Double
        Public Obj As Object
        Public PickX As Double
        Public PickY As Double
    End Class

    Private Shared Sub FillCollectionCountsTyped(viewIdx As Integer, dv As DrawingView, sum As ViewGeometrySummary, logger As Logger)
        Dim sb As New StringBuilder()
        sb.Append("[GEOM][COUNT] Vista#" & viewIdx.ToString(CultureInfo.InvariantCulture))

        Try
            Dim c As DVLines2d = dv.DVLines2d
            sum.Counts.DVLines2d = If(c Is Nothing, 0, c.Count)
        Catch ex As Exception
            sum.Counts.DVLines2d = 0
            logger.Log("[GEOM][COUNT] Vista#" & viewIdx.ToString(CultureInfo.InvariantCulture) & " DVLines2d: no disponible — " & ex.Message)
        End Try

        Try
            Dim c As DVArcs2d = dv.DVArcs2d
            sum.Counts.DVArcs2d = If(c Is Nothing, 0, c.Count)
        Catch ex As Exception
            sum.Counts.DVArcs2d = 0
            logger.Log("[GEOM][COUNT] Vista#" & viewIdx.ToString(CultureInfo.InvariantCulture) & " DVArcs2d: no disponible — " & ex.Message)
        End Try

        Try
            Dim c As DVCircles2d = dv.DVCircles2d
            sum.Counts.DVCircles2d = If(c Is Nothing, 0, c.Count)
        Catch ex As Exception
            sum.Counts.DVCircles2d = 0
            logger.Log("[GEOM][COUNT] Vista#" & viewIdx.ToString(CultureInfo.InvariantCulture) & " DVCircles2d: no disponible — " & ex.Message)
        End Try

        Try
            Dim c As DVEllipses2d = dv.DVEllipses2d
            sum.Counts.DVEllipses2d = If(c Is Nothing, 0, c.Count)
        Catch ex As Exception
            sum.Counts.DVEllipses2d = 0
            logger.Log("[GEOM][COUNT] Vista#" & viewIdx.ToString(CultureInfo.InvariantCulture) & " DVEllipses2d: no disponible — " & ex.Message)
        End Try

        Try
            Dim c As DVPoints2d = dv.DVPoints2d
            sum.Counts.DVPoints2d = If(c Is Nothing, 0, c.Count)
        Catch ex As Exception
            sum.Counts.DVPoints2d = 0
            logger.Log("[GEOM][COUNT] Vista#" & viewIdx.ToString(CultureInfo.InvariantCulture) & " DVPoints2d: no disponible — " & ex.Message)
        End Try

        Try
            Dim c As DVLineStrings2d = dv.DVLineStrings2d
            sum.Counts.DVLineStrings2d = If(c Is Nothing, 0, c.Count)
        Catch ex As Exception
            sum.Counts.DVLineStrings2d = 0
            logger.Log("[GEOM][COUNT] Vista#" & viewIdx.ToString(CultureInfo.InvariantCulture) & " DVLineStrings2d: no disponible — " & ex.Message)
        End Try

        Try
            Dim c As DVBSplineCurves2d = dv.DVBSplineCurves2d
            sum.Counts.DVBSplineCurves2d = If(c Is Nothing, 0, c.Count)
        Catch ex As Exception
            sum.Counts.DVBSplineCurves2d = 0
            logger.Log("[GEOM][COUNT] Vista#" & viewIdx.ToString(CultureInfo.InvariantCulture) & " DVBSplineCurves2d: no disponible — " & ex.Message)
        End Try

        sb.Append(" DVLines2d=").Append(sum.Counts.DVLines2d.ToString(CultureInfo.InvariantCulture))
        sb.Append(" DVArcs2d=").Append(sum.Counts.DVArcs2d.ToString(CultureInfo.InvariantCulture))
        sb.Append(" DVCircles2d=").Append(sum.Counts.DVCircles2d.ToString(CultureInfo.InvariantCulture))
        sb.Append(" DVEllipses2d=").Append(sum.Counts.DVEllipses2d.ToString(CultureInfo.InvariantCulture))
        sb.Append(" DVPoints2d=").Append(sum.Counts.DVPoints2d.ToString(CultureInfo.InvariantCulture))
        sb.Append(" DVLineStrings2d=").Append(sum.Counts.DVLineStrings2d.ToString(CultureInfo.InvariantCulture))
        sb.Append(" DVBSplineCurves2d=").Append(sum.Counts.DVBSplineCurves2d.ToString(CultureInfo.InvariantCulture))
        logger.Log(sb.ToString())
    End Sub

    Private Shared Sub FillViewHeader(idx As Integer, dv As DrawingView, sum As ViewGeometrySummary, logger As Logger)
        sum.ViewIndex = idx
        Try : sum.Name = CStr(dv.Name) : Catch : sum.Name = "?" : End Try
        Try : sum.DrawingViewType = dv.DrawingViewType.ToString() : Catch : sum.DrawingViewType = "?" : End Try
        Try
            Dim vx As Double, vy As Double, vz As Double, lx As Double, ly As Double, lz As Double
            Dim vo As ViewOrientationConstants
            dv.ViewOrientation(vx, vy, vz, lx, ly, lz, vo)
            sum.ViewOrientation = vo.ToString()
        Catch
            sum.ViewOrientation = "?"
        End Try

        sum.ScaleFactor = 1.0R
        Try
            sum.ScaleFactor = CDbl(dv.Scale)
        Catch
        End Try

        sum.RotationNote = "(n/d)"
        Try
            sum.RotationNote = FormatInv(CDbl(dv.Angle))
        Catch
        End Try

        sum.HasOrigin = TryGetOrigin(dv, sum.OriginSheetX, sum.OriginSheetY)
        Dim orgStr As String = If(sum.HasOrigin, FormatInv(sum.OriginSheetX) & "," & FormatInv(sum.OriginSheetY), "(n/d)")

        Dim rngOk As Boolean = False
        Try
            Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
            dv.Range(x1, y1, x2, y2)
            sum.RangeMinX = Math.Min(x1, x2)
            sum.RangeMinY = Math.Min(y1, y2)
            sum.RangeMaxX = Math.Max(x1, x2)
            sum.RangeMaxY = Math.Max(y1, y2)
            sum.ViewWidth = Math.Abs(sum.RangeMaxX - sum.RangeMinX)
            sum.ViewHeight = Math.Abs(sum.RangeMaxY - sum.RangeMinY)
            rngOk = True
        Catch ex As Exception
            logger.Log("[GEOM][VIEW] Vista#" & idx.ToString(CultureInfo.InvariantCulture) & " Range error: " & ex.Message)
        End Try

        sum.HasModelLink = False
        sum.ModelLinkFileName = ""
        Try
            Dim ml As ModelLink = dv.ModelLink
            sum.HasModelLink = (ml IsNot Nothing)
            If ml IsNot Nothing Then
                Try : sum.ModelLinkFileName = CStr(ml.FileName) : Catch : sum.ModelLinkFileName = "(ilegible)" : End Try
            End If
        Catch
        End Try

        Dim rngStr As String = If(rngOk,
            "RangeSheet=(" & FormatInv(sum.RangeMinX) & "," & FormatInv(sum.RangeMinY) & ")-(" & FormatInv(sum.RangeMaxX) & "," & FormatInv(sum.RangeMaxY) & ") W=" & FormatInv(sum.ViewWidth) & " H=" & FormatInv(sum.ViewHeight),
            "Range=(n/d)")

        Dim mlTail As String = If(String.IsNullOrEmpty(sum.ModelLinkFileName), "", " ModelLinkFile=" & sum.ModelLinkFileName)
        logger.Log("[GEOM][VIEW] --- Vista " & idx.ToString(CultureInfo.InvariantCulture) & " ---")
        logger.Log("[GEOM][VIEW] Name=" & sum.Name &
                   " DrawingViewType=" & sum.DrawingViewType &
                   " ViewOrientation=" & sum.ViewOrientation &
                   " ScaleFactor=" & FormatInv(sum.ScaleFactor) &
                   " Rotation=" & sum.RotationNote &
                   " OriginSheet=" & orgStr &
                   " " & rngStr &
                   " ModelLink=" & sum.HasModelLink.ToString() & mlTail)
    End Sub

    Private Shared Sub ProcessLinesTyped(idx As Integer, dv As DrawingView, sum As ViewGeometrySummary, tolAxis As Double, opt As DiagOptions, logger As Logger, boxes As List(Of LabeledBBox))
        Dim linesCol As DVLines2d = Nothing
        Try
            linesCol = dv.DVLines2d
        Catch
            Return
        End Try
        If linesCol Is Nothing Then Return

        Dim n As Integer = 0
        Try : n = linesCol.Count : Catch : Return : End Try
        Dim logged As Integer = 0
        For i As Integer = 1 To n
            Dim ln As DVLine2d = Nothing
            Try
                ln = CType(linesCol.Item(i), DVLine2d)
            Catch
                Continue For
            End Try
            If ln Is Nothing Then Continue For

            Dim vx1 As Double, vy1 As Double, vx2 As Double, vy2 As Double
            Try
                ln.GetStartPoint(vx1, vy1)
                ln.GetEndPoint(vx2, vy2)
            Catch
                Continue For
            End Try

            Dim sx1 As Double, sy1 As Double, sx2 As Double, sy2 As Double
            Try
                dv.ViewToSheet(vx1, vy1, sx1, sy1)
                dv.ViewToSheet(vx2, vy2, sx2, sy2)
            Catch
                Continue For
            End Try

            Dim dx As Double = sx2 - sx1
            Dim dy As Double = sy2 - sy1
            Dim len As Double = Math.Sqrt(dx * dx + dy * dy)
            Dim orient As String = ClassifyLineAxis(dx, dy, len, tolAxis)
            Dim ang As Double = If(len > 1.0E-12, Math.Atan2(dy, dx), 0R)
            Dim info As New LineGeometryInfo With {
                .Index = i,
                .X1 = sx1, .Y1 = sy1, .X2 = sx2, .Y2 = sy2,
                .Length = len, .DeltaX = dx, .DeltaY = dy,
                .AngleRadians = ang,
                .AngleDegrees = ang * 180.0R / Math.PI,
                .Orientation = orient,
                .MidX = (sx1 + sx2) / 2.0R, .MidY = (sy1 + sy2) / 2.0R,
                .BboxMinX = Math.Min(sx1, sx2), .BboxMaxX = Math.Max(sx1, sx2),
                .BboxMinY = Math.Min(sy1, sy2), .BboxMaxY = Math.Max(sy1, sy2)
            }
            sum.Lines.Add(info)
            boxes.Add(New LabeledBBox With {
                .Label = "DVLine2d#" & i.ToString(CultureInfo.InvariantCulture),
                .MinX = info.BboxMinX, .MaxX = info.BboxMaxX, .MinY = info.BboxMinY, .MaxY = info.BboxMaxY,
                .Obj = ln,
                .PickX = info.MidX, .PickY = info.MidY
            })

            Select Case orient
                Case "horizontal" : sum.ClassifyLinesH += 1
                Case "vertical" : sum.ClassifyLinesV += 1
                Case "inclinada" : sum.ClassifyLinesI += 1
            End Select

            If logged < MaxEntitiesDetailedPerType Then
                logger.Log("[GEOM][LINE] Vista#" & idx.ToString(CultureInfo.InvariantCulture) &
                           " idx=" & i.ToString(CultureInfo.InvariantCulture) &
                           " sheet=(" & FormatInv(sx1) & "," & FormatInv(sy1) & ")-(" & FormatInv(sx2) & "," & FormatInv(sy2) & ")" &
                           " L=" & FormatInv(len) & " dX=" & FormatInv(dx) & " dY=" & FormatInv(dy) &
                           " angRad=" & FormatInv(ang) & " cls=" & orient &
                           " mid=(" & FormatInv(info.MidX) & "," & FormatInv(info.MidY) & ")" &
                           " xmin=" & FormatInv(info.BboxMinX) & " xmax=" & FormatInv(info.BboxMaxX) &
                           " ymin=" & FormatInv(info.BboxMinY) & " ymax=" & FormatInv(info.BboxMaxY))
                logged += 1
            End If
        Next
        If n > MaxEntitiesDetailedPerType Then
            logger.Log("[GEOM][LINE] Vista#" & idx.ToString(CultureInfo.InvariantCulture) & " ... truncado " & MaxEntitiesDetailedPerType.ToString(CultureInfo.InvariantCulture) & "/" & n.ToString(CultureInfo.InvariantCulture))
        End If
    End Sub

    Private Shared Function ClassifyLineAxis(dx As Double, dy As Double, len As Double, tolAxis As Double) As String
        If len < 1.0E-12 Then Return "degenerada"
        If Math.Abs(dx) < tolAxis AndAlso Math.Abs(dy) < tolAxis Then Return "degenerada"
        If Math.Abs(dy) <= tolAxis Then Return "horizontal"
        If Math.Abs(dx) <= tolAxis Then Return "vertical"
        Return "inclinada"
    End Function

    Private Shared Sub ProcessArcsTyped(idx As Integer, dv As DrawingView, sum As ViewGeometrySummary, opt As DiagOptions, logger As Logger, boxes As List(Of LabeledBBox))
        Dim arcsCol As DVArcs2d = Nothing
        Try
            arcsCol = dv.DVArcs2d
        Catch
            Return
        End Try
        If arcsCol Is Nothing Then Return

        Dim n As Integer = 0
        Try : n = arcsCol.Count : Catch : Return : End Try
        Dim logged As Integer = 0
        For i As Integer = 1 To n
            Dim a As DVArc2d = Nothing
            Try
                a = CType(arcsCol.Item(i), DVArc2d)
            Catch
                Continue For
            End Try
            If a Is Nothing Then Continue For

            Dim cx As Double, cy As Double, rad As Double
            Dim okC As Boolean = TryArcCenterRadius(a, cx, cy, rad)
            Dim startRad As Double, sweepRad As Double
            Dim angNote As String = ""
            Dim okAng As Boolean = TryGetArcAnglesRadians(a, startRad, sweepRad, angNote)
            Dim endRad As Double = If(okAng, startRad + sweepRad, 0R)

            Dim sx0 As Double = 0, sy0 As Double = 0, ex0 As Double = 0, ey0 As Double = 0
            Dim okSe As Boolean = False
            Try
                Dim vx1 As Double, vy1 As Double, vx2 As Double, vy2 As Double
                a.GetStartPoint(vx1, vy1)
                a.GetEndPoint(vx2, vy2)
                dv.ViewToSheet(vx1, vy1, sx0, sy0)
                dv.ViewToSheet(vx2, vy2, ex0, ey0)
                okSe = True
            Catch
            End Try

            Dim csx As Double = 0, csy As Double = 0
            If okC Then
                Try
                    dv.ViewToSheet(cx, cy, csx, csy)
                Catch
                    okC = False
                End Try
            End If

            Dim bbox As Tuple(Of Double, Double, Double, Double) = ComputeArcSheetBBox(dv, a, cx, cy, rad, startRad, sweepRad, okC, okAng)
            Dim bminx As Double = bbox.Item1, bmaxx As Double = bbox.Item2, bminy As Double = bbox.Item3, bmaxy As Double = bbox.Item4

            Dim arcLen As Double = 0R
            If okC AndAlso rad > 1.0E-12 AndAlso okAng Then
                arcLen = Math.Abs(sweepRad) * rad
            End If

            Dim arcCls As String = "arco_abierto"
            If okAng AndAlso Math.Abs(Math.Abs(sweepRad) - TwoPi) < 0.02R Then
                arcCls = "arco_cerrado≈360°"
            End If

            If okC Then
                If rad < opt.ArcSmallRadiusM Then
                    sum.ClassifyArcSmall += 1
                ElseIf rad >= opt.ArcLargeRadiusM Then
                    sum.ClassifyArcLarge += 1
                Else
                    sum.ClassifyArcOther += 1
                End If
            Else
                sum.ClassifyArcOther += 1
            End If
            sum.ClassifyCurveOpen += If(arcCls.StartsWith("arco_abierto", StringComparison.OrdinalIgnoreCase), 1, 0)
            If arcCls <> "arco_abierto" Then sum.ClassifyCurveClosed += 1

            Dim info As New ArcGeometryInfo With {
                .Index = i,
                .CenterSheetX = csx, .CenterSheetY = csy, .Radius = rad,
                .StartSheetX = sx0, .StartSheetY = sy0, .EndSheetX = ex0, .EndSheetY = ey0,
                .StartAngleRaw = startRad, .EndAngleRaw = endRad, .SweepAngleRaw = sweepRad, .AngleUnitNote = angNote,
                .SweepSignNote = If(okAng, "sign(sweep)=" & Math.Sign(sweepRad).ToString(CultureInfo.InvariantCulture), "?"),
                .ArcLengthApprox = arcLen,
                .BboxMinX = bminx, .BboxMaxX = bmaxx, .BboxMinY = bminy, .BboxMaxY = bmaxy,
                .CurveClass = arcCls,
                .ApiNote = If(okC AndAlso okSe, "", "centro/radio o extremos parcialmente no leídos")
            }
            sum.Arcs.Add(info)
            boxes.Add(New LabeledBBox With {
                .Label = "DVArc2d#" & i.ToString(CultureInfo.InvariantCulture),
                .MinX = bminx, .MaxX = bmaxx, .MinY = bminy, .MaxY = bmaxy,
                .Obj = a,
                .PickX = (bminx + bmaxx) / 2.0R, .PickY = (bminy + bmaxy) / 2.0R
            })

            If logged < MaxEntitiesDetailedPerType Then
                logger.Log("[GEOM][ARC] Vista#" & idx.ToString(CultureInfo.InvariantCulture) &
                           " idx=" & i.ToString(CultureInfo.InvariantCulture) &
                           " centroSheet=(" & FormatInv(csx) & "," & FormatInv(csy) & ") R=" & FormatInv(rad) &
                           " P0=(" & FormatInv(sx0) & "," & FormatInv(sy0) & ") P1=(" & FormatInv(ex0) & "," & FormatInv(ey0) & ")" &
                           " startRad=" & FormatInv(startRad) & " endRad=" & FormatInv(endRad) & " sweepRad=" & FormatInv(sweepRad) & " (" & angNote & ")" &
                           " Larc≈" & FormatInv(arcLen) &
                           " xmin=" & FormatInv(bminx) & " xmax=" & FormatInv(bmaxx) & " ymin=" & FormatInv(bminy) & " ymax=" & FormatInv(bmaxy) &
                           " cls=" & arcCls)
                logged += 1
            End If
        Next
        If n > MaxEntitiesDetailedPerType Then
            logger.Log("[GEOM][ARC] Vista#" & idx.ToString(CultureInfo.InvariantCulture) & " ... truncado detalle")
        End If
    End Sub

    Private Shared Sub ExpandBBoxSheet(ByRef has As Boolean, ByRef minX As Double, ByRef maxX As Double, ByRef minY As Double, ByRef maxY As Double, sx As Double, sy As Double)
        If Not has Then
            minX = sx : maxX = sx : minY = sy : maxY = sy : has = True
        Else
            minX = Math.Min(minX, sx) : maxX = Math.Max(maxX, sx)
            minY = Math.Min(minY, sy) : maxY = Math.Max(maxY, sy)
        End If
    End Sub

    Private Shared Function ComputeArcSheetBBox(dv As DrawingView, a As DVArc2d, cx As Double, cy As Double, rad As Double,
                                                startRad As Double, sweepRad As Double, okC As Boolean, okAng As Boolean) As Tuple(Of Double, Double, Double, Double)
        Dim minX As Double = Double.MaxValue, maxX As Double = Double.MinValue, minY As Double = Double.MaxValue, maxY As Double = Double.MinValue
        Dim has As Boolean = False

        Try
            Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
            a.Range(x1, y1, x2, y2)
            Dim corners As Tuple(Of Double, Double)() = New Tuple(Of Double, Double)() {
                Tuple.Create(x1, y1),
                Tuple.Create(x2, y1),
                Tuple.Create(x1, y2),
                Tuple.Create(x2, y2)
            }
            For Each p As Tuple(Of Double, Double) In corners
                Dim sx As Double, sy As Double
                dv.ViewToSheet(p.Item1, p.Item2, sx, sy)
                ExpandBBoxSheet(has, minX, maxX, minY, maxY, sx, sy)
            Next
        Catch
        End Try

        If okC AndAlso rad > 1.0E-12 AndAlso okAng Then
            Dim steps As Integer = 24
            For k As Integer = 0 To steps
                Dim t As Double = startRad + sweepRad * (k / steps)
                Dim vx As Double = cx + rad * Math.Cos(t)
                Dim vy As Double = cy + rad * Math.Sin(t)
                Dim sx As Double, sy As Double
                Try
                    dv.ViewToSheet(vx, vy, sx, sy)
                    ExpandBBoxSheet(has, minX, maxX, minY, maxY, sx, sy)
                Catch
                End Try
            Next
        End If

        Try
            Dim vx1 As Double, vy1 As Double, vx2 As Double, vy2 As Double
            a.GetStartPoint(vx1, vy1)
            a.GetEndPoint(vx2, vy2)
            Dim sx1 As Double, sy1 As Double, sx2 As Double, sy2 As Double
            dv.ViewToSheet(vx1, vy1, sx1, sy1)
            dv.ViewToSheet(vx2, vy2, sx2, sy2)
            ExpandBBoxSheet(has, minX, maxX, minY, maxY, sx1, sy1)
            ExpandBBoxSheet(has, minX, maxX, minY, maxY, sx2, sy2)
        Catch
        End Try

        If Not has Then Return Tuple.Create(0R, 0R, 0R, 0R)
        Return Tuple.Create(minX, maxX, minY, maxY)
    End Function

    Private Shared Sub ProcessCirclesTyped(idx As Integer, dv As DrawingView, sum As ViewGeometrySummary, opt As DiagOptions, logger As Logger, boxes As List(Of LabeledBBox))
        Dim circCol As DVCircles2d = Nothing
        Try
            circCol = dv.DVCircles2d
        Catch
            Return
        End Try
        If circCol Is Nothing Then Return

        Dim n As Integer = 0
        Try : n = circCol.Count : Catch : Return : End Try
        Dim logged As Integer = 0
        For i As Integer = 1 To n
            Dim c As DVCircle2d = Nothing
            Try
                c = CType(circCol.Item(i), DVCircle2d)
            Catch
                Continue For
            End Try
            If c Is Nothing Then Continue For

            Dim cx As Double, cy As Double, rad As Double
            Dim ok As Boolean = TryCircleCenterRadius(c, cx, cy, rad)
            Dim csx As Double = 0, csy As Double = 0
            If ok Then
                Try
                    dv.ViewToSheet(cx, cy, csx, csy)
                Catch
                    ok = False
                End Try
            End If
            Dim diam As Double = If(ok, 2.0R * rad, 0R)
            Dim bminx As Double = 0, bmaxx As Double = 0, bminy As Double = 0, bmaxy As Double = 0
            If ok AndAlso rad > 1.0E-12 Then
                bminx = Double.MaxValue : bmaxx = Double.MinValue : bminy = Double.MaxValue : bmaxy = Double.MinValue
                For k As Integer = 0 To 7
                    Dim ang As Double = k * Math.PI / 4.0R
                    Dim sx As Double, sy As Double
                    dv.ViewToSheet(cx + rad * Math.Cos(ang), cy + rad * Math.Sin(ang), sx, sy)
                    bminx = Math.Min(bminx, sx) : bmaxx = Math.Max(bmaxx, sx)
                    bminy = Math.Min(bminy, sy) : bmaxy = Math.Max(bmaxy, sy)
                Next
            End If

            Dim ccls As String = If(ok AndAlso diam > 1.0E-12 AndAlso diam <= opt.HoleCandidateDiamM, "circulo_agujero_candidato", "circulo_otro")
            If ccls = "circulo_agujero_candidato" Then
                sum.ClassifyCircHoleCandidate += 1
            Else
                sum.ClassifyCircOther += 1
            End If
            sum.ClassifyCurveClosed += 1

            Dim info As New CircleGeometryInfo With {
                .Index = i,
                .CenterSheetX = csx, .CenterSheetY = csy, .Radius = rad, .Diameter = diam,
                .BboxMinX = bminx, .BboxMaxX = bmaxx, .BboxMinY = bminy, .BboxMaxY = bmaxy,
                .CurveClass = ccls
            }
            sum.Circles.Add(info)
            If ok AndAlso rad > 1.0E-12 Then
                boxes.Add(New LabeledBBox With {
                    .Label = "DVCircle2d#" & i.ToString(CultureInfo.InvariantCulture),
                    .MinX = bminx, .MaxX = bmaxx, .MinY = bminy, .MaxY = bmaxy,
                    .Obj = c,
                    .PickX = csx, .PickY = csy
                })
            End If

            If logged < MaxEntitiesDetailedPerType Then
                logger.Log("[GEOM][CIRCLE] Vista#" & idx.ToString(CultureInfo.InvariantCulture) &
                           " idx=" & i.ToString(CultureInfo.InvariantCulture) &
                           " centroSheet=(" & FormatInv(csx) & "," & FormatInv(csy) & ") R=" & FormatInv(rad) & " D=" & FormatInv(diam) &
                           " xmin=" & FormatInv(bminx) & " xmax=" & FormatInv(bmaxx) & " ymin=" & FormatInv(bminy) & " ymax=" & FormatInv(bmaxy) &
                           " cls=" & ccls)
                logged += 1
            End If
        Next
        If n > MaxEntitiesDetailedPerType Then
            logger.Log("[GEOM][CIRCLE] Vista#" & idx.ToString(CultureInfo.InvariantCulture) & " ... truncado detalle")
        End If
    End Sub

    Private Shared Sub ProcessEllipsesTyped(idx As Integer, dv As DrawingView, sum As ViewGeometrySummary, logger As Logger, boxes As List(Of LabeledBBox))
        Dim ellCol As DVEllipses2d = Nothing
        Try
            ellCol = dv.DVEllipses2d
        Catch
            Return
        End Try
        If ellCol Is Nothing Then Return

        Dim n As Integer = 0
        Try : n = ellCol.Count : Catch : Return : End Try
        Dim logged As Integer = 0
        For i As Integer = 1 To n
            Dim el As DVEllipse2d = Nothing
            Try
                el = CType(ellCol.Item(i), DVEllipse2d)
            Catch
                Continue For
            End Try
            If el Is Nothing Then Continue For

            Dim maj As Double = 0, minr As Double = 0, orient As Double = 0
            Dim cx As Double = 0, cy As Double = 0
            Dim apiNote As String = TryReadEllipseTyped(el, cx, cy, maj, minr, orient)

            Dim csx As Double = 0, csy As Double = 0
            Try
                dv.ViewToSheet(cx, cy, csx, csy)
            Catch
            End Try

            Dim bminx As Double, bmaxx As Double, bminy As Double, bmaxy As Double
            Dim okB As Boolean = TryRangeToSheetBBox(dv, el, bminx, bmaxx, bminy, bmaxy)
            If okB Then
                boxes.Add(New LabeledBBox With {
                    .Label = "DVEllipse2d#" & i.ToString(CultureInfo.InvariantCulture),
                    .MinX = bminx, .MaxX = bmaxx, .MinY = bminy, .MaxY = bmaxy,
                    .Obj = el,
                    .PickX = csx, .PickY = csy
                })
            Else
                bminx = 0 : bmaxx = 0 : bminy = 0 : bmaxy = 0
            End If

            sum.Ellipses.Add(New EllipseGeometryInfo With {
                .Index = i,
                .CenterSheetX = csx, .CenterSheetY = csy,
                .MajorAxis = maj, .MinorAxis = minr, .OrientationRadians = orient,
                .ApiNote = apiNote,
                .BboxMinX = bminx, .BboxMaxX = bmaxx, .BboxMinY = bminy, .BboxMaxY = bmaxy
            })
            sum.ClassifyCurveClosed += 1

            If logged < MaxEntitiesDetailedPerType Then
                logger.Log("[GEOM][ELLIPSE] Vista#" & idx.ToString(CultureInfo.InvariantCulture) &
                           " idx=" & i.ToString(CultureInfo.InvariantCulture) &
                           " centroSheet=(" & FormatInv(csx) & "," & FormatInv(csy) & ")" &
                           " majorR=" & FormatInv(maj) & " minorR=" & FormatInv(minr) & " orientRad=" & FormatInv(orient) &
                           " xmin=" & FormatInv(bminx) & " xmax=" & FormatInv(bmaxx) & " ymin=" & FormatInv(bminy) & " ymax=" & FormatInv(bmaxy) &
                           " api=" & apiNote)
                logged += 1
            End If
        Next
        If n > MaxEntitiesDetailedPerType Then
            logger.Log("[GEOM][ELLIPSE] Vista#" & idx.ToString(CultureInfo.InvariantCulture) & " ... truncado detalle")
        End If
    End Sub

    Private Shared Function TryReadEllipseTyped(e As DVEllipse2d, ByRef cx As Double, ByRef cy As Double, ByRef maj As Double, ByRef minr As Double, ByRef orient As Double) As String
        cx = 0 : cy = 0 : maj = 0 : minr = 0 : orient = 0
        Try
            e.GetCenterPoint(cx, cy)
            maj = CDbl(e.MajorRadius)
            minr = CDbl(e.MinorRadius)
            Try : orient = CDbl(e.RotationAngle) : Catch : orient = 0 : End Try
            Return "GetCenterPoint+MajorRadius+MinorRadius+RotationAngle"
        Catch
            Return "propiedades elipse no legibles vía interop tipado; usar bbox Range"
        End Try
    End Function

    Private Shared Function TryRangeToSheetBBox(dv As DrawingView, geom As Object, ByRef minX As Double, ByRef maxX As Double, ByRef minY As Double, ByRef maxY As Double) As Boolean
        minX = 0 : maxX = 0 : minY = 0 : maxY = 0
        Try
            Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
            If TypeOf geom Is DVEllipse2d Then
                DirectCast(geom, DVEllipse2d).Range(x1, y1, x2, y2)
            ElseIf TypeOf geom Is DVLineString2d Then
                DirectCast(geom, DVLineString2d).Range(x1, y1, x2, y2)
            ElseIf TypeOf geom Is DVBSplineCurve2d Then
                DirectCast(geom, DVBSplineCurve2d).Range(x1, y1, x2, y2)
            Else
                Return False
            End If
            Dim sx1 As Double, sy1 As Double, sx2 As Double, sy2 As Double, sx3 As Double, sy3 As Double, sx4 As Double, sy4 As Double
            dv.ViewToSheet(x1, y1, sx1, sy1)
            dv.ViewToSheet(x2, y1, sx2, sy2)
            dv.ViewToSheet(x1, y2, sx3, sy3)
            dv.ViewToSheet(x2, y2, sx4, sy4)
            minX = Math.Min(Math.Min(sx1, sx2), Math.Min(sx3, sx4))
            maxX = Math.Max(Math.Max(sx1, sx2), Math.Max(sx3, sx4))
            minY = Math.Min(Math.Min(sy1, sy2), Math.Min(sy3, sy4))
            maxY = Math.Max(Math.Max(sy1, sy2), Math.Max(sy3, sy4))
            Return True
        Catch
            Return False
        End Try
    End Function

    Private Shared Function TryGetOrigin(dv As DrawingView, ByRef originSheetX As Double, ByRef originSheetY As Double) As Boolean
        originSheetX = 0 : originSheetY = 0
        Try
            Dim ox As Double = 0, oy As Double = 0
            dv.GetOrigin(ox, oy)
            dv.ViewToSheet(ox, oy, originSheetX, originSheetY)
            Return True
        Catch
            Return False
        End Try
    End Function

    Private Shared Function TryArcCenterRadius(a As DVArc2d, ByRef cx As Double, ByRef cy As Double, ByRef rad As Double) As Boolean
        cx = 0 : cy = 0 : rad = 0
        Try
            Dim obj As Object = a
            Try
                obj.GetCenterPoint(cx, cy)
            Catch
                ' Si no existe GetCenterPoint, dejamos que el fallo de abajo dispare el Catch final.
            End Try

            Try
                rad = CDbl(obj.Radius)
            Catch
                Try
                    Dim diam As Double = CDbl(obj.Diameter)
                    rad = diam / 2.0R
                Catch
                End Try
            End Try

            Return (rad > 0)
        Catch
            Return False
        End Try
    End Function

    Private Shared Function TryGetArcAnglesRadians(a As DVArc2d, ByRef startRad As Double, ByRef sweepRad As Double, ByRef angNote As String) As Boolean
        startRad = 0 : sweepRad = 0
        angNote = ""
        Try
            Dim obj As Object = a
            ' Intento 1: lectura directa si el interop expone StartAngle/SweepAngle.
            Try
                Dim startCandidate As Double = CDbl(obj.StartAngle)
                Dim sweepCandidate As Double = CDbl(obj.SweepAngle)

                ' Heurística: si parece estar en grados (normalmente mucho más grande que 2*PI).
                If Math.Abs(sweepCandidate) > TwoPi * 2.0R Then
                    startRad = startCandidate * Math.PI / 180.0R
                    sweepRad = sweepCandidate * Math.PI / 180.0R
                    angNote = "StartAngle/SweepAngle (grados->rad)"
                Else
                    startRad = startCandidate
                    sweepRad = sweepCandidate
                    angNote = "StartAngle/SweepAngle (rad asumidos)"
                End If
                Return True
            Catch
            End Try

            ' Intento 2: calcular a partir de centro y extremos.
            Dim cx As Double = 0, cy As Double = 0, radTmp As Double = 0
            If Not TryArcCenterRadius(a, cx, cy, radTmp) Then Return False

            Dim vxS As Double = 0, vyS As Double = 0, vxE As Double = 0, vyE As Double = 0
            obj.GetStartPoint(vxS, vyS)
            obj.GetEndPoint(vxE, vyE)

            Dim endRadComputed As Double = Math.Atan2(vyE - cy, vxE - cx)
            startRad = Math.Atan2(vyS - cy, vxS - cx)
            sweepRad = endRadComputed - startRad
            angNote = "atan2(centro+extremos) sweep=end-start (aprox.)"
            Return True
        Catch ex As Exception
            angNote = "angulos no legibles: " & ex.Message
            Return False
        End Try
    End Function

    Private Shared Function TryCircleCenterRadius(c As DVCircle2d, ByRef cx As Double, ByRef cy As Double, ByRef rad As Double) As Boolean
        cx = 0 : cy = 0 : rad = 0
        Try
            Dim obj As Object = c
            obj.GetCenterPoint(cx, cy)
            Try
                rad = CDbl(obj.Radius)
            Catch
                Dim diam As Double = CDbl(obj.Diameter)
                rad = diam / 2.0R
            End Try
            Return (rad > 0)
        Catch
            Return False
        End Try
    End Function

    Private Shared Sub ProcessPointsTyped(idx As Integer, dv As DrawingView, sum As ViewGeometrySummary, logger As Logger, boxes As List(Of LabeledBBox))
        Dim ptsCol As DVPoints2d = Nothing
        Try
            ptsCol = dv.DVPoints2d
        Catch
            Return
        End Try
        If ptsCol Is Nothing Then Return

        Dim n As Integer = 0
        Try : n = ptsCol.Count : Catch : Return : End Try

        Dim logged As Integer = 0
        For i As Integer = 1 To n
            Dim p As DVPoint2d = Nothing
            Try
                p = CType(ptsCol.Item(i), DVPoint2d)
            Catch
                Continue For
            End Try
            If p Is Nothing Then Continue For

            Dim vx As Double = 0, vy As Double = 0
            Dim okV As Boolean = False
            Dim relNote As String = ""
            Try
                Dim obj As Object = p
                obj.GetPoint(vx, vy)
                okV = True
                relNote = "GetPoint"
            Catch ex As Exception
                relNote = "no se pudo leer punto: " & ex.Message
            End Try

            Dim sx As Double = 0, sy As Double = 0
            Dim okS As Boolean = False
            If okV Then
                Try
                    dv.ViewToSheet(vx, vy, sx, sy)
                    okS = True
                Catch
                End Try
            End If

            sum.Points.Add(New PointGeometryInfo With {
                .Index = i,
                .SheetX = sx,
                .SheetY = sy,
                .RelationNote = relNote
            })

            If okS Then
                boxes.Add(New LabeledBBox With {
                    .Label = "DVPoint2d#" & i.ToString(CultureInfo.InvariantCulture),
                    .MinX = sx, .MaxX = sx, .MinY = sy, .MaxY = sy,
                    .Obj = p,
                    .PickX = sx, .PickY = sy
                })
            End If

            If logged < MaxEntitiesDetailedPerType Then
                logger.Log("[GEOM][POINT] Vista#" & idx.ToString(CultureInfo.InvariantCulture) &
                           " idx=" & i.ToString(CultureInfo.InvariantCulture) &
                           " sheet=(" & FormatInv(sx) & "," & FormatInv(sy) & ")" &
                           " note=" & relNote)
                logged += 1
            End If
        Next

        If n > MaxEntitiesDetailedPerType Then
            logger.Log("[GEOM][POINT] Vista#" & idx.ToString(CultureInfo.InvariantCulture) & " ... truncado detalle")
        End If
    End Sub

    Private Shared Sub ProcessLineStringsTyped(idx As Integer, dv As DrawingView, sum As ViewGeometrySummary, logger As Logger, boxes As List(Of LabeledBBox))
        Dim lsCol As DVLineStrings2d = Nothing
        Try
            lsCol = dv.DVLineStrings2d
        Catch
            Return
        End Try
        If lsCol Is Nothing Then Return

        Dim n As Integer = 0
        Try : n = lsCol.Count : Catch : Return : End Try

        Dim logged As Integer = 0
        For i As Integer = 1 To n
            Dim ls As DVLineString2d = Nothing
            Try
                ls = CType(lsCol.Item(i), DVLineString2d)
            Catch
                Continue For
            End Try
            If ls Is Nothing Then Continue For

            Dim obj As Object = ls
            Dim vx1 As Double = 0, vy1 As Double = 0, vx2 As Double = 0, vy2 As Double = 0
            Dim okA As Boolean = False
            Dim okB As Boolean = False
            Dim apiNote As String = ""

            Try
                obj.GetStartPoint(vx1, vy1)
                okA = True
            Catch ex As Exception
                apiNote &= If(apiNote.Length > 0, "; ", "") & "GetStartPoint: " & ex.Message
            End Try
            Try
                obj.GetEndPoint(vx2, vy2)
                okB = True
            Catch ex As Exception
                apiNote &= If(apiNote.Length > 0, "; ", "") & "GetEndPoint: " & ex.Message
            End Try

            Dim sx1 As Double = 0, sy1 As Double = 0, sx2 As Double = 0, sy2 As Double = 0
            If okA Then
                Try : dv.ViewToSheet(vx1, vy1, sx1, sy1) : Catch : sx1 = 0 : sy1 = 0 : End Try
            End If
            If okB Then
                Try : dv.ViewToSheet(vx2, vy2, sx2, sy2) : Catch : sx2 = 0 : sy2 = 0 : End Try
            End If

            Dim nodeCount As Integer = 0
            Dim lengthApprox As Double = 0
            Dim closed As Boolean = False

            Try : nodeCount = CInt(obj.NodeCount) : Catch : End Try
            Try : lengthApprox = CDbl(obj.Length) : Catch : End Try
            Try : closed = CBool(obj.Closed) : Catch : End Try

            Dim bminx As Double = 0, bmaxx As Double = 0, bminy As Double = 0, bmaxy As Double = 0
            Dim okBbox As Boolean = TryRangeToSheetBBox(dv, ls, bminx, bmaxx, bminy, bmaxy)
            If Not okBbox Then
                apiNote &= If(apiNote.Length > 0, "; ", "") & "Range->bbox: no legible"
            End If

            sum.LineStrings.Add(New LineStringGeometryInfo With {
                .Index = i,
                .NodeCount = nodeCount,
                .LengthApprox = lengthApprox,
                .StartSheetX = sx1,
                .StartSheetY = sy1,
                .EndSheetX = sx2,
                .EndSheetY = sy2,
                .Closed = closed,
                .BboxMinX = bminx,
                .BboxMaxX = bmaxx,
                .BboxMinY = bminy,
                .BboxMaxY = bmaxy,
                .ApiNote = apiNote
            })

            If okBbox Then
                boxes.Add(New LabeledBBox With {
                    .Label = "DVLineString2d#" & i.ToString(CultureInfo.InvariantCulture),
                    .MinX = bminx, .MaxX = bmaxx, .MinY = bminy, .MaxY = bmaxy,
                    .Obj = ls,
                    .PickX = (bminx + bmaxx) / 2.0R, .PickY = (bminy + bmaxy) / 2.0R
                })
            End If

            If logged < MaxEntitiesDetailedPerType Then
                logger.Log("[GEOM][LINESTRING] Vista#" & idx.ToString(CultureInfo.InvariantCulture) &
                           " idx=" & i.ToString(CultureInfo.InvariantCulture) &
                           " P0=(" & FormatInv(sx1) & "," & FormatInv(sy1) & ")" &
                           " P1=(" & FormatInv(sx2) & "," & FormatInv(sy2) & ")" &
                           " L≈" & FormatInv(lengthApprox) &
                           " xmin=" & FormatInv(bminx) & " xmax=" & FormatInv(bmaxx) &
                           " ymin=" & FormatInv(bminy) & " ymax=" & FormatInv(bmaxy))
                logged += 1
            End If
        Next

        If n > MaxEntitiesDetailedPerType Then
            logger.Log("[GEOM][LINESTRING] Vista#" & idx.ToString(CultureInfo.InvariantCulture) & " ... truncado detalle")
        End If
    End Sub

    Private Shared Sub ProcessBSplinesTyped(idx As Integer, dv As DrawingView, sum As ViewGeometrySummary, logger As Logger, boxes As List(Of LabeledBBox))
        Dim bsCol As DVBSplineCurves2d = Nothing
        Try
            bsCol = dv.DVBSplineCurves2d
        Catch
            Return
        End Try
        If bsCol Is Nothing Then Return

        Dim n As Integer = 0
        Try : n = bsCol.Count : Catch : Return : End Try

        Dim logged As Integer = 0
        For i As Integer = 1 To n
            Dim bs As DVBSplineCurve2d = Nothing
            Try
                bs = CType(bsCol.Item(i), DVBSplineCurve2d)
            Catch
                Continue For
            End Try
            If bs Is Nothing Then Continue For

            Dim obj As Object = bs
            Dim vx1 As Double = 0, vy1 As Double = 0, vx2 As Double = 0, vy2 As Double = 0
            Dim okA As Boolean = False
            Dim okB As Boolean = False
            Dim apiNote As String = ""

            Try
                obj.GetStartPoint(vx1, vy1)
                okA = True
            Catch ex As Exception
                apiNote &= If(apiNote.Length > 0, "; ", "") & "GetStartPoint: " & ex.Message
            End Try
            Try
                obj.GetEndPoint(vx2, vy2)
                okB = True
            Catch ex As Exception
                apiNote &= If(apiNote.Length > 0, "; ", "") & "GetEndPoint: " & ex.Message
            End Try

            Dim sx1 As Double = 0, sy1 As Double = 0, sx2 As Double = 0, sy2 As Double = 0
            If okA Then
                Try : dv.ViewToSheet(vx1, vy1, sx1, sy1) : Catch : sx1 = 0 : sy1 = 0 : End Try
            End If
            If okB Then
                Try : dv.ViewToSheet(vx2, vy2, sx2, sy2) : Catch : sx2 = 0 : sy2 = 0 : End Try
            End If

            Dim nodeCount As Integer = 0
            Dim lengthApprox As Double = 0
            Dim closed As Boolean = False

            Try : nodeCount = CInt(obj.NodeCount) : Catch : End Try
            Try : lengthApprox = CDbl(obj.Length) : Catch : End Try
            Try : closed = CBool(obj.Closed) : Catch : End Try

            Dim bminx As Double = 0, bmaxx As Double = 0, bminy As Double = 0, bmaxy As Double = 0
            Dim okBbox As Boolean = TryRangeToSheetBBox(dv, bs, bminx, bmaxx, bminy, bmaxy)
            If Not okBbox Then
                apiNote &= If(apiNote.Length > 0, "; ", "") & "Range->bbox: no legible"
            End If

            sum.BSplines.Add(New BSplineGeometryInfo With {
                .Index = i,
                .NodeCount = nodeCount,
                .LengthApprox = lengthApprox,
                .StartSheetX = sx1,
                .StartSheetY = sy1,
                .EndSheetX = sx2,
                .EndSheetY = sy2,
                .Closed = closed,
                .BboxMinX = bminx,
                .BboxMaxX = bmaxx,
                .BboxMinY = bminy,
                .BboxMaxY = bmaxy,
                .ApiNote = apiNote
            })

            If okBbox Then
                boxes.Add(New LabeledBBox With {
                    .Label = "DVBSplineCurve2d#" & i.ToString(CultureInfo.InvariantCulture),
                    .MinX = bminx, .MaxX = bmaxx, .MinY = bminy, .MaxY = bmaxy,
                    .Obj = bs,
                    .PickX = (bminx + bmaxx) / 2.0R, .PickY = (bminy + bmaxy) / 2.0R
                })
            End If

            If logged < MaxEntitiesDetailedPerType Then
                logger.Log("[GEOM][BSPLINE] Vista#" & idx.ToString(CultureInfo.InvariantCulture) &
                           " idx=" & i.ToString(CultureInfo.InvariantCulture) &
                           " P0=(" & FormatInv(sx1) & "," & FormatInv(sy1) & ")" &
                           " P1=(" & FormatInv(sx2) & "," & FormatInv(sy2) & ")" &
                           " L≈" & FormatInv(lengthApprox) &
                           " xmin=" & FormatInv(bminx) & " xmax=" & FormatInv(bmaxx) &
                           " ymin=" & FormatInv(bminy) & " ymax=" & FormatInv(bmaxy))
                logged += 1
            End If
        Next

        If n > MaxEntitiesDetailedPerType Then
            logger.Log("[GEOM][BSPLINE] Vista#" & idx.ToString(CultureInfo.InvariantCulture) & " ... truncado detalle")
        End If
    End Sub

    Private Shared Sub FinalizeExtentsAndClassifyLog(viewIdx As Integer, sum As ViewGeometrySummary, boxes As List(Of LabeledBBox), tolAxis As Double, logger As Logger)
        Dim eps As Double = Math.Max(1.0E-12, tolAxis * 1.0E-6R)
        If boxes Is Nothing OrElse boxes.Count = 0 Then
            sum.EntityUnionHasData = False
            sum.EntityUnionMinX = 0 : sum.EntityUnionMaxX = 0
            sum.EntityUnionMinY = 0 : sum.EntityUnionMaxY = 0
            sum.ContributorMinX = "(sin bbox)"
            sum.ContributorMaxX = "(sin bbox)"
            sum.ContributorMinY = "(sin bbox)"
            sum.ContributorMaxY = "(sin bbox)"
            logger.Log("[GEOM][EXTREME] Vista#" & viewIdx.ToString(CultureInfo.InvariantCulture) & " sin entidades con bbox legible.")
            Return
        End If

        Dim minX As Double = Double.MaxValue
        Dim maxX As Double = Double.MinValue
        Dim minY As Double = Double.MaxValue
        Dim maxY As Double = Double.MinValue

        Dim cMinX As New List(Of String)()
        Dim cMaxX As New List(Of String)()
        Dim cMinY As New List(Of String)()
        Dim cMaxY As New List(Of String)()

        Dim minXCandidates As New List(Of LabeledBBox)()
        Dim maxXCandidates As New List(Of LabeledBBox)()
        Dim minYCandidates As New List(Of LabeledBBox)()
        Dim maxYCandidates As New List(Of LabeledBBox)()

        For Each b In boxes
            If b.MinX < minX - eps Then
                minX = b.MinX : cMinX.Clear() : cMinX.Add(b.Label)
                minXCandidates.Clear() : minXCandidates.Add(b)
            ElseIf Math.Abs(b.MinX - minX) <= eps Then
                If Not cMinX.Contains(b.Label) Then cMinX.Add(b.Label)
                minXCandidates.Add(b)
            End If

            If b.MaxX > maxX + eps Then
                maxX = b.MaxX : cMaxX.Clear() : cMaxX.Add(b.Label)
                maxXCandidates.Clear() : maxXCandidates.Add(b)
            ElseIf Math.Abs(b.MaxX - maxX) <= eps Then
                If Not cMaxX.Contains(b.Label) Then cMaxX.Add(b.Label)
                maxXCandidates.Add(b)
            End If

            If b.MinY < minY - eps Then
                minY = b.MinY : cMinY.Clear() : cMinY.Add(b.Label)
                minYCandidates.Clear() : minYCandidates.Add(b)
            ElseIf Math.Abs(b.MinY - minY) <= eps Then
                If Not cMinY.Contains(b.Label) Then cMinY.Add(b.Label)
                minYCandidates.Add(b)
            End If

            If b.MaxY > maxY + eps Then
                maxY = b.MaxY : cMaxY.Clear() : cMaxY.Add(b.Label)
                maxYCandidates.Clear() : maxYCandidates.Add(b)
            ElseIf Math.Abs(b.MaxY - maxY) <= eps Then
                If Not cMaxY.Contains(b.Label) Then cMaxY.Add(b.Label)
                maxYCandidates.Add(b)
            End If
        Next

        sum.EntityUnionHasData = True
        sum.EntityUnionMinX = minX : sum.EntityUnionMaxX = maxX
        sum.EntityUnionMinY = minY : sum.EntityUnionMaxY = maxY

        sum.ContributorMinX = String.Join(",", cMinX)
        sum.ContributorMaxX = String.Join(",", cMaxX)
        sum.ContributorMinY = String.Join(",", cMinY)
        sum.ContributorMaxY = String.Join(",", cMaxY)

        ' Prioridad segura para acotado: DVLine2d > DVArc2d > DVCircle2d.
        Dim bestMinX As LabeledBBox = PickBestForDimension(minXCandidates)
        Dim bestMaxX As LabeledBBox = PickBestForDimension(maxXCandidates)
        Dim bestMinY As LabeledBBox = PickBestForDimension(minYCandidates)
        Dim bestMaxY As LabeledBBox = PickBestForDimension(maxYCandidates)

        If bestMinX IsNot Nothing Then
            sum.ExtremeMinXObject = bestMinX.Obj
            sum.ExtremeMinXPickX = bestMinX.PickX
            sum.ExtremeMinXPickY = bestMinX.PickY
        Else
            sum.ExtremeMinXObject = Nothing
        End If
        If bestMaxX IsNot Nothing Then
            sum.ExtremeMaxXObject = bestMaxX.Obj
            sum.ExtremeMaxXPickX = bestMaxX.PickX
            sum.ExtremeMaxXPickY = bestMaxX.PickY
        Else
            sum.ExtremeMaxXObject = Nothing
        End If
        If bestMinY IsNot Nothing Then
            sum.ExtremeMinYObject = bestMinY.Obj
            sum.ExtremeMinYPickX = bestMinY.PickX
            sum.ExtremeMinYPickY = bestMinY.PickY
        Else
            sum.ExtremeMinYObject = Nothing
        End If
        If bestMaxY IsNot Nothing Then
            sum.ExtremeMaxYObject = bestMaxY.Obj
            sum.ExtremeMaxYPickX = bestMaxY.PickX
            sum.ExtremeMaxYPickY = bestMaxY.PickY
        Else
            sum.ExtremeMaxYObject = Nothing
        End If

        logger.Log("[GEOM][EXTREME] Vista#" & viewIdx.ToString(CultureInfo.InvariantCulture) &
                   " xmin=" & FormatInv(minX) & " (" & sum.ContributorMinX & ")" &
                   " xmax=" & FormatInv(maxX) & " (" & sum.ContributorMaxX & ")" &
                   " ymin=" & FormatInv(minY) & " (" & sum.ContributorMinY & ")" &
                   " ymax=" & FormatInv(maxY) & " (" & sum.ContributorMaxY & ")")
        logger.Log("[GEOM][EXTREME] Vista#" & viewIdx.ToString(CultureInfo.InvariantCulture) &
                   " selección_acotado: xmin=" & GetObjectTypeName(sum.ExtremeMinXObject) &
                   " xmax=" & GetObjectTypeName(sum.ExtremeMaxXObject) &
                   " ymin=" & GetObjectTypeName(sum.ExtremeMinYObject) &
                   " ymax=" & GetObjectTypeName(sum.ExtremeMaxYObject))
    End Sub

    Private Shared Function PickBestForDimension(candidates As List(Of LabeledBBox)) As LabeledBBox
        If candidates Is Nothing OrElse candidates.Count = 0 Then Return Nothing
        Dim best As LabeledBBox = Nothing
        Dim bestRank As Integer = Integer.MaxValue
        For Each c In candidates
            Dim r As Integer = GetDimensionPriorityRank(c.Obj)
            If r < bestRank Then
                best = c
                bestRank = r
            End If
        Next
        If bestRank = Integer.MaxValue Then Return Nothing
        Return best
    End Function

    Private Shared Function GetDimensionPriorityRank(obj As Object) As Integer
        If TypeOf obj Is DVLine2d Then Return 1
        If TypeOf obj Is DVArc2d Then Return 2
        If TypeOf obj Is DVCircle2d Then Return 3
        Return Integer.MaxValue
    End Function

    Private Shared Function GetObjectTypeName(obj As Object) As String
        If obj Is Nothing Then Return "(sin entidad prioritaria compatible)"
        Try
            Return obj.GetType().Name
        Catch
            Return "(tipo no legible)"
        End Try
    End Function

    Private Shared Function FormatInv(v As Double) As String
        If Double.IsNaN(v) OrElse Double.IsInfinity(v) Then Return "NaN"
        Return v.ToString("0.###############", CultureInfo.InvariantCulture)
    End Function

    Private Shared Sub TryWriteGeomReportFile(logger As Logger, config As JobConfiguration, summaries As List(Of ViewGeometrySummary))
        If config Is Nothing OrElse summaries Is Nothing Then Return
        If String.IsNullOrWhiteSpace(config.OutputFolder) Then Return

        Dim baseName As String = "PIEZA"
        Try
            If Not String.IsNullOrWhiteSpace(config.InputFile) Then
                baseName = System.IO.Path.GetFileNameWithoutExtension(config.InputFile)
            End If
        Catch
        End Try

        Dim outDir As String = System.IO.Path.Combine(config.OutputFolder, "GEOM")
        If Not Directory.Exists(outDir) Then Directory.CreateDirectory(outDir)

        Dim outPath As String = System.IO.Path.Combine(outDir, baseName & "_GEOM.txt")
        Dim sb As New StringBuilder()
        sb.AppendLine("[GEOM][REPORT] " & baseName)
        sb.AppendLine("GeneratedUtc=" & DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture))
        sb.AppendLine("ExperimentalDraftGeometryDiagnostics=" & config.ExperimentalDraftGeometryDiagnostics.ToString())
        sb.AppendLine("EnableAutoDimensioning=" & config.EnableAutoDimensioning.ToString())
        sb.AppendLine("")

        For Each v In summaries
            sb.AppendLine("=== VISTA index=" & v.ViewIndex & " ===")
            sb.AppendLine("Name=" & v.Name)
            sb.AppendLine("DrawingViewType=" & v.DrawingViewType)
            sb.AppendLine("ViewOrientation=" & v.ViewOrientation)
            sb.AppendLine("ScaleFactor=" & FormatInv(v.ScaleFactor))
            sb.AppendLine("Rotation=" & v.RotationNote)
            sb.AppendLine("OriginSheet=" & If(v.HasOrigin, FormatInv(v.OriginSheetX) & "," & FormatInv(v.OriginSheetY), "(n/d)"))
            sb.AppendLine("RangeSheetMin=(" & FormatInv(v.RangeMinX) & "," & FormatInv(v.RangeMinY) & ") Max=(" & FormatInv(v.RangeMaxX) & "," & FormatInv(v.RangeMaxY) & ")")
            sb.AppendLine("ViewWidth=" & FormatInv(v.ViewWidth) & " ViewHeight=" & FormatInv(v.ViewHeight))
            sb.AppendLine("ModelLink=" & If(v.HasModelLink, v.ModelLinkFileName, "(no)"))
            sb.AppendLine("TolAxisM=" & FormatInv(v.TolAxisM))

            sb.AppendLine("")
            sb.AppendLine("Counts: " &
                          "Lines2d=" & v.Counts.DVLines2d &
                          " Arcs2d=" & v.Counts.DVArcs2d &
                          " Circles2d=" & v.Counts.DVCircles2d &
                          " Ellipses2d=" & v.Counts.DVEllipses2d &
                          " Points2d=" & v.Counts.DVPoints2d &
                          " LineStrings2d=" & v.Counts.DVLineStrings2d &
                          " BSplineCurves2d=" & v.Counts.DVBSplineCurves2d)
            sb.AppendLine("Classify: " &
                          "LinesH=" & v.ClassifyLinesH &
                          " LinesV=" & v.ClassifyLinesV &
                          " LinesI=" & v.ClassifyLinesI &
                          " ArcSmall=" & v.ClassifyArcSmall &
                          " ArcLarge=" & v.ClassifyArcLarge &
                          " ArcOther=" & v.ClassifyArcOther &
                          " CircHoleCandidate=" & v.ClassifyCircHoleCandidate &
                          " CurveOpen=" & v.ClassifyCurveOpen &
                          " CurveClosed=" & v.ClassifyCurveClosed)

            sb.AppendLine("")
            sb.AppendLine("ExtremesGlobal (entities real bbox): " &
                          "xmin=" & FormatInv(v.EntityUnionMinX) & " [" & v.ContributorMinX & "]" &
                          " xmax=" & FormatInv(v.EntityUnionMaxX) & " [" & v.ContributorMaxX & "]" &
                          " ymin=" & FormatInv(v.EntityUnionMinY) & " [" & v.ContributorMinY & "]" &
                          " ymax=" & FormatInv(v.EntityUnionMaxY) & " [" & v.ContributorMaxY & "]" &
                          " hasData=" & v.EntityUnionHasData.ToString())

            sb.AppendLine("")
            sb.AppendLine("Lines (DVLine2d):")
            For Each ln In v.Lines
                sb.AppendLine("-#" & ln.Index &
                               " P0=(" & FormatInv(ln.X1) & "," & FormatInv(ln.Y1) & ")" &
                               " P1=(" & FormatInv(ln.X2) & "," & FormatInv(ln.Y2) & ")" &
                               " L=" & FormatInv(ln.Length) &
                               " dx=" & FormatInv(ln.DeltaX) &
                               " dy=" & FormatInv(ln.DeltaY) &
                               " angRad=" & FormatInv(ln.AngleRadians) &
                               " angDeg=" & FormatInv(ln.AngleDegrees) &
                               " cls=" & ln.Orientation &
                               " mid=(" & FormatInv(ln.MidX) & "," & FormatInv(ln.MidY) & ")" &
                               " bbox=(" & FormatInv(ln.BboxMinX) & "," & FormatInv(ln.BboxMinY) & ")-(" & FormatInv(ln.BboxMaxX) & "," & FormatInv(ln.BboxMaxY) & ")")
            Next

            sb.AppendLine("")
            sb.AppendLine("Arcs (DVArc2d):")
            For Each a In v.Arcs
                sb.AppendLine("-#" & a.Index &
                               " center=(" & FormatInv(a.CenterSheetX) & "," & FormatInv(a.CenterSheetY) & ")" &
                               " R=" & FormatInv(a.Radius) &
                               " P0=(" & FormatInv(a.StartSheetX) & "," & FormatInv(a.StartSheetY) & ")" &
                               " P1=(" & FormatInv(a.EndSheetX) & "," & FormatInv(a.EndSheetY) & ")" &
                               " ang0=" & FormatInv(a.StartAngleRaw) &
                               " ang1=" & FormatInv(a.EndAngleRaw) &
                               " sweep=" & FormatInv(a.SweepAngleRaw) &
                               " Larc≈" & FormatInv(a.ArcLengthApprox) &
                               " bbox=(" & FormatInv(a.BboxMinX) & "," & FormatInv(a.BboxMinY) & ")-(" & FormatInv(a.BboxMaxX) & "," & FormatInv(a.BboxMaxY) & ")" &
                               " cls=" & a.CurveClass &
                               " api=" & a.ApiNote)
            Next

            sb.AppendLine("")
            sb.AppendLine("Circles (DVCircle2d):")
            For Each c In v.Circles
                sb.AppendLine("-#" & c.Index &
                               " center=(" & FormatInv(c.CenterSheetX) & "," & FormatInv(c.CenterSheetY) & ")" &
                               " R=" & FormatInv(c.Radius) &
                               " D=" & FormatInv(c.Diameter) &
                               " bbox=(" & FormatInv(c.BboxMinX) & "," & FormatInv(c.BboxMinY) & ")-(" & FormatInv(c.BboxMaxX) & "," & FormatInv(c.BboxMaxY) & ")" &
                               " cls=" & c.CurveClass)
            Next

            sb.AppendLine("")
            sb.AppendLine("Ellipses (DVEllipse2d):")
            For Each el In v.Ellipses
                sb.AppendLine("-#" & el.Index &
                               " center=(" & FormatInv(el.CenterSheetX) & "," & FormatInv(el.CenterSheetY) & ")" &
                               " major=" & FormatInv(el.MajorAxis) &
                               " minor=" & FormatInv(el.MinorAxis) &
                               " orientRad=" & FormatInv(el.OrientationRadians) &
                               " bbox=(" & FormatInv(el.BboxMinX) & "," & FormatInv(el.BboxMinY) & ")-(" & FormatInv(el.BboxMaxX) & "," & FormatInv(el.BboxMaxY) & ")" &
                               " api=" & el.ApiNote)
            Next

            sb.AppendLine("")
            sb.AppendLine("Points (DVPoint2d):")
            For Each p In v.Points
                sb.AppendLine("-#" & p.Index &
                               " sheet=(" & FormatInv(p.SheetX) & "," & FormatInv(p.SheetY) & ")" &
                               " note=" & p.RelationNote)
            Next

            sb.AppendLine("")
            sb.AppendLine("LineStrings (DVLineString2d):")
            For Each ls In v.LineStrings
                sb.AppendLine("-#" & ls.Index &
                               " nodeCount=" & ls.NodeCount &
                               " L≈" & FormatInv(ls.LengthApprox) &
                               " P0=(" & FormatInv(ls.StartSheetX) & "," & FormatInv(ls.StartSheetY) & ")" &
                               " P1=(" & FormatInv(ls.EndSheetX) & "," & FormatInv(ls.EndSheetY) & ")" &
                               " closed=" & ls.Closed.ToString() &
                               " bbox=(" & FormatInv(ls.BboxMinX) & "," & FormatInv(ls.BboxMinY) & ")-(" & FormatInv(ls.BboxMaxX) & "," & FormatInv(ls.BboxMaxY) & ")" &
                               " api=" & ls.ApiNote)
            Next

            sb.AppendLine("")
            sb.AppendLine("BSplines (DVBSplineCurve2d):")
            For Each bs In v.BSplines
                sb.AppendLine("-#" & bs.Index &
                               " nodeCount=" & bs.NodeCount &
                               " L≈" & FormatInv(bs.LengthApprox) &
                               " P0=(" & FormatInv(bs.StartSheetX) & "," & FormatInv(bs.StartSheetY) & ")" &
                               " P1=(" & FormatInv(bs.EndSheetX) & "," & FormatInv(bs.EndSheetY) & ")" &
                               " closed=" & bs.Closed.ToString() &
                               " bbox=(" & FormatInv(bs.BboxMinX) & "," & FormatInv(bs.BboxMinY) & ")-(" & FormatInv(bs.BboxMaxX) & "," & FormatInv(bs.BboxMaxY) & ")" &
                               " api=" & bs.ApiNote)
            Next

            sb.AppendLine("")
        Next

        System.IO.File.WriteAllText(outPath, sb.ToString(), Encoding.UTF8)
        logger.Log("[GEOM][FILE] Archivo de diagnóstico: " & System.IO.Path.GetFileName(outPath))
    End Sub

    ''' <summary>
    ''' Captura geometría completa de UNA vista (sin generar fichero).
    ''' Útil para acotado: proporciona EntityUnion* y Extreme*Object/Pick* en <see cref="ViewGeometrySummary"/>.
    ''' </summary>
    Public Shared Function CaptureViewGeometry(dv As DrawingView, logger As Logger, Optional config As JobConfiguration = Nothing) As ViewGeometrySummary
        If dv Is Nothing OrElse logger Is Nothing Then Return Nothing
        Dim opt = ResolveOptions(config)

        Dim sum As New ViewGeometrySummary()
        FillViewHeader(0, dv, sum, logger)
        Dim tolAxis As Double = Math.Max(sum.ViewWidth * opt.TolAxisWidthFraction, opt.TolAxisMinM)
        sum.TolAxisM = tolAxis

        Dim boxes As New List(Of LabeledBBox)()
        FillCollectionCountsTyped(0, dv, sum, logger)
        ProcessLinesTyped(0, dv, sum, tolAxis, opt, logger, boxes)
        ProcessArcsTyped(0, dv, sum, opt, logger, boxes)
        ProcessCirclesTyped(0, dv, sum, opt, logger, boxes)
        ProcessEllipsesTyped(0, dv, sum, logger, boxes)
        ProcessPointsTyped(0, dv, sum, logger, boxes)
        ProcessLineStringsTyped(0, dv, sum, logger, boxes)
        ProcessBSplinesTyped(0, dv, sum, logger, boxes)
        FinalizeExtentsAndClassifyLog(0, sum, boxes, tolAxis, logger)
        Return sum
    End Function

End Class
