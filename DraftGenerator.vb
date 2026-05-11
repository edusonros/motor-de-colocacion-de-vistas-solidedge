Option Strict Off
Imports System.IO
Imports System.Runtime.InteropServices
Imports SolidEdgeDraft
Imports SolidEdgeFramework
Imports SolidEdgeAssembly
Imports SolidEdgePart
Imports System.Linq
Imports IOPath = System.IO.Path

''' <summary>Motor de generación automática de Drafts usando AddByFold.
''' Orquesta medición, selección de layout e inserción de vistas.</summary>
Public Module DraftGenerator

#Region "Constantes"

    ' Origen fijo de vistas (mm)
    Private Const BASE_ORIGIN_X_MM As Double = 0.04    ' 40 mm
    Private Const BASE_ORIGIN_Y_MM As Double = 0.26    ' 260 mm (210 + 50)
    Private Const GAP_BASE_RIGHT_MM As Double = 0.025  ' 25 mm entre Base y Right
    Private Const GAP_BASE_DOWN_MM As Double = 0.025   ' 25 mm entre Base y Down
    Private Const ISO_ORIGIN_X_MM As Double = 0.32     ' 320 mm
    Private Const ISO_ORIGIN_Y_MM As Double = 0.14     ' 140 mm
    Private Const FLAT_ORIGIN_X_MM As Double = 0.035   ' 35 mm - origen fijo vista Flat
    Private Const FLAT_ORIGIN_Y_MM As Double = 0.10   ' 100 mm
    Private Const ISO_FACTOR As Double = 0.45
    Private Const ISO_MARGIN_MM As Double = 0.01   ' 10 mm espacio vacío antes del marco
    Private Const ISO_SEP_ABOVE_CAJETIN_MM As Double = 0.02   ' 20 mm separación sobre el cajetín

    Private ReadOnly StandardScales As Double() = {
        1.0, 0.5, 0.2, 0.1, 0.05, 0.04, 1.0 / 30.0, 0.025, 0.02, 1.0 / 75.0, 0.01
    }

    ''' <summary>
    ''' <c>True</c>: solo se inserta la vista base en el DFT (sin proyecciones AddByFold, sin ISO ni vista flat).
    ''' <c>False</c>: flujo completo como hasta ahora. Quitar la prueba poniendo <c>False</c>.
    ''' </summary>
    Private Const InsertOnlyBaseViewForTesting As Boolean = False

#End Region

#Region "Logs"

    Private Sub Log(msg As String)
        Console.WriteLine($"[DRAFT] {msg}")
    End Sub

    Private Sub LogEx(context As String, ex As Exception)
        Dim hr As String = ""
        If TypeOf ex Is COMException Then hr = $" HR=0x{DirectCast(ex, COMException).ErrorCode:X8}"
        Console.WriteLine($"[EX] {context}: {ex.GetType().Name}{hr} -> {ex.Message}")
    End Sub

#End Region

#Region "Helpers Solid Edge"

    Private Sub DoIdle(app As SolidEdgeFramework.Application)
        Try : app.DoIdle() : Catch : End Try
    End Sub

    ''' <summary>Refresca la geometría 2D proyectada de la vista. El acotado automático corre una sola vez al final en <see cref="DraftGenerationEngine"/> vía <see cref="DimensioningEngine.ApplyAutoDimensioning"/> para no duplicar cotas.</summary>
    Private Sub SafeUpdateView(dv As DrawingView)
        If dv Is Nothing Then Return
        Try : dv.Update() : Catch : End Try
    End Sub

    Private Function TryGetViewRange(dv As DrawingView, ByRef xmin As Double, ByRef ymin As Double, ByRef xmax As Double, ByRef ymax As Double) As Boolean
        xmin = 0 : ymin = 0 : xmax = 0 : ymax = 0
        Try
            dv.Range(xmin, ymin, xmax, ymax)
            Dim x1 = xmin : Dim y1 = ymin : Dim x2 = xmax : Dim y2 = ymax
            xmin = Math.Min(x1, x2) : ymin = Math.Min(y1, y2)
            xmax = Math.Max(x1, x2) : ymax = Math.Max(y1, y2)
            Return True
        Catch
            Return False
        End Try
    End Function

    Private Sub GetViewSize(dv As DrawingView, ByRef w As Double, ByRef h As Double)
        w = 0 : h = 0
        Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
        If Not TryGetViewRange(dv, x1, y1, x2, y2) Then Return
        w = Math.Abs(x2 - x1) : h = Math.Abs(y2 - y1)
    End Sub

    Private Sub MoveViewTopLeft(app As SolidEdgeFramework.Application, dv As DrawingView, targetLeft As Double, targetTop As Double)
        If dv Is Nothing Then Return
        Dim xmin As Double, ymin As Double, xmax As Double, ymax As Double
        If Not TryGetViewRange(dv, xmin, ymin, xmax, ymax) Then Return
        Dim left As Double = Math.Min(xmin, xmax)
        Dim top As Double = Math.Max(ymin, ymax)
        Dim ox As Double = 0, oy As Double = 0
        dv.GetOrigin(ox, oy)
        dv.SetOrigin(ox + (targetLeft - left), oy + (targetTop - top))
        DoIdle(app)
    End Sub

    Private Function GetViewTopLeft(dv As DrawingView, ByRef outLeft As Double, ByRef outTop As Double) As Boolean
        If dv Is Nothing Then Return False
        Dim xmin As Double, ymin As Double, xmax As Double, ymax As Double
        If Not TryGetViewRange(dv, xmin, ymin, xmax, ymax) Then Return False
        outLeft = Math.Min(xmin, xmax)
        outTop = Math.Max(ymin, ymax)
        Return True
    End Function

    Private Sub LogViewTopLeft(dv As DrawingView, viewName As String)
        If dv Is Nothing Then Return
        Dim l As Double = 0, t As Double = 0
        If GetViewTopLeft(dv, l, t) Then Log($"[FOLD] {viewName} final top-left = ({l * 1000:0},{t * 1000:0})mm")
    End Sub

    Private Sub ForceViewOrientation(dv As DrawingView, ori As Integer)
        If dv Is Nothing Then Return
        Try : dv.SetViewOrientationStandard(CType(ori, ViewOrientationConstants)) : Catch : End Try
    End Sub

#End Region

#Region "Medición"

    ''' <summary>Mide Front, Top, Left, Right, Bottom a escala 1.</summary>
    Public Function MeasureAllViewSizes(app As SolidEdgeFramework.Application, modelPath As String,
                                        templatePath As String, isSheetMetal As Boolean,
                                        Optional isAssembly As Boolean = False) As ViewSizesAt1
        Dim r As New ViewSizesAt1
        If Not File.Exists(modelPath) OrElse Not File.Exists(templatePath) Then Return r

        Dim oriMap As New Dictionary(Of Integer, String) From {
            {CInt(ViewOrientationConstants.igFrontView), "Front"},
            {CInt(ViewOrientationConstants.igTopView), "Top"},
            {CInt(ViewOrientationConstants.igLeftView), "Left"},
            {CInt(ViewOrientationConstants.igRightView), "Right"},
            {CInt(ViewOrientationConstants.igBottomView), "Bottom"}
        }
        Dim toMeasure As Integer() = {
            CInt(ViewOrientationConstants.igFrontView),
            CInt(ViewOrientationConstants.igTopView),
            CInt(ViewOrientationConstants.igLeftView),
            CInt(ViewOrientationConstants.igRightView),
            CInt(ViewOrientationConstants.igBottomView)
        }

        Dim modelFull As String = modelPath
        Try
            modelFull = IOPath.GetFullPath(modelPath)
        Catch
        End Try

        For Each ori In toMeasure
            Dim dft As DraftDocument = Nothing
            Try
                dft = CType(app.Documents.Add("SolidEdge.DraftDocument", templatePath), DraftDocument)
                If dft Is Nothing Then
                    Log($"Measure {oriMap(ori)}: Documents.Add devolvió Nothing.")
                    Continue For
                End If

                DoIdle(app)
                DoIdle(app)

                ' El Draft debe ser el documento activo antes de ModelLinks / DrawingViews (evita E_POINTER 0x80004003).
                Try
                    Dim sed As SolidEdgeFramework.SolidEdgeDocument = TryCast(dft, SolidEdgeFramework.SolidEdgeDocument)
                    If sed IsNot Nothing Then sed.Activate()
                Catch
                End Try
                DoIdle(app)

                Dim sheet As Sheet = Nothing
                Try
                    sheet = dft.ActiveSheet
                Catch exSh As Exception
                    LogEx($"Measure {oriMap(ori)} ActiveSheet", exSh)
                    Continue For
                End Try
                If sheet Is Nothing Then
                    Log($"Measure {oriMap(ori)}: ActiveSheet Nothing.")
                    Continue For
                End If

                Dim dvs As DrawingViews = Nothing
                Try
                    dvs = sheet.DrawingViews
                Catch exDvs As Exception
                    LogEx($"Measure {oriMap(ori)} DrawingViews", exDvs)
                    Continue For
                End Try
                If dvs Is Nothing Then
                    Log($"Measure {oriMap(ori)}: DrawingViews Nothing.")
                    Continue For
                End If

                Dim link As ModelLink = Nothing
                Try
                    link = dft.ModelLinks.Add(modelFull)
                Catch exL As Exception
                    LogEx($"Measure {oriMap(ori)} ModelLinks.Add", exL)
                    Continue For
                End Try
                If link Is Nothing Then
                    Log($"Measure {oriMap(ori)}: ModelLinks.Add devolvió Nothing.")
                    Continue For
                End If

                DoIdle(app)
                Try
                    link.UpdateViews()
                Catch
                End Try
                DoIdle(app)
                DoIdle(app)

                Dim dv As DrawingView = Nothing
                Dim x As Double = 0.15 : Dim y As Double = 0.2
                If isAssembly Then
                    dv = dvs.AddAssemblyView(link, CType(ori, ViewOrientationConstants), 1.0, x, y, AssemblyDrawingViewTypeConstants.seAssemblyDesignedView)
                ElseIf Not isSheetMetal Then
                    dv = dvs.AddPartView(link, ori, 1.0, x, y, PartDrawingViewTypeConstants.sePartDesignedView)
                Else
                    Try
                        dv = dvs.AddSheetMetalView(link, ori, 1.0, x, y, SheetMetalDrawingViewTypeConstants.seSheetMetalDesignedView)
                    Catch exSm As Exception
                        LogEx($"Measure {oriMap(ori)} AddSheetMetalView (primer intento)", exSm)
                        DoIdle(app)
                        DoIdle(app)
                        Try
                            dv = dvs.AddSheetMetalView(link, ori, 1.0, x, y, SheetMetalDrawingViewTypeConstants.seSheetMetalDesignedView)
                        Catch exSm2 As Exception
                            Log($"Measure {oriMap(ori)}: AddSheetMetalView falló de nuevo; prueba AddPartView solo para medición.")
                            Try
                                dv = dvs.AddPartView(link, ori, 1.0, x, y, PartDrawingViewTypeConstants.sePartDesignedView)
                            Catch exPv As Exception
                                LogEx($"Measure {oriMap(ori)} AddPartView (fallback medición)", exPv)
                            End Try
                        End Try
                    End Try
                End If

                If dv Is Nothing Then
                    Log($"Measure {oriMap(ori)}: no se creó DrawingView.")
                    Continue For
                End If

                DoIdle(app)
                SafeUpdateView(dv)
                DoIdle(app)
                Dim w As Double = 0, h As Double = 0
                GetViewSize(dv, w, h)
                r.SetSize(ori, w, h)
                Log($"Measure {oriMap(ori)}: W={w * 1000:0}mm H={h * 1000:0}mm Area={w * h * 1E6:0}mm²")
            Catch ex As Exception
                LogEx($"Measure {oriMap(ori)}", ex)
            Finally
                If dft IsNot Nothing Then Try : dft.Close(False) : Catch : End Try
            End Try
        Next
        Return r
    End Function

#End Region

#Region "Candidatos"

    ''' <summary>Genera 6 candidatos: Front 0°/90°, Top 0°/90°, Left 0°/90°.</summary>
    Public Function GenerateBaseViewCandidates(sizes As ViewSizesAt1) As List(Of BaseViewCandidate)
        Dim list As New List(Of BaseViewCandidate)
        Dim baseConfigs As (Integer, String, Integer, Integer)() = {
            (CInt(ViewOrientationConstants.igFrontView), "Front", CInt(ViewOrientationConstants.igRightView), CInt(ViewOrientationConstants.igTopView)),
            (CInt(ViewOrientationConstants.igTopView), "Top", CInt(ViewOrientationConstants.igRightView), CInt(ViewOrientationConstants.igFrontView)),
            (CInt(ViewOrientationConstants.igLeftView), "Left", CInt(ViewOrientationConstants.igFrontView), CInt(ViewOrientationConstants.igBottomView))
        }
        For Each cfg In baseConfigs
            Dim baseOri As Integer = cfg.Item1
            Dim baseName As String = cfg.Item2
            Dim rightOri As Integer = cfg.Item3
            Dim downOri As Integer = cfg.Item4

            Dim baseW As Double = sizes.GetWidth(baseOri)
            Dim baseH As Double = sizes.GetHeight(baseOri)
            Dim rightW As Double = sizes.GetWidth(rightOri)
            Dim rightH As Double = sizes.GetHeight(rightOri)
            Dim downW As Double = sizes.GetWidth(downOri)
            Dim downH As Double = sizes.GetHeight(downOri)

            For rotIdx As Integer = 0 To 1
                Dim rotated As Boolean = (rotIdx = 1)
                Dim bw, bh, rw, rh, dw, dh As Double
                If rotated Then
                    bw = baseH : bh = baseW : rw = rightH : rh = rightW : dw = downH : dh = downW
                Else
                    bw = baseW : bh = baseH : rw = rightW : rh = rightH : dw = downW : dh = downH
                End If
                Dim cand As New BaseViewCandidate With {
                    .BaseOri = baseOri,
                    .BaseOriName = baseName,
                    .Rotated90 = rotated,
                    .BaseSize = ViewSize.Create(bw, bh),
                    .RightSize = ViewSize.Create(rw, rh),
                    .DownSize = ViewSize.Create(dw, dh)
                }
                list.Add(cand)
                Log($"Candidate {baseName} {(If(rotated, "90°", "0°"))}: base=({bw * 1000:0}x{bh * 1000:0}) right=({rw * 1000:0}x{rh * 1000:0}) down=({dw * 1000:0}x{dh * 1000:0})")
            Next
        Next
        Return list
    End Function

#End Region

#Region "Resolución de layout"

    ''' <summary>Resuelve el mejor layout y devuelve posiciones exactas.
    ''' Prioriza: mayor área de vista base + componente horizontal más larga.</summary>
    Public Function ResolveBestLayout(candidates As List(Of BaseViewCandidate),
                                      templates As String(),
                                      usable As LayoutEngine.UsableArea,
                                      isSheetMetal As Boolean) As ResolvedLayout
        If candidates Is Nothing OrElse candidates.Count = 0 Then Return Nothing
        If templates Is Nothing OrElse templates.Length = 0 Then Return Nothing

        Dim tplInfo As LayoutEngine.TemplateInfo = LayoutEngine.GetTemplateInfo(templates(0))
        Dim baseX As Double = BASE_ORIGIN_X_MM
        Dim baseY As Double = BASE_ORIGIN_Y_MM
        Dim effW As Double = usable.MaxX - baseX
        Dim effH As Double = baseY - usable.MinY
        If effW <= 0 OrElse effH <= 0 Then Return Nothing

        Dim best As ResolvedLayout = Nothing
        Dim bestScore As Double = Double.MinValue

        For Each cand In candidates
            Dim baseW1 As Double = cand.BaseSize.Width
            Dim baseH1 As Double = cand.BaseSize.Height
            Dim rightW1 As Double = cand.RightSize.Width
            Dim rightH1 As Double = cand.RightSize.Height
            Dim downW1 As Double = cand.DownSize.Width
            Dim downH1 As Double = cand.DownSize.Height
            If baseW1 <= 0 OrElse baseH1 <= 0 Then Continue For

            Dim blockW1 As Double = baseW1 + GAP_BASE_RIGHT_MM + rightW1
            Dim blockH1 As Double = baseH1 + GAP_BASE_DOWN_MM + downH1
            Dim scaleMaxW As Double = effW / blockW1
            Dim scaleMaxH As Double = effH / blockH1
            Dim scaleMax As Double = Math.Min(scaleMaxW, scaleMaxH)

            Dim scale As Double = 0.01
            For i As Integer = 0 To StandardScales.Length - 1
                If StandardScales(i) <= scaleMax + 0.000001 Then
                    scale = StandardScales(i)
                    Exit For
                End If
            Next
            If scale <= 0 Then Continue For

            Dim baseW As Double = baseW1 * scale
            Dim baseH As Double = baseH1 * scale
            Dim rightW As Double = rightW1 * scale
            Dim rightH As Double = rightH1 * scale
            Dim downW As Double = downW1 * scale
            Dim downH As Double = downH1 * scale
            Dim blockW As Double = baseW + GAP_BASE_RIGHT_MM + rightW
            Dim blockH As Double = baseH + GAP_BASE_DOWN_MM + downH

            If blockW > effW OrElse blockH > effH Then Continue For

            Dim scaleFactor As Double = 1.0 / scale
            Dim rightX As Double = baseX + baseW + GAP_BASE_RIGHT_MM
            Dim downX As Double = baseX

            Dim areaUtil As Double = (blockW * blockH) / (effW * effH)
            Dim scaleBonus As Double = Math.Log10(scale + 0.001) * 0.5
            ' Priorizar: mayor área de vista base + componente horizontal más larga
            Dim baseAreaAt1 As Double = baseW1 * baseH1
            Dim score As Double = baseAreaAt1 * 1E6 + baseW1 * 500 + areaUtil * 2.0 + scaleBonus

            If score > bestScore Then
                bestScore = score
                Dim isoSize As Double = Math.Max(baseW1, baseH1) * ISO_FACTOR * scale
                ' Base: origen fijo. Right se alinea al top real de base.
                Dim rightYCorr As Double = baseY
                ' Down provisional: top de down desde bottom de base (sin usar downH en Y)
                Dim downYCorr As Double = baseY - baseH - GAP_BASE_DOWN_MM
                best = New ResolvedLayout With {
                    .TemplatePath = templates(0),
                    .Scale = scale,
                    .BaseOri = cand.BaseOri,
                    .BaseOriName = cand.BaseOriName,
                    .Rotated90 = cand.Rotated90,
                    .BaseTopLeftX = baseX,
                    .BaseTopLeftY = baseY,
                    .RightTopLeftX = rightX,
                    .RightTopLeftY = rightYCorr,
                    .DownTopLeftX = downX,
                    .DownTopLeftY = downYCorr,
                    .BaseWidth = baseW,
                    .BaseHeight = baseH,
                    .BlockWidth = blockW,
                    .BlockHeight = blockH,
                    .BaseWidthAt1 = baseW1,
                    .BaseHeightAt1 = baseH1,
                    .RightWidthAt1 = rightW1,
                    .RightHeightAt1 = rightH1,
                    .DownWidthAt1 = downW1,
                    .DownHeightAt1 = downH1,
                    .IsoWidth = isoSize,
                    .IsoHeight = isoSize,
                    .FlatZoneLeft = baseX,
                    .FlatZoneWidth = effW,
                    .IncludeIso = True,
                    .IncludeFlat = isSheetMetal
                }
                ' ISO y Flat: orígenes fijos
                best.IsoTopLeftX = ISO_ORIGIN_X_MM
                best.IsoTopLeftY = ISO_ORIGIN_Y_MM
                best.FlatTopLeftX = FLAT_ORIGIN_X_MM
                best.FlatTopLeftY = FLAT_ORIGIN_Y_MM
            End If
        Next

        If best Is Nothing AndAlso candidates.Count > 0 Then
            Dim cand = candidates(0)
            Dim baseW1 = cand.BaseSize.Width
            Dim baseH1 = cand.BaseSize.Height
            Dim rightW1 = cand.RightSize.Width
            Dim rightH1 = cand.RightSize.Height
            Dim downW1 = cand.DownSize.Width
            Dim downH1 = cand.DownSize.Height
            Dim scale As Double = 0.05
            Dim scaleFactor As Double = 1.0 / scale
            Dim baseW = baseW1 * scale
            Dim baseH = baseH1 * scale
            Dim rightW = rightW1 * scale
            Dim rightH = rightH1 * scale
            Dim downW = downW1 * scale
            Dim downH = downH1 * scale
            ' baseY ya definido al inicio = BASE_ORIGIN_Y_MM
            Dim downYCorr As Double = baseY - baseH - GAP_BASE_DOWN_MM
            Dim rightYCorr As Double = baseY
            Dim rX = baseX + baseW + GAP_BASE_RIGHT_MM
            Dim dX = baseX
            Dim isoSize As Double = Math.Max(baseW1, baseH1) * ISO_FACTOR * scale
            best = New ResolvedLayout With {
                .TemplatePath = templates(0),
                .Scale = scale,
                .BaseOri = cand.BaseOri,
                .BaseOriName = cand.BaseOriName,
                .Rotated90 = cand.Rotated90,
                .BaseTopLeftX = baseX,
                .BaseTopLeftY = baseY,
                .RightTopLeftX = rX,
                .RightTopLeftY = rightYCorr,
                .DownTopLeftX = dX,
                .DownTopLeftY = downYCorr,
                .BaseWidth = baseW,
                .BaseHeight = baseH,
                .BlockWidth = baseW + GAP_BASE_RIGHT_MM + rightW,
                .BlockHeight = baseH + GAP_BASE_DOWN_MM + downH,
                .BaseWidthAt1 = baseW1,
                .BaseHeightAt1 = baseH1,
                .RightWidthAt1 = rightW1,
                .RightHeightAt1 = rightH1,
                .DownWidthAt1 = downW1,
                .DownHeightAt1 = downH1,
                .IsoWidth = isoSize,
                .IsoHeight = isoSize,
                .FlatZoneLeft = baseX,
                .FlatZoneWidth = effW,
                .IncludeIso = True,
                .IncludeFlat = isSheetMetal
            }
            ' Fallback: ISO y Flat con orígenes fijos
            best.IsoTopLeftX = ISO_ORIGIN_X_MM
            best.IsoTopLeftY = ISO_ORIGIN_Y_MM
            best.FlatTopLeftX = FLAT_ORIGIN_X_MM
            best.FlatTopLeftY = FLAT_ORIGIN_Y_MM
        End If
        Return best
    End Function

#End Region

#Region "Inserción"

    ''' <summary>Inserta la vista base e la fija en (leftEdge, topEdge) usando Range real + MoveViewTopLeft.
    ''' NO usa fórmulas con scaleFactor: la posición final depende exclusivamente del target y del Range real de la vista.</summary>
    Private Function InsertBaseView(app As SolidEdgeFramework.Application, sheet As Sheet, modelLink As ModelLink,
                                    isSheetMetal As Boolean, isAssembly As Boolean, layout As ResolvedLayout,
                                    leftEdge As Double, topEdge As Double,
                                    ByRef vBase As DrawingView) As Boolean
        vBase = Nothing
        Try
            Dim dvws As DrawingViews = sheet.DrawingViews
            ' Inserción provisional: posición inicial genérica (Solid Edge necesita un centro para AddPartView)
            Dim cxProv As Double = 0.15 : Dim cyProv As Double = 0.2
            If isAssembly Then
                Try
                    Log($"[ASM] InsertBaseView: AddAssemblyView ori={layout.BaseOri} scale={layout.Scale}")
                    vBase = dvws.AddAssemblyView(modelLink, CType(layout.BaseOri, ViewOrientationConstants), layout.Scale, cxProv, cyProv, AssemblyDrawingViewTypeConstants.seAssemblyDesignedView)
                Catch exAsm As Exception
                    LogEx("InsertBaseView AddAssemblyView (1er intento)", exAsm)
                    DoIdle(app) : DoIdle(app)
                    Try
                        Log($"[ASM] InsertBaseView: reintento AddAssemblyView ori={layout.BaseOri}")
                        vBase = dvws.AddAssemblyView(modelLink, CType(layout.BaseOri, ViewOrientationConstants), layout.Scale, cxProv, cyProv, AssemblyDrawingViewTypeConstants.seAssemblyDesignedView)
                    Catch exAsm2 As Exception
                        LogEx("InsertBaseView AddAssemblyView (2º intento)", exAsm2)
                    End Try
                End Try
            ElseIf Not isSheetMetal Then
                vBase = dvws.AddPartView(modelLink, layout.BaseOri, layout.Scale, cxProv, cyProv, PartDrawingViewTypeConstants.sePartDesignedView)
            Else
                vBase = dvws.AddSheetMetalView(modelLink, layout.BaseOri, layout.Scale, cxProv, cyProv, SheetMetalDrawingViewTypeConstants.seSheetMetalDesignedView)
            End If
            If vBase Is Nothing Then
                Log($"[WARN] InsertBaseView: vista base nula tras la inserción (isAssembly={isAssembly} ori={layout.BaseOri} scale={layout.Scale}).")
                Return False
            End If
            ForceViewOrientation(vBase, layout.BaseOri)
            SafeUpdateView(vBase)
            DoIdle(app)

            If layout.Rotated90 Then
                Try : vBase.SetRotationAngle(-Math.PI / 2.0) : Catch : End Try
                SafeUpdateView(vBase)
                DoIdle(app)
            End If

            ' Leer Range REAL de la vista insertada (antes de mover)
            Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
            If TryGetViewRange(vBase, x1, y1, x2, y2) Then
                Dim leftBef As Double = Math.Min(x1, x2) : Dim topBef As Double = Math.Max(y1, y2)
                Log($"[FOLD] base range before move = left={leftBef * 1000:0} top={topBef * 1000:0}mm")
            End If

            ' Fijar posición final: esquina superior izquierda REAL en (leftEdge, topEdge)
            MoveViewTopLeft(app, vBase, leftEdge, topEdge)
            SafeUpdateView(vBase)

            ' Verificar posición final
            If TryGetViewRange(vBase, x1, y1, x2, y2) Then
                Dim leftAft As Double = Math.Min(x1, x2) : Dim topAft As Double = Math.Max(y1, y2)
                Log($"[FOLD] base range after move = left={leftAft * 1000:0} top={topAft * 1000:0}mm")
                Log($"[FOLD] base final top-left = ({leftAft * 1000:0},{topAft * 1000:0})mm (target was {leftEdge * 1000:0},{topEdge * 1000:0})")
            End If
            Return True
        Catch ex As Exception
            LogEx("InsertBaseView", ex)
            Return False
        End Try
    End Function

    Private Function InsertFoldedView(app As SolidEdgeFramework.Application, sheet As Sheet, baseView As DrawingView,
                                      foldDir As FoldTypeConstants, targetLeft As Double, targetTop As Double,
                                      ByRef vOut As DrawingView) As Boolean
        vOut = Nothing
        Try
            Dim dvws As DrawingViews = sheet.DrawingViews
            vOut = dvws.AddByFold(baseView, foldDir, targetLeft + 0.01, targetTop - 0.01)
            SafeUpdateView(vOut)
            DoIdle(app)
            MoveViewTopLeft(app, vOut, targetLeft, targetTop)
            SafeUpdateView(vOut)
            Return True
        Catch ex As Exception
            LogEx("InsertFoldedView", ex)
            Return False
        End Try
    End Function

    Private Function InsertIsoView(app As SolidEdgeFramework.Application, sheet As Sheet, modelLink As ModelLink,
                                   isSheetMetal As Boolean, isAssembly As Boolean, scale As Double, layout As ResolvedLayout,
                                   ByRef vIso As DrawingView) As Boolean
        vIso = Nothing
        Try
            Dim dvws As DrawingViews = sheet.DrawingViews
            If isAssembly Then
                vIso = dvws.AddAssemblyView(modelLink, ViewOrientationConstants.igTopFrontRightView, scale * ISO_FACTOR, layout.IsoTopLeftX, layout.IsoTopLeftY, AssemblyDrawingViewTypeConstants.seAssemblyDesignedView)
            ElseIf Not isSheetMetal Then
                vIso = dvws.AddPartView(modelLink, CInt(ViewOrientationConstants.igTopFrontRightView), scale * ISO_FACTOR, layout.IsoTopLeftX, layout.IsoTopLeftY, PartDrawingViewTypeConstants.sePartDesignedView)
            Else
                vIso = dvws.AddSheetMetalView(modelLink, CInt(ViewOrientationConstants.igTopFrontRightView), scale * ISO_FACTOR, layout.IsoTopLeftX, layout.IsoTopLeftY, SheetMetalDrawingViewTypeConstants.seSheetMetalDesignedView)
            End If
            SafeUpdateView(vIso)
            DoIdle(app)
            MoveViewTopLeft(app, vIso, layout.IsoTopLeftX, layout.IsoTopLeftY)
            SafeUpdateView(vIso)
            Log($"ISO at ({layout.IsoTopLeftX * 1000:0},{layout.IsoTopLeftY * 1000:0})mm")
            Return True
        Catch ex As Exception
            LogEx("InsertIsoView", ex)
            Return False
        End Try
    End Function

#End Region

#Region "Función principal"

    ''' <summary>Crea un Draft automático desde .par, .psm o .asm.
    ''' Sustituye a CreateDraftAlzadoPrimerDiedro.</summary>
    Public Function CreateAutomaticDraftFromModel(app As SolidEdgeFramework.Application,
                                                  modelPath As String,
                                                  templates As String(),
                                                  cleanTemplatePath As String,
                                                  ByRef flatInserted As Boolean,
                                                  ByRef mainDrawingView As DrawingView,
                                                  Optional enableSlotBBoxViewLayout As Boolean = True,
                                                  Optional viewLayoutLog As Action(Of String) = Nothing) As DraftDocument
        flatInserted = False
        mainDrawingView = Nothing
        If app Is Nothing OrElse String.IsNullOrWhiteSpace(modelPath) OrElse Not File.Exists(modelPath) Then Return Nothing
        If templates Is Nothing OrElse templates.Length = 0 Then Return Nothing
        If String.IsNullOrWhiteSpace(cleanTemplatePath) OrElse Not File.Exists(cleanTemplatePath) Then Return Nothing

        Dim isAssembly As Boolean = modelPath.EndsWith(".asm", StringComparison.OrdinalIgnoreCase)
        Dim isSheetMetal As Boolean = modelPath.EndsWith(".psm", StringComparison.OrdinalIgnoreCase)
        Log($"CreateAutomaticDraft: {IOPath.GetFileName(modelPath)} isAssembly={isAssembly} isSheetMetal={isSheetMetal}")

        Dim usable As LayoutEngine.UsableArea = LayoutEngine.GetUsableAreaForTemplate(templates(0))
        Dim sizes As ViewSizesAt1 = MeasureAllViewSizes(app, modelPath, cleanTemplatePath, isSheetMetal, isAssembly)
        Dim candidates As List(Of BaseViewCandidate) = GenerateBaseViewCandidates(sizes)
        Dim layout As ResolvedLayout = ResolveBestLayout(candidates, templates, usable, isSheetMetal)

        If layout Is Nothing OrElse String.IsNullOrWhiteSpace(layout.TemplatePath) Then
            Log("No layout válido.")
            Return Nothing
        End If

        Log($"Resolved: {IOPath.GetFileName(layout.TemplatePath)} Base={layout.BaseOri} Rot={If(layout.Rotated90, "90°", "0°")} Scale={layout.Scale}")

        ' --- Variables y fórmulas para orígenes de cada vista (ver de dónde sale cada coordenada) ---
        Dim tplInfo As LayoutEngine.TemplateInfo = LayoutEngine.GetTemplateInfo(layout.TemplatePath)
        Dim usableArea As LayoutEngine.UsableArea = LayoutEngine.GetUsableAreaForTemplate(layout.TemplatePath)
        Dim baseX As Double = BASE_ORIGIN_X_MM
        Dim layoutScaleFactor As Double = 1.0 / layout.Scale
        Dim effW As Double = usableArea.MaxX - baseX
        Dim baseY As Double = layout.BaseTopLeftY
        Log($"[ORIGIN VARS] ===== Variables para cálculo de orígenes =====")
        Log($"[ORIGIN VARS] Constantes: BASE=(40,260)mm ISO=(320,140)mm FLAT=(35,100)mm GAP_RIGHT=25mm GAP_DOWN=25mm")
        Log($"[ORIGIN VARS] Área usable: MinX={usableArea.MinX * 1000:0}mm MaxX={usableArea.MaxX * 1000:0}mm MinY={usableArea.MinY * 1000:0}mm MaxY={usableArea.MaxY * 1000:0}mm")
        Log($"[ORIGIN VARS] --- BASE: X=40mm Y=260mm (origen fijo)")
        Log($"[ORIGIN VARS] --- RIGHT: X=40+baseW+25 = {layout.RightTopLeftX * 1000:0}mm | Y=baseY (misma que Base)")
        Log($"[ORIGIN VARS] --- DOWN (provisional): X=40 (misma que Base) | Y=baseTop - baseH - 25 = {layout.DownTopLeftY * 1000:0}mm")
        Log($"[ORIGIN VARS] --- ISO: X=320mm Y=140mm (origen fijo)")
        Log($"[ORIGIN VARS] --- FLAT: X=35mm Y=100mm (origen fijo)")
        Log($"[ORIGIN VARS] ===== Fin variables orígenes =====")

        ' Candidate elegido como vista base
        Log($"Chosen Candidate: {layout.BaseOriName} {(If(layout.Rotated90, "90°", "0°"))}: base=({layout.BaseWidthAt1 * 1000:0}x{layout.BaseHeightAt1 * 1000:0}) right=({layout.RightWidthAt1 * 1000:0}x{layout.RightHeightAt1 * 1000:0}) down=({layout.DownWidthAt1 * 1000:0}x{layout.DownHeightAt1 * 1000:0})")
        Log($"Candidate base view dimensions: W={layout.BaseWidth * 1000:0}mm L={layout.BaseHeight * 1000:0}mm (scaled)")

        ' Template: dimensiones, origen, área útil (tplInfo ya declarado arriba)
        Log($"Template: Size={tplInfo.TemplateWidth:0}x{tplInfo.TemplateHeight:0}mm Origin=({tplInfo.TemplateOriginX * 1000:0},{tplInfo.TemplateOriginY * 1000:0})mm")
        Log($"Template usable area: W={tplInfo.UsableWidth * 1000:0}mm H={tplInfo.UsableHeight * 1000:0}mm Origin=({tplInfo.UsableMinX * 1000:0},{tplInfo.UsableMinY * 1000:0})mm")

        ' Cajetín a respetar
        Log($"Cajetín to respect: Origin=({tplInfo.CajetinOriginX * 1000:0},{tplInfo.CajetinOriginY * 1000:0})mm Size={tplInfo.CajetinWidth * 1000:0}x{tplInfo.CajetinHeight * 1000:0}mm")

        ' Dimensiones vista ISO
        Log($"ISO view: W={layout.IsoWidth * 1000:0}mm H={layout.IsoHeight * 1000:0}mm at ({layout.IsoTopLeftX * 1000:0},{layout.IsoTopLeftY * 1000:0})mm")

        Log($"Pos (provisional): Base=({layout.BaseTopLeftX * 1000:0},{layout.BaseTopLeftY * 1000:0}) Right=({layout.RightTopLeftX * 1000:0},{layout.RightTopLeftY * 1000:0}) Down=({layout.DownTopLeftX * 1000:0},{layout.DownTopLeftY * 1000:0})")

        Dim dft As DraftDocument = Nothing
        Try
            dft = CType(app.Documents.Add("SolidEdge.DraftDocument", layout.TemplatePath), DraftDocument)
            DoIdle(app)
            Dim sheet As Sheet = dft.ActiveSheet
            Dim modelLink As ModelLink = dft.ModelLinks.Add(modelPath)
            DoIdle(app)

            ' [FOLD] Layout por AddByFold: target fijo, sin fórmulas scaleFactor. No se aplica relayout global.
            Dim requestedLeftEdge As Double = layout.BaseTopLeftX
            Dim requestedTopEdge As Double = layout.BaseTopLeftY
            Dim leftEdge As Double = Math.Max(usableArea.MinX, Math.Min(requestedLeftEdge, usableArea.MaxX))
            Dim topEdge As Double = Math.Max(usableArea.MinY, Math.Min(requestedTopEdge, usableArea.MaxY))
            Log($"[FOLD] usable area = MinX={usableArea.MinX * 1000:0} MaxX={usableArea.MaxX * 1000:0} MinY={usableArea.MinY * 1000:0} MaxY={usableArea.MaxY * 1000:0}mm")
            Log($"[FOLD] requested base top-left = ({requestedLeftEdge * 1000:0},{requestedTopEdge * 1000:0})mm")
            Log($"[FOLD] clamped base top-left = ({leftEdge * 1000:0},{topEdge * 1000:0})mm")

            Dim vBase As DrawingView = Nothing
            If Not InsertBaseView(app, sheet, modelLink, isSheetMetal, isAssembly, layout, leftEdge, topEdge, vBase) Then Return Nothing

            Dim vRight As DrawingView = Nothing
            Dim vBelow As DrawingView = Nothing
            Dim vIso As DrawingView = Nothing
            Dim vFlat As DrawingView = Nothing

            ' Releer base ya movida y derivar targets reales para folded views.
            Dim bx1 As Double = 0, by1 As Double = 0, bx2 As Double = 0, by2 As Double = 0
            Dim baseLeft As Double = leftEdge, baseRight As Double = leftEdge + layout.BaseWidth
            Dim baseTop As Double = topEdge, baseBottom As Double = topEdge - layout.BaseHeight
            If TryGetViewRange(vBase, bx1, by1, bx2, by2) Then
                baseLeft = Math.Min(bx1, bx2)
                baseRight = Math.Max(bx1, bx2)
                baseTop = Math.Max(by1, by2)
                baseBottom = Math.Min(by1, by2)
            End If
            Log($"[FOLD] base edges = left={baseLeft * 1000:0} right={baseRight * 1000:0} top={baseTop * 1000:0} bottom={baseBottom * 1000:0}mm")

            If InsertOnlyBaseViewForTesting Then
                Log("[TEST] InsertOnlyBaseViewForTesting=True → solo vista base en el DFT (sin derecha/abajo AddByFold, sin ISO, sin flat). Poner Const=False en DraftGenerator para restaurar el flujo completo.")
            Else
                Dim rightTargetLeft As Double = baseRight + GAP_BASE_RIGHT_MM
                Dim rightTargetTop As Double = baseTop
                Dim downTargetLeft As Double = baseLeft
                Dim downTargetTop As Double = baseBottom - GAP_BASE_DOWN_MM
                Log($"[FOLD] right target from baseRight = ({rightTargetLeft * 1000:0},{rightTargetTop * 1000:0})mm (gap={GAP_BASE_RIGHT_MM * 1000:0}mm)")
                Log($"[FOLD] down target from baseBottom = ({downTargetLeft * 1000:0},{downTargetTop * 1000:0})mm (gap={GAP_BASE_DOWN_MM * 1000:0}mm)")

                If Not InsertFoldedView(app, sheet, vBase, FoldTypeConstants.igFoldRight, rightTargetLeft, rightTargetTop, vRight) Then Return Nothing
                LogViewTopLeft(vRight, "right")

                If Not InsertFoldedView(app, sheet, vBase, FoldTypeConstants.igFoldDown, downTargetLeft, downTargetTop, vBelow) Then Return Nothing
                LogViewTopLeft(vBelow, "down")

                Log($"[FOLD] skipping global relayout for folded main block")
                Log($"[FOLD] apply global layout only to iso/flat (already at fixed positions)")

                If layout.IncludeIso Then
                    InsertIsoView(app, sheet, modelLink, isSheetMetal, isAssembly, layout.Scale, layout, vIso)
                End If

                If isSheetMetal AndAlso layout.IncludeFlat Then
                    If CojonudoBestFit_Bueno.CreateFlatViewForDraft(app, modelPath, sheet.DrawingViews, modelLink, layout.Scale, vFlat) Then
                        flatInserted = True
                        SafeUpdateView(vFlat)
                        DoIdle(app)
                        Dim flatW As Double = 0, flatH As Double = 0
                        GetViewSize(vFlat, flatW, flatH)
                        ' Si la vista flat es más alta que ancha (ej. 58x293), girarla 90° para orientación horizontal
                        If flatH > flatW Then
                            Try : vFlat.SetRotationAngle(-Math.PI / 2.0) : Catch : End Try
                            SafeUpdateView(vFlat)
                            DoIdle(app)
                            GetViewSize(vFlat, flatW, flatH)
                            Log($"Flat rotated 90° (was taller than wide). New size={flatW * 1000:0}x{flatH * 1000:0}mm")
                        End If
                        ' Posición fija (35, 100) mm
                        Dim flatX As Double = layout.FlatTopLeftX
                        Dim flatY As Double = layout.FlatTopLeftY
                        MoveViewTopLeft(app, vFlat, flatX, flatY)
                        SafeUpdateView(vFlat)
                        GetViewSize(vFlat, flatW, flatH)

                        ' Verificar si la Flat cabe en el área usable; si no, escalar hacia abajo
                        Dim usableMinY As Double = usableArea.MinY
                        Dim usableMaxX As Double = usableArea.MaxX
                        Dim flatBottom As Double = flatY - flatH
                        Dim flatRight As Double = flatX + flatW
                        Dim scaleY As Double = 1.0
                        If flatBottom < usableMinY AndAlso flatH > 0.0001 Then
                            scaleY = Math.Max(0.05, (flatY - usableMinY) / flatH)
                            Log($"[FLAT] Overflow vertical: bottom={flatBottom * 1000:0}mm < MinY={usableMinY * 1000:0}mm -> scaleY={scaleY:0.00}")
                        End If
                        Dim scaleX As Double = 1.0
                        If flatRight > usableMaxX AndAlso flatW > 0.0001 Then
                            scaleX = Math.Max(0.05, (usableMaxX - flatX) / flatW)
                            Log($"[FLAT] Overflow horizontal: right={flatRight * 1000:0}mm > MaxX={usableMaxX * 1000:0}mm -> scaleX={scaleX:0.00}")
                        End If
                        Dim scaleDown As Double = Math.Min(Math.Min(scaleY, scaleX), 1.0)
                        If scaleDown < 0.99 Then
                            Dim curSf As Double = layout.Scale
                            Try
                                vFlat.ScaleFactor = curSf * scaleDown
                                SafeUpdateView(vFlat)
                                DoIdle(app)
                                GetViewSize(vFlat, flatW, flatH)
                                MoveViewTopLeft(app, vFlat, flatX, flatY)
                                SafeUpdateView(vFlat)
                                Log($"[FLAT] Scaled down to fit: ScaleFactor {curSf:0.00} -> {curSf * scaleDown:0.00} | new size={flatW * 1000:0}x{flatH * 1000:0}mm")
                            Catch exScale As Exception
                                LogEx("[FLAT] ScaleFactor", exScale)
                            End Try
                        End If

                        Log($"Flat at ({flatX * 1000:0},{flatY * 1000:0})mm size={flatW * 1000:0}x{flatH * 1000:0}mm")
                    End If
                End If
            End If

            Dim vlLog As Action(Of String) = If(viewLayoutLog, Sub(m) Console.WriteLine(m))
            If enableSlotBBoxViewLayout AndAlso Not InsertOnlyBaseViewForTesting AndAlso vBase IsNot Nothing AndAlso vRight IsNot Nothing AndAlso vBelow IsNot Nothing Then
                Dim applied As Boolean = False
                Try
                    Try
                        dft.UpdateAll(True)
                    Catch
                    End Try
                    DoIdle(app)
                    applied = SlotBBoxViewLayout.TryApplySlotBBoxLayout(
                        app, dft, sheet, layout,
                        vBase, vRight, vBelow,
                        If(layout.IncludeIso, vIso, Nothing),
                        vFlat,
                        flatInserted AndAlso layout.IncludeFlat,
                        vlLog)
                Catch ex As Exception
                    applied = False
                    vlLog("[VIEWLAYOUT][FALLBACK_TO_LEGACY] excepción en TryApplySlotBBoxLayout: " & ex.Message)
                    LogEx("SlotBBoxViewLayout", ex)
                End Try
                If Not applied Then
                    vlLog("[VIEWLAYOUT][FALLBACK_TO_LEGACY] Motor slot no aplicado; se conserva layout AddByFold + posiciones ISO/Flat previas")
                    vlLog("[VIEWLAYOUT][SUMMARY] fallback_a_layout_anterior=true (sin motor slot o error previo)")
                End If
            ElseIf enableSlotBBoxViewLayout AndAlso (InsertOnlyBaseViewForTesting OrElse vRight Is Nothing OrElse vBelow Is Nothing) Then
                vlLog("[VIEWLAYOUT][FALLBACK_TO_LEGACY] Omitido slot layout (solo base, sin desplegado completo o vistas principales faltantes)")
            End If

            mainDrawingView = vBase
            Return dft
        Catch ex As Exception
            LogEx("CreateAutomaticDraftFromModel", ex)
            If dft IsNot Nothing Then Try : dft.Close(False) : Catch : End Try
            mainDrawingView = Nothing
            Return Nothing
        End Try
    End Function

#End Region

End Module
