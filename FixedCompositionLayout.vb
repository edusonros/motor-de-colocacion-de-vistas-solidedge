Option Strict Off
Imports System.IO
Imports SolidEdgeDraft
Imports SolidEdgeFramework
Imports SolidEdgePart

''' <summary>Composición de plano DETERMINISTA basada en PDF 005 correcto.
''' Vista principal por mayor área -> girar a horizontal si H>2.5*W -> derecha + debajo desde mapa europeo.
''' ISO encima del cajetín. Flat centrado en hueco restante.
''' SIN heurísticas genéricas: lógica geométrica fija.</summary>
Public Module FixedCompositionLayout

    Private Sub FclLog(msg As String)
        Console.WriteLine(msg)
    End Sub

    ' Constantes de layout según reglas del usuario (PDF 005)
    Private Const GAP_H As Double = 0.012           ' Separación horizontal entre vistas
    Private Const GAP_H_ROTATED As Double = 0.006   ' Menor separación cuando la pieza se gira (más larga)
    Private Const GAP_V As Double = 0.012           ' Separación vertical entre vistas
    Private Const MARGIN_INNER As Double = 0.020     ' Márgen interno (20 mm)
    Private Const ISO_FACTOR As Double = 0.45       ' ISO respecto a escala principal

    ' Rotación: si altura >= 2.5 * anchura -> girar para dejar horizontal
    Private Const HORIZONTAL_ASPECT_RATIO As Double = 2.5

    ' Regla horizontal: 5 partes. main+gap+right ≈ 3/5 del ancho libre. A3: 400mm libre -> 240mm para bloque
    Private Const HORIZONTAL_BLOCK_FRACTION As Double = 3.0 / 5.0

    ' Regla vertical: main+bottom ≈ 150mm del hueco de 277mm; 20mm para flat
    Private Const VERTICAL_MAIN_BLOCK_MM As Double = 150.0
    Private Const VERTICAL_FLAT_ZONE_MM As Double = 20.0
    Private Const VERTICAL_TOP_MARGIN_MM As Double = 20.0

    ' A3 Plantilla SIEMPRE. Dimensiones en metros (420x297 mm)
    Private Const A3_SHEET_W As Double = 0.42
    Private Const A3_SHEET_H As Double = 0.297
    Private Const A3_TOP_MARGIN_MM As Double = 0.04   ' 40mm del borde superior (vista principal más abajo)
    Private Const A3_CAJETIN_LEFT As Double = 0.23    ' Cajetín 190mm: empieza en X=0.23
    Private Const A3_CAJETIN_TOP As Double = 0.035    ' Cajetín 30mm alto: hasta Y=0.035

    Private ReadOnly StandardScales As Double() = {
        1.0, 0.5, 0.2, 0.1, 0.05, 0.04, 1.0 / 30.0, 0.025, 0.02, 1.0 / 75.0, 0.01, 1.0 / 150.0, 1.0 / 200.0
    }

#Region "Enums y clases"

    Public Enum MainViewKind
        Front
        Top
        Right
        Left
        Back
        Bottom
    End Enum

    Public Enum ViewRotation
        Rot0
        RotPlus90
        RotMinus90
    End Enum

    Public Class ViewBox
        Public Property Name As String
        Public Property Ori As Integer
        Public Property Width As Double
        Public Property Height As Double
        Public ReadOnly Property Area As Double
            Get
                Return Width * Height
            End Get
        End Property
    End Class

    Public Class LayoutZones
        Public Property MainTopLeftX As Double
        Public Property MainTopLeftY As Double
        Public Property RightTopLeftX As Double
        Public Property RightTopLeftY As Double
        Public Property BottomTopLeftX As Double
        Public Property BottomTopLeftY As Double
        Public Property IsoTopLeftX As Double
        Public Property IsoTopLeftY As Double
        Public Property FlatTopLeftX As Double
        Public Property FlatTopLeftY As Double
        Public Property GapRight As Double
        Public Property GapBelow As Double
    End Class

    Public Class FinalLayoutPlan
        Public Property TemplatePath As String
        Public Property MainView As MainViewKind
        Public Property MainViewOri As Integer
        Public Property Rotation As CojonudoBestFit_Bueno.ViewRotation
        Public Property RightProjectedOri As Integer
        Public Property BottomProjectedOri As Integer
        Public Property RightProjectedView As String
        Public Property BottomProjectedView As String
        Public Property Scale As Double
        Public Property IncludeIso As Boolean
        Public Property IncludeFlat As Boolean
        Public Property Zones As LayoutZones
        Public Property MainWidthAt1 As Double
        Public Property MainHeightAt1 As Double
        Public Property RightWidthAt1 As Double
        Public Property RightHeightAt1 As Double
        Public Property BottomWidthAt1 As Double
        Public Property BottomHeightAt1 As Double
    End Class

#End Region

#Region "Helpers"

    Private Function MainViewKindToOri(kind As MainViewKind) As Integer
        Select Case kind
            Case MainViewKind.Front : Return CInt(ViewOrientationConstants.igFrontView)
            Case MainViewKind.Top : Return CInt(ViewOrientationConstants.igTopView)
            Case MainViewKind.Right : Return CInt(ViewOrientationConstants.igRightView)
            Case MainViewKind.Left : Return CInt(ViewOrientationConstants.igLeftView)
            Case MainViewKind.Back : Return CInt(ViewOrientationConstants.igBackView)
            Case MainViewKind.Bottom : Return CInt(ViewOrientationConstants.igBottomView)
            Case Else : Return CInt(ViewOrientationConstants.igFrontView)
        End Select
    End Function

    Private Function OriToMainViewKind(ori As Integer) As MainViewKind
        Select Case CType(ori, ViewOrientationConstants)
            Case ViewOrientationConstants.igFrontView : Return MainViewKind.Front
            Case ViewOrientationConstants.igTopView : Return MainViewKind.Top
            Case ViewOrientationConstants.igRightView : Return MainViewKind.Right
            Case ViewOrientationConstants.igLeftView : Return MainViewKind.Left
            Case ViewOrientationConstants.igBackView : Return MainViewKind.Back
            Case ViewOrientationConstants.igBottomView : Return MainViewKind.Bottom
            Case Else : Return MainViewKind.Front
        End Select
    End Function

    Private Function OriToName(ori As Integer) As String
        Select Case CType(ori, ViewOrientationConstants)
            Case ViewOrientationConstants.igFrontView : Return "Front"
            Case ViewOrientationConstants.igTopView : Return "Top"
            Case ViewOrientationConstants.igRightView : Return "Right"
            Case ViewOrientationConstants.igLeftView : Return "Left"
            Case ViewOrientationConstants.igBottomView : Return "Bottom"
            Case ViewOrientationConstants.igBackView : Return "Back"
            Case Else : Return "Front"
        End Select
    End Function

    Private Sub GetViewSizeAt1(sizes As CojonudoBestFit_Bueno.BaseViewSizesAtScale1, ori As Integer,
                              ByRef w As Double, ByRef h As Double)
        w = 0 : h = 0
        Select Case CType(ori, ViewOrientationConstants)
            Case ViewOrientationConstants.igFrontView : w = sizes.W_Front : h = sizes.H_Front
            Case ViewOrientationConstants.igTopView : w = sizes.W_Top : h = sizes.H_Top
            Case ViewOrientationConstants.igRightView : w = sizes.W_Right : h = sizes.H_Right
            Case ViewOrientationConstants.igLeftView : w = sizes.W_Right : h = sizes.H_Right
            Case ViewOrientationConstants.igBottomView : w = sizes.W_Top : h = sizes.H_Top
            Case ViewOrientationConstants.igBackView : w = sizes.W_Front : h = sizes.H_Front
            Case Else : w = sizes.W_Front : h = sizes.H_Front
        End Select
    End Sub

#End Region

#Region "PASO 1 - Elegir vista principal por mayor área"

    ''' <summary>Candidatos Front, Top, Right (Left usa mismas dimensiones que Right).</summary>
    Public Function GetCandidateMainViews(sizes As CojonudoBestFit_Bueno.BaseViewSizesAtScale1) As List(Of ViewBox)
        Dim list As New List(Of ViewBox)
        Dim oris As (Integer, String)() = {
            (CInt(ViewOrientationConstants.igFrontView), "Front"),
            (CInt(ViewOrientationConstants.igTopView), "Top"),
            (CInt(ViewOrientationConstants.igRightView), "Right"),
            (CInt(ViewOrientationConstants.igLeftView), "Left")
        }
        For Each o In oris
            Dim w As Double, h As Double
            GetViewSizeAt1(sizes, o.Item1, w, h)
            list.Add(New ViewBox With {.Name = o.Item2, .Ori = o.Item1, .Width = w, .Height = h})
            FclLog($"[MAINVIEW] {o.Item2} area={w * h:0.0000} (W={w:0.000} H={h:0.000})")
        Next
        Return list
    End Function

    ''' <summary>Elige la vista principal por mayor área útil proyectada.</summary>
    Public Function ChooseMainViewByLargestArea(candidates As List(Of ViewBox)) As MainViewKind
        If candidates Is Nothing OrElse candidates.Count = 0 Then Return MainViewKind.Front
        Dim best = candidates.OrderByDescending(Function(c) c.Area).FirstOrDefault()
        Dim kind = OriToMainViewKind(best.Ori)
        FclLog($"[MAINVIEW] Winner={best.Name} area={best.Area:0.0000}")
        Return kind
    End Function

    ''' <summary>Girar vista principal si H/W > 2.5 para dejarla horizontal. +90 o -90 según convenga.</summary>
    Public Function ChooseRotationToMakeMainViewHorizontal(mainView As MainViewKind, mainBox As ViewBox) As CojonudoBestFit_Bueno.ViewRotation
        If mainBox Is Nothing OrElse mainBox.Width <= 0 Then Return CojonudoBestFit_Bueno.ViewRotation.Rot0
        Dim ratio As Double = mainBox.Height / mainBox.Width
        If ratio > HORIZONTAL_ASPECT_RATIO Then
            ' Vista muy vertical -> girar -90° para que quede horizontal (mejor aprovechamiento)
            FclLog($"[ROTATION] Winner=-90 because main H/W={ratio:0.00} > {HORIZONTAL_ASPECT_RATIO}")
            Return CojonudoBestFit_Bueno.ViewRotation.RotMinus90
        End If
        If mainBox.Width > 0 AndAlso mainBox.Height / mainBox.Width < (1.0 / HORIZONTAL_ASPECT_RATIO) Then
            ' Vista muy ancha -> girar +90° para aprovechar formato
            FclLog($"[ROTATION] Winner=+90 because main very wide (H/W={ratio:0.00})")
            Return CojonudoBestFit_Bueno.ViewRotation.RotPlus90
        End If
        FclLog($"[ROTATION] Winner=0 (main already horizontal)")
        Return CojonudoBestFit_Bueno.ViewRotation.Rot0
    End Function

#End Region

#Region "PASO 2 - Vistas proyectadas desde principal (sistema europeo)"

    ''' <summary>Derecha e inferior desde mapa europeo. Usa GetProjectedViewMap de CojonudoBestFit.
    ''' Right = vista a la derecha de la principal; Down = vista debajo de la principal.</summary>
    Public Function GetProjectedViewsFromMain(mainView As MainViewKind, rotation As CojonudoBestFit_Bueno.ViewRotation) As Tuple(Of Integer, Integer)
        Dim mainOri As Integer = MainViewKindToOri(mainView)
        Dim map = CojonudoBestFit_Bueno.GetProjectedViewMap(mainOri, rotation)
        Dim oriRight As Integer = CojonudoBestFit_Bueno.OrthoViewToSolidEdge(map.Right)
        Dim oriBelow As Integer = CojonudoBestFit_Bueno.OrthoViewToSolidEdge(map.Down)
        FclLog($"[PROJECTED] RightOfMain={OriToName(oriRight)} BelowMain={OriToName(oriBelow)}")
        Return Tuple.Create(oriRight, oriBelow)
    End Function

#End Region

#Region "PASO 3-6 - Plan, escala global, zonas"

    ''' <summary>Construye el plan fijo determinista. Composición: principal + derecha + debajo + iso + flat.</summary>
    Public Function BuildFixedCompositionPlan(app As SolidEdgeFramework.Application,
                                               modelPath As String,
                                               templates As String(),
                                               cleanTemplatePath As String,
                                               isSheetMetal As Boolean,
                                               sizes As CojonudoBestFit_Bueno.BaseViewSizesAtScale1) As FinalLayoutPlan

        Dim candidates = GetCandidateMainViews(sizes)
        Dim mainView As MainViewKind = ChooseMainViewByLargestArea(candidates)
        Dim mainOri As Integer = MainViewKindToOri(mainView)
        Dim mainBox As ViewBox = candidates.FirstOrDefault(Function(c) c.Ori = mainOri)
        If mainBox Is Nothing Then mainBox = New ViewBox With {.Width = 0.1, .Height = 0.1}

        Dim rotation As CojonudoBestFit_Bueno.ViewRotation = ChooseRotationToMakeMainViewHorizontal(mainView, mainBox)
        Dim projected = GetProjectedViewsFromMain(mainView, rotation)
        Dim oriRight As Integer = projected.Item1
        Dim oriBelow As Integer = projected.Item2

        ' Dimensiones a escala 1 tras rotación
        Dim mainW1 As Double = mainBox.Width
        Dim mainH1 As Double = mainBox.Height
        If rotation = CojonudoBestFit_Bueno.ViewRotation.RotMinus90 OrElse rotation = CojonudoBestFit_Bueno.ViewRotation.RotPlus90 Then
            Dim t = mainW1 : mainW1 = mainH1 : mainH1 = t
        End If

        Dim rightW1 As Double, rightH1 As Double
        GetViewSizeAt1(sizes, oriRight, rightW1, rightH1)
        If rotation = CojonudoBestFit_Bueno.ViewRotation.RotMinus90 OrElse rotation = CojonudoBestFit_Bueno.ViewRotation.RotPlus90 Then
            Dim t = rightW1 : rightW1 = rightH1 : rightH1 = t
        End If

        Dim belowW1 As Double, belowH1 As Double
        GetViewSizeAt1(sizes, oriBelow, belowW1, belowH1)
        If rotation = CojonudoBestFit_Bueno.ViewRotation.RotMinus90 OrElse rotation = CojonudoBestFit_Bueno.ViewRotation.RotPlus90 Then
            Dim t = belowW1 : belowW1 = belowH1 : belowH1 = t
        End If

        Dim plan As New FinalLayoutPlan
        plan.MainView = mainView
        plan.MainViewOri = mainOri
        plan.Rotation = rotation
        plan.RightProjectedOri = oriRight
        plan.BottomProjectedOri = oriBelow
        plan.RightProjectedView = OriToName(oriRight)
        plan.BottomProjectedView = OriToName(oriBelow)
        plan.MainWidthAt1 = mainW1
        plan.MainHeightAt1 = mainH1
        plan.RightWidthAt1 = rightW1
        plan.RightHeightAt1 = rightH1
        plan.BottomWidthAt1 = belowW1
        plan.BottomHeightAt1 = belowH1
        plan.IncludeFlat = isSheetMetal

        ' A3 Plantilla únicamente
        Dim chosenTemplate As String = Nothing
        Dim bestScale As Double = 0
        Dim usable As LayoutEngine.UsableArea = Nothing

        For Each tpl In templates
            If String.IsNullOrWhiteSpace(tpl) OrElse Not File.Exists(tpl) Then Continue For
            Dim u = LayoutEngine.GetUsableAreaForTemplate(tpl)
            plan.TemplatePath = tpl
            plan.Scale = ComputeScaleForFixedComposition(plan, u)
            plan.IncludeIso = CanIncludeIso(plan, u)

            If plan.Scale >= 0.02 AndAlso plan.Scale > bestScale Then
                bestScale = plan.Scale
                chosenTemplate = tpl
                usable = u
            End If
        Next

        If String.IsNullOrEmpty(chosenTemplate) AndAlso templates IsNot Nothing AndAlso templates.Length > 0 Then
            chosenTemplate = templates(0)
            usable = LayoutEngine.GetUsableAreaForTemplate(chosenTemplate)
            plan.TemplatePath = chosenTemplate
            plan.Scale = ComputeScaleForFixedComposition(plan, usable)
            plan.IncludeIso = False
        End If

        If usable IsNot Nothing AndAlso TemplateBboxLayout.UseSlotBasedLayout Then
            plan.Scale = TemplateBboxLayout.ComputeUneScaleForThreeViewBlock(plan, usable)
            If plan.IncludeIso Then plan.IncludeIso = TemplateBboxLayout.IsoFitsInSlot(plan, usable)
            plan.Zones = TemplateBboxLayout.ComputeZonesFromBboxSlots(plan, usable)
            FclLog("[BBOX-LAYOUT] Escala UNE por slots + zonas desde bbox REF (sin x2)")
        Else
            plan.Scale = plan.Scale * 2.0  ' Ampliar piezas al doble de tamaño
            plan.Zones = ComputeLayoutZones(plan, usable)
        End If

        FclLog($"[TEMPLATE] Winner={Path.GetFileName(plan.TemplatePath)}")
        FclLog($"[SCALE] Final={plan.Scale}")
        FclLog($"[ISO] Included={plan.IncludeIso}")
        FclLog($"[FLAT] Included={plan.IncludeFlat}")
        FclLog("[LAYOUT] Fixed composition applied successfully")

        Return plan
    End Function

    ''' <summary>Escala global para A3: bloque principal + derecha + debajo.</summary>
    Private Function ComputeScaleForFixedComposition(plan As FinalLayoutPlan, usable As LayoutEngine.UsableArea) As Double
        If usable Is Nothing Then Return 0.05
        ' A3: ancho útil ~0.38 (hasta cajetín), alto útil ~0.22 (40mm arriba, 45mm abajo)
        Dim blockWMax As Double = (A3_CAJETIN_LEFT - 0.02) * HORIZONTAL_BLOCK_FRACTION
        Dim blockHMax As Double = (VERTICAL_MAIN_BLOCK_MM / 1000.0)

        Dim blockW1 As Double = plan.MainWidthAt1 + GAP_H + plan.RightWidthAt1
        Dim blockH1 As Double = plan.MainHeightAt1 + GAP_V + plan.BottomHeightAt1

        If blockW1 <= 0 OrElse blockH1 <= 0 Then Return 0.05

        Dim scaleMax As Double = Math.Min(blockWMax / blockW1, blockHMax / blockH1)

        For i As Integer = 0 To StandardScales.Length - 1
            If StandardScales(i) <= scaleMax + 0.000001 Then Return StandardScales(i)
        Next
        Return StandardScales(StandardScales.Length - 1)
    End Function

    ''' <summary>ISO solo si cabe sin forzar reducción importante de las principales.</summary>
    Private Function CanIncludeIso(plan As FinalLayoutPlan, usable As LayoutEngine.UsableArea) As Boolean
        If usable Is Nothing Then Return False
        Dim isoSize As Double = Math.Max(plan.MainWidthAt1, plan.MainHeightAt1) * ISO_FACTOR * plan.Scale
        Dim blockW As Double = (plan.MainWidthAt1 + GAP_H + plan.RightWidthAt1) * plan.Scale + GAP_H + isoSize
        Dim blockH As Double = (plan.MainHeightAt1 + GAP_V + plan.BottomHeightAt1) * plan.Scale
        If blockH < isoSize Then blockH = isoSize
        Return blockW <= usable.Width AndAlso blockH <= usable.Height
    End Function

    ''' <summary>Zonas A3. Origen Solid Edge: abajo-izq, Y hacia arriba.
    ''' Vista principal: 40mm del borde superior. ISO: encima del cajetín. Flat: dentro del marco.</summary>
    Public Function ComputeLayoutZones(plan As FinalLayoutPlan, usable As LayoutEngine.UsableArea) As LayoutZones
        If usable Is Nothing Then
            Dim u As New LayoutEngine.UsableArea
            u.MinX = 0.02 : u.MinY = 0.045 : u.MaxX = 0.38 : u.MaxY = 0.257
            usable = u
        End If

        Dim eff = LayoutEngine.GetEffectiveArea(usable)
        Dim mw = plan.MainWidthAt1 * plan.Scale
        Dim mh = plan.MainHeightAt1 * plan.Scale
        Dim rw = plan.RightWidthAt1 * plan.Scale
        Dim bh = plan.BottomHeightAt1 * plan.Scale

        Dim isRotated As Boolean = (plan.Rotation = CojonudoBestFit_Bueno.ViewRotation.RotPlus90 OrElse
                                    plan.Rotation = CojonudoBestFit_Bueno.ViewRotation.RotMinus90)
        Dim gapH As Double = If(isRotated, GAP_H_ROTATED, GAP_H)

        Dim blockW As Double = mw + gapH + rw
        Dim leftX As Double
        If isRotated Then
            ' Pieza girada (larga): vista principal dentro del marco, no centrar (evitar que se salga por la izquierda)
            leftX = eff.MinX + 0.028
        Else
            Dim leftMargin As Double = (eff.Width - blockW) / 2.0
            leftX = eff.MinX + Math.Max(0, leftMargin)
        End If

        ' Vista principal 40mm del borde superior (usable.MaxY = 0.257)
        Dim mainTopY As Double = usable.MaxY
        Dim mainBottomY As Double = mainTopY - mh

        Dim zones As New LayoutZones
        zones.MainTopLeftX = leftX
        zones.MainTopLeftY = mainTopY
        zones.RightTopLeftX = leftX + mw + gapH
        zones.RightTopLeftY = mainTopY
        zones.BottomTopLeftX = leftX
        zones.BottomTopLeftY = mainBottomY - GAP_V
        zones.GapRight = mw + gapH
        zones.GapBelow = mh + GAP_V

        ' ISO: encima del cajetín. Para piezas giradas, pegada al margen derecho.
        If plan.IncludeIso Then
            Dim isoSize As Double = Math.Max(plan.MainWidthAt1, plan.MainHeightAt1) * ISO_FACTOR * plan.Scale
            zones.IsoTopLeftY = A3_CAJETIN_TOP + 0.02 + isoSize   ' Encima del cajetín (altura bien)
            If isRotated Then
                zones.IsoTopLeftX = A3_SHEET_W - 0.008 - isoSize   ' Casi al final del formato (8mm del borde derecho)
            Else
                zones.IsoTopLeftX = A3_SHEET_W - 0.015 - isoSize   ' Aproximada al lado derecho también
            End If
        Else
            zones.IsoTopLeftX = 0
            zones.IsoTopLeftY = 0
        End If

        ' Flat: DENTRO del marco.
        Dim flatZoneBottom As Double = usable.MinY + 0.015
        Dim flatZoneTop As Double = mainBottomY - GAP_V - bh - 0.025
        zones.FlatTopLeftY = flatZoneBottom + Math.Max(0.03, (flatZoneTop - flatZoneBottom) * 0.6)
        If isRotated Then
            zones.FlatTopLeftX = eff.MinX + (eff.Width - 0.14) / 2.0 - 0.02   ' Desplazada a la izquierda (un poco)
        Else
            zones.FlatTopLeftX = eff.MinX + (eff.Width - 0.14) / 2.0
        End If

        Return zones
    End Function

#End Region

#Region "Aplicar al Draft"

    ''' <summary>Devuelve (x0, y0, gapR, gapB) para InsertStandard3Internal.
    ''' InsertStandard3 usa (x,y) como origen inferior-izq de la vista. Y0 = base de la vista principal.</summary>
    Public Sub GetInsertParamsForFixedComposition(plan As FinalLayoutPlan, ByRef x0 As Double, ByRef y0 As Double,
                                                  ByRef gapR As Double, ByRef gapB As Double)
        Dim z = plan.Zones
        If z Is Nothing Then Exit Sub
        x0 = z.MainTopLeftX
        ' La vista se inserta por su origen (esquina inf-izq). Top de main = MainTopLeftY, base = Top - altura.
        Dim mh As Double = plan.MainHeightAt1 * plan.Scale
        y0 = z.MainTopLeftY - mh
        gapR = z.GapRight
        gapB = z.GapBelow
    End Sub

    ''' <summary>Posición ISO para composición fija: encima del cajetín, a la derecha.</summary>
    Public Function GetIsoPositionForFixedComposition(plan As FinalLayoutPlan) As LayoutEngine.Point2D
        If plan.Zones Is Nothing OrElse Not plan.IncludeIso Then Return New LayoutEngine.Point2D(0, 0)
        Return New LayoutEngine.Point2D(plan.Zones.IsoTopLeftX, plan.Zones.IsoTopLeftY)
    End Function

    ''' <summary>Posición ISO para layout CLÁSICO (cuando no usamos FixedComposition): pegada al margen derecho, encima del cajetín. Evita que CenterAllViews desplace la ISO.</summary>
    Public Function GetIsoPositionForClassicLayout(sheetW As Double, isoW As Double, isoH As Double) As LayoutEngine.Point2D
        Dim isoX As Double = sheetW - 0.008 - isoW   ' 8mm del borde derecho (misma regla que composición fija)
        Dim isoY As Double = A3_CAJETIN_TOP + 0.02 + isoH   ' Encima del cajetín
        Return New LayoutEngine.Point2D(isoX, isoY)
    End Function

    ''' <summary>Posición Flat para composición fija: centrada en hueco restante.</summary>
    Public Function GetFlatPositionForFixedComposition(plan As FinalLayoutPlan) As LayoutEngine.Point2D
        If plan.Zones Is Nothing Then Return New LayoutEngine.Point2D(0.05, 0.05)
        Return New LayoutEngine.Point2D(plan.Zones.FlatTopLeftX, plan.Zones.FlatTopLeftY)
    End Function

    ''' <summary>Posiciones para reposicionar las 3 vistas tras rotación del bloque. (MainX, MainY, RightX, RightY, BottomX, BottomY)</summary>
    Public Sub GetBlockRepositionParams(plan As FinalLayoutPlan,
                                        ByRef mainX As Double, ByRef mainY As Double,
                                        ByRef rightX As Double, ByRef rightY As Double,
                                        ByRef bottomX As Double, ByRef bottomY As Double)
        If plan.Zones Is Nothing Then Exit Sub
        mainX = plan.Zones.MainTopLeftX
        mainY = plan.Zones.MainTopLeftY
        rightX = plan.Zones.RightTopLeftX
        rightY = plan.Zones.RightTopLeftY
        bottomX = plan.Zones.BottomTopLeftX
        bottomY = plan.Zones.BottomTopLeftY
    End Sub

#End Region

End Module
