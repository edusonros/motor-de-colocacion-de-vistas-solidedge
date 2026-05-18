Option Strict Off
Imports System.IO
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports SolidEdgeDraft
Imports SolidEdgeFramework
Imports SolidEdgePart
Imports System.Linq

Public Module CojonudoBestFit_Bueno

    '=========================
    ' LOG
    '=========================
    Private stepId As Integer = 0

    Private Sub StepLog(msg As String)
        ' Comentado: reemplazado por barra de progreso
        'stepId += 1
        'Console.WriteLine($"[{stepId:000}] {msg}")
    End Sub

    Private Sub LogEx(context As String, ex As Exception)
        Dim hr As String = ""
        If TypeOf ex Is COMException Then
            Dim cex = DirectCast(ex, COMException)
            hr = $" HR=0x{cex.ErrorCode:X8}"
        End If
        Console.WriteLine($"[EX] {context}: {ex.GetType().Name}{hr} -> {ex.Message}")
    End Sub

    '=========================
    ' CONFIG INTERNA DE LAYOUT
    '=========================
    ' OJO:
    ' - Templates NO van aquí. Llegan por parámetro.
    ' - Scales NO van aquí. Llegan por parámetro.

    ' Separaciones
    Private Const GAP_H As Double = 0.01
    Private Const GAP_V As Double = 0.01
    Private Const EPS As Double = 0.000001

    ' ISO respecto a la escala principal
    Private Const ISO_FACTOR As Double = 0.45

    ' Queremos que el layout "aproveche" al menos este % del lado más crítico
    Private Const TARGET_UTIL As Double = 0.1

    ' Margen de seguridad para escalar la FLAT al hueco
    Private Const FLAT_FIT_SAFETY As Double = 0.3

    ' Rotación Right
    Private Const RIGHT_ROT_RAD As Double = 1.5707963267949 ' 90º
    Private Const NEG_90_RAD As Double = -1.5707963267949 ' -90º (para vistas auxiliares cuando base<>Front)

    ' Posiciones semilla para inserción de vistas (InsertStandard4)
    Private Const INSERT_OFFSET_FROM_MARGIN As Double = 0.05
    Private Const INSERT_GAP_BELOW As Double = 0.06
    Private Const INSERT_GAP_RIGHT As Double = 0.12
    Private Const INSERT_GAP_ISO As Double = 0.24
    ' InsertStandard3 (medición)
    Private Const LAYOUT_SEED_RIGHT As Double = 0.15
    Private Const LAYOUT_SEED_BELOW As Double = 0.15
    ' Alzado 3 vistas: gaps iniciales y separación mínima entre bordes (según rangos)
    Private Const ALZADO_GAP_RIGHT As Double = 0.15
    Private Const ALZADO_GAP_BELOW As Double = 0.12
    ' Separación mínima entre el borde de una vista y el borde de la adyacente (según rangos)
    Private Const ALZADO_MIN_EDGE_GAP As Double = 0.03
    ' Si H >= ALZADO_TALL_RATIO * W, rotar el bloque de 3 vistas -90° para aprovechar mejor la hoja
    Private Const ALZADO_TALL_RATIO As Double = 2.5

    Private Structure Margins
        Public Left As Double
        Public Right As Double
        Public Top As Double
        Public Bottom As Double
    End Structure

    ' Tamaños de Front/Top/Right a escala 1 (modelo)
    Public Structure BaseViewSizesAtScale1
        Public W_Front As Double
        Public H_Front As Double
        Public W_Top As Double
        Public H_Top As Double
        Public W_Right As Double
        Public H_Right As Double
    End Structure

    Private Const GAP_MODEL_UNITS As Double = 0.1
    Private Const NUM_GAPS_HORIZONTAL As Integer = 3
    Private Const ASPECT_ROTATE_THRESHOLD As Double = 2.0  ' Si H >= 2*W -> rotar 90º (sistema europeo)

    ' Escalas estándar: 1:1, 1:2, 1:5, 1:10, 1:20, 1:25, 1:30, 1:40, 1:50, 1:75, 1:100
    Private ReadOnly StandardScales As Double() = {
        1.0, 0.5, 0.2, 0.1, 0.05, 0.04, 1.0 / 30.0, 0.025, 0.02, 1.0 / 75.0, 0.01, 1.0 / 150.0, 1.0 / 200.0
    }

    ' Escalas UNE normalizadas preferentes para reducción en Draft (1:n).
    Private ReadOnly UneNormalizedReductionScales As Double() = {
        1.0, 0.5, 0.4, 0.2, 0.1, 0.05, 0.04, 1.0 / 30.0, 0.025, 0.02, 1.0 / 75.0, 0.01, 1.0 / 150.0, 1.0 / 200.0, 0.004, 1.0 / 300.0, 0.0025, 0.002
    }

    ' A3 Conrad: inicio típico de cajetín desde la izquierda (m). Usado para control FLAT.
    Private Const A3_CAJETIN_LEFT_X As Double = 0.23R

    Private Structure TemplateUsableArea
        Public Name As String
        Public TemplatePath As String
        Public UsableW As Double
        Public UsableH As Double
    End Structure

    ' Constante: si vista es >= 2.5 veces más alta que ancha, usar esa como base
    Private Const ASPECT_BASE_THRESHOLD As Double = 2.5

    ' ========== LAYOUT BY FOLD (AddByFold) - Nueva lógica 30mm/20mm ==========
    ' Si True: usa AddByFold para vistas proyectadas, base en 30mm del borde, gaps 20mm.
    ' Si False: mantiene flujo anterior (InsertStandard3 + alineaciones).
    Private Const USE_LAYOUT_BY_FOLD As Boolean = True
    ' Unidades: Solid Edge Draft trabaja en METROS (0.03 = 30mm, 0.02 = 20mm).
    Private Const LAYOUT_BASE_OFFSET_MM As Double = 0.03   ' 30 mm del borde superior e izquierdo útil
    Private Const LAYOUT_GAP_MM As Double = 0.02            ' 20 mm entre vista principal y proyectadas

    '======================================================================
    ' SISTEMA DE VISTAS PROYECTADAS (Primer Diedro / Europeo)
    ' Gestiona correctamente qué vista va en cada posición (Up/Down/Left/Right)
    ' cuando la vista base está rotada 0°, +90° o -90°.
    '======================================================================

    ''' <summary>Vistas ortográficas estándar (coinciden con Solid Edge ViewOrientationConstants).</summary>
    Public Enum OrthoView
        Front = 0
        Top = 1
        Bottom = 2
        Right = 3
        Left = 4
        Back = 5
    End Enum

    ''' <summary>Rotación de la vista base: 0°, +90° (antihorario), -90° (horario).</summary>
    Public Enum ViewRotation
        Rot0 = 0
        RotPlus90 = 1   ' Antihorario: Up->Right, Right->Down, Down->Left, Left->Up
        RotMinus90 = 2  ' Horario: Up->Left, Right->Up, Down->Right, Left->Down
    End Enum

    ''' <summary>Mapa de vistas proyectadas en cada posición respecto a la base (sistema europeo).</summary>
    Public Class ProjectedViewMap
        Public Property Up As OrthoView
        Public Property Down As OrthoView
        Public Property Left As OrthoView
        Public Property Right As OrthoView
        Public Overrides Function ToString() As String
            Return $"Up={Up} Down={Down} Left={Left} Right={Right}"
        End Function
    End Class

    ''' <summary>Convierte OrthoView a constante de orientación de Solid Edge.</summary>
    ''' <remarks>ADAPTAR si los nombres de Solid Edge difieren en tu versión.</remarks>
    Public Function OrthoViewToSolidEdge(ov As OrthoView) As Integer
        Select Case ov
            Case OrthoView.Front : Return CInt(ViewOrientationConstants.igFrontView)
            Case OrthoView.Top : Return CInt(ViewOrientationConstants.igTopView)
            Case OrthoView.Bottom : Return CInt(ViewOrientationConstants.igBottomView)
            Case OrthoView.Right : Return CInt(ViewOrientationConstants.igRightView)
            Case OrthoView.Left : Return CInt(ViewOrientationConstants.igLeftView)
            Case OrthoView.Back : Return CInt(ViewOrientationConstants.igBackView)
            Case Else : Return CInt(ViewOrientationConstants.igFrontView)
        End Select
    End Function

    ''' <summary>Convierte constante Solid Edge a OrthoView.</summary>
    Public Function SolidEdgeToOrthoView(ori As Integer) As OrthoView
        Select Case CType(ori, ViewOrientationConstants)
            Case ViewOrientationConstants.igFrontView : Return OrthoView.Front
            Case ViewOrientationConstants.igTopView : Return OrthoView.Top
            Case ViewOrientationConstants.igBottomView : Return OrthoView.Bottom
            Case ViewOrientationConstants.igRightView : Return OrthoView.Right
            Case ViewOrientationConstants.igLeftView : Return OrthoView.Left
            Case ViewOrientationConstants.igBackView : Return OrthoView.Back
            Case Else : Return OrthoView.Front
        End Select
    End Function

    ''' <summary>Obtiene el mapa base sin rotación para la vista indicada (sistema europeo / primer diedro).</summary>
    Private Function GetBaseMap(baseView As OrthoView) As ProjectedViewMap
        Dim m As New ProjectedViewMap
        Select Case baseView
            Case OrthoView.Front
                m.Up = OrthoView.Bottom : m.Down = OrthoView.Top : m.Left = OrthoView.Right : m.Right = OrthoView.Left
            Case OrthoView.Top
                m.Up = OrthoView.Back : m.Down = OrthoView.Front : m.Left = OrthoView.Right : m.Right = OrthoView.Left
            Case OrthoView.Bottom
                m.Up = OrthoView.Front : m.Down = OrthoView.Back : m.Left = OrthoView.Left : m.Right = OrthoView.Right
            Case OrthoView.Right
                m.Up = OrthoView.Bottom : m.Down = OrthoView.Top : m.Left = OrthoView.Front : m.Right = OrthoView.Back
            Case OrthoView.Left
                m.Up = OrthoView.Bottom : m.Down = OrthoView.Top : m.Left = OrthoView.Back : m.Right = OrthoView.Front
            Case OrthoView.Back
                m.Up = OrthoView.Bottom : m.Down = OrthoView.Top : m.Left = OrthoView.Left : m.Right = OrthoView.Right
            Case Else
                m.Up = OrthoView.Bottom : m.Down = OrthoView.Top : m.Left = OrthoView.Right : m.Right = OrthoView.Left
        End Select
        Return m
    End Function

    ''' <summary>Rota el mapa espacialmente. +90° antihorario: Up->Right, Right->Down, Down->Left, Left->Up.</summary>
    Private Function RotateViewMap(map As ProjectedViewMap, rotation As ViewRotation) As ProjectedViewMap
        If rotation = ViewRotation.Rot0 Then Return map
        Dim r As New ProjectedViewMap
        If rotation = ViewRotation.RotPlus90 Then
            r.Right = map.Up : r.Down = map.Right : r.Left = map.Down : r.Up = map.Left
        Else ' RotMinus90
            r.Right = map.Down : r.Up = map.Right : r.Down = map.Left : r.Left = map.Up
        End If
        Return r
    End Function

    ''' <summary>Devuelve el mapa final de vistas proyectadas según base y rotación.</summary>
    ''' <param name="baseView">Vista base (Front, Top, Bottom, Right, Left, Back).</param>
    ''' <param name="rotation">Rotación de la vista base: 0°, +90°, -90°.</param>
    ''' <returns>Mapa con Up, Down, Left, Right indicando qué vista va en cada posición.</returns>
    Public Function GetProjectedViewMap(baseView As OrthoView, rotation As ViewRotation) As ProjectedViewMap
        Dim baseMap As ProjectedViewMap = GetBaseMap(baseView)
        Return RotateViewMap(baseMap, rotation)
    End Function

    ''' <summary>Versión con base y rotación como Integer (para compatibilidad con código existente).</summary>
    Public Function GetProjectedViewMap(baseOri As Integer, rotation As ViewRotation) As ProjectedViewMap
        Return GetProjectedViewMap(SolidEdgeToOrthoView(baseOri), rotation)
    End Function

    ''' <summary>Rutina de depuración: imprime en consola los mapas para Base=Front y las 3 rotaciones.</summary>
    Public Sub DebugPrintProjectedViewMaps()
        For Each rot In {ViewRotation.Rot0, ViewRotation.RotPlus90, ViewRotation.RotMinus90}
            Dim rotStr As String = If(rot = ViewRotation.Rot0, "0", If(rot = ViewRotation.RotPlus90, "+90", "-90"))
            Dim m As ProjectedViewMap = GetProjectedViewMap(OrthoView.Front, rot)
            Console.WriteLine($"Base=Front, Rot={rotStr} => Up={m.Up}, Down={m.Down}, Left={m.Left}, Right={m.Right}")
        Next
    End Sub

    '======================================================================
    ' LAYOUT BY FOLD: Base por área (Front, Top, Left), AddByFold para proyectadas.
    ' Posición base: 30mm del borde superior e izquierdo útil.
    ' Gaps: 20mm entre vista principal y derecha/inferior.
    '======================================================================
#Region "Layout por AddByFold"

    ''' <summary>Log específico para layout por fold (siempre visible en consola).</summary>
    Private Sub LayoutLog(msg As String)
        Console.WriteLine($"[LAYOUT] {msg}")
    End Sub

    ''' <summary>Resultado de la selección de vista base para LayoutByFold.</summary>
    Private Structure LayoutByFoldBaseResult
        Public BaseOri As Integer
        Public BaseOriName As String
        Public Rotated As Boolean
        Public BaseWidth As Double
        Public BaseHeight As Double
    End Structure

    ''' <summary>Selecciona vista base entre Front, Top, Left por mayor área proyectada.
    ''' Evalúa si conviene girar 90° para mejorar distribución (H>=2.5*W).
    ''' Left usa las dimensiones de Right (symmetric).</summary>
    Private Function SelectBaseViewForLayoutByFold(sizes As BaseViewSizesAtScale1,
                                                   usable As LayoutEngine.UsableArea) As LayoutByFoldBaseResult
        Dim r As New LayoutByFoldBaseResult
        r.BaseOri = CInt(ViewOrientationConstants.igFrontView)
        r.BaseOriName = "Front"
        r.Rotated = False
        r.BaseWidth = sizes.W_Front
        r.BaseHeight = sizes.H_Front

        ' Candidatos: Front, Top, Left (Left = dimensiones de Right)
        Dim candidates As (Integer, String, Double, Double)() = {
            (CInt(ViewOrientationConstants.igFrontView), "Front", sizes.W_Front, sizes.H_Front),
            (CInt(ViewOrientationConstants.igTopView), "Top", sizes.W_Top, sizes.H_Top),
            (CInt(ViewOrientationConstants.igLeftView), "Left", sizes.W_Right, sizes.H_Right)
        }

        Dim bestArea As Double = 0
        Dim bestRotated As Boolean = False
        Dim bestW As Double = 0, bestH As Double = 0

        For Each c In candidates
            Dim w As Double = c.Item3, h As Double = c.Item4
            Dim area As Double = w * h
            LayoutLog($"Candidate={c.Item2} Rot=0 Width={w:0.000000} Height={h:0.000000} Area={area:0.000000}")
            If area > bestArea Then
                bestArea = area
                r.BaseOri = c.Item1
                r.BaseOriName = c.Item2
                r.Rotated = False
                r.BaseWidth = w
                r.BaseHeight = h
                bestW = w : bestH = h : bestRotated = False
            End If
            ' Probar giro 90° si H/W >= 2.5
            If w > 0.000001 AndAlso h >= 2.5 * w Then
                Dim wRot As Double = h, hRot As Double = w
                Dim areaRot As Double = wRot * hRot
                LayoutLog($"Candidate={c.Item2} Rot=90 Width={wRot:0.000000} Height={hRot:0.000000} Area={areaRot:0.000000}")
                If areaRot > bestArea Then
                    bestArea = areaRot
                    r.BaseOri = c.Item1
                    r.BaseOriName = c.Item2
                    r.Rotated = True
                    r.BaseWidth = wRot
                    r.BaseHeight = hRot
                    bestW = wRot : bestH = hRot : bestRotated = True
                End If
            End If
        Next

        LayoutLog($"Selected={r.BaseOriName} Rot={If(r.Rotated, "90", "0")} FinalWidth={r.BaseWidth:0.000000} FinalHeight={r.BaseHeight:0.000000}")
        Return r
    End Function

    ''' <summary>Obtiene oriRight y oriBelow que corresponden a igFoldRight e igFoldDown desde la base.
    ''' Proyección europea (primer diedro): Right=derecha, Down=debajo.</summary>
    Private Sub GetOrisForBaseFoldEuropean(baseOri As Integer, ByRef oriRight As Integer, ByRef oriBelow As Integer)
        Select Case CType(baseOri, ViewOrientationConstants)
            Case ViewOrientationConstants.igFrontView
                oriRight = CInt(ViewOrientationConstants.igRightView)
                oriBelow = CInt(ViewOrientationConstants.igTopView)
            Case ViewOrientationConstants.igTopView
                oriRight = CInt(ViewOrientationConstants.igRightView)
                oriBelow = CInt(ViewOrientationConstants.igFrontView)
            Case ViewOrientationConstants.igLeftView
                oriRight = CInt(ViewOrientationConstants.igFrontView)
                oriBelow = CInt(ViewOrientationConstants.igBottomView)
            Case Else
                oriRight = CInt(ViewOrientationConstants.igRightView)
                oriBelow = CInt(ViewOrientationConstants.igTopView)
        End Select
    End Sub

    ''' <summary>Inserta las 3 vistas usando AddByFold: base en (30mm,30mm), derecha e inferior con AddByFold, gaps 20mm.</summary>
    Private Function InsertThreeViewsByFold(app As SolidEdgeFramework.Application,
                                           sheet As Sheet,
                                           modelLink As ModelLink,
                                           isSheetMetal As Boolean,
                                           scale As Double,
                                           baseOri As Integer,
                                           baseRotated As Boolean,
                                           baseWidthAt1 As Double,
                                           baseHeightAt1 As Double,
                                           chosenTemplate As String,
                                           ByRef vBase As DrawingView,
                                           ByRef vRight As DrawingView,
                                           ByRef vBelow As DrawingView,
                                           ByRef oriRight As Integer,
                                           ByRef oriBelow As Integer) As Boolean
        vBase = Nothing : vRight = Nothing : vBelow = Nothing
        Dim usable As LayoutEngine.UsableArea = LayoutEngine.GetUsableAreaForTemplate(chosenTemplate)
        Dim leftEdge As Double = usable.MinX + LAYOUT_BASE_OFFSET_MM
        Dim topEdge As Double = usable.MaxY - LAYOUT_BASE_OFFSET_MM
        LayoutLog($"Base TopLeft target: left={leftEdge * 1000:0}mm top={topEdge * 1000:0}mm (usable MaxY={usable.MaxY:0.000})")

        ' 1) Insertar vista base. Centro aproximado para AddPartView (usa origen); MoveViewTopLeft corregirá a top-left exacto.
        Dim dvws As DrawingViews = sheet.DrawingViews
        Dim baseW1 As Double = If(baseWidthAt1 > 0, baseWidthAt1, 0.1)
        Dim baseH1 As Double = If(baseHeightAt1 > 0, baseHeightAt1, 0.1)
        Dim cx As Double = leftEdge + baseW1 * scale / 2.0
        Dim cy As Double = topEdge - baseH1 * scale / 2.0
        Try
            If Not isSheetMetal Then
                vBase = dvws.AddPartView(modelLink, baseOri, scale, cx, cy, PartDrawingViewTypeConstants.sePartDesignedView)
            Else
                vBase = dvws.AddSheetMetalView(modelLink, baseOri, scale, cx, cy, SheetMetalDrawingViewTypeConstants.seSheetMetalDesignedView)
            End If
        Catch ex As Exception
            LogEx("InsertThreeViewsByFold.Base", ex)
            Return False
        End Try
        ForceViewOrientationStandard(vBase, baseOri, "Base(Fold)")
        SafeUpdateView(vBase, "Base(Fold)")
        DoIdleSafe(app, "LayoutByFold Base")

        Dim baseW As Double = 0, baseH As Double = 0
        GetViewSizeSmart(vBase, baseW, baseH)
        LayoutLog($"Base TopLeft=({leftEdge * 1000:0}mm,{topEdge * 1000:0}mm) size=({baseW * 1000:0}mm x {baseH * 1000:0}mm)")
        MoveViewTopLeft(app, vBase, leftEdge, topEdge, "Base(LayoutByFold)")
        SafeUpdateView(vBase, "Base(pos)")
        DoIdleSafe(app, "LayoutByFold Base move")

        ' Obtener rango real de la base tras posicionar
        Dim baseXmin As Double, baseYmin As Double, baseXmax As Double, baseYmax As Double
        If Not TryGetViewRange(vBase, baseXmin, baseYmin, baseXmax, baseYmax) Then Return False
        baseW = baseXmax - baseXmin
        baseH = baseYmax - baseYmin
        Dim baseRight As Double = baseXmax
        Dim baseBottom As Double = baseYmin
        Dim baseCX As Double = (baseXmin + baseXmax) / 2.0
        Dim baseCY As Double = (baseYmin + baseYmax) / 2.0
        LayoutLog($"Base final: left={baseXmin * 1000:0} right={baseRight * 1000:0} top={baseYmax * 1000:0} bottom={baseBottom * 1000:0} Center=({baseCX:0.000},{baseCY:0.000})")

        ' 2) Vista derecha con AddByFold. Borde izquierdo = baseRight + 20mm
        Dim rightTargetLeft As Double = baseRight + LAYOUT_GAP_MM
        Dim rightX As Double = rightTargetLeft + 0.02
        Dim rightY As Double = baseCY
        Try
            vRight = dvws.AddByFold(vBase, FoldTypeConstants.igFoldRight, rightX, rightY)
        Catch ex As Exception
            LogEx("InsertThreeViewsByFold.AddByFold(Right)", ex)
            Return False
        End Try
        SafeUpdateView(vRight, "Right(Fold)")
        DoIdleSafe(app, "LayoutByFold Right")
        Dim rightW As Double = 0, rightH As Double = 0
        GetViewSizeSmart(vRight, rightW, rightH)
        Dim rightTargetTop As Double = baseCY + rightH / 2.0
        MoveViewTopLeft(app, vRight, rightTargetLeft, rightTargetTop, "Right(LayoutByFold)")
        SafeUpdateView(vRight, "Right(pos)")
        LayoutLog($"RightView gap=20mm targetLeft={rightTargetLeft * 1000:0}mm")

        ' 3) Vista inferior con AddByFold. Borde superior = baseBottom - 20mm
        Dim belowTargetTop As Double = baseBottom - LAYOUT_GAP_MM
        Dim belowX As Double = baseCX
        Dim belowY As Double = belowTargetTop - 0.02
        Try
            vBelow = dvws.AddByFold(vBase, FoldTypeConstants.igFoldDown, belowX, belowY)
        Catch ex As Exception
            LogEx("InsertThreeViewsByFold.AddByFold(Down)", ex)
            Return False
        End Try
        SafeUpdateView(vBelow, "Below(Fold)")
        DoIdleSafe(app, "LayoutByFold Below")
        Dim belowW As Double = 0, belowH As Double = 0
        GetViewSizeSmart(vBelow, belowW, belowH)
        Dim belowTargetLeft As Double = baseCX - belowW / 2.0
        MoveViewTopLeft(app, vBelow, belowTargetLeft, belowTargetTop, "Below(LayoutByFold)")
        SafeUpdateView(vBelow, "Below(pos)")
        LayoutLog($"BottomView gap=20mm targetTop={belowTargetTop * 1000:0}mm")

        GetOrisForBaseFoldEuropean(baseOri, oriRight, oriBelow)
        Return True
    End Function

#End Region

    '======================================================================
    ' ALZADO PRIMER DIEDRO: delegado a DraftGenerator.CreateAutomaticDraftFromModel.
    ' Motor modular con AddByFold, selección de base 6 candidatos, layout por bloque.
    '======================================================================
    Public Function CreateDraftAlzadoPrimerDiedro(app As SolidEdgeFramework.Application,
                                                  modelPath As String,
                                                  templates As String(),
                                                  cleanTemplatePath As String,
                                                  ByRef flatInserted As Boolean,
                                                  ByRef mainDrawingView As SolidEdgeDraft.DrawingView,
                                                  Optional enableSlotBBoxViewLayout As Boolean = True,
                                                  Optional viewLayoutLog As Action(Of String) = Nothing) As SolidEdgeDraft.DraftDocument
        Return DraftGenerator.CreateAutomaticDraftFromModel(app, modelPath, templates, cleanTemplatePath, flatInserted, mainDrawingView, enableSlotBBoxViewLayout, viewLayoutLog)
    End Function

    '======================================================================
    ' [LEGACY] Código anterior de CreateDraftAlzadoPrimerDiedro - mantiene lógica antigua
    ' para referencia o fallback. No se usa si DraftGenerator funciona correctamente.
    '======================================================================
    Private Function CreateDraftAlzadoPrimerDiedro_Legacy(app As SolidEdgeFramework.Application,
                                                  modelPath As String,
                                                  templates As String(),
                                                  cleanTemplatePath As String,
                                                  ByRef flatInserted As Boolean) As SolidEdgeDraft.DraftDocument
        flatInserted = False
        If app Is Nothing OrElse String.IsNullOrWhiteSpace(modelPath) OrElse Not File.Exists(modelPath) Then Return Nothing
        If templates Is Nothing OrElse templates.Length = 0 Then Return Nothing
        If String.IsNullOrWhiteSpace(cleanTemplatePath) OrElse Not File.Exists(cleanTemplatePath) Then Return Nothing

        Dim isSheetMetal As Boolean = modelPath.EndsWith(".psm", StringComparison.OrdinalIgnoreCase)
        StepLog($"[ALZADO-PD] modelPath={modelPath} isSheetMetal={isSheetMetal} templates={templates.Length}")
        If isSheetMetal Then EnsureFlatPatternReady(app, modelPath)

        ' Obtener tamaños de Front/Top/Right a escala 1 para elegir base (Right = dims de Left)
        Dim baseOri As Integer = CInt(ViewOrientationConstants.igFrontView)
        Dim sizesOpt = GetBaseViewSizesAtScale1(app, modelPath, cleanTemplatePath, isSheetMetal)
        Dim useLayoutByFold As Boolean = USE_LAYOUT_BY_FOLD AndAlso sizesOpt.HasValue
        Dim layoutByFoldBase As LayoutByFoldBaseResult = Nothing

        If useLayoutByFold Then
            ' NUEVA LÓGICA: candidatos Front, Top, Left; AddByFold para proyectadas; 30mm/20mm
            Dim usableForLayout = LayoutEngine.GetUsableAreaForTemplate(If(templates.Length > 0, templates(0), ""))
            layoutByFoldBase = SelectBaseViewForLayoutByFold(sizesOpt.Value, usableForLayout)
            baseOri = layoutByFoldBase.BaseOri
            StepLog($"[ALZADO-PD] LayoutByFold: Base={layoutByFoldBase.BaseOriName} Rot={If(layoutByFoldBase.Rotated, "90", "0")}")
        ElseIf sizesOpt.HasValue Then
            baseOri = SelectBaseViewByAreaOnly(sizesOpt.Value)
            StepLog($"[ALZADO-PD] baseOri={OriToConstantName(baseOri)} (mayor área)")
        Else
            StepLog("[ALZADO-PD] No se pudieron leer tamaños -> base=Front por defecto")
        End If
        DoIdleSafe(app, "ALZADO-PD post-sizes")

        Dim oriRight As Integer, oriBelow As Integer
        Dim blockRotation As ViewRotation = ViewRotation.Rot0
        Dim chosenTemplate As String = Nothing
        Dim chosenScale As Double = 0
        Dim x0Insert As Double = 0, y0Insert As Double = 0, gapR As Double = ALZADO_GAP_RIGHT, gapB As Double = ALZADO_GAP_BELOW
        Dim useFixedComposition As Boolean = False
        Dim fclPlan As FixedCompositionLayout.FinalLayoutPlan = Nothing

        ' FASE FIXED COMPOSITION (solo si NO usamos LayoutByFold): composición determinista PDF 005
        If sizesOpt.HasValue AndAlso Not useLayoutByFold Then
            fclPlan = FixedCompositionLayout.BuildFixedCompositionPlan(app, modelPath, templates, cleanTemplatePath, isSheetMetal, sizesOpt.Value)
            If fclPlan IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(fclPlan.TemplatePath) AndAlso fclPlan.Scale >= 0.02 Then
                useFixedComposition = True
                chosenTemplate = fclPlan.TemplatePath
                chosenScale = fclPlan.Scale
                baseOri = fclPlan.MainViewOri
                blockRotation = fclPlan.Rotation
                oriRight = fclPlan.RightProjectedOri
                oriBelow = fclPlan.BottomProjectedOri
                FixedCompositionLayout.GetInsertParamsForFixedComposition(fclPlan, x0Insert, y0Insert, gapR, gapB)
                StepLog($"[ALZADO-PD] FIXED COMPOSITION: {Path.GetFileName(chosenTemplate)} Main={OriToConstantName(baseOri)} Rot={blockRotation} Scale={chosenScale}")
            End If
        End If

        ' Fallback: lógica clásica (rango + PickFormatAndScale)
        If Not useFixedComposition Then
            If useLayoutByFold Then
                GetOrisForBaseFoldEuropean(baseOri, oriRight, oriBelow)
                If layoutByFoldBase.Rotated Then blockRotation = ViewRotation.RotMinus90
            Else
                GetOrisForBaseEuropean(baseOri, oriRight, oriBelow)
            End If
            Dim rangeH As Double = 0, rangeV As Double = 0
            If Not MeasureRange3ViewsAtScale1(app, modelPath, cleanTemplatePath, isSheetMetal, baseOri, oriRight, oriBelow, False, rangeH, rangeV) Then
                StepLog("[ALZADO-PD] No se pudo medir rango.")
                Return Nothing
            End If
            StepLog($"[ALZADO-PD] rangeH={rangeH:0.000000} rangeV={rangeV:0.000000}")
            If sizesOpt.HasValue AndAlso Not useLayoutByFold Then
                Dim baseW As Double = 0, baseH As Double = 0
                Select Case CType(baseOri, ViewOrientationConstants)
                    Case ViewOrientationConstants.igFrontView : baseW = sizesOpt.Value.W_Front : baseH = sizesOpt.Value.H_Front
                    Case ViewOrientationConstants.igTopView : baseW = sizesOpt.Value.W_Top : baseH = sizesOpt.Value.H_Top
                    Case ViewOrientationConstants.igRightView : baseW = sizesOpt.Value.W_Right : baseH = sizesOpt.Value.H_Right
                    Case ViewOrientationConstants.igLeftView : baseW = sizesOpt.Value.W_Right : baseH = sizesOpt.Value.H_Right
                    Case Else : baseW = sizesOpt.Value.W_Front : baseH = sizesOpt.Value.H_Front
                End Select
                If baseW > 0 AndAlso baseH >= ALZADO_TALL_RATIO * baseW Then
                    blockRotation = ViewRotation.RotMinus90
                    Dim mapRot As ProjectedViewMap = GetProjectedViewMap(baseOri, blockRotation)
                    oriRight = OrthoViewToSolidEdge(mapRot.Right)
                    oriBelow = OrthoViewToSolidEdge(mapRot.Down)
                End If
            End If
            Dim templateAreas = GetTemplateUsableAreas(templates)
            PickFormatAndScale(rangeH, rangeV, templateAreas, chosenTemplate, chosenScale)
        End If

        If String.IsNullOrWhiteSpace(chosenTemplate) OrElse chosenScale <= 0 Then
            StepLog("[ALZADO-PD] No se encontró formato/escala válido.")
            Return Nothing
        End If
        StepLog($"[ALZADO-PD] Formato={Path.GetFileName(chosenTemplate)} escala={chosenScale}")
        StepLog($"[EUROPEAN] Base={OriToConstantName(baseOri)} Rot={blockRotation} -> derecha={OriToConstantName(oriRight)} debajo={OriToConstantName(oriBelow)}")

        Dim m As Margins = GetMarginsByTemplate(chosenTemplate)
        Dim dft As DraftDocument = Nothing
        Try
            dft = CType(app.Documents.Add("SolidEdge.DraftDocument", chosenTemplate), DraftDocument)
            DoIdleSafe(app, "ALZADO-PD Draft Add")
            Dim sheet As Sheet = dft.ActiveSheet
            Dim modelLink As ModelLink = dft.ModelLinks.Add(modelPath)
            DoIdleSafe(app, "ALZADO-PD ModelLink")

            Dim H As Double = sheet.SheetSetup.SheetHeight
            Dim x0 As Double, y0 As Double
            If useFixedComposition Then
                x0 = x0Insert : y0 = y0Insert
                StepLog($"[ALZADO-PD] Insert pos FixedComposition: x0={x0:0.000} y0={y0:0.000} gapR={gapR:0.000} gapB={gapB:0.000}")
            Else
                x0 = m.Left + INSERT_OFFSET_FROM_MARGIN
                y0 = (H - m.Top) - INSERT_OFFSET_FROM_MARGIN
            End If

            ' Log cambio de plano / sistema de coordenadas (antes/después)
            StepLog("[SISTEMA-COORDS] === Cambio de plano de vista ===")
            StepLog($"[SISTEMA-COORDS] Plano ACTUAL del modelo (alzado por defecto): {OriToConstantName(CInt(ViewOrientationConstants.igFrontView))}")
            StepLog($"[SISTEMA-COORDS] Plano ELEGIDO como nuevo Alzado (por área): {OriToConstantName(baseOri)}")
            If baseOri = CInt(ViewOrientationConstants.igFrontView) Then
                StepLog("[SISTEMA-COORDS] No hay cambio: el modelo ya tiene Front como alzado.")
            Else
                StepLog($"[SISTEMA-COORDS] Cambio: de {OriToConstantName(CInt(ViewOrientationConstants.igFrontView))} -> a {OriToConstantName(baseOri)} como alzado.")
                StepLog("[SISTEMA-COORDS] Rotaremos vistas auxiliares (derecha/debajo) para alinearlas; NO tocamos SCoords de la pieza.")
            End If
            StepLog($"[SISTEMA-COORDS] Planos que COLOCAMOS: Alzado={OriToConstantName(baseOri)} | Derecha={OriToConstantName(oriRight)} | Debajo={OriToConstantName(oriBelow)}")

            ' Log escala antes de insertar
            StepLog($"[SCALE] computed={chosenScale}")
            Dim vBase As DrawingView = Nothing, vRight As DrawingView = Nothing, vBelow As DrawingView = Nothing
            If useLayoutByFold Then
                ' NUEVA LÓGICA: AddByFold, base 30mm del borde, gaps 20mm
                If Not InsertThreeViewsByFold(app, sheet, modelLink, isSheetMetal, chosenScale, baseOri, layoutByFoldBase.Rotated,
                                             layoutByFoldBase.BaseWidth, layoutByFoldBase.BaseHeight,
                                             chosenTemplate, vBase, vRight, vBelow, oriRight, oriBelow) Then
                    StepLog("[ALZADO-PD] ERROR: InsertThreeViewsByFold falló.")
                    Return Nothing
                End If
            Else
                InsertStandard3Internal(app, sheet, modelLink, isSheetMetal, chosenScale, baseOri, oriRight, oriBelow,
                                       x0, y0, gapR, gapB, vBase, vRight, vBelow)
            End If
            DoIdleSafe(app, "ALZADO-PD Insert3")
            StepLog($"[SCALE] applied to baseView={chosenScale} (vistas proyectadas usan misma escala)")

            If vBase Is Nothing OrElse vRight Is Nothing OrElse vBelow Is Nothing Then
                StepLog("[ALZADO-PD] ERROR: alguna vista no se insertó correctamente (Base/Right/Below)")
            End If

            SafeUpdateView(vBase, "Base")
            DoIdleSafe(app, "ALZADO-PD after Base.Update")
            SafeUpdateView(vRight, "Derecha")
            DoIdleSafe(app, "ALZADO-PD after Right.Update")
            SafeUpdateView(vBelow, "Debajo")
            DoIdleSafe(app, "ALZADO-PD Update views")

            If Not useFixedComposition AndAlso Not useLayoutByFold Then
                Align3ViewsToBase(app, vBase, vRight, vBelow)
                DoIdleSafe(app, "ALZADO-PD Align")
                RepositionViewsWithGap(app, vBase, vRight, vBelow)
                DoIdleSafe(app, "ALZADO-PD RepositionWithGap")
            End If

            ' Rotar vistas auxiliares cuando base<>Front (alinear proyecciones al sistema europeo).
            ' Con LayoutByFold, AddByFold crea proyecciones ya orientadas → no aplicar rotación auxiliar.
            If baseOri <> CInt(ViewOrientationConstants.igFrontView) AndAlso Not useLayoutByFold Then
                Dim rotRight As Double = GetAuxViewRotationRadians(baseOri, oriRight, True)
                Dim rotBelow As Double = GetAuxViewRotationRadians(baseOri, oriBelow, False)
                StepLog($"[ROT-AUX] Base={OriToConstantName(baseOri)} -> Derecha({OriToConstantName(oriRight)}) rot={rotRight * 180 / Math.PI:0}º | Debajo({OriToConstantName(oriBelow)}) rot={rotBelow * 180 / Math.PI:0}º")
                SafeSetRotationAngle(vRight, rotRight)
                DoIdleSafe(app, "ALZADO-PD RotRight")
                SafeSetRotationAngle(vBelow, rotBelow)
                DoIdleSafe(app, "ALZADO-PD RotBelow")
                SafeUpdateView(vRight, "Derecha tras rot")
                SafeUpdateView(vBelow, "Debajo tras rot")
                DoIdleSafe(app, "ALZADO-PD Update post-rot")
            End If

            ' Rotar bloque según LayoutEngine o condición H>=2.5*W
            Dim blockRotated As Boolean = False
            Dim angleRad As Double = 0
            If blockRotation = ViewRotation.RotMinus90 Then angleRad = -Math.PI / 2.0
            If blockRotation = ViewRotation.RotPlus90 Then angleRad = Math.PI / 2.0
            If Math.Abs(angleRad) > 0.0001 Then
                blockRotated = RotateBlock3ViewsByAngle(app, vBase, vRight, vBelow, angleRad, True)
            End If
            DoIdleSafe(app, "ALZADO-PD RotBlockIfTall")

            ' Reposicionar las 3 vistas tras rotación para que coincidan con zonas del plan (piezas giradas)
            If useFixedComposition AndAlso blockRotated AndAlso fclPlan IsNot Nothing AndAlso fclPlan.Zones IsNot Nothing Then
                Dim mx As Double, my As Double, rx As Double, ry As Double, bx As Double, by As Double
                FixedCompositionLayout.GetBlockRepositionParams(fclPlan, mx, my, rx, ry, bx, by)
                MoveViewTopLeft(app, vBase, mx, my, "Base(Reposicion tras rot)")
                MoveViewTopLeft(app, vRight, rx, ry, "Right(Reposicion tras rot)")
                MoveViewTopLeft(app, vBelow, bx, by, "Below(Reposicion tras rot)")
                DoIdleSafe(app, "ALZADO-PD Reposicion block")
            End If

            ' Segundo pase de Update para forzar refresco completo de todas las vistas
            SafeUpdateView(vBase, "Base final")
            SafeUpdateView(vRight, "Derecha final")
            SafeUpdateView(vBelow, "Debajo final")
            DoIdleSafe(app, "ALZADO-PD Update final")

            ' Vista isométrica (igTopFrontRightView)
            Dim vIso As DrawingView = Nothing
            Dim includeIso As Boolean = True
            Dim isoX As Double = Double.NaN, isoY As Double = Double.NaN
            If useFixedComposition AndAlso fclPlan IsNot Nothing Then
                includeIso = fclPlan.IncludeIso
                If includeIso Then
                    Dim isoPt = FixedCompositionLayout.GetIsoPositionForFixedComposition(fclPlan)
                    isoX = isoPt.X : isoY = isoPt.Y
                End If
            End If
            If includeIso Then
                Dim isoDrawingScale As Double = Double.NaN
                If useFixedComposition AndAlso TemplateBboxLayout.UseSlotBasedLayout AndAlso fclPlan IsNot Nothing Then
                    Dim uIso = LayoutEngine.GetUsableAreaForTemplate(chosenTemplate)
                    If uIso IsNot Nothing Then
                        isoDrawingScale = TemplateBboxLayout.ComputeIsoUneScale(
                            chosenScale, fclPlan.MainWidthAt1, fclPlan.MainHeightAt1, uIso)
                    End If
                End If
                InsertIsoView(app, sheet, modelLink, isSheetMetal, chosenScale, m, vBase, vRight, vBelow, vIso, isoX, isoY, isoDrawingScale)
                ' Con FixedComposition, asegurar que ISO quede en la posición calculada
                If useFixedComposition AndAlso vIso IsNot Nothing AndAlso Not Double.IsNaN(isoX) AndAlso Not Double.IsNaN(isoY) Then
                    MoveViewTopLeft(app, vIso, isoX, isoY, "Iso(FixedComposition)")
                    SafeUpdateView(vIso, "Iso.Update(pos)")
                End If
            End If
            DoIdleSafe(app, "ALZADO-PD InsertIso")

            ' Vista Flat (desarrollada) si PSM - centrada en hueco restante (FixedComposition) o layout clásico
            Dim vFlat As DrawingView = Nothing
            If isSheetMetal Then
                flatInserted = TryCreateFlatView_Safe(sheet.DrawingViews, modelLink, chosenScale, vFlat, app, chosenTemplate)
                If flatInserted AndAlso vFlat IsNot Nothing Then
                    SafeUpdateView(vFlat, "Flat.Update")
                    DoIdleSafe(app, "ALZADO-PD Flat")
                    Dim flatW As Double = 0, flatH As Double = 0
                    GetViewSizeSmart(vFlat, flatW, flatH)
                    ' Rotar Flat SOLO si la vista principal se ha girado (no por aspect ratio de la flat)
                    Dim needFlatRotate As Boolean = blockRotated
                    If needFlatRotate Then
                        Dim currentFlatRad As Double = 0
                        Try : vFlat.GetRotationAngle(currentFlatRad) : Catch : End Try
                        SafeSetRotationAngle(vFlat, currentFlatRad - Math.PI / 2.0)
                        DoIdleSafe(app, "ALZADO-PD Flat rotate")
                        SafeUpdateView(vFlat, "Flat.Update rotated")
                    End If
                    ' Posicionar Flat según FixedComposition si aplica
                    If useFixedComposition AndAlso fclPlan IsNot Nothing AndAlso fclPlan.IncludeFlat Then
                        Dim flatPt = FixedCompositionLayout.GetFlatPositionForFixedComposition(fclPlan)
                        MoveViewTopLeft(app, vFlat, flatPt.X, flatPt.Y, "Flat(FixedComposition)")
                        SafeUpdateView(vFlat, "Flat.Update(pos)")
                    End If
                End If
            End If

            ' Layout: posicionar 3 vistas + ISO + Flat según plantilla (Base=front, Below=top, Right=right)
            ' FixedComposition ya colocó correctamente; solo aplicar layout clásico si no usamos FixedComposition
            If Not useFixedComposition Then
                Dim sheetW As Double, sheetH As Double
                GetSheetSize(sheet, sheetW, sheetH)
                Dim bw As Double, bh As Double : GetViewSizeSmart(vBase, bw, bh)
                Dim rw As Double, rh As Double : GetViewSizeSmart(vRight, rw, rh)
                Dim blw As Double, blh As Double : GetViewSizeSmart(vBelow, blw, blh)
                Dim iw As Double = 0, ih As Double = 0
                If vIso IsNot Nothing Then GetViewSizeSmart(vIso, iw, ih)
                ApplyLayout_IsoTop_FlatBottom(app, sheetW, sheetH, m, vBase, bw, bh, vBelow, blw, blh, vRight, rw, rh, vIso, iw, ih, vFlat, GAP_H, GAP_V, flatInserted)
                CenterAllViewsOnUsableArea(app, sheet.DrawingViews, sheetW, sheetH, m)
                ' La ISO debe quedar pegada al margen derecho (CenterAllViews la desplaza incorrectamente)
                If vIso IsNot Nothing AndAlso iw > 0 AndAlso ih > 0 Then
                    Dim isoPt = FixedCompositionLayout.GetIsoPositionForClassicLayout(sheetW, iw, ih)
                    MoveViewTopLeft(app, vIso, isoPt.X, isoPt.Y, "Iso(ClassicLayout-pegadaDer)")
                    SafeUpdateView(vIso, "Iso.Update(posClassic)")
                End If
            ElseIf vIso IsNot Nothing Then
                ' FixedComposition colocó ISO; MoveViewTopLeft ya aplicado por posición
                SafeUpdateView(vIso, "Iso.Update(final)")
            End If
            DoIdleSafe(app, "ALZADO-PD CenterAll")

            If useFixedComposition AndAlso fclPlan IsNot Nothing AndAlso TemplateBboxLayout.UseSlotBasedLayout Then
                TemplateBboxLayout.ApplySlotCentersAfterInsert(
                    app, chosenTemplate, fclPlan, vBase, vRight, vBelow, vIso, vFlat, flatInserted,
                    Sub(msg) StepLog(msg))
                If vBase IsNot Nothing Then SafeUpdateView(vBase, "Base.BboxSlot")
                If vRight IsNot Nothing Then SafeUpdateView(vRight, "Right.BboxSlot")
                If vBelow IsNot Nothing Then SafeUpdateView(vBelow, "Below.BboxSlot")
                If vIso IsNot Nothing Then SafeUpdateView(vIso, "Iso.BboxSlot")
                If flatInserted AndAlso vFlat IsNot Nothing Then SafeUpdateView(vFlat, "Flat.BboxSlot")
                DoIdleSafe(app, "ALZADO-PD BboxSlot centers")
            End If

            Try
                dft.Activate()
            Catch ex As Exception
                LogEx("dft.Activate", ex)
            End Try
            DoIdleSafe(app, "ALZADO-PD Activate")

            LogViewOriginsAndDistances(vBase, vRight, vBelow, baseOri, oriRight, oriBelow)

            StepLog($"[ALZADO-PD] 3 vistas + ISO + {If(flatInserted, "Flat", "sin Flat")}. Primer diedro.")
            Return dft
        Catch ex As Exception
            LogEx("CreateDraftAlzadoPrimerDiedro", ex)
            If dft IsNot Nothing Then Try : dft.Close(False) : Catch : End Try
            Return Nothing
        End Try
    End Function

    '======================================================================
    ' FUNCIÓN PÚBLICA REUTILIZABLE
    ' - NO selecciona archivo
    ' - NO arranca Solid Edge
    ' - NO registra/revoca OleMessageFilter
    ' - DEVUELVE el DraftDocument final SIN cerrarlo
    '======================================================================
    Public Function CreateBestFitDraft(app As SolidEdgeFramework.Application,
                                       modelPath As String,
                                       templates As String(),
                                       scales As Double(),
                                       ByRef flatInserted As Boolean,
                                       Optional cleanTemplatePath As String = Nothing) As SolidEdgeDraft.DraftDocument

        flatInserted = False

        If app Is Nothing Then Throw New ArgumentNullException(NameOf(app))
        If String.IsNullOrWhiteSpace(modelPath) Then Throw New ArgumentException("modelPath vacío.")
        If templates Is Nothing OrElse templates.Length = 0 Then Throw New ArgumentException("No se han proporcionado templates.")
        If scales Is Nothing OrElse scales.Length = 0 Then Throw New ArgumentException("No se han proporcionado escalas.")
        If Not File.Exists(modelPath) Then Throw New FileNotFoundException("No existe el modelo.", modelPath)

        StepLog($"CreateBestFitDraft() -> modelPath={modelPath}")

        Dim isSheetMetal As Boolean = modelPath.EndsWith(".psm", StringComparison.OrdinalIgnoreCase)
        StepLog($"isSheetMetal={isSheetMetal}")

        Dim useReverseAlgo As Boolean = Not String.IsNullOrWhiteSpace(cleanTemplatePath) AndAlso File.Exists(cleanTemplatePath)
        If useReverseAlgo Then
            StepLog("[REVERSE] Usando algoritmo inverso con DXF_LIMPIO para medición.")
            Dim result = CreateBestFitDraftReverse(app, modelPath, templates, cleanTemplatePath, flatInserted, isSheetMetal, 0)
            If result IsNot Nothing Then Return result
            StepLog("[REVERSE] Falló, continúo con algoritmo original.")
        End If

        Dim scalesDescending As Boolean = True
        For si As Integer = 1 To scales.Length - 1
            If scales(si) > scales(si - 1) Then
                scalesDescending = False
                Exit For
            End If
        Next

        Dim success As Boolean = False
        Dim finalDoc As DraftDocument = Nothing

        Try
            app.DisplayAlerts = False

            For Each tpl In templates
                If String.IsNullOrWhiteSpace(tpl) Then Continue For

                If Not File.Exists(tpl) Then
                    StepLog($"[SKIP] No existe template: {tpl}")
                    Continue For
                End If

                Dim m As Margins = GetMarginsByTemplate(tpl)
                StepLog($"Plantilla: {tpl}")
                StepLog($"Margins(m): L={m.Left} R={m.Right} T={m.Top} B={m.Bottom}")

                Dim sheetW0 As Double, sheetH0 As Double
                Dim dftTemp As DraftDocument = Nothing
                Try
                    dftTemp = CType(app.Documents.Add("SolidEdge.DraftDocument", tpl), DraftDocument)
                    GetSheetSize(dftTemp.ActiveSheet, sheetW0, sheetH0)
                Finally
                    If dftTemp IsNot Nothing Then Try : dftTemp.Close(False) : Catch : End Try
                End Try
                Dim usableW0 As Double = sheetW0 - m.Left - m.Right
                Dim usableH0 As Double = sheetH0 - m.Top - m.Bottom

                Dim baseOri As Integer = CInt(ViewOrientationConstants.igRightView)
                Dim baseRotated As Boolean = False
                Dim initialScale As Double = scales(0)
                Dim sizesOpt = GetBaseViewSizesAtScale1(app, modelPath, tpl, isSheetMetal)
                If sizesOpt.HasValue Then
                    Dim baseW As Double = 0, baseH As Double = 0
                    SelectBaseViewByArea(sizesOpt.Value, baseOri, baseW, baseH, baseRotated)
                    Dim adjW As Double = GetAdjacentW(sizesOpt.Value, baseOri)
                    initialScale = ComputeInitialScale(usableW0, baseW, adjW)
                    StepLog($"[ALGO] baseOri={baseOri} rotated={baseRotated} initialScale={initialScale:0.000}")
                End If

                Dim scalesToTry As Double() = {initialScale}.Concat(scales.Where(Function(s) Math.Abs(s - initialScale) > 0.001)).ToArray()

                Dim bestScale As Double = 0
                Dim bestUtil As Double = -1
                Dim bestDraft As DraftDocument = Nothing

                For Each sc In scalesToTry
                    Dim dft As DraftDocument = Nothing

                    Try
                        StepLog($"--- PROBANDO: tpl='{Path.GetFileName(tpl)}' escala={sc} ---")

                        dft = CType(app.Documents.Add("SolidEdge.DraftDocument", tpl), DraftDocument)
                        DoIdleSafe(app, "after Draft Add")

                        LogSheetWindowOrigin(app, dft, "AFTER Draft Add")

                        Dim sheet As Sheet = dft.ActiveSheet
                        Dim sheetW As Double, sheetH As Double
                        GetSheetSize(sheet, sheetW, sheetH)

                        Dim usableW As Double = sheetW - m.Left - m.Right
                        Dim usableH As Double = sheetH - m.Top - m.Bottom
                        StepLog($"UsableW/H = ({usableW} x {usableH})")

                        StepLog("Creando ModelLink...")
                        Dim modelLink As ModelLink = dft.ModelLinks.Add(modelPath)
                        StepLog("ModelLink OK.")

                        Dim vFront As DrawingView = Nothing
                        Dim vTop As DrawingView = Nothing
                        Dim vRight As DrawingView = Nothing
                        Dim vIso As DrawingView = Nothing

                        StepLog($"Insertando 4 vistas (baseOri={baseOri} baseRotated={baseRotated})...")
                        InsertStandard4(app, sheet, modelLink, isSheetMetal, sc, baseOri, baseRotated, m, vFront, vTop, vRight, vIso)
                        DoIdleSafe(app, "after InsertStandard4")

                        LogSheetWindowOrigin(app, dft, "AFTER InsertStandard4")

                        StepLog("=== DUMP (just created) ===")
                        DumpOneView("FRONT just created", vFront)
                        DumpOneView("TOP   just created", vTop)
                        DumpOneView("RIGHT just created", vRight)
                        DumpOneView("ISO   just created", vIso)

                        SafeUpdateView(vFront, "Front.Update") : DoIdleSafe(app, "after Front.Update")
                        SafeUpdateView(vTop, "Top.Update") : DoIdleSafe(app, "after Top.Update")
                        SafeUpdateView(vRight, "Right.Update") : DoIdleSafe(app, "after Right.Update")
                        SafeUpdateView(vIso, "Iso.Update") : DoIdleSafe(app, "after Iso.Update")

                        If baseRotated Then
                            StepLog($"Vista base (H>=2*W) -> rotar bloque de 3 vistas -90º")
                            ApplyRotationToPrincipalViews(app, baseOri, vFront, vTop, vRight)
                        End If

                        Dim rw0 As Double, rh0 As Double
                        GetViewSizeSmart(vRight, rw0, rh0)

                        ' Solo rotar vista derecha por aspecto si NO hemos rotado ya el bloque
                        If Not baseRotated AndAlso ShouldRotateRight(rw0, rh0) Then
                            StepLog($"Right: w={rw0:0.000000} h={rh0:0.000000} -> ROTAR 90º")
                            SafeSetRotationAngle(vRight, -RIGHT_ROT_RAD)

                            ForceSameOriginY(app, vFront, vRight, "Right same row as Front")
                            SafeUpdateView(vRight, "Right.Update (after align to Right)")
                            DoIdleSafe(app, "after Front align")
                        Else
                            StepLog($"Right: w={rw0:0.000000} h={rh0:0.000000} -> NO rotar, solo centrar/alinear")
                        End If

                        StepLog("=== DUMP (after update + rotation) ===")
                        DumpOneView("FRONT after update", vFront)
                        DumpOneView("TOP   after update", vTop)
                        DumpOneView("RIGHT after rotation", vRight)
                        DumpOneView("ISO   after update", vIso)

                        Dim fw As Double, fh As Double : GetViewSizeSmart(vFront, fw, fh)
                        Dim tw As Double, th As Double : GetViewSizeSmart(vTop, tw, th)
                        Dim rw As Double, rh As Double : GetViewSizeSmart(vRight, rw, rh)
                        Dim iw As Double, ih As Double : GetViewSizeSmart(vIso, iw, ih)

                        Dim fits4 As Boolean = LayoutFits_4(usableW, usableH, fw, fh, tw, th, rw, rh, iw, ih, GAP_H, GAP_V)
                        If Not fits4 Then
                            StepLog("No cabe (4 vistas). Cierro y pruebo otra escala.")
                            CloseDraftNoSave(dft) : dft = Nothing
                            Continue For
                        End If

                        Dim util As Double = LayoutUtil_4(usableW, usableH, fw, fh, tw, th, rw, rh, iw, ih, GAP_H, GAP_V)
                        StepLog($"FIT=TRUE (4 vistas). Utilización={util:0.000}")

                        If util > bestUtil + 0.000001 Then
                            If bestDraft IsNot Nothing Then
                                CloseDraftNoSave(bestDraft)
                                bestDraft = Nothing
                            End If

                            bestUtil = util
                            bestScale = sc
                            bestDraft = dft
                            dft = Nothing

                            StepLog($"*** NUEVO BEST: escala={bestScale} util={bestUtil:0.000}")

                            If bestUtil >= TARGET_UTIL Then
                                StepLog($"Alcanza TARGET_UTIL={TARGET_UTIL}. No sigo probando escalas.")
                                Exit For
                            End If
                        Else
                            CloseDraftNoSave(dft) : dft = Nothing
                        End If

                        ' OPTIMIZACIÓN:
                        ' Si las escalas vienen en orden descendente (grande->pequeña),
                        ' la primera que “cabe” es la MEJOR posible para esta plantilla.
                        ' No tiene sentido probar escalas más pequeñas (solo empeora).
                        '
                        ' EXCEPCIÓN: si es PSM y NO se ha podido insertar FLAT todavía,
                        ' dejamos que siga probando escalas más pequeñas para ver si entra el FLAT.
                        If scalesDescending Then
                            Dim isPsm As Boolean = modelPath.ToLowerInvariant().EndsWith(".psm")
                            If (Not isPsm) OrElse flatInserted Then
                                StepLog("   [OPT] Ya cabe y escalas descendentes -> no pruebo escalas más pequeñas.")
                                Exit For
                            End If
                        End If

                    Catch ex As Exception
                        LogEx($"Probe tpl={tpl} sc={sc}", ex)
                        If dft IsNot Nothing Then CloseDraftNoSave(dft)
                    End Try
                Next

                If bestDraft Is Nothing Then
                    StepLog($"[WARN] No encontré escala válida en plantilla {Path.GetFileName(tpl)}")
                    Continue For
                End If

                StepLog($"Usando BEST escala={bestScale} util={bestUtil:0.000} en plantilla {Path.GetFileName(tpl)}")
                finalDoc = bestDraft
                success = True

                ' Layout final + FLAT si PSM
                Try
                    Dim sheet As Sheet = finalDoc.ActiveSheet
                    Dim sheetW As Double, sheetH As Double
                    GetSheetSize(sheet, sheetW, sheetH)

                    Dim modelLink As ModelLink = finalDoc.ModelLinks.Item(1)
                    Dim dvws As DrawingViews = sheet.DrawingViews

                    Dim created As DrawingView() = dvws.OfType(Of DrawingView)().ToArray()
                    If created.Length < 4 Then Throw New Exception("No encuentro las 4 vistas creadas.")

                    Dim vFront = created(created.Length - 4)
                    Dim vTop = created(created.Length - 3)
                    Dim vRight = created(created.Length - 2)
                    Dim vIso = created(created.Length - 1)

                    Dim fw As Double, fh As Double : GetViewSizeSmart(vFront, fw, fh)
                    Dim tw As Double, th As Double : GetViewSizeSmart(vTop, tw, th)
                    Dim rw As Double, rh As Double : GetViewSizeSmart(vRight, rw, rh)
                    Dim iw As Double, ih As Double : GetViewSizeSmart(vIso, iw, ih)

                    Dim hasFlat As Boolean = False
                    Dim vFlat As DrawingView = Nothing
                    Dim flw As Double = 0, flh As Double = 0

                    If isSheetMetal Then

                        StepLog("Creando vista DESARROLLADA (Flat)...")
                        hasFlat = TryCreateFlatView_Safe(dvws, modelLink, bestScale, vFlat, app)
                        DoIdleSafe(app, "after TryCreateFlatView")

                        If hasFlat AndAlso vFlat IsNot Nothing Then
                            SafeUpdateView(vFlat, "Flat.Update (initial)")
                            DoIdleSafe(app, "after Flat.Update initial")
                            If baseRotated AndAlso (CType(baseOri, ViewOrientationConstants) = ViewOrientationConstants.igRightView OrElse CType(baseOri, ViewOrientationConstants) = ViewOrientationConstants.igFrontView) Then
                                SafeSetRotationAngle(vFlat, -RIGHT_ROT_RAD)
                                DoIdleSafe(app, "Flat rotate (main)")
                                SafeUpdateView(vFlat, "Flat.Update rotated (main)")
                            End If
                            GetViewSizeSmart(vFlat, flw, flh)
                            StepLog($"Flat initial size=({flw} x {flh})")
                        Else
                            StepLog("[WARN] No pude crear Flat (no disponible). Continuo sin Flat.")
                        End If
                    End If

                    flatInserted = hasFlat

                    ApplyLayout_IsoTop_FlatBottom(
                        app,
                        sheetW, sheetH, m,
                        vFront, fw, fh,
                        vTop, tw, th,
                        vRight, rw, rh,
                        vIso, iw, ih,
                        vFlat,
                        GAP_H, GAP_V,
                        hasFlat
                    )
                    DoIdleSafe(app, "after ApplyLayout main")

                    If hasFlat AndAlso vFlat IsNot Nothing Then
                        StepLog("Ajustando FLAT bajo ISO (recrear con escala final)...")

                        Dim xMinIso As Double, yMinIso As Double, xMaxIso As Double, yMaxIso As Double
                        If Not TryGetViewRange(vIso, xMinIso, yMinIso, xMaxIso, yMaxIso) Then
                            StepLog("[WARN] No puedo leer RANGE de ISO para calcular hueco de FLAT.")
                        Else
                            Dim isoLeft As Double = xMinIso
                            Dim isoBottom As Double = yMinIso

                            Dim availW As Double = (sheetW - m.Right) - isoLeft
                            Dim availH As Double = (isoBottom - GAP_V) - m.Bottom

                            GetViewSizeSmart(vFlat, flw, flh)

                            If flw > EPS AndAlso flh > EPS AndAlso availW > EPS AndAlso availH > EPS Then
                                Dim factorW As Double = availW / flw
                                Dim factorH As Double = availH / flh
                                Dim factor As Double = Math.Min(factorW, factorH) * FLAT_FIT_SAFETY

                                StepLog($"Hueco Flat: availW={availW} availH={availH} | flatW={flw} flatH={flh} => factor={factor}")

                                Dim finalFlatScale As Double = bestScale * factor

                                Try
                                    vFlat.Delete()
                                    StepLog("Flat provisional borrada (Delete).")
                                Catch ex As Exception
                                    LogEx("vFlat.Delete()", ex)
                                End Try

                                vFlat = Nothing
                                hasFlat = False
                                DoIdleSafe(app, "after Flat.Delete")

                                StepLog($"Creando FLAT definitiva con escala={finalFlatScale}...")


                                hasFlat = TryCreateFlatView_Safe(dvws, modelLink, finalFlatScale, vFlat, app)
                                DoIdleSafe(app, "after TryCreateFlatView (final)")

                                If hasFlat AndAlso vFlat IsNot Nothing Then
                                    SafeUpdateView(vFlat, "Flat.Update (final)")
                                    DoIdleSafe(app, "after Flat.Update final")
                                    If baseRotated AndAlso (CType(baseOri, ViewOrientationConstants) = ViewOrientationConstants.igRightView OrElse CType(baseOri, ViewOrientationConstants) = ViewOrientationConstants.igFrontView) Then
                                        SafeSetRotationAngle(vFlat, -RIGHT_ROT_RAD)
                                        DoIdleSafe(app, "Flat rotate (final)")
                                        SafeUpdateView(vFlat, "Flat.Update rotated (final)")
                                    End If

                                    Dim x4 As Double = isoLeft
                                    Dim yTop As Double = sheetH - m.Top

                                    Dim isoW2 As Double, isoH2 As Double
                                    GetViewSizeSmart(vIso, isoW2, isoH2)

                                    MoveViewTopLeft(app, vIso, x4, yTop, "Iso(repack)")
                                    SafeUpdateView(vIso, "Iso.Update(repack)")
                                    DoIdleSafe(app, "after Iso repack")

                                    Dim yFlatTop As Double = yTop - isoH2 - GAP_V
                                    MoveViewTopLeft(app, vFlat, x4, yFlatTop, "Flat(repack)")
                                    SafeUpdateView(vFlat, "Flat.Update(repack)")
                                    DoIdleSafe(app, "after Flat repack")
                                Else
                                    StepLog("[WARN] No pude recrear Flat definitiva.")
                                End If

                                flatInserted = hasFlat
                            Else
                                StepLog("[WARN] No puedo calcular ajuste FLAT (dimensiones/hueco inválidos).")
                            End If
                        End If
                    End If

                    StepLog("Centrando todas las vistas dentro del área útil...")
                    CenterAllViewsOnUsableArea(app, sheet.DrawingViews, sheetW, sheetH, m)
                    DoIdleSafe(app, "after CenterAllViews")

                    StepLog("=== ESTADO FINAL ===")
                    DumpAllViews(sheet.DrawingViews)

                Catch ex As Exception
                    LogEx("PostLayout", ex)
                End Try

                Exit For
            Next

            If Not success OrElse finalDoc Is Nothing Then
                Throw New Exception("No pude crear un Draft válido con ninguna plantilla/escala.")
            End If

            StepLog("FIN OK: Draft creado, rotación aplicada, layout + (Flat si PSM) ajustado y centrado. (No guardo automáticamente).")
            Return finalDoc

        Catch ex As Exception
            LogEx("CreateBestFitDraft()", ex)

            If finalDoc IsNot Nothing Then
                Try
                    finalDoc.Close(False)
                Catch
                End Try
            End If

            Throw
        End Try
    End Function

    '======================================================================
    ' ALGORITMO INVERSO: medir rango a escala 1 con DXF_LIMPIO, elegir formato y escala estándar
    '======================================================================
    Private Function CreateBestFitDraftReverse(app As SolidEdgeFramework.Application,
                                               modelPath As String,
                                               templates As String(),
                                               cleanTemplatePath As String,
                                               ByRef flatInserted As Boolean,
                                               isSheetMetal As Boolean, veDu As Double) As DraftDocument
        flatInserted = False
        Dim sizesOpt = GetBaseViewSizesAtScale1(app, modelPath, cleanTemplatePath, isSheetMetal)
        If Not sizesOpt.HasValue Then Return Nothing

        Dim baseOri As Integer = CInt(ViewOrientationConstants.igRightView)
        Dim baseRotated As Boolean = False
        Dim baseW As Double = 0, baseH As Double = 0
        SelectBaseViewByArea(sizesOpt.Value, baseOri, baseW, baseH, baseRotated)

        Dim oriRight As Integer, oriBelow As Integer
        ' Sistema europeo (primer diedro) - mismo que CreateDraftAlzadoPrimerDiedro
        GetOrisForBaseEuropean(baseOri, oriRight, oriBelow)

        ' Si base rotada, usar mapa rotado para vistas proyectadas
        Dim blockRotation As ViewRotation = If(baseRotated, ViewRotation.RotMinus90, ViewRotation.Rot0)
        If blockRotation = ViewRotation.RotMinus90 Then
            Dim mapRot As ProjectedViewMap = GetProjectedViewMap(baseOri, blockRotation)
            oriRight = OrthoViewToSolidEdge(mapRot.Right)
            oriBelow = OrthoViewToSolidEdge(mapRot.Down)
        End If

        Dim rangeH As Double = 0, rangeV As Double = 0
        If Not MeasureRange3ViewsAtScale1(app, modelPath, cleanTemplatePath, isSheetMetal, baseOri, oriRight, oriBelow, baseRotated, rangeH, rangeV) Then
            StepLog("[REVERSE] MeasureRange3ViewsAtScale1 falló.")
            Return Nothing
        End If

        StepLog($"[REVERSE] rangeH={rangeH:0.000000} rangeV={rangeV:0.000000}")

        Dim templateAreas = GetTemplateUsableAreas(templates)
        Dim chosenTemplate As String = Nothing
        Dim chosenScale As Double = 0
        PickFormatAndScale(rangeH, rangeV, templateAreas, chosenTemplate, chosenScale)

        If String.IsNullOrWhiteSpace(chosenTemplate) OrElse chosenScale <= 0 Then
            StepLog("[REVERSE] No se encontró formato/escala válido.")
            Return Nothing
        End If

        StepLog($"[REVERSE] Formato elegido: {Path.GetFileName(chosenTemplate)} escala={chosenScale}")

        Dim m As Margins = GetMarginsByTemplate(chosenTemplate)
        Dim dft As DraftDocument = Nothing
        Try
            dft = CType(app.Documents.Add("SolidEdge.DraftDocument", chosenTemplate), DraftDocument)
            DoIdleSafe(app, "REVERSE Draft Add")
            Dim sheet As Sheet = dft.ActiveSheet
            Dim modelLink As ModelLink = dft.ModelLinks.Add(modelPath)
            DoIdleSafe(app, "REVERSE ModelLink")

            Dim vFront As DrawingView = Nothing
            Dim vTop As DrawingView = Nothing
            Dim vRight As DrawingView = Nothing
            Dim vIso As DrawingView = Nothing
            InsertStandard4(app, sheet, modelLink, isSheetMetal, chosenScale, baseOri, baseRotated, m, vFront, vTop, vRight, vIso)
            DoIdleSafe(app, "REVERSE InsertStandard4")

            SafeUpdateView(vFront, "Front.Update")
            SafeUpdateView(vTop, "Top.Update")
            SafeUpdateView(vRight, "Right.Update")
            SafeUpdateView(vIso, "Iso.Update")
            DoIdleSafe(app, "REVERSE Update views")

            If baseRotated Then
                StepLog($"REVERSE: baseRotated -> rotar bloque de 3 vistas -90º")
                ApplyRotationToPrincipalViews(app, baseOri, vFront, vTop, vRight)
            Else
                ' Solo girar vista a la derecha por aspecto si aplica (nunca la inferior por aspecto)
                Dim rw0 As Double, rh0 As Double
                GetViewSizeSmart(vRight, rw0, rh0)
                If ShouldRotateRight(rw0, rh0) AndAlso vRight IsNot Nothing Then
                    SafeSetRotationAngle(vRight, -RIGHT_ROT_RAD)
                    DoIdleSafe(app, "REVERSE rotate viewRight by aspect")
                    If vFront IsNot Nothing Then
                        AlignViewCenterY(app, vRight, GetViewCenterY(vFront), "Right align CY")
                    End If
                    SafeUpdateView(vRight, "Right.Update rotated by aspect")
                End If
            End If

            Dim sheetW As Double, sheetH As Double
            GetSheetSize(sheet, sheetW, sheetH)

            Dim vFlat As DrawingView = Nothing
            If isSheetMetal Then
                Dim hasFlat = TryCreateFlatView_Safe(sheet.DrawingViews, modelLink, chosenScale, vFlat, app)
                flatInserted = hasFlat
                    If hasFlat Then
                        SafeUpdateView(vFlat, "Flat.Update")
                        DoIdleSafe(app, "REVERSE Flat")
                        ' Si hemos rotado el bloque de vistas, girar también la Flat para alineación
                        If baseRotated AndAlso (CType(baseOri, ViewOrientationConstants) = ViewOrientationConstants.igRightView OrElse CType(baseOri, ViewOrientationConstants) = ViewOrientationConstants.igFrontView) Then
                        SafeSetRotationAngle(vFlat, -RIGHT_ROT_RAD)
                        DoIdleSafe(app, "REVERSE Flat rotate")
                        SafeUpdateView(vFlat, "Flat.Update rotated")
                    End If
                End If
            End If

            Dim fw As Double, fh As Double : GetViewSizeSmart(vFront, fw, fh)
            Dim tw As Double, th As Double : GetViewSizeSmart(vTop, tw, th)
            Dim rw As Double, rh As Double : GetViewSizeSmart(vRight, rw, rh)
            Dim iw As Double, ih As Double : GetViewSizeSmart(vIso, iw, ih)
            ApplyLayout_IsoTop_FlatBottom(app, sheetW, sheetH, m, vFront, fw, fh, vTop, tw, th, vRight, rw, rh, vIso, iw, ih, vFlat, GAP_H, GAP_V, flatInserted)
            CenterAllViewsOnUsableArea(app, sheet.DrawingViews, sheetW, sheetH, m)
            DoIdleSafe(app, "REVERSE CenterAll")

            Return dft
        Catch ex As Exception
            LogEx("CreateBestFitDraftReverse", ex)
            If dft IsNot Nothing Then Try : dft.Close(False) : Catch : End Try
            Return Nothing
        End Try
    End Function

    Private Function MeasureRange3ViewsAtScale1(app As SolidEdgeFramework.Application,
                                                modelPath As String,
                                                cleanTemplatePath As String,
                                                isSheetMetal As Boolean,
                                                baseOri As Integer,
                                                oriRight As Integer,
                                                oriBelow As Integer,
                                                baseRotated As Boolean,
                                                ByRef rangeH As Double,
                                                ByRef rangeV As Double) As Boolean
        rangeH = 0 : rangeV = 0
        Dim dft As DraftDocument = Nothing
        Try
            dft = CType(app.Documents.Add("SolidEdge.DraftDocument", cleanTemplatePath), DraftDocument)
            DoIdleSafe(app, "MeasureRange Draft Add")
            Dim sheet As Sheet = dft.ActiveSheet
            Dim modelLink As ModelLink = dft.ModelLinks.Add(modelPath)
            DoIdleSafe(app, "MeasureRange ModelLink")

            Dim x0 As Double = 0.05
            Dim y0 As Double = 0.05
            Dim scale As Double = 1.0

            Dim vBase As DrawingView = Nothing
            Dim v2 As DrawingView = Nothing
            Dim v3 As DrawingView = Nothing

            InsertStandard3(app, sheet, modelLink, isSheetMetal, scale, baseOri, oriRight, oriBelow, x0, y0, vBase, v2, v3)
            DoIdleSafe(app, "MeasureRange Insert3")

            SafeUpdateView(vBase, "M base")
            SafeUpdateView(v2, "M v2")
            SafeUpdateView(v3, "M v3")
            DoIdleSafe(app, "MeasureRange Update")

            If baseRotated Then
                ApplyRotationToPrincipalViews(app, baseOri, vBase, v2, v3)
            End If

            Dim xmin As Double = Double.PositiveInfinity
            Dim ymin As Double = Double.PositiveInfinity
            Dim xmax As Double = Double.NegativeInfinity
            Dim ymax As Double = Double.NegativeInfinity
            For Each v In {vBase, v2, v3}
                If v Is Nothing Then Continue For
                Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
                If TryGetViewRange(v, x1, y1, x2, y2) Then
                    xmin = Math.Min(xmin, Math.Min(x1, x2))
                    ymin = Math.Min(ymin, Math.Min(y1, y2))
                    xmax = Math.Max(xmax, Math.Max(x1, x2))
                    ymax = Math.Max(ymax, Math.Max(y1, y2))
                End If
            Next

            rangeH = Math.Max(0, xmax - xmin) + GAP_H * 2
            rangeV = Math.Max(0, ymax - ymin) + GAP_V * 2
            Return (rangeH > 0 AndAlso rangeV > 0)
        Catch ex As Exception
            LogEx("MeasureRange3ViewsAtScale1", ex)
            Return False
        Finally
            If dft IsNot Nothing Then Try : dft.Close(False) : Catch : End Try
        End Try
    End Function

    ''' <summary>Alinea vRight y vBelow respecto a vBase (sistema europeo: derecha misma fila, debajo centrado).</summary>
    Private Sub Align3ViewsToBase(app As SolidEdgeFramework.Application,
                                 vBase As DrawingView,
                                 vRight As DrawingView,
                                 vBelow As DrawingView)
        If vBase Is Nothing Then Return
        Dim baseCX As Double, baseCY As Double
        If Not TryGetViewCenter(vBase, baseCX, baseCY) Then Return
        If vRight IsNot Nothing Then
            AlignViewCenterY(app, vRight, baseCY, "Right alineado a base")
            SafeUpdateView(vRight, "Right.Update(align)")
        End If
        If vBelow IsNot Nothing Then
            AlignViewCenterX(app, vBelow, baseCX, "Below alineado a base")
            SafeUpdateView(vBelow, "Below.Update(align)")
        End If
    End Sub

    ''' <summary>Si el bloque de las 3 vistas se sale de la hoja (fuera de márgenes), lo mueve para que quede dentro.</summary>
    Private Sub EnsureBlockWithinSheet(app As SolidEdgeFramework.Application,
                                       sheet As Sheet,
                                       m As Margins,
                                       vBase As DrawingView,
                                       vRight As DrawingView,
                                       vBelow As DrawingView)
        If sheet Is Nothing OrElse vBase Is Nothing Then Return
        Dim W As Double = sheet.SheetSetup.SheetWidth
        Dim H As Double = sheet.SheetSetup.SheetHeight
        Dim xmin As Double = Double.PositiveInfinity
        Dim ymin As Double = Double.PositiveInfinity
        Dim xmax As Double = Double.NegativeInfinity
        Dim ymax As Double = Double.NegativeInfinity
        For Each v As DrawingView In {vBase, vRight, vBelow}
            If v Is Nothing Then Continue For
            Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
            If TryGetViewRange(v, x1, y1, x2, y2) Then
                xmin = Math.Min(xmin, Math.Min(x1, x2))
                ymin = Math.Min(ymin, Math.Min(y1, y2))
                xmax = Math.Max(xmax, Math.Max(x1, x2))
                ymax = Math.Max(ymax, Math.Max(y1, y2))
            End If
        Next
        If xmin > xmax OrElse ymin > ymax Then Return
        Dim blockLeft As Double = Math.Min(xmin, xmax)
        Dim blockRight As Double = Math.Max(xmin, xmax)
        Dim blockBottom As Double = Math.Min(ymin, ymax)
        Dim blockTop As Double = Math.Max(ymin, ymax)
        Dim usableLeft As Double = m.Left
        Dim usableRight As Double = W - m.Right
        Dim usableBottom As Double = m.Bottom
        Dim usableTop As Double = H - m.Top
        Dim dx As Double = 0
        Dim dy As Double = 0
        If blockLeft < usableLeft Then
            dx = usableLeft - blockLeft
        ElseIf blockRight > usableRight Then
            dx = usableRight - blockRight
        End If
        If blockBottom < usableBottom Then
            dy = usableBottom - blockBottom
        ElseIf blockTop > usableTop Then
            dy = usableTop - blockTop
        End If
        If Math.Abs(dx) < EPS AndAlso Math.Abs(dy) < EPS Then Return
        StepLog($"[SHEET] Bloque fuera de hoja: mover dx={dx:0.0000} dy={dy:0.0000}")
        For Each v As DrawingView In {vBase, vRight, vBelow}
            If v Is Nothing Then Continue For
            Dim ox As Double = 0, oy As Double = 0
            v.GetOrigin(ox, oy)
            v.SetOrigin(ox + dx, oy + dy)
        Next
        DoIdleSafe(app, "EnsureBlockWithinSheet")
        SafeUpdateView(vBase, "Base dentro")
        SafeUpdateView(vRight, "Right dentro")
        SafeUpdateView(vBelow, "Below dentro")
    End Sub

    ''' <summary>Reposiciona Right y Below según rangos reales: gap mínimo entre bordes de vistas adyacentes.</summary>
    Private Sub RepositionViewsWithGap(app As SolidEdgeFramework.Application,
                                       vBase As DrawingView,
                                       vRight As DrawingView,
                                       vBelow As DrawingView)
        If vBase Is Nothing Then Return
        Dim xminB As Double, yminB As Double, xmaxB As Double, ymaxB As Double
        If Not TryGetViewRange(vBase, xminB, yminB, xmaxB, ymaxB) Then Return
        Dim baseCX As Double = (xminB + xmaxB) / 2.0
        Dim baseCY As Double = (yminB + ymaxB) / 2.0
        StepLog($"[GAP] Base rango: x=[{xminB:0.0000},{xmaxB:0.0000}] y=[{yminB:0.0000},{ymaxB:0.0000}] | minEdgeGap={ALZADO_MIN_EDGE_GAP}")

        If vRight IsNot Nothing Then
            Dim xminR As Double, yminR As Double, xmaxR As Double, ymaxR As Double
            If TryGetViewRange(vRight, xminR, yminR, xmaxR, ymaxR) Then
                Dim targetLeft As Double = xmaxB + ALZADO_MIN_EDGE_GAP
                If xminR < targetLeft Then
                    Dim dx As Double = targetLeft - xminR
                    StepLog($"[GAP] Right: xminR={xminR:0.0000} < baseRight+GAP={targetLeft:0.0000} => mover dx={dx:0.0000}")
                    Dim ox As Double = 0, oy As Double = 0
                    vRight.GetOrigin(ox, oy)
                    vRight.SetOrigin(ox + dx, oy)
                    SafeUpdateView(vRight, "Right.Update(gap)")
                    DoIdleSafe(app, "after Right gap")
                End If
            End If
            AlignViewCenterY(app, vRight, baseCY, "Right CY después de gap")
            SafeUpdateView(vRight, "Right.Update(CY)")
        End If

        If vBelow IsNot Nothing Then
            Dim xminBl As Double, yminBl As Double, xmaxBl As Double, ymaxBl As Double
            If TryGetViewRange(vBelow, xminBl, yminBl, xmaxBl, ymaxBl) Then
                Dim targetTop As Double = yminB - ALZADO_MIN_EDGE_GAP
                If ymaxBl > targetTop Then
                    Dim dy As Double = targetTop - ymaxBl
                    StepLog($"[GAP] Below: ymaxBl={ymaxBl:0.0000} > baseBottom-GAP={targetTop:0.0000} => mover dy={dy:0.0000}")
                    Dim ox As Double = 0, oy As Double = 0
                    vBelow.GetOrigin(ox, oy)
                    vBelow.SetOrigin(ox, oy + dy)
                    SafeUpdateView(vBelow, "Below.Update(gap)")
                    DoIdleSafe(app, "after Below gap")
                End If
            End If
            AlignViewCenterX(app, vBelow, baseCX, "Below CX después de gap")
            SafeUpdateView(vBelow, "Below.Update(CX)")
        End If
    End Sub

    ''' <summary>Obtiene el rango combinado de las 3 vistas. Range devuelve coords de hoja (comparables entre vistas).</summary>
    Private Function GetCombinedRange3Views(vBase As DrawingView, vRight As DrawingView, vBelow As DrawingView,
                                           ByRef xmin As Double, ByRef ymin As Double,
                                           ByRef xmax As Double, ByRef ymax As Double) As Boolean
        xmin = Double.MaxValue : ymin = Double.MaxValue
        xmax = Double.MinValue : ymax = Double.MinValue
        Dim anyOk As Boolean = False
        For Each v As DrawingView In {vBase, vRight, vBelow}
            If v Is Nothing Then Continue For
            Dim vx1 As Double, vy1 As Double, vx2 As Double, vy2 As Double
            If Not TryGetViewRange(v, vx1, vy1, vx2, vy2) Then Continue For
            anyOk = True
            Dim left As Double = Math.Min(vx1, vx2)
            Dim right As Double = Math.Max(vx1, vx2)
            Dim bot As Double = Math.Min(vy1, vy2)
            Dim top As Double = Math.Max(vy1, vy2)
            If left < xmin Then xmin = left
            If right > xmax Then xmax = right
            If bot < ymin Then ymin = bot
            If top > ymax Then ymax = top
        Next
        Return anyOk
    End Function

    ''' <summary>Rota vistas -90° SOLO cuando forceRotateBlock=True (vista base cumple H>=2.5*W).
    ''' Si forceRotateBlock=False, NO rotar ninguna vista (evita desalinear proyecciones).</summary>
    Private Function RotateBlock3ViewsIfTall(app As SolidEdgeFramework.Application,
                                            vBase As DrawingView, vRight As DrawingView, vBelow As DrawingView,
                                            Optional forceRotateBlock As Boolean = False) As Boolean
        Return RotateBlock3ViewsByAngle(app, vBase, vRight, vBelow, -Math.PI / 2.0, forceRotateBlock)
    End Function

    ''' <summary>Rota las 3 vistas por angleRad. Si force=False, no hace nada.</summary>
    Private Function RotateBlock3ViewsByAngle(app As SolidEdgeFramework.Application,
                                             vBase As DrawingView, vRight As DrawingView, vBelow As DrawingView,
                                             angleRad As Double, force As Boolean) As Boolean
        If Not force OrElse Math.Abs(angleRad) < 0.0001 Then Return False
        Dim anyRotated As Boolean = False
        Try
            For Each v As DrawingView In {vBase, vRight, vBelow}
                If v Is Nothing Then Continue For
                Dim currentRad As Double = 0
                Try : v.GetRotationAngle(currentRad) : Catch : End Try
                SafeSetRotationAngle(v, currentRad + angleRad, "block")
                DoIdleSafe(app, "ROT-BLOCK view")
                anyRotated = True
            Next
            If anyRotated Then
                SafeUpdateView(vBase, "Base tras rot-block")
                SafeUpdateView(vRight, "Derecha tras rot-block")
                SafeUpdateView(vBelow, "Debajo tras rot-block")
            End If
            Return anyRotated
        Catch ex As Exception
            LogEx("RotateBlock3ViewsByAngle", ex)
            Return False
        End Try
    End Function

    ''' <summary>Expone rotación sincronizada de las 3 vistas ortográficas (p. ej. <see cref="SlotBBoxViewLayout"/>).</summary>
    Public Function RotateDrawingViewsBlockByAngle(app As SolidEdgeFramework.Application,
                                                   vBase As DrawingView, vRight As DrawingView, vBelow As DrawingView,
                                                   angleRad As Double) As Boolean
        Return RotateBlock3ViewsByAngle(app, vBase, vRight, vBelow, angleRad, True)
    End Function

    ''' <summary>Log por pantalla: orígenes, rangos y distancias. Incluye nombres constantes (igFrontView etc).</summary>
    Private Sub LogViewOriginsAndDistances(vBase As DrawingView, vRight As DrawingView, vBelow As DrawingView,
                                           baseOri As Integer, oriRight As Integer, oriBelow As Integer)
        Dim oxBase As Double = 0, oyBase As Double = 0
        Dim oxRight As Double = 0, oyRight As Double = 0
        Dim oxBelow As Double = 0, oyBelow As Double = 0
        If vBase IsNot Nothing Then vBase.GetOrigin(oxBase, oyBase)
        If vRight IsNot Nothing Then vRight.GetOrigin(oxRight, oyRight)
        If vBelow IsNot Nothing Then vBelow.GetOrigin(oxBelow, oyBelow)
        StepLog($"[ORIGINS] Base ({OriToConstantName(baseOri)}):  origen=({oxBase:0.0000}, {oyBase:0.0000})")
        LogViewRange("Base", vBase, baseOri)
        StepLog($"[ORIGINS] Right ({OriToConstantName(oriRight)}): origen=({oxRight:0.0000}, {oyRight:0.0000}) alineada a Base")
        LogViewRange("Right", vRight, oriRight)
        StepLog($"[ORIGINS] Below ({OriToConstantName(oriBelow)}): origen=({oxBelow:0.0000}, {oyBelow:0.0000}) alineada a Base")
        LogViewRange("Below", vBelow, oriBelow)
        Dim dBaseRight As Double = Math.Sqrt((oxRight - oxBase) ^ 2 + (oyRight - oyBase) ^ 2)
        Dim dBaseBelow As Double = Math.Sqrt((oxBelow - oxBase) ^ 2 + (oyBelow - oyBase) ^ 2)
        Dim dRightBelow As Double = Math.Sqrt((oxBelow - oxRight) ^ 2 + (oyBelow - oyRight) ^ 2)
        StepLog($"[ORIGINS] Distancia Base-Right:  {dBaseRight:0.0000}")
        StepLog($"[ORIGINS] Distancia Base-Below: {dBaseBelow:0.0000}")
        StepLog($"[ORIGINS] Distancia Right-Below: {dRightBelow:0.0000}")
    End Sub

    Private Sub LogViewRange(tag As String, v As DrawingView, Optional ori As Integer = -999)
        Dim oriStr As String = If(ori <> -999, $" {OriToConstantName(ori)}", "")
        If v Is Nothing Then
            StepLog($"[RANGE] {tag}{oriStr}: (Nothing)")
            Return
        End If
        Dim xmin As Double, ymin As Double, xmax As Double, ymax As Double
        If TryGetViewRange(v, xmin, ymin, xmax, ymax) Then
            Dim w As Double = xmax - xmin
            Dim h As Double = ymax - ymin
            StepLog($"[RANGE] {tag}{oriStr}: [{xmin:0.0000},{ymin:0.0000}]->[{xmax:0.0000},{ymax:0.0000}] tamaño=({w:0.0000} x {h:0.0000})")
        Else
            StepLog($"[RANGE] {tag}{oriStr}: <<NO>>")
        End If
    End Sub

    Private Function TryGetViewCenter(dv As DrawingView, ByRef cx As Double, ByRef cy As Double) As Boolean
        Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
        If Not TryGetViewRange(dv, x1, y1, x2, y2) Then Return False
        cx = (x1 + x2) / 2.0
        cy = (y1 + y2) / 2.0
        Return True
    End Function

    ''' <summary>Fuerza la orientación de la vista con SetViewOrientationStandard (API Solid Edge).
    ''' Útil cuando AddPartView no aplica correctamente la orientación al cambiar de base (Top/Right en vez de Front).</summary>
    Private Sub ForceViewOrientationStandard(dv As DrawingView, ori As Integer, tag As String)
        If dv Is Nothing Then Return
        Try
            dv.SetViewOrientationStandard(CType(ori, ViewOrientationConstants))
            StepLog($"[ORIENT] {tag}: SetViewOrientationStandard({OriToConstantName(ori)}) OK")
        Catch ex As Exception
            LogEx($"ForceViewOrientationStandard({tag})", ex)
        End Try
    End Sub

    Private Sub InsertStandard3(app As SolidEdgeFramework.Application,
                                sheet As Sheet,
                                modelLink As ModelLink,
                                isSheetMetal As Boolean,
                                scale As Double,
                                baseOri As Integer,
                                oriRight As Integer,
                                oriBelow As Integer,
                                x0 As Double,
                                y0 As Double,
                                ByRef vBase As DrawingView,
                                ByRef v2 As DrawingView,
                                ByRef v3 As DrawingView)
        InsertStandard3Internal(app, sheet, modelLink, isSheetMetal, scale, baseOri, oriRight, oriBelow,
                               x0, y0, LAYOUT_SEED_RIGHT, LAYOUT_SEED_BELOW, vBase, v2, v3)
    End Sub

    ''' <summary>Inserta vista isométrica (igTopFrontRightView). Opcional: posición (xIso,yIso) o auto a la derecha del bloque.
    ''' <paramref name="isoDrawingScale"/>: si es NaN, se usa scale*ISO_FACTOR; si no, escala UNE independiente de las 3 principales.</summary>
    Private Sub InsertIsoView(app As SolidEdgeFramework.Application,
                             sheet As Sheet,
                             modelLink As ModelLink,
                             isSheetMetal As Boolean,
                             scale As Double,
                             m As Margins,
                             vBase As DrawingView, vRight As DrawingView, vBelow As DrawingView,
                             ByRef vIso As DrawingView,
                             Optional xIsoOverride As Double = Double.NaN,
                             Optional yIsoOverride As Double = Double.NaN,
                             Optional isoDrawingScale As Double = Double.NaN)
        vIso = Nothing
        Dim xIso As Double
        Dim yIso As Double
        If Double.IsNaN(xIsoOverride) OrElse Double.IsNaN(yIsoOverride) Then
            Dim xminB As Double, yminB As Double, xmaxB As Double, ymaxB As Double
            If Not GetCombinedRange3Views(vBase, vRight, vBelow, xminB, yminB, xmaxB, ymaxB) Then Return
            xIso = xmaxB + INSERT_GAP_ISO
            Dim cy As Double = (yminB + ymaxB) / 2.0
            yIso = cy
        Else
            xIso = xIsoOverride
            yIso = yIsoOverride
        End If
        Dim effIso As Double = If(Double.IsNaN(isoDrawingScale), scale * ISO_FACTOR, isoDrawingScale)
        Dim dvws As DrawingViews = sheet.DrawingViews
        Try
            If Not isSheetMetal Then
                Dim vt As PartDrawingViewTypeConstants = PartDrawingViewTypeConstants.sePartDesignedView
                vIso = dvws.AddPartView(modelLink, CInt(ViewOrientationConstants.igTopFrontRightView), effIso, xIso, yIso, vt)
            Else
                Dim vt As SheetMetalDrawingViewTypeConstants = SheetMetalDrawingViewTypeConstants.seSheetMetalDesignedView
                vIso = dvws.AddSheetMetalView(modelLink, CInt(ViewOrientationConstants.igTopFrontRightView), effIso, xIso, yIso, vt)
            End If
            If vIso IsNot Nothing Then SafeUpdateView(vIso, "Iso.Update")
            StepLog($"[INSERT-ISO] igTopFrontRightView pos=({xIso:0.000},{yIso:0.000}) escala_iso={effIso:0.000} escala_principal={scale:0.000}")
        Catch ex As Exception
            LogEx("InsertIsoView", ex)
        End Try
    End Sub

    Private Sub InsertStandard3Internal(app As SolidEdgeFramework.Application,
                                        sheet As Sheet,
                                        modelLink As ModelLink,
                                        isSheetMetal As Boolean,
                                        scale As Double,
                                        baseOri As Integer,
                                        oriRight As Integer,
                                        oriBelow As Integer,
                                        x0 As Double,
                                        y0 As Double,
                                        gapRight As Double,
                                        gapBelow As Double,
                                        ByRef vBase As DrawingView,
                                        ByRef v2 As DrawingView,
                                        ByRef v3 As DrawingView)
        Dim dvws As DrawingViews = sheet.DrawingViews
        Dim xBase = x0
        Dim yBase = y0
        Dim xRightPos = x0 + gapRight
        Dim yRightPos = y0
        Dim xBelowPos = x0
        Dim yBelowPos = y0 - gapBelow
        StepLog($"[INSERT3] Pos: Base({xBase:0.000},{yBase:0.000}) Right({xRightPos:0.000},{yRightPos:0.000}) Below({xBelowPos:0.000},{yBelowPos:0.000}) | ori: base={OriToConstantName(baseOri)} right={OriToConstantName(oriRight)} below={OriToConstantName(oriBelow)}")

        If Not isSheetMetal Then
            Dim vt = PartDrawingViewTypeConstants.sePartDesignedView
            vBase = dvws.AddPartView(modelLink, baseOri, scale, xBase, yBase, vt)
            ForceViewOrientationStandard(vBase, baseOri, "Base")
            DoIdleSafe(app, "Insert3 Base")
            v2 = dvws.AddPartView(modelLink, oriRight, scale, xRightPos, yRightPos, vt)
            ForceViewOrientationStandard(v2, oriRight, "Right")
            DoIdleSafe(app, "Insert3 VRight")
            v3 = dvws.AddPartView(modelLink, oriBelow, scale, xBelowPos, yBelowPos, vt)
            ForceViewOrientationStandard(v3, oriBelow, "Below")
            DoIdleSafe(app, "Insert3 VBelow")
        Else
            Dim vt = SheetMetalDrawingViewTypeConstants.seSheetMetalDesignedView
            vBase = dvws.AddSheetMetalView(modelLink, baseOri, scale, xBase, yBase, vt)
            ForceViewOrientationStandard(vBase, baseOri, "Base")
            DoIdleSafe(app, "Insert3 Base")
            v2 = dvws.AddSheetMetalView(modelLink, oriRight, scale, xRightPos, yRightPos, vt)
            ForceViewOrientationStandard(v2, oriRight, "Right")
            DoIdleSafe(app, "Insert3 VRight")
            v3 = dvws.AddSheetMetalView(modelLink, oriBelow, scale, xBelowPos, yBelowPos, vt)
            ForceViewOrientationStandard(v3, oriBelow, "Below")
            DoIdleSafe(app, "Insert3 VBelow")
        End If
    End Sub

    Private Function GetTemplateUsableAreas(templates As String()) As List(Of TemplateUsableArea)
        Dim list As New List(Of TemplateUsableArea)
        For Each tpl In templates
            If String.IsNullOrWhiteSpace(tpl) OrElse Not File.Exists(tpl) Then
                Continue For
            End If
            Dim name = Path.GetFileNameWithoutExtension(tpl).ToLowerInvariant()
            Dim m = GetMarginsByTemplate(tpl)
            Dim sw As Double = 0.297
            Dim sh As Double = 0.42
            If name.Contains("a4") Then
                sw = 0.21
                sh = 0.297
            ElseIf name.Contains("a2") Then
                sw = 0.42
                sh = 0.594
            ElseIf name.Contains("a3") Then
                sw = 0.297
                sh = 0.42
            ElseIf name.Contains("a1") Then
                sw = 0.594
                sh = 0.841
            End If
            Dim uw = sw - m.Left - m.Right
            Dim uh = sh - m.Top - m.Bottom
            list.Add(New TemplateUsableArea With {.Name = Path.GetFileName(tpl), .TemplatePath = tpl, .UsableW = uw, .UsableH = uh})
        Next
        Return list.OrderBy(Function(x) x.UsableW * x.UsableH).ToList()
    End Function

    Private Function PickStandardScale(rangeH As Double, rangeV As Double, usableW As Double, usableH As Double) As Double
        If rangeH <= 0 OrElse rangeV <= 0 Then Return 1.0
        Dim scaleMax = Math.Min(usableW / rangeH, usableH / rangeV)
        For i As Integer = 0 To StandardScales.Length - 1
            If StandardScales(i) <= scaleMax + EPS Then Return StandardScales(i)
        Next
        Return StandardScales(StandardScales.Length - 1)
    End Function

    Private Sub PickFormatAndScale(rangeH As Double, rangeV As Double,
                                   templateAreas As List(Of TemplateUsableArea),
                                   ByRef chosenTemplate As String,
                                   ByRef chosenScale As Double)
        chosenTemplate = Nothing
        chosenScale = 0
        For Each t In templateAreas
            If t.UsableW <= 0 OrElse t.UsableH <= 0 Then Continue For
            Dim sc = PickStandardScale(rangeH, rangeV, t.UsableW, t.UsableH)
            If sc > 0 AndAlso rangeH * sc <= t.UsableW + EPS AndAlso rangeV * sc <= t.UsableH + EPS Then
                chosenTemplate = t.TemplatePath
                chosenScale = sc
                Return
            End If
        Next
    End Sub

    ' Tercer diedro (sistema americano): vista a la derecha y vista debajo de la base.
    ' Front base: derecha = Right, debajo = Top.
    ' Top base: derecha = Right, debajo = Front.
    ' Right base: derecha = Front, debajo = Top.
    ' (Solid Edge suele usar tercer diedro por defecto en plantillas)
    Private Sub GetOrisForBaseAmerican(baseOri As Integer, ByRef oriRight As Integer, ByRef oriBelow As Integer)
        Select Case CType(baseOri, ViewOrientationConstants)
            Case ViewOrientationConstants.igTopView
                oriRight = CInt(ViewOrientationConstants.igRightView)
                oriBelow = CInt(ViewOrientationConstants.igFrontView)
            Case ViewOrientationConstants.igFrontView
                oriRight = CInt(ViewOrientationConstants.igRightView)
                oriBelow = CInt(ViewOrientationConstants.igTopView)
            Case ViewOrientationConstants.igRightView
                oriRight = CInt(ViewOrientationConstants.igFrontView)
                oriBelow = CInt(ViewOrientationConstants.igTopView)
            Case ViewOrientationConstants.igBackView
                oriRight = CInt(ViewOrientationConstants.igLeftView)
                oriBelow = CInt(ViewOrientationConstants.igTopView)
            Case ViewOrientationConstants.igLeftView
                oriRight = CInt(ViewOrientationConstants.igFrontView)
                oriBelow = CInt(ViewOrientationConstants.igTopView)
            Case ViewOrientationConstants.igBottomView
                oriRight = CInt(ViewOrientationConstants.igLeftView)
                oriBelow = CInt(ViewOrientationConstants.igFrontView)
            Case Else
                oriRight = CInt(ViewOrientationConstants.igRightView)
                oriBelow = CInt(ViewOrientationConstants.igTopView)
        End Select
    End Sub

    ' Rotación de vistas auxiliares cuando base<>Front. Sin tocar SCoords de la pieza.
    ' SOLO UNA de las dos vistas (derecha o debajo) necesita rotación; cuál depende de la base.
    ' base=Top: solo Derecha(Left) rota -90º; Below(Back) 0º.
    ' Para otras bases: cuál rota debe determinarse por prueba (derecha vs debajo).
    Private Function GetAuxViewRotationRadians(baseOri As Integer, auxOri As Integer, isRightView As Boolean) As Double
        If baseOri = CInt(ViewOrientationConstants.igFrontView) Then Return 0.0
        Select Case CType(baseOri, ViewOrientationConstants)
            Case ViewOrientationConstants.igTopView
                ' Solo la derecha (Left) rota -90º; debajo (Back) no
                If isRightView Then Return NEG_90_RAD Else Return 0.0
            Case ViewOrientationConstants.igBottomView
                ' Solo una rota: prueba derecha
                If isRightView Then Return RIGHT_ROT_RAD Else Return 0.0
            Case ViewOrientationConstants.igRightView
                ' Solo una rota: prueba debajo
                If isRightView Then Return 0.0 Else Return NEG_90_RAD
            Case ViewOrientationConstants.igLeftView
                ' Solo una rota: prueba debajo
                If isRightView Then Return 0.0 Else Return RIGHT_ROT_RAD
            Case ViewOrientationConstants.igBackView
                ' Solo una rota: prueba derecha
                If isRightView Then Return NEG_90_RAD Else Return 0.0
            Case Else
                Return 0.0
        End Select
    End Function

    ' Primer diedro - Sistema Europeo (ISO/UNE) - First Angle Projection.
    ' Regla básica: la vista se coloca en el lado CONTRARIO al que se mira.
    '   Derecha→se dibuja a la izquierda | Izquierda→se dibuja a la derecha
    '   Superior→se dibuja debajo        | Inferior→se dibuja encima
    ' Tabla según norma europea:
    '   Vista base | Vista derecha (se coloca) | Vista inferior (se coloca)
    '   Front      | Left                      | Top
    '   Top        | Left                      | Back
    '   Bottom     | Right                     | Front
    '   Right      | Front                     | Top
    '   Left       | Back                      | Top
    '   Back       | Right                     | Top
    Private Sub GetOrisForBaseEuropean(baseOri As Integer, ByRef oriRight As Integer, ByRef oriBelow As Integer)
        Select Case CType(baseOri, ViewOrientationConstants)
            Case ViewOrientationConstants.igFrontView
                oriRight = CInt(ViewOrientationConstants.igLeftView)
                oriBelow = CInt(ViewOrientationConstants.igTopView)
            Case ViewOrientationConstants.igTopView
                oriRight = CInt(ViewOrientationConstants.igLeftView)
                oriBelow = CInt(ViewOrientationConstants.igBackView)
            Case ViewOrientationConstants.igBottomView
                oriRight = CInt(ViewOrientationConstants.igRightView)
                oriBelow = CInt(ViewOrientationConstants.igFrontView)
            Case ViewOrientationConstants.igRightView
                oriRight = CInt(ViewOrientationConstants.igFrontView)
                oriBelow = CInt(ViewOrientationConstants.igTopView)
            Case ViewOrientationConstants.igLeftView
                oriRight = CInt(ViewOrientationConstants.igBackView)
                oriBelow = CInt(ViewOrientationConstants.igTopView)
            Case ViewOrientationConstants.igBackView
                oriRight = CInt(ViewOrientationConstants.igRightView)
                oriBelow = CInt(ViewOrientationConstants.igTopView)
            Case Else
                oriRight = CInt(ViewOrientationConstants.igLeftView)
                oriBelow = CInt(ViewOrientationConstants.igTopView)
        End Select
    End Sub

    Private Sub InsertStandard4(app As SolidEdgeFramework.Application,
                            sheet As Sheet,
                            modelLink As ModelLink,
                            isSheetMetal As Boolean,
                            scale As Double,
                            baseOri As Integer,
                            baseRotated As Boolean,
                            m As Margins,
                            ByRef vFront As DrawingView,
                            ByRef vTop As DrawingView,
                            ByRef vRight As DrawingView,
                            ByRef vIso As DrawingView)

        Dim dvws As DrawingViews = sheet.DrawingViews
        Dim W As Double = sheet.SheetSetup.SheetWidth
        Dim H As Double = sheet.SheetSetup.SheetHeight

        Dim xSafe As Double = m.Left + INSERT_OFFSET_FROM_MARGIN
        Dim ySafe As Double = (H - m.Top) - INSERT_OFFSET_FROM_MARGIN

        Dim xFront As Double = xSafe
        Dim yFront As Double = ySafe

        Dim xBelow As Double = xSafe
        Dim yBelow As Double = ySafe - INSERT_GAP_BELOW

        Dim xRight As Double = xSafe + INSERT_GAP_RIGHT
        Dim yRight As Double = ySafe

        Dim xIso As Double = xSafe + INSERT_GAP_ISO
        Dim yIso As Double = ySafe - INSERT_GAP_BELOW

        Dim oriRight As Integer, oriBelow As Integer
        ' Sistema europeo (primer diedro) - consistente con CreateDraftAlzadoPrimerDiedro
        GetOrisForBaseEuropean(baseOri, oriRight, oriBelow)

        ' Si base rotada, usar mapa rotado para vistas proyectadas
        If baseRotated Then
            Dim mapRot As ProjectedViewMap = GetProjectedViewMap(SolidEdgeToOrthoView(baseOri), ViewRotation.RotMinus90)
            oriRight = OrthoViewToSolidEdge(mapRot.Right)
            oriBelow = OrthoViewToSolidEdge(mapRot.Down)
        End If

        Dim posRightOri As Integer = oriRight
        Dim posBelowOri As Integer = oriBelow

        StepLog($"InsertStandard4: W={W} H={H} baseRotated={baseRotated} European")
        StepLog($"Insert pos: Base({xFront},{yFront}) DEBAJO({xBelow},{yBelow}) DERECHA({xRight},{yRight}) Iso({xIso},{yIso}) scale={scale}")

        If Not isSheetMetal Then
            Dim vt As PartDrawingViewTypeConstants = PartDrawingViewTypeConstants.sePartDesignedView

            vFront = dvws.AddPartView(modelLink, baseOri, scale, xFront, yFront, vt)
            DoIdleSafe(app, "InsertStandard4 after Base")
            vTop = dvws.AddPartView(modelLink, posBelowOri, scale, xBelow, yBelow, vt)
            DoIdleSafe(app, "InsertStandard4 view below")
            vRight = dvws.AddPartView(modelLink, posRightOri, scale, xRight, yRight, vt)
            DoIdleSafe(app, "InsertStandard4 view right")
            vIso = dvws.AddPartView(modelLink,
                                CInt(ViewOrientationConstants.igTopFrontLeftView),
                                scale * ISO_FACTOR,
                                xIso, yIso, vt)
            DoIdleSafe(app, "InsertStandard4 after Iso")
        Else
            Dim vt As SheetMetalDrawingViewTypeConstants = SheetMetalDrawingViewTypeConstants.seSheetMetalDesignedView

            vFront = dvws.AddSheetMetalView(modelLink, baseOri, scale, xFront, yFront, vt)
            DoIdleSafe(app, "InsertStandard4 after Base")
            vTop = dvws.AddSheetMetalView(modelLink, posBelowOri, scale, xBelow, yBelow, vt)
            DoIdleSafe(app, "InsertStandard4 view below")
            vRight = dvws.AddSheetMetalView(modelLink, posRightOri, scale, xRight, yRight, vt)
            DoIdleSafe(app, "InsertStandard4 view right")
            vIso = dvws.AddSheetMetalView(modelLink,
                                      CInt(ViewOrientationConstants.igTopFrontLeftView),
                                      scale * ISO_FACTOR,
                                      xIso, yIso, vt)
            DoIdleSafe(app, "InsertStandard4 after Iso")
        End If

    End Sub

    Private Function GetOriginY(v As DrawingView) As Double
        Dim ox As Double = 0, oy As Double = 0
        v.GetOrigin(ox, oy)
        Return oy
    End Function

    Private Sub SetOriginY(app As SolidEdgeFramework.Application, v As DrawingView, targetY As Double, tag As String)
        If v Is Nothing Then Return
        Dim ox As Double = 0, oy As Double = 0
        v.GetOrigin(ox, oy)
        v.SetOrigin(ox, targetY)
        StepLog($"{tag}: SetOriginY -> ({ox},{oy}) => ({ox},{targetY})")
        DoIdleSafe(app, $"after SetOriginY {tag}")
    End Sub

    ''' <summary>Girar vista derecha -90º cuando es muy alargada (ancha O alta) para mejor layout.</summary>
    Private Function ShouldRotateRight(w As Double, h As Double) As Boolean
        Const RATIO As Double = 1.15
        If w <= EPS OrElse h <= EPS Then Return False
        ' Muy ancha (w>=1.15*h) O muy alta (h>=1.15*w): girar -90º para normalizar
        Return (w / h) >= RATIO OrElse (h / w) >= RATIO
    End Function

    Private Function GetViewCenterY(dv As DrawingView) As Double
        Dim xmin As Double, ymin As Double, xmax As Double, ymax As Double
        If Not TryGetViewRange(dv, xmin, ymin, xmax, ymax) Then Return 0
        Return (ymin + ymax) / 2.0
    End Function

    Private Sub EnsureFlatPatternReady(app As SolidEdgeFramework.Application, modelPath As String)
        If Not modelPath.EndsWith(".psm", StringComparison.OrdinalIgnoreCase) Then Exit Sub

        Dim doc As Object = Nothing
        Try
            StepLog("[FLAT-PRE] Abriendo PSM para preparar FlatPattern...")
            doc = app.Documents.Open(modelPath)

            ' Intento forzar cálculo interno
            Try
                Dim sm = TryCast(doc, SolidEdgePart.SheetMetalDocument)
                If sm IsNot Nothing Then
                    Dim cnt As Integer = 0
                    Try : cnt = sm.FlatPatternModels.Count : Catch : End Try
                    StepLog($"[FLAT-PRE] FlatPatternModels.Count = {cnt}")
                Else
                    StepLog("[FLAT-PRE] WARN: No pude castear a SheetMetalDocument.")
                End If
            Catch ex As Exception
                StepLog("[FLAT-PRE] WARN: " & ex.Message)
            End Try

            app.DoIdle()
            Try : doc.UpdateAll() : Catch : End Try
            app.DoIdle()

        Catch ex As Exception
            LogEx("[FLAT-PRE] Error preparando PSM", ex)
        Finally
            Try
                If doc IsNot Nothing Then CType(doc, Object).Close(False)
            Catch
            End Try
            app.DoIdle()
        End Try
    End Sub

    Private Sub AlignViewsCenterY_MoveOnly(app As SolidEdgeFramework.Application,
                                           dvRef As DrawingView,
                                           dvToMove As DrawingView,
                                           tag As String)

        If dvRef Is Nothing OrElse dvToMove Is Nothing Then Exit Sub

        Dim cyRef As Double = GetViewCenterY(dvRef)
        Dim cyMov As Double = GetViewCenterY(dvToMove)

        Dim dy As Double = cyRef - cyMov

        Dim ox As Double = 0, oy As Double = 0
        dvToMove.GetOrigin(ox, oy)
        dvToMove.SetOrigin(ox, oy + dy)

        DoIdleSafe(app, $"after AlignViewsCenterY_MoveOnly {tag}")
    End Sub

    Private Sub ForceSameOriginY(app As SolidEdgeFramework.Application,
                                 ref As SolidEdgeDraft.DrawingView,
                                 mov As SolidEdgeDraft.DrawingView,
                                 tag As String)

        If ref Is Nothing OrElse mov Is Nothing Then Exit Sub

        Try
            Dim rx As Double = 0, ry As Double = 0
            Dim mx As Double = 0, my As Double = 0

            ref.GetOrigin(rx, ry)
            mov.GetOrigin(mx, my)

            mov.SetOrigin(mx, ry)

            DoIdleSafe(app, $"after ForceSameOriginY {tag}")

            Try
                mov.Update()
            Catch
            End Try

        Catch ex As Exception
            LogEx($"ForceSameOriginY({tag})", ex)
        End Try
    End Sub

    Private Sub DumpOneView(tag As String, v As DrawingView)
        If v Is Nothing Then
            StepLog($"{tag}: (Nothing)")
            Return
        End If

        Try
            Dim ox As Double = 0, oy As Double = 0
            v.GetOrigin(ox, oy)

            Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
            If Not TryGetViewRange(v, x1, y1, x2, y2) Then
                StepLog($"{tag}: ORIGIN=({ox},{oy}) RANGE=<<NO>>")
                Return
            End If

            StepLog($"{tag}: ORIGIN=({ox},{oy}) RANGE=[{x1},{y1}]->[{x2},{y2}] SIZE=({x2 - x1} x {y2 - y1})")
        Catch ex As Exception
            LogEx($"DumpOneView({tag})", ex)
        End Try
    End Sub

    Private Sub CloseDraftNoSave(dft As DraftDocument)
        Try
            dft.Close(False)
        Catch
        End Try
    End Sub

    Private Sub DoIdleSafe(app As SolidEdgeFramework.Application, tag As String)
        Try
            app.DoIdle()
            StepLog($"DoIdle: OK ({tag})")
        Catch ex As Exception
            LogEx($"DoIdle({tag})", ex)
        End Try
    End Sub

    Private Sub GetSheetSize(sheet As Sheet, ByRef w As Double, ByRef h As Double)
        Dim ss As SheetSetup = sheet.SheetSetup
        w = ss.SheetWidth
        h = ss.SheetHeight
        StepLog($"SHEET size = ({w} x {h})")
    End Sub

    ''' <summary>Márgenes del área útil (valores originales que funcionaban).</summary>
    Private Function GetMarginsByTemplate(templatePath As String) As Margins
        Dim m As New Margins

        If templatePath.EndsWith("A4 Plantilla.dft", StringComparison.OrdinalIgnoreCase) OrElse
           templatePath.EndsWith("A4 Plantilla.dxf", StringComparison.OrdinalIgnoreCase) Then
            m.Left = 0.015 : m.Right = 0.01 : m.Top = 0.01 : m.Bottom = 0.03
        ElseIf templatePath.EndsWith("A3 Plantilla.dft", StringComparison.OrdinalIgnoreCase) Then
            m.Left = 0.02 : m.Right = 0.01 : m.Top = 0.01 : m.Bottom = 0.04
        ElseIf templatePath.EndsWith("A2 Plantilla.dft", StringComparison.OrdinalIgnoreCase) Then
            m.Left = 0.02 : m.Right = 0.01 : m.Top = 0.01 : m.Bottom = 0.045
        Else
            m.Left = 0.02 : m.Right = 0.01 : m.Top = 0.01 : m.Bottom = 0.04
        End If

        Return m
    End Function

    'Private Function GetBaseOriCandidates() As Integer()
    '    ' 6 vistas del cubo (las “planas”)
    '    Return New Integer() {
    '    CInt(ViewOrientationConstants.igFrontView),
    '    CInt(ViewOrientationConstants.igTopView),
    '    CInt(ViewOrientationConstants.igRightView),
    '    CInt(ViewOrientationConstants.igBackView),
    '    CInt(ViewOrientationConstants.igBottomView),
    '    CInt(ViewOrientationConstants.igLeftView)
    '}
    'End Function

    Private Sub LogSheetWindowOrigin(app As SolidEdgeFramework.Application, dft As SolidEdgeDraft.DraftDocument, tag As String)
        Try
            dft.Activate()
            app.DoIdle()

            Dim winObj As Object = app.ActiveWindow
            Dim sw As SolidEdgeDraft.SheetWindow = TryCast(winObj, SolidEdgeDraft.SheetWindow)

            If sw Is Nothing Then
                StepLog($"{tag}: ActiveWindow no es SheetWindow.")
                Return
            End If

            Dim ox As Double = 0, oy As Double = 0
            sw.GetOrigin(ox, oy)

            StepLog($"{tag}: SheetWindow.GetOrigin -> ({ox}, {oy})")

        Catch ex As Exception
            LogEx($"LogSheetWindowOrigin({tag})", ex)
        End Try
    End Sub

    '=========================================================
    ' CANDIDATAS DE VISTA BASE
    '=========================================================
    Private Structure BaseViewCandidate
        Public Orientation As Integer
        Public OrientationName As String

        Public Rotated As Boolean
        Public RotationRad As Double

        Public Width As Double
        Public Height As Double
        Public Area As Double
        Public Aspect As Double   ' width / height

        Public FitW As Double
        Public FitH As Double
        Public AreaUse As Double

        Public Score As Double
    End Structure

    Private Function GetBaseOriCandidates() As Integer()
        Return New Integer() {
            CInt(ViewOrientationConstants.igFrontView),
            CInt(ViewOrientationConstants.igTopView),
            CInt(ViewOrientationConstants.igRightView)
                }
    End Function

    Private Function OriToText(ori As Integer) As String
        Select Case CType(ori, ViewOrientationConstants)
            Case ViewOrientationConstants.igFrontView : Return "Front"
            Case ViewOrientationConstants.igBackView : Return "Back"
            Case ViewOrientationConstants.igTopView : Return "Top"
            Case ViewOrientationConstants.igBottomView : Return "Bottom"
            Case ViewOrientationConstants.igRightView : Return "Right"
            Case ViewOrientationConstants.igLeftView : Return "Left"
            Case Else : Return $"Ori({ori})"
        End Select
    End Function

    ''' <summary>Nombre de la constante (ej: igFrontView) para log.</summary>
    Private Function OriToConstantName(ori As Integer) As String
        Select Case CType(ori, ViewOrientationConstants)
            Case ViewOrientationConstants.igFrontView : Return "igFrontView"
            Case ViewOrientationConstants.igBackView : Return "igBackView"
            Case ViewOrientationConstants.igTopView : Return "igTopView"
            Case ViewOrientationConstants.igBottomView : Return "igBottomView"
            Case ViewOrientationConstants.igRightView : Return "igRightView"
            Case ViewOrientationConstants.igLeftView : Return "igLeftView"
            Case Else : Return $"Ori({ori})"
        End Select
    End Function

    Private Function ScoreBaseCandidate(viewW As Double,
                                    viewH As Double,
                                    usableW As Double,
                                    usableH As Double) As Double

        If viewW <= EPS OrElse viewH <= EPS Then Return -999999.0
        If usableW <= EPS OrElse usableH <= EPS Then Return -999999.0

        Dim fitW As Double = viewW / usableW
        Dim fitH As Double = viewH / usableH
        Dim areaUse As Double = (viewW * viewH) / (usableW * usableH)

        ' Si se sale de hoja, muy penalizada
        If viewW > usableW + EPS OrElse viewH > usableH + EPS Then
            Return -1000.0 - Math.Abs(viewW - usableW) * 100.0 - Math.Abs(viewH - usableH) * 100.0
        End If

        ' Penalizar vistas “de testa” demasiado pequeñas
        ' Si no ocupan al menos un 35% del ancho útil, suelen ser mala vista principal
        Dim tooSmallPenalty As Double = 0.0
        If fitW < 0.35 Then
            tooSmallPenalty = (0.35 - fitW) * 2.0
        End If

        ' Bonus moderado si es más horizontal que vertical
        Dim aspectBonus As Double = 0.0
        If viewH > EPS Then
            Dim aspect = viewW / viewH
            If aspect >= 1.0 Then
                aspectBonus = Math.Min(aspect, 4.0) / 4.0
            End If
        End If

        Return (areaUse * 0.55) + (fitW * 0.25) + (fitH * 0.1) + (aspectBonus * 0.1) - tooSmallPenalty
    End Function

    Private Function BuildCandidateFromView(ori As Integer,
                                            rotated As Boolean,
                                            rotationRad As Double,
                                            dv As DrawingView,
                                            usableW As Double,
                                            usableH As Double) As BaseViewCandidate

        Dim c As New BaseViewCandidate

        Dim w As Double = 0, h As Double = 0
        GetViewSizeSmart(dv, w, h)

        c.Orientation = ori
        c.OrientationName = OriToText(ori)
        c.Rotated = rotated
        c.RotationRad = rotationRad

        c.Width = w
        c.Height = h
        c.Area = w * h

        If h > EPS Then
            c.Aspect = w / h
        Else
            c.Aspect = 0
        End If

        If usableW > EPS Then c.FitW = w / usableW Else c.FitW = 0
        If usableH > EPS Then c.FitH = h / usableH Else c.FitH = 0
        If usableW > EPS AndAlso usableH > EPS Then
            c.AreaUse = (w * h) / (usableW * usableH)
        Else
            c.AreaUse = 0
        End If

        c.Score = ScoreBaseCandidate(w, h, usableW, usableH)

        Return c
    End Function

    Private Function ProbeSingleBaseOrientation(app As SolidEdgeFramework.Application,
                                                modelPath As String,
                                                templatePath As String,
                                                ori As Integer,
                                                scale As Double,
                                                isSheetMetal As Boolean,
                                                usableW As Double,
                                                usableH As Double) As BaseViewCandidate

        Dim best As New BaseViewCandidate
        best.Score = -999999.0
        best.Orientation = ori
        best.OrientationName = OriToText(ori)

        Dim dft As DraftDocument = Nothing

        Try
            StepLog($"[PROBE-BASE] Ori={OriToText(ori)} scale={scale}")

            dft = CType(app.Documents.Add("SolidEdge.DraftDocument", templatePath), DraftDocument)
            DoIdleSafe(app, $"probe draft add {OriToText(ori)}")

            Dim sheet As Sheet = dft.ActiveSheet
            Dim modelLink As ModelLink = dft.ModelLinks.Add(modelPath)
            DoIdleSafe(app, $"probe model link {OriToText(ori)}")

            Dim dvws As DrawingViews = sheet.DrawingViews
            Dim dv As DrawingView = Nothing

            Dim x As Double = 0.12
            Dim y As Double = 0.2

            If Not isSheetMetal Then
                Dim vt As PartDrawingViewTypeConstants = PartDrawingViewTypeConstants.sePartDesignedView
                dv = dvws.AddPartView(modelLink, ori, scale, x, y, vt)
            Else
                Dim vt As SheetMetalDrawingViewTypeConstants = SheetMetalDrawingViewTypeConstants.seSheetMetalDesignedView
                dv = dvws.AddSheetMetalView(modelLink, ori, scale, x, y, vt)
            End If

            DoIdleSafe(app, $"probe insert base {OriToText(ori)}")
            SafeUpdateView(dv, $"Probe.Update base {OriToText(ori)}")
            DoIdleSafe(app, $"probe update base {OriToText(ori)}")

            Dim normalCand As BaseViewCandidate =
                BuildCandidateFromView(ori, False, 0.0, dv, usableW, usableH)

            StepLog($"[PROBE-BASE] {normalCand.OrientationName} normal -> W={normalCand.Width:0.000000} H={normalCand.Height:0.000000} Score={normalCand.Score:0.000000}")

            Dim rotatedCand As BaseViewCandidate
            rotatedCand.Score = -999999.0
            rotatedCand.Orientation = ori
            rotatedCand.OrientationName = OriToText(ori)

            Try
                SafeSetRotationAngle(dv, -RIGHT_ROT_RAD)
                DoIdleSafe(app, $"probe rotate base {OriToText(ori)}")
                SafeUpdateView(dv, $"Probe.Update rotated {OriToText(ori)}")
                DoIdleSafe(app, $"probe update rotated {OriToText(ori)}")

                rotatedCand = BuildCandidateFromView(ori, True, -RIGHT_ROT_RAD, dv, usableW, usableH)
                StepLog($"[PROBE-BASE] {rotatedCand.OrientationName} rotated -> W={rotatedCand.Width:0.000000} H={rotatedCand.Height:0.000000} Score={rotatedCand.Score:0.000000}")

            Catch ex As Exception
                LogEx($"Probe rotate {OriToText(ori)}", ex)
            End Try

            If rotatedCand.Score > normalCand.Score Then
                best = rotatedCand
                StepLog($"[PROBE-BASE] BEST for {OriToText(ori)} = ROTATED")
            Else
                best = normalCand
                StepLog($"[PROBE-BASE] BEST for {OriToText(ori)} = NORMAL")
            End If

        Catch ex As Exception
            LogEx($"ProbeSingleBaseOrientation({OriToText(ori)})", ex)
        Finally
            If dft IsNot Nothing Then
                Try
                    dft.Close(False)
                Catch
                End Try
            End If
        End Try

        Return best
    End Function

    Private Function GetBestBaseCandidates(app As SolidEdgeFramework.Application,
                                           modelPath As String,
                                           templatePath As String,
                                           scale As Double,
                                           isSheetMetal As Boolean,
                                           usableW As Double,
                                           usableH As Double) As List(Of BaseViewCandidate)

        Dim result As New List(Of BaseViewCandidate)
        Dim oris = GetBaseOriCandidates()

        StepLog("==============================================")
        StepLog("[PROBE-BASE] INICIO análisis de orientaciones base")
        StepLog("==============================================")

        For Each ori In oris
            Dim cand = ProbeSingleBaseOrientation(app, modelPath, templatePath, ori, scale, isSheetMetal, usableW, usableH)
            result.Add(cand)
        Next

        result = result.
            OrderByDescending(Function(c) c.Score).
            ThenByDescending(Function(c) c.Area).
            ToList()

        StepLog("==============================================")
        StepLog("[PROBE-BASE] RANKING FINAL")
        StepLog("==============================================")

        For i As Integer = 0 To result.Count - 1
            Dim c = result(i)
            StepLog($"#{i + 1} Ori={c.OrientationName} Rot={c.Rotated} W={c.Width:0.000000} H={c.Height:0.000000} Aspect={c.Aspect:0.000000} Score={c.Score:0.000000}")
        Next

        Return result
    End Function

    Private Sub LogBaseViewsSizeOnly(app As SolidEdgeFramework.Application,
                                 modelPath As String,
                                 templatePath As String)

        StepLog("==============================================")
        StepLog("[SIZE-ONLY] INICIO lectura de tamaños base")
        StepLog("==============================================")

        If Not File.Exists(modelPath) Then
            StepLog("[SIZE-ONLY] ERROR: no existe modelPath.")
            Exit Sub
        End If

        If Not File.Exists(templatePath) Then
            StepLog("[SIZE-ONLY] ERROR: no existe templatePath.")
            Exit Sub
        End If

        Dim isSheetMetal As Boolean = modelPath.EndsWith(".psm", StringComparison.OrdinalIgnoreCase)

        Dim oriNames As New Dictionary(Of Integer, String) From {
        {CInt(ViewOrientationConstants.igFrontView), "Front"},
        {CInt(ViewOrientationConstants.igTopView), "Top"},
        {CInt(ViewOrientationConstants.igRightView), "Right"}
    }

        Dim baseOris As Integer() = {
        CInt(ViewOrientationConstants.igFrontView),
        CInt(ViewOrientationConstants.igTopView),
        CInt(ViewOrientationConstants.igRightView)
    }

        For Each ori In baseOris
            Dim probeDoc As DraftDocument = Nothing

            Try
                StepLog($"[SIZE-ONLY] Ori={oriNames(ori)}")

                probeDoc = CType(app.Documents.Add("SolidEdge.DraftDocument", templatePath), DraftDocument)
                DoIdleSafe(app, $"size-only Draft Add {oriNames(ori)}")

                Dim sheet As Sheet = probeDoc.ActiveSheet
                Dim modelLink As ModelLink = probeDoc.ModelLinks.Add(modelPath)
                DoIdleSafe(app, $"size-only ModelLink {oriNames(ori)}")

                Dim dv As DrawingView = Nothing
                Dim x As Double = 0.15
                Dim y As Double = 0.2
                Dim scale As Double = 1.0

                If Not isSheetMetal Then
                    dv = sheet.DrawingViews.AddPartView(
                    modelLink,
                    ori,
                    scale,
                    x,
                    y,
                    PartDrawingViewTypeConstants.sePartDesignedView
                )
                Else
                    dv = sheet.DrawingViews.AddSheetMetalView(
                    modelLink,
                    ori,
                    scale,
                    x,
                    y,
                    SheetMetalDrawingViewTypeConstants.seSheetMetalDesignedView
                )
                End If

                DoIdleSafe(app, $"size-only InsertView {oriNames(ori)}")
                SafeUpdateView(dv, $"SizeOnly.Update {oriNames(ori)}")
                DoIdleSafe(app, $"size-only Update {oriNames(ori)}")

                Dim w As Double = 0
                Dim h As Double = 0
                GetViewSizeSmart(dv, w, h)

                StepLog($"[SIZE-ONLY] {oriNames(ori)} -> W={w:0.000000}  H={h:0.000000}")

            Catch ex As Exception
                LogEx($"[SIZE-ONLY] {oriNames(ori)}", ex)
            Finally
                If probeDoc IsNot Nothing Then
                    Try
                        probeDoc.Close(False)
                    Catch
                    End Try
                End If
            End Try
        Next

        StepLog("==============================================")
        StepLog("[SIZE-ONLY] FIN lectura de tamaños base")
        StepLog("==============================================")
    End Sub

    ''' <summary>Obtiene los tamaños de Front/Top/Right a escala 1. Devuelve Nothing si falla.</summary>
    Private Function GetBaseViewSizesAtScale1(app As SolidEdgeFramework.Application,
                                              modelPath As String,
                                              templatePath As String,
                                              isSheetMetal As Boolean) As BaseViewSizesAtScale1?
        If Not File.Exists(modelPath) OrElse Not File.Exists(templatePath) Then Return Nothing
        Dim r As New BaseViewSizesAtScale1
        Dim oriNames As New Dictionary(Of Integer, String) From {
            {CInt(ViewOrientationConstants.igFrontView), "Front"},
            {CInt(ViewOrientationConstants.igTopView), "Top"},
            {CInt(ViewOrientationConstants.igRightView), "Right"}
        }
        Dim baseOris As Integer() = {
            CInt(ViewOrientationConstants.igFrontView),
            CInt(ViewOrientationConstants.igTopView),
            CInt(ViewOrientationConstants.igRightView)
        }
        For Each ori In baseOris
            Dim probeDoc As DraftDocument = Nothing
            Try
                probeDoc = CType(app.Documents.Add("SolidEdge.DraftDocument", templatePath), DraftDocument)
                DoIdleSafe(app, $"size-only {oriNames(ori)}")
                Dim sheet As Sheet = probeDoc.ActiveSheet
                Dim modelLink As ModelLink = probeDoc.ModelLinks.Add(modelPath)
                DoIdleSafe(app, $"size-only link {oriNames(ori)}")
                Dim dv As DrawingView = Nothing
                Dim x As Double = 0.15 : Dim y As Double = 0.2 : Dim sc As Double = 1.0
                If Not isSheetMetal Then
                    dv = sheet.DrawingViews.AddPartView(modelLink, ori, sc, x, y, PartDrawingViewTypeConstants.sePartDesignedView)
                Else
                    dv = sheet.DrawingViews.AddSheetMetalView(modelLink, ori, sc, x, y, SheetMetalDrawingViewTypeConstants.seSheetMetalDesignedView)
                End If
                DoIdleSafe(app, $"size-only insert {oriNames(ori)}")
                SafeUpdateView(dv, $"SizeOnly {oriNames(ori)}")
                DoIdleSafe(app, $"size-only update {oriNames(ori)}")
                Dim w As Double = 0 : Dim h As Double = 0
                GetViewSizeSmart(dv, w, h)
                If ori = CInt(ViewOrientationConstants.igFrontView) Then r.W_Front = w : r.H_Front = h
                If ori = CInt(ViewOrientationConstants.igTopView) Then r.W_Top = w : r.H_Top = h
                If ori = CInt(ViewOrientationConstants.igRightView) Then r.W_Right = w : r.H_Right = h
            Catch ex As Exception
                LogEx($"[GetBaseViewSizes] {oriNames(ori)}", ex)
            Finally
                If probeDoc IsNot Nothing Then Try : probeDoc.Close(False) : Catch : End Try
            End Try
        Next
        Return r
    End Function

    ''' <summary>Base = vista de mayor área (Front, Top o Right). Sin rotaciones.</summary>
    Private Function SelectBaseViewByAreaOnly(sizes As BaseViewSizesAtScale1) As Integer
        Dim areaFront As Double = sizes.W_Front * sizes.H_Front
        Dim areaTop As Double = sizes.W_Top * sizes.H_Top
        Dim areaRight As Double = sizes.W_Right * sizes.H_Right
        Dim maxArea As Double = Math.Max(Math.Max(areaFront, areaTop), areaRight)
        StepLog($"[BASE-AREA] Front: área={areaFront:0.0000} Top: área={areaTop:0.0000} Right: área={areaRight:0.0000}")
        If maxArea <= 0 Then
            StepLog("[BASE-AREA] -> Base=Front (por defecto)")
            Return CInt(ViewOrientationConstants.igFrontView)
        End If
        Dim bestOri As Integer
        If maxArea = areaFront Then
            bestOri = CInt(ViewOrientationConstants.igFrontView)
        ElseIf maxArea = areaTop Then
            bestOri = CInt(ViewOrientationConstants.igTopView)
        Else
            bestOri = CInt(ViewOrientationConstants.igRightView)
        End If
        StepLog($"[BASE-AREA] -> Base={OriToConstantName(bestOri)} (mayor área={maxArea:0.0000})")
        Return bestOri
    End Function

    ''' <summary>Vista principal = mayor área. Si H >= 3*W -> rotar 90º.</summary>
    Private Sub SelectBaseViewByArea(sizes As BaseViewSizesAtScale1,
                                     ByRef baseOri As Integer,
                                     ByRef baseW As Double,
                                     ByRef baseH As Double,
                                     ByRef rotated As Boolean)
        Dim areaFront As Double = sizes.W_Front * sizes.H_Front
        Dim areaTop As Double = sizes.W_Top * sizes.H_Top
        Dim areaRight As Double = sizes.W_Right * sizes.H_Right
        Dim maxArea As Double = Math.Max(Math.Max(areaFront, areaTop), areaRight)
        If maxArea <= 0 Then
            baseOri = CInt(ViewOrientationConstants.igRightView)
            baseW = sizes.H_Right : baseH = sizes.W_Right : rotated = False
            Return
        End If
        Dim w As Double, h As Double
        If maxArea = areaFront Then
            baseOri = CInt(ViewOrientationConstants.igFrontView)
            w = sizes.W_Front : h = sizes.H_Front
        ElseIf maxArea = areaTop Then
            baseOri = CInt(ViewOrientationConstants.igTopView)
            w = sizes.W_Top : h = sizes.H_Top
        Else
            baseOri = CInt(ViewOrientationConstants.igRightView)
            w = sizes.W_Right : h = sizes.H_Right
        End If
        If h >= ASPECT_ROTATE_THRESHOLD * w Then
            rotated = True
            baseW = h : baseH = w
        Else
            rotated = False
            baseW = w : baseH = h
        End If
    End Sub

    ''' <summary>Ancho efectivo de la vista adyacente en horizontal (la menor de las dos ortogonales para no sobrestimar).</summary>
    Private Function GetAdjacentW(sizes As BaseViewSizesAtScale1, baseOri As Integer) As Double
        Dim oriRight As Integer, oriBelow As Integer
        GetOrisForBaseEuropean(baseOri, oriRight, oriBelow)
        Dim maxRight As Double = 0.0 : Dim maxBelow As Double = 0.0
        Dim mF = Math.Max(sizes.W_Front, sizes.H_Front)
        Dim mT = Math.Max(sizes.W_Top, sizes.H_Top)
        Dim mR = Math.Max(sizes.W_Right, sizes.H_Right)
        Dim fallback = (mF + mT + mR) / 3.0

        If oriRight = CInt(ViewOrientationConstants.igFrontView) Then
            maxRight = mF
        ElseIf oriRight = CInt(ViewOrientationConstants.igTopView) Then
            maxRight = mT
        ElseIf oriRight = CInt(ViewOrientationConstants.igRightView) Then
            maxRight = mR
        Else : maxRight = fallback
        End If

        If oriBelow = CInt(ViewOrientationConstants.igFrontView) Then
            maxBelow = mF
        ElseIf oriBelow = CInt(ViewOrientationConstants.igTopView) Then
            maxBelow = mT
        ElseIf oriBelow = CInt(ViewOrientationConstants.igRightView) Then
            maxBelow = mR
        Else
            maxBelow = fallback
        End If

        Return Math.Min(maxRight, maxBelow)
    End Function

    ''' <summary>Escala inicial: usableW / (baseW + adjacentW + numGaps * gap).</summary>
    Private Function ComputeInitialScale(usableW As Double, baseW As Double, adjacentW As Double) As Double
        Dim total As Double = baseW + adjacentW + NUM_GAPS_HORIZONTAL * GAP_MODEL_UNITS
        If total <= 0 Then Return 0.5
        Return usableW / total
    End Function




    'Private Structure ViewMetrics
    '    Public Ori As Integer
    '    Public W As Double
    '    Public H As Double
    '    Public Score As Double
    'End Structure

    'Private Function MeasureView(veDu As DrawingView) As (w As Double, h As Double)
    '    Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
    '    veDu.GetRange(x1, y1, x2, y2)
    '    Dim w = Math.Abs(x2 - x1)
    '    Dim h = Math.Abs(y2 - y1)
    '    Return (w, h)
    'End Function

    'Private Function ScoreLandscape(w As Double, h As Double) As Double
    '    ' Prioriza que sea más ancha que alta, pero también que tenga “tamaño” (área).
    '    If h <= 0 Then Return 0
    '    Dim aspect = w / h             ' >1 = horizontal
    '    Dim area = w * h
    '    Return aspect * Math.Sqrt(Math.Max(area, 0.000000000001))
    'End Function



    'Private Sub InsertStandard4(sheet As Sheet,
    '                            modelLink As ModelLink,
    '                            isSheetMetal As Boolean,
    '                            scale As Double,
    '                            baseOri As Integer,
    '                            m As Margins,
    '                            ByRef vFront As DrawingView,
    '                            ByRef vTop As DrawingView,
    '                            ByRef vRight As DrawingView,
    '                            ByRef vIso As DrawingView)

    '    Dim dvws As DrawingViews = sheet.DrawingViews
    '    Dim W As Double = sheet.SheetSetup.SheetWidth
    '    Dim H As Double = sheet.SheetSetup.SheetHeight
    '    Dim xSafe As Double = m.Left + 0.05
    '    Dim ySafe As Double = (H - m.Top) - 0.05

    '    Dim xFront As Double = xSafe
    '    Dim yFront As Double = ySafe
    '    Dim xTop As Double = xSafe
    '    Dim yTop As Double = ySafe - 0.06
    '    Dim xRight As Double = xSafe + 0.12
    '    Dim yRight As Double = ySafe
    '    Dim xIso As Double = xSafe + 0.24
    '    Dim yIso As Double = ySafe - 0.06

    '    StepLog($"InsertStandard4: W={W} H={H}")
    '    StepLog($"Insert pos: Front({xFront},{yFront}) Top({xTop},{yTop}) Right({xRight},{yRight}) Iso({xIso},{yIso}) scale={scale}")

    '    If Not isSheetMetal Then
    '        Dim vt As PartDrawingViewTypeConstants = PartDrawingViewTypeConstants.sePartDesignedView
    '        vFront = dvws.AddPartView(modelLink, baseOri, scale, xFront, yFront, vt)
    '        vTop = dvws.AddPartView(modelLink, CInt(ViewOrientationConstants.igTopView), scale, xTop, yTop, vt)
    '        vRight = dvws.AddPartView(modelLink, CInt(ViewOrientationConstants.igRightView), scale, xRight, yRight, vt)
    '        vIso = dvws.AddPartView(modelLink, CInt(ViewOrientationConstants.igTrimetricTopFrontLeftView), scale * ISO_FACTOR, xIso, yIso, vt)
    '    Else
    '        Dim vt As SheetMetalDrawingViewTypeConstants = SheetMetalDrawingViewTypeConstants.seSheetMetalDesignedView
    '        vFront = dvws.AddSheetMetalView(modelLink, baseOri, scale, xFront, yFront, vt)
    '        vTop = dvws.AddSheetMetalView(modelLink, CInt(ViewOrientationConstants.igTopView), scale, xTop, yTop, vt)
    '        vRight = dvws.AddSheetMetalView(modelLink, CInt(ViewOrientationConstants.igRightView), scale, xRight, yRight, vt)
    '        vIso = dvws.AddSheetMetalView(modelLink, CInt(ViewOrientationConstants.igTrimetricTopFrontLeftView), scale * ISO_FACTOR, xIso, yIso, vt)
    '    End If
    'End Sub



    ''' <summary>Vista principal láser = misma heurística que DFT (mayor área en Front/Top/Left + giro 90° si aplica).</summary>
    Public Structure LaserPrincipalViewChoice
        Public BaseOri As Integer
        Public ApplyRotMinus90 As Boolean
        Public BaseOriName As String
    End Structure

    Public Function ResolveLaserPrincipalViewLayout(app As SolidEdgeFramework.Application,
                                                    modelPath As String,
                                                    cleanTemplatePath As String) As LaserPrincipalViewChoice
        Dim choice As New LaserPrincipalViewChoice With {
            .BaseOri = CInt(ViewOrientationConstants.igFrontView),
            .ApplyRotMinus90 = False,
            .BaseOriName = "Front"
        }
        If String.IsNullOrWhiteSpace(modelPath) Then Return choice

        Dim isSheetMetal As Boolean = modelPath.EndsWith(".psm", StringComparison.OrdinalIgnoreCase)
        Dim sizesOpt = GetBaseViewSizesAtScale1(app, modelPath, cleanTemplatePath, isSheetMetal)
        If Not sizesOpt.HasValue Then Return choice

        If Not String.IsNullOrWhiteSpace(cleanTemplatePath) AndAlso File.Exists(cleanTemplatePath) Then
            Dim usable = LayoutEngine.GetUsableAreaForTemplate(cleanTemplatePath)
            Dim fold = SelectBaseViewForLayoutByFold(sizesOpt.Value, usable)
            choice.BaseOri = fold.BaseOri
            choice.BaseOriName = fold.BaseOriName
            choice.ApplyRotMinus90 = fold.Rotated
            Return choice
        End If

        choice.BaseOri = SelectBaseViewByAreaOnly(sizesOpt.Value)
        choice.BaseOriName = OriToConstantName(choice.BaseOri)
        Return choice
    End Function

    ''' <summary>Orientación de vista principal para plano láser (delega en <see cref="ResolveLaserPrincipalViewLayout"/>).</summary>
    Public Function ResolveLaserPartMainOrientation(app As SolidEdgeFramework.Application,
                                                   modelPath As String,
                                                   cleanTemplatePath As String) As Integer
        Return ResolveLaserPrincipalViewLayout(app, modelPath, cleanTemplatePath).BaseOri
    End Function

    ''' <summary>Wrapper público para DraftGenerator: prepara FlatPattern y crea la vista desarrollada.</summary>
    Public Function CreateFlatViewForDraft(app As SolidEdgeFramework.Application,
                                           modelPath As String,
                                           dvs As DrawingViews,
                                           modelLink As ModelLink,
                                           baseScale As Double,
                                           ByRef flat As DrawingView) As Boolean
        flat = Nothing
        If app Is Nothing OrElse dvs Is Nothing OrElse modelLink Is Nothing Then Return False
        If Not modelPath.EndsWith(".psm", StringComparison.OrdinalIgnoreCase) Then Return False
        EnsureFlatPatternReady(app, modelPath)
        Return TryCreateFlatView_Safe(dvs, modelLink, baseScale, flat, app)
    End Function

    Private Function TryCreateFlatView_Safe(
        dvs As DrawingViews,
        modelLink As ModelLink,
        baseScale As Double,
        ByRef flat As DrawingView,
        app As SolidEdgeFramework.Application,
        Optional templatePath As String = Nothing
    ) As Boolean

        flat = Nothing

        Try
            Dim x As Double = 0.1, y As Double = 0.1
            Dim useScale As Double = baseScale

            ' Primer intento con la escala solicitada.
            flat = CType(
                dvs.AddSheetMetalView(
                    modelLink,
                    ViewOrientationConstants.igTopView,
                    useScale,
                    x, y,
                    SheetMetalDrawingViewTypeConstants.seSheetMetalFlatView
                ),
                DrawingView
            )

            If flat Is Nothing Then Return False
            SafeUpdateView(flat, "Flat.Update(scale-initial)")

            ' Escala UNE independiente de las 3 principales: slot bbox (preferente) o regla A3 legada.
            Dim targetScale As Double
            If TemplateBboxLayout.UseSlotBasedLayout AndAlso Not String.IsNullOrWhiteSpace(templatePath) Then
                targetScale = TemplateBboxLayout.ResolveFlatScaleForBboxSlot(
                    dvs, flat, templatePath, useScale, Sub(msg) StepLog(msg))
            Else
                targetScale = ResolveA3FlatScaleIfNeeded(dvs, flat, useScale)
            End If
            If targetScale > 0.0R AndAlso Math.Abs(targetScale - useScale) > EPS Then
                Try
                    flat.Delete()
                Catch
                End Try
                flat = Nothing

                flat = CType(
                    dvs.AddSheetMetalView(
                        modelLink,
                        ViewOrientationConstants.igTopView,
                        targetScale,
                        x, y,
                        SheetMetalDrawingViewTypeConstants.seSheetMetalFlatView
                    ),
                    DrawingView
                )
                If flat IsNot Nothing Then
                    SafeUpdateView(flat, "Flat.Update(scale-une-adjusted)")
                    StepLog("Flat recreado con escala UNE ajustada=" & targetScale.ToString("0.######", Globalization.CultureInfo.InvariantCulture))
                End If
            End If

            StepLog("Flat creado con AddSheetMetalView (seSheetMetalFlatView).")
            Return (flat IsNot Nothing)

        Catch ex As Exception
            LogEx("TryCreateFlatView(AddSheetMetalView seSheetMetalFlatView)", ex)
        End Try

        Try
            StepLog("Fallback Flat por reflexión...")
            Return TryCreateFlatView_ByReflection(dvs, modelLink, baseScale, flat)
        Catch ex As Exception
            LogEx("TryCreateFlatView_ByReflection", ex)
        End Try

        Return False
    End Function

    Private Function ResolveA3FlatScaleIfNeeded(ByVal dvs As DrawingViews, ByVal flat As DrawingView, ByVal currentScale As Double) As Double
        If dvs Is Nothing OrElse flat Is Nothing OrElse currentScale <= 0.0R Then Return currentScale
        Try
            Dim parentSheet As Sheet = TryCast(CallByName(dvs, "Parent", CallType.Get), Sheet)
            If parentSheet Is Nothing Then Return currentScale

            Dim sheetW As Double = 0.0R, sheetH As Double = 0.0R
            GetSheetSize(parentSheet, sheetW, sheetH)
            If sheetW <= 0.0R OrElse sheetH <= 0.0R Then Return currentScale

            ' Solo aplicar regla explícita para A3 (con tolerancia).
            Dim isA3 As Boolean = (Math.Abs(sheetW - 0.42R) <= 0.01R AndAlso Math.Abs(sheetH - 0.297R) <= 0.01R) OrElse
                                  (Math.Abs(sheetW - 0.297R) <= 0.01R AndAlso Math.Abs(sheetH - 0.42R) <= 0.01R)
            If Not isA3 Then Return currentScale

            Dim m As Margins = GetMarginsByTemplate("A3 Plantilla.dft")
            Dim availableW As Double = A3_CAJETIN_LEFT_X - m.Left
            If availableW <= 0.0R Then Return currentScale

            Dim flatW As Double = 0.0R, flatH As Double = 0.0R
            GetViewSizeSmart(flat, flatW, flatH)
            If flatW <= 0.0R Then Return currentScale

            If flatW <= availableW + EPS Then Return currentScale

            Dim rawTargetScale As Double = currentScale * (availableW / flatW) * FLAT_FIT_SAFETY
            Dim snapped As Double = SnapToUneReductionScale(rawTargetScale)
            If snapped <= 0.0R Then Return currentScale

            StepLog("A3 FLAT width check: flatW=" & flatW.ToString("0.######", Globalization.CultureInfo.InvariantCulture) &
                    " availW=" & availableW.ToString("0.######", Globalization.CultureInfo.InvariantCulture) &
                    " currentScale=" & currentScale.ToString("0.######", Globalization.CultureInfo.InvariantCulture) &
                    " targetRaw=" & rawTargetScale.ToString("0.######", Globalization.CultureInfo.InvariantCulture) &
                    " targetUNE=" & snapped.ToString("0.######", Globalization.CultureInfo.InvariantCulture))
            Return snapped
        Catch
            Return currentScale
        End Try
    End Function

    Private Function SnapToUneReductionScale(ByVal maxScale As Double) As Double
        If maxScale <= 0.0R Then Return 0.0R
        For Each s In UneNormalizedReductionScales
            If s <= maxScale + EPS Then Return s
        Next
        Return UneNormalizedReductionScales(UneNormalizedReductionScales.Length - 1)
    End Function

    Private Function TryCreateFlatView_ByReflection(dvs As DrawingViews, modelLink As ModelLink, scale As Double, ByRef flat As DrawingView) As Boolean
        flat = Nothing

        Dim preferredNames As String() = {
            "AddSheetMetalDevelopedView",
            "AddSheetMetalFlatPatternView",
            "AddFlatPatternView",
            "AddDevelopedView",
            "AddDevelopedDrawingView"
        }

        Dim x As Double = 0.1, y As Double = 0.1

        Dim argSets As Object()() = {
            New Object() {modelLink, x, y, scale},
            New Object() {modelLink, scale, x, y},
            New Object() {modelLink, x, y},
            New Object() {modelLink, CInt(ViewOrientationConstants.igFrontView), scale, x, y}
        }

        Dim t = dvs.GetType()
        Dim methods = t.GetMethods(BindingFlags.Instance Or BindingFlags.Public)

        Dim candidates = methods.
            Where(Function(mi)
                      Dim n = mi.Name
                      Return preferredNames.Any(Function(p) String.Equals(n, p, StringComparison.OrdinalIgnoreCase)) _
                             OrElse n.IndexOf("Develop", StringComparison.OrdinalIgnoreCase) >= 0 _
                             OrElse n.IndexOf("Flat", StringComparison.OrdinalIgnoreCase) >= 0
                  End Function).
            OrderBy(Function(mi)
                        Dim n = mi.Name.ToLowerInvariant()
                        If n.Contains("sheetmetal") AndAlso n.Contains("develop") Then Return 0
                        If n.Contains("flatpattern") Then Return 1
                        If n.Contains("develop") Then Return 2
                        If n.Contains("flat") Then Return 3
                        Return 9
                    End Function).ToList()

        For Each mi In candidates
            Dim pc = mi.GetParameters().Length
            StepLog($"TryFlat(reflection): {mi.Name} params={pc}")

            For Each args In argSets
                If args.Length <> pc Then Continue For
                Try
                    Dim obj = mi.Invoke(dvs, args)
                    flat = TryCast(obj, DrawingView)
                    If flat IsNot Nothing Then
                        StepLog($"Flat creado con reflexión: {mi.Name}")
                        Return True
                    End If
                Catch ex As Exception
                    Dim inner = TryCast(ex, TargetInvocationException)
                    If inner IsNot Nothing AndAlso inner.InnerException IsNot Nothing Then
                        LogEx($"Flat Invoke {mi.Name}", inner.InnerException)
                    Else
                        LogEx($"Flat Invoke {mi.Name}", ex)
                    End If
                End Try
            Next
        Next

        Return False
    End Function

    Private Function LayoutFits_4(usableW As Double, usableH As Double,
                                  frontW As Double, frontH As Double,
                                  topW As Double, topH As Double,
                                  rightW As Double, rightH As Double,
                                  isoW As Double, isoH As Double,
                                  gapH As Double, gapV As Double) As Boolean

        Dim col1W = Math.Max(frontW, topW)
        Dim col2W = rightW
        Dim col3W = isoW

        Dim needW = col1W + gapH + col2W + gapH + col3W

        Dim leftStackH = frontH + gapV + topH
        Dim needH = Math.Max(rightH, Math.Max(leftStackH, isoH))

        StepLog($"LayoutFits(4): needW={needW} usableW={usableW} | needH={needH} usableH={usableH}")
        Return (needW <= usableW + EPS) AndAlso (needH <= usableH + EPS)
    End Function

    Private Function LayoutUtil_4(usableW As Double, usableH As Double,
                                  frontW As Double, frontH As Double,
                                  topW As Double, topH As Double,
                                  rightW As Double, rightH As Double,
                                  isoW As Double, isoH As Double,
                                  gapH As Double, gapV As Double) As Double

        Dim col1W = Math.Max(frontW, topW)
        Dim col2W = rightW
        Dim col3W = isoW
        Dim needW = col1W + gapH + col2W + gapH + col3W

        Dim leftStackH = frontH + gapV + topH
        Dim needH = Math.Max(rightH, Math.Max(leftStackH, isoH))

        Dim uW = needW / usableW
        Dim uH = needH / usableH
        Return Math.Max(uW, uH)
    End Function

    Private Sub ApplyLayout_IsoTop_FlatBottom(app As SolidEdgeFramework.Application,
                                              sheetW As Double, sheetH As Double, m As Margins,
                                              front As DrawingView, frontW As Double, frontH As Double,
                                              top As DrawingView, topW As Double, topH As Double,
                                              right As DrawingView, rightW As Double, rightH As Double,
                                              iso As DrawingView, isoW As Double, isoH As Double,
                                              flat As DrawingView,
                                              gapH As Double, gapV As Double,
                                              hasFlat As Boolean)

        Dim x0 As Double = m.Left
        Dim yTopSheet As Double = sheetH - m.Top

        Dim col1W As Double = Math.Max(frontW, topW)
        Dim x2 As Double = x0 + col1W + gapH
        Dim x3 As Double = x2 + rightW + gapH

        StepLog($"ApplyLayout: x0={x0} yTopSheet={yTopSheet} x2={x2} x3={x3}")

        MoveViewTopLeft(app, front, x0, yTopSheet, "Front")
        SafeUpdateView(front, "Front.Update(layout)") : DoIdleSafe(app, "after Front move")

        Dim frontCX As Double, frontCY As Double
        Dim fx1 As Double, fy1 As Double, fx2 As Double, fy2 As Double
        If TryGetViewRange(front, fx1, fy1, fx2, fy2) Then
            frontCX = (fx1 + fx2) / 2.0
            frontCY = (fy1 + fy2) / 2.0
        Else
            frontCX = x0 + frontW / 2.0
            frontCY = yTopSheet - frontH / 2.0
        End If

        Dim yTop2 As Double = yTopSheet - frontH - gapV
        MoveViewTopLeft(app, top, x0, yTop2, "Top")
        AlignViewCenterX(app, top, frontCX, "Top(align CX con base)")
        SafeUpdateView(top, "Top.Update(layout)") : DoIdleSafe(app, "after Top move")

        MoveViewTopLeft(app, right, x2, yTopSheet, "Right(pre)")
        AlignViewCenterY(app, right, frontCY, "Right(align CY con base)")
        SafeUpdateView(right, "Right.Update(layout)") : DoIdleSafe(app, "after Right align")

        ' ISO más arriba (esquina superior derecha) para no invadir el cajetín
        Dim yIsoTop As Double = yTopSheet - isoH - gapV
        MoveViewTopLeft(app, iso, x3, yIsoTop, "Iso")
        SafeUpdateView(iso, "Iso.Update(layout)") : DoIdleSafe(app, "after Iso move")

        ' Chapa desarrollada en el hueco inferior entre cajetín y borde izquierdo del template
        If hasFlat AndAlso flat IsNot Nothing Then
            Dim flatW As Double, flatH As Double : GetViewSizeSmart(flat, flatW, flatH)
            Dim yFlatTop As Double = m.Bottom + gapV + flatH
            MoveViewTopLeft(app, flat, x0, yFlatTop, "Flat")
            SafeUpdateView(flat, "Flat.Update(layout)") : DoIdleSafe(app, "after Flat move")
        End If
    End Sub

    ''' <summary>Centra todas las vistas en el área útil del template usando el rango de los dibujos y el marco (márgenes).</summary>
    Private Sub CenterAllViewsOnUsableArea(app As SolidEdgeFramework.Application, dvs As DrawingViews, sheetW As Double, sheetH As Double, m As Margins)
        Dim views = dvs.OfType(Of DrawingView)().ToList()
        If views.Count = 0 Then Return

        Dim xminAll As Double = Double.PositiveInfinity
        Dim yminAll As Double = Double.PositiveInfinity
        Dim xmaxAll As Double = Double.NegativeInfinity
        Dim ymaxAll As Double = Double.NegativeInfinity

        For Each v In views
            Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
            If Not TryGetViewRange(v, x1, y1, x2, y2) Then Continue For
            xminAll = Math.Min(xminAll, Math.Min(x1, x2))
            yminAll = Math.Min(yminAll, Math.Min(y1, y2))
            xmaxAll = Math.Max(xmaxAll, Math.Max(x1, x2))
            ymaxAll = Math.Max(ymaxAll, Math.Max(y1, y2))
        Next

        If xminAll > xmaxAll OrElse yminAll > ymaxAll Then Return

        Dim curCX As Double = (xminAll + xmaxAll) / 2.0
        Dim curCY As Double = (yminAll + ymaxAll) / 2.0

        Dim usableLeft As Double = m.Left
        Dim usableRight As Double = sheetW - m.Right
        Dim usableBottom As Double = m.Bottom
        Dim usableTop As Double = sheetH - m.Top

        ' Centro del área útil. Corrección empírica que funcionaba.
        Dim centerOffsetX As Double = 0.025
        Dim centerOffsetY As Double = -0.02
        Dim tgtCX As Double = (usableLeft + usableRight) / 2.0 + centerOffsetX
        Dim tgtCY As Double = (usableBottom + usableTop) / 2.0 + centerOffsetY

        Dim dx As Double = tgtCX - curCX
        Dim dy As Double = tgtCY - curCY

        StepLog($"CenterAllViews: curC=({curCX},{curCY}) tgtC=({tgtCX},{tgtCY}) => d=({dx},{dy})")

        For Each v In views
            Dim ox As Double = 0, oy As Double = 0
            v.GetOrigin(ox, oy)
            v.SetOrigin(ox + dx, oy + dy)
        Next

        DoIdleSafe(app, "after CenterAllViews SetOrigin")
    End Sub

    Private Function TryGetViewRange(dv As DrawingView,
                                     ByRef xmin As Double, ByRef ymin As Double,
                                     ByRef xmax As Double, ByRef ymax As Double) As Boolean
        xmin = 0 : ymin = 0 : xmax = 0 : ymax = 0

        Try
            dv.Range(xmin, ymin, xmax, ymax)

            Dim x1 = xmin, y1 = ymin, x2 = xmax, y2 = ymax
            xmin = Math.Min(x1, x2) : ymin = Math.Min(y1, y2)
            xmax = Math.Max(x1, x2) : ymax = Math.Max(y1, y2)

            Return True
        Catch ex As Exception
            LogEx("TryGetViewRange(dv.Range)", ex)
            Return False
        End Try
    End Function

    ''' <summary>Expone lectura de Range para módulos de layout (p. ej. TemplateBboxLayout).</summary>
    Public Function TryGetViewRangePublic(dv As DrawingView,
                                          ByRef xmin As Double, ByRef ymin As Double,
                                          ByRef xmax As Double, ByRef ymax As Double) As Boolean
        Return TryGetViewRange(dv, xmin, ymin, xmax, ymax)
    End Function

    Private Sub GetViewSizeSmart(dv As DrawingView, ByRef w As Double, ByRef h As Double)
        Dim xmin As Double, ymin As Double, xmax As Double, ymax As Double
        If Not TryGetViewRange(dv, xmin, ymin, xmax, ymax) Then
            w = 0 : h = 0
            Exit Sub
        End If
        w = xmax - xmin
        h = ymax - ymin
    End Sub

    Private Sub MoveViewTopLeft(app As SolidEdgeFramework.Application, dv As DrawingView, targetLeft As Double, targetTop As Double, tag As String)
        If dv Is Nothing Then Return

        Dim xmin As Double, ymin As Double, xmax As Double, ymax As Double
        If Not TryGetViewRange(dv, xmin, ymin, xmax, ymax) Then
            StepLog($"MoveViewTopLeft[{tag}]: NO RANGE")
            Return
        End If

        Dim ox As Double = 0, oy As Double = 0
        dv.GetOrigin(ox, oy)

        Dim dx As Double = targetLeft - xmin
        Dim dy As Double = targetTop - ymax

        dv.SetOrigin(ox + dx, oy + dy)
        DoIdleSafe(app, $"after SetOrigin {tag}")
    End Sub

    Private Sub AlignViewCenterY(app As SolidEdgeFramework.Application, dv As DrawingView, targetCY As Double, tag As String)
        If dv Is Nothing Then Return

        Dim xmin As Double, ymin As Double, xmax As Double, ymax As Double
        If Not TryGetViewRange(dv, xmin, ymin, xmax, ymax) Then
            StepLog($"AlignViewCenterY[{tag}]: NO RANGE")
            Return
        End If

        Dim cy As Double = (ymin + ymax) / 2.0
        Dim dy As Double = targetCY - cy

        Dim ox As Double = 0, oy As Double = 0
        dv.GetOrigin(ox, oy)
        dv.SetOrigin(ox, oy + dy)
        DoIdleSafe(app, $"after AlignCenterY {tag}")
    End Sub

    Private Sub AlignViewCenterX(app As SolidEdgeFramework.Application, dv As DrawingView, targetCX As Double, tag As String)
        If dv Is Nothing Then Return

        Dim xmin As Double, ymin As Double, xmax As Double, ymax As Double
        If Not TryGetViewRange(dv, xmin, ymin, xmax, ymax) Then
            StepLog($"AlignViewCenterX[{tag}]: NO RANGE")
            Return
        End If

        Dim cx As Double = (xmin + xmax) / 2.0
        Dim dx As Double = targetCX - cx

        Dim ox As Double = 0, oy As Double = 0
        dv.GetOrigin(ox, oy)
        dv.SetOrigin(ox + dx, oy)
        DoIdleSafe(app, $"after AlignCenterX {tag}")
    End Sub

    ''' <summary>Rota -90º el bloque de las tres vistas (base + derecha + debajo). Siempre las tres para mantener primer diedro.</summary>
    Private Sub ApplyRotationToPrincipalViews(app As SolidEdgeFramework.Application,
                                             baseOri As Integer,
                                             vBase As DrawingView,
                                             vViewRight As DrawingView,
                                             vViewBelow As DrawingView)
        If vBase IsNot Nothing Then
            SafeSetRotationAngle(vBase, -RIGHT_ROT_RAD)
            DoIdleSafe(app, "rotate base")
            SafeUpdateView(vBase, "Base.Update rotated")
        End If
        If vViewRight IsNot Nothing Then
            SafeSetRotationAngle(vViewRight, -RIGHT_ROT_RAD)
            DoIdleSafe(app, "rotate viewRight")
            SafeUpdateView(vViewRight, "ViewRight.Update rotated")
        End If
        If vViewBelow IsNot Nothing Then
            SafeSetRotationAngle(vViewBelow, -RIGHT_ROT_RAD)
            DoIdleSafe(app, "rotate viewBelow")
            SafeUpdateView(vViewBelow, "ViewBelow.Update rotated")
        End If
    End Sub

    Private Sub SafeSetRotationAngle(v As DrawingView, angleRad As Double, Optional tag As String = "")
        If v Is Nothing Then Return
        Try
            v.SetRotationAngle(angleRad)
            StepLog($"SetRotationAngle({tag},{angleRad * 180 / Math.PI:0}º) OK")
        Catch ex As Exception
            LogEx($"SetRotationAngle({tag})", ex)
        End Try
    End Sub

    Private Sub SafeUpdateView(v As DrawingView, tag As String)
        If v Is Nothing Then
            StepLog($"{tag}: (Nothing)")
            Return
        End If
        Try
            v.Update()
            StepLog($"{tag}: OK")
        Catch ex As Exception
            LogEx(tag, ex)
        End Try
    End Sub

    Private Sub DumpAllViews(dvs As DrawingViews)
        Try
            Dim i As Integer = 0
            For Each v As DrawingView In dvs.OfType(Of DrawingView)()
                i += 1
                Dim ox As Double = 0, oy As Double = 0
                v.GetOrigin(ox, oy)
                Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
                If Not TryGetViewRange(v, x1, y1, x2, y2) Then
                    StepLog($"View#{i}: ORIGIN=({ox},{oy}) RANGE=<<NO>>")
                Else
                    StepLog($"View#{i}: ORIGIN=({ox},{oy}) RANGE=[{x1},{y1}]->[{x2},{y2}] SIZE=({x2 - x1} x {y2 - y1})")
                End If
            Next
        Catch ex As Exception
            LogEx("DumpAllViews", ex)
        End Try
    End Sub

    ''' <summary>Inserta vista 1:1 en plano de corte láser (sin recálculo UNE de escala).</summary>
    Public Function InsertLaserCutPieceView(app As SolidEdgeFramework.Application,
                                            draftDoc As DraftDocument,
                                            sheet As Sheet,
                                            modelPath As String,
                                            templatePath As String,
                                            insertX As Double,
                                            insertY As Double,
                                            scale As Double,
                                            preferSheetMetalFlat As Boolean,
                                            log As Action(Of String)) As DrawingView
        If app Is Nothing OrElse draftDoc Is Nothing OrElse sheet Is Nothing OrElse String.IsNullOrWhiteSpace(modelPath) Then Return Nothing
        Dim fullPath As String = IO.Path.GetFullPath(modelPath.Trim())
        If Not File.Exists(fullPath) Then
            log?.Invoke("[LASER][VIEW][FAIL] file not found path=" & fullPath)
            Return Nothing
        End If

        Dim isPsm As Boolean = fullPath.EndsWith(".psm", StringComparison.OrdinalIgnoreCase)
        If isPsm AndAlso preferSheetMetalFlat Then EnsureFlatPatternReady(app, fullPath)

        Dim link As ModelLink = Nothing
        Try
            link = draftDoc.ModelLinks.Add(fullPath)
            DoIdleSafe(app, "Laser ModelLink")
        Catch ex As Exception
            LogEx("InsertLaserCutPieceView ModelLinks.Add", ex)
            log?.Invoke("[LASER][VIEW][FAIL] ModelLink " & ex.Message)
            Return Nothing
        End Try
        If link Is Nothing Then
            log?.Invoke("[LASER][VIEW][FAIL] ModelLink=null path=" & fullPath)
            Return Nothing
        End If

        Dim dvs As DrawingViews = sheet.DrawingViews
        Dim dv As DrawingView = Nothing
        If isPsm Then
            If preferSheetMetalFlat Then
                dv = TryInsertLaserPsmFlatView(app, dvs, link, scale, insertX, insertY, log)
                If dv Is Nothing Then
                    log?.Invoke("[LASER][VIEW][FALLBACK] PSM sin flat usable, probando vista principal")
                    dv = TryInsertLaserPsmMainView(app, dvs, link, fullPath, templatePath, scale, insertX, insertY, log)
                End If
            Else
                dv = TryInsertLaserPsmMainView(app, dvs, link, fullPath, templatePath, scale, insertX, insertY, log)
            End If
        Else
            dv = TryInsertLaserParMainView(app, dvs, link, fullPath, templatePath, scale, insertX, insertY, log)
        End If

        If dv Is Nothing Then Return Nothing

        SafeUpdateView(dv, "LaserCut.Update")
        DoIdleSafe(app, "LaserCut idle1")
        SafeUpdateView(dv, "LaserCut.Update2")
        DoIdleSafe(app, "LaserCut idle2")

        Dim xmin As Double, ymin As Double, xmax As Double, ymax As Double
        If TryGetViewRange(dv, xmin, ymin, xmax, ymax) Then
            Dim w As Double = Math.Max(0.0R, xmax - xmin)
            Dim h As Double = Math.Max(0.0R, ymax - ymin)
            log?.Invoke("[LASER][VIEW][RANGE] w=" & w.ToString("0.0000", Globalization.CultureInfo.InvariantCulture) &
                        " h=" & h.ToString("0.0000", Globalization.CultureInfo.InvariantCulture))
            If w < 0.000001R OrElse h < 0.000001R Then
                log?.Invoke("[LASER][VIEW][FAIL] zero range path=" & IO.Path.GetFileName(fullPath))
                Return Nothing
            End If
        Else
            log?.Invoke("[LASER][VIEW][WARN] no Range() path=" & IO.Path.GetFileName(fullPath))
        End If

        log?.Invoke("[LASER][VIEW][ADD] file=" & IO.Path.GetFileName(fullPath) & " scale=" & scale.ToString("0.####", Globalization.CultureInfo.InvariantCulture))
        Return dv
    End Function

    Private Function TryInsertLaserPsmFlatView(app As SolidEdgeFramework.Application,
                                              dvs As DrawingViews,
                                              link As ModelLink,
                                              scale As Double,
                                              insertX As Double,
                                              insertY As Double,
                                              log As Action(Of String)) As DrawingView
        Dim dv As DrawingView = Nothing
        Try
            dv = CType(
                dvs.AddSheetMetalView(
                    link,
                    ViewOrientationConstants.igTopView,
                    scale,
                    insertX,
                    insertY,
                    SheetMetalDrawingViewTypeConstants.seSheetMetalFlatView),
                DrawingView)
            If dv IsNot Nothing Then
                log?.Invoke("[LASER][VIEW][INSERT] PSM flat AddSheetMetalView")
                Return dv
            End If
        Catch ex As Exception
            LogEx("TryInsertLaserPsmFlatView AddSheetMetalView", ex)
            log?.Invoke("[LASER][VIEW][FAIL] PSM flat " & ex.Message)
        End Try

        Dim flat As DrawingView = Nothing
        If TryCreateFlatView_ByReflection(dvs, link, scale, flat) AndAlso flat IsNot Nothing Then
            log?.Invoke("[LASER][VIEW][INSERT] PSM flat reflection")
            Return flat
        End If

        log?.Invoke("[LASER][VIEW][FAIL] PSM all flat strategies failed")
        Return Nothing
    End Function

    Private Function TryInsertLaserPsmMainView(app As SolidEdgeFramework.Application,
                                               dvs As DrawingViews,
                                               link As ModelLink,
                                               modelPath As String,
                                               templatePath As String,
                                               scale As Double,
                                               insertX As Double,
                                               insertY As Double,
                                               log As Action(Of String)) As DrawingView
        Return TryInsertLaserDesignedMainView(app, dvs, link, modelPath, templatePath, scale, insertX, insertY, isSheetMetal:=True, log)
    End Function

    Private Function TryInsertLaserParMainView(app As SolidEdgeFramework.Application,
                                               dvs As DrawingViews,
                                               link As ModelLink,
                                               modelPath As String,
                                               templatePath As String,
                                               scale As Double,
                                               insertX As Double,
                                               insertY As Double,
                                               log As Action(Of String)) As DrawingView
        Return TryInsertLaserDesignedMainView(app, dvs, link, modelPath, templatePath, scale, insertX, insertY, isSheetMetal:=False, log)
    End Function

    ''' <summary>Inserta vista principal 1:1 (mayor área / LayoutByFold, igual que motor DFT).</summary>
    Private Function TryInsertLaserDesignedMainView(app As SolidEdgeFramework.Application,
                                                    dvs As DrawingViews,
                                                    link As ModelLink,
                                                    modelPath As String,
                                                    templatePath As String,
                                                    scale As Double,
                                                    insertX As Double,
                                                    insertY As Double,
                                                    isSheetMetal As Boolean,
                                                    log As Action(Of String)) As DrawingView
        Dim layout As LaserPrincipalViewChoice
        Try
            layout = ResolveLaserPrincipalViewLayout(app, modelPath, templatePath)
        Catch ex As Exception
            LogEx("ResolveLaserPrincipalViewLayout", ex)
            layout = New LaserPrincipalViewChoice With {.BaseOri = CInt(ViewOrientationConstants.igFrontView), .BaseOriName = "Front"}
        End Try

        log?.Invoke("[LASER][VIEW][MAIN] file=" & IO.Path.GetFileName(modelPath) &
                    " base=" & layout.BaseOriName &
                    " ori=" & layout.BaseOri.ToString(Globalization.CultureInfo.InvariantCulture) &
                    " rot90=" & layout.ApplyRotMinus90.ToString())

        Dim oris As New List(Of Integer) From {layout.BaseOri}
        For Each fallbackOri In New Integer() {
            CInt(ViewOrientationConstants.igFrontView),
            CInt(ViewOrientationConstants.igTopView),
            CInt(ViewOrientationConstants.igRightView)}
            If Not oris.Contains(fallbackOri) Then oris.Add(fallbackOri)
        Next

        Dim kindTag As String = If(isSheetMetal, "PSM", "PAR")
        For Each ori In oris
            Dim dv As DrawingView = Nothing
            Dim useLayoutRotation As Boolean = (ori = layout.BaseOri) AndAlso layout.ApplyRotMinus90
            Try
                If isSheetMetal Then
                    dv = dvs.AddSheetMetalView(
                        link, ori, scale, insertX, insertY,
                        SheetMetalDrawingViewTypeConstants.seSheetMetalDesignedView)
                Else
                    dv = dvs.AddPartView(link, ori, scale, insertX, insertY, PartDrawingViewTypeConstants.sePartDesignedView)
                End If
                If dv Is Nothing Then
                    log?.Invoke("[LASER][VIEW][FAIL] " & kindTag & " principal null ori=" & ori.ToString(Globalization.CultureInfo.InvariantCulture))
                    Continue For
                End If
                ForceViewOrientationStandard(dv, ori, "Laser" & kindTag)
                If useLayoutRotation Then
                    SafeSetRotationAngle(dv, -RIGHT_ROT_RAD)
                    log?.Invoke("[LASER][VIEW][ROTATE] " & kindTag & " -90° (LayoutByFold)")
                End If
                SafeUpdateView(dv, "Laser" & kindTag & ".Update")
                DoIdleSafe(app, "Laser" & kindTag & " idle")
                Dim xmin As Double, ymin As Double, xmax As Double, ymax As Double
                If TryGetViewRange(dv, xmin, ymin, xmax, ymax) Then
                    Dim w As Double = Math.Max(0.0R, xmax - xmin)
                    Dim h As Double = Math.Max(0.0R, ymax - ymin)
                    If w >= 0.000001R AndAlso h >= 0.000001R Then
                        log?.Invoke("[LASER][VIEW][INSERT] " & kindTag & " principal ori=" & ori.ToString(Globalization.CultureInfo.InvariantCulture) &
                                    If(useLayoutRotation, " rot=-90", ""))
                        Return dv
                    End If
                End If
                log?.Invoke("[LASER][VIEW][FAIL] " & kindTag & " principal zero range ori=" & ori.ToString(Globalization.CultureInfo.InvariantCulture))
                Try : dv.Delete() : Catch : End Try
            Catch ex As Exception
                LogEx("TryInsertLaserDesignedMainView " & kindTag, ex)
                log?.Invoke("[LASER][VIEW][FAIL] " & kindTag & " principal ori=" & ori.ToString(Globalization.CultureInfo.InvariantCulture) & " " & ex.Message)
            End Try
        Next

        log?.Invoke("[LASER][VIEW][FAIL] " & kindTag & " principal all orientations failed path=" & IO.Path.GetFileName(modelPath))
        Return Nothing
    End Function

End Module