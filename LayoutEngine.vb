Option Strict Off
Imports System.IO
Imports System.Runtime.InteropServices
Imports SolidEdgeDraft
Imports SolidEdgeFramework
Imports SolidEdgePart
Imports System.Linq

''' <summary>Motor de layout geométrico con restricciones duras y criterios de calidad industrial.</summary>
Public Module LayoutEngine

    Private Sub LayoutLog(msg As String)
        Console.WriteLine($"[LAYOUT] {msg}")
    End Sub

#Region "Configuración - Umbrales y márgenes"

    ' Margenes interiores mínimos (m): ninguna vista puede tocar el borde
    Public Const USABLE_MARGIN_X As Double = 0.012   ' 12 mm
    Public Const USABLE_MARGIN_Y As Double = 0.012   ' 12 mm

    ' Holgura mínima aceptable respecto al área efectiva (m)
    Private Const MIN_CLEARANCE As Double = 0.008   ' 8 mm

    ' Umbrales mínimos de aprovechamiento para aceptar un layout
    Private Const MIN_WIDTH_USAGE As Double = 0.45   ' 45% del ancho efectivo
    Private Const MIN_HEIGHT_USAGE As Double = 0.35   ' 35% del alto efectivo
    Private Const MIN_AREA_USAGE As Double = 0.20    ' 20% del área efectiva

    ' Base dominance: umbral relajado para chapas/perfiles (restricción BLANDA)
    Private Const MIN_BASE_DOMINANCE_STRICT As Double = 0.20   ' rechazo fuerte solo si < 20%
    Private Const MIN_BASE_DOMINANCE_RELAXED As Double = 0.08   ' fallback: aceptar si >= 8%

    ' Descentrado máximo aceptable (m): |centro_layout - centro_efectivo|
    Private Const MAX_CENTER_OFFSET As Double = 0.025   ' 25 mm

    ' Escala mínima aceptable: no layouts con escala ridícula
    Private Const MIN_ACCEPTABLE_SCALE As Double = 0.02   ' 1:50 mínimo

    ' Escalas estándar en orden descendente
    Private ReadOnly StandardScales As Double() = {
        1.0, 0.5, 0.2, 0.1, 0.05, 0.04, 1.0 / 30.0, 0.025, 0.02, 1.0 / 75.0, 0.01
    }

    Private Const GAP_H As Double = 0.012
    Private Const GAP_V As Double = 0.012

#End Region

#Region "Estructuras"

    ''' <summary>Punto 2D en metros (coordenadas Solid Edge Draft).</summary>
    Public Structure Point2D
        Public X As Double
        Public Y As Double
        Public Sub New(x As Double, y As Double)
            Me.X = x
            Me.Y = y
        End Sub
    End Structure

    ''' <summary>Rectángulo útil bruto del template. Origen Solid Edge: esquina inferior izquierda, Y hacia arriba.</summary>
    Public Class UsableArea
        Public Property MinX As Double
        Public Property MinY As Double
        Public Property MaxX As Double
        Public Property MaxY As Double

        Public ReadOnly Property Width As Double
            Get
                Return Math.Max(0, MaxX - MinX)
            End Get
        End Property

        Public ReadOnly Property Height As Double
            Get
                Return Math.Max(0, MaxY - MinY)
            End Get
        End Property

        Public ReadOnly Property Center As Point2D
            Get
                Return New Point2D((MinX + MaxX) / 2.0, (MinY + MaxY) / 2.0)
            End Get
        End Property
    End Class

    ''' <summary>Área efectiva: usable menos márgenes interiores. Ninguna vista puede tocar sus bordes.</summary>
    Public Class EffectiveArea
        Public Property MinX As Double
        Public Property MinY As Double
        Public Property MaxX As Double
        Public Property MaxY As Double

        Public ReadOnly Property Width As Double
            Get
                Return Math.Max(0, MaxX - MinX)
            End Get
        End Property

        Public ReadOnly Property Height As Double
            Get
                Return Math.Max(0, MaxY - MinY)
            End Get
        End Property

        Public ReadOnly Property Center As Point2D
            Get
                Return New Point2D((MinX + MaxX) / 2.0, (MinY + MaxY) / 2.0)
            End Get
        End Property
    End Class

    ''' <summary>Candidato de layout con todas las métricas obligatorias para validación industrial.</summary>
    Public Class CandidateLayout
        Public Property TemplatePath As String
        Public Property TemplateName As String
        Public Property BaseViewOri As Integer
        Public Property BaseViewName As String
        Public Property RotationDeg As Integer
        Public Property Rotation As CojonudoBestFit_Bueno.ViewRotation
        Public Property Scale As Double

        ' Dimensiones a escala 1
        Public Property BaseWidthAt1 As Double
        Public Property BaseHeightAt1 As Double
        Public Property SideWidthAt1 As Double
        Public Property SideHeightAt1 As Double
        Public Property BelowWidthAt1 As Double
        Public Property BelowHeightAt1 As Double
        Public Property IsoWidthAt1 As Double
        Public Property IsoHeightAt1 As Double

        ' Dimensiones escaladas (m)
        Public Property BaseWidth As Double
        Public Property BaseHeight As Double
        Public Property SideWidth As Double
        Public Property SideHeight As Double
        Public Property BelowWidth As Double
        Public Property BelowHeight As Double
        Public Property IsoWidth As Double
        Public Property IsoHeight As Double

        Public Property TotalWidth As Double
        Public Property TotalHeight As Double

        Public Property IncludeIso As Boolean

        ' Métricas obligatorias
        Public Property LayoutCenterX As Double
        Public Property LayoutCenterY As Double
        Public Property LeftClearance As Double
        Public Property RightClearance As Double
        Public Property TopClearance As Double
        Public Property BottomClearance As Double
        Public Property WidthUsage As Double
        Public Property HeightUsage As Double
        Public Property AreaUsage As Double
        Public Property BaseViewArea As Double
        Public Property SecondaryViewsArea As Double
        Public Property MinClearance As Double   ' Mínimo de los 4 clearances

        Public Property Fits As Boolean
        Public Property HardRejected As Boolean
        Public Property HardRejectReason As String
        Public Property SoftPenalty As String
        Public Property CompactnessScore As Double
        Public Property BaseDominanceScore As Double
        Public Property Score As Double
    End Class

#End Region

#Region "GetUsableAreaForTemplate / GetEffectiveArea"

    ''' <summary>Rectángulo útil bruto del template. Coordenadas en metros.
    ''' Origen Solid Edge: esquina inferior izquierda, Y hacia arriba.
    ''' A3: 420x297 mm. Marco interior con márgenes. Cajetín 190x30 abajo-derecha.</summary>
    Public Function GetUsableAreaForTemplate(templatePath As String) As UsableArea
        Dim ua As New UsableArea
        Dim name As String = If(String.IsNullOrEmpty(templatePath), "", Path.GetFileName(templatePath)).ToLowerInvariant()

        If name.Contains("a3") Then
            ' A3: 420x297mm. Vista principal 40mm del borde superior. Cajetín 190x30 abajo-derecha (X>=0.23).
            ua.MinX = 0.02
            ua.MaxX = 0.38   ' Ancho útil (cajetín empieza ~230mm, ISO en esa zona)
            ua.MinY = 0.045  ' Por encima del cajetín (35mm)
            ua.MaxY = 0.257  ' 40mm del borde superior
        ElseIf name.Contains("a2") Then
            ' A2: ~594x420 mm en horizontal
            ua.MinX = 0.02
            ua.MaxX = 0.58
            ua.MinY = 0.04
            ua.MaxY = 0.42 - 0.02
        Else
            ua.MinX = 0.02 : ua.MinY = 0.04 : ua.MaxX = 0.27 : ua.MaxY = 0.38
        End If

        Return ua
    End Function

    ''' <summary>Área efectiva = usable - márgenes. Todo el layout debe caber DENTRO de este rectángulo.</summary>
    Public Function GetEffectiveArea(usable As UsableArea, Optional marginX As Double = USABLE_MARGIN_X, Optional marginY As Double = USABLE_MARGIN_Y) As EffectiveArea
        Dim eff As New EffectiveArea
        eff.MinX = usable.MinX + marginX
        eff.MinY = usable.MinY + marginY
        eff.MaxX = usable.MaxX - marginX
        eff.MaxY = usable.MaxY - marginY
        Return eff
    End Function

    ''' <summary>Información del template para logging: dimensiones de la hoja, área útil y cajetín.
    ''' Unidades en metros. Origen Solid Edge: esquina inferior izquierda, Y hacia arriba.</summary>
    Public Class TemplateInfo
        Public Property TemplateWidth As Double   ' mm
        Public Property TemplateHeight As Double  ' mm
        Public Property TemplateOriginX As Double  ' m
        Public Property TemplateOriginY As Double ' m
        Public Property UsableMinX As Double      ' m
        Public Property UsableMinY As Double      ' m
        Public Property UsableWidth As Double     ' m
        Public Property UsableHeight As Double    ' m
        Public Property CajetinOriginX As Double  ' m
        Public Property CajetinOriginY As Double  ' m
        Public Property CajetinWidth As Double    ' m
        Public Property CajetinHeight As Double   ' m
    End Class

    ''' <summary>Obtiene dimensiones del template, área útil y cajetín para logging.</summary>
    Public Function GetTemplateInfo(templatePath As String) As TemplateInfo
        Dim ti As New TemplateInfo
        ti.TemplateOriginX = 0
        ti.TemplateOriginY = 0
        Dim name As String = If(String.IsNullOrEmpty(templatePath), "", Path.GetFileName(templatePath)).ToLowerInvariant()

        If name.Contains("a3") Then
            ti.TemplateWidth = 420 : ti.TemplateHeight = 297
            Dim ua = GetUsableAreaForTemplate(templatePath)
            ti.UsableMinX = ua.MinX : ti.UsableMinY = ua.MinY
            ti.UsableWidth = ua.Width : ti.UsableHeight = ua.Height
            ti.CajetinOriginX = 0.23 : ti.CajetinOriginY = 0
            ti.CajetinWidth = 0.19 : ti.CajetinHeight = 0.03
        ElseIf name.Contains("a2") Then
            ti.TemplateWidth = 594 : ti.TemplateHeight = 420
            Dim ua = GetUsableAreaForTemplate(templatePath)
            ti.UsableMinX = ua.MinX : ti.UsableMinY = ua.MinY
            ti.UsableWidth = ua.Width : ti.UsableHeight = ua.Height
            ti.CajetinOriginX = 0.404 : ti.CajetinOriginY = 0
            ti.CajetinWidth = 0.19 : ti.CajetinHeight = 0.03
        Else
            ti.TemplateWidth = 297 : ti.TemplateHeight = 210
            Dim ua = GetUsableAreaForTemplate(templatePath)
            ti.UsableMinX = ua.MinX : ti.UsableMinY = ua.MinY
            ti.UsableWidth = ua.Width : ti.UsableHeight = ua.Height
            ti.CajetinOriginX = 0.107 : ti.CajetinOriginY = 0
            ti.CajetinWidth = 0.19 : ti.CajetinHeight = 0.03
        End If
        Return ti
    End Function

#End Region

#Region "ComputeGlobalLayoutSize / BuildCandidateLayouts"

    ''' <summary>Tamaño global del layout (base + side + below + opcional iso) en metros.</summary>
    Public Sub ComputeGlobalLayoutSize(cand As CandidateLayout, includeIso As Boolean,
                                       ByRef totalW As Double, ByRef totalH As Double)
        totalW = cand.BaseWidth + GAP_H + cand.SideWidth
        If includeIso Then totalW += GAP_H + cand.IsoWidth
        totalH = Math.Max(cand.BaseHeight + GAP_V + cand.BelowHeight, Math.Max(cand.SideHeight, If(includeIso, cand.IsoHeight, 0)))
    End Sub

    ''' <summary>Construye candidatos para cada template × base × rotación.</summary>
    Public Function BuildCandidateLayouts(app As SolidEdgeFramework.Application,
                                         modelPath As String,
                                         templates As String(),
                                         cleanTemplatePath As String,
                                         isSheetMetal As Boolean,
                                         sizesAt1 As CojonudoBestFit_Bueno.BaseViewSizesAtScale1) As List(Of CandidateLayout)

        Dim candidates As New List(Of CandidateLayout)
        Dim baseOris As Integer() = {
            CInt(ViewOrientationConstants.igFrontView),
            CInt(ViewOrientationConstants.igTopView),
            CInt(ViewOrientationConstants.igRightView)
        }
        Dim rotations As (CojonudoBestFit_Bueno.ViewRotation, Integer)() = {
            (CojonudoBestFit_Bueno.ViewRotation.Rot0, 0),
            (CojonudoBestFit_Bueno.ViewRotation.RotPlus90, 90),
            (CojonudoBestFit_Bueno.ViewRotation.RotMinus90, -90)
        }

        For Each tpl In templates
            If String.IsNullOrWhiteSpace(tpl) OrElse Not File.Exists(tpl) Then Continue For
            Dim usable As UsableArea = GetUsableAreaForTemplate(tpl)
            Dim effective As EffectiveArea = GetEffectiveArea(usable)

            For Each baseOri In baseOris
                For Each rot In rotations
                    Dim cand = BuildSingleCandidate(sizesAt1, tpl, baseOri, rot.Item1, rot.Item2, effective)
                    If cand IsNot Nothing Then candidates.Add(cand)
                Next
            Next
        Next

        Return candidates
    End Function

    Private Function BuildSingleCandidate(sizesAt1 As CojonudoBestFit_Bueno.BaseViewSizesAtScale1,
                                          templatePath As String,
                                          baseOri As Integer,
                                          rotation As CojonudoBestFit_Bueno.ViewRotation,
                                          rotationDeg As Integer,
                                          effective As EffectiveArea) As CandidateLayout

        Dim map As ProjectedViewMap = CojonudoBestFit_Bueno.GetProjectedViewMap(baseOri, rotation)
        Dim oriRight As Integer = CojonudoBestFit_Bueno.OrthoViewToSolidEdge(map.Right)
        Dim oriBelow As Integer = CojonudoBestFit_Bueno.OrthoViewToSolidEdge(map.Down)

        Dim baseW1 As Double, baseH1 As Double
        GetViewSizeAt1ForOri(sizesAt1, baseOri, baseW1, baseH1)
        If rotation = CojonudoBestFit_Bueno.ViewRotation.RotMinus90 OrElse rotation = CojonudoBestFit_Bueno.ViewRotation.RotPlus90 Then
            Dim tmp = baseW1 : baseW1 = baseH1 : baseH1 = tmp
        End If

        Dim sideW1 As Double, sideH1 As Double
        GetViewSizeAt1ForOri(sizesAt1, oriRight, sideW1, sideH1)
        If rotation = CojonudoBestFit_Bueno.ViewRotation.RotMinus90 OrElse rotation = CojonudoBestFit_Bueno.ViewRotation.RotPlus90 Then
            Dim tmp = sideW1 : sideW1 = sideH1 : sideH1 = tmp
        End If

        Dim belowW1 As Double, belowH1 As Double
        GetViewSizeAt1ForOri(sizesAt1, oriBelow, belowW1, belowH1)
        If rotation = CojonudoBestFit_Bueno.ViewRotation.RotMinus90 OrElse rotation = CojonudoBestFit_Bueno.ViewRotation.RotPlus90 Then
            Dim tmp = belowW1 : belowW1 = belowH1 : belowH1 = tmp
        End If

        Dim isoW1 As Double = Math.Max(baseW1, baseH1) * 0.45
        Dim isoH1 As Double = isoW1

        Dim cand As New CandidateLayout
        cand.TemplatePath = templatePath
        cand.TemplateName = Path.GetFileName(templatePath)
        cand.BaseViewOri = baseOri
        cand.BaseViewName = OriToName(baseOri)
        cand.Rotation = rotation
        cand.RotationDeg = rotationDeg
        cand.BaseWidthAt1 = baseW1 : cand.BaseHeightAt1 = baseH1
        cand.SideWidthAt1 = sideW1 : cand.SideHeightAt1 = sideH1
        cand.BelowWidthAt1 = belowW1 : cand.BelowHeightAt1 = belowH1
        cand.IsoWidthAt1 = isoW1 : cand.IsoHeightAt1 = isoH1

        ' Escala: priorizar 3 vistas principales sobre área efectiva
        Dim scale As Double = ComputeBestScaleForLayout(cand, effective, False)
        cand.Scale = scale

        cand.BaseWidth = baseW1 * scale : cand.BaseHeight = baseH1 * scale
        cand.SideWidth = sideW1 * scale : cand.SideHeight = sideH1 * scale
        cand.BelowWidth = belowW1 * scale : cand.BelowHeight = belowH1 * scale
        cand.IsoWidth = isoW1 * scale : cand.IsoHeight = isoH1 * scale

        ' Sin ISO primero
        cand.IncludeIso = False
        ComputeGlobalLayoutSize(cand, False, cand.TotalWidth, cand.TotalHeight)
        cand.Fits = FitsInEffectiveArea(cand, effective, False)

        ' ISO solo si cabe sin reducir escala de las principales
        If cand.Fits Then
            Dim withIsoW As Double, withIsoH As Double
            ComputeGlobalLayoutSize(cand, True, withIsoW, withIsoH)
            If withIsoW <= effective.Width AndAlso withIsoH <= effective.Height Then
                cand.IncludeIso = True
                cand.TotalWidth = withIsoW
                cand.TotalHeight = withIsoH
            End If
        End If

        ' Métricas obligatorias
        ComputeLayoutUsageMetrics(cand, effective)
        ComputeLayoutCenteringMetrics(cand, effective)
        cand.MinClearance = Math.Min(Math.Min(cand.LeftClearance, cand.RightClearance),
                                    Math.Min(cand.TopClearance, cand.BottomClearance))
        cand.BaseDominanceScore = ComputeBaseDominanceScore(cand)
        cand.CompactnessScore = ComputeCompactnessScore(cand)

        ' Rechazo DURO (solo restricciones absolutas)
        cand.HardRejected = ShouldHardRejectLayout(cand, effective)
        cand.HardRejectReason = If(cand.HardRejected, GetHardRejectReason(cand, effective), "")
        cand.SoftPenalty = If(Not cand.HardRejected, ComputeSoftPenalties(cand, effective), "")
        cand.Score = If(cand.HardRejected, -9999.0, ScoreLayout(cand, effective))

        LayoutLog(FormatCandidateLog(cand, effective))

        Return cand
    End Function

    Private Sub GetViewSizeAt1ForOri(sizes As CojonudoBestFit_Bueno.BaseViewSizesAtScale1, ori As Integer,
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

#End Region

#Region "ComputeBestScaleForLayout / FitsInEffectiveArea"

    ''' <summary>Escala máxima para que el layout quepa en el área efectiva.</summary>
    Public Function ComputeBestScaleForLayout(cand As CandidateLayout, effective As EffectiveArea, includeIso As Boolean) As Double
        Dim needW As Double = cand.BaseWidthAt1 + GAP_H + cand.SideWidthAt1
        If includeIso Then needW += GAP_H + cand.IsoWidthAt1
        Dim needH As Double = cand.BaseHeightAt1 + GAP_V + cand.BelowHeightAt1

        If needW <= 0 OrElse needH <= 0 OrElse effective.Width <= 0 OrElse effective.Height <= 0 Then Return 0.05

        Dim scaleMax As Double = Math.Min(effective.Width / needW, effective.Height / needH)
        For i As Integer = 0 To StandardScales.Length - 1
            If StandardScales(i) <= scaleMax + 0.000001 Then Return StandardScales(i)
        Next
        Return StandardScales(StandardScales.Length - 1)
    End Function

    ''' <summary>Comprueba si el layout cabe en el área efectiva.</summary>
    Public Function FitsInEffectiveArea(cand As CandidateLayout, effective As EffectiveArea, includeIso As Boolean) As Boolean
        Dim needW As Double = cand.BaseWidth + GAP_H + cand.SideWidth
        If includeIso Then needW += GAP_H + cand.IsoWidth
        Dim needH As Double = cand.BaseHeight + GAP_V + cand.BelowHeight
        Return needW <= effective.Width AndAlso needH <= effective.Height
    End Function

#End Region

#Region "ComputeLayoutUsageMetrics / ComputeLayoutCenteringMetrics"

    Private Sub ComputeLayoutUsageMetrics(cand As CandidateLayout, effective As EffectiveArea)
        If effective.Width > 0 Then cand.WidthUsage = cand.TotalWidth / effective.Width Else cand.WidthUsage = 0
        If effective.Height > 0 Then cand.HeightUsage = cand.TotalHeight / effective.Height Else cand.HeightUsage = 0
        Dim effArea As Double = effective.Width * effective.Height
        If effArea > 0 Then cand.AreaUsage = (cand.TotalWidth * cand.TotalHeight) / effArea Else cand.AreaUsage = 0

        cand.BaseViewArea = cand.BaseWidth * cand.BaseHeight
        cand.SecondaryViewsArea = (cand.SideWidth * cand.SideHeight) + (cand.BelowWidth * cand.BelowHeight)
        If cand.IncludeIso Then cand.SecondaryViewsArea += cand.IsoWidth * cand.IsoHeight
    End Sub

    Private Sub ComputeLayoutCenteringMetrics(cand As CandidateLayout, effective As EffectiveArea)
        cand.LayoutCenterX = effective.MinX + (effective.Width - cand.TotalWidth) / 2.0 + cand.TotalWidth / 2.0
        cand.LayoutCenterY = effective.MinY + (effective.Height - cand.TotalHeight) / 2.0 + cand.TotalHeight / 2.0

        Dim layoutLeft As Double = effective.MinX + (effective.Width - cand.TotalWidth) / 2.0
        Dim layoutBottom As Double = effective.MinY + (effective.Height - cand.TotalHeight) / 2.0
        Dim layoutRight As Double = layoutLeft + cand.TotalWidth
        Dim layoutTop As Double = layoutBottom + cand.TotalHeight

        cand.LeftClearance = layoutLeft - effective.MinX
        cand.RightClearance = effective.MaxX - layoutRight
        cand.BottomClearance = layoutBottom - effective.MinY
        cand.TopClearance = effective.MaxY - layoutTop
    End Sub

#End Region

#Region "ComputeBaseDominanceScore / IsBaseDominanceAcceptable"

    ''' <summary>Score adaptativo de dominancia de la vista base. Chapas largas: menor área puede ser válida.</summary>
    Private Function ComputeBaseDominanceScore(candidate As CandidateLayout) As Double
        Dim totalViewArea As Double = candidate.BaseViewArea + candidate.SecondaryViewsArea
        If totalViewArea <= 0 Then Return 0.5
        Dim baseAreaRatio As Double = candidate.BaseViewArea / totalViewArea

        ' Longitud principal representada: base con mayor dimensión aporta más relevancia
        Dim baseMaxDim As Double = Math.Max(candidate.BaseWidthAt1, candidate.BaseHeightAt1)
        Dim totalMaxDim As Double = Math.Max(Math.Max(candidate.BaseWidthAt1 + candidate.SideWidthAt1,
            candidate.BaseHeightAt1 + candidate.BelowHeightAt1), baseMaxDim)
        Dim lengthRelevance As Double = If(totalMaxDim > 0, baseMaxDim / totalMaxDim, 0.5)

        ' Combinación: 70% área + 30% relevancia geométrica (piezas alargadas)
        Dim score As Double = baseAreaRatio * 0.7 + lengthRelevance * 0.3
        Return Math.Min(1.0, score * 1.2)
    End Function

    ''' <summary>True si la dominancia de la base es aceptable (adaptativo para chapas).</summary>
    Private Function IsBaseDominanceAcceptable(candidate As CandidateLayout) As Boolean
        Dim dominanceScore As Double = ComputeBaseDominanceScore(candidate)
        ' Aceptar si score >= 0.15 (muy relajado para chapas largas)
        Return dominanceScore >= 0.12
    End Function

#End Region

#Region "ShouldHardRejectLayout - solo restricciones duras"

    ''' <summary>Rechazo DURO: solo fallos geométricos absolutos. No incluye BaseDominance.</summary>
    Public Function ShouldHardRejectLayout(candidate As CandidateLayout, effective As EffectiveArea) As Boolean
        If Not candidate.Fits Then Return True
        If candidate.Scale < MIN_ACCEPTABLE_SCALE Then Return True
        If candidate.MinClearance < MIN_CLEARANCE Then Return True
        ' Base extremadamente residual (solo rechazo duro si < 5%)
        Dim totalViewArea As Double = candidate.BaseViewArea + candidate.SecondaryViewsArea
        If totalViewArea > 0 Then
            Dim baseRatio As Double = candidate.BaseViewArea / totalViewArea
            If baseRatio < MIN_BASE_DOMINANCE_RELAXED Then Return True
        End If
        ' ISO perjudica mucho
        If candidate.IncludeIso AndAlso totalViewArea > 0 Then
            Dim mainArea As Double = candidate.BaseViewArea + (candidate.SideWidth * candidate.SideHeight) + (candidate.BelowWidth * candidate.BelowHeight)
            If mainArea / totalViewArea < 0.55 Then Return True
        End If
        Return False
    End Function

    Private Function GetHardRejectReason(cand As CandidateLayout, effective As EffectiveArea) As String
        If Not cand.Fits Then Return "DoesNotFit"
        If cand.Scale < MIN_ACCEPTABLE_SCALE Then Return "ScaleTooSmall"
        If cand.MinClearance < MIN_CLEARANCE Then Return "LowClearance"
        Dim totalViewArea As Double = cand.BaseViewArea + cand.SecondaryViewsArea
        If totalViewArea > 0 AndAlso cand.BaseViewArea / totalViewArea < MIN_BASE_DOMINANCE_RELAXED Then
            Return $"BaseResidual({cand.BaseViewArea / totalViewArea:0.00})"
        End If
        Return ""
    End Function

#End Region

#Region "ComputeSoftPenalties / ComputeCompactnessScore"

    Private Function ComputeCompactnessScore(candidate As CandidateLayout) As Double
        ' Penalizar huecos muertos: área útil de vistas / área del bounding box
        Dim viewArea As Double = candidate.BaseViewArea + (candidate.SideWidth * candidate.SideHeight) + (candidate.BelowWidth * candidate.BelowHeight)
        Dim bboxArea As Double = candidate.TotalWidth * candidate.TotalHeight
        If bboxArea <= 0 Then Return 0.8
        Return Math.Min(1.0, viewArea / bboxArea * 1.2)
    End Function

    Private Function ComputeSoftPenalties(candidate As CandidateLayout, effective As EffectiveArea) As String
        Dim penalties As New List(Of String)
        If candidate.WidthUsage < MIN_WIDTH_USAGE Then penalties.Add($"LowWidthUsage({candidate.WidthUsage:0.00})")
        If candidate.HeightUsage < MIN_HEIGHT_USAGE Then penalties.Add($"LowHeightUsage({candidate.HeightUsage:0.00})")
        If candidate.AreaUsage < MIN_AREA_USAGE Then penalties.Add($"LowAreaUsage({candidate.AreaUsage:0.00})")
        If candidate.BaseDominanceScore < MIN_BASE_DOMINANCE_STRICT Then penalties.Add($"BaseDominanceLow({candidate.BaseDominanceScore:0.00})")
        If candidate.CompactnessScore < 0.6 Then penalties.Add($"CompactnessLow({candidate.CompactnessScore:0.00})")
        Dim effCenter As Point2D = effective.Center
        Dim dx As Double = Math.Abs(candidate.LayoutCenterX - effCenter.X)
        Dim dy As Double = Math.Abs(candidate.LayoutCenterY - effCenter.Y)
        If dx + dy > MAX_CENTER_OFFSET Then penalties.Add($"Decentered({(dx + dy) * 1000:0}mm)")
        If penalties.Count = 0 Then Return "None"
        Return String.Join(",", penalties)
    End Function

#End Region

#Region "ScoreLayout"

    ''' <summary>Score ponderado. Incluye penalización por SoftPenalties.</summary>
    Private Function ScoreLayout(cand As CandidateLayout, effective As EffectiveArea) As Double
        Dim scaleScore As Double = Math.Min(1.0, cand.Scale * 2.5)
        Dim widthScore As Double = Math.Min(1.0, cand.WidthUsage / 0.8)
        Dim heightScore As Double = Math.Min(1.0, cand.HeightUsage / 0.8)
        Dim areaScore As Double = Math.Min(1.0, cand.AreaUsage / 0.5)
        Dim effCenter As Point2D = effective.Center
        Dim dx As Double = Math.Abs(cand.LayoutCenterX - effCenter.X)
        Dim dy As Double = Math.Abs(cand.LayoutCenterY - effCenter.Y)
        Dim centerScore As Double = Math.Max(0, 1.0 - (dx + dy) / (MAX_CENTER_OFFSET * 2))
        Dim clearanceScore As Double = Math.Min(1.0, cand.MinClearance / (MIN_CLEARANCE * 2))
        Dim compactScore As Double = cand.CompactnessScore

        Dim total As Double = scaleScore * 0.35 + widthScore * 0.20 + heightScore * 0.12 +
                             areaScore * 0.10 + centerScore * 0.08 + clearanceScore * 0.05 + compactScore * 0.10

        ' Penalización por soft penalties
        If cand.SoftPenalty <> "None" AndAlso cand.SoftPenalty.Length > 0 Then
            If cand.SoftPenalty.Contains("BaseDominanceLow") Then total *= 0.92
            If cand.SoftPenalty.Contains("LowWidthUsage") Then total *= 0.95
            If cand.SoftPenalty.Contains("LowHeightUsage") Then total *= 0.95
        End If

        If cand.TemplateName.ToLowerInvariant().Contains("a3") AndAlso total > 0.55 Then total *= 1.03
        Return Math.Min(1.0, total)
    End Function

#End Region

#Region "ChooseBestLayoutWithFallback"

    ''' <summary>Nunca devuelve Nothing si existe candidato que cabe. Fallback con restricciones relajadas.</summary>
    Public Function ChooseBestLayoutWithFallback(candidates As List(Of CandidateLayout)) As CandidateLayout
        If candidates Is Nothing OrElse candidates.Count = 0 Then
            LayoutLog("[WINNER] Sin candidatos.")
            Return Nothing
        End If

        Dim thatFit = candidates.Where(Function(c) c.Fits).ToList()
        If thatFit.Count = 0 Then
            LayoutLog("[WINNER] Ningún candidato cabe geométricamente.")
            Return Nothing
        End If

        ' FALLBACK A: candidatos no rechazados por restricciones duras
        Dim strictValid = thatFit.Where(Function(c) Not c.HardRejected).ToList()
        If strictValid.Count > 0 Then
            Dim best = strictValid.OrderByDescending(Function(c) c.Score).
                                  ThenByDescending(Function(c) c.Scale).
                                  ThenByDescending(Function(c) c.AreaUsage).
                                  FirstOrDefault()
            LayoutLog($"[WINNER] Template={best.TemplateName} Base={best.BaseViewName} Rot={best.RotationDeg} Scale={best.Scale} reason=BestStrict Score={best.Score:0.000}")
            Return best
        End If

        ' FALLBACK B: todos rechazados por duro -> elegir "menos malo" de los que caben
        Dim relaxed = thatFit.OrderByDescending(Function(c) c.Scale).
                             ThenByDescending(Function(c) c.AreaUsage).
                             ThenByDescending(Function(c) c.MinClearance).
                             FirstOrDefault()
        LayoutLog($"[FALLBACK] No strict winner. Using relaxed: Template={relaxed.TemplateName} Base={relaxed.BaseViewName} Rot={relaxed.RotationDeg} Scale={relaxed.Scale} HardReject={relaxed.HardRejectReason}")
        Return relaxed
    End Function

#End Region

#Region "Posiciones finales / GetInsertParamsFromLayout / ApplyLayoutToDraft"

    ''' <summary>Posición esquina superior-izquierda de la vista base. Usa área EFECTIVA para centrado.</summary>
    Public Function GetBaseTopLeftPoint(best As CandidateLayout, effective As EffectiveArea) As Point2D
        Dim blockW As Double = best.BaseWidth + GAP_H + best.SideWidth
        If best.IncludeIso Then blockW += GAP_H + best.IsoWidth
        Dim blockH As Double = best.BaseHeight + GAP_V + best.BelowHeight
        Dim leftX As Double = effective.MinX + (effective.Width - blockW) / 2.0
        Dim topY As Double = effective.MinY + (effective.Height - blockH) / 2.0 + blockH
        Return New Point2D(leftX, topY)
    End Function

    Public Function GetSideViewTopLeftPoint(best As CandidateLayout, effective As EffectiveArea) As Point2D
        Dim basePt = GetBaseTopLeftPoint(best, effective)
        Return New Point2D(basePt.X + best.BaseWidth + GAP_H, basePt.Y)
    End Function

    Public Function GetBelowViewTopLeftPoint(best As CandidateLayout, effective As EffectiveArea) As Point2D
        Dim basePt = GetBaseTopLeftPoint(best, effective)
        Return New Point2D(basePt.X, basePt.Y - best.BaseHeight - GAP_V)
    End Function

    Public Function GetIsoViewTopLeftPoint(best As CandidateLayout, effective As EffectiveArea) As Point2D
        Dim sidePt = GetSideViewTopLeftPoint(best, effective)
        Return New Point2D(sidePt.X + best.SideWidth + GAP_H, sidePt.Y)
    End Function

    ''' <summary>Parámetros para InsertStandard3Internal.</summary>
    Public Sub GetInsertParamsFromLayout(best As CandidateLayout, effective As EffectiveArea,
                                        ByRef x0 As Double, ByRef y0 As Double,
                                        ByRef gapRight As Double, ByRef gapBelow As Double)
        Dim basePt = GetBaseTopLeftPoint(best, effective)
        Dim sidePt = GetSideViewTopLeftPoint(best, effective)
        Dim belowPt = GetBelowViewTopLeftPoint(best, effective)
        x0 = basePt.X
        y0 = basePt.Y
        gapRight = sidePt.X - basePt.X
        gapBelow = basePt.Y - belowPt.Y
    End Sub

#End Region

#Region "Log format"

    Private Function FormatCandidateLog(cand As CandidateLayout, effective As EffectiveArea) As String
        Dim rotStr As String = If(cand.RotationDeg = 0, "0", If(cand.RotationDeg > 0, $"+{cand.RotationDeg}", cand.RotationDeg.ToString()))
        Dim hardStr As String = If(cand.HardRejected, $" HardReject=True {cand.HardRejectReason}", " HardReject=False")
        Dim softStr As String = If(String.IsNullOrEmpty(cand.SoftPenalty) OrElse cand.SoftPenalty = "None", " SoftPenalty=None", $" SoftPenalty={cand.SoftPenalty}")
        Dim centerOff As Double = Math.Sqrt((cand.LayoutCenterX - effective.Center.X) ^ 2 + (cand.LayoutCenterY - effective.Center.Y) ^ 2) * 1000
        Return $"Template={cand.TemplateName} Base={cand.BaseViewName} Rot={rotStr} Scale={cand.Scale:0.000} Total={cand.TotalWidth * 1000:0}x{cand.TotalHeight * 1000:0} Fits={cand.Fits}{hardStr}{softStr} Compactness={cand.CompactnessScore:0.00} Usage={cand.AreaUsage:0.00} Score={cand.Score:0.000}"
    End Function

#End Region

End Module
