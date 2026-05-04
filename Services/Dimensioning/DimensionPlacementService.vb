Option Strict Off

Imports System.Globalization
Imports SolidEdgeDraft

''' <summary>Seguimiento de filas/columnas de cotas exteriores para separar H/V (ISO 129 pragmático).</summary>
Public Class Iso129PlacementContext
    Public Property BottomRowIndex As Integer
    Public Property TopRowIndex As Integer
    Public Property RightColIndex As Integer
    Public Property LeftColIndex As Integer
End Class

Public NotInheritable Class DimensionPlacementService
    Private Sub New()
    End Sub

    ''' <summary>Rectángulo de hoja en metros (origen típico abajo-izquierda). TODO: márgenes finos desde SheetSetup.</summary>
    Public Shared Function TryGetSheetWorkArea(sheet As Sheet, log As Action(Of String)) As BoundingBox2D
        Dim b As New BoundingBox2D With {.MinX = 0, .MinY = 0, .MaxX = 0.297, .MaxY = 0.21}
        If sheet Is Nothing Then Return b
        Try
            Dim su = sheet.SheetSetup
            If su IsNot Nothing Then
                b.MaxX = su.SheetWidth
                b.MaxY = su.SheetHeight
            End If
        Catch ex As Exception
            log?.Invoke("[DIM][ISO129][PLACE] SheetSetup: " & ex.Message)
        End Try
        Return b
    End Function

    ''' <summary>Zona conservadora de cajetín (esquina inferior derecha). TODO: leer bbox real del cajetín.</summary>
    Public Shared Function GetTitleBlockAvoidanceBox(sheetWork As BoundingBox2D, config As DimensioningNormConfig) As BoundingBox2D
        If sheetWork Is Nothing OrElse Not config.AvoidTitleBlock Then Return Nothing
        Dim w As Double = sheetWork.Width
        Dim h As Double = sheetWork.Height
        Return New BoundingBox2D With {
            .MinX = sheetWork.MinX + w * 0.42R,
            .MinY = sheetWork.MinY,
            .MaxX = sheetWork.MaxX,
            .MaxY = sheetWork.MinY + h * 0.35R
        }
    End Function

    Public Shared Function CalcularPosicionExteriorCota(
        cand As DimensionCandidate,
        bbox As BoundingBox2D,
        ctx As Iso129PlacementContext,
        config As DimensioningNormConfig,
        sheetArea As BoundingBox2D,
        log As Action(Of String)) As Point2D

        If cand Is Nothing OrElse bbox Is Nothing OrElse config Is Nothing Then Return New Point2D()
        If ctx Is Nothing Then ctx = New Iso129PlacementContext()

        Dim gap As Double = Math.Max(config.MinGapFromView, 1.0E-6)
        Dim rowGap As Double = Math.Max(config.GapBetweenDimensionRows, 1.0E-6)
        Dim cx As Double = (bbox.MinX + bbox.MaxX) / 2.0R
        Dim cy As Double = (bbox.MinY + bbox.MaxY) / 2.0R

        Dim pt As New Point2D()

        Select Case cand.Type
            Case DimensionCandidateType.TotalHorizontal
                Dim row As Integer = ctx.BottomRowIndex
                ctx.BottomRowIndex += 1
                pt.X = cx
                pt.Y = bbox.MinY - gap - row * rowGap
                cand.PlacementSide = DimensionSide.Bottom
                log?.Invoke(String.Format(CultureInfo.InvariantCulture,
                    "[DIM][ISO129][PLACE] H_TOTAL bottom row={0} Y={1:0.######}", row, pt.Y))

            Case DimensionCandidateType.PartialHorizontal
                Dim row As Integer = ctx.TopRowIndex
                ctx.TopRowIndex += 1
                pt.X = cx
                pt.Y = bbox.MaxY + gap + row * rowGap
                cand.PlacementSide = DimensionSide.Top
                log?.Invoke(String.Format(CultureInfo.InvariantCulture,
                    "[DIM][ISO129][PLACE] H_PARTIAL top row={0} Y={1:0.######}", row, pt.Y))

            Case DimensionCandidateType.TotalVertical
                Dim col As Integer = ctx.RightColIndex
                ctx.RightColIndex += 1
                pt.X = bbox.MaxX + gap + col * rowGap
                pt.Y = cy
                cand.PlacementSide = DimensionSide.Right
                log?.Invoke(String.Format(CultureInfo.InvariantCulture,
                    "[DIM][ISO129][PLACE] V_TOTAL right col={0} X={1:0.######}", col, pt.X))

            Case DimensionCandidateType.PartialVertical
                Dim col As Integer = ctx.LeftColIndex
                ctx.LeftColIndex += 1
                pt.X = bbox.MinX - gap - col * rowGap
                pt.Y = cy
                cand.PlacementSide = DimensionSide.Left
                log?.Invoke(String.Format(CultureInfo.InvariantCulture,
                    "[DIM][ISO129][PLACE] V_PARTIAL left col={0} X={1:0.######}", col, pt.X))

            Case Else
                pt.X = cx
                pt.Y = bbox.MaxY + gap
                cand.PlacementSide = DimensionSide.Unknown
        End Select

        If sheetArea IsNot Nothing AndAlso config.AvoidTitleBlock Then
            Dim tb = GetTitleBlockAvoidanceBox(sheetArea, config)
            If tb IsNot Nothing AndAlso PointInsideBox(pt, tb) Then
                log?.Invoke("[DIM][ISO129][PLACE] Invasión cajetín; intento lado opuesto (heurística).")
                If cand.Type = DimensionCandidateType.TotalHorizontal Then
                    ctx.BottomRowIndex = Math.Max(0, ctx.BottomRowIndex - 1)
                    ctx.TopRowIndex += 1
                    pt.Y = bbox.MaxY + gap + ctx.TopRowIndex * rowGap
                    cand.PlacementSide = DimensionSide.Top
                ElseIf cand.Type = DimensionCandidateType.TotalVertical Then
                    ctx.RightColIndex = Math.Max(0, ctx.RightColIndex - 1)
                    ctx.LeftColIndex += 1
                    pt.X = bbox.MinX - gap - ctx.LeftColIndex * rowGap
                    cand.PlacementSide = DimensionSide.Left
                End If
            End If
        End If

        cand.PlacementPoint = pt
        Return pt
    End Function

    Private Shared Function PointInsideBox(p As Point2D, b As BoundingBox2D) As Boolean
        If p Is Nothing OrElse b Is Nothing Then Return False
        Return p.X >= b.MinX AndAlso p.X <= b.MaxX AndAlso p.Y >= b.MinY AndAlso p.Y <= b.MaxY
    End Function

    ''' <summary>Ajusta Y de picks sobre líneas verticales; X sobre horizontales.</summary>
    Public Shared Sub RefinarPicksCotaExterior(cand As DimensionCandidate, placementLineY As Double, placementLineX As Double)
        If cand Is Nothing Then Return
        If cand.Orientation = DimensionOrientation.Horizontal Then
            If cand.P1 Is Nothing Then cand.P1 = New Point2D()
            If cand.P2 Is Nothing Then cand.P2 = New Point2D()
            Dim yUse As Double = placementLineY
            cand.P1.Y = yUse
            cand.P2.Y = yUse
        ElseIf cand.Orientation = DimensionOrientation.Vertical Then
            If cand.P1 Is Nothing Then cand.P1 = New Point2D()
            If cand.P2 Is Nothing Then cand.P2 = New Point2D()
            Dim xUse As Double = placementLineX
            cand.P1.X = xUse
            cand.P2.X = xUse
        End If
    End Sub

    Public Shared Function ClampPickYToVerticalLine(yDesired As Double, line As DvLineGeomInfo) As Double
        If line Is Nothing Then Return yDesired
        Return Math.Max(line.MinYs, Math.Min(line.MaxYs, yDesired))
    End Function

    Public Shared Function ClampPickXToHorizontalLine(xDesired As Double, line As DvLineGeomInfo) As Double
        If line Is Nothing Then Return xDesired
        Return Math.Max(line.MinXs, Math.Min(line.MaxXs, xDesired))
    End Function
End Class
