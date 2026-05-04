Option Strict Off

Imports System.Globalization

''' <summary>
''' Marco único de colocación para autoacotado: origen en la esquina inferior izquierda del
''' <see cref="SolidEdgeDraft.DrawingView.Range"/> de la vista base (coordenadas de hoja).
''' Todas las cotas expresan primero posiciones en este marco local y se convierten a hoja con <see cref="ToSheetX"/> / <see cref="ToSheetY"/>.
''' </summary>
Friend NotInheritable Class ViewPlacementFrame

    Public ReadOnly Property OriginX As Double
    Public ReadOnly Property OriginY As Double
    Public ReadOnly Property MinX As Double
    Public ReadOnly Property MinY As Double
    Public ReadOnly Property MaxX As Double
    Public ReadOnly Property MaxY As Double
    Public ReadOnly Property CenterX As Double
    Public ReadOnly Property CenterY As Double
    Public ReadOnly Property Width As Double
    Public ReadOnly Property Height As Double

    Private Sub New(
        originX As Double,
        originY As Double,
        minX As Double,
        minY As Double,
        maxX As Double,
        maxY As Double)

        Me.OriginX = originX
        Me.OriginY = originY
        Me.MinX = minX
        Me.MinY = minY
        Me.MaxX = maxX
        Me.MaxY = maxY
        Me.CenterX = (minX + maxX) / 2.0R
        Me.CenterY = (minY + maxY) / 2.0R
        Me.Width = Math.Abs(maxX - minX)
        Me.Height = Math.Abs(maxY - minY)
    End Sub

    ''' <summary>Crea el marco desde el rectángulo de hoja de la vista base (<c>DrawingView.Range</c> normalizado).</summary>
    Friend Shared Function TryCreateFromBaseViewSheetBox(box As ViewSheetBoundingBox, log As DimensionLogger, ByRef frame As ViewPlacementFrame) As Boolean
        frame = Nothing
        If box.Width <= 1.0E-9 OrElse box.Height <= 1.0E-9 Then
            log?.Err("ViewPlacementFrame: rango de vista base degenerado; no se crea marco.")
            Return False
        End If

        Dim ox As Double = box.MinX
        Dim oy As Double = box.MinY
        frame = New ViewPlacementFrame(ox, oy, box.MinX, box.MinY, box.MaxX, box.MaxY)
        frame.LogFrameHeader(log)
        Return True
    End Function

    ''' <summary>Equivale a <c>localX + OriginX</c> (desplazamiento puro desde el origen de la vista base).</summary>
    Public Function ToSheetX(localX As Double) As Double
        Return OriginX + localX
    End Function

    ''' <summary>Equivale a <c>localY + OriginY</c>.</summary>
    Public Function ToSheetY(localY As Double) As Double
        Return OriginY + localY
    End Function

    Public Function FromSheetX(sheetX As Double) As Double
        Return sheetX - OriginX
    End Function

    Public Function FromSheetY(sheetY As Double) As Double
        Return sheetY - OriginY
    End Function

    Public Function GetSheetBoundingBox() As ViewSheetBoundingBox
        Return New ViewSheetBoundingBox With {
            .MinX = MinX,
            .MinY = MinY,
            .MaxX = MaxX,
            .MaxY = MaxY
        }
    End Function

    Friend Sub LogLocalToSheetX(log As DimensionLogger, localX As Double)
        If log Is Nothing Then Return
        Dim s As Double = ToSheetX(localX)
        log.Frame(String.Format(CultureInfo.InvariantCulture, "local->sheet X: local={0:0.######} => sheet={1:0.######}", localX, s))
    End Sub

    Friend Sub LogLocalToSheetY(log As DimensionLogger, localY As Double)
        If log Is Nothing Then Return
        Dim s As Double = ToSheetY(localY)
        log.Frame(String.Format(CultureInfo.InvariantCulture, "local->sheet Y: local={0:0.######} => sheet={1:0.######}", localY, s))
    End Sub

    Private Sub LogFrameHeader(log As DimensionLogger)
        If log Is Nothing Then Return
        log.Frame(String.Format(CultureInfo.InvariantCulture,
            "baseView.Range=({0:0.######},{1:0.######})-({2:0.######},{3:0.######})",
            MinX, MinY, MaxX, MaxY))
        log.Frame(String.Format(CultureInfo.InvariantCulture,
            "origin=({0:0.######},{1:0.######})",
            OriginX, OriginY))
        log.Frame(String.Format(CultureInfo.InvariantCulture,
            "center=({0:0.######},{1:0.######}) size=({2:0.######},{3:0.######})",
            CenterX, CenterY, Width, Height))
    End Sub

End Class
