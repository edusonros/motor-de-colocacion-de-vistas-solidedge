Option Strict Off

Namespace Services.Dimensioning.Labs

    ''' <summary>Modo de ejecución del laboratorio DIMLAB (no afecta al motor principal de acotación).</summary>
    Public Enum DimLabMode
        HorizontalOnly = 0
        VerticalOnly = 1
        Full = 2
    ForensicHorizontal = 3
    CleanFull = 4
    CleanFullStrict = 5
    End Enum

    ''' <summary>Extremos del contorno 2D de la vista (m) según DVLine2d + lectura tipada.</summary>
    Public Structure DVBounds
        Public MinX As Double
        Public MaxX As Double
        Public MinY As Double
        Public MaxY As Double
        ''' <summary>Ancho horizontal total esperado en coordenadas de vista (≈ bounds.MaxX - MinX).</summary>
        Public ExpectedWidth As Double
        ''' <summary>Alto vertical total esperado (≈ bounds.MaxY - MinY).</summary>
        Public ExpectedHeight As Double
    End Structure

End Namespace
