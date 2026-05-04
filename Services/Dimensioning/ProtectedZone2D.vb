Option Strict Off

''' <summary>Zona rectangular en coordenadas de hoja (m) para evitar vistas/cotas (PartsList, cajetín, etc.).</summary>
Public Class ProtectedZone2D
    Public Property Name As String
    Public Property MinX As Double
    Public Property MinY As Double
    Public Property MaxX As Double
    Public Property MaxY As Double

    Public Function ToBoundingBox2D() As BoundingBox2D
        Return New BoundingBox2D With {
            .MinX = MinX,
            .MinY = MinY,
            .MaxX = MaxX,
            .MaxY = MaxY
        }
    End Function
End Class
