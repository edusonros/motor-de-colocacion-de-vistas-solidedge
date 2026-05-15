Option Strict Off
Namespace Extraer_dft_dxf_flatdxf.LegacyV02Dimensioning
' --- Copia aislada desde Planos_Automaticos_v02 (no editar v02); motor independiente del arbol Services\Dimensioning vigente. ---
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

End Namespace

