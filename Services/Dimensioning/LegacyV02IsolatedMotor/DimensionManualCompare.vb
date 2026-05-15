Option Strict Off
Imports SolidEdgeFrameworkSupport
Imports FrameworkDimension = SolidEdgeFrameworkSupport.Dimension
Namespace Extraer_dft_dxf_flatdxf.LegacyV02Dimensioning
' --- Copia aislada desde Planos_Automaticos_v02 (no editar v02); motor independiente del arbol Services\Dimensioning vigente. ---
''' <summary>
''' Comparación entre una cota insertada por código y una cota "buena" creada manualmente en Solid Edge.
''' La API no ofrece un flujo fiable para "la cota actualmente seleccionada en la UI" sin automatizar selección;
''' este módulo documenta la lectura de <see cref="Dimension"/> ya existente en la hoja y los límites del interop.
''' </summary>
Friend NotInheritable Class DimensionManualCompare

    ''' <summary>
    ''' Intenta localizar la última dimensión de la colección (heurística débil) y volcar Range/KeyPoints/TrackDistance.
    ''' No sustituye a una selección explícita del usuario; usar solo como apoyo de laboratorio.
    ''' </summary>
    Friend Shared Sub TryLogLastSheetDimensionForComparison(log As DimensionLogger, tag As String)
        If log Is Nothing Then Return
        log.LogLine("[DIM][MANUAL_CMP] " & tag & " nota: la API COM no expone de forma estable 'cota igual a la selección manual actual' sin UI automation.")
        log.LogLine("[DIM][MANUAL_CMP] " & tag & " workaround: inspeccionar Sheet.Dimensions.Item(n) o la última creada; puede no coincidir con la cota que el usuario quiere comparar.")
    End Sub

    Friend Shared Sub TryLogDimensionGeometry(d As FrameworkDimension, log As DimensionLogger, tag As String)
        If d Is Nothing OrElse log Is Nothing Then Return
        DimensionCoordinateDiagnostics.LogCreatedDimensionState(d, log, "[MANUAL_CMP]" & tag, 0, 0, 0, 0)
    End Sub

End Class

End Namespace

