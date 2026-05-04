Option Strict Off

Imports System.Globalization

''' <summary>
''' Configuración global de inserción de cotas exteriores (sin depender de JobConfiguration para no romper firmas existentes).
''' Ajustar desde UI o código al iniciar la app si hace falta.
''' </summary>
Friend Module DimensionInsertionConfig

    ''' <summary>Versión extendida de log de coordenadas, round-trip vista/hoja y evidencia post-creación.</summary>
    Friend EnableDimensionInsertionDiagnostics As Boolean = True

    ''' <summary>Estrategia activa para picks de proximidad en Dimensions.AddDistanceBetweenObjects.</summary>
    Friend ActiveExteriorPickStrategy As ExteriorPickStrategy = ExteriorPickStrategy.SilhouetteProximityB

    ''' <summary>Si True, los extremos exteriores priorizan aristas DVLine2d; si no hay suficientes datos, se usa el resolvedor geométrico completo.</summary>
    Friend PreferLineOnlyExteriorReferences As Boolean = True

    ''' <summary>Separación mínima (m) legacy; la colocación exterior principal usa ExteriorDimensionOffsetFromRangeM.</summary>
    Friend SilhouettePickInsetM As Double = 0.0004R

    ''' <summary>
    ''' Separación de la línea de cota respecto al marco <c>DrawingView.Range</c> en hoja (m). Cotas exteriores:
    ''' horizontal → <c>Y = MaxY + esto</c>; vertical → <c>X = MaxX + esto</c>. Por defecto 0,02 m (2 cm).
    ''' Ajustado a 0,012 m para acercar la línea de cota.
    ''' </summary>
    Friend ExteriorDimensionOffsetFromRangeM As Double = 0.012R

    ''' <summary>
    ''' Si True, tras fallar el intento en HOJA se reintenta con picks en espacio vista (SheetToView).
    ''' El orden es siempre hoja primero; la API encaja mejor y el Range de la cota queda en coordenadas de pliego.
    ''' </summary>
    Friend UseViewSpaceProximityForAddDistance As Boolean = True

End Module

' --- Auto-ajuste de rango para que las dimensiones recién creadas sean visibles ---
Friend Module DimensioningViewRangeTuning
    ' Si True, tras insertar una dimensión se intenta ampliar el DrawingView.Range para incluir su Range real.
    Friend AutoExpandDrawingViewRangeForDimensions As Boolean = True

    ' Margen extra al ampliar el rango (m). Pequeño para que no afecte al encuadre global.
    Friend DrawingViewRangeExpansionPaddingM As Double = 0.003R
End Module

''' <summary>Estrategias documentadas para experimentación; en producción usar <see cref="ExteriorPickStrategy.SilhouetteProximityB"/>.</summary>
Friend Enum ExteriorPickStrategy
    ''' <summary>Picks en hoja con offset fuerte respecto al bbox de la vista (comportamiento histórico problemático con vistas pequeñas).</summary>
    SheetOffsetBandA = 0
    ''' <summary>Picks en hoja pegados a la silueta real (top X/Y de extremos) + pequeño inset; recomendado.</summary>
    SilhouetteProximityB = 1
    ''' <summary>Solo diagnóstico: misma inserción que B pero registro explícito round-trip Sheet↔View (no cambia picks respecto a B si la API usa hoja).</summary>
    RoundTripDiagnosticC = 2
    ''' <summary>Solo referencias desde líneas cuando el resolvedor estable lo permite; picks = misma regla que B.</summary>
    LineOnlyReferencesD = 3
End Enum

Friend NotInheritable Class DimensionInsertionConfigFormat

    Friend Shared Function StrategyName(s As ExteriorPickStrategy) As String
        Select Case s
            Case ExteriorPickStrategy.SheetOffsetBandA : Return "A_SheetOffsetBand"
            Case ExteriorPickStrategy.SilhouetteProximityB : Return "B_SilhouetteProximity"
            Case ExteriorPickStrategy.RoundTripDiagnosticC : Return "C_RoundTripDiag"
            Case ExteriorPickStrategy.LineOnlyReferencesD : Return "D_LineOnlyRefs"
            Case Else : Return s.ToString()
        End Select
    End Function

    Friend Shared Function Bool(b As Boolean) As String
        Return b.ToString(CultureInfo.InvariantCulture)
    End Function

End Class
