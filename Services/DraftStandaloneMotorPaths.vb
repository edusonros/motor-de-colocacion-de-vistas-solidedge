Option Strict Off

Imports System.IO
Imports Extraer_dft_dxf_flatdxf.Services.Dimensioning.Labs

''' <summary>
''' Rutas de DFT esperadas por los motores sueltos (metadatos / acotación), alineadas con <see cref="DraftGenerationEngine"/>.
''' </summary>
Public NotInheritable Class DraftStandaloneMotorPaths
    Private Sub New()
    End Sub

    Public Shared Function EffectiveRunLab(config As JobConfiguration) As Boolean
        If config Is Nothing Then Return False
        Return config.EnableDrawingViewDimensioningLab OrElse DimensionInsertionConfig.EnableDrawingViewDimensioningLab
    End Function

    ''' <remarks>El nombre en disco es fijo (<c>_DIMLAB_REF_TEST</c>) para no mezclar builds; <paramref name="mode"/> solo afecta al laboratorio interno.</remarks>
    Public Shared Function DimLabStemForDraft(baseStem As String, mode As DimLabMode) As String
        Return baseStem & "_DIMLAB_REF_TEST"
    End Function

    ''' <summary>Ruta absoluta del .dft que abrirán los motores Metadatos o Acotación (sin considerar sufijos _001 si overwrite=False).</summary>
    Public Shared Function GetExpectedStandaloneDraftPath(config As JobConfiguration, modelPath As String) As String
        If config Is Nothing OrElse String.IsNullOrWhiteSpace(config.OutputFolder) OrElse String.IsNullOrWhiteSpace(modelPath) Then Return ""
        Dim baseName As String = Path.GetFileNameWithoutExtension(modelPath)
        Dim stem As String = If(EffectiveRunLab(config), DimLabStemForDraft(baseName, config.DimLabMode), baseName)
        Dim root As String = config.OutputFolder.Trim().TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
        Return Path.Combine(root, "DFT", stem & ".dft")
    End Function

    ''' <summary>
    ''' Si la entrada ya es un <c>.dft</c> existente, los motores sueltos abren ese archivo.
    ''' Si no, se usa el DFT esperado bajo <c>OutputFolder\DFT\</c> (misma regla que la generación previa).
    ''' </summary>
    Public Shared Function ResolveStandaloneDraftFullPath(config As JobConfiguration, modelOrDraftPath As String) As String
        If String.IsNullOrWhiteSpace(modelOrDraftPath) Then Return ""
        Try
            If String.Equals(Path.GetExtension(modelOrDraftPath), ".dft", StringComparison.OrdinalIgnoreCase) AndAlso File.Exists(modelOrDraftPath) Then
                Return Path.GetFullPath(modelOrDraftPath)
            End If
        Catch
            Return ""
        End Try
        Return GetExpectedStandaloneDraftPath(config, modelOrDraftPath)
    End Function

    Public Shared Function FormatMissingDraftHint(expectedPath As String) As String
        Return "Para ""Gestor Metadatos"" o ""Motor Acotación"" debe existir el DFT generado previamente:" &
            Environment.NewLine & "  " & expectedPath &
            Environment.NewLine & "Pulse primero ""Generador de vistas"" o ""GENERAR"" con la misma carpeta de salida y la misma pieza," &
            Environment.NewLine & "o elija un archivo .dft existente como entrada y use esos motores."
    End Function
End Class
