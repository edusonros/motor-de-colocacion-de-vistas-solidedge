Option Strict Off

Imports System.Globalization

''' <summary>
''' Flags efectivas durante una corrida del motor de generación (producción vs diagnóstico vs laboratorio COM).
''' Se establecen al inicio de <see cref="DraftGenerationEngine.ProcessModel"/> y se consultan desde servicios de dimensión / PART_LIST.
''' </summary>
Friend NotInheritable Class GenerationEngineRuntime
    ''' <summary>Flujo estándar (DFT+vistas+acotación+guardar) sin rutas DIMLAB exclusivas ni forense.</summary>
    Friend Shared ProductionMode As Boolean = True

    ''' <summary>Barridos extra, colecciones fallidas conocidas y logs de troubleshooting detallados.</summary>
    Friend Shared DebugDiagnosticsMode As Boolean = False

    ''' <summary>Laboratorio DIMLAB (AdditiveLength forense, MsgBox opcional); no confundir con <see cref="Labs.DimLabMode"/> enum de escenario.</summary>
    Friend Shared DimLaboratoryMode As Boolean = False

    Friend Shared Sub ResetForRun()
        ProductionMode = True
        DebugDiagnosticsMode = False
        DimLaboratoryMode = False
    End Sub

    Friend Shared Sub ApplyFromJob(config As JobConfiguration, runExclusiveDimLab As Boolean)
        DimLaboratoryMode = runExclusiveDimLab
        ProductionMode = Not runExclusiveDimLab
        If config Is Nothing Then Return
        DebugDiagnosticsMode =
            config.DetailedLog OrElse
            config.DebugTemplatesInspection OrElse
            config.ExperimentalDraftGeometryDiagnostics OrElse
            config.EnablePmiRetrievalProbe OrElse
            config.ExperimentalPmiProjectionDiagnostics OrElse
            config.ExperimentalCreatePMIModelViewIfMissing OrElse
            config.ExperimentalProbeCreatePMIModelViewOnly OrElse
            config.ExperimentalPmiTryAddPMIModelViewView OrElse
            config.ExperimentalPmiSyncPMIModelViewOrientationBeforeDraft OrElse
            False
        If runExclusiveDimLab Then DebugDiagnosticsMode = True
    End Sub

    Friend Shared Function FormatFlagsLog() As String
        Return String.Format(CultureInfo.InvariantCulture,
            "[ENGINE][MODE] ProductionMode={0} DebugDiagnosticsMode={1} DimLaboratoryMode={2}",
            ProductionMode, DebugDiagnosticsMode, DimLaboratoryMode)
    End Function
End Class
