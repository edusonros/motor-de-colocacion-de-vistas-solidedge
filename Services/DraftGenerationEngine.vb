Option Strict Off

Imports System.Collections.Generic
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports SolidEdgeAssembly
Imports SolidEdgeDraft
Imports SolidEdgeFramework
Imports Extraer_dft_dxf_flatdxf.Services.Dimensioning.Labs
Imports IO = System.IO
Imports IOPath = System.IO.Path
Imports Microsoft.VisualBasic

Public Class EngineProgressInfo
    Public Property Percent As Integer
    Public Property Status As String
End Class

Public Class EngineRunResult
    Public Property Success As Boolean
    Public Property ProcessedCount As Integer
    Public Property ErrorCount As Integer
    Public Property DraftCreatedCount As Integer
    Public Property PdfCreatedCount As Integer
    Public Property DxfCreatedCount As Integer
    Public Property FlatDxfCreatedCount As Integer
    Public Property SkippedCount As Integer
    ''' <summary>Ruta guardada del DFT de referencia DIMLAB (*_DIMLAB_REF_TEST.dft).</summary>
    Public Property DimLabReferenceDftFullPath As String
    ''' <summary>Botón LAB pulsado pero flags efectivos impidieron ejecutar DIMLAB.</summary>
    Public Property DimLabRunAbortedMisconfigured As Boolean
    ''' <summary>Si true, no se llama app.Quit al final para no cerrar el DFT DIMLAB abierto en revisión.</summary>
    Public Property KeepSolidEdgeOpenForDimLab As Boolean
    ''' <summary>Última ruta absoluta de PDF exportado correctamente en la corrida (p. ej. vista previa en UI).</summary>
    Public Property LastExportedPdfFullPath As String
End Class

Public Class DraftGenerationEngine
    Private Shared _stopRequested As Boolean = False
    Private Shared _pauseRequested As Boolean = False
    Private ReadOnly _logger As Logger
    Private ReadOnly _progress As Action(Of EngineProgressInfo)
    Private _diagnostics As RunDiagnosticsContext = Nothing

    Private ReadOnly _excludeKeywords As String() = {
        "skf", "nut", "2026", "2026_02", "screw", "duin", "iso", "bolt", "whaser", "washer",
        "snl", "sleeve", "22210", "22211", "22212", "fnl", "motor", "prensa", "estopada", "tornillo",
        "tuerca", "arandela", "fag"
    }

    Private Class FlatError
        Public Property FilePath As String
        Public Property StepName As String
        Public Property Message As String
    End Class

    Public Sub New(logger As Logger, progressReporter As Action(Of EngineProgressInfo))
        _logger = logger
        _progress = progressReporter
    End Sub

    Public Shared Sub RequestStop()
        _stopRequested = True
    End Sub

    Public Shared Sub RequestPause()
        _pauseRequested = True
    End Sub

    Public Shared Sub RequestResume()
        _pauseRequested = False
    End Sub

    Public Shared Function IsPaused() As Boolean
        Return _pauseRequested
    End Function

    Private Sub WaitIfPausedOrStopped()
        Do While _pauseRequested AndAlso Not _stopRequested
            Threading.Thread.Sleep(250)
        Loop
    End Sub

    Public Function Run(config As JobConfiguration) As EngineRunResult
        _stopRequested = False
        _pauseRequested = False
        Dim result As New EngineRunResult()
        Dim app As SolidEdgeFramework.Application = Nothing
        Dim appCreatedByUs As Boolean = False
        Dim flatErrors As New List(Of FlatError)()

        LogUtil.ResetStepCounter()
        LogUtil.SetLogSink(Sub(m) _logger.LogRaw(m))

        Try
            OleMessageFilter.Register()

            Try
                _logger.Log("[BOOT][EXE_PATH] " & Assembly.GetExecutingAssembly().Location)
                _logger.Log("[BOOT][CURRENT_DIR] " & System.Environment.CurrentDirectory)
                _logger.Log("[BOOT][STARTUP_PATH] " & System.Windows.Forms.Application.StartupPath)
            Catch bootEx As Exception
                _logger.Log("[BOOT][WARN] " & bootEx.Message)
            End Try
            _logger.Log("[BOOT][OUTPUT_ROOT] " & If(config?.OutputFolder, ""))

            ' Diagnósticos siempre en repositorio\docs\logs (run_log, com_errors; audit/geometry por DFT en ProcessModel).
            _diagnostics = RunDiagnosticsContext.Initialize(If(config?.OutputFolder, ""), _logger)

            Report(1, "Conectando a Solid Edge...")
            If Not ConnectSolidEdge(config, app, appCreatedByUs) Then
                Throw New Exception("No fue posible conectar con Solid Edge.")
            End If

            Dim inputKind As SourceFileKind = config.DetectInputKind()
            config.LastDetectedSourceKind = inputKind
            _logger.Log($"Leyendo archivo origen: {config.InputFile}")
            _logger.Log($"Tipo detectado: {inputKind}")
            _logger.Log($"Opciones -> DFT={config.CreateDraft}, PDF={config.CreatePdf}, DXF_DRAFT={config.CreateDxfFromDraft}, DXF_FLAT={config.CreateFlatDxf}")
            _logger.Log($"Opciones -> InsertProps={config.InsertPropertiesInTitleBlock}, SourceMode={config.TitleBlockPropertySourceMode}")
            _logger.Log($"[TITLEBLOCK] Mode={config.TitleBlockPropertySourceMode}")
            _logger.Log($"[MOTOR][CONFIG] MotorPhase={config.MotorPhase}")

            Dim outDftDir As String = IOPath.Combine(config.OutputFolder, "DFT")
            Dim outDxfDir As String = IOPath.Combine(config.OutputFolder, "DXF")
            Dim outPdfDir As String = IOPath.Combine(config.OutputFolder, "PDF de DFT")
            EnsureDir(outDftDir)
            EnsureDir(outDxfDir)
            EnsureDir(outPdfDir)

            Dim dftTemplates As String() = BuildTemplateList(config)
            Dim needsDftTemplatesForNewDraft As Boolean =
                (config.MotorPhase = DraftMotorPhase.FullSequence OrElse config.MotorPhase = DraftMotorPhase.ViewGeneration) AndAlso
                (config.CreateDraft OrElse config.CreatePdf OrElse config.CreateDxfFromDraft)
            If needsDftTemplatesForNewDraft Then
                If dftTemplates Is Nothing OrElse dftTemplates.Length = 0 Then
                    Throw New Exception("No hay templates A4/A3/A2 válidos para crear DFT/PDF/DXF.")
                End If
                ' Inspección profunda de plantillas solo con diagnóstico explícito (no en modo normal).
                If config.DebugTemplatesInspection Then
                    Dim templateList = dftTemplates.Distinct(StringComparer.OrdinalIgnoreCase).ToList()
                    For Each tpl In templateList
                        TemplateInspectionCache.LogTemplatePropertyInspectWithDiskCache(tpl, _logger)
                    Next
                End If
            End If
            If config.CreateDxfFromDraft AndAlso String.IsNullOrWhiteSpace(config.TemplateDxf) Then
                Throw New Exception("Template DXF limpio obligatorio para DXF desde Draft.")
            End If

            Dim targets As New List(Of String)()
            If inputKind = SourceFileKind.AssemblyFile Then
                Report(5, "Leyendo componentes del ASM...")
                If config.UseSelectedComponents Then
                    targets = If(config.SelectedComponentPaths, New List(Of String)()).
                        Distinct(StringComparer.OrdinalIgnoreCase).ToList()
                    _logger.Log($"Selección manual ASM activa: {targets.Count} componentes elegidos por el usuario.")
                Else
                    targets = ResolveAssemblyTargets(app, config.InputFile, config.ProcessRepeatedComponentsOnce, config)
                    If targets Is Nothing Then targets = New List(Of String)()
                    If targets.Count = 0 Then
                        _logger.Log("[ASM][FALLBACK] ResolveAssemblyTargets devolvió 0; usando AssemblyComponentService.LoadAssemblyComponentItems")
                        Try
                            Dim items = AssemblyComponentService.LoadAssemblyComponentItems(
                                config.InputFile,
                                config.ProcessRepeatedComponentsOnce,
                                showSolidEdge:=False,
                                logger:=_logger,
                                progress:=Nothing)
                            If items IsNot Nothing AndAlso items.Count > 0 Then
                                targets = items.
                                    Where(Function(it) it IsNot Nothing AndAlso
                                                         (String.Equals(it.Kind, "PAR", StringComparison.OrdinalIgnoreCase) OrElse
                                                          String.Equals(it.Kind, "PSM", StringComparison.OrdinalIgnoreCase)) AndAlso
                                                         Not String.IsNullOrWhiteSpace(it.FullPath)).
                                    Select(Function(it) it.FullPath).
                                    Distinct(StringComparer.OrdinalIgnoreCase).
                                    ToList()
                                _logger.Log($"[ASM][FALLBACK] Items={items.Count} -> targets(PAR/PSM)={targets.Count}")
                            End If
                        Catch exFallback As Exception
                            _logger.LogException("[ASM][FALLBACK]", exFallback)
                        End Try
                    End If
                End If
            Else
                targets.Add(config.InputFile)
            End If

            targets = ExpandSelectedTargets(app, targets, config.ProcessRepeatedComponentsOnce, config.UseSelectedComponents, config)

            Dim runAsmOverview As Boolean =
                inputKind = SourceFileKind.AssemblyFile AndAlso
                IO.File.Exists(config.InputFile) AndAlso
                (
                    config.CreateDraft OrElse config.CreatePdf OrElse config.CreateDxfFromDraft OrElse
                    config.MotorPhase = DraftMotorPhase.MetadataManagement OrElse
                    config.MotorPhase = DraftMotorPhase.Dimensioning
                )

            ' Respetar la selección manual de componentes:
            '   - Si el usuario está usando selección manual (UseSelectedComponents=True), el "overview"
            '     del ASM completo solo se ejecuta si el propio archivo ASM raíz fue marcado en la lista
            '     (lo añade la UI a SelectedComponentPaths). En cualquier otro caso (ningún componente
            '     marcado, o solo algunas piezas marcadas), NO se procesa el ASM completo.
            If runAsmOverview AndAlso config.UseSelectedComponents Then
                Dim asmRootMarked As Boolean = False
                If config.SelectedComponentPaths IsNot Nothing Then
                    asmRootMarked = config.SelectedComponentPaths.
                        Any(Function(p) Not String.IsNullOrWhiteSpace(p) AndAlso
                                          String.Equals(p, config.InputFile, StringComparison.OrdinalIgnoreCase))
                End If
                If Not asmRootMarked Then
                    _logger.Log("[ASM_OVERVIEW][SKIP] reason=manual_selection_active asm_root_not_checked")
                    runAsmOverview = False
                End If
            End If

            If targets.Count = 0 AndAlso Not runAsmOverview Then
                If config.UseSelectedComponents Then
                    _logger.Log("[ASM][SELECTION] 0 componentes marcados; nada que procesar (selección manual estricta).")
                Else
                    _logger.Log("No hay modelos para procesar.")
                End If
                result.Success = True
                Return result
            End If

            Dim totalJobs As Integer = targets.Count + If(runAsmOverview, 1, 0)
            Dim i As Integer = 0

            If runAsmOverview Then
                WaitIfPausedOrStopped()
                If _stopRequested Then
                    _logger.Log("[ENG][STOP] requested=True stage=before_asm_overview")
                Else
                    Report(0, $"Procesando 1/{totalJobs} - {IOPath.GetFileName(config.InputFile)}")
                    _logger.Log("[ASM_OVERVIEW][START] path=" & config.InputFile)
                    Try
                        ProcessModel(app, config.InputFile, config, dftTemplates, outDftDir, outDxfDir, outPdfDir, flatErrors, result, 1, totalJobs)
                        result.ProcessedCount += 1
                    Catch exAsm As Exception
                        result.ErrorCount += 1
                        _logger.LogException("ASM overview", exAsm)
                    End Try
                    Report(CInt(100.0 / Math.Max(1, totalJobs)), $"Finalizado 1/{totalJobs}")
                End If
            End If

            For Each modelPath In targets
                WaitIfPausedOrStopped()
                If _stopRequested Then
                    _logger.Log("[ENG][STOP] requested=True stage=before_process_model")
                    Exit For
                End If
                i += 1
                Dim jobIndex As Integer = i + If(runAsmOverview, 1, 0)
                Dim progressBase As Integer = CInt((jobIndex - 1) * 100.0 / Math.Max(1, totalJobs))
                Report(progressBase, $"Procesando {jobIndex}/{totalJobs} - {IOPath.GetFileName(modelPath)}")
                Try
                    ProcessModel(app, modelPath, config, dftTemplates, outDftDir, outDxfDir, outPdfDir, flatErrors, result, jobIndex, totalJobs)
                    result.ProcessedCount += 1
                Catch exModel As Exception
                    result.ErrorCount += 1
                    _logger.LogException($"Modelo {IOPath.GetFileName(modelPath)}", exModel)
                End Try
                Report(CInt(jobIndex * 100.0 / Math.Max(1, totalJobs)), $"Finalizado {jobIndex}/{totalJobs}")
            Next

            If flatErrors.Count > 0 Then
                WriteFlatErrorsFile(outDxfDir, flatErrors)
            End If

            result.Success = (result.ErrorCount = 0)
            _logger.Log($"RESUMEN -> Procesados={result.ProcessedCount}, OK={result.ProcessedCount - result.ErrorCount}, Error={result.ErrorCount}, DFT={result.DraftCreatedCount}, PDF={result.PdfCreatedCount}, DXF={result.DxfCreatedCount}, FLAT={result.FlatDxfCreatedCount}, Omitidos={result.SkippedCount}")
            If config.OpenOutputFolderWhenDone Then
                Try
                    Process.Start("explorer.exe", config.OutputFolder)
                Catch exOpen As Exception
                    _logger.LogException("Abrir carpeta de salida", exOpen)
                End Try
            End If

            Return result

        Catch ex As Exception
            result.Success = False
            result.ErrorCount += 1
            _logger.LogException("Run()", ex)
            Return result

        Finally
            ' Volcado final del log + resumen de rutas de diagnósticos (run_log/audit/geometry/com_errors).
            Try
                If _diagnostics IsNot Nothing Then _diagnostics.Finish()
            Catch exDiag As Exception
                Try : _logger.LogException("RunDiagnosticsContext.Finish", exDiag) : Catch : End Try
            End Try

            Try
                If app IsNot Nothing AndAlso appCreatedByUs Then
                    If result.KeepSolidEdgeOpenForDimLab Then
                        _logger.Log("[DIMLAB][KEEP_SE] Omitiendo Quit(): instancia creada por la app pero se deja abierta para revisar el DFT DIMLAB.")
                    Else
                        app.Quit()
                        _logger.Log("Solid Edge cerrado (instancia creada por la aplicación).")
                    End If
                ElseIf app IsNot Nothing Then
                    _logger.Log("Solid Edge permanece abierto (instancia preexistente).")
                End If
            Catch exQuit As Exception
                _logger.LogException("Cerrar Solid Edge", exQuit)
            End Try
            TryReleaseComObject(app)
            Try : OleMessageFilter.Revoke() : Catch : End Try
        End Try
    End Function

    ''' <summary>Acotación automática principal, DIMLAB exclusivo y laboratorios auxiliares (mismo comportamiento que en la rama CREATE DFT).</summary>
    ''' <returns>False si el flujo debe abortarse (ej. DIMLAB solicitado desde UI pero runLab efectivo fue False).</returns>
    Private Function ApplyDimensioningAndLabsBlock(
        app As SolidEdgeFramework.Application,
        dftDoc As DraftDocument,
        modelPath As String,
        baseName As String,
        outDft As String,
        config As JobConfiguration,
        runLab As Boolean,
        isAssemblyDraft As Boolean,
        mainDrawingView As DrawingView,
        runResult As EngineRunResult) As Boolean

        Dim normCfg As DimensioningNormConfig = If(config.DimensioningNormConfig Is Nothing,
            DimensioningNormConfig.DefaultUneLegacyConfig(), config.DimensioningNormConfig)
        normCfg.EnableSesdkPostDimensionIntrospection =
            normCfg.EnableSesdkPostDimensionIntrospection OrElse config.EnableSesdkPostDimensionIntrospection
        If config.PreferSweepAllDrawingDimensions Then
            normCfg.DimensionCreationMode = DimensioningNormConfig.ModeSweepAll
            normCfg.KeepIntentionalDuplicateDimensions = False
            normCfg.MaxTotalDimensionsTarget = Math.Max(normCfg.MaxTotalDimensionsTarget, 24)
            normCfg.MaxLinearDimensionsTarget = Math.Max(normCfg.MaxLinearDimensionsTarget, 18)
            normCfg.MaxRadialDimensionsTarget = Math.Max(normCfg.MaxRadialDimensionsTarget, 8)
            _logger.Log("[DIM][NORM] PreferSweepAllDrawingDimensions=True → SweepAllEntities keep_intentional_dupes=False caps>=24/18/8 bandas moderadas (~mitad densidad)")
        End If
        normCfg.SuppressDimensionTrackDistanceSpacing = config.SuppressDimensionTrackDistanceSpacing
        normCfg.EnableKeypointValueDuplicateCleanup = config.EnableKeypointValueDuplicateCleanup
        _logger.Log("[DIM][NORM] suppress_track_spacing=" & normCfg.SuppressDimensionTrackDistanceSpacing.ToString() &
                    " keypoint_value_dedup=" & normCfg.EnableKeypointValueDuplicateCleanup.ToString())

        Dim protectedZones As IList(Of ProtectedZone2D) = Nothing
        If dftDoc IsNot Nothing AndAlso normCfg.PrepareExistingPartsListTop Then
            Try
                Dim shPl As Sheet = Nothing
                Try : shPl = dftDoc.ActiveSheet : Catch : End Try
                If shPl IsNot Nothing Then
                    Dim sheetViews As New List(Of DrawingView)()
                    Try
                        Dim nDv As Integer = shPl.DrawingViews.Count
                        For idv As Integer = 1 To nDv
                            Try
                                sheetViews.Add(CType(shPl.DrawingViews.Item(idv), DrawingView))
                            Catch
                            End Try
                        Next
                    Catch
                    End Try
                    Dim plZone As ProtectedZone2D = PartsListSuperiorService.PrepararPartsListSuperiorExistente(
                        dftDoc, shPl, modelPath, normCfg, Sub(m) _logger.Log(m), sheetViews)
                    protectedZones = PartsListSuperiorService.DetectarZonasProtegidas(shPl, plZone, normCfg, Sub(m) _logger.Log(m))
                    PartsListSuperiorService.TryNudgeViewsOutsideZone(shPl, plZone, Sub(m) _logger.Log(m))
                End If
            Catch exPl As Exception
                _logger.Log("[PARTSLIST][ERR] " & exPl.Message)
            End Try
        End If

        Dim dimLabForensicInteractive As Boolean = runLab AndAlso config.EnableDimLabInteractivePause
        If runLab Then
            Try
                app.Visible = True
                app.DisplayAlerts = True
            Catch exVis As Exception
                _logger.Log("[DIMLAB][INTERACTIVE] WARN set Visible/DisplayAlerts: " & exVis.Message)
            End Try
            _logger.Log("[DIMLAB][INTERACTIVE] app.Visible=True DisplayAlerts=True (modo laboratorio)")
            _logger.Log("[DIMLAB][EXPORT_POLICY] skip_pdf_dxf_flat=True keep_dft_open_until_review=True")
            Dim outLabHint As String = config.OutputFolder.Trim().TrimEnd(IO.Path.DirectorySeparatorChar, IO.Path.AltDirectorySeparatorChar)
            Dim outLeaf As String = ""
            Try
                outLeaf = IOPath.GetFileName(outLabHint)
            Catch
                outLeaf = ""
            End Try
            If Not String.Equals(outLeaf, "OUT_DIMLAB", StringComparison.OrdinalIgnoreCase) Then
                outLabHint = IO.Path.Combine(config.OutputFolder.Trim(), "OUT_DIMLAB")
            End If
            _logger.Log("[DIMLAB][LOG_ROOT_HINT] texto DimLab=" & outLabHint)
        End If
        If dimLabForensicInteractive Then
            _logger.Log("[DIMLAB][INTERACTIVE] forensic_MsgBox_PAUSE=True")
            _logger.Log("[DIMLAB][INTERACTIVE] skipExportClose=True")
        End If
        Dim runProductionDvRefClean As Boolean = (dftDoc IsNot Nothing AndAlso config.EnableProductionDvRefCleanEngine AndAlso Not runLab)
        Dim runMainDimensioning As Boolean = (dftDoc IsNot Nothing AndAlso config.EnableAutoDimensioning AndAlso Not runLab AndAlso Not runProductionDvRefClean)
        ' Siempre en run_log: evita confundir corridas sin DIMLAB vs con laboratorio (no depender de ProductionMode).
        _logger.Log("[DIM][FLAGS] Config_EnableAutoDimensioning=" & config.EnableAutoDimensioning.ToString())
        _logger.Log("[DIM][FLAGS] Config_AutoDimensioningMotor=" & config.AutoDimensioningMotor.ToString())
        _logger.Log("[DIM][FLAGS] Config_EnableProductionDvRefCleanEngine=" & config.EnableProductionDvRefCleanEngine.ToString())
        _logger.Log("[DIM][FLAGS] Config_EnableDrawingViewDimensioningLab=" & config.EnableDrawingViewDimensioningLab.ToString())
        _logger.Log("[DIM][FLAGS] DimensionInsertionConfig_EnableDrawingViewDimensioningLab=" & DimensionInsertionConfig.EnableDrawingViewDimensioningLab.ToString())
        _logger.Log("[DIM][FLAGS] Effective_runLab=" & runLab.ToString())
        _logger.Log("[DIM][FLAGS] Effective_runProductionDvRefClean=" & runProductionDvRefClean.ToString())
        _logger.Log("[DIM][FLAGS] Effective_runMainDimensioning=" & runMainDimensioning.ToString())
        Dim decisionReason As String = If(runLab, "lab_mode_exclusive", If(runProductionDvRefClean, "production_dvref_clean_exclusive", "normal"))
        _logger.Log("[DIM][DECISION] runLab=" & runLab.ToString() & " runProductionDvRefClean=" & runProductionDvRefClean.ToString() & " runMainDimensioning=" & runMainDimensioning.ToString() & " reason=" & decisionReason)

        If config.RequestedDimLabFromDedicatedButton AndAlso Not runLab Then
            _logger.Log("[DIMLAB][ABORT] requested_lab_button_but_effective_runLab_false")
            runResult.DimLabRunAbortedMisconfigured = True
            runResult.ErrorCount += 1
            Return False
        End If

        If runLab Then
            _logger.Log("[DIMLAB][RUN] exclusive=True")
        End If

        If runLab AndAlso dftDoc IsNot Nothing Then
            Try
                Dim outLab As String = config.OutputFolder.Trim().TrimEnd(IO.Path.DirectorySeparatorChar, IO.Path.AltDirectorySeparatorChar)
                Dim leaf As String = ""
                Try
                    leaf = IOPath.GetFileName(outLab)
                Catch
                    leaf = ""
                End Try
                If Not String.Equals(leaf, "OUT_DIMLAB", StringComparison.OrdinalIgnoreCase) Then
                    outLab = IO.Path.Combine(config.OutputFolder.Trim(), "OUT_DIMLAB")
                End If
                Try
                    IO.Directory.CreateDirectory(outLab)
                Catch
                End Try
                DrawingViewDimensioningLab.Run(
                    app, dftDoc, Nothing, False, Sub(m) _logger.Log(m), outLab, dimLabForensicInteractive,
                    config.DimLabMode, config.EnableDimLabVisibleProbe, config.EnableDimLabAlternativePlacement,
                    config.EnableDimLabHorizontalControlInVerticalOnly,
                    config.DimLabKeepFailedDimensions,
                    config.DimLabCleanPreviousLabDimensions)
            Catch exLab As Exception
                _logger.Log("[DIMLAB][FATAL] " & exLab.ToString())
            End Try
            _logger.Log("[DIMLAB][EXIT_BEFORE_MAIN_ENGINE] true exclusive_block_main_dimensioning=True")
            _logger.Log("[DIM][GATE_SKIPPED] main_engine_blocked reason=lab_mode_exclusive")
        ElseIf runProductionDvRefClean Then
            _logger.Log("[DIM][PIPE] ProductionDvRefCleanDimensionEngine.Run exclusive=True")
            Try
                Dim dimProd As New DimensionLogger(_logger)
                ProductionDvRefCleanDimensionEngine.Run(app, dftDoc, dimProd)
            Catch exProd As Exception
                _logger.Log("[PRODDIM][FATAL] " & exProd.ToString())
            End Try
            _logger.Log("[DIM][GATE_SKIPPED] labs_and_secondary dimensioning_blocked reason=production_dvref_clean_exclusive")
        ElseIf runMainDimensioning Then
            Try
                Dim pipe As String
                Select Case config.AutoDimensioningMotor
                    Case AutoDimensioningMotorKind.LegacyV02IsolatedCopy
                        pipe = "[DIM][PIPE] DimensioningEngine → motor LegacyV02IsolatedCopy (LegacyV02DimensionMotorBridge)"
                    Case AutoDimensioningMotorKind.AlternatePlugIn
                        pipe = "[DIM][PIPE] DimensioningEngine → motor AlternatePlugIn (plugin alternativo, sin UniqueDv principal)"
                    Case Else
                        pipe = "[DIM][PIPE] DimensioningEngine → motor CurrentMainPipeline (UniqueDv + UNE/ISO 129)"
                End Select
                _logger.Log(pipe)
                DimensioningEngine.RunAutoDimensioning(dftDoc, mainDrawingView, _logger, normCfg, protectedZones)
            Catch exDim As Exception
                _logger.Log("[DIM][FATAL] " & exDim.ToString())
            End Try
        Else
            _logger.Log("[DIM][GATE] Acotado automático no ejecutado (deshabilitado, sin DFT o laboratorio sin documento).")
        End If

        Dim envDrop2D As String = ""
        Try : envDrop2D = System.Environment.GetEnvironmentVariable("DROP2D_LAB") : Catch : envDrop2D = "" : End Try
        Dim runDvMethodLab As Boolean = config.RunDVGeometryMethodDiscoveryLab
        Dim runDvDimLab As Boolean = config.RunDVGeometryDimensionPlacementLab
        Dim runDropSheetsLab As Boolean = config.RunDropCreatedSheetsDimensionLab
        Dim runDrop2D As Boolean = Not runDvMethodLab AndAlso Not runDvDimLab AndAlso Not runDropSheetsLab AndAlso (
            config.RunDropViewsTo2DModelLab OrElse
            String.Equals(envDrop2D, "1", StringComparison.OrdinalIgnoreCase) OrElse
            String.Equals(envDrop2D, "true", StringComparison.OrdinalIgnoreCase))
        If runDrop2D OrElse runDropSheetsLab OrElse runDvDimLab OrElse runDvMethodLab Then
            runMainDimensioning = False
        End If

        If runProductionDvRefClean Then
            If GenerationEngineRuntime.DebugDiagnosticsMode OrElse Not GenerationEngineRuntime.ProductionMode Then
                _logger.Log("[PRODDIM][SKIP] DV_METHODLAB / DV_DIMLAB / DROP labs — production exclusive")
            End If
            GoTo AfterDimensionLabs
        End If

        If runDvMethodLab AndAlso Not isAssemblyDraft AndAlso dftDoc IsNot Nothing Then
            Try
                _logger.Log("[DV_METHODLAB][RUN] enabled=True source=config")
                Dim prevAlerts As Boolean = True
                Try : prevAlerts = app.DisplayAlerts : Catch : End Try
                Try
                    app.DisplayAlerts = False
                    _logger.Log("[DV_METHODLAB][ALERTS] DisplayAlerts=False (suppress interactive dialogs during probing)")
                Catch exSetAlerts As Exception
                    _logger.Log("[DV_METHODLAB][ALERTS][WARN] " & exSetAlerts.Message)
                End Try
                Try
                    DVGeometryMethodDiscoveryLab.Run(app, dftDoc, Sub(m) _logger.Log(m), False)
                Finally
                    Try
                        app.DisplayAlerts = prevAlerts
                        _logger.Log("[DV_METHODLAB][ALERTS] DisplayAlerts restored=" & prevAlerts.ToString())
                    Catch
                    End Try
                End Try
            Catch exDvMethod As Exception
                _logger.Log("[DV_METHODLAB][FATAL] " & exDvMethod.ToString())
            End Try
        ElseIf runDvDimLab AndAlso Not isAssemblyDraft AndAlso dftDoc IsNot Nothing Then
            Try
                _logger.Log("[DV_DIMLAB][RUN] enabled=True source=config")
                DVGeometryDimensionPlacementLab.Run(app, dftDoc, Sub(m) _logger.Log(m), False)
            Catch exDv As Exception
                _logger.Log("[DV_DIMLAB][FATAL] " & exDv.ToString())
            End Try
        ElseIf runDropSheetsLab AndAlso Not isAssemblyDraft AndAlso dftDoc IsNot Nothing Then
            Try
                _logger.Log("[DROP_SHEETS][RUN] enabled=True source=config")
                If config.RunDropViewsTo2DModelLab OrElse
                    String.Equals(envDrop2D, "1", StringComparison.OrdinalIgnoreCase) OrElse
                    String.Equals(envDrop2D, "true", StringComparison.OrdinalIgnoreCase) Then
                    _logger.Log("[DROP_SHEETS][NOTE] omitiendo DROP2D (RunDropCreatedSheetsDimensionLab tiene prioridad; no se usa hoja 2D Model)")
                End If
                DropCreatedSheetsDimensionLab.Run(app, dftDoc, Sub(m) _logger.Log(m), config.DropCreatedSheetsDimensionLabDebugSave)
            Catch exSheets As Exception
                _logger.Log("[DROP_SHEETS][FATAL] " & exSheets.ToString())
            End Try
        ElseIf runDrop2D AndAlso Not isAssemblyDraft AndAlso dftDoc IsNot Nothing Then
            Try
                _logger.Log("[DROP2D][RUN] enabled=True source=" &
                            If(config.RunDropViewsTo2DModelLab, "config", "env:DROP2D_LAB"))
                DropViewsTo2DModelLab.Run(app, dftDoc, Sub(m) _logger.Log(m), False)
            Catch exDrop As Exception
                _logger.Log("[DROP2D][FATAL] " & exDrop.ToString())
            End Try
        ElseIf GenerationEngineRuntime.DebugDiagnosticsMode Then
            _logger.Log("[DV_METHODLAB][RUN] enabled=False")
            _logger.Log("[DV_DIMLAB][RUN] enabled=False")
            _logger.Log("[DROP2D][RUN] enabled=False")
            _logger.Log("[DROP_SHEETS][RUN] enabled=False")
        End If

AfterDimensionLabs:
        If Not runLab AndAlso GenerationEngineRuntime.ProductionMode Then
            WriteProductionSummary(modelPath, outDft)
        End If

        If Not runLab AndAlso Not isAssemblyDraft AndAlso config.RunUnitHorizontalExteriorDimensionTest AndAlso dftDoc IsNot Nothing Then
            Try
                UnitHorizontalExteriorDimensionTest.Run(
                    dftDoc,
                    Sub(m As String) _logger.Log(m)
                )
            Catch exUnit As Exception
                _logger.Log("[DIM][UNIT][ERR] " & exUnit.Message)
            End Try
        End If

        Return True
    End Function

    ''' <summary>
    ''' Escribe Material/L/H/D y fuentes PART_LIST desde el job sobre el DFT abierto.
    ''' Independiente de <see cref="JobConfiguration.InsertPropertiesInTitleBlock"/> (cajetín).
    ''' </summary>
    Private Sub TryApplyPartListMetadataAfterDraftDimensioning(
        app As SolidEdgeFramework.Application,
        dftDoc As DraftDocument,
        modelPath As String,
        config As JobConfiguration,
        runLab As Boolean)

        If runLab OrElse dftDoc Is Nothing Then Return
        Dim extLc As String = IOPath.GetExtension(modelPath).ToLowerInvariant()
        If extLc <> ".par" AndAlso extLc <> ".psm" Then Return

        Dim partSnap As DrawingMetadataInput = DrawingMetadataService.BuildFromJobConfiguration(config)
        Dim modelDoc As Object = Nothing
        Try
            modelDoc = ExecuteComWithRetry(Function() app.Documents.Open(modelPath), "[ENGINE][PARTLIST] Open modelo")
            SolidEdgeSessionVisibility.SuppressForegroundIfConfigured(app, config, _logger)
            If String.IsNullOrWhiteSpace(partSnap.LargoL) OrElse String.IsNullOrWhiteSpace(partSnap.AltoH) OrElse String.IsNullOrWhiteSpace(partSnap.DatoD) Then
                Dim lV As String = "", hV As String = "", dV As String = "", lhdSrc As String = ""
                If DrawingMetadataService.TryComputeLhdSameAsCalcButton(dftDoc, modelDoc, _logger, lV, hV, dV, lhdSrc) Then
                    If String.IsNullOrWhiteSpace(partSnap.LargoL) Then partSnap.LargoL = lV
                    If String.IsNullOrWhiteSpace(partSnap.AltoH) Then partSnap.AltoH = hV
                    If String.IsNullOrWhiteSpace(partSnap.DatoD) Then partSnap.DatoD = dV
                    _logger.Log("[ENGINE][PARTLIST][LHD] " & lhdSrc & " → L=" & lV & " H=" & hV & " D=" & dV)
                End If
            End If
            _logger.Log("[ENGINE][PARTLIST] Actualizando PART_LIST desde configuración del job (no depende de Insertar propiedades en cajetín).")
            DrawingMetadataService.ApplyPartListSourceProperties(modelDoc, dftDoc, partSnap, _logger)
        Catch ex As Exception
            _logger.LogException("[ENGINE][PARTLIST]", ex)
        Finally
            If modelDoc IsNot Nothing Then
                SolidEdgePropertyService.TryCloseComDocument(modelDoc, False)
                TryReleaseComObject(modelDoc)
            End If
            SolidEdgeSessionVisibility.SuppressForegroundIfConfigured(app, config, _logger)
        End Try
    End Sub

    Private Sub ProcessMetadataMotorOnly(
        app As SolidEdgeFramework.Application,
        modelPath As String,
        config As JobConfiguration,
        baseName As String,
        runLab As Boolean,
        outDftDir As String,
        runResult As EngineRunResult,
        itemIndex As Integer,
        totalItems As Integer)

        WaitIfPausedOrStopped()
        If _stopRequested Then Return

        Dim outDft As String = DraftStandaloneMotorPaths.ResolveStandaloneDraftFullPath(config, modelPath)
        _logger.Log("[DFT][FULLPATH] " & outDft)
        _logger.Log($"[MOTOR][GESTOR_META][{itemIndex}/{totalItems}] DFT esperado={outDft}")

        If Not IO.File.Exists(outDft) Then
            _logger.Log("[MOTOR][GESTOR_META][ERR] No existe DFT previo en carpeta DFT. " &
                        "Genere el DFT con ""Generador de vistas"" o ""GENERAR"" (misma salida y pieza).")
            runResult.ErrorCount += 1
            Return
        End If

        If Not config.InsertPropertiesInTitleBlock Then
            _logger.Log("[MOTOR][GESTOR_META][SKIP_CAJETIN] InsertPropertiesInTitleBlock=False (cajetín y ApplyPropertiesToSavedDraft omitidos; PART_LIST sí se actualiza si aplica).")
        End If

        If config.InsertPropertiesInTitleBlock AndAlso config.TitleBlockPropertySourceMode = TitleBlockPropertySource.FromModelLink Then
            Dim extIn As String = IOPath.GetExtension(modelPath).ToLowerInvariant()
            If extIn = ".dft" Then
                _logger.Log("[MOTOR][GESTOR_META][SKIP] Entrada .dft: no se abre ""modelo"" por ruta de entrada (solo DFT).")
            Else
                Dim modelPreUpdates As Integer = SolidEdgePropertyService.ApplyPropertiesToOpenModelDocument(app, modelPath, config, _logger)
                _logger.Log($"[PROPS][MODEL] Escritura previa al draft completada (motor metadatos). Campos={modelPreUpdates}")
            End If
        End If

        Dim partListSnap As DrawingMetadataInput = DrawingMetadataService.BuildFromJobConfiguration(config)

        Dim dftDoc As DraftDocument = Nothing
        Dim modelDoc As Object = Nothing
        Try
            dftDoc = ExecuteComWithRetry(Function() CType(app.Documents.Open(outDft), DraftDocument), "[GESTOR_META] Open DFT")
            SolidEdgeSessionVisibility.SuppressForegroundIfConfigured(app, config, _logger)
            If dftDoc Is Nothing Then
                runResult.ErrorCount += 1
                Return
            End If

            If IO.File.Exists(modelPath) Then
                Dim extLc As String = IOPath.GetExtension(modelPath).ToLowerInvariant()
                If extLc = ".par" OrElse extLc = ".psm" OrElse extLc = ".asm" Then
                    Try
                        modelDoc = ExecuteComWithRetry(Function() app.Documents.Open(modelPath), "[GESTOR_META] Open modelo (PART_LIST)")
                        SolidEdgeSessionVisibility.SuppressForegroundIfConfigured(app, config, _logger)
                        _logger.Log("[MOTOR][GESTOR_META][PARTSLIST] Documento modelo abierto para Custom.* coherentes con el enlace.")
                    Catch exMo As Exception
                        _logger.Log("[MOTOR][GESTOR_META][PARTSLIST][WARN] Modelo no abierto; se aplicará PART_LIST sólo sobre DFT: " & exMo.Message)
                        modelDoc = Nothing
                    End Try
                End If
            End If

            If String.IsNullOrWhiteSpace(partListSnap.LargoL) OrElse String.IsNullOrWhiteSpace(partListSnap.AltoH) OrElse String.IsNullOrWhiteSpace(partListSnap.DatoD) Then
                Dim lV As String = "", hV As String = "", dV As String = "", lhdSrc As String = ""
                If DrawingMetadataService.TryComputeLhdSameAsCalcButton(dftDoc, modelDoc, _logger, lV, hV, dV, lhdSrc) Then
                    If String.IsNullOrWhiteSpace(partListSnap.LargoL) Then partListSnap.LargoL = lV
                    If String.IsNullOrWhiteSpace(partListSnap.AltoH) Then partListSnap.AltoH = hV
                    If String.IsNullOrWhiteSpace(partListSnap.DatoD) Then partListSnap.DatoD = dV
                    _logger.Log("[MOTOR][GESTOR_META][LHD] " & lhdSrc & " → L=" & lV & " H=" & hV & " D=" & dV)
                End If
            End If

            _logger.Log("[MOTOR][GESTOR_META][PARTSLIST] Actualizando fuentes PART_LIST + Refresh Native PartsLists (misma ruta que la UI).")
            DrawingMetadataService.ApplyPartListSourceProperties(modelDoc, dftDoc, partListSnap, _logger)

            If modelDoc IsNot Nothing Then
                Try
                    ExecuteComWithRetry(Sub() CallByName(modelDoc, "Save", CallType.Method), "[GESTOR_META] Save modelo")
                Catch exSaveM As Exception
                    _logger.Log("[GESTOR_META][MODEL][SAVE][WARN] " & exSaveM.Message)
                End Try
                SolidEdgePropertyService.TryCloseComDocument(modelDoc, False)
                TryReleaseComObject(modelDoc)
                modelDoc = Nothing
            End If

            If config.InsertPropertiesInTitleBlock Then
                Dim syncCount As Integer = SolidEdgePropertyService.ApplyTitleBlockPropertyStrategy(dftDoc, modelPath, config, _logger)
                _logger.Log($"[MOTOR][GESTOR_META] Sincronización cajetín aplicada. Campos escritos~={syncCount}")
            Else
                _logger.Log("[MOTOR][GESTOR_META][SKIP_CAJETIN] Sin ApplyTitleBlockPropertyStrategy.")
            End If
            ExecuteComWithRetry(Sub() dftDoc.Save(), "[GESTOR_META] Save DFT")
            Try : dftDoc.Close(False) : Catch : End Try
            TryReleaseComObject(dftDoc)
            dftDoc = Nothing
            If config.InsertPropertiesInTitleBlock Then
                ExecuteComWithRetry(Sub() SolidEdgePropertyService.ApplyPropertiesToSavedDraft(app, outDft, config, _logger), "[GESTOR_META] Apply props saved DFT")
            End If
        Catch ex As Exception
            _logger.LogException("[MOTOR][GESTOR_META]", ex)
            runResult.ErrorCount += 1
        Finally
            If modelDoc IsNot Nothing Then
                SolidEdgePropertyService.TryCloseComDocument(modelDoc, False)
                TryReleaseComObject(modelDoc)
            End If
            If dftDoc IsNot Nothing Then
                Try : dftDoc.Close(False) : Catch : End Try
                TryReleaseComObject(dftDoc)
            End If
            SolidEdgeSessionVisibility.SuppressForegroundIfConfigured(app, config, _logger)
        End Try
    End Sub

    Private Sub ProcessDimensioningMotorOnly(
        app As SolidEdgeFramework.Application,
        modelPath As String,
        config As JobConfiguration,
        baseName As String,
        isAssemblyDraft As Boolean,
        outDftDir As String,
        outPdfDir As String,
        outDxfDir As String,
        runLab As Boolean,
        runResult As EngineRunResult,
        itemIndex As Integer,
        totalItems As Integer)

        WaitIfPausedOrStopped()
        If _stopRequested Then Return

        Dim outDft As String = DraftStandaloneMotorPaths.ResolveStandaloneDraftFullPath(config, modelPath)
        _logger.Log("[DFT][FULLPATH] " & outDft)
        _logger.Log($"[MOTOR][ACOTACION][{itemIndex}/{totalItems}] DFT esperado={outDft}")

        If Not IO.File.Exists(outDft) Then
            _logger.Log("[MOTOR][ACOTACION][ERR] No existe DFT previo en carpeta DFT. " &
                        "Genere el DFT con ""Generador de vistas"" o ""GENERAR"" (misma salida y pieza).")
            runResult.ErrorCount += 1
            Return
        End If

        Dim dftDoc As DraftDocument = Nothing
        Try
            If Not runLab Then DimensionProductionRunSummary.Reset()
            dftDoc = ExecuteComWithRetry(Function() CType(app.Documents.Open(outDft), DraftDocument), "[ACOTACION] Open DFT")
            SolidEdgeSessionVisibility.SuppressForegroundIfConfigured(app, config, _logger)
            If dftDoc Is Nothing Then
                runResult.ErrorCount += 1
                Return
            End If
            If Not runLab Then
                Try
                    Dim shv As Sheet = Nothing
                    Try : shv = dftDoc.ActiveSheet : Catch : End Try
                    If shv IsNot Nothing Then DimensionProductionRunSummary.ViewsPlanned = shv.DrawingViews.Count
                Catch
                End Try
            End If

            Dim mainDrawingView As DrawingView = Nothing
            If config.KeepSolidEdgeVisible Then FitDraftView(app, dftDoc)

            If Not ApplyDimensioningAndLabsBlock(app, dftDoc, modelPath, baseName, outDft, config, runLab, isAssemblyDraft, mainDrawingView, runResult) Then
                Try
                    If dftDoc IsNot Nothing Then dftDoc.Close(False)
                Catch
                End Try
                TryReleaseComObject(dftDoc)
                Return
            End If

            TryApplyPartListMetadataAfterDraftDimensioning(app, dftDoc, modelPath, config, runLab)

            Dim savedDraftReady As Boolean = False
            Dim diagCapturedForOpenDft As Boolean = False
            _logger.Log("[MOTOR][ACOTACION] Guardando DFT...")
            ExecuteComWithRetry(Sub() dftDoc.SaveAs(outDft), "[ACOTACION] Save DFT")
            savedDraftReady = IO.File.Exists(outDft)

            If _diagnostics IsNot Nothing AndAlso dftDoc IsNot Nothing Then
                Try
                    diagCapturedForOpenDft = _diagnostics.CaptureDiagnosticsForDft(dftDoc, baseName)
                Catch exDiagCapture As Exception
                    _logger.LogException("RunDiagnosticsContext.CaptureDiagnosticsForDft (motor acotación)", exDiagCapture)
                End Try
            End If

            If config.InsertPropertiesInTitleBlock AndAlso savedDraftReady AndAlso Not runLab Then
                Try : dftDoc.Close(False) : Catch : End Try
                TryReleaseComObject(dftDoc)
                dftDoc = Nothing
                ExecuteComWithRetry(Sub() SolidEdgePropertyService.ApplyPropertiesToSavedDraft(app, outDft, config, _logger), "[ACOTACION] Apply props saved DFT")
            End If

            If savedDraftReady Then runResult.DraftCreatedCount += 1
            If savedDraftReady AndAlso (runLab OrElse config.KeepDftOpenAfterRun) Then runResult.KeepSolidEdgeOpenForDimLab = True
            If runLab AndAlso savedDraftReady Then runResult.DimLabReferenceDftFullPath = outDft

            If Not diagCapturedForOpenDft AndAlso _diagnostics IsNot Nothing AndAlso dftDoc IsNot Nothing Then
                Try
                    diagCapturedForOpenDft = _diagnostics.CaptureDiagnosticsForDft(dftDoc, baseName)
                    If diagCapturedForOpenDft Then
                        _logger.Log("[DIAG][FALLBACK] audit+geometry (motor acotación, DFT abierto).")
                    End If
                Catch exDiagLate As Exception
                    _logger.LogException("RunDiagnosticsContext.CaptureDiagnosticsForDft (abierto)", exDiagLate)
                End Try
            End If

            Dim exportDoc As DraftDocument = dftDoc
            Dim exportDocOpened As Boolean = False
            Try
                If Not runLab AndAlso (config.CreatePdf OrElse config.CreateDxfFromDraft) AndAlso IO.File.Exists(outDft) Then
                    exportDoc = ExecuteComWithRetry(Function() CType(app.Documents.Open(outDft), DraftDocument), "[ACOTACION] Open saved DFT for export")
                    exportDocOpened = True
                End If

                If Not runLab AndAlso config.CreatePdf Then
                    SolidEdgePropertyService.ApplyDirectSummaryInfoToDraft(exportDoc, config, _logger)
                    Dim outPdf As String = GetOutputPath(outPdfDir, baseName, ".pdf", config.OverwriteExisting)
                    _logger.Log("Exportando PDF...")
                    ExecuteComWithRetry(Sub() ExportDraftToPdf(app, exportDoc, outPdf, config.KeepSolidEdgeVisible), "Export PDF")
                    _logger.Log($"PDF: {IOPath.GetFileName(outPdf)}")
                    runResult.PdfCreatedCount += 1
                    runResult.LastExportedPdfFullPath = outPdf
                End If

                If Not runLab AndAlso config.CreateDxfFromDraft Then
                    Dim outDxfDraft As String = GetOutputPath(outDxfDir, baseName, ".dxf", config.OverwriteExisting)
                    _logger.Log("Exportando DXF desde DFT...")
                    ExecuteComWithRetry(Sub() ExportDraftToDxf(exportDoc, outDxfDraft), "Export DXF Draft")
                    _logger.Log($"DXF(Draft): {IOPath.GetFileName(outDxfDraft)}")
                    runResult.DxfCreatedCount += 1
                End If
            Finally
                If exportDocOpened AndAlso exportDoc IsNot Nothing Then
                    Try : exportDoc.Close(False) : Catch : End Try
                    TryReleaseComObject(exportDoc)
                End If
            End Try

        Catch ex As Exception
            _logger.LogException("[MOTOR][ACOTACION]", ex)
            runResult.ErrorCount += 1
        Finally
            If Not runResult.KeepSolidEdgeOpenForDimLab Then
                Try
                    If dftDoc IsNot Nothing Then dftDoc.Close(False)
                Catch
                End Try
                TryReleaseComObject(dftDoc)
            ElseIf runLab Then
                _logger.Log("[DIMLAB][KEEP_OPEN] No se cierra el DFT automáticamente (motor acotación).")
            Else
                _logger.Log("[ENG][KEEP_OPEN] KeepDftOpenAfterRun=true (motor acotación).")
            End If
        End Try
    End Sub

    Private Sub ProcessModel(app As SolidEdgeFramework.Application,
                             modelPath As String,
                             config As JobConfiguration,
                             dftTemplates As String(),
                             outDftDir As String,
                             outDxfDir As String,
                             outPdfDir As String,
                             flatErrors As List(Of FlatError),
                             runResult As EngineRunResult,
                             itemIndex As Integer,
                             totalItems As Integer)
        If Not IO.File.Exists(modelPath) Then
            _logger.Log($"[SKIP] No existe archivo: {modelPath}")
            runResult.SkippedCount += 1
            Return
        End If

        WaitIfPausedOrStopped()
        If _stopRequested Then
            _logger.Log("[ENG][STOP] requested=True stage=process_model_enter")
            Return
        End If

        Dim ext As String = IOPath.GetExtension(modelPath).ToLowerInvariant()
        If ext = ".dft" AndAlso config.MotorPhase <> DraftMotorPhase.MetadataManagement AndAlso config.MotorPhase <> DraftMotorPhase.Dimensioning Then
            _logger.Log("[SKIP] Entrada .dft solo es valida con ""Gestor Metadatos DFT"" o ""Motor Acotación"". Use PAR/PSM/ASM para generar vistas.")
            runResult.SkippedCount += 1
            Return
        End If

        Dim isAssemblyDraft As Boolean = (ext = ".asm")
        Dim baseName As String = IOPath.GetFileNameWithoutExtension(modelPath)
        Dim runLab As Boolean = config.EnableDrawingViewDimensioningLab OrElse DimensionInsertionConfig.EnableDrawingViewDimensioningLab
        ' FORZADO TEMPORAL: ejecutar SIEMPRE DVGeometryMethodDiscoveryLab en modo exclusivo.
        Dim forceDvMethodDiscoveryLabOnly As Boolean = False
        If forceDvMethodDiscoveryLabOnly AndAlso Not isAssemblyDraft Then
            config.RunDVGeometryMethodDiscoveryLab = True
            config.RunDVGeometryDimensionPlacementLab = False
            config.RunDropCreatedSheetsDimensionLab = False
            config.RunDropViewsTo2DModelLab = False
            config.EnableAutoDimensioning = False
            config.EnableDrawingViewDimensioningLab = False
            DimensionInsertionConfig.EnableDrawingViewDimensioningLab = False
            runLab = False
            _logger.Log("[DV_METHODLAB][FORCE] enabled=True source=temporary_hard_force")
            _logger.Log("[DV_METHODLAB][FORCE] bypass_ui_config_gates=True")
        End If
        If isAssemblyDraft Then
            runLab = False
            _logger.Log("[ASM_OVERVIEW][GATE] labs_off=True (los laboratorios DIMLAB siguen desactivados para .asm)")
            _logger.Log("[ASM_OVERVIEW][ENGINE] Aplicando el mismo motor de vistas que para PAR/PSM (DraftGenerator.CreateAutomaticDraftFromModel + AddAssemblyView).")
            _logger.Log("[ASM_OVERVIEW][DIM] Acotado automático ACTIVO para el ensamblaje (mismo motor DimensioningEngine que para PAR/PSM).")
        End If

        GenerationEngineRuntime.ResetForRun()
        GenerationEngineRuntime.ApplyFromJob(config, runLab)
        _logger.Log(GenerationEngineRuntime.FormatFlagsLog())
        runResult.KeepSolidEdgeOpenForDimLab = False

        _logger.Log($"[{itemIndex}/{totalItems}] Procesando: {IOPath.GetFileName(modelPath)}")
        _logger.Log($"[MOTOR][PHASE] {config.MotorPhase}")
        Dim viewMotorOnly As Boolean = (config.MotorPhase = DraftMotorPhase.ViewGeneration)

        If config.MotorPhase = DraftMotorPhase.MetadataManagement Then
            ProcessMetadataMotorOnly(app, modelPath, config, baseName, runLab, outDftDir, runResult, itemIndex, totalItems)
            Return
        End If

        If config.MotorPhase = DraftMotorPhase.Dimensioning Then
            ProcessDimensioningMotorOnly(app, modelPath, config, baseName, isAssemblyDraft, outDftDir, outPdfDir, outDxfDir, runLab, runResult, itemIndex, totalItems)
            Return
        End If

        If config.InsertPropertiesInTitleBlock AndAlso config.TitleBlockPropertySourceMode = TitleBlockPropertySource.FromModelLink Then
            Dim modelPreUpdates As Integer = SolidEdgePropertyService.ApplyPropertiesToOpenModelDocument(app, modelPath, config, _logger)
            _logger.Log($"[PROPS][MODEL] Escritura previa al draft completada. Campos={modelPreUpdates}")
        End If

        Dim runDraftOutputBranch As Boolean =
            config.MotorPhase <> DraftMotorPhase.MetadataManagement AndAlso
            config.MotorPhase <> DraftMotorPhase.Dimensioning AndAlso
            (config.CreateDraft OrElse config.CreatePdf OrElse config.CreateDxfFromDraft)

        If runDraftOutputBranch Then
            Dim dftStem As String = If(runLab, DraftStandaloneMotorPaths.DimLabStemForDraft(baseName, config.DimLabMode), baseName)
            Dim outDft As String = GetOutputPath(outDftDir, dftStem, ".dft", config.OverwriteExisting)
            _logger.Log("[DFT][FULLPATH] " & outDft)
            If runLab Then _logger.Log("[DIMLAB][DFT][TARGET_FULLPATH] " & outDft)
            Dim dftDoc As DraftDocument = Nothing
            Try
                If Not runLab Then DimensionProductionRunSummary.Reset()
                _logger.Log("Creando draft...")
                Dim flatInserted As Boolean = False
                Dim mainDrawingView As DrawingView = Nothing
                dftDoc = ExecuteComWithRetry(
                    Function() CojonudoBestFit_Bueno.CreateDraftAlzadoPrimerDiedro(app, modelPath, dftTemplates, config.TemplateDxf, flatInserted, mainDrawingView, config.EnableSlotBBoxViewLayout, Sub(m) _logger.Log(m)),
                    "CreateDraft")
                SolidEdgeSessionVisibility.SuppressForegroundIfConfigured(app, config, _logger)
                If dftDoc IsNot Nothing Then
                    If Not runLab Then
                        Try
                            Dim shv As Sheet = Nothing
                            Try : shv = dftDoc.ActiveSheet : Catch : End Try
                            If shv IsNot Nothing Then DimensionProductionRunSummary.ViewsPlanned = shv.DrawingViews.Count
                        Catch
                        End Try
                    End If
                    Dim savedDraftReady As Boolean = False
                    Dim diagCapturedForOpenDft As Boolean = False
                    If config.InsertPropertiesInTitleBlock Then
                        If Not viewMotorOnly Then
                            Dim needPreSaveSync As Boolean =
                                (config.TitleBlockPropertySourceMode = TitleBlockPropertySource.FromModelLink) OrElse
                                (config.TitleBlockPropertySourceMode = TitleBlockPropertySource.FromDraft AndAlso Not config.CreateDraft)
                            If needPreSaveSync Then
                                Dim syncCount As Integer = SolidEdgePropertyService.ApplyTitleBlockPropertyStrategy(dftDoc, modelPath, config, _logger)
                                _logger.Log($"[PROPS] Sincronización pre-guardado completada. Campos escritos={syncCount}")
                            Else
                                _logger.Log("[PROPS] Modo FromDraft: se aplicará sobre DFT guardado (FileProperties) y fallback a documento abierto.")
                            End If
                        Else
                            _logger.Log("[MOTOR][VISTAS] Metadatos de cajetín omitidos (ejecutar ""Gestor Metadatos DFT"" si aplica).")
                        End If
                    End If

                        If config.KeepSolidEdgeVisible Then FitDraftView(app, dftDoc)

                    If Not viewMotorOnly Then
                        If Not ApplyDimensioningAndLabsBlock(app, dftDoc, modelPath, baseName, outDft, config, runLab, isAssemblyDraft, mainDrawingView, runResult) Then
                            Try
                                If dftDoc IsNot Nothing Then dftDoc.Close(False)
                            Catch
                            End Try
                            TryReleaseComObject(dftDoc)
                            Return
                        End If
                    Else
                        _logger.Log("[MOTOR][VISTAS] Acotación y laboratorios omitidos.")
                    End If

                    TryApplyPartListMetadataAfterDraftDimensioning(app, dftDoc, modelPath, config, runLab)

                    If config.CreateDraft Then
                        _logger.Log("Guardando DFT...")
                        ExecuteComWithRetry(Sub() dftDoc.SaveAs(outDft), "Save DFT")
                        savedDraftReady = IO.File.Exists(outDft)

                        ' DIAGNÓSTICOS AUTOMÁTICOS: audit + geometry del DFT ANTES de cualquier cierre/reapertura.
                        ' El DFT NO se debe cerrar antes de generar estos dos archivos.
                        If _diagnostics IsNot Nothing AndAlso dftDoc IsNot Nothing Then
                            Try
                                diagCapturedForOpenDft = _diagnostics.CaptureDiagnosticsForDft(dftDoc, baseName)
                            Catch exDiagCapture As Exception
                                _logger.LogException("RunDiagnosticsContext.CaptureDiagnosticsForDft", exDiagCapture)
                            End Try
                        End If

                        ' Asegurar persistencia/refresco de propiedades en el DFT ya guardado cuando aplica.
                        If Not viewMotorOnly AndAlso config.InsertPropertiesInTitleBlock AndAlso savedDraftReady AndAlso Not runLab Then
                            Try
                                dftDoc.Close(False)
                            Catch
                            End Try
                            TryReleaseComObject(dftDoc)
                            dftDoc = Nothing
                            ExecuteComWithRetry(Sub() SolidEdgePropertyService.ApplyPropertiesToSavedDraft(app, outDft, config, _logger), "Apply props saved DFT")
                        End If
                        _logger.Log($"DFT: {IOPath.GetFileName(outDft)}")
                        runResult.DraftCreatedCount += 1
                        If savedDraftReady AndAlso (runLab OrElse config.KeepDftOpenAfterRun) Then
                            runResult.KeepSolidEdgeOpenForDimLab = True
                        End If
                        If runLab AndAlso savedDraftReady Then
                            runResult.DimLabReferenceDftFullPath = outDft
                        End If
                    End If

                    If Not diagCapturedForOpenDft AndAlso _diagnostics IsNot Nothing AndAlso dftDoc IsNot Nothing Then
                        Try
                            diagCapturedForOpenDft = _diagnostics.CaptureDiagnosticsForDft(dftDoc, baseName)
                            If diagCapturedForOpenDft Then
                                _logger.Log("[DIAG][FALLBACK] audit+geometry capturados sobre DFT abierto (sin SaveAs previo o captura omitida).")
                            End If
                        Catch exDiagLate As Exception
                            _logger.LogException("RunDiagnosticsContext.CaptureDiagnosticsForDft (DFT abierto)", exDiagLate)
                        End Try
                    End If

                    Dim exportDoc As DraftDocument = dftDoc
                    Dim exportDocOpened As Boolean = False
                    Try
                        ' Tras SaveAs, algunos RCW de Draft pueden quedar desconectados.
                        ' Reabrimos el DFT guardado para exportaciones PDF/DXF cuando exista.
                        If Not runLab AndAlso config.CreateDraft AndAlso (config.CreatePdf OrElse config.CreateDxfFromDraft) AndAlso IO.File.Exists(outDft) Then
                            exportDoc = ExecuteComWithRetry(Function() CType(app.Documents.Open(outDft), DraftDocument), "Open saved DFT for export")
                            exportDocOpened = True
                        End If

                        If Not runLab AndAlso config.CreatePdf Then
                            SolidEdgePropertyService.ApplyDirectSummaryInfoToDraft(exportDoc, config, _logger)
                            Dim outPdf As String = GetOutputPath(outPdfDir, baseName, ".pdf", config.OverwriteExisting)
                            _logger.Log("Exportando PDF...")
                            ExecuteComWithRetry(Sub() ExportDraftToPdf(app, exportDoc, outPdf, config.KeepSolidEdgeVisible), "Export PDF")
                            _logger.Log($"PDF: {IOPath.GetFileName(outPdf)}")
                            runResult.PdfCreatedCount += 1
                            runResult.LastExportedPdfFullPath = outPdf
                        End If

                        If Not runLab AndAlso config.CreateDxfFromDraft Then
                            Dim outDxfDraft As String = GetOutputPath(outDxfDir, baseName, ".dxf", config.OverwriteExisting)
                            _logger.Log("Exportando DXF desde DFT...")
                            ExecuteComWithRetry(Sub() ExportDraftToDxf(exportDoc, outDxfDraft), "Export DXF Draft")
                            _logger.Log($"DXF(Draft): {IOPath.GetFileName(outDxfDraft)}")
                            runResult.DxfCreatedCount += 1
                        End If
                    Finally
                        If exportDocOpened AndAlso exportDoc IsNot Nothing Then
                            Try : exportDoc.Close(False) : Catch : End Try
                            TryReleaseComObject(exportDoc)
                        End If
                    End Try
                Else
                    If isAssemblyDraft Then
                        _logger.Log("[WARN][ASM_OVERVIEW] CreateAutomaticDraftFromModel devolvió Nothing para el ensamblaje. Posibles causas: AddAssemblyView falló (componentes no resueltos, sin geometría visible o configuración activa incompatible), no había vistas medibles, o el área útil de la plantilla es insuficiente. Revisa los mensajes [DRAFT] y [EX] previos.")
                    Else
                        _logger.Log("[WARN] CreateDraftAlzadoPrimerDiedro devolvió Nothing.")
                    End If
                End If
            Catch ex As Exception
                _logger.LogException("DFT/PDF/DXF Draft", ex)
            Finally
                If Not runResult.KeepSolidEdgeOpenForDimLab Then
                    Try
                        If dftDoc IsNot Nothing Then dftDoc.Close(False)
                        _logger.Log($"Documento DFT cerrado para: {IOPath.GetFileName(modelPath)}")
                    Catch
                    End Try
                    TryReleaseComObject(dftDoc)
                ElseIf runLab Then
                    _logger.Log("[DIMLAB][KEEP_OPEN] No se cierra el DFT automáticamente; revisar/grabar en Solid Edge.")
                Else
                    _logger.Log("[ENG][KEEP_OPEN] KeepDftOpenAfterRun=true: DFT dejado abierto para revisión.")
                End If
            End Try
        End If

        If ext = ".psm" AndAlso config.CreateFlatDxf AndAlso Not runLab Then
            Dim outFlat As String = GetOutputPath(outDxfDir, baseName & "_FLAT", ".dxf", config.OverwriteExisting)
            Try
                Dim errMsg As String = ""
                Dim ok As Boolean = ExecuteComWithRetry(
                    Function() FlatDxfExporter.ExportFlatDxf(app, modelPath, outFlat, errMsg, Sub(m) _logger.Log(m)),
                    "Export Flat DXF")
                If ok Then
                    _logger.Log($"DXF(Flat): {IOPath.GetFileName(outFlat)}")
                    runResult.FlatDxfCreatedCount += 1
                Else
                    flatErrors.Add(New FlatError With {
                        .FilePath = modelPath,
                        .StepName = "DXF: SaveAsFlatDXFEx",
                        .Message = If(String.IsNullOrWhiteSpace(errMsg), "Fallo desconocido en SaveAsFlatDXFEx.", errMsg)
                    })
                    _logger.Log($"[WARN] Flat no generado: {IOPath.GetFileName(modelPath)}")
                End If
            Catch ex As Exception
                _logger.LogException("DXF flat", ex)
            End Try
        ElseIf ext = ".par" AndAlso config.CreateFlatDxf Then
            _logger.Log("[INFO] DXF de chapa desarrollada no aplica para .par.")
            runResult.SkippedCount += 1
        End If

        SolidEdgeSessionVisibility.SuppressForegroundIfConfigured(app, config, _logger)
    End Sub

    Private Function BuildTemplateList(config As JobConfiguration) As String()
        Dim templates As New List(Of String)()
        Select Case config.PreferredFormat
            Case PreferredSheetFormat.A4
                If IO.File.Exists(config.TemplateA4) Then templates.Add(config.TemplateA4)
            Case PreferredSheetFormat.A3
                If IO.File.Exists(config.TemplateA3) Then templates.Add(config.TemplateA3)
            Case PreferredSheetFormat.A2
                If IO.File.Exists(config.TemplateA2) Then templates.Add(config.TemplateA2)
            Case Else
                If IO.File.Exists(config.TemplateA3) Then templates.Add(config.TemplateA3)
                If IO.File.Exists(config.TemplateA2) Then templates.Add(config.TemplateA2)
                If IO.File.Exists(config.TemplateA4) Then templates.Add(config.TemplateA4)
        End Select
        Return templates.ToArray()
    End Function

    Private Function ExpandSelectedTargets(app As SolidEdgeFramework.Application, rawTargets As List(Of String), uniqueOnly As Boolean, useManualSelection As Boolean, config As JobConfiguration) As List(Of String)
        Dim expanded As New List(Of String)()
        For Each p In rawTargets
            If String.IsNullOrWhiteSpace(p) Then Continue For
            Dim ext As String = IOPath.GetExtension(p).ToLowerInvariant()
            If ext = ".par" OrElse ext = ".psm" Then
                expanded.Add(p)
            ElseIf ext = ".dft" Then
                expanded.Add(p)
            ElseIf ext = ".asm" Then
                If useManualSelection Then
                    ' En selección manual los hijos del ASM ya aparecen aplanados en la UI;
                    ' expandirlo aquí ignoraría los desmarcados que el usuario hizo a nivel de pieza.
                    ' El ASM se trata aparte vía runAsmOverview cuando está marcado el ASM raíz.
                    _logger.Log($"[ASM][SELECTION] No se expande subensamblaje en modo manual: {IOPath.GetFileName(p)} (se respetan los desmarcados de sus piezas)")
                Else
                    _logger.Log($"Expandiendo ASM seleccionado: {IOPath.GetFileName(p)}")
                    expanded.AddRange(ResolveAssemblyTargets(app, p, uniqueOnly, config))
                End If
            Else
                _logger.Log($"[SKIP] Tipo no procesable en selección: {IOPath.GetFileName(p)}")
            End If
        Next
        If uniqueOnly Then
            expanded = expanded.Distinct(StringComparer.OrdinalIgnoreCase).ToList()
        End If
        Return expanded
    End Function

    Private Function ResolveAssemblyTargets(app As SolidEdgeFramework.Application,
                                            asmPath As String,
                                            uniqueOnly As Boolean,
                                            config As JobConfiguration) As List(Of String)
        Dim asmDoc As AssemblyDocument = Nothing
        Dim found As IDictionary(Of String, String)
        If uniqueOnly Then
            found = New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
        Else
            found = New Dictionary(Of String, String)(StringComparer.Ordinal)
        End If
        Dim countPar As Integer = 0
        Dim countPsm As Integer = 0

        Try
            OleMessageFilter.Register()
            asmDoc = ExecuteComWithRetry(Function() CType(app.Documents.Open(asmPath), AssemblyDocument), "Open ASM")
            ProcessOccurrencesRecursive(asmDoc.Occurrences, found, countPar, countPsm)
            _logger.Log($"TOTAL .par: {countPar} | TOTAL .psm: {countPsm} | únicos: {found.Count}")
            Return found.Keys.ToList()
        Catch ex As Exception
            _logger.LogException("ResolveAssemblyTargets", ex)
            Return New List(Of String)()
        Finally
            Try
                If asmDoc IsNot Nothing Then asmDoc.Close(False)
            Catch
            End Try
            TryReleaseComObject(asmDoc)
            SolidEdgeSessionVisibility.SuppressForegroundIfConfigured(app, config, _logger)
        End Try
    End Function

    Private Sub ProcessOccurrencesRecursive(occurs As Occurrences,
                                            found As IDictionary(Of String, String),
                                            ByRef countPar As Integer,
                                            ByRef countPsm As Integer)
        For Each occ As Occurrence In occurs
            Try
                Dim docObj As Object = occ.OccurrenceDocument
                Dim fullName As String = docObj.FullName
                Dim ext As String = IOPath.GetExtension(fullName).ToLowerInvariant()

                If ext = ".par" OrElse ext = ".psm" Then
                    If ShouldExcludeFile(fullName) Then Continue For
                    If Not found.ContainsKey(fullName) Then
                        found(fullName) = fullName
                        If ext = ".par" Then countPar += 1 Else countPsm += 1
                    End If
                ElseIf ext = ".asm" Then
                    Dim subAsm As AssemblyDocument = TryCast(docObj, AssemblyDocument)
                    If subAsm IsNot Nothing Then ProcessOccurrencesRecursive(subAsm.Occurrences, found, countPar, countPsm)
                End If
            Catch
            End Try
        Next
    End Sub

    Private Function ShouldExcludeFile(filePath As String) As Boolean
        Dim baseName As String = IOPath.GetFileNameWithoutExtension(filePath).ToLowerInvariant()
        Return _excludeKeywords.Any(Function(k) baseName.Contains(k))
    End Function

    Private Sub EnsureDir(dirPath As String)
        If Not IO.Directory.Exists(dirPath) Then IO.Directory.CreateDirectory(dirPath)
    End Sub

    Private Function GetOutputPath(dirPath As String, baseName As String, extWithDot As String, overwrite As Boolean) As String
        Dim candidate As String = IOPath.Combine(dirPath, baseName & extWithDot)
        If overwrite OrElse Not IO.File.Exists(candidate) Then Return candidate

        Dim i As Integer = 1
        Do
            candidate = IOPath.Combine(dirPath, $"{baseName}_{i:000}{extWithDot}")
            i += 1
        Loop While IO.File.Exists(candidate)
        Return candidate
    End Function

    Private Sub WriteFlatErrorsFile(outDxfDir As String, flatErrors As List(Of FlatError))
        Dim outPath As String = GetOutputPath(outDxfDir, "ERRORES_CHAPA_DESARROLLADA", ".txt", False)
        Using sw As New IO.StreamWriter(outPath, append:=False)
            sw.WriteLine("Errores al insertar/exportar chapa desarrollada")
            sw.WriteLine($"Fecha: {DateTime.Now:yyyy-MM-dd HH:mm:ss}")
            sw.WriteLine("")
            For Each e In flatErrors
                sw.WriteLine($"{e.StepName} | {e.FilePath} | {e.Message}")
            Next
        End Using
        _logger.Log($"Reporte de errores de chapa: {IOPath.GetFileName(outPath)}")
    End Sub

    Private Sub FitDraftView(app As SolidEdgeFramework.Application, draftDoc As DraftDocument)
        Try
            draftDoc.Activate()
            Dim winObj As Object = app.ActiveWindow
            Dim sw As SheetWindow = TryCast(winObj, SheetWindow)
            If sw IsNot Nothing Then sw.Fit()
        Catch
        End Try
    End Sub

    Private Sub ExportDraftToDxf(draftDoc As DraftDocument, dxfPath As String)
        Try
            Try
                draftDoc.SaveCopyAs(dxfPath)
            Catch
                draftDoc.SaveAs(dxfPath)
            End Try
        Catch ex As Exception
            _logger.LogException("ExportDraftToDxf", ex)
            Throw
        End Try
    End Sub

    Private Sub ExportDraftToPdf(app As SolidEdgeFramework.Application, draftDoc As DraftDocument, pdfPath As String, keepSolidEdgeVisible As Boolean)
        If draftDoc Is Nothing Then Return
        Dim exported As Boolean = False
        Dim firstEx As Exception = Nothing
        Try
            draftDoc.SaveAs(pdfPath, Nothing, False)
            exported = IO.File.Exists(pdfPath)
        Catch ex As Exception
            firstEx = ex
        End Try

        If Not exported Then
            Try
                draftDoc.SaveAs(pdfPath)
                exported = IO.File.Exists(pdfPath)
            Catch ex2 As Exception
                If firstEx Is Nothing Then firstEx = ex2
            End Try
        End If

        If Not exported AndAlso keepSolidEdgeVisible Then
            Try
                Dim prevVisible As Boolean = app.Visible
                app.Visible = True
                draftDoc.SaveAs(pdfPath, Nothing, False)
                exported = IO.File.Exists(pdfPath)
                app.Visible = prevVisible
            Catch
            End Try
        End If

        If Not exported Then
            If firstEx IsNot Nothing Then Throw firstEx
            Throw New Exception($"No se pudo generar PDF: {pdfPath}")
        End If
    End Sub

    Private Sub WriteProductionSummary(modelFullPath As String, dftOutputPath As String)
        _logger.Log("[SUMMARY][DOC] file=" & If(modelFullPath, ""))
        _logger.Log("[SUMMARY][VIEWS] count=" & DimensionProductionRunSummary.ViewsPlanned.ToString() &
                    " layout=" & If(DimensionProductionRunSummary.LayoutOk, "OK", "?"))
        _logger.Log("[SUMMARY][PARTSLIST] created=" & DimensionProductionRunSummary.PartsListCreated.ToString() &
                    " rows=" & DimensionProductionRunSummary.PartsListRows.ToString() &
                    " cols=" & DimensionProductionRunSummary.PartsListCols.ToString())
        _logger.Log("[SUMMARY][DIMS] created=" & DimensionProductionRunSummary.DimsCreated.ToString() &
                    " connected=" & DimensionProductionRunSummary.DimsConnectedOk.ToString() &
                    " visible=" & DimensionProductionRunSummary.DimsVisibleOk.ToString() &
                    " failed=" & DimensionProductionRunSummary.DimsFailed.ToString())
        Dim sty As String = If(String.IsNullOrWhiteSpace(DimensionProductionRunSummary.StyleAppliedName),
            "U3,5", DimensionProductionRunSummary.StyleAppliedName)
        _logger.Log("[SUMMARY][STYLE] " & sty)
        _logger.Log("[SUMMARY][OUTPUT] dft=" & If(dftOutputPath, ""))
        _logger.Log("[SUMMARY][RESULT] " & If(DimensionProductionRunSummary.ResultOk, "OK", "WARN"))
    End Sub

    Private Function ConnectSolidEdge(config As JobConfiguration,
                                      ByRef app As SolidEdgeFramework.Application,
                                      ByRef createdByUs As Boolean) As Boolean
        app = Nothing
        createdByUs = False
        _logger.Log("[SE][CONNECT] Reutilizar instancia activa (GetActiveObject)...")
        Try
            app = CType(Marshal.GetActiveObject("SolidEdge.Application"), SolidEdgeFramework.Application)
            _logger.Log("Solid Edge: instancia existente reutilizada.")
        Catch exAttach As Exception
            _logger.Log("[SE][CONNECT] Sin instancia activa (" & exAttach.GetType().Name & ": " & exAttach.Message & "). Creando proceso (puede tardar 30–90 s si Solid Edge no estaba abierto).")
            Dim t = Type.GetTypeFromProgID("SolidEdge.Application", throwOnError:=False)
            If t Is Nothing Then
                _logger.Log("[SE][CONNECT] ERROR: ProgID ""SolidEdge.Application"" no encontrado. ¿Solid Edge instalado y registrado COM?")
                Return False
            End If
            Try
                app = CType(Activator.CreateInstance(t), SolidEdgeFramework.Application)
                createdByUs = True
                _logger.Log("Solid Edge: nueva instancia creada por la aplicación.")
            Catch exNew As Exception
                _logger.LogException("ConnectSolidEdge CreateInstance", exNew)
                Return False
            End Try
        End Try
        If app Is Nothing Then Return False

        SolidEdgeSessionVisibility.ApplyApplicationVisibility(app, config, _logger)
        Return True
    End Function

    Private Sub Report(percent As Integer, status As String)
        If _progress Is Nothing Then Return
        _progress.Invoke(New EngineProgressInfo With {
            .Percent = Math.Max(0, Math.Min(100, percent)),
            .Status = status
        })
    End Sub

    ' Reintento COM para rechazos temporales RPC_E_CALL_REJECTED / SERVERCALL_RETRYLATER.
    Private Function ExecuteComWithRetry(Of T)(work As Func(Of T), operationName As String, Optional retries As Integer = 3, Optional waitMs As Integer = 120) As T
        Dim lastEx As Exception = Nothing
        For i As Integer = 1 To retries
            Try
                Return work.Invoke()
            Catch ex As Exception
                lastEx = ex
                Dim cex As COMException = TryCast(ex, COMException)
                Dim retryable As Boolean = (cex IsNot Nothing AndAlso (cex.ErrorCode = &H80010001 OrElse cex.ErrorCode = &H8001010A))
                If Not retryable OrElse i >= retries Then Throw
                _logger.Log($"[COM][RETRY] {operationName} intento {i}/{retries} HR=0x{cex.ErrorCode:X8}")
                Threading.Thread.Sleep(waitMs)
            End Try
        Next
        Throw lastEx
    End Function

    Private Sub ExecuteComWithRetry(work As Action, operationName As String, Optional retries As Integer = 3, Optional waitMs As Integer = 120)
        ExecuteComWithRetry(Of Boolean)(Function()
                                            work.Invoke()
                                            Return True
                                        End Function, operationName, retries, waitMs)
    End Sub

    Private Sub TryReleaseComObject(obj As Object)
        If obj Is Nothing Then Return
        Try
            If Marshal.IsComObject(obj) Then Marshal.ReleaseComObject(obj)
        Catch
        End Try
    End Sub

End Class
