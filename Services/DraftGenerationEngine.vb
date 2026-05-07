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
End Class

Public Class DraftGenerationEngine
    Private ReadOnly _logger As Logger
    Private ReadOnly _progress As Action(Of EngineProgressInfo)

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

    Public Function Run(config As JobConfiguration) As EngineRunResult
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

            Dim outDftDir As String = IOPath.Combine(config.OutputFolder, "DFT")
            Dim outDxfDir As String = IOPath.Combine(config.OutputFolder, "DXF")
            Dim outPdfDir As String = IOPath.Combine(config.OutputFolder, "PDF de DFT")
            EnsureDir(outDftDir)
            EnsureDir(outDxfDir)
            EnsureDir(outPdfDir)

            Dim dftTemplates As String() = BuildTemplateList(config)
            If config.CreateDraft OrElse config.CreatePdf OrElse config.CreateDxfFromDraft Then
                If dftTemplates Is Nothing OrElse dftTemplates.Length = 0 Then
                    Throw New Exception("No hay templates A4/A3/A2 válidos para crear DFT/PDF/DXF.")
                End If
                ' Inspección profunda de plantillas solo con diagnóstico explícito (no en modo normal).
                If config.DebugTemplatesInspection Then
                    Dim templateList = dftTemplates.Distinct(StringComparer.OrdinalIgnoreCase).ToList()
                    For Each tpl In templateList
                        SolidEdgePropertyService.LogAllPropertySetsFromFile(tpl, _logger)
                    Next
                End If
            End If
            If config.CreateDxfFromDraft AndAlso String.IsNullOrWhiteSpace(config.TemplateDxf) Then
                Throw New Exception("Template DXF limpio obligatorio para DXF desde Draft.")
            End If

            Dim targets As New List(Of String)()
            If inputKind = SourceFileKind.AssemblyFile Then
                Report(5, "Leyendo componentes del ASM...")
                If config.UseSelectedComponents AndAlso config.SelectedComponentPaths IsNot Nothing AndAlso config.SelectedComponentPaths.Count > 0 Then
                    targets = config.SelectedComponentPaths.Distinct(StringComparer.OrdinalIgnoreCase).ToList()
                    _logger.Log($"Selección manual activa: {targets.Count} componentes elegidos por el usuario.")
                Else
                    targets = ResolveAssemblyTargets(app, config.InputFile, config.ProcessRepeatedComponentsOnce)
                End If
            Else
                targets.Add(config.InputFile)
            End If

            targets = ExpandSelectedTargets(app, targets, config.ProcessRepeatedComponentsOnce)

            If targets.Count = 0 Then
                _logger.Log("No hay modelos para procesar.")
                result.Success = True
                Return result
            End If

            Dim i As Integer = 0
            For Each modelPath In targets
                i += 1
                Dim progressBase As Integer = CInt((i - 1) * 100.0 / targets.Count)
                Report(progressBase, $"Procesando {i}/{targets.Count} - {IOPath.GetFileName(modelPath)}")
                Try
                    ProcessModel(app, modelPath, config, dftTemplates, outDftDir, outDxfDir, outPdfDir, flatErrors, result, i, targets.Count)
                    result.ProcessedCount += 1
                Catch exModel As Exception
                    result.ErrorCount += 1
                    _logger.LogException($"Modelo {IOPath.GetFileName(modelPath)}", exModel)
                End Try
                Report(CInt(i * 100.0 / targets.Count), $"Finalizado {i}/{targets.Count}")
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

        Dim ext As String = IOPath.GetExtension(modelPath).ToLowerInvariant()
        Dim baseName As String = IOPath.GetFileNameWithoutExtension(modelPath)
        Dim runLab As Boolean = config.EnableDrawingViewDimensioningLab OrElse DimensionInsertionConfig.EnableDrawingViewDimensioningLab
        ' FORZADO TEMPORAL: ejecutar SIEMPRE DVGeometryMethodDiscoveryLab en modo exclusivo.
        Dim forceDvMethodDiscoveryLabOnly As Boolean = True
        If forceDvMethodDiscoveryLabOnly Then
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

        GenerationEngineRuntime.ResetForRun()
        GenerationEngineRuntime.ApplyFromJob(config, runLab)
        _logger.Log(GenerationEngineRuntime.FormatFlagsLog())
        runResult.KeepSolidEdgeOpenForDimLab = False

        _logger.Log($"[{itemIndex}/{totalItems}] Procesando: {IOPath.GetFileName(modelPath)}")

        If config.InsertPropertiesInTitleBlock AndAlso config.TitleBlockPropertySourceMode = TitleBlockPropertySource.FromModelLink Then
            Dim modelPreUpdates As Integer = SolidEdgePropertyService.ApplyPropertiesToOpenModelDocument(app, modelPath, config, _logger)
            _logger.Log($"[PROPS][MODEL] Escritura previa al draft completada. Campos={modelPreUpdates}")
        End If

        If config.CreateDraft OrElse config.CreatePdf OrElse config.CreateDxfFromDraft Then
            Dim dftStem As String = If(runLab, GetDimLabDraftStemPiece(baseName, config.DimLabMode), baseName)
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
                    If config.InsertPropertiesInTitleBlock Then
                        Dim needPreSaveSync As Boolean =
                            (config.TitleBlockPropertySourceMode = TitleBlockPropertySource.FromModelLink) OrElse
                            (config.TitleBlockPropertySourceMode = TitleBlockPropertySource.FromDraft AndAlso Not config.CreateDraft)
                        If needPreSaveSync Then
                            Dim syncCount As Integer = SolidEdgePropertyService.ApplyTitleBlockPropertyStrategy(dftDoc, modelPath, config, _logger)
                            _logger.Log($"[PROPS] Sincronización pre-guardado completada. Campos escritos={syncCount}")
                        Else
                            _logger.Log("[PROPS] Modo FromDraft: se aplicará sobre DFT guardado (FileProperties) y fallback a documento abierto.")
                        End If
                    End If

                        If config.KeepSolidEdgeVisible Then FitDraftView(app, dftDoc)

                    Dim normCfg As DimensioningNormConfig = If(config.DimensioningNormConfig Is Nothing,
                        DimensioningNormConfig.DefaultUneLegacyConfig(), config.DimensioningNormConfig)

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
                        _logger.Log("[DIMLAB][INTERACTIVE] app.Visible=True DisplayAlerts=True")
                        _logger.Log("[DIMLAB][EXPORT_POLICY] skip_pdf_dxf_flat=True keep_dft_open_until_review=True")
                        _logger.Log("[DIMLAB][LOG_ROOT_HINT] texto DimLab=" & IO.Path.Combine(config.OutputFolder, "OUT_DIMLAB"))
                    End If
                    If dimLabForensicInteractive Then
                        _logger.Log("[DIMLAB][INTERACTIVE] forensic_MsgBox_PAUSE=True")
                        _logger.Log("[DIMLAB][INTERACTIVE] skipExportClose=True")
                    End If
                    Dim runMainDimensioning As Boolean = (dftDoc IsNot Nothing AndAlso config.EnableAutoDimensioning AndAlso Not runLab)
                    If GenerationEngineRuntime.DebugDiagnosticsMode OrElse Not GenerationEngineRuntime.ProductionMode Then
                        _logger.Log("[DIM][FLAGS] Config_EnableAutoDimensioning=" & config.EnableAutoDimensioning.ToString())
                        _logger.Log("[DIM][FLAGS] Config_EnableDrawingViewDimensioningLab=" & config.EnableDrawingViewDimensioningLab.ToString())
                        _logger.Log("[DIM][FLAGS] DimensionInsertionConfig_EnableDrawingViewDimensioningLab=" & DimensionInsertionConfig.EnableDrawingViewDimensioningLab.ToString())
                        _logger.Log("[DIM][FLAGS] Effective_runLab=" & runLab.ToString())
                        _logger.Log("[DIM][FLAGS] Effective_runMainDimensioning=" & runMainDimensioning.ToString())
                        Dim decisionReason As String = If(runLab, "lab_mode_exclusive", "normal")
                        _logger.Log("[DIM][DECISION] runLab=" & runLab.ToString() & " runMainDimensioning=" & runMainDimensioning.ToString() & " reason=" & decisionReason)
                    End If

                    If config.RequestedDimLabFromDedicatedButton AndAlso Not runLab Then
                        _logger.Log("[DIMLAB][ABORT] requested_lab_button_but_effective_runLab_false")
                        runResult.DimLabRunAbortedMisconfigured = True
                        runResult.ErrorCount += 1
                        Try
                            If dftDoc IsNot Nothing Then dftDoc.Close(False)
                        Catch
                        End Try
                        TryReleaseComObject(dftDoc)
                        dftDoc = Nothing
                        Return
                    End If

                    If runLab AndAlso dftDoc IsNot Nothing Then
                        _logger.Log("[DIMLAB][RUN] exclusive=True")
                        Try
                            Dim outLab As String = IO.Path.Combine(config.OutputFolder, "OUT_DIMLAB")
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
                    ElseIf runMainDimensioning Then
                        Try
                            _logger.Log("[DIM][PIPE] DimensioningEngine.RunAutoDimensioning → motor DV*2d + UNE/ISO 129")
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

                    If runDvMethodLab AndAlso dftDoc IsNot Nothing Then
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
                    ElseIf runDvDimLab AndAlso dftDoc IsNot Nothing Then
                        Try
                            _logger.Log("[DV_DIMLAB][RUN] enabled=True source=config")
                            DVGeometryDimensionPlacementLab.Run(app, dftDoc, Sub(m) _logger.Log(m), False)
                        Catch exDv As Exception
                            _logger.Log("[DV_DIMLAB][FATAL] " & exDv.ToString())
                        End Try
                    ElseIf runDropSheetsLab AndAlso dftDoc IsNot Nothing Then
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
                    ElseIf runDrop2D AndAlso dftDoc IsNot Nothing Then
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

                    If Not runLab AndAlso GenerationEngineRuntime.ProductionMode Then
                        WriteProductionSummary(modelPath, outDft)
                    End If

                    If Not runLab AndAlso config.RunUnitHorizontalExteriorDimensionTest AndAlso dftDoc IsNot Nothing Then
                        Try
                            UnitHorizontalExteriorDimensionTest.Run(
                                dftDoc,
                                Sub(m As String) _logger.Log(m)
                            )
                        Catch exUnit As Exception
                            _logger.Log("[DIM][UNIT][ERR] " & exUnit.Message)
                        End Try
                    End If

                    If config.CreateDraft Then
                        _logger.Log("Guardando DFT...")
                        ExecuteComWithRetry(Sub() dftDoc.SaveAs(outDft), "Save DFT")
                        savedDraftReady = IO.File.Exists(outDft)

                        ' Asegurar persistencia/refresco de propiedades en el DFT ya guardado cuando aplica.
                        If config.InsertPropertiesInTitleBlock AndAlso savedDraftReady AndAlso Not runLab Then
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
                            ExecuteComWithRetry(Sub() ExportDraftToPdf(app, exportDoc, outPdf), "Export PDF")
                            _logger.Log($"PDF: {IOPath.GetFileName(outPdf)}")
                            runResult.PdfCreatedCount += 1
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
                    _logger.Log("[WARN] CreateDraftAlzadoPrimerDiedro devolvió Nothing.")
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
                Dim ok As Boolean = ExecuteComWithRetry(Function() FlatDxfExporter.ExportFlatDxf(app, modelPath, outFlat, errMsg), "Export Flat DXF")
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

    Private Function ExpandSelectedTargets(app As SolidEdgeFramework.Application, rawTargets As List(Of String), uniqueOnly As Boolean) As List(Of String)
        Dim expanded As New List(Of String)()
        For Each p In rawTargets
            If String.IsNullOrWhiteSpace(p) Then Continue For
            Dim ext As String = IOPath.GetExtension(p).ToLowerInvariant()
            If ext = ".par" OrElse ext = ".psm" Then
                expanded.Add(p)
            ElseIf ext = ".asm" Then
                _logger.Log($"Expandiendo ASM seleccionado: {IOPath.GetFileName(p)}")
                expanded.AddRange(ResolveAssemblyTargets(app, p, uniqueOnly))
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
                                            uniqueOnly As Boolean) As List(Of String)
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

    Private Sub ExportDraftToPdf(app As SolidEdgeFramework.Application, draftDoc As DraftDocument, pdfPath As String)
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

        If Not exported Then
            Try
                ' Último fallback: mostrar temporalmente Solid Edge y reintentar exportación.
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

        ' Producción estable: interfaz SIEMPRE visible (sin ocultar ni suprimir alertas).
        Try
            app.Visible = True
            app.DisplayAlerts = True
        Catch visEx As Exception
            _logger.Log("[SE][VISIBLE] fallo al establecer visibilidad: " & visEx.Message)
        End Try
        _logger.Log("[SE][VISIBLE] True")
        _logger.Log("[SE][DISPLAY_ALERTS] True")
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

    Private Shared Function GetDimLabDraftStemPiece(baseStem As String, mode As DimLabMode) As String
        Select Case mode
            Case DimLabMode.VerticalOnly
                Return baseStem & "_DIMLAB_REF_VERTICAL"
            Case DimLabMode.Full
                Return baseStem & "_DIMLAB_REF_FULL"
            Case DimLabMode.CleanFull
                Return baseStem & "_DIMLAB_REF_CLEANFULL"
            Case DimLabMode.CleanFullStrict
                Return baseStem & "_DIMLAB_REF_CLEANFULLSTRICT"
            Case Else
                Return baseStem & "_DIMLAB_REF_HORIZONTAL"
        End Select
    End Function
End Class
