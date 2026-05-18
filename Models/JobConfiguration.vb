Option Strict Off

Imports System
Imports Extraer_dft_dxf_flatdxf.Services.Dimensioning.Labs

Public Enum SourceFileKind
    Unknown = 0
    AssemblyFile = 1   ' .asm
    PartFile = 2       ' .par
    SheetMetalFile = 3 ' .psm
    ''' <summary>Borrador Solid Edge (.dft) como archivo de entrada (solo motores que operan sobre DFT existente).</summary>
    DraftFile = 4
End Enum

Public Enum PreferredSheetFormat
    Auto = 0
    A4 = 1
    A3 = 2
    A2 = 3
End Enum

Public Enum TitleBlockPropertySource
    FromModelLink = 0
    FromDraft = 1
End Enum

''' <summary>Origen del título mostrado en cajetín antes de escribir al modelo.</summary>
Public Enum TitleSourceMode
    Manual = 0
    AutoFromFileName = 1
End Enum

''' <summary>
''' Desacopla la ejecución en tres motores lógicos (misma implementación interna en <c>DraftGenerationEngine</c>, distinto encadenamiento).
''' </summary>
Public Enum DraftMotorPhase
    ''' <summary>Secuencia histórica: vistas → metadatos (según flags) → acotación (según flags).</summary>
    FullSequence = 0
    ''' <summary>Sólo creación de DFT, vistas, layout y piezas anexas (p.ej. PartsList); sin escritura de metadatos ni acotación.</summary>
    ViewGeneration = 1
    ''' <summary>Abre el DFT ya generado en la carpeta de salida y aplica metadatos/cajetín (requiere DFT previo).</summary>
    MetadataManagement = 2
    ''' <summary>Abre el DFT existente y ejecuta acotación / laboratorios según la UI (requiere DFT previo).</summary>
    Dimensioning = 3
End Enum

''' <summary>Implementación del acotado automático DV 2D: pipeline actual vs copia aislada vs enganche alternativo.</summary>
Public Enum AutoDimensioningMotorKind
    ''' <summary>Motor en <c>Services\Dimensioning</c> (UniqueDv + UNE129) vigente en este repositorio.</summary>
    CurrentMainPipeline = 0
    ''' <summary>Copia bajo <c>LegacyV02Dimensioning</c> (carpeta LegacyV02IsolatedMotor); tipos distintos del motor principal.</summary>
    LegacyV02IsolatedCopy = 1
    ''' <summary>Segundo motor propio (COM/cotas) vía factoría <c>DrawingViewAutoDimensioningMotorFactory</c>; no usa <c>UniqueDvAutoDimensioningEngine</c> del pipeline principal.</summary>
    AlternatePlugIn = 2
End Enum

Public Class JobConfiguration
    Public Property InputFile As String = ""
    Public Property OutputFolder As String = ""
    ''' <summary>Motor activo para esta ejecución (UI: GENERAR = <see cref="DraftMotorPhase.FullSequence"/>).</summary>
    Public Property MotorPhase As DraftMotorPhase = DraftMotorPhase.FullSequence

    Public Property TemplateA4 As String = ""
    Public Property TemplateA3 As String = ""
    Public Property TemplateA2 As String = ""
    Public Property TemplateDxf As String = ""

    Public Property CreateDraft As Boolean = True
    Public Property CreatePdf As Boolean = True
    Public Property CreateDxfFromDraft As Boolean = True
    Public Property CreateFlatDxf As Boolean = True

    Public Property OpenOutputFolderWhenDone As Boolean = True
    Public Property OverwriteExisting As Boolean = False
    Public Property ProcessRepeatedComponentsOnce As Boolean = True
    Public Property DetailedLog As Boolean = True
    ''' <summary>Inspección profunda de plantillas .dft (PropertySets) y volcado de enlaces del cajetín en cada DFT. Por defecto desactivado.</summary>
    Public Property DebugTemplatesInspection As Boolean = False
    Public Property KeepSolidEdgeVisible As Boolean = False
    ''' <summary>HWND de la ventana WinForms del generador; para devolver el foco tras operaciones COM si Solid Edge no debe estar al frente.</summary>
    Public Property OwnerWindowHandle As IntPtr = IntPtr.Zero
    Public Property InsertPropertiesInTitleBlock As Boolean = False
    Public Property TitleBlockPropertySourceMode As TitleBlockPropertySource = TitleBlockPropertySource.FromModelLink

    Public Property PreferredFormat As PreferredSheetFormat = PreferredSheetFormat.Auto
    Public Property UseAutomaticScale As Boolean = True
    Public Property ManualScale As Double = 1.0
    Public Property IncludeIsometric As Boolean = True
    Public Property IncludeProjectedViews As Boolean = True
    Public Property IncludeFlatInDraftWhenPsm As Boolean = True
    ''' <summary>Si está activo, tras InsertByFold aplica slots proporcionales (referencia A3→hoja real desde SheetSetup) centrando cada vista con el centro geométrico de <see cref="SolidEdgeDraft.DrawingView.Range"/>.</summary>
    Public Property EnableSlotBBoxViewLayout As Boolean = True
    ''' <summary>Habilita el único motor de acotado automático DV*2d durante la generación del DFT.</summary>
    Public Property EnableAutoDimensioning As Boolean = True
    ''' <summary>Motor exclusivo DVRef minimalista (prefijo log [PRODDIM]); si está activo no se ejecuta el motor principal ni laboratorios de acotación secundarios.</summary>
    Public Property EnableProductionDvRefCleanEngine As Boolean = False
    ''' <summary>Solo si <see cref="EnableAutoDimensioning"/> y no hay laboratorio exclusivo: elige entre motor principal, copia V02 aislada o plugin alternativo.</summary>
    Public Property AutoDimensioningMotor As AutoDimensioningMotorKind = AutoDimensioningMotorKind.CurrentMainPipeline
    ''' <summary>Tras el acotado automático, volcar en el log introspección SDK sobre geometría DV y cotas (prefijo [SESDK_PROBE]). También <c>SE_SESDK_INTROSPECT</c>.</summary>
    Public Property EnableSesdkPostDimensionIntrospection As Boolean = False
    ''' <summary>Si True (y acotado activo): fuerza modo barrido completo de entidades DV (<c>SweepAllEntities</c>) para crear cotas sobre muchas líneas, arcos, círculos y elipses por vista, con límites de banda amplios.</summary>
    Public Property PreferSweepAllDrawingDimensions As Boolean = False
    ''' <summary>Prueba: no alejar cotas con TrackDistance escalonado (motor LegacyV02).</summary>
    Public Property SuppressDimensionTrackDistanceSpacing As Boolean = False
    ''' <summary>Eliminar cotas duplicadas (mismo valor y keypoints) tras el barrido.</summary>
    Public Property EnableKeypointValueDuplicateCleanup As Boolean = True
    ''' <summary>Laboratorio DIMLAB: DVLine2d.Reference + AddDistanceBetweenObjects (exclusivo; no ejecuta el motor principal de acotación).</summary>
    Public Property EnableDrawingViewDimensioningLab As Boolean = False
    ''' <summary>Modo forense: VIS0 en Hoja1, pausa MsgBox, sin cerrar DFT ni exportar PDF/DXF, validación por Range+DisplayData.</summary>
    Public Property EnableDimLabInteractivePause As Boolean = True
    ''' <summary>Solo True si el usuario pulsó el botón dedicado [LAB] DIMLAB (validar que Effective_runLab coincide).</summary>
    Public Property RequestedDimLabFromDedicatedButton As Boolean = False
    ''' <summary>Escenario del laboratorio DIMLAB (horizontal, vertical, completo o forense).</summary>
    Public Property DimLabMode As DimLabMode = DimLabMode.Full
    ''' <summary>Si True, ejecuta la sonda VIS0 (línea en hoja + AddLength 60 mm).</summary>
    Public Property EnableDimLabVisibleProbe As Boolean = False
    ''' <summary>Si True, solo registra [PLACE][NOTE] si el carril superior es estrecho (no cambia TrackDistance salvo lógica futura).</summary>
    Public Property EnableDimLabAlternativePlacement As Boolean = False
    ''' <summary>En DimLabMode.VerticalOnly, crear antes la horizontal DVRef como control (una cadena exclusiva).</summary>
    Public Property EnableDimLabHorizontalControlInVerticalOnly As Boolean = True
    ''' <summary>Si False, borra cotas DIMLAB fallidas (WRONG/NOT_VISIBLE/FAIL/GHOST) tras inspeccionarlas.</summary>
    Public Property DimLabKeepFailedDimensions As Boolean = False
    ''' <summary>Si True, limpia cotas previas de la hoja activa al arrancar DIMLAB.</summary>
    Public Property DimLabCleanPreviousLabDimensions As Boolean = True
    ''' <summary>Parámetros ISO 129 (primera iteración). Nothing = <see cref="DimensioningNormConfig.DefaultConfig"/> en tiempo de ejecución.</summary>
    Public Property DimensioningNormConfig As DimensioningNormConfig
    ''' <summary>Prueba aislada: cota horizontal exterior (vista superior), sin el motor de acotado automático.</summary>
    Public Property RunUnitHorizontalExteriorDimensionTest As Boolean = False
    ''' <summary>Prueba aislada: PMI del modelo + DrawingView.RetrieveDimensions (no sustituye acotación geométrica).</summary>
    Public Property EnablePmiRetrievalProbe As Boolean = False
    ''' <summary>Laboratorio experimental independiente: Drop de DrawingViews a 2D Model para probar acotación sobre geometría 2D real.</summary>
    Public Property RunDropViewsTo2DModelLab As Boolean = False
    ''' <summary>Laboratorio Drop + detección de sheets creadas por SE y acotación sobre Lines2d reales (prefijo log [DROP_SHEETS]).</summary>
    Public Property RunDropCreatedSheetsDimensionLab As Boolean = False
    ''' <summary>Si True con el lab de sheets, ejecuta draft.Save() al final del laboratorio (solo diagnóstico).</summary>
    Public Property DropCreatedSheetsDimensionLabDebugSave As Boolean = False
    ''' <summary>Laboratorio exclusivo de acotación sobre DVGeometry (sin Drop, sin 2D Model), con log [DV_DIMLAB].</summary>
    Public Property RunDVGeometryDimensionPlacementLab As Boolean = True
    ''' <summary>Laboratorio de descubrimiento de métodos SESDK sobre DVGeometry (prefijo [DV_METHODLAB]).</summary>
    Public Property RunDVGeometryMethodDiscoveryLab As Boolean = False
    ''' <summary>Si PMI.PMIModelViews.Count=0 y hay PMI.Dimensions, intenta crear un PMIModelView temporal en el modelo (sin guardar) y reintentar RetrieveDimensions.</summary>
    Public Property ExperimentalCreatePMIModelViewIfMissing As Boolean = False
    ''' <summary>Solo PMI: abrir modelo, ejecutar creación fuertemente tipada de PMIModelView y cerrar; sin Draft ni RetrieveDimensions.</summary>
    Public Property ExperimentalProbeCreatePMIModelViewOnly As Boolean = False
    ''' <summary>Solo PMI: si el binding por nombre falla, intenta DrawingViews.AddPMIModelView en una vista de diagnóstico (no sustituye la vista base).</summary>
    Public Property ExperimentalPmiTryAddPMIModelViewView As Boolean = False
    ''' <summary>Solo PMI: antes de crear el Draft, con el modelo activo, SetViewOrientationToCurrentView+Apply sobre el PMIModelView por nombre (alinear PMI con la vista 3D actual).</summary>
    Public Property ExperimentalPmiSyncPMIModelViewOrientationBeforeDraft As Boolean = False
    ''' <summary>Solo PMI: diagnóstico de proyección (ViewOrientation, SetViewOrientationFromNamedView, reintento Front, AddPMIModelView experimental si falla RetrieveDimensions).</summary>
    Public Property ExperimentalPmiProjectionDiagnostics As Boolean = False
    ''' <summary>Tras crear el DFT y las vistas: inventario geométrico 2D completo por vista (solo log + modelos; no acotación en este pase).</summary>
    Public Property ExperimentalDraftGeometryDiagnostics As Boolean = False
    ''' <summary>Fracción del ancho de vista (hoja, m) para tolerancia de eje en clasificación H/V: tolAxis = max(ancho*esta_fracción, GeometryDiagnosticsTolAxisMinM).</summary>
    Public Property GeometryDiagnosticsTolAxisWidthFraction As Double = 0.001
    ''' <summary>Tolerancia mínima absoluta (m) en hoja para clasificar líneas horizontales/verticales.</summary>
    Public Property GeometryDiagnosticsTolAxisMinM As Double = 0.0000001
    ''' <summary>Radio (m) por debajo del cual un arco se clasifica como «pequeño» en [GEOM][CLASSIFY].</summary>
    Public Property GeometryDiagnosticsArcSmallRadiusM As Double = 0.008
    ''' <summary>Radio (m) por encima o igual del cual un arco se clasifica como «grande».</summary>
    Public Property GeometryDiagnosticsArcLargeRadiusM As Double = 0.03
    ''' <summary>Diámetro (m) máximo para clasificar un círculo como candidato a agujero.</summary>
    Public Property GeometryDiagnosticsHoleCandidateDiamM As Double = 0.025
    Public Property UseBestBaseViewLogic As Boolean = True

    Public Property ClientName As String = ""
    Public Property ProjectName As String = ""
    Public Property DrawingTitle As String = ""
    ''' <summary>Manual: usa DrawingTitle de la UI. AutoFromFileName: nombre de archivo sin extensión.</summary>
    Public Property TitleSourceMode As TitleSourceMode = TitleSourceMode.Manual
    Public Property Material As String = ""
    ''' <summary>Espesor / calibre chapa (MechanicalModeling.Sheet Metal Gauge + Custom.Espesor).</summary>
    Public Property Thickness As String = ""
    ''' <summary>Referencia de pedido (Custom.Pedido).</summary>
    Public Property Pedido As String = ""
    ''' <summary>Autor en propiedades; vacío = usuario Windows.</summary>
    Public Property AuthorName As String = ""
    Public Property Weight As String = ""
    Public Property Equipment As String = ""
    Public Property DrawingNumber As String = ""
    Public Property Revision As String = ""
    Public Property Notes As String = ""

    ''' <summary>Fecha del plano (texto ISO o corta) para Custom.FechaPlano.</summary>
    Public Property FechaPlano As String = ""
    ''' <summary>Metadatos PART_LIST: cotas mayor/menor en mm como texto.</summary>
    Public Property PartListL As String = ""
    Public Property PartListH As String = ""
    Public Property PartListD As String = ""
    Public Property PartListNombreArchivo As String = ""
    Public Property PartListCantidad As String = "1"
    ''' <summary>Si es True, ValidateConfiguration bloquea con campos de metadatos incompletos.</summary>
    Public Property StrictMetadataValidation As Boolean = False

    ''' <summary>Tras guardar el DFT en modo normal no cerrar documento hasta revisión manual en Solid Edge (no es laboratorio).</summary>
    Public Property KeepDftOpenAfterRun As Boolean = False

    Public Property LastDetectedSourceKind As SourceFileKind = SourceFileKind.Unknown
    Public Property UseSelectedComponents As Boolean = False
    Public Property SelectedComponentPaths As New List(Of String)()

    Public Function DetectInputKind() As SourceFileKind
        If String.IsNullOrWhiteSpace(InputFile) Then Return SourceFileKind.Unknown
        Dim ext As String = IO.Path.GetExtension(InputFile).ToLowerInvariant()
        Select Case ext
            Case ".asm" : Return SourceFileKind.AssemblyFile
            Case ".par" : Return SourceFileKind.PartFile
            Case ".psm" : Return SourceFileKind.SheetMetalFile
            Case ".dft" : Return SourceFileKind.DraftFile
            Case Else : Return SourceFileKind.Unknown
        End Select
    End Function

    ''' <summary>Copia superficial para aplicar título/autor efectivos sin mutar la configuración de la UI.</summary>
    Public Function CloneForExecution() As JobConfiguration
        Return New JobConfiguration With {
            .InputFile = InputFile,
            .OutputFolder = OutputFolder,
            .MotorPhase = MotorPhase,
            .TemplateA4 = TemplateA4,
            .TemplateA3 = TemplateA3,
            .TemplateA2 = TemplateA2,
            .TemplateDxf = TemplateDxf,
            .CreateDraft = CreateDraft,
            .CreatePdf = CreatePdf,
            .CreateDxfFromDraft = CreateDxfFromDraft,
            .CreateFlatDxf = CreateFlatDxf,
            .OpenOutputFolderWhenDone = OpenOutputFolderWhenDone,
            .OverwriteExisting = OverwriteExisting,
            .ProcessRepeatedComponentsOnce = ProcessRepeatedComponentsOnce,
            .DetailedLog = DetailedLog,
            .DebugTemplatesInspection = DebugTemplatesInspection,
            .KeepSolidEdgeVisible = KeepSolidEdgeVisible,
            .InsertPropertiesInTitleBlock = InsertPropertiesInTitleBlock,
            .TitleBlockPropertySourceMode = TitleBlockPropertySourceMode,
            .PreferredFormat = PreferredFormat,
            .UseAutomaticScale = UseAutomaticScale,
            .ManualScale = ManualScale,
            .IncludeIsometric = IncludeIsometric,
            .IncludeProjectedViews = IncludeProjectedViews,
            .IncludeFlatInDraftWhenPsm = IncludeFlatInDraftWhenPsm,
            .EnableSlotBBoxViewLayout = EnableSlotBBoxViewLayout,
            .EnableAutoDimensioning = EnableAutoDimensioning,
            .AutoDimensioningMotor = AutoDimensioningMotor,
            .EnableSesdkPostDimensionIntrospection = EnableSesdkPostDimensionIntrospection,
            .PreferSweepAllDrawingDimensions = PreferSweepAllDrawingDimensions,
            .SuppressDimensionTrackDistanceSpacing = SuppressDimensionTrackDistanceSpacing,
            .EnableKeypointValueDuplicateCleanup = EnableKeypointValueDuplicateCleanup,
            .EnableProductionDvRefCleanEngine = EnableProductionDvRefCleanEngine,
            .EnableDrawingViewDimensioningLab = EnableDrawingViewDimensioningLab,
            .EnableDimLabInteractivePause = EnableDimLabInteractivePause,
            .RequestedDimLabFromDedicatedButton = RequestedDimLabFromDedicatedButton,
            .DimLabMode = DimLabMode,
            .EnableDimLabVisibleProbe = EnableDimLabVisibleProbe,
            .EnableDimLabAlternativePlacement = EnableDimLabAlternativePlacement,
            .EnableDimLabHorizontalControlInVerticalOnly = EnableDimLabHorizontalControlInVerticalOnly,
            .DimLabKeepFailedDimensions = DimLabKeepFailedDimensions,
            .DimLabCleanPreviousLabDimensions = DimLabCleanPreviousLabDimensions,
            .DimensioningNormConfig = If(DimensioningNormConfig Is Nothing, Nothing, DimensioningNormConfig.Clone()),
            .RunUnitHorizontalExteriorDimensionTest = RunUnitHorizontalExteriorDimensionTest,
            .EnablePmiRetrievalProbe = EnablePmiRetrievalProbe,
            .RunDropViewsTo2DModelLab = RunDropViewsTo2DModelLab,
            .RunDropCreatedSheetsDimensionLab = RunDropCreatedSheetsDimensionLab,
            .DropCreatedSheetsDimensionLabDebugSave = DropCreatedSheetsDimensionLabDebugSave,
            .RunDVGeometryDimensionPlacementLab = RunDVGeometryDimensionPlacementLab,
            .RunDVGeometryMethodDiscoveryLab = RunDVGeometryMethodDiscoveryLab,
            .ExperimentalCreatePMIModelViewIfMissing = ExperimentalCreatePMIModelViewIfMissing,
            .ExperimentalProbeCreatePMIModelViewOnly = ExperimentalProbeCreatePMIModelViewOnly,
            .ExperimentalPmiTryAddPMIModelViewView = ExperimentalPmiTryAddPMIModelViewView,
            .ExperimentalPmiSyncPMIModelViewOrientationBeforeDraft = ExperimentalPmiSyncPMIModelViewOrientationBeforeDraft,
            .ExperimentalPmiProjectionDiagnostics = ExperimentalPmiProjectionDiagnostics,
            .ExperimentalDraftGeometryDiagnostics = ExperimentalDraftGeometryDiagnostics,
            .GeometryDiagnosticsTolAxisWidthFraction = GeometryDiagnosticsTolAxisWidthFraction,
            .GeometryDiagnosticsTolAxisMinM = GeometryDiagnosticsTolAxisMinM,
            .GeometryDiagnosticsArcSmallRadiusM = GeometryDiagnosticsArcSmallRadiusM,
            .GeometryDiagnosticsArcLargeRadiusM = GeometryDiagnosticsArcLargeRadiusM,
            .GeometryDiagnosticsHoleCandidateDiamM = GeometryDiagnosticsHoleCandidateDiamM,
            .UseBestBaseViewLogic = UseBestBaseViewLogic,
            .ClientName = ClientName,
            .ProjectName = ProjectName,
            .DrawingTitle = DrawingTitle,
            .TitleSourceMode = TitleSourceMode,
            .Material = Material,
            .Thickness = Thickness,
            .Pedido = Pedido,
            .AuthorName = AuthorName,
            .Weight = Weight,
            .Equipment = Equipment,
            .DrawingNumber = DrawingNumber,
            .Revision = Revision,
            .Notes = Notes,
            .FechaPlano = FechaPlano,
            .PartListL = PartListL,
            .PartListH = PartListH,
            .PartListD = PartListD,
            .PartListNombreArchivo = PartListNombreArchivo,
            .PartListCantidad = PartListCantidad,
            .StrictMetadataValidation = StrictMetadataValidation,
            .KeepDftOpenAfterRun = KeepDftOpenAfterRun,
            .LastDetectedSourceKind = LastDetectedSourceKind,
            .UseSelectedComponents = UseSelectedComponents,
            .SelectedComponentPaths = If(SelectedComponentPaths Is Nothing, New List(Of String)(), New List(Of String)(SelectedComponentPaths))
        }
    End Function
End Class
