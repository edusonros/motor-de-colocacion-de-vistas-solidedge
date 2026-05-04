Option Strict Off

''' <summary>Parámetros de la primera iteración de acotación inspirada en UNE-EN ISO 129-1 (pragmática, extensible).</summary>
Public Class DimensioningNormConfig

    Public Property EnableISO129Rules As Boolean = True
    Public Property AutoDimPragmaticVisibleFirst As Boolean = True

    Public Property PreferOutsideDimensions As Boolean = True
    Public Property AvoidHiddenGeometry As Boolean = True
    Public Property AvoidDuplicateDimensions As Boolean = True
    Public Property AvoidInsideContour As Boolean = True
    Public Property AvoidTitleBlock As Boolean = True
    Public Property AvoidBorder As Boolean = True
    Public Property AvoidOverlaps As Boolean = True

    Public Property UseTotalDimensionsFirst As Boolean = True
    Public Property UseParallelDimensioning As Boolean = True
    Public Property UseChainDimensioningOnlyIfNeeded As Boolean = True
    Public Property UseRepeatedFeatureNotation As Boolean = True

    Public Property MinGapFromView As Double = 0.012
    Public Property GapBetweenDimensionRows As Double = 0.008
    Public Property MinFeatureSeparation As Double = 0.001
    Public Property MaxDimensionsPerViewInitial As Integer = 4

    ''' <summary>Si True, el postproceso puede intentar fusionar/eliminar cotas duplicadas (desaconsejado si se editan keypoints a mano).</summary>
    Public Property EnableDuplicateDimensionCleanup As Boolean = False

    ''' <summary>Si True, permitir borrar duplicados “inseguros” en postproceso. Por defecto False: no se borran por ser duplicados.</summary>
    Public Property EnableDeleteUnsafeDuplicates As Boolean = False

    ''' <summary>Si True (por defecto), se mantienen todas las cotas con el mismo valor nominal para flujo manual (p. ej. varias 340 y recolocar keypoints).</summary>
    Public Property KeepIntentionalDuplicateDimensions As Boolean = True

    ''' <summary>Si True, el paso UNE129 posterior solo ordena/recoloca/separa carriles; no simplifica el número de cotas ni las trata como error.</summary>
    Public Property OnlyArrangeExistingDimensions As Boolean = True

    ''' <summary>No omitir creación en barrido por magnitud ya vista en otra vista (solo si además <see cref="EnableDuplicateDimensionCleanup"/> y sin intención de duplicados).</summary>
    Public Property UneDedupeNominalAcrossOrthogonalViews As Boolean = False

    ''' <summary>Procesar primero la vista ortogonal de mayor área en hoja.</summary>
    Public Property UneProcessLargestOrthogonalViewFirst As Boolean = True

    Public Const ModeTargetReference As String = "TargetDrawingLikeReference"
    Public Const ModeSweepAll As String = "SweepAllEntities"

    ''' <summary>TargetDrawingLikeReference = pocas cotas como plano de referencia; SweepAllEntities = barrido DV completo (legado).</summary>
    Public Property DimensionCreationMode As String = ModeTargetReference

    Public Property CreateAllLineLengths As Boolean = False
    Public Property CreateAllArcRadii As Boolean = False
    Public Property SweepAllDVLines As Boolean = False
    Public Property SweepAllDVArcs As Boolean = False

    ''' <summary>Equilibrio referencia vs. barrido: valores por defecto más permisivos para conservar rectas principales.</summary>
    Public Property MaxTotalDimensionsTarget As Integer = 28
    Public Property MaxLinearDimensionsTarget As Integer = 20
    Public Property MaxRadialDimensionsTarget As Integer = 8

    Public Property AllowSomeDuplicate340 As Boolean = True
    Public Property AllowSomeDuplicate90 As Boolean = True
    Public Property AllowSomeDuplicate102 As Boolean = True
    Public Property KeepManualEditableDuplicates As Boolean = True

    Public Property PrepareExistingPartsListTop As Boolean = True
    Public Property PartsListMarginXm As Double = 0.006
    Public Property PartsListMarginYm As Double = 0.006
    Public Property ProtectedZoneSafetyMarginM As Double = 0.005

    ''' <summary>Si False, no se inserta PartsList nativa; solo zona heurística superior.</summary>
    Public Property EnablePartsListCreation As Boolean = True
    ''' <summary>Nombre exacto del SavedSettings de lista de piezas en Solid Edge (cuadro Editar definición).</summary>
    Public Property PartsListSavedSettingsName As String = "PART_LIST"
    ''' <summary>Nombre del estilo de tabla (TableStyle) asociado a la plantilla PART_LIST.</summary>
    Public Property PartsListTableStyleName As String = "PART_LIST"
    ''' <summary>Si True, intentar primero <c>PartsLists.AddEx</c>; si falla, se prueba <c>Add</c>.</summary>
    Public Property PartsListUseAddEx As Boolean = True
    Public Property PartsListAutoBalloon As Long = 0
    Public Property PartsListCreatePartsList As Long = 1
    ''' <summary>Origen de la PartsList en coordenadas de hoja (m), eje Y hacia arriba (Solid Edge).</summary>
    Public Property PartsListOriginX As Double = 0.01
    Public Property PartsListOriginY As Double = 0.287
    ''' <summary>Si True, elimina listas existentes antes de crear para evitar duplicados.</summary>
    Public Property DeleteExistingPartsListsBeforeCreate As Boolean = True
    ''' <summary>Solo si True: tras fallar PART_LIST, intento diagnóstico con estilo ANSI (no usar en producción).</summary>
    Public Property AllowPartsListFallbackAnsi As Boolean = False
    ''' <summary>Mínimo de columnas esperado para PART_LIST (plantilla real suele ser &gt; 6; 6 suele ser genérico).</summary>
    Public Property PartsListMinExpectedColumns As Integer = 7

    ''' <summary>Altura estimada de la tabla (m) para colocar el borde inferior cerca del margen superior de hoja.</summary>
    Public Property PartsListEstimatedHeightM As Double = 0.028
    ''' <summary>Si True, SetOrigin Y = SheetHeight - margenSuperior - PartsListEstimatedHeightM (ancla típica esquina inferior-izquierda de bloque).</summary>
    Public Property PartsListOriginIsTableBottomLeft As Boolean = True

    ''' <summary>Incluir aristas extremas de vista (igual que motor previo) con máxima prioridad.</summary>
    Public Property ReferenceIncludeViewExtents As Boolean = True
    ''' <summary>Tolerancia H/V proporcional: fracción del mayor lado de la vista en hoja.</summary>
    Public Property ReferenceAxisToleranceFraction As Double = 0.0025R
    ''' <summary>Si |dx|/|dy| o |dy|/|dx| &lt;= este valor, la línea se trata como horizontal o vertical.</summary>
    Public Property ReferenceObliqueAsAxisMaxRatio As Double = 0.12R
    ''' <summary>Longitud mínima de línea candidata como fracción del mayor lado de la vista.</summary>
    Public Property ReferenceMinLineLengthFraction As Double = 0.015R
    ''' <summary>Cotas lineales “genéricas” (prioridad alta no tabulada): máximo por mismo nominal mm.</summary>
    Public Property ReferenceGenericLineDupCap As Integer = 2

    Public Shared Function DefaultConfig() As DimensioningNormConfig
        Return New DimensioningNormConfig()
    End Function

    Public Function Clone() As DimensioningNormConfig
        Return New DimensioningNormConfig With {
            .EnableISO129Rules = EnableISO129Rules,
            .AutoDimPragmaticVisibleFirst = AutoDimPragmaticVisibleFirst,
            .PreferOutsideDimensions = PreferOutsideDimensions,
            .AvoidHiddenGeometry = AvoidHiddenGeometry,
            .AvoidDuplicateDimensions = AvoidDuplicateDimensions,
            .AvoidInsideContour = AvoidInsideContour,
            .AvoidTitleBlock = AvoidTitleBlock,
            .AvoidBorder = AvoidBorder,
            .AvoidOverlaps = AvoidOverlaps,
            .UseTotalDimensionsFirst = UseTotalDimensionsFirst,
            .UseParallelDimensioning = UseParallelDimensioning,
            .UseChainDimensioningOnlyIfNeeded = UseChainDimensioningOnlyIfNeeded,
            .UseRepeatedFeatureNotation = UseRepeatedFeatureNotation,
            .MinGapFromView = MinGapFromView,
            .GapBetweenDimensionRows = GapBetweenDimensionRows,
            .MinFeatureSeparation = MinFeatureSeparation,
            .MaxDimensionsPerViewInitial = MaxDimensionsPerViewInitial,
            .EnableDuplicateDimensionCleanup = EnableDuplicateDimensionCleanup,
            .EnableDeleteUnsafeDuplicates = EnableDeleteUnsafeDuplicates,
            .KeepIntentionalDuplicateDimensions = KeepIntentionalDuplicateDimensions,
            .OnlyArrangeExistingDimensions = OnlyArrangeExistingDimensions,
            .UneDedupeNominalAcrossOrthogonalViews = UneDedupeNominalAcrossOrthogonalViews,
            .UneProcessLargestOrthogonalViewFirst = UneProcessLargestOrthogonalViewFirst,
            .DimensionCreationMode = DimensionCreationMode,
            .CreateAllLineLengths = CreateAllLineLengths,
            .CreateAllArcRadii = CreateAllArcRadii,
            .SweepAllDVLines = SweepAllDVLines,
            .SweepAllDVArcs = SweepAllDVArcs,
            .MaxTotalDimensionsTarget = MaxTotalDimensionsTarget,
            .MaxLinearDimensionsTarget = MaxLinearDimensionsTarget,
            .MaxRadialDimensionsTarget = MaxRadialDimensionsTarget,
            .AllowSomeDuplicate340 = AllowSomeDuplicate340,
            .AllowSomeDuplicate90 = AllowSomeDuplicate90,
            .AllowSomeDuplicate102 = AllowSomeDuplicate102,
            .KeepManualEditableDuplicates = KeepManualEditableDuplicates,
            .PrepareExistingPartsListTop = PrepareExistingPartsListTop,
            .PartsListMarginXm = PartsListMarginXm,
            .PartsListMarginYm = PartsListMarginYm,
            .ProtectedZoneSafetyMarginM = ProtectedZoneSafetyMarginM,
            .EnablePartsListCreation = EnablePartsListCreation,
            .PartsListSavedSettingsName = PartsListSavedSettingsName,
            .PartsListTableStyleName = PartsListTableStyleName,
            .PartsListUseAddEx = PartsListUseAddEx,
            .PartsListAutoBalloon = PartsListAutoBalloon,
            .PartsListCreatePartsList = PartsListCreatePartsList,
            .PartsListOriginX = PartsListOriginX,
            .PartsListOriginY = PartsListOriginY,
            .DeleteExistingPartsListsBeforeCreate = DeleteExistingPartsListsBeforeCreate,
            .AllowPartsListFallbackAnsi = AllowPartsListFallbackAnsi,
            .PartsListMinExpectedColumns = PartsListMinExpectedColumns,
            .PartsListEstimatedHeightM = PartsListEstimatedHeightM,
            .PartsListOriginIsTableBottomLeft = PartsListOriginIsTableBottomLeft,
            .ReferenceIncludeViewExtents = ReferenceIncludeViewExtents,
            .ReferenceAxisToleranceFraction = ReferenceAxisToleranceFraction,
            .ReferenceObliqueAsAxisMaxRatio = ReferenceObliqueAsAxisMaxRatio,
            .ReferenceMinLineLengthFraction = ReferenceMinLineLengthFraction,
            .ReferenceGenericLineDupCap = ReferenceGenericLineDupCap
        }
    End Function
End Class
