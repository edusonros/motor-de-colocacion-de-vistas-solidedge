Option Strict Off

Imports System.Runtime.InteropServices
Imports System.Text.RegularExpressions
Imports System.Text
Imports SolidEdgeFramework
Imports SolidEdgeFileProperties

Public Class SourcePropertiesData
    Public Property DrawingTitle As String = ""
    Public Property ProjectName As String = ""
    Public Property Material As String = ""
    Public Property Thickness As String = ""
    Public Property Weight As String = ""
    ''' <summary>Origen del peso (p. ej. Model.MassProperties, Custom.Peso).</summary>
    Public Property WeightSource As String = ""
    Public Property Equipment As String = ""
    Public Property DrawingNumber As String = ""
    Public Property Revision As String = ""
    Public Property Notes As String = ""
    Public Property ClientName As String = ""
    ''' <summary>L/H/D en mm desde caja 3D del modelo (.par/.psm) durante la lectura del archivo.</summary>
    Public Property PartListL As String = ""
    Public Property PartListH As String = ""
    Public Property PartListD As String = ""
End Class

Public Class SolidEdgePropertyService
    Private Enum PropertyWriteStatus
        Updated = 0
        Added = 1
        SetNotFound = 2
        PropertyNotFound = 3
        ReadOnlyOrTypeMismatch = 4
        UnexpectedError = 5
    End Enum

    Private Enum PropertySyncTarget
        ModelOnly = 1
        DraftOnly = 2
        Both = 3
    End Enum

    Private Class PropertyBinding
        Public Property SetName As String
        Public Property PropertyName As String
        Public Property AllowCreate As Boolean
    End Class

    Private Class PropertySyncEntry
        Public Property LogicalName As String
        Public Property Value As String
        Public Property Target As PropertySyncTarget
        Public ReadOnly Property StandardBindings As New List(Of PropertyBinding)()
        Public ReadOnly Property CustomBindings As New List(Of PropertyBinding)()
    End Class

    Public Shared Function TryReadSourceProperties(filePath As String,
                                                   showSolidEdge As Boolean,
                                                   logger As Logger,
                                                   ByRef data As SourcePropertiesData) As Boolean
        data = New SourcePropertiesData()
        If String.IsNullOrWhiteSpace(filePath) Then Return False

        Dim app As Application = Nothing
        Dim createdByUs As Boolean = False
        Dim doc As SolidEdgeDocument = Nothing

        Try
            OleMessageFilter.Register()
            If Not TryConnectApplication(showSolidEdge, logger, app, createdByUs) Then Return False

            logger.Log($"Abriendo documento para leer propiedades: {IO.Path.GetFileName(filePath)}")
            doc = CType(app.Documents.Open(filePath), SolidEdgeDocument)

            data.DrawingTitle = FirstNonEmpty(
                GetDocumentProperty(doc, "SummaryInformation", {"Title", "Document Title"}),
                GetDocumentProperty(doc, "Custom", {"Titulo", "Título", "Title"})
            )
            data.ProjectName = FirstNonEmpty(
                GetDocumentProperty(doc, "ProjectInformation", {"Project Name", "Project", "Document Number"}),
                GetDocumentProperty(doc, "Custom", {"PROYECTO", "Proyecto", "Project"})
            )
            data.Material = FirstNonEmpty(
                GetDocumentProperty(doc, "MechanicalModeling", {"Material"}),
                GetDocumentProperty(doc, "ProjectInformation", {"Material"}),
                GetDocumentProperty(doc, "Custom", {"Material"})
            )
            data.Thickness = FirstNonEmpty(
                GetDocumentProperty(doc, "MechanicalModeling", {"Sheet Metal Gauge"}),
                GetDocumentProperty(doc, "Custom", {"Espesor"})
            )
            ' Peso: no usar Mass de PropertySets estándar; solo Custom o geometría (GetPesoModelo).
            data.Weight = FirstNonEmpty(GetDocumentProperty(doc, "Custom", {"Peso", "Weight"}))
            If Not String.IsNullOrWhiteSpace(data.Weight) Then data.WeightSource = "Custom.Peso"
            data.Equipment = GetDocumentProperty(doc, "Custom", {"Equipo", "Equipment"})
            data.DrawingNumber = GetDocumentProperty(doc, "ProjectInformation", {"Document Number", "Part Number"})
            If String.IsNullOrWhiteSpace(data.DrawingNumber) Then
                data.DrawingNumber = DrawingMetadataService.InferPlanoFromFileName(filePath)
            End If
            data.Revision = GetDocumentProperty(doc, "ProjectInformation", {"Revision", "Revision Number"})
            data.Notes = GetDocumentProperty(doc, "SummaryInformation", {"Comments"})
            If String.IsNullOrWhiteSpace(data.Notes) Then data.Notes = GetDocumentProperty(doc, "Custom", {"Observaciones", "Notes"})
            data.ClientName = FirstNonEmpty(
                GetDocumentProperty(doc, "Custom", {"Cliente", "Client", "Empresa"}),
                GetDocumentProperty(doc, "DocumentSummaryInformation", {"Company"})
            )

            If String.IsNullOrWhiteSpace(data.Weight) Then
                Dim pesoProbe As New DrawingMetadataInput()
                DrawingMetadataService.TryDetectPeso(doc, logger, pesoProbe)
                If Not String.IsNullOrWhiteSpace(pesoProbe.Peso) Then
                    data.Weight = pesoProbe.Peso.Trim()
                    data.WeightSource = If(String.IsNullOrWhiteSpace(pesoProbe.PesoSource), "Model.MassProperties", pesoProbe.PesoSource)
                End If
            End If

            Dim ext = IO.Path.GetExtension(filePath).ToLowerInvariant()
            If ext = ".par" OrElse ext = ".psm" Then
                Dim L As String = "", H As String = "", D As String = ""
                If DrawingMetadataService.TryComputeLhdFromModelDoc(doc, logger, L, H, D) Then
                    data.PartListL = L
                    data.PartListH = H
                    data.PartListD = D
                    If logger IsNot Nothing Then logger.Log("[PARTLISTDATA][LHD][READ_FILE] L=" & L & " H=" & H & " D=" & D)
                End If
            End If

            Return True

        Catch ex As Exception
            If logger IsNot Nothing Then logger.LogException("TryReadSourceProperties", ex)
            Return False
        Finally
            Try
                If doc IsNot Nothing Then
                    doc.Close(False)
                    If logger IsNot Nothing Then logger.Log($"Documento cerrado tras lectura de propiedades: {IO.Path.GetFileName(filePath)}")
                End If
            Catch
            End Try
            Try
                If app IsNot Nothing AndAlso createdByUs Then
                    app.Quit()
                    If logger IsNot Nothing Then logger.Log("Instancia de Solid Edge cerrada (creada por la aplicación).")
                ElseIf app IsNot Nothing Then
                    If logger IsNot Nothing Then logger.Log("Instancia de Solid Edge preexistente: se mantiene abierta.")
                End If
            Catch
            End Try
            Try : OleMessageFilter.Revoke() : Catch : End Try
        End Try
    End Function

    Public Shared Function ApplyTitleBlockPropertyStrategy(dftDoc As Object,
                                                           modelPath As String,
                                                           config As JobConfiguration,
                                                           logger As Logger) As Integer
        If dftDoc Is Nothing OrElse config Is Nothing Then Return 0
        If Not config.InsertPropertiesInTitleBlock Then Return 0
        If logger IsNot Nothing Then logger.Log("[PROPS][TITLEBLOCK][WRITE] Inicio sincronización cajetín/propiedades")

        Dim profile = BuildPropertySyncProfile(config)
        Dim totalUpdated As Integer = 0
        Dim mode As TitleBlockPropertySource = config.TitleBlockPropertySourceMode
        If logger IsNot Nothing Then logger.Log($"[PROPS][MODE] Estrategia configurada: {mode}")

        If mode = TitleBlockPropertySource.FromModelLink Then
            Dim linkCount As Integer = GetModelLinkCount(dftDoc)
            If logger IsNot Nothing Then logger.Log($"[PROPS][MODE] ModelLinks detectados para estrategia: {linkCount}")
            If linkCount <= 0 Then
                If logger IsNot Nothing Then logger.Log("[PROPS][MODE][WARN] FromModelLink sin enlaces. Se aplica fallback automático a FromDraft.")
                mode = TitleBlockPropertySource.FromDraft
            End If
        End If
        If logger IsNot Nothing Then logger.Log($"[PROPS][MODE] Estrategia efectiva: {mode}")

        If mode = TitleBlockPropertySource.FromModelLink Then
            If logger IsNot Nothing Then
                logger.Log("[PROPS][MODE] FromModelLink: la escritura de propiedades se realiza antes de crear el DFT. En esta fase solo se refresca el draft.")
            End If
            RefreshDraftFromModelLinks(dftDoc, logger, config.DebugTemplatesInspection)
            If ShouldFallbackToDraftProperties(dftDoc, logger) Then
                Dim draftFallbackCount As Integer = ApplyProfileToOpenDraft(dftDoc, profile, logger)
                RefreshDraftPropertyTextOnly(dftDoc, logger)
                totalUpdated += draftFallbackCount
                If logger IsNot Nothing Then
                    logger.Log($"[PROPS][MODE][FALLBACK] Se aplicó fallback a FromDraft en DFT actual. Campos={draftFallbackCount}")
                End If
            End If
        Else
            totalUpdated += ApplyProfileToOpenDraft(dftDoc, profile, logger)
            RefreshDraftPropertyTextOnly(dftDoc, logger)
        End If

        If config.DebugTemplatesInspection Then
            LogTemplateTitleBlockBindings(dftDoc, logger)
            LogDraftTitleBlockDataSources(dftDoc, logger)
        End If
        Return totalUpdated
    End Function

    Public Shared Function ApplyDirectSummaryInfoToDraft(draftDoc As Object,
                                                          config As JobConfiguration,
                                                          logger As Logger) As Boolean
        If draftDoc Is Nothing OrElse config Is Nothing Then Return False
        Dim subjectDenom As String = If(config.DrawingTitle, "").Trim()
        Dim titlePlano As String = If(config.DrawingNumber, "").Trim()
        Return ApplyDirectSummaryInfoToDraft(draftDoc, subjectDenom, titlePlano, config.ClientName, config.ProjectName, logger)
    End Function

    Public Shared Function ApplyDirectSummaryInfoToDraft(draftDoc As Object,
                                                          subjectValue As String,
                                                          titleValue As String,
                                                          companyValue As String,
                                                          projectNameValue As String,
                                                          logger As Logger) As Boolean
        If draftDoc Is Nothing Then Return False
        If logger IsNot Nothing Then logger.Log("[DFT][SUMMARYINFO] Iniciando escritura directa sobre DFT.")

        Dim summaryInfo As Object = Nothing
        Try
            summaryInfo = CallByName(draftDoc, "SummaryInfo", CallType.Get)
        Catch ex As Exception
            If logger IsNot Nothing Then logger.Log($"[DFT][SUMMARYINFO][WARN] SummaryInfo -> {ex.GetType().Name}: {ex.Message}")
            Return False
        End Try

        Dim ok As Boolean = True
        ok = SetSummaryInfoField(summaryInfo, "Subject", subjectValue, logger) AndAlso ok
        ok = SetSummaryInfoField(summaryInfo, "Title", titleValue, logger) AndAlso ok
        ok = SetSummaryInfoField(summaryInfo, "Company", companyValue, logger) AndAlso ok
        ok = SetSummaryInfoField(summaryInfo, "ProjectName", projectNameValue, logger) AndAlso ok

        Dim saved As Boolean = False
        Try
            CallByName(draftDoc, "Save", CallType.Method)
            saved = True
        Catch ex As Exception
            ok = False
            If logger IsNot Nothing Then logger.Log($"[DFT][SUMMARYINFO][WARN] Save -> {ex.GetType().Name}: {ex.Message}")
        End Try
        If logger IsNot Nothing Then logger.Log($"[DFT][SUMMARYINFO] Save={saved}")
        Return ok AndAlso saved
    End Function

    Public Shared Function ApplyDirectSummaryInfoToDraftFile(dftPath As String,
                                                              showSolidEdge As Boolean,
                                                              config As JobConfiguration,
                                                              logger As Logger) As Boolean
        If String.IsNullOrWhiteSpace(dftPath) OrElse config Is Nothing Then Return False
        If Not IO.File.Exists(dftPath) Then
            If logger IsNot Nothing Then logger.Log($"[UI][DFT][WARN] No existe DFT: {dftPath}")
            Return False
        End If

        Dim app As Application = Nothing
        Dim createdByUs As Boolean = False
        Dim dftDoc As Object = Nothing
        Try
            OleMessageFilter.Register()
            If Not TryConnectApplication(showSolidEdge, logger, app, createdByUs) Then Return False

            dftDoc = app.Documents.Open(dftPath)
            Return ApplyDirectSummaryInfoToDraft(dftDoc, config, logger)
        Catch ex As Exception
            If logger IsNot Nothing Then logger.LogException("ApplyDirectSummaryInfoToDraftFile", ex)
            Return False
        Finally
            Try
                If dftDoc IsNot Nothing Then CallByName(dftDoc, "Close", CallType.Method, False)
            Catch
            End Try
            Try
                If app IsNot Nothing AndAlso createdByUs Then app.Quit()
            Catch
            End Try
            Try : OleMessageFilter.Revoke() : Catch : End Try
        End Try
    End Function

    Private Shared Function ShouldFallbackToDraftProperties(dftDoc As Object, logger As Logger) As Boolean
        If dftDoc Is Nothing Then Return False
        Try
            Dim title As String = GetDocumentProperty(dftDoc, "SummaryInformation", {"Title", "Document Title", "Título", "Titulo"})
            Dim project As String = GetDocumentProperty(dftDoc, "ProjectInformation", {"Project Name", "Nombre de proyecto", "Project"})
            Dim company As String = GetDocumentProperty(dftDoc, "DocumentSummaryInformation", {"Company", "Empresa"})

            Dim score As Integer = 0
            If IsLikelyTemplatePlaceholder(title, {"TITULO", "TÍTULO", "TITLE"}) Then score += 1
            If IsLikelyTemplatePlaceholder(project, {"PROYECTO", "PROJECT"}) Then score += 1
            If IsLikelyTemplatePlaceholder(company, {"CLIENTE", "COMPANY", "EMPRESA"}) Then score += 1

            If logger IsNot Nothing Then
                logger.Log($"[PROPS][MODE][CHECK] PlaceholderScore={score} Title='{title}' Project='{project}' Company='{company}'")
            End If
            Return score >= 2
        Catch ex As Exception
            If logger IsNot Nothing Then logger.LogException("ShouldFallbackToDraftProperties", ex)
            Return False
        End Try
    End Function

    Private Shared Function SetSummaryInfoField(summaryInfo As Object, fieldName As String, fieldValue As String, logger As Logger) As Boolean
        If summaryInfo Is Nothing Then Return False
        If fieldValue Is Nothing Then fieldValue = ""
        Try
            CallByName(summaryInfo, fieldName, CallType.Let, fieldValue)
            If logger IsNot Nothing Then logger.Log($"[DFT][SUMMARYINFO] {fieldName}='{fieldValue}'")
            Return True
        Catch ex As Exception
            If logger IsNot Nothing Then logger.Log($"[DFT][SUMMARYINFO][WARN] {fieldName} -> {ex.GetType().Name}: {ex.Message}")
            Return False
        End Try
    End Function

    Private Shared Function IsLikelyTemplatePlaceholder(value As String, placeholders As IEnumerable(Of String)) As Boolean
        If String.IsNullOrWhiteSpace(value) Then Return False
        Dim normalized As String = value.Trim().ToUpperInvariant()
        For Each p In placeholders
            If normalized = p.Trim().ToUpperInvariant() Then Return True
        Next
        Return False
    End Function

    Private Shared Function GetModelLinkCount(dftDoc As Object) As Integer
        If dftDoc Is Nothing Then Return 0
        Try
            Dim modelLinks As Object = Nothing
            Try : modelLinks = CallByName(dftDoc, "ModelLinks", CallType.Get) : Catch : End Try
            If modelLinks Is Nothing Then Return 0
            Try : Return CInt(CallByName(modelLinks, "Count", CallType.Get)) : Catch : End Try
        Catch
        End Try
        Return 0
    End Function

    Public Shared Sub ApplyPropertiesToDraftDocument(dftDoc As Object, config As JobConfiguration, logger As Logger)
        If dftDoc Is Nothing OrElse config Is Nothing Then Return

        Dim profile = BuildPropertySyncProfile(config)
        Dim updatedCount As Integer = ApplyProfileToOpenDraft(dftDoc, profile, logger)
        Try
            CallByName(dftDoc, "Update", CallType.Method)
        Catch
        End Try

        If logger IsNot Nothing Then
            logger.Log($"[PROPS][DRAFT] Propiedades aplicadas al DFT abierto: {updatedCount}")
            If updatedCount = 0 Then
                logger.Log("[PROPS][WARN] No se pudo actualizar ninguna propiedad del DFT (revisar mapeo de PropertySets/plantilla).")
            End If
        End If
    End Sub

    Public Shared Sub ApplyPropertiesToSavedDraft(app As Application, dftPath As String, config As JobConfiguration, logger As Logger)
        If app Is Nothing OrElse String.IsNullOrWhiteSpace(dftPath) OrElse Not IO.File.Exists(dftPath) Then Return
        Dim dftDoc As Object = Nothing
        Try
            If logger IsNot Nothing AndAlso config IsNot Nothing Then
                logger.Log($"[PROPS][MODE] ApplyPropertiesToSavedDraft -> {config.TitleBlockPropertySourceMode}")
            End If
            Dim filePropsUpdated As Integer = 0
            If config IsNot Nothing AndAlso config.TitleBlockPropertySourceMode = TitleBlockPropertySource.FromDraft Then
                filePropsUpdated = ApplyPropertiesUsingFilePropertySets(dftPath, config, logger)
                If logger IsNot Nothing Then logger.Log($"[PROPS][DFT_FILE] FileProperties actualizadas en DFT guardado: {filePropsUpdated}")
            End If

            dftDoc = app.Documents.Open(dftPath)
            If config IsNot Nothing AndAlso config.TitleBlockPropertySourceMode = TitleBlockPropertySource.FromDraft AndAlso filePropsUpdated <= 0 Then
                ApplyPropertiesToDraftDocument(dftDoc, config, logger)
            End If
            If config IsNot Nothing AndAlso config.TitleBlockPropertySourceMode = TitleBlockPropertySource.FromModelLink Then
                RefreshDraftFromModelLinks(dftDoc, logger, config.DebugTemplatesInspection)
            Else
                RefreshDraftPropertyTextOnly(dftDoc, logger)
            End If
            RefreshNativePartsListsAndUpdateAll(dftDoc, logger)
            Try : CallByName(dftDoc, "Update", CallType.Method) : Catch : End Try
            Try : CallByName(dftDoc, "Save", CallType.Method) : Catch : End Try
            If logger IsNot Nothing Then logger.Log($"Propiedades guardadas en DFT final: {IO.Path.GetFileName(dftPath)}")
        Catch ex As Exception
            If logger IsNot Nothing Then logger.LogException("ApplyPropertiesToSavedDraft", ex)
        Finally
            Try
                If dftDoc IsNot Nothing Then
                    CallByName(dftDoc, "Close", CallType.Method, False)
                End If
            Catch
            End Try
        End Try
    End Sub

    Public Shared Function ApplyPropertiesToModelFile(modelPath As String, config As JobConfiguration, logger As Logger) As Integer
        If config Is Nothing Then Return 0
        Dim profile = BuildPropertySyncProfile(config)
        Return ApplyProfileToFile(modelPath, profile, PropertySyncTarget.ModelOnly, logger, "MODEL")
    End Function

    Public Shared Function ApplyPropertiesToOpenModelDocument(app As Application,
                                                              modelPath As String,
                                                              config As JobConfiguration,
                                                              logger As Logger) As Integer
        If app Is Nothing OrElse config Is Nothing Then Return 0
        If String.IsNullOrWhiteSpace(modelPath) OrElse Not IO.File.Exists(modelPath) Then Return 0

        Dim doc As Object = Nothing
        Dim openedByUs As Boolean = False
        Dim written As Integer = 0
        Try
            doc = TryGetOpenDocumentByPath(app, modelPath)
            If doc Is Nothing Then
                doc = app.Documents.Open(modelPath)
                openedByUs = True
            End If
            If doc Is Nothing Then
                If logger IsNot Nothing Then logger.Log($"[PROPS][DOC][WARN] No se pudo abrir documento modelo: {modelPath}")
                Return 0
            End If

            If config.DebugTemplatesInspection Then
                LogAllPropertySetsFromOpenDocument(doc, logger)
            End If
            Dim profile = BuildPropertySyncProfile(config)
            written = ApplyProfileToOpenModelDocument(doc, profile, logger)

            Dim saved As Boolean = False
            If written > 0 Then
                Try
                    CallByName(doc, "Save", CallType.Method)
                    saved = True
                Catch exSave As Exception
                    If logger IsNot Nothing Then logger.LogException("ApplyPropertiesToOpenModelDocument.Save", exSave)
                End Try
            End If
            If logger IsNot Nothing Then
                logger.Log($"[PROPS][WRITE] Campos escritos en modelo={written}")
                logger.Log($"[PROPS][WRITE] Save={saved}")
            End If
        Catch ex As Exception
            If logger IsNot Nothing Then logger.LogException("ApplyPropertiesToOpenModelDocument", ex)
        Finally
            If openedByUs AndAlso doc IsNot Nothing Then
                Try : CallByName(doc, "Close", CallType.Method, False) : Catch : End Try
            End If
        End Try
        Return written
    End Function

    Public Shared Sub RefreshDraftFromModelLinks(dftDoc As Object, logger As Logger, Optional verboseLinks As Boolean = False)
        If dftDoc Is Nothing Then Return
        Try
            Dim modelLinks As Object = Nothing
            Try : modelLinks = CallByName(dftDoc, "ModelLinks", CallType.Get) : Catch : End Try
            Dim linkCount As Integer = 0
            Try : linkCount = CInt(CallByName(modelLinks, "Count", CallType.Get)) : Catch : End Try
            If logger IsNot Nothing Then logger.Log($"[PROPS][MLINK] ModelLinks={linkCount}")
            If linkCount = 0 AndAlso logger IsNot Nothing Then
                logger.Log("[PROPS][MLINK][WARN] Sin ModelLinks: el cajetín con |R1 puede no resolver propiedades del modelo.")
            End If

            For i As Integer = 1 To linkCount
                Dim linkObj As Object = GetCollectionItem(modelLinks, i)
                If linkObj Is Nothing Then Continue For
                Dim linkedName As String = ""
                Try : linkedName = CStr(CallByName(linkObj, "FileName", CallType.Get)) : Catch : End Try
                If verboseLinks AndAlso logger IsNot Nothing Then
                    logger.Log($"[PROPS][MLINK] Link#{i}: {If(String.IsNullOrWhiteSpace(linkedName), "(sin nombre)", linkedName)}")
                End If
                Try : CallByName(linkObj, "UpdateViews", CallType.Method) : Catch : End Try
                Try : CallByName(linkObj, "ForceUpdateViews", CallType.Method) : Catch : End Try
            Next

            RefreshDraftPropertyTextOnly(dftDoc, logger)
        Catch ex As Exception
            If logger IsNot Nothing Then logger.LogException("RefreshDraftFromModelLinks", ex)
        End Try
    End Sub

    Public Shared Sub LogAllPropertySetsFromOpenDocument(doc As Object, logger As Logger, Optional maxPropsPerSet As Integer = 200)
        If doc Is Nothing OrElse logger Is Nothing Then Return
        Try
            Dim fullName As String = ""
            Dim docType As String = ""
            Try : fullName = CStr(CallByName(doc, "FullName", CallType.Get)) : Catch : End Try
            Try : docType = TypeName(doc) : Catch : End Try
            logger.Log($"[PROPS][DOC][INSPECT] Archivo={fullName}")
            logger.Log($"[PROPS][DOC][INSPECT] Tipo={docType}")

            Dim sets As Object = Nothing
            Try : sets = CallByName(doc, "Properties", CallType.Get) : Catch : End Try
            If sets Is Nothing Then
                Try : sets = CallByName(doc, "PropertySets", CallType.Get) : Catch : End Try
            End If
            If sets Is Nothing Then
                logger.Log("[PROPS][DOC][INSPECT][WARN] Documento sin colección Properties/PropertySets accesible.")
                Return
            End If

            Dim setCount As Integer = 0
            Try : setCount = CInt(CallByName(sets, "Count", CallType.Get)) : Catch : End Try
            logger.Log($"[PROPS][DOC][INSPECT] PropertySets detectados={setCount}")

            For si As Integer = 1 To setCount
                Dim setObj As Object = GetCollectionItem(sets, si)
                If setObj Is Nothing Then Continue For
                Dim setName As String = $"Set#{si}"
                Try : setName = CStr(CallByName(setObj, "Name", CallType.Get)) : Catch : End Try
                Dim propCount As Integer = 0
                Try : propCount = CInt(CallByName(setObj, "Count", CallType.Get)) : Catch : End Try
                logger.Log($"[PROPS][DOC][INSPECT] Set={setName} Count={propCount}")

                Dim printed As Integer = 0
                For pi As Integer = 1 To propCount
                    If printed >= Math.Max(1, maxPropsPerSet) Then
                        logger.Log($"[PROPS][DOC][INSPECT] Set={setName} ... truncado en {maxPropsPerSet} propiedades.")
                        Exit For
                    End If
                    Dim propObj As Object = GetCollectionItem(setObj, pi)
                    If propObj Is Nothing Then Continue For
                    Dim propName As String = ""
                    Dim propValue As String = ""
                    Try : propName = CStr(CallByName(propObj, "Name", CallType.Get)) : Catch : End Try
                    Try
                        Dim v As Object = CallByName(propObj, "Value", CallType.Get)
                        If v IsNot Nothing Then propValue = v.ToString()
                    Catch
                    End Try
                    propValue = If(propValue, "").Replace(vbCr, " ").Replace(vbLf, " ")
                    logger.Log($"[PROPS][DOC][INSPECT] {setName}.{propName}='{propValue}'")
                    printed += 1
                Next
            Next
        Catch ex As Exception
            logger.LogException("LogAllPropertySetsFromOpenDocument", ex)
        End Try
    End Sub

    Public Shared Sub LogDraftTitleBlockDataSources(dftDoc As Object, logger As Logger)
        If dftDoc Is Nothing OrElse logger Is Nothing Then Return
        Try
            Dim sections As Object = Nothing
            Try : sections = CallByName(dftDoc, "Sections", CallType.Get) : Catch : End Try
            If sections Is Nothing Then
                logger.Log("[TPL][SRC][WARN] No se pudo acceder a Sections del draft.")
                Return
            End If

            Dim secCount As Integer = 0
            Try : secCount = CInt(CallByName(sections, "Count", CallType.Get)) : Catch : End Try
            logger.Log($"[TPL][SRC] Secciones={secCount}")
            For si As Integer = 1 To secCount
                Dim section As Object = GetCollectionItem(sections, si)
                If section Is Nothing Then Continue For
                Dim sheets As Object = Nothing
                Try : sheets = CallByName(section, "Sheets", CallType.Get) : Catch : End Try
                If sheets Is Nothing Then Continue For
                Dim sheetCount As Integer = 0
                Try : sheetCount = CInt(CallByName(sheets, "Count", CallType.Get)) : Catch : End Try
                For shi As Integer = 1 To sheetCount
                    Dim sheet As Object = GetCollectionItem(sheets, shi)
                    If sheet Is Nothing Then Continue For
                    InspectTitleBlockCollectionSource(sheet, "TitleBlocks", logger, si, shi)
                    InspectTextCollectionSource(sheet, "Texts", logger, si, shi)
                    InspectTextCollectionSource(sheet, "TextBoxes", logger, si, shi)
                Next
            Next
        Catch ex As Exception
            logger.LogException("LogDraftTitleBlockDataSources", ex)
        End Try
    End Sub

    Public Shared Sub RefreshDraftPropertyTextOnly(dftDoc As Object, logger As Logger)
        If dftDoc Is Nothing Then Return
        Try
            Dim okCache As Boolean = True
            Dim okDisplay As Boolean = True
            Try : CallByName(dftDoc, "UpdatePropertyTextCacheAndDisplay", CallType.Method) : Catch : okCache = False : End Try
            Try : CallByName(dftDoc, "UpdatePropertyTextDisplay", CallType.Method) : Catch : okDisplay = False : End Try
            Try : CallByName(dftDoc, "Update", CallType.Method) : Catch : End Try
            If logger IsNot Nothing Then
                logger.Log($"[PROPS][REFRESH] UpdatePropertyTextCacheAndDisplay={okCache}, UpdatePropertyTextDisplay={okDisplay}")
            End If
        Catch ex As Exception
            If logger IsNot Nothing Then logger.LogException("RefreshDraftPropertyTextOnly", ex)
        End Try
    End Sub

    Private Shared Function ApplyPropertiesUsingFilePropertySets(filePath As String, config As JobConfiguration, logger As Logger) As Integer
        If config Is Nothing Then Return 0
        Dim profile = BuildPropertySyncProfile(config)
        Return ApplyProfileToFile(filePath, profile, PropertySyncTarget.DraftOnly, logger, "DFT_FILE")
    End Function

    Public Shared Sub LogAllPropertySetsFromFile(filePath As String, logger As Logger, Optional maxPropsPerSet As Integer = 250)
        If logger Is Nothing Then Return
        If String.IsNullOrWhiteSpace(filePath) Then
            logger.Log("[PROPS][INSPECT][WARN] Ruta vacía.")
            Return
        End If
        If Not IO.File.Exists(filePath) Then
            logger.Log($"[PROPS][INSPECT][WARN] No existe archivo: {filePath}")
            Return
        End If

        Dim psets As SolidEdgeFileProperties.PropertySets = Nothing
        Try
            psets = New SolidEdgeFileProperties.PropertySets()
            psets.Open(filePath, True)

            Dim setCount As Integer = 0
            Try : setCount = CInt(CallByName(psets, "Count", CallType.Get)) : Catch : End Try
            logger.Log($"[PROPS][INSPECT] Archivo={filePath}")
            logger.Log($"[PROPS][INSPECT] PropertySets detectados={setCount}")

            For si As Integer = 1 To setCount
                Dim setObj As Object = Nothing
                Try : setObj = GetCollectionItem(psets, si) : Catch : End Try
                If setObj Is Nothing Then Continue For

                Dim setName As String = ""
                Try : setName = CStr(CallByName(setObj, "Name", CallType.Get)) : Catch : End Try
                If String.IsNullOrWhiteSpace(setName) Then setName = $"Set#{si}"

                Dim propCount As Integer = 0
                Try : propCount = CInt(CallByName(setObj, "Count", CallType.Get)) : Catch : End Try
                logger.Log($"[PROPS][INSPECT] Set={setName} Count={propCount}")

                Dim printed As Integer = 0
                For pi As Integer = 1 To propCount
                    If printed >= Math.Max(1, maxPropsPerSet) Then
                        logger.Log($"[PROPS][INSPECT] Set={setName} ... truncado en {maxPropsPerSet} propiedades.")
                        Exit For
                    End If

                    Dim propObj As Object = Nothing
                    Try : propObj = GetCollectionItem(setObj, pi) : Catch : End Try
                    If propObj Is Nothing Then Continue For

                    Dim propName As String = ""
                    Dim propValue As String = ""
                    Try : propName = CStr(CallByName(propObj, "Name", CallType.Get)) : Catch : End Try
                    Try
                        Dim v As Object = CallByName(propObj, "Value", CallType.Get)
                        If v IsNot Nothing Then propValue = v.ToString()
                    Catch
                    End Try
                    If propName Is Nothing Then propName = ""
                    If propValue Is Nothing Then propValue = ""
                    propValue = propValue.Replace(vbCr, " ").Replace(vbLf, " ")
                    logger.Log($"[PROPS][INSPECT] {setName}.{propName} = '{propValue}'")
                    printed += 1
                Next
            Next
            Try : psets.Close() : Catch : End Try
        Catch ex As Exception
            logger.LogException("LogAllPropertySetsFromFile", ex)
        Finally
            Try
                If psets IsNot Nothing AndAlso Marshal.IsComObject(psets) Then Marshal.ReleaseComObject(psets)
            Catch
            End Try
        End Try
    End Sub

    Private Shared Function BuildPropertySyncProfile(config As JobConfiguration) As List(Of PropertySyncEntry)
        Dim list As New List(Of PropertySyncEntry)()
        If config Is Nothing Then Return list

        ' Plano (número / nombre archivo de plano) → Summary.Title + Custom.Plano
        Dim plano = New PropertySyncEntry With {.LogicalName = "Plano", .Value = config.DrawingNumber, .Target = PropertySyncTarget.Both}
        plano.StandardBindings.Add(New PropertyBinding With {.SetName = "SummaryInformation", .PropertyName = "Title", .AllowCreate = False})
        plano.StandardBindings.Add(New PropertyBinding With {.SetName = "SummaryInformation", .PropertyName = "Document Title", .AllowCreate = False})
        plano.StandardBindings.Add(New PropertyBinding With {.SetName = "SummaryInformation", .PropertyName = "Título", .AllowCreate = False})
        plano.StandardBindings.Add(New PropertyBinding With {.SetName = "SummaryInformation", .PropertyName = "Titulo", .AllowCreate = False})
        plano.CustomBindings.Add(New PropertyBinding With {.SetName = "Custom", .PropertyName = "Plano", .AllowCreate = True})
        list.Add(plano)

        ' Denominación / título pieza → Subject + Custom.Denominacion (+ Titulo custom legacy)
        Dim denom = New PropertySyncEntry With {.LogicalName = "DenominacionTitulo", .Value = config.DrawingTitle, .Target = PropertySyncTarget.Both}
        denom.StandardBindings.Add(New PropertyBinding With {.SetName = "SummaryInformation", .PropertyName = "Subject", .AllowCreate = False})
        denom.CustomBindings.Add(New PropertyBinding With {.SetName = "Custom", .PropertyName = "Denominacion", .AllowCreate = True})
        denom.CustomBindings.Add(New PropertyBinding With {.SetName = "Custom", .PropertyName = "Titulo", .AllowCreate = True})
        list.Add(denom)

        Dim project = New PropertySyncEntry With {.LogicalName = "NombreProyecto", .Value = config.ProjectName, .Target = PropertySyncTarget.Both}
        project.StandardBindings.Add(New PropertyBinding With {.SetName = "ProjectInformation", .PropertyName = "Project Name", .AllowCreate = False})
        project.StandardBindings.Add(New PropertyBinding With {.SetName = "ProjectInformation", .PropertyName = "Nombre de proyecto", .AllowCreate = False})
        project.CustomBindings.Add(New PropertyBinding With {.SetName = "Custom", .PropertyName = "Proyecto", .AllowCreate = True})
        list.Add(project)

        Dim material = New PropertySyncEntry With {.LogicalName = "Material", .Value = config.Material, .Target = PropertySyncTarget.Both}
        material.StandardBindings.Add(New PropertyBinding With {.SetName = "MechanicalModeling", .PropertyName = "Material", .AllowCreate = False})
        material.StandardBindings.Add(New PropertyBinding With {.SetName = "ProjectInformation", .PropertyName = "Material", .AllowCreate = False})
        material.CustomBindings.Add(New PropertyBinding With {.SetName = "Custom", .PropertyName = "Material", .AllowCreate = True})
        list.Add(material)

        Dim client = New PropertySyncEntry With {.LogicalName = "EmpresaCliente", .Value = config.ClientName, .Target = PropertySyncTarget.Both}
        client.StandardBindings.Add(New PropertyBinding With {.SetName = "DocumentSummaryInformation", .PropertyName = "Company", .AllowCreate = False})
        client.StandardBindings.Add(New PropertyBinding With {.SetName = "DocumentSummaryInformation", .PropertyName = "Empresa", .AllowCreate = False})
        client.CustomBindings.Add(New PropertyBinding With {.SetName = "Custom", .PropertyName = "Cliente", .AllowCreate = True})
        list.Add(client)

        Dim drawingNum = New PropertySyncEntry With {.LogicalName = "NumeroDocumento", .Value = config.DrawingNumber, .Target = PropertySyncTarget.Both}
        drawingNum.StandardBindings.Add(New PropertyBinding With {.SetName = "ProjectInformation", .PropertyName = "Document Number", .AllowCreate = False})
        drawingNum.StandardBindings.Add(New PropertyBinding With {.SetName = "ProjectInformation", .PropertyName = "Número de documento", .AllowCreate = False})
        drawingNum.CustomBindings.Add(New PropertyBinding With {.SetName = "Custom", .PropertyName = "Numero de plano", .AllowCreate = True})
        list.Add(drawingNum)

        Dim revision = New PropertySyncEntry With {.LogicalName = "Revision", .Value = config.Revision, .Target = PropertySyncTarget.Both}
        revision.StandardBindings.Add(New PropertyBinding With {.SetName = "ProjectInformation", .PropertyName = "Revision", .AllowCreate = False})
        revision.StandardBindings.Add(New PropertyBinding With {.SetName = "ProjectInformation", .PropertyName = "Revisión", .AllowCreate = False})
        revision.CustomBindings.Add(New PropertyBinding With {.SetName = "Custom", .PropertyName = "Revision", .AllowCreate = True})
        list.Add(revision)

        Dim author = New PropertySyncEntry With {.LogicalName = "Autor", .Value = TitleBlockFieldCoordinator.ResolveEffectiveAuthor(config), .Target = PropertySyncTarget.Both}
        author.StandardBindings.Add(New PropertyBinding With {.SetName = "SummaryInformation", .PropertyName = "Author", .AllowCreate = False})
        author.StandardBindings.Add(New PropertyBinding With {.SetName = "SummaryInformation", .PropertyName = "Autor", .AllowCreate = False})
        list.Add(author)

        Dim thickness = New PropertySyncEntry With {.LogicalName = "Espesor", .Value = If(config.Thickness, "").Trim(), .Target = PropertySyncTarget.Both}
        thickness.StandardBindings.Add(New PropertyBinding With {.SetName = "MechanicalModeling", .PropertyName = "Sheet Metal Gauge", .AllowCreate = False})
        thickness.CustomBindings.Add(New PropertyBinding With {.SetName = "Custom", .PropertyName = "Espesor", .AllowCreate = True})
        list.Add(thickness)

        Dim order = New PropertySyncEntry With {.LogicalName = "Pedido", .Value = If(config.Pedido, "").Trim(), .Target = PropertySyncTarget.Both}
        ' Variantes habituales enlazadas en cajetines (incl. PEDIDO / Order). El perfil intenta escribir todas si aplica espejo.
        order.CustomBindings.Add(New PropertyBinding With {.SetName = "Custom", .PropertyName = "Pedido", .AllowCreate = True})
        order.CustomBindings.Add(New PropertyBinding With {.SetName = "Custom", .PropertyName = "PEDIDO", .AllowCreate = True})
        order.CustomBindings.Add(New PropertyBinding With {.SetName = "Custom", .PropertyName = "Order", .AllowCreate = True})
        list.Add(order)

        Dim fechaVal = If(config.FechaPlano, "").Trim()
        If fechaVal <> "" Then
            Dim fecha = New PropertySyncEntry With {.LogicalName = "FechaPlano", .Value = fechaVal, .Target = PropertySyncTarget.Both}
            fecha.CustomBindings.Add(New PropertyBinding With {.SetName = "Custom", .PropertyName = "FechaPlano", .AllowCreate = True})
            list.Add(fecha)
        End If

        Dim pesoVal = If(config.Weight, "").Trim()
        If pesoVal <> "" Then
            Dim peso = New PropertySyncEntry With {.LogicalName = "PesoMeta", .Value = pesoVal, .Target = PropertySyncTarget.Both}
            peso.StandardBindings.Add(New PropertyBinding With {.SetName = "MechanicalModeling", .PropertyName = "Mass", .AllowCreate = False})
            peso.StandardBindings.Add(New PropertyBinding With {.SetName = "ProjectInformation", .PropertyName = "Mass", .AllowCreate = False})
            peso.CustomBindings.Add(New PropertyBinding With {.SetName = "Custom", .PropertyName = "Peso", .AllowCreate = True})
            list.Add(peso)
        End If

        SubAddPartMeta(list, "L", config.PartListL)
        SubAddPartMeta(list, "H", config.PartListH)
        SubAddPartMeta(list, "D", config.PartListD)
        SubAddPartMeta(list, "NombreArchivo", config.PartListNombreArchivo)
        SubAddPartMeta(list, "Cantidad", config.PartListCantidad)

        Return list
    End Function

    Private Shared Sub SubAddPartMeta(list As List(Of PropertySyncEntry), propName As String, value As String)
        If list Is Nothing OrElse String.IsNullOrWhiteSpace(value) Then Return
        Dim e = New PropertySyncEntry With {.LogicalName = "PartMeta_" & propName, .Value = value.Trim(), .Target = PropertySyncTarget.Both}
        e.CustomBindings.Add(New PropertyBinding With {.SetName = "Custom", .PropertyName = propName, .AllowCreate = True})
        list.Add(e)
    End Sub

    Private Shared Function ApplyProfileToOpenModelDocument(doc As Object,
                                                            profile As IEnumerable(Of PropertySyncEntry),
                                                            logger As Logger) As Integer
        Dim updates As Integer = 0
        If doc Is Nothing OrElse profile Is Nothing Then Return updates

        For Each entry In profile
            If entry Is Nothing Then Continue For
            If Not ShouldWriteEntry(entry.Target, PropertySyncTarget.ModelOnly) Then Continue For
            If String.IsNullOrWhiteSpace(entry.Value) Then
                If logger IsNot Nothing Then logger.Log($"[PROPS][WRITE][SKIP] {entry.LogicalName}: vacío")
                Continue For
            End If

            Dim wroteStandard As Boolean = False
            Dim lastStandardReason As String = ""
            For Each binding In entry.StandardBindings
                Dim status As PropertyWriteStatus = PropertyWriteStatus.UnexpectedError
                Dim reason As String = ""
                If TrySetDocumentPropertyDetailed(doc, binding.SetName, binding.PropertyName, entry.Value, binding.AllowCreate, status, reason) Then
                    updates += 1
                    wroteStandard = True
                    If logger IsNot Nothing Then logger.Log($"[PROPS][WRITE] {entry.LogicalName} ← {binding.SetName}.{binding.PropertyName} OK")
                    Exit For
                Else
                    lastStandardReason = $"{binding.SetName}.{binding.PropertyName} -> {status}"
                End If
            Next

            If wroteStandard Then Continue For

            Dim wroteCustom As Boolean = False
            Dim mirrorAllPedidoCustom As Boolean = String.Equals(entry.LogicalName, "Pedido", StringComparison.OrdinalIgnoreCase)
            For Each binding In entry.CustomBindings
                Dim status As PropertyWriteStatus = PropertyWriteStatus.UnexpectedError
                Dim reason As String = ""
                If TrySetDocumentPropertyDetailed(doc, binding.SetName, binding.PropertyName, entry.Value, True, status, reason) Then
                    updates += 1
                    wroteCustom = True
                    If logger IsNot Nothing Then logger.Log($"[PROPS][WRITE] {entry.LogicalName} ← Custom {binding.SetName}.{binding.PropertyName} OK")
                    If Not mirrorAllPedidoCustom Then Exit For
                End If
            Next

            If Not wroteCustom AndAlso logger IsNot Nothing Then
                logger.Log($"[PROPS][WRITE][WARN] {entry.LogicalName} no se pudo escribir. Último intento estándar: {lastStandardReason}")
            End If
        Next

        Return updates
    End Function

    Private Shared Function ApplyProfileToOpenDraft(dftDoc As Object,
                                                    profile As IEnumerable(Of PropertySyncEntry),
                                                    logger As Logger) As Integer
        Dim updates As Integer = 0
        If dftDoc Is Nothing OrElse profile Is Nothing Then Return updates
        For Each entry In profile
            If entry Is Nothing Then Continue For
            If String.IsNullOrWhiteSpace(entry.Value) Then
                If logger IsNot Nothing Then logger.Log($"[PROPS][DFT_OPEN][SKIP] {entry.LogicalName}: valor vacío, no se escribe.")
                Continue For
            End If
            If entry.Target = PropertySyncTarget.ModelOnly Then Continue For
            updates += SetProfileEntryOnOpenDocument(dftDoc, entry, logger, "DFT_OPEN")
        Next
        Return updates
    End Function

    Private Shared Function ApplyProfileToFile(filePath As String,
                                               profile As IEnumerable(Of PropertySyncEntry),
                                               filter As PropertySyncTarget,
                                               logger As Logger,
                                               contextTag As String) As Integer
        If String.IsNullOrWhiteSpace(filePath) OrElse Not IO.File.Exists(filePath) Then Return 0
        If profile Is Nothing Then Return 0

        Dim updated As Integer = 0
        Dim psets As SolidEdgeFileProperties.PropertySets = Nothing
        Try
            psets = New SolidEdgeFileProperties.PropertySets()
            psets.Open(filePath, False)
            For Each entry In profile
                If entry Is Nothing Then Continue For
                If String.IsNullOrWhiteSpace(entry.Value) Then
                    If logger IsNot Nothing Then logger.Log($"[PROPS][{contextTag}][SKIP] {entry.LogicalName}: valor vacío, no se escribe.")
                    Continue For
                End If
                If Not ShouldWriteEntry(entry.Target, filter) Then Continue For

                For Each binding In entry.StandardBindings
                    updated += SetFileProperty(psets, binding.SetName, binding.PropertyName, entry.Value, allowCreate:=binding.AllowCreate, logger:=logger)
                Next
                For Each binding In entry.CustomBindings
                    updated += SetFileProperty(psets, binding.SetName, binding.PropertyName, entry.Value, allowCreate:=binding.AllowCreate, logger:=logger)
                Next
                If logger IsNot Nothing Then logger.Log($"[PROPS][{contextTag}] {entry.LogicalName}='{entry.Value}' procesado.")
            Next
            psets.Save()
            psets.Close()
            If logger IsNot Nothing Then logger.Log($"[PROPS][{contextTag}] Escritura completada en {IO.Path.GetFileName(filePath)}. Campos={updated}")
        Catch ex As Exception
            If logger IsNot Nothing Then logger.LogException($"ApplyProfileToFile[{contextTag}]", ex)
        Finally
            Try
                If psets IsNot Nothing AndAlso Marshal.IsComObject(psets) Then Marshal.ReleaseComObject(psets)
            Catch
            End Try
        End Try
        Return updated
    End Function

    Private Shared Function ApplyProfileToLinkedModelDocuments(dftDoc As Object,
                                                               profile As IEnumerable(Of PropertySyncEntry),
                                                               logger As Logger) As Integer
        Dim updates As Integer = 0
        If dftDoc Is Nothing OrElse profile Is Nothing Then Return updates
        Try
            Dim modelLinks As Object = Nothing
            Try : modelLinks = CallByName(dftDoc, "ModelLinks", CallType.Get) : Catch : End Try
            Dim linkCount As Integer = 0
            Try : linkCount = CInt(CallByName(modelLinks, "Count", CallType.Get)) : Catch : End Try
            If logger IsNot Nothing Then logger.Log($"[PROPS][MODEL_LINK_FILE] ModelLinks para escritura por archivo: {linkCount}")

            Dim processed As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            For i As Integer = 1 To linkCount
                Dim linkObj As Object = GetCollectionItem(modelLinks, i)
                If linkObj Is Nothing Then Continue For

                Dim modelName As String = ""
                Try : modelName = CStr(CallByName(linkObj, "FileName", CallType.Get)) : Catch : End Try
                If String.IsNullOrWhiteSpace(modelName) Then
                    If logger IsNot Nothing Then logger.Log($"[PROPS][MODEL_LINK_FILE][WARN] Link#{i} sin FileName; se omite.")
                    Continue For
                End If
                If logger IsNot Nothing Then logger.Log($"[PROPS][MODEL_LINK_FILE] Link#{i} -> {modelName}")
                If Not IO.File.Exists(modelName) Then
                    If logger IsNot Nothing Then logger.Log($"[PROPS][MODEL_LINK_FILE][WARN] No existe archivo enlazado: {modelName}")
                    Continue For
                End If
                If processed.Contains(modelName) Then Continue For
                processed.Add(modelName)

                updates += ApplyProfileToFile(modelName, profile, PropertySyncTarget.ModelOnly, logger, "MODEL_LINK_FILE")
            Next
        Catch ex As Exception
            If logger IsNot Nothing Then logger.LogException("ApplyProfileToLinkedModelDocuments", ex)
        End Try
        Return updates
    End Function

    Private Shared Function SetProfileEntryOnOpenDocument(dftDoc As Object,
                                                          entry As PropertySyncEntry,
                                                          logger As Logger,
                                                          contextTag As String) As Integer
        Dim updates As Integer = 0
        For Each binding In entry.StandardBindings
            updates += SetDocumentProperty(dftDoc, binding.SetName, binding.PropertyName, entry.Value, logger)
        Next
        For Each binding In entry.CustomBindings
            updates += SetDocumentProperty(dftDoc, binding.SetName, binding.PropertyName, entry.Value, logger)
        Next
        If logger IsNot Nothing Then logger.Log($"[PROPS][{contextTag}] {entry.LogicalName}='{entry.Value}' procesado.")
        Return updates
    End Function

    Private Shared Function TrySetDocumentPropertyDetailed(doc As Object,
                                                           setName As String,
                                                           propName As String,
                                                           value As String,
                                                           allowCreate As Boolean,
                                                           ByRef status As PropertyWriteStatus,
                                                           ByRef reason As String) As Boolean
        status = PropertyWriteStatus.UnexpectedError
        reason = ""
        If doc Is Nothing Then
            status = PropertyWriteStatus.UnexpectedError
            reason = "documento_nulo"
            Return False
        End If
        If String.IsNullOrWhiteSpace(value) Then
            status = PropertyWriteStatus.PropertyNotFound
            reason = "valor_vacio"
            Return False
        End If

        Try
            Dim pset As Object = ResolvePropertySet(doc, setName)
            If pset Is Nothing Then
                Dim memberFromSummary As String = ""
                If TrySetViaSummaryInfo(doc, setName, propName, value, memberFromSummary, reason) Then
                    status = PropertyWriteStatus.Updated
                    reason = $"summaryinfo.{memberFromSummary}"
                    Return True
                End If
                status = PropertyWriteStatus.SetNotFound
                reason = If(String.IsNullOrWhiteSpace(reason), "set_no_encontrado", reason)
                Return False
            End If

            Dim prop As Object = FindNamedItem(pset, {propName})
            If prop IsNot Nothing Then
                Try
                    Dim memberUsed As String = ""
                    If Not TrySetComPropertyValue(prop, value, memberUsed, reason) Then
                        Throw New MissingMemberException($"No se pudo escribir usando miembros estándar. {reason}")
                    End If
                    status = PropertyWriteStatus.Updated
                    reason = $"actualizada_via_{memberUsed}"
                    Return True
                Catch exSet As Exception
                    Dim memberFromSummary As String = ""
                    Dim reasonSummary As String = ""
                    If TrySetViaSummaryInfo(doc, setName, propName, value, memberFromSummary, reasonSummary) Then
                        status = PropertyWriteStatus.Updated
                        reason = $"summaryinfo.{memberFromSummary}"
                        Return True
                    End If
                    status = PropertyWriteStatus.ReadOnlyOrTypeMismatch
                    reason = $"{exSet.GetType().Name}: {exSet.Message}"
                    Return False
                End Try
            End If

            If Not allowCreate Then
                Dim memberFromSummary As String = ""
                Dim reasonSummary As String = ""
                If TrySetViaSummaryInfo(doc, setName, propName, value, memberFromSummary, reasonSummary) Then
                    status = PropertyWriteStatus.Updated
                    reason = $"summaryinfo.{memberFromSummary}"
                    Return True
                End If
                status = PropertyWriteStatus.PropertyNotFound
                reason = "propiedad_no_encontrada"
                Return False
            End If

            Try
                CallByName(pset, "Add", CallType.Method, value, propName)
                status = PropertyWriteStatus.Added
                reason = "creada"
                Return True
            Catch
                Try
                    CallByName(pset, "Add", CallType.Method, propName, value)
                    status = PropertyWriteStatus.Added
                    reason = "creada_alt"
                    Return True
                Catch exAdd As Exception
                    status = PropertyWriteStatus.ReadOnlyOrTypeMismatch
                    reason = exAdd.Message
                    Return False
                End Try
            End Try
        Catch ex As Exception
            status = PropertyWriteStatus.UnexpectedError
            reason = ex.Message
            Return False
        End Try
    End Function

    Private Shared Function TrySetComPropertyValue(propObj As Object,
                                                   value As Object,
                                                   ByRef memberUsed As String,
                                                   ByRef reason As String) As Boolean
        memberUsed = ""
        reason = ""
        If propObj Is Nothing Then
            reason = "propiedad_nula"
            Return False
        End If

        Dim diagnostics As New List(Of String)()

        Try
            CallByName(propObj, "Value", CallType.Let, value)
            memberUsed = "Value(Let)"
            Return True
        Catch ex As Exception
            diagnostics.Add($"Value(Let):{ex.GetType().Name}:{ex.Message}")
        End Try
        Try
            CallByName(propObj, "Value", CallType.Set, value)
            memberUsed = "Value(Set)"
            Return True
        Catch ex As Exception
            diagnostics.Add($"Value(Set):{ex.GetType().Name}:{ex.Message}")
        End Try
        Try
            CallByName(propObj, "set_Value", CallType.Method, value)
            memberUsed = "set_Value(Method)"
            Return True
        Catch ex As Exception
            diagnostics.Add($"set_Value(Method):{ex.GetType().Name}:{ex.Message}")
        End Try
        Try
            CallByName(propObj, "PutValue", CallType.Method, value)
            memberUsed = "PutValue(Method)"
            Return True
        Catch ex As Exception
            diagnostics.Add($"PutValue(Method):{ex.GetType().Name}:{ex.Message}")
        End Try

        reason = String.Join(" | ", diagnostics)
        Return False
    End Function

    Private Shared Function TrySetViaSummaryInfo(doc As Object,
                                                 setName As String,
                                                 propName As String,
                                                 value As String,
                                                 ByRef memberUsed As String,
                                                 ByRef reason As String) As Boolean
        memberUsed = ""
        reason = ""
        If doc Is Nothing Then Return False

        Dim summaryInfo As Object = Nothing
        Try : summaryInfo = CallByName(doc, "SummaryInfo", CallType.Get) : Catch : End Try
        If summaryInfo Is Nothing Then
            reason = "summaryinfo_no_disponible"
            Return False
        End If

        Dim setKey As String = NormalizeKey(setName)
        Dim propKey As String = NormalizeKey(propName)
        Dim candidates As New List(Of String)

        If setKey = "summaryinformation" Then
            If propKey = "title" OrElse propKey = "documenttitle" OrElse propKey = "titulo" Then candidates.Add("Title")
            If propKey = "author" OrElse propKey = "autor" Then candidates.Add("Author")
        ElseIf setKey = "projectinformation" Then
            If propKey = "projectname" OrElse propKey = "nombredeproyecto" OrElse propKey = "project" Then candidates.Add("ProjectName")
            If propKey = "revision" OrElse propKey = "revisionnumber" Then candidates.Add("RevisionNumber")
            If propKey = "documentnumber" OrElse propKey = "numerodedocumento" OrElse propKey = "partnumber" Then candidates.Add("DocumentNumber")
        ElseIf setKey = "documentsummaryinformation" Then
            If propKey = "company" OrElse propKey = "empresa" Then
                candidates.Add("CompanyName")
                candidates.Add("Company")
            End If
        End If

        For Each member In candidates
            Try
                CallByName(summaryInfo, member, CallType.Let, value)
                memberUsed = member
                Return True
            Catch ex As Exception
                reason = $"{member}:{ex.GetType().Name}"
            End Try
        Next
        Return False
    End Function

    Private Shared Function NormalizeKey(text As String) As String
        If String.IsNullOrWhiteSpace(text) Then Return ""
        Dim s As String = text.Trim().ToLowerInvariant()
        s = s.Replace("á", "a").Replace("é", "e").Replace("í", "i").Replace("ó", "o").Replace("ú", "u").Replace("ü", "u").Replace("ñ", "n")
        Dim sb As New StringBuilder()
        For Each ch As Char In s
            If Char.IsLetterOrDigit(ch) Then sb.Append(ch)
        Next
        Return sb.ToString()
    End Function

    Private Shared Sub InspectTitleBlockCollectionSource(sheet As Object,
                                                         collectionName As String,
                                                         logger As Logger,
                                                         sectionIndex As Integer,
                                                         sheetIndex As Integer)
        If sheet Is Nothing Then Return
        Dim coll As Object = Nothing
        Try : coll = CallByName(sheet, collectionName, CallType.Get) : Catch : End Try
        If coll Is Nothing Then Return
        Dim count As Integer = 0
        Try : count = CInt(CallByName(coll, "Count", CallType.Get)) : Catch : End Try
        For i As Integer = 1 To count
            Dim tb As Object = GetCollectionItem(coll, i)
            If tb Is Nothing Then Continue For
            Dim tbType As String = TypeName(tb)
            logger.Log($"[TPL][SRC] Section={sectionIndex} Sheet={sheetIndex} {collectionName}#{i} Type={tbType}")
            Dim textBoxes As Object = Nothing
            Try : textBoxes = CallByName(tb, "TextBoxes", CallType.Get) : Catch : End Try
            If textBoxes Is Nothing Then Continue For
            Dim txCount As Integer = 0
            Try : txCount = CInt(CallByName(textBoxes, "Count", CallType.Get)) : Catch : End Try
            For txi As Integer = 1 To txCount
                Dim tx As Object = GetCollectionItem(textBoxes, txi)
                If tx Is Nothing Then Continue For
                LogTextObjectSource(tx, logger, $"Section={sectionIndex} Sheet={sheetIndex} {collectionName}#{i}.TextBox#{txi}")
            Next
        Next
    End Sub

    Private Shared Sub InspectTextCollectionSource(sheet As Object,
                                                   collectionName As String,
                                                   logger As Logger,
                                                   sectionIndex As Integer,
                                                   sheetIndex As Integer)
        If sheet Is Nothing Then Return
        Dim coll As Object = Nothing
        Try : coll = CallByName(sheet, collectionName, CallType.Get) : Catch : End Try
        If coll Is Nothing Then Return
        Dim count As Integer = 0
        Try : count = CInt(CallByName(coll, "Count", CallType.Get)) : Catch : End Try
        For i As Integer = 1 To count
            Dim tx As Object = GetCollectionItem(coll, i)
            If tx Is Nothing Then Continue For
            LogTextObjectSource(tx, logger, $"Section={sectionIndex} Sheet={sheetIndex} {collectionName}#{i}")
        Next
    End Sub

    Private Shared Sub LogTextObjectSource(obj As Object, logger As Logger, location As String)
        If obj Is Nothing OrElse logger Is Nothing Then Return
        Dim objType As String = TypeName(obj)
        Dim objName As String = GetStringMember(obj, "Name")
        Dim txt As String = GetStringMember(obj, "Text")
        Dim value As String = GetStringMember(obj, "Value")
        Dim formula As String = GetStringMember(obj, "Formula")
        Dim visible As String = FirstNonEmpty(value, txt)
        Dim token As String = FirstNonEmpty(formula, value, txt)
        Dim srcType As String = "Texto fijo"
        Dim srcGuess As String = "No detectable"
        If Not String.IsNullOrWhiteSpace(token) Then
            Dim up As String = token.ToUpperInvariant()
            If up.Contains("|R") Then
                srcType = "PropertyText"
                srcGuess = "ModelLink (R#)"
            ElseIf up.Contains("CUSTOM") Then
                srcType = "PropertyText"
                srcGuess = "Custom"
            ElseIf up.Contains("SUMMARYINFORMATION") Then
                srcType = "PropertyText"
                srcGuess = "SummaryInformation"
            ElseIf up.Contains("PROJECTINFORMATION") Then
                srcType = "PropertyText"
                srcGuess = "ProjectInformation"
            ElseIf up.Contains("DOCUMENTSUMMARYINFORMATION") Then
                srcType = "PropertyText"
                srcGuess = "DocumentSummaryInformation"
            ElseIf up.Contains("MECHANICALMODELING") Then
                srcType = "PropertyText"
                srcGuess = "MechanicalModeling"
            ElseIf up.Contains("ACTIVE DOCUMENT") OrElse up.Contains("%{") Then
                srcType = "PropertyText"
                srcGuess = "Active Document"
            End If
        End If
        logger.Log($"[TPL][SRC] {location} Type={objType} Name='{objName}' Visible='{visible}' Formula='{formula}' SourceType={srcType} SourceGuess={srcGuess}")
    End Sub

    Private Shared Function ShouldWriteEntry(entryTarget As PropertySyncTarget, filter As PropertySyncTarget) As Boolean
        If filter = PropertySyncTarget.Both Then Return True
        If entryTarget = PropertySyncTarget.Both Then Return True
        Return entryTarget = filter
    End Function

    Private Shared Function SetFileProperty(psets As SolidEdgeFileProperties.PropertySets,
                                            setName As String,
                                            propName As String,
                                            value As String,
                                            allowCreate As Boolean,
                                            logger As Logger) As Integer
        If psets Is Nothing Then Return 0
        If value Is Nothing Then value = ""
        If value.Trim() = "" Then Return 0

        Dim setObj As Object = Nothing
        Dim propObj As Object = Nothing
        Try
            setObj = ResolveFilePropertySet(psets, setName)
            If setObj Is Nothing Then Return 0

            Dim found As Boolean = False
            Dim cnt As Integer = 0
            Try : cnt = CInt(CallByName(setObj, "Count", CallType.Get)) : Catch : End Try
            For i As Integer = 1 To cnt
                Try
                    propObj = GetCollectionItem(setObj, i)
                    If propObj Is Nothing Then Continue For
                    Dim n As String = ""
                    Try : n = CStr(CallByName(propObj, "Name", CallType.Get)) : Catch : End Try
                    If String.Equals(n, propName, StringComparison.OrdinalIgnoreCase) Then
                        Dim memberUsed As String = ""
                        Dim reason As String = ""
                        If TrySetComPropertyValue(propObj, value, memberUsed, reason) Then
                            If logger IsNot Nothing Then logger.Log($"[PROPS][FILE] {setName}.{propName}='{value}' (update:{memberUsed})")
                            found = True
                            Exit For
                        Else
                            If logger IsNot Nothing Then logger.Log($"[PROPS][FILE][WARN] update falló {setName}.{propName} -> {reason}")
                        End If
                    End If
                Finally
                    Try
                        If propObj IsNot Nothing AndAlso Marshal.IsComObject(propObj) Then Marshal.ReleaseComObject(propObj)
                    Catch
                    End Try
                    propObj = Nothing
                End Try
            Next

            If Not found AndAlso allowCreate Then
                Try
                    CallByName(setObj, "Add", CallType.Method, propName, value)
                    If logger IsNot Nothing Then logger.Log($"[PROPS][FILE] {setName}.{propName}='{value}' (add)")
                    Return 1
                Catch
                    Try
                        CallByName(setObj, "Add", CallType.Method, value, propName)
                        If logger IsNot Nothing Then logger.Log($"[PROPS][FILE] {setName}.{propName}='{value}' (add-alt)")
                        Return 1
                    Catch
                    End Try
                End Try
            End If

            If found Then Return 1
            If logger IsNot Nothing Then logger.Log($"[PROPS][FILE][WARN] No se pudo escribir {setName}.{propName}")
            Return 0
        Catch
            If logger IsNot Nothing Then logger.Log($"[PROPS][FILE][WARN] Fallo en {setName}.{propName}")
            Return 0
        Finally
            Try
                If setObj IsNot Nothing AndAlso Marshal.IsComObject(setObj) Then Marshal.ReleaseComObject(setObj)
            Catch
            End Try
        End Try
    End Function

    ''' <summary>Lectura segura desde PropertySets de archivo (solo lectura). Usado por trazabilidad UI.</summary>
    Friend Shared Function TryGetFilePropertyValue(psets As SolidEdgeFileProperties.PropertySets,
                                                   setName As String,
                                                   propName As String) As String
        If psets Is Nothing OrElse String.IsNullOrWhiteSpace(propName) Then Return ""
        Dim setObj As Object = Nothing
        Dim propObj As Object = Nothing
        Try
            setObj = ResolveFilePropertySet(psets, setName)
            If setObj Is Nothing Then Return ""
            Dim cnt As Integer = 0
            Try : cnt = CInt(CallByName(setObj, "Count", CallType.Get)) : Catch : End Try
            For i As Integer = 1 To cnt
                Try
                    propObj = GetCollectionItem(setObj, i)
                    If propObj Is Nothing Then Continue For
                    Dim n As String = ""
                    Try : n = CStr(CallByName(propObj, "Name", CallType.Get)) : Catch : End Try
                    If String.Equals(n, propName, StringComparison.OrdinalIgnoreCase) Then
                        Dim v As Object = Nothing
                        Try : v = CallByName(propObj, "Value", CallType.Get) : Catch : End Try
                        If v IsNot Nothing Then
                            Dim s As String = v.ToString().Trim()
                            If s <> "" Then Return s
                        End If
                        Return ""
                    End If
                Finally
                    Try
                        If propObj IsNot Nothing AndAlso Marshal.IsComObject(propObj) Then Marshal.ReleaseComObject(propObj)
                    Catch
                    End Try
                    propObj = Nothing
                End Try
            Next
        Catch
        Finally
            Try
                If setObj IsNot Nothing AndAlso Marshal.IsComObject(setObj) Then Marshal.ReleaseComObject(setObj)
            Catch
            End Try
        End Try
        Return ""
    End Function

    Private Shared Function TryGetOpenDocumentByPath(app As Application, filePath As String) As Object
        If app Is Nothing OrElse String.IsNullOrWhiteSpace(filePath) Then Return Nothing
        Try
            Dim docs As Object = Nothing
            Try : docs = CallByName(app, "Documents", CallType.Get) : Catch : End Try
            If docs Is Nothing Then Return Nothing
            Dim count As Integer = 0
            Try : count = CInt(CallByName(docs, "Count", CallType.Get)) : Catch : End Try
            For i As Integer = 1 To count
                Dim d As Object = GetCollectionItem(docs, i)
                If d Is Nothing Then Continue For
                Dim fullName As String = ""
                Try : fullName = CStr(CallByName(d, "FullName", CallType.Get)) : Catch : End Try
                If String.Equals(fullName, filePath, StringComparison.OrdinalIgnoreCase) Then Return d
            Next
        Catch
        End Try
        Return Nothing
    End Function

    Friend Shared Function ResolveFilePropertySetForInspection(psets As Object, canonicalName As String) As Object
        Return ResolveFilePropertySet(psets, canonicalName)
    End Function

    Friend Shared Function GetCollectionItemForInspection(collection As Object, key As Object) As Object
        Return GetCollectionItem(collection, key)
    End Function

    Private Shared Function ResolveFilePropertySet(psets As Object, canonicalName As String) As Object
        If psets Is Nothing OrElse String.IsNullOrWhiteSpace(canonicalName) Then Return Nothing
        Dim candidates As New List(Of String)
        Select Case canonicalName.ToLowerInvariant()
            Case "summaryinformation"
                candidates.AddRange({"SummaryInformation", "Summary Information", "Información de resumen", "Informacion de resumen"})
            Case "extendedsummaryinformation"
                candidates.AddRange({"ExtendedSummaryInformation", "Extended Summary Information", "Información de resumen extendida", "Informacion de resumen extendida"})
            Case "projectinformation"
                candidates.AddRange({"ProjectInformation", "Project Information", "Información del proyecto", "Informacion del proyecto"})
            Case "documentsummaryinformation"
                candidates.AddRange({"DocumentSummaryInformation", "Document Summary Information", "Resumen del documento", "Información del documento", "Informacion del documento"})
            Case "mechanicalmodeling"
                candidates.AddRange({"MechanicalModeling", "Mechanical Modeling", "Modelado mecánico", "Modelado mecanico"})
            Case "custom"
                candidates.AddRange({"Custom", "CustomInformation", "Custom Properties", "Propiedades personalizadas", "Personalizado"})
            Case Else
                candidates.Add(canonicalName)
        End Select

        Try
            For Each c In candidates
                Try
                    Dim byName As Object = CallByName(psets, "Item", CallType.Method, c)
                    If byName IsNot Nothing Then Return byName
                Catch
                End Try
            Next
        Catch
        End Try

        Try
            Dim count As Integer = CInt(CallByName(psets, "Count", CallType.Get))
            For i As Integer = 1 To count
                Dim setObj As Object = GetCollectionItem(psets, i)
                If setObj Is Nothing Then Continue For
                Dim setName As String = ""
                Try : setName = CStr(CallByName(setObj, "Name", CallType.Get)) : Catch : End Try
                If String.IsNullOrWhiteSpace(setName) Then Continue For
                For Each c In candidates
                    If setName.Equals(c, StringComparison.OrdinalIgnoreCase) Then Return setObj
                    If setName.IndexOf(c, StringComparison.OrdinalIgnoreCase) >= 0 Then Return setObj
                Next
            Next
        Catch
        End Try
        Return Nothing
    End Function

    Private Shared Function TryConnectApplication(showSolidEdge As Boolean,
                                                  logger As Logger,
                                                  ByRef app As Application,
                                                  ByRef createdByUs As Boolean) As Boolean
        app = Nothing
        createdByUs = False
        Try
            app = CType(Marshal.GetActiveObject("SolidEdge.Application"), Application)
            If logger IsNot Nothing Then logger.Log("Solid Edge: instancia existente detectada.")
        Catch
            Dim t = Type.GetTypeFromProgID("SolidEdge.Application")
            app = CType(Activator.CreateInstance(t), Application)
            createdByUs = True
            If logger IsNot Nothing Then logger.Log("Solid Edge: nueva instancia COM creada.")
        End Try

        If app Is Nothing Then Return False
        app.Visible = showSolidEdge
        app.DisplayAlerts = False
        If logger IsNot Nothing Then logger.Log($"Solid Edge modo oculto={Not showSolidEdge}, DisplayAlerts=False")
        Return True
    End Function

    Private Shared Function GetDocumentProperty(doc As Object, setName As String, propNames As IEnumerable(Of String)) As String
        Try
            Dim pset As Object = ResolvePropertySet(doc, setName)
            If pset Is Nothing Then Return ""
            For Each pn In propNames
                Try
                    Dim prop As Object = FindNamedItem(pset, {pn})
                    If prop Is Nothing Then Continue For
                    Dim value As Object = GetPropertyValueSafe(prop)
                    If value IsNot Nothing Then
                        Dim s As String = value.ToString().Trim()
                        If s <> "" Then Return s
                    End If
                Catch
                End Try
            Next
        Catch
        End Try
        Return ""
    End Function

    Private Shared Function SetDocumentProperty(doc As Object, setName As String, propName As String, value As String, logger As Logger) As Integer
        If value Is Nothing Then value = ""
        If value.Trim() = "" Then Return 0
        Try
            Dim pset As Object = ResolvePropertySet(doc, setName)
            If pset Is Nothing Then
                If logger IsNot Nothing Then
                    If setName.Equals("Custom", StringComparison.OrdinalIgnoreCase) Then
                        logger.Log($"[PROPS][WARN] PropertySet no encontrado: {setName}")
                    Else
                        logger.Log($"[PROPS][INFO] PropertySet no disponible en este documento: {setName}")
                    End If
                End If
                Return 0
            End If
            Try
                Dim prop As Object = FindNamedItem(pset, {propName})
                If prop IsNot Nothing Then
                    Dim memberUsed As String = ""
                    Dim reason As String = ""
                    If TrySetComPropertyValue(prop, value, memberUsed, reason) Then
                        If logger IsNot Nothing Then logger.Log($"[PROPS] {setName}.{propName} = '{value}' (update:{memberUsed})")
                        Return 1
                    End If
                    If logger IsNot Nothing Then logger.Log($"[PROPS][WARN] Falló update {setName}.{propName} -> {reason}")
                End If
            Catch
            End Try
            Try
                Try
                    CallByName(pset, "Add", CallType.Method, value, propName)
                    If logger IsNot Nothing Then logger.Log($"[PROPS] {setName}.{propName} = '{value}' (add)")
                    Return 1
                Catch
                    Try
                        CallByName(pset, "Add", CallType.Method, propName, value)
                        If logger IsNot Nothing Then logger.Log($"[PROPS] {setName}.{propName} = '{value}' (add-alt)")
                        Return 1
                    Catch
                        If logger IsNot Nothing Then
                            If setName.Equals("Custom", StringComparison.OrdinalIgnoreCase) Then
                                logger.Log($"[PROPS][WARN] No se pudo escribir {setName}.{propName}")
                            Else
                                logger.Log($"[PROPS][INFO] No se pudo escribir {setName}.{propName} en documento abierto.")
                            End If
                        End If
                        Return 0
                    End Try
                End Try
            Catch
                Return 0
            End Try
        Catch
            Return 0
        End Try
    End Function

    Private Shared Function ResolvePropertySet(doc As Object, canonicalName As String) As Object
        Dim candidates As New List(Of String)
        Select Case canonicalName.ToLowerInvariant()
            Case "summaryinformation"
                candidates.AddRange({"SummaryInformation", "Summary Information", "Información de resumen", "Informacion de resumen"})
            Case "projectinformation"
                candidates.AddRange({"ProjectInformation", "Project Information", "Información del proyecto", "Informacion del proyecto"})
            Case "documentsummaryinformation"
                candidates.AddRange({"DocumentSummaryInformation", "Document Summary Information", "Resumen del documento", "Información del documento", "Informacion del documento"})
            Case "mechanicalmodeling"
                candidates.AddRange({"MechanicalModeling", "Mechanical Modeling", "Modelado mecánico", "Modelado mecanico"})
            Case "custom"
                candidates.AddRange({"Custom", "CustomInformation", "Custom Properties", "Propiedades personalizadas", "Personalizado"})
            Case Else
                candidates.Add(canonicalName)
        End Select

        Dim collections As New List(Of Object)
        Try : collections.Add(CallByName(doc, "Properties", CallType.Get)) : Catch : End Try
        Try : collections.Add(CallByName(doc, "PropertySets", CallType.Get)) : Catch : End Try

        For Each sets In collections
            If sets Is Nothing Then Continue For
            Dim psetByName As Object = FindNamedItem(sets, candidates)
            If psetByName IsNot Nothing Then Return psetByName

            Try
                Dim count As Integer = CInt(CallByName(sets, "Count", CallType.Get))
                For i As Integer = 1 To count
                    Dim pset = GetCollectionItem(sets, i)
                    Dim n As String = ""
                    Try : n = CStr(CallByName(pset, "Name", CallType.Get)) : Catch : End Try
                    For Each c In candidates
                        If Not String.IsNullOrWhiteSpace(n) AndAlso n.IndexOf(c, StringComparison.OrdinalIgnoreCase) >= 0 Then
                            Return pset
                        End If
                    Next
                Next
            Catch
            End Try
        Next

        Return Nothing
    End Function

    Private Shared Function GetPropertyValueSafe(prop As Object) As Object
        If prop Is Nothing Then Return Nothing
        Try
            Return CallByName(prop, "Value", CallType.Get)
        Catch
            Return Nothing
        End Try
    End Function

    Private Shared Function ApplyPropertiesToTitleBlocks(dftDoc As Object, config As JobConfiguration, logger As Logger) As Integer
        Dim updates As Integer = 0
        Try
            Dim sections As Object = Nothing
            Try : sections = CallByName(dftDoc, "Sections", CallType.Get) : Catch : End Try
            If sections Is Nothing Then Return 0

            Dim sectionCount As Integer = 0
            Try : sectionCount = CInt(CallByName(sections, "Count", CallType.Get)) : Catch : End Try
            If logger IsNot Nothing Then logger.Log($"[TPL] ApplyPropertiesToTitleBlocks: Secciones={sectionCount}")
            For si As Integer = 1 To sectionCount
                Dim section As Object = GetCollectionItem(sections, si)
                If section Is Nothing Then Continue For
                Dim sheets As Object = Nothing
                Try : sheets = CallByName(section, "Sheets", CallType.Get) : Catch : End Try
                If sheets Is Nothing Then Continue For

                Dim sheetCount As Integer = 0
                Try : sheetCount = CInt(CallByName(sheets, "Count", CallType.Get)) : Catch : End Try
                If logger IsNot Nothing Then logger.Log($"[TPL] ApplyPropertiesToTitleBlocks: Section={si} Sheets={sheetCount}")
                For shi As Integer = 1 To sheetCount
                    Dim sheet As Object = GetCollectionItem(sheets, shi)
                    If sheet Is Nothing Then Continue For
                    Dim titleBlocks As Object = Nothing
                    Try : titleBlocks = CallByName(sheet, "TitleBlocks", CallType.Get) : Catch : End Try
                    If titleBlocks Is Nothing Then Continue For

                    Dim tbCount As Integer = 0
                    Try : tbCount = CInt(CallByName(titleBlocks, "Count", CallType.Get)) : Catch : End Try
                    If logger IsNot Nothing Then logger.Log($"[TPL] ApplyPropertiesToTitleBlocks: Section={si} Sheet={shi} TitleBlocks={tbCount}")
                    For tbi As Integer = 1 To tbCount
                        Dim tb As Object = GetCollectionItem(titleBlocks, tbi)
                        If tb Is Nothing Then Continue For
                        Dim textBoxes As Object = Nothing
                        Try : textBoxes = CallByName(tb, "TextBoxes", CallType.Get) : Catch : End Try
                        If textBoxes Is Nothing Then Continue For

                        Dim txCount As Integer = 0
                        Try : txCount = CInt(CallByName(textBoxes, "Count", CallType.Get)) : Catch : End Try
                        If logger IsNot Nothing Then logger.Log($"[TPL] ApplyPropertiesToTitleBlocks: Section={si} Sheet={shi} TitleBlock={tbi} TextBoxes={txCount}")
                        For txi As Integer = 1 To txCount
                            Dim tx As Object = GetCollectionItem(textBoxes, txi)
                            Dim newValue As String = ResolveTitleBlockValue(tx, config)
                            If String.IsNullOrWhiteSpace(newValue) Then Continue For
                            If SetTextBoxValue(tx, newValue) Then
                                updates += 1
                            End If
                        Next
                    Next
                Next
            Next
        Catch ex As Exception
            If logger IsNot Nothing Then logger.LogException("ApplyPropertiesToTitleBlocks", ex)
        End Try
        If logger IsNot Nothing Then logger.Log($"[PROPS] Campos aplicados directamente en cajetin: {updates}")
        Return updates
    End Function

    Private Shared Function StampTitleBlockWithDiagnosticKeywords(dftDoc As Object, logger As Logger) As Integer
        Dim updates As Integer = 0
        Dim tokens As String() = {"KW_CLIENTE", "KW_PROYECTO", "KW_TITULO", "KW_MATERIAL"}
        Dim tokenIndex As Integer = 0

        Try
            Dim sections As Object = Nothing
            Try : sections = CallByName(dftDoc, "Sections", CallType.Get) : Catch : End Try
            If sections Is Nothing Then Return 0

            Dim sectionCount As Integer = 0
            Try : sectionCount = CInt(CallByName(sections, "Count", CallType.Get)) : Catch : End Try
            For si As Integer = 1 To sectionCount
                Dim section As Object = GetCollectionItem(sections, si)
                If section Is Nothing Then Continue For

                Dim sheets As Object = Nothing
                Try : sheets = CallByName(section, "Sheets", CallType.Get) : Catch : End Try
                If sheets Is Nothing Then Continue For

                Dim sheetCount As Integer = 0
                Try : sheetCount = CInt(CallByName(sheets, "Count", CallType.Get)) : Catch : End Try
                For shi As Integer = 1 To sheetCount
                    Dim sheet As Object = GetCollectionItem(sheets, shi)
                    If sheet Is Nothing Then Continue For

                    Dim titleBlocks As Object = Nothing
                    Try : titleBlocks = CallByName(sheet, "TitleBlocks", CallType.Get) : Catch : End Try
                    If titleBlocks Is Nothing Then Continue For

                    Dim tbCount As Integer = 0
                    Try : tbCount = CInt(CallByName(titleBlocks, "Count", CallType.Get)) : Catch : End Try
                    For tbi As Integer = 1 To tbCount
                        Dim tb As Object = GetCollectionItem(titleBlocks, tbi)
                        If tb Is Nothing Then Continue For

                        Dim textBoxes As Object = Nothing
                        Try : textBoxes = CallByName(tb, "TextBoxes", CallType.Get) : Catch : End Try
                        If textBoxes Is Nothing Then Continue For

                        Dim txCount As Integer = 0
                        Try : txCount = CInt(CallByName(textBoxes, "Count", CallType.Get)) : Catch : End Try
                        For txi As Integer = 1 To txCount
                            Dim tx As Object = GetCollectionItem(textBoxes, txi)
                            If tx Is Nothing Then Continue For
                            tokenIndex += 1
                            Dim token As String = $"{tokens((tokenIndex - 1) Mod tokens.Length)}_{tokenIndex:00}"
                            If SetTextBoxValue(tx, token) Then
                                updates += 1
                            End If
                        Next
                    Next
                Next
            Next
        Catch ex As Exception
            If logger IsNot Nothing Then logger.LogException("StampTitleBlockWithDiagnosticKeywords", ex)
        End Try

        Return updates
    End Function

    Public Shared Sub LogTemplateTitleBlockBindings(dftDoc As Object, logger As Logger)
        If dftDoc Is Nothing OrElse logger Is Nothing Then Return
        Try
            Dim sections As Object = Nothing
            Try : sections = CallByName(dftDoc, "Sections", CallType.Get) : Catch : End Try
            If sections Is Nothing Then
                logger.Log("[TPL] No se pudo leer Sections del DFT.")
                Return
            End If

            Dim sectionCount As Integer = 0
            Try : sectionCount = CInt(CallByName(sections, "Count", CallType.Get)) : Catch : End Try
            logger.Log($"[TPL] Analizando cajetín del template. Secciones={sectionCount}")
            For si As Integer = 1 To sectionCount
                Dim section As Object = GetCollectionItem(sections, si)
                If section Is Nothing Then Continue For
                Dim sheets As Object = Nothing
                Try : sheets = CallByName(section, "Sheets", CallType.Get) : Catch : End Try
                If sheets Is Nothing Then Continue For

                Dim sheetCount As Integer = 0
                Try : sheetCount = CInt(CallByName(sheets, "Count", CallType.Get)) : Catch : End Try
                For shi As Integer = 1 To sheetCount
                    Dim sheet As Object = GetCollectionItem(sheets, shi)
                    If sheet Is Nothing Then Continue For
                    Dim titleBlocks As Object = Nothing
                    Try : titleBlocks = CallByName(sheet, "TitleBlocks", CallType.Get) : Catch : End Try
                    If titleBlocks Is Nothing Then Continue For

                    Dim tbCount As Integer = 0
                    Try : tbCount = CInt(CallByName(titleBlocks, "Count", CallType.Get)) : Catch : End Try
                    For tbi As Integer = 1 To tbCount
                        Dim tb As Object = GetCollectionItem(titleBlocks, tbi)
                        If tb Is Nothing Then Continue For
                        Dim textBoxes As Object = Nothing
                        Try : textBoxes = CallByName(tb, "TextBoxes", CallType.Get) : Catch : End Try
                        If textBoxes Is Nothing Then Continue For

                        Dim txCount As Integer = 0
                        Try : txCount = CInt(CallByName(textBoxes, "Count", CallType.Get)) : Catch : End Try
                        logger.Log($"[TPL] Section={si} Sheet={shi} TitleBlock={tbi} TextBoxes={txCount}")
                        For txi As Integer = 1 To txCount
                            Dim tx As Object = GetCollectionItem(textBoxes, txi)
                            If tx Is Nothing Then Continue For
                            Dim dump As String = BuildTextBoxDiagnosticDump(tx)
                            If Not String.IsNullOrWhiteSpace(dump) Then
                                logger.Log($"[TPL] TX#{txi}: {dump}")
                            End If
                        Next
                    Next
                Next
            Next
        Catch ex As Exception
            logger.LogException("LogTemplateTitleBlockBindings", ex)
        End Try
    End Sub

    Private Shared Function BuildTextBoxDiagnosticDump(textBoxObj As Object) As String
        If textBoxObj Is Nothing Then Return ""
        Dim parts As New List(Of String)()
        Dim keys As String() = {"Name", "Title", "PromptText", "PrimaryCaption", "SecondaryCaption", "Text", "Value", "Formula"}
        For Each k In keys
            Dim v As String = GetStringMember(textBoxObj, k)
            If Not String.IsNullOrWhiteSpace(v) Then
                Dim safeVal As String = v.Replace(vbCr, " ").Replace(vbLf, " ")
                parts.Add($"{k}='{safeVal}'")
            End If
        Next
        Return String.Join(" | ", parts)
    End Function

    Private Shared Function ApplyPropertiesToSheetTexts(dftDoc As Object, config As JobConfiguration, logger As Logger) As Integer
        Dim updates As Integer = 0
        Try
            Dim sections As Object = Nothing
            Try : sections = CallByName(dftDoc, "Sections", CallType.Get) : Catch : End Try
            If sections Is Nothing Then Return 0

            Dim sectionCount As Integer = 0
            Try : sectionCount = CInt(CallByName(sections, "Count", CallType.Get)) : Catch : End Try
            For si As Integer = 1 To sectionCount
                Dim section As Object = GetCollectionItem(sections, si)
                If section Is Nothing Then Continue For

                Dim sheets As Object = Nothing
                Try : sheets = CallByName(section, "Sheets", CallType.Get) : Catch : End Try
                If sheets Is Nothing Then Continue For

                Dim sheetCount As Integer = 0
                Try : sheetCount = CInt(CallByName(sheets, "Count", CallType.Get)) : Catch : End Try
                For shi As Integer = 1 To sheetCount
                    Dim sheet As Object = GetCollectionItem(sheets, shi)
                    If sheet Is Nothing Then Continue For
                    updates += ApplyPropertiesToTextCollection(sheet, "TextBoxes", config, logger, si, shi)
                    updates += ApplyPropertiesToTextCollection(sheet, "Texts", config, logger, si, shi)
                Next
            Next
        Catch ex As Exception
            If logger IsNot Nothing Then logger.LogException("ApplyPropertiesToSheetTexts", ex)
        End Try
        If logger IsNot Nothing Then logger.Log($"[TPL] Campos actualizados por texto de hoja: {updates}")
        Return updates
    End Function

    Private Shared Function ApplyPropertiesToTextCollection(owner As Object,
                                                            collectionName As String,
                                                            config As JobConfiguration,
                                                            logger As Logger,
                                                            sectionIndex As Integer,
                                                            sheetIndex As Integer) As Integer
        Dim updates As Integer = 0
        If owner Is Nothing Then Return 0

        Dim coll As Object = Nothing
        Try : coll = CallByName(owner, collectionName, CallType.Get) : Catch : End Try
        If coll Is Nothing Then Return 0

        Dim count As Integer = 0
        Try : count = CInt(CallByName(coll, "Count", CallType.Get)) : Catch : End Try
        If count <= 0 Then Return 0

        If logger IsNot Nothing Then logger.Log($"[TPL] Escaneando {collectionName}: Section={sectionIndex} Sheet={sheetIndex} Count={count}")
        For i As Integer = 1 To count
            Dim obj As Object = GetCollectionItem(coll, i)
            If obj Is Nothing Then Continue For

            Dim currentText As String = GetObjectTextValue(obj)
            If String.IsNullOrWhiteSpace(currentText) Then Continue For

            Dim replacement As String = ResolveSheetTextReplacement(currentText, config)
            If String.IsNullOrWhiteSpace(replacement) Then Continue For
            If String.Equals(currentText.Trim(), replacement.Trim(), StringComparison.Ordinal) Then Continue For

            If SetTextBoxValue(obj, replacement) Then
                updates += 1
                If logger IsNot Nothing Then logger.Log($"[TPL] {collectionName}#{i}: '{currentText}' -> '{replacement}'")
            End If
        Next
        Return updates
    End Function

    Private Shared Function GetObjectTextValue(obj As Object) As String
        If obj Is Nothing Then Return ""
        Dim candidates As String() = {"Text", "Value", "PrimaryCaption", "Title", "Name"}
        For Each m In candidates
            Dim s As String = GetStringMember(obj, m)
            If Not String.IsNullOrWhiteSpace(s) Then Return s
        Next
        Return ""
    End Function

    Private Shared Function ResolveSheetTextReplacement(currentText As String, config As JobConfiguration) As String
        If String.IsNullOrWhiteSpace(currentText) Then Return ""
        Dim raw As String = currentText.Trim()
        Dim upper As String = raw.ToUpperInvariant()

        If upper.Contains("|R") OrElse upper.Contains("%{") Then Return ""

        If upper.Contains("CLIENTE") Then
            If upper.Contains(":") Then
                Return Regex.Replace(raw, "(?i)(CLIENTE\s*:?\s*).*", "$1" & config.ClientName)
            End If
            Return config.ClientName
        End If

        If upper.Contains("PROYECTO") Then
            If upper.Contains(":") Then
                Return Regex.Replace(raw, "(?i)(PROYECTO\s*:?\s*).*", "$1" & config.ProjectName)
            End If
            Return config.ProjectName
        End If

        If upper.Contains("TITULO") OrElse upper.Contains("TÍTULO") Then
            If upper.Contains(":") Then
                Return Regex.Replace(raw, "(?i)(T[IÍ]TULO\s*:?\s*).*", "$1" & config.DrawingTitle)
            End If
            Return config.DrawingTitle
        End If

        If upper.Contains("MATERIAL") Then
            If upper.Contains(":") Then
                Return Regex.Replace(raw, "(?i)(MATERIAL\s*:?\s*).*", "$1" & config.Material)
            End If
            Return config.Material
        End If

        Return ""
    End Function

    Private Shared Function ResolveTitleBlockValue(textBoxObj As Object, config As JobConfiguration) As String
        If textBoxObj Is Nothing Then Return ""
        Dim candidates As New List(Of String)
        candidates.Add(GetStringMember(textBoxObj, "Name"))
        candidates.Add(GetStringMember(textBoxObj, "Title"))
        candidates.Add(GetStringMember(textBoxObj, "PromptText"))
        candidates.Add(GetStringMember(textBoxObj, "PrimaryCaption"))
        candidates.Add(GetStringMember(textBoxObj, "SecondaryCaption"))
        candidates.Add(GetStringMember(textBoxObj, "Text"))
        candidates.Add(GetStringMember(textBoxObj, "Value"))
        candidates.Add(GetStringMember(textBoxObj, "Formula"))

        For Each raw In candidates
            If String.IsNullOrWhiteSpace(raw) Then Continue For
            Dim k As String = raw.ToLowerInvariant()
            If k.Contains("cliente") OrElse k.Contains("client") Then Return config.ClientName
            If k.Contains("proyecto") OrElse k.Contains("project") Then Return config.ProjectName
            If k.Contains("titulo") OrElse k.Contains("title") Then Return config.DrawingTitle
            If k.Contains("material") Then Return config.Material
            If k.Contains("peso") OrElse k.Contains("weight") OrElse k.Contains("mass") Then Return config.Weight
            If k.Contains("equipo") OrElse k.Contains("equipment") Then Return config.Equipment
            If k.Contains("revision") OrElse k = "rev" Then Return config.Revision
            If k.Contains("plano") OrElse k.Contains("drawing number") OrElse k.Contains("part number") OrElse k.Contains("numero") Then Return config.DrawingNumber
            If k.Contains("observ") OrElse k.Contains("notes") OrElse k.Contains("coment") Then Return config.Notes
        Next
        Return ""
    End Function

    Private Shared Function SetTextBoxValue(textBoxObj As Object, value As String) As Boolean
        If textBoxObj Is Nothing OrElse value Is Nothing Then Return False
        Try
            CallByName(textBoxObj, "Value", CallType.Set, value)
            Return True
        Catch
        End Try
        Try
            CallByName(textBoxObj, "Text", CallType.Set, value)
            Return True
        Catch
        End Try
        Try
            CallByName(textBoxObj, "PrimaryCaption", CallType.Set, value)
            Return True
        Catch
        End Try
        Try
            CallByName(textBoxObj, "PrimaryText", CallType.Set, value)
            Return True
        Catch
        End Try
        Return False
    End Function

    Private Shared Function GetStringMember(obj As Object, memberName As String) As String
        If obj Is Nothing Then Return ""
        Try
            Dim v As Object = CallByName(obj, memberName, CallType.Get)
            If v Is Nothing Then Return ""
            Return v.ToString().Trim()
        Catch
            Return ""
        End Try
    End Function

    Private Shared Function FindNamedItem(collection As Object, candidateNames As IEnumerable(Of String)) As Object
        If collection Is Nothing Then Return Nothing
        Dim wanted As New List(Of String)
        For Each n In candidateNames
            If Not String.IsNullOrWhiteSpace(n) Then wanted.Add(n.Trim())
        Next
        If wanted.Count = 0 Then Return Nothing

        Try
            Dim count As Integer = CInt(CallByName(collection, "Count", CallType.Get))
            For i As Integer = 1 To count
                Dim item As Object = GetCollectionItem(collection, i)
                If item Is Nothing Then Continue For
                Dim itemName As String = ""
                Try : itemName = CStr(CallByName(item, "Name", CallType.Get)) : Catch : End Try
                If String.IsNullOrWhiteSpace(itemName) Then Continue For
                For Each w In wanted
                    If itemName.Equals(w, StringComparison.OrdinalIgnoreCase) Then
                        Return item
                    End If
                Next
            Next
        Catch
        End Try

        Return Nothing
    End Function

    Private Shared Function GetCollectionItem(collection As Object, key As Object) As Object
        If collection Is Nothing Then Return Nothing
        Try
            Return CallByName(collection, "Item", CallType.Get, key)
        Catch
            Try
                Return CallByName(collection, "Item", CallType.Method, key)
            Catch
                Return Nothing
            End Try
        End Try
    End Function

    Private Shared Function FirstNonEmpty(ParamArray values() As String) As String
        If values Is Nothing Then Return ""
        For Each value In values
            If Not String.IsNullOrWhiteSpace(value) Then Return value.Trim()
        Next
        Return ""
    End Function

    Friend Shared Function GetDocumentPropertyForMetadata(doc As Object, setName As String, propNames As IEnumerable(Of String)) As String
        Return GetDocumentProperty(doc, setName, propNames)
    End Function

    Friend Shared Function TrySetCustomProperty(doc As Object, propName As String, value As String, logger As Logger) As Boolean
        If doc Is Nothing OrElse String.IsNullOrWhiteSpace(propName) Then Return False
        If value Is Nothing Then value = ""
        If value.Trim() = "" Then Return False
        Return SetDocumentProperty(doc, "Custom", propName, value.Trim(), logger) > 0
    End Function

    Friend Shared Function TryFindMaterialByPropertyScan(doc As Object) As String
        If doc Is Nothing Then Return ""
        Try
            Dim sets As Object = Nothing
            Try : sets = CallByName(doc, "Properties", CallType.Get) : Catch : End Try
            If sets Is Nothing Then Try : sets = CallByName(doc, "PropertySets", CallType.Get) : Catch : End Try
            If sets Is Nothing Then Return ""
            Dim setCount As Integer = 0
            Try : setCount = CInt(CallByName(sets, "Count", CallType.Get)) : Catch : End Try
            For si As Integer = 1 To setCount
                Dim setObj As Object = GetCollectionItem(sets, si)
                If setObj Is Nothing Then Continue For
                Dim pc As Integer = 0
                Try : pc = CInt(CallByName(setObj, "Count", CallType.Get)) : Catch : End Try
                For pi As Integer = 1 To pc
                    Dim propObj As Object = GetCollectionItem(setObj, pi)
                    If propObj Is Nothing Then Continue For
                    Dim pn As String = ""
                    Try : pn = CStr(CallByName(propObj, "Name", CallType.Get)) : Catch : End Try
                    If String.IsNullOrWhiteSpace(pn) Then Continue For
                    If pn.IndexOf("material", StringComparison.OrdinalIgnoreCase) < 0 Then Continue For
                    Dim v As Object = Nothing
                    Try : v = CallByName(propObj, "Value", CallType.Get) : Catch : End Try
                    If v Is Nothing Then Continue For
                    Dim s As String = Convert.ToString(v).Trim()
                    If s <> "" Then Return s
                Next
            Next
        Catch
        End Try
        Return ""
    End Function

    Friend Shared Sub TrySamplePartsListCells(listsObj As Object, logger As Logger)
        If listsObj Is Nothing OrElse logger Is Nothing Then Return
        Try
            Dim nLists As Integer = 0
            Try : nLists = CInt(CallByName(listsObj, "Count", CallType.Get)) : Catch : End Try
            If nLists <= 0 Then Return
            Dim pl As Object = GetCollectionItem(listsObj, 1)
            If pl Is Nothing Then Return
            Dim nRows As Integer = 0, nCols As Integer = 0
            Try : nRows = CInt(CallByName(pl, "Rows", CallType.Get)) : Catch : End Try
            Try : nCols = CInt(CallByName(pl, "Columns", CallType.Get)) : Catch : End Try
            Dim maxR = Math.Min(Math.Max(nRows, 0), 3)
            Dim maxC = Math.Min(Math.Max(nCols, 0), 8)
            If maxR <= 0 OrElse maxC <= 0 Then Return
            Dim parts As New List(Of String)()
            For r As Integer = 1 To maxR
                For c As Integer = 1 To maxC
                    Dim t = TryGetPartsListCellTextInternal(pl, r, c)
                    If Not String.IsNullOrWhiteSpace(t) Then
                        parts.Add("r" & r.ToString() & "c" & c.ToString() & "=" & t)
                    End If
                Next
            Next
            If parts.Count > 0 Then
                logger.Log("[PARTSLIST][CELL_CHECK] " & String.Join("; ", parts.Take(16)))
            Else
                logger.Log("[PARTSLIST][CELL_CHECK][WARN] muestra vacía")
            End If
        Catch ex As Exception
            logger.Log("[PARTSLIST][CELL_CHECK][ERR] " & ex.Message)
        End Try
    End Sub

    Private Shared Function TryGetPartsListCellTextInternal(pl As Object, row1 As Integer, col1 As Integer) As String
        If pl Is Nothing Then Return ""
        Dim cellObj As Object = Nothing
        Try
            cellObj = CallByName(pl, "Cell", CallType.Get, row1, col1)
        Catch
            cellObj = Nothing
        End Try
        If cellObj Is Nothing Then
            Try
                cellObj = CallByName(pl, "Cell", CallType.Method, row1, col1)
            Catch
                cellObj = Nothing
            End Try
        End If
        If cellObj Is Nothing Then Return ""
        Try
            Dim s As String = Convert.ToString(CallByName(cellObj, "Text", CallType.Get))
            If String.IsNullOrWhiteSpace(s) Then s = Convert.ToString(CallByName(cellObj, "Value", CallType.Get))
            If Not String.IsNullOrWhiteSpace(s) Then
                s = s.Replace(vbCr, " ").Replace(vbLf, " ").Trim()
                If s.Length > 48 Then s = s.Substring(0, 48) & "…"
                Return s
            End If
        Catch
        End Try
        Return ""
    End Function

    Friend Shared Sub RefreshNativePartsListsAndUpdateAll(dftDoc As Object, logger As Logger)
        If dftDoc Is Nothing Then Return
        Try
            Dim lists As Object = Nothing
            Try : lists = CallByName(dftDoc, "PartsLists", CallType.Get) : Catch : End Try
            If lists IsNot Nothing Then
                Dim n As Integer = 0
                Try : n = CInt(CallByName(lists, "Count", CallType.Get)) : Catch : End Try
                For i As Integer = 1 To n
                    Dim pl As Object = GetCollectionItem(lists, i)
                    If pl Is Nothing Then Continue For
                    Try
                        CallByName(pl, "Update", CallType.Method)
                        If logger IsNot Nothing Then logger.Log("[PARTSLIST][UPDATE][OK] post_props index=" & i.ToString())
                    Catch ex As Exception
                        If logger IsNot Nothing Then logger.Log("[PARTSLIST][UPDATE][WARN] " & ex.Message)
                    End Try
                Next
                TrySamplePartsListCells(lists, logger)
            End If
        Catch ex As Exception
            If logger IsNot Nothing Then logger.Log("[PARTSLIST][REFRESH][ERR] " & ex.Message)
        End Try
        Try
            CallByName(dftDoc, "UpdateAll", CallType.Method, True)
        Catch ex As Exception
            If logger IsNot Nothing Then logger.Log("[PARTSLIST][UPDATEALL][WARN] " & ex.Message)
        End Try
    End Sub
End Class
