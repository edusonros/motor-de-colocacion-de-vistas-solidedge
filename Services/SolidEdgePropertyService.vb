Option Strict Off

Imports System
Imports System.Collections.Generic
Imports System.Globalization
Imports System.Linq
Imports System.Text
Imports System.Runtime.InteropServices
Imports System.Reflection
Imports System.Text.RegularExpressions
Imports SolidEdgeDraft
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

    ''' <summary>Valor de campo «Plano/Nº documento» desde UI suele llegar como nombre completo (.psm + denominación); ProjectInformation debe llevar sólo código.</summary>
    Private Shared Function NormalizeProjectDocumentNumberForProperty(raw As String) As String
        If String.IsNullOrWhiteSpace(raw) Then Return ""
        Dim s As String = raw.Trim()
        Dim ext As String = ""
        Try
            ext = IO.Path.GetExtension(s)
        Catch
            ext = ""
        End Try
        If Not String.IsNullOrWhiteSpace(ext) Then
            Dim el = ext.ToLowerInvariant()
            If el = ".psm" OrElse el = ".par" OrElse el = ".asm" OrElse el = ".dft" Then
                Try
                    s = IO.Path.GetFileNameWithoutExtension(s)
                Catch
                End Try
            End If
        End If
        Dim sp As Integer = s.IndexOfAny(New Char() {" "c, ChrW(9)})
        If sp > 0 Then s = s.Substring(0, sp).TrimEnd()
        Return s.Trim()
    End Function

    ''' <summary>Save() puntual suele dar E_FAIL con SE oculto o documento sin activar; reintentos + SaveAs al mismo FullName como respaldo COM.</summary>
    Private Shared Function TryPersistDraftComSave(draftDoc As Object, logger As Logger) As Boolean
        Const maxAttempts As Integer = 4
        For attempt As Integer = 1 To maxAttempts
            Try
                CallByName(draftDoc, "Activate", CallType.Method)
            Catch
            End Try
            Try
                CallByName(draftDoc, "UpdateAll", CallType.Method, False)
            Catch
            End Try
            Try
                CallByName(draftDoc, "Save", CallType.Method)
                If logger IsNot Nothing Then
                    logger.Log("[DFT][SUMMARYINFO] Save=True intento=" & attempt.ToString(CultureInfo.InvariantCulture))
                End If
                Return True
            Catch ex As Exception
                If logger IsNot Nothing Then logger.Log("[DFT][SUMMARYINFO][SAVE_RETRY] #" & attempt.ToString(CultureInfo.InvariantCulture) & " Save: " & ex.Message)
            End Try
            Dim fp As String = ""
            Try : fp = Convert.ToString(CallByName(draftDoc, "FullName", CallType.Get)) : Catch : End Try
            If Not String.IsNullOrWhiteSpace(fp) Then
                Try
                    CallByName(draftDoc, "SaveAs", CallType.Method, fp)
                    If logger IsNot Nothing Then logger.Log("[DFT][SUMMARYINFO] SaveAs(mismo FullName) OK")
                    Return True
                Catch ex2 As Exception
                    If logger IsNot Nothing Then logger.Log("[DFT][SUMMARYINFO][SAVE_RETRY] SaveAs: " & ex2.Message)
                End Try
            End If
            Try
                Threading.Thread.Sleep(120 * attempt)
            Catch
            End Try
        Next
        Return False
    End Function

    Public Shared Function ApplyDirectSummaryInfoToDraft(draftDoc As Object,
                                                          config As JobConfiguration,
                                                          logger As Logger) As Boolean
        If draftDoc Is Nothing OrElse config Is Nothing Then Return False
        Dim denom As String = If(config.DrawingTitle, "").Trim()
        Dim drawingNum As String = If(config.DrawingNumber, "").Trim()
        Dim pedidoVal As String = If(config.Pedido, "").Trim()
        Return ApplyDirectSummaryInfoToDraft(draftDoc, denom, drawingNum, pedidoVal, config.ClientName, config.ProjectName, logger)
    End Function

    ''' <summary>Relleno «ligero» de SummaryInfo desde la UI sin pasar por el perfil completo.
    ''' Cajetín CADEBRO: el campo visible «PEDIDO» suele enlazar a Summary Asunto (= Subject COM), no a Custom.Pedido;
    ''' la denominación va en Título (= Title).</summary>
    Public Shared Function ApplyDirectSummaryInfoToDraft(draftDoc As Object,
                                                          denomination As String,
                                                          drawingNumber As String,
                                                          pedido As String,
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

        ' Subject (Asunto): sólo pedido; no mezclar con denominación (muchas plantillas enlazan «PEDIDO» al Asunto).
        If Not String.IsNullOrWhiteSpace(pedido) Then
            SetSummaryInfoField(summaryInfo, "Subject", pedido.Trim(), logger)
        Else
            If logger IsNot Nothing Then logger.Log("[DFT][SUMMARYINFO][SKIP] Subject: Pedido vacío en UI (no se escribe denominación aquí).")
        End If

        Dim titleEffective As String = If(Not String.IsNullOrWhiteSpace(denomination), denomination.Trim(), drawingNumber.Trim())
        If Not String.IsNullOrWhiteSpace(titleEffective) Then
            SetSummaryInfoField(summaryInfo, "Title", titleEffective, logger)
        Else
            If logger IsNot Nothing Then logger.Log("[DFT][SUMMARYINFO][SKIP] Title: valor vacío (no se fuerza borrado).")
        End If

        If Not String.IsNullOrWhiteSpace(drawingNumber) Then
            Dim docNum As String = NormalizeProjectDocumentNumberForProperty(drawingNumber)
            If String.IsNullOrWhiteSpace(docNum) Then docNum = drawingNumber.Trim()
            If logger IsNot Nothing AndAlso Not String.Equals(docNum, drawingNumber.Trim(), StringComparison.Ordinal) Then
                logger.Log("[DFT][SUMMARYINFO] Document Number normalizado: '" & drawingNumber.Trim() & "' -> '" & docNum & "'")
            End If
            SetDocumentProperty(draftDoc, "ProjectInformation", "Document Number", docNum, logger)
            SetDocumentProperty(draftDoc, "ProjectInformation", "Número de documento", docNum, logger)
        End If

        If Not String.IsNullOrWhiteSpace(companyValue) Then
            SetSummaryInfoField(summaryInfo, "Company", companyValue, logger)
        End If
        If Not String.IsNullOrWhiteSpace(projectNameValue) Then
            SetSummaryInfoField(summaryInfo, "ProjectName", projectNameValue, logger)
        End If

        Dim saved As Boolean = TryPersistDraftComSave(draftDoc, logger)
        If logger IsNot Nothing Then logger.Log("[DFT][SUMMARYINFO] Save=" & saved.ToString())
        Return saved
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
        Dim persisted As Boolean = False
        Try
            OleMessageFilter.Register()
            If Not TryConnectApplication(showSolidEdge, logger, app, createdByUs) Then Return False

            dftDoc = app.Documents.Open(dftPath)
            persisted = ApplyDirectSummaryInfoToDraft(dftDoc, config, logger)
            Return persisted
        Catch ex As Exception
            If logger IsNot Nothing Then logger.LogException("ApplyDirectSummaryInfoToDraftFile", ex)
            Return False
        Finally
            ' Si Save falló, Close(True) a veces consigue persistir; si ya guardó, Close(False) evita doble escritura.
            TryCloseComDocument(dftDoc, saveChanges:=Not persisted)
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

    ''' <summary>Primera ruta de pieza/conjunto (.par/.psm/.asm) encontrada por PartsList o ModelLinks; vacío si no hay archivo accesible.</summary>
    Public Shared Function TryGetPrimaryLinkedModelFullPath(draftDoc As Object) As String
        If draftDoc Is Nothing Then Return ""
        Try
            Dim lists As Object = Nothing
            Try : lists = CallByName(draftDoc, "PartsLists", CallType.Get) : Catch : End Try
            If lists IsNot Nothing Then
                Dim nLists As Integer = 0
                Try : nLists = CInt(CallByName(lists, "Count", CallType.Get)) : Catch : End Try
                If nLists >= 1 Then
                    Dim pl As Object = GetCollectionItem(lists, 1)
                    If pl IsNot Nothing Then
                        Dim asm As String = ""
                        Try : asm = CStr(CallByName(pl, "AssemblyFileName", CallType.Get)) : Catch : End Try
                        asm = If(asm, "").Trim()
                        If asm.Length > 0 Then Return asm
                    End If
                End If
            End If
        Catch
        End Try
        Try
            Dim modelLinks As Object = Nothing
            Try : modelLinks = CallByName(draftDoc, "ModelLinks", CallType.Get) : Catch : End Try
            If modelLinks Is Nothing Then Return ""
            Dim n As Integer = 0
            Try : n = CInt(CallByName(modelLinks, "Count", CallType.Get)) : Catch : End Try
            If n < 1 Then Return ""
            Dim linkObj As Object = GetCollectionItem(modelLinks, 1)
            If linkObj Is Nothing Then Return ""
            Dim fn As String = ""
            Try : fn = CStr(CallByName(linkObj, "FileName", CallType.Get)) : Catch : End Try
            fn = If(fn, "").Trim()
            If fn.Length > 0 Then Return fn
        Catch
        End Try
        Return ""
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
            TryCloseComDocument(dftDoc, False)
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
            If openedByUs Then TryCloseComDocument(doc, False)
            SolidEdgeSessionVisibility.SuppressForegroundIfConfigured(app, config, logger)
        End Try
        Return written
    End Function

    ''' <summary>Escribe el perfil estándar (cajetín + Custom usados por la aplicación) en un modelo PAR/PSM ya abierto; opcionalmente guarda.</summary>
    Public Shared Function ApplyPropertiesToExistingOpenModelDocument(modelDoc As Object, config As JobConfiguration, logger As Logger, Optional saveAfter As Boolean = True) As Integer
        If modelDoc Is Nothing OrElse config Is Nothing Then Return 0
        Dim profile = BuildPropertySyncProfile(config)
        Dim written = ApplyProfileToOpenModelDocument(modelDoc, profile, logger)
        Dim saved As Boolean = False
        If saveAfter AndAlso written > 0 Then
            Try
                CallByName(modelDoc, "Save", CallType.Method)
                saved = True
            Catch ex As Exception
                If logger IsNot Nothing Then logger.LogException("ApplyPropertiesToExistingOpenModelDocument.Save", ex)
            End Try
            If logger IsNot Nothing Then logger.Log($"[PROPS][EXISTING_MODEL] escritas={written} save={saved}")
        End If
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
            AppendOpenDocumentPropertySets(doc, Sub(line) logger.Log(line), maxPropsPerSet)
        Catch ex As Exception
            logger.LogException("LogAllPropertySetsFromOpenDocument", ex)
        End Try
    End Sub

    ''' <summary>Volcado de PropertySets del documento abierto (p. ej. auditoría DFT en archivo .txt).</summary>
    Public Shared Sub AppendPropertySetsAuditToStringBuilder(doc As Object, sb As StringBuilder, Optional maxPropsPerSet As Integer = 250)
        If doc Is Nothing OrElse sb Is Nothing Then Return
        Try
            sb.AppendLine("")
            sb.AppendLine("[PROPERTIES] PropertySets del documento DFT (valores actuales; el cajetín puede enlazar cualquiera de estos nombres).")
            AppendOpenDocumentPropertySets(doc, Sub(line) sb.AppendLine(line), maxPropsPerSet)
        Catch ex As Exception
            sb.AppendLine("[PROPERTIES][ERR] " & ex.Message)
        End Try
    End Sub

    Private Shared Sub AppendOpenDocumentPropertySets(doc As Object, emit As Action(Of String), maxPropsPerSet As Integer)
        If doc Is Nothing OrElse emit Is Nothing Then Return
        Dim fullName As String = ""
        Dim docType As String = ""
        Try : fullName = CStr(CallByName(doc, "FullName", CallType.Get)) : Catch : End Try
        Try : docType = TypeName(doc) : Catch : End Try
        emit($"[PROPS][DOC][INSPECT] Archivo={fullName}")
        emit($"[PROPS][DOC][INSPECT] Tipo={docType}")

        Dim sets As Object = Nothing
        Try : sets = CallByName(doc, "Properties", CallType.Get) : Catch : End Try
        If sets Is Nothing Then
            Try : sets = CallByName(doc, "PropertySets", CallType.Get) : Catch : End Try
        End If
        If sets Is Nothing Then
            emit("[PROPS][DOC][INSPECT][WARN] Documento sin colección Properties/PropertySets accesible.")
            Return
        End If

        Dim setCount As Integer = 0
        Try : setCount = CInt(CallByName(sets, "Count", CallType.Get)) : Catch : End Try
        emit($"[PROPS][DOC][INSPECT] PropertySets detectados={setCount}")

        For si As Integer = 1 To setCount
            Dim setObj As Object = GetCollectionItem(sets, si)
            If setObj Is Nothing Then Continue For
            Dim setName As String = $"Set#{si}"
            Try : setName = CStr(CallByName(setObj, "Name", CallType.Get)) : Catch : End Try
            Dim propCount As Integer = 0
            Try : propCount = CInt(CallByName(setObj, "Count", CallType.Get)) : Catch : End Try
            emit($"[PROPS][DOC][INSPECT] Set={setName} Count={propCount}")

            Dim printed As Integer = 0
            For pi As Integer = 1 To propCount
                If printed >= Math.Max(1, maxPropsPerSet) Then
                    emit($"[PROPS][DOC][INSPECT] Set={setName} ... truncado en {maxPropsPerSet} propiedades.")
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
                emit($"[PROPS][DOC][INSPECT] {setName}.{propName}='{propValue}'")
                printed += 1
            Next
        Next
    End Sub

    Public Shared Sub LogDraftTitleBlockDataSources(dftDoc As Object, logger As Logger)
        If dftDoc Is Nothing OrElse logger Is Nothing Then Return
        Try
            WalkDraftTitleBlockTextSources(dftDoc, Sub(line) logger.Log(line))
        Catch ex As Exception
            logger.LogException("LogDraftTitleBlockDataSources", ex)
        End Try
    End Sub

    ''' <summary>Fuentes PropertyText del cajetín y textos sueltos (misma lógica que el log de plantilla).</summary>
    Public Shared Sub AppendTitleBlockSourcesAuditToStringBuilder(dftDoc As Object, sb As StringBuilder)
        If dftDoc Is Nothing OrElse sb Is Nothing Then Return
        Try
            sb.AppendLine("")
            sb.AppendLine("[TITLEBLOCK_TEXT] Textos con fórmula / PropertyText (enlazan Custom, Summary, |R model link, etc.).")
            WalkDraftTitleBlockTextSources(dftDoc, Sub(line) sb.AppendLine(line))
        Catch ex As Exception
            sb.AppendLine("[TITLEBLOCK_TEXT][ERR] " & ex.Message)
        End Try
    End Sub

    Private Shared Sub WalkDraftTitleBlockTextSources(dftDoc As Object, emit As Action(Of String))
        If dftDoc Is Nothing OrElse emit Is Nothing Then Return
        Dim sections As Object = Nothing
        Try : sections = CallByName(dftDoc, "Sections", CallType.Get) : Catch : End Try
        If sections Is Nothing Then
            emit("[TPL][SRC][WARN] No se pudo acceder a Sections del draft.")
            Return
        End If

        Dim secCount As Integer = 0
        Try : secCount = CInt(CallByName(sections, "Count", CallType.Get)) : Catch : End Try
        emit($"[TPL][SRC] Secciones={secCount}")
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
                InspectTitleBlockCollectionSource(sheet, "TitleBlocks", emit, si, shi)
                InspectTextCollectionSource(sheet, "Texts", emit, si, shi)
                InspectTextCollectionSource(sheet, "TextBoxes", emit, si, shi)
            Next
        Next
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

        ' Plano (código dibujo): ProjectInformation + Custom; NO Summary.Título (en español suele llevar denominación pieza — ver entrada DenominacionTitulo).
        Dim plano = New PropertySyncEntry With {.LogicalName = "Plano", .Value = config.DrawingNumber, .Target = PropertySyncTarget.Both}
        plano.CustomBindings.Add(New PropertyBinding With {.SetName = "Custom", .PropertyName = "Plano", .AllowCreate = True})
        ' Alias cortos habituales en plantillas locales (campo texto cajetín).
        plano.CustomBindings.Add(New PropertyBinding With {.SetName = "Custom", .PropertyName = "plan", .AllowCreate = True})
        list.Add(plano)

        ' Denominación / descripción corta → Subject/Asunto + Summary.Título (Solid Edge español — ver audits CADEBRO).
        ' Denominación → Título Summary; NO escribir Subject (Asunto): en CADEBRO «PEDIDO» enlaza ahí y lleva Custom.Pedido / número.
        Dim denom = New PropertySyncEntry With {.LogicalName = "DenominacionTitulo", .Value = config.DrawingTitle, .Target = PropertySyncTarget.Both}
        denom.StandardBindings.Add(New PropertyBinding With {.SetName = "SummaryInformation", .PropertyName = "Title", .AllowCreate = False})
        denom.StandardBindings.Add(New PropertyBinding With {.SetName = "SummaryInformation", .PropertyName = "Document Title", .AllowCreate = False})
        denom.StandardBindings.Add(New PropertyBinding With {.SetName = "SummaryInformation", .PropertyName = "Título", .AllowCreate = False})
        denom.StandardBindings.Add(New PropertyBinding With {.SetName = "SummaryInformation", .PropertyName = "Titulo", .AllowCreate = False})
        denom.CustomBindings.Add(New PropertyBinding With {.SetName = "Custom", .PropertyName = "Denominacion", .AllowCreate = True})
        denom.CustomBindings.Add(New PropertyBinding With {.SetName = "Custom", .PropertyName = "Titulo", .AllowCreate = True})
        denom.CustomBindings.Add(New PropertyBinding With {.SetName = "Custom", .PropertyName = "nom", .AllowCreate = True})
        list.Add(denom)

        Dim project = New PropertySyncEntry With {.LogicalName = "NombreProyecto", .Value = config.ProjectName, .Target = PropertySyncTarget.Both}
        project.StandardBindings.Add(New PropertyBinding With {.SetName = "ProjectInformation", .PropertyName = "Project Name", .AllowCreate = False})
        project.StandardBindings.Add(New PropertyBinding With {.SetName = "ProjectInformation", .PropertyName = "Nombre de proyecto", .AllowCreate = False})
        project.CustomBindings.Add(New PropertyBinding With {.SetName = "Custom", .PropertyName = "Proyecto", .AllowCreate = True})
        project.CustomBindings.Add(New PropertyBinding With {.SetName = "Custom", .PropertyName = "proy", .AllowCreate = True})
        list.Add(project)

        Dim material = New PropertySyncEntry With {.LogicalName = "Material", .Value = config.Material, .Target = PropertySyncTarget.Both}
        material.StandardBindings.Add(New PropertyBinding With {.SetName = "MechanicalModeling", .PropertyName = "Material", .AllowCreate = False})
        material.StandardBindings.Add(New PropertyBinding With {.SetName = "ProjectInformation", .PropertyName = "Material", .AllowCreate = False})
        material.CustomBindings.Add(New PropertyBinding With {.SetName = "Custom", .PropertyName = "Material", .AllowCreate = True})
        material.CustomBindings.Add(New PropertyBinding With {.SetName = "Custom", .PropertyName = "mat", .AllowCreate = True})
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
        drawingNum.CustomBindings.Add(New PropertyBinding With {.SetName = "Custom", .PropertyName = "CODIGO", .AllowCreate = True})
        list.Add(drawingNum)

        Dim revision = New PropertySyncEntry With {.LogicalName = "Revision", .Value = config.Revision, .Target = PropertySyncTarget.Both}
        revision.StandardBindings.Add(New PropertyBinding With {.SetName = "SummaryInformation", .PropertyName = "Número de revisión", .AllowCreate = False})
        revision.StandardBindings.Add(New PropertyBinding With {.SetName = "SummaryInformation", .PropertyName = "Revision Number", .AllowCreate = False})
        revision.StandardBindings.Add(New PropertyBinding With {.SetName = "ProjectInformation", .PropertyName = "Revision", .AllowCreate = False})
        revision.StandardBindings.Add(New PropertyBinding With {.SetName = "ProjectInformation", .PropertyName = "Revisión", .AllowCreate = False})
        revision.CustomBindings.Add(New PropertyBinding With {.SetName = "Custom", .PropertyName = "Revision", .AllowCreate = True})
        list.Add(revision)

        Dim author = New PropertySyncEntry With {.LogicalName = "Autor", .Value = TitleBlockFieldCoordinator.ResolveEffectiveAuthor(config), .Target = PropertySyncTarget.Both}
        author.StandardBindings.Add(New PropertyBinding With {.SetName = "SummaryInformation", .PropertyName = "Author", .AllowCreate = False})
        author.StandardBindings.Add(New PropertyBinding With {.SetName = "SummaryInformation", .PropertyName = "Autor", .AllowCreate = False})
        author.CustomBindings.Add(New PropertyBinding With {.SetName = "Custom", .PropertyName = "Autor", .AllowCreate = True})
        author.CustomBindings.Add(New PropertyBinding With {.SetName = "Custom", .PropertyName = "autor", .AllowCreate = True})
        author.CustomBindings.Add(New PropertyBinding With {.SetName = "Custom", .PropertyName = "aut", .AllowCreate = True})
        list.Add(author)

        Dim thickness = New PropertySyncEntry With {.LogicalName = "Espesor", .Value = If(config.Thickness, "").Trim(), .Target = PropertySyncTarget.Both}
        thickness.StandardBindings.Add(New PropertyBinding With {.SetName = "MechanicalModeling", .PropertyName = "Sheet Metal Gauge", .AllowCreate = False})
        thickness.CustomBindings.Add(New PropertyBinding With {.SetName = "Custom", .PropertyName = "Espesor", .AllowCreate = True})
        list.Add(thickness)

        Dim order = New PropertySyncEntry With {.LogicalName = "Pedido", .Value = If(config.Pedido, "").Trim(), .Target = PropertySyncTarget.Both}
        order.StandardBindings.Add(New PropertyBinding With {.SetName = "SummaryInformation", .PropertyName = "Subject", .AllowCreate = False})
        order.StandardBindings.Add(New PropertyBinding With {.SetName = "SummaryInformation", .PropertyName = "Asunto", .AllowCreate = False})
        ' Variantes habituales enlazadas en cajetines (incl. PEDIDO / Order). El perfil intenta escribir todas si aplica espejo.
        order.CustomBindings.Add(New PropertyBinding With {.SetName = "Custom", .PropertyName = "Pedido", .AllowCreate = True})
        order.CustomBindings.Add(New PropertyBinding With {.SetName = "Custom", .PropertyName = "PEDIDO", .AllowCreate = True})
        order.CustomBindings.Add(New PropertyBinding With {.SetName = "Custom", .PropertyName = "Order", .AllowCreate = True})
        list.Add(order)

        Dim fechaVal = If(config.FechaPlano, "").Trim()
        If fechaVal <> "" Then
            Dim fecha = New PropertySyncEntry With {.LogicalName = "FechaPlano", .Value = fechaVal, .Target = PropertySyncTarget.Both}
            fecha.CustomBindings.Add(New PropertyBinding With {.SetName = "Custom", .PropertyName = "FechaPlano", .AllowCreate = True})
            fecha.CustomBindings.Add(New PropertyBinding With {.SetName = "Custom", .PropertyName = "fec", .AllowCreate = True})
            list.Add(fecha)
        End If

        ' Escala de dibujo (texto 1:n) en Custom; muchos cajetines enlazan Escala/esc al documento.
        If Not config.UseAutomaticScale AndAlso config.ManualScale > 1.0E-9 Then
            Dim escStr = FormatScaleOneToN(config.ManualScale)
            If Not String.IsNullOrWhiteSpace(escStr) Then
                Dim escEntry = New PropertySyncEntry With {.LogicalName = "EscalaDibujo", .Value = escStr.Trim(), .Target = PropertySyncTarget.Both}
                escEntry.CustomBindings.Add(New PropertyBinding With {.SetName = "Custom", .PropertyName = "Escala", .AllowCreate = True})
                escEntry.CustomBindings.Add(New PropertyBinding With {.SetName = "Custom", .PropertyName = "esc", .AllowCreate = True})
                list.Add(escEntry)
            End If
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

    ''' <summary>Convierte factor de escala del motor (modelo→hoja, p. ej. 0,4) a cadena «1:2,5» para propiedades de texto.</summary>
    Private Shared Function FormatScaleOneToN(scaleFactor As Double) As String
        If scaleFactor <= 1.0E-9 Then Return ""
        Dim inv = 1.0R / scaleFactor
        Dim s = inv.ToString("0.####", CultureInfo.InvariantCulture).TrimEnd("0"c).TrimEnd("."c)
        If String.IsNullOrEmpty(s) Then s = "1"
        Return "1:" & s
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

            Dim wroteAny As Boolean = False
            Dim lastStandardReason As String = ""
            For Each binding In entry.StandardBindings
                Dim status As PropertyWriteStatus = PropertyWriteStatus.UnexpectedError
                Dim reason As String = ""
                If TrySetDocumentPropertyDetailed(doc, binding.SetName, binding.PropertyName, entry.Value, binding.AllowCreate, status, reason) Then
                    updates += 1
                    wroteAny = True
                    If logger IsNot Nothing Then logger.Log($"[PROPS][WRITE] {entry.LogicalName} ← {binding.SetName}.{binding.PropertyName} OK")
                    Exit For
                Else
                    lastStandardReason = $"{binding.SetName}.{binding.PropertyName} -> {status}"
                End If
            Next

            ' Customs (aliases proy/plan/nom…) deben aplicarse aunque un estándar ya se haya escrito.
            Dim wroteCustomOk As Boolean = False
            For Each binding In entry.CustomBindings
                Dim status As PropertyWriteStatus = PropertyWriteStatus.UnexpectedError
                Dim reason As String = ""
                If TrySetDocumentPropertyDetailed(doc, binding.SetName, binding.PropertyName, entry.Value, True, status, reason) Then
                    updates += 1
                    wroteAny = True
                    wroteCustomOk = True
                    If logger IsNot Nothing Then logger.Log($"[PROPS][WRITE] {entry.LogicalName} ← Custom {binding.SetName}.{binding.PropertyName} OK")
                End If
            Next

            If Not wroteAny AndAlso logger IsNot Nothing Then
                logger.Log($"[PROPS][WRITE][WARN] {entry.LogicalName} sin escrituras útiles. Estándar: {lastStandardReason}")
            ElseIf wroteAny AndAlso Not wroteCustomOk AndAlso entry.CustomBindings.Count > 0 AndAlso logger IsNot Nothing Then
                logger.Log($"[PROPS][WRITE][WARN] {entry.LogicalName}: no se pudieron crear/actualizar Customs (Último estándar: {lastStandardReason})")
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
                                                         emit As Action(Of String),
                                                         sectionIndex As Integer,
                                                         sheetIndex As Integer)
        If sheet Is Nothing OrElse emit Is Nothing Then Return
        Dim coll As Object = Nothing
        Try : coll = CallByName(sheet, collectionName, CallType.Get) : Catch : End Try
        If coll Is Nothing Then Return
        Dim count As Integer = 0
        Try : count = CInt(CallByName(coll, "Count", CallType.Get)) : Catch : End Try
        For i As Integer = 1 To count
            Dim tb As Object = GetCollectionItem(coll, i)
            If tb Is Nothing Then Continue For
            Dim tbType As String = TypeName(tb)
            emit($"[TPL][SRC] Section={sectionIndex} Sheet={sheetIndex} {collectionName}#{i} Type={tbType}")
            Dim textBoxes As Object = Nothing
            Try : textBoxes = CallByName(tb, "TextBoxes", CallType.Get) : Catch : End Try
            If textBoxes Is Nothing Then Continue For
            Dim txCount As Integer = 0
            Try : txCount = CInt(CallByName(textBoxes, "Count", CallType.Get)) : Catch : End Try
            For txi As Integer = 1 To txCount
                Dim tx As Object = GetCollectionItem(textBoxes, txi)
                If tx Is Nothing Then Continue For
                LogTextObjectSource(tx, emit, $"Section={sectionIndex} Sheet={sheetIndex} {collectionName}#{i}.TextBox#{txi}")
            Next
        Next
    End Sub

    Private Shared Sub InspectTextCollectionSource(sheet As Object,
                                                   collectionName As String,
                                                   emit As Action(Of String),
                                                   sectionIndex As Integer,
                                                   sheetIndex As Integer)
        If sheet Is Nothing OrElse emit Is Nothing Then Return
        Dim coll As Object = Nothing
        Try : coll = CallByName(sheet, collectionName, CallType.Get) : Catch : End Try
        If coll Is Nothing Then Return
        Dim count As Integer = 0
        Try : count = CInt(CallByName(coll, "Count", CallType.Get)) : Catch : End Try
        For i As Integer = 1 To count
            Dim tx As Object = GetCollectionItem(coll, i)
            If tx Is Nothing Then Continue For
            LogTextObjectSource(tx, emit, $"Section={sectionIndex} Sheet={sheetIndex} {collectionName}#{i}")
        Next
    End Sub

    Private Shared Sub LogTextObjectSource(obj As Object, emit As Action(Of String), location As String)
        If obj Is Nothing OrElse emit Is Nothing Then Return
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
        emit($"[TPL][SRC] {location} Type={objType} Name='{objName}' Visible='{visible}' Formula='{formula}' SourceType={srcType} SourceGuess={srcGuess}")
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

    ''' <summary>Reutiliza el DFT ya abierto en Solid Edge; prioriza la instancia con <c>PartsLists.Count &gt; 0</c>.</summary>
    Friend Shared Function TryGetOrOpenDraftDocumentByPath(app As Application, filePath As String, logger As Logger, ByRef openedByUs As Boolean) As Object
        openedByUs = False
        If app Is Nothing OrElse String.IsNullOrWhiteSpace(filePath) OrElse Not IO.File.Exists(filePath) Then Return Nothing
        Dim path As String = TryNormalizeExistingFilePath(filePath)

        Dim active As Object = Nothing
        Try : active = CallByName(app, "ActiveDocument", CallType.Get) : Catch : End Try
        If active IsNot Nothing AndAlso LooksLikeDraftDocument(active) AndAlso DocumentPathsMatch(TryGetDocumentFullName(active), path) Then
            If logger IsNot Nothing Then logger.Log("[PARTLISTDATA][DFT_OPEN] Reutilizando DFT activo en Solid Edge: " & path)
            Return TryPreferDraftWithPartsLists(app, active, path, logger, openedByUs)
        End If

        Dim existing As Object = TryFindOpenDocumentByPath(app, path, logger, isDraft:=True, requirePartsList:=False)
        If existing IsNot Nothing Then
            If logger IsNot Nothing Then logger.Log("[PARTLISTDATA][DFT_OPEN] Reutilizando DFT ya abierto: " & path)
            Return TryPreferDraftWithPartsLists(app, existing, path, logger, openedByUs)
        End If

        Dim doc As Object = Nothing
        Try
            doc = app.Documents.Open(path)
            openedByUs = (doc IsNot Nothing)
            If logger IsNot Nothing AndAlso openedByUs Then logger.Log("[PARTLISTDATA][DFT_OPEN] " & path)
        Catch ex As Exception
            If logger IsNot Nothing Then logger.Log("[PARTLISTDATA][DFT_OPEN][ERR] " & ex.Message)
            Return Nothing
        End Try

        Return TryPreferDraftWithPartsLists(app, doc, path, logger, openedByUs)
    End Function

    ''' <summary>Si <paramref name="draftDoc"/> no expone PartsLists, cambia a otra ventana COM del mismo .dft que sí.</summary>
    Friend Shared Function TryEnsureDraftWithPartsLists(draftDoc As Object, logger As Logger) As Object
        If draftDoc Is Nothing Then Return Nothing
        If TryGetDraftPartsListCount(draftDoc) > 0 Then Return draftDoc
        Dim app As Application = TryGetApplicationFromDocument(draftDoc)
        If app Is Nothing Then Return draftDoc
        Dim path As String = TryGetDocumentFullName(draftDoc)
        If String.IsNullOrWhiteSpace(path) Then Return draftDoc
        Dim openedByUs As Boolean = False
        Return TryPreferDraftWithPartsLists(app, draftDoc, TryNormalizeExistingFilePath(path), logger, openedByUs)
    End Function

    Friend Shared Function TryGetOrOpenModelDocumentByPath(app As Application, filePath As String, logger As Logger, ByRef openedByUs As Boolean) As Object
        openedByUs = False
        If app Is Nothing OrElse String.IsNullOrWhiteSpace(filePath) OrElse Not IO.File.Exists(filePath) Then Return Nothing
        Dim path As String = TryNormalizeExistingFilePath(filePath)
        Dim doc As Object = TryFindOpenDocumentByPath(app, path, logger, isDraft:=False, requirePartsList:=False)
        If doc IsNot Nothing Then
            If logger IsNot Nothing Then logger.Log("[PARTLISTDATA][MODEL_OPEN] Reutilizando documento ya abierto: " & path)
            Return doc
        End If
        Try
            doc = app.Documents.Open(path)
            openedByUs = (doc IsNot Nothing)
            If logger IsNot Nothing AndAlso openedByUs Then logger.Log("[PARTLISTDATA][MODEL_OPEN] " & path)
        Catch ex As Exception
            If logger IsNot Nothing Then logger.Log("[PARTLISTDATA][MODEL_OPEN][ERR] " & ex.Message)
            doc = Nothing
        End Try
        Return doc
    End Function

    Private Shared Function TryPreferDraftWithPartsLists(app As Application, current As Object, normalizedPath As String, logger As Logger,
                                                        ByRef openedByUs As Boolean) As Object
        If current Is Nothing Then Return Nothing
        If TryGetDraftPartsListCount(current) > 0 Then Return current

        Dim withList As Object = TryFindOpenDocumentByPath(app, normalizedPath, logger, isDraft:=True, requirePartsList:=True)
        If withList Is Nothing OrElse Object.ReferenceEquals(withList, current) Then
            LogOpenDraftCandidates(app, normalizedPath, logger)
            Return current
        End If

        If logger IsNot Nothing Then
            logger.Log("[PARTLISTDATA][DFT_RESOLVE] Cambio a instancia con PartsLists (Count=" &
                       TryGetDraftPartsListCount(withList).ToString(CultureInfo.InvariantCulture) &
                       "); la copia abierta por API tenía Count=0.")
        End If
        If openedByUs AndAlso current IsNot Nothing Then
            TryCloseComDocument(current, False)
            openedByUs = False
        End If
        Return withList
    End Function

    Friend Shared Sub TryActivateDraftDocument(draftDoc As Object, logger As Logger)
        If draftDoc Is Nothing Then Return
        Try
            CallByName(draftDoc, "Activate", CallType.Method)
        Catch ex As Exception
            If logger IsNot Nothing Then logger.Log("[PARTLISTDATA][DFT_ACTIVATE][WARN] " & ex.Message)
        End Try
        Try
            CallByName(draftDoc, "UpdateAll", CallType.Method, False)
        Catch
        End Try
    End Sub

    Friend Shared Function TryGetDraftPartsLists(draftDoc As Object, logger As Logger) As Object
        If draftDoc Is Nothing Then Return Nothing
        draftDoc = TryEnsureDraftWithPartsLists(draftDoc, logger)
        TryActivateDraftDocument(draftDoc, logger)
        TryActivateWorkingDrawingSheet(draftDoc, logger)

        Dim lists As Object = TryGetPartsListsCollectionWithCount(draftDoc, logger, "working-sheet")
        If lists IsNot Nothing Then
            TryLogPartsListParentSheets(lists, logger)
            Return lists
        End If

        LogDraftSheetScanForPartsList(draftDoc, logger)

        If logger IsNot Nothing Then
            Dim fn As String = TryGetDocumentFullName(draftDoc)
            logger.Log("[PARTLISTDATA][PARTSLISTS][WARN] Sin listas de piezas en el DFT (Count=0). FullName=" & fn)
            logger.Log("[PARTLISTDATA][PARTSLISTS][HINT] La PART_LIST suele estar en «Hoja 1» (no en plantillas A4/A3/A2). Active esa pestaña en SE y reintente.")
        End If
        Return Nothing
    End Function

    ''' <summary>Activa la hoja de trabajo (p. ej. «Hoja 1»), no plantillas A4/A3 ni «2D Model».</summary>
    Friend Shared Function TryActivateWorkingDrawingSheet(draftDoc As Object, logger As Logger) As Boolean
        If draftDoc Is Nothing Then Return False
        Dim working As Object = TryResolveWorkingDrawingSheet(draftDoc, logger)
        If working Is Nothing Then
            If logger IsNot Nothing Then logger.Log("[PARTLISTDATA][SHEET][WARN] No se encontró hoja de trabajo (Hoja 1 / con vistas).")
            Return False
        End If
        Dim name As String = TryGetSheetName(working)
        Try
            CallByName(draftDoc, "Activate", CallType.Method)
        Catch
        End Try
        Try
            CallByName(working, "Activate", CallType.Method)
        Catch ex As Exception
            If logger IsNot Nothing Then logger.Log("[PARTLISTDATA][SHEET][ACTIVATE][WARN] " & ex.Message)
        End Try
        Try
            CallByName(draftDoc, "ActiveSheet", CallType.Set, working)
        Catch
        End Try
        If logger IsNot Nothing Then
            logger.Log("[PARTLISTDATA][SHEET] Hoja activa para PART_LIST: '" & name & "'")
        End If
        Return True
    End Function

    Private Shared Function TryResolveWorkingDrawingSheet(draftDoc As Object, logger As Logger) As Object
        If draftDoc Is Nothing Then Return Nothing
        Dim preferred As Object = Nothing
        Dim firstWithViews As Object = Nothing

        Dim visit = Sub(sh As Object)
                        If sh Is Nothing Then Return
                        Dim nm As String = TryGetSheetName(sh)
                        Dim nDv As Integer = TryGetDrawingViewCountOnSheet(sh)
                        If logger IsNot Nothing Then
                            logger.Log("[PARTLISTDATA][SHEET_SCAN] name='" & nm & "' drawingViews=" & nDv.ToString(CultureInfo.InvariantCulture) &
                                       " template=" & IsTemplateSheetName(nm).ToString(CultureInfo.InvariantCulture))
                        End If
                        If IsTemplateSheetName(nm) Then Return
                        If nDv > 0 AndAlso firstWithViews Is Nothing Then firstWithViews = sh
                        If nDv > 0 AndAlso IsPreferredWorkingSheetName(nm) AndAlso preferred Is Nothing Then preferred = sh
                        If preferred Is Nothing AndAlso IsPreferredWorkingSheetName(nm) Then preferred = sh
                    End Sub

        Dim scanned As Boolean = False
        Try
            Dim sections As Object = CallByName(draftDoc, "Sections", CallType.Get)
            If sections IsNot Nothing Then
                Dim nSec As Integer = CInt(CallByName(sections, "Count", CallType.Get))
                For si As Integer = 1 To nSec
                    Dim sec As Object = GetCollectionItem(sections, si)
                    If sec Is Nothing Then Continue For
                    Dim sheetsCol As Object = Nothing
                    Try : sheetsCol = CallByName(sec, "Sheets", CallType.Get) : Catch : End Try
                    If sheetsCol Is Nothing Then Continue For
                    Dim nSh As Integer = CInt(CallByName(sheetsCol, "Count", CallType.Get))
                    For shi As Integer = 1 To nSh
                        scanned = True
                        visit(GetCollectionItem(sheetsCol, shi))
                    Next
                Next
            End If
        Catch
        End Try

        If Not scanned Then
            Try
                Dim sheets As Object = CallByName(draftDoc, "Sheets", CallType.Get)
                If sheets IsNot Nothing Then
                    Dim n As Integer = CInt(CallByName(sheets, "Count", CallType.Get))
                    For i As Integer = 1 To n
                        visit(GetCollectionItem(sheets, i))
                    Next
                End If
            Catch
            End Try
        End If

        Return If(preferred, If(firstWithViews, Nothing))
    End Function

    Private Shared Function TryGetSheetName(sh As Object) As String
        If sh Is Nothing Then Return ""
        Try
            Return Convert.ToString(CallByName(sh, "Name", CallType.Get)).Trim()
        Catch
            Return ""
        End Try
    End Function

    Private Shared Function IsTwoDModelSheetName(name As String) As Boolean
        Return String.Equals(If(name, "").Trim(), "2D Model", StringComparison.OrdinalIgnoreCase)
    End Function

    Private Shared Function IsTemplateSheetName(name As String) As Boolean
        Dim n As String = If(name, "").Trim()
        If n.Length = 0 Then Return True
        If IsTwoDModelSheetName(n) Then Return True
        Dim u As String = n.ToUpperInvariant()
        If u.StartsWith("HOJA A", StringComparison.Ordinal) Then Return True
        If u.StartsWith("A1_", StringComparison.Ordinal) OrElse u.StartsWith("A2_", StringComparison.Ordinal) OrElse
            u.StartsWith("A3_", StringComparison.Ordinal) OrElse u.StartsWith("A4_", StringComparison.Ordinal) Then Return True
        Return False
    End Function

    Private Shared Function IsPreferredWorkingSheetName(name As String) As Boolean
        Dim n As String = If(name, "").Trim()
        If n.Length = 0 Then Return False
        If String.Equals(n, "Hoja 1", StringComparison.OrdinalIgnoreCase) Then Return True
        If String.Equals(n, "Hoja1", StringComparison.OrdinalIgnoreCase) Then Return True
        If n.StartsWith("Hoja 1", StringComparison.OrdinalIgnoreCase) AndAlso Not IsTemplateSheetName(n) Then Return True
        Return False
    End Function

    Private Shared Function TryGetDrawingViewCountOnSheet(sh As Object) As Integer
        If sh Is Nothing Then Return 0
        Try
            Dim dvs As Object = CallByName(sh, "DrawingViews", CallType.Get)
            If dvs Is Nothing Then Return 0
            Return CInt(CallByName(dvs, "Count", CallType.Get))
        Catch
            Return 0
        End Try
    End Function

    Private Shared Sub LogDraftSheetScanForPartsList(draftDoc As Object, logger As Logger)
        If logger Is Nothing OrElse draftDoc Is Nothing Then Return
        logger.Log("[PARTLISTDATA][SHEET_SCAN] --- Pestañas del DFT (la PART_LIST no está en plantillas A4/A3) ---")
        TryResolveWorkingDrawingSheet(draftDoc, logger)
    End Sub

    Private Shared Sub TryLogPartsListParentSheets(lists As Object, logger As Logger)
        If lists Is Nothing OrElse logger Is Nothing Then Return
        Try
            Dim n As Integer = CInt(CallByName(lists, "Count", CallType.Get))
            For i As Integer = 1 To n
                Dim pl As Object = GetCollectionItem(lists, i)
                If pl Is Nothing Then Continue For
                Dim parentName As String = ""
                Try
                    Dim par As Object = CallByName(pl, "Parent", CallType.Get)
                    parentName = TryGetSheetName(par)
                Catch
                End Try
                logger.Log("[PARTLISTDATA][PARTSLIST] index=" & i.ToString(CultureInfo.InvariantCulture) &
                           " parentSheet='" & parentName & "'")
            Next
        Catch
        End Try
    End Sub

    Private Shared Function TryGetPartsListsCollectionWithCount(draftDoc As Object, logger As Logger, scope As String) As Object
        If draftDoc Is Nothing Then Return Nothing
        For attempt As Integer = 1 To 2
            Dim lists As Object = Nothing
            Try : lists = CallByName(draftDoc, "PartsLists", CallType.Get) : Catch : End Try
            Dim n As Integer = 0
            If lists IsNot Nothing Then
                Try : n = CInt(CallByName(lists, "Count", CallType.Get)) : Catch : End Try
            End If
            If n > 0 Then
                If logger IsNot Nothing Then
                    logger.Log("[PARTLISTDATA][PARTSLISTS] Count=" & n.ToString(CultureInfo.InvariantCulture) &
                               " scope=" & scope &
                               If(attempt > 1, " retry=" & attempt.ToString(CultureInfo.InvariantCulture), ""))
                End If
                Return lists
            End If
            If attempt = 1 Then TryActivateDraftDocument(draftDoc, logger)
        Next
        Return Nothing
    End Function

    ''' <summary>Fuerza PartsList.Update en todas las listas accesibles; registra si Count=0.</summary>
    Friend Shared Sub TryUpdateAllPartsListsOnDraft(draftDoc As Object, logger As Logger)
        If draftDoc Is Nothing Then Return
        draftDoc = TryEnsureDraftWithPartsLists(draftDoc, logger)
        TryActivateWorkingDrawingSheet(draftDoc, logger)
        Dim lists As Object = TryGetDraftPartsLists(draftDoc, logger)
        If lists Is Nothing Then
            If logger IsNot Nothing Then
                logger.Log("[PARTSLIST][UPDATE][SKIP] No hay PartsLists; la PART_LIST del plano no se refrescará por API.")
                LogDraftTableInventory(draftDoc, logger)
            End If
            Return
        End If
        Dim n As Integer = 0
        Try : n = CInt(CallByName(lists, "Count", CallType.Get)) : Catch : End Try
        For i As Integer = 1 To n
            Dim pl As Object = GetCollectionItem(lists, i)
            If pl Is Nothing Then Continue For
            Try
                CallByName(pl, "Update", CallType.Method)
                If logger IsNot Nothing Then logger.Log("[PARTSLIST][UPDATE][OK] post_props index=" & i.ToString(CultureInfo.InvariantCulture))
            Catch ex As Exception
                If logger IsNot Nothing Then logger.Log("[PARTSLIST][UPDATE][WARN] " & ex.Message)
            End Try
        Next
        TrySamplePartsListCells(lists, logger)
    End Sub

    Private Shared Function TryNormalizeExistingFilePath(filePath As String) As String
        If String.IsNullOrWhiteSpace(filePath) Then Return ""
        Dim t As String = filePath.Trim()
        Try
            If IO.File.Exists(t) Then Return IO.Path.GetFullPath(t)
        Catch
        End Try
        Try
            Return IO.Path.GetFullPath(t)
        Catch
            Return t
        End Try
    End Function

    Private Shared Function DocumentPathsMatch(pathA As String, pathB As String) As Boolean
        If String.IsNullOrWhiteSpace(pathA) OrElse String.IsNullOrWhiteSpace(pathB) Then Return False
        Dim a As String = TryNormalizeExistingFilePath(pathA)
        Dim b As String = TryNormalizeExistingFilePath(pathB)
        If a.Length > 0 AndAlso b.Length > 0 AndAlso String.Equals(a, b, StringComparison.OrdinalIgnoreCase) Then Return True
        Return String.Equals(IO.Path.GetFileName(a), IO.Path.GetFileName(b), StringComparison.OrdinalIgnoreCase)
    End Function

    Private Shared Function TryGetDocumentFullName(doc As Object) As String
        If doc Is Nothing Then Return ""
        Try
            Return Convert.ToString(CallByName(doc, "FullName", CallType.Get)).Trim()
        Catch
            Return ""
        End Try
    End Function

    Private Shared Function TryGetApplicationFromDocument(doc As Object) As Application
        If doc Is Nothing Then Return Nothing
        Try
            Return TryCast(CallByName(doc, "Application", CallType.Get), Application)
        Catch
            Return Nothing
        End Try
    End Function

    Friend Shared Function TryGetDraftPartsListCount(draftDoc As Object) As Integer
        If draftDoc Is Nothing Then Return 0
        TryActivateWorkingDrawingSheet(draftDoc, Nothing)
        Try
            Dim lists As Object = Nothing
            Try : lists = CallByName(draftDoc, "PartsLists", CallType.Get) : Catch : End Try
            If lists Is Nothing Then Return 0
            Return CInt(CallByName(lists, "Count", CallType.Get))
        Catch
            Return 0
        End Try
    End Function

    Private Shared Function TryFindOpenDocumentByPath(app As Application, normalizedPath As String, logger As Logger,
                                                     isDraft As Boolean, requirePartsList As Boolean) As Object
        If app Is Nothing OrElse String.IsNullOrWhiteSpace(normalizedPath) Then Return Nothing

        Dim best As Object = Nothing
        Dim bestScore As Integer = -1
        Dim active As Object = Nothing
        Try : active = CallByName(app, "ActiveDocument", CallType.Get) : Catch : End Try

        Dim docs As Object = Nothing
        Try : docs = CallByName(app, "Documents", CallType.Get) : Catch : End Try
        If docs Is Nothing Then Return Nothing

        Dim count As Integer = 0
        Try : count = CInt(CallByName(docs, "Count", CallType.Get)) : Catch : End Try

        For i As Integer = 1 To count
            Dim d As Object = GetCollectionItem(docs, i)
            If d Is Nothing Then Continue For
            Dim fullName As String = TryGetDocumentFullName(d)
            If Not DocumentPathsMatch(fullName, normalizedPath) Then Continue For

            Dim plCount As Integer = 0
            If isDraft Then plCount = TryGetDraftPartsListCount(d)

            If requirePartsList AndAlso plCount <= 0 Then Continue For

            Dim score As Integer = 0
            If plCount > 0 Then score += 100
            If active IsNot Nothing AndAlso Object.ReferenceEquals(d, active) Then score += 20
            If isDraft AndAlso LooksLikeDraftDocument(d) Then score += 5

            If score > bestScore Then
                bestScore = score
                best = d
            End If
        Next

        If active IsNot Nothing AndAlso best Is Nothing AndAlso Not requirePartsList Then
            If DocumentPathsMatch(TryGetDocumentFullName(active), normalizedPath) Then Return active
        End If

        Return best
    End Function

    Private Shared Function LooksLikeDraftDocument(doc As Object) As Boolean
        If doc Is Nothing Then Return False
        Try
            Dim lists As Object = CallByName(doc, "PartsLists", CallType.Get)
            If lists IsNot Nothing Then Return True
        Catch
        End Try
        Try
            Dim sheets As Object = CallByName(doc, "Sheets", CallType.Get)
            If sheets IsNot Nothing Then Return True
        Catch
        End Try
        Return False
    End Function

    Private Shared Sub LogOpenDraftCandidates(app As Application, normalizedPath As String, logger As Logger)
        If logger Is Nothing OrElse app Is Nothing Then Return
        Try
            Dim docs As Object = CallByName(app, "Documents", CallType.Get)
            Dim count As Integer = CInt(CallByName(docs, "Count", CallType.Get))
            logger.Log("[PARTLISTDATA][DFT_RESOLVE][SCAN] Buscando '" & IO.Path.GetFileName(normalizedPath) & "' entre " &
                       count.ToString(CultureInfo.InvariantCulture) & " documento(s) abiertos:")
            For i As Integer = 1 To count
                Dim d As Object = GetCollectionItem(docs, i)
                If d Is Nothing Then Continue For
                Dim fn As String = TryGetDocumentFullName(d)
                Dim pl As Integer = TryGetDraftPartsListCount(d)
                Dim match As Boolean = DocumentPathsMatch(fn, normalizedPath)
                logger.Log("[PARTLISTDATA][DFT_RESOLVE][SCAN] #" & i.ToString(CultureInfo.InvariantCulture) &
                           " match=" & match.ToString(CultureInfo.InvariantCulture) &
                           " PartsLists=" & pl.ToString(CultureInfo.InvariantCulture) &
                           " '" & fn & "'")
            Next
        Catch ex As Exception
            logger.Log("[PARTLISTDATA][DFT_RESOLVE][SCAN][ERR] " & ex.Message)
        End Try
    End Sub

    Private Shared Function TryGetOpenDocumentByPath(app As Application, filePath As String) As Object
        If app Is Nothing OrElse String.IsNullOrWhiteSpace(filePath) Then Return Nothing
        Return TryFindOpenDocumentByPath(app, TryNormalizeExistingFilePath(filePath), Nothing, isDraft:=False, requirePartsList:=False)
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
                    n = If(n, "").Trim()
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
                itemName = itemName.Trim()
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

    ''' <summary>
    ''' Cierra documentos devueltos como <c>Object</c> (p. ej. <c>Documents.Open</c>) sin <c>CallByName</c> en <c>Close</c>:
    ''' el enlace en retardado a <c>Close</c> suele disparar <see cref="MissingMemberException"/> o <c>TargetParameterCountException</c> según el TLB de Solid Edge.
    ''' </summary>
    Friend Shared Sub TryCloseComDocument(doc As Object, Optional saveChanges As Boolean = False)
        If doc Is Nothing Then Return
        Dim typed = TryCast(doc, SolidEdgeDocument)
        If typed IsNot Nothing Then
            Try
                typed.Close(saveChanges)
                Return
            Catch
            End Try
        End If
        Try
            CallByName(doc, "Close", CallType.Method, saveChanges)
        Catch
            Try
                CallByName(doc, "Close", CallType.Method)
            Catch
            End Try
        End Try
    End Sub

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

    Friend Shared Function SetDocumentPropertyCount(doc As Object, setName As String, propName As String, value As String, logger As Logger) As Integer
        Return SetDocumentProperty(doc, setName, propName, value, logger)
    End Function

    Friend Shared Function TrySetCustomProperty(doc As Object, propName As String, value As String, logger As Logger) As Boolean
        If doc Is Nothing OrElse String.IsNullOrWhiteSpace(propName) Then Return False
        If value Is Nothing Then value = ""
        If value.Trim() = "" Then Return False
        Return SetDocumentProperty(doc, "Custom", propName, value.Trim(), logger) > 0
    End Function

    ''' <summary>Listas PART_LIST enlazadas al modelo: refuerzo de calibre y masa en MechanicalModeling tras Custom.</summary>
    ''' <param name="mirrorMaterialToMechanical">False evita escribir Material en MechanicalModeling: Solid Edge suele normalizar a nombre de stock (p. ej. AISI304) y la lista nativa lo muestra en lugar del texto de la UI.</param>
    Friend Shared Function TryMirrorPartListOntoMechanicalModel(doc As Object, material As String, thicknessGauge As String, weightRaw As String, logger As Logger,
                                                              Optional mirrorMaterialToMechanical As Boolean = True) As Integer
        Dim cnt As Integer = 0
        If doc Is Nothing Then Return 0
        If mirrorMaterialToMechanical AndAlso Not String.IsNullOrWhiteSpace(material) Then
            Dim m = material.Trim()
            If SetDocumentProperty(doc, "MechanicalModeling", "Material", m, logger) > 0 Then
                cnt += 1
                If logger IsNot Nothing Then logger.Log("[PARTLISTDATA][WRITE_MODEL_MM] MechanicalModeling.Material=" & m)
            End If
        ElseIf Not mirrorMaterialToMechanical AndAlso Not String.IsNullOrWhiteSpace(material) AndAlso logger IsNot Nothing Then
            logger.Log("[PARTLISTDATA][WRITE_MODEL_MM][SKIP] MechanicalModeling.Material (evitar stock abreviado en PART_LIST; usar Custom + celdas)")
        End If
        If Not String.IsNullOrWhiteSpace(thicknessGauge) Then
            Dim g = thicknessGauge.Trim()
            If SetDocumentProperty(doc, "MechanicalModeling", "Sheet Metal Gauge", g, logger) > 0 Then
                cnt += 1
                If logger IsNot Nothing Then logger.Log("[PARTLISTDATA][WRITE_MODEL_MM] MechanicalModeling.Sheet Metal Gauge=" & g)
            End If
        End If
        If Not String.IsNullOrWhiteSpace(weightRaw) Then
            Dim w = weightRaw.Trim().Replace(",", ".")
            If SetDocumentProperty(doc, "MechanicalModeling", "Mass", w, logger) > 0 Then
                cnt += 1
                If logger IsNot Nothing Then logger.Log("[PARTLISTDATA][WRITE_MODEL_MM] MechanicalModeling.Mass=" & w)
            End If
        End If
        Return cnt
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

    Private Shared Function TryGetPartsListCellTextInternal(pl As Object, row1 As Integer, col1 As Integer, Optional maxCellLen As Integer = 48) As String
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
        Dim tc As TableCell = TryCast(cellObj, TableCell)
        If tc IsNot Nothing Then
            Try
                Dim sv As String = Convert.ToString(tc.value, CultureInfo.CurrentCulture).Trim()
                If Not String.IsNullOrWhiteSpace(sv) Then
                    If maxCellLen > 0 AndAlso sv.Length > maxCellLen Then sv = sv.Substring(0, maxCellLen) & "…"
                    Return sv
                End If
            Catch
            End Try
        End If
        Try
            Dim s As String = Convert.ToString(CallByName(cellObj, "Text", CallType.Get))
            If String.IsNullOrWhiteSpace(s) Then s = Convert.ToString(CallByName(cellObj, "Value", CallType.Get))
            If String.IsNullOrWhiteSpace(s) Then
                Try : s = Convert.ToString(CallByName(cellObj, "value", CallType.Get)) : Catch : End Try
            End If
            If Not String.IsNullOrWhiteSpace(s) Then
                s = s.Replace(vbCr, " ").Replace(vbLf, " ").Trim()
                If maxCellLen > 0 AndAlso s.Length > maxCellLen Then s = s.Substring(0, maxCellLen) & "…"
                Return s
            End If
        Catch
        End Try
        Return ""
    End Function

    Friend Shared Sub TryLogPartsListCellReadback(pl As Object, listIdx As Integer, cols As IEnumerable(Of Integer), logger As Logger,
                                                  Optional maxRows As Integer = 3)
        If pl Is Nothing OrElse logger Is Nothing OrElse cols Is Nothing Then Return
        Dim colList = cols.Where(Function(c) c > 0).Distinct().ToList()
        If colList.Count = 0 Then Return
        Dim hdrRows As Integer = TryGetPartsListHeaderRowCount(pl)
        If hdrRows <= 0 Then hdrRows = 1
        Dim rowCount As Integer = TryGetPartsListRowCount(pl)
        Dim maxR As Integer = Math.Max(hdrRows + Math.Max(rowCount, 1), maxRows)
        maxR = Math.Min(maxR, hdrRows + 4)
        Dim parts As New List(Of String)()
        For r As Integer = 1 To maxR
            For Each c In colList
                Dim t As String = TryGetPartsListCellTextInternal(pl, r, c, 64)
                If String.IsNullOrWhiteSpace(t) Then Continue For
                parts.Add("r" & r.ToString(CultureInfo.InvariantCulture) & "c" & c.ToString(CultureInfo.InvariantCulture) & "='" & t.Replace("'", "") & "'")
            Next
        Next
        If parts.Count > 0 Then
            logger.Log("[PARTSLIST][CELL_READBACK] list=" & listIdx.ToString(CultureInfo.InvariantCulture) &
                       " hdrRows=" & hdrRows.ToString(CultureInfo.InvariantCulture) &
                       " rows=" & rowCount.ToString(CultureInfo.InvariantCulture) & " " &
                       String.Join("; ", parts.Take(20)))
        Else
            logger.Log("[PARTSLIST][CELL_READBACK][WARN] list=" & listIdx.ToString(CultureInfo.InvariantCulture) &
                       " sin texto en filas 1.." & maxR.ToString(CultureInfo.InvariantCulture))
        End If
    End Sub

    Friend Shared Function TryReadPartsListCellText(pl As Object, row1 As Integer, col1 As Integer, Optional maxCellLen As Integer = 640) As String
        Return TryGetPartsListCellTextInternal(pl, row1, col1, maxCellLen)
    End Function

    ''' <summary>Obtiene <see cref="TableCell"/> desde <c>PartsList</c> o <c>Table</c> (pestaña Datos del SDK).</summary>
    Friend Shared Function TryGetTableCell(tableOrPartsList As Object, row1 As Integer, col1 As Integer) As Object
        If tableOrPartsList Is Nothing OrElse row1 <= 0 OrElse col1 <= 0 Then Return Nothing
        Try
            Return CallByName(tableOrPartsList, "Cell", CallType.Get, row1, col1)
        Catch
        End Try
        Try
            Return CallByName(tableOrPartsList, "Cell", CallType.Method, row1, col1)
        Catch
            Return Nothing
        End Try
    End Function

    Friend Shared Function TryReadTableCellPropertyText(cellObj As Object) As String
        If cellObj Is Nothing Then Return ""
        Dim tc As TableCell = TryCast(cellObj, TableCell)
        If tc IsNot Nothing Then
            Try : Return If(tc.PropertyText, "").Trim() : Catch : End Try
        End If
        Try
            Return Convert.ToString(CallByName(cellObj, "PropertyText", CallType.Get)).Trim()
        Catch
            Return ""
        End Try
    End Function

    Friend Shared Function TryReadTableCellIsOverridden(cellObj As Object) As Boolean?
        If cellObj Is Nothing Then Return Nothing
        Dim tc As TableCell = TryCast(cellObj, TableCell)
        If tc IsNot Nothing Then
            Try : Return tc.IsOverridden : Catch : End Try
        End If
        Try
            Return CBool(CallByName(cellObj, "IsOverridden", CallType.Get))
        Catch
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Equivalente API de la pestaña «Datos» en Propiedades de tabla (SDK <c>TableCell</c>):
    ''' <c>AllowCellOverrides=True</c>, <c>value</c>; comprobar <c>IsOverridden</c> tras escribir.
    ''' </summary>
    Friend Shared Function TryWriteTableCellDataTabOverride(tableOrPartsList As Object, row1 As Integer, col1 As Integer,
                                                            value As String, logger As Logger, logVerb As String) As Boolean
        Dim cellObj As Object = TryGetTableCell(tableOrPartsList, row1, col1)
        If cellObj Is Nothing Then
            If logger IsNot Nothing Then
                logger.Log(logVerb & "[CELL_WRITE][SKIP] r=" & row1.ToString(CultureInfo.InvariantCulture) &
                           " c=" & col1.ToString(CultureInfo.InvariantCulture) & " sin TableCell.")
            End If
            Return False
        End If
        Return TryWriteInteropTableCellString(cellObj, value, logger, logVerb, row1, col1)
    End Function

    Friend Shared Function TryWriteInteropTableCellString(cellObj As Object, value As String, logger As Logger, logVerb As String,
                                                        Optional row1 As Integer = 0, Optional col1 As Integer = 0) As Boolean
        If cellObj Is Nothing OrElse String.IsNullOrWhiteSpace(value) Then Return False
        Dim v As String = value.Trim()
        Dim pos As String = ""
        If row1 > 0 AndAlso col1 > 0 Then
            pos = " r=" & row1.ToString(CultureInfo.InvariantCulture) & " c=" & col1.ToString(CultureInfo.InvariantCulture)
        End If

        Dim tc As TableCell = TryCast(cellObj, TableCell)
        If tc IsNot Nothing Then
            Try
                tc.AllowCellOverrides = True
                tc.value = v
                If logger IsNot Nothing Then
                    Dim pt As String = ""
                    Try : pt = If(tc.PropertyText, "").Replace("'", "").Trim() : Catch : End Try
                    Dim over As Boolean = False
                    Try : over = tc.IsOverridden : Catch : End Try
                    logger.Log(logVerb & "[CELL_WRITE][OK]" & pos & " value='" & v.Replace("'", "") &
                               "' IsOverridden=" & over.ToString(CultureInfo.InvariantCulture) &
                               If(pt.Length > 0, " PropertyText='" & pt & "'", ""))
                End If
                Return True
            Catch ex As Exception
                If logger IsNot Nothing Then logger.Log(logVerb & "[CELL_WRITE][TableCell]" & pos & " " & ex.Message)
            End Try
        End If

        Try
            CallByName(cellObj, "AllowCellOverrides", CallType.Set, True)
        Catch
        End Try
        Try
            CallByName(cellObj, "value", CallType.Set, v)
            If logger IsNot Nothing Then
                Dim overLate As Boolean? = TryReadTableCellIsOverridden(cellObj)
                logger.Log(logVerb & "[CELL_WRITE][OK][late]" & pos & " value='" & v.Replace("'", "") &
                           "' IsOverridden=" & If(overLate.HasValue, overLate.Value.ToString(CultureInfo.InvariantCulture), "?"))
            End If
            Return True
        Catch ex As Exception
            If logger IsNot Nothing Then logger.Log(logVerb & "[CELL_WRITE][late_value]" & pos & " " & ex.Message)
        End Try

        Dim rt As Type = cellObj.GetType()
        For Each pi As PropertyInfo In rt.GetProperties(BindingFlags.Public Or BindingFlags.Instance)
            If Not pi.CanWrite Then Continue For
            If Not String.Equals(pi.Name, "value", StringComparison.OrdinalIgnoreCase) Then Continue For
            Try
                pi.SetValue(cellObj, v, Nothing)
                If logger IsNot Nothing Then logger.Log(logVerb & "[CELL_WRITE][OK][reflect]" & pos)
                Return True
            Catch ex As Exception
                If logger IsNot Nothing Then logger.Log(logVerb & "[CELL_WRITE][reflect_value]" & pos & " " & ex.Message)
            End Try
            Exit For
        Next
        Return False
    End Function

    ''' <summary>Cuando las columnas están enlazadas a propiedades, forzar texto: en el SDK de Draft, TableCell.AllowCellOverrides permite sobrescribir el valor mostrado.</summary>
    Friend Shared Function TryWritePartsListCellText(pl As Object, row1 As Integer, col1 As Integer, value As String, logger As Logger) As Boolean
        If pl Is Nothing OrElse col1 <= 0 OrElse row1 <= 0 Then Return False
        If value Is Nothing OrElse value.Trim() = "" Then Return False
        If TryWriteTableCellDataTabOverride(pl, row1, col1, value, logger, "[PARTSLIST]") Then Return True
        Dim v As String = value.Trim()
        Dim cellObj As Object = TryGetTableCell(pl, row1, col1)
        If cellObj Is Nothing Then Return False

        If TryWriteInteropTableCellString(cellObj, v, logger, "[PARTSLIST]", row1, col1) Then Return True

        Try
            CallByName(cellObj, "AllowCellOverrides", CallType.Set, True)
        Catch
        End Try

        Try
            CallByName(cellObj, "value", CallType.Set, v)
            Return True
        Catch ex As Exception
            If logger IsNot Nothing Then logger.Log("[PARTSLIST][CELL_WRITE][value] r=" & row1.ToString(CultureInfo.InvariantCulture) & " c=" & col1.ToString(CultureInfo.InvariantCulture) & " " & ex.Message)
        End Try
        Try
            CallByName(cellObj, "Text", CallType.Set, v)
            Return True
        Catch ex As Exception
            If logger IsNot Nothing Then logger.Log("[PARTSLIST][CELL_WRITE][Text] r=" & row1.ToString(CultureInfo.InvariantCulture) & " c=" & col1.ToString(CultureInfo.InvariantCulture) & " " & ex.Message)
        End Try
        Try
            CallByName(cellObj, "Value", CallType.Set, v)
            Return True
        Catch ex As Exception
            If logger IsNot Nothing Then logger.Log("[PARTSLIST][CELL_WRITE][Value] r=" & row1.ToString(CultureInfo.InvariantCulture) & " c=" & col1.ToString(CultureInfo.InvariantCulture) & " " & ex.Message)
        End Try
        Try
            CallByName(cellObj, "Formula", CallType.Set, "")
            CallByName(cellObj, "value", CallType.Set, v)
            Return True
        Catch ex As Exception
            If logger IsNot Nothing Then logger.Log("[PARTSLIST][CELL_WRITE][FormulaClear] r=" & row1.ToString(CultureInfo.InvariantCulture) & " c=" & col1.ToString(CultureInfo.InvariantCulture) & " " & ex.Message)
        End Try
        Return False
    End Function

    ''' <summary>
    ''' Una fila de log por columna: cabecera visible y <see cref="SolidEdgeDraft.TableColumn.PropertyText"/> (plantilla que rellena filas; enlaces tipo propiedad personalizada).
    ''' Ver ayuda Draft: <c>SolidEdgeDraft~TableColumn~PropertyText.html</c> en <c>docs/SDK_HTML</c>.
    ''' </summary>
    Friend Shared Sub TryLogPartsListColumnDefinitions(pl As Object, listIndex As Integer, logger As Logger,
                                                       Optional maxPropTextChars As Integer = 360)
        If pl Is Nothing OrElse logger Is Nothing Then Return
        Dim nc As Integer = TryGetPartsListColumnCount(pl)
        If nc <= 0 Then Return
        Dim lim As Integer = Math.Max(80, maxPropTextChars)
        For c As Integer = 1 To nc
            Dim header As String = TryReadPartsListColumnHeader(pl, c)
            Dim propText As String = ""
            Dim showCol As String = ""
            Try
                Dim cols As Object = Nothing
                Try : cols = CallByName(pl, "Columns", CallType.Get) : Catch : End Try
                If cols IsNot Nothing Then
                    Dim col As Object = GetCollectionItem(cols, c)
                    If col IsNot Nothing Then
                        Try : propText = Convert.ToString(CallByName(col, "PropertyText", CallType.Get)) : Catch : End Try
                        Try : showCol = Convert.ToString(CallByName(col, "Show", CallType.Get)) : Catch : End Try
                    End If
                End If
            Catch
            End Try
            propText = If(propText, "").Replace(vbCr, " ").Replace(vbLf, " ").Trim()
            If propText.Length > lim Then propText = propText.Substring(0, lim) & "…"
            logger.Log("[PARTSLIST][COL_DIAG] list=" & listIndex.ToString(CultureInfo.InvariantCulture) &
                       " col=" & c.ToString(CultureInfo.InvariantCulture) &
                       " show=" & If(showCol, "").Trim() &
                       " header='" & header.Replace("'", "") & "'" &
                       " PropertyText='" & propText.Replace("'", "") & "'")
        Next
    End Sub

    Friend Shared Function TryReadPartsListColumnHeader(pl As Object, col1 As Integer) As String
        If pl Is Nothing Then Return ""
        Try
            Dim cols As Object = Nothing
            Try : cols = CallByName(pl, "Columns", CallType.Get) : Catch : End Try
            If cols Is Nothing Then Return ""
            Dim col As Object = GetCollectionItem(cols, col1)
            If col Is Nothing Then Return ""
            Dim h As String = ""
            Try : h = Convert.ToString(CallByName(col, "Header", CallType.Get)) : Catch : End Try
            h = If(h, "").Replace(vbCr, " ").Replace(vbLf, " ").Trim()
            If h.Length > 0 Then Return h
            Try : h = Convert.ToString(CallByName(col, "HeaderRowValue", CallType.Get)) : Catch : End Try
            h = If(h, "").Replace(vbCr, " ").Replace(vbLf, " ").Trim()
            If h.Length > 0 Then Return h
            Return TryReadPartsListCellText(pl, 1, col1, 256).Replace(vbCr, " ").Replace(vbLf, " ").Trim()
        Catch
            Return ""
        End Try
    End Function

    Friend Shared Function TryGetPartsListColumnCount(pl As Object) As Integer
        If pl Is Nothing Then Return 0
        Try
            Dim cols As Object = Nothing
            Try : cols = CallByName(pl, "Columns", CallType.Get) : Catch : End Try
            If cols Is Nothing Then Return 0
            Return CInt(CallByName(cols, "Count", CallType.Get))
        Catch
            Return 0
        End Try
    End Function

    Friend Shared Function TryGetPartsListRowCount(pl As Object) As Integer
        If pl Is Nothing Then Return 0
        Try
            Dim rows As Object = Nothing
            Try : rows = CallByName(pl, "Rows", CallType.Get) : Catch : End Try
            If rows Is Nothing Then Return 0
            Return CInt(CallByName(rows, "Count", CallType.Get))
        Catch
            Try
                Return CInt(CallByName(pl, "RowCount", CallType.Get))
            Catch
                Return 0
            End Try
        End Try
    End Function

    Friend Shared Function TryGetPartsListHeaderRowCount(pl As Object) As Integer
        If pl Is Nothing Then Return 0
        Try
            Return CInt(CallByName(pl, "NumberOfHeaderRows", CallType.Get))
        Catch
            Return 1
        End Try
    End Function

    ''' <summary>
    ''' Índice de fila para <see cref="PartsList.Cell"/>. Con <c>Rows.Count=1</c> la única fila COM es la de datos (fila 1);
    ''' no usar <c>hdrRows+1</c> (fila 2 no tiene TableCell — ver log <c>sin TableCell</c>).
    ''' </summary>
    Friend Shared Function ResolvePartsListDataRowIndex(hdrRows As Integer, rowCount As Integer) As Integer
        If rowCount <= 0 Then Return 1
        If rowCount = 1 Then Return 1
        If hdrRows <= 0 Then hdrRows = 1
        Return Math.Max(1, hdrRows + 1)
    End Function

    ''' <summary>Candidatos de fila: con una sola fila de datos solo <c>1</c>; si hay varias, <c>hdrRows+1</c> y <c>1</c>.</summary>
    Friend Shared Function ResolvePartsListDataRowCandidates(hdrRows As Integer, rowCount As Integer) As Integer()
        Dim primary As Integer = ResolvePartsListDataRowIndex(hdrRows, rowCount)
        Dim alt As New List(Of Integer) From {primary}
        If rowCount > 1 AndAlso primary <> 1 Then alt.Add(1)
        Return alt.Distinct().ToArray()
    End Function

    Friend Shared Function TryActivatePartsListParentSheet(pl As Object, draftDoc As Object, logger As Logger) As Boolean
        If pl Is Nothing Then Return False
        Dim par As Object = Nothing
        Try : par = CallByName(pl, "Parent", CallType.Get) : Catch : End Try
        If par Is Nothing Then Return False
        Dim sheetName As String = TryGetSheetName(par)
        Try
            If draftDoc IsNot Nothing Then
                Try : CallByName(draftDoc, "Activate", CallType.Method) : Catch : End Try
                Try : CallByName(draftDoc, "ActiveSheet", CallType.Set, par) : Catch : End Try
            End If
            CallByName(par, "Activate", CallType.Method)
        Catch ex As Exception
            If logger IsNot Nothing Then logger.Log("[PARTLISTDATA][SHEET][PARTSLIST_PARENT][WARN] " & ex.Message)
            Return False
        End Try
        If logger IsNot Nothing AndAlso sheetName.Length > 0 Then
            logger.Log("[PARTLISTDATA][SHEET] PartsList parent='" & sheetName.Replace("'", "") & "'")
        End If
        Return True
    End Function

    Friend Shared Function TryPartsListCellDisplaysValue(pl As Object, row1 As Integer, col1 As Integer, expected As String) As Boolean
        If pl Is Nothing OrElse col1 <= 0 OrElse String.IsNullOrWhiteSpace(expected) Then Return False
        Dim cellText As String = TryReadPartsListCellText(pl, row1, col1, 256)
        If String.IsNullOrWhiteSpace(cellText) Then Return False
        Dim a As String = expected.Trim().Replace(" ", "").Replace(",", ".")
        Dim b As String = cellText.Trim().Replace(" ", "").Replace(",", ".")
        If b.IndexOf(a, StringComparison.OrdinalIgnoreCase) >= 0 Then Return True
        If a.IndexOf(b, StringComparison.OrdinalIgnoreCase) >= 0 AndAlso b.Length > 0 Then Return True
        Return False
    End Function

    Friend Shared Function TryGetDraftDocumentTableCount(dftDoc As Object) As Integer
        If dftDoc Is Nothing Then Return 0
        Try
            Dim tablesObj As Object = Nothing
            Try : tablesObj = CallByName(dftDoc, "Tables", CallType.Get) : Catch : End Try
            If tablesObj Is Nothing Then Return 0
            Return CInt(CallByName(tablesObj, "Count", CallType.Get))
        Catch
            Return 0
        End Try
    End Function

    Friend Shared Function TryGetDraftDraftTablesCount(dftDoc As Object) As Integer
        If dftDoc Is Nothing Then Return 0
        Try
            Dim tablesObj As Object = Nothing
            Try : tablesObj = CallByName(dftDoc, "DraftTables", CallType.Get) : Catch : End Try
            If tablesObj Is Nothing Then Return 0
            Return CInt(CallByName(tablesObj, "Count", CallType.Get))
        Catch
            Return 0
        End Try
    End Function

    Friend Shared Sub LogDraftTableInventory(dftDoc As Object, logger As Logger)
        If dftDoc Is Nothing OrElse logger Is Nothing Then Return
        logger.Log("[PARTLISTDATA][TABLE_INVENTORY] Tables=" & TryGetDraftDocumentTableCount(dftDoc).ToString(CultureInfo.InvariantCulture) &
                   " DraftTables=" & TryGetDraftDraftTablesCount(dftDoc).ToString(CultureInfo.InvariantCulture))
        Dim idx As Integer = 0
        For Each tbl As Object In TryEnumerateAllDraftTables(dftDoc)
            idx += 1
            Dim sheetName As String = TryGetDraftTableSheetName(tbl)
            Dim cols As Integer = TryGetDraftTableColumnCount(tbl)
            Dim rows As Integer = TryGetDraftTableRowCount(tbl)
            logger.Log("[PARTLISTDATA][TABLE_INVENTORY] #" & idx.ToString(CultureInfo.InvariantCulture) &
                       " sheet='" & sheetName.Replace("'", "") & "'" &
                       " cols=" & cols.ToString(CultureInfo.InvariantCulture) &
                       " rows=" & rows.ToString(CultureInfo.InvariantCulture))
        Next
    End Sub

    Friend Shared Function TryEnumerateAllDraftTables(dftDoc As Object) As IEnumerable(Of Object)
        Dim result As New List(Of Object)()
        If dftDoc Is Nothing Then Return result
        For Each collName As String In New String() {"Tables", "DraftTables"}
            Dim tablesObj As Object = Nothing
            Try : tablesObj = CallByName(dftDoc, collName, CallType.Get) : Catch : End Try
            If tablesObj Is Nothing Then Continue For
            Dim n As Integer = 0
            Try : n = CInt(CallByName(tablesObj, "Count", CallType.Get)) : Catch : End Try
            For i As Integer = 1 To n
                Dim tbl As Object = GetCollectionItem(tablesObj, i)
                If tbl IsNot Nothing Then result.Add(tbl)
            Next
        Next
        Return result
    End Function

    Friend Shared Function TryGetDraftTableSheetName(tbl As Object) As String
        If tbl Is Nothing Then Return ""
        Try
            Dim sh As Object = CallByName(tbl, "Sheet", CallType.Get)
            Return TryGetSheetName(sh)
        Catch
            Return ""
        End Try
    End Function

    Friend Shared Function TryGetResolvedWorkingDrawingSheet(draftDoc As Object, logger As Logger) As Object
        If draftDoc Is Nothing Then Return Nothing
        TryActivateWorkingDrawingSheet(draftDoc, logger)
        Return TryResolveWorkingDrawingSheet(draftDoc, logger)
    End Function

    Friend Shared Function TryScoreDraftTableAsPartList(tbl As Object) As Integer
        If tbl Is Nothing Then Return 0
        Dim colCount As Integer = TryGetDraftTableColumnCount(tbl)
        If colCount <= 0 Then Return 0
        Dim score As Integer = 0
        For c As Integer = 1 To colCount
            Dim h As String = TryReadDraftTableColumnHeader(tbl, c)
            If h.Length = 0 Then Continue For
            Dim u As String = h.ToUpperInvariant()
            If u.Contains("ESPESOR") OrElse u.Contains("THICK") OrElse u.Contains("CALIBRE") Then score += 2
            If h.Length = 1 AndAlso "LHD".IndexOf(Char.ToUpperInvariant(h(0))) >= 0 Then score += 2
            If u.StartsWith("LARGO") OrElse u.StartsWith("ALTO") OrElse u.StartsWith("PROF") Then score += 1
            If u.Contains("MATERIAL") Then score += 1
        Next
        Return score
    End Function

    Friend Shared Sub TryRefreshDraftTable(tbl As Object, logger As Logger)
        If tbl Is Nothing Then Return
        Try
            CallByName(tbl, "Update", CallType.Method)
            If logger IsNot Nothing Then logger.Log("[PARTLISTDATA][TABLE_FALLBACK][UPDATE][OK]")
        Catch ex As Exception
            If logger IsNot Nothing Then logger.Log("[PARTLISTDATA][TABLE_FALLBACK][UPDATE][WARN] " & ex.Message)
        End Try
    End Sub

    ''' <summary>Convierte la primera lista de piezas en tabla genérica y devuelve la nueva <c>Table</c> (última de <c>DraftDocument.Tables</c> si el conteo aumenta en uno).</summary>
    Friend Shared Function TryConvertFirstPartsListToTable(dftDoc As Object, logger As Logger) As Object
        If dftDoc Is Nothing Then Return Nothing
        Dim nBefore As Integer = TryGetDraftDocumentTableCount(dftDoc)
        Dim lists As Object = Nothing
        Try : lists = CallByName(dftDoc, "PartsLists", CallType.Get) : Catch : End Try
        If lists Is Nothing Then Return Nothing
        Dim nPl As Integer = 0
        Try : nPl = CInt(CallByName(lists, "Count", CallType.Get)) : Catch : End Try
        If nPl < 1 Then Return Nothing
        Dim pl As Object = GetCollectionItem(lists, 1)
        If pl Is Nothing Then Return Nothing
        Try
            CallByName(pl, "ConvertToTable", CallType.Method)
        Catch ex As Exception
            If logger IsNot Nothing Then logger.Log("[PARTSLIST][CONVERT_TABLE][ERR] " & ex.Message)
            Return Nothing
        End Try
        Dim tablesObj As Object = Nothing
        Try : tablesObj = CallByName(dftDoc, "Tables", CallType.Get) : Catch : End Try
        If tablesObj Is Nothing Then Return Nothing
        Dim nAfter As Integer = 0
        Try : nAfter = CInt(CallByName(tablesObj, "Count", CallType.Get)) : Catch : End Try
        If nAfter <= nBefore Then
            If logger IsNot Nothing Then
                logger.Log("[PARTSLIST][CONVERT_TABLE][WARN] tablas " & nBefore.ToString(CultureInfo.InvariantCulture) & " → " & nAfter.ToString(CultureInfo.InvariantCulture) & " (sin nueva tabla).")
            End If
            Return Nothing
        End If
        Return GetCollectionItem(tablesObj, nAfter)
    End Function

    Friend Shared Function TryGetDraftTableColumnCount(tbl As Object) As Integer
        If tbl Is Nothing Then Return 0
        Try
            Dim cols As Object = Nothing
            Try : cols = CallByName(tbl, "Columns", CallType.Get) : Catch : End Try
            If cols Is Nothing Then Return 0
            Return CInt(CallByName(cols, "Count", CallType.Get))
        Catch
            Return 0
        End Try
    End Function

    Friend Shared Function TryGetDraftTableRowCount(tbl As Object) As Integer
        If tbl Is Nothing Then Return 0
        Try
            Dim rows As Object = Nothing
            Try : rows = CallByName(tbl, "Rows", CallType.Get) : Catch : End Try
            If rows Is Nothing Then Return 0
            Return CInt(CallByName(rows, "Count", CallType.Get))
        Catch
            Return 0
        End Try
    End Function

    Friend Shared Function TryGetDraftTableHeaderRowCount(tbl As Object) As Integer
        If tbl Is Nothing Then Return 0
        Try
            Return CInt(CallByName(tbl, "NumberOfHeaderRows", CallType.Get))
        Catch
            Return 1
        End Try
    End Function

    Friend Shared Function TryReadDraftTableColumnPropertyText(tbl As Object, col1 As Integer) As String
        If tbl Is Nothing OrElse col1 <= 0 Then Return ""
        Try
            Dim cols As Object = Nothing
            Try : cols = CallByName(tbl, "Columns", CallType.Get) : Catch : End Try
            If cols Is Nothing Then Return ""
            Dim col As Object = GetCollectionItem(cols, col1)
            If col Is Nothing Then Return ""
            Return Convert.ToString(CallByName(col, "PropertyText", CallType.Get)).Trim()
        Catch
            Return ""
        End Try
    End Function

    Friend Shared Sub TryLogDraftTableColumnDefinitions(tbl As Object, logger As Logger)
        If tbl Is Nothing OrElse logger Is Nothing Then Return
        Dim nc As Integer = TryGetDraftTableColumnCount(tbl)
        For c As Integer = 1 To nc
            Dim header As String = TryReadDraftTableColumnHeader(tbl, c)
            Dim propText As String = TryReadDraftTableColumnPropertyText(tbl, c)
            propText = propText.Replace(vbCr, " ").Replace(vbLf, " ").Trim()
            If propText.Length > 200 Then propText = propText.Substring(0, 200) & "…"
            logger.Log("[TABLE][COL_DIAG] col=" & c.ToString(CultureInfo.InvariantCulture) &
                       " header='" & header.Replace("'", "") & "'" &
                       " PropertyText='" & propText.Replace("'", "") & "'")
        Next
    End Sub

    Friend Shared Function TryReadDraftTableColumnHeader(tbl As Object, col1 As Integer) As String
        If tbl Is Nothing OrElse col1 <= 0 Then Return ""
        Try
            Dim cols As Object = Nothing
            Try : cols = CallByName(tbl, "Columns", CallType.Get) : Catch : End Try
            If cols Is Nothing Then Return ""
            Dim col As Object = GetCollectionItem(cols, col1)
            If col Is Nothing Then Return ""
            Dim h As String = ""
            Try : h = Convert.ToString(CallByName(col, "HeaderRowValue", CallType.Get)) : Catch : End Try
            If String.IsNullOrWhiteSpace(h) Then
                Try : h = Convert.ToString(CallByName(col, "Header", CallType.Get)) : Catch : End Try
            End If
            Return If(h, "").Replace(vbCr, " ").Replace(vbLf, " ").Trim()
        Catch
            Return ""
        End Try
    End Function

    Friend Shared Function TryWriteDraftTableCellValue(tbl As Object, row1 As Integer, col1 As Integer, value As String, logger As Logger) As Boolean
        If tbl Is Nothing OrElse col1 <= 0 OrElse row1 <= 0 Then Return False
        If value Is Nothing OrElse value.Trim() = "" Then Return False
        If TryWriteTableCellDataTabOverride(tbl, row1, col1, value, logger, "[TABLE]") Then Return True
        Dim v As String = value.Trim()
        Dim cellObj As Object = TryGetTableCell(tbl, row1, col1)
        If cellObj Is Nothing Then Return False
        If TryWriteInteropTableCellString(cellObj, v, logger, "[TABLE]", row1, col1) Then Return True
        Try
            CallByName(cellObj, "AllowCellOverrides", CallType.Set, True)
        Catch
        End Try
        Try
            CallByName(cellObj, "value", CallType.Set, v)
            Return True
        Catch ex As Exception
            If logger IsNot Nothing Then logger.Log("[TABLE][CELL_WRITE][value] r=" & row1.ToString(CultureInfo.InvariantCulture) & " c=" & col1.ToString(CultureInfo.InvariantCulture) & " " & ex.Message)
        End Try
        Try
            CallByName(cellObj, "Value", CallType.Set, v)
            Return True
        Catch ex As Exception
            If logger IsNot Nothing Then logger.Log("[TABLE][CELL_WRITE][Value] r=" & row1.ToString(CultureInfo.InvariantCulture) & " c=" & col1.ToString(CultureInfo.InvariantCulture) & " " & ex.Message)
        End Try
        Try
            CallByName(cellObj, "Text", CallType.Set, v)
            Return True
        Catch ex As Exception
            If logger IsNot Nothing Then logger.Log("[TABLE][CELL_WRITE][Text] r=" & row1.ToString(CultureInfo.InvariantCulture) & " c=" & col1.ToString(CultureInfo.InvariantCulture) & " " & ex.Message)
        End Try
        Return False
    End Function

    ''' <summary>Busca en todos los conjuntos una propiedad cuyo <i>Name</i> contenga la cadena (p. ej. «pedido» en claves locales).</summary>
    Friend Shared Function TryFindDocumentPropertyValueByLoosePropertyName(doc As Object, nameNeedle As String) As String
        If doc Is Nothing OrElse String.IsNullOrWhiteSpace(nameNeedle) Then Return ""
        Dim needle As String = nameNeedle.Trim()
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
                    pn = pn.Trim()
                    If pn.IndexOf(needle, StringComparison.OrdinalIgnoreCase) < 0 Then Continue For
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

    Friend Shared Sub RefreshNativePartsListsAndUpdateAll(dftDoc As Object, logger As Logger)
        If dftDoc Is Nothing Then Return
        Try
            TryUpdateAllPartsListsOnDraft(dftDoc, logger)
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
