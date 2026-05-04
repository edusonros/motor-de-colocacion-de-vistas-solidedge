Option Strict On

Imports System.Collections.Generic
Imports System.Globalization
Imports System.IO
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports SolidEdgeFramework
Imports SolidEdgeFrameworkSupport
Imports SolidEdgePart

''' <summary>
''' Creación de <see cref="PMIModelView"/> solo con el documento PAR/PSM como <see cref="Application.ActiveDocument"/>.
''' Con el Draft activo, <c>Activate()</c> en el modelo suele fallar (E_FAIL) y <c>PMIModelViews.Add</c> no debe invocarse en ese contexto.
''' </summary>
Public NotInheritable Class PmiModelViewExperimental

    Public Const TemporaryPmiModelViewName As String = "AUTO_TEMP_PMI_VIEW"

    ''' <summary>Ejemplo SDK: sombreado con aristas visibles.</summary>
    Public Shared ReadOnly SdkPrimaryAddRenderMode As PMIRenderModeConstants =
        PMIRenderModeConstants.sePMIModelViewRenderModeShadedWithVisibleEdges

    ''' <summary>Segundo intento explícito si el primero falla.</summary>
    Public Shared ReadOnly SdkAlternateAddRenderMode As PMIRenderModeConstants =
        PMIRenderModeConstants.sePMIModelViewRenderModeVisibleEdges

    ''' <summary>Asignación opcional de RenderMode tras Add (visible edges).</summary>
    Public Shared ReadOnly SdkAssignRenderModeVisibleEdges As PMIRenderModeConstants =
        PMIRenderModeConstants.sePMIModelViewRenderModeVisibleEdges

    Public Class CreationResult
        Public Property CreatedNew As Boolean
        Public Property ModelView As PMIModelView
        Public Property ResolvedName As String = ""
        Public Property IndexInCollection As Integer = -1
        Public Property RenderModeDescription As String = ""
        Public Property AddSucceeded As Boolean
        Public Property DimensionsAddedOk As Integer
        Public Property DimensionsAddAttempts As Integer
        Public Property DimensionsAddFailed As Integer
        Public Property DocumentWasSavedBefore As Boolean = True
        Public Property DocumentAppearsModifiedAfter As Boolean
        Public Property Messages As New List(Of String)()
    End Class

    Private Sub New()
    End Sub

    ''' <summary>
    ''' Antes de crear el Draft: abrir o localizar el PAR/PSM; si es el documento activo y no hay PMIModelViews pero hay dimensiones PMI, crear vista temporal.
    ''' No usa Draft ni <c>Activate()</c> (el modelo debe ser activo tras <c>Documents.Open</c> o ya ser la ventana activa).
    ''' </summary>
    Public Shared Sub TryPreCreatePMIModelViewBeforeDraft(app As SolidEdgeFramework.Application, modelPath As String, logger As Logger)
        Dim mv As Action(Of String) = Sub(m) If logger IsNot Nothing Then logger.Log("[PMI][PRE-DRAFT] " & m)
        If app Is Nothing OrElse String.IsNullOrWhiteSpace(modelPath) Then
            mv("Omitido: Application o ruta vacía.")
            Return
        End If

        Dim modelDoc As Object = TryFindOpenDocumentByPath(app, modelPath)
        If modelDoc Is Nothing Then
            Try
                modelDoc = app.Documents.Open(modelPath)
            Catch ex As Exception
                mv("Documents.Open falló: " & ex.Message)
                Return
            End Try
        End If

        If modelDoc Is Nothing Then
            mv("No se pudo obtener documento modelo.")
            Return
        End If

        Dim activeObj As Object = Nothing
        Try
            activeObj = app.ActiveDocument
        Catch
        End Try
        If Not Object.ReferenceEquals(activeObj, modelDoc) Then
            mv("Omitido: el documento activo no es el modelo (p. ej. otro documento abierto primero). " &
               "Cierre el Draft u otros documentos o ejecute primero ExperimentalProbeCreatePMIModelViewOnly. " &
               "Active=" & TryGetDocumentFullName(activeObj) & " | modelo=" & TryGetDocumentFullName(modelDoc))
            Return
        End If

        Dim smDoc As SheetMetalDocument = TryCast(modelDoc, SheetMetalDocument)
        Dim partDoc As PartDocument = TryCast(modelDoc, PartDocument)
        If smDoc Is Nothing AndAlso partDoc Is Nothing Then
            mv("Omitido: no es PAR/PSM.")
            Return
        End If

        Dim pmiObj As PMI = Nothing
        Try
            If smDoc IsNot Nothing Then
                smDoc.PMI_ByModelState(pmiObj)
            Else
                partDoc.PMI_ByModelState(pmiObj)
            End If
        Catch ex As Exception
            mv("PMI_ByModelState: " & ex.Message)
            Return
        End Try

        If pmiObj Is Nothing Then
            mv("Omitido: PMI Nothing.")
            Return
        End If

        Dim dimCount As Integer = 0
        Try
            Dim dco As Object = pmiObj.Dimensions
            If dco IsNot Nothing Then
                dimCount = CType(dco, Dimensions).Count
            End If
        Catch
        End Try

        Dim mvCount As Integer = 0
        Try
            Dim mvsO As Object = pmiObj.PMIModelViews
            If mvsO IsNot Nothing Then
                mvCount = CType(mvsO, PMIModelViews).Count
            End If
        Catch
        End Try

        mv("PMI.Dimensions=" & dimCount.ToString(CultureInfo.InvariantCulture) & ", PMIModelViews.Count=" & mvCount.ToString(CultureInfo.InvariantCulture))

        If dimCount <= 0 OrElse mvCount > 0 Then
            mv("Pre-creación omitida (sin dimensiones PMI o ya existe PMIModelView).")
            Return
        End If

        mv("Invocando creación PMIModelView (contexto modelo activo, antes del Draft)...")
        Dim res = TryCreateTemporaryPMIModelViewForModel(modelDoc, logger)
        If res IsNot Nothing AndAlso res.AddSucceeded Then
            mv("Pre-creación OK. PMIModelView listo para asociar en Draft.")
        Else
            mv("Pre-creación no completada (ver [PMI][MODELVIEW]).")
        End If
    End Sub

    ''' <summary>
    ''' Secuencia tipo SDK: <c>PMI_ByModelState</c>, <c>PMIModelViews.Add</c> (ShadedWithVisibleEdges, luego VisibleEdges),
    ''' lectura/asignación Name, RenderMode, <c>SetShowHidePMISections(0)</c>. Sin <c>Apply</c> ni <c>SetViewOrientationToCurrentView</c>.
    ''' Si <see cref="Application.ActiveDocument"/> no es el modelo, no se llama a <c>Activate()</c> (evita E_FAIL) y se devuelve Nothing.
    ''' </summary>
    Public Shared Function TryCreateTemporaryPMIModelViewStronglyTyped(modelDoc As Object, log As Action(Of String)) As PMIModelView
        If log Is Nothing Then
            log = Sub(_msg As String)
                  End Sub
        End If

        log("[PMI][MODELVIEW] === model-only: FASE Add (SDK) ===")
        log("TypeName(modelDoc)=" & If(modelDoc Is Nothing, "(Nothing)", TypeName(modelDoc)))

        If modelDoc Is Nothing Then
            log("ERROR: modelDoc es Nothing.")
            log("Add ejecutado=False")
            Return Nothing
        End If

        Dim sed As SolidEdgeDocument = TryCast(modelDoc, SolidEdgeDocument)
        If sed Is Nothing Then
            log("ERROR: modelDoc no castea a SolidEdgeDocument.")
            log("Add ejecutado=False")
            Return Nothing
        End If

        Dim modelFull As String = TryGetDocumentFullName(modelDoc)
        log("modelDoc.FullName=" & modelFull)

        Dim app As SolidEdgeFramework.Application = Nothing
        Try
            app = sed.Application
        Catch ex As Exception
            LogFailure(log, ex, "obtener Application")
            log("Add ejecutado=False")
            Return Nothing
        End Try

        Dim activeObj As Object = Nothing
        Try
            activeObj = app.ActiveDocument
        Catch ex As Exception
            log("WARN ActiveDocument: " & ex.Message)
        End Try

        Dim activeBefore As String = TryGetDocumentFullName(activeObj)
        log("documento activo antes=" & activeBefore)

        Dim sameRef As Boolean = Object.ReferenceEquals(activeObj, modelDoc)
        log("modelo es documento activo (ReferenceEquals)=" & sameRef.ToString())

        If Not sameRef Then
            log("BLOQUEADO: el documento activo no es el modelo (p. ej. Draft u otro documento). " &
                "No se llama SolidEdgeDocument.Activate() aquí (suele fallar con E_FAIL 0x80004005 con Draft activo). " &
                "Use ExperimentalProbeCreatePMIModelViewOnly o pre-creación antes del Draft con el modelo como único/activo.")
            log("documento activo después (sin Activate; operación cancelada)=" & TryGetDocumentFullName(activeObj))
            log("Add ejecutado=False")
            Return Nothing
        End If

        log("documento activo después (ya era el modelo; sin Activate)=" & TryGetDocumentFullName(activeObj))

        Dim pmiObj As PMI = Nothing
        Dim phase As String = "PMI_ByModelState"
        Try
            pmiObj = ObtainPmiByModelState(modelDoc, log)
        Catch ex As Exception
            LogFailure(log, ex, phase)
            log("PMI obtenido=False")
            log("Add ejecutado=False")
            Return Nothing
        End Try

        log("PMI obtenido=" & (pmiObj IsNot Nothing).ToString())
        If pmiObj Is Nothing Then
            log("ERROR: PMI Nothing.")
            log("PMIModelViews obtenido=False")
            log("Add ejecutado=False")
            Return Nothing
        End If
        log("TypeName(PMI)=" & TypeName(pmiObj))

        Dim mvsObj As Object = Nothing
        phase = "PMI.PMIModelViews"
        Try
            mvsObj = pmiObj.PMIModelViews
        Catch ex As Exception
            LogFailure(log, ex, phase)
            log("PMIModelViews obtenido=False")
            log("Add ejecutado=False")
            Return Nothing
        End Try

        log("PMIModelViews obtenido=" & (mvsObj IsNot Nothing).ToString())
        If mvsObj Is Nothing Then
            log("Add ejecutado=False")
            Return Nothing
        End If

        Dim mvs As PMIModelViews = Nothing
        Try
            mvs = CType(mvsObj, PMIModelViews)
        Catch ex As Exception
            LogFailure(log, ex, "CType PMIModelViews")
            log("Add ejecutado=False")
            Return Nothing
        End Try

        log("TypeName(PMIModelViews)=" & TypeName(mvs))

        Dim cntBefore As Integer = 0
        phase = "PMIModelViews.Count (antes)"
        Try
            cntBefore = mvs.Count
        Catch ex As Exception
            LogFailure(log, ex, phase)
            log("Add ejecutado=False")
            Return Nothing
        End Try
        log("PMIModelViews.Count antes=" & cntBefore.ToString(CultureInfo.InvariantCulture))

        If cntBefore > 0 Then
            log("Add omitido: Count>0.")
            log("Add ejecutado=False")
            Return Nothing
        End If

        Dim newMv As PMIModelView = Nothing
        phase = "PMIModelViews.Add (sePMIModelViewRenderModeShadedWithVisibleEdges)"
        Try
            newMv = mvs.Add(TemporaryPmiModelViewName, SdkPrimaryAddRenderMode)
        Catch ex As Exception
            LogFailure(log, ex, phase)
            newMv = Nothing
        End Try

        If newMv Is Nothing Then
            phase = "PMIModelViews.Add (sePMIModelViewRenderModeVisibleEdges)"
            Try
                newMv = mvs.Add(TemporaryPmiModelViewName, SdkAlternateAddRenderMode)
            Catch ex2 As Exception
                LogFailure(log, ex2, phase)
                log("Add ejecutado=False")
                Return Nothing
            End Try
        End If

        If newMv Is Nothing Then
            log("Add ejecutado=False")
            Return Nothing
        End If

        log("Add ejecutado=True")

        Dim cntAfter As Integer = 0
        phase = "PMIModelViews.Count (después)"
        Try
            cntAfter = mvs.Count
        Catch ex As Exception
            LogFailure(log, ex, phase)
        End Try
        log("PMIModelViews.Count después=" & cntAfter.ToString(CultureInfo.InvariantCulture))

        phase = "Item(Count)"
        Try
            Dim verify As PMIModelView = CType(mvs.Item(cntAfter), PMIModelView)
            log("Item(Count) OK, TypeName=" & TypeName(verify))
        Catch ex As Exception
            LogFailure(log, ex, phase)
        End Try

        phase = "leer Name"
        Try
            log("objPMIMV.Name (leído)=" & newMv.Name)
        Catch ex As Exception
            LogFailure(log, ex, phase)
        End Try

        phase = "leer RenderMode"
        Try
            log("objPMIMV.RenderMode (leído)=" & newMv.RenderMode.ToString())
        Catch ex As Exception
            LogFailure(log, ex, phase)
        End Try

        phase = "asignar Name"
        Try
            newMv.Name = TemporaryPmiModelViewName
            log("Name asignado=" & TemporaryPmiModelViewName)
        Catch ex As Exception
            LogFailure(log, ex, phase)
        End Try

        phase = "asignar RenderMode VisibleEdges"
        Try
            newMv.RenderMode = SdkAssignRenderModeVisibleEdges
            log("RenderMode asignado=" & SdkAssignRenderModeVisibleEdges.ToString())
        Catch ex As Exception
            LogFailure(log, ex, phase)
        End Try

        phase = "SetShowHidePMISections(0)"
        Try
            newMv.SetShowHidePMISections(0)
            log("SetShowHidePMISections(0) OK.")
        Catch ex As Exception
            LogFailure(log, ex, phase)
        End Try

        log("[PMI][MODELVIEW] === Fin FASE Add (sin Apply / SetViewOrientationToCurrentView) ===")
        Return newMv
    End Function

    ''' <summary>
    ''' Tras Add OK: dimensiones PMI al PMIModelView. Sin <c>Apply</c> ni <c>SetViewOrientationToCurrentView</c> (reintegrar en Draft si aplica).
    ''' </summary>
    Public Shared Function TryCreateTemporaryPMIModelViewForModel(modelDoc As Object, logger As Logger) As CreationResult
        Dim r As New CreationResult()
        Dim mv As Action(Of String) = Sub(m) If logger IsNot Nothing Then logger.Log("[PMI][MODELVIEW] " & m)

        r.DocumentWasSavedBefore = True

        Dim newMv As PMIModelView = TryCreateTemporaryPMIModelViewStronglyTyped(modelDoc, mv)
        If newMv Is Nothing Then
            r.Messages.Add("TryCreateTemporaryPMIModelViewStronglyTyped devolvió Nothing.")
            Return r
        End If

        r.CreatedNew = True
        r.AddSucceeded = True
        r.ModelView = newMv

        Try
            r.ResolvedName = newMv.Name
        Catch
            r.ResolvedName = TemporaryPmiModelViewName
        End Try

        Try
            r.RenderModeDescription = newMv.RenderMode.ToString()
        Catch
        End Try

        Try
            Dim pmiIdx As PMI = ObtainPmiByModelState(modelDoc, mv)
            If pmiIdx IsNot Nothing Then
                Dim mvsIdxObj As Object = pmiIdx.PMIModelViews
                If mvsIdxObj IsNot Nothing Then
                    r.IndexInCollection = CType(mvsIdxObj, PMIModelViews).Count
                End If
            End If
        Catch
            r.IndexInCollection = -1
        End Try

        Dim pmiForDims As PMI = Nothing
        Try
            pmiForDims = ObtainPmiByModelState(modelDoc, mv)
        Catch ex As Exception
            mv("PMI para dimensiones: " & ex.Message)
        End Try

        If pmiForDims IsNot Nothing Then
            Dim dimsCol As Dimensions = Nothing
            Try
                Dim dimsObj As Object = pmiForDims.Dimensions
                If dimsObj IsNot Nothing Then dimsCol = CType(dimsObj, Dimensions)
            Catch ex As Exception
                mv("PMI.Dimensions no accesible: " & ex.Message)
            End Try

            If dimsCol IsNot Nothing Then
                Dim nDim As Integer = 0
                Try
                    nDim = dimsCol.Count
                Catch
                    nDim = 0
                End Try
                r.DimensionsAddAttempts = nDim
                mv("PMI.Dimensions total (modelo)=" & nDim.ToString(CultureInfo.InvariantCulture))

                For i As Integer = 1 To nDim
                    Dim d As Dimension = Nothing
                    Try
                        d = CType(dimsCol.Item(i), Dimension)
                    Catch
                        d = Nothing
                    End Try
                    If d Is Nothing Then Continue For
                    Try
                        newMv.AddDimAnnotOrSectionToView(d)
                        r.DimensionsAddedOk += 1
                    Catch ex As Exception
                        r.DimensionsAddFailed += 1
                        mv("AddDimAnnotOrSectionToView ítem " & i.ToString(CultureInfo.InvariantCulture) & " ERROR: " & ex.Message)
                        Dim cex = TryCast(ex, COMException)
                        If cex IsNot Nothing Then
                            mv("HRESULT=0x" & cex.ErrorCode.ToString("X8", CultureInfo.InvariantCulture))
                        End If
                    End Try
                Next

                mv("Dimensiones: OK=" & r.DimensionsAddedOk.ToString(CultureInfo.InvariantCulture) &
                   ", fallidas=" & r.DimensionsAddFailed.ToString(CultureInfo.InvariantCulture) &
                   ", total=" & nDim.ToString(CultureInfo.InvariantCulture))
            End If
        End If

        r.DocumentAppearsModifiedAfter = r.CreatedNew AndAlso r.AddSucceeded
        If r.DocumentAppearsModifiedAfter Then
            mv("Modelo posiblemente modificado en memoria (sin Save automático).")
        End If

        Return r
    End Function

    Private Shared Function ObtainPmiByModelState(modelDoc As Object, log As Action(Of String)) As PMI
        Dim pmiObj As PMI = Nothing
        Dim smDoc As SheetMetalDocument = TryCast(modelDoc, SheetMetalDocument)
        Dim partDoc As PartDocument = TryCast(modelDoc, PartDocument)
        If smDoc IsNot Nothing Then
            smDoc.PMI_ByModelState(pmiObj)
            log("PMI_ByModelState (SheetMetalDocument).")
        ElseIf partDoc IsNot Nothing Then
            partDoc.PMI_ByModelState(pmiObj)
            log("PMI_ByModelState (PartDocument).")
        Else
            Throw New InvalidOperationException("modelDoc no es SheetMetalDocument ni PartDocument.")
        End If
        Return pmiObj
    End Function

    Private Shared Function TryGetDocumentFullName(doc As Object) As String
        If doc Is Nothing Then Return "(Nothing)"
        Dim sed As SolidEdgeDocument = TryCast(doc, SolidEdgeDocument)
        If sed Is Nothing Then Return "(no SolidEdgeDocument)"
        Try
            Return CStr(sed.FullName)
        Catch ex As Exception
            Return "(FullName error: " & ex.Message & ")"
        End Try
    End Function

    Private Shared Sub LogFailure(log As Action(Of String), ex As Exception, phase As String)
        Dim msg As String = "ERROR [" & phase & "]: " & ex.GetType().FullName & " — " & ex.Message
        Dim cex = TryCast(ex, COMException)
        If cex IsNot Nothing Then
            msg &= " | HRESULT=0x" & cex.ErrorCode.ToString("X8", CultureInfo.InvariantCulture)
        End If
        log(msg)
    End Sub

    ''' <summary>
    ''' Sincroniza la definición del PMIModelView con la vista 3D actual del modelo (API típica PMI) y aplica.
    ''' Usar con el documento modelo activo antes de generar el Draft si la PMI se definió en otra orientación.
    ''' </summary>
    Public Shared Sub TryPMIModelViewSetViewOrientationToCurrentViewApply(pmv As PMIModelView, log As Action(Of String))
        If log Is Nothing Then
            log = Sub(_s As String)
                  End Sub
        End If
        If pmv Is Nothing Then
            log("[PMI][MODELVIEW][ORIENT] PMIModelView Nothing.")
            Return
        End If
        Dim t As Type = pmv.GetType()
        Dim mi1 As MethodInfo = t.GetMethod("SetViewOrientationToCurrentView", BindingFlags.Public Or BindingFlags.Instance)
        If mi1 Is Nothing Then
            log("[PMI][MODELVIEW][ORIENT] SetViewOrientationToCurrentView no encontrado (reflexión).")
            Return
        End If
        Try
            mi1.Invoke(pmv, Nothing)
            log("[PMI][MODELVIEW][ORIENT] SetViewOrientationToCurrentView: OK")
        Catch ex As Exception
            log("[PMI][MODELVIEW][ORIENT] SetViewOrientationToCurrentView: " & ex.GetType().Name & " — " & ex.Message)
            Return
        End Try
        Dim mi2 As MethodInfo = t.GetMethod("Apply", BindingFlags.Public Or BindingFlags.Instance)
        If mi2 Is Nothing Then
            log("[PMI][MODELVIEW][ORIENT] Apply no encontrado.")
            Return
        End If
        Try
            mi2.Invoke(pmv, Nothing)
            log("[PMI][MODELVIEW][ORIENT] Apply: OK")
        Catch ex As Exception
            log("[PMI][MODELVIEW][ORIENT] Apply: " & ex.GetType().Name & " — " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Abre o localiza el PAR/PSM, exige documento activo, localiza PMIModelView por nombre y ejecuta SetViewOrientationToCurrentView+Apply.
    ''' Llamar antes de CreateDraft si la opción ExperimentalPmiSyncPMIModelViewOrientationBeforeDraft está activa.
    ''' </summary>
    Public Shared Sub TrySyncPMIModelViewOrientationInOpenModel(
        app As SolidEdgeFramework.Application,
        modelPath As String,
        pmiViewName As String,
        logger As Logger)

        Dim mv As Action(Of String) = Sub(m) If logger IsNot Nothing Then logger.Log("[PMI][PRE-ORIENT] " & m)
        If app Is Nothing OrElse String.IsNullOrWhiteSpace(modelPath) OrElse String.IsNullOrWhiteSpace(pmiViewName) Then
            mv("Omitido: parámetros vacíos.")
            Return
        End If

        Dim modelDoc As Object = TryFindOpenDocumentByPath(app, modelPath)
        If modelDoc Is Nothing Then
            Try
                modelDoc = app.Documents.Open(modelPath)
            Catch ex As Exception
                mv("Documents.Open: " & ex.Message)
                Return
            End Try
        End If
        If modelDoc Is Nothing Then
            mv("Sin documento modelo.")
            Return
        End If

        Dim activeObj As Object = Nothing
        Try
            activeObj = app.ActiveDocument
        Catch
        End Try
        If Not Object.ReferenceEquals(activeObj, modelDoc) Then
            mv("Omitido: el documento activo no es el modelo (active=" & TryGetDocumentFullName(activeObj) & ").")
            Return
        End If

        Dim pmiObj As PMI = Nothing
        Try
            Dim smDoc As SheetMetalDocument = TryCast(modelDoc, SheetMetalDocument)
            Dim partDoc As PartDocument = TryCast(modelDoc, PartDocument)
            If smDoc IsNot Nothing Then
                smDoc.PMI_ByModelState(pmiObj)
            ElseIf partDoc IsNot Nothing Then
                partDoc.PMI_ByModelState(pmiObj)
            Else
                mv("No es PAR/PSM.")
                Return
            End If
        Catch ex As Exception
            mv("PMI_ByModelState: " & ex.Message)
            Return
        End Try
        If pmiObj Is Nothing Then
            mv("PMI Nothing.")
            Return
        End If

        Dim pmv As PMIModelView = TryGetPMIModelViewByName(pmiObj, pmiViewName, mv)
        If pmv Is Nothing Then
            mv("PMIModelView """ & pmiViewName & """ no encontrado.")
            Return
        End If

        mv("Sincronizando orientación PMIModelView con vista 3D actual…")
        TryPMIModelViewSetViewOrientationToCurrentViewApply(pmv, mv)
    End Sub

    ''' <summary>Busca un PMIModelView por nombre (p. ej. <see cref="TemporaryPmiModelViewName"/>).</summary>
    Public Shared Function FindPMIModelViewByName(pmi As PMI, name As String) As PMIModelView
        Return TryGetPMIModelViewByName(pmi, name, Nothing)
    End Function

    Private Shared Function TryGetPMIModelViewByName(pmi As PMI, name As String, log As Action(Of String)) As PMIModelView
        Try
            Dim mvsO As Object = pmi.PMIModelViews
            If mvsO Is Nothing Then Return Nothing
            Dim mvs As PMIModelViews = CType(mvsO, PMIModelViews)
            Dim n As Integer = mvs.Count
            For i As Integer = 1 To n
                Try
                    Dim pmv As PMIModelView = CType(mvs.Item(i), PMIModelView)
                    If pmv Is Nothing Then Continue For
                    Dim nm As String = Nothing
                    Try
                        nm = pmv.Name
                    Catch
                    End Try
                    If String.Equals(nm, name, StringComparison.OrdinalIgnoreCase) Then
                        Return pmv
                    End If
                Catch ex As Exception
                    If log IsNot Nothing Then log("Item(" & i.ToString(CultureInfo.InvariantCulture) & "): " & ex.Message)
                End Try
            Next
        Catch ex As Exception
            If log IsNot Nothing Then log("TryGetPMIModelViewByName: " & ex.Message)
        End Try
        Return Nothing
    End Function

    Private Shared Function TryFindOpenDocumentByPath(app As SolidEdgeFramework.Application, fullPath As String) As Object
        If app Is Nothing OrElse String.IsNullOrWhiteSpace(fullPath) Then Return Nothing
        Dim target As String
        Try
            target = Path.GetFullPath(fullPath)
        Catch
            target = fullPath
        End Try
        Try
            Dim n As Integer = app.Documents.Count
            For i As Integer = 1 To n
                Try
                    Dim doc As Object = app.Documents.Item(i)
                    Dim sedDoc As SolidEdgeDocument = TryCast(doc, SolidEdgeDocument)
                    If sedDoc Is Nothing Then Continue For
                    Dim name As String = Nothing
                    Try
                        name = CStr(sedDoc.FullName)
                    Catch
                        Continue For
                    End Try
                    If String.IsNullOrWhiteSpace(name) Then Continue For
                    Dim full As String
                    Try
                        full = Path.GetFullPath(name)
                    Catch
                        full = name
                    End Try
                    If String.Equals(full, target, StringComparison.OrdinalIgnoreCase) Then
                        Return doc
                    End If
                Catch
                End Try
            Next
        Catch
        End Try
        Return Nothing
    End Function

End Class
