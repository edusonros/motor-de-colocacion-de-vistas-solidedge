Option Strict Off

Imports System.Globalization
Imports System.IO
Imports System.Runtime.InteropServices
Imports SolidEdgeDraft
Imports SolidEdgeFramework
Imports SolidEdgeFrameworkSupport
Imports SolidEdgePart

''' <summary>
''' Proyección PMI en Draft mediante <see cref="DrawingViews.AddPMIModelView"/> únicamente (sin <c>RetrieveDimensions</c>).
''' </summary>
Public NotInheritable Class PmiRetrievalProbe

    Private Sub New()
    End Sub

    ''' <param name="experimentalCreatePMIModelViewIfMissing">Reservado; no usado en este flujo (compatibilidad con llamadas existentes).</param>
    ''' <param name="experimentalPmiTryAddPMIModelViewView">Reservado; no usado.</param>
    ''' <param name="experimentalPmiProjectionDiagnostics">Reservado; no usado.</param>
    Public Shared Sub RunProbe(
        app As SolidEdgeFramework.Application,
        draft As DraftDocument,
        mainView As DrawingView,
        modelPath As String,
        sourceKind As SourceFileKind,
        logger As Logger,
        Optional experimentalCreatePMIModelViewIfMissing As Boolean = False,
        Optional experimentalPmiTryAddPMIModelViewView As Boolean = False,
        Optional experimentalPmiProjectionDiagnostics As Boolean = False)

        If logger Is Nothing Then Return

        Dim lg As Action(Of String) = Sub(m) logger.Log("[PMI] " & m)
        Dim cfg As Action(Of String) = Sub(m) logger.Log("[PMI][CONFIG] " & m)
        Dim chk As Action(Of String) = Sub(m) logger.Log("[PMI][CHECK] " & m)
        Dim fb As Action(Of String) = Sub(m) logger.Log("[PMI][FALLBACK] " & m)
        Dim pmiErr As Action(Of String) = Sub(m) logger.Log("[PMI][ERR] " & m)
        Dim pmiWarn As Action(Of String) = Sub(m) logger.Log("[PMI][WARN] " & m)
        Dim proj As Action(Of String) = Sub(m) logger.Log("[PMI][PROJECT] " & m)

        Dim openedModelHere As Boolean = False
        Dim modelDoc As Object = Nothing

        Try
            If mainView Is Nothing Then
                pmiErr("Vista principal Nothing; abortando.")
                Return
            End If

            Dim resolvedPath As String = Nothing
            Dim link As ModelLink = Nothing
            Dim linkOk As Boolean = False

            Try
                link = mainView.ModelLink
                linkOk = (link IsNot Nothing)
            Catch ex As Exception
                LogCom(pmiErr, "DrawingView.ModelLink", ex)
                linkOk = False
            End Try

            chk("ModelLink: " & linkOk.ToString())

            If linkOk AndAlso link IsNot Nothing Then
                Try
                    Dim fn As String = Nothing
                    Try
                        fn = CStr(link.FileName)
                    Catch
                    End Try
                    If Not String.IsNullOrWhiteSpace(fn) Then
                        Try
                            resolvedPath = Path.GetFullPath(fn)
                        Catch
                            resolvedPath = fn
                        End Try
                    End If
                Catch ex As Exception
                    pmiWarn("ModelLink.FileName: " & ex.Message)
                End Try

                If modelDoc Is Nothing Then
                    Try
                        modelDoc = link.ModelDocument
                    Catch ex As Exception
                        pmiWarn("ModelLink.ModelDocument: " & ex.Message)
                    End Try
                End If
            End If

            If String.IsNullOrWhiteSpace(resolvedPath) AndAlso Not String.IsNullOrWhiteSpace(modelPath) Then
                Try
                    resolvedPath = Path.GetFullPath(modelPath)
                Catch
                    resolvedPath = modelPath
                End Try
            End If

            chk("Documento origen: " & If(String.IsNullOrWhiteSpace(resolvedPath), "(sin ruta)", resolvedPath))

            If modelDoc Is Nothing AndAlso app IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(resolvedPath) Then
                modelDoc = TryFindOpenDocumentByPath(app, resolvedPath)
            End If

            Dim docOpen As Boolean = (modelDoc IsNot Nothing)
            chk("Modelo abierto: " & docOpen.ToString())

            If Not docOpen AndAlso app IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(resolvedPath) AndAlso File.Exists(resolvedPath) Then
                Try
                    modelDoc = app.Documents.Open(resolvedPath)
                    openedModelHere = (modelDoc IsNot Nothing)
                    docOpen = openedModelHere
                    chk("Modelo abierto por esta prueba: " & openedModelHere.ToString())
                Catch ex As Exception
                    LogCom(pmiErr, "Documents.Open(modelo)", ex)
                    modelDoc = Nothing
                    docOpen = False
                End Try
            End If

            Dim smDoc As SheetMetalDocument = TryCast(modelDoc, SheetMetalDocument)
            Dim partDoc As PartDocument = TryCast(modelDoc, PartDocument)
            Dim docKindLabel As String = "otro"
            If smDoc IsNot Nothing Then
                docKindLabel = "PSM"
            ElseIf partDoc IsNot Nothing Then
                docKindLabel = "PAR"
            ElseIf modelDoc IsNot Nothing Then
                docKindLabel = "otro"
            End If
            chk("Tipo documento: " & docKindLabel & " (solo PAR/PSM cuentan PMI_ByModelState)")

            Dim pmiObj As SolidEdgeFrameworkSupport.PMI = Nothing
            Dim pmiDimCount As Integer? = Nothing
            Dim pmiModelViewCount As Integer? = Nothing

            If modelDoc IsNot Nothing AndAlso (partDoc IsNot Nothing OrElse smDoc IsNot Nothing) Then
                Try
                    If smDoc IsNot Nothing Then
                        smDoc.PMI_ByModelState(pmiObj)
                    Else
                        partDoc.PMI_ByModelState(pmiObj)
                    End If
                Catch ex As Exception
                    LogCom(pmiErr, "PMI_ByModelState", ex)
                    pmiObj = Nothing
                End Try

                If pmiObj Is Nothing Then
                    pmiWarn("PMI_ByModelState no devolvió objeto PMI.")
                Else
                    Try
                        Dim dc As Dimensions = pmiObj.Dimensions
                        If dc IsNot Nothing Then pmiDimCount = dc.Count
                    Catch ex As Exception
                        LogCom(pmiWarn, "PMI.Dimensions.Count", ex)
                    End Try
                    Try
                        Dim mvCol = pmiObj.PMIModelViews
                        If mvCol IsNot Nothing Then pmiModelViewCount = mvCol.Count
                    Catch
                        pmiModelViewCount = Nothing
                    End Try
                End If
            ElseIf modelDoc Is Nothing Then
                pmiWarn("Sin modelo abierto: no se evalúa PMI.")
            Else
                pmiWarn("Documento no PAR/PSM: se omite PMI_ByModelState en esta prueba.")
            End If

            chk("PMI.Dimensions (modelo): " & If(pmiDimCount.HasValue, pmiDimCount.Value.ToString(CultureInfo.InvariantCulture), "(n/d)"))
            If pmiModelViewCount.HasValue Then
                chk("PMI.PMIModelViews.Count: " & pmiModelViewCount.Value.ToString(CultureInfo.InvariantCulture))
            Else
                chk("PMI.PMIModelViews: (no leído o no disponible)")
            End If

            Dim willRunPmiProjection As Boolean =
                draft IsNot Nothing AndAlso
                mainView IsNot Nothing AndAlso
                modelDoc IsNot Nothing AndAlso
                (partDoc IsNot Nothing OrElse smDoc IsNot Nothing) AndAlso
                pmiModelViewCount.HasValue AndAlso
                pmiModelViewCount.Value > 0

            chk("Proyección PMI vía AddPMIModelView: " & willRunPmiProjection.ToString())

            If Not willRunPmiProjection Then
                If pmiDimCount.HasValue AndAlso pmiDimCount.Value > 0 AndAlso
                    (Not pmiModelViewCount.HasValue OrElse pmiModelViewCount.Value = 0) Then
                    LogPmiFallbackToAutoDimensionFuture(fb, pmiModelViewCount, True)
                    fb("Hay PMI.Dimensions pero PMIModelViews.Count=0; no se puede proyectar con AddPMIModelView hasta exista al menos una vista PMI en el modelo.")
                End If
                Return
            End If

            Dim sheet As Sheet = Nothing
            Try
                sheet = draft.ActiveSheet
            Catch ex As Exception
                pmiErr("ActiveSheet — " & ex.Message)
                Return
            End Try
            If sheet Is Nothing Then
                pmiErr("ActiveSheet Nothing.")
                Return
            End If

            Dim sheetDims As Dimensions = Nothing
            Dim countBefore As Integer = 0
            Try
                sheetDims = CType(sheet.Dimensions, Dimensions)
                countBefore = sheetDims.Count
            Catch ex As Exception
                pmiErr("Sheet.Dimensions — " & ex.Message)
                Return
            End Try

            Dim bindLog As Action(Of String) = Sub(m) logger.Log(m)
            Dim pmiViewName As String = ResolvePmiModelViewName(pmiObj, chk)
            proj("Vista PMI usada para AddPMIModelView: """ & pmiViewName & """")

            TryConfigureDraftFilePreferencesForPmi(draft, cfg, pmiWarn)
            PmiDraftContext.PrepareDraftContextBeforeRetrieve(app, draft, sheet, mainView, cfg, pmiWarn)

            Dim newView As DrawingView = PmiDrawingViewBinding.AddPMIModelViewForPmiProjection(
                sheet, mainView, pmiViewName, bindLog, countBefore)

            If newView Is Nothing Then
                LogPmiFallbackToAutoDimensionFuture(fb, pmiModelViewCount, True)
                fb("AddPMIModelView no devolvió vista o falló; ver [PMI][AddPMIModelView].")
                Return
            End If

            TryConfigureDrawingViewPmiFlags(newView, cfg, pmiWarn)
            PmiDraftContext.PrepareDraftContextBeforeRetrieve(app, draft, sheet, newView, cfg, pmiWarn)
            TryUpdateDrawingView(newView, cfg, pmiWarn)

            LogSheetDimensionDeltaAfterProjection(sheetDims, countBefore, proj, pmiWarn)
            proj("Proyección PMI completada sin RetrieveDimensions.")

        Catch ex As Exception
            LogCom(pmiErr, "PmiRetrievalProbe.RunProbe", ex)
            LogPmiFallbackToAutoDimensionFuture(fb, Nothing, True)
        Finally
            If openedModelHere AndAlso modelDoc IsNot Nothing Then
                lg("Nota: modelo abierto solo para la prueba PMI; no se cierra aquí.")
            End If
        End Try
    End Sub

    Private Shared Function ResolvePmiModelViewName(pmi As SolidEdgeFrameworkSupport.PMI, chk As Action(Of String)) As String
        If pmi Is Nothing Then Return PmiModelViewExperimental.TemporaryPmiModelViewName

        Dim pmv = PmiModelViewExperimental.FindPMIModelViewByName(pmi, PmiModelViewExperimental.TemporaryPmiModelViewName)
        If pmv IsNot Nothing Then
            Try
                Dim n As String = pmv.Name
                If Not String.IsNullOrWhiteSpace(n) Then Return n
            Catch
            End Try
        End If

        Try
            Dim mvs = pmi.PMIModelViews
            If mvs IsNot Nothing AndAlso mvs.Count >= 1 Then
                Dim first = CType(mvs.Item(1), PMIModelView)
                If first IsNot Nothing Then
                    Dim n2 As String = first.Name
                    If Not String.IsNullOrWhiteSpace(n2) Then Return n2
                End If
            End If
        Catch ex As Exception
            chk("WARN: resolver nombre PMIModelView: " & ex.Message)
        End Try

        Return PmiModelViewExperimental.TemporaryPmiModelViewName
    End Function

    ''' <summary>Preferencias de documento útiles para PMI; sin <c>DVRetrieveDimensionsOnViewCreation</c>.</summary>
    Private Shared Sub TryConfigureDraftFilePreferencesForPmi(draft As DraftDocument,
                                                              cfg As Action(Of String), warn As Action(Of String))
        Try
            Dim prefs As Object = Nothing
            Try
                prefs = draft.DraftFilePreferences
            Catch
            End Try
            If prefs IsNot Nothing Then
                TrySetPrefBool(prefs, "DVIncludePMIDimensions", True, cfg, warn)
                TrySetPrefBool(prefs, "DVIncludePMIAnnotations", True, cfg, warn)
            End If
        Catch ex As Exception
            warn("DraftFilePreferences: " & ex.Message)
        End Try
    End Sub

    Private Shared Sub TrySetPrefBool(prefs As Object, propName As String, value As Boolean,
                                      cfg As Action(Of String), warn As Action(Of String))
        Try
            CallByName(prefs, propName, CallType.Let, value)
            cfg("DraftFilePreferences." & propName & " = " & value.ToString())
        Catch ex As Exception
            warn("DraftFilePreferences." & propName & " no disponible: " & ex.Message)
        End Try
    End Sub

    Private Shared Sub TryUpdateDrawingView(dv As DrawingView, cfg As Action(Of String), warn As Action(Of String))
        If dv Is Nothing Then Return
        Try
            dv.Update()
            cfg("DrawingView.Update() ejecutado (vista AddPMIModelView).")
        Catch ex As Exception
            warn("DrawingView.Update: " & ex.Message)
        End Try
    End Sub

    Private Shared Sub TryConfigureDrawingViewPmiFlags(dv As DrawingView, cfg As Action(Of String), warn As Action(Of String))
        If dv Is Nothing Then Return
        Try
            dv.IncludePMIDimensions = True
            cfg("DrawingView.IncludePMIDimensions = True")
        Catch ex As Exception
            warn("IncludePMIDimensions: " & ex.Message)
        End Try
        Try
            dv.IncludePMIAnnotations = True
            cfg("DrawingView.IncludePMIAnnotations = True")
        Catch ex As Exception
            warn("IncludePMIAnnotations: " & ex.Message)
        End Try
    End Sub

    Private Shared Sub LogSheetDimensionDeltaAfterProjection(sheetDims As Dimensions, countBefore As Integer,
                                                             proj As Action(Of String), warn As Action(Of String))
        If sheetDims Is Nothing Then Return
        Dim countAfter As Integer = countBefore
        Try
            countAfter = sheetDims.Count
        Catch ex As Exception
            warn("Sheet.Dimensions.Count tras Update: " & ex.Message)
            Return
        End Try
        Dim delta As Integer = countAfter - countBefore
        proj("Sheet.Dimensions: antes=" & countBefore.ToString(CultureInfo.InvariantCulture) &
            ", después=" & countAfter.ToString(CultureInfo.InvariantCulture) &
            ", Δ=" & delta.ToString(CultureInfo.InvariantCulture))
    End Sub

    Private Shared Sub LogPmiFallbackToAutoDimensionFuture(fb As Action(Of String),
                                                           pmiModelViewCount As Integer?,
                                                           projectionFailed As Boolean)
        If Not projectionFailed Then Return
        fb("Proyección PMI vía AddPMIModelView no disponible o falló.")
        fb("Próximo paso previsto: acotado automático propio (p. ej. DimensioningEngine / geometría de vista) cuando se integre en el pipeline.")
        If pmiModelViewCount.HasValue AndAlso pmiModelViewCount.Value = 0 Then
            fb("Contexto: PMIModelViews.Count=0 en el modelo.")
        End If
    End Sub

    Private Shared Sub LogCom(sink As Action(Of String), ctx As String, ex As Exception)
        Dim msg As String = ctx & " — " & ex.GetType().Name & " — " & ex.Message
        Dim cex = TryCast(ex, COMException)
        If cex IsNot Nothing Then
            msg &= " (HRESULT=0x" & cex.ErrorCode.ToString("X8", CultureInfo.InvariantCulture) & ")"
        End If
        sink(msg)
    End Sub

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
                    Dim name As String = Nothing
                    Try
                        name = CStr(doc.FullName)
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
