Option Strict Off

Imports System.Globalization
Imports System.IO
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports SolidEdgeFramework
Imports SolidEdgeFrameworkSupport
Imports SolidEdgePart

''' <summary>
''' Secuencia antes de crear el Draft cuando se va a recuperar PMI: modelo abierto, activo, PMI visible y depuración de cotas.
''' No sustituye TryPreCreatePMIModelViewBeforeDraft; se ejecuta adicionalmente si EnablePmiRetrievalProbe está activo.
''' </summary>
Public NotInheritable Class PmiModelPreDraftPrep

    Private Sub New()
    End Sub

    ''' <summary>
    ''' Orden: abrir modelo → activar → PMI (document.PMI si existe, si no PMI_ByModelState) → Show/ShowDimensions → log dimensiones.
    ''' </summary>
    Public Shared Sub PrepareModelBeforeDraftForPmi(app As Application, modelPath As String, logger As Logger)
        If logger Is Nothing OrElse app Is Nothing OrElse String.IsNullOrWhiteSpace(modelPath) Then Return
        Dim log As Action(Of String) = Sub(m) logger.Log("[PMI][PRE-DRAFT] " & m)

        If Not File.Exists(modelPath) Then
            log("WARN: archivo modelo no encontrado: " & modelPath)
            Return
        End If

        Dim modelDoc As Object = FindOpenDocumentByPath(app, modelPath)
        If modelDoc Is Nothing Then
            Try
                modelDoc = app.Documents.Open(modelPath)
                log("Documents.Open(modelo) OK")
            Catch ex As Exception
                LogCom(log, ex, "Documents.Open")
                Return
            End Try
        Else
            log("Modelo ya estaba abierto en sesión.")
        End If

        If modelDoc Is Nothing Then
            log("WARN: documento modelo Nothing.")
            Return
        End If

        Dim sed As SolidEdgeDocument = TryCast(modelDoc, SolidEdgeDocument)
        If sed Is Nothing Then
            log("WARN: no castea a SolidEdgeDocument.")
            Return
        End If

        Try
            sed.Activate()
            log("SolidEdgeDocument.Activate() OK")
        Catch ex As Exception
            LogCom(log, ex, "SolidEdgeDocument.Activate")
        End Try

        Try
            log("Modelo FullName=" & CStr(sed.FullName))
        Catch
        End Try

        Dim smDoc As SheetMetalDocument = TryCast(modelDoc, SheetMetalDocument)
        Dim partDoc As PartDocument = TryCast(modelDoc, PartDocument)
        If smDoc Is Nothing AndAlso partDoc Is Nothing Then
            log("WARN: no es PartDocument ni SheetMetalDocument.")
            Return
        End If

        Dim pmiObj As PMI = Nothing

        ' 1) API tipo partDoc.PMI / smDoc.PMI (según interop; reflexión si hace falta)
        pmiObj = TryGetPmiViaDocumentProperty(modelDoc, log)
        If pmiObj Is Nothing Then
            Try
                If smDoc IsNot Nothing Then
                    smDoc.PMI_ByModelState(pmiObj)
                    log("PMI vía SheetMetalDocument.PMI_ByModelState(out).")
                Else
                    partDoc.PMI_ByModelState(pmiObj)
                    log("PMI vía PartDocument.PMI_ByModelState(out).")
                End If
            Catch ex As Exception
                LogCom(log, ex, "PMI_ByModelState")
                Return
            End Try
        End If

        If pmiObj Is Nothing Then
            log("WARN: PMI Nothing (ni .PMI ni PMI_ByModelState).")
            Return
        End If

        ' 2) Visibilidad: asignación directa (late binding) + fallback CallByName
        TryPmiVisibilityShowDimensionsDirect(pmiObj, log)
        TryPmiVisibilityFlagsFallback(pmiObj, log)
        LogPmiDimensionsDebug(pmiObj, log)
    End Sub

    ''' <summary>
    ''' Equivale a partDoc.PMI / smDoc.PMI cuando el interop lo expone como propiedad.
    ''' </summary>
    Private Shared Function TryGetPmiViaDocumentProperty(modelDoc As Object, log As Action(Of String)) As PMI
        If modelDoc Is Nothing Then Return Nothing
        Try
            Dim t As Type = modelDoc.GetType()
            Dim pi As PropertyInfo = t.GetProperty("PMI", BindingFlags.Public Or BindingFlags.Instance)
            If pi Is Nothing Then
                log("documento.PMI: propiedad no encontrada (reflexión); se usará PMI_ByModelState.")
                Return Nothing
            End If
            Dim o As Object = pi.GetValue(modelDoc, Nothing)
            If o Is Nothing Then
                log("documento.PMI: Nothing.")
                Return Nothing
            End If
            Dim p As PMI = TryCast(o, PMI)
            If p IsNot Nothing Then
                log("PMI obtenido vía documento.PMI (reflexión → PMI).")
                Return p
            End If
            log("WARN: documento.PMI no castea a PMI; se usará PMI_ByModelState.")
        Catch ex As Exception
            LogCom(log, ex, "documento.PMI")
        End Try
        Return Nothing
    End Function

    Private Shared Sub TryPmiVisibilityShowDimensionsDirect(pmi As PMI, log As Action(Of String))
        If pmi Is Nothing Then Return
        Dim late As Object = pmi
        Try
            late.Show = True
            log("PMI.Show = True OK (directo).")
        Catch ex As Exception
            LogCom(log, ex, "PMI.Show")
        End Try
        Try
            late.ShowDimensions = True
            log("PMI.ShowDimensions = True OK (directo).")
        Catch ex As Exception
            LogCom(log, ex, "PMI.ShowDimensions")
        End Try
    End Sub

    Private Shared Sub TryPmiVisibilityFlagsFallback(pmi As PMI, log As Action(Of String))
        If pmi Is Nothing Then Return
        For Each propName In New String() {"Show", "ShowDimensions", "ShowPMI", "ShowDimensionsInPMI"}
            Try
                CallByName(pmi, propName, CallType.Let, True)
                log("PMI." & propName & " = True (CallByName OK)")
            Catch
            End Try
        Next
    End Sub

    Private Shared Sub LogPmiDimensionsDebug(pmi As PMI, log As Action(Of String))
        Try
            Dim dimsObj As Object = Nothing
            Try
                dimsObj = pmi.Dimensions
            Catch ex As Exception
                log("WARN: no se pudo leer PMI.Dimensions: " & ex.Message)
                Return
            End Try
            If dimsObj Is Nothing Then
                log("WARN: PMI.Dimensions Nothing")
                Return
            End If
            Dim dims As Dimensions = Nothing
            Try
                dims = CType(dimsObj, Dimensions)
            Catch ex As Exception
                log("WARN: PMI.Dimensions no castea a Dimensions: " & ex.Message)
                Return
            End Try
            Dim n As Integer = 0
            Try
                n = dims.Count
            Catch ex As Exception
                log("WARN: Dimensions.Count: " & ex.Message)
                Return
            End Try
            log("PMI.Dimensions.Count = " & n.ToString(CultureInfo.InvariantCulture))
            If n = 0 Then
                log("WARN: PMI.Dimensions.Count=0 — no hay cotas PMI en el modelo para recuperar en Draft.")
                Return
            End If
            Dim maxShow As Integer = Math.Min(n, 20)
            For i As Integer = 1 To maxShow
                Try
                    Dim d As Object = dims.Item(i)
                    Dim info As String = TypeName(d)
                    Try
                        info &= " | Value=" & CStr(CallByName(d, "Value", CallType.Get))
                    Catch
                    End Try
                    log("  Dim[" & i.ToString(CultureInfo.InvariantCulture) & "] " & info)
                Catch ex As Exception
                    log("  Dim[" & i.ToString(CultureInfo.InvariantCulture) & "] ERROR: " & ex.Message)
                End Try
            Next
        Catch ex As Exception
            LogCom(log, ex, "LogPmiDimensionsDebug")
        End Try
    End Sub

    Private Shared Sub LogCom(log As Action(Of String), ex As Exception, ctx As String)
        Dim msg As String = ctx & " — " & ex.GetType().Name & " — " & ex.Message
        Dim cex = TryCast(ex, COMException)
        If cex IsNot Nothing Then
            msg &= " | HRESULT=0x" & cex.ErrorCode.ToString("X8", CultureInfo.InvariantCulture)
        End If
        log(msg)
    End Sub

    Private Shared Function FindOpenDocumentByPath(app As Application, fullPath As String) As Object
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
