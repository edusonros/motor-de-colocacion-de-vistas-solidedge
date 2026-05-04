Option Strict On

Imports System.Collections.Generic
Imports System.Globalization
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports SolidEdgeDraft
Imports SolidEdgeFrameworkSupport

''' <summary>
''' Enlace DrawingView ↔ nombre de vista PMI.
''' Documentación Siemens (tipo de propiedad en VB): <c>Public Property PMIModelView As String</c> — getter/setter trabajan con el
''' <b>nombre</b> del modelo de vista PMI en el documento de pieza/chapa, no con el objeto COM <see cref="PMIModelView"/>.
''' Referencia: https://support.industrysoftware.automation.siemens.com/trainings/se/107/api/SolidEdgeDraft~DrawingView~PMIModelView.html
''' Asignar un objeto <see cref="PMIModelView"/> provoca error del estilo «conversión de PMIModelView a String no es válida».
''' <see cref="DrawingViews.AddPMIModelView"/> (interop SE 2026): firma usada aquí
''' <c>AddPMIModelView(From As ModelLink, ModelViewName As String, Scale As Double, x As Double, y As Double, IncludePMIDimensions As Boolean, IncludePMIAnnotations As Boolean, LastParam As Integer)</c>
''' (último parámetro según sobrecarga del tlb; ver log reflexión <c>[PMI][BIND][AddPMIModelView]</c>).
''' Referencia: https://support.industrysoftware.automation.siemens.com/trainings/se/107/api/SolidEdgeDraft~DrawingViews~AddPMIModelView.html
''' </summary>
Public NotInheritable Class PmiDrawingViewBinding

    Private Sub New()
    End Sub

    ''' <summary>
    ''' Inspección en tiempo de ejecución del interop cargado (CLR). Los tipos deben coincidir con la documentación Siemens (String).
    ''' </summary>
    Public Shared Sub LogDrawingViewPMIModelViewInterop(drawingView As DrawingView, log As Action(Of String))
        If log Is Nothing Then Return
        If drawingView Is Nothing Then
            log("[PMI][BIND] drawingView Nothing.")
            Return
        End If

        Dim t As Type = drawingView.GetType()
        log("[PMI][BIND] RuntimeType DrawingView = " & t.FullName)

        Dim pi As PropertyInfo = FindPropertyPMIModelView(t)
        If pi Is Nothing Then
            log("[PMI][BIND] PropertyInfo 'PMIModelView' no encontrada (reflexión).")
            Return
        End If

        log("[PMI][BIND] Tipo CLR PropertyType (declarado) = " & pi.PropertyType.FullName)
        log("[PMI][BIND] Tipo propiedad getter DrawingView.PMIModelView (ReturnType) = " &
            If(pi.GetGetMethod() IsNot Nothing, pi.GetGetMethod().ReturnType.FullName, "(sin getter)"))

        Dim sm = pi.GetSetMethod()
        If sm IsNot Nothing AndAlso sm.GetParameters().Length > 0 Then
            log("[PMI][BIND] Tipo propiedad setter DrawingView.PMIModelView (parámetro 0) = " & sm.GetParameters()(0).ParameterType.FullName)
        Else
            log("[PMI][BIND] Setter no disponible o sin parámetros visibles.")
        End If

        log("[PMI][BIND] Conclusión interop: en Solid Edge Draft, PMIModelView es el <b>nombre</b> (String) de la vista PMI en el modelo; no asignar el RCW PMIModelView.")
    End Sub

    Private Shared Function FindPropertyPMIModelView(t As Type) As PropertyInfo
        Dim pi = t.GetProperty("PMIModelView", BindingFlags.Public Or BindingFlags.Instance)
        If pi IsNot Nothing Then Return pi
        For Each iface As Type In t.GetInterfaces()
            pi = iface.GetProperty("PMIModelView", BindingFlags.Public Or BindingFlags.Instance)
            If pi IsNot Nothing Then Return pi
        Next
        Return Nothing
    End Function

    ''' <summary>
    ''' RUTA B (recomendada por SDK): <c>DrawingView.PMIModelView = nombre</c> (String).
    ''' </summary>
    Public Shared Function TryAssignPMIModelViewByName(drawingView As DrawingView, modelViewName As String, log As Action(Of String)) As Boolean
        If drawingView Is Nothing OrElse log Is Nothing Then Return False
        log("[PMI][BIND] Valor candidato a asignar (nombre) = " & If(modelViewName, "(Nothing)"))
        Try
            drawingView.PMIModelView = modelViewName
            log("[PMI][BIND] RUTA B — Asignación por nombre (String): OK")
            Return True
        Catch ex As Exception
            log("[PMI][BIND] RUTA B — Asignación por nombre (String): ERROR — " & ex.GetType().FullName & " — " & ex.Message)
            Dim cex = TryCast(ex, COMException)
            If cex IsNot Nothing Then
                log("[PMI][BIND] HRESULT=0x" & cex.ErrorCode.ToString("X8", CultureInfo.InvariantCulture))
            End If
            Return False
        End Try
    End Function

    ''' <summary>
    ''' RUTA A (diagnóstico): intento de asignar el objeto <see cref="PMIModelView"/>; suele fallar si el setter es String.
    ''' </summary>
    Public Shared Sub TryAssignPMIModelViewByObjectDiagnostic(drawingView As DrawingView, modelView As PMIModelView, log As Action(Of String))
        If drawingView Is Nothing OrElse log Is Nothing Then Return
        log("[PMI][BIND] Tipo del objeto PMIModelView creado = " & If(modelView Is Nothing, "(Nothing)", modelView.GetType().FullName))
        If modelView Is Nothing Then Return
        Dim pi As PropertyInfo = FindPropertyPMIModelView(drawingView.GetType())
        If pi Is Nothing Then
            log("[PMI][BIND] RUTA A — PropertyInfo no encontrada.")
            Return
        End If
        Try
            pi.SetValue(drawingView, modelView, Nothing)
            log("[PMI][BIND] RUTA A — PropertyInfo.SetValue(objeto PMIModelView): OK (inesperado si setter es String).")
        Catch ex As Exception
            log("[PMI][BIND] RUTA A — PropertyInfo.SetValue(objeto PMIModelView): ERROR (esperado si setter es String) — " &
                ex.GetType().FullName & " — " & ex.Message)
        End Try
    End Sub

    Public Shared Sub LogReadBackPMIModelView(drawingView As DrawingView, log As Action(Of String))
        If drawingView Is Nothing OrElse log Is Nothing Then Return
        Try
            Dim v As String = drawingView.PMIModelView
            Dim tname As String = If(v Is Nothing, "(Nothing)", v.GetType().FullName)
            log("[PMI][BIND] Lectura posterior DrawingView.PMIModelView = """ & If(v, "(Nothing)") & """ | Tipo = " & tname)
        Catch ex As Exception
            log("[PMI][BIND] Lectura posterior DrawingView.PMIModelView: ERROR — " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Proyección PMI en el Draft: nueva vista con <see cref="DrawingViews.AddPMIModelView"/> (mismo <see cref="ModelLink"/> que la vista de referencia).
    ''' Tras crear la vista, el llamador debe asignar <c>IncludePMIDimensions</c>/<c>IncludePMIAnnotations</c> y <c>Update()</c>.
    ''' </summary>
    Public Shared Function AddPMIModelViewForPmiProjection(
        sheet As Sheet,
        referenceView As DrawingView,
        modelViewName As String,
        log As Action(Of String),
        Optional sheetDimensionCountBefore As Integer = -1) As DrawingView

        If sheet Is Nothing OrElse referenceView Is Nothing OrElse log Is Nothing Then Return Nothing

        Dim link As ModelLink = Nothing
        Try
            Dim linkObj As Object = referenceView.ModelLink
            link = TryCast(linkObj, ModelLink)
        Catch ex As Exception
            log("[PMI][AddPMIModelView] ModelLink: " & ex.Message)
            Return Nothing
        End Try
        If link Is Nothing Then
            log("[PMI][AddPMIModelView] ModelLink Nothing.")
            Return Nothing
        End If

        Dim dvs As DrawingViews = Nothing
        Try
            dvs = sheet.DrawingViews
        Catch ex As Exception
            log("[PMI][AddPMIModelView] DrawingViews: " & ex.Message)
            Return Nothing
        End Try
        If dvs Is Nothing Then
            log("[PMI][AddPMIModelView] DrawingViews Nothing.")
            Return Nothing
        End If

        LogAddPMIModelViewMethodSignatures(dvs, log)

        ' Scale en el interop COM puede exponerse como invocación dinámica; Option Strict On prohíbe leer referenceView.Scale directamente.
        Dim scale As Double = TryGetDoublePropertyViaReflection(referenceView, "Scale", 1.0R)

        Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
        Dim xPlace As Double = 0.1R
        Dim yPlace As Double = 0.1R
        Try
            referenceView.Range(x1, y1, x2, y2)
            Dim maxX = Math.Max(x1, x2)
            Dim minY = Math.Min(y1, y2)
            Dim maxY = Math.Max(y1, y2)
            xPlace = maxX + 0.05R
            yPlace = (minY + maxY) / 2.0R
        Catch
        End Try

        log("[PMI][AddPMIModelView] Firma SDK: AddPMIModelView(From, ModelViewName, Scale, x, y, …); Include* se aplican en la vista tras crear.")
        log("[PMI][AddPMIModelView] Parámetros: ModelViewName=""" & modelViewName & """, Scale=" & scale.ToString(CultureInfo.InvariantCulture) &
            ", x=" & xPlace.ToString(CultureInfo.InvariantCulture) & ", y=" & yPlace.ToString(CultureInfo.InvariantCulture))

        Dim nv As DrawingView = Nothing
        Try
            ' Incluir en Add solo si la sobrecarga lo exige; la proyección PMI se confirma con IncludePMI* + Update en el llamador.
            nv = dvs.AddPMIModelView(link, modelViewName, scale, xPlace, yPlace, False, False, 0)
            If nv IsNot Nothing Then
                log("[PMI][AddPMIModelView] Vista creada OK. TypeName=" & nv.GetType().FullName)
                Try
                    log("[PMI][AddPMIModelView] Nueva vista PMIModelView (String)=" & nv.PMIModelView)
                Catch ex As Exception
                    log("[PMI][AddPMIModelView] Leer PMIModelView en nueva vista: " & ex.Message)
                End Try
            Else
                log("[PMI][AddPMIModelView] AddPMIModelView devolvió Nothing.")
            End If
        Catch ex As Exception
            log("[PMI][AddPMIModelView] ERROR — " & ex.GetType().FullName & " — " & ex.Message)
            Dim cex = TryCast(ex, COMException)
            If cex IsNot Nothing Then
                log("[PMI][AddPMIModelView] HRESULT=0x" & cex.ErrorCode.ToString("X8", CultureInfo.InvariantCulture))
            End If
            nv = Nothing
        End Try

        If sheetDimensionCountBefore >= 0 AndAlso nv IsNot Nothing Then
            Try
                Dim dimsObj As Object = sheet.Dimensions
                If dimsObj IsNot Nothing Then
                    Dim cnt As Integer = CType(dimsObj, Dimensions).Count
                    Dim delta As Integer = cnt - sheetDimensionCountBefore
                    log("[PMI][AddPMIModelView] Sheet.Dimensions: antes=" & sheetDimensionCountBefore.ToString(CultureInfo.InvariantCulture) &
                        ", después=" & cnt.ToString(CultureInfo.InvariantCulture) & ", Δ=" & delta.ToString(CultureInfo.InvariantCulture) &
                        " (tras crear vista PMI).")
                End If
            Catch ex As Exception
                log("[PMI][AddPMIModelView] Sheet.Dimensions (post): " & ex.Message)
            End Try
        End If

        Return nv
    End Function

    ''' <summary>Lee una propiedad numérica de un RCW COM sin late binding (BC30574 con Option Strict On).</summary>
    Private Shared Function TryGetDoublePropertyViaReflection(target As Object, propertyName As String, defaultValue As Double) As Double
        If target Is Nothing Then Return defaultValue
        Try
            Dim t As Type = target.GetType()
            Dim pi As PropertyInfo = t.GetProperty(propertyName, BindingFlags.Public Or BindingFlags.Instance)
            If pi Is Nothing Then Return defaultValue
            Dim raw As Object = pi.GetValue(target, Nothing)
            If raw Is Nothing Then Return defaultValue
            Return System.Convert.ToDouble(raw)
        Catch
            Return defaultValue
        End Try
    End Function

    Private Shared Sub LogAddPMIModelViewMethodSignatures(dvs As DrawingViews, log As Action(Of String))
        Try
            Dim t As Type = dvs.GetType()
            Dim count As Integer = 0
            For Each mi As MethodInfo In t.GetMethods()
                If mi.Name <> "AddPMIModelView" Then Continue For
                count += 1
                Dim ps = mi.GetParameters()
                Dim parts As New List(Of String)()
                For Each p As ParameterInfo In ps
                    parts.Add(p.ParameterType.Name & " " & p.Name)
                Next
                log("[PMI][AddPMIModelView]   " & mi.Name & "(" & String.Join(", ", parts.ToArray()) & ")")
            Next
            log("[PMI][AddPMIModelView] Sobrecargas AddPMIModelView (reflexión): " & count.ToString(CultureInfo.InvariantCulture))
        Catch ex As Exception
            log("[PMI][AddPMIModelView] Reflexión métodos: " & ex.Message)
        End Try
    End Sub

End Class
