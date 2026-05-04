Option Strict On

Imports System.Collections.Generic
Imports System.Globalization
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports SolidEdgeDraft
Imports SolidEdgeFrameworkSupport

''' <summary>
''' Diagnóstico PMI → Draft: orientación de <see cref="DrawingView"/> vs <see cref="PMIModelView"/>,
''' y pruebas de alineación antes de <see cref="DrawingView.RetrieveDimensions"/>.
''' Referencia: <see href="https://support.industrysoftware.automation.siemens.com/trainings/se/107/api/SolidEdgeDraft~DrawingView~SetViewOrientationFromNamedView.html">SetViewOrientationFromNamedView</see>.
''' </summary>
Public NotInheritable Class PmiProjectionDiagnostics

    Private Sub New()
    End Sub

    ''' <summary>
    ''' Registra orientación/tipo de vista Draft (reflexión donde el interop expone propiedades como Object o métodos con out-params).
    ''' </summary>
    Public Shared Sub LogDrawingViewProjectionState(dv As DrawingView, phase As String, log As Action(Of String))
        If dv Is Nothing OrElse log Is Nothing Then Return
        log("[PMI][PROJ] === " & phase & " ===")

        Dim t As Type = dv.GetType()
        TryInvokeGetterLog(dv, t, "GetViewOrientationStandard", log, "[PMI][PROJ] GetViewOrientationStandard")
        LogViewOrientationSub(dv, log)
        TryInvokeGetterLog(dv, t, "ViewType", log, "[PMI][PROJ] ViewType")
        TryInvokeGetterLog(dv, t, "UsePerspective", log, "[PMI][PROJ] UsePerspective")

        Dim cls = ClassifyFromViewOrientationConstants(dv)
        If Not String.IsNullOrEmpty(cls) Then
            log("[PMI][PROJ] Clasificación heurística: " & cls)
        End If
    End Sub

    ''' <summary>API real: Sub ViewOrientation(vx,vy,vz, lx,ly,lz, ByRef ori) — igual que DimensioningEngine.</summary>
    Private Shared Sub LogViewOrientationSub(dv As DrawingView, log As Action(Of String))
        If log Is Nothing Then Return
        Try
            Dim vx As Double, vy As Double, vz As Double, lx As Double, ly As Double, lz As Double
            Dim ori As SolidEdgeDraft.ViewOrientationConstants = SolidEdgeDraft.ViewOrientationConstants.igFrontView
            dv.ViewOrientation(vx, vy, vz, lx, ly, lz, ori)
            log("[PMI][PROJ] ViewOrientation(Sub): ori=" & ori.ToString() &
                " view=(" & vx.ToString(CultureInfo.InvariantCulture) & "," & vy.ToString(CultureInfo.InvariantCulture) & "," & vz.ToString(CultureInfo.InvariantCulture) & ")" &
                " local=(" & lx.ToString(CultureInfo.InvariantCulture) & "," & ly.ToString(CultureInfo.InvariantCulture) & "," & lz.ToString(CultureInfo.InvariantCulture) & ")")
        Catch ex As Exception
            log("[PMI][PROJ] ViewOrientation(Sub): " & ex.Message)
        End Try
    End Sub

    Private Shared Function ClassifyFromViewOrientationConstants(dv As DrawingView) As String
        Try
            Dim vx As Double, vy As Double, vz As Double, lx As Double, ly As Double, lz As Double
            Dim ori As SolidEdgeDraft.ViewOrientationConstants = SolidEdgeDraft.ViewOrientationConstants.igFrontView
            dv.ViewOrientation(vx, vy, vz, lx, ly, lz, ori)
            If IsIsometricLike(ori) Then
                Return "vista 3D / trimétrica o isométrica (" & ori.ToString() & ") — la PMI puede no proyectarse con RetrieveDimensions si el motor exige vista plana alineada al PMIModelView."
            End If
            If IsOrthographicStandard(ori) Then
                Return "vista ortográfica estándar (" & ori.ToString() & ")."
            End If
            Return "ViewOrientationConstants=" & ori.ToString()
        Catch
            Return ""
        End Try
    End Function

    Private Shared Function IsIsometricLike(voc As SolidEdgeDraft.ViewOrientationConstants) As Boolean
        Select Case voc
            Case SolidEdgeDraft.ViewOrientationConstants.igTopFrontRightView,
                 SolidEdgeDraft.ViewOrientationConstants.igTopFrontLeftView,
                 SolidEdgeDraft.ViewOrientationConstants.igBottomFrontLeftView,
                 SolidEdgeDraft.ViewOrientationConstants.igDimetricTopFrontLeftView
                Return True
            Case Else
                Return False
        End Select
    End Function

    Private Shared Function IsOrthographicStandard(voc As SolidEdgeDraft.ViewOrientationConstants) As Boolean
        Select Case voc
            Case SolidEdgeDraft.ViewOrientationConstants.igFrontView, SolidEdgeDraft.ViewOrientationConstants.igTopView, SolidEdgeDraft.ViewOrientationConstants.igBottomView,
                 SolidEdgeDraft.ViewOrientationConstants.igRightView, SolidEdgeDraft.ViewOrientationConstants.igLeftView, SolidEdgeDraft.ViewOrientationConstants.igBackView
                Return True
            Case Else
                Return False
        End Select
    End Function

    Private Shared Sub TryInvokeGetterLog(target As DrawingView, t As Type, memberName As String, log As Action(Of String), prefix As String)
        Try
            Dim pi As PropertyInfo = t.GetProperty(memberName, BindingFlags.Public Or BindingFlags.Instance)
            If pi IsNot Nothing AndAlso pi.GetGetMethod() IsNot Nothing AndAlso pi.GetIndexParameters().Length = 0 Then
                Dim v As Object = pi.GetValue(target, Nothing)
                log(prefix & " = " & FormatProjectionValue(v))
                Return
            End If
        Catch ex As Exception
            log(prefix & " (propiedad): " & ex.Message)
        End Try

        Try
            Dim mi As MethodInfo = t.GetMethod(memberName, BindingFlags.Public Or BindingFlags.Instance)
            If mi IsNot Nothing AndAlso mi.GetParameters().Length = 0 Then
                Dim v As Object = mi.Invoke(target, Nothing)
                log(prefix & " = " & FormatProjectionValue(v))
            End If
        Catch ex As Exception
            log(prefix & " (método): " & ex.Message)
        End Try
    End Sub

    Private Shared Function FormatProjectionValue(v As Object) As String
        If v Is Nothing Then Return "(Nothing)"
        Return v.ToString() & " (" & v.GetType().Name & ")"
    End Function

    ''' <summary>
    ''' Métodos API en el tipo runtime útiles para orientación PMI (reflexión).
    ''' </summary>
    Public Shared Sub LogPMIModelViewOrientationApi(pmv As PMIModelView, log As Action(Of String))
        If log Is Nothing Then Return
        If pmv Is Nothing Then
            log("[PMI][PROJ][PMIModelView] (Nothing) — sin objeto vista PMI.")
            Return
        End If
        Dim t As Type = pmv.GetType()
        log("[PMI][PROJ][PMIModelView] RuntimeType=" & t.FullName)
        Try
            log("[PMI][PROJ][PMIModelView] Name=" & pmv.Name)
        Catch ex As Exception
            log("[PMI][PROJ][PMIModelView] Name: " & ex.Message)
        End Try
        Try
            Dim names As New List(Of String)()
            For Each mi As MethodInfo In t.GetMethods(BindingFlags.Public Or BindingFlags.Instance)
                If mi.Name.IndexOf("Orientation", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                   (mi.Name.IndexOf("View", StringComparison.OrdinalIgnoreCase) >= 0 AndAlso mi.Name.IndexOf("PMI", StringComparison.OrdinalIgnoreCase) >= 0) OrElse
                   String.Equals(mi.Name, "Apply", StringComparison.Ordinal) Then
                    names.Add(mi.Name)
                End If
            Next
            names.Sort(StringComparer.Ordinal)
            log("[PMI][PROJ][PMIModelView] Métodos relevantes (Orientation/View+PMI/Apply): " &
                If(names.Count = 0, "(ninguno listado)", String.Join(", ", names.ToArray())))
        Catch ex As Exception
            log("[PMI][PROJ][PMIModelView] Reflexión métodos: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Alinea la vista de plano con la “vista con nombre” del modelo (suele coincidir con el nombre del PMIModelView).
    ''' UsePerspective=False favorece proyección ortográfica para cotas.
    ''' </summary>
    Public Shared Function TrySetViewOrientationFromNamedView(
        dv As DrawingView,
        namedView As String,
        usePerspective As Boolean,
        log As Action(Of String)) As Boolean

        If dv Is Nothing OrElse log Is Nothing Then Return False
        log("[PMI][PROJ] SetViewOrientationFromNamedView(NamedView=""" & namedView & """, UsePerspective=" & usePerspective.ToString() & ") …")
        Try
            dv.SetViewOrientationFromNamedView(namedView, usePerspective)
            log("[PMI][PROJ] SetViewOrientationFromNamedView: OK")
            Return True
        Catch ex As Exception
            log("[PMI][PROJ] SetViewOrientationFromNamedView: ERROR — " & ex.GetType().Name & " — " & ex.Message)
            Dim cex = TryCast(ex, COMException)
            If cex IsNot Nothing Then
                log("[PMI][PROJ] HRESULT=0x" & cex.ErrorCode.ToString("X8", CultureInfo.InvariantCulture))
            End If
            Return False
        End Try
    End Function

    Public Shared Function TrySetViewOrientationStandard(
        dv As DrawingView,
        ori As SolidEdgeDraft.ViewOrientationConstants,
        log As Action(Of String)) As Boolean

        If dv Is Nothing OrElse log Is Nothing Then Return False
        log("[PMI][PROJ] SetViewOrientationStandard(" & ori.ToString() & ") …")
        Try
            dv.SetViewOrientationStandard(ori)
            log("[PMI][PROJ] SetViewOrientationStandard: OK")
            Return True
        Catch ex As Exception
            log("[PMI][PROJ] SetViewOrientationStandard: ERROR — " & ex.Message)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Mensaje de cierre cuando E_ABORT persiste con binding OK (orientación / geometría / límite API).
    ''' </summary>
    Public Shared Sub LogProjectionEAbortConclusion(log As Action(Of String))
        If log Is Nothing Then Return
        log("[PMI][PROJ] CONCLUSIÓN: E_ABORT con PMIModelView enlazado suele indicar proyección no resoluble (orientación distinta a la PMI, " &
            "vista isométrica, chapa curva, o límite del motor RetrieveDimensions — no solo fallo de binding.")
    End Sub

End Class
