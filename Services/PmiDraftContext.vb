Option Strict Off

Imports System.Globalization
Imports System.Runtime.InteropServices
Imports SolidEdgeDraft
Imports SolidEdgeFramework

''' <summary>
''' Contexto Draft antes de actualizar la vista PMI (p. ej. <c>DrawingView.Update</c>): documento activo, hoja activa, metadatos de vista.
''' En el interop Draft, <see cref="DrawingView"/> no expone normalmente <c>Activate()</c> sin parámetros; no usar reflexión para ello.
''' </summary>
Public NotInheritable Class PmiDraftContext

    Private Sub New()
    End Sub

    ''' <summary>
    ''' Secuencia mínima tipada: activar DFT, comprobar hoja activa, comprobar ActiveDocument, nombre/tipo de vista, estado de actualización si existe la propiedad.
    ''' </summary>
    Public Shared Sub PrepareDraftContextBeforeRetrieve(
        app As SolidEdgeFramework.Application,
        draft As DraftDocument,
        sheet As Sheet,
        targetView As DrawingView,
        cfg As Action(Of String),
        warn As Action(Of String))

        If draft Is Nothing Then
            warn("[PMI][CTX] DraftDocument Nothing.")
            Return
        End If

        Try
            Dim sed As SolidEdgeDocument = TryCast(draft, SolidEdgeDocument)
            If sed IsNot Nothing Then
                sed.Activate()
                cfg("[PMI][CTX] SolidEdgeDocument.Activate() OK (Draft).")
            End If
        Catch ex As Exception
            LogComWarn(warn, ex, "Draft.Activate")
        End Try

        Try
            Dim cur As Sheet = draft.ActiveSheet
            If cur Is Nothing Then
                warn("[PMI][CTX] draft.ActiveSheet es Nothing.")
            ElseIf sheet IsNot Nothing Then
                cfg("[PMI][CTX] ActiveSheet misma referencia que hoja del probe = " &
                    Object.ReferenceEquals(cur, sheet).ToString())
            Else
                cfg("[PMI][CTX] draft.ActiveSheet OK (sheet probe no pasada).")
            End If
        Catch ex As Exception
            LogComWarn(warn, ex, "draft.ActiveSheet")
        End Try

        If app IsNot Nothing Then
            Try
                Dim ad As Object = app.ActiveDocument
                cfg("[PMI][CTX] Application.ActiveDocument es este Draft (ReferenceEquals) = " &
                    Object.ReferenceEquals(ad, draft).ToString())
            Catch ex As Exception
                LogComWarn(warn, ex, "Application.ActiveDocument")
            End Try
        End If

        If targetView Is Nothing Then
            warn("[PMI][CTX] DrawingView destino Nothing.")
            Return
        End If

        cfg("[PMI][CTX] DrawingView CLR: " & targetView.GetType().FullName)

        Dim late As Object = targetView
        Try
            cfg("[PMI][CTX] DrawingView.Name = " & CStr(late.Name))
        Catch ex As Exception
            warn("[PMI][CTX] DrawingView.Name no disponible: " & ex.Message)
        End Try

        Try
            cfg("[PMI][CTX] DrawingView.ViewsNeedUpdate = " & CStr(late.ViewsNeedUpdate))
        Catch ex As Exception
            warn("[PMI][CTX] DrawingView.ViewsNeedUpdate no expuesto en interop (omitido): " & ex.Message)
        End Try

        cfg("[PMI][CTX] Nota: no se usa reflexión sobre Activate; en Draft, DrawingView no suele tener Sub Activate() sin parámetros.")
    End Sub

    Private Shared Sub LogComWarn(warn As Action(Of String), ex As Exception, ctx As String)
        Dim msg As String = "[PMI][CTX] " & ctx & " — " & ex.GetType().Name & " — " & ex.Message
        Dim cex = TryCast(ex, COMException)
        If cex IsNot Nothing Then
            msg &= " | HRESULT=0x" & cex.ErrorCode.ToString("X8", CultureInfo.InvariantCulture)
        End If
        warn(msg)
    End Sub

End Class
