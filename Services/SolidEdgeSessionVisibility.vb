Option Strict Off

Imports System.Runtime.InteropServices
Imports SolidEdgeFramework

''' <summary>Visibilidad y foco de Solid Edge durante generación batch (sin robar primer plano salvo opción UI o DIMLAB).</summary>
Public NotInheritable Class SolidEdgeSessionVisibility

    Private Const SW_RESTORE As Integer = 9

    <DllImport("user32.dll")>
    Private Shared Function SetForegroundWindow(hWnd As IntPtr) As Boolean
    End Function

    <DllImport("user32.dll")>
    Private Shared Function ShowWindow(hWnd As IntPtr, nCmdShow As Integer) As Boolean
    End Function

    Private Sub New()
    End Sub

    Public Shared Function ShouldKeepSolidEdgeVisible(config As JobConfiguration) As Boolean
        If config Is Nothing Then Return False
        If config.KeepSolidEdgeVisible Then Return True
        If config.EnableDrawingViewDimensioningLab OrElse config.RequestedDimLabFromDedicatedButton Then Return True
        Return False
    End Function

    Public Shared Sub ApplyApplicationVisibility(app As Application, config As JobConfiguration, logger As Logger)
        If app Is Nothing Then Return
        Dim visible As Boolean = ShouldKeepSolidEdgeVisible(config)
        Try
            app.Visible = visible
            app.DisplayAlerts = True
        Catch ex As Exception
            If logger IsNot Nothing Then logger.Log("[SE][VISIBLE] fallo al establecer visibilidad: " & ex.Message)
        End Try
        If logger IsNot Nothing Then
            logger.Log("[SE][VISIBLE] " & visible.ToString())
            logger.Log("[SE][DISPLAY_ALERTS] True")
        End If
    End Sub

    ''' <summary>Oculta Solid Edge y devuelve el foco a la ventana del generador tras abrir/activar documentos COM.</summary>
    Public Shared Sub SuppressForegroundIfConfigured(app As Application, config As JobConfiguration, logger As Logger)
        If app Is Nothing OrElse ShouldKeepSolidEdgeVisible(config) Then Return
        Try
            app.Visible = False
        Catch ex As Exception
            If logger IsNot Nothing Then logger.Log("[SE][VISIBLE][SUPPRESS] " & ex.Message)
        End Try
        TryRestoreOwnerForeground(config, logger)
    End Sub

    Public Shared Sub TryRestoreOwnerForeground(config As JobConfiguration, logger As Logger)
        If config Is Nothing Then Return
        Dim hwnd As IntPtr = config.OwnerWindowHandle
        If hwnd = IntPtr.Zero Then Return
        Try
            ShowWindow(hwnd, SW_RESTORE)
            SetForegroundWindow(hwnd)
            If logger IsNot Nothing Then logger.Log("[SE][FOREGROUND] Foco devuelto a la aplicación.")
        Catch ex As Exception
            If logger IsNot Nothing Then logger.Log("[SE][FOREGROUND][WARN] " & ex.Message)
        End Try
    End Sub

End Class
