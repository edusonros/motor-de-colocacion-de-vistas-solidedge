Option Strict Off

Imports System.Collections.Generic
Imports System.IO

''' <summary>
''' Evita reinspeccionar los mismos templates en cada proceso: clave = ruta completa + fecha de modificación.
''' </summary>
Friend NotInheritable Class TemplateInspectionCache
    Private Shared ReadOnly SyncRoot As New Object()
    Private Shared ReadOnly LastLoggedTicks As New Dictionary(Of String, Long)(StringComparer.OrdinalIgnoreCase)
    Private Shared _initialInspectionLogged As Boolean = False

    Private Sub New()
    End Sub

    ''' <summary>True si todos los paths existentes ya están cacheados con el mismo LastWriteUtc.</summary>
    Public Shared Function AllTemplatesReuseSessionCache(templatePaths As IEnumerable(Of String)) As Boolean
        If templatePaths Is Nothing Then Return False
        Dim any As Boolean = False
        SyncLock SyncRoot
            For Each p In templatePaths
                If String.IsNullOrWhiteSpace(p) OrElse Not File.Exists(p) Then Continue For
                any = True
                Dim full As String = Path.GetFullPath(p)
                Dim ticks As Long = GetTicksSafe(full)
                Dim old As Long = 0
                If Not LastLoggedTicks.TryGetValue(full, old) OrElse old <> ticks Then
                    Return False
                End If
            Next
        End SyncLock
        Return any
    End Function

    ''' <summary>Registra propiedades del template con caché por sesión y por mtime.</summary>
    ''' <param name="enableInspection">Solo True en modo diagnóstico de plantillas (equivalente a DebugTemplatesInspection).</param>
    Public Shared Sub EnsureTemplatePropertyLog(templatePath As String, logger As Logger, Optional maxPropsPerSet As Integer = 250, Optional enableInspection As Boolean = False)
        If Not enableInspection Then Return
        If logger Is Nothing Then Return
        If String.IsNullOrWhiteSpace(templatePath) OrElse Not File.Exists(templatePath) Then Return

        Dim full As String = Path.GetFullPath(templatePath)
        Dim ticks As Long = GetTicksSafe(full)

        SyncLock SyncRoot
            Dim old As Long = 0
            If LastLoggedTicks.TryGetValue(full, old) Then
                If old = ticks Then
                    Return
                End If
                logger.Log("[TPL][DEBUG] Template modificado, reanalizando: " & full)
            End If
        End SyncLock

        SolidEdgePropertyService.LogAllPropertySetsFromFile(full, logger, maxPropsPerSet)

        SyncLock SyncRoot
            LastLoggedTicks(full) = ticks
            If Not _initialInspectionLogged Then
                logger.Log("[TPL][DEBUG] Inspección de templates (diagnóstico) completada.")
                _initialInspectionLogged = True
            End If
        End SyncLock
    End Sub

    Private Shared Function GetTicksSafe(fullPath As String) As Long
        Try
            Return File.GetLastWriteTimeUtc(fullPath).Ticks
        Catch
            Return 0
        End Try
    End Function

End Class
