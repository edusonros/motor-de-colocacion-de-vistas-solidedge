Option Strict Off

Imports System.Collections.Generic
Imports System.Globalization
Imports System.IO
Imports System.Linq
Imports System.Text

''' <summary>
''' Evita volcar [PROPS][INSPECT] repetidamente sobre los mismos .dft de plantilla: caché en disco por ruta + fecha de modificación.
''' </summary>
Friend NotInheritable Class TemplateInspectionCache
    Private Shared ReadOnly SyncRoot As New Object()
    Private Shared ReadOnly LastLoggedTicks As New Dictionary(Of String, Long)(StringComparer.OrdinalIgnoreCase)
    Private Shared _initialInspectionLogged As Boolean = False
    Private Shared ReadOnly PersistLock As New Object()

    Private Sub New()
    End Sub

    Private Shared Function GetPersistCacheFilePath() As String
        Dim root = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Extraer_dft_dxf_flatdxf")
        Return Path.Combine(root, "template_property_inspect_mtime.txt")
    End Function

    Private Shared Function LoadPersistedMtimes() As Dictionary(Of String, Long)
        Dim d As New Dictionary(Of String, Long)(StringComparer.OrdinalIgnoreCase)
        Dim fp = GetPersistCacheFilePath()
        Try
            If Not File.Exists(fp) Then Return d
            For Each line In File.ReadAllLines(fp, Encoding.UTF8)
                If String.IsNullOrWhiteSpace(line) Then Continue For
                Dim idx = line.IndexOf("|"c)
                If idx <= 0 Then Continue For
                Dim p = line.Substring(0, idx).Trim()
                Dim tPart = line.Substring(idx + 1).Trim()
                Dim t As Long = 0
                If p.Length > 0 AndAlso Long.TryParse(tPart, NumberStyles.Integer, CultureInfo.InvariantCulture, t) Then
                    d(p) = t
                End If
            Next
        Catch
        End Try
        Return d
    End Function

    Private Shared Sub SavePersistedMtimes(map As Dictionary(Of String, Long))
        If map Is Nothing Then Return
        Dim fp = GetPersistCacheFilePath()
        Try
            Dim dir = Path.GetDirectoryName(fp)
            If Not String.IsNullOrWhiteSpace(dir) AndAlso Not Directory.Exists(dir) Then Directory.CreateDirectory(dir)
            Dim lines = map.OrderBy(Function(kv) kv.Key, StringComparer.OrdinalIgnoreCase).
                Select(Function(kv) kv.Key & "|" & kv.Value.ToString(CultureInfo.InvariantCulture))
            File.WriteAllLines(fp, lines.ToArray(), Encoding.UTF8)
        Catch
        End Try
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

    ''' <summary>
    ''' Con diagnóstico de plantillas activo: vuelca PropertySets la primera vez (o si el .dft cambió); si no, una línea [CACHE] y no abre el fichero.
    ''' </summary>
    Public Shared Sub LogTemplatePropertyInspectWithDiskCache(templatePath As String, logger As Logger, Optional maxPropsPerSet As Integer = 250)
        If logger Is Nothing Then Return
        If String.IsNullOrWhiteSpace(templatePath) OrElse Not File.Exists(templatePath) Then Return

        Dim full As String = Path.GetFullPath(templatePath)
        Dim ticks As Long = GetTicksSafe(full)
        Dim skipInspect As Boolean = False

        SyncLock PersistLock
            Dim diskMap = LoadPersistedMtimes()
            Dim oldT As Long = 0
            If diskMap.TryGetValue(full, oldT) AndAlso oldT = ticks Then
                skipInspect = True
            End If
        End SyncLock

        If skipInspect Then
            logger.Log("[PROPS][INSPECT][CACHE] Archivo=" & full & " (sin cambios desde última inspección; omitiendo volcado)")
            Return
        End If

        SolidEdgePropertyService.LogAllPropertySetsFromFile(full, logger, maxPropsPerSet)

        SyncLock PersistLock
            Dim diskMap = LoadPersistedMtimes()
            diskMap(full) = ticks
            SavePersistedMtimes(diskMap)
        End SyncLock
    End Sub

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

        LogTemplatePropertyInspectWithDiskCache(full, logger, maxPropsPerSet)

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
