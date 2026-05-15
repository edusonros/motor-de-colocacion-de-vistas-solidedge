Option Strict Off

Imports System.Collections.Generic
Imports System.Globalization
Imports System.IO
Imports System.Text
Imports SolidEdgeDraft

''' <summary>
''' Contexto de diagnósticos de una ejecución del motor.
''' Crea y gestiona los archivos bajo la carpeta fija del repositorio:
'''   docs\logs\
''' con los nombres:
'''   run_log_yyyyMMdd_HHmmss.txt      → log completo (mirror del Logger + snapshot final).
'''   com_errors_yyyyMMdd_HHmmss.txt   → errores COM (mirror dedicado del Logger).
'''   audit_&lt;base&gt;_yyyyMMdd_HHmmss.txt → DftAuditService.ExportAuditFromOpenDocument por cada DFT.
'''   geometry_&lt;base&gt;_yyyyMMdd_HHmmss.txt → DraftGeometryReporter.ExportDraftGeometryLog por cada DFT.
''' </summary>
Public NotInheritable Class RunDiagnosticsContext
    Public ReadOnly Property OutputRoot As String
    Public ReadOnly Property Timestamp As String
    Public ReadOnly Property RunLogPath As String
    Public ReadOnly Property ComErrorsPath As String

    Private ReadOnly _logger As Logger
    Private ReadOnly _auditPaths As New List(Of String)()
    Private ReadOnly _geometryPaths As New List(Of String)()

    Private Sub New(outputRoot As String, ts As String, runLog As String, comErr As String, logger As Logger)
        Me.OutputRoot = outputRoot
        Me.Timestamp = ts
        Me.RunLogPath = runLog
        Me.ComErrorsPath = comErr
        _logger = logger
    End Sub

    Public ReadOnly Property AuditPathsGenerated As IReadOnlyList(Of String)
        Get
            Return _auditPaths.AsReadOnly()
        End Get
    End Property

    Public ReadOnly Property GeometryPathsGenerated As IReadOnlyList(Of String)
        Get
            Return _geometryPaths.AsReadOnly()
        End Get
    End Property

    ''' <summary>
    ''' Inicializa el contexto: crea <c>docs\logs</c> si no existe y adjunta los mirrors al logger.
    ''' <paramref name="jobOutputRoot"/> solo se registra en el log (DFT/PDF/etc. siguen en la carpeta de trabajo del trabajo); las trazas van siempre a repositorio\docs\logs.
    ''' Devuelve Nothing si <paramref name="logger"/> es Nothing.
    ''' </summary>
    Public Shared Function Initialize(jobOutputRoot As String, logger As Logger) As RunDiagnosticsContext
        If logger Is Nothing Then Return Nothing

        Dim logsRoot As String = DiagnosticsLogPaths.GetRepositoryDocsLogsDirectory()
        Try
            If Not Directory.Exists(logsRoot) Then Directory.CreateDirectory(logsRoot)
        Catch ex As Exception
            logger.LogException("RunDiagnosticsContext.Initialize CreateDirectory docs\logs", ex)
            Return Nothing
        End Try

        ' Sesión ya abierta (p. ej. arranque de MainForm): reutilizar los mismos ficheros timestamped.
        If Not String.IsNullOrWhiteSpace(logger.RunLogPath) AndAlso Not String.IsNullOrWhiteSpace(logger.ComErrorsPath) Then
            Dim tsReuse As String = ParseTimestampFromRunLogPath(logger.RunLogPath)
            If String.IsNullOrWhiteSpace(tsReuse) Then
                tsReuse = DateTime.Now.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture)
            End If
            logger.Log("[DIAG][INIT][REUSE] runLog=" & logger.RunLogPath)
            logger.Log("[DIAG][INIT][REUSE] comErrors=" & logger.ComErrorsPath)
            logger.Log("[DIAG][INIT] jobOutputRoot=" & If(jobOutputRoot, ""))
            Return New RunDiagnosticsContext(logsRoot, tsReuse, logger.RunLogPath, logger.ComErrorsPath, logger)
        End If

        Dim ts As String = DateTime.Now.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture)
        Dim runLog As String = Path.Combine(logsRoot, "run_log_" & ts & ".txt")
        Dim comErr As String = Path.Combine(logsRoot, "com_errors_" & ts & ".txt")

        logger.AttachRunLogFile(runLog)
        logger.AttachComErrorsFile(comErr)

        Dim ctx As New RunDiagnosticsContext(logsRoot, ts, runLog, comErr, logger)
        logger.Log("[DIAG][INIT] logsRoot=" & logsRoot)
        logger.Log("[DIAG][INIT] jobOutputRoot=" & If(jobOutputRoot, ""))
        logger.Log("[DIAG][INIT] runLog=" & runLog)
        logger.Log("[DIAG][INIT] comErrors=" & comErr)
        logger.Log("[DIAG][INIT] sdkDocHint=docs\SDK_HTML (fuente prioritaria API Solid Edge)")
        Return ctx
    End Function

    ''' <summary>
    ''' Crea audit + geometry para el <paramref name="dftDoc"/> aún abierto.
    ''' Usa <paramref name="pieceBaseName"/> como prefijo para distinguir múltiples DFTs en lote.
    ''' Devuelve True si al menos uno de los archivos se generó.
    ''' </summary>
    Public Function CaptureDiagnosticsForDft(dftDoc As DraftDocument, pieceBaseName As String) As Boolean
        If dftDoc Is Nothing Then
            _logger.Log("[DIAG][SKIP] dftDoc=Nothing pieceBaseName=" & If(pieceBaseName, ""))
            Return False
        End If

        Dim safeBase As String = SanitizeBaseName(pieceBaseName)
        Dim auditPath As String = Path.Combine(OutputRoot, "audit_" & safeBase & "_" & Timestamp & ".txt")
        Dim geomPath As String = Path.Combine(OutputRoot, "geometry_" & safeBase & "_" & Timestamp & ".txt")

        Dim anyOk As Boolean = False
        ' AUDIT
        Try
            DftAuditService.ExportAuditFromOpenDocument(dftDoc, auditPath, _logger)
            _auditPaths.Add(auditPath)
            _logger.Log("[DIAG][AUDIT][OK] " & auditPath)
            anyOk = True
        Catch ex As Exception
            _logger.LogException("DftAuditService.ExportAuditFromOpenDocument", ex)
        End Try

        ' GEOMETRY
        Try
            DraftGeometryReporter.ExportDraftGeometryLog(dftDoc, geomPath, _logger)
            _geometryPaths.Add(geomPath)
            _logger.Log("[DIAG][GEOM][OK] " & geomPath)
            anyOk = True
        Catch ex As Exception
            _logger.LogException("DraftGeometryReporter.ExportDraftGeometryLog", ex)
        End Try

        Return anyOk
    End Function

    ''' <summary>Vuelca el snapshot completo del logger al run_log para cerrarlo (líneas que pudieran haberse perdido por excepción del file sink).</summary>
    Public Sub Finish()
        If _logger Is Nothing Then Return
        Try
            If Not String.IsNullOrWhiteSpace(RunLogPath) Then
                ' El mirror append-only ya escribió línea a línea. Como cierre, sobreescribimos con
                ' el snapshot final del logger para garantizar persistencia completa y orden absoluto.
                _logger.SaveToFile(RunLogPath)
            End If
        Catch ex As Exception
            _logger.LogException("RunDiagnosticsContext.Finish SaveLog", ex)
        End Try

        _logger.Log("[SUMMARY][DIAG] runLog=" & RunLogPath)
        _logger.Log("[SUMMARY][DIAG] comErrors=" & ComErrorsPath)
        If _auditPaths.Count = 0 Then
            _logger.Log("[SUMMARY][DIAG] audit=(no se generó ningún DFT)")
        Else
            For Each p In _auditPaths
                _logger.Log("[SUMMARY][DIAG] audit=" & p)
            Next
        End If
        If _geometryPaths.Count = 0 Then
            _logger.Log("[SUMMARY][DIAG] geometry=(no se generó ningún DFT)")
        Else
            For Each p In _geometryPaths
                _logger.Log("[SUMMARY][DIAG] geometry=" & p)
            Next
        End If
    End Sub

    Private Shared Function ParseTimestampFromRunLogPath(runLogFullPath As String) As String
        If String.IsNullOrWhiteSpace(runLogFullPath) Then Return ""
        Dim baseName As String = Path.GetFileNameWithoutExtension(runLogFullPath)
        Const prefix As String = "run_log_"
        If baseName.Length <= prefix.Length Then Return ""
        If Not baseName.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) Then Return ""
        Return baseName.Substring(prefix.Length)
    End Function

    Friend Shared Function SanitizeBaseName(name As String) As String
        If String.IsNullOrWhiteSpace(name) Then Return "dft"
        Dim invalid As Char() = Path.GetInvalidFileNameChars()
        Dim sb As New StringBuilder(name.Length)
        For Each ch As Char In name
            If Array.IndexOf(invalid, ch) >= 0 OrElse ch = " "c Then
                sb.Append("_"c)
            Else
                sb.Append(ch)
            End If
        Next
        Dim s As String = sb.ToString().Trim("_"c)
        If String.IsNullOrWhiteSpace(s) Then s = "dft"
        If s.Length > 80 Then s = s.Substring(0, 80)
        Return s
    End Function
End Class
