Option Strict Off

Imports System.Collections.Concurrent
Imports System.Collections.Generic
Imports System.Globalization
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading

''' <summary>
''' Logger central del proyecto.
''' - Mantiene un buffer en memoria de todas las líneas (`GetAllText`, `SaveToFile`).
''' - Emite por consola/UI a través del <c>uiSink</c>.
''' - Mirror opcional al fichero <c>run_log_yyyyMMdd_HHmmss.txt</c> bajo <c>docs\logs</c> (ver <see cref="RunDiagnosticsContext"/>).
''' - Mirror opcional de errores COM a <c>com_errors_yyyyMMdd_HHmmss.txt</c> en la misma carpeta.
''' - <see cref="LogComException"/> detecta el método SE que falló y emite la línea
'''   <c>[COM][SDK_REQUEST_REQUIRED]</c> obligatoria para que el usuario aporte la doc del SDK.
''' </summary>
Public Class Logger
    Private _step As Integer = 0
    Private _lines As ConcurrentQueue(Of String) = New ConcurrentQueue(Of String)()
    Private ReadOnly _sync As New Object()
    Private _uiSink As Action(Of String) = Nothing

    Private _runLogPath As String = ""
    Private _comErrorsPath As String = ""
    Private ReadOnly _fileLock As New Object()

    ''' <summary>Palabras clave SE para parsear el método SE en LogComException cuando solo hay <c>context</c>.</summary>
    Private Shared ReadOnly SeMethodHintRegex As Regex = New Regex(
        "(Add[A-Z][A-Za-z0-9_]+|Get[A-Z][A-Za-z0-9_]+|Set[A-Z][A-Za-z0-9_]+|" &
        "Apply[A-Z][A-Za-z0-9_]+|Open|Close|Save|SaveAs|Update|UpdateAll|UpdateViews|" &
        "Item|Quit|Activate|DoIdle|Recompute|Range|GetOrigin|SetOrigin|RetrieveDimensions|" &
        "SetViewOrientationStandard|SetRotationAngle|ScaleFactor|TrackDistance|" &
        "AddByFold|AddAssemblyView|AddPartView|AddSheetMetalView|AddDistanceBetweenObjects|" &
        "GetReferenceToGraphicMember|DVLines2d|DVArcs2d|DVCircles2d|DVBSplineCurves2d|DVPoints2d|" &
        "ModelLinks|Documents|DrawingViews|Sheets|Dimensions|PartsLists)",
        RegexOptions.Compiled)

    Public Sub New(Optional uiSink As Action(Of String) = Nothing)
        _uiSink = uiSink
    End Sub

    Public Sub SetUiSink(sink As Action(Of String))
        _uiSink = sink
    End Sub

    ''' <summary>Habilita el mirror automático del log de ejecución al archivo indicado.</summary>
    Public Sub AttachRunLogFile(path As String)
        SyncLock _fileLock
            _runLogPath = If(path, "")
            If Not String.IsNullOrWhiteSpace(_runLogPath) Then
                Try
                    Dim dir As String = IO.Path.GetDirectoryName(_runLogPath)
                    If Not String.IsNullOrWhiteSpace(dir) AndAlso Not Directory.Exists(dir) Then
                        Directory.CreateDirectory(dir)
                    End If
                    File.AppendAllText(_runLogPath,
                        "=== run_log open " & DateTime.Now.ToString("o", CultureInfo.InvariantCulture) & Environment.NewLine,
                        Encoding.UTF8)
                Catch
                    ' No bloquear si I/O falla.
                End Try
            End If
        End SyncLock
    End Sub

    ''' <summary>Habilita el mirror automático de errores COM al archivo indicado.</summary>
    Public Sub AttachComErrorsFile(path As String)
        SyncLock _fileLock
            _comErrorsPath = If(path, "")
            If Not String.IsNullOrWhiteSpace(_comErrorsPath) Then
                Try
                    Dim dir As String = IO.Path.GetDirectoryName(_comErrorsPath)
                    If Not String.IsNullOrWhiteSpace(dir) AndAlso Not Directory.Exists(dir) Then
                        Directory.CreateDirectory(dir)
                    End If
                    File.AppendAllText(_comErrorsPath,
                        "=== com_errors open " & DateTime.Now.ToString("o", CultureInfo.InvariantCulture) & Environment.NewLine,
                        Encoding.UTF8)
                Catch
                End Try
            End If
        End SyncLock
    End Sub

    Public ReadOnly Property RunLogPath As String
        Get
            Return _runLogPath
        End Get
    End Property

    Public ReadOnly Property ComErrorsPath As String
        Get
            Return _comErrorsPath
        End Get
    End Property

    Public Sub Reset()
        SyncLock _sync
            _step = 0
            _lines = New ConcurrentQueue(Of String)()
        End SyncLock
    End Sub

    Public Sub Log(message As String)
        Dim currentStep As Integer = Interlocked.Increment(_step)
        Dim line As String = $"[{currentStep:000}] {message}"
        EmitLine(line)
    End Sub

    Public Sub LogRaw(line As String)
        EmitLine(line)
    End Sub

    ''' <summary>Pone una excepción en el log con HRESULT si es COM. Si es COMException, además vuelca al com_errors y emite SDK_REQUEST_REQUIRED.</summary>
    Public Sub LogException(context As String, ex As Exception)
        If TypeOf ex Is COMException Then
            LogComException(context, ex)
            Return
        End If
        LogRaw($"[EX] {context}: {ex.GetType().Name} -> {ex.Message}")
    End Sub

    ''' <summary>Registra un error COM con detalle completo + SDK_REQUEST_REQUIRED.
    ''' Si <paramref name="methodName"/> u <paramref name="objectName"/> son vacíos se intenta inferir desde <paramref name="context"/>.</summary>
    Public Sub LogComException(context As String, ex As Exception,
                                Optional methodName As String = "",
                                Optional objectName As String = "",
                                Optional itemIndex As Integer = -1)
        If ex Is Nothing Then Return
        Dim cex As COMException = TryCast(ex, COMException)
        Dim hr As String = ""
        If cex IsNot Nothing Then hr = $"0x{cex.ErrorCode:X8}"

        ' Inferencia de método/objeto desde el contexto si no se pasaron.
        Dim resolvedMethod As String = methodName
        Dim resolvedObject As String = objectName
        If String.IsNullOrWhiteSpace(resolvedMethod) Then
            ParseMethodAndObjectFromContext(context, resolvedMethod, resolvedObject)
        End If

        Dim message As String = If(ex.Message, "")
        Dim source As String = If(ex.Source, "")
        Dim targetSite As String = ""
        Try
            If ex.TargetSite IsNot Nothing Then
                targetSite = ex.TargetSite.DeclaringType?.FullName & "." & ex.TargetSite.Name
            End If
        Catch
        End Try
        Dim stackTrace As String = If(ex.StackTrace, "")

        ' Línea resumida en el log principal: identifica claramente qué método de SE falló.
        Dim summary As New StringBuilder()
        summary.Append("[COM][ERROR]")
        If Not String.IsNullOrWhiteSpace(resolvedMethod) Then summary.Append(" method=").Append(resolvedMethod)
        If Not String.IsNullOrWhiteSpace(resolvedObject) Then summary.Append(" object=").Append(resolvedObject)
        If itemIndex >= 0 Then summary.Append(" index=").Append(itemIndex.ToString(CultureInfo.InvariantCulture))
        If Not String.IsNullOrWhiteSpace(hr) Then summary.Append(" HRESULT=").Append(hr)
        summary.Append(" context=").Append(SafeOneLine(context))
        summary.Append(" message=").Append(SafeOneLine(message))
        EmitLine(summary.ToString())

        ' Petición explícita de documentación SDK obligatoria si pudimos identificar el método.
        If Not String.IsNullOrWhiteSpace(resolvedMethod) Then
            Dim sdkReq As New StringBuilder()
            sdkReq.Append("[COM][SDK_REQUEST_REQUIRED] method=").Append(resolvedMethod)
            If Not String.IsNullOrWhiteSpace(resolvedObject) Then sdkReq.Append(" object=").Append(resolvedObject)
            sdkReq.Append(" reason=Necesaria documentación SDK/API exacta")
            EmitLine(sdkReq.ToString())
        End If

        ' Entrada completa en com_errors_*.txt (con stack).
        Try
            SyncLock _fileLock
                If Not String.IsNullOrWhiteSpace(_comErrorsPath) Then
                    Dim sb As New StringBuilder()
                    sb.AppendLine("[COM][ERROR] " & DateTime.Now.ToString("o", CultureInfo.InvariantCulture))
                    sb.AppendLine("context=" & context)
                    sb.AppendLine("method=" & resolvedMethod)
                    sb.AppendLine("object=" & resolvedObject)
                    If itemIndex >= 0 Then sb.AppendLine("index=" & itemIndex.ToString(CultureInfo.InvariantCulture))
                    sb.AppendLine("hresult=" & hr)
                    sb.AppendLine("type=" & ex.GetType().FullName)
                    sb.AppendLine("message=" & message)
                    sb.AppendLine("source=" & source)
                    sb.AppendLine("targetSite=" & targetSite)
                    sb.AppendLine("stackTrace=")
                    sb.AppendLine(stackTrace)
                    sb.AppendLine("---")
                    File.AppendAllText(_comErrorsPath, sb.ToString(), Encoding.UTF8)
                End If
            End SyncLock
        Catch
            ' No bloquear si I/O falla.
        End Try
    End Sub

    ''' <summary>Atajo para registrar un error COM cuando ya conoces el método y el objeto SE.</summary>
    Public Sub LogComMethodFailure(methodName As String, objectName As String, ex As Exception,
                                    Optional itemIndex As Integer = -1,
                                    Optional context As String = "")
        Dim ctx As String = If(String.IsNullOrWhiteSpace(context), objectName & "." & methodName, context)
        LogComException(ctx, ex, methodName, objectName, itemIndex)
    End Sub

    Public Function GetAllText() As String
        Dim snapshot As String() = _lines.ToArray()
        Return String.Join(Environment.NewLine, snapshot)
    End Function

    Public Sub SaveToFile(filePath As String)
        Dim info As New FileInfo(filePath)
        Dim dir As String = info.DirectoryName
        If Not String.IsNullOrWhiteSpace(dir) AndAlso Not Directory.Exists(dir) Then
            Directory.CreateDirectory(dir)
        End If
        File.WriteAllText(filePath, GetAllText(), Encoding.UTF8)
    End Sub

    ' --- Internos --------------------------------------------------------

    Private Sub EmitLine(line As String)
        _lines.Enqueue(line)
        Dim sink As Action(Of String) = _uiSink
        If sink IsNot Nothing Then
            Try
                sink.Invoke(line)
            Catch
            End Try
        End If
        ' Mirror al run_log si está activo.
        Try
            SyncLock _fileLock
                If Not String.IsNullOrWhiteSpace(_runLogPath) Then
                    File.AppendAllText(_runLogPath, line & Environment.NewLine, Encoding.UTF8)
                End If
            End SyncLock
        Catch
        End Try
    End Sub

    Private Shared Sub ParseMethodAndObjectFromContext(context As String,
                                                       ByRef methodName As String,
                                                       ByRef objectName As String)
        methodName = ""
        objectName = ""
        If String.IsNullOrWhiteSpace(context) Then Return

        ' Patrón explícito "method=X object=Y" si el llamante ya lo proporciona.
        Dim mMethodKv As Match = Regex.Match(context, "method\s*=\s*([A-Za-z0-9_\.]+)", RegexOptions.IgnoreCase)
        If mMethodKv.Success Then methodName = mMethodKv.Groups(1).Value
        Dim mObjectKv As Match = Regex.Match(context, "object\s*=\s*([A-Za-z0-9_\.]+)", RegexOptions.IgnoreCase)
        If mObjectKv.Success Then objectName = mObjectKv.Groups(1).Value
        If Not String.IsNullOrWhiteSpace(methodName) Then Return

        ' Patrón "Class.Method" (lo más común que ya usamos en los Try-Catch).
        Dim mDot As Match = Regex.Match(context, "([A-Z][A-Za-z0-9_]+)\.([A-Za-z][A-Za-z0-9_]+)")
        If mDot.Success Then
            objectName = mDot.Groups(1).Value
            methodName = mDot.Groups(2).Value
            Return
        End If

        ' Fallback: cualquier identificador SE conocido en el contexto.
        Dim mHint As Match = SeMethodHintRegex.Match(context)
        If mHint.Success Then methodName = mHint.Value
    End Sub

    Private Shared Function SafeOneLine(s As String) As String
        If s Is Nothing Then Return ""
        Return s.Replace(ControlChars.CrLf, " ").Replace(ControlChars.Cr, " ").Replace(ControlChars.Lf, " ")
    End Function
End Class
