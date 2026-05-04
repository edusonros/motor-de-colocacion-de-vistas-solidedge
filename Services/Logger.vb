Option Strict Off

Imports System.Collections.Concurrent
Imports System.IO
Imports System.Text
Imports System.Threading

Public Class Logger
    Private _step As Integer = 0
    Private _lines As ConcurrentQueue(Of String) = New ConcurrentQueue(Of String)()
    Private ReadOnly _sync As New Object()
    Private _uiSink As Action(Of String) = Nothing

    Public Sub New(Optional uiSink As Action(Of String) = Nothing)
        _uiSink = uiSink
    End Sub

    Public Sub SetUiSink(sink As Action(Of String))
        _uiSink = sink
    End Sub

    Public Sub Reset()
        SyncLock _sync
            _step = 0
            _lines = New ConcurrentQueue(Of String)()
        End SyncLock
    End Sub

    Public Sub Log(message As String)
        Dim currentStep As Integer = Interlocked.Increment(_step)
        Dim line As String = $"[{currentStep:000}] {message}"
        _lines.Enqueue(line)
        Dim sink As Action(Of String) = _uiSink
        If sink IsNot Nothing Then sink.Invoke(line)
    End Sub

    Public Sub LogRaw(line As String)
        _lines.Enqueue(line)
        Dim sink As Action(Of String) = _uiSink
        If sink IsNot Nothing Then sink.Invoke(line)
    End Sub

    Public Sub LogException(context As String, ex As Exception)
        Dim hr As String = ""
        If TypeOf ex Is Runtime.InteropServices.COMException Then
            Dim cex = CType(ex, Runtime.InteropServices.COMException)
            hr = $" HR=0x{cex.ErrorCode:X8}"
        End If
        LogRaw($"[EX] {context}: {ex.GetType().Name}{hr} -> {ex.Message}")
    End Sub

    Public Function GetAllText() As String
        Dim snapshot As String() = _lines.ToArray()
        Return String.Join(Environment.NewLine, snapshot)
    End Function

    Public Sub SaveToFile(filePath As String)
        ' Evitar ambigüedades con nombres "Path" en el proyecto.
        Dim info As New FileInfo(filePath)
        Dim dir As String = info.DirectoryName
        If Not String.IsNullOrWhiteSpace(dir) AndAlso Not Directory.Exists(dir) Then
            Directory.CreateDirectory(dir)
        End If
        File.WriteAllText(filePath, GetAllText(), Encoding.UTF8)
    End Sub
End Class
