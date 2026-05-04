Option Strict Off

Public Module LogUtil
    Private _stepId As Integer = 0
    Private _sink As Action(Of String) = Nothing

    Public Sub SetLogSink(sink As Action(Of String))
        _sink = sink
    End Sub

    Public Sub ResetStepCounter()
        _stepId = 0
    End Sub

    Public Sub StepLog(msg As String)
        _stepId += 1
        Dim line As String = $"[{_stepId:000}] {msg}"
        Try
            If _sink IsNot Nothing Then
                _sink.Invoke(line)
            Else
                Console.WriteLine(line)
            End If
        Catch
            Console.WriteLine(line)
        End Try
    End Sub
End Module

'Module LogUtil
'    Private _stepId As Integer = 0

'    Public Sub StepLog(msg As String)
'        _stepId += 1
'        Console.WriteLine($"[{_stepId:000}] {msg}")
'    End Sub

'    Public Sub Log(msg As String)
'        Console.WriteLine(msg)
'    End Sub

'    Public Sub LogEx(context As String, ex As Exception)
'        Console.WriteLine($"[EX] {context}: {ex.GetType().Name} -> {ex.Message}")
'    End Sub
'End Module