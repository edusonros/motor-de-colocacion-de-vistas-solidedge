Option Strict Off

Imports System.Globalization
Imports System.IO
Imports System.Text

Namespace Services.Dimensioning.Labs

    ''' <summary>Logger de archivo para DIMLAB: OUT_DIMLAB/DIMLAB_yyyyMMdd_HHmmss.txt + reenvío opcional al log de aplicación.</summary>
    Friend NotInheritable Class DimLabLogger

        Private ReadOnly _filePath As String
        Private ReadOnly _appLog As Action(Of String)
        Private ReadOnly _sync As New Object()

        Public Sub New(outRoot As String, appLog As Action(Of String))
            _appLog = appLog
            Dim root = If(String.IsNullOrWhiteSpace(outRoot), Path.Combine(System.Environment.CurrentDirectory, "OUT_DIMLAB"), outRoot.Trim())
            Try
                Directory.CreateDirectory(root)
            Catch
            End Try
            Dim stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture)
            _filePath = Path.Combine(root, "DIMLAB_" & stamp & ".txt")
        End Sub

        Public ReadOnly Property TextFilePath As String
            Get
                Return _filePath
            End Get
        End Property

        Public Sub Log(area As String, ev As String, Optional detail As String = Nothing)
            Dim sb As New StringBuilder()
            sb.Append("[DIMLAB][").Append(area).Append("][").Append(ev).Append("]")
            If Not String.IsNullOrWhiteSpace(detail) Then
                sb.Append(" ").Append(detail)
            End If
            Dim line = sb.ToString()
            SyncLock _sync
                Try
                    File.AppendAllText(_filePath, line & Environment.NewLine, Encoding.UTF8)
                Catch
                End Try
            End SyncLock
            Try
                _appLog?.Invoke(line)
            Catch
            End Try
        End Sub

    End Class

End Namespace
