Option Strict Off
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Diagnostics

''' <summary>Formulario con barra de progreso 0-100%, tiempo transcurrido y log de ficheros procesados.</summary>
Public Class ProgressForm
    Inherits Form

    Private _progressBar As New ProgressBar()
    Private _label As New Label()
    Private _timeLabel As New Label()
    Private _logBox As New TextBox()
    Private _stopwatch As New Stopwatch()
    Private _timer As New Timer()

    Public Sub New()
        Me.Text = "Extrayendo planos..."
        Me.FormBorderStyle = FormBorderStyle.Sizable
        Me.Size = New Size(700, 480)
        Me.MinimumSize = New Size(550, 350)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.ControlBox = True
        Me.MinimizeBox = False
        Me.MaximizeBox = True
        Me.TopMost = True

        _timeLabel.Dock = DockStyle.Top
        _timeLabel.Height = 24
        _timeLabel.TextAlign = ContentAlignment.MiddleCenter
        _timeLabel.Text = "Tiempo: 0:00"
        _timeLabel.Font = New Font(_timeLabel.Font.FontFamily, 10, FontStyle.Bold)

        _label.Dock = DockStyle.Top
        _label.Height = 28
        _label.TextAlign = ContentAlignment.MiddleCenter
        _label.Text = "0 %"

        _progressBar.Minimum = 0
        _progressBar.Maximum = 100
        _progressBar.Value = 0
        _progressBar.Style = ProgressBarStyle.Continuous
        _progressBar.Dock = DockStyle.Top
        _progressBar.Height = 24

        _logBox.Multiline = True
        _logBox.ReadOnly = True
        _logBox.ScrollBars = ScrollBars.Vertical
        _logBox.Dock = DockStyle.Fill
        _logBox.Font = New Font("Consolas", 9)
        _logBox.BackColor = Color.White

        Me.Controls.Add(_logBox)
        Me.Controls.Add(_progressBar)
        Me.Controls.Add(_label)
        Me.Controls.Add(_timeLabel)

        AddHandler Me.Shown, Sub()
            _stopwatch.Start()
            _timer.Interval = 500
            AddHandler _timer.Tick, Sub(s, e) UpdateElapsedTime()
            _timer.Start()
        End Sub

        AddHandler Me.FormClosing, Sub(s As Object, e As FormClosingEventArgs)
            _timer.Stop()
            _stopwatch.Stop()
        End Sub
    End Sub

    Private Sub UpdateElapsedTime()
        Dim ts = _stopwatch.Elapsed
        Dim msg = FormatElapsed(ts)
        If _timeLabel.InvokeRequired Then
            Try
                _timeLabel.Invoke(New Action(Sub() _timeLabel.Text = "Tiempo transcurrido: " & msg))
            Catch
            End Try
        Else
            _timeLabel.Text = "Tiempo transcurrido: " & msg
        End If
    End Sub

    Private Shared Function FormatElapsed(ts As TimeSpan) As String
        If ts.TotalHours >= 1 Then
            Return $"{ts.Hours}:{ts.Minutes:D2}:{ts.Seconds:D2}"
        Else
            Return $"{ts.Minutes}:{ts.Seconds:D2}"
        End If
    End Function

    ''' <summary>Añade una línea al log de ficheros procesados.</summary>
    Public Sub AppendLog(line As String)
        If _logBox.InvokeRequired Then
            _logBox.BeginInvoke(New Action(Of String)(AddressOf AppendLogInternal), line)
        Else
            AppendLogInternal(line)
        End If
    End Sub

    Private Sub AppendLogInternal(line As String)
        _logBox.AppendText(line & Environment.NewLine)
        _logBox.SelectionStart = _logBox.Text.Length
        _logBox.ScrollToCaret()
    End Sub

    Public Sub SetProgress(percent As Integer)
        Dim v As Integer = Math.Max(0, Math.Min(100, percent))
        If _progressBar.InvokeRequired Then
            _progressBar.BeginInvoke(New Action(Of Integer)(AddressOf SetProgressInternal), v)
        Else
            SetProgressInternal(v)
        End If
    End Sub

    Private Sub SetProgressInternal(percent As Integer)
        _progressBar.Value = percent
        _label.Text = $"{percent} %"
    End Sub

    ''' <summary>Detiene el cronómetro y muestra el tiempo total. Llamar al finalizar el proceso.</summary>
    Public Sub SetFinished()
        _timer.Stop()
        _stopwatch.Stop()
        Dim ts = _stopwatch.Elapsed
        Dim msg = "Finalizado en " & FormatElapsed(ts)
        If _timeLabel.InvokeRequired Then
            Try
                _timeLabel.Invoke(New Action(Sub()
                    _timeLabel.Text = msg
                    Me.Text = "Extrayendo planos... - " & msg
                End Sub))
            Catch
            End Try
        Else
            _timeLabel.Text = msg
            Me.Text = "Extrayendo planos... - " & msg
        End If
    End Sub
End Class
