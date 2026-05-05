Option Strict Off

Imports System.Threading
Imports System.Threading.Tasks

''' <summary>
''' Ejecuta trabajo COM (Solid Edge) en un subproceso <see cref="ApartmentState.STA"/> aparte para no bloquear el hilo de la interfaz WinForms.
''' Sin esto <c>Documents.Open</c> / ocurrencias ASM suelen «congelar» el formulario y los cuadros de diálogo de SE no reciben mensajes.
''' Un único trabajo COM a la vez evita corrupción al abrir/cierra documentos desde varios llamadores paralelos.
''' </summary>
Public NotInheritable Class StaComInvoker
    Private Shared ReadOnly _comSerial As New SemaphoreSlim(1, 1)

    Private Sub New()
    End Sub

    Public Shared Function Run(Of T)(work As Func(Of T)) As Task(Of T)
        If work Is Nothing Then Throw New ArgumentNullException(NameOf(work))
        Dim tcs As New TaskCompletionSource(Of T)(TaskCreationOptions.RunContinuationsAsynchronously)
        Dim th As New Thread(
            Sub()
                _comSerial.Wait()
                Try
                    tcs.TrySetResult(work())
                Catch ex As Exception
                    tcs.TrySetException(ex)
                Finally
                    Try
                        _comSerial.Release()
                    Catch
                    End Try
                End Try
            End Sub)
        th.SetApartmentState(ApartmentState.STA)
        th.IsBackground = True
        th.Name = "StaComInvoker STA"
        th.Start()
        Return tcs.Task
    End Function
End Class
