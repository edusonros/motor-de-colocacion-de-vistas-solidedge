Option Strict Off
Imports System.Diagnostics
Imports System.Reflection
Imports System.Windows.Forms

Module Program
    ' Nota: este proyecto se mantiene en .NET Framework 4.8 para compatibilidad
    ' estable con referencias COM de Solid Edge (Interop clásico).
    <STAThread()>
    Sub Main()
        Try
            Debug.WriteLine("[BOOT][EXE_PATH] " & Assembly.GetExecutingAssembly().Location)
            Debug.WriteLine("[BOOT][CURRENT_DIR] " & Environment.CurrentDirectory)
            Debug.WriteLine("[BOOT][STARTUP_PATH] " & Application.StartupPath)
        Catch
        End Try
        Application.EnableVisualStyles()
        Application.SetCompatibleTextRenderingDefault(False)
        Application.Run(New MainForm())
    End Sub
End Module