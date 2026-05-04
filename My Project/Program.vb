Option Strict Off
Imports System.Windows.Forms

Module Program
    ' Nota: este proyecto se mantiene en .NET Framework 4.8 para compatibilidad
    ' estable con referencias COM de Solid Edge (Interop clásico).
    <STAThread()>
    Sub Main()
        Application.EnableVisualStyles()
        Application.SetCompatibleTextRenderingDefault(False)
        Application.Run(New MainForm())
    End Sub
End Module