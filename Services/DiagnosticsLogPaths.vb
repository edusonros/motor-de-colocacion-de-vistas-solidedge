Option Strict On

Imports System.IO

''' <summary>
''' Resuelve <c>&lt;raíz_repo&gt;\docs\logs</c> subiendo desde el directorio base del ejecutable
''' (<c>AppDomain.CurrentDomain.BaseDirectory</c>) hasta encontrar marcadores del proyecto (<c>Extraer_dft_dxf_flatdxf.vbproj</c> o <c>docs\SDK_HTML</c>),
''' así funciona igual desde <c>bin\Debug</c>, <c>bin\x64\Debug</c>, <c>bin\Release</c>, etc.
''' </summary>
Public NotInheritable Class DiagnosticsLogPaths
    Private Const VbProjFile As String = "Extraer_dft_dxf_flatdxf.vbproj"

    Private Sub New()
    End Sub

    Public Shared Function GetRepositoryDocsLogsDirectory() As String
        Dim baseDir As String = AppDomain.CurrentDomain.BaseDirectory
        Dim d As DirectoryInfo = New DirectoryInfo(baseDir)
        For depth As Integer = 0 To 16
            If d Is Nothing Then Exit For
            Dim vk As String = Path.Combine(d.FullName, VbProjFile)
            Dim sdkHtml As String = Path.Combine(d.FullName, "docs", "SDK_HTML")
            If File.Exists(vk) OrElse Directory.Exists(sdkHtml) Then
                Return Path.Combine(d.FullName, "docs", "logs")
            End If
            d = d.Parent
        Next
        Return Path.GetFullPath(Path.Combine(baseDir, "..", "..", "docs", "logs"))
    End Function
End Class
