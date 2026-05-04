Option Strict Off

Imports System.IO

''' <summary>
''' Resuelve valores efectivos del cajetín (título automático, autor) alineados con la UI.
''' </summary>
Friend NotInheritable Class TitleBlockFieldCoordinator

    Private Sub New()
    End Sub

    Public Shared Function ResolveEffectiveTitle(config As JobConfiguration, modelPath As String) As String
        If config Is Nothing Then Return ""
        If config.TitleSourceMode = TitleSourceMode.AutoFromFileName Then
            If Not String.IsNullOrWhiteSpace(modelPath) Then
                Try
                    Return Path.GetFileNameWithoutExtension(modelPath)
                Catch
                End Try
            End If
        End If
        Return If(config.DrawingTitle, "").Trim()
    End Function

    Public Shared Function ResolveEffectiveAuthor(config As JobConfiguration) As String
        If config Is Nothing Then Return Environment.UserName
        Dim a As String = If(config.AuthorName, "").Trim()
        If a <> "" Then Return a
        Return Environment.UserName
    End Function

    ''' <summary>Aplica título y autor efectivos a una copia de configuración para escritura en modelo/draft.</summary>
    Public Shared Function PrepareTitleBlockFields(config As JobConfiguration, modelPath As String) As JobConfiguration
        If config Is Nothing Then Return Nothing
        Dim c As JobConfiguration = config.CloneForExecution()
        c.DrawingTitle = ResolveEffectiveTitle(config, modelPath)
        Dim auth As String = ResolveEffectiveAuthor(config)
        c.AuthorName = auth
        Return c
    End Function

End Class
