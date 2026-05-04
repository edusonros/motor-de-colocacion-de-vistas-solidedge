Option Strict Off

Imports SolidEdgeDraft

Friend NotInheritable Class DimensionStyleResolver
    Private Sub New()
    End Sub

    Public Shared Function ResolvePreferredStyle(draft As DraftDocument, sheet As Sheet, styleName As String, log As DimensionLogger) As Object
        If draft Is Nothing AndAlso sheet Is Nothing Then Return Nothing
        Try
            Dim styles As Object = Nothing
            If draft IsNot Nothing Then
                Try
                    styles = CallByName(draft, "DimensionStyles", CallType.Get)
                Catch
                    styles = Nothing
                End Try
            End If
            If styles Is Nothing AndAlso sheet IsNot Nothing Then
                Try
                    styles = CallByName(sheet, "DimensionStyles", CallType.Get)
                Catch
                    styles = Nothing
                End Try
            End If
            If styles Is Nothing Then
                log?.LogLine("[DIM][STYLE][ERR] DimensionStyles no disponible.")
                Return Nothing
            End If

            Dim names As String = ListStyleNames(styles)
            If Not String.IsNullOrWhiteSpace(names) Then
                log?.LogLine("[DIM][STYLE][LIST] " & names)
            End If

            Dim styleObj As Object = FindStyleByName(styles, styleName)
            If styleObj Is Nothing Then
                log?.LogLine("[DIM][STYLE][WARN] No se encontró estilo preferido '" & styleName & "'. Se continúa con estilo por defecto.")
                Return Nothing
            End If
            Dim resolvedName As String = ""
            Try
                resolvedName = Convert.ToString(CallByName(styleObj, "Name", CallType.Get), Globalization.CultureInfo.InvariantCulture)
            Catch
                resolvedName = styleName
            End Try
            log?.LogLine("[DIM][STYLE] resolved=" & resolvedName)
            Return styleObj
        Catch ex As Exception
            log?.ComFail("ResolvePreferredStyle", "DimensionStyles", ex)
            Return Nothing
        End Try
    End Function

    Private Shared Function ListStyleNames(styles As Object) As String
        If styles Is Nothing Then Return ""
        Try
            Dim n As Integer = CInt(CallByName(styles, "Count", CallType.Get))
            If n <= 0 Then Return "(sin estilos)"
            Dim parts As New List(Of String)()
            For i As Integer = 1 To n
                Try
                    Dim it As Object = CallByName(styles, "Item", CallType.Method, i)
                    Dim nm As String = Convert.ToString(CallByName(it, "Name", CallType.Get), Globalization.CultureInfo.InvariantCulture)
                    If Not String.IsNullOrWhiteSpace(nm) Then parts.Add(nm)
                Catch
                End Try
            Next
            Return String.Join(", ", parts)
        Catch
            Return ""
        End Try
    End Function

    Private Shared Function FindStyleByName(styles As Object, styleName As String) As Object
        If styles Is Nothing OrElse String.IsNullOrWhiteSpace(styleName) Then Return Nothing
        Try
            Dim n As Integer = CInt(CallByName(styles, "Count", CallType.Get))
            For i As Integer = 1 To n
                Try
                    Dim it As Object = CallByName(styles, "Item", CallType.Method, i)
                    Dim nm As String = Convert.ToString(CallByName(it, "Name", CallType.Get), Globalization.CultureInfo.InvariantCulture)
                    If String.Equals(NormalizeStyleName(nm), NormalizeStyleName(styleName), StringComparison.OrdinalIgnoreCase) Then
                        Return it
                    End If
                Catch
                End Try
            Next
        Catch
        End Try
        Return Nothing
    End Function

    Private Shared Function NormalizeStyleName(s As String) As String
        If String.IsNullOrWhiteSpace(s) Then Return ""
        Dim t As String = s.Trim().Replace(ChrW(&HA0), " ")
        Do While t.Contains("  ")
            t = t.Replace("  ", " ")
        Loop
        Return t
    End Function
End Class
