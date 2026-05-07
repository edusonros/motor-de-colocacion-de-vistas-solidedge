Option Strict Off

Imports System.Collections.Generic
Imports System.Globalization
Imports SolidEdgeDraft
Imports SolidEdgeFrameworkSupport

Friend NotInheritable Class DimensionStyleResolver
    Private Sub New()
    End Sub

    ''' <summary>Sólo <see cref="DraftDocument.DimensionStyles"/> en producción; sin barridos de colecciones inválidas.</summary>
    Public Shared Function ResolveDimensionStyle(draft As DraftDocument, styleName As String, log As DimensionLogger) As Object
        If draft Is Nothing OrElse String.IsNullOrWhiteSpace(styleName) Then
            log?.LogLine("[DIM][STYLE][RESOLVE] requested=" & If(styleName, "").Trim() & " found=False")
            Return Nothing
        End If

        Dim styles As Object = Nothing
        Try
            styles = CallByName(draft, "DimensionStyles", CallType.Get)
        Catch
            styles = Nothing
        End Try

        If styles Is Nothing Then
            log?.LogLine("[DIM][STYLE][RESOLVE] requested=" & If(styleName, "").Trim() & " found=False")
            Return Nothing
        End If

        Dim styleObj As Object = FindStyleByName(styles, styleName)
        Dim found As Boolean = styleObj IsNot Nothing
        log?.LogLine("[DIM][STYLE][RESOLVE] requested=" & If(styleName, "").Trim() & " found=" & found.ToString(CultureInfo.InvariantCulture))
        If GenerationEngineRuntime.DebugDiagnosticsMode Then
            Dim names As String = ListStyleNames(styles)
            If Not String.IsNullOrWhiteSpace(names) Then log?.LogLine("[DIM][STYLE][LIST] " & names)
        End If
        Return styleObj
    End Function

    ''' <summary>Aplicación de estilo por late-binding para evitar cast COM rígido (E_NOINTERFACE).</summary>
    Public Shared Function ApplyDimensionStyle(dimObj As Dimension, resolvedStyle As Object, log As DimensionLogger) As Boolean
        If dimObj Is Nothing OrElse resolvedStyle Is Nothing Then
            log?.LogLine("[DIM][STYLE][APPLY] dim=? style=? ok=False")
            Return False
        End If
        Dim styleName As String = ""
        Try
            styleName = Convert.ToString(CallByName(resolvedStyle, "Name", CallType.Get), CultureInfo.InvariantCulture)
        Catch
            styleName = ""
        End Try
        Try
            CallByName(dimObj, "Style", CallType.Let, resolvedStyle)
        Catch ex As Exception
            If GenerationEngineRuntime.DebugDiagnosticsMode Then
                log?.ComFail("ApplyDimensionStyle", "Dimension.Style=styleObject(late_binding)", ex)
            End If
            If Not String.IsNullOrWhiteSpace(styleName) Then
                Try
                    CallByName(dimObj, "StyleName", CallType.Let, styleName)
                Catch ex2 As Exception
                    If GenerationEngineRuntime.DebugDiagnosticsMode Then
                        log?.ComFail("ApplyDimensionStyle", "Dimension.StyleName=string", ex2)
                    End If
                    log?.LogLine("[DIM][STYLE][APPLY] dim=com_object style=? ok=False")
                    Return False
                End Try
            Else
                log?.LogLine("[DIM][STYLE][APPLY] dim=com_object style=? ok=False")
                Return False
            End If
        End Try
        Dim finalName As String = ReadDimensionStyleName(dimObj)
        Dim ok As Boolean = Not String.IsNullOrWhiteSpace(finalName) OrElse Not String.IsNullOrWhiteSpace(styleName)
        If String.IsNullOrWhiteSpace(finalName) Then finalName = styleName
        log?.LogLine("[DIM][STYLE][APPLY] dim=com_object style=" & finalName & " ok=" & ok.ToString(CultureInfo.InvariantCulture))
        Return ok
    End Function

    ''' <summary>Delega en <see cref="ResolveDimensionStyle"/> para compatibilidad.</summary>
    Public Shared Function ResolvePreferredStyle(draft As DraftDocument, sheet As Sheet, styleName As String, log As DimensionLogger) As Object
        Dim o = ResolveDimensionStyle(draft, styleName, log)
        If o Is Nothing AndAlso sheet IsNot Nothing AndAlso GenerationEngineRuntime.DebugDiagnosticsMode Then
            Try
                Dim shStyles As Object = CallByName(sheet, "DimensionStyles", CallType.Get)
                If shStyles IsNot Nothing Then
                    log?.LogLine("[DIM][STYLE][DEBUG][SHEET] DimensionStyles objeto presente en diagnóstico (no usado en producción).")
                End If
            Catch
            End Try
        End If
        Return o
    End Function

    Private Shared Function ReadDimensionStyleName(dimObj As Dimension) As String
        If dimObj Is Nothing Then Return ""
        Try
            Dim st As Object = CallByName(dimObj, "Style", CallType.Get)
            If st Is Nothing Then Return ""
            Return Convert.ToString(CallByName(st, "Name", CallType.Get), CultureInfo.InvariantCulture)
        Catch
            Return ""
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
                    Dim nm As String = Convert.ToString(CallByName(it, "Name", CallType.Get), CultureInfo.InvariantCulture)
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
                    Dim nm As String = Convert.ToString(CallByName(it, "Name", CallType.Get), CultureInfo.InvariantCulture)
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
