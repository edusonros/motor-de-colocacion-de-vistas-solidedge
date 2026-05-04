Option Strict Off

Imports System.Globalization
Imports SolidEdgeDraft

Friend NotInheritable Class DraftViewCollector
    Private Sub New()
    End Sub

    Public Shared Function CollectOrthogonalViews(sheet As Sheet, log As DimensionLogger) As List(Of DrawingView)
        Dim result As New List(Of DrawingView)()
        If sheet Is Nothing Then Return result

        Dim count As Integer = 0
        Try
            count = sheet.DrawingViews.Count
        Catch ex As Exception
            log?.ComFail("Sheet.DrawingViews.Count", "DrawingViews", ex)
            Return result
        End Try
        log?.LogLine("[DIM][VIEW][FOUND] total=" & count.ToString(CultureInfo.InvariantCulture))

        For i As Integer = 1 To count
            Dim dv As DrawingView = Nothing
            Try
                dv = CType(sheet.DrawingViews.Item(i), DrawingView)
            Catch ex As Exception
                log?.ComFail("DrawingViews.Item", "DrawingViews", ex)
                Continue For
            End Try
            If dv Is Nothing Then Continue For

            Dim name As String = SafeToString(CallByNameSafe(dv, "Name"))
            Dim ori As String = SafeToString(CallByNameSafe(dv, "ViewOrientation"))
            Dim dvt As String = SafeToString(CallByNameSafe(dv, "DrawingViewType"))
            Dim dvtNum As Integer = SafeToInt(CallByNameSafe(dv, "DrawingViewType"))
            Dim isIso As Boolean =
                ori.IndexOf("iso", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                ori.IndexOf("topfrontright", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                dvt.IndexOf("iso", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                (dvtNum > 0 AndAlso dvtNum <> 1)

            If isIso Then
                log?.LogLine("[DIM][VIEW][SKIP] idx=" & i.ToString(CultureInfo.InvariantCulture) & " name=" & name & " reason=isometric_or_non_orthogonal type=" & dvtNum.ToString(CultureInfo.InvariantCulture))
                Continue For
            End If

            result.Add(dv)
            log?.LogLine("[DIM][VIEW][PROCESS] idx=" & i.ToString(CultureInfo.InvariantCulture) & " name=" & name & " orientation=" & ori)
        Next

        Return result
    End Function

    Private Shared Function CallByNameSafe(obj As Object, member As String) As Object
        Try
            Return CallByName(obj, member, CallType.Get)
        Catch
            Return Nothing
        End Try
    End Function

    Private Shared Function SafeToString(v As Object) As String
        If v Is Nothing Then Return ""
        Try
            Return Convert.ToString(v, CultureInfo.InvariantCulture)
        Catch
            Return ""
        End Try
    End Function

    Private Shared Function SafeToInt(v As Object) As Integer
        If v Is Nothing Then Return 0
        Try
            Return Convert.ToInt32(v, CultureInfo.InvariantCulture)
        Catch
            Return 0
        End Try
    End Function
End Class
