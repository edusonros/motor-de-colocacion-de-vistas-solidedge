Option Strict Off
Option Explicit On

Imports System.Collections.Generic
Imports System.Globalization
Imports System.Windows.Forms
Imports SolidEdgeDraft

Public Module ManualAdboGuidedLab

    Public Sub RunInteractive(app As SolidEdgeFramework.Application, draftDoc As DraftDocument, log As Action(Of String))
        If app Is Nothing OrElse draftDoc Is Nothing Then Exit Sub
        If log Is Nothing Then log = Sub(_m As String)
                                     End Sub

        log("[ADBO_GUIDED][START]")
        Dim sh As Sheet = Nothing
        Try : sh = draftDoc.ActiveSheet : Catch : sh = Nothing : End Try
        If sh Is Nothing Then
            log("[ADBO_GUIDED][ABORT] reason=no_active_sheet")
            Exit Sub
        End If

        Dim dims As Object = Nothing
        Try : dims = sh.Dimensions : Catch : dims = Nothing : End Try
        If dims Is Nothing Then
            log("[ADBO_GUIDED][ABORT] reason=no_dimensions_collection")
            Exit Sub
        End If

        Dim pickedObjects As List(Of Object) = CaptureFromSelectSet(
            app,
            draftDoc,
            sh,
            2,
            "Paso 1/2: Selecciona 2 objetos",
            "Selecciona en Solid Edge los 2 objetos para acotar (por ejemplo 2 líneas), y después vuelve aquí y pulsa Reintentar.",
            log)
        If pickedObjects Is Nothing OrElse pickedObjects.Count < 2 Then
            log("[ADBO_GUIDED][ABORT] reason=objects_not_selected")
            Exit Sub
        End If
        Dim obj1 As Object = pickedObjects(0)
        Dim obj2 As Object = pickedObjects(1)
        log("[ADBO_GUIDED][OBJECTS] o1=" & SafeTypeName(obj1) & " o2=" & SafeTypeName(obj2))

        Dim pickedPoints As List(Of Object) = CaptureFromSelectSet(
            app,
            draftDoc,
            sh,
            2,
            "Paso 2/2: Selecciona 2 puntos",
            "Ahora selecciona 2 puntos pertenecientes a esos objetos (uno por objeto), y después vuelve aquí y pulsa Reintentar.",
            log)
        If pickedPoints Is Nothing OrElse pickedPoints.Count < 2 Then
            log("[ADBO_GUIDED][ABORT] reason=points_not_selected")
            Exit Sub
        End If

        Dim x1 As Double = 0.0R, y1 As Double = 0.0R
        Dim x2 As Double = 0.0R, y2 As Double = 0.0R
        If Not TryExtractPointXY(pickedPoints(0), x1, y1) Then
            log("[ADBO_GUIDED][ABORT] reason=point1_invalid type=" & SafeTypeName(pickedPoints(0)))
            Exit Sub
        End If
        If Not TryExtractPointXY(pickedPoints(1), x2, y2) Then
            log("[ADBO_GUIDED][ABORT] reason=point2_invalid type=" & SafeTypeName(pickedPoints(1)))
            Exit Sub
        End If
        log("[ADBO_GUIDED][POINTS] p1=(" & F(x1) & "," & F(y1) & ") p2=(" & F(x2) & "," & F(y2) & ")")

        ' ADBO principal solicitado por el usuario.
        Dim d As Object = Nothing
        Try
            d = CallByName(dims, "AddDistanceBetweenObjects", CallType.Method, obj1, x1, y1, 0.0R, True, obj2, x2, y2, 0.0R, True)
            log("[ADBO_GUIDED][ADBO][OK] flags=True type=" & SafeTypeName(d) & " value=" & F(SafeToDouble(CallByName(d, "Value", CallType.Get))))
        Catch ex As Exception
            log("[ADBO_GUIDED][ADBO][FAIL] flags=True error=" & ex.Message)
        End Try

        ' También intenta variante flags=False para comparar.
        Try
            Dim d2 As Object = CallByName(dims, "AddDistanceBetweenObjects", CallType.Method, obj1, x1, y1, 0.0R, False, obj2, x2, y2, 0.0R, False)
            log("[ADBO_GUIDED][ADBO][OK] flags=False type=" & SafeTypeName(d2) & " value=" & F(SafeToDouble(CallByName(d2, "Value", CallType.Get))))
        Catch ex As Exception
            log("[ADBO_GUIDED][ADBO][FAIL] flags=False error=" & ex.Message)
        End Try

        Try : CallByName(draftDoc, "UpdateAll", CallType.Method, True) : Catch : End Try
        log("[ADBO_GUIDED][END]")
    End Sub

    Private Function CaptureFromSelectSet(app As SolidEdgeFramework.Application,
                                          draftDoc As DraftDocument,
                                          sh As Sheet,
                                          minCount As Integer,
                                          title As String,
                                          prompt As String,
                                          log As Action(Of String)) As List(Of Object)
        If MessageBox.Show(prompt, title, MessageBoxButtons.OKCancel, MessageBoxIcon.Information) <> DialogResult.OK Then Return Nothing

        Do
            Dim picked As List(Of Object) = ReadSelectionFromAnySource(app, draftDoc, sh, log)
            If picked.Count >= minCount Then Return picked
            Dim msg As String = "No hay suficientes elementos seleccionados (" & picked.Count.ToString(CultureInfo.InvariantCulture) &
                                "). Selecciona al menos " & minCount.ToString(CultureInfo.InvariantCulture) &
                                " en Solid Edge y pulsa Reintentar."
            Dim ans = MessageBox.Show(msg, title, MessageBoxButtons.RetryCancel, MessageBoxIcon.Warning)
            If ans <> DialogResult.Retry Then Exit Do
        Loop
        Return Nothing
    End Function

    Private Function ReadSelectionFromAnySource(app As SolidEdgeFramework.Application, draftDoc As DraftDocument, sh As Sheet, log As Action(Of String)) As List(Of Object)
        Dim out As New List(Of Object)()
        Dim seen As New HashSet(Of IntPtr)()

        Dim sources As New List(Of Tuple(Of String, Object))()
        Try : sources.Add(Tuple.Create("app.ActiveSelectSet", CallByName(app, "ActiveSelectSet", CallType.Get))) : Catch : End Try
        Try : sources.Add(Tuple.Create("doc.SelectSet", CallByName(draftDoc, "SelectSet", CallType.Get))) : Catch : End Try
        Try : sources.Add(Tuple.Create("sheet.SelectSet", CallByName(sh, "SelectSet", CallType.Get))) : Catch : End Try

        For Each src In sources
            Dim ss As Object = src.Item2
            Dim c As Integer = 0
            Try : c = Convert.ToInt32(CallByName(ss, "Count", CallType.Get), CultureInfo.InvariantCulture) : Catch : c = 0 : End Try
            Try : log?.Invoke("[ADBO_GUIDED][SELECTSET] source=" & src.Item1 & " count=" & c.ToString(CultureInfo.InvariantCulture)) : Catch : End Try
            For i As Integer = 1 To c
                Try
                    Dim o As Object = CallByName(ss, "Item", CallType.Method, i)
                    If o Is Nothing Then Continue For
                    Dim key As IntPtr = IntPtr.Zero
                    Try
                        key = Runtime.InteropServices.Marshal.GetIUnknownForObject(o)
                        If key <> IntPtr.Zero AndAlso Not seen.Contains(key) Then
                            seen.Add(key)
                            out.Add(o)
                        End If
                    Finally
                        If key <> IntPtr.Zero Then
                            Try : Runtime.InteropServices.Marshal.Release(key) : Catch : End Try
                        End If
                    End Try
                Catch
                End Try
            Next
        Next
        Return out
    End Function

    Private Function TryExtractPointXY(ptObj As Object, ByRef x As Double, ByRef y As Double) As Boolean
        x = 0.0R : y = 0.0R
        If ptObj Is Nothing Then Return False

        Try
            x = SafeToDouble(CallByName(ptObj, "X", CallType.Get))
            y = SafeToDouble(CallByName(ptObj, "Y", CallType.Get))
            Return True
        Catch
        End Try

        Try
            Dim px As Double = 0.0R, py As Double = 0.0R
            CallByName(ptObj, "GetPoint", CallType.Method, px, py)
            x = px : y = py
            Return True
        Catch
        End Try

        Try
            Dim p As Object = CallByName(ptObj, "Point", CallType.Get)
            If p IsNot Nothing Then
                x = SafeToDouble(CallByName(p, "X", CallType.Get))
                y = SafeToDouble(CallByName(p, "Y", CallType.Get))
                Return True
            End If
        Catch
        End Try

        Return False
    End Function

    Private Function SafeToDouble(v As Object) As Double
        If v Is Nothing Then Return 0.0R
        Try
            Return Convert.ToDouble(v, CultureInfo.InvariantCulture)
        Catch
            Return 0.0R
        End Try
    End Function

    Private Function SafeTypeName(o As Object) As String
        Try : Return If(o Is Nothing, "Nothing", o.GetType().Name) : Catch : Return "?" : End Try
    End Function

    Private Function F(v As Double) As String
        Return v.ToString("0.###############", CultureInfo.InvariantCulture)
    End Function

End Module

