Option Strict Off

Imports System
Imports System.ComponentModel
Imports System.Globalization
Imports System.Linq
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Text
Imports Microsoft.VisualBasic
Imports SolidEdgeDraft
Imports SolidEdgeFrameworkSupport

''' <summary>
''' Volcado de diagnóstico post-acotado: propiedades legibles (y una muestra de métodos CLR declarados) sobre
''' <see cref="DrawingView"/>, colecciones DV*2d y <see cref="Dimension"/>. No parsea <c>F:\sesdk_extraido</c>;
''' la ayuda CHM/HTML allí describe el mismo contrato COM que aquí se inspecciona en vivo.
''' </summary>
Public NotInheritable Class SesdkDraftGeometryDimensionIntrospection

    Private Const MaxPropsPerObject As Integer = 72
    Private Const MaxMethodsPerObject As Integer = 36
    Private Const MaxStringSample As Integer = 120
    Private Const MaxDimensionsDetailed As Integer = 3

    Private Sub New()
    End Sub

    ''' <summary>
    ''' Ejecuta el volcado si <see cref="DimensioningNormConfig.EnableSesdkPostDimensionIntrospection"/> es True
    ''' o si existe <c>SE_SESDK_INTROSPECT=1</c> (o "true").
    ''' </summary>
    Public Shared Sub RunIfRequested(draft As DraftDocument, norm As DimensioningNormConfig, appLogger As Logger)
        Dim cfg As DimensioningNormConfig = If(norm, DimensioningNormConfig.DefaultConfig())
        Dim env As String = ""
        Try
            env = Environment.GetEnvironmentVariable("SE_SESDK_INTROSPECT")
        Catch
            env = ""
        End Try
        Dim want As Boolean = cfg.EnableSesdkPostDimensionIntrospection OrElse
            String.Equals(env, "1", StringComparison.OrdinalIgnoreCase) OrElse
            String.Equals(env, "true", StringComparison.OrdinalIgnoreCase)
        If Not want Then Return
        If appLogger Is Nothing Then Return
        If draft Is Nothing Then
            appLogger.Log("[SESDK_PROBE][SKIP] DraftDocument Nothing.")
            Return
        End If

        Try
            Dim sb As New StringBuilder()
            sb.AppendLine("[SESDK_PROBE][START] Introspección post-acotado (geometría DV + cotas). Ayuda de referencia: F:\sesdk_extraido\sesdk.chm")

            Dim sh As Sheet = Nothing
            Try
                sh = draft.ActiveSheet
            Catch
                sh = Nothing
            End Try
            If sh Is Nothing Then
                appLogger.Log(sb.ToString() & "[SESDK_PROBE][SKIP] ActiveSheet Nothing.")
                Return
            End If

            AppendLine(sb, "[SESDK_PROBE][SHEET] name=" & SafeName(sh))

            Dim dvs As DrawingViews = Nothing
            Try
                dvs = sh.DrawingViews
            Catch
                dvs = Nothing
            End Try
            If dvs Is Nothing Then
                appLogger.Log(sb.ToString() & "[SESDK_PROBE][SKIP] DrawingViews Nothing.")
                Return
            End If

            Dim nViews As Integer = 0
            Try
                nViews = dvs.Count
            Catch
                nViews = 0
            End Try
            AppendLine(sb, "[SESDK_PROBE][VIEWS] count=" & nViews.ToString(CultureInfo.InvariantCulture))

            For vi As Integer = 1 To Math.Min(nViews, 4)
                Dim dv As DrawingView = Nothing
                Try
                    dv = CType(dvs.Item(vi), DrawingView)
                Catch
                    dv = Nothing
                End Try
                If dv Is Nothing Then Continue For
                AppendLine(sb, "[SESDK_PROBE][VIEW] index=" & vi.ToString(CultureInfo.InvariantCulture) & " name=" & SafeName(dv))
                DumpComObjectProps(sb, "DrawingView", dv, MaxPropsPerObject)
                DumpDeclaredMethods(sb, "DrawingView", dv, MaxMethodsPerObject)

                TrySampleCollectionItem(sb, dv, "DVLines2d", "DVLine2d", vi)
                TrySampleCollectionItem(sb, dv, "DVArcs2d", "DVArc2d", vi)
                TrySampleCollectionItem(sb, dv, "DVCircles2d", "DVCircle2d", vi)
                TrySampleCollectionItem(sb, dv, "DVLineStrings2d", "DVLineString2d", vi)
                TrySampleCollectionItem(sb, dv, "DVBSplineCurves2d", "DVBSplineCurve2d", vi)
            Next

            Dim dims As Dimensions = Nothing
            Try
                dims = CType(sh.Dimensions, Dimensions)
            Catch
                dims = Nothing
            End Try
            If dims IsNot Nothing Then
                Dim nd As Integer = 0
                Try
                    nd = dims.Count
                Catch
                    nd = 0
                End Try
                AppendLine(sb, "[SESDK_PROBE][DIMENSIONS] count=" & nd.ToString(CultureInfo.InvariantCulture))
                For di As Integer = 1 To Math.Min(nd, 12)
                    Dim dm As Dimension = Nothing
                    Try
                        dm = CType(dims.Item(di), Dimension)
                    Catch
                        dm = Nothing
                    End Try
                    If dm Is Nothing Then Continue For
                    If di <= MaxDimensionsDetailed Then
                        AppendLine(sb, "[SESDK_PROBE][DIM] index=" & di.ToString(CultureInfo.InvariantCulture))
                        DumpComObjectProps(sb, "Dimension", dm, MaxPropsPerObject)
                        DumpDeclaredMethods(sb, "Dimension", dm, MaxMethodsPerObject)
                    Else
                        AppendLine(sb, "[SESDK_PROBE][DIM] index=" & di.ToString(CultureInfo.InvariantCulture) & " typename=" & TypeName(dm))
                    End If
                Next
            Else
                AppendLine(sb, "[SESDK_PROBE][DIMENSIONS] (no disponible)")
            End If

            sb.AppendLine("[SESDK_PROBE][END]")
            appLogger.Log(sb.ToString().TrimEnd())
        Catch ex As Exception
            appLogger.Log("[SESDK_PROBE][FATAL] " & ex.ToString())
        End Try
    End Sub

    Private Shared Sub AppendLine(sb As StringBuilder, line As String)
        If sb Is Nothing Then Return
        sb.AppendLine(line)
    End Sub

    Private Shared Function SafeName(o As Object) As String
        If o Is Nothing Then Return ""
        Try
            Return Convert.ToString(CallByName(o, "Name", CallType.Get), CultureInfo.InvariantCulture)
        Catch
            Return TypeName(o)
        End Try
    End Function

    Private Shared Sub TrySampleCollectionItem(sb As StringBuilder, view As DrawingView, collectionProp As String, labelShort As String, viewIndex As Integer)
        If sb Is Nothing OrElse view Is Nothing Then Return
        Dim col As Object = Nothing
        Try
            col = CallByName(view, collectionProp, CallType.Get)
        Catch ex As Exception
            AppendLine(sb, "[SESDK_PROBE][" & labelShort & "][view " & viewIndex.ToString(CultureInfo.InvariantCulture) & "] no_col err=" & ex.Message)
            Return
        End Try
        If col Is Nothing Then
            AppendLine(sb, "[SESDK_PROBE][" & labelShort & "][view " & viewIndex.ToString(CultureInfo.InvariantCulture) & "] col=null")
            Return
        End If
        Dim cnt As Integer = 0
        Try
            cnt = CInt(CallByName(col, "Count", CallType.Get))
        Catch
            cnt = 0
        End Try
        AppendLine(sb, "[SESDK_PROBE][" & labelShort & "][view " & viewIndex.ToString(CultureInfo.InvariantCulture) & "] count=" & cnt.ToString(CultureInfo.InvariantCulture))
        If cnt <= 0 Then Return
        Dim it As Object = Nothing
        Try
            it = CallByName(col, "Item", CallType.Method, 1)
        Catch ex As Exception
            AppendLine(sb, "[SESDK_PROBE][" & labelShort & "][Item(1)] err=" & ex.Message)
            Return
        End Try
        If it Is Nothing Then Return
        DumpComObjectProps(sb, labelShort & ".Item(1)", it, MaxPropsPerObject)
        DumpDeclaredMethods(sb, labelShort & ".Item(1)", it, MaxMethodsPerObject)
    End Sub

    Private Shared Sub DumpComObjectProps(sb As StringBuilder, label As String, obj As Object, maxProps As Integer)
        If sb Is Nothing OrElse obj Is Nothing Then Return
        AppendLine(sb, "[SESDK_PROBE][PROPS] " & label & " TypeName=" & TypeName(obj))
        Dim n As Integer = 0
        Try
            For Each pd As PropertyDescriptor In TypeDescriptor.GetProperties(obj)
                If n >= maxProps Then Exit For
                If pd Is Nothing Then Continue For
                Dim pname As String = pd.Name
                If String.IsNullOrWhiteSpace(pname) Then Continue For
                Try
                    Dim v As Object = pd.GetValue(obj)
                    AppendLine(sb, "  " & pname & "=" & FormatValueSample(v))
                Catch ex As Exception
                    AppendLine(sb, "  " & pname & "=<get " & ex.Message & ">")
                End Try
                n += 1
            Next
        Catch ex As Exception
            AppendLine(sb, "[SESDK_PROBE][PROPS_ERR] " & label & " " & ex.Message)
        End Try
    End Sub

    Private Shared Sub DumpDeclaredMethods(sb As StringBuilder, label As String, obj As Object, maxMethods As Integer)
        If sb Is Nothing OrElse obj Is Nothing Then Return
        Try
            Dim t As Type = obj.GetType()
            AppendLine(sb, "[SESDK_PROBE][METHODS] " & label & " CLR=" & t.FullName)
            Dim n As Integer = 0
            Dim arr = t.GetMethods(BindingFlags.Public Or BindingFlags.Instance Or BindingFlags.DeclaredOnly).
                Where(Function(m) m IsNot Nothing AndAlso Not m.IsSpecialName).
                OrderBy(Function(m) m.Name, StringComparer.OrdinalIgnoreCase).
                ToArray()
            For Each m As MethodInfo In arr
                If n >= maxMethods Then Exit For
                Dim ps = m.GetParameters()
                Dim sig As String = String.Join(", ", ps.Select(Function(p) p.ParameterType.Name))
                AppendLine(sb, "  " & m.Name & "(" & sig & ")")
                n += 1
            Next
        Catch ex As Exception
            AppendLine(sb, "[SESDK_PROBE][METHODS_ERR] " & label & " " & ex.Message)
        End Try
    End Sub

    Private Shared Function FormatValueSample(v As Object) As String
        If v Is Nothing Then Return "<null>"
        Try
            If TypeOf v Is String Then
                Dim s As String = CStr(v)
                If s.Length > MaxStringSample Then Return s.Substring(0, MaxStringSample) & "…"
                Return s
            End If
            If TypeOf v Is Double OrElse TypeOf v Is Single Then
                Return Convert.ToString(v, CultureInfo.InvariantCulture)
            End If
            If TypeOf v Is Integer OrElse TypeOf v Is Long OrElse TypeOf v Is Short OrElse TypeOf v Is Byte Then
                Return Convert.ToString(v, CultureInfo.InvariantCulture)
            End If
            If TypeOf v Is Boolean Then
                Return If(CBool(v), "True", "False")
            End If
            If Marshal.IsComObject(v) Then
                Return "<COM:" & TypeName(v) & ">"
            End If
            Dim s2 As String = Convert.ToString(v, CultureInfo.InvariantCulture)
            If s2.Length > MaxStringSample Then Return s2.Substring(0, MaxStringSample) & "…"
            Return s2
        Catch ex As Exception
            Return "<fmt " & ex.Message & ">"
        End Try
    End Function
End Class
