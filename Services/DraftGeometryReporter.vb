Option Strict Off

Imports System.Globalization
Imports System.IO
Imports System.Text
Imports SolidEdgeDraft
Imports SolidEdgeFramework

''' <summary>
''' Genera el log de geometría existente del DFT (DVLines2d, DVArcs2d, DVCircles2d, etc.).
''' Pensado para invocarse JUSTO DESPUÉS de crear el DFT y antes de cerrarlo (no abre nada por sí mismo).
'''
''' Si la lectura de una propiedad COM falla, registra <c>[GEOM][WARN] method=... object=... index=...</c>
''' en el logger y continúa con el resto. NUNCA detiene el proceso.
''' </summary>
Public NotInheritable Class DraftGeometryReporter
    Private Sub New()
    End Sub

    Public Shared Sub ExportDraftGeometryLog(dft As DraftDocument, outFilePath As String, logger As Logger)
        If dft Is Nothing Then
            logger?.Log("[GEOM][ERR] DraftDocument Nothing.")
            Return
        End If
        If String.IsNullOrWhiteSpace(outFilePath) Then
            logger?.Log("[GEOM][ERR] outFilePath vacío.")
            Return
        End If

        Dim sb As New StringBuilder()
        sb.AppendLine("=== DFT GEOMETRY LOG ===")
        sb.AppendLine("File=" & SafeRead(Function() dft.FullName, "FullName", "DraftDocument", logger))
        sb.AppendLine("GeneratedUtc=" & DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture))

        Dim sheets As Sheets = Nothing
        Try
            sheets = dft.Sheets
        Catch ex As Exception
            logger?.LogComException("DraftDocument.Sheets", ex, "Sheets", "DraftDocument")
        End Try
        Dim sheetCount As Integer = SafeCount(sheets)
        sb.AppendLine("SheetCount=" & sheetCount.ToString(CultureInfo.InvariantCulture))

        For si As Integer = 1 To sheetCount
            Dim sh As Sheet = Nothing
            Try
                sh = CType(sheets.Item(si), Sheet)
            Catch ex As Exception
                logger?.LogComException("Sheets.Item", ex, "Item", "Sheets", si)
                Continue For
            End Try
            If sh Is Nothing Then Continue For

            Dim shName As String = SafeRead(Function() sh.Name, "Name", "Sheet", logger)
            sb.AppendLine("")
            sb.AppendLine($"[SHEET] idx={si} name={shName}")

            Dim views As DrawingViews = Nothing
            Try
                views = sh.DrawingViews
            Catch ex As Exception
                logger?.LogComException("Sheet.DrawingViews", ex, "DrawingViews", "Sheet", si)
            End Try
            Dim viewCount As Integer = SafeCount(views)
            sb.AppendLine($"  [SHEET][VIEWS] count={viewCount}")

            For vi As Integer = 1 To viewCount
                Dim dv As DrawingView = Nothing
                Try
                    dv = CType(views.Item(vi), DrawingView)
                Catch ex As Exception
                    logger?.LogComException("DrawingViews.Item", ex, "Item", "DrawingViews", vi)
                    Continue For
                End Try
                If dv Is Nothing Then Continue For

                Dim dvName As String = SafeRead(Function() dv.Name, "Name", "DrawingView", logger)
                Dim dvScale As String = SafeRead(Function() dv.ScaleFactor.ToString("0.######", CultureInfo.InvariantCulture), "ScaleFactor", "DrawingView", logger)
                Dim dvRange As String = TryReadRange(dv, logger)
                sb.AppendLine("")
                sb.AppendLine($"  [VIEW] idx={vi} name={dvName} scaleFactor={dvScale} range={dvRange}")

                DumpDVLines2d(sb, dv, vi, logger)
                DumpDVArcs2d(sb, dv, vi, logger)
                DumpDVCircles2d(sb, dv, vi, logger)
            Next
        Next

        Try
            Dim dir As String = Path.GetDirectoryName(outFilePath)
            If Not String.IsNullOrWhiteSpace(dir) AndAlso Not Directory.Exists(dir) Then Directory.CreateDirectory(dir)
            File.WriteAllText(outFilePath, sb.ToString(), Encoding.UTF8)
            logger?.Log("[GEOM][OK] " & outFilePath)
        Catch ex As Exception
            logger?.LogException("DraftGeometryReporter.WriteAllText", ex)
        End Try
    End Sub

    Private Shared Sub DumpDVLines2d(sb As StringBuilder, dv As DrawingView, viewIndex As Integer, logger As Logger)
        Dim col As Object = Nothing
        Try
            col = CallByName(dv, "DVLines2d", CallType.Get)
        Catch ex As Exception
            logger?.LogComException("DrawingView.DVLines2d", ex, "DVLines2d", "DrawingView", viewIndex)
        End Try
        Dim n As Integer = SafeCount(col)
        sb.AppendLine($"    [DVLines2d] count={n}")
        For i As Integer = 1 To n
            Dim ln As Object = Nothing
            Try
                ln = CallByName(col, "Item", CallType.Method, i)
            Catch ex As Exception
                logger?.LogComException("DVLines2d.Item", ex, "Item", "DVLines2d", i)
                Continue For
            End Try
            If ln Is Nothing Then Continue For

            Dim sx As Double = 0R, sy As Double = 0R, ex2 As Double = 0R, ey As Double = 0R
            Dim hasStart As Boolean = TryGet2dPoint(ln, "StartPoint", sx, sy)
            Dim hasEnd As Boolean = TryGet2dPoint(ln, "EndPoint", ex2, ey)
            Dim startStr As String = "?"
            If hasStart Then startStr = String.Format(CultureInfo.InvariantCulture, "({0:0.######},{1:0.######})", sx, sy)
            Dim endStr As String = "?"
            If hasEnd Then endStr = String.Format(CultureInfo.InvariantCulture, "({0:0.######},{1:0.######})", ex2, ey)
            Dim lenStr As String = SafeRead(Function() CallByName(ln, "Length", CallType.Get).ToString(), "Length", "DVLine2d", logger, i)
            Dim kpCount As String = SafeRead(Function() CallByName(ln, "KeyPointCount", CallType.Get).ToString(), "KeyPointCount", "DVLine2d", logger, i)
            Dim refIsNothing As Boolean = CheckReferenceIsNothing(ln, "DVLine2d", i, logger)
            Dim modelMemberIsNothing As Boolean = CheckModelMemberIsNothing(ln, "DVLine2d", i, logger)
            Dim edgeType As String = SafeRead(Function() CallByName(ln, "EdgeType", CallType.Get).ToString(), "EdgeType", "DVLine2d", logger, i)

            sb.AppendLine(String.Format(CultureInfo.InvariantCulture,
                "      [LINE] idx={0} start={1} end={2} length={3} keyPoints={4} refIsNothing={5} modelMemberIsNothing={6} edgeType={7}",
                i, startStr, endStr, lenStr, kpCount, refIsNothing, modelMemberIsNothing, edgeType))
        Next
    End Sub

    Private Shared Sub DumpDVArcs2d(sb As StringBuilder, dv As DrawingView, viewIndex As Integer, logger As Logger)
        Dim col As Object = Nothing
        Try
            col = CallByName(dv, "DVArcs2d", CallType.Get)
        Catch ex As Exception
            logger?.LogComException("DrawingView.DVArcs2d", ex, "DVArcs2d", "DrawingView", viewIndex)
        End Try
        Dim n As Integer = SafeCount(col)
        sb.AppendLine($"    [DVArcs2d] count={n}")
        For i As Integer = 1 To n
            Dim arc As Object = Nothing
            Try
                arc = CallByName(col, "Item", CallType.Method, i)
            Catch ex As Exception
                logger?.LogComException("DVArcs2d.Item", ex, "Item", "DVArcs2d", i)
                Continue For
            End Try
            If arc Is Nothing Then Continue For

            Dim cx As Double = 0R, cy As Double = 0R
            Dim hasCenter As Boolean = TryGet2dPoint(arc, "CenterPoint", cx, cy)
            Dim centerStr As String = "?"
            If hasCenter Then centerStr = String.Format(CultureInfo.InvariantCulture, "({0:0.######},{1:0.######})", cx, cy)
            Dim radius As String = SafeRead(Function() CallByName(arc, "Radius", CallType.Get).ToString(), "Radius", "DVArc2d", logger, i)
            Dim sweep As String = SafeRead(Function() CallByName(arc, "SweepAngle", CallType.Get).ToString(), "SweepAngle", "DVArc2d", logger, i)
            Dim start As String = SafeRead(Function() CallByName(arc, "StartAngle", CallType.Get).ToString(), "StartAngle", "DVArc2d", logger, i)
            Dim refIsNothing As Boolean = CheckReferenceIsNothing(arc, "DVArc2d", i, logger)
            sb.AppendLine(String.Format(CultureInfo.InvariantCulture,
                "      [ARC] idx={0} center={1} radius={2} startAngle={3} sweepAngle={4} refIsNothing={5}",
                i, centerStr, radius, start, sweep, refIsNothing))
        Next
    End Sub

    Private Shared Sub DumpDVCircles2d(sb As StringBuilder, dv As DrawingView, viewIndex As Integer, logger As Logger)
        Dim col As Object = Nothing
        Try
            col = CallByName(dv, "DVCircles2d", CallType.Get)
        Catch ex As Exception
            logger?.LogComException("DrawingView.DVCircles2d", ex, "DVCircles2d", "DrawingView", viewIndex)
        End Try
        Dim n As Integer = SafeCount(col)
        sb.AppendLine($"    [DVCircles2d] count={n}")
        For i As Integer = 1 To n
            Dim c As Object = Nothing
            Try
                c = CallByName(col, "Item", CallType.Method, i)
            Catch ex As Exception
                logger?.LogComException("DVCircles2d.Item", ex, "Item", "DVCircles2d", i)
                Continue For
            End Try
            If c Is Nothing Then Continue For

            Dim cx As Double = 0R, cy As Double = 0R
            Dim hasCenter As Boolean = TryGet2dPoint(c, "CenterPoint", cx, cy)
            Dim centerStr As String = "?"
            If hasCenter Then centerStr = String.Format(CultureInfo.InvariantCulture, "({0:0.######},{1:0.######})", cx, cy)
            Dim radius As String = SafeRead(Function() CallByName(c, "Radius", CallType.Get).ToString(), "Radius", "DVCircle2d", logger, i)
            Dim refIsNothing As Boolean = CheckReferenceIsNothing(c, "DVCircle2d", i, logger)
            sb.AppendLine(String.Format(CultureInfo.InvariantCulture,
                "      [CIRCLE] idx={0} center={1} radius={2} refIsNothing={3}",
                i, centerStr, radius, refIsNothing))
        Next
    End Sub

    ' --- helpers ---------------------------------------------------------

    Private Shared Function SafeCount(obj As Object) As Integer
        If obj Is Nothing Then Return 0
        Try
            Return CInt(CallByName(obj, "Count", CallType.Get))
        Catch
            Return 0
        End Try
    End Function

    Private Shared Function TryReadRange(dv As DrawingView, logger As Logger) As String
        If dv Is Nothing Then Return ""
        Try
            Dim x1 As Double = 0R, y1 As Double = 0R, x2 As Double = 0R, y2 As Double = 0R
            dv.Range(x1, y1, x2, y2)
            Return String.Format(CultureInfo.InvariantCulture,
                "({0:0.######},{1:0.######})-({2:0.######},{3:0.######})",
                x1, y1, x2, y2)
        Catch ex As Exception
            logger?.Log("[GEOM][WARN] method=Range object=DrawingView ex=" & ex.Message)
            Return ""
        End Try
    End Function

    Private Shared Function TryGet2dPoint(obj As Object, propertyName As String, ByRef x As Double, ByRef y As Double) As Boolean
        x = 0 : y = 0
        If obj Is Nothing Then Return False
        Try
            Dim pt As Object = CallByName(obj, propertyName, CallType.Get)
            If pt Is Nothing Then Return False
            x = CDbl(CallByName(pt, "X", CallType.Get))
            y = CDbl(CallByName(pt, "Y", CallType.Get))
            Return True
        Catch
            Return False
        End Try
    End Function

    Private Shared Function CheckReferenceIsNothing(obj As Object, objectKind As String, index As Integer, logger As Logger) As Boolean
        If obj Is Nothing Then Return True
        Try
            Dim ref As Object = CallByName(obj, "Reference", CallType.Get)
            Return ref Is Nothing
        Catch ex As Exception
            logger?.Log($"[GEOM][WARN] method=Reference object={objectKind} index={index} ex={ex.Message}")
            Return True
        End Try
    End Function

    Private Shared Function CheckModelMemberIsNothing(obj As Object, objectKind As String, index As Integer, logger As Logger) As Boolean
        If obj Is Nothing Then Return True
        Try
            Dim mm As Object = CallByName(obj, "ModelMember", CallType.Get)
            Return mm Is Nothing
        Catch ex As Exception
            logger?.Log($"[GEOM][WARN] method=ModelMember object={objectKind} index={index} ex={ex.Message}")
            Return True
        End Try
    End Function

    Private Shared Function SafeRead(reader As Func(Of String),
                                     methodName As String,
                                     objectKind As String,
                                     logger As Logger,
                                     Optional index As Integer = -1) As String
        If reader Is Nothing Then Return ""
        Try
            Return reader.Invoke()
        Catch ex As Exception
            Dim suffix As String = If(index >= 0, " index=" & index.ToString(CultureInfo.InvariantCulture), "")
            logger?.Log($"[GEOM][WARN] method={methodName} object={objectKind}{suffix} ex={ex.Message}")
            Return ""
        End Try
    End Function
End Class
