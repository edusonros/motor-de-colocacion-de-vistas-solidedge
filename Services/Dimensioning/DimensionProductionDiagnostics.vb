Option Strict Off

Imports System.Globalization
Imports SolidEdgeDraft
Imports SolidEdgeFrameworkSupport
Imports FrameworkDimension = SolidEdgeFrameworkSupport.Dimension

''' <summary>Métricas de una corrida para el bloque final <c>[SUMMARY]</c> en modo producción.</summary>
Friend NotInheritable Class DimensionProductionRunSummary
    Friend Shared ViewsPlanned As Integer
    Friend Shared LayoutOk As Boolean = True
    Friend Shared PartsListCreated As Boolean
    Friend Shared PartsListRows As Integer
    Friend Shared PartsListCols As Integer
    Friend Shared DimsCreated As Integer
    Friend Shared DimsConnectedOk As Integer
    Friend Shared DimsVisibleOk As Integer
    Friend Shared DimsFailed As Integer
    Friend Shared StyleAppliedName As String = ""
    Friend Shared ResultOk As Boolean = True

    Friend Shared Sub RecordPartsList(created As Boolean, rows As Integer, cols As Integer)
        PartsListCreated = created
        PartsListRows = rows
        PartsListCols = cols
    End Sub

    ''' <summary>
    ''' Tras <c>ProductionDvRefCleanDimensionEngine.Run</c>: alinea <c>[SUMMARY][DIMS]</c> con el resumen interno <c>[PRODDIM][SUMMARY]</c>.
    ''' </summary>
    Friend Shared Sub RecordProddimRun(created As Integer, kept As Integer, deleted As Integer, bothAxesSuccess As Boolean, Optional resolvedStyleName As String = Nothing)
        DimsCreated = Math.Max(0, kept)
        DimsConnectedOk = Math.Max(0, kept)
        DimsVisibleOk = Math.Max(0, kept)
        DimsFailed = Math.Max(0, deleted)
        If Not bothAxesSuccess Then ResultOk = False
        If Not String.IsNullOrWhiteSpace(resolvedStyleName) Then StyleAppliedName = resolvedStyleName
    End Sub

    Friend Shared Sub Reset()
        ViewsPlanned = 0
        LayoutOk = True
        PartsListCreated = False
        PartsListRows = 0
        PartsListCols = 0
        DimsCreated = 0
        DimsConnectedOk = 0
        DimsVisibleOk = 0
        DimsFailed = 0
        StyleAppliedName = ""
        ResultOk = True
    End Sub
End Class

''' <summary>Validación única de cota (creación + estado final).</summary>
Friend Class DimensionDiagnosticResult
    Public Property Name As String
    Public Property CreatedOk As Boolean
    Public Property ValueOk As Boolean
    Public Property VisibleOk As Boolean
    Public Property MaterializedOk As Boolean
    Public Property ConnectedOk As Boolean
    Public Property FinalValue As Double
    Public Property ExpectedValue As Double
    Public Property Delta As Double
    Public Property RangeMinX As Double
    Public Property RangeMinY As Double
    Public Property RangeMaxX As Double
    Public Property RangeMaxY As Double
    Public Property StyleName As String
    Public Property FinalResult As Boolean
End Class

Friend NotInheritable Class DimensionProductionDiagnostics
    Private Const TolValue As Double = 0.004R

    Private Sub New()
    End Sub

    Friend Shared Function Diagnose(dimObj As FrameworkDimension, draft As DraftDocument, sheet As Sheet, dv As DrawingView,
                                    name As String, Optional expectedNominalM As Nullable(Of Double) = Nothing) As DimensionDiagnosticResult
        Dim r As New DimensionDiagnosticResult With {.Name = name, .CreatedOk = dimObj IsNot Nothing}
        If dimObj Is Nothing Then
            r.FinalResult = False
            Return r
        End If

        Try
            Dim v As Double = Convert.ToDouble(CallByName(dimObj, "Value", CallType.Get), CultureInfo.InvariantCulture)
            r.FinalValue = v
        Catch
            r.FinalValue = Double.NaN
        End Try

        If expectedNominalM.HasValue AndAlso expectedNominalM.Value > 1.0E-9 Then
            r.ExpectedValue = expectedNominalM.Value
            r.Delta = r.FinalValue - r.ExpectedValue
            r.ValueOk = (Not Double.IsNaN(r.FinalValue)) AndAlso Math.Abs(r.Delta) <= TolValue
        Else
            r.ExpectedValue = 0R
            r.Delta = 0R
            r.ValueOk = True
        End If

        Try
            Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
            dimObj.Range(x1, y1, x2, y2)
            r.RangeMinX = Math.Min(x1, x2)
            r.RangeMaxX = Math.Max(x1, x2)
            r.RangeMinY = Math.Min(y1, y2)
            r.RangeMaxY = Math.Max(y1, y2)
            Dim w = Math.Abs(x2 - x1)
            Dim h = Math.Abs(y2 - y1)
            r.MaterializedOk = (w > 1.0E-8 OrElse h > 1.0E-8)
            r.VisibleOk = IsLikelyInsideSheetRough(sheet, r.RangeMinX, r.RangeMinY, r.RangeMaxX, r.RangeMaxY)
        Catch
            r.MaterializedOk = False
            r.VisibleOk = False
        End Try

        r.StyleName = SafeReadDimensionStyleName(dimObj)
        r.ConnectedOk = If(draft Is Nothing OrElse dv Is Nothing, True, DrawingViewDimensionCreator.IsDimensionConnectedToViewForDiagnostics(draft, sheet, dv, dimObj))

        r.FinalResult = r.CreatedOk AndAlso r.MaterializedOk AndAlso r.VisibleOk AndAlso r.ConnectedOk AndAlso r.ValueOk
        Return r
    End Function

    Private Shared Function IsLikelyInsideSheetRough(sheet As Sheet, mnX As Double, mnY As Double, mxX As Double, mxY As Double) As Boolean
        If sheet Is Nothing Then Return True
        Try
            Dim w As Double = 0, h As Double = 0
            Dim su = sheet.SheetSetup
            If su IsNot Nothing Then
                w = su.SheetWidth
                h = su.SheetHeight
            End If
            If w <= 1.0E-6 OrElse h <= 1.0E-6 Then Return True
            Dim mrg As Double = 0.08R ' margen laxo sólo anti-falso positivo fuera-hoja grosera
            Return mnX >= -mrg AndAlso mnY >= -mrg AndAlso mxX <= w + mrg AndAlso mxY <= h + mrg
        Catch
            Return True
        End Try
    End Function

    Private Shared Function SafeReadDimensionStyleName(d As FrameworkDimension) As String
        If d Is Nothing Then Return ""
        Try
            Dim st As Object = CallByName(d, "Style", CallType.Get)
            If st Is Nothing Then Return ""
            Return Convert.ToString(CallByName(st, "Name", CallType.Get), CultureInfo.InvariantCulture)
        Catch
            Return ""
        End Try
    End Function

    Friend Shared Sub LogValidate(tag As String, res As DimensionDiagnosticResult, log As DimensionLogger)
        If log Is Nothing OrElse res Is Nothing Then Return
        log.LogLine(String.Format(CultureInfo.InvariantCulture,
            "[DIM][VALIDATE][{0}] name={1} visible={2} connected={3} value={4:0.#####} delta={5:0.#####} materialized={6} ok={7}",
            tag,
            If(res.Name, ""),
            res.VisibleOk,
            res.ConnectedOk,
            res.FinalValue,
            res.Delta,
            res.MaterializedOk,
            res.FinalResult))
    End Sub
End Class
