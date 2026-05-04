Option Strict Off

Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Globalization
Imports System.Linq
Imports SolidEdgeDraft

Friend NotInheritable Class UniqueDvAutoDimensioningEngine
    Private Const RequiredStyleName As String = "U3,5"

    Private Sub New()
    End Sub

    Public Shared Sub Run(draft As DraftDocument, log As DimensionLogger, baseLogger As Logger, Optional norm As DimensioningNormConfig = Nothing, Optional protectedZones As IList(Of ProtectedZone2D) = Nothing)
        Dim sw As Stopwatch = Stopwatch.StartNew()
        If norm Is Nothing Then norm = DimensioningNormConfig.DefaultConfig()
        log?.LogLine("[DIM][ENTER] único motor DV*2d")
        log?.LogLine("[DIM][UNE] dedupe_cleanup=" & norm.EnableDuplicateDimensionCleanup.ToString(CultureInfo.InvariantCulture) &
                     " dedupe_cross_view=" & norm.UneDedupeNominalAcrossOrthogonalViews.ToString(CultureInfo.InvariantCulture) &
                     " keep_intentional_dupes=" & norm.KeepIntentionalDuplicateDimensions.ToString(CultureInfo.InvariantCulture) &
                     " largest_first=" & norm.UneProcessLargestOrthogonalViewFirst.ToString(CultureInfo.InvariantCulture) &
                     " iso129=" & norm.EnableISO129Rules.ToString(CultureInfo.InvariantCulture))

        Dim ctx As SolidEdgeContext = Nothing
        If Not SolidEdgeContext.TryCreate(draft, log, ctx) Then Return
        log?.LogLine("[DIM][DOC] name=" & SafeToString(CallByNameSafe(ctx.Draft, "Name")))
        log?.LogLine("[DIM][SHEET] name=" & SafeToString(CallByNameSafe(ctx.Sheet, "Name")))

        Dim styleObj As Object = DimensionStyleResolver.ResolvePreferredStyle(ctx.Draft, ctx.Sheet, RequiredStyleName, log)
        Dim resolvedStyleName As String = ReadStyleName(styleObj)

        Dim views As List(Of DrawingView) = DraftViewCollector.CollectOrthogonalViews(ctx.Sheet, log)
        If views.Count = 0 Then
            log?.LogLine("[DIM][SUMMARY][DOC] views=0 processed=0 skipped=0 created=0 errors=0")
            Return
        End If

        Dim crossKeys As HashSet(Of String) = Nothing
        If norm.EnableDuplicateDimensionCleanup AndAlso norm.UneDedupeNominalAcrossOrthogonalViews AndAlso
           norm.AvoidDuplicateDimensions AndAlso Not norm.KeepIntentionalDuplicateDimensions Then
            crossKeys = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        End If

        Dim workList As New List(Of DrawingViewGeometryInfo)()
        For i As Integer = 0 To views.Count - 1
            Dim info As DrawingViewGeometryInfo = DrawingViewGeometryReader.Read(views(i), i + 1, log)
            If info Is Nothing Then
                Continue For
            End If
            workList.Add(info)
        Next

        If norm.UneProcessLargestOrthogonalViewFirst AndAlso workList.Count > 1 Then
            workList = workList.OrderByDescending(Function(inf) inf.Box.Width * inf.Box.Height).ToList()
            log?.LogLine("[DIM][UNE][ORDER] vistas_ortogonales por area en hoja (mayor primero)")
        End If

        Dim useTargetRef As Boolean = String.Equals(norm.DimensionCreationMode, DimensioningNormConfig.ModeTargetReference, StringComparison.OrdinalIgnoreCase)
        If useTargetRef Then
            ReferenceDrawingDimensioningService.Run(ctx, workList, styleObj, norm, log, baseLogger, protectedZones)
            sw.Stop()
            log?.LogLine("[DIM][SUMMARY][DOC] mode=TargetDrawingLikeReference viewsProcessed=" & workList.Count.ToString(CultureInfo.InvariantCulture) &
                         " ms=" & sw.ElapsedMilliseconds.ToString(CultureInfo.InvariantCulture))
            baseLogger?.Log("[DIM] Motor referencia finalizado.")
            Return
        End If

        Dim totalLines As Integer = 0
        Dim totalArcs As Integer = 0
        Dim totalCircles As Integer = 0
        Dim totalLineStrings As Integer = 0
        Dim totalSplines As Integer = 0
        Dim totalPoints As Integer = 0
        Dim processed As Integer = 0
        Dim skipped As Integer = views.Count - workList.Count
        Dim createdLinear As Integer = 0
        Dim createdRadial As Integer = 0
        Dim discarded As Integer = 0
        Dim errors As Integer = 0
        For Each info As DrawingViewGeometryInfo In workList
            processed += 1
            totalLines += info.CountLines
            totalArcs += info.CountArcs
            totalCircles += info.CountCircles
            totalLineStrings += info.CountLineStrings
            totalSplines += info.CountSplines
            totalPoints += info.CountPoints

            Dim createdView As Integer = DrawingViewDimensionCreator.CreateDimensionsByViewSweep(
                ctx.Draft, ctx.Sheet, info, styleObj, log, createdLinear, createdRadial, discarded, errors, crossKeys, norm)

            log?.LogLine("[DIM][SUMMARY][VIEW] idx=" & info.ViewIndex.ToString(CultureInfo.InvariantCulture) &
                         " name=" & info.ViewName &
                         " lines=" & info.CountLines.ToString(CultureInfo.InvariantCulture) &
                         " arcs=" & info.CountArcs.ToString(CultureInfo.InvariantCulture) &
                         " circles=" & info.CountCircles.ToString(CultureInfo.InvariantCulture) &
                         " created=" & createdView.ToString(CultureInfo.InvariantCulture))
        Next

        If norm.EnableISO129Rules Then
            Dim arrangeViews As New List(Of DrawingView)()
            For Each inf As DrawingViewGeometryInfo In workList
                If inf IsNot Nothing AndAlso inf.View IsNot Nothing Then arrangeViews.Add(inf.View)
            Next
            Dim zoneBoxes As List(Of BoundingBox2D) = Nothing
            If protectedZones IsNot Nothing AndAlso protectedZones.Count > 0 Then
                zoneBoxes = New List(Of BoundingBox2D)()
                For Each z In protectedZones
                    If z IsNot Nothing Then zoneBoxes.Add(z.ToBoundingBox2D())
                Next
            End If
            Une129ArrangeExistingDimensions.OrdenarCotasExistentesUNE129(
                ctx.Draft, ctx.Sheet, arrangeViews, zoneBoxes, norm,
                Sub(m) log?.LogLine(m))
        End If

        sw.Stop()
        log?.LogLine("[DIM][SUMMARY][DOC] viewsTotal=" & views.Count.ToString(CultureInfo.InvariantCulture) &
                     " viewsProcessed=" & processed.ToString(CultureInfo.InvariantCulture) &
                     " viewsSkipped=" & skipped.ToString(CultureInfo.InvariantCulture) &
                     " dvLines=" & totalLines.ToString(CultureInfo.InvariantCulture) &
                     " dvCircles=" & totalCircles.ToString(CultureInfo.InvariantCulture) &
                     " dvArcs=" & totalArcs.ToString(CultureInfo.InvariantCulture) &
                     " dvLineStrings=" & totalLineStrings.ToString(CultureInfo.InvariantCulture) &
                     " dvSplines=" & totalSplines.ToString(CultureInfo.InvariantCulture) &
                     " dvPoints=" & totalPoints.ToString(CultureInfo.InvariantCulture) &
                     " linearCreated=" & createdLinear.ToString(CultureInfo.InvariantCulture) &
                     " radialCreated=" & createdRadial.ToString(CultureInfo.InvariantCulture) &
                     " totalCreated=" & (createdLinear + createdRadial).ToString(CultureInfo.InvariantCulture) &
                     " discarded=" & discarded.ToString(CultureInfo.InvariantCulture) &
                     " errors=" & errors.ToString(CultureInfo.InvariantCulture) &
                     " ms=" & sw.ElapsedMilliseconds.ToString(CultureInfo.InvariantCulture))

        baseLogger?.Log("[DIM] Motor único DV*2d finalizado.")
    End Sub

    Private Shared Function ReadStyleName(styleObj As Object) As String
        If styleObj Is Nothing Then Return ""
        Try
            Return Convert.ToString(CallByName(styleObj, "Name", CallType.Get), CultureInfo.InvariantCulture)
        Catch
            Return ""
        End Try
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
End Class
