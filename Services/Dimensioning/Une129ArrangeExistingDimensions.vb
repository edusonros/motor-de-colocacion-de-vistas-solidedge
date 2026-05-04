Option Strict Off

Imports System.Collections.Generic
Imports System.Globalization
Imports System.Linq
Imports SolidEdgeDraft
Imports SolidEdgeFrameworkSupport
Imports FrameworkDimension = SolidEdgeFrameworkSupport.Dimension

''' <summary>
''' Postproceso UNE-EN ISO 129-1 sobre cotas ya existentes en hoja: orden, carriles, separación y zonas protegidas.
''' No elimina cotas por estar duplicadas; solo informa. Borrado solo si la cota es claramente corrupta.
''' </summary>
Public NotInheritable Class Une129ArrangeExistingDimensions

    Private Sub New()
    End Sub

    ''' <param name="protectedZones">Zonas prohibidas en metros (hoja); Nothing = cajetín + banda superior (tabla) desde plantilla.</param>
    Public Shared Sub OrdenarCotasExistentesUNE129(
        draftDoc As DraftDocument,
        sheet As Sheet,
        drawingViews As IList(Of DrawingView),
        protectedZones As IList(Of BoundingBox2D),
        config As DimensioningNormConfig,
        log As Action(Of String))

        If sheet Is Nothing OrElse config Is Nothing Then Return
        Dim Lg = Sub(m As String) log?.Invoke(m)
        Dim vCount As Integer = If(drawingViews Is Nothing, 0, drawingViews.Count)
        Lg("[DIM][UNE129][ENTER] OrdenarCotasExistentesUNE129 only_arrange=" & config.OnlyArrangeExistingDimensions.ToString(CultureInfo.InvariantCulture) &
           " keep_duplicates=" & config.KeepIntentionalDuplicateDimensions.ToString(CultureInfo.InvariantCulture) &
           " views_ctx=" & vCount.ToString(CultureInfo.InvariantCulture))

        Dim sheetArea = DimensionPlacementService.TryGetSheetWorkArea(sheet, log)
        Dim zones As List(Of BoundingBox2D) = MergeProtectedZones(sheetArea, protectedZones, config, log)

        Dim dimsObj As Object = Nothing
        Try
            dimsObj = sheet.Dimensions
        Catch ex As Exception
            Lg("[DIM][UNE129][ERR] Dimensions: " & ex.Message)
            Return
        End Try
        Dim dims As Dimensions = TryCast(dimsObj, Dimensions)
        If dims Is Nothing Then Return

        Dim snapshot As New List(Of FrameworkDimension)()
        Try
            For i As Integer = 1 To dims.Count
                Try
                    Dim dx = TryCast(dims.Item(i), FrameworkDimension)
                    If dx IsNot Nothing Then snapshot.Add(dx)
                Catch
                End Try
            Next
        Catch
        End Try

        Dim initialCount As Integer = snapshot.Count
        Dim items As New List(Of DimLayoutItem)()
        Dim corruptDeleted As Integer = 0
        For idx As Integer = 0 To snapshot.Count - 1
            Dim d As FrameworkDimension = snapshot(idx)
            If d Is Nothing Then Continue For

            If IsCorruptOrAbsurdDimension(d, sheetArea, Lg) Then
                Try
                    d.Delete()
                    corruptDeleted += 1
                    Lg("[DIM][UNE129][DELETE][ONLY_IF_CORRUPT] snapshot_idx=" & idx.ToString(CultureInfo.InvariantCulture))
                Catch ex As Exception
                    Lg("[DIM][UNE129][DELETE][FAIL] corrupt snapshot_idx=" & idx.ToString(CultureInfo.InvariantCulture) & " " & ex.Message)
                End Try
                Continue For
            End If

            Dim it As DimLayoutItem = Nothing
            If TryBuildLayoutItem(d, it) Then items.Add(it)
        Next

        Dim duplicatesDetected As Integer = AuditDuplicateValuesInformative(items, Lg)

        Dim moved As Integer = 0
        Dim kept As Integer = 0
        Dim overlapsAvoided As Integer = 0
        Dim zoneNudges As Integer = 0

        If config.AvoidOverlaps Then
            ResolveOverlapsByTrackDistance(items, sheetArea, config, overlapsAvoided)
        End If

        If zones.Count > 0 AndAlso (config.AvoidTitleBlock OrElse config.AvoidBorder) Then
            zoneNudges = NudgeOutOfProtectedZones(items, zones, sheetArea, config, Lg)
        End If

        For Each it In items
            If it.MovedInPass Then
                moved += 1
            Else
                kept += 1
            End If
        Next

        Dim finalCount As Integer = 0
        Try
            finalCount = dims.Count
        Catch
            finalCount = initialCount - corruptDeleted
        End Try

        Lg(String.Format(CultureInfo.InvariantCulture,
            "[DIM][UNE129][SUMMARY] initial={0} final={1} moved={2} kept={3} overlapsAvoided={4} zoneNudges={5} duplicatesDetected={6} duplicatesDeleted={7} corruptDeleted={8}",
            initialCount, finalCount, moved, kept, overlapsAvoided, zoneNudges, duplicatesDetected, 0, corruptDeleted))
        If draftDoc IsNot Nothing Then
            Try
                CallByName(draftDoc, "UpdateAll", CallType.Method, False)
            Catch
            End Try
        End If
        Lg("[DIM][UNE129][EXIT]")
    End Sub

    Private Class DimLayoutItem
        Public D As FrameworkDimension
        Public MinX As Double
        Public MaxX As Double
        Public MinY As Double
        Public MaxY As Double
        Public IsHorizontal As Boolean
        Public NominalMm As Integer?
        Public Track0 As Double
        Public MovedInPass As Boolean

        Public ReadOnly Property MidX As Double
            Get
                Return (MinX + MaxX) * 0.5R
            End Get
        End Property

        Public ReadOnly Property MidY As Double
            Get
                Return (MinY + MaxY) * 0.5R
            End Get
        End Property
    End Class

    Private Shared Function MergeProtectedZones(
        sheetArea As BoundingBox2D,
        userZones As IList(Of BoundingBox2D),
        config As DimensioningNormConfig,
        log As Action(Of String)) As List(Of BoundingBox2D)

        Dim outL As New List(Of BoundingBox2D)()
        If userZones IsNot Nothing Then
            For Each z In userZones
                If z IsNot Nothing Then outL.Add(z)
            Next
        End If
        If sheetArea Is Nothing Then Return outL

        If config.AvoidTitleBlock Then
            Dim tb = DimensionPlacementService.GetTitleBlockAvoidanceBox(sheetArea, config)
            If tb IsNot Nothing Then outL.Add(tb)
        End If

        ' Banda superior conservadora (tabla de piezas / revisiones): ~12 % alto de hoja.
        If config.AvoidBorder Then
            Dim h As Double = sheetArea.Height
            If h > 1.0E-6 Then
                outL.Add(New BoundingBox2D With {
                    .MinX = sheetArea.MinX,
                    .MaxX = sheetArea.MaxX,
                    .MinY = sheetArea.MaxY - h * 0.12R,
                    .MaxY = sheetArea.MaxY
                })
            End If
        End If

        Return outL
    End Function

    Private Shared Function TryBuildLayoutItem(d As FrameworkDimension, ByRef it As DimLayoutItem) As Boolean
        it = Nothing
        If d Is Nothing Then Return False
        Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
        Try
            d.Range(x1, y1, x2, y2)
        Catch
            Return False
        End Try
        Dim minX = Math.Min(x1, x2), maxX = Math.Max(x1, x2)
        Dim minY = Math.Min(y1, y2), maxY = Math.Max(y1, y2)
        Dim sx = maxX - minX, sy = maxY - minY
        Dim isH As Boolean = sx >= sy

        Dim td As Double = 0R
        Try
            td = SafeDouble(CallByName(d, "TrackDistance", CallType.Get))
        Catch
            td = 0R
        End Try

        Dim nmm As Integer? = TryReadNominalMillimeters(d)
        it = New DimLayoutItem With {
            .D = d,
            .MinX = minX, .MaxX = maxX, .MinY = minY, .MaxY = maxY,
            .IsHorizontal = isH,
            .NominalMm = nmm,
            .Track0 = td,
            .MovedInPass = False
        }
        Return True
    End Function

    Private Shared Function TryReadNominalMillimeters(d As FrameworkDimension) As Integer?
        If d Is Nothing Then Return Nothing
        Try
            Dim v As Double = Convert.ToDouble(CallByName(d, "Value", CallType.Get), CultureInfo.InvariantCulture)
            If Double.IsNaN(v) OrElse Double.IsInfinity(v) Then Return Nothing
            Return CInt(Math.Round(v * 1000.0R))
        Catch
            Return Nothing
        End Try
    End Function

    Private Shared Function IsCorruptOrAbsurdDimension(d As FrameworkDimension, sheetArea As BoundingBox2D, log As Action(Of String)) As Boolean
        If d Is Nothing Then Return True
        Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
        Try
            d.Range(x1, y1, x2, y2)
        Catch
            Return True
        End Try
        If Not AreFinite4(x1, y1, x2, y2) Then Return True
        Dim minX = Math.Min(x1, x2), maxX = Math.Max(x1, x2)
        Dim minY = Math.Min(y1, y2), maxY = Math.Max(y1, y2)
        If (maxX - minX) + (maxY - minY) < 1.0E-12 Then
            Dim hasVal As Boolean = TryReadNominalMillimeters(d).HasValue
            If Not hasVal Then Return True
        End If
        If sheetArea Is Nothing Then Return False
        Dim m As Double = Math.Max(sheetArea.Width, sheetArea.Height) * 2.0R + 0.05R
        If maxX < sheetArea.MinX - m OrElse minX > sheetArea.MaxX + m OrElse
           maxY < sheetArea.MinY - m OrElse minY > sheetArea.MaxY + m Then
            Return True
        End If
        Return False
    End Function

    Private Shared Function AreFinite4(a As Double, b As Double, c As Double, d As Double) As Boolean
        Return Not (Double.IsNaN(a) OrElse Double.IsNaN(b) OrElse Double.IsNaN(c) OrElse Double.IsNaN(d))
    End Function

    Private Shared Function AuditDuplicateValuesInformative(items As List(Of DimLayoutItem), log As Action(Of String)) As Integer
        If items Is Nothing OrElse items.Count = 0 Then Return 0
        Dim dupDims As Integer = 0
        Dim groups = items.Where(Function(x) x.NominalMm.HasValue).GroupBy(Function(x) x.NominalMm.Value)
        For Each g In groups
            If g.Count() <= 1 Then Continue For
            Dim v As Integer = g.Key
            dupDims += g.Count()
            log?.Invoke(String.Format(CultureInfo.InvariantCulture,
                "[DIM][UNE129][DUPLICATE][INFO] value={0} keep=True reason=intentional_manual_keypoint_workflow count={1}",
                v, g.Count()))
        Next
        Return dupDims
    End Function

    Private Shared Sub ResolveOverlapsByTrackDistance(
        items As List(Of DimLayoutItem),
        sheetArea As BoundingBox2D,
        config As DimensioningNormConfig,
        ByRef overlapsAvoided As Integer)

        overlapsAvoided = 0
        If items Is Nothing OrElse items.Count < 2 Then Return
        Dim margin As Double = Math.Max(config.MinFeatureSeparation, 0.0004R)
        Dim trackStep As Double = Math.Max(config.GapBetweenDimensionRows * 0.35R, 0.001R)
        If sheetArea IsNot Nothing Then
            trackStep = Math.Max(trackStep, Math.Min(sheetArea.Width, sheetArea.Height) * 0.012R)
        End If

        Dim rounds As Integer = 0
        Dim total As Integer = 0
        While rounds < 4
            rounds += 1
            Dim hit As Boolean = False
            Dim ordered = items.OrderBy(Function(x)
                                            If x.IsHorizontal Then Return x.MidY
                                            Return x.MidX
                                        End Function).ToList()
            For idxA As Integer = 0 To ordered.Count - 2
                For idxB As Integer = idxA + 1 To ordered.Count - 1
                    Dim ia = ordered(idxA)
                    Dim ib = ordered(idxB)
                    If ia.IsHorizontal <> ib.IsHorizontal Then Continue For
                    If Not BboxesOverlap(ia.MinX, ia.MaxX, ia.MinY, ia.MaxY, ib.MinX, ib.MaxX, ib.MinY, ib.MaxY, margin) Then Continue For

                    Dim victim As DimLayoutItem = ib
                    Try
                        Dim td As Double = SafeDouble(CallByName(victim.D, "TrackDistance", CallType.Get))
                        Dim tdNew As Double = td + trackStep
                        CallByName(victim.D, "TrackDistance", CallType.Let, tdNew)
                        victim.MovedInPass = True
                        RefreshItemRange(victim)
                        hit = True
                        total += 1
                    Catch
                    End Try
                Next idxB
            Next idxA
            If Not hit Then Exit While
        End While
        overlapsAvoided = total
    End Sub

    Private Shared Sub RefreshItemRange(it As DimLayoutItem)
        If it Is Nothing OrElse it.D Is Nothing Then Return
        Try
            Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
            it.D.Range(x1, y1, x2, y2)
            it.MinX = Math.Min(x1, x2) : it.MaxX = Math.Max(x1, x2)
            it.MinY = Math.Min(y1, y2) : it.MaxY = Math.Max(y1, y2)
        Catch
        End Try
    End Sub

    Private Shared Function NudgeOutOfProtectedZones(
        items As List(Of DimLayoutItem),
        zones As List(Of BoundingBox2D),
        sheetArea As BoundingBox2D,
        config As DimensioningNormConfig,
        log As Action(Of String)) As Integer

        Dim n As Integer = 0
        Dim nudgeStep As Double = Math.Max(config.GapBetweenDimensionRows * 0.4R, 0.0012R)
        For Each it In items
            If it Is Nothing OrElse it.D Is Nothing Then Continue For
            Dim cx = (it.MinX + it.MaxX) * 0.5R
            Dim cy = (it.MinY + it.MaxY) * 0.5R
            For Each z In zones
                If z Is Nothing Then Continue For
                If Not PointInBox(cx, cy, z.MinX, z.MaxX, z.MinY, z.MaxY) Then Continue For
                Try
                    Dim td As Double = SafeDouble(CallByName(it.D, "TrackDistance", CallType.Get))
                    CallByName(it.D, "TrackDistance", CallType.Let, td + nudgeStep)
                    it.MovedInPass = True
                    RefreshItemRange(it)
                    n += 1
                Catch
                End Try
                Exit For
            Next
        Next
        Return n
    End Function

    Private Shared Function PointInBox(px As Double, py As Double, minX As Double, maxX As Double, minY As Double, maxY As Double) As Boolean
        Return px >= minX AndAlso px <= maxX AndAlso py >= minY AndAlso py <= maxY
    End Function

    Private Shared Function BboxesOverlap(ax1 As Double, ax2 As Double, ay1 As Double, ay2 As Double,
                                          bx1 As Double, bx2 As Double, by1 As Double, by2 As Double,
                                          m As Double) As Boolean
        Return Not (ax2 + m < bx1 OrElse ax1 - m > bx2 OrElse ay2 + m < by1 OrElse ay1 - m > by2)
    End Function

    Private Shared Function SafeDouble(o As Object) As Double
        If o Is Nothing Then Return 0R
        Try
            Return Convert.ToDouble(o, CultureInfo.InvariantCulture)
        Catch
            Return 0R
        End Try
    End Function

End Class
