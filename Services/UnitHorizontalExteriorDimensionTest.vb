Option Strict Off
Option Explicit On

Imports System.Globalization
Imports System.Linq
Imports SolidEdgeDraft
Imports SolidEdgeFrameworkSupport
Imports SolidEdgeConstants
Imports System.Runtime.InteropServices

Public Module UnitHorizontalExteriorDimensionTest

    Private Const EPS As Double = 0.000001
    Private Const Y_TOL As Double = 0.00001
    Private Const USE_GLOBAL_DV_SWEEP As Boolean = True

    Public Sub Run(
        ByVal draftDoc As SolidEdgeDraft.DraftDocument,
        ByVal log As Action(Of String)
    )
        log("[DIM][KPTEST][START] Inicio prueba aislada comparativa View vs Sheet")

        Dim sheet As SolidEdgeDraft.Sheet = draftDoc.ActiveSheet
        Create2DModelSheetDropsWithoutIsometric(draftDoc, sheet, log)
        Dim model2dSheet As SolidEdgeDraft.Sheet = TryGet2DModelSheet(draftDoc, log)
        If model2dSheet IsNot Nothing Then
            log("[DIM][KPTEST][2D_MODEL_SHEET] name=" & SafeSheetName(model2dSheet))
        End If
        Dim dims As SolidEdgeFrameworkSupport.Dimensions = CType(sheet.Dimensions, SolidEdgeFrameworkSupport.Dimensions)

        If USE_GLOBAL_DV_SWEEP Then
            RunGlobalDvSweep(draftDoc, sheet, dims, log)
            log("[DIM][KPTEST][END] Fin barrido global DV")
            Exit Sub
        End If

        Dim view As SolidEdgeDraft.DrawingView = BuscarVistaSuperior(sheet)
        If view Is Nothing Then
            log("[DIM][KPTEST][ERROR] No se encontró vista superior")
            Exit Sub
        End If
        log("[DIM][KPTEST][VIEW_SELECTED] name=" & SafeViewName(view))

        ProbarAddLengthEnHorizontalMasLarga(draftDoc, sheet, dims, view, log)

        Dim lineLeft As Object = Nothing
        Dim lineRight As Object = Nothing
        Dim idxLeft As Integer = 0
        Dim idxRight As Integer = 0
        If Not BuscarVerticalesExteriores(view, lineLeft, lineRight, idxLeft, idxRight) Then
            log("[DIM][KPTEST][ERROR] No se encontraron líneas verticales exteriores")
            Exit Sub
        End If

        DumpDvLine2dMembers(lineLeft, "LEFT", idxLeft, log)
        DumpDvLine2dMembers(lineRight, "RIGHT", idxRight, log)
        DumpDisplayDataForDvLine(lineLeft, "LEFT", idxLeft, log)
        DumpDisplayDataForDvLine(lineRight, "RIGHT", idxRight, log)

        Dim leftEnd As Pt2 = GetLineEnd(lineLeft)
        Dim rightEnd As Pt2 = GetLineEnd(lineRight)
        If Math.Abs(leftEnd.Y - rightEnd.Y) > Y_TOL Then
            log("[DIM][KPTEST][ERROR] End-End no comparte Y. dy=" & Math.Abs(leftEnd.Y - rightEnd.Y).ToString("G17", CultureInfo.InvariantCulture))
            Exit Sub
        End If

        Dim leftEndSheet As Pt2 = ViewToSheetPoint(view, leftEnd)
        Dim rightEndSheet As Pt2 = ViewToSheetPoint(view, rightEnd)

        log("[DIM][KPTEST][LINE_LEFT] idx=" & idxLeft.ToString(CultureInfo.InvariantCulture) & " type=" & SafeTypeName(lineLeft))
        log("[DIM][KPTEST][LINE_RIGHT] idx=" & idxRight.ToString(CultureInfo.InvariantCulture) & " type=" & SafeTypeName(lineRight))
        log("[DIM][KPTEST][LEFT_END_VIEW] x=" & leftEnd.X.ToString("G17", CultureInfo.InvariantCulture) & " y=" & leftEnd.Y.ToString("G17", CultureInfo.InvariantCulture))
        log("[DIM][KPTEST][RIGHT_END_VIEW] x=" & rightEnd.X.ToString("G17", CultureInfo.InvariantCulture) & " y=" & rightEnd.Y.ToString("G17", CultureInfo.InvariantCulture))
        log("[DIM][KPTEST][LEFT_END_SHEET] x=" & leftEndSheet.X.ToString("G17", CultureInfo.InvariantCulture) & " y=" & leftEndSheet.Y.ToString("G17", CultureInfo.InvariantCulture))
        log("[DIM][KPTEST][RIGHT_END_SHEET] x=" & rightEndSheet.X.ToString("G17", CultureInfo.InvariantCulture) & " y=" & rightEndSheet.Y.ToString("G17", CultureInfo.InvariantCulture))

        Dim viewResult As DimensionResult = CrearEInspeccionarDimension(
            dims, draftDoc, sheet, view, lineLeft, lineRight, leftEnd, rightEnd, "CREATE_VIEW", log)

        Dim sheetResult As DimensionResult = CrearEInspeccionarDimension(
            dims, draftDoc, sheet, view, lineLeft, lineRight, leftEndSheet, rightEndSheet, "CREATE_SHEET", log)

        log("[DIM][KPTEST][SUMMARY] " &
            "view_created=" & viewResult.Created.ToString() & " " &
            "sheet_created=" & sheetResult.Created.ToString() & " " &
            "visual_compare=manual_required")

        log("[DIM][KPTEST][END] Fin prueba aislada comparativa")
    End Sub

    Private Sub RunGlobalDvSweep(
        ByVal draftDoc As SolidEdgeDraft.DraftDocument,
        ByVal sheet As SolidEdgeDraft.Sheet,
        ByVal dims As SolidEdgeFrameworkSupport.Dimensions,
        ByVal log As Action(Of String)
    )
        log("[DIM][SWEEP][START] barrido de todas las vistas no isométricas")
        If sheet Is Nothing OrElse dims Is Nothing Then
            log("[DIM][SWEEP][ERROR] Sheet o Dimensions = Nothing")
            Exit Sub
        End If

        Dim methodNames As String() = {
            "AddAngle", "AddAngleBetween3Objects", "AddAngleBetweenObjects",
            "AddAngularCoordinate", "AddAngularCoordinateOrigin", "AddChamfer",
            "AddCircularDiameter", "AddCoordinate", "AddCoordinateEx",
            "AddCoordinateOrigin", "AddCoordinateOriginEx", "AddDimension",
            "AddDistanceBetweenObjects", "AddDistanceBetweenObjectsEX",
            "AddDistanceIntersectionToIntersection", "AddDistanceIntersectionToObject",
            "AddDistanceObjectToIntersection", "AddLength", "AddRadialDiameter",
            "AddRadius", "AddSymmetricalDiameter"
        }

        Dim totalOk As Integer = 0
        Dim totalFail As Integer = 0
        Dim viewIndex As Integer = 0

        For Each dv As SolidEdgeDraft.DrawingView In sheet.DrawingViews
            viewIndex += 1
            If IsIsometricDrawingViewForDrop(dv) Then
                log("[DIM][SWEEP][VIEW][SKIP] idx=" & viewIndex.ToString(CultureInfo.InvariantCulture) & " reason=isometric name=" & SafeViewName(dv))
                Continue For
            End If

            log("[DIM][SWEEP][VIEW][START] idx=" & viewIndex.ToString(CultureInfo.InvariantCulture) & " name=" & SafeViewName(dv))
            Dim elements As List(Of Object) = CollectDvElementsForSweep(dv)
            log("[DIM][SWEEP][VIEW][ELEMENTS] idx=" & viewIndex.ToString(CultureInfo.InvariantCulture) & " count=" & elements.Count.ToString(CultureInfo.InvariantCulture))

            For e As Integer = 0 To elements.Count - 1
                Dim primary As Object = elements(e)
                Dim secondary As Object = If(e + 1 < elements.Count, elements(e + 1), primary)
                Dim p1 As Pt2 = GetRepresentativePoint(primary)
                Dim p2 As Pt2 = GetRepresentativePoint(secondary)
                log("[DIM][SWEEP][ELEMENT][START] viewIdx=" & viewIndex.ToString(CultureInfo.InvariantCulture) &
                    " elemIdx=" & (e + 1).ToString(CultureInfo.InvariantCulture) &
                    " type=" & SafeTypeName(primary))

                For Each methodName In methodNames
                    Dim ok As Boolean = TryInvokeDimensionApiOverloads(draftDoc, sheet, dims, methodName, primary, secondary, p1, p2, log, viewIndex, e + 1)
                    If ok Then
                        totalOk += 1
                    Else
                        totalFail += 1
                    End If
                Next

                log("[DIM][SWEEP][ELEMENT][END] viewIdx=" & viewIndex.ToString(CultureInfo.InvariantCulture) &
                    " elemIdx=" & (e + 1).ToString(CultureInfo.InvariantCulture))
            Next

            log("[DIM][SWEEP][VIEW][END] idx=" & viewIndex.ToString(CultureInfo.InvariantCulture))
        Next

        log("[DIM][SWEEP][SUMMARY] ok=" & totalOk.ToString(CultureInfo.InvariantCulture) &
            " fail=" & totalFail.ToString(CultureInfo.InvariantCulture))
        log("[DIM][SWEEP][END]")
    End Sub

    Private Function CollectDvElementsForSweep(ByVal dv As SolidEdgeDraft.DrawingView) As List(Of Object)
        Dim out As New List(Of Object)()
        AddCollectionItems(out, TryGetCollection(dv, "DVLines2d"))
        AddCollectionItems(out, TryGetCollection(dv, "DVArcs2d"))
        AddCollectionItems(out, TryGetCollection(dv, "DVCircles2d"))
        AddCollectionItems(out, TryGetCollection(dv, "DVEllipses2d"))
        Return out
    End Function

    Private Function TryGetCollection(ByVal owner As Object, ByVal name As String) As Object
        Try
            Return CallByName(owner, name, CallType.Get)
        Catch
            Return Nothing
        End Try
    End Function

    Private Sub AddCollectionItems(ByVal dst As List(Of Object), ByVal collectionObj As Object)
        If collectionObj Is Nothing Then Return
        Try
            For Each it As Object In collectionObj
                If it IsNot Nothing Then dst.Add(it)
            Next
        Catch
        End Try
    End Sub

    Private Function TryInvokeDimensionApiOverloads(
        ByVal draftDoc As SolidEdgeDraft.DraftDocument,
        ByVal sheet As SolidEdgeDraft.Sheet,
        ByVal dims As SolidEdgeFrameworkSupport.Dimensions,
        ByVal methodName As String,
        ByVal primary As Object,
        ByVal secondary As Object,
        ByVal p1 As Pt2,
        ByVal p2 As Pt2,
        ByVal log As Action(Of String),
        ByVal viewIndex As Integer,
        ByVal elemIndex As Integer
    ) As Boolean
        Dim t As Type = dims.GetType()
        Dim methodOverloads = t.GetMethods().Where(Function(m) m.Name = methodName).ToArray()
        If methodOverloads.Length = 0 Then
            log("[DIM][SWEEP][API][MISS] method=" & methodName)
            Return False
        End If

        Dim lastErr As String = ""
        For Each mi In methodOverloads
            Try
                Dim args = BuildArgsForDimensionMethod(mi.GetParameters(), primary, secondary, p1, p2)
                Dim result As Object = mi.Invoke(dims, args)

                Dim createdDim As SolidEdgeFrameworkSupport.Dimension = TryCast(result, SolidEdgeFrameworkSupport.Dimension)
                If createdDim IsNot Nothing Then
                    DrawingViewDimensionCreator.ApplyU35ForIsolatedTest(draftDoc, sheet, createdDim, Sub(_m As String)
                                                                                                     ' Reutiliza aplicación de estilo U3,5.
                                                                                                 End Sub)
                End If

                log("[DIM][SWEEP][API][OK] viewIdx=" & viewIndex.ToString(CultureInfo.InvariantCulture) &
                    " elemIdx=" & elemIndex.ToString(CultureInfo.InvariantCulture) &
                    " method=" & methodName &
                    " sig=" & mi.ToString() &
                    " resultType=" & SafeTypeName(result) &
                    " style=" & If(createdDim Is Nothing, "<n/a>", ReadStyleName(createdDim)))
                Return True
            Catch ex As Exception
                lastErr = ex.Message
            End Try
        Next

        log("[DIM][SWEEP][API][FAIL] viewIdx=" & viewIndex.ToString(CultureInfo.InvariantCulture) &
            " elemIdx=" & elemIndex.ToString(CultureInfo.InvariantCulture) &
            " method=" & methodName &
            " msg=" & lastErr)
        Return False
    End Function

    Private Function BuildArgsForDimensionMethod(
        ByVal ps As Reflection.ParameterInfo(),
        ByVal primary As Object,
        ByVal secondary As Object,
        ByVal p1 As Pt2,
        ByVal p2 As Pt2
    ) As Object()
        Dim args(ps.Length - 1) As Object
        Dim objCount As Integer = 0
        Dim doubleCount As Integer = 0

        For i As Integer = 0 To ps.Length - 1
            Dim rawType As Type = ps(i).ParameterType
            Dim t As Type = If(rawType.IsByRef, rawType.GetElementType(), rawType)

            If t Is GetType(Object) OrElse (Not t.IsValueType AndAlso t IsNot GetType(String)) Then
                objCount += 1
                If objCount = 1 Then
                    args(i) = primary
                Else
                    args(i) = secondary
                End If
            ElseIf t Is GetType(Double) Then
                doubleCount += 1
                Select Case doubleCount
                    Case 1 : args(i) = p1.X
                    Case 2 : args(i) = p1.Y
                    Case 3 : args(i) = 0.0R
                    Case 4 : args(i) = p2.X
                    Case 5 : args(i) = p2.Y
                    Case Else : args(i) = 0.0R
                End Select
            ElseIf t Is GetType(Single) Then
                args(i) = CSng(0.0R)
            ElseIf t Is GetType(Boolean) Then
                args(i) = False
            ElseIf t Is GetType(Integer) OrElse t Is GetType(Long) OrElse t Is GetType(Short) Then
                args(i) = 0
            ElseIf t Is GetType(String) Then
                args(i) = ""
            ElseIf t.IsEnum Then
                args(i) = Activator.CreateInstance(t)
            ElseIf t.IsValueType Then
                args(i) = Activator.CreateInstance(t)
            Else
                args(i) = Nothing
            End If
        Next

        Return args
    End Function

    Private Function GetRepresentativePoint(ByVal obj As Object) As Pt2
        Try
            Dim p As Pt2 = GetLineStart(obj)
            Return p
        Catch
        End Try

        Try
            Dim minX As Double = 0, minY As Double = 0, maxX As Double = 0, maxY As Double = 0
            CallByName(obj, "Range", CallType.Method, minX, minY, maxX, maxY)
            Return New Pt2 With {.X = (minX + maxX) / 2.0R, .Y = (minY + maxY) / 2.0R}
        Catch
        End Try

        Return New Pt2 With {.X = 0.0R, .Y = 0.0R}
    End Function

    Private Sub ProbarAddLengthEnHorizontalMasLarga(
        ByVal draftDoc As SolidEdgeDraft.DraftDocument,
        ByVal sheet As SolidEdgeDraft.Sheet,
        ByVal dims As SolidEdgeFrameworkSupport.Dimensions,
        ByVal drawingView As SolidEdgeDraft.DrawingView,
        ByVal log As Action(Of String)
    )
        log("[DIM][ADDLENGTH][START] prueba con DVLine2d horizontal más larga")

        Dim targetLine As Object = Nothing
        Dim targetIdx As Integer = -1
        Dim targetLength As Double = -1.0R
        Dim i As Integer = 0

        For Each lineObj As Object In drawingView.DVLines2d
            i += 1
            Dim p1 As Pt2 = GetLineStart(lineObj)
            Dim p2 As Pt2 = GetLineEnd(lineObj)
            Dim dx As Double = Math.Abs(p2.X - p1.X)
            Dim dy As Double = Math.Abs(p2.Y - p1.Y)
            If dy < EPS AndAlso dx > EPS Then
                Dim len As Double = Math.Sqrt((p2.X - p1.X) * (p2.X - p1.X) + (p2.Y - p1.Y) * (p2.Y - p1.Y))
                If len > targetLength Then
                    targetLength = len
                    targetLine = lineObj
                    targetIdx = i
                End If
            End If
        Next

        If targetLine Is Nothing Then
            log("[DIM][ADDLENGTH][ERROR] No se encontró línea horizontal válida.")
            log("[DIM][ADDLENGTH][END]")
            Exit Sub
        End If

        log("[DIM][ADDLENGTH][LINE] idx=" & targetIdx.ToString(CultureInfo.InvariantCulture) &
            " length=" & targetLength.ToString("G17", CultureInfo.InvariantCulture) &
            " type=" & SafeTypeName(targetLine))

        Dim refObj As Object = Nothing
        Try
            refObj = CallByName(targetLine, "Reference", CallType.Get)
        Catch ex As Exception
            log("[DIM][ADDLENGTH][ERROR] No se pudo obtener DVLine2d.Reference: " & ex.Message)
            log("[DIM][ADDLENGTH][END]")
            Exit Sub
        End Try

        Dim dimObj As SolidEdgeFrameworkSupport.Dimension = Nothing
        Try
            dimObj = TryCast(CallByName(dims, "AddLength", CallType.Method, refObj), SolidEdgeFrameworkSupport.Dimension)
            If dimObj Is Nothing Then
                log("[DIM][ADDLENGTH][ERROR] AddLength devolvió Nothing.")
                log("[DIM][ADDLENGTH][END]")
                Exit Sub
            End If
            log("[DIM][ADDLENGTH][OK] created=True")
        Catch ex As Exception
            log("[DIM][ADDLENGTH][ERROR] AddLength: " & ex.Message)
            log("[DIM][ADDLENGTH][END]")
            Exit Sub
        End Try

        Try
            CallByName(dimObj, "ProjectionLineDirection", CallType.Let, True)
            log("[DIM][ADDLENGTH][SET] ProjectionLineDirection=True")
        Catch ex As Exception
            log("[DIM][ADDLENGTH][SET][ERROR] ProjectionLineDirection: " & ex.Message)
        End Try

        Try
            CallByName(dimObj, "TrackDistance", CallType.Let, 0.02R)
            log("[DIM][ADDLENGTH][SET] TrackDistance=0.02")
        Catch ex As Exception
            log("[DIM][ADDLENGTH][SET][ERROR] TrackDistance: " & ex.Message)
        End Try

        DrawingViewDimensionCreator.ApplyU35ForIsolatedTest(draftDoc, sheet, dimObj, Sub(_m As String)
                                                                                           ' Reutiliza ruta de estilo U3,5.
                                                                                       End Sub)
        log("[DIM][ADDLENGTH][STYLE] " & ReadStyleName(dimObj))

        RunConnectedSelectSetApis(draftDoc, drawingView, "ADDLENGTH_LONGEST_HORIZONTAL", log)
        log("[DIM][ADDLENGTH][END]")
    End Sub

    Private Function CrearEInspeccionarDimension(
        ByVal dims As SolidEdgeFrameworkSupport.Dimensions,
        ByVal draftDoc As SolidEdgeDraft.DraftDocument,
        ByVal sheet As SolidEdgeDraft.Sheet,
        ByVal drawingView As SolidEdgeDraft.DrawingView,
        ByVal lineLeft As Object,
        ByVal lineRight As Object,
        ByVal p1 As Pt2,
        ByVal p2 As Pt2,
        ByVal tag As String,
        ByVal log As Action(Of String)
    ) As DimensionResult
        Dim r As New DimensionResult()

        log("[DIM][KPTEST][" & tag & "][TRY]")
        If tag = "CREATE_VIEW" Then
            log("[DIM][KPTEST][" & tag & "][TRY] p1_view=(x=" & p1.X.ToString("G17", CultureInfo.InvariantCulture) & ", y=" & p1.Y.ToString("G17", CultureInfo.InvariantCulture) & ", z=0)")
            log("[DIM][KPTEST][" & tag & "][TRY] p2_view=(x=" & p2.X.ToString("G17", CultureInfo.InvariantCulture) & ", y=" & p2.Y.ToString("G17", CultureInfo.InvariantCulture) & ", z=0)")
        Else
            log("[DIM][KPTEST][" & tag & "][TRY] p1_sheet=(x=" & p1.X.ToString("G17", CultureInfo.InvariantCulture) & ", y=" & p1.Y.ToString("G17", CultureInfo.InvariantCulture) & ", z=0)")
            log("[DIM][KPTEST][" & tag & "][TRY] p2_sheet=(x=" & p2.X.ToString("G17", CultureInfo.InvariantCulture) & ", y=" & p2.Y.ToString("G17", CultureInfo.InvariantCulture) & ", z=0)")
        End If
        log("[DIM][KPTEST][" & tag & "][TRY] keyPoint1=True")
        log("[DIM][KPTEST][" & tag & "][TRY] keyPoint2=True")

        Dim dimObj As SolidEdgeFrameworkSupport.Dimension = Nothing
        Try

            dimObj = dims.AddDistanceBetweenObjects(
                lineLeft, p1.X, p1.Y, 0.0, False,
                lineRight, p2.X, p2.Y, 0.0, False
            )
            r.Created = (dimObj IsNot Nothing)
            log("[DIM][KPTEST][" & tag & "][OK] created=" & r.Created.ToString())
        Catch ex As Exception
            log("[DIM][KPTEST][" & tag & "][FAIL] " & ex.Message)
            Return r
        End Try

        If dimObj Is Nothing Then Return r

        DrawingViewDimensionCreator.ApplyU35ForIsolatedTest(draftDoc, sheet, dimObj, Sub(_m As String)
                                                                                          ' Se reutiliza la implementación existente; el log de la comparación se normaliza abajo.
                                                                                      End Sub)
        log("[DIM][KPTEST][" & tag & "][STYLE] " & ReadStyleName(dimObj))

        '*****VAS A PONEWR AQUI LAS DOS FUNCIONES SIGUIENTES: AddConnectedAnnotationsToSelectSet  Y AddConnectedDimensionsToSelectSet Method
        RunConnectedSelectSetApis(draftDoc, drawingView, tag, log)

        Try
            Dim status = dimObj.UpdateStatus()
            log("[DIM][KPTEST][" & tag & "][STATUS] " & status.ToString())
        Catch ex As Exception
            log("[DIM][KPTEST][" & tag & "][STATUS] error=" & ex.Message)
        End Try

        Return r
    End Function

    Private Sub RunConnectedSelectSetApis(
        ByVal draftDoc As SolidEdgeDraft.DraftDocument,
        ByVal drawingView As SolidEdgeDraft.DrawingView,
        ByVal tag As String,
        ByVal log As Action(Of String)
    )
        If draftDoc Is Nothing OrElse drawingView Is Nothing Then
            log("[DIM][KPTEST][" & tag & "][SELECTSET][ERROR] draftDoc o drawingView = Nothing")
            Exit Sub
        End If

        Dim selectSetObj As Object = Nothing
        Try
            selectSetObj = draftDoc.SelectSet
        Catch ex As Exception
            log("[DIM][KPTEST][" & tag & "][SELECTSET][ERROR] No se pudo leer SelectSet: " & ex.Message)
            Exit Sub
        End Try

        Dim countBefore As Integer = SafeSelectSetCount(selectSetObj)
        log("[DIM][KPTEST][" & tag & "][SELECTSET][BEFORE] count=" & countBefore.ToString(CultureInfo.InvariantCulture))

        Try
            drawingView.AddConnectedDimensionsToSelectSet()
            Dim countAfterDims As Integer = SafeSelectSetCount(selectSetObj)
            log("[DIM][KPTEST][" & tag & "][SELECTSET][AddConnectedDimensionsToSelectSet] before=" &
                countBefore.ToString(CultureInfo.InvariantCulture) &
                " after=" & countAfterDims.ToString(CultureInfo.InvariantCulture) &
                " delta=" & (countAfterDims - countBefore).ToString(CultureInfo.InvariantCulture))
            countBefore = countAfterDims
        Catch ex As Exception
            log("[DIM][KPTEST][" & tag & "][SELECTSET][AddConnectedDimensionsToSelectSet][ERROR] " & ex.Message)
        End Try

        Try
            drawingView.AddConnectedAnnotationsToSelectSet()
            Dim countAfterAnn As Integer = SafeSelectSetCount(selectSetObj)
            log("[DIM][KPTEST][" & tag & "][SELECTSET][AddConnectedAnnotationsToSelectSet] before=" &
                countBefore.ToString(CultureInfo.InvariantCulture) &
                " after=" & countAfterAnn.ToString(CultureInfo.InvariantCulture) &
                " delta=" & (countAfterAnn - countBefore).ToString(CultureInfo.InvariantCulture))
        Catch ex As Exception
            log("[DIM][KPTEST][" & tag & "][SELECTSET][AddConnectedAnnotationsToSelectSet][ERROR] " & ex.Message)
        End Try
    End Sub

    Private Function SafeSelectSetCount(ByVal selectSetObj As Object) As Integer
        If selectSetObj Is Nothing Then Return -1
        Try
            Return Convert.ToInt32(CallByName(selectSetObj, "Count", CallType.Get), CultureInfo.InvariantCulture)
        Catch
            Return -1
        End Try
    End Function

    Private Function BuscarVistaSuperior(ByVal sheet As SolidEdgeDraft.Sheet) As SolidEdgeDraft.DrawingView
        Dim bestView As SolidEdgeDraft.DrawingView = Nothing
        Dim bestScore As Double = Double.MinValue
        For Each v As SolidEdgeDraft.DrawingView In sheet.DrawingViews
            Dim count As Integer = 0
            Try
                count = v.DVLines2d.Count
            Catch
                count = 0
            End Try
            If count <= 0 Then Continue For
            Dim bbox As BBox = CalcularBBoxDVLines(v)
            Dim score As Double = (bbox.MaxX - bbox.MinX) - (bbox.MaxY - bbox.MinY)
            If score > bestScore Then
                bestScore = score
                bestView = v
            End If
        Next
        Return bestView
    End Function

    Private Function BuscarVerticalesExteriores(
        ByVal view As SolidEdgeDraft.DrawingView,
        ByRef leftLine As Object,
        ByRef rightLine As Object,
        ByRef outLeftIndex As Integer,
        ByRef outRightIndex As Integer
    ) As Boolean

        Dim candidates As New List(Of LineInfo)
        Dim i As Integer = 0

        For Each lineObj As Object In view.DVLines2d
            i += 1
            Dim p1 As Pt2 = GetLineStart(lineObj)
            Dim p2 As Pt2 = GetLineEnd(lineObj)
            Dim dx As Double = Math.Abs(p2.X - p1.X)
            Dim dy As Double = Math.Abs(p2.Y - p1.Y)
            If dx < EPS AndAlso dy > EPS Then
                candidates.Add(New LineInfo With {
                    .Index = i,
                    .LineObject = lineObj,
                    .XMean = (p1.X + p2.X) / 2.0
                })
            End If
        Next

        If candidates.Count < 2 Then Return False

        Dim left = candidates.OrderBy(Function(c) c.XMean).First()
        Dim right = candidates.OrderByDescending(Function(c) c.XMean).First()
        leftLine = left.LineObject
        rightLine = right.LineObject
        outLeftIndex = left.Index
        outRightIndex = right.Index
        Return True
    End Function

    Private Function ViewToSheetPoint(
        ByVal view As SolidEdgeDraft.DrawingView,
        ByVal p As Pt2
    ) As Pt2
        Dim xSheet As Double
        Dim ySheet As Double
        view.ViewToSheet(p.X, p.Y, xSheet, ySheet)
        Return New Pt2 With {.X = xSheet, .Y = ySheet}
    End Function

    Private Function ReadStyleName(ByVal dimObj As SolidEdgeFrameworkSupport.Dimension) As String
        Try
            Dim st As Object = CallByName(dimObj, "Style", CallType.Get)
            If st Is Nothing Then Return "name=<none>"
            Dim nm As String = Convert.ToString(CallByName(st, "Name", CallType.Get), CultureInfo.InvariantCulture)
            Return "name=" & nm
        Catch ex As Exception
            Return "error=" & ex.Message
        End Try
    End Function

    Private Sub Create2DModelSheetDropsWithoutIsometric(
        ByVal draftDoc As SolidEdgeDraft.DraftDocument,
        ByVal sourceSheet As SolidEdgeDraft.Sheet,
        ByVal log As Action(Of String)
    )
        log("[DIM][DROP][START] Crear drops en 2D Model Sheet (sin isométrica)")

        Dim modelSheet As SolidEdgeDraft.Sheet = TryGet2DModelSheet(draftDoc, log)
        If modelSheet Is Nothing Then
            log("[DIM][DROP][ERROR] No disponible Get2DModelSheet; se omite creación de drops.")
            log("[DIM][DROP][END]")
            Exit Sub
        End If

        Dim dstViews As SolidEdgeDraft.DrawingViews = Nothing
        Try
            dstViews = modelSheet.DrawingViews
        Catch ex As Exception
            log("[DIM][DROP][ERROR] No se pudo leer DrawingViews destino: " & ex.Message)
            log("[DIM][DROP][END]")
            Exit Sub
        End Try

        Dim createdCount As Integer = 0
        Dim skippedIsoCount As Integer = 0
        Dim errorCount As Integer = 0
        Dim placeIndex As Integer = 0

        For Each src As SolidEdgeDraft.DrawingView In sourceSheet.DrawingViews
            Dim srcName As String = SafeViewName(src)
            If IsIsometricDrawingViewForDrop(src) Then
                skippedIsoCount += 1
                log("[DIM][DROP][SKIP] view=" & srcName & " reason=isometric")
                Continue For
            End If

            Dim link As Object = Nothing
            Try
                link = src.ModelLink
            Catch ex As Exception
                errorCount += 1
                log("[DIM][DROP][ERROR] view=" & srcName & " ModelLink: " & ex.Message)
                Continue For
            End Try
            If link Is Nothing Then
                errorCount += 1
                log("[DIM][DROP][ERROR] view=" & srcName & " ModelLink=Nothing")
                Continue For
            End If

            Dim ori As SolidEdgeConstants.ViewOrientationConstants = SolidEdgeConstants.ViewOrientationConstants.igFrontView
            Try
                Dim vx As Double = 0, vy As Double = 0, vz As Double = 0, lx As Double = 0, ly As Double = 0, lz As Double = 0
                src.ViewOrientation(vx, vy, vz, lx, ly, lz, ori)
            Catch
            End Try

            Dim scale As Double = 1.0R
            Try
                scale = Convert.ToDouble(CallByName(src, "ScaleFactor", CallType.Get), CultureInfo.InvariantCulture)
            Catch
                Try
                    scale = Convert.ToDouble(CallByName(src, "Scale", CallType.Get), CultureInfo.InvariantCulture)
                Catch
                    scale = 1.0R
                End Try
            End Try

            Dim row As Integer = placeIndex \ 3
            Dim col As Integer = placeIndex Mod 3
            Dim x As Double = 0.12R + (0.16R * col)
            Dim y As Double = 0.12R + (0.16R * row)
            placeIndex += 1

            Dim newView As Object = Nothing
            Try
                newView = dstViews.AddPartView(
                    link,
                    CInt(ori),
                    scale,
                    x,
                    y,
                    SolidEdgeDraft.PartDrawingViewTypeConstants.sePartDesignedView
                )
            Catch exPart As Exception
                Try
                    newView = dstViews.AddSheetMetalView(
                        link,
                        CInt(ori),
                        scale,
                        x,
                        y,
                        SolidEdgeDraft.SheetMetalDrawingViewTypeConstants.seSheetMetalDesignedView
                    )
                Catch exSm As Exception
                    errorCount += 1
                    log("[DIM][DROP][ERROR] view=" & srcName & " AddPartView=" & exPart.Message & " AddSheetMetalView=" & exSm.Message)
                    Continue For
                End Try
            End Try

            If newView Is Nothing Then
                errorCount += 1
                log("[DIM][DROP][ERROR] view=" & srcName & " create_result=Nothing")
                Continue For
            End If

            createdCount += 1
            log("[DIM][DROP][CREATED] src_view=" & srcName &
                " dst_sheet=" & SafeSheetName(modelSheet) &
                " ori=" & ori.ToString() &
                " scale=" & scale.ToString("G17", CultureInfo.InvariantCulture) &
                " x=" & x.ToString("G17", CultureInfo.InvariantCulture) &
                " y=" & y.ToString("G17", CultureInfo.InvariantCulture))
        Next

        log("[DIM][DROP][SUMMARY] created=" & createdCount.ToString(CultureInfo.InvariantCulture) &
            " skipped_isometric=" & skippedIsoCount.ToString(CultureInfo.InvariantCulture) &
            " errors=" & errorCount.ToString(CultureInfo.InvariantCulture))
        log("[DIM][DROP][END]")
    End Sub

    Private Function TryGet2DModelSheet(
        ByVal draftDoc As SolidEdgeDraft.DraftDocument,
        ByVal log As Action(Of String)
    ) As SolidEdgeDraft.Sheet
        Dim objSheets As SolidEdgeDraft.Sheets = Nothing
        Dim objSections As SolidEdgeDraft.Sections = Nothing
        Dim fromSheets As SolidEdgeDraft.Sheet = Nothing
        Dim fromSections As SolidEdgeDraft.Sheet = Nothing

        Try
            objSheets = draftDoc.Sheets
            fromSheets = objSheets.Get2DModelSheet
            log("[DIM][KPTEST][2D_MODEL_SHEET][SHEETS][OK] name=" & SafeSheetName(fromSheets))
        Catch ex As Exception
            log("[DIM][KPTEST][2D_MODEL_SHEET][SHEETS][ERROR] " & ex.Message)
        End Try

        Try
            objSections = draftDoc.Sections
            fromSections = objSections.Get2DModelSheet
            log("[DIM][KPTEST][2D_MODEL_SHEET][SECTIONS][OK] name=" & SafeSheetName(fromSections))
        Catch ex As Exception
            log("[DIM][KPTEST][2D_MODEL_SHEET][SECTIONS][ERROR] " & ex.Message)
        End Try

        If fromSheets IsNot Nothing Then Return fromSheets
        If fromSections IsNot Nothing Then Return fromSections
        Return Nothing
    End Function

    Private Sub DumpDvLine2dMembers(
        ByVal lineObj As Object,
        ByVal side As String,
        ByVal idx As Integer,
        ByVal log As Action(Of String)
    )
        log($"[DIM][DVLINE2D][START] side={side} idx={idx}")
        Try
            log($"[DIM][DVLINE2D][TYPE] side={side} idx={idx} type={SafeTypeName(lineObj)}")

            DumpDvLine2dProperty(lineObj, side, idx, "Angle", log)
            DumpDvLine2dProperty(lineObj, side, idx, "Application", log)
            DumpDvLine2dProperty(lineObj, side, idx, "AttributeSets", log)
            DumpDvLine2dProperty(lineObj, side, idx, "BendAngle", log)
            DumpDvLine2dProperty(lineObj, side, idx, "BendQuantity", log)
            DumpDvLine2dProperty(lineObj, side, idx, "BendRadius", log)
            DumpDvLine2dProperty(lineObj, side, idx, "Document", log)
            DumpDvLine2dProperty(lineObj, side, idx, "DrawingView", log)
            DumpDvLine2dProperty(lineObj, side, idx, "EdgeType", log)
            DumpDvLine2dProperty(lineObj, side, idx, "IsAttributeSetPresent", log)
            DumpDvLine2dProperty(lineObj, side, idx, "Key", log)
            DumpDvLine2dProperty(lineObj, side, idx, "KeyPointCount", log)
            DumpDvLine2dProperty(lineObj, side, idx, "Length", log)
            DumpDvLine2dProperty(lineObj, side, idx, "ModelMember", log)
            DumpDvLine2dProperty(lineObj, side, idx, "ModelWelds", log)
            DumpDvLine2dProperty(lineObj, side, idx, "Name", log)
            DumpDvLine2dProperty(lineObj, side, idx, "Reference", log)
            DumpDvLine2dProperty(lineObj, side, idx, "Relationships", log)
            DumpDvLine2dProperty(lineObj, side, idx, "SegmentedStyleCount", log)
            DumpDvLine2dProperty(lineObj, side, idx, "ShowHideEdgeOverride", log)
            DumpDvLine2dProperty(lineObj, side, idx, "Type", log)

            Try
                Dim x As Double = 0, y As Double = 0
                lineObj.GetStartPoint(x, y)
                log($"[DIM][DVLINE2D][METHOD][GetStartPoint] side={side} idx={idx} x={x.ToString("G17", CultureInfo.InvariantCulture)} y={y.ToString("G17", CultureInfo.InvariantCulture)}")
            Catch ex As Exception
                log($"[DIM][DVLINE2D][METHOD][GetStartPoint][ERROR] side={side} idx={idx} msg={ex.Message}")
            End Try

            Try
                Dim x As Double = 0, y As Double = 0
                lineObj.GetEndPoint(x, y)
                log($"[DIM][DVLINE2D][METHOD][GetEndPoint] side={side} idx={idx} x={x.ToString("G17", CultureInfo.InvariantCulture)} y={y.ToString("G17", CultureInfo.InvariantCulture)}")
            Catch ex As Exception
                log($"[DIM][DVLINE2D][METHOD][GetEndPoint][ERROR] side={side} idx={idx} msg={ex.Message}")
            End Try

            Try
                Dim lowX As Double = 0, lowY As Double = 0, highX As Double = 0, highY As Double = 0
                lineObj.Range(lowX, lowY, highX, highY)
                log($"[DIM][DVLINE2D][METHOD][Range] side={side} idx={idx} lowX={lowX.ToString("G17", CultureInfo.InvariantCulture)} lowY={lowY.ToString("G17", CultureInfo.InvariantCulture)} highX={highX.ToString("G17", CultureInfo.InvariantCulture)} highY={highY.ToString("G17", CultureInfo.InvariantCulture)}")
            Catch ex As Exception
                log($"[DIM][DVLINE2D][METHOD][Range][ERROR] side={side} idx={idx} msg={ex.Message}")
            End Try

            Try
                Dim keyPointCount As Integer = Convert.ToInt32(CallByName(lineObj, "KeyPointCount", CallType.Get), CultureInfo.InvariantCulture)
                For i As Integer = 0 To Math.Max(0, keyPointCount - 1)
                    Try
                        Dim x As Double = 0, y As Double = 0, z As Double = 0
                        Dim kpType As Integer = 0
                        Dim handleType As Integer = 0
                        lineObj.GetKeyPoint(i, x, y, z, kpType, handleType)
                        log($"[DIM][DVLINE2D][METHOD][GetKeyPoint] side={side} idx={idx} kpIndex={i} x={x.ToString("G17", CultureInfo.InvariantCulture)} y={y.ToString("G17", CultureInfo.InvariantCulture)} z={z.ToString("G17", CultureInfo.InvariantCulture)} keyPointType={kpType.ToString(CultureInfo.InvariantCulture)} handleType={handleType.ToString(CultureInfo.InvariantCulture)}")
                    Catch exKp As Exception
                        log($"[DIM][DVLINE2D][METHOD][GetKeyPoint][ERROR] side={side} idx={idx} kpIndex={i} msg={exKp.Message}")
                    End Try
                Next
            Catch ex As Exception
                log($"[DIM][DVLINE2D][METHOD][GetKeyPoint][ERROR] side={side} idx={idx} msg={ex.Message}")
            End Try

            Try
                Dim refKey As Object = CallByName(lineObj, "GetReferenceKey", CallType.Method)
                log($"[DIM][DVLINE2D][METHOD][GetReferenceKey] side={side} idx={idx} raw={FormatRawValues(NormalizeResultArray(refKey))}")
            Catch ex As Exception
                log($"[DIM][DVLINE2D][METHOD][GetReferenceKey][ERROR] side={side} idx={idx} msg={ex.Message}")
            End Try

            Try
                Dim segCount As Integer = Convert.ToInt32(CallByName(lineObj, "SegmentedStyleCount", CallType.Get), CultureInfo.InvariantCulture)
                For i As Integer = 1 To segCount
                    Try
                        Dim res As Object = CallByName(lineObj, "GetSegmentedStyle", CallType.Method, i)
                        log($"[DIM][DVLINE2D][METHOD][GetSegmentedStyle] side={side} idx={idx} segIndex={i} raw={FormatRawValues(NormalizeResultArray(res))}")
                    Catch exSeg As Exception
                        log($"[DIM][DVLINE2D][METHOD][GetSegmentedStyle][ERROR] side={side} idx={idx} segIndex={i} msg={exSeg.Message}")
                    End Try
                Next
            Catch ex As Exception
                log($"[DIM][DVLINE2D][METHOD][GetSegmentedStyle][ERROR] side={side} idx={idx} msg={ex.Message}")
            End Try

        Catch ex As Exception
            log($"[DIM][DVLINE2D][FATAL] side={side} idx={idx} msg={ex.Message}")
        Finally
            log($"[DIM][DVLINE2D][END] side={side} idx={idx}")
        End Try
    End Sub

    Private Sub DumpDvLine2dProperty(
        ByVal lineObj As Object,
        ByVal side As String,
        ByVal idx As Integer,
        ByVal propertyName As String,
        ByVal log As Action(Of String)
    )
        If String.Equals(propertyName, "IsAttributeSetPresent", StringComparison.OrdinalIgnoreCase) Then
            DumpIsAttributeSetPresent(lineObj, side, idx, log)
            Exit Sub
        End If

        Try
            Dim value As Object = CallByName(lineObj, propertyName, CallType.Get)
            log($"[DIM][DVLINE2D][PROPERTY] side={side} idx={idx} name={propertyName} value={FormatOne(value)}")
        Catch ex As Exception
            log($"[DIM][DVLINE2D][PROPERTY][ERROR] side={side} idx={idx} name={propertyName} msg={ex.Message}")
        End Try
    End Sub

    Private Sub DumpIsAttributeSetPresent(
        ByVal lineObj As Object,
        ByVal side As String,
        ByVal idx As Integer,
        ByVal log As Action(Of String)
    )
        Dim candidateNames As String() = {"DIM_FORENSIC", "DEFAULT", ""}

        For Each setName In candidateNames
            Try
                Dim value As Object = CallByName(lineObj, "IsAttributeSetPresent", CallType.Get, setName)
                log($"[DIM][DVLINE2D][PROPERTY] side={side} idx={idx} name=IsAttributeSetPresent attrSet=""{setName}"" value={FormatOne(value)}")
                Return
            Catch ex As Exception
                log($"[DIM][DVLINE2D][PROPERTY][TRY_ERROR] side={side} idx={idx} name=IsAttributeSetPresent attrSet=""{setName}"" msg={ex.Message}")
            End Try
        Next

        log($"[DIM][DVLINE2D][PROPERTY][ERROR] side={side} idx={idx} name=IsAttributeSetPresent msg=No se pudo evaluar con ningún nombre de AttributeSet.")
    End Sub

    Private Sub DumpDisplayDataForDvLine(
        ByVal lineObj As Object,
        ByVal side As String,
        ByVal idx As Integer,
        ByVal log As Action(Of String)
    )
        log($"[DIM][DISPLAYDATA][START] side={side} idx={idx}")

        Try
            Dim dd As Object = Nothing
            Try
                dd = CallByName(lineObj, "DisplayData", CallType.Get)
            Catch ex As Exception
                log($"[DIM][DISPLAYDATA][ERROR] side={side} idx={idx} no_DisplayData msg={ex.Message}")
                Exit Sub
            End Try

            If dd Is Nothing Then
                log($"[DIM][DISPLAYDATA][ERROR] side={side} idx={idx} DisplayData=Nothing")
                Exit Sub
            End If

            log($"[DIM][DISPLAYDATA][TYPE] side={side} idx={idx} type={dd.GetType().FullName}")

            Dim lineCount As Integer = GetDisplayDataCount(dd, "GetLineCount", "LINE_COUNT", side, idx, log)
            For i As Integer = 1 To lineCount
                LogDisplayDataAtIndex(dd, "GetLineAtIndex", "LINE", side, idx, i, log)
            Next

            Dim arcCount As Integer = GetDisplayDataCount(dd, "GetArcCount", "ARC_COUNT", side, idx, log)
            For i As Integer = 1 To arcCount
                LogDisplayDataAtIndex(dd, "GetArcAtIndex", "ARC", side, idx, i, log)
            Next

            Dim ellipseCount As Integer = GetDisplayDataCount(dd, "GetEllipseCount", "ELLIPSE_COUNT", side, idx, log)
            For i As Integer = 1 To ellipseCount
                LogDisplayDataAtIndex(dd, "GetEllipseAtIndex", "ELLIPSE", side, idx, i, log)
            Next

            Dim lsCount As Integer = GetDisplayDataCount(dd, "GetLinestringCount", "LINESTRING_COUNT", side, idx, log)
            For i As Integer = 1 To lsCount
                Dim lsSize As Integer = GetDisplayDataCount(dd, "GetLinestringSizeAtIndex", "LINESTRING_SIZE", side, idx, log, i)
                If lsSize >= 0 Then
                    log($"[DIM][DISPLAYDATA][LINESTRING_SIZE] side={side} idx={idx} lsIndex={i} size={lsSize}")
                End If
                LogDisplayDataAtIndex(dd, "GetLinestringAtIndex", "LINESTRING_POINT", side, idx, i, log)
            Next

            Dim termCount As Integer = GetDisplayDataCount(dd, "GetTerminatorCount", "TERMINATOR_COUNT", side, idx, log)
            For i As Integer = 1 To termCount
                LogDisplayDataAtIndex(dd, "GetTerminatorAtIndex", "TERMINATOR", side, idx, i, log)
            Next

            Dim textCount As Integer = GetDisplayDataCount(dd, "GetTextCount", "TEXT_COUNT", side, idx, log)
            For i As Integer = 1 To textCount
                LogDisplayDataAtIndex(dd, "GetTextAtIndex", "TEXT", side, idx, i, log)
            Next

            Dim textCountEx As Integer = GetDisplayDataCount(dd, "GetTextCountEx", "TEXT_COUNT_EX", side, idx, log)
            For i As Integer = 1 To textCountEx
                LogDisplayDataAtIndex(dd, "GetTextAtIndexEx", "TEXT_EX", side, idx, i, log)
                LogDisplayDataAtIndex(dd, "GetTextAndFontAtIndex", "TEXT_AND_FONT", side, idx, i, log)
                LogDisplayDataAtIndex(dd, "GetTextAndFontAtIndexEx", "TEXT_AND_FONT_EX", side, idx, i, log)
            Next

        Catch ex As Exception
            log($"[DIM][DISPLAYDATA][FATAL] side={side} idx={idx} msg={ex.Message}")
        Finally
            log($"[DIM][DISPLAYDATA][END] side={side} idx={idx}")
        End Try
    End Sub

    Private Function GetDisplayDataCount(
        ByVal dd As Object,
        ByVal methodName As String,
        ByVal label As String,
        ByVal side As String,
        ByVal idx As Integer,
        ByVal log As Action(Of String),
        Optional ByVal atIndex As Integer? = Nothing
    ) As Integer
        Dim outValues As Object() = Nothing
        Dim signature As String = ""
        Dim returnValue As Object = Nothing
        Dim err As String = ""

        If TryInvokeDisplayDataMethod(dd, methodName, atIndex, outValues, signature, returnValue, err) Then
            Dim count As Integer = ExtractFirstInteger(outValues, returnValue)
            log($"[DIM][DISPLAYDATA][{label}] side={side} idx={idx} count={count} sig={signature}")
            Return count
        End If

        log($"[DIM][DISPLAYDATA][{label}][ERROR] side={side} idx={idx} msg={err}")
        Return -1
    End Function

    Private Sub LogDisplayDataAtIndex(
        ByVal dd As Object,
        ByVal methodName As String,
        ByVal label As String,
        ByVal side As String,
        ByVal idx As Integer,
        ByVal itemIndex As Integer,
        ByVal log As Action(Of String)
    )
        Dim outValues As Object() = Nothing
        Dim signature As String = ""
        Dim returnValue As Object = Nothing
        Dim err As String = ""

        If TryInvokeDisplayDataMethod(dd, methodName, itemIndex, outValues, signature, returnValue, err) Then
            Dim merged As Object() = MergeOutputs(outValues, returnValue)
            log($"[DIM][DISPLAYDATA][{label}] side={side} idx={idx} itemIndex={itemIndex} sig={signature} raw={FormatRawValues(merged)}")
        Else
            log($"[DIM][DISPLAYDATA][{label}][ERROR] side={side} idx={idx} itemIndex={itemIndex} msg={err}")
        End If
    End Sub

    Private Function TryInvokeDisplayDataMethod(
        ByVal dd As Object,
        ByVal methodName As String,
        ByVal atIndex As Integer?,
        ByRef outValues As Object(),
        ByRef signature As String,
        ByRef returnValue As Object,
        ByRef errorMessage As String
    ) As Boolean
        outValues = Array.Empty(Of Object)()
        signature = ""
        returnValue = Nothing
        errorMessage = "No overload found."

        Try
            Dim methods = dd.GetType().GetMethods().Where(Function(m) m.Name = methodName).ToArray()
            If methods Is Nothing OrElse methods.Length = 0 Then
                errorMessage = "Método no encontrado."
                Return False
            End If

            For Each mi In methods
                Try
                    Dim ps = mi.GetParameters()
                    Dim args(ps.Length - 1) As Object
                    For i As Integer = 0 To ps.Length - 1
                        Dim pType As Type = ps(i).ParameterType
                        Dim baseType As Type = If(pType.IsByRef, pType.GetElementType(), pType)
                        If i = 0 AndAlso atIndex.HasValue Then
                            args(i) = Convert.ChangeType(atIndex.Value, baseType, CultureInfo.InvariantCulture)
                        Else
                            args(i) = CreateDefaultForType(baseType)
                        End If
                    Next

                    returnValue = mi.Invoke(dd, args)
                    signature = mi.ToString()
                    If atIndex.HasValue Then
                        outValues = args.Skip(1).ToArray()
                    Else
                        outValues = args
                    End If
                    Return True
                Catch exOne As Exception
                    errorMessage = exOne.Message
                End Try
            Next
        Catch ex As Exception
            errorMessage = ex.Message
        End Try

        Return False
    End Function

    Private Function CreateDefaultForType(ByVal t As Type) As Object
        If t Is GetType(String) Then Return ""
        If t Is GetType(Boolean) Then Return False
        If t.IsValueType Then Return Activator.CreateInstance(t)
        Return Nothing
    End Function

    Private Function ExtractFirstInteger(ByVal outValues As Object(), ByVal returnValue As Object) As Integer
        If outValues IsNot Nothing Then
            For Each v In outValues
                If IsNumeric(v) Then
                    Return Convert.ToInt32(v, CultureInfo.InvariantCulture)
                End If
            Next
        End If
        If IsNumeric(returnValue) Then
            Return Convert.ToInt32(returnValue, CultureInfo.InvariantCulture)
        End If
        Return -1
    End Function

    Private Function MergeOutputs(ByVal outValues As Object(), ByVal returnValue As Object) As Object()
        If returnValue Is Nothing Then
            Return If(outValues, Array.Empty(Of Object)())
        End If
        Dim baseVals As Object() = If(outValues, Array.Empty(Of Object)())
        Dim result(baseVals.Length) As Object
        For i As Integer = 0 To baseVals.Length - 1
            result(i) = baseVals(i)
        Next
        result(baseVals.Length) = returnValue
        Return result
    End Function

    Private Function InvokeCountMethod(ByVal target As Object, ByVal methodName As String, ParamArray args() As Object) As Integer
        Dim result As Object
        Select Case args.Length
            Case 0
                result = CallByName(target, methodName, CallType.Method)
            Case 1
                result = CallByName(target, methodName, CallType.Method, args(0))
            Case 2
                result = CallByName(target, methodName, CallType.Method, args(0), args(1))
            Case Else
                Throw New InvalidOperationException("InvokeCountMethod solo soporta hasta 2 argumentos.")
        End Select
        If result Is Nothing Then Return 0
        Return Convert.ToInt32(result, CultureInfo.InvariantCulture)
    End Function

    Private Function NormalizeResultArray(ByVal result As Object) As Object()
        If result Is Nothing Then Return Array.Empty(Of Object)()
        If TypeOf result Is Array Then
            Dim arr As Array = DirectCast(result, Array)
            Dim vals(arr.Length - 1) As Object
            For i As Integer = 0 To arr.Length - 1
                vals(i) = arr.GetValue(i)
            Next
            Return vals
        End If
        Return New Object() {result}
    End Function

    Private Function FormatValueTuple(ByVal vals As Object(), ParamArray names() As String) As String
        Dim parts As New List(Of String)
        For i As Integer = 0 To names.Length - 1
            Dim v As Object = If(i < vals.Length, vals(i), Nothing)
            parts.Add(names(i) & "=" & FormatOne(v))
        Next
        Return String.Join(" ", parts)
    End Function

    Private Function FormatRawValues(ByVal vals As Object()) As String
        If vals Is Nothing OrElse vals.Length = 0 Then Return "<empty>"
        Dim parts As New List(Of String)
        For i As Integer = 0 To vals.Length - 1
            parts.Add($"v{i + 1}={FormatOne(vals(i))}")
        Next
        Return String.Join(" ", parts)
    End Function

    Private Function FormatOne(ByVal value As Object) As String
        If value Is Nothing Then Return "Nothing"
        If TypeOf value Is Double Then Return DirectCast(value, Double).ToString("G17", CultureInfo.InvariantCulture)
        If TypeOf value Is Single Then Return DirectCast(value, Single).ToString("G9", CultureInfo.InvariantCulture)
        If TypeOf value Is IFormattable Then Return DirectCast(value, IFormattable).ToString(Nothing, CultureInfo.InvariantCulture)
        Return Convert.ToString(value, CultureInfo.InvariantCulture)
    End Function

    Private Function SafeTypeName(obj As Object) As String
        If obj Is Nothing Then Return "Nothing"
        Try
            Return obj.GetType().Name
        Catch
            Return "?"
        End Try
    End Function

    Private Function SafeViewName(ByVal view As SolidEdgeDraft.DrawingView) As String
        Try
            Return view.Name
        Catch
            Return "?"
        End Try
    End Function

    Private Function IsIsometricDrawingViewForDrop(ByVal dv As SolidEdgeDraft.DrawingView) As Boolean
        Try
            Dim ori As SolidEdgeConstants.ViewOrientationConstants = SolidEdgeConstants.ViewOrientationConstants.igFrontView
            Dim vx As Double = 0, vy As Double = 0, vz As Double = 0, lx As Double = 0, ly As Double = 0, lz As Double = 0
            dv.ViewOrientation(vx, vy, vz, lx, ly, lz, ori)
            Return (ori = SolidEdgeConstants.ViewOrientationConstants.igTopFrontRightView)
        Catch
            Return False
        End Try
    End Function

    Private Function SafeSheetName(ByVal sheet As SolidEdgeDraft.Sheet) As String
        If sheet Is Nothing Then Return "<nothing>"
        Try
            Return sheet.Name
        Catch
            Return "?"
        End Try
    End Function

    Private Function GetLineStart(ByVal lineObj As Object) As Pt2
        Dim x As Double = 0
        Dim y As Double = 0
        Try
            lineObj.GetStartPoint(x, y)
        Catch
            x = CDbl(CallByName(lineObj, "StartX", CallType.Get))
            y = CDbl(CallByName(lineObj, "StartY", CallType.Get))
        End Try
        Return New Pt2 With {.X = x, .Y = y}
    End Function

    Private Function GetLineEnd(ByVal lineObj As Object) As Pt2
        Dim x As Double = 0
        Dim y As Double = 0
        Try
            lineObj.GetEndPoint(x, y)
        Catch
            x = CDbl(CallByName(lineObj, "EndX", CallType.Get))
            y = CDbl(CallByName(lineObj, "EndY", CallType.Get))
        End Try
        Return New Pt2 With {.X = x, .Y = y}
    End Function

    Private Function CalcularBBoxDVLines(ByVal view As SolidEdgeDraft.DrawingView) As BBox
        Dim bbox As New BBox With {
            .MinX = Double.PositiveInfinity,
            .MinY = Double.PositiveInfinity,
            .MaxX = Double.NegativeInfinity,
            .MaxY = Double.NegativeInfinity
        }
        For Each lineObj As Object In view.DVLines2d
            Dim p1 As Pt2 = GetLineStart(lineObj)
            Dim p2 As Pt2 = GetLineEnd(lineObj)
            bbox.MinX = Math.Min(bbox.MinX, Math.Min(p1.X, p2.X))
            bbox.MinY = Math.Min(bbox.MinY, Math.Min(p1.Y, p2.Y))
            bbox.MaxX = Math.Max(bbox.MaxX, Math.Max(p1.X, p2.X))
            bbox.MaxY = Math.Max(bbox.MaxY, Math.Max(p1.Y, p2.Y))
        Next
        Return bbox
    End Function

    Private Structure Pt2
        Public X As Double
        Public Y As Double
    End Structure

    Private Structure BBox
        Public MinX As Double
        Public MinY As Double
        Public MaxX As Double
        Public MaxY As Double
    End Structure

    Private Class LineInfo
        Public Property Index As Integer
        Public Property LineObject As Object
        Public Property XMean As Double
    End Class

    Private Class DimensionResult
        Public Property Created As Boolean
    End Class

End Module
