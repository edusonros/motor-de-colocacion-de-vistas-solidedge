Option Strict Off

Imports System
Imports System.Collections.Generic
Imports System.Globalization
Imports System.IO
Imports System.Linq
Imports SolidEdgeDraft

''' <summary>Colocación de PartsList nativa en la parte superior (convención usuario: margen desde esquina superior izquierda útil; COM Solid Edge: origen hoja inferior izquierda, Y hacia arriba).</summary>
Public NotInheritable Class PartsListSuperiorService
    Private Sub New()
    End Sub

    ''' <summary>Ruta histórica: prepara zona y crea/ajusta PartsList según <see cref="DimensioningNormConfig"/>.</summary>
    Public Shared Function PrepararPartsListSuperiorExistente(
        draftDoc As DraftDocument,
        sheet As Sheet,
        modelPath As String,
        config As DimensioningNormConfig,
        log As Action(Of String),
        Optional drawingViews As IList(Of DrawingView) = Nothing) As ProtectedZone2D

        Dim Lg = Sub(m As String) log?.Invoke(m)
        If draftDoc Is Nothing OrElse sheet Is Nothing OrElse config Is Nothing Then Return Nothing

        Lg("[PARTSLIST][COORD] sheet_origin=top_left_A3_horizontal sheet_se=bottom_left_y_up")

        If Not config.EnablePartsListCreation Then
            Lg("[PARTSLIST][CREATE][SKIP] reason=EnablePartsListCreation_False")
            Return BuildHeuristicTopZoneOnly(sheet, config, Lg)
        End If

        Try
            CallByName(sheet, "Activate", CallType.Method)
        Catch
        End Try
        Try
            If draftDoc IsNot Nothing Then CallByName(draftDoc, "ActiveSheet", CallType.Let, sheet)
        Catch
        End Try

        Dim listsObj As Object = Nothing
        Try
            listsObj = CallByName(draftDoc, "PartsLists", CallType.Get)
        Catch ex As Exception
            Lg("[PARTSLIST][ERR] PartsLists: " & ex.Message)
            Return Nothing
        End Try

        Dim nLists As Integer = SafeCount(listsObj)
        Lg("[PARTSLIST][EXISTING][FOUND] count=" & nLists.ToString(CultureInfo.InvariantCulture))

        If config.DeleteExistingPartsListsBeforeCreate AndAlso nLists > 0 Then
            TryDeleteAllPartsLists(listsObj, Lg)
            nLists = SafeCount(listsObj)
            Lg("[PARTSLIST][DELETE][DONE] remaining_count=" & nLists.ToString(CultureInfo.InvariantCulture))
        End If

        Dim viewList As List(Of DrawingView) = NormalizeDrawingViewList(sheet, drawingViews)

        Dim pl As Object = Nothing
        If nLists > 0 Then
            Lg("[PARTSLIST][CREATE][SKIP] reason=existing_partslist_reposition_only")
            Try
                pl = CallByName(listsObj, "Item", CallType.Method, 1)
            Catch ex As Exception
                Lg("[PARTSLIST][ERR] Item(1): " & ex.Message)
                Return Nothing
            End Try
            If pl Is Nothing Then Return Nothing
            ApplyPartsListPlacementAndRefresh(pl, sheet, config, Lg)
        Else
            Dim z = CrearPartsListSuperiorDesdeVista(draftDoc, sheet, viewList, config, Lg)
            If z Is Nothing Then
                Lg("[PARTSLIST][CREATE][FALLBACK] reason=native_failed_heuristic_zone")
                Return BuildHeuristicTopZoneOnly(sheet, config, Lg)
            End If
            If Not String.IsNullOrWhiteSpace(modelPath) Then
                Try
                    pl = CallByName(listsObj, "Item", CallType.Method, 1)
                Catch
                    pl = Nothing
                End Try
                LogAssemblyLinkAudit(pl, modelPath, Lg)
            End If
            Return z
        End If

        If pl Is Nothing Then Return Nothing
        If Not String.IsNullOrWhiteSpace(modelPath) Then LogAssemblyLinkAudit(pl, modelPath, Lg)
        Return GetPartsListProtectedZone(pl, sheet, config, Lg, useHeuristic:=True)
    End Function

    ''' <summary>Inserta PartsList nativa vinculada a la vista principal (API <c>PartsLists.Add</c>/<c>AddEx</c> con <see cref="DrawingView"/>).</summary>
    Public Shared Function CrearPartsListSuperiorDesdeVista(
        draftDoc As DraftDocument,
        sheet As Sheet,
        drawingViews As List(Of DrawingView),
        config As DimensioningNormConfig,
        log As Action(Of String)) As ProtectedZone2D

        Dim Lg = Sub(m As String) log?.Invoke(m)
        If draftDoc Is Nothing OrElse sheet Is Nothing OrElse config Is Nothing Then Return Nothing

        Lg("[PARTSLIST][CREATE][ENTER]")

        Dim listsObj As Object = Nothing
        Try
            listsObj = CallByName(draftDoc, "PartsLists", CallType.Get)
        Catch ex As Exception
            Lg("[PARTSLIST][ADD][FAIL] PartsLists: " & ex.Message)
            Return Nothing
        End Try

        Dim viewPrincipal As DrawingView = SelectPrincipalOrthogonalView(sheet, drawingViews, Lg)
        If viewPrincipal Is Nothing Then
            Lg("[PARTSLIST][ADD][FAIL] reason=no_principal_orthogonal_view")
            Return Nothing
        End If

        Dim savedName As String = If(config.PartsListSavedSettingsName, "").Trim()
        If String.IsNullOrWhiteSpace(savedName) Then savedName = "PART_LIST"

        Dim pl As Object = Nothing
        If config.PartsListUseAddEx Then
            pl = TryPartsListsAddEx(listsObj, viewPrincipal, config, savedName, Lg)
        End If
        If pl Is Nothing Then
            Lg("[PARTSLIST][ADD][RETRY] method=Add reason=AddEx_failed_or_skipped")
            pl = TryPartsListsAdd(listsObj, viewPrincipal, config, savedName, Lg)
        End If
        If pl Is Nothing AndAlso config.AllowPartsListFallbackAnsi Then
            Lg("[PARTSLIST][ADD][FALLBACK] savedSettings=ANSI reason=AllowPartsListFallbackAnsi_diagnostic")
            pl = TryPartsListsAddEx(listsObj, viewPrincipal, config, "ANSI", Lg)
            If pl Is Nothing Then pl = TryPartsListsAdd(listsObj, viewPrincipal, config, "ANSI", Lg)
        End If

        If pl Is Nothing Then
            Lg("[PARTSLIST][ADD][FAIL] reason=AddEx_and_Add_exhausted")
            Return Nothing
        End If

        Lg("[PARTSLIST][ADD][OK]")

        Try
            CallByName(pl, "Active", CallType.Let, True)
        Catch ex As Exception
            Lg("[PARTSLIST][SET][WARN] property=Active error=" & ex.Message)
        End Try
        Try
            CallByName(pl, "ShowColumnHeader", CallType.Let, True)
        Catch ex As Exception
            Lg("[PARTSLIST][SET][WARN] property=ShowColumnHeader error=" & ex.Message)
        End Try

        ApplyPartsListTableStyleAndSavedSettings(pl, config, Lg)

        Dim oxT As Double = config.PartsListOriginX
        Dim oyT As Double = config.PartsListOriginY
        Lg("[PARTSLIST][ORIGIN][SET] x=" & FormatInv(oxT) & " y=" & FormatInv(oyT))
        TryEnsurePartsListOnActiveSheet(pl, sheet, Lg)
        Try
            CallByName(pl, "SetOrigin", CallType.Method, oxT, oyT)
        Catch ex As Exception
            Lg("[PARTSLIST][ORIGIN][FAIL] " & ex.Message)
        End Try

        Try
            CallByName(pl, "Update", CallType.Method)
            Lg("[PARTSLIST][UPDATE][OK]")
        Catch ex As Exception
            Lg("[PARTSLIST][UPDATE][FAIL] " & ex.Message)
        End Try

        Dim oxA As Double, oyA As Double
        TryGetOrigin(pl, oxA, oyA)
        Lg("[PARTSLIST][ORIGIN][AFTER] x=" & FormatInv(oxA) & " y=" & FormatInv(oyA))

        LogGetListOfSavedSettingsAllVariants(pl, config, Lg)

        LogPartsListAuditDetailed(pl, config, Lg)
        LogPartsListCellSample(pl, Lg)
        RunPartsListTemplateRefcheck(pl, config, Lg)

        Try
            CallByName(draftDoc, "UpdateAll", CallType.Method, False)
        Catch
        End Try

        Dim zone = GetPartsListProtectedZone(pl, sheet, config, Lg, useHeuristic:=True)
        Return zone
    End Function

    Private Shared Function TryPartsListsAddEx(
        listsObj As Object,
        viewPrincipal As DrawingView,
        config As DimensioningNormConfig,
        savedSettings As String,
        log As Action(Of String)) As Object

        If listsObj Is Nothing OrElse viewPrincipal Is Nothing Then Return Nothing
        Dim ab = CLng(config.PartsListAutoBalloon)
        Dim cpl = CLng(config.PartsListCreatePartsList)
        log?.Invoke("[PARTSLIST][ADD][TRY] method=AddEx savedSettings=" & savedSettings)
        Try
            Return CallByName(listsObj, "AddEx", CallType.Method, viewPrincipal, 0L, savedSettings, ab, cpl)
        Catch ex As Exception
            log?.Invoke("[PARTSLIST][ADD][FAIL] AddEx " & ex.Message)
            Return Nothing
        End Try
    End Function

    Private Shared Function TryPartsListsAdd(
        listsObj As Object,
        viewPrincipal As DrawingView,
        config As DimensioningNormConfig,
        savedSettings As String,
        log As Action(Of String)) As Object

        If listsObj Is Nothing OrElse viewPrincipal Is Nothing Then Return Nothing
        Dim ab = CLng(config.PartsListAutoBalloon)
        Dim cpl = CLng(config.PartsListCreatePartsList)
        log?.Invoke("[PARTSLIST][ADD][TRY] method=Add savedSettings=" & savedSettings)
        Try
            Return CallByName(listsObj, "Add", CallType.Method, viewPrincipal, savedSettings, ab, cpl)
        Catch ex As Exception
            log?.Invoke("[PARTSLIST][ADD][FAIL] Add " & ex.Message)
            Return Nothing
        End Try
    End Function

    Private Shared Sub ApplyPartsListTableStyleAndSavedSettings(pl As Object, config As DimensioningNormConfig, log As Action(Of String))
        If pl Is Nothing OrElse config Is Nothing Then Return
        Dim tsName = If(config.PartsListTableStyleName, "").Trim()
        Dim ssName = If(config.PartsListSavedSettingsName, "").Trim()
        If String.IsNullOrWhiteSpace(tsName) Then tsName = "PART_LIST"
        If String.IsNullOrWhiteSpace(ssName) Then ssName = "PART_LIST"

        Try
            CallByName(pl, "TableStyle", CallType.Let, tsName)
            log?.Invoke("[PARTSLIST][TABLESTYLE][OK] " & tsName)
        Catch ex As Exception
            log?.Invoke("[PARTSLIST][TABLESTYLE][WARN] " & ex.Message)
        End Try
        Try
            CallByName(pl, "SavedSettings", CallType.Let, ssName)
            log?.Invoke("[PARTSLIST][SAVEDSETTINGS][OK] " & ssName)
        Catch ex As Exception
            log?.Invoke("[PARTSLIST][SAVEDSETTINGS][WARN] " & ex.Message)
        End Try
        Try
            CallByName(pl, "Update", CallType.Method)
            log?.Invoke("[PARTSLIST][UPDATE][OK]")
        Catch ex As Exception
            log?.Invoke("[PARTSLIST][UPDATE][FAIL] " & ex.Message)
        End Try
    End Sub

    Private Shared Function NormalizeDrawingViewList(sheet As Sheet, drawingViews As IList(Of DrawingView)) As List(Of DrawingView)
        Dim result As New List(Of DrawingView)()
        If drawingViews IsNot Nothing Then
            For Each dv In drawingViews
                If dv IsNot Nothing Then result.Add(dv)
            Next
            If result.Count > 0 Then Return result
        End If
        If sheet Is Nothing Then Return result
        Try
            Dim n = sheet.DrawingViews.Count
            For i As Integer = 1 To n
                Try
                    result.Add(CType(sheet.DrawingViews.Item(i), DrawingView))
                Catch
                End Try
            Next
        Catch
        End Try
        Return result
    End Function

    Private Shared Function SelectPrincipalOrthogonalView(sheet As Sheet, drawingViews As List(Of DrawingView), log As Action(Of String)) As DrawingView
        Dim candidates As New List(Of DrawingView)()
        For Each dv In drawingViews
            If dv Is Nothing Then Continue For
            If IsIsometricLikeDrawingView(dv) Then Continue For
            candidates.Add(dv)
        Next
        If candidates.Count = 0 Then Return Nothing

        Dim best As DrawingView = Nothing
        Dim bestArea As Double = -1R
        For Each dv In candidates
            Dim a = TryViewUsefulArea(dv)
            If a > bestArea Then
                bestArea = a
                best = dv
            End If
        Next
        If best IsNot Nothing Then
            Dim nm = SafeViewName(best)
            Dim tp = SafeDrawingViewTypeString(best)
            log?.Invoke("[PARTSLIST][VIEW][SELECTED] name=" & nm & " type=" & tp & " area=" & FormatInv(bestArea))
        End If
        Return best
    End Function

    Private Shared Function TryViewUsefulArea(dv As DrawingView) As Double
        If dv Is Nothing Then Return 0R
        Try
            Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
            dv.Range(x1, y1, x2, y2)
            Dim w = Math.Abs(x2 - x1)
            Dim h = Math.Abs(y2 - y1)
            Return w * h
        Catch
            Return 0R
        End Try
    End Function

    Private Shared Function IsIsometricLikeDrawingView(dv As DrawingView) As Boolean
        If dv Is Nothing Then Return False
        Try
            Dim ori As String = Convert.ToString(CallByName(dv, "ViewOrientation", CallType.Get), CultureInfo.InvariantCulture)
            Dim dvtStr As String = Convert.ToString(CallByName(dv, "DrawingViewType", CallType.Get), CultureInfo.InvariantCulture)
            Dim dvtNum As Integer = 0
            Try
                dvtNum = Convert.ToInt32(CallByName(dv, "DrawingViewType", CallType.Get), CultureInfo.InvariantCulture)
            Catch
                dvtNum = 0
            End Try
            If ori.IndexOf("iso", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
               ori.IndexOf("topfrontright", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
               dvtStr.IndexOf("iso", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
               (dvtNum > 0 AndAlso dvtNum <> 1) Then
                Return True
            End If
        Catch
        End Try
        Return False
    End Function

    Private Shared Function SafeViewName(dv As DrawingView) As String
        If dv Is Nothing Then Return "?"
        Try
            Return Convert.ToString(CallByName(dv, "Name", CallType.Get), CultureInfo.InvariantCulture)
        Catch
            Return "?"
        End Try
    End Function

    Private Shared Function SafeDrawingViewTypeString(dv As DrawingView) As String
        If dv Is Nothing Then Return "?"
        Try
            Return Convert.ToString(CallByName(dv, "DrawingViewType", CallType.Get), CultureInfo.InvariantCulture)
        Catch
            Return "?"
        End Try
    End Function

    Private Shared Sub TryDeleteAllPartsLists(listsObj As Object, log As Action(Of String))
        If listsObj Is Nothing Then Return
        Dim guard As Integer = 0
        While SafeCount(listsObj) > 0 AndAlso guard < 64
            guard += 1
            Try
                Dim pl = CallByName(listsObj, "Item", CallType.Method, 1)
                If pl Is Nothing Then Exit While
                CallByName(pl, "Delete", CallType.Method)
                log?.Invoke("[PARTSLIST][DELETE] item=1 ok")
            Catch ex As Exception
                log?.Invoke("[PARTSLIST][DELETE][FAIL] " & ex.Message)
                Exit While
            End Try
        End While
    End Sub

    Private Shared Sub ApplyPartsListPlacementAndRefresh(pl As Object, sheet As Sheet, config As DimensioningNormConfig, log As Action(Of String))
        If pl Is Nothing OrElse config Is Nothing Then Return
        log?.Invoke("[PARTSLIST][EXISTING][REPOSITION] using PartsListOriginX/Y")
        Try
            CallByName(pl, "Active", CallType.Let, True)
        Catch
        End Try
        Try
            CallByName(pl, "ShowColumnHeader", CallType.Let, True)
        Catch
        End Try
        ApplyPartsListTableStyleAndSavedSettings(pl, config, log)
        Dim oxT = config.PartsListOriginX
        Dim oyT = config.PartsListOriginY
        log?.Invoke("[PARTSLIST][ORIGIN][SET] x=" & FormatInv(oxT) & " y=" & FormatInv(oyT))
        TryEnsurePartsListOnActiveSheet(pl, sheet, log)
        Try
            CallByName(pl, "SetOrigin", CallType.Method, oxT, oyT)
        Catch ex As Exception
            log?.Invoke("[PARTSLIST][ORIGIN][FAIL] " & ex.Message)
        End Try
        Try
            CallByName(pl, "Update", CallType.Method)
            log?.Invoke("[PARTSLIST][UPDATE][OK]")
        Catch ex As Exception
            log?.Invoke("[PARTSLIST][UPDATE][FAIL] " & ex.Message)
        End Try
        Dim oxA As Double, oyA As Double
        TryGetOrigin(pl, oxA, oyA)
        log?.Invoke("[PARTSLIST][ORIGIN][AFTER] x=" & FormatInv(oxA) & " y=" & FormatInv(oyA))
        LogGetListOfSavedSettingsAllVariants(pl, config, log)
        LogPartsListAuditDetailed(pl, config, log)
        LogPartsListCellSample(pl, log)
        RunPartsListTemplateRefcheck(pl, config, log)
    End Sub

    Private Shared Sub LogGetListOfSavedSettingsAllVariants(pl As Object, config As DimensioningNormConfig, log As Action(Of String))
        If pl Is Nothing OrElse config Is Nothing Then Return
        Dim marker = If(config.PartsListSavedSettingsName, "PART_LIST").Trim()
        If String.IsNullOrWhiteSpace(marker) Then marker = "PART_LIST"

        Dim foundAny As Boolean = False

        log?.Invoke("[PARTSLIST][SETTINGS][TRY_A]")
        Try
            Dim cntA As Long = 0
            Dim arrA As String() = Nothing
            pl.GetListOfSavedSettings(cntA, arrA)
            log?.Invoke("[PARTSLIST][SETTINGS][COUNT] try=A count=" & cntA.ToString(CultureInfo.InvariantCulture))
            If LogSavedSettingsArrayItems(arrA, marker, log) Then foundAny = True
        Catch ex As Exception
            log?.Invoke("[PARTSLIST][SETTINGS][TRY_A][ERR] " & ex.Message)
        End Try

        log?.Invoke("[PARTSLIST][SETTINGS][TRY_B]")
        Try
            Dim cntB As Object = CLng(0)
            Dim arrB As Object = Nothing
            CallByName(pl, "GetListOfSavedSettings", CallType.Method, cntB, arrB)
            Dim cLong As Long = Convert.ToInt64(cntB, CultureInfo.InvariantCulture)
            log?.Invoke("[PARTSLIST][SETTINGS][COUNT] try=B count=" & cLong.ToString(CultureInfo.InvariantCulture))
            If LogSavedSettingsObjectAsArray(arrB, marker, log) Then foundAny = True
        Catch ex As Exception
            log?.Invoke("[PARTSLIST][SETTINGS][TRY_B][ERR] " & ex.Message)
        End Try

        log?.Invoke("[PARTSLIST][SETTINGS][TRY_C]")
        Try
            Dim cntC As Long = 0
            Dim arrC(100) As String
            pl.GetListOfSavedSettings(cntC, arrC)
            log?.Invoke("[PARTSLIST][SETTINGS][COUNT] try=C count=" & cntC.ToString(CultureInfo.InvariantCulture))
            If LogSavedSettingsFixedStringArray(arrC, cntC, marker, log) Then foundAny = True
        Catch ex As Exception
            log?.Invoke("[PARTSLIST][SETTINGS][TRY_C][ERR] " & ex.Message)
        End Try

        log?.Invoke("[PARTSLIST][SETTINGS][FOUND_PART_LIST] " & foundAny.ToString(CultureInfo.InvariantCulture))
    End Sub

    Private Shared Function LogSavedSettingsArrayItems(arr As String(), marker As String, log As Action(Of String)) As Boolean
        Dim found As Boolean = False
        If arr Is Nothing Then Return False
        For i As Integer = 0 To arr.Length - 1
            Dim nm = If(arr(i), "").Trim()
            If String.IsNullOrEmpty(nm) Then Continue For
            log?.Invoke("[PARTSLIST][SETTINGS][ITEM] name=" & nm)
            If String.Equals(nm, marker, StringComparison.OrdinalIgnoreCase) Then found = True
        Next
        Return found
    End Function

    Private Shared Function LogSavedSettingsObjectAsArray(arrObj As Object, marker As String, log As Action(Of String)) As Boolean
        If arrObj Is Nothing Then Return False
        Dim a = TryCast(arrObj, Array)
        If a Is Nothing Then Return False
        Dim found As Boolean = False
        For Each o In a
            Dim nm = Convert.ToString(o, CultureInfo.InvariantCulture).Trim()
            If String.IsNullOrEmpty(nm) Then Continue For
            log?.Invoke("[PARTSLIST][SETTINGS][ITEM] name=" & nm)
            If String.Equals(nm, marker, StringComparison.OrdinalIgnoreCase) Then found = True
        Next
        Return found
    End Function

    Private Shared Function LogSavedSettingsFixedStringArray(arr As String(), count As Long, marker As String, log As Action(Of String)) As Boolean
        If arr Is Nothing Then Return False
        Dim n As Integer = CInt(Math.Min(count, arr.Length))
        If n < 0 Then n = 0
        Dim found As Boolean = False
        For i As Integer = 0 To n - 1
            Dim nm = If(arr(i), "").Trim()
            If String.IsNullOrEmpty(nm) Then Continue For
            log?.Invoke("[PARTSLIST][SETTINGS][ITEM] name=" & nm)
            If String.Equals(nm, marker, StringComparison.OrdinalIgnoreCase) Then found = True
        Next
        Return found
    End Function

    Private Shared Sub LogPartsListAuditDetailed(pl As Object, config As DimensioningNormConfig, log As Action(Of String))
        If pl Is Nothing Then Return
        Dim rows As Integer = SafeCount(CallByNameSafe(pl, "Rows"))
        Dim cols As Integer = SafeCount(CallByNameSafe(pl, "Columns"))
        Dim up As String = SafeProp(pl, "IsUpToDate")
        Dim tsDisplay = TryGetTableStyleDisplayName(pl)
        Dim savedDisp = SafeProp(pl, "SavedSettings")
        log?.Invoke("[PARTSLIST][AUDIT] rows=" & rows.ToString(CultureInfo.InvariantCulture) &
                    " cols=" & cols.ToString(CultureInfo.InvariantCulture) &
                    " isUpToDate=" & up &
                    " tableStyle=" & tsDisplay &
                    " savedSettingsProp=" & savedDisp)
        Dim asm As String = SafeProp(pl, "AssemblyFileName")
        log?.Invoke("[PARTSLIST][EXISTING][ASSEMBLYFILE] " & asm)
    End Sub

    Private Shared Function TryGetTableStyleDisplayName(pl As Object) As String
        If pl Is Nothing Then Return ""
        Try
            Dim ts = CallByName(pl, "TableStyle", CallType.Get)
            If ts Is Nothing Then Return ""
            Dim nm = Convert.ToString(CallByName(ts, "Name", CallType.Get), CultureInfo.InvariantCulture)
            If Not String.IsNullOrWhiteSpace(nm) Then Return nm.Trim()
        Catch
        End Try
        Try
            Return Convert.ToString(CallByName(pl, "TableStyle", CallType.Get), CultureInfo.InvariantCulture).Trim()
        Catch
            Return ""
        End Try
    End Function

    Private Shared Sub RunPartsListTemplateRefcheck(pl As Object, config As DimensioningNormConfig, log As Action(Of String))
        If pl Is Nothing OrElse config Is Nothing Then Return
        Dim wantStyle = If(config.PartsListTableStyleName, "PART_LIST").Trim()
        Dim wantSaved = If(config.PartsListSavedSettingsName, "PART_LIST").Trim()
        If String.IsNullOrWhiteSpace(wantStyle) Then wantStyle = "PART_LIST"
        If String.IsNullOrWhiteSpace(wantSaved) Then wantSaved = "PART_LIST"

        Dim styleActual = TryGetTableStyleDisplayName(pl)
        Dim styleOk = String.Equals(styleActual, wantStyle, StringComparison.OrdinalIgnoreCase)

        Dim cols As Integer = SafeCount(CallByNameSafe(pl, "Columns"))
        Dim minCols As Integer = Math.Max(7, config.PartsListMinExpectedColumns)
        Dim colsOk As Boolean = cols >= minCols

        Dim savedProp = SafeProp(pl, "SavedSettings")
        Dim savedOk As Boolean = Not String.IsNullOrWhiteSpace(savedProp) AndAlso
            String.Equals(savedProp.Trim(), wantSaved, StringComparison.OrdinalIgnoreCase)

        If Not styleOk Then
            log?.Invoke("[PARTSLIST][REFCHECK][ERROR] wrong_table_style actual=" & styleActual & " expected=" & wantStyle)
        End If
        If Not colsOk Then
            log?.Invoke("[PARTSLIST][REFCHECK][ERROR] wrong_saved_settings_or_table_style cols=" & cols.ToString(CultureInfo.InvariantCulture) & " minExpected=" & minCols.ToString(CultureInfo.InvariantCulture))
        End If
        If String.IsNullOrWhiteSpace(savedProp) OrElse Not savedOk Then
            log?.Invoke("[PARTSLIST][REFCHECK][WARN] saved_settings_not_confirmed actual=" & savedProp & " expected=" & wantSaved)
        End If

        If styleOk AndAlso colsOk Then
            Dim savedOut As String = If(savedOk, wantSaved, If(String.IsNullOrWhiteSpace(savedProp), "?", savedProp))
            log?.Invoke("[PARTSLIST][REFCHECK][OK] style=" & wantStyle & " savedSettings=" & savedOut)
        End If
    End Sub

    ''' <summary>Validación de plantilla PART_LIST para audit final (mismo criterio que tras crear).</summary>
    Public Shared Sub RunPartsListStructureRefcheck(pl As Object, config As DimensioningNormConfig, log As Action(Of String))
        RunPartsListTemplateRefcheck(pl, config, log)
    End Sub

    Private Shared Sub LogPartsListCellSample(pl As Object, log As Action(Of String))
        If pl Is Nothing Then Return
        Dim nRows As Integer = SafeCount(CallByNameSafe(pl, "Rows"))
        Dim nCols As Integer = SafeCount(CallByNameSafe(pl, "Columns"))
        Dim maxR = Math.Min(Math.Max(nRows, 0), 3)
        Dim maxC = Math.Min(Math.Max(nCols, 0), 6)
        If maxR <= 0 OrElse maxC <= 0 Then Return
        Dim parts As New List(Of String)()
        For r As Integer = 1 To maxR
            For c As Integer = 1 To maxC
                Dim t = TryGetPartsListCellText(pl, r, c)
                If Not String.IsNullOrWhiteSpace(t) Then parts.Add("r" & r.ToString(CultureInfo.InvariantCulture) & "c" & c.ToString(CultureInfo.InvariantCulture) & "=" & t)
            Next
        Next
        If parts.Count > 0 Then
            log?.Invoke("[PARTSLIST][CELLS][SAMPLE] " & String.Join("; ", parts.Take(12)))
        End If
    End Sub

    Private Shared Function TryGetPartsListCellText(pl As Object, row1 As Integer, col1 As Integer) As String
        If pl Is Nothing Then Return ""
        Dim cellObj As Object = Nothing
        Try
            cellObj = CallByName(pl, "Cell", CallType.Get, row1, col1)
        Catch
            cellObj = Nothing
        End Try
        If cellObj Is Nothing Then
            Try
                cellObj = CallByName(pl, "Cell", CallType.Method, row1, col1)
            Catch
                cellObj = Nothing
            End Try
        End If
        If cellObj Is Nothing Then Return ""
        Try
            Dim s As String = Convert.ToString(CallByName(cellObj, "Text", CallType.Get), CultureInfo.InvariantCulture)
            If String.IsNullOrWhiteSpace(s) Then s = Convert.ToString(CallByName(cellObj, "Value", CallType.Get), CultureInfo.InvariantCulture)
            If Not String.IsNullOrWhiteSpace(s) Then Return Truncate(s, 48)
        Catch
        End Try
        Return ""
    End Function

    Private Shared Function Truncate(s As String, maxLen As Integer) As String
        If s Is Nothing Then Return ""
        If s.Length <= maxLen Then Return s
        Return s.Substring(0, maxLen) & "..."
    End Function

    Private Shared Sub LogAssemblyLinkAudit(pl As Object, modelPath As String, log As Action(Of String))
        If pl Is Nothing OrElse String.IsNullOrWhiteSpace(modelPath) Then Return
        Dim asm As String = SafeProp(pl, "AssemblyFileName")
        If String.IsNullOrWhiteSpace(asm) Then
            log?.Invoke("[PARTSLIST][REF][WARN] assembly_empty")
            Return
        End If
        If ModelPathLikelyMatches(asm, modelPath) Then
            log?.Invoke("[PARTSLIST][REF][OK] assembly_linked")
        Else
            log?.Invoke("[PARTSLIST][REF][WARN] assembly_mismatch model=" & modelPath & " partsList=" & asm)
        End If
    End Sub

    Private Shared Function ModelPathLikelyMatches(assemblyFile As String, modelPath As String) As Boolean
        Try
            Dim a = Path.GetFullPath(assemblyFile)
            Dim b = Path.GetFullPath(modelPath)
            Return String.Equals(a, b, StringComparison.OrdinalIgnoreCase)
        Catch
            Return String.Equals(Path.GetFileName(assemblyFile), Path.GetFileName(modelPath), StringComparison.OrdinalIgnoreCase)
        End Try
    End Function

    Public Shared Function GetPartsListProtectedZone(
        pl As Object,
        sheet As Sheet,
        config As DimensioningNormConfig,
        log As Action(Of String),
        Optional useHeuristic As Boolean = True) As ProtectedZone2D

        Dim Lg = Sub(m As String) log?.Invoke(m)
        If pl Is Nothing OrElse config Is Nothing Then Return Nothing

        Dim ox As Double, oy As Double
        TryGetOrigin(pl, ox, oy)
        Dim estimatedHeight As Double = Math.Max(0.012R, config.PartsListEstimatedHeightM)
        Dim estimatedWidth As Double = 0.31R
        Dim mrg As Double = config.ProtectedZoneSafetyMarginM

        Dim yTop As Double = oy
        Dim yBot As Double = oy - estimatedHeight
        Dim minY As Double = Math.Min(yTop, yBot)
        Dim maxY As Double = Math.Max(yTop, yBot)

        Dim z As New ProtectedZone2D With {
            .Name = "PartsListTop",
            .MinX = ox - mrg,
            .MinY = minY - mrg,
            .MaxX = ox + estimatedWidth + mrg,
            .MaxY = maxY + mrg
        }

        ClampZoneToSheet(z, sheet)
        Lg("[ZONE][PARTSLIST] minX=" & FormatInv(z.MinX) & " minY=" & FormatInv(z.MinY) &
           " maxX=" & FormatInv(z.MaxX) & " maxY=" & FormatInv(z.MaxY))
        Return z
    End Function

    Public Shared Function DetectarZonasProtegidas(
        sheet As Sheet,
        partsListZone As ProtectedZone2D,
        config As DimensioningNormConfig,
        log As Action(Of String)) As List(Of ProtectedZone2D)

        Dim Lg = Sub(m As String) log?.Invoke(m)
        Dim list As New List(Of ProtectedZone2D)()
        If partsListZone IsNot Nothing Then list.Add(partsListZone)

        Dim sw As BoundingBox2D = DimensionPlacementService.TryGetSheetWorkArea(sheet, Lg)
        Dim w As Double = If(sw IsNot Nothing AndAlso sw.Width > 0.01R, sw.Width, 0.42R)
        Dim h As Double = If(sw IsNot Nothing AndAlso sw.Height > 0.01R, sw.Height, 0.297R)

        Dim title As New ProtectedZone2D With {
            .Name = "TitleBlock",
            .MinX = 0.25R,
            .MinY = 0.245R,
            .MaxX = 0.415R,
            .MaxY = 0.292R
        }
        Lg("[ZONE][TITLEBLOCK] minX=" & FormatInv(title.MinX) & " minY=" & FormatInv(title.MinY) &
           " maxX=" & FormatInv(title.MaxX) & " maxY=" & FormatInv(title.MaxY))
        list.Add(title)

        Dim rev As New ProtectedZone2D With {
            .Name = "RevisionTable",
            .MinX = 0.21R,
            .MinY = 0.26R,
            .MaxX = 0.32R,
            .MaxY = 0.292R
        }
        Lg("[ZONE][REVISION] minX=" & FormatInv(rev.MinX) & " minY=" & FormatInv(rev.MinY) &
           " maxX=" & FormatInv(rev.MaxX) & " maxY=" & FormatInv(rev.MaxY))
        list.Add(rev)

        Lg("[ZONE][BORDER] sheetApproxW=" & FormatInv(w) & " sheetApproxH=" & FormatInv(h) & " margin=0.005 (UNE129+AvoidBorder)")
        Return list
    End Function

    Public Shared Sub TryNudgeViewsOutsideZone(sheet As Sheet, zone As ProtectedZone2D, log As Action(Of String))
        Dim Lg = Sub(m As String) log?.Invoke(m)
        If sheet Is Nothing OrElse zone Is Nothing Then Return
        Dim cnt As Integer = 0
        Try
            cnt = sheet.DrawingViews.Count
        Catch
            Return
        End Try

        For i As Integer = 1 To cnt
            Dim dv As DrawingView = Nothing
            Try
                dv = CType(sheet.DrawingViews.Item(i), DrawingView)
            Catch
                dv = Nothing
            End Try
            If dv Is Nothing Then Continue For

            Dim vx1 As Double, vy1 As Double, vx2 As Double, vy2 As Double
            Try
                dv.Range(vx1, vy1, vx2, vy2)
            Catch
                Continue For
            End Try
            Dim vMinX = Math.Min(vx1, vx2), vMaxX = Math.Max(vx1, vx2)
            Dim vMinY = Math.Min(vy1, vy2), vMaxY = Math.Max(vy1, vy2)

            If Not RectsOverlap(vMinX, vMaxX, vMinY, vMaxY, zone.MinX, zone.MaxX, zone.MinY, zone.MaxY) Then Continue For

            Dim overlapY As Double = Math.Min(vMaxY, zone.MaxY) - Math.Max(vMinY, zone.MinY)
            If overlapY <= 0R Then Continue For

            Dim ox As Double, oy As Double
            Try
                dv.GetOrigin(ox, oy)
            Catch
                Continue For
            End Try

            Dim dy As Double = -(overlapY + 0.003R)
            Try
                dv.SetOrigin(ox, oy + dy)
                Lg("[DIM][PLACE][MOVE] reason=view_invaded_PartsListTop viewIdx=" & i.ToString(CultureInfo.InvariantCulture) & " dy=" & FormatInv(dy))
            Catch ex As Exception
                Lg("[DIM][PLACE][WARN] nudge_view_fail idx=" & i.ToString(CultureInfo.InvariantCulture) & " " & ex.Message)
            End Try
        Next
    End Sub

    Private Shared Function BuildHeuristicTopZoneOnly(sheet As Sheet, config As DimensioningNormConfig, log As Action(Of String)) As ProtectedZone2D
        If config Is Nothing Then Return Nothing
        Dim sheetH As Double = ReadSheetHeightM(sheet)
        Dim mrg As Double = config.ProtectedZoneSafetyMarginM
        Dim estH0 As Double = Math.Max(0.012R, config.PartsListEstimatedHeightM)
        Dim z As New ProtectedZone2D With {
            .Name = "PartsListTop",
            .MinX = config.PartsListMarginXm - mrg,
            .MinY = sheetH - config.PartsListMarginYm - estH0 - mrg,
            .MaxX = config.PartsListMarginXm + 0.31R + mrg,
            .MaxY = sheetH - config.PartsListMarginYm + mrg
        }
        ClampZoneToSheet(z, sheet)
        log?.Invoke("[ZONE][PARTSLIST_TOP][ESTIMATED] minX=" & FormatInv(z.MinX) & " minY=" & FormatInv(z.MinY) &
                    " maxX=" & FormatInv(z.MaxX) & " maxY=" & FormatInv(z.MaxY))
        Return z
    End Function

    Private Shared Sub TryEnsurePartsListOnActiveSheet(pl As Object, sheet As Sheet, log As Action(Of String))
        If pl Is Nothing OrElse sheet Is Nothing Then Return
        Try
            Dim par As Object = CallByName(pl, "Parent", CallType.Get)
            If TypeOf par Is Sheet AndAlso Not Object.ReferenceEquals(par, sheet) Then
                log?.Invoke("[PARTSLIST][WARN] PartsList.Parent <> ActiveSheet; intentando MoveToSheet")
                Try
                    CallByName(pl, "MoveToSheet", CallType.Method, sheet)
                    log?.Invoke("[PARTSLIST][MOVE][OK] MoveToSheet")
                Catch ex As Exception
                    log?.Invoke("[PARTSLIST][MOVE][WARN] " & ex.Message)
                End Try
            End If
        Catch ex As Exception
            log?.Invoke("[PARTSLIST][PARENT][WARN] " & ex.Message)
        End Try
    End Sub

    Private Shared Sub TryGetOrigin(pl As Object, ByRef x As Double, ByRef y As Double)
        x = 0R : y = 0R
        If pl Is Nothing Then Return
        Try
            CallByName(pl, "GetOrigin", CallType.Method, x, y)
        Catch
        End Try
    End Sub

    Private Shared Function SafeProp(o As Object, name As String) As String
        If o Is Nothing OrElse name Is Nothing Then Return ""
        Try
            Return Convert.ToString(CallByName(o, name, CallType.Get), CultureInfo.InvariantCulture)
        Catch
            Return ""
        End Try
    End Function

    Private Shared Function CallByNameSafe(o As Object, m As String) As Object
        Try
            Return CallByName(o, m, CallType.Get)
        Catch
            Return Nothing
        End Try
    End Function

    Private Shared Function SafeCount(o As Object) As Integer
        If o Is Nothing Then Return 0
        Try
            Return CInt(CallByName(o, "Count", CallType.Get))
        Catch
            Return 0
        End Try
    End Function

    Private Shared Function FormatInv(v As Double) As String
        Return v.ToString("0.#####", CultureInfo.InvariantCulture)
    End Function

    Private Shared Function ReadSheetHeightM(sheet As Sheet) As Double
        If sheet Is Nothing Then Return 0.297R
        Try
            Dim su = sheet.SheetSetup
            If su IsNot Nothing Then Return su.SheetHeight
        Catch
        End Try
        Return 0.297R
    End Function

    Private Shared Sub ClampZoneToSheet(z As ProtectedZone2D, sheet As Sheet)
        If z Is Nothing Then Return
        Dim sw As BoundingBox2D = DimensionPlacementService.TryGetSheetWorkArea(
            sheet,
            Sub(s As String)
            End Sub)
        If sw Is Nothing Then Return
        z.MinX = Math.Max(sw.MinX, Math.Min(z.MinX, sw.MaxX))
        z.MaxX = Math.Min(sw.MaxX, Math.Max(z.MaxX, sw.MinX))
        z.MinY = Math.Max(sw.MinY, Math.Min(z.MinY, sw.MaxY))
        z.MaxY = Math.Min(sw.MaxY, Math.Max(z.MaxY, sw.MinY))
    End Sub

    Private Shared Function RectsOverlap(ax1 As Double, ax2 As Double, ay1 As Double, ay2 As Double, bx1 As Double, bx2 As Double, by1 As Double, by2 As Double) As Boolean
        Return ax1 < bx2 AndAlso ax2 > bx1 AndAlso ay1 < by2 AndAlso ay2 > by1
    End Function
End Class
