Option Strict Off

Imports System.Globalization
Imports System.IO
Imports System.Linq
Imports System.Runtime.InteropServices
Imports System.Text
Imports SolidEdgeDraft
Imports SolidEdgeFramework

Public NotInheritable Class DftAuditService
    Private Sub New()
    End Sub

    Public Shared Function AnalyzeDft(dftPath As String, logger As Logger, keepSolidEdgeVisible As Boolean, outputFolder As String) As Boolean
        If String.IsNullOrWhiteSpace(dftPath) OrElse Not File.Exists(dftPath) Then
            logger?.Log("[DFT][AUDIT][ERR] Ruta DFT inválida.")
            Return False
        End If

        Dim app As Application = Nothing
        Dim createdApp As Boolean = False
        Dim dft As DraftDocument = Nothing
        Dim sb As New StringBuilder()

        Try
            OleMessageFilter.Register()
            ConnectSolidEdge(app, createdApp, keepSolidEdgeVisible, logger)
            If app Is Nothing Then
                logger?.Log("[DFT][AUDIT][ERR] No se pudo obtener instancia de Solid Edge.")
                Return False
            End If

            dft = CType(app.Documents.Open(dftPath), DraftDocument)
            If dft Is Nothing Then
                logger?.Log("[DFT][AUDIT][ERR] No se pudo abrir el DFT.")
                Return False
            End If

            sb.AppendLine("=== DFT AUDIT ===")
            sb.AppendLine("File=" & dftPath)
            sb.AppendLine("GeneratedUtc=" & DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture))
            sb.AppendLine("Name=" & SafeToString(CallByNameSafe(dft, "Name")))
            sb.AppendLine("FullName=" & SafeToString(CallByNameSafe(dft, "FullName")))

            Dim sheets As Sheets = dft.Sheets
            Dim totalSheets As Integer = SafeCount(sheets)
            Dim totalViews As Integer = 0
            Dim totalDims As Integer = 0
            Dim totalDvLines As Integer = 0
            Dim totalDvArcs As Integer = 0
            Dim totalDvCircles As Integer = 0
            Dim totalDvSplines As Integer = 0
            Dim totalDvLineStrings As Integer = 0
            Dim totalDvPoints As Integer = 0
            Dim totalDraftTables As Integer = 0
            Dim totalPartsLists As Integer = 0

            sb.AppendLine("Sheets=" & totalSheets.ToString(CultureInfo.InvariantCulture))

            totalDraftTables = AppendDraftDocumentTablesSection(sb, dft)
            totalPartsLists = AppendDraftDocumentPartsListsSection(sb, dft)

            For i As Integer = 1 To totalSheets
                Dim sh As Sheet = Nothing
                Try
                    sh = CType(sheets.Item(i), Sheet)
                Catch
                    sh = Nothing
                End Try
                If sh Is Nothing Then Continue For

                Dim shName As String = SafeToString(CallByNameSafe(sh, "Name"))
                Dim dimsObj As Object = CallByNameSafe(sh, "Dimensions")
                Dim dimsCount As Integer = SafeCount(dimsObj)
                Dim views As DrawingViews = Nothing
                Try
                    views = sh.DrawingViews
                Catch
                    views = Nothing
                End Try
                Dim viewCount As Integer = SafeCount(views)

                totalDims += dimsCount
                totalViews += viewCount

                sb.AppendLine("")
                sb.AppendLine(String.Format(CultureInfo.InvariantCulture, "[SHEET] idx={0} name={1} views={2} dimensions={3}", i, shName, viewCount, dimsCount))
                AppendDimensionDetails(sb, dimsObj)

                For v As Integer = 1 To viewCount
                    Dim dv As DrawingView = Nothing
                    Try
                        dv = CType(views.Item(v), DrawingView)
                    Catch
                        dv = Nothing
                    End Try
                    If dv Is Nothing Then Continue For

                    Dim dvName As String = SafeToString(CallByNameSafe(dv, "Name"))
                    Dim dvType As String = SafeToString(CallByNameSafe(dv, "DrawingViewType"))
                    Dim dvOri As String = SafeToString(CallByNameSafe(dv, "ViewOrientation"))
                    Dim nLines As Integer = SafeCount(CallByNameSafe(dv, "DVLines2d"))
                    Dim nArcs As Integer = SafeCount(CallByNameSafe(dv, "DVArcs2d"))
                    Dim nCircles As Integer = SafeCount(CallByNameSafe(dv, "DVCircles2d"))
                    Dim nSplines As Integer = SafeCount(CallByNameSafe(dv, "DVBSplineCurves2d"))
                    Dim nLineStrings As Integer = SafeCount(CallByNameSafe(dv, "DVLineStrings2d"))
                    Dim nPoints As Integer = SafeCount(CallByNameSafe(dv, "DVPoints2d"))

                    totalDvLines += nLines
                    totalDvArcs += nArcs
                    totalDvCircles += nCircles
                    totalDvSplines += nSplines
                    totalDvLineStrings += nLineStrings
                    totalDvPoints += nPoints

                    sb.AppendLine(String.Format(CultureInfo.InvariantCulture,
                        "  [VIEW] idx={0} name={1} type={2} orientation={3} lines={4} arcs={5} circles={6} splines={7} linestrings={8} points={9}",
                        v, dvName, dvType, dvOri, nLines, nArcs, nCircles, nSplines, nLineStrings, nPoints))
                Next
            Next

            sb.AppendLine("")
            sb.AppendLine("=== SUMMARY ===")
            sb.AppendLine("TotalSheets=" & totalSheets.ToString(CultureInfo.InvariantCulture))
            sb.AppendLine("TotalViews=" & totalViews.ToString(CultureInfo.InvariantCulture))
            sb.AppendLine("TotalDimensions=" & totalDims.ToString(CultureInfo.InvariantCulture))
            sb.AppendLine("TotalDVLines2d=" & totalDvLines.ToString(CultureInfo.InvariantCulture))
            sb.AppendLine("TotalDVArcs2d=" & totalDvArcs.ToString(CultureInfo.InvariantCulture))
            sb.AppendLine("TotalDVCircles2d=" & totalDvCircles.ToString(CultureInfo.InvariantCulture))
            sb.AppendLine("TotalDVBSplineCurves2d=" & totalDvSplines.ToString(CultureInfo.InvariantCulture))
            sb.AppendLine("TotalDVLineStrings2d=" & totalDvLineStrings.ToString(CultureInfo.InvariantCulture))
            sb.AppendLine("TotalDVPoints2d=" & totalDvPoints.ToString(CultureInfo.InvariantCulture))
            sb.AppendLine("TotalDraftTables=" & totalDraftTables.ToString(CultureInfo.InvariantCulture))
            sb.AppendLine("TotalPartsLists=" & totalPartsLists.ToString(CultureInfo.InvariantCulture))

            Dim outRoot As String = If(String.IsNullOrWhiteSpace(outputFolder), Path.GetDirectoryName(dftPath), outputFolder)
            Dim outDir As String = Path.Combine(outRoot, "DFT_INSPECT")
            Directory.CreateDirectory(outDir)
            Dim outFile As String = Path.Combine(outDir, Path.GetFileNameWithoutExtension(dftPath) & "_AUDIT.txt")
            File.WriteAllText(outFile, sb.ToString(), Encoding.UTF8)

            logger?.Log("[DFT][AUDIT][OK] Informe generado: " & outFile)
            Return True
        Catch ex As Exception
            logger?.LogException("DftAuditService.AnalyzeDft", ex)
            Return False
        Finally
            Try
                If dft IsNot Nothing Then dft.Close(False)
            Catch
            End Try
            TryReleaseComObject(dft)
            If createdApp AndAlso app IsNot Nothing Then
                Try
                    app.Quit()
                Catch
                End Try
            End If
            TryReleaseComObject(app)
            Try
                OleMessageFilter.Revoke()
            Catch
            End Try
        End Try
    End Function

    Private Shared Sub ConnectSolidEdge(ByRef app As Application, ByRef createdByUs As Boolean, keepVisible As Boolean, logger As Logger)
        app = Nothing
        createdByUs = False
        Try
            app = CType(Marshal.GetActiveObject("SolidEdge.Application"), Application)
        Catch
            Dim t = Type.GetTypeFromProgID("SolidEdge.Application")
            If t IsNot Nothing Then
                app = CType(Activator.CreateInstance(t), Application)
                createdByUs = True
            End If
        End Try
        If app IsNot Nothing Then
            Try : app.Visible = keepVisible : Catch : End Try
            Try : app.DisplayAlerts = False : Catch : End Try
            logger?.Log("[DFT][AUDIT] Conectado a Solid Edge.")
        End If
    End Sub

    Private Shared Function SafeCount(obj As Object) As Integer
        If obj Is Nothing Then Return 0
        Try
            Return CInt(CallByName(obj, "Count", CallType.Get))
        Catch
            Return 0
        End Try
    End Function

    Private Shared Function SafeInt32(o As Object) As Integer
        If o Is Nothing Then Return 0
        Try
            Return Convert.ToInt32(o, CultureInfo.InvariantCulture)
        Catch
            Return 0
        End Try
    End Function

    ''' <summary>Inventaria <see cref="DraftDocument.Tables"/> (tablas de usuario del pliego: materiales, piezas, etc.).</summary>
    Private Shared Function AppendDraftDocumentTablesSection(sb As StringBuilder, dft As DraftDocument) As Integer
        If sb Is Nothing OrElse dft Is Nothing Then Return 0
        Dim tablesObj As Object = Nothing
        Try
            tablesObj = CallByName(dft, "Tables", CallType.Get)
        Catch ex As Exception
            sb.AppendLine("")
            sb.AppendLine("[TABLES] DraftDocument.Tables no disponible en este entorno: " & ex.Message)
            Return 0
        End Try
        Dim n As Integer = SafeCount(tablesObj)
        sb.AppendLine("")
        sb.AppendLine(String.Format(CultureInfo.InvariantCulture,
            "[TABLES] DraftDocument.Tables Count={0} (colección de tablas de usuario en el documento; p. ej. tabla de materiales)",
            n))
        For ti As Integer = 1 To n
            Dim tbl As Object = Nothing
            Try
                tbl = CallByName(tablesObj, "Item", CallType.Method, ti)
            Catch
                tbl = Nothing
            End Try
            If tbl Is Nothing Then Continue For

            Dim tName As String = SafeToString(CallByNameSafe(tbl, "Name"))
            Dim tTitle As String = SafeToString(CallByNameSafe(tbl, "Title"))
            Dim tType As String = SafeToString(CallByNameSafe(tbl, "Type"))
            If String.IsNullOrWhiteSpace(tType) Then tType = SafeToString(CallByNameSafe(tbl, "TableType"))

            Dim nRows As Integer = SafeInt32(CallByNameSafe(tbl, "RowCount"))
            If nRows <= 0 Then nRows = SafeCount(CallByNameSafe(tbl, "Rows"))
            Dim nCols As Integer = SafeInt32(CallByNameSafe(tbl, "ColumnCount"))
            If nCols <= 0 Then nCols = SafeCount(CallByNameSafe(tbl, "Columns"))

            Dim sheetName As String = ""
            Try
                Dim shObj As Object = CallByNameSafe(tbl, "Sheet")
                sheetName = SafeToString(CallByNameSafe(shObj, "Name"))
            Catch
                sheetName = ""
            End Try

            Dim rangeTxt As String = TryTableRangeString(tbl)
            Dim typeName As String = ""
            Try
                typeName = TypeName(tbl)
            Catch
                typeName = ""
            End Try

            sb.AppendLine(String.Format(CultureInfo.InvariantCulture,
                "  [TABLE] idx={0} name={1} title={2} tableType={3} comType={4} sheet={5} rows={6} cols={7} range={8}",
                ti, tName, tTitle, tType, typeName, sheetName, nRows, nCols, rangeTxt))

            AppendTableCellPeek(sb, tbl, ti, nRows, nCols)
        Next
        Return n
    End Function

    ''' <summary>Inventaria <see cref="DraftDocument.PartsLists"/> (listas de piezas / BOM en el pliego).</summary>
    Private Shared Function AppendDraftDocumentPartsListsSection(sb As StringBuilder, dft As DraftDocument) As Integer
        If sb Is Nothing OrElse dft Is Nothing Then Return 0
        Dim listsObj As Object = Nothing
        Try
            listsObj = CallByName(dft, "PartsLists", CallType.Get)
        Catch ex As Exception
            sb.AppendLine("")
            sb.AppendLine("[PARTSLISTS] DraftDocument.PartsLists no disponible: " & ex.Message)
            Return 0
        End Try
        Dim n As Integer = SafeCount(listsObj)
        sb.AppendLine("")
        sb.AppendLine(String.Format(CultureInfo.InvariantCulture,
            "[PARTSLISTS] DraftDocument.PartsLists Count={0} (listas de piezas; objeto PartsList en Solid Edge Draft)",
            n))
        For pi As Integer = 1 To n
            Dim pl As Object = Nothing
            Try
                pl = CallByName(listsObj, "Item", CallType.Method, pi)
            Catch
                pl = Nothing
            End Try
            If pl Is Nothing Then Continue For

            Dim comType As String = ""
            Try
                comType = TypeName(pl)
            Catch
                comType = ""
            End Try

            Dim listType As String = SafeToString(CallByNameSafe(pl, "ListType"))
            Dim asmFile As String = SafeToString(CallByNameSafe(pl, "AssemblyFileName"))
            Dim upToDate As String = SafeToString(CallByNameSafe(pl, "IsUpToDate"))
            Dim anchor As String = SafeToString(CallByNameSafe(pl, "AnchorPoint"))
            Dim configuration As String = SafeToString(CallByNameSafe(pl, "Configuration"))
            Dim activePl As String = SafeToString(CallByNameSafe(pl, "Active"))

            Dim nRows As Integer = SafeCount(CallByNameSafe(pl, "Rows"))
            Dim nCols As Integer = SafeCount(CallByNameSafe(pl, "Columns"))
            Dim nPages As Integer = SafeCount(CallByNameSafe(pl, "Pages"))
            Dim nGroups As Integer = SafeCount(CallByNameSafe(pl, "Groups"))
            Dim nTitles As Integer = SafeCount(CallByNameSafe(pl, "Titles"))

            Dim sheetName As String = ""
            Try
                Dim par As Object = CallByNameSafe(pl, "Parent")
                sheetName = SafeToString(CallByNameSafe(par, "Name"))
            Catch
                sheetName = ""
            End Try

            Dim modelHint As String = TryModelLinkAuditHint(pl)
            Dim originTxt As String = TryPartsListGetOriginString(pl)
            Dim styleName As String = ""
            Try
                Dim ts As Object = CallByNameSafe(pl, "TableStyle")
                styleName = SafeToString(CallByNameSafe(ts, "Name"))
            Catch
                styleName = ""
            End Try

            sb.AppendLine(String.Format(CultureInfo.InvariantCulture,
                "  [PARTSLIST] idx={0} comType={1} listType={2} assemblyFile={3} isUpToDate={4} active={5} sheetOrParent={6} anchor={7} configuration={8} rows={9} cols={10} pages={11} groups={12} titles={13} tableStyle={14} modelLink={15} origin={16}",
                pi, comType, listType, asmFile, upToDate, activePl, sheetName, anchor, configuration,
                nRows, nCols, nPages, nGroups, nTitles, styleName, modelHint, originTxt))

            AppendPartsListSettingsLine(sb, pl, pi)
            AppendPartsListCellPeek(sb, pl, pi, nRows, nCols)
        Next
        Return n
    End Function

    Private Shared Function TryModelLinkAuditHint(pl As Object) As String
        If pl Is Nothing Then Return ""
        Try
            Dim ml As Object = CallByNameSafe(pl, "ModelLink")
            If ml Is Nothing Then Return ""
            Dim p1 As String = SafeToString(CallByNameSafe(ml, "Name"))
            Dim p2 As String = SafeToString(CallByNameSafe(ml, "FullName"))
            If Not String.IsNullOrWhiteSpace(p2) Then Return TruncateOneLine(p2, 120)
            Return TruncateOneLine(p1, 120)
        Catch
            Return ""
        End Try
    End Function

    Private Shared Function TryPartsListGetOriginString(pl As Object) As String
        If pl Is Nothing Then Return ""
        Try
            Dim x As Double = 0R, y As Double = 0R
            CallByName(pl, "GetOrigin", CallType.Method, x, y)
            Return String.Format(CultureInfo.InvariantCulture, "({0:0.#####},{1:0.#####})", x, y)
        Catch
            Return ""
        End Try
    End Function

    Private Shared Sub AppendPartsListSettingsLine(sb As StringBuilder, pl As Object, idx As Integer)
        If sb Is Nothing OrElse pl Is Nothing Then Return
        Dim keys As String() = {
            "ShowColumnHeader", "ShowTopAssembly", "ShowSheetBackground", "EnableGroups", "CreateNewSheetsForTable",
            "MarkUnballoonedItems", "MarkAmbiguousValues", "MarkAmbiguousValuesForTube", "UseAssemblyItemNumbers", "UseLevelBasedItemNumbers",
            "RenumberAccordingToSortOrder", "ReverseDisplayOrder", "ConvertDeletedPartsIntoUserDefinedRows",
            "FillEndOfTableWithBlankRows", "WrapDataCellsToNewRow", "MaintainSheetsWithTableSize",
            "ExpandWeldmentSubAssemblies", "TotalLengthPartsList", "TotalLengthPartsList_ItemNumberSeparator",
            "ItemNumberStart", "ItemNumberIncrement",
            "MaximumRows", "MaximumRowsFirstPage", "MaximumRowsAdditionalPages", "NumberOfHeaderRows",
            "MaximumHeightFirstPage", "MaximumHeightAdditionalPages", "PageGap", "ColumnHeaderPosition", "GroupByColumn",
            "DataFixedRowHeight", "HeaderFixedRowHeight", "TitleFixedRowHeight",
            "MarginHorizontal", "MarginVertical",
            "MarkAmbiguousValuesString", "MarkAmbiguousValuesStringForTube", "MarkUnballoonedItemsString",
            "FirstSheetBackgroundName", "AdditionalSheetsBackgroundName",
            "FrameRoughCutEndClearance", "PipeRoughCutEndClearance",
            "UseUniquenessCriteria_CutLength", "UseUniquenessCriteria_Mass", "UseUniquenessCriteria_Miter",
            "UseUniquenessCriteria_TubeFlatLength", "UseUniquenessCriteria_TubeMass"
        }
        Dim parts As New List(Of String)()
        For Each k In keys
            Try
                Dim v As Object = CallByName(pl, k, CallType.Get)
                Dim s As String = SafeToString(v)
                If Not String.IsNullOrWhiteSpace(s) Then parts.Add(k & "=" & TruncateOneLine(s, 48))
            Catch
            End Try
        Next
        If parts.Count > 0 Then
            sb.AppendLine("    [PARTSLIST][SETTINGS] idx=" & idx.ToString(CultureInfo.InvariantCulture) & " " & String.Join(" | ", parts))
        End If
    End Sub

    Private Shared Sub AppendPartsListCellPeek(sb As StringBuilder, pl As Object, partsListIdx As Integer, nRows As Integer, nCols As Integer)
        If sb Is Nothing OrElse pl Is Nothing Then Return
        Dim maxRows As Integer = Math.Min(Math.Max(nRows, 0), 4)
        Dim maxCols As Integer = Math.Min(Math.Max(nCols, 0), 8)
        If maxRows <= 0 OrElse maxCols <= 0 Then Return

        For r As Integer = 1 To maxRows
            Dim cells As New List(Of String)()
            For c As Integer = 1 To maxCols
                Dim cellTxt As String = TryGetPartsListCellText(pl, r, c)
                If Not String.IsNullOrEmpty(cellTxt) Then cells.Add(cellTxt)
            Next
            If cells.Count > 0 Then
                sb.AppendLine(String.Format(CultureInfo.InvariantCulture,
                    "    [PARTSLIST][CELLS] idx={0} row={1} sample={2}",
                    partsListIdx, r, String.Join(" | ", cells)))
            End If
        Next
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
        If cellObj IsNot Nothing Then
            Dim s As String = SafeToString(CallByNameSafe(cellObj, "Text"))
            If String.IsNullOrWhiteSpace(s) Then s = SafeToString(CallByNameSafe(cellObj, "Value"))
            If Not String.IsNullOrWhiteSpace(s) Then Return TruncateOneLine(s, 72)
        End If
        Return ""
    End Function

    Private Shared Function TryTableRangeString(tbl As Object) As String
        If tbl Is Nothing Then Return ""
        Try
            Dim x1 As Double = 0, y1 As Double = 0, x2 As Double = 0, y2 As Double = 0
            CallByName(tbl, "Range", CallType.Method, x1, y1, x2, y2)
            Return String.Format(CultureInfo.InvariantCulture, "({0:0.#####},{1:0.#####})-({2:0.#####},{3:0.#####})", x1, y1, x2, y2)
        Catch
            Return ""
        End Try
    End Function

    ''' <summary>Muestra unas pocas celdas si la API expone lectura por fila/columna (varía según versión SE).</summary>
    Private Shared Sub AppendTableCellPeek(sb As StringBuilder, tbl As Object, tableIdx As Integer, nRows As Integer, nCols As Integer)
        If sb Is Nothing OrElse tbl Is Nothing Then Return
        Dim maxRows As Integer = Math.Min(Math.Max(nRows, 0), 4)
        Dim maxCols As Integer = Math.Min(Math.Max(nCols, 0), 6)
        If maxRows <= 0 OrElse maxCols <= 0 Then Return

        For r As Integer = 1 To maxRows
            Dim cells As New List(Of String)()
            For c As Integer = 1 To maxCols
                Dim cellTxt As String = TryGetTableCellText(tbl, r, c)
                If Not String.IsNullOrEmpty(cellTxt) Then cells.Add(cellTxt)
            Next
            If cells.Count > 0 Then
                sb.AppendLine(String.Format(CultureInfo.InvariantCulture,
                    "    [TABLE][CELLS] tableIdx={0} row={1} sample={2}",
                    tableIdx, r, String.Join(" | ", cells)))
            End If
        Next
    End Sub

    Private Shared Function TryGetTableCellText(tbl As Object, row1 As Integer, col1 As Integer) As String
        If tbl Is Nothing Then Return ""
        Dim methods As String() = {
            "GetCellString", "CellString", "GetText", "GetValueAsString"
        }
        For Each m In methods
            Try
                Dim v As Object = CallByName(tbl, m, CallType.Method, row1, col1)
                Dim s As String = SafeToString(v)
                If Not String.IsNullOrWhiteSpace(s) Then Return TruncateOneLine(s, 80)
            Catch
            End Try
        Next
        Try
            Dim cellObj As Object = CallByName(tbl, "Cell", CallType.Method, row1, col1)
            If cellObj IsNot Nothing Then
                Dim s2 As String = SafeToString(CallByNameSafe(cellObj, "Text"))
                If String.IsNullOrWhiteSpace(s2) Then s2 = SafeToString(CallByNameSafe(cellObj, "Value"))
                If Not String.IsNullOrWhiteSpace(s2) Then Return TruncateOneLine(s2, 80)
            End If
        Catch
        End Try
        Return ""
    End Function

    Private Shared Function TruncateOneLine(s As String, maxLen As Integer) As String
        If String.IsNullOrEmpty(s) Then Return ""
        s = s.Replace(vbCr, " ").Replace(vbLf, " ").Trim()
        If s.Length <= maxLen Then Return s
        Return s.Substring(0, maxLen) & "…"
    End Function

    Private Shared Sub AppendDimensionDetails(sb As StringBuilder, dimsObj As Object)
        If sb Is Nothing OrElse dimsObj Is Nothing Then Return
        Dim n As Integer = SafeCount(dimsObj)
        If n <= 0 Then Return
        Dim typeCounters As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
        For d As Integer = 1 To n
            Try
                Dim dimObj As Object = CallByName(dimsObj, "Item", CallType.Method, d)
                If dimObj Is Nothing Then Continue For
                Dim dimType As String = SafeToString(CallByNameSafe(dimObj, "DimensionType"))
                If String.IsNullOrWhiteSpace(dimType) Then dimType = "(unknown)"
                If Not typeCounters.ContainsKey(dimType) Then typeCounters(dimType) = 0
                typeCounters(dimType) += 1
                Dim valueTxt As String = SafeToString(CallByNameSafe(dimObj, "Value"))
                Dim styleName As String = ""
                Try
                    Dim st As Object = CallByNameSafe(dimObj, "Style")
                    styleName = SafeToString(CallByNameSafe(st, "Name"))
                Catch
                    styleName = ""
                End Try
                Dim parentName As String = ""
                Try
                    parentName = SafeToString(CallByNameSafe(CallByNameSafe(dimObj, "Parent"), "Name"))
                Catch
                    parentName = ""
                End Try
                sb.AppendLine(String.Format(CultureInfo.InvariantCulture,
                    "  [DIM] idx={0} type={1} value={2} style={3} parent={4}",
                    d, dimType, valueTxt, styleName, parentName))
            Catch
            End Try
        Next
        If typeCounters.Count > 0 Then
            Dim pairs As New List(Of String)()
            For Each kv In typeCounters.OrderBy(Function(x) x.Key, StringComparer.OrdinalIgnoreCase)
                pairs.Add(kv.Key & "=" & kv.Value.ToString(CultureInfo.InvariantCulture))
            Next
            sb.AppendLine("  [DIM][SUMMARY] " & String.Join(", ", pairs))
        End If
    End Sub

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

    Private Shared Sub TryReleaseComObject(obj As Object)
        If obj Is Nothing Then Return
        Try
            If Marshal.IsComObject(obj) Then Marshal.ReleaseComObject(obj)
        Catch
        End Try
    End Sub
End Class
