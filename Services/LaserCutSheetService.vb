Option Strict Off

Imports System.Globalization
Imports System.IO
Imports System.Linq
Imports IOPath = System.IO.Path
Imports System.Runtime.InteropServices
Imports SolidEdgeAssembly
Imports SolidEdgeDraft
Imports SolidEdgeFramework
Imports SolidEdgeFrameworkSupport
Imports SolidEdgePart

''' <summary>Plano de corte láser 1:1 agrupado por espesor (módulo aislado del motor DFT individual).</summary>
Public Class LaserCutSheetService
    Public Const LaserStartX As Double = 0.02
    Public Const LaserStartY As Double = 0.02
    Public Const LaserPieceGapX As Double = 0.05
    Public Const LaserRowGapY As Double = 0.2
    ''' <summary>Espacio reservado a la izquierda para la etiqueta de espesor en cada fila (m).</summary>
    Public Const LaserThicknessLabelWidth As Double = 0.12

    Private Const ScaleOne As Double = 1.0
    Private Const LaserDftFileName As String = "Piezas a pedir de Corte"
    Private Const LaserOutputSubfolder As String = "Corte Laser"
    Private Const BendListFileName As String = "Piezas a plegar.txt"

    Private ReadOnly _logger As Logger
    Private _laserFileLog As StreamWriter = Nothing

    Public Sub New(logger As Logger)
        _logger = logger
    End Sub

    Public Sub LogLaser(msg As String)
        If _logger IsNot Nothing Then _logger.Log(msg)
        Try
            If _laserFileLog IsNot Nothing Then
                _laserFileLog.WriteLine(DateTime.Now.ToString("HH:mm:ss.fff", CultureInfo.InvariantCulture) & " " & msg)
                _laserFileLog.Flush()
            End If
        Catch
        End Try
    End Sub

    ''' <summary>Usa <see cref="AssemblyComponentService"/> (mismo recorrido que la columna de componentes). Sin lectura COM por pieza.</summary>
    Public Function ScanAssembly(asmPath As String,
                               outputFolder As String,
                               templateDxfPath As String,
                               showSolidEdge As Boolean) As List(Of LaserCutPieceInfo)
        LogLaser("[LASER][ENTER]")
        LogLaser("[LASER][ASM][SCAN_START] path=" & asmPath)
        If String.IsNullOrWhiteSpace(asmPath) OrElse Not File.Exists(asmPath) Then Return New List(Of LaserCutPieceInfo)()

        Dim items = AssemblyComponentService.LoadAssemblyComponentItems(
            asmPath,
            uniqueOnly:=True,
            showSolidEdge,
            _logger,
            Sub(phase As String, current As Integer, total As Integer)
                LogLaser("[LASER][ASM][PROGRESS] " & phase & " n=" & current.ToString(CultureInfo.InvariantCulture))
            End Sub)
        LogLaser("[LASER][ASM][SCAN] componentes=" & items.Count.ToString(CultureInfo.InvariantCulture))
        Return BuildPieceListFromComponentItems(items, outputFolder, flatByPath:=Nothing)
    End Function

    ''' <summary>Tabla láser desde lista ASM ya cargada (no abre Solid Edge).</summary>
    Public Function BuildPieceListFromComponentItems(items As IEnumerable(Of AssemblyComponentItem),
                                                       outputFolder As String,
                                                       flatByPath As Dictionary(Of String, Boolean?)) As List(Of LaserCutPieceInfo)
        Dim pieces As New Dictionary(Of String, LaserCutPieceInfo)(StringComparer.OrdinalIgnoreCase)
        If items Is Nothing Then Return New List(Of LaserCutPieceInfo)()

        For Each it In items
            If it Is Nothing OrElse String.IsNullOrWhiteSpace(it.FullPath) Then Continue For
            If String.Equals(it.Kind, "ASM", StringComparison.OrdinalIgnoreCase) Then Continue For
            If Not String.Equals(it.Kind, "PAR", StringComparison.OrdinalIgnoreCase) AndAlso
               Not String.Equals(it.Kind, "PSM", StringComparison.OrdinalIgnoreCase) Then Continue For

            Dim fullPath As String = it.FullPath.Trim()
            If LaserCutPartFilters.ShouldExcludeFile(fullPath) Then
                LogLaser("[LASER][FILTER][EXCLUDE] file=" & IOPath.GetFileName(fullPath))
                Continue For
            End If
            LogLaser("[LASER][FILTER][INCLUDE] file=" & IOPath.GetFileName(fullPath))

            If pieces.ContainsKey(fullPath) Then
                pieces(fullPath).Quantity += Math.Max(1, it.OccurrenceCount)
                LogLaser("[LASER][QTY][MERGE] file=" & IOPath.GetFileName(fullPath) &
                         " qty=" & pieces(fullPath).Quantity.ToString(CultureInfo.InvariantCulture))
                Continue For
            End If

            Dim isPsm As Boolean = String.Equals(it.Kind, "PSM", StringComparison.OrdinalIgnoreCase)
            Dim qty As Integer = Math.Max(1, it.OccurrenceCount)
            Dim info As New LaserCutPieceInfo With {
                .Include = True,
                .FilePath = fullPath,
                .FileName = IOPath.GetFileName(fullPath),
                .FileNameNoExt = IOPath.GetFileNameWithoutExtension(fullPath),
                .FileType = If(isPsm, "PSM", "PAR"),
                .Quantity = qty,
                .IsSheetMetal = isPsm,
                .Status = "Falta espesor",
                .Notes = "Indique espesor en la tabla (mm)"
            }

            If isPsm Then
                info.IsBent = Nothing
                If flatByPath IsNot Nothing AndAlso flatByPath.ContainsKey(fullPath) Then
                    info.HasFlatPattern = flatByPath(fullPath)
                    If info.HasFlatPattern.HasValue AndAlso Not info.HasFlatPattern.Value Then
                        info.Status = "SIN FLAT"
                        info.Notes = "Sin desarrollo flat"
                    End If
                    LogLaser("[LASER][PSM][FLAT] file=" & info.FileName & " hasFlat=" & info.HasFlatPattern.ToString())
                Else
                    info.HasFlatPattern = Nothing
                End If
            Else
                info.HasFlatPattern = Nothing
                info.IsBent = Nothing
            End If

            If Not String.IsNullOrWhiteSpace(outputFolder) Then
                Dim cfg As New JobConfiguration With {.OutputFolder = outputFolder}
                info.SourceDftPath = DraftStandaloneMotorPaths.GetExpectedStandaloneDraftPath(cfg, fullPath)
            End If

            pieces(fullPath) = info
            LogLaser("[LASER][QTY] file=" & info.FileName & " occurrences=" & qty.ToString(CultureInfo.InvariantCulture))
        Next

        Dim list = pieces.Values.OrderBy(Function(p) p.ThicknessMm.GetValueOrDefault(Double.MaxValue)).
            ThenBy(Function(p) p.FileNameNoExt, StringComparer.OrdinalIgnoreCase).ToList()
        WriteSummaryLogs(list)
        LogLaser("[LASER][UI][TABLE_LOAD] count=" & list.Count.ToString(CultureInfo.InvariantCulture))
        Return list
    End Function

    ''' <summary>Espesor desde DFT generado (ruta esperada) o propiedades del PAR/PSM.</summary>
    Public Sub FillThicknessesFromDraftsAndModels(pieces As IList(Of LaserCutPieceInfo), showSolidEdge As Boolean)
        If pieces Is Nothing OrElse pieces.Count = 0 Then Return
        For Each p In pieces
            If p Is Nothing Then Continue For
            If p.ThicknessMm.HasValue AndAlso p.ThicknessMm.Value > 0 Then Continue For

            Try
                Dim md As DrawingMetadataInput = Nothing
                If DrawingMetadataService.TryLoadMetadataFromModelFile(p.FilePath, showSolidEdge, _logger, md) AndAlso md IsNot Nothing Then
                    ApplyThicknessFromMetadata(p, md.Espesor, If(md.EspesorSource, "modelo"))
                End If
            Catch ex As Exception
                If _logger IsNot Nothing Then _logger.LogException("FillThickness model " & p.FileName, ex)
            End Try

            If p.ThicknessMm.HasValue AndAlso p.ThicknessMm.Value > 0 Then Continue For
            If String.IsNullOrWhiteSpace(p.SourceDftPath) OrElse Not File.Exists(p.SourceDftPath) Then Continue For
            Try
                Dim mdDft As DrawingMetadataInput = Nothing
                If DrawingMetadataService.TryLoadMetadataFromDraftPath(p.SourceDftPath, showSolidEdge, _logger, mdDft) AndAlso mdDft IsNot Nothing Then
                    ApplyThicknessFromMetadata(p, mdDft.Espesor, If(mdDft.EspesorSource, "DFT"))
                End If
            Catch ex As Exception
                If _logger IsNot Nothing Then _logger.LogException("FillThickness dft " & p.FileName, ex)
            End Try
        Next
    End Sub

    Private Sub ApplyThicknessFromMetadata(piece As LaserCutPieceInfo, thicknessText As String, source As String)
        If piece Is Nothing OrElse String.IsNullOrWhiteSpace(thicknessText) Then Return
        Dim parsed = TryParseThicknessMm(thicknessText)
        If parsed.HasValue AndAlso parsed.Value > 0 Then
            piece.ThicknessMm = parsed
            piece.ThicknessText = parsed.Value.ToString("0.###", CultureInfo.InvariantCulture)
            If String.Equals(piece.Status, "Falta espesor", StringComparison.OrdinalIgnoreCase) Then piece.Status = "Listo"
            LogLaser("[LASER][THICKNESS][" & source & "] file=" & piece.FileName &
                     " mm=" & parsed.Value.ToString("0.###", CultureInfo.InvariantCulture))
        ElseIf String.IsNullOrWhiteSpace(piece.ThicknessText) Then
            piece.ThicknessText = thicknessText.Trim()
            LogLaser("[LASER][THICKNESS][" & source & "][RAW] file=" & piece.FileName & " text=" & thicknessText.Trim())
        End If
    End Sub

    Public Function GenerateLaserCutSheet(app As Application,
                                          pieces As IList(Of LaserCutPieceInfo),
                                          outputFolder As String,
                                          templateDxfPath As String,
                                          showSolidEdge As Boolean) As LaserCutGenerateResult
        Dim result As New LaserCutGenerateResult()
        LogLaser("[LASER][DFT][CREATE]")
        If app Is Nothing Then
            result.ErrorMessage = "Solid Edge no conectado."
            Return result
        End If
        If String.IsNullOrWhiteSpace(templateDxfPath) OrElse Not File.Exists(templateDxfPath) Then
            result.ErrorMessage = "Template DXF_LIMPIO no encontrado: " & templateDxfPath
            Return result
        End If

        Dim included = pieces.Where(Function(p) p IsNot Nothing AndAlso p.Include).ToList()
        If included.Count = 0 Then
            result.ErrorMessage = "No hay piezas incluidas."
            Return result
        End If

        Dim outDir As String = IOPath.Combine(outputFolder.Trim(), LaserOutputSubfolder)
        Directory.CreateDirectory(outDir)
        Dim laserLogPath As String = IOPath.Combine(outDir, "laser_cut_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".txt")
        Try
            _laserFileLog = New StreamWriter(laserLogPath, append:=False, encoding:=System.Text.Encoding.UTF8)
            _laserFileLog.WriteLine("=== laser_cut " & DateTime.Now.ToString("o", CultureInfo.InvariantCulture) & " ===")
        Catch
            _laserFileLog = Nothing
        End Try
        result.LaserLogPath = laserLogPath

        Dim dftPath As String = IOPath.Combine(outDir, LaserDftFileName & ".dft")
        Dim dxfPath As String = IOPath.Combine(outDir, LaserDftFileName & ".dxf")
        Dim draftDoc As DraftDocument = Nothing

        Try
            app.Visible = showSolidEdge
            app.DisplayAlerts = False
            draftDoc = CType(app.Documents.Add("SolidEdge.DraftDocument", templateDxfPath), DraftDocument)
            DoIdle(app)
            Dim sheet As Sheet = draftDoc.ActiveSheet
            Dim viewsBefore As Integer = 0
            Try : viewsBefore = sheet.DrawingViews.Count : Catch : End Try
            LogLaser("[LASER][DFT][VIEWS_BEFORE] count=" & viewsBefore.ToString(CultureInfo.InvariantCulture))

            Dim groups = included.
                GroupBy(Function(p) p.ThicknessMm.GetValueOrDefault(0)).
                OrderBy(Function(g) g.Key).
                ToList()

            Dim rowTopY As Double = LaserStartY
            LogLaser("[LASER][COORD] row advance: nextRowTop = currentRowTop - maxRowHeight - LaserRowGapY (Y decrece hacia abajo en hoja)")

            For Each grp In groups
                Dim thicknessKey As Double = grp.Key
                LogLaser("[LASER][ROW][START] espesor=" & thicknessKey.ToString("0.###", CultureInfo.InvariantCulture))
                Dim rowLeftX As Double = LaserStartX + LaserThicknessLabelWidth
                Dim rowMaxHeight As Double = 0

                InsertSheetText(sheet, LaserStartX, rowTopY, "Espesor " & thicknessKey.ToString("0.###", CultureInfo.InvariantCulture) & " mm")

                For Each piece In grp.OrderBy(Function(p) p.FileNameNoExt, StringComparer.OrdinalIgnoreCase)
                    Dim useFlat As Boolean = True
                    If piece.IsSheetMetal AndAlso piece.HasFlatPattern.HasValue AndAlso Not piece.HasFlatPattern.Value Then
                        useFlat = False
                        LogLaser("[LASER][FLAT][MISSING] file=" & piece.FileName & " -> vista principal 1:1")
                        piece.Notes = "Sin flat: vista principal 1:1"
                    End If

                    Dim dv As DrawingView = Nothing
                    Try
                        LogLaser("[LASER][VIEW][TRY] name=" & piece.FileNameNoExt & " path=" & piece.FilePath &
                                 " mode=" & If(useFlat, "flat", "principal"))
                        dv = CojonudoBestFit_Bueno.InsertLaserCutPieceView(
                            app, draftDoc, sheet, piece.FilePath, templateDxfPath,
                            0.15, 0.2, ScaleOne, useFlat, AddressOf LogLaser)
                        If dv Is Nothing Then
                            LogLaser("[LASER][VIEW][FAIL] name=" & piece.FileNameNoExt)
                            piece.Status = "Error vista"
                            piece.Notes = "No se pudo insertar vista 1:1"
                            Continue For
                        End If

                        MoveViewTopLeft(app, dv, rowLeftX, rowTopY)
                        SafeUpdateView(dv)
                        DoIdle(app)

                        Dim xmin As Double, ymin As Double, xmax As Double, ymax As Double
                        If Not CojonudoBestFit_Bueno.TryGetViewRangePublic(dv, xmin, ymin, xmax, ymax) Then
                            LogLaser("[LASER][VIEW][FAIL] no range after move name=" & piece.FileNameNoExt)
                            piece.Status = "Sin range"
                            Continue For
                        End If
                        Dim w As Double = Math.Abs(xmax - xmin)
                        Dim h As Double = Math.Abs(ymax - ymin)
                        rowMaxHeight = Math.Max(rowMaxHeight, h)
                        LogLaser("[LASER][VIEW][RANGE] name=" & piece.FileNameNoExt &
                                 " x1=" & xmin.ToString("0.0000", CultureInfo.InvariantCulture) &
                                 " y1=" & ymin.ToString("0.0000", CultureInfo.InvariantCulture) &
                                 " x2=" & xmax.ToString("0.0000", CultureInfo.InvariantCulture) &
                                 " y2=" & ymax.ToString("0.0000", CultureInfo.InvariantCulture))
                        LogLaser("[LASER][ROW][PIECE] name=" & piece.FileNameNoExt &
                                 " x=" & rowLeftX.ToString("0.0000", CultureInfo.InvariantCulture) &
                                 " y=" & rowTopY.ToString("0.0000", CultureInfo.InvariantCulture) &
                                 " rangeW=" & w.ToString("0.0000", CultureInfo.InvariantCulture) &
                                 " rangeH=" & h.ToString("0.0000", CultureInfo.InvariantCulture))

                        Dim labelY As Double = Math.Min(ymin, ymax) - 0.012
                        Dim centerX As Double = (Math.Min(xmin, xmax) + Math.Max(xmin, xmax)) / 2.0
                        Dim labelText As String = FormatLaserPieceLabel(piece)
                        InsertSheetText(sheet, centerX, labelY, labelText)
                        LogLaser("[LASER][TEXT][NAME] text=" & labelText)

                        If piece.IsSheetMetal AndAlso piece.IsBent.GetValueOrDefault(False) Then
                            InsertSheetText(sheet, centerX, labelY - 0.01, "LLEVA PLIEGUE")
                            LogLaser("[LASER][TEXT][BEND] name=" & piece.FileNameNoExt)
                        End If

                        rowLeftX = Math.Max(xmin, xmax) + LaserPieceGapX
                        piece.Status = "Colocada"
                        LogLaser("[LASER][VIEW][ADD] name=" & piece.FileNameNoExt)
                    Catch ex As Exception
                        piece.Status = "Error"
                        piece.Notes = ex.Message
                        If _logger IsNot Nothing Then _logger.LogException("LaserCut place piece", ex)
                    End Try
                Next

                ' Siguiente fila: altura máxima de la fila anterior + 20 cm (LaserRowGapY = 0,20 m).
                Dim nextY As Double = rowTopY - rowMaxHeight - LaserRowGapY
                LogLaser("[LASER][ROW][END] espesor=" & thicknessKey.ToString("0.###", CultureInfo.InvariantCulture) &
                         " maxHeight=" & rowMaxHeight.ToString("0.0000", CultureInfo.InvariantCulture) &
                         " nextY=" & nextY.ToString("0.0000", CultureInfo.InvariantCulture))
                rowTopY = nextY
            Next

            WriteSummaryLogs(included)

            Dim viewsAfter As Integer = 0
            Try : viewsAfter = sheet.DrawingViews.Count : Catch : End Try
            LogLaser("[LASER][DFT][VIEWS_AFTER] count=" & viewsAfter.ToString(CultureInfo.InvariantCulture))

            draftDoc.SaveAs(dftPath)
            LogLaser("[LASER][DFT][SAVE] path=" & dftPath)
            result.DftPath = dftPath

            LogLaser("[LASER][DXF][EXPORT]")
            ExportDraftToDxf(draftDoc, dxfPath)
            LogLaser("[LASER][DXF][OK] path=" & dxfPath)
            result.DxfPath = dxfPath

            result.BendListPath = SaveBendList(outDir, included)
            CopyFlatPieceDraftsToLaserFolder(outDir, outputFolder.Trim(), included)
            result.Success = True
        Catch ex As Exception
            result.ErrorMessage = ex.Message
            If _logger IsNot Nothing Then _logger.LogException("GenerateLaserCutSheet", ex)
        Finally
            Try
                If draftDoc IsNot Nothing Then draftDoc.Close(True)
            Catch
            End Try
            Try
                If _laserFileLog IsNot Nothing Then
                    _laserFileLog.Dispose()
                    _laserFileLog = Nothing
                End If
            Catch
            End Try
            LogLaser("[LASER][EXIT]")
        End Try

        Return result
    End Function

    ''' <summary>Copia a <paramref name="laserOutDir"/> los DFT individuales de piezas PSM con desarrollo flat.</summary>
    Private Sub CopyFlatPieceDraftsToLaserFolder(laserOutDir As String,
                                                  outputFolder As String,
                                                  pieces As IList(Of LaserCutPieceInfo))
        If String.IsNullOrWhiteSpace(laserOutDir) OrElse pieces Is Nothing Then Return
        Directory.CreateDirectory(laserOutDir)
        Dim copied As Integer = 0
        Dim skipped As Integer = 0

        For Each p In pieces
            If p Is Nothing OrElse Not p.Include OrElse Not p.IsSheetMetal Then Continue For
            If p.HasFlatPattern.HasValue AndAlso Not p.HasFlatPattern.Value Then
                LogLaser("[LASER][COPY][SKIP] sin flat file=" & p.FileName)
                skipped += 1
                Continue For
            End If

            Dim src As String = ResolveExistingPieceDraftPath(p, outputFolder)
            If String.IsNullOrWhiteSpace(src) Then
                LogLaser("[LASER][COPY][SKIP] DFT no encontrado file=" & p.FileName)
                skipped += 1
                Continue For
            End If

            Dim dest As String = IOPath.Combine(laserOutDir, IOPath.GetFileName(src))
            Try
                If String.Equals(IOPath.GetFullPath(src), IOPath.GetFullPath(dest), StringComparison.OrdinalIgnoreCase) Then
                    LogLaser("[LASER][COPY][SKIP] mismo archivo file=" & p.FileName)
                    Continue For
                End If
                File.Copy(src, dest, overwrite:=True)
                copied += 1
                LogLaser("[LASER][COPY][DFT] src=" & src & " -> " & dest)
            Catch ex As Exception
                skipped += 1
                LogLaser("[LASER][COPY][FAIL] file=" & p.FileName & " " & ex.Message)
                If _logger IsNot Nothing Then _logger.LogException("CopyFlatPieceDrafts", ex)
            End Try
        Next

        LogLaser("[LASER][COPY][SUMMARY] copied=" & copied.ToString(CultureInfo.InvariantCulture) &
                 " skipped=" & skipped.ToString(CultureInfo.InvariantCulture) &
                 " dir=" & laserOutDir)
    End Sub

    ''' <summary>Localiza el DFT generado bajo OutputFolder\DFT (incluye sufijos _001 y DIMLAB).</summary>
    Private Shared Function ResolveExistingPieceDraftPath(piece As LaserCutPieceInfo, outputFolder As String) As String
        If piece Is Nothing Then Return ""
        If Not String.IsNullOrWhiteSpace(piece.SourceDftPath) AndAlso File.Exists(piece.SourceDftPath) Then
            Return IOPath.GetFullPath(piece.SourceDftPath)
        End If
        If String.IsNullOrWhiteSpace(outputFolder) OrElse String.IsNullOrWhiteSpace(piece.FilePath) Then Return ""

        Dim dftDir As String = IOPath.Combine(outputFolder.Trim().TrimEnd(
            IOPath.DirectorySeparatorChar, IOPath.AltDirectorySeparatorChar), "DFT")
        If Not Directory.Exists(dftDir) Then Return ""

        Dim baseStem As String = IOPath.GetFileNameWithoutExtension(piece.FilePath)
        Dim directNames As String() = {
            baseStem & ".dft",
            baseStem & "_DIMLAB_REF_TEST.dft"
        }
        For Each name In directNames
            Dim candidate As String = IOPath.Combine(dftDir, name)
            If File.Exists(candidate) Then Return IOPath.GetFullPath(candidate)
        Next

        Dim numbered = Directory.GetFiles(dftDir, baseStem & "_*.dft", SearchOption.TopDirectoryOnly).
            OrderByDescending(Function(f) File.GetLastWriteTimeUtc(f)).
            ToList()
        If numbered.Count > 0 Then Return numbered(0)

        Dim anyMatch = Directory.GetFiles(dftDir, baseStem & "*.dft", SearchOption.TopDirectoryOnly).
            OrderByDescending(Function(f) File.GetLastWriteTimeUtc(f)).
            ToList()
        If anyMatch.Count > 0 Then Return anyMatch(0)

        Return ""
    End Function

    Private Shared Function FormatLaserPieceLabel(piece As LaserCutPieceInfo) As String
        If piece Is Nothing Then Return ""
        Dim qty As Integer = Math.Max(1, piece.Quantity)
        Return piece.FileNameNoExt & "  (" & qty.ToString(CultureInfo.InvariantCulture) & " ud.)"
    End Function

    Private Sub InsertSheetText(sheet As Sheet, x As Double, y As Double, text As String)
        If sheet Is Nothing OrElse String.IsNullOrWhiteSpace(text) Then Return
        Try
            Dim boxes As TextBoxes = sheet.TextBoxes
            Dim tb As TextBox = boxes.Add(x, y, 0)
            tb.TextScale = 1
            tb.Text = text
            LogLaser("[LASER][TEXT][ADD] """ & text & """ @ " & x.ToString("0.0000", CultureInfo.InvariantCulture) &
                     "," & y.ToString("0.0000", CultureInfo.InvariantCulture))
        Catch ex As Exception
            LogLaser("[LASER][TEXT][ERR] " & ex.Message)
            If _logger IsNot Nothing Then _logger.LogException("InsertSheetText", ex)
        End Try
    End Sub

    Private Function SaveBendList(outDir As String, pieces As IList(Of LaserCutPieceInfo)) As String
        Dim path As String = IOPath.Combine(outDir, BendListFileName)
        Using sw As New StreamWriter(path, False, System.Text.Encoding.UTF8)
            sw.WriteLine("Piezas a plegar")
            sw.WriteLine("Fecha: " & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
            sw.WriteLine("")
            For Each p In pieces
                If Not p.Include OrElse Not p.IsSheetMetal Then Continue For
                If Not p.IsBent.GetValueOrDefault(False) Then Continue For
                LogLaser("[LASER][BEND_LIST][ADD] " & p.FileName)
                sw.WriteLine("Pieza: " & p.FileNameNoExt)
                sw.WriteLine("  Ruta PSM: " & p.FilePath)
                sw.WriteLine("  Espesor: " & If(p.ThicknessMm.HasValue, p.ThicknessMm.Value.ToString("0.###", CultureInfo.InvariantCulture) & " mm", p.ThicknessText))
                sw.WriteLine("  Cantidad: " & p.Quantity.ToString(CultureInfo.InvariantCulture))
                sw.WriteLine("  DFT individual: " & If(p.SourceDftPath, ""))
                sw.WriteLine("")
            Next
        End Using
        LogLaser("[LASER][BEND_LIST][SAVE] path=" & path)
        Return path
    End Function

    Private Sub ExportDraftToDxf(draftDoc As DraftDocument, dxfPath As String)
        Try
            draftDoc.SaveCopyAs(dxfPath)
        Catch
            draftDoc.SaveAs(dxfPath)
        End Try
    End Sub

    Private Sub WriteSummaryLogs(list As List(Of LaserCutPieceInfo))
        Dim par = list.Where(Function(p) String.Equals(p.FileType, "PAR", StringComparison.OrdinalIgnoreCase)).Count()
        Dim psm = list.Where(Function(p) String.Equals(p.FileType, "PSM", StringComparison.OrdinalIgnoreCase)).Count()
        Dim inc = list.Where(Function(p) p.Include).Count()
        Dim qty = list.Where(Function(p) p.Include).Sum(Function(p) p.Quantity)
        Dim bend = list.Where(Function(p) p.IsSheetMetal AndAlso p.IsBent.GetValueOrDefault(False)).Count()
        LogLaser("[LASER][SUMMARY] par=" & par.ToString(CultureInfo.InvariantCulture) &
                 " psm=" & psm.ToString(CultureInfo.InvariantCulture) &
                 " included=" & inc.ToString(CultureInfo.InvariantCulture) &
                 " totalQty=" & qty.ToString(CultureInfo.InvariantCulture))
        LogLaser("[LASER][SUMMARY][BEND] count=" & bend.ToString(CultureInfo.InvariantCulture))
        For Each g In list.Where(Function(p) p.Include).GroupBy(Function(p) p.ThicknessMm.GetValueOrDefault(0))
            Dim c = g.Count()
            Dim q = g.Sum(Function(p) p.Quantity)
            LogLaser("[LASER][SUMMARY][THICKNESS] espesor=" & g.Key.ToString("0.###", CultureInfo.InvariantCulture) &
                     " count=" & c.ToString(CultureInfo.InvariantCulture) &
                     " qty=" & q.ToString(CultureInfo.InvariantCulture))
        Next
    End Sub

    Public Shared Function TryParseThicknessMm(text As String) As Double?
        If String.IsNullOrWhiteSpace(text) Then Return Nothing
        Dim s As String = text.Trim().Replace(",", ".")
        Dim m = System.Text.RegularExpressions.Regex.Match(s, "(\d+([.]\d+)?)")
        If Not m.Success Then Return Nothing
        Dim v As Double
        If Double.TryParse(m.Groups(1).Value, NumberStyles.Any, CultureInfo.InvariantCulture, v) AndAlso v > 0 Then
            Return v
        End If
        Return Nothing
    End Function

    Private Shared Function ConnectApplication(showSolidEdge As Boolean, ByRef createdByUs As Boolean) As Application
        createdByUs = False
        Try
            Return CType(Marshal.GetActiveObject("SolidEdge.Application"), Application)
        Catch
            Dim t = Type.GetTypeFromProgID("SolidEdge.Application")
            If t Is Nothing Then Return Nothing
            createdByUs = True
            Return CType(Activator.CreateInstance(t), Application)
        End Try
    End Function

    Private Shared Sub DoIdle(app As Application)
        Try : app.DoIdle() : Catch : End Try
    End Sub

    Private Shared Sub SafeUpdateView(dv As DrawingView)
        If dv Is Nothing Then Return
        Try : dv.Update() : Catch : End Try
    End Sub

    Private Shared Function TryGetViewRange(dv As DrawingView, ByRef xmin As Double, ByRef ymin As Double, ByRef xmax As Double, ByRef ymax As Double) As Boolean
        xmin = 0 : ymin = 0 : xmax = 0 : ymax = 0
        Try
            dv.Range(xmin, ymin, xmax, ymax)
            Return True
        Catch
            Return False
        End Try
    End Function

    Private Shared Sub MoveViewTopLeft(app As Application, dv As DrawingView, targetLeft As Double, targetTop As Double)
        If dv Is Nothing Then Return
        Dim xmin As Double, ymin As Double, xmax As Double, ymax As Double
        If Not TryGetViewRange(dv, xmin, ymin, xmax, ymax) Then Return
        Dim left As Double = Math.Min(xmin, xmax)
        Dim top As Double = Math.Max(ymin, ymax)
        Dim ox As Double = 0, oy As Double = 0
        dv.GetOrigin(ox, oy)
        dv.SetOrigin(ox + (targetLeft - left), oy + (targetTop - top))
        DoIdle(app)
    End Sub
End Class
