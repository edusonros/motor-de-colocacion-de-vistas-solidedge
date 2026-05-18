Option Strict Off

Imports System.Globalization
Imports System.Runtime.InteropServices
Imports System.IO
Imports System.Linq
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports SolidEdgeFramework

Partial Public Class MainForm

    Private _laserPieces As New List(Of LaserCutPieceInfo)()
    Private _laserService As LaserCutSheetService = Nothing

    Private Sub EnsureLaserService()
        If _laserService Is Nothing Then _laserService = New LaserCutSheetService(_logger)
    End Sub

    ''' <summary>Replica en la tabla láser el estado «Incluir» de componentes ASM.</summary>
    Private Sub ApplyAsmSelectionToLaserPieces()
        If _laserPieces Is Nothing OrElse _laserPieces.Count = 0 Then Return
        For Each p In _laserPieces
            p.Include = IsAsmPathSelectedForLaser(p.FilePath)
        Next
    End Sub

    Private Sub SyncLaserIncludeFromAsmGrid()
        If _laserPieces Is Nothing OrElse _laserPieces.Count = 0 Then Return
        ApplyAsmSelectionToLaserPieces()
        BindLaserGrid()
    End Sub

    Private Async Function FinalizeLaserPieceListAfterScanAsync(outFolder As String, showSe As Boolean) As Task
        ApplyAsmSelectionToLaserPieces()
        UpdateStatus("Corte láser — leyendo espesores...")
        Await StaComInvoker.Run(Function()
                                    _laserService.FillThicknessesFromDraftsAndModels(_laserPieces, showSe)
                                    Return True
                                End Function).ConfigureAwait(True)
        BindLaserGrid()
    End Function

    Private Sub SetupLaserDataGridView()
        If dgvLaserPieces Is Nothing Then Return
        dgvLaserPieces.SuspendLayout()
        dgvLaserPieces.Columns.Clear()
        dgvLaserPieces.AutoGenerateColumns = False
        dgvLaserPieces.AllowUserToAddRows = False
        dgvLaserPieces.RowHeadersVisible = False
        dgvLaserPieces.SelectionMode = DataGridViewSelectionMode.FullRowSelect

        dgvLaserPieces.Columns.Add(New DataGridViewCheckBoxColumn With {
            .Name = "colLaserInclude", .HeaderText = "Incluir", .Width = 52})
        dgvLaserPieces.Columns.Add(New DataGridViewTextBoxColumn With {
            .Name = "colLaserName", .HeaderText = "Nombre pieza", .ReadOnly = True, .MinimumWidth = 120, .FillWeight = 80})
        dgvLaserPieces.Columns.Add(New DataGridViewTextBoxColumn With {
            .Name = "colLaserPath", .HeaderText = "Ruta fichero", .ReadOnly = True, .MinimumWidth = 140, .FillWeight = 100})
        dgvLaserPieces.Columns.Add(New DataGridViewTextBoxColumn With {
            .Name = "colLaserType", .HeaderText = "Tipo", .ReadOnly = True, .Width = 44})
        dgvLaserPieces.Columns.Add(New DataGridViewTextBoxColumn With {
            .Name = "colLaserThickness", .HeaderText = "Espesor", .Width = 64})
        dgvLaserPieces.Columns.Add(New DataGridViewComboBoxColumn With {
            .Name = "colLaserBent", .HeaderText = "Va plegada", .Width = 88,
            .DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton})
        CType(dgvLaserPieces.Columns("colLaserBent"), DataGridViewComboBoxColumn).Items.AddRange({"Sí", "No", "?"})
        dgvLaserPieces.Columns.Add(New DataGridViewTextBoxColumn With {
            .Name = "colLaserFlat", .HeaderText = "Tiene flat", .ReadOnly = True, .Width = 72})
        dgvLaserPieces.Columns.Add(New DataGridViewTextBoxColumn With {
            .Name = "colLaserQty", .HeaderText = "Cantidad", .ReadOnly = True, .Width = 72})
        dgvLaserPieces.Columns.Add(New DataGridViewTextBoxColumn With {
            .Name = "colLaserStatus", .HeaderText = "Estado", .ReadOnly = True, .Width = 90})
        dgvLaserPieces.Columns.Add(New DataGridViewTextBoxColumn With {
            .Name = "colLaserNotes", .HeaderText = "Observaciones", .MinimumWidth = 80, .FillWeight = 60})

        For Each col As DataGridViewColumn In dgvLaserPieces.Columns
            col.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
        Next
        CType(dgvLaserPieces.Columns("colLaserName"), DataGridViewColumn).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
        CType(dgvLaserPieces.Columns("colLaserPath"), DataGridViewColumn).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
        CType(dgvLaserPieces.Columns("colLaserNotes"), DataGridViewColumn).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
        dgvLaserPieces.RowTemplate.Height = 30
        dgvLaserPieces.ResumeLayout(False)
    End Sub

    Private Sub BindLaserGrid()
        If dgvLaserPieces Is Nothing Then Return
        dgvLaserPieces.Rows.Clear()
        For i As Integer = 0 To _laserPieces.Count - 1
            Dim p = _laserPieces(i)
            Dim flatTxt As String = "No aplica"
            If p.HasFlatPattern.HasValue Then flatTxt = If(p.HasFlatPattern.Value, "Sí", "No")
            Dim bentTxt As String = "?"
            If p.IsBent.HasValue Then bentTxt = If(p.IsBent.Value, "Sí", "No")
            Dim thickTxt As String = If(p.ThicknessMm.HasValue, p.ThicknessMm.Value.ToString("0.###", CultureInfo.InvariantCulture), p.ThicknessText)
            Dim rowIdx = dgvLaserPieces.Rows.Add(
                p.Include, p.FileNameNoExt, p.FilePath, p.FileType, thickTxt, bentTxt, flatTxt,
                p.Quantity.ToString(CultureInfo.InvariantCulture), p.Status, p.Notes)
            dgvLaserPieces.Rows(rowIdx).Tag = i
        Next
        UpdateLaserSummaryLabel()
    End Sub

    Private Sub UpdateLaserSummaryLabel()
        If lblLaserSummary Is Nothing Then Return
        Dim par = _laserPieces.Where(Function(p) String.Equals(p.FileType, "PAR", StringComparison.OrdinalIgnoreCase)).Count()
        Dim psm = _laserPieces.Where(Function(p) String.Equals(p.FileType, "PSM", StringComparison.OrdinalIgnoreCase)).Count()
        Dim inc = _laserPieces.Where(Function(p) p.Include).Count()
        Dim qty = _laserPieces.Where(Function(p) p.Include).Sum(Function(p) p.Quantity)
        Dim bend = _laserPieces.Where(Function(p) p.IsSheetMetal AndAlso p.IsBent.GetValueOrDefault(False)).Count()
        lblLaserSummary.Text = $"PAR={par} PSM={psm} | Incluidas={inc} | Cantidad acumulada={qty} | Plegadas={bend}"
    End Sub

    Private Sub SyncLaserGridToModel()
        If dgvLaserPieces Is Nothing Then Return
        For Each row As DataGridViewRow In dgvLaserPieces.Rows
            If row.IsNewRow OrElse row.Tag Is Nothing Then Continue For
            Dim idx As Integer = CInt(row.Tag)
            If idx < 0 OrElse idx >= _laserPieces.Count Then Continue For
            Dim p = _laserPieces(idx)
            p.Include = CBool(row.Cells("colLaserInclude").Value)
            Dim thickCell = Convert.ToString(row.Cells("colLaserThickness").Value)
            p.ThicknessText = thickCell
            Dim parsed = LaserCutSheetService.TryParseThicknessMm(thickCell)
            If parsed.HasValue Then
                p.ThicknessMm = parsed
                If p.Status = "Falta espesor" Then p.Status = "Listo"
            ElseIf String.IsNullOrWhiteSpace(thickCell) Then
                p.ThicknessMm = Nothing
                p.Status = "Falta espesor"
            End If
            Dim bentVal = Convert.ToString(row.Cells("colLaserBent").Value)
            Select Case bentVal
                Case "Sí" : p.IsBent = True
                Case "No" : p.IsBent = False
                Case Else : p.IsBent = Nothing
            End Select
        Next
        UpdateLaserSummaryLabel()
    End Sub

    Private Function ValidateLaserThicknesses() As Boolean
        SyncLaserGridToModel()
        For Each p In _laserPieces
            If Not p.Include Then Continue For
            If Not p.ThicknessMm.HasValue OrElse p.ThicknessMm.Value <= 0 Then
                EnsureLaserService()
                _laserService.LogLaser("[LASER][UI][MISSING_THICKNESS] file=" & p.FileName)
                Dim input = Microsoft.VisualBasic.Interaction.InputBox(
                    "Indica el espesor de la pieza (mm): " & p.FileNameNoExt,
                    "Espesor faltante",
                    If(p.ThicknessText, ""))
                If String.IsNullOrWhiteSpace(input) Then Return False
                p.ThicknessText = input.Trim()
                p.ThicknessMm = LaserCutSheetService.TryParseThicknessMm(p.ThicknessText)
                If Not p.ThicknessMm.HasValue Then
                    MessageBox.Show("Espesor no válido para: " & p.FileNameNoExt, "Corte láser", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return False
                End If
                p.Status = "Listo"
            End If
        Next
        BindLaserGrid()
        Return True
    End Function

    Private Async Sub btnLaserScan_Click(sender As Object, e As EventArgs) Handles btnLaserScan.Click
        Dim asmPath As String = txtInputFile.Text.Trim()
        If String.IsNullOrWhiteSpace(asmPath) OrElse Not asmPath.EndsWith(".asm", StringComparison.OrdinalIgnoreCase) Then
            MessageBox.Show("Seleccione un archivo ASM como entrada.", "Corte láser", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If
        EnsureAutoOutputFolderForInput()
        EnsureLaserService()
        SetBusy(True, "Escaneando piezas láser...", True)
        Try
            Dim outFolder = txtOutputFolder.Text.Trim()
            If Not Await EnsureAsmComponentsForLaserAsync(asmPath).ConfigureAwait(True) Then
                MessageBox.Show(
                    "No se pudo leer el ensamblaje. Compruebe Solid Edge (diálogos de rutas/enlaces) y el log.",
                    "Corte láser",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning)
                Return
            End If

            _laserPieces = _laserService.BuildPieceListFromComponentItems(_asmComponents, outFolder, _flatAvailabilityByPath)
            Dim showSe = chkKeepSolidEdgeVisible.Checked
            Await FinalizeLaserPieceListAfterScanAsync(outFolder, showSe).ConfigureAwait(True)
            _logger.Log("[LASER][UI] Escaneo completado: " & _laserPieces.Count.ToString(CultureInfo.InvariantCulture) & " piezas únicas")
        Catch ex As Exception
            _logger.LogException("btnLaserScan", ex)
            MessageBox.Show(ex.Message, "Corte láser", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            RestoreSolidEdgeVisibleIfHidden()
            SetBusy(False, "Preparado", False)
        End Try
    End Sub

    ''' <summary>Carga componentes ASM con el mismo flujo probado del motor (STA + progreso en UI).</summary>
    Private Async Function EnsureAsmComponentsForLaserAsync(asmPath As String) As Task(Of Boolean)
        If String.Equals(asmPath, _loadedAsmComponentPath, StringComparison.OrdinalIgnoreCase) AndAlso
           _asmComponents IsNot Nothing AndAlso _asmComponents.Count > 0 Then
            Return True
        End If

        Dim showSe = chkKeepSolidEdgeVisible.Checked
        Dim capture As SynchronizationContext = SynchronizationContext.Current

        _logger.Log("[LASER][ASM] Cargando componentes (AssemblyComponentService, STA)...")
        Try
            Dim items = Await StaComInvoker.Run(Function() AssemblyComponentService.LoadAssemblyComponentItems(
                asmPath,
                uniqueOnly:=True,
                showSe,
                _logger,
                Sub(phase As String, current As Integer, total As Integer)
                    If capture IsNot Nothing Then
                        capture.Post(Sub() UpdateStatus("Corte láser — " & phase), Nothing)
                    Else
                        Try
                            BeginInvoke(Sub() UpdateStatus("Corte láser — " & phase))
                        Catch
                        End Try
                    End If
                End Sub)).ConfigureAwait(True)

            FinishAsmComponentsLoadTask(Task.FromResult(items), asmPath)
            Return _asmComponents IsNot Nothing AndAlso _asmComponents.Count > 0
        Catch ex As Exception
            _logger.LogException("EnsureAsmComponentsForLaserAsync", ex)
            Return False
        End Try
    End Function

    Private Async Sub btnLaserGenerate_Click(sender As Object, e As EventArgs) Handles btnLaserGenerate.Click
        Await RunLaserCutGenerationAsync(requireDftCreatedConfirm:=True).ConfigureAwait(True)
    End Sub

    ''' <summary>Genera Piezas a pedir de Corte.dft/.dxf en «Corte Laser» (template DXF_LIMPIO, 1:1 por espesor).</summary>
    Private Async Function RunLaserCutGenerationAsync(requireDftCreatedConfirm As Boolean) As Task
        If _laserPieces Is Nothing OrElse _laserPieces.Count = 0 Then
            MessageBox.Show("Escanee primero las piezas del ASM (botón «Escanear ASM»).", "Corte láser", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If
        SyncLaserGridToModel()
        If Not ValidateLaserThicknesses() Then Return

        EnsureLaserService()
        If requireDftCreatedConfirm Then
            Dim confirm = MessageBox.Show(
                "¿Confirmas que los planos DFT de las piezas ya han sido creados?",
                "Corte láser",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question)
            If confirm <> DialogResult.Yes Then
                _laserService.LogLaser("[LASER][CONFIRM][DFT_CREATED] user=NO")
                Return
            End If
            _laserService.LogLaser("[LASER][CONFIRM][DFT_CREATED] user=YES")
        Else
            _laserService.LogLaser("[LASER][CONFIRM][DFT_CREATED] user=YES after_main_generate")
        End If

        EnsureAutoOutputFolderForInput()
        Dim templateDxf = txtTemplateDxf.Text.Trim()
        If String.IsNullOrWhiteSpace(templateDxf) Then templateDxf = Path.Combine(DefaultTemplateFolder, "DXF_LIMPIO.dft")
        If Not File.Exists(templateDxf) Then
            MessageBox.Show("No existe el template DXF_LIMPIO: " & templateDxf, "Corte láser", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        SetBusy(True, "Generando plano de corte láser...", True)
        Try
            Dim outFolder = txtOutputFolder.Text.Trim()
            Dim showSe = chkKeepSolidEdgeVisible.Checked
            Dim piecesCopy = _laserPieces.ToList()
            Dim genResult = Await StaComInvoker.Run(Function() RunLaserGenerateOnSta(outFolder, templateDxf, piecesCopy, showSe))
            BindLaserGrid()
            If genResult.Success Then
                _logger.Log("[LASER] DFT: " & genResult.DftPath)
                _logger.Log("[LASER] DXF: " & genResult.DxfPath)
                If Not String.IsNullOrWhiteSpace(genResult.BendListPath) Then
                    _logger.Log("[LASER] Lista plegado: " & genResult.BendListPath)
                End If
                If Not String.IsNullOrWhiteSpace(genResult.DftPath) AndAlso File.Exists(genResult.DftPath) Then
                    Try
                        OpenDocumentInSolidEdge(genResult.DftPath)
                        _logger.Log("[LASER][UI][OPEN_DFT] " & genResult.DftPath)
                    Catch ex As Exception
                        _logger.LogException("OpenLaserCutDft", ex)
                    End Try
                End If
                MessageBox.Show(
                    "Plano de corte generado." & System.Environment.NewLine & System.Environment.NewLine &
                    "DFT: " & genResult.DftPath & System.Environment.NewLine &
                    "DXF: " & genResult.DxfPath,
                    "Corte láser",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information)
            Else
                MessageBox.Show(If(genResult.ErrorMessage, "Error desconocido"), "Corte láser", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Catch ex As Exception
            _logger.LogException("RunLaserCutGenerationAsync", ex)
            MessageBox.Show(ex.Message, "Corte láser", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            RestoreSolidEdgeVisibleIfHidden()
            SetBusy(False, "Preparado", False)
        End Try
    End Function

    ''' <summary>Tras generar DFT de un ASM, ofrece crear el plano 1:1 de corte láser.</summary>
    Friend Async Function OfferLaserCutAfterAssemblyGenerationAsync() As Task
        Dim asmPath As String = txtInputFile.Text.Trim()
        If String.IsNullOrWhiteSpace(asmPath) OrElse Not asmPath.EndsWith(".asm", StringComparison.OrdinalIgnoreCase) Then Return

        Dim offer = MessageBox.Show(
            "¿Generar ahora el plano «Piezas a pedir de Corte» (DFT/DXF 1:1 agrupado por espesor) en la carpeta Corte Laser?",
            "Corte láser",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question)
        If offer <> DialogResult.Yes Then Return

        If tabMotors IsNot Nothing AndAlso tabPageLaserCut IsNot Nothing Then
            tabMotors.SelectedTab = tabPageLaserCut
        End If

        EnsureLaserService()
        If _laserPieces Is Nothing OrElse _laserPieces.Count = 0 Then
            SetBusy(True, "Escaneando piezas para corte láser...", True)
            Try
                If Await EnsureAsmComponentsForLaserAsync(asmPath).ConfigureAwait(True) Then
                    _laserPieces = _laserService.BuildPieceListFromComponentItems(
                        _asmComponents, txtOutputFolder.Text.Trim(), _flatAvailabilityByPath)
                    Dim showSe = chkKeepSolidEdgeVisible.Checked
                    Await FinalizeLaserPieceListAfterScanAsync(txtOutputFolder.Text.Trim(), showSe).ConfigureAwait(True)
                End If
            Catch ex As Exception
                _logger.LogException("OfferLaserCutAfterAssemblyGenerationAsync.scan", ex)
                MessageBox.Show(ex.Message, "Corte láser", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            Finally
                SetBusy(False, "Preparado", False)
            End Try
        End If

        Await RunLaserCutGenerationAsync(requireDftCreatedConfirm:=False).ConfigureAwait(True)
    End Function

    Private Function RunLaserGenerateOnSta(outputFolder As String,
                                            templateDxf As String,
                                            pieces As List(Of LaserCutPieceInfo),
                                            showSe As Boolean) As LaserCutGenerateResult
        Dim app As SolidEdgeFramework.Application = Nothing
        Dim created As Boolean = False
        Try
            OleMessageFilter.Register()
            Try
                app = CType(Marshal.GetActiveObject("SolidEdge.Application"), SolidEdgeFramework.Application)
            Catch
                Dim t = Type.GetTypeFromProgID("SolidEdge.Application")
                app = CType(Activator.CreateInstance(t), SolidEdgeFramework.Application)
                created = True
            End Try
            app.Visible = showSe
            Return _laserService.GenerateLaserCutSheet(app, pieces, outputFolder, templateDxf, showSe)
        Finally
            Try
                If app IsNot Nothing AndAlso created Then app.Quit()
            Catch
            End Try
            Try : OleMessageFilter.Revoke() : Catch : End Try
        End Try
    End Function

    Private Sub dgvLaserPieces_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles dgvLaserPieces.CellValueChanged
        If e.RowIndex < 0 Then Return
        SyncLaserGridToModel()
    End Sub

    Private Sub btnMotorLaser_Click(sender As Object, e As EventArgs) Handles btnMotorLaser.Click
        If tabMotors IsNot Nothing AndAlso tabPageLaserCut IsNot Nothing Then
            tabMotors.SelectedTab = tabPageLaserCut
        End If
    End Sub

    Private Sub InitLaserTabOnce()
        Static done As Boolean = False
        If done Then Return
        done = True
        SetupLaserDataGridView()
        If IsWindowsDarkMode() Then
            StyleLaserDataGridViewDark()
        Else
            StyleLaserDataGridViewLight()
        End If
    End Sub
End Class
