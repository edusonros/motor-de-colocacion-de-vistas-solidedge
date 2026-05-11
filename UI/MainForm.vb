Option Strict Off

Imports System.Collections.Generic
Imports System.IO
Imports System.Reflection
Imports System.Diagnostics
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Data
Imports System.Runtime.InteropServices
Imports Microsoft.Win32
Imports SolidEdgeDraft
Imports Extraer_dft_dxf_flatdxf.Services.Dimensioning.Labs

Partial Public Class MainForm
    Private ReadOnly _logger As Logger
    Private _isRunning As Boolean = False
    Private _loadedSettings As PersistedAppSettings = Nothing
    Private Const FORCE_DARK_THEME As Boolean = True
    Private Const DEFAULT_TITLE_BLOCK_SOURCE As TitleBlockPropertySource = TitleBlockPropertySource.FromModelLink
    Private Const ForceTitleBlockModeForDebug As Boolean = False
    Private Const ForcedTitleBlockMode As TitleBlockPropertySource = TitleBlockPropertySource.FromDraft
    Private Const MaxVisibleLogLines As Integer = 1500
    Private Const LogTrimCheckInterval As Integer = 40
    ''' <summary>Escala de la barra "Pieza": ~líneas de log típicas PAR/PSM por pieza.</summary>
    Private Const ExpectedLogLinesPerPiece As Integer = 360
    Private _pendingLogLinesForTrim As Integer = 0
    Private _logLinesForCurrentPiece As Integer
    Private _accumulatePieceLogLines As Boolean
    Private _runIsAssemblyJob As Boolean
    Private _progressUiTimer As System.Windows.Forms.Timer
    Private ReadOnly _runStopwatch As New Stopwatch()
    Private ReadOnly _pieceStopwatch As New Stopwatch()
    Private _lastPieceElapsed As TimeSpan = TimeSpan.Zero
    Private _currentPieceKey As String = ""
    Private _lastUiPulseTick As Integer = Environment.TickCount
    ''' <summary>Evita aplicar metadatos obsoletos si cambió rápido el archivo de entrada (carga COM en segundo plano).</summary>
    Private _metadataLoadGeneration As Long
    Private Shared ReadOnly DefaultTemplateFolder As String = "C:\Program Files\Siemens\Solid Edge 2026\Template\Conrad"
    Private _asmComponents As New List(Of AssemblyComponentItem)()
    Private ReadOnly _componentMetadataStates As New Dictionary(Of String, ComponentMetadataState)(StringComparer.OrdinalIgnoreCase)
    Private ReadOnly _flatAvailabilityByPath As New Dictionary(Of String, Boolean?)(StringComparer.OrdinalIgnoreCase)
    Private ReadOnly _componentDirtyPaths As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
    Private ReadOnly _componentExecutedPaths As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
    Private ReadOnly _manualCajetinFields As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
    Private _cajetinManualHooksRegistered As Boolean
    Private _componentDirtyHooksRegistered As Boolean
    Private _loadedAsmComponentPath As String = ""
    Private _currentRunningComponentPath As String = ""
    Private _asmUiToolTip As ToolTip
    Private _templatesEmbedded As Boolean = False
    Private _requestedDimLabFromDedicatedButton As Boolean
    Private ReadOnly _dgvTraceability As New DataGridView()
    Private _btnAnalyzeDft As Button
    Private _btnDimRelinkLab As Button
    Private _btnAdboGuidedLab As Button
    Private _btnBringSolidEdgeFront As Button
    Private _btnStopRun As Button
    Private _runDropViewsTo2DModelLab As Boolean = False
    Private _runDropCreatedSheetsDimensionLab As Boolean = False
    Private _dropCreatedSheetsLabDebugSave As Boolean = False
    Private _runDVGeometryDimensionPlacementLab As Boolean = True

    <DllImport("user32.dll")>
    Private Shared Function SetForegroundWindow(hWnd As IntPtr) As Boolean
    End Function

    <DllImport("user32.dll")>
    Private Shared Function ShowWindow(hWnd As IntPtr, nCmdShow As Integer) As Boolean
    End Function

    Private Const SW_RESTORE As Integer = 9
    Private Shared ReadOnly _excludeKeywordsForSelection As String() = {
        "skf", "nut", "2026", "2026_02", "screw", "duin", "iso", "bolt", "whaser", "washer",
        "snl", "sleeve", "22210", "22211", "22212", "fnl", "motor", "prensa", "estopada", "tornillo",
        "tuerca", "arandela", "fag"
    }

    Public Sub New()
        InitializeComponent()
        _logger = New Logger(AddressOf AppendLogLine)
        Me.Text = "Solid Edge - Generador Automatico de DFT / PDF / DXF"
        SetupAsmDataGridView()
        SetupTraceabilityDataGridView()
    End Sub

    Private Sub SetupAsmDataGridView()
        If dgvAsmComponents Is Nothing Then Return
        dgvAsmComponents.SuspendLayout()
        dgvAsmComponents.Columns.Clear()
        Dim colChk As New DataGridViewCheckBoxColumn With {
            .Name = "colSel",
            .HeaderText = "",
            .Width = 32,
            .MinimumWidth = 28,
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.None,
            .ThreeState = False
        }
        Dim colName As New DataGridViewTextBoxColumn With {
            .Name = "colName",
            .HeaderText = "Componente",
            .ReadOnly = True,
            .MinimumWidth = 250,
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
            .FillWeight = 100.0F
        }
        Dim colDatos As New DataGridViewButtonColumn With {
            .Name = "colDatos",
            .HeaderText = "",
            .Text = "Datos",
            .UseColumnTextForButtonValue = True,
            .Width = 76,
            .MinimumWidth = 60,
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.None
        }
        Dim colGuardar As New DataGridViewButtonColumn With {
            .Name = "colGuardarDatos",
            .HeaderText = "",
            .Text = "Guardar",
            .UseColumnTextForButtonValue = True,
            .Width = 78,
            .MinimumWidth = 62,
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.None
        }
        Dim colFlat As New DataGridViewButtonColumn With {
            .Name = "colFlat",
            .HeaderText = "Flat",
            .Text = "Flat?",
            .UseColumnTextForButtonValue = True,
            .Width = 60,
            .MinimumWidth = 52,
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.None
        }
        Dim colDone As New DataGridViewTextBoxColumn With {
            .Name = "colDone",
            .HeaderText = "✓",
            .ReadOnly = True,
            .Width = 34,
            .MinimumWidth = 30,
            .AutoSizeMode = DataGridViewAutoSizeColumnMode.None
        }
        dgvAsmComponents.Columns.AddRange({colChk, colName, colDatos, colGuardar, colFlat, colDone})
        dgvAsmComponents.DefaultCellStyle.WrapMode = DataGridViewTriState.True
        dgvAsmComponents.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        dgvAsmComponents.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
        dgvAsmComponents.ShowCellToolTips = True
        dgvAsmComponents.ResumeLayout(False)
    End Sub

    Private Sub ClearManualCajetinFieldTracking()
        _manualCajetinFields.Clear()
    End Sub

    Private Sub EnsureCajetinManualHooks()
        If _cajetinManualHooksRegistered Then Return
        _cajetinManualHooksRegistered = True
        AddHandler txtClient.TextChanged, Sub(s, ev) MarkCajetinManual("Cliente")
        AddHandler txtProject.TextChanged, Sub(s, ev) MarkCajetinManual("Proyecto")
        AddHandler txtPedido.TextChanged, Sub(s, ev) MarkCajetinManual("Pedido")
        AddHandler txtDrawingNumber.TextChanged, Sub(s, ev) MarkCajetinManual("Plano")
        AddHandler txtTitle.TextChanged, Sub(s, ev) MarkCajetinManual("Titulo")
        AddHandler txtRevision.TextChanged, Sub(s, ev) MarkCajetinManual("Revision")
        AddHandler txtAuthor.TextChanged, Sub(s, ev) MarkCajetinManual("Autor")
    End Sub

    Private Sub EnsureComponentDirtyHooks()
        If _componentDirtyHooksRegistered Then Return
        _componentDirtyHooksRegistered = True
        Dim markDirty As EventHandler = Sub(sender As Object, e As EventArgs) MarkLoadedComponentDirty()
        AddHandler txtMaterial.TextChanged, markDirty
        AddHandler txtThickness.TextChanged, markDirty
        AddHandler txtPartL.TextChanged, markDirty
        AddHandler txtPartH.TextChanged, markDirty
        AddHandler txtPartD.TextChanged, markDirty
        AddHandler txtPartPeso.TextChanged, markDirty
        AddHandler txtPartCant.TextChanged, markDirty
        AddHandler txtPartNombreArchivo.TextChanged, markDirty
        AddHandler txtTitle.TextChanged, markDirty
        AddHandler txtDrawingNumber.TextChanged, markDirty
    End Sub

    Private Sub MarkLoadedComponentDirty()
        If _loadingMetadataProgrammatically Then Return
        If String.IsNullOrWhiteSpace(_loadedAsmComponentPath) Then Return
        _componentDirtyPaths.Add(_loadedAsmComponentPath)
        RefreshGuardarDatosButtonByPath(_loadedAsmComponentPath)
    End Sub

    Private Sub MarkCajetinManual(fieldKey As String)
        Try
            If _loadingMetadataProgrammatically Then Return
        Catch
        End Try
        If String.IsNullOrWhiteSpace(fieldKey) Then Return
        _manualCajetinFields.Add(fieldKey)
    End Sub

    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        cmbPreferredFormat.Items.Clear()
        cmbPreferredFormat.Items.AddRange(New Object() {"Auto", "A4", "A3", "A2"})
        cmbPreferredFormat.SelectedIndex = 0
        PopulateTitleBlockSourceOptions()
        PopulateTitleSourceCombo()

        _loadedSettings = AppSettingsManager.LoadSettings()
        EnsureMainLayoutGeometry()
        EnsureDrawingPlanMetadataPanel()
        If chkStrictMetadata Is Nothing Then
            chkStrictMetadata = New CheckBox With {.Text = "Validación estricta de metadatos (bloquea generar si falta dato obligatorio)", .AutoSize = True, .Checked = False}
            flowGeneration.Controls.Add(chkStrictMetadata)
        End If
        ApplySettingsToUi(_loadedSettings)
        EnsureSolidEdgeForegroundButton()
        RemoveLabControlsFromUi()
        DimensionInsertionConfig.EnableDrawingViewDimensioningLab = False
        If ForceTitleBlockModeForDebug Then
            SetSelectedTitleBlockPropertySource(ForcedTitleBlockMode)
            cmbTitleBlockSource.Enabled = False
        End If
        ' El usuario pide iniciar siempre con entrada en blanco.
        txtInputFile.Clear()
        txtOutputFolder.Clear()
        ApplyDefaultTemplatesIfEmpty()
        ValidateTemplatePaths(showMessage:=False)
        UpdateDetectedType()
        UpdateScaleMode()
        UpdateStatus("Esperando archivo...")
        RenameUiOptionTexts()
        UpdateAsmComponentPanelVisibility()
        EmbedTemplatesInAdvancedPanel()

        ' Forzar layout inicial robusto (evita paneles apilados por splitter).
        AddHandler Me.Resize, Sub() EnsureMainLayoutGeometry()
        EnsureMainLayoutGeometry()
        ApplyWindowsTheme()
        _logger.Log("[UI][GENERATE_BUTTON][VISIBLE]")
        EnsureCajetinManualHooks()
        EnsureComponentDirtyHooks()
        ClearPlanMetadataUi("[UI][METADATA][CLEAR_ON_START]")
        If Not String.IsNullOrWhiteSpace(txtInputFile.Text) AndAlso File.Exists(txtInputFile.Text) Then
            LoadSourcePropertiesToUi()
            RefreshAsmComponentsIfNeeded()
        End If
        RefreshTitleModeUi()
        UpdateTitleBlockOriginHints()
        chkPmiRetrievalProbe.Visible = False : chkPmiRetrievalProbe.Checked = False
        chkExperimentalPmiModelView.Visible = False : chkExperimentalPmiModelView.Checked = False
        chkExperimentalDraftGeometryDiagnostics.Visible = False : chkExperimentalDraftGeometryDiagnostics.Checked = False
        EnsureAnalyzeDftButton()
        EnsureDimRelinkLabButton()
        EnsureAdboGuidedLabButton()
        EnsureStopRunButton()

        _progressUiTimer = New System.Windows.Forms.Timer()
        _progressUiTimer.Interval = 1000
        AddHandler _progressUiTimer.Tick, AddressOf RefreshProgressTelemetry
        _progressUiTimer.Start()
        ResetProgressTelemetry()
        SetProgressDeterminateDefaults()
        LogBootPathsBanner()
    End Sub

    Private Sub RemoveLabControlsFromUi()
        If flowGeneration Is Nothing Then Return

        Dim controlsToRemove As New List(Of Control)()
        If chkUnitHorizontalExteriorTest IsNot Nothing Then controlsToRemove.Add(chkUnitHorizontalExteriorTest)
        If chkDrawingViewDimensioningLab IsNot Nothing Then controlsToRemove.Add(chkDrawingViewDimensioningLab)
        If chkDimLabInteractivePause IsNot Nothing Then controlsToRemove.Add(chkDimLabInteractivePause)
        If lblDimLabMode IsNot Nothing Then controlsToRemove.Add(lblDimLabMode)
        If cmbDimLabMode IsNot Nothing Then controlsToRemove.Add(cmbDimLabMode)
        If chkDimLabVisibleProbe IsNot Nothing Then controlsToRemove.Add(chkDimLabVisibleProbe)
        If chkDimLabAlternativePlacement IsNot Nothing Then controlsToRemove.Add(chkDimLabAlternativePlacement)
        If btnDimLabRun IsNot Nothing Then controlsToRemove.Add(btnDimLabRun)

        For Each ctl In controlsToRemove
            If ctl Is Nothing Then Continue For
            Try
                ctl.Visible = False
                ctl.Enabled = False
                If ctl.Parent IsNot Nothing Then
                    ctl.Parent.Controls.Remove(ctl)
                End If
            Catch
            End Try
        Next

        Try
            chkDrawingViewDimensioningLab.Checked = False
        Catch
        End Try
        Try
            chkDimLabInteractivePause.Checked = False
        Catch
        End Try
        Try
            chkDimLabVisibleProbe.Checked = False
        Catch
        End Try
        Try
            chkDimLabAlternativePlacement.Checked = False
        Catch
        End Try
    End Sub

    Private Sub EnsureSolidEdgeForegroundButton()
        If flowGeneration Is Nothing Then Return
        If _btnBringSolidEdgeFront IsNot Nothing Then Return

        _btnBringSolidEdgeFront = New Button With {
            .Name = "btnBringSolidEdgeFront",
            .Text = "Abrir Solid Edge en primer plano",
            .AutoSize = True,
            .MinimumSize = New Size(220, 30)
        }
        AddHandler _btnBringSolidEdgeFront.Click, AddressOf btnBringSolidEdgeFront_Click
        flowGeneration.Controls.Add(_btnBringSolidEdgeFront)
    End Sub

    Private Sub btnBringSolidEdgeFront_Click(sender As Object, e As EventArgs)
        BringSolidEdgeToFront()
    End Sub

    Private Sub BringSolidEdgeToFront()
        Dim app As SolidEdgeFramework.Application = Nothing
        Dim created As Boolean = False
        Try
            Try
                app = CType(Marshal.GetActiveObject("SolidEdge.Application"), SolidEdgeFramework.Application)
            Catch
                Dim t = Type.GetTypeFromProgID("SolidEdge.Application", throwOnError:=False)
                If t IsNot Nothing Then
                    app = CType(Activator.CreateInstance(t), SolidEdgeFramework.Application)
                    created = True
                End If
            End Try

            If app Is Nothing Then
                MessageBox.Show("No se pudo abrir/conectar Solid Edge.", "Solid Edge", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Try : app.Visible = True : Catch : End Try
            Dim hwnd As IntPtr = IntPtr.Zero
            Try
                hwnd = New IntPtr(Convert.ToInt64(CallByName(app, "hWnd", CallType.Get)))
            Catch
                hwnd = IntPtr.Zero
            End Try
            If hwnd <> IntPtr.Zero Then
                Try : ShowWindow(hwnd, SW_RESTORE) : Catch : End Try
                Try : SetForegroundWindow(hwnd) : Catch : End Try
            End If

            _logger.Log("[UI][SOLIDEDGE][FOREGROUND] ok created=" & created.ToString())
        Catch ex As Exception
            _logger.LogException("BringSolidEdgeToFront", ex)
            MessageBox.Show(ex.Message, "Solid Edge", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub MainForm_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        ' En Shown ya existe tamaño real de ventana: fijar splitter aquí es más fiable.
        BeginInvoke(New Action(Sub()
                                   EnsureMainLayoutGeometry()
                                   EnsureMainLayoutGeometry()
                               End Sub))
    End Sub

    Private Sub MainForm_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        Try
            If _progressUiTimer IsNot Nothing Then
                _progressUiTimer.Stop()
                RemoveHandler _progressUiTimer.Tick, AddressOf RefreshProgressTelemetry
                _progressUiTimer.Dispose()
            End If
            Dim settings As PersistedAppSettings = BuildSettingsFromUi()
            settings.WindowLeft = Me.Left
            settings.WindowTop = Me.Top
            settings.WindowWidth = Me.Width
            settings.WindowHeight = Me.Height
            settings.WindowStateValue = CInt(Me.WindowState)
            AppSettingsManager.SaveSettings(settings)
        Catch
        End Try
    End Sub

    Private Sub EnsureMainLayoutGeometry()
        Try
            If mainLayout Is Nothing OrElse mainLayout.RowStyles.Count < 3 Then Return
            mainLayout.RowStyles(0).SizeType = SizeType.Absolute
            mainLayout.RowStyles(0).Height = 70.0F
            mainLayout.RowStyles(1).SizeType = SizeType.Absolute
            mainLayout.RowStyles(1).Height = 160.0F
            mainLayout.RowStyles(2).SizeType = SizeType.Percent
            mainLayout.RowStyles(2).Height = 100.0F
            If bodyHost IsNot Nothing AndAlso bodyHost.RowStyles.Count >= 2 Then
                bodyHost.RowStyles(1).SizeType = SizeType.Absolute
                bodyHost.RowStyles(1).Height = 58.0F
            End If
            RebalanceMainLayoutHeights()
        Catch
        End Try
    End Sub

    Private Sub RebalanceMainLayoutHeights()
        Try
            If mainLayout Is Nothing OrElse mainLayout.RowStyles.Count < 3 Then Return
            mainLayout.RowStyles(1).SizeType = SizeType.Absolute
            mainLayout.RowStyles(1).Height = 160.0F
        Catch
        End Try
    End Sub

    Private Sub EmbedTemplatesInAdvancedPanel()
        If _templatesEmbedded Then Return
        Try
            ' Opciones de procesado (tblAdvanced) solo existen como controles ocultos; templates ya están en rightLayout (Designer).
            If tblAdvanced.Parent IsNot Nothing AndAlso Not ReferenceEquals(tblAdvanced.Parent, pnlHiddenProcessing) Then
                tblAdvanced.Parent.Controls.Remove(tblAdvanced)
            End If
            If pnlHiddenProcessing IsNot Nothing AndAlso Not pnlHiddenProcessing.Controls.Contains(tblAdvanced) Then
                pnlHiddenProcessing.Controls.Add(tblAdvanced)
            End If
            _templatesEmbedded = True
            RebalanceMainLayoutHeights()
        Catch ex As Exception
            _logger.LogException("EmbedTemplatesInAdvancedPanel", ex)
        End Try
    End Sub

    Private Sub ApplyWindowsTheme()
        Dim dark As Boolean = FORCE_DARK_THEME OrElse IsWindowsDarkMode()
        If dark Then
            ApplyDarkTheme()
        Else
            ApplyLightTheme()
        End If
    End Sub

    Private Sub RenameUiOptionTexts()
        chkKeepSolidEdgeVisible.Text = "Mostrar Solid Edge mientras se genera"
        chkAutoDimensioning.Text = "Generar acotado automático"
        chkUnitHorizontalExteriorTest.Text = "Prueba aislada cota horizontal exterior"
        chkDrawingViewDimensioningLab.Text = ""
        chkIncludeIso.Text = "Incluir vista isométrica"
        chkIncludeProjected.Text = "Incluir vistas proyectadas"
    End Sub

    Private Function IsWindowsDarkMode() As Boolean
        Try
            Using key = Registry.CurrentUser.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\Themes\Personalize")
                If key IsNot Nothing Then
                    Dim v = key.GetValue("AppsUseLightTheme")
                    If v IsNot Nothing Then
                        Return (Convert.ToInt32(v) = 0)
                    End If
                End If
            End Using
        Catch
        End Try
        Return False
    End Function

    Private Sub ApplyThemeRecursive(parent As Control, back As Drawing.Color, fore As Drawing.Color)
        parent.BackColor = back
        parent.ForeColor = fore
        For Each c As Control In parent.Controls
            ApplyThemeRecursive(c, back, fore)
        Next
    End Sub

    Private Sub ApplyDarkTheme()
        Dim back As Drawing.Color = Drawing.Color.FromArgb(32, 32, 32)
        Dim panel As Drawing.Color = Drawing.Color.FromArgb(45, 45, 48)
        Dim fore As Drawing.Color = Drawing.Color.Gainsboro
        ApplyThemeRecursive(Me, back, fore)

        pnlHeader.BackColor = Drawing.Color.FromArgb(25, 45, 75)
        lblTitle.ForeColor = Drawing.Color.White
        lblSubTitle.ForeColor = Drawing.Color.LightGray

        For Each gb As GroupBox In New GroupBox() {grpInput, grpAsmComponents, grpTemplates, grpTraceability, grpGeneration, grpAdvanced, grpProgress, grpLog}
            gb.BackColor = panel
            gb.ForeColor = fore
        Next
        If grpPlanCajetinBox IsNot Nothing Then
            grpPlanCajetinBox.BackColor = Drawing.Color.FromArgb(52, 52, 58)
            grpPlanCajetinBox.ForeColor = fore
        End If
        If grpPlanPartListBox IsNot Nothing Then
            grpPlanPartListBox.BackColor = Drawing.Color.FromArgb(40, 40, 46)
            grpPlanPartListBox.ForeColor = fore
        End If

        txtLog.BackColor = Drawing.Color.FromArgb(28, 28, 28)
        txtLog.ForeColor = Drawing.Color.LightGray
        If pnlGenerateBar IsNot Nothing Then
            pnlGenerateBar.BackColor = back
        End If
        StyleAsmDataGridViewDark()
    End Sub

    Private Sub StyleAsmDataGridViewDark()
        If dgvAsmComponents Is Nothing Then Return
        dgvAsmComponents.BackgroundColor = Drawing.Color.FromArgb(45, 45, 48)
        dgvAsmComponents.BorderStyle = BorderStyle.None
        dgvAsmComponents.EnableHeadersVisualStyles = False
        dgvAsmComponents.ColumnHeadersDefaultCellStyle.BackColor = Drawing.Color.FromArgb(60, 60, 64)
        dgvAsmComponents.ColumnHeadersDefaultCellStyle.ForeColor = Drawing.Color.Gainsboro
        dgvAsmComponents.DefaultCellStyle.BackColor = Drawing.Color.FromArgb(45, 45, 48)
        dgvAsmComponents.DefaultCellStyle.ForeColor = Drawing.Color.Gainsboro
        dgvAsmComponents.DefaultCellStyle.SelectionBackColor = Drawing.Color.FromArgb(70, 90, 120)
        dgvAsmComponents.DefaultCellStyle.SelectionForeColor = Drawing.Color.White
    End Sub

    Private Sub StyleAsmDataGridViewLight()
        If dgvAsmComponents Is Nothing Then Return
        dgvAsmComponents.BackgroundColor = SystemColors.Window
        dgvAsmComponents.BorderStyle = BorderStyle.None
        dgvAsmComponents.EnableHeadersVisualStyles = True
        dgvAsmComponents.DefaultCellStyle.BackColor = SystemColors.Window
        dgvAsmComponents.DefaultCellStyle.ForeColor = SystemColors.ControlText
    End Sub

    Private Sub ApplyLightTheme()
        Dim back As Drawing.Color = Drawing.SystemColors.Control
        Dim fore As Drawing.Color = Drawing.SystemColors.ControlText
        ApplyThemeRecursive(Me, back, fore)
        pnlHeader.BackColor = Drawing.Color.FromArgb(33, 56, 86)
        lblTitle.ForeColor = Drawing.Color.White
        lblSubTitle.ForeColor = Drawing.Color.Gainsboro
        txtLog.BackColor = Drawing.Color.White
        txtLog.ForeColor = Drawing.Color.Black
        If grpPlanCajetinBox IsNot Nothing Then
            grpPlanCajetinBox.BackColor = Drawing.Color.FromArgb(245, 245, 248)
            grpPlanCajetinBox.ForeColor = fore
        End If
        If grpPlanPartListBox IsNot Nothing Then
            grpPlanPartListBox.BackColor = Drawing.Color.FromArgb(235, 236, 240)
            grpPlanPartListBox.ForeColor = fore
        End If
        StyleAsmDataGridViewLight()
    End Sub

    Private Sub btnBrowseInput_Click(sender As Object, e As EventArgs) Handles btnBrowseInput.Click
        Using ofd As New OpenFileDialog()
            ofd.Title = "Selecciona archivo de entrada"
            ofd.Filter = "Solid Edge (*.asm;*.par;*.psm)|*.asm;*.par;*.psm|ASM (*.asm)|*.asm|PAR (*.par)|*.par|PSM (*.psm)|*.psm"
            ofd.Multiselect = False
            If ofd.ShowDialog() = DialogResult.OK Then
                txtInputFile.Text = ofd.FileName
                ResetTransientUiForNewInput()
                EnsureAutoOutputFolderForInput()
                UpdateDetectedType()
                RefreshAsmComponentsIfNeeded()
            End If
        End Using
    End Sub

    Private Sub btnBrowseOut_Click(sender As Object, e As EventArgs) Handles btnBrowseOut.Click
        Using fbd As New FolderBrowserDialog()
            fbd.Description = "Selecciona carpeta de salida"
            If Directory.Exists(txtOutputFolder.Text) Then fbd.SelectedPath = txtOutputFolder.Text
            If fbd.ShowDialog() = DialogResult.OK Then
                txtOutputFolder.Text = fbd.SelectedPath
            End If
        End Using
    End Sub

    Private Sub BrowseTemplate(targetTextBox As TextBox)
        Using ofd As New OpenFileDialog()
            ofd.Title = "Selecciona template DFT"
            ofd.Filter = "Draft Template (*.dft)|*.dft|Todos los archivos|*.*"
            ofd.Multiselect = False
            If ofd.ShowDialog() = DialogResult.OK Then
                targetTextBox.Text = ofd.FileName
            End If
        End Using
    End Sub

    Private Sub btnBrowseA4_Click(sender As Object, e As EventArgs) Handles btnBrowseA4.Click
        BrowseTemplate(txtTemplateA4)
    End Sub

    Private Sub btnBrowseA3_Click(sender As Object, e As EventArgs) Handles btnBrowseA3.Click
        BrowseTemplate(txtTemplateA3)
    End Sub

    Private Sub btnBrowseA2_Click(sender As Object, e As EventArgs) Handles btnBrowseA2.Click
        BrowseTemplate(txtTemplateA2)
    End Sub

    Private Sub btnBrowseDxf_Click(sender As Object, e As EventArgs) Handles btnBrowseDxf.Click
        BrowseTemplate(txtTemplateDxf)
    End Sub

    Private Sub btnDimLabRun_Click(sender As Object, e As EventArgs) Handles btnDimLabRun.Click
        _requestedDimLabFromDedicatedButton = True
        chkDrawingViewDimensioningLab.Checked = True
        chkAutoDimensioning.Checked = False
        chkUnitHorizontalExteriorTest.Checked = False
        chkCreatePdf.Checked = False
        chkCreateDxfDraft.Checked = False
        chkCreateFlatDxf.Checked = False
        chkCreateDft.Checked = True
        If cmbDimLabMode IsNot Nothing Then cmbDimLabMode.SelectedIndex = CInt(DimLabMode.CleanFullStrict)
        If chkDimLabVisibleProbe IsNot Nothing Then chkDimLabVisibleProbe.Checked = False
        If chkDimLabInteractivePause IsNot Nothing Then chkDimLabInteractivePause.Checked = False
        DimensionInsertionConfig.EnableDrawingViewDimensioningLab = True
        btnGenerate.PerformClick()
    End Sub

    Private Sub btnGenerate_Click(sender As Object, e As EventArgs) Handles btnGenerate.Click
        If _isRunning Then Return
        Dim unused = GenerateWorkAsync()
    End Sub

    ''' <summary>Ejecuta el motor fuera del hilo UI para que ""Conectando a Solid Edge"" no congele la ventana durante el arranque COM.</summary>
    Private Async Function GenerateWorkAsync() As Task
        If _isRunning Then Return
        EnsureAutoOutputFolderForInput()

        Dim reviewMessage As String = "Quieres actualizar/revisar las propiedades antes de generar?" &
                                      Environment.NewLine &
                                      "Si eliges SI, se cancelara la generacion y volveras al bloque de Propiedades."
        Dim reviewProps As DialogResult = MessageBox.Show(reviewMessage, "Confirmacion previa", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If reviewProps = DialogResult.Yes Then
            UpdateStatus("Generacion cancelada para editar propiedades.")
            _logger.Log("Generacion cancelada por usuario para revisar propiedades del cajetin.")
            Try
                grpTraceability.Select()
                txtTitle.Focus()
            Catch
            End Try
            Return
        End If

        Dim config As JobConfiguration = BuildConfigurationFromUi()
        _logger.Log("[UI][GENERATE][VALIDATE_METADATA]")
        If Not ValidateMetadataBeforeGenerate(config) Then Return

        Dim validationErrors As List(Of String) = ValidateConfiguration(config)
        If validationErrors.Count > 0 Then
            MessageBox.Show(String.Join(Environment.NewLine, validationErrors), "Validacion", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Try
            _isRunning = True
            ToggleUi(False)
            If _btnStopRun IsNot Nothing Then _btnStopRun.Enabled = True
            txtLog.Clear()
            _logger.Reset()
            _logLinesForCurrentPiece = 0
            _accumulatePieceLogLines = False
            _componentExecutedPaths.Clear()
            _currentRunningComponentPath = ""
            ResetAsmExecutionTicks()
            _runIsAssemblyJob = (config.DetectInputKind() = SourceFileKind.AssemblyFile)
            If lblProgressAsm IsNot Nothing Then lblProgressAsm.Visible = _runIsAssemblyJob
            If progressBarAsm IsNot Nothing Then
                progressBarAsm.Visible = _runIsAssemblyJob
                progressBarAsm.Minimum = 0
                progressBarAsm.Maximum = 100
                progressBarAsm.Value = 0
            End If
            If _runIsAssemblyJob AndAlso lblProgressAsm IsNot Nothing Then
                lblProgressAsm.Text = "Ensamblaje: 0 de 0"
            End If
            ConfigurePieceLogProgressBar()
            StartProgressTelemetry()
            _logger.Log("Inicio de proceso.")

            Dim result As EngineRunResult = Await RunDraftEngineOnBackgroundStaThreadAsync(config).ConfigureAwait(True)

            If Not String.IsNullOrWhiteSpace(result.DimLabReferenceDftFullPath) Then
                _logger.Log("[DIMLAB][DONE] DFT referencia guardado")
                MessageBox.Show(
                    "Abre/revisa este archivo:" & Environment.NewLine & result.DimLabReferenceDftFullPath,
                    "DIMLAB",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information)
            ElseIf result.DimLabRunAbortedMisconfigured Then
                _logger.Log("[DIMLAB][UI] Sin MsgBox de ruta: abort por flags")
                MessageBox.Show(
                    "DIMLAB no se ejecutó: el botón [LAB] exigía laboratorio pero Effective_runLab quedó en False." & Environment.NewLine &
                    "Revise el log con [DIMLAB][ABORT].",
                    "DIMLAB",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning)
            ElseIf result.Success Then
                _logger.Log($"Proceso finalizado OK. Procesados={result.ProcessedCount}, Errores={result.ErrorCount}")
                MessageBox.Show("Proceso completado.", "Generacion", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                _logger.Log($"Proceso finalizado con incidencias. Procesados={result.ProcessedCount}, Errores={result.ErrorCount}")
                MessageBox.Show("Proceso finalizado con errores. Revisa el log.", "Generacion", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If

        Catch ex As Exception
            _logger.LogException("GenerateWorkAsync", ex)
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            _requestedDimLabFromDedicatedButton = False
            _isRunning = False
            If _btnStopRun IsNot Nothing Then _btnStopRun.Enabled = False
            ToggleUi(True)
            StopProgressTelemetry()
            ResetProgressBarsAfterJob()
            UpdateStatus("Proceso finalizado.")
        End Try
    End Function

    ''' <summary>Solid Edge COM suele requerir apartamento STA; el arranque puede tardar minutos sin bloquear la ventana.</summary>
    Private Async Function RunDraftEngineOnBackgroundStaThreadAsync(config As JobConfiguration) As Task(Of EngineRunResult)
        Return Await Task.Run(Function() RunDraftEngineOnStaThread(config)).ConfigureAwait(False)
    End Function

    Private Function RunDraftEngineOnStaThread(config As JobConfiguration) As EngineRunResult
        Dim result As EngineRunResult = Nothing
        Dim fault As Exception = Nothing
        Dim th As New Thread(
            Sub()
                Try
                    Dim eng As New DraftGenerationEngine(_logger, AddressOf HandleEngineProgress)
                    result = eng.Run(config)
                Catch ex As Exception
                    fault = ex
                End Try
            End Sub)
        th.SetApartmentState(ApartmentState.STA)
        th.IsBackground = False
        th.Name = "DraftGenSolidEdgeSta"
        th.Start()
        th.Join()
        If fault IsNot Nothing Then Throw fault
        Return result
    End Function

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        txtInputFile.Clear()
        lblDetectedTypeValue.Text = "-"
        txtLog.Clear()
        ResetTransientUiForNewInput()
    End Sub

    Private Sub btnOpenOutput_Click(sender As Object, e As EventArgs) Handles btnOpenOutput.Click
        Try
            If Directory.Exists(txtOutputFolder.Text) Then
                Process.Start("explorer.exe", txtOutputFolder.Text)
            Else
                MessageBox.Show("La carpeta de salida no existe.", "Abrir carpeta", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Abrir carpeta", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnSaveConfig_Click(sender As Object, e As EventArgs) Handles btnSaveConfig.Click
        Try
            Dim settings As PersistedAppSettings = BuildSettingsFromUi()
            AppSettingsManager.SaveSettings(settings)
            _loadedSettings = settings
            _logger.Log("Configuracion guardada.")
            MessageBox.Show("Configuracion guardada.", "Configuracion", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Guardar configuracion", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnLoadConfig_Click(sender As Object, e As EventArgs) Handles btnLoadConfig.Click
        Try
            _loadedSettings = AppSettingsManager.LoadSettings()
            ApplySettingsToUi(_loadedSettings)
            txtOutputFolder.Clear()
            ValidateTemplatePaths(showMessage:=False)
            _logger.Log("Configuracion cargada.")
            MessageBox.Show("Configuracion cargada.", "Configuracion", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Cargar configuracion", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnReloadSourceProps_Click(sender As Object, e As EventArgs) Handles btnReloadSourceProps.Click
        LoadSourcePropertiesToUi()
    End Sub

    Private Sub btnApplyTraceability_Click(sender As Object, e As EventArgs) Handles btnApplyTraceability.Click
        Try
            Dim cfg As JobConfiguration = BuildConfigurationFromUi()
            _logger.Log($"[UI][DFT] Aplicación manual iniciada para: {cfg.InputFile}")

            Dim dftPath As String = ResolveManualDftPath(cfg)
            If String.IsNullOrWhiteSpace(dftPath) Then
                _logger.Log("[UI][DFT][WARN] No se encontró DFT para el archivo actual.")
                MessageBox.Show("No se encontró el DFT correspondiente al archivo seleccionado.", "Aplicar propiedades al DFT", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            _logger.Log($"[UI][DFT] DFT encontrado: {dftPath}")
            Dim ok As Boolean = SolidEdgePropertyService.ApplyDirectSummaryInfoToDraftFile(dftPath, chkKeepSolidEdgeVisible.Checked, cfg, _logger)
            _logger.Log($"[UI][DFT] Resultado={ok}")
            UpdateStatus(If(ok, "Propiedades SummaryInfo aplicadas al DFT.", "No se pudieron aplicar propiedades al DFT."))

            If ok Then
                MessageBox.Show("Propiedades aplicadas correctamente al DFT.", "Aplicar propiedades al DFT", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                MessageBox.Show("No se pudieron aplicar propiedades al DFT. Revisa el log para más detalle.", "Aplicar propiedades al DFT", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
        Catch ex As Exception
            _logger.LogException("btnApplyTraceability_Click", ex)
            MessageBox.Show(ex.Message, "Aplicar propiedades al DFT", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnRestoreDefaultTemplates_Click(sender As Object, e As EventArgs) Handles btnRestoreDefaultTemplates.Click
        RestoreDefaultTemplatePaths()
        ValidateTemplatePaths(showMessage:=True)
    End Sub

    Private Sub btnLoadAsmComponents_Click(sender As Object, e As EventArgs) Handles btnLoadAsmComponents.Click
        RefreshAsmComponentsIfNeeded(forceReload:=True)
    End Sub

    Private Sub btnSelectAllComponents_Click(sender As Object, e As EventArgs) Handles btnSelectAllComponents.Click
        If dgvAsmComponents Is Nothing Then Return
        For Each r As DataGridViewRow In dgvAsmComponents.Rows
            If r.IsNewRow Then Continue For
            r.Cells("colSel").Value = True
        Next
    End Sub

    Private Sub btnSelectNoneComponents_Click(sender As Object, e As EventArgs) Handles btnSelectNoneComponents.Click
        If dgvAsmComponents Is Nothing Then Return
        For Each r As DataGridViewRow In dgvAsmComponents.Rows
            If r.IsNewRow Then Continue For
            r.Cells("colSel").Value = False
        Next
    End Sub

    Private Sub dgvAsmComponents_CurrentCellDirtyStateChanged(sender As Object, e As EventArgs) Handles dgvAsmComponents.CurrentCellDirtyStateChanged
        If dgvAsmComponents Is Nothing OrElse Not dgvAsmComponents.IsCurrentCellDirty Then Return
        If TypeOf dgvAsmComponents.CurrentCell Is DataGridViewCheckBoxCell Then
            dgvAsmComponents.CommitEdit(DataGridViewDataErrorContexts.Commit)
        End If
    End Sub

    ''' <summary>Evita reentrada cuando se aplica la cascada de marcado/desmarcado.</summary>
    Private _suppressAsmCascade As Boolean = False

    Private Sub dgvAsmComponents_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles dgvAsmComponents.CellValueChanged
        If _suppressAsmCascade Then Return
        If dgvAsmComponents Is Nothing OrElse e.RowIndex < 0 OrElse e.ColumnIndex < 0 Then Return
        If dgvAsmComponents.Columns(e.ColumnIndex).Name <> "colSel" Then Return

        Dim r As DataGridViewRow = dgvAsmComponents.Rows(e.RowIndex)
        If r Is Nothing OrElse r.IsNewRow OrElse r.Tag Is Nothing Then Return
        Dim idx As Integer = CInt(r.Tag)
        If idx < 0 OrElse idx >= _asmComponents.Count Then Return
        Dim it As AssemblyComponentItem = _asmComponents(idx)
        If it Is Nothing Then Return

        ' Solo cascada cuando se marca/desmarca un subensamblaje (ASM): se aplica a sus descendientes.
        If Not String.Equals(it.Kind, "ASM", StringComparison.OrdinalIgnoreCase) Then Return

        Dim newValue As Boolean
        Try
            newValue = Convert.ToBoolean(r.Cells("colSel").Value)
        Catch
            newValue = False
        End Try

        Dim parentLevel As Integer = it.Level
        Dim affected As Integer = 0
        _suppressAsmCascade = True
        Try
            For i As Integer = e.RowIndex + 1 To dgvAsmComponents.Rows.Count - 1
                Dim child As DataGridViewRow = dgvAsmComponents.Rows(i)
                If child Is Nothing OrElse child.IsNewRow OrElse child.Tag Is Nothing Then Continue For
                Dim cIdx As Integer = CInt(child.Tag)
                If cIdx < 0 OrElse cIdx >= _asmComponents.Count Then Continue For
                Dim cItem As AssemblyComponentItem = _asmComponents(cIdx)
                If cItem Is Nothing Then Continue For
                ' Detener al volver a un nivel hermano o superior al del ASM modificado.
                If cItem.Level <= parentLevel Then Exit For
                Dim cell = child.Cells("colSel")
                Dim current As Boolean
                Try
                    current = Convert.ToBoolean(cell.Value)
                Catch
                    current = False
                End Try
                If current <> newValue Then
                    cell.Value = newValue
                    affected += 1
                End If
            Next
        Finally
            _suppressAsmCascade = False
        End Try

        If affected > 0 Then
            _logger.Log($"[ASM][CASCADE] {(If(newValue, "Marcadas", "Desmarcadas"))} {affected} piezas hijas de {IO.Path.GetFileName(it.FullPath)}")
        End If
    End Sub

    Private Sub dgvAsmComponents_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvAsmComponents.CellContentClick
        If dgvAsmComponents Is Nothing OrElse e.RowIndex < 0 OrElse e.ColumnIndex < 0 Then Return
        Dim r As DataGridViewRow = dgvAsmComponents.Rows(e.RowIndex)
        If r.Tag Is Nothing Then Return
        Dim idx As Integer = CInt(r.Tag)
        Dim colName As String = dgvAsmComponents.Columns(e.ColumnIndex).Name

        If String.Equals(colName, "colFlat", StringComparison.Ordinal) Then
            HandleAsmFlatButton(idx)
            Return
        End If

        If String.Equals(colName, "colGuardarDatos", StringComparison.Ordinal) Then
            HandleAsmGuardarDatosButton(idx)
            Return
        End If

        If Not String.Equals(colName, "colDatos", StringComparison.Ordinal) Then Return
        If idx >= 0 AndAlso idx < _asmComponents.Count Then
            Dim fp As String = _asmComponents(idx).FullPath
            Dim st As ComponentMetadataState = Nothing
            If _componentMetadataStates.TryGetValue(fp, st) AndAlso st IsNot Nothing AndAlso st.Status = ComponentMetadataStatus.Complete Then
                _logger.Log("[UI][COMPONENT][DATA_BUTTON_DISABLED] reason=metadata_complete path=" & fp)
                Return
            End If
        End If
        AsmComponentReviewForIndex(idx)
    End Sub

    Private Sub dgvAsmComponents_CellToolTipTextNeeded(sender As Object, e As DataGridViewCellToolTipTextNeededEventArgs) Handles dgvAsmComponents.CellToolTipTextNeeded
        If e.RowIndex < 0 OrElse _asmComponents Is Nothing Then Return
        Dim r As DataGridViewRow = dgvAsmComponents.Rows(e.RowIndex)
        If r.Tag Is Nothing Then Return
        Dim idx As Integer = CInt(r.Tag)
        If idx < 0 OrElse idx >= _asmComponents.Count Then Return
        Dim it As AssemblyComponentItem = _asmComponents(idx)
        If it Is Nothing Then Return
        e.ToolTipText = If(String.IsNullOrWhiteSpace(it.FullPath), it.DisplayName, it.FullPath)
    End Sub

    Private Async Sub btnSaveLog_Click(sender As Object, e As EventArgs) Handles btnSaveLog.Click
        Using sfd As New SaveFileDialog()
            sfd.Title = "Guardar log"
            sfd.Filter = "Log (*.txt)|*.txt|Todos los archivos|*.*"
            sfd.FileName = $"run_log_{DateTime.Now:yyyyMMdd_HHmmss}.txt"
            If sfd.ShowDialog() = DialogResult.OK Then
                btnSaveLog.Enabled = False
                Try
                    Dim outPath As String = sfd.FileName
                    Await Task.Run(Sub() _logger.SaveToFile(outPath))
                    MessageBox.Show("Log guardado.", "Log", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Catch ex As Exception
                    MessageBox.Show(ex.Message, "Guardar log", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Finally
                    btnSaveLog.Enabled = True
                End Try
            End If
        End Using
    End Sub

    Private Sub btnClearLog_Click(sender As Object, e As EventArgs) Handles btnClearLog.Click
        txtLog.Clear()
    End Sub

    Private Sub EnsureAnalyzeDftButton()
        If flowLogButtons Is Nothing Then Return
        If _btnAnalyzeDft IsNot Nothing Then Return
        _btnAnalyzeDft = New Button With {.Text = "Analizar DFT", .AutoSize = True}
        AddHandler _btnAnalyzeDft.Click, AddressOf btnAnalyzeDft_Click
        flowLogButtons.Controls.Add(_btnAnalyzeDft)
    End Sub

    Private Sub EnsureDimRelinkLabButton()
        If flowLogButtons Is Nothing Then Return
        If _btnDimRelinkLab IsNot Nothing Then Return
        _btnDimRelinkLab = New Button With {.Text = "Lab Reenganche Cotas", .AutoSize = True}
        AddHandler _btnDimRelinkLab.Click, AddressOf btnDimRelinkLab_Click
        flowLogButtons.Controls.Add(_btnDimRelinkLab)
    End Sub

    Private Sub EnsureAdboGuidedLabButton()
        If flowLogButtons Is Nothing Then Return
        If _btnAdboGuidedLab IsNot Nothing Then Return
        _btnAdboGuidedLab = New Button With {.Text = "Lab ADBO guiado", .AutoSize = True}
        AddHandler _btnAdboGuidedLab.Click, AddressOf btnAdboGuidedLab_Click
        flowLogButtons.Controls.Add(_btnAdboGuidedLab)
    End Sub

    Private Sub EnsureStopRunButton()
        If pnlGenerateBar Is Nothing Then Return
        If _btnStopRun IsNot Nothing Then Return
        _btnStopRun = New Button With {
            .Text = "STOP",
            .Dock = DockStyle.Right,
            .Width = 120,
            .Enabled = False,
            .BackColor = Color.FromArgb(150, 35, 35),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat
        }
        AddHandler _btnStopRun.Click, AddressOf btnStopRun_Click
        pnlGenerateBar.Controls.Add(_btnStopRun)
    End Sub

    Private Sub btnStopRun_Click(sender As Object, e As EventArgs)
        If Not _isRunning Then Return
        Dim msg As String = "¿Qué deseas hacer?" & System.Environment.NewLine &
                            "Sí = Pausar/Reanudar" & System.Environment.NewLine &
                            "No = Parar ejecución" & System.Environment.NewLine &
                            "Cancelar = Seguir"
        Dim ans = MessageBox.Show(msg, "Control de ejecución", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
        If ans = DialogResult.Cancel Then Return
        If ans = DialogResult.No Then
            DraftGenerationEngine.RequestStop()
            _logger.Log("[UI][STOP] requested=True")
            Return
        End If
        If DraftGenerationEngine.IsPaused() Then
            DraftGenerationEngine.RequestResume()
            _logger.Log("[UI][PAUSE] resume=True")
        Else
            DraftGenerationEngine.RequestPause()
            _logger.Log("[UI][PAUSE] requested=True")
        End If
    End Sub

    Private Sub btnAnalyzeDft_Click(sender As Object, e As EventArgs)
        Try
            Using ofd As New OpenFileDialog()
                ofd.Title = "Selecciona DFT a auditar"
                ofd.Filter = "Draft (*.dft)|*.dft|Todos los archivos|*.*"
                ofd.Multiselect = False
                If ofd.ShowDialog() <> DialogResult.OK Then Return
                Dim dftPath As String = ofd.FileName
                _logger.Log("[DFT][AUDIT] Inicio: " & dftPath)
                Dim ok As Boolean = DftAuditService.AnalyzeDft(dftPath, _logger, chkKeepSolidEdgeVisible.Checked, txtOutputFolder.Text.Trim())
                If ok Then
                    MessageBox.Show("Audit DFT completado. Revisa DFT_INSPECT.", "Audit DFT", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    MessageBox.Show("Audit DFT falló. Revisa el log.", "Audit DFT", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                End If
            End Using
        Catch ex As Exception
            _logger.LogException("btnAnalyzeDft_Click", ex)
            MessageBox.Show(ex.Message, "Audit DFT", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnDimRelinkLab_Click(sender As Object, e As EventArgs)
        Const debugSave As Boolean = False
        Const debugCleanup As Boolean = True

        Dim app As SolidEdgeFramework.Application = Nothing
        Dim createdApp As Boolean = False
        Dim dftDoc As DraftDocument = Nothing

        Try
            Using ofd As New OpenFileDialog()
                ofd.Title = "Selecciona DFT para laboratorio reenganche"
                ofd.Filter = "Draft (*.dft)|*.dft|Todos los archivos|*.*"
                ofd.Multiselect = False
                If ofd.ShowDialog() <> DialogResult.OK Then Return

                Dim dftPath As String = ofd.FileName
                _logger.Log("[DIMRELINK][UI][START] dft=" & dftPath)
                _logger.Log("[DIMRELINK][UI][FLAGS] DebugSave=" & debugSave.ToString() & " DebugCleanup=" & debugCleanup.ToString())

                Try
                    app = CType(Marshal.GetActiveObject("SolidEdge.Application"), SolidEdgeFramework.Application)
                Catch
                    Dim t As Type = Type.GetTypeFromProgID("SolidEdge.Application")
                    If t IsNot Nothing Then
                        app = CType(Activator.CreateInstance(t), SolidEdgeFramework.Application)
                        createdApp = True
                    End If
                End Try

                If app Is Nothing Then
                    MessageBox.Show("No se pudo abrir/conectar Solid Edge.", "Lab Reenganche Cotas", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return
                End If

                Try : app.Visible = chkKeepSolidEdgeVisible.Checked : Catch : End Try
                Try : app.DisplayAlerts = False : Catch : End Try

                dftDoc = CType(app.Documents.Open(dftPath), DraftDocument)
                If dftDoc Is Nothing Then
                    MessageBox.Show("No se pudo abrir el DFT seleccionado.", "Lab Reenganche Cotas", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return
                End If

                DimensionKeypointRelinkLabV2.Run(
                    dftDoc,
                    Sub(m As String) _logger.Log(m),
                    DebugSave:=debugSave,
                    DebugCleanup:=debugCleanup)

                MessageBox.Show("Laboratorio ejecutado. Revisa el log [DIMLAB].", "Lab Reenganche Cotas", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End Using
        Catch ex As Exception
            _logger.LogException("btnDimRelinkLab_Click", ex)
            MessageBox.Show(ex.Message, "Lab Reenganche Cotas", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Try
                If dftDoc IsNot Nothing Then dftDoc.Close(False)
            Catch
            End Try
            Try
                If createdApp AndAlso app IsNot Nothing Then app.Quit()
            Catch
            End Try
            Try
                If dftDoc IsNot Nothing AndAlso Marshal.IsComObject(dftDoc) Then Marshal.ReleaseComObject(dftDoc)
            Catch
            End Try
            Try
                If app IsNot Nothing AndAlso Marshal.IsComObject(app) Then Marshal.ReleaseComObject(app)
            Catch
            End Try
            _logger.Log("[DIMRELINK][UI][END]")
        End Try
    End Sub

    Private Sub btnAdboGuidedLab_Click(sender As Object, e As EventArgs)
        Dim app As SolidEdgeFramework.Application = Nothing
        Dim createdApp As Boolean = False
        Dim dftDoc As DraftDocument = Nothing
        Dim openedByLab As Boolean = False
        Try
            Try
                app = CType(Marshal.GetActiveObject("SolidEdge.Application"), SolidEdgeFramework.Application)
            Catch
                Try
                    Dim t As Type = Type.GetTypeFromProgID("SolidEdge.Application")
                    If t IsNot Nothing Then
                        app = CType(Activator.CreateInstance(t), SolidEdgeFramework.Application)
                        createdApp = True
                    End If
                Catch
                    app = Nothing
                End Try
            End Try
            If app Is Nothing Then
                MessageBox.Show("Abre Solid Edge y un DFT antes de ejecutar el lab guiado.", "Lab ADBO guiado", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Try
                dftDoc = TryCast(app.ActiveDocument, DraftDocument)
            Catch exActive As Exception
                _logger.Log("[ADBO_GUIDED][ERR] ActiveDocument: " & exActive.Message)
                dftDoc = Nothing
            End Try
            If dftDoc Is Nothing Then
                Using ofd As New OpenFileDialog()
                    ofd.Title = "Selecciona DFT para Lab ADBO guiado"
                    ofd.Filter = "Draft (*.dft)|*.dft|Todos los archivos|*.*"
                    ofd.Multiselect = False
                    If ofd.ShowDialog() <> DialogResult.OK Then
                        MessageBox.Show("El documento activo no es un DFT y no se seleccionó ninguno.", "Lab ADBO guiado", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        Return
                    End If
                    Try
                        dftDoc = CType(app.Documents.Open(ofd.FileName), DraftDocument)
                        openedByLab = (dftDoc IsNot Nothing)
                        _logger.Log("[ADBO_GUIDED][OPEN_DFT] " & ofd.FileName)
                    Catch exOpen As Exception
                        _logger.Log("[ADBO_GUIDED][ERR] Open DFT: " & exOpen.Message)
                        MessageBox.Show("No se pudo abrir el DFT seleccionado.", "Lab ADBO guiado", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Return
                    End Try
                End Using
            End If

            ManualAdboGuidedLab.RunInteractive(app, dftDoc, Sub(m As String) _logger.Log(m))
            MessageBox.Show("Lab ADBO guiado finalizado. Revisa el log [ADBO_GUIDED].", "Lab ADBO guiado", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            _logger.LogException("btnAdboGuidedLab_Click", ex)
            MessageBox.Show("Error en Lab ADBO guiado: " & ex.Message, "Lab ADBO guiado", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Try
                If openedByLab AndAlso dftDoc IsNot Nothing Then dftDoc.Close(False)
            Catch
            End Try
            Try
                If dftDoc IsNot Nothing AndAlso Marshal.IsComObject(dftDoc) Then Marshal.ReleaseComObject(dftDoc)
            Catch
            End Try
            Try
                If createdApp AndAlso app IsNot Nothing Then app.Quit()
            Catch
            End Try
            Try
                If app IsNot Nothing AndAlso Marshal.IsComObject(app) Then Marshal.ReleaseComObject(app)
            Catch
            End Try
        End Try
    End Sub

    Private Sub txtInputFile_TextChanged(sender As Object, e As EventArgs) Handles txtInputFile.TextChanged
        ResetTransientUiForNewInput()
        Dim p = If(txtInputFile?.Text, "").Trim()
        If String.IsNullOrWhiteSpace(p) OrElse Not File.Exists(p) Then
            ClearPlanMetadataUi()
            _componentMetadataStates.Clear()
        _flatAvailabilityByPath.Clear()
        _componentDirtyPaths.Clear()
        _componentExecutedPaths.Clear()
        _currentRunningComponentPath = ""
            _loadedAsmComponentPath = ""
        Else
            LoadSourcePropertiesToUi()
        End If
        UpdateDetectedType()
        UpdateAsmComponentPanelVisibility()
        EnsureAutoOutputFolderForInput()
        RefreshTraceabilityDataGridSafe()
        RefreshTitleModeUi()
        UpdateTitleBlockOriginHints()
    End Sub

    Private Sub chkAutoScale_CheckedChanged(sender As Object, e As EventArgs) Handles chkAutoScale.CheckedChanged
        UpdateScaleMode()
    End Sub

    Private Sub chkUniqueComponents_CheckedChanged(sender As Object, e As EventArgs) Handles chkUniqueComponents.CheckedChanged
        If (New JobConfiguration With {.InputFile = txtInputFile.Text}).DetectInputKind() = SourceFileKind.AssemblyFile Then
            RefreshAsmComponentsIfNeeded(forceReload:=True)
        End If
    End Sub

    Private Sub UpdateScaleMode()
        txtManualScale.Enabled = Not chkAutoScale.Checked
    End Sub

    Private Sub UpdateDetectedType()
        Dim kind As SourceFileKind = (New JobConfiguration With {.InputFile = txtInputFile.Text}).DetectInputKind()
        Select Case kind
            Case SourceFileKind.AssemblyFile : lblDetectedTypeValue.Text = ".asm (Ensamblaje)"
            Case SourceFileKind.PartFile : lblDetectedTypeValue.Text = ".par (Pieza)"
            Case SourceFileKind.SheetMetalFile : lblDetectedTypeValue.Text = ".psm (Chapa)"
            Case Else : lblDetectedTypeValue.Text = "-"
        End Select
    End Sub

    Private Sub UpdateAsmComponentPanelVisibility()
        Dim kind As SourceFileKind = (New JobConfiguration With {.InputFile = txtInputFile.Text}).DetectInputKind()
        grpAsmComponents.Enabled = (kind = SourceFileKind.AssemblyFile)
        If kind <> SourceFileKind.AssemblyFile Then
            ClearAsmComponentListUi()
            _asmComponents.Clear()
            lblAsmComponentHint.Text = "Solo aplica para entrada ASM."
        End If
    End Sub

    Private Sub ClearAsmComponentListUi()
        If dgvAsmComponents Is Nothing Then Return
        dgvAsmComponents.Rows.Clear()
    End Sub

    Private Sub RebuildAsmComponentListUi()
        ClearAsmComponentListUi()
        If dgvAsmComponents Is Nothing Then Return
        If _asmComponents Is Nothing OrElse _asmComponents.Count = 0 Then Return

        If _asmUiToolTip Is Nothing Then
            _asmUiToolTip = New ToolTip With {.AutoPopDelay = 60000, .InitialDelay = 350, .ReshowDelay = 200, .ShowAlways = True}
        End If

        For idx As Integer = 0 To _asmComponents.Count - 1
            Dim it As AssemblyComponentItem = _asmComponents(idx)
            Dim disp As String = If(it Is Nothing, "", it.DisplayName)
            Dim n As Integer = dgvAsmComponents.Rows.Add(Not ShouldAutoUncheckByKeyword(it), disp, "Datos", "Guardar", "Flat?", "")
            dgvAsmComponents.Rows(n).Tag = idx
            SetAsmComponentRowStatus(idx, ComponentMetadataStatus.Pending)
            ApplyFlatCellStyle(dgvAsmComponents.Rows(n), it, Nothing)
            RefreshGuardarDatosButtonByPath(If(it Is Nothing, "", it.FullPath))
        Next
        If FORCE_DARK_THEME Then StyleAsmDataGridViewDark()
    End Sub

    Private Sub AsmComponentReviewForIndex(idx As Integer)
        BringSolidEdgeToFront()
        If idx < 0 OrElse idx >= _asmComponents.Count Then Return
        Dim it As AssemblyComponentItem = _asmComponents(idx)
        If it Is Nothing Then Return

        Dim isParOrPsm As Boolean =
            String.Equals(it.Kind, "PAR", StringComparison.OrdinalIgnoreCase) OrElse
            String.Equals(it.Kind, "PSM", StringComparison.OrdinalIgnoreCase)
        If Not isParOrPsm Then
            SetAsmComponentRowStatus(idx, ComponentMetadataStatus.Pending)
            MessageBox.Show(
                "«" & it.DisplayName & "» es " & it.Kind & ". La lectura de datos de plano para el cajetín/lista aplica a piezas PAR/PSM. Los subensamblajes no tienen esos metadatos de pieza única aquí.",
                "Datos de plano (componente)",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information)
            Return
        End If

        Dim path As String = it.FullPath
        _logger.Log("[UI][METADATA][COMPONENT_SELECTED] path=" & path)
        _logger.Log("[UI][METADATA][LOAD_COMPONENT_DATA] path=" & path)

        ClearPartListMetadataUiOnly()

        Dim pathSnap As String = path
        Dim idxSnap As Integer = idx
        Dim showSe As Boolean = chkKeepSolidEdgeVisible.Checked
        SetBusy(True, "Leyendo metadatos del componente...", True)
        StaComInvoker.Run(Function() As Tuple(Of Boolean, DrawingMetadataInput)
                               Dim data As DrawingMetadataInput = Nothing
                               Dim ok = DrawingMetadataService.TryLoadMetadataFromModelFile(pathSnap, showSe, _logger, data)
                               Return Tuple.Create(ok, data)
                           End Function).
            ContinueWith(Sub(t As Task(Of Tuple(Of Boolean, DrawingMetadataInput)))
                               BeginInvoke(New Action(Sub() FinishAsmComponentMetadataTask(t, pathSnap, idxSnap)))
                           End Sub)
    End Sub

    Private Sub FinishAsmComponentMetadataTask(t As Task(Of Tuple(Of Boolean, DrawingMetadataInput)), pathSnap As String, idxSnap As Integer)
        Try
            SetBusy(False, "Preparado", False)

            If dgvAsmComponents Is Nothing Then Return
            If idxSnap < 0 OrElse idxSnap >= _asmComponents.Count Then Return
            If Not String.Equals(_asmComponents(idxSnap).FullPath, pathSnap, StringComparison.OrdinalIgnoreCase) Then Return

            Dim compData As DrawingMetadataInput = Nothing
            If t.IsFaulted Then
                Dim ex As Exception = t.Exception
                Dim agg As AggregateException = TryCast(ex, AggregateException)
                If agg IsNot Nothing Then ex = agg.GetBaseException()
                _logger.LogException("AsmComponentReviewForMetadata", ex)
                SetAsmComponentRowStatus(idxSnap, ComponentMetadataStatus.PartialComplete)
                MessageBox.Show(
                    "Error al leer metadatos:" & Environment.NewLine & ex.Message,
                    "Datos de plano (error)",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning)
                Return
            End If

            Dim tup = t.Result
            If Not tup.Item1 OrElse tup.Item2 Is Nothing Then
                SetAsmComponentRowStatus(idxSnap, ComponentMetadataStatus.PartialComplete)
                MessageBox.Show(
                    "No se pudieron leer metadatos desde:" & Environment.NewLine & pathSnap & Environment.NewLine & "Revisa el log.",
                    "Datos de plano (error)",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning)
                Return
            End If

            compData = tup.Item2

        DrawingMetadataService.ApplyToUi(Me, compData, applyCajetin:=False, applyPartList:=True)

        _loadingMetadataProgrammatically = True
        Try
            If Not _manualCajetinFields.Contains("Plano") Then
                MetadataPlano = compData.Plano
                MetadataPlanoSourceLabel = If(String.IsNullOrWhiteSpace(compData.PlanoSource), "vacío", compData.PlanoSource)
                _logger.Log("[UI][METADATA][FIELD_FROM_COMPONENT] field=Plano")
            End If
            If Not _manualCajetinFields.Contains("Titulo") Then
                MetadataTitulo = compData.Titulo
                MetadataTituloSourceLabel = If(String.IsNullOrWhiteSpace(compData.TituloSource), "vacío", compData.TituloSource)
                _logger.Log("[UI][METADATA][FIELD_FROM_COMPONENT] field=Titulo")
            End If
            If Not _manualCajetinFields.Contains("Material") Then
                MetadataMaterial = compData.Material
                MetadataMaterialSourceLabel = If(String.IsNullOrWhiteSpace(compData.MaterialSource), "vacío", compData.MaterialSource)
                _logger.Log("[UI][METADATA][FIELD_FROM_COMPONENT] field=Material")
            End If
            If Not _manualCajetinFields.Contains("Peso") Then
                MetadataPeso = compData.Peso
                MetadataPesoSourceLabel = If(String.IsNullOrWhiteSpace(compData.PesoSource), "vacío", compData.PesoSource)
                _logger.Log("[UI][METADATA][FIELD_FROM_COMPONENT] field=Peso")
            End If
        Finally
            _loadingMetadataProgrammatically = False
        End Try

        _loadedAsmComponentPath = pathSnap
        _componentDirtyPaths.Remove(pathSnap)

        Dim snapshot As DrawingMetadataInput = DrawingMetadataService.BuildFromUi(Me)
        Dim strict As Boolean = If(chkStrictMetadata Is Nothing, False, chkStrictMetadata.Checked)
        Dim vr As MetadataValidationResult = DrawingMetadataService.ValidateMetadataForComponent(pathSnap, snapshot, strict, _logger)
        Dim st As New ComponentMetadataState With {
            .ComponentPath = pathSnap,
            .Metadata = snapshot,
            .MissingRequiredFields = vr.MissingRequiredFields,
            .MissingWarningFields = vr.MissingWarningFields,
            .LastLoadedUtc = DateTime.UtcNow
        }
        If vr.Outcome = MetadataValidationOutcome.Complete Then
            st.Status = ComponentMetadataStatus.Complete
        Else
            st.Status = ComponentMetadataStatus.PartialComplete
        End If
        _componentMetadataStates(pathSnap) = st
        SetAsmComponentRowStatus(idxSnap, st.Status)
        HighlightAsmComponentRow(idxSnap)

        If st.Status = ComponentMetadataStatus.Complete Then
            _logger.Log("[UI][COMPONENT][DATA_BUTTON_DISABLED] reason=metadata_complete path=" & pathSnap)
        End If
        Catch ex As Exception
            _logger.LogException("FinishAsmComponentMetadataTask", ex)
            Try
                SetBusy(False, "Preparado", False)
            Catch
            End Try
        End Try
    End Sub

    Private Sub SetAsmComponentRowStatus(componentIndex As Integer, status As ComponentMetadataStatus)
        If dgvAsmComponents Is Nothing Then Return
        For Each r As DataGridViewRow In dgvAsmComponents.Rows
            If r.IsNewRow OrElse r.Tag Is Nothing Then Continue For
            If CInt(r.Tag) <> componentIndex Then Continue For
            Dim c As DataGridViewCell = r.Cells("colDatos")
            Select Case status
                Case ComponentMetadataStatus.Pending
                    c.Value = "Datos"
                    c.ReadOnly = False
                    c.Style.BackColor = dgvAsmComponents.DefaultCellStyle.BackColor
                    c.Style.ForeColor = dgvAsmComponents.DefaultCellStyle.ForeColor
                Case ComponentMetadataStatus.PartialComplete
                    c.Value = "Datos"
                    c.ReadOnly = False
                    c.Style.BackColor = Color.FromArgb(255, 200, 120)
                    c.Style.ForeColor = Color.Black
                Case ComponentMetadataStatus.Complete
                    c.Value = "OK"
                    c.ReadOnly = True
                    c.Style.BackColor = Color.FromArgb(180, 230, 180)
                    c.Style.ForeColor = Color.Black
            End Select
            Dim pathLog As String = ""
            If componentIndex >= 0 AndAlso componentIndex < _asmComponents.Count Then pathLog = _asmComponents(componentIndex).FullPath
            _logger.Log("[UI][COMPONENT][STATUS] path=" & pathLog & " status=" & status.ToString())
            RefreshGuardarDatosButtonByPath(pathLog)
            Return
        Next
    End Sub

    Private Sub RefreshGuardarDatosButtonByPath(componentPath As String)
        If dgvAsmComponents Is Nothing OrElse String.IsNullOrWhiteSpace(componentPath) Then Return
        For Each r As DataGridViewRow In dgvAsmComponents.Rows
            If r.IsNewRow OrElse r.Tag Is Nothing Then Continue For
            Dim idx As Integer = CInt(r.Tag)
            If idx < 0 OrElse idx >= _asmComponents.Count Then Continue For
            Dim it = _asmComponents(idx)
            If it Is Nothing OrElse Not String.Equals(it.FullPath, componentPath, StringComparison.OrdinalIgnoreCase) Then Continue For
            Dim c As DataGridViewCell = r.Cells("colGuardarDatos")
            Dim isLoaded As Boolean = _componentMetadataStates.ContainsKey(componentPath)
            Dim isDirty As Boolean = _componentDirtyPaths.Contains(componentPath)
            c.Value = If(isDirty, "Guardar*", "Guardar")
            c.ReadOnly = Not (isLoaded AndAlso isDirty)
            If Not isLoaded Then
                c.Style.BackColor = dgvAsmComponents.DefaultCellStyle.BackColor
                c.Style.ForeColor = dgvAsmComponents.DefaultCellStyle.ForeColor
            ElseIf isDirty Then
                c.Style.BackColor = Color.FromArgb(120, 180, 255)
                c.Style.ForeColor = Color.Black
            Else
                c.Style.BackColor = Color.FromArgb(180, 230, 180)
                c.Style.ForeColor = Color.Black
            End If
            Return
        Next
    End Sub

    Private Sub ApplyFlatCellStyle(row As DataGridViewRow, item As AssemblyComponentItem, hasFlat As Boolean?)
        If row Is Nothing OrElse item Is Nothing Then Return
        Dim c As DataGridViewCell = row.Cells("colFlat")
        If Not String.Equals(item.Kind, "PSM", StringComparison.OrdinalIgnoreCase) Then
            c.Value = "-"
            c.ReadOnly = True
            c.Style.BackColor = dgvAsmComponents.DefaultCellStyle.BackColor
            c.Style.ForeColor = dgvAsmComponents.DefaultCellStyle.ForeColor
            Return
        End If
        If Not hasFlat.HasValue Then
            c.Value = "Flat?"
            c.ReadOnly = False
            c.Style.BackColor = Color.FromArgb(255, 180, 0)
            c.Style.ForeColor = Color.Black
        ElseIf hasFlat.Value Then
            c.Value = "Flat OK"
            c.ReadOnly = False
            c.Style.BackColor = Color.FromArgb(0, 140, 0)
            c.Style.ForeColor = Color.White
        Else
            c.Value = "SIN FLAT"
            c.ReadOnly = False
            c.Style.BackColor = Color.FromArgb(190, 0, 0)
            c.Style.ForeColor = Color.White
        End If
    End Sub

    Private Sub BeginFlatAvailabilityScan()
        If _asmComponents Is Nothing OrElse _asmComponents.Count = 0 Then Return
        Dim psmPaths As List(Of String) = _asmComponents.
            Where(Function(it) it IsNot Nothing AndAlso String.Equals(it.Kind, "PSM", StringComparison.OrdinalIgnoreCase)).
            Select(Function(it) it.FullPath).
            Where(Function(p) Not String.IsNullOrWhiteSpace(p) AndAlso File.Exists(p)).
            Distinct(StringComparer.OrdinalIgnoreCase).
            ToList()
        If psmPaths.Count = 0 Then Return

        Dim capture As SynchronizationContext = SynchronizationContext.Current
        StaComInvoker.Run(Function() ScanFlatAvailability(psmPaths)).
            ContinueWith(Sub(t As Task(Of Dictionary(Of String, Boolean?)))
                             If capture IsNot Nothing Then
                                 capture.Post(
                                     Sub()
                                         If t Is Nothing OrElse t.IsFaulted OrElse t.Result Is Nothing Then Return
                                         For Each kv In t.Result
                                             _flatAvailabilityByPath(kv.Key) = kv.Value
                                         Next
                                         RepaintFlatColumnFromCache()
                                     End Sub, Nothing)
                             Else
                                 BeginInvoke(
                                     Sub()
                                         If t Is Nothing OrElse t.IsFaulted OrElse t.Result Is Nothing Then Return
                                         For Each kv In t.Result
                                             _flatAvailabilityByPath(kv.Key) = kv.Value
                                         Next
                                         RepaintFlatColumnFromCache()
                                     End Sub)
                             End If
                         End Sub)
    End Sub

    Private Function ScanFlatAvailability(psmPaths As List(Of String)) As Dictionary(Of String, Boolean?)
        Dim outMap As New Dictionary(Of String, Boolean?)(StringComparer.OrdinalIgnoreCase)
        Dim app As SolidEdgeFramework.Application = Nothing
        Dim created As Boolean = False
        Try
            OleMessageFilter.Register()
            Try
                app = CType(Marshal.GetActiveObject("SolidEdge.Application"), SolidEdgeFramework.Application)
            Catch
                Dim t = Type.GetTypeFromProgID("SolidEdge.Application", throwOnError:=False)
                If t Is Nothing Then Return outMap
                app = CType(Activator.CreateInstance(t), SolidEdgeFramework.Application)
                created = True
            End Try
            If app Is Nothing Then Return outMap
            app.DisplayAlerts = False
            app.Visible = False

            For Each p In psmPaths
                Dim doc As SolidEdgePart.SheetMetalDocument = Nothing
                Try
                    doc = TryCast(app.Documents.Open(p), SolidEdgePart.SheetMetalDocument)
                    Dim hasFlat As Boolean = False
                    If doc IsNot Nothing Then
                        Try
                            hasFlat = (doc.FlatPatternModels.Count > 0)
                        Catch
                            hasFlat = False
                        End Try
                        outMap(p) = hasFlat
                    Else
                        outMap(p) = Nothing
                    End If
                Catch
                    outMap(p) = Nothing
                Finally
                    Try
                        If doc IsNot Nothing Then doc.Close(False)
                    Catch
                    End Try
                    Try
                        If doc IsNot Nothing AndAlso Marshal.IsComObject(doc) Then Marshal.ReleaseComObject(doc)
                    Catch
                    End Try
                End Try
            Next
        Finally
            Try
                If app IsNot Nothing AndAlso created Then app.Quit()
            Catch
            End Try
            Try : OleMessageFilter.Revoke() : Catch : End Try
        End Try
        Return outMap
    End Function

    Private Sub RepaintFlatColumnFromCache()
        If dgvAsmComponents Is Nothing Then Return
        For Each r As DataGridViewRow In dgvAsmComponents.Rows
            If r.IsNewRow OrElse r.Tag Is Nothing Then Continue For
            Dim idx As Integer = CInt(r.Tag)
            If idx < 0 OrElse idx >= _asmComponents.Count Then Continue For
            Dim it = _asmComponents(idx)
            If it Is Nothing Then Continue For
            Dim hasFlat As Boolean? = Nothing
            If Not String.IsNullOrWhiteSpace(it.FullPath) AndAlso _flatAvailabilityByPath.ContainsKey(it.FullPath) Then
                hasFlat = _flatAvailabilityByPath(it.FullPath)
            End If
            ApplyFlatCellStyle(r, it, hasFlat)
        Next
    End Sub

    Private Sub HandleAsmFlatButton(componentIndex As Integer)
        BringSolidEdgeToFront()
        If componentIndex < 0 OrElse componentIndex >= _asmComponents.Count Then Return
        Dim it = _asmComponents(componentIndex)
        If it Is Nothing OrElse String.IsNullOrWhiteSpace(it.FullPath) Then Return
        If Not String.Equals(it.Kind, "PSM", StringComparison.OrdinalIgnoreCase) Then Return

        Dim hasFlat As Boolean? = Nothing
        If _flatAvailabilityByPath.ContainsKey(it.FullPath) Then hasFlat = _flatAvailabilityByPath(it.FullPath)
        If Not hasFlat.HasValue Then
            Dim probe = ScanFlatAvailability(New List(Of String) From {it.FullPath})
            If probe IsNot Nothing AndAlso probe.ContainsKey(it.FullPath) Then
                hasFlat = probe(it.FullPath)
                _flatAvailabilityByPath(it.FullPath) = hasFlat
                RepaintFlatColumnFromCache()
            End If
        End If
        If hasFlat.HasValue AndAlso hasFlat.Value Then
            _logger.Log("[UI][FLAT] componente con desarrollo OK: " & it.FullPath)
            MessageBox.Show("Este componente sí tiene chapa desarrollada.", "Flat", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        _logger.Log("[UI][FLAT][OPEN] sin desarrollo detectado, abriendo en Solid Edge: " & it.FullPath)
        OpenDocumentInSolidEdge(it.FullPath)
    End Sub

    Private Sub OpenDocumentInSolidEdge(filePath As String)
        If String.IsNullOrWhiteSpace(filePath) OrElse Not File.Exists(filePath) Then Return
        Try
            Dim app As SolidEdgeFramework.Application = Nothing
            Try
                app = CType(Marshal.GetActiveObject("SolidEdge.Application"), SolidEdgeFramework.Application)
            Catch
                Dim t = Type.GetTypeFromProgID("SolidEdge.Application", throwOnError:=False)
                If t Is Nothing Then Return
                app = CType(Activator.CreateInstance(t), SolidEdgeFramework.Application)
            End Try
            If app Is Nothing Then Return
            app.Visible = True
            app.DisplayAlerts = True
            app.Documents.Open(filePath)
        Catch ex As Exception
            _logger.LogException("OpenDocumentInSolidEdge", ex)
        End Try
    End Sub

    Private Sub HandleAsmGuardarDatosButton(componentIndex As Integer)
        BringSolidEdgeToFront()
        If componentIndex < 0 OrElse componentIndex >= _asmComponents.Count Then Return
        Dim it = _asmComponents(componentIndex)
        If it Is Nothing OrElse String.IsNullOrWhiteSpace(it.FullPath) OrElse Not File.Exists(it.FullPath) Then Return
        If Not _componentDirtyPaths.Contains(it.FullPath) Then
            _logger.Log("[UI][COMPONENT][SAVE][SKIP] sin cambios: " & it.FullPath)
            Return
        End If
        Dim app As SolidEdgeFramework.Application = Nothing
        Dim created As Boolean = False
        Dim modelDoc As Object = Nothing
        Try
            OleMessageFilter.Register()
            Try
                app = CType(Marshal.GetActiveObject("SolidEdge.Application"), SolidEdgeFramework.Application)
            Catch
                Dim t = Type.GetTypeFromProgID("SolidEdge.Application", throwOnError:=False)
                If t Is Nothing Then Throw New Exception("No se pudo obtener Solid Edge.Application.")
                app = CType(Activator.CreateInstance(t), SolidEdgeFramework.Application)
                created = True
            End Try
            app.Visible = chkKeepSolidEdgeVisible.Checked
            app.DisplayAlerts = False

            modelDoc = app.Documents.Open(it.FullPath)
            Dim data = DrawingMetadataService.BuildFromUi(Me)
            DrawingMetadataService.ApplyPartListSourceProperties(modelDoc, Nothing, data, _logger)
            Dim cfg As JobConfiguration = BuildConfigurationFromUi()
            cfg.InputFile = it.FullPath
            SolidEdgePropertyService.ApplyPropertiesToOpenModelDocument(app, it.FullPath, cfg, _logger)
            Try : CallByName(modelDoc, "Save", CallType.Method) : Catch : End Try

            _componentDirtyPaths.Remove(it.FullPath)
            RefreshGuardarDatosButtonByPath(it.FullPath)
            _logger.Log("[UI][COMPONENT][SAVE][OK] " & it.FullPath)
        Catch ex As Exception
            _logger.LogException("HandleAsmGuardarDatosButton", ex)
            MessageBox.Show("No se pudieron guardar datos del componente." & Environment.NewLine & ex.Message, "Guardar datos", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Finally
            Try : SolidEdgePropertyService.TryCloseComDocument(modelDoc, False) : Catch : End Try
            Try
                If app IsNot Nothing AndAlso created Then app.Quit()
            Catch
            End Try
            Try : OleMessageFilter.Revoke() : Catch : End Try
        End Try
    End Sub

    Private Sub HighlightAsmComponentRow(componentIndex As Integer)
        If dgvAsmComponents Is Nothing Then Return
        For Each r As DataGridViewRow In dgvAsmComponents.Rows
            If r.IsNewRow OrElse r.Tag Is Nothing Then Continue For
            Dim idx As Integer = CInt(r.Tag)
            If idx = componentIndex Then
                r.Cells("colName").Style.BackColor = Color.FromArgb(120, 160, 220)
            Else
                r.Cells("colName").Style.BackColor = dgvAsmComponents.DefaultCellStyle.BackColor
                r.Cells("colName").Style.ForeColor = dgvAsmComponents.DefaultCellStyle.ForeColor
            End If
        Next
    End Sub

    Private Sub EnsureAutoOutputFolderForInput()
        Try
            If String.IsNullOrWhiteSpace(txtInputFile.Text) OrElse Not File.Exists(txtInputFile.Text) Then Return
            Dim parentDir As String = Path.GetDirectoryName(txtInputFile.Text)
            If String.IsNullOrWhiteSpace(parentDir) Then Return

            Dim current As String = txtOutputFolder.Text.Trim()
            Dim shouldAuto As Boolean =
                String.IsNullOrWhiteSpace(current) OrElse
                current.Equals(parentDir, StringComparison.OrdinalIgnoreCase) OrElse
                IsAutoPlanosFolder(current, parentDir)

            If shouldAuto Then
                txtOutputFolder.Text = GetNextAutomaticPlanosFolder(parentDir)
            End If
        Catch
        End Try
    End Sub

    Private Function IsAutoPlanosFolder(folderPath As String, parentDir As String) As Boolean
        Try
            If String.IsNullOrWhiteSpace(folderPath) OrElse String.IsNullOrWhiteSpace(parentDir) Then Return False
            Dim parent As String = Path.GetDirectoryName(folderPath)
            If String.IsNullOrWhiteSpace(parent) Then Return False
            If Not parent.Equals(parentDir, StringComparison.OrdinalIgnoreCase) Then Return False
            Dim name As String = Path.GetFileName(folderPath)
            Return Regex.IsMatch(name, "^PLANOS AUTOMATICOS V\d{2}$", RegexOptions.IgnoreCase)
        Catch
            Return False
        End Try
    End Function

    Private Function GetNextAutomaticPlanosFolder(parentDir As String) As String
        For i As Integer = 0 To 999
            Dim candidate As String = Path.Combine(parentDir, $"PLANOS AUTOMATICOS V{i:00}")
            If Not Directory.Exists(candidate) Then
                Return candidate
            End If
        Next
        Return Path.Combine(parentDir, $"PLANOS AUTOMATICOS V{DateTime.Now:HHmmss}")
    End Function

    Private Sub RefreshAsmComponentsIfNeeded(Optional forceReload As Boolean = False)
        Dim kind As SourceFileKind = (New JobConfiguration With {.InputFile = txtInputFile.Text}).DetectInputKind()
        If kind <> SourceFileKind.AssemblyFile Then
            UpdateAsmComponentPanelVisibility()
            Return
        End If
        If Not File.Exists(txtInputFile.Text) Then Return
        If Not forceReload AndAlso _asmComponents.Count > 0 Then Return

        Dim asmPathSnap As String = txtInputFile.Text.Trim()
        Dim uniqueSnap As Boolean = chkUniqueComponents.Checked
        Dim showSeSnap As Boolean = chkKeepSolidEdgeVisible.Checked
        Dim capture As SynchronizationContext = SynchronizationContext.Current

        Try
            SetBusy(True, "Leyendo ensamblaje para selección...", True)
            _logger.Log("[ASM] Inicio lectura ensamblaje (subproceso STA COM)")
            _componentMetadataStates.Clear()
            _loadedAsmComponentPath = ""

            StaComInvoker.Run(Function() AssemblyComponentService.LoadAssemblyComponentItems(
                                  asmPathSnap, uniqueSnap, showSeSnap, _logger,
                                  Sub(phase As String, current As Integer, total As Integer)
                                      If capture IsNot Nothing Then
                                          capture.Post(Sub() HandleAsmReadProgress(phase, current, total), Nothing)
                                      Else
                                          BeginInvoke(Sub() HandleAsmReadProgress(phase, current, total))
                                      End If
                                  End Sub)).
                ContinueWith(Sub(t As Task(Of List(Of AssemblyComponentItem)))
                                  BeginInvoke(Sub() FinishAsmComponentsLoadTask(t, asmPathSnap))
                              End Sub)
        Catch ex As Exception
            _logger.LogException("RefreshAsmComponentsIfNeeded", ex)
            SetBusy(False, "Preparado", False)
        End Try
    End Sub

    Private Sub FinishAsmComponentsLoadTask(task As Task(Of List(Of AssemblyComponentItem)), asmPathSnap As String)
        Try
            If Not String.Equals(asmPathSnap, If(txtInputFile?.Text, "").Trim(), StringComparison.OrdinalIgnoreCase) Then
                _logger.Log("[ASM][SKIP] El archivo de entrada cambió antes de aplicar la lista de componentes.")
                Return
            End If

            If task.IsFaulted Then
                Dim ex As Exception = task.Exception
                Dim agg As AggregateException = TryCast(ex, AggregateException)
                If agg IsNot Nothing Then ex = agg.GetBaseException()
                _logger.LogException("RefreshAsmComponentsIfNeeded", ex)
                Return
            End If

            _asmComponents = If(task.Result, New List(Of AssemblyComponentItem)())

            ' Insertar el ASM raíz como primera fila virtual del listado.
            ' Si el usuario lo deja marcado, el motor genera el plano del ensamblaje completo (overview)
            ' aplicando el mismo flujo de vistas que para PAR/PSM (CreateAutomaticDraftFromModel + AddAssemblyView).
            ' Su Level=-1 hace que la cascada (al marcar/desmarcar) propague el cambio a TODOS los hijos.
            Dim rootItem As New AssemblyComponentItem With {
                .FullPath = asmPathSnap,
                .Kind = "ASM",
                .DisplayName = "[ASM RAÍZ] " & IO.Path.GetFileName(asmPathSnap),
                .Level = -1
            }
            _asmComponents.Insert(0, rootItem)

            RebuildAsmComponentListUi()
            BeginFlatAvailabilityScan()
            _logger.Log("[ASM][COMPONENTS][LOADED] count=" & _asmComponents.Count.ToString(Globalization.CultureInfo.InvariantCulture) & " (incluye fila virtual del ASM raíz)")
            Dim autoUnchecked As Integer = 0
            For Each it In _asmComponents
                If ShouldAutoUncheckByKeyword(it) Then autoUnchecked += 1
            Next
            Dim totalAsm As Integer = 0
            Dim totalPar As Integer = 0
            Dim totalPsm As Integer = 0
            For Each it In _asmComponents
                If it.Kind = "ASM" Then totalAsm += 1
                If it.Kind = "PAR" Then totalPar += 1
                If it.Kind = "PSM" Then totalPsm += 1
            Next
            lblAsmComponentHint.Text = $"Total: {_asmComponents.Count} (ASM={totalAsm}, PAR={totalPar}, PSM={totalPsm}) | Auto desmarcados={autoUnchecked} | El ASM raíz es la 1ª fila"
        Finally
            SetBusy(False, "Preparado", False)
        End Try
    End Sub

    Private Function ShouldAutoUncheckByKeyword(item As AssemblyComponentItem) As Boolean
        If item Is Nothing Then Return False
        ' El ASM raíz (Level=-1) no se autodesmarca nunca: si el usuario lo desactiva, lo hace explícitamente.
        If item.Level < 0 Then Return False
        Dim name As String = ""
        Try
            name = Path.GetFileNameWithoutExtension(item.FullPath)
        Catch
            name = item.DisplayName
        End Try
        If String.IsNullOrWhiteSpace(name) Then name = item.DisplayName
        If String.IsNullOrWhiteSpace(name) Then Return False

        Dim normalized As String = name.ToLowerInvariant()
        For Each kw In _excludeKeywordsForSelection
            If normalized.Contains(kw.ToLowerInvariant()) Then Return True
        Next
        Return False
    End Function

    Private Function ResolveDimLabModeFromCombo() As DimLabMode
        If cmbDimLabMode Is Nothing OrElse cmbDimLabMode.SelectedIndex < 0 OrElse cmbDimLabMode.SelectedIndex > 5 Then
            Return DimLabMode.Full
        End If
        Return CType(cmbDimLabMode.SelectedIndex, DimLabMode)
    End Function

    Private Function BuildConfigurationFromUi() As JobConfiguration
        EnsureAutoOutputFolderForInput()
        Dim runUnitHorizontalTest As Boolean = chkUnitHorizontalExteriorTest.Checked
        Dim runDrawingViewLab As Boolean = chkDrawingViewDimensioningLab.Checked
        Dim scaleValue As Double = 1.0
        Double.TryParse(txtManualScale.Text.Replace(",", "."), Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, scaleValue)

        Dim preferred As PreferredSheetFormat = PreferredSheetFormat.Auto
        Select Case cmbPreferredFormat.SelectedItem?.ToString()
            Case "A4" : preferred = PreferredSheetFormat.A4
            Case "A3" : preferred = PreferredSheetFormat.A3
            Case "A2" : preferred = PreferredSheetFormat.A2
            Case Else : preferred = PreferredSheetFormat.Auto
        End Select

        Dim cfg = New JobConfiguration With {
            .InputFile = txtInputFile.Text.Trim(),
            .OutputFolder = txtOutputFolder.Text.Trim(),
            .TemplateA4 = txtTemplateA4.Text.Trim(),
            .TemplateA3 = txtTemplateA3.Text.Trim(),
            .TemplateA2 = txtTemplateA2.Text.Trim(),
            .TemplateDxf = txtTemplateDxf.Text.Trim(),
            .CreateDraft = chkCreateDft.Checked,
            .CreatePdf = chkCreatePdf.Checked,
            .CreateDxfFromDraft = chkCreateDxfDraft.Checked,
            .CreateFlatDxf = chkCreateFlatDxf.Checked,
            .OpenOutputFolderWhenDone = chkOpenOutput.Checked,
            .OverwriteExisting = chkOverwrite.Checked,
            .ProcessRepeatedComponentsOnce = chkUniqueComponents.Checked,
            .DetailedLog = chkDetailedLog.Checked,
            .DebugTemplatesInspection = chkDebugTemplates.Checked,
            .KeepSolidEdgeVisible = chkKeepSolidEdgeVisible.Checked,
            .InsertPropertiesInTitleBlock = chkInsertProperties.Checked,
            .TitleBlockPropertySourceMode = ResolveTitleBlockPropertySourceMode(),
            .PreferredFormat = preferred,
            .UseAutomaticScale = chkAutoScale.Checked,
            .ManualScale = Math.Max(0.01, scaleValue),
            .IncludeIsometric = chkIncludeIso.Checked,
            .IncludeProjectedViews = chkIncludeProjected.Checked,
            .IncludeFlatInDraftWhenPsm = chkIncludeFlatInDraft.Checked,
            .EnableAutoDimensioning = If(runUnitHorizontalTest, False, chkAutoDimensioning.Checked),
            .EnableDrawingViewDimensioningLab = runDrawingViewLab,
            .RunDropViewsTo2DModelLab = _runDropViewsTo2DModelLab,
            .RunDropCreatedSheetsDimensionLab = _runDropCreatedSheetsDimensionLab,
            .DropCreatedSheetsDimensionLabDebugSave = _dropCreatedSheetsLabDebugSave,
            .RunDVGeometryDimensionPlacementLab = _runDVGeometryDimensionPlacementLab,
            .EnableDimLabInteractivePause = If(chkDimLabInteractivePause Is Nothing, True, chkDimLabInteractivePause.Checked),
            .RunUnitHorizontalExteriorDimensionTest = runUnitHorizontalTest,
            .EnablePmiRetrievalProbe = False,
            .ExperimentalCreatePMIModelViewIfMissing = False,
            .ExperimentalDraftGeometryDiagnostics = False,
            .UseBestBaseViewLogic = chkUseBestBase.Checked,
            .ClientName = txtClient.Text.Trim(),
            .ProjectName = txtProject.Text.Trim(),
            .DrawingTitle = txtTitle.Text.Trim(),
            .TitleSourceMode = GetSelectedTitleSourceMode(),
            .Material = txtMaterial.Text.Trim(),
            .Thickness = txtThickness.Text.Trim(),
            .Pedido = txtPedido.Text.Trim(),
            .AuthorName = txtAuthor.Text.Trim(),
            .Equipment = "",
            .DrawingNumber = txtDrawingNumber.Text.Trim(),
            .Revision = txtRevision.Text.Trim(),
            .Notes = txtNotes.Text.Trim(),
            .FechaPlano = MetadataFecha.ToString("yyyy-MM-dd"),
            .Weight = MetadataPeso.Trim(),
            .PartListL = MetadataL.Trim(),
            .PartListH = MetadataH.Trim(),
            .PartListD = MetadataD.Trim(),
            .PartListNombreArchivo = MetadataNombreArchivo.Trim(),
            .PartListCantidad = If(String.IsNullOrWhiteSpace(MetadataCantidad), "1", MetadataCantidad.Trim()),
            .StrictMetadataValidation = If(chkStrictMetadata Is Nothing, False, chkStrictMetadata.Checked),
            .DimLabMode = ResolveDimLabModeFromCombo(),
            .EnableDimLabVisibleProbe = If(chkDimLabVisibleProbe Is Nothing, False, chkDimLabVisibleProbe.Checked),
            .EnableDimLabAlternativePlacement = If(chkDimLabAlternativePlacement Is Nothing, False, chkDimLabAlternativePlacement.Checked),
            .EnableDimLabHorizontalControlInVerticalOnly = True,
            .DimLabKeepFailedDimensions = False,
            .DimLabCleanPreviousLabDimensions = True
        }

        If cfg.DimLabMode = DimLabMode.CleanFull Then
            cfg.EnableAutoDimensioning = False
            cfg.EnableDimLabVisibleProbe = False
            cfg.DimLabKeepFailedDimensions = False
            cfg.DimLabCleanPreviousLabDimensions = True
            cfg.CreatePdf = False
            cfg.CreateDxfFromDraft = False
            cfg.CreateFlatDxf = False
        End If

        Dim kind As SourceFileKind = cfg.DetectInputKind()
        If kind = SourceFileKind.AssemblyFile AndAlso _asmComponents.Count > 0 AndAlso dgvAsmComponents IsNot Nothing AndAlso dgvAsmComponents.Rows.Count = _asmComponents.Count Then
            cfg.SelectedComponentPaths = New List(Of String)()
            Dim totalRows As Integer = 0
            Dim totalChecked As Integer = 0
            Dim partsChecked As Integer = 0
            Dim subAsmsChecked As Integer = 0
            Dim asmRootChecked As Boolean = False
            Dim partsUnchecked As Integer = 0
            Dim subAsmsUnchecked As Integer = 0
            For Each r As DataGridViewRow In dgvAsmComponents.Rows
                If r.IsNewRow OrElse r.Tag Is Nothing Then Continue For
                Dim idx As Integer = CInt(r.Tag)
                If idx < 0 OrElse idx >= _asmComponents.Count Then Continue For
                Dim it = _asmComponents(idx)
                If it Is Nothing Then Continue For
                totalRows += 1
                Dim checkedObj = r.Cells("colSel").Value
                Dim checked As Boolean = checkedObj IsNot Nothing AndAlso Convert.ToBoolean(checkedObj)
                Dim isPart As Boolean = String.Equals(it.Kind, "PAR", StringComparison.OrdinalIgnoreCase) OrElse String.Equals(it.Kind, "PSM", StringComparison.OrdinalIgnoreCase)
                Dim isAsm As Boolean = String.Equals(it.Kind, "ASM", StringComparison.OrdinalIgnoreCase)
                Dim isAsmRoot As Boolean = isAsm AndAlso it.Level < 0 AndAlso String.Equals(it.FullPath, cfg.InputFile, StringComparison.OrdinalIgnoreCase)
                If checked Then
                    totalChecked += 1
                    If isAsmRoot Then
                        ' El ASM raíz va al target list para que el engine genere el plano del ensamblaje
                        ' completo aplicando el mismo motor de vistas que para PAR/PSM (CreateAutomaticDraftFromModel).
                        asmRootChecked = True
                        cfg.SelectedComponentPaths.Add(it.FullPath)
                    ElseIf isPart Then
                        partsChecked += 1
                        cfg.SelectedComponentPaths.Add(it.FullPath)
                    ElseIf isAsm Then
                        subAsmsChecked += 1
                        ' Los hijos del subensamblaje se procesan vía sus propias filas marcadas;
                        ' aquí no añadimos el .asm al target list para no expandirlo en el motor.
                    End If
                Else
                    If isPart Then
                        partsUnchecked += 1
                    ElseIf isAsm AndAlso Not isAsmRoot Then
                        subAsmsUnchecked += 1
                    End If
                End If
            Next
            cfg.UseSelectedComponents = True
            _logger.Log($"Selección manual ASM (estricta): marcados={totalChecked} (ASM raíz={If(asmRootChecked, "Sí", "No")}, PAR/PSM={partsChecked}, sub-ASM={subAsmsChecked}); desmarcados PAR/PSM={partsUnchecked}, sub-ASM={subAsmsUnchecked}; total componentes={totalRows}. Las piezas desmarcadas NO se procesan.")
            If totalChecked = 0 Then
                _logger.Log("[ASM][SELECTION] Ningún componente marcado: no se generará nada (ni el plano del ensamblaje completo).")
            ElseIf asmRootChecked Then
                _logger.Log("[ASM][SELECTION] ASM raíz marcado: se generará también el plano del ensamblaje completo con el motor de vistas (vista base + AddByFold + ISO).")
            End If
        End If
        If _requestedDimLabFromDedicatedButton Then
            cfg.CreateDraft = True
            cfg.EnableDrawingViewDimensioningLab = True
            cfg.EnableAutoDimensioning = False
            cfg.CreatePdf = False
            cfg.CreateDxfFromDraft = False
            cfg.CreateFlatDxf = False
            cfg.DimLabMode = DimLabMode.CleanFullStrict
            cfg.EnableDimLabVisibleProbe = False
            cfg.EnableDimLabInteractivePause = False
            cfg.DimLabKeepFailedDimensions = False
            cfg.DimLabCleanPreviousLabDimensions = True
            cfg.RequestedDimLabFromDedicatedButton = True
            DimensionInsertionConfig.EnableDrawingViewDimensioningLab = True
            _logger.Log("[DIMLAB][UI_FORCE] desde botón [LAB]: mode=CleanFullStrict CreateDraft=True EnableDrawingViewDimensioningLab=True DimensionInsertionConfig=True EnableAutoDimensioning=False CreatePdf/DxfDraft/FlatDxf=False VisibleProbe=False InteractivePause=False")
        End If

        DimensionInsertionConfig.EnableDrawingViewDimensioningLab = cfg.EnableDrawingViewDimensioningLab
        _logger.Log("[DIM] Config desde UI: EnableAutoDimensioning=" & cfg.EnableAutoDimensioning.ToString() &
                     " (chkAutoDimensioning=" & chkAutoDimensioning.Checked.ToString() & ", RunUnitHorizontalExterior=" & cfg.RunUnitHorizontalExteriorDimensionTest.ToString() & ")" &
                     ", EnableDrawingViewDimensioningLab=" & cfg.EnableDrawingViewDimensioningLab.ToString())
        _logger.Log("[PROPS][UI] Título modo=" & cfg.TitleSourceMode.ToString() & ", plantillas diagnóstico=" & cfg.DebugTemplatesInspection.ToString())
        Return cfg
    End Function

    Private Sub LogBootPathsBanner()
        Try
            _logger.Log("[BOOT][EXE_PATH] " & Assembly.GetExecutingAssembly().Location)
            _logger.Log("[BOOT][CURRENT_DIR] " & Environment.CurrentDirectory)
            _logger.Log("[BOOT][STARTUP_PATH] " & Application.StartupPath)
            _logger.Log("[BOOT][OUTPUT_ROOT] " & If(txtOutputFolder?.Text, "").Trim())
        Catch ex As Exception
            _logger.Log("[BOOT][WARN] " & ex.Message)
        End Try
    End Sub

    Private Function ResolveManualDftPath(cfg As JobConfiguration) As String
        If cfg Is Nothing Then Return ""
        If String.IsNullOrWhiteSpace(cfg.InputFile) OrElse Not File.Exists(cfg.InputFile) Then Return ""

        Dim inputExt As String = Path.GetExtension(cfg.InputFile).ToLowerInvariant()
        If inputExt = ".dft" Then Return cfg.InputFile

        Dim baseName As String = Path.GetFileNameWithoutExtension(cfg.InputFile)

        Dim candidates As New List(Of String)()
        If Not String.IsNullOrWhiteSpace(cfg.OutputFolder) Then
            Dim outDftDir As String = Path.Combine(cfg.OutputFolder, "DFT")
            If Directory.Exists(outDftDir) Then
                candidates.AddRange(SafeFindFiles(outDftDir, baseName & "*.dft"))
                Dim exact As String = Path.Combine(outDftDir, baseName & ".dft")
                If File.Exists(exact) Then Return exact
            End If
        End If

        Dim inputDir As String = Path.GetDirectoryName(cfg.InputFile)
        If Not String.IsNullOrWhiteSpace(inputDir) AndAlso Directory.Exists(inputDir) Then
            candidates.AddRange(SafeFindFiles(inputDir, baseName & "*.dft"))
            Dim exactLocal As String = Path.Combine(inputDir, baseName & ".dft")
            If File.Exists(exactLocal) Then Return exactLocal
        End If

        Dim best As String = ""
        Dim bestTime As DateTime = DateTime.MinValue
        For Each c In candidates
            Try
                If Not File.Exists(c) Then Continue For
                Dim t As DateTime = File.GetLastWriteTimeUtc(c)
                If best = "" OrElse t > bestTime Then
                    best = c
                    bestTime = t
                End If
            Catch
            End Try
        Next
        Return best
    End Function

    Private Function SafeFindFiles(folder As String, pattern As String) As IEnumerable(Of String)
        Try
            Return Directory.GetFiles(folder, pattern, SearchOption.TopDirectoryOnly)
        Catch
            Return Array.Empty(Of String)()
        End Try
    End Function

    Private Function ResolveTitleBlockPropertySourceMode() As TitleBlockPropertySource
        Dim mode As TitleBlockPropertySource = GetSelectedTitleBlockPropertySource()
        If ForceTitleBlockModeForDebug Then mode = ForcedTitleBlockMode
        Return mode
    End Function

    Private Function ValidateConfiguration(config As JobConfiguration) As List(Of String)
        Dim errors As New List(Of String)()

        If String.IsNullOrWhiteSpace(config.InputFile) Then
            errors.Add("- Debes seleccionar un archivo de entrada.")
        ElseIf Not File.Exists(config.InputFile) Then
            errors.Add("- El archivo de entrada no existe.")
        End If

        If config.DetectInputKind() = SourceFileKind.Unknown Then
            errors.Add("- Extension no valida. Debe ser .asm, .par o .psm.")
        End If

        If String.IsNullOrWhiteSpace(config.OutputFolder) Then
            errors.Add("- Debes indicar carpeta de salida.")
        Else
            Try
                If Not Directory.Exists(config.OutputFolder) Then Directory.CreateDirectory(config.OutputFolder)
            Catch ex As Exception
                errors.Add("- Carpeta de salida invalida: " & ex.Message)
            End Try
        End If

        If config.CreateDraft OrElse config.CreatePdf OrElse config.CreateDxfFromDraft Then
            If String.IsNullOrWhiteSpace(config.TemplateA4) AndAlso String.IsNullOrWhiteSpace(config.TemplateA3) AndAlso String.IsNullOrWhiteSpace(config.TemplateA2) Then
                errors.Add("- Debes informar al menos un template A4/A3/A2 para generar Draft/PDF/DXF.")
            End If
        End If
        Dim dftTplPaths As String() = {config.TemplateA4, config.TemplateA3, config.TemplateA2}
        Dim anyDftTemplateExisting As Boolean = dftTplPaths.Any(Function(p) Not String.IsNullOrWhiteSpace(p) AndAlso File.Exists(p))
        If (config.CreateDraft OrElse config.CreatePdf OrElse config.CreateDxfFromDraft) AndAlso Not anyDftTemplateExisting Then
            errors.Add("- Debes tener al menos un template A4/A3/A2 existente en disco.")
        End If

        If config.CreateDxfFromDraft AndAlso String.IsNullOrWhiteSpace(config.TemplateDxf) Then
            errors.Add("- Debes informar Template DXF limpio para generar DXF del draft.")
        ElseIf config.CreateDxfFromDraft AndAlso Not File.Exists(config.TemplateDxf) Then
            errors.Add("- El Template DXF limpio no existe en disco.")
        End If

        If Not config.UseAutomaticScale AndAlso config.ManualScale <= 0 Then
            errors.Add("- La escala manual debe ser mayor que cero.")
        End If

        Return errors
    End Function

    Private Sub HandleEngineProgress(info As EngineProgressInfo)
        If info Is Nothing Then Return
        If Me.InvokeRequired Then
            Me.BeginInvoke(New Action(Of EngineProgressInfo)(AddressOf HandleEngineProgress), info)
            Return
        End If
        If IsAsmStatus(info.Status) Then
            SetBusy(True, info.Status, True)
            Return
        End If
        SetProgressDeterminateDefaults()
        UpdatePieceTelemetry(info.Status)
        UpdateAsmJobProgressFromStatus(info.Status)
        UpdateStatus(info.Status)
        RefreshProgressTelemetry(Nothing, EventArgs.Empty)
    End Sub

    Private Sub AppendLogLine(line As String)
        If InvokeRequired Then
            BeginInvoke(New Action(Of String)(AddressOf AppendLogLine), line)
            Return
        End If
        If txtLog Is Nothing OrElse txtLog.IsDisposed Then Return
        txtLog.AppendText(line & Environment.NewLine)
        If _isRunning AndAlso _accumulatePieceLogLines Then
            _logLinesForCurrentPiece += 1
            UpdatePieceProgressBarFromLogLines()
        End If
        _pendingLogLinesForTrim += 1
        If _pendingLogLinesForTrim >= LogTrimCheckInterval Then
            _pendingLogLinesForTrim = 0
            TrimVisibleLogIfNeeded()
        End If
        txtLog.SelectionStart = txtLog.TextLength
        txtLog.ScrollToCaret()
    End Sub

    Private Sub TrimVisibleLogIfNeeded()
        If txtLog Is Nothing Then Return
        Dim lines As String() = txtLog.Lines
        If lines Is Nothing OrElse lines.Length <= MaxVisibleLogLines Then Return
        Dim keepFrom As Integer = Math.Max(0, lines.Length - MaxVisibleLogLines)
        Dim sb As New StringBuilder()
        For i As Integer = keepFrom To lines.Length - 1
            sb.AppendLine(lines(i))
        Next
        txtLog.Text = sb.ToString()
        txtLog.SelectionStart = txtLog.TextLength
        txtLog.ScrollToCaret()
    End Sub

    Private Sub ToggleUi(enabled As Boolean)
        btnGenerate.Enabled = enabled
        btnClear.Enabled = enabled
        btnOpenOutput.Enabled = enabled
        btnSaveConfig.Enabled = enabled
        btnLoadConfig.Enabled = enabled
        grpInput.Enabled = enabled
        grpAsmComponents.Enabled = enabled
        grpTemplates.Enabled = enabled
        grpTraceability.Enabled = enabled
        grpGeneration.Enabled = enabled
        grpAdvanced.Enabled = enabled
        If _btnAnalyzeDft IsNot Nothing Then _btnAnalyzeDft.Enabled = enabled
        If _btnDimRelinkLab IsNot Nothing Then _btnDimRelinkLab.Enabled = enabled
        If _btnAdboGuidedLab IsNot Nothing Then _btnAdboGuidedLab.Enabled = enabled
        If chkStrictMetadata IsNot Nothing Then chkStrictMetadata.Enabled = enabled
    End Sub

    Private Sub HandleAsmReadProgress(phase As String, current As Integer, total As Integer)
        If InvokeRequired Then
            BeginInvoke(New Action(Sub() HandleAsmReadProgress(phase, current, total)))
            Return
        End If
        Dim msg As String = $"[ASM] {phase}"
        If total > 0 Then
            SetProgressDeterminateDefaults()
            UpdateProgress(current, total, msg)
        Else
            SetBusy(True, msg, True)
        End If
        PulseUiIfNeeded()
    End Sub

    Private Sub SetBusy(isBusy As Boolean, status As String, useMarquee As Boolean)
        If isBusy AndAlso useMarquee Then
            progressBar.Style = ProgressBarStyle.Marquee
            progressBar.MarqueeAnimationSpeed = 18
            UpdateStatus(status)
            lblStatusValue.Text = status
            PulseUiIfNeeded()
            Return
        End If

        SetProgressDeterminateDefaults()
        If Not String.IsNullOrWhiteSpace(status) Then
            UpdateStatus(status)
        End If
        PulseUiIfNeeded()
    End Sub

    Private Sub SetProgressDeterminateDefaults()
        progressBar.MarqueeAnimationSpeed = 0
        progressBar.Style = ProgressBarStyle.Continuous
        progressBar.Minimum = 0
        If progressBar.Maximum <= 0 Then progressBar.Maximum = 100
    End Sub

    Private Sub UpdateProgress(current As Integer, total As Integer, message As String)
        SetProgressDeterminateDefaults()
        If total > 0 Then
            progressBar.Maximum = Math.Max(1, total)
            progressBar.Value = Math.Max(0, Math.Min(progressBar.Maximum, current))
        Else
            progressBar.Maximum = 100
            progressBar.Value = Math.Max(0, Math.Min(100, current))
        End If
        UpdateStatus(message)
    End Sub

    Private Function IsAsmStatus(status As String) As Boolean
        If String.IsNullOrWhiteSpace(status) Then Return False
        Return status.IndexOf("[ASM]", StringComparison.OrdinalIgnoreCase) >= 0
    End Function

    Private Sub PulseUiIfNeeded()
        Dim nowTick As Integer = Environment.TickCount
        If Math.Abs(nowTick - _lastUiPulseTick) < 90 Then Return
        _lastUiPulseTick = nowTick
        Application.DoEvents()
    End Sub

    Private Sub ResetTransientUiForNewInput()
        If _isRunning Then Return
        ResetProgressBarsAfterJob()
        ResetProgressTelemetry()
        _currentPieceKey = ""
        UpdateStatus("Preparado")
        lblStatusValue.Text = "Preparado"
        RefreshTitleModeUi()
    End Sub

    Private Sub StartProgressTelemetry()
        _currentPieceKey = ""
        _lastPieceElapsed = TimeSpan.Zero
        _pieceStopwatch.Reset()
        _runStopwatch.Reset()
        _runStopwatch.Start()
        RefreshProgressTelemetry(Nothing, EventArgs.Empty)
    End Sub

    Private Sub StopProgressTelemetry()
        If _pieceStopwatch.IsRunning Then
            _lastPieceElapsed = _pieceStopwatch.Elapsed
            _pieceStopwatch.Stop()
        End If
        If _runStopwatch.IsRunning Then
            _runStopwatch.Stop()
        End If
        SetProgressDeterminateDefaults()
        RefreshProgressTelemetry(Nothing, EventArgs.Empty)
    End Sub

    Private Sub ResetProgressTelemetry()
        _currentPieceKey = ""
        _lastPieceElapsed = TimeSpan.Zero
        _pieceStopwatch.Reset()
        _runStopwatch.Reset()
        lblPieceTimeValue.Text = "00:00"
        lblTotalTimeValue.Text = "00:00"
        RefreshProgressTelemetry(Nothing, EventArgs.Empty)
    End Sub

    Private Sub UpdatePieceTelemetry(status As String)
        If String.IsNullOrWhiteSpace(status) Then Return

        Dim processingMatch As Match = Regex.Match(status, "^Procesando\s+(\d+)/(\d+)\s+-\s+(.+)$", RegexOptions.IgnoreCase)
        If processingMatch.Success Then
            Dim currentFileName As String = processingMatch.Groups(3).Value.Trim()
            Dim pieceKey As String = $"{processingMatch.Groups(1).Value}/{processingMatch.Groups(2).Value}:{currentFileName}"
            If Not String.Equals(pieceKey, _currentPieceKey, StringComparison.OrdinalIgnoreCase) Then
                _currentPieceKey = pieceKey
                _logLinesForCurrentPiece = 0
                _accumulatePieceLogLines = True
                UpdatePieceProgressBarFromLogLines()
                _lastPieceElapsed = TimeSpan.Zero
                _pieceStopwatch.Reset()
                _pieceStopwatch.Start()
            End If
            _currentRunningComponentPath = ResolveAsmComponentPathFromFileName(currentFileName)
            MarkComponentExecutionState(_currentRunningComponentPath, inProgress:=True, completed:=False)
            Return
        End If

        Dim completedMatch As Match = Regex.Match(status, "^Finalizado\s+(\d+)/(\d+)", RegexOptions.IgnoreCase)
        If completedMatch.Success Then
            _accumulatePieceLogLines = False
            If progressBar IsNot Nothing AndAlso _isRunning Then
                progressBar.Value = progressBar.Maximum
            End If
            If _pieceStopwatch.IsRunning Then
                _lastPieceElapsed = _pieceStopwatch.Elapsed
                _pieceStopwatch.Stop()
            End If
            MarkComponentExecutionState(_currentRunningComponentPath, inProgress:=False, completed:=True)
            _currentRunningComponentPath = ""
        End If
    End Sub

    Private Function ResolveAsmComponentPathFromFileName(fileName As String) As String
        If String.IsNullOrWhiteSpace(fileName) OrElse _asmComponents Is Nothing OrElse _asmComponents.Count = 0 Then Return ""
        Dim exact As String = _asmComponents.
            Where(Function(it) it IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(it.FullPath)).
            Select(Function(it) it.FullPath).
            FirstOrDefault(Function(p) String.Equals(Path.GetFileName(p), fileName, StringComparison.OrdinalIgnoreCase) AndAlso Not _componentExecutedPaths.Contains(p))
        If Not String.IsNullOrWhiteSpace(exact) Then Return exact
        Return _asmComponents.
            Where(Function(it) it IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(it.FullPath)).
            Select(Function(it) it.FullPath).
            FirstOrDefault(Function(p) String.Equals(Path.GetFileName(p), fileName, StringComparison.OrdinalIgnoreCase))
    End Function

    Private Sub MarkComponentExecutionState(componentPath As String, inProgress As Boolean, completed As Boolean)
        If dgvAsmComponents Is Nothing OrElse String.IsNullOrWhiteSpace(componentPath) Then Return
        If completed Then _componentExecutedPaths.Add(componentPath)
        For Each r As DataGridViewRow In dgvAsmComponents.Rows
            If r.IsNewRow OrElse r.Tag Is Nothing Then Continue For
            Dim idx As Integer = CInt(r.Tag)
            If idx < 0 OrElse idx >= _asmComponents.Count Then Continue For
            Dim it = _asmComponents(idx)
            If it Is Nothing OrElse Not String.Equals(it.FullPath, componentPath, StringComparison.OrdinalIgnoreCase) Then Continue For
            Dim cDone As DataGridViewCell = r.Cells("colDone")
            If completed Then
                cDone.Value = "✔"
                cDone.Style.ForeColor = Color.FromArgb(140, 220, 140)
                cDone.Style.BackColor = dgvAsmComponents.DefaultCellStyle.BackColor
            ElseIf inProgress Then
                cDone.Value = "…"
                cDone.Style.ForeColor = Color.FromArgb(245, 220, 100)
                cDone.Style.BackColor = dgvAsmComponents.DefaultCellStyle.BackColor
            End If
            Return
        Next
    End Sub

    Private Sub ConfigurePieceLogProgressBar()
        If progressBar Is Nothing Then Return
        progressBar.Minimum = 0
        progressBar.Maximum = Math.Max(1, ExpectedLogLinesPerPiece)
        progressBar.Value = 0
    End Sub

    Private Sub UpdatePieceProgressBarFromLogLines()
        If progressBar Is Nothing Then Return
        If Not _isRunning Then Return
        progressBar.Minimum = 0
        progressBar.Maximum = Math.Max(1, ExpectedLogLinesPerPiece)
        Dim cap As Integer = ExpectedLogLinesPerPiece
        Dim v As Integer = Math.Max(0, Math.Min(cap, _logLinesForCurrentPiece))
        progressBar.Value = v
    End Sub

    Private Sub UpdateAsmJobProgressFromStatus(status As String)
        If Not _runIsAssemblyJob OrElse Not _isRunning Then Return
        If progressBarAsm Is Nothing Then Return
        If String.IsNullOrWhiteSpace(status) Then Return

        Dim proc As Match = Regex.Match(status, "^Procesando\s+(\d+)/(\d+)\s", RegexOptions.IgnoreCase)
        If proc.Success Then
            Dim i As Integer = Integer.Parse(proc.Groups(1).Value, Globalization.CultureInfo.InvariantCulture)
            Dim n As Integer = Integer.Parse(proc.Groups(2).Value, Globalization.CultureInfo.InvariantCulture)
            If n > 0 Then
                If lblProgressAsm IsNot Nothing Then lblProgressAsm.Text = $"Ensamblaje: {Math.Max(0, i - 1)} de {n}"
                progressBarAsm.Maximum = 100
                Dim pct As Integer = CInt(Math.Floor(100.0 * (i - 1) / n))
                progressBarAsm.Value = Math.Max(0, Math.Min(100, pct))
            End If
            Return
        End If

        Dim fin As Match = Regex.Match(status, "^Finalizado\s+(\d+)/(\d+)", RegexOptions.IgnoreCase)
        If fin.Success Then
            Dim i As Integer = Integer.Parse(fin.Groups(1).Value, Globalization.CultureInfo.InvariantCulture)
            Dim n As Integer = Integer.Parse(fin.Groups(2).Value, Globalization.CultureInfo.InvariantCulture)
            If n > 0 Then
                If lblProgressAsm IsNot Nothing Then lblProgressAsm.Text = $"Ensamblaje: {Math.Max(0, i)} de {n}"
                progressBarAsm.Maximum = 100
                Dim pct As Integer = CInt(Math.Ceiling(100.0 * i / n))
                progressBarAsm.Value = Math.Max(0, Math.Min(100, pct))
            End If
        End If
    End Sub

    Private Sub ResetProgressBarsAfterJob()
        _accumulatePieceLogLines = False
        _logLinesForCurrentPiece = 0
        _runIsAssemblyJob = False
        If lblProgressAsm IsNot Nothing Then
            lblProgressAsm.Visible = False
            lblProgressAsm.Text = "Ensamblaje (piezas marcadas)"
        End If
        If progressBarAsm IsNot Nothing Then
            progressBarAsm.Visible = False
            progressBarAsm.Value = 0
        End If
        SetProgressDeterminateDefaults()
        If progressBar IsNot Nothing Then
            progressBar.Maximum = 100
            progressBar.Value = 0
        End If
    End Sub

    Private Sub ResetAsmExecutionTicks()
        If dgvAsmComponents Is Nothing Then Return
        For Each r As DataGridViewRow In dgvAsmComponents.Rows
            If r.IsNewRow Then Continue For
            Try
                r.Cells("colDone").Value = ""
            Catch
            End Try
        Next
    End Sub

    Private Sub RefreshProgressTelemetry(sender As Object, e As EventArgs)
        If lblCurrentTimeValue Is Nothing OrElse lblCurrentTimeValue.IsDisposed Then Return

        lblCurrentTimeValue.Text = DateTime.Now.ToString("HH:mm:ss")
        Dim pieceElapsed As TimeSpan = If(_pieceStopwatch.IsRunning, _pieceStopwatch.Elapsed, _lastPieceElapsed)
        lblPieceTimeValue.Text = FormatElapsedTime(pieceElapsed)
        lblTotalTimeValue.Text = FormatElapsedTime(_runStopwatch.Elapsed)
    End Sub

    Private Function FormatElapsedTime(value As TimeSpan) As String
        If value.TotalHours >= 1 Then
            Return value.ToString("hh\:mm\:ss")
        End If
        Return value.ToString("mm\:ss")
    End Function

    Private Sub UpdateStatus(text As String)
        lblStatusValue.Text = text
    End Sub

    Private Function BuildSettingsFromUi() As PersistedAppSettings
        Dim cfg As JobConfiguration = BuildConfigurationFromUi()
        Return New PersistedAppSettings With {
            .LastInputFile = "",
            .LastOutputFolder = cfg.OutputFolder,
            .TemplateA4 = cfg.TemplateA4,
            .TemplateA3 = cfg.TemplateA3,
            .TemplateA2 = cfg.TemplateA2,
            .TemplateDxf = cfg.TemplateDxf,
            .CreateDraft = cfg.CreateDraft,
            .CreatePdf = cfg.CreatePdf,
            .CreateDxfFromDraft = cfg.CreateDxfFromDraft,
            .CreateFlatDxf = cfg.CreateFlatDxf,
            .OpenOutputFolderWhenDone = cfg.OpenOutputFolderWhenDone,
            .OverwriteExisting = cfg.OverwriteExisting,
            .ProcessRepeatedComponentsOnce = cfg.ProcessRepeatedComponentsOnce,
            .DetailedLog = cfg.DetailedLog,
            .DebugTemplatesInspection = cfg.DebugTemplatesInspection,
            .KeepSolidEdgeVisible = cfg.KeepSolidEdgeVisible,
            .InsertPropertiesInTitleBlock = cfg.InsertPropertiesInTitleBlock,
            .TitleBlockPropertySourceMode = cfg.TitleBlockPropertySourceMode,
            .PreferredFormat = cfg.PreferredFormat,
            .UseAutomaticScale = cfg.UseAutomaticScale,
            .ManualScale = cfg.ManualScale,
            .IncludeIsometric = cfg.IncludeIsometric,
            .IncludeProjectedViews = cfg.IncludeProjectedViews,
            .IncludeFlatInDraftWhenPsm = cfg.IncludeFlatInDraftWhenPsm,
            .EnableAutoDimensioning = cfg.EnableAutoDimensioning,
            .EnableDrawingViewDimensioningLab = cfg.EnableDrawingViewDimensioningLab,
            .RunDropViewsTo2DModelLab = cfg.RunDropViewsTo2DModelLab,
            .RunDropCreatedSheetsDimensionLab = cfg.RunDropCreatedSheetsDimensionLab,
            .DropCreatedSheetsDimensionLabDebugSave = cfg.DropCreatedSheetsDimensionLabDebugSave,
            .RunDVGeometryDimensionPlacementLab = cfg.RunDVGeometryDimensionPlacementLab,
            .EnableDimLabInteractivePause = cfg.EnableDimLabInteractivePause,
            .DimLabMode = CInt(cfg.DimLabMode),
            .EnableDimLabVisibleProbe = cfg.EnableDimLabVisibleProbe,
            .EnableDimLabAlternativePlacement = cfg.EnableDimLabAlternativePlacement,
            .EnableDimLabHorizontalControlInVerticalOnly = cfg.EnableDimLabHorizontalControlInVerticalOnly,
            .DimLabKeepFailedDimensions = cfg.DimLabKeepFailedDimensions,
            .DimLabCleanPreviousLabDimensions = cfg.DimLabCleanPreviousLabDimensions,
            .EnablePmiRetrievalProbe = cfg.EnablePmiRetrievalProbe,
            .ExperimentalCreatePMIModelViewIfMissing = cfg.ExperimentalCreatePMIModelViewIfMissing,
            .ExperimentalDraftGeometryDiagnostics = cfg.ExperimentalDraftGeometryDiagnostics,
            .UseBestBaseViewLogic = cfg.UseBestBaseViewLogic,
            .ClientName = cfg.ClientName,
            .ProjectName = cfg.ProjectName,
            .DrawingTitle = cfg.DrawingTitle,
            .TitleSourceMode = cfg.TitleSourceMode,
            .Material = cfg.Material,
            .Thickness = cfg.Thickness,
            .Pedido = cfg.Pedido,
            .AuthorName = cfg.AuthorName,
            .Weight = cfg.Weight,
            .Equipment = cfg.Equipment,
            .DrawingNumber = cfg.DrawingNumber,
            .Revision = cfg.Revision,
            .Notes = cfg.Notes,
            .StrictMetadataValidation = If(chkStrictMetadata Is Nothing, False, chkStrictMetadata.Checked)
        }
    End Function

    Private Sub ApplySettingsToUi(settings As PersistedAppSettings)
        If settings Is Nothing Then Return

        txtInputFile.Text = settings.LastInputFile
        txtOutputFolder.Text = settings.LastOutputFolder
        txtTemplateA4.Text = settings.TemplateA4
        txtTemplateA3.Text = settings.TemplateA3
        txtTemplateA2.Text = settings.TemplateA2
        txtTemplateDxf.Text = settings.TemplateDxf

        chkCreateDft.Checked = settings.CreateDraft
        chkCreatePdf.Checked = settings.CreatePdf
        chkCreateDxfDraft.Checked = settings.CreateDxfFromDraft
        chkCreateFlatDxf.Checked = settings.CreateFlatDxf
        chkOpenOutput.Checked = settings.OpenOutputFolderWhenDone
        chkOverwrite.Checked = settings.OverwriteExisting
        chkUniqueComponents.Checked = settings.ProcessRepeatedComponentsOnce
        chkDetailedLog.Checked = settings.DetailedLog
        chkDebugTemplates.Checked = settings.DebugTemplatesInspection
        chkKeepSolidEdgeVisible.Checked = settings.KeepSolidEdgeVisible
        chkInsertProperties.Checked = settings.InsertPropertiesInTitleBlock
        If settings.TitleBlockPropertySourceMode <> TitleBlockPropertySource.FromModelLink AndAlso
           settings.TitleBlockPropertySourceMode <> TitleBlockPropertySource.FromDraft Then
            settings.TitleBlockPropertySourceMode = DEFAULT_TITLE_BLOCK_SOURCE
        End If
        SetSelectedTitleBlockPropertySource(settings.TitleBlockPropertySourceMode)
        Dim tsm As TitleSourceMode = settings.TitleSourceMode
        If tsm <> TitleSourceMode.Manual AndAlso tsm <> TitleSourceMode.AutoFromFileName Then tsm = TitleSourceMode.Manual
        SetSelectedTitleSourceMode(tsm)

        Select Case settings.PreferredFormat
            Case PreferredSheetFormat.A4 : cmbPreferredFormat.SelectedItem = "A4"
            Case PreferredSheetFormat.A3 : cmbPreferredFormat.SelectedItem = "A3"
            Case PreferredSheetFormat.A2 : cmbPreferredFormat.SelectedItem = "A2"
            Case Else : cmbPreferredFormat.SelectedItem = "Auto"
        End Select

        chkAutoScale.Checked = settings.UseAutomaticScale
        txtManualScale.Text = settings.ManualScale.ToString(Globalization.CultureInfo.InvariantCulture)
        chkIncludeIso.Checked = settings.IncludeIsometric
        chkIncludeProjected.Checked = settings.IncludeProjectedViews
        chkIncludeFlatInDraft.Checked = settings.IncludeFlatInDraftWhenPsm
        chkAutoDimensioning.Checked = settings.EnableAutoDimensioning
        chkDrawingViewDimensioningLab.Checked = settings.EnableDrawingViewDimensioningLab
        _runDropViewsTo2DModelLab = settings.RunDropViewsTo2DModelLab
        _runDropCreatedSheetsDimensionLab = settings.RunDropCreatedSheetsDimensionLab
        _dropCreatedSheetsLabDebugSave = settings.DropCreatedSheetsDimensionLabDebugSave
        _runDVGeometryDimensionPlacementLab = settings.RunDVGeometryDimensionPlacementLab
        If chkDimLabInteractivePause IsNot Nothing Then chkDimLabInteractivePause.Checked = settings.EnableDimLabInteractivePause
        If cmbDimLabMode IsNot Nothing Then
            Dim imx = settings.DimLabMode
            If imx < 0 OrElse imx > 5 Then imx = CInt(DimLabMode.Full)
            cmbDimLabMode.SelectedIndex = imx
        End If
        If chkDimLabVisibleProbe IsNot Nothing Then chkDimLabVisibleProbe.Checked = settings.EnableDimLabVisibleProbe
        If chkDimLabAlternativePlacement IsNot Nothing Then chkDimLabAlternativePlacement.Checked = settings.EnableDimLabAlternativePlacement
        DimensionInsertionConfig.EnableDrawingViewDimensioningLab = settings.EnableDrawingViewDimensioningLab
        RemoveLabControlsFromUi()
        chkPmiRetrievalProbe.Checked = False
        chkExperimentalPmiModelView.Checked = False
        chkExperimentalDraftGeometryDiagnostics.Checked = False
        chkUseBestBase.Checked = settings.UseBestBaseViewLogic

        If chkStrictMetadata IsNot Nothing Then chkStrictMetadata.Checked = settings.StrictMetadataValidation

        If settings.WindowWidth > 200 AndAlso settings.WindowHeight > 200 Then
            Me.StartPosition = FormStartPosition.Manual
            Me.Left = settings.WindowLeft
            Me.Top = settings.WindowTop
            Me.Width = settings.WindowWidth
            Me.Height = settings.WindowHeight
            If settings.WindowStateValue >= 0 AndAlso settings.WindowStateValue <= 2 Then
                Me.WindowState = CType(settings.WindowStateValue, FormWindowState)
            End If
        End If
    End Sub

    Private Sub ApplyDefaultTemplatesIfEmpty()
        If String.IsNullOrWhiteSpace(txtTemplateA4.Text) Then txtTemplateA4.Text = Path.Combine(DefaultTemplateFolder, "A4 Plantilla.dft")
        If String.IsNullOrWhiteSpace(txtTemplateA3.Text) Then txtTemplateA3.Text = Path.Combine(DefaultTemplateFolder, "A3 Plantilla.dft")
        If String.IsNullOrWhiteSpace(txtTemplateA2.Text) Then txtTemplateA2.Text = Path.Combine(DefaultTemplateFolder, "A2 Plantilla.dft")
        If String.IsNullOrWhiteSpace(txtTemplateDxf.Text) Then txtTemplateDxf.Text = Path.Combine(DefaultTemplateFolder, "DXF_LIMPIO.dft")
    End Sub

    Private Sub RestoreDefaultTemplatePaths()
        txtTemplateA4.Text = Path.Combine(DefaultTemplateFolder, "A4 Plantilla.dft")
        txtTemplateA3.Text = Path.Combine(DefaultTemplateFolder, "A3 Plantilla.dft")
        txtTemplateA2.Text = Path.Combine(DefaultTemplateFolder, "A2 Plantilla.dft")
        txtTemplateDxf.Text = Path.Combine(DefaultTemplateFolder, "DXF_LIMPIO.dft")
        _logger.Log("Templates restaurados a rutas por defecto.")
    End Sub

    Private Function ValidateTemplatePaths(showMessage As Boolean) As List(Of String)
        Dim warnings As New List(Of String)()
        Dim checks As (String, String)() = {
            ("A4", txtTemplateA4.Text.Trim()),
            ("A3", txtTemplateA3.Text.Trim()),
            ("A2", txtTemplateA2.Text.Trim()),
            ("DXF", txtTemplateDxf.Text.Trim())
        }
        For Each c In checks
            If Not String.IsNullOrWhiteSpace(c.Item2) AndAlso Not File.Exists(c.Item2) Then
                warnings.Add($"- Template {c.Item1} no existe: {c.Item2}")
            End If
        Next
        If warnings.Count > 0 Then
            For Each e In warnings
                _logger.Log("[VALIDACION] " & e)
            Next
            If showMessage Then
                MessageBox.Show(String.Join(Environment.NewLine, warnings), "Templates", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
        End If
        Return warnings
    End Function

    Private Sub LoadSourcePropertiesToUi()
        If String.IsNullOrWhiteSpace(txtInputFile.Text) OrElse Not File.Exists(txtInputFile.Text) Then
            _logger.Log("[INFO] No hay archivo válido para cargar propiedades.")
            Return
        End If

        Dim inputSnap As String = txtInputFile.Text.Trim()
        Dim myGen As Long = Interlocked.Increment(_metadataLoadGeneration)
        Dim showSe As Boolean = chkKeepSolidEdgeVisible.Checked

        Try
            _logger.Log("[UI][METADATA][CLEAR_BEFORE_LOAD] scope=ALL")
            ClearPlanMetadataUi()
            _loadedAsmComponentPath = ""
            SetBusy(True, "Leyendo metadatos del modelo...", True)

            StaComInvoker.Run(Function() As Tuple(Of Boolean, DrawingMetadataInput)
                                   Dim data As DrawingMetadataInput = Nothing
                                   Dim ok = DrawingMetadataService.TryLoadMetadataFromModelFile(inputSnap, showSe, _logger, data)
                                   Return Tuple.Create(ok, data)
                               End Function).
                ContinueWith(Sub(t As Task(Of Tuple(Of Boolean, DrawingMetadataInput)))
                                  BeginInvoke(New Action(Sub() FinishLoadSourceMetadataTask(t, inputSnap, myGen)))
                              End Sub)
        Catch ex As Exception
            _logger.LogException("LoadSourcePropertiesToUi", ex)
            SetBusy(False, "Preparado", False)
        End Try
    End Sub

    Private Sub FinishLoadSourceMetadataTask(t As Task(Of Tuple(Of Boolean, DrawingMetadataInput)), inputSnap As String, myGen As Long)
        Try
            SetBusy(False, "Preparado", False)

            If Interlocked.Read(_metadataLoadGeneration) <> myGen Then Return

            Dim pathNow As String = If(txtInputFile?.Text, "").Trim()
            If Not String.Equals(inputSnap, pathNow, StringComparison.OrdinalIgnoreCase) Then Return

            If t.IsFaulted Then
                Dim ex As Exception = t.Exception
                Dim agg As AggregateException = TryCast(ex, AggregateException)
                If agg IsNot Nothing Then ex = agg.GetBaseException()
                _logger.LogException("LoadSourcePropertiesToUi", ex)
                RefreshTraceabilityDataGridSafe()
                RefreshTitleModeUi()
                UpdateTitleBlockOriginHints()
                Return
            End If

            Dim tup = t.Result
            If Not tup.Item1 OrElse tup.Item2 Is Nothing Then
                _logger.Log("[WARN] No se pudieron leer metadatos del archivo de entrada.")
                RefreshTraceabilityDataGridSafe()
                RefreshTitleModeUi()
                UpdateTitleBlockOriginHints()
                Return
            End If

            DrawingMetadataService.ApplyToUi(Me, tup.Item2, applyCajetin:=True, applyPartList:=True)
            _logger.Log("[UI][METADATA] Metadatos aplicados desde modelo de entrada.")

            Dim kindNow As SourceFileKind = (New JobConfiguration With {.InputFile = pathNow}).DetectInputKind()
            If kindNow = SourceFileKind.AssemblyFile Then
                RefreshAsmComponentsIfNeeded(forceReload:=True)
            End If
            RefreshTraceabilityDataGridSafe()
            RefreshTitleModeUi()
            UpdateTitleBlockOriginHints()
        Catch ex As Exception
            _logger.LogException("FinishLoadSourceMetadataTask", ex)
        End Try
    End Sub

    Private Function ValidateMetadataBeforeGenerate(config As JobConfiguration) As Boolean
        If config Is Nothing Then Return True
        Dim strict As Boolean = config.StrictMetadataValidation
        Dim kind As SourceFileKind = config.DetectInputKind()
        If kind = SourceFileKind.PartFile OrElse kind = SourceFileKind.SheetMetalFile Then
            Dim uiData As DrawingMetadataInput = DrawingMetadataService.BuildFromUi(Me)
            Dim vr As MetadataValidationResult = DrawingMetadataService.ValidateMetadataForComponent(config.InputFile, uiData, strict, _logger)
            If vr.Outcome = MetadataValidationOutcome.RequiredMissing Then
                _logger.Log("[UI][GENERATE][BLOCKED_METADATA]")
                MessageBox.Show("Faltan metadatos obligatorios:" & Environment.NewLine & String.Join(Environment.NewLine, vr.MissingRequiredFields),
                                "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return False
            End If
            If vr.Outcome = MetadataValidationOutcome.WarningMissing Then
                _logger.Log("[UI][GENERATE][WARN_METADATA]")
                Dim r As DialogResult = MessageBox.Show(
                    "Faltan metadatos (validación no estricta). ¿Continuar?" & Environment.NewLine & String.Join(", ", vr.MissingWarningFields),
                    "Aviso", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
                Return r = DialogResult.Yes
            End If
            Return True
        End If

        If kind = SourceFileKind.AssemblyFile Then
            If config.SelectedComponentPaths Is Nothing OrElse config.SelectedComponentPaths.Count = 0 Then Return True
            Dim issues As New List(Of String)()
            For Each fp As String In config.SelectedComponentPaths
                Dim st As ComponentMetadataState = Nothing
                If Not _componentMetadataStates.TryGetValue(fp, st) OrElse st Is Nothing OrElse st.Status = ComponentMetadataStatus.Pending Then
                    issues.Add(Path.GetFileName(fp) & " — requiere «Datos»")
                    Continue For
                End If
                If st.Metadata Is Nothing Then
                    issues.Add(Path.GetFileName(fp) & " — sin metadatos")
                    Continue For
                End If
                Dim vr As MetadataValidationResult = DrawingMetadataService.ValidateMetadataForComponent(fp, st.Metadata, strict, _logger)
                If vr.Outcome <> MetadataValidationOutcome.Complete Then
                    issues.Add(Path.GetFileName(fp) & ": " & String.Join(", ", vr.MissingRequiredFields))
                End If
            Next
            If issues.Count = 0 Then Return True
            If strict Then
                _logger.Log("[UI][GENERATE][BLOCKED_METADATA]")
                MessageBox.Show(String.Join(Environment.NewLine, issues), "Validación ASM", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return False
            End If
            _logger.Log("[UI][GENERATE][WARN_METADATA]")
            Dim r2 As DialogResult = MessageBox.Show(
                "Hay componentes con metadatos incompletos. ¿Continuar?" & Environment.NewLine & String.Join(Environment.NewLine, issues),
                "Aviso", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
            Return r2 = DialogResult.Yes
        End If
        Return True
    End Function

    Private Sub SetupTraceabilityDataGridView()
        Try
            ' La tabla de inspección de trazabilidad ya no se muestra en el grupo «Datos de plano»
            ' (liberaba ~220 px vacíos y reducía el área útil del formulario).
            If grpTraceability.Controls.Contains(_dgvTraceability) Then
                grpTraceability.Controls.Remove(_dgvTraceability)
            End If
            _dgvTraceability.Visible = False
        Catch ex As Exception
            _logger.LogException("SetupTraceabilityDataGridView", ex)
        End Try
    End Sub

    Private Sub RefreshTraceabilityDataGridSafe()
        If _dgvTraceability Is Nothing OrElse Not _dgvTraceability.Visible Then Return
        Try
            If Me.IsHandleCreated AndAlso Me.InvokeRequired Then
                Me.BeginInvoke(New Action(AddressOf RefreshTraceabilityDataGridCore))
            Else
                RefreshTraceabilityDataGridCore()
            End If
        Catch
            Try
                RefreshTraceabilityDataGridCore()
            Catch
            End Try
        End Try
    End Sub

    Private Sub RefreshTraceabilityDataGridCore()
        Try
            If String.IsNullOrWhiteSpace(txtInputFile.Text) OrElse Not File.Exists(txtInputFile.Text) Then
                _dgvTraceability.DataSource = Nothing
                Return
            End If
            Dim cfgProbe As New JobConfiguration With {.InputFile = txtInputFile.Text, .OutputFolder = txtOutputFolder.Text.Trim()}
            Dim kind As SourceFileKind = cfgProbe.DetectInputKind()
            Dim cfgFull As JobConfiguration = BuildConfigurationFromUi()
            Dim dftPath As String = ResolveManualDftPath(cfgFull)

            Dim dt As DataTable = TraceabilityInspectionService.BuildTraceabilityDataTable(
                txtInputFile.Text, dftPath, kind, _logger, onlyCajetinFields:=True)
            _dgvTraceability.DataSource = dt

            For Each col As DataGridViewColumn In _dgvTraceability.Columns
                Select Case col.Name
                    Case "CampoLogico" : col.HeaderText = "Campo lógico"
                    Case "PropertySet" : col.HeaderText = "PropertySet"
                    Case "NombrePropiedad" : col.HeaderText = "Nombre propiedad"
                    Case "Alcance" : col.HeaderText = "Origen / Draft"
                    Case "ValorModelo" : col.HeaderText = "Valor modelo"
                    Case "ValorDraft" : col.HeaderText = "Valor draft"
                    Case "Escribible" : col.HeaderText = "Escribible"
                    Case "Sobrescribible" : col.HeaderText = "Sobrescribible"
                    Case "OrigenRecomendado" : col.HeaderText = "Origen recomendado"
                    Case "Observaciones" : col.HeaderText = "Observaciones"
                End Select
                If col.Name = "Observaciones" Then
                    col.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                End If
            Next
        Catch ex As Exception
            _logger.LogException("RefreshTraceabilityDataGridCore", ex)
        End Try
    End Sub

    Private Sub CopyTraceabilityGridToClipboard(sender As Object, e As EventArgs)
        Try
            Dim dt = TryCast(_dgvTraceability.DataSource, DataTable)
            If dt Is Nothing OrElse dt.Rows.Count = 0 Then Return
            Dim sb As New StringBuilder()
            For c As Integer = 0 To dt.Columns.Count - 1
                If c > 0 Then sb.Append(vbTab)
                sb.Append(dt.Columns(c).ColumnName)
            Next
            sb.AppendLine()
            For Each r As DataRow In dt.Rows
                For c As Integer = 0 To dt.Columns.Count - 1
                    If c > 0 Then sb.Append(vbTab)
                    sb.Append(If(r.IsNull(c), "", r(c).ToString().Replace(vbTab, " ")))
                Next
                sb.AppendLine()
            Next
            Clipboard.SetText(sb.ToString())
            _logger.Log("[INFO] Tabla de trazabilidad copiada al portapapeles.")
        Catch ex As Exception
            _logger.LogException("CopyTraceabilityGridToClipboard", ex)
        End Try
    End Sub

    Private Sub PopulateTitleBlockSourceOptions()
        cmbTitleBlockSource.Items.Clear()
        cmbTitleBlockSource.Items.Add(New TitleBlockSourceItem(TitleBlockPropertySource.FromModelLink, "Modelo enlazado (FromModelLink)"))
        cmbTitleBlockSource.Items.Add(New TitleBlockSourceItem(TitleBlockPropertySource.FromDraft, "Draft (FromDraft)"))
        cmbTitleBlockSource.SelectedIndex = 0
    End Sub

    Private Function GetSelectedTitleBlockPropertySource() As TitleBlockPropertySource
        Dim mode As TitleBlockPropertySource = DEFAULT_TITLE_BLOCK_SOURCE
        Try
            Dim item As TitleBlockSourceItem = TryCast(cmbTitleBlockSource.SelectedItem, TitleBlockSourceItem)
            If item IsNot Nothing Then
                mode = item.Mode
            End If
        Catch
        End Try
        If mode <> TitleBlockPropertySource.FromModelLink AndAlso mode <> TitleBlockPropertySource.FromDraft Then
            mode = DEFAULT_TITLE_BLOCK_SOURCE
        End If
        Return mode
    End Function

    Private Sub SetSelectedTitleBlockPropertySource(mode As TitleBlockPropertySource)
        If mode <> TitleBlockPropertySource.FromModelLink AndAlso mode <> TitleBlockPropertySource.FromDraft Then
            mode = DEFAULT_TITLE_BLOCK_SOURCE
        End If
        For i As Integer = 0 To cmbTitleBlockSource.Items.Count - 1
            Dim item As TitleBlockSourceItem = TryCast(cmbTitleBlockSource.Items(i), TitleBlockSourceItem)
            If item IsNot Nothing AndAlso item.Mode = mode Then
                cmbTitleBlockSource.SelectedIndex = i
                Exit Sub
            End If
        Next
        If cmbTitleBlockSource.Items.Count > 0 Then cmbTitleBlockSource.SelectedIndex = 0
    End Sub

    Private Class TitleBlockSourceItem
        Public ReadOnly Property Mode As TitleBlockPropertySource
        Public ReadOnly Property Display As String

        Public Sub New(mode As TitleBlockPropertySource, display As String)
            Me.Mode = mode
            Me.Display = display
        End Sub

        Public Overrides Function ToString() As String
            Return Display
        End Function
    End Class

    Private Sub PopulateTitleSourceCombo()
        cmbTitleSource.Items.Clear()
        cmbTitleSource.Items.Add(New TitleSourceItem(TitleSourceMode.Manual, "Manual (campo Título)"))
        cmbTitleSource.Items.Add(New TitleSourceItem(TitleSourceMode.AutoFromFileName, "Automático desde nombre de archivo"))
        If cmbTitleSource.Items.Count > 0 Then cmbTitleSource.SelectedIndex = 0
    End Sub

    Private Function GetSelectedTitleSourceMode() As TitleSourceMode
        Dim mode As TitleSourceMode = TitleSourceMode.Manual
        Try
            Dim it As TitleSourceItem = TryCast(cmbTitleSource.SelectedItem, TitleSourceItem)
            If it IsNot Nothing Then mode = it.Mode
        Catch
        End Try
        Return mode
    End Function

    Private Sub SetSelectedTitleSourceMode(mode As TitleSourceMode)
        For i As Integer = 0 To cmbTitleSource.Items.Count - 1
            Dim it As TitleSourceItem = TryCast(cmbTitleSource.Items(i), TitleSourceItem)
            If it IsNot Nothing AndAlso it.Mode = mode Then
                cmbTitleSource.SelectedIndex = i
                Exit Sub
            End If
        Next
        If cmbTitleSource.Items.Count > 0 Then cmbTitleSource.SelectedIndex = 0
    End Sub

    Private Sub RefreshTitleModeUi()
        Try
            Dim autoMode As Boolean = (GetSelectedTitleSourceMode() = TitleSourceMode.AutoFromFileName)
            txtTitle.ReadOnly = autoMode
            If autoMode AndAlso Not String.IsNullOrWhiteSpace(txtInputFile.Text) AndAlso File.Exists(txtInputFile.Text) Then
                txtTitle.Text = Path.GetFileNameWithoutExtension(txtInputFile.Text)
            End If
        Catch
        End Try
    End Sub

    ''' <summary>Indica de forma breve el origen previsto de cada campo del cajetín.</summary>
    Private Sub UpdateTitleBlockOriginHints()
        If _unifiedPlanMetadataDone Then Return
        Dim hasFile As Boolean = Not String.IsNullOrWhiteSpace(txtInputFile.Text) AndAlso File.Exists(txtInputFile.Text)
        Dim modelo As String = If(hasFile, "Documento modelo", "—")
        lblOriginTitle.Text = If(GetSelectedTitleSourceMode() = TitleSourceMode.AutoFromFileName,
            "Autocalculado (nombre de archivo)", "Usuario / " & modelo)
        lblOriginProject.Text = modelo & " / usuario"
        lblOriginMaterial.Text = modelo
        lblOriginClient.Text = modelo & " / usuario"
        lblOriginDocNum.Text = modelo
        lblOriginRevision.Text = modelo
        lblOriginAuthor.Text = If(String.IsNullOrWhiteSpace(txtAuthor.Text), "Usuario Windows (si vacío al generar)", "Usuario")
        lblOriginThickness.Text = modelo & " (calibre) / usuario"
        lblOriginPedido.Text = If(String.IsNullOrWhiteSpace(txtPedido.Text), "(vacío si no se informa)", "Usuario")
    End Sub

    Private Sub cmbTitleSource_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbTitleSource.SelectedIndexChanged
        RefreshTitleModeUi()
        UpdateTitleBlockOriginHints()
    End Sub

    Private Sub TitleBlockOriginFields_TextChanged(sender As Object, e As EventArgs) Handles txtAuthor.TextChanged, txtPedido.TextChanged, txtThickness.TextChanged
        UpdateTitleBlockOriginHints()
    End Sub

    Private Class TitleSourceItem
        Public ReadOnly Property Mode As TitleSourceMode
        Public ReadOnly Property Display As String

        Public Sub New(mode As TitleSourceMode, display As String)
            Me.Mode = mode
            Me.Display = display
        End Sub

        Public Overrides Function ToString() As String
            Return Display
        End Function
    End Class
End Class
