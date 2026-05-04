Imports System.Drawing
Imports System.Windows.Forms

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class MainForm
    Inherits Form

    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    Private components As System.ComponentModel.IContainer

    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.mainLayout = New TableLayoutPanel()
        Me.pnlHeader = New Panel()
        Me.lblTitle = New Label()
        Me.lblSubTitle = New Label()
        Me.bodyHost = New TableLayoutPanel()
        Me.threeColumnLayout = New TableLayoutPanel()
        Me.pnlGenerateBar = New Panel()
        Me.rightLayout = New TableLayoutPanel()
        Me.tblRightLogProgress = New TableLayoutPanel()
        Me.topInputOptionsLayout = New TableLayoutPanel()
        Me.grpInput = New GroupBox()
        Me.grpAsmComponents = New GroupBox()
        Me.grpTemplates = New GroupBox()
        Me.grpTraceability = New GroupBox()
        Me.grpGeneration = New GroupBox()
        Me.grpAdvanced = New GroupBox()
        Me.pnlHiddenLegacyGeneration = New Panel()
        Me.pnlHiddenProcessing = New Panel()
        Me.grpProgress = New GroupBox()
        Me.grpLog = New GroupBox()

        Me.tblInput = New TableLayoutPanel()
        Me.lblInput = New Label()
        Me.txtInputFile = New TextBox()
        Me.btnBrowseInput = New Button()
        Me.lblDetectedType = New Label()
        Me.lblDetectedTypeValue = New Label()

        Me.tblTemplates = New TableLayoutPanel()
        Me.lblTplA4 = New Label() : Me.txtTemplateA4 = New TextBox() : Me.btnBrowseA4 = New Button()
        Me.lblTplA3 = New Label() : Me.txtTemplateA3 = New TextBox() : Me.btnBrowseA3 = New Button()
        Me.lblTplA2 = New Label() : Me.txtTemplateA2 = New TextBox() : Me.btnBrowseA2 = New Button()
        Me.lblTplDxf = New Label() : Me.txtTemplateDxf = New TextBox() : Me.btnBrowseDxf = New Button()

        Me.tblTrace = New TableLayoutPanel()
        Me.tblAsmComponents = New TableLayoutPanel()
        Me.dgvAsmComponents = New DataGridView()
        Me.flowAsmButtons = New FlowLayoutPanel()
        Me.btnLoadAsmComponents = New Button()
        Me.btnSelectAllComponents = New Button()
        Me.btnSelectNoneComponents = New Button()
        Me.lblAsmComponentHint = New Label()
        Me.lblClient = New Label() : Me.txtClient = New TextBox()
        Me.lblProject = New Label() : Me.txtProject = New TextBox()
        Me.lblDrawingTitle = New Label() : Me.txtTitle = New TextBox()
        Me.lblMaterial = New Label() : Me.txtMaterial = New TextBox()
        Me.lblDrawingNumber = New Label() : Me.txtDrawingNumber = New TextBox()
        Me.lblRevision = New Label() : Me.txtRevision = New TextBox()
        Me.lblAuthor = New Label() : Me.txtAuthor = New TextBox()
        Me.lblThickness = New Label() : Me.txtThickness = New TextBox()
        Me.lblPedido = New Label() : Me.txtPedido = New TextBox()
        Me.cmbTitleSource = New ComboBox()
        Me.lblOriginTitle = New Label()
        Me.lblOriginProject = New Label()
        Me.lblOriginMaterial = New Label()
        Me.lblOriginClient = New Label()
        Me.lblOriginDocNum = New Label()
        Me.lblOriginRevision = New Label()
        Me.lblOriginAuthor = New Label()
        Me.lblOriginThickness = New Label()
        Me.lblOriginPedido = New Label()
        Me.lblNotes = New Label() : Me.txtNotes = New TextBox()
        Me.flowTraceButtons = New FlowLayoutPanel()
        Me.btnApplyTraceability = New Button()

        Me.flowGeneration = New FlowLayoutPanel()
        Me.lblTitleBlockSource = New Label()
        Me.cmbTitleBlockSource = New ComboBox()
        Me.chkCreateDft = New CheckBox()
        Me.chkCreatePdf = New CheckBox()
        Me.chkCreateDxfDraft = New CheckBox()
        Me.chkAutoDimensioning = New CheckBox()
        Me.chkUnitHorizontalExteriorTest = New CheckBox()
        Me.chkPmiRetrievalProbe = New CheckBox()
        Me.chkExperimentalPmiModelView = New CheckBox()
        Me.chkCreateFlatDxf = New CheckBox()
        Me.chkOpenOutput = New CheckBox()
        Me.chkOverwrite = New CheckBox()
        Me.chkUniqueComponents = New CheckBox()
        Me.chkDetailedLog = New CheckBox()
        Me.chkKeepSolidEdgeVisible = New CheckBox()
        Me.chkInsertProperties = New CheckBox()
        Me.chkDebugTemplates = New CheckBox()
        Me.chkExperimentalDraftGeometryDiagnostics = New CheckBox()

        Me.tblAdvanced = New TableLayoutPanel()
        Me.lblOutput = New Label()
        Me.txtOutputFolder = New TextBox()
        Me.btnBrowseOut = New Button()
        Me.lblPreferredFormat = New Label()
        Me.cmbPreferredFormat = New ComboBox()
        Me.chkAutoScale = New CheckBox()
        Me.txtManualScale = New TextBox()
        Me.chkIncludeIso = New CheckBox()
        Me.chkIncludeProjected = New CheckBox()
        Me.chkIncludeFlatInDraft = New CheckBox()
        Me.chkUseBestBase = New CheckBox()

        Me.btnGenerate = New Button()
        Me.btnClear = New Button()
        Me.btnOpenOutput = New Button()
        Me.btnSaveConfig = New Button()
        Me.btnLoadConfig = New Button()
        Me.btnReloadSourceProps = New Button()
        Me.btnRestoreDefaultTemplates = New Button()

        Me.tblProgress = New TableLayoutPanel()
        Me.lblCurrentTime = New Label()
        Me.lblCurrentTimeValue = New Label()
        Me.lblPieceTime = New Label()
        Me.lblPieceTimeValue = New Label()
        Me.lblTotalTime = New Label()
        Me.lblTotalTimeValue = New Label()
        Me.lblStatus = New Label()
        Me.lblStatusValue = New Label()
        Me.pnlProgressBars = New Panel()
        Me.lblProgressPiece = New Label()
        Me.progressBar = New ProgressBar()
        Me.lblProgressAsm = New Label()
        Me.progressBarAsm = New ProgressBar()

        Me.tblLog = New TableLayoutPanel()
        Me.txtLog = New TextBox()
        Me.flowLogButtons = New FlowLayoutPanel()
        Me.btnSaveLog = New Button()
        Me.btnClearLog = New Button()

        CType(Me.dgvAsmComponents, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()

        Me.mainLayout.Dock = DockStyle.Fill
        Me.mainLayout.ColumnCount = 1
        Me.mainLayout.ColumnStyles.Clear()
        Me.mainLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0!))
        Me.mainLayout.RowCount = 3
        Me.mainLayout.RowStyles.Clear()
        Me.mainLayout.RowStyles.Add(New RowStyle(SizeType.Absolute, 70.0!))
        Me.mainLayout.RowStyles.Add(New RowStyle(SizeType.Absolute, 160.0!))
        Me.mainLayout.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0!))

        Me.pnlHeader.Dock = DockStyle.Fill
        Me.pnlHeader.BackColor = Color.FromArgb(33, 56, 86)
        Me.lblTitle.AutoSize = True
        Me.lblTitle.Text = "Solid Edge - Generador Automatico de DFT / PDF / DXF"
        Me.lblTitle.Font = New Font("Segoe UI Semibold", 14.0!, FontStyle.Bold)
        Me.lblTitle.ForeColor = Color.White
        Me.lblTitle.Location = New Point(12, 10)
        Me.lblSubTitle.AutoSize = True
        Me.lblSubTitle.Text = "ASM / PAR / PSM"
        Me.lblSubTitle.ForeColor = Color.Gainsboro
        Me.lblSubTitle.Location = New Point(14, 42)
        Me.pnlHeader.Controls.Add(Me.lblTitle)
        Me.pnlHeader.Controls.Add(Me.lblSubTitle)

        Me.bodyHost.Dock = DockStyle.Fill
        Me.bodyHost.ColumnCount = 1
        Me.bodyHost.ColumnStyles.Clear()
        Me.bodyHost.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0!))
        Me.bodyHost.RowCount = 2
        Me.bodyHost.RowStyles.Clear()
        Me.bodyHost.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0!))
        Me.bodyHost.RowStyles.Add(New RowStyle(SizeType.Absolute, 58.0!))
        Me.bodyHost.Controls.Add(Me.threeColumnLayout, 0, 0)
        Me.bodyHost.Controls.Add(Me.pnlGenerateBar, 0, 1)

        Me.threeColumnLayout.Dock = DockStyle.Fill
        Me.threeColumnLayout.ColumnCount = 3
        Me.threeColumnLayout.ColumnStyles.Clear()
        Me.threeColumnLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 40.0!))
        Me.threeColumnLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 30.0!))
        Me.threeColumnLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 30.0!))
        Me.threeColumnLayout.RowCount = 1
        Me.threeColumnLayout.RowStyles.Clear()
        Me.threeColumnLayout.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0!))
        Me.threeColumnLayout.Controls.Add(Me.grpAsmComponents, 0, 0)
        Me.threeColumnLayout.Controls.Add(Me.grpTraceability, 1, 0)
        Me.threeColumnLayout.Controls.Add(Me.rightLayout, 2, 0)

        Me.pnlGenerateBar.Dock = DockStyle.Fill
        Me.pnlGenerateBar.Padding = New Padding(8, 6, 8, 6)
        Me.pnlGenerateBar.Controls.Add(Me.btnGenerate)

        Me.rightLayout.Dock = DockStyle.Fill
        Me.rightLayout.ColumnCount = 1
        Me.rightLayout.ColumnStyles.Clear()
        Me.rightLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0!))
        Me.rightLayout.RowCount = 1
        Me.rightLayout.RowStyles.Clear()
        Me.rightLayout.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0!))

        Me.tblRightLogProgress.Dock = DockStyle.Fill
        Me.tblRightLogProgress.ColumnCount = 1
        Me.tblRightLogProgress.ColumnStyles.Clear()
        Me.tblRightLogProgress.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0!))
        Me.tblRightLogProgress.RowCount = 2
        Me.tblRightLogProgress.RowStyles.Clear()
        Me.tblRightLogProgress.RowStyles.Add(New RowStyle(SizeType.Percent, 65.0!))
        Me.tblRightLogProgress.RowStyles.Add(New RowStyle(SizeType.Percent, 35.0!))
        Me.tblRightLogProgress.Margin = New Padding(0)
        Me.tblRightLogProgress.Padding = New Padding(0)
        Me.tblRightLogProgress.Controls.Add(Me.grpLog, 0, 0)
        Me.tblRightLogProgress.Controls.Add(Me.grpProgress, 0, 1)

        For Each gb In New GroupBox() {Me.grpInput, Me.grpAsmComponents, Me.grpTemplates, Me.grpTraceability, Me.grpGeneration, Me.grpAdvanced, Me.grpProgress, Me.grpLog}
            gb.Dock = DockStyle.Fill
            gb.Font = New Font("Segoe UI", 9.0!)
            gb.Margin = New Padding(8)
        Next
        Me.grpGeneration.Margin = New Padding(4, 6, 2, 6)
        Me.grpTemplates.Margin = New Padding(2, 6, 4, 6)
        Me.grpLog.Margin = New Padding(8, 2, 8, 4)
        Me.grpProgress.Margin = New Padding(8, 2, 8, 8)
        Me.grpInput.Text = "Archivo de entrada"
        Me.grpAsmComponents.Text = "Componentes detectados del ASM (marcar los que deseas procesar)"
        Me.grpTemplates.Text = "Templates"
        Me.grpTraceability.Text = "Metadatos (plano y pieza)"
        Me.grpGeneration.Text = "Opciones de generación"
        Me.grpAdvanced.Text = "Opciones de procesado"
        Me.grpProgress.Text = "Progreso y estado"
        Me.grpLog.Text = "Log"

        Me.rightLayout.Controls.Add(Me.tblRightLogProgress, 0, 0)

        Me.topInputOptionsLayout.Dock = DockStyle.Fill
        Me.topInputOptionsLayout.ColumnCount = 3
        Me.topInputOptionsLayout.ColumnStyles.Clear()
        Me.topInputOptionsLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 78.0!))
        Me.topInputOptionsLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 11.0!))
        Me.topInputOptionsLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 11.0!))
        Me.topInputOptionsLayout.RowCount = 1
        Me.topInputOptionsLayout.RowStyles.Clear()
        Me.topInputOptionsLayout.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0!))
        Me.topInputOptionsLayout.Margin = New Padding(0)
        Me.topInputOptionsLayout.Controls.Add(Me.grpInput, 0, 0)
        Me.topInputOptionsLayout.Controls.Add(Me.grpGeneration, 1, 0)
        Me.topInputOptionsLayout.Controls.Add(Me.grpTemplates, 2, 0)

        Me.tblInput.Dock = DockStyle.Fill
        Me.tblInput.ColumnCount = 3
        Me.tblInput.ColumnStyles.Clear()
        Me.tblInput.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 90.0!))
        Me.tblInput.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0!))
        Me.tblInput.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 100.0!))
        Me.tblInput.RowCount = 3
        Me.tblInput.RowStyles.Clear()
        Me.tblInput.RowStyles.Add(New RowStyle(SizeType.Absolute, 32.0!))
        Me.tblInput.RowStyles.Add(New RowStyle(SizeType.Absolute, 32.0!))
        Me.tblInput.RowStyles.Add(New RowStyle(SizeType.Absolute, 32.0!))
        Me.lblInput.Text = "Archivo origen"
        Me.lblDetectedType.Text = "Tipo detectado"
        Me.lblDetectedTypeValue.Text = "-"
        Me.txtInputFile.Dock = DockStyle.Fill : Me.txtInputFile.ReadOnly = True
        Me.btnBrowseInput.Text = "Examinar..." : Me.btnBrowseInput.Dock = DockStyle.Fill
        Me.tblInput.Controls.Add(Me.lblInput, 0, 0)
        Me.tblInput.Controls.Add(Me.txtInputFile, 1, 0)
        Me.tblInput.Controls.Add(Me.btnBrowseInput, 2, 0)
        Me.tblInput.Controls.Add(Me.lblDetectedType, 0, 1)
        Me.tblInput.Controls.Add(Me.lblDetectedTypeValue, 1, 1)
        Me.lblOutput.Text = "Carpeta salida"
        Me.txtOutputFolder.Dock = DockStyle.Fill
        Me.btnBrowseOut.Text = "Examinar..."
        Me.tblInput.Controls.Add(Me.lblOutput, 0, 2)
        Me.tblInput.Controls.Add(Me.txtOutputFolder, 1, 2)
        Me.tblInput.Controls.Add(Me.btnBrowseOut, 2, 2)
        Me.grpInput.Controls.Add(Me.tblInput)

        Me.tblAsmComponents.Dock = DockStyle.Fill
        Me.tblAsmComponents.ColumnCount = 1
        Me.tblAsmComponents.RowCount = 3
        Me.tblAsmComponents.RowStyles.Clear()
        Me.tblAsmComponents.RowStyles.Add(New RowStyle(SizeType.Absolute, 30.0!))
        Me.tblAsmComponents.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0!))
        Me.tblAsmComponents.RowStyles.Add(New RowStyle(SizeType.Absolute, 32.0!))
        Me.lblAsmComponentHint.Text = "Solo aplica para entrada ASM."
        Me.lblAsmComponentHint.Dock = DockStyle.Fill
        Me.dgvAsmComponents.Dock = DockStyle.Fill
        Me.dgvAsmComponents.AllowUserToAddRows = False
        Me.dgvAsmComponents.AllowUserToDeleteRows = False
        Me.dgvAsmComponents.AllowUserToResizeRows = True
        Me.dgvAsmComponents.MultiSelect = False
        Me.dgvAsmComponents.ReadOnly = False
        Me.dgvAsmComponents.RowHeadersVisible = False
        Me.dgvAsmComponents.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        Me.dgvAsmComponents.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        Me.dgvAsmComponents.BorderStyle = BorderStyle.None
        Me.flowAsmButtons.Dock = DockStyle.Fill
        Me.btnLoadAsmComponents.Text = "Recargar"
        Me.btnSelectAllComponents.Text = "Marcar"
        Me.btnSelectNoneComponents.Text = "Desmarcar"
        Me.flowAsmButtons.Controls.Add(Me.btnLoadAsmComponents)
        Me.flowAsmButtons.Controls.Add(Me.btnSelectAllComponents)
        Me.flowAsmButtons.Controls.Add(Me.btnSelectNoneComponents)
        Me.tblAsmComponents.Controls.Add(Me.lblAsmComponentHint, 0, 0)
        Me.tblAsmComponents.Controls.Add(Me.dgvAsmComponents, 0, 1)
        Me.tblAsmComponents.Controls.Add(Me.flowAsmButtons, 0, 2)
        Me.grpAsmComponents.Controls.Add(Me.tblAsmComponents)

        Me.tblTemplates.Dock = DockStyle.Fill
        Me.tblTemplates.ColumnCount = 3
        Me.tblTemplates.ColumnStyles.Clear()
        Me.tblTemplates.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 108.0!))
        Me.tblTemplates.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0!))
        Me.tblTemplates.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 82.0!))
        Me.tblTemplates.RowCount = 4
        Me.tblTemplates.RowStyles.Clear()
        For i As Integer = 1 To 4 : Me.tblTemplates.RowStyles.Add(New RowStyle(SizeType.Absolute, 32.0!)) : Next
        Me.lblTplA4.Text = "Template A4" : Me.btnBrowseA4.Text = "Examinar..."
        Me.lblTplA3.Text = "Template A3" : Me.btnBrowseA3.Text = "Examinar..."
        Me.lblTplA2.Text = "Template A2" : Me.btnBrowseA2.Text = "Examinar..."
        Me.lblTplDxf.Text = "Template DXF limpio" : Me.btnBrowseDxf.Text = "Examinar..."
        Me.btnBrowseA4.Width = 80 : Me.btnBrowseA4.Dock = DockStyle.Fill
        Me.btnBrowseA3.Width = 80 : Me.btnBrowseA3.Dock = DockStyle.Fill
        Me.btnBrowseA2.Width = 80 : Me.btnBrowseA2.Dock = DockStyle.Fill
        Me.btnBrowseDxf.Width = 80 : Me.btnBrowseDxf.Dock = DockStyle.Fill
        Me.txtTemplateA4.Dock = DockStyle.Fill : Me.txtTemplateA3.Dock = DockStyle.Fill
        Me.txtTemplateA2.Dock = DockStyle.Fill : Me.txtTemplateDxf.Dock = DockStyle.Fill
        Me.tblTemplates.Controls.Add(Me.lblTplA4, 0, 0) : Me.tblTemplates.Controls.Add(Me.txtTemplateA4, 1, 0) : Me.tblTemplates.Controls.Add(Me.btnBrowseA4, 2, 0)
        Me.tblTemplates.Controls.Add(Me.lblTplA3, 0, 1) : Me.tblTemplates.Controls.Add(Me.txtTemplateA3, 1, 1) : Me.tblTemplates.Controls.Add(Me.btnBrowseA3, 2, 1)
        Me.tblTemplates.Controls.Add(Me.lblTplA2, 0, 2) : Me.tblTemplates.Controls.Add(Me.txtTemplateA2, 1, 2) : Me.tblTemplates.Controls.Add(Me.btnBrowseA2, 2, 2)
        Me.tblTemplates.Controls.Add(Me.lblTplDxf, 0, 3) : Me.tblTemplates.Controls.Add(Me.txtTemplateDxf, 1, 3) : Me.tblTemplates.Controls.Add(Me.btnBrowseDxf, 2, 3)
        Me.grpTemplates.Controls.Add(Me.tblTemplates)

        Me.tblTrace.Dock = DockStyle.Fill
        Me.tblTrace.ColumnCount = 3
        Me.tblTrace.ColumnStyles.Clear()
        Me.tblTrace.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 132.0!))
        Me.tblTrace.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0!))
        Me.tblTrace.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 128.0!))
        Me.tblTrace.RowCount = 10
        Me.tblTrace.RowStyles.Clear()
        Me.tblTrace.RowStyles.Add(New RowStyle(SizeType.Absolute, 30.0!))
        Me.tblTrace.RowStyles.Add(New RowStyle(SizeType.Absolute, 28.0!))
        For i As Integer = 1 To 6
            Me.tblTrace.RowStyles.Add(New RowStyle(SizeType.Absolute, 30.0!))
        Next
        Me.tblTrace.RowStyles.Add(New RowStyle(SizeType.Absolute, 10.0!))
        Me.tblTrace.RowStyles.Add(New RowStyle(SizeType.Absolute, 38.0!))

        Me.lblDrawingTitle.Text = "Título / Denominación" : Me.lblDrawingTitle.AutoSize = True : Me.lblDrawingTitle.Dock = DockStyle.Fill
        Me.txtTitle.Dock = DockStyle.Fill
        Me.lblOriginTitle.AutoSize = False : Me.lblOriginTitle.Dock = DockStyle.Fill : Me.lblOriginTitle.TextAlign = ContentAlignment.MiddleLeft
        Me.lblOriginTitle.Font = New Font("Segoe UI", 8.0!, FontStyle.Italic)
        Me.tblTrace.Controls.Add(Me.lblDrawingTitle, 0, 0)
        Me.tblTrace.Controls.Add(Me.txtTitle, 1, 0)
        Me.tblTrace.Controls.Add(Me.lblOriginTitle, 2, 0)

        Me.cmbTitleSource.DropDownStyle = ComboBoxStyle.DropDownList
        Me.cmbTitleSource.Dock = DockStyle.Fill
        Me.tblTrace.Controls.Add(Me.cmbTitleSource, 1, 1)
        Me.tblTrace.SetColumnSpan(Me.cmbTitleSource, 2)

        Me.lblProject.Text = "Proyecto" : Me.lblProject.AutoSize = True : Me.lblProject.Dock = DockStyle.Fill
        Me.txtProject.Dock = DockStyle.Fill
        Me.lblOriginProject.AutoSize = False : Me.lblOriginProject.Dock = DockStyle.Fill : Me.lblOriginProject.TextAlign = ContentAlignment.MiddleLeft
        Me.lblOriginProject.Font = New Font("Segoe UI", 8.0!, FontStyle.Italic)
        Me.tblTrace.Controls.Add(Me.lblProject, 0, 2)
        Me.tblTrace.Controls.Add(Me.txtProject, 1, 2)
        Me.tblTrace.Controls.Add(Me.lblOriginProject, 2, 2)

        Me.lblMaterial.Text = "Material (PART_LIST)" : Me.lblMaterial.AutoSize = True : Me.lblMaterial.Dock = DockStyle.Fill
        Me.txtMaterial.Dock = DockStyle.Fill
        Me.lblOriginMaterial.AutoSize = False : Me.lblOriginMaterial.Dock = DockStyle.Fill : Me.lblOriginMaterial.TextAlign = ContentAlignment.MiddleLeft
        Me.lblOriginMaterial.Font = New Font("Segoe UI", 8.0!, FontStyle.Italic)

        Me.lblClient.Text = "Cliente / empresa" : Me.lblClient.AutoSize = True : Me.lblClient.Dock = DockStyle.Fill
        Me.txtClient.Dock = DockStyle.Fill
        Me.lblOriginClient.AutoSize = False : Me.lblOriginClient.Dock = DockStyle.Fill : Me.lblOriginClient.TextAlign = ContentAlignment.MiddleLeft
        Me.lblOriginClient.Font = New Font("Segoe UI", 8.0!, FontStyle.Italic)
        Me.tblTrace.Controls.Add(Me.lblClient, 0, 3)
        Me.tblTrace.Controls.Add(Me.txtClient, 1, 3)
        Me.tblTrace.Controls.Add(Me.lblOriginClient, 2, 3)

        Me.lblDrawingNumber.Text = "Plano" : Me.lblDrawingNumber.AutoSize = True : Me.lblDrawingNumber.Dock = DockStyle.Fill
        Me.txtDrawingNumber.Dock = DockStyle.Fill
        Me.lblOriginDocNum.AutoSize = False : Me.lblOriginDocNum.Dock = DockStyle.Fill : Me.lblOriginDocNum.TextAlign = ContentAlignment.MiddleLeft
        Me.lblOriginDocNum.Font = New Font("Segoe UI", 8.0!, FontStyle.Italic)
        Me.tblTrace.Controls.Add(Me.lblDrawingNumber, 0, 4)
        Me.tblTrace.Controls.Add(Me.txtDrawingNumber, 1, 4)
        Me.tblTrace.Controls.Add(Me.lblOriginDocNum, 2, 4)

        Me.lblRevision.Text = "Revisión" : Me.lblRevision.AutoSize = True : Me.lblRevision.Dock = DockStyle.Fill
        Me.txtRevision.Dock = DockStyle.Fill
        Me.lblOriginRevision.AutoSize = False : Me.lblOriginRevision.Dock = DockStyle.Fill : Me.lblOriginRevision.TextAlign = ContentAlignment.MiddleLeft
        Me.lblOriginRevision.Font = New Font("Segoe UI", 8.0!, FontStyle.Italic)
        Me.tblTrace.Controls.Add(Me.lblRevision, 0, 5)
        Me.tblTrace.Controls.Add(Me.txtRevision, 1, 5)
        Me.tblTrace.Controls.Add(Me.lblOriginRevision, 2, 5)

        Me.lblAuthor.Text = "Autor" : Me.lblAuthor.AutoSize = True : Me.lblAuthor.Dock = DockStyle.Fill
        Me.txtAuthor.Dock = DockStyle.Fill
        Me.lblOriginAuthor.AutoSize = False : Me.lblOriginAuthor.Dock = DockStyle.Fill : Me.lblOriginAuthor.TextAlign = ContentAlignment.MiddleLeft
        Me.lblOriginAuthor.Font = New Font("Segoe UI", 8.0!, FontStyle.Italic)
        Me.tblTrace.Controls.Add(Me.lblAuthor, 0, 6)
        Me.tblTrace.Controls.Add(Me.txtAuthor, 1, 6)
        Me.tblTrace.Controls.Add(Me.lblOriginAuthor, 2, 6)

        Me.lblThickness.Text = "Espesor (PART_LIST)" : Me.lblThickness.AutoSize = True : Me.lblThickness.Dock = DockStyle.Fill
        Me.txtThickness.Dock = DockStyle.Fill
        Me.lblOriginThickness.AutoSize = False : Me.lblOriginThickness.Dock = DockStyle.Fill : Me.lblOriginThickness.TextAlign = ContentAlignment.MiddleLeft
        Me.lblOriginThickness.Font = New Font("Segoe UI", 8.0!, FontStyle.Italic)

        Me.lblPedido.Text = "Pedido" : Me.lblPedido.AutoSize = True : Me.lblPedido.Dock = DockStyle.Fill
        Me.txtPedido.Dock = DockStyle.Fill
        Me.lblOriginPedido.AutoSize = False : Me.lblOriginPedido.Dock = DockStyle.Fill : Me.lblOriginPedido.TextAlign = ContentAlignment.MiddleLeft
        Me.lblOriginPedido.Font = New Font("Segoe UI", 8.0!, FontStyle.Italic)
        Me.tblTrace.Controls.Add(Me.lblPedido, 0, 7)
        Me.tblTrace.Controls.Add(Me.txtPedido, 1, 7)
        Me.tblTrace.Controls.Add(Me.lblOriginPedido, 2, 7)

        Me.flowTraceButtons.Dock = DockStyle.Fill
        Me.flowTraceButtons.FlowDirection = FlowDirection.RightToLeft
        Me.flowTraceButtons.WrapContents = False
        Me.flowTraceButtons.Padding = New Padding(0, 4, 0, 0)
        Me.btnApplyTraceability.Text = "Aplicar propiedades al DFT"
        Me.btnApplyTraceability.Size = New Size(220, 30)
        Me.flowTraceButtons.Controls.Add(Me.btnApplyTraceability)
        Me.tblTrace.Controls.Add(Me.flowTraceButtons, 0, 9)
        Me.tblTrace.SetColumnSpan(Me.flowTraceButtons, 3)

        Me.grpTraceability.Controls.Add(Me.tblTrace)

        Me.flowGeneration.Dock = DockStyle.Fill : Me.flowGeneration.Padding = New Padding(5)
        Me.grpGeneration.Padding = New Padding(5, 8, 5, 5)
        Me.grpTemplates.Padding = New Padding(5, 8, 5, 5)
        Me.flowGeneration.FlowDirection = FlowDirection.TopDown
        Me.flowGeneration.WrapContents = False
        Me.flowGeneration.AutoScroll = True
        Me.lblTitleBlockSource.Text = "Origen propiedades cajetín"
        Me.lblTitleBlockSource.AutoSize = True
        Me.cmbTitleBlockSource.DropDownStyle = ComboBoxStyle.DropDownList
        Me.cmbTitleBlockSource.Width = 220
        Me.flowGeneration.Controls.Add(Me.lblTitleBlockSource)
        Me.flowGeneration.Controls.Add(Me.cmbTitleBlockSource)
        Me.flowGeneration.Controls.Add(Me.chkUnitHorizontalExteriorTest)
        Dim gc As CheckBox() = {Me.chkCreateDft, Me.chkCreatePdf, Me.chkCreateDxfDraft, Me.chkAutoDimensioning, Me.chkPmiRetrievalProbe, Me.chkExperimentalPmiModelView, Me.chkCreateFlatDxf, Me.chkOpenOutput, Me.chkOverwrite, Me.chkUniqueComponents, Me.chkDetailedLog, Me.chkDebugTemplates, Me.chkExperimentalDraftGeometryDiagnostics, Me.chkKeepSolidEdgeVisible, Me.chkInsertProperties}
        Dim gt As String() = {"Crear DFT", "Crear PDF", "Crear DXF del DFT", "Generar acotado automático", "Probar recuperación PMI en Draft (experimental)", "Experimental: crear PMIModelView si falta (sin guardar modelo)", "Crear DXF de chapa desarrollada", "Abrir carpeta de salida al finalizar", "Sobrescribir archivos existentes", "Procesar componentes repetidos solo una vez", "Registrar log detallado", "Diagnóstico: inspeccionar plantillas DFT (PropertySets, cajetín)", "Diagnóstico geométrico Draft (2D, vistas, motor acotado)", "Mantener visible Solid Edge durante el proceso", "Insertar propiedades en el cajetin del draft, si procede"}
        For i As Integer = 0 To gc.Length - 1
            gc(i).Text = gt(i) : gc(i).AutoSize = True
            If i <> 4 AndAlso i <> 5 AndAlso i <> 8 AndAlso i <> 11 AndAlso i <> 12 AndAlso i <> 14 Then gc(i).Checked = True
            Me.pnlHiddenLegacyGeneration.Controls.Add(gc(i))
        Next
        Me.grpGeneration.Controls.Add(Me.flowGeneration)
        Me.pnlHiddenLegacyGeneration.Visible = False
        Me.pnlHiddenLegacyGeneration.Size = New Size(1, 1)
        Me.pnlHiddenLegacyGeneration.Location = New Point(-8000, -8000)
        Me.pnlHiddenLegacyGeneration.TabStop = False

        Me.tblAdvanced.Dock = DockStyle.Fill
        Me.tblAdvanced.ColumnCount = 3
        Me.tblAdvanced.ColumnStyles.Clear()
        Me.tblAdvanced.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 160.0!))
        Me.tblAdvanced.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0!))
        Me.tblAdvanced.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 100.0!))
        Me.tblAdvanced.RowCount = 6
        Me.tblAdvanced.RowStyles.Clear()
        For i As Integer = 1 To 6 : Me.tblAdvanced.RowStyles.Add(New RowStyle(SizeType.Absolute, 30.0!)) : Next
        Me.lblPreferredFormat.Text = "Formato preferido"
        Me.chkAutoScale.Text = "Escalado automatico" : Me.chkAutoScale.Checked = True
        Me.txtManualScale.Text = "1.0"
        Me.chkIncludeIso.Text = "Incluir vista isometrica" : Me.chkIncludeIso.Checked = True
        Me.chkIncludeProjected.Text = "Incluir vistas proyectadas" : Me.chkIncludeProjected.Checked = True
        Me.chkIncludeFlatInDraft.Text = "Incluir chapa desarrollada en draft (.psm)" : Me.chkIncludeFlatInDraft.Checked = True
        Me.chkUseBestBase.Text = "Usar logica actual de mejor vista base" : Me.chkUseBestBase.Checked = True
        Me.cmbPreferredFormat.DropDownStyle = ComboBoxStyle.DropDownList
        Me.tblAdvanced.Controls.Add(Me.lblPreferredFormat, 0, 0) : Me.tblAdvanced.Controls.Add(Me.cmbPreferredFormat, 1, 0)
        Me.tblAdvanced.Controls.Add(Me.chkAutoScale, 0, 1) : Me.tblAdvanced.Controls.Add(Me.txtManualScale, 1, 1)
        Me.tblAdvanced.Controls.Add(Me.chkIncludeIso, 0, 2) : Me.tblAdvanced.Controls.Add(Me.chkIncludeProjected, 1, 2)
        Me.tblAdvanced.Controls.Add(Me.chkIncludeFlatInDraft, 0, 3)
        Me.tblAdvanced.SetColumnSpan(Me.chkIncludeFlatInDraft, 3)
        Me.tblAdvanced.Controls.Add(Me.chkUseBestBase, 0, 4)
        Me.lblNotes.Text = "Observaciones (asunto PDF opc.)"
        Me.txtNotes.Dock = DockStyle.Fill
        Me.tblAdvanced.Controls.Add(Me.lblNotes, 0, 5)
        Me.tblAdvanced.Controls.Add(Me.txtNotes, 1, 5)
        Me.tblAdvanced.SetColumnSpan(Me.txtNotes, 2)
        Me.pnlHiddenProcessing.Controls.Add(Me.tblAdvanced)
        Me.pnlHiddenProcessing.Visible = False
        Me.pnlHiddenProcessing.Size = New Size(1, 1)
        Me.pnlHiddenProcessing.Location = New Point(-8000, -8000)
        Me.pnlHiddenProcessing.TabStop = False
        Me.grpAdvanced.Visible = False
        Me.grpAdvanced.Size = New Size(1, 1)

        Me.btnGenerate.Text = "GENERAR"
        Me.btnGenerate.Dock = DockStyle.Fill
        Me.btnGenerate.MinimumSize = New Size(280, 50)
        Me.btnGenerate.Height = 50
        Me.btnGenerate.Font = New Font("Segoe UI", 12.0!, FontStyle.Bold)
        Me.btnGenerate.BackColor = Color.FromArgb(39, 116, 57) : Me.btnGenerate.ForeColor = Color.White : Me.btnGenerate.FlatStyle = FlatStyle.Flat
        Me.btnGenerate.Margin = New Padding(0)
        Me.btnClear.Visible = False
        Me.btnOpenOutput.Visible = False
        Me.btnSaveConfig.Visible = False
        Me.btnLoadConfig.Visible = False
        Me.btnReloadSourceProps.Visible = False
        Me.btnRestoreDefaultTemplates.Visible = False
        Me.tblProgress.Dock = DockStyle.Fill
        Me.tblProgress.ColumnCount = 3
        Me.tblProgress.ColumnStyles.Clear()
        Me.tblProgress.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 110.0!))
        Me.tblProgress.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 90.0!))
        Me.tblProgress.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0!))
        Me.tblProgress.RowCount = 4
        Me.tblProgress.RowStyles.Clear()
        Me.tblProgress.RowStyles.Add(New RowStyle(SizeType.Absolute, 22.0!))
        Me.tblProgress.RowStyles.Add(New RowStyle(SizeType.Absolute, 22.0!))
        Me.tblProgress.RowStyles.Add(New RowStyle(SizeType.Absolute, 22.0!))
        Me.tblProgress.RowStyles.Add(New RowStyle(SizeType.Absolute, 22.0!))
        Me.lblCurrentTime.Text = "Hora actual:" : Me.lblCurrentTimeValue.Text = "-"
        Me.lblPieceTime.Text = "Tiempo pieza:" : Me.lblPieceTimeValue.Text = "00:00"
        Me.lblTotalTime.Text = "Tiempo total:" : Me.lblTotalTimeValue.Text = "00:00"
        Me.lblStatus.Text = "Estado:" : Me.lblStatusValue.Text = "-"
        Me.tblProgress.Controls.Add(Me.lblCurrentTime, 0, 0) : Me.tblProgress.Controls.Add(Me.lblCurrentTimeValue, 1, 0)
        Me.tblProgress.Controls.Add(Me.lblPieceTime, 0, 1) : Me.tblProgress.Controls.Add(Me.lblPieceTimeValue, 1, 1)
        Me.tblProgress.Controls.Add(Me.lblTotalTime, 0, 2) : Me.tblProgress.Controls.Add(Me.lblTotalTimeValue, 1, 2)
        Me.tblProgress.Controls.Add(Me.lblStatus, 0, 3) : Me.tblProgress.Controls.Add(Me.lblStatusValue, 1, 3)
        Me.pnlProgressBars.Dock = DockStyle.Fill
        Me.pnlProgressBars.Padding = New Padding(6, 4, 6, 4)
        Me.lblProgressPiece.AutoSize = True
        Me.lblProgressPiece.Dock = DockStyle.Top
        Me.lblProgressPiece.Margin = New Padding(0, 0, 0, 2)
        Me.lblProgressPiece.Text = "Pieza (log ~360 líneas)"
        Me.progressBar.Dock = DockStyle.Top
        Me.progressBar.Height = 14
        Me.progressBar.Margin = New Padding(0, 0, 0, 8)
        Me.lblProgressAsm.AutoSize = True
        Me.lblProgressAsm.Dock = DockStyle.Top
        Me.lblProgressAsm.Margin = New Padding(0, 0, 0, 2)
        Me.lblProgressAsm.Text = "Ensamblaje (piezas marcadas)"
        Me.lblProgressAsm.Visible = False
        Me.progressBarAsm.Dock = DockStyle.Top
        Me.progressBarAsm.Height = 14
        Me.progressBarAsm.Margin = New Padding(0, 0, 0, 0)
        Me.progressBarAsm.Visible = False
        Me.pnlProgressBars.Controls.Add(Me.lblProgressPiece)
        Me.pnlProgressBars.Controls.Add(Me.progressBar)
        Me.pnlProgressBars.Controls.Add(Me.lblProgressAsm)
        Me.pnlProgressBars.Controls.Add(Me.progressBarAsm)
        Me.tblProgress.Controls.Add(Me.pnlProgressBars, 2, 0)
        Me.tblProgress.SetRowSpan(Me.pnlProgressBars, 4)
        Me.tblProgress.Margin = New Padding(0)
        Me.grpProgress.Controls.Add(Me.tblProgress)

        Me.tblLog.Dock = DockStyle.Fill
        Me.tblLog.ColumnCount = 1
        Me.tblLog.ColumnStyles.Clear()
        Me.tblLog.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0!))
        Me.tblLog.RowCount = 2
        Me.tblLog.RowStyles.Clear()
        Me.tblLog.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0!))
        Me.tblLog.RowStyles.Add(New RowStyle(SizeType.Absolute, 36.0!))
        Me.txtLog.Dock = DockStyle.Fill : Me.txtLog.Multiline = True : Me.txtLog.ScrollBars = ScrollBars.Vertical : Me.txtLog.Font = New Font("Consolas", 9.0!)
        Me.flowLogButtons.Dock = DockStyle.Fill : Me.flowLogButtons.Padding = New Padding(8, 4, 8, 4)
        Me.btnSaveLog.Text = "Guardar log" : Me.btnClearLog.Text = "Limpiar log"
        Me.flowLogButtons.Controls.Add(Me.btnSaveLog) : Me.flowLogButtons.Controls.Add(Me.btnClearLog)
        Me.tblLog.Controls.Add(Me.txtLog, 0, 0) : Me.tblLog.Controls.Add(Me.flowLogButtons, 0, 1)
        Me.grpLog.Controls.Add(Me.tblLog)

        Me.mainLayout.Controls.Add(Me.pnlHeader, 0, 0)
        Me.mainLayout.Controls.Add(Me.topInputOptionsLayout, 0, 1)
        Me.mainLayout.Controls.Add(Me.bodyHost, 0, 2)
        Me.Controls.Add(Me.mainLayout)
        Me.Controls.Add(Me.pnlHiddenLegacyGeneration)
        Me.Controls.Add(Me.pnlHiddenProcessing)

        Me.AutoScaleMode = AutoScaleMode.Font
        Me.ClientSize = New Size(1400, 900)
        Me.Font = New Font("Segoe UI", 9.0!)
        Me.MinimumSize = New Size(1280, 820)
        Me.Name = "MainForm"
        Me.StartPosition = FormStartPosition.CenterScreen

        CType(Me.dgvAsmComponents, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
    End Sub

    Friend WithEvents mainLayout As TableLayoutPanel
    Friend WithEvents pnlHeader As Panel
    Friend WithEvents lblTitle As Label
    Friend WithEvents lblSubTitle As Label
    Friend WithEvents bodyHost As TableLayoutPanel
    Friend WithEvents threeColumnLayout As TableLayoutPanel
    Friend WithEvents pnlGenerateBar As Panel
    Friend WithEvents rightLayout As TableLayoutPanel
    Friend WithEvents tblRightLogProgress As TableLayoutPanel
    Friend WithEvents topInputOptionsLayout As TableLayoutPanel
    Friend WithEvents grpInput As GroupBox
    Friend WithEvents grpAsmComponents As GroupBox
    Friend WithEvents grpTemplates As GroupBox
    Friend WithEvents grpTraceability As GroupBox
    Friend WithEvents grpGeneration As GroupBox
    Friend WithEvents grpAdvanced As GroupBox
    Friend WithEvents pnlHiddenLegacyGeneration As Panel
    Friend WithEvents pnlHiddenProcessing As Panel
    Friend WithEvents grpProgress As GroupBox
    Friend WithEvents grpLog As GroupBox
    Friend WithEvents tblInput As TableLayoutPanel
    Friend WithEvents tblAsmComponents As TableLayoutPanel
    Friend WithEvents dgvAsmComponents As DataGridView
    Friend WithEvents flowAsmButtons As FlowLayoutPanel
    Friend WithEvents btnLoadAsmComponents As Button
    Friend WithEvents btnSelectAllComponents As Button
    Friend WithEvents btnSelectNoneComponents As Button
    Friend WithEvents lblAsmComponentHint As Label
    Friend WithEvents lblInput As Label
    Friend WithEvents txtInputFile As TextBox
    Friend WithEvents btnBrowseInput As Button
    Friend WithEvents lblDetectedType As Label
    Friend WithEvents lblDetectedTypeValue As Label
    Friend WithEvents tblTemplates As TableLayoutPanel
    Friend WithEvents lblTplA4 As Label
    Friend WithEvents lblTplA3 As Label
    Friend WithEvents lblTplA2 As Label
    Friend WithEvents lblTplDxf As Label
    Friend WithEvents txtTemplateA4 As TextBox
    Friend WithEvents txtTemplateA3 As TextBox
    Friend WithEvents txtTemplateA2 As TextBox
    Friend WithEvents txtTemplateDxf As TextBox
    Friend WithEvents btnBrowseA4 As Button
    Friend WithEvents btnBrowseA3 As Button
    Friend WithEvents btnBrowseA2 As Button
    Friend WithEvents btnBrowseDxf As Button
    Friend WithEvents tblTrace As TableLayoutPanel
    Friend WithEvents lblClient As Label
    Friend WithEvents txtClient As TextBox
    Friend WithEvents lblProject As Label
    Friend WithEvents txtProject As TextBox
    Friend WithEvents lblDrawingTitle As Label
    Friend WithEvents txtTitle As TextBox
        Friend WithEvents lblMaterial As Label
        Friend WithEvents txtMaterial As TextBox
        Friend WithEvents lblDrawingNumber As Label
        Friend WithEvents txtDrawingNumber As TextBox
        Friend WithEvents lblRevision As Label
        Friend WithEvents txtRevision As TextBox
        Friend WithEvents lblAuthor As Label
        Friend WithEvents txtAuthor As TextBox
        Friend WithEvents lblThickness As Label
        Friend WithEvents txtThickness As TextBox
        Friend WithEvents lblPedido As Label
        Friend WithEvents txtPedido As TextBox
        Friend WithEvents cmbTitleSource As ComboBox
        Friend WithEvents lblOriginTitle As Label
        Friend WithEvents lblOriginProject As Label
        Friend WithEvents lblOriginMaterial As Label
        Friend WithEvents lblOriginClient As Label
        Friend WithEvents lblOriginDocNum As Label
        Friend WithEvents lblOriginRevision As Label
        Friend WithEvents lblOriginAuthor As Label
        Friend WithEvents lblOriginThickness As Label
        Friend WithEvents lblOriginPedido As Label
        Friend WithEvents lblNotes As Label
        Friend WithEvents txtNotes As TextBox
    Friend WithEvents flowTraceButtons As FlowLayoutPanel
    Friend WithEvents btnApplyTraceability As Button
    Friend WithEvents flowGeneration As FlowLayoutPanel
    Friend WithEvents lblTitleBlockSource As Label
    Friend WithEvents cmbTitleBlockSource As ComboBox
    Friend WithEvents chkCreateDft As CheckBox
    Friend WithEvents chkCreatePdf As CheckBox
    Friend WithEvents chkCreateDxfDraft As CheckBox
    Friend WithEvents chkAutoDimensioning As CheckBox
    Friend WithEvents chkUnitHorizontalExteriorTest As CheckBox
        Friend WithEvents chkPmiRetrievalProbe As CheckBox
        Friend WithEvents chkExperimentalPmiModelView As CheckBox
        Friend WithEvents chkCreateFlatDxf As CheckBox
    Friend WithEvents chkOpenOutput As CheckBox
    Friend WithEvents chkOverwrite As CheckBox
    Friend WithEvents chkUniqueComponents As CheckBox
    Friend WithEvents chkDetailedLog As CheckBox
    Friend WithEvents chkKeepSolidEdgeVisible As CheckBox
        Friend WithEvents chkInsertProperties As CheckBox
        Friend WithEvents chkDebugTemplates As CheckBox
        Friend WithEvents chkExperimentalDraftGeometryDiagnostics As CheckBox
        Friend WithEvents tblAdvanced As TableLayoutPanel
    Friend WithEvents lblOutput As Label
    Friend WithEvents txtOutputFolder As TextBox
    Friend WithEvents btnBrowseOut As Button
    Friend WithEvents lblPreferredFormat As Label
    Friend WithEvents cmbPreferredFormat As ComboBox
    Friend WithEvents chkAutoScale As CheckBox
    Friend WithEvents txtManualScale As TextBox
    Friend WithEvents chkIncludeIso As CheckBox
    Friend WithEvents chkIncludeProjected As CheckBox
    Friend WithEvents chkIncludeFlatInDraft As CheckBox
    Friend WithEvents chkUseBestBase As CheckBox
    Friend WithEvents btnGenerate As Button
    Friend WithEvents btnClear As Button
    Friend WithEvents btnOpenOutput As Button
    Friend WithEvents btnSaveConfig As Button
    Friend WithEvents btnLoadConfig As Button
    Friend WithEvents btnReloadSourceProps As Button
    Friend WithEvents btnRestoreDefaultTemplates As Button
    Friend WithEvents tblProgress As TableLayoutPanel
    Friend WithEvents lblCurrentTime As Label
    Friend WithEvents lblCurrentTimeValue As Label
    Friend WithEvents lblPieceTime As Label
    Friend WithEvents lblPieceTimeValue As Label
    Friend WithEvents lblTotalTime As Label
    Friend WithEvents lblTotalTimeValue As Label
    Friend WithEvents lblStatus As Label
    Friend WithEvents lblStatusValue As Label
    Friend WithEvents pnlProgressBars As Panel
    Friend WithEvents lblProgressPiece As Label
    Friend WithEvents progressBar As ProgressBar
    Friend WithEvents lblProgressAsm As Label
    Friend WithEvents progressBarAsm As ProgressBar
    Friend WithEvents tblLog As TableLayoutPanel
    Friend WithEvents txtLog As TextBox
    Friend WithEvents flowLogButtons As FlowLayoutPanel
    Friend WithEvents btnSaveLog As Button
    Friend WithEvents btnClearLog As Button
End Class
