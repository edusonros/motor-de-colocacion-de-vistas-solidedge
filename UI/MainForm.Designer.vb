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
        Me.twoColumnBodyLayout = New TableLayoutPanel()
        Me.pnlSharedSidebar = New TableLayoutPanel()
        Me.tabMotors = New TabControl()
        Me.tabPageMotorViews = New TabPage()
        Me.tabPageMotorMetadata = New TabPage()
        Me.tabPageMotorDimensions = New TabPage()
        Me.tabPageLaserCut = New TabPage()
        Me.tblLaserTabHost = New TableLayoutPanel()
        Me.tblLaserTabRightColumn = New TableLayoutPanel()
        Me.grpLaserPieces = New GroupBox()
        Me.tblLaserPiecesInner = New TableLayoutPanel()
        Me.flowLaserButtons = New FlowLayoutPanel()
        Me.btnLaserScan = New Button()
        Me.btnLaserGenerate = New Button()
        Me.lblLaserSummary = New Label()
        Me.dgvLaserPieces = New DataGridView()
        Me.tblMetadataTabHost = New TableLayoutPanel()
        Me.tblMetadataTabRightColumn = New TableLayoutPanel()
        Me.tblMetadataPlanTwoCols = New TableLayoutPanel()
        Me.tblDimensionTabHost = New TableLayoutPanel()
        Me.tblDimensionTabRightColumn = New TableLayoutPanel()
        Me.tblMotorTab1Host = New TableLayoutPanel()
        Me.tblMotorTab1LeftColumn = New TableLayoutPanel()
        Me.tblMotorTab1RightColumn = New TableLayoutPanel()
        Me.tblMotorRightTemplatesFormat = New TableLayoutPanel()
        Me.pnlMotorViewsAdvancedHost = New Panel()
        Me.grpPdfPreview = New GroupBox()
        Me.tblPdfPreviewInner = New TableLayoutPanel()
        Me.flowPdfPreviewBar = New FlowLayoutPanel()
        Me.btnRefreshPdfPreview = New Button()
        Me.btnOpenLastPdfExternal = New Button()
        Me.wbPdfPreview = New WebBrowser()
        Me.pnlGenerateBar = New Panel()
        Me.flowGenerateBar = New FlowLayoutPanel()
        Me.btnMotorViews = New Button()
        Me.btnMotorMetadata = New Button()
        Me.btnMotorDimensioning = New Button()
        Me.btnOpenVariableTable = New Button()
        Me.btnMotorLaser = New Button()
        Me.tblRightLogProgress = New TableLayoutPanel()
        Me.grpInput = New GroupBox()
        Me.grpAsmComponents = New GroupBox()
        Me.grpTemplates = New GroupBox()
        Me.grpTraceability = New GroupBox()
        Me.grpGeneration = New GroupBox()
        Me.grpAdvanced = New GroupBox()
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

        Me.tblGenerationTwoCols = New TableLayoutPanel()
        Me.flowGenerationLeft = New FlowLayoutPanel()
        Me.flowGenerationRight = New FlowLayoutPanel()
        Me.lblTitleBlockSource = New Label()
        Me.cmbTitleBlockSource = New ComboBox()
        Me.chkCreateDft = New CheckBox()
        Me.chkCreatePdf = New CheckBox()
        Me.chkCreateDxfDraft = New CheckBox()
        Me.chkAutoDimensioning = New CheckBox()
        Me.pnlAutoDimensionMotorChoice = New FlowLayoutPanel()
        Me.lblAutoDimensionMotor = New Label()
        Me.radAutoDimMotorMain = New RadioButton()
        Me.radAutoDimMotorLegacyV02 = New RadioButton()
        Me.radAutoDimMotorAlternatePlugIn = New RadioButton()
        Me.chkSesdkPostDimensionIntrospection = New CheckBox()
        Me.chkPreferSweepAllDrawingDimensions = New CheckBox()
        Me.chkSuppressDimTrackSpacing = New CheckBox()
        Me.chkDedupDimensionsByKeypoints = New CheckBox()
        Me.chkUnitHorizontalExteriorTest = New CheckBox()
        Me.chkDrawingViewDimensioningLab = New CheckBox()
        Me.chkDimLabInteractivePause = New CheckBox()
        Me.lblDimLabMode = New Label()
        Me.cmbDimLabMode = New ComboBox()
        Me.chkDimLabVisibleProbe = New CheckBox()
        Me.chkDimLabAlternativePlacement = New CheckBox()
        Me.btnDimLabRun = New Button()
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
        CType(Me.dgvLaserPieces, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()

        Me.mainLayout.Dock = DockStyle.Fill
        Me.mainLayout.ColumnCount = 1
        Me.mainLayout.ColumnStyles.Clear()
        Me.mainLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0!))
        Me.mainLayout.RowCount = 2
        Me.mainLayout.RowStyles.Clear()
        Me.mainLayout.RowStyles.Add(New RowStyle(SizeType.Absolute, 70.0!))
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
        Me.bodyHost.RowStyles.Add(New RowStyle(SizeType.Absolute, 74.0!))
        Me.bodyHost.Controls.Add(Me.twoColumnBodyLayout, 0, 0)
        Me.bodyHost.Controls.Add(Me.pnlGenerateBar, 0, 1)

        Me.twoColumnBodyLayout.Dock = DockStyle.Fill
        Me.twoColumnBodyLayout.ColumnCount = 2
        Me.twoColumnBodyLayout.ColumnStyles.Clear()
        Me.twoColumnBodyLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 46.0!))
        Me.twoColumnBodyLayout.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 54.0!))
        Me.twoColumnBodyLayout.RowCount = 1
        Me.twoColumnBodyLayout.RowStyles.Clear()
        Me.twoColumnBodyLayout.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0!))
        Me.twoColumnBodyLayout.Margin = New Padding(0)
        Me.twoColumnBodyLayout.Controls.Add(Me.tabMotors, 0, 0)
        Me.twoColumnBodyLayout.Controls.Add(Me.pnlSharedSidebar, 1, 0)

        Me.tabMotors.Dock = DockStyle.Fill
        Me.tabMotors.Margin = New Padding(4, 4, 2, 4)
        Me.tabPageMotorViews.Text = "Motor de vistas"
        Me.tabPageMotorMetadata.Text = "Motor de metadatos"
        Me.tabPageMotorDimensions.Text = "Motor de acotación"
        Me.tabPageLaserCut.Text = "Piezas a Corte Laser"
        Me.tabMotors.Controls.Add(Me.tabPageMotorViews)
        Me.tabMotors.Controls.Add(Me.tabPageMotorMetadata)
        Me.tabMotors.Controls.Add(Me.tabPageMotorDimensions)
        Me.tabMotors.Controls.Add(Me.tabPageLaserCut)

        Me.tblMotorTab1LeftColumn.Dock = DockStyle.Fill
        Me.tblMotorTab1LeftColumn.Margin = New Padding(0, 0, 4, 0)
        Me.tblMotorTab1LeftColumn.ColumnCount = 1
        Me.tblMotorTab1LeftColumn.ColumnStyles.Clear()
        Me.tblMotorTab1LeftColumn.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0!))
        Me.tblMotorTab1LeftColumn.RowCount = 3
        Me.tblMotorTab1LeftColumn.RowStyles.Clear()
        Me.tblMotorTab1LeftColumn.RowStyles.Add(New RowStyle(SizeType.Absolute, 108.0!))
        Me.tblMotorTab1LeftColumn.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0!))
        Me.tblMotorTab1LeftColumn.RowStyles.Add(New RowStyle(SizeType.Absolute, 132.0!))

        Me.pnlMotorViewsAdvancedHost.Dock = DockStyle.Fill
        Me.pnlMotorViewsAdvancedHost.Padding = New Padding(2, 0, 0, 0)
        Me.pnlMotorViewsAdvancedHost.Margin = New Padding(0)

        Me.grpPdfPreview.Dock = DockStyle.Fill
        Me.grpPdfPreview.Text = "Vista previa del último PDF generado"
        Me.grpPdfPreview.Padding = New Padding(6, 8, 6, 6)
        Me.grpPdfPreview.Margin = New Padding(0, 4, 0, 0)
        Me.tblPdfPreviewInner.Dock = DockStyle.Fill
        Me.tblPdfPreviewInner.ColumnCount = 1
        Me.tblPdfPreviewInner.ColumnStyles.Clear()
        Me.tblPdfPreviewInner.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0!))
        Me.tblPdfPreviewInner.RowCount = 2
        Me.tblPdfPreviewInner.RowStyles.Clear()
        Me.tblPdfPreviewInner.RowStyles.Add(New RowStyle(SizeType.Absolute, 36.0!))
        Me.tblPdfPreviewInner.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0!))
        Me.tblPdfPreviewInner.Margin = New Padding(0)
        Me.flowPdfPreviewBar.Dock = DockStyle.Fill
        Me.flowPdfPreviewBar.FlowDirection = FlowDirection.LeftToRight
        Me.flowPdfPreviewBar.WrapContents = False
        Me.btnRefreshPdfPreview.Text = "Actualizar vista previa"
        Me.btnRefreshPdfPreview.AutoSize = True
        Me.btnRefreshPdfPreview.Margin = New Padding(0, 4, 12, 4)
        Me.btnOpenLastPdfExternal.Text = "Abrir PDF en el visor del sistema"
        Me.btnOpenLastPdfExternal.AutoSize = True
        Me.btnOpenLastPdfExternal.Margin = New Padding(0, 4, 0, 4)
        Me.flowPdfPreviewBar.Controls.Add(Me.btnRefreshPdfPreview)
        Me.flowPdfPreviewBar.Controls.Add(Me.btnOpenLastPdfExternal)
        Me.wbPdfPreview.Dock = DockStyle.Fill
        Me.wbPdfPreview.Margin = New Padding(0)
        Me.wbPdfPreview.ScriptErrorsSuppressed = True
        Me.tblPdfPreviewInner.Controls.Add(Me.flowPdfPreviewBar, 0, 0)
        Me.tblPdfPreviewInner.Controls.Add(Me.wbPdfPreview, 0, 1)
        Me.grpPdfPreview.Controls.Add(Me.tblPdfPreviewInner)

        Me.tblMotorTab1LeftColumn.Controls.Add(Me.grpPdfPreview, 0, 2)

        Me.tblMotorRightTemplatesFormat.Dock = DockStyle.Fill
        Me.tblMotorRightTemplatesFormat.Margin = New Padding(0)
        Me.tblMotorRightTemplatesFormat.ColumnCount = 2
        Me.tblMotorRightTemplatesFormat.ColumnStyles.Clear()
        Me.tblMotorRightTemplatesFormat.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50.0!))
        Me.tblMotorRightTemplatesFormat.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50.0!))
        Me.tblMotorRightTemplatesFormat.RowCount = 1
        Me.tblMotorRightTemplatesFormat.RowStyles.Clear()
        Me.tblMotorRightTemplatesFormat.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0!))
        Me.tblMotorRightTemplatesFormat.Controls.Add(Me.grpTemplates, 0, 0)
        Me.tblMotorRightTemplatesFormat.Controls.Add(Me.pnlMotorViewsAdvancedHost, 1, 0)

        Me.tblMotorTab1Host.Dock = DockStyle.Fill
        Me.tblMotorTab1Host.Margin = New Padding(0)
        Me.tblMotorTab1Host.ColumnCount = 2
        Me.tblMotorTab1Host.ColumnStyles.Clear()
        Me.tblMotorTab1Host.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50.0!))
        Me.tblMotorTab1Host.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50.0!))
        Me.tblMotorTab1Host.RowCount = 1
        Me.tblMotorTab1Host.RowStyles.Clear()
        Me.tblMotorTab1Host.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0!))
        Me.tblMotorTab1Host.Controls.Add(Me.tblMotorTab1LeftColumn, 0, 0)
        Me.tblMotorTab1Host.Controls.Add(Me.tblMotorTab1RightColumn, 1, 0)

        Me.tblMotorTab1RightColumn.Dock = DockStyle.Fill
        Me.tblMotorTab1RightColumn.Margin = New Padding(4, 0, 0, 0)
        Me.tblMotorTab1RightColumn.ColumnCount = 1
        Me.tblMotorTab1RightColumn.ColumnStyles.Clear()
        Me.tblMotorTab1RightColumn.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0!))
        Me.tblMotorTab1RightColumn.RowCount = 2
        Me.tblMotorTab1RightColumn.RowStyles.Clear()
        Me.tblMotorTab1RightColumn.RowStyles.Add(New RowStyle(SizeType.Absolute, 204.0!))
        Me.tblMotorTab1RightColumn.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0!))
        Me.tblMotorTab1RightColumn.Controls.Add(Me.tblMotorRightTemplatesFormat, 0, 0)

        Me.tabPageMotorViews.Controls.Add(Me.tblMotorTab1Host)

        Me.tblMetadataTabHost.Dock = DockStyle.Fill
        Me.tblMetadataTabHost.Margin = New Padding(0)
        Me.tblMetadataTabHost.ColumnCount = 2
        Me.tblMetadataTabHost.ColumnStyles.Clear()
        Me.tblMetadataTabHost.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50.0!))
        Me.tblMetadataTabHost.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50.0!))
        Me.tblMetadataTabHost.RowCount = 1
        Me.tblMetadataTabHost.RowStyles.Clear()
        Me.tblMetadataTabHost.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0!))

        Me.tblMetadataTabRightColumn.Dock = DockStyle.Fill
        Me.tblMetadataTabRightColumn.Margin = New Padding(4, 0, 0, 0)
        Me.tblMetadataTabRightColumn.ColumnCount = 1
        Me.tblMetadataTabRightColumn.ColumnStyles.Clear()
        Me.tblMetadataTabRightColumn.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0!))
        Me.tblMetadataTabRightColumn.RowCount = 2
        Me.tblMetadataTabRightColumn.RowStyles.Clear()
        Me.tblMetadataTabRightColumn.RowStyles.Add(New RowStyle(SizeType.Percent, 36.0!))
        Me.tblMetadataTabRightColumn.RowStyles.Add(New RowStyle(SizeType.Percent, 64.0!))

        Me.tblMetadataPlanTwoCols.Dock = DockStyle.Fill
        Me.tblMetadataPlanTwoCols.Margin = New Padding(0)
        Me.tblMetadataPlanTwoCols.ColumnCount = 2
        Me.tblMetadataPlanTwoCols.ColumnStyles.Clear()
        Me.tblMetadataPlanTwoCols.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50.0!))
        Me.tblMetadataPlanTwoCols.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50.0!))
        Me.tblMetadataPlanTwoCols.RowCount = 1
        Me.tblMetadataPlanTwoCols.RowStyles.Clear()
        Me.tblMetadataPlanTwoCols.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0!))
        Me.tblMetadataTabRightColumn.Controls.Add(Me.tblMetadataPlanTwoCols, 0, 0)

        Me.tblMetadataTabHost.Controls.Add(Me.tblMetadataTabRightColumn, 1, 0)
        Me.tabPageMotorMetadata.Controls.Add(Me.tblMetadataTabHost)
        Me.tabPageMotorMetadata.Controls.Add(Me.grpTraceability)
        Me.grpTraceability.Visible = False

        Me.tblDimensionTabHost.Dock = DockStyle.Fill
        Me.tblDimensionTabHost.Margin = New Padding(0)
        Me.tblDimensionTabHost.ColumnCount = 2
        Me.tblDimensionTabHost.ColumnStyles.Clear()
        Me.tblDimensionTabHost.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50.0!))
        Me.tblDimensionTabHost.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50.0!))
        Me.tblDimensionTabHost.RowCount = 1
        Me.tblDimensionTabHost.RowStyles.Clear()
        Me.tblDimensionTabHost.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0!))

        Me.tblDimensionTabRightColumn.Dock = DockStyle.Fill
        Me.tblDimensionTabRightColumn.Margin = New Padding(4, 0, 0, 0)
        Me.tblDimensionTabRightColumn.ColumnCount = 1
        Me.tblDimensionTabRightColumn.ColumnStyles.Clear()
        Me.tblDimensionTabRightColumn.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0!))
        Me.tblDimensionTabRightColumn.RowCount = 2
        Me.tblDimensionTabRightColumn.RowStyles.Clear()
        Me.tblDimensionTabRightColumn.RowStyles.Add(New RowStyle(SizeType.Percent, 36.0!))
        Me.tblDimensionTabRightColumn.RowStyles.Add(New RowStyle(SizeType.Percent, 64.0!))
        Me.grpGeneration.Dock = DockStyle.Fill
        Me.tblDimensionTabRightColumn.Controls.Add(Me.grpGeneration, 0, 0)
        Me.tblDimensionTabHost.Controls.Add(Me.tblDimensionTabRightColumn, 1, 0)
        Me.tabPageMotorDimensions.Controls.Add(Me.tblDimensionTabHost)

        Me.tblLaserTabHost.Dock = DockStyle.Fill
        Me.tblLaserTabHost.Margin = New Padding(0)
        Me.tblLaserTabHost.ColumnCount = 1
        Me.tblLaserTabHost.ColumnStyles.Clear()
        Me.tblLaserTabHost.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0!))
        Me.tblLaserTabHost.RowCount = 3
        Me.tblLaserTabHost.RowStyles.Clear()
        Me.tblLaserTabHost.RowStyles.Add(New RowStyle(SizeType.Absolute, 118.0!))
        Me.tblLaserTabHost.RowStyles.Add(New RowStyle(SizeType.Percent, 72.0!))
        Me.tblLaserTabHost.RowStyles.Add(New RowStyle(SizeType.Percent, 28.0!))

        Me.tblLaserTabRightColumn.Dock = DockStyle.Fill
        Me.tblLaserTabRightColumn.Margin = New Padding(0)
        Me.tblLaserTabRightColumn.ColumnCount = 1
        Me.tblLaserTabRightColumn.ColumnStyles.Clear()
        Me.tblLaserTabRightColumn.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0!))
        Me.tblLaserTabRightColumn.RowCount = 1
        Me.tblLaserTabRightColumn.RowStyles.Clear()
        Me.tblLaserTabRightColumn.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0!))

        Me.grpLaserPieces.Dock = DockStyle.Fill
        Me.grpLaserPieces.Text = "Piezas para corte láser (editar antes de generar)"
        Me.grpLaserPieces.Margin = New Padding(0, 0, 0, 4)
        Me.tblLaserPiecesInner.Dock = DockStyle.Fill
        Me.tblLaserPiecesInner.ColumnCount = 1
        Me.tblLaserPiecesInner.ColumnStyles.Clear()
        Me.tblLaserPiecesInner.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0!))
        Me.tblLaserPiecesInner.RowCount = 3
        Me.tblLaserPiecesInner.RowStyles.Clear()
        Me.tblLaserPiecesInner.RowStyles.Add(New RowStyle(SizeType.Absolute, 40.0!))
        Me.tblLaserPiecesInner.RowStyles.Add(New RowStyle(SizeType.Absolute, 28.0!))
        Me.tblLaserPiecesInner.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0!))
        Me.flowLaserButtons.Dock = DockStyle.Fill
        Me.flowLaserButtons.FlowDirection = FlowDirection.LeftToRight
        Me.flowLaserButtons.WrapContents = False
        Me.btnLaserScan.Text = "Escanear ASM"
        Me.btnLaserScan.AutoSize = True
        Me.btnLaserScan.Margin = New Padding(0, 6, 8, 0)
        Me.btnLaserGenerate.Text = "Generar DFT/DXF corte"
        Me.btnLaserGenerate.AutoSize = True
        Me.btnLaserGenerate.Margin = New Padding(0, 6, 0, 0)
        Me.flowLaserButtons.Controls.Add(Me.btnLaserScan)
        Me.flowLaserButtons.Controls.Add(Me.btnLaserGenerate)
        Me.lblLaserSummary.Dock = DockStyle.Fill
        Me.lblLaserSummary.Text = "Escanee un ASM para cargar la tabla."
        Me.lblLaserSummary.AutoEllipsis = True
        Me.dgvLaserPieces.Dock = DockStyle.Fill
        Me.dgvLaserPieces.MinimumSize = New Size(400, 220)
        Me.tblLaserPiecesInner.Controls.Add(Me.flowLaserButtons, 0, 0)
        Me.tblLaserPiecesInner.Controls.Add(Me.lblLaserSummary, 0, 1)
        Me.tblLaserPiecesInner.Controls.Add(Me.dgvLaserPieces, 0, 2)
        Me.grpLaserPieces.Controls.Add(Me.tblLaserPiecesInner)
        Me.tblLaserTabHost.Controls.Add(Me.grpInput, 0, 0)
        Me.tblLaserTabHost.Controls.Add(Me.grpLaserPieces, 0, 1)
        Me.tblLaserTabHost.Controls.Add(Me.grpProgress, 0, 2)
        Me.tabPageLaserCut.Controls.Add(Me.tblLaserTabHost)

        Me.pnlSharedSidebar.Dock = DockStyle.Fill
        Me.pnlSharedSidebar.ColumnCount = 1
        Me.pnlSharedSidebar.ColumnStyles.Clear()
        Me.pnlSharedSidebar.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0!))
        Me.pnlSharedSidebar.RowCount = 3
        Me.pnlSharedSidebar.RowStyles.Clear()
        Me.pnlSharedSidebar.RowStyles.Add(New RowStyle(SizeType.Absolute, 148.0!))
        Me.pnlSharedSidebar.RowStyles.Add(New RowStyle(SizeType.Percent, 38.0!))
        Me.pnlSharedSidebar.RowStyles.Add(New RowStyle(SizeType.Percent, 62.0!))
        Me.pnlSharedSidebar.Margin = New Padding(0, 4, 6, 4)
        Me.pnlSharedSidebar.Controls.Add(Me.grpInput, 0, 0)
        Me.pnlSharedSidebar.Controls.Add(Me.grpAsmComponents, 0, 1)
        Me.pnlSharedSidebar.Controls.Add(Me.tblRightLogProgress, 0, 2)

        Me.pnlGenerateBar.Dock = DockStyle.Fill
        Me.pnlGenerateBar.Padding = New Padding(8, 6, 8, 6)
        Me.flowGenerateBar.Dock = DockStyle.Fill
        Me.flowGenerateBar.FlowDirection = FlowDirection.LeftToRight
        Me.flowGenerateBar.WrapContents = True
        Me.flowGenerateBar.AutoScroll = True
        Me.flowGenerateBar.Padding = New Padding(0, 2, 0, 2)
        Me.btnMotorViews.Text = "Generador de vistas"
        Me.btnMotorViews.AutoSize = True
        Me.btnMotorViews.MinimumSize = New Size(160, 44)
        Me.btnMotorViews.Margin = New Padding(0, 0, 8, 6)
        Me.btnMotorMetadata.Text = "Gestor Metadatos DFT"
        Me.btnMotorMetadata.AutoSize = True
        Me.btnMotorMetadata.MinimumSize = New Size(180, 44)
        Me.btnMotorMetadata.Margin = New Padding(0, 0, 8, 6)
        Me.btnMotorDimensioning.Text = "Motor Acotación"
        Me.btnMotorDimensioning.AutoSize = True
        Me.btnMotorDimensioning.MinimumSize = New Size(150, 44)
        Me.btnMotorDimensioning.Margin = New Padding(0, 0, 8, 6)
        Me.btnOpenVariableTable.Text = "Tabla de variables (SE)"
        Me.btnOpenVariableTable.AutoSize = True
        Me.btnOpenVariableTable.MinimumSize = New Size(170, 44)
        Me.btnOpenVariableTable.Margin = New Padding(0, 0, 8, 6)
        Me.btnMotorLaser.Text = "Corte láser"
        Me.btnMotorLaser.AutoSize = True
        Me.btnMotorLaser.MinimumSize = New Size(120, 44)
        Me.btnMotorLaser.Margin = New Padding(0, 0, 8, 6)
        Me.flowGenerateBar.Controls.Add(Me.btnMotorViews)
        Me.flowGenerateBar.Controls.Add(Me.btnMotorMetadata)
        Me.flowGenerateBar.Controls.Add(Me.btnMotorDimensioning)
        Me.flowGenerateBar.Controls.Add(Me.btnMotorLaser)
        Me.flowGenerateBar.Controls.Add(Me.btnOpenVariableTable)
        Me.flowGenerateBar.Controls.Add(Me.btnGenerate)
        Me.pnlGenerateBar.Controls.Add(Me.flowGenerateBar)

        Me.tblRightLogProgress.Dock = DockStyle.Fill
        Me.tblRightLogProgress.ColumnCount = 1
        Me.tblRightLogProgress.ColumnStyles.Clear()
        Me.tblRightLogProgress.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0!))
        Me.tblRightLogProgress.RowCount = 2
        Me.tblRightLogProgress.RowStyles.Clear()
        Me.tblRightLogProgress.RowStyles.Add(New RowStyle(SizeType.Percent, 57.0!))
        Me.tblRightLogProgress.RowStyles.Add(New RowStyle(SizeType.Percent, 43.0!))
        Me.tblRightLogProgress.Margin = New Padding(0)
        Me.tblRightLogProgress.Padding = New Padding(0)
        Me.tblRightLogProgress.Controls.Add(Me.grpLog, 0, 0)
        Me.tblRightLogProgress.Controls.Add(Me.grpProgress, 0, 1)

        For Each gb In New GroupBox() {Me.grpInput, Me.grpAsmComponents, Me.grpTemplates, Me.grpTraceability, Me.grpGeneration, Me.grpAdvanced, Me.grpProgress, Me.grpLog, Me.grpPdfPreview, Me.grpLaserPieces}
            gb.Dock = DockStyle.Fill
            gb.Font = New Font("Segoe UI", 9.0!)
            gb.Margin = New Padding(8)
        Next
        Me.grpGeneration.Margin = New Padding(4, 6, 2, 6)
        Me.grpTemplates.Margin = New Padding(2, 6, 4, 6)
        Me.grpInput.Margin = New Padding(4, 4, 6, 4)
        Me.grpAsmComponents.Margin = New Padding(4, 4, 6, 4)
        Me.grpLog.Margin = New Padding(4, 2, 6, 4)
        Me.grpProgress.Margin = New Padding(4, 2, 6, 6)
        Me.grpInput.Text = "Archivo de entrada"
        Me.grpAsmComponents.Text = "Componentes detectados del ASM (marcar los que deseas procesar)"
        Me.grpTemplates.Text = "Templates"
        Me.grpTraceability.Text = "Metadatos (plano y pieza)"
        Me.grpGeneration.Text = "Opciones de generación"
        Me.grpAdvanced.Text = "Opciones de procesado"
        Me.grpProgress.Text = "Progreso y estado de generación"
        Me.grpLog.Text = "Log"

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
        Me.tblAsmComponents.RowStyles.Add(New RowStyle(SizeType.Absolute, 24.0!))
        Me.tblAsmComponents.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0!))
        Me.tblAsmComponents.RowStyles.Add(New RowStyle(SizeType.Absolute, 30.0!))
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
        Me.btnApplyTraceability.Text = "Aplicar cajetín + PART_LIST al DFT"
        Me.btnApplyTraceability.Size = New Size(220, 30)
        Me.flowTraceButtons.Controls.Add(Me.btnApplyTraceability)
        Me.tblTrace.Controls.Add(Me.flowTraceButtons, 0, 9)
        Me.tblTrace.SetColumnSpan(Me.flowTraceButtons, 3)

        Me.grpTraceability.Controls.Add(Me.tblTrace)

        Me.tblGenerationTwoCols.Dock = DockStyle.Fill
        Me.tblGenerationTwoCols.Margin = New Padding(0)
        Me.tblGenerationTwoCols.ColumnCount = 2
        Me.tblGenerationTwoCols.ColumnStyles.Clear()
        Me.tblGenerationTwoCols.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50.0!))
        Me.tblGenerationTwoCols.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50.0!))
        Me.tblGenerationTwoCols.RowCount = 1
        Me.tblGenerationTwoCols.RowStyles.Clear()
        Me.tblGenerationTwoCols.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0!))
        Me.flowGenerationLeft.Dock = DockStyle.Fill
        Me.flowGenerationLeft.Padding = New Padding(5, 4, 4, 4)
        Me.flowGenerationRight.Dock = DockStyle.Fill
        Me.flowGenerationRight.Padding = New Padding(4, 4, 5, 4)
        Me.grpGeneration.Padding = New Padding(5, 8, 5, 5)
        Me.grpTemplates.Padding = New Padding(5, 8, 5, 5)
        Me.flowGenerationLeft.FlowDirection = FlowDirection.TopDown
        Me.flowGenerationLeft.WrapContents = False
        Me.flowGenerationLeft.AutoScroll = True
        Me.flowGenerationRight.FlowDirection = FlowDirection.TopDown
        Me.flowGenerationRight.WrapContents = False
        Me.flowGenerationRight.AutoScroll = True
        Dim gc As CheckBox() = {Me.chkCreateDft, Me.chkCreatePdf, Me.chkCreateDxfDraft, Me.chkAutoDimensioning, Me.chkPmiRetrievalProbe, Me.chkExperimentalPmiModelView, Me.chkCreateFlatDxf, Me.chkOpenOutput, Me.chkOverwrite, Me.chkUniqueComponents, Me.chkDetailedLog, Me.chkDebugTemplates, Me.chkExperimentalDraftGeometryDiagnostics, Me.chkKeepSolidEdgeVisible, Me.chkInsertProperties}
        Dim gt As String() = {"Crear DFT", "Crear PDF", "Crear DXF del DFT", "Generar acotado automático", "Probar recuperación PMI en Draft (experimental)", "Experimental: crear PMIModelView si falta (sin guardar modelo)", "Crear DXF de chapa desarrollada", "Abrir carpeta de salida al finalizar", "Sobrescribir archivos existentes", "Procesar componentes repetidos solo una vez", "Registrar log detallado", "Diagnóstico: inspeccionar plantillas DFT (PropertySets, cajetín)", "Diagnóstico geométrico Draft (2D, vistas, motor acotado)", "Mantener visible Solid Edge durante el proceso", "Insertar propiedades en el cajetin del draft, si procede"}
        For i As Integer = 0 To gc.Length - 1
            gc(i).Text = gt(i) : gc(i).AutoSize = True
            If i <> 4 AndAlso i <> 5 AndAlso i <> 8 AndAlso i <> 11 AndAlso i <> 12 AndAlso i <> 14 Then gc(i).Checked = True
            If i <= 7 Then
                Me.flowGenerationLeft.Controls.Add(gc(i))
            Else
                Me.flowGenerationRight.Controls.Add(gc(i))
            End If
        Next
        Me.lblTitleBlockSource.Text = "Origen propiedades cajetín"
        Me.lblTitleBlockSource.AutoSize = True
        Me.cmbTitleBlockSource.DropDownStyle = ComboBoxStyle.DropDownList
        Me.cmbTitleBlockSource.Width = 220
        Me.flowGenerationLeft.Controls.Add(Me.lblTitleBlockSource)
        Me.flowGenerationLeft.Controls.Add(Me.cmbTitleBlockSource)
        Me.pnlAutoDimensionMotorChoice.FlowDirection = FlowDirection.TopDown
        Me.pnlAutoDimensionMotorChoice.WrapContents = False
        Me.pnlAutoDimensionMotorChoice.AutoSize = True
        Me.pnlAutoDimensionMotorChoice.AutoSizeMode = AutoSizeMode.GrowAndShrink
        Me.pnlAutoDimensionMotorChoice.Margin = New Padding(0, 2, 0, 0)
        Me.lblAutoDimensionMotor.AutoSize = True
        Me.lblAutoDimensionMotor.Margin = New Padding(0, 0, 0, 2)
        Me.lblAutoDimensionMotor.Text = "Motor DV (solo uno activo):"
        Me.radAutoDimMotorMain.AutoSize = True
        Me.radAutoDimMotorMain.Text = "Principal (UniqueDv + UNE/ISO 129)"
        Me.radAutoDimMotorMain.Checked = True
        Me.radAutoDimMotorLegacyV02.AutoSize = True
        Me.radAutoDimMotorLegacyV02.Text = "Copia V02 aislada (LegacyV02Dimensioning)"
        Me.radAutoDimMotorAlternatePlugIn.AutoSize = True
        Me.radAutoDimMotorAlternatePlugIn.Text = "Plugin alternativo (enganche propio, sin motor principal)"
        Me.pnlAutoDimensionMotorChoice.Controls.Add(Me.lblAutoDimensionMotor)
        Me.pnlAutoDimensionMotorChoice.Controls.Add(Me.radAutoDimMotorMain)
        Me.pnlAutoDimensionMotorChoice.Controls.Add(Me.radAutoDimMotorLegacyV02)
        Me.pnlAutoDimensionMotorChoice.Controls.Add(Me.radAutoDimMotorAlternatePlugIn)
        Me.flowGenerationLeft.Controls.Add(Me.pnlAutoDimensionMotorChoice)
        Me.chkSesdkPostDimensionIntrospection.AutoSize = True
        Me.chkSesdkPostDimensionIntrospection.Text = "Tras acotado: volcar introspección SDK (DV + cotas, log [SESDK_PROBE])"
        Me.chkSesdkPostDimensionIntrospection.Checked = False
        Me.chkSesdkPostDimensionIntrospection.Margin = New Padding(18, 0, 0, 0)
        Me.flowGenerationLeft.Controls.Add(Me.chkSesdkPostDimensionIntrospection)
        Me.chkPreferSweepAllDrawingDimensions.AutoSize = True
        Me.chkPreferSweepAllDrawingDimensions.Text = "Más cotas DV: barrer líneas, arcos y círculos (modo SweepAll; puede saturar el pliego)"
        Me.chkPreferSweepAllDrawingDimensions.Checked = False
        Me.chkPreferSweepAllDrawingDimensions.Margin = New Padding(18, 0, 0, 0)
        Me.flowGenerationLeft.Controls.Add(Me.chkPreferSweepAllDrawingDimensions)
        Me.chkSuppressDimTrackSpacing.AutoSize = True
        Me.chkSuppressDimTrackSpacing.Text = "Sin espaciado 12/10 mm (TrackDistance y carriles desactivados)"
        Me.chkSuppressDimTrackSpacing.Checked = False
        Me.chkSuppressDimTrackSpacing.Margin = New Padding(18, 0, 0, 0)
        Me.flowGenerationLeft.Controls.Add(Me.chkSuppressDimTrackSpacing)
        Me.chkDedupDimensionsByKeypoints.AutoSize = True
        Me.chkDedupDimensionsByKeypoints.Text = "Quitar cotas duplicadas (mismo valor y keypoints)"
        Me.chkDedupDimensionsByKeypoints.Checked = True
        Me.chkDedupDimensionsByKeypoints.Margin = New Padding(18, 0, 0, 0)
        Me.flowGenerationLeft.Controls.Add(Me.chkDedupDimensionsByKeypoints)
        Me.flowGenerationLeft.Controls.Add(Me.chkUnitHorizontalExteriorTest)
        Me.chkDrawingViewDimensioningLab.AutoSize = True
        Me.chkDrawingViewDimensioningLab.Text = ""
        Me.chkDrawingViewDimensioningLab.Checked = False
        Me.flowGenerationRight.Controls.Add(Me.chkDrawingViewDimensioningLab)
        Me.chkDimLabInteractivePause.AutoSize = True
        Me.chkDimLabInteractivePause.Text = ""
        Me.chkDimLabInteractivePause.Checked = True
        Me.flowGenerationRight.Controls.Add(Me.chkDimLabInteractivePause)
        Me.lblDimLabMode.AutoSize = True
        Me.lblDimLabMode.Text = ""
        Me.flowGenerationRight.Controls.Add(Me.lblDimLabMode)
        Me.cmbDimLabMode.DropDownStyle = ComboBoxStyle.DropDownList
        Me.cmbDimLabMode.Width = 280
        Me.cmbDimLabMode.Items.AddRange(New Object() {"HorizontalOnly", "VerticalOnly", "Full", "ForensicHorizontal", "CleanFull", "CleanFullStrict"})
        Me.cmbDimLabMode.SelectedIndex = 2
        Me.flowGenerationRight.Controls.Add(Me.cmbDimLabMode)
        Me.chkDimLabVisibleProbe.AutoSize = True
        Me.chkDimLabVisibleProbe.Text = ""
        Me.chkDimLabVisibleProbe.Checked = False
        Me.flowGenerationRight.Controls.Add(Me.chkDimLabVisibleProbe)
        Me.chkDimLabAlternativePlacement.AutoSize = True
        Me.chkDimLabAlternativePlacement.Text = ""
        Me.chkDimLabAlternativePlacement.Checked = False
        Me.flowGenerationRight.Controls.Add(Me.chkDimLabAlternativePlacement)
        Me.btnDimLabRun.AutoSize = True
        Me.btnDimLabRun.Text = ""
        Me.btnDimLabRun.Margin = New Padding(0, 4, 0, 0)
        Me.flowGenerationRight.Controls.Add(Me.btnDimLabRun)
        Me.tblGenerationTwoCols.Controls.Add(Me.flowGenerationLeft, 0, 0)
        Me.tblGenerationTwoCols.Controls.Add(Me.flowGenerationRight, 1, 0)
        Me.grpGeneration.Controls.Add(Me.tblGenerationTwoCols)

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
        Me.pnlMotorViewsAdvancedHost.Controls.Add(Me.tblAdvanced)
        Me.grpAdvanced.Visible = False
        Me.grpAdvanced.Size = New Size(1, 1)

        Me.btnGenerate.Text = "GENERAR"
        Me.btnGenerate.AutoSize = True
        Me.btnGenerate.MinimumSize = New Size(200, 50)
        Me.btnGenerate.Height = 50
        Me.btnGenerate.Margin = New Padding(0, 0, 8, 6)
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
        Me.mainLayout.Controls.Add(Me.bodyHost, 0, 1)
        Me.Controls.Add(Me.mainLayout)

        Me.AutoScaleMode = AutoScaleMode.Font
        Me.ClientSize = New Size(1400, 900)
        Me.Font = New Font("Segoe UI", 9.0!)
        Me.MinimumSize = New Size(1280, 820)
        Me.Name = "MainForm"
        Me.StartPosition = FormStartPosition.CenterScreen

        CType(Me.dgvAsmComponents, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvLaserPieces, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
    End Sub

    Friend WithEvents mainLayout As TableLayoutPanel
    Friend WithEvents pnlHeader As Panel
    Friend WithEvents lblTitle As Label
    Friend WithEvents lblSubTitle As Label
    Friend WithEvents bodyHost As TableLayoutPanel
    Friend WithEvents twoColumnBodyLayout As TableLayoutPanel
    Friend WithEvents pnlSharedSidebar As TableLayoutPanel
    Friend WithEvents tabMotors As TabControl
    Friend WithEvents tabPageMotorViews As TabPage
    Friend WithEvents tabPageMotorMetadata As TabPage
    Friend WithEvents tabPageMotorDimensions As TabPage
    Friend WithEvents tabPageLaserCut As TabPage
    Friend WithEvents tblLaserTabHost As TableLayoutPanel
    Friend WithEvents tblLaserTabRightColumn As TableLayoutPanel
    Friend WithEvents grpLaserPieces As GroupBox
    Friend WithEvents tblLaserPiecesInner As TableLayoutPanel
    Friend WithEvents flowLaserButtons As FlowLayoutPanel
    Friend WithEvents btnLaserScan As Button
    Friend WithEvents btnLaserGenerate As Button
    Friend WithEvents lblLaserSummary As Label
    Friend WithEvents dgvLaserPieces As DataGridView
    Friend WithEvents btnMotorLaser As Button
    Friend WithEvents tblMetadataTabHost As TableLayoutPanel
    Friend WithEvents tblMetadataTabRightColumn As TableLayoutPanel
    Friend WithEvents tblMetadataPlanTwoCols As TableLayoutPanel
    Friend WithEvents tblDimensionTabHost As TableLayoutPanel
    Friend WithEvents tblDimensionTabRightColumn As TableLayoutPanel
    Friend WithEvents tblMotorTab1Host As TableLayoutPanel
    Friend WithEvents tblMotorTab1LeftColumn As TableLayoutPanel
    Friend WithEvents tblMotorTab1RightColumn As TableLayoutPanel
    Friend WithEvents tblMotorRightTemplatesFormat As TableLayoutPanel
    Friend WithEvents pnlMotorViewsAdvancedHost As Panel
    Friend WithEvents grpPdfPreview As GroupBox
    Friend WithEvents tblPdfPreviewInner As TableLayoutPanel
    Friend WithEvents flowPdfPreviewBar As FlowLayoutPanel
    Friend WithEvents btnRefreshPdfPreview As Button
    Friend WithEvents btnOpenLastPdfExternal As Button
    Friend WithEvents wbPdfPreview As WebBrowser
    Friend WithEvents pnlGenerateBar As Panel
    Friend WithEvents flowGenerateBar As FlowLayoutPanel
    Friend WithEvents btnMotorViews As Button
    Friend WithEvents btnMotorMetadata As Button
    Friend WithEvents btnMotorDimensioning As Button
    Friend WithEvents btnOpenVariableTable As Button
    Friend WithEvents tblRightLogProgress As TableLayoutPanel
    Friend WithEvents grpInput As GroupBox
    Friend WithEvents grpAsmComponents As GroupBox
    Friend WithEvents grpTemplates As GroupBox
    Friend WithEvents grpTraceability As GroupBox
    Friend WithEvents grpGeneration As GroupBox
    Friend WithEvents grpAdvanced As GroupBox
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
    Friend WithEvents tblGenerationTwoCols As TableLayoutPanel
    Friend WithEvents flowGenerationLeft As FlowLayoutPanel
    Friend WithEvents flowGenerationRight As FlowLayoutPanel
    Friend WithEvents lblTitleBlockSource As Label
    Friend WithEvents cmbTitleBlockSource As ComboBox
    Friend WithEvents chkCreateDft As CheckBox
    Friend WithEvents chkCreatePdf As CheckBox
    Friend WithEvents chkCreateDxfDraft As CheckBox
    Friend WithEvents chkAutoDimensioning As CheckBox
    Friend WithEvents pnlAutoDimensionMotorChoice As FlowLayoutPanel
    Friend WithEvents lblAutoDimensionMotor As Label
    Friend WithEvents radAutoDimMotorMain As RadioButton
    Friend WithEvents radAutoDimMotorLegacyV02 As RadioButton
    Friend WithEvents radAutoDimMotorAlternatePlugIn As RadioButton
    Friend WithEvents chkSesdkPostDimensionIntrospection As CheckBox
    Friend WithEvents chkPreferSweepAllDrawingDimensions As CheckBox
    Friend WithEvents chkSuppressDimTrackSpacing As CheckBox
    Friend WithEvents chkDedupDimensionsByKeypoints As CheckBox
    Friend WithEvents chkUnitHorizontalExteriorTest As CheckBox
    Friend WithEvents chkDrawingViewDimensioningLab As CheckBox
    Friend WithEvents chkDimLabInteractivePause As CheckBox
    Friend WithEvents lblDimLabMode As Label
    Friend WithEvents cmbDimLabMode As ComboBox
    Friend WithEvents chkDimLabVisibleProbe As CheckBox
    Friend WithEvents chkDimLabAlternativePlacement As CheckBox
    Friend WithEvents btnDimLabRun As Button
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
