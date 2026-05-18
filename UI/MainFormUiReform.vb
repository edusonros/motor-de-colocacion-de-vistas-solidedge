Option Strict Off

Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Windows.Forms

''' <summary>Reorganización visual según «Reforma UI.docx» (pestaña Opciones, progreso bajo ASM, colores por pestaña).</summary>
Partial Public Class MainForm

    Private _uiReformApplied As Boolean = False
    Private _tblMotorViewsRightSplit As TableLayoutPanel
    Private _tblOptionsTabHost As TableLayoutPanel
    Private _tblOptionsTabRightColumn As TableLayoutPanel
    Private _pnlOptionsScroll As Panel
    Private _pnlOptionsStack As TableLayoutPanel
    Private _pnlMetaButtonsMetadata As Panel
    Private _pnlMetaButtonsOptions As Panel
    Friend tabPageGenerationOptions As TabPage

    Private Shared ReadOnly TabCortenColors As Color() = {
        Color.FromArgb(72, 52, 42),
        Color.FromArgb(88, 58, 48),
        Color.FromArgb(62, 68, 58),
        Color.FromArgb(58, 62, 72),
        Color.FromArgb(48, 52, 56)
    }

    Private Sub ApplyUiReform()
        If _uiReformApplied Then Return
        _uiReformApplied = True
        Try
            EnsureGenerationOptionsTab()
            ConfigureSharedLeftColumn()
            ConfigureMotorViewsRightColumn()
            RelocateOptionsTabContent()
            ConfigureBottomActionBar()
            ApplyIndustrialTabColors()
            UpdateProgressBarExpectations()
            EnsureStopRunButton()
            If _btnStopRun IsNot Nothing Then
                _btnStopRun.Text = "STOP"
                _btnStopRun.MinimumSize = New Size(120, 50)
                _btnStopRun.Height = 50
                _btnStopRun.Margin = New Padding(12, 0, 0, 0)
                _btnStopRun.BackColor = Color.FromArgb(168, 42, 42)
            End If
            _lastAppliedMotorChromeTabIndex = -1
            ApplySharedSidebarDockingForActiveTab()
        Catch ex As Exception
            Try
                _logger?.LogException("ApplyUiReform", ex)
            Catch
            End Try
        End Try
    End Sub

    Private Sub EnsureGenerationOptionsTab()
        If tabMotors Is Nothing Then Return
        If tabPageGenerationOptions IsNot Nothing Then Return

        tabPageGenerationOptions = New TabPage With {.Text = "Opciones de Configuracion", .Padding = New Padding(4)}
        _tblOptionsTabHost = New TableLayoutPanel With {.Dock = DockStyle.Fill, .ColumnCount = 2, .Margin = Padding.Empty}
        _tblOptionsTabHost.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50.0F))
        _tblOptionsTabHost.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50.0F))
        _tblOptionsTabHost.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))

        _tblOptionsTabRightColumn = New TableLayoutPanel With {.Dock = DockStyle.Fill, .ColumnCount = 1, .RowCount = 1, .Margin = New Padding(4, 0, 0, 0)}
        _tblOptionsTabRightColumn.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))

        _pnlOptionsScroll = New Panel With {.Dock = DockStyle.Fill, .AutoScroll = True, .Padding = New Padding(4)}
        _pnlOptionsStack = New TableLayoutPanel With {
            .Dock = DockStyle.Top,
            .AutoSize = True,
            .AutoSizeMode = AutoSizeMode.GrowAndShrink,
            .ColumnCount = 1,
            .Padding = Padding.Empty,
            .MinimumSize = New Size(520, 0)
        }
        _pnlOptionsStack.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
        _pnlOptionsScroll.Controls.Add(_pnlOptionsStack)
        _tblOptionsTabRightColumn.Controls.Add(_pnlOptionsScroll, 0, 0)
        _tblOptionsTabHost.Controls.Add(_tblOptionsTabRightColumn, 1, 0)
        tabPageGenerationOptions.Controls.Add(_tblOptionsTabHost)
        tabMotors.Controls.Add(tabPageGenerationOptions)
    End Sub

    Private Sub ConfigureSharedLeftColumn()
        If tblMotorTab1LeftColumn Is Nothing Then Return
        While tblMotorTab1LeftColumn.Controls.Contains(grpPdfPreview)
            tblMotorTab1LeftColumn.Controls.Remove(grpPdfPreview)
        End While
        tblMotorTab1LeftColumn.RowCount = 3
        tblMotorTab1LeftColumn.RowStyles.Clear()
        tblMotorTab1LeftColumn.RowStyles.Add(New RowStyle(SizeType.Absolute, 108.0F))
        tblMotorTab1LeftColumn.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
        tblMotorTab1LeftColumn.RowStyles.Add(New RowStyle(SizeType.Absolute, 132.0F))
        ConfigureAsmComponentsInnerLayout()
        If grpInput IsNot Nothing AndAlso Not tblMotorTab1LeftColumn.Controls.Contains(grpInput) Then
            tblMotorTab1LeftColumn.Controls.Add(grpInput, 0, 0)
        End If
        If grpAsmComponents IsNot Nothing AndAlso Not tblMotorTab1LeftColumn.Controls.Contains(grpAsmComponents) Then
            tblMotorTab1LeftColumn.Controls.Add(grpAsmComponents, 0, 1)
        End If
        If grpProgress IsNot Nothing Then
            If tblRightLogProgress IsNot Nothing AndAlso tblRightLogProgress.Controls.Contains(grpProgress) Then
                tblRightLogProgress.Controls.Remove(grpProgress)
            End If
            If Not tblMotorTab1LeftColumn.Controls.Contains(grpProgress) Then
                tblMotorTab1LeftColumn.Controls.Add(grpProgress, 0, 2)
            End If
            grpProgress.Dock = DockStyle.Fill
            grpProgress.Text = "Progreso y estado de la generación"
            grpProgress.Margin = New Padding(4, 2, 6, 4)
        End If
    End Sub

    Private Sub ConfigureMotorViewsRightColumn()
        If tblMotorTab1RightColumn Is Nothing Then Return
        If _tblMotorViewsRightSplit IsNot Nothing Then Return

        While tblMotorTab1RightColumn.Controls.Contains(tblMotorRightTemplatesFormat)
            tblMotorTab1RightColumn.Controls.Remove(tblMotorRightTemplatesFormat)
        End While

        _tblMotorViewsRightSplit = New TableLayoutPanel With {.Dock = DockStyle.Fill, .ColumnCount = 1, .RowCount = 2, .Margin = Padding.Empty}
        _tblMotorViewsRightSplit.RowStyles.Add(New RowStyle(SizeType.Percent, 40.0F))
        _tblMotorViewsRightSplit.RowStyles.Add(New RowStyle(SizeType.Percent, 60.0F))

        If grpLog IsNot Nothing Then
            If tblRightLogProgress IsNot Nothing AndAlso tblRightLogProgress.Controls.Contains(grpLog) Then
                tblRightLogProgress.Controls.Remove(grpLog)
            End If
            grpLog.Dock = DockStyle.Fill
            grpLog.Margin = New Padding(4, 2, 6, 2)
            _tblMotorViewsRightSplit.Controls.Add(grpLog, 0, 0)
        End If
        If grpPdfPreview IsNot Nothing Then
            grpPdfPreview.Dock = DockStyle.Fill
            grpPdfPreview.Margin = New Padding(4, 2, 6, 4)
            _tblMotorViewsRightSplit.Controls.Add(grpPdfPreview, 0, 1)
        End If

        tblMotorTab1RightColumn.RowCount = 1
        tblMotorTab1RightColumn.RowStyles.Clear()
        tblMotorTab1RightColumn.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
        tblMotorTab1RightColumn.Controls.Add(_tblMotorViewsRightSplit, 0, 0)
    End Sub

    Private Sub ConfigureAsmComponentsInnerLayout()
        If tblAsmComponents Is Nothing Then Return
        tblAsmComponents.RowStyles.Clear()
        tblAsmComponents.RowCount = 3
        tblAsmComponents.RowStyles.Add(New RowStyle(SizeType.Absolute, 24.0F))
        tblAsmComponents.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
        tblAsmComponents.RowStyles.Add(New RowStyle(SizeType.Absolute, 30.0F))
        If lblAsmComponentHint IsNot Nothing Then lblAsmComponentHint.Margin = New Padding(0, 0, 0, 2)
    End Sub

    Private Sub ApplyMotorViewsTabLayout()
        ConfigureSharedLeftColumn()
        If _tblMotorViewsRightSplit IsNot Nothing Then
            _tblMotorViewsRightSplit.RowStyles(0).Height = 40.0F
            _tblMotorViewsRightSplit.RowStyles(1).Height = 60.0F
        End If
    End Sub

    Private Sub RelocateOptionsTabContent()
        If _pnlOptionsStack Is Nothing Then Return

        If tblDimensionTabRightColumn IsNot Nothing AndAlso grpGeneration IsNot Nothing Then
            tblDimensionTabRightColumn.Controls.Remove(grpGeneration)
        End If

        _templatesEmbedded = False
        EmbedTemplatesInOptionsPanel()

        PrepareTemplatesGroupForOptionsTab()
        PrepareAdvancedGroupForOptionsTab()
        PrepareGenerationGroupForOptionsTab()

        AddOptionsSection(grpTemplates, "Plantillas DFT/DXF (rutas con Examinar…).", 178)
        AddOptionsSection(grpAdvanced, "Formato de hoja, escala y vistas incluidas en el DFT.", 218)
        AddOptionsSection(grpGeneration, "Qué archivos crear (DFT, PDF, DXF) y opciones del motor de acotado.", 420)

        EnsureOptionsUtilityButtonsPanel()
        EnsureMetadataSecondaryButtonsOnOptions()
        EnsureOptionsPickerControlsVisible()
    End Sub

    Private Sub EmbedTemplatesInOptionsPanel()
        Try
            If tblMotorRightTemplatesFormat IsNot Nothing AndAlso grpTemplates IsNot Nothing Then
                If tblMotorRightTemplatesFormat.Controls.Contains(grpTemplates) Then
                    tblMotorRightTemplatesFormat.Controls.Remove(grpTemplates)
                End If
            End If

            If tblAdvanced IsNot Nothing AndAlso grpAdvanced IsNot Nothing Then
                If tblAdvanced.Parent IsNot Nothing AndAlso Not ReferenceEquals(tblAdvanced.Parent, grpAdvanced) Then
                    tblAdvanced.Parent.Controls.Remove(tblAdvanced)
                End If
                If pnlMotorViewsAdvancedHost IsNot Nothing AndAlso pnlMotorViewsAdvancedHost.Controls.Contains(tblAdvanced) Then
                    pnlMotorViewsAdvancedHost.Controls.Remove(tblAdvanced)
                End If
                grpAdvanced.Controls.Clear()
                grpAdvanced.Controls.Add(tblAdvanced)
                grpAdvanced.Visible = True
                grpAdvanced.AutoSize = False
                grpAdvanced.MinimumSize = New Size(420, 180)
                grpAdvanced.Size = New Size(480, 200)
                tblAdvanced.Dock = DockStyle.Fill
                tblAdvanced.Visible = True
            End If

            _templatesEmbedded = False
        Catch ex As Exception
            _logger?.LogException("EmbedTemplatesInOptionsPanel", ex)
        End Try
    End Sub

    Private Sub PrepareTemplatesGroupForOptionsTab()
        If grpTemplates Is Nothing OrElse tblTemplates Is Nothing Then Return
        grpTemplates.Text = "Templates (A4 / A3 / A2 / DXF limpio)"
        grpTemplates.Visible = True
        grpTemplates.AutoSize = False
        grpTemplates.Padding = New Padding(8, 10, 8, 8)
        If Not grpTemplates.Controls.Contains(tblTemplates) Then
            grpTemplates.Controls.Add(tblTemplates)
        End If
        tblTemplates.Dock = DockStyle.Fill
        tblTemplates.Visible = True
        tblTemplates.ColumnStyles(2).SizeType = SizeType.Absolute
        tblTemplates.ColumnStyles(2).Width = 88.0F
        For Each btn As Button In New Button() {btnBrowseA4, btnBrowseA3, btnBrowseA2, btnBrowseDxf}
            If btn Is Nothing Then Continue For
            btn.Visible = True
            btn.Enabled = True
            btn.Text = "Examinar..."
            btn.Dock = DockStyle.Fill
        Next
        _templatesEmbedded = True
    End Sub

    Private Sub PrepareAdvancedGroupForOptionsTab()
        If grpAdvanced Is Nothing OrElse tblAdvanced Is Nothing Then Return
        Try
            If tblAdvanced.Parent IsNot Nothing AndAlso Not ReferenceEquals(tblAdvanced.Parent, grpAdvanced) Then
                tblAdvanced.Parent.Controls.Remove(tblAdvanced)
            End If
            grpAdvanced.Controls.Clear()
            grpAdvanced.Controls.Add(tblAdvanced)
            grpAdvanced.Text = "Formato y procesado"
            grpAdvanced.Visible = True
            grpAdvanced.AutoSize = False
            grpAdvanced.Padding = New Padding(8, 10, 8, 8)
            tblAdvanced.Dock = DockStyle.Fill
            tblAdvanced.Visible = True
            tblAdvanced.ColumnStyles(2).SizeType = SizeType.Absolute
            tblAdvanced.ColumnStyles(2).Width = 100.0F
            _templatesEmbedded = True
        Catch ex As Exception
            _logger?.LogException("PrepareAdvancedGroupForOptionsTab", ex)
        End Try
    End Sub

    Private Sub PrepareGenerationGroupForOptionsTab()
        If grpGeneration Is Nothing OrElse tblGenerationTwoCols Is Nothing Then Return
        grpGeneration.Visible = True
        grpGeneration.AutoSize = False
        grpGeneration.Padding = New Padding(8, 10, 8, 8)
        If Not grpGeneration.Controls.Contains(tblGenerationTwoCols) Then
            grpGeneration.Controls.Add(tblGenerationTwoCols)
        End If
        tblGenerationTwoCols.Dock = DockStyle.Fill
        tblGenerationTwoCols.Visible = True
    End Sub

    Private Sub EnsureOptionsPickerControlsVisible()
        For Each btn As Button In New Button() {btnBrowseA4, btnBrowseA3, btnBrowseA2, btnBrowseDxf, btnBrowseInput, btnBrowseOut}
            If btn IsNot Nothing Then
                btn.Visible = True
                btn.Enabled = True
            End If
        Next
        For Each txt As TextBox In New TextBox() {txtTemplateA4, txtTemplateA3, txtTemplateA2, txtTemplateDxf}
            If txt IsNot Nothing Then txt.Visible = True
        Next
    End Sub

    Private Sub AddOptionsSection(gb As GroupBox, helpText As String, minHeightPx As Integer)
        If gb Is Nothing OrElse _pnlOptionsStack Is Nothing Then Return
        DetachControlFromParent(gb)

        Dim row As Integer = _pnlOptionsStack.RowCount
        _pnlOptionsStack.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        _pnlOptionsStack.RowCount = row + 1

        Dim host As New TableLayoutPanel With {
            .Dock = DockStyle.Top,
            .AutoSize = False,
            .ColumnCount = 2,
            .Margin = New Padding(0, 0, 0, 10),
            .MinimumSize = New Size(500, minHeightPx + 8),
            .Height = minHeightPx + 12
        }
        host.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
        host.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 36.0F))

        gb.Dock = DockStyle.Fill
        gb.AutoSize = False
        gb.Margin = New Padding(0)
        gb.MinimumSize = New Size(460, minHeightPx)
        gb.Height = minHeightPx
        gb.Visible = True

        Dim gbHost As New Panel With {.Dock = DockStyle.Fill, .MinimumSize = New Size(460, minHeightPx), .Height = minHeightPx}
        gbHost.Controls.Add(gb)

        host.Controls.Add(gbHost, 0, 0)
        host.Controls.Add(CreateComicHelpButton(helpText), 1, 0)
        host.SetRowSpan(gbHost, 1)
        _pnlOptionsStack.Controls.Add(host, 0, row)
    End Sub

    Private Shared Sub DetachControlFromParent(ctrl As Control)
        If ctrl Is Nothing Then Return
        Dim p = ctrl.Parent
        If p IsNot Nothing Then p.Controls.Remove(ctrl)
    End Sub

    Private Function IsGroupBoxInOptionsStack(gb As GroupBox) As Boolean
        If gb Is Nothing OrElse _pnlOptionsStack Is Nothing Then Return False
        Dim p As Control = gb.Parent
        While p IsNot Nothing
            If ReferenceEquals(p, _pnlOptionsStack) Then Return True
            p = p.Parent
        End While
        Return False
    End Function

    Private Sub RemoveGroupBoxFromOptionsStack(gb As GroupBox)
        If gb Is Nothing OrElse _pnlOptionsStack Is Nothing Then Return
        For Each host As Control In _pnlOptionsStack.Controls
            Dim tbl = TryCast(host, TableLayoutPanel)
            If tbl Is Nothing Then Continue For
            For Each child As Control In tbl.Controls
                Dim wrap = TryCast(child, Panel)
                If wrap IsNot Nothing AndAlso wrap.Controls.Contains(gb) Then
                    _pnlOptionsStack.Controls.Remove(host)
                    Return
                End If
            Next
        Next
    End Sub

    Friend Sub MoveGrpInputToLeftColumn()
        If grpInput Is Nothing OrElse tblMotorTab1LeftColumn Is Nothing Then Return
        RemoveGroupBoxFromOptionsStack(grpInput)
        DetachControlFromParent(grpInput)
        If Not tblMotorTab1LeftColumn.Controls.Contains(grpInput) Then
            tblMotorTab1LeftColumn.Controls.Add(grpInput, 0, 0)
        End If
        grpInput.Dock = DockStyle.Fill
        grpInput.Visible = True
        If tblInput IsNot Nothing Then
            tblInput.ColumnStyles(2).SizeType = SizeType.Absolute
            tblInput.ColumnStyles(2).Width = 100.0F
        End If
        If btnBrowseInput IsNot Nothing Then btnBrowseInput.Visible = True
        If btnBrowseOut IsNot Nothing Then btnBrowseOut.Visible = True
    End Sub

    Friend Sub MoveGrpInputToOptionsStack()
        If grpInput Is Nothing OrElse _pnlOptionsStack Is Nothing Then Return
        If IsGroupBoxInOptionsStack(grpInput) Then Return
        If tblMotorTab1LeftColumn IsNot Nothing AndAlso tblMotorTab1LeftColumn.Controls.Contains(grpInput) Then
            tblMotorTab1LeftColumn.Controls.Remove(grpInput)
        End If
        AddOptionsSection(grpInput, "Archivo ASM de entrada y carpeta de salida (Examinar…).", 108)
    End Sub

    Private Sub EnsureOptionsUtilityButtonsPanel()
        If _pnlOptionsStack Is Nothing Then Return
        Dim row As Integer = _pnlOptionsStack.RowCount
        _pnlOptionsStack.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        _pnlOptionsStack.RowCount = row + 1

        Dim flow As New FlowLayoutPanel With {
            .Dock = DockStyle.Top,
            .AutoSize = True,
            .FlowDirection = FlowDirection.LeftToRight,
            .WrapContents = True,
            .Padding = New Padding(0, 4, 0, 8)
        }

        Dim nav As New FlowLayoutPanel With {.AutoSize = True, .FlowDirection = FlowDirection.LeftToRight, .WrapContents = True}
        If btnMotorViews IsNot Nothing Then nav.Controls.Add(btnMotorViews)
        If btnMotorMetadata IsNot Nothing Then nav.Controls.Add(btnMotorMetadata)
        If btnMotorDimensioning IsNot Nothing Then nav.Controls.Add(btnMotorDimensioning)
        If btnMotorLaser IsNot Nothing Then nav.Controls.Add(btnMotorLaser)
        If btnOpenVariableTable IsNot Nothing Then nav.Controls.Add(btnOpenVariableTable)

        Dim util As New FlowLayoutPanel With {.AutoSize = True, .FlowDirection = FlowDirection.LeftToRight, .WrapContents = True}
        btnClear.Visible = True : btnOpenOutput.Visible = True
        btnSaveConfig.Visible = True : btnLoadConfig.Visible = True
        btnReloadSourceProps.Visible = True : btnRestoreDefaultTemplates.Visible = True
        util.Controls.Add(btnClear)
        util.Controls.Add(btnOpenOutput)
        util.Controls.Add(btnSaveConfig)
        util.Controls.Add(btnLoadConfig)
        util.Controls.Add(btnReloadSourceProps)
        util.Controls.Add(btnRestoreDefaultTemplates)

        flow.Controls.Add(nav)
        flow.Controls.Add(New Label With {.Text = "  ", .AutoSize = True})
        flow.Controls.Add(util)
        flow.Controls.Add(CreateComicHelpButton("Accesos rápidos a otras pestañas y utilidades de configuración que antes estaban en la barra inferior."))

        _pnlOptionsStack.Controls.Add(flow, 0, row)
    End Sub

    Private Sub EnsureMetadataSecondaryButtonsOnOptions()
        If _pnlMetaButtonsOptions IsNot Nothing Then Return
        If pnlPlanMetaActions Is Nothing Then Return

        _pnlMetaButtonsMetadata = New Panel With {.Dock = DockStyle.Top, .Height = 44, .Padding = New Padding(0, 4, 0, 0)}
        _pnlMetaButtonsOptions = New Panel With {.Dock = DockStyle.Top, .AutoSize = True, .Padding = New Padding(0, 6, 0, 4)}

        Dim tblMd As New TableLayoutPanel With {.Dock = DockStyle.Fill, .ColumnCount = 2, .RowCount = 1}
        tblMd.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50.0F))
        tblMd.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50.0F))
        If btnMetaPickDftLoad IsNot Nothing Then
            RemoveButtonFromParents(btnMetaPickDftLoad)
            btnMetaPickDftLoad.Dock = DockStyle.Fill
            btnMetaPickDftLoad.Margin = New Padding(0, 0, 4, 0)
            tblMd.Controls.Add(btnMetaPickDftLoad, 0, 0)
        End If
        If btnApplyTraceability IsNot Nothing Then
            RemoveButtonFromParents(btnApplyTraceability)
            btnApplyTraceability.Dock = DockStyle.Fill
            btnApplyTraceability.Margin = New Padding(4, 0, 0, 0)
            tblMd.Controls.Add(btnApplyTraceability, 1, 0)
        End If
        _pnlMetaButtonsMetadata.Controls.Add(tblMd)

        pnlPlanMetaActions.Controls.Clear()
        Dim secHost As New TableLayoutPanel With {.Dock = DockStyle.Top, .AutoSize = True, .ColumnCount = 2}
        secHost.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0F))
        secHost.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, 34.0F))
        Dim flowOpt As New FlowLayoutPanel With {.Dock = DockStyle.Fill, .AutoSize = True, .WrapContents = True}
        For Each b As Button In New Button() {btnMetaReadModel, btnMetaFillEmpty, btnMetaApplyDocs, btnMetaPreview}
            If b Is Nothing Then Continue For
            RemoveButtonFromParents(b)
            b.AutoSize = True
            b.Margin = New Padding(0, 0, 8, 4)
            flowOpt.Controls.Add(b)
        Next
        secHost.Controls.Add(flowOpt, 0, 0)
        secHost.Controls.Add(CreateComicHelpButton("Metadatos avanzados. En Metadatos solo quedan elegir DFT y aplicar cajetín + PART_LIST."), 1, 0)
        _pnlMetaButtonsOptions.Controls.Add(secHost)

        If _pnlOptionsStack IsNot Nothing Then
            Dim row As Integer = _pnlOptionsStack.RowCount
            _pnlOptionsStack.RowStyles.Add(New RowStyle(SizeType.AutoSize))
            _pnlOptionsStack.RowCount = row + 1
            _pnlOptionsStack.Controls.Add(_pnlMetaButtonsOptions, 0, row)
        End If
    End Sub

    Private Shared Sub RemoveButtonFromParents(btn As Button)
        If btn Is Nothing Then Return
        Dim p As Control = btn.Parent
        If p IsNot Nothing Then p.Controls.Remove(btn)
    End Sub

    Private Sub ConfigureBottomActionBar()
        If flowGenerateBar Is Nothing Then Return
        flowGenerateBar.Controls.Clear()
        flowGenerateBar.WrapContents = False
        flowGenerateBar.FlowDirection = FlowDirection.LeftToRight
        Dim spacerL As New Panel With {.Width = 40, .Height = 1, .Margin = Padding.Empty}
        Dim spacerR As New Panel With {.Width = 40, .Height = 1, .Margin = Padding.Empty}
        flowGenerateBar.Controls.Add(spacerL)
        EnsureSolidEdgeQuickAccessButtons()
        If _btnBringSolidEdgeFront IsNot Nothing Then
            _btnBringSolidEdgeFront.Margin = New Padding(0, 0, 8, 0)
            _btnBringSolidEdgeFront.Height = 44
            flowGenerateBar.Controls.Add(_btnBringSolidEdgeFront)
        End If
        If _btnToggleSolidEdgeVisible IsNot Nothing Then
            _btnToggleSolidEdgeVisible.Margin = New Padding(0, 0, 12, 0)
            _btnToggleSolidEdgeVisible.Height = 44
            flowGenerateBar.Controls.Add(_btnToggleSolidEdgeVisible)
        End If
        If btnGenerate IsNot Nothing Then
            btnGenerate.MinimumSize = New Size(220, 52)
            btnGenerate.Height = 52
            btnGenerate.Margin = New Padding(0)
            flowGenerateBar.Controls.Add(btnGenerate)
        End If
        EnsureStopRunButton()
        If _btnStopRun IsNot Nothing Then flowGenerateBar.Controls.Add(_btnStopRun)
        flowGenerateBar.Controls.Add(spacerR)
    End Sub

    Private Sub ApplyIndustrialTabColors()
        If tabMotors Is Nothing Then Return
        tabMotors.DrawMode = TabDrawMode.OwnerDrawFixed
        RemoveHandler tabMotors.DrawItem, AddressOf tabMotors_DrawItemReform
        AddHandler tabMotors.DrawItem, AddressOf tabMotors_DrawItemReform
        For i As Integer = 0 To Math.Min(tabMotors.TabPages.Count - 1, TabCortenColors.Length - 1)
            tabMotors.TabPages(i).BackColor = TabCortenColors(i)
            tabMotors.TabPages(i).ForeColor = Color.FromArgb(245, 238, 230)
        Next
    End Sub

    Private Sub tabMotors_DrawItemReform(sender As Object, e As DrawItemEventArgs)
        Dim tc = TryCast(sender, TabControl)
        If tc Is Nothing OrElse e.Index < 0 OrElse e.Index >= tc.TabPages.Count Then Return
        Dim page = tc.TabPages(e.Index)
        Dim baseColor = TabCortenColors(Math.Min(e.Index, TabCortenColors.Length - 1))
        Dim back = If((e.State And DrawItemState.Selected) = DrawItemState.Selected, ControlPaint.Light(baseColor, 0.15F), baseColor)
        Using br As New SolidBrush(back)
            e.Graphics.FillRectangle(br, e.Bounds)
        End Using
        Dim flags As TextFormatFlags = TextFormatFlags.HorizontalCenter Or TextFormatFlags.VerticalCenter
        TextRenderer.DrawText(e.Graphics, page.Text, New Font("Segoe UI Semibold", 9.0F, FontStyle.Bold), e.Bounds, Color.FromArgb(248, 242, 235), flags)
        e.DrawFocusRectangle()
    End Sub

    Private Function CreateComicHelpButton(explanation As String) As Button
        Dim btn As New Button With {
            .Text = "?",
            .Width = 30,
            .Height = 28,
            .FlatStyle = FlatStyle.Flat,
            .BackColor = Color.FromArgb(210, 180, 140),
            .ForeColor = Color.FromArgb(48, 32, 24),
            .Font = New Font("Segoe UI", 10.0F, FontStyle.Bold),
            .Margin = New Padding(4, 2, 0, 0),
            .TabStop = False
        }
        Dim tip As New ToolTip With {.IsBalloon = True, .ToolTipTitle = "Ayuda", .AutoPopDelay = 16000}
        tip.SetToolTip(btn, explanation)
        AddHandler btn.Click,
            Sub(s, ev)
                tip.Show(explanation, btn, 0, -btn.Height - 8, 12000)
            End Sub
        Return btn
    End Function

    Private Sub UpdateProgressBarExpectations()
        ExpectedLogLinesPerPiece = 1000
        If lblProgressPiece IsNot Nothing Then lblProgressPiece.Text = "Pieza (log ~1000 líneas)"
        If progressBar IsNot Nothing Then progressBar.Maximum = 1000
    End Sub

    Private Sub AttachLogOnlyToHost(host As TableLayoutPanel)
        If host Is Nothing OrElse grpLog Is Nothing Then Return
        host.Controls.Clear()
        host.ColumnCount = 1
        host.RowCount = 1
        host.RowStyles.Clear()
        host.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
        If grpLog.Parent IsNot Nothing Then grpLog.Parent.Controls.Remove(grpLog)
        grpLog.Dock = DockStyle.Fill
        host.Controls.Add(grpLog, 0, 0)
    End Sub

    Private Sub AttachMetadataLayoutToHost(host As TableLayoutPanel)
        If host Is Nothing Then Return
        host.Controls.Clear()
        host.ColumnCount = 1
        host.RowCount = 3
        host.RowStyles.Clear()
        host.RowStyles.Add(New RowStyle(SizeType.Percent, 52.0F))
        host.RowStyles.Add(New RowStyle(SizeType.Absolute, 44.0F))
        host.RowStyles.Add(New RowStyle(SizeType.Percent, 48.0F))

        If tblMetadataTabRightColumn IsNot Nothing Then
            tblMetadataTabRightColumn.RowCount = 1
            tblMetadataTabRightColumn.RowStyles.Clear()
            tblMetadataTabRightColumn.RowStyles.Add(New RowStyle(SizeType.Percent, 100.0F))
        End If
        If tblMetadataPlanTwoCols IsNot Nothing Then
            If tblMetadataPlanTwoCols.Parent IsNot Nothing Then tblMetadataPlanTwoCols.Parent.Controls.Remove(tblMetadataPlanTwoCols)
            tblMetadataPlanTwoCols.Dock = DockStyle.Fill
            host.Controls.Add(tblMetadataPlanTwoCols, 0, 0)
        End If
        If _pnlMetaButtonsMetadata IsNot Nothing Then
            host.Controls.Add(_pnlMetaButtonsMetadata, 0, 1)
        End If
        If grpLog IsNot Nothing Then
            If grpLog.Parent IsNot Nothing Then grpLog.Parent.Controls.Remove(grpLog)
            grpLog.Dock = DockStyle.Fill
            host.Controls.Add(grpLog, 0, 2)
        End If
    End Sub

End Class
