Option Strict Off

Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports System.IO
Imports System.Linq
Imports System.Windows.Forms

''' <summary>Metadatos de plano unificados en <see cref="tblTrace"/> (cajetín + PART_LIST + acciones).</summary>
Partial Public Class MainForm

    Private _unifiedPlanMetadataDone As Boolean
    Private _suspendDenomSync As Boolean
    Private tblTracePartInner As TableLayoutPanel

    Friend grpPlanCajetinBox As GroupBox
    Friend grpPlanPartListBox As GroupBox

    Private Const MetaEditorWidthPx As Integer = 156
    Private Const MetaLabelColWidthPx As Integer = 122

    Private dtpFechaPlano As DateTimePicker
    Private lblOriginFecha As Label
    Private txtPartNum As TextBox
    Private lblOriNum As Label
    Private txtPartCant As TextBox
    Private txtPartNombreArchivo As TextBox
    Private txtPartDenominacion As TextBox
    Private txtPartL As TextBox
    Private txtPartH As TextBox
    Private txtPartD As TextBox
    Private txtPartPeso As TextBox

    Private lblOriCant As Label
    Private lblOriNomArch As Label
    Private lblOriDenom As Label
    Private lblOriL As Label
    Private lblOriH As Label
    Private lblOriD As Label
    Private lblOriPeso As Label

    Friend chkStrictMetadata As CheckBox

    Private btnMetaReadModel As Button
    Private btnMetaCalcLhd As Button
    Private btnMetaFillEmpty As Button
    Private btnMetaApplyDocs As Button
    Private btnMetaPreview As Button

    ''' <summary>Evita marcar campos como «manuales» durante limpieza o carga programática.</summary>
    Private _loadingMetadataProgrammatically As Boolean

    Private Sub EnsureDrawingPlanMetadataPanel()
        If _unifiedPlanMetadataDone Then Return
        _unifiedPlanMetadataDone = True
        grpTraceability.Text = "Metadatos (plano y pieza)"
        grpTraceability.Padding = New Padding(4, 6, 4, 6)

        tblTrace.SuspendLayout()
        Try
            tblTrace.Controls.Remove(flowTraceButtons)
            If flowTraceButtons.Controls.Contains(btnApplyTraceability) Then
                flowTraceButtons.Controls.Remove(btnApplyTraceability)
            End If
            flowTraceButtons.Visible = False

            Dim tblCaj As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .AutoScroll = True,
                .ColumnCount = 3,
                .Padding = New Padding(2, 2, 4, 2)
            }
            tblCaj.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, CSng(MetaLabelColWidthPx)))
            tblCaj.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, CSng(MetaEditorWidthPx)))
            tblCaj.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0!))
            For i As Integer = 0 To 7
                tblCaj.RowStyles.Add(New RowStyle(SizeType.Absolute, If(i = 1, 27.0F, 29.0F)))
            Next
            tblCaj.RowCount = 8

            Dim toMove As List(Of Control) = tblTrace.Controls.Cast(Of Control)().ToList()
            For Each ctrl As Control In toMove
                Dim pos = tblTrace.GetPositionFromControl(ctrl)
                If pos.Row >= 0 AndAlso pos.Row <= 7 Then
                    tblTrace.Controls.Remove(ctrl)
                    tblCaj.Controls.Add(ctrl, pos.Column, pos.Row)
                    If ReferenceEquals(ctrl, cmbTitleSource) Then tblCaj.SetColumnSpan(ctrl, 2)
                End If
            Next

            tblCaj.RowStyles.Add(New RowStyle(SizeType.Absolute, 29.0!))
            tblCaj.RowCount = 9
            Dim fechaRow As Integer = 8
            tblCaj.Controls.Add(New Label With {.Text = "Fecha plano", .Dock = DockStyle.Fill, .TextAlign = ContentAlignment.MiddleLeft}, 0, fechaRow)
            dtpFechaPlano = New DateTimePicker With {.Format = DateTimePickerFormat.Short, .Value = Date.Today}
            tblCaj.Controls.Add(dtpFechaPlano, 1, fechaRow)
            lblOriginFecha = New Label With {.Text = "vacío", .Dock = DockStyle.Fill, .Font = ItalicHintFont(), .TextAlign = ContentAlignment.MiddleLeft}
            tblCaj.Controls.Add(lblOriginFecha, 2, fechaRow)
            StyleCompactEditor(dtpFechaPlano, MetaEditorWidthPx)
            AddHandler dtpFechaPlano.ValueChanged,
                Sub(s, ev)
                    If _loadingMetadataProgrammatically Then Return
                    MarkCajetinManual("Fecha")
                    MetadataFechaSourceLabel = "manual"
                End Sub

            lblDrawingTitle.Text = "Título"
            lblClient.Text = "Cliente"

            tblTracePartInner = New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .AutoScroll = True,
                .ColumnCount = 3,
                .Padding = New Padding(2, 2, 2, 2)
            }
            tblTracePartInner.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, CSng(MetaLabelColWidthPx)))
            tblTracePartInner.ColumnStyles.Add(New ColumnStyle(SizeType.Absolute, CSng(MetaEditorWidthPx)))
            tblTracePartInner.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0!))
            tblTracePartInner.RowCount = 12
            tblTracePartInner.RowStyles.Clear()
            tblTracePartInner.RowStyles.Add(New RowStyle(SizeType.Absolute, 40.0!))
            For j As Integer = 1 To 10
                tblTracePartInner.RowStyles.Add(New RowStyle(SizeType.Absolute, 28.0!))
            Next
            tblTracePartInner.RowStyles.Add(New RowStyle(SizeType.Absolute, 118.0!))

            Dim r As Integer = 0
            Dim lblPartListAviso As New Label With {
                .Text = "PART LIST: material y espesor se consolidan al generar el DFT. CAJETÍN (arriba) ≠ datos de pieza (aquí).",
                .Dock = DockStyle.Fill,
                .TextAlign = ContentAlignment.TopLeft,
                .AutoSize = False,
                .Font = New Font("Segoe UI", 8.25!, FontStyle.Italic),
                .Padding = New Padding(0, 2, 0, 4)
            }
            tblTracePartInner.SetColumnSpan(lblPartListAviso, 3)
            tblTracePartInner.Controls.Add(lblPartListAviso, 0, r)
            r += 1

            lblMaterial.Text = "Material"
            tblTracePartInner.Controls.Add(lblMaterial, 0, r)
            tblTracePartInner.Controls.Add(txtMaterial, 1, r)
            tblTracePartInner.Controls.Add(lblOriginMaterial, 2, r)
            StyleCompactEditor(txtMaterial, MetaEditorWidthPx)
            r += 1

            lblThickness.Text = "Espesor"
            tblTracePartInner.Controls.Add(lblThickness, 0, r)
            tblTracePartInner.Controls.Add(txtThickness, 1, r)
            tblTracePartInner.Controls.Add(lblOriginThickness, 2, r)
            StyleCompactEditor(txtThickness, MetaEditorWidthPx)
            r += 1

            txtPartNum = New TextBox With {.Text = ""}
            SubAddFieldRowToPartTable(r, "Nº", txtPartNum, lblOriNum) : r += 1

            txtPartCant = New TextBox With {.Text = ""}
            SubAddFieldRowToPartTable(r, "Cantidad", txtPartCant, lblOriCant) : r += 1

            txtPartNombreArchivo = New TextBox()
            SubAddFieldRowToPartTable(r, "Nombre archivo", txtPartNombreArchivo, lblOriNomArch) : r += 1

            txtPartDenominacion = New TextBox()
            SubAddFieldRowToPartTable(r, "Denominación", txtPartDenominacion, lblOriDenom) : r += 1

            txtPartL = New TextBox()
            SubAddFieldRowToPartTable(r, "L (mm)", txtPartL, lblOriL) : r += 1

            txtPartH = New TextBox()
            SubAddFieldRowToPartTable(r, "H (mm)", txtPartH, lblOriH) : r += 1

            txtPartD = New TextBox()
            SubAddFieldRowToPartTable(r, "D (mm)", txtPartD, lblOriD) : r += 1

            txtPartPeso = New TextBox()
            SubAddFieldRowToPartTable(r, "Peso", txtPartPeso, lblOriPeso) : r += 1

            Dim pnlMetaActions As New TableLayoutPanel With {
                .Dock = DockStyle.Fill,
                .ColumnCount = 1,
                .RowCount = 3,
                .Padding = New Padding(0, 4, 0, 0)
            }
            pnlMetaActions.RowStyles.Add(New RowStyle(SizeType.Absolute, 34.0!))
            pnlMetaActions.RowStyles.Add(New RowStyle(SizeType.Absolute, 38.0!))
            pnlMetaActions.RowStyles.Add(New RowStyle(SizeType.Absolute, 34.0!))

            Dim tblBtn3 As New TableLayoutPanel With {.Dock = DockStyle.Fill, .ColumnCount = 3, .RowCount = 1, .Margin = New Padding(0)}
            tblBtn3.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 33.33333!))
            tblBtn3.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 33.33333!))
            tblBtn3.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 33.33333!))
            btnMetaReadModel = New Button With {.Text = "Leer del modelo", .Dock = DockStyle.Fill, .Margin = New Padding(0, 0, 4, 0), .Height = 30}
            btnMetaCalcLhd = New Button With {.Text = "Calcular L/H/D", .Dock = DockStyle.Fill, .Margin = New Padding(2, 0, 2, 0), .Height = 30}
            btnMetaFillEmpty = New Button With {.Text = "Rellenar vacíos", .Dock = DockStyle.Fill, .Margin = New Padding(4, 0, 0, 0), .Height = 30}
            tblBtn3.Controls.Add(btnMetaReadModel, 0, 0)
            tblBtn3.Controls.Add(btnMetaCalcLhd, 1, 0)
            tblBtn3.Controls.Add(btnMetaFillEmpty, 2, 0)

            btnApplyTraceability.Text = "Aplicar propiedades al DFT"
            btnApplyTraceability.Dock = DockStyle.Fill
            btnApplyTraceability.Margin = New Padding(0, 4, 0, 4)
            btnApplyTraceability.Height = 32

            Dim tblBtnSec As New TableLayoutPanel With {.Dock = DockStyle.Fill, .ColumnCount = 2, .RowCount = 1, .Margin = New Padding(0)}
            tblBtnSec.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50.0!))
            tblBtnSec.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50.0!))
            btnMetaApplyDocs = New Button With {.Text = "Aplicar a modelo/DFT", .Dock = DockStyle.Fill, .Margin = New Padding(0, 0, 3, 0)}
            btnMetaPreview = New Button With {.Text = "Previsualizar", .Dock = DockStyle.Fill, .Margin = New Padding(3, 0, 0, 0)}
            tblBtnSec.Controls.Add(btnMetaApplyDocs, 0, 0)
            tblBtnSec.Controls.Add(btnMetaPreview, 1, 0)

            pnlMetaActions.Controls.Add(tblBtn3, 0, 0)
            pnlMetaActions.Controls.Add(btnApplyTraceability, 0, 1)
            pnlMetaActions.Controls.Add(tblBtnSec, 0, 2)

            AddHandler btnMetaReadModel.Click, AddressOf btnMetaReadModel_Click
            AddHandler btnMetaCalcLhd.Click, AddressOf btnMetaCalcLhd_Click
            AddHandler btnMetaFillEmpty.Click, AddressOf btnMetaFillEmpty_Click
            AddHandler btnMetaApplyDocs.Click, AddressOf btnMetaApplyDocs_Click
            AddHandler btnMetaPreview.Click, AddressOf btnMetaPreview_Click

            tblTracePartInner.SetColumnSpan(pnlMetaActions, 3)
            tblTracePartInner.Controls.Add(pnlMetaActions, 0, 11)

            ApplyCompactEditorsInTable(tblCaj)

            grpPlanCajetinBox = New GroupBox With {
                .Text = "Datos del plano (CAJETÍN)",
                .Dock = DockStyle.Fill,
                .Padding = New Padding(5, 10, 5, 14),
                .Margin = New Padding(0, 0, 0, 0)
            }
            grpPlanPartListBox = New GroupBox With {
                .Text = "Datos de pieza (PART LIST)",
                .Dock = DockStyle.Fill,
                .Padding = New Padding(5, 10, 5, 8),
                .Margin = New Padding(0, 10, 0, 0)
            }
            tblCaj.Dock = DockStyle.Fill
            grpPlanCajetinBox.Controls.Add(tblCaj)
            tblTracePartInner.Dock = DockStyle.Fill
            grpPlanPartListBox.Controls.Add(tblTracePartInner)

            While tblTrace.Controls.Count > 0
                tblTrace.Controls.RemoveAt(0)
            End While
            tblTrace.ColumnStyles.Clear()
            tblTrace.RowStyles.Clear()
            tblTrace.ColumnCount = 1
            tblTrace.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100.0!))
            tblTrace.RowCount = 2
            tblTrace.RowStyles.Add(New RowStyle(SizeType.Percent, 50.0!))
            tblTrace.RowStyles.Add(New RowStyle(SizeType.Percent, 50.0!))
            tblTrace.Controls.Add(grpPlanCajetinBox, 0, 0)
            tblTrace.Controls.Add(grpPlanPartListBox, 0, 1)
        Finally
            tblTrace.ResumeLayout(True)
        End Try

        AddHandler txtTitle.TextChanged, AddressOf SyncPartDenominacionFromTitle
    End Sub

    Private Sub ApplyCompactEditorsInTable(t As TableLayoutPanel)
        If t Is Nothing Then Return
        For Each c As Control In t.Controls
            Dim p = t.GetPositionFromControl(c)
            If p.Row < 0 OrElse p.Column <> 1 Then Continue For
            Dim span = t.GetColumnSpan(c)
            Dim w = MetaEditorWidthPx
            If span >= 2 Then w = MetaEditorWidthPx * 2 + 8
            StyleCompactEditor(c, w)
        Next
    End Sub

    Private Shared Sub StyleCompactEditor(c As Control, widthPx As Integer)
        If c Is Nothing Then Return
        If TypeOf c Is TextBox Then
            Dim tb = DirectCast(c, TextBox)
            tb.Dock = DockStyle.None
            tb.Anchor = AnchorStyles.Left Or AnchorStyles.Top
            tb.Width = widthPx
            tb.Margin = New Padding(0, 2, 0, 2)
        ElseIf TypeOf c Is ComboBox Then
            Dim cb = DirectCast(c, ComboBox)
            cb.Dock = DockStyle.None
            cb.Anchor = AnchorStyles.Left Or AnchorStyles.Top
            cb.Width = widthPx
            cb.Margin = New Padding(0, 2, 0, 2)
        ElseIf TypeOf c Is DateTimePicker Then
            Dim d = DirectCast(c, DateTimePicker)
            d.Dock = DockStyle.None
            d.Anchor = AnchorStyles.Left Or AnchorStyles.Top
            d.Width = widthPx
            d.Margin = New Padding(0, 2, 0, 2)
        End If
    End Sub

    Private Shared Function ItalicHintFont() As Font
        Return New Font("Segoe UI", 8.0!, FontStyle.Italic)
    End Function

    Private Shared Sub DetachControlFromTable(t As TableLayoutPanel, c As Control)
        If c Is Nothing OrElse t Is Nothing Then Return
        If c.Parent Is t Then t.Controls.Remove(c)
    End Sub

    Private Sub SubAddFieldRowToPartTable(row As Integer, caption As String, editor As Control, ByRef ori As Label)
        If tblTracePartInner Is Nothing Then Return
        tblTracePartInner.Controls.Add(New Label With {.Text = caption, .Dock = DockStyle.Fill, .TextAlign = ContentAlignment.MiddleLeft}, 0, row)
        editor.Dock = DockStyle.None
        editor.Anchor = AnchorStyles.Left Or AnchorStyles.Top
        editor.Width = MetaEditorWidthPx
        editor.Margin = New Padding(0, 2, 0, 2)
        If TypeOf editor Is TextBox Then
            Dim tb = DirectCast(editor, TextBox)
            If tb.Height < 12 Then tb.Height = 23
        End If
        tblTracePartInner.Controls.Add(editor, 1, row)
        If ori Is Nothing Then ori = New Label()
        ori.Dock = DockStyle.Fill
        ori.Font = ItalicHintFont()
        ori.TextAlign = ContentAlignment.MiddleLeft
        ori.Text = "vacío"
        tblTracePartInner.Controls.Add(ori, 2, row)
    End Sub

    Private Sub SyncPartDenominacionFromTitle(sender As Object, e As EventArgs)
        If _suspendDenomSync Then Return
        If txtPartDenominacion Is Nothing OrElse txtTitle Is Nothing Then Return
        If txtPartDenominacion.Focused Then Return
        txtPartDenominacion.Text = txtTitle.Text
    End Sub

    Public Property MetadataCliente As String
        Get
            Return If(txtClient?.Text, "").Trim()
        End Get
        Set(value As String)
            If txtClient IsNot Nothing Then txtClient.Text = value
        End Set
    End Property

    Public Property MetadataProyecto As String
        Get
            Return If(txtProject?.Text, "").Trim()
        End Get
        Set(value As String)
            If txtProject IsNot Nothing Then txtProject.Text = value
        End Set
    End Property

    Public Property MetadataPedido As String
        Get
            Return If(txtPedido?.Text, "").Trim()
        End Get
        Set(value As String)
            If txtPedido IsNot Nothing Then txtPedido.Text = value
        End Set
    End Property

    Public Property MetadataPedidoSourceLabel As String
        Get
            Return If(lblOriginPedido?.Text, "").Trim()
        End Get
        Set(value As String)
            If lblOriginPedido IsNot Nothing Then lblOriginPedido.Text = If(String.IsNullOrWhiteSpace(value), "vacío", value)
        End Set
    End Property

    Public Property MetadataClienteSourceLabel As String
        Get
            Return If(lblOriginClient?.Text, "").Trim()
        End Get
        Set(value As String)
            If lblOriginClient IsNot Nothing Then lblOriginClient.Text = If(String.IsNullOrWhiteSpace(value), "vacío", value)
        End Set
    End Property

    Public Property MetadataProyectoSourceLabel As String
        Get
            Return If(lblOriginProject?.Text, "").Trim()
        End Get
        Set(value As String)
            If lblOriginProject IsNot Nothing Then lblOriginProject.Text = If(String.IsNullOrWhiteSpace(value), "vacío", value)
        End Set
    End Property

    Public Property MetadataPlanoSourceLabel As String
        Get
            Return If(lblOriginDocNum?.Text, "").Trim()
        End Get
        Set(value As String)
            If lblOriginDocNum IsNot Nothing Then lblOriginDocNum.Text = If(String.IsNullOrWhiteSpace(value), "vacío", value)
        End Set
    End Property

    Public Property MetadataTituloSourceLabel As String
        Get
            Return If(lblOriginTitle?.Text, "").Trim()
        End Get
        Set(value As String)
            If lblOriginTitle IsNot Nothing Then lblOriginTitle.Text = If(String.IsNullOrWhiteSpace(value), "vacío", value)
        End Set
    End Property

    Public Property MetadataRevisionSourceLabel As String
        Get
            Return If(lblOriginRevision?.Text, "").Trim()
        End Get
        Set(value As String)
            If lblOriginRevision IsNot Nothing Then lblOriginRevision.Text = If(String.IsNullOrWhiteSpace(value), "vacío", value)
        End Set
    End Property

    Public Property MetadataAutorSourceLabel As String
        Get
            Return If(lblOriginAuthor?.Text, "").Trim()
        End Get
        Set(value As String)
            If lblOriginAuthor IsNot Nothing Then lblOriginAuthor.Text = If(String.IsNullOrWhiteSpace(value), "vacío", value)
        End Set
    End Property

    Public Property MetadataFechaSourceLabel As String
        Get
            Return If(lblOriginFecha?.Text, "").Trim()
        End Get
        Set(value As String)
            If lblOriginFecha IsNot Nothing Then lblOriginFecha.Text = If(String.IsNullOrWhiteSpace(value), "vacío", value)
        End Set
    End Property

    Public Property MetadataNumeroSourceLabel As String
        Get
            Return If(lblOriNum?.Text, "").Trim()
        End Get
        Set(value As String)
            If lblOriNum IsNot Nothing Then lblOriNum.Text = If(String.IsNullOrWhiteSpace(value), "vacío", value)
        End Set
    End Property

    Public Property MetadataCantidadSourceLabel As String
        Get
            Return If(lblOriCant?.Text, "").Trim()
        End Get
        Set(value As String)
            If lblOriCant IsNot Nothing Then lblOriCant.Text = If(String.IsNullOrWhiteSpace(value), "vacío", value)
        End Set
    End Property

    Public Property MetadataNombreArchivoSourceLabel As String
        Get
            Return If(lblOriNomArch?.Text, "").Trim()
        End Get
        Set(value As String)
            If lblOriNomArch IsNot Nothing Then lblOriNomArch.Text = If(String.IsNullOrWhiteSpace(value), "vacío", value)
        End Set
    End Property

    Public Property MetadataDenominacionSourceLabel As String
        Get
            Return If(lblOriDenom?.Text, "").Trim()
        End Get
        Set(value As String)
            If lblOriDenom IsNot Nothing Then lblOriDenom.Text = If(String.IsNullOrWhiteSpace(value), "vacío", value)
        End Set
    End Property

    Public Property MetadataPlano As String
        Get
            Return If(txtDrawingNumber?.Text, "").Trim()
        End Get
        Set(value As String)
            If txtDrawingNumber IsNot Nothing Then txtDrawingNumber.Text = value
        End Set
    End Property

    Public Property MetadataTitulo As String
        Get
            Return If(txtTitle?.Text, "").Trim()
        End Get
        Set(value As String)
            If txtTitle IsNot Nothing Then txtTitle.Text = value
        End Set
    End Property

    Public Property MetadataRevision As String
        Get
            Return If(txtRevision?.Text, "").Trim()
        End Get
        Set(value As String)
            If txtRevision IsNot Nothing Then txtRevision.Text = value
        End Set
    End Property

    Public Property MetadataAutor As String
        Get
            Return If(txtAuthor?.Text, "").Trim()
        End Get
        Set(value As String)
            If txtAuthor IsNot Nothing Then txtAuthor.Text = value
        End Set
    End Property

    Public Property MetadataFecha As Date
        Get
            If dtpFechaPlano Is Nothing Then Return Date.Today
            Return dtpFechaPlano.Value.Date
        End Get
        Set(value As Date)
            If dtpFechaPlano IsNot Nothing Then dtpFechaPlano.Value = value
        End Set
    End Property

    Public Property MetadataNumeroPartList As String
        Get
            Return If(txtPartNum?.Text, "").Trim()
        End Get
        Set(value As String)
            If txtPartNum IsNot Nothing Then txtPartNum.Text = If(value, "")
        End Set
    End Property

    Public Property MetadataCantidad As String
        Get
            Return If(txtPartCant?.Text, "1").Trim()
        End Get
        Set(value As String)
            If txtPartCant IsNot Nothing Then txtPartCant.Text = value
        End Set
    End Property

    Public Property MetadataNombreArchivo As String
        Get
            Return If(txtPartNombreArchivo?.Text, "").Trim()
        End Get
        Set(value As String)
            If txtPartNombreArchivo IsNot Nothing Then txtPartNombreArchivo.Text = value
        End Set
    End Property

    Public Property MetadataDenominacion As String
        Get
            Return If(txtPartDenominacion?.Text, "").Trim()
        End Get
        Set(value As String)
            If txtPartDenominacion IsNot Nothing Then
                _suspendDenomSync = True
                txtPartDenominacion.Text = value
                _suspendDenomSync = False
            End If
        End Set
    End Property

    Public Property MetadataMaterial As String
        Get
            Return If(txtMaterial?.Text, "").Trim()
        End Get
        Set(value As String)
            If txtMaterial IsNot Nothing Then txtMaterial.Text = value
        End Set
    End Property

    Public Property MetadataEspesor As String
        Get
            Return If(txtThickness?.Text, "").Trim()
        End Get
        Set(value As String)
            If txtThickness IsNot Nothing Then txtThickness.Text = value
        End Set
    End Property

    Public Property MetadataL As String
        Get
            Return If(txtPartL?.Text, "").Trim()
        End Get
        Set(value As String)
            If txtPartL IsNot Nothing Then txtPartL.Text = value
        End Set
    End Property

    Public Property MetadataH As String
        Get
            Return If(txtPartH?.Text, "").Trim()
        End Get
        Set(value As String)
            If txtPartH IsNot Nothing Then txtPartH.Text = value
        End Set
    End Property

    Public Property MetadataD As String
        Get
            Return If(txtPartD?.Text, "").Trim()
        End Get
        Set(value As String)
            If txtPartD IsNot Nothing Then txtPartD.Text = value
        End Set
    End Property

    Public Property MetadataPeso As String
        Get
            Return If(txtPartPeso?.Text, "").Trim()
        End Get
        Set(value As String)
            If txtPartPeso IsNot Nothing Then txtPartPeso.Text = value
        End Set
    End Property

    Public Property MetadataMaterialSourceLabel As String
        Get
            Return If(lblOriginMaterial?.Text, "").Trim()
        End Get
        Set(value As String)
            If lblOriginMaterial IsNot Nothing Then lblOriginMaterial.Text = If(String.IsNullOrWhiteSpace(value), "vacío", value)
        End Set
    End Property

    Public Property MetadataEspesorSourceLabel As String
        Get
            Return If(lblOriginThickness?.Text, "").Trim()
        End Get
        Set(value As String)
            If lblOriginThickness IsNot Nothing Then lblOriginThickness.Text = If(String.IsNullOrWhiteSpace(value), "vacío", value)
        End Set
    End Property

    Public Property MetadataLhdSourceLabel As String
        Get
            Return If(lblOriL?.Text, "").Trim()
        End Get
        Set(value As String)
            Dim t = If(String.IsNullOrWhiteSpace(value), "vacío", value)
            If lblOriL IsNot Nothing Then lblOriL.Text = t
            If lblOriH IsNot Nothing Then lblOriH.Text = t
            If lblOriD IsNot Nothing Then lblOriD.Text = t
        End Set
    End Property

    Public Property MetadataPesoSourceLabel As String
        Get
            Return If(lblOriPeso?.Text, "").Trim()
        End Get
        Set(value As String)
            If lblOriPeso IsNot Nothing Then lblOriPeso.Text = If(String.IsNullOrWhiteSpace(value), "vacío", value)
        End Set
    End Property

    Private Sub ApplyMetadataDataToForm(data As DrawingMetadataInput)
        If data Is Nothing Then Return
        DrawingMetadataService.ApplyToUi(Me, data)
    End Sub

    Friend Sub ClearPlanMetadataUi(Optional logLine As String = Nothing)
        EnsureDrawingPlanMetadataPanel()
        _loadingMetadataProgrammatically = True
        Try
            ClearManualCajetinFieldTracking()
            txtClient.Clear()
            txtProject.Clear()
            txtPedido.Clear()
            txtDrawingNumber.Clear()
            txtTitle.Clear()
            txtRevision.Clear()
            txtAuthor.Clear()
            txtNotes.Clear()
            If dtpFechaPlano IsNot Nothing Then dtpFechaPlano.Value = Date.Today
            If txtPartNum IsNot Nothing Then txtPartNum.Clear()
            If txtPartCant IsNot Nothing Then txtPartCant.Clear()
            If txtPartNombreArchivo IsNot Nothing Then txtPartNombreArchivo.Clear()
            If txtPartDenominacion IsNot Nothing Then txtPartDenominacion.Clear()
            If txtPartL IsNot Nothing Then txtPartL.Clear()
            If txtPartH IsNot Nothing Then txtPartH.Clear()
            If txtPartD IsNot Nothing Then txtPartD.Clear()
            If txtPartPeso IsNot Nothing Then txtPartPeso.Clear()
            If txtMaterial IsNot Nothing Then txtMaterial.Clear()
            If txtThickness IsNot Nothing Then txtThickness.Clear()

            MetadataClienteSourceLabel = "vacío"
            MetadataProyectoSourceLabel = "vacío"
            MetadataPedidoSourceLabel = "vacío"
            MetadataPlanoSourceLabel = "vacío"
            MetadataTituloSourceLabel = "vacío"
            MetadataRevisionSourceLabel = "vacío"
            MetadataAutorSourceLabel = "vacío"
            MetadataFechaSourceLabel = "vacío"
            MetadataNumeroSourceLabel = "vacío"
            MetadataCantidadSourceLabel = "vacío"
            MetadataNombreArchivoSourceLabel = "vacío"
            MetadataDenominacionSourceLabel = "vacío"
            MetadataMaterialSourceLabel = "vacío"
            MetadataEspesorSourceLabel = "vacío"
            MetadataLhdSourceLabel = "vacío"
            MetadataPesoSourceLabel = "vacío"
        Finally
            _loadingMetadataProgrammatically = False
        End Try
        If Not String.IsNullOrWhiteSpace(logLine) AndAlso _logger IsNot Nothing Then _logger.Log(logLine)
    End Sub

    Friend Sub ClearPartListMetadataUiOnly()
        EnsureDrawingPlanMetadataPanel()
        _loadingMetadataProgrammatically = True
        Try
            If txtPartNum IsNot Nothing Then txtPartNum.Clear()
            If txtPartCant IsNot Nothing Then txtPartCant.Clear()
            If txtPartNombreArchivo IsNot Nothing Then txtPartNombreArchivo.Clear()
            If txtPartDenominacion IsNot Nothing Then txtPartDenominacion.Clear()
            If txtPartL IsNot Nothing Then txtPartL.Clear()
            If txtPartH IsNot Nothing Then txtPartH.Clear()
            If txtPartD IsNot Nothing Then txtPartD.Clear()
            If txtPartPeso IsNot Nothing Then txtPartPeso.Clear()
            If txtMaterial IsNot Nothing Then txtMaterial.Clear()
            If txtThickness IsNot Nothing Then txtThickness.Clear()
            MetadataNumeroSourceLabel = "vacío"
            MetadataCantidadSourceLabel = "vacío"
            MetadataNombreArchivoSourceLabel = "vacío"
            MetadataDenominacionSourceLabel = "vacío"
            MetadataMaterialSourceLabel = "vacío"
            MetadataEspesorSourceLabel = "vacío"
            MetadataLhdSourceLabel = "vacío"
            MetadataPesoSourceLabel = "vacío"
        Finally
            _loadingMetadataProgrammatically = False
        End Try
        If _logger IsNot Nothing Then _logger.Log("[UI][METADATA][CLEAR_BEFORE_LOAD] scope=PART_LIST")
    End Sub

    Private Sub btnMetaReadModel_Click(sender As Object, e As EventArgs)
        If String.IsNullOrWhiteSpace(txtInputFile.Text) OrElse Not File.Exists(txtInputFile.Text) Then
            MessageBox.Show("Selecciona un archivo PAR/PSM/ASM válido.", "Metadatos", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        _logger.Log("[UI][METADATA][READ_MODEL_CURRENT_CONTEXT]")
        Dim cfg As New JobConfiguration With {.InputFile = txtInputFile.Text}
        Dim kind As SourceFileKind = cfg.DetectInputKind()
        Dim path As String = txtInputFile.Text.Trim()

        If kind = SourceFileKind.AssemblyFile Then
            Dim dataAsm As DrawingMetadataInput = Nothing
            If Not DrawingMetadataService.TryLoadMetadataFromModelFile(path, chkKeepSolidEdgeVisible.Checked, _logger, dataAsm) OrElse dataAsm Is Nothing Then
                MessageBox.Show("No se pudieron leer metadatos del ensamblaje.", "Metadatos", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
            DrawingMetadataService.ApplyToUi(Me, dataAsm, applyCajetin:=True, applyPartList:=False)
            _logger.Log("[UI][METADATA] Cajetín ASM releído (PART_LIST no modificado).")
            Return
        End If

        If kind = SourceFileKind.PartFile OrElse kind = SourceFileKind.SheetMetalFile Then
            Dim data As DrawingMetadataInput = Nothing
            If Not String.IsNullOrWhiteSpace(_loadedAsmComponentPath) AndAlso File.Exists(_loadedAsmComponentPath) Then
                If DrawingMetadataService.TryLoadMetadataFromModelFile(_loadedAsmComponentPath, chkKeepSolidEdgeVisible.Checked, _logger, data) Then
                    DrawingMetadataService.ApplyToUi(Me, data, applyCajetin:=False, applyPartList:=True)
                    _logger.Log("[UI][METADATA] PART_LIST releído desde pieza en contexto.")
                Else
                    MessageBox.Show("No se pudieron leer metadatos del modelo.", "Metadatos", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                End If
            Else
                If DrawingMetadataService.TryLoadMetadataFromModelFile(path, chkKeepSolidEdgeVisible.Checked, _logger, data) Then
                    DrawingMetadataService.ApplyToUi(Me, data, applyCajetin:=True, applyPartList:=True)
                    _logger.Log("[UI][METADATA] Metadatos PAR/PSM aplicados (cajetín + PART_LIST).")
                Else
                    MessageBox.Show("No se pudieron leer metadatos del modelo.", "Metadatos", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                End If
            End If
            Return
        End If

        MessageBox.Show("Tipo de archivo no soportado para leer metadatos.", "Metadatos", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub btnMetaCalcLhd_Click(sender As Object, e As EventArgs)
        Dim dftPath As String = ResolveManualDftPath(BuildConfigurationFromUi())
        Dim modelPath As String = If(txtInputFile?.Text, "").Trim()
        Dim haveDft As Boolean = Not String.IsNullOrWhiteSpace(dftPath) AndAlso File.Exists(dftPath)
        If Not haveDft AndAlso (String.IsNullOrWhiteSpace(modelPath) OrElse Not File.Exists(modelPath)) Then
            MessageBox.Show(
                "No hay DFT guardado ni archivo de entrada PAR/PSM." & Environment.NewLine &
                "Opciones: genere o guarde un DFT, o seleccione un modelo .par/.psm como archivo de entrada.",
                "L/H/D", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Dim app As SolidEdgeFramework.Application = Nothing
        Dim created As Boolean = False
        Dim doc As Object = Nothing
        Try
            OleMessageFilter.Register()
            If Not ConnectSeForLhd(chkKeepSolidEdgeVisible.Checked, app, created) Then Return

            Dim L As String = "", H As String = "", D As String = ""

            If haveDft Then
                doc = app.Documents.Open(dftPath)
                If DrawingMetadataService.TryComputeLhdFromDraft(doc, _logger, L, H, D) Then
                    MetadataL = L
                    MetadataH = H
                    MetadataD = D
                    MetadataLhdSourceLabel = "Calculado (DFT)"
                    _logger.Log("[PARTLISTDATA][LHD][UI] Valores aplicados desde DFT.")
                    Return
                End If
                Try
                    If doc IsNot Nothing Then CallByName(doc, "Close", CallType.Method, False)
                Catch
                Finally
                    doc = Nothing
                End Try
                _logger.Log("[PARTLISTDATA][LHD][UI] DFT sin cotas/bbox útiles; se intenta modelo 3D.")
            End If

            If String.IsNullOrWhiteSpace(modelPath) OrElse Not File.Exists(modelPath) Then
                MessageBox.Show("No hay archivo PAR/PSM de entrada para calcular la caja 3D.", "L/H/D", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Dim cfg As New JobConfiguration With {.InputFile = modelPath}
            If cfg.DetectInputKind() = SourceFileKind.AssemblyFile Then
                MessageBox.Show("Para ensamblajes (.asm) calcule L/H/D desde un DFT o elija un componente PAR/PSM.", "L/H/D", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            doc = app.Documents.Open(modelPath)
            If DrawingMetadataService.TryComputeLhdFromModelDoc(doc, _logger, L, H, D) Then
                MetadataL = L
                MetadataH = H
                MetadataD = D
                MetadataLhdSourceLabel = "Caja 3D (modelo)"
                _logger.Log("[PARTLISTDATA][LHD][UI] Valores aplicados desde rango del modelo.")
            Else
                MessageBox.Show(
                    "No se pudieron calcular L/H/D." & Environment.NewLine &
                    If(haveDft, "El DFT no aportó cotas ni contornos de vista; el modelo no devolvió rango 3D legible.", "No se pudo leer la caja 3D del modelo."),
                    "L/H/D", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
        Catch ex As Exception
            _logger.LogException("btnMetaCalcLhd_Click", ex)
            MessageBox.Show(ex.Message, "L/H/D", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Try
                If doc IsNot Nothing Then CallByName(doc, "Close", CallType.Method, False)
            Catch
            End Try
            Try
                If app IsNot Nothing AndAlso created Then app.Quit()
            Catch
            End Try
            Try : OleMessageFilter.Revoke() : Catch : End Try
        End Try
    End Sub

    Private Function ConnectSeForLhd(visible As Boolean, ByRef app As SolidEdgeFramework.Application, ByRef created As Boolean) As Boolean
        app = Nothing
        created = False
        Try
            app = CType(Runtime.InteropServices.Marshal.GetActiveObject("SolidEdge.Application"), SolidEdgeFramework.Application)
        Catch
            Dim t = Type.GetTypeFromProgID("SolidEdge.Application")
            app = CType(Activator.CreateInstance(t), SolidEdgeFramework.Application)
            created = True
        End Try
        If app Is Nothing Then Return False
        app.Visible = visible
        app.DisplayAlerts = False
        Return True
    End Function

    Private Sub btnMetaFillEmpty_Click(sender As Object, e As EventArgs)
        If String.IsNullOrWhiteSpace(txtInputFile.Text) OrElse Not File.Exists(txtInputFile.Text) Then Return
        _logger.Log("[UI][METADATA][FILL_EMPTY_ONLY]")
        Dim cfg As New JobConfiguration With {.InputFile = txtInputFile.Text}
        Dim kind As SourceFileKind = cfg.DetectInputKind()
        Dim snap As DrawingMetadataInput = DrawingMetadataService.BuildFromUi(Me)

        If kind = SourceFileKind.AssemblyFile Then
            Dim mdl As DrawingMetadataInput = Nothing
            If Not DrawingMetadataService.TryLoadMetadataFromModelFile(txtInputFile.Text, chkKeepSolidEdgeVisible.Checked, _logger, mdl) OrElse mdl Is Nothing Then Return
            MergeAsmCajetinFieldIfEmpty(snap, mdl, "Cliente")
            MergeAsmCajetinFieldIfEmpty(snap, mdl, "Proyecto")
            MergeAsmCajetinFieldIfEmpty(snap, mdl, "Pedido")
            MergeAsmCajetinFieldIfEmpty(snap, mdl, "Plano")
            MergeAsmCajetinFieldIfEmpty(snap, mdl, "Titulo")
            MergeAsmCajetinFieldIfEmpty(snap, mdl, "Revision")
            MergeAsmCajetinFieldIfEmpty(snap, mdl, "Autor")
            _logger.Log("[UI][METADATA] Rellenar vacíos (cajetín ASM) completado.")
            Return
        End If

        Dim modelPath As String = txtInputFile.Text.Trim()
        If Not String.IsNullOrWhiteSpace(_loadedAsmComponentPath) AndAlso File.Exists(_loadedAsmComponentPath) Then
            modelPath = _loadedAsmComponentPath
        End If
        Dim mdlPart As DrawingMetadataInput = Nothing
        DrawingMetadataService.TryLoadMetadataFromModelFile(modelPath, chkKeepSolidEdgeVisible.Checked, _logger, mdlPart)
        If mdlPart Is Nothing Then mdlPart = New DrawingMetadataInput()

        If String.IsNullOrWhiteSpace(snap.Material) AndAlso Not String.IsNullOrWhiteSpace(mdlPart.Material) Then
            MetadataMaterial = mdlPart.Material
            MetadataMaterialSourceLabel = mdlPart.MaterialSource
        ElseIf Not String.IsNullOrWhiteSpace(snap.Material) Then
            _logger.Log("[UI][METADATA][SKIP_NOT_EMPTY] field=Material")
        End If
        If String.IsNullOrWhiteSpace(snap.Espesor) AndAlso Not String.IsNullOrWhiteSpace(mdlPart.Espesor) Then
            MetadataEspesor = mdlPart.Espesor
            MetadataEspesorSourceLabel = mdlPart.EspesorSource
        ElseIf Not String.IsNullOrWhiteSpace(snap.Espesor) Then
            _logger.Log("[UI][METADATA][SKIP_NOT_EMPTY] field=Espesor")
        End If
        If String.IsNullOrWhiteSpace(snap.Peso) AndAlso Not String.IsNullOrWhiteSpace(mdlPart.Peso) Then
            MetadataPeso = mdlPart.Peso
            MetadataPesoSourceLabel = mdlPart.PesoSource
        ElseIf Not String.IsNullOrWhiteSpace(snap.Peso) Then
            _logger.Log("[UI][METADATA][SKIP_NOT_EMPTY] field=Peso")
        End If
        If String.IsNullOrWhiteSpace(snap.Pedido) Then
            Dim p = DrawingMetadataService.InferPedidoFromPath(modelPath, _logger)
            If p <> "" Then
                MetadataPedido = p
                MetadataPedidoSourceLabel = "Inferido (carpeta o nombre de archivo)"
                _logger.Log("[METADATA][PEDIDO][INFERRED]")
            End If
        Else
            _logger.Log("[UI][METADATA][SKIP_NOT_EMPTY] field=Pedido")
        End If
        If String.IsNullOrWhiteSpace(snap.Plano) Then
            Dim p = DrawingMetadataService.InferPlanoFromFileName(modelPath)
            If Not String.IsNullOrWhiteSpace(p) Then
                MetadataPlano = p
                MetadataPlanoSourceLabel = "calculado (nombre fichero)"
                _logger.Log("[METADATA][PLANO][FROM_FILENAME] " & MetadataPlano)
            End If
        Else
            _logger.Log("[UI][METADATA][SKIP_NOT_EMPTY] field=Plano")
        End If
        If String.IsNullOrWhiteSpace(snap.Denominacion) AndAlso Not String.IsNullOrWhiteSpace(mdlPart.Denominacion) Then
            MetadataDenominacion = mdlPart.Denominacion
            MetadataDenominacionSourceLabel = mdlPart.DenominacionSource
        ElseIf String.IsNullOrWhiteSpace(snap.Denominacion) Then
            Dim d = DrawingMetadataService.InferDenominacionListaFromFileName(modelPath)
            If Not String.IsNullOrWhiteSpace(d) Then
                MetadataDenominacion = d
                MetadataDenominacionSourceLabel = "calculado (nombre fichero)"
                _logger.Log("[METADATA][DENOM_LIST][FROM_FILENAME] " & d)
            End If
        Else
            _logger.Log("[UI][METADATA][SKIP_NOT_EMPTY] field=Denominacion")
        End If
        If String.IsNullOrWhiteSpace(snap.NombreArchivo) Then
            Try
                MetadataNombreArchivo = Path.GetFileName(modelPath)
                MetadataNombreArchivoSourceLabel = "calculado (nombre fichero)"
            Catch
            End Try
        Else
            _logger.Log("[UI][METADATA][SKIP_NOT_EMPTY] field=NombreArchivo")
        End If
        _logger.Log("[UI][METADATA] Rellenar vacíos completado.")
    End Sub

    Private Sub MergeAsmCajetinFieldIfEmpty(snap As DrawingMetadataInput, mdl As DrawingMetadataInput, field As String)
        If snap Is Nothing OrElse mdl Is Nothing Then Return
        Dim cur As String = ""
        Dim neu As String = ""
        Dim ori As String = ""
        Select Case field
            Case "Cliente" : cur = snap.Cliente : neu = mdl.Cliente : ori = mdl.ClienteSource
            Case "Proyecto" : cur = snap.Proyecto : neu = mdl.Proyecto : ori = mdl.ProyectoSource
            Case "Pedido" : cur = snap.Pedido : neu = mdl.Pedido : ori = mdl.PedidoSource
            Case "Plano" : cur = snap.Plano : neu = mdl.Plano : ori = mdl.PlanoSource
            Case "Titulo" : cur = snap.Titulo : neu = mdl.Titulo : ori = mdl.TituloSource
            Case "Revision" : cur = snap.Revision : neu = mdl.Revision : ori = mdl.RevisionSource
            Case "Autor" : cur = snap.Autor : neu = mdl.Autor : ori = mdl.AutorSource
            Case Else : Return
        End Select
        If Not String.IsNullOrWhiteSpace(cur) Then
            _logger.Log("[UI][METADATA][SKIP_NOT_EMPTY] field=" & field)
            Return
        End If
        If String.IsNullOrWhiteSpace(neu) Then Return
        Dim o = If(String.IsNullOrWhiteSpace(ori), "modelo", ori)
        Select Case field
            Case "Cliente" : MetadataCliente = neu : MetadataClienteSourceLabel = o
            Case "Proyecto" : MetadataProyecto = neu : MetadataProyectoSourceLabel = o
            Case "Pedido" : MetadataPedido = neu : MetadataPedidoSourceLabel = o
            Case "Plano" : MetadataPlano = neu : MetadataPlanoSourceLabel = o
            Case "Titulo" : MetadataTitulo = neu : MetadataTituloSourceLabel = o
            Case "Revision" : MetadataRevision = neu : MetadataRevisionSourceLabel = o
            Case "Autor" : MetadataAutor = neu : MetadataAutorSourceLabel = o
        End Select
    End Sub

    Private Sub btnMetaApplyDocs_Click(sender As Object, e As EventArgs)
        Try
            Dim cfg As JobConfiguration = BuildConfigurationFromUi()
            Dim dftPath As String = ResolveManualDftPath(cfg)
            If String.IsNullOrWhiteSpace(dftPath) OrElse Not File.Exists(dftPath) Then
                MessageBox.Show("No se encontró DFT para aplicar propiedades.", "Aplicar", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
            Dim data = DrawingMetadataService.BuildFromUi(Me)
            Dim app As SolidEdgeFramework.Application = Nothing
            Dim created As Boolean = False
            Dim dftDoc As Object = Nothing
            Dim modelDoc As Object = Nothing
            Try
                OleMessageFilter.Register()
                If Not ConnectSeForLhd(chkKeepSolidEdgeVisible.Checked, app, created) Then Return
                dftDoc = app.Documents.Open(dftPath)
                Dim modelPath As String = cfg.InputFile
                If File.Exists(modelPath) AndAlso cfg.DetectInputKind() <> SourceFileKind.AssemblyFile Then
                    modelDoc = app.Documents.Open(modelPath)
                End If
                DrawingMetadataService.ApplyPartListSourceProperties(modelDoc, dftDoc, data, _logger)
                Try
                    CallByName(dftDoc, "Save", CallType.Method)
                Catch
                End Try
                MessageBox.Show("Propiedades aplicadas. Revise el log y la PartsList.", "Aplicar", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Finally
                Try
                    If modelDoc IsNot Nothing Then CallByName(modelDoc, "Close", CallType.Method, True)
                Catch
                End Try
                Try
                    If dftDoc IsNot Nothing Then CallByName(dftDoc, "Close", CallType.Method, False)
                Catch
                End Try
                Try
                    If app IsNot Nothing AndAlso created Then app.Quit()
                Catch
                End Try
                Try : OleMessageFilter.Revoke() : Catch : End Try
            End Try
        Catch ex As Exception
            _logger.LogException("btnMetaApplyDocs_Click", ex)
            MessageBox.Show(ex.Message, "Aplicar", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnMetaPreview_Click(sender As Object, e As EventArgs)
        Dim data = DrawingMetadataService.BuildFromUi(Me)
        Dim sb As New System.Text.StringBuilder()
        sb.AppendLine("Cajetín:")
        sb.AppendLine($"  Cliente={data.Cliente}")
        sb.AppendLine($"  Proyecto={data.Proyecto}")
        sb.AppendLine($"  Pedido={data.Pedido}")
        sb.AppendLine($"  Plano={data.Plano}")
        sb.AppendLine($"  Título={data.Titulo}")
        sb.AppendLine($"  Rev={data.Revision}")
        sb.AppendLine($"  Autor={data.Autor}")
        sb.AppendLine($"  Fecha={data.Fecha:d}")
        sb.AppendLine("PartsList:")
        sb.AppendLine($"  Cant={data.Cantidad}  Archivo={data.NombreArchivo}")
        sb.AppendLine($"  Denom={data.Denominacion}")
        sb.AppendLine($"  Mat={data.Material} ({data.MaterialSource})")
        sb.AppendLine($"  Esp={data.Espesor} ({data.EspesorSource})")
        sb.AppendLine($"  L/H/D={data.LargoL} / {data.AltoH} / {data.DatoD} ({data.LHDSource})")
        sb.AppendLine($"  Peso={data.Peso} ({data.PesoSource})")
        MessageBox.Show(sb.ToString(), "Previsualización metadatos", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

End Class
