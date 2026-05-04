Option Strict Off

Imports System.Collections.Generic
Imports System.Globalization
Imports System.IO
Imports System.Linq
Imports System.Text
Imports System.Text.RegularExpressions
Imports SolidEdgeDraft
Imports SolidEdgeFramework
Imports SolidEdgeFrameworkSupport
Imports SolidEdgeGeometry
Imports SolidEdgePart

''' <summary>Carga, detección L/H/D, validación y escritura de metadatos para cajetín y PART_LIST.</summary>
Public NotInheritable Class DrawingMetadataService
    Private Const TolMm As Double = 0.5R

    Private Sub New()
    End Sub

    Public Shared Function BuildFromUi(main As MainForm) As DrawingMetadataInput
        If main Is Nothing Then Return New DrawingMetadataInput()
        Dim m As New DrawingMetadataInput With {
            .Cliente = main.MetadataCliente,
            .Proyecto = main.MetadataProyecto,
            .Pedido = main.MetadataPedido,
            .Plano = main.MetadataPlano,
            .Titulo = main.MetadataTitulo,
            .Revision = main.MetadataRevision,
            .Autor = main.MetadataAutor,
            .Fecha = main.MetadataFecha,
            .FechaSource = main.MetadataFechaSourceLabel,
            .Numero = main.MetadataNumeroPartList,
            .Cantidad = main.MetadataCantidad,
            .NombreArchivo = main.MetadataNombreArchivo,
            .Denominacion = main.MetadataDenominacion,
            .Material = main.MetadataMaterial,
            .Espesor = main.MetadataEspesor,
            .LargoL = main.MetadataL,
            .AltoH = main.MetadataH,
            .DatoD = main.MetadataD,
            .Peso = main.MetadataPeso,
            .ClienteSource = main.MetadataClienteSourceLabel,
            .ProyectoSource = main.MetadataProyectoSourceLabel,
            .PlanoSource = main.MetadataPlanoSourceLabel,
            .TituloSource = main.MetadataTituloSourceLabel,
            .RevisionSource = main.MetadataRevisionSourceLabel,
            .AutorSource = main.MetadataAutorSourceLabel,
            .PedidoSource = main.MetadataPedidoSourceLabel,
            .NumeroSource = main.MetadataNumeroSourceLabel,
            .CantidadSource = main.MetadataCantidadSourceLabel,
            .NombreArchivoSource = main.MetadataNombreArchivoSourceLabel,
            .DenominacionSource = main.MetadataDenominacionSourceLabel,
            .MaterialSource = main.MetadataMaterialSourceLabel,
            .EspesorSource = main.MetadataEspesorSourceLabel,
            .LHDSource = main.MetadataLhdSourceLabel,
            .PesoSource = main.MetadataPesoSourceLabel
        }
        Return m
    End Function

    Public Shared Sub ApplyToUi(main As MainForm, data As DrawingMetadataInput,
                                Optional applyCajetin As Boolean = True,
                                Optional applyPartList As Boolean = True)
        If main Is Nothing OrElse data Is Nothing Then Return
        If applyCajetin Then
            main.MetadataCliente = data.Cliente
            main.MetadataProyecto = data.Proyecto
            main.MetadataPedido = data.Pedido
            main.MetadataPedidoSourceLabel = If(String.IsNullOrWhiteSpace(data.PedidoSource), "vacío", data.PedidoSource.Trim())
            main.MetadataPlano = data.Plano
            main.MetadataTitulo = data.Titulo
            main.MetadataRevision = data.Revision
            main.MetadataAutor = data.Autor
            main.MetadataFecha = data.Fecha
            main.MetadataClienteSourceLabel = data.ClienteSource
            main.MetadataProyectoSourceLabel = data.ProyectoSource
            main.MetadataPlanoSourceLabel = data.PlanoSource
            main.MetadataTituloSourceLabel = data.TituloSource
            main.MetadataRevisionSourceLabel = data.RevisionSource
            main.MetadataAutorSourceLabel = data.AutorSource
            main.MetadataFechaSourceLabel = data.FechaSource
        End If
        If Not applyPartList Then Return
        main.MetadataNumeroPartList = data.Numero
        main.MetadataCantidad = data.Cantidad
        main.MetadataNombreArchivo = data.NombreArchivo
        main.MetadataDenominacion = data.Denominacion
        main.MetadataMaterial = data.Material
        main.MetadataEspesor = data.Espesor
        main.MetadataL = data.LargoL
        main.MetadataH = data.AltoH
        main.MetadataD = data.DatoD
        main.MetadataPeso = data.Peso
        main.MetadataMaterialSourceLabel = data.MaterialSource
        main.MetadataEspesorSourceLabel = data.EspesorSource
        main.MetadataLhdSourceLabel = data.LHDSource
        main.MetadataPesoSourceLabel = data.PesoSource
        main.MetadataNumeroSourceLabel = data.NumeroSource
        main.MetadataCantidadSourceLabel = data.CantidadSource
        main.MetadataNombreArchivoSourceLabel = data.NombreArchivoSource
        main.MetadataDenominacionSourceLabel = data.DenominacionSource
    End Sub

    ''' <summary>Validación de checklist cajetín + PART_LIST para un contexto o archivo de componente.</summary>
    Public Shared Function ValidateMetadataForComponent(componentPath As String, data As DrawingMetadataInput, strict As Boolean, logger As Logger) As MetadataValidationResult
        Dim r As New MetadataValidationResult()
        If data Is Nothing Then
            r.Outcome = MetadataValidationOutcome.RequiredMissing
            r.MissingRequiredFields.Add("metadatos")
            If logger IsNot Nothing Then logger.Log("[UI][GENERATE][VALIDATE_METADATA] path=" & componentPath & " outcome=RequiredMissing")
            Return r
        End If

        SubAddIfEmpty(r.MissingRequiredFields, data.Cliente, "Cliente")
        SubAddIfEmpty(r.MissingRequiredFields, data.Proyecto, "Proyecto")
        SubAddIfEmpty(r.MissingRequiredFields, data.Pedido, "Pedido")
        SubAddIfEmpty(r.MissingRequiredFields, data.Plano, "Plano")
        SubAddIfEmpty(r.MissingRequiredFields, data.Revision, "Revisión")
        SubAddIfEmpty(r.MissingRequiredFields, data.Autor, "Autor")
        If String.IsNullOrWhiteSpace(data.FechaSource) OrElse String.Equals(data.FechaSource, "vacío", StringComparison.OrdinalIgnoreCase) Then
            r.MissingRequiredFields.Add("Fecha")
        End If

        SubAddIfEmpty(r.MissingRequiredFields, data.Numero, "Nº")
        SubAddIfEmpty(r.MissingRequiredFields, data.Cantidad, "Cantidad")
        SubAddIfEmpty(r.MissingRequiredFields, data.NombreArchivo, "Nombre archivo")
        SubAddIfEmpty(r.MissingRequiredFields, data.Denominacion, "Denominación")
        SubAddIfEmpty(r.MissingRequiredFields, data.Material, "Material")
        SubAddIfEmpty(r.MissingRequiredFields, data.Espesor, "Espesor")
        SubAddIfEmpty(r.MissingRequiredFields, data.LargoL, "L")
        SubAddIfEmpty(r.MissingRequiredFields, data.AltoH, "H")
        SubAddIfEmpty(r.MissingRequiredFields, data.DatoD, "D")
        SubAddIfEmpty(r.MissingRequiredFields, data.Peso, "Peso")

        r.MissingWarningFields.AddRange(r.MissingRequiredFields)

        If r.MissingRequiredFields.Count = 0 Then
            r.Outcome = MetadataValidationOutcome.Complete
        ElseIf strict Then
            r.Outcome = MetadataValidationOutcome.RequiredMissing
        Else
            r.Outcome = MetadataValidationOutcome.WarningMissing
        End If

        If logger IsNot Nothing Then
            logger.Log("[UI][GENERATE][VALIDATE_METADATA] path=" & If(componentPath, "") & " outcome=" & r.Outcome.ToString() &
                       " missing=" & r.MissingRequiredFields.Count.ToString(CultureInfo.InvariantCulture))
        End If
        Return r
    End Function

    Private Shared Sub SubAddIfEmpty(list As List(Of String), value As String, label As String)
        If list Is Nothing Then Return
        If String.IsNullOrWhiteSpace(value) Then list.Add(label)
    End Sub

    ''' <summary>Rellena cajetín; opcionalmente PART_LIST (PAR/PSM). ASM usa fillPartList=False.</summary>
    Public Shared Sub LoadTitleBlockFromOpenModel(doc As Object, inputPath As String, logger As Logger, ByRef data As DrawingMetadataInput,
                                                  Optional fillPartList As Boolean = True,
                                                  Optional inferPedidoFromPath As Boolean = False,
                                                  Optional traceUiLog As Boolean = False,
                                                  Optional traceKind As String = "FIELD")
        If data Is Nothing Then data = New DrawingMetadataInput()
        If doc Is Nothing Then Return

        If logger IsNot Nothing Then logger.Log("[UI][TITLEBLOCK][LOAD] Origen=documento abierto fillPartList=" & fillPartList.ToString())

        If Not fillPartList Then
            data.Material = ""
            data.Espesor = ""
            data.LargoL = ""
            data.AltoH = ""
            data.DatoD = ""
            data.Peso = ""
            data.Denominacion = ""
            data.NombreArchivo = ""
            data.Cantidad = ""
            data.Numero = ""
            data.MaterialSource = ""
            data.EspesorSource = ""
            data.LHDSource = ""
            data.PesoSource = ""
            data.NumeroSource = ""
            data.CantidadSource = ""
            data.NombreArchivoSource = ""
            data.DenominacionSource = ""
        End If

        data.Cliente = FirstNonEmpty(
            SolidEdgePropertyService.GetDocumentPropertyForMetadata(doc, "DocumentSummaryInformation", {"Company", "Empresa"}),
            SolidEdgePropertyService.GetDocumentPropertyForMetadata(doc, "Custom", {"Cliente", "Client"}))
        data.ClienteSource = If(String.IsNullOrWhiteSpace(data.Cliente), "vacío", "modelo")
        UiMetadataFieldTrace(logger, traceUiLog, traceKind, "Cliente", data.Cliente)

        data.Proyecto = FirstNonEmpty(
            SolidEdgePropertyService.GetDocumentPropertyForMetadata(doc, "ProjectInformation", {"Project Name", "Project", "Nombre de proyecto"}),
            SolidEdgePropertyService.GetDocumentPropertyForMetadata(doc, "Custom", {"Proyecto", "PROYECTO"}))
        data.ProyectoSource = If(String.IsNullOrWhiteSpace(data.Proyecto), "vacío", "modelo")
        UiMetadataFieldTrace(logger, traceUiLog, traceKind, "Proyecto", data.Proyecto)

        ResolvePedidoFromDocument(doc, inputPath, logger, data, inferPedidoFromPath)
        If String.IsNullOrWhiteSpace(data.Pedido) Then data.PedidoSource = "vacío"
        UiMetadataFieldTrace(logger, traceUiLog, traceKind, "Pedido", data.Pedido)

        data.Plano = FirstNonEmpty(
            SolidEdgePropertyService.GetDocumentPropertyForMetadata(doc, "SummaryInformation", {"Title", "Document Title", "Titulo", "Título"}),
            SolidEdgePropertyService.GetDocumentPropertyForMetadata(doc, "Custom", {"Plano"}))
        data.PlanoSource = If(String.IsNullOrWhiteSpace(data.Plano), "vacío", "modelo")
        UiMetadataFieldTrace(logger, traceUiLog, traceKind, "Plano", data.Plano)

        data.Titulo = FirstNonEmpty(
            SolidEdgePropertyService.GetDocumentPropertyForMetadata(doc, "SummaryInformation", {"Subject"}),
            SolidEdgePropertyService.GetDocumentPropertyForMetadata(doc, "Custom", {"Titulo", "Título"}))
        data.TituloSource = If(String.IsNullOrWhiteSpace(data.Titulo), "vacío", "modelo")
        UiMetadataFieldTrace(logger, traceUiLog, traceKind, "Titulo", data.Titulo)

        If fillPartList Then
            data.Denominacion = FirstNonEmpty(
                SolidEdgePropertyService.GetDocumentPropertyForMetadata(doc, "Custom", {"Denominacion", "Denominación"}))
            data.DenominacionSource = If(String.IsNullOrWhiteSpace(data.Denominacion), "vacío", "modelo")
            UiMetadataFieldTrace(logger, traceUiLog, traceKind, "Denominacion", data.Denominacion)
        End If

        data.Revision = FirstNonEmpty(
            SolidEdgePropertyService.GetDocumentPropertyForMetadata(doc, "ProjectInformation", {"Revision", "Revision Number", "Revisión"}),
            SolidEdgePropertyService.GetDocumentPropertyForMetadata(doc, "Custom", {"Revision"}))
        data.RevisionSource = If(String.IsNullOrWhiteSpace(data.Revision), "vacío", "modelo")
        UiMetadataFieldTrace(logger, traceUiLog, traceKind, "Revision", data.Revision)

        data.Autor = FirstNonEmpty(
            SolidEdgePropertyService.GetDocumentPropertyForMetadata(doc, "SummaryInformation", {"Author", "Autor"}),
            SolidEdgePropertyService.GetDocumentPropertyForMetadata(doc, "Custom", {"Autor"}))
        data.AutorSource = If(String.IsNullOrWhiteSpace(data.Autor), "vacío", "modelo")
        UiMetadataFieldTrace(logger, traceUiLog, traceKind, "Autor", data.Autor)

        Dim fechaTxt = SolidEdgePropertyService.GetDocumentPropertyForMetadata(doc, "Custom", {"FechaPlano", "Fecha"})
        Dim parsed As Date
        If Not String.IsNullOrWhiteSpace(fechaTxt) AndAlso Date.TryParse(fechaTxt, CultureInfo.CurrentCulture, DateTimeStyles.None, parsed) Then
            data.Fecha = parsed.Date
            data.FechaSource = "modelo"
            UiMetadataFieldTrace(logger, traceUiLog, traceKind, "Fecha", fechaTxt)
        Else
            data.Fecha = Date.Today
            data.FechaSource = "vacío"
            UiMetadataFieldEmpty(logger, traceUiLog, traceKind, "Fecha")
        End If

        If fillPartList Then
            TryDetectMaterial(doc, logger, data)
            UiMetadataFieldTrace(logger, traceUiLog, traceKind, "Material", data.Material)
            TryDetectEspesor(doc, Nothing, logger, data)
            If String.IsNullOrWhiteSpace(data.Espesor) Then data.EspesorSource = "Missing"
            UiMetadataFieldTrace(logger, traceUiLog, traceKind, "Espesor", data.Espesor)
            TryDetectPeso(doc, logger, data)
            If String.IsNullOrWhiteSpace(data.Peso) Then
                data.PesoSource = "Missing"
                UiMetadataFieldEmpty(logger, traceUiLog, traceKind, "Peso")
            Else
                UiMetadataFieldTrace(logger, traceUiLog, traceKind, "Peso", data.Peso)
            End If
            Try
                If Not String.IsNullOrWhiteSpace(inputPath) Then
                    data.NombreArchivo = Path.GetFileName(inputPath)
                    data.NombreArchivoSource = "modelo"
                End If
            Catch
            End Try
            UiMetadataFieldTrace(logger, traceUiLog, traceKind, "NombreArchivo", data.NombreArchivo)
            UiMetadataFieldTrace(logger, traceUiLog, traceKind, "L", data.LargoL)
            UiMetadataFieldTrace(logger, traceUiLog, traceKind, "H", data.AltoH)
            UiMetadataFieldTrace(logger, traceUiLog, traceKind, "D", data.DatoD)
        End If
    End Sub

    Private Shared Sub UiMetadataFieldTrace(logger As Logger, traceUiLog As Boolean, traceKind As String, fieldName As String, value As String)
        If Not traceUiLog OrElse logger Is Nothing Then Return
        If String.IsNullOrWhiteSpace(value) Then
            logger.Log("[UI][METADATA][" & traceKind & "_EMPTY] field=" & fieldName)
        Else
            logger.Log("[UI][METADATA][" & traceKind & "_FOUND] field=" & fieldName)
        End If
    End Sub

    Private Shared Sub UiMetadataFieldEmpty(logger As Logger, traceUiLog As Boolean, traceKind As String, fieldName As String)
        If Not traceUiLog OrElse logger Is Nothing Then Return
        logger.Log("[UI][METADATA][" & traceKind & "_EMPTY] field=" & fieldName)
    End Sub

    ''' <summary>Reglas explícitas PAR/PSM: nombre fichero, plano si vacío, Nº y cantidad = 1.</summary>
    Public Shared Sub ApplySinglePartExplicitRules(modelPath As String, data As DrawingMetadataInput, logger As Logger)
        If data Is Nothing OrElse String.IsNullOrWhiteSpace(modelPath) Then Return
        Try
            data.NombreArchivo = Path.GetFileName(modelPath)
            data.NombreArchivoSource = "calculado (nombre fichero)"
        Catch
        End Try
        If String.IsNullOrWhiteSpace(data.Plano) Then
            data.Plano = InferPlanoFromFileName(modelPath)
            data.PlanoSource = If(String.IsNullOrWhiteSpace(data.Plano), "vacío", "calculado (nombre fichero)")
            If logger IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(data.Plano) Then
                logger.Log("[METADATA][PLANO][FROM_FILENAME] " & data.Plano)
            End If
        End If
        data.Cantidad = "1"
        data.CantidadSource = "calculado"
        data.Numero = "1"
        data.NumeroSource = "calculado"
    End Sub

    ''' <summary>Resuelve Pedido desde propiedades del documento (mismas claves que suelen enlazarse al cajetín «PEDIDO»).</summary>
    Private Shared Sub ResolvePedidoFromDocument(doc As Object, inputPath As String, logger As Logger, data As DrawingMetadataInput, Optional inferFromPath As Boolean = False)
        If data Is Nothing Then Return
        data.Pedido = ""
        data.PedidoSource = ""

        Dim customKeys As String() = {
            "Pedido", "PEDIDO",
            "Order", "ORDER", "Purchase Order", "PurchaseOrder",
            "NumeroPedido", "Número de pedido", "Numero de pedido",
            "Nº pedido", "No pedido", "No. pedido", "N.º pedido",
            "RefPedido", "REF_PEDIDO", "Ref. pedido", "PedidoCliente",
            "OC", "Orden de compra", "OrdenCompra"
        }
        For Each ck In customKeys
            Dim v = SolidEdgePropertyService.GetDocumentPropertyForMetadata(doc, "Custom", {ck})
            If Not String.IsNullOrWhiteSpace(v) Then
                data.Pedido = v.Trim()
                data.PedidoSource = "Custom." & ck
                If logger IsNot Nothing Then logger.Log("[METADATA][PEDIDO][FROM] " & data.PedidoSource & " = " & data.Pedido)
                Return
            End If
        Next

        ' Algunas plantillas enlazan «pedido» al nº de existencias / stock (convención ERP).
        Dim stockKeys As String() = {"Stock Number", "Número de existencias"}
        Dim stock = SolidEdgePropertyService.GetDocumentPropertyForMetadata(doc, "ProjectInformation", stockKeys)
        If Not String.IsNullOrWhiteSpace(stock) Then
            data.Pedido = stock.Trim()
            data.PedidoSource = "ProjectInformation.Stock Number (o equivalente)"
            If logger IsNot Nothing Then logger.Log("[METADATA][PEDIDO][FROM] " & data.PedidoSource & " = " & data.Pedido)
            Return
        End If

        If Not inferFromPath Then Return
        data.Pedido = InferPedidoFromPath(inputPath, logger)
        If Not String.IsNullOrWhiteSpace(data.Pedido) Then
            data.PedidoSource = "Inferido (carpeta o nombre de archivo)"
        End If
    End Sub

    Public Shared Function InferPedidoFromPath(inputPath As String, logger As Logger) As String
        If String.IsNullOrWhiteSpace(inputPath) Then Return ""
        Dim folder As String = ""
        Try
            folder = Path.GetFileName(Path.GetDirectoryName(inputPath))
        Catch
            folder = ""
        End Try
        Dim candidates As String() = {folder, Path.GetFileNameWithoutExtension(inputPath)}
        For Each c In candidates
            If String.IsNullOrWhiteSpace(c) Then Continue For
            If Regex.IsMatch(c, "\d{3,}[-_]\d{2,}[-_]\d{2,}", RegexOptions.IgnoreCase) Then
                If logger IsNot Nothing Then logger.Log("[METADATA][PEDIDO][INFERRED] " & c)
                Return c
            End If
        Next
        Return ""
    End Function

    ''' <summary>Nº de plano desde nombre de archivo: texto antes del primer espacio (p. ej. <c>1084...001 LM1 Chapa</c> → <c>1084...001</c>); sin espacios, todo el nombre sin extensión.</summary>
    Public Shared Function InferPlanoFromFileName(filePath As String) As String
        If String.IsNullOrWhiteSpace(filePath) Then Return ""
        Dim name As String = ""
        Try
            name = Path.GetFileNameWithoutExtension(filePath).Trim()
        Catch
            Return ""
        End Try
        If name.Length = 0 Then Return ""
        Dim sp = name.IndexOfAny(New Char() {" "c, ChrW(9)})
        If sp > 0 Then Return name.Substring(0, sp).TrimEnd()
        Return name
    End Function

    ''' <summary>Denominación PART_LIST: texto tras el primer espacio (<c>...001 LM1 Chapa Cuello</c> → <c>LM1 Chapa Cuello</c>).</summary>
    Public Shared Function InferDenominacionListaFromFileName(filePath As String) As String
        If String.IsNullOrWhiteSpace(filePath) Then Return ""
        Dim name As String = ""
        Try
            name = Path.GetFileNameWithoutExtension(filePath).Trim()
        Catch
            Return ""
        End Try
        If name.Length = 0 Then Return ""
        Dim sp = name.IndexOfAny(New Char() {" "c, ChrW(9)})
        If sp < 0 OrElse sp >= name.Length - 1 Then Return ""
        Return name.Substring(sp + 1).Trim()
    End Function

    Public Shared Sub TryDetectMaterial(doc As Object, logger As Logger, data As DrawingMetadataInput)
        If doc Is Nothing OrElse data Is Nothing Then Return
        Dim v = FirstNonEmpty(
            SolidEdgePropertyService.GetDocumentPropertyForMetadata(doc, "MechanicalModeling", {"Material"}),
            SolidEdgePropertyService.GetDocumentPropertyForMetadata(doc, "Custom", {"Material"}),
            SolidEdgePropertyService.GetDocumentPropertyForMetadata(doc, "ProjectInformation", {"Material"}))

        If Not String.IsNullOrWhiteSpace(v) Then
            data.Material = v.Trim()
            data.MaterialSource = "Model.MechanicalModeling.Material"
            If logger IsNot Nothing Then logger.Log("[PARTLISTDATA][MATERIAL][FOUND] " & v)
            Return
        End If

        Dim inspected = SolidEdgePropertyService.TryFindMaterialByPropertyScan(doc)
        If Not String.IsNullOrWhiteSpace(inspected) Then
            data.Material = inspected.Trim()
            data.MaterialSource = "Model.PropertyScan"
            If logger IsNot Nothing Then logger.Log("[PARTLISTDATA][MATERIAL][FOUND] scan=" & inspected)
            Return
        End If

        data.Material = ""
        data.MaterialSource = "Missing"
        If logger IsNot Nothing Then logger.Log("[PARTLISTDATA][MATERIAL][MISSING]")
    End Sub

    Public Shared Sub TryDetectEspesor(modelDoc As Object, draftDoc As Object, logger As Logger, data As DrawingMetadataInput)
        If data Is Nothing Then Return
        Dim v = ""
        If modelDoc IsNot Nothing Then
            v = FirstNonEmpty(
                SolidEdgePropertyService.GetDocumentPropertyForMetadata(modelDoc, "Custom", {"Espesor"}),
                SolidEdgePropertyService.GetDocumentPropertyForMetadata(modelDoc, "Custom", {"Espesor del material"}),
                SolidEdgePropertyService.GetDocumentPropertyForMetadata(modelDoc, "MechanicalModeling", {"Sheet Metal Gauge", "Thickness"}))
        End If
        If Not String.IsNullOrWhiteSpace(v) Then
            data.Espesor = NormalizeThicknessDisplay(v)
            data.EspesorSource = "Model"
            If logger IsNot Nothing Then logger.Log("[PARTLISTDATA][ESPESOR][FOUND_MODEL] " & data.Espesor)
            Return
        End If

        If draftDoc IsNot Nothing Then
            Dim fromDraft = TryInferThicknessFromDraftDimensions(draftDoc, logger)
            If Not String.IsNullOrWhiteSpace(fromDraft) Then
                data.Espesor = fromDraft
                data.EspesorSource = "DFT"
                If logger IsNot Nothing Then logger.Log("[PARTLISTDATA][ESPESOR][FOUND_DRAFT] " & fromDraft)
                Return
            End If
        End If

        data.Espesor = ""
        data.EspesorSource = "Missing"
        If logger IsNot Nothing Then logger.Log("[PARTLISTDATA][ESPESOR][MISSING]")
    End Sub

    Private Shared Function TryInferThicknessFromDraftDimensions(draftDoc As Object, logger As Logger) As String
        If draftDoc Is Nothing Then Return ""
        Try
            Dim sections As Object = Nothing
            Try : sections = CallByName(draftDoc, "Sections", CallType.Get) : Catch : End Try
            If sections Is Nothing Then Return ""
            Dim secCount As Integer = SafeInt(CallByName(sections, "Count", CallType.Get), 0)
            Dim smallMm As New List(Of Double)()
            For si As Integer = 1 To secCount
                Dim section As Object = GetItem(sections, si)
                If section Is Nothing Then Continue For
                Dim sheets As Object = Nothing
                Try : sheets = CallByName(section, "Sheets", CallType.Get) : Catch : End Try
                If sheets Is Nothing Then Continue For
                Dim shCount As Integer = SafeInt(CallByName(sheets, "Count", CallType.Get), 0)
                For shi As Integer = 1 To shCount
                    Dim sheet As Object = GetItem(sheets, shi)
                    If sheet Is Nothing Then Continue For
                    CollectSmallLinearDimensionsMm(sheet, smallMm)
                Next
            Next
            If smallMm.Count = 0 Then Return ""
            Dim pick = smallMm.Where(Function(x) x > 0.2 AndAlso x <= 30).OrderBy(Function(x) x).FirstOrDefault()
            If pick <= 0 Then Return ""
            Dim s = FormatNumberEs(pick, 2)
            If logger IsNot Nothing Then logger.Log("[PARTLISTDATA][ESPESOR][DRAFT_SMALL_DIM] mm=" & s)
            Return s
        Catch
            Return ""
        End Try
    End Function

    Private Shared Sub CollectSmallLinearDimensionsMm(sheet As Object, sink As List(Of Double))
        If sheet Is Nothing OrElse sink Is Nothing Then Return
        Dim dimsObj As Object = Nothing
        Try : dimsObj = CallByName(sheet, "Dimensions", CallType.Get) : Catch : End Try
        If dimsObj Is Nothing Then Return
        Dim n As Integer = SafeInt(CallByName(dimsObj, "Count", CallType.Get), 0)
        For i As Integer = 1 To n
            Dim d As Object = Nothing
            Try : d = GetItem(dimsObj, i) : Catch : End Try
            If d Is Nothing Then Continue For
            If IsRadialDimension(d) Then Continue For
            Dim mm As Double = ReadDimensionValueMm(d)
            If mm > 0 Then sink.Add(mm)
        Next
    End Sub

    Private Shared Function IsRadialDimension(d As Object) As Boolean
        If d Is Nothing Then Return False
        Try
            Dim t = Convert.ToString(CallByName(d, "DimensionType", CallType.Get), CultureInfo.InvariantCulture).ToUpperInvariant()
            If t.Contains("RADIUS") OrElse t.Contains("DIAMETER") OrElse t.Contains("RADIAL") Then Return True
        Catch
        End Try
        Return False
    End Function

    Private Shared Function ReadDimensionValueMm(d As Object) As Double
        If d Is Nothing Then Return 0
        Try
            Dim v As Double = Convert.ToDouble(CallByName(d, "Value", CallType.Get), CultureInfo.InvariantCulture)
            If Double.IsNaN(v) OrElse Double.IsInfinity(v) Then Return 0
            Return v * 1000.0R
        Catch
            Return 0
        End Try
    End Function

    Private Const PhysPropAccuracy As Double = 0.0001R

    ''' <summary>Masa del primer Model (PAR/PSM): MassProperties.Calculate si existe; si no ComputePhysicalProperties / GetPhysicalProperties.</summary>
    Public Shared Function GetPesoModelo(doc As Object, logger As Logger) As String
        If doc Is Nothing Then Return ""
        Dim m As Model = Nothing
        Dim pd As PartDocument = TryCast(doc, PartDocument)
        If pd IsNot Nothing AndAlso pd.Models IsNot Nothing AndAlso pd.Models.Count >= 1 Then
            m = pd.Models.Item(1)
        Else
            Dim sm As SheetMetalDocument = TryCast(doc, SheetMetalDocument)
            If sm IsNot Nothing AndAlso sm.Models IsNot Nothing AndAlso sm.Models.Count >= 1 Then
                m = sm.Models.Item(1)
            End If
        End If
        If m Is Nothing Then
            If logger IsNot Nothing Then logger.Log("[PESO][MISSING] Sin PartDocument/SheetMetalDocument o Models(1).")
            Return ""
        End If

        Dim massKg As Double = 0
        If Not TryGetMassKgFromModel(m, logger, massKg) OrElse massKg <= 0 OrElse Double.IsNaN(massKg) OrElse Double.IsInfinity(massKg) Then
            If logger IsNot Nothing Then logger.Log("[PESO][MISSING] Masa no disponible desde geometría.")
            Return ""
        End If

        Dim formatted = FormatPesoKgEs(massKg)
        If logger IsNot Nothing Then
            logger.Log("[PESO][FOUND] value=" & massKg.ToString(CultureInfo.InvariantCulture) & " source=Model.MassProperties")
        End If
        Return formatted
    End Function

    Private Shared Function FormatPesoKgEs(massKg As Double) As String
        Return FormatNumberEs(massKg, 3) & " kg"
    End Function

    Private Shared Function TryGetMassKgFromModel(m As Model, logger As Logger, ByRef massKg As Double) As Boolean
        massKg = 0
        If m Is Nothing Then Return False

        Try
            Dim mp As Object = Nothing
            Try : mp = CallByName(m, "MassProperties", CallType.Get) : Catch : End Try
            If mp IsNot Nothing Then
                Try : CallByName(mp, "Calculate", CallType.Method) : Catch : End Try
                Try
                    massKg = Convert.ToDouble(CallByName(mp, "Mass", CallType.Get), CultureInfo.InvariantCulture)
                    If massKg > 0 Then Return True
                Catch
                End Try
            End If
        Catch ex As Exception
            If logger IsNot Nothing Then logger.Log("[PESO][ERROR] MassProperties: " & ex.Message)
        End Try

        Dim st As Integer
        Dim dens As Double, acc As Double, vol As Double, area As Double, mass As Double, accOut As Double
        Dim cog As Double() = New Double(2) {}
        Dim cov As Double() = New Double(2) {}
        Dim g6 As Double() = New Double(5) {}
        Dim p3 As Double() = New Double(2) {}
        Dim p9 As Double() = New Double(8) {}
        Dim r3 As Double() = New Double(2) {}

        Try
            m.GetPhysicalProperties(st, dens, acc, vol, area, mass, cog, cov, g6, p3, p9, r3, accOut)
            If mass > 0 AndAlso st = 1 Then
                massKg = mass
                Return True
            End If
        Catch ex As Exception
            If logger IsNot Nothing Then logger.Log("[PESO][ERROR] GetPhysicalProperties: " & ex.Message)
        End Try

        Try
            Dim densityUse As Double = dens
            If densityUse < 0 Then densityUse = 0
            mass = 0 : vol = 0 : area = 0 : accOut = 0 : st = 0
            m.ComputePhysicalProperties(densityUse, PhysPropAccuracy, vol, area, mass, cog, cov, g6, p3, p9, r3, accOut, st)
            If mass > 0 Then
                massKg = mass
                Return True
            End If

            If densityUse <= 0 Then
                mass = 0 : vol = 0 : area = 0 : accOut = 0 : st = 0
                m.ComputePhysicalProperties(7850.0R, PhysPropAccuracy, vol, area, mass, cog, cov, g6, p3, p9, r3, accOut, st)
                If mass > 0 Then
                    massKg = mass
                    Return True
                End If
            End If
        Catch ex As Exception
            If logger IsNot Nothing Then logger.Log("[PESO][ERROR] ComputePhysicalProperties: " & ex.Message)
        End Try

        Return False
    End Function

    Public Shared Sub TryDetectPeso(doc As Object, logger As Logger, data As DrawingMetadataInput)
        If doc Is Nothing OrElse data Is Nothing Then Return

        Dim fromGeom = GetPesoModelo(doc, logger)
        If Not String.IsNullOrWhiteSpace(fromGeom) Then
            data.Peso = fromGeom
            data.PesoSource = "Model.MassProperties"
            Return
        End If

        Dim customP = SolidEdgePropertyService.GetDocumentPropertyForMetadata(doc, "Custom", {"Peso"})
        If Not String.IsNullOrWhiteSpace(customP) Then
            data.Peso = FormatWeightForUi(customP)
            data.PesoSource = "Custom.Peso"
            If logger IsNot Nothing Then logger.Log("[PESO][FOUND] value=" & data.Peso & " source=Custom.Peso")
            Return
        End If

        data.Peso = ""
        data.PesoSource = "Missing"
        If logger IsNot Nothing Then logger.Log("[PESO][MISSING]")
    End Sub

    Public Shared Function TryComputeLhdFromDraft(draftDoc As Object, logger As Logger, ByRef L As String, ByRef H As String, ByRef D As String) As Boolean
        L = "" : H = "" : D = ""
        If draftDoc Is Nothing Then Return False
        Dim values As List(Of Double) = Nothing
        If TryCollectDistinctLengthsFromDimensions(draftDoc, logger, values) AndAlso values.Count >= 3 Then
            AssignLhd(values, L, H, D)
            If logger IsNot Nothing Then
                logger.Log("[PARTLISTDATA][LHD][FROM_DIMENSIONS]")
                logger.Log("[PARTLISTDATA][LHD][VALUES_UNIQUE] " & String.Join(",", values.Take(12).Select(Function(x) FormatNumberEs(x, 2))))
                logger.Log("[PARTLISTDATA][LHD][ASSIGN] L=" & L & " H=" & H & " D=" & D)
            End If
            Return True
        End If

        If TryCollectFromViewBboxes(draftDoc, logger, values) AndAlso values.Count >= 3 Then
            AssignLhd(values, L, H, D)
            If logger IsNot Nothing Then
                logger.Log("[PARTLISTDATA][LHD][FROM_DV_BBOX]")
                logger.Log("[PARTLISTDATA][LHD][VALUES_UNIQUE] " & String.Join(",", values.Take(12).Select(Function(x) FormatNumberEs(x, 2))))
                logger.Log("[PARTLISTDATA][LHD][ASSIGN] L=" & L & " H=" & H & " D=" & D)
            End If
            Return True
        End If
        Return False
    End Function

    ''' <summary>Caja envolvente 3D: Model.Body (SolidEdgeGeometry.Body) y GetRange (metros → mm), L≥H≥D. PAR/PSM.</summary>
    Public Shared Function TryComputeLhdFromModelDoc(modelDoc As Object, logger As Logger, ByRef L As String, ByRef H As String, ByRef D As String) As Boolean
        L = "" : H = "" : D = ""
        If modelDoc Is Nothing Then Return False
        Try
            Dim models As Models = Nothing
            Dim pd As PartDocument = TryCast(modelDoc, PartDocument)
            If pd IsNot Nothing Then
                models = pd.Models
            Else
                Dim sm As SheetMetalDocument = TryCast(modelDoc, SheetMetalDocument)
                If sm IsNot Nothing Then models = sm.Models
            End If
            If models Is Nothing OrElse models.Count < 1 Then Return False
            Dim m0 As Model = models.Item(1)
            If m0 Is Nothing Then Return False
            Dim bodyObj As Body = m0.Body
            If bodyObj Is Nothing Then Return False

            Dim minPt As Double() = New Double(2) {}
            Dim maxPt As Double() = New Double(2) {}
            bodyObj.GetRange(minPt, maxPt)

            Dim dx = Math.Abs(maxPt(0) - minPt(0)) * 1000.0R
            Dim dy = Math.Abs(maxPt(1) - minPt(1)) * 1000.0R
            Dim dz = Math.Abs(maxPt(2) - minPt(2)) * 1000.0R
            If dx < 0.01 AndAlso dy < 0.01 AndAlso dz < 0.01 Then Return False
            Dim dims As New List(Of Double) From {dx, dy, dz}
            dims = dims.OrderByDescending(Function(x) x).ToList()
            AssignLhd(dims, L, H, D)
            If logger IsNot Nothing Then
                logger.Log("[PARTLISTDATA][LHD][FROM_BODY_GETRANGE]")
                logger.Log("[PARTLISTDATA][LHD][ASSIGN] L=" & L & " H=" & H & " D=" & D)
            End If
            Return True
        Catch ex As Exception
            If logger IsNot Nothing Then logger.Log("[PARTLISTDATA][LHD][MODEL][ERR] " & ex.Message)
            Return False
        End Try
    End Function

    Private Shared Function TryCollectDistinctLengthsFromDimensions(draftDoc As Object, logger As Logger, ByRef mmValues As List(Of Double)) As Boolean
        mmValues = New List(Of Double)()
        Try
            Dim sections As Object = Nothing
            Try : sections = CallByName(draftDoc, "Sections", CallType.Get) : Catch : End Try
            If sections Is Nothing Then Return False
            For si As Integer = 1 To SafeInt(CallByName(sections, "Count", CallType.Get), 0)
                Dim section As Object = GetItem(sections, si)
                If section Is Nothing Then Continue For
                Dim sheets As Object = Nothing
                Try : sheets = CallByName(section, "Sheets", CallType.Get) : Catch : End Try
                If sheets Is Nothing Then Continue For
                For shi As Integer = 1 To SafeInt(CallByName(sheets, "Count", CallType.Get), 0)
                    Dim sheet As Object = GetItem(sheets, shi)
                    If sheet Is Nothing Then Continue For
                    Dim dimsObj As Object = Nothing
                    Try : dimsObj = CallByName(sheet, "Dimensions", CallType.Get) : Catch : End Try
                    If dimsObj Is Nothing Then Continue For
                    For di As Integer = 1 To SafeInt(CallByName(dimsObj, "Count", CallType.Get), 0)
                        Dim d As Object = GetItem(dimsObj, di)
                        If d Is Nothing OrElse IsRadialDimension(d) Then Continue For
                        Dim mm = ReadDimensionValueMm(d)
                        If mm > 0.05 Then mmValues.Add(mm)
                    Next
                Next
            Next
        Catch
        End Try
        mmValues = DeduplicateMm(mmValues, TolMm)
        mmValues = mmValues.OrderByDescending(Function(x) x).ToList()
        Return mmValues.Count >= 3
    End Function

    Private Shared Function TryCollectFromViewBboxes(draftDoc As Object, logger As Logger, ByRef mmValues As List(Of Double)) As Boolean
        mmValues = New List(Of Double)()
        Try
            Dim sections As Object = CallByName(draftDoc, "Sections", CallType.Get)
            If sections Is Nothing Then Return False
            For si As Integer = 1 To SafeInt(CallByName(sections, "Count", CallType.Get), 0)
                Dim section As Object = GetItem(sections, si)
                If section Is Nothing Then Continue For
                Dim sheets As Object = CallByName(section, "Sheets", CallType.Get)
                If sheets Is Nothing Then Continue For
                For shi As Integer = 1 To SafeInt(CallByName(sheets, "Count", CallType.Get), 0)
                    Dim sheet As Sheet = TryCast(GetItem(sheets, shi), Sheet)
                    If sheet Is Nothing Then Continue For
                    For vi As Integer = 1 To sheet.DrawingViews.Count
                        Dim dv As DrawingView = Nothing
                        Try : dv = CType(sheet.DrawingViews.Item(vi), DrawingView) : Catch : End Try
                        If dv Is Nothing Then Continue For
                        Dim sc As Double = 1.0
                        Try : sc = CDbl(dv.Scale) : Catch : End Try
                        If sc <= 1.0E-9 Then sc = 1.0
                        Dim box As ViewSheetBoundingBox
                        If Not ViewGeometryReader.TryReadBoundingBox(dv, Nothing, box) Then Continue For
                        Dim wM = box.Width / sc
                        Dim hM = box.Height / sc
                        If wM > 1.0E-6 Then mmValues.Add(wM * 1000.0R)
                        If hM > 1.0E-6 Then mmValues.Add(hM * 1000.0R)
                    Next
                Next
            Next
        Catch ex As Exception
            If logger IsNot Nothing Then logger.Log("[PARTLISTDATA][LHD][BBOX][ERR] " & ex.Message)
        End Try
        mmValues = DeduplicateMm(mmValues, TolMm)
        mmValues = mmValues.OrderByDescending(Function(x) x).ToList()
        Return mmValues.Count >= 3
    End Function

    Private Shared Sub AssignLhd(sortedDesc As List(Of Double), ByRef L As String, ByRef H As String, ByRef D As String)
        If sortedDesc Is Nothing OrElse sortedDesc.Count < 3 Then
            L = "" : H = "" : D = ""
            Return
        End If
        L = FormatIntMm(sortedDesc(0))
        H = FormatIntMm(sortedDesc(1))
        D = FormatIntMm(sortedDesc(2))
    End Sub

    Private Shared Function FormatIntMm(x As Double) As String
        Return CInt(Math.Round(x)).ToString(CultureInfo.InvariantCulture)
    End Function

    Private Shared Function DeduplicateMm(values As IEnumerable(Of Double), tol As Double) As List(Of Double)
        Dim sorted = values.OrderByDescending(Function(x) x).ToList()
        Dim out As New List(Of Double)()
        For Each v In sorted
            Dim dup As Boolean = False
            For Each o In out
                If Math.Abs(o - v) <= tol Then dup = True : Exit For
            Next
            If Not dup Then out.Add(v)
        Next
        Return out
    End Function

    Public Shared Sub ApplyPartListSourceProperties(modelDoc As Object, draftDoc As Object, data As DrawingMetadataInput, logger As Logger)
        If data Is Nothing Then Return
        If logger IsNot Nothing Then logger.Log("[PARTLISTDATA][WRITE] Inicio ApplyPartListSourceProperties")

        Dim pairs As New List(Of KeyValuePair(Of String, String))()
        SubAdd(pairs, "Material", data.Material)
        SubAdd(pairs, "Espesor", data.Espesor)
        SubAdd(pairs, "L", data.LargoL)
        SubAdd(pairs, "H", data.AltoH)
        SubAdd(pairs, "D", data.DatoD)
        SubAdd(pairs, "Peso", data.Peso)
        SubAdd(pairs, "Denominacion", If(String.IsNullOrWhiteSpace(data.Denominacion), data.Titulo, data.Denominacion))
        SubAdd(pairs, "NombreArchivo", data.NombreArchivo)
        SubAdd(pairs, "Cantidad", data.Cantidad)

        For Each kv In pairs
            If String.IsNullOrWhiteSpace(kv.Value) Then Continue For
            If modelDoc IsNot Nothing Then
                If SolidEdgePropertyService.TrySetCustomProperty(modelDoc, kv.Key, kv.Value, logger) Then
                    If logger IsNot Nothing Then logger.Log("[PARTLISTDATA][WRITE_MODEL_PROP] Custom." & kv.Key & "=" & kv.Value)
                End If
            End If
            If draftDoc IsNot Nothing Then
                If SolidEdgePropertyService.TrySetCustomProperty(draftDoc, kv.Key, kv.Value, logger) Then
                    If logger IsNot Nothing Then logger.Log("[PARTLISTDATA][WRITE_DFT_PROP] Custom." & kv.Key & "=" & kv.Value)
                End If
            End If
        Next

        If draftDoc IsNot Nothing Then
            SolidEdgePropertyService.RefreshDraftFromModelLinks(draftDoc, logger, False)
            SolidEdgePropertyService.RefreshNativePartsListsAndUpdateAll(draftDoc, logger)
        End If
        If logger IsNot Nothing Then logger.Log("[PARTLISTDATA][WRITE] Fin")
    End Sub

    Private Shared Sub SubAdd(list As List(Of KeyValuePair(Of String, String)), k As String, v As String)
        If list Is Nothing OrElse String.IsNullOrWhiteSpace(v) Then Return
        list.Add(New KeyValuePair(Of String, String)(k, v.Trim()))
    End Sub

    Public Shared Function TryLoadMetadataFromModelFile(modelPath As String, showSe As Boolean, logger As Logger, ByRef data As DrawingMetadataInput) As Boolean
        data = New DrawingMetadataInput()
        If String.IsNullOrWhiteSpace(modelPath) OrElse Not File.Exists(modelPath) Then Return False
        Dim ext As String = Path.GetExtension(modelPath).ToLowerInvariant()
        Dim isAsm As Boolean = ext = ".asm"
        Dim app As Application = Nothing
        Dim created As Boolean = False
        Dim doc As Object = Nothing
        Try
            OleMessageFilter.Register()
            If Not TryConnectForMetadata(showSe, logger, app, created) Then Return False
            doc = app.Documents.Open(modelPath)
            If isAsm Then
                If logger IsNot Nothing Then logger.Log("[UI][METADATA][LOAD_ASM] path=" & modelPath)
                LoadTitleBlockFromOpenModel(doc, modelPath, logger, data, fillPartList:=False, inferPedidoFromPath:=False, traceUiLog:=True, traceKind:="ASM_FIELD")
                If logger IsNot Nothing Then
                    For Each fn In New String() {"Material", "Espesor", "Denominacion", "NombreArchivo", "Cantidad", "Numero", "L", "H", "D", "Peso"}
                        logger.Log("[UI][METADATA][ASM_FIELD_EMPTY] field=" & fn)
                    Next
                End If
            Else
                If logger IsNot Nothing Then logger.Log("[UI][METADATA][LOAD_SINGLE_PART] path=" & modelPath)
                LoadTitleBlockFromOpenModel(doc, modelPath, logger, data, fillPartList:=True, inferPedidoFromPath:=False, traceUiLog:=True, traceKind:="FIELD")
                ApplySinglePartExplicitRules(modelPath, data, logger)
            End If
            Return True
        Catch ex As Exception
            If logger IsNot Nothing Then logger.LogException("TryLoadMetadataFromModelFile", ex)
            Return False
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
    End Function

    Private Shared Function TryConnectForMetadata(showSe As Boolean, logger As Logger, ByRef app As Application, ByRef created As Boolean) As Boolean
        app = Nothing
        created = False
        Try
            app = CType(System.Runtime.InteropServices.Marshal.GetActiveObject("SolidEdge.Application"), Application)
        Catch
            Dim t = Type.GetTypeFromProgID("SolidEdge.Application")
            app = CType(Activator.CreateInstance(t), Application)
            created = True
        End Try
        If app Is Nothing Then Return False
        app.Visible = showSe
        app.DisplayAlerts = False
        Return True
    End Function

    Public Shared Function ValidateMetadata(data As DrawingMetadataInput, strict As Boolean, logger As Logger) As List(Of String)
        Dim msgs As New List(Of String)()
        Dim vr = ValidateMetadataForComponent("", data, strict, Nothing)
        For Each f In vr.MissingRequiredFields
            msgs.Add(If(strict, "ERROR: " & f & " vacío.", "AVISO: " & f & " vacío."))
        Next
        Dim hasErr As Boolean = msgs.Any(Function(m) m.StartsWith("ERROR:", StringComparison.OrdinalIgnoreCase))
        If logger IsNot Nothing Then
            If msgs.Count = 0 OrElse (Not strict AndAlso Not hasErr) Then
                logger.Log("[METADATA][VALIDATE][OK]")
            Else
                For Each m In msgs
                    logger.Log("[METADATA][VALIDATE][WARN] " & m)
                Next
            End If
        End If
        Return msgs
    End Function

    ''' <summary>Etiquetas de campos vacíos en cajetín/PART_LIST (informativo; no bloquea generación de DFT).</summary>
    Public Shared Function GetPlanoMetadataGapLabels(data As DrawingMetadataInput) As List(Of String)
        Dim gaps As New List(Of String)()
        If data Is Nothing Then Return gaps
        If String.IsNullOrWhiteSpace(data.Cliente) Then gaps.Add("Cliente")
        If String.IsNullOrWhiteSpace(data.Proyecto) Then gaps.Add("Proyecto")
        If String.IsNullOrWhiteSpace(data.Pedido) Then gaps.Add("Pedido")
        If String.IsNullOrWhiteSpace(data.Plano) Then gaps.Add("Plano")
        If String.IsNullOrWhiteSpace(data.Titulo) Then gaps.Add("Título")
        If String.IsNullOrWhiteSpace(data.Revision) Then gaps.Add("Revisión")
        If String.IsNullOrWhiteSpace(data.Material) Then gaps.Add("Material")
        If String.IsNullOrWhiteSpace(data.Espesor) Then gaps.Add("Espesor")
        If String.IsNullOrWhiteSpace(data.LargoL) OrElse String.IsNullOrWhiteSpace(data.AltoH) OrElse String.IsNullOrWhiteSpace(data.DatoD) Then gaps.Add("L/H/D")
        If String.IsNullOrWhiteSpace(data.Peso) Then gaps.Add("Peso")
        If String.IsNullOrWhiteSpace(data.Denominacion) Then gaps.Add("Denominación (lista)")
        If String.IsNullOrWhiteSpace(data.NombreArchivo) Then gaps.Add("Nombre archivo (lista)")
        If String.IsNullOrWhiteSpace(data.Numero) Then gaps.Add("Nº")
        If String.IsNullOrWhiteSpace(data.Cantidad) Then gaps.Add("Cantidad")
        If String.IsNullOrWhiteSpace(data.FechaSource) OrElse String.Equals(data.FechaSource, "vacío", StringComparison.OrdinalIgnoreCase) Then gaps.Add("Fecha")
        Return gaps
    End Function

    Public Shared Function FormatMetadataSnapshotLines(data As DrawingMetadataInput) As String
        If data Is Nothing Then Return ""
        Dim sb As New StringBuilder()
        sb.AppendLine("Plano: " & If(data.Plano, "").Trim())
        sb.AppendLine("Título: " & If(data.Titulo, "").Trim())
        sb.AppendLine("Denominación (lista): " & If(data.Denominacion, "").Trim())
        sb.AppendLine("Material / Espesor: " & If(data.Material, "").Trim() & " | " & If(data.Espesor, "").Trim())
        sb.AppendLine("L/H/D: " & If(data.LargoL, "").Trim() & " × " & If(data.AltoH, "").Trim() & " × " & If(data.DatoD, "").Trim())
        sb.AppendLine("Peso: " & If(data.Peso, "").Trim())
        sb.AppendLine("Cliente / Proyecto / Pedido: " & If(data.Cliente, "").Trim() & " | " & If(data.Proyecto, "").Trim() & " | " & If(data.Pedido, "").Trim())
        If Not String.IsNullOrWhiteSpace(data.PedidoSource) Then
            sb.AppendLine("Origen campo Pedido (propiedad Solid Edge): " & data.PedidoSource.Trim())
        End If
        Return sb.ToString().TrimEnd()
    End Function

    Private Shared Sub SubVal(msgs As List(Of String), strict As Boolean, value As String, label As String)
        If Not String.IsNullOrWhiteSpace(value) Then Return
        msgs.Add(If(strict, "ERROR: " & label & " vacío.", "AVISO: " & label & " vacío."))
    End Sub

    Private Shared Function FirstNonEmpty(ParamArray values() As String) As String
        If values Is Nothing Then Return ""
        For Each value In values
            If Not String.IsNullOrWhiteSpace(value) Then Return value.Trim()
        Next
        Return ""
    End Function

    Private Shared Function NormalizeThicknessDisplay(s As String) As String
        If String.IsNullOrWhiteSpace(s) Then Return ""
        Return s.Trim()
    End Function

    Private Shared Function FormatWeightForUi(raw As String) As String
        If String.IsNullOrWhiteSpace(raw) Then Return ""
        Dim s = raw.Trim()
        If s.ToLowerInvariant().Contains("kg") Then Return s
        Dim d As Double
        If Double.TryParse(s.Replace(",", "."), NumberStyles.Any, CultureInfo.InvariantCulture, d) Then
            Return FormatNumberEs(d, 3) & " kg"
        End If
        Return s
    End Function

    Private Shared Function FormatNumberEs(value As Double, decimals As Integer) As String
        Return Math.Round(value, decimals).ToString("0." & New String("0"c, decimals), New CultureInfo("es-ES"))
    End Function

    Private Shared Function SafeInt(o As Object, def As Integer) As Integer
        Try
            Return CInt(o)
        Catch
            Return def
        End Try
    End Function

    Private Shared Function GetItem(coll As Object, index As Integer) As Object
        If coll Is Nothing Then Return Nothing
        Try
            Return CallByName(coll, "Item", CallType.Get, index)
        Catch
            Try
                Return CallByName(coll, "Item", CallType.Method, index)
            Catch
                Return Nothing
            End Try
        End Try
    End Function

End Class
