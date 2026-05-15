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

    ''' <summary>Metadatos de cajetín y fuente PART_LIST a partir del mismo snapshot que llega del motor (<see cref="JobConfiguration"/> en ejecución).</summary>
    Public Shared Function BuildFromJobConfiguration(job As JobConfiguration) As DrawingMetadataInput
        If job Is Nothing Then Return New DrawingMetadataInput()
        Dim fechaVal As DateTime = DateTime.Today
        If Not String.IsNullOrWhiteSpace(job.FechaPlano) Then
            Dim d As DateTime
            If DateTime.TryParse(job.FechaPlano, CultureInfo.InvariantCulture, DateTimeStyles.None, d) Then
                fechaVal = d.Date
            ElseIf DateTime.TryParse(job.FechaPlano, d) Then
                fechaVal = d.Date
            End If
        End If
        Dim nombreArchivo As String = If(String.IsNullOrWhiteSpace(job.PartListNombreArchivo), "", job.PartListNombreArchivo.Trim())
        If nombreArchivo = "" AndAlso Not String.IsNullOrWhiteSpace(job.InputFile) Then
            nombreArchivo = Path.GetFileName(job.InputFile)
        End If
        Dim numPartList As String = job.DrawingNumber
        If String.IsNullOrWhiteSpace(numPartList) Then numPartList = "1"
        Return New DrawingMetadataInput With {
            .Cliente = job.ClientName,
            .Proyecto = job.ProjectName,
            .Pedido = job.Pedido,
            .Plano = job.DrawingNumber,
            .Titulo = job.DrawingTitle,
            .Revision = job.Revision,
            .Autor = job.AuthorName,
            .Fecha = fechaVal,
            .Numero = numPartList.Trim(),
            .Cantidad = If(String.IsNullOrWhiteSpace(job.PartListCantidad), "1", job.PartListCantidad.Trim()),
            .NombreArchivo = nombreArchivo,
            .Denominacion = job.DrawingTitle,
            .Material = job.Material,
            .Espesor = job.Thickness,
            .LargoL = job.PartListL,
            .AltoH = job.PartListH,
            .DatoD = job.PartListD,
            .Peso = job.Weight
        }
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
            SolidEdgePropertyService.GetDocumentPropertyForMetadata(doc, "ProjectInformation", {"Document Number", "Número de documento", "Part Number", "Número de pieza"}),
            SolidEdgePropertyService.GetDocumentPropertyForMetadata(doc, "Custom", {"Plano", "plan", "CODIGO", "Numero de plano", "Número de plano"}))
        data.PlanoSource = If(String.IsNullOrWhiteSpace(data.Plano), "vacío", "modelo")
        UiMetadataFieldTrace(logger, traceUiLog, traceKind, "Plano", data.Plano)

        ' En español Summary «Título» / COM Title suele guardar denominación (no confundir con número de plano).
        data.Titulo = FirstNonEmpty(
            SolidEdgePropertyService.GetDocumentPropertyForMetadata(doc, "SummaryInformation", {"Subject", "Title", "Document Title", "Título", "Titulo"}),
            SolidEdgePropertyService.GetDocumentPropertyForMetadata(doc, "Custom", {"Titulo", "Título"}))
        data.TituloSource = If(String.IsNullOrWhiteSpace(data.Titulo), "vacío", "modelo")
        UiMetadataFieldTrace(logger, traceUiLog, traceKind, "Titulo", data.Titulo)

        If fillPartList Then
            data.Denominacion = FirstNonEmpty(
                SolidEdgePropertyService.GetDocumentPropertyForMetadata(doc, "Custom", {"Denominacion", "Denominación"}),
                If(String.IsNullOrWhiteSpace(data.Titulo), "", data.Titulo))
            data.DenominacionSource = If(String.IsNullOrWhiteSpace(data.Denominacion), "vacío", "modelo")
            UiMetadataFieldTrace(logger, traceUiLog, traceKind, "Denominacion", data.Denominacion)
        End If

        data.Revision = FirstNonEmpty(
            SolidEdgePropertyService.GetDocumentPropertyForMetadata(doc, "SummaryInformation", {"Número de revisión", "Revision Number", "Revision", "Revisión"}),
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
            ' Propiedades personalizadas L/H/D/Material… en la pieza o DFT (prioridad sobre caja 3D / cotas).
            TryFillPartListFieldsFromDocumentCustomIfEmpty(doc, data, logger, "modelo")

            TryDetectMaterial(doc, logger, data)
            UiMetadataFieldTrace(logger, traceUiLog, traceKind, "Material", data.Material)
            TryDetectEspesor(doc, doc, logger, data)
            If String.IsNullOrWhiteSpace(data.Espesor) Then data.EspesorSource = "Missing"
            UiMetadataFieldTrace(logger, traceUiLog, traceKind, "Espesor", data.Espesor)
            TryDetectPeso(doc, logger, data)
            If String.IsNullOrWhiteSpace(data.Peso) Then
                data.PesoSource = "Missing"
                UiMetadataFieldEmpty(logger, traceUiLog, traceKind, "Peso")
            Else
                UiMetadataFieldTrace(logger, traceUiLog, traceKind, "Peso", data.Peso)
            End If

            ' L/H/D con la misma prioridad que en carga de metadatos (DFT: cotas/vistas; PAR/PSM: caja 3D).
            Try
                Dim safePath = If(inputPath, "").Trim()
                Dim extOpen = If(safePath.Length > 0, Path.GetExtension(safePath).ToLowerInvariant(), "")
                Dim lGeom As String = "", hGeom As String = "", dGeom As String = "", lhdSrcGeom As String = ""
                If extOpen = ".dft" Then
                    If TryComputeLhdSameAsCalcButton(doc, Nothing, logger, lGeom, hGeom, dGeom, lhdSrcGeom) Then
                        If String.IsNullOrWhiteSpace(data.LargoL) Then data.LargoL = lGeom
                        If String.IsNullOrWhiteSpace(data.AltoH) Then data.AltoH = hGeom
                        If String.IsNullOrWhiteSpace(data.DatoD) Then data.DatoD = dGeom
                        If Not String.IsNullOrWhiteSpace(lhdSrcGeom) Then data.LHDSource = lhdSrcGeom
                    End If
                ElseIf extOpen = ".par" OrElse extOpen = ".psm" Then
                    If TryComputeLhdSameAsCalcButton(Nothing, doc, logger, lGeom, hGeom, dGeom, lhdSrcGeom) Then
                        If String.IsNullOrWhiteSpace(data.LargoL) Then data.LargoL = lGeom
                        If String.IsNullOrWhiteSpace(data.AltoH) Then data.AltoH = hGeom
                        If String.IsNullOrWhiteSpace(data.DatoD) Then data.DatoD = dGeom
                        If Not String.IsNullOrWhiteSpace(lhdSrcGeom) Then data.LHDSource = lhdSrcGeom
                    End If
                End If
            Catch ex As Exception
                If logger IsNot Nothing Then logger.Log("[PARTLISTDATA][LHD][LOAD_TITLEBLOCK][ERR] " & ex.Message)
            End Try

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

        ' Parte del cajetín enlaza contra Project / Summary cuando no hay entrada en Personalizado (nombres se recortan al leer desde COM).
        For Each ck In customKeys
            Dim vProj = SolidEdgePropertyService.GetDocumentPropertyForMetadata(doc, "ProjectInformation", {ck})
            If Not String.IsNullOrWhiteSpace(vProj) Then
                data.Pedido = vProj.Trim()
                data.PedidoSource = "ProjectInformation." & ck
                If logger IsNot Nothing Then logger.Log("[METADATA][PEDIDO][FROM] " & data.PedidoSource & " = " & data.Pedido)
                Return
            End If
            Dim vSum = SolidEdgePropertyService.GetDocumentPropertyForMetadata(doc, "SummaryInformation", {ck})
            If Not String.IsNullOrWhiteSpace(vSum) Then
                data.Pedido = vSum.Trim()
                data.PedidoSource = "SummaryInformation." & ck
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

        Dim looseNeedles As String() = {"pedido", "Purchase Order", "OrdenCompra", "orden de compra", "número de pedido"}
        For Each nd In looseNeedles
            Dim pv = SolidEdgePropertyService.TryFindDocumentPropertyValueByLoosePropertyName(doc, nd)
            If Not String.IsNullOrWhiteSpace(pv) Then
                data.Pedido = pv.Trim()
                data.PedidoSource = "propiedad (nombre contiene «" & nd & "»)"
                If logger IsNot Nothing Then logger.Log("[METADATA][PEDIDO][FROM_LOOSE] " & data.PedidoSource & " = " & data.Pedido)
                Return
            End If
        Next

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

    ''' <summary>Misma prioridad que el cálculo L/H/D al cargar metadatos: DFT (<see cref="TryComputeLhdFromDraft"/>), si falla el modelo (<see cref="TryComputeLhdFromModelDoc"/>).</summary>
    Public Shared Function TryComputeLhdSameAsCalcButton(draftDoc As Object, modelDoc As Object, logger As Logger,
                                                         ByRef L As String, ByRef H As String, ByRef D As String,
                                                         ByRef lhdSourceLabel As String) As Boolean
        L = ""
        H = ""
        D = ""
        lhdSourceLabel = ""
        If draftDoc IsNot Nothing Then
            If TryComputeLhdFromDraft(draftDoc, logger, L, H, D) Then
                lhdSourceLabel = "Calculado (DFT)"
                Return True
            End If
        End If
        If modelDoc IsNot Nothing Then
            If TryComputeLhdFromModelDoc(modelDoc, logger, L, H, D) Then
                lhdSourceLabel = "Caja 3D (modelo)"
                Return True
            End If
        End If
        Return False
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

    ''' <summary>Rellena huecos del snapshot PART_LIST desde <c>Custom.*</c> del documento (pieza o DFT); no pisa valores ya informados por job/UI.</summary>
    Private Shared Sub TryFillPartListFieldsFromDocumentCustomIfEmpty(doc As Object, data As DrawingMetadataInput, logger As Logger, sourceShort As String)
        If doc Is Nothing OrElse data Is Nothing Then Return
        Dim tagBase As String = If(String.IsNullOrWhiteSpace(sourceShort), "doc", sourceShort.Trim())

        Dim v As String
        If String.IsNullOrWhiteSpace(data.LargoL) Then
            v = SolidEdgePropertyService.GetDocumentPropertyForMetadata(doc, "Custom", {"L"})
            If Not String.IsNullOrWhiteSpace(v) Then
                data.LargoL = v.Trim()
                data.LHDSource = tagBase & " (Custom.L)"
                If logger IsNot Nothing Then logger.Log("[PARTLISTDATA][ENRICH][" & tagBase & "] Custom.L=" & data.LargoL)
            End If
        End If
        If String.IsNullOrWhiteSpace(data.AltoH) Then
            v = SolidEdgePropertyService.GetDocumentPropertyForMetadata(doc, "Custom", {"H"})
            If Not String.IsNullOrWhiteSpace(v) Then
                data.AltoH = v.Trim()
                data.LHDSource = tagBase & " (Custom.L/H/D)"
                If logger IsNot Nothing Then logger.Log("[PARTLISTDATA][ENRICH][" & tagBase & "] Custom.H=" & data.AltoH)
            End If
        End If
        If String.IsNullOrWhiteSpace(data.DatoD) Then
            v = SolidEdgePropertyService.GetDocumentPropertyForMetadata(doc, "Custom", {"D"})
            If Not String.IsNullOrWhiteSpace(v) Then
                data.DatoD = v.Trim()
                data.LHDSource = tagBase & " (Custom.L/H/D)"
                If logger IsNot Nothing Then logger.Log("[PARTLISTDATA][ENRICH][" & tagBase & "] Custom.D=" & data.DatoD)
            End If
        End If

        If String.IsNullOrWhiteSpace(data.Material) Then
            v = SolidEdgePropertyService.GetDocumentPropertyForMetadata(doc, "Custom", {"Material", "MATERIAL", "MAT", "MATERIA"})
            If Not String.IsNullOrWhiteSpace(v) Then
                data.Material = v.Trim()
                data.MaterialSource = tagBase & " (Custom.Material)"
                If logger IsNot Nothing Then logger.Log("[PARTLISTDATA][ENRICH][" & tagBase & "] Custom.Material=" & data.Material)
            End If
        End If

        If String.IsNullOrWhiteSpace(data.Espesor) OrElse String.Equals(data.EspesorSource, "Missing", StringComparison.OrdinalIgnoreCase) Then
            v = SolidEdgePropertyService.GetDocumentPropertyForMetadata(doc, "Custom", {"Espesor", "ESPESOR", "Thickness", "THICKNESS", "Thick", "Calib"})
            If Not String.IsNullOrWhiteSpace(v) Then
                data.Espesor = NormalizeThicknessDisplay(v)
                data.EspesorSource = tagBase & " (Custom.Espesor)"
                If logger IsNot Nothing Then logger.Log("[PARTLISTDATA][ENRICH][" & tagBase & "] Custom.Espesor=" & data.Espesor)
            End If
        End If

        If String.IsNullOrWhiteSpace(data.Peso) Then
            v = SolidEdgePropertyService.GetDocumentPropertyForMetadata(doc, "Custom", {"Peso", "Weight", "WEIGHT", "Mass"})
            If Not String.IsNullOrWhiteSpace(v) Then
                data.Peso = FormatWeightForUi(v)
                data.PesoSource = tagBase & " (Custom.Peso)"
                If logger IsNot Nothing Then logger.Log("[PARTLISTDATA][ENRICH][" & tagBase & "] Custom.Peso=" & data.Peso)
            End If
        End If

        If String.IsNullOrWhiteSpace(data.Denominacion) Then
            v = SolidEdgePropertyService.GetDocumentPropertyForMetadata(doc, "Custom", {"Denominacion", "Denominación"})
            If Not String.IsNullOrWhiteSpace(v) Then
                data.Denominacion = v.Trim()
                data.DenominacionSource = tagBase & " (Custom.Denominacion)"
            End If
        End If

        If String.IsNullOrWhiteSpace(data.Plano) Then
            v = SolidEdgePropertyService.GetDocumentPropertyForMetadata(doc, "Custom", {"CODIGO", "Codigo", "Código"})
            If Not String.IsNullOrWhiteSpace(v) Then
                data.Plano = v.Trim()
                data.PlanoSource = tagBase & " (Custom.CODIGO)"
            End If
        End If
    End Sub

    ''' <summary>Antes de escribir Custom en pieza/DFT: metadatos de la pieza (Custom + detectores) y refuerzo desde DFT; L/H/D geométricos si siguen vacíos.</summary>
    Private Shared Sub TryEnrichPartListSnapshotFromLinkedDocuments(modelDoc As Object, draftDoc As Object, data As DrawingMetadataInput, logger As Logger)
        If data Is Nothing Then Return

        If modelDoc IsNot Nothing Then
            TryFillPartListFieldsFromDocumentCustomIfEmpty(modelDoc, data, logger, "pieza")
        End If
        If draftDoc IsNot Nothing Then
            TryFillPartListFieldsFromDocumentCustomIfEmpty(draftDoc, data, logger, "DFT")
        End If

        If String.IsNullOrWhiteSpace(data.Material) Then
            If modelDoc IsNot Nothing Then TryDetectMaterial(modelDoc, logger, data)
            If String.IsNullOrWhiteSpace(data.Material) AndAlso draftDoc IsNot Nothing Then TryDetectMaterial(draftDoc, logger, data)
        End If

        If String.IsNullOrWhiteSpace(data.Espesor) OrElse String.Equals(data.EspesorSource, "Missing", StringComparison.OrdinalIgnoreCase) Then
            TryDetectEspesor(modelDoc, draftDoc, logger, data)
        End If

        If String.IsNullOrWhiteSpace(data.Peso) Then
            If modelDoc IsNot Nothing Then TryDetectPeso(modelDoc, logger, data)
            If String.IsNullOrWhiteSpace(data.Peso) AndAlso draftDoc IsNot Nothing Then TryDetectPeso(draftDoc, logger, data)
        End If

        If String.IsNullOrWhiteSpace(data.LargoL) OrElse String.IsNullOrWhiteSpace(data.AltoH) OrElse String.IsNullOrWhiteSpace(data.DatoD) Then
            Dim lV As String = "", hV As String = "", dV As String = "", lhdSrc As String = ""
            If TryComputeLhdSameAsCalcButton(draftDoc, modelDoc, logger, lV, hV, dV, lhdSrc) Then
                If String.IsNullOrWhiteSpace(data.LargoL) AndAlso Not String.IsNullOrWhiteSpace(lV) Then data.LargoL = lV
                If String.IsNullOrWhiteSpace(data.AltoH) AndAlso Not String.IsNullOrWhiteSpace(hV) Then data.AltoH = hV
                If String.IsNullOrWhiteSpace(data.DatoD) AndAlso Not String.IsNullOrWhiteSpace(dV) Then data.DatoD = dV
                If Not String.IsNullOrWhiteSpace(lhdSrc) AndAlso String.IsNullOrWhiteSpace(data.LHDSource) Then data.LHDSource = lhdSrc
                If logger IsNot Nothing Then
                    logger.Log("[PARTLISTDATA][ENRICH][LHD_GEOM] " & lhdSrc & " → L=" & If(data.LargoL, "") & " H=" & If(data.AltoH, "") & " D=" & If(data.DatoD, ""))
                End If
            End If
        End If
    End Sub

    ''' <summary>
    ''' Sesión COM única: SummaryInfo del cajetín + Custom/PART_LIST (mismos datos que la UI de metadatos).
    ''' Escribe en log <c>[PARTLISTDATA][SNAPSHOT]</c> y <c>[PARTSLIST][COL_DIAG]</c> para alinear nombres con la plantilla.
    ''' </summary>
    Public Shared Function ApplyCajetinSummaryAndPartListToFiles(
        dftPath As String,
        modelPath As String,
        showSolidEdge As Boolean,
        config As JobConfiguration,
        metadata As DrawingMetadataInput,
        logger As Logger) As ApplyMetadataToDraftFilesResult

        Dim result As New ApplyMetadataToDraftFilesResult()
        If String.IsNullOrWhiteSpace(dftPath) OrElse Not File.Exists(dftPath) Then
            If logger IsNot Nothing Then logger.Log("[UI][DFT][ERR] DFT no encontrado: " & If(dftPath, ""))
            Return result
        End If
        If metadata Is Nothing Then metadata = New DrawingMetadataInput()
        If config Is Nothing Then config = New JobConfiguration()

        Dim app As Application = Nothing
        Dim created As Boolean = False
        Dim dftDoc As Object = Nothing
        Dim modelDoc As Object = Nothing
        Dim openedModel As Boolean = False
        Dim openedDft As Boolean = False

        Try
            OleMessageFilter.Register()
            If Not TryConnectForMetadata(showSolidEdge, logger, app, created) Then Return result

            If logger IsNot Nothing Then
                logger.Log("[UI][DFT][APPLY] Abriendo DFT: " & dftPath)
                If Not String.IsNullOrWhiteSpace(modelPath) AndAlso File.Exists(modelPath) Then
                    logger.Log("[UI][DFT][APPLY] Modelo enlazado: " & modelPath)
                End If
            End If

            dftDoc = SolidEdgePropertyService.TryGetOrOpenDraftDocumentByPath(app, dftPath, logger, openedDft)
            If dftDoc IsNot Nothing Then
                dftDoc = SolidEdgePropertyService.TryEnsureDraftWithPartsLists(dftDoc, logger)
            End If
            If logger IsNot Nothing AndAlso Not openedDft Then
                logger.Log("[UI][DFT][APPLY] DFT ya estaba abierto en Solid Edge; se reutiliza esa ventana (no una copia COM).")
            End If

            If String.IsNullOrWhiteSpace(modelPath) OrElse Not File.Exists(modelPath) Then
                Dim linkOnly As String = SolidEdgePropertyService.TryGetPrimaryLinkedModelFullPath(dftDoc)
                If Not String.IsNullOrWhiteSpace(linkOnly) AndAlso File.Exists(linkOnly) Then
                    modelPath = linkOnly
                    If logger IsNot Nothing Then logger.Log("[UI][DFT][APPLY] Modelo desde enlace DFT: " & modelPath)
                End If
            End If

            If Not String.IsNullOrWhiteSpace(modelPath) AndAlso File.Exists(modelPath) Then
                Dim ext As String = Path.GetExtension(modelPath).ToLowerInvariant()
                If ext = ".par" OrElse ext = ".psm" Then
                    Try
                        modelDoc = SolidEdgePropertyService.TryGetOrOpenModelDocumentByPath(app, modelPath, logger, openedModel)
                    Catch exMo As Exception
                        If logger IsNot Nothing Then logger.Log("[UI][DFT][APPLY][WARN] No se abrió modelo: " & exMo.Message)
                    End Try
                End If
            End If

            NormalizeMetadataPlaceholders(metadata)
            If PartsListPropertyTextWriter.IsMetadataPlaceholder(metadata.LargoL) OrElse
               PartsListPropertyTextWriter.IsMetadataPlaceholder(metadata.AltoH) OrElse
               PartsListPropertyTextWriter.IsMetadataPlaceholder(metadata.DatoD) Then
                Dim lV As String = "", hV As String = "", dV As String = "", lhdSrc As String = ""
                If TryComputeLhdSameAsCalcButton(dftDoc, modelDoc, logger, lV, hV, dV, lhdSrc) Then
                    If String.IsNullOrWhiteSpace(metadata.LargoL) Then metadata.LargoL = lV
                    If String.IsNullOrWhiteSpace(metadata.AltoH) Then metadata.AltoH = hV
                    If String.IsNullOrWhiteSpace(metadata.DatoD) Then metadata.DatoD = dV
                    If Not String.IsNullOrWhiteSpace(lhdSrc) AndAlso String.IsNullOrWhiteSpace(metadata.LHDSource) Then metadata.LHDSource = lhdSrc
                    If logger IsNot Nothing Then
                        logger.Log("[UI][DFT][APPLY][LHD] " & lhdSrc & " → L=" & lV & " H=" & hV & " D=" & dV)
                    End If
                End If
            End If

            result.SummaryInfoOk = SolidEdgePropertyService.ApplyDirectSummaryInfoToDraft(dftDoc, config, logger)
            If logger IsNot Nothing Then logger.Log("[UI][DFT][APPLY] SummaryInfo=" & result.SummaryInfoOk.ToString(CultureInfo.InvariantCulture))

            LogPartListSnapshotBeforeApply(metadata, logger)
            ApplyPartListSourceProperties(modelDoc, dftDoc, metadata, logger)
            result.PartListPipelineRan = True

            If modelDoc IsNot Nothing Then
                Try
                    CallByName(modelDoc, "Save", CallType.Method)
                    result.ModelSaved = True
                    If logger IsNot Nothing Then
                        logger.Log("[UI][DFT][APPLY] Modelo guardado." & If(Not openedModel, " (documento ya abierto en SE)", ""))
                    End If
                Catch exMs As Exception
                    If logger IsNot Nothing Then logger.Log("[UI][DFT][APPLY][WARN] Save modelo: " & exMs.Message)
                End Try
            End If

            Try
                CallByName(dftDoc, "Save", CallType.Method)
                result.DftSaved = True
                If logger IsNot Nothing Then
                    logger.Log("[UI][DFT][APPLY] DFT guardado." & If(Not openedDft, " (ventana que tenía abierta)", ""))
                End If
            Catch exDs As Exception
                If logger IsNot Nothing Then logger.Log("[UI][DFT][APPLY][WARN] Save DFT: " & exDs.Message)
            End Try

            If logger IsNot Nothing Then
                logger.Log("[UI][DFT][APPLY] Fin SummaryInfo=" & result.SummaryInfoOk.ToString(CultureInfo.InvariantCulture) &
                           " PartList=" & result.PartListPipelineRan.ToString(CultureInfo.InvariantCulture) &
                           " Revise [PARTSLIST][COL_DIAG] para ver PropertyText de cada columna.")
            End If
        Catch ex As Exception
            If logger IsNot Nothing Then logger.LogException("ApplyCajetinSummaryAndPartListToFiles", ex)
        Finally
            ' No cerrar documentos que el usuario ya tenía abiertos en Solid Edge.
            If openedModel Then SolidEdgePropertyService.TryCloseComDocument(modelDoc, False)
            If openedDft Then SolidEdgePropertyService.TryCloseComDocument(dftDoc, False)
            Try
                If app IsNot Nothing AndAlso created Then app.Quit()
            Catch
            End Try
            Try : OleMessageFilter.Revoke() : Catch : End Try
        End Try

        Return result
    End Function

    Private Shared Sub LogPartListSnapshotBeforeApply(data As DrawingMetadataInput, logger As Logger)
        If logger Is Nothing OrElse data Is Nothing Then Return
        logger.Log("[PARTLISTDATA][SNAPSHOT] --- Valores UI → Custom.* (pieza + DFT) y fila PART_LIST ---")
        For Each line As String In FormatMetadataSnapshotLines(data).Split(New String() {vbCrLf, vbLf}, StringSplitOptions.RemoveEmptyEntries)
            logger.Log("[PARTLISTDATA][SNAPSHOT] " & line.Trim())
        Next
        logger.Log("[PARTLISTDATA][SNAPSHOT] Claves Custom escritas: CODIGO, Material, Espesor, L, H, D, Peso, Denominacion, NombreArchivo, Nombre, Cantidad, Numero")
        logger.Log("[PARTLISTDATA][SNAPSHOT] Tras aplicar: busque [PARTSLIST][COL_DIAG] (header + PropertyText); deben coincidir nombres con Custom.")
    End Sub

    Public Shared Sub ApplyPartListSourceProperties(modelDoc As Object, draftDoc As Object, data As DrawingMetadataInput, logger As Logger)
        If data Is Nothing Then Return
        If logger IsNot Nothing Then logger.Log("[PARTLISTDATA][WRITE] Inicio ApplyPartListSourceProperties")

        NormalizeMetadataPlaceholders(data)
        If draftDoc IsNot Nothing Then
            draftDoc = SolidEdgePropertyService.TryEnsureDraftWithPartsLists(draftDoc, logger)
        End If
        TryEnrichPartListSnapshotFromLinkedDocuments(modelDoc, draftDoc, data, logger)

        Dim openedLinkedModel As Boolean = False
        If modelDoc Is Nothing AndAlso draftDoc IsNot Nothing Then
            modelDoc = PartsListPropertyTextWriter.TryOpenLinkedPartDocument(draftDoc, logger)
            openedLinkedModel = (modelDoc IsNot Nothing)
        End If

        Dim propTextWrites As Integer = 0
        If modelDoc IsNot Nothing AndAlso draftDoc IsNot Nothing Then
            If logger IsNot Nothing Then logger.Log("[PARTLISTDATA][WRITE] Ruta PropertyText → documento de pieza (columnas %{…|G} / %{…/CP|G}).")
            SolidEdgePropertyService.TryActivateDraftDocument(draftDoc, logger)
            propTextWrites = PartsListPropertyTextWriter.ApplyAllPartsListsOnDraft(modelDoc, draftDoc, data, logger)
            If logger IsNot Nothing Then
                logger.Log("[PARTLISTDATA][PROP_TEXT_WRITE] total=" & propTextWrites.ToString(CultureInfo.InvariantCulture))
            End If
        ElseIf logger IsNot Nothing Then
            logger.Log("[PARTLISTDATA][WRITE][WARN] Sin modelo .par/.psm: PropertyText no se puede propagar; solo Custom en DFT.")
        End If

        Dim pairs As New List(Of KeyValuePair(Of String, String))()
        SubAdd(pairs, "CODIGO", data.Plano)
        SubAdd(pairs, "Material", data.Material)
        SubAdd(pairs, "Espesor", data.Espesor)
        SubAdd(pairs, "L", data.LargoL)
        SubAdd(pairs, "H", data.AltoH)
        SubAdd(pairs, "D", data.DatoD)
        SubAdd(pairs, "Peso", data.Peso)
        SubAdd(pairs, "Denominacion", If(String.IsNullOrWhiteSpace(data.Denominacion), data.Titulo, data.Denominacion))
        SubAdd(pairs, "NombreArchivo", data.NombreArchivo)
        ' Columna tipo «Nombre» en plantillas enlazadas a Custom distinto del nombre de archivo.
        SubAdd(pairs, "Nombre", data.NombreArchivo)
        SubAdd(pairs, "Cantidad", data.Cantidad)
        SubAdd(pairs, "Numero", data.Numero)

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

        If modelDoc IsNot Nothing Then
            Dim mmN = SolidEdgePropertyService.TryMirrorPartListOntoMechanicalModel(modelDoc, data.Material, data.Espesor, data.Peso, logger, mirrorMaterialToMechanical:=False)
            If logger IsNot Nothing AndAlso mmN > 0 Then
                logger.Log("[PARTLISTDATA][WRITE_MODEL_MM] escritas=" & mmN.ToString(CultureInfo.InvariantCulture))
            End If
            Try
                CallByName(modelDoc, "Save", CallType.Method)
            Catch ex As Exception
                If logger IsNot Nothing Then logger.Log("[PARTLISTDATA][WRITE_MODEL_MM][SAVE][WARN] " & ex.Message)
            End Try
        End If

        If openedLinkedModel AndAlso modelDoc IsNot Nothing Then
            Try
                CallByName(modelDoc, "Save", CallType.Method)
            Catch ex As Exception
                If logger IsNot Nothing Then logger.Log("[PARTLISTDATA][MODEL_SAVE][WARN] " & ex.Message)
            End Try
        End If

        If draftDoc IsNot Nothing Then
            SolidEdgePropertyService.RefreshDraftFromModelLinks(draftDoc, logger, False)
            SolidEdgePropertyService.RefreshNativePartsListsAndUpdateAll(draftDoc, logger)
            SolidEdgePropertyService.RefreshDraftPropertyTextOnly(draftDoc, logger)
            Dim needConv As Boolean = False
            TryPushNativePartsListCellsFromMetadata(draftDoc, data, logger, needConv, accumulateNeedConv:=False)
            If Not needConv Then
                needConv = Not TryVerifyPartsListCriticalFieldsVisible(draftDoc, data, logger)
            End If
            If needConv Then
                If logger IsNot Nothing Then
                    logger.Log("[PARTLISTDATA][CONVERT_TABLE] Enlaces/override no muestran Espesor o L/H/D; convirtiendo a tabla y escribiendo celdas.")
                End If
                TryPushPartListFieldsThroughConvertedTable(draftDoc, data, logger)
            End If
            If SolidEdgePropertyService.TryGetDraftPartsListCount(draftDoc) <= 0 Then
                TryPushPartListFieldsToExistingDraftTables(draftDoc, data, logger)
            End If
        End If
        If logger IsNot Nothing Then logger.Log("[PARTLISTDATA][WRITE] Fin")
    End Sub

    ''' <summary>Convierte «--» y similares en vacío para no escribirlos en Comentarios/L/H/D de la pieza.</summary>
    Friend Shared Sub NormalizeMetadataPlaceholders(data As DrawingMetadataInput)
        If data Is Nothing Then Return
        If PartsListPropertyTextWriter.IsMetadataPlaceholder(data.Espesor) Then data.Espesor = ""
        If PartsListPropertyTextWriter.IsMetadataPlaceholder(data.AltoH) Then data.AltoH = ""
        If PartsListPropertyTextWriter.IsMetadataPlaceholder(data.LargoL) Then data.LargoL = ""
        If PartsListPropertyTextWriter.IsMetadataPlaceholder(data.DatoD) Then data.DatoD = ""
        If PartsListPropertyTextWriter.IsMetadataPlaceholder(data.Material) Then data.Material = ""
        If PartsListPropertyTextWriter.IsMetadataPlaceholder(data.Peso) Then data.Peso = ""
    End Sub

    Private Shared Sub SubAdd(list As List(Of KeyValuePair(Of String, String)), k As String, v As String)
        If list Is Nothing OrElse PartsListPropertyTextWriter.IsMetadataPlaceholder(v) Then Return
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
            If logger IsNot Nothing Then logger.Log("[UI][METADATA][STEP] Conectando a Solid Edge...")
            If Not TryConnectForMetadata(showSe, logger, app, created) Then Return False
            If logger IsNot Nothing Then logger.Log("[UI][METADATA][STEP] Conexión lista (nueva_instancia=" & created.ToString() & "), abriendo: " & modelPath)
            doc = app.Documents.Open(modelPath)
            If logger IsNot Nothing Then logger.Log("[UI][METADATA][STEP] Documents.Open terminado.")
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
            SolidEdgePropertyService.TryCloseComDocument(doc, False)
            Try
                If app IsNot Nothing AndAlso created Then app.Quit()
            Catch
            End Try
            Try : OleMessageFilter.Revoke() : Catch : End Try
        End Try
    End Function

    ''' <summary>Si hay pieza enlazada al DFT, rellena Material/Espesor/Peso/LHD y corrige nombre de archivo cuando el DFT no aporta pieza.</summary>
    Private Shared Sub TryEnrichDraftMetadataFromLinkedModel(app As Application, draftDoc As Object, dftPath As String, data As DrawingMetadataInput, logger As Logger,
                                                             Optional traceKind As String = "DFT_FIELD")
        If app Is Nothing OrElse draftDoc Is Nothing OrElse data Is Nothing Then Return
        Dim rawLink As String = SolidEdgePropertyService.TryGetPrimaryLinkedModelFullPath(draftDoc)
        If String.IsNullOrWhiteSpace(rawLink) Then
            If logger IsNot Nothing Then logger.Log("[UI][METADATA][DFT][LINK] Sin vínculo a modelo (.par/.psm/.asm) en el plano.")
            Return
        End If

        Dim candidates As New List(Of String)()
        Dim primary As String = rawLink.Trim()
        candidates.Add(primary)
        Try
            Dim bn As String = Path.GetFileName(primary)
            If bn.Length > 0 AndAlso Not String.IsNullOrWhiteSpace(dftPath) Then
                Dim dn As String = Path.GetDirectoryName(dftPath)
                If Not String.IsNullOrWhiteSpace(dn) Then
                    Dim alt As String = Path.Combine(dn, bn)
                    If Not candidates.Contains(alt, StringComparer.OrdinalIgnoreCase) Then candidates.Add(alt)
                End If
            End If
        Catch
        End Try

        Dim mdoc As Object = Nothing
        Dim linkPath As String = ""
        For Each cand As String In candidates
            If String.IsNullOrWhiteSpace(cand) Then Continue For
            Try
                mdoc = app.Documents.Open(cand.Trim())
                linkPath = cand.Trim()
                Exit For
            Catch ex As Exception
                If logger IsNot Nothing Then logger.Log("[UI][METADATA][DFT][LINK][OPEN_TRY] " & cand & " -> " & ex.Message)
            End Try
        Next

        If mdoc Is Nothing Then
            If logger IsNot Nothing Then logger.Log("[UI][METADATA][DFT][LINK] No se pudo abrir el modelo enlazado (intentos=" & candidates.Count.ToString(CultureInfo.InvariantCulture) & ").")
            Return
        End If
        If logger IsNot Nothing Then logger.Log("[UI][METADATA][DFT][LINK] Enriqueciendo desde: " & linkPath)

        Try
            Dim ext = Path.GetExtension(linkPath).ToLowerInvariant()
            Dim isPart As Boolean = ext = ".par" OrElse ext = ".psm"

            Dim needsNameFix As Boolean = String.IsNullOrWhiteSpace(data.NombreArchivo) OrElse
                Path.GetExtension(data.NombreArchivo).Equals(".dft", StringComparison.OrdinalIgnoreCase)
            If needsNameFix AndAlso isPart Then
                data.NombreArchivo = Path.GetFileName(linkPath)
                data.NombreArchivoSource = "modelo enlazado"
                UiMetadataFieldTrace(logger, True, traceKind, "NombreArchivo", data.NombreArchivo)
            End If

            If String.IsNullOrWhiteSpace(data.Material) Then
                TryDetectMaterial(mdoc, logger, data)
                If Not String.IsNullOrWhiteSpace(data.Material) Then data.MaterialSource = "modelo enlazado"
                UiMetadataFieldTrace(logger, True, traceKind, "Material", data.Material)
            End If

            If String.IsNullOrWhiteSpace(data.Espesor) OrElse String.Equals(data.EspesorSource, "Missing", StringComparison.OrdinalIgnoreCase) Then
                TryDetectEspesor(mdoc, draftDoc, logger, data)
                UiMetadataFieldTrace(logger, True, traceKind, "Espesor", data.Espesor)
            End If

            If String.IsNullOrWhiteSpace(data.Peso) Then
                TryDetectPeso(mdoc, logger, data)
                If String.IsNullOrWhiteSpace(data.Peso) Then
                    UiMetadataFieldEmpty(logger, True, traceKind, "Peso")
                Else
                    UiMetadataFieldTrace(logger, True, traceKind, "Peso", data.Peso)
                End If
            End If

            If String.IsNullOrWhiteSpace(data.Titulo) AndAlso isPart Then
                Dim tAlt = FirstNonEmpty(
                    SolidEdgePropertyService.GetDocumentPropertyForMetadata(mdoc, "SummaryInformation", {"Subject", "Title", "Document Title", "Título", "Titulo"}),
                    SolidEdgePropertyService.GetDocumentPropertyForMetadata(mdoc, "Custom", {"Titulo", "Título"}))
                If Not String.IsNullOrWhiteSpace(tAlt) Then
                    data.Titulo = tAlt.Trim()
                    data.TituloSource = "modelo enlazado"
                    UiMetadataFieldTrace(logger, True, traceKind, "Titulo", data.Titulo)
                    If String.IsNullOrWhiteSpace(data.Denominacion) Then
                        data.Denominacion = data.Titulo
                        data.DenominacionSource = "modelo enlazado"
                        UiMetadataFieldTrace(logger, True, traceKind, "Denominacion", data.Denominacion)
                    End If
                End If
            End If

            If isPart Then
                Dim L As String = "", Hh As String = "", D As String = "", lhdSrc As String = ""
                If TryComputeLhdSameAsCalcButton(draftDoc, mdoc, logger, L, Hh, D, lhdSrc) Then
                    If String.IsNullOrWhiteSpace(data.LargoL) Then data.LargoL = L
                    If String.IsNullOrWhiteSpace(data.AltoH) Then data.AltoH = Hh
                    If String.IsNullOrWhiteSpace(data.DatoD) Then data.DatoD = D
                    If Not (String.IsNullOrWhiteSpace(data.LargoL) AndAlso String.IsNullOrWhiteSpace(data.AltoH) AndAlso String.IsNullOrWhiteSpace(data.DatoD)) Then
                        If Not String.IsNullOrWhiteSpace(lhdSrc) Then data.LHDSource = lhdSrc
                        UiMetadataFieldTrace(logger, True, traceKind, "L", data.LargoL)
                        UiMetadataFieldTrace(logger, True, traceKind, "H", data.AltoH)
                        UiMetadataFieldTrace(logger, True, traceKind, "D", data.DatoD)
                    End If
                End If
            End If
        Catch ex As Exception
            If logger IsNot Nothing Then logger.LogException("TryEnrichDraftMetadataFromLinkedModel", ex)
        Finally
            SolidEdgePropertyService.TryCloseComDocument(mdoc, False)
        End Try
    End Sub

    ''' <summary>Lee cajetín y campos de PART_LIST desde un <c>.dft</c> (mismos PropertySets que <see cref="LoadTitleBlockFromOpenModel"/> + claves <c>Custom.*</c> que escribe <see cref="ApplyPartListSourceProperties"/>).</summary>
    Public Shared Function TryLoadMetadataFromDraftPath(dftPath As String, showSe As Boolean, logger As Logger, ByRef data As DrawingMetadataInput) As Boolean
        data = Nothing
        If String.IsNullOrWhiteSpace(dftPath) OrElse Not File.Exists(dftPath) Then Return False
        If Not String.Equals(Path.GetExtension(dftPath), ".dft", StringComparison.OrdinalIgnoreCase) Then Return False

        Dim app As Application = Nothing
        Dim created As Boolean = False
        Dim doc As Object = Nothing
        Try
            OleMessageFilter.Register()
            If logger IsNot Nothing Then logger.Log("[UI][METADATA][DFT][STEP] Conectando y abriendo: " & dftPath)
            If Not TryConnectForMetadata(showSe, logger, app, created) Then Return False
            doc = app.Documents.Open(dftPath)
            data = New DrawingMetadataInput()
            LoadTitleBlockFromOpenModel(doc, dftPath, logger, data, fillPartList:=True, inferPedidoFromPath:=False, traceUiLog:=True, traceKind:="DFT_FIELD")
            OverlayDraftPartListCustomKeys(doc, data, logger)
            OverlayDraftNativePartsListFirstDataRow(doc, data, logger)
            TryEnrichDraftMetadataFromLinkedModel(app, doc, dftPath, data, logger)
            RetagMetadataSourcesDraftFile(data)
            If String.IsNullOrWhiteSpace(data.Pedido) Then
                data.Pedido = InferPedidoFromPath(dftPath, logger)
                If Not String.IsNullOrWhiteSpace(data.Pedido) Then
                    data.PedidoSource = "Inferido (carpeta o nombre de archivo)"
                End If
            End If
            Return True
        Catch ex As Exception
            If logger IsNot Nothing Then logger.LogException("TryLoadMetadataFromDraftPath", ex)
            data = Nothing
            Return False
        Finally
            SolidEdgePropertyService.TryCloseComDocument(doc, False)
            Try
                If app IsNot Nothing AndAlso created Then app.Quit()
            Catch
            End Try
            Try : OleMessageFilter.Revoke() : Catch : End Try
        End Try
    End Function

    Private Shared Function PartsListHeaderLooksLikeMaterial(header As String) As Boolean
        Dim u As String = If(header, "").Trim().ToUpperInvariant()
        Return u.Length > 0 AndAlso (u.Contains("MATERIAL") OrElse u.Contains("MATERIA"))
    End Function

    Private Shared Function PartsListHeaderLooksLikeEspesor(header As String) As Boolean
        Dim u As String = If(header, "").Trim().ToUpperInvariant()
        If u.Length = 0 Then Return False
        Return u.Contains("ESPESOR") OrElse u.Contains("THICK") OrElse u.Contains("CALIBRE") OrElse u.Contains("GAUGE")
    End Function

    Private Shared Function PartsListHeaderLooksLikeCodigo(header As String) As Boolean
        Dim h As String = If(header, "").Trim()
        If h.Length = 0 Then Return False
        If h.StartsWith("Nº", StringComparison.Ordinal) OrElse h.StartsWith("N°", StringComparison.Ordinal) Then Return False
        If h.IndexOf("codig", StringComparison.OrdinalIgnoreCase) >= 0 Then Return True
        Return String.Equals(h, "CODE", StringComparison.OrdinalIgnoreCase)
    End Function

    Private Shared Function PartsListHeaderLooksLikeDenominacion(header As String) As Boolean
        Dim u As String = If(header, "").Trim().ToUpperInvariant()
        If u.Length = 0 Then Return False
        Return u.Contains("DENOMINAC") OrElse u.Contains("DESIGNATION") OrElse u.Contains("DESIGNACI") OrElse u.StartsWith("DESCRIP")
    End Function

    Private Shared Function PartsListHeaderLooksLikePeso(header As String) As Boolean
        Dim u As String = If(header, "").Trim().ToUpperInvariant()
        If u.Length = 0 Then Return False
        Return u.StartsWith("PESO", StringComparison.Ordinal) OrElse u.Contains("WEIGHT") OrElse u.StartsWith("MASS", StringComparison.Ordinal)
    End Function

    ''' <summary>Después de propiedades y Update, escribe valores en la fila de datos de <b>cada</b> PartsList (plantillas con enlaces que no propagan Custom/MM).</summary>
    ''' <param name="needsConvertToTableForLhd">True si en alguna lista hay columna Material/Espesor/L/H/D con dato en metadatos pero la escritura en PartsList falló (entonces <c>ConvertToTable</c>).</param>
    Private Shared Sub TryPushNativePartsListCellsFromMetadata(draftDoc As Object, data As DrawingMetadataInput, logger As Logger,
                                                                ByRef needsConvertToTableForLhd As Boolean,
                                                                Optional accumulateNeedConv As Boolean = False)
        If Not accumulateNeedConv Then needsConvertToTableForLhd = False
        If draftDoc Is Nothing OrElse data Is Nothing Then Return
        Dim lists As Object = Nothing
        Try
            lists = CallByName(draftDoc, "PartsLists", CallType.Get)
        Catch
            lists = Nothing
        End Try
        If lists Is Nothing Then Return
        Dim nLists As Integer = SafeInt(CallByName(lists, "Count", CallType.Get), 0)
        If nLists < 1 Then Return

        For li As Integer = 1 To nLists
            Dim pl As Object = GetItem(lists, li)
            If pl Is Nothing Then Continue For
            TryPushNativePartsListCellsOneList(draftDoc, pl, li, data, logger, needsConvertToTableForLhd)
        Next
    End Sub

    Private Shared Function TryVerifyPartsListCriticalFieldsVisible(draftDoc As Object, data As DrawingMetadataInput, logger As Logger) As Boolean
        If draftDoc Is Nothing OrElse data Is Nothing Then Return True
        Dim lists As Object = Nothing
        Try : lists = CallByName(draftDoc, "PartsLists", CallType.Get) : Catch : End Try
        If lists Is Nothing Then Return True
        Dim nLists As Integer = SafeInt(CallByName(lists, "Count", CallType.Get), 0)
        If nLists < 1 Then Return True
        Dim pl As Object = GetItem(lists, 1)
        If pl Is Nothing Then Return True
        Dim rowCount As Integer = SolidEdgePropertyService.TryGetPartsListRowCount(pl)
        Dim hdrRows As Integer = SolidEdgePropertyService.TryGetPartsListHeaderRowCount(pl)
        Dim dataRow As Integer = SolidEdgePropertyService.ResolvePartsListDataRowIndex(hdrRows, rowCount)
        Dim colCount As Integer = SolidEdgePropertyService.TryGetPartsListColumnCount(pl)
        Dim iEsp As Integer = 0, iL As Integer = 0, iH As Integer = 0, iD As Integer = 0
        For c As Integer = 1 To colCount
            Dim h As String = SolidEdgePropertyService.TryReadPartsListColumnHeader(pl, c)
            If iEsp = 0 AndAlso PartsListHeaderLooksLikeEspesor(h) Then iEsp = c
            If iL = 0 AndAlso PartsListHeaderMatchesDimLetter(h, "L"c) Then iL = c
            If iH = 0 AndAlso PartsListHeaderMatchesDimLetter(h, "H"c) Then iH = c
            If iD = 0 AndAlso PartsListHeaderMatchesDimLetter(h, "D"c) Then iD = c
        Next
        Dim missing As New List(Of String)()
        If iEsp > 0 AndAlso Not PartsListPropertyTextWriter.IsMetadataPlaceholder(data.Espesor) AndAlso
            Not SolidEdgePropertyService.TryPartsListCellDisplaysValue(pl, dataRow, iEsp, data.Espesor) Then missing.Add("Espesor")
        If iL > 0 AndAlso Not PartsListPropertyTextWriter.IsMetadataPlaceholder(data.LargoL) AndAlso
            Not SolidEdgePropertyService.TryPartsListCellDisplaysValue(pl, dataRow, iL, data.LargoL) Then missing.Add("L")
        If iH > 0 AndAlso Not PartsListPropertyTextWriter.IsMetadataPlaceholder(data.AltoH) AndAlso
            Not SolidEdgePropertyService.TryPartsListCellDisplaysValue(pl, dataRow, iH, data.AltoH) Then missing.Add("H")
        If iD > 0 AndAlso Not PartsListPropertyTextWriter.IsMetadataPlaceholder(data.DatoD) AndAlso
            Not SolidEdgePropertyService.TryPartsListCellDisplaysValue(pl, dataRow, iD, data.DatoD) Then missing.Add("D")
        If missing.Count = 0 Then Return True
        If logger IsNot Nothing Then
            logger.Log("[PARTLISTDATA][VERIFY][FAIL] fila=" & dataRow.ToString(CultureInfo.InvariantCulture) &
                       " sin valor visible en tabla: " & String.Join(", ", missing))
        End If
        Return False
    End Function

    Private Shared Sub TryPushNativePartsListCellsOneList(draftDoc As Object, pl As Object, listIdx As Integer, data As DrawingMetadataInput, logger As Logger,
                                                          ByRef needsConvertToTableForLhd As Boolean)
        If pl Is Nothing OrElse data Is Nothing Then Return
        Try
            CallByName(pl, "Active", CallType.Set, True)
        Catch
        End Try
        SolidEdgePropertyService.TryActivatePartsListParentSheet(pl, draftDoc, logger)
        Try
            CallByName(pl, "Update", CallType.Method)
        Catch ex As Exception
            If logger IsNot Nothing Then logger.Log("[PARTSLIST][UPDATE][PRE_WRITE] " & ex.Message)
        End Try

        Dim colCount As Integer = SolidEdgePropertyService.TryGetPartsListColumnCount(pl)
        Dim rowCount As Integer = SolidEdgePropertyService.TryGetPartsListRowCount(pl)
        If colCount <= 0 OrElse rowCount <= 0 Then Return

        If logger IsNot Nothing Then
            SolidEdgePropertyService.TryLogPartsListColumnDefinitions(pl, listIdx, logger)
        End If

        Dim hdrRows As Integer = SolidEdgePropertyService.TryGetPartsListHeaderRowCount(pl)
        If hdrRows <= 0 Then hdrRows = 1
        Dim dataRow As Integer = SolidEdgePropertyService.ResolvePartsListDataRowIndex(hdrRows, rowCount)
        Dim dataRowCandidates As Integer() = SolidEdgePropertyService.ResolvePartsListDataRowCandidates(hdrRows, rowCount)
        If logger IsNot Nothing Then
            logger.Log("[PARTLISTDATA][CELLS] list=" & listIdx.ToString(CultureInfo.InvariantCulture) &
                       " rows=" & rowCount.ToString(CultureInfo.InvariantCulture) &
                       " hdrRows=" & hdrRows.ToString(CultureInfo.InvariantCulture) &
                       " dataRow=" & dataRow.ToString(CultureInfo.InvariantCulture) &
                       " candidates=" & String.Join(",", dataRowCandidates.Select(Function(r) r.ToString(CultureInfo.InvariantCulture))))
        End If
        If dataRow < 1 Then Return

        Dim iCod As Integer = 0, iDen As Integer = 0, iMat As Integer = 0, iEsp As Integer = 0
        Dim iL As Integer = 0, iH As Integer = 0, iD As Integer = 0, iPeso As Integer = 0
        For c As Integer = 1 To colCount
            Dim h As String = SolidEdgePropertyService.TryReadPartsListColumnHeader(pl, c)
            If iCod = 0 AndAlso PartsListHeaderLooksLikeCodigo(h) Then iCod = c
            If iDen = 0 AndAlso PartsListHeaderLooksLikeDenominacion(h) Then iDen = c
            If iMat = 0 AndAlso PartsListHeaderLooksLikeMaterial(h) Then iMat = c
            If iEsp = 0 AndAlso PartsListHeaderLooksLikeEspesor(h) Then iEsp = c
            If iL = 0 AndAlso PartsListHeaderMatchesDimLetter(h, "L"c) Then iL = c
            If iH = 0 AndAlso PartsListHeaderMatchesDimLetter(h, "H"c) Then iH = c
            If iD = 0 AndAlso PartsListHeaderMatchesDimLetter(h, "D"c) Then iD = c
            If iPeso = 0 AndAlso PartsListHeaderLooksLikePeso(h) Then iPeso = c
        Next

        Dim filled As New List(Of String)()
        Dim okMat As Boolean = False, okEsp As Boolean = False
        Dim okL As Boolean = False, okH As Boolean = False, okD As Boolean = False

        Dim pushCols As New List(Of Tuple(Of Integer, String, String))()
        SubAddPush(pushCols, iCod, If(data.Plano, "").Trim(), "CODIGO")
        Dim denom As String = If(String.IsNullOrWhiteSpace(data.Denominacion), If(data.Titulo, "").Trim(), data.Denominacion.Trim())
        SubAddPush(pushCols, iDen, denom, "DENOMINACION")
        SubAddPush(pushCols, iMat, MetadataValueOrEmpty(data.Material), "Material")
        SubAddPush(pushCols, iEsp, MetadataValueOrEmpty(data.Espesor), "Espesor")
        SubAddPush(pushCols, iL, MetadataValueOrEmpty(data.LargoL), "L")
        SubAddPush(pushCols, iH, MetadataValueOrEmpty(data.AltoH), "H")
        SubAddPush(pushCols, iD, MetadataValueOrEmpty(data.DatoD), "D")
        SubAddPush(pushCols, iPeso, If(data.Peso, "").Trim(), "Peso")

        Dim writeRowUsed As Integer = dataRow
        For Each writeRow In dataRowCandidates
            writeRowUsed = writeRow
            For Each t In pushCols
                If PartsListPropertyTextWriter.ColumnUsesLinkedPropertyText(pl, t.Item1) Then Continue For
                If SolidEdgePropertyService.TryWritePartsListCellText(pl, writeRow, t.Item1, t.Item2, logger) Then
                    If Not filled.Contains(t.Item3) Then filled.Add(t.Item3)
                End If
            Next

            For Each t In pushCols
                If Not PartsListPropertyTextWriter.ColumnUsesLinkedPropertyText(pl, t.Item1) Then Continue For
                If String.IsNullOrWhiteSpace(t.Item2) Then Continue For
                If SolidEdgePropertyService.TryWriteTableCellDataTabOverride(pl, writeRow, t.Item1, t.Item2, logger, "[PARTSLIST][DATA_TAB]") Then
                    If Not filled.Contains(t.Item3) Then filled.Add(t.Item3)
                ElseIf writeRow = dataRow AndAlso logger IsNot Nothing Then
                    logger.Log("[PARTLISTDATA][DATA_TAB][WARN] col=" & t.Item1.ToString(CultureInfo.InvariantCulture) &
                               " " & t.Item3 & " (PropertyText en modelo OK; sobrescritura de celda falló).")
                End If
            Next
        Next

        okMat = iMat <= 0 OrElse PartsListPropertyTextWriter.IsMetadataPlaceholder(data.Material) OrElse
            SolidEdgePropertyService.TryPartsListCellDisplaysValue(pl, writeRowUsed, iMat, data.Material)
        okEsp = iEsp <= 0 OrElse PartsListPropertyTextWriter.IsMetadataPlaceholder(data.Espesor) OrElse
            SolidEdgePropertyService.TryPartsListCellDisplaysValue(pl, writeRowUsed, iEsp, data.Espesor)
        okL = iL <= 0 OrElse PartsListPropertyTextWriter.IsMetadataPlaceholder(data.LargoL) OrElse
            SolidEdgePropertyService.TryPartsListCellDisplaysValue(pl, writeRowUsed, iL, data.LargoL)
        okH = iH <= 0 OrElse PartsListPropertyTextWriter.IsMetadataPlaceholder(data.AltoH) OrElse
            SolidEdgePropertyService.TryPartsListCellDisplaysValue(pl, writeRowUsed, iH, data.AltoH)
        okD = iD <= 0 OrElse PartsListPropertyTextWriter.IsMetadataPlaceholder(data.DatoD) OrElse
            SolidEdgePropertyService.TryPartsListCellDisplaysValue(pl, writeRowUsed, iD, data.DatoD)

        If iMat > 0 AndAlso Not PartsListPropertyTextWriter.IsMetadataPlaceholder(data.Material) AndAlso Not okMat Then needsConvertToTableForLhd = True
        If iEsp > 0 AndAlso Not PartsListPropertyTextWriter.IsMetadataPlaceholder(data.Espesor) AndAlso Not okEsp Then needsConvertToTableForLhd = True
        If iL > 0 AndAlso Not PartsListPropertyTextWriter.IsMetadataPlaceholder(data.LargoL) AndAlso Not okL Then needsConvertToTableForLhd = True
        If iH > 0 AndAlso Not PartsListPropertyTextWriter.IsMetadataPlaceholder(data.AltoH) AndAlso Not okH Then needsConvertToTableForLhd = True
        If iD > 0 AndAlso Not PartsListPropertyTextWriter.IsMetadataPlaceholder(data.DatoD) AndAlso Not okD Then needsConvertToTableForLhd = True

        If logger IsNot Nothing AndAlso filled.Count > 0 Then
            logger.Log("[PARTLISTDATA][CELLS_PUSH] list=" & listIdx.ToString(CultureInfo.InvariantCulture) &
                       " fila=" & writeRowUsed.ToString(CultureInfo.InvariantCulture) & " " & String.Join(", ", filled) &
                       " visibleEsp=" & okEsp.ToString(CultureInfo.InvariantCulture) &
                       " visibleLHD=" & okL.ToString(CultureInfo.InvariantCulture) & okH.ToString(CultureInfo.InvariantCulture) & okD.ToString(CultureInfo.InvariantCulture))
        End If

        Dim readbackCols As New List(Of Integer)()
        For Each c In New Integer() {iCod, iDen, iMat, iEsp, iL, iH, iD, iPeso}
            If c > 0 Then readbackCols.Add(c)
        Next
        SolidEdgePropertyService.TryLogPartsListCellReadback(pl, listIdx, readbackCols, logger)
    End Sub

    ''' <summary>Si PartsList nativa no acepta escritura (enlaces), <c>ConvertToTable</c> y rellena Material, Espesor y L/H/D en la <c>Table</c>.</summary>
    Private Shared Sub TryPushPartListFieldsThroughConvertedTable(draftDoc As Object, data As DrawingMetadataInput, logger As Logger)
        If draftDoc Is Nothing OrElse data Is Nothing Then Return
        If String.IsNullOrWhiteSpace(data.Material) AndAlso String.IsNullOrWhiteSpace(data.Espesor) AndAlso
            String.IsNullOrWhiteSpace(data.LargoL) AndAlso String.IsNullOrWhiteSpace(data.AltoH) AndAlso String.IsNullOrWhiteSpace(data.DatoD) Then Return

        SolidEdgePropertyService.TryActivateWorkingDrawingSheet(draftDoc, logger)
        Dim tbl As Object = SolidEdgePropertyService.TryConvertFirstPartsListToTable(draftDoc, logger)
        If tbl Is Nothing Then Return
        SolidEdgePropertyService.TryActivatePartsListParentSheet(tbl, draftDoc, logger)
        TryPushPartListFieldsToDraftTable(tbl, data, logger, "ConvertToTable")
    End Sub

    ''' <summary>Cuando <c>PartsLists.Count=0</c>, busca tablas en <c>Tables</c>/<c>DraftTables</c> (PART_LIST como tabla de usuario) y escribe celdas.</summary>
    Friend Shared Sub TryPushPartListFieldsToExistingDraftTables(draftDoc As Object, data As DrawingMetadataInput, logger As Logger)
        If draftDoc Is Nothing OrElse data Is Nothing Then Return
        If String.IsNullOrWhiteSpace(data.Material) AndAlso String.IsNullOrWhiteSpace(data.Espesor) AndAlso
            String.IsNullOrWhiteSpace(data.LargoL) AndAlso String.IsNullOrWhiteSpace(data.AltoH) AndAlso String.IsNullOrWhiteSpace(data.DatoD) Then Return

        SolidEdgePropertyService.TryActivateWorkingDrawingSheet(draftDoc, logger)
        Dim workingSheet As Object = SolidEdgePropertyService.TryGetResolvedWorkingDrawingSheet(draftDoc, logger)
        Dim workingName As String = ""
        If workingSheet IsNot Nothing Then
            Try
                workingName = Convert.ToString(CallByName(workingSheet, "Name", CallType.Get)).Trim()
            Catch
                workingName = ""
            End Try
        End If

        SolidEdgePropertyService.LogDraftTableInventory(draftDoc, logger)

        Dim bestTbl As Object = Nothing
        Dim bestScore As Integer = 0
        For Each tbl As Object In SolidEdgePropertyService.TryEnumerateAllDraftTables(draftDoc)
            Dim score As Integer = SolidEdgePropertyService.TryScoreDraftTableAsPartList(tbl)
            If score <= 0 Then Continue For
            Dim sheetName As String = SolidEdgePropertyService.TryGetDraftTableSheetName(tbl)
            If workingName.Length > 0 AndAlso sheetName.Length > 0 AndAlso
                String.Equals(sheetName, workingName, StringComparison.OrdinalIgnoreCase) Then
                score += 50
            End If
            If score > bestScore Then
                bestScore = score
                bestTbl = tbl
            End If
        Next

        If bestTbl Is Nothing OrElse bestScore < 3 Then
            If logger IsNot Nothing Then
                logger.Log("[PARTLISTDATA][TABLE_FALLBACK][WARN] Sin tabla PART_LIST en Tables/DraftTables (score máx=" &
                           bestScore.ToString(CultureInfo.InvariantCulture) &
                           "). Los datos están en el .par; en SE use Actualizar lista de piezas o inserte PartsList nativa.")
            End If
            Return
        End If

        If logger IsNot Nothing Then
            logger.Log("[PARTLISTDATA][TABLE_FALLBACK] Tabla candidata sheet='" &
                       SolidEdgePropertyService.TryGetDraftTableSheetName(bestTbl).Replace("'", "") &
                       "' score=" & bestScore.ToString(CultureInfo.InvariantCulture) &
                       " (pestaña «Datos» de Propiedades de tabla)")
            SolidEdgePropertyService.TryLogDraftTableColumnDefinitions(bestTbl, logger)
        End If
        TryPushPartListFieldsToDraftTable(bestTbl, data, logger, "TABLE_FALLBACK")
        SolidEdgePropertyService.TryRefreshDraftTable(bestTbl, logger)
        Dim pl As PartsList = TryCast(bestTbl, PartsList)
        If pl IsNot Nothing Then
            Try
                CallByName(pl, "Update", CallType.Method)
                If logger IsNot Nothing Then logger.Log("[PARTSLIST][UPDATE][OK] TABLE_FALLBACK PartsList")
            Catch ex As Exception
                If logger IsNot Nothing Then logger.Log("[PARTSLIST][UPDATE][WARN] TABLE_FALLBACK " & ex.Message)
            End Try
        End If
    End Sub

    Private Shared Sub TryPushPartListFieldsToDraftTable(tbl As Object, data As DrawingMetadataInput, logger As Logger, logTag As String)
        If tbl Is Nothing OrElse data Is Nothing Then Return

        Dim colCount As Integer = SolidEdgePropertyService.TryGetDraftTableColumnCount(tbl)
        Dim rowCount As Integer = SolidEdgePropertyService.TryGetDraftTableRowCount(tbl)
        If colCount <= 0 OrElse rowCount <= 0 Then
            If logger IsNot Nothing Then logger.Log("[PARTLISTDATA][" & logTag & "][WARN] tabla sin filas/columnas.")
            Return
        End If

        Dim hdrRows As Integer = SolidEdgePropertyService.TryGetDraftTableHeaderRowCount(tbl)
        If hdrRows <= 0 Then hdrRows = 1
        Dim dataRow As Integer = SolidEdgePropertyService.ResolvePartsListDataRowIndex(hdrRows, rowCount)
        If dataRow < 1 Then Return

        Dim iMat As Integer = 0, iEsp As Integer = 0
        Dim iL As Integer = 0, iH As Integer = 0, iD As Integer = 0
        For c As Integer = 1 To colCount
            Dim h As String = SolidEdgePropertyService.TryReadDraftTableColumnHeader(tbl, c)
            If iMat = 0 AndAlso PartsListHeaderLooksLikeMaterial(h) Then iMat = c
            If iEsp = 0 AndAlso PartsListHeaderLooksLikeEspesor(h) Then iEsp = c
            If iL = 0 AndAlso PartsListHeaderMatchesDimLetter(h, "L"c) Then iL = c
            If iH = 0 AndAlso PartsListHeaderMatchesDimLetter(h, "H"c) Then iH = c
            If iD = 0 AndAlso PartsListHeaderMatchesDimLetter(h, "D"c) Then iD = c
        Next

        Dim iDen As Integer = 0, iPeso As Integer = 0
        For c As Integer = 1 To colCount
            Dim h As String = SolidEdgePropertyService.TryReadDraftTableColumnHeader(tbl, c)
            If iDen = 0 AndAlso PartsListHeaderLooksLikeDenominacion(h) Then iDen = c
            If iPeso = 0 AndAlso PartsListHeaderLooksLikePeso(h) Then iPeso = c
        Next
        Dim denom As String = If(String.IsNullOrWhiteSpace(data.Denominacion), If(data.Titulo, "").Trim(), data.Denominacion.Trim())

        Dim filled As New List(Of String)()
        If iMat > 0 AndAlso SolidEdgePropertyService.TryWriteDraftTableCellValue(tbl, dataRow, iMat, MetadataValueOrEmpty(data.Material), logger) Then filled.Add("Material")
        If iEsp > 0 AndAlso SolidEdgePropertyService.TryWriteDraftTableCellValue(tbl, dataRow, iEsp, MetadataValueOrEmpty(data.Espesor), logger) Then filled.Add("Espesor")
        If iL > 0 AndAlso SolidEdgePropertyService.TryWriteDraftTableCellValue(tbl, dataRow, iL, MetadataValueOrEmpty(data.LargoL), logger) Then filled.Add("L")
        If iH > 0 AndAlso SolidEdgePropertyService.TryWriteDraftTableCellValue(tbl, dataRow, iH, MetadataValueOrEmpty(data.AltoH), logger) Then filled.Add("H")
        If iD > 0 AndAlso SolidEdgePropertyService.TryWriteDraftTableCellValue(tbl, dataRow, iD, MetadataValueOrEmpty(data.DatoD), logger) Then filled.Add("D")
        If iDen > 0 AndAlso Not String.IsNullOrWhiteSpace(denom) AndAlso
            SolidEdgePropertyService.TryWriteDraftTableCellValue(tbl, dataRow, iDen, denom, logger) Then filled.Add("DENOMINACION")
        If iPeso > 0 AndAlso Not String.IsNullOrWhiteSpace(data.Peso) AndAlso
            SolidEdgePropertyService.TryWriteDraftTableCellValue(tbl, dataRow, iPeso, data.Peso.Trim(), logger) Then filled.Add("Peso")

        If logger IsNot Nothing Then
            If filled.Count > 0 Then
                logger.Log("[PARTLISTDATA][" & logTag & "] fila=" & dataRow.ToString(CultureInfo.InvariantCulture) & " " & String.Join(", ", filled))
            Else
                logger.Log("[PARTLISTDATA][" & logTag & "][WARN] Sin columnas coincidentes o escritura sin efecto (celdas enlazadas).")
            End If
        End If
    End Sub

    Private Shared Function MetadataValueOrEmpty(value As String) As String
        If PartsListPropertyTextWriter.IsMetadataPlaceholder(value) Then Return ""
        Return If(value, "").Trim()
    End Function

    Private Shared Sub SubAddPush(list As List(Of Tuple(Of Integer, String, String)), colIdx As Integer, val As String, label As String)
        If list Is Nothing OrElse colIdx <= 0 OrElse String.IsNullOrWhiteSpace(val) Then Return
        list.Add(New Tuple(Of Integer, String, String)(colIdx, val.Trim(), label))
    End Sub

    ''' <summary>Reconoce cabeceras «L», «L (mm)», «Largo»… sin confundir con «Documento» (D + letra).</summary>
    Private Shared Function PartsListHeaderMatchesDimLetter(header As String, letter As Char) As Boolean
        Dim t As String = If(header, "").Trim()
        If t.Length = 0 Then Return False
        Dim c0 As Char = Char.ToUpperInvariant(t(0))
        If c0 = Char.ToUpperInvariant(letter) AndAlso (t.Length = 1 OrElse Not Char.IsLetter(t(1))) Then Return True
        Dim u As String = t.ToUpperInvariant()
        If letter = "L"c AndAlso (u.StartsWith("LARGO") OrElse u.StartsWith("LONG") OrElse u.StartsWith("LONGITUD")) Then Return True
        If letter = "H"c AndAlso (u.StartsWith("ALTO") OrElse u.StartsWith("ALTURA") OrElse u.StartsWith("HEIGHT")) Then Return True
        If letter = "D"c AndAlso (u.StartsWith("PROF") OrElse u.StartsWith("FONDO") OrElse u.StartsWith("DEPTH")) Then Return True
        Return False
    End Function

    ''' <summary>Primera fila de datos nativa «PartsList» cuando Custom no trae MATERIAL/ESP/L/H/D.</summary>
    Private Shared Sub OverlayDraftNativePartsListFirstDataRow(draftDoc As Object, data As DrawingMetadataInput, logger As Logger)
        If draftDoc Is Nothing OrElse data Is Nothing Then Return
        Dim lists As Object = Nothing
        Try
            lists = CallByName(draftDoc, "PartsLists", CallType.Get)
        Catch
            lists = Nothing
        End Try
        If lists Is Nothing Then Return
        Dim nLists As Integer = 0
        Try
            nLists = CInt(CallByName(lists, "Count", CallType.Get))
        Catch
            nLists = 0
        End Try
        If nLists < 1 Then Return
        Dim pl As Object = Nothing
        Try
            pl = CallByName(lists, "Item", CallType.Get, 1)
        Catch
            pl = Nothing
        End Try
        If pl Is Nothing Then
            Try
                pl = CallByName(lists, "Item", CallType.Method, 1)
            Catch
                pl = Nothing
            End Try
        End If
        If pl Is Nothing Then Return
        Try
            CallByName(pl, "Update", CallType.Method)
        Catch ex As Exception
            If logger IsNot Nothing Then logger.Log("[PARTSLIST][UPDATE][DFT_LOAD] " & ex.Message)
        End Try

        Dim colCount As Integer = SolidEdgePropertyService.TryGetPartsListColumnCount(pl)
        Dim rowCount As Integer = SolidEdgePropertyService.TryGetPartsListRowCount(pl)
        If colCount <= 0 OrElse rowCount <= 0 Then Return

        Dim hdrRows As Integer = SolidEdgePropertyService.TryGetPartsListHeaderRowCount(pl)
        If hdrRows <= 0 Then hdrRows = 1
        Dim dataRow As Integer = SolidEdgePropertyService.ResolvePartsListDataRowIndex(hdrRows, rowCount)
        If dataRow < 1 Then Return

        Dim iMat As Integer = 0, iEsp As Integer = 0, iL As Integer = 0, iH As Integer = 0, iD As Integer = 0
        For c As Integer = 1 To colCount
            Dim h As String = SolidEdgePropertyService.TryReadPartsListColumnHeader(pl, c)
            If iMat = 0 AndAlso PartsListHeaderLooksLikeMaterial(h) Then iMat = c
            If iEsp = 0 AndAlso PartsListHeaderLooksLikeEspesor(h) Then iEsp = c
            If iL = 0 AndAlso PartsListHeaderMatchesDimLetter(h, "L"c) Then iL = c
            If iH = 0 AndAlso PartsListHeaderMatchesDimLetter(h, "H"c) Then iH = c
            If iD = 0 AndAlso PartsListHeaderMatchesDimLetter(h, "D"c) Then iD = c
        Next

        Dim filled As New List(Of String)()
        If iMat > 0 AndAlso String.IsNullOrWhiteSpace(data.Material) Then
            Dim cellV As String = SolidEdgePropertyService.TryReadPartsListCellText(pl, dataRow, iMat)
            If Not String.IsNullOrWhiteSpace(cellV) Then
                data.Material = cellV.Trim()
                data.MaterialSource = "DFT (PartsList)"
                filled.Add("Material")
            End If
        End If
        If iEsp > 0 AndAlso (String.IsNullOrWhiteSpace(data.Espesor) OrElse String.Equals(data.EspesorSource, "Missing", StringComparison.OrdinalIgnoreCase)) Then
            Dim cellV As String = SolidEdgePropertyService.TryReadPartsListCellText(pl, dataRow, iEsp)
            If Not String.IsNullOrWhiteSpace(cellV) Then
                data.Espesor = cellV.Trim()
                data.EspesorSource = "DFT (PartsList)"
                filled.Add("Espesor")
            End If
        End If
        If iL > 0 AndAlso String.IsNullOrWhiteSpace(data.LargoL) Then
            Dim cellV As String = SolidEdgePropertyService.TryReadPartsListCellText(pl, dataRow, iL)
            If Not String.IsNullOrWhiteSpace(cellV) Then
                data.LargoL = cellV.Trim()
                data.LHDSource = "DFT (PartsList)"
                filled.Add("L")
            End If
        End If
        If iH > 0 AndAlso String.IsNullOrWhiteSpace(data.AltoH) Then
            Dim cellV As String = SolidEdgePropertyService.TryReadPartsListCellText(pl, dataRow, iH)
            If Not String.IsNullOrWhiteSpace(cellV) Then
                data.AltoH = cellV.Trim()
                data.LHDSource = "DFT (PartsList)"
                filled.Add("H")
            End If
        End If
        If iD > 0 AndAlso String.IsNullOrWhiteSpace(data.DatoD) Then
            Dim cellV As String = SolidEdgePropertyService.TryReadPartsListCellText(pl, dataRow, iD)
            If Not String.IsNullOrWhiteSpace(cellV) Then
                data.DatoD = cellV.Trim()
                data.LHDSource = "DFT (PartsList)"
                filled.Add("D")
            End If
        End If

        If logger IsNot Nothing AndAlso filled.Count > 0 Then
            logger.Log("[UI][METADATA][DFT][PARTSLIST_ROW] " & String.Join(", ", filled))
        End If
    End Sub

    ''' <summary>Refuerza L/H/D, cantidad, nº pieza y nombre de archivo desde propiedades personalizadas del DFT (convención de esta aplicación).</summary>
    Private Shared Sub OverlayDraftPartListCustomKeys(doc As Object, data As DrawingMetadataInput, logger As Logger)
        If doc Is Nothing OrElse data Is Nothing Then Return

        Dim v As String = SolidEdgePropertyService.GetDocumentPropertyForMetadata(doc, "Custom", {"L"})
        If Not String.IsNullOrWhiteSpace(v) Then
            data.LargoL = v.Trim()
            data.LHDSource = "DFT (Custom.L)"
        End If
        v = SolidEdgePropertyService.GetDocumentPropertyForMetadata(doc, "Custom", {"H"})
        If Not String.IsNullOrWhiteSpace(v) Then
            data.AltoH = v.Trim()
            data.LHDSource = "DFT (Custom.L/H/D)"
        End If
        v = SolidEdgePropertyService.GetDocumentPropertyForMetadata(doc, "Custom", {"D"})
        If Not String.IsNullOrWhiteSpace(v) Then
            data.DatoD = v.Trim()
            data.LHDSource = "DFT (Custom.L/H/D)"
        End If

        If String.IsNullOrWhiteSpace(data.Material) Then
            v = SolidEdgePropertyService.GetDocumentPropertyForMetadata(doc, "Custom", {"Material", "MATERIAL", "MAT", "MATERIA"})
            If Not String.IsNullOrWhiteSpace(v) Then
                data.Material = v.Trim()
                data.MaterialSource = "DFT (Custom.Material)"
            End If
        End If
        If String.IsNullOrWhiteSpace(data.Espesor) OrElse String.Equals(data.EspesorSource, "Missing", StringComparison.OrdinalIgnoreCase) Then
            v = SolidEdgePropertyService.GetDocumentPropertyForMetadata(doc, "Custom", {"Espesor", "ESPESOR", "Thickness", "THICKNESS", "Thick", "Calib"})
            If Not String.IsNullOrWhiteSpace(v) Then
                data.Espesor = v.Trim()
                data.EspesorSource = "DFT (Custom.Espesor)"
            End If
        End If

        v = SolidEdgePropertyService.GetDocumentPropertyForMetadata(doc, "Custom", {"Cantidad"})
        If Not String.IsNullOrWhiteSpace(v) Then
            data.Cantidad = v.Trim()
            data.CantidadSource = "DFT (Custom.Cantidad)"
        End If

        v = SolidEdgePropertyService.GetDocumentPropertyForMetadata(doc, "Custom", {"CODIGO"})
        If Not String.IsNullOrWhiteSpace(v) AndAlso String.IsNullOrWhiteSpace(data.Plano) Then
            data.Plano = v.Trim()
            data.PlanoSource = "DFT (Custom.CODIGO)"
        End If

        v = SolidEdgePropertyService.GetDocumentPropertyForMetadata(doc, "Custom", {"NombreArchivo"})
        If Not String.IsNullOrWhiteSpace(v) Then
            data.NombreArchivo = v.Trim()
            data.NombreArchivoSource = "DFT (Custom.NombreArchivo)"
        End If
        If String.IsNullOrWhiteSpace(data.NombreArchivo) Then
            v = SolidEdgePropertyService.GetDocumentPropertyForMetadata(doc, "Custom", {"Nombre"})
            If Not String.IsNullOrWhiteSpace(v) Then
                data.NombreArchivo = v.Trim()
                data.NombreArchivoSource = "DFT (Custom.Nombre)"
            End If
        End If

        v = SolidEdgePropertyService.GetDocumentPropertyForMetadata(doc, "Custom", {"Numero", "Número", "NumeroPieza"})
        If String.IsNullOrWhiteSpace(v) Then
            v = SolidEdgePropertyService.GetDocumentPropertyForMetadata(doc, "ProjectInformation", {"Document Number", "Part Number", "Número de pieza"})
        End If
        If Not String.IsNullOrWhiteSpace(v) Then
            data.Numero = v.Trim()
            data.NumeroSource = "DFT"
        End If

        v = SolidEdgePropertyService.GetDocumentPropertyForMetadata(doc, "Custom", {"Denominacion", "Denominación"})
        If Not String.IsNullOrWhiteSpace(v) Then
            data.Denominacion = v.Trim()
            data.DenominacionSource = "DFT (Custom.Denominacion)"
        End If
        If logger IsNot Nothing Then logger.Log("[UI][METADATA][DFT][OVERLAY] Custom PART_LIST aplicado si existía en DFT.")
    End Sub

    Private Shared Sub RetagMetadataSourcesDraftFile(data As DrawingMetadataInput)
        If data Is Nothing Then Return
        If String.Equals(data.ClienteSource, "modelo", StringComparison.OrdinalIgnoreCase) Then data.ClienteSource = "DFT"
        If String.Equals(data.ProyectoSource, "modelo", StringComparison.OrdinalIgnoreCase) Then data.ProyectoSource = "DFT"
        If String.Equals(data.PlanoSource, "modelo", StringComparison.OrdinalIgnoreCase) Then data.PlanoSource = "DFT"
        If String.Equals(data.TituloSource, "modelo", StringComparison.OrdinalIgnoreCase) Then data.TituloSource = "DFT"
        If String.Equals(data.RevisionSource, "modelo", StringComparison.OrdinalIgnoreCase) Then data.RevisionSource = "DFT"
        If String.Equals(data.AutorSource, "modelo", StringComparison.OrdinalIgnoreCase) Then data.AutorSource = "DFT"
        If String.Equals(data.FechaSource, "modelo", StringComparison.OrdinalIgnoreCase) Then data.FechaSource = "DFT"
        If String.Equals(data.MaterialSource, "modelo", StringComparison.OrdinalIgnoreCase) Then data.MaterialSource = "DFT"
        If String.Equals(data.EspesorSource, "modelo", StringComparison.OrdinalIgnoreCase) Then data.EspesorSource = "DFT"
        If String.Equals(data.PesoSource, "modelo", StringComparison.OrdinalIgnoreCase) Then data.PesoSource = "DFT"
        If String.Equals(data.DenominacionSource, "modelo", StringComparison.OrdinalIgnoreCase) Then data.DenominacionSource = "DFT"
        If String.Equals(data.NombreArchivoSource, "modelo", StringComparison.OrdinalIgnoreCase) Then data.NombreArchivoSource = "DFT"
        If String.Equals(data.CantidadSource, "modelo", StringComparison.OrdinalIgnoreCase) Then data.CantidadSource = "DFT"
        If String.Equals(data.NumeroSource, "modelo", StringComparison.OrdinalIgnoreCase) Then data.NumeroSource = "DFT"
        If String.Equals(data.PedidoSource, "modelo", StringComparison.OrdinalIgnoreCase) Then data.PedidoSource = "DFT"
    End Sub

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
