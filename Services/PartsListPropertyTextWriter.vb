Option Strict Off

Imports System.Globalization
Imports System.IO
Imports System.Text.RegularExpressions
Imports SolidEdgeFramework

''' <summary>
''' Escribe metadatos en el documento de pieza según <c>TableColumn.PropertyText</c> de la PART_LIST
''' (p. ej. <c>%{Comentarios|G}</c>, <c>%{L/CP|G}</c>), no solo Custom.* genérico.
''' </summary>
Public NotInheritable Class PartsListPropertyTextWriter
    Private Shared ReadOnly PropertyTextRegex As New Regex("%\{([^}/|]+)(?:/CP)?\|G\}", RegexOptions.Compiled Or RegexOptions.CultureInvariant)

    Private Sub New()
    End Sub

    Public Shared Function TryOpenLinkedPartDocument(draftDoc As Object, logger As Logger) As Object
        If draftDoc Is Nothing Then Return Nothing
        Dim link As String = SolidEdgePropertyService.TryGetPrimaryLinkedModelFullPath(draftDoc)
        If String.IsNullOrWhiteSpace(link) OrElse Not File.Exists(link) Then
            If logger IsNot Nothing Then logger.Log("[PARTLISTDATA][MODEL_OPEN][SKIP] Sin .par/.psm enlazado accesible.")
            Return Nothing
        End If
        Dim ext As String = Path.GetExtension(link).ToLowerInvariant()
        If ext <> ".par" AndAlso ext <> ".psm" Then Return Nothing
        Try
            Dim appObj As Object = Nothing
            Try : appObj = CallByName(draftDoc, "Application", CallType.Get) : Catch : End Try
            Dim app As SolidEdgeFramework.Application = TryCast(appObj, SolidEdgeFramework.Application)
            If app Is Nothing Then Return Nothing
            Dim openedByUs As Boolean = False
            Return SolidEdgePropertyService.TryGetOrOpenModelDocumentByPath(app, link.Trim(), logger, openedByUs)
        Catch ex As Exception
            If logger IsNot Nothing Then logger.Log("[PARTLISTDATA][MODEL_OPEN][ERR] " & ex.Message)
            Return Nothing
        End Try
    End Function

    ''' <summary>Plantilla CADEBRO cuando el DFT no expone <c>PartsLists</c> por COM (p. ej. DFT ya abierto en SE).</summary>
    Friend Shared Function ApplyFallbackCadebroPropertyTextsToModel(modelDoc As Object, data As DrawingMetadataInput, logger As Logger) As Integer
        If modelDoc Is Nothing OrElse data Is Nothing Then Return 0
        Dim specs As Tuple(Of String, Boolean)() = {
            Tuple.Create("Número de artículo", False),
            Tuple.Create("Cantidad", False),
            Tuple.Create("Título", False),
            Tuple.Create("Comentarios", False),
            Tuple.Create("Espesor", False),
            Tuple.Create("L", True),
            Tuple.Create("H", True),
            Tuple.Create("D", True),
            Tuple.Create("Masa (Cantidad)", False)
        }
        Dim written As Integer = 0
        For Each spec In specs
            Dim val As String = MapMetadataToPropertyField(spec.Item1, data)
            If String.IsNullOrWhiteSpace(val) Then Continue For
            If TryWritePropertyOnDocument(modelDoc, spec.Item1, spec.Item2, val, logger) Then
                written += 1
                If logger IsNot Nothing Then
                    logger.Log("[PARTLISTDATA][PROP_TEXT_FALLBACK] field='" & spec.Item1 & "' value='" & val.Replace("'", "") & "'")
                End If
            End If
        Next
        Return written
    End Function

    ''' <returns>Número de columnas escritas correctamente en el modelo.</returns>
    Public Shared Function ApplyAllPartsListsOnDraft(modelDoc As Object, draftDoc As Object, data As DrawingMetadataInput, logger As Logger) As Integer
        If draftDoc Is Nothing OrElse modelDoc Is Nothing Then Return 0
        draftDoc = SolidEdgePropertyService.TryEnsureDraftWithPartsLists(draftDoc, logger)
        SolidEdgePropertyService.TryActivateWorkingDrawingSheet(draftDoc, logger)
        Dim lists As Object = SolidEdgePropertyService.TryGetDraftPartsLists(draftDoc, logger)
        Dim n As Integer = 0
        If lists IsNot Nothing Then
            Try : n = CInt(CallByName(lists, "Count", CallType.Get)) : Catch : End Try
        End If
        Dim written As Integer = 0
        For i As Integer = 1 To n
            Dim pl As Object = Nothing
            Try
                pl = CallByName(lists, "Item", CallType.Get, i)
            Catch
                Try : pl = CallByName(lists, "Item", CallType.Method, i) : Catch : End Try
            End Try
            If pl Is Nothing Then Continue For
            written += ApplyColumnPropertyTextsToModel(modelDoc, pl, data, logger)
        Next
        If written <= 0 Then
            If logger IsNot Nothing Then
                logger.Log("[PARTLISTDATA][PROP_TEXT_WRITE] Sin escritura por columnas; plantilla CADEBRO por defecto (PartsLists=" &
                           n.ToString(CultureInfo.InvariantCulture) & ").")
            End If
            written = ApplyFallbackCadebroPropertyTextsToModel(modelDoc, data, logger)
        End If
        Return written
    End Function

    Private Shared Function ApplyColumnPropertyTextsToModel(modelDoc As Object, partsList As Object, data As DrawingMetadataInput, logger As Logger) As Integer
        If modelDoc Is Nothing OrElse partsList Is Nothing OrElse data Is Nothing Then Return 0
        Dim colCount As Integer = SolidEdgePropertyService.TryGetPartsListColumnCount(partsList)
        If colCount <= 0 Then Return 0

        Dim written As Integer = 0
        For c As Integer = 1 To colCount
            Dim propText As String = TryReadColumnPropertyText(partsList, c)
            If String.IsNullOrWhiteSpace(propText) Then Continue For
            If Not propText.Contains("%{") Then Continue For

            Dim fieldName As String = ""
            Dim isCustom As Boolean = False
            If Not TryParsePropertyText(propText, fieldName, isCustom) Then Continue For

            Dim value As String = MapMetadataToPropertyField(fieldName, data)
            If String.IsNullOrWhiteSpace(value) Then Continue For

            Dim ok As Boolean = TryWritePropertyOnDocument(modelDoc, fieldName, isCustom, value, logger)
            If ok Then written += 1
            If logger IsNot Nothing Then
                Dim hdr As String = SolidEdgePropertyService.TryReadPartsListColumnHeader(partsList, c)
                logger.Log("[PARTLISTDATA][PROP_TEXT_WRITE] col=" & c.ToString(CultureInfo.InvariantCulture) &
                           " header='" & hdr.Replace("'", "") & "'" &
                           " field='" & fieldName & "'" &
                           " custom=" & isCustom.ToString(CultureInfo.InvariantCulture) &
                           " ok=" & ok.ToString(CultureInfo.InvariantCulture) &
                           " value='" & value.Replace("'", "") & "'")
            End If
        Next
        Return written
    End Function

    Public Shared Function ColumnUsesLinkedPropertyText(partsList As Object, col1 As Integer) As Boolean
        Dim pt As String = TryReadColumnPropertyText(partsList, col1)
        Return Not String.IsNullOrWhiteSpace(pt) AndAlso pt.IndexOf("%{", StringComparison.Ordinal) >= 0 AndAlso pt.IndexOf("|G", StringComparison.Ordinal) >= 0
    End Function

    Private Shared Function TryReadColumnPropertyText(partsList As Object, col1 As Integer) As String
        If partsList Is Nothing OrElse col1 <= 0 Then Return ""
        Try
            Dim cols As Object = Nothing
            Try : cols = CallByName(partsList, "Columns", CallType.Get) : Catch : End Try
            If cols Is Nothing Then Return ""
            Dim col As Object = Nothing
            Try
                col = CallByName(cols, "Item", CallType.Get, col1)
            Catch
                Try : col = CallByName(cols, "Item", CallType.Method, col1) : Catch : End Try
            End Try
            If col Is Nothing Then Return ""
            Return Convert.ToString(CallByName(col, "PropertyText", CallType.Get)).Trim()
        Catch
            Return ""
        End Try
    End Function

    Friend Shared Function TryParsePropertyText(propertyText As String, ByRef fieldName As String, ByRef isCustomProperty As Boolean) As Boolean
        fieldName = ""
        isCustomProperty = False
        If String.IsNullOrWhiteSpace(propertyText) Then Return False
        Dim m As Match = PropertyTextRegex.Match(propertyText.Trim())
        If Not m.Success Then Return False
        fieldName = m.Groups(1).Value.Trim()
        isCustomProperty = propertyText.IndexOf("/CP|", StringComparison.OrdinalIgnoreCase) >= 0
        Return fieldName.Length > 0
    End Function

    Private Shared Function MapMetadataToPropertyField(fieldName As String, data As DrawingMetadataInput) As String
        If data Is Nothing OrElse String.IsNullOrWhiteSpace(fieldName) Then Return ""
        Dim key As String = fieldName.Trim()
        Dim kl As String = key.ToLowerInvariant()

        If kl = "comentarios" OrElse kl = "comments" Then
            Return If(data.Material, "").Trim()
        End If
        If kl = "título" OrElse kl = "titulo" OrElse kl = "title" OrElse kl = "document title" Then
            Dim d As String = If(data.Denominacion, "").Trim()
            If d = "" Then d = If(data.Titulo, "").Trim()
            Return d
        End If
        If kl = "l" OrElse kl = "largo" OrElse kl = "longitud" Then Return If(data.LargoL, "").Trim()
        If kl = "h" OrElse kl = "alto" OrElse kl = "altura" Then Return If(data.AltoH, "").Trim()
        If kl = "d" OrElse kl = "profundidad" OrElse kl = "fondo" Then Return If(data.DatoD, "").Trim()
        If kl = "espesor" OrElse kl = "thickness" Then Return If(data.Espesor, "").Trim()
        If kl.Contains("masa") OrElse kl.Contains("mass") OrElse kl.Contains("weight") OrElse kl.Contains("peso") Then
            Return ExtractMassNumericForProperty(If(data.Peso, ""))
        End If
        If kl = "cantidad" OrElse kl = "quantity" Then Return If(data.Cantidad, "").Trim()
        If kl.Contains("número de artículo") OrElse kl.Contains("numero de articulo") OrElse kl.Contains("article") Then
            Return FirstNonEmpty(If(data.Numero, "").Trim(), If(data.Plano, "").Trim())
        End If
        If kl.Contains("nombre de archivo (sin extensión)") OrElse kl.Contains("nombre de archivo (sin extension)") Then
            Return FileNameWithoutExtension(If(data.NombreArchivo, "").Trim())
        End If
        If kl.Contains("nombre de archivo") OrElse kl = "filename" Then
            Return If(data.NombreArchivo, "").Trim()
        End If
        If kl.Contains("denominacion") OrElse kl.Contains("denominación") Then
            Return FirstNonEmpty(If(data.Denominacion, "").Trim(), If(data.Titulo, "").Trim())
        End If
        If kl = "material" Then Return If(data.Material, "").Trim()
        If kl = "codigo" OrElse kl = "código" Then Return If(data.Plano, "").Trim()

        Return ""
    End Function

    Friend Shared Function IsMetadataPlaceholder(value As String) As Boolean
        If String.IsNullOrWhiteSpace(value) Then Return True
        Dim t As String = value.Trim()
        If t = "--" OrElse t = "-" OrElse t = "—" OrElse t = "n/a" OrElse t = "na" Then Return True
        Return False
    End Function

    Private Shared Function TryWritePropertyOnDocument(doc As Object, fieldName As String, isCustom As Boolean, value As String, logger As Logger) As Boolean
        If doc Is Nothing OrElse IsMetadataPlaceholder(value) Then Return False
        Dim v As String = value.Trim()
        If isCustom Then
            Return SolidEdgePropertyService.TrySetCustomProperty(doc, fieldName, v, logger)
        End If

        Dim kl As String = fieldName.Trim().ToLowerInvariant()
        Select Case kl
            Case "comentarios"
                ' Plantilla CADEBRO: columna MATERIAL → %{Comentarios|G}
                If TrySetDocProp(doc, "SummaryInformation", "Comments", v, logger) Then Return True
                If TrySetDocProp(doc, "SummaryInformation", "Comentarios", v, logger) Then Return True
                If SolidEdgePropertyService.TrySetCustomProperty(doc, "Material", v, logger) Then Return True
                Return False
            Case "título", "titulo"
                If TrySetDocProp(doc, "SummaryInformation", "Title", v, logger) Then Return True
                If TrySetDocProp(doc, "SummaryInformation", "Título", v, logger) Then Return True
                If TrySetDocProp(doc, "SummaryInformation", "Document Title", v, logger) Then Return True
                Return SolidEdgePropertyService.TrySetCustomProperty(doc, "Titulo", v, logger)
            Case "espesor"
                If SolidEdgePropertyService.TrySetCustomProperty(doc, "Espesor", v, logger) Then Return True
                If TrySetDocProp(doc, "MechanicalModeling", "Sheet Metal Gauge", v, logger) Then Return True
                Return TrySetDocProp(doc, "MechanicalModeling", "Thickness", v, logger)
            Case Else
                If TrySetDocProp(doc, "SummaryInformation", fieldName, v, logger) Then Return True
                If TrySetDocProp(doc, "ProjectInformation", fieldName, v, logger) Then Return True
                If TrySetDocProp(doc, "MechanicalModeling", fieldName, v, logger) Then Return True
                Return SolidEdgePropertyService.TrySetCustomProperty(doc, fieldName, v, logger)
        End Select
    End Function

    Private Shared Function TrySetDocProp(doc As Object, setName As String, propName As String, value As String, logger As Logger) As Boolean
        Return SolidEdgePropertyService.SetDocumentPropertyCount(doc, setName, propName, value, logger) > 0
    End Function

    Private Shared Function ExtractMassNumericForProperty(peso As String) As String
        If String.IsNullOrWhiteSpace(peso) Then Return ""
        Dim s As String = peso.Trim().ToLowerInvariant()
        s = s.Replace("kg", "").Replace(" ", "").Trim()
        s = s.Replace(",", ".")
        Dim d As Double
        If Double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, d) Then
            Return d.ToString("G", CultureInfo.InvariantCulture)
        End If
        Return peso.Trim()
    End Function

    Private Shared Function FileNameWithoutExtension(fileName As String) As String
        If String.IsNullOrWhiteSpace(fileName) Then Return ""
        Try
            Return Path.GetFileNameWithoutExtension(fileName.Trim())
        Catch
            Return fileName.Trim()
        End Try
    End Function

    Private Shared Function FirstNonEmpty(ParamArray values() As String) As String
        If values Is Nothing Then Return ""
        For Each v In values
            If Not String.IsNullOrWhiteSpace(v) Then Return v.Trim()
        Next
        Return ""
    End Function
End Class
