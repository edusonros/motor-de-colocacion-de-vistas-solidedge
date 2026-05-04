Option Strict Off

Imports System.Data
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text.RegularExpressions
Imports SolidEdgeFileProperties

''' <summary>
''' Construye la vista de trazabilidad a partir de PropertySets reales en disco (origen y DFT opcional).
''' </summary>
Public NotInheritable Class TraceabilityInspectionService

    Private Sub New()
    End Sub

    Public Shared Function BuildTraceabilityDataTable(sourcePath As String,
                                                      draftPath As String,
                                                      kind As SourceFileKind,
                                                      logger As Logger,
                                                      Optional onlyCajetinFields As Boolean = True) As DataTable
        Dim dt As New DataTable("Trazabilidad")
        dt.Columns.Add("CampoLogico", GetType(String))
        dt.Columns.Add("PropertySet", GetType(String))
        dt.Columns.Add("NombrePropiedad", GetType(String))
        dt.Columns.Add("Alcance", GetType(String))
        dt.Columns.Add("ValorModelo", GetType(String))
        dt.Columns.Add("ValorDraft", GetType(String))
        dt.Columns.Add("Escribible", GetType(String))
        dt.Columns.Add("Sobrescribible", GetType(String))
        dt.Columns.Add("OrigenRecomendado", GetType(String))
        dt.Columns.Add("Observaciones", GetType(String))

        Dim usedBindings As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Dim modelSets As PropertySets = Nothing
        Dim draftSets As PropertySets = Nothing

        Try
            If Not String.IsNullOrWhiteSpace(sourcePath) AndAlso File.Exists(sourcePath) Then
                modelSets = OpenPropertySetsReadOnly(sourcePath, logger)
            End If
            If Not String.IsNullOrWhiteSpace(draftPath) AndAlso File.Exists(draftPath) Then
                draftSets = OpenPropertySetsReadOnly(draftPath, logger)
            End If

            Dim mapped As Integer = 0
            For Each def In TraceabilityFieldCatalog.GetDefinitions()
                If def Is Nothing Then Continue For
                If onlyCajetinFields AndAlso Not TraceabilityFieldCatalog.CajetinLogicalKeys.Contains(def.LogicalKey) Then Continue For
                If Not TraceabilityFieldCatalog.AppliesToKind(def, kind) Then Continue For

                Dim hit As BindingHit = ResolveBinding(modelSets, draftSets, def, usedBindings)
                Dim scope As String = DescribeScope(modelSets IsNot Nothing, draftSets IsNot Nothing, hit.ModelValue, hit.DraftValue)
                Dim writable As String = If(def.WritableStandard OrElse def.AllowCustomCreate, "Sí", "Parcial")
                Dim over As String = If(def.AllowCustomCreate OrElse def.WritableStandard, "Sí", "Parcial")
                Dim note As String = If(def.Important, "Campo prioritario para cajetín.", "")
                If hit.BindingUsed IsNot Nothing Then
                    note = If(String.IsNullOrWhiteSpace(note), "", note & " ") & "Resuelto vía " & hit.BindingUsed.SetKey & "." & hit.BindingUsed.PropertyName
                End If

                dt.Rows.Add(
                    def.DisplayLabel,
                    If(hit.BindingUsed IsNot Nothing, hit.BindingUsed.SetKey, ""),
                    If(hit.BindingUsed IsNot Nothing, hit.BindingUsed.PropertyName, ""),
                    scope,
                    hit.ModelValue,
                    hit.DraftValue,
                    writable,
                    over,
                    def.RecommendedOrigin,
                    note.Trim())
                If (Not String.IsNullOrWhiteSpace(hit.ModelValue)) OrElse (Not String.IsNullOrWhiteSpace(hit.DraftValue)) Then
                    mapped += 1
                End If
            Next

            If Not onlyCajetinFields Then
                AppendMergedCustomRows(dt, modelSets, draftSets, usedBindings, logger)
            End If

            ' [TRACE] suprimido en UI: la tabla y la lectura de propiedades no cambian.

        Catch ex As Exception
            If logger IsNot Nothing Then logger.LogException("BuildTraceabilityDataTable", ex)
        Finally
            CloseRelease(modelSets)
            CloseRelease(draftSets)
        End Try

        Return dt
    End Function

    Private Shared Function KindToLabel(kind As SourceFileKind) As String
        Select Case kind
            Case SourceFileKind.PartFile : Return "PAR"
            Case SourceFileKind.SheetMetalFile : Return "PSM"
            Case SourceFileKind.AssemblyFile : Return "ASM"
            Case Else : Return "Desconocido"
        End Select
    End Function

    Private Class BindingHit
        Public ModelValue As String = ""
        Public DraftValue As String = ""
        Public BindingUsed As TraceabilityFieldCatalog.BindingInfo
    End Class

    Private Shared Function ResolveBinding(modelSets As PropertySets, draftSets As PropertySets,
                                          def As TraceabilityFieldCatalog.FieldDefinition,
                                          usedBindings As HashSet(Of String)) As BindingHit
        Dim hit As New BindingHit()
        For Each b In def.Bindings
            If b Is Nothing Then Continue For
            Dim key = b.SetKey & "|" & b.PropertyName
            Dim mv As String = ""
            Dim dv As String = ""
            If modelSets IsNot Nothing Then
                mv = SolidEdgePropertyService.TryGetFilePropertyValue(modelSets, b.SetKey, b.PropertyName)
            End If
            If draftSets IsNot Nothing Then
                dv = SolidEdgePropertyService.TryGetFilePropertyValue(draftSets, b.SetKey, b.PropertyName)
            End If
            If Not String.IsNullOrWhiteSpace(mv) OrElse Not String.IsNullOrWhiteSpace(dv) Then
                hit.ModelValue = mv
                hit.DraftValue = dv
                hit.BindingUsed = b
                usedBindings.Add(key)
                Return hit
            End If
        Next
        ' Sin valor: usar primer binding para mostrar dónde escribiría el programa
        If def.Bindings.Count > 0 Then
            hit.BindingUsed = def.Bindings(0)
            usedBindings.Add(def.Bindings(0).SetKey & "|" & def.Bindings(0).PropertyName)
        End If
        Return hit
    End Function

    Private Shared Function DescribeScope(hasModel As Boolean, hasDraft As Boolean, mv As String, dv As String) As String
        If hasModel AndAlso hasDraft Then Return "Modelo + Draft"
        If hasModel Then Return "Modelo"
        If hasDraft Then Return "Draft"
        Return "—"
    End Function

    Private Shared Sub AppendMergedCustomRows(dt As DataTable,
                                              modelSets As PropertySets,
                                              draftSets As PropertySets,
                                              usedBindings As HashSet(Of String),
                                              logger As Logger)
        Try
            Dim customNames As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            For Each u In usedBindings
                If u.StartsWith("Custom|", StringComparison.OrdinalIgnoreCase) Then
                    customNames.Add(u.Substring("Custom|".Length))
                End If
            Next

            Dim modelVals = ReadCustomDictionary(modelSets)
            Dim draftVals = ReadCustomDictionary(draftSets)

            Dim names As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            For Each kv In modelVals.Keys
                names.Add(kv)
            Next
            For Each kv In draftVals.Keys
                names.Add(kv)
            Next

            For Each n In names
                If customNames.Contains(n) Then Continue For
                Dim mv As String = ""
                Dim dv As String = ""
                modelVals.TryGetValue(n, mv)
                draftVals.TryGetValue(n, dv)
                If String.IsNullOrWhiteSpace(mv) AndAlso String.IsNullOrWhiteSpace(dv) Then Continue For
                If Not String.IsNullOrWhiteSpace(mv) AndAlso mv.Length > 280 Then Continue For
                If Not String.IsNullOrWhiteSpace(dv) AndAlso dv.Length > 280 Then Continue For
                If IsNoisePropertyName(n) Then Continue For

                Dim scope As String = "Modelo + Draft"
                If modelSets Is Nothing Then
                    scope = "Draft"
                ElseIf draftSets Is Nothing Then
                    scope = "Modelo"
                End If

                dt.Rows.Add(
                    "Custom: " & n,
                    "Custom",
                    n,
                    scope,
                    mv,
                    dv,
                    "Sí",
                    "Sí",
                    "Custom",
                    "Propiedad personalizada con valor (no estaba en el catálogo principal).")
            Next
        Catch ex As Exception
            If logger IsNot Nothing Then logger.Log($"[WARN] Trazabilidad custom extra: {ex.Message}")
        End Try
    End Sub

    Private Shared Function ReadCustomDictionary(psets As PropertySets) As Dictionary(Of String, String)
        Dim d As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
        If psets Is Nothing Then Return d
        Dim setObj As Object = SolidEdgePropertyService.ResolveFilePropertySetForInspection(psets, "Custom")
        If setObj Is Nothing Then Return d
        Dim cnt As Integer = 0
        Try : cnt = CInt(CallByName(setObj, "Count", CallType.Get)) : Catch : End Try
        For i As Integer = 1 To cnt
            Dim propObj As Object = Nothing
            Try
                propObj = SolidEdgePropertyService.GetCollectionItemForInspection(setObj, i)
                If propObj Is Nothing Then Continue For
                Dim n As String = ""
                Try : n = CStr(CallByName(propObj, "Name", CallType.Get)) : Catch : End Try
                If String.IsNullOrWhiteSpace(n) Then Continue For
                Dim v As String = ""
                Try
                    Dim vo As Object = CallByName(propObj, "Value", CallType.Get)
                    If vo IsNot Nothing Then v = vo.ToString().Trim()
                Catch
                End Try
                If Not String.IsNullOrWhiteSpace(v) Then d(n) = v
            Finally
                Try
                    If propObj IsNot Nothing AndAlso Marshal.IsComObject(propObj) Then Marshal.ReleaseComObject(propObj)
                Catch
                End Try
            End Try
        Next
        Try
            If setObj IsNot Nothing AndAlso Marshal.IsComObject(setObj) Then Marshal.ReleaseComObject(setObj)
        Catch
        End Try
        Return d
    End Function

    Private Shared Function IsNoisePropertyName(n As String) As Boolean
        If n.Length > 60 Then Return True
        If Regex.IsMatch(n, "^[0-9A-Fa-f\-]{32,}$") Then Return True
        Return False
    End Function

    Private Shared Function OpenPropertySetsReadOnly(path As String, logger As Logger) As PropertySets
        Dim p As PropertySets = Nothing
        Try
            p = New PropertySets()
            p.Open(path, True)
            Return p
        Catch ex As Exception
            If logger IsNot Nothing Then logger.Log($"[WARN] No se pudieron abrir PropertySets: {path} -> {ex.Message}")
            Try
                If p IsNot Nothing AndAlso Marshal.IsComObject(p) Then Marshal.ReleaseComObject(p)
            Catch
            End Try
            Return Nothing
        End Try
    End Function

    Private Shared Sub CloseRelease(psets As PropertySets)
        If psets Is Nothing Then Return
        Try : psets.Close() : Catch : End Try
        Try
            If Marshal.IsComObject(psets) Then Marshal.ReleaseComObject(psets)
        Catch
        End Try
    End Sub

    Private Shared Function CountPropertiesRough(psets As PropertySets) As Integer
        Dim total As Integer = 0
        Try
            Dim sc As Integer = 0
            Try : sc = CInt(CallByName(psets, "Count", CallType.Get)) : Catch : End Try
            For si As Integer = 1 To sc
                Dim setObj As Object = Nothing
                Try
                    setObj = SolidEdgePropertyService.GetCollectionItemForInspection(psets, si)
                    If setObj Is Nothing Then Continue For
                    Dim pc As Integer = 0
                    Try : pc = CInt(CallByName(setObj, "Count", CallType.Get)) : Catch : End Try
                    total += pc
                Finally
                    Try
                        If setObj IsNot Nothing AndAlso Marshal.IsComObject(setObj) Then Marshal.ReleaseComObject(setObj)
                    Catch
                    End Try
                End Try
            Next
        Catch
        End Try
        Return total
    End Function

End Class
