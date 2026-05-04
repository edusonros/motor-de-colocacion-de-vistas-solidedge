Option Strict Off

Imports System.Collections.Generic
Imports System.Collections.ObjectModel

''' <summary>
''' Catálogo central de campos lógicos de trazabilidad/cajetín y sus enlaces a PropertySets de Solid Edge.
''' Mantener alineado con la lógica de sincronización en SolidEdgePropertyService.BuildPropertySyncProfile.
''' </summary>
Public NotInheritable Class TraceabilityFieldCatalog

    Public Class BindingInfo
        Public Property SetKey As String
        Public Property PropertyName As String
        Public Property IsCustomSet As Boolean
    End Class

    Public Class FieldDefinition
        Public Property LogicalKey As String
        Public Property DisplayLabel As String
        Public Property AppliesToPar As Boolean = True
        Public Property AppliesToPsm As Boolean = True
        Public Property AppliesToAsm As Boolean = True
        Public Property Important As Boolean
        ''' <summary>Model / Draft / SummaryInfo / Custom según uso recomendado en el flujo actual.</summary>
        Public Property RecommendedOrigin As String
        Public Property WritableStandard As Boolean
        Public Property AllowCustomCreate As Boolean
        Public ReadOnly Property Bindings As New List(Of BindingInfo)()
    End Class

    Private Shared ReadOnly _fields As List(Of FieldDefinition) = BuildDefinitions()

    ''' <summary>Campos del bloque «Datos de cajetín» (tabla resumida en UI).</summary>
    Public Shared ReadOnly CajetinLogicalKeys As ReadOnlyCollection(Of String) =
        New ReadOnlyCollection(Of String)(New String() {
            "Titulo", "Proyecto", "ClienteEmpresa", "NumeroDocumento", "Revision", "Autor", "Material", "Espesor", "Pedido"
        })

    Public Shared Function GetDefinitions() As IList(Of FieldDefinition)
        Return _fields
    End Function

    Public Shared Function AppliesToKind(def As FieldDefinition, kind As SourceFileKind) As Boolean
        If def Is Nothing Then Return False
        Select Case kind
            Case SourceFileKind.PartFile : Return def.AppliesToPar
            Case SourceFileKind.SheetMetalFile : Return def.AppliesToPsm
            Case SourceFileKind.AssemblyFile : Return def.AppliesToAsm
            Case Else : Return True
        End Select
    End Function

    Private Shared Function BuildDefinitions() As List(Of FieldDefinition)
        Dim list As New List(Of FieldDefinition)()

        Dim title As New FieldDefinition With {
            .LogicalKey = "Titulo",
            .DisplayLabel = "Título",
            .Important = True,
            .RecommendedOrigin = "SummaryInfo",
            .WritableStandard = True,
            .AllowCustomCreate = True}
        title.Bindings.Add(New BindingInfo With {.SetKey = "SummaryInformation", .PropertyName = "Title", .IsCustomSet = False})
        title.Bindings.Add(New BindingInfo With {.SetKey = "SummaryInformation", .PropertyName = "Document Title", .IsCustomSet = False})
        title.Bindings.Add(New BindingInfo With {.SetKey = "SummaryInformation", .PropertyName = "Título", .IsCustomSet = False})
        title.Bindings.Add(New BindingInfo With {.SetKey = "Custom", .PropertyName = "Titulo", .IsCustomSet = True})
        title.Bindings.Add(New BindingInfo With {.SetKey = "Custom", .PropertyName = "Título", .IsCustomSet = True})
        list.Add(title)

        Dim proj As New FieldDefinition With {
            .LogicalKey = "Proyecto",
            .DisplayLabel = "Proyecto",
            .Important = True,
            .RecommendedOrigin = "Model",
            .WritableStandard = True,
            .AllowCustomCreate = True}
        proj.Bindings.Add(New BindingInfo With {.SetKey = "ProjectInformation", .PropertyName = "Project Name", .IsCustomSet = False})
        proj.Bindings.Add(New BindingInfo With {.SetKey = "ProjectInformation", .PropertyName = "Nombre de proyecto", .IsCustomSet = False})
        proj.Bindings.Add(New BindingInfo With {.SetKey = "Custom", .PropertyName = "Proyecto", .IsCustomSet = True})
        list.Add(proj)

        Dim client As New FieldDefinition With {
            .LogicalKey = "ClienteEmpresa",
            .DisplayLabel = "Cliente / Empresa",
            .Important = True,
            .RecommendedOrigin = "DocumentSummary",
            .WritableStandard = True,
            .AllowCustomCreate = True}
        client.Bindings.Add(New BindingInfo With {.SetKey = "DocumentSummaryInformation", .PropertyName = "Company", .IsCustomSet = False})
        client.Bindings.Add(New BindingInfo With {.SetKey = "DocumentSummaryInformation", .PropertyName = "Empresa", .IsCustomSet = False})
        client.Bindings.Add(New BindingInfo With {.SetKey = "Custom", .PropertyName = "Cliente", .IsCustomSet = True})
        list.Add(client)

        Dim docNum As New FieldDefinition With {
            .LogicalKey = "NumeroDocumento",
            .DisplayLabel = "Número de documento",
            .Important = True,
            .RecommendedOrigin = "Model",
            .WritableStandard = True,
            .AllowCustomCreate = True}
        docNum.Bindings.Add(New BindingInfo With {.SetKey = "ProjectInformation", .PropertyName = "Document Number", .IsCustomSet = False})
        docNum.Bindings.Add(New BindingInfo With {.SetKey = "ProjectInformation", .PropertyName = "Número de documento", .IsCustomSet = False})
        docNum.Bindings.Add(New BindingInfo With {.SetKey = "Custom", .PropertyName = "Numero de plano", .IsCustomSet = True})
        list.Add(docNum)

        Dim rev As New FieldDefinition With {
            .LogicalKey = "Revision",
            .DisplayLabel = "Revisión",
            .Important = True,
            .RecommendedOrigin = "Model",
            .WritableStandard = True,
            .AllowCustomCreate = True}
        rev.Bindings.Add(New BindingInfo With {.SetKey = "ProjectInformation", .PropertyName = "Revision", .IsCustomSet = False})
        rev.Bindings.Add(New BindingInfo With {.SetKey = "ProjectInformation", .PropertyName = "Revisión", .IsCustomSet = False})
        rev.Bindings.Add(New BindingInfo With {.SetKey = "Custom", .PropertyName = "Revision", .IsCustomSet = True})
        list.Add(rev)

        Dim author As New FieldDefinition With {
            .LogicalKey = "Autor",
            .DisplayLabel = "Autor",
            .Important = False,
            .RecommendedOrigin = "SummaryInfo",
            .WritableStandard = True,
            .AllowCustomCreate = False}
        author.Bindings.Add(New BindingInfo With {.SetKey = "SummaryInformation", .PropertyName = "Author", .IsCustomSet = False})
        author.Bindings.Add(New BindingInfo With {.SetKey = "SummaryInformation", .PropertyName = "Autor", .IsCustomSet = False})
        list.Add(author)

        Dim subject As New FieldDefinition With {
            .LogicalKey = "Asunto",
            .DisplayLabel = "Asunto / Subject",
            .Important = False,
            .RecommendedOrigin = "SummaryInfo",
            .WritableStandard = True,
            .AllowCustomCreate = False}
        subject.Bindings.Add(New BindingInfo With {.SetKey = "SummaryInformation", .PropertyName = "Subject", .IsCustomSet = False})
        list.Add(subject)

        Dim comments As New FieldDefinition With {
            .LogicalKey = "Comentarios",
            .DisplayLabel = "Comentarios",
            .Important = False,
            .RecommendedOrigin = "SummaryInfo",
            .WritableStandard = True,
            .AllowCustomCreate = True}
        comments.Bindings.Add(New BindingInfo With {.SetKey = "SummaryInformation", .PropertyName = "Comments", .IsCustomSet = False})
        comments.Bindings.Add(New BindingInfo With {.SetKey = "Custom", .PropertyName = "Observaciones", .IsCustomSet = True})
        list.Add(comments)

        Dim mat As New FieldDefinition With {
            .LogicalKey = "Material",
            .DisplayLabel = "Material",
            .Important = True,
            .RecommendedOrigin = "MechanicalModeling",
            .WritableStandard = True,
            .AllowCustomCreate = True}
        mat.Bindings.Add(New BindingInfo With {.SetKey = "MechanicalModeling", .PropertyName = "Material", .IsCustomSet = False})
        mat.Bindings.Add(New BindingInfo With {.SetKey = "ProjectInformation", .PropertyName = "Material", .IsCustomSet = False})
        mat.Bindings.Add(New BindingInfo With {.SetKey = "Custom", .PropertyName = "Material", .IsCustomSet = True})
        list.Add(mat)

        Dim thick As New FieldDefinition With {
            .LogicalKey = "Espesor",
            .DisplayLabel = "Espesor / chapa",
            .Important = True,
            .RecommendedOrigin = "MechanicalModeling",
            .WritableStandard = True,
            .AllowCustomCreate = True}
        thick.Bindings.Add(New BindingInfo With {.SetKey = "MechanicalModeling", .PropertyName = "Sheet Metal Gauge", .IsCustomSet = False})
        thick.Bindings.Add(New BindingInfo With {.SetKey = "Custom", .PropertyName = "Espesor", .IsCustomSet = True})
        list.Add(thick)

        Dim pedido As New FieldDefinition With {
            .LogicalKey = "Pedido",
            .DisplayLabel = "Pedido",
            .Important = False,
            .RecommendedOrigin = "Custom",
            .WritableStandard = False,
            .AllowCustomCreate = True}
        pedido.Bindings.Add(New BindingInfo With {.SetKey = "Custom", .PropertyName = "Pedido", .IsCustomSet = True})
        pedido.Bindings.Add(New BindingInfo With {.SetKey = "Custom", .PropertyName = "PEDIDO", .IsCustomSet = True})
        pedido.Bindings.Add(New BindingInfo With {.SetKey = "Custom", .PropertyName = "Order", .IsCustomSet = True})
        pedido.Bindings.Add(New BindingInfo With {.SetKey = "ProjectInformation", .PropertyName = "Stock Number", .IsCustomSet = False})
        list.Add(pedido)

        Dim mass As New FieldDefinition With {
            .LogicalKey = "Peso",
            .DisplayLabel = "Masa / Peso",
            .Important = False,
            .RecommendedOrigin = "MechanicalModeling",
            .WritableStandard = True,
            .AllowCustomCreate = True}
        mass.Bindings.Add(New BindingInfo With {.SetKey = "MechanicalModeling", .PropertyName = "Mass", .IsCustomSet = False})
        mass.Bindings.Add(New BindingInfo With {.SetKey = "MechanicalModeling", .PropertyName = "Weight", .IsCustomSet = False})
        mass.Bindings.Add(New BindingInfo With {.SetKey = "Custom", .PropertyName = "Peso", .IsCustomSet = True})
        list.Add(mass)

        Dim equip As New FieldDefinition With {
            .LogicalKey = "Equipo",
            .DisplayLabel = "Equipo",
            .Important = False,
            .RecommendedOrigin = "Custom",
            .WritableStandard = False,
            .AllowCustomCreate = True}
        equip.Bindings.Add(New BindingInfo With {.SetKey = "Custom", .PropertyName = "Equipo", .IsCustomSet = True})
        list.Add(equip)

        Return list
    End Function

End Class
