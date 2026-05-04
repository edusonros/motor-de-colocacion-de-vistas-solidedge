Option Strict Off

Imports System

''' <summary>Metadatos unificados para cajetín DFT y PartsList PART_LIST (propiedades de documento).</summary>
Public Class DrawingMetadataInput

    ' --- CAJETÍN ---
    Public Property Cliente As String
    Public Property Proyecto As String
    Public Property Pedido As String
    ''' <summary>Origen legible del pedido (p. ej. Custom.Pedido, Custom.Order), para UI y resumen «Datos de plano».</summary>
    Public Property PedidoSource As String
    Public Property Plano As String
    Public Property Titulo As String
    Public Property Revision As String
    Public Property Autor As String
    Public Property Fecha As Date = Date.Today
    ''' <summary>Origen de la fecha (p. ej. modelo, vacío).</summary>
    Public Property FechaSource As String

    ' --- PARTSLIST (columnas visibles / Custom.*) ---
    Public Property Numero As String = "1"
    Public Property Cantidad As String = "1"
    Public Property NombreArchivo As String
    Public Property Denominacion As String
    Public Property Material As String
    Public Property Espesor As String
    Public Property LargoL As String
    Public Property AltoH As String
    Public Property DatoD As String
    Public Property Peso As String

    ' --- Origen cajetín (UI) ---
    Public Property ClienteSource As String
    Public Property ProyectoSource As String
    Public Property PlanoSource As String
    Public Property TituloSource As String
    Public Property RevisionSource As String
    Public Property AutorSource As String

    ' --- TRAZABILIDAD UI (origen del dato) ---
    Public Property NumeroSource As String
    Public Property CantidadSource As String
    Public Property NombreArchivoSource As String
    Public Property DenominacionSource As String
    Public Property MaterialSource As String
    Public Property EspesorSource As String
    Public Property LHDSource As String
    Public Property PesoSource As String

    Public Function CloneShallow() As DrawingMetadataInput
        Return DirectCast(Me.MemberwiseClone(), DrawingMetadataInput)
    End Function

End Class
