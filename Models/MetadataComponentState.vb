Option Strict Off

Imports System
Imports System.Collections.Generic

''' <summary>Estado del botón «Datos» por componente ASM.</summary>
Public Enum ComponentMetadataStatus
    Pending
    PartialComplete
    Complete
End Enum

''' <summary>Resultado global de validación de metadatos para una pieza/contexto.</summary>
Public Enum MetadataValidationOutcome
    Complete
    RequiredMissing
    WarningMissing
End Enum

''' <summary>Resultado de <see cref="DrawingMetadataService.ValidateMetadataForComponent"/>.</summary>
Public Class MetadataValidationResult
    Public Property Outcome As MetadataValidationOutcome = MetadataValidationOutcome.Complete
    Public Property MissingRequiredFields As New List(Of String)()
    Public Property MissingWarningFields As New List(Of String)()
End Class

''' <summary>Metadatos y validación asociados a un archivo de componente (clave = ruta completa).</summary>
Public Class ComponentMetadataState
    Public Property ComponentPath As String = ""
    Public Property Metadata As DrawingMetadataInput
    Public Property Status As ComponentMetadataStatus = ComponentMetadataStatus.Pending
    Public Property MissingRequiredFields As New List(Of String)()
    Public Property MissingWarningFields As New List(Of String)()
    Public Property LastLoadedUtc As DateTime?
End Class
