''' <summary>Resultado de aplicar cajetín (SummaryInfo) y PART_LIST sobre un DFT (y modelo enlazado si existe).</summary>
Public Class ApplyMetadataToDraftFilesResult
    Public Property SummaryInfoOk As Boolean
    Public Property PartListPipelineRan As Boolean
    Public Property DftSaved As Boolean
    Public Property ModelSaved As Boolean

    Public ReadOnly Property OverallOk As Boolean
        Get
            Return SummaryInfoOk OrElse PartListPipelineRan
        End Get
    End Property
End Class
