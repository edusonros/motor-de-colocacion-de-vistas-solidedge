Option Strict Off

''' <summary>
''' Tres motores lógicos documentados por la UI (Generador vistas / Gestor metadatos / Acotación).
''' La ejecución sigue canalizándose en <see cref="DraftGenerationEngine"/> mediante <see cref="DraftMotorPhase"/>.
''' Este archivo no altera algoritmos; sólo marca la fase sobre una copia de <see cref="JobConfiguration"/>.
''' </summary>
Public NotInheritable Class DraftViewGenerationMotor
    Private Sub New()
    End Sub

    Public Shared Function PrepareCopy(uiConfiguration As JobConfiguration) As JobConfiguration
        Dim c As JobConfiguration = uiConfiguration.CloneForExecution()
        c.MotorPhase = DraftMotorPhase.ViewGeneration
        Return c
    End Function
End Class

Public NotInheritable Class DraftMetadataMotor
    Private Sub New()
    End Sub

    Public Shared Function PrepareCopy(uiConfiguration As JobConfiguration) As JobConfiguration
        Dim c As JobConfiguration = uiConfiguration.CloneForExecution()
        c.MotorPhase = DraftMotorPhase.MetadataManagement
        Return c
    End Function
End Class

Public NotInheritable Class DraftDimensioningMotor
    Private Sub New()
    End Sub

    Public Shared Function PrepareCopy(uiConfiguration As JobConfiguration) As JobConfiguration
        Dim c As JobConfiguration = uiConfiguration.CloneForExecution()
        c.MotorPhase = DraftMotorPhase.Dimensioning
        Return c
    End Function
End Class

''' <summary>Secuencia completa histórica (botón GENERAR).</summary>
Public NotInheritable Class DraftFullSequenceMotor
    Private Sub New()
    End Sub

    Public Shared Function PrepareCopy(uiConfiguration As JobConfiguration) As JobConfiguration
        Dim c As JobConfiguration = uiConfiguration.CloneForExecution()
        c.MotorPhase = DraftMotorPhase.FullSequence
        Return c
    End Function
End Class
