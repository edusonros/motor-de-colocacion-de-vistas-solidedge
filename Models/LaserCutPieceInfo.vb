Option Strict Off

''' <summary>Pieza detectada para el plano de corte láser (resumen editable en UI).</summary>
Public Class LaserCutPieceInfo
    Public Property Include As Boolean = True
    Public Property FilePath As String = ""
    Public Property FileName As String = ""
    Public Property FileNameNoExt As String = ""
    ''' <summary>PAR o PSM.</summary>
    Public Property FileType As String = ""
    Public Property ThicknessMm As Double? = Nothing
    Public Property ThicknessText As String = ""
    Public Property Material As String = ""
    Public Property Quantity As Integer = 1
    Public Property IsSheetMetal As Boolean = False
    ''' <summary>Nothing = no aplica (PAR).</summary>
    Public Property HasFlatPattern As Boolean? = Nothing
    ''' <summary>Nothing = desconocido; True = lleva pliegue.</summary>
    Public Property IsBent As Boolean? = Nothing
    Public Property SourceDftPath As String = ""
    Public Property Status As String = "Pendiente"
    Public Property Notes As String = ""
End Class

Public Class LaserCutGenerateResult
    Public Property Success As Boolean
    Public Property DftPath As String = ""
    Public Property DxfPath As String = ""
    Public Property BendListPath As String = ""
    Public Property LaserLogPath As String = ""
    Public Property ErrorMessage As String = ""
End Class
