Option Strict Off

''' <summary>Filtros de exclusión compartidos con <see cref="DraftGenerationEngine"/> y la UI ASM.</summary>
Public NotInheritable Class LaserCutPartFilters
    Public Shared ReadOnly ExcludeKeywords As String() = {
        "skf", "nut", "2026", "2026_02", "screw", "duin", "iso", "bolt", "whaser", "washer",
        "snl", "sleeve", "22210", "22211", "22212", "fnl", "motor", "prensa", "estopada", "tornillo",
        "tuerca", "arandela", "fag"
    }

    Private Sub New()
    End Sub

    Public Shared Function ShouldExcludeFile(filePath As String) As Boolean
        If String.IsNullOrWhiteSpace(filePath) Then Return True
        Dim baseName As String = IO.Path.GetFileNameWithoutExtension(filePath).ToLowerInvariant()
        Return ExcludeKeywords.Any(Function(k) baseName.Contains(k))
    End Function
End Class
