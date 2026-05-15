Option Strict Off
Imports SolidEdgeDraft
Namespace Extraer_dft_dxf_flatdxf.LegacyV02Dimensioning
' --- Copia aislada desde Planos_Automaticos_v02 (no editar v02); motor independiente del arbol Services\Dimensioning vigente. ---
Friend NotInheritable Class SolidEdgeContext
    Public Property Draft As DraftDocument
    Public Property Sheet As Sheet

    Private Sub New()
    End Sub

    Public Shared Function TryCreate(draft As DraftDocument, log As DimensionLogger, ByRef ctx As SolidEdgeContext) As Boolean
        ctx = Nothing
        If draft Is Nothing Then
            log?.LogLine("[DIM][DOC][ERR] DraftDocument activo no disponible.")
            Return False
        End If

        Dim sh As Sheet = Nothing
        Try
            sh = draft.ActiveSheet
        Catch ex As Exception
            log?.ComFail("DraftDocument.ActiveSheet", "DraftDocument", ex)
            Return False
        End Try
        If sh Is Nothing Then
            log?.LogLine("[DIM][SHEET][ERR] Hoja activa no disponible.")
            Return False
        End If

        ctx = New SolidEdgeContext With {
            .Draft = draft,
            .Sheet = sh
        }
        Return True
    End Function
End Class

End Namespace

