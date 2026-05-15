Option Strict Off

Imports System
Imports SolidEdgeFramework
Imports SolidEdgePart
Imports System.Runtime.InteropServices

''' <summary>Exportación DXF de chapa desarrollada (.psm) sin lógica legacy "Conrad".</summary>
Module FlatDxfExporter

    ''' <summary>Motivo devuelto en errorMessage cuando no hay flat pattern (no es fallo de COM).</summary>
    Public Const SkipReasonNoFlatPattern As String = "NO_FLAT_PATTERN"

    ''' <summary>Exporta DXF flat reutilizando el desarrollo existente. Sin Face/Edge arbitrarios.</summary>
    ''' <param name="logLine">Si se proporciona (p.ej. <c>logger.Log</c>), el número de línea sigue la secuencia del <see cref="Logger"/> del motor sin reiniciar en [001].</param>
    Public Function ExportFlatDxf(app As Application,
                                  modelPath As String,
                                  outPathDxf As String,
                                  ByRef errorMessage As String,
                                  Optional logLine As Action(Of String) = Nothing) As Boolean

        errorMessage = ""
        Dim psmDoc As SheetMetalDocument = Nothing
        Dim models As Models = Nothing

        Try
            FLog(logLine, "[FLAT] Analizando pieza de chapa: " & modelPath)
            Dim obj = app.Documents.Open(modelPath)
            psmDoc = TryCast(obj, SheetMetalDocument)

            Dim isSheetMetal As Boolean = (psmDoc IsNot Nothing)
            FLog(logLine, "[FLAT] Es documento SheetMetal: " & isSheetMetal.ToString())
            If Not isSheetMetal Then
                errorMessage = "No se pudo convertir a SheetMetalDocument (E_NOINTERFACE)."
                FLog(logLine, "[FLAT][ERR] " & errorMessage)
                Return False
            End If

            Dim hasFlat As Boolean = False
            Try
                hasFlat = (psmDoc.FlatPatternModels.Count > 0)
            Catch
                hasFlat = False
            End Try
            FLog(logLine, "[FLAT] Tiene flat existente: " & hasFlat.ToString())

            If Not hasFlat Then
                errorMessage = SkipReasonNoFlatPattern
                FLog(logLine, "[FLAT][WARN] El fichero .psm NO tiene la chapa desarrollada: " & modelPath & ". Tendrías que desarrollarla.")
                FLog(logLine, "[FLAT] Exportación omitida por falta de desarrollo")
                Return False
            End If

            FLog(logLine, "[FLAT] Se reutilizará flat existente")
            FLog(logLine, "[FLAT] No se intentará regenerar flat")

            Try
                models = psmDoc.Models
                If models Is Nothing Then
                    errorMessage = "Models del PSM es Nothing."
                    FLog(logLine, "[FLAT][ERR] " & errorMessage)
                    Return False
                End If

                FLog(logLine, "[FLAT] Método de exportación utilizado: Models.SaveAsFlatDXFEx (sin referencias Face/Edge forzadas)")
                ' Parámetros Nothing: Solid Edge usa el flat pattern existente (mismo criterio estable que reutilizar desarrollo).
                models.SaveAsFlatDXFEx(outPathDxf, Nothing, Nothing, Nothing, True)
                FLog(logLine, "[FLAT] DXF Flat exportado correctamente: " & outPathDxf)
                Return True

            Catch ex As Exception
                Dim cex As COMException = TryCast(ex, COMException)
                If cex IsNot Nothing Then
                    errorMessage = $"{ex.Message} (HR=0x{cex.ErrorCode:X8})"
                Else
                    errorMessage = ex.Message
                End If
                FLog(logLine, "[FLAT][ERR] Error real: " & errorMessage)
                Return False
            End Try

        Finally
            Try
                If psmDoc IsNot Nothing Then psmDoc.Close(False)
            Catch
            End Try
            TryReleaseComObject(models)
            TryReleaseComObject(psmDoc)
        End Try

    End Function

    Private Sub FLog(logLine As Action(Of String), msg As String)
        If logLine IsNot Nothing Then
            logLine.Invoke(msg)
        Else
            LogUtil.EmitUnnumbered(msg)
        End If
    End Sub

    Private Sub TryReleaseComObject(obj As Object)
        If obj Is Nothing Then Return
        Try
            If Marshal.IsComObject(obj) Then Marshal.ReleaseComObject(obj)
        Catch
        End Try
    End Sub

End Module
