Option Strict Off

Imports System.Runtime.InteropServices
Imports SolidEdgeAssembly
Imports SolidEdgeFramework

Public Class AssemblyComponentItem
    Public Property FullPath As String = ""
    Public Property Kind As String = "" ' ASM / PAR / PSM
    Public Property DisplayName As String = ""
    Public Property Level As Integer = 0
End Class

Public Class AssemblyComponentService
    Public Shared Function LoadAssemblyComponentItems(asmPath As String,
                                                      uniqueOnly As Boolean,
                                                      showSolidEdge As Boolean,
                                                      logger As Logger,
                                                      Optional progress As Action(Of String, Integer, Integer) = Nothing) As List(Of AssemblyComponentItem)
        Dim result As New List(Of AssemblyComponentItem)()
        If String.IsNullOrWhiteSpace(asmPath) OrElse Not IO.File.Exists(asmPath) Then Return result

        Dim app As Application = Nothing
        Dim asmDoc As AssemblyDocument = Nothing
        Dim createdByUs As Boolean = False
        Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Dim processedCount As Integer = 0

        Try
            OleMessageFilter.Register()
            logger.Log("[ASM] Inicio lectura ensamblaje")
            If progress IsNot Nothing Then progress.Invoke("Inicio lectura ensamblaje", 0, -1)
            Try
                app = CType(Marshal.GetActiveObject("SolidEdge.Application"), Application)
                logger.Log("Solid Edge: instancia existente detectada para lectura de componentes ASM.")
            Catch
                Dim t = Type.GetTypeFromProgID("SolidEdge.Application")
                app = CType(Activator.CreateInstance(t), Application)
                createdByUs = True
                logger.Log("Solid Edge: nueva instancia creada para lectura de componentes ASM.")
            End Try

            app.Visible = showSolidEdge
            app.DisplayAlerts = False

            logger.Log("[ASM] Obteniendo ocurrencias")
            If progress IsNot Nothing Then progress.Invoke("Obteniendo ocurrencias", 0, -1)
            asmDoc = CType(app.Documents.Open(asmPath), AssemblyDocument)
            logger.Log("[ASM] Recorriendo componentes")
            If progress IsNot Nothing Then progress.Invoke("Recorriendo componentes", 0, -1)
            RecurseOccurrences(asmDoc.Occurrences, result, seen, uniqueOnly, 0, processedCount, progress)
            logger.Log($"Componentes ASM detectados: {result.Count}")
            logger.Log("[ASM] Lectura finalizada")
            If progress IsNot Nothing Then progress.Invoke("Lectura finalizada", processedCount, Math.Max(1, processedCount))
            Return result

        Catch ex As Exception
            logger.LogException("LoadAssemblyComponentItems", ex)
            Return result
        Finally
            Try
                If asmDoc IsNot Nothing Then asmDoc.Close(False)
            Catch
            End Try
            Try
                If app IsNot Nothing AndAlso createdByUs Then app.Quit()
            Catch
            End Try
            Try : OleMessageFilter.Revoke() : Catch : End Try
        End Try
    End Function

    Private Shared Sub RecurseOccurrences(occurs As Occurrences,
                                          output As List(Of AssemblyComponentItem),
                                          seen As HashSet(Of String),
                                          uniqueOnly As Boolean,
                                          depth As Integer,
                                          ByRef processedCount As Integer,
                                          progress As Action(Of String, Integer, Integer))
        For Each occ As Occurrence In occurs
            Try
                processedCount += 1
                If progress IsNot Nothing AndAlso (processedCount <= 10 OrElse (processedCount Mod 25) = 0) Then
                    progress.Invoke("Procesando componente", processedCount, -1)
                End If

                Dim docObj As Object = occ.OccurrenceDocument
                If docObj Is Nothing Then Continue For
                Dim fullName As String = CStr(CallByName(docObj, "FullName", CallType.Get))
                If String.IsNullOrWhiteSpace(fullName) Then Continue For

                Dim ext As String = IO.Path.GetExtension(fullName).ToLowerInvariant()
                Dim kind As String = ""
                If ext = ".asm" Then kind = "ASM"
                If ext = ".par" Then kind = "PAR"
                If ext = ".psm" Then kind = "PSM"
                If kind = "" Then Continue For
                If progress IsNot Nothing AndAlso (processedCount <= 10 OrElse (processedCount Mod 25) = 0) Then
                    progress.Invoke("Clasificando documentos", processedCount, -1)
                End If

                Dim addIt As Boolean = True
                If uniqueOnly Then
                    addIt = Not seen.Contains(fullName)
                End If
                If addIt Then
                    seen.Add(fullName)
                    Dim indent As String = New String(" "c, Math.Max(0, depth) * 4)
                    output.Add(New AssemblyComponentItem With {
                        .FullPath = fullName,
                        .Kind = kind,
                        .DisplayName = $"{indent}[{kind}] {IO.Path.GetFileName(fullName)}",
                        .Level = depth
                    })
                End If

                If kind = "ASM" Then
                    Dim subAsm As AssemblyDocument = TryCast(docObj, AssemblyDocument)
                    If subAsm IsNot Nothing Then
                        RecurseOccurrences(subAsm.Occurrences, output, seen, uniqueOnly, depth + 1, processedCount, progress)
                    End If
                End If
            Catch
            End Try
        Next
    End Sub
End Class
