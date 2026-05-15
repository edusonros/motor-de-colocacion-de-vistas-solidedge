Option Strict Off

Imports System.Collections.Generic
Imports System.Reflection
Imports SolidEdgeDraft

''' <summary>
''' Entrada pública al motor copiado desde <c>Planos_Automaticos_v02</c>, aislado en
''' <see cref="Extraer_dft_dxf_flatdxf.LegacyV02Dimensioning"/>. Convierte norma y zonas del job principal por reflexión
''' para no compartir tipos con el motor vigente.
''' </summary>
Public NotInheritable Class LegacyV02DimensionMotorBridge
    Private Sub New()
    End Sub

    Public Shared Sub Run(draft As DraftDocument, appLogger As Logger, norm As DimensioningNormConfig, protectedZones As IList(Of ProtectedZone2D))
        If draft Is Nothing Then Return
        Dim legacyNorm = CloneNormReflection(norm)
        Dim legacyZones = CloneProtectedZones(protectedZones)
        Dim log As New Extraer_dft_dxf_flatdxf.LegacyV02Dimensioning.DimensionLogger(appLogger)
        appLogger?.Log("[DIM][PIPE][V02] LegacyV02DimensionMotorBridge → namespace LegacyV02Dimensioning.UniqueDvAutoDimensioningEngine")
        Extraer_dft_dxf_flatdxf.LegacyV02Dimensioning.UniqueDvAutoDimensioningEngine.Run(draft, log, appLogger, legacyNorm, legacyZones)
    End Sub

    Private Shared Function CloneNormReflection(src As DimensioningNormConfig) As Extraer_dft_dxf_flatdxf.LegacyV02Dimensioning.DimensioningNormConfig
        If src Is Nothing Then Return Extraer_dft_dxf_flatdxf.LegacyV02Dimensioning.DimensioningNormConfig.DefaultConfig()
        Dim dst As New Extraer_dft_dxf_flatdxf.LegacyV02Dimensioning.DimensioningNormConfig()
        Dim tSrc As Type = GetType(DimensioningNormConfig)
        Dim tDst As Type = GetType(Extraer_dft_dxf_flatdxf.LegacyV02Dimensioning.DimensioningNormConfig)
        For Each p As PropertyInfo In tSrc.GetProperties(BindingFlags.Public Or BindingFlags.Instance)
            Try
                If Not p.CanRead OrElse p.GetIndexParameters().Length <> 0 Then Continue For
                Dim pd As PropertyInfo = tDst.GetProperty(p.Name, BindingFlags.Public Or BindingFlags.Instance)
                If pd Is Nothing OrElse Not pd.CanWrite Then Continue For
                Dim v As Object = p.GetValue(src, Nothing)
                If v Is Nothing Then
                    pd.SetValue(dst, Nothing, Nothing)
                    Continue For
                End If
                If Not pd.PropertyType.IsAssignableFrom(v.GetType()) Then Continue For
                pd.SetValue(dst, v, Nothing)
            Catch
            End Try
        Next
        Return dst
    End Function

    Private Shared Function CloneProtectedZones(src As IList(Of ProtectedZone2D)) As IList(Of Extraer_dft_dxf_flatdxf.LegacyV02Dimensioning.ProtectedZone2D)
        If src Is Nothing OrElse src.Count = 0 Then Return Nothing
        Dim list As New List(Of Extraer_dft_dxf_flatdxf.LegacyV02Dimensioning.ProtectedZone2D)()
        For Each z As ProtectedZone2D In src
            If z Is Nothing Then Continue For
            list.Add(New Extraer_dft_dxf_flatdxf.LegacyV02Dimensioning.ProtectedZone2D With {
                .Name = z.Name,
                .MinX = z.MinX,
                .MinY = z.MinY,
                .MaxX = z.MaxX,
                .MaxY = z.MaxY
            })
        Next
        Return list
    End Function
End Class
