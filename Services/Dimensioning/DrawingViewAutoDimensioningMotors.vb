Option Strict Off

Imports System.Collections.Generic
Imports System.Globalization
Imports SolidEdgeDraft

''' <summary>
''' Contrato aislado para colocar cotas en vistas de pliego. Un solo <see cref="IDrawingViewAutoDimensioningMotor"/>
''' se ejecuta por pasada; la factoría elige implementación según <see cref="AutoDimensioningMotorKind"/>.
''' </summary>
Friend Interface IDrawingViewAutoDimensioningMotor
    Sub Run(
        draft As DraftDocument,
        mainView As DrawingView,
        log As DimensionLogger,
        appLogger As Logger,
        norm As DimensioningNormConfig,
        protectedZones As IList(Of ProtectedZone2D))
End Interface

Friend NotInheritable Class DrawingViewAutoDimensioningMotorFactory
    Private Sub New()
    End Sub

    Friend Shared Function Create(kind As AutoDimensioningMotorKind) As IDrawingViewAutoDimensioningMotor
        Select Case kind
            Case AutoDimensioningMotorKind.LegacyV02IsolatedCopy
                Return LegacyV02IsolatedDrawingViewAutoDimensioningMotor.Instance
            Case AutoDimensioningMotorKind.AlternatePlugIn
                Return AlternatePlugInDrawingViewAutoDimensioningMotor.Instance
            Case Else
                Return MainPipelineDrawingViewAutoDimensioningMotor.Instance
        End Select
    End Function
End Class

Friend NotInheritable Class MainPipelineDrawingViewAutoDimensioningMotor
    Implements IDrawingViewAutoDimensioningMotor

    Friend Shared ReadOnly Instance As New MainPipelineDrawingViewAutoDimensioningMotor()

    Private Sub New()
    End Sub

    Public Sub Run(
        draft As DraftDocument,
        mainView As DrawingView,
        log As DimensionLogger,
        appLogger As Logger,
        norm As DimensioningNormConfig,
        protectedZones As IList(Of ProtectedZone2D)) Implements IDrawingViewAutoDimensioningMotor.Run
        UniqueDvAutoDimensioningEngine.Run(draft, log, appLogger, norm, protectedZones)
    End Sub
End Class

Friend NotInheritable Class LegacyV02IsolatedDrawingViewAutoDimensioningMotor
    Implements IDrawingViewAutoDimensioningMotor

    Friend Shared ReadOnly Instance As New LegacyV02IsolatedDrawingViewAutoDimensioningMotor()

    Private Sub New()
    End Sub

    Public Sub Run(
        draft As DraftDocument,
        mainView As DrawingView,
        log As DimensionLogger,
        appLogger As Logger,
        norm As DimensioningNormConfig,
        protectedZones As IList(Of ProtectedZone2D)) Implements IDrawingViewAutoDimensioningMotor.Run
        LegacyV02DimensionMotorBridge.Run(draft, appLogger, norm, protectedZones)
    End Sub
End Class

''' <summary>
''' Enganche para un segundo motor (solo lógica COM/cotas propia). No invoca <see cref="UniqueDvAutoDimensioningEngine"/>.
''' Amplía esta clase con tu pipeline; el valor <see cref="AutoDimensioningMotorKind.AlternatePlugIn"/> ya está cableado en UI y factoría.
''' </summary>
Friend NotInheritable Class AlternatePlugInDrawingViewAutoDimensioningMotor
    Implements IDrawingViewAutoDimensioningMotor

    Friend Shared ReadOnly Instance As New AlternatePlugInDrawingViewAutoDimensioningMotor()

    Private Sub New()
    End Sub

    Public Sub Run(
        draft As DraftDocument,
        mainView As DrawingView,
        log As DimensionLogger,
        appLogger As Logger,
        norm As DimensioningNormConfig,
        protectedZones As IList(Of ProtectedZone2D)) Implements IDrawingViewAutoDimensioningMotor.Run
        appLogger?.Log("[DIM][ALT_PLUGIN] Motor alternativo (sin UniqueDv): enganche listo; sustituir por tu acotación COM.")
        If draft Is Nothing Then
            appLogger?.Log("[DIM][ALT_PLUGIN][SKIP] DraftDocument Nothing.")
            Return
        End If
        Try
            Dim viewCount As Integer = 0
            Dim sh As Sheet = Nothing
            Try
                sh = draft.ActiveSheet
            Catch
                sh = Nothing
            End Try
            If sh IsNot Nothing Then
                Try
                    viewCount = sh.DrawingViews.Count
                Catch
                    viewCount = -1
                End Try
            End If
            appLogger?.Log("[DIM][ALT_PLUGIN] hoja_activa_vistas=" & viewCount.ToString(CultureInfo.InvariantCulture))
        Catch ex As Exception
            appLogger?.Log("[DIM][ALT_PLUGIN][ERR] " & ex.Message)
        End Try
    End Sub
End Class
