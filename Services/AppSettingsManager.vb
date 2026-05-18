Option Strict Off

Imports System.IO
Imports System.Runtime.Serialization
Imports System.Runtime.Serialization.Json
Imports System.Text

<DataContract>
Public Class PersistedAppSettings
    <DataMember> Public Property LastInputFile As String = ""
    <DataMember> Public Property LastOutputFolder As String = ""
    <DataMember> Public Property TemplateA4 As String = ""
    <DataMember> Public Property TemplateA3 As String = ""
    <DataMember> Public Property TemplateA2 As String = ""
    <DataMember> Public Property TemplateDxf As String = ""

    <DataMember> Public Property CreateDraft As Boolean = True
    <DataMember> Public Property CreatePdf As Boolean = True
    <DataMember> Public Property CreateDxfFromDraft As Boolean = True
    <DataMember> Public Property CreateFlatDxf As Boolean = True
    <DataMember> Public Property OpenOutputFolderWhenDone As Boolean = True
    <DataMember> Public Property OverwriteExisting As Boolean = False
    <DataMember> Public Property ProcessRepeatedComponentsOnce As Boolean = True
    <DataMember> Public Property DetailedLog As Boolean = True
    <DataMember> Public Property DebugTemplatesInspection As Boolean = False
    <DataMember> Public Property KeepSolidEdgeVisible As Boolean = False
    <DataMember> Public Property InsertPropertiesInTitleBlock As Boolean = False
    <DataMember> Public Property TitleBlockPropertySourceMode As TitleBlockPropertySource = TitleBlockPropertySource.FromModelLink

    <DataMember> Public Property PreferredFormat As PreferredSheetFormat = PreferredSheetFormat.Auto
    <DataMember> Public Property UseAutomaticScale As Boolean = True
    <DataMember> Public Property ManualScale As Double = 1.0
    <DataMember> Public Property IncludeIsometric As Boolean = True
    <DataMember> Public Property IncludeProjectedViews As Boolean = True
    <DataMember> Public Property IncludeFlatInDraftWhenPsm As Boolean = True
    <DataMember> Public Property EnableAutoDimensioning As Boolean = False
    <DataMember> Public Property AutoDimensioningMotor As Integer = 0
    <DataMember> Public Property EnableProductionDvRefCleanEngine As Boolean = False
    <DataMember> Public Property EnableSesdkPostDimensionIntrospection As Boolean = False
    <DataMember> Public Property PreferSweepAllDrawingDimensions As Boolean = False
    <DataMember> Public Property SuppressDimensionTrackDistanceSpacing As Boolean = False
    <DataMember> Public Property EnableKeypointValueDuplicateCleanup As Boolean = True

    <DataMember> Public Property EnableDrawingViewDimensioningLab As Boolean = False
    <DataMember> Public Property RunDropViewsTo2DModelLab As Boolean = False
    <DataMember> Public Property RunDropCreatedSheetsDimensionLab As Boolean = False
    <DataMember> Public Property DropCreatedSheetsDimensionLabDebugSave As Boolean = False
    <DataMember> Public Property RunDVGeometryDimensionPlacementLab As Boolean = True
    <DataMember> Public Property RunDVGeometryMethodDiscoveryLab As Boolean = False

    <DataMember> Public Property EnableDimLabInteractivePause As Boolean = False
    ''' <summary>Modo DIMLAB serializado como entero (enum DimLabMode). Por defecto Full (=2).</summary>
    <DataMember> Public Property DimLabMode As Integer = 2
    <DataMember> Public Property EnableDimLabVisibleProbe As Boolean = False
    <DataMember> Public Property EnableDimLabAlternativePlacement As Boolean = False
    <DataMember> Public Property EnableDimLabHorizontalControlInVerticalOnly As Boolean = False
    <DataMember> Public Property DimLabKeepFailedDimensions As Boolean = False
    <DataMember> Public Property DimLabCleanPreviousLabDimensions As Boolean = False
    <DataMember> Public Property EnablePmiRetrievalProbe As Boolean = False
    <DataMember> Public Property ExperimentalCreatePMIModelViewIfMissing As Boolean = False
    <DataMember> Public Property ExperimentalDraftGeometryDiagnostics As Boolean = False
    <DataMember> Public Property UseBestBaseViewLogic As Boolean = True

    <DataMember> Public Property ClientName As String = ""
    <DataMember> Public Property ProjectName As String = ""
    <DataMember> Public Property DrawingTitle As String = ""
    <DataMember> Public Property TitleSourceMode As TitleSourceMode = TitleSourceMode.Manual
    <DataMember> Public Property Material As String = ""
    <DataMember> Public Property Thickness As String = ""
    <DataMember> Public Property Pedido As String = ""
    <DataMember> Public Property AuthorName As String = ""
    <DataMember> Public Property Weight As String = ""
    <DataMember> Public Property Equipment As String = ""
    <DataMember> Public Property DrawingNumber As String = ""
    <DataMember> Public Property Revision As String = ""
    <DataMember> Public Property Notes As String = ""
    <DataMember> Public Property StrictMetadataValidation As Boolean = False

    <DataMember> Public Property WindowLeft As Integer = -1
    <DataMember> Public Property WindowTop As Integer = -1
    <DataMember> Public Property WindowWidth As Integer = 1400
    <DataMember> Public Property WindowHeight As Integer = 900
    <DataMember> Public Property WindowStateValue As Integer = 0
End Class

Public Class AppSettingsManager
    Private Shared ReadOnly SettingsDirectory As String =
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Conrad", "DraftAutomation")

    Private Shared ReadOnly SettingsFilePath As String =
        Path.Combine(SettingsDirectory, "settings.json")

    Public Shared Function LoadSettings() As PersistedAppSettings
        Try
            If Not File.Exists(SettingsFilePath) Then Return New PersistedAppSettings()
            Using fs As New FileStream(SettingsFilePath, FileMode.Open, FileAccess.Read, FileShare.Read)
                Dim ser As New DataContractJsonSerializer(GetType(PersistedAppSettings))
                Return CType(ser.ReadObject(fs), PersistedAppSettings)
            End Using
        Catch
            Return New PersistedAppSettings()
        End Try
    End Function

    Public Shared Sub SaveSettings(settings As PersistedAppSettings)
        If Not TrySaveSettings(settings, Nothing) Then
            Throw New IOException("No se pudo guardar la configuracion. Revise permisos y antivirus en: " & SettingsFilePath)
        End If
    End Sub

    ''' <summary>
    ''' Escribe <c>settings.json</c> en un temporal y sustituye el destino (reduce fallos por archivo bloqueado).
    ''' </summary>
    ''' <param name="errorDetail">Si no es Nothing, recibe el mensaje de error (sin lanzar).</param>
    Public Shared Function TrySaveSettings(settings As PersistedAppSettings, ByRef errorDetail As String) As Boolean
        errorDetail = Nothing
        If settings Is Nothing Then Return True
        Try
            If Not Directory.Exists(SettingsDirectory) Then Directory.CreateDirectory(SettingsDirectory)
        Catch ex As Exception
            errorDetail = ex.Message
            Return False
        End Try

        Dim tmpPath As String = SettingsFilePath & ".tmp"
        Try
            Using fs As New FileStream(tmpPath, FileMode.Create, FileAccess.Write, FileShare.Read)
                Dim ser As New DataContractJsonSerializer(GetType(PersistedAppSettings))
                ser.WriteObject(fs, settings)
                fs.Flush(True)
            End Using

            If File.Exists(SettingsFilePath) Then
                File.Replace(tmpPath, SettingsFilePath, Nothing)
            Else
                File.Move(tmpPath, SettingsFilePath)
            End If
            Return True
        Catch ex As Exception
            errorDetail = ex.Message
            Try
                If File.Exists(tmpPath) Then File.Delete(tmpPath)
            Catch
            End Try
            Return False
        End Try
    End Function

    Public Shared Function GetSettingsFilePath() As String
        Return SettingsFilePath
    End Function
End Class
