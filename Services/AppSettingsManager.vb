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
    <DataMember> Public Property EnableAutoDimensioning As Boolean = True
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
        If settings Is Nothing Then Return
        If Not Directory.Exists(SettingsDirectory) Then Directory.CreateDirectory(SettingsDirectory)
        Using fs As New FileStream(SettingsFilePath, FileMode.Create, FileAccess.Write, FileShare.None)
            Dim ser As New DataContractJsonSerializer(GetType(PersistedAppSettings))
            ser.WriteObject(fs, settings)
        End Using
    End Sub

    Public Shared Function GetSettingsFilePath() As String
        Return SettingsFilePath
    End Function
End Class
