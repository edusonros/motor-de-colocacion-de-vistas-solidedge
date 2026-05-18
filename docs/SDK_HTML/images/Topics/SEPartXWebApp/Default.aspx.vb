Imports System

Public Class WebForm1
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        ' Store the filename to be displayed in the SEPartX ActiveX Control in strFilename variable
        Dim strFilename As String = "C:\\TestPartCP.par"

        Try
            ' Check if the filename exists. If true call the ViewInputFile procedure to display the
            ' file in the SEPartX ActiveX Control, else display appropriate error message to the user
            ' that the file doe not exist.
            If My.Computer.FileSystem.FileExists(strFilename) = True Then
                ViewInputFile(strFilename)
            Else
                MsgBox("The file : " & strFilename & " does not exist. Please set a valid filename.")
            End If

        Catch ex As Exception
            'do nothing
        End Try
    End Sub

    Public Sub ViewInputFile(ByVal strFilename As String)

        ' Create the OBJECT tag with the various attributes for the SEPartX ActiveX Control and
        ' set it to the label's TEXT property i.e. lblSEPartxView.Text
        lblSEPartxView.Text = " <OBJECT id='separtx' style='LEFT:10px; WIDTH:100%; TOP:10px; HEIGHT:100%' "
        lblSEPartxView.Text += "align='middle' classid='clsid:E03935A5-FBAD-11D0-8AC7-0800362FB302' VIEWASTEXT >"
        lblSEPartxView.Text += "<PARAM NAME='BackColor' VALUE='14466712'>" ' set the default 3D fileype color
        lblSEPartxView.Text += "<PARAM NAME='BackColor2D' VALUE='12632256'>" ' set the default 2D fileype color
        lblSEPartxView.Text += "<PARAM NAME='BorderStyle' VALUE='0'>" ' set the border style
        lblSEPartxView.Text += "<PARAM NAME='PartFile' VALUE='" & strFilename & "'>" ' set the filename
        lblSEPartxView.Text += "<PARAM NAME='ViewType' VALUE='iso'>" ' set the view projection type
        lblSEPartxView.Text += "<PARAM NAME='MouseAction' VALUE='none'>" ' set the mouse dynamics operation
        lblSEPartxView.Text += "<PARAM NAME='ViewPerspective' VALUE='0'>" ' set the flag which indicates if file is to be displayed in perspective projection or not
        lblSEPartxView.Text += "<PARAM NAME='ShowToolbar' VALUE='1'>" ' set the flag indicating if the toolbar is displayed or not
        lblSEPartxView.Text += "</OBJECT>"

    End Sub
End Class
