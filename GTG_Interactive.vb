Imports System
Imports System.Xml
Imports Microsoft.VisualBasic
Imports System.Windows
Imports Microsoft.VisualBasic.ApplicationServices
Imports Microsoft.SqlServer
Imports System.Configuration
Imports System.IO
Imports System.Collections.Specialized
Imports System.Data.OleDb
Imports System.Data.SqlClient

Public Class GTGForm

    Private Event Startup As StartupEventHandler
    Private currentConnectionString
    Private currentSQLState
    Private commandLine
    Private exePath As String = Application.ExecutablePath()
    Private FilePath As String = Application.StartupPath() + "\ScriptData"
    Private settingsPath As String
    Private errorLoad As Boolean = False
    Private sqlValid As Boolean = True
    Private sqlCon As SqlConnection
    Private sqlCom As SqlCommand

    Private Sub GTGForm_Startup(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Load

        FolderCheck()

        Dim CommandLineArgs As System.Collections.ObjectModel.ReadOnlyCollection(Of String) = My.Application.CommandLineArgs
        commandLine = CommandLineArgs.Item(0)
        settingsPath = FilePath + "\" + CommandLineArgs.Item(0) + ".xml"

        Dim loadFailed As Boolean = False

        Dim doc As XmlDocument = New XmlDocument()

        Try
            doc.Load(settingsPath)
        Catch ex As Exception
            loadFailed = True
            MessageBox.Show("Error Loading Settings File Located at " + FilePath.ToString(), "Error...")
        End Try

        If (loadFailed = False) Then
            Read_Settings()
            Update_Textbox()
            ExecuteStatement()
        End If
    End Sub

    'Updates the text boxes in the application to show the values stored in the Settings.XML file
    Private Sub Update_Textbox()

        SettingsFileName.Text = commandLine.ToString() + ".xml"

        ConnectionString.Text = currentConnectionString.ToString

        SQLStatement.Text = currentSQLState.ToString

    End Sub

    'Writes to the Settings.XML file the new settings specified by the user
    Private Sub UpdateSettings()
        Dim settings As XmlWriterSettings = New XmlWriterSettings()
        settings.Indent = True
        Using writer As XmlWriter = XmlWriter.Create(FilePath, settings)
            'Begin writing
            Try
                writer.WriteStartDocument()
                writer.WriteStartElement("Settings")
                writer.WriteElementString("ConnectionString", ConnectionString.Text)
                writer.WriteElementString("SQLStatement", SQLStatement.Text)
                writer.WriteEndElement()
                writer.WriteEndDocument()
            Catch ex As Exception
                MessageBox.Show(ex.ToString, "Error Writing Settings to Settings File")
            End Try
        End Using
        currentConnectionString = ConnectionString.Text
        currentSQLState = SQLStatement.Text
        Read_Settings()
        Update_Textbox()
    End Sub

    'Reads the settings from Settings.XML and stores them into variables
    Private Sub Read_Settings()

        Dim settingsDoc As XmlDocument = New XmlDocument

        Dim nodeList As XmlNodeList

        Dim node As XmlNode

        Try
            settingsDoc.Load(settingsPath)
            nodeList = settingsDoc.SelectNodes("Settings")

            Try
                For Each node In nodeList
                    currentConnectionString = node.Item("ConnectionString").InnerText
                    currentSQLState = node.Item("SQLStatement").InnerText
                Next

            Catch ex As Exception
                MessageBox.Show("Error Reading From Settings File", "Error...", MessageBoxButtons.OK)
            End Try

        Catch ex As Exception
            Dim readSettings As DialogResult = MessageBox.Show("Error Reading Settings from " + settingsPath.ToString(), "Error...")
        End Try

    End Sub

    'Informs the user that closing the application without saving will lose all changes made.
    Private Sub GTGForm_Closing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        If (errorLoad = False) Then

            Dim saveExit As DialogResult = MessageBox.Show("Are you sure you want to exit and lose all changes?", "Warning...", MessageBoxButtons.YesNo)

            If (saveExit = Forms.DialogResult.No) Then
                e.Cancel = True
            End If

        End If

    End Sub

    'Checks to see if the folder "ScriptData" exists in the Application Path. If not, it will ask the user if they want to create it
    Private Sub FolderCheck()

        If (Not System.IO.Directory.Exists(FilePath)) Then

            Dim folderMissing As DialogResult = MessageBox.Show(FilePath + "\ScriptData doesn't exist. Would you like to create it?", "Warning...", MessageBoxButtons.YesNo)

            If (folderMissing = Forms.DialogResult.Yes) Then
                System.IO.Directory.CreateDirectory(FilePath)
            End If

        End If

    End Sub


    Private Sub SQLCheck()
        Dim tempConnection As OleDbConnection = New OleDbConnection()
        tempConnection.ConnectionString = currentConnectionString
        Dim tempSQLCommand = New OleDbCommand("SET NOEXEC ON", tempConnection)
        Try
            tempConnection.Open()
            tempSQLCommand.ExecuteNonQuery()
            tempConnection.Close()
            tempSQLCommand = New OleDbCommand(currentSQLState, tempConnection)
            tempConnection.Open()
            tempSQLCommand.ExecuteNonQuery()
            tempSQLCommand.CommandText = currentSQLState
            tempSQLCommand.ExecuteNonQuery()
        Catch ex As Exception
            sqlValid = False
            MessageBox.Show(ex.ToString(), "Invalid SQL Statement")
        End Try
        tempConnection.Close()
    End Sub

    Private Sub ExecuteStatement()

        Dim connection As OleDbConnection = New OleDbConnection()

        SQLCheck()
        If (sqlValid = True) Then
            Using (connection)

                Try

                    connection.ConnectionString = currentConnectionString

                    connection.Open()

                    Dim sqlCom As OleDbCommand = New OleDbCommand(currentSQLState, connection)

                    sqlCom.ExecuteNonQuery()

                    connection.Close()

                Catch ex As Exception

                    MessageBox.Show(ex.ToString())

                End Try

            End Using

        End If

    End Sub

    Private Sub CheckSQLState_Click(sender As Object, e As EventArgs) Handles CheckSQLState.Click

        SQLCheck()

        If (sqlValid = True) Then

            MessageBox.Show("SQL Statement is Valid", "SQL Check")

        End If

    End Sub

    Private Sub ChangeUDL_Click(sender As Object, e As EventArgs) Handles ChangeUDL.Click
        DataLinkChange()
    End Sub

    'Allows the user to create a new connection string or update an existing valid connection string
    Private Sub DataLinkChange()
        Dim mydlg As New MSDASC.DataLinks()
        Dim OleCon As New OleDbConnection()
        Dim ADOcon As New ADODB.Connection()

        Try
            ADOcon.ConnectionString = currentConnectionString
            OleCon.ConnectionString = ADOcon.ConnectionString
            OleCon.Open()
            If OleCon.State = 1 Then
                OleCon.Close()
                mydlg.PromptEdit(ADOcon)
                OleCon.ConnectionString = ADOcon.ConnectionString
                OleCon.Open()
                If OleCon.State = 1 Then
                    currentConnectionString = OleCon.ConnectionString
                    Update_Textbox()
                    OleCon.Close()
                Else
                    MsgBox("Connection Failed to Open")
                End If
            Else
                ADOcon = mydlg.PromptNew
                OleCon.ConnectionString = ADOcon.ConnectionString
                OleCon.Open()
                If OleCon.State = 1 Then
                    currentConnectionString = OleCon.ConnectionString
                    Update_Textbox()
                    OleCon.Close()
                Else
                    MsgBox("Connection Failed to Open")
                End If
            End If
        Catch ex As Exception
            Dim resetConString As DialogResult = MessageBox.Show("Previous Connection String Is Not Valid. Update Connection String?", "Error...", MessageBoxButtons.OKCancel)
            If resetConString = Forms.DialogResult.OK Then
                ADOcon = mydlg.PromptNew
                Try
                    OleCon.ConnectionString = ADOcon.ConnectionString
                    OleCon.Open()
                Catch er As Exception
                    MessageBox.Show("Connection String is Invalid")
                End Try

                If OleCon.State = 1 Then
                    currentConnectionString = OleCon.ConnectionString
                    Update_Textbox()
                    OleCon.Close()
                Else
                    MsgBox("Connection Failed to Open")
                End If
            Else
                MessageBox.Show("Datalink Update Cancelled")
            End If
        End Try
    End Sub

End Class
