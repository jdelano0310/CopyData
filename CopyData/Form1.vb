Imports System.Data.OleDb
Imports System.IO
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Menu

Public Class frmMain
    Dim logFile As StreamWriter
    Dim startupParameter As String
    Dim configFileName As String

    Private Sub btnSelectSourceDB_Click(sender As Object, e As EventArgs) Handles btnSelectSourceDB.Click

        ' allow the user to select the source database
        With ofd
            .Title = "Select Source Access ACCDB"
            .DefaultExt = "*.accdb"

            If .ShowDialog Then
                txtSourceDB.Text = .FileName
                WriteToLog($"Selected source DB of { .FileName} ")
            Else
                txtSourceDB.Text = ""
            End If
        End With

        FillSourceTableDropDown(txtSourceDB.Text)

    End Sub

    Private Sub FillSourceTableDropDown(sourceDBFileName As String)

        ' open the Source Access database and list the tables that are available
        Dim sourceDBCon As OleDbConnection
        Try
            Dim conStrBuilder As New OleDbConnectionStringBuilder With {
                .Provider = "Microsoft.ACE.OLEDB.16.0",
                .DataSource = sourceDBFileName
            }

            sourceDBCon = New OleDbConnection(conStrBuilder.ConnectionString)
            sourceDBCon.Open()

            Dim sourceDBTables As DataTable
            sourceDBTables = sourceDBCon.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, New Object() {Nothing, Nothing, Nothing, "TABLE"})

            Dim dtRow As DataRow
            cboTableToCopy.Items.Clear()

            For Each dtRow In sourceDBTables.Rows
                cboTableToCopy.Items.Add(dtRow(2))
            Next

        Catch ex As Exception
            MsgBox("Could not fill table dropdown " & vbCrLf & ex.Message)
            WriteToLog("Could not fill table dropdown " & ex.Message)

        End Try

        If Not sourceDBCon Is Nothing Then sourceDBCon.Close()

    End Sub

    Private Sub btnSelectDestinationDB_Click(sender As Object, e As EventArgs) Handles btnSelectDestinationDB.Click
        ' allow the user to select the destination database
        With ofd
            .Title = "Select Destination Access MDB"
            .DefaultExt = "Old Access | *.mdb"

            If .ShowDialog Then
                txtDestinationDB.Text = .FileName
            Else
                txtDestinationDB.Text = ""
            End If
        End With

    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click

        ' save this information to be used later
        Dim configFile As StreamWriter

        If File.Exists(configFileName) Then
            ' there is one of these, just delete it
            File.Delete(configFileName)
            WriteToLog("Deleted previous configuration")
        End If

        ' save the settings
        configFile = New StreamWriter(configFileName)
        configFile.WriteLine($"SourceDB,{txtSourceDB.Text}")
        configFile.WriteLine($"SourceTB,{cboTableToCopy.Text}")
        configFile.WriteLine($"DestinationDB,{txtDestinationDB.Text}")
        configFile.Close()

        WriteToLog("Saved configuration")
        WriteToLog($"    Of {txtSourceDB.Text} {cboTableToCopy.Text} {txtDestinationDB.Text}")

    End Sub

    Private Sub WriteToLog(logMessage As String)

        logFile.WriteLine($"{Format(Now, "MM-dd-yy HH:mm:ss")} - {logMessage}")
    End Sub

    Private Sub frmMain_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        WriteToLog("**************** Closing")
        logFile.Close()

    End Sub

    Private Sub cboTableToCopy_Click(sender As Object, e As EventArgs) Handles cboTableToCopy.Click
        WriteToLog($"Clicked Source table {cboTableToCopy.Text} ")
    End Sub

    Private Sub btnRun_Click(sender As Object, e As EventArgs) Handles btnRun.Click

        Dim sourceDBFile As String = ""
        Dim sourceTableName As String = ""
        Dim destinationDBFile As String = ""

        If startupParameter Is Nothing Then
            ' this is running from the form
            sourceDBFile = txtSourceDB.Text
            sourceTableName = cboTableToCopy.Text
            destinationDBFile = txtDestinationDB.Text
            WriteToLog("**************** Running from the form")
        Else
            ' started via a command line 
            LoadConfiguration("CMD", sourceDBFile, sourceTableName, destinationDBFile)
        End If

        WriteToLog("Opening databases")
        WriteToLog($"Source: {sourceDBFile} table: {sourceTableName}")
        WriteToLog($"Destination: {destinationDBFile}")

        Dim sourceDBCon As OleDbConnection
        Dim destinationDBCon As OleDbConnection

        Try
            ' open the databases
            Dim sourceConStr As New OleDbConnectionStringBuilder With {
                .Provider = "Microsoft.ACE.OLEDB.16.0",
                .DataSource = sourceDBFile
            }
            sourceDBCon = New OleDbConnection(sourceConStr.ConnectionString)
            sourceDBCon.Open()

            Dim DestinationConStr As New OleDbConnectionStringBuilder With {
                .Provider = "Microsoft.ACE.OLEDB.12.0",
                .DataSource = destinationDBFile
            }
            destinationDBCon = New OleDbConnection(DestinationConStr.ConnectionString)
            destinationDBCon.Open()

        Catch ex As Exception

            WriteToLog("Could not open " & IIf(sourceDBCon Is Nothing, "source", "destination") & " database due to error " & ex.Message)

            If startupParameter Is Nothing Then
                ' running from the form, display a message box
                MsgBox(MsgBox("Could not open " & IIf(sourceDBCon Is Nothing, "source", "destination") & " database due to error " & vbCrLf & ex.Message))
                Exit Sub
            Else
                Me.Close()
                End
            End If

            Exit Sub
        End Try

        ' export the table to a text file
        Try
            WriteToLog($"Exporting Table")
            File.Delete(Application.StartupPath & "\TableExport.csv")
            Dim exportSourceTable As New OleDbCommand($"SELECT * INTO [Text;HDR=Yes;DATABASE={Application.StartupPath & "\"}].TableExport.csv FROM {sourceTableName}", sourceDBCon)
            exportSourceTable.ExecuteNonQuery()

        Catch ex As Exception

            WriteToLog($"Could not export table {sourceTableName}" & ex.Message)
            If startupParameter Is Nothing Then
                ' running from the form, display a message box
                MsgBox($"Could not export table {sourceTableName}" & vbCrLf & ex.Message)
                Exit Sub
            Else
                Me.Close()
                End
            End If

        End Try

        ' import into the destination MDB
        WriteToLog("Drop existing table")
        Dim dropTableCMD As New OleDbCommand($"DROP TABLE {sourceTableName}", destinationDBCon)

        ' the error would be that the table doesn't exist, and that is fine.
        Try
            dropTableCMD.ExecuteNonQuery()
        Catch
        End Try

        Try
            WriteToLog($"Importing Table")
            Dim importTableCMD As New OleDbCommand("SELECT * " &
            $"INTO {sourceTableName} FROM [Text;FMT=Delimited;HDR=Yes;CharacterSet=850;DATABASE={Application.StartupPath & "\"}].TableExport.csv;", destinationDBCon)
            importTableCMD.ExecuteNonQuery()

        Catch ex As Exception

            WriteToLog($"Could not import table {sourceTableName}" & ex.Message)

            If startupParameter Is Nothing Then
                ' running from the form, display a message box
                MsgBox($"Could not import table {sourceTableName}" & vbCrLf & ex.Message)
                Exit Sub
            Else
                Me.Close()
                End

            End If

        End Try

        WriteToLog("The process has completed")

        ' if this was executed via command line, then close the form
        If Not startupParameter Is Nothing Then
            Me.Close()
            End
        Else
            MsgBox("Process complete")
        End If

    End Sub

    Private Sub LoadConfiguration(calledFrom As String, Optional ByRef sourceDBFile As String = "", Optional ByRef sourceTableName As String = "", Optional ByRef destinationDBFile As String = "")
        ' open the configuration file, if running this on the form then use 
        ' the form controls for the config data, else use the variables
        ' passed ByRef so the values are passwed back to the calling sub
        Dim configFile As New StreamReader(configFileName)

        Select Case calledFrom
            Case "form"

                ' load the saved data
                WriteToLog("Loading saved data to the form")

                Dim cboValue As String
                txtSourceDB.Text = Split(configFile.ReadLine, ",")(1)

                ' check to see if the source db is still there before attempting to use it
                If File.Exists(txtSourceDB.Text) Then
                    FillSourceTableDropDown(txtSourceDB.Text)

                    cboValue = Split(configFile.ReadLine, ",")(1)
                    cboTableToCopy.SelectedIndex = cboTableToCopy.Items.IndexOf(cboValue)

                    If cboTableToCopy.SelectedIndex = -1 Then
                        MsgBox($"The table {cboValue} was not found in the source db")
                    End If
                Else
                    WriteToLog($"Loaded source db {txtSourceDB.Text} doesn't exist.")
                    MsgBox($"The source db {txtSourceDB.Text} does not exist. Please select another.")
                    txtSourceDB.Text = ""
                End If

                txtDestinationDB.Text = Split(configFile.ReadLine, ",")(1)
                If Not File.Exists(txtDestinationDB.Text) Then
                    WriteToLog($"Loaded destination db {txtDestinationDB.Text} doesn't exist.")
                    MsgBox($"The destination db {txtDestinationDB.Text} does not exist. Please select another.")
                    txtDestinationDB.Text = ""
                End If

            Case "CMD"

                ' load the values to the variables needed
                WriteToLog("Loading saved config for command line run")
                sourceDBFile = Split(configFile.ReadLine, ",")(1)
                sourceTableName = Split(configFile.ReadLine, ",")(1)
                destinationDBFile = Split(configFile.ReadLine, ",")(1)

        End Select

        configFile.Close()

    End Sub

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ' open the log file
        Dim logFileName As String = Application.StartupPath & "\CopyData.log"
        configFileName = Application.StartupPath & "\CopyData.cfg"

        logFile = New StreamWriter(logFileName, True)

        If My.Application.CommandLineArgs.Count > 0 Then
            WriteToLog("**************** Running from command line")
            startupParameter = My.Application.CommandLineArgs(0)
            If startupParameter = "RUN" Then
                ' the command line to run and then close has been passed
                If File.Exists(configFileName) Then
                    btnRun_Click(Nothing, Nothing)
                Else
                    WriteToLog($"The configuration file is missing, unable to continue.")
                    Me.Close()
                    End
                End If
            Else
                ' something other than RUN was found in the first parameter
                WriteToLog($"Invalid argument; {startupParameter} was passed.")
                Me.Close()
                End
            End If
        Else
            If File.Exists(configFileName) Then
                ' there is a config file available, load the values
                LoadConfiguration("form")

            End If
        End If

    End Sub

    Private Sub btnViewLog_Click(sender As Object, e As EventArgs) Handles btnViewLog.Click
        ' open notepad with the log file
        Shell($"notepad.exe {Application.StartupPath & "\CopyData.log"}")
    End Sub
End Class
