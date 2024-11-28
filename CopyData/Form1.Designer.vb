<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMain
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtSourceDB = New System.Windows.Forms.TextBox()
        Me.btnSelectSourceDB = New System.Windows.Forms.Button()
        Me.btnSelectDestinationDB = New System.Windows.Forms.Button()
        Me.txtDestinationDB = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.cboTableToCopy = New System.Windows.Forms.ComboBox()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.btnRun = New System.Windows.Forms.Button()
        Me.btnViewLog = New System.Windows.Forms.Button()
        Me.ofd = New System.Windows.Forms.OpenFileDialog()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(15, 14)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(90, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Source Database"
        '
        'txtSourceDB
        '
        Me.txtSourceDB.Location = New System.Drawing.Point(15, 32)
        Me.txtSourceDB.Margin = New System.Windows.Forms.Padding(2)
        Me.txtSourceDB.Name = "txtSourceDB"
        Me.txtSourceDB.Size = New System.Drawing.Size(509, 20)
        Me.txtSourceDB.TabIndex = 1
        '
        'btnSelectSourceDB
        '
        Me.btnSelectSourceDB.Location = New System.Drawing.Point(415, 55)
        Me.btnSelectSourceDB.Margin = New System.Windows.Forms.Padding(2)
        Me.btnSelectSourceDB.Name = "btnSelectSourceDB"
        Me.btnSelectSourceDB.Size = New System.Drawing.Size(107, 23)
        Me.btnSelectSourceDB.TabIndex = 2
        Me.btnSelectSourceDB.Text = "Select Source DB"
        Me.btnSelectSourceDB.UseVisualStyleBackColor = True
        '
        'btnSelectDestinationDB
        '
        Me.btnSelectDestinationDB.Location = New System.Drawing.Point(392, 164)
        Me.btnSelectDestinationDB.Margin = New System.Windows.Forms.Padding(2)
        Me.btnSelectDestinationDB.Name = "btnSelectDestinationDB"
        Me.btnSelectDestinationDB.Size = New System.Drawing.Size(129, 23)
        Me.btnSelectDestinationDB.TabIndex = 5
        Me.btnSelectDestinationDB.Text = "Select Destination DB"
        Me.btnSelectDestinationDB.UseVisualStyleBackColor = True
        '
        'txtDestinationDB
        '
        Me.txtDestinationDB.Location = New System.Drawing.Point(15, 141)
        Me.txtDestinationDB.Margin = New System.Windows.Forms.Padding(2)
        Me.txtDestinationDB.Name = "txtDestinationDB"
        Me.txtDestinationDB.Size = New System.Drawing.Size(509, 20)
        Me.txtDestinationDB.TabIndex = 4
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(15, 123)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(109, 13)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Destination Database"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(22, 66)
        Me.Label3.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(73, 13)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "Table to Copy"
        '
        'cboTableToCopy
        '
        Me.cboTableToCopy.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboTableToCopy.FormattingEnabled = True
        Me.cboTableToCopy.Location = New System.Drawing.Point(97, 64)
        Me.cboTableToCopy.Margin = New System.Windows.Forms.Padding(2)
        Me.cboTableToCopy.Name = "cboTableToCopy"
        Me.cboTableToCopy.Size = New System.Drawing.Size(257, 21)
        Me.cboTableToCopy.TabIndex = 7
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(452, 235)
        Me.btnSave.Margin = New System.Windows.Forms.Padding(2)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(69, 23)
        Me.btnSave.TabIndex = 8
        Me.btnSave.Text = "Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'btnRun
        '
        Me.btnRun.Location = New System.Drawing.Point(15, 235)
        Me.btnRun.Margin = New System.Windows.Forms.Padding(2)
        Me.btnRun.Name = "btnRun"
        Me.btnRun.Size = New System.Drawing.Size(69, 23)
        Me.btnRun.TabIndex = 9
        Me.btnRun.Text = "Run"
        Me.btnRun.UseVisualStyleBackColor = True
        '
        'btnViewLog
        '
        Me.btnViewLog.Location = New System.Drawing.Point(88, 235)
        Me.btnViewLog.Margin = New System.Windows.Forms.Padding(2)
        Me.btnViewLog.Name = "btnViewLog"
        Me.btnViewLog.Size = New System.Drawing.Size(69, 23)
        Me.btnViewLog.TabIndex = 10
        Me.btnViewLog.Text = "View Log"
        Me.btnViewLog.UseVisualStyleBackColor = True
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(533, 265)
        Me.Controls.Add(Me.btnViewLog)
        Me.Controls.Add(Me.btnRun)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.cboTableToCopy)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.btnSelectDestinationDB)
        Me.Controls.Add(Me.txtDestinationDB)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.btnSelectSourceDB)
        Me.Controls.Add(Me.txtSourceDB)
        Me.Controls.Add(Me.Label1)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.Name = "frmMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Copy Data"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label1 As Label
    Friend WithEvents txtSourceDB As TextBox
    Friend WithEvents btnSelectSourceDB As Button
    Friend WithEvents btnSelectDestinationDB As Button
    Friend WithEvents txtDestinationDB As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents cboTableToCopy As ComboBox
    Friend WithEvents btnSave As Button
    Friend WithEvents btnRun As Button
    Friend WithEvents btnViewLog As Button
    Protected WithEvents ofd As OpenFileDialog
End Class
