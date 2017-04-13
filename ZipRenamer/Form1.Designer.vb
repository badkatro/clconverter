<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
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
        Me.components = New System.ComponentModel.Container()
        Me.ChooseZip_Btn = New System.Windows.Forms.Button()
        Me.FilesList_Lview = New System.Windows.Forms.ListView()
        Me.ClearList_Lbl = New System.Windows.Forms.Label()
        Me.Process_Btn = New System.Windows.Forms.Button()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.Message_Lbl = New System.Windows.Forms.Label()
        Me.ToggleOptions_Lbl = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.AutoClean_Folders_Cbox = New System.Windows.Forms.CheckBox()
        Me.OpenOutputFolder_Cbox = New System.Windows.Forms.CheckBox()
        Me.ChooseWorkingFolder_Lbl = New System.Windows.Forms.Label()
        Me.CtWorkingFolder_TxtBox = New System.Windows.Forms.TextBox()
        Me.WorkingFolder_Lbl = New System.Windows.Forms.Label()
        Me.WorkingFolderBrowserDialog = New System.Windows.Forms.FolderBrowserDialog()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'ChooseZip_Btn
        '
        Me.ChooseZip_Btn.Location = New System.Drawing.Point(12, 12)
        Me.ChooseZip_Btn.Name = "ChooseZip_Btn"
        Me.ChooseZip_Btn.Size = New System.Drawing.Size(75, 23)
        Me.ChooseZip_Btn.TabIndex = 0
        Me.ChooseZip_Btn.Text = "Choose Zip"
        Me.ChooseZip_Btn.UseVisualStyleBackColor = True
        '
        'FilesList_Lview
        '
        Me.FilesList_Lview.AllowDrop = True
        Me.FilesList_Lview.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.FilesList_Lview.Location = New System.Drawing.Point(93, 12)
        Me.FilesList_Lview.Name = "FilesList_Lview"
        Me.FilesList_Lview.Size = New System.Drawing.Size(358, 149)
        Me.FilesList_Lview.TabIndex = 1
        Me.FilesList_Lview.UseCompatibleStateImageBehavior = False
        Me.FilesList_Lview.View = System.Windows.Forms.View.List
        '
        'ClearList_Lbl
        '
        Me.ClearList_Lbl.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.ClearList_Lbl.Font = New System.Drawing.Font("Wingdings", 21.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(2, Byte))
        Me.ClearList_Lbl.ForeColor = System.Drawing.Color.Red
        Me.ClearList_Lbl.Location = New System.Drawing.Point(453, 7)
        Me.ClearList_Lbl.Name = "ClearList_Lbl"
        Me.ClearList_Lbl.Size = New System.Drawing.Size(38, 33)
        Me.ClearList_Lbl.TabIndex = 3
        Me.ClearList_Lbl.Text = "ý"
        '
        'Process_Btn
        '
        Me.Process_Btn.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Process_Btn.Location = New System.Drawing.Point(12, 130)
        Me.Process_Btn.Name = "Process_Btn"
        Me.Process_Btn.Size = New System.Drawing.Size(75, 31)
        Me.Process_Btn.TabIndex = 4
        Me.Process_Btn.Text = "Process"
        Me.Process_Btn.UseVisualStyleBackColor = True
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.ProgressBar1.Location = New System.Drawing.Point(93, 146)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(357, 15)
        Me.ProgressBar1.TabIndex = 5
        '
        'Message_Lbl
        '
        Me.Message_Lbl.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Message_Lbl.BackColor = System.Drawing.SystemColors.GradientInactiveCaption
        Me.Message_Lbl.Location = New System.Drawing.Point(15, 172)
        Me.Message_Lbl.Name = "Message_Lbl"
        Me.Message_Lbl.Size = New System.Drawing.Size(470, 15)
        Me.Message_Lbl.TabIndex = 6
        '
        'ToggleOptions_Lbl
        '
        Me.ToggleOptions_Lbl.BackColor = System.Drawing.Color.Transparent
        Me.ToggleOptions_Lbl.Font = New System.Drawing.Font("Webdings", 20.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(2, Byte))
        Me.ToggleOptions_Lbl.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.ToggleOptions_Lbl.ImageAlign = System.Drawing.ContentAlignment.TopLeft
        Me.ToggleOptions_Lbl.Location = New System.Drawing.Point(452, 135)
        Me.ToggleOptions_Lbl.Name = "ToggleOptions_Lbl"
        Me.ToggleOptions_Lbl.Size = New System.Drawing.Size(38, 33)
        Me.ToggleOptions_Lbl.TabIndex = 7
        Me.ToggleOptions_Lbl.Text = "@"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.AutoClean_Folders_Cbox)
        Me.GroupBox1.Controls.Add(Me.OpenOutputFolder_Cbox)
        Me.GroupBox1.Controls.Add(Me.ChooseWorkingFolder_Lbl)
        Me.GroupBox1.Controls.Add(Me.CtWorkingFolder_TxtBox)
        Me.GroupBox1.Controls.Add(Me.WorkingFolder_Lbl)
        Me.GroupBox1.Location = New System.Drawing.Point(498, 5)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(304, 184)
        Me.GroupBox1.TabIndex = 8
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Options"
        '
        'AutoClean_Folders_Cbox
        '
        Me.AutoClean_Folders_Cbox.AutoSize = True
        Me.AutoClean_Folders_Cbox.Location = New System.Drawing.Point(9, 96)
        Me.AutoClean_Folders_Cbox.Name = "AutoClean_Folders_Cbox"
        Me.AutoClean_Folders_Cbox.Size = New System.Drawing.Size(220, 17)
        Me.AutoClean_Folders_Cbox.TabIndex = 4
        Me.AutoClean_Folders_Cbox.Text = "Auto-clean folders when processing ends"
        Me.AutoClean_Folders_Cbox.UseVisualStyleBackColor = True
        '
        'OpenOutputFolder_Cbox
        '
        Me.OpenOutputFolder_Cbox.AutoSize = True
        Me.OpenOutputFolder_Cbox.Location = New System.Drawing.Point(9, 72)
        Me.OpenOutputFolder_Cbox.Name = "OpenOutputFolder_Cbox"
        Me.OpenOutputFolder_Cbox.Size = New System.Drawing.Size(170, 17)
        Me.OpenOutputFolder_Cbox.TabIndex = 3
        Me.OpenOutputFolder_Cbox.Text = "Open output folder when done"
        Me.OpenOutputFolder_Cbox.UseVisualStyleBackColor = True
        '
        'ChooseWorkingFolder_Lbl
        '
        Me.ChooseWorkingFolder_Lbl.BackColor = System.Drawing.Color.Lavender
        Me.ChooseWorkingFolder_Lbl.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.ChooseWorkingFolder_Lbl.Font = New System.Drawing.Font("Wingdings", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(2, Byte))
        Me.ChooseWorkingFolder_Lbl.ForeColor = System.Drawing.Color.Black
        Me.ChooseWorkingFolder_Lbl.Location = New System.Drawing.Point(266, 38)
        Me.ChooseWorkingFolder_Lbl.Name = "ChooseWorkingFolder_Lbl"
        Me.ChooseWorkingFolder_Lbl.Size = New System.Drawing.Size(30, 21)
        Me.ChooseWorkingFolder_Lbl.TabIndex = 2
        Me.ChooseWorkingFolder_Lbl.Text = "1"
        Me.ChooseWorkingFolder_Lbl.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'CtWorkingFolder_TxtBox
        '
        Me.CtWorkingFolder_TxtBox.AccessibleDescription = "Main folder of app, where it copies input and produces output"
        Me.CtWorkingFolder_TxtBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CtWorkingFolder_TxtBox.Location = New System.Drawing.Point(6, 38)
        Me.CtWorkingFolder_TxtBox.Name = "CtWorkingFolder_TxtBox"
        Me.CtWorkingFolder_TxtBox.Size = New System.Drawing.Size(255, 21)
        Me.CtWorkingFolder_TxtBox.TabIndex = 1
        '
        'WorkingFolder_Lbl
        '
        Me.WorkingFolder_Lbl.AutoSize = True
        Me.WorkingFolder_Lbl.Location = New System.Drawing.Point(6, 22)
        Me.WorkingFolder_Lbl.Name = "WorkingFolder_Lbl"
        Me.WorkingFolder_Lbl.Size = New System.Drawing.Size(76, 13)
        Me.WorkingFolder_Lbl.TabIndex = 0
        Me.WorkingFolder_Lbl.Text = "Working folder"
        '
        'Timer1
        '
        Me.Timer1.Enabled = True
        Me.Timer1.Interval = 1000
        '
        'Form1
        '
        Me.AllowDrop = True
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(809, 196)
        Me.Controls.Add(Me.FilesList_Lview)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.ToggleOptions_Lbl)
        Me.Controls.Add(Me.Message_Lbl)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Controls.Add(Me.Process_Btn)
        Me.Controls.Add(Me.ClearList_Lbl)
        Me.Controls.Add(Me.ChooseZip_Btn)
        Me.Name = "Form1"
        Me.Text = "CL Converter"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents ChooseZip_Btn As System.Windows.Forms.Button
    Friend WithEvents FilesList_Lview As System.Windows.Forms.ListView
    Friend WithEvents ClearList_Lbl As System.Windows.Forms.Label
    Friend WithEvents Process_Btn As System.Windows.Forms.Button
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents Message_Lbl As System.Windows.Forms.Label

    Private Sub Form1_Activated(sender As Object, e As EventArgs) Handles Me.Activated

        Me.ProgressBar1.Minimum = 0
        Me.ProgressBar1.Maximum = 100

        Me.Message_Lbl.Text = "Ready"


    End Sub
    Friend WithEvents ToggleOptions_Lbl As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents WorkingFolderBrowserDialog As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents CtWorkingFolder_TxtBox As System.Windows.Forms.TextBox
    Friend WithEvents WorkingFolder_Lbl As System.Windows.Forms.Label
    Friend WithEvents ChooseWorkingFolder_Lbl As System.Windows.Forms.Label
    Friend WithEvents OpenOutputFolder_Cbox As System.Windows.Forms.CheckBox
    Friend WithEvents AutoClean_Folders_Cbox As System.Windows.Forms.CheckBox
    Friend WithEvents Timer1 As System.Windows.Forms.Timer


End Class
