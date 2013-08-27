<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SettDlg
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(SettDlg))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.cleardataB = New System.Windows.Forms.Button()
        Me.importB = New System.Windows.Forms.Button()
        Me.exportB = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.enableB = New System.Windows.Forms.CheckBox()
        Me.excludeL = New System.Windows.Forms.CheckedListBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.syncintervalN = New System.Windows.Forms.NumericUpDown()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.saveB = New System.Windows.Forms.Button()
        Me.cancelB = New System.Windows.Forms.Button()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.clrAcctB = New System.Windows.Forms.Button()
        Me.aboutB = New System.Windows.Forms.Button()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        CType(Me.syncintervalN, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox3.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.cleardataB)
        Me.GroupBox1.Controls.Add(Me.importB)
        Me.GroupBox1.Controls.Add(Me.exportB)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(12, 166)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(308, 146)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Export Sync Data"
        '
        'cleardataB
        '
        Me.cleardataB.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cleardataB.Location = New System.Drawing.Point(205, 108)
        Me.cleardataB.Name = "cleardataB"
        Me.cleardataB.Size = New System.Drawing.Size(97, 23)
        Me.cleardataB.TabIndex = 4
        Me.cleardataB.Text = "Clear sync data"
        Me.cleardataB.UseVisualStyleBackColor = True
        '
        'importB
        '
        Me.importB.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.importB.Location = New System.Drawing.Point(102, 108)
        Me.importB.Name = "importB"
        Me.importB.Size = New System.Drawing.Size(97, 23)
        Me.importB.TabIndex = 3
        Me.importB.Text = "Import from file"
        Me.importB.UseVisualStyleBackColor = True
        '
        'exportB
        '
        Me.exportB.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.exportB.Location = New System.Drawing.Point(9, 108)
        Me.exportB.Name = "exportB"
        Me.exportB.Size = New System.Drawing.Size(87, 23)
        Me.exportB.TabIndex = 2
        Me.exportB.Text = "Export to file"
        Me.exportB.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(6, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(296, 89)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = resources.GetString("Label1.Text")
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.enableB)
        Me.GroupBox2.Controls.Add(Me.excludeL)
        Me.GroupBox2.Controls.Add(Me.Label2)
        Me.GroupBox2.Controls.Add(Me.syncintervalN)
        Me.GroupBox2.Controls.Add(Me.Label3)
        Me.GroupBox2.Controls.Add(Me.Label4)
        Me.GroupBox2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox2.Location = New System.Drawing.Point(12, 37)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(411, 123)
        Me.GroupBox2.TabIndex = 4
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Sync Options"
        '
        'enableB
        '
        Me.enableB.AutoSize = True
        Me.enableB.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.enableB.Location = New System.Drawing.Point(9, 15)
        Me.enableB.Name = "enableB"
        Me.enableB.Size = New System.Drawing.Size(133, 17)
        Me.enableB.TabIndex = 7
        Me.enableB.Text = "Enable automatic sync"
        Me.enableB.UseVisualStyleBackColor = True
        '
        'excludeL
        '
        Me.excludeL.CheckOnClick = True
        Me.excludeL.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.excludeL.FormattingEnabled = True
        Me.excludeL.Location = New System.Drawing.Point(6, 53)
        Me.excludeL.Name = "excludeL"
        Me.excludeL.ScrollAlwaysVisible = True
        Me.excludeL.Size = New System.Drawing.Size(399, 64)
        Me.excludeL.TabIndex = 10
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(6, 39)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(93, 13)
        Me.Label2.TabIndex = 9
        Me.Label2.Text = "Exclude from sync"
        '
        'syncintervalN
        '
        Me.syncintervalN.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.syncintervalN.Location = New System.Drawing.Point(301, 14)
        Me.syncintervalN.Maximum = New Decimal(New Integer() {800, 0, 0, 0})
        Me.syncintervalN.Minimum = New Decimal(New Integer() {10, 0, 0, 0})
        Me.syncintervalN.Name = "syncintervalN"
        Me.syncintervalN.Size = New System.Drawing.Size(46, 20)
        Me.syncintervalN.TabIndex = 7
        Me.syncintervalN.Value = New Decimal(New Integer() {10, 0, 0, 0})
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(202, 16)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(97, 13)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Perform sync every"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(351, 16)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(43, 13)
        Me.Label4.TabIndex = 6
        Me.Label4.Text = "minutes"
        '
        'saveB
        '
        Me.saveB.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.saveB.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.saveB.Location = New System.Drawing.Point(348, 288)
        Me.saveB.Name = "saveB"
        Me.saveB.Size = New System.Drawing.Size(81, 27)
        Me.saveB.TabIndex = 0
        Me.saveB.Text = "Save"
        Me.saveB.UseVisualStyleBackColor = True
        '
        'cancelB
        '
        Me.cancelB.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cancelB.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cancelB.Location = New System.Drawing.Point(348, 255)
        Me.cancelB.Name = "cancelB"
        Me.cancelB.Size = New System.Drawing.Size(81, 27)
        Me.cancelB.TabIndex = 6
        Me.cancelB.Text = "Cancel"
        Me.cancelB.UseVisualStyleBackColor = True
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.clrAcctB)
        Me.GroupBox3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox3.Location = New System.Drawing.Point(326, 174)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(103, 66)
        Me.GroupBox3.TabIndex = 5
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Clear Account Data"
        '
        'clrAcctB
        '
        Me.clrAcctB.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.clrAcctB.Location = New System.Drawing.Point(6, 35)
        Me.clrAcctB.Name = "clrAcctB"
        Me.clrAcctB.Size = New System.Drawing.Size(91, 23)
        Me.clrAcctB.TabIndex = 4
        Me.clrAcctB.Text = "Clear"
        Me.clrAcctB.UseVisualStyleBackColor = True
        '
        'aboutB
        '
        Me.aboutB.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.aboutB.Image = Global.OutlookCalendarSync.My.Resources.Resources.infoIcon
        Me.aboutB.Location = New System.Drawing.Point(270, 8)
        Me.aboutB.Name = "aboutB"
        Me.aboutB.Size = New System.Drawing.Size(153, 23)
        Me.aboutB.TabIndex = 7
        Me.aboutB.Text = "About the software..."
        Me.aboutB.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.aboutB.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.aboutB.UseVisualStyleBackColor = True
        '
        'SettDlg
        '
        Me.AcceptButton = Me.saveB
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(435, 324)
        Me.ControlBox = False
        Me.Controls.Add(Me.aboutB)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.cancelB)
        Me.Controls.Add(Me.saveB)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "SettDlg"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Outlook Calendar Sync Settings"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        CType(Me.syncintervalN, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox3.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents importB As System.Windows.Forms.Button
    Friend WithEvents exportB As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cleardataB As System.Windows.Forms.Button
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents syncintervalN As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents excludeL As System.Windows.Forms.CheckedListBox
    Friend WithEvents saveB As System.Windows.Forms.Button
    Friend WithEvents cancelB As System.Windows.Forms.Button
    Friend WithEvents enableB As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents clrAcctB As System.Windows.Forms.Button
    Friend WithEvents aboutB As System.Windows.Forms.Button

End Class
