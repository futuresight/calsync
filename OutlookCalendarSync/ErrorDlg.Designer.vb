<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ErrorDlg
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
        Me.cancelB = New System.Windows.Forms.Button()
        Me.copyB = New System.Windows.Forms.Button()
        Me.errorL = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'cancelB
        '
        Me.cancelB.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cancelB.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cancelB.Location = New System.Drawing.Point(356, 207)
        Me.cancelB.Name = "cancelB"
        Me.cancelB.Size = New System.Drawing.Size(67, 23)
        Me.cancelB.TabIndex = 1
        Me.cancelB.Text = "Close"
        '
        'copyB
        '
        Me.copyB.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.copyB.Location = New System.Drawing.Point(238, 207)
        Me.copyB.Name = "copyB"
        Me.copyB.Size = New System.Drawing.Size(112, 23)
        Me.copyB.TabIndex = 2
        Me.copyB.Text = "Copy to clipboard"
        Me.copyB.UseVisualStyleBackColor = True
        '
        'errorL
        '
        Me.errorL.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.errorL.Location = New System.Drawing.Point(12, 27)
        Me.errorL.Name = "errorL"
        Me.errorL.Size = New System.Drawing.Size(411, 177)
        Me.errorL.TabIndex = 3
        Me.errorL.Text = "Unspecified Error"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Maroon
        Me.Label1.Location = New System.Drawing.Point(12, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(295, 15)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Outlook Calendar Sync encountered an error!"
        '
        'ErrorDlg
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.cancelB
        Me.ClientSize = New System.Drawing.Size(435, 242)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.errorL)
        Me.Controls.Add(Me.copyB)
        Me.Controls.Add(Me.cancelB)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "ErrorDlg"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Outlook Calendar Sync - Error Message"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cancelB As System.Windows.Forms.Button
    Friend WithEvents copyB As System.Windows.Forms.Button
    Friend WithEvents errorL As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label

End Class
