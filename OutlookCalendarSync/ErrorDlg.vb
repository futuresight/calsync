Imports System.Windows.Forms

Public Class ErrorDlg

    Private Sub closeB_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cancelB.Click
        Me.Close()
    End Sub

    Private Sub copyB_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles copyB.Click
        My.Computer.Clipboard.SetText(errorL.Text)
    End Sub
End Class
