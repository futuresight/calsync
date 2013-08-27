Imports System.Windows.Forms

Public Class AboutDlg

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.Close()
    End Sub

    Private Sub licenceB_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles licenceB.Click
        Dim dlg As LicenseDlg = New LicenseDlg
        dlg.ShowDialog()
    End Sub

    Private Sub siteB_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles siteB.Click
        Diagnostics.Process.Start("http://futuresight.org/")
    End Sub

    Private Sub AboutDlg_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ' load version number
        versionL.Text = "You are running FutureSight Cal Sync version " & My.Application.Info.Version.ToString
    End Sub
End Class
