Imports System.Windows.Forms

Public Class FirstRun

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.Close()
        Globals.ThisAddIn.ConnectToGoogle()
    End Sub

    Private Sub importB_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles importB.LinkClicked
        ImportSyncfile()
    End Sub

    Private Sub FirstRun_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class
