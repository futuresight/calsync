Imports Microsoft.Office.Tools.Ribbon

Public Class ribb

    Private Sub ribb_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub rawB_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles rawB.Click
        ' force a complete and total sync, ignoring last update time

        ' warn that it might take a while depending on quantity of events in calendar
        Dim res As MsgBoxResult = MsgBox("Warning: performing a full sync means that the utility will attempt to synchronise all past and future events, regardless of when they were last updated. This may take a while depending on the quantity of events and calendars. Use this only if you need old events that are not being synchronised.", MsgBoxStyle.OkCancel, "Full sync")
        If res = MsgBoxResult.Ok Then
            Dim bw As Threading.Timer = New Threading.Timer(AddressOf Globals.ThisAddIn.SyncNowFull)
            bw.Change(0, Threading.Timeout.Infinite)
        End If
    End Sub

    Private Sub syncB_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles syncB.Click
        Dim bw As Threading.Timer = New Threading.Timer(AddressOf Globals.ThisAddIn.SyncNow)
        bw.Change(0, Threading.Timeout.Infinite)
    End Sub

    Private Sub settB_Click(ByVal sender As System.Object, ByVal e As Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs) Handles settB.Click
        Dim dlg As SettDlg = New SettDlg
        dlg.Show()
    End Sub
End Class
