Imports System.Windows.Forms
Imports Google.Apis.Samples.Helper

Public Class SettDlg
    Dim SettingInit As Boolean = False
    Friend CalList As List(Of String) = New List(Of String)

    Private Sub SettDlg_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ' populate exclude list with correct values
        ' set checkbox and numeric values to current setting

        Dim settfile As String = ""
        If IO.File.Exists(AppData.SpecificPath & "\_CalendarSync.ini") Then
            settfile = IO.File.ReadAllText(AppData.SpecificPath & "\_CalendarSync.ini")
        Else ' Could be first run, do not schedule syncs and display setup dialog
            IO.File.WriteAllText(AppData.SpecificPath & "\_CalendarSync.ini", "n20")
            Dim dlg As SettDlg = New SettDlg
            dlg.Show()
            Return
        End If
        If settfile.Length < 2 Then
            'invalid file
            IO.File.WriteAllText(AppData.SpecificPath & "\_CalendarSync.ini", "n20")
            ErrorMsg("Invalid settings file. Default settings file created, visit settings to change back.", 16)
        End If
        If Not (settfile.StartsWith("n") Or settfile.StartsWith("y")) Then
            'invalid file
            IO.File.WriteAllText(AppData.SpecificPath & "\_CalendarSync.ini", "n20")
            ErrorMsg("Invalid settings file. Default settings file created, visit settings to change back.", 17)
        End If
        If Not IsNumeric(settfile.Substring(1)) Then
            'invalid file
            IO.File.WriteAllText(AppData.SpecificPath & "\_CalendarSync.ini", "n20")
            ErrorMsg("Invalid settings file. Default settings file created, visit settings to change back.", 18)
        End If
        settfile = IO.File.ReadAllText(AppData.SpecificPath & "\_CalendarSync.ini")


        enableB.Checked = settfile.StartsWith("y") 'enableB
        syncintervalN.Value = Math.Min(CInt(settfile.Substring(1)), 800) 'syncintervalB
        Dim excludeStrings As List(Of String) = New List(Of String)
        If AppData.Exists("_ExcludeList.ini") Then
            For Each el As String In IO.File.ReadAllLines(AppData.SpecificPath & "\_ExcludeList.ini")
                excludeStrings.Add(el)
            Next
        End If
        excludeL.Items.Clear()
        CalList.Clear()
        For Each el As Outlook.Folder In DiscoverCalendars()
            excludeL.Items.Add(el.Name & " (ID" & el.EntryID & ")", excludeStrings.Contains(el.EntryID))
            CalList.Add(el.EntryID)
        Next
    End Sub

    Private Sub saveB_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles saveB.Click
        ' process values, save settings, close

        ' write settings file
        Dim setdata As String = ""
        If enableB.Checked Then
            setdata = "y"
        Else
            setdata = "n"
        End If
        setdata &= syncintervalN.Value.ToString
        Try
            IO.File.WriteAllText(AppData.SpecificPath & "\_CalendarSync.ini", setdata)
        Catch ex As Exception
            ErrorMsg("Could not save settings (IO err): " & ex.Message, 17)
            Return
        End Try

        'write exclude list
        Try
            Dim resstring As String = ""
            For Each el As String In excludeL.Items
                If excludeL.GetItemChecked(excludeL.Items.IndexOf(el)) Then
                    If resstring = "" Then
                        resstring = CalList.Item(excludeL.Items.IndexOf(el))
                    Else
                        resstring &= vbCrLf & CalList.Item(excludeL.Items.IndexOf(el))
                    End If
                End If
            Next
            IO.File.WriteAllText(AppData.SpecificPath & "\_ExcludeList.ini", resstring)
        Catch ex As Exception
            ErrorMsg("Failed to save settings, could not write exclude list. " & ex.Message, 18)
        End Try

        ' close the UI
        Close()

        ' update timer
        Globals.ThisAddIn.IntervalTimer.Interval = syncintervalN.Value
        If enableB.Checked Then
            Globals.ThisAddIn.IntervalTimer.Enabled = True
        Else
            Globals.ThisAddIn.IntervalTimer.Enabled = False
        End If
    End Sub

    Private Sub cancelB_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cancelB.Click
        Close()
    End Sub

    Private Sub exportB_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles exportB.Click
        ExportSyncfile()
    End Sub

    Private Sub importB_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles importB.Click
        ImportSyncfile()
    End Sub

    Private Sub cleardataB_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cleardataB.Click
        'warn first, then delete all data
        If MsgBox("Warning: deleting synchronisation data will probably cause terrible side effects to the synchronisation of your calendars. Only use this option if you need to make a clean start or to use a different account.", MsgBoxStyle.Exclamation + MsgBoxStyle.OkCancel, "Clear all data?") = MsgBoxResult.Ok Then
            ' loop through appdata and delete all sync files (keep setting files)
            Dim delCount As Integer = 0
            For Each el As String In IO.Directory.GetFiles(AppData.SpecificPath)
                If el.StartsWith("GoogID") Or el.StartsWith("ApptID") Or el = "org.futuresight.google.calendarsync.auth" Or el = "LastSyncTime" Then
                    Try
                        IO.File.Delete(AppData.SpecificPath & "\" & el)
                        delCount += 1
                    Catch ex As Exception
                        ErrorMsg("Error clearing data member '" & el & "': " & ex.Message, 15)
                    End Try
                End If
            Next
            MsgBox("Successfully deleted " & delCount & " sync items. As a preventive action, auto sync has been disabled. You may reenable it in settings.", MsgBoxStyle.Information, "Data cleared")
        End If
    End Sub

    Private Sub clrAcctB_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles clrAcctB.Click
        If Not MsgBox("Clearing account cache will require signing in again via Google to authorise connection. This is useful if you wish to change account to sync, but should not be done lightly.", MsgBoxStyle.OkCancel, "Clear account data") = MsgBoxResult.Ok Then
            Return
        End If
        If AppData.Exists("org.futuresight.google.calendarsync.auth") Then
            Try
                IO.File.Delete(AppData.SpecificPath & "\org.futuresight.google.calendarsync.auth")
            Catch ex As Exception
                ErrorMsg("Could not clear account data: IO error. " & ex.Message, 19)
            End Try
        End If
    End Sub

    Private Sub aboutB_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles aboutB.Click
        Dim dlg As AboutDlg = New AboutDlg
        dlg.ShowDialog()
    End Sub
End Class
