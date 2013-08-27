Imports Ionic

Module Misc

    Sub ErrorMsg(ByVal Msg As String, ByVal Code As Integer)
        'MsgBox("0x0" & Code & ": " & Msg, MsgBoxStyle.Critical, "Calendar Sync Error")
        Dim dlg As ErrorDlg = New ErrorDlg
        dlg.errorL.Text = "0x" & Code.ToString("D4") & ":" & vbCrLf & Msg
        dlg.ShowDialog()
    End Sub

    Function GetMyFolder(ByVal FolderPath) As Outlook.Folder
        ' folder path needs to be something like 
        '   "Public Folders\All Public Folders\Company\Sales"
        Dim aFolders
        Dim fldr
        Dim i
        Dim objNS

        On Error Resume Next
        Dim strFolderPath As String = Replace(FolderPath, "/", "\")
        aFolders = Split(FolderPath, "\")

        'get the Outlook objects
        ' use intrinsic Application object in form script
        objNS = Globals.ThisAddIn.Application.GetNamespace("MAPI")

        'set the root folder
        fldr = objNS.Folders(aFolders(0))

        'loop through the array to get the subfolder
        'loop is skipped when there is only one element in the array
        For i = 1 To UBound(aFolders)
            fldr = fldr.Folders(aFolders(i))
            'check for errors
            'If Err() &lt;&gt; 0 Then Exit Function
        Next
        GetMyFolder = fldr

        ' dereference objects
        objNS = Nothing
    End Function

    Function DiscoverCalendars() As List(Of Outlook.Folder)
        DiscoverCalendars = New List(Of Outlook.Folder)
        ' Fetch the list of local cals
        Dim app As Outlook.Application = Globals.ThisAddIn.Application
        If app.GetNamespace("MAPI").Folders.Count = 0 Then
            'no folders, exit app with error
            ErrorMsg("Could not find any folders in Outlook, please check that the program is set up correctly.", 2)
            Exit Function
        End If
        ' Add default calendar to list of locals
        Dim calFolder As Outlook.Folder = app.GetNamespace("MAPI").GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar)
        DiscoverCalendars.Add(calFolder)
        ' Look for extra local calendars (both in parent dir and in subdirs)
        Try
            For index = 1 To calFolder.Parent.Folders.Count '// Parent
                If calFolder.Parent.Folders.Item(index).DefaultItemType = Outlook.OlItemType.olAppointmentItem Then
                    ' should contain appointments, let's try it out
                    If Not calFolder.Parent.Folders.Item(index).EntryID = calFolder.EntryID Then
                        DiscoverCalendars.Add(calFolder.Parent.Folders.Item(index))
                    End If
                End If
            Next
            For index = 1 To calFolder.Folders.Count '// Subdirs
                If calFolder.Folders.Item(index).DefaultItemType = Outlook.OlItemType.olAppointmentItem Then
                    ' should contain appointments, let's try it out
                    DiscoverCalendars.Add(calFolder.Folders.Item(index))
                End If
            Next
        Catch ex As Exception
            ' Didn't go as expected, don't lose sleep over it
        End Try
    End Function

    Sub ImportSyncfile()
        MsgBox("Please find your ZIP file with the backed up contents of the program's App Data, and unzip it to User_folder\AppData\Roaming\Google.Apis.Samples.Helper with no extra directories in the way (for example, the following path should be valid: User_folder\AppData\Roaming\Google.Apis.Samples.Helper\_CalendarSync.ini).", MsgBoxStyle.Information, "Import Sync Data")
    End Sub

    Sub ExportSyncfile()
        ' Pack all sync files in a zip (including exclude list) and save it
        'Dim dlg As Windows.Forms.SaveFileDialog = New Windows.Forms.SaveFileDialog
        'dlg.Title = "Export sync data"
        'dlg.Filter = "Sync file (.sync)|*.sync"
        'dlg.DefaultExt = ".sync"
        'If dlg.ShowDialog = Windows.Forms.DialogResult.OK Then
        '    Try

        '    Catch ex As Exception

        '    End Try
        'End If

        MsgBox("The current version of Outlook Calendar Sync does not support automatic backup due to limitations with Microsoft Office Plugins. To manually back up your data, you must zip the contents of the following directory and keep them in a safe place:" & vbCrLf & "User_folder\AppData\Roaming\Google.Apis.Samples.Helper" & vbCrLf & "Future versions will include automatic backup. Sorry for the inconvenience.", MsgBoxStyle.Information, "Export Sync Data")
    End Sub

    Function Return3339Date(ByVal inDate As Date) As String
        Return inDate.Year & "-" & inDate.Month.ToString("D2") & "-" & inDate.Day.ToString("D2")
    End Function

    Function Return3339Datetime(ByVal inDate As Date) As String
        Return inDate.Year & "-" & inDate.Month.ToString("D2") & "-" & inDate.Day.ToString("D2") _
            & "T" & inDate.Hour.ToString("D2") & ":" & inDate.Minute.ToString("D2") & ":" & inDate.Second.ToString("D2") _
            & "Z" '& offsetH.ToString("D2") & ":" & offsetM.ToString("D2")
    End Function
    ', ByVal offsetH As Integer, ByVal offsetM As Integer

    Sub ReplaceOutlookTodayBirthdays(NewString As String)
        Try
            Dim filepath As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\Outlook Files\Outlook Today.htm"
            Dim res As String = IO.File.ReadAllText(filepath)
            If Not res.Contains("<table width=""380px"" id=""BIRTHDAYS"">") Or Not res.Contains("</table>") Then
                Return
            End If
            Dim startI As Integer = res.IndexOf("<table width=""380px"" id=""BIRTHDAYS"">") + 36
            Dim endI As Integer = res.IndexOf("</table>", startI)
            res = res.Substring(0, startI) & vbCrLf & NewString & res.Substring(endI)
            IO.File.WriteAllText(filepath, res)
        Catch ex As Exception
            ErrorMsg("Failed to update upcoming birthdays; " & ex.Message, 20)
        End Try
    End Sub
End Module
