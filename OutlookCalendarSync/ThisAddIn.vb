
Imports System.Collections.Generic

Imports DotNetOpenAuth.OAuth2

Imports Google.Apis.Authentication.OAuth2
Imports Google.Apis.Authentication.OAuth2.DotNetOpenAuth
Imports Google.Apis.Calendar.v3
Imports Google.Apis.Calendar.v3.Data
Imports Google.Apis.Calendar.v3.EventsResource
Imports Google.Apis.Samples.Helper
Imports Google.Apis.Services
Imports Google.Apis.Util

Public Class ThisAddIn

    '' Calendar scopes which is initialized on the main method
    Dim scopes As IList(Of String) = New List(Of String)()

    '' Calendar service
    Dim service As CalendarService

    '' List of calendars
    Dim list As IList(Of CalendarListEntry)
    WithEvents cd As AppDomain = AppDomain.CurrentDomain
    Dim localcache As List(Of Outlook.AppointmentItem) = New List(Of Outlook.AppointmentItem)

    '' Sync interval timer
    Public IntervalTimer As Timers.Timer = New Timers.Timer(20 * 60000)

    ' Fool the loader for assembly version errors
    Function LoadAsm(ByVal sender As Object, ByVal args As ResolveEventArgs) As Reflection.Assembly Handles cd.AssemblyResolve
        Dim asmName = New Reflection.AssemblyName(args.Name)
        Dim shortName = asmName.Name
        If {"System.Net.Http.Primitives"}.Contains(shortName) Then
            ' compare against list of assemblies we ignore revisions for
            ' check if this assembly is already loaded under a different version #
            Dim allAss As Reflection.Assembly() = AppDomain.CurrentDomain.GetAssemblies()
            Dim list As New List(Of Reflection.Assembly)(allAss)
            Dim loadedAssembly = list.Find(Function(ass) New Reflection.AssemblyName(ass.FullName).Name = shortName)
            ' check if we have any version loaded yet
            If loadedAssembly IsNot Nothing Then
                Return loadedAssembly
            Else
                ' assembly has not yet been loaded in this domain
                ' probe for assembly by name
                Dim probedAssembly As Reflection.Assembly = Reflection.Assembly.LoadFrom(String.Format("{0}.dll", shortName))

                ' probe for any match on assembly.dll 
                Return probedAssembly
            End If
        End If
        ' ignore binding failure -> pass up the stack
        Return Nothing
    End Function

    Sub ThisAddIn_Startup() Handles Me.Startup
        ' Set app name
            AuthorizationMgr.SetCustomAppName("FutureSight.OutlookCalendarSync")

            '// Check and load run settings
            If Not IO.Directory.Exists(AppData.SpecificPath) Then
                IO.Directory.CreateDirectory(AppData.SpecificPath)
            End If
            Dim settfile As String = "n20"
            If IO.File.Exists(AppData.SpecificPath & "\_CalendarSync.ini") Then
                settfile = IO.File.ReadAllText(AppData.SpecificPath & "\_CalendarSync.ini")
            Else ' Could be first run, do not schedule syncs and display setup dialog
                IO.File.WriteAllText(AppData.SpecificPath & "\_CalendarSync.ini", "n20")
                Dim dlg As FirstRun = New FirstRun
                dlg.Show()
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

            ' Set timer settings
            Dim interv As Integer = Math.Max(CInt(settfile.Substring(1)), 10)
            IntervalTimer.Interval = interv

            '// Connect to Google Calendar API if I haven't yet done so, and perform a first sync
            If settfile.StartsWith("y") Then
                Dim bw As Threading.Timer = New Threading.Timer(AddressOf DoAutosync)
                bw.Change(1000, Threading.Timeout.Infinite)
                IntervalTimer.Start()
            End If

    End Sub

    Sub ThisAddIn_Shutdown() Handles Me.Shutdown
        '// Perform last sync before closing
    End Sub

    Sub DoAutosync()
        ' check if sync is enabled
        Dim settfile As String = ""
        If IO.File.Exists(AppData.SpecificPath & "\_CalendarSync.ini") Then
            settfile = IO.File.ReadAllText(AppData.SpecificPath & "\_CalendarSync.ini")
        End If
        If settfile.Length < 2 Then
            'invalid file
            IO.File.WriteAllText(AppData.SpecificPath & "\_CalendarSync.ini", "y10")
            ErrorMsg("Invalid settings file. Default settings file created, visit settings to change back.", 16)
        End If
        If Not (settfile.StartsWith("n") Or settfile.StartsWith("y")) Then
            'invalid file
            IO.File.WriteAllText(AppData.SpecificPath & "\_CalendarSync.ini", "y10")
            ErrorMsg("Invalid settings file. Default settings file created, visit settings to change back.", 17)
        End If
        If Not IsNumeric(settfile.Substring(1)) Then
            'invalid file
            IO.File.WriteAllText(AppData.SpecificPath & "\_CalendarSync.ini", "y10")
            ErrorMsg("Invalid settings file. Default settings file created, visit settings to change back.", 18)
        End If
        settfile = IO.File.ReadAllText(AppData.SpecificPath & "\_CalendarSync.ini")

        If settfile.Substring(0, 1) = "y" Then
            ' autosync enabled, do sync
            SyncNow()
        End If
    End Sub

    Sub SyncNow()
        Globals.Ribbons.ribb.syncB.Image = My.Resources.syncingIcon
        Globals.Ribbons.ribb.syncB.Label = "Syncing"
        SyncCalendars()
        Globals.Ribbons.ribb.syncB.Image = My.Resources.syncIcon
        Globals.Ribbons.ribb.syncB.Label = "Sync Now"
    End Sub

    Sub SyncNowFull()
        Globals.Ribbons.ribb.syncB.Image = My.Resources.syncingIcon
            Globals.Ribbons.ribb.syncB.Label = "Syncing All"
        SyncCalendars(True)
        Globals.Ribbons.ribb.syncB.Image = My.Resources.syncIcon
        Globals.Ribbons.ribb.syncB.Label = "Sync Now"
    End Sub

    Sub ConnectToGoogle()
        ' Add the calendar specific scope to the scopes list
        scopes.Add(CalendarService.Scopes.Calendar.GetStringValue())

        ' Create the authenticator
        Dim apiKey As String = APIConfig.APIKey
        Dim clientID As String = APIConfig.ClientID
        Dim secret As String = APIConfig.Secret
        Dim credentials As FullClientCredentials = New FullClientCredentials With _
                                   {.ApiKey = apiKey, .ClientId = clientID, .ClientSecret = secret}
        Dim provider = New NativeApplicationClient(GoogleAuthenticationServer.Description)
        provider.ClientIdentifier = credentials.ClientId
        provider.ClientSecret = credentials.ClientSecret
        Dim auth As New OAuth2Authenticator(Of NativeApplicationClient)(provider, AddressOf GetAuthorization)

        ' Create the calendar service using an initializer instance
        Dim initializer As New BaseClientService.Initializer()
        initializer.Authenticator = auth
        service = New CalendarService(initializer)
    End Sub

    Function GetAuthorization(ByVal client As NativeApplicationClient) As IAuthorizationState
        ' You should use a more secure way of storing the key here as
        ' .NET applications can be disassembled using a reflection tool.
        Const STORAGE As String = "org.futuresight.google.calendarsync"
        Const KEY As String = APIConfig.AuthStorageEncryptionKey

        ' Check if there is a cached refresh token available.
        Dim state As IAuthorizationState = AuthorizationMgr.GetCachedRefreshToken(STORAGE, KEY)
        If Not state Is Nothing Then
            Try
                client.RefreshToken(state)
                Return state ' we are done
            Catch ex As DotNetOpenAuth.Messaging.ProtocolException
                MsgBox("Using an existing refresh token failed: " + ex.Message)
            End Try
        End If

        ' Otherwise, retrieve the authorization from the user
        state = AuthorizationMgr.RequestNativeAuthorization(client, scopes.ToArray())
        AuthorizationMgr.SetCachedRefreshToken(STORAGE, KEY, state)
        Return state
    End Function

    ''' <summary> Fetches a list of calendars from Google for later use and adds missing ones either way </summary>
    Sub SyncCalendars(Optional ByVal SyncAll As Boolean = False)
        ' Authenticate
        Try
            ConnectToGoogle()
        Catch ex As Exception
            MsgBox("err" & ex.Message)
        End Try

        ' Prepare exclude list
        Dim excludeIDs As List(Of String) = New List(Of String)
        If AppData.Exists("_ExcludeList.ini") Then
            For Each el As String In IO.File.ReadAllLines(AppData.SpecificPath & "\_ExcludeList.ini")
                excludeIDs.Add(el)
            Next
        End If

        ' Prepare birthday calendar variable
        Dim birthdayCal As CalendarListEntry

        ' Fetch the list of google calendars
        Try
            list = service.CalendarList.List.Execute.Items
        Catch ex As Exception
            ' failed?
            ErrorMsg("Fetching list of calendars failed: " & ex.Message, 1)
            Return
        End Try

        ' Fetch the list of local cals
        Dim app As Outlook.Application = Application
        If app.GetNamespace("MAPI").Folders.Count = 0 Then
            'no folders, exit app with error
            ErrorMsg("Could not find any folders in Outlook, please check that the program is set up correctly.", 2)
            Return
        End If
        ' Add default calendar to list of locals
        Dim localCals As List(Of Outlook.Folder) = New List(Of Outlook.Folder)
        Dim calFolder As Outlook.Folder = app.GetNamespace("MAPI").GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar)
        localCals.Add(calFolder)
        ' Look for extra local calendars (both in parent dir and in subdirs)
        Try
            For index = 1 To calFolder.Parent.Folders.Count '// Parent
                If calFolder.Parent.Folders.Item(index).DefaultItemType = Outlook.OlItemType.olAppointmentItem Then
                    ' should contain appointments, let's try it out
                    If Not calFolder.Parent.Folders.Item(index).EntryID = calFolder.EntryID Then
                        localCals.Add(calFolder.Parent.Folders.Item(index))
                    End If
                End If
            Next
            For index = 1 To calFolder.Folders.Count '// Subdirs
                If calFolder.Folders.Item(index).DefaultItemType = Outlook.OlItemType.olAppointmentItem Then
                    ' should contain appointments, let's try it out
                    localCals.Add(calFolder.Folders.Item(index))
                End If
            Next
        Catch ex As Exception
            ' Didn't go as expected, don't lose sleep over it
        End Try

        ' Compare and add missing ones that exist on server but not locally
        localCals.Remove(calFolder)

        For Each cal As CalendarListEntry In list
            If cal.Primary Then
                ' Tick off primary one but make sure it's registered
                AppData.WriteFile("GoogID_" & calFolder.EntryID, (New System.Text.UTF8Encoding).GetBytes(cal.Id))
            ElseIf cal.AccessRole <> "owner" Then
                ' don't need shared cals
                If IsNothing(birthdayCal) Then
                    If cal.Summary.Contains("birthdays") Then
                        birthdayCal = cal
                    End If
                End If
            Else
                ' Check if it exists
                Dim correctItem As Integer = -1
                For Each el As Outlook.Folder In localCals
                    Dim rawdata = AppData.ReadFile("GoogID_" & el.EntryID)
                    If Not IsNothing(rawdata) Then
                        If (New System.Text.UTF8Encoding).GetString(rawdata) = cal.Id Then
                            correctItem = localCals.IndexOf(el)
                        End If
                    End If

                Next
                If correctItem <> -1 Then
                    ' Exists, tick off
                    localCals.RemoveAt(correctItem)
                Else
                    ' Does not exist locally, create it
                    Dim newname As String = ""
                    Dim canuse As Boolean = False
                    Do Until canuse
                        canuse = True
                        For Each fold As Outlook.Folder In calFolder.Folders
                            If fold.Name = cal.Summary & newname Then
                                canuse = False
                                newname = newname & "+"
                            End If
                        Next
                    Loop
                    Dim newcal As Outlook.Folder = calFolder.Folders.Add(cal.Summary & newname, Outlook.OlDefaultFolders.olFolderCalendar)
                    AppData.WriteFile("GoogID_" & newcal.EntryID, (New System.Text.UTF8Encoding).GetBytes(cal.Id))
                End If
            End If
        Next

        ' Add calendars missing on server (if we didnt exlude them from sync)
        For Each cal As Outlook.Folder In localCals
            If Not cal.EntryID = calFolder.EntryID And Not excludeIDs.Contains(cal.EntryID) Then
                AppData.WriteFile("GoogID_" & cal.EntryID, (New System.Text.UTF8Encoding).GetBytes(service.Calendars.Insert(New Calendar With {.Summary = cal.Name}).Execute().Id))
            End If
        Next

        ' Repopulate list of local calendars after filtering for sync
        localCals = New List(Of Outlook.Folder)
        localCals.Add(calFolder)
        ' Look for extra local calendars (both in parent dir and in subdirs)
        Try
            For index = 1 To calFolder.Parent.Folders.Count '// Parent
                If calFolder.Parent.Folders.Item(index).DefaultItemType = Outlook.OlItemType.olAppointmentItem Then
                    ' should contain appointments, let's try it out
                    If Not calFolder.Parent.Folders.Item(index).EntryID = calFolder.EntryID Then
                        localCals.Add(calFolder.Parent.Folders.Item(index))
                    End If
                End If
            Next
            For index = 1 To calFolder.Folders.Count '// Subdirs
                If calFolder.Folders.Item(index).DefaultItemType = Outlook.OlItemType.olAppointmentItem Then
                    ' should contain appointments, let's try it out
                    localCals.Add(calFolder.Folders.Item(index))
                End If
            Next
        Catch ex As Exception
            ' Didn't go as expected, don't lose sleep over it
        End Try

        ' Repopulate online list after sync
        Try
            list = service.CalendarList.List.Execute.Items
        Catch ex As Exception
            ' failed?
            ErrorMsg("Fetching second list of calendars failed: " & ex.Message, 3)
            Return
        End Try

        ' Synchronise each calendar with the server
        For Each cal As Outlook.Folder In localCals
            If Not excludeIDs.Contains(cal.EntryID) Then
                SyncAppointments(cal, SyncAll)
            End If
        Next

        ' Try to get contact birthday calendar if possible
        If IO.File.Exists(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\Outlook Files\Outlook Today.htm") Then
            If Not IsNothing(birthdayCal) Then
                Dim ld As SortedDictionary(Of Date, String) = New SortedDictionary(Of Date, String)
                Dim listreq = service.Events.List(birthdayCal.Id)
                listreq.TimeMin = Return3339Datetime(Today.AddDays(-1))
                listreq.TimeMax = Return3339Datetime(Today.AddMonths(2))
                For Each appt As Data.Event In listreq.Execute.Items
                    ld.Add(appt.Start.Date, appt.Summary.Replace("'s birthday", ""))
                Next

                ' make a simple string array in HTML from dictionary
                Dim l As List(Of String) = New List(Of String)
                For Each el As KeyValuePair(Of Date, String) In ld
                    'l.Add("<tr><td>" & Globalization.CultureInfo.CurrentCulture.DateTimeFormat.GetDayName(el.Key.DayOfWeek) & ", " _
                    '      & el.Key.Day & " " & Globalization.CultureInfo.CurrentCulture.DateTimeFormat.GetDayName(el.Key.Month) & "</td><td>" & el.Value & "</tr>")
                    Dim thisString As String = ""
                    If el.Key.Subtract(Today) < New TimeSpan(6, 0, 0, 0) Then
                        thisString = "This "
                    End If
                    If el.Key.Subtract(Today) < New TimeSpan(3, 0, 0, 0) Then
                        l.Add("<tr style=""font-weight: bold; color: #800;""><td>" & el.Value & "</td><td>" & thisString & el.Key.ToString("ddd, d MMMM") & "</td></tr>")
                    Else
                        l.Add("<tr><td>" & el.Value & "</td><td>" & thisString & el.Key.ToString("ddd, d MMMM") & "</td><td>(in " & el.Key.Subtract(Today).Days & " days)</td></tr>")
                    End If
                Next

                'MsgBox(Join(l.ToArray, "<br />"))
                ' Add to HTML
                ReplaceOutlookTodayBirthdays(Join(l.ToArray, vbCrLf))
            End If
        End If

        '' Display all calendars
        'DisplayList(list)
        'If list.Count > 0 Then
        '    DisplayFirstCalendarEvents(list.First)
        'End If
    End Sub

    ''' <summary>Synchronise all appointments in a given calendar</summary>
    Sub SyncAppointments(ByVal Cal As Outlook.Folder, Optional ByVal SyncAll As Boolean = False)
        Dim syncTime As String = Return3339Datetime(DateTime.UtcNow.AddHours(-1))
        ' Get corresponding google cal entry
        Dim GCal As CalendarListEntry = Nothing
        If Cal.EntryID = Application.GetNamespace("MAPI").GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar).EntryID Then
            For Each el As CalendarListEntry In list
                If el.Primary Then
                    GCal = el
                    Exit For
                End If
            Next
        End If
        For Each el As CalendarListEntry In list
            Dim rawdata = AppData.ReadFile("GoogID_" & Cal.EntryID)
            If Not IsNothing(rawdata) Then
                If (New System.Text.UTF8Encoding).GetString(rawdata) = el.Id Then
                    GCal = el
                    Exit For
                End If
            End If
        Next

        If GCal Is Nothing Then
            ' Calendar does not exist, return error
            ErrorMsg("Appointment sync failed: Google Calendar for '" & Cal.Name & "' was not found", 4)
        End If

        ' Prepare variables
        Dim listreq = service.Events.List(GCal.Id)
        listreq.ShowDeleted = True
        Dim synctimeraw = AppData.ReadFile("LastSyncTime")
        Dim synctimestr As String = DateTime.MinValue.ToString
        If Not IsNothing(synctimeraw) And Not SyncAll Then
            listreq.UpdatedMin = (New System.Text.UTF8Encoding).GetString(synctimeraw)
            synctimestr = (New System.Text.UTF8Encoding).GetString(synctimeraw)
        End If

        ' populate future missing entries
        Dim missingEntries As List(Of Outlook.AppointmentItem) = New List(Of Outlook.AppointmentItem)
        For Each cld As Outlook.AppointmentItem In Cal.Items
            ' add to missing entries only if after last sync time
            If SyncAll Or cld.LastModificationTime.AddHours(1) > Convert.ToDateTime(synctimestr) Then
                missingEntries.Add(cld)
            End If
            Try : RemoveHandler cld.BeforeDelete, AddressOf HandleDeleteAppointment
            Catch ex As Exception : End Try
            AddHandler cld.BeforeDelete, AddressOf HandleDeleteAppointment
        Next

        ' Download missing entries (server -> local)

        For Each srv As Data.Event In listreq.Execute.Items
            ' Look for existing entry to update
            Dim foundLoc As Outlook.AppointmentItem = Nothing
            For Each loc As Outlook.AppointmentItem In missingEntries
                Dim rawdata = AppData.ReadFile("ApptID_" & loc.EntryID)
                If Not IsNothing(rawdata) Then
                    If (New System.Text.UTF8Encoding).GetString(rawdata) = srv.Id Then
                        foundLoc = loc
                        Exit For
                    End If
                End If
            Next
            If Not foundLoc Is Nothing Then ' // Found, update
                missingEntries.Remove(foundLoc)
                If srv.Status = "cancelled" Then
                    ' Deleted on server, delete here
                    Try : RemoveHandler foundLoc.BeforeDelete, AddressOf HandleDeleteAppointment
                    Catch ex As Exception : End Try
                    If AppData.Exists("ApptID_" & foundLoc.EntryID) Then
                        Try
                            IO.File.Delete(AppData.SpecificPath & "\ApptID_" & foundLoc.EntryID)
                        Catch ex As Exception
                            ErrorMsg("Could not delete old appt ID (nothing serious). " & ex.Message, 14)
                        End Try
                    End If
                    foundLoc.Delete()
                Else
                    ' Sync the two appointment entries
                    SyncAppointment(srv, foundLoc)
                End If
            Else ' // Not found, create if not deleted on server already
                If Not srv.Status = "cancelled" Then
                    Try
                        Dim newentry As Outlook.AppointmentItem = Application.CreateItem(Outlook.OlItemType.olAppointmentItem)
                        newentry.Subject = srv.Summary
                        newentry.Body = srv.Description
                        newentry.Location = srv.Location
                        If Not IsNothing(srv.Start.Date) Then
                            newentry.Start = Convert.ToDateTime(srv.Start.Date)
                            newentry.AllDayEvent = True
                        Else
                            newentry.Start = Convert.ToDateTime(srv.Start.DateTime)
                            If Not IsNothing(srv.Start.Date) Then
                                newentry.End = Convert.ToDateTime(srv.End.Date)
                            Else
                                newentry.End = Convert.ToDateTime(srv.End.DateTime)
                            End If
                        End If

                        ' Save to correct calendar (copy + delete, instead of move, to be safe)
                        newentry.Save()
                        Dim newentry2 As Outlook.AppointmentItem = newentry.CopyTo(Cal, Outlook.OlAppointmentCopyOptions.olCreateAppointment)
                        newentry.Delete()
                        RefreshAllHandlers(Cal)
                        ' Register into appdata
                        AppData.WriteFile("ApptID_" & newentry2.EntryID, (New System.Text.UTF8Encoding).GetBytes(srv.Id))
                    Catch ex As Exception
                        ErrorMsg("Could not create calendar: " & ex.Message, 5)
                    End Try
                End If
            End If
        Next

        ' Upload missing entries (local -> server)
        For Each entry As Outlook.AppointmentItem In missingEntries
            If Not AppData.Exists("ApptID_" & entry.EntryID) Then
                ' Prepare insert request
                Dim calId As String = ""
                Try
                    Dim rawdata = AppData.ReadFile("GoogID_" & Cal.EntryID)
                    If Not IsNothing(rawdata) Then
                        calId = (New System.Text.UTF8Encoding).GetString(rawdata)
                    End If
                Catch ex As Exception
                    ErrorMsg("calId could not be resolved: " & ex.Message, 10)
                    Return
                End Try
                Dim newev As Data.Event = New Data.Event
                newev.Summary = entry.Subject
                newev.Description = entry.Body
                newev.Location = entry.Location
                If entry.AllDayEvent Then
                    newev.Start = New Data.EventDateTime With {.Date = Return3339Date(entry.StartInStartTimeZone)}
                    newev.End = New Data.EventDateTime With {.Date = Return3339Date(entry.StartInStartTimeZone.AddDays(1))}
                Else
                    newev.Start = New Data.EventDateTime With {.DateTime = Return3339Datetime(entry.StartUTC)}
                    newev.End = New Data.EventDateTime With {.DateTime = Return3339Datetime(entry.EndUTC)}
                End If
                Dim createdev = service.Events.Insert(newev, calId).Execute()
                ' Register into appdata
                AppData.WriteFile("ApptID_" & entry.EntryID, (New System.Text.UTF8Encoding).GetBytes(createdev.Id))
            End If
        Next

        ' update last sync time
        AppData.WriteFile("LastSyncTime", (New System.Text.UTF8Encoding).GetBytes(syncTime))
    End Sub

    ''' <summary>Synchronise a single appointment</summary>
    Sub SyncAppointment(ByVal ServerCopy As Data.Event, ByVal LocalCopy As Outlook.AppointmentItem)
        ' First determine which way to sync
        Dim servNewer As Boolean = False
        Try
            If Convert.ToDateTime(ServerCopy.Updated) > LocalCopy.LastModificationTime Then
                ' server -> local
                servNewer = True
                ' Else :: local -> server
            End If
        Catch ex As Exception
            ErrorMsg("Could not sync existing appointment '" & LocalCopy.Subject & "' (get mod time): " & ex.Message, 11)
        End Try

        ' Sync data
        If servNewer Then
            ' write to local
            Try
                LocalCopy.Subject = ServerCopy.Summary
                LocalCopy.Body = ServerCopy.Description
                LocalCopy.Location = ServerCopy.Location
                If Not IsNothing(ServerCopy.Start.Date) Then
                    LocalCopy.Start = Convert.ToDateTime(ServerCopy.Start.Date)
                    LocalCopy.AllDayEvent = True
                Else
                    LocalCopy.Start = Convert.ToDateTime(ServerCopy.Start.DateTime)
                    If Not IsNothing(ServerCopy.Start.Date) Then
                        LocalCopy.End = Convert.ToDateTime(ServerCopy.End.Date)
                    Else
                        LocalCopy.End = Convert.ToDateTime(ServerCopy.End.DateTime)
                    End If
                End If
            Catch ex As Exception
                ErrorMsg("Could not sync '" & LocalCopy.Subject & "' (inbound): " & ex.Message, 12)
                Return
            End Try
            LocalCopy.Save()
        Else
            ' check for equality first to avoid unnecessary server queries
            Dim copiesMatch As Boolean = True
            If Not LocalCopy.Subject = ServerCopy.Summary Then
                copiesMatch = False
            End If
            If Not LocalCopy.Body = ServerCopy.Description Then
                copiesMatch = False
            End If
            If Not LocalCopy.Location = ServerCopy.Location Then
                copiesMatch = False
            End If
            If LocalCopy.AllDayEvent Then
                If IsNothing(ServerCopy.Start.Date) Then
                    copiesMatch = False
                ElseIf LocalCopy.Start <> Convert.ToDateTime(ServerCopy.Start.Date) Then
                    'MsgBox(LocalCopy.StartUTC.ToString & vbCrLf & ServerCopy.Start.Date)
                    copiesMatch = False
                End If
            Else
                If IsNothing(ServerCopy.Start.DateTime) Then
                    copiesMatch = False
                ElseIf LocalCopy.Start <> Convert.ToDateTime(ServerCopy.Start.DateTime) Then
                    copiesMatch = False
                End If
            End If
            ' create and upload new copy if they don't match
            If Not copiesMatch Then
                ServerCopy.Summary = LocalCopy.Subject
                ServerCopy.Description = LocalCopy.Body
                ServerCopy.Location = LocalCopy.Location
                If LocalCopy.AllDayEvent Then
                    ServerCopy.Start = New Data.EventDateTime With {.Date = Return3339Date(LocalCopy.StartInStartTimeZone)}
                    ServerCopy.End = New Data.EventDateTime With {.Date = Return3339Date(LocalCopy.StartInStartTimeZone.AddDays(1))}
                Else
                    ServerCopy.Start = New Data.EventDateTime With {.DateTime = Return3339Datetime(LocalCopy.StartUTC)}
                    ServerCopy.End = New Data.EventDateTime With {.DateTime = Return3339Datetime(LocalCopy.EndUTC)}
                End If
                ' push request to server
                Dim calId As String = ""
                Try
                    Dim rawdata = AppData.ReadFile("GoogID_" & LocalCopy.Parent.EntryID)
                    If Not IsNothing(rawdata) Then
                        calId = (New System.Text.UTF8Encoding).GetString(rawdata)
                    End If
                Catch ex As Exception
                    ErrorMsg("calId parent could not be resolved for '" & LocalCopy.Subject & "': " & ex.Message, 14)
                    Return
                End Try
                If calId = "" Then
                    ErrorMsg("Could not sync '" & LocalCopy.Subject & "' (outbound): could not resolve calId", 15)
                    Return
                End If
                If IsNothing(service.Events.Update(ServerCopy, calId, ServerCopy.Id).Execute()) Then
                    ErrorMsg("Could not sync '" & LocalCopy.Subject & "' (outbound): empty result", 13)
                End If
            End If
        End If
    End Sub

    Sub RefreshAllHandlers(ByVal Cal As Outlook.Folder)
        For Each cld As Outlook.AppointmentItem In Cal.Items
            Try : RemoveHandler cld.BeforeDelete, AddressOf HandleDeleteAppointment
            Catch ex As Exception : End Try
            AddHandler cld.BeforeDelete, AddressOf HandleDeleteAppointment
            localcache.Add(cld)
        Next
    End Sub

    ''' <summary>Handle the delete event of an outlook appointment</summary>
    Sub HandleDeleteAppointment(ByVal ItemObj As Object, ByRef Cancel As Boolean)
        Dim Item As Outlook.AppointmentItem = Nothing
        Try
            Item = ItemObj
        Catch ex As Exception
            ErrorMsg("Could not handle appointment deletion (cast error): " & ex.Message, 6)
            Return
        End Try
        ' first of all remove handler
        Try : RemoveHandler Item.BeforeDelete, AddressOf HandleDeleteAppointment
        Catch ex As Exception : End Try

        ' prepare Cal ID for deletion
        Dim calId As String = ""
        Try
            Dim rawdata = AppData.ReadFile("GoogID_" & Item.Parent.EntryID)
            If Not IsNothing(rawdata) Then
                calId = (New System.Text.UTF8Encoding).GetString(rawdata)
            End If
        Catch ex As Exception
            ErrorMsg("calId could not be resolved: " & ex.Message, 6)
            Return
        End Try

        ' prepare Event ID for deletion
        Dim evId As String = ""
        Try
            Dim rawdata = AppData.ReadFile("ApptID_" & Item.EntryID)
            If Not IsNothing(rawdata) Then
                evId = (New System.Text.UTF8Encoding).GetString(rawdata)
            End If
        Catch ex As Exception
            ErrorMsg("eventId could not be resolved: " & ex.Message, 7)
            Return
        End Try

        ' Prepare exclude list
        Dim excludeIDs As List(Of String) = New List(Of String)
        If AppData.Exists("_ExcludeList.ini") Then
            For Each el As String In IO.File.ReadAllLines(AppData.SpecificPath & "\_ExcludeList.ini")
                excludeIDs.Add(el)
            Next
        End If

        ' request deletion from server (if we didnt exlude it from sync)
        If excludeIDs.Contains(calId) Then
            calId = ""
        End If
        If calId <> "" And evId <> "" Then
            Try
                service.Events.Delete(calId, evId).Execute()
            Catch ex As Exception
                ErrorMsg("Failed to delete event '" & Item.Subject & "' from server: " & ex.Message, 8)
                Return
            End Try
        End If

        ' if everything went well, delete appdata entry
        If AppData.Exists("ApptID_" & Item.EntryID) Then
            IO.File.Delete(AppData.SpecificPath & "\ApptID_" & Item.EntryID)
        End If

        ' try to remove from cache
        If localcache.Contains(Item) Then
            localcache.Remove(Item)
        End If
    End Sub

    ''' <summary> Displays all calendars.</summary>
    Sub DisplayList(ByVal list As IList(Of CalendarListEntry))
        'calendarsL.Items.Clear()
        For Each item As CalendarListEntry In list
            'calendarsL.Items.Add(item.Summary & vbCrLf & "Location: " & item.Location & ", TimeZone: " & item.TimeZone)
            MsgBox(item.Summary)
        Next
    End Sub

    ''' <summary>Displays the calendar's events.</summary>
    Sub DisplayFirstCalendarEvents(ByVal list As CalendarListEntry)
        'fromT.Text = list.Summary
        Dim requeust As ListRequest = service.Events.List(list.Id)
        ' Set MaxResults and TimeMin with sample values
        requeust.MaxResults = 10
        requeust.TimeMin = "2012-01-01T00:00:00-00:00"
        ' Fetch the list of events
        'eventsL.Items.Clear()
        For Each calendarEvent As Data.Event In requeust.Execute().Items
            Dim startDate As String = "Unspecified"
            If (Not calendarEvent.Start Is Nothing) Then
                If (Not calendarEvent.Start.Date Is Nothing) Then
                    startDate = calendarEvent.Start.Date.ToString()
                ElseIf (Not calendarEvent.Start.DateTime Is Nothing) Then
                    startDate = calendarEvent.Start.DateTime.ToString()
                End If
            ElseIf (Not calendarEvent.End Is Nothing) Then
                If (Not calendarEvent.End.Date Is Nothing) Then
                    startDate = calendarEvent.End.Date.ToString()
                End If
            End If

            'eventsL.Items.Add(calendarEvent.Summary & ", Start at: " & startDate)
        Next
    End Sub

End Class
