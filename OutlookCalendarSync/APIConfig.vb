Public Class APIConfig
    '
    ' Before you begin, you must register this program under your API Console in your Google Developer Account.
    ' Go to https://code.google.com/apis/console/
    ' Create a new project and enable it in the APIs tab
    ' In the services tab, enable the following services: CalDAV API, Calendar API, Tasks API
    ' In the API Access tab, click "Create an OAuth 2.0 Client ID..."
    ' On the first screen enter something like "Outlook Calendar Sync Spinoff" for the name, and you don't need a logo
    ' On the second screen select "Installed application"
    ' Click "Create Client ID"


    ' APIKey should be the long alphanumeric jumble found under Simple API Access (in the API Access tab)
    Public Shared ReadOnly Property APIKey As String
        Get
            Return ""
        End Get
    End Property

    ' ClientID should be the multi-section ID found under Client ID for installed applications (in the API Access tab)
    ' It should be in the format 000000000.apps.googleusercontent.com
    Public Shared ReadOnly Property ClientID As String
        Get
            Return ""
        End Get
    End Property

    ' Secret should be the shorter alphanumeric jumble found under Client ID for installed applications (in the API Access tab)
    ' This value should be (as the name implies) kept secret, so make sure you don't publish this in an open source repository!
    Public Shared ReadOnly Property Secret As String
        Get
            Return ""
        End Get
    End Property

    ' The AuthStorageEncryptionKey is a string used to encrypt the authentication key given by Google upon connection of this program.
    ' You may use the default value below for testing purposes as well as everyday use, but you should not publish it.
    '
    '
    Public Const AuthStorageEncryptionKey As String = "CALENDARS_R_US_1337"

End Class
