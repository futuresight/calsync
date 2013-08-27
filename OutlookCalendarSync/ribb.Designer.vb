Partial Class ribb
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()> _
   Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'This call is required by the Component Designer.
        InitializeComponent()

    End Sub

    'Component overrides dispose to clean up the component list.
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

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Tab1 = Me.Factory.CreateRibbonTab
        Me.Group1 = Me.Factory.CreateRibbonGroup
        Me.syncB = Me.Factory.CreateRibbonButton
        Me.rawB = Me.Factory.CreateRibbonButton
        Me.settB = Me.Factory.CreateRibbonButton
        Me.Tab1.SuspendLayout()
        Me.Group1.SuspendLayout()
        '
        'Tab1
        '
        Me.Tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.Tab1.ControlId.OfficeId = "TabCalendar"
        Me.Tab1.Groups.Add(Me.Group1)
        Me.Tab1.Label = "TabCalendar"
        Me.Tab1.Name = "Tab1"
        '
        'Group1
        '
        Me.Group1.Items.Add(Me.syncB)
        Me.Group1.Items.Add(Me.rawB)
        Me.Group1.Items.Add(Me.settB)
        Me.Group1.Label = "Sync with Google"
        Me.Group1.Name = "Group1"
        '
        'syncB
        '
        Me.syncB.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.syncB.Image = Global.OutlookCalendarSync.My.Resources.Resources.syncIcon
        Me.syncB.Label = "Sync Now"
        Me.syncB.Name = "syncB"
        Me.syncB.ShowImage = True
        Me.syncB.SuperTip = "Synchronise now with Google"
        '
        'rawB
        '
        Me.rawB.Image = Global.OutlookCalendarSync.My.Resources.Resources.rawIcon
        Me.rawB.Label = "Force Resync"
        Me.rawB.Name = "rawB"
        Me.rawB.ShowImage = True
        Me.rawB.SuperTip = "Force a resync of all past calendar entries. Can be very slow."
        '
        'settB
        '
        Me.settB.Image = Global.OutlookCalendarSync.My.Resources.Resources.settingsIcon
        Me.settB.Label = "Settings"
        Me.settB.Name = "settB"
        Me.settB.ShowImage = True
        Me.settB.SuperTip = "Edit preferences for the Outlook Calendar Sync program"
        '
        'ribb
        '
        Me.Name = "ribb"
        Me.RibbonType = "Microsoft.Outlook.Explorer"
        Me.Tabs.Add(Me.Tab1)
        Me.Tab1.ResumeLayout(False)
        Me.Tab1.PerformLayout()
        Me.Group1.ResumeLayout(False)
        Me.Group1.PerformLayout()

    End Sub

    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents settB As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents syncB As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents rawB As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property ribb() As ribb
        Get
            Return Me.GetRibbon(Of ribb)()
        End Get
    End Property
End Class
