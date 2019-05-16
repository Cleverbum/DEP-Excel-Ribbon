Partial Class Ribbon1
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'This call is required by the Component Designer.
        InitializeComponent()

    End Sub

    'Component overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Ribbon1))
        Me.Tab1 = Me.Factory.CreateRibbonTab
        Me.Group1 = Me.Factory.CreateRibbonGroup
        Me.Group3 = Me.Factory.CreateRibbonGroup
        Me.Group4 = Me.Factory.CreateRibbonGroup
        Me.Group2 = Me.Factory.CreateRibbonGroup
        Me.ChkDebug = Me.Factory.CreateRibbonCheckBox
        Me.CreateNew = Me.Factory.CreateRibbonButton
        Me.TDOnly = Me.Factory.CreateRibbonButton
        Me.BtnWCOnly = Me.Factory.CreateRibbonButton
        Me.CloseStale = Me.Factory.CreateRibbonButton
        Me.BtnCheckRegistrations = Me.Factory.CreateRibbonButton
        Me.WriteMails = Me.Factory.CreateRibbonButton
        Me.FindIgnored = Me.Factory.CreateRibbonButton
        Me.btnTestChrome = Me.Factory.CreateRibbonButton
        Me.Button2 = Me.Factory.CreateRibbonButton
        Me.Button3 = Me.Factory.CreateRibbonButton
        Me.Tab1.SuspendLayout()
        Me.Group1.SuspendLayout()
        Me.Group3.SuspendLayout()
        Me.Group4.SuspendLayout()
        Me.Group2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Tab1
        '
        Me.Tab1.Groups.Add(Me.Group1)
        Me.Tab1.Groups.Add(Me.Group3)
        Me.Tab1.Groups.Add(Me.Group4)
        Me.Tab1.Groups.Add(Me.Group2)
        Me.Tab1.Label = "DEP Tools"
        Me.Tab1.Name = "Tab1"
        '
        'Group1
        '
        Me.Group1.Items.Add(Me.CreateNew)
        Me.Group1.Items.Add(Me.TDOnly)
        Me.Group1.Items.Add(Me.BtnWCOnly)
        Me.Group1.Label = "Registration Tools"
        Me.Group1.Name = "Group1"
        '
        'Group3
        '
        Me.Group3.Items.Add(Me.CloseStale)
        Me.Group3.Items.Add(Me.BtnCheckRegistrations)
        Me.Group3.Label = "Ticket Tasks"
        Me.Group3.Name = "Group3"
        '
        'Group4
        '
        Me.Group4.Items.Add(Me.WriteMails)
        Me.Group4.Items.Add(Me.FindIgnored)
        Me.Group4.Label = "Sales Engagement Tools"
        Me.Group4.Name = "Group4"
        '
        'Group2
        '
        Me.Group2.Items.Add(Me.btnTestChrome)
        Me.Group2.Items.Add(Me.ChkDebug)
        Me.Group2.Items.Add(Me.Button2)
        Me.Group2.Items.Add(Me.Button3)
        Me.Group2.Label = "Debug Tools"
        Me.Group2.Name = "Group2"
        '
        'ChkDebug
        '
        Me.ChkDebug.Label = "Debug Mode"
        Me.ChkDebug.Name = "ChkDebug"
        '
        'CreateNew
        '
        Me.CreateNew.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.CreateNew.Image = CType(resources.GetObject("CreateNew.Image"), System.Drawing.Image)
        Me.CreateNew.Label = "Create New Tickets"
        Me.CreateNew.Name = "CreateNew"
        Me.CreateNew.ShowImage = True
        Me.CreateNew.SuperTip = "Create new nextdesk tickets based on the below DEP Spreadsheet"
        '
        'TDOnly
        '
        Me.TDOnly.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.TDOnly.Image = CType(resources.GetObject("TDOnly.Image"), System.Drawing.Image)
        Me.TDOnly.Label = "TechData Registrations"
        Me.TDOnly.Name = "TDOnly"
        Me.TDOnly.ShowImage = True
        '
        'BtnWCOnly
        '
        Me.BtnWCOnly.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.BtnWCOnly.Image = CType(resources.GetObject("BtnWCOnly.Image"), System.Drawing.Image)
        Me.BtnWCOnly.Label = "Westcoast Registrations"
        Me.BtnWCOnly.Name = "BtnWCOnly"
        Me.BtnWCOnly.ShowImage = True
        '
        'CloseStale
        '
        Me.CloseStale.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.CloseStale.Image = CType(resources.GetObject("CloseStale.Image"), System.Drawing.Image)
        Me.CloseStale.Label = "Close Stale Ticket"
        Me.CloseStale.Name = "CloseStale"
        Me.CloseStale.ShowImage = True
        '
        'BtnCheckRegistrations
        '
        Me.BtnCheckRegistrations.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.BtnCheckRegistrations.Image = CType(resources.GetObject("BtnCheckRegistrations.Image"), System.Drawing.Image)
        Me.BtnCheckRegistrations.Label = "Check Reg Status"
        Me.BtnCheckRegistrations.Name = "BtnCheckRegistrations"
        Me.BtnCheckRegistrations.ShowImage = True
        '
        'WriteMails
        '
        Me.WriteMails.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.WriteMails.Image = CType(resources.GetObject("WriteMails.Image"), System.Drawing.Image)
        Me.WriteMails.Label = "Write Pivot Emails"
        Me.WriteMails.Name = "WriteMails"
        Me.WriteMails.ShowImage = True
        '
        'FindIgnored
        '
        Me.FindIgnored.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.FindIgnored.Image = CType(resources.GetObject("FindIgnored.Image"), System.Drawing.Image)
        Me.FindIgnored.Label = "Find Ignored Tickets"
        Me.FindIgnored.Name = "FindIgnored"
        Me.FindIgnored.ShowImage = True
        '
        'btnTestChrome
        '
        Me.btnTestChrome.Label = "Test Chrome Control"
        Me.btnTestChrome.Name = "btnTestChrome"
        '
        'Button2
        '
        Me.Button2.Label = "Recheck Intranet"
        Me.Button2.Name = "Button2"
        '
        'Button3
        '
        Me.Button3.Label = "Version Number"
        Me.Button3.Name = "Button3"
        '
        'Ribbon1
        '
        Me.Name = "Ribbon1"
        Me.RibbonType = "Microsoft.Excel.Workbook"
        Me.Tabs.Add(Me.Tab1)
        Me.Tab1.ResumeLayout(False)
        Me.Tab1.PerformLayout()
        Me.Group1.ResumeLayout(False)
        Me.Group1.PerformLayout()
        Me.Group3.ResumeLayout(False)
        Me.Group3.PerformLayout()
        Me.Group4.ResumeLayout(False)
        Me.Group4.PerformLayout()
        Me.Group2.ResumeLayout(False)
        Me.Group2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents CreateNew As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents CloseStale As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents FindIgnored As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents WriteMails As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents TDOnly As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnTestChrome As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ChkDebug As Microsoft.Office.Tools.Ribbon.RibbonCheckBox
    Friend WithEvents BtnWCOnly As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group3 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents BtnCheckRegistrations As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group4 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Group2 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button2 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button3 As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As Ribbon1
        Get
            Return Me.GetRibbon(Of Ribbon1)()
        End Get
    End Property
End Class
