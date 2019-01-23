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
        Me.ChkDebug = Me.Factory.CreateRibbonCheckBox
        Me.CreateNew = Me.Factory.CreateRibbonButton
        Me.CloseStale = Me.Factory.CreateRibbonButton
        Me.FindIgnored = Me.Factory.CreateRibbonButton
        Me.WriteMails = Me.Factory.CreateRibbonButton
        Me.TDOnly = Me.Factory.CreateRibbonButton
        Me.Button1 = Me.Factory.CreateRibbonButton
        Me.BtnWCOnly = Me.Factory.CreateRibbonButton
        Me.Tab1.SuspendLayout()
        Me.Group1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Tab1
        '
        Me.Tab1.Groups.Add(Me.Group1)
        Me.Tab1.Label = "DEP Tools"
        Me.Tab1.Name = "Tab1"
        '
        'Group1
        '
        Me.Group1.Items.Add(Me.CreateNew)
        Me.Group1.Items.Add(Me.CloseStale)
        Me.Group1.Items.Add(Me.FindIgnored)
        Me.Group1.Items.Add(Me.WriteMails)
        Me.Group1.Items.Add(Me.TDOnly)
        Me.Group1.Items.Add(Me.Button1)
        Me.Group1.Items.Add(Me.ChkDebug)
        Me.Group1.Items.Add(Me.BtnWCOnly)
        Me.Group1.Name = "Group1"
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
        '
        'CloseStale
        '
        Me.CloseStale.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.CloseStale.Image = CType(resources.GetObject("CloseStale.Image"), System.Drawing.Image)
        Me.CloseStale.Label = "Close Stale Ticket"
        Me.CloseStale.Name = "CloseStale"
        Me.CloseStale.ShowImage = True
        '
        'FindIgnored
        '
        Me.FindIgnored.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.FindIgnored.Image = CType(resources.GetObject("FindIgnored.Image"), System.Drawing.Image)
        Me.FindIgnored.Label = "Find Ignored Tickets"
        Me.FindIgnored.Name = "FindIgnored"
        Me.FindIgnored.ShowImage = True
        '
        'WriteMails
        '
        Me.WriteMails.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.WriteMails.Image = CType(resources.GetObject("WriteMails.Image"), System.Drawing.Image)
        Me.WriteMails.Label = "Write Pivot Emails"
        Me.WriteMails.Name = "WriteMails"
        Me.WriteMails.ShowImage = True
        '
        'TDOnly
        '
        Me.TDOnly.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.TDOnly.Image = CType(resources.GetObject("TDOnly.Image"), System.Drawing.Image)
        Me.TDOnly.Label = "TechData Registrations"
        Me.TDOnly.Name = "TDOnly"
        Me.TDOnly.ShowImage = True
        '
        'Button1
        '
        Me.Button1.Label = "Test Chrome Control"
        Me.Button1.Name = "Button1"
        '
        'BtnWCOnly
        '
        Me.BtnWCOnly.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.BtnWCOnly.Image = CType(resources.GetObject("BtnWCOnly.Image"), System.Drawing.Image)
        Me.BtnWCOnly.Label = "Westcoast Registrations"
        Me.BtnWCOnly.Name = "BtnWCOnly"
        Me.BtnWCOnly.ShowImage = True
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
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents CreateNew As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents CloseStale As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents FindIgnored As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents WriteMails As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents TDOnly As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button1 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ChkDebug As Microsoft.Office.Tools.Ribbon.RibbonCheckBox
    Friend WithEvents BtnWCOnly As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As Ribbon1
        Get
            Return Me.GetRibbon(Of Ribbon1)()
        End Get
    End Property
End Class
