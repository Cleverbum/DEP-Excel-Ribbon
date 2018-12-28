﻿Partial Class Ribbon1
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
        Me.CreateNew = Me.Factory.CreateRibbonButton
        Me.CloseStale = Me.Factory.CreateRibbonButton
        Me.FindIgnored = Me.Factory.CreateRibbonButton
        Me.Button4 = Me.Factory.CreateRibbonButton
        Me.WriteMails = Me.Factory.CreateRibbonButton
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
        Me.Group1.Items.Add(Me.Button4)
        Me.Group1.Name = "Group1"
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
        'Button4
        '
        Me.Button4.Label = "Test Chrome Control"
        Me.Button4.Name = "Button4"
        '
        'WriteMails
        '
        Me.WriteMails.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.WriteMails.Image = CType(resources.GetObject("WriteMails.Image"), System.Drawing.Image)
        Me.WriteMails.Label = "Write Pivot Emails"
        Me.WriteMails.Name = "WriteMails"
        Me.WriteMails.ShowImage = True
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
    Friend WithEvents Button4 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents WriteMails As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As Ribbon1
        Get
            Return Me.GetRibbon(Of Ribbon1)()
        End Get
    End Property
End Class
