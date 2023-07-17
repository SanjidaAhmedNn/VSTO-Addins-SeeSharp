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
        Me.Tab1 = Me.Factory.CreateRibbonTab
        Me.Group1 = Me.Factory.CreateRibbonGroup
        Me.Group2 = Me.Factory.CreateRibbonGroup
        Me.DropDown1 = Me.Factory.CreateRibbonDropDown
        Me.ComboBox1 = Me.Factory.CreateRibbonComboBox
        Me.Button8 = Me.Factory.CreateRibbonButton
        Me.Button7 = Me.Factory.CreateRibbonButton
        Me.Button6 = Me.Factory.CreateRibbonButton
        Me.Button5 = Me.Factory.CreateRibbonButton
        Me.Button4 = Me.Factory.CreateRibbonButton
        Me.Button1 = Me.Factory.CreateRibbonButton
        Me.Transform = Me.Factory.CreateRibbonButton
        Me.Button3 = Me.Factory.CreateRibbonButton
        Me.Menu1 = Me.Factory.CreateRibbonMenu
        Me.SplitButton1 = Me.Factory.CreateRibbonSplitButton
        Me.Button9 = Me.Factory.CreateRibbonButton
        Me.Button10 = Me.Factory.CreateRibbonButton
        Me.Tab1.SuspendLayout()
        Me.Group1.SuspendLayout()
        Me.Group2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Tab1
        '
        Me.Tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.Tab1.Groups.Add(Me.Group1)
        Me.Tab1.Groups.Add(Me.Group2)
        Me.Tab1.Label = "TabAddIns"
        Me.Tab1.Name = "Tab1"
        '
        'Group1
        '
        Me.Group1.Items.Add(Me.Button8)
        Me.Group1.Items.Add(Me.Button7)
        Me.Group1.Items.Add(Me.Button6)
        Me.Group1.Items.Add(Me.Button5)
        Me.Group1.Items.Add(Me.Button4)
        Me.Group1.Items.Add(Me.Button1)
        Me.Group1.Items.Add(Me.Transform)
        Me.Group1.Items.Add(Me.Button3)
        Me.Group1.Label = "Group1"
        Me.Group1.Name = "Group1"
        '
        'Group2
        '
        Me.Group2.Items.Add(Me.Menu1)
        Me.Group2.Items.Add(Me.DropDown1)
        Me.Group2.Items.Add(Me.ComboBox1)
        Me.Group2.Label = "Group2"
        Me.Group2.Name = "Group2"
        '
        'DropDown1
        '
        Me.DropDown1.Label = "DropDown1"
        Me.DropDown1.Name = "DropDown1"
        '
        'ComboBox1
        '
        Me.ComboBox1.Label = "ComboBox1"
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Text = Nothing
        '
        'Button8
        '
        Me.Button8.Label = "Form 11"
        Me.Button8.Name = "Button8"
        '
        'Button7
        '
        Me.Button7.Label = "Button7"
        Me.Button7.Name = "Button7"
        '
        'Button6
        '
        Me.Button6.Label = "Unmerge"
        Me.Button6.Name = "Button6"
        '
        'Button5
        '
        Me.Button5.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button5.Label = "Merge Cells with Same Value"
        Me.Button5.Name = "Button5"
        Me.Button5.ShowImage = True
        '
        'Button4
        '
        Me.Button4.Label = "Form 12"
        Me.Button4.Name = "Button4"
        '
        'Button1
        '
        Me.Button1.Label = "Flip Design 1"
        Me.Button1.Name = "Button1"
        '
        'Transform
        '
        Me.Transform.Label = "Transform"
        Me.Transform.Name = "Transform"
        '
        'Button3
        '
        Me.Button3.Label = "Transpose"
        Me.Button3.Name = "Button3"
        '
        'Menu1
        '
        Me.Menu1.Items.Add(Me.SplitButton1)
        Me.Menu1.Label = "Range"
        Me.Menu1.Name = "Menu1"
        '
        'SplitButton1
        '
        Me.SplitButton1.Items.Add(Me.Button9)
        Me.SplitButton1.Items.Add(Me.Button10)
        Me.SplitButton1.Label = "Specify Scroll Area"
        Me.SplitButton1.Name = "SplitButton1"
        '
        'Button9
        '
        Me.Button9.Label = "Set"
        Me.Button9.Name = "Button9"
        Me.Button9.ShowImage = True
        '
        'Button10
        '
        Me.Button10.Label = "Clear"
        Me.Button10.Name = "Button10"
        Me.Button10.ShowImage = True
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
        Me.Group2.ResumeLayout(False)
        Me.Group2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button1 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Transform As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button3 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button4 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button5 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button6 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button7 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group2 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents DropDown1 As Microsoft.Office.Tools.Ribbon.RibbonDropDown
    Friend WithEvents ComboBox1 As Microsoft.Office.Tools.Ribbon.RibbonComboBox
    Friend WithEvents Menu1 As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents SplitButton1 As Microsoft.Office.Tools.Ribbon.RibbonSplitButton
    Friend WithEvents Button9 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button10 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button8 As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As Ribbon1
        Get
            Return Me.GetRibbon(Of Ribbon1)()
        End Get
    End Property
End Class
