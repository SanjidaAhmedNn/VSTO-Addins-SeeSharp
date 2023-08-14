Partial Class Ribbon1
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
        Me.Button8 = Me.Factory.CreateRibbonButton
        Me.Button7 = Me.Factory.CreateRibbonButton
        Me.Button6 = Me.Factory.CreateRibbonButton
        Me.Button5 = Me.Factory.CreateRibbonButton
        Me.Button4 = Me.Factory.CreateRibbonButton
        Me.Button1 = Me.Factory.CreateRibbonButton
        Me.Button3 = Me.Factory.CreateRibbonButton
        Me.Group2 = Me.Factory.CreateRibbonGroup
        Me.Menu1 = Me.Factory.CreateRibbonMenu
        Me.SplitButton1 = Me.Factory.CreateRibbonSplitButton
        Me.Button9 = Me.Factory.CreateRibbonButton
        Me.Button10 = Me.Factory.CreateRibbonButton
        Me.DropDown1 = Me.Factory.CreateRibbonDropDown
        Me.ComboBox1 = Me.Factory.CreateRibbonComboBox
        Me.Group3 = Me.Factory.CreateRibbonGroup
        Me.SplitButton2 = Me.Factory.CreateRibbonSplitButton
        Me.Button16 = Me.Factory.CreateRibbonButton
        Me.Button17 = Me.Factory.CreateRibbonButton
        Me.Button18 = Me.Factory.CreateRibbonButton
        Me.Button11 = Me.Factory.CreateRibbonButton
        Me.Button12 = Me.Factory.CreateRibbonButton
        Me.Button13 = Me.Factory.CreateRibbonButton
        Me.Button14 = Me.Factory.CreateRibbonButton
        Me.Button15 = Me.Factory.CreateRibbonButton
        Me.Button19 = Me.Factory.CreateRibbonButton
        Me.Button20 = Me.Factory.CreateRibbonButton
        Me.Button21 = Me.Factory.CreateRibbonButton
        Me.Group4 = Me.Factory.CreateRibbonGroup
        Me.Button22 = Me.Factory.CreateRibbonButton
        Me.Button23 = Me.Factory.CreateRibbonButton
        Me.Menu2 = Me.Factory.CreateRibbonMenu
        Me.Button24 = Me.Factory.CreateRibbonButton
        Me.Button25 = Me.Factory.CreateRibbonButton
        Me.Button26 = Me.Factory.CreateRibbonButton
        Me.Button27 = Me.Factory.CreateRibbonButton
        Me.Menu3 = Me.Factory.CreateRibbonMenu
        Me.Button28 = Me.Factory.CreateRibbonButton
        Me.Button29 = Me.Factory.CreateRibbonButton
        Me.Button30 = Me.Factory.CreateRibbonButton
        Me.Tab1.SuspendLayout()
        Me.Group1.SuspendLayout()
        Me.Group2.SuspendLayout()
        Me.Group3.SuspendLayout()
        Me.Group4.SuspendLayout()
        Me.SuspendLayout()
        '
        'Tab1
        '
        Me.Tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.Tab1.Groups.Add(Me.Group1)
        Me.Tab1.Groups.Add(Me.Group2)
        Me.Tab1.Groups.Add(Me.Group3)
        Me.Tab1.Groups.Add(Me.Group4)
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
        Me.Group1.Items.Add(Me.Button3)
        Me.Group1.Label = "Group1"
        Me.Group1.Name = "Group1"
        '
        'Button8
        '
        Me.Button8.Label = "Swap"
        Me.Button8.Name = "Button8"
        '
        'Button7
        '
        Me.Button7.Label = "Tranform"
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
        Me.Button1.Label = "Flip"
        Me.Button1.Name = "Button1"
        '
        'Button3
        '
        Me.Button3.Label = "Transpose"
        Me.Button3.Name = "Button3"
        '
        'Group2
        '
        Me.Group2.Items.Add(Me.Menu1)
        Me.Group2.Items.Add(Me.DropDown1)
        Me.Group2.Items.Add(Me.ComboBox1)
        Me.Group2.Label = "Group2"
        Me.Group2.Name = "Group2"
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
        'Group3
        '
        Me.Group3.Items.Add(Me.SplitButton2)
        Me.Group3.Items.Add(Me.Button11)
        Me.Group3.Items.Add(Me.Button12)
        Me.Group3.Items.Add(Me.Button13)
        Me.Group3.Items.Add(Me.Button14)
        Me.Group3.Items.Add(Me.Button15)
        Me.Group3.Items.Add(Me.Button19)
        Me.Group3.Items.Add(Me.Button20)
        Me.Group3.Items.Add(Me.Button21)
        Me.Group3.Label = "Test"
        Me.Group3.Name = "Group3"
        '
        'SplitButton2
        '
        Me.SplitButton2.Items.Add(Me.Button16)
        Me.SplitButton2.Items.Add(Me.Button17)
        Me.SplitButton2.Items.Add(Me.Button18)
        Me.SplitButton2.Label = "Combine Range"
        Me.SplitButton2.Name = "SplitButton2"
        '
        'Button16
        '
        Me.Button16.Label = "Form 18"
        Me.Button16.Name = "Button16"
        Me.Button16.ShowImage = True
        '
        'Button17
        '
        Me.Button17.Label = "Form 19"
        Me.Button17.Name = "Button17"
        Me.Button17.ShowImage = True
        '
        'Button18
        '
        Me.Button18.Label = "Form 20"
        Me.Button18.Name = "Button18"
        Me.Button18.ShowImage = True
        '
        'Button11
        '
        Me.Button11.Label = "Form 13"
        Me.Button11.Name = "Button11"
        '
        'Button12
        '
        Me.Button12.Label = "Form 14"
        Me.Button12.Name = "Button12"
        '
        'Button13
        '
        Me.Button13.Label = "Form15"
        Me.Button13.Name = "Button13"
        '
        'Button14
        '
        Me.Button14.Label = "Form 16"
        Me.Button14.Name = "Button14"
        '
        'Button15
        '
        Me.Button15.Label = "Form 17"
        Me.Button15.Name = "Button15"
        '
        'Button19
        '
        Me.Button19.Label = "Form 21 Fill Emty Cells"
        Me.Button19.Name = "Button19"
        '
        'Button20
        '
        Me.Button20.Label = "Form 22 Merge Dupli"
        Me.Button20.Name = "Button20"
        '
        'Button21
        '
        Me.Button21.Label = "Form 23 Merge Dupli Row"
        Me.Button21.Name = "Button21"
        '
        'Group4
        '
        Me.Group4.Items.Add(Me.Button22)
        Me.Group4.Items.Add(Me.Button23)
        Me.Group4.Items.Add(Me.Menu2)
        Me.Group4.Items.Add(Me.Button27)
        Me.Group4.Items.Add(Me.Menu3)
        Me.Group4.Label = "Group4"
        Me.Group4.Name = "Group4"
        '
        'Button22
        '
        Me.Button22.Label = "Form 24 Split Cells"
        Me.Button22.Name = "Button22"
        '
        'Button23
        '
        Me.Button23.Label = "Form 25 Split Range"
        Me.Button23.Name = "Button23"
        '
        'Menu2
        '
        Me.Menu2.Items.Add(Me.Button24)
        Me.Menu2.Items.Add(Me.Button25)
        Me.Menu2.Items.Add(Me.Button26)
        Me.Menu2.Label = "Split_26-27-28"
        Me.Menu2.Name = "Menu2"
        '
        'Button24
        '
        Me.Button24.Label = "Form 26"
        Me.Button24.Name = "Button24"
        Me.Button24.ShowImage = True
        '
        'Button25
        '
        Me.Button25.Label = "Form 27"
        Me.Button25.Name = "Button25"
        Me.Button25.ShowImage = True
        '
        'Button26
        '
        Me.Button26.Label = "Form 28"
        Me.Button26.Name = "Button26"
        Me.Button26.ShowImage = True
        '
        'Button27
        '
        Me.Button27.Label = "Form 29"
        Me.Button27.Name = "Button27"
        '
        'Menu3
        '
        Me.Menu3.Items.Add(Me.Button28)
        Me.Menu3.Items.Add(Me.Button29)
        Me.Menu3.Items.Add(Me.Button30)
        Me.Menu3.Label = "Dynamic Drop-down List"
        Me.Menu3.Name = "Menu3"
        '
        'Button28
        '
        Me.Button28.Label = "Create"
        Me.Button28.Name = "Button28"
        Me.Button28.ShowImage = True
        '
        'Button29
        '
        Me.Button29.Label = "Update"
        Me.Button29.Name = "Button29"
        Me.Button29.ShowImage = True
        '
        'Button30
        '
        Me.Button30.Label = "Expand"
        Me.Button30.Name = "Button30"
        Me.Button30.ShowImage = True
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
        Me.Group3.ResumeLayout(False)
        Me.Group3.PerformLayout()
        Me.Group4.ResumeLayout(False)
        Me.Group4.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button1 As Microsoft.Office.Tools.Ribbon.RibbonButton
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
    Friend WithEvents Group3 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button11 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button12 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button13 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button14 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button15 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents SplitButton2 As Microsoft.Office.Tools.Ribbon.RibbonSplitButton
    Friend WithEvents Button16 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button17 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button18 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button19 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button20 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button21 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group4 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button22 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button23 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Menu2 As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents Button24 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button25 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button26 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button27 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Menu3 As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents Button28 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button29 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button30 As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As Ribbon1
        Get
            Return Me.GetRibbon(Of Ribbon1)()
        End Get
    End Property
End Class
