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
        Me.Button1 = Me.Factory.CreateRibbonButton
        Me.Button8 = Me.Factory.CreateRibbonButton
        Me.Button3 = Me.Factory.CreateRibbonButton
        Me.Separator1 = Me.Factory.CreateRibbonSeparator
        Me.Button7 = Me.Factory.CreateRibbonButton
        Me.Button13 = Me.Factory.CreateRibbonButton
        Me.Menu2 = Me.Factory.CreateRibbonMenu
        Me.Button12 = Me.Factory.CreateRibbonButton
        Me.Button10 = Me.Factory.CreateRibbonButton
        Me.Button14 = Me.Factory.CreateRibbonButton
        Me.Group2 = Me.Factory.CreateRibbonGroup
        Me.Button5 = Me.Factory.CreateRibbonButton
        Me.Menu5 = Me.Factory.CreateRibbonMenu
        Me.Button16 = Me.Factory.CreateRibbonButton
        Me.Button17 = Me.Factory.CreateRibbonButton
        Me.Button18 = Me.Factory.CreateRibbonButton
        Me.Menu8 = Me.Factory.CreateRibbonMenu
        Me.Button20 = Me.Factory.CreateRibbonButton
        Me.Button21 = Me.Factory.CreateRibbonButton
        Me.Button6 = Me.Factory.CreateRibbonButton
        Me.Menu10 = Me.Factory.CreateRibbonMenu
        Me.Button23 = Me.Factory.CreateRibbonButton
        Me.Button22 = Me.Factory.CreateRibbonButton
        Me.Menu9 = Me.Factory.CreateRibbonMenu
        Me.Button45 = Me.Factory.CreateRibbonButton
        Me.Button46 = Me.Factory.CreateRibbonButton
        Me.Button47 = Me.Factory.CreateRibbonButton
        Me.Button15 = Me.Factory.CreateRibbonButton
        Me.Group5 = Me.Factory.CreateRibbonGroup
        Me.Menu11 = Me.Factory.CreateRibbonMenu
        Me.Button54 = Me.Factory.CreateRibbonButton
        Me.Button11 = Me.Factory.CreateRibbonButton
        Me.Menu4 = Me.Factory.CreateRibbonMenu
        Me.Button31 = Me.Factory.CreateRibbonButton
        Me.Button32 = Me.Factory.CreateRibbonButton
        Me.Group3 = Me.Factory.CreateRibbonGroup
        Me.Menu7 = Me.Factory.CreateRibbonMenu
        Me.Button37 = Me.Factory.CreateRibbonButton
        Me.Button38 = Me.Factory.CreateRibbonButton
        Me.Button39 = Me.Factory.CreateRibbonButton
        Me.Button40 = Me.Factory.CreateRibbonButton
        Me.Menu6 = Me.Factory.CreateRibbonMenu
        Me.Button33 = Me.Factory.CreateRibbonButton
        Me.Button34 = Me.Factory.CreateRibbonButton
        Me.Button35 = Me.Factory.CreateRibbonButton
        Me.Button36 = Me.Factory.CreateRibbonButton
        Me.Button41 = Me.Factory.CreateRibbonButton
        Me.Separator2 = Me.Factory.CreateRibbonSeparator
        Me.Button19 = Me.Factory.CreateRibbonButton
        Me.Group4 = Me.Factory.CreateRibbonGroup
        Me.Menu1 = Me.Factory.CreateRibbonMenu
        Me.Button2 = Me.Factory.CreateRibbonButton
        Me.Button9 = Me.Factory.CreateRibbonButton
        Me.Button49 = Me.Factory.CreateRibbonButton
        Me.Menu3 = Me.Factory.CreateRibbonMenu
        Me.Button28 = Me.Factory.CreateRibbonButton
        Me.Button29 = Me.Factory.CreateRibbonButton
        Me.Button30 = Me.Factory.CreateRibbonButton
        Me.SplitButton7 = Me.Factory.CreateRibbonSplitButton
        Me.Button24 = Me.Factory.CreateRibbonButton
        Me.Button25 = Me.Factory.CreateRibbonButton
        Me.Button26 = Me.Factory.CreateRibbonButton
        Me.Button27 = Me.Factory.CreateRibbonButton
        Me.Tab2 = Me.Factory.CreateRibbonTab
        Me.Tab1.SuspendLayout()
        Me.Group1.SuspendLayout()
        Me.Group2.SuspendLayout()
        Me.Group5.SuspendLayout()
        Me.Group3.SuspendLayout()
        Me.Group4.SuspendLayout()
        Me.Tab2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Tab1
        '
        Me.Tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.Tab1.Groups.Add(Me.Group1)
        Me.Tab1.Groups.Add(Me.Group2)
        Me.Tab1.Groups.Add(Me.Group5)
        Me.Tab1.Groups.Add(Me.Group3)
        Me.Tab1.Groups.Add(Me.Group4)
        Me.Tab1.Label = "TabAddIns"
        Me.Tab1.Name = "Tab1"
        '
        'Group1
        '
        Me.Group1.Items.Add(Me.Button1)
        Me.Group1.Items.Add(Me.Button8)
        Me.Group1.Items.Add(Me.Button3)
        Me.Group1.Items.Add(Me.Separator1)
        Me.Group1.Items.Add(Me.Button7)
        Me.Group1.Items.Add(Me.Button13)
        Me.Group1.Items.Add(Me.Menu2)
        Me.Group1.Items.Add(Me.Button14)
        Me.Group1.Label = "Range"
        Me.Group1.Name = "Group1"
        '
        'Button1
        '
        Me.Button1.Label = "    Flip"
        Me.Button1.Name = "Button1"
        '
        'Button8
        '
        Me.Button8.Label = "  Swap"
        Me.Button8.Name = "Button8"
        '
        'Button3
        '
        Me.Button3.Label = "Transpose"
        Me.Button3.Name = "Button3"
        '
        'Separator1
        '
        Me.Separator1.Name = "Separator1"
        '
        'Button7
        '
        Me.Button7.Label = "Transform"
        Me.Button7.Name = "Button7"
        '
        'Button13
        '
        Me.Button13.Label = "Compare Cells"
        Me.Button13.Name = "Button13"
        '
        'Menu2
        '
        Me.Menu2.Items.Add(Me.Button12)
        Me.Menu2.Items.Add(Me.Button10)
        Me.Menu2.Label = "Specify Scroll Area"
        Me.Menu2.Name = "Menu2"
        '
        'Button12
        '
        Me.Button12.Label = "Set"
        Me.Button12.Name = "Button12"
        Me.Button12.ShowImage = True
        '
        'Button10
        '
        Me.Button10.Label = "Clear"
        Me.Button10.Name = "Button10"
        Me.Button10.ShowImage = True
        '
        'Button14
        '
        Me.Button14.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button14.Label = "Paste into Visible Range"
        Me.Button14.Name = "Button14"
        Me.Button14.ShowImage = True
        '
        'Group2
        '
        Me.Group2.Items.Add(Me.Button5)
        Me.Group2.Items.Add(Me.Menu5)
        Me.Group2.Items.Add(Me.Menu8)
        Me.Group2.Items.Add(Me.Button6)
        Me.Group2.Items.Add(Me.Menu10)
        Me.Group2.Items.Add(Me.Menu9)
        Me.Group2.Items.Add(Me.Button15)
        Me.Group2.Label = "Merge && Unmerge"
        Me.Group2.Name = "Group2"
        '
        'Button5
        '
        Me.Button5.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button5.Label = "Merge Cells with Same Value"
        Me.Button5.Name = "Button5"
        Me.Button5.ShowImage = True
        '
        'Menu5
        '
        Me.Menu5.Items.Add(Me.Button16)
        Me.Menu5.Items.Add(Me.Button17)
        Me.Menu5.Items.Add(Me.Button18)
        Me.Menu5.Label = "Combine Range"
        Me.Menu5.Name = "Menu5"
        '
        'Button16
        '
        Me.Button16.Label = "Combine Ranges into Column"
        Me.Button16.Name = "Button16"
        Me.Button16.ShowImage = True
        '
        'Button17
        '
        Me.Button17.Label = "Combine Ranges into Row"
        Me.Button17.Name = "Button17"
        Me.Button17.ShowImage = True
        '
        'Button18
        '
        Me.Button18.Label = "Combine Ranges into Cells"
        Me.Button18.Name = "Button18"
        Me.Button18.ShowImage = True
        '
        'Menu8
        '
        Me.Menu8.Items.Add(Me.Button20)
        Me.Menu8.Items.Add(Me.Button21)
        Me.Menu8.Label = "Combine Duplicate"
        Me.Menu8.Name = "Menu8"
        '
        'Button20
        '
        Me.Button20.Label = "Combine Duplicate Rows"
        Me.Button20.Name = "Button20"
        Me.Button20.ShowImage = True
        '
        'Button21
        '
        Me.Button21.Label = "Combine Duplicate Columns"
        Me.Button21.Name = "Button21"
        Me.Button21.ShowImage = True
        '
        'Button6
        '
        Me.Button6.Label = "Unmerge Cells with Value"
        Me.Button6.Name = "Button6"
        '
        'Menu10
        '
        Me.Menu10.Items.Add(Me.Button23)
        Me.Menu10.Items.Add(Me.Button22)
        Me.Menu10.Label = "Split Data"
        Me.Menu10.Name = "Menu10"
        '
        'Button23
        '
        Me.Button23.Label = "Split Range"
        Me.Button23.Name = "Button23"
        Me.Button23.ShowImage = True
        '
        'Button22
        '
        Me.Button22.Label = "Split Cells"
        Me.Button22.Name = "Button22"
        Me.Button22.ShowImage = True
        '
        'Menu9
        '
        Me.Menu9.Items.Add(Me.Button45)
        Me.Menu9.Items.Add(Me.Button46)
        Me.Menu9.Items.Add(Me.Button47)
        Me.Menu9.Label = "Split Text"
        Me.Menu9.Name = "Menu9"
        '
        'Button45
        '
        Me.Button45.Label = "Split Text by Characters"
        Me.Button45.Name = "Button45"
        Me.Button45.ShowImage = True
        '
        'Button46
        '
        Me.Button46.Label = "Split Text by Strings"
        Me.Button46.Name = "Button46"
        Me.Button46.ShowImage = True
        '
        'Button47
        '
        Me.Button47.Label = "Split Text by Pattern"
        Me.Button47.Name = "Button47"
        Me.Button47.ShowImage = True
        '
        'Button15
        '
        Me.Button15.Label = "Divide Names"
        Me.Button15.Name = "Button15"
        '
        'Group5
        '
        Me.Group5.Items.Add(Me.Menu11)
        Me.Group5.Items.Add(Me.Menu4)
        Me.Group5.Label = "Hide && Unhide"
        Me.Group5.Name = "Group5"
        '
        'Menu11
        '
        Me.Menu11.Items.Add(Me.Button54)
        Me.Menu11.Items.Add(Me.Button11)
        Me.Menu11.Label = "Hide Ranges"
        Me.Menu11.Name = "Menu11"
        '
        'Button54
        '
        Me.Button54.Label = "Hide only the selected range"
        Me.Button54.Name = "Button54"
        Me.Button54.ShowImage = True
        '
        'Button11
        '
        Me.Button11.Label = "Hide all except the selected range"
        Me.Button11.Name = "Button11"
        Me.Button11.ShowImage = True
        '
        'Menu4
        '
        Me.Menu4.Items.Add(Me.Button31)
        Me.Menu4.Items.Add(Me.Button32)
        Me.Menu4.Label = "Unhide Ranges"
        Me.Menu4.Name = "Menu4"
        '
        'Button31
        '
        Me.Button31.Label = "Unhide All Ranges"
        Me.Button31.Name = "Button31"
        Me.Button31.ShowImage = True
        '
        'Button32
        '
        Me.Button32.Label = "Unhide Ranges from the Selection"
        Me.Button32.Name = "Button32"
        Me.Button32.ShowImage = True
        '
        'Group3
        '
        Me.Group3.Items.Add(Me.Menu7)
        Me.Group3.Items.Add(Me.Menu6)
        Me.Group3.Items.Add(Me.Button41)
        Me.Group3.Items.Add(Me.Separator2)
        Me.Group3.Items.Add(Me.Button19)
        Me.Group3.Label = "Remove Blanks"
        Me.Group3.Name = "Group3"
        '
        'Menu7
        '
        Me.Menu7.Items.Add(Me.Button37)
        Me.Menu7.Items.Add(Me.Button38)
        Me.Menu7.Items.Add(Me.Button39)
        Me.Menu7.Items.Add(Me.Button40)
        Me.Menu7.Label = "Empty Rows"
        Me.Menu7.Name = "Menu7"
        Me.Menu7.ShowImage = True
        '
        'Button37
        '
        Me.Button37.Label = "From Selected Range"
        Me.Button37.Name = "Button37"
        Me.Button37.ShowImage = True
        '
        'Button38
        '
        Me.Button38.Label = "From Active Sheet"
        Me.Button38.Name = "Button38"
        Me.Button38.ShowImage = True
        '
        'Button39
        '
        Me.Button39.Label = "From Selected Sheets"
        Me.Button39.Name = "Button39"
        Me.Button39.ShowImage = True
        '
        'Button40
        '
        Me.Button40.Label = "From All Sheets"
        Me.Button40.Name = "Button40"
        Me.Button40.ShowImage = True
        '
        'Menu6
        '
        Me.Menu6.Items.Add(Me.Button33)
        Me.Menu6.Items.Add(Me.Button34)
        Me.Menu6.Items.Add(Me.Button35)
        Me.Menu6.Items.Add(Me.Button36)
        Me.Menu6.Label = "Empty Columns"
        Me.Menu6.Name = "Menu6"
        Me.Menu6.ShowImage = True
        '
        'Button33
        '
        Me.Button33.Label = "From Selected Range"
        Me.Button33.Name = "Button33"
        Me.Button33.ShowImage = True
        '
        'Button34
        '
        Me.Button34.Label = "From Active Sheet"
        Me.Button34.Name = "Button34"
        Me.Button34.ShowImage = True
        '
        'Button35
        '
        Me.Button35.Label = "From Selected Sheets"
        Me.Button35.Name = "Button35"
        Me.Button35.ShowImage = True
        '
        'Button36
        '
        Me.Button36.Label = "From All Sheets"
        Me.Button36.Name = "Button36"
        Me.Button36.ShowImage = True
        '
        'Button41
        '
        Me.Button41.Label = "Empty Sheets"
        Me.Button41.Name = "Button41"
        Me.Button41.ShowImage = True
        '
        'Separator2
        '
        Me.Separator2.Name = "Separator2"
        '
        'Button19
        '
        Me.Button19.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button19.Label = "Fill Empty Cells"
        Me.Button19.Name = "Button19"
        Me.Button19.ShowImage = True
        '
        'Group4
        '
        Me.Group4.Items.Add(Me.Menu1)
        Me.Group4.Items.Add(Me.Button49)
        Me.Group4.Items.Add(Me.Menu3)
        Me.Group4.Items.Add(Me.SplitButton7)
        Me.Group4.Label = "Drop-down List"
        Me.Group4.Name = "Group4"
        '
        'Menu1
        '
        Me.Menu1.Items.Add(Me.Button2)
        Me.Menu1.Items.Add(Me.Button9)
        Me.Menu1.Label = "Create Drop-down List"
        Me.Menu1.Name = "Menu1"
        '
        'Button2
        '
        Me.Button2.Label = "Simple Drop-down List"
        Me.Button2.Name = "Button2"
        Me.Button2.ShowImage = True
        '
        'Button9
        '
        Me.Button9.Label = "Picture Based Drop-down List"
        Me.Button9.Name = "Button9"
        Me.Button9.ShowImage = True
        '
        'Button49
        '
        Me.Button49.Label = "Color Based Drop-down List"
        Me.Button49.Name = "Button49"
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
        Me.Button30.Label = "Extend"
        Me.Button30.Name = "Button30"
        Me.Button30.ShowImage = True
        '
        'SplitButton7
        '
        Me.SplitButton7.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.SplitButton7.Items.Add(Me.Button24)
        Me.SplitButton7.Items.Add(Me.Button25)
        Me.SplitButton7.Items.Add(Me.Button26)
        Me.SplitButton7.Items.Add(Me.Button27)
        Me.SplitButton7.Label = "Advanced Drop-down List"
        Me.SplitButton7.Name = "SplitButton7"
        '
        'Button24
        '
        Me.Button24.Label = "Multi-selection Based Drop-down List"
        Me.Button24.Name = "Button24"
        Me.Button24.ShowImage = True
        '
        'Button25
        '
        Me.Button25.Label = "Drop-down List with Checkbox"
        Me.Button25.Name = "Button25"
        Me.Button25.ShowImage = True
        '
        'Button26
        '
        Me.Button26.Label = "Drop-down List with Search Option"
        Me.Button26.Name = "Button26"
        Me.Button26.ShowImage = True
        '
        'Button27
        '
        Me.Button27.Label = "Remove Advanced Drop-down List"
        Me.Button27.Name = "Button27"
        Me.Button27.ShowImage = True
        '
        'Tab2
        '
        Me.Tab2.Label = "Tab2"
        Me.Tab2.Name = "Tab2"
        '
        'Ribbon1
        '
        Me.Name = "Ribbon1"
        Me.RibbonType = "Microsoft.Excel.Workbook"
        Me.Tabs.Add(Me.Tab1)
        Me.Tabs.Add(Me.Tab2)
        Me.Tab1.ResumeLayout(False)
        Me.Tab1.PerformLayout()
        Me.Group1.ResumeLayout(False)
        Me.Group1.PerformLayout()
        Me.Group2.ResumeLayout(False)
        Me.Group2.PerformLayout()
        Me.Group5.ResumeLayout(False)
        Me.Group5.PerformLayout()
        Me.Group3.ResumeLayout(False)
        Me.Group3.PerformLayout()
        Me.Group4.ResumeLayout(False)
        Me.Group4.PerformLayout()
        Me.Tab2.ResumeLayout(False)
        Me.Tab2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button1 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button3 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button5 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button6 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button7 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group2 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button8 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group3 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button13 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button14 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button15 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button19 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group4 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Menu3 As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents Button28 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button29 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button30 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Menu4 As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents Button31 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button32 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Menu6 As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents Button33 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button34 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button35 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button36 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Menu7 As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents Button41 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button37 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button38 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button39 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button40 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group5 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button49 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents SplitButton7 As Microsoft.Office.Tools.Ribbon.RibbonSplitButton
    Friend WithEvents Separator1 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents Separator2 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents Menu1 As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents Button2 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button9 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Menu2 As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents Button12 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button10 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Menu5 As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents Button16 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button17 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button18 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Menu8 As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents Button20 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button21 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Menu10 As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents Button23 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button22 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Menu9 As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents Button45 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button46 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button47 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Menu11 As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents Button54 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button11 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button24 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button25 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button26 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button27 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Tab2 As Microsoft.Office.Tools.Ribbon.RibbonTab
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As Ribbon1
        Get
            Return Me.GetRibbon(Of Ribbon1)()
        End Get
    End Property
End Class
