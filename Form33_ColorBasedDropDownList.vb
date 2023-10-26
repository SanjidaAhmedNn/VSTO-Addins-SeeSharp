Imports System.ComponentModel
Imports System.Drawing
Imports System.Linq.Expressions
Imports System.Reflection
Imports System.Reflection.Emit
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Button

Public Class Form33_ColorBasedDropDownList
    Dim WithEvents excelApp As Excel.Application
    Dim workBook As Excel.Workbook
    Public Shared workSheet As Excel.Worksheet
    Dim workSheet2 As Excel.Worksheet
    Dim workSheet3 As Excel.Worksheet
    Dim src_rng As Excel.Range
    Public des_rng As Excel.Range
    Dim selectedRange As Excel.Range
    Public ax As String

    Dim opened As Integer
    Dim objectPosition As New Point() ' For 2D
    Public mybtn As Object
    Public focuschange As Boolean
    Public form As Form42 = Nothing
    Dim flag As Boolean = False
    Public form2 As Form43 = Nothing
    Dim flag2 As Boolean = False



    Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As IntPtr, ByVal hWndInsertAfter As IntPtr, ByVal X As Integer, ByVal Y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal uFlags As UInteger) As Boolean
    Private Const SWP_NOMOVE As UInteger = &H2
    Private Const SWP_NOSIZE As UInteger = &H1
    Private Const SWP_NOACTIVATE As UInteger = &H10
    Private Const HWND_TOPMOST As Integer = -1
    ' Declare the tooltip at class level
    Private tooltip As New ToolTip()
    Private Sub Form1_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Enter Then
            btn_OK.PerformClick()


        End If
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'ReDim mybtn(List_Preview.Items.Count)
        'MsgBox
        Me.KeyPreview = True
        TB_des_rng.Enabled = False
        Selection_destination.Enabled = False

        ' Define the first 42 colors from the Visual Studio Custom tab with their names
        Dim vsColors As Dictionary(Of String, Color) = New Dictionary(Of String, Color) From {
        {"White", Color.FromArgb(255, 255, 255)},    '1
        {"Aqua Light", Color.FromArgb(228, 239, 240)}, '2
        {"Blue Light", Color.FromArgb(127, 127, 127)},     '3
        {"Rose", Color.FromArgb(250, 214, 212)},   '4
        {"Light Yellow", Color.FromArgb(255, 255, 235)},  '5
        {"Lavender", Color.FromArgb(255, 153, 255)},    '6
        {"Lime", Color.FromArgb(233, 249, 198)},     '7
        {"Light Gray", Color.FromArgb(217, 217, 217)},     '8
        {"Aqua", Color.FromArgb(188, 215, 218)},     '9
        {"Light Torquoise", Color.FromArgb(102, 204, 255)},      '10
        {"Light Red", Color.FromArgb(245, 174, 169)},  '11
        {"Light Medium Yellow", Color.FromArgb(255, 255, 204)},          '12
        {"Pink", Color.FromArgb(255, 153, 204)},            '13
        {"Light Green", Color.FromArgb(200, 249, 207)},         '14
        {"Gray", Color.FromArgb(166, 166, 166)},        '15
        {"Teal", Color.FromArgb(122, 174, 181)},     '16
        {"Blue", Color.FromArgb(51, 102, 255)},             '17
        {"Medium Red", Color.FromArgb(241, 133, 127)},  '18
        {"Yellow", Color.FromArgb(255, 255, 153)},          '19
        {"Medium Pink", Color.FromArgb(255, 51, 204)},    '20
        {"Medium Green", Color.FromArgb(91, 138, 212)},         '21
        {"Dark Gray", Color.FromArgb(20, 26, 26)},   '22
        {"Aqua Medium", Color.FromArgb(71, 121, 128)},        '23
        {"Royal Blue", Color.FromArgb(0, 0, 255)},    '24
        {"Red", Color.FromArgb(183, 30, 21)},          '25
        {"Medium Yellow", Color.FromArgb(255, 255, 102)},          '26
        {"Dark Pink", Color.FromArgb(204, 51, 153)},        '27
        {"Green", Color.FromArgb(51, 153, 102)},    '28
        {"Black", Color.FromArgb(64, 64, 64)},     '29
        {"Dark Aqua", Color.FromArgb(49, 83, 88)},      '30
        {"Medium Dark Blue", Color.FromArgb(0, 51, 153)},         '31
        {"Dark Red", Color.FromArgb(122, 20, 14)},    '32
        {"Dark Yellow", Color.FromArgb(255, 255, 0)},         '33
        {"Plum", Color.FromArgb(153, 44, 98)},         '34
        {"Medium Dark Green", Color.FromArgb(15, 141, 33)},   '35
        {"Dark Black", Color.FromArgb(13, 13, 13)},    '36
        {"Dark Teal", Color.FromArgb(20, 26, 26)},    '37
        {"Dark Blue", Color.FromArgb(0, 32, 96)},          '38
        {"Brown", Color.FromArgb(106, 53, 12)},           '39
        {"Gold", Color.FromArgb(255, 204, 0)},               '40
        {"Dark Purple", Color.FromArgb(128, 0, 128)},          '41
        {"Dark Green", Color.FromArgb(10, 94, 22)}}          '42
        '... Add more colors if you have them
        ' }

        ' Generate color palette buttons
        For Each colorEntry In vsColors
            Dim btn As New Button With {
            .Width = 20,
            .Height = 20,
            .BackColor = colorEntry.Value,
            .Tag = colorEntry.Key,  ' Store color name in the Tag property
            .FlatStyle = FlatStyle.Popup
        }

            ' Attach events to the button
            AddHandler btn.Click, AddressOf ColorButton_Click
            AddHandler btn.MouseHover, AddressOf ColorButton_MouseHover

            ' Add the button to the flow layout panel
            FlowLayoutPanel1.Controls.Add(btn)
        Next



        Try

            excelApp = Globals.ThisAddIn.Application

            AddHandler excelApp.SheetSelectionChange, AddressOf excelApp_SheetSelectionChange

            opened = opened + 1

            If excelApp.Selection IsNot Nothing Then
                selectedRange = excelApp.Selection
                src_rng = selectedRange
                TB_src_rng.Text = selectedRange.Address

            End If
            TB_src_rng.Focus()

        Catch ex As Exception
            TB_src_rng.Focus()
        End Try


    End Sub

    Private Sub ColorButton_Click(sender As Object, e As EventArgs)

        Dim clickedButton As Button = CType(sender, Button)
        objectPosition = clickedButton.Location
        Dim c As Color
        'MsgBox(8)
        c = clickedButton.BackColor


        Dim index As Integer = List_Preview.SelectedIndex
        If index < 0 Then Return

        Dim item As ColoredItem = CType(List_Preview.Items(index), ColoredItem)
        item.Color = c
        clickedButton.Focus()

        mybtn(List_Preview.SelectedIndex) = clickedButton
        Btn_color.BackColor = c
        Me.Refresh()
    End Sub
    Private Sub List_Box_IndexChanged() Handles List_Preview.SelectedIndexChanged

        Dim item As ColoredItem = CType(List_Preview.Items(List_Preview.SelectedIndex), ColoredItem)

        If item.Color = Color.White Then
            Btn_NC.Focus()
        Else
            mybtn(List_Preview.SelectedIndex).Focus()
        End If
    End Sub

    Private Sub ListBox1_DrawItem(sender As Object, e As DrawItemEventArgs) Handles List_Preview.DrawItem
        If e.Index < 0 Then Return

        Dim item As ColoredItem = CType(List_Preview.Items(e.Index), ColoredItem)

        Dim textColor As Color = Color.Black
        Dim backColor As Color = item.Color

        If (e.State And DrawItemState.Selected) = DrawItemState.Selected And item.Color = Color.White Then
            ' If item is selected, we'll use system colors to highlight.
            backColor = SystemColors.Highlight
            textColor = SystemColors.HighlightText
        End If

        ' Use the determined colors.
        Using brush As New SolidBrush(backColor)
            e.Graphics.FillRectangle(brush, e.Bounds)
        End Using

        ' Draw the text in the determined text color.
        Using brush As New SolidBrush(textColor)
            e.Graphics.DrawString(item.Text, e.Font, brush, e.Bounds.Left, e.Bounds.Top)
        End Using

        e.DrawFocusRectangle()
    End Sub

    Private Sub ColorButton_MouseHover(sender As Object, e As EventArgs)
        Dim hoveredButton As Button = CType(sender, Button)
        ' Display the color name from the Tag property of the button
        tooltip.SetToolTip(hoveredButton, hoveredButton.Tag.ToString())
    End Sub


    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If ColorDialog1.ShowDialog <> Windows.Forms.DialogResult.Cancel Then
            'Label1.ForeColor = ColorDialog1.Color

            Dim clickedButton As Button = CType(sender, Button)
            objectPosition = clickedButton.Location
            Dim index As Integer = List_Preview.SelectedIndex

            Dim item As ColoredItem = CType(List_Preview.Items(index), ColoredItem)
            item.Color = ColorDialog1.Color
            Button2.Focus()

            mybtn(List_Preview.SelectedIndex) = Button2
            Btn_color.BackColor = ColorDialog1.Color
            Me.Refresh()
        End If
    End Sub

    Private Sub Selection_source_Click(sender As Object, e As EventArgs) Handles Selection_source.Click
        Try
            If selectedRange Is Nothing Then
            Else

                'MsgBox(List_Preview.Items.Count)
                TB_src_rng.Text = selectedRange.Address


                'FocusedTextBox = 1
                Me.Hide()

                excelApp = Globals.ThisAddIn.Application
                workBook = excelApp.ActiveWorkbook

                Dim userInput As Excel.Range = excelApp.InputBox("Select a range", "Select a range", "=$A$1", Type:=8)
                src_rng = userInput

                Dim sheetName As String
                sheetName = Split(src_rng.Address(True, True, Excel.XlReferenceStyle.xlA1, True), "]")(1)
                sheetName = Split(sheetName, "!")(0)

                If Mid(sheetName, Len(sheetName), 1) = "'" Then
                    sheetName = Mid(sheetName, 1, Len(sheetName) - 1)
                End If
                workSheet = workBook.Worksheets(sheetName)
                workSheet.Activate()

                src_rng.Select()

                TB_src_rng.Text = src_rng.Address

                Me.Show()
                TB_src_rng.Focus()

                Dim ran As Excel.Range = src_rng(1, 1)



                ' Clear the ListBox
                List_Preview.Items.Clear()
                'MsgBox(ran.Address)

                'If range.Validation.Type = Excel.XlDVType.xlValidateList Then
                Dim formula As String = ran.Validation.Formula1
                'MsgBox(formula)
                Dim items As New List(Of String)()
                If formula.Contains(":") Then
                    Dim range As Excel.Range = excelApp.Range(formula)
                    For Each r In range
                        items.Add(r.Value.ToString())
                    Next
                Else
                    ' Else, split the formula to get the individual items
                    items.AddRange(formula.Split(","))
                End If


                For Each item As String In items
                    List_Preview.Items.Add(New ColoredItem(item.Trim))
                Next


                'Next

                ReDim mybtn(List_Preview.Items.Count)

            End If

        Catch ex As Exception

            Me.Show()
            TB_src_rng.Focus()

        End Try
    End Sub

    Private Sub Label_Click(sender As Object, e As EventArgs)
        Dim lbl As Windows.Forms.Label = DirectCast(sender, Windows.Forms.Label)
        ' Reset all labels to their default colors
        For Each control As Control In Me.Controls
            If TypeOf control Is Windows.Forms.Label Then
                Dim lbl1 As Windows.Forms.Label = DirectCast(control, Windows.Forms.Label)
                lbl1.BackColor = Color.White
                lbl1.ForeColor = Color.Black

            End If
        Next

        lbl.BackColor = SystemColors.Highlight
        lbl.ForeColor = SystemColors.HighlightText
        Btn_NC.Focus()

    End Sub



    Private Sub Selection_destination_Click(sender As Object, e As EventArgs) Handles Selection_destination.Click
        If selectedRange Is Nothing Then
        Else
            ' TB_src_range.Text = selectedRange.Address


            Me.Hide()

            excelApp = Globals.ThisAddIn.Application
            workBook = excelApp.ActiveWorkbook

            'Dim userInput As String = excelApp.InputBox("Select a range", "Select range", "=$A$1")


            Dim userInput As Excel.Range = excelApp.InputBox("Select a range", "Select a range", "=$A$1", Type:=8)
            des_rng = userInput

            Dim sheetName As String
            sheetName = Split(des_rng.Address(True, True, Excel.XlReferenceStyle.xlA1, True), "]")(1)
            sheetName = Split(sheetName, "!")(0)

            If Mid(sheetName, Len(sheetName), 1) = "'" Then
                sheetName = Mid(sheetName, 1, Len(sheetName) - 1)
            End If

            workSheet = workBook.Worksheets(sheetName)
            workSheet.Activate()

            des_rng.Select()
            'MsgBox(src_rng.Address)

            TB_des_rng.Text = des_rng.Address

            Me.Show()
            TB_des_rng.Focus()

        End If
    End Sub

    Private Sub excelApp_SheetSelectionChange(ByVal Sh As Object, ByVal selectionRange1 As Excel.Range) Handles excelApp.SheetSelectionChange
        Try

            excelApp = Globals.ThisAddIn.Application
            If focuschange = False Then

                If Me.ActiveControl Is TB_des_rng Then
                    des_rng = selectionRange1
                    ' This will run on the Excel thread, so you need to use Invoke to update the UI
                    'Me.BeginInvoke(New System.Action(Sub() TB_dest_range.Text = selectionRange1.Address))
                    Me.Activate()
                    Me.BeginInvoke(New System.Action(Sub()
                                                         TB_des_rng.Text = des_rng.Address
                                                         SetWindowPos(Me.Handle, New IntPtr(HWND_TOPMOST), 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOSIZE)
                                                     End Sub))

                ElseIf Me.ActiveControl Is TB_src_rng Then
                    src_rng = selectionRange1
                    Me.Activate()


                    Me.BeginInvoke(New System.Action(Sub()
                                                         TB_src_rng.Text = src_rng.Address
                                                         SetWindowPos(Me.Handle, New IntPtr(HWND_TOPMOST), 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOSIZE)
                                                     End Sub))
                End If

            End If

        Catch ex As Exception

        End Try

    End Sub

    Protected Overrides Function ProcessCmdKey(ByRef msg As Message, keyData As Keys) As Boolean
        ' Check if Enter is pressed and btn1 is focused
        If keyData = Keys.Enter AndAlso Me.ActiveControl Is Btn_NC Then
            btn_OK.PerformClick() ' Perform the btn2 click operation
            Return True ' The key is handled
        End If
        For Each ctrl As Control In FlowLayoutPanel1.Controls
            If keyData = Keys.Enter AndAlso TypeOf ctrl Is Button Then
                btn_OK.PerformClick() ' Perform the btn2 click operation
                Return True ' The key is handled
            End If
        Next
        Return MyBase.ProcessCmdKey(msg, keyData)
    End Function


    Private Sub btn_OK_Click(sender As Object, e As EventArgs) Handles btn_OK.Click
        excelApp = Globals.ThisAddIn.Application
        Dim workbook As Excel.Workbook = excelApp.ActiveWorkbook
        Dim worksheet As Excel.Worksheet = workbook.ActiveSheet
        'Try
        If des_rng IsNot Nothing Then
            des_rng.FormatConditions.Delete()
        End If
        If RB_Row.Checked = False Then
            des_rng = Nothing
        End If


        If TB_src_rng.Text = "" And RB_cell.Checked = True Then
            MessageBox.Show("Select all necessary options", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TB_src_rng.Focus()
            'Me.Close()
            Exit Sub

        ElseIf TB_src_rng.Text <> "" And IsValidExcelCellReference(TB_src_rng.Text) = False Then
            MessageBox.Show("Select a valid data validation range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TB_src_rng.Focus()
            'Me.Close()
            Exit Sub

        ElseIf TB_src_rng.Text = "" And RB_Row.Checked = True And TB_des_rng.Text = "" Then
            MessageBox.Show("Select all necessary options", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TB_src_rng.Focus()
            'Me.Close()
            Exit Sub


        ElseIf TB_des_rng.Text = "" And RB_Row.Checked = True Then
            MessageBox.Show("Select the Destination Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TB_des_rng.Focus()
            'Me.Close()
            Exit Sub


        ElseIf IsValidExcelCellReference(TB_des_rng.Text) = False And RB_Row.Checked = True Then

            MessageBox.Show("Select a valid range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TB_des_rng.Focus()
            'Me.Close()
            Exit Sub

        ElseIf src_rng.Areas.Count > 1 Then
            MessageBox.Show("Multiple selection is not possible in the Data Validation Range field.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TB_src_rng.Focus()
            Exit Sub




            'ElseIf RB_Row.Checked = True And src_rng.row <> des_rng.row Then
            '    MsgBox("so")
            'ElseIf RB_Row.Checked = True And (src_rng.Row >= des_rng.Row) AndAlso
            '   ((src_rng.Row + src_rng.Rows.Count - 1) <= (des_rng.Row + des_rng.Rows.Count - 1)) AndAlso
            '   (src_rng.Column >= des_rng.Column) AndAlso
            '   ((src_rng.Column + src_rng.Columns.Count - 1) <= (des_rng.Column + des_rng.Columns.Count - 1)) Then



        ElseIf RB_Row.Checked = True Then
            If workSheet3.Name <> des_rng.Worksheet.Name Then
                MessageBox.Show("Please select the range of the same worksheet", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                TB_des_rng.Focus()
                Exit Sub

            ElseIf ((src_rng.Row + src_rng.Rows.Count - 1) < (des_rng.Row + des_rng.Rows.Count - 1)) Or (src_rng.Row <> des_rng.Row) Or excelApp.Intersect(src_rng, des_rng) Is Nothing Then


                MessageBox.Show("Please select the range of the same data table.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                TB_des_rng.Focus()
                'Me.Close()
                Exit Sub

            ElseIf RB_Row.Checked = True AndAlso des_rng.Areas.Count > 1 Then
                MessageBox.Show("Select Case a valid range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                TB_src_rng.Focus()
                Exit Sub
            Else
                GoTo GotoExpression
            End If

        Else

GotoExpression:

            If Backup_sheet.Checked = True Then
                    workSheet.Copy(After:=workBook.Sheets(workSheet.Name))
                End If

                workBook.Sheets(workSheet.Name).Activate

                ' Retrieve data validation items

                ' Clear any existing conditional formats

                Dim formula As String = src_rng.Validation.Formula1
                Dim items As String() = formula.Split(","c)

                src_rng.FormatConditions.Delete()

                If RB_cell.Checked = True Then
                    For Each item As ColoredItem In List_Preview.Items
                        'Only drop-down cell
                        AddColorCondition(src_rng, item.ToString, item.Color)

                    Next

                ElseIf RB_Row.Checked = True Then
                    For Each item As ColoredItem In List_Preview.Items
                        ' Color the destination range
                        Dim i As Integer = 0
                        For Each cell In src_rng
                        i = i + 1
                        If item.Color <> Color.White Then
                            AddColorCondition2(des_rng.Rows(i), cell, item.ToString, item.Color)
                        End If
                    Next
                    Next
                End If


            End If
            src_rng.Select()


        For Each cell In src_rng
            If cell.Validation.Type = Excel.XlDVType.xlValidateList Then
                flag = True
                Exit For
            End If

        Next


        For Each item In List_Preview.Items
                If item.color = Color.White Or item.color = SystemColors.Highlight Then
                    flag2 = False
                Else
                    flag2 = True
                    Exit For
                End If
            Next


            If flag = False And sessionflag2 = True Then
                form = New Form42
                form.Show()
                Me.Hide()

            ElseIf flag2 = False And sessionflag1 = True Then
                form2 = New Form43
                form2.Show()
                Me.Hide()


            Else

                Me.Close()

            End If

            'Me.Show()

            ' Catch ex As Exception
            'If flag = False Then
            '    Me.Hide()
            '    form = New Form42
            '    form.Show()
            '    If form.IsDisposed Or form Is Nothing Then
            '        Me.Show()
            '    End If

            'ElseIf flag2 = False Then
            '    Me.Hide()
            '    form2 = New Form43
            '    form2.Show()
            '    If form2.IsDisposed Or form2 Is Nothing Then
            '        Me.Show()
            '    End If
            'End If
            Me.Close()
        'End Try
    End Sub
    Private Function GetColorForItem(index As Integer) As Color
        ' This function maps an index to a color
        ' You can adjust or expand this as needed
        Dim colors As Color() = {Color.Red, Color.Green, Color.Blue, Color.Yellow, Color.Purple}

        If index >= 0 And index <colors.Length Then
            Return colors(index)
        Else
            Return Color.White
        End If
    End Function

    Private Sub AddColorCondition(targetRange As Excel.Range, value As String, color As Color)
        Dim condition As Excel.FormatCondition = CType(targetRange.FormatConditions.Add(Type:=Excel.XlFormatConditionType.xlCellValue, Operator:=Excel.XlFormatConditionOperator.xlEqual, Formula1:=value), Excel.FormatCondition)
        condition.Interior.Color = ColorTranslator.ToOle(color)
    End Sub

    Private Sub AddColorCondition2(targetRange As Excel.Range, controlCell As Excel.Range, value As String, color As Color)
        If IsNumeric(value) = False Then
            Dim formula As String = "=" & controlCell.Address & " = """ & value & """"
            Dim condition As Excel.FormatCondition = CType(targetRange.FormatConditions.Add(Type:=Excel.XlFormatConditionType.xlExpression, Formula1:=formula), Excel.FormatCondition)
            condition.Interior.Color = ColorTranslator.ToOle(color)
        Else
            Dim formula As String = "=" & controlCell.Address & " = " & value & ""
            Dim condition As Excel.FormatCondition = CType(targetRange.FormatConditions.Add(Type:=Excel.XlFormatConditionType.xlExpression, Formula1:=formula), Excel.FormatCondition)
            condition.Interior.Color = ColorTranslator.ToOle(color)
        End If

    End Sub


    Private Sub RB_Row_CheckedChanged(sender As Object, e As EventArgs) Handles RB_Row.CheckedChanged
        If RB_Row.Checked = True Then
            Selection_destination.Enabled = True
            TB_des_rng.Enabled = True
            TB_des_rng.Focus()
        End If
    End Sub

    Private Sub btn_Cancel_Click(sender As Object, e As EventArgs) Handles btn_Cancel.Click
        Me.Close()
    End Sub

    Private Sub TB_src_rng_TextChanged(sender As Object, e As EventArgs) Handles TB_src_rng.TextChanged
        Try

            excelApp = Globals.ThisAddIn.Application
            Dim workbook As Excel.Workbook = excelApp.ActiveWorkbook
            Dim worksheet As Excel.Worksheet = workbook.ActiveSheet

            TB_src_rng.Focus()


            If TB_src_rng.Text IsNot Nothing And IsValidExcelCellReference(TB_src_rng.Text) = True Then
                focuschange = True

                ' Define the range of cells to read (for example, cells A1 to A10)
                src_rng = excelApp.Range(TB_src_rng.Text)
                src_rng.Select()
                Dim range As Excel.Range = src_rng

                Me.Activate()
                'TB_src_range.Focus()
                TB_src_rng.SelectionStart = TB_src_rng.Text.Length
                focuschange = False



                Dim ran As Excel.Range = src_rng(1, 1)

                ' Clear the ListBox
                List_Preview.Items.Clear()
                'MsgBox(ran.Address)

                'If range.Validation.Type = Excel.XlDVType.xlValidateList Then
                Dim formula As String = ran.Validation.Formula1
                'MsgBox(formula)
                Dim items As New List(Of String)()
                If formula.Contains(":") Then
                    range = excelApp.Range(formula)
                    For Each r In range
                        items.Add(r.Value.ToString())
                    Next
                Else
                    ' Else, split the formula to get the individual items
                    items.AddRange(formula.Split(","))
                End If


                For Each item As String In items
                    'MsgBox(item.ToString)
                    List_Preview.Items.Add(New ColoredItem(item.Trim))
                Next

                TB_src_rng.Focus()
                ReDim mybtn(List_Preview.Items.Count)
                workSheet3 = worksheet

            End If

        Catch ex As Exception
            TB_src_rng.Focus()
        End Try
    End Sub




    'Private Sub Form_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown

    '    Try
    '        If e.KeyCode = Keys.Enter Then

    '            Call btn_OK_Click(sender, e)

    '        End If

    '    Catch ex As Exception

    '    End Try

    'End Sub

    'Private Sub RB_cell_KeyDown(sender As Object, e As KeyEventArgs) Handles RB_cell.KeyDown

    '    Try
    '        If e.KeyCode = Keys.Enter Then

    '            Call btn_OK_Click(sender, e)

    '        End If

    '    Catch ex As Exception

    '    End Try

    'End Sub


    'Private Sub RB_row_KeyDown(sender As Object, e As KeyEventArgs) Handles Btn_NC.KeyDown

    '    Try
    '        If e.KeyCode = Keys.Enter Then

    '            btn_OK.PerformClick()

    '        End If

    '    Catch ex As Exception

    '    End Try

    'End Sub
    'Private Sub Sample_Image_KeyDown(sender As Object, e As KeyEventArgs) Handles Btn_NC.KeyDown

    '    Try
    '        If e.KeyCode = Keys.Enter Then

    '            Call btn_OK_Click(sender, e)

    '        End If

    '    Catch ex As Exception
    '        MsgBox(1)
    '    End Try

    'End Sub


    'Private Sub TextBox_Destination_KeyDown(sender As Object, e As KeyEventArgs) Handles TB_des_rng.KeyDown

    '    Try
    '        If e.KeyCode = Keys.Enter Then

    '            Call btn_OK_Click(sender, e)

    '        End If

    '    Catch ex As Exception

    '    End Try

    'End Sub

    'Private Sub TextBox_Source_KeyDown(sender As Object, e As KeyEventArgs) Handles TB_src_rng.KeyDown

    '    Try
    '        If e.KeyCode = Keys.Enter Then

    '            Call btn_OK_Click(sender, e)

    '        End If

    '    Catch ex As Exception

    '    End Try

    'End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim index As Integer = 0
        For Each item In List_Preview.Items
            item = CType(List_Preview.Items(index), ColoredItem)
            item.Color = Color.White
            index = index + 1
        Next
        Me.Refresh()
    End Sub

    Private Sub Btn_color_Click(sender As Object, e As EventArgs) Handles Btn_color.Click
        Dim clickedButton As Button = CType(sender, Button)
        objectPosition = clickedButton.Location
        Dim c As Color
        'MsgBox(8)
        c = Btn_color.BackColor


        Dim index As Integer = List_Preview.SelectedIndex
        If index < 0 Then Return

        Dim item As ColoredItem = CType(List_Preview.Items(index), ColoredItem)
        item.Color = c
        clickedButton.Focus()

        mybtn(List_Preview.SelectedIndex) = clickedButton
        Btn_color.BackColor = c
        Me.Refresh()
    End Sub

    Private Function IsValidExcelCellReference(cellReference As String) As Boolean

        ' Regular expression pattern for a valid sheet name. This is a simplified version and might not cover all edge cases.
        ' Excel sheet names cannot contain the characters \, /, *, [, ], :, ?, and cannot be 'History'.
        Dim sheetNamePattern As String = "(?i)(?![\/*[\]:?])(?!History)[^\/\[\]*?:\\]+"

        ' Regular expression pattern for a cell reference.
        ' This pattern will match references like A1, $A$1, etc.
        Dim cellPattern As String = "(\$?[A-Z]+\$?[0-9]+)"

        ' Regular expression pattern for an Excel reference.
        ' This pattern will match references like A1:B13, $A$1:$B$13, A1, $B$1, etc.
        Dim singleReferencePattern As String = cellPattern + "(:" + cellPattern + ")?"

        ' Regular expression pattern to allow the sheet name, followed by '!', before the cell reference
        Dim fullPattern As String = "^(" + sheetNamePattern + "!)?(" + singleReferencePattern + ")(," + singleReferencePattern + ")*$"

        ' Create a regex object with the pattern.
        Dim regex As New Regex(fullPattern)

        ' Test the input string against the regex pattern.
        Return regex.IsMatch(cellReference.ToUpper)

    End Function

    'Private Sub TB_dest_range_Enter(sender As Object, e As KeyEventArgs) Handles TB_des_rng.KeyDown
    '    'If Enter key is pressed then check if the text is a valid address
    '    If IsValidExcelCellReference(TB_des_rng.Text) = True And e.KeyCode = Keys.Enter Then
    '        des_rng = excelApp.Range(TB_des_rng.Text)
    '        TB_des_rng.Focus()
    '        des_rng.Select()

    '        Call btn_OK_Click(sender, e)   'OK button click event called

    '        'MsgBox(des_rng.Address)
    '    ElseIf IsValidExcelCellReference(TB_des_rng.Text) = False And e.KeyCode = Keys.Enter Then
    '        MessageBox.Show("Select the valid Destination Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '        TB_des_rng.Text = ""
    '        TB_des_rng.Focus()
    '        'Me.Close()
    '        Exit Sub
    '    End If
    'End Sub

    'Private Sub TB_src_range_Enter(sender As Object, e As KeyEventArgs) Handles TB_src_rng.KeyDown
    '    'If Enter key is pressed then check if the text is a valid address

    '    If IsValidExcelCellReference(TB_src_rng.Text) = True And e.KeyCode = Keys.Enter Then
    '        src_rng = excelApp.Range(TB_src_rng.Text)
    '        TB_src_rng.Focus()
    '        src_rng.Select()

    '        Call btn_OK_Click(sender, e)   'OK button click event called

    '        'MsgBox(des_rng.Address)
    '    ElseIf IsValidExcelCellReference(TB_src_rng.Text) = False And e.KeyCode = Keys.Enter Then
    '        MessageBox.Show("Select the valid Source Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '        TB_src_rng.Text = ""
    '        TB_src_rng.Focus()
    '        'Me.Close()
    '        Exit Sub
    '    End If
    'End Sub
    'Private Sub FlowLayout_KeyDown(sender As Object, e As KeyEventArgs) Handles FlowLayoutPanel1.KeyDown

    '    Try
    '        If e.KeyCode = Keys.Enter Then

    '            Call btn_OK_Click(sender, e)

    '        End If

    '    Catch ex As Exception

    '    End Try

    'End Sub



    'Private Sub Form_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown

    '    Try
    '        If e.KeyCode = Keys.Enter And TB_src_rng.Focus = False And TB_des_rng.Focus = False Then

    '            Call btn_OK_Click(sender, e)

    '        End If

    '    Catch ex As Exception

    '    End Try

    'End Sub

    Private Sub TB_des_rng_TextChanged(sender As Object, e As EventArgs) Handles TB_des_rng.TextChanged
        Try
            excelApp = Globals.ThisAddIn.Application
            Dim workbook As Excel.Workbook = excelApp.ActiveWorkbook
            Dim worksheet As Excel.Worksheet = workbook.ActiveSheet
            'workSheet3 = worksheet

            If TB_des_rng.Text IsNot Nothing And IsValidExcelCellReference(TB_des_rng.Text) = True Then
                focuschange = True
                Try
                    ' Define the range of cells to read (for example, cells A1 to A10)
                    des_rng = excelApp.Range(TB_des_rng.Text)
                    des_rng.Select()
                    'Dim range As Excel.Range = des_rng
                Catch
                    ' Split the string into sheet name and cell address
                    Dim parts As String() = TB_des_rng.Text.Split("!"c)
                    Dim sheetName As String = parts(0)
            Dim cellAddress As String = parts(1)

                    des_rng = excelApp.Range(cellAddress)
                    des_rng.Select()
                End Try

        If worksheet.Name <> workSheet3.Name Then
                    TB_des_rng.Text = worksheet.Name & "!" & des_rng.Address
                    des_rng = excelApp.Range(TB_des_rng.Text)
                End If
                Me.Activate()
                    'TB_src_range.Focus()
                    TB_des_rng.SelectionStart = TB_des_rng.Text.Length
                    focuschange = False
                ax = worksheet.Name
            End If

        Catch ex As Exception
            ax = ""
        End Try
    End Sub


    Private Sub Form33_ColorBasedDropDownList_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        form_flag = False
    End Sub

    Private Sub Form33_ColorBasedDropDownList_Disposed(sender As Object, e As EventArgs) Handles Me.Disposed
        form_flag = False
    End Sub

    Private Sub Form33_ColorBasedDropDownList_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        Me.Focus()
        Me.BringToFront()
        Me.Activate()
        Me.BeginInvoke(New System.Action(Sub()
                                             TB_src_rng.Text = src_rng.Address
                                             SetWindowPos(Me.Handle, New IntPtr(HWND_TOPMOST), 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOSIZE)
                                         End Sub))
        TB_src_rng.Focus()
    End Sub

    Private Sub RB_cell_CheckedChanged(sender As Object, e As EventArgs) Handles RB_cell.CheckedChanged
        If RB_cell.Checked = True Then
            Selection_destination.Enabled = False
            TB_des_rng.Enabled = False
        End If
    End Sub

    Private Sub Btn_NC_Click(sender As Object, e As EventArgs) Handles Btn_NC.Click
        Dim clickedButton As Button = CType(sender, Button)
        objectPosition = clickedButton.Location
        Dim index As Integer = List_Preview.SelectedIndex

        Dim item As ColoredItem = CType(List_Preview.Items(index), ColoredItem)
        item.Color = Color.White
        Button2.Focus()

        mybtn(List_Preview.SelectedIndex) = Button2
        Btn_color.BackColor = Color.White
        Me.Refresh()
    End Sub


End Class

Public Class ColoredItem
    Public Property Text As String
    Public Property Color As Color = Color.White

    Public Sub New(t As String)
        Text = t
    End Sub

    Public Sub New(t As String, c As Color)
        Text = t
        Color = c
    End Sub

    Public Overrides Function ToString() As String
        Return Text
    End Function
End Class
