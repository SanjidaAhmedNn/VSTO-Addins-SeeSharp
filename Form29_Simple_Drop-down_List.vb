Imports System.Drawing
Imports System.Security.Cryptography
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Windows.Forms
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports Microsoft.Office.Interop



Public Class Form29_Simple_Drop_down_List

    Dim WithEvents excelApp As Excel.Application
    Dim workBook As Excel.Workbook
    Dim workSheet As Excel.Worksheet
    Dim workSheet2 As Excel.Worksheet
    Dim src_rng As Excel.Range
    Public des_rng As Excel.Range
    Dim selectedRange As Excel.Range
    Public focuschange As Boolean = False


    Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As IntPtr, ByVal hWndInsertAfter As IntPtr, ByVal X As Integer, ByVal Y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal uFlags As UInteger) As Boolean
    Private Const SWP_NOMOVE As UInteger = &H2
    Private Const SWP_NOSIZE As UInteger = &H1
    Private Const SWP_NOACTIVATE As UInteger = &H10
    Private Const HWND_TOPMOST As Integer = -1


    Dim opened As Integer
    Private Sub Info_Click(sender As Object, e As EventArgs) Handles Info.Click

    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        ' Clear the list box
        List_Preview.Items.Clear()
        Dim selectedItem As String = ListBox1.SelectedItem.ToString()
        ' Split the string into an array of strings
        Dim items As String() = selectedItem.Split(","c)

        List_Preview.Items.AddRange(items)
        Label7.Visible = True
        Label7.Text = items.Count

    End Sub


    Private Sub ComboBox1_MouseClick(sender As Object, e As MouseEventArgs) Handles ComboBox1.MouseClick
        If ComboBox1.Text = "" Then
            'Do nothing
        Else
            ' Clear the list box
            List_Preview.Items.Clear()
            Dim selectedItem As String = ComboBox1.Text
            ' Split the string into an array of strings
            Dim items As String() = selectedItem.Split(","c)

            List_Preview.Items.AddRange(items)
            Label7.Visible = True
            Label7.Text = items.Count
        End If
    End Sub

    Private Sub ComboBox1_Enter(sender As Object, e As EventArgs) Handles ComboBox1.Enter
        If ComboBox1.Text = "" Then
            'Do nothing
        Else
            ' Clear the list box
            List_Preview.Items.Clear()
            Dim selectedItem As String = ComboBox1.Text
            ' Split the string into an array of strings
            Dim items As String() = selectedItem.Split(","c)

            List_Preview.Items.AddRange(items)
            Label7.Visible = True
            Label7.Text = items.Count
        End If
    End Sub


    Private Sub ComboBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles ComboBox1.KeyPress
        If ComboBox1.Text = "" Then
            'Do nothing
        Else
            ' Clear the list box
            List_Preview.Items.Clear()
            Dim selectedItem As String = ComboBox1.Text
            ' Split the string into an array of strings
            Dim items As String() = selectedItem.Split(","c)

            For i As Integer = 0 To items.Length - 1
                items(i) = items(i).TrimStart()
            Next


            'ComboBox1.Items.AddRange(items)
            List_Preview.Items.AddRange(items)
            Label7.Visible = True
            Label7.Text = items.Count
        End If
    End Sub

    Private Sub ComboBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox1.KeyDown
        ' Check if the key pressed was 'Enter'
        If e.KeyCode = Keys.Enter Then
            AddNewItem(ComboBox1.Text)
        End If
    End Sub

    Private Sub ComboBox1_Leave(sender As Object, e As EventArgs) Handles ComboBox1.Leave
        AddNewItem(ComboBox1.Text)
    End Sub

    Private Sub AddNewItem(item As String)
        ' Check if the item is not already in the ComboBox
        If Not ComboBox1.Items.Contains(item) Then
            ComboBox1.Items.Add(item)
        End If
    End Sub
    Private Sub Selection() Handles ComboBox1.SelectedValueChanged
        If ComboBox1.Text = "" Then
            'Do nothing
        Else
            ' Clear the list box
            List_Preview.Items.Clear()
            Dim selectedItem As String = ComboBox1.Text
            ' Split the string into an array of strings
            Dim items As String() = selectedItem.Split(","c)

            For i As Integer = 0 To items.Length - 1
                items(i) = items(i).TrimStart()
            Next


            'ComboBox1.Items.AddRange(items)
            List_Preview.Items.AddRange(items)
            Label7.Visible = True
            Label7.Text = items.Count
        End If
    End Sub



    Public Sub Btn_OK_Click(sender As Object, e As EventArgs) Handles Btn_OK.Click
        excelApp = Globals.ThisAddIn.Application
        workBook = excelApp.ActiveWorkbook
        workSheet = workBook.ActiveSheet

        If TB_dest_range.Text = "" Then
            MessageBox.Show("Select the Destination Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TB_dest_range.Focus()
            'Me.Close()
            Exit Sub
        ElseIf List_Preview.Items.Count = 0 Then
            MessageBox.Show("Input for Drop-down list is missing.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TB_src_range.Focus()
            Exit Sub
            'ElseIf IsValidExcelCellReference(TB_dest_range.Text) = False Then
            '   MessageBox.Show("Select a Valid Destination Cell.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            '  TB_dest_range.Focus()
            'Me.Close()
            ' Exit Sub
            'ElseIf IsValidExcelCellReference(TB_src_range.Text) = False Then
            '   MessageBox.Show("Select a Valid Source Cell.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            '  TB_src_range.Focus()
            'Me.Close()
            ' Exit Sub
        Else
            Dim stringItems As New List(Of String)()

            For Each item As Object In List_Preview.Items
                stringItems.Add(item.ToString())
            Next

            ' Join the string representations into a single string
            Dim items As String = String.Join(", ", stringItems)

            des_rng.Validation.Delete()

            ' Create a new validation rule
            Dim validation As Excel.Validation = des_rng.Validation

            ' Add a drop-down list validation rule
            validation.Delete()
            validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, items, Type.Missing)
            validation.IgnoreBlank = True
            validation.InCellDropdown = True

            des_rng.Select()

            Me.Close()
        End If
    End Sub

    Private Sub Btn_Cancel_Click(sender As Object, e As EventArgs) Handles Btn_Cancel.Click
        Me.Close()
    End Sub

    Private Sub Selection_Source_Click(sender As Object, e As EventArgs) Handles Selection_Source.Click
        If selectedRange Is Nothing Then
            TB_src_range.Focus()
        Else
            ' TB_src_range.Text = selectedRange.Address


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
            'MsgBox(src_rng.Address)

            TB_src_range.Text = src_rng.Address

            Me.Show()
            TB_src_range.Focus()
            TB_src_range.Focus()

            ' Define the range of cells to read (for example, cells A1 to A10)
            Dim range As Excel.Range = src_rng

            ' Clear the ListBox
            List_Preview.Items.Clear()

            ' Iterate over each cell in the range
            For Each cell As Excel.Range In range
                ' Add the cell's value to the ListBox
                If cell.Value IsNot Nothing Then
                    List_Preview.Items.Add(cell.Value)
                End If
            Next

            Label7.Visible = True
            Label7.Text = List_Preview.Items.Count
            TB_src_range.Focus()
            TB_src_range.Focus()
            'Me.Activate()

        End If

    End Sub


    Private Sub Selection_Click(sender As Object, e As EventArgs) Handles Selection_destination.Click
        Try
            If selectedRange Is Nothing Then
                TB_dest_range.Focus()
            Else

                TB_dest_range.Text = selectedRange.Address


                'FocusedTextBox = 1
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

                TB_dest_range.Text = des_rng.Address

                Me.Show()
                TB_dest_range.Focus()
                TB_dest_range.Focus()
            End If

        Catch ex As Exception

            Me.Show()
            TB_dest_range.Focus()

        End Try
    End Sub



    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load

        Try

            excelApp = Globals.ThisAddIn.Application

            AddHandler excelApp.SheetSelectionChange, AddressOf excelApp_SheetSelectionChange

            opened = opened + 1

            If excelApp.Selection IsNot Nothing Then
                selectedRange = excelApp.Selection
                des_rng = selectedRange
                TB_dest_range.Text = selectedRange.Address
            End If

            If RadioButton1.Checked = True Then
                ComboBox1.Enabled = False
                ListBox1.Enabled = False
                TB_src_range.Enabled = True
                Selection_Source.Enabled = True
            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub ListBox1_DrawItem(sender As Object, e As DrawItemEventArgs) Handles ListBox1.DrawItem
        ' If the index is invalid, exit
        If e.Index < 0 Then Exit Sub
        Dim backColor As Color

        ' Determine the color based on even or odd index
        If e.Index Mod 2 = 0 Then
            ' Odd lines
            e.Graphics.FillRectangle(Brushes.White, e.Bounds)
            backColor = Color.White
        Else
            ' Even lines
            e.Graphics.FillRectangle(Brushes.LightGray, e.Bounds)
            backColor = Color.LightGray
        End If

        Dim textColor As Color = Color.Black


        If (e.State And DrawItemState.Selected) = DrawItemState.Selected Then
            ' If item is selected, we'll use system colors to highlight.
            BackColor = SystemColors.Highlight
            textColor = SystemColors.HighlightText
        End If

        ' Draw the text
        ' e.Graphics.DrawString(ListBox1.Items(e.Index).ToString(), e.Font, Brushes.Black, e.Bounds)
        Using brush As New SolidBrush(backColor)
            e.Graphics.FillRectangle(brush, e.Bounds)
        End Using

        Using brush As New SolidBrush(textColor)
            e.Graphics.DrawString(ListBox1.Items(e.Index).ToString(), e.Font, brush, e.Bounds.Left, e.Bounds.Top)
        End Using

        ' If the ListBox has focus, draw a focus rectangle around the selected item.
        e.DrawFocusRectangle()

    End Sub

    'Private Sub ListBox12_DrawItem(sender As Object, e As DrawItemEventArgs) Handles ListBox1.DrawItem
    '    If e.Index < 0 Then Return

    '    '  Dim item As ColoredItem = CType(ListBox1.Items(e.Index), ColoredItem)

    '    Dim textColor As Color = Color.Black
    '    Dim backColor As Color = Color.White

    '    If (e.State And DrawItemState.Selected) = DrawItemState.Selected Then
    '        ' If item is selected, we'll use system colors to highlight.
    '        backColor = SystemColors.Highlight
    '        textColor = SystemColors.HighlightText
    '    End If

    '    ' Use the determined colors.
    '    Using brush As New SolidBrush(backColor)
    '        e.Graphics.FillRectangle(brush, e.Bounds)
    '    End Using

    '    ' Draw the text in the determined text color.
    '    Using brush As New SolidBrush(textColor)
    '        e.Graphics.DrawString(ListBox1.Items(e.Index).ToString(), e.Font, brush, e.Bounds.Left, e.Bounds.Top)
    '    End Using

    '    e.DrawFocusRectangle()
    'End Sub


    Private Sub excelApp_SheetSelectionChange(ByVal Sh As Object, ByVal selectionRange1 As Excel.Range) Handles excelApp.SheetSelectionChange
        Try

            excelApp = Globals.ThisAddIn.Application

            '  If Me.ActiveControl Is TB_dest_range Then
            If focuschange = False Then
                If TB_dest_range.Focused = True Or Me.ActiveControl Is TB_dest_range Then
                    If TB_dest_range.Focused = True Then
                        des_rng = selectionRange1
                    End If
                    Me.Activate()
                    Me.BeginInvoke(New System.Action(Sub()
                                                         TB_dest_range.Text = des_rng.Address
                                                         SetWindowPos(Me.Handle, New IntPtr(HWND_TOPMOST), 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOSIZE)
                                                     End Sub))

                    ' ElseIf Me.ActiveControl Is TB_src_range Then
                ElseIf TB_src_range.Focused = True Or Me.ActiveControl Is TB_src_range Then
                    If TB_src_range.Focused = True Then
                        src_rng = selectionRange1
                    End If
                    Me.Activate()
                    Me.BeginInvoke(New System.Action(Sub()
                                                         TB_src_range.Text = src_rng.Address
                                                         SetWindowPos(Me.Handle, New IntPtr(HWND_TOPMOST), 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOSIZE)
                                                     End Sub))

                End If
            End If



        Catch ex As Exception

        End Try

    End Sub

    Private Sub TB_src_range_TextChanged(sender As Object, e As EventArgs) Handles TB_src_range.TextChanged
        Try

            If TB_src_range.Text IsNot Nothing And IsValidExcelCellReference(TB_src_range.Text) = True Then
                focuschange = True

                ' Define the range of cells to read (for example, cells A1 to A10)
                src_rng = excelApp.Range(TB_src_range.Text)
                src_rng.Select()
                Dim range As Excel.Range = src_rng

                ' Clear the ListBox
                List_Preview.Items.Clear()

                ' Iterate over each cell in the range
                For Each cell As Excel.Range In range
                    ' Add the cell's value to the ListBox
                    If cell.Value IsNot Nothing Then
                        List_Preview.Items.Add(cell.Value)
                    End If
                Next

                Label7.Visible = True
                Label7.Text = List_Preview.Items.Count
                Me.Activate()
                'TB_src_range.Focus()
                TB_src_range.SelectionStart = TB_src_range.Text.Length
                focuschange = False

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub TB_dest_rane_TextChanged(sender As Object, e As EventArgs) Handles TB_dest_range.TextChanged
        Try

            If TB_dest_range.Text IsNot Nothing And IsValidExcelCellReference(TB_dest_range.Text) = True Then
                focuschange = True

                ' Define the range of cells to read (for example, cells A1 to A10)
                des_rng = excelApp.Range(TB_dest_range.Text)
                des_rng.Select()
                Dim range As Excel.Range = des_rng

                ' Clear the ListBox
                'List_Preview.Items.Clear()

                '' Iterate over each cell in the range
                'For Each cell As Excel.Range In range
                '    ' Add the cell's value to the ListBox
                '    If cell.Value IsNot Nothing Then
                '        List_Preview.Items.Add(cell.Value)
                '    End If
                'Next

                'Label7.Visible = True
                'Label7.Text = List_Preview.Items.Count
                Me.Activate()
                'TB_src_range.Focus()
                TB_dest_range.SelectionStart = TB_dest_range.Text.Length
                focuschange = False

            End If

        Catch ex As Exception

        End Try
    End Sub


    Private Sub form(sender As Object, e As KeyEventArgs) Handles Me.KeyDown

        Try
            If e.KeyCode = Keys.Enter Then

                Call Btn_OK_Click(sender, e)

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub Listbox(sender As Object, e As KeyEventArgs) Handles ListBox1.KeyDown

        Try
            If e.KeyCode = Keys.Enter Then

                Call Btn_OK_Click(sender, e)

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub Listboxx2(sender As Object, e As KeyEventArgs) Handles List_Preview.KeyDown

        Try
            If e.KeyCode = Keys.Enter Then

                Call Btn_OK_Click(sender, e)

            End If

        Catch ex As Exception

        End Try

    End Sub


    Private Sub destination(sender As Object, e As KeyEventArgs) Handles Selection_destination.KeyDown

        Try
            If e.KeyCode = Keys.Enter Then

                Call Btn_OK_Click(sender, e)

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub source(sender As Object, e As KeyEventArgs) Handles Selection_Source.KeyDown

        Try
            If e.KeyCode = Keys.Enter Then

                Call Btn_OK_Click(sender, e)

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub TB_dest(sender As Object, e As KeyEventArgs) Handles TB_dest_range.KeyDown

        'Try
        '    If e.KeyCode = Keys.Enter Then

        '        Call Btn_OK_Click(sender, e)

        '    End If

        'Catch ex As Exception

        'End Try

    End Sub


    Private Sub TB_src(sender As Object, e As KeyEventArgs) Handles TB_src_range.KeyDown

        'Try
        '    If e.KeyCode = Keys.Enter Then

        '        Call Btn_OK_Click(sender, e)

        '    End If

        'Catch ex As Exception

        'End Try

    End Sub

    Private Function IsValidExcelCellReference(cellReference As String) As Boolean

        ' Regular expression pattern for a cell reference.
        ' This pattern will match references like A1, $A$1, etc.
        Dim cellPattern As String = "(\$?[A-Z]+\$?[0-9]+)"

        ' Regular expression pattern for an Excel reference.
        ' This pattern will match references like A1:B13, $A$1:$B$13, A1, $B$1, etc.
        Dim referencePattern As String = "^" + cellPattern + "(:" + cellPattern + ")?$"

        ' Create a regex object with the pattern.
        Dim regex As New Regex(referencePattern)

        ' Test the input string against the regex pattern.
        If regex.IsMatch(cellReference) Then
            Return True
        Else
            Return False
        End If


    End Function

    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged
        If RadioButton3.Checked = True Then
            ComboBox1.Enabled = True
            ListBox1.Enabled = False
            TB_src_range.Enabled = False
            Selection_Source.Enabled = False

        End If
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        If RadioButton2.Checked = True Then
            ComboBox1.Enabled = False
            ListBox1.Enabled = True
            TB_src_range.Enabled = False
            Selection_Source.Enabled = False
        End If
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked = True Then
            ComboBox1.Enabled = False
            ListBox1.Enabled = False
            TB_src_range.Enabled = True
            Selection_Source.Enabled = True
        End If
    End Sub



    Private Sub TB_dest_range_Enter(sender As Object, e As KeyEventArgs) Handles TB_dest_range.KeyDown
        'If Enter key is pressed then check if the text is a valid address
        If IsValidExcelCellReference(TB_dest_range.Text) = True And e.KeyCode = Keys.Enter Then
            des_rng = excelApp.Range(TB_dest_range.Text)
            TB_dest_range.Focus()
            des_rng.Select()

            Call Btn_OK_Click(sender, e)   'OK button click event called

            'MsgBox(des_rng.Address)
        ElseIf IsValidExcelCellReference(TB_dest_range.Text) = False And e.KeyCode = Keys.Enter Then
            MessageBox.Show("Select the valid Destination Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TB_dest_range.Text = ""
            TB_dest_range.Focus()
            'Me.Close()
            Exit Sub
        End If
    End Sub

    Private Sub TB_src_range_Enter(sender As Object, e As KeyEventArgs) Handles TB_src_range.KeyDown
        'If Enter key is pressed then check if the text is a valid address

        If IsValidExcelCellReference(TB_src_range.Text) = True And e.KeyCode = Keys.Enter Then
            src_rng = excelApp.Range(TB_src_range.Text)
            TB_src_range.Focus()
            src_rng.Select()

            Call Btn_OK_Click(sender, e)   'OK button click event called

            'MsgBox(des_rng.Address)
        ElseIf IsValidExcelCellReference(TB_src_range.Text) = False And e.KeyCode = Keys.Enter Then
            MessageBox.Show("Select the valid Source Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TB_src_range.Text = ""
            TB_src_range.Focus()
            'Me.Close()
            Exit Sub
        End If
    End Sub

    Private Sub ComboBox1_TextUpdate(sender As Object, e As EventArgs) Handles ComboBox1.TextUpdate
        If ComboBox1.Text = "" Then
            'Do nothing
        Else
            ' Clear the list box
            List_Preview.Items.Clear()
            Dim selectedItem As String = ComboBox1.Text
            ' Split the string into an array of strings
            Dim items As String() = selectedItem.Split(","c)

            For i As Integer = 0 To items.Length - 1
                items(i) = items(i).TrimStart()
            Next


            'ComboBox1.Items.AddRange(items)
            List_Preview.Items.AddRange(items)
            Label7.Visible = True
            Label7.Text = items.Count
        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

    End Sub
End Class

Public Class ColoredItem1
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


