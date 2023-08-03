Imports System.Windows.Forms
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports Microsoft.Office.Interop

Public Class Form29_Simple_Drop_down_List

    Dim WithEvents excelApp As Excel.Application
    Dim workBook As Excel.Workbook
    Dim workSheet As Excel.Worksheet
    Dim workSheet2 As Excel.Worksheet
    Dim src_rng As Excel.Range
    Dim des_rng As Excel.Range
    Dim selectedRange As Excel.Range
    Private Sub Info_Click(sender As Object, e As EventArgs) Handles Info.Click

    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        ' Clear the list box
        ListBox2.Items.Clear()
        Dim selectedItem As String = ListBox1.SelectedItem.ToString()
        ' Split the string into an array of strings
        Dim items As String() = selectedItem.Split(","c)

        ListBox2.Items.AddRange(items)
        Label7.Visible = True
        Label7.Text = items.Count
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

    End Sub

    Private Sub ComboBox1_TextUpdate(sender As Object, e As EventArgs) Handles ComboBox1.TextUpdate

    End Sub

    Private Sub ComboBox1_MouseClick(sender As Object, e As MouseEventArgs) Handles ComboBox1.MouseClick
        If ComboBox1.Text = "" Then
            'Do nothing
        Else
            ' Clear the list box
            ListBox2.Items.Clear()
            Dim selectedItem As String = ComboBox1.Text
            ' Split the string into an array of strings
            Dim items As String() = selectedItem.Split(","c)

            ListBox2.Items.AddRange(items)
            Label7.Visible = True
            Label7.Text = items.Count
        End If
    End Sub

    Private Sub ComboBox1_Enter(sender As Object, e As EventArgs) Handles ComboBox1.Enter
        If ComboBox1.Text = "" Then
            'Do nothing
        Else
            ' Clear the list box
            ListBox2.Items.Clear()
            Dim selectedItem As String = ComboBox1.Text
            ' Split the string into an array of strings
            Dim items As String() = selectedItem.Split(","c)

            ListBox2.Items.AddRange(items)
            Label7.Visible = True
            Label7.Text = items.Count
        End If
    End Sub

    Private Sub ComboBox1_DragLeave(sender As Object, e As EventArgs) Handles ComboBox1.DragLeave

    End Sub

    Private Sub ComboBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles ComboBox1.KeyPress
        If ComboBox1.Text = "" Then
            'Do nothing
        Else
            ' Clear the list box
            ListBox2.Items.Clear()
            Dim selectedItem As String = ComboBox1.Text
            ' Split the string into an array of strings
            Dim items As String() = selectedItem.Split(","c)

            ListBox2.Items.AddRange(items)
            Label7.Visible = True
            Label7.Text = items.Count
        End If
    End Sub

    Private Sub Btn_OK_Click(sender As Object, e As EventArgs) Handles Btn_OK.Click
        excelApp = Globals.ThisAddIn.Application
        workBook = excelApp.ActiveWorkbook
        workSheet = workBook.ActiveSheet

        'Dim items As String() = {"Item 1", "Item 2", "Item 3"}
        Dim stringItems As New List(Of String)()

        ' Join the items into a comma-separated string
        'Dim formula As String = String.Join(",", ListBox2.Items)
        For Each item As Object In ListBox2.Items
            stringItems.Add(item.ToString())
        Next

        ' Join the string representations into a single string
        Dim items As String = String.Join(", ", stringItems)

        'Dim items As String = String.Join(",", ListBox2.Items.Cast(Of String)().ToArray())


        ' Define the cell that will contain the drop-down list (for example, cell A1)
        Dim range As Excel.Range = des_rng

        ' Delete existing validation rules
        range.Validation.Delete()

        ' Create a new validation rule
        Dim validation As Excel.Validation = range.Validation

        ' Add a drop-down list validation rule
        validation.Delete()
        validation.Add(Excel.XlDVType.xlValidateList, Excel.XlDVAlertStyle.xlValidAlertStop, Excel.XlFormatConditionOperator.xlBetween, items, Type.Missing)
        validation.IgnoreBlank = True
        validation.InCellDropdown = True

        Me.Close()

    End Sub

    Private Sub Btn_Cancel_Click(sender As Object, e As EventArgs) Handles Btn_Cancel.Click
        Me.Close()
    End Sub

    Private Sub Selection_Source_Click(sender As Object, e As EventArgs) Handles Selection_Source.Click
        Me.Hide()

        excelApp = Globals.ThisAddIn.Application
        workBook = excelApp.ActiveWorkbook
        workSheet = workBook.ActiveSheet

        'workSheet.Range("A1").Select()
        'Dim userInput As Excel.Range = excelApp.InputBox("Select a range", Type:=8)
        'Dim userInput As Excel.Range = excelApp.InputBox("Select a range", "Select range", "=$A$1")
        Dim userInput As Excel.Range = excelApp.InputBox("Select a range", "Select range", "=$A$1", Type:=8)

        src_rng = userInput
        'MsgBox(src_rng)

        Dim sheetName As String
        sheetName = Split(src_rng.Address(True, True, Excel.XlReferenceStyle.xlA1, True), "]")(1)
        sheetName = Split(sheetName, "!")(0)
        workSheet = workBook.Worksheets(sheetName)
        workSheet.Activate()

        src_rng.Select()

        TB_src_range.Text = src_rng.Address

        Me.Show()
        TB_src_range.Focus()

        ' Define the range of cells to read (for example, cells A1 to A10)
        Dim range As Excel.Range = src_rng

        ' Clear the ListBox
        ListBox2.Items.Clear()

        ' Iterate over each cell in the range
        For Each cell As Excel.Range In range
            ' Add the cell's value to the ListBox
            ListBox2.Items.Add(cell.Value)
        Next

        Label7.Visible = True
        Label7.Text = ListBox2.Items.Count

    End Sub

    Private Sub Selection_Click(sender As Object, e As EventArgs) Handles Selection.Click
        Try
            TB_src_range = selectedRange
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

        Catch ex As Exception

            Me.Show()
            TB_dest_range.Focus()

        End Try
    End Sub

    Private Sub TB_dest_range_TextChanged(sender As Object, e As EventArgs) Handles TB_dest_range.TextChanged

    End Sub

    Private Sub Form29_Load(sender As Object, e As EventArgs) Handles Me.Load
        TB_src_range.Focus()
        'selectedRange = excelApp.Selection
    End Sub

    Private Sub excelApp_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Excel.Range)

        Try

            excelApp = Globals.ThisAddIn.Application
            Dim selectedRange As Excel.Range
            selectedRange = excelApp.Selection

            TB_dest_range.Text = selectedRange.Address
            workSheet = workBook.ActiveSheet
            src_rng = selectedRange
            TB_dest_range.Focus()


        Catch ex As Exception

        End Try

    End Sub
End Class

