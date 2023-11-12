Imports System.Data
Imports System.Drawing
Imports System.Windows.Forms
Imports Microsoft.Office.Interop

Public Class Form31_2_updated_selection

    Dim WithEvents excelApp As Excel.Application
    Dim workBook As Excel.Workbook
    Public Shared workSheet As Excel.Worksheet
    Dim workSheet2 As Excel.Worksheet
    Private Form As Form31_UpdateDynamicDropdownList

    Private Sub Form31_2_updated_selection_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        PopulateDataGridViewWithExcelData()

        ' Set only the non-checkbox columns to read-only
        For Each column As DataGridViewColumn In DataGridView1.Columns
            If Not TypeOf column Is DataGridViewCheckBoxColumn Then
                column.ReadOnly = True
            End If
        Next

    End Sub

    Private Function CreateTextBoxColumn(bindingName As String, headerText As String) As DataGridViewTextBoxColumn
        Dim column As New DataGridViewTextBoxColumn()
        column.DataPropertyName = bindingName ' This should match the property name of the data you're binding to
        column.HeaderText = headerText
        column.Name = bindingName
        Return column
    End Function

    ' This subroutine would be called to populate your DataGridView, assuming it's named dataGridView1.
    Private Sub PopulateDataGridViewWithExcelData()
        excelApp = Globals.ThisAddIn.Application
        workBook = excelApp.ActiveWorkbook
        workSheet = workBook.ActiveSheet
        Try
            ' Create a new DataTable.
            Dim dataTable As New DataTable()

            ' Define columns for the DataTable to match your Excel structure.
            'dataTable.Columns.Add("DataRange", GetType(String))
            dataTable.Columns.Add("OriginalDataRange", GetType(String))
            dataTable.Columns.Add("OutputRange", GetType(String))
            dataTable.Columns.Add("Level", GetType(Integer))
            dataTable.Columns.Add("Select", GetType(Boolean)) ' For the CheckBox

            Dim targetWorksheet As Excel.Worksheet = Nothing
            For Each ws As Excel.Worksheet In excelApp.Worksheets
                If ws.Name = "MySpecialSheet" Then
                    targetWorksheet = ws
                    Exit For
                End If
            Next

            Dim Label As Int16



            If targetWorksheet.Range("A1").Value <> "" Then
                If targetWorksheet.Range("A7").Value = True Then
                    Label = 5 ' Replace with the label you want if the condition is true
                Else
                    Label = 2 ' Replace with the label you want if the condition is false
                End If

                dataTable.Rows.Add(targetWorksheet.Range("A1").Value, targetWorksheet.Range("A2").Value, Label, False)

            End If

            If targetWorksheet.Range("B1").Value <> "" Then

                If targetWorksheet.Range("B7").Value = True Then
                    Label = 5 ' Replace with the label you want if the condition is true
                Else
                    Label = 2 ' Replace with the label you want if the condition is false
                End If

                dataTable.Rows.Add(targetWorksheet.Range("B1").Value, targetWorksheet.Range("B2").Value, Label, False)

            End If

            If targetWorksheet.Range("C1").Value <> "" Then

                If targetWorksheet.Range("C7").Value = True Then
                    Label = 5 ' Replace with the label you want if the condition is true
                Else
                    Label = 2 ' Replace with the label you want if the condition is false
                End If

                dataTable.Rows.Add(targetWorksheet.Range("C1").Value, targetWorksheet.Range("C2").Value, Label, False)

            End If

            If targetWorksheet.Range("D1").Value <> "" Then

                If targetWorksheet.Range("D7").Value = True Then
                    Label = 5 ' Replace with the label you want if the condition is true
                Else
                    Label = 2 ' Replace with the label you want if the condition is false
                End If

                dataTable.Rows.Add(targetWorksheet.Range("D1").Value, targetWorksheet.Range("D2").Value, Label, False)

            End If
            If targetWorksheet.Range("E1").Value <> "" Then

                If targetWorksheet.Range("E7").Value = True Then
                    Label = 5 ' Replace with the label you want if the condition is true
                Else
                    Label = 2 ' Replace with the label you want if the condition is false
                End If

                dataTable.Rows.Add(targetWorksheet.Range("E1").Value, targetWorksheet.Range("E2").Value, Label, False)

            End If

            ' Set the DataGridView's DataSource to the DataTable.
            DataGridView1.DataSource = dataTable

            ' Adjusting the DataGridView properties for better appearance
            DataGridView1.AutoResizeColumns()
            DataGridView1.Columns("Select").DisplayIndex = 0 ' To show the checkbox column as the first column
        Catch ex As Exception
            MsgBox("Dynamic Drop-down List is not available")
        End Try
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

        ' If the clicked cell is in the checkbox column.
        If e.ColumnIndex = DataGridView1.Columns("Select").Index Then
            DataGridView1.CommitEdit(DataGridViewDataErrorContexts.Commit)
        End If

    End Sub
    ' The event handler for CellValueChanged
    Private Sub dataGridView1_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellValueChanged
        ' Check if the change happened in the checkbox column
        If e.ColumnIndex = DataGridView1.Columns("Select").Index Then
            UpdateRowColor(e.RowIndex)
        End If

    End Sub

    ' Call this method in the CellValueChanged event handler to update the row color.
    Private Sub UpdateRowColor(rowIndex As Integer)
        If rowIndex < 0 Then Exit Sub

        Dim row As DataGridViewRow = DataGridView1.Rows(rowIndex)
        Dim isChecked As Boolean = Convert.ToBoolean(row.Cells("Select").Value)

        If isChecked Then
            row.DefaultCellStyle.BackColor = SystemColors.Highlight
            row.DefaultCellStyle.ForeColor = Color.White
            row.DefaultCellStyle.Font = New Font("Segoe UI", 10)
            'row.DefaultCellStyle.Font = New Font(DataGridView1.DefaultCellStyle.Font, FontStyle.Bold)
        Else
            row.DefaultCellStyle.BackColor = Color.White
            row.DefaultCellStyle.ForeColor = Color.Black
            row.DefaultCellStyle.Font = New Font("Segoe UI", 10)
            'row.DefaultCellStyle.Font = New Font(DataGridView1.DefaultCellStyle.Font, FontStyle.Regular)
        End If
    End Sub

    ' Handle the event when the DataGridView data binding is complete to color initial rows (if necessary)
    Private Sub dataGridView1_DataBindingComplete(sender As Object, e As DataGridViewBindingCompleteEventArgs) Handles DataGridView1.DataBindingComplete

        For Each row As DataGridViewRow In DataGridView1.Rows
            UpdateRowColor(row.Index)
        Next

    End Sub

    ' Handle the CellClick event for the DataGridView.
    Private Sub dataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick

        ' Ignore header clicks or clicks on the checkbox cell itself
        If e.RowIndex < 0 Or e.ColumnIndex = DataGridView1.Columns("Select").Index Then Return

        ' Get the checkbox cell
        Dim checkBoxCell As DataGridViewCheckBoxCell = TryCast(DataGridView1.Rows(e.RowIndex).Cells("Select"), DataGridViewCheckBoxCell)

        If checkBoxCell IsNot Nothing AndAlso Not checkBoxCell.ReadOnly Then
            ' Toggle the checkbox value
            checkBoxCell.Value = Not Convert.ToBoolean(checkBoxCell.Value)
            ' Commit the edit immediately
            DataGridView1.CommitEdit(DataGridViewDataErrorContexts.Commit)
        End If

    End Sub

    Private Sub Btn_OK_Click(sender As Object, e As EventArgs) Handles Btn_OK.Click

        Dim targetWorksheet As Excel.Worksheet
        Dim i As Integer = 1
        For Each ws In excelApp.ActiveWorkbook.Worksheets
            If ws.name = "MySpecialSheet" Then
                targetWorksheet = ws
                Exit For
            End If
        Next

        ' This list will hold the values of the checked rows
        Dim checkedRowsValues As New List(Of String)

        ' Iterate over each row to check if the checkbox is checked
        For Each row As DataGridViewRow In DataGridView1.Rows
            Dim isSelected As Boolean = Convert.ToBoolean(row.Cells("Select").Value) ' Replace "Select" with your checkbox column's name
            If isSelected And i = 1 Then

                Variable1 = targetWorksheet.Range("A1").Value.ToString()
                Variable2 = targetWorksheet.Range("A2").Value.ToString()
                Header = targetWorksheet.Range("A3").Value.ToString()
                Ascending = targetWorksheet.Range("A4").Value.ToString()
                Descending = targetWorksheet.Range("A5").Value.ToString()
                TextConvert = targetWorksheet.Range("A6").Value.ToString()
                OptionType = targetWorksheet.Range("A7").Value.ToString()
                Horizontal_CreateDP = targetWorksheet.Range("A8").Value.ToString()
                Flag_CreateDDDL = targetWorksheet.Range("A9").Value.ToString
                sheetName10 = targetWorksheet.Range("A10").Value.ToString
                sheetName11 = targetWorksheet.Range("A11").Value.ToString
                Form = New Form31_UpdateDynamicDropdownList
                Form.Show()
                Form.TextBox1.Text = i

            ElseIf isSelected And i = 2 Then

                Variable1 = targetWorksheet.Range("B1").Value.ToString()
                Variable2 = targetWorksheet.Range("B2").Value.ToString()
                Header = targetWorksheet.Range("B3").Value.ToString()
                Ascending = targetWorksheet.Range("B4").Value.ToString()
                Descending = targetWorksheet.Range("B5").Value.ToString()
                TextConvert = targetWorksheet.Range("B6").Value.ToString()
                OptionType = targetWorksheet.Range("B7").Value.ToString()
                Horizontal_CreateDP = targetWorksheet.Range("B8").Value.ToString()
                Flag_CreateDDDL = targetWorksheet.Range("B9").Value.ToString
                sheetName10 = targetWorksheet.Range("B10").Value.ToString
                sheetName11 = targetWorksheet.Range("B11").Value.ToString
                Form = New Form31_UpdateDynamicDropdownList
                Form.Show()
                Form.TextBox1.Text = i

            ElseIf isSelected And i = 3 Then
                Variable1 = targetWorksheet.Range("C1").Value.ToString()
                Variable2 = targetWorksheet.Range("C2").Value.ToString()
                Header = targetWorksheet.Range("C3").Value.ToString()
                Ascending = targetWorksheet.Range("C4").Value.ToString()
                Descending = targetWorksheet.Range("C5").Value.ToString()
                TextConvert = targetWorksheet.Range("C6").Value.ToString()
                OptionType = targetWorksheet.Range("C7").Value.ToString()
                Horizontal_CreateDP = targetWorksheet.Range("C8").Value.ToString()
                Flag_CreateDDDL = targetWorksheet.Range("C9").Value.ToString
                sheetName10 = targetWorksheet.Range("C10").Value.ToString
                sheetName11 = targetWorksheet.Range("C11").Value.ToString
                Form = New Form31_UpdateDynamicDropdownList
                Form.Show()
                Form.TextBox1.Text = i

            ElseIf isSelected And i = 4 Then

                Variable1 = targetWorksheet.Range("D1").Value.ToString()
                Variable2 = targetWorksheet.Range("D2").Value.ToString()
                Header = targetWorksheet.Range("D3").Value.ToString()
                Ascending = targetWorksheet.Range("D4").Value.ToString()
                Descending = targetWorksheet.Range("D5").Value.ToString()
                TextConvert = targetWorksheet.Range("D6").Value.ToString()
                OptionType = targetWorksheet.Range("D7").Value.ToString()
                Horizontal_CreateDP = targetWorksheet.Range("D8").Value.ToString()
                Flag_CreateDDDL = targetWorksheet.Range("D9").Value.ToString
                sheetName10 = targetWorksheet.Range("D10").Value.ToString
                sheetName11 = targetWorksheet.Range("D11").Value.ToString
                Form = New Form31_UpdateDynamicDropdownList
                Form.Show()
                Form.TextBox1.Text = i

            ElseIf isSelected And i = 5 Then
                Variable1 = targetWorksheet.Range("E1").Value.ToString()
                Variable2 = targetWorksheet.Range("E2").Value.ToString()
                Header = targetWorksheet.Range("E3").Value.ToString()
                Ascending = targetWorksheet.Range("E4").Value.ToString()
                Descending = targetWorksheet.Range("E5").Value.ToString()
                TextConvert = targetWorksheet.Range("E6").Value.ToString()
                OptionType = targetWorksheet.Range("E7").Value.ToString()
                Horizontal_CreateDP = targetWorksheet.Range("E8").Value.ToString()
                Flag_CreateDDDL = targetWorksheet.Range("E9").Value.ToString
                sheetName10 = targetWorksheet.Range("E10").Value.ToString
                sheetName11 = targetWorksheet.Range("E11").Value.ToString
                Form = New Form31_UpdateDynamicDropdownList
                Form.Show()
                Form.TextBox1.Text = i

            End If
            i = i + 1

        Next
        Me.Close()

    End Sub

    Private Sub Btn_Cancel_Click(sender As Object, e As EventArgs) Handles Btn_Cancel.Click
        Me.Dispose()
    End Sub

End Class