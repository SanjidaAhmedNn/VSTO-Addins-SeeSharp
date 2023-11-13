Imports System.ComponentModel
Imports System.ComponentModel.Design
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices
Imports System.Runtime.Remoting.Contexts
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports Microsoft.Office.Interop

Imports Microsoft.Office.Interop.Excel
Public Class Form31_UpdateDynamicDropdownList
    Dim WithEvents excelApp As Excel.Application
    Dim workBook As Excel.Workbook
    Public Shared workSheet As Excel.Worksheet
    Dim workSheet2 As Excel.Worksheet
    Dim workSheet3 As Excel.Worksheet
    Dim src_rng As Excel.Range
    Public des_rng As Excel.Range
    Dim selectedRange As Excel.Range
    Public ax As String
    Public focuschange As Boolean
    Dim form As Form30_Create_Dynamic_Drop_down_List


    Dim opened As Integer


    Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As IntPtr, ByVal hWndInsertAfter As IntPtr, ByVal X As Integer, ByVal Y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal uFlags As UInteger) As Boolean
    Private Const SWP_NOMOVE As UInteger = &H2
    Private Const SWP_NOSIZE As UInteger = &H1
    Private Const SWP_NOACTIVATE As UInteger = &H10
    Private Const HWND_TOPMOST As Integer = -1


    Private Sub Selection_source_Click(sender As Object, e As EventArgs) Handles Selection_source.Click
        Try
            If selectedRange Is Nothing Then
            Else

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
                Variable1 = TB_src_rng.Text
                'MsgBox(Variable1)
                Me.Show()
                TB_src_rng.Focus()
            End If

        Catch ex As Exception

            Me.Show()
            TB_src_rng.Focus()

        End Try
    End Sub

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click
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

            TB_des_rng2.Text = des_rng.Address

            Me.Show()
            TB_des_rng2.Focus()

        End If
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Dim workbook As Excel.Workbook = excelApp.ActiveWorkbook
        'Dim worksheet As Excel.Worksheet = workbook.ActiveSheet
        Try

            excelApp = Globals.ThisAddIn.Application

            AddHandler excelApp.SheetSelectionChange, AddressOf excelApp_SheetSelectionChange

            opened = opened + 1

            If excelApp.Selection IsNot Nothing Then
                selectedRange = excelApp.Selection
                src_rng = selectedRange
                TB_src_rng.Text = selectedRange.Address
            Else
                selectedRange = excelApp.Range(Variable1)
                src_rng = selectedRange
                TB_src_rng.Text = selectedRange.Address

            End If


            If RB_same_source.Checked = True Then
                TB_des_rng1.Enabled = True
                TB_des_rng2.Enabled = False
                PictureBox3.Enabled = False
                PictureBox2.Enabled = False
                L_select.Enabled = False
                If Variable2 IsNot Nothing Then
                    TB_des_rng1.Text = Variable2
                    des_rng = excelApp.ActiveSheet.Range(TB_des_rng1.Text)
                End If

            ElseIf RB_diff_rng.Checked = True Then
                TB_des_rng1.Enabled = False
                TB_des_rng2.Enabled = True
                PictureBox3.Enabled = True
                PictureBox2.Enabled = True
                L_select.Enabled = True

            End If

        Catch ex As Exception

        End Try

    End Sub


    Private Sub excelApp_SheetSelectionChange(ByVal Sh As Object, ByVal selectionRange1 As Excel.Range) Handles excelApp.SheetSelectionChange
        Try

            excelApp = Globals.ThisAddIn.Application
            If focuschange = False Then

                If Me.ActiveControl Is TB_des_rng2 Then
                    des_rng = selectionRange1
                    ' This will run on the Excel thread, so you need to use Invoke to update the UI
                    'Me.BeginInvoke(New System.Action(Sub() TB_dest_range.Text = selectionRange1.Address))
                    Me.Activate()
                    Me.BeginInvoke(New System.Action(Sub()
                                                         TB_des_rng2.Text = des_rng.Address
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

    Private Sub Btn_Cancel_Click(sender As Object, e As EventArgs) Handles Btn_Cancel.Click
        Me.Close()
    End Sub


    Private Sub Btn_OK_Click(sender As Object, e As EventArgs) Handles Btn_OK.Click
        excelApp = Globals.ThisAddIn.Application
        Dim workbook As Excel.Workbook = excelApp.ActiveWorkbook
        Dim worksheet As Excel.Worksheet = workbook.ActiveSheet

        Dim r1 As Excel.Range
        r1 = workSheet2.Range(TB_src_rng.Text)


        If TB_src_rng.Text = "" And TB_des_rng2.Text = "" And TB_des_rng2.Enabled = True Then
            MessageBox.Show("Please, Select updated source range and destination range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TB_src_rng.Focus()
            'Me.Close()
            Exit Sub

        ElseIf TB_src_rng.Text = "" Then
            'MsgBox(100)
            MessageBox.Show("Check your Updated Source Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TB_src_rng.Focus()
            'Me.Close()
            Exit Sub

        ElseIf IsValidExcelCellReference(TB_src_rng.Text) = False Then
            MessageBox.Show("Select a valid Source Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TB_src_rng.Focus()
            'Me.Close()
            Exit Sub



        ElseIf TB_des_rng2.Enabled = True And TB_des_rng2.Text = "" Then
            MessageBox.Show("Please, Select destination range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TB_src_rng.Focus()
            'Me.Close()
            Exit Sub
        ElseIf TB_des_rng2.Enabled = True And IsValidExcelCellReference(TB_des_rng2.Text) = False Then
            MessageBox.Show("Select a valid Destination Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TB_src_rng.Focus()
            'Me.Close()
            Exit Sub

        ElseIf RB_diff_rng.Checked = False And RB_same_source.Checked = False Then
            MessageBox.Show("Select Destination Range", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TB_des_rng2.Focus()
            'Me.Close()
            Exit Sub

        ElseIf src_rng.Areas.Count > 1 Then
            MessageBox.Show("Please Select dynamic drop-down list range from same worksheet.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TB_src_rng.Focus()

        ElseIf r1.Columns.Count <> des_rng.Columns.Count Then
            MessageBox.Show("Check your Updated Source Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TB_des_rng2.Focus()
        Else
            Try
                Dim result As DialogResult = MessageBox.Show("The Original Source Range is :" & Variable1 & ". AND the Drop-down list is in :" & Variable2 & "Do you want to continue?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

                ' Check if the user clicked 'Yes'
                If result = DialogResult.Yes Then

                    Variable1 = TB_src_rng.Text
                    If des_rng.Worksheet.Name <> src_rng.Worksheet.Name Then
                        Variable1 = src_rng.Worksheet.Name & "!" & TB_src_rng.Text
                        Variable2 = des_rng.Worksheet.Name & "!" & des_rng.Address

                    Else
                        Variable1 = src_rng.Worksheet.Name & "!" & TB_src_rng.Text
                        Variable2 = des_rng.Worksheet.Name & "!" & des_rng.Address
                    End If

                    OutPut()                                                                        'Main Function

                    Dim targetWorksheet As Excel.Worksheet

                    For Each ws In excelApp.ActiveWorkbook.Worksheets
                        If ws.name = "MySpecialSheet" Then
                            targetWorksheet = ws
                            Exit For
                        End If
                    Next
                    If TextBox1.Text = 1 Then

                        targetWorksheet.Range("A1").Value = Variable1
                        targetWorksheet.Range("A2").Value = Variable2
                        targetWorksheet.Range("A10").Value = sheetName10
                        targetWorksheet.Range("A11").Value = sheetName11

                    ElseIf TextBox1.Text = 2 Then
                        targetWorksheet.Range("B1").Value = Variable1
                        targetWorksheet.Range("B2").Value = Variable2
                        targetWorksheet.Range("B10").Value = sheetName10
                        targetWorksheet.Range("B11").Value = sheetName11

                    ElseIf TextBox1.Text = 3 Then
                        targetWorksheet.Range("C1").Value = Variable1
                        targetWorksheet.Range("C2").Value = Variable2
                        targetWorksheet.Range("C10").Value = sheetName10
                        targetWorksheet.Range("C11").Value = sheetName11

                    ElseIf TextBox1.Text = 4 Then
                        targetWorksheet.Range("D1").Value = Variable1
                        targetWorksheet.Range("D2").Value = Variable2
                        targetWorksheet.Range("D10").Value = sheetName10
                        targetWorksheet.Range("D11").Value = sheetName11

                    ElseIf TextBox1.Text = 5 Then
                        targetWorksheet.Range("E1").Value = Variable1
                        targetWorksheet.Range("E2").Value = Variable2
                        targetWorksheet.Range("E10").Value = sheetName10
                        targetWorksheet.Range("E11").Value = sheetName11
                    End If

                End If
                Me.Close()
            Catch ex As Exception
                des_rng.Select()
                Me.Close()
            End Try
        End If

    End Sub

    Private Sub TB_dest_range_Enter(sender As Object, e As KeyEventArgs) Handles TB_des_rng2.KeyDown
        'If Enter key is pressed then check if the text is a valid address
        If IsValidExcelCellReference(TB_des_rng2.Text) = True And e.KeyCode = Keys.Enter Then
            des_rng = excelApp.Range(TB_des_rng2.Text)
            TB_des_rng2.Focus()
            des_rng.Select()

            Call Btn_OK_Click(sender, e)   'OK button click event called

            'MsgBox(des_rng.Address)
        ElseIf IsValidExcelCellReference(TB_des_rng2.Text) = False And e.KeyCode = Keys.Enter Then
            MessageBox.Show("Select the valid Destination Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TB_des_rng2.Text = ""
            TB_des_rng2.Focus()
            'Me.Close()
            Exit Sub
        End If
    End Sub

    Private Sub TB_src_range_Enter(sender As Object, e As KeyEventArgs) Handles TB_src_rng.KeyDown
        'If Enter key is pressed then check if the text is a valid address

        If IsValidExcelCellReference(TB_src_rng.Text) = True And e.KeyCode = Keys.Enter Then
            src_rng = excelApp.Range(TB_src_rng.Text)
            TB_src_rng.Focus()
            src_rng.Select()

            Call Btn_OK_Click(sender, e)   'OK button click event called

            'MsgBox(des_rng.Address)
        ElseIf IsValidExcelCellReference(TB_src_rng.Text) = False And e.KeyCode = Keys.Enter Then
            MessageBox.Show("Select the valid Source Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TB_src_rng.Text = ""
            TB_src_rng.Focus()
            'Me.Close()
            Exit Sub
        End If
    End Sub


    Sub CreateValidationList(cell As Excel.Range, listValues As String)
        With cell.Validation
            .Delete()
            .Add(Type:=Excel.XlDVType.xlValidateList, AlertStyle:=Excel.XlDVAlertStyle.xlValidAlertStop, Operator:=Excel.XlFormatConditionOperator.xlBetween, Formula1:=listValues)
            .ShowInput = True
            .ShowError = True
        End With
    End Sub
    Private Sub OutPut()

        Try
            excelApp = Globals.ThisAddIn.Application
            Dim workbook As Excel.Workbook = excelApp.ActiveWorkbook
            Dim worksheet As Excel.Worksheet = workbook.ActiveSheet


            Dim rng As Excel.Range
            If Header = True Then
                'Dim adjustRange As Excel.Range
                rng = src_rng.Offset(1, 0).Resize(src_rng.Rows.Count - 1, src_rng.Columns.Count)

            Else

                rng = src_rng 'Assuming you have a range from A1 to A100
            End If

            Dim uniqueValues As New List(Of String)

            'Extract unique values from the range
            For Each cell As Excel.Range In rng.Columns(1).Cells
                Dim value As String = cell.Value
                If Not uniqueValues.Contains(value) Then
                    uniqueValues.Add(value)
                End If
            Next

            If Ascending = True Then
                'Sort the list in ascending order
                uniqueValues.Sort()
            ElseIf Descending = True Then
                'Sort the list in ascending order
                uniqueValues.Sort()
                uniqueValues.Reverse()
            End If

            'Create drop-down list at B1 with the unique values
            Dim dropDownRange As Excel.Range = des_rng.Columns(1)
            Dim validation As Excel.Validation = dropDownRange.Validation
            validation.Delete() 'Remove any existing validation
            validation.Add(Excel.XlDVType.xlValidateList, Formula1:=String.Join(",", uniqueValues))
            Dim range1 As Excel.Range = excelApp.Range(TB_des_rng1.Text)
            'Dim range2 As Excel.Range = range1.Rows(1)
            'MsgBox(range1.Address)
            'MsgBox(des_rng.Address)
            If RB_diff_rng.Checked = True And range1.Address(1, 1) <> des_rng.Address(1, 1) Then

                'MsgBox(range1.Address)
                'If des_rng.Rows.Count < range1.Rows.Count Then
                '    Dim difference As Integer = range1.Rows.Count - des_rng.Rows.Count
                '    Dim startRowToDelete As Integer = range1.Rows.Count - difference + 1
                '    Dim endRowToDelete As Integer = range1.Rows.Count
                '    range1.Rows(String.Format("{0}:{1}", startRowToDelete, endRowToDelete)).Delete(Excel.XlDeleteShiftDirection.xlShiftUp)
                '    range1 = range1.Resize(des_rng.Rows.Count, range1.Columns.Count)
                'End If

                'range1.Rows(1).Cut(des_rng)

                'For i As Integer = 1 To des_rng.Rows.Count

                '    range1.Rows(1).Copy(des_rng.Rows(i))
                'Next

                'des_rng.Rows(1).cut(range1.Rows(1))


                form = New Form30_Create_Dynamic_Drop_down_List
                form.TB_src_range.Text = TB_src_rng.Text
                form.TB_dest_range.Text = TB_des_rng2.Text
                If OptionType = True Then
                    form.RB_Dropdown_35_Labels.Checked = True
                End If
                If Header = True Then
                    form.CB_header.Checked = True
                End If
                If Ascending = True Then
                    form.CB_ascending.Checked = True
                End If
                If Descending = True Then
                    form.CB_descending.Checked = True
                End If
                If TextConvert = True Then
                    form.CB_text.Checked = True
                End If
                If Horizontal_CreateDP = True Then
                    form.RB_Horizon.Checked = True
                End If

                form.Btn_OK_Click(form.Btn_OK, New EventArgs())

            End If


            Variable1 = TB_src_rng.Text
            If RB_diff_rng.Checked = True Then
                Variable2 = TB_des_rng2.Text
            End If
            des_rng.Select()

            des_rng.Value = Nothing
            sheetName10 = workSheet2.Name
            If RB_diff_rng.Checked = True Then
                sheetName11 = workSheet3.Name
            End If

        Catch ex As Exception

        End Try
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


    Private Sub RB_same_source_CheckedChanged(sender As Object, e As EventArgs) Handles RB_same_source.CheckedChanged
        excelApp = Globals.ThisAddIn.Application
        If RB_same_source.Checked = True Then

            TB_des_rng1.Enabled = True

            TB_des_rng2.Enabled = False
            PictureBox3.Enabled = False
            PictureBox2.Enabled = False
            L_select.Enabled = False
            'MsgBox(L_select.Enabled)
            If Variable2 IsNot Nothing Then
                TB_des_rng1.Text = Variable2
                'MsgBox(Variable2)
                des_rng = excelApp.Range(Variable2)
            End If
        End If
    End Sub

    Private Sub RB_diff_rng_CheckedChanged(sender As Object, e As EventArgs) Handles RB_diff_rng.CheckedChanged
        If RB_diff_rng.Checked = True Then
            TB_des_rng1.Enabled = False
            TB_des_rng2.Enabled = True
            PictureBox3.Enabled = True
            PictureBox2.Enabled = True
            L_select.Enabled = True
            TB_des_rng2.Focus()

        End If

    End Sub

    Private Sub OK(sender As Object, e As KeyEventArgs) Handles Btn_OK.KeyDown

        Try
            If e.KeyCode = Keys.Enter Then

                Call Btn_OK_Click(sender, e)

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub Cancel(sender As Object, e As KeyEventArgs) Handles Btn_Cancel.KeyDown

        Try
            If e.KeyCode = Keys.Enter Then

                Call Btn_OK_Click(sender, e)

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub Form_load(sender As Object, e As KeyEventArgs) Handles Me.KeyDown

        Try
            If e.KeyCode = Keys.Enter Then

                Call Btn_OK_Click(sender, e)

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub RB_Different(sender As Object, e As KeyEventArgs) Handles RB_diff_rng.KeyDown

        Try
            If e.KeyCode = Keys.Enter Then

                Call Btn_OK_Click(sender, e)

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub RB_same(sender As Object, e As KeyEventArgs) Handles RB_same_source.KeyDown

        Try
            If e.KeyCode = Keys.Enter Then

                Call Btn_OK_Click(sender, e)

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub CustomGroupBox2_Enter(sender As Object, e As EventArgs) Handles CustomGroupBox2.Enter

    End Sub

    Private Sub TB_src_rng_TextChanged(sender As Object, e As EventArgs) Handles TB_src_rng.TextChanged
        excelApp = Globals.ThisAddIn.Application
        workBook = excelApp.ActiveWorkbook
        workSheet = workBook.ActiveSheet

        Try
            If TB_src_rng.Text IsNot Nothing And IsValidExcelCellReference(TB_src_rng.Text) = True Then
                focuschange = True

                ' Define the range of cells to read (for example, cells A1 to A10)
                TB_src_rng.Text = TB_src_rng.Text.ToUpper
                src_rng = excelApp.Range(TB_src_rng.Text)
                src_rng.Select()
                Dim range As Excel.Range = src_rng


                Me.Activate()
                'TB_src_range.Focus()
                TB_src_rng.SelectionStart = TB_src_rng.Text.Length
                focuschange = False
                workSheet2 = workSheet


            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub TB_des_rng2_TextChanged(sender As Object, e As EventArgs) Handles TB_des_rng2.TextChanged

        excelApp = Globals.ThisAddIn.Application
        Dim workbook As Excel.Workbook = excelApp.ActiveWorkbook
        Dim worksheet As Excel.Worksheet = workbook.ActiveSheet
        Try
            If TB_des_rng2.Text IsNot Nothing And IsValidExcelCellReference(TB_des_rng2.Text) = True Then
                focuschange = True

                ' Define the range of cells to read (for example, cells A1 to A10)
                Try
                    TB_des_rng2.Text = TB_des_rng2.Text
                    des_rng = excelApp.Range(TB_des_rng2.Text)
                    des_rng.Select()

                Catch ex As Exception
                    ' Split the string into sheet name and cell address
                    Dim parts As String() = TB_des_rng2.Text.Split("!"c)
                    Dim sheetName As String = parts(0)
                    Dim cellAddress As String = parts(1)

                    des_rng = excelApp.Range(cellAddress)
                    des_rng.Select()

                End Try

                If workSheet2.Name <> worksheet.Name Then
                    TB_des_rng2.Text = worksheet.Name & "!" & des_rng.Address
                    'src_rng = excelApp.Range(TB_src_range.Text)


                End If
                Me.Activate()
                TB_des_rng2.SelectionStart = TB_des_rng2.Text.Length
                focuschange = False
                ax = worksheet.Name
                workSheet3 = worksheet
                'MsgBox(workSheet3.Name)
            End If
        Catch ex As Exception
            ax = ""
            workSheet3 = worksheet
        End Try
    End Sub

    Private Sub Form31_UpdateDynamicDropdownList_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        form_flag = False
    End Sub

    Private Sub Form31_UpdateDynamicDropdownList_Disposed(sender As Object, e As EventArgs) Handles Me.Disposed
        form_flag = False
    End Sub

    Private Sub Form31_UpdateDynamicDropdownList_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        Me.Focus()
        Me.BringToFront()
        Me.Activate()
        Me.BeginInvoke(New System.Action(Sub()
                                             TB_src_rng.Text = src_rng.Address
                                             SetWindowPos(Me.Handle, New IntPtr(HWND_TOPMOST), 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOSIZE)
                                         End Sub))
    End Sub

End Class