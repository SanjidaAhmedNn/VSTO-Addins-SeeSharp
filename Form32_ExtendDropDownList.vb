Imports System.ComponentModel
Imports System.Runtime.InteropServices
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports Microsoft.Office.Interop

Imports Microsoft.Office.Interop.Excel
Public Class Form32_ExtendDropDownList

    Dim WithEvents excelApp As Excel.Application
    Dim workBook As Excel.Workbook
    Public Shared workSheet As Excel.Worksheet
    Dim workSheet2 As Excel.Worksheet
    Dim worksheet3 As Excel.Worksheet
    Dim src_rng As Excel.Range
    Public des_rng As Excel.Range
    Dim selectedRange As Excel.Range
    Public ax As String
    Public firstRow As Excel.Range

    Dim opened As Integer
    Public focuschange As Boolean

    Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As IntPtr, ByVal hWndInsertAfter As IntPtr, ByVal X As Integer, ByVal Y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal uFlags As UInteger) As Boolean
    Private Const SWP_NOMOVE As UInteger = &H2
    Private Const SWP_NOSIZE As UInteger = &H1
    Private Const SWP_NOACTIVATE As UInteger = &H10
    Private Const HWND_TOPMOST As Integer = -1
    Private Sub Selection_source_Click(sender As Object, e As EventArgs) Handles Source_selection.Click
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

                Me.Show()
                TB_src_rng.Focus()

                firstRow = src_rng.Rows(1)
                'MsgBox(firstRow.Address)
            End If

        Catch ex As Exception

            Me.Show()
            TB_src_rng.Focus()

        End Try
    End Sub

    Private Sub Form1_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Enter Then
            Btn_OK.Focus()
            Btn_OK.PerformClick()
        End If
    End Sub

    Private Sub Dest_selection_Click(sender As Object, e As EventArgs) Handles Dest_selection.Click
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

    Private Sub Form32_ExtendDropDownList_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
        Catch ex As Exception
        End Try
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

    Private Sub Btn_OK_Click(sender As Object, e As EventArgs) Handles Btn_OK.Click

        excelApp = Globals.ThisAddIn.Application
        Dim workbook As Excel.Workbook = excelApp.ActiveWorkbook
        Dim worksheet As Excel.Worksheet = workbook.ActiveSheet

        If TB_src_rng.Text = "" And TB_des_rng.Text = "" Then
            MessageBox.Show("Please select all necessary options.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TB_src_rng.Focus()
            'Me.Close()
            Exit Sub

        ElseIf TB_src_rng.Text = "" Then
            MessageBox.Show("Please, Select updated source range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TB_src_rng.Focus()
            'Me.Close()
            Exit Sub

        ElseIf IsValidExcelCellReference(TB_src_rng.Text) = False Then
            MessageBox.Show("Select a valid Source Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TB_src_rng.Focus()
            'Me.Close()
            Exit Sub



        ElseIf TB_des_rng.Text = "" Then
            MessageBox.Show("Please, Select destination range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TB_src_rng.Focus()
            'Me.Close()
            Exit Sub
        ElseIf IsValidExcelCellReference(TB_des_rng.Text) = False Then
            MessageBox.Show("Select a valid Destination Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TB_src_rng.Focus()
            'Me.Close()
            Exit Sub


        ElseIf src_rng.Areas.Count > 1 Or des_rng.Areas.Count > 1 Then
            MessageBox.Show("Multiple selection is not possible in the Source Range field.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TB_src_rng.Focus()
            Exit Sub


        ElseIf ax <> workSheet2.Name Then
            MessageBox.Show("Please select the range of the same worksheet", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TB_des_rng.Focus()
            Exit Sub

            'ElseIf src_rng.Column <> des_rng.Column Then
            '    MessageBox.Show("1st column of source range and destination range should be same.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            '    TB_des_rng.Focus()
            '    Exit Sub

        ElseIf excelApp.Intersect(src_rng, des_rng) Is Nothing Then
            MessageBox.Show(" Please select a valid expanded dynamic drop-down list range that intersects each other.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TB_src_rng.Focus()
            'Me.Close()
            Exit Sub

        Else
            Try
                Dim targetWorksheet As Excel.Worksheet
                ' Dim i As Integer = 1
                For Each ws In excelApp.ActiveWorkbook.Worksheets
                    If ws.name = "MySpecialSheet" Then
                        targetWorksheet = ws
                        Exit For
                    End If
                Next

                Dim k As Integer = 0
                'For i = 1 To targetWorksheet.Columns.Count
                If worksheet.Name = targetWorksheet.Range("A11").Value And excelApp.Intersect(src_rng, excelApp.Range(targetWorksheet.Range("A2").Value)) IsNot Nothing Then
                    Variable1 = targetWorksheet.Range("A1").Value.ToString()
                    Variable2 = targetWorksheet.Range("A2").Value.ToString()
                    Header = targetWorksheet.Range("A3").Value.ToString()
                    Ascending = targetWorksheet.Range("A4").Value.ToString()
                    Descending = targetWorksheet.Range("A5").Value.ToString()
                    TextConvert = targetWorksheet.Range("A6").Value.ToString()
                    OptionType = targetWorksheet.Range("A7").Value.ToString()
                    Horizontal_CreateDP = targetWorksheet.Range("A8").Value.ToString()
                    'Flag_CreateDDDL = targetWorksheet.Range("A9").Value.ToString
                    sheetName10 = targetWorksheet.Range("A10").Value.ToString
                    sheetName11 = targetWorksheet.Range("A11").Value.ToString
                    k = 1

                ElseIf worksheet.Name = targetWorksheet.Range("B11").Value And excelApp.Intersect(src_rng, excelApp.Range(targetWorksheet.Range("B2").Value)) IsNot Nothing Then
                    Variable1 = targetWorksheet.Range("B1").Value.ToString()
                    Variable2 = targetWorksheet.Range("B2").Value.ToString()
                    Header = targetWorksheet.Range("B3").Value.ToString()
                    Ascending = targetWorksheet.Range("B4").Value.ToString()
                    Descending = targetWorksheet.Range("B5").Value.ToString()
                    TextConvert = targetWorksheet.Range("B6").Value.ToString()
                    OptionType = targetWorksheet.Range("B7").Value.ToString()
                    Horizontal_CreateDP = targetWorksheet.Range("B8").Value.ToString()
                    'Flag_CreateDDDL = targetWorksheet.Range("B9").Value.ToString
                    sheetName10 = targetWorksheet.Range("B10").Value.ToString
                    sheetName11 = targetWorksheet.Range("B11").Value.ToString
                    k = 2

                ElseIf worksheet.Name = targetWorksheet.Range("C11").Value And excelApp.Intersect(src_rng, excelApp.Range(targetWorksheet.Range("C2").Value)) IsNot Nothing Then
                    Variable1 = targetWorksheet.Range("C1").Value.ToString()
                    Variable2 = targetWorksheet.Range("C2").Value.ToString()
                    Header = targetWorksheet.Range("C3").Value.ToString()
                    Ascending = targetWorksheet.Range("C4").Value.ToString()
                    Descending = targetWorksheet.Range("C5").Value.ToString()
                    TextConvert = targetWorksheet.Range("C6").Value.ToString()
                    OptionType = targetWorksheet.Range("C7").Value.ToString()
                    Horizontal_CreateDP = targetWorksheet.Range("C8").Value.ToString()
                    'Flag_CreateDDDL = targetWorksheet.Range("C9").Value.ToString
                    sheetName10 = targetWorksheet.Range("C10").Value.ToString
                    sheetName11 = targetWorksheet.Range("C11").Value.ToString
                    k = 3

                ElseIf worksheet.Name = targetWorksheet.Range("D11").Value And excelApp.Intersect(src_rng, excelApp.Range(targetWorksheet.Range("D2").Value)) IsNot Nothing Then
                    Variable1 = targetWorksheet.Range("D1").Value.ToString()
                    Variable2 = targetWorksheet.Range("D2").Value.ToString()
                    Header = targetWorksheet.Range("D3").Value.ToString()
                    Ascending = targetWorksheet.Range("D4").Value.ToString()
                    Descending = targetWorksheet.Range("D5").Value.ToString()
                    TextConvert = targetWorksheet.Range("D6").Value.ToString()
                    OptionType = targetWorksheet.Range("D7").Value.ToString()
                    Horizontal_CreateDP = targetWorksheet.Range("D8").Value.ToString()
                    'Flag_CreateDDDL = targetWorksheet.Range("D9").Value.ToString
                    sheetName10 = targetWorksheet.Range("D10").Value.ToString
                    sheetName11 = targetWorksheet.Range("D11").Value.ToString
                    k = 4

                ElseIf worksheet.Name = targetWorksheet.Range("E11").Value And excelApp.Intersect(src_rng, excelApp.Range(targetWorksheet.Range("E2").Value)) IsNot Nothing Then
                    Variable1 = targetWorksheet.Range("E1").Value.ToString()
                    Variable2 = targetWorksheet.Range("E2").Value.ToString()
                    Header = targetWorksheet.Range("E3").Value.ToString()
                    Ascending = targetWorksheet.Range("E4").Value.ToString()
                    Descending = targetWorksheet.Range("E5").Value.ToString()
                    TextConvert = targetWorksheet.Range("E6").Value.ToString()
                    OptionType = targetWorksheet.Range("E7").Value.ToString()
                    Horizontal_CreateDP = targetWorksheet.Range("E8").Value.ToString()
                    'Flag_CreateDDDL = targetWorksheet.Range("E9").Value.ToString
                    sheetName10 = targetWorksheet.Range("E10").Value.ToString
                    sheetName11 = targetWorksheet.Range("E11").Value.ToString
                    k = 5

                End If


                ' Get the validation formula from the source cell
                Dim validationFormula As String = des_rng(1, 1).Validation.Formula1

                ' Apply the validation to the target range
                With des_rng.Columns(1).Validation
                    .Delete() ' Clear any existing validation
                    .Add(Type:=Excel.XlDVType.xlValidateList,
                     AlertStyle:=Excel.XlDVAlertStyle.xlValidAlertStop,
                     Operator:=Excel.XlFormatConditionOperator.xlBetween,
                     Formula1:=validationFormula)
                    .IgnoreBlank = True
                    .InCellDropdown = True
                    .ShowInput = True
                    .ShowError = True
                End With

                If k = 1 Then

                    targetWorksheet.Range("A2").Value = excelApp.Union(worksheet.Range(targetWorksheet.Range("A2").Value), des_rng).Address
                    ' Header = targetWorksheet.Range("A3").Value.ToString()
                    'Ascending = targetWorksheet.Range("A4").Value.ToString()
                    'Descending = targetWorksheet.Range("A5").Value.ToString()
                    'TextConvert = targetWorksheet.Range("A6").Value.ToString()
                    'OptionType = targetWorksheet.Range("A7").Value.ToString()
                    'Horizontal_CreateDP = targetWorksheet.Range("A8").Value.ToString()
                    'Flag_CreateDDDL = targetWorksheet.Range("A9").Value.ToString
                    'sheetName10 = targetWorksheet.Range("A10").Value.ToString
                    'sheetName11 = targetWorksheet.Range("A11").Value.ToString

                ElseIf k = 2 Then
                    targetWorksheet.Range("B2").Value = excelApp.Union(worksheet.Range(targetWorksheet.Range("B2").Value), des_rng).Address
                    'Header = targetWorksheet.Range("B3").Value.ToString()
                    'Ascending = targetWorksheet.Range("B4").Value.ToString()
                    'Descending = targetWorksheet.Range("B5").Value.ToString()
                    'TextConvert = targetWorksheet.Range("B6").Value.ToString()
                    'OptionType = targetWorksheet.Range("B7").Value.ToString()
                    'Horizontal_CreateDP = targetWorksheet.Range("B8").Value.ToString()
                    'Flag_CreateDDDL = targetWorksheet.Range("B9").Value.ToString
                    'sheetName10 = targetWorksheet.Range("B10").Value.ToString
                    'sheetName11 = targetWorksheet.Range("B11").Value.ToString

                ElseIf k = 3 Then
                    targetWorksheet.Range("C2").Value = excelApp.Union(worksheet.Range(targetWorksheet.Range("C2").Value), des_rng).Address
                    'Header = targetWorksheet.Range("C3").Value.ToString()
                    'Ascending = targetWorksheet.Range("C4").Value.ToString()
                    'Descending = targetWorksheet.Range("C5").Value.ToString()
                    'TextConvert = targetWorksheet.Range("C6").Value.ToString()
                    'OptionType = targetWorksheet.Range("C7").Value.ToString()
                    'Horizontal_CreateDP = targetWorksheet.Range("C8").Value.ToString()
                    'Flag_CreateDDDL = targetWorksheet.Range("C9").Value.ToString
                    'sheetName10 = targetWorksheet.Range("C10").Value.ToString
                    'sheetName11 = targetWorksheet.Range("C11").Value.ToString

                ElseIf k = 4 Then
                    targetWorksheet.Range("D2").Value = excelApp.Union(worksheet.Range(targetWorksheet.Range("D2").Value), des_rng).Address
                    'Header = targetWorksheet.Range("D3").Value.ToString()
                    'Ascending = targetWorksheet.Range("D4").Value.ToString()
                    'Descending = targetWorksheet.Range("D5").Value.ToString()
                    'TextConvert = targetWorksheet.Range("D6").Value.ToString()
                    'OptionType = targetWorksheet.Range("D7").Value.ToString()
                    'Horizontal_CreateDP = targetWorksheet.Range("D8").Value.ToString()
                    'Flag_CreateDDDL = targetWorksheet.Range("D9").Value.ToString
                    'sheetName10 = targetWorksheet.Range("D10").Value.ToString
                    'sheetName11 = targetWorksheet.Range("D11").Value.ToString

                ElseIf k = 5 Then
                    targetWorksheet.Range("E2").Value = excelApp.Union(worksheet.Range(targetWorksheet.Range("E2").Value), des_rng).Address
                    'Header = targetWorksheet.Range("E3").Value.ToString()
                    'Ascending = targetWorksheet.Range("E4").Value.ToString()
                    'Descending = targetWorksheet.Range("E5").Value.ToString()
                    'TextConvert = targetWorksheet.Range("E6").Value.ToString()
                    'OptionType = targetWorksheet.Range("E7").Value.ToString()
                    'Horizontal_CreateDP = targetWorksheet.Range("E8").Value.ToString()
                    'Flag_CreateDDDL = targetWorksheet.Range("E9").Value.ToString
                    'sheetName10 = targetWorksheet.Range("E10").Value.ToString
                    'sheetName11 = targetWorksheet.Range("E11").Value.ToString
                End If
                src_rng.Select()
                Me.Refresh()
                Me.Hide()
                MessageBox.Show("Your Dynamic Drop-down List is extended successfully.", "Softeko", MessageBoxButtons.OK, MessageBoxIcon.Information)

                Me.Close()
            Catch ex As Exception
                Me.Close()
            End Try
        End If
    End Sub

    Private Sub Btn_Cancel_Click(sender As Object, e As EventArgs) Handles Btn_Cancel.Click
        Close()
    End Sub

    Private Sub Form32_ExtendDropDownList_Disposed(sender As Object, e As EventArgs) Handles Me.Disposed
        form_flag = False
    End Sub

    Private Sub Form32_ExtendDropDownList_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        form_flag = False
    End Sub

    Private Sub Form32_ExtendDropDownList_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        Me.Focus()
        Me.BringToFront()
        Me.Activate()
        Me.BeginInvoke(New System.Action(Sub()
                                             TB_src_rng.Text = src_rng.Address
                                             SetWindowPos(Me.Handle, New IntPtr(HWND_TOPMOST), 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOSIZE)
                                         End Sub))
    End Sub

    Private Sub TB_des_rng_TextChanged(sender As Object, e As EventArgs) Handles TB_des_rng.TextChanged
        excelApp = Globals.ThisAddIn.Application
        Dim workbook As Excel.Workbook = excelApp.ActiveWorkbook
        Dim worksheet As Excel.Worksheet = workbook.ActiveSheet
        Try
            If TB_des_rng.Text IsNot Nothing And IsValidExcelCellReference(TB_des_rng.Text) = True Then
                focuschange = True
                Dim sheetname As String = ""

                Try

                    des_rng = worksheet.Range(TB_des_rng.Text)
                    des_rng.Select()

                Catch
                    ' Split the string into sheet name and cell address
                    Dim parts As String() = TB_des_rng.Text.Split("!"c)
                    sheetname = parts(0)
                    Dim cellAddress As String = parts(1)
                    worksheet = CType(workbook.Worksheets(sheetname), Worksheet)
                    worksheet.Activate()
                    des_rng = worksheet.Range(cellAddress)
                    des_rng.Select()
                End Try

                If workSheet2.Name <> worksheet.Name And TB_des_rng.Text.Contains("!") = False Then

                    TB_des_rng.Text = worksheet.Name & "!" & TB_des_rng.Text

                End If

                Me.Activate()
                TB_des_rng.Focus()
                TB_des_rng.SelectionStart = TB_des_rng.Text.Length

                focuschange = False
                ax = worksheet.Name
                workSheet3 = worksheet
            End If
        Catch ex As Exception
            focuschange = False
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

    Private Sub TB_src_rng_TextChanged(sender As Object, e As EventArgs) Handles TB_src_rng.TextChanged
        excelApp = Globals.ThisAddIn.Application
        workBook = excelApp.ActiveWorkbook
        workSheet = workBook.ActiveSheet

        If TB_src_rng.Text IsNot Nothing And IsValidExcelCellReference(TB_src_rng.Text.ToUpper) = True Then
            focuschange = True

            ' Define the range of cells to read (for example, cells A1 to A10)
            src_rng = excelApp.Range(TB_src_rng.Text)
            src_rng.Select()

            Me.Activate()
            'TB_src_range.Focus()
            TB_src_rng.SelectionStart = TB_src_rng.Text.Length
            focuschange = False
            workSheet2 = workSheet
        End If
    End Sub
End Class
