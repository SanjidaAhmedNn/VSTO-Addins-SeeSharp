Imports System.Threading
Imports System.Windows.Forms
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop.Excel
Imports System.Reflection.Emit
Imports System.Linq
Imports System.Media
Imports System.Security.Cryptography.X509Certificates
Imports System.Data

Public Class Form30_Create_Dynamic_Drop_down_List

    Dim WithEvents excelApp As Excel.Application
    Dim workBook As Excel.Workbook
    Dim workSheet As Excel.Worksheet
    Dim workSheet2 As Excel.Worksheet
    Dim src_rng As Excel.Range
    Public des_rng As Excel.Range
    Dim selectedRange As Excel.Range

    Dim opened As Integer


    Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As IntPtr, ByVal hWndInsertAfter As IntPtr, ByVal X As Integer, ByVal Y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal uFlags As UInteger) As Boolean
    Private Const SWP_NOMOVE As UInteger = &H2
    Private Const SWP_NOSIZE As UInteger = &H1
    Private Const SWP_NOACTIVATE As UInteger = &H10
    Private Const HWND_TOPMOST As Integer = -1

    Private Sub CB_ascending_CheckedChanged(sender As Object, e As EventArgs) Handles CB_ascending.CheckedChanged
        If CB_ascending.Checked = True Then
            CB_descending.Checked = False
        End If
    End Sub

    Private Sub CB_descending_CheckedChanged(sender As Object, e As EventArgs) Handles CB_descending.CheckedChanged
        If CB_descending.Checked = True Then
            CB_ascending.Checked = False
        End If
    End Sub

    Private Sub RB_columns_CheckedChanged(sender As Object, e As EventArgs) Handles RB_Dropdown_35_Labels.CheckedChanged
        If RB_Dropdown_35_Labels.Checked = True Then

            CB_header.Enabled = True
            CB_ascending.Enabled = True
            CB_descending.Enabled = True
            CB_text.Enabled = True
            GB_list_option.Enabled = False

        End If
    End Sub

    Private Sub RB_rows_CheckedChanged(sender As Object, e As EventArgs) Handles RB_Dropdown_2_Labels.CheckedChanged
        If RB_Dropdown_2_Labels.Checked = True Then
            GB_list_option.Enabled = True
            CB_header.Enabled = False
            CB_ascending.Enabled = False
            CB_descending.Enabled = False
            CB_text.Enabled = False

        End If
    End Sub



    Private Sub Selection_source_Click(sender As Object, e As EventArgs) Handles Selection_source.Click
        Try
            If selectedRange Is Nothing Then
            Else

                TB_src_range.Text = selectedRange.Address


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

                TB_src_range.Text = src_rng.Address

                Me.Show()
                TB_src_range.Focus()
            End If

        Catch ex As Exception

            Me.Show()
            TB_src_range.Focus()

        End Try
    End Sub

    ' Event handler to detect changes in E1 and adjust dropdown in E2


    Private Sub Btn_OK_Click(sender As Object, e As EventArgs) Handles Btn_OK.Click

        Dim workbook As Excel.Workbook = excelApp.ActiveWorkbook
        Dim worksheet As Excel.Worksheet = workbook.ActiveSheet
        If RB_Dropdown_35_Labels.Checked = True Then
            Dim rng As Excel.Range
            If CB_header.Checked = True Then
                'Dim adjustRange As Excel.Range
                rng = src_rng.Offset(1, 0).Resize(src_rng.Rows.Count - 1, src_rng.Columns.Count)

            Else

                rng = src_rng 'Assuming you have a range from A1 to A100
            End If
            'Dim rng2 As Excel.Range = excelApp.Range("B1:B16")
            'Dim rng3 As Excel.Range = excelApp.Range("C1:C16")

            Dim uniqueValues As New List(Of String)

            'Extract unique values from the range
            For Each cell As Excel.Range In rng.Columns(1).Cells
                Dim value As String = cell.Value
                If Not uniqueValues.Contains(value) Then
                    uniqueValues.Add(value)
                End If
            Next

            If CB_ascending.Checked = True Then
                'Sort the list in ascending order
                uniqueValues.Sort()
            ElseIf CB_descending.Checked = True Then
                'Sort the list in ascending order
                uniqueValues.Sort()
                uniqueValues.Reverse()
            End If

            'Create drop-down list at B1 with the unique values
            Dim dropDownRange As Excel.Range = des_rng.Columns(1)
            Dim validation As Excel.Validation = dropDownRange.Validation
            validation.Delete() 'Remove any existing validation
            validation.Add(Excel.XlDVType.xlValidateList, Formula1:=String.Join(",", uniqueValues))


            'MsgBox(i)

            AddHandler worksheet.Change, AddressOf worksheet_Change

        ElseIf RB_Dropdown_2_Labels.Checked = True Then
            ' Extract headers from A1:C1
            Dim headersRange As Excel.Range = src_rng.Rows(1)
            Dim headers As List(Of String) = New List(Of String)
            ' Dim workbook As excelapp.workbook

            For Each cell As Excel.Range In headersRange.Cells
                headers.Add(cell.Value.ToString())
            Next
            'Dim workbook As Excel.Workbook = excelApp.ActiveWorkbook
            'Dim worksheet As Excel.Worksheet = workbook.ActiveSheet
            ' Create the dropdown list with headers in cell E1
            'CreateValidationList(excelApp.ActiveSheet.Range("$E$1"), String.Join(",", headers))
            'Create drop-down list at B1 with the unique values
            Dim dropDownRange As Excel.Range = des_rng(1, 1)
            Dim validation As Excel.Validation = dropDownRange.Validation
            validation.Delete() 'Remove any existing validation
            validation.Add(Excel.XlDVType.xlValidateList, Formula1:=String.Join(",", headers))

            ' Add event handler to listen for changes in E1

            AddHandler worksheet.Change, AddressOf worksheet_Change
        End If
        Me.Close()


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

            TB_dest_range.Text = des_rng.Address

            Me.Show()
            TB_dest_range.Focus()

        End If
    End Sub


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Try

            excelApp = Globals.ThisAddIn.Application

            AddHandler excelApp.SheetSelectionChange, AddressOf excelApp_SheetSelectionChange

            opened = opened + 1

            If excelApp.Selection IsNot Nothing Then
                selectedRange = excelApp.Selection
                src_rng = selectedRange
                TB_src_range.Text = selectedRange.Address
            End If

        Catch ex As Exception

        End Try

    End Sub


    Private Sub excelApp_SheetSelectionChange(ByVal Sh As Object, ByVal selectionRange1 As Excel.Range) Handles excelApp.SheetSelectionChange
        Try

            excelApp = Globals.ThisAddIn.Application

            If Me.ActiveControl Is TB_dest_range Then
                'des_rng = selectionRange1
                ' This will run on the Excel thread, so you need to use Invoke to update the UI
                'Me.BeginInvoke(New System.Action(Sub() TB_dest_range.Text = selectionRange1.Address))
                'Me.Activate()
                'Me.BeginInvoke(New System.Action(Sub()

            ElseIf Me.ActiveControl Is TB_src_range Then
                'src_rng = selectionRange1
                'workSheet = workBook.ActiveSheet
                'TB_src_range.Text = src_rng.Address
                'TB_src_range.Focus()
                'Me.Activate()
                'ActiveForm.Select()


                'Me.BeginInvoke(New System.Action(Sub()
                'TB_src_range.Text = src_rng.Address
                ' SetWindowPos(Me.Handle, New IntPtr(HWND_TOPMOST), 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOSIZE)
                'End Sub))
            End If



        Catch ex As Exception

        End Try

    End Sub

    Private Sub Btn_Cancel_Click(sender As Object, e As EventArgs) Handles Btn_Cancel.Click

        Me.Close()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs)


    End Sub


    Private Sub worksheet_Change(ByVal Target As Excel.Range)
        Dim rng As Excel.Range
        If RB_Dropdown_35_Labels.Checked = True Then
            If CB_header.Checked = True Then
                'Dim adjustRange As Excel.Range
                rng = src_rng.Offset(1, 0).Resize(src_rng.Rows.Count - 1, src_rng.Columns.Count)

            Else

                rng = src_rng 'Assuming you have a range from A1 to A100
            End If
            ' MsgBox(des_rng.Rows.Count)
            For k = 1 To des_rng.Rows.Count
                Dim matchedValues As New List(Of String)
                'Dim i As Integer = 0
                'MsgBox(i)
                If des_rng(k, 1).Value IsNot Nothing Then
                    For i = 1 To rng.Rows.Count
                        If rng(i, 1).Value = des_rng(k, 1).Value Then
                            If Not matchedValues.Contains(rng(i, 2).Value) Then
                                matchedValues.Add(rng(i, 2).Value)
                            End If
                            'matchedValues.Add(rng(i, 2).Value)
                        End If
                    Next


                    If CB_ascending.Checked = True Then
                        'Sort the list in ascending order
                        matchedValues.Sort()
                    ElseIf CB_descending.Checked = True Then
                        'Sort the list in ascending order
                        matchedValues.Sort()
                        matchedValues.Reverse()
                    End If

                    'MsgBox(i)
                    'Create drop-down list at B1 with the unique values
                    'Dim dropDownRange As Excel.Range = des_rng(k, 2)
                    Dim dropDownRange As Excel.Range = des_rng(k, 2)
                    Dim Validation As Excel.Validation = dropDownRange.Validation
                    Validation.Delete() 'Remove any existing validation
                    Validation.Add(Excel.XlDVType.xlValidateList, Formula1:=String.Join(",", matchedValues))
                    matchedValues.Clear()
                    'MsgBox(k)
                End If

                Dim sec_matchedValues As New List(Of String)
                'Dim i As Integer = 0
                'MsgBox(i)

                If des_rng(k, 2).Value IsNot Nothing Then
                    For i = 1 To rng.Rows.Count
                        If rng(i, 1).Value = des_rng(k, 1).Value And rng(i, 2).Value = des_rng(k, 2).Value Then
                            sec_matchedValues.Add(rng(i, 3).Value)
                        End If
                    Next


                    If CB_ascending.Checked = True Then
                        'Sort the list in ascending order
                        sec_matchedValues.Sort()
                    ElseIf CB_descending.Checked = True Then
                        'Sort the list in ascending order
                        sec_matchedValues.Sort()
                        sec_matchedValues.Reverse()
                    End If

                    'MsgBox(i)
                    'Create drop-down list at B1 with the unique values
                    'dropDownRange = des_rng(k, 3)
                    Dim dropDownRange As Excel.Range = des_rng(k, 3)
                    Dim Validation As Excel.Validation = dropDownRange.Validation
                    Validation.Delete() 'Remove any existing validation
                    Validation.Add(Excel.XlDVType.xlValidateList, Formula1:=String.Join(",", sec_matchedValues))
                    sec_matchedValues.Clear()
                End If
            Next

        ElseIf RB_Dropdown_2_Labels.Checked = True Then
            If RB_Horizon.Checked = True Then
                If Target.Address = des_rng(1, 1).Address Then
                    Dim worksheet As Excel.Worksheet = CType(Target.Worksheet, Excel.Worksheet)
                    Dim col As Integer = src_rng.Rows().Find(Target.Value).Column
                    Dim sourceRng As Excel.Range = src_rng.Cells(2, col).Resize(worksheet.Cells(worksheet.Rows.Count, col).End(Excel.XlDirection.xlUp).Row - 1, 1)
                    Dim dropDownRange As Excel.Range = des_rng(1, 2)
                    Dim Validation As Excel.Validation = dropDownRange.Validation
                    Validation.Delete() 'Remove any existing validation
                    Validation.Add(Excel.XlDVType.xlValidateList, Formula1:="=" & sourceRng.Address)
                    'CreateValidationList(worksheet.Cells(2, 5), "=" & sourceRng.Address)
                End If

            ElseIf RB_Verti.Checked = True Then
                If Target.Address = des_rng(1, 1).Address Then
                    Dim worksheet As Excel.Worksheet = CType(Target.Worksheet, Excel.Worksheet)
                    Dim col As Integer = src_rng.Rows().Find(Target.Value).Column
                    Dim sourceRng As Excel.Range = src_rng.Cells(2, col).Resize(worksheet.Cells(worksheet.Rows.Count, col).End(Excel.XlDirection.xlUp).Row - 1, 1)
                    Dim dropDownRange As Excel.Range = des_rng(2, 1)
                    Dim Validation As Excel.Validation = dropDownRange.Validation
                    Validation.Delete() 'Remove any existing validation
                    Validation.Add(Excel.XlDVType.xlValidateList, Formula1:="=" & sourceRng.Address)
                    'CreateValidationList(worksheet.Cells(2, 5), "=" & sourceRng.Address)
                End If
            End If

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
End Class