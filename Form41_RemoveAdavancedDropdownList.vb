Imports System.Windows.Forms
Imports System.ComponentModel
Imports System.Diagnostics
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop

Imports Microsoft.Office.Interop.Excel
Imports System.Data

Public Class Form41_RemoveAdavancedDropdownList

    Dim WithEvents excelApp As Excel.Application
    Dim workBook As Excel.Workbook
    Public Shared workSheet As Excel.Worksheet
    Private SheetHandlers As New List(Of WorksheetHandler)
    Private EventDel_CellsChange As Excel.DocEvents_ChangeEventHandler

    Private WithEvents CurrentSheet As Excel.Worksheet
    Private WithEvents WorkbookEvents As Excel.Workbook

    Dim srcRng1 As String
    Dim srcRng2 As String
    Dim srcRng3 As String

    Dim form As Form36 = Nothing

    Dim src_rng As Excel.Range
    Dim frm1 As Form35Multi_SelectionbasedDropdown = Nothing

    Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As IntPtr, ByVal hWndInsertAfter As IntPtr, ByVal X As Integer, ByVal Y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal uFlags As UInteger) As Boolean
    Private Const SWP_NOMOVE As UInteger = &H2
    Private Const SWP_NOSIZE As UInteger = &H1
    Private Const SWP_NOACTIVATE As UInteger = &H10
    Private Const HWND_TOPMOST As Integer = -1

    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CB_search.CheckedChanged

    End Sub

    Private Sub CB_multiselect_CheckedChanged(sender As Object, e As EventArgs) Handles CB_multiselect.CheckedChanged

    End Sub

    Private Sub Btn_OK_Click(sender As Object, e As EventArgs) Handles Btn_OK.Click
        excelApp = Globals.ThisAddIn.Application
        Dim workbook As Excel.Workbook = excelApp.ActiveWorkbook
        Dim worksheet As Excel.Worksheet = workbook.ActiveSheet

        If CB_Source.Text = "Select Range" And TB_src_rng.Text = "" Then
            MessageBox.Show("Select a Source Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TB_src_rng.Focus()
            'Me.Close()
            Exit Sub
        ElseIf CB_multiselect.Checked = False And CB_checkbox.Checked = False And CB_search.Checked = False Then
            MessageBox.Show("Please, select the Data Validation List type.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TB_src_rng.Focus()
            'Me.Close()
            Exit Sub

        Else

            If CB_Source.Text <> "Select Range" Then
                src_rng = worksheet.Range("A1", worksheet.Cells(excelApp.Rows.Count, excelApp.Columns.Count))
            End If

            If CB_Source.Text.Contains("Active Workbook") And CB_multiselect.Checked = True Then
                For Each sheet As Excel.Worksheet In workbook.Sheets
                    If sheet.Name = "Newwwwwwwwww" Then
                        sheet.Delete()
                        Exit For
                    End If
                Next
                GB_CB_Source1 = ""
                SR1 = ""
                Horizontal1 = False
                Separator1 = ""
                Search1 = False
                Flag1 = False
                TargetVar1 = ""
                RangeType1 = ""

            ElseIf CB_Source.Text.Contains("Active Workbook") And CB_checkbox.Checked = True Then
                For Each sheet As Excel.Worksheet In workbook.Sheets
                    If sheet.Name = "SofTekoSofTeko" Then
                        sheet.Delete()
                        Exit For
                    End If
                Next
                GB_CB_Source2 = ""
                SR2 = ""
                Horizontal2 = False
                Separator2 = ""
                Search2 = False
                Flag2 = False
                TargetVar2 = ""
                RangeType2 = ""

            ElseIf CB_Source.Text.Contains("Active Workbook") And CB_search.Checked = True Then
                For Each sheet As Excel.Worksheet In workbook.Sheets
                    If sheet.Name = "SofTekoSofTekoSofteko" Then
                        sheet.Delete()
                        Exit For
                    End If
                Next
                GB_CB_Source3 = ""
                SR3 = ""
                Horizontal3 = False
                Separator3 = ""
                Search3 = False
                Flag3 = False
                TargetVar3 = ""
                RangeType3 = ""

            ElseIf CB_Source.Text = "Select Range" Or CB_Source.Text.Contains("Active Sheet") Then

                If CB_Source.Text.Contains("Active Sheet") Then
                    src_rng = worksheet.Range("A1", worksheet.Cells(excelApp.Rows.Count, excelApp.Columns.Count))
                End If

                If CB_multiselect.Checked = True Then
                    'RemoveHandler worksheet.SelectionChange, AddressOf sheet_SelectionChange
                    srcRng1 = GB_CB_Source1


                    If IsCellInsideRange(src_rng, worksheet.Range(srcRng1)) = True Then
                        'Dim result As Excel.Range = SubtractRanges(rng1, rng2)
                        'Dim results As List(Of Excel.Range) = SubtractRanges(worksheet.Range(srcRng1), src_rng)
                        'Dim addressList As New List(Of String)

                        'Dim combinedAddress As String = ""
                        'For Each r In results
                        '    addressList.Add(r.Address)

                        'Next

                        'combinedAddress = String.Join(",", addressList)
                        'targetWorksheet.Name & "!" & targetRange.Address(External:=False)
                        'GB_CB_Dlt = combinedAddress
                        GB_CB_Dlt1 = src_rng.Address
                        Nam1 = workbook.ActiveSheet.name
                        'MsgBox(GB_CB_Source1)
                    End If



                End If

                If CB_checkbox.Checked = True Then
                    srcRng2 = GB_CB_Source2


                    If IsCellInsideRange(src_rng, worksheet.Range(srcRng2)) = True Then
                        'Dim result As Excel.Range = SubtractRanges(rng1, rng2)
                        'Dim results As List(Of Excel.Range) = SubtractRanges(worksheet.Range(srcRng2), src_rng)
                        'Dim addressList As New List(Of String)

                        'Dim combinedAddress As String = ""
                        'For Each r In results
                        '    addressList.Add(r.Address)

                        'Next

                        'combinedAddress = String.Join(",", addressList)
                        'targetWorksheet.Name & "!" & targetRange.Address(External:=False)
                        'GB_CB_Dlt = combinedAddress
                        GB_CB_Dlt2 = src_rng.Address
                        Nam2 = workbook.ActiveSheet.name
                        'MsgBox(GB_CB_Source1)
                    End If

                End If

                If CB_search.Checked = True Then
                    srcRng3 = GB_CB_Source3


                    If IsCellInsideRange(src_rng, worksheet.Range(srcRng3)) = True Then
                        'Dim result As Excel.Range = SubtractRanges(rng1, rng2)
                        'Dim results As List(Of Excel.Range) = SubtractRanges(workSheet.Range(srcRng3), src_rng)
                        'Dim addressList As New List(Of String)

                        'Dim combinedAddress As String = ""
                        'For Each r In results
                        '    addressList.Add(r.Address)

                        'Next

                        'combinedAddress = String.Join(",", addressList)
                        'GB_CB_Source3 = combinedAddress

                        GB_CB_Dlt3 = src_rng.Address
                        Nam3 = workbook.ActiveSheet.name

                    End If

                End If

            Else
                'MsgBox(1)
                'Dim targetWorksheet As Excel.Worksheet = Nothing
                'targetWorksheet = CType(workbook.Sheets(CB_Source.Text), Excel.Worksheet)
                ''MsgBox(2)
                ''src_rng = worksheet.Range(CB_Source.Text)
                ''src_rng = workbook.Sheet(CB_Source.Text).src_rng
                'src_rng = targetWorksheet.Range(src_rng.Address) ' Change the range as needed


                If CB_multiselect.Checked = True Then
                    'srcRng1 = GB_CB_Source1
                    'Dim srcRng1_prime As Excel.Range = workbook.ActiveSheet.Range(srcRng1)

                    'If IsCellInsideRange(src_rng, srcRng1_prime) = True Then
                    '    'Dim result As Excel.Range = SubtractRanges(rng1, rng2)
                    '    Dim results As List(Of Excel.Range) = SubtractRanges(srcRng1_prime, src_rng)
                    '    Dim addressList As New List(Of String)

                    '    Dim combinedAddress As String = ""
                    '    For Each r In results
                    '        addressList.Add(r.Address)

                    '    Next

                    '    combinedAddress = String.Join(",", addressList)
                    '    GB_CB_Source1 = combinedAddress
                    '    ' MsgBox()

                    'End If


                    src_rng = worksheet.Range("A1", worksheet.Cells(excelApp.Rows.Count, excelApp.Columns.Count))
                    GB_CB_Dlt1 = src_rng.Address
                    Nam1 = CB_Source.Text


                End If

                If CB_checkbox.Checked = True Then
                    'srcRng2 = GB_CB_Source2
                    'Dim srcRng2_prime As Excel.Range = workbook.ActiveSheet.Range(srcRng2)


                    'If IsCellInsideRange(src_rng, srcRng2_prime) = True Then
                    '    'Dim result As Excel.Range = SubtractRanges(rng1, rng2)
                    '    Dim results As List(Of Excel.Range) = SubtractRanges(srcRng2_prime, src_rng)
                    '    Dim addressList As New List(Of String)

                    '    Dim combinedAddress As String = ""
                    '    For Each r In results
                    '        addressList.Add(r.Address)

                    '    Next

                    '    combinedAddress = String.Join(",", addressList)
                    '    GB_CB_Source2 = combinedAddress

                    'End If

                    src_rng = worksheet.Range("A1", worksheet.Cells(excelApp.Rows.Count, excelApp.Columns.Count))
                    GB_CB_Dlt2 = src_rng.Address
                    Nam2 = CB_Source.Text

                End If

                If CB_search.Checked = True Then
                    'srcRng3 = GB_CB_Source3
                    'Dim srcRng3_prime As Excel.Range = workbook.ActiveSheet.Range(srcRng3)


                    'If IsCellInsideRange(src_rng, srcRng3_prime) = True Then
                    '    'Dim result As Excel.Range = SubtractRanges(rng1, rng2)
                    '    Dim results As List(Of Excel.Range) = SubtractRanges(srcRng3_prime, src_rng)
                    '    Dim addressList As New List(Of String)

                    '    Dim combinedAddress As String = ""
                    '    For Each r In results
                    '        addressList.Add(r.Address)

                    '    Next

                    '    combinedAddress = String.Join(",", addressList)
                    '    GB_CB_Source3 = combinedAddress

                    'End If

                    src_rng = worksheet.Range("A1", worksheet.Cells(excelApp.Rows.Count, excelApp.Columns.Count))
                    GB_CB_Dlt3 = src_rng.Address
                    Nam3 = CB_Source.Text

                End If
                'MsgBox(src_rng.Address)
            End If

            If CB_multiselect.Checked = True Then
                TType1 = CB_Source.Text
            End If

            If CB_checkbox.Checked = True Then
                TType2 = CB_Source.Text
            End If

            If CB_search.Checked = True Then
                TType3 = CB_Source.Text
            End If


            Close()

        End If
    End Sub

    'Private Sub sheet_SelectionChange(ByVal Target As Excel.Range)
    '    excelApp = Globals.ThisAddIn.Application
    '    workBook = excelApp.ActiveWorkbook
    '    workSheet = workBook.ActiveSheet
    '    If GB_CB_Source1 IsNot Nothing Then

    '        ' src_rng = workSheet.Range(GB_CB_Source1)
    '        src_rng = workSheet.Range(GB_CB_Source1)

    '        'MsgBox(workSheet.Name)
    '        'MsgBox(src_rng.Worksheet.Name)

    '        If CB_Source.Text.Contains("Active Workbook") Then
    '            src_rng = workSheet.Range("A1", workSheet.Cells(excelApp.Rows.Count, excelApp.Columns.Count))
    '        Else

    '        End If

    '        src_rng = workSheet.Range(GB_CB_Source1)

    '        src_rng = workBook.ActiveSheet.range(src_rng.Address)


    '        ' MsgBox(src_rng.Worksheet.Name)
    '        'Recheck: Newly added
    '        If CB_Source.Text.Contains("Active Sheet") And Nam <> workSheet.Name Then
    '            Exit Sub
    '        End If

    '        If IsCellInsideRange(Target, src_rng) And Target.Cells.Count = 1 And HasDataValidationList(Target) Then
    '            'If Target.Cells.Count = 1 Then ' Ensure only one cell is selected
    '            TargetVar1 = Target.Address
    '            If form Is Nothing OrElse form.IsDisposed Then
    '                form = New Form36()
    '                form.Show()
    '                form.Refresh()
    '            Else
    '                ' If form is already open, bring it to the front
    '                'Form = Form36()
    '                'Form.Refresh()
    '                'Form.BringToFront()
    '                'Form.Refresh()
    '                form.Dispose()
    '                form = New Form36()
    '                form.Show()
    '            End If
    '        End If

    '        'Dim form As New Form36()
    '        'form.Show()
    '        'form.Focus()
    '        ''form.TopMost = True
    '        ''form.Activate()
    '        'form.BringToFront()
    '        'End If
    '    End If

    'End Sub


    Private Function HasDataValidationList(ByVal cell As Excel.Range) As Boolean
        Dim hasValidation As Boolean = False

        Try
            If Not cell.Validation Is Nothing AndAlso cell.Validation.Type = Excel.XlDVType.xlValidateList Then
                hasValidation = True
            End If
        Catch ex As Exception
            ' Exception will be thrown if cell doesn't have validation. No action needed.
        End Try

        Return hasValidation
    End Function


    Private Sub Btn_Cancel_Click(sender As Object, e As EventArgs) Handles Btn_Cancel.Click
        Me.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) 
        excelApp = Globals.ThisAddIn.Application
        Dim workbook As Excel.Workbook = excelApp.ActiveWorkbook
        Dim worksheet As Excel.Worksheet = workbook.ActiveSheet

        'Dim results As List(Of Excel.Range)


        Dim rng1 As Excel.Range = worksheet.Range("A1:D10")
        Dim rng2 As Excel.Range = worksheet.Range("A1:B5")
        If IsCellInsideRange(rng2, rng1) = True Then
            'Dim result As Excel.Range = SubtractRanges(rng1, rng2)
            Dim results As List(Of Excel.Range) = SubtractRanges(rng1, rng2)
            Dim addressList As New List(Of String)

            Dim combinedAddress As String = ""
            For Each r In results
                addressList.Add(r.Address)


            Next

            combinedAddress = String.Join(",", addressList)
            worksheet.Range(combinedAddress).Select()

        End If

        'worksheet.Select("A1:A3", "D1:D2")
        'If Not result Is Nothing Then
        '    ' Do something with the result range
        '    MessageBox.Show(result.Address)
        'Else
        '    MessageBox.Show("Ranges are either equivalent or do not have a direct subtraction result.")
        'End If

        Dim rng3 As Excel.Range = worksheet.Range("A1:A10")
        Dim rng4 As Excel.Range = worksheet.Range("C1:C10")

        Dim combinedRange As Excel.Range = excelApp.Union(rng3, rng4) ' Assuming ExcelApp is your Excel.Application object

        ' combinedRange.Select()
        'MsgBox(combinedRange.Address)
    End Sub

    Function SubtractRanges(rng1 As Excel.Range, rng2 As Excel.Range) As List(Of Excel.Range)
        Dim result As New List(Of Excel.Range)()

        ' Top-left and bottom-right cells of rng1
        Dim tl1 As Excel.Range = rng1.Cells(1, 1)
        Dim br1 As Excel.Range = rng1.Cells(rng1.Rows.Count, rng1.Columns.Count)

        ' Top-left and bottom-right cells of rng2
        Dim tl2 As Excel.Range = rng2.Cells(1, 1)
        Dim br2 As Excel.Range = rng2.Cells(rng2.Rows.Count, rng2.Columns.Count)

        ' Check rows above rng2
        If tl1.Row < tl2.Row Then
            result.Add(rng1.Worksheet.Range(tl1, rng1.Cells(tl2.Row - 1, br1.Column)))
        End If

        ' Check rows below rng2
        If br1.Row > br2.Row Then
            result.Add(rng1.Worksheet.Range(rng1.Cells(br2.Row + 1, tl1.Column), br1))
        End If

        ' Check columns to the left of rng2
        If tl1.Column < tl2.Column Then
            result.Add(rng1.Worksheet.Range(tl1, rng1.Cells(br1.Row, tl2.Column - 1)))
        End If

        ' Check columns to the right of rng2
        If br1.Column > br2.Column Then
            result.Add(rng1.Worksheet.Range(rng1.Cells(tl1.Row, br2.Column + 1), br1))
        End If

        Return result
    End Function


    Private Function IsCellInsideRange(ByVal cell As Excel.Range, ByVal targetRange As Excel.Range) As Boolean
        'MsgBox(cell.Address)
        'MsgBox(targetRange.Address)
        Try
            Dim intersectRange As Excel.Range = Globals.ThisAddIn.Application.Intersect(cell, targetRange)
            'MsgBox(intersectRange.Address)
            Return Not intersectRange Is Nothing
        Catch ex As Exception
            'MsgBox(cell.Address)
            'MsgBox(targetRange.Address)
            Return False
        End Try
    End Function

    Private Sub Selection_source_Click(sender As Object, e As EventArgs) Handles Selection_source.Click
        Try

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

        Catch ex As Exception

            Me.Show()
            TB_src_rng.Focus()

        End Try
    End Sub

    Private Sub Form41_RemoveAdavancedDropdownList_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.KeyPreview = True

        excelApp = Globals.ThisAddIn.Application
        Dim workbook As Excel.Workbook = excelApp.ActiveWorkbook
        Dim worksheet As Excel.Worksheet = workbook.ActiveSheet

        'Timer1.Start()
        CB_Source.Items.Add("Select Range")
        CB_Source.Items.Add("Active Workbook :" & workbook.Name)
        CB_Source.Items.Add("Active Sheet :" & worksheet.Name)

        Dim i As Integer = 0
        ' Loop through each worksheet in the workbook.
        For Each WS In workbook.Sheets
            ' Check if the worksheet is not hidden.
            If WS.Visible = Excel.XlSheetVisibility.xlSheetVisible And WS.name <> worksheet.Name Then
                CB_Source.Items.Add(WS.Name)
                i = i + 1
            End If
        Next

        'Only Enable when select Range is selected in combobox
        If CB_Source.Text = "Select Range" Then
            TB_src_rng.Enabled = True
            Selection_source.Enabled = True
        Else
            TB_src_rng.Enabled = False
            Selection_source.Enabled = False
        End If
    End Sub

    Private Sub TB_src_rng_TextChanged(sender As Object, e As EventArgs) Handles TB_src_rng.TextChanged

    End Sub

    Private Sub Form41_RemoveAdavancedDropdownList_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Enter Then
            Btn_OK.PerformClick()
        End If
    End Sub

    Private Sub Form41_RemoveAdavancedDropdownList_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        form_flag = False
    End Sub

    Private Sub Form41_RemoveAdavancedDropdownList_Disposed(sender As Object, e As EventArgs) Handles Me.Disposed
        form_flag = False
    End Sub


    Private Sub Form41_RemoveAdavancedDropdownList_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        TB_src_rng.Focus()
        Me.Focus()
        Me.BringToFront()
        Me.Activate()
        Me.BeginInvoke(New System.Action(Sub()
                                             TB_src_rng.Text = src_rng.Address
                                             SetWindowPos(Me.Handle, New IntPtr(HWND_TOPMOST), 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOSIZE)
                                         End Sub))
    End Sub
End Class