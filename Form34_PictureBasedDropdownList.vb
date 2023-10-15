Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Runtime.InteropServices
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports Microsoft.Office.Interop

Imports Microsoft.Office.Interop.Excel
Imports System.ComponentModel
Imports System.Linq.Expressions

Public Class Form34_PictureBasedDropdownList
    Dim WithEvents excelApp As Excel.Application
    Dim workBook As Excel.Workbook
    Public Shared workSheet As Excel.Worksheet

    Dim src_rng As Excel.Range
    Public des_rng As Excel.Range
    Dim selectedRange As Excel.Range

    Public validationRange As Excel.Range

    Private processingEvent As Boolean = False
    Public focuschange As Boolean

    Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As IntPtr, ByVal hWndInsertAfter As IntPtr, ByVal X As Integer, ByVal Y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal uFlags As UInteger) As Boolean
    Private Const SWP_NOMOVE As UInteger = &H2
    Private Const SWP_NOSIZE As UInteger = &H1
    Private Const SWP_NOACTIVATE As UInteger = &H10
    Private Const HWND_TOPMOST As Integer = -1



    Private Sub Btn_OK_Click(sender As Object, e As EventArgs) Handles Btn_OK.Click
        excelApp = Globals.ThisAddIn.Application
        Dim workbook As Excel.Workbook = excelApp.ActiveWorkbook
        Dim worksheet As Excel.Worksheet = workbook.ActiveSheet

        sheetName2 = worksheet.Name


        Dim x As Boolean = False

        If IsValidExcelCellReference(TB_src_rng.Text) = True Then

            For Each pic As Excel.Shape In worksheet.Shapes
                For i = 1 To src_rng.Rows.Count
                    If pic.TopLeftCell.Address = src_rng(i, 2).Address Then

                        x = True
                        GoTo BreakAllLoops
                    Else
                        x = False

                    End If
                Next i

            Next
BreakAllLoops:
        End If

        If TB_src_rng.Text = "" And TB_des_rng.Text = "" Then
            MessageBox.Show("Select all necessary options.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TB_src_rng.Focus()
            'Me.Close()
            Exit Sub

        ElseIf TB_src_rng.Text = "" Then
            MessageBox.Show("Select a Source Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TB_src_rng.Focus()
            'Me.Close()
            Exit Sub

        ElseIf TB_des_rng.Text = "" Then
            MessageBox.Show("Select the Destination Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TB_des_rng.Focus()
            'Me.Close()
            Exit Sub


        ElseIf IsValidExcelCellReference(TB_src_rng.Text) = False Then
            MessageBox.Show("Select a valid source cell range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TB_src_rng.Focus()
            Exit Sub

        ElseIf IsValidExcelCellReference(TB_des_rng.Text) = False Then
            MessageBox.Show("Select a valid destination cell range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TB_des_rng.Focus()
            Exit Sub
        ElseIf src_rng.Areas.Count > 1 Then
            MessageBox.Show("Multiple selection is not possible in the Source Range field. Please select two columns.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TB_des_rng.Focus()

            Exit Sub

        ElseIf src_rng.Columns.Count = 1 Then
            MessageBox.Show("Please select both of the columns that contain the data and the relevant images.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TB_src_rng.Focus()
            'Me.Close()
            Exit Sub

        ElseIf src_rng.Columns.Count > 2 Then
            MessageBox.Show("Please, Select two columns.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TB_src_rng.Focus()
            'Me.Close()
            Exit Sub

        ElseIf des_rng.Columns.Count <> 2 Then
            MessageBox.Show("Please, Select two columns.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TB_des_rng.Focus()
            'Me.Close()
            Exit Sub


        ElseIf x = False Then
            MessageBox.Show("Please select both of the columns that contain the data And the relevant images.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TB_des_rng.Focus()

            Exit Sub

        Else
            Try
                ' Set up validation list for 1st Column
                Dim rangeValues As Excel.Range = src_rng.Columns(1).cells
                Dim listString As String = ""
                'MsgBox(rangeValues.Address)
                For Each cell As Excel.Range In rangeValues
                    If listString <> "" Then
                        listString &= ","
                    End If
                    listString &= cell.Value
                Next

                ' Set data validation in C1
                validationRange = des_rng.Columns(1).cells
                With validationRange.Validation
                    .Delete() ' Delete any previous validation
                    .Add(Type:=Excel.XlDVType.xlValidateList, AlertStyle:=Excel.XlDVAlertStyle.xlValidAlertStop, Operator:=Excel.XlFormatConditionOperator.xlBetween, Formula1:=listString)
                    .IgnoreBlank = True
                    .ShowInput = True
                    .ShowError = True
                End With
                ' MsgBox(2)
                des_rng.Columns(2).ColumnWidth = src_rng.Columns(2).ColumnWidth
                des_rng.Rows.RowHeight = src_rng.Rows.RowHeight



                AddHandler worksheet.Change, AddressOf worksheet1_Change

                '2 ta event handler dile valo vabe kaj korena. Seijonno ektar event handler er moddhe arekta call kora hoise.

                'AddHandler worksheet.Change, AddressOf worksheet2_Change


                Dim targetWorksheet As Excel.Worksheet = Nothing
                For Each ws As Excel.Worksheet In excelApp.Worksheets
                    If ws.Name = "SoftekoPictureBasedDropDown" Then
                        targetWorksheet = ws
                        Exit For
                    End If
                Next

                ' If "MySpecialSheet" does not exist, add it
                If targetWorksheet Is Nothing Then
                    targetWorksheet = CType(excelApp.Worksheets.Add(After:=excelApp.Worksheets(excelApp.Worksheets.Count)), Excel.Worksheet)
                    targetWorksheet.Name = "SoftekoPictureBasedDropDown"
                End If


                Flag_Picture = True
                sheetName2 = worksheet.Name
                Src_Rng_of_PictureDDL = TB_src_rng.Text
                Des_Rng_of_PictureDDL = TB_des_rng.Text

                ' Write something in cell A1 of the target worksheet
                targetWorksheet.Range("A1").Value = "Do not delete the sheet!"
                targetWorksheet.Range("A2").Value = Flag_Picture
                targetWorksheet.Range("A3").Value = sheetName2
                targetWorksheet.Range("A4").Value = Src_Rng_of_PictureDDL
                targetWorksheet.Range("A5").Value = Des_Rng_of_PictureDDL
                targetWorksheet.Visible = Excel.XlSheetVisibility.xlSheetHidden

            Catch ex As Exception
            End Try
            Me.Dispose()
        End If

    End Sub


    Private Sub worksheet2_Change(ByVal Target As Excel.Range)
        excelApp = Globals.ThisAddIn.Application
        Dim workbook As Excel.Workbook = excelApp.ActiveWorkbook
        Dim worksheet As Excel.Worksheet = workbook.ActiveSheet

        Try

            For Each pic As Excel.Shape In worksheet.Shapes
                'MsgBox(pic.TopLeftCell.Address)
                If pic.TopLeftCell.Address = Target.Offset(0, 1).Address Then

                    pic.Delete()
                    'Exit For
                End If
            Next
        Catch ex As Exception
        End Try
    End Sub


    Private Sub worksheet1_Change(ByVal Target As Excel.Range)

        excelApp = Globals.ThisAddIn.Application
        Dim workbook As Excel.Workbook = excelApp.ActiveWorkbook
        Dim worksheet As Excel.Worksheet = workbook.ActiveSheet
        Try

            For i = 1 To src_rng.Rows.Count
                If src_rng(i, 1).Value = Target.Value Then
                    ' MsgBox(3)
                    Try
                        worksheet2_Change(Target)
                        '    MsgBox(5)
                    Catch ex As Exception
                        ' MsgBox(15)
                    End Try

                    'MsgBox(6)

                    '            Dim imageCell As Excel.Range = src_rng(i, 2)
                    '            imageCell.CopyPicture(
                    'Appearance:=Excel.XlPictureAppearance.xlScreen,
                    'Format:=Excel.XlCopyPictureFormat.xlPicture)
                    '            workSheet.Paste(Target.Offset(0, 1))

                    Dim x As Boolean = False
                    'Try
                    For Each pic As Excel.Shape In worksheet.Shapes
                        'MsgBox(pic.TopLeftCell.Address)
                        If pic.TopLeftCell.Address = src_rng(i, 2).Address Then
                            pic.CopyPicture()
                            worksheet.Paste(Target.Offset(0, 1))
                            Target.Offset(0, 1).RowHeight = src_rng(i, 2).RowHeight
                            ' Target.Offset(0, 1).RowHeight = src_rng(i, 2).C
                            x = True
                            Exit For
                        Else
                            x = False

                        End If
                        'x = x + 1
                    Next

                    excelApp.CutCopyMode = False


                End If
            Next
        Catch ex As Exception
        End Try


    End Sub


    Private Sub Form34_PictureBasedDropdownList_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try

            excelApp = Globals.ThisAddIn.Application
            Dim workbook As Excel.Workbook = excelApp.ActiveWorkbook
            Dim worksheet As Excel.Worksheet = workbook.ActiveSheet

            AddHandler excelApp.SheetSelectionChange, AddressOf excelApp_SheetSelectionChange

            'opened = opened + 1

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

    Private Sub excelApp_SheetSelectionChange(ByVal Sh As Object, ByVal selectionRange1 As Excel.Range) Handles excelApp.SheetSelectionChange
        Try

            excelApp = Globals.ThisAddIn.Application
            If focuschange = False Then
                If focuschange = False Then
                    If TB_des_rng.Focused = True Or Me.ActiveControl Is TB_des_rng Then
                        If TB_des_rng.Focused = True Then
                            des_rng = selectionRange1
                        End If
                        Me.Activate()
                        Me.BeginInvoke(New System.Action(Sub()
                                                             TB_des_rng.Text = des_rng.Address
                                                             SetWindowPos(Me.Handle, New IntPtr(HWND_TOPMOST), 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOSIZE)
                                                         End Sub))

                        ' ElseIf Me.ActiveControl Is TB_src_range Then
                    ElseIf TB_src_rng.Focused = True Or Me.ActiveControl Is TB_src_rng Then
                        If TB_src_rng.Focused = True Then
                            src_rng = selectionRange1
                        End If
                        Me.Activate()
                        Me.BeginInvoke(New System.Action(Sub()
                                                             TB_src_rng.Text = src_rng.Address
                                                             SetWindowPos(Me.Handle, New IntPtr(HWND_TOPMOST), 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOSIZE)
                                                         End Sub))

                    End If
                End If
            End If



        Catch ex As Exception

        End Try

    End Sub

    Private Sub PictureBox9_Click(sender As Object, e As EventArgs) Handles Src_selection.Click
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


                Dim ran As Excel.Range = src_rng(1, 1)





            End If

        Catch ex As Exception

            Me.Show()
            TB_src_rng.Focus()

        End Try
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles Des_selection.Click
        Try
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
        Catch ex As Exception
            TB_des_rng.Focus()
        End Try
    End Sub



    Private Sub Btn_Cancel_Click(sender As Object, e As EventArgs) Handles Btn_Cancel.Click
        Me.Dispose()

    End Sub



    Private Sub TB_src_rng_TextChanged(sender As Object, e As EventArgs) Handles TB_src_rng.TextChanged
        Try

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

            End If

        Catch ex As Exception

        End Try
    End Sub
    Private Function IsValidExcelCellReference(cellReference As String) As Boolean

        ' Regular expression pattern for a cell reference.
        ' This pattern will match references like A1, $A$1, etc.
        Dim cellPattern As String = "(\$?[A-Z]+\$?[0-9]+)"

        ' Regular expression pattern for an Excel reference.
        ' This pattern will match references like A1:B13, $A$1:$B$13, A1, $B$1, etc.
        Dim singleReferencePattern As String = cellPattern + "(:" + cellPattern + ")?"

        ' Regular expression pattern to allow multiple cell references separated by commas
        Dim referencePattern As String = "^(" + singleReferencePattern + ")(," + singleReferencePattern + ")*$"

        ' Create a regex object with the pattern.
        Dim regex As New Regex(referencePattern)

        ' Test the input string against the regex pattern.
        Return regex.IsMatch(cellReference)

    End Function


    Private Sub TB_des_rng_TextChanged(sender As Object, e As EventArgs) Handles TB_des_rng.TextChanged
        Try

            If TB_des_rng.Text IsNot Nothing And IsValidExcelCellReference(TB_des_rng.Text) = True Then
                focuschange = True

                ' Define the range of cells to read (for example, cells A1 to A10)
                des_rng = excelApp.Range(TB_des_rng.Text)
                des_rng.Select()
                Dim range As Excel.Range = des_rng

                Me.Activate()
                'TB_src_range.Focus()
                TB_des_rng.SelectionStart = TB_des_rng.Text.Length
                focuschange = False

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub source(sender As Object, e As KeyEventArgs) Handles Src_selection.KeyDown

        Try
            If e.KeyCode = Keys.Enter Then

                Call Btn_OK_Click(sender, e)

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub Destination(sender As Object, e As KeyEventArgs) Handles Des_selection.KeyDown

        Try
            If e.KeyCode = Keys.Enter Then

                Call Btn_OK_Click(sender, e)

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub source_TextBox(sender As Object, e As KeyEventArgs) Handles TB_src_rng.KeyDown

        Try
            If e.KeyCode = Keys.Enter Then

                Call Btn_OK_Click(sender, e)

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub destination_TextBox(sender As Object, e As KeyEventArgs) Handles TB_des_rng.KeyDown

        Try
            If e.KeyCode = Keys.Enter Then

                Call Btn_OK_Click(sender, e)

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub form_enter(sender As Object, e As KeyEventArgs) Handles Me.KeyDown

        Try
            If e.KeyCode = Keys.Enter Then

                Call Btn_OK_Click(sender, e)

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub Combobox1_enter(sender As Object, e As KeyEventArgs) Handles ComboBox1.KeyDown

        Try
            If e.KeyCode = Keys.Enter Then

                Call Btn_OK_Click(sender, e)

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub Form34_PictureBasedDropdownList_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        form_flag = False
    End Sub

    Private Sub Form34_PictureBasedDropdownList_Disposed(sender As Object, e As EventArgs) Handles Me.Disposed
        form_flag = False
    End Sub

    Private Sub Form34_PictureBasedDropdownList_Shown(sender As Object, e As EventArgs) Handles Me.Shown

        Me.Focus()
        Me.BringToFront()
        Me.Activate()
        Try
            If TB_src_rng.Text <> "" Then

                Me.BeginInvoke(New System.Action(Sub()
                                                 TB_src_rng.Text = src_rng.Address
                                                 SetWindowPos(Me.Handle, New IntPtr(HWND_TOPMOST), 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOSIZE)
                                             End Sub))


            End If

        Catch ex As Exception
            TB_src_rng.Focus()

        End Try
    End Sub

End Class