Imports System.ComponentModel.Design
Imports System.Security.Cryptography
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Excel
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Drawing
Imports System.ComponentModel
Imports System.Linq.Expressions
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Button

Public Class Form13HideAllExceptSelectedRange
    Dim WithEvents excelApp As Excel.Application
    Dim workbook As Excel.Workbook
    Dim worksheet, worksheet1 As Excel.Worksheet
    Dim outWorksheet As Excel.Worksheet
    Dim inputRng As Excel.Range
    Dim FocusedTxtBox As Integer
    Dim selectedRange As Excel.Range
    Dim txtChanged As Boolean = False

    Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As IntPtr, ByVal hWndInsertAfter As IntPtr, ByVal X As Integer, ByVal Y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal uFlags As UInteger) As Boolean
    Private Const SWP_NOMOVE As UInteger = &H2
    Private Const SWP_NOSIZE As UInteger = &H1
    Private Const SWP_NOACTIVATE As UInteger = &H10
    Private Const HWND_TOPMOST As Integer = -1


    Private Sub Form1_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Enter Then
            btnOK.PerformClick()
        End If
    End Sub

    Private Sub Form13HideAllExceptSelectedRange_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        excelApp = Globals.ThisAddIn.Application
        workbook = excelApp.ActiveWorkbook
        worksheet = workbook.ActiveSheet

        Dim selectedRng As Excel.Range = excelApp.Selection
        txtSourceRange.Text = selectedRng.Address

        Me.KeyPreview = True




    End Sub

    Private Sub txtSourceRange_TextChanged(sender As Object, e As EventArgs) Handles txtSourceRange.TextChanged
        Try
            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            worksheet = workbook.ActiveSheet



            txtChanged = True

            inputRng = worksheet.Range(txtSourceRange.Text)
            inputRng.Select()



        Catch ex As Exception

        End Try


        txtChanged = False
        txtSourceRange.Focus()

    End Sub

    Private Sub pctBoxSelectRange_Click(sender As Object, e As EventArgs) Handles pctBoxSelectRange.Click

        Try

            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            worksheet = workbook.ActiveSheet
            selectedRange = excelApp.Selection
            txtSourceRange.Focus()

            Me.Hide()
            inputRng = excelApp.InputBox("Please Select a Range", "Range Selection", selectedRange.Address, Type:=8)
            Me.Show()

            inputRng.Worksheet.Activate()

            txtSourceRange.Text = inputRng.Address

            inputRng.Select()
            txtSourceRange.Focus()

        Catch ex As Exception

            txtSourceRange.Focus()

        End Try




    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click

        Me.Dispose()

    End Sub
    Public Function IsValidRng(input As String) As Boolean

        Dim pattern As String = "^(\$?[A-Z]+\$?[0-9]+(:\$?[A-Z]+\$?[0-9]+)?)(,\$?[A-Z]+\$?[0-9]+(:\$?[A-Z]+\$?[0-9]+)?)*$"
        Return System.Text.RegularExpressions.Regex.IsMatch(input, pattern)

    End Function

    Private Sub btnOK_Click(sender As Object, e As EventArgs) Handles btnOK.Click
        Try
            Dim inputWsName As String
            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            worksheet = workbook.ActiveSheet
            inputWsName = worksheet.Name

            If txtSourceRange.Text = "" Then
                MsgBox("Please select the Source Range.", MsgBoxStyle.Exclamation, "Error!")
                txtSourceRange.Focus()
                Exit Sub
            ElseIf IsValidRng(txtSourceRange.Text.ToUpper) = False Then
                MsgBox("Please use a valid range.", MsgBoxStyle.Exclamation, "Error!")
                txtSourceRange.Text = ""
                txtSourceRange.Focus()
                Exit Sub
            End If


            Dim rngCount As Integer
            rngCount = 0

            For Each c As Char In txtSourceRange.Text

                If c = "," Then
                    rngCount = rngCount + 1
                End If

            Next

            If rngCount = 0 Then

                Call singleRng()
            Else
                Call multiRng()
            End If

            Me.Dispose()


        Catch ex As Exception

        End Try



    End Sub

    Private Sub singleRng()

        Try

            'this sub will be called when user selected a single range as input

            Dim inputWsName As String
            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            worksheet = workbook.ActiveSheet
            inputWsName = worksheet.Name
            Dim selectedRng As Excel.Range
            selectedRng = worksheet.Range(txtSourceRange.Text)



            Dim temp As String
            temp = txtSourceRange.Text
            worksheet1 = inputRng.Worksheet

            If checkBoxCopyWorksheet.Checked = True Then

                workbook.ActiveSheet.Copy(After:=workbook.Sheets(workbook.Sheets.Count))
                outWorksheet = workbook.Sheets(workbook.Sheets.Count)

                worksheet1.Activate()
                txtSourceRange.Text = temp

            End If


            Dim lastCell() As String
            Dim firstRowNum, firstColNum, lastRowNum, lastColNum As Integer

            lastCell = worksheet.UsedRange.Address.Split(":"c)
            firstRowNum = worksheet.Range(lastCell(0)).Row
            firstColNum = worksheet.Range(lastCell(0)).Column
            lastRowNum = worksheet.Range(lastCell(1)).Row
            lastColNum = worksheet.Range(lastCell(1)).Column

            Dim i As Integer

            'cellCount variable is used to count the number of cells in users' selection.
            'Our goal is to check whether the cellCount is <= 4 or not in the next block.
            'if the cellCount exceeds 5 then exit from the loop.
            Dim cellCount As Integer = 0
            For i = 1 To selectedRng.Rows.Count
                For j = 1 To selectedRng.Columns.Count
                    cellCount += 1
                    If cellCount > 5 Then Exit For
                Next
                If cellCount > 5 Then Exit For
            Next

            'checks if the cellCount is <=6 or not. If yes then show a YesNo msgbox as warning.
            'If user select yes then continue excecuting next lines, else dispose the form
            If cellCount <= 4 Then
                Dim answer As MsgBoxResult
                answer = MsgBox("Do you really want to hide everything except " & cellCount & " cells." & vbCrLf & "If yes, hide every cell except the selected cell range. If no, close the add-in.", MsgBoxStyle.YesNo, "Warning!")
                If answer = MsgBoxResult.Yes Then
                    GoTo Proceed
                Else
                    GoTo break
                End If
            End If

Proceed:


            If checkBox_Header.Checked = True Then


                'find first row with data and exit from loop after finding the first data
                For i = 1 To worksheet.Rows.Count
                    For j = 1 To worksheet.Columns.Count
                        If worksheet.Cells(i, j).value IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(worksheet.Cells(i, j).value.ToString()) Then
                            GoTo exitLoop
                        End If
                    Next
                Next
exitLoop:
                'hide all rows and columns of the used range of the worksheet
                worksheet.UsedRange.EntireRow.Hidden = True
                worksheet.UsedRange.EntireColumn.Hidden = True

                'unhide the header row
                worksheet.Rows(i).entirerow.hidden = False

                'unhide users' selected range
                selectedRng.EntireRow.Hidden = False
                selectedRng.EntireColumn.Hidden = False
                selectedRng = worksheet.Range(worksheet.Cells(i, selectedRng.Column), selectedRng.Cells(1, 1).offset(selectedRng.Rows.Count - 1, selectedRng.Columns.Count - 1))
                selectedRng.Select()

            Else

                'hide all rows and columns of the used range of the worksheet
                worksheet.UsedRange.EntireRow.Hidden = True
                worksheet.UsedRange.EntireColumn.Hidden = True

                'unhide users' selected range
                selectedRng.EntireRow.Hidden = False
                selectedRng.EntireColumn.Hidden = False
                selectedRng.Select()

            End If

break:

            Me.Dispose()

        Catch ex As Exception

        End Try


    End Sub

    Private Sub multiRng()

        'this sub will be called when user selected multiple ranges as input

        Try

            Dim WsName As String
            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            worksheet = workbook.ActiveSheet
            Dim selectedRng As Excel.Range
            selectedRng = worksheet.Range(txtSourceRange.Text)
            WsName = worksheet.Name

            'keeps the range address from the textbox in a variable and keeps the worksheet info in another variable named "worksheet1"
            Dim i As Integer
            Dim temp As String
            temp = txtSourceRange.Text
            worksheet1 = inputRng.Worksheet

            'checks if user opted to backup the sheet. If yes then create a copy and reactivate the original worksheet
            If checkBoxCopyWorksheet.Checked = True Then

                workbook.ActiveSheet.Copy(After:=workbook.Sheets(workbook.Sheets.Count))
                outWorksheet = workbook.Sheets(workbook.Sheets.Count)

                worksheet1.Activate()
                txtSourceRange.Text = temp

            End If

            'keeps each of the range addresses from users' selecion in separate array elements of the arrRng array
            Dim arrRng As String() = Split(txtSourceRange.Text, ",")

            'finds the start and end row, column numbers and store the range in scrollArea variable as range
            Dim minRow As Integer = Integer.MaxValue
            Dim maxRow As Integer = Integer.MinValue
            Dim minCol As Integer = Integer.MaxValue
            Dim maxCol As Integer = Integer.MinValue

            For Each address In arrRng
                Dim range As Excel.Range = worksheet.Range(address)
                minRow = Math.Min(minRow, range.Row)
                maxRow = Math.Max(maxRow, range.Row + range.Rows.Count - 1)
                minCol = Math.Min(minCol, range.Column)
                maxCol = Math.Max(maxCol, range.Column + range.Columns.Count - 1)
            Next
            Dim visibleRange As Excel.Range = worksheet.Range(worksheet.Cells(minRow, minCol), worksheet.Cells(maxRow, maxCol))


            'declare a booolean variable named "flag" with Fasle value
            'if the number of rows and the row number of 1st row of each range is same then flag will be True
            'if the number of columns and the column number of 1st column of each range is same then flag will be True
            'otherwise it flag will be false
            Dim flag As Boolean = False
            For i = 0 To UBound(arrRng) - 1
                If worksheet.Range(arrRng(i)).Rows.Count = worksheet.Range(arrRng(i + 1)).Rows.Count And worksheet.Range(arrRng(i)).Row = worksheet.Range(arrRng(i + 1)).Row Then
                    flag = True
                ElseIf worksheet.Range(arrRng(i)).Columns.Count = worksheet.Range(arrRng(i + 1)).Columns.Count And worksheet.Range(arrRng(i)).Column = worksheet.Range(arrRng(i + 1)).Column Then
                    flag = True
                Else
                    flag = False
                End If
            Next

            'checks if the flag is true or false
            'muiltiple ranges will be hidden only if the the flag is true
            'otherwise a msgbox will open and give user another chance to enter correct inputs
            If flag = False Then
                MsgBox("Multiple selection is not possible with this source range.", MsgBoxStyle.Exclamation, "Error!")
                txtSourceRange.Clear()
                txtSourceRange.Focus()
            Else

                If checkBox_Header.Checked = True Then
                    'find first row with data and exit from loop after finding the first data
                    For i = 1 To worksheet.Rows.Count
                        For j = 1 To worksheet.Columns.Count
                            If worksheet.Cells(i, j).value IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(worksheet.Cells(i, j).value.ToString()) Then
                                GoTo exitLoop
                            End If
                        Next
                    Next
exitLoop:
                    'hide all rows and columns of the used range of the worksheet
                    worksheet.UsedRange.EntireRow.Hidden = True
                    worksheet.UsedRange.EntireColumn.Hidden = True

                    'unhide the header row
                    worksheet.Rows(i).entirerow.hidden = False

                    'unhide users' selected ranges
                    For k = 0 To UBound(arrRng)
                        worksheet.Range(arrRng(k)).EntireRow.Hidden = False
                        worksheet.Range(arrRng(k)).EntireColumn.Hidden = False
                    Next

                    selectedRng = worksheet.Range(worksheet.Cells(i, minCol), worksheet.Cells(maxRow, maxCol))
                    selectedRng.Select()


                Else
                    'hide all rows and columns of the used range of the worksheet
                    worksheet.UsedRange.EntireRow.Hidden = True
                    worksheet.UsedRange.EntireColumn.Hidden = True

                    For k = 0 To UBound(arrRng)
                        worksheet.Range(arrRng(k)).EntireRow.Hidden = False
                        worksheet.Range(arrRng(k)).EntireColumn.Hidden = False
                    Next

                    visibleRange.Select()

                End If

                Me.Dispose()

            End If


        Catch ex As Exception

        End Try


    End Sub

    Private Sub txtSourceRange_GotFocus(sender As Object, e As EventArgs) Handles txtSourceRange.GotFocus
        Try

            FocusedTxtBox = 1


        Catch ex As Exception

        End Try
    End Sub
    Private Sub Form1_Activated(sender As Object, e As EventArgs) Handles Me.Activated
        Try

            excelApp = Globals.ThisAddIn.Application

            AddHandler excelApp.SheetSelectionChange, AddressOf rngSelectionFromTxtBox

        Catch ex As Exception

        End Try

    End Sub
    Private Sub rngSelectionFromTxtBox(ByVal Sh As Object, ByVal Target As Excel.Range)

        Try

            excelApp = Globals.ThisAddIn.Application
            selectedRange = excelApp.Selection
            selectedRange.Select()

            txtSourceRange.Focus()


            If txtChanged = False Then

                If FocusedTxtBox = 1 Then

                    txtSourceRange.Text = selectedRange.Address
                    worksheet = workbook.ActiveSheet
                    inputRng = selectedRange
                    txtSourceRange.Focus()


                End If

            End If



        Catch ex As Exception


        End Try

    End Sub

    Private Sub Form13HideAllExceptSelectedRange_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        form_flag = False
    End Sub

    Private Sub Form13HideAllExceptSelectedRange_Disposed(sender As Object, e As EventArgs) Handles Me.Disposed
        form_flag = False
    End Sub

    Private Sub Form13HideAllExceptSelectedRange_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        Me.Focus()
        Me.BringToFront()
        Me.Activate()
        Me.BeginInvoke(New System.Action(Sub()
                                             txtSourceRange.Text = inputRng.Address
                                             SetWindowPos(Me.Handle, New IntPtr(HWND_TOPMOST), 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOSIZE)
                                         End Sub))
    End Sub
End Class