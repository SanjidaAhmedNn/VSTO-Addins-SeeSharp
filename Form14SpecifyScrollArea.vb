Imports System.ComponentModel.Design
Imports System.Security.Cryptography
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Excel
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Drawing
Imports System.ComponentModel
Imports System.Linq.Expressions

Public Class Form14SpecifyScrollArea

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
            Btn_OK.PerformClick()
        End If
    End Sub

    Private Sub Form14SpecifyScrollArea_Load(sender As Object, e As EventArgs) Handles MyBase.Load

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

    Private Sub Selection_Click(sender As Object, e As EventArgs) Handles Selection.Click

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

    Private Sub Btn_Cancel_Click(sender As Object, e As EventArgs) Handles Btn_Cancel.Click

        Me.Dispose()

    End Sub

    Public Function IsValidRng(input As String) As Boolean

        Dim pattern As String = "^(\$?[A-Z]+\$?[0-9]+(:\$?[A-Z]+\$?[0-9]+)?)(,\$?[A-Z]+\$?[0-9]+(:\$?[A-Z]+\$?[0-9]+)?)*$"
        Return System.Text.RegularExpressions.Regex.IsMatch(input, pattern)

    End Function


    Private Sub Btn_OK_Click(sender As Object, e As EventArgs) Handles Btn_OK.Click

        Try
            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            worksheet = workbook.ActiveSheet


            'checks if the user clicked OK button with an empty sourceRange textbox
            'if it is non-empty, then checks is the used range is a valid range or not
            'if any of these are true then it will give user another chance to enter correct input
            If txtSourceRange.Text = "" Then
                MsgBox("Please provide source range.", MsgBoxStyle.Exclamation, "Error!")
                txtSourceRange.Focus()
                Exit Sub
            ElseIf IsValidRng(txtSourceRange.Text.ToUpper) = False Then
                MsgBox("Please provide a valid source range.", MsgBoxStyle.Exclamation, "Error!")
                txtSourceRange.Text = ""
                txtSourceRange.Focus()
                Exit Sub
            End If

            'counts the number of ranges used by user
            Dim rngCount As Integer
            rngCount = 0

            For Each c As Char In txtSourceRange.Text

                If c = "," Then
                    rngCount = rngCount + 1
                End If

            Next

            'calls different subs based on number of ranges in users' selection 
            If rngCount = 0 Then
                Call singleRng()
            Else
                Call multiRng()
            End If


        Catch ex As Exception

        End Try



    End Sub

    Private Sub singleRng()

        'this sub will be called when user selected a single range as input

        Try
            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            worksheet = workbook.ActiveSheet

            Dim selectedRng As Excel.Range
            selectedRng = worksheet.Range(txtSourceRange.Text)

            'keeps the range address from the textbox in a variable and keeps the worksheet info in another variable named "worksheet1"
            Dim temp As String
            temp = txtSourceRange.Text
            worksheet1 = inputRng.Worksheet

            'checks if user opted to backup the sheet. If yes then create a copy and reactivate the original worksheet
            If CheckBox.Checked = True Then

                workbook.ActiveSheet.Copy(After:=workbook.Sheets(workbook.Sheets.Count))
                outWorksheet = workbook.Sheets(workbook.Sheets.Count)

                worksheet1.Activate()
                txtSourceRange.Text = temp

            End If

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
            worksheet.Rows.Hidden = True
            worksheet.Columns.Hidden = True


            selectedRng.EntireRow.Hidden = False
            selectedRng.EntireColumn.Hidden = False

            selectedRng.Select()

break:

            Me.Dispose()


        Catch ex As Exception

        End Try


    End Sub


    Private Sub multiRng()

        'this sub will be called when user selected multiple ranges as input

        excelApp = Globals.ThisAddIn.Application
        workbook = excelApp.ActiveWorkbook
        worksheet = workbook.ActiveSheet

        Try

            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            worksheet = workbook.ActiveSheet

            'keeps the range address from the textbox in a variable and keeps the worksheet info in another variable named "worksheet1"
            Dim temp As String
            temp = txtSourceRange.Text
            worksheet1 = inputRng.Worksheet

            'checks if user opted to backup the sheet. If yes then create a copy and reactivate the original worksheet
            If CheckBox.Checked = True Then

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
            Dim scrollArea As Excel.Range = worksheet.Range(worksheet.Cells(minRow, minCol), worksheet.Cells(maxRow, maxCol))


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
                worksheet.Rows.Hidden = True
                worksheet.Columns.Hidden = True

                For i = 0 To UBound(arrRng)
                    worksheet.Range(arrRng(i)).EntireRow.Hidden = False
                    worksheet.Range(arrRng(i)).EntireColumn.Hidden = False
                Next

                scrollArea.Select()
                Me.Dispose()

            End If


        Catch ex As Exception

        End Try

    End Sub
    Private Sub Form14SpecifyScrollArea_Disposed(sender As Object, e As EventArgs) Handles Me.Disposed
        form_flag = False
    End Sub

    Private Sub Form14SpecifyScrollArea_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        form_flag = False
    End Sub



    Private Sub Form14SpecifyScrollArea_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        Me.Focus()
        Me.BringToFront()
        Me.Activate()
        Me.BeginInvoke(New System.Action(Sub()
                                             txtSourceRange.Text = inputRng.Address
                                             SetWindowPos(Me.Handle, New IntPtr(HWND_TOPMOST), 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOSIZE)
                                         End Sub))
    End Sub
End Class