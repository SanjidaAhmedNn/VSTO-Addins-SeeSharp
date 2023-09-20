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
    Dim textChanged As Boolean = False

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

            textChanged = True

            inputRng = worksheet.Range(txtSourceRange.Text)
            inputRng.Select()


        Catch ex As Exception

        End Try

        textChanged = False
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

            If textChanged = False Then

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
            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            worksheet = workbook.ActiveSheet

            Dim selectedRng As Excel.Range
            selectedRng = worksheet.Range(txtSourceRange.Text)

            Dim temp As String
            temp = txtSourceRange.Text
            worksheet1 = inputRng.Worksheet

            If CheckBox.Checked = True Then

                workbook.ActiveSheet.Copy(After:=workbook.Sheets(workbook.Sheets.Count))
                outWorksheet = workbook.Sheets(workbook.Sheets.Count)

                worksheet1.Activate()
                txtSourceRange.Text = temp

            End If

            If selectedRng.Rows.Count <= 2 And selectedRng.Columns.Count <= 2 Then
                Dim answer As MsgBoxResult
                answer = MsgBox("You are about to set Scroll Area for only " & selectedRng.Rows.Count & " Rows and " & selectedRng.Columns.Count & " Columns." & vbCrLf & "Do you want to proceed?", MsgBoxStyle.YesNo, "Warning!")
                If answer = MsgBoxResult.Yes Then
                    GoTo Proceed
                Else
                    GoTo break
                End If
            End If

Proceed:
            worksheet.Rows.Hidden = True
            worksheet.Columns.Hidden = True

            For i As Integer = 1 To selectedRng.Rows.Count
                selectedRng.Rows(i).EntireRow.Hidden = False
            Next

            For i As Integer = 1 To selectedRng.Columns.Count
                selectedRng.Columns(i).EntireColumn.Hidden = False
            Next

            selectedRng.Select()

break:

            Me.Dispose()


        Catch ex As Exception

        End Try


    End Sub


    Private Sub multiRng()

        excelApp = Globals.ThisAddIn.Application
        workbook = excelApp.ActiveWorkbook
        worksheet = workbook.ActiveSheet

        Try

            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            worksheet = workbook.ActiveSheet


            Dim temp As String
            temp = txtSourceRange.Text
            worksheet1 = inputRng.Worksheet

            If CheckBox.Checked = True Then

                workbook.ActiveSheet.Copy(After:=workbook.Sheets(workbook.Sheets.Count))
                outWorksheet = workbook.Sheets(workbook.Sheets.Count)

                worksheet1.Activate()
                txtSourceRange.Text = temp

            End If

            Dim arrRng As String() = Split(txtSourceRange.Text, ",")

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

            worksheet.Rows.Hidden = True
            worksheet.Columns.Hidden = True


            For i As Integer = minRow To maxRow
                worksheet.Rows(i).EntireRow.Hidden = False
            Next

            For i As Integer = minCol To maxCol
                worksheet.Columns(i).EntireColumn.Hidden = False
            Next

            scrollArea.Select()


            Me.Dispose()

        Catch ex As Exception

        End Try

    End Sub

End Class