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


            'Single rows Or Columns validation
            If selectedRng.Rows.Count <= 2 And selectedRng.Columns.Count <= 2 Then
                Dim answer As MsgBoxResult
                answer = MsgBox("You are about to hide all cells except " & selectedRng.Rows.Count & " Rows and " & selectedRng.Columns.Count & " Columns." & vbCrLf & "Do you want to proceed?", MsgBoxStyle.YesNo, "Warning!")
                If answer = MsgBoxResult.Yes Then
                    GoTo Proceed
                Else
                    GoTo break
                End If
            End If

Proceed:
            worksheet.Range(worksheet.Cells(firstRowNum + 1, firstColNum), worksheet.Cells(lastRowNum, 1)).EntireRow.Hidden = True
            worksheet.Range(worksheet.Cells(firstRowNum + 1, firstColNum), worksheet.Cells(1, lastColNum)).EntireColumn.Hidden = True

            selectedRng.EntireRow.Hidden = False
            selectedRng.EntireColumn.Hidden = False
            selectedRng.Select()

break:

            Me.Dispose()


        Catch ex As Exception

        End Try


    End Sub

    Private Sub multiRng()

        Dim WsName As String
        excelApp = Globals.ThisAddIn.Application
        workbook = excelApp.ActiveWorkbook
        worksheet = workbook.ActiveSheet
        WsName = worksheet.Name



        Try

            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            worksheet = workbook.ActiveSheet

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

            worksheet.Range(worksheet.Cells(firstRowNum + 1, firstColNum), worksheet.Cells(lastRowNum, 1)).EntireRow.Hidden = True
            worksheet.Range(worksheet.Cells(firstRowNum + 1, firstColNum), worksheet.Cells(1, lastColNum)).EntireColumn.Hidden = True



            Dim arrRng As String() = Split(txtSourceRange.Text, ",")

            For i = 0 To UBound(arrRng)

                worksheet.Range(arrRng(i)).EntireRow.Hidden = False
                worksheet.Range(arrRng(i)).EntireColumn.Hidden = False

            Next


            worksheet.Range(arrRng(0)).Cells(1, 1).select()


            Me.Dispose()

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


End Class