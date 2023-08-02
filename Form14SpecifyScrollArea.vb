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
    Dim worksheet As Excel.Worksheet
    Dim outWorksheet As Excel.Worksheet
    Dim inputRng As Excel.Range
    Dim FocusedTxtBox As Integer
    Dim selectedRange As Excel.Range


    Private Sub Form14SpecifyScrollArea_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        excelApp = Globals.ThisAddIn.Application
        workbook = excelApp.ActiveWorkbook
        worksheet = workbook.ActiveSheet

        Dim selectedRng As Excel.Range = excelApp.Selection
        txtSourceRange.Text = selectedRng.Address



    End Sub

    Private Sub txtSourceRange_TextChanged(sender As Object, e As EventArgs) Handles txtSourceRange.TextChanged

        Try
            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            worksheet = workbook.ActiveSheet


            txtSourceRange.Focus()


            inputRng = worksheet.Range(txtSourceRange.Text)



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


            If FocusedTxtBox = 1 Then

                txtSourceRange.Text = selectedRange.Address
                worksheet = workbook.ActiveSheet
                inputRng = selectedRange
                txtSourceRange.Focus()

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

            inputRng = excelApp.InputBox("Please Select a Range", "Range Selection", selectedRange.Address, Type:=8)
            inputRng.Select()
            txtSourceRange.Text = inputRng.Address
            txtSourceRange.Focus()


        Catch ex As Exception

            txtSourceRange.Focus()

        End Try


    End Sub

    Private Sub Btn_Cancel_Click(sender As Object, e As EventArgs) Handles Btn_Cancel.Click

        Me.Dispose()

    End Sub

    Private Sub Btn_OK_Click(sender As Object, e As EventArgs) Handles Btn_OK.Click

        Try
            Dim inputWsName As String
            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            worksheet = workbook.ActiveSheet
            inputWsName = worksheet.Name

            Dim rngCount As Integer
            rngCount = 0

            For Each c As Char In txtSourceRange.Text

                If c = "," Then
                    rngCount = rngCount + 1
                End If

            Next

            Dim arrRng As String() = Split(txtSourceRange.Text, ",")
            For i = 0 To UBound(arrRng) - 1
                If Not worksheet.Range(arrRng(i)).Column = worksheet.Range(arrRng(i + 1)).Column And worksheet.Range(arrRng(i)).Columns.Count = worksheet.Range(arrRng(i + 1)).Columns.Count Then
                    Dim columnMisMatchAnswer As MsgBoxResult
                    columnMisMatchAnswer = MsgBox("To continue, the number of Columns must be same in each selection." & vbCrLf & "Please retry!", MsgBoxStyle.OkOnly, "Warning!")

                    GoTo break
                End If
            Next


            If rngCount = 0 Then

                Call singleRng()
            Else
                Call multiRng()
            End If

break:

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






            Dim firstRow, lastRow, firstColumn, lastColumn As Integer
            Dim selectedRng As Excel.Range

            selectedRng = worksheet.Range(txtSourceRange.Text)
            firstRow = selectedRng.Row
            lastRow = firstRow + selectedRng.Rows.Count - 1
            firstColumn = selectedRng.Column
            lastColumn = firstColumn + selectedRng.Columns.Count - 1

            'Single rows Or Columns validation
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
            worksheet.Range(worksheet.Cells(1, 1), worksheet.Cells(firstRow - 1, 1)).EntireRow.Hidden = True
            worksheet.Range(worksheet.Cells(lastRow + 1, 1), worksheet.Cells(lastRow + 1, 1).end(Excel.XlDirection.xlDown)).EntireRow.Hidden = True


            worksheet.Range(worksheet.Cells(1, 1), worksheet.Cells(1, firstColumn - 1)).EntireColumn.Hidden = True
            worksheet.Range(worksheet.Cells(1, lastColumn + 1), worksheet.Cells(1, lastColumn + 1).end(Excel.XlDirection.xlToRight)).EntireColumn.Hidden = True

            If CheckBox.Checked = True Then

                workbook.ActiveSheet.Copy(After:=workbook.Sheets(workbook.Sheets.Count))
                outWorksheet = workbook.Sheets(workbook.Sheets.Count)
                outWorksheet.Range("A1").Select()
                worksheet.Cells.EntireColumn.Hidden = False
                worksheet.Cells.EntireRow.Hidden = False

                worksheet = workbook.Sheets(inputWsName)
                worksheet.Activate()

            End If

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

            Dim arrRng As String() = Split(txtSourceRange.Text, ",")
            Dim visRows, followingRows As Integer
            Dim visColumns, followingColumns As Integer


            For i = 0 To UBound(arrRng) + 1


                If i > UBound(arrRng) Then

                    worksheet.Range(worksheet.Cells(followingRows + 1, visColumns), worksheet.Cells(1048576, visColumns)).EntireRow.Hidden = True
                    worksheet.Range(worksheet.Cells(followingRows, followingColumns + 1), worksheet.Cells(followingRows, followingColumns + 1).end(Excel.XlDirection.xlToRight)).EntireColumn.Hidden = True

                    Exit For

                End If

                visRows = worksheet.Range(arrRng(i)).Row
                visColumns = worksheet.Range(arrRng(i)).Column
                followingColumns = visColumns + worksheet.Range(arrRng(i)).Columns.Count - 1

                If i = 0 Then
                    worksheet.Range(worksheet.Cells(1, 1), worksheet.Cells(visRows - 1, 1)).EntireRow.Hidden = True
                    worksheet.Range(worksheet.Cells(1, 1), worksheet.Cells(1, visColumns - 1)).EntireColumn.Hidden = True
                    worksheet.Range(worksheet.Cells(1, followingColumns + 1), worksheet.Cells(1, followingColumns + 1).End(XlDirection.xlToRight)).EntireColumn.Hidden = True



                Else
                    worksheet.Range(worksheet.Cells(followingRows + 1, 1), worksheet.Cells(visRows - 1, 1)).EntireRow.Hidden = True
                    worksheet.Range(worksheet.Cells(visRows, 1), worksheet.Cells(visRows, visColumns - 1)).EntireColumn.Hidden = True
                    worksheet.Range(worksheet.Cells(visRows, followingColumns + 1), worksheet.Cells(visRows, followingColumns + 1).End(XlDirection.xlToRight)).EntireColumn.Hidden = True


                End If

                followingRows = visRows + worksheet.Range(arrRng(i)).Rows.Count - 1



            Next


            If CheckBox.Checked = True Then

                workbook.ActiveSheet.Copy(After:=workbook.Sheets(workbook.Sheets.Count))
                outWorksheet = workbook.Sheets(workbook.Sheets.Count)
                outWorksheet.Range("A1").Select()
                worksheet.Cells.EntireColumn.Hidden = False
                worksheet.Cells.EntireRow.Hidden = False
                worksheet = workbook.Sheets(WsName)
                worksheet.Activate()

            End If




            Me.Dispose()

        Catch ex As Exception

        End Try



    End Sub

End Class