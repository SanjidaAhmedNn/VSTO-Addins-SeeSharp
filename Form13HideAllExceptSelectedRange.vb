Imports System.ComponentModel.Design
Imports System.Security.Cryptography
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Excel
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Drawing
Imports System.ComponentModel
Imports System.Linq.Expressions

Public Class Form13HideAllExceptSelectedRange
    Dim WithEvents excelApp As Excel.Application
    Dim workbook As Excel.Workbook
    Dim worksheet As Excel.Worksheet
    Dim outWorksheet As Excel.Worksheet
    Dim inputRng As Excel.Range
    Dim FocusedTxtBox As Integer

    Private Sub Form13HideAllExceptSelectedRange_Load(sender As Object, e As EventArgs) Handles MyBase.Load

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




            inputRng = worksheet.Range(txtSourceRange.Text)
            txtSourceRange.Focus()

            'inputRng.Select()


        Catch ex As Exception

        End Try


    End Sub

    Private Sub pctBoxSelectRange_Click(sender As Object, e As EventArgs) Handles pctBoxSelectRange.Click

        Try

            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            worksheet = workbook.ActiveSheet


            inputRng = excelApp.InputBox("Please Select a Range", "Range Selection", Type:=8)




            inputRng.Select()

            txtSourceRange.Text = inputRng.Address
            txtSourceRange.Focus()

        Catch ex As Exception

            txtSourceRange.Focus()

        End Try




    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click

        Me.Dispose()

    End Sub

    Private Sub btnOK_Click(sender As Object, e As EventArgs) Handles btnOK.Click
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






            Dim firstRow, lastRow, firstColumn, lastColumn As Integer
            Dim selectedRng As Excel.Range

            selectedRng = worksheet.Range(txtSourceRange.Text)
            firstRow = selectedRng.Row
            lastRow = firstRow + selectedRng.Rows.Count - 1
            firstColumn = selectedRng.Column
            lastColumn = firstColumn + selectedRng.Columns.Count - 1

            'Single rows Or Columns validation
            If selectedRng.Rows.Count <= 3 And selectedRng.Columns.Count <= 3 Then
                Dim answer As MsgBoxResult
                answer = MsgBox("You are about to hide all cells except " & selectedRng.Rows.Count & " Rows and " & selectedRng.Columns.Count & " Columns." & vbCrLf & "Do you want to proceed?", MsgBoxStyle.YesNo, "Warning!")
                If answer = MsgBoxResult.Yes Then
                    GoTo Proceed
                Else
                    GoTo break
                End If
            End If

Proceed:
            worksheet.Range(worksheet.Cells(1, 1), worksheet.Cells(firstRow - 1, 1)).EntireRow.Hidden = True

            If worksheet.Range(worksheet.Cells(lastRow, firstColumn), worksheet.Cells(1048576, firstColumn).end(Excel.XlDirection.xlUp)).Rows.Count = 1 Then

                worksheet.Range(worksheet.Cells(lastRow + 1, 1), worksheet.Cells(lastRow + 1, 1)).EntireRow.Hidden = False

            Else
                worksheet.Range(worksheet.Cells(lastRow + 1, 1), worksheet.Cells(firstRow, firstColumn).end(Excel.XlDirection.xlDown)).EntireRow.Hidden = True

            End If

            worksheet.Range(worksheet.Cells(1, 1), worksheet.Cells(1, firstColumn - 1)).EntireColumn.Hidden = True

            If worksheet.Range(worksheet.Cells(firstRow, lastColumn), worksheet.Cells(firstRow, 16384).end(Excel.XlDirection.xlToLeft)).Columns.Count = 1 Then
                worksheet.Range(worksheet.Cells(1, lastColumn + 1), worksheet.Cells(1, lastColumn + 1)).EntireColumn.Hidden = False
            Else

                worksheet.Range(worksheet.Cells(1, lastColumn + 1), worksheet.Cells(firstRow, firstColumn).end(Excel.XlDirection.xlToRight)).EntireColumn.Hidden = True


            End If




            If checkBoxCopyWorksheet.Checked = True Then

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

                    If worksheet.Range(worksheet.Cells(followingRows, visColumns), worksheet.Cells(1048576, visColumns).end(Excel.XlDirection.xlUp)).Rows.Count = 1 Then

                        worksheet.Range(worksheet.Cells(followingRows, visColumns), worksheet.Cells(followingRows, visColumns)).EntireRow.Hidden = False
                    Else
                        worksheet.Range(worksheet.Cells(followingRows + 1, visColumns), worksheet.Cells(followingRows, visColumns).end(Excel.XlDirection.xlDown)).EntireRow.Hidden = True

                    End If

                    Exit For
                End If

                visRows = worksheet.Range(arrRng(i)).Row
                visColumns = worksheet.Range(arrRng(i)).Column
                followingColumns = visColumns + worksheet.Range(arrRng(i)).Columns.Count - 1

                If i = 0 Then
                    worksheet.Range(worksheet.Cells(1, 1), worksheet.Cells(visRows - 1, 1)).EntireRow.Hidden = True
                    worksheet.Range(worksheet.Cells(1, 1), worksheet.Cells(1, visColumns - 1)).EntireColumn.Hidden = True

                    If worksheet.Range(worksheet.Cells(visRows, followingColumns), worksheet.Cells(visRows, 16384).End(XlDirection.xlToLeft)).Columns.Count = 1 Then

                        worksheet.Range(worksheet.Cells(visRows, followingColumns), worksheet.Cells(visRows, followingColumns)).EntireColumn.Hidden = False
                    Else
                        worksheet.Range(worksheet.Cells(visRows, followingColumns + 1), worksheet.Cells(visRows, followingColumns).End(XlDirection.xlToRight)).EntireColumn.Hidden = True

                    End If

                Else
                    worksheet.Range(worksheet.Cells(followingRows + 1, 1), worksheet.Cells(visRows - 1, 1)).EntireRow.Hidden = True
                    worksheet.Range(worksheet.Cells(visRows, 1), worksheet.Cells(visRows, visColumns - 1)).EntireColumn.Hidden = True

                    If worksheet.Range(worksheet.Cells(visRows, followingColumns), worksheet.Cells(visRows, 16384).End(XlDirection.xlToLeft)).Columns.Count = 1 Then

                        worksheet.Range(worksheet.Cells(visRows, followingColumns), worksheet.Cells(visRows, followingColumns)).EntireColumn.Hidden = False
                    Else
                        worksheet.Range(worksheet.Cells(visRows, followingColumns + 1), worksheet.Cells(visRows, followingColumns).End(XlDirection.xlToRight)).EntireColumn.Hidden = True

                    End If


                End If

                followingRows = visRows + worksheet.Range(arrRng(i)).Rows.Count - 1



            Next


            If checkBoxCopyWorksheet.Checked = True Then

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

    Private Sub Display()

        custPanInputRange.Controls.Clear()
        custPanExcpectedOutput.Controls.Clear()

        excelApp = Globals.ThisAddIn.Application
        workbook = excelApp.ActiveWorkbook
        worksheet = workbook.ActiveSheet
        Dim nRows, nColumns, unitHeight, unitWidth As Integer


        inputRng = worksheet.Range(txtSourceRange.Text)
        If inputRng.Rows.Count > 50 Then

            inputRng = worksheet.Range(inputRng.Cells(1, 1), inputRng.Cells(50, inputRng.Columns.Count))

        ElseIf inputRng.Columns.Count > 50 Then

            inputRng = worksheet.Range(inputRng.Cells(1, 1), inputRng.Cells(inputRng.Rows.Count, 50))

        End If

        nRows = inputRng.Rows.Count
        nColumns = inputRng.Columns.Count

        If nRows > 1 And nRows <= 6 Then
            unitHeight = custPanInputRange.Height / nRows
        Else
            unitHeight = custPanInputRange.Height / 6
        End If

        If nColumns > 1 And nColumns <= 6 Then
            unitWidth = custPanInputRange.Width / nColumns
        Else
            unitWidth = custPanInputRange.Width / 6
        End If

        For i = 1 To nRows
            For j = 1 To nColumns
                Dim label As New System.Windows.Forms.Label
                label.Text = inputRng.Cells(i, j).Value
                If nRows <> 1 And nColumns = 1 Then
                    label.Location = New System.Drawing.Point((3.25 - 1) * unitWidth, (i - 1) * unitHeight)
                ElseIf nRows = 1 And nColumns <> 1 Then
                    label.Location = New System.Drawing.Point((j - 1) * unitWidth, (3.5 - 1) * unitHeight)
                Else
                    label.Location = New System.Drawing.Point((j - 1) * unitWidth, (i - 1) * unitHeight)
                End If

                label.Height = unitHeight
                label.Width = unitWidth
                label.BorderStyle = BorderStyle.FixedSingle
                label.TextAlign = ContentAlignment.MiddleCenter

                custPanInputRange.Controls.Add(label)


            Next
        Next

        custPanInputRange.AutoScroll = True



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
            Dim selectedRange As Excel.Range
            selectedRange = excelApp.Selection
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


End Class