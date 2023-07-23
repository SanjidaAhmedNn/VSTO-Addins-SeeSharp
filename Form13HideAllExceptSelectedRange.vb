Imports System.ComponentModel.Design
Imports System.Security.Cryptography
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Excel
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Drawing



Public Class Form13HideAllExceptSelectedRange
    Dim excelApp As Excel.Application
    Dim workbook As Excel.Workbook
    Dim worksheet As Excel.Worksheet
    Dim outWorksheet As Excel.Worksheet
    Dim inputRng As Excel.Range
    Dim outputRng As Excel.Range



    Private Sub Form13HideAllExceptSelectedRange_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub txtSourceRange_TextChanged(sender As Object, e As EventArgs) Handles txtSourceRange.TextChanged
        Try
            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            worksheet = workbook.ActiveSheet



            inputRng = worksheet.Range(txtSourceRange.Text)
            inputRng.Select()

            Call Display()



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




            'If checkBoxCopyWorksheet.Checked = True Then

            '    workbook.Worksheets.Copy(After:=workbook.Sheets(workbook.Sheets.Count))
            '    outWorksheet = workbook.Sheets(workbook.Sheets.Count)
            '    outWorksheet.Range("A1").Select()

            'End If

            'worksheet = workbook.Sheets(inputWsName)
            'worksheet.Activate()

            'Dim firstRow, lastRow, firstColumn, lastColumn As Integer
            'Dim selectedRng As Excel.Range

            'selectedRng = worksheet.Range(txtSourceRange.Text)
            'firstRow = selectedRng.Row
            'lastRow = firstRow + selectedRng.Rows.Count - 1
            'firstColumn = selectedRng.Column
            'lastColumn = firstColumn + selectedRng.Columns.Count - 1


            'worksheet.Range(worksheet.Cells(1, 1), worksheet.Cells(firstRow - 1, 1)).EntireRow.Hidden = True
            'worksheet.Range(worksheet.Cells(lastRow + 1, 1), worksheet.Cells(firstRow, firstColumn).end(Excel.XlDirection.xlDown)).EntireRow.Hidden = True
            'worksheet.Range(worksheet.Cells(1, 1), worksheet.Cells(1, firstColumn - 1)).EntireColumn.Hidden = True
            'worksheet.Range(worksheet.Cells(1, lastColumn + 1), worksheet.Cells(firstRow, firstColumn).end(Excel.XlDirection.xlToRight)).EntireColumn.Hidden = True


            'Me.Dispose()


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

            If checkBoxCopyWorksheet.Checked = True Then

                workbook.Worksheets.Copy(After:=workbook.Sheets(workbook.Sheets.Count))
                outWorksheet = workbook.Sheets(workbook.Sheets.Count)
                outWorksheet.Range("A1").Select()

            End If

            worksheet = workbook.Sheets(inputWsName)
            worksheet.Activate()

            Dim firstRow, lastRow, firstColumn, lastColumn As Integer
            Dim selectedRng As Excel.Range

            selectedRng = worksheet.Range(txtSourceRange.Text)
            firstRow = selectedRng.Row
            lastRow = firstRow + selectedRng.Rows.Count - 1
            firstColumn = selectedRng.Column
            lastColumn = firstColumn + selectedRng.Columns.Count - 1


            worksheet.Range(worksheet.Cells(1, 1), worksheet.Cells(firstRow - 1, 1)).EntireRow.Hidden = True
            worksheet.Range(worksheet.Cells(lastRow + 1, 1), worksheet.Cells(firstRow, firstColumn).end(Excel.XlDirection.xlDown)).EntireRow.Hidden = True
            worksheet.Range(worksheet.Cells(1, 1), worksheet.Cells(1, firstColumn - 1)).EntireColumn.Hidden = True
            worksheet.Range(worksheet.Cells(1, lastColumn + 1), worksheet.Cells(firstRow, firstColumn).end(Excel.XlDirection.xlToRight)).EntireColumn.Hidden = True


            Me.Dispose()


        Catch ex As Exception

        End Try


    End Sub

    Private Sub multiRng()

        Dim inputWsName As String
        excelApp = Globals.ThisAddIn.Application
        workbook = excelApp.ActiveWorkbook
        worksheet = workbook.ActiveSheet
        inputWsName = worksheet.Name



        Try
            Dim WsName As String
            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            worksheet = workbook.ActiveSheet
            WsName = worksheet.Name

            If checkBoxCopyWorksheet.Checked = True Then

                workbook.Worksheets.Copy(After:=workbook.Sheets(workbook.Sheets.Count))
                outWorksheet = workbook.Sheets(workbook.Sheets.Count)
                outWorksheet.Range("A1").Select()

            End If

            worksheet = workbook.Sheets(WsName)
            worksheet.Activate()


            Dim arrRng As String() = Split(txtSourceRange.Text, ",")
            Dim visRows, followingRows As Integer

            For i = 0 To UBound(arrRng) + 1


                If i > UBound(arrRng) Then
                    worksheet.Range(worksheet.Cells(followingRows + 1, 1), worksheet.Cells(followingRows + 1, 1).end(Excel.XlDirection.xlDown)).EntireRow.Hidden = True
                End If

                visRows = worksheet.Range(arrRng(i)).Row

                If i = 0 Then
                    worksheet.Range(worksheet.Cells(1, 1), worksheet.Cells(visRows - 1, 1)).EntireRow.Hidden = True
                Else
                    worksheet.Range(worksheet.Cells(followingRows + 1, 1), worksheet.Cells(visRows - 1, 1)).EntireRow.Hidden = True
                End If

                followingRows = visRows + worksheet.Range(arrRng(i)).Rows.Count - 1

                'integrate visColumns



            Next





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


End Class