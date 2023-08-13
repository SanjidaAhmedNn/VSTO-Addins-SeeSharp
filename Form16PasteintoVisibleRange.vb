Imports System.ComponentModel.Design
Imports System.Security.Cryptography
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Excel
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Drawing
Imports System.ComponentModel
Imports System.Linq.Expressions



Public Class Form16PasteintoVisibleRange
    Dim WithEvents excelApp As Excel.Application
    Dim workbook As Excel.Workbook
    Dim worksheet As Excel.Worksheet
    Dim outWorksheet As Excel.Worksheet
    Dim inputRng As Excel.Range
    Dim FocusedTxtBox As Integer
    Dim selectedRange As Excel.Range
    Dim destRange As Excel.Range
    Dim outputRng As Excel.Range
    Dim WsName As String

    Private Sub Form16PasteintoVisibleRange_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        excelApp = Globals.ThisAddIn.Application
        workbook = excelApp.ActiveWorkbook
        worksheet = workbook.ActiveSheet

        Dim selectedRng As Excel.Range = excelApp.Selection
        txtSourceRange.Text = selectedRng.Address
        txtSourceRange.Focus()


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

    Private Sub destinationSelection_Click(sender As Object, e As EventArgs) Handles destinationSelection.Click

        Try

            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            worksheet = workbook.ActiveSheet
            selectedRange = excelApp.Selection
            txtDestRange.Focus()

            outputRng = excelApp.InputBox("Please Select a Destination Range", "Destination Range Selection", selectedRange.Address, Type:=8)
            outputRng.Select()
            txtDestRange.Text = outputRng.Address
            txtDestRange.Focus()



        Catch ex As Exception

            txtDestRange.Focus()

        End Try


    End Sub



    Private Sub txtSourceRange_GotFocus(sender As Object, e As EventArgs) Handles txtSourceRange.GotFocus
        Try

            FocusedTxtBox = 1


        Catch ex As Exception

        End Try
    End Sub

    Private Sub txtDestRange_GotFocus(sender As Object, e As EventArgs) Handles txtDestRange.GotFocus

        Try

            FocusedTxtBox = 2


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


            If FocusedTxtBox = 1 Then

                txtSourceRange.Text = selectedRange.Address
                worksheet = workbook.ActiveSheet
                inputRng = selectedRange
                txtSourceRange.Focus()

            ElseIf FocusedTxtBox = 2 Then
                txtDestRange.Text = selectedRange.Address
                worksheet = workbook.ActiveSheet
                destRange = selectedRange
                txtDestRange.Focus()

            End If


        Catch ex As Exception


        End Try

    End Sub

    Private Sub AutoSelection_Click(sender As Object, e As EventArgs) Handles AutoSelection.Click

        Try

            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            worksheet = workbook.ActiveSheet
            selectedRange = excelApp.Selection
            selectedRange.Select()

            Dim topLeft, bottomRight As String



            If selectedRange.Offset(0, -1).Value = Nothing And selectedRange.Offset(0, 1).Value = Nothing And selectedRange.Offset(-1, 0).Value = Nothing Then
                topLeft = selectedRange.Address
                bottomRight = worksheet.Range(topLeft).End(XlDirection.xlDown).Address
                selectedRange = worksheet.Range(worksheet.Range(topLeft), worksheet.Range(bottomRight))

            ElseIf selectedRange.Offset(-1, 0).Value = Nothing And selectedRange.Offset(1, 0).Value = Nothing And selectedRange.Offset(0, -1).Value = Nothing Then

                topLeft = selectedRange.Address
                bottomRight = worksheet.Range(topLeft).End(XlDirection.xlToRight).Address
                selectedRange = worksheet.Range(worksheet.Range(topLeft), worksheet.Range(bottomRight))

            ElseIf selectedRange.Offset(0, -1).Value = Nothing And selectedRange.Offset(-1, 0).Value = Nothing Then
                bottomRight = selectedRange.End(XlDirection.xlToRight).Address
                bottomRight = worksheet.Range(bottomRight).End(XlDirection.xlDown).Address

                selectedRange = worksheet.Range(selectedRange, worksheet.Range(bottomRight))

            ElseIf selectedRange.Offset(0, -1).Value = Nothing And selectedRange.Offset(0, 1).Value = Nothing Then

                topLeft = selectedRange.End(XlDirection.xlUp).Address
                bottomRight = worksheet.Range(topLeft).End(XlDirection.xlDown).Address
                selectedRange = worksheet.Range(worksheet.Range(topLeft), worksheet.Range(bottomRight))

            ElseIf selectedRange.Offset(-1, 0).Value = Nothing And selectedRange.Offset(1, 0).Value = Nothing Then
                topLeft = selectedRange.End(XlDirection.xlToLeft).Address
                bottomRight = worksheet.Range(topLeft).End(XlDirection.xlToRight).Address
                selectedRange = worksheet.Range(worksheet.Range(topLeft), worksheet.Range(bottomRight))

            ElseIf selectedRange.Offset(0, -1).Value = Nothing Then
                topLeft = selectedRange.End(XlDirection.xlUp).Address
                bottomRight = worksheet.Range(topLeft).End(XlDirection.xlToRight).Address
                bottomRight = worksheet.Range(bottomRight).End(XlDirection.xlDown).Address
                selectedRange = worksheet.Range(worksheet.Range(topLeft), worksheet.Range(bottomRight))


            ElseIf selectedRange.Offset(-1, 0).Value = Nothing Then

                topLeft = selectedRange.End(XlDirection.xlToLeft).Address
                bottomRight = worksheet.Range(topLeft).End(XlDirection.xlToRight).Address
                bottomRight = worksheet.Range(bottomRight).End(XlDirection.xlDown).Address
                selectedRange = worksheet.Range(worksheet.Range(topLeft), worksheet.Range(bottomRight))



            Else
                topLeft = selectedRange.End(XlDirection.xlToLeft).Address
                topLeft = worksheet.Range(topLeft).End(XlDirection.xlUp).Address
                bottomRight = worksheet.Range(topLeft).End(XlDirection.xlToRight).Address
                bottomRight = worksheet.Range(bottomRight).End(XlDirection.xlDown).Address

                selectedRange = worksheet.Range(worksheet.Range(topLeft), worksheet.Range(bottomRight))


            End If

            selectedRange.Select()





        Catch ex As Exception

        End Try




    End Sub

    Private Sub txtDestRange_TextChanged(sender As Object, e As EventArgs) Handles txtDestRange.TextChanged

        Try
            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            worksheet = workbook.ActiveSheet


            txtDestRange.Focus()


            outputRng = worksheet.Range(txtDestRange.Text)



        Catch ex As Exception

        End Try


    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.Dispose()
    End Sub

    Private Sub btnOK_Click(sender As Object, e As EventArgs) Handles btnOK.Click
        Dim i, j, count, rngCount As Integer
        Dim lastRow As String
        excelApp = Globals.ThisAddIn.Application
        workbook = excelApp.ActiveWorkbook
        worksheet = workbook.ActiveSheet
        WsName = worksheet.Name


        If CB_copyWs.Checked = True Then

            workbook.ActiveSheet.Copy(After:=workbook.Sheets(workbook.Sheets.Count))
            outWorksheet = workbook.Sheets(workbook.Sheets.Count)


            worksheet = workbook.Sheets(WsName)
            worksheet.Activate()


        End If

        If CB_keepFormat.Checked = False Then

            inputRng.ClearFormats()

        End If




        inputRng = worksheet.Range(txtSourceRange.Text)

        outputRng = worksheet.Range(txtDestRange.Text).SpecialCells(Excel.XlCellType.xlCellTypeVisible)
        selectedRange = excelApp.Selection


        rngCount = 0

        For Each c As Char In outputRng.Address

            If c = "," Then
                rngCount = rngCount + 1
            End If

        Next


        Dim arrRng As String() = Split(outputRng.Address, ",")


        lastRow = worksheet.Range(arrRng(rngCount)).End(XlDirection.xlDown).Address
        lastRow = worksheet.Range(lastRow).End(XlDirection.xlUp).Address

        count = 0

        While worksheet.Range(lastRow).Offset(count, 0).Value <> Nothing

            count = count + 1

        End While









        If inputRng.Rows.Count - 1 <= rngCount Then

            For i = 0 To inputRng.Rows.Count - 1


                worksheet.Range(arrRng(i)).Value = inputRng.Cells.Offset(i, 0).Value


            Next


        Else

            For j = 0 To inputRng.Rows.Count
                If j <= rngCount Then
                    worksheet.Range(arrRng(j)).Value = inputRng.Cells.Offset(j, 0).Value
                ElseIf j > rngCount Then
                    For k = 0 To inputRng.Columns.Count - 1
                        worksheet.Range(lastRow).Offset(count, k).Value = inputRng.Cells.Offset(j, k).Value

                    Next
                    count = count + 1
                End If
            Next


        End If


        Me.Dispose()



    End Sub


End Class