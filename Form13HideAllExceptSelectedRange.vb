Public Class Form13HideAllExceptSelectedRange
    Dim excelApp As Excel.Application
    Dim workbook As Excel.Workbook
    Dim worksheet As Excel.Worksheet
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




        Catch ex As Exception

        End Try
    End Sub
End Class