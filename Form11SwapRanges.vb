Public Class Form11SwapRanges

    Dim excelApp As Excel.Application
    Dim workbook As Excel.Workbook
    Dim inputWs As Excel.Worksheet
    Dim outWs As Excel.Worksheet
    Dim inRng As Excel.Range
    Dim outRng As Excel.Range





    Private Sub PictureBox9_Click(sender As Object, e As EventArgs) Handles PictureBox9.Click

    End Sub

    Private Sub Form11SwapRanges_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        Try

            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            inputWs = workbook.ActiveSheet


            inRng = inputWs.Range(TextBox1.Text)





        Catch ex As Exception

        End Try
    End Sub
    Private Sub Display()




    End Sub

End Class