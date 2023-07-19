Imports Microsoft.Office.Tools.Ribbon

Public Class Ribbon1

    Dim excelApp As Excel.Application
    Dim workbook As Excel.Workbook
    Dim worksheet As Excel.Worksheet
    Public MyForm3 As New Form3

    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As RibbonControlEventArgs) Handles Button1.Click

        Dim MyForm1 As New Form1

        excelApp = Globals.ThisAddIn.Application
        workbook = excelApp.ActiveWorkbook
        worksheet = workbook.ActiveSheet

        Dim selection As Excel.Range = CType(excelApp.Selection, Excel.Range)

        MyForm1.TextBox1.Text = selection.Address
        MyForm1.ComboBox1.SelectedIndex = -1
        MyForm1.ComboBox1.Text = "SOFTEKO"

        MyForm1.Show()

    End Sub

    Private Sub Button2_Click(sender As Object, e As RibbonControlEventArgs) Handles Transform.Click

        Dim MyForm7 As New Form7

        excelApp = Globals.ThisAddIn.Application
        workbook = excelApp.ActiveWorkbook
        worksheet = workbook.ActiveSheet

        Dim selection As Excel.Range = CType(excelApp.Selection, Excel.Range)

        MyForm7.TextBox1.Text = selection.Address
        MyForm7.Show()

    End Sub

    Private Sub Button3_Click(sender As Object, e As RibbonControlEventArgs) Handles Button3.Click




        excelApp = Globals.ThisAddIn.Application
        workbook = excelApp.ActiveWorkbook
        worksheet = workbook.ActiveSheet

        Dim selection As Excel.Range = CType(excelApp.Selection, Excel.Range)

        MyForm3.TextBox1.Text = selection.Address
        MyForm3.ComboBox1.SelectedIndex = -1
        MyForm3.ComboBox1.Text = "SOFTEKO"

        MyForm3.Show()

    End Sub

End Class
