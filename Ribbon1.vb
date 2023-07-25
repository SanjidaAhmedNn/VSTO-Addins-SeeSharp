Imports Microsoft.Office.Tools.Ribbon
Imports Excel = Microsoft.Office.Interop.Excel

Public Class Ribbon1

    Dim excelApp As Excel.Application
    Dim workbook As Excel.Workbook
    Dim worksheet As Excel.Worksheet


    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub Flip_Click(sender As Object, e As RibbonControlEventArgs) Handles Button1.Click

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

    Private Sub Transform_Click(sender As Object, e As RibbonControlEventArgs) Handles Transform.Click

        Dim MyForm7 As New Form7

        excelApp = Globals.ThisAddIn.Application
        workbook = excelApp.ActiveWorkbook
        worksheet = workbook.ActiveSheet

        Dim selection As Excel.Range = CType(excelApp.Selection, Excel.Range)

        MyForm7.TextBox1.Text = selection.Address
        MyForm7.Show()

    End Sub

    Private Sub Transpose_Click(sender As Object, e As RibbonControlEventArgs) Handles Button3.Click


        Dim MyForm3 As New Form3

        excelApp = Globals.ThisAddIn.Application
        workbook = excelApp.ActiveWorkbook
        worksheet = workbook.ActiveSheet

        MyForm3.excelApp = excelApp
        MyForm3.workbook = workbook
        MyForm3.worksheet = worksheet
        MyForm3.workbook2 = workbook
        MyForm3.worksheet2 = worksheet

        MyForm3.FocusedTextBox = 0
        MyForm3.Form4Open = 0
        MyForm3.Workbook2Opened = False

        Dim selection As Excel.Range = CType(excelApp.Selection, Excel.Range)

        MyForm3.TextBox1.Text = selection.Address
        MyForm3.ComboBox1.SelectedIndex = -1
        MyForm3.ComboBox1.Text = "SOFTEKO"

        MyForm3.Show()

    End Sub

    Private Sub Unmerge_Click(sender As Object, e As RibbonControlEventArgs) Handles Button2.Click

        Dim MyForm10 As New Form10

        excelApp = Globals.ThisAddIn.Application
        workbook = excelApp.ActiveWorkbook
        worksheet = workbook.ActiveSheet

        Dim selection As Excel.Range = CType(excelApp.Selection, Excel.Range)

        MyForm10.TextBox1.Text = selection.Address
        MyForm10.Show()

    End Sub
End Class
