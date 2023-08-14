Imports Microsoft.Office.Tools.Ribbon
Imports System.Drawing
Imports System.Windows.Forms
'Imports Microsoft.Office.Tools.Ribbon
Imports Excel = Microsoft.Office.Interop.Excel
'Imports Microsoft.Office.Tools.Ribbon
Public Class Ribbon1

    Dim WithEvents excelApp As Excel.Application
    Dim workBook As Excel.Workbook
    Dim workSheet As Excel.Worksheet

    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As RibbonControlEventArgs) Handles Button1.Click
        Dim MyForm1 As New Form1

        excelApp = Globals.ThisAddIn.Application
        workBook = excelApp.ActiveWorkbook
        workSheet = workBook.ActiveSheet

        Dim selection As Excel.Range = CType(excelApp.Selection, Excel.Range)

        MyForm1.TextBox1.Text = selection.Address
        MyForm1.ComboBox1.SelectedIndex = -1
        MyForm1.ComboBox1.Text = "SOFTEKO"

        MyForm1.Show()
    End Sub

    Private Sub Button2_Click(sender As Object, e As RibbonControlEventArgs)
        Dim form As New Form2
        form.Show()
    End Sub

    Private Sub Button3_Click(sender As Object, e As RibbonControlEventArgs) Handles Button3.Click
        Dim MyForm3 As New Form3

        excelApp = Globals.ThisAddIn.Application
        workBook = excelApp.ActiveWorkbook
        workSheet = workBook.ActiveSheet

        MyForm3.excelApp = excelApp
        MyForm3.workbook = workBook
        MyForm3.worksheet = workSheet
        MyForm3.workbook2 = workBook
        MyForm3.worksheet2 = workSheet

        MyForm3.FocusedTextBox = 0
        MyForm3.Form4Open = 0
        MyForm3.Workbook2Opened = False

        Dim selection As Excel.Range = CType(excelApp.Selection, Excel.Range)

        MyForm3.TextBox1.Text = selection.Address
        MyForm3.ComboBox1.SelectedIndex = -1
        MyForm3.ComboBox1.Text = "SOFTEKO"

        MyForm3.Show()
    End Sub

    Private Sub Button5_Click(sender As Object, e As RibbonControlEventArgs) Handles Button5.Click
        Dim MyForm8 As New Form8

        excelApp = Globals.ThisAddIn.Application
        workBook = excelApp.ActiveWorkbook
        workSheet = workBook.ActiveSheet

        Dim selection As Excel.Range = CType(excelApp.Selection, Excel.Range)

        MyForm8.TextBox1.Text = selection.Address
        MyForm8.Show()
    End Sub

    Private Sub Button6_Click(sender As Object, e As RibbonControlEventArgs) Handles Button6.Click
        Dim MyForm10 As New Form10

        excelApp = Globals.ThisAddIn.Application
        workBook = excelApp.ActiveWorkbook
        workSheet = workBook.ActiveSheet

        Dim selection As Excel.Range = CType(excelApp.Selection, Excel.Range)

        MyForm10.TextBox1.Text = selection.Address
        MyForm10.Show()
    End Sub

    Private Sub Button7_Click(sender As Object, e As RibbonControlEventArgs) Handles Button7.Click
        Dim MyForm7 As New Form7

        excelApp = Globals.ThisAddIn.Application
        workBook = excelApp.ActiveWorkbook
        workSheet = workBook.ActiveSheet

        Dim selection As Excel.Range = CType(excelApp.Selection, Excel.Range)

        MyForm7.TextBox1.Text = selection.Address
        MyForm7.Show()
    End Sub

    Private Sub Button8_Click(sender As Object, e As RibbonControlEventArgs) Handles Button8.Click
        Dim MyForm11 As New Form11SwapRanges

        excelApp = Globals.ThisAddIn.Application
        workBook = excelApp.ActiveWorkbook
        workSheet = workBook.ActiveSheet

        Dim selection As Excel.Range = CType(excelApp.Selection, Excel.Range)

        MyForm11.Show()
    End Sub

    Private Sub Button4_Click(sender As Object, e As RibbonControlEventArgs) Handles Button4.Click
        Dim form As New Form12HideRanges

        form.Show()
    End Sub

    Private Sub Button11_Click(sender As Object, e As RibbonControlEventArgs) Handles Button11.Click
        Dim form As New Form13HideAllExceptSelectedRange

        form.Show()
    End Sub

    Private Sub Button12_Click(sender As Object, e As RibbonControlEventArgs) Handles Button12.Click
        Dim form As New Form14SpecifyScrollArea
        form.Show()
    End Sub

    Private Sub Button13_Click(sender As Object, e As RibbonControlEventArgs) Handles Button13.Click

        Dim form As New Form15CompareCells
        form.Show()

    End Sub

    Private Sub Button14_Click(sender As Object, e As RibbonControlEventArgs) Handles Button14.Click

        Dim form As New Form16PasteintoVisibleRange
        form.Show()

    End Sub

    Private Sub Button15_Click(sender As Object, e As RibbonControlEventArgs) Handles Button15.Click

        Dim form As New Form17DivideNames
        form.Show()

    End Sub

    Private Sub Button16_Click_1(sender As Object, e As RibbonControlEventArgs) Handles Button16.Click
        Dim form As New Form18_CombineRanges
        form.Show()
    End Sub

    Private Sub Button19_Click(sender As Object, e As RibbonControlEventArgs) Handles Button19.Click
        Dim form As New Form21FillEmtyCells
        form.Show()
    End Sub

    Private Sub Button20_Click(sender As Object, e As RibbonControlEventArgs) Handles Button20.Click
        Dim MyForm22 As New Form22_Merge_Duplicate_Rows

        excelApp = Globals.ThisAddIn.Application
        workBook = excelApp.ActiveWorkbook
        workSheet = workBook.ActiveSheet

        Dim selection As Excel.Range = CType(excelApp.Selection, Excel.Range)

        MyForm22.TextBox1.Text = selection.Address

        MyForm22.Show()
    End Sub

    Private Sub Button21_Click(sender As Object, e As RibbonControlEventArgs) Handles Button21.Click
        Dim form As New Form23_Merge_Duplicate_Columns
        form.Show()
    End Sub

    Private Sub Button22_Click(sender As Object, e As RibbonControlEventArgs) Handles Button22.Click
        Dim form As New Form24_Split_Cells
        form.Show()
    End Sub

    Private Sub Button23_Click(sender As Object, e As RibbonControlEventArgs) Handles Button23.Click
        Dim form As New Form25_Split_Range
        form.Show()
    End Sub

    Private Sub Button24_Click(sender As Object, e As RibbonControlEventArgs) Handles Button24.Click
        Dim form As New Form26_split_text_bycharacters
        form.Show()
    End Sub

    Private Sub Button25_Click(sender As Object, e As RibbonControlEventArgs) Handles Button25.Click
        Dim form As New Form27_Split_text_bystrings
        form.Show()
    End Sub

    Private Sub Button26_Click(sender As Object, e As RibbonControlEventArgs) Handles Button26.Click
        Dim form As New Form28_Split_text_bypattern
        form.Show()
    End Sub

    Private Sub Button27_Click(sender As Object, e As RibbonControlEventArgs) Handles Button27.Click
        Dim form As New Form29_Simple_Drop_down_List
        form.Show()
    End Sub

    Private Sub Button28_Click(sender As Object, e As RibbonControlEventArgs) Handles Button28.Click
        Dim form As New Form30_Create_Dynamic_Drop_down_List
        form.Show()
    End Sub

    Private Sub Button29_Click(sender As Object, e As RibbonControlEventArgs) Handles Button29.Click
        Dim form As New Form31_Update_Dynamic_Drop_down_List
        form.Show()
    End Sub
End Class
