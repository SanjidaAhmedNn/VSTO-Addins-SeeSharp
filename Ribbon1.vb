Imports System.Drawing
Imports System.Windows.Forms
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

    Private Sub Merge_Click(sender As Object, e As RibbonControlEventArgs) Handles Button5.Click

        Dim MyForm8 As New Form8

        excelApp = Globals.ThisAddIn.Application
        workbook = excelApp.ActiveWorkbook
        worksheet = workbook.ActiveSheet

        Dim selection As Excel.Range = CType(excelApp.Selection, Excel.Range)

        MyForm8.TextBox1.Text = selection.Address
        MyForm8.Show()

    End Sub

    Private Sub Combine_Rows_Click(sender As Object, e As RibbonControlEventArgs) Handles Button6.Click

        Dim MyForm22 As New Form22_Merge_Duplicate_Rows

        MyForm22.excelApp = Globals.ThisAddIn.Application
        MyForm22.workBook = MyForm22.excelApp.ActiveWorkbook
        MyForm22.workSheet = MyForm22.workBook.ActiveSheet

        Dim selection As Excel.Range = CType(MyForm22.excelApp.Selection, Excel.Range)

        MyForm22.rng = selection

        MyForm22.TextBox1.Text = selection.Address

        Dim height As Single = MyForm22.Label3.Height
        Dim width As Single = MyForm22.Label2.Width - 3.5

        Dim i As Integer

        For i = 1 To selection.Columns.Count

            Dim label As New System.Windows.Forms.Label
            label.Text = selection.Cells(1, i).Value
            label.Location = New System.Drawing.Point(0, (i - 1) * height)
            label.Height = height
            label.Width = width
            label.Font = New Font(label.Font.FontFamily, 9)
            label.BorderStyle = BorderStyle.FixedSingle
            label.TextAlign = ContentAlignment.MiddleCenter
            label.BorderStyle = BorderStyle.None
            label.TextAlign = ContentAlignment.MiddleLeft
            MyForm22.CustomGroupBox7.Controls.Add(label)

            Dim label2 As New System.Windows.Forms.Label
            label2.Text = selection.Cells(2, i).Value
            label2.Location = New System.Drawing.Point(width, (i - 1) * height)
            label2.Height = height
            label2.Width = width
            label2.Font = New Font(label2.Font.FontFamily, 9)
            label2.BorderStyle = BorderStyle.FixedSingle
            label2.TextAlign = ContentAlignment.MiddleCenter
            label2.BorderStyle = BorderStyle.None
            label2.TextAlign = ContentAlignment.MiddleLeft
            MyForm22.CustomGroupBox7.Controls.Add(label2)

        Next

        MyForm22.Show()

    End Sub

End Class
