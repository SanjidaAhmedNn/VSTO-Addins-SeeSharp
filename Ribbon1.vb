Imports Microsoft.Office.Tools.Ribbon
Imports System.ComponentModel.Design
Imports System.Security.Cryptography
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Excel
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Drawing
Imports System.ComponentModel
Imports System.Linq.Expressions
Public Class Ribbon1

    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As RibbonControlEventArgs) Handles Button1.Click
        'Dim form As New Form1
        'form.Show()
    End Sub

    Private Sub Button2_Click(sender As Object, e As RibbonControlEventArgs) Handles Button2.Click
        'Dim form As New Form2
        'form.Show()
    End Sub

    Private Sub Button3_Click(sender As Object, e As RibbonControlEventArgs) Handles Button3.Click
        'Dim form As New Form3
        'form.Show()
    End Sub

    Private Sub Button5_Click(sender As Object, e As RibbonControlEventArgs) Handles Button5.Click
        'Dim form As New Form8
        'form.Show()
    End Sub

    Private Sub Button6_Click(sender As Object, e As RibbonControlEventArgs) Handles Button6.Click
        'Dim form As New Form10

        'form.Show()
    End Sub

    Private Sub Button7_Click(sender As Object, e As RibbonControlEventArgs) Handles Button7.Click
        'Dim form As New Form7

        'form.Show()
    End Sub

    Private Sub Button8_Click(sender As Object, e As RibbonControlEventArgs) Handles Button8.Click
        Dim form As New Form11SwapRanges

        form.Show()
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

        'Dim form As New Form17DivideNames
        'form.Show()

    End Sub

    Private Sub Button16_Click_1(sender As Object, e As RibbonControlEventArgs) Handles Button16.Click
        'Dim form As New Form18_CombineRanges
        'form.Show()
    End Sub

    Private Sub Button19_Click(sender As Object, e As RibbonControlEventArgs) Handles Button19.Click
        'Dim form As New Form21FillEmtyCells
        'form.Show()
    End Sub

    Private Sub Button20_Click(sender As Object, e As RibbonControlEventArgs) Handles Button20.Click
        'Dim form As New Form22_Merge_Duplicate_Rows
        'form.Show()
    End Sub

    Private Sub Button21_Click(sender As Object, e As RibbonControlEventArgs) Handles Button21.Click
        'Dim form As New Form23_Merge_Duplicate_Columns
        'form.Show()
    End Sub

    Private Sub Button22_Click(sender As Object, e As RibbonControlEventArgs) Handles Button22.Click
        'Dim form As New Form24_Split_Cells
        'form.Show()
    End Sub

    Private Sub Button23_Click(sender As Object, e As RibbonControlEventArgs) Handles Button23.Click
        'Dim form As New Form25_Split_Range
        'form.Show()
    End Sub

    Private Sub Button24_Click(sender As Object, e As RibbonControlEventArgs) Handles Button24.Click
        'Dim form As New Form26_split_text_bycharacters
        'form.Show()
    End Sub

    Private Sub Button25_Click(sender As Object, e As RibbonControlEventArgs) Handles Button25.Click
        'Dim form As New Form27_Split_text_bystrings
        'form.Show()
    End Sub

    Private Sub Button26_Click(sender As Object, e As RibbonControlEventArgs) Handles Button26.Click
        'Dim form As New Form28_Split_text_bypattern
        'form.Show()
    End Sub

    Private Sub Button27_Click(sender As Object, e As RibbonControlEventArgs) Handles Button27.Click
        'Dim form As New Form29_Simple_Drop_down_List
        'form.Show()
    End Sub

    Private Sub Button28_Click(sender As Object, e As RibbonControlEventArgs) Handles Button28.Click
        'Dim form As New Form30_Create_Dynamic_Drop_down_List
        'form.Show()
    End Sub

    Private Sub Button31_Click(sender As Object, e As RibbonControlEventArgs) Handles Button31.Click

        'Unhide All

        Dim excelApp As Excel.Application
        Dim workbook As Excel.Workbook
        Dim worksheet As Excel.Worksheet


        excelApp = Globals.ThisAddIn.Application
        Workbook = excelApp.ActiveWorkbook
        worksheet = workbook.ActiveSheet

        ' Declare variables to store row and column numbers
        Dim data_Row_Num, data_Col_Num, last_Data_Row_Num, last_Data_Col_Num As Integer


        'Calculate the last used row and column number in the worksheet
        last_Data_Row_Num = worksheet.UsedRange.Rows.Count + worksheet.UsedRange.Row - 1
        last_Data_Col_Num = worksheet.UsedRange.Columns.Count + worksheet.UsedRange.Column - 1


        ' Loop through each row in the used range of the worksheet
        ' Check if the entire row is hidden
        ' If the row is hidden, unhide it
        For data_Row_Num = worksheet.UsedRange.Row To last_Data_Row_Num
            If worksheet.Range(worksheet.Cells(data_Row_Num, 1), worksheet.Cells(data_Row_Num, 3)).EntireRow.Hidden = True Then
                worksheet.Range(worksheet.Cells(data_Row_Num, 1), worksheet.Cells(data_Row_Num, 3)).EntireRow.Hidden = False
            End If
        Next


        ' Loop through each column in the used range of the worksheet
        ' Check if the entire column is hidden
        ' If the column is hidden, unhide it
        For data_Col_Num = worksheet.UsedRange.Column To last_Data_Col_Num
            If worksheet.Range(worksheet.Cells(1, data_Col_Num), worksheet.Cells(3, data_Col_Num)).EntireColumn.Hidden = True Then
                worksheet.Range(worksheet.Cells(1, data_Col_Num), worksheet.Cells(3, data_Col_Num)).EntireColumn.Hidden = False
            End If
        Next

    End Sub

    Private Sub Button32_Click(sender As Object, e As RibbonControlEventArgs) Handles Button32.Click

        'Unhide from Selected Range

        Dim excelApp As Excel.Application
        Dim workbook As Excel.Workbook
        Dim worksheet As Excel.Worksheet
        Dim selectedRange As Excel.Range

        excelApp = Globals.ThisAddIn.Application
        workbook = excelApp.ActiveWorkbook
        worksheet = workbook.ActiveSheet
        selectedRange = excelApp.Selection

        'Define varibales to store row and column numbers of the selected range
        Dim rowNum, colNum As Integer

        'Loop through each row of the selected range
        ' Check if the entire row is hidden
        ' If the row is hidden, unhide it
        For rowNum = 1 To selectedRange.Rows.Count
            If worksheet.Range(selectedRange.Cells(rowNum, 1), selectedRange.Cells(rowNum, 3)).EntireRow.Hidden = True Then
                worksheet.Range(selectedRange.Cells(rowNum, 1), selectedRange.Cells(rowNum, 3)).EntireRow.Hidden = False
            End If
        Next

        ' Loop through each column in the selected range
        ' Check if the entire column is hidden
        ' If the column is hidden, unhide it
        For colNum = 1 To selectedRange.Columns.Count
            If worksheet.Range(selectedRange.Cells(1, colNum), selectedRange.Cells(3, colNum)).EntireColumn.Hidden = True Then
                worksheet.Range(selectedRange.Cells(1, colNum), selectedRange.Cells(3, colNum)).EntireColumn.Hidden = False
            End If
        Next

    End Sub
End Class
