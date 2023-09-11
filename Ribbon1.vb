Imports Microsoft.Office.Tools.Ribbon
Imports System.ComponentModel.Design
Imports System.Security.Cryptography
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Excel
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Drawing
Imports System.ComponentModel
Imports System.Linq.Expressions
Imports System.Data.Common

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

    ''' <summary>
    ''' Unhides any hidden rows and columns from the entire sheet.
    ''' </summary>

    Private Sub Button31_Click(sender As Object, e As RibbonControlEventArgs) Handles Button31.Click

        Try

            Dim excelApp As Excel.Application
            Dim workbook As Excel.Workbook
            Dim worksheet As Excel.Worksheet


            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
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

        Catch ex As Exception

        End Try


    End Sub

    ''' <summary>
    ''' Unhides any hidden rows and columns from a Selected Range.
    ''' </summary>

    Private Sub Button32_Click(sender As Object, e As RibbonControlEventArgs) Handles Button32.Click

        Try

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

        Catch ex As Exception

        End Try


    End Sub

    ''' <summary>
    ''' Checks if a specific column in the active worksheet of the current workbook is empty.
    ''' </summary>
    ''' <param name="columnIndex"> The index of the column to check, starting from 1 </param>
    ''' <returns> True if the column is empty, otherwise False. </returns>

    Private Function IsColumnEmpty(columnIndex As Integer) As Boolean
        ' Assume the column is empty by default
        Dim isEmpty As Boolean = True

        Dim excelApp As Excel.Application
        Dim workbook As Excel.Workbook
        Dim worksheet As Excel.Worksheet


        excelApp = Globals.ThisAddIn.Application
        workbook = excelApp.ActiveWorkbook
        worksheet = workbook.ActiveSheet


        Dim lastCell() As String
        Dim lastRowNum, lastColNum As Integer

        lastCell = worksheet.UsedRange.Address.Split(":"c)

        lastRowNum = worksheet.Range(lastCell(1)).Row
        lastColNum = worksheet.Range(lastCell(1)).Column


        ' Extract the entire column specified by the columnIndex from the used range
        Dim range As Excel.Range = worksheet.Range(worksheet.Cells(1, 1), worksheet.Cells(lastRowNum, lastColNum))
        Dim column As Excel.Range = range.Columns(columnIndex)

        ' Check if there are any non-blank cells in the column using WorksheetFunction.CountA
        ' If the count is not zero, the column isn't empty
        If excelApp.WorksheetFunction.CountA(column) <> 0 Then
            isEmpty = False
        End If

        ' Return the result as boolean
        Return isEmpty
    End Function

    ''' <summary>
    ''' Removes blank columns from selected range
    ''' </summary>
    Private Sub Button33_Click(sender As Object, e As RibbonControlEventArgs) Handles Button33.Click

        Try

            Dim excelApp As Excel.Application
            Dim workbook As Excel.Workbook
            Dim worksheet As Excel.Worksheet
            Dim selectedRng As Excel.Range
            Dim blankColumnList As String = ""
            Dim blankColCount As Integer = 0
            Dim flag As String = "Empty"


            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            worksheet = workbook.ActiveSheet
            selectedRng = excelApp.Selection

            '"rngCount" variable indicates the number of ranges in the users' selection.
            '0 means a single continuous selection
            '> 0 means multiple disjoint range selection
            Dim rngCount As Integer
            rngCount = 0

            For Each c As Char In selectedRng.Address

                If c = "," Then
                    rngCount = rngCount + 1
                End If

            Next

            'if user select a single continuous range "rngCount" will be 0
            If rngCount = 0 Then

                'if user select only a single cell the following warning will pop up and exit from the code 
                If selectedRng.Rows.Count = 1 And selectedRng.Columns.Count = 1 Then
                    MsgBox("This Add-in doesn't work for single cell. Please Select a Range and try again!", MsgBoxStyle.Exclamation, "Warning")
                    Exit Sub
                End If

                'loop through each cells of a column of the selection and check if the column is empty or not.
                'if all the cells of the column of that seleted range are blank then the "flag" variable remains Empty. If any of the cell of that column is non-empty then "flag" will be "NotEmpty"
                'if flag is "Empty" then store the address of that column in the "blankColumnList" string and increase the value of "blankColCount" by 1.
                For i = 1 To selectedRng.Columns.Count
                    flag = "Empty"
                    For j = 1 To selectedRng.Rows.Count
                        If Not selectedRng.Cells(j, i).value Is Nothing Then

                            flag = "NotEmpty"

                        End If

                    Next


                    If flag = "Empty" Then

                        blankColumnList = blankColumnList & "," & worksheet.Range(selectedRng.Cells(1, i), selectedRng.Cells(selectedRng.Rows.Count, i)).Address
                        blankColCount = blankColCount + 1

                    End If

                Next

                'remove the leading comma (,) from the "blankColumnList" string
                blankColumnList = Right(blankColumnList, Len(blankColumnList) - 1)

                'removes the empty columns
                worksheet.Range(blankColumnList).Delete(Excel.XlDeleteShiftDirection.xlShiftToLeft)

                selectedRng.Cells(1, 1).select

                'displays a msgbox that shows how many columns are deleted
                MsgBox(blankColCount & " Column(s) are deleted.", MsgBoxStyle.Information, "SOFTEKO")





                'user selected multiple disjoint ranges
            Else
                'an array named "arrRng" is used to separately store all  the addresses of the selection 
                Dim arrRng As String() = Split(selectedRng.Address, ",")

                'loop through each address from the selection and check if any range is a single cell or not. If so, then the following warning will pop up.
                'Then exit sub.
                For i = 0 To UBound(arrRng)
                    selectedRng = worksheet.Range(arrRng(i))
                    If selectedRng.Rows.Count = 1 And selectedRng.Columns.Count = 1 Then
                        MsgBox("This Add-in doesn't work for single cell. Please select a Range and try again!", MsgBoxStyle.Exclamation, "Warning")
                        Exit Sub
                    End If
                Next


                'loop through each range of the selection and remove blank columns
                For i = 0 To UBound(arrRng)

                    selectedRng = worksheet.Range(arrRng(i))

                    'loop through each cells of a column of the selection and check if the column is empty or not.
                    'if all the cells of the column of that seleted range are blank then the "flag" variable remains Empty. If any of the cell of that column is non-empty then "flag" will be "NotEmpty"
                    'if flag is "Empty" then store the address of that column in the "blankColumnList" string and increase the value of "blankColCount" by 1.
                    For k = 1 To selectedRng.Columns.Count
                        flag = "Empty"
                        For j = 1 To selectedRng.Rows.Count
                            If Not selectedRng.Cells(j, k).value Is Nothing Then

                                flag = "NotEmpty"

                            End If

                        Next

                        If flag = "Empty" Then

                            blankColumnList = blankColumnList & "," & worksheet.Range(selectedRng.Cells(1, k), selectedRng.Cells(selectedRng.Rows.Count, k)).Address
                            blankColCount = blankColCount + 1

                        End If

                    Next

                Next


                'remove the leading comma (,) from the "blankColumnList" string
                blankColumnList = Right(blankColumnList, Len(blankColumnList) - 1)

                'removes the empty columns
                worksheet.Range(blankColumnList).Delete(Excel.XlDeleteShiftDirection.xlShiftToLeft)

                selectedRng.Cells(1, 1).select

                'displays a msgbox that shows how many columns are deleted
                MsgBox(blankColCount & " Column(s) are deleted.", MsgBoxStyle.Information, "SOFTEKO")

            End If






            'Dim arrBlankCol As String() = Split(blankColumnList, ",")
            'For i = 0 To UBound(arrBlankCol)

            '    worksheet.Range(arrBlankCol(i)).Delete(Excel.XlDeleteShiftDirection.xlShiftToLeft)

            'Next



            'loop through each each columns of the selection to check is the column is empty using the "IsColumnEmpty" function
            'If the column is empty then store the address of a cell of that column in the "blankColumnList" string
            'For k = 1 To selectedRng.Columns.Count

            '    If columnFlag = "Empty" Then

            '        blankColumnList = blankColumnList & "," & selectedRng.Cells(1, k).address
            '        blankColCount = blankColCount + 1
            '    End If

            'Next



            ''delete the blank columns
            'worksheet.Range(blankColumnList).EntireColumn.Delete()




        Catch ex As Exception

        End Try


    End Sub

    ''' <summary>
    ''' removes blank columns from Active Sheet
    ''' </summary>
    Private Sub Button34_Click(sender As Object, e As RibbonControlEventArgs) Handles Button34.Click
        Try

            Dim excelApp As Excel.Application
            Dim workbook As Excel.Workbook
            Dim worksheet As Excel.Worksheet
            Dim selectedRng As Excel.Range
            Dim blankColumnList As String = ""
            Dim blankColCount As Integer = 0


            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            worksheet = workbook.ActiveSheet
            selectedRng = excelApp.Selection

            'use the UsedRange method to find the address of the range used in the active sheet
            'use split function to get 2nd portion of the range which is the last cell of the used range
            'Use this addrees to to find column number of last cell
            Dim lastCell() As String
            Dim lastColNum As Integer

            lastCell = worksheet.UsedRange.Address.Split(":"c)
            lastColNum = worksheet.Range(lastCell(1)).Column

            'loop through each columns of the active sheet upto the last column number
            'check if the entire column is empty or not by using the "IsEmptyColumn" function
            For i = 1 To lastColNum
                If IsColumnEmpty(i) = True Then
                    blankColumnList = blankColumnList & "," & worksheet.Range(worksheet.Cells(1, i), worksheet.Cells(2, i)).Address
                    blankColCount = blankColCount + 1
                End If
            Next

            'remove the leading comma (,) from the "blankColumnList" string
            blankColumnList = Right(blankColumnList, Len(blankColumnList) - 1)

            'removes the empty columns
            worksheet.Range(blankColumnList).EntireColumn.Delete()
            worksheet.Cells(1, 1).select

            'displays a msgbox that shows how many columns are deleted
            MsgBox(blankColCount & " Column(s) are deleted.", MsgBoxStyle.Information, "SOFTEKO")

        Catch ex As Exception

        End Try


    End Sub

    Private Sub Button35_Click(sender As Object, e As RibbonControlEventArgs) Handles Button35.Click

        Try


            Dim excelApp As Excel.Application
            Dim workbook As Excel.Workbook
            Dim worksheet As Excel.Worksheet
            Dim blankColumnList As String
            Dim confirmationMsg As String = ""
            Dim blankColCount As Integer
            Dim i As Integer = 0


            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook


            Dim selectedSheets As Excel.Sheets = excelApp.ActiveWindow.SelectedSheets
            Dim sheetName As String = ""

            For Each sheet As Excel.Worksheet In selectedSheets
                sheetName = sheetName & "," & sheet.Name
            Next
            sheetName = Right(sheetName, Len(sheetName) - 1)

            Dim arrSheetName As String() = Split(sheetName, ",")

            For i = 0 To UBound(arrSheetName)
                blankColumnList = ""
                blankColCount = 0
                worksheet = workbook.Sheets(arrSheetName(i))
                worksheet.Activate()


                'use the UsedRange method to find the address of the range used in the active sheet
                'use split function to get 2nd portion of the range which is the last cell of the used range
                'Use this addrees to to find column number of last cell
                Dim lastCell() As String
                Dim lastColNum As Integer

                lastCell = worksheet.UsedRange.Address.Split(":"c)
                lastColNum = worksheet.Range(lastCell(1)).Column

                'loop through each columns of the active sheet upto the last column number
                'check if the entire column is empty or not by using the "IsEmptyColumn" function
                For j = 1 To lastColNum
                    If IsColumnEmpty(j) = True Then
                        blankColumnList = blankColumnList & "," & worksheet.Range(worksheet.Cells(1, j), worksheet.Cells(2, j)).Address
                        blankColCount = blankColCount + 1
                    End If
                Next

                'remove the leading comma (,) from the "blankColumnList" string
                blankColumnList = Right(blankColumnList, Len(blankColumnList) - 1)

                'removes the empty columns
                worksheet.Range(blankColumnList).EntireColumn.Delete()

                'stores information about how many columns deleted from which sheet
                confirmationMsg = confirmationMsg & blankColCount & " Column(s) are deleted from " & arrSheetName(i) & vbCrLf

            Next

            'finally this msgBox is shown
            MsgBox(confirmationMsg, MsgBoxStyle.Information, "SOFTEKO")

        Catch ex As Exception

        End Try



    End Sub

    Private Sub Button36_Click(sender As Object, e As RibbonControlEventArgs) Handles Button36.Click



        Try


            Dim excelApp As Excel.Application
            Dim workbook As Excel.Workbook
            Dim worksheet As Excel.Worksheet
            Dim blankColumnList As String
            Dim confirmationMsg As String = ""
            Dim blankColCount As Integer
            Dim i As Integer = 0


            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook


            Dim selectedSheets As Excel.Sheets = excelApp.Sheets
            Dim sheetName As String = ""

            For Each sheet As Excel.Worksheet In selectedSheets
                sheetName = sheetName & "," & sheet.Name
            Next
            sheetName = Right(sheetName, Len(sheetName) - 1)

            Dim arrSheetName As String() = Split(sheetName, ",")

            For i = 0 To UBound(arrSheetName)
                blankColumnList = ""
                blankColCount = 0
                worksheet = workbook.Sheets(arrSheetName(i))
                worksheet.Activate()


                'use the UsedRange method to find the address of the range used in the active sheet
                'use split function to get 2nd portion of the range which is the last cell of the used range
                'Use this addrees to to find column number of last cell
                Dim lastCell() As String
                Dim lastColNum As Integer

                lastCell = worksheet.UsedRange.Address.Split(":"c)
                lastColNum = worksheet.Range(lastCell(1)).Column

                'loop through each columns of the active sheet upto the last column number
                'check if the entire column is empty or not by using the "IsEmptyColumn" function
                For j = 1 To lastColNum
                    If IsColumnEmpty(j) = True Then
                        blankColumnList = blankColumnList & "," & worksheet.Range(worksheet.Cells(1, j), worksheet.Cells(2, j)).Address
                        blankColCount = blankColCount + 1
                    End If
                Next

                'remove the leading comma (,) from the "blankColumnList" string
                blankColumnList = Right(blankColumnList, Len(blankColumnList) - 1)

                'removes the empty columns
                worksheet.Range(blankColumnList).EntireColumn.Delete()

                'stores information about how many columns deleted from which sheet
                confirmationMsg = confirmationMsg & blankColCount & " Column(s) are deleted from " & arrSheetName(i) & vbCrLf

            Next

            'finally this msgBox is shown
            MsgBox(confirmationMsg, MsgBoxStyle.Information, "SOFTEKO")

        Catch ex As Exception

        End Try


    End Sub
End Class
