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
Imports System.Runtime.InteropServices
Imports Application = System.Windows.Forms.Application



Public Class Ribbon1

    Dim excelApp As Excel.Application
    Dim workbook As Excel.Workbook
    Dim worksheet As Excel.Worksheet



    'Public Function ConvertImage(ByVal image As Image) As stdole.IPictureDisp
    '    Return DirectCast(AxHost.GetIPictureDispFromPicture(image), stdole.IPictureDisp)
    'End Function


    Private Function SplitText(Source, Pattern, Consecutive, KeepSeparator, Before)

        Dim SplitValues(0) As String
        Dim Index As Integer = -1
        Dim Start As Integer = 1

        For i = 1 To Len(Pattern)
            If Mid(Pattern, i, 1) <> "*" Then
                Dim SeparatorLength As Integer = 1
                While Mid(Pattern, i + SeparatorLength, 1) <> "*"
                    SeparatorLength = SeparatorLength + 1
                End While
                Dim separator As String = Mid(Pattern, i, SeparatorLength)
                Dim Ending As Integer = InStr(Source, separator)
                MsgBox(Ending)
                Index = Index + 1
                ReDim Preserve SplitValues(Index)
                SplitValues(Index) = Mid(Source, Start, Ending - Start)
                Start = Ending + Len(separator)
            End If
        Next

        SplitText = SplitValues

    End Function

    Function IsRangeEmpty(rng As Excel.Range) As Boolean

        Dim Result As Boolean = True

        For Each cell In rng
            If cell.Value IsNot Nothing AndAlso cell.Value.ToString() <> String.Empty Then
                Result = False
                Exit For
            End If
        Next

        Return Result

    End Function


    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As RibbonControlEventArgs) Handles Button1.Click
        'Dim form As New Form1
        'form.Show()

        Dim MyForm1 As New Form1
        If form_flag = False Then

            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            worksheet = workbook.ActiveSheet

            Dim selection As Excel.Range = CType(excelApp.Selection, Excel.Range)

            If IsRangeEmpty(selection) = True Then
                MessageBox.Show("You have not selected any data.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)

            Else
                MyForm1.TextBox1.Text = selection.Address
                MyForm1.ComboBox1.SelectedIndex = -1
                MyForm1.ComboBox1.Text = "SOFTEKO"

                MyForm1.Show()
                form_flag = True
            End If
        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As RibbonControlEventArgs)
        'Dim form As New Form2
        'form.Show()
    End Sub

    Private Sub Button3_Click(sender As Object, e As RibbonControlEventArgs) Handles Button3.Click
        'Dim form As New Form3
        'form.Show()
        If form_flag = False Then   'For avoiding multiple occurrence
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

            If IsRangeEmpty(selection) = True Then
                MessageBox.Show("You have not selected any data.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)

            Else
                MyForm3.TextBox1.Text = selection.Address
                MyForm3.ComboBox1.SelectedIndex = -1
                MyForm3.ComboBox1.Text = "SOFTEKO"

                MyForm3.Show()
                form_flag = True
            End If
        End If

    End Sub

    Private Sub Button5_Click(sender As Object, e As RibbonControlEventArgs) Handles Button5.Click
        'Dim form As New Form8
        'form.Show()
        If form_flag = False Then
            Dim MyForm8 As New Form8

            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            worksheet = workbook.ActiveSheet

            Dim selection As Excel.Range = CType(excelApp.Selection, Excel.Range)

            If IsRangeEmpty(selection) = True Then
                MessageBox.Show("You have not selected any data.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)

            Else
                MyForm8.TextBox1.Text = selection.Address
                MyForm8.ComboBox1.SelectedIndex = -1
                MyForm8.ComboBox1.Text = "SOFTEKO"
                MyForm8.Show()
                form_flag = True
            End If
        End If

    End Sub

    Private Sub Button6_Click(sender As Object, e As RibbonControlEventArgs) Handles Button6.Click
        'Dim form As New Form10

        'form.Show()
        If form_flag = False Then
            Dim MyForm10 As New Form10

            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            worksheet = workbook.ActiveSheet

            Dim selection As Excel.Range = CType(excelApp.Selection, Excel.Range)

            MyForm10.TextBox1.Text = selection.Address
            MyForm10.ComboBox1.SelectedIndex = -1
            MyForm10.ComboBox1.Text = "SOFTEKO"
            MyForm10.Show()
            form_flag = True
        End If
    End Sub

    Private Sub Button7_Click(sender As Object, e As RibbonControlEventArgs) Handles Button7.Click
        'Dim form As New Form7

        'form.Show()
        If form_flag = False Then
            Dim MyForm7 As New Form7

            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            worksheet = workbook.ActiveSheet

            Dim selection As Excel.Range = CType(excelApp.Selection, Excel.Range)

            If IsRangeEmpty(selection) = True Then
                MessageBox.Show("You have not selected any data.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)

            Else
                MyForm7.TextBox1.Text = selection.Address
                MyForm7.Show()
                form_flag = True
            End If

        End If
    End Sub

    Private Sub Button8_Click(sender As Object, e As RibbonControlEventArgs) Handles Button8.Click
        If form_flag = False Then
            Dim form As New Form11SwapRanges

            form.Show()
            form_flag = True
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As RibbonControlEventArgs)

    End Sub

    Private Sub Button11_Click(sender As Object, e As RibbonControlEventArgs)

    End Sub

    Private Sub Button12_Click(sender As Object, e As RibbonControlEventArgs)
        Dim form As New Form14SpecifyScrollArea
        form.Show()
    End Sub

    Private Sub Button13_Click(sender As Object, e As RibbonControlEventArgs) Handles Button13.Click
        'MsgBox(form_flag)
        If form_flag = False Then
            Dim form As New Form15CompareCells
            form.Show()
            form_flag = True
        End If

    End Sub

    Private Sub Button14_Click(sender As Object, e As RibbonControlEventArgs) Handles Button14.Click
        If form_flag = False Then
            Dim form As New Form16PasteintoVisibleRange
            form.Show()
            form_flag = True
        End If

    End Sub

    Private Sub Button15_Click(sender As Object, e As RibbonControlEventArgs) Handles Button15.Click
        If form_flag = False Then
            Dim form As New Form17DivideNames
            form.Show()
            form_flag = True
        End If

    End Sub

    Private Sub Button16_Click_1(sender As Object, e As RibbonControlEventArgs)
        'Dim form As New Form18_CombineRanges
        'form.Show()
    End Sub

    Private Sub Button19_Click(sender As Object, e As RibbonControlEventArgs) Handles Button19.Click
        If form_flag = False Then
            Dim form As New Form21FillEmtyCells
            form.Show()
            form_flag = True
        End If
    End Sub

    Private Sub Button20_Click(sender As Object, e As RibbonControlEventArgs)
        'If form_flag = False
        'Dim form As New Form22_Merge_Duplicate_Rows
        'form.Show()
    End Sub

    Private Sub Button21_Click(sender As Object, e As RibbonControlEventArgs)
        'Dim form As New Form23_Merge_Duplicate_Columns
        'form.Show()
    End Sub

    Private Sub Button22_Click(sender As Object, e As RibbonControlEventArgs)
        'Dim form As New Form24_Split_Cells
        'form.Show()
    End Sub

    Private Sub Button23_Click(sender As Object, e As RibbonControlEventArgs)
        'Dim form As New Form25_Split_Range
        'form.Show()
    End Sub

    Private Sub Button24_Click(sender As Object, e As RibbonControlEventArgs)
        'Dim form As New Form26_split_text_bycharacters
        'form.Show()
    End Sub

    Private Sub Button25_Click(sender As Object, e As RibbonControlEventArgs)
        'Dim form As New Form27_Split_text_bystrings
        'form.Show()
    End Sub

    Private Sub Button26_Click(sender As Object, e As RibbonControlEventArgs)
        'Dim form As New Form28_Split_text_bypattern
        'form.Show()
    End Sub

    Private Sub Button27_Click(sender As Object, e As RibbonControlEventArgs)
        'Dim form As New Form29_Simple_Drop_down_List
        'form.Show()
    End Sub

    Private Sub Button28_Click(sender As Object, e As RibbonControlEventArgs) Handles Button28.Click
        If form_flag = False Then
            Dim form As New Form30_Create_Dynamic_Drop_down_List

            form.Show()
            form_flag = True
        End If
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

            'takes all the ranges selected by user into an array named arrRng
            Dim arrRng As String() = Split(selectedRange.Address, ",")

            'loops through each range selected by user, which is stored in arrRng array
            For p = 0 To UBound(arrRng)

                selectedRange = worksheet.Range(arrRng(p))

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
            Next

        Catch ex As Exception

        End Try


    End Sub

    ''' <summary>
    ''' Removes blank columns from selected range
    ''' </summary>
    Private Sub Button33_Click(sender As Object, e As RibbonControlEventArgs) Handles Button33.Click

        Try

            Dim excelApp As Excel.Application
            Dim workbook As Excel.Workbook
            Dim worksheet As Excel.Worksheet
            Dim selectedRng As Excel.Range
            Dim blankColCount As Integer = 0
            Dim flag As String = "Empty"
            Dim ValueFlag As String = "Empty"


            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            worksheet = workbook.ActiveSheet
            selectedRng = excelApp.Selection

            '"rngCount" variable indicates the number of ranges in the users' selection.
            '0 means a single continuous selection
            '> 0 means user selectd multiple disjoint ranges
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


                'loops through the entire selection and if the entire selection is empty or not
                'if the entire selection is blank then valueFlag will become "NotEmpty"  
                For i = 1 To selectedRng.Rows.Count
                    For j = 1 To selectedRng.Columns.Count
                        If Not selectedRng.Cells(i, j).value Is Nothing Then
                            ValueFlag = "NotEmpty"
                        End If
                    Next
                Next

                'loop through each cells of a column of the selection and check if the column is empty or not.
                'if all the cells of the column of that seleted range are blank then the "flag" variable remains Empty. If any of the cell of that column is non-empty then "flag" will be "NotEmpty"
                'if flag is "Empty" then the respective cells of that column of the selection will be deleted.
                'Note that any cells of the same column that is outside the selection will be deleted even if it is empty
                'after checking a column, the "flag" variable resets to "Empty"
                For i = selectedRng.Columns.Count To 1 Step -1
                    flag = "Empty"
                    For j = selectedRng.Rows.Count To 1 Step -1
                        If Not selectedRng.Cells(j, i).value Is Nothing Then

                            flag = "NotEmpty"

                        End If

                    Next


                    If flag = "Empty" Then

                        worksheet.Range(selectedRng.Cells(1, i), selectedRng.Cells(selectedRng.Rows.Count, i)).Delete(Excel.XlDeleteShiftDirection.xlShiftToLeft)
                        blankColCount += 1

                    End If

                Next

                'if no blank columns are found in a sheet then go to the "nextloop" section and skip the lines in between
                If blankColCount = 0 Then
                    GoTo break1
                End If

                'valueFlag is "Empty" means the entire selection is blank
                'so the msgbox will be shown  and then exit sub                
                If ValueFlag = "Empty" Then
                    MsgBox(blankColCount & " Column(s) are deleted.", MsgBoxStyle.Information, "SOFTEKO")
                    Exit Sub
                End If

                selectedRng.Cells(1, 1).select
break1:
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


                'loops through the entire selection and if the entire selection is empty or not
                'if the entire selection is blank then valueFlag will become "NotEmpty"  
                For i = 0 To UBound(arrRng)
                    For j = 1 To selectedRng.Rows.Count
                        For k = 1 To selectedRng.Columns.Count
                            If Not selectedRng.Cells(j, k).value Is Nothing Then
                                ValueFlag = "NotEmpty"
                            End If
                        Next
                    Next
                Next


                'loop through each range of the selection and remove blank columns
                For i = 0 To UBound(arrRng)

                    selectedRng = worksheet.Range(arrRng(i))

                    'loop through each cells of a column of the selection and check if the column is empty or not.
                    'if all the cells of the column of that seleted range are blank then the "flag" variable remains "Empty". If any of the cell of that column is non-empty then "flag" will be "NotEmpty"
                    'if flag is "Empty" then the respective cells of that column of the selection will be deleted.
                    'Note that any cells of the same column that is outside the selection will be deleted even if it is empty
                    'after checking a column, the "flag" variable resets to "Empty"
                    For k = selectedRng.Columns.Count To 1 Step -1
                        flag = "Empty"

                        For j = selectedRng.Rows.Count To 1 Step -1

                            If Not selectedRng.Cells(j, k).value Is Nothing Then

                                flag = "NotEmpty"

                            End If

                        Next

                        If flag = "Empty" Then

                            worksheet.Range(selectedRng.Cells(1, k), selectedRng.Cells(selectedRng.Rows.Count, k)).Delete(Excel.XlDeleteShiftDirection.xlShiftToLeft)
                            blankColCount += 1

                        End If

                    Next

                Next

                'if no blank columns are found in a sheet then go to the "nextloop" section and skip the lines in between
                If blankColCount = 0 Then
                    GoTo break2
                End If

                'valueFlag is "Empty" means the entire selection is blank
                'so the msgbox will be shown  and then exit sub                
                If ValueFlag = "Empty" Then
                    MsgBox(blankColCount & " Column(s) are deleted.", MsgBoxStyle.Information, "SOFTEKO")
                    Exit Sub
                End If

                selectedRng.Cells(1, 1).select

break2:
                'displays a msgbox that shows how many columns are deleted
                MsgBox(blankColCount & " Column(s) are deleted.", MsgBoxStyle.Information, "SOFTEKO")

            End If


        Catch ex As Exception

        End Try


    End Sub

    ''' <summary>
    ''' removes all blank columns from Active Sheet
    ''' </summary>
    Private Sub Button34_Click(sender As Object, e As RibbonControlEventArgs) Handles Button34.Click
        Try

            Dim excelApp As Excel.Application
            Dim workbook As Excel.Workbook
            Dim worksheet As Excel.Worksheet
            Dim selectedRng As Excel.Range
            Dim blankColCount As Integer = 0
            Dim flag As String = "Empty"

            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            worksheet = workbook.ActiveSheet
            selectedRng = excelApp.Selection

            'use the UsedRange method to find the address of the range used in the active sheet
            'use split function to get 2nd portion of the range which is the last cell of the used range
            'Use this addrees to to find row and column number of last cell
            Dim lastCell() As String
            Dim lastColNum As Integer
            Dim lastRowNum As Integer

            lastCell = worksheet.UsedRange.Address.Split(":"c)
            lastColNum = worksheet.Range(lastCell(1)).Column
            lastRowNum = worksheet.Range(lastCell(1)).Row


            'loop through each cells of a column of the active sheet and check if the column is empty or not.
            'if all the cells of the column of that seleted range are blank then the "flag" variable remains "Empty". If any of the cell of that column is non-empty then "flag" will be "NotEmpty"
            'if flag is "Empty" then the respective column of the active sheet will be deleted.
            'after checking a column, the "flag" variable resets to "Empty"
            For i = lastColNum To 1 Step -1
                flag = "Empty"
                For j = lastRowNum To 1 Step -1
                    If Not worksheet.Cells(j, i).value Is Nothing Then

                        flag = "NotEmpty"

                    End If

                Next

                If flag = "Empty" Then

                    worksheet.Cells(1, i).entirecolumn.delete()

                    blankColCount += 1

                End If
            Next

            'if no blank columns are found in a sheet then go to the "nextloop" section and skip the lines in between
            If blankColCount = 0 Then
                GoTo break
            End If

            worksheet.Cells(1, 1).select

break:
            'displays a msgbox that shows how many columns are deleted
            MsgBox(blankColCount & " Column(s) are deleted.", MsgBoxStyle.Information, "SOFTEKO")

        Catch ex As Exception

        End Try


    End Sub


    ''' <summary>
    ''' removes blank columns from the selected worksheets
    ''' </summary>

    Private Sub Button35_Click(sender As Object, e As RibbonControlEventArgs) Handles Button35.Click

        Try


            Dim excelApp As Excel.Application
            Dim workbook As Excel.Workbook
            Dim worksheet As Excel.Worksheet
            Dim confirmationMsg As String = ""
            Dim blankColCount As Integer
            Dim i As Integer = 0
            Dim flag As String = "NotEmpty"


            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook

            'takes the sheet names of the selected worksheets
            Dim selectedSheets As Excel.Sheets = excelApp.ActiveWindow.SelectedSheets
            Dim sheetName As String = ""

            'loops through each selected worksheet and concatenate all the sheet names togehter in the "sheetName" variable
            'then Right function removes the leading comma (,) from the "sheetName" variable
            For Each sheet As Excel.Worksheet In selectedSheets
                sheetName = sheetName & "," & sheet.Name
            Next
            sheetName = Right(sheetName, Len(sheetName) - 1)

            'new array (arrSheetName) stores all the selected sheet names separately
            Dim arrSheetName As String() = Split(sheetName, ",")


            'loops through each selected sheet name from the "arrSheetName" array
            '"worksheet" variable takes the sheets name from the array and makes it active worksheet
            'each time a new sheet is taken from the slected sheets, "blankColCount" resets to 0
            For i = 0 To UBound(arrSheetName)
                blankColCount = 0
                worksheet = workbook.Sheets(arrSheetName(i))
                worksheet.Activate()


                'use the UsedRange method to find the address of the range used in the active sheet
                'use split function to get 2nd portion of the range which is the last cell of the used range
                'Use this addrees to to find row and column number of last cell
                Dim lastCell() As String
                Dim lastRowNum As Integer
                Dim lastColNum As Integer

                lastCell = worksheet.UsedRange.Address.Split(":"c)
                lastRowNum = worksheet.Range(lastCell(1)).Row
                lastColNum = worksheet.Range(lastCell(1)).Column

                'loop through each cells of a column of the active sheet and check if the column is empty or not.
                'if all the cells of the column of that seleted range are blank then the "flag" variable remains "Empty". If any of the cell of that column is non-empty then "flag" will be "NotEmpty"
                'if flag is "Empty" then the respective column of the active sheet will be deleted.
                'after checking a column, the "flag" variable resets to "Empty"
                For j = lastColNum To 1 Step -1
                    flag = "Empty"
                    For k = lastRowNum To 1 Step -1
                        If Not worksheet.Cells(k, j).value Is Nothing Then

                            flag = "NotEmpty"

                        End If

                    Next

                    If flag = "Empty" Then

                        worksheet.Cells(1, j).entirecolumn.delete()

                        blankColCount += 1

                    End If
                Next

                'if no blank columns are found in a sheet then go to the "nextloop" section and skip the lines in between
                If blankColCount = 0 Then
                    GoTo nextloop
                End If


nextloop:
                'stores information about how many columns deleted from which sheet
                confirmationMsg = confirmationMsg & blankColCount & " Column(s) are deleted from " & arrSheetName(i) & vbCrLf

            Next

            'finally this msgBox is shown
            MsgBox(confirmationMsg, MsgBoxStyle.Information, "SOFTEKO")

        Catch ex As Exception

        End Try



    End Sub


    ''' <summary>
    ''' removes blank columns from all worksheets from the active workbook
    ''' </summary>

    Private Sub Button36_Click(sender As Object, e As RibbonControlEventArgs) Handles Button36.Click

        Try

            Dim excelApp As Excel.Application
            Dim workbook As Excel.Workbook
            Dim worksheet As Excel.Worksheet
            Dim confirmationMsg As String = ""
            Dim blankColCount As Integer
            Dim i As Integer = 0
            Dim flag As String = "Empty"


            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook

            'takes the sheet names of all worksheets of the workbook
            Dim selectedSheets As Excel.Sheets = excelApp.Sheets
            Dim sheetName As String = ""

            'loops through each selected worksheet and concatenate all the sheet names togehter in the "sheetName" variable
            'then Right function removes the leading comma (,) from the "sheetName" variable
            For Each sheet As Excel.Worksheet In selectedSheets
                sheetName = sheetName & "," & sheet.Name
            Next
            sheetName = Right(sheetName, Len(sheetName) - 1)

            'new array (arrSheetName) stores all the sheet names separately
            Dim arrSheetName As String() = Split(sheetName, ",")


            'loops through each sheet name from the "arrSheetName" array
            '"worksheet" variable takes the sheet names from the array and makes it active worksheet
            'each time a new sheet is taken by "worksheet" variable, "blankColList" and "blankColCount" resets to 0
            For i = 0 To UBound(arrSheetName)
                blankColCount = 0
                worksheet = workbook.Sheets(arrSheetName(i))
                worksheet.Activate()


                'use the UsedRange method to find the address of the range used in the active sheet
                'use split function to get 2nd portion of the range which is the last cell of the used range
                'Use this addrees to to find column number of last cell
                Dim lastCell() As String
                Dim lastRowNum As Integer
                Dim lastColNum As Integer

                lastCell = worksheet.UsedRange.Address.Split(":"c)
                lastRowNum = worksheet.Range(lastCell(1)).Row
                lastColNum = worksheet.Range(lastCell(1)).Column

                'loop through each cells of a column of the active sheet and check if the column is empty or not.
                'if all the cells of the column of that seleted range are blank then the "flag" variable remains "Empty". If any of the cell of that column is non-empty then "flag" will be "NotEmpty"
                'if flag is "Empty" then the respective column of the active sheet will be deleted.
                'after checking a column, the "flag" variable resets to "Empty"
                For j = lastColNum To 1 Step -1
                    flag = "Empty"
                    For k = lastRowNum To 1 Step -1
                        If Not worksheet.Cells(k, j).value Is Nothing Then

                            flag = "NotEmpty"

                        End If

                    Next

                    If flag = "Empty" Then

                        worksheet.Cells(1, j).entirecolumn.delete()

                        blankColCount += 1

                    End If
                Next

                'if no blank columns are found in a sheet then go to the "nextloop" section and skip the lines in between
                If blankColCount = 0 Then
                    GoTo nextloop
                End If

nextloop:
                'stores information about how many columns deleted from which sheet
                confirmationMsg = confirmationMsg & blankColCount & " Column(s) are deleted from " & arrSheetName(i) & vbCrLf

            Next

            'finally this msgBox is shown
            MsgBox(confirmationMsg, MsgBoxStyle.Information, "SOFTEKO")

        Catch ex As Exception

        End Try


    End Sub


    ''' <summary>
    ''' removes blank rows from selected range of the active worksheet
    ''' </summary>

    Private Sub Button37_Click(sender As Object, e As RibbonControlEventArgs) Handles Button37.Click


        Try

            Dim excelApp As Excel.Application
            Dim workbook As Excel.Workbook
            Dim worksheet As Excel.Worksheet
            Dim selectedRng As Excel.Range
            Dim blankRowCount As Integer = 0
            Dim flag As String = "Empty"
            Dim ValueFlag As String = "Empty"

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


                'loops through the entire selection and if the entire selection is empty or not
                'if the entire selection is blank then valueFlag will become "NotEmpty"
                For i = 1 To selectedRng.Rows.Count
                    For j = 1 To selectedRng.Columns.Count
                        If Not selectedRng.Cells(i, j).value Is Nothing Then
                            ValueFlag = "NotEmpty"
                        End If
                    Next
                Next



                'loop through each cells of a row of the selection and check if the row is empty or not.
                'if all the cells of the row of that seleted range are blank then the "flag" variable remains Empty. If any of the cell of that row is non-empty then "flag" will be "NotEmpty"
                'if flag is "Empty" the blank row is deleted and increase the value of "blankRowCount" by 1
                'after checking a row, the "flag" variable resets to "Empty"
                For i = selectedRng.Rows.Count To 1 Step -1
                    flag = "Empty"
                    For j = selectedRng.Columns.Count To 1 Step -1
                        If Not selectedRng.Cells(i, j).value Is Nothing Then

                            flag = "NotEmpty"

                        End If

                    Next


                    If flag = "Empty" Then

                        worksheet.Range(selectedRng.Cells(i, 1), selectedRng.Cells(i, selectedRng.Columns.Count)).Delete(Excel.XlDeleteShiftDirection.xlShiftUp)
                        blankRowCount = blankRowCount + 1

                    End If

                Next

                'if no blank rows are found in a sheet then go to the "break1" section and skip the lines in between
                If blankRowCount = 0 Then
                    GoTo break1
                End If


                'valueFlag is "Empty" means the entire selection is blank
                'so the msgbox will be shown  and then exit sub   
                If ValueFlag = "Empty" Then
                    MsgBox(blankRowCount & " Row(s) are deleted.", MsgBoxStyle.Information, "SOFTEKO")
                    Exit Sub
                End If

                selectedRng.Cells(1, 1).select
break1:
                'displays a msgbox that shows how many rows are deleted
                MsgBox(blankRowCount & " Row(s) are deleted.", MsgBoxStyle.Information, "SOFTEKO")



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

                'loops through the entire selection and if the entire selection is empty or not
                'if the entire selection is blank then valueFlag will become "NotEmpty"
                For i = 0 To UBound(arrRng)
                    selectedRng = worksheet.Range(arrRng(i))
                    For j = 1 To selectedRng.Rows.Count
                        For k = 1 To selectedRng.Columns.Count
                            If Not selectedRng.Cells(j, k).value Is Nothing Then
                                ValueFlag = "NotEmpty"
                            End If
                        Next
                    Next
                Next

                'loop through each range of the selection and remove blank rows
                For i = 0 To UBound(arrRng)

                    selectedRng = worksheet.Range(arrRng(i))

                    'loop through each cells of a row of the selection and check if the row is empty or not.
                    'if all the cells of the row of that seleted range are blank then the "flag" variable remains Empty. If any of the cell of that row is non-empty then "flag" will be "NotEmpty"
                    'if flag is "Empty" the blank row is deleted and increase the value of "blankRowCount" by 1
                    'after checking a row, the "flag" variable resets to "Empty"
                    For j = selectedRng.Rows.Count To 1 Step -1
                        flag = "Empty"
                        For k = selectedRng.Columns.Count To 1 Step -1
                            If Not selectedRng.Cells(j, k).value Is Nothing Then

                                flag = "NotEmpty"

                            End If

                        Next

                        If flag = "Empty" Then

                            worksheet.Range(selectedRng.Cells(j, 1), selectedRng.Cells(j, selectedRng.Columns.Count)).Delete(Excel.XlDeleteShiftDirection.xlShiftUp)
                            blankRowCount = blankRowCount + 1

                        End If

                    Next

                Next

                'if no blank rows are found in a sheet then go to the "break2" section and skip the lines in between
                If blankRowCount = 0 Then
                    GoTo break2
                End If

                'valueFlag is "Empty" means the entire selection is blank
                'so the msgbox will be shown  and then exit sub  
                If ValueFlag = "Empty" Then
                    MsgBox(blankRowCount & " Row(s) are deleted.", MsgBoxStyle.Information, "SOFTEKO")
                    Exit Sub
                End If

                selectedRng.Cells(1, 1).select

break2:
                'displays a msgbox that shows how many rows are deleted
                MsgBox(blankRowCount & " Row(s) are deleted.", MsgBoxStyle.Information, "SOFTEKO")

            End If


        Catch ex As Exception

        End Try


    End Sub

    ''' <summary>
    ''' removes all blank rows from active worksheet
    ''' </summary>

    Private Sub Button38_Click(sender As Object, e As RibbonControlEventArgs) Handles Button38.Click
        Try

            Dim excelApp As Excel.Application
            Dim workbook As Excel.Workbook
            Dim worksheet As Excel.Worksheet
            Dim selectedRng As Excel.Range
            Dim blankRowList As String = ""
            Dim blankRowCount As Integer = 0
            Dim i As Integer
            Dim flag As String = "Empty"

            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            worksheet = workbook.ActiveSheet
            selectedRng = excelApp.Selection

            'use the UsedRange method to find the address of the range used in the active sheet
            'use split function to get 2nd portion of the range which is the last cell of the used range
            'Use this addrees to to find row number of last cell
            Dim lastCell() As String
            Dim lastRowNum As Integer
            Dim lastColNum As Integer

            lastCell = worksheet.UsedRange.Address.Split(":"c)
            lastRowNum = worksheet.Range(lastCell(1)).Row
            lastColNum = worksheet.Range(lastCell(1)).Column



            'loop through each rows of the active sheet upto the last row number
            'check if the entire row is empty or not by using the "IsEmptyRow" function



            For i = lastRowNum To 1 Step -1
                flag = "Empty"
                For j = lastColNum To 1 Step -1
                    If Not worksheet.Cells(i, j).value Is Nothing Then

                        flag = "NotEmpty"

                    End If

                Next

                If flag = "Empty" Then

                    worksheet.Cells(i, 1).entirerow.delete()

                    blankRowCount += 1

                End If
            Next

            'if no blank rows are found in a sheet then go to the "nextloop" section and skip the lines in between
            If blankRowCount = 0 Then
                GoTo break
            End If

            worksheet.Cells(1, 1).select

break:
            'displays a msgbox that shows how many rows are deleted
            MsgBox(blankRowCount & " Row(s) are deleted.", MsgBoxStyle.Information, "SOFTEKO")

        Catch ex As Exception

        End Try




        'Dim excelApp As Excel.Application = Nothing
        'Dim workbook As Excel.Workbook = Nothing
        'Dim worksheet As Excel.Worksheet = Nothing
        'Dim selectedRng As Excel.Range = Nothing
        'Dim range As Excel.Range = Nothing
        'Dim blankRowList As String = ""
        'Dim blankRowCount As Integer = 0

        'Try
        '    excelApp = Globals.ThisAddIn.Application
        '    workbook = excelApp.ActiveWorkbook
        '    worksheet = workbook.ActiveSheet
        '    selectedRng = excelApp.Selection

        '    Dim lastCell() As String
        '    Dim lastRowNum As Long
        '    Dim lastColNum As Long

        '    lastCell = worksheet.UsedRange.Address.Split(":"c)
        '    lastRowNum = worksheet.Range(lastCell(1)).Row
        '    lastColNum = worksheet.Range(lastCell(1)).Column

        '    range = worksheet.Range(worksheet.Cells(1, 1), worksheet.Cells(lastRowNum, lastColNum))

        '    MsgBox(lastRowNum)
        '    MsgBox(lastColNum)



        '    For i = 1 To lastRowNum
        '        For j = 1 To lastColNum


        '            Dim currentRow As Excel.Range = range.Rows(i)
        '            'If excelApp.WorksheetFunction.CountA(currentRow) = 0 Then
        '            If Not worksheet.Cells(i, j).value IsNot Nothing AndAlso worksheet.Cells(i, j).ToString().Trim() <> "" Then
        '                If i = 128 Then
        '                    MsgBox("y")
        '                End If
        '                Exit Sub
        '                'blankRowList &= "," & worksheet.Range(worksheet.Cells(i, 1), worksheet.Cells(i, 2)).Address
        '                'blankRowCount += 1
        '            End If
        '            'Marshal.ReleaseComObject(currentRow)
        '        Next
        '    Next

        '        If blankRowCount > 0 Then
        '        blankRowList = Right(blankRowList, Len(blankRowList) - 1)
        '        worksheet.Range(blankRowList).EntireRow.Delete()
        '        worksheet.Cells(1, 1).Select()
        '        MsgBox($"{blankRowCount} Row(s) are deleted.", MsgBoxStyle.Information, "SOFTEKO")
        '    End If

        'Catch ex As Exception
        '    ' Handle exceptions
        'Finally
        '    ' Release and cleanup COM objects
        '    If Not range Is Nothing Then Marshal.ReleaseComObject(range)
        '    If Not selectedRng Is Nothing Then Marshal.ReleaseComObject(selectedRng)
        '    If Not worksheet Is Nothing Then Marshal.ReleaseComObject(worksheet)
        '    If Not workbook Is Nothing Then Marshal.ReleaseComObject(workbook)

        '    GC.Collect()
        '    GC.WaitForPendingFinalizers()
        '    GC.Collect()
        '    GC.WaitForPendingFinalizers()
        'End Try









    End Sub

    ''' <summary>
    ''' removes all blank rows from all selected worksheets
    ''' </summary>

    Private Sub Button39_Click(sender As Object, e As RibbonControlEventArgs) Handles Button39.Click


        Try


            Dim excelApp As Excel.Application
            Dim workbook As Excel.Workbook
            Dim worksheet As Excel.Worksheet
            Dim blankRowList As String
            Dim confirmationMsg As String = ""
            Dim blankRowCount As Integer
            Dim i As Integer = 0
            Dim flag As String = "Empty"


            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook

            'takes the sheet names of the selected worksheets
            Dim selectedSheets As Excel.Sheets = excelApp.ActiveWindow.SelectedSheets
            Dim sheetName As String = ""

            'loops through each selected worksheet and concatenate all the sheet names togehter in the "sheetName" variable
            'then Right function removes the leading comma (,) from the "sheetName" variable
            For Each sheet As Excel.Worksheet In selectedSheets
                sheetName = sheetName & "," & sheet.Name
            Next
            sheetName = Right(sheetName, Len(sheetName) - 1)

            'new array (arrSheetName) stores all the selected sheet names separately
            Dim arrSheetName As String() = Split(sheetName, ",")


            'loops through each selected sheet name from the "arrSheetName" array
            '"worksheet" variable takes the sheets name from the array and makes it active worksheet
            'each time a new sheet is taken from the slected sheets, "blankRowList" resets to empty string and "blankRowCount" resets to 0
            For i = 0 To UBound(arrSheetName)
                blankRowList = ""
                blankRowCount = 0
                worksheet = workbook.Sheets(arrSheetName(i))
                worksheet.Activate()


                'use the UsedRange method to find the address of the range used in the active sheet
                'use split function to get 2nd portion of the range which is the last cell of the used range
                'Use this addrees to to find row number of last cell
                Dim lastCell() As String
                Dim lastRowNum As Integer
                Dim lastColNum As Integer

                lastCell = worksheet.UsedRange.Address.Split(":"c)
                lastRowNum = worksheet.Range(lastCell(1)).Row
                lastColNum = worksheet.Range(lastCell(1)).Column


                'loop through each rows of the active sheet upto the last row number
                'check if the entire column is empty or not by using the "IsRowEmpty" function
                For j = lastRowNum To 1 Step -1
                    flag = "Empty"
                    For k = lastColNum To 1 Step -1
                        If Not worksheet.Cells(j, k).value Is Nothing Then

                            flag = "NotEmpty"

                        End If

                    Next

                    If flag = "Empty" Then

                        worksheet.Cells(j, 1).entirerow.delete()

                        blankRowCount += 1

                    End If
                Next

                'if no blank rows are found in a sheet then go to the "nextloop" section and skip the lines in between
                If blankRowCount = 0 Then
                    GoTo nextloop
                End If


nextloop:
                'stores information about how many rows deleted from which sheet
                confirmationMsg = confirmationMsg & blankRowCount & " Row(s) are deleted from " & arrSheetName(i) & vbCrLf

            Next

            'finally this msgBox is shown
            MsgBox(confirmationMsg, MsgBoxStyle.Information, "SOFTEKO")

        Catch ex As Exception

        End Try

    End Sub

    ''' <summary>
    ''' removes all blank rows from all worksheets from active workbook
    ''' </summary>

    Private Sub Button40_Click(sender As Object, e As RibbonControlEventArgs) Handles Button40.Click


        Try

            Dim excelApp As Excel.Application
            Dim workbook As Excel.Workbook
            Dim worksheet As Excel.Worksheet
            Dim confirmationMsg As String = ""
            Dim blankRowCount As Integer
            Dim i As Integer = 0
            Dim flag As String = "Empty"

            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook

            'takes the sheet names of all worksheets of the workbook
            Dim selectedSheets As Excel.Sheets = excelApp.Sheets
            Dim sheetName As String = ""

            'loops through each selected worksheet and concatenate all the sheet names togehter in the "sheetName" variable
            'then Right function removes the leading comma (,) from the "sheetName" variable
            For Each sheet As Excel.Worksheet In selectedSheets
                sheetName = sheetName & "," & sheet.Name
            Next
            sheetName = Right(sheetName, Len(sheetName) - 1)

            'new array (arrSheetName) stores all the sheet names separately
            Dim arrSheetName As String() = Split(sheetName, ",")


            'loops through each sheet name from the "arrSheetName" array
            '"worksheet" variable takes the sheet names from the array and makes it active worksheet
            'each time a new sheet is taken by "worksheet" variable, "blankRowList" resets to empty string and "blankRowCount" resets to 0
            For i = 0 To UBound(arrSheetName)
                blankRowCount = 0
                worksheet = workbook.Sheets(arrSheetName(i))
                worksheet.Activate()


                'use the UsedRange method to find the address of the range used in the active sheet
                'use split function to get 2nd portion of the range which is the last cell of the used range
                'Use this addrees to to find row number of last cell
                Dim lastCell() As String
                Dim lastRowNum As Integer
                Dim lastColNum As Integer

                lastCell = worksheet.UsedRange.Address.Split(":"c)
                lastRowNum = worksheet.Range(lastCell(1)).Row
                lastColNum = worksheet.Range(lastCell(1)).Column


                'loop through each rows of the active sheet upto the last row number
                'check if the entire column is empty or not by using the "IsRowEmpty" function
                For j = lastRowNum To 1 Step -1
                    flag = "Empty"
                    For k = lastColNum To 1 Step -1
                        If Not worksheet.Cells(j, k).value Is Nothing Then

                            flag = "NotEmpty"

                        End If

                    Next

                    If flag = "Empty" Then

                        worksheet.Cells(j, 1).entirerow.delete()

                        blankRowCount += 1

                    End If
                Next


                'if no blank rows are found in a sheet then go to the "nextloop" section and skip the lines in between
                If blankRowCount = 0 Then
                    GoTo nextloop
                End If

nextloop:
                'stores information about how many rows deleted from which sheet
                confirmationMsg = confirmationMsg & blankRowCount & " Row(s) are deleted from " & arrSheetName(i) & vbCrLf

            Next

            'finally this msgBox is shown
            MsgBox(confirmationMsg, MsgBoxStyle.Information, "SOFTEKO")

        Catch ex As Exception

        End Try


    End Sub

    ''' <summary>
    ''' removes empty sheets from the active workbook
    ''' </summary>

    Private Sub Button41_Click(sender As Object, e As RibbonControlEventArgs) Handles Button41.Click

        Try

            Dim excelApp As Excel.Application
            Dim workbook As Excel.Workbook
            Dim worksheet As Excel.Worksheet
            Dim blankWsCount As Integer = 0
            Dim i As Integer = 0
            Dim flag As String
            Dim answer As MsgBoxResult
            Dim initialWs As String

            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook

            '"initialWs" variable stores the name of the worksheet, where the button event was clicked
            initialWs = excelApp.ActiveSheet.name

            'takes the sheet names of all worksheets of the workbook
            Dim selectedSheets As Excel.Sheets = excelApp.Sheets
            Dim sheetName As String = ""

            'loops through each selected worksheet and concatenate all the sheet names togehter in the "sheetName" variable
            'then Right function removes the leading comma (,) from the "sheetName" variable
            For Each sheet As Excel.Worksheet In selectedSheets
                sheetName = sheetName & "," & sheet.Name
            Next
            sheetName = Right(sheetName, Len(sheetName) - 1)

            'new array (arrSheetName) stores all the sheet names separately
            Dim arrSheetName As String() = Split(sheetName, ",")



            'this loops only counts the number of empty WS present in the active workbook
            'loops through each selected sheet name from the "arrSheetName" array
            '"worksheet" variable takes the sheets name from the array and makes it active worksheet
            For i = 0 To UBound(arrSheetName)
                flag = "NotEmpty"
                worksheet = workbook.Sheets(arrSheetName(i))
                worksheet.Activate()


                'loop thorugh the characters of address of the used range of a worksheet
                'check if it conrians ":". If the WS is empty, used range will be single cell and the address will not have any ":" in it
                'so, "usedCellCount" will be 0 for an empty WS
                Dim usedCellCount As Integer = 0
                For Each c As Char In worksheet.UsedRange.Address

                    If c = ":" Then
                        usedCellCount += 1
                    End If

                Next


                'make sure the WS is actually empty or not by checking the value of the used range (which is already a single cell)
                'if there is no value then flag becomes "Empty"
                If usedCellCount = 0 Then
                    If worksheet.UsedRange.Value IsNot Nothing Then
                        flag = "NotEmpty"
                    Else
                        flag = "Empty"
                    End If

                End If

                'increase the value of "blankWsCount" by 1 if an empty WS is found
                If flag = "Empty" Then

                    blankWsCount += 1

                End If

            Next

            'if no blank WS is found then display this message and exit sub
            If blankWsCount = 0 Then
                MsgBox("No empty worksheet is found.", MsgBoxStyle.Information, "SOFTEKO")
                Exit Sub
            End If

            workbook.Sheets(initialWs).activate()

            'assign the reponse of user from the msgbox to the "answer" variable
            answer = MsgBox(blankWsCount & " empty worksheet(s) will be deleted. Please click Yes to continue.", MsgBoxStyle.YesNo, "SOFTEKO")

            If answer = MsgBoxResult.Yes Then

                'this loop deletes the empty worksheets
                'mechanism is same as previous loop
                For i = 0 To UBound(arrSheetName)
                    flag = "NotEmpty"
                    worksheet = workbook.Sheets(arrSheetName(i))
                    worksheet.Activate()

                    Dim usedCellCount As Integer = 0
                    For Each c As Char In worksheet.UsedRange.Address

                        If c = ":" Then
                            usedCellCount += 1
                        End If

                    Next

                    If usedCellCount = 0 Then
                        If worksheet.UsedRange.Value IsNot Nothing Then
                            flag = "NotEmpty"
                        Else
                            flag = "Empty"
                        End If

                    End If

                    If flag = "Empty" Then

                        worksheet.Delete()

                    End If

                Next

                workbook.Sheets(initialWs).activate()

                'finally this msgBox is shown
                MsgBox(blankWsCount & " worksheet(s) are deleted.", MsgBoxStyle.Information, "SOFTEKO")
            Else
                Exit Sub
            End If


        Catch ex As Exception

        End Try

    End Sub

    Private Sub DropDown1_SelectionChanged(sender As Object, e As RibbonControlEventArgs)

    End Sub

    Private Sub ComboBox1_TextChanged(sender As Object, e As RibbonControlEventArgs)

    End Sub

    Private Sub Button9_Click(sender As Object, e As RibbonControlEventArgs)

    End Sub

    Private Sub Button10_Click(sender As Object, e As RibbonControlEventArgs)

    End Sub

    Private Sub Button55_Click(sender As Object, e As RibbonControlEventArgs)

    End Sub

    Private Sub Button44_Click(sender As Object, e As RibbonControlEventArgs)

    End Sub

    Private Sub Button43_Click(sender As Object, e As RibbonControlEventArgs)

    End Sub

    Private Sub Button42_Click(sender As Object, e As RibbonControlEventArgs)

    End Sub

    Private Sub Button2_Click_1(sender As Object, e As RibbonControlEventArgs)

    End Sub

    Private Sub Button16_Click(sender As Object, e As RibbonControlEventArgs) Handles Button16.Click

        If form_flag = False Then
            Dim MyForm18 As New Form18_CombineRanges

            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            worksheet = workbook.ActiveSheet

            Dim selection As Excel.Range = CType(excelApp.Selection, Excel.Range)

            If IsRangeEmpty(selection) = True Then
                MessageBox.Show("You have not selected any data.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)

            Else
                MyForm18.TextBox1.Text = selection.Address
                MyForm18.ComboBox1.SelectedIndex = -1
                MyForm18.ComboBox1.Text = "SOFTEKO"
                MyForm18.Show()
                MyForm18.RadioButton1.Checked = True
                form_flag = True
            End If
        End If
    End Sub

    Private Sub Button20_Click_1(sender As Object, e As RibbonControlEventArgs) Handles Button20.Click
        If form_flag = False Then
            Dim MyForm22 As New Form22_Merge_Duplicate_Rows

            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            worksheet = workbook.ActiveSheet

            Dim selection As Excel.Range = CType(excelApp.Selection, Excel.Range)

            If IsRangeEmpty(selection) = True Then
                MessageBox.Show("You have not selected any data.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)

            Else
                MyForm22.TextBox1.Text = selection.Address
                MyForm22.ComboBox1.SelectedIndex = -1
                MyForm22.ComboBox1.Text = "SOFTEKO"
                MyForm22.Show()
                form_flag = True
            End If
        End If
    End Sub

    Private Sub Button21_Click_1(sender As Object, e As RibbonControlEventArgs) Handles Button21.Click
        If form_flag = False Then
            Dim MyForm23 As New Form23_Merge_Duplicate_Columns

            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            worksheet = workbook.ActiveSheet

            Dim selection As Excel.Range = CType(excelApp.Selection, Excel.Range)

            If IsRangeEmpty(selection) = True Then
                MessageBox.Show("You have not selected any data.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)

            Else
                MyForm23.TextBox1.Text = selection.Address
                MyForm23.ComboBox1.SelectedIndex = -1
                MyForm23.ComboBox1.Text = "SOFTEKO"
                MyForm23.Show()
                form_flag = True
            End If
        End If
    End Sub

    Private Sub Button23_Click_1(sender As Object, e As RibbonControlEventArgs) Handles Button23.Click
        If form_flag = False Then
            Dim MyForm25 As New Form25_Split_Range

            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            worksheet = workbook.ActiveSheet

            Dim selection As Excel.Range = CType(excelApp.Selection, Excel.Range)

            If IsRangeEmpty(selection) = True Then
                MessageBox.Show("You have not selected any data.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)

            Else
                MyForm25.TextBox1.Text = selection.Address
                MyForm25.ComboBox1.SelectedIndex = -1
                MyForm25.ComboBox1.Text = "SOFTEKO"
                MyForm25.Show()
                form_flag = True
            End If
        End If
    End Sub

    Private Sub Button22_Click_1(sender As Object, e As RibbonControlEventArgs) Handles Button22.Click
        If form_flag = False Then
            Dim MyForm24 As New Form24_Split_Cells

            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            worksheet = workbook.ActiveSheet

            Dim selection As Excel.Range = CType(excelApp.Selection, Excel.Range)

            If IsRangeEmpty(selection) = True Then
                MessageBox.Show("You have not selected any data.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)

            Else
                MyForm24.TextBox1.Text = selection.Address
                MyForm24.ComboBox1.SelectedIndex = -1
                MyForm24.ComboBox1.Text = "SOFTEKO"
                MyForm24.Show()
                form_flag = True
            End If
        End If
    End Sub

    Private Sub Button45_Click(sender As Object, e As RibbonControlEventArgs) Handles Button45.Click
        If form_flag = False Then
            Dim MyForm26 As New Form26_split_text_bycharacters

        excelApp = Globals.ThisAddIn.Application
        workbook = excelApp.ActiveWorkbook
        worksheet = workbook.ActiveSheet

        Dim selection As Excel.Range = CType(excelApp.Selection, Excel.Range)

            If IsRangeEmpty(selection) = True Then
                MessageBox.Show("You have not selected any data.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)

            Else
                MyForm26.TB_source_range.Text = selection.Address
                MyForm26.ComboBox1.SelectedIndex = -1
                MyForm26.ComboBox1.Text = "SOFTEKO"
                MyForm26.Show()
                form_flag = True
            End If
        End If
    End Sub

    Private Sub Button46_Click(sender As Object, e As RibbonControlEventArgs) Handles Button46.Click
        If form_flag = False Then
            Dim Source As String = "Absbsjdwd,hdwdiqd,djd"
        Dim pattern As String = "***,*,"
        Dim KeepSeparator As Boolean = True
        Dim Consecutive As Boolean = True
        Dim Before As Boolean = True

        Dim Values() As String
        Values = SplitText(Source, pattern, Consecutive, KeepSeparator, Before)

        For i = LBound(Values) To UBound(Values)
            MsgBox(Values(i))
        Next

        Dim MyForm27 As New Form27_Split_text_bystrings

        excelApp = Globals.ThisAddIn.Application
        workbook = excelApp.ActiveWorkbook
        worksheet = workbook.ActiveSheet

        Dim selection As Excel.Range = CType(excelApp.Selection, Excel.Range)

            If IsRangeEmpty(selection) = True Then
                MessageBox.Show("You have not selected any data.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)

            Else
                MyForm27.TB_source_range.Text = selection.Address
                MyForm27.ComboBox1.SelectedIndex = -1
                MyForm27.ComboBox1.Text = "SOFTEKO"
                MyForm27.Show()
                form_flag = True
            End If
        End If

    End Sub


    Private Sub Button49_Click(sender As Object, e As RibbonControlEventArgs) Handles Button49.Click
        If form_flag = False Then
            Dim MyForm33 As New Form33_ColorBasedDropDownList
            MyForm33.Show()
            form_flag = True
        End If
    End Sub

    Private Sub Button54_Click(sender As Object, e As RibbonControlEventArgs) Handles Button54.Click
        If form_flag = False Then
            Dim form As New Form12HideRanges

            form.Show()
            form_flag = True
        End If
    End Sub

    Private Sub Button11_Click_1(sender As Object, e As RibbonControlEventArgs) Handles Button11.Click
        If form_flag = False Then
            Dim form As New Form13HideAllExceptSelectedRange

            form.Show()
            form_flag = True
        End If
    End Sub

    Private Sub Button12_Click_1(sender As Object, e As RibbonControlEventArgs) Handles Button12.Click
        If form_flag = False Then
            Dim Myform As New Form14SpecifyScrollArea
            Myform.Show()
            form_flag = True
        End If
    End Sub

    Private Sub Button2_Click_2(sender As Object, e As RibbonControlEventArgs) Handles Button2.Click
        If form_flag = False Then
            Dim form As New Form29_Simple_Drop_down_List
            form.Show()
            form_flag = True
        End If
    End Sub

    Private Sub Button9_Click_1(sender As Object, e As RibbonControlEventArgs) Handles Button9.Click
        If form_flag = False Then
            Dim form As New Form34_PictureBasedDropdownList

            form.Show()
            form_flag = True
        End If
    End Sub

    Private Sub Button29_Click(sender As Object, e As RibbonControlEventArgs) Handles Button29.Click
        If form_flag = False Then
            Dim form As New Form31_UpdateDynamicDropdownList

            form.Show()
            form_flag = True
        End If
    End Sub

    Private Sub Button30_Click(sender As Object, e As RibbonControlEventArgs) Handles Button30.Click
        If form_flag = False Then
            Dim form As New Form32_ExtendDropDownList

            form.Show()
            form_flag = True
        End If
    End Sub

    Private Sub Button24_Click_1(sender As Object, e As RibbonControlEventArgs) Handles Button24.Click
        If form_flag = False Then
            settingflag1 = False
            Dim form As New Form35Multi_SelectionbasedDropdown

            form.Show()
            form_flag = True
        End If
    End Sub

    Private Sub Button25_Click_1(sender As Object, e As RibbonControlEventArgs) Handles Button25.Click
        If form_flag = False Then
            settingflag2 = False
            Dim form As New Form37_MSDropDownCheckBox

            form.Show()
            form_flag = True
        End If
    End Sub

    Private Sub Button26_Click_1(sender As Object, e As RibbonControlEventArgs) Handles Button26.Click
        If form_flag = False Then
            Dim form As New Form39_DropdownlistwithSearchOption

            form.Show()
            form_flag = True
        End If
    End Sub

    Private Sub Button27_Click_1(sender As Object, e As RibbonControlEventArgs) Handles Button27.Click
        If form_flag = False Then
            Dim form As New Form41_RemoveAdavancedDropdownList

            form.Show()
            form_flag = True
        End If
    End Sub

    Private Sub Button17_Click(sender As Object, e As RibbonControlEventArgs) Handles Button17.Click
        If form_flag = False Then
            Dim MyForm18 As New Form18_CombineRanges

            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            worksheet = workbook.ActiveSheet

            Dim selection As Excel.Range = CType(excelApp.Selection, Excel.Range)

            If IsRangeEmpty(selection) = True Then
                MessageBox.Show("You have not selected any data.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)

            Else
                MyForm18.TextBox1.Text = selection.Address
                MyForm18.ComboBox1.SelectedIndex = -1
                MyForm18.ComboBox1.Text = "SOFTEKO"
                MyForm18.Show()
                MyForm18.RadioButton2.Checked = True
                form_flag = True
            End If
        End If
    End Sub

    Private Sub Button18_Click(sender As Object, e As RibbonControlEventArgs) Handles Button18.Click

        If form_flag = False Then

            Dim MyForm18 As New Form18_CombineRanges

            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            worksheet = workbook.ActiveSheet

            Dim selection As Excel.Range = CType(excelApp.Selection, Excel.Range)

            If IsRangeEmpty(selection) = True Then
                MessageBox.Show("You have not selected any data.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)

            Else
                MyForm18.TextBox1.Text = selection.Address
                MyForm18.ComboBox1.SelectedIndex = -1
                MyForm18.ComboBox1.Text = "SOFTEKO"
                MyForm18.Show()
                MyForm18.RadioButton3.Checked = True
                form_flag = True
            End If
        End If
    End Sub

    Private Sub Button47_Click(sender As Object, e As RibbonControlEventArgs) Handles Button47.Click
        If form_flag = False Then
            Dim MyForm28 As New Form28_Split_text_bypattern

            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            worksheet = workbook.ActiveSheet

            Dim selection As Excel.Range = CType(excelApp.Selection, Excel.Range)

            If IsRangeEmpty(selection) = True Then
                MessageBox.Show("You have not selected any data.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)

            Else
                MyForm28.TB_source_range.Text = selection.Address
                MyForm28.ComboBox1.SelectedIndex = -1
                MyForm28.ComboBox1.Text = "SOFTEKO"
                MyForm28.Show()
                form_flag = True
            End If
            End If

    End Sub
End Class



