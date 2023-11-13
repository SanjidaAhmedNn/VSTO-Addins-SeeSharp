Imports System.ComponentModel
Imports System.Diagnostics
Imports System.Drawing
Imports System.Runtime.ConstrainedExecution
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Excel

Public Class ThisAddIn
    Dim WithEvents excelApp As Excel.Application
    Dim workBook As Excel.Workbook
    Public Shared worksheet As Excel.Worksheet
    ' Class-level variable for your form
    Private Form As Form36 = Nothing
    Private Form2 As Form38 = Nothing
    Private Form3 As Form40 = Nothing
    Public src_rng As Excel.Range
    Public src_rng1 As Excel.Range
    Public src_rng2 As Excel.Range
    Public src_rng3 As Excel.Range
    Public src_rng4 As Excel.Range
    Public src_rng5 As Excel.Range
    Public des_rng As Excel.Range
    Public des_rng1 As Excel.Range
    Public des_rng2 As Excel.Range
    Public des_rng3 As Excel.Range
    Public des_rng4 As Excel.Range
    Public des_rng5 As Excel.Range

    Public sheetName3 As String
    Public sheetName4 As String

    Public range1 As Excel.Range
    Public range2 As Excel.Range

    Private WithEvents wsEvent1 As Excel.Worksheet
    Private WithEvents wsEvent2 As Excel.Worksheet
    Private WithEvents wsEvent3 As Excel.Worksheet
    Private WithEvents wsEvent4_1 As Excel.Worksheet
    Private WithEvents wsEvent4_2 As Excel.Worksheet
    Private WithEvents wsEvent4_3 As Excel.Worksheet
    Private WithEvents wsEvent4_4 As Excel.Worksheet
    Private WithEvents wsEvent4_5 As Excel.Worksheet
    Private WithEvents wsEvent5 As Excel.Worksheet


    Private WithEvents CurrentSheet As Excel.Worksheet
    Private WithEvents WorkbookEvents As Excel.Workbook



    Private Sub ThisAddIn_Startup(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Startup

        Globals.ThisAddIn.Application.DisplayAlerts = False
        Application.EnableEvents = True
        form_flag = False
        sessionflag1 = True
        sessionflag2 = True


        AddHandler Globals.ThisAddIn.Application.WorkbookActivate, AddressOf Workbook_Activated
    End Sub

    Private Sub Workbook_Activated(ByVal Wb As Excel.Workbook)
        excelApp = Globals.ThisAddIn.Application
        Dim workBook As Excel.Workbook = excelApp.ActiveWorkbook
        Dim worksheet As Excel.Worksheet = workBook.ActiveSheet
        RemovePreviousEventHandler() ' We'll define this function to ensure we don't attach multiple handlers.

        CheckForNewwWorksheet()
        ' MsgBox(3)
        ' MsgBox(Flag)

        HideNewwwWorksheet()

        'MsgBox(Flag)

        If Flag1 = True Then
            'MsgBox(4)

            worksheet.Range("B1").Select()  'Randomly select a cell. If nothing is selected, addhandler show error

            ' Define an array of type Excel.Worksheet
            Dim sheetsArray() As Excel.Worksheet

            ' Resize the array based on the number of sheets
            ReDim sheetsArray(workBook.Worksheets.Count)

            If SR1.Contains("Active Workbook") Then
                'MsgBox(1)
                ' Assuming you're working with the active workbook:
                AddHandler worksheet.SelectionChange, AddressOf sheet_SelectionChange1

                WorkbookEvents = Globals.ThisAddIn.Application.ActiveWorkbook
                CurrentSheet = CType(Globals.ThisAddIn.Application.ActiveSheet, Excel.Worksheet)


            ElseIf SR1 = "Select Range" Or SR1.Contains("Active Sheet") Then
                For i = 1 To excelApp.ActiveWorkbook.Worksheets.Count
                    sheetsArray(i) = CType(workBook.Worksheets(i), Excel.Worksheet)
                    If shName1 = sheetsArray(i).Name Then
                        'MsgBox(shName1)
                        wsEvent1 = DirectCast(sheetsArray(i), Excel.Worksheet)
                        'AddHandler sheetsArray(i).SelectionChange, AddressOf sheet_SelectionChange1
                        'src_rng = sheetsArray(i).Range("A1", workSheet.Cells(excelApp.Rows.Count, excelApp.Columns.Count))
                    End If
                    'i = i + 1

                Next

                'AddHandler worksheet.SelectionChange, AddressOf sheet_SelectionChange1

                'ElseIf CB_Source.Text.Contains("Active Sheet") Then

                '    AddHandler workSheet.SelectionChange, AddressOf sheet_SelectionChange
            Else

                For i = 1 To excelApp.ActiveWorkbook.Worksheets.Count
                    sheetsArray(i) = CType(workBook.Worksheets(i), Excel.Worksheet)
                    If SR1 = sheetsArray(i).Name Then
                        AddHandler sheetsArray(i).SelectionChange, AddressOf sheet_SelectionChange1
                        'src_rng = sheetsArray(i).Range("A1", workSheet.Cells(excelApp.Rows.Count, excelApp.Columns.Count))
                    End If
                    'i = i + 1

                Next
                GB_CB_Source1 = worksheet.Range("A1", worksheet.Cells(excelApp.Rows.Count, excelApp.Columns.Count)).Address

            End If
        End If

        If Flag2 = True Then
            'Form2.Refresh()

            worksheet.Range("B1").Select()  'Randomly select a cell. If nothing is selected, addhandler show error

            ' Define an array of type Excel.Worksheet
            Dim sheetsArray() As Excel.Worksheet

            ' Resize the array based on the number of sheets
            ReDim sheetsArray(workBook.Worksheets.Count)

            If SR2.Contains("Active Workbook") Then
                'MsgBox(1)
                ' Assuming you're working with the active workbook:
                AddHandler worksheet.SelectionChange, AddressOf sheet_SelectionChange2
                WorkbookEvents = Globals.ThisAddIn.Application.ActiveWorkbook
                CurrentSheet = CType(Globals.ThisAddIn.Application.ActiveSheet, Excel.Worksheet)

                'For i = 1 To excelApp.ActiveWorkbook.Worksheets.Count
                'MsgBox(excelApp.ActiveWorkbook.Worksheets.Count)
                'sheetsArray(i) = CType(workBook.Worksheets(i), Excel.Worksheet)

                ''Needs to add handler to each worksheet
                'If sheetsArray(i).Visible = Excel.XlSheetVisibility.xlSheetVisible Then
                '    AddHandler sheetsArray(i).SelectionChange, AddressOf sheet_SelectionChange
                'End If
                'Next


            ElseIf SR2 = "Select Range" Or SR2.Contains("Active Sheet") Then

                For i = 1 To excelApp.ActiveWorkbook.Worksheets.Count
                    sheetsArray(i) = CType(workBook.Worksheets(i), Excel.Worksheet)
                    If shName2 = sheetsArray(i).Name Then
                        'MsgBox(shName2)
                        wsEvent2 = DirectCast(sheetsArray(i), Excel.Worksheet)
                        'AddHandler sheetsArray(i).SelectionChange, AddressOf sheet_SelectionChange2
                        'src_rng = sheetsArray(i).Range("A1", workSheet.Cells(excelApp.Rows.Count, excelApp.Columns.Count))
                    End If
                    'i = i + 1

                Next
            Else

                For i = 1 To excelApp.ActiveWorkbook.Worksheets.Count
                    sheetsArray(i) = CType(workBook.Worksheets(i), Excel.Worksheet)
                    If SR2 = sheetsArray(i).Name Then
                        AddHandler sheetsArray(i).SelectionChange, AddressOf sheet_SelectionChange2
                        'src_rng = sheetsArray(i).Range("A1", workSheet.Cells(excelApp.Rows.Count, excelApp.Columns.Count))
                    End If
                    'i = i + 1

                Next
                GB_CB_Source2 = worksheet.Range("A1", worksheet.Cells(excelApp.Rows.Count, excelApp.Columns.Count)).Address

            End If
        End If

        If Flag3 = True Then
            'MsgBox(6)

            worksheet.Range("B1").Select()  'Randomly select a cell. If nothing is selected, addhandler show error

            ' Define an array of type Excel.Worksheet
            Dim sheetsArray() As Excel.Worksheet

            ' Resize the array based on the number of sheets
            ReDim sheetsArray(workBook.Worksheets.Count)

            If SR3.Contains("Active Workbook") Then
                'MsgBox(1)
                ' Assuming you're working with the active workbook:
                AddHandler worksheet.SelectionChange, AddressOf sheet_SelectionChange3
                WorkbookEvents = Globals.ThisAddIn.Application.ActiveWorkbook
                CurrentSheet = CType(Globals.ThisAddIn.Application.ActiveSheet, Excel.Worksheet)

                'For i = 1 To excelApp.ActiveWorkbook.Worksheets.Count
                'MsgBox(excelApp.ActiveWorkbook.Worksheets.Count)
                'sheetsArray(i) = CType(workBook.Worksheets(i), Excel.Worksheet)

                ''Needs to add handler to each worksheet
                'If sheetsArray(i).Visible = Excel.XlSheetVisibility.xlSheetVisible Then
                '    AddHandler sheetsArray(i).SelectionChange, AddressOf sheet_SelectionChange
                'End If
                'Next


            ElseIf SR3 = "Select Range" Or SR3.Contains("Active Sheet") Then

                For i = 1 To excelApp.ActiveWorkbook.Worksheets.Count
                    sheetsArray(i) = CType(workBook.Worksheets(i), Excel.Worksheet)
                    If shName2 = sheetsArray(i).Name Then
                        wsEvent3 = DirectCast(sheetsArray(i), Excel.Worksheet)
                        'AddHandler sheetsArray(i).SelectionChange, AddressOf sheet_SelectionChange3
                        'src_rng = sheetsArray(i).Range("A1", workSheet.Cells(excelApp.Rows.Count, excelApp.Columns.Count))
                    End If
                    'i = i + 1

                Next
            Else

                For i = 1 To excelApp.ActiveWorkbook.Worksheets.Count
                    sheetsArray(i) = CType(workBook.Worksheets(i), Excel.Worksheet)
                    If SR3 = sheetsArray(i).Name Then
                        AddHandler sheetsArray(i).SelectionChange, AddressOf sheet_SelectionChange3
                        'src_rng = sheetsArray(i).Range("A1", workSheet.Cells(excelApp.Rows.Count, excelApp.Columns.Count))
                    End If
                    'i = i + 1

                Next
                GB_CB_Source3 = worksheet.Range("A1", worksheet.Cells(excelApp.Rows.Count, excelApp.Columns.Count)).Address

            End If

        End If

        If Flag_CreateDDDL = True Then
            Dim ws1 As Excel.Worksheet
            Dim ws2 As Excel.Worksheet
            For Each ws In excelApp.ActiveWorkbook.Worksheets
                If ws.name = "MySpecialSheet" Then
                    If ws.Range("A1").Value <> "" Then
                        Variable1 = ws.Range("A1").Value.ToString()
                        Variable2 = ws.Range("A2").Value.ToString()
                        Header = ws.Range("A3").Value.ToString()
                        Ascending = ws.Range("A4").Value.ToString()
                        Descending = ws.Range("A5").Value.ToString()
                        TextConvert = ws.Range("A6").Value.ToString()
                        OptionType = ws.Range("A7").Value.ToString()
                        Horizontal_CreateDP = ws.Range("A8").Value.ToString()
                        Flag_CreateDDDL = ws.Range("A9").value.ToString
                        sheetName3 = ws.Range("A10").value.ToString
                        sheetName4 = ws.Range("A11").value.ToString



                        ws1 = CType(workBook.Worksheets(sheetName4), Excel.Worksheet)
                        ws2 = CType(workBook.Worksheets(sheetName3), Excel.Worksheet)
                        src_rng1 = ws2.Range(Variable1)
                        des_rng1 = ws1.Range(Variable2)
                        'AddHandler ws1.Change, AddressOf worksheet5_1_Change


                        wsEvent4_1 = DirectCast(ws1, Excel.Worksheet)
                        'MsgBox(1)
                    End If

                    If ws.Range("B1").Value <> "" Then
                        Variable1 = ws.Range("B1").Value.ToString()
                        Variable2 = ws.Range("B2").Value.ToString()
                        Header = ws.Range("B3").Value.ToString()
                        Ascending = ws.Range("B4").Value.ToString()
                        Descending = ws.Range("B5").Value.ToString()
                        TextConvert = ws.Range("B6").Value.ToString()
                        OptionType = ws.Range("B7").Value.ToString()
                        Horizontal_CreateDP = ws.Range("B8").Value.ToString()
                        Flag_CreateDDDL = ws.Range("B9").value.ToString
                        sheetName3 = ws.Range("B10").value.ToString
                        sheetName4 = ws.Range("B11").value.ToString
                        ws1 = CType(workBook.Worksheets(sheetName4), Excel.Worksheet)
                        'range1 = ws1.Range(Variable2)

                        ws2 = CType(workBook.Worksheets(sheetName3), Excel.Worksheet)
                        'range2 = ws2.Range(Variable1)

                        src_rng2 = ws2.Range(Variable1)
                        des_rng2 = ws1.Range(Variable2)

                        'AddHandler ws1.Change, AddressOf worksheet5_2_Change

                        wsEvent4_2 = DirectCast(ws1, Excel.Worksheet)
                        'MsgBox(2)

                    End If

                    If ws.Range("C1").Value <> "" Then
                        Variable1 = ws.Range("C1").Value.ToString()
                        Variable2 = ws.Range("C2").Value.ToString()
                        Header = ws.Range("C3").Value.ToString()
                        Ascending = ws.Range("C4").Value.ToString()
                        Descending = ws.Range("C5").Value.ToString()
                        TextConvert = ws.Range("C6").Value.ToString()
                        OptionType = ws.Range("C7").Value.ToString()
                        Horizontal_CreateDP = ws.Range("C8").Value.ToString()
                        Flag_CreateDDDL = ws.Range("C9").value.ToString
                        sheetName3 = ws.Range("C10").value.ToString
                        sheetName4 = ws.Range("C11").value.ToString
                        ws1 = CType(workBook.Worksheets(sheetName4), Excel.Worksheet)
                        'range1 = ws1.Range(Variable2)

                        ws2 = CType(workBook.Worksheets(sheetName3), Excel.Worksheet)
                        src_rng3 = ws2.Range(Variable1)
                        des_rng3 = ws1.Range(Variable2)
                        'range2 = ws2.Range(Variable1)
                        'AddHandler ws1.Change, AddressOf worksheet5_2_Change

                        wsEvent4_3 = DirectCast(ws1, Excel.Worksheet)
                        'MsgBox(3)
                    End If

                    If ws.Range("D1").Value <> "" Then
                        Variable1 = ws.Range("D1").Value.ToString()
                        Variable2 = ws.Range("D2").Value.ToString()
                        Header = ws.Range("D3").Value.ToString()
                        Ascending = ws.Range("D4").Value.ToString()
                        Descending = ws.Range("D5").Value.ToString()
                        TextConvert = ws.Range("D6").Value.ToString()
                        OptionType = ws.Range("D7").Value.ToString()
                        Horizontal_CreateDP = ws.Range("D8").Value.ToString()
                        Flag_CreateDDDL = ws.Range("D9").value.ToString
                        sheetName3 = ws.Range("D10").value.ToString
                        sheetName4 = ws.Range("D11").value.ToString

                        ws1 = CType(workBook.Worksheets(sheetName4), Excel.Worksheet)
                        ws2 = CType(workBook.Worksheets(sheetName3), Excel.Worksheet)
                        src_rng4 = ws2.Range(Variable1)
                        des_rng4 = ws1.Range(Variable2)
                        ' AddHandler ws1.Change, AddressOf worksheet5_2_Change

                        wsEvent4_4 = DirectCast(ws1, Excel.Worksheet)
                        ' MsgBox(4)
                    End If

                    If ws.Range("E1").Value <> "" Then
                        Variable1 = ws.Range("E1").Value.ToString()
                        Variable2 = ws.Range("E2").Value.ToString()
                        Header = ws.Range("E3").Value.ToString()
                        Ascending = ws.Range("E4").Value.ToString()
                        Descending = ws.Range("E5").Value.ToString()
                        TextConvert = ws.Range("E6").Value.ToString()
                        OptionType = ws.Range("E7").Value.ToString()
                        Horizontal_CreateDP = ws.Range("E8").Value.ToString()
                        Flag_CreateDDDL = ws.Range("E9").value.ToString
                        sheetName3 = ws.Range("E10").value.ToString
                        sheetName4 = ws.Range("E11").value.ToString

                        ws1 = CType(workBook.Worksheets(sheetName4), Excel.Worksheet)
                        ws2 = CType(workBook.Worksheets(sheetName3), Excel.Worksheet)
                        src_rng5 = ws2.Range(Variable1)
                        des_rng5 = ws1.Range(Variable2)
                        ' AddHandler ws1.Change, AddressOf worksheet5_2_Change
                        wsEvent4_5 = DirectCast(ws1, Excel.Worksheet)
                    End If
                End If
            Next


        End If

        If Flag_Picture = True Then
            ' Define an array of type Excel.Worksheet
            Dim sheetsArray() As Excel.Worksheet

            ' Resize the array based on the number of sheets
            ReDim sheetsArray(workBook.Worksheets.Count)
            For i = 1 To excelApp.ActiveWorkbook.Worksheets.Count
                sheetsArray(i) = CType(workBook.Worksheets(i), Excel.Worksheet)
                If sheetName2 = sheetsArray(i).Name Then
                    'wsEvent5 = DirectCast(sheetsArray(i), Excel.Worksheet)
                    excelApp.Range(Des_Rng_of_PictureDDL).Columns(2).ColumnWidth = excelApp.Range(Src_Rng_of_PictureDDL).Columns(2).ColumnWidth
                    excelApp.Range(Des_Rng_of_PictureDDL).Rows.RowHeight = excelApp.Range(Src_Rng_of_PictureDDL).RowHeight

                    'worksheet7_Change(Target)
                    'worksheet6_Change(Target)
                    'AddHandler sheetsArray(i).Change, AddressOf worksheet7_Change
                    AddHandler sheetsArray(i).Change, AddressOf worksheet6_Change
                    'MsgBox(sheetsArray(i).Name)
                    'worksheet.Change
                    'src_rng = sheetsArray(i).Range("A1", workSheet.Cells(excelApp.Rows.Count, excelApp.Columns.Count))
                End If
                'i = i + 1

            Next
            'wsEvent4 = DirectCast(worksheet, Excel.Worksheet)
            'AddHandler worksheet.Change, AddressOf worksheet6_Change
        End If



    End Sub

    ' This event will trigger when a cell in the worksheet is selected.
    Private Sub wsEvent1_SelectionChange(ByVal Target As Excel.Range) Handles wsEvent1.SelectionChange
        ' For testing purposes, we'll just show a message box.
        'MsgBox("Cell selected: " & Target.Address)
        sheet_SelectionChange1(Target)
    End Sub

    Private Sub wsEvent2_SelectionChange(ByVal Target As Excel.Range) Handles wsEvent2.SelectionChange
        ' For testing purposes, we'll just show a message box.
        'MsgBox("Cell selected: " & Target.Address)
        sheet_SelectionChange2(Target)
    End Sub

    Private Sub wsEvent3_SelectionChange(ByVal Target As Excel.Range) Handles wsEvent3.SelectionChange
        ' For testing purposes, we'll just show a message box.
        'MsgBox("Cell selected: " & Target.Address)
        sheet_SelectionChange3(Target)
    End Sub

    'For Create Dynamic List
    Private Sub wsEvent4_SelectionChange(ByVal Target As Excel.Range) Handles wsEvent4_1.Change
        Try
            excelApp = Globals.ThisAddIn.Application
            Dim workBook As Excel.Workbook = excelApp.ActiveWorkbook
            'Dim worksheet As Excel.Worksheet = workBook.ActiveSheet

            Dim targetsheet As Excel.Worksheet
            For Each ws In excelApp.ActiveWorkbook.Worksheets
                If ws.name = "MySpecialSheet" Then
                    targetsheet = ws
                End If
            Next

            Header = targetsheet.Range("A3").Value.ToString()
            Ascending = targetsheet.Range("A4").Value.ToString()
            Descending = targetsheet.Range("A5").Value.ToString()
            TextConvert = targetsheet.Range("A6").Value.ToString()
            OptionType = targetsheet.Range("A7").Value.ToString()
            Horizontal_CreateDP = targetsheet.Range("A8").Value.ToString()
            sheetName10 = targetsheet.Range("A10").Value.ToString()
            sheetName11 = targetsheet.Range("A11").Value.ToString()

            src_rng = src_rng1
            des_rng = des_rng1
            'If excelApp.Intersect(Target, des_rng) IsNot Nothing Then
            'MsgBox(1)
            worksheet5_2_Change(Target)
            'End If
        Catch ex As Exception
        End Try

    End Sub

    Private Sub wsEvent4_2_SelectionChange(ByVal Target As Excel.Range) Handles wsEvent4_2.Change

        excelApp = Globals.ThisAddIn.Application
        Dim workBook As Excel.Workbook = excelApp.ActiveWorkbook
        'Dim worksheet As Excel.Worksheet = workBook.ActiveSheet

        Dim targetsheet As Excel.Worksheet
        For Each ws In excelApp.ActiveWorkbook.Worksheets
            If ws.name = "MySpecialSheet" Then
                targetsheet = ws
            End If
        Next

        Header = targetsheet.Range("B3").Value.ToString()
        Ascending = targetsheet.Range("B4").Value.ToString()
        Descending = targetsheet.Range("B5").Value.ToString()
        TextConvert = targetsheet.Range("B6").Value.ToString()
        OptionType = targetsheet.Range("B7").Value.ToString()
        Horizontal_CreateDP = targetsheet.Range("B8").Value.ToString()
        sheetName10 = targetsheet.Range("B10").Value.ToString()
        sheetName11 = targetsheet.Range("B11").Value.ToString()

        src_rng = src_rng2
        des_rng = des_rng2
        'MsgBox(src_rng.Worksheet.Name)
        'MsgBox(des_rng.Worksheet.Name)

        'If excelApp.Intersect(Target, des_rng) IsNot Nothing Then
        worksheet5_2_Change(Target)
        'End If
    End Sub

    Private Sub wsEvent4_3_SelectionChange(ByVal Target As Excel.Range) Handles wsEvent4_3.Change

        excelApp = Globals.ThisAddIn.Application
        Dim workBook As Excel.Workbook = excelApp.ActiveWorkbook
        'Dim worksheet As Excel.Worksheet = workBook.ActiveSheet

        Dim targetsheet As Excel.Worksheet
        For Each ws In excelApp.ActiveWorkbook.Worksheets
            If ws.name = "MySpecialSheet" Then
                targetsheet = ws
            End If
        Next

        Header = targetsheet.Range("C3").Value.ToString()
        Ascending = targetsheet.Range("C4").Value.ToString()
        Descending = targetsheet.Range("C5").Value.ToString()
        TextConvert = targetsheet.Range("C6").Value.ToString()
        OptionType = targetsheet.Range("C7").Value.ToString()
        Horizontal_CreateDP = targetsheet.Range("C8").Value.ToString()
        sheetName10 = targetsheet.Range("C10").Value.ToString()
        sheetName11 = targetsheet.Range("C11").Value.ToString()

        src_rng = src_rng3
        des_rng = des_rng3
        ' MsgBox(des_rng.Address)
        'MsgBox(src_rng.Worksheet.Name)
        'MsgBox(des_rng.Worksheet.Name)

        ' If excelApp.Intersect(Target, des_rng) IsNot Nothing Then
        worksheet5_2_Change(Target)
        ' End If
    End Sub

    Private Sub wsEvent4_4_SelectionChange(ByVal Target As Excel.Range) Handles wsEvent4_4.Change

        excelApp = Globals.ThisAddIn.Application
        Dim workBook As Excel.Workbook = excelApp.ActiveWorkbook
        'Dim worksheet As Excel.Worksheet = workBook.ActiveSheet

        Dim targetsheet As Excel.Worksheet
        For Each ws In excelApp.ActiveWorkbook.Worksheets
            If ws.name = "MySpecialSheet" Then
                targetsheet = ws
            End If
        Next

        Header = targetsheet.Range("D3").Value.ToString()
        Ascending = targetsheet.Range("D4").Value.ToString()
        Descending = targetsheet.Range("D5").Value.ToString()
        TextConvert = targetsheet.Range("D6").Value.ToString()
        OptionType = targetsheet.Range("D7").Value.ToString()
        Horizontal_CreateDP = targetsheet.Range("D8").Value.ToString()
        sheetName10 = targetsheet.Range("D10").Value.ToString()
        sheetName11 = targetsheet.Range("D11").Value.ToString()

        src_rng = src_rng4
        des_rng = des_rng4
        'MsgBox(src_rng.Worksheet.Name)
        'MsgBox(des_rng.Worksheet.Name)

        ' If excelApp.Intersect(Target, des_rng) IsNot Nothing Then
        worksheet5_2_Change(Target)
        '  End If

    End Sub

    Private Sub wsEvent4_5_SelectionChange(ByVal Target As Excel.Range) Handles wsEvent4_5.Change

        excelApp = Globals.ThisAddIn.Application
        Dim workBook As Excel.Workbook = excelApp.ActiveWorkbook
        'Dim worksheet As Excel.Worksheet = workBook.ActiveSheet

        Dim targetsheet As Excel.Worksheet
        For Each ws In excelApp.ActiveWorkbook.Worksheets
            If ws.name = "MySpecialSheet" Then
                targetsheet = ws
            End If
        Next

        Header = targetsheet.Range("E3").Value.ToString()
        Ascending = targetsheet.Range("E4").Value.ToString()
        Descending = targetsheet.Range("E5").Value.ToString()
        TextConvert = targetsheet.Range("E6").Value.ToString()
        OptionType = targetsheet.Range("E7").Value.ToString()
        Horizontal_CreateDP = targetsheet.Range("E8").Value.ToString()
        sheetName10 = targetsheet.Range("E10").Value.ToString()
        sheetName11 = targetsheet.Range("E11").Value.ToString()

        src_rng = src_rng5
        des_rng = des_rng5
        'MsgBox(src_rng.Worksheet.Name)
        'MsgBox(des_rng.Worksheet.Name)

        '  If excelApp.Intersect(Target, des_rng) IsNot Nothing Then
        worksheet5_2_Change(Target)
        '  End If

    End Sub



    'For Picture Drop-down List
    Private Sub wsEvent5_SelectionChange(ByVal Target As Excel.Range) Handles wsEvent5.SelectionChange

        excelApp.Range(Des_Rng_of_PictureDDL).Columns(2).ColumnWidth = excelApp.Range(Src_Rng_of_PictureDDL).Columns(2).ColumnWidth
        excelApp.Range(Des_Rng_of_PictureDDL).Rows.RowHeight = excelApp.Range(Src_Rng_of_PictureDDL).RowHeight

        ' worksheet7_Change(Target)
        worksheet6_Change(Target)

    End Sub


    ' Event handler for when any sheet in the workbook is activated
    Private Sub WorkbookEvents_SheetActivate(ByVal Sh As Object) Handles WorkbookEvents.SheetActivate
        ' Detach event from previous sheet
        If CurrentSheet IsNot Nothing Then

            RemoveHandler CurrentSheet.SelectionChange, AddressOf sheet_SelectionChange1
        End If

        ' Attach event to the new active sheet
        CurrentSheet = CType(Sh, Excel.Worksheet)
        AddHandler CurrentSheet.SelectionChange, AddressOf sheet_SelectionChange1
        ' MsgBox(CurrentSheet.Name)
    End Sub

    Private Sub RemovePreviousEventHandler()

        ' This function ensures that we remove previously attached event handlers
        ' to avoid multiple event triggers. 
        ' This is a simplistic approach and may need refining based on your exact needs.
        For Each wb In excelApp.Workbooks
            For Each worksheet In wb.Worksheets
                RemoveHandler CType(worksheet, Excel.Worksheet).SelectionChange, AddressOf sheet_SelectionChange1
                RemoveHandler CType(worksheet, Excel.Worksheet).SelectionChange, AddressOf sheet_SelectionChange2
                RemoveHandler CType(worksheet, Excel.Worksheet).SelectionChange, AddressOf sheet_SelectionChange3
            Next
        Next
    End Sub

    Private Sub CheckForNewwWorksheet()

        excelApp = Globals.ThisAddIn.Application
        Dim workBook As Excel.Workbook = excelApp.ActiveWorkbook
        Dim worksheet As Excel.Worksheet = workBook.ActiveSheet
        'excelApp = Globals.ThisAddIn.Application

        ' Loop through each worksheet in the active workbook
        For Each ws In excelApp.ActiveWorkbook.Worksheets
            If ws.Name = "Newwwwwwwwww" Then
                ' If worksheet "Neww" is found, store the value of A1 in the variable
                'MsgBox(1)
                GB_CB_Source1 = ws.Range("A2").Value.ToString()
                ' MsgBox(GB_CB_Source)
                SR1 = ws.Range("A3").Value.ToString()
                Horizontal1 = ws.Range("A4").Value.ToString()
                Separator1 = ws.Range("A5").Value.ToString()
                Search1 = ws.Range("A6").Value.ToString()
                Flag1 = ws.Range("A7").Value.ToString()
                shName1 = ws.Range("A9").Value.ToString
                RangeType1 = ws.Range("B2").value.ToString
                ' TargetVar = ws.Range("A8").Value.ToString

                Dim src_rng As Excel.Range = worksheet.Range(GB_CB_Source1)
                TType1 = ""
                'Exit Sub ' No need to check other sheets once "Neww" is found
                'MsgBox(1)
            End If
            If ws.Name = "SoftekoSofteko" Then   'For checkbox drop-down list
                'MsgBox(2)

                ' If worksheet "Neww" is found, store the value of A1 in the variable
                GB_CB_Source2 = ws.Range("A2").Value.ToString()
                ' MsgBox(GB_CB_Source)
                SR2 = ws.Range("A3").Value.ToString()
                Horizontal2 = ws.Range("A4").Value.ToString()
                Separator2 = ws.Range("A5").Value.ToString()
                Search2 = ws.Range("A6").Value.ToString()
                Flag2 = ws.Range("A7").Value.ToString()
                shName2 = ws.Range("A9").Value.ToString
                RangeType2 = ws.Range("B2").value.ToString
                ' TargetVar = ws.Range("A8").Value.ToString

                Dim src_rng As Excel.Range = worksheet.Range(GB_CB_Source2)
                TType2 = ""
                'MsgBox(2)
            End If
            If ws.Name = "SoftekoSoftekoSofteko" Then   'For checkbox drop-down list
                'MsgBox(3)
                ' If worksheet "Neww" is found, store the value of A1 in the variable
                GB_CB_Source3 = ws.Range("A2").Value.ToString()
                ' MsgBox(GB_CB_Source)
                SR3 = ws.Range("A3").Value.ToString()
                'Horizontal2 = ws.Range("A4").Value.ToString()
                'Separator2 = ws.Range("A5").Value.ToString()
                'Search2 = ws.Range("A6").Value.ToString()
                Flag3 = ws.Range("A7").Value.ToString()
                shName3 = ws.Range("A9").Value.ToString
                RangeType3 = ws.Range("B2").value.ToString
                ' TargetVar = ws.Range("A8").Value.ToString

                Dim src_rng As Excel.Range = worksheet.Range(GB_CB_Source3)
                TType3 = ""
                'MsgBox(3)

            End If

            If ws.name = "MySpecialSheet" Then
                Variable1 = ws.Range("A1").Value.ToString()
                Variable2 = ws.Range("A2").Value.ToString()
                Header = ws.Range("A3").Value.ToString()
                Ascending = ws.Range("A4").Value.ToString()
                Descending = ws.Range("A5").Value.ToString()
                TextConvert = ws.Range("A6").Value.ToString()
                OptionType = ws.Range("A7").Value.ToString()
                Horizontal_CreateDP = ws.Range("A8").Value.ToString()
                Flag_CreateDDDL = ws.Range("A9").value.ToString
                sheetName3 = ws.Range("A10").value.ToString
                sheetName4 = ws.Range("A11").value.ToString

            End If

            If ws.name = "SoftekoPictureBasedDropDown" Then
                Flag_Picture = ws.Range("A2").value.ToString
                sheetName2 = ws.Range("A3").value.ToString
                Src_Rng_of_PictureDDL = ws.Range("A4").value.ToString
                Des_Rng_of_PictureDDL = ws.Range("A5").value.ToString
                'MsgBox(1)
            End If
        Next
    End Sub

    Private Sub HideNewwwWorksheet()
        Dim ws As Excel.Worksheet

        ' Loop through each worksheet in the active workbook
        For Each ws In Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets
            If ws.Name = "Newwwwwwwwww" Then
                'MsgBox(4)
                ' If worksheet "Newww" is found, hide it
                ws.Visible = Excel.XlSheetVisibility.xlSheetHidden
                'Exit Sub ' No need to check other sheets once "Newww" is found
            End If

            If ws.Name = "SoftekoSofteko" Then
                ' If worksheet "SoftekoSofteko" is found, hide it
                'MsgBox(5)
                ws.Visible = Excel.XlSheetVisibility.xlSheetHidden
                'Exit Sub ' No need to check other sheets once "Newww" is found
            End If

            If ws.Name = "SoftekoSoftekoSofteko" Then
                ' If worksheet "SoftekoSofteko" is found, hide it
                ws.Visible = Excel.XlSheetVisibility.xlSheetHidden
                'Exit Sub ' No need to check other sheets once "Newww" is found
            End If
            If ws.Name = "SoftekoPictureBasedDropDown" Then
                ws.Visible = Excel.XlSheetVisibility.xlSheetHidden
            End If
        Next
    End Sub

    'For multi drop-down list
    Private Sub sheet_SelectionChange1(ByVal Target As Excel.Range)

        excelApp = Globals.ThisAddIn.Application
        workBook = excelApp.ActiveWorkbook
        worksheet = workBook.ActiveSheet

        Try


            Dim src_rng_concate As Excel.Range
            'MsgBox(1)

            If GB_CB_Source1 IsNot Nothing Then
                Dim src_rng As Excel.Range = worksheet.Range(GB_CB_Source1)
                'MsgBox(src_rng.Address)

                'MsgBox(src_rng.Address)
                'src_rng = workSheet.Range(GB_CB_Source1)


                If SR1.Contains("Active Workbook") Then
                    src_rng = worksheet.Range("A1", worksheet.Cells(excelApp.Rows.Count, excelApp.Columns.Count))
                End If
                src_rng = workBook.ActiveSheet.range(src_rng.Address)
                ' MsgBox(src_rng.Worksheet.Name)

                'Change starts from here
                If (Nam1 = worksheet.Name And TType1 = "Select Range") Or (Nam1 = worksheet.Name And TType1.Contains("Active Sheet")) Or (Nam1 = worksheet.Name And TType1 = worksheet.Name) Then

                    src_rng_concate = worksheet.Range(GB_CB_Dlt1)
                    'MsgBox(src_rng_concate.Address)

                    If IsCellInsideRange(Target, src_rng) = True And Target.Cells.Count = 1 And HasDataValidationList(Target) And IsCellInsideRange(Target, src_rng_concate) = True Then
                        'MsgBox(1)
                        'If Target.Cells.Count = 1 Then ' Ensure only one cell is selected
                        TargetVar1 = Target.Address
                        If Form IsNot Nothing Then
                            'Form = Nothing

                            Form.Dispose()
                            'MsgBox(2)
                        End If


                    Else


                        If IsCellInsideRange(Target, src_rng) And Target.Cells.Count = 1 And HasDataValidationList(Target) Then
                            'MsgBox(2)
                            'If Target.Cells.Count = 1 Then ' Ensure only one cell is selected
                            TargetVar1 = Target.Address
                            If Form Is Nothing OrElse Form.IsDisposed Then
                                Form = New Form36()
                                Form.Show()
                                Form.BringToFront()
                                Form.Refresh()
                            Else
                                ' If form is already open, bring it to the front

                                Form.Dispose()
                                Form = New Form36()
                                Form.Show()
                                Form.BringToFront()

                            End If
                        End If

                        'Dim form As New Form36()
                        'form.Show()
                        'form.Focus()
                        ''form.TopMost = True
                        ''form.Activate()
                        'form.BringToFront()
                        'End If
                    End If
                Else

                    If IsCellInsideRange(Target, src_rng) And Target.Cells.Count = 1 And HasDataValidationList(Target) Then
                        'MsgBox(2)
                        'MsgBox(10)
                        'If Target.Cells.Count = 1 Then ' Ensure only one cell is selected
                        TargetVar1 = Target.Address
                        If Form Is Nothing OrElse Form.IsDisposed Then
                            Form = New Form36()
                            Form.Show()
                            Form.BringToFront()
                            Form.Refresh()
                        Else
                            ' If form is already open, bring it to the front

                            Form.Dispose()
                            Form = New Form36()
                            Form.Show()
                            Form.BringToFront()
                            'MsgBox(3)

                        End If
                    End If
                End If
            Else

                'If Form IsNot Nothing Then
                '    'Form = Nothing

                '    Form.Dispose()
                '    'MsgBox(2)
                'End If


            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)

        End Try
    End Sub

    'For CheckBox drop-down List
    Private Sub sheet_SelectionChange2(ByVal Target As Excel.Range)
        excelApp = Globals.ThisAddIn.Application
        workBook = excelApp.ActiveWorkbook
        worksheet = workBook.ActiveSheet
        Try



            Dim src_rng_concate As Excel.Range
            If GB_CB_Source2 IsNot Nothing Then
                Dim src_rng As Excel.Range = worksheet.Range(GB_CB_Source2)

                If SR2.Contains("Active Workbook") Then
                    src_rng = worksheet.Range("A1", worksheet.Cells(excelApp.Rows.Count, excelApp.Columns.Count))
                End If
                src_rng = workBook.ActiveSheet.range(src_rng.Address)
                ' MsgBox(src_rng.Worksheet.Name)

                'Change starts from here
                If (Nam2 = worksheet.Name And TType2 = "Select Range") Or (Nam2 = worksheet.Name And TType2.Contains("Active Sheet")) Or (Nam2 = worksheet.Name And TType2 = worksheet.Name) Then

                    src_rng_concate = worksheet.Range(GB_CB_Dlt2)
                    'MsgBox(src_rng_concate.Address)

                    If IsCellInsideRange(Target, src_rng) = True And Target.Cells.Count = 1 And HasDataValidationList(Target) And IsCellInsideRange(Target, src_rng_concate) = True Then
                        'MsgBox(1)
                        'If Target.Cells.Count = 1 Then ' Ensure only one cell is selected
                        TargetVar2 = Target.Address
                        If Form IsNot Nothing Then
                            'Form = Nothing

                            Form.Dispose()
                            'MsgBox(2)
                        End If





                    Else


                        If IsCellInsideRange(Target, src_rng) And Target.Cells.Count = 1 And HasDataValidationList(Target) Then
                            'MsgBox(2)
                            'If Target.Cells.Count = 1 Then ' Ensure only one cell is selected
                            TargetVar2 = Target.Address
                            If Form2 Is Nothing OrElse Form.IsDisposed Then
                                Form2 = New Form38()
                                Form2.Show()
                                Form2.BringToFront()
                                Form2.Refresh()
                            Else
                                ' If form is already open, bring it to the front

                                Form2.Dispose()
                                Form2 = New Form38()
                                Form2.Show()
                                Form2.BringToFront()

                            End If
                        End If

                        'Dim form As New Form36()
                        'form.Show()
                        'form.Focus()
                        ''form.TopMost = True
                        ''form.Activate()
                        'form.BringToFront()
                        'End If
                    End If
                Else

                    If IsCellInsideRange(Target, src_rng) And Target.Cells.Count = 1 And HasDataValidationList(Target) Then
                        'MsgBox(2)
                        'MsgBox(10)
                        'If Target.Cells.Count = 1 Then ' Ensure only one cell is selected
                        TargetVar2 = Target.Address
                        If Form2 Is Nothing OrElse Form2.IsDisposed Then
                            Form2 = New Form38()
                            Form2.Show()
                            Form2.BringToFront()
                            Form2.Refresh()
                        Else
                            ' If form is already open, bring it to the front

                            Form2.Dispose()
                            Form2 = New Form38()
                            Form2.Show()
                            Form2.BringToFront()
                            'MsgBox(3)

                        End If
                    End If
                End If
            Else

                'If Form IsNot Nothing Then
                '    'Form = Nothing

                '    Form.Dispose()
                '    'MsgBox(2)
                'End If


            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)

        End Try
    End Sub

    'For search drop-down
    Private Sub sheet_SelectionChange3(ByVal Target As Excel.Range)
        excelApp = Globals.ThisAddIn.Application
        workBook = excelApp.ActiveWorkbook
        worksheet = workBook.ActiveSheet

        Dim src_rng_concate As Excel.Range
        'MsgBox(workSheet.Name)
        'MsgBox(src_rng.Worksheet.Name)

        'src_rng = workSheet.Range(GB_CB_Source1)

        If GB_CB_Source3 IsNot Nothing Then
            Dim src_rng As Excel.Range = worksheet.Range(GB_CB_Source3)
            'MsgBox(src_rng.Address)

            'MsgBox(src_rng.Address)
            'src_rng = workSheet.Range(GB_CB_Source1)


            If SR3.Contains("Active Workbook") Then
                src_rng = worksheet.Range("A1", worksheet.Cells(excelApp.Rows.Count, excelApp.Columns.Count))
            End If
            src_rng = workBook.ActiveSheet.range(src_rng.Address)
            ' MsgBox(src_rng.Worksheet.Name)

            'Change starts from here
            If (Nam3 = worksheet.Name And TType3 = "Select Range") Or (Nam3 = worksheet.Name And TType3.Contains("Active Sheet")) Or (Nam3 = worksheet.Name And TType3 = worksheet.Name) Then

                src_rng_concate = worksheet.Range(GB_CB_Dlt3)
                'MsgBox(src_rng_concate.Address)

                If IsCellInsideRange(Target, src_rng) = True And Target.Cells.Count = 1 And HasDataValidationList(Target) And IsCellInsideRange(Target, src_rng_concate) = True Then
                    'MsgBox(1)
                    'If Target.Cells.Count = 1 Then ' Ensure only one cell is selected
                    TargetVar3 = Target.Address
                    If Form3 IsNot Nothing Then
                        'Form = Nothing

                        Form3.Dispose()
                        'MsgBox(2)
                    End If


                Else


                    If IsCellInsideRange(Target, src_rng) And Target.Cells.Count = 1 And HasDataValidationList(Target) Then
                        'MsgBox(2)
                        'If Target.Cells.Count = 1 Then ' Ensure only one cell is selected
                        TargetVar3 = Target.Address
                        If Form3 Is Nothing OrElse Form.IsDisposed Then
                            Form3 = New Form40()
                            Form3.Show()
                            Form3.BringToFront()
                            Form3.Refresh()
                        Else
                            ' If form is already open, bring it to the front

                            Form3.Dispose()
                            Form3 = New Form40()
                            Form3.Show()
                            Form3.BringToFront()

                        End If
                    End If

                    'Dim form As New Form36()
                    'form.Show()
                    'form.Focus()
                    ''form.TopMost = True
                    ''form.Activate()
                    'form.BringToFront()
                    'End If
                End If
            Else

                If IsCellInsideRange(Target, src_rng) And Target.Cells.Count = 1 And HasDataValidationList(Target) Then
                    'MsgBox(2)
                    'MsgBox(10)
                    'If Target.Cells.Count = 1 Then ' Ensure only one cell is selected
                    TargetVar3 = Target.Address
                    If Form3 Is Nothing OrElse Form.IsDisposed Then
                        Form3 = New Form40()
                        Form3.Show()
                        Form3.BringToFront()
                        Form3.Refresh()
                    Else
                        ' If form is already open, bring it to the front

                        Form3.Dispose()
                        Form3 = New Form40()
                        Form3.Show()
                        Form3.BringToFront()
                        'MsgBox(3)

                    End If
                End If
            End If
        Else

            'If Form IsNot Nothing Then
            '    'Form = Nothing

            '    Form.Dispose()
            '    'MsgBox(2)
            'End If


        End If
    End Sub
    Private Function IsCellInsideRange(ByVal cell As Excel.Range, ByVal targetRange As Excel.Range) As Boolean
        'MsgBox(cell.Address)
        'MsgBox(targetRange.Address)
        Try
            Dim intersectRange As Excel.Range = Globals.ThisAddIn.Application.Intersect(cell, targetRange)
            'MsgBox(intersectRange.Address)
            Return Not intersectRange Is Nothing
        Catch ex As Exception
            'MsgBox(cell.Address)
            'MsgBox(targetRange.Address)
            Return False
        End Try
    End Function

    Private Function HasDataValidationList(ByVal cell As Excel.Range) As Boolean
        Dim hasValidation As Boolean = False

        Try
            If Not cell.Validation Is Nothing AndAlso cell.Validation.Type = Excel.XlDVType.xlValidateList Then
                hasValidation = True
            End If
        Catch ex As Exception
            ' Exception will be thrown if cell doesn't have validation. No action needed.
        End Try

        Return hasValidation
    End Function


    'Create Dynamic Drop-down list
    Public Sub worksheet5_1_Change(ByVal Target As Excel.Range)
        excelApp = Globals.ThisAddIn.Application
        workBook = excelApp.ActiveWorkbook
        worksheet = workBook.ActiveSheet

        'src_rng = excelApp.Range(Variable1)
        'des_rng = excelApp.Range(Variable2)
        'Dim src_sheet As Excel.Worksheet = CType(workBook.Worksheets(sheetName3), Excel.Worksheet)
        'Dim des_sheet As Excel.Worksheet = CType(workBook.Worksheets(sheetName4), Excel.Worksheet)

        ' src_rng = src_sheet.Range(src_rng.Address)

        'des_rng = des_sheet.Range(des_rng.Address)
        'MsgBox(1)
        'MsgBox(des_rng.Address)
        Dim rng As Excel.Range
        'des_rng.ClearContents()

        If OptionType = True Then
            If Header = True Then
                'Dim adjustRange As Excel.Range
                rng = src_rng.Offset(1, 0).Resize(src_rng.Rows.Count - 1, src_rng.Columns.Count)

            Else

                rng = src_rng 'Assuming you have a range from A1 to A100
            End If
            ' MsgBox(des_rng.Rows.Count)
            Dim col_dif As Integer
            col_dif = Target.Column - worksheet.Range(des_rng.Address).Column + 1
            'MsgBox(col_dif)

            'For k = 1 To des_rng.Rows.Count
            Dim matchedValues As New List(Of String)
            Dim sec_matchedValues As New List(Of String)
            Dim thrd_matchedValues As New List(Of String)
            Dim four_matchedValues As New List(Of String)
            Dim k As Integer = Target.Row - worksheet.Range(des_rng.Address).Row + 1
            'MsgBox(k)
            'MsgBox(i)
            If col_dif = 1 Then

                If des_rng(k, 1).Value IsNot Nothing Then
                    For i = 1 To rng.Rows.Count
                        If rng(i, 1).Value = des_rng(k, 1).Value Then
                            If Not matchedValues.Contains(rng(i, 2).Value) Then
                                matchedValues.Add(rng(i, 2).Value)
                            End If
                            'matchedValues.Add(rng(i, 2).Value)
                        End If
                    Next


                    If Ascending = True Then
                        'Sort the list in ascending order
                        matchedValues.Sort()
                    ElseIf Descending = True Then
                        'Sort the list in ascending order
                        matchedValues.Sort()
                        matchedValues.Reverse()
                    End If

                    'Dim dropDownRange As Excel.Range = des_rng(k, 2)
                    Dim dropDownRange As Excel.Range = Target(1, 2)
                    '  MsgBox(Target.Address)
                    Dim Validation As Excel.Validation = dropDownRange.Validation
                    Validation.Delete() 'Remove any existing validation
                    Validation.Add(Excel.XlDVType.xlValidateList, Formula1:=String.Join(",", matchedValues))
                    matchedValues.Clear()
                    'MsgBox(k)
                End If

                '  Dim sec_matchedValues As New List(Of String)
            ElseIf col_dif = 2 Then
                If des_rng(k, 2).Value IsNot Nothing Then
                    For i = 1 To rng.Rows.Count
                        If rng(i, 1).Value = des_rng(k, 1).Value And rng(i, 2).Value = des_rng(k, 2).Value Then
                            If Not sec_matchedValues.Contains(rng(i, 3).Value) Then
                                sec_matchedValues.Add(rng(i, 3).Value)
                            End If

                        End If
                    Next


                    If Ascending = True Then
                        'Sort the list in ascending order
                        sec_matchedValues.Sort()
                    ElseIf Descending = True Then
                        'Sort the list in ascending order
                        sec_matchedValues.Sort()
                        sec_matchedValues.Reverse()
                    End If


                    ' Dim dropDownRange As Excel.Range = des_rng(k, 3)
                    Dim dropDownRange As Excel.Range = Target(, 2)
                    Dim Validation As Excel.Validation = dropDownRange.Validation
                    Validation.Delete() 'Remove any existing validation
                    Validation.Add(Excel.XlDVType.xlValidateList, Formula1:=String.Join(",", sec_matchedValues))
                    sec_matchedValues.Clear()
                End If
            ElseIf col_dif = 3 Then
                '       Dim thrd_matchedValues As New List(Of String)

                If des_rng(k, 3).Value IsNot Nothing Then
                    For i = 1 To rng.Rows.Count
                        If rng(i, 1).Value = des_rng(k, 1).Value And rng(i, 2).Value = des_rng(k, 2).Value And rng(i, 3).Value = des_rng(k, 3).Value Then
                            If Not thrd_matchedValues.Contains(rng(i, 4).Value) Then
                                thrd_matchedValues.Add(rng(i, 4).Value)
                            End If

                        End If
                    Next


                    If Ascending = True Then
                        'Sort the list in ascending order
                        thrd_matchedValues.Sort()
                    ElseIf Descending = True Then
                        'Sort the list in ascending order
                        thrd_matchedValues.Sort()
                        thrd_matchedValues.Reverse()
                    End If


                    'Dim dropDownRange As Excel.Range = des_rng(k, 4)
                    Dim dropDownRange As Excel.Range = Target(, 2)
                    Dim Validation As Excel.Validation = dropDownRange.Validation
                    Validation.Delete() 'Remove any existing validation
                    Validation.Add(Excel.XlDVType.xlValidateList, Formula1:=String.Join(",", thrd_matchedValues))
                    thrd_matchedValues.Clear()
                End If


                '  Dim four_matchedValues As New List(Of String)
            ElseIf col_dif = 4 Then
                If des_rng(k, 4).Value IsNot Nothing Then
                    For i = 1 To rng.Rows.Count
                        If rng(i, 1).Value = des_rng(k, 1).Value And rng(i, 2).Value = des_rng(k, 2).Value And rng(i, 3).Value = des_rng(k, 3).Value And rng(i, 4).Value = des_rng(k, 4).Value Then

                            If Not four_matchedValues.Contains(rng(i, 5).Value) Then
                                four_matchedValues.Add(rng(i, 5).Value)
                            End If


                        End If
                    Next


                    If Ascending = True Then
                        'Sort the list in ascending order
                        four_matchedValues.Sort()
                    ElseIf Descending = True Then
                        'Sort the list in ascending order
                        four_matchedValues.Sort()
                        four_matchedValues.Reverse()
                    End If


                    Dim dropDownRange As Excel.Range = des_rng(k, 5)
                    Dim Validation As Excel.Validation = dropDownRange.Validation
                    Validation.Delete() 'Remove any existing validation
                    Validation.Add(Excel.XlDVType.xlValidateList, Formula1:=String.Join(",", four_matchedValues))
                    four_matchedValues.Clear()
                End If
            End If

            'Next

        ElseIf OptionType = False Then
            If Horizontal_CreateDP = True Then
                If Target.Address = des_rng(1, 1).Address Then

                    Dim worksheet As Excel.Worksheet = CType(Target.Worksheet, Excel.Worksheet)
                    Dim col As Integer = src_rng.Rows().Find(Target.Value).Column - src_rng.Column + 1
                    'MsgBox(col)
                    'Dim ab As Integer = col - src_rng.Column
                    Dim sourceRng As Excel.Range = src_rng.Cells(2, col).Resize(src_rng(src_rng.Rows.Count, col).row - 2, 1)
                    'MsgBox(sourceRng.Address)
                    'Dim sourceRng As Excel.Range = src_rng.Cells(2, col).Resize(worksheet.Cells(worksheet.Rows.Count, col), 1)
                    Dim dropDownRange As Excel.Range = des_rng(1, 2)
                    Dim Validation As Excel.Validation = dropDownRange.Validation
                    Validation.Delete() 'Remove any existing validation
                    Dim formula As String = "='" & sourceRng.Worksheet.Name & "'!" & sourceRng.Address(External:=False)
                    Validation.Add(Excel.XlDVType.xlValidateList, Formula1:=formula)
                    'CreateValidationList(worksheet.Cells(2, 5), "=" & sourceRng.Address)
                End If

            ElseIf Horizontal_CreateDP = False Then
                If Target.Address = des_rng(1, 1).Address Then
                    Dim worksheet As Excel.Worksheet = CType(Target.Worksheet, Excel.Worksheet)
                    Dim col As Integer = src_rng.Rows().Find(Target.Value).Column - src_rng.Column + 1
                    'MsgBox(col)
                    'Dim ab As Integer = col - src_rng.Column
                    Dim sourceRng As Excel.Range = src_rng.Cells(2, col).Resize(src_rng(src_rng.Rows.Count, col).row - 2, 1)
                    'MsgBox(sourceRng.Address)
                    'Dim sourceRng As Excel.Range = src_rng.Cells(2, col).Resize(worksheet.Cells(worksheet.Rows.Count, col), 1)
                    Dim dropDownRange As Excel.Range = des_rng(2, 1)
                    Dim Validation As Excel.Validation = dropDownRange.Validation
                    Validation.Delete() 'Remove any existing validation
                    Dim formula As String = "='" & sourceRng.Worksheet.Name & "'!" & sourceRng.Address(External:=False)
                    Validation.Add(Excel.XlDVType.xlValidateList, Formula1:=formula)
                End If
            End If

        End If
        ' Catch ex As Exception
        'MsgBox("error")
        'End Try


    End Sub


    Public Sub worksheet5_2_Change(ByVal Target As Excel.Range)
        excelApp = Globals.ThisAddIn.Application
        workBook = excelApp.ActiveWorkbook
        worksheet = workBook.ActiveSheet

        Dim targetWorksheet As Excel.Worksheet
        Dim i As Integer = 1
        For Each ws In excelApp.ActiveWorkbook.Worksheets
            If ws.name = "MySpecialSheet" Then
                targetWorksheet = ws
                Exit For
            End If
        Next


        'For i = 1 To targetWorksheet.Columns.Count
        If Target.Worksheet.Name = targetWorksheet.Range("A11").Value And excelApp.Intersect(Target, excelApp.Range(targetWorksheet.Range("A2").Value)) IsNot Nothing Then
            Variable1 = targetWorksheet.Range("A1").Value.ToString()
            Variable2 = targetWorksheet.Range("A2").Value.ToString()
            Header = targetWorksheet.Range("A3").Value.ToString()
            Ascending = targetWorksheet.Range("A4").Value.ToString()
            Descending = targetWorksheet.Range("A5").Value.ToString()
            TextConvert = targetWorksheet.Range("A6").Value.ToString()
            OptionType = targetWorksheet.Range("A7").Value.ToString()
            Horizontal_CreateDP = targetWorksheet.Range("A8").Value.ToString()
            Flag_CreateDDDL = targetWorksheet.Range("A9").Value.ToString
            sheetName10 = targetWorksheet.Range("A10").Value.ToString
            sheetName11 = targetWorksheet.Range("A11").Value.ToString

        ElseIf Target.Worksheet.Name = targetWorksheet.Range("B11").Value And excelApp.Intersect(Target, excelApp.Range(targetWorksheet.Range("B2").Value)) IsNot Nothing Then
            Variable1 = targetWorksheet.Range("B1").Value.ToString()
            Variable2 = targetWorksheet.Range("B2").Value.ToString()
            Header = targetWorksheet.Range("B3").Value.ToString()
            Ascending = targetWorksheet.Range("B4").Value.ToString()
            Descending = targetWorksheet.Range("B5").Value.ToString()
            TextConvert = targetWorksheet.Range("B6").Value.ToString()
            OptionType = targetWorksheet.Range("B7").Value.ToString()
            Horizontal_CreateDP = targetWorksheet.Range("B8").Value.ToString()
            Flag_CreateDDDL = targetWorksheet.Range("B9").Value.ToString
            sheetName10 = targetWorksheet.Range("B10").Value.ToString
            sheetName11 = targetWorksheet.Range("B11").Value.ToString


        ElseIf Target.Worksheet.Name = targetWorksheet.Range("C11").Value And excelApp.Intersect(Target, excelApp.Range(targetWorksheet.Range("C2").Value)) IsNot Nothing Then
            Variable1 = targetWorksheet.Range("C1").Value.ToString()
            Variable2 = targetWorksheet.Range("C2").Value.ToString()
            Header = targetWorksheet.Range("C3").Value.ToString()
            Ascending = targetWorksheet.Range("C4").Value.ToString()
            Descending = targetWorksheet.Range("C5").Value.ToString()
            TextConvert = targetWorksheet.Range("C6").Value.ToString()
            OptionType = targetWorksheet.Range("C7").Value.ToString()
            Horizontal_CreateDP = targetWorksheet.Range("C8").Value.ToString()
            Flag_CreateDDDL = targetWorksheet.Range("C9").Value.ToString
            sheetName10 = targetWorksheet.Range("C10").Value.ToString
            sheetName11 = targetWorksheet.Range("C11").Value.ToString

        ElseIf Target.Worksheet.Name = targetWorksheet.Range("D11").Value And excelApp.Intersect(Target, excelApp.Range(targetWorksheet.Range("D2").Value)) IsNot Nothing Then
            Variable1 = targetWorksheet.Range("D1").Value.ToString()
            Variable2 = targetWorksheet.Range("D2").Value.ToString()
            Header = targetWorksheet.Range("D3").Value.ToString()
            Ascending = targetWorksheet.Range("D4").Value.ToString()
            Descending = targetWorksheet.Range("D5").Value.ToString()
            TextConvert = targetWorksheet.Range("D6").Value.ToString()
            OptionType = targetWorksheet.Range("D7").Value.ToString()
            Horizontal_CreateDP = targetWorksheet.Range("D8").Value.ToString()
            Flag_CreateDDDL = targetWorksheet.Range("D9").Value.ToString
            sheetName10 = targetWorksheet.Range("D10").Value.ToString
            sheetName11 = targetWorksheet.Range("D11").Value.ToString

        ElseIf Target.Worksheet.Name = targetWorksheet.Range("E11").Value And excelApp.Intersect(Target, excelApp.Range(targetWorksheet.Range("E2").Value)) IsNot Nothing Then
            Variable1 = targetWorksheet.Range("E1").Value.ToString()
            Variable2 = targetWorksheet.Range("E2").Value.ToString()
            Header = targetWorksheet.Range("E3").Value.ToString()
            Ascending = targetWorksheet.Range("E4").Value.ToString()
            Descending = targetWorksheet.Range("E5").Value.ToString()
            TextConvert = targetWorksheet.Range("E6").Value.ToString()
            OptionType = targetWorksheet.Range("E7").Value.ToString()
            Horizontal_CreateDP = targetWorksheet.Range("E8").Value.ToString()
            Flag_CreateDDDL = targetWorksheet.Range("E9").Value.ToString
            sheetName10 = targetWorksheet.Range("E10").Value.ToString
            sheetName11 = targetWorksheet.Range("E11").Value.ToString
        End If

        src_rng = excelApp.Range(Variable1)
        Dim src_ws As Excel.Worksheet = CType(workBook.Worksheets(sheetName10), Excel.Worksheet)
        Dim des_ws As Excel.Worksheet = CType(workBook.Worksheets(sheetName11), Excel.Worksheet)
        src_rng = src_ws.Range(Variable1)


        'des_rng = des_ws.Range(des_rng.Address)
        des_rng = des_ws.Range(Variable2)
        'MsgBox(src_rng.Address)
        'MsgBox(des_rng.Address)

        If excelApp.Intersect(Target, des_rng) IsNot Nothing Then


            Dim rng As Excel.Range

            ' Dim rng As Excel.Range
            'des_rng.ClearContents()

            If OptionType = True Then
                If Header = True Then
                    'Dim adjustRange As Excel.Range
                    rng = src_rng.Offset(1, 0).Resize(src_rng.Rows.Count - 1, src_rng.Columns.Count)

                Else

                    rng = src_rng 'Assuming you have a range from A1 to A100
                End If
                ' MsgBox(des_rng.Rows.Count)
                Dim col_dif As Integer
                col_dif = Target.Column - worksheet.Range(des_rng.Address).Column + 1
                'MsgBox(col_dif)

                'For k = 1 To des_rng.Rows.Count
                Dim matchedValues As New List(Of String)
                Dim sec_matchedValues As New List(Of String)
                Dim thrd_matchedValues As New List(Of String)
                Dim four_matchedValues As New List(Of String)
                Dim k As Integer = Target.Row - worksheet.Range(des_rng.Address).Row + 1
                'MsgBox(k)
                'MsgBox(i)
                If col_dif = 1 Then

                    If des_rng(k, 1).Value IsNot Nothing Then
                        For i = 1 To rng.Rows.Count
                            If rng(i, 1).Value = des_rng(k, 1).Value Then
                                If Not matchedValues.Contains(rng(i, 2).Value) Then
                                    matchedValues.Add(rng(i, 2).Value)
                                End If
                                'matchedValues.Add(rng(i, 2).Value)
                            End If
                        Next


                        If Ascending = True Then
                            'Sort the list in ascending order
                            matchedValues.Sort()
                        ElseIf Descending = True Then
                            'Sort the list in ascending order
                            matchedValues.Sort()
                            matchedValues.Reverse()
                        End If

                        'Dim dropDownRange As Excel.Range = des_rng(k, 2)
                        Dim dropDownRange As Excel.Range = Target(1, 2)
                        '  MsgBox(Target.Address)
                        Dim Validation As Excel.Validation = dropDownRange.Validation
                        Validation.Delete() 'Remove any existing validation
                        Validation.Add(Excel.XlDVType.xlValidateList, Formula1:=String.Join(",", matchedValues))
                        matchedValues.Clear()
                        'MsgBox(100)
                    End If

                    '  Dim sec_matchedValues As New List(Of String)
                ElseIf col_dif = 2 Then
                    If des_rng(k, 2).Value IsNot Nothing Then
                        For i = 1 To rng.Rows.Count
                            If rng(i, 1).Value = des_rng(k, 1).Value And rng(i, 2).Value = des_rng(k, 2).Value Then
                                If Not sec_matchedValues.Contains(rng(i, 3).Value) Then
                                    sec_matchedValues.Add(rng(i, 3).Value)
                                End If

                            End If
                        Next


                        If Ascending = True Then
                            'Sort the list in ascending order
                            sec_matchedValues.Sort()
                        ElseIf Descending = True Then
                            'Sort the list in ascending order
                            sec_matchedValues.Sort()
                            sec_matchedValues.Reverse()
                        End If


                        ' Dim dropDownRange As Excel.Range = des_rng(k, 3)
                        Dim dropDownRange As Excel.Range = Target(, 2)
                        Dim Validation As Excel.Validation = dropDownRange.Validation
                        Validation.Delete() 'Remove any existing validation
                        Validation.Add(Excel.XlDVType.xlValidateList, Formula1:=String.Join(",", sec_matchedValues))
                        sec_matchedValues.Clear()
                    End If
                ElseIf col_dif = 3 Then
                    '       Dim thrd_matchedValues As New List(Of String)

                    If des_rng(k, 3).Value IsNot Nothing Then
                        For i = 1 To rng.Rows.Count
                            If rng(i, 1).Value = des_rng(k, 1).Value And rng(i, 2).Value = des_rng(k, 2).Value And rng(i, 3).Value = des_rng(k, 3).Value Then
                                If Not thrd_matchedValues.Contains(rng(i, 4).Value) Then
                                    thrd_matchedValues.Add(rng(i, 4).Value)
                                End If

                            End If
                        Next


                        If Ascending = True Then
                            'Sort the list in ascending order
                            thrd_matchedValues.Sort()
                        ElseIf Descending = True Then
                            'Sort the list in ascending order
                            thrd_matchedValues.Sort()
                            thrd_matchedValues.Reverse()
                        End If


                        'Dim dropDownRange As Excel.Range = des_rng(k, 4)
                        Dim dropDownRange As Excel.Range = Target(, 2)
                        Dim Validation As Excel.Validation = dropDownRange.Validation
                        Validation.Delete() 'Remove any existing validation
                        Validation.Add(Excel.XlDVType.xlValidateList, Formula1:=String.Join(",", thrd_matchedValues))
                        thrd_matchedValues.Clear()
                    End If


                    '  Dim four_matchedValues As New List(Of String)
                ElseIf col_dif = 4 Then
                    If des_rng(k, 4).Value IsNot Nothing Then
                        For i = 1 To rng.Rows.Count
                            If rng(i, 1).Value = des_rng(k, 1).Value And rng(i, 2).Value = des_rng(k, 2).Value And rng(i, 3).Value = des_rng(k, 3).Value And rng(i, 4).Value = des_rng(k, 4).Value Then

                                If Not four_matchedValues.Contains(rng(i, 5).Value) Then
                                    four_matchedValues.Add(rng(i, 5).Value)
                                End If


                            End If
                        Next


                        If Ascending = True Then
                            'Sort the list in ascending order
                            four_matchedValues.Sort()
                        ElseIf Descending = True Then
                            'Sort the list in ascending order
                            four_matchedValues.Sort()
                            four_matchedValues.Reverse()
                        End If


                        Dim dropDownRange As Excel.Range = des_rng(k, 5)
                        Dim Validation As Excel.Validation = dropDownRange.Validation
                        Validation.Delete() 'Remove any existing validation
                        Validation.Add(Excel.XlDVType.xlValidateList, Formula1:=String.Join(",", four_matchedValues))
                        four_matchedValues.Clear()
                    End If
                End If

                'Next

            ElseIf OptionType = False Then
                If Horizontal_CreateDP = True Then
                    If Target.Address = des_rng(1, 1).Address Then

                        Dim worksheet As Excel.Worksheet = CType(Target.Worksheet, Excel.Worksheet)
                        Dim col As Integer = src_rng.Rows().Find(Target.Value).Column - src_rng.Column + 1
                        'MsgBox(col)
                        'Dim ab As Integer = col - src_rng.Column
                        Dim sourceRng As Excel.Range = src_rng.Cells(2, col).Resize(src_rng(src_rng.Rows.Count, col).row - 2, 1)
                        'MsgBox(sourceRng.Address)
                        'Dim sourceRng As Excel.Range = src_rng.Cells(2, col).Resize(worksheet.Cells(worksheet.Rows.Count, col), 1)
                        Dim dropDownRange As Excel.Range = des_rng(1, 2)
                        Dim Validation As Excel.Validation = dropDownRange.Validation
                        Validation.Delete() 'Remove any existing validation
                        Dim formula As String = "='" & sourceRng.Worksheet.Name & "'!" & sourceRng.Address(External:=False)
                        Validation.Add(Excel.XlDVType.xlValidateList, Formula1:=formula)
                        'CreateValidationList(worksheet.Cells(2, 5), "=" & sourceRng.Address)
                    End If

                ElseIf Horizontal_CreateDP = False Then
                    If Target.Address = des_rng(1, 1).Address Then
                        Dim worksheet As Excel.Worksheet = CType(Target.Worksheet, Excel.Worksheet)
                        Dim col As Integer = src_rng.Rows().Find(Target.Value).Column - src_rng.Column + 1
                        'MsgBox(col)
                        'Dim ab As Integer = col - src_rng.Column
                        Dim sourceRng As Excel.Range = src_rng.Cells(2, col).Resize(src_rng(src_rng.Rows.Count, col).row - 2, 1)
                        'MsgBox(sourceRng.Address)
                        'Dim sourceRng As Excel.Range = src_rng.Cells(2, col).Resize(worksheet.Cells(worksheet.Rows.Count, col), 1)
                        Dim dropDownRange As Excel.Range = des_rng(2, 1)
                        Dim Validation As Excel.Validation = dropDownRange.Validation
                        Validation.Delete() 'Remove any existing validation
                        Dim formula As String = "='" & sourceRng.Worksheet.Name & "'!" & sourceRng.Address(External:=False)
                        Validation.Add(Excel.XlDVType.xlValidateList, Formula1:=formula)
                    End If
                End If

            End If

            'MsgBox(5)

        End If
        ' Catch ex As Exception
        'MsgBox("error")
        'End Try
        'MsgBox(3)

    End Sub

    'For picturebox Drop-down
    Private Sub worksheet6_Change(ByVal Target As Excel.Range)

        excelApp = Globals.ThisAddIn.Application
        Dim workbook As Excel.Workbook = excelApp.ActiveWorkbook
        Dim worksheet As Excel.Worksheet = workbook.ActiveSheet

        Dim src_rng As Excel.Range = excelApp.Range(Src_Rng_of_PictureDDL)
        'MsgBox(worksheet.Name)
        'Target = worksheet.Range(Target.Address)

        For i = 1 To src_rng.Rows.Count
            If src_rng(i, 1).Value = Target.Value Then
                'MsgBox(3)
                Try
                    worksheet7_Change(Target)
                    'MsgBox(5)
                Catch ex As Exception
                    'MsgBox(15)
                End Try

                'MsgBox(6)

                '            Dim imageCell As Excel.Range = worksheet.Range(src_rng(i, 2).address)
                '            imageCell.CopyPicture(
                'Appearance:=Excel.XlPictureAppearance.xlScreen,
                'Format:=Excel.XlCopyPictureFormat.xlPicture)
                '            worksheet.Paste(Target.Offset(0, 1))
                '            Me.Refresh()
                '            MsgBox(2)

                Dim x As Boolean = False

                For Each pic As Excel.Shape In worksheet.Shapes
                    'MsgBox(pic.TopLeftCell.Address)
                    If pic.TopLeftCell.Address = src_rng(i, 2).Address Then

                        pic.CopyPicture()
                        worksheet.Paste(Target.Offset(0, 1))
                        Target.Offset(0, 1).RowHeight = src_rng(i, 2).RowHeight
                        ' Target.Offset(0, 1).RowHeight = src_rng(i, 2).C
                        x = True
                        Exit For
                    End If
                    'x = x + 1
                Next

                excelApp.CutCopyMode = False
                'Exit Sub

            End If
        Next


    End Sub

    Private Sub worksheet7_Change(ByVal Target As Excel.Range)

        excelApp = Globals.ThisAddIn.Application
        Dim workbook As Excel.Workbook = excelApp.ActiveWorkbook
        Dim worksheet As Excel.Worksheet = workbook.ActiveSheet

        'MsgBox(worksheet.Shapes.Count)
        'MsgBox(worksheet.Name)

        For Each pic As Excel.Shape In worksheet.Shapes
            'MsgBox(pic.TopLeftCell.Address)
            If pic.TopLeftCell.Address = Target.Offset(0, 1).Address Then

                pic.Delete()
                'Exit For
            End If
        Next
        'MsgBox(4)
        ' End Sub
    End Sub


End Class




