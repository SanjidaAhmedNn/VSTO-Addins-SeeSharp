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

    Private WithEvents wsEvent1 As Excel.Worksheet
    Private WithEvents wsEvent2 As Excel.Worksheet
    Private WithEvents wsEvent3 As Excel.Worksheet
    Private WithEvents wsEvent4 As Excel.Worksheet
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
            ' Define an array of type Excel.Worksheet
            Dim sheetsArray() As Excel.Worksheet

            ' Resize the array based on the number of sheets
            ReDim sheetsArray(workBook.Worksheets.Count)
            For i = 1 To excelApp.ActiveWorkbook.Worksheets.Count
                sheetsArray(i) = CType(workBook.Worksheets(i), Excel.Worksheet)
                If sheetName = sheetsArray(i).Name Then
                    wsEvent4 = DirectCast(sheetsArray(i), Excel.Worksheet)
                    ' AddHandler sheetsArray(i).Change, AddressOf sheet_SelectionChange3
                    'src_rng = sheetsArray(i).Range("A1", workSheet.Cells(excelApp.Rows.Count, excelApp.Columns.Count))
                End If
                'i = i + 1

            Next
            'wsEvent4 = DirectCast(worksheet, Excel.Worksheet)
            'AddHandler worksheet.Change, AddressOf worksheet5_Change
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
    Private Sub wsEvent4_SelectionChange(ByVal Target As Excel.Range) Handles wsEvent4.Change
        ' For testing purposes, we'll just show a message box.
        'MsgBox("Cell selected: " & Target.Address)
        'MsgBox(Target.Address)
        worksheet5_Change(Target)
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

        'workSheet2.Range("A2").Value = GB_CB_Source1   'For Source Range
        'workSheet2.Range("A3").Value = SR1             '  Range Type
        'workSheet2.Range("A4").Value = Horizontal1
        'workSheet2.Range("A5").Value = Separator1
        'workSheet2.Range("A6").Value = Search1
        'workSheet2.Range("A7").Value = Flag1            'Activated
        'workSheet2.Range("A8").Value = TargetVar1

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
                ' Write something in cell A1 of the target worksheet
                'targetWorksheet.Range("A1").Value = Variable1
                'targetWorksheet.Range("A2").Value = Variable2
                'targetWorksheet.Range("A3").Value = Header
                'targetWorksheet.Range("A4").Value = Ascending
                'targetWorksheet.Range("A5").Value = Descending
                'targetWorksheet.Range("A6").Value = TextConvert
                'targetWorksheet.Range("A7").Value = OptionType
                'targetWorksheet.Range("A8").Value = Horizontal_CreateDP
                Variable1 = ws.Range("A1").Value.ToString()
                Variable2 = ws.Range("A2").Value.ToString()
                Header = ws.Range("A3").Value.ToString()
                Ascending = ws.Range("A4").Value.ToString()
                Descending = ws.Range("A5").Value.ToString()
                TextConvert = ws.Range("A6").Value.ToString()
                OptionType = ws.Range("A7").Value.ToString()
                Horizontal_CreateDP = ws.Range("A8").Value.ToString()
                Flag_CreateDDDL = ws.Range("A9").value.ToString
                sheetName = ws.Range("A10").value.ToString
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

            'MsgBox(2)
            Dim src_rng_concate As Excel.Range
            'MsgBox(workSheet.Name)
            'MsgBox(src_rng.Worksheet.Name)

            'src_rng = workSheet.Range(GB_CB_Source1)

            If GB_CB_Source2 IsNot Nothing Then
                Dim src_rng As Excel.Range = worksheet.Range(GB_CB_Source2)
                'MsgBox(src_rng.Address)

                'MsgBox(src_rng.Address)
                'src_rng = workSheet.Range(GB_CB_Source1)


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
    Public Sub worksheet5_Change(ByVal Target As Excel.Range)
        Try
            Dim des_rng As Excel.Range = excelApp.Range(Variable2)
            Dim src_rng As Excel.Range = excelApp.Range(Variable1)
            src_rng = excelApp.Range(Variable1)
            'MsgBox(src_rng.Address)
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
                col_dif = Target.Column - worksheet.Range(Variable2).Column + 1
                'MsgBox(col_dif)

                'For k = 1 To des_rng.Rows.Count
                Dim matchedValues As New List(Of String)
                Dim sec_matchedValues As New List(Of String)
                Dim thrd_matchedValues As New List(Of String)
                Dim four_matchedValues As New List(Of String)
                Dim k As Integer = Target.Row - worksheet.Range(Variable2).Row + 1
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
                        Validation.Add(Excel.XlDVType.xlValidateList, Formula1:="=" & sourceRng.Address)
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
                        Validation.Add(Excel.XlDVType.xlValidateList, Formula1:="=" & sourceRng.Address)
                    End If
                End If

            End If
        Catch ex As Exception
            'MsgBox("error")
        End Try


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




