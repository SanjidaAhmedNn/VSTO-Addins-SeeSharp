
Imports System.ComponentModel
Imports System.Reflection.Emit
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports Microsoft.Office.Interop

Imports Microsoft.Office.Interop.Excel


Public Class Form30_Create_Dynamic_Drop_down_List

    Dim WithEvents excelApp As Excel.Application
    Dim workBook As Excel.Workbook
    Public Shared workSheet As Excel.Worksheet
    Dim workSheet2 As Excel.Worksheet
    Dim workSheet3 As Excel.Worksheet
    Dim src_rng As Excel.Range
    Public des_rng As Excel.Range
    Dim selectedRange As Excel.Range
    Public ax As String
    Public focuschange As Boolean

    Dim opened As Integer
    'Public WithEvents Btn_OK As System.Windows.Forms.Button


    Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As IntPtr, ByVal hWndInsertAfter As IntPtr, ByVal X As Integer, ByVal Y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal uFlags As UInteger) As Boolean
    Private Const SWP_NOMOVE As UInteger = &H2
    Private Const SWP_NOSIZE As UInteger = &H1
    Private Const SWP_NOACTIVATE As UInteger = &H10
    Private Const HWND_TOPMOST As Integer = -1

    Private Sub CB_ascending_CheckedChanged(sender As Object, e As EventArgs) Handles CB_ascending.CheckedChanged
        If CB_ascending.Checked = True Then
            CB_descending.Checked = False
        End If
    End Sub

    Private Sub CB_descending_CheckedChanged(sender As Object, e As EventArgs) Handles CB_descending.CheckedChanged
        If CB_descending.Checked = True Then
            CB_ascending.Checked = False
        End If
    End Sub

    Private Sub RB_columns_CheckedChanged(sender As Object, e As EventArgs) Handles RB_Dropdown_35_Labels.CheckedChanged
        If RB_Dropdown_35_Labels.Checked = True Then

            CB_header.Enabled = True
            CB_ascending.Enabled = True
            CB_descending.Enabled = True
            CB_text.Enabled = True
            GB_list_option.Enabled = False

        End If
    End Sub

    Private Sub RB_rows_CheckedChanged(sender As Object, e As EventArgs) Handles RB_Dropdown_2_Labels.CheckedChanged
        If RB_Dropdown_2_Labels.Checked = True Then
            GB_list_option.Enabled = True
            CB_header.Enabled = False
            CB_ascending.Enabled = False
            CB_descending.Enabled = False
            CB_text.Enabled = False

        End If
    End Sub



    Private Sub Selection_source_Click(sender As Object, e As EventArgs) Handles Selection_source.Click
        Try
            If selectedRange Is Nothing Then
            Else

                TB_src_range.Text = selectedRange.Address


                'FocusedTextBox = 1
                Me.Hide()

                excelApp = Globals.ThisAddIn.Application
                workBook = excelApp.ActiveWorkbook

                Dim userInput As Excel.Range = excelApp.InputBox("Select a range", "Select a range", "=$A$1", Type:=8)
                src_rng = userInput

                Dim sheetName As String
                sheetName = Split(src_rng.Address(True, True, Excel.XlReferenceStyle.xlA1, True), "]")(1)
                sheetName = Split(sheetName, "!")(0)

                If Mid(sheetName, Len(sheetName), 1) = "'" Then
                    sheetName = Mid(sheetName, 1, Len(sheetName) - 1)
                End If

                workSheet = workBook.Worksheets(sheetName)
                workSheet.Activate()

                src_rng.Select()

                TB_src_range.Text = src_rng.Address

                Me.Show()
                TB_src_range.Focus()
            End If

        Catch ex As Exception

            Me.Show()
            TB_src_range.Focus()

        End Try
    End Sub

    ' Event handler to detect changes in E1 and adjust dropdown in E2

    Public Sub Btn_OK_Click(sender As Object, e As EventArgs) Handles Btn_OK.Click

        excelApp = Globals.ThisAddIn.Application
        Dim workbook As Excel.Workbook = excelApp.ActiveWorkbook
        Dim worksheet As Excel.Worksheet = workbook.ActiveSheet

        If TB_src_range.Text = "" And TB_dest_range.Text = "" Then
            MessageBox.Show("Please select all necessary options.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TB_src_range.Focus()
            'Me.Close()
            Exit Sub

        ElseIf TB_src_range.Text = "" Then
            MessageBox.Show("Please select the Source Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TB_src_range.Focus()
            'Me.Close()
            Exit Sub
            'End If

        ElseIf IsValidExcelCellReference(TB_src_range.Text) = False Then
            MessageBox.Show("Select a Valid Source Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TB_src_range.Focus()
            'Me.Close()
            Exit Sub
            ' End If

        ElseIf TB_dest_range.Text = "" Then
            MessageBox.Show("Select a Destination Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TB_dest_range.Focus()
            'Me.Close()
            Exit Sub
            ' End If

        ElseIf IsValidExcelCellReference(TB_dest_range.Text) = False Then
            MessageBox.Show("Select a Valid Destination Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TB_dest_range.Focus()
            'Me.Close()
            Exit Sub
            ' End If


        ElseIf RB_Dropdown_2_Labels.Checked = False And RB_Dropdown_35_Labels.Checked = False Then
            MessageBox.Show("Select a Drop-down List type.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            worksheet.Activate()
            src_rng.Select()
            'Me.Close()
            '   Exit Sub
            Exit Sub

        ElseIf RB_Dropdown_2_Labels.Checked = True And RB_Horizon.Checked = False And RB_Verti.Checked = False Then
            MessageBox.Show("Select a Flip Option.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            worksheet.Activate()
            src_rng.Select()
            'Me.Close()
            Exit Sub
            ' End If

        ElseIf RB_Dropdown_35_Labels.Checked = True And src_rng.Columns.Count > 5 Then
            MessageBox.Show("You can maximum select 5 columns.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            worksheet.Activate()
            src_rng.Select()
            'Me.Close()
            Exit Sub

        ElseIf src_rng.Areas.Count > 1 Then
            MessageBox.Show("Multiple selection is not possible in the Source Range field.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TB_src_range.Focus()
            Exit Sub


        ElseIf RB_Dropdown_2_Labels.Checked = True And src_rng.Rows.Count < 2 Then
            MessageBox.Show("Select a valid Source Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TB_src_range.Focus()
            Exit Sub

        ElseIf RB_Dropdown_2_Labels.Checked = True And RB_Horizon.Checked = True And des_rng.Columns.Count <> 2 Then
            MessageBox.Show("Select a valid Destination Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TB_dest_range.Focus()
            Exit Sub

        ElseIf RB_Dropdown_2_Labels.Checked = True And RB_Verti.Checked = True And des_rng.Rows.Count <> 2 Then
            MessageBox.Show("Select a valid Destination Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TB_dest_range.Focus()
            Exit Sub

        Else
            Try
                If RB_Dropdown_35_Labels.Checked = True Then
                    Dim rng As Excel.Range
                    If CB_header.Checked = True Then
                        'Dim adjustRange As Excel.Range
                        rng = src_rng.Offset(1, 0).Resize(src_rng.Rows.Count - 1, src_rng.Columns.Count)

                    Else

                        rng = src_rng 'Assuming you have a range from A1 to A100

                    End If

                    Dim uniqueValues As New List(Of String)

                    'Extract unique values from the range
                    For Each cell As Excel.Range In rng.Columns(1).Cells
                        Dim value As String = cell.Value
                        If Not uniqueValues.Contains(value) Then
                            uniqueValues.Add(value)
                        End If
                    Next

                    If CB_ascending.Checked = True Then
                        'Sort the list in ascending order
                        uniqueValues.Sort()
                    ElseIf CB_descending.Checked = True Then
                        'Sort the list in ascending order
                        uniqueValues.Sort()
                        uniqueValues.Reverse()
                    End If

                    'Create drop-down list at B1 with the unique values
                    Dim dropDownRange As Excel.Range = des_rng.Columns(1)
                    Dim validation As Excel.Validation = dropDownRange.Validation
                    validation.Delete() 'Remove any existing validation
                    validation.Add(Excel.XlDVType.xlValidateList, Formula1:=String.Join(",", uniqueValues))

                    AddHandler workSheet3.Change, AddressOf worksheet1_Change


                ElseIf RB_Dropdown_2_Labels.Checked = True Then
                    ' Extract headers from A1:C1
                    'MsgBox(src_rng.Address)
                    src_rng = workSheet2.Range(src_rng.Address)

                    'MsgBox(workSheet2.Name)
                    Dim headersRange As Excel.Range = src_rng.Rows(1)
                    Dim headers As List(Of String) = New List(Of String)
                    ' Dim workbook As excelapp.workbook

                    For Each cell As Excel.Range In headersRange.Cells
                        headers.Add(cell.Value.ToString())
                    Next
                    'Dim workbook As Excel.Workbook = excelApp.ActiveWorkbook
                    'Dim worksheet As Excel.Worksheet = workbook.ActiveSheet
                    ' Create the dropdown list with headers in cell E1
                    'CreateValidationList(excelApp.ActiveSheet.Range("$E$1"), String.Join(",", headers))
                    'Create drop-down list at B1 with the unique values

                    Dim dropDownRange As Excel.Range = des_rng(1, 1)
                    Dim validation As Excel.Validation = dropDownRange.Validation
                    validation.Delete() 'Remove any existing validation
                    validation.Add(Excel.XlDVType.xlValidateList, Formula1:=String.Join(",", headers))

                    ' Add event handler to listen for changes in E1

                    ' AddHandler worksheet.Change, AddressOf worksheet1_Change
                    AddHandler workSheet3.Change, AddressOf worksheet1_Change
                End If

                If CB_text.Checked = True Then
                    des_rng.NumberFormat = "@"
                End If

                If des_rng.Worksheet.Name <> src_rng.Worksheet.Name Then
                    Variable1 = src_rng.Worksheet.Name & "!" & TB_src_range.Text
                    Variable2 = TB_dest_range.Text
                Else
                    Variable1 = src_rng.Worksheet.Name & "!" & TB_src_range.Text
                    Variable2 = des_rng.Worksheet.Name & "!" & TB_dest_range.Text
                End If

                Header = CB_header.Checked
                Ascending = CB_ascending.Checked
                Descending = CB_descending.Checked
                TextConvert = CB_text.Checked
                ' MsgBox(CB_header.Checked)



                Dim targetWorksheet As Excel.Worksheet = Nothing
                For Each ws As Excel.Worksheet In excelApp.Worksheets
                    If ws.Name = "MySpecialSheet" Then
                        targetWorksheet = ws
                        Exit For
                    End If
                Next

                ' If "MySpecialSheet" does not exist, add it
                If targetWorksheet Is Nothing Then
                    targetWorksheet = CType(excelApp.Worksheets.Add(After:=excelApp.Worksheets(excelApp.Worksheets.Count)), Excel.Worksheet)
                    targetWorksheet.Name = "MySpecialSheet"
                End If

                If RB_Dropdown_2_Labels.Checked = True Then
                    OptionType = False     '2 label=False
                Else
                    OptionType = True      '3-5 label=true

                End If

                If RB_Horizon.Checked = True And CustomGroupBox5.Enabled = True Then
                    Horizontal_CreateDP = True
                ElseIf RB_Verti.Checked = True And CustomGroupBox5.Enabled = True Then
                    Horizontal_CreateDP = False
                End If

                Flag_CreateDDDL = True
                'sheetName = worksheet.Name
                sheetName10 = workSheet2.Name
                sheetName11 = workSheet3.Name
                Dim sheetName1 As String = src_rng.Worksheet.Name
                Dim sheetName2 As String = des_rng.Worksheet.Name

                If targetWorksheet.Range("A1").Value = "" Then
                    ' Write something in cell A1 of the target worksheet
                    targetWorksheet.Range("A1").Value = Variable1
                    targetWorksheet.Range("A2").Value = Variable2
                    targetWorksheet.Range("A3").Value = Header
                    targetWorksheet.Range("A4").Value = Ascending
                    targetWorksheet.Range("A5").Value = Descending
                    targetWorksheet.Range("A6").Value = TextConvert
                    targetWorksheet.Range("A7").Value = OptionType
                    targetWorksheet.Range("A8").Value = Horizontal_CreateDP
                    targetWorksheet.Range("A9").Value = Flag_CreateDDDL
                    targetWorksheet.Range("A10").Value = sheetName1
                    targetWorksheet.Range("A11").Value = sheetName2

                ElseIf targetWorksheet.Range("B1").Value = "" Then

                    targetWorksheet.Range("B1").Value = Variable1
                    targetWorksheet.Range("B2").Value = Variable2
                    targetWorksheet.Range("B3").Value = Header
                    targetWorksheet.Range("B4").Value = Ascending
                    targetWorksheet.Range("B5").Value = Descending
                    targetWorksheet.Range("B6").Value = TextConvert
                    targetWorksheet.Range("B7").Value = OptionType
                    targetWorksheet.Range("B8").Value = Horizontal_CreateDP
                    targetWorksheet.Range("B9").Value = Flag_CreateDDDL
                    targetWorksheet.Range("B10").Value = sheetName1
                    targetWorksheet.Range("B11").Value = sheetName2

                ElseIf targetWorksheet.Range("C1").Value = "" Then

                    targetWorksheet.Range("C1").Value = Variable1
                    targetWorksheet.Range("C2").Value = Variable2
                    targetWorksheet.Range("C3").Value = Header
                    targetWorksheet.Range("C4").Value = Ascending
                    targetWorksheet.Range("C5").Value = Descending
                    targetWorksheet.Range("C6").Value = TextConvert
                    targetWorksheet.Range("C7").Value = OptionType
                    targetWorksheet.Range("C8").Value = Horizontal_CreateDP
                    targetWorksheet.Range("C9").Value = Flag_CreateDDDL
                    targetWorksheet.Range("C10").Value = sheetName1
                    targetWorksheet.Range("C11").Value = sheetName2

                ElseIf targetWorksheet.Range("D1").Value = "" Then

                    targetWorksheet.Range("D1").Value = Variable1
                    targetWorksheet.Range("D2").Value = Variable2
                    targetWorksheet.Range("D3").Value = Header
                    targetWorksheet.Range("D4").Value = Ascending
                    targetWorksheet.Range("D5").Value = Descending
                    targetWorksheet.Range("D6").Value = TextConvert
                    targetWorksheet.Range("D7").Value = OptionType
                    targetWorksheet.Range("D8").Value = Horizontal_CreateDP
                    targetWorksheet.Range("D9").Value = Flag_CreateDDDL
                    targetWorksheet.Range("D10").Value = sheetName1
                    targetWorksheet.Range("D11").Value = sheetName2

                ElseIf targetWorksheet.Range("E1").Value = "" Then

                    targetWorksheet.Range("E1").Value = Variable1
                    targetWorksheet.Range("E2").Value = Variable2
                    targetWorksheet.Range("E3").Value = Header
                    targetWorksheet.Range("E4").Value = Ascending
                    targetWorksheet.Range("E5").Value = Descending
                    targetWorksheet.Range("E6").Value = TextConvert
                    targetWorksheet.Range("E7").Value = OptionType
                    targetWorksheet.Range("E8").Value = Horizontal_CreateDP
                    targetWorksheet.Range("E9").Value = Flag_CreateDDDL
                    targetWorksheet.Range("E10").Value = sheetName1
                    targetWorksheet.Range("E11").Value = sheetName2
                Else
                    ' Cut range D1:D10
                    targetWorksheet.Range("B1:E11").Copy()

                    ' Paste to range E1
                    targetWorksheet.Range("A1:D11").PasteSpecial(Excel.XlPasteType.xlPasteAll)
                    excelApp.CutCopyMode = False

                    targetWorksheet.Range("E1:E11").Value = ""
                    targetWorksheet.Range("E1").Value = Variable1
                    targetWorksheet.Range("E2").Value = Variable2
                    targetWorksheet.Range("E3").Value = Header
                    targetWorksheet.Range("E4").Value = Ascending
                    targetWorksheet.Range("E5").Value = Descending
                    targetWorksheet.Range("E6").Value = TextConvert
                    targetWorksheet.Range("E7").Value = OptionType
                    targetWorksheet.Range("E8").Value = Horizontal_CreateDP
                    targetWorksheet.Range("E9").Value = Flag_CreateDDDL
                    targetWorksheet.Range("E10").Value = sheetName1
                    targetWorksheet.Range("E11").Value = sheetName2
                    'MsgBox(105)

                End If
                ' Hide the target worksheet
                targetWorksheet.Visible = Excel.XlSheetVisibility.xlSheetHidden


                des_rng.Value = Nothing
                des_rng.Select()
                MsgBox(100)
                Me.Dispose()
            Catch ex As Exception
                Me.Dispose()
            End Try
        End If

    End Sub

    Private Sub Selection_destination_Click(sender As Object, e As EventArgs) Handles Selection_destination.Click
        If selectedRange Is Nothing Then
        Else
            ' TB_src_range.Text = selectedRange.Address


            Me.Hide()

            excelApp = Globals.ThisAddIn.Application
            workBook = excelApp.ActiveWorkbook

            'Dim userInput As String = excelApp.InputBox("Select a range", "Select range", "=$A$1")


            Dim userInput As Excel.Range = excelApp.InputBox("Select a range", "Select a range", "=$A$1", Type:=8)
            des_rng = userInput

            Dim sheetName As String
            sheetName = Split(des_rng.Address(True, True, Excel.XlReferenceStyle.xlA1, True), "]")(1)
            sheetName = Split(sheetName, "!")(0)

            If Mid(sheetName, Len(sheetName), 1) = "'" Then
                sheetName = Mid(sheetName, 1, Len(sheetName) - 1)
            End If

            workSheet = workBook.Worksheets(sheetName)
            workSheet.Activate()

            des_rng.Select()
            'MsgBox(src_rng.Address)

            TB_dest_range.Text = des_rng.Address

            Me.Show()
            TB_dest_range.Focus()

        End If
    End Sub


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load

        Try

            excelApp = Globals.ThisAddIn.Application

            AddHandler excelApp.SheetSelectionChange, AddressOf excelApp_SheetSelectionChange

            opened = opened + 1

            If excelApp.Selection IsNot Nothing Then
                selectedRange = excelApp.Selection
                src_rng = selectedRange
                TB_src_range.Text = selectedRange.Address
            End If

        Catch ex As Exception

        End Try

    End Sub


    Private Sub excelApp_SheetSelectionChange(ByVal Sh As Object, ByVal selectionRange1 As Excel.Range) Handles excelApp.SheetSelectionChange
        Try

            excelApp = Globals.ThisAddIn.Application
            If focuschange = False Then
                If Me.ActiveControl Is TB_dest_range Then
                    des_rng = selectionRange1
                    ' This will run on the Excel thread, so you need to use Invoke to update the UI
                    'Me.BeginInvoke(New System.Action(Sub() TB_dest_range.Text = selectionRange1.Address))
                    Me.Activate()
                    Me.BeginInvoke(New System.Action(Sub()
                                                         TB_dest_range.Text = des_rng.Address
                                                         SetWindowPos(Me.Handle, New IntPtr(HWND_TOPMOST), 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOSIZE)
                                                     End Sub))

                ElseIf Me.ActiveControl Is TB_src_range Then
                    src_rng = selectionRange1
                    Me.Activate()


                    Me.BeginInvoke(New System.Action(Sub()
                                                         TB_src_range.Text = src_rng.Address
                                                         SetWindowPos(Me.Handle, New IntPtr(HWND_TOPMOST), 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOSIZE)
                                                     End Sub))
                End If
            End If


        Catch ex As Exception

        End Try

    End Sub

    Private Sub Btn_Cancel_Click(sender As Object, e As EventArgs) Handles Btn_Cancel.Click

        Me.Close()
    End Sub



    Public Sub worksheet1_Change(ByVal Target As Excel.Range)
        Try
            excelApp = Globals.ThisAddIn.Application
            workBook = excelApp.ActiveWorkbook
            workSheet = workBook.ActiveSheet



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
            ' Next




            src_rng = excelApp.Range(Variable1)
            Dim src_ws As Excel.Worksheet = CType(workBook.Worksheets(sheetName10), Excel.Worksheet)
            Dim des_ws As Excel.Worksheet = CType(workBook.Worksheets(sheetName11), Excel.Worksheet)
            src_rng = src_ws.Range(Variable1)

                des_rng = des_ws.Range(des_rng.Address)

            If excelApp.Intersect(Target, des_rng) IsNot Nothing Then
                Dim rng As Excel.Range

                'MsgBox(src_rng.Address)
                'MsgBox(des_rng.Address)

                If RB_Dropdown_35_Labels.Checked = True Then
                    If CB_header.Checked = True Then
                        'Dim adjustRange As Excel.Range
                        rng = src_rng.Offset(1, 0).Resize(src_rng.Rows.Count - 1, src_rng.Columns.Count)

                    Else

                        rng = src_rng 'Assuming you have a range from A1 to A100
                    End If

                    Dim col_dif As Integer
                    col_dif = Target.Column - des_rng.Column + 1
                    'MsgBox(col_dif)

                    'For k = 1 To des_rng.Rows.Count
                    Dim matchedValues As New List(Of String)
                    Dim sec_matchedValues As New List(Of String)
                    Dim thrd_matchedValues As New List(Of String)
                    Dim four_matchedValues As New List(Of String)
                    Dim k As Integer = Target.Row - des_rng.Row + 1

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


                            If CB_ascending.Checked = True Then
                                'Sort the list in ascending order
                                matchedValues.Sort()
                            ElseIf CB_descending.Checked = True Then
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


                            If CB_ascending.Checked = True Then
                                'Sort the list in ascending order
                                sec_matchedValues.Sort()
                            ElseIf CB_descending.Checked = True Then
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


                            If CB_ascending.Checked = True Then
                                'Sort the list in ascending order
                                thrd_matchedValues.Sort()
                            ElseIf CB_descending.Checked = True Then
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


                            If CB_ascending.Checked = True Then
                                'Sort the list in ascending order
                                four_matchedValues.Sort()
                            ElseIf CB_descending.Checked = True Then
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

                ElseIf RB_Dropdown_2_Labels.Checked = True Then
                    If RB_Horizon.Checked = True Then
                        If Target.Address = des_rng(1, 1).Address Then

                            'Dim worksheet As Excel.Worksheet = CType(Target.Worksheet, Excel.Worksheet)
                            Dim col As Integer = src_rng.Rows().Find(Target.Value).Column - src_rng.Column + 1

                            Dim sourceRng As Excel.Range = src_rng.Cells(2, col).Resize(src_rng(src_rng.Rows.Count, col).row - 2, 1)

                            Dim dropDownRange As Excel.Range = des_rng(1, 2)
                            Dim Validation As Excel.Validation = dropDownRange.Validation
                            Validation.Delete() 'Remove any existing validation
                            Dim formula As String = "='" & sheetName10 & "'!" & sourceRng.Address(External:=False)
                            'MsgBox(formula)
                            Validation.Add(Excel.XlDVType.xlValidateList, Formula1:=formula)

                        End If

                    ElseIf RB_Verti.Checked = True Then
                        If Target.Address = des_rng(1, 1).Address Then

                            Dim col As Integer = src_rng.Rows().Find(Target.Value).Column - src_rng.Column + 1

                            Dim sourceRng As Excel.Range = src_rng.Cells(2, col).Resize(src_rng(src_rng.Rows.Count, col).row - 2, 1)

                            Dim dropDownRange As Excel.Range = des_rng(2, 1)
                            Dim Validation As Excel.Validation = dropDownRange.Validation
                            Validation.Delete() 'Remove any existing validation

                            Dim formula As String = "='" & sourceRng.Worksheet.Name & "'!" & sourceRng.Address(External:=False)
                            Validation.Add(Excel.XlDVType.xlValidateList, Formula1:=formula)
                        End If
                    End If

                End If

            End If
        Catch ex As Exception

        End Try

        'MsgBox(src_rng.Address)
        'MsgBox(des_rng.Address)
    End Sub
    Sub CreateValidationList(cell As Excel.Range, listValues As String)
        With cell.Validation
            .Delete()
            .Add(Type:=Excel.XlDVType.xlValidateList, AlertStyle:=Excel.XlDVAlertStyle.xlValidAlertStop, Operator:=Excel.XlFormatConditionOperator.xlBetween, Formula1:=listValues)
            .ShowInput = True
            .ShowError = True
        End With
    End Sub

    Private Function IsValidExcelCellReference(cellReference As String) As Boolean

        ' Regular expression pattern for a valid sheet name. This is a simplified version and might not cover all edge cases.
        ' Excel sheet names cannot contain the characters \, /, *, [, ], :, ?, and cannot be 'History'.
        Dim sheetNamePattern As String = "(?i)(?![\/*[\]:?])(?!History)[^\/\[\]*?:\\]+"

        ' Regular expression pattern for a cell reference.
        ' This pattern will match references like A1, $A$1, etc.
        Dim cellPattern As String = "(\$?[A-Z]+\$?[0-9]+)"

        ' Regular expression pattern for an Excel reference.
        ' This pattern will match references like A1:B13, $A$1:$B$13, A1, $B$1, etc.
        Dim singleReferencePattern As String = cellPattern + "(:" + cellPattern + ")?"

        ' Regular expression pattern to allow the sheet name, followed by '!', before the cell reference
        Dim fullPattern As String = "^(" + sheetNamePattern + "!)?(" + singleReferencePattern + ")(," + singleReferencePattern + ")*$"

        ' Create a regex object with the pattern.
        Dim regex As New Regex(fullPattern)

        ' Test the input string against the regex pattern.
        Return regex.IsMatch(cellReference.ToUpper)

    End Function


    Private Sub form(sender As Object, e As KeyEventArgs) Handles Me.KeyDown

        Try
            If e.KeyCode = Keys.Enter Then

                Call Btn_OK_Click(sender, e)

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub CB_asceding(sender As Object, e As KeyEventArgs) Handles CB_ascending.KeyDown

        Try
            If e.KeyCode = Keys.Enter Then

                Call Btn_OK_Click(sender, e)

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub CB_desceding(sender As Object, e As KeyEventArgs) Handles CB_descending.KeyDown

        Try
            If e.KeyCode = Keys.Enter Then

                Call Btn_OK_Click(sender, e)

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub CB_head(sender As Object, e As KeyEventArgs) Handles CB_header.KeyDown

        Try
            If e.KeyCode = Keys.Enter Then

                Call Btn_OK_Click(sender, e)

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub CB_texting(sender As Object, e As KeyEventArgs) Handles CB_text.KeyDown

        Try
            If e.KeyCode = Keys.Enter Then

                Call Btn_OK_Click(sender, e)

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub RB_Label2(sender As Object, e As KeyEventArgs) Handles RB_Dropdown_2_Labels.KeyDown

        Try
            If e.KeyCode = Keys.Enter Then

                Call Btn_OK_Click(sender, e)

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub RB_35(sender As Object, e As KeyEventArgs) Handles RB_Dropdown_35_Labels.KeyDown

        Try
            If e.KeyCode = Keys.Enter Then

                Call Btn_OK_Click(sender, e)

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub RB_horiz(sender As Object, e As KeyEventArgs) Handles RB_Horizon.KeyDown

        Try
            If e.KeyCode = Keys.Enter Then

                Call Btn_OK_Click(sender, e)

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub RB_verticalll(sender As Object, e As KeyEventArgs) Handles RB_Verti.KeyDown

        Try
            If e.KeyCode = Keys.Enter Then

                Call Btn_OK_Click(sender, e)

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub TB_dest_range_Enter(sender As Object, e As KeyEventArgs) Handles TB_dest_range.KeyDown
        'If Enter key is pressed then check if the text is a valid address
        If IsValidExcelCellReference(TB_dest_range.Text) = True And e.KeyCode = Keys.Enter Then
            des_rng = excelApp.Range(TB_dest_range.Text)
            TB_dest_range.Focus()
            des_rng.Select()

            Call Btn_OK_Click(sender, e)   'OK button click event called

            'MsgBox(des_rng.Address)
        ElseIf IsValidExcelCellReference(TB_dest_range.Text) = False And e.KeyCode = Keys.Enter Then
            MessageBox.Show("Select the valid Destination Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TB_dest_range.Text = ""
            TB_dest_range.Focus()
            'Me.Close()
            Exit Sub
        End If
    End Sub

    Private Sub TB_src_range_Enter(sender As Object, e As KeyEventArgs) Handles TB_src_range.KeyDown
        'If Enter key is pressed then check if the text is a valid address

        If IsValidExcelCellReference(TB_src_range.Text) = True And e.KeyCode = Keys.Enter Then
            src_rng = excelApp.Range(TB_src_range.Text)
            TB_src_range.Focus()
            src_rng.Select()

            Call Btn_OK_Click(sender, e)   'OK button click event called

            'MsgBox(des_rng.Address)
        ElseIf IsValidExcelCellReference(TB_src_range.Text) = False And e.KeyCode = Keys.Enter Then
            MessageBox.Show("Select the valid Source Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TB_src_range.Text = ""
            TB_src_range.Focus()
            'Me.Close()
            Exit Sub
        End If
    End Sub

    Private Sub TB_dest_range_TextChanged(sender As Object, e As EventArgs) Handles TB_dest_range.TextChanged
        excelApp = Globals.ThisAddIn.Application
        Dim workbook As Excel.Workbook = excelApp.ActiveWorkbook
        Dim worksheet As Excel.Worksheet = workbook.ActiveSheet
        Try
            If TB_dest_range.Text IsNot Nothing And IsValidExcelCellReference(TB_dest_range.Text) = True Then
                focuschange = True
                Dim sheetname As String = ""
                'MsgBox(1)

                Try
                    'des_rng = excelApp.Range(TB_dest_range.Text)
                    des_rng = worksheet.Range(TB_dest_range.Text)
                    des_rng.Select()
                    'sheetname = ""
                Catch
                    ' Split the string into sheet name and cell address
                    Dim parts As String() = TB_dest_range.Text.Split("!"c)
                    sheetname = parts(0)
                    Dim cellAddress As String = parts(1)
                    worksheet = CType(workbook.Worksheets(sheetName), Worksheet)
                    worksheet.Activate()
                    des_rng = worksheet.Range(cellAddress)
                    des_rng.Select()
                End Try
                'MsgBox(sheetname)
                ' Define the range of cells to read (for example, cells A1 to A10)
                If workSheet2.Name <> worksheet.Name And TB_dest_range.Text.Contains("!") = False Then
                    TB_dest_range.Text = worksheet.Name & "!" & TB_dest_range.Text
                    'src_rng = excelApp.Range(TB_src_range.Text)

                End If

                Me.Activate()
                TB_dest_range.Focus()
                TB_dest_range.SelectionStart = TB_dest_range.Text.Length

                focuschange = False
                ax = worksheet.Name
                workSheet3 = worksheet
            End If
        Catch ex As Exception
            focuschange = False
        End Try
    End Sub

    Private Sub TB_src_range_TextChanged(sender As Object, e As EventArgs) Handles TB_src_range.TextChanged
        excelApp = Globals.ThisAddIn.Application
        Dim workbook As Excel.Workbook = excelApp.ActiveWorkbook
        Dim worksheet As Excel.Worksheet = workbook.ActiveSheet
        Try
            If TB_src_range.Text IsNot Nothing And IsValidExcelCellReference(TB_src_range.Text) = True Then
                focuschange = True

                ' Define the range of cells to read (for example, cells A1 to A10)
                src_rng = excelApp.Range(TB_src_range.Text)
                src_rng = worksheet.Range(TB_src_range.Text)
                src_rng.Select()
                Dim range As Excel.Range = src_rng


                Me.Activate()
                'TB_src_range.Focus()
                TB_src_range.SelectionStart = TB_src_range.Text.Length
                focuschange = False
                workSheet2 = worksheet

            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub Form1_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Enter Then
            Btn_OK.Focus()
            Btn_OK.PerformClick()
        End If
    End Sub

    Private Sub Form30_Create_Dynamic_Drop_down_List_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        form_flag = False
    End Sub

    Private Sub Form30_Create_Dynamic_Drop_down_List_Disposed(sender As Object, e As EventArgs) Handles Me.Disposed
        form_flag = False
    End Sub

    Private Sub Form30_Create_Dynamic_Drop_down_List_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        Me.Focus()
        Me.BringToFront()
        Me.Activate()
        Me.BeginInvoke(New System.Action(Sub()
                                             TB_src_range.Text = src_rng.Address
                                             SetWindowPos(Me.Handle, New IntPtr(HWND_TOPMOST), 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOSIZE)
                                         End Sub))
    End Sub
End Class