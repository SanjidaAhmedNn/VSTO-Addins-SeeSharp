Imports System.ComponentModel
Imports System.Diagnostics
Imports System.Windows.Forms
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop

Imports Microsoft.Office.Interop.Excel
Imports System.Drawing
Imports Microsoft.Office
Imports System.Runtime
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports System.CodeDom.Compiler
Imports System.Text.RegularExpressions

Public Class Form37_MSDropDownCheckBox


    Dim WithEvents excelApp As Excel.Application
    Dim workBook As Excel.Workbook
    Public Shared workSheet As Excel.Worksheet
    Private SheetHandlers As New List(Of WorksheetHandler)
    Private EventDel_CellsChange As Excel.DocEvents_ChangeEventHandler

    Private WithEvents CurrentSheet As Excel.Worksheet
    Private WithEvents WorkbookEvents As Excel.Workbook


    Private Form As Form38 = Nothing

    Dim src_rng As Excel.Range
    Public des_rng As Excel.Range
    Dim selectedRange As Excel.Range

    Public validationRange As Excel.Range

    Private processingEvent As Boolean = False
    Public focuschange As Boolean

    Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As IntPtr, ByVal hWndInsertAfter As IntPtr, ByVal X As Integer, ByVal Y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal uFlags As UInteger) As Boolean
    Private Const SWP_NOMOVE As UInteger = &H2
    Private Const SWP_NOSIZE As UInteger = &H1
    Private Const SWP_NOACTIVATE As UInteger = &H10
    Private Const HWND_TOPMOST As Integer = -1

    <DllImport("user32.dll")>
    Public Shared Function GetAsyncKeyState(ByVal vKey As Integer) As Short
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function GetForegroundWindow() As IntPtr
    End Function


    Private WithEvents Timer1 As New Timer With {.Interval = 100}
    'xlApp = Globals.ThisAddIn.Application
    'Private xlWorkbook As Excel.Workbook
    'Private xlWorksheet As Excel.Worksheet

    Private Sub Form1_HelpButtonClicked(sender As Object, e As CancelEventArgs) Handles MyBase.HelpButtonClicked
        Process.Start("https://www.softeko.co/")
        e.Cancel = True ' This will suppress any additional event handling for the Help button
    End Sub


    Private Sub YourForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.KeyPreview = True
        excelApp = Globals.ThisAddIn.Application
        Dim workbook As Excel.Workbook = excelApp.ActiveWorkbook
        Dim worksheet As Excel.Worksheet = workbook.ActiveSheet

        'Timer1.Start()
        CB_Source.Items.Add("Select Range")
        CB_Source.Items.Add("Active Sheet :" & worksheet.Name)
        CB_Source.Items.Add("Active Workbook :" & workbook.Name)

        Dim i As Integer = 0
        ' Loop through each worksheet in the workbook.
        For Each WS In workbook.Sheets
            ' Check if the worksheet is not hidden.
            If WS.Visible = Excel.XlSheetVisibility.xlSheetVisible And WS.name <> worksheet.Name Then
                CB_Source.Items.Add(WS.Name)
                i = i + 1
            End If
        Next

        'Only Enable when select Range is selected in combobox
        If CB_Source.Text = "Select Range" Then
            TB_src_rng.Enabled = True
            Selection_source.Enabled = True
        Else
            TB_src_rng.Enabled = False
            Selection_source.Enabled = False
        End If

        Try

            excelApp = Globals.ThisAddIn.Application

            AddHandler excelApp.SheetSelectionChange, AddressOf excelApp_SheetSelectionChange

            'opened = opened + 1

            If excelApp.Selection IsNot Nothing Then
                selectedRange = excelApp.Selection
                src_rng = selectedRange
                TB_src_rng.Text = selectedRange.Address
                TB_src_rng.Focus()
                TB_src_rng.SelectionStart = TB_src_rng.Text.Length
                'MsgBox(TB_src_rng.Text.Length)
            End If


        Catch ex As Exception
            TB_src_rng.Focus()
        End Try
    End Sub

    Private Sub excelApp_SheetSelectionChange(ByVal Sh As Object, ByVal selectionRange1 As Excel.Range) Handles excelApp.SheetSelectionChange
        Try

            excelApp = Globals.ThisAddIn.Application
            If focuschange = False Then


                If Me.ActiveControl Is TB_src_rng Then
                    src_rng = selectionRange1
                    Me.Activate()


                    Me.BeginInvoke(New System.Action(Sub()
                                                         TB_src_rng.Text = src_rng.Address
                                                         SetWindowPos(Me.Handle, New IntPtr(HWND_TOPMOST), 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOSIZE)
                                                     End Sub))
                    TB_src_rng.Focus()
                    TB_src_rng.SelectionStart = TB_src_rng.Text.Length
                    'TB_src_rng.Focus()
                End If



            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub CB_Source_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CB_Source.SelectedIndexChanged
        If CB_Source.Text = "Select Range" Then
            TB_src_rng.Enabled = True
            Selection_source.Enabled = True
        Else
            TB_src_rng.Enabled = False
            Selection_source.Enabled = False
        End If
    End Sub

    Private Sub Selection_source_Click(sender As Object, e As EventArgs) Handles Selection_source.Click
        Try
            ' If selectedRange Is Nothing Then
            'MsgBox(1)
            'Else

            'TB_src_rng.Text = selectedRange.Address


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

            TB_src_rng.Text = src_rng.Address

            Me.Show()
            TB_src_rng.Focus()


            'End If

        Catch ex As Exception

            Me.Show()
            TB_src_rng.Focus()

        End Try
    End Sub


    ' Event handler for when any sheet in the workbook is activated
    Private Sub WorkbookEvents_SheetActivate(ByVal Sh As Object) Handles WorkbookEvents.SheetActivate
        ' Detach event from previous sheet
        If CurrentSheet IsNot Nothing Then
            RemoveHandler CurrentSheet.SelectionChange, AddressOf sheet_SelectionChange
        End If

        ' Attach event to the new active sheet
        CurrentSheet = CType(Sh, Excel.Worksheet)
        AddHandler CurrentSheet.SelectionChange, AddressOf sheet_SelectionChange
        'MsgBox(CurrentSheet.Name)
    End Sub



    Private Sub Btn_OK_Click(sender As Object, e As EventArgs) Handles Btn_OK.Click
        excelApp = Globals.ThisAddIn.Application
        workBook = excelApp.ActiveWorkbook
        workSheet = workBook.ActiveSheet

        If CB_Source.Text = "Select Range" And TB_src_rng.Text = "" Then
            MessageBox.Show("Select a Source Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TB_src_rng.Focus()
            'Me.Close()
            Exit Sub

        ElseIf CB_Source.Text = "Select Range" And IsValidExcelCellReference(TB_src_rng.Text) = False Then
            MessageBox.Show("Select a Valid Source Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TB_src_rng.Focus()
            'Me.Close()
            Exit Sub

        Else

            If settingflag2 = True Then

                If RB_Horizontal.Checked = True Then
                    Horizontal2 = True
                Else
                    Horizontal2 = False
                End If

                If CB_Search.Checked = True Then
                    Search2 = True
                Else
                    Search2 = False
                End If

                Flag2 = True

                Separator2 = CB_Separator.Text
            Else
                If CB_Source.Text.Contains("Active Sheet") Then

                    src_rng = workSheet.Range("A1", workSheet.Cells(excelApp.Rows.Count, excelApp.Columns.Count))

                ElseIf CB_Source.Text.Contains("Active Workbook") Then

                    src_rng = workSheet.Range("A1", workSheet.Cells(excelApp.Rows.Count, excelApp.Columns.Count))

                End If

                If RB_Horizontal.Checked = True Then
                    Horizontal2 = True
                Else
                    Horizontal2 = False
                End If

                If CB_Search.Checked = True Then
                    Search2 = True
                Else
                    Search2 = False
                End If

                Flag2 = True
                GB_CB_Source2 = TB_src_rng.Text
                Separator2 = CB_Separator.Text

                'Private EventDel_CellsChange As Excel.DocEvents_ChangeEventHandler
                Dim i As Integer = 1

                workSheet.Range("B1").Select()  'Randomly select a cell. If nothing is selected, addhandler show error

                ' Define an array of type Excel.Worksheet
                Dim sheetsArray() As Excel.Worksheet

                ' Resize the array based on the number of sheets
                ReDim sheetsArray(workBook.Worksheets.Count)

                If CB_Source.Text.Contains("Active Workbook") Then
                    'MsgBox(1)
                    ' Assuming you're working with the active workbook:
                    AddHandler workSheet.SelectionChange, AddressOf sheet_SelectionChange
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


                ElseIf CB_Source.Text = "Select Range" Or CB_Source.Text.Contains("Active Sheet") Then

                    AddHandler workSheet.SelectionChange, AddressOf sheet_SelectionChange

                    'ElseIf CB_Source.Text.Contains("Active Sheet") Then

                    '    AddHandler workSheet.SelectionChange, AddressOf sheet_SelectionChange
                Else

                    For i = 1 To excelApp.ActiveWorkbook.Worksheets.Count
                        sheetsArray(i) = CType(workBook.Worksheets(i), Excel.Worksheet)
                        If CB_Source.Text = sheetsArray(i).Name Then
                            AddHandler sheetsArray(i).SelectionChange, AddressOf sheet_SelectionChange
                            'src_rng = sheetsArray(i).Range("A1", workSheet.Cells(excelApp.Rows.Count, excelApp.Columns.Count))
                        End If
                        'i = i + 1

                    Next
                    src_rng = workSheet.Range("A1", workSheet.Cells(excelApp.Rows.Count, excelApp.Columns.Count))

                End If

                'EventDel_CellsChange = New Excel.DocEvents_SelectionChangeEventHandler(AddressOf sheet_SelectionChange)
                'AddHandler xlSheet1.Change, EventDel_CellsChange


                If TB_src_rng.Enabled = True Then
                    GB_CB_Source2 = TB_src_rng.Text 'SR is the global variable for Source Range
                Else
                    GB_CB_Source2 = src_rng.Address
                End If

                GB_CB_Source2 = src_rng.Address

                RangeType2 = CB_Source.Text


                TType2 = ""
                SR2 = CB_Source.Text
                shName2 = workSheet.Name


            End If


            Dim workSheet2 As Excel.Worksheet = workBook.ActiveSheet



            ' Check if "Neww" worksheet exists and delete it if it does
            For Each ws In workBook.Sheets
                If ws.Name = "SoftekoSofteko" Then
                    ws.Delete()
                    Exit For
                End If
            Next

            ' Add a new worksheet named "Neww"
            workSheet2 = CType(workBook.Worksheets.Add(), Excel.Worksheet)
            workSheet2.Name = "SoftekoSofteko"

            ' Add your values (here's an example to set A1 to "Sample Value")
            workSheet2.Range("A1").Value = "Do not Delete thesheet!"
            ' ... Add more values as required ...



            ' Hide the worksheet
            workSheet2.Visible = Excel.XlSheetVisibility.xlSheetHidden

            workSheet2.Range("A2").Value = GB_CB_Source2   'For Source Range
            workSheet2.Range("A3").Value = SR2             '  Range Type
            workSheet2.Range("A4").Value = Horizontal2
            workSheet2.Range("A5").Value = Separator2
            workSheet2.Range("A6").Value = Search2
            workSheet2.Range("A7").Value = Flag2            'Activated
            workSheet2.Range("A8").Value = TargetVar2
            workSheet2.Range("A9").Value = shName2
            workSheet2.Range("A2").Value = GB_CB_Source2

            workSheet2.Range("B2").Value = CB_Source.Text

            Me.Close()

        End If
    End Sub

    Private Sub Btn_Cancel_Click(sender As Object, e As EventArgs) Handles Btn_Cancel.Click
        Me.Close()
    End Sub


    Private Sub sheet_SelectionChange(ByVal Target As Excel.Range)
        excelApp = Globals.ThisAddIn.Application
        workBook = excelApp.ActiveWorkbook
        workSheet = workBook.ActiveSheet

        Dim src_rng_concate As Excel.Range
        'MsgBox(workSheet.Name)
        'MsgBox(src_rng.Worksheet.Name)

        'src_rng = workSheet.Range(GB_CB_Source1)

        If GB_CB_Source2 IsNot Nothing Then
            src_rng = workSheet.Range(GB_CB_Source2)
            'MsgBox(src_rng.Address)

            'MsgBox(src_rng.Address)
            'src_rng = workSheet.Range(GB_CB_Source1)


            If CB_Source.Text.Contains("Active Workbook") Then
                src_rng = workSheet.Range("A1", workSheet.Cells(excelApp.Rows.Count, excelApp.Columns.Count))
            End If
            src_rng = workBook.ActiveSheet.range(src_rng.Address)
            ' MsgBox(src_rng.Worksheet.Name)

            'Change starts from here
            If (Nam2 = workSheet.Name And TType2 = "Select Range") Or (Nam2 = workSheet.Name And TType2.Contains("Active Sheet")) Or (Nam2 = workSheet.Name And TType2 = workSheet.Name) Then

                src_rng_concate = workSheet.Range(GB_CB_Dlt2)
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
                        If Form Is Nothing OrElse Form.IsDisposed Then
                            Form = New Form38()
                            Form.Show()
                            Form.BringToFront()
                            Form.Refresh()
                        Else
                            ' If form is already open, bring it to the front

                            Form.Dispose()
                            Form = New Form38()
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
                    TargetVar2 = Target.Address
                    If Form Is Nothing OrElse Form.IsDisposed Then
                        Form = New Form38()
                        Form.Show()
                        Form.BringToFront()
                        Form.Refresh()
                    Else
                        ' If form is already open, bring it to the front

                        Form.Dispose()
                        Form = New Form38()
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

    Private Sub CB_Separator_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CB_Separator.SelectedIndexChanged

    End Sub

    Private Sub CB_Separator_KeyUp(sender As Object, e As KeyEventArgs) Handles CB_Separator.KeyUp
        ' Check if the Enter key was pressed
        If e.KeyCode = Keys.Enter Then
            ' Add the current text in the ComboBox to the items collection
            If Not String.IsNullOrEmpty(CB_Separator.Text) AndAlso Not CB_Separator.Items.Contains(CB_Separator.Text) Then
                CB_Separator.Items.Add(CB_Separator.Text)
            End If


        End If
    End Sub

    Private Sub Form37_MSDropDownCheckBox_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Enter Then
            Btn_OK.PerformClick()
        End If
    End Sub

    Private Sub TB_src_rng_TextChanged(sender As Object, e As EventArgs) Handles TB_src_rng.TextChanged
        Try
            TB_src_rng.Focus()
            TB_src_rng.SelectionStart = TB_src_rng.Text.Length

            If TB_src_rng.Text IsNot Nothing And IsValidExcelCellReference(TB_src_rng.Text) = True Then
                focuschange = True

                ' Define the range of cells to read (for example, cells A1 to A10)
                src_rng = excelApp.Range(TB_src_rng.Text)
                src_rng.Select()
                Dim range As Excel.Range = src_rng

                Me.Activate()
                TB_src_rng.Focus()
                TB_src_rng.SelectionStart = TB_src_rng.Text.Length
                focuschange = False

            End If

        Catch ex As Exception
            TB_src_rng.Focus()

        End Try
    End Sub

    Private Function IsValidExcelCellReference(cellReference As String) As Boolean

        ' Regular expression pattern for a cell reference.
        ' This pattern will match references like A1, $A$1, etc.
        Dim cellPattern As String = "(\$?[A-Z]+\$?[0-9]+)"

        ' Regular expression pattern for an Excel reference.
        ' This pattern will match references like A1:B13, $A$1:$B$13, A1, $B$1, etc.
        Dim referencePattern As String = "^" + cellPattern + "(:" + cellPattern + ")?$"

        ' Create a regex object with the pattern.
        Dim regex As New Regex(referencePattern)

        ' Test the input string against the regex pattern.
        If regex.IsMatch(cellReference) Then
            Return True
        Else
            Return False
        End If


    End Function

    Private Sub Form37_MSDropDownCheckBox_Activated(sender As Object, e As EventArgs) Handles MyBase.Activated
        TB_src_rng.Focus()
        TB_src_rng.SelectionStart = TB_src_rng.Text.Length
    End Sub

    Private Sub Form37_MSDropDownCheckBox_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        TB_src_rng.Focus()
    End Sub
End Class

Public Class WorksheetHandler
    Public WithEvents Sheet As Excel.Worksheet

    Public Sub New(ByRef ws As Excel.Worksheet)
        Sheet = ws
    End Sub

    Private Sub Worksheet_SelectionChange(ByVal Target As Excel.Range) Handles Sheet.SelectionChange
        ' The event code goes here
    End Sub
End Class
