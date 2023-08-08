Imports System.Threading
Imports System.Windows.Forms
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop.Excel
Imports System.Reflection.Emit
Imports System.Linq
Imports System.Media

Public Class Form30_Create_Dynamic_Drop_down_List

    Dim WithEvents excelApp As Excel.Application
    Dim workBook As Excel.Workbook
    Dim workSheet As Excel.Worksheet
    Dim workSheet2 As Excel.Worksheet
    Dim src_rng As Excel.Range
    Public des_rng As Excel.Range
    Dim selectedRange As Excel.Range

    Dim opened As Integer


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

    Private Sub RB_columns_CheckedChanged(sender As Object, e As EventArgs) Handles RB_2_5_levels.CheckedChanged
        If RB_2_5_levels.Checked = True Then

            CB_header.Enabled = True
            CB_ascending.Enabled = True
            CB_descending.Enabled = True
            CB_text.Enabled = True
            GB_list_option.Enabled = False

        End If
    End Sub

    Private Sub RB_rows_CheckedChanged(sender As Object, e As EventArgs) Handles RB_2_levels.CheckedChanged
        If RB_2_levels.Checked = True Then
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

    Private Sub Btn_OK_Click(sender As Object, e As EventArgs) Handles Btn_OK.Click
        ' Dim worksheet As Excel.Worksheet
        excelApp = Globals.ThisAddIn.Application
        Dim rng1 As Excel.Range = excelApp.Range("A12", "A27")
        Dim destRange As Excel.Range = excelApp.Range("Z100", "Z104")
        ' Dim destRange As Excel.Range = xlWorksheet.Range("Z1:Z100")
        src_rng.Columns(1).AdvancedFilter(Action:=Excel.XlFilterAction.xlFilterCopy, CopyToRange:=destRange, Unique:=True)
        '=UNIQUE(FILTER(B12:B27, A12: A27 = I24))
        ' Get the filter criteria from cell I24
        Dim filterCriteria As String = destRange(1, 1).Value
        Dim filterCriteria2 As String = destRange(1, 2).value
        ' Create a new range for filtered values
        Dim filteredRange As Excel.Range = src_rng.Columns(2).SpecialCells(Excel.XlCellType.xlCellTypeConstants, Type.Missing).
                                    Offset(, -1).Resize(src_rng.Columns(2).SpecialCells(Excel.XlCellType.xlCellTypeConstants).Count)

        ' Loop through the column A values and check if they match the filter criteria
        For Each cell In src_rng.Columns(1).Cells
            If cell.Value = filterCriteria Then
                Dim valueToAdd As Object = cell.Offset(, 1).Value
                filteredRange.Value = valueToAdd
                filteredRange = filteredRange.Offset(1)
            End If
        Next


        ' Create a new range for filtered values
        Dim filteredRange2 As Excel.Range = src_rng.Columns(3).SpecialCells(Excel.XlCellType.xlCellTypeConstants, Type.Missing).
                                    Offset(, -2).Resize(src_rng.Columns(2).SpecialCells(Excel.XlCellType.xlCellTypeConstants).Count)

        ' Loop through the values in column A and check if they match the filter criteria
        For Each cellA In src_rng.Columns(1).Cells
            Dim cellB As Excel.Range = src_rng.Columns(2).Cells(cellA.Row - src_rng.Columns(1).Row + 1)
            If cellA.Value = filterCriteria And cellB.Value = filterCriteria2 Then
                Dim valueToAdd As Object = cellA.Offset(, 2).Value
                filteredRange2.Value = valueToAdd
                filteredRange2 = filteredRange2.Offset(1)
            End If
        Next

        ' Define the cell reference
        'Dim cellI24 As Excel.Range = des_rng

        ' Delete any existing validation
        des_rng.Validation.Delete()

        ' Set new validation
        des_rng.Validation.Add(Excel.XlDVType.xlValidateList,
                              Excel.XlDVAlertStyle.xlValidAlertStop,
                              Excel.XlFormatConditionOperator.xlBetween,
                              Formula1:=destRange(1, 1) & "#")

        ' Configure additional validation settings
        With des_rng.Validation
            .IgnoreBlank = True
            .InCellDropdown = True
            .InputTitle = ""
            .ErrorTitle = ""
            .InputMessage = ""
            .ErrorMessage = ""
            .ShowInput = True
            .ShowError = True
        End With

        ' Delete any existing validation
        des_rng.Columns(2).Validation.Delete()

        ' Set new validation
        des_rng.Columns(2).Validation.Add(Excel.XlDVType.xlValidateList,
                              Excel.XlDVAlertStyle.xlValidAlertStop,
                              Excel.XlFormatConditionOperator.xlBetween,
                              Formula1:=destRange(1, 2) & "#")

        ' Configure additional validation settings
        With des_rng.Columns(2).Validation
            .IgnoreBlank = True
            .InCellDropdown = True
            .InputTitle = ""
            .ErrorTitle = ""
            .InputMessage = ""
            .ErrorMessage = ""
            .ShowInput = True
            .ShowError = True
        End With

        ' Delete any existing validation
        des_rng.Columns(3).Validation.Delete()

        ' Set new validation
        des_rng.Columns(3).Validation.Add(Excel.XlDVType.xlValidateList,
                              Excel.XlDVAlertStyle.xlValidAlertStop,
                              Excel.XlFormatConditionOperator.xlBetween,
                              Formula1:=destRange(1, 3) & "#")

        ' Configure additional validation settings
        With des_rng.Columns(3).Validation
            .IgnoreBlank = True
            .InCellDropdown = True
            .InputTitle = ""
            .ErrorTitle = ""
            .InputMessage = ""
            .ErrorMessage = ""
            .ShowInput = True
            .ShowError = True
        End With

        Me.Close()
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


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

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

            If Me.ActiveControl Is TB_dest_range Then
                'des_rng = selectionRange1
                ' This will run on the Excel thread, so you need to use Invoke to update the UI
                'Me.BeginInvoke(New System.Action(Sub() TB_dest_range.Text = selectionRange1.Address))
                'Me.Activate()
                'Me.BeginInvoke(New System.Action(Sub()

            ElseIf Me.ActiveControl Is TB_src_range Then
                'src_rng = selectionRange1
                'workSheet = workBook.ActiveSheet
                'TB_src_range.Text = src_rng.Address
                'TB_src_range.Focus()
                'Me.Activate()
                'ActiveForm.Select()


                'Me.BeginInvoke(New System.Action(Sub()
                'TB_src_range.Text = src_rng.Address
                ' SetWindowPos(Me.Handle, New IntPtr(HWND_TOPMOST), 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOSIZE)
                'End Sub))
            End If



        Catch ex As Exception

        End Try

    End Sub

    Private Sub Btn_Cancel_Click(sender As Object, e As EventArgs) Handles Btn_Cancel.Click
        Me.Close()
    End Sub
End Class