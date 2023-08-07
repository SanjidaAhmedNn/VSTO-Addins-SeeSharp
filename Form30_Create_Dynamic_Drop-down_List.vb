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

        ' Use AdvancedFilter to copy unique values to the target range.
        'Dim xlApp.CutCopyMode As Excel.application = Excel.XlCutCopyMode.xlCopy
        src_rng.AdvancedFilter(Action:=Excel.XlFilterAction.xlFilterCopy, CopyToRange:=des_rng, Unique:=True)

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
                des_rng = selectionRange1
                ' This will run on the Excel thread, so you need to use Invoke to update the UI
                'Me.BeginInvoke(New System.Action(Sub() TB_dest_range.Text = selectionRange1.Address))
                'Me.Activate()
                'Me.BeginInvoke(New System.Action(Sub()

            ElseIf Me.ActiveControl Is TB_src_range Then
                src_rng = selectionRange1
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