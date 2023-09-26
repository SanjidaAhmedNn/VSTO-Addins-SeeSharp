Imports System.ComponentModel
Imports System.ComponentModel.Design
Imports System.Runtime.InteropServices
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports Microsoft.Office.Interop

Imports Microsoft.Office.Interop.Excel
Public Class Form31_UpdateDynamicDropdownList
    Dim WithEvents excelApp As Excel.Application
    Dim workBook As Excel.Workbook
    Public Shared workSheet As Excel.Worksheet
    Dim workSheet2 As Excel.Worksheet
    Dim src_rng As Excel.Range
    Public des_rng As Excel.Range
    Dim selectedRange As Excel.Range
    Public ax As String
    Public focuschange As Boolean

    Dim opened As Integer


    Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As IntPtr, ByVal hWndInsertAfter As IntPtr, ByVal X As Integer, ByVal Y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal uFlags As UInteger) As Boolean
    Private Const SWP_NOMOVE As UInteger = &H2
    Private Const SWP_NOSIZE As UInteger = &H1
    Private Const SWP_NOACTIVATE As UInteger = &H10
    Private Const HWND_TOPMOST As Integer = -1


    Private Sub Selection_source_Click(sender As Object, e As EventArgs) Handles Selection_source.Click
        Try
            If selectedRange Is Nothing Then
            Else

                TB_src_rng.Text = selectedRange.Address


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
                Variable1 = TB_src_rng.Text
                'MsgBox(Variable1)
                Me.Show()
                TB_src_rng.Focus()
            End If

        Catch ex As Exception

            Me.Show()
            TB_src_rng.Focus()

        End Try
    End Sub

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click
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

            TB_des_rng2.Text = des_rng.Address

            Me.Show()
            TB_des_rng2.Focus()

        End If
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Dim workbook As Excel.Workbook = excelApp.ActiveWorkbook
        'Dim worksheet As Excel.Worksheet = workbook.ActiveSheet
        Try

            excelApp = Globals.ThisAddIn.Application

            AddHandler excelApp.SheetSelectionChange, AddressOf excelApp_SheetSelectionChange

            opened = opened + 1

            If excelApp.Selection IsNot Nothing Then
                selectedRange = excelApp.Selection
                src_rng = selectedRange
                TB_src_rng.Text = selectedRange.Address
            Else
                selectedRange = excelApp.Range(Variable1)
                src_rng = selectedRange
                TB_src_rng.Text = selectedRange.Address

            End If


            If RB_same_source.Checked = True Then
                TB_des_rng1.Enabled = True
                TB_des_rng2.Enabled = False
                PictureBox3.Enabled = False
                PictureBox2.Enabled = False
                L_select.Enabled = False
                If Variable2 IsNot Nothing Then
                    TB_des_rng1.Text = Variable2
                    des_rng = excelApp.ActiveSheet.Range(TB_des_rng1.Text)
                End If

            ElseIf RB_diff_rng.Checked = True Then
                TB_des_rng1.Enabled = False
                TB_des_rng2.Enabled = True
                PictureBox3.Enabled = True
                PictureBox2.Enabled = True
                L_select.Enabled = True

            End If

        Catch ex As Exception

        End Try

    End Sub


    Private Sub excelApp_SheetSelectionChange(ByVal Sh As Object, ByVal selectionRange1 As Excel.Range) Handles excelApp.SheetSelectionChange
        Try

            excelApp = Globals.ThisAddIn.Application
            If focuschange = False Then

                If Me.ActiveControl Is TB_des_rng2 Then
                    des_rng = selectionRange1
                    ' This will run on the Excel thread, so you need to use Invoke to update the UI
                    'Me.BeginInvoke(New System.Action(Sub() TB_dest_range.Text = selectionRange1.Address))
                    Me.Activate()
                    Me.BeginInvoke(New System.Action(Sub()
                                                         TB_des_rng2.Text = des_rng.Address
                                                         SetWindowPos(Me.Handle, New IntPtr(HWND_TOPMOST), 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOSIZE)
                                                     End Sub))

                ElseIf Me.ActiveControl Is TB_src_rng Then
                    src_rng = selectionRange1
                    Me.Activate()


                    Me.BeginInvoke(New System.Action(Sub()
                                                         TB_src_rng.Text = src_rng.Address
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

    Private Sub Btn_OK_Click(sender As Object, e As EventArgs) Handles Btn_OK.Click

        Try
            Dim result As DialogResult = MessageBox.Show("The Original Source Range is :" & Variable1 & ". AND the Drop-down list is in :" & Variable2 & "Do you want to continue?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

            ' Check if the user clicked 'Yes'
            If result = DialogResult.Yes Then

                Variable1 = TB_src_rng.Text
                'For Each cell In des_rng
                '    cell.Validation.Delete()
                'Next
                'Btn_OK.PerformClick()

                OutPut()
                'Variable1 = TB_src_rng.Text

            End If
        Catch ex As Exception
            des_rng.Select()
        End Try

    End Sub

    Private Sub TB_dest_range_Enter(sender As Object, e As KeyEventArgs) Handles TB_des_rng2.KeyDown
        'If Enter key is pressed then check if the text is a valid address
        If IsValidExcelCellReference(TB_des_rng2.Text) = True And e.KeyCode = Keys.Enter Then
            des_rng = excelApp.Range(TB_des_rng2.Text)
            TB_des_rng2.Focus()
            des_rng.Select()

            Call Btn_OK_Click(sender, e)   'OK button click event called

            'MsgBox(des_rng.Address)
        ElseIf IsValidExcelCellReference(TB_des_rng2.Text) = False And e.KeyCode = Keys.Enter Then
            MessageBox.Show("Select the valid Destination Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TB_des_rng2.Text = ""
            TB_des_rng2.Focus()
            'Me.Close()
            Exit Sub
        End If
    End Sub

    Private Sub TB_src_range_Enter(sender As Object, e As KeyEventArgs) Handles TB_src_rng.KeyDown
        'If Enter key is pressed then check if the text is a valid address

        If IsValidExcelCellReference(TB_src_rng.Text) = True And e.KeyCode = Keys.Enter Then
            src_rng = excelApp.Range(TB_src_rng.Text)
            TB_src_rng.Focus()
            src_rng.Select()

            Call Btn_OK_Click(sender, e)   'OK button click event called

            'MsgBox(des_rng.Address)
        ElseIf IsValidExcelCellReference(TB_src_rng.Text) = False And e.KeyCode = Keys.Enter Then
            MessageBox.Show("Select the valid Source Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TB_src_rng.Text = ""
            TB_src_rng.Focus()
            'Me.Close()
            Exit Sub
        End If
    End Sub


    Sub CreateValidationList(cell As Excel.Range, listValues As String)
        With cell.Validation
            .Delete()
            .Add(Type:=Excel.XlDVType.xlValidateList, AlertStyle:=Excel.XlDVAlertStyle.xlValidAlertStop, Operator:=Excel.XlFormatConditionOperator.xlBetween, Formula1:=listValues)
            .ShowInput = True
            .ShowError = True
        End With
    End Sub
    Private Sub OutPut()

        Try
            Dim workbook As Excel.Workbook = excelApp.ActiveWorkbook
            Dim worksheet As Excel.Worksheet = workbook.ActiveSheet
            'MsgBox(src_rng.Address)
            ' MsgBox(des_rng.Address)

            If TB_src_rng.Text = "" Then
                MessageBox.Show("Select a Source Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                TB_src_rng.Focus()
                'Me.Close()
                Exit Sub
                'End If

            ElseIf IsValidExcelCellReference(TB_src_rng.Text) = False Then
                MessageBox.Show("Select a Valid Source Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                TB_src_rng.Focus()
                'Me.Close()
                Exit Sub
            ElseIf RB_diff_rng.Checked = False And RB_same_source.Checked = False Then
                MessageBox.Show("Select Destination Range", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                TB_des_rng2.Focus()
                'Me.Close()
                Exit Sub

            Else
                Dim range1 As Excel.Range = excelApp.Range(Variable2)
                If RB_diff_rng.Checked = True And range1.Address(1, 1) <> des_rng.Address(1, 1) Then

                    range1.Cut(des_rng)
                End If

                Variable1 = TB_src_rng.Text
                If RB_diff_rng.Checked = True Then
                    Variable2 = TB_des_rng2.Text
                End If
                des_rng.Select()


                Me.Close()
            End If
        Catch ex As Exception
            Me.Close()
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


    Private Sub RB_same_source_CheckedChanged(sender As Object, e As EventArgs) Handles RB_same_source.CheckedChanged
        excelApp = Globals.ThisAddIn.Application
        If RB_same_source.Checked = True Then

            TB_des_rng1.Enabled = True

            TB_des_rng2.Enabled = False
            PictureBox3.Enabled = False
            PictureBox2.Enabled = False
            L_select.Enabled = False
            'MsgBox(L_select.Enabled)
            If Variable2 IsNot Nothing Then
                TB_des_rng1.Text = Variable2
                'MsgBox(Variable2)
                des_rng = excelApp.Range(Variable2)
            End If
        End If
    End Sub

    Private Sub RB_diff_rng_CheckedChanged(sender As Object, e As EventArgs) Handles RB_diff_rng.CheckedChanged
        If RB_diff_rng.Checked = True Then
            TB_des_rng1.Enabled = False
            TB_des_rng2.Enabled = True
            PictureBox3.Enabled = True
            PictureBox2.Enabled = True
            L_select.Enabled = True

        End If

    End Sub

    Private Sub OK(sender As Object, e As KeyEventArgs) Handles Btn_OK.KeyDown

        Try
            If e.KeyCode = Keys.Enter Then

                Call Btn_OK_Click(sender, e)

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub Cancel(sender As Object, e As KeyEventArgs) Handles Btn_Cancel.KeyDown

        Try
            If e.KeyCode = Keys.Enter Then

                Call Btn_OK_Click(sender, e)

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub Form_load(sender As Object, e As KeyEventArgs) Handles Me.KeyDown

        Try
            If e.KeyCode = Keys.Enter Then

                Call Btn_OK_Click(sender, e)

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub RB_Different(sender As Object, e As KeyEventArgs) Handles RB_diff_rng.KeyDown

        Try
            If e.KeyCode = Keys.Enter Then

                Call Btn_OK_Click(sender, e)

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub RB_same(sender As Object, e As KeyEventArgs) Handles RB_same_source.KeyDown

        Try
            If e.KeyCode = Keys.Enter Then

                Call Btn_OK_Click(sender, e)

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub CustomGroupBox2_Enter(sender As Object, e As EventArgs) Handles CustomGroupBox2.Enter

    End Sub

    Private Sub TB_src_rng_TextChanged(sender As Object, e As EventArgs) Handles TB_src_rng.TextChanged
        If TB_src_rng.Text IsNot Nothing And IsValidExcelCellReference(TB_src_rng.Text) = True Then
            focuschange = True

            ' Define the range of cells to read (for example, cells A1 to A10)
            src_rng = excelApp.Range(TB_src_rng.Text)
            src_rng.Select()
            Dim range As Excel.Range = src_rng


            Me.Activate()
            'TB_src_range.Focus()
            TB_src_rng.SelectionStart = TB_src_rng.Text.Length
            focuschange = False

        End If
    End Sub

    Private Sub TB_des_rng2_TextChanged(sender As Object, e As EventArgs) Handles TB_des_rng2.TextChanged
        If TB_des_rng2.Text IsNot Nothing And IsValidExcelCellReference(TB_des_rng2.Text) = True Then
            focuschange = True

            ' Define the range of cells to read (for example, cells A1 to A10)
            TB_des_rng2 = excelApp.Range(TB_des_rng2.Text)
            src_rng.Select()
            Dim range As Excel.Range = des_rng


            Me.Activate()
            'TB_src_range.Focus()
            TB_des_rng2.SelectionStart = TB_des_rng2.Text.Length
            focuschange = False

        End If
    End Sub
End Class