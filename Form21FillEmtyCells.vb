Imports System.ComponentModel.Design
Imports System.Security.Cryptography
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Excel
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Drawing
Imports System.ComponentModel
Imports System.Linq.Expressions



Public Class Form21FillEmtyCells

    Dim WithEvents excelApp As Excel.Application
    Dim workbook As Excel.Workbook
    Dim worksheet, worksheet1 As Excel.Worksheet
    Dim outWorksheet As Excel.Worksheet
    Dim inputRng As Excel.Range
    Dim FocusedTxtBox As Integer
    Dim selectedRange As Excel.Range
    Dim textChanged As Boolean = False


    Private Sub RB_Linear_values_CheckedChanged(sender As Object, e As EventArgs) Handles RB_Linear_values.CheckedChanged
        If RB_Linear_values.Checked = True Then
            ComboBox_Options.Items.Clear()
            ComboBox_Options.Items.Add("Top to Buttom")
            ComboBox_Options.Items.Add("Left to Right")
            ComboBox_Options.SelectedIndex = 0
            TextBox_Value.Enabled = False
            L_Fill_Value.Enabled = False
            ComboBox_Options.Enabled = True
            L_Fill_Options.Enabled = True

        End If

    End Sub

    Private Sub RB_Values_fromselected_range_CheckedChanged(sender As Object, e As EventArgs) Handles RB_Values_fromselected_range.CheckedChanged
        If RB_Values_fromselected_range.Checked = True Then
            ComboBox_Options.Items.Clear()
            ComboBox_Options.Items.Add("Downwards")
            ComboBox_Options.Items.Add("Upwards")
            ComboBox_Options.Items.Add("Towards the Right")
            ComboBox_Options.Items.Add("Towards the Left")
            ComboBox_Options.SelectedIndex = 0
            TextBox_Value.Enabled = False
            L_Fill_Value.Enabled = False
            ComboBox_Options.Enabled = True
            L_Fill_Options.Enabled = True

        End If
    End Sub

    Private Sub RB_Certain_value_CheckedChanged(sender As Object, e As EventArgs) Handles RB_Certain_value.CheckedChanged
        If RB_Certain_value.Checked = True Then
            ComboBox_Options.Items.Clear()
            ComboBox_Options.SelectedItem = ""
            TextBox_Value.Enabled = True
            L_Fill_Value.Enabled = True
            ComboBox_Options.Enabled = False
            L_Fill_Options.Enabled = False
        End If
    End Sub


    Private Sub Form21FillEmtyCells_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Try

            excelApp = Globals.ThisAddIn.Application
            Workbook = excelApp.ActiveWorkbook
            Worksheet = Workbook.ActiveSheet

            Dim selectedRng As Excel.Range = excelApp.Selection
            txtSourceRange.Text = selectedRng.Address


        Catch ex As Exception

        End Try


    End Sub

    Private Sub Selection_Click(sender As Object, e As EventArgs) Handles Selection.Click


        Try

            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            worksheet = workbook.ActiveSheet
            selectedRange = excelApp.Selection
            txtSourceRange.Focus()

            Me.Hide()
            inputRng = excelApp.InputBox("Please Select a Range", "Range Selection", selectedRange.Address, Type:=8)
            Me.Show()

            inputRng.Worksheet.Activate()

            txtSourceRange.Text = inputRng.Address
            inputRng.Select()


        Catch ex As Exception

            txtSourceRange.Focus()

        End Try




    End Sub

    Private Sub AutoSelection_Click(sender As Object, e As EventArgs) Handles AutoSelection.Click


        Try

            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            worksheet = workbook.ActiveSheet
            selectedRange = excelApp.Selection
            selectedRange = selectedRange.Cells(1, 1)
            selectedRange.Select()

            Dim topLeft, bottomRight As String


            'firstrow first column
            If selectedRange.Cells(1, 1).row = 1 And selectedRange.Cells(1, 1).column = 1 Then
                MsgBox("1")


                bottomRight = selectedRange.End(XlDirection.xlToRight).Address
                bottomRight = worksheet.Range(bottomRight).End(XlDirection.xlDown).Address

                selectedRange = worksheet.Range(selectedRange, worksheet.Range(bottomRight))

            ElseIf selectedRange.Cells(1, 1).row = 1 And selectedRange.Cells(1, 1).column <> 1 Then
                MsgBox("y")
            ElseIf selectedRange.Cells(1, 1).row <> 1 And selectedRange.Cells(1, 1).column = 1 Then
                MsgBox("z")
            Else




                If selectedRange.Offset(0, -1).Value = Nothing And selectedRange.Offset(0, 1).Value = Nothing And selectedRange.Offset(-1, 0).Value = Nothing Then
                    topLeft = selectedRange.Address
                    bottomRight = worksheet.Range(topLeft).End(XlDirection.xlDown).Address
                    selectedRange = worksheet.Range(worksheet.Range(topLeft), worksheet.Range(bottomRight))

                ElseIf selectedRange.Offset(-1, 0).Value = Nothing And selectedRange.Offset(1, 0).Value = Nothing And selectedRange.Offset(0, -1).Value = Nothing Then
                    topLeft = selectedRange.Address
                    bottomRight = worksheet.Range(topLeft).End(XlDirection.xlToRight).Address
                    selectedRange = worksheet.Range(worksheet.Range(topLeft), worksheet.Range(bottomRight))

                ElseIf selectedRange.Offset(0, -1).Value = Nothing And selectedRange.Offset(-1, 0).Value = Nothing Then
                    bottomRight = selectedRange.End(XlDirection.xlToRight).Address
                    bottomRight = worksheet.Range(bottomRight).End(XlDirection.xlDown).Address

                    selectedRange = worksheet.Range(selectedRange, worksheet.Range(bottomRight))

                ElseIf selectedRange.Offset(0, -1).Value = Nothing And selectedRange.Offset(0, 1).Value = Nothing Then
                    topLeft = selectedRange.End(XlDirection.xlUp).Address
                    bottomRight = worksheet.Range(topLeft).End(XlDirection.xlDown).Address
                    selectedRange = worksheet.Range(worksheet.Range(topLeft), worksheet.Range(bottomRight))

                ElseIf selectedRange.Offset(-1, 0).Value = Nothing And selectedRange.Offset(1, 0).Value = Nothing Then
                    topLeft = selectedRange.End(XlDirection.xlToLeft).Address
                    bottomRight = worksheet.Range(topLeft).End(XlDirection.xlToRight).Address
                    selectedRange = worksheet.Range(worksheet.Range(topLeft), worksheet.Range(bottomRight))

                ElseIf selectedRange.Offset(0, -1).Value = Nothing Then
                    topLeft = selectedRange.End(XlDirection.xlUp).Address
                    bottomRight = worksheet.Range(topLeft).End(XlDirection.xlToRight).Address
                    bottomRight = worksheet.Range(bottomRight).End(XlDirection.xlDown).Address
                    selectedRange = worksheet.Range(worksheet.Range(topLeft), worksheet.Range(bottomRight))


                ElseIf selectedRange.Offset(-1, 0).Value = Nothing Then
                    topLeft = selectedRange.End(XlDirection.xlToLeft).Address
                    bottomRight = worksheet.Range(topLeft).End(XlDirection.xlToRight).Address
                    bottomRight = worksheet.Range(bottomRight).End(XlDirection.xlDown).Address
                    selectedRange = worksheet.Range(worksheet.Range(topLeft), worksheet.Range(bottomRight))



                Else
                    topLeft = selectedRange.End(XlDirection.xlToLeft).Address
                    topLeft = worksheet.Range(topLeft).End(XlDirection.xlUp).Address
                    bottomRight = worksheet.Range(topLeft).End(XlDirection.xlToRight).Address
                    bottomRight = worksheet.Range(bottomRight).End(XlDirection.xlDown).Address

                    selectedRange = worksheet.Range(worksheet.Range(topLeft), worksheet.Range(bottomRight))


                End If


            End If


            selectedRange.Select()





        Catch ex As Exception

        End Try



    End Sub

    Private Sub Textbox1_TextChanged(sender As Object, e As EventArgs) Handles txtSourceRange.TextChanged

        Try
            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            worksheet = workbook.ActiveSheet

            textChanged = True

            inputRng = worksheet.Range(txtSourceRange.Text)
            inputRng.Select()


        Catch ex As Exception

        End Try

        textChanged = False
        txtSourceRange.Focus()



    End Sub



    Private Sub txtSourceRange_GotFocus(sender As Object, e As EventArgs) Handles txtSourceRange.GotFocus
        Try

            FocusedTxtBox = 1


        Catch ex As Exception

        End Try
    End Sub
    Private Sub Form1_Activated(sender As Object, e As EventArgs) Handles Me.Activated
        Try

            excelApp = Globals.ThisAddIn.Application

            AddHandler excelApp.SheetSelectionChange, AddressOf rngSelectionFromTxtBox

        Catch ex As Exception

        End Try

    End Sub

    Private Sub rngSelectionFromTxtBox(ByVal Sh As Object, ByVal Target As Excel.Range)

        Try

            excelApp = Globals.ThisAddIn.Application
            selectedRange = excelApp.Selection
            selectedRange.Select()

            txtSourceRange.Focus()

            If textChanged = False Then

                If FocusedTxtBox = 1 Then

                    txtSourceRange.Text = selectedRange.Address
                    worksheet = workbook.ActiveSheet
                    inputRng = selectedRange
                    txtSourceRange.Focus()

                End If

            End If



        Catch ex As Exception


        End Try

    End Sub



    Private Sub btn_OK_Click(sender As Object, e As EventArgs) Handles btn_OK.Click



    End Sub


    Private Sub btn_Cancel_Click(sender As Object, e As EventArgs) Handles btn_Cancel.Click

        Me.Dispose()

    End Sub

End Class