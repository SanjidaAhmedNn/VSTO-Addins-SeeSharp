﻿Imports System.ComponentModel.Design
Imports System.Security.Cryptography
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Excel
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Drawing
Imports System.ComponentModel
Imports System.Linq.Expressions
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Button



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
            txtFillValue.Enabled = False
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
            txtFillValue.Enabled = False
            L_Fill_Value.Enabled = False
            ComboBox_Options.Enabled = True
            L_Fill_Options.Enabled = True

        End If
    End Sub

    Private Sub RB_Certain_value_CheckedChanged(sender As Object, e As EventArgs) Handles RB_Certain_value.CheckedChanged
        If RB_Certain_value.Checked = True Then
            ComboBox_Options.Items.Clear()
            ComboBox_Options.SelectedItem = ""
            txtFillValue.Enabled = True
            L_Fill_Value.Enabled = True
            ComboBox_Options.Enabled = False
            L_Fill_Options.Enabled = False
        End If
    End Sub


    Private Sub Form21FillEmtyCells_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Try

            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            worksheet = Workbook.ActiveSheet

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

    Public Function IsValidRng(input As String) As Boolean

        Dim pattern As String = "^(\$?[A-Z]+\$?[0-9]+(:\$?[A-Z]+\$?[0-9]+)?)(,\$?[A-Z]+\$?[0-9]+(:\$?[A-Z]+\$?[0-9]+)?)*$"
        Return System.Text.RegularExpressions.Regex.IsMatch(input, pattern)

    End Function


    Private Sub btn_OK_Click(sender As Object, e As EventArgs) Handles btn_OK.Click

        Try

            Dim inputWsName As String
            Dim fillValue As String
            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            worksheet = workbook.ActiveSheet
            selectedRange = excelApp.Selection
            inputWsName = worksheet.Name

            'checks if an empty source range is used or not
            'if it is blank then a warning msgbox will appear and give user another chance to enter source range
            'if it is not blank then it checks the used range is valid range or not by using IsValidRng() function
            'IsValidRng() function is a custom function (see line 200)
            'using invalid range will give a warning to user and give another chance to enter range correctly
            If txtSourceRange.Text = "" Then
                MsgBox("Please select the Source Range.", MsgBoxStyle.Exclamation, "Error!")
                txtSourceRange.Focus()
                Exit Sub
            ElseIf IsValidRng(txtSourceRange.Text.ToUpper) = False Then
                MsgBox("Please use a valid range.", MsgBoxStyle.Exclamation, "Error!")
                txtSourceRange.Text = ""
                txtSourceRange.Focus()
                Exit Sub
            End If

            'stores the text value of the textbox in "temp" variable to use it later
            'store the active worksheet into "worksheet1" variable
            Dim temp As String
            temp = txtSourceRange.Text
            worksheet1 = inputRng.Worksheet

            'if CB_Backup_Sheet is checked then this will copy the active sheet and reactivate the original worksheet
            'replace the text of the txtSourceRange textbox by "temp" variable
            If CB_Backup_Sheet.Checked = True Then

                workbook.ActiveSheet.Copy(After:=workbook.Sheets(workbook.Sheets.Count))
                outWorksheet = workbook.Sheets(workbook.Sheets.Count)

                worksheet1.Activate()
                txtSourceRange.Text = temp

            End If

            'RB_Values_fromselected_range with Downwards fill option 
            If RB_Values_fromselected_range.Checked = True Then

                If ComboBox_Options.SelectedIndex = 0 Then

                    'loops through the cells of the selected range column by column
                    For j = 1 To selectedRange.Columns.Count

                        'checks if the first cell of the column is blank or not
                        'if so then value of fillValue var will be blank
                        'if not, fillValue will be the value of the first cell
                        If selectedRange.Cells(1, j).value Is Nothing Then
                            fillValue = ""
                        Else
                            fillValue = selectedRange.Cells(1, j).value
                        End If


                        For i = 1 To selectedRange.Rows.Count

                            'checks if the current cell is blank or not. this condition only passes from 2nd row (i=2)
                            'if the current cell is not blank then, replace the value of fillValue by the cuurent cell value. Then it can be copied to the following cells if they are blank
                            If selectedRange.Cells(i, j).value Is Nothing And i > 1 Then

                                'checks if the CB_Keepformatting is checked
                                'if so then, copy the cell of the previous row and same column (i-1,j) and paste it in current cell. This will copy both the value and format
                                'if CB_Keepformatting is not checked then, cuurent cell's value will be the value of fillValue
                                If CB_Keepformatting.Checked = True Then
                                    selectedRange.Cells(i - 1, j).copy(selectedRange.Cells(i, j))
                                Else
                                    selectedRange.Cells(i, j).value = fillValue
                                End If

                            Else
                                fillValue = selectedRange.Cells(i, j).value
                            End If

                        Next
                    Next



                    'RB_Values_fromselected_range with Upwards fill option 
                ElseIf ComboBox_Options.SelectedIndex = 1 Then

                    Dim rowCount As Integer = selectedRange.Rows.Count

                    'loops through the cells of the selected range column by column
                    For j = 1 To selectedRange.Columns.Count

                        'checks if the last cell of the column is blank or not
                        'if so then value of fillValue var will be blank
                        'if not, fillValue will be the value of the last cell
                        If selectedRange.Cells(rowCount, j).value Is Nothing Then
                            fillValue = ""
                        Else
                            fillValue = selectedRange.Cells(rowCount, j).value
                        End If


                        For i = rowCount To 1 Step -1

                            'checks if the current cell is blank or not. this condition only passes from 2nd from last row (i < rowCount)
                            'if the current cell is not blank then, replace the value of fillValue by the cuurent cell value. Then it can be copied to the previous cells if they are blank
                            If selectedRange.Cells(i, j).value Is Nothing And i < rowCount Then

                                'checks if the CB_Keepformatting is checked
                                'if so then, copy the cell of the next row and same column (i+1,j) and paste it in current cell. This will copy both the value and format
                                'if CB_Keepformatting is not checked then, cuurent cell's value will be the value of fillValue
                                If CB_Keepformatting.Checked = True Then
                                    selectedRange.Cells(i + 1, j).copy(selectedRange.Cells(i, j))
                                Else
                                    selectedRange.Cells(i, j).value = fillValue
                                End If

                            Else
                                fillValue = selectedRange.Cells(i, j).value
                            End If

                        Next
                    Next


                    'RB_Values_fromselected_range with Towards Right fill option 
                ElseIf ComboBox_Options.SelectedIndex = 2 Then
                    'loops through the cells of the selected range row by row
                    For i = 1 To selectedRange.Rows.Count

                        'checks if the first cell of the row is blank or not
                        'if so then value of fillValue var will be blank
                        'if not, fillValue will be the value of the first cell
                        If selectedRange.Cells(i, 1).value Is Nothing Then
                            fillValue = ""
                        Else
                            fillValue = selectedRange.Cells(i, 1).value
                        End If


                        For j = 1 To selectedRange.Columns.Count

                            'checks if the current cell is blank or not. this condition only passes from 2nd column(j > 1)
                            'if the current cell is not blank then, replace the value of fillValue by the cuurent cell value. Then it can be copied to the previous cells if they are blank
                            If selectedRange.Cells(i, j).value Is Nothing And j > 1 Then

                                'checks if the CB_Keepformatting is checked
                                'if so then, copy the cell of the previous column and same row(i,j-1) and paste it in current cell. This will copy both the value and format
                                'if CB_Keepformatting is not checked then, cuurent cell's value will be the value of fillValue
                                If CB_Keepformatting.Checked = True Then
                                    selectedRange.Cells(i, j - 1).copy(selectedRange.Cells(i, j))
                                Else
                                    selectedRange.Cells(i, j).value = fillValue
                                End If

                            Else
                                fillValue = selectedRange.Cells(i, j).value
                            End If

                        Next
                    Next



                    'RB_Values_fromselected_range with Towards Left fill option 
                ElseIf ComboBox_Options.SelectedIndex = 3 Then

                    Dim colCount As Integer = selectedRange.Columns.Count

                    'loops through the cells of the selected range row by row
                    For i = 1 To selectedRange.Rows.Count

                        'checks if the last cell of the row is blank or not
                        'if so then value of fillValue var will be blank
                        'if not, fillValue will be the value of the last cell
                        If selectedRange.Cells(i, colCount).value Is Nothing Then
                            fillValue = ""
                        Else
                            fillValue = selectedRange.Cells(i, colCount).value
                        End If


                        For j = colCount To 1 Step -1

                            'checks if the current cell is blank or not. this condition only passes from 2nd last column(j < colCount)
                            'if the current cell is not blank then, replace the value of fillValue by the cuurent cell value. Then it can be copied to the previous cells if they are blank
                            If selectedRange.Cells(i, j).value Is Nothing And j < colCount Then

                                'checks if the CB_Keepformatting is checked
                                'if so then, copy the cell of the next column and same row(i,j+1) and paste it in current cell. This will copy both the value and format
                                'if CB_Keepformatting is not checked then, cuurent cell's value will be the value of fillValue
                                If CB_Keepformatting.Checked = True Then
                                    selectedRange.Cells(i, j + 1).copy(selectedRange.Cells(i, j))
                                Else
                                    selectedRange.Cells(i, j).value = fillValue
                                End If

                            Else
                                fillValue = selectedRange.Cells(i, j).value
                            End If

                        Next
                    Next





                End If

            ElseIf RB_Linear_values.Checked = True Then

                'code




                'RB_Certain_value selected
            ElseIf RB_Certain_value.Checked = True Then

                'checks if the an empty Fill Value is used or not
                'if so then, a warning msgbox will pop up and give user another chance to enter Fill Value
                If txtFillValue.Text = "" Then
                    MsgBox("Please enter a Fill Value.", MsgBoxStyle.Exclamation, "Error!")
                    txtFillValue.Focus()
                    Exit Sub
                End If

                'loops through each cell of the selected range
                For i = 1 To selectedRange.Rows.Count
                    For j = 1 To selectedRange.Columns.Count

                        'checks if the current cell is blank or not
                        'if so then, its cell value will be the specified Fill Value
                        If selectedRange.Cells(i, j).value Is Nothing Then
                            selectedRange.Cells(i, j).value = txtFillValue.Text
                        End If
                    Next
                Next


            End If


            Me.Dispose()


        Catch ex As Exception

        End Try

    End Sub


    Private Sub btn_Cancel_Click(sender As Object, e As EventArgs) Handles btn_Cancel.Click

        Me.Dispose()

    End Sub

    Public Sub copyCell(ByVal destRng As Range, ByVal destOff1 As Integer, ByVal destOff2 As Integer, ByVal srcRng As Range, ByVal srcOff1 As Integer, ByVal srcOff2 As Integer)

        destRng.Offset(destOff1, destOff2).Font.Name = srcRng.Offset(srcOff1, srcOff2).Font.Name
        destRng.Offset(destOff1, destOff2).Font.Size = srcRng.Offset(srcOff1, srcOff2).Font.Size
        destRng.Offset(destOff1, destOff2).Font.Color = srcRng.Offset(srcOff1, srcOff2).Font.Color
        destRng.Offset(destOff1, destOff2).NumberFormat = srcRng.Offset(srcOff1, srcOff2).NumberFormat
        destRng.Offset(destOff1, destOff2).Interior.Color = srcRng.Offset(srcOff1, srcOff2).Interior.Color

        'bold,italic,underline
        destRng.Offset(destOff1, destOff2).Font.FontStyle = srcRng.Offset(srcOff1, srcOff2).Font.FontStyle
        destRng.Offset(destOff1, destOff2).Font.Underline = srcRng.Offset(srcOff1, srcOff2).Font.Underline




        'border

        destRng.Offset(destOff1, destOff2).Borders.LineStyle = srcRng.Offset(srcOff1, srcOff2).Borders.LineStyle
        destRng.Offset(destOff1, destOff2).Borders.Weight = srcRng.Offset(srcOff1, srcOff2).Borders.Weight


        'value
        'destRng.Offset(destOff1, destOff2).Value = srcRng.Offset(srcOff1, srcOff2).Value

    End Sub


End Class