Imports System.ComponentModel.Design
Imports System.Security.Cryptography
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Excel
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Drawing
Imports System.ComponentModel
Imports System.Linq.Expressions
Imports System.Threading
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports Microsoft.VisualBasic

Public Class Form11SwapRanges
    Dim WithEvents excelApp As Excel.Application
    Dim workbook As Excel.Workbook
    Dim worksheet As Excel.Worksheet
    Dim worksheet1, worksheet2 As Excel.Worksheet
    Dim outWorksheet As Excel.Worksheet
    Dim firstInputRng As Excel.Range
    Dim secondInputRng As Excel.Range
    Dim FocusedTxtBox As Integer
    Dim selectedRange As Excel.Range
    Dim firstRngRows, firstRngCols As Integer
    Dim tempRng As Excel.Range
    Dim rng1_Address, rng2_Address As String
    Dim changeState As Boolean = False
    Dim txtChanged As Boolean = False


    Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As IntPtr, ByVal hWndInsertAfter As IntPtr, ByVal X As Integer, ByVal Y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal uFlags As UInteger) As Boolean
    Private Const SWP_NOMOVE As UInteger = &H2
    Private Const SWP_NOSIZE As UInteger = &H1
    Private Const SWP_NOACTIVATE As UInteger = &H10
    Private Const HWND_TOPMOST As Integer = -1

    Private Sub Form1_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Enter Then
            btnOK.PerformClick()
        End If
    End Sub

    Private Sub Form11SwapRanges_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        excelApp = Globals.ThisAddIn.Application
        workbook = excelApp.ActiveWorkbook
        worksheet = workbook.ActiveSheet

        Dim selectedRng As Excel.Range = excelApp.Selection
        txtSourceRange1.Text = selectedRng.Address
        txtSourceRange1.Focus()

        radBtnValues.Checked = True

        Me.KeyPreview = True


    End Sub

    Private Sub txtSourceRange1_TextChanged(sender As Object, e As EventArgs) Handles txtSourceRange1.TextChanged


        Try

            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            worksheet = workbook.ActiveSheet


            'MsgBox(txtSourceRange1.Text)
            txtChanged = True
            firstInputRng = worksheet.Range(txtSourceRange1.Text)


            lblSourceRng1.Text = "1st Source Range (" & firstInputRng.Rows.Count & " rows x " & firstInputRng.Columns.Count & " columns)"

            firstInputRng.Select()


            firstRngRows = worksheet.Range(txtSourceRange1.Text).Rows.Count
            firstRngCols = worksheet.Range(txtSourceRange1.Text).Columns.Count


            If changeState = True Then


                If secondInputRng.Worksheet.Name <> firstInputRng.Worksheet.Name Then

                    txtSourceRange2.Text = secondInputRng.Worksheet.Name & "!" & secondInputRng.Address

                End If


            End If


        Catch ex As Exception

        End Try

        txtChanged = False

        txtSourceRange1.Focus()
    End Sub

    Private Sub txtSourceRange2_TextChanged(sender As Object, e As EventArgs) Handles txtSourceRange2.TextChanged

        Try
            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            worksheet = workbook.ActiveSheet

            changeState = True

            txtChanged = True
            secondInputRng = worksheet.Range(txtSourceRange2.Text)

            lblSourceRng2.Text = "2nd Source Range (" & secondInputRng.Rows.Count & " rows x " & secondInputRng.Columns.Count & " columns)"



            secondInputRng.Select()


            If secondInputRng.Worksheet.Name <> firstInputRng.Worksheet.Name Then

                txtSourceRange2.Text = secondInputRng.Worksheet.Name & "!" & secondInputRng.Address

            End If


        Catch ex As Exception

        End Try

        txtChanged = False
        txtSourceRange2.Focus()


    End Sub

    Private Sub rngSelection1_Click(sender As Object, e As EventArgs) Handles rngSelection1.Click

        Try

            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            worksheet = workbook.ActiveSheet
            selectedRange = excelApp.Selection
            txtSourceRange1.Focus()

            Me.Hide()
            firstInputRng = excelApp.InputBox("Please Select the First Range", "First Range Selection", selectedRange.Address, Type:=8)
            Me.Show()



            'firstInputRng.Worksheet.Activate()


            txtSourceRange1.Text = firstInputRng.Worksheet.Name & "!" & firstInputRng.Address

            firstInputRng.Select()

            txtSourceRange1.Focus()



        Catch ex As Exception

            txtSourceRange1.Focus()

        End Try


    End Sub

    Private Sub rngSelection2_Click(sender As Object, e As EventArgs) Handles rngSelection2.Click

        Try
            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            worksheet = workbook.ActiveSheet
            selectedRange = excelApp.Selection
            txtSourceRange2.Focus()

            Me.Hide()
            secondInputRng = excelApp.InputBox("Please Select the Second Range", "Second Range Selection", selectedRange.Address, Type:=8)
            Me.Show()




            txtSourceRange2.Text = secondInputRng.Worksheet.Name & "!" & secondInputRng.Address

            secondInputRng.Select()
            txtSourceRange2.Focus()




        Catch ex As Exception

            txtSourceRange2.Focus()

        End Try


    End Sub



    Private Sub AutoSelection1_Click(sender As Object, e As EventArgs) Handles AutoSelection1.Click

        Try

            'excelApp = Globals.ThisAddIn.Application
            'workbook = excelApp.ActiveWorkbook
            'worksheet = workbook.ActiveSheet
            'selectedRange = excelApp.Selection
            'selectedRange = selectedRange.Cells(1, 1)
            'selectedRange.Select()

            'Dim topLeft, bottomRight As String



            'If selectedRange.Offset(0, -1).Value = Nothing And selectedRange.Offset(0, 1).Value = Nothing And selectedRange.Offset(-1, 0).Value = Nothing Then
            '    topLeft = selectedRange.Address
            '    bottomRight = worksheet.Range(topLeft).End(XlDirection.xlDown).Address
            '    selectedRange = worksheet.Range(worksheet.Range(topLeft), worksheet.Range(bottomRight))

            'ElseIf selectedRange.Offset(-1, 0).Value = Nothing And selectedRange.Offset(1, 0).Value = Nothing And selectedRange.Offset(0, -1).Value = Nothing Then

            '    topLeft = selectedRange.Address
            '    bottomRight = worksheet.Range(topLeft).End(XlDirection.xlToRight).Address
            '    selectedRange = worksheet.Range(worksheet.Range(topLeft), worksheet.Range(bottomRight))

            'ElseIf selectedRange.Offset(0, -1).Value = Nothing And selectedRange.Offset(-1, 0).Value = Nothing Then
            '    bottomRight = selectedRange.End(XlDirection.xlToRight).Address
            '    bottomRight = worksheet.Range(bottomRight).End(XlDirection.xlDown).Address

            '    selectedRange = worksheet.Range(selectedRange, worksheet.Range(bottomRight))

            'ElseIf selectedRange.Offset(0, -1).Value = Nothing And selectedRange.Offset(0, 1).Value = Nothing Then

            '    topLeft = selectedRange.End(XlDirection.xlUp).Address
            '    bottomRight = worksheet.Range(topLeft).End(XlDirection.xlDown).Address
            '    selectedRange = worksheet.Range(worksheet.Range(topLeft), worksheet.Range(bottomRight))

            'ElseIf selectedRange.Offset(-1, 0).Value = Nothing And selectedRange.Offset(1, 0).Value = Nothing Then
            '    topLeft = selectedRange.End(XlDirection.xlToLeft).Address
            '    bottomRight = worksheet.Range(topLeft).End(XlDirection.xlToRight).Address
            '    selectedRange = worksheet.Range(worksheet.Range(topLeft), worksheet.Range(bottomRight))

            'ElseIf selectedRange.Offset(0, -1).Value = Nothing Then
            '    topLeft = selectedRange.End(XlDirection.xlUp).Address
            '    bottomRight = worksheet.Range(topLeft).End(XlDirection.xlToRight).Address
            '    bottomRight = worksheet.Range(bottomRight).End(XlDirection.xlDown).Address
            '    selectedRange = worksheet.Range(worksheet.Range(topLeft), worksheet.Range(bottomRight))


            'ElseIf selectedRange.Offset(-1, 0).Value = Nothing Then

            '    topLeft = selectedRange.End(XlDirection.xlToLeft).Address
            '    bottomRight = worksheet.Range(topLeft).End(XlDirection.xlToRight).Address
            '    bottomRight = worksheet.Range(bottomRight).End(XlDirection.xlDown).Address
            '    selectedRange = worksheet.Range(worksheet.Range(topLeft), worksheet.Range(bottomRight))



            'Else
            '    topLeft = selectedRange.End(XlDirection.xlToLeft).Address
            '    topLeft = worksheet.Range(topLeft).End(XlDirection.xlUp).Address
            '    bottomRight = worksheet.Range(topLeft).End(XlDirection.xlToRight).Address
            '    bottomRight = worksheet.Range(bottomRight).End(XlDirection.xlDown).Address

            '    selectedRange = worksheet.Range(worksheet.Range(topLeft), worksheet.Range(bottomRight))


            'End If

            'selectedRange.Select()

            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            worksheet = workbook.ActiveSheet
            selectedRange = excelApp.Selection

            Dim activeRange As Excel.Range = excelApp.ActiveCell

            Dim startRow As Integer = activeRange.Row
            Dim startColumn As Integer = activeRange.Column
            Dim endRow As Integer = activeRange.Row
            Dim endColumn As Integer = activeRange.Column

            'Find the upper boundary
            Do While startRow > 1 AndAlso Not IsNothing(worksheet.Cells(startRow - 1, startColumn).Value)
                startRow -= 1
            Loop

            'Find the lower boundary
            Do While Not IsNothing(worksheet.Cells(endRow + 1, endColumn).Value)
                endRow += 1
            Loop

            'Find the left boundary
            Do While startColumn > 1 AndAlso Not IsNothing(worksheet.Cells(startRow, startColumn - 1).Value)
                startColumn -= 1
            Loop

            'Find the right boundary
            Do While Not IsNothing(worksheet.Cells(endRow, endColumn + 1).Value)
                endColumn += 1
            Loop

            'Select the determined range
            worksheet.Range(worksheet.Cells(startRow, startColumn), worksheet.Cells(endRow, endColumn)).Select()





            firstInputRng = selectedRange
            txtSourceRange1.Text = firstInputRng.Address

            firstRngRows = selectedRange.Rows.Count
            firstRngCols = selectedRange.Columns.Count



        Catch ex As Exception

        End Try


    End Sub

    Private Sub AutoSelection2_Click(sender As Object, e As EventArgs) Handles AutoSelection2.Click


        Dim firstCell As Excel.Range

        excelApp = Globals.ThisAddIn.Application
        workbook = excelApp.ActiveWorkbook
        worksheet = workbook.ActiveSheet
        selectedRange = excelApp.Selection
        selectedRange.Select()

        Dim bottomRight As String
        firstCell = selectedRange.Cells(1, 1)

        If selectedRange.Cells(1, 1).Offset(1, 0).Value = Nothing Then

            For i = 0 To firstRngCols - 1
                If selectedRange.Cells(1, 1).offset(0, i).value <> Nothing Then
                    selectedRange = worksheet.Range(selectedRange.Cells(1, 1), selectedRange.Cells(1, 1).Offset(0, i))
                End If
                selectedRange.Select()
            Next

        ElseIf selectedRange.Cells(1, 1).Offset(0, 1).Value = Nothing Then
            For i = 0 To firstRngRows - 1
                If selectedRange.Cells(1, 1).offset(i, 0).value <> Nothing Then
                    selectedRange = worksheet.Range(selectedRange.Cells(1, 1), selectedRange.Cells(1, 1).Offset(i, 0))
                End If
                selectedRange.Select()
            Next

        Else

            bottomRight = firstCell.End(XlDirection.xlToRight).Address
            bottomRight = worksheet.Range(bottomRight).End(XlDirection.xlDown).Address

            selectedRange = worksheet.Range(firstCell, worksheet.Range(bottomRight))

            If selectedRange.Rows.Count = 1 And selectedRange.Columns.Count >= firstRngCols Then
                selectedRange = worksheet.Range(selectedRange.Cells(1, 1), selectedRange.Cells(1, 1).Offset(0, firstRngCols - 1))
                selectedRange.Select()

            ElseIf selectedRange.Rows.Count = 1 And selectedRange.Columns.Count < firstRngCols Then
                selectedRange = worksheet.Range(selectedRange.Cells(1, 1), selectedRange.Cells(1, 1).Offset(0, selectedRange.Columns.Count - 1))
                selectedRange.Select()

            ElseIf selectedRange.Columns.Count = 1 And selectedRange.Rows.Count >= firstRngRows Then
                selectedRange = worksheet.Range(selectedRange.Cells(1, 1), selectedRange.Cells(1, 1).Offset(firstRngRows - 1, 0))
                selectedRange.Select()

            ElseIf selectedRange.Columns.Count = 1 And selectedRange.Rows.Count < firstRngRows Then
                selectedRange = worksheet.Range(selectedRange.Cells(1, 1), selectedRange.Cells(1, 1).Offset(selectedRange.Rows.Count - 1, 0))
                selectedRange.Select()


            Else
                bottomRight = firstCell.End(XlDirection.xlToRight).Address
                bottomRight = worksheet.Range(bottomRight).End(XlDirection.xlDown).Address

                selectedRange = worksheet.Range(firstCell, worksheet.Range(bottomRight))

                If selectedRange.Rows.Count = firstRngRows And selectedRange.Columns.Count = firstRngCols Then
                    firstCell = selectedRange.Cells(1, 1)
                    selectedRange = worksheet.Range(firstCell.Offset(0, 0), firstCell.Offset(firstRngRows - 1, firstRngCols - 1))
                    selectedRange.Select()

                ElseIf selectedRange.Rows.Count = firstRngRows And selectedRange.Columns.Count > firstRngCols Then
                    firstCell = selectedRange.Cells(1, 1)
                    selectedRange = worksheet.Range(firstCell.Offset(0, 0), firstCell.Offset(firstRngRows - 1, firstRngCols - 1))
                    selectedRange.Select()

                ElseIf selectedRange.Rows.Count = firstRngRows And selectedRange.Columns.Count < firstRngCols Then
                    firstCell = selectedRange.Cells(1, 1)
                    bottomRight = firstCell.End(XlDirection.xlToRight).Address
                    bottomRight = worksheet.Range(bottomRight).End(XlDirection.xlDown).Address

                    selectedRange = worksheet.Range(firstCell.Offset(0, 0), worksheet.Range(bottomRight))
                    selectedRange.Select()

                ElseIf selectedRange.Rows.Count > firstRngRows And selectedRange.Columns.Count = firstRngCols Then
                    firstCell = selectedRange.Cells(1, 1)
                    selectedRange = worksheet.Range(firstCell.Offset(0, 0), firstCell.Offset(firstRngRows - 1, firstRngCols - 1))
                    selectedRange.Select()

                ElseIf selectedRange.Rows.Count > firstRngRows And selectedRange.Columns.Count > firstRngCols Then
                    firstCell = selectedRange.Cells(1, 1)
                    selectedRange = worksheet.Range(firstCell.Offset(0, 0), firstCell.Offset(firstRngRows - 1, firstRngCols - 1))
                    selectedRange.Select()

                ElseIf selectedRange.Rows.Count > firstRngRows And selectedRange.Columns.Count < firstRngCols Then
                    firstCell = selectedRange.Cells(1, 1)
                    bottomRight = firstCell.End(XlDirection.xlToRight).Address
                    bottomRight = worksheet.Range(bottomRight).Offset(firstRngRows - 1, 0).Address

                    selectedRange = worksheet.Range(firstCell.Offset(0, 0), worksheet.Range(bottomRight))
                    selectedRange.Select()

                ElseIf selectedRange.Rows.Count < firstRngRows And selectedRange.Columns.Count = firstRngCols Then
                    firstCell = selectedRange.Cells(1, 1)
                    bottomRight = firstCell.End(XlDirection.xlToRight).Address
                    bottomRight = worksheet.Range(bottomRight).End(XlDirection.xlDown).Address

                    selectedRange = worksheet.Range(firstCell.Offset(0, 0), worksheet.Range(bottomRight))
                    selectedRange.Select()
                ElseIf selectedRange.Rows.Count < firstRngRows And selectedRange.Columns.Count > firstRngCols Then

                    firstCell = selectedRange.Cells(1, 1)
                    bottomRight = firstCell.Offset(0, firstRngCols - 1).Address
                    bottomRight = worksheet.Range(bottomRight).End(XlDirection.xlDown).Address

                    selectedRange = worksheet.Range(firstCell.Offset(0, 0), worksheet.Range(bottomRight))
                    selectedRange.Select()


                ElseIf selectedRange.Rows.Count < firstRngRows And selectedRange.Columns.Count < firstRngCols Then
                    firstCell = selectedRange.Cells(1, 1)
                    bottomRight = firstCell.End(XlDirection.xlToRight).Address
                    bottomRight = worksheet.Range(bottomRight).End(XlDirection.xlDown).Address

                    selectedRange = worksheet.Range(firstCell.Offset(0, 0), worksheet.Range(bottomRight))
                    selectedRange.Select()

                End If
            End If

        End If

        secondInputRng = selectedRange
        txtSourceRange2.Text = secondInputRng.Address


    End Sub


    Private Sub txtSourceRange1_GotFocus(sender As Object, e As EventArgs) Handles txtSourceRange1.GotFocus
        Try

            FocusedTxtBox = 1


        Catch ex As Exception

        End Try
    End Sub
    Private Sub txtSourceRange2_GotFocus(sender As Object, e As EventArgs) Handles txtSourceRange2.GotFocus
        Try

            FocusedTxtBox = 2


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
            worksheet = workbook.ActiveSheet
            selectedRange = excelApp.Selection
            selectedRange.Select()


            If txtChanged = False Then


                If FocusedTxtBox = 1 Then
                    txtSourceRange1.Text = selectedRange.Address
                    txtSourceRange1.Focus()

                ElseIf FocusedTxtBox = 2 Then
                    txtSourceRange2.Text = selectedRange.Address
                End If

            End If


        Catch ex As Exception

        End Try

    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click

        Me.Dispose()

    End Sub
    Public Function IsValidRng(input As String) As Boolean

        Dim pattern As String = "^(\$?[A-Z]+\$?[0-9]+(:\$?[A-Z]+\$?[0-9]+)?)(,\$?[A-Z]+\$?[0-9]+(:\$?[A-Z]+\$?[0-9]+)?)*$"
        Return System.Text.RegularExpressions.Regex.IsMatch(input, pattern)

    End Function

    Private Sub btnOK_Click(sender As Object, e As EventArgs) Handles btnOK.Click



        Try

            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            worksheet = workbook.ActiveSheet
            selectedRange = excelApp.Selection

            If txtSourceRange1.Text = "" And txtSourceRange2.Text = "" Then

                MsgBox("Please select the first and the second range.", MsgBoxStyle.Exclamation, "Error!")
                txtSourceRange1.Focus()
                Exit Sub
            ElseIf txtSourceRange1.Text = "" And txtSourceRange2.Text <> "" Then

                If IsValidRng(txtSourceRange2.Text.ToUpper) = True Then
                    MsgBox("Please select the first range.", MsgBoxStyle.Exclamation, "Error!")
                    txtSourceRange1.Focus()
                    Exit Sub
                Else
                    MsgBox("Please use a valid range in the 2nd Source Range.", MsgBoxStyle.Exclamation, "Error!")
                    txtSourceRange2.Text = ""
                    txtSourceRange2.Focus()
                    Exit Sub
                End If

            ElseIf txtSourceRange2.Text = "" And txtSourceRange1.Text <> "" Then
                If IsValidRng(txtSourceRange1.Text.ToUpper) = True Then
                    MsgBox("Please select the second range.", MsgBoxStyle.Exclamation, "Error!")
                    txtSourceRange2.Focus()
                    Exit Sub
                Else
                    MsgBox("Please use a valid range in the 1st Source Range.", MsgBoxStyle.Exclamation, "Error!")
                    txtSourceRange1.Text = ""
                    txtSourceRange1.Focus()
                    Exit Sub
                End If

            ElseIf txtSourceRange1.Text <> "" And txtSourceRange2.Text <> "" Then
                If IsValidRng(txtSourceRange1.Text.ToUpper) = False And IsValidRng(txtSourceRange2.Text.ToUpper) = True Then
                    MsgBox("Please use a valid range in the 1st Source Range.", MsgBoxStyle.Exclamation, "Error!")
                    txtSourceRange1.Text = ""
                    txtSourceRange1.Focus()
                    Exit Sub

                ElseIf IsValidRng(txtSourceRange1.Text.ToUpper) = True And IsValidRng(txtSourceRange2.Text.ToUpper) = False Then
                    MsgBox("Please use a valid range in the 2nd Source Range.", MsgBoxStyle.Exclamation, "Error!")
                    txtSourceRange2.Text = ""
                    txtSourceRange2.Focus()
                    Exit Sub
                ElseIf IsValidRng(txtSourceRange1.Text.ToUpper) = False And IsValidRng(txtSourceRange2.Text.ToUpper) = False Then
                    MsgBox("Please use a valid range in the Source Ranges.", MsgBoxStyle.Exclamation, "Error!")
                    txtSourceRange1.Text = ""
                    txtSourceRange2.Text = ""
                    txtSourceRange1.Focus()
                    Exit Sub

                End If
            End If

            If firstInputRng.Rows.Count <> secondInputRng.Rows.Count And firstInputRng.Columns.Count <> secondInputRng.Columns.Count Then

                MsgBox("You must use same number of rows and columns in both ranges.",, "Warning!")
                txtSourceRange2.Focus()
                Exit Sub

            ElseIf firstInputRng.Rows.Count <> secondInputRng.Rows.Count And firstInputRng.Columns.Count = secondInputRng.Columns.Count Then
                MsgBox("Please match the source range row size.",, "Warning!")
                txtSourceRange2.Focus()
                'Me.Dispose()
                Exit Sub
            ElseIf firstInputRng.Rows.Count = secondInputRng.Rows.Count And firstInputRng.Columns.Count <> secondInputRng.Columns.Count Then
                MsgBox("Please match the source range column size.",, "Warning!")
                txtSourceRange2.Focus()
                Exit Sub

            End If

            worksheet1 = workbook.Sheets(firstInputRng.Worksheet.Name)
            worksheet2 = workbook.Sheets(secondInputRng.Worksheet.Name)

            'firstInputRng = worksheet.Range(txtSourceRange1.Text)
            'secondInputRng = worksheet.Range(txtSourceRange2.Text)

            'MsgBox(worksheet1.Name)
            'MsgBox(worksheet2.Name)



            Dim temp As Object
            tempRng = worksheet1.Range("A10000")
            tempRng = worksheet1.Range(tempRng.Cells(1, 1).offset(0, 0), tempRng.Cells(1, 1).offset(firstInputRng.Rows.Count - 1, firstInputRng.Columns.Count - 1))



            If CB_CopyWs.Checked = True Then

                workbook.ActiveSheet.Copy(After:=workbook.Sheets(workbook.Sheets.Count))
                outWorksheet = workbook.Sheets(workbook.Sheets.Count)


                worksheet = workbook.Sheets(firstInputRng.Worksheet.Name)
                worksheet.Activate()


            End If

            If radBtnValues.Checked = True Then
                If CB_KeepFormatting.Checked = True Then

                    temp = firstInputRng.Value
                    firstInputRng.Value = secondInputRng.Value
                    secondInputRng.Value = temp

                    For i = 0 To firstInputRng.Rows.Count - 1
                        For j = 0 To firstInputRng.Columns.Count - 1


                            Call copyCell(tempRng.Cells(1, 1), i, j, worksheet1.Range(firstInputRng.Address).Cells(1, 1), i, j)
                            Call copyCell(worksheet1.Range(firstInputRng.Address).Cells(1, 1), i, j, worksheet2.Range(secondInputRng.Address).Cells(1, 1), i, j)
                            Call copyCell(worksheet2.Range(secondInputRng.Address).Cells(1, 1), i, j, tempRng.Cells(1, 1), i, j)

                        Next
                    Next
                    tempRng.Delete()


                Else
                    firstInputRng.ClearFormats()
                    secondInputRng.ClearFormats()

                    temp = firstInputRng.Value
                    firstInputRng.Value = secondInputRng.Value
                    secondInputRng.Value = temp

                End If
                worksheet1.Activate()
                firstInputRng.Select()



            ElseIf radBtnKeepRef.Checked = True Then
                Dim modifiedFormula1, modifiedFormula2 As String
                If CB_KeepFormatting.Checked = True Then
                    For i = 0 To firstInputRng.Rows.Count - 1
                        For j = 0 To firstInputRng.Columns.Count - 1

                            Call copyCell(tempRng.Cells(1, 1), i, j, worksheet1.Range(firstInputRng.Address).Cells(1, 1), i, j)
                            Call copyCell(worksheet1.Range(firstInputRng.Address).Cells(1, 1), i, j, worksheet2.Range(secondInputRng.Address).Cells(1, 1), i, j)
                            Call copyCell(worksheet2.Range(secondInputRng.Address).Cells(1, 1), i, j, tempRng.Cells(1, 1), i, j)

                            modifiedFormula1 = swapFormulaWithSheetName(worksheet1.Range(firstInputRng.Address).Cells(1, 1).offset(i, j).formula, worksheet1.Name)
                            modifiedFormula2 = swapFormulaWithSheetName(worksheet2.Range(secondInputRng.Address).Cells(1, 1).offset(i, j).formula, worksheet2.Name)
                            worksheet1.Range(firstInputRng.Address).Cells(1, 1).offset(i, j).formula = modifiedFormula2
                            worksheet2.Range(secondInputRng.Address).Cells(1, 1).offset(i, j).formula = modifiedFormula1

                        Next
                    Next
                    tempRng.Delete()

                Else
                    firstInputRng.ClearFormats()
                    secondInputRng.ClearFormats()

                    For i = 0 To firstInputRng.Rows.Count - 1
                        For j = 0 To firstInputRng.Columns.Count - 1

                            modifiedFormula1 = swapFormulaWithSheetName(worksheet1.Range(firstInputRng.Address).Cells(1, 1).offset(i, j).formula, worksheet1.Name)
                            modifiedFormula2 = swapFormulaWithSheetName(worksheet2.Range(secondInputRng.Address).Cells(1, 1).offset(i, j).formula, worksheet2.Name)
                            worksheet1.Range(firstInputRng.Address).Cells(1, 1).offset(i, j).formula = modifiedFormula2
                            worksheet2.Range(secondInputRng.Address).Cells(1, 1).offset(i, j).formula = modifiedFormula1

                        Next
                    Next

                End If

                worksheet1.Activate()
                firstInputRng.Select()


            ElseIf radBtnAdjustRef.Checked = True Then

                If CB_KeepFormatting.Checked = True Then
                    worksheet1.Range(firstInputRng.Address).Copy(tempRng)
                    worksheet2.Range(secondInputRng.Address).Copy(worksheet1.Range(firstInputRng.Address))
                    tempRng.Copy(worksheet2.Range(secondInputRng.Address))
                    tempRng.Delete()

                Else
                    firstInputRng.ClearFormats()
                    secondInputRng.ClearFormats()

                    worksheet1.Range(firstInputRng.Address).Copy(tempRng)
                    worksheet2.Range(secondInputRng.Address).Copy(worksheet1.Range(firstInputRng.Address))
                    tempRng.Copy(worksheet2.Range(secondInputRng.Address))

                    tempRng.Delete()

                End If
                worksheet1.Activate()
                firstInputRng.Select()

            End If

            Me.Dispose()


        Catch ex As Exception

        End Try


    End Sub
    Public Function swapFormulaWithSheetName(currentFormula As String, sheetName As String) As String
        Dim pattern As String = "\b([A-Z]+[0-9]+(:[A-Z]+[0-9]+)?)\b"
        Dim replacement As String = ""
        Dim charToFind As Char = " "c
        Dim index As Integer

        If changeState = True Then
            If worksheet2.Name <> worksheet1.Name Then
                index = sheetName.IndexOf(charToFind)
                If index >= 0 Then
                    replacement = "'" & sheetName & "'!$1"
                Else
                    replacement = sheetName & "!$1"
                End If

            Else
                replacement = "$1"

            End If



        End If

        Return System.Text.RegularExpressions.Regex.Replace(currentFormula, pattern, replacement)


    End Function

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

    Private Sub Form11SwapRanges_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        form_flag = False
    End Sub

    Private Sub Form11SwapRanges_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        Me.Focus()
        Me.BringToFront()
        Me.Activate()
        Me.BeginInvoke(New System.Action(Sub()
                                             txtSourceRange1.Text = firstInputRng.Address
                                             SetWindowPos(Me.Handle, New IntPtr(HWND_TOPMOST), 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOSIZE)
                                         End Sub))
    End Sub

    Private Sub Form11SwapRanges_Disposed(sender As Object, e As EventArgs) Handles Me.Disposed
        form_flag = False
    End Sub
End Class