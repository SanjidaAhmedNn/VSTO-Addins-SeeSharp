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
Imports Microsoft.VisualBasic.Devices

Public Class Form15CompareCells
    Dim WithEvents excelApp As Excel.Application
    Dim workbook As Excel.Workbook
    Dim worksheet, worksheet1, worksheet2 As Excel.Worksheet
    Dim outWorksheet As Excel.Worksheet
    Dim firstInputRng As Excel.Range
    Dim secondInputRng As Excel.Range
    Dim FocusedTxtBox As Integer
    Dim selectedRange As Excel.Range
    Dim firstRngRows, firstRngCols As Integer
    Dim colorPick As DialogResult
    Dim count As Integer
    Dim rng1CellValue, rng2CellValue, WsName, coloredRng, rngKeyBoard, output As String
    Dim changeState As Boolean = False
    Dim txtChanged As Boolean = False

    Private Sub Form1_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Enter Then
            btnOK.PerformClick()
        End If
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
                    '    secondInputRng = worksheet.Range(Microsoft.VisualBasic.Right(txtSourceRange2.Text, Len(txtSourceRange2.Text) - txtSourceRange2.Text.IndexOf("!") - 1))
                    '    lblSourceRng2.Text = "2nd Source Range (" & secondInputRng.Rows.Count & " rows x " & secondInputRng.Columns.Count & " columns)"
                    'Else
                    '    txtSourceRange2.Text = secondInputRng.Address
                    '    lblSourceRng2.Text = "2nd Source Range (" & secondInputRng.Rows.Count & " rows x " & secondInputRng.Columns.Count & " columns)"
                End If


            End If





        Catch ex As Exception

        End Try

        Call Display()
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
                '    secondInputRng = worksheet.Range(Microsoft.VisualBasic.Right(txtSourceRange2.Text, Len(txtSourceRange2.Text) - txtSourceRange2.Text.IndexOf("!") - 1))
                '    lblSourceRng2.Text = "2nd Source Range (" & secondInputRng.Rows.Count & " rows x " & secondInputRng.Columns.Count & " columns)"
                'Else
                '    txtSourceRange2.Text = secondInputRng.Address
                '    lblSourceRng2.Text = "2nd Source Range (" & secondInputRng.Rows.Count & " rows x " & secondInputRng.Columns.Count & " columns)"

            End If








        Catch ex As Exception

        End Try

        Call Display()
        txtChanged = False
        txtSourceRange2.Focus()

    End Sub




    Private Sub Form15CompareCells_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        excelApp = Globals.ThisAddIn.Application
        workbook = excelApp.ActiveWorkbook
        worksheet = workbook.ActiveSheet

        Dim selectedRng As Excel.Range = excelApp.Selection

        txtSourceRange1.Focus()
        txtSourceRange1.Text = selectedRng.Address

        radBtnSameValues.Checked = True


        Me.KeyPreview = True



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

            'Call Display()

            'txtSourceRange1.Text = selectedRange.Worksheet.Name & "!" & selectedRange.Address




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
            'Call Display()

        Catch ex As Exception

        End Try
    End Sub
    Private Sub txtSourceRange2_GotFocus(sender As Object, e As EventArgs) Handles txtSourceRange2.GotFocus
        Try

            FocusedTxtBox = 2
            'Call Display()

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

                    'Call Display()
                    'worksheet = workbook.ActiveSheet
                    'firstInputRng = selectedRange
                    'firstInputRng.Select()
                    'firstInputRng = worksheet.Range(selectedRange.Address)

                    txtSourceRange1.Focus()

                ElseIf FocusedTxtBox = 2 Then
                    txtSourceRange2.Text = selectedRange.Address

                    'Call Display()
                    'worksheet = workbook.ActiveSheet
                    ''secondInputRng = selectedRange
                    'secondInputRng = worksheet.Range(txtSourceRange2.Text)

                    'txtSourceRange2.Focus()


                End If



            End If



        Catch ex As Exception


        End Try

    End Sub

    Private Sub btnCanecl_Click(sender As Object, e As EventArgs) Handles btnCanecl.Click
        Me.Dispose()
    End Sub
    Public Function IsValidRng(input As String) As Boolean
        Dim pattern As String = "^\$?[A-Z]+\$?[0-9]+(:\$?[A-Z]+\$?[0-9]+)?$"
        Return System.Text.RegularExpressions.Regex.IsMatch(input, pattern)
    End Function

    Private Sub btnOK_Click(sender As Object, e As EventArgs) Handles btnOK.Click




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

        excelApp = Globals.ThisAddIn.Application
        Dim i, j As Integer
        Dim rng1CellValue, rng2CellValue As String
        Dim coloredRng As String
        Dim temp As String

        worksheet1 = firstInputRng.Worksheet
        worksheet2 = secondInputRng.Worksheet


        count = 0
        coloredRng = ""
        temp = txtSourceRange2.Text


        If checkBoxCopyWs.Checked = True Then

            worksheet1.Copy(After:=workbook.Sheets(workbook.Sheets.Count))
            outWorksheet = workbook.Sheets(workbook.Sheets.Count)

            worksheet2.Activate()

            txtSourceRange2.Text = temp

        End If


        If checkBoxFormatting.Checked = False Then

            firstInputRng.ClearFormats()

        End If



        If radBtnSameValues.Checked = True Then
            If checkBoxCase.Checked = True Then

                '1st Range >> 2nd Range >> radbtnSameValues checked >> case sensitive checked >> fill/font color both are selected >> OK
                If checkBoxFillBack.Checked = True And checkBoxFillFont.Checked = True Then
                    For i = 1 To firstInputRng.Rows.Count
                        For j = 1 To firstInputRng.Columns.Count

                            'handles comparison of different type o variables
                            If VarType(firstInputRng.Cells(i, j).value) <> VarType(secondInputRng.Cells(i, j).value) Then
                                GoTo nextLoop1

                            ElseIf firstInputRng.Cells(i, j).value = secondInputRng.Cells(i, j).value Then

                                firstInputRng.Cells(i, j).Interior.Color = CBFillBackground.BackColor

                                firstInputRng.Cells(i, j).Font.Color = CbFillFont.BackColor
                                count = count + 1
                                coloredRng = coloredRng & "," & firstInputRng.Cells(i, j).address
                            End If
nextLoop1:
                        Next
                    Next

                    '1st Range >> 2nd Range >> radbtnSameValues checked >> case sensitive checked >> only fill color is selected >> OK
                ElseIf checkBoxFillBack.Checked = True And checkBoxFillFont.Checked = False Then


                    For i = 1 To firstInputRng.Rows.Count
                        For j = 1 To firstInputRng.Columns.Count

                            If VarType(firstInputRng.Cells(i, j).value) <> VarType(secondInputRng.Cells(i, j).value) Then
                                GoTo nextLoop2

                            ElseIf firstInputRng.Cells(i, j).value = secondInputRng.Cells(i, j).value Then

                                firstInputRng.Cells(i, j).Interior.Color = CBFillBackground.BackColor
                                count = count + 1
                                coloredRng = coloredRng & "," & firstInputRng.Cells(i, j).address
                            End If
nextLoop2:
                        Next
                    Next

                    '1st Range >> 2nd Range >> radbtnSameValues checked >> case sensitive checked >> only font color is selected >> OK
                ElseIf checkBoxFillBack.Checked = False And checkBoxFillFont.Checked = True Then

                    For i = 1 To firstInputRng.Rows.Count
                        For j = 1 To firstInputRng.Columns.Count

                            If VarType(firstInputRng.Cells(i, j).value) <> VarType(secondInputRng.Cells(i, j).value) Then
                                GoTo nextLoop3

                            ElseIf firstInputRng.Cells(i, j).value = secondInputRng.Cells(i, j).value Then

                                firstInputRng.Cells(i, j).Font.Color = CbFillFont.BackColor

                                count = count + 1
                                coloredRng = coloredRng & "," & firstInputRng.Cells(i, j).address
                            End If
nextLoop3:
                        Next
                    Next

                    '1st Range >> 2nd Range >> radbtnSameValues checked >> case sensitive checked >> fill/font color is not selected >> OK
                Else

                    For i = 1 To firstInputRng.Rows.Count
                        For j = 1 To firstInputRng.Columns.Count

                            'If variable type of two compared cell are different
                            If VarType(firstInputRng.Cells(i, j).value) <> VarType(secondInputRng.Cells(i, j).value) Then
                                GoTo nextLoop4

                            ElseIf firstInputRng.Cells(i, j).value = secondInputRng.Cells(i, j).value Then
                                count = count + 1
                                coloredRng = coloredRng & "," & firstInputRng.Cells(i, j).address
                            End If
nextLoop4:
                        Next
                    Next

                End If
            Else

                '1st Range >> 2nd Range >> radbtnSameValues checked >> case sensitive unchecked >> fill/font color both are selected >> OK
                If checkBoxFillBack.Checked = True And checkBoxFillFont.Checked = True Then
                    For i = 1 To firstInputRng.Rows.Count
                        For j = 1 To firstInputRng.Columns.Count
                            rng1CellValue = firstInputRng.Cells(i, j).value
                            rng2CellValue = secondInputRng.Cells(i, j).value

                            If VarType(firstInputRng.Cells(i, j).value) <> VarType(secondInputRng.Cells(i, j).value) Then
                                GoTo nextLoop5

                            ElseIf rng1CellValue.ToUpper = rng2CellValue.ToUpper Then

                                firstInputRng.Cells(i, j).Interior.Color = CBFillBackground.BackColor

                                firstInputRng.Cells(i, j).Font.Color = CbFillFont.BackColor
                                count = count + 1
                                coloredRng = coloredRng & "," & firstInputRng.Cells(i, j).address
                            End If
nextLoop5:
                        Next
                    Next

                    '1st Range >> 2nd Range >> radbtnSameValues checked >> case sensitive unchecked >> only fill color is selected >> OK
                ElseIf checkBoxFillBack.Checked = True And checkBoxFillFont.Checked = False Then
                    For i = 1 To firstInputRng.Rows.Count
                        For j = 1 To firstInputRng.Columns.Count
                            rng1CellValue = firstInputRng.Cells(i, j).value
                            rng2CellValue = secondInputRng.Cells(i, j).value

                            If VarType(firstInputRng.Cells(i, j).value) <> VarType(secondInputRng.Cells(i, j).value) Then
                                GoTo nextLoop6

                            ElseIf rng1CellValue.ToUpper = rng2CellValue.ToUpper Then

                                firstInputRng.Cells(i, j).Interior.Color = CBFillBackground.BackColor

                                count = count + 1
                                coloredRng = coloredRng & "," & firstInputRng.Cells(i, j).address
                            End If
nextLoop6:
                        Next
                    Next

                    '1st Range >> 2nd Range >> radbtnSameValues checked >> case sensitive unchecked >> only font color is selected >> OK
                ElseIf checkBoxFillBack.Checked = False And checkBoxFillFont.Checked = True Then
                    For i = 1 To firstInputRng.Rows.Count
                        For j = 1 To firstInputRng.Columns.Count
                            rng1CellValue = firstInputRng.Cells(i, j).value
                            rng2CellValue = secondInputRng.Cells(i, j).value

                            If VarType(firstInputRng.Cells(i, j).value) <> VarType(secondInputRng.Cells(i, j).value) Then
                                GoTo nextLoop7

                            ElseIf rng1CellValue.ToUpper = rng2CellValue.ToUpper Then

                                firstInputRng.Cells(i, j).Font.Color = CbFillFont.BackColor

                                count = count + 1
                                coloredRng = coloredRng & "," & firstInputRng.Cells(i, j).address
                            End If
nextLoop7:
                        Next
                    Next


                Else
                    '1st Range >> 2nd Range >> radbtnSameValues checked >> case sensitive unchecked >> fill/font color not selected >> OK
                    For i = 1 To firstInputRng.Rows.Count
                        For j = 1 To firstInputRng.Columns.Count
                            rng1CellValue = firstInputRng.Cells(i, j).value
                            rng2CellValue = secondInputRng.Cells(i, j).value
                            If VarType(firstInputRng.Cells(i, j).value) <> VarType(secondInputRng.Cells(i, j).value) Then
                                GoTo nextLoop8

                            ElseIf rng1CellValue.ToUpper = rng2CellValue.ToUpper Then
                                count = count + 1
                                coloredRng = coloredRng & "," & firstInputRng.Cells(i, j).address
                            End If
nextLoop8:
                        Next
                    Next


                End If
            End If

        ElseIf radBtnDifferentValues.Checked = True Then
            If checkBoxCase.Checked = True Then

                '1st Range >> 2nd Range >> radBtnDifferentValues checked >> case sensitive checked >> fill/font color both are selected >> OK
                If checkBoxFillBack.Checked = True And checkBoxFillFont.Checked = True Then
                    For i = 1 To firstInputRng.Rows.Count
                        For j = 1 To firstInputRng.Columns.Count

                            If VarType(firstInputRng.Cells(i, j).value) <> VarType(secondInputRng.Cells(i, j).value) Then
                                GoTo nextLoop9

                            ElseIf firstInputRng.Cells(i, j).value <> secondInputRng.Cells(i, j).value Then
nextLoop9:
                                firstInputRng.Cells(i, j).Interior.Color = CBFillBackground.BackColor
                                firstInputRng.Cells(i, j).Font.Color = CbFillFont.BackColor
                                count = count + 1
                                coloredRng = coloredRng & "," & firstInputRng.Cells(i, j).address

                            End If

                        Next
                    Next

                    '1st Range >> 2nd Range >> radBtnDifferentValues checked >> case sensitive checked >> only fill color is selected >> OK
                ElseIf checkBoxFillBack.Checked = True And checkBoxFillFont.Checked = False Then
                    For i = 1 To firstInputRng.Rows.Count
                        For j = 1 To firstInputRng.Columns.Count

                            If VarType(firstInputRng.Cells(i, j).value) <> VarType(secondInputRng.Cells(i, j).value) Then
                                GoTo nextLoop10

                            ElseIf firstInputRng.Cells(i, j).value <> secondInputRng.Cells(i, j).value Then
nextLoop10:
                                firstInputRng.Cells(i, j).Interior.Color = CBFillBackground.BackColor

                                count = count + 1
                                coloredRng = coloredRng & "," & firstInputRng.Cells(i, j).address

                            End If

                        Next
                    Next

                    '1st Range >> 2nd Range >> radBtnDifferentValues checked >> case sensitive checked >> only font color is selected >> OK
                ElseIf checkBoxFillBack.Checked = False And checkBoxFillFont.Checked = True Then
                    For i = 1 To firstInputRng.Rows.Count
                        For j = 1 To firstInputRng.Columns.Count

                            If VarType(firstInputRng.Cells(i, j).value) <> VarType(secondInputRng.Cells(i, j).value) Then
                                GoTo nextLoop11

                            ElseIf firstInputRng.Cells(i, j).value <> secondInputRng.Cells(i, j).value Then
nextLoop11:
                                firstInputRng.Cells(i, j).Font.Color = CbFillFont.BackColor

                                count = count + 1
                                coloredRng = coloredRng & "," & firstInputRng.Cells(i, j).address

                            End If
                        Next
                    Next
                Else
                    '1st Range >> 2nd Range >> radBtnDifferentValues checked >> case sensitive checked >> fill/font color not selected >> OK

                    For i = 1 To firstInputRng.Rows.Count
                        For j = 1 To firstInputRng.Columns.Count

                            If VarType(firstInputRng.Cells(i, j).value) <> VarType(secondInputRng.Cells(i, j).value) Then
                                GoTo nextLoop12

                            ElseIf firstInputRng.Cells(i, j).value <> secondInputRng.Cells(i, j).value Then
nextLoop12:
                                count = count + 1
                                coloredRng = coloredRng & "," & firstInputRng.Cells(i, j).address

                            End If

                        Next
                    Next


                End If
            Else

                '1st Range >> 2nd Range >> radBtnDifferentValues checked >> case sensitive unchecked >> fill/font color both selected >> OK
                If checkBoxFillBack.Checked = True And checkBoxFillFont.Checked = True Then
                    For i = 1 To firstInputRng.Rows.Count
                        For j = 1 To firstInputRng.Columns.Count
                            rng1CellValue = firstInputRng.Cells(i, j).value
                            rng2CellValue = secondInputRng.Cells(i, j).value

                            If VarType(firstInputRng.Cells(i, j).value) <> VarType(secondInputRng.Cells(i, j).value) Then
                                GoTo nextLoop13

                            ElseIf rng1CellValue.ToUpper <> rng2CellValue.ToUpper Then
nextLoop13:
                                firstInputRng.Cells(i, j).Interior.Color = CBFillBackground.BackColor

                                firstInputRng.Cells(i, j).Font.Color = CbFillFont.BackColor
                                count = count + 1
                                coloredRng = coloredRng & "," & firstInputRng.Cells(i, j).address

                            End If

                        Next
                    Next

                    '1st Range >> 2nd Range >> radBtnDifferentValues checked >> case sensitive unchecked >> only fill color is selected >> OK
                ElseIf checkBoxFillBack.Checked = True And checkBoxFillFont.Checked = False Then
                    For i = 1 To firstInputRng.Rows.Count
                        For j = 1 To firstInputRng.Columns.Count
                            rng1CellValue = firstInputRng.Cells(i, j).value
                            rng2CellValue = secondInputRng.Cells(i, j).value

                            If VarType(firstInputRng.Cells(i, j).value) <> VarType(secondInputRng.Cells(i, j).value) Then
                                GoTo nextLoop14

                            ElseIf rng1CellValue.ToUpper <> rng2CellValue.ToUpper Then
nextLoop14:
                                firstInputRng.Cells(i, j).Interior.Color = CBFillBackground.BackColor

                                count = count + 1
                                coloredRng = coloredRng & "," & firstInputRng.Cells(i, j).address

                            End If

                        Next
                    Next

                    '1st Range >> 2nd Range >> radBtnDifferentValues checked >> case sensitive unchecked >> only font color is selected >> OK
                ElseIf checkBoxFillBack.Checked = False And checkBoxFillFont.Checked = True Then
                    For i = 1 To firstInputRng.Rows.Count
                        For j = 1 To firstInputRng.Columns.Count
                            rng1CellValue = firstInputRng.Cells(i, j).value
                            rng2CellValue = secondInputRng.Cells(i, j).value

                            If VarType(firstInputRng.Cells(i, j).value) <> VarType(secondInputRng.Cells(i, j).value) Then
                                GoTo nextLoop15

                            ElseIf rng1CellValue.ToUpper <> rng2CellValue.ToUpper Then
nextLoop15:
                                firstInputRng.Cells(i, j).Font.Color = CbFillFont.BackColor
                                count = count + 1
                                coloredRng = coloredRng & "," & firstInputRng.Cells(i, j).address

                            End If

                        Next

                    Next



                Else
                    '1st Range >> 2nd Range >> radBtnDifferentValues checked >> case sensitive unchecked >> fill/font color not selected >> OK
                    For i = 1 To firstInputRng.Rows.Count
                        For j = 1 To firstInputRng.Columns.Count
                            rng1CellValue = firstInputRng.Cells(i, j).value
                            rng2CellValue = secondInputRng.Cells(i, j).value

                            If VarType(firstInputRng.Cells(i, j).value) <> VarType(secondInputRng.Cells(i, j).value) Then
                                GoTo nextLoop16

                            ElseIf rng1CellValue.ToUpper <> rng2CellValue.ToUpper Then
nextLoop16:
                                count = count + 1
                                coloredRng = coloredRng & "," & firstInputRng.Cells(i, j).address

                            End If

                        Next
                    Next

                End If
            End If

        End If

        Me.Dispose()


        firstInputRng.Worksheet.Activate()

        MsgBox(count & " cell(s) found.", MsgBoxStyle.Information, "SOFTEKO")
        If coloredRng = "" Then
            Exit Sub
        Else
            coloredRng = Microsoft.VisualBasic.Right(coloredRng, Len(coloredRng) - 1)
            firstInputRng.Worksheet.Range(coloredRng).Select()
        End If



    End Sub


    Sub Display()

        Try

            CP_Input_Range1.Controls.Clear()
            CP_Input_Range2.Controls.Clear()
            CP_Output_Range.Controls.Clear()


            Dim displayRng As Excel.Range
            Dim lblColor As Long
            Dim rgbColor As System.Drawing.Color

            If txtSourceRange1.Text = "" Or firstInputRng Is Nothing Then
                CP_Input_Range1.Controls.Clear()
                GoTo secondDisplay
            End If


            If firstInputRng.Rows.Count > 50 Then
                displayRng = firstInputRng.Rows("1:50")
            Else
                displayRng = firstInputRng
            End If


            Dim height As Double
            Dim width As Double

            If displayRng.Rows.Count <= 4 Then
                height = CP_Input_Range1.Height / displayRng.Rows.Count
            Else
                height = (119 / 4)
            End If

            If displayRng.Columns.Count <= 3 Then
                width = CP_Input_Range1.Width / displayRng.Columns.Count
            Else
                width = (260 / 3)
            End If

            For i = 1 To displayRng.Rows.Count
                For j = 1 To displayRng.Columns.Count
                    Dim label As New System.Windows.Forms.Label
                    label.Text = displayRng.Cells(i, j).Value
                    label.Location = New System.Drawing.Point((j - 1) * width, (i - 1) * height)
                    label.Height = height
                    label.Width = width
                    label.BorderStyle = BorderStyle.FixedSingle
                    label.TextAlign = ContentAlignment.MiddleCenter

                    If checkBoxFormatting.Checked = True Then

                        'background fill color
                        lblColor = displayRng.Cells(i, j).Interior.Color
                        rgbColor = System.Drawing.Color.FromArgb(CInt(lblColor Mod 256), CInt((lblColor \ 256) Mod 256), CInt((lblColor \ 65536) Mod 256))
                        label.BackColor = rgbColor

                        'font color
                        lblColor = displayRng.Cells(i, j).Font.Color
                        rgbColor = System.Drawing.Color.FromArgb(CInt(lblColor Mod 256), CInt((lblColor \ 256) Mod 256), CInt((lblColor \ 65536) Mod 256))
                        label.ForeColor = rgbColor

                    End If
                    'label.BackColor = displayRng.Cells(i, j).interior.color
                    'label.ForeColor = displayRng.Cells(i, j).font.color

                    CP_Input_Range1.Controls.Add(label)
                Next
            Next

            CP_Input_Range1.AutoScroll = True

secondDisplay:

            If txtSourceRange2.Text = "" Or secondInputRng Is Nothing Then
                CP_Input_Range2.Controls.Clear()
                Exit Sub
            End If

            Dim displayRng2 As Excel.Range
            If secondInputRng.Rows.Count > 50 Then
                displayRng2 = secondInputRng.Rows("1:50")
            Else
                displayRng2 = secondInputRng
            End If


            Dim height2 As Double
            Dim width2 As Double

            If displayRng2.Rows.Count <= 4 Then
                height2 = CP_Input_Range2.Height / displayRng2.Rows.Count
            Else
                height2 = (119 / 4)
            End If

            If displayRng2.Columns.Count <= 3 Then
                width2 = CP_Input_Range2.Width / displayRng2.Columns.Count
            Else
                width2 = (260 / 3)
            End If

            For i = 1 To displayRng2.Rows.Count
                For j = 1 To displayRng2.Columns.Count
                    Dim label As New System.Windows.Forms.Label
                    label.Text = displayRng2.Cells(i, j).Value
                    label.Location = New System.Drawing.Point((j - 1) * width2, (i - 1) * height2)
                    label.Height = height2
                    label.Width = width2
                    label.BorderStyle = BorderStyle.FixedSingle
                    label.TextAlign = ContentAlignment.MiddleCenter

                    If checkBoxFormatting.Checked = True Then

                        'backgroud fill color
                        lblColor = displayRng2.Cells(i, j).Interior.Color
                        rgbColor = System.Drawing.Color.FromArgb(CInt(lblColor Mod 256), CInt((lblColor \ 256) Mod 256), CInt((lblColor \ 65536) Mod 256))
                        label.BackColor = rgbColor

                        'font color
                        lblColor = displayRng2.Cells(i, j).Font.Color
                        rgbColor = System.Drawing.Color.FromArgb(CInt(lblColor Mod 256), CInt((lblColor \ 256) Mod 256), CInt((lblColor \ 65536) Mod 256))
                        label.ForeColor = rgbColor

                    End If



                    CP_Input_Range2.Controls.Add(label)
                Next
            Next

            CP_Input_Range2.AutoScroll = True

            If txtSourceRange1.Text = "" Or firstInputRng Is Nothing Then
                Exit Sub
            End If

            If firstInputRng.Rows.Count > 50 Then
                displayRng = firstInputRng.Rows("1:50")
            Else
                displayRng = firstInputRng
            End If

            If displayRng.Rows.Count <> displayRng2.Rows.Count Or displayRng.Columns.Count <> displayRng2.Columns.Count Then
                Exit Sub
            End If


            If radBtnSameValues.Checked = True Then

                If checkBoxCase.Checked = True Then

                    '1st range >> 2nd range >> radBtnSameValues checked >> case sensitive checked >> both fill/font color selected
                    If checkBoxFillBack.Checked = True And checkBoxFillFont.Checked = True Then
                        For i = 1 To displayRng.Rows.Count
                            For j = 1 To displayRng.Columns.Count

                                If VarType(displayRng.Cells(i, j).value) = VarType(displayRng2.Cells(i, j).value) Then
                                    If displayRng.Cells(i, j).value = displayRng2.Cells(i, j).value Then

                                        Dim label As New System.Windows.Forms.Label
                                        label.Text = displayRng.Cells(i, j).Value
                                        label.Location = New System.Drawing.Point((j - 1) * width, (i - 1) * height)
                                        label.Height = height
                                        label.Width = width
                                        label.BorderStyle = BorderStyle.FixedSingle
                                        label.TextAlign = ContentAlignment.MiddleCenter
                                        label.BackColor = CBFillBackground.BackColor
                                        label.ForeColor = CbFillFont.BackColor

                                        CP_Output_Range.Controls.Add(label)
                                    Else
                                        Dim label As New System.Windows.Forms.Label
                                        label.Text = displayRng.Cells(i, j).Value
                                        label.Location = New System.Drawing.Point((j - 1) * width, (i - 1) * height)
                                        label.Height = height
                                        label.Width = width
                                        label.BorderStyle = BorderStyle.FixedSingle
                                        label.TextAlign = ContentAlignment.MiddleCenter
                                        'label.BackColor = Color.Transparent
                                        'label.ForeColor = Nothing
                                        If checkBoxFormatting.Checked = True Then

                                            'background fill color
                                            lblColor = displayRng.Cells(i, j).Interior.Color
                                            rgbColor = System.Drawing.Color.FromArgb(CInt(lblColor Mod 256), CInt((lblColor \ 256) Mod 256), CInt((lblColor \ 65536) Mod 256))
                                            label.BackColor = rgbColor

                                            'font color
                                            lblColor = displayRng.Cells(i, j).Font.Color
                                            rgbColor = System.Drawing.Color.FromArgb(CInt(lblColor Mod 256), CInt((lblColor \ 256) Mod 256), CInt((lblColor \ 65536) Mod 256))
                                            label.ForeColor = rgbColor

                                        Else
                                            label.BackColor = Color.Transparent
                                            label.ForeColor = Nothing

                                        End If

                                        CP_Output_Range.Controls.Add(label)

                                    End If

                                Else
                                    Dim label As New System.Windows.Forms.Label
                                    label.Text = displayRng.Cells(i, j).Value
                                    label.Location = New System.Drawing.Point((j - 1) * width, (i - 1) * height)
                                    label.Height = height
                                    label.Width = width
                                    label.BorderStyle = BorderStyle.FixedSingle
                                    label.TextAlign = ContentAlignment.MiddleCenter
                                    'label.BackColor = Color.Transparent
                                    'label.ForeColor = Nothing


                                    If checkBoxFormatting.Checked = True Then

                                        'background fill color
                                        lblColor = displayRng.Cells(i, j).Interior.Color
                                        rgbColor = System.Drawing.Color.FromArgb(CInt(lblColor Mod 256), CInt((lblColor \ 256) Mod 256), CInt((lblColor \ 65536) Mod 256))
                                        label.BackColor = rgbColor

                                        'font color
                                        lblColor = displayRng.Cells(i, j).Font.Color
                                        rgbColor = System.Drawing.Color.FromArgb(CInt(lblColor Mod 256), CInt((lblColor \ 256) Mod 256), CInt((lblColor \ 65536) Mod 256))
                                        label.ForeColor = rgbColor

                                    Else
                                        label.BackColor = Color.Transparent
                                        label.ForeColor = Nothing

                                    End If

                                    CP_Output_Range.Controls.Add(label)
                                End If
                            Next
                        Next

                        '1st range >> 2nd range >> radBtnSameValues checked >> case sensitive checked >> only fill color is selected
                    ElseIf checkBoxFillBack.Checked = True And checkBoxFillFont.Checked = False Then
                        For i = 1 To displayRng.Rows.Count
                            For j = 1 To displayRng.Columns.Count

                                If VarType(displayRng.Cells(i, j).value) = VarType(displayRng2.Cells(i, j).value) Then
                                    If displayRng.Cells(i, j).value = displayRng2.Cells(i, j).value Then

                                        Dim label As New System.Windows.Forms.Label
                                        label.Text = displayRng.Cells(i, j).Value
                                        label.Location = New System.Drawing.Point((j - 1) * width, (i - 1) * height)
                                        label.Height = height
                                        label.Width = width
                                        label.BorderStyle = BorderStyle.FixedSingle
                                        label.TextAlign = ContentAlignment.MiddleCenter
                                        label.BackColor = CBFillBackground.BackColor
                                        label.ForeColor = Nothing

                                        CP_Output_Range.Controls.Add(label)
                                    Else
                                        Dim label As New System.Windows.Forms.Label
                                        label.Text = displayRng.Cells(i, j).Value
                                        label.Location = New System.Drawing.Point((j - 1) * width, (i - 1) * height)
                                        label.Height = height
                                        label.Width = width
                                        label.BorderStyle = BorderStyle.FixedSingle
                                        label.TextAlign = ContentAlignment.MiddleCenter
                                        'label.BackColor = Color.Transparent
                                        'label.ForeColor = Nothing

                                        If checkBoxFormatting.Checked = True Then

                                            'background fill color
                                            lblColor = displayRng.Cells(i, j).Interior.Color
                                            rgbColor = System.Drawing.Color.FromArgb(CInt(lblColor Mod 256), CInt((lblColor \ 256) Mod 256), CInt((lblColor \ 65536) Mod 256))
                                            label.BackColor = rgbColor

                                            'font color
                                            lblColor = displayRng.Cells(i, j).Font.Color
                                            rgbColor = System.Drawing.Color.FromArgb(CInt(lblColor Mod 256), CInt((lblColor \ 256) Mod 256), CInt((lblColor \ 65536) Mod 256))
                                            label.ForeColor = rgbColor

                                        Else
                                            label.BackColor = Color.Transparent
                                            label.ForeColor = Nothing

                                        End If

                                        CP_Output_Range.Controls.Add(label)

                                    End If

                                Else
                                    Dim label As New System.Windows.Forms.Label
                                    label.Text = displayRng.Cells(i, j).Value
                                    label.Location = New System.Drawing.Point((j - 1) * width, (i - 1) * height)
                                    label.Height = height
                                    label.Width = width
                                    label.BorderStyle = BorderStyle.FixedSingle
                                    label.TextAlign = ContentAlignment.MiddleCenter
                                    'label.BackColor = Color.Transparent
                                    'label.ForeColor = Nothing


                                    If checkBoxFormatting.Checked = True Then

                                        'background fill color
                                        lblColor = displayRng.Cells(i, j).Interior.Color
                                        rgbColor = System.Drawing.Color.FromArgb(CInt(lblColor Mod 256), CInt((lblColor \ 256) Mod 256), CInt((lblColor \ 65536) Mod 256))
                                        label.BackColor = rgbColor

                                        'font color
                                        lblColor = displayRng.Cells(i, j).Font.Color
                                        rgbColor = System.Drawing.Color.FromArgb(CInt(lblColor Mod 256), CInt((lblColor \ 256) Mod 256), CInt((lblColor \ 65536) Mod 256))
                                        label.ForeColor = rgbColor

                                    Else
                                        label.BackColor = Color.Transparent
                                        label.ForeColor = Nothing

                                    End If

                                    CP_Output_Range.Controls.Add(label)

                                End If
                            Next
                        Next

                        '1st range >> 2nd range >> radBtnSameValues checked >> case sensitive checked >> only font color is selected
                    ElseIf checkBoxFillBack.Checked = False And checkBoxFillFont.Checked = True Then
                        For i = 1 To displayRng.Rows.Count
                            For j = 1 To displayRng.Columns.Count

                                If VarType(displayRng.Cells(i, j).value) = VarType(displayRng2.Cells(i, j).value) Then
                                    If displayRng.Cells(i, j).value = displayRng2.Cells(i, j).value Then

                                        Dim label As New System.Windows.Forms.Label
                                        label.Text = displayRng.Cells(i, j).Value
                                        label.Location = New System.Drawing.Point((j - 1) * width, (i - 1) * height)
                                        label.Height = height
                                        label.Width = width
                                        label.BorderStyle = BorderStyle.FixedSingle
                                        label.TextAlign = ContentAlignment.MiddleCenter
                                        label.BackColor = Color.Transparent
                                        label.ForeColor = CbFillFont.BackColor

                                        CP_Output_Range.Controls.Add(label)
                                    Else
                                        Dim label As New System.Windows.Forms.Label
                                        label.Text = displayRng.Cells(i, j).Value
                                        label.Location = New System.Drawing.Point((j - 1) * width, (i - 1) * height)
                                        label.Height = height
                                        label.Width = width
                                        label.BorderStyle = BorderStyle.FixedSingle
                                        label.TextAlign = ContentAlignment.MiddleCenter
                                        'label.BackColor = Color.Transparent
                                        'label.ForeColor = Nothing

                                        If checkBoxFormatting.Checked = True Then

                                            'background fill color
                                            lblColor = displayRng.Cells(i, j).Interior.Color
                                            rgbColor = System.Drawing.Color.FromArgb(CInt(lblColor Mod 256), CInt((lblColor \ 256) Mod 256), CInt((lblColor \ 65536) Mod 256))
                                            label.BackColor = rgbColor

                                            'font color
                                            lblColor = displayRng.Cells(i, j).Font.Color
                                            rgbColor = System.Drawing.Color.FromArgb(CInt(lblColor Mod 256), CInt((lblColor \ 256) Mod 256), CInt((lblColor \ 65536) Mod 256))
                                            label.ForeColor = rgbColor

                                        Else
                                            label.BackColor = Color.Transparent
                                            label.ForeColor = Nothing

                                        End If

                                        CP_Output_Range.Controls.Add(label)

                                    End If

                                Else
                                    Dim label As New System.Windows.Forms.Label
                                    label.Text = displayRng.Cells(i, j).Value
                                    label.Location = New System.Drawing.Point((j - 1) * width, (i - 1) * height)
                                    label.Height = height
                                    label.Width = width
                                    label.BorderStyle = BorderStyle.FixedSingle
                                    label.TextAlign = ContentAlignment.MiddleCenter
                                    'label.BackColor = Color.Transparent
                                    'label.ForeColor = Nothing

                                    If checkBoxFormatting.Checked = True Then

                                        'background fill color
                                        lblColor = displayRng.Cells(i, j).Interior.Color
                                        rgbColor = System.Drawing.Color.FromArgb(CInt(lblColor Mod 256), CInt((lblColor \ 256) Mod 256), CInt((lblColor \ 65536) Mod 256))
                                        label.BackColor = rgbColor

                                        'font color
                                        lblColor = displayRng.Cells(i, j).Font.Color
                                        rgbColor = System.Drawing.Color.FromArgb(CInt(lblColor Mod 256), CInt((lblColor \ 256) Mod 256), CInt((lblColor \ 65536) Mod 256))
                                        label.ForeColor = rgbColor

                                    Else
                                        label.BackColor = Color.Transparent
                                        label.ForeColor = Nothing

                                    End If

                                    CP_Output_Range.Controls.Add(label)

                                End If
                            Next
                        Next

                    Else
                        '1st range >> 2nd range >> radBtnSameValues checked >> case sensitive checked >> fill/font color not selected
                        For i = 1 To displayRng.Rows.Count
                            For j = 1 To displayRng.Columns.Count
                                Dim label As New System.Windows.Forms.Label
                                label.Text = displayRng.Cells(i, j).Value
                                label.Location = New System.Drawing.Point((j - 1) * width, (i - 1) * height)
                                label.Height = height
                                label.Width = width
                                label.BorderStyle = BorderStyle.FixedSingle
                                label.TextAlign = ContentAlignment.MiddleCenter
                                'label.BackColor = Color.Transparent
                                'label.ForeColor = Nothing

                                If checkBoxFormatting.Checked = True Then

                                    'background fill color
                                    lblColor = displayRng.Cells(i, j).Interior.Color
                                    rgbColor = System.Drawing.Color.FromArgb(CInt(lblColor Mod 256), CInt((lblColor \ 256) Mod 256), CInt((lblColor \ 65536) Mod 256))
                                    label.BackColor = rgbColor

                                    'font color
                                    lblColor = displayRng.Cells(i, j).Font.Color
                                    rgbColor = System.Drawing.Color.FromArgb(CInt(lblColor Mod 256), CInt((lblColor \ 256) Mod 256), CInt((lblColor \ 65536) Mod 256))
                                    label.ForeColor = rgbColor

                                Else
                                    label.BackColor = Color.Transparent
                                    label.ForeColor = Nothing

                                End If

                                CP_Output_Range.Controls.Add(label)

                            Next
                        Next

                    End If
                Else
                    '1st range >> 2nd range >> radBtnSameValues checked >> case sensitive unchecked >> fill/font color both are selected
                    If checkBoxFillBack.Checked = True And checkBoxFillFont.Checked = True Then
                        For i = 1 To displayRng.Rows.Count
                            For j = 1 To displayRng.Columns.Count
                                rng1CellValue = displayRng.Cells(i, j).value
                                rng2CellValue = displayRng2.Cells(i, j).value

                                If rng1CellValue Is Nothing Or rng2CellValue Is Nothing Then
                                    Exit Sub
                                End If


                                If VarType(displayRng.Cells(i, j).value) = VarType(displayRng2.Cells(i, j).value) Then
                                    If rng1CellValue.ToUpper = rng2CellValue.ToUpper Then

                                        Dim label As New System.Windows.Forms.Label
                                        label.Text = displayRng.Cells(i, j).Value
                                        label.Location = New System.Drawing.Point((j - 1) * width, (i - 1) * height)
                                        label.Height = height
                                        label.Width = width
                                        label.BorderStyle = BorderStyle.FixedSingle
                                        label.TextAlign = ContentAlignment.MiddleCenter
                                        label.BackColor = CBFillBackground.BackColor
                                        label.ForeColor = CbFillFont.BackColor

                                        CP_Output_Range.Controls.Add(label)
                                    Else
                                        Dim label As New System.Windows.Forms.Label
                                        label.Text = displayRng.Cells(i, j).Value
                                        label.Location = New System.Drawing.Point((j - 1) * width, (i - 1) * height)
                                        label.Height = height
                                        label.Width = width
                                        label.BorderStyle = BorderStyle.FixedSingle
                                        label.TextAlign = ContentAlignment.MiddleCenter
                                        'label.BackColor = Color.Transparent
                                        'label.ForeColor = Nothing

                                        If checkBoxFormatting.Checked = True Then

                                            'background fill color
                                            lblColor = displayRng.Cells(i, j).Interior.Color
                                            rgbColor = System.Drawing.Color.FromArgb(CInt(lblColor Mod 256), CInt((lblColor \ 256) Mod 256), CInt((lblColor \ 65536) Mod 256))
                                            label.BackColor = rgbColor

                                            'font color
                                            lblColor = displayRng.Cells(i, j).Font.Color
                                            rgbColor = System.Drawing.Color.FromArgb(CInt(lblColor Mod 256), CInt((lblColor \ 256) Mod 256), CInt((lblColor \ 65536) Mod 256))
                                            label.ForeColor = rgbColor

                                        Else
                                            label.BackColor = Color.Transparent
                                            label.ForeColor = Nothing

                                        End If

                                        CP_Output_Range.Controls.Add(label)

                                    End If

                                Else
                                    Dim label As New System.Windows.Forms.Label
                                    label.Text = displayRng.Cells(i, j).Value
                                    label.Location = New System.Drawing.Point((j - 1) * width, (i - 1) * height)
                                    label.Height = height
                                    label.Width = width
                                    label.BorderStyle = BorderStyle.FixedSingle
                                    label.TextAlign = ContentAlignment.MiddleCenter
                                    'label.BackColor = Color.Transparent
                                    'label.ForeColor = Nothing

                                    If checkBoxFormatting.Checked = True Then

                                        'background fill color
                                        lblColor = displayRng.Cells(i, j).Interior.Color
                                        rgbColor = System.Drawing.Color.FromArgb(CInt(lblColor Mod 256), CInt((lblColor \ 256) Mod 256), CInt((lblColor \ 65536) Mod 256))
                                        label.BackColor = rgbColor

                                        'font color
                                        lblColor = displayRng.Cells(i, j).Font.Color
                                        rgbColor = System.Drawing.Color.FromArgb(CInt(lblColor Mod 256), CInt((lblColor \ 256) Mod 256), CInt((lblColor \ 65536) Mod 256))
                                        label.ForeColor = rgbColor

                                    Else
                                        label.BackColor = Color.Transparent
                                        label.ForeColor = Nothing

                                    End If

                                    CP_Output_Range.Controls.Add(label)

                                End If
                            Next
                        Next

                        '1st range >> 2nd range >> radBtnSameValues checked >> case sensitive unchecked >> only fill color is selected
                    ElseIf checkBoxFillBack.Checked = True And checkBoxFillFont.Checked = False Then
                        For i = 1 To displayRng.Rows.Count
                            For j = 1 To displayRng.Columns.Count
                                rng1CellValue = displayRng.Cells(i, j).value
                                rng2CellValue = displayRng2.Cells(i, j).value

                                If rng1CellValue Is Nothing Or rng2CellValue Is Nothing Then
                                    Exit Sub
                                End If


                                If VarType(displayRng.Cells(i, j).value) = VarType(displayRng2.Cells(i, j).value) Then
                                    If rng1CellValue.ToUpper = rng2CellValue.ToUpper Then

                                        Dim label As New System.Windows.Forms.Label
                                        label.Text = displayRng.Cells(i, j).Value
                                        label.Location = New System.Drawing.Point((j - 1) * width, (i - 1) * height)
                                        label.Height = height
                                        label.Width = width
                                        label.BorderStyle = BorderStyle.FixedSingle
                                        label.TextAlign = ContentAlignment.MiddleCenter
                                        label.BackColor = CBFillBackground.BackColor
                                        label.ForeColor = Nothing

                                        CP_Output_Range.Controls.Add(label)
                                    Else
                                        Dim label As New System.Windows.Forms.Label
                                        label.Text = displayRng.Cells(i, j).Value
                                        label.Location = New System.Drawing.Point((j - 1) * width, (i - 1) * height)
                                        label.Height = height
                                        label.Width = width
                                        label.BorderStyle = BorderStyle.FixedSingle
                                        label.TextAlign = ContentAlignment.MiddleCenter
                                        'label.BackColor = Color.Transparent
                                        'label.ForeColor = Nothing

                                        If checkBoxFormatting.Checked = True Then

                                            'background fill color
                                            lblColor = displayRng.Cells(i, j).Interior.Color
                                            rgbColor = System.Drawing.Color.FromArgb(CInt(lblColor Mod 256), CInt((lblColor \ 256) Mod 256), CInt((lblColor \ 65536) Mod 256))
                                            label.BackColor = rgbColor

                                            'font color
                                            lblColor = displayRng.Cells(i, j).Font.Color
                                            rgbColor = System.Drawing.Color.FromArgb(CInt(lblColor Mod 256), CInt((lblColor \ 256) Mod 256), CInt((lblColor \ 65536) Mod 256))
                                            label.ForeColor = rgbColor

                                        Else
                                            label.BackColor = Color.Transparent
                                            label.ForeColor = Nothing

                                        End If

                                        CP_Output_Range.Controls.Add(label)

                                    End If

                                Else
                                    Dim label As New System.Windows.Forms.Label
                                    label.Text = displayRng.Cells(i, j).Value
                                    label.Location = New System.Drawing.Point((j - 1) * width, (i - 1) * height)
                                    label.Height = height
                                    label.Width = width
                                    label.BorderStyle = BorderStyle.FixedSingle
                                    label.TextAlign = ContentAlignment.MiddleCenter
                                    'label.BackColor = Color.Transparent
                                    'label.ForeColor = Nothing

                                    If checkBoxFormatting.Checked = True Then

                                        'background fill color
                                        lblColor = displayRng.Cells(i, j).Interior.Color
                                        rgbColor = System.Drawing.Color.FromArgb(CInt(lblColor Mod 256), CInt((lblColor \ 256) Mod 256), CInt((lblColor \ 65536) Mod 256))
                                        label.BackColor = rgbColor

                                        'font color
                                        lblColor = displayRng.Cells(i, j).Font.Color
                                        rgbColor = System.Drawing.Color.FromArgb(CInt(lblColor Mod 256), CInt((lblColor \ 256) Mod 256), CInt((lblColor \ 65536) Mod 256))
                                        label.ForeColor = rgbColor

                                    Else
                                        label.BackColor = Color.Transparent
                                        label.ForeColor = Nothing

                                    End If


                                    CP_Output_Range.Controls.Add(label)

                                End If
                            Next
                        Next

                        '1st range >> 2nd range >> radBtnSameValues checked >> case sensitive unchecked >> only font color is selected
                    ElseIf checkBoxFillBack.Checked = False And checkBoxFillFont.Checked = True Then
                        For i = 1 To displayRng.Rows.Count
                            For j = 1 To displayRng.Columns.Count
                                rng1CellValue = displayRng.Cells(i, j).value
                                rng2CellValue = displayRng2.Cells(i, j).value

                                If rng1CellValue Is Nothing Or rng2CellValue Is Nothing Then
                                    Exit Sub
                                End If



                                If VarType(displayRng.Cells(i, j).value) = VarType(displayRng2.Cells(i, j).value) Then
                                    If rng1CellValue.ToUpper = rng2CellValue.ToUpper Then

                                        Dim label As New System.Windows.Forms.Label
                                        label.Text = displayRng.Cells(i, j).Value
                                        label.Location = New System.Drawing.Point((j - 1) * width, (i - 1) * height)
                                        label.Height = height
                                        label.Width = width
                                        label.BorderStyle = BorderStyle.FixedSingle
                                        label.TextAlign = ContentAlignment.MiddleCenter
                                        label.BackColor = Color.Transparent
                                        label.ForeColor = CbFillFont.BackColor

                                        CP_Output_Range.Controls.Add(label)
                                    Else
                                        Dim label As New System.Windows.Forms.Label
                                        label.Text = displayRng.Cells(i, j).Value
                                        label.Location = New System.Drawing.Point((j - 1) * width, (i - 1) * height)
                                        label.Height = height
                                        label.Width = width
                                        label.BorderStyle = BorderStyle.FixedSingle
                                        label.TextAlign = ContentAlignment.MiddleCenter
                                        'label.BackColor = Color.Transparent
                                        'label.ForeColor = Nothing

                                        If checkBoxFormatting.Checked = True Then

                                            'background fill color
                                            lblColor = displayRng.Cells(i, j).Interior.Color
                                            rgbColor = System.Drawing.Color.FromArgb(CInt(lblColor Mod 256), CInt((lblColor \ 256) Mod 256), CInt((lblColor \ 65536) Mod 256))
                                            label.BackColor = rgbColor

                                            'font color
                                            lblColor = displayRng.Cells(i, j).Font.Color
                                            rgbColor = System.Drawing.Color.FromArgb(CInt(lblColor Mod 256), CInt((lblColor \ 256) Mod 256), CInt((lblColor \ 65536) Mod 256))
                                            label.ForeColor = rgbColor

                                        Else
                                            label.BackColor = Color.Transparent
                                            label.ForeColor = Nothing

                                        End If

                                        CP_Output_Range.Controls.Add(label)

                                    End If

                                Else
                                    Dim label As New System.Windows.Forms.Label
                                    label.Text = displayRng.Cells(i, j).Value
                                    label.Location = New System.Drawing.Point((j - 1) * width, (i - 1) * height)
                                    label.Height = height
                                    label.Width = width
                                    label.BorderStyle = BorderStyle.FixedSingle
                                    label.TextAlign = ContentAlignment.MiddleCenter
                                    'label.BackColor = Color.Transparent
                                    'label.ForeColor = Nothing

                                    If checkBoxFormatting.Checked = True Then

                                        'background fill color
                                        lblColor = displayRng.Cells(i, j).Interior.Color
                                        rgbColor = System.Drawing.Color.FromArgb(CInt(lblColor Mod 256), CInt((lblColor \ 256) Mod 256), CInt((lblColor \ 65536) Mod 256))
                                        label.BackColor = rgbColor

                                        'font color
                                        lblColor = displayRng.Cells(i, j).Font.Color
                                        rgbColor = System.Drawing.Color.FromArgb(CInt(lblColor Mod 256), CInt((lblColor \ 256) Mod 256), CInt((lblColor \ 65536) Mod 256))
                                        label.ForeColor = rgbColor

                                    Else
                                        label.BackColor = Color.Transparent
                                        label.ForeColor = Nothing

                                    End If

                                    CP_Output_Range.Controls.Add(label)

                                End If
                            Next
                        Next

                        '1st range >> 2nd range >> radBtnSameValues checked >> case sensitive unchecked >> fill/font color not selected
                    Else
                        For i = 1 To displayRng.Rows.Count
                            For j = 1 To displayRng.Columns.Count
                                rng1CellValue = displayRng.Cells(i, j).value
                                rng2CellValue = displayRng2.Cells(i, j).value

                                If rng1CellValue Is Nothing Or rng2CellValue Is Nothing Then
                                    Exit Sub
                                End If


                                Dim label As New System.Windows.Forms.Label
                                label.Text = displayRng.Cells(i, j).Value
                                label.Location = New System.Drawing.Point((j - 1) * width, (i - 1) * height)
                                label.Height = height
                                label.Width = width
                                label.BorderStyle = BorderStyle.FixedSingle
                                label.TextAlign = ContentAlignment.MiddleCenter
                                'label.BackColor = Color.Transparent
                                'label.ForeColor = Nothing
                                If checkBoxFormatting.Checked = True Then

                                    'background fill color
                                    lblColor = displayRng.Cells(i, j).Interior.Color
                                    rgbColor = System.Drawing.Color.FromArgb(CInt(lblColor Mod 256), CInt((lblColor \ 256) Mod 256), CInt((lblColor \ 65536) Mod 256))
                                    label.BackColor = rgbColor

                                    'font color
                                    lblColor = displayRng.Cells(i, j).Font.Color
                                    rgbColor = System.Drawing.Color.FromArgb(CInt(lblColor Mod 256), CInt((lblColor \ 256) Mod 256), CInt((lblColor \ 65536) Mod 256))
                                    label.ForeColor = rgbColor

                                Else
                                    label.BackColor = Color.Transparent
                                    label.ForeColor = Nothing

                                End If

                                CP_Output_Range.Controls.Add(label)
                            Next
                        Next

                    End If

                End If

            ElseIf radBtnDifferentValues.Checked = True Then

                If checkBoxCase.Checked = True Then

                    '1st range >> 2nd range >> radBtnDifferentValues checked >> case sensitive checked >> fill/font color both are selected
                    If checkBoxFillBack.Checked = True And checkBoxFillFont.Checked = True Then
                        For i = 1 To displayRng.Rows.Count
                            For j = 1 To displayRng.Columns.Count

                                If VarType(displayRng.Cells(i, j).value) = VarType(displayRng2.Cells(i, j).value) Then
                                    If displayRng.Cells(i, j).value <> displayRng2.Cells(i, j).value Then

                                        Dim label As New System.Windows.Forms.Label
                                        label.Text = displayRng.Cells(i, j).Value
                                        label.Location = New System.Drawing.Point((j - 1) * width, (i - 1) * height)
                                        label.Height = height
                                        label.Width = width
                                        label.BorderStyle = BorderStyle.FixedSingle
                                        label.TextAlign = ContentAlignment.MiddleCenter
                                        label.BackColor = CBFillBackground.BackColor
                                        label.ForeColor = CbFillFont.BackColor

                                        CP_Output_Range.Controls.Add(label)
                                    Else
                                        Dim label As New System.Windows.Forms.Label
                                        label.Text = displayRng.Cells(i, j).Value
                                        label.Location = New System.Drawing.Point((j - 1) * width, (i - 1) * height)
                                        label.Height = height
                                        label.Width = width
                                        label.BorderStyle = BorderStyle.FixedSingle
                                        label.TextAlign = ContentAlignment.MiddleCenter
                                        'label.BackColor = Color.Transparent
                                        'label.ForeColor = Nothing
                                        If checkBoxFormatting.Checked = True Then

                                            'background fill color
                                            lblColor = displayRng.Cells(i, j).Interior.Color
                                            rgbColor = System.Drawing.Color.FromArgb(CInt(lblColor Mod 256), CInt((lblColor \ 256) Mod 256), CInt((lblColor \ 65536) Mod 256))
                                            label.BackColor = rgbColor

                                            'font color
                                            lblColor = displayRng.Cells(i, j).Font.Color
                                            rgbColor = System.Drawing.Color.FromArgb(CInt(lblColor Mod 256), CInt((lblColor \ 256) Mod 256), CInt((lblColor \ 65536) Mod 256))
                                            label.ForeColor = rgbColor

                                        Else
                                            label.BackColor = Color.Transparent
                                            label.ForeColor = Nothing

                                        End If

                                        CP_Output_Range.Controls.Add(label)

                                    End If

                                Else
                                    Dim label As New System.Windows.Forms.Label
                                    label.Text = displayRng.Cells(i, j).Value
                                    label.Location = New System.Drawing.Point((j - 1) * width, (i - 1) * height)
                                    label.Height = height
                                    label.Width = width
                                    label.BorderStyle = BorderStyle.FixedSingle
                                    label.TextAlign = ContentAlignment.MiddleCenter
                                    label.BackColor = CBFillBackground.BackColor
                                    label.ForeColor = CbFillFont.BackColor



                                    CP_Output_Range.Controls.Add(label)

                                End If
                            Next
                        Next

                        '1st range >> 2nd range >> radBtnDifferentValues checked >> case sensitive checked >> only fill color is selected
                    ElseIf checkBoxFillBack.Checked = True And checkBoxFillFont.Checked = False Then
                        For i = 1 To displayRng.Rows.Count
                            For j = 1 To displayRng.Columns.Count

                                If VarType(displayRng.Cells(i, j).value) = VarType(displayRng2.Cells(i, j).value) Then
                                    If displayRng.Cells(i, j).value <> displayRng2.Cells(i, j).value Then

                                        Dim label As New System.Windows.Forms.Label
                                        label.Text = displayRng.Cells(i, j).Value
                                        label.Location = New System.Drawing.Point((j - 1) * width, (i - 1) * height)
                                        label.Height = height
                                        label.Width = width
                                        label.BorderStyle = BorderStyle.FixedSingle
                                        label.TextAlign = ContentAlignment.MiddleCenter
                                        label.BackColor = CBFillBackground.BackColor
                                        label.ForeColor = Nothing

                                        CP_Output_Range.Controls.Add(label)
                                    Else
                                        Dim label As New System.Windows.Forms.Label
                                        label.Text = displayRng.Cells(i, j).Value
                                        label.Location = New System.Drawing.Point((j - 1) * width, (i - 1) * height)
                                        label.Height = height
                                        label.Width = width
                                        label.BorderStyle = BorderStyle.FixedSingle
                                        label.TextAlign = ContentAlignment.MiddleCenter
                                        'label.BackColor = Color.Transparent
                                        'label.ForeColor = Nothing

                                        If checkBoxFormatting.Checked = True Then

                                            'background fill color
                                            lblColor = displayRng.Cells(i, j).Interior.Color
                                            rgbColor = System.Drawing.Color.FromArgb(CInt(lblColor Mod 256), CInt((lblColor \ 256) Mod 256), CInt((lblColor \ 65536) Mod 256))
                                            label.BackColor = rgbColor

                                            'font color
                                            lblColor = displayRng.Cells(i, j).Font.Color
                                            rgbColor = System.Drawing.Color.FromArgb(CInt(lblColor Mod 256), CInt((lblColor \ 256) Mod 256), CInt((lblColor \ 65536) Mod 256))
                                            label.ForeColor = rgbColor

                                        Else
                                            label.BackColor = Color.Transparent
                                            label.ForeColor = Nothing

                                        End If

                                        CP_Output_Range.Controls.Add(label)

                                    End If

                                Else
                                    Dim label As New System.Windows.Forms.Label
                                    label.Text = displayRng.Cells(i, j).Value
                                    label.Location = New System.Drawing.Point((j - 1) * width, (i - 1) * height)
                                    label.Height = height
                                    label.Width = width
                                    label.BorderStyle = BorderStyle.FixedSingle
                                    label.TextAlign = ContentAlignment.MiddleCenter
                                    label.BackColor = CBFillBackground.BackColor
                                    label.ForeColor = Nothing


                                    CP_Output_Range.Controls.Add(label)

                                End If
                            Next
                        Next

                        '1st range >> 2nd range >> radBtnDifferentValues checked >> case sensitive checked >> only font color is selected
                    ElseIf checkBoxFillBack.Checked = False And checkBoxFillFont.Checked = True Then
                        For i = 1 To displayRng.Rows.Count
                            For j = 1 To displayRng.Columns.Count

                                If VarType(displayRng.Cells(i, j).value) = VarType(displayRng2.Cells(i, j).value) Then
                                    If displayRng.Cells(i, j).value <> displayRng2.Cells(i, j).value Then

                                        Dim label As New System.Windows.Forms.Label
                                        label.Text = displayRng.Cells(i, j).Value
                                        label.Location = New System.Drawing.Point((j - 1) * width, (i - 1) * height)
                                        label.Height = height
                                        label.Width = width
                                        label.BorderStyle = BorderStyle.FixedSingle
                                        label.TextAlign = ContentAlignment.MiddleCenter
                                        label.BackColor = Color.Transparent
                                        label.ForeColor = CbFillFont.BackColor

                                        CP_Output_Range.Controls.Add(label)
                                    Else
                                        Dim label As New System.Windows.Forms.Label
                                        label.Text = displayRng.Cells(i, j).Value
                                        label.Location = New System.Drawing.Point((j - 1) * width, (i - 1) * height)
                                        label.Height = height
                                        label.Width = width
                                        label.BorderStyle = BorderStyle.FixedSingle
                                        label.TextAlign = ContentAlignment.MiddleCenter
                                        'label.BackColor = Color.Transparent
                                        'label.ForeColor = Nothing

                                        If checkBoxFormatting.Checked = True Then

                                            'background fill color
                                            lblColor = displayRng.Cells(i, j).Interior.Color
                                            rgbColor = System.Drawing.Color.FromArgb(CInt(lblColor Mod 256), CInt((lblColor \ 256) Mod 256), CInt((lblColor \ 65536) Mod 256))
                                            label.BackColor = rgbColor

                                            'font color
                                            lblColor = displayRng.Cells(i, j).Font.Color
                                            rgbColor = System.Drawing.Color.FromArgb(CInt(lblColor Mod 256), CInt((lblColor \ 256) Mod 256), CInt((lblColor \ 65536) Mod 256))
                                            label.ForeColor = rgbColor

                                        Else
                                            label.BackColor = Color.Transparent
                                            label.ForeColor = Nothing

                                        End If

                                        CP_Output_Range.Controls.Add(label)

                                    End If

                                Else
                                    Dim label As New System.Windows.Forms.Label
                                    label.Text = displayRng.Cells(i, j).Value
                                    label.Location = New System.Drawing.Point((j - 1) * width, (i - 1) * height)
                                    label.Height = height
                                    label.Width = width
                                    label.BorderStyle = BorderStyle.FixedSingle
                                    label.TextAlign = ContentAlignment.MiddleCenter
                                    label.BackColor = Color.Transparent
                                    label.ForeColor = CbFillFont.BackColor



                                    CP_Output_Range.Controls.Add(label)

                                End If
                            Next
                        Next

                        '1st range >> 2nd range >> radBtnDifferentValues checked >> case sensitive checked >> fill/font color not selected
                    Else
                        For i = 1 To displayRng.Rows.Count
                            For j = 1 To displayRng.Columns.Count

                                Dim label As New System.Windows.Forms.Label
                                label.Text = displayRng.Cells(i, j).Value
                                label.Location = New System.Drawing.Point((j - 1) * width, (i - 1) * height)
                                label.Height = height
                                label.Width = width
                                label.BorderStyle = BorderStyle.FixedSingle
                                label.TextAlign = ContentAlignment.MiddleCenter
                                'label.BackColor = Color.Transparent
                                'label.ForeColor = Nothing

                                If checkBoxFormatting.Checked = True Then

                                    'background fill color
                                    lblColor = displayRng.Cells(i, j).Interior.Color
                                    rgbColor = System.Drawing.Color.FromArgb(CInt(lblColor Mod 256), CInt((lblColor \ 256) Mod 256), CInt((lblColor \ 65536) Mod 256))
                                    label.BackColor = rgbColor

                                    'font color
                                    lblColor = displayRng.Cells(i, j).Font.Color
                                    rgbColor = System.Drawing.Color.FromArgb(CInt(lblColor Mod 256), CInt((lblColor \ 256) Mod 256), CInt((lblColor \ 65536) Mod 256))
                                    label.ForeColor = rgbColor

                                Else
                                    label.BackColor = Color.Transparent
                                    label.ForeColor = Nothing

                                End If

                                CP_Output_Range.Controls.Add(label)

                            Next
                        Next

                    End If




                Else

                    '1st range >> 2nd range >> radBtnDifferentValues checked >> case sensitive unchecked >> fill/font color both are selected
                    If checkBoxFillBack.Checked = True And checkBoxFillFont.Checked = True Then
                        For i = 1 To displayRng.Rows.Count
                            For j = 1 To displayRng.Columns.Count
                                rng1CellValue = displayRng.Cells(i, j).value
                                rng2CellValue = displayRng2.Cells(i, j).value
                                If rng1CellValue Is Nothing Or rng2CellValue Is Nothing Then
                                    Exit Sub
                                End If
                                If VarType(displayRng.Cells(i, j).value) = VarType(displayRng2.Cells(i, j).value) Then
                                    If rng1CellValue.ToUpper <> rng2CellValue.ToUpper Then

                                        Dim label As New System.Windows.Forms.Label
                                        label.Text = displayRng.Cells(i, j).Value
                                        label.Location = New System.Drawing.Point((j - 1) * width, (i - 1) * height)
                                        label.Height = height
                                        label.Width = width
                                        label.BorderStyle = BorderStyle.FixedSingle
                                        label.TextAlign = ContentAlignment.MiddleCenter
                                        label.BackColor = CBFillBackground.BackColor
                                        label.ForeColor = CbFillFont.BackColor

                                        CP_Output_Range.Controls.Add(label)
                                    Else
                                        Dim label As New System.Windows.Forms.Label
                                        label.Text = displayRng.Cells(i, j).Value
                                        label.Location = New System.Drawing.Point((j - 1) * width, (i - 1) * height)
                                        label.Height = height
                                        label.Width = width
                                        label.BorderStyle = BorderStyle.FixedSingle
                                        label.TextAlign = ContentAlignment.MiddleCenter
                                        'label.BackColor = Color.Transparent
                                        'label.ForeColor = Nothing

                                        If checkBoxFormatting.Checked = True Then

                                            'background fill color
                                            lblColor = displayRng.Cells(i, j).Interior.Color
                                            rgbColor = System.Drawing.Color.FromArgb(CInt(lblColor Mod 256), CInt((lblColor \ 256) Mod 256), CInt((lblColor \ 65536) Mod 256))
                                            label.BackColor = rgbColor

                                            'font color
                                            lblColor = displayRng.Cells(i, j).Font.Color
                                            rgbColor = System.Drawing.Color.FromArgb(CInt(lblColor Mod 256), CInt((lblColor \ 256) Mod 256), CInt((lblColor \ 65536) Mod 256))
                                            label.ForeColor = rgbColor

                                        Else
                                            label.BackColor = Color.Transparent
                                            label.ForeColor = Nothing

                                        End If

                                        CP_Output_Range.Controls.Add(label)

                                    End If

                                Else
                                    Dim label As New System.Windows.Forms.Label
                                    label.Text = displayRng.Cells(i, j).Value
                                    label.Location = New System.Drawing.Point((j - 1) * width, (i - 1) * height)
                                    label.Height = height
                                    label.Width = width
                                    label.BorderStyle = BorderStyle.FixedSingle
                                    label.TextAlign = ContentAlignment.MiddleCenter
                                    label.BackColor = CBFillBackground.BackColor
                                    label.ForeColor = CbFillFont.BackColor


                                    CP_Output_Range.Controls.Add(label)

                                End If
                            Next
                        Next

                        '1st range >> 2nd range >> radBtnDifferentValues checked >> case sensitive unchecked >> only fill color is selected
                    ElseIf checkBoxFillBack.Checked = True And checkBoxFillFont.Checked = False Then
                        For i = 1 To displayRng.Rows.Count
                            For j = 1 To displayRng.Columns.Count
                                rng1CellValue = displayRng.Cells(i, j).value
                                rng2CellValue = displayRng2.Cells(i, j).value
                                If rng1CellValue Is Nothing Or rng2CellValue Is Nothing Then
                                    Exit Sub
                                End If
                                If VarType(displayRng.Cells(i, j).value) = VarType(displayRng2.Cells(i, j).value) Then
                                    If rng1CellValue.ToUpper <> rng2CellValue.ToUpper Then

                                        Dim label As New System.Windows.Forms.Label
                                        label.Text = displayRng.Cells(i, j).Value
                                        label.Location = New System.Drawing.Point((j - 1) * width, (i - 1) * height)
                                        label.Height = height
                                        label.Width = width
                                        label.BorderStyle = BorderStyle.FixedSingle
                                        label.TextAlign = ContentAlignment.MiddleCenter
                                        label.BackColor = CBFillBackground.BackColor
                                        label.ForeColor = Nothing

                                        CP_Output_Range.Controls.Add(label)
                                    Else
                                        Dim label As New System.Windows.Forms.Label
                                        label.Text = displayRng.Cells(i, j).Value
                                        label.Location = New System.Drawing.Point((j - 1) * width, (i - 1) * height)
                                        label.Height = height
                                        label.Width = width
                                        label.BorderStyle = BorderStyle.FixedSingle
                                        label.TextAlign = ContentAlignment.MiddleCenter
                                        'label.BackColor = Color.Transparent
                                        'label.ForeColor = Nothing

                                        If checkBoxFormatting.Checked = True Then

                                            'background fill color
                                            lblColor = displayRng.Cells(i, j).Interior.Color
                                            rgbColor = System.Drawing.Color.FromArgb(CInt(lblColor Mod 256), CInt((lblColor \ 256) Mod 256), CInt((lblColor \ 65536) Mod 256))
                                            label.BackColor = rgbColor

                                            'font color
                                            lblColor = displayRng.Cells(i, j).Font.Color
                                            rgbColor = System.Drawing.Color.FromArgb(CInt(lblColor Mod 256), CInt((lblColor \ 256) Mod 256), CInt((lblColor \ 65536) Mod 256))
                                            label.ForeColor = rgbColor

                                        Else
                                            label.BackColor = Color.Transparent
                                            label.ForeColor = Nothing

                                        End If

                                        CP_Output_Range.Controls.Add(label)

                                    End If

                                Else
                                    Dim label As New System.Windows.Forms.Label
                                    label.Text = displayRng.Cells(i, j).Value
                                    label.Location = New System.Drawing.Point((j - 1) * width, (i - 1) * height)
                                    label.Height = height
                                    label.Width = width
                                    label.BorderStyle = BorderStyle.FixedSingle
                                    label.TextAlign = ContentAlignment.MiddleCenter
                                    label.BackColor = CBFillBackground.BackColor
                                    label.ForeColor = Nothing


                                    CP_Output_Range.Controls.Add(label)

                                End If
                            Next
                        Next

                        '1st range >> 2nd range >> radBtnDifferentValues checked >> case sensitive unchecked >> only font color is selected
                    ElseIf checkBoxFillBack.Checked = False And checkBoxFillFont.Checked = True Then
                        For i = 1 To displayRng.Rows.Count
                            For j = 1 To displayRng.Columns.Count
                                rng1CellValue = displayRng.Cells(i, j).value
                                rng2CellValue = displayRng2.Cells(i, j).value
                                If rng1CellValue Is Nothing Or rng2CellValue Is Nothing Then
                                    Exit Sub
                                End If
                                If VarType(displayRng.Cells(i, j).value) = VarType(displayRng2.Cells(i, j).value) Then
                                    If rng1CellValue.ToUpper <> rng2CellValue.ToUpper Then

                                        Dim label As New System.Windows.Forms.Label
                                        label.Text = displayRng.Cells(i, j).Value
                                        label.Location = New System.Drawing.Point((j - 1) * width, (i - 1) * height)
                                        label.Height = height
                                        label.Width = width
                                        label.BorderStyle = BorderStyle.FixedSingle
                                        label.TextAlign = ContentAlignment.MiddleCenter
                                        label.BackColor = Color.Transparent
                                        label.ForeColor = CbFillFont.BackColor

                                        CP_Output_Range.Controls.Add(label)
                                    Else
                                        Dim label As New System.Windows.Forms.Label
                                        label.Text = displayRng.Cells(i, j).Value
                                        label.Location = New System.Drawing.Point((j - 1) * width, (i - 1) * height)
                                        label.Height = height
                                        label.Width = width
                                        label.BorderStyle = BorderStyle.FixedSingle
                                        label.TextAlign = ContentAlignment.MiddleCenter
                                        'label.BackColor = Color.Transparent
                                        'label.ForeColor = Nothing

                                        If checkBoxFormatting.Checked = True Then

                                            'background fill color
                                            lblColor = displayRng.Cells(i, j).Interior.Color
                                            rgbColor = System.Drawing.Color.FromArgb(CInt(lblColor Mod 256), CInt((lblColor \ 256) Mod 256), CInt((lblColor \ 65536) Mod 256))
                                            label.BackColor = rgbColor

                                            'font color
                                            lblColor = displayRng.Cells(i, j).Font.Color
                                            rgbColor = System.Drawing.Color.FromArgb(CInt(lblColor Mod 256), CInt((lblColor \ 256) Mod 256), CInt((lblColor \ 65536) Mod 256))
                                            label.ForeColor = rgbColor

                                        Else
                                            label.BackColor = Color.Transparent
                                            label.ForeColor = Nothing

                                        End If

                                        CP_Output_Range.Controls.Add(label)

                                    End If

                                Else
                                    Dim label As New System.Windows.Forms.Label
                                    label.Text = displayRng.Cells(i, j).Value
                                    label.Location = New System.Drawing.Point((j - 1) * width, (i - 1) * height)
                                    label.Height = height
                                    label.Width = width
                                    label.BorderStyle = BorderStyle.FixedSingle
                                    label.TextAlign = ContentAlignment.MiddleCenter
                                    label.BackColor = Color.Transparent
                                    label.ForeColor = CbFillFont.BackColor


                                    CP_Output_Range.Controls.Add(label)

                                End If
                            Next
                        Next

                        '1st range >> 2nd range >> radBtnDifferentValues checked >> case sensitive unchecked >> fill/font color not selected
                    Else
                        For i = 1 To displayRng.Rows.Count
                            For j = 1 To displayRng.Columns.Count
                                rng1CellValue = displayRng.Cells(i, j).value
                                rng2CellValue = displayRng2.Cells(i, j).value
                                If rng1CellValue Is Nothing Or rng2CellValue Is Nothing Then
                                    Exit Sub
                                End If

                                Dim label As New System.Windows.Forms.Label
                                label.Text = displayRng.Cells(i, j).Value
                                label.Location = New System.Drawing.Point((j - 1) * width, (i - 1) * height)
                                label.Height = height
                                label.Width = width
                                label.BorderStyle = BorderStyle.FixedSingle
                                label.TextAlign = ContentAlignment.MiddleCenter


                                If checkBoxFormatting.Checked = True Then

                                    'background fill color
                                    lblColor = displayRng.Cells(i, j).Interior.Color
                                    rgbColor = System.Drawing.Color.FromArgb(CInt(lblColor Mod 256), CInt((lblColor \ 256) Mod 256), CInt((lblColor \ 65536) Mod 256))
                                    label.BackColor = rgbColor

                                    'font color
                                    lblColor = displayRng.Cells(i, j).Font.Color
                                    rgbColor = System.Drawing.Color.FromArgb(CInt(lblColor Mod 256), CInt((lblColor \ 256) Mod 256), CInt((lblColor \ 65536) Mod 256))
                                    label.ForeColor = rgbColor

                                Else
                                    label.BackColor = Color.Transparent
                                    label.ForeColor = Nothing

                                End If


                                CP_Output_Range.Controls.Add(label)


                            Next
                        Next


                    End If

                End If


            End If



            CP_Output_Range.AutoScroll = True


        Catch ex As Exception

        End Try

    End Sub

    Private Sub CBFillBackground_Click(sender As Object, e As EventArgs) Handles CBFillBackground.Click
        Call Display()
        If checkBoxFillBack.Checked = True Then

            colorPick = CD_Fill_Background.ShowDialog()

            If colorPick = DialogResult.OK Then
                CBFillBackground.BackColor = CD_Fill_Background.Color
                Call Display()

            End If
            CP_Input_Range1.Focus()
        Else
            CP_Input_Range1.Focus()


        End If
    End Sub


    Private Sub CbFillFont_Click(sender As Object, e As EventArgs) Handles CbFillFont.Click

        Call Display()
        If checkBoxFillFont.Checked = True Then

            colorPick = CD_Fill_Font.ShowDialog()


            If colorPick = DialogResult.OK Then
                CbFillFont.BackColor = CD_Fill_Font.Color
                Call Display()

            End If
            CP_Input_Range1.Focus()
        Else
            CP_Input_Range1.Focus()


        End If

    End Sub

    Private Sub CBFillBackground_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CBFillBackground.SelectedIndexChanged
        Call Display()

    End Sub

    Private Sub radBtnSameValues_CheckedChanged(sender As Object, e As EventArgs) Handles radBtnSameValues.CheckedChanged
        Call Display()
    End Sub

    Private Sub radBtnDifferentValues_CheckedChanged(sender As Object, e As EventArgs) Handles radBtnDifferentValues.CheckedChanged
        Call Display()
    End Sub

    Private Sub CustomPanel1_Paint(sender As Object, e As PaintEventArgs) Handles CustomPanel1.Paint

    End Sub

    Private Sub checkBoxCase_CheckedChanged(sender As Object, e As EventArgs) Handles checkBoxCase.CheckedChanged
        Call Display()
    End Sub

    Private Sub CBFillBackground_BackColorChanged(sender As Object, e As EventArgs) Handles CBFillBackground.BackColorChanged

        If CBFillBackground.BackColor.Name = "LightSteelBlue" And GB_Display_Result.BackColor <> CBFillBackground.BackColor Then

            Exit Sub

        End If


        Call Display()



    End Sub

    Private Sub CbFillFont_BackColorChanged(sender As Object, e As EventArgs) Handles CbFillFont.BackColorChanged

        If CbFillFont.BackColor.Name = "MidnightBlue" And GB_Display_Result.BackColor <> CBFillBackground.BackColor Then

            Exit Sub

        End If


        Call Display()



    End Sub

    Private Sub checkBoxFormatting_CheckedChanged(sender As Object, e As EventArgs) Handles checkBoxFormatting.CheckedChanged

        Call Display()


    End Sub

    Private Sub checkBoxFillBack_CheckedChanged(sender As Object, e As EventArgs) Handles checkBoxFillBack.CheckedChanged
        Call Display()

    End Sub

    Private Sub checkBoxFillFont_CheckedChanged(sender As Object, e As EventArgs) Handles checkBoxFillFont.CheckedChanged

        Call Display()

    End Sub

    Private Sub txtSourceRange1_Click(sender As Object, e As EventArgs) Handles txtSourceRange1.Click
        'txtSourceRange1.SelectionStart = txtSourceRange1.TextLength
        'txtSourceRange1.ScrollToCaret()
    End Sub

    Private Sub txtSourceRange2_Click(sender As Object, e As EventArgs) Handles txtSourceRange2.Click
        'txtSourceRange2.SelectionStart = txtSourceRange2.TextLength
        'txtSourceRange2.ScrollToCaret()
    End Sub


End Class