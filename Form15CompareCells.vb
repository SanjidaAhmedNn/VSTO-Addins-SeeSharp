Imports System.ComponentModel.Design
Imports System.Security.Cryptography
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Excel
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Drawing
Imports System.ComponentModel
Imports System.Linq.Expressions
Imports System.Threading


Public Class Form15CompareCells
    Dim WithEvents excelApp As Excel.Application
    Dim workbook As Excel.Workbook
    Dim worksheet As Excel.Worksheet
    Dim outWorksheet As Excel.Worksheet
    Dim firstInputRng As Excel.Range
    Dim secondInputRng As Excel.Range
    Dim FocusedTxtBox As Integer
    Dim selectedRange As Excel.Range
    Dim firstRngRows, firstRngCols As Integer
    Dim colorPick As DialogResult
    Dim count As Integer
    Dim rng1CellValue, rng2CellValue, WsName, coloredRng As String



    Private Sub txtSourceRange1_TextChanged(sender As Object, e As EventArgs) Handles txtSourceRange1.TextChanged

        Try
            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            worksheet = workbook.ActiveSheet


            txtSourceRange1.Focus()


            firstInputRng = worksheet.Range(txtSourceRange1.Text)

            lblSourceRng1.Text = "1st Source Range (" & firstInputRng.Rows.Count & " rows x " & firstInputRng.Columns.Count & " columns)"

            firstRngRows = worksheet.Range(txtSourceRange1.Text).Rows.Count
            firstRngCols = worksheet.Range(txtSourceRange1.Text).Columns.Count


            Call Display()


        Catch ex As Exception

        End Try




    End Sub


    Private Sub txtSourceRange2_TextChanged(sender As Object, e As EventArgs) Handles txtSourceRange2.TextChanged

        Try
            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            worksheet = workbook.ActiveSheet


            txtSourceRange2.Focus()


            secondInputRng = worksheet.Range(txtSourceRange2.Text)

            lblSourceRng2.Text = "2nd Source Range (" & secondInputRng.Rows.Count & " rows x " & secondInputRng.Columns.Count & " columns)"



            Call Display()

        Catch ex As Exception

        End Try



    End Sub



    Private Sub Form15CompareCells_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        excelApp = Globals.ThisAddIn.Application
        workbook = excelApp.ActiveWorkbook
        worksheet = workbook.ActiveSheet

        Dim selectedRng As Excel.Range = excelApp.Selection
        txtSourceRange1.Text = selectedRng.Address
        txtSourceRange1.Focus()









    End Sub





    Private Sub rngSelection1_Click(sender As Object, e As EventArgs) Handles rngSelection1.Click

        Try

            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            worksheet = workbook.ActiveSheet
            selectedRange = excelApp.Selection
            txtSourceRange1.Focus()

            firstInputRng = excelApp.InputBox("Please Select the First Range", "First Range Selection", selectedRange.Address, Type:=8)
            firstInputRng.Select()
            txtSourceRange1.Text = firstInputRng.Address
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

            secondInputRng = excelApp.InputBox("Please Select the Second Range", "Second Range Selection", selectedRange.Address, Type:=8)
            secondInputRng.Select()
            txtSourceRange2.Text = secondInputRng.Address
            txtSourceRange2.Focus()




        Catch ex As Exception

            txtSourceRange2.Focus()

        End Try
    End Sub





    Private Sub AutoSelection1_Click(sender As Object, e As EventArgs) Handles AutoSelection1.Click

        Try

            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            worksheet = workbook.ActiveSheet
            selectedRange = excelApp.Selection
            selectedRange.Select()

            Dim topLeft, bottomRight As String



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

            selectedRange.Select()

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
        firstCell = selectedRange

        selectedRange = worksheet.Range(firstCell.Offset(0, 0), firstCell.Offset(firstRngRows - 1, firstRngCols - 1))

        selectedRange.Select()







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
            selectedRange = excelApp.Selection
            selectedRange.Select()



            If FocusedTxtBox = 1 Then

                txtSourceRange1.Text = selectedRange.Address
                worksheet = workbook.ActiveSheet
                firstInputRng = selectedRange

                txtSourceRange1.Focus()

            ElseIf FocusedTxtBox = 2 Then
                txtSourceRange2.Text = selectedRange.Address
                worksheet = workbook.ActiveSheet
                secondInputRng = selectedRange


                txtSourceRange2.Focus()

            End If




        Catch ex As Exception


        End Try

    End Sub

    Private Sub btnCanecl_Click(sender As Object, e As EventArgs) Handles btnCanecl.Click
        Me.Dispose()
    End Sub

    Private Sub btnOK_Click(sender As Object, e As EventArgs) Handles btnOK.Click


        If Not firstInputRng.Rows.Count = secondInputRng.Rows.Count And firstInputRng.Columns.Count = secondInputRng.Columns.Count Then

            MsgBox("You must use same number of rows and columns in both ranges.",, "Warning!")

            Me.Dispose()
            Exit Sub


        End If

        excelApp = Globals.ThisAddIn.Application
        Dim i, j As Integer
        Dim rng1CellValue, rng2CellValue, WsName As String
        Dim coloredRng As String

        worksheet = workbook.ActiveSheet
        WsName = worksheet.Name

        count = 0
        coloredRng = ""



        If Not firstInputRng.Rows.Count = secondInputRng.Rows.Count And firstInputRng.Columns.Count = secondInputRng.Columns.Count Then

            MsgBox("You must use same number of rows and columns in both ranges.",, "Warning!")
            Me.Dispose()
            Exit Sub

        End If

        If checkBoxFormatting.Checked = False Then

            firstInputRng.ClearFormats()

        End If



        If radBtnSameValues.Checked = True Then
            If checkBoxCase.Checked = True Then
                If checkBoxFillBack.Checked = True Or checkBoxFillFont.Checked = True Then

                    For i = 1 To firstInputRng.Rows.Count
                        For j = 1 To firstInputRng.Columns.Count
                            If firstInputRng.Cells(i, j).value = secondInputRng.Cells(i, j).value Then

                                firstInputRng.Cells(i, j).Interior.Color = CBFillBackground.BackColor

                                firstInputRng.Cells(i, j).Font.Color = CbFillFont.BackColor
                                count = count + 1
                                coloredRng = coloredRng & "," & firstInputRng.Cells(i, j).address
                            End If
                        Next
                    Next

                Else
                    MsgBox("Please select a Color for 'Cell Background' or 'Font Color'.")
                    Exit Sub

                End If
            Else
                If checkBoxFillBack.Checked = True Or checkBoxFillFont.Checked = True Then
                    For i = 1 To firstInputRng.Rows.Count
                        For j = 1 To firstInputRng.Columns.Count
                            rng1CellValue = firstInputRng.Cells(i, j).value
                            rng2CellValue = secondInputRng.Cells(i, j).value
                            If rng1CellValue.ToUpper = rng2CellValue.ToUpper Then

                                firstInputRng.Cells(i, j).Interior.Color = CBFillBackground.BackColor

                                firstInputRng.Cells(i, j).Font.Color = CbFillFont.BackColor
                                count = count + 1
                                coloredRng = coloredRng & "," & firstInputRng.Cells(i, j).address
                            End If
                        Next
                    Next
                Else
                    MsgBox("Please select a Color for 'Cell Background' or 'Font Color'.")
                    Exit Sub
                End If
            End If

        ElseIf radBtnDifferentValues.Checked = True Then
            If checkBoxCase.Checked = True Then
                If checkBoxFillBack.Checked = True Or checkBoxFillFont.Checked = True Then
                    For i = 1 To firstInputRng.Rows.Count
                        For j = 1 To firstInputRng.Columns.Count
                            If firstInputRng.Cells(i, j).value <> secondInputRng.Cells(i, j).value Then

                                firstInputRng.Cells(i, j).Interior.Color = CBFillBackground.BackColor

                                firstInputRng.Cells(i, j).Font.Color = CbFillFont.BackColor
                                count = count + 1
                                coloredRng = coloredRng & "," & firstInputRng.Cells(i, j).address

                            End If
                        Next
                    Next
                Else
                    MsgBox("Please select a Color for 'Cell Background' or 'Font Color'.")
                    Exit Sub
                End If
            Else
                If checkBoxFillBack.Checked = True Or checkBoxFillFont.Checked = True Then
                    For i = 1 To firstInputRng.Rows.Count
                        For j = 1 To firstInputRng.Columns.Count
                            rng1CellValue = firstInputRng.Cells(i, j).value
                            rng2CellValue = secondInputRng.Cells(i, j).value
                            If rng1CellValue.ToUpper <> rng2CellValue.ToUpper Then

                                firstInputRng.Cells(i, j).Interior.Color = CBFillBackground.BackColor

                                firstInputRng.Cells(i, j).Font.Color = CbFillFont.BackColor
                                count = count + 1
                                coloredRng = coloredRng & "," & firstInputRng.Cells(i, j).address

                            End If
                        Next
                    Next
                Else
                    MsgBox("Please select a Color for 'Cell Background' or 'Font Color'.")
                    Exit Sub
                End If
            End If

        End If

        If checkBoxCopyWs.Checked = True Then

            workbook.ActiveSheet.Copy(After:=workbook.Sheets(workbook.Sheets.Count))
            outWorksheet = workbook.Sheets(workbook.Sheets.Count)
            outWorksheet.Range("A1").Select()

            For i = 1 To worksheet.Range(txtSourceRange1.Text).Rows.Count
                For j = 1 To worksheet.Range(txtSourceRange1.Text).Columns.Count


                    worksheet.Range(txtSourceRange1.Text).Cells(i, j).Interior.Colorindex = -4142

                    worksheet.Range(txtSourceRange1.Text).Cells(i, j).Font.Color = Nothing

                Next
            Next

            worksheet = workbook.Sheets(WsName)
            worksheet.Activate()

        End If

        Me.Dispose()





        MsgBox(count & " cell(s) found.", MsgBoxStyle.Information, "SOFTEKO")

        coloredRng = Microsoft.VisualBasic.Right(coloredRng, Len(coloredRng) - 1)
        worksheet.Range(coloredRng).Select()


    End Sub


    Sub Display()

        Try

            CP_Input_Range1.Controls.Clear()
            CP_Input_Range2.Controls.Clear()
            CP_Output_Range.Controls.Clear()


            Dim displayRng As Excel.Range
            Dim displayRng2 As Excel.Range




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

                    CP_Input_Range1.Controls.Add(label)
                Next
            Next

            CP_Input_Range1.AutoScroll = True


            If secondInputRng Is Nothing Then
                Exit Sub
            End If


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

            If displayRng.Columns.Count <= 3 Then
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

                    CP_Input_Range2.Controls.Add(label)
                Next
            Next

            CP_Input_Range2.AutoScroll = True




            If radBtnSameValues.Checked = True Then
                    If checkBoxCase.Checked = True Then
                        If checkBoxFillBack.Checked = True Or checkBoxFillFont.Checked = True Then
                            For i = 1 To displayRng.Rows.Count
                                For j = 1 To displayRng.Columns.Count


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
                                        label.BackColor = Color.Transparent
                                        label.ForeColor = Nothing

                                        CP_Output_Range.Controls.Add(label)

                                    End If
                                Next
                            Next
                        End If
                    Else
                        If checkBoxFillBack.Checked = True Or checkBoxFillFont.Checked = True Then
                            For i = 1 To displayRng.Rows.Count
                                For j = 1 To displayRng.Columns.Count
                                    rng1CellValue = displayRng.Cells(i, j).value
                                rng2CellValue = displayRng2.Cells(i, j).value

                                If rng1CellValue Is Nothing Or rng2CellValue Is Nothing Then
                                    Exit Sub
                                End If

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
                                        label.BackColor = Color.Transparent
                                        label.ForeColor = Nothing

                                        CP_Output_Range.Controls.Add(label)

                                    End If
                                Next
                            Next
                        End If
                    End If

                ElseIf radBtnDifferentValues.Checked = True Then
                    If checkBoxCase.Checked = True Then
                        If checkBoxFillBack.Checked = True Or checkBoxFillFont.Checked = True Then
                            For i = 1 To displayRng.Rows.Count
                                For j = 1 To displayRng.Columns.Count

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
                                        label.BackColor = Color.Transparent
                                        label.ForeColor = Nothing

                                        CP_Output_Range.Controls.Add(label)

                                    End If
                                Next
                            Next
                        End If
                    Else
                        If checkBoxFillBack.Checked = True Or checkBoxFillFont.Checked = True Then
                            For i = 1 To displayRng.Rows.Count
                                For j = 1 To displayRng.Columns.Count
                                    rng1CellValue = displayRng.Cells(i, j).value
                                rng2CellValue = displayRng2.Cells(i, j).value
                                If rng1CellValue Is Nothing Or rng2CellValue Is Nothing Then
                                    Exit Sub
                                End If

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
                                        label.BackColor = Color.Transparent
                                        label.ForeColor = Nothing

                                        CP_Output_Range.Controls.Add(label)

                                    End If
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

        If checkBoxFillBack.Checked = True Then

            colorPick = CD_Fill_Background.ShowDialog()

            If colorPick = DialogResult.OK Then
                CBFillBackground.BackColor = CD_Fill_Background.Color
            End If


        End If
    End Sub


    Private Sub CbFillFont_Click(sender As Object, e As EventArgs) Handles CbFillFont.Click

        If checkBoxFillFont.Checked = True Then

            colorPick = CD_Fill_Font.ShowDialog()


            If colorPick = DialogResult.OK Then
                CbFillFont.BackColor = CD_Fill_Font.Color
            End If


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
End Class