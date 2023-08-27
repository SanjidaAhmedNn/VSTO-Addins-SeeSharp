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
    Dim WsName1, WsName2, rng1_Address, rng2_Address As String


    Private Sub Form11SwapRanges_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        excelApp = Globals.ThisAddIn.Application
        workbook = excelApp.ActiveWorkbook
        worksheet = workbook.ActiveSheet

        Dim selectedRng As Excel.Range = excelApp.Selection
        txtSourceRange1.Text = selectedRng.Address
        txtSourceRange1.Focus()

        radBtnValues.Checked = True




    End Sub

    Private Sub txtSourceRange1_TextChanged(sender As Object, e As EventArgs) Handles txtSourceRange1.TextChanged

        Try

            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            worksheet = workbook.ActiveSheet

            'WsName1 = Microsoft.VisualBasic.Left(txtSourceRange1.Text, txtSourceRange2.Text.IndexOf("!"))
            WsName1 = worksheet.Name
            worksheet1 = workbook.Sheets(WsName1)



            txtSourceRange1.Text = WsName1 & "!" & excelApp.Selection.Address



            rng1_Address = Microsoft.VisualBasic.Right(txtSourceRange1.Text, Len(txtSourceRange1.Text) - txtSourceRange1.Text.IndexOf("!") - 1)
            firstInputRng = worksheet1.Range(rng1_Address)
            'firstInputRng = worksheet.Range(txtSourceRange1.Text)

            txtSourceRange1.Focus()
            txtSourceRange1.SelectionStart = txtSourceRange1.TextLength
            txtSourceRange1.ScrollToCaret()

            'txtSourceRange1.Text = WsName1 & "!" & firstInputRng.Address

            lblSourceRng1.Text = "1st Source Range (" & firstInputRng.Rows.Count & " rows x " & firstInputRng.Columns.Count & " columns)"

            firstRngRows = worksheet.Range(txtSourceRange1.Text).Rows.Count
            firstRngCols = worksheet.Range(txtSourceRange1.Text).Columns.Count



            'firstInputRng.Select()


        Catch ex As Exception

        End Try
        If txtSourceRange1.Text = "" Or firstInputRng Is Nothing Then
            Exit Sub
        End If
        firstInputRng.Select()
        txtSourceRange1.Focus()
    End Sub

    Private Sub txtSourceRange2_TextChanged(sender As Object, e As EventArgs) Handles txtSourceRange2.TextChanged

        Try
            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            worksheet = workbook.ActiveSheet

            WsName2 = worksheet.Name
            'WsName2 = Microsoft.VisualBasic.Left(txtSourceRange2.Text, txtSourceRange2.Text.IndexOf("!"))
            'MsgBox(WsName2)
            worksheet2 = workbook.Sheets(WsName2)

            txtSourceRange2.Text = WsName2 & "!" & secondInputRng.Address


            txtSourceRange2.Focus()
            txtSourceRange2.SelectionStart = txtSourceRange2.TextLength
            txtSourceRange2.ScrollToCaret()



            rng2_Address = Microsoft.VisualBasic.Right(txtSourceRange2.Text, Len(txtSourceRange2.Text) - txtSourceRange2.Text.IndexOf("!") - 1)
            secondInputRng = worksheet2.Range(rng2_Address)
            'MsgBox(address)
            'secondInputRng = worksheet.Range(txtSourceRange2.Text)

            'txtSourceRange2.Text = WsName2 & "!" & secondInputRng.Address
            lblSourceRng2.Text = "2nd Source Range (" & secondInputRng.Rows.Count & " rows x " & secondInputRng.Columns.Count & " columns)"




        Catch ex As Exception

        End Try

        If txtSourceRange2.Text = "" Or secondInputRng Is Nothing Then
            Exit Sub
        End If

        secondInputRng.Select()
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

            firstInputRng.Select()
            'WsName1 = firstInputRng.Worksheet.Name

            'txtSourceRange1.Text = WsName1 & "!" & firstInputRng.Address


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

            secondInputRng.Select()
            'WsName2 = secondInputRng.Worksheet.Name

            'txtSourceRange2.Text = WsName2 & "!" & secondInputRng.Address
            'txtSourceRange2.Focus()



        Catch ex As Exception

            txtSourceRange2.Focus()

        End Try


    End Sub



    Private Sub AutoSelection1_Click(sender As Object, e As EventArgs) Handles AutoSelection1.Click

        Try

            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            worksheet = workbook.ActiveSheet

            'WsName1 = worksheet.Name

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
        'worksheet1 = workbook.Sheets(WsName1)
        'worksheet2 = workbook.Sheets(WsName2)

        selectedRange = excelApp.Selection
        selectedRange.Select()

        Dim bottomRight As String
        firstCell = selectedRange.Cells(1, 1)


        If selectedRange.Cells(1, 1).Offset(1, 0).Value = Nothing Then

            For i = 0 To firstInputRng.Columns.Count - 1
                If selectedRange.Cells(1, 1).offset(0, i).value <> Nothing Then
                    selectedRange = worksheet.Range(selectedRange.Cells(1, 1), selectedRange.Cells(1, 1).Offset(0, i))
                End If
                selectedRange.Select()
            Next

        ElseIf selectedRange.Cells(1, 1).Offset(0, 1).Value = Nothing Then
            For i = 0 To firstInputRng.Rows.Count - 1
                If selectedRange.Cells(1, 1).offset(i, 0).value <> Nothing Then
                    selectedRange = worksheet.Range(selectedRange.Cells(1, 1), selectedRange.Cells(1, 1).Offset(i, 0))
                End If
                selectedRange.Select()
            Next

        Else

            bottomRight = firstCell.End(XlDirection.xlToRight).Address
            bottomRight = worksheet.Range(bottomRight).End(XlDirection.xlDown).Address

            selectedRange = worksheet.Range(firstCell, worksheet.Range(bottomRight))


            If selectedRange.Rows.Count = 1 And selectedRange.Columns.Count >= firstInputRng.Columns.Count Then
                selectedRange = worksheet.Range(selectedRange.Cells(1, 1), selectedRange.Cells(1, 1).Offset(0, firstInputRng.Columns.Count - 1))
                selectedRange.Select()

            ElseIf selectedRange.Rows.Count = 1 And selectedRange.Columns.Count < firstInputRng.Columns.Count Then
                selectedRange = worksheet.Range(selectedRange.Cells(1, 1), selectedRange.Cells(1, 1).Offset(0, selectedRange.Columns.Count - 1))
                selectedRange.Select()

            ElseIf selectedRange.Columns.Count = 1 And selectedRange.Rows.Count >= firstInputRng.Rows.Count Then
                selectedRange = worksheet.Range(selectedRange.Cells(1, 1), selectedRange.Cells(1, 1).Offset(firstInputRng.Rows.Count - 1, 0))
                selectedRange.Select()

            ElseIf selectedRange.Columns.Count = 1 And selectedRange.Rows.Count < firstInputRng.Rows.Count Then
                selectedRange = worksheet.Range(selectedRange.Cells(1, 1), selectedRange.Cells(1, 1).Offset(selectedRange.Rows.Count - 1, 0))
                selectedRange.Select()


            Else

                If selectedRange.Rows.Count = firstInputRng.Rows.Count And selectedRange.Columns.Count = firstInputRng.Columns.Count Then


                    firstCell = selectedRange.Cells(1, 1)
                    selectedRange = worksheet.Range(firstCell.Offset(0, 0), firstCell.Offset(firstInputRng.Rows.Count - 1, firstInputRng.Columns.Count - 1))
                    selectedRange.Select()

                ElseIf selectedRange.Rows.Count = firstInputRng.Rows.Count And selectedRange.Columns.Count > firstInputRng.Columns.Count Then



                    firstCell = selectedRange.Cells(1, 1)
                    selectedRange = worksheet.Range(firstCell.Offset(0, 0), firstCell.Offset(firstInputRng.Rows.Count - 1, firstInputRng.Columns.Count - 1))
                    selectedRange.Select()

                ElseIf selectedRange.Rows.Count = firstInputRng.Rows.Count And selectedRange.Columns.Count < firstInputRng.Columns.Count Then
                    firstCell = selectedRange.Cells(1, 1)
                    bottomRight = firstCell.End(XlDirection.xlToRight).Address
                    bottomRight = worksheet.Range(bottomRight).End(XlDirection.xlDown).Address

                    selectedRange = worksheet.Range(firstCell.Offset(0, 0), worksheet.Range(bottomRight))
                    selectedRange.Select()

                ElseIf selectedRange.Rows.Count > firstInputRng.Rows.Count And selectedRange.Columns.Count = firstInputRng.Columns.Count Then

                    firstCell = selectedRange.Cells(1, 1)
                    selectedRange = worksheet.Range(firstCell.Offset(0, 0), firstCell.Offset(firstInputRng.Rows.Count - 1, firstInputRng.Columns.Count - 1))
                    selectedRange.Select()

                ElseIf selectedRange.Rows.Count > firstInputRng.Rows.Count And selectedRange.Columns.Count > firstInputRng.Columns.Count Then

                    firstCell = selectedRange.Cells(1, 1)
                    selectedRange = worksheet.Range(firstCell.Offset(0, 0), firstCell.Offset(firstInputRng.Rows.Count - 1, firstInputRng.Columns.Count - 1))
                    selectedRange.Select()

                ElseIf selectedRange.Rows.Count > firstInputRng.Rows.Count And selectedRange.Columns.Count < firstInputRng.Columns.Count Then

                    firstCell = selectedRange.Cells(1, 1)
                    bottomRight = firstCell.End(XlDirection.xlToRight).Address
                    bottomRight = worksheet.Range(bottomRight).Offset(firstInputRng.Rows.Count - 1, 0).Address

                    selectedRange = worksheet.Range(firstCell.Offset(0, 0), worksheet.Range(bottomRight))
                    selectedRange.Select()

                ElseIf selectedRange.Rows.Count < firstInputRng.Rows.Count And selectedRange.Columns.Count = firstInputRng.Columns.Count Then

                    firstCell = selectedRange.Cells(1, 1)
                    bottomRight = firstCell.End(XlDirection.xlToRight).Address
                    bottomRight = worksheet.Range(bottomRight).End(XlDirection.xlDown).Address

                    selectedRange = worksheet.Range(firstCell.Offset(0, 0), worksheet.Range(bottomRight))
                    selectedRange.Select()
                ElseIf selectedRange.Rows.Count < firstInputRng.Rows.Count And selectedRange.Columns.Count > firstInputRng.Columns.Count Then

                    firstCell = selectedRange.Cells(1, 1)
                    bottomRight = firstCell.Offset(0, firstInputRng.Columns.Count - 1).Address
                    bottomRight = worksheet.Range(bottomRight).End(XlDirection.xlDown).Address

                    selectedRange = worksheet.Range(firstCell.Offset(0, 0), worksheet.Range(bottomRight))
                    selectedRange.Select()


                ElseIf selectedRange.Rows.Count < firstInputRng.Rows.Count And selectedRange.Columns.Count < firstInputRng.Columns.Count Then

                    firstCell = selectedRange.Cells(1, 1)
                    bottomRight = firstCell.End(XlDirection.xlToRight).Address
                    bottomRight = worksheet.Range(bottomRight).End(XlDirection.xlDown).Address

                    selectedRange = worksheet.Range(firstCell.Offset(0, 0), worksheet.Range(bottomRight))
                    selectedRange.Select()

                End If
            End If

        End If


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

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click

        Me.Dispose()

    End Sub

    Private Sub btnOK_Click(sender As Object, e As EventArgs) Handles btnOK.Click

        Try

            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            worksheet = workbook.ActiveSheet
            selectedRange = excelApp.Selection

            worksheet1 = workbook.Sheets(WsName1)
            worksheet2 = workbook.Sheets(WsName2)

            'firstInputRng = worksheet.Range(txtSourceRange1.Text)
            'secondInputRng = worksheet.Range(txtSourceRange2.Text)




            Dim temp As Object
            tempRng = worksheet1.Range("A10000")
            tempRng = worksheet1.Range(tempRng.Cells(1, 1).offset(0, 0), tempRng.Cells(1, 1).offset(firstInputRng.Rows.Count - 1, firstInputRng.Columns.Count - 1))

            If firstInputRng.Rows.Count <> secondInputRng.Rows.Count And firstInputRng.Columns.Count <> secondInputRng.Columns.Count Then

                MsgBox("You must use same number of rows and columns in both ranges.",, "Warning!")

                Me.Dispose()
                Exit Sub

            ElseIf firstInputRng.Rows.Count <> secondInputRng.Rows.Count And firstInputRng.Columns.Count = secondInputRng.Columns.Count Then
                MsgBox("Please match the source range row size.",, "Warning!")

                Me.Dispose()
                Exit Sub
            ElseIf firstInputRng.Rows.Count = secondInputRng.Rows.Count And firstInputRng.Columns.Count <> secondInputRng.Columns.Count Then
                MsgBox("Please match the source range column size.",, "Warning!")

                Me.Dispose()
                Exit Sub

            End If



            If CB_CopyWs.Checked = True Then

                workbook.ActiveSheet.Copy(After:=workbook.Sheets(workbook.Sheets.Count))
                outWorksheet = workbook.Sheets(workbook.Sheets.Count)


                worksheet = workbook.Sheets(WsName1)
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

                            modifiedFormula1 = swapFormulaWithSheetName(worksheet1.Range(firstInputRng.Address).Cells(1, 1).offset(i, j).formula, WsName1)
                            modifiedFormula2 = swapFormulaWithSheetName(worksheet2.Range(secondInputRng.Address).Cells(1, 1).offset(i, j).formula, WsName2)
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

                            modifiedFormula1 = swapFormulaWithSheetName(worksheet1.Range(firstInputRng.Address).Cells(1, 1).offset(i, j).formula, WsName1)
                            modifiedFormula2 = swapFormulaWithSheetName(worksheet2.Range(secondInputRng.Address).Cells(1, 1).offset(i, j).formula, WsName2)
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
        Dim replacement As String

        Dim charToFind As Char = " "c
        Dim index As Integer = sheetName.IndexOf(charToFind)
        If index >= 0 Then
            replacement = "'" & sheetName & "'!$1"
        Else
            replacement = sheetName & "!$1"
        End If

        Return System.Text.RegularExpressions.Regex.Replace(currentFormula, pattern, replacement)
    End Function


    Private Sub txtSourceRange1_Click(sender As Object, e As EventArgs) Handles txtSourceRange1.Click
        txtSourceRange1.SelectionStart = txtSourceRange1.TextLength
        txtSourceRange1.ScrollToCaret()
    End Sub

    Private Sub txtSourceRange2_Click(sender As Object, e As EventArgs) Handles txtSourceRange2.Click
        txtSourceRange2.SelectionStart = txtSourceRange2.TextLength
        txtSourceRange2.ScrollToCaret()
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