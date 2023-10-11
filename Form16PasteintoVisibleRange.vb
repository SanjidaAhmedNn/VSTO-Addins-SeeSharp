Imports System.ComponentModel.Design
Imports System.Security.Cryptography
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Excel
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Drawing
Imports System.ComponentModel
Imports System.Linq.Expressions



Public Class Form16PasteintoVisibleRange
    Dim WithEvents excelApp As Excel.Application
    Dim workbook As Excel.Workbook
    Dim worksheet As Excel.Worksheet
    Dim outWorksheet As Excel.Worksheet
    Dim FocusedTxtBox As Integer
    Dim selectedRange As Excel.Range
    Dim sourceRange, destRange As Excel.Range
    Dim WsName As String
    Dim changeState As Boolean = False
    Dim txtChanged As Boolean = False

    Private Sub Form1_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Enter Then
            btnOK.PerformClick()
        End If
    End Sub

    Private Sub Form16PasteintoVisibleRange_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        excelApp = Globals.ThisAddIn.Application
        workbook = excelApp.ActiveWorkbook
        worksheet = workbook.ActiveSheet

        'Define a varibale to access a selected range
        Dim selectedRng As Excel.Range = excelApp.Selection

        'Assign the address of selected range that is selcted before loading the form in the textbox "txtSourceRange" 
        'Give foucs to the textbox "txtSourceRange" after the form loads
        txtSourceRange.Text = selectedRng.Address
        txtSourceRange.Focus()

        Me.KeyPreview = True

    End Sub

    Private Sub txtSourceRange_TextChanged(sender As Object, e As EventArgs) Handles txtSourceRange.TextChanged

        Try

            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            worksheet = workbook.ActiveSheet


            'MsgBox(txtSourceRange1.Text)
            txtChanged = True
            sourceRange = worksheet.Range(txtSourceRange.Text)


            sourceRange.Select()




            'If changeState = True Then


            '    If destRange.Worksheet.Name <> sourceRange.Worksheet.Name Then

            '        txtDestRange.Text = destRange.Worksheet.Name & "!" & destRange.Address

            '    End If


            'End If



        Catch ex As Exception

        End Try

        txtChanged = False

        txtSourceRange.Focus()


    End Sub
    Private Sub txtDestRange_TextChanged(sender As Object, e As EventArgs) Handles txtDestRange.TextChanged

        Try
            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            worksheet = workbook.ActiveSheet

            changeState = True

            txtChanged = True
            destRange = worksheet.Range(txtDestRange.Text)




            destRange.Select()


            If destRange.Worksheet.Name <> sourceRange.Worksheet.Name Then

                txtDestRange.Text = destRange.Worksheet.Name & "!" & destRange.Address

            End If


        Catch ex As Exception

        End Try

        txtChanged = False
        txtDestRange.Focus()


    End Sub

    Private Sub Selection_Click(sender As Object, e As EventArgs) Handles Selection.Click

        Try

            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            worksheet = workbook.ActiveSheet
            selectedRange = excelApp.Selection
            txtSourceRange.Focus()

            Me.Hide()
            sourceRange = excelApp.InputBox("Please Select the First Range", "First Range Selection", selectedRange.Address, Type:=8)
            Me.Show()



            sourceRange.Worksheet.Activate()


            txtSourceRange.Text = sourceRange.Address

            sourceRange.Select()

            txtSourceRange.Focus()



        Catch ex As Exception

            txtSourceRange.Focus()

        End Try


    End Sub

    Private Sub destinationSelection_Click(sender As Object, e As EventArgs) Handles destinationSelection.Click

        Try
            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            worksheet = workbook.ActiveSheet
            selectedRange = excelApp.Selection
            txtDestRange.Focus()

            Me.Hide()
            destRange = excelApp.InputBox("Please Select the Second Range", "Second Range Selection", selectedRange.Address, Type:=8)
            Me.Show()


            destRange.Worksheet.Activate()

            'txtDestRange.Text = destRange.Address
            txtDestRange.Text = destRange.Worksheet.Name & "!" & destRange.Address

            destRange.Select()
            txtDestRange.Focus()




        Catch ex As Exception

            txtDestRange.Focus()

        End Try




    End Sub



    Private Sub txtSourceRange_GotFocus(sender As Object, e As EventArgs) Handles txtSourceRange.GotFocus
        Try

            'If txtSourceRange textbox got focus, assign 1 to the global variable "FocusedTxtBox"
            FocusedTxtBox = 1


        Catch ex As Exception

        End Try
    End Sub

    Private Sub txtDestRange_GotFocus(sender As Object, e As EventArgs) Handles txtDestRange.GotFocus

        Try

            'If txtDestRange textbox got focus, assign 2 to the global variable "FocusedTxtBox"
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
                    txtSourceRange.Text = selectedRange.Address
                    txtSourceRange.Focus()

                ElseIf FocusedTxtBox = 2 Then
                    txtDestRange.Text = selectedRange.Address
                End If

            End If


        Catch ex As Exception

        End Try


    End Sub

    Private Sub AutoSelection_Click(sender As Object, e As EventArgs) Handles AutoSelection.Click

        Try
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

            sourceRange = selectedRange
            txtSourceRange.Text = sourceRange.Address



        Catch ex As Exception

        End Try


    End Sub



    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.Dispose()
    End Sub


    Public Function IsValidRng(input As String) As Boolean
        '"^(\$?[A-Z]+\$?[0-9]+(:\$?[A-Z]+\$?[0-9]+)?)(,\$?[A-Z]+\$?[0-9]+(:\$?[A-Z]+\$?[0-9]+)?)*$"

        Dim pattern As String = "^(.*!)?(\$?[A-Z]+\$?[0-9]+(:\$?[A-Z]+\$?[0-9]+)?)(,\$?[A-Z]+\$?[0-9]+(:\$?[A-Z]+\$?[0-9]+)?)*$"
        Return System.Text.RegularExpressions.Regex.IsMatch(input, pattern)

    End Function

    Private Sub btnOK_Click(sender As Object, e As EventArgs) Handles btnOK.Click



        If txtSourceRange.Text = "" And txtDestRange.Text = "" Then

            MsgBox("Please select the Source Range and the Destination Range.", MsgBoxStyle.Exclamation, "Error!")
            txtSourceRange.Focus()
            Exit Sub
        ElseIf txtSourceRange.Text = "" And txtDestRange.Text <> "" Then

            If IsValidRng(txtDestRange.Text.ToUpper) = True Then
                MsgBox("Please select data to be copied.", MsgBoxStyle.Exclamation, "Error!")
                txtSourceRange.Focus()
                Exit Sub
            Else
                MsgBox("Please use a valid range in the Destination Range.", MsgBoxStyle.Exclamation, "Error!")
                txtDestRange.Text = ""
                txtDestRange.Focus()
                Exit Sub
            End If

        ElseIf txtDestRange.Text = "" And txtSourceRange.Text <> "" Then
            If IsValidRng(txtSourceRange.Text.ToUpper) = True Then
                MsgBox("Please select the Destination Range.", MsgBoxStyle.Exclamation, "Error!")
                txtDestRange.Focus()
                Exit Sub
            Else
                MsgBox("Please select a valid cell range for data to be copied.", MsgBoxStyle.Exclamation, "Error!")
                txtSourceRange.Text = ""
                txtSourceRange.Focus()
                Exit Sub
            End If

        ElseIf txtSourceRange.Text <> "" And txtDestRange.Text <> "" Then
            If IsValidRng(txtSourceRange.Text.ToUpper) = False And IsValidRng(txtDestRange.Text.ToUpper) = True Then
                MsgBox("Please select a valid cell range for data to be copied.", MsgBoxStyle.Exclamation, "Error!")
                txtSourceRange.Text = ""
                txtSourceRange.Focus()
                Exit Sub

            ElseIf IsValidRng(txtSourceRange.Text.ToUpper) = True And IsValidRng(txtDestRange.Text.ToUpper) = False Then
                MsgBox("Please select a valid cell range for data to be copied.", MsgBoxStyle.Exclamation, "Error!")
                txtDestRange.Text = ""
                txtDestRange.Focus()
                Exit Sub
            ElseIf IsValidRng(txtSourceRange.Text.ToUpper) = False And IsValidRng(txtDestRange.Text.ToUpper) = False Then
                MsgBox("Please select a valid cell range for data to be copied.", MsgBoxStyle.Exclamation, "Error!")
                txtSourceRange.Text = ""
                txtDestRange.Text = ""
                txtSourceRange.Focus()
                Exit Sub

            End If
        End If


        Dim i, j, count, pasteValue, lastRowNum, lastColNum As Integer
        Dim rowNum, colNum As Integer
        Dim rowOff, colOff As Integer
        Dim lastRow, lastCol As String
        excelApp = Globals.ThisAddIn.Application
        workbook = excelApp.ActiveWorkbook
        worksheet = workbook.ActiveSheet
        WsName = worksheet.Name
        destRange = worksheet.Range(txtDestRange.Text).Cells(1, 1)


        If CB_copyWs.Checked = True Then

            workbook.ActiveSheet.Copy(After:=workbook.Sheets(workbook.Sheets.Count))
            outWorksheet = workbook.Sheets(workbook.Sheets.Count)


            worksheet = workbook.Sheets(WsName)
            worksheet.Activate()


        End If


        lastRowNum = 0
        If destRange.End(XlDirection.xlDown).Value Is Nothing Then

            While destRange.Offset(lastRowNum, 0).Value IsNot Nothing

                lastRowNum = lastRowNum + 1

            End While

            lastRowNum = destRange.Row + lastRowNum

        Else
            lastRow = destRange.End(XlDirection.xlDown).Address
            While worksheet.Range(lastRow).Offset(lastRowNum, 0).Value IsNot Nothing

                lastRowNum = lastRowNum + 1

            End While

            lastRowNum = worksheet.Range(lastRow).Row + lastRowNum
        End If

        'finding last column number
        lastColNum = 0
        If destRange.End(XlDirection.xlToRight).Value Is Nothing Then

            While destRange.Offset(0, lastColNum).Value IsNot Nothing

                lastColNum = lastColNum + 1

            End While

            lastColNum = destRange.Column + lastColNum

        Else
            lastCol = destRange.End(XlDirection.xlToRight).Address
            While worksheet.Range(lastCol).Offset(0, lastColNum).Value IsNot Nothing

                lastColNum = lastColNum + 1

            End While

            lastColNum = worksheet.Range(lastCol).Column + lastColNum
        End If





        'finding the total visible rows
        Dim visibleRows As Integer = 0
        For i = destRange.Row To lastRowNum

            If worksheet.Range(worksheet.Cells(i, 1), worksheet.Cells(i, 2)).EntireRow.Hidden = False Then
                visibleRows = visibleRows + 1
            End If


        Next
        visibleRows = visibleRows - 1



        'finding total visible columns
        Dim visibleCols As Integer = 0
        For i = destRange.Column To lastColNum

            If worksheet.Range(worksheet.Cells(1, i), worksheet.Cells(2, i)).EntireColumn.Hidden = False Then
                visibleCols = visibleCols + 1
            End If


        Next
        visibleCols = visibleCols - 1


        count = 0
        pasteValue = 0

        If sourceRange.Rows.Count <= visibleRows And sourceRange.Columns.Count <= visibleCols Then

            rowOff = 0

            rowNum = destRange.Row
            colNum = destRange.Column

            For i = rowNum To lastRowNum



                If worksheet.Cells(i, 1).EntireRow.Hidden = False Then
                    rowOff += 1

                ElseIf worksheet.Cells(i, 1).EntireRow.Hidden = True Then

                    GoTo nextLoop4
                End If
                If rowOff > sourceRange.Rows.Count Then
                    Exit For
                End If

                colOff = 0

                For j = colNum To lastColNum

                    If colOff + 1 > sourceRange.Columns.Count Then
                        Exit For
                    End If

                    '' Check if the destination cell (or its row/column) is not hidden.
                    If worksheet.Cells(i, j).EntireColumn.Hidden = False Then
                        colOff += 1


                        If CB_keepFormat.Checked = True Then
                            sourceRange.Cells(1, 1).offset(rowOff - 1, colOff - 1).Copy(worksheet.Cells(i, j))

                        ElseIf CB_keepFormat.Checked = False Then

                            worksheet.Cells(i, j).value = sourceRange.Cells(1, 1).offset(rowOff - 1, colOff - 1).value

                        End If

                    End If

                Next
nextLoop4:

            Next

            worksheet.Range(worksheet.Cells(destRange.Row, destRange.Column), worksheet.Cells(i - 1, j - 1)).Select()




        ElseIf sourceRange.Rows.Count <= visibleRows And sourceRange.Columns.Count > visibleCols Then

            rowOff = 0

            rowNum = destRange.Row
            colNum = destRange.Column

            For i = rowNum To lastRowNum



                If worksheet.Cells(i, 1).EntireRow.Hidden = False Then
                    rowOff += 1

                ElseIf worksheet.Cells(i, 1).EntireRow.Hidden = True Then

                    GoTo nextLoop
                End If
                If rowOff > sourceRange.Rows.Count Then
                    Exit For
                End If

                colOff = 0

                For j = colNum To lastColNum + sourceRange.Columns.Count - visibleCols - 1



                    If colOff > sourceRange.Columns.Count Then
                        Exit For
                    End If


                    '' Check if the destination cell (or its row/column) is not hidden.
                    If worksheet.Cells(i, j).EntireColumn.Hidden = False Then
                        colOff += 1


                        If CB_keepFormat.Checked = True Then
                            sourceRange.Cells(1, 1).offset(rowOff - 1, colOff - 1).Copy(worksheet.Cells(i, j))

                        ElseIf CB_keepFormat.Checked = False Then

                            worksheet.Cells(i, j).value = sourceRange.Cells(1, 1).offset(rowOff - 1, colOff - 1).value

                        End If

                    End If

                Next
nextLoop:

            Next

            worksheet.Range(worksheet.Cells(destRange.Row, destRange.Column), worksheet.Cells(i - 1, j - 1)).Select()

        ElseIf sourceRange.Rows.Count > visibleRows And sourceRange.Columns.Count <= visibleCols Then



            rowOff = 0

            rowNum = destRange.Row
            colNum = destRange.Column

            For i = rowNum To lastRowNum + sourceRange.Rows.Count - visibleRows - 1


                If worksheet.Cells(i, 1).EntireRow.Hidden = False Then
                    rowOff += 1

                ElseIf worksheet.Cells(i, 1).EntireRow.Hidden = True Then

                    GoTo nextLoop2
                End If

                If rowOff > sourceRange.Rows.Count Then
                    Exit For
                End If


                colOff = 0

                For j = colNum To lastColNum

                    If colOff > sourceRange.Columns.Count Then
                        Exit For
                    End If


                    '' Check if the destination cell (or its row/column) is not hidden.
                    If worksheet.Cells(i, j).EntireColumn.Hidden = False Then
                        colOff += 1


                        If CB_keepFormat.Checked = True Then
                            sourceRange.Cells(1, 1).offset(rowOff - 1, colOff - 1).Copy(worksheet.Cells(i, j))

                        ElseIf CB_keepFormat.Checked = False Then

                            worksheet.Cells(i, j).value = sourceRange.Cells(1, 1).offset(rowOff - 1, colOff - 1).value

                        End If

                    End If

                Next
nextLoop2:

            Next

            worksheet.Range(worksheet.Cells(destRange.Row, destRange.Column), worksheet.Cells(i - 1, j - 1)).Select()


        Else

            rowOff = 0

            rowNum = destRange.Row
            colNum = destRange.Column

            For i = rowNum To lastRowNum + sourceRange.Rows.Count - visibleRows - 1

                If rowOff > sourceRange.Rows.Count Then
                    Exit For
                End If

                If worksheet.Cells(i, 1).EntireRow.Hidden = False Then
                    rowOff += 1

                ElseIf worksheet.Cells(i, 1).EntireRow.Hidden = True Then

                    GoTo nextLoop3
                End If

                colOff = 0

                For j = colNum To lastColNum + sourceRange.Columns.Count - visibleCols - 1

                    If colOff > sourceRange.Columns.Count Then
                        Exit For
                    End If


                    '' Check if the destination cell (or its row/column) is not hidden.
                    If worksheet.Cells(i, j).EntireColumn.Hidden = False Then
                        colOff += 1

                        If CB_keepFormat.Checked = True Then
                            sourceRange.Cells(1, 1).offset(rowOff - 1, colOff - 1).Copy(worksheet.Cells(i, j))

                        ElseIf CB_keepFormat.Checked = False Then

                            worksheet.Cells(i, j).value = sourceRange.Cells(1, 1).offset(rowOff - 1, colOff - 1).value

                        End If

                    End If

                Next
nextLoop3:

            Next

            worksheet.Range(worksheet.Cells(destRange.Row, destRange.Column), worksheet.Cells(i - 1, j - 1)).Select()

        End If


        Me.Dispose()





    End Sub

End Class

'UNUSED CODES

'While destRange.Offset(count, 0).Value IsNot Nothing

'    If destRange.Offset(count, count2).EntireRow.Hidden = False Then
'        pasteValue = pasteValue + 1
'        count2 = 0
'        pasteValue2 = 0

'    End If
'    If pasteValue > sourceRange.Rows.Count Then
'        Exit While
'    End If

'    While destRange.Offset(count, count2).Value <> Nothing
'        If pasteValue2 + 1 > sourceRange.Columns.Count Then
'            Exit While
'        End If


'        If destRange.Offset(count, count2).EntireRow.Hidden = False And destRange.Offset(count, count2).EntireColumn.Hidden = False Then
'            pasteValue2 = pasteValue2 + 1


'            If CB_keepFormat.Checked = True Then

'                'Call copyCell(destRange, count, count2, worksheet.Range(txtSourceRange.Text).Cells(1, 1), pasteValue - 1, pasteValue2 - 1)
'                sourceRange.Cells(1, 1).offset(pasteValue - 1, pasteValue2 - 1).copy(destRange.Cells(1, 1).offset(count, count2))

'            Else
'                'sourceRange.Cells(1, 1).offset(pasteValue - 1, pasteValue2 - 1).copy
'                'destRange.Cells(1, 1).offset(count, count2).PasteSpecial(Excel.XlPasteType.xlPasteValues)
'                destRange.Offset(count, count2).Value = sourceRange.Cells(1, 1).offset(pasteValue - 1, pasteValue2 - 1).value


'            End If


'        End If

'        count2 += 1

'    End While

'    count += 1

'End While





'            For j = destRange.Row To lastRowNum

'                While destRange.Offset(count, 0).Value <> Nothing

'                    If destRange.Offset(count, count2).EntireRow.Hidden = False Then
'                        pasteValue = pasteValue + 1
'                        count2 = 0
'                        pasteValue2 = 0

'                    End If
'                    If pasteValue > sourceRange.Rows.Count Then
'                        Exit While
'                    End If

'                    While destRange.Offset(count, count2).Value <> Nothing
'                        If pasteValue2 + 1 > sourceRange.Columns.Count Then
'                            Exit While
'                        End If


'                        If destRange.Offset(count, count2).EntireRow.Hidden = False And destRange.Offset(count, count2).EntireColumn.Hidden = False Then
'                            pasteValue2 = pasteValue2 + 1
'                            'If CB_keepFormat.Checked = True Then

'                            '    Call copyCell(destRange, count, count2, worksheet.Range(txtSourceRange.Text).Cells(1, 1), pasteValue - 1, pasteValue2 - 1)

'                            '    'Dim borderIndices As Excel.XlBordersIndex() = {Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBordersIndex.xlEdgeTop}
'                            '    'Dim sourceCell As Excel.Range = sourceRange.Cells(1, 1).Offset(pasteValue - 1, pasteValue2 - 1)
'                            '    'Dim destCell As Excel.Range = destRange.Cells(1, 1).Offset(count, count2)


'                            '    'For Each borderIndex As Excel.XlBordersIndex In borderIndices
'                            '    '    Dim sourceBorder As Excel.Border = sourceCell.Borders(borderIndex)
'                            '    '    Dim destBorder As Excel.Border = destCell.Borders(borderIndex)

'                            '    '    If sourceBorder.LineStyle = Excel.XlLineStyle.xlLineStyleNone Then
'                            '    '        destBorder.LineStyle = Excel.XlLineStyle.xlLineStyleNone
'                            '    '    Else
'                            '    '        ' Copying the line style
'                            '    '        destBorder.LineStyle = sourceBorder.LineStyle

'                            '    '        ' Copying the color
'                            '    '        destBorder.Color = sourceBorder.Color

'                            '    '        ' Copying the weight
'                            '    '        destBorder.Weight = sourceBorder.Weight

'                            '    '        ' Copying the TintAndShade
'                            '    '        destBorder.TintAndShade = sourceBorder.TintAndShade
'                            '    '    End If

'                            '    'Next

'                            'Else
'                            '    destRange.Offset(count, count2).Value = worksheet.Range(txtSourceRange.Text).Cells(1, 1).offset(pasteValue - 1, pasteValue2 - 1).value

'                            'End If

'                            destRange.Offset(count, count2).Value = worksheet.Range(txtSourceRange.Text).Cells(1, 1).offset(pasteValue - 1, pasteValue2 - 1).value



'                        End If

'                        count2 = count2 + 1

'                    End While

'                    count = count + 1

'                End While

'            Next





'            Dim count3, count4, count5, l As Integer

'            count3 = 0

'            For k = lastRowNum To lastRowNum + sourceRange.Rows.Count - visibleRows - 1
'                count4 = 0
'                count5 = 0
'                For l = 1 To lastColNum + sourceRange.Columns.Count - visibleCols - 1

'                    If worksheet.Cells(lastRowNum, destRange.Column).Offset(count3, l - 1).EntireColumn.Hidden = False Then
'                        count5 = count5 + 1
'                    End If


'                    If count5 > sourceRange.Columns.Count Then
'                        Exit For
'                    End If

'                    If worksheet.Cells(lastRowNum, destRange.Column).Offset(count3, l - 1).EntireColumn.Hidden = False Then

'                        'If CB_keepFormat.Checked = True Then

'                        '    Call copyCell(worksheet.Cells(lastRowNum, destRange.Column), count3, l - 1, worksheet.Range(txtSourceRange.Text).Cells(1, 1), visibleRows + count3, count4)

'                        '    'Dim borderIndices As Excel.XlBordersIndex() = {Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBordersIndex.xlEdgeTop}
'                        '    'Dim sourceCell As Excel.Range = sourceRange.Cells(1, 1).Offset(visibleRows + count3, count4)
'                        '    'Dim destCell As Excel.Range = destRange.Cells(1, 1).Offset(count3, l - 1)


'                        '    'For Each borderIndex As Excel.XlBordersIndex In borderIndices
'                        '    '    Dim sourceBorder As Excel.Border = sourceCell.Borders(borderIndex)
'                        '    '    Dim destBorder As Excel.Border = destCell.Borders(borderIndex)

'                        '    '    If sourceBorder.LineStyle = Excel.XlLineStyle.xlLineStyleNone Then
'                        '    '        destBorder.LineStyle = Excel.XlLineStyle.xlLineStyleNone
'                        '    '    Else
'                        '    '        ' Copying the line style
'                        '    '        destBorder.LineStyle = sourceBorder.LineStyle

'                        '    '        ' Copying the color
'                        '    '        destBorder.Color = sourceBorder.Color

'                        '    '        ' Copying the weight
'                        '    '        destBorder.Weight = sourceBorder.Weight

'                        '    '        ' Copying the TintAndShade
'                        '    '        destBorder.TintAndShade = sourceBorder.TintAndShade
'                        '    '    End If

'                        '    'Next


'                        'Else
'                        '    worksheet.Cells(lastRowNum, destRange.Column).Offset(count3, l - 1).Value = worksheet.Range(txtSourceRange.Text).Cells(1, 1).offset(visibleRows + count3, count4).value

'                        'End If

'                        worksheet.Cells(lastRowNum, destRange.Column).Offset(count3, l - 1).Value = worksheet.Range(txtSourceRange.Text).Cells(1, 1).offset(visibleRows + count3, count4).value



'                        count4 = count4 + 1
'                    End If

'                Next
'                count3 = count3 + 1
'            Next




'            rowNum = destRange.Row
'            colNum = destRange.Column
'            count3 = 0
'            count4 = visibleCols
'            For k = destRange.Row To lastRowNum - 1

'                If worksheet.Range(worksheet.Cells(k, 1), worksheet.Cells(k, 2)).EntireRow.Hidden = False Then

'                    rowNum = worksheet.Range(worksheet.Cells(k, 1), worksheet.Cells(k, 2)).Row

'                End If

'                If count3 + 1 > sourceRange.Rows.Count Then
'                    Exit For
'                End If


'                If Not worksheet.Range(worksheet.Cells(k, 1), worksheet.Cells(k, 2)).EntireRow.Hidden = False And worksheet.Range(worksheet.Cells(k, 1), worksheet.Cells(k + 1, 1)).EntireColumn.Hidden = False Then

'                    GoTo exitLoop

'                End If

'                count4 = visibleCols


'                For l = lastColNum To lastColNum + sourceRange.Columns.Count - visibleCols - 1


'                    If worksheet.Range(worksheet.Cells(k, l), worksheet.Cells(k + 1, l)).EntireColumn.Hidden = False Then

'                        colNum = worksheet.Range(worksheet.Cells(k, l), worksheet.Cells(k + 1, l)).Column

'                    End If
'                    If count4 + 1 > sourceRange.Columns.Count Then
'                        Exit For
'                    End If


'                    If worksheet.Range(worksheet.Cells(k, l), worksheet.Cells(k, l + 1)).EntireRow.Hidden = False And worksheet.Range(worksheet.Cells(k, l), worksheet.Cells(k + 1, l)).EntireColumn.Hidden = False Then

'                        'If CB_keepFormat.Checked = True Then

'                        '    Call copyCell(worksheet.Range(worksheet.Cells(rowNum, colNum), worksheet.Cells(rowNum, colNum)), 0, 0, worksheet.Range(txtSourceRange.Text).Cells(1, 1), count3, count4)

'                        '    'Dim borderIndices As Excel.XlBordersIndex() = {Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBordersIndex.xlEdgeTop}
'                        '    'Dim sourceCell As Excel.Range = sourceRange.Cells(1, 1).Offset(count3, count4)
'                        '    'Dim destCell As Excel.Range = destRange.Cells(1, 1).Offset(0, 0)


'                        '    'For Each borderIndex As Excel.XlBordersIndex In borderIndices
'                        '    '    Dim sourceBorder As Excel.Border = sourceCell.Borders(borderIndex)
'                        '    '    Dim destBorder As Excel.Border = destCell.Borders(borderIndex)

'                        '    '    If sourceBorder.LineStyle = Excel.XlLineStyle.xlLineStyleNone Then
'                        '    '        destBorder.LineStyle = Excel.XlLineStyle.xlLineStyleNone
'                        '    '    Else
'                        '    '        ' Copying the line style
'                        '    '        destBorder.LineStyle = sourceBorder.LineStyle

'                        '    '        ' Copying the color
'                        '    '        destBorder.Color = sourceBorder.Color

'                        '    '        ' Copying the weight
'                        '    '        destBorder.Weight = sourceBorder.Weight

'                        '    '        ' Copying the TintAndShade
'                        '    '        destBorder.TintAndShade = sourceBorder.TintAndShade
'                        '    '    End If

'                        '    'Next


'                        'Else
'                        '    worksheet.Range(worksheet.Cells(rowNum, colNum), worksheet.Cells(rowNum, colNum)).Offset(0, 0).Value = worksheet.Range(txtSourceRange.Text).Cells(1, 1).offset(count3, count4).value

'                        'End If

'                        worksheet.Range(worksheet.Cells(rowNum, colNum), worksheet.Cells(rowNum, colNum)).Offset(0, 0).Value = worksheet.Range(txtSourceRange.Text).Cells(1, 1).offset(count3, count4).value



'                        'worksheet.Range(worksheet.Cells(rowNum, colNum), worksheet.Cells(rowNum, colNum)).Value = sourceRange.Cells.Offset(count3, count4).Value
'                        'sourceRange.Cells.Offset(count3, count4).Copy(worksheet.Cells(rowNum, colNum))

'                    End If
'                    count4 = count4 + 1

'                Next
'                count3 = count3 + 1
'exitLoop:
'            Next

'            If CB_keepFormat.Checked = True Then
'                Dim rowOff, colOff As Integer
'                rowOff = 0

'                rowNum = destRange.Row
'                colNum = destRange.Column

'                For i = rowNum To lastRowNum + sourceRange.Rows.Count - visibleRows - 1

'                    If worksheet.Cells(i, 1).EntireRow.Hidden = False Then
'                        rowOff += 1

'                    ElseIf worksheet.Cells(i, 1).EntireRow.Hidden = True Then

'                        GoTo nextLoop
'                    End If
'                    colOff = 0

'                    For j = colNum To lastColNum + sourceRange.Columns.Count - visibleCols - 1

'                        'Dim sourceCell As Excel.Range = sourceRange.Cells(i, j)
'                        'Dim destCell As Excel.Range = destRange.Cells(i, j)




'                        '' Check if the destination cell (or its row/column) is not hidden.
'                        If worksheet.Cells(i, j).EntireColumn.Hidden = False Then
'                            colOff += 1
'                            ' Copy only the formatting.
'                            sourceRange.Cells(1, 1).offset(rowOff - 1, colOff - 1).Copy()
'                            worksheet.Cells(i, j).PasteSpecial(Excel.XlPasteType.xlPasteFormats)


'                        End If



'                    Next
'nextLoop:


'                Next


'            End If


'worksheet.Range(worksheet.Cells(destRange.Row, destRange.Column), worksheet.Cells(lastRowNum + sourceRange.Rows.Count - visibleRows - 1, lastColNum + sourceRange.Columns.Count - visibleCols - 1)).Select()















'Public Sub copyCell(ByVal destRng As Range, ByVal destOff1 As Integer, ByVal destOff2 As Integer, ByVal srcRng As Range, ByVal srcOff1 As Integer, ByVal srcOff2 As Integer)

'    destRng.Offset(destOff1, destOff2).Font.Name = srcRng.Offset(srcOff1, srcOff2).Font.Name
'    destRng.Offset(destOff1, destOff2).Font.Size = srcRng.Offset(srcOff1, srcOff2).Font.Size
'    destRng.Offset(destOff1, destOff2).Font.Color = srcRng.Offset(srcOff1, srcOff2).Font.Color
'    destRng.Offset(destOff1, destOff2).NumberFormat = srcRng.Offset(srcOff1, srcOff2).NumberFormat

'    If Not srcRng.Offset(srcOff1, srcOff2).Interior.ColorIndex = -4142 Then

'        destRng.Offset(destOff1, destOff2).Interior.Color = srcRng.Offset(srcOff1, srcOff2).Interior.Color

'    End If


'    'bold,italic,underline
'    destRng.Offset(destOff1, destOff2).Font.FontStyle = srcRng.Offset(srcOff1, srcOff2).Font.FontStyle
'    destRng.Offset(destOff1, destOff2).Font.Underline = srcRng.Offset(srcOff1, srcOff2).Font.Underline




'    'border

'    Dim borderIndices As Excel.XlBordersIndex() = {Excel.XlBordersIndex.xlEdgeBottom, Excel.XlBordersIndex.xlEdgeLeft, Excel.XlBordersIndex.xlEdgeRight, Excel.XlBordersIndex.xlEdgeTop}
'    Dim sourceCell As Excel.Range = sourceRange.Cells(1, 1).Offset(srcOff1, srcOff2)
'    Dim destCell As Excel.Range = destRange.Cells(1, 1).Offset(destOff1, destOff2)


'    For Each borderIndex As Excel.XlBordersIndex In borderIndices
'        Dim sourceBorder As Excel.Border = sourceCell.Borders(borderIndex)
'        Dim destBorder As Excel.Border = destCell.Borders(borderIndex)

'        If sourceBorder.LineStyle = Excel.XlLineStyle.xlLineStyleNone Then

'            destBorder.LineStyle = Excel.XlLineStyle.xlLineStyleNone
'        Else
'            ' Copying the line style
'            destBorder.LineStyle = sourceBorder.LineStyle

'            ' Copying the color
'            destBorder.Color = sourceBorder.Color

'            ' Copying the weight
'            destBorder.Weight = sourceBorder.Weight

'            ' Copying the TintAndShade
'            destBorder.TintAndShade = sourceBorder.TintAndShade
'        End If

'    Next


'    'value
'    destRng.Offset(destOff1, destOff2).Value = srcRng.Offset(srcOff1, srcOff2).Value

'    'gridline
'    excelApp.ActiveWindow.DisplayGridlines = True




'End Sub

