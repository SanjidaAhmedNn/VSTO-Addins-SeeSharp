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
    Dim inputRng As Excel.Range
    Dim FocusedTxtBox As Integer
    Dim selectedRange As Excel.Range
    Dim destRange As Excel.Range
    Dim outputRng As Excel.Range
    Dim WsName As String


    Private Sub Form16PasteintoVisibleRange_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        excelApp = Globals.ThisAddIn.Application
        workbook = excelApp.ActiveWorkbook
        worksheet = workbook.ActiveSheet

        'Define a varibale to access a selected range
        Dim selectedRng As Excel.Range = excelApp.Selection

        'Assign the address of selected range that is selcted before loading the form in the textbox "txtSourceRange" 
        'keep foucs on the textbox "txtSourceRange" after the form loads
        txtSourceRange.Text = selectedRng.Address
        txtSourceRange.Focus()


    End Sub

    Private Sub txtSourceRange_TextChanged(sender As Object, e As EventArgs) Handles txtSourceRange.TextChanged

        Try
            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            worksheet = workbook.ActiveSheet

            'keep foucs on the textbox "txtSourceRange" if any text is changed in it
            txtSourceRange.Focus()

            'convert the text of the textbox to range and store it in the public variable "inputRng"
            inputRng = worksheet.Range(txtSourceRange.Text)


        Catch ex As Exception

        End Try




    End Sub

    Private Sub Selection_Click(sender As Object, e As EventArgs) Handles Selection.Click
        Try

            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            worksheet = workbook.ActiveSheet
            selectedRange = excelApp.Selection

            ' Give focus to "txtSourceRange" textbox
            txtSourceRange.Focus()

            ' Display an InputBox to prompt the user to select a range and store the address in the public variable "inputRng" as range
            ' Select the range provided by the user
            ' Display the selected range address in the "txtSourceRange" textbox
            inputRng = excelApp.InputBox("Please Select a Range", "Range Selection", selectedRange.Address, Type:=8)
            inputRng.Select()
            txtSourceRange.Text = inputRng.Address

            'Return focus to "txtSourceRange" textbox after displaying the selected range
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

            ' Give focus to "txtDestRange" textbox
            txtDestRange.Focus()

            ' Display an InputBox to prompt the user to select a range and store the address in the public variable "outputRng" as range
            ' Select the range provided by the user
            ' Display the selected range address in the "txtDestRange" textbox
            outputRng = excelApp.InputBox("Please Select a Destination Range", "Destination Range Selection", selectedRange.Address, Type:=8)
            outputRng.Select()
            txtDestRange.Text = outputRng.Address

            'Return focus to "txtDestRange" textbox after displaying the selected range
            txtDestRange.Focus()


        Catch ex As Exception

            'Return focus to "txtDestRange" textbox if any error occurs 
            'This will keep the form visible on the scrren and allow user to enter correct destination range
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

            selectedRange = excelApp.Selection
            selectedRange.Select()

            'If value of FocusedTxtBox is 1, that means txtSourceRange textbox is selected, and 2 means txtDestRange textbox is selected
            If FocusedTxtBox = 1 Then

                'Display the address of the selected range in the txtSourceRange textbox
                'Store the address in the public variable "inputRng" as range
                txtSourceRange.Text = selectedRange.Address
                worksheet = workbook.ActiveSheet
                inputRng = selectedRange

                'Return focus to "txtSourceRange" textbox after displaying the selected range
                txtSourceRange.Focus()

            ElseIf FocusedTxtBox = 2 Then

                'Display the address of the selected range in the txtDestRange textbox
                'Store the address in the public variable "destRange" as range
                txtDestRange.Text = selectedRange.Address
                worksheet = workbook.ActiveSheet
                destRange = selectedRange

                'Return focus to "txtDestRange" textbox after displaying the selected range
                txtDestRange.Focus()

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





        Catch ex As Exception

        End Try




    End Sub

    Private Sub txtDestRange_TextChanged(sender As Object, e As EventArgs) Handles txtDestRange.TextChanged

        Try
            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            worksheet = workbook.ActiveSheet
            selectedRange = excelApp.Selection

            txtDestRange.Focus()


            outputRng = worksheet.Range(txtDestRange.Text)



        Catch ex As Exception

        End Try


    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.Dispose()
    End Sub

    Private Sub btnOK_Click(sender As Object, e As EventArgs) Handles btnOK.Click


        If txtDestRange.Text = Nothing Or txtSourceRange.Text = Nothing Then

            MsgBox("Please enter a valid Range.",, "Warning!")
            Me.Dispose()
            Exit Sub

        End If


        Dim i, j, count, pasteValue, pasteValue2, count2, lastRowNum, lastColNum As Integer
        Dim lastRow, lastCol As String
        excelApp = Globals.ThisAddIn.Application
        workbook = excelApp.ActiveWorkbook
        worksheet = workbook.ActiveSheet
        WsName = worksheet.Name



        If CB_copyWs.Checked = True Then

            workbook.ActiveSheet.Copy(After:=workbook.Sheets(workbook.Sheets.Count))
            outWorksheet = workbook.Sheets(workbook.Sheets.Count)


            worksheet = workbook.Sheets(WsName)
            worksheet.Activate()


        End If


        lastRowNum = 0
        If worksheet.Range(txtDestRange.Text).End(XlDirection.xlDown).Value Is Nothing Then

            While worksheet.Range(txtDestRange.Text).Offset(lastRowNum, 0).Value <> Nothing

                lastRowNum = lastRowNum + 1

            End While

            lastRowNum = worksheet.Range(txtDestRange.Text).Row + lastRowNum

        Else
            lastRow = worksheet.Range(txtDestRange.Text).End(XlDirection.xlDown).Address
            While worksheet.Range(lastRow).Offset(lastRowNum, 0).Value <> Nothing

                lastRowNum = lastRowNum + 1

            End While

            lastRowNum = worksheet.Range(lastRow).Row + lastRowNum
        End If

        'finding last column number
        lastColNum = 0
        If worksheet.Range(txtDestRange.Text).End(XlDirection.xlToRight).Value Is Nothing Then

            While worksheet.Range(txtDestRange.Text).Offset(0, lastColNum).Value <> Nothing

                lastColNum = lastColNum + 1

            End While

            lastColNum = worksheet.Range(txtDestRange.Text).Column + lastColNum

        Else
            lastCol = worksheet.Range(txtDestRange.Text).End(XlDirection.xlToRight).Address
            While worksheet.Range(lastCol).Offset(0, lastColNum).Value <> Nothing

                lastColNum = lastColNum + 1

            End While

            lastColNum = worksheet.Range(lastCol).Column + lastColNum
        End If





        'finding the total visible rows
        Dim visibleRows As Integer = 0
        For i = worksheet.Range(txtDestRange.Text).Row To lastRowNum

            If worksheet.Range(worksheet.Cells(i, 1), worksheet.Cells(i, 2)).EntireRow.Hidden = False Then
                visibleRows = visibleRows + 1
            End If


        Next
        visibleRows = visibleRows - 1



        'finding total visible columns
        Dim visibleCols As Integer = 0
        For i = worksheet.Range(txtDestRange.Text).Column To lastColNum

            If worksheet.Range(worksheet.Cells(1, i), worksheet.Cells(2, i)).EntireColumn.Hidden = False Then
                visibleCols = visibleCols + 1
            End If


        Next
        visibleCols = visibleCols - 1


        count = 0
        pasteValue = 0

        If inputRng.Rows.Count <= visibleRows And inputRng.Columns.Count <= visibleCols Then



            While worksheet.Range(txtDestRange.Text).Offset(count, 0).Value <> Nothing

                If worksheet.Range(txtDestRange.Text).Offset(count, count2).EntireRow.Hidden = False Then
                    pasteValue = pasteValue + 1
                    count2 = 0
                    pasteValue2 = 0

                End If
                If pasteValue > inputRng.Rows.Count Then
                    Exit While
                End If

                While worksheet.Range(txtDestRange.Text).Offset(count, count2).Value <> Nothing
                    If pasteValue2 + 1 > inputRng.Columns.Count Then
                        Exit While
                    End If


                    If worksheet.Range(txtDestRange.Text).Offset(count, count2).EntireRow.Hidden = False And worksheet.Range(txtDestRange.Text).Offset(count, count2).EntireColumn.Hidden = False Then
                        pasteValue2 = pasteValue2 + 1


                        If CB_keepFormat.Checked = True Then

                            Call copyCell(worksheet.Range(txtDestRange.Text), count, count2, worksheet.Range(txtSourceRange.Text).Cells(1, 1), pasteValue - 1, pasteValue2 - 1)


                        Else

                            worksheet.Range(txtSourceRange.Text).Cells(1, 1).offset(pasteValue - 1, pasteValue2 - 1).Value = worksheet.Range(txtSourceRange.Text).Cells(1, 1).offset(pasteValue - 1, pasteValue2 - 1).value

                        End If

                    End If

                    count2 = count2 + 1

                End While

                count = count + 1

            End While



        Else

            For j = worksheet.Range(txtDestRange.Text).Row To lastRowNum

                While worksheet.Range(txtDestRange.Text).Offset(count, 0).Value <> Nothing

                    If worksheet.Range(txtDestRange.Text).Offset(count, count2).EntireRow.Hidden = False Then
                        pasteValue = pasteValue + 1
                        count2 = 0
                        pasteValue2 = 0

                    End If
                    If pasteValue > inputRng.Rows.Count Then
                        Exit While
                    End If

                    While worksheet.Range(txtDestRange.Text).Offset(count, count2).Value <> Nothing
                        If pasteValue2 + 1 > inputRng.Columns.Count Then
                            Exit While
                        End If


                        If worksheet.Range(txtDestRange.Text).Offset(count, count2).EntireRow.Hidden = False And worksheet.Range(txtDestRange.Text).Offset(count, count2).EntireColumn.Hidden = False Then
                            pasteValue2 = pasteValue2 + 1
                            If CB_keepFormat.Checked = True Then

                                Call copyCell(worksheet.Range(txtDestRange.Text), count, count2, worksheet.Range(txtSourceRange.Text).Cells(1, 1), pasteValue - 1, pasteValue2 - 1)
                            Else
                                worksheet.Range(txtDestRange.Text).Offset(count, count2).Value = worksheet.Range(txtSourceRange.Text).Cells(1, 1).offset(pasteValue - 1, pasteValue2 - 1).value

                            End If
                        End If

                        count2 = count2 + 1

                    End While

                    count = count + 1

                End While

            Next





            Dim count3, count4, count5, l As Integer

            count3 = 0

            For k = lastRowNum To lastRowNum + inputRng.Rows.Count - visibleRows - 1
                count4 = 0
                count5 = 0
                For l = 1 To lastColNum + inputRng.Columns.Count - visibleCols - 1

                    If worksheet.Cells(lastRowNum, worksheet.Range(txtDestRange.Text).Column).Offset(count3, l - 1).EntireColumn.Hidden = False Then
                        count5 = count5 + 1
                    End If


                    If count5 > inputRng.Columns.Count Then
                        Exit For
                    End If

                    If worksheet.Cells(lastRowNum, worksheet.Range(txtDestRange.Text).Column).Offset(count3, l - 1).EntireColumn.Hidden = False Then

                        If CB_keepFormat.Checked = True Then

                            Call copyCell(worksheet.Cells(lastRowNum, worksheet.Range(txtDestRange.Text).Column), count3, l - 1, worksheet.Range(txtSourceRange.Text).Cells(1, 1), visibleRows + count3, count4)
                        Else
                            worksheet.Cells(lastRowNum, worksheet.Range(txtDestRange.Text).Column).Offset(count3, l - 1).Value = worksheet.Range(txtSourceRange.Text).Cells(1, 1).offset(visibleRows + count3, count4).value

                        End If

                        count4 = count4 + 1
                    End If

                Next
                count3 = count3 + 1
            Next







            Dim rowNum, colNum As Integer
            rowNum = worksheet.Range(txtDestRange.Text).Row
            colNum = worksheet.Range(txtDestRange.Text).Column
            count3 = 0
            count4 = visibleCols
            For k = worksheet.Range(txtDestRange.Text).Row To lastRowNum - 1

                If worksheet.Range(worksheet.Cells(k, 1), worksheet.Cells(k, 2)).EntireRow.Hidden = False Then

                    rowNum = worksheet.Range(worksheet.Cells(k, 1), worksheet.Cells(k, 2)).Row

                End If

                If count3 + 1 > inputRng.Rows.Count Then
                    Exit For
                End If


                If Not worksheet.Range(worksheet.Cells(k, 1), worksheet.Cells(k, 2)).EntireRow.Hidden = False And worksheet.Range(worksheet.Cells(k, 1), worksheet.Cells(k + 1, 1)).EntireColumn.Hidden = False Then

                    GoTo exitLoop

                End If

                count4 = visibleCols


                For l = lastColNum To lastColNum + inputRng.Columns.Count - visibleCols - 1


                    If worksheet.Range(worksheet.Cells(k, l), worksheet.Cells(k + 1, l)).EntireColumn.Hidden = False Then

                        colNum = worksheet.Range(worksheet.Cells(k, l), worksheet.Cells(k + 1, l)).Column

                    End If
                    If count4 + 1 > inputRng.Columns.Count Then
                        Exit For
                    End If


                    If worksheet.Range(worksheet.Cells(k, l), worksheet.Cells(k, l + 1)).EntireRow.Hidden = False And worksheet.Range(worksheet.Cells(k, l), worksheet.Cells(k + 1, l)).EntireColumn.Hidden = False Then

                        If CB_keepFormat.Checked = True Then

                            Call copyCell(worksheet.Range(worksheet.Cells(rowNum, colNum), worksheet.Cells(rowNum, colNum)), 0, 0, worksheet.Range(txtSourceRange.Text).Cells(1, 1), count3, count4)
                        Else
                            worksheet.Range(worksheet.Cells(rowNum, colNum), worksheet.Cells(rowNum, colNum)).Offset(0, 0).Value = worksheet.Range(txtSourceRange.Text).Cells(1, 1).offset(count3, count4).value

                        End If




                        'worksheet.Range(worksheet.Cells(rowNum, colNum), worksheet.Cells(rowNum, colNum)).Value = inputRng.Cells.Offset(count3, count4).Value
                        'inputRng.Cells.Offset(count3, count4).Copy(worksheet.Cells(rowNum, colNum))

                    End If
                    count4 = count4 + 1

                Next
                count3 = count3 + 1
exitLoop:
            Next


        End If








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
        destRng.Offset(destOff1, destOff2).Value = srcRng.Offset(srcOff1, srcOff2).Value
    End Sub

End Class