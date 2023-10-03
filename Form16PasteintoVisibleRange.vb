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




            If changeState = True Then


                If destRange.Worksheet.Name <> sourceRange.Worksheet.Name Then

                    txtDestRange.Text = destRange.Worksheet.Name & "!" & destRange.Address

                End If


            End If



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



            'firstInputRng.Worksheet.Activate()


            txtSourceRange.Text = sourceRange.Worksheet.Name & "!" & sourceRange.Address

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

        Dim pattern As String = "^(\$?[A-Z]+\$?[0-9]+(:\$?[A-Z]+\$?[0-9]+)?)(,\$?[A-Z]+\$?[0-9]+(:\$?[A-Z]+\$?[0-9]+)?)*$"
        Return System.Text.RegularExpressions.Regex.IsMatch(input, pattern)

    End Function

    Private Sub btnOK_Click(sender As Object, e As EventArgs) Handles btnOK.Click



        If txtSourceRange.Text = "" And txtDestRange.Text = "" Then

            MsgBox("Please select the Source Range and the Destination Range.", MsgBoxStyle.Exclamation, "Error!")
            txtSourceRange.Focus()
            Exit Sub
        ElseIf txtSourceRange.Text = "" And txtDestRange.Text <> "" Then

            If IsValidRng(txtDestRange.Text.ToUpper) = True Then
                MsgBox("Please select the Source Range.", MsgBoxStyle.Exclamation, "Error!")
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
                MsgBox("Please use a valid range in the Source Range.", MsgBoxStyle.Exclamation, "Error!")
                txtSourceRange.Text = ""
                txtSourceRange.Focus()
                Exit Sub
            End If

        ElseIf txtSourceRange.Text <> "" And txtDestRange.Text <> "" Then
            If IsValidRng(txtSourceRange.Text.ToUpper) = False And IsValidRng(txtDestRange.Text.ToUpper) = True Then
                MsgBox("Please use a valid range in the Source Range.", MsgBoxStyle.Exclamation, "Error!")
                txtSourceRange.Text = ""
                txtSourceRange.Focus()
                Exit Sub

            ElseIf IsValidRng(txtSourceRange.Text.ToUpper) = True And IsValidRng(txtDestRange.Text.ToUpper) = False Then
                MsgBox("Please use a valid range in the Destination Range.", MsgBoxStyle.Exclamation, "Error!")
                txtDestRange.Text = ""
                txtDestRange.Focus()
                Exit Sub
            ElseIf IsValidRng(txtSourceRange.Text.ToUpper) = False And IsValidRng(txtDestRange.Text.ToUpper) = False Then
                MsgBox("Please use valid ranges in the Source Range and in the Destination Range.", MsgBoxStyle.Exclamation, "Error!")
                txtSourceRange.Text = ""
                txtDestRange.Text = ""
                txtSourceRange.Focus()
                Exit Sub

            End If
        End If


        Dim i, j, count, pasteValue, pasteValue2, count2, lastRowNum, lastColNum As Integer
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



            While destRange.Offset(count, 0).Value IsNot Nothing

                If destRange.Offset(count, count2).EntireRow.Hidden = False Then
                    pasteValue = pasteValue + 1
                    count2 = 0
                    pasteValue2 = 0

                End If
                If pasteValue > sourceRange.Rows.Count Then
                    Exit While
                End If

                While destRange.Offset(count, count2).Value <> Nothing
                    If pasteValue2 + 1 > sourceRange.Columns.Count Then
                        Exit While
                    End If


                    If destRange.Offset(count, count2).EntireRow.Hidden = False And destRange.Offset(count, count2).EntireColumn.Hidden = False Then
                        pasteValue2 = pasteValue2 + 1


                        If CB_keepFormat.Checked = True Then

                            Call copyCell(destRange, count, count2, worksheet.Range(txtSourceRange.Text).Cells(1, 1), pasteValue - 1, pasteValue2 - 1)


                        Else
                            destRange.Offset(count, count2).Value = sourceRange.Cells(1, 1).offset(pasteValue - 1, pasteValue2 - 1).value


                        End If

                    End If

                    count2 = count2 + 1

                End While

                count = count + 1

            End While



        Else

            For j = destRange.Row To lastRowNum

                While destRange.Offset(count, 0).Value <> Nothing

                    If destRange.Offset(count, count2).EntireRow.Hidden = False Then
                        pasteValue = pasteValue + 1
                        count2 = 0
                        pasteValue2 = 0

                    End If
                    If pasteValue > sourceRange.Rows.Count Then
                        Exit While
                    End If

                    While destRange.Offset(count, count2).Value <> Nothing
                        If pasteValue2 + 1 > sourceRange.Columns.Count Then
                            Exit While
                        End If


                        If destRange.Offset(count, count2).EntireRow.Hidden = False And destRange.Offset(count, count2).EntireColumn.Hidden = False Then
                            pasteValue2 = pasteValue2 + 1
                            If CB_keepFormat.Checked = True Then

                                Call copyCell(destRange, count, count2, worksheet.Range(txtSourceRange.Text).Cells(1, 1), pasteValue - 1, pasteValue2 - 1)
                            Else
                                destRange.Offset(count, count2).Value = worksheet.Range(txtSourceRange.Text).Cells(1, 1).offset(pasteValue - 1, pasteValue2 - 1).value

                            End If
                        End If

                        count2 = count2 + 1

                    End While

                    count = count + 1

                End While

            Next





            Dim count3, count4, count5, l As Integer

            count3 = 0

            For k = lastRowNum To lastRowNum + sourceRange.Rows.Count - visibleRows - 1
                count4 = 0
                count5 = 0
                For l = 1 To lastColNum + sourceRange.Columns.Count - visibleCols - 1

                    If worksheet.Cells(lastRowNum, destRange.Column).Offset(count3, l - 1).EntireColumn.Hidden = False Then
                        count5 = count5 + 1
                    End If


                    If count5 > sourceRange.Columns.Count Then
                        Exit For
                    End If

                    If worksheet.Cells(lastRowNum, destRange.Column).Offset(count3, l - 1).EntireColumn.Hidden = False Then

                        If CB_keepFormat.Checked = True Then

                            Call copyCell(worksheet.Cells(lastRowNum, destRange.Column), count3, l - 1, worksheet.Range(txtSourceRange.Text).Cells(1, 1), visibleRows + count3, count4)
                        Else
                            worksheet.Cells(lastRowNum, destRange.Column).Offset(count3, l - 1).Value = worksheet.Range(txtSourceRange.Text).Cells(1, 1).offset(visibleRows + count3, count4).value

                        End If

                        count4 = count4 + 1
                    End If

                Next
                count3 = count3 + 1
            Next



            Dim rowNum, colNum As Integer
            rowNum = destRange.Row
            colNum = destRange.Column
            count3 = 0
            count4 = visibleCols
            For k = destRange.Row To lastRowNum - 1

                If worksheet.Range(worksheet.Cells(k, 1), worksheet.Cells(k, 2)).EntireRow.Hidden = False Then

                    rowNum = worksheet.Range(worksheet.Cells(k, 1), worksheet.Cells(k, 2)).Row

                End If

                If count3 + 1 > sourceRange.Rows.Count Then
                    Exit For
                End If


                If Not worksheet.Range(worksheet.Cells(k, 1), worksheet.Cells(k, 2)).EntireRow.Hidden = False And worksheet.Range(worksheet.Cells(k, 1), worksheet.Cells(k + 1, 1)).EntireColumn.Hidden = False Then

                    GoTo exitLoop

                End If

                count4 = visibleCols


                For l = lastColNum To lastColNum + sourceRange.Columns.Count - visibleCols - 1


                    If worksheet.Range(worksheet.Cells(k, l), worksheet.Cells(k + 1, l)).EntireColumn.Hidden = False Then

                        colNum = worksheet.Range(worksheet.Cells(k, l), worksheet.Cells(k + 1, l)).Column

                    End If
                    If count4 + 1 > sourceRange.Columns.Count Then
                        Exit For
                    End If


                    If worksheet.Range(worksheet.Cells(k, l), worksheet.Cells(k, l + 1)).EntireRow.Hidden = False And worksheet.Range(worksheet.Cells(k, l), worksheet.Cells(k + 1, l)).EntireColumn.Hidden = False Then

                        If CB_keepFormat.Checked = True Then

                            Call copyCell(worksheet.Range(worksheet.Cells(rowNum, colNum), worksheet.Cells(rowNum, colNum)), 0, 0, worksheet.Range(txtSourceRange.Text).Cells(1, 1), count3, count4)
                        Else
                            worksheet.Range(worksheet.Cells(rowNum, colNum), worksheet.Cells(rowNum, colNum)).Offset(0, 0).Value = worksheet.Range(txtSourceRange.Text).Cells(1, 1).offset(count3, count4).value

                        End If




                        'worksheet.Range(worksheet.Cells(rowNum, colNum), worksheet.Cells(rowNum, colNum)).Value = sourceRange.Cells.Offset(count3, count4).Value
                        'sourceRange.Cells.Offset(count3, count4).Copy(worksheet.Cells(rowNum, colNum))

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

    Private Sub Form16PasteintoVisibleRange_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        form_flag = False
    End Sub

    Private Sub Form16PasteintoVisibleRange_Disposed(sender As Object, e As EventArgs) Handles Me.Disposed
        form_flag = False
    End Sub

    Private Sub Form16PasteintoVisibleRange_Shown(sender As Object, e As EventArgs) Handles Me.Shown

        Me.Focus()
        Me.BringToFront()
        Me.Activate()
        Me.BeginInvoke(New System.Action(Sub()
                                             txtSourceRange.Text = sourceRange.Address
                                             SetWindowPos(Me.Handle, New IntPtr(HWND_TOPMOST), 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOSIZE)
                                         End Sub))
    End Sub
End Class