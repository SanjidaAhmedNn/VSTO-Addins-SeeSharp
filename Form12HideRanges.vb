Imports System.ComponentModel.Design
Imports System.Security.Cryptography
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Excel
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Drawing
Imports System.ComponentModel
Imports System.Linq.Expressions


Public Class Form12HideRanges
    Dim WithEvents excelApp As Excel.Application
    Dim workbook As Excel.Workbook
    Dim worksheet, worksheet1 As Excel.Worksheet
    Dim outWorksheet As Excel.Worksheet
    Dim inputRng As Excel.Range
    Dim FocusedTxtBox As Integer
    Dim selectedRange As Excel.Range
    Dim textChanged As Boolean = False

    Private Sub Form1_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Enter Then
            btn_OK.PerformClick()
        End If
    End Sub

    Private Sub Form12HideRanges_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        excelApp = Globals.ThisAddIn.Application
        workbook = excelApp.ActiveWorkbook
        worksheet = workbook.ActiveSheet

        Dim selectedRng As Excel.Range = excelApp.Selection
        txtSourceRange.Text = selectedRng.Address

        RB_Single_Range.Checked = True
        RB_Row.Checked = True

        Me.KeyPreview = True
    End Sub
    Private Sub txtSourceRange_TextChanged(sender As Object, e As EventArgs) Handles txtSourceRange.TextChanged

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

    Private Sub AutoSelection_Click(sender As Object, e As EventArgs) Handles AutoSelection.Click

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




            '    Try

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

            '    Catch ex As Exception

            '    End Try



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
            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            worksheet = workbook.ActiveSheet
            inputWsName = worksheet.Name

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



            Dim rngCount As Integer
            rngCount = 0
            For Each c As Char In txtSourceRange.Text

                If c = "," Then
                    rngCount = rngCount + 1
                End If

            Next


            If rngCount = 0 And RB_Single_Range.Checked = True Then
                Call singleRng()
                Me.Dispose()
            ElseIf rngCount = 0 And RB_Multiple_Range.Checked = True Then
                MsgBox("Please select the correct Range Type.", MsgBoxStyle.Exclamation, "Error!")
                RB_Single_Range.Focus()
            ElseIf rngCount <> 0 And RB_Multiple_Range.Checked = True Then
                Call multiRng()
                Me.Dispose()
            ElseIf rngCount <> 0 And RB_Single_Range.Checked = True Then
                MsgBox("Please select the correct Range Type.", MsgBoxStyle.Exclamation, "Error!")
                RB_Multiple_Range.Focus()
            End If




        Catch ex As Exception

        End Try
    End Sub



    Private Sub singleRng()

        Try
            Dim inputWsName As String
            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            worksheet = workbook.ActiveSheet
            inputWsName = worksheet.Name

            Dim temp As String
            temp = txtSourceRange.Text
            worksheet1 = inputRng.Worksheet

            If CheckBox1.Checked = True Then

                workbook.ActiveSheet.Copy(After:=workbook.Sheets(workbook.Sheets.Count))
                outWorksheet = workbook.Sheets(workbook.Sheets.Count)

                worksheet1.Activate()
                txtSourceRange.Text = temp

            End If

            Dim firstRow, lastRow, firstColumn, lastColumn As Integer


            selectedRange = worksheet.Range(txtSourceRange.Text)
            firstRow = selectedRange.Row
            lastRow = firstRow + selectedRange.Rows.Count - 1
            firstColumn = selectedRange.Column
            lastColumn = firstColumn + selectedRange.Columns.Count - 1

            If RB_Single_Range.Checked = True And RB_Row.Checked = True Then
                If selectedRange.Rows.Count <= 2 Then
                    Dim answer As MsgBoxResult
                    answer = MsgBox("You are about to hide " & selectedRange.Rows.Count & " Rows." & vbCrLf & "Do you want to proceed?", MsgBoxStyle.YesNo, "Warning!")
                    If answer = MsgBoxResult.Yes Then
                        GoTo Proceed1
                    Else
                        GoTo break1
                    End If
                End If
Proceed1:
                worksheet.Range(worksheet.Cells(firstRow, firstColumn), worksheet.Cells(lastRow, lastColumn)).EntireRow.Hidden = True
break1:
                Me.Dispose()

            ElseIf RB_Single_Range.Checked = True And RB_Column.Checked = True Then
                If selectedRange.Columns.Count <= 2 Then
                    Dim answer As MsgBoxResult
                    answer = MsgBox("You are about to hide " & selectedRange.Columns.Count & " Columns." & vbCrLf & "Do you want to proceed?", MsgBoxStyle.YesNo, "Warning!")
                    If answer = MsgBoxResult.Yes Then
                        GoTo Proceed2
                    Else
                        GoTo break2
                    End If
                End If
Proceed2:
                worksheet.Range(worksheet.Cells(firstRow, firstColumn), worksheet.Cells(lastRow, lastColumn)).EntireColumn.Hidden = True
break2:
                Me.Dispose()

            ElseIf RB_Single_Range.Checked = True And RB_bidirection.Checked = True Then
                If selectedRange.Columns.Count <= 2 Then
                    Dim answer As MsgBoxResult
                    answer = MsgBox("You are about to hide " & selectedRange.Rows.Count & " Rows and" & selectedRange.Columns.Count & " Columns." & vbCrLf & "Do you want to proceed?", MsgBoxStyle.YesNo, "Warning!")
                    If answer = MsgBoxResult.Yes Then
                        GoTo Proceed3
                    Else
                        GoTo break3
                    End If
                End If
Proceed3:
                worksheet.Range(worksheet.Cells(firstRow, 1), worksheet.Cells(lastRow, 1)).EntireRow.Hidden = True
                worksheet.Range(worksheet.Cells(1, firstColumn), worksheet.Cells(1, lastColumn)).EntireColumn.Hidden = True

break3:
                Me.Dispose()
            End If



        Catch ex As Exception

        End Try


    End Sub

    Private Sub multiRng()

        Try

            Dim inputWsName As String
            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            worksheet = workbook.ActiveSheet
            inputWsName = worksheet.Name

            Dim temp As String
            temp = txtSourceRange.Text
            worksheet1 = inputRng.Worksheet


            If CheckBox1.Checked = True Then

                workbook.ActiveSheet.Copy(After:=workbook.Sheets(workbook.Sheets.Count))
                outWorksheet = workbook.Sheets(workbook.Sheets.Count)

                worksheet1.Activate()
                txtSourceRange.Text = temp

            End If


            Dim visRows, followingRows As Integer
            Dim visColumns, followingColumns As Integer
            Dim arrRng As String() = Split(txtSourceRange.Text, ",")

            If RB_Multiple_Range.Checked = True And RB_Row.Checked = True Then
                For i = 0 To UBound(arrRng)
                    visRows = worksheet.Range(arrRng(i)).Row
                    followingRows = visRows + worksheet.Range(arrRng(i)).Rows.Count - 1
                    visColumns = worksheet.Range(arrRng(i)).Column
                    followingColumns = visColumns + worksheet.Range(arrRng(i)).Columns.Count - 1

                    worksheet.Range(worksheet.Cells(visRows, 1), worksheet.Cells(followingRows, 1)).EntireRow.Hidden = True

                Next



            ElseIf RB_Multiple_Range.Checked = True And RB_Column.Checked = True Then
                For i = 0 To UBound(arrRng)
                    visRows = worksheet.Range(arrRng(i)).Row
                    followingRows = visRows + worksheet.Range(arrRng(i)).Rows.Count - 1
                    visColumns = worksheet.Range(arrRng(i)).Column
                    followingColumns = visColumns + worksheet.Range(arrRng(i)).Columns.Count - 1

                    worksheet.Range(worksheet.Cells(1, visColumns), worksheet.Cells(1, followingColumns)).EntireColumn.Hidden = True


                Next


            Else
                For i = 0 To UBound(arrRng)
                    visRows = worksheet.Range(arrRng(i)).Row
                    followingRows = visRows + worksheet.Range(arrRng(i)).Rows.Count - 1
                    visColumns = worksheet.Range(arrRng(i)).Column
                    followingColumns = visColumns + worksheet.Range(arrRng(i)).Columns.Count - 1

                    worksheet.Range(worksheet.Cells(visRows, 1), worksheet.Cells(followingRows, 1)).EntireRow.Hidden = True
                    worksheet.Range(worksheet.Cells(1, visColumns), worksheet.Cells(1, followingColumns)).EntireColumn.Hidden = True


                Next


            End If






        Catch ex As Exception

        End Try
    End Sub

    Private Sub btn_Cancel_Click(sender As Object, e As EventArgs) Handles btn_Cancel.Click
        Me.Dispose()
    End Sub
End Class