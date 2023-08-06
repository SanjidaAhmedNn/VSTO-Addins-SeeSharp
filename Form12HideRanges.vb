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
    Dim worksheet As Excel.Worksheet
    Dim outWorksheet As Excel.Worksheet
    Dim inputRng As Excel.Range
    Dim FocusedTxtBox As Integer
    Dim selectedRange As Excel.Range

    Private Sub txtSourceRange_TextChanged(sender As Object, e As EventArgs) Handles txtSourceRange.TextChanged

        Try
            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            worksheet = workbook.ActiveSheet


            txtSourceRange.Focus()


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
            txtSourceRange.Focus()

            inputRng = excelApp.InputBox("Please Select a Range", "Range Selection", selectedRange.Address, Type:=8)
            inputRng.Select()
            txtSourceRange.Text = inputRng.Address
            txtSourceRange.Focus()



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


            If FocusedTxtBox = 1 Then

                txtSourceRange.Text = selectedRange.Address
                worksheet = workbook.ActiveSheet
                inputRng = selectedRange
                txtSourceRange.Focus()

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

    Private Sub btn_OK_Click(sender As Object, e As EventArgs) Handles btn_OK.Click
        Try

            Dim inputWsName As String
            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            worksheet = workbook.ActiveSheet
            inputWsName = worksheet.Name





            Dim rngCount As Integer
            rngCount = 0
            For Each c As Char In txtSourceRange.Text

                If c = "," Then
                    rngCount = rngCount + 1
                End If

            Next


            If rngCount = 0 Then

                Call singleRng()
            Else
                Call multiRng()
            End If
break:

            Me.Dispose()


        Catch ex As Exception

        End Try
    End Sub

    Private Sub Form12HideRanges_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        excelApp = Globals.ThisAddIn.Application
        workbook = excelApp.ActiveWorkbook
        worksheet = workbook.ActiveSheet

        Dim selectedRng As Excel.Range = excelApp.Selection
        txtSourceRange.Text = selectedRng.Address


    End Sub

    Private Sub singleRng()

        Try
            Dim inputWsName As String
            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            worksheet = workbook.ActiveSheet
            inputWsName = worksheet.Name


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


            If CheckBox1.Checked = True Then

                workbook.ActiveSheet.Copy(After:=workbook.Sheets(workbook.Sheets.Count))
                outWorksheet = workbook.Sheets(workbook.Sheets.Count)
                outWorksheet.Range("A1").Select()
                worksheet.Cells.EntireColumn.Hidden = False
                worksheet.Cells.EntireRow.Hidden = False

                worksheet = workbook.Sheets(inputWsName)
                worksheet.Activate()

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


            If CheckBox1.Checked = True Then

                workbook.ActiveSheet.Copy(After:=workbook.Sheets(workbook.Sheets.Count))
                outWorksheet = workbook.Sheets(workbook.Sheets.Count)
                outWorksheet.Range("A1").Select()
                worksheet.Cells.EntireColumn.Hidden = False
                worksheet.Cells.EntireRow.Hidden = False

                worksheet = workbook.Sheets(inputWsName)
                worksheet.Activate()

            End If




        Catch ex As Exception

        End Try
    End Sub

    Private Sub btn_Cancel_Click(sender As Object, e As EventArgs) Handles btn_Cancel.Click
        Me.Dispose()
    End Sub
End Class