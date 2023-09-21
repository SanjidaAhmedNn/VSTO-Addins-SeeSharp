Imports System.ComponentModel.Design
Imports System.Security.Cryptography
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Excel
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Drawing
Imports System.ComponentModel
Imports System.Linq.Expressions
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Button

Public Class Form17DivideNames

    Dim WithEvents excelApp As Excel.Application
    Dim workbook As Excel.Workbook
    Dim worksheet, worksheet1 As Excel.Worksheet
    Dim outWorksheet As Excel.Worksheet
    Dim inputRng As Excel.Range
    Dim FocusedTxtBox As Integer
    Dim sourceRange, destRange As Excel.Range
    Dim selectedRange As Excel.Range
    Dim changeState As Boolean = False
    Dim textChanged As Boolean = False

    Private Sub Form17DivideNames_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Try

            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            worksheet = workbook.ActiveSheet

            Dim selectedRng As Excel.Range = excelApp.Selection
            txtSourceRange.Text = selectedRng.Address
            RB_Same_As_Source_Range.Checked = True

            Me.KeyPreview = True

        Catch ex As Exception

        End Try

    End Sub


    Private Sub txtSourceRange_TextChanged(sender As Object, e As EventArgs) Handles txtSourceRange.TextChanged

        Try

            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            worksheet = workbook.ActiveSheet


            textChanged = True
            sourceRange = worksheet.Range(txtSourceRange.Text)


            sourceRange.Select()




            If changeState = True Then


                If destRange.Worksheet.Name <> sourceRange.Worksheet.Name Then

                    txtDestRange.Text = destRange.Worksheet.Name & "!" & destRange.Address

                End If


            End If



        Catch ex As Exception

        End Try

        textChanged = False

        txtSourceRange.Focus()


    End Sub

    Private Sub txtDestRange_TextChanged(sender As Object, e As EventArgs) Handles txtDestRange.TextChanged

        Try
            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            worksheet = workbook.ActiveSheet

            changeState = True

            textChanged = True
            destRange = worksheet.Range(txtDestRange.Text)




            destRange.Select()


            If destRange.Worksheet.Name <> sourceRange.Worksheet.Name Then

                txtDestRange.Text = destRange.Worksheet.Name & "!" & destRange.Address

            End If


        Catch ex As Exception

        End Try

        textChanged = False
        txtDestRange.Focus()


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



            txtSourceRange.Text = sourceRange.Worksheet.Name & "!" & sourceRange.Address

            sourceRange.Select()

            txtSourceRange.Focus()



        Catch ex As Exception

            txtSourceRange.Focus()

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


            If textChanged = False Then


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


    Private Sub btnOK_Click(sender As Object, e As EventArgs) Handles btnOK.Click

        Try

            'excelApp = Globals.ThisAddIn.Application
            'workbook = excelApp.ActiveWorkbook
            'worksheet = workbook.ActiveSheet
            'selectedRange = excelApp.Selection
            'Dim arrRng As String()
            'Dim arrName As String()
            'Dim arrTitle() = {"Sir", "Miss", "Lady", "Lord", "Madam", "Master"}
            'Dim arrHeader() = {"Full Name", "Title", "First Name", "Middle Name", "Last Name Prefix", "Last Name", "Abbreviations", "Name Suffix"}

            'If RB_Same_As_Source_Range.Checked = True Then






            'ElseIf RB_Different_Range.Checked = True Then

            '    If CB_Keep_Formatting.Checked = True Then
            '        If CB_Add_Header.Checked = True Then
            '            For i = 1 To 8
            '                destRange.Cells(1, i).value = arrHeader(i - 1)
            '            Next
            '        End If
            '    End If

            '    Exit Sub



            '    'arrRng = Split(txtSourceRange.Text, ",")

            '    'For i = 0 To UBound(arrRng)

            '    '    selectedRange = worksheet.Range(arrRng(i))

            '    '    For j = 1 To selectedRange.Rows.Count

            '    '        arrName = Split(selectedRange.Cells(j, 1).value, " ")

            '    '        Dim dotCount As Integer
            '    '        dotCount = 0
            '    '        For Each c As Char In arrName(0)

            '    '            If c = "." Then
            '    '                dotCount += 1
            '    '            End If

            '    '        Next
            '    '        If dotCount > 0 Or arrTitle.Contains(arrName(0), StringComparer.OrdinalIgnoreCase) Then
            '    '            MsgBox("Title")






            '    '        Else
            '    '            MsgBox("no title")
            '    '        End If


            '    '        dotCount = 0
            '    '        For Each c As Char In arrName(UBound(arrName))

            '    '            If c = "." Then
            '    '                dotCount += 1
            '    '            End If

            '    '        Next


            '    '        If dotCount > 0 Then
            '    '            MsgBox("suffix")
            '    '        Else
            '    '            MsgBox("No suffix")
            '    '        End If


            '    '    Next


            '    'Next



            'End If


            Call nameSplitter()




        Catch ex As Exception

        End Try




    End Sub


    Private Sub nameSplitter()
        excelApp = Globals.ThisAddIn.Application
        workbook = excelApp.ActiveWorkbook
        worksheet = workbook.ActiveSheet
        selectedRange = excelApp.Selection
        Dim arrRng As String()
        Dim arrName As String()
        Dim arrTitle() = {"Sir", "Miss", "Lady", "Lord", "Madam", "Master"}
        Dim arrHeader() = {"Full Name", "Title", "First Name", "Middle Name", "Last Name Prefix", "Last Name", "Abbreviations", "Name Suffix"}
        'Dim arrSplitName As String()

        arrRng = Split(txtSourceRange.Text, ",")

        For i = 0 To UBound(arrRng)

            selectedRange = worksheet.Range(arrRng(i))

            For j = 1 To selectedRange.Rows.Count

                arrName = Split(selectedRange.Cells(j, 1).value, " ")

                Dim dotCount As Integer
                dotCount = 0
                For Each c As Char In arrName(0)

                    If c = "." Then
                        dotCount += 1
                    End If

                Next
                If dotCount > 0 Or arrTitle.Contains(arrName(0), StringComparer.OrdinalIgnoreCase) Then
                    MsgBox("Title")
                    For p = 1 To destRange.Rows.Count
                        For q = 1 To 8
                            destRange.Cells(p, 1).value = arrName(0)
                        Next
                    Next




                Else
                    MsgBox("no title")
                End If


                'dotCount = 0
                'For Each c As Char In arrName(UBound(arrName))

                '    If c = "." Then
                '        dotCount += 1
                '    End If

                'Next


                'If dotCount > 0 Then
                '    MsgBox("suffix")
                'Else
                '    MsgBox("No suffix")
                'End If


            Next


        Next


    End Sub

End Class