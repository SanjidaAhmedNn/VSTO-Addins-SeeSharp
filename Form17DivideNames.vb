Imports System.ComponentModel.Design
Imports System.Security.Cryptography
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Excel
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Drawing
Imports System.ComponentModel
Imports System.Linq.Expressions
Imports Microsoft.VisualBasic
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Button
Imports System.Text.RegularExpressions

Public Class Form17DivideNames

    Dim WithEvents excelApp As Excel.Application
    Dim workbook As Excel.Workbook
    Dim worksheet, worksheet1 As Excel.Worksheet
    Dim outWorksheet As Excel.Worksheet
    Dim inputRng As Excel.Range
    Dim FocusedTxtBox As Integer
    Dim sourceRange, destRange As Excel.Range
    Dim selectedRange As Excel.Range
    Dim mainArr(6) As String
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
        Call display()

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

    Public Function IsValidRng(input As String) As Boolean
        '"^(([A-Za-z]+[0-9]*( \([0-9]+\))?!)?\$?[A-Z]+\$?[0-9]+(:\$?[A-Z]+\$?[0-9]+)?)(,([A-Za-z]+[0-9]*( \([0-9]+\))?!)?\$?[A-Z]+\$?[0-9]+(:\$?[A-Z]+\$?[0-9]+)?)*$"     
        Dim pattern As String = "^(\$?[A-Z]+\$?[0-9]+(:\$?[A-Z]+\$?[0-9]+)?)(,\$?[A-Z]+\$?[0-9]+(:\$?[A-Z]+\$?[0-9]+)?)*$"
        Return System.Text.RegularExpressions.Regex.IsMatch(input, pattern)

    End Function

    Private Sub btnOK_Click(sender As Object, e As EventArgs) Handles btnOK.Click

        Try
            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            worksheet = workbook.ActiveSheet
            selectedRange = excelApp.Selection

            Dim checkBox_checked_count As Integer = 0
            For Each ctrl As Control In CustomGroupBox7.Controls
                If TypeOf ctrl Is System.Windows.Forms.CheckBox Then
                    Dim chk As System.Windows.Forms.CheckBox = DirectCast(ctrl, System.Windows.Forms.CheckBox)
                    If chk.Checked Then
                        checkBox_checked_count += 1
                    End If
                End If
            Next


            If RB_Same_As_Source_Range.Checked = True Then

                If txtSourceRange.Text = "" Then
                    MsgBox("Please select the Source Range.", MsgBoxStyle.Exclamation, "Error!")
                    txtSourceRange.Focus()
                    Exit Sub
                Else
                    If IsValidRng(txtSourceRange.Text.ToUpper) = False Then
                        MsgBox("Please use a valid range in the Source Range.", MsgBoxStyle.Exclamation, "Error!")
                        txtDestRange.Focus()
                        Exit Sub


                    ElseIf checkBox_checked_count = 0 Then
                        MsgBox("Please check least one checkbox to divide names.", MsgBoxStyle.Exclamation, "Error!")

                        CustomGroupBox7.Focus()
                        Exit Sub

                    End If

                End If

            ElseIf RB_Different_Range.Checked = True Then

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
                    If checkBox_checked_count = 0 Then
                        MsgBox("Please check least one checkbox to divide names.", MsgBoxStyle.Exclamation, "Error!")
                        CustomGroupBox7.Focus()
                        Exit Sub

                    ElseIf IsValidRng(txtSourceRange.Text.ToUpper) = False And IsValidRng(txtDestRange.Text.ToUpper) = True Then
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

            End If


            Dim arrRng As String()
            Dim temp As String
            temp = txtSourceRange.Text
            worksheet1 = sourceRange.Worksheet


            If CB_Backup_Sheet.Checked = True Then

                workbook.ActiveSheet.Copy(After:=workbook.Sheets(workbook.Sheets.Count))
                outWorksheet = workbook.Sheets(workbook.Sheets.Count)


                worksheet1.Activate()
                txtSourceRange.Text = temp

            End If



            If RB_Same_As_Source_Range.Checked = True Then

                Dim outputColumn As Integer

                arrRng = Split(txtSourceRange.Text, ",")

                For i = 0 To UBound(arrRng)

                    sourceRange = worksheet.Range(arrRng(i))

                    For j = 1 To sourceRange.Rows.Count

                        mainArr = {"", "", "", "", "", "", "", ""}
                        name = sourceRange.Cells(j, 1).value

                        Call nameSplitter()

                        Dim headerIndex As Integer
                        Dim headerStr As String = ""
                        For Each ctrl As Control In CustomGroupBox7.Controls
                            If TypeOf ctrl Is System.Windows.Forms.CheckBox Then
                                Dim chk As System.Windows.Forms.CheckBox = DirectCast(ctrl, System.Windows.Forms.CheckBox)
                                If chk.Checked Then
                                    headerStr = headerStr & "," & chk.Text
                                End If
                            End If
                        Next

                        headerStr = headerStr.Replace("Select All,", String.Empty)
                        headerStr = Microsoft.VisualBasic.Right(headerStr, Len(headerStr) - 1)

                        Dim arrHeaderStr As String() = Split(headerStr, ",")
                        outputColumn = UBound(arrHeaderStr) + 1
                        For k = 0 To UBound(arrHeaderStr)
                            worksheet.Cells(100000, k + 1).value = arrHeaderStr(k)

                            Select Case arrHeaderStr(k)
                                Case = "Title"
                                    headerIndex = 0
                                Case = "First Name"
                                    headerIndex = 1
                                Case = "Middle Name"
                                    headerIndex = 2
                                Case = "Last Name Prefix"
                                    headerIndex = 3
                                Case = "Last Name"
                                    headerIndex = 4
                                Case = "Name Suffix"
                                    headerIndex = 5
                                Case = "Name Abbreviations"
                                    headerIndex = 6

                            End Select
                            If CB_Keep_Formatting.Checked = True Then
                                sourceRange.Cells(j, 1).copy(worksheet.Cells(100000 + j, k + 1))
                            End If
                            worksheet.Cells(100000 + j, k + 1).value = mainArr(headerIndex)
                        Next





                    Next

                    worksheet.Range(worksheet.Cells(100000, 1), worksheet.Cells(100000, outputColumn)).Font.Bold = True
                    worksheet.Range(worksheet.Cells(100000, 1), worksheet.Cells(100000 + sourceRange.Rows.Count, outputColumn)).Copy(sourceRange.Cells(1, 1))

                    worksheet.Range(worksheet.Cells(100000, 1), worksheet.Cells(100000 + sourceRange.Rows.Count, outputColumn)).Delete(Excel.XlDeleteShiftDirection.xlShiftUp)
                    worksheet.Range(sourceRange.Cells(1, 1), sourceRange.Cells(sourceRange.Rows.Count + 1, outputColumn)).Select()

                    Dim border As Excel.Borders = selectedRange.Borders
                    border.LineStyle = Excel.XlLineStyle.xlContinuous
                    border.Weight = Excel.XlBorderWeight.xlThin

                    selectedRange.EntireColumn.AutoFit()

                    If CB_Add_Header.Checked = False Then

                        worksheet.Range(selectedRange.Cells(1, 1), selectedRange.Cells(1, outputColumn)).Delete(Excel.XlDeleteShiftDirection.xlShiftUp)
                        worksheet.Range(sourceRange.Cells(1, 1), sourceRange.Cells(sourceRange.Rows.Count + 1, outputColumn)).Select()

                    End If

                Next



            ElseIf RB_Different_Range.Checked = True Then


                Dim outputColumn As Integer

                arrRng = Split(txtSourceRange.Text, ",")

                For i = 0 To UBound(arrRng)

                    sourceRange = worksheet.Range(arrRng(i))

                    For j = 1 To sourceRange.Rows.Count

                        mainArr = {"", "", "", "", "", "", "", ""}
                        name = sourceRange.Cells(j, 1).value

                        Call nameSplitter()

                        Dim headerIndex As Integer
                        Dim headerStr As String = ""
                        For Each ctrl As Control In CustomGroupBox7.Controls
                            If TypeOf ctrl Is System.Windows.Forms.CheckBox Then
                                Dim chk As System.Windows.Forms.CheckBox = DirectCast(ctrl, System.Windows.Forms.CheckBox)
                                If chk.Checked Then
                                    headerStr = headerStr & "," & chk.Text
                                End If
                            End If
                        Next

                        headerStr = headerStr.Replace("Select All,", String.Empty)
                        headerStr = Microsoft.VisualBasic.Right(headerStr, Len(headerStr) - 1)

                        Dim arrHeaderStr As String() = Split(headerStr, ",")
                        outputColumn = UBound(arrHeaderStr) + 1
                        For k = 0 To UBound(arrHeaderStr)
                            worksheet.Cells(100000, k + 1).value = arrHeaderStr(k)

                            Select Case arrHeaderStr(k)
                                Case = "Title"
                                    headerIndex = 0
                                Case = "First Name"
                                    headerIndex = 1
                                Case = "Middle Name"
                                    headerIndex = 2
                                Case = "Last Name Prefix"
                                    headerIndex = 3
                                Case = "Last Name"
                                    headerIndex = 4
                                Case = "Name Suffix"
                                    headerIndex = 5
                                Case = "Name Abbreviations"
                                    headerIndex = 6

                            End Select
                            If CB_Keep_Formatting.Checked = True Then
                                sourceRange.Cells(j, 1).copy(worksheet.Cells(100000 + j, k + 1))
                            End If
                            worksheet.Cells(100000 + j, k + 1).value = mainArr(headerIndex)
                        Next





                    Next

                    worksheet.Range(worksheet.Cells(100000, 1), worksheet.Cells(100000, outputColumn)).Font.Bold = True
                    worksheet.Range(worksheet.Cells(100000, 1), worksheet.Cells(100000 + sourceRange.Rows.Count, outputColumn)).Copy(destRange.Cells(1, 1))

                    worksheet.Range(worksheet.Cells(100000, 1), worksheet.Cells(100000 + sourceRange.Rows.Count, outputColumn)).Delete(Excel.XlDeleteShiftDirection.xlShiftUp)
                    worksheet.Range(destRange.Cells(1, 1), destRange.Cells(sourceRange.Rows.Count + 1, outputColumn)).Select()

                    Dim border As Excel.Borders = selectedRange.Borders
                    border.LineStyle = Excel.XlLineStyle.xlContinuous
                    border.Weight = Excel.XlBorderWeight.xlThin

                    selectedRange.EntireColumn.AutoFit()

                    If CB_Add_Header.Checked = False Then

                        worksheet.Range(selectedRange.Cells(1, 1), selectedRange.Cells(1, outputColumn)).Delete(Excel.XlDeleteShiftDirection.xlShiftUp)
                        worksheet.Range(destRange.Cells(1, 1), destRange.Cells(sourceRange.Rows.Count + 1, outputColumn)).Select()

                    End If

                Next




            End If


            Me.Dispose()




        Catch ex As Exception

        End Try

        Me.Dispose()



    End Sub




    Function checkTitle(ByVal inputStr As String) As Boolean

        Dim dotCount As Integer

        Dim arrTitle() = {
                                "Mr", "Mister", "Mrs", "Missus", "Miss", "Ms", "Dr", "Doctor",
                                "Prof", "Professor", "Sir", "Lady", "Lord", "Madam",
                                "Mdm", "Count", "Madame", "Master", "Rev", "Reverend", "Fr",
                                "Father", "Sr", "Sister", "Pvt", "Private", "Esq", "Esquire",
                                "Imam", "Sheikh", "Capt", "Captain", "Cpl", "Corporal",
                                "Sgt", "Sergeant", "Gen", "General", "Lt", "Lieutenant",
                                "Eng", "Engineer", "Hon", "Honorable", "Pres", "President",
                                "VP", "Vice President", "Gov", "Governor", "Sen", "Senator",
                                "Rep", "Representative", "Mx", "Herr", "Frau", "Duke",
                                "Señor", "Señora", "Señorita", "Dott", "Dottore", "Mlle", "Mademoiselle",
                                "Maestro", "Don", "Doña", "Smt", "Shrimati", "Shri", "Guru", "Sensei"
                          }

        dotCount = 0

        'count if there is any periods in the first word of the name
        'if a period is there then it will be considered as a title
        For Each c As Char In inputStr

            If c = "." Then
                dotCount += 1
            End If

        Next

        'checks if there are period(s) in the first word
        'OR the first word matches with any of the word from the arrTitle array (case insensitively)
        'if any one of the 2 conditon is true then, assign the first word as value in the first column of the destRange
        'otherwise assign a blank value
        If dotCount > 0 Or arrTitle.Contains(inputStr, StringComparer.OrdinalIgnoreCase) Then
            Return True
        Else
            Return False
        End If

    End Function


    Function checkSuffix(ByVal inputStr As String) As Boolean

        Dim dotCount As Integer

        Dim arrSuffix() = {
                                "Jr", "Sr", "II", "III", "IV", "V",
                                "VI", "VII", "VIII", "IX", "X", "MD",
                                "PhD", "Esq", "DDS", "RN", "CPA",
                                "DVM", "JD", "LLB", "LLM", "BA",
                                "BS", "MA", "MS", "PsyD", "OD",
                                "DO", "EdD", "DPhil", "PE", "CFA",
                                "MBA", "MPH", "BEd", "MFA", "ThD",
                                "DMin", "DPT", "BBA", "MDiv", "RPh",
                                "OBE", "KBE", "DC", "NP", "PA",
                                "CNM", "FACP", "DABR"
                            }

        'Name Suffix


        'count if there is any periods in the last word of the name
        'if a period is there then it will be considered as a Name Suffix
        dotCount = 0
        For Each c As Char In inputStr

            If c = "." Then
                dotCount += 1
            End If

        Next

        'checks if there are period(s) in the last word
        'OR the last word matches with any of the word from the arrSuffix array (case insensitively)
        'if any one of the 2 conditon is true then, assign the last word as value in the last column of the destRange
        'otherwise assign a blank value
        If dotCount > 0 Or arrSuffix.Contains(inputStr, StringComparer.OrdinalIgnoreCase) Then
            Return True
        Else
            Return False
        End If

    End Function




    Sub nameSplitter()

        excelApp = Globals.ThisAddIn.Application
        workbook = excelApp.ActiveWorkbook
        worksheet = workbook.ActiveSheet
        selectedRange = excelApp.Selection
        'Dim arrRng As String()
        Dim arrName As String()

        'Dim mainArr(7) As String


        Dim arrTitle() = {
                        "Mr", "Mister", "Mrs", "Missus", "Miss", "Ms", "Dr", "Doctor",
                        "Prof", "Professor", "Sir", "Lady", "Lord", "Madam",
                        "Mdm", "Count", "Madame", "Master", "Rev", "Reverend", "Fr",
                        "Father", "Sr", "Sister", "Pvt", "Private", "Esq", "Esquire",
                        "Imam", "Sheikh", "Capt", "Captain", "Cpl", "Corporal",
                        "Sgt", "Sergeant", "Gen", "General", "Lt", "Lieutenant",
                        "Eng", "Engineer", "Hon", "Honorable", "Pres", "President",
                        "VP", "Vice President", "Gov", "Governor", "Sen", "Senator",
                        "Rep", "Representative", "Mx", "Herr", "Frau", "Duke",
                        "Señor", "Señora", "Señorita", "Dott", "Dottore", "Mlle", "Mademoiselle",
                        "Maestro", "Don", "Doña", "Smt", "Shrimati", "Shri", "Guru", "Sensei"
                  }


        Dim arrSuffix() = {
                                "Jr", "Sr", "II", "III", "IV", "V",
                                "VI", "VII", "VIII", "IX", "X", "MD",
                                "PhD", "Esq", "DDS", "RN", "CPA",
                                "DVM", "JD", "LLB", "LLM", "BA",
                                "BS", "MA", "MS", "PsyD", "OD",
                                "DO", "EdD", "DPhil", "PE", "CFA",
                                "MBA", "MPH", "BEd", "MFA", "ThD",
                                "DMin", "DPT", "BBA", "MDiv", "RPh",
                                "OBE", "KBE", "DC", "NP", "PA",
                                "CNM", "FACP", "DABR"
                            }


        'Dim arrHeader() = {"Full Name", "Title", "First Name", "Middle Name", "Last Name Prefix", "Last Name", "Abbreviations", "Name Suffix"}
        'Dim arrSplitName As String()


        'arrRng = Split(txtSourceRange.Text, ",")

        'For i = 0 To UBound(arrRng)

        '    selectedRange = worksheet.Range(arrRng(i))

        'For j = 1 To selectedRange.Rows.Count

        'mainArr = {"", "", "", "", "", "", "", ""}
        '    mainArr(0) = selectedRange.Cells(j, 1).value

        'arrName = Split(selectedRange.Cells(j, 1).value, " ")
        arrName = Split(name, " ")



        If UBound(arrName) = 0 Then

            If checkTitle(arrName(0)) = True Then
                mainArr(0) = arrName(0)

            ElseIf checkSuffix(arrName(0)) = True Then
                mainArr(5) = arrName(0)

            Else
                mainArr(1) = arrName(0)

            End If



        ElseIf UBound(arrName) = 1 Then

            If checkTitle(arrName(0)) = True And checkSuffix(arrName(1)) = True Then
                'Dr. PhD

                'add title to 1st place in Mainarray
                mainArr(0) = arrName(0)

                'add suffix to 6th  place in main array
                mainArr(5) = arrName(1)

            ElseIf checkTitle(arrName(0)) = True And checkSuffix(arrName(1)) = False Then
                'Dr. John

                'add title in the 1st place
                mainArr(0) = arrName(0)

                'add first name to the 2nd place 
                mainArr(1) = arrName(1)

            ElseIf checkTitle(arrName(0)) = False And checkSuffix(arrName(1)) = True Then
                'John PhD

                'add first name to the 2nd place
                mainArr(1) = arrName(0)

                'add suffix to the 6th place
                mainArr(6) = arrName(1)

            ElseIf checkTitle(arrName(0)) = False And checkSuffix(arrName(1)) = False Then
                'John Smith

                'add first name to the 2nd place
                mainArr(1) = arrName(0)

                'add last name in the 5th place
                mainArr(4) = arrName(1)

            End If


        ElseIf UBound(arrName) = 2 Then

            If checkTitle(arrName(0)) = True And checkSuffix(arrName(2)) = True Then
                'Dr. John PhD

                'add title to 1st place in Mainarray
                mainArr(0) = arrName(0)

                'add first name to the 2nd plcae in the main array
                mainArr(1) = arrName(1)

                'add suffix to 6th place in main array
                mainArr(5) = arrName(2)


            ElseIf checkTitle(arrName(0)) = True And checkSuffix(arrName(2)) = False Then
                'Dr. John Smith

                'add title to the 1st place
                mainArr(0) = arrName(0)

                'add frist name to the 2nd place
                mainArr(1) = arrName(1)

                'add last name to the 5th place
                mainArr(4) = arrName(2)

            ElseIf checkTitle(arrName(0)) = False And checkSuffix(arrName(2)) = True Then
                'John Smith PhD

                'add first name to the 2nd place
                mainArr(1) = arrName(0)

                'add last name to the 5th place
                mainArr(4) = arrName(1)

                'add suffix to the 6th place
                mainArr(5) = arrName(2)

            ElseIf checkTitle(arrName(0)) = False And checkSuffix(arrName(2)) = False Then
                'John Phillip Smith

                'add first name to the 2nd place
                mainArr(1) = arrName(0)

                'add middle name to the 3rd place
                mainArr(2) = arrName(1)

                'add last name to the 5th plcae
                mainArr(4) = arrName(2)

                'add abbreviation to the 7th place
                mainArr(6) = Microsoft.VisualBasic.Left(mainArr(1), 1) & "." & Microsoft.VisualBasic.Left(mainArr(2), 1) & ". " & mainArr(4)


            End If


        ElseIf UBound(arrName) = 3 Then


            If checkTitle(arrName(0)) = True And checkSuffix(arrName(3)) = True Then
                'Dr. John Smith PhD

                'add title to 1st place in Main array
                mainArr(0) = arrName(0)

                'add first name to the 2nd plcae in the main array
                mainArr(1) = arrName(1)

                'add last name to the 5th place
                mainArr(4) = arrName(2)

                'add suffix to 6th place in main array
                mainArr(5) = arrName(3)


            ElseIf checkTitle(arrName(0)) = True And checkSuffix(arrName(3)) = False Then
                'Dr. John Phillip Smith

                'add title to the 1st place
                mainArr(0) = arrName(0)

                'add frist name to the 2nd place
                mainArr(1) = arrName(1)

                'add middle name to the 3rd place
                mainArr(2) = arrName(2)

                'add last name to the 5th place
                mainArr(4) = arrName(3)

                'add abbreviation to the 7th place
                mainArr(6) = Microsoft.VisualBasic.Left(mainArr(1), 1) & "." & Microsoft.VisualBasic.Left(mainArr(2), 1) & ". " & mainArr(4)


            ElseIf checkTitle(arrName(0)) = False And checkSuffix(arrName(3)) = True Then
                'John Phillip Smith PhD

                'add first name to the 2nd place
                mainArr(1) = arrName(0)

                'add middle name to the 3rd place
                mainArr(2) = arrName(1)

                'add last name to the 5th place
                mainArr(4) = arrName(2)

                'add suffix to the 6th place
                mainArr(5) = arrName(3)

                'add abbreviation to the 7th place
                mainArr(6) = Microsoft.VisualBasic.Left(mainArr(1), 1) & "." & Microsoft.VisualBasic.Left(mainArr(2), 1) & ". " & mainArr(4)

            ElseIf checkTitle(arrName(0)) = False And checkSuffix(arrName(3)) = False Then
                'John Phillip Van Smith

                'add first name to the 2nd place
                mainArr(1) = arrName(0)

                'add middle name to the 3rd place
                mainArr(2) = arrName(1)

                'add last name prefix in 4th place
                mainArr(3) = arrName(2)

                'add last name to the 5th plcae
                mainArr(4) = arrName(3)

                'add abbreviation to the 7th place
                mainArr(6) = Microsoft.VisualBasic.Left(mainArr(1), 1) & "." & Microsoft.VisualBasic.Left(mainArr(2), 1) & ". " & mainArr(3) & " " & mainArr(4)


            End If

        ElseIf UBound(arrName) = 4 Then

            If checkTitle(arrName(0)) = True And checkSuffix(arrName(4)) = True Then
                'Dr. John Phillip Smith PhD

                'add title to 1st place in Main array
                mainArr(0) = arrName(0)

                'add first name to the 2nd plcae in the main array
                mainArr(1) = arrName(1)

                'add middle name to the 3rd place
                mainArr(2) = arrName(2)


                'add last name to the 5th place
                mainArr(4) = arrName(3)

                'add suffix to 6th place in main array
                mainArr(5) = arrName(4)

                'add abbreviation to the 7th place
                mainArr(6) = Microsoft.VisualBasic.Left(mainArr(1), 1) & "." & Microsoft.VisualBasic.Left(mainArr(2), 1) & ". " & mainArr(4)


            ElseIf checkTitle(arrName(0)) = True And checkSuffix(arrName(4)) = False Then
                'Dr. John Phillip Van Smith

                'add title to the 1st place
                mainArr(0) = arrName(0)

                'add frist name to the 2nd place
                mainArr(1) = arrName(1)

                'add middle name to the 3rd place
                mainArr(2) = arrName(2)

                'add last name prefix in 4th place
                mainArr(3) = arrName(3)

                'add last name to the 5th place
                mainArr(4) = arrName(4)

                'add abbreviation to the 7th place
                mainArr(6) = Microsoft.VisualBasic.Left(mainArr(1), 1) & "." & Microsoft.VisualBasic.Left(mainArr(2), 1) & ". " & mainArr(3) & " " & mainArr(4)


            ElseIf checkTitle(arrName(0)) = False And checkSuffix(arrName(4)) = True Then
                'John Phillip Van Smith PhD

                'add first name to the 2nd place
                mainArr(1) = arrName(0)

                'add middle name to the 3rd place
                mainArr(2) = arrName(1)

                'add last name prefix in 4th place
                mainArr(3) = arrName(2)

                'add last name to the 5th place
                mainArr(4) = arrName(3)

                'add suffix to the 6th place
                mainArr(5) = arrName(4)

                'add abbreviation to the 7th place
                mainArr(6) = Microsoft.VisualBasic.Left(mainArr(1), 1) & "." & Microsoft.VisualBasic.Left(mainArr(2), 1) & ". " & mainArr(3) & " " & mainArr(4)

            ElseIf checkTitle(arrName(0)) = False And checkSuffix(arrName(4)) = False Then
                'John Phillip Van Der Smith

                'add first name to the 2nd place
                mainArr(1) = arrName(0)

                'add middle name to the 3rd place
                mainArr(2) = arrName(1)

                'add last name prefix in 4th place
                mainArr(3) = arrName(2) & " " & arrName(3)

                'add last name to the 5th plcae
                mainArr(4) = arrName(4)

                'add abbreviation to the 7th place
                mainArr(6) = Microsoft.VisualBasic.Left(mainArr(1), 1) & "." & Microsoft.VisualBasic.Left(mainArr(2), 1) & ". " & mainArr(3) & " " & mainArr(4)


            End If

        ElseIf UBound(arrName) >= 5 Then


            If checkTitle(arrName(0)) = True And checkSuffix(arrName(UBound(arrName))) = True Then
                'Dr. John Phillip Van ... Smith PhD

                'add title to 1st place in Main array
                mainArr(0) = arrName(0)

                'add first name to the 2nd plcae in the main array
                mainArr(1) = arrName(1)

                'add middle name to the 3rd place
                mainArr(2) = arrName(2)

                'add last name prefix in 4th place
                For k = 3 To UBound(arrName) - 2
                    mainArr(3) = mainArr(3) & " " & arrName(k)
                Next
                'remove any extra leading and trailing spaces
                mainArr(3) = Trim(mainArr(3))


                'add last name to the 5th place
                mainArr(4) = arrName(UBound(arrName) - 1)

                'mainArr(5) = arrName(UBound(arrName) - 1)


                'add abbreviation to the 7th place
                mainArr(6) = Microsoft.VisualBasic.Left(mainArr(1), 1) & "." & Microsoft.VisualBasic.Left(mainArr(2), 1) & ". " & mainArr(3) & " " & mainArr(4)


                'add suffix to 6th place in main array
                mainArr(5) = arrName(UBound(arrName))


            ElseIf checkTitle(arrName(0)) = True And checkSuffix(arrName(UBound(arrName))) = False Then
                'Dr. John Phillip Van Der ... Smith

                'add title to the 1st place
                mainArr(0) = arrName(0)

                'add first name to the 2nd place
                mainArr(1) = arrName(1)

                'add middle name to the 3rd place
                mainArr(2) = arrName(2)

                'add last name prefix in 4th place
                For k = 3 To UBound(arrName) - 1
                    mainArr(3) = mainArr(3) & " " & arrName(k)
                Next
                'remove any extra leading and trailing spaces
                mainArr(3) = Trim(mainArr(3))

                'add last name to the 5th place
                mainArr(4) = arrName(UBound(arrName))

                'add abbreviation to the 7th place
                mainArr(6) = Microsoft.VisualBasic.Left(mainArr(1), 1) & "." & Microsoft.VisualBasic.Left(mainArr(2), 1) & ". " & mainArr(3) & " " & mainArr(4)


            ElseIf checkTitle(arrName(0)) = False And checkSuffix(arrName(UBound(arrName))) = True Then
                'John Phillip Van Der ... Smith PhD

                'add first name to the 2nd place
                mainArr(1) = arrName(0)

                'add middle name to the 3rd place
                mainArr(2) = arrName(1)

                'add last name prefix in 4th place
                For k = 2 To UBound(arrName) - 2
                    mainArr(3) = mainArr(3) & " " & arrName(k)
                Next
                'remove any extra leading and trailing spaces
                mainArr(3) = Trim(mainArr(3))

                'add last name to the 5th place
                mainArr(4) = arrName(UBound(arrName) - 1)

                'add suffix to the 6th place
                mainArr(5) = arrName(UBound(arrName))

                'add abbreviation to the 7th place
                mainArr(6) = Microsoft.VisualBasic.Left(mainArr(1), 1) & "." & Microsoft.VisualBasic.Left(mainArr(2), 1) & ". " & mainArr(3) & " " & mainArr(4)

            ElseIf checkTitle(arrName(0)) = False And checkSuffix(arrName(UBound(arrName))) = False Then
                'John Phillip Van Der James ... Smith 

                'add first name to the 2nd place
                mainArr(1) = arrName(0)

                'add middle name to the 3rd place
                mainArr(2) = arrName(1)

                'add last name prefix in 4th place
                For k = 2 To UBound(arrName) - 1
                    mainArr(3) = mainArr(3) & " " & arrName(k)
                Next
                'remove any extra leading and trailing spaces
                mainArr(3) = Trim(mainArr(3))

                'add last name to the 5th plcae
                mainArr(4) = arrName(UBound(arrName))

                'add abbreviation to the 7th place
                mainArr(6) = Microsoft.VisualBasic.Left(mainArr(1), 1) & "." & Microsoft.VisualBasic.Left(mainArr(2), 1) & ". " & mainArr(3) & " " & mainArr(4)

            End If





        End If



    End Sub

    Private Sub display()

        Try


            CustomPanel2.Controls.Clear()



            Dim displayRng As Excel.Range = worksheet.Cells(1, 1)
            Dim arrRng As String()


            Dim outputColumn As Integer

            arrRng = Split(txtSourceRange.Text, ",")

            For i = 0 To UBound(arrRng)

                sourceRange = worksheet.Range(arrRng(i))

                For j = 1 To sourceRange.Rows.Count

                    mainArr = {"", "", "", "", "", "", "", ""}
                    name = sourceRange.Cells(j, 1).value

                    Call nameSplitter()

                    Dim headerIndex As Integer
                    Dim headerStr As String = ""
                    For Each ctrl As Control In CustomGroupBox7.Controls
                        If TypeOf ctrl Is System.Windows.Forms.CheckBox Then
                            Dim chk As System.Windows.Forms.CheckBox = DirectCast(ctrl, System.Windows.Forms.CheckBox)
                            If chk.Checked Then
                                headerStr = headerStr & "," & chk.Text
                            End If
                        End If
                    Next

                    headerStr = headerStr.Replace("Select All,", String.Empty)
                    headerStr = Microsoft.VisualBasic.Right(headerStr, Len(headerStr) - 1)

                    Dim arrHeaderStr As String() = Split(headerStr, ",")
                    outputColumn = UBound(arrHeaderStr) + 1
                    For k = 0 To UBound(arrHeaderStr)
                        worksheet.Cells(100000, k + 1).value = arrHeaderStr(k)

                        Select Case arrHeaderStr(k)
                            Case = "Title"
                                headerIndex = 0
                            Case = "First Name"
                                headerIndex = 1
                            Case = "Middle Name"
                                headerIndex = 2
                            Case = "Last Name Prefix"
                                headerIndex = 3
                            Case = "Last Name"
                                headerIndex = 4
                            Case = "Name Suffix"
                                headerIndex = 5
                            Case = "Name Abbreviations"
                                headerIndex = 6

                        End Select

                        worksheet.Cells(100000 + j, k + 1).value = mainArr(headerIndex)

                    Next





                Next

                displayRng = worksheet.Range(worksheet.Cells(100000, 1), worksheet.Cells(100000 + sourceRange.Rows.Count, outputColumn))

                If CB_Add_Header.Checked = False Then

                    displayRng = worksheet.Range(worksheet.Cells(100001, 1), worksheet.Cells(100000 + sourceRange.Rows.Count, outputColumn))

                End If

            Next


            If txtSourceRange.Text = "" Or displayRng Is Nothing Then
                CustomPanel2.Controls.Clear()
                Exit Sub
            End If


            If displayRng.Rows.Count > 50 Then
                displayRng = displayRng.Rows("1:50")
            End If


            Dim height As Double
            Dim width As Double

            If displayRng.Rows.Count <= 4 Then
                height = CustomPanel2.Height / displayRng.Rows.Count
            Else
                height = (119 / 4)
            End If

            If displayRng.Columns.Count <= 3 Then
                width = CustomPanel2.Width / displayRng.Columns.Count
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

                    If CB_Keep_Formatting.Checked = True Then
                        If CB_Add_Header.Checked = True Then

                            Dim cellColor As Color = ColorTranslator.FromOle(CType(sourceRange.Cells(i - 1, 1).Font.Color, Integer))
                            Dim fillColor As Color = ColorTranslator.FromOle(CType(sourceRange.Cells(i - 1, 1).interior.Color, Integer))


                            Dim cell As Excel.Range = sourceRange.Cells(i - 1, 1)

                            Dim cellFontName As String = cell.Font.Name
                            Dim cellFontSize As Single = Convert.ToSingle(10)
                            Dim cellFontColor As Color = ColorTranslator.FromOle(CType(cell.Font.Color, Integer))
                            Dim cellFontStyle As FontStyle = fontStyle.Regular
                            If cell.Font.Bold Then cellFontStyle = cellFontStyle Or fontStyle.Bold
                            If cell.Font.Italic Then cellFontStyle = cellFontStyle Or fontStyle.Italic
                            If cell.Font.Underline <> Excel.XlUnderlineStyle.xlUnderlineStyleNone Then cellFontStyle = cellFontStyle Or fontStyle.Underline

                            label.Font = New System.Drawing.Font(cellFontName, cellFontSize, cellFontStyle)


                            label.ForeColor = cellColor
                            label.BackColor = fillColor

                            'bold header
                            If i = 1 Then

                                cellFontStyle = cellFontStyle Or FontStyle.Bold
                                label.Font = New System.Drawing.Font(cellFontName, cellFontSize, cellFontStyle)
                            End If

                        Else

                            Dim cellColor As Color = ColorTranslator.FromOle(CType(sourceRange.Cells(i, 1).Font.Color, Integer))
                            Dim fillColor As Color = ColorTranslator.FromOle(CType(sourceRange.Cells(i, 1).interior.Color, Integer))


                            Dim cell As Excel.Range = sourceRange.Cells(i, 1)

                            Dim cellFontName As String = cell.Font.Name
                            Dim cellFontSize As Single = Convert.ToSingle(10)
                            Dim cellFontColor As Color = ColorTranslator.FromOle(CType(cell.Font.Color, Integer))
                            Dim cellFontStyle As FontStyle = FontStyle.Regular
                            If cell.Font.Bold Then cellFontStyle = cellFontStyle Or FontStyle.Bold
                            If cell.Font.Italic Then cellFontStyle = cellFontStyle Or FontStyle.Italic
                            If cell.Font.Underline <> Excel.XlUnderlineStyle.xlUnderlineStyleNone Then cellFontStyle = cellFontStyle Or FontStyle.Underline

                            label.Font = New System.Drawing.Font(cellFontName, cellFontSize, cellFontStyle)


                            label.ForeColor = cellColor
                            label.BackColor = fillColor


                        End If


                    Else
                        label.BackColor = Color.Transparent
                        label.ForeColor = Nothing



                    End If


                    CustomPanel2.Controls.Add(label)
                Next
            Next

            CustomPanel2.AutoScroll = True

            worksheet.Range(displayRng.Cells(1, 1).offset(-1, 0), displayRng.Cells(displayRng.Rows.Count, displayRng.Columns.Count)).EntireRow.Delete()

        Catch ex As Exception

        End Try



    End Sub



    Private Sub uncheck_CB_Select_All()
        For Each ctrl As Control In CustomGroupBox7.Controls
            If TypeOf ctrl Is System.Windows.Forms.CheckBox Then
                Dim chk As System.Windows.Forms.CheckBox = DirectCast(ctrl, System.Windows.Forms.CheckBox)
                'Do something with chk
                If chk.Checked = False Then
                    CB_Select_All.Checked = False
                End If
            End If
        Next
        Call display()

    End Sub


    Private Sub CB_Select_All_CheckedChanged(sender As Object, e As EventArgs) Handles CB_Select_All.CheckedChanged
        If CB_Select_All.Checked = True Then
            For Each ctrl As Control In CustomGroupBox7.Controls
                If TypeOf ctrl Is System.Windows.Forms.CheckBox Then
                    Dim chk As System.Windows.Forms.CheckBox = DirectCast(ctrl, System.Windows.Forms.CheckBox)

                    chk.Checked = True
                    chk.Enabled = False

                End If
            Next
            CB_Select_All.Checked = True
            CB_Select_All.Enabled = True

        ElseIf CB_Select_All.Checked = False Then
            For Each ctrl As Control In CustomGroupBox7.Controls
                If TypeOf ctrl Is System.Windows.Forms.CheckBox Then
                    Dim chk As System.Windows.Forms.CheckBox = DirectCast(ctrl, System.Windows.Forms.CheckBox)

                    chk.Checked = False
                    chk.Enabled = True

                End If
            Next


        End If

        Call display()


    End Sub

    Private Sub CB_Title_CheckedChanged(sender As Object, e As EventArgs) Handles CB_Title.CheckedChanged

        Call uncheck_CB_Select_All()

    End Sub

    Private Sub CB_First_Name_CheckedChanged(sender As Object, e As EventArgs) Handles CB_First_Name.CheckedChanged
        Call uncheck_CB_Select_All()
    End Sub

    Private Sub CB_Middle_Name_CheckedChanged(sender As Object, e As EventArgs) Handles CB_Middle_Name.CheckedChanged
        Call uncheck_CB_Select_All()
    End Sub

    Private Sub CB_Last_Name_Prefix_CheckedChanged(sender As Object, e As EventArgs) Handles CB_Last_Name_Prefix.CheckedChanged
        Call uncheck_CB_Select_All()
    End Sub

    Private Sub CB_Last_Name_CheckedChanged(sender As Object, e As EventArgs) Handles CB_Last_Name.CheckedChanged
        Call uncheck_CB_Select_All()
    End Sub



    Private Sub CB_Name_Abbreviations_CheckedChanged(sender As Object, e As EventArgs) Handles CB_Name_Abbreviations.CheckedChanged
        Call uncheck_CB_Select_All()
    End Sub



    Private Sub CB_Name_Suffix_CheckedChanged(sender As Object, e As EventArgs) Handles CB_Name_Suffix.CheckedChanged
        Call uncheck_CB_Select_All()
    End Sub

    Private Sub RB_Same_As_Source_Range_CheckedChanged(sender As Object, e As EventArgs) Handles RB_Same_As_Source_Range.CheckedChanged

        If RB_Same_As_Source_Range.Checked = True Then

            txtDestRange.Enabled = False
            destinationSelection.Enabled = False
            lbl_destRange_Selection.Enabled = False

        ElseIf RB_Same_As_Source_Range.Checked = False Then

            txtDestRange.Enabled = True
            destinationSelection.Enabled = True
            lbl_destRange_Selection.Enabled = True

        End If

    End Sub

    Private Sub CB_Add_Header_CheckedChanged(sender As Object, e As EventArgs) Handles CB_Add_Header.CheckedChanged
        Call display()
    End Sub

    Private Sub CB_Keep_Formatting_CheckedChanged(sender As Object, e As EventArgs) Handles CB_Keep_Formatting.CheckedChanged
        Call display()
    End Sub

    Private Sub Form17DivideNames_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        form_flag = False
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.Dispose()
    End Sub

    Private Sub Form17DivideNames_Disposed(sender As Object, e As EventArgs) Handles Me.Disposed
        form_flag = False
    End Sub

    Private Sub Form17DivideNames_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        Me.Focus()
        Me.BringToFront()
        Me.Activate()
        Me.BeginInvoke(New System.Action(Sub()
                                             txtSourceRange.Text = sourceRange.Address
                                             SetWindowPos(Me.Handle, New IntPtr(HWND_TOPMOST), 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOSIZE)
                                         End Sub))
    End Sub
End Class