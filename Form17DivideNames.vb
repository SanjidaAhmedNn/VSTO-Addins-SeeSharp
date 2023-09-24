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

            Call nameSplitter2()



            Dim headerStr As String = ""
            For Each ctrl As Control In CustomGroupBox7.Controls
                If TypeOf ctrl Is System.Windows.Forms.CheckBox Then
                    Dim chk As System.Windows.Forms.CheckBox = DirectCast(ctrl, System.Windows.Forms.CheckBox)
                    'Do something with chk
                    If chk.Checked Then
                        'For example, print the name of the checkbox that's checked
                        'Console.WriteLine(chk.Name & " is checked")
                        headerStr = headerStr & "," & chk.Text
                        'MsgBox(chk.Text)
                    End If
                End If
            Next
            headerStr = "Full Name" & headerStr
            headerStr = headerStr.Replace("Select All,", String.Empty)

            Dim arrHeaderStr As String() = Split(headerStr, ",")


            If RB_Same_As_Source_Range.Checked = True Then
                If CB_Add_Header.Checked = True Then
                    For i = 0 To UBound(arrHeaderStr)
                        sourceRange.Cells(1, i + 1).value = arrHeaderStr(i)
                    Next


                End If


            ElseIf RB_Different_Range.Checked = True Then
                If CB_Add_Header.Checked = True Then

                    For i = 0 To UBound(arrHeaderStr)
                        destRange.Cells(1, i + 1).value = arrHeaderStr(i)
                    Next


                End If

            End If




            'MsgBox(headerStr)






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




        Dim arrHeader() = {"Full Name", "Title", "First Name", "Middle Name", "Last Name Prefix", "Last Name", "Abbreviations", "Name Suffix"}
        'Dim arrSplitName As String()


        arrRng = Split(txtSourceRange.Text, ",")

        For i = 0 To UBound(arrRng)

            selectedRange = worksheet.Range(arrRng(i))

            For j = 1 To selectedRange.Rows.Count

                destRange.Cells(j, 1).value = selectedRange.Cells(j, 1).value
                arrName = Split(selectedRange.Cells(j, 1).value, " ")





                Dim dotCount As Integer


                'Name Suffix


                'count if there is any periods in the last word of the name
                'if a period is there then it will be considered as a Name Suffix
                dotCount = 0
                For Each c As Char In arrName(UBound(arrName))

                    If c = "." Then
                        dotCount += 1
                    End If

                Next

                'checks if there are period(s) in the last word
                'OR the last word matches with any of the word from the arrSuffix array (case insensitively)
                'if any one of the 2 conditon is true then, assign the last word as value in the last column of the destRange
                'otherwise assign a blank value
                If dotCount > 0 Or arrSuffix.Contains(arrName(UBound(arrName)), StringComparer.OrdinalIgnoreCase) Then
                    'MsgBox("Title")
                    'For p = 1 To sourceRange.Rows.Count
                    'For q = 1 To 8
                    destRange.Cells(j, 8).value = arrName(UBound(arrName))
                    'Next
                    'Next
                    destRange.Cells(j, 6).value = arrName(UBound(arrName) - 1)

                Else
                    destRange.Cells(j, 8).value = ""
                    destRange.Cells(j, 6).value = arrName(UBound(arrName))
                End If



                'Title


                dotCount = 0

                'count if there is any periods in the first word of the name
                'if a period is there then it will be considered as a title
                For Each c As Char In arrName(0)

                    If c = "." Then
                        dotCount += 1
                    End If

                Next

                'checks if there are period(s) in the first word
                'OR the first word matches with any of the word from the arrTitle array (case insensitively)
                'if any one of the 2 conditon is true then, assign the first word as value in the first column of the destRange
                'otherwise assign a blank value
                If dotCount > 0 Or arrTitle.Contains(arrName(0), StringComparer.OrdinalIgnoreCase) Then
                    'MsgBox("Title")
                    'For p = 1 To sourceRange.Rows.Count
                    'For q = 1 To 8
                    destRange.Cells(j, 2).value = arrName(0)
                    'Next
                    'Next
                    destRange.Cells(j, 3).value = arrName(1)
                    destRange.Cells(j, 4).value = arrName(2)
                    destRange.Cells(j, 5).value = arrName(3) & " " & arrName(4)
                    destRange.Cells(j, 7).value = Microsoft.VisualBasic.Left(destRange.Cells(j, 3).value, 1) & "." & Microsoft.VisualBasic.Left(destRange.Cells(j, 4).value, 1) & ". " & destRange.Cells(j, 5).value & " " & destRange.Cells(j, 6).value


                Else
                    destRange.Cells(j, 2).value = ""
                    destRange.Cells(j, 3).value = arrName(0)
                    destRange.Cells(j, 4).value = arrName(1)
                    destRange.Cells(j, 5).value = arrName(2) & " " & arrName(3)
                    destRange.Cells(j, 7).value = Microsoft.VisualBasic.Left(destRange.Cells(j, 3).value, 1) & "." & Microsoft.VisualBasic.Left(destRange.Cells(j, 4).value, 1) & ". " & destRange.Cells(j, 5).value & " " & destRange.Cells(j, 6).value
                End If




            Next


        Next


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




    Sub nameSplitter2()

        excelApp = Globals.ThisAddIn.Application
        workbook = excelApp.ActiveWorkbook
        worksheet = workbook.ActiveSheet
        selectedRange = excelApp.Selection
        Dim arrRng As String()
        Dim arrName As String()

        Dim mainArr(7) As String


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


        Dim arrHeader() = {"Full Name", "Title", "First Name", "Middle Name", "Last Name Prefix", "Last Name", "Abbreviations", "Name Suffix"}
        'Dim arrSplitName As String()


        arrRng = Split(txtSourceRange.Text, ",")

        For i = 0 To UBound(arrRng)

            selectedRange = worksheet.Range(arrRng(i))

            For j = 1 To selectedRange.Rows.Count

                mainArr = {"", "", "", "", "", "", "", ""}
                mainArr(0) = selectedRange.Cells(j, 1).value

                arrName = Split(selectedRange.Cells(j, 1).value, " ")



                If UBound(arrName) = 0 Then

                    If checkTitle(arrName(0)) = True Then
                        mainArr(1) = arrName(0)

                    ElseIf checkSuffix(arrName(0)) = True Then
                        mainArr(7) = arrName(0)

                    Else
                        mainArr(2) = arrName(0)

                    End If


                ElseIf UBound(arrName) = 1 Then

                    If checkTitle(arrName(0)) = True And checkSuffix(arrName(1)) = True Then
                        'Dr. PhD

                        'add title to 2nd place in Mainarray
                        mainArr(1) = arrName(0)

                        'add suffix to last place in main array
                        mainArr(7) = arrName(1)

                    ElseIf checkTitle(arrName(0)) = True And checkSuffix(arrName(1)) = False Then
                        'Dr. John

                        'add title in the 2nd place
                        mainArr(1) = arrName(0)

                        'add first name to the 3rd place 
                        mainArr(2) = arrName(1)

                    ElseIf checkTitle(arrName(0)) = False And checkSuffix(arrName(1)) = True Then
                        'John PhD

                        'add first name to the 3rd place
                        mainArr(2) = arrName(0)

                        'add suffix to the last place
                        mainArr(7) = arrName(1)

                    ElseIf checkTitle(arrName(0)) = False And checkSuffix(arrName(1)) = False Then
                        'John Smith

                        'add first name to the 3rd place
                        mainArr(2) = arrName(0)

                        'add last name in the 6th place
                        mainArr(5) = arrName(1)

                    End If

                ElseIf UBound(arrName) = 2 Then

                    If checkTitle(arrName(0)) = True And checkSuffix(arrName(2)) = True Then
                        'Dr. John PhD

                        'add title to 2nd place in Mainarray
                        mainArr(1) = arrName(0)

                        'add first name to the 3rd plcae in the main array
                        mainArr(2) = arrName(1)

                        'add suffix to last place in main array
                        mainArr(7) = arrName(2)


                    ElseIf checkTitle(arrName(0)) = True And checkSuffix(arrName(2)) = False Then
                        'Dr. John Smith

                        'add title to the 2nd place
                        mainArr(1) = arrName(0)

                        'add frist name to the 3rd place
                        mainArr(2) = arrName(1)

                        'add last name to the 6th place
                        mainArr(5) = arrName(2)

                    ElseIf checkTitle(arrName(0)) = False And checkSuffix(arrName(2)) = True Then
                        'John Smith PhD

                        'add first name to the 3rd place
                        mainArr(2) = arrName(0)

                        'add last name to the 6th place
                        mainArr(5) = arrName(1)

                        'add suffix to the last place
                        mainArr(7) = arrName(2)
                    ElseIf checkTitle(arrName(0)) = False And checkSuffix(arrName(2)) = False Then
                        'John Phillip Smith

                        'add first name to the 3rd place
                        mainArr(2) = arrName(0)

                        'add middle name to the 4th place
                        mainArr(3) = arrName(1)

                        'add last name to the 6th plcae
                        mainArr(5) = arrName(2)

                        'add abbreviation to the 7th place
                        mainArr(6) = Microsoft.VisualBasic.Left(mainArr(2), 1) & "." & Microsoft.VisualBasic.Left(mainArr(3), 1) & ". " & mainArr(5)


                    End If


                ElseIf UBound(arrName) = 3 Then


                    If checkTitle(arrName(0)) = True And checkSuffix(arrName(3)) = True Then
                        'Dr. John Smith PhD

                        'add title to 2nd place in Main array
                        mainArr(1) = arrName(0)

                        'add first name to the 3rd plcae in the main array
                        mainArr(2) = arrName(1)

                        'add last name to the 6th place
                        mainArr(5) = arrName(2)

                        'add suffix to last place in main array
                        mainArr(7) = arrName(3)


                    ElseIf checkTitle(arrName(0)) = True And checkSuffix(arrName(3)) = False Then
                        'Dr. John Phillip Smith

                        'add title to the 2nd place
                        mainArr(1) = arrName(0)

                        'add frist name to the 3rd place
                        mainArr(2) = arrName(1)

                        'add middle name to the 4th place
                        mainArr(3) = arrName(2)

                        'add last name to the 6th place
                        mainArr(5) = arrName(3)

                        'add abbreviation to the 7th place
                        mainArr(6) = Microsoft.VisualBasic.Left(mainArr(2), 1) & "." & Microsoft.VisualBasic.Left(mainArr(3), 1) & ". " & mainArr(5)


                    ElseIf checkTitle(arrName(0)) = False And checkSuffix(arrName(3)) = True Then
                        'John Phillip Smith PhD

                        'add first name to the 3rd place
                        mainArr(2) = arrName(0)

                        'add middle name to the 4th place
                        mainArr(3) = arrName(1)

                        'add last name to the 6th place
                        mainArr(5) = arrName(2)

                        'add suffix to the last place
                        mainArr(7) = arrName(3)

                        'add abbreviation to the 7th place
                        mainArr(6) = Microsoft.VisualBasic.Left(mainArr(2), 1) & "." & Microsoft.VisualBasic.Left(mainArr(3), 1) & ". " & mainArr(5)

                    ElseIf checkTitle(arrName(0)) = False And checkSuffix(arrName(3)) = False Then
                        'John Phillip Van Smith

                        'add first name to the 3rd place
                        mainArr(2) = arrName(0)

                        'add middle name to the 4th place
                        mainArr(3) = arrName(1)

                        'add last name prefix in 5 th place
                        mainArr(4) = arrName(2)

                        'add last name to the 6th plcae
                        mainArr(5) = arrName(3)

                        'add abbreviation to the 7th place
                        mainArr(6) = Microsoft.VisualBasic.Left(mainArr(2), 1) & "." & Microsoft.VisualBasic.Left(mainArr(3), 1) & ". " & mainArr(4) & mainArr(5)


                    End If

                ElseIf UBound(arrName) = 4 Then

                    If checkTitle(arrName(0)) = True And checkSuffix(arrName(4)) = True Then
                        'Dr. John Phillip Smith PhD

                        'add title to 2nd place in Main array
                        mainArr(1) = arrName(0)

                        'add first name to the 3rd plcae in the main array
                        mainArr(2) = arrName(1)

                        'add middle name to the 4th place
                        mainArr(3) = arrName(2)

                        'add last name prefix in 5 th place


                        'If UBound(arrName) - 2 - 3 >= 0 Then
                        '    For k = 3 To UBound(arrName) - 2
                        '        mainArr(4) = mainArr(4) & " " & arrName(k)
                        '    Next
                        '    mainArr(4) = Trim(mainArr(4))
                        'End If



                        'add last name to the 5th place
                        mainArr(5) = arrName(3)

                        'mainArr(5) = arrName(UBound(arrName) - 1)


                        'add abbreviation to the 7th place
                        mainArr(6) = Microsoft.VisualBasic.Left(mainArr(2), 1) & "." & Microsoft.VisualBasic.Left(mainArr(3), 1) & ". " & mainArr(5)


                        'add suffix to last place in main array
                        mainArr(7) = arrName(4)


                    ElseIf checkTitle(arrName(0)) = True And checkSuffix(arrName(4)) = False Then
                        'Dr. John Phillip Van Smith

                        'add title to the 2nd place
                        mainArr(1) = arrName(0)

                        'add frist name to the 3rd place
                        mainArr(2) = arrName(1)

                        'add middle name to the 4th place
                        mainArr(3) = arrName(2)

                        'add last name prefix in 5 th place
                        mainArr(4) = arrName(3)

                        'add last name to the 6th place
                        mainArr(5) = arrName(4)

                        'add abbreviation to the 7th place
                        mainArr(6) = Microsoft.VisualBasic.Left(mainArr(2), 1) & "." & Microsoft.VisualBasic.Left(mainArr(3), 1) & ". " & mainArr(4) & " " & mainArr(5)


                    ElseIf checkTitle(arrName(0)) = False And checkSuffix(arrName(4)) = True Then
                        'John Phillip Van Smith PhD

                        'add first name to the 3rd place
                        mainArr(2) = arrName(0)

                        'add middle name to the 4th place
                        mainArr(3) = arrName(1)

                        'add last name prefix in 5 th place
                        mainArr(4) = arrName(2)

                        'add last name to the 6th place
                        mainArr(5) = arrName(3)

                        'add suffix to the last place
                        mainArr(7) = arrName(4)

                        'add abbreviation to the 7th place
                        mainArr(6) = Microsoft.VisualBasic.Left(mainArr(2), 1) & "." & Microsoft.VisualBasic.Left(mainArr(3), 1) & ". " & mainArr(4) & " " & mainArr(5)

                    ElseIf checkTitle(arrName(0)) = False And checkSuffix(arrName(4)) = False Then
                        'John Phillip Van Der Smith

                        'add first name to the 3rd place
                        mainArr(2) = arrName(0)

                        'add middle name to the 4th place
                        mainArr(3) = arrName(1)

                        'add last name prefix in 5 th place
                        mainArr(4) = arrName(2) & " " & arrName(3)

                        'add last name to the 6th plcae
                        mainArr(5) = arrName(4)

                        'add abbreviation to the 7th place
                        mainArr(6) = Microsoft.VisualBasic.Left(mainArr(2), 1) & "." & Microsoft.VisualBasic.Left(mainArr(3), 1) & ". " & mainArr(4) & " " & mainArr(5)


                    End If

                ElseIf UBound(arrName) >= 5 Then


                    If checkTitle(arrName(0)) = True And checkSuffix(arrName(UBound(arrName))) = True Then
                        'Dr. John Phillip Van ... Smith PhD

                        'add title to 2nd place in Main array
                        mainArr(1) = arrName(0)

                        'add first name to the 3rd plcae in the main array
                        mainArr(2) = arrName(1)

                        'add middle name to the 4th place
                        mainArr(3) = arrName(2)

                        'add last name prefix in 5 th place
                        For k = 3 To UBound(arrName) - 2
                            mainArr(4) = mainArr(4) & " " & arrName(k)
                        Next
                        'remove any extra leading and trailing spaces
                        mainArr(4) = Trim(mainArr(4))


                        'add last name to the 5th place
                        mainArr(5) = arrName(UBound(arrName) - 1)

                        'mainArr(5) = arrName(UBound(arrName) - 1)


                        'add abbreviation to the 7th place
                        mainArr(6) = Microsoft.VisualBasic.Left(mainArr(2), 1) & "." & Microsoft.VisualBasic.Left(mainArr(3), 1) & ". " & mainArr(4) & " " & mainArr(5)


                        'add suffix to last place in main array
                        mainArr(7) = arrName(UBound(arrName))


                    ElseIf checkTitle(arrName(0)) = True And checkSuffix(arrName(UBound(arrName))) = False Then
                        'Dr. John Phillip Van Der ... Smith

                        'add title to the 2nd place
                        mainArr(1) = arrName(0)

                        'add frist name to the 3rd place
                        mainArr(2) = arrName(1)

                        'add middle name to the 4th place
                        mainArr(3) = arrName(2)

                        'add last name prefix in 5 th place
                        For k = 3 To UBound(arrName) - 2
                            mainArr(4) = mainArr(4) & " " & arrName(k)
                        Next
                        'remove any extra leading and trailing spaces
                        mainArr(4) = Trim(mainArr(4))

                        'add last name to the 6th place
                        mainArr(5) = arrName(UBound(arrName))

                        'add abbreviation to the 7th place
                        mainArr(6) = Microsoft.VisualBasic.Left(mainArr(2), 1) & "." & Microsoft.VisualBasic.Left(mainArr(3), 1) & ". " & mainArr(4) & " " & mainArr(5)


                    ElseIf checkTitle(arrName(0)) = False And checkSuffix(arrName(UBound(arrName))) = True Then
                        'John Phillip Van Der ... Smith PhD

                        'add first name to the 3rd place
                        mainArr(2) = arrName(0)

                        'add middle name to the 4th place
                        mainArr(3) = arrName(1)

                        'add last name prefix in 5 th place
                        For k = 2 To UBound(arrName) - 2
                            mainArr(4) = mainArr(4) & " " & arrName(k)
                        Next
                        'remove any extra leading and trailing spaces
                        mainArr(4) = Trim(mainArr(4))

                        'add last name to the 6th place
                        mainArr(5) = arrName(UBound(arrName) - 1)

                        'add suffix to the last place
                        mainArr(7) = arrName(UBound(arrName))

                        'add abbreviation to the 7th place
                        mainArr(6) = Microsoft.VisualBasic.Left(mainArr(2), 1) & "." & Microsoft.VisualBasic.Left(mainArr(3), 1) & ". " & mainArr(4) & " " & mainArr(5)

                    ElseIf checkTitle(arrName(0)) = False And checkSuffix(arrName(UBound(arrName))) = False Then
                        'John Phillip Van Der James ... Smith 

                        'add first name to the 3rd place
                        mainArr(2) = arrName(0)

                        'add middle name to the 4th place
                        mainArr(3) = arrName(1)

                        'add last name prefix in 5 th place
                        For k = 2 To UBound(arrName) - 2
                            mainArr(4) = mainArr(4) & " " & arrName(k)
                        Next
                        'remove any extra leading and trailing spaces
                        mainArr(4) = Trim(mainArr(4))

                        'add last name to the 6th plcae
                        mainArr(5) = arrName(UBound(arrName) - 1)

                        'add abbreviation to the 7th place
                        mainArr(6) = Microsoft.VisualBasic.Left(mainArr(2), 1) & "." & Microsoft.VisualBasic.Left(mainArr(3), 1) & ". " & mainArr(4) & " " & mainArr(5)


                    End If





                End If





                MsgBox("Full name is " & mainArr(0))
                MsgBox("title is " & mainArr(1))
                MsgBox("First name is " & mainArr(2))
                MsgBox("Middle name is " & mainArr(3))
                MsgBox("Last name prefix is " & mainArr(4))
                MsgBox("Last name is " & mainArr(5))
                MsgBox("Abbreviation is " & mainArr(6))
                MsgBox("suffix is " & mainArr(7))
            Next

        Next



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
    End Sub


    Private Sub CB_Select_All_CheckedChanged(sender As Object, e As EventArgs) Handles CB_Select_All.CheckedChanged
        If CB_Select_All.Checked = True Then
            For Each ctrl As Control In CustomGroupBox7.Controls
                If TypeOf ctrl Is System.Windows.Forms.CheckBox Then
                    Dim chk As System.Windows.Forms.CheckBox = DirectCast(ctrl, System.Windows.Forms.CheckBox)

                    chk.Checked = True

                End If
            Next
            CB_Select_All.Checked = True
        End If


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


End Class