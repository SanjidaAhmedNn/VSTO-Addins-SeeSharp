Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Runtime.InteropServices
Imports System.Drawing
Imports Microsoft.Office.Interop.Excel
Imports System.Net.Mime.MediaTypeNames
Imports System.Reflection
Imports System.Windows.Forms
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports System.Diagnostics
Imports System.Text.RegularExpressions


Public Class Form1
    Dim WithEvents excelApp As Excel.Application
    Dim workBook As Excel.Workbook
    Dim workSheet As Excel.Worksheet
    Dim workSheet2 As Excel.Worksheet
    Dim rng As Excel.Range
    Dim rng2 As Excel.Range
    Dim selectedRange As Excel.Range

    Dim opened As Integer
    Dim FocusedTextBox As Integer
    Private Function Overlap(excelApp As Excel.Application, sheet1 As Excel.Worksheet, sheet2 As Excel.Worksheet, rng1 As Excel.Range, rng2 As Excel.Range) As Boolean

        If sheet1.Name <> sheet2.Name Then
            Return False

        Else
            Dim activesheet As Excel.Worksheet = CType(excelApp.ActiveSheet, Excel.Worksheet)

            Dim rng3 As Excel.Range = activesheet.Range(rng1.Address)
            Dim rng4 As Excel.Range = activesheet.Range(rng2.Address)

            Dim intersectRange As Range = excelApp.Intersect(rng3, rng4)

            If intersectRange Is Nothing Then
                Return False
            Else
                Return True
            End If
        End If

    End Function
    Private Function IsValidExcelCellReference(cellReference As String) As Boolean

        ' Regular expression pattern for a cell reference.
        ' This pattern will match references like A1, $A$1, etc.
        Dim cellPattern As String = "(\$?[A-Z]+\$?[0-9]+)"

        ' Regular expression pattern for an Excel reference.
        ' This pattern will match references like A1:B13, $A$1:$B$13, A1, $B$1, etc.
        Dim referencePattern As String = "^" + cellPattern + "(:" + cellPattern + ")?$"

        ' Create a regex object with the pattern.
        Dim regex As New Regex(referencePattern)

        ' Test the input string against the regex pattern.
        If regex.IsMatch(cellReference) Then
            Return True
        Else
            Return False
        End If


    End Function
    Private Function IsWithin(rng1 As Excel.Range, rng2 As Excel.Range)

        Dim Result As Boolean

            Dim sr1 As Integer
            Dim sr2 As Integer

            Dim er1 As Integer
            Dim er2 As Integer

            Dim sc1 As Integer
            Dim sc2 As Integer

            Dim ec1 As Integer
            Dim ec2 As Integer

            sr1 = rng1.Cells(1, 1).Row
            sc1 = rng1.Cells(1, 1).Column

            er1 = rng1.Cells(rng1.Rows.Count, rng1.Columns.Count).Row
            ec1 = rng1.Cells(rng1.Rows.Count, rng1.Columns.Count).Column

            sr2 = rng2.Cells(1, 1).Row
            sc2 = rng2.Cells(1, 1).Column

            er2 = rng2.Cells(rng2.Rows.Count, rng2.Columns.Count).Row
            ec2 = rng2.Cells(rng2.Rows.Count, rng2.Columns.Count).Column

            If sr1 >= sr2 And sc1 >= sc2 And er1 <= er2 And ec1 <= ec2 Then
                Result = True
            Else
                Result = False
            End If

            IsWithin = Result


    End Function
    Private Function ReplaceNotInRange(input As String, find As String, replaceWith As String) As String

        ' Build the regex pattern to exclude range notation 
        Dim pattern As String = String.Format("(?<!:)\b{0}\b(?!:)", Regex.Escape(find))

            ' Create a Regex object.
            Dim reg As New Regex(pattern)

        ' Call the Regex.Replace method to replace matching text.
        Return reg.Replace(input, replaceWith)
    End Function
    Private Function ReplaceReference(Ref As String, rng As Excel.Range, rng2 As Excel.Range, type As Integer)

        If InStr(1, Ref, "!") > 0 Then
            ReplaceReference = Ref
        Else

            Dim activesheet As Excel.Worksheet = CType(excelApp.ActiveSheet, Excel.Worksheet)

            Dim colNum As Integer
            Dim rowNum As Integer
            Dim colNum2 As Integer
            Dim rowNum2 As Integer
            Dim colName As String
            Dim rowName As String
            Dim colName2 As String
            Dim rowName2 As String
            Dim expRange As Excel.Range
            Dim Ext As Integer
            Dim Ext2 As Integer
            Dim Ref2 As String
            Dim Ref3 As String
            Dim distance1 As Integer
            Dim distance2 As Integer

            distance1 = rng2.Cells(1, 1).Row - rng.Cells(1, 1).Row
            distance2 = rng2.Cells(1, 1).Column - rng.Cells(1, 1).Column

            expRange = activesheet.Range(Ref)

            If type = 1 Then
                colNum = expRange.Column
                colName = Split(activesheet.Cells(1, colNum).Address, "$")(1)
                Ext = colNum - rng.Cells(1, 1).Column + 1
                Ext2 = rng.Columns.Count - Ext + 1
                colNum2 = rng.Cells(1, 1).Column - 1 + Ext2
                colName2 = Split(activesheet.Cells(1, colNum2).Address, "$")(1)
                Ref2 = Replace(Ref, colName, colName2)
                expRange = activesheet.Range(Ref2)
                rowNum = expRange.Row
                colNum = expRange.Column
                rowNum2 = rowNum + distance1
                colNum2 = colNum + distance2
                rowName = Split(activesheet.Cells(rowNum, 1).Address, "$")(2)
                rowName2 = Split(activesheet.Cells(rowNum2, 1).Address, "$")(2)
                colName = Split(activesheet.Cells(1, colNum).Address, "$")(1)
                colName2 = Split(activesheet.Cells(1, colNum2).Address, "$")(1)
                Ref3 = Replace(Ref2, rowName, rowName2)
                Ref3 = Replace(Ref3, colName, colName2)
            Else
                rowNum = expRange.Row
                rowName = Split(activesheet.Cells(rowNum, 1).Address, "$")(2)
                Ext = rowNum - rng.Cells(1, 1).Row + 1
                Ext2 = rng.Rows.Count - Ext + 1
                rowNum2 = rng.Cells(1, 1).Row - 1 + Ext2
                rowName2 = Split(activesheet.Cells(rowNum2, 1).Address, "$")(2)
                Ref2 = Replace(Ref, rowName, rowName2)
                expRange = activesheet.Range(Ref2)
                rowNum = expRange.Row
                colNum = expRange.Column
                rowNum2 = rowNum + distance1
                colNum2 = colNum + distance2
                rowName = Split(activesheet.Cells(rowNum, 1).Address, "$")(2)
                rowName2 = Split(activesheet.Cells(rowNum2, 1).Address, "$")(2)
                colName = Split(activesheet.Cells(1, colNum).Address, "$")(1)
                colName2 = Split(activesheet.Cells(1, colNum2).Address, "$")(1)
                Ref3 = Replace(Ref2, rowName, rowName2)
                Ref3 = Replace(Ref3, colName, colName2)
            End If


            ReplaceReference = Ref3
        End If
    End Function
    Private Function ReplaceRange(Ref As String, rng As Excel.Range, rng2 As Excel.Range, Type As Integer)

        If InStr(1, Ref, "!") > 0 Then
                ReplaceRange = Ref
            Else
                Dim Ref1 As String
                Dim Ref2 As String

                Dim R1() As String
                R1 = Split(Ref, ":")
                Ref1 = R1(0)
                Ref2 = R1(1)

                Ref1 = ReplaceReference(Ref1, rng, rng2, Type)
                Ref2 = ReplaceReference(Ref2, rng, rng2, Type)

                Dim NewRef As String
                NewRef = Ref1 & ":" & Ref2

                ReplaceRange = NewRef
            End If

    End Function
    Private Function ReplaceFormula(Formula As String, Rng As Excel.Range, rng2 As Excel.Range, Type As Integer, sheet1 As Excel.Worksheet, sheet2 As Excel.Worksheet)


        Dim activesheet As Excel.Worksheet = CType(excelApp.ActiveSheet, Excel.Worksheet)

            Dim Starters As String() = New String() {"=", "(", ",", " ", "+", "-", "*", "/", "^", ")"}

            Dim Arr() As String

            Dim Index As Integer
            Index = -1


            Dim Arr1() As Integer

            Dim Index1 As Integer
            Index1 = -1

            Dim Refs() As String

            Dim i As Integer
            Dim j As Integer


            For i = 1 To Len(Formula)
                For j = LBound(Starters) To UBound(Starters)
                    If Mid(Formula, i, 1) = Starters(j) Then
                        Index1 = Index1 + 1
                        ReDim Preserve Arr1(Index1)
                        Arr1(Index1) = i
                        Exit For
                    End If
                Next j
            Next i

            Index1 = Index1 + 1
            ReDim Preserve Arr1(Index1)
            Arr1(Index1) = Len(Formula) + 1

            Dim Start As Integer
            Dim Ending As Integer
            Dim Ref As String

            For i = LBound(Arr1) To UBound(Arr1) - 1
                Index = Index + 1
                Start = Arr1(i)
                Ending = Arr1(i + 1)
                Ref = Mid(Formula, Start + 1, Ending - Start - 1)
                ReDim Preserve Arr(Index)
                Arr(Index) = Ref
            Next i

            Index = -1

            Dim C1 As Boolean
            Dim C2 As Boolean
            Dim C3 As Boolean

            For i = LBound(Arr) To UBound(Arr)

                If Arr(i) <> "" Then
                    C1 = Asc(Mid(Arr(i), Len(Arr(i)), 1)) >= 48 And Asc(Mid(Arr(i), Len(Arr(i)), 1)) <= 57
                    C2 = Asc(Mid(Arr(i), 1, 1)) >= 65 And Asc(Mid(Arr(i), 1, 1)) <= 90
                    C3 = Asc(Mid(Arr(i), 1, 1)) >= 97 And Asc(Mid(Arr(i), 1, 1)) <= 122

                    If (C1 And (C2 Or C3)) Then
                        Index = Index + 1
                        ReDim Preserve Refs(Index)
                        Refs(Index) = Arr(i)
                    End If
                End If
            Next i

            Dim expRange As Excel.Range

            For Each Ref In Refs

                If InStr(1, Ref, ":") = 0 Then
                    If InStr(1, Ref, "!") = 0 Then
                        expRange = activesheet.Range(Ref)
                    Else
                        Dim exp() As String
                        exp = Split(Ref, "!")
                    expRange = activesheet.Range(exp(1))
                End If
                    If IsWithin(expRange, Rng) = True Then
                        Dim Ref2 As String
                        Ref2 = ReplaceReference(Ref, Rng, rng2, Type)
                        Formula = ReplaceNotInRange(Formula, Ref, Ref2)
                    Else
                        If sheet1.Name <> sheet2.Name Then
                            Dim Ref2 As String
                            Ref2 = sheet1.Name & "!" & Ref
                            Formula = ReplaceNotInRange(Formula, Ref, Ref2)
                        End If
                    End If
                Else
                    If InStr(1, Ref, "!") = 0 Then
                        expRange = activesheet.Range(Ref)
                    Else
                        Dim e1() As String
                        Dim exp1() As String
                        Dim exp2() As String

                        e1 = Split(Ref, ":")
                        Dim S1 As String = e1(0)
                        Dim S2 As String = e1(1)

                        exp1 = Split(S1, "!")
                        exp2 = Split(S2, "!")

                        Dim S3 As String = exp1(1)
                        Dim S4 As String = exp2(1)

                        expRange = activesheet.Range(S3 & ":" & S4)
                    End If
                    If IsWithin(expRange, Rng) = True Then
                        Dim Ref2 As String
                        Ref2 = ReplaceRange(Ref, Rng, rng2, Type)
                        Formula = Replace(Formula, Ref, Ref2)
                    Else
                        If sheet1.Name <> sheet2.Name Then
                            Dim R1() As String
                            R1 = Split(Ref, ":")
                            Dim Rf1 As String
                            Dim Rf2 As String
                            Rf1 = R1(0)
                            Rf2 = R1(1)
                            Dim Ref2 As String
                            Ref2 = sheet1.Name & "!" & Rf1 & ":" & sheet1.Name & "!" & Rf2
                            Formula = Replace(Formula, Ref, Ref2)
                        End If
                    End If
                End If

            Next Ref

            ReplaceFormula = Formula

    End Function

    Private Sub Display()

        Try

            panel1.Controls.Clear()
            panel2.Controls.Clear()

            Dim displayRng As Excel.Range

            If rng.Rows.Count > 50 Then
                displayRng = rng.Rows("1:50")
            Else
                displayRng = rng
            End If


            Dim height As Double
            Dim width As Double

            If displayRng.Rows.Count <= 4 Then
                height = panel1.Height / displayRng.Rows.Count
            Else
                height = (119 / 4)
            End If

            If displayRng.Columns.Count <= 3 Then
                width = panel1.Width / displayRng.Columns.Count
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

                    If CheckBox2.Checked = True Then
                        Dim cell As Excel.Range = displayRng.Cells(i, j)
                        Dim font As Excel.Font = cell.Font

                        Dim fontStyle As FontStyle = FontStyle.Regular
                        If cell.Font.Bold Then fontStyle = fontStyle Or FontStyle.Bold
                        If cell.Font.Italic Then fontStyle = fontStyle Or FontStyle.Italic


                        Dim fontSize As Single = Convert.ToSingle(font.Size)

                        label.Font = New System.Drawing.Font(font.ToString, fontSize, fontStyle)
                        If Not cell.Interior.ColorIndex = Excel.XlColorIndex.xlColorIndexNone Then
                            Dim colorValue1 As Long = CLng(cell.Interior.Color)
                            Dim red1 As Integer = colorValue1 Mod 256
                            Dim green1 As Integer = (colorValue1 \ 256) Mod 256
                            Dim blue1 As Integer = (colorValue1 \ 256 \ 256) Mod 256
                            label.BackColor = System.Drawing.Color.FromArgb(red1, green1, blue1)
                        End If
                        If Not cell.Font.ColorIndex = Excel.XlColorIndex.xlColorIndexNone Then
                            Dim colorValue2 As Long = CLng(cell.Font.Color)
                            Dim red2 As Integer = colorValue2 Mod 256
                            Dim green2 As Integer = (colorValue2 \ 256) Mod 256
                            Dim blue2 As Integer = (colorValue2 \ 256 \ 256) Mod 256
                            label.ForeColor = System.Drawing.Color.FromArgb(red2, green2, blue2)
                        End If
                    End If
                    panel1.Controls.Add(label)
                Next
            Next

            panel1.AutoScroll = True

            If (RadioButton1.Checked = True Or RadioButton4.Checked = True Or RadioButton5.Checked = True) And (RadioButton3.Checked = True Or RadioButton2.Checked = True) Then

                If RadioButton3.Checked = True Then

                    For i = 1 To displayRng.Rows.Count
                        For j = 1 To displayRng.Columns.Count
                            Dim label As New System.Windows.Forms.Label
                            label.Text = displayRng.Cells(i, displayRng.Columns.Count - j + 1).Value
                            label.Location = New System.Drawing.Point((j - 1) * width, (i - 1) * height)
                            label.Height = height
                            label.Width = width
                            label.BorderStyle = BorderStyle.FixedSingle
                            label.TextAlign = ContentAlignment.MiddleCenter

                            If CheckBox2.Checked = True Then
                                Dim cell As Excel.Range = displayRng.Cells(i, displayRng.Columns.Count - j + 1)
                                Dim font As Excel.Font = cell.Font

                                Dim fontStyle As FontStyle = FontStyle.Regular
                                If cell.Font.Bold Then fontStyle = fontStyle Or FontStyle.Bold
                                If cell.Font.Italic Then fontStyle = fontStyle Or FontStyle.Italic

                                Dim fontSize As Single = Convert.ToSingle(font.Size)

                                label.Font = New System.Drawing.Font(font.ToString, fontSize, fontStyle)
                                If Not cell.Interior.ColorIndex = Excel.XlColorIndex.xlColorIndexNone Then
                                    Dim colorValue1 As Long = CLng(cell.Interior.Color)
                                    Dim red1 As Integer = colorValue1 Mod 256
                                    Dim green1 As Integer = (colorValue1 \ 256) Mod 256
                                    Dim blue1 As Integer = (colorValue1 \ 256 \ 256) Mod 256
                                    label.BackColor = System.Drawing.Color.FromArgb(red1, green1, blue1)
                                End If
                                If Not cell.Font.ColorIndex = Excel.XlColorIndex.xlColorIndexNone Then
                                    Dim colorValue2 As Long = CLng(cell.Font.Color)
                                    Dim red2 As Integer = colorValue2 Mod 256
                                    Dim green2 As Integer = (colorValue2 \ 256) Mod 256
                                    Dim blue2 As Integer = (colorValue2 \ 256 \ 256) Mod 256
                                    label.ForeColor = System.Drawing.Color.FromArgb(red2, green2, blue2)
                                End If
                            End If

                            panel2.Controls.Add(label)
                        Next
                    Next

                End If


                If RadioButton2.Checked = True Then

                    For i = 1 To displayRng.Rows.Count
                        For j = 1 To displayRng.Columns.Count
                            Dim label As New System.Windows.Forms.Label
                            label.Text = displayRng.Cells(displayRng.Rows.Count - i + 1, j).Value
                            label.Location = New System.Drawing.Point((j - 1) * width, (i - 1) * height)
                            label.Height = height
                            label.Width = width
                            label.BorderStyle = BorderStyle.FixedSingle
                            label.TextAlign = ContentAlignment.MiddleCenter

                            If CheckBox2.Checked = True Then
                                Dim cell As Excel.Range = displayRng.Cells(i, rng.Columns.Count - j + 1)
                                Dim font As Excel.Font = cell.Font

                                Dim fontStyle As FontStyle = FontStyle.Regular
                                If cell.Font.Bold Then fontStyle = fontStyle Or FontStyle.Bold
                                If cell.Font.Italic Then fontStyle = fontStyle Or FontStyle.Italic

                                Dim fontSize As Single = Convert.ToSingle(font.Size)

                                label.Font = New System.Drawing.Font(font.ToString, fontSize, fontStyle)

                                If Not cell.Interior.ColorIndex = Excel.XlColorIndex.xlColorIndexNone Then
                                    Dim colorValue1 As Long = CLng(cell.Interior.Color)
                                    Dim red1 As Integer = colorValue1 Mod 256
                                    Dim green1 As Integer = (colorValue1 \ 256) Mod 256
                                    Dim blue1 As Integer = (colorValue1 \ 256 \ 256) Mod 256
                                    label.BackColor = System.Drawing.Color.FromArgb(red1, green1, blue1)
                                End If
                                If Not cell.Font.ColorIndex = Excel.XlColorIndex.xlColorIndexNone Then
                                    Dim colorValue2 As Long = CLng(cell.Font.Color)
                                    Dim red2 As Integer = colorValue2 Mod 256
                                    Dim green2 As Integer = (colorValue2 \ 256) Mod 256
                                    Dim blue2 As Integer = (colorValue2 \ 256 \ 256) Mod 256
                                    label.ForeColor = System.Drawing.Color.FromArgb(red2, green2, blue2)
                                End If
                            End If

                            panel2.Controls.Add(label)
                        Next
                    Next

                End If

                panel2.AutoScroll = True

            End If

        Catch ex As Exception

        End Try

    End Sub
    Private Sub PictureBox4_Click(sender As Object, e As EventArgs) Handles PictureBox4.Click

        Try
            FocusedTextBox = 1
            Me.Hide()

            excelApp = Globals.ThisAddIn.Application
            workBook = excelApp.ActiveWorkbook

            Dim userInput As Excel.Range = excelApp.InputBox("Select a range", Type:=8)
            rng = userInput


            Dim sheetName As String
            sheetName = Split(rng.Address(True, True, Excel.XlReferenceStyle.xlA1, True), "]")(1)
            sheetName = Split(sheetName, "!")(0)
            workSheet = workBook.Worksheets(sheetName)
            workSheet.Activate()


            rng.Select()

            rng = excelApp.Range(rng, rng.End(Microsoft.Office.Interop.Excel.XlDirection.xlDown))
            rng = excelApp.Range(rng, rng.End(Microsoft.Office.Interop.Excel.XlDirection.xlToRight))

            rng.Select()
            Me.TextBox1.Text = rng.Address

            Me.Show()
            Me.TextBox1.Focus()

        Catch ex As Exception

            Me.Show()
            TextBox1.Focus()

        End Try

    End Sub

    Private Sub PictureBox8_Click(sender As Object, e As EventArgs) Handles PictureBox8.Click

        Try
            FocusedTextBox = 1
            Me.Hide()

            excelApp = Globals.ThisAddIn.Application
            workBook = excelApp.ActiveWorkbook

            Dim userInput As Excel.Range = excelApp.InputBox("Select a range", Type:=8)
            rng = userInput


            Dim sheetName As String
            sheetName = Split(rng.Address(True, True, Excel.XlReferenceStyle.xlA1, True), "]")(1)
            sheetName = Split(sheetName, "!")(0)
            workSheet = workBook.Worksheets(sheetName)
            workSheet.Activate()

            rng.Select()

            TextBox1.Text = rng.Address

            Me.Show()
            TextBox1.Focus()

        Catch ex As Exception

            Me.Show()
            TextBox1.Focus()

        End Try


    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

        Try
            excelApp = Globals.ThisAddIn.Application
            workBook = excelApp.ActiveWorkbook
            workSheet = workBook.ActiveSheet

            rng = workSheet.Range(TextBox1.Text)
            rng.Select()

            Call Display()

        Catch ex As Exception

        End Try

    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        Try
            Call Display()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        Try
            Call Display()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        Try
            Call Display()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub RadioButton4_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton4.CheckedChanged
        Try
            Call Display()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged
        Try
            Call Display()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub btn_OK_Click(sender As Object, e As EventArgs) Handles btn_OK.Click

        Try

            If TextBox1.Text = "" Then
                MessageBox.Show("Select a Source Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                TextBox1.Focus()
                Exit Sub
            End If

            If IsValidExcelCellReference(TextBox1.Text) = False Then
                MessageBox.Show("Select a Valid Source Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                TextBox1.Focus()
                Exit Sub
            End If

            If TextBox2.Text = "" Then
                MessageBox.Show("Select a Destination Cell.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                TextBox2.Focus()
                Exit Sub
            End If

            If IsValidExcelCellReference(TextBox2.Text) = False Then
                MessageBox.Show("Select a Valid Destination Cell.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                TextBox2.Focus()
                Exit Sub
            End If

            If RadioButton2.Checked = False And RadioButton3.Checked = False Then
                MessageBox.Show("Select a Flip Type.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                workSheet.Activate()
                rng.Select()
                Exit Sub

            ElseIf RadioButton1.Checked = False And RadioButton4.Checked = False And RadioButton5.Checked = False Then
                MessageBox.Show("Select a Flip Option.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                workSheet.Activate()
                rng.Select()
                Exit Sub
            End If

            If CheckBox1.Checked = True Then
                workSheet.Copy(After:=workBook.Sheets(workSheet.Name))
                workSheet2.Activate()
            End If

            rng2 = workSheet2.Range(rng2.Cells(1, 1), rng2.Cells(rng.Rows.Count, rng.Columns.Count))

            rng2.Select()


            If (RadioButton1.Checked = True Or RadioButton4.Checked = True Or RadioButton5.Checked = True) And (RadioButton3.Checked = True Or RadioButton2.Checked = True) Then

                If Overlap(excelApp, workSheet, workSheet2, rng, rng2) = False Then

                    If RadioButton3.Checked = True Then

                        For i = 1 To rng.Rows.Count
                            For j = 1 To rng.Columns.Count
                                If RadioButton1.Checked = True Then
                                    rng2.Cells(i, j).Value = rng.Cells(i, rng.Columns.Count - j + 1).Value
                                End If

                                If RadioButton4.Checked = True Then
                                    If rng.Cells(i, rng.Columns.Count - j + 1).HasFormula = True Then
                                        rng2.Cells(i, j).Formula = ReplaceFormula(rng.Cells(i, rng.Columns.Count - j + 1).Formula, rng, rng2, 1, workSheet, workSheet2)
                                    Else
                                        rng2.Cells(i, j).Value = rng.Cells(i, rng.Columns.Count - j + 1).Value
                                    End If
                                End If

                                If RadioButton5.Checked = True Then
                                    If rng.Cells(i, rng.Columns.Count - j + 1).HasFormula = True Then
                                        rng2.Cells(i, j).Formula = rng.Cells(i, rng.Columns.Count - j + 1).Formula
                                    Else
                                        rng2.Cells(i, j).Value = rng.Cells(i, rng.Columns.Count - j + 1).Value
                                    End If
                                End If

                                If CheckBox2.Checked = True Then
                                    rng.Cells(i, rng.Columns.Count - j + 1).Copy
                                    rng2.Cells(i, j).PasteSpecial(Excel.XlPasteType.xlPasteFormats)
                                End If
                            Next
                        Next
                    End If


                    If RadioButton2.Checked = True Then

                        For i = 1 To rng.Rows.Count
                            For j = 1 To rng.Columns.Count

                                If RadioButton1.Checked = True Then
                                    rng2.Cells(i, j).Value = rng.Cells(rng.Rows.Count - i + 1, j).Value

                                End If

                                If RadioButton4.Checked = True Then
                                    If rng.Cells(rng.Rows.Count - i + 1, j).HasFormula = True Then
                                        rng2.Cells(i, j).Formula = ReplaceFormula(rng.Cells(rng.Rows.Count - i + 1, j).Formula, rng, rng2, 1, workSheet, workSheet2)
                                    Else
                                        rng2.Cells(i, j).Value = rng.Cells(rng.Rows.Count - i + 1, j).Value
                                    End If
                                End If

                                If RadioButton5.Checked = True Then
                                    If rng.Cells(rng.Rows.Count - i + 1, j).HasFormula = True Then
                                        rng2.Cells(i, j).Formula = rng.Cells(rng.Rows.Count - i + 1, j).Formula
                                    Else
                                        rng2.Cells(i, j).Value = rng.Cells(rng.Rows.Count - i + 1, j).Value
                                    End If
                                End If

                                If CheckBox2.Checked = True Then
                                    rng.Cells(rng.Rows.Count - i + 1, j).Copy
                                    rng2.Cells(i, j).PasteSpecial(Excel.XlPasteType.xlPasteFormats)
                                End If
                            Next
                        Next
                    End If

                Else

                    Dim Arr(rng.Rows.Count - 1, rng.Columns.Count - 1) As Object

                    For i = LBound(Arr, 1) To UBound(Arr, 1)
                        For j = LBound(Arr, 2) To UBound(Arr, 2)
                            Arr(i, j) = rng.Cells(i + 1, j + 1).Value
                        Next
                    Next

                    Dim FontNames(rng.Rows.Count - 1, rng.Columns.Count - 1) As String
                    Dim HasFormulas(rng.Rows.Count - 1, rng.Columns.Count - 1) As Boolean
                    Dim Formulas(rng.Rows.Count - 1, rng.Columns.Count - 1) As String
                    Dim FontSizes(rng.Rows.Count - 1, rng.Columns.Count - 1) As Single

                    Dim FontBolds(rng.Rows.Count - 1, rng.Columns.Count - 1) As Boolean
                    Dim Fontitalics(rng.Rows.Count - 1, rng.Columns.Count - 1) As Boolean
                    Dim Red1s(rng.Rows.Count - 1, rng.Columns.Count - 1) As Integer
                    Dim Green1s(rng.Rows.Count - 1, rng.Columns.Count - 1) As Integer
                    Dim Blue1s(rng.Rows.Count - 1, rng.Columns.Count - 1) As Integer
                    Dim Red2s(rng.Rows.Count - 1, rng.Columns.Count - 1) As Integer
                    Dim Green2s(rng.Rows.Count - 1, rng.Columns.Count - 1) As Integer
                    Dim Blue2s(rng.Rows.Count - 1, rng.Columns.Count - 1) As Integer

                    For i = LBound(FontSizes, 1) To UBound(FontSizes, 1)
                        For j = LBound(FontSizes, 2) To UBound(FontSizes, 2)

                            Dim cell As Excel.Range = rng.Cells(i + 1, j + 1)
                            If cell.HasFormula Then
                                HasFormulas(i, j) = True
                            Else
                                HasFormulas(i, j) = False
                            End If

                            Formulas(i, j) = cell.Formula
                            Dim font As Excel.Font = cell.Font
                            FontNames(i, j) = font.Name
                            FontBolds(i, j) = cell.Font.Bold
                            Fontitalics(i, j) = cell.Font.Italic


                            Dim fontSize As Single = Convert.ToSingle(font.Size)
                            FontSizes(i, j) = fontSize

                            Dim colorValue1 As Long = CLng(cell.Interior.Color)
                            Dim red1 As Integer = colorValue1 Mod 256
                            Dim green1 As Integer = (colorValue1 \ 256) Mod 256
                            Dim blue1 As Integer = (colorValue1 \ 256 \ 256) Mod 256
                            Red1s(i, j) = red1
                            Green1s(i, j) = green1
                            Blue1s(i, j) = blue1
                            Dim colorValue2 As Long = CLng(cell.Font.Color)
                            Dim red2 As Integer = colorValue2 Mod 256
                            Dim green2 As Integer = (colorValue2 \ 256) Mod 256
                            Dim blue2 As Integer = (colorValue2 \ 256 \ 256) Mod 256
                            Red2s(i, j) = red2
                            Green2s(i, j) = green2
                            Blue2s(i, j) = blue2

                        Next
                    Next
                    If RadioButton3.Checked = True Then
                        For i = 1 To rng.Rows.Count
                            For j = 1 To rng.Columns.Count

                                If RadioButton1.Checked = True Then
                                    rng2.Cells(i, j).Value = Arr(i - 1, rng.Columns.Count - j + 1 - 1)
                                End If

                                If RadioButton4.Checked = True Then
                                    If HasFormulas(i - 1, rng.Columns.Count - j + 1 - 1) = True Then
                                        rng2.Cells(i, j).Formula = ReplaceFormula(Formulas(i - 1, rng.Columns.Count - j + 1 - 1), rng, rng2, 1, workSheet, workSheet2)
                                    Else
                                        rng2.Cells(i, j) = Arr(i - 1, rng.Columns.Count - j + 1 - 1)
                                    End If
                                End If

                                If RadioButton5.Checked = True Then
                                    If HasFormulas(i - 1, rng.Columns.Count - j + 1 - 1) = True Then
                                        rng2.Cells(i, j).Formula = Formulas(i - 1, rng.Columns.Count - j + 1 - 1)
                                    Else
                                        rng2.Cells(i, j) = Arr(i - 1, rng.Columns.Count - j + 1 - 1)
                                    End If
                                End If

                                If CheckBox2.Checked = True Then
                                    Dim x As Integer = i - 1
                                    Dim y As Integer = rng.Columns.Count - j + 1 - 1

                                    rng2.Cells(i, j).Font.Name = FontNames(x, y)
                                    rng2.Cells(i, j).Font.Size = FontSizes(x, y)

                                    If FontBolds(x, y) Then rng2.Cells(i, j).Font.Bold = True
                                    If Fontitalics(x, y) Then rng2.Cells(i, j).Font.Italic = True


                                    rng2.Cells(i, j).Interior.Color = System.Drawing.Color.FromArgb(Red1s(x, y), Green1s(x, y), Blue1s(x, y))

                                    rng2.Cells(i, j).Font.Color = System.Drawing.Color.FromArgb(Red2s(x, y), Green2s(x, y), Blue2s(x, y))

                                End If

                            Next
                        Next

                    End If


                    If RadioButton2.Checked = True Then

                        For i = 1 To rng.Rows.Count
                            For j = 1 To rng.Columns.Count

                                If RadioButton1.Checked = True Then
                                    rng2.Cells(i, j).Value = Arr(rng.Rows.Count - i + 1 - 1, j - 1)
                                End If

                                If RadioButton4.Checked = True Then
                                    If HasFormulas(rng.Rows.Count - i + 1 - 1, j - 1) = True Then
                                        rng2.Cells(i, j).Formula = ReplaceFormula(Formulas(rng.Rows.Count - i + 1 - 1, j - 1), rng, rng2, 2, workSheet, workSheet2)
                                    Else
                                        rng2.Cells(i, j) = Arr(rng.Rows.Count - i + 1 - 1, j - 1)
                                    End If
                                End If

                                If RadioButton5.Checked = True Then
                                    If HasFormulas(rng.Rows.Count - i + 1 - 1, j - 1) = True Then
                                        rng2.Cells(i, j).Formula = Formulas(rng.Rows.Count - i + 1 - 1, j - 1)
                                    Else
                                        rng2.Cells(i, j) = Arr(rng.Rows.Count - i + 1 - 1, j - 1)
                                    End If
                                End If

                                If CheckBox2.Checked = True Then
                                    Dim x As Integer = rng.Rows.Count - i + 1 - 1
                                    Dim y As Integer = j - 1

                                    Dim fontStyle As FontStyle = FontStyle.Regular

                                    If FontBolds(x, y) Then fontStyle = fontStyle Or FontStyle.Bold
                                    If Fontitalics(x, y) Then fontStyle = fontStyle Or FontStyle.Italic


                                    rng2.Cells(i, j).Font.Name = FontNames(x, y)
                                    rng2.Cells(i, j).Font.Size = FontSizes(x, y)

                                    If FontBolds(x, y) Then rng2.Cells(i, j).Font.Bold = True
                                    If Fontitalics(x, y) Then rng2.Cells(i, j).Font.Italic = True


                                    rng2.Cells(i, j).Interior.Color = System.Drawing.Color.FromArgb(Red1s(x, y), Green1s(x, y), Blue1s(x, y))
                                    rng2.Cells(i, j).Font.Color = System.Drawing.Color.FromArgb(Red2s(x, y), Green2s(x, y), Blue2s(x, y))

                                End If

                            Next
                        Next

                    End If
                End If

                Me.Close()

            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub btn_cancel_Click(sender As Object, e As EventArgs) Handles btn_cancel.Click
        Try
            Me.Close()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub PictureBox9_Click(sender As Object, e As EventArgs) Handles PictureBox9.Click

        Try
            FocusedTextBox = 2
            Me.Hide()

            excelApp = Globals.ThisAddIn.Application
            workBook = excelApp.ActiveWorkbook

            Dim userInput As Excel.Range = excelApp.InputBox("Select a range", Type:=8)
            rng2 = userInput


            Dim sheetName As String
            sheetName = Split(rng2.Address(True, True, Excel.XlReferenceStyle.xlA1, True), "]")(1)
            sheetName = Split(sheetName, "!")(0)
            workSheet2 = workBook.Worksheets(sheetName)
            workSheet2.Activate()

            rng2.Select()

            TextBox2.Text = rng2.Address

            Me.Show()
            TextBox2.Focus()

        Catch ex As Exception

            Me.Show()
            TextBox2.Focus()

        End Try

    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

        Try
            excelApp = Globals.ThisAddIn.Application
            workBook = excelApp.ActiveWorkbook
            workSheet2 = workBook.ActiveSheet

            rng2 = workSheet2.Range(TextBox2.Text)
            rng2.Select()

        Catch ex As Exception

        End Try

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Try
            If ComboBox1.SelectedItem = "SOFTEKO" And opened >= 1 Then

                Dim url As String = "https://www.softeko.co"
                Process.Start(url)

            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub Form1_Activated(sender As Object, e As EventArgs) Handles Me.Activated
        Try

            excelApp = Globals.ThisAddIn.Application

            AddHandler excelApp.SheetSelectionChange, AddressOf excelApp_SheetSelectionChange

            opened = opened + 1

        Catch ex As Exception

        End Try

    End Sub

    Private Sub excelApp_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Excel.Range)

        Try

            excelApp = Globals.ThisAddIn.Application
            Dim selectedRange As Excel.Range
            selectedRange = excelApp.Selection
            If FocusedTextBox = 1 Then
                TextBox1.Text = selectedRange.Address
                workSheet = workBook.ActiveSheet
                rng = selectedRange
                TextBox1.Focus()
            ElseIf FocusedTextBox = 2 Then
                TextBox2.Text = selectedRange.Address
                workSheet2 = workBook.ActiveSheet
                rng2 = selectedRange
                TextBox2.Focus()
            End If

        Catch ex As Exception

        End Try

    End Sub


    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown

        Try
            If e.KeyCode = Keys.Enter Then

                Call btn_OK_Click(sender, e)

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub PictureBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles PictureBox4.KeyDown

        Try

            If e.KeyCode = Keys.Enter Then

                Call btn_OK_Click(sender, e)

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub PictureBox8_KeyDown(sender As Object, e As KeyEventArgs) Handles PictureBox8.KeyDown

        Try

            If e.KeyCode = Keys.Enter Then

                Call btn_OK_Click(sender, e)

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub RadioButton3_KeyDown(sender As Object, e As KeyEventArgs) Handles RadioButton3.KeyDown

        Try

            If e.KeyCode = Keys.Enter Then

                Call btn_OK_Click(sender, e)

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub RadioButton2_KeyDown(sender As Object, e As KeyEventArgs) Handles RadioButton2.KeyDown

        Try

            If e.KeyCode = Keys.Enter Then

                Call btn_OK_Click(sender, e)

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub RadioButton1_KeyDown(sender As Object, e As KeyEventArgs) Handles RadioButton1.KeyDown
        Try

            If e.KeyCode = Keys.Enter Then

                Call btn_OK_Click(sender, e)

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub RadioButton4_KeyDown(sender As Object, e As KeyEventArgs) Handles RadioButton4.KeyDown

        Try

            If e.KeyCode = Keys.Enter Then

                Call btn_OK_Click(sender, e)

            End If

        Catch ex As Exception

        End Try

    End Sub


    Private Sub PictureBox5_KeyDown(sender As Object, e As KeyEventArgs) Handles PictureBox5.KeyDown

        Try
            If e.KeyCode = Keys.Enter Then

                Call btn_OK_Click(sender, e)

            End If
        Catch ex As Exception

        End Try

    End Sub


    Private Sub PictureBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles PictureBox1.KeyDown

        If e.KeyCode = Keys.Enter Then

            Call btn_OK_Click(sender, e)

        End If
    End Sub

    Private Sub PictureBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles PictureBox3.KeyDown

        Try

            If e.KeyCode = Keys.Enter Then

                Call btn_OK_Click(sender, e)

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub PictureBox6_KeyDown(sender As Object, e As KeyEventArgs) Handles PictureBox6.KeyDown

        Try
            If e.KeyCode = Keys.Enter Then

                Call btn_OK_Click(sender, e)

            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub CheckBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles CheckBox2.KeyDown

        Try
            If e.KeyCode = Keys.Enter Then

                Call btn_OK_Click(sender, e)

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub PictureBox10_KeyDown(sender As Object, e As KeyEventArgs) Handles PictureBox9.KeyDown

        Try
            If e.KeyCode = Keys.Enter Then

                Call btn_OK_Click(sender, e)

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub PictureBox9_KeyDown(sender As Object, e As KeyEventArgs) Handles PictureBox9.KeyDown

        Try
            If e.KeyCode = Keys.Enter Then

                Call btn_OK_Click(sender, e)

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub CheckBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles CheckBox1.KeyDown

        Try
            If e.KeyCode = Keys.Enter Then

                Call btn_OK_Click(sender, e)

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub ComboBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox1.KeyDown

        Try
            If e.KeyCode = Keys.Enter Then

                Call btn_OK_Click(sender, e)

            End If

        Catch ex As Exception

        End Try

    End Sub


    Private Sub CustomGroupBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles CustomGroupBox1.KeyDown

        Try
            If e.KeyCode = Keys.Enter Then

                Call btn_OK_Click(sender, e)

            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub panel1_KeyDown(sender As Object, e As KeyEventArgs) Handles panel1.KeyDown

        Try
            If e.KeyCode = Keys.Enter Then

                Call btn_OK_Click(sender, e)

            End If
        Catch ex As Exception

        End Try

    End Sub


    Private Sub PictureBox7_KeyDown(sender As Object, e As KeyEventArgs) Handles PictureBox7.KeyDown

        Try
            If e.KeyCode = Keys.Enter Then

                Call btn_OK_Click(sender, e)

            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub CustomGroupBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles CustomGroupBox2.KeyDown

        Try
            If e.KeyCode = Keys.Enter Then

                Call btn_OK_Click(sender, e)

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub panel2_KeyDown(sender As Object, e As KeyEventArgs) Handles panel2.KeyDown

        Try
            If e.KeyCode = Keys.Enter Then

                Call btn_OK_Click(sender, e)

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub btn_OK_KeyDown(sender As Object, e As KeyEventArgs) Handles btn_OK.KeyDown

        Try
            If e.KeyCode = Keys.Enter Then

                Call btn_OK_Click(sender, e)

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub btn_cancel_KeyDown(sender As Object, e As KeyEventArgs) Handles btn_cancel.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then

                Call btn_OK_Click(sender, e)

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub Form1_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown

        Try
            If e.KeyCode = Keys.Enter Then

                Call btn_OK_Click(sender, e)

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub TextBox1_GotFocus(sender As Object, e As EventArgs) Handles TextBox1.GotFocus

        Try
            FocusedTextBox = 1

        Catch ex As Exception

        End Try

    End Sub

    Private Sub TextBox2_GotFocus(sender As Object, e As EventArgs) Handles TextBox2.GotFocus
        Try
            FocusedTextBox = 2

        Catch ex As Exception

        End Try
    End Sub

    Private Sub PictureBox4_GotFocus(sender As Object, e As EventArgs) Handles PictureBox4.GotFocus
        Try
            FocusedTextBox = 1
        Catch ex As Exception

        End Try
    End Sub

    Private Sub PictureBox8_GotFocus(sender As Object, e As EventArgs) Handles PictureBox8.GotFocus
        Try
            FocusedTextBox = 1

        Catch ex As Exception

        End Try
    End Sub

    Private Sub PictureBox10_GotFocus(sender As Object, e As EventArgs) Handles PictureBox9.GotFocus
        Try
            FocusedTextBox = 2
        Catch ex As Exception

        End Try
    End Sub

    Private Sub PictureBox9_GotFocus(sender As Object, e As EventArgs) Handles PictureBox9.GotFocus
        Try
            FocusedTextBox = 2

        Catch ex As Exception

        End Try
    End Sub


    Private Sub btn_OK_MouseHover(sender As Object, e As EventArgs) Handles btn_OK.MouseHover

        Try

            btn_OK.BackColor = Color.FromArgb(65, 105, 225)
            btn_OK.ForeColor = Color.FromArgb(255, 255, 255)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub btn_OK_MouseLeave(sender As Object, e As EventArgs) Handles btn_OK.MouseLeave
        Try

            btn_OK.BackColor = Color.FromArgb(255, 255, 255)
            btn_OK.ForeColor = Color.FromArgb(70, 70, 70)
        Catch ex As Exception

        End Try

    End Sub

    Private Sub btn_cancel_MouseHover(sender As Object, e As EventArgs) Handles btn_cancel.MouseHover

        Try
            btn_cancel.BackColor = Color.FromArgb(65, 105, 225)
            btn_cancel.ForeColor = Color.FromArgb(255, 255, 255)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub btn_cancel_MouseLeave(sender As Object, e As EventArgs) Handles btn_cancel.MouseLeave

        Try
            btn_cancel.BackColor = Color.FromArgb(255, 255, 255)
            btn_cancel.ForeColor = Color.FromArgb(70, 70, 70)

        Catch ex As Exception

        End Try
    End Sub

    Private Sub RadioButton2_GotFocus(sender As Object, e As EventArgs) Handles RadioButton2.GotFocus
        Try
            FocusedTextBox = 0
        Catch ex As Exception

        End Try
    End Sub

    Private Sub RadioButton1_GotFocus(sender As Object, e As EventArgs) Handles RadioButton1.GotFocus
        Try
            FocusedTextBox = 0
        Catch ex As Exception

        End Try
    End Sub

    Private Sub RadioButton4_GotFocus(sender As Object, e As EventArgs) Handles RadioButton4.GotFocus
        Try
            FocusedTextBox = 0
        Catch ex As Exception

        End Try
    End Sub

    Private Sub RadioButton5_GotFocus(sender As Object, e As EventArgs) Handles RadioButton5.GotFocus
        Try
            FocusedTextBox = 0
        Catch ex As Exception

        End Try
    End Sub

    Private Sub PictureBox5_GotFocus(sender As Object, e As EventArgs) Handles PictureBox5.GotFocus
        Try
            FocusedTextBox = 0
        Catch ex As Exception

        End Try
    End Sub


    Private Sub PictureBox1_GotFocus(sender As Object, e As EventArgs) Handles PictureBox1.GotFocus
        Try
            FocusedTextBox = 0
        Catch ex As Exception

        End Try
    End Sub


    Private Sub PictureBox3_GotFocus(sender As Object, e As EventArgs) Handles PictureBox3.GotFocus
        Try
            FocusedTextBox = 0
        Catch ex As Exception

        End Try
    End Sub

    Private Sub PictureBox6_GotFocus(sender As Object, e As EventArgs) Handles PictureBox6.GotFocus
        Try
            FocusedTextBox = 0
        Catch ex As Exception

        End Try
    End Sub

    Private Sub RadioButton3_GotFocus(sender As Object, e As EventArgs) Handles RadioButton3.GotFocus
        Try
            FocusedTextBox = 0
        Catch ex As Exception

        End Try
    End Sub


    Private Sub PictureBox2_GotFocus(sender As Object, e As EventArgs) Handles PictureBox2.GotFocus
        Try
            FocusedTextBox = 0
        Catch ex As Exception

        End Try
    End Sub

    Private Sub CheckBox1_GotFocus(sender As Object, e As EventArgs) Handles CheckBox1.GotFocus
        Try
            FocusedTextBox = 0
        Catch ex As Exception

        End Try
    End Sub

    Private Sub ComboBox1_GotFocus(sender As Object, e As EventArgs) Handles ComboBox1.GotFocus
        Try
            FocusedTextBox = 0
        Catch ex As Exception

        End Try
    End Sub

    Private Sub CustomGroupBox1_GotFocus(sender As Object, e As EventArgs) Handles CustomGroupBox1.GotFocus
        Try
            FocusedTextBox = 0
        Catch ex As Exception

        End Try
    End Sub

    Private Sub CustomGroupBox2_GotFocus(sender As Object, e As EventArgs) Handles CustomGroupBox2.GotFocus
        Try
            FocusedTextBox = 0
        Catch ex As Exception

        End Try
    End Sub

    Private Sub panel1_GotFocus(sender As Object, e As EventArgs) Handles panel1.GotFocus
        Try
            FocusedTextBox = 0
        Catch ex As Exception

        End Try
    End Sub

    Private Sub panel2_GotFocus(sender As Object, e As EventArgs) Handles panel2.GotFocus
        Try
            FocusedTextBox = 0
        Catch ex As Exception

        End Try
    End Sub

    Private Sub PictureBox7_GotFocus(sender As Object, e As EventArgs) Handles PictureBox7.GotFocus
        Try
            FocusedTextBox = 0
        Catch ex As Exception

        End Try
    End Sub

    Private Sub btn_OK_GotFocus(sender As Object, e As EventArgs) Handles btn_OK.GotFocus
        Try
            FocusedTextBox = 0
        Catch ex As Exception

        End Try
    End Sub

    Private Sub btn_cancel_GotFocus(sender As Object, e As EventArgs) Handles btn_cancel.GotFocus
        Try
            FocusedTextBox = 0
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub
End Class
