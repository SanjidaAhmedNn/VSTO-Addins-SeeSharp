Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Runtime.InteropServices
Imports System.Drawing
Imports Microsoft.Office.Interop.Excel
Imports System.Net.Mime.MediaTypeNames
Imports System.Reflection
Imports System.Windows.Forms
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports System.Diagnostics

Public Class Form1
    Dim WithEvents excelApp As Excel.Application
    Dim workBook As Excel.Workbook
    Dim workSheet As Excel.Worksheet
    Dim workSheet2 As Excel.Worksheet
    Dim rng As Excel.Range
    Dim rng2 As Excel.Range
    Dim selectedRange As Excel.Range

    Dim opened As Integer
    Dim FocuesdTextBox As Integer

    Public Function ReplaceFormula(Formula As String, Rng As Excel.Range, rng2 As Excel.Range, Type As Integer)

        Dim activesheet As Excel.Worksheet = CType(excelApp.ActiveSheet, Excel.Worksheet)

        Dim Starters As String() = New String() {"=", "(", ",", ":", " ", "+", "-", "*", "/", "^", ")"}

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

        For i = LBound(Arr) To UBound(Arr)
            If Arr(i) <> "" Then
                If Asc(Mid(Arr(i), Len(Arr(i)), 1)) >= 48 And Asc(Mid(Arr(i), Len(Arr(i)), 1)) <= 57 Then
                    Index = Index + 1
                    ReDim Preserve Refs(Index)
                    Refs(Index) = Arr(i)
                End If
            End If
        Next i

        Dim Work As Boolean
        Dim SheetName As String
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

        distance1 = rng2.Cells(1, 1).Row - Rng.Cells(1, 1).Row
        distance2 = rng2.Cells(1, 1).Column - Rng.Cells(1, 1).Column

        For Each Ref In Refs
            Work = True
            If InStr(1, Ref, "!") > 0 Then
                SheetName = Split(Ref, "!")(0)
                If SheetName = activesheet.Name Then
                    Ref = Split(Ref, "!")(0)
                    Work = True
                Else
                    Work = False
                End If
            End If

            If InStr(1, Ref, ":") > 0 Then
                Dim FirstCell As String
                Dim LastCell As String
                FirstCell = Split(Ref, ":")(0)
                LastCell = Split(Ref, ":")(1)

            End If

            If Work = True Then
                expRange = activesheet.Range(Ref)
                If Type = 1 Then
                    colNum = expRange.Column
                    If colNum >= Rng.Cells(1, 1).Column And colNum <= Rng.Cells(1, Rng.Columns.Count).Column Then
                        colName = Split(activesheet.Cells(1, colNum).Address, "$")(1)
                        Ext = colNum - Rng.Cells(1, 1).Column + 1
                        Ext2 = Rng.Columns.Count - Ext + 1
                        colNum2 = Rng.Cells(1, 1).Column - 1 + Ext2
                        colName2 = Split(activesheet.Cells(1, colNum2).Address, "$")(1)
                        Ref2 = Replace(Ref, colName, colName2)
                        Formula = Replace(Formula, Ref, Ref2)
                    End If
                ElseIf Type = 2 Then
                    rowNum = expRange.Row
                    If rowNum >= Rng.Cells(1, 1).Row And rowNum <= Rng.Cells(Rng.Rows.Count, 1).Row Then
                        rowName = Split(activesheet.Cells(rowNum, 1).Address, "$")(2)
                        Ext = rowNum - Rng.Cells(1, 1).Row + 1
                        Ext2 = Rng.Rows.Count - Ext + 1
                        rowNum2 = Rng.Cells(1, 1).Row - 1 + Ext2
                        rowName2 = Split(activesheet.Cells(rowNum2, 1).Address, "$")(2)
                        Ref2 = Replace(Ref, rowName, rowName2)
                        Formula = Replace(Formula, Ref, Ref2)
                    End If
                End If
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
                Formula = Replace(Formula, Ref2, Ref3)
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

            If (RadioButton1.Checked = True Or RadioButton4.Checked = True) And (RadioButton3.Checked = True Or RadioButton2.Checked = True) Then

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

        FocuesdTextBox = 1
        Try

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
        FocuesdTextBox = 1
        Try

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

        Call Display()

    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged

        Call Display()

    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged

        Call Display()

    End Sub

    Private Sub RadioButton4_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton4.CheckedChanged

        Call Display()

    End Sub

    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged

        Call Display()

    End Sub

    Private Sub btn_OK_Click(sender As Object, e As EventArgs) Handles btn_OK.Click

        Try


            If TextBox1.Text = "" Then
                MessageBox.Show("Select a Source Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                TextBox1.Focus()
                Exit Sub
            End If

            If TextBox2.Text = "" Then
                MessageBox.Show("Select a Destination Cell.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                TextBox2.Focus()
                Exit Sub
            End If

            If RadioButton2.Checked = False And RadioButton3.Checked = False Then
                MessageBox.Show("Select a Flip Type.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                workSheet.Activate()
                rng.Select()
                Exit Sub
            ElseIf RadioButton1.Checked = False And RadioButton4.Checked = False Then
                MessageBox.Show("Select a Flip Option.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                workSheet.Activate()
                rng.Select()
                Exit Sub
            End If

            If CheckBox1.Checked = True Then
                workSheet.Copy(After:=workBook.Sheets(workSheet.Name))
                workSheet2.Activate()
            End If

            Dim Arr(rng.Rows.Count - 1, rng.Columns.Count - 1) As Object

            rng2 = workSheet2.Range(rng2.Cells(1, 1), rng2.Cells(rng.Rows.Count, rng.Columns.Count))

            rng2.Select()

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


            If (RadioButton1.Checked = True Or RadioButton4.Checked = True) And (RadioButton3.Checked = True Or RadioButton2.Checked = True) Then
                If RadioButton3.Checked = True Then
                    For i = 1 To rng.Rows.Count
                        For j = 1 To rng.Columns.Count
                            If RadioButton1.Checked = True Then
                                rng2.Cells(i, j).Value = Arr(i - 1, rng.Columns.Count - j + 1 - 1)
                            End If

                            If RadioButton4.Checked = True Then
                                If HasFormulas(i - 1, rng.Columns.Count - j + 1 - 1) = True Then
                                    rng2.Cells(i, j).Formula = ReplaceFormula(Formulas(i - 1, rng.Columns.Count - j + 1 - 1), rng, rng2, 1)
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
                                    rng2.Cells(i, j).Formula = ReplaceFormula(Formulas(rng.Rows.Count - i + 1 - 1, j - 1), rng, rng2, 2)
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

                Me.Close()

            Else

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub btn_cancel_Click(sender As Object, e As EventArgs) Handles btn_cancel.Click

        Me.Close()

    End Sub


    'Private Sub PictureBox9_Click(sender As Object, e As EventArgs) Handles PictureBox9.Click
    'FocuesdTextBox = 2
    'Try

    'Me.Hide()

    'excelApp = Globals.ThisAddIn.Application
    'workBook = excelApp.ActiveWorkbook

    'Dim userInput As Excel.Range = excelApp.InputBox("Select a range", Type:=8)
    'rng2 = userInput


    'Dim sheetName As String
    'sheetName = Split(rng2.Address(True, True, Excel.XlReferenceStyle.xlA1, True), "]")(1)
    'sheetName = Split(sheetName, "!")(0)
    'workSheet2 = workBook.Worksheets(sheetName)
    'workSheet2.Activate()

    'rng2.Select()

    'TextBox2.Text = rng2.Address

    'Me.Show()
    'TextBox2.Focus()

    'Catch ex As Exception

    'Me.Show()
    'TextBox2.Focus()

    'End Try

    'End Sub

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

        If ComboBox1.SelectedItem = "SOFTEKO" And opened >= 1 Then

            Dim url As String = "https://www.softeko.co"
            Process.Start(url)

        End If
    End Sub

    Private Sub Form1_Activated(sender As Object, e As EventArgs) Handles Me.Activated

        excelApp = Globals.ThisAddIn.Application

        AddHandler excelApp.SheetSelectionChange, AddressOf excelApp_SheetSelectionChange

        opened = opened + 1

    End Sub

    Private Sub excelApp_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Excel.Range)

        excelApp = Globals.ThisAddIn.Application
        Dim selectedRange As Excel.Range
        selectedRange = excelApp.Selection
        If FocuesdTextBox = 1 Then
            TextBox1.Text = selectedRange.Address
            workSheet = workBook.ActiveSheet
            rng = selectedRange
            TextBox1.Focus()
        ElseIf FocuesdTextBox = 2 Then
            TextBox2.Text = selectedRange.Address
            workSheet2 = workBook.ActiveSheet
            rng2 = selectedRange
            TextBox2.Focus()
        End If

    End Sub


    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown

        If e.KeyCode = Keys.Enter Then

            Call btn_OK_Click(sender, e)

        End If

    End Sub

    Private Sub PictureBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles PictureBox4.KeyDown

        If e.KeyCode = Keys.Enter Then

            Call btn_OK_Click(sender, e)

        End If

    End Sub

    Private Sub PictureBox8_KeyDown(sender As Object, e As KeyEventArgs) Handles PictureBox8.KeyDown

        If e.KeyCode = Keys.Enter Then

            Call btn_OK_Click(sender, e)

        End If

    End Sub

    Private Sub RadioButton3_KeyDown(sender As Object, e As KeyEventArgs) Handles RadioButton3.KeyDown

        If e.KeyCode = Keys.Enter Then

            Call btn_OK_Click(sender, e)

        End If

    End Sub

    Private Sub RadioButton2_KeyDown(sender As Object, e As KeyEventArgs) Handles RadioButton2.KeyDown

        If e.KeyCode = Keys.Enter Then

            Call btn_OK_Click(sender, e)

        End If

    End Sub

    Private Sub RadioButton1_KeyDown(sender As Object, e As KeyEventArgs) Handles RadioButton1.KeyDown

        If e.KeyCode = Keys.Enter Then

            Call btn_OK_Click(sender, e)

        End If

    End Sub

    Private Sub RadioButton4_KeyDown(sender As Object, e As KeyEventArgs) Handles RadioButton4.KeyDown

        If e.KeyCode = Keys.Enter Then

            Call btn_OK_Click(sender, e)

        End If

    End Sub


    Private Sub PictureBox5_KeyDown(sender As Object, e As KeyEventArgs) Handles PictureBox5.KeyDown

        If e.KeyCode = Keys.Enter Then

            Call btn_OK_Click(sender, e)

        End If

    End Sub


    Private Sub PictureBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles PictureBox1.KeyDown

        If e.KeyCode = Keys.Enter Then

            Call btn_OK_Click(sender, e)

        End If
    End Sub

    Private Sub PictureBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles PictureBox3.KeyDown

        If e.KeyCode = Keys.Enter Then

            Call btn_OK_Click(sender, e)

        End If

    End Sub

    Private Sub PictureBox6_KeyDown(sender As Object, e As KeyEventArgs) Handles PictureBox6.KeyDown

        If e.KeyCode = Keys.Enter Then

            Call btn_OK_Click(sender, e)

        End If

    End Sub

    Private Sub CheckBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles CheckBox2.KeyDown

        If e.KeyCode = Keys.Enter Then

            Call btn_OK_Click(sender, e)

        End If

    End Sub

    'Private Sub PictureBox10_KeyDown(sender As Object, e As KeyEventArgs) Handles PictureBox10.KeyDown

    'If e.KeyCode = Keys.Enter Then

    'Call btn_OK_Click(sender, e)

    'End If

    'End Sub

    Private Sub PictureBox9_KeyDown(sender As Object, e As KeyEventArgs) Handles PictureBox9.KeyDown

        If e.KeyCode = Keys.Enter Then

            Call btn_OK_Click(sender, e)

        End If

    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged

    End Sub

    Private Sub CheckBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles CheckBox1.KeyDown

        If e.KeyCode = Keys.Enter Then

            Call btn_OK_Click(sender, e)

        End If

    End Sub

    Private Sub ComboBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox1.KeyDown

        If e.KeyCode = Keys.Enter Then

            Call btn_OK_Click(sender, e)

        End If

    End Sub


    Private Sub CustomGroupBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles CustomGroupBox1.KeyDown

        If e.KeyCode = Keys.Enter Then

            Call btn_OK_Click(sender, e)

        End If
    End Sub

    Private Sub panel1_KeyDown(sender As Object, e As KeyEventArgs) Handles panel1.KeyDown

        If e.KeyCode = Keys.Enter Then

            Call btn_OK_Click(sender, e)

        End If

    End Sub


    Private Sub PictureBox7_KeyDown(sender As Object, e As KeyEventArgs) Handles PictureBox7.KeyDown

        If e.KeyCode = Keys.Enter Then

            Call btn_OK_Click(sender, e)

        End If

    End Sub

    Private Sub CustomGroupBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles CustomGroupBox2.KeyDown

        If e.KeyCode = Keys.Enter Then

            Call btn_OK_Click(sender, e)

        End If

    End Sub

    Private Sub panel2_KeyDown(sender As Object, e As KeyEventArgs) Handles panel2.KeyDown

        If e.KeyCode = Keys.Enter Then

            Call btn_OK_Click(sender, e)

        End If

    End Sub

    Private Sub btn_OK_KeyDown(sender As Object, e As KeyEventArgs) Handles btn_OK.KeyDown

        If e.KeyCode = Keys.Enter Then

            Call btn_OK_Click(sender, e)

        End If

    End Sub

    Private Sub btn_cancel_KeyDown(sender As Object, e As KeyEventArgs) Handles btn_cancel.KeyDown

        If e.KeyCode = Keys.Enter Then

            Call btn_OK_Click(sender, e)

        End If

    End Sub

    Private Sub Form1_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown

        If e.KeyCode = Keys.Enter Then

            Call btn_OK_Click(sender, e)

        End If

    End Sub

    Private Sub TextBox1_GotFocus(sender As Object, e As EventArgs) Handles TextBox1.GotFocus

        FocuesdTextBox = 1

    End Sub

    Private Sub TextBox2_GotFocus(sender As Object, e As EventArgs) Handles TextBox2.GotFocus
        FocuesdTextBox = 2
    End Sub

    Private Sub PictureBox4_GotFocus(sender As Object, e As EventArgs) Handles PictureBox4.GotFocus
        FocuesdTextBox = 1
    End Sub

    Private Sub PictureBox8_GotFocus(sender As Object, e As EventArgs) Handles PictureBox8.GotFocus
        FocuesdTextBox = 1
    End Sub

    'Private Sub PictureBox10_GotFocus(sender As Object, e As EventArgs) Handles PictureBox10.GotFocus
    'FocuesdTextBox = 2
    'End Sub

    Private Sub PictureBox9_GotFocus(sender As Object, e As EventArgs) Handles PictureBox9.GotFocus
        FocuesdTextBox = 2
    End Sub

End Class
