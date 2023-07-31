﻿Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Runtime.InteropServices
Imports System.Drawing
Imports Microsoft.Office.Interop.Excel
Imports System.Net.Mime.MediaTypeNames
Imports System.Reflection
Imports System.Windows.Forms
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports System.Diagnostics
Imports System.Text.RegularExpressions

Public Class Form8
    Dim WithEvents excelApp As Excel.Application

    Dim workBook As Excel.Workbook
    Dim workbook2 As Excel.Workbook

    Dim workSheet As Excel.Worksheet
    Dim workSheet2 As Excel.Worksheet

    Dim rng As Excel.Range
    Dim rng2 As Excel.Range

    Dim opened As Integer
    Dim FocusedTextBox As Integer
    Function FindMinValue(arr() As Integer) As Integer

        Dim min As Integer = arr(0)

        For Each num In arr
            If num < min Then
                min = num
            End If
        Next

        Return min

    End Function
    Private Function SearchAlongRow(Rng As Excel.Range, r As Integer, C As Integer)

        Dim i As Integer = 1

        Dim search As Boolean = True

        Dim Type1 As Type
        Dim Type2 As Type

        While search = True

            If Rng.Cells(r, C + i).Value Is Nothing Then
                Type1 = GetType(String)
            Else
                Type1 = Rng.Cells(r, C + i).Value.GetType()
            End If

            If Rng.Cells(r, C).Value Is Nothing Then
                Type2 = GetType(String)
            Else
                Type2 = Rng.Cells(r, C).Value.GetType()
            End If

            If Type1.Equals(Type2) Then
                If Rng.Cells(r, C + i).Value = Rng.Cells(r, C).value And (C + i) <= Rng.Columns.Count And Rng.Cells(r, C).MergeCells = False And Rng.Cells(r, C + i).MergeCells = False Then
                    i = i + 1
                    search = True
                Else
                    search = False
                End If
            Else
                search = False
            End If

        End While

        SearchAlongRow = i

    End Function
    Private Function SearchAlongColumn(Rng As Excel.Range, r As Integer, C As Integer)

        Dim i As Integer = 1

        Dim search As Boolean = True

        Dim Type1 As Type
        Dim Type2 As Type

        While search = True

            If Rng.Cells(r + i, C).Value Is Nothing Then
                Type1 = GetType(String)
            Else
                Type1 = Rng.Cells(r + i, C).Value.GetType()
            End If

            If Rng.Cells(r, C).Value Is Nothing Then
                Type2 = GetType(String)
            Else
                Type2 = Rng.Cells(r, C).Value.GetType()
            End If

            If Type1.Equals(Type2) Then
                If Rng.Cells(r + i, C).value = Rng.Cells(r, C).value And (r + i) <= Rng.Rows.Count And Rng.Cells(r, C).MergeCells = False And Rng.Cells(r + i, C).MergeCells = False Then
                    i = i + 1
                    search = True
                Else
                    search = False
                End If
            Else
                search = False
            End If

        End While

        SearchAlongColumn = i

    End Function

    Private Function FindInArray(i, j, Arr)

        Dim Result As Boolean = False
        Dim count As Integer

        For count = LBound(Arr, 1) To UBound(Arr, 1)
            If Arr(count, 0) = i And Arr(count, 1) = j Then
                Result = True
                Exit For
            End If
        Next count

        FindInArray = Result

    End Function
    Private Function SearchAlongColumn2(Rng, r, C, Arr)

        Dim i As Integer = 1

        Dim search As Boolean = True

        Dim Type1 As Type
        Dim Type2 As Type

        While search = True

            If Rng.Cells(r + i, C).Value Is Nothing Then
                Type1 = GetType(String)
            Else
                Type1 = Rng.Cells(r + i, C).Value.GetType()
            End If

            If Rng.Cells(r, C).Value Is Nothing Then
                Type2 = GetType(String)
            Else
                Type2 = Rng.Cells(r, C).Value.GetType()
            End If

            If Type1.Equals(Type2) Then
                If Rng.Cells(r + i, C).value = Rng.Cells(r, C).value And (r + i) <= Rng.Rows.Count And FindInArray(r, C, Arr) = False And FindInArray(r + i, C, Arr) = False And Rng.Cells(r, C).MergeCells = False And Rng.Cells(r + i, C).MergeCells = False Then
                    i = i + 1
                    search = True
                Else
                    search = False
                End If
            Else
                search = False
            End If

        End While

        SearchAlongColumn2 = i

    End Function
    Private Function SearchAlongRow2(Rng, r, C, Arr)

        Dim i As Integer = 1

        Dim search As Boolean = True

        Dim Type1 As Type
        Dim Type2 As Type

        While search = True

            If Rng.Cells(r, C + i).Value Is Nothing Then
                Type1 = GetType(String)
            Else
                Type1 = Rng.Cells(r, C + i).Value.GetType()
            End If

            If Rng.Cells(r, C).Value Is Nothing Then
                Type2 = GetType(String)
            Else
                Type2 = Rng.Cells(r, C).Value.GetType()
            End If

            If Type1.Equals(Type2) Then

                If Rng.Cells(r, C + i).value = Rng.Cells(r, C).value And (C + i) <= Rng.Columns.Count And FindInArray(r, C, Arr) = False And FindInArray(r, C + i, Arr) = False And Rng.Cells(r, C).MergeCells = False And Rng.Cells(r, C + i).MergeCells = False Then
                    i = i + 1
                    search = True
                Else
                    search = False
                End If
            Else
                search = False
            End If

        End While

        SearchAlongRow2 = i

    End Function

    Private Function SearchDiagonally(Rng, r, c, Arr)

        Dim rowEqual As Integer = SearchAlongRow2(Rng, r, c, Arr)

        Dim activesheet As Excel.Worksheet = CType(excelApp.ActiveSheet, Excel.Worksheet)

        Dim Rng2 As Excel.Range
        Rng2 = activesheet.Range(Rng.Cells(1, 1), Rng.Cells(1, rowEqual))

        Dim Output(1) As Integer
        Output(0) = 1
        Output(1) = rowEqual

        Dim TotalCells As Integer = Rng2.Cells.Count

        Dim j As Integer

        j = 0

        While SearchAlongColumn2(Rng, r, c + j, Arr) > 1 And j + 1 <= rowEqual
            If activesheet.Range(Rng.Cells(1, 1), Rng.Cells(SearchAlongColumn2(Rng, r, c + j, Arr), j + 1)).Cells.Count >= TotalCells Then
                Rng2 = activesheet.Range(Rng.Cells(1, 1), Rng.Cells(SearchAlongColumn2(Rng, r, c + j, Arr), j + 1))
                Output(0) = SearchAlongColumn2(Rng, r, c + j, Arr)
                Output(1) = j + 1
                TotalCells = Rng2.Cells.Count
            End If
            j = j + 1

        End While

        SearchDiagonally = Output

    End Function

    Private Sub Display()

        CustomPanel1.Controls.Clear()
        CustomPanel2.Controls.Clear()

        excelApp = Globals.ThisAddIn.Application
        workBook = excelApp.ActiveWorkbook
        workSheet = workBook.ActiveSheet


        Dim Rng As Excel.Range = workSheet.Range(TextBox1.Text)
            Dim displayRng As Excel.Range
            Rng.Select()

            If Rng.Rows.Count > 50 Then
                displayRng = workSheet.Range(Rng.Cells(1, 1), Rng.Cells(50, Rng.Columns.Count))
            Else
                displayRng = workSheet.Range(Rng.Cells(1, 1), Rng.Cells(Rng.Rows.Count, Rng.Columns.Count))
            End If

            Dim r As Integer = displayRng.Rows.Count
            Dim C As Integer = displayRng.Columns.Count

            Dim height As Single
            Dim width As Single

            If r <= 6 Then
                height = CustomPanel1.Height / r
            Else
                height = CustomPanel1.Height / 6
            End If

            If C <= 4 Then
                width = CustomPanel1.Width / C
            Else
                width = CustomPanel1.Width / 4
            End If

        Dim i As Integer
        Dim j As Integer

        For i = 1 To r
            For j = 1 To C
                Dim label As New System.Windows.Forms.Label
                label.Text = displayRng.Cells(i, j).Value
                label.Location = New System.Drawing.Point((j - 1) * width, (i - 1) * height)
                label.Height = height
                label.Width = width
                label.BorderStyle = BorderStyle.FixedSingle
                label.TextAlign = ContentAlignment.MiddleCenter

                If CheckBox1.Checked = True Then
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
                CustomPanel1.Controls.Add(label)
            Next
        Next

        CustomPanel1.AutoScroll = True

        If RadioButton1.Checked = True Or RadioButton2.Checked = True Or RadioButton3.Checked = True Then

            If RadioButton1.Checked = True Then
                For i = 1 To r
                    For j = 1 To C
                        Dim rowEqual As Integer = SearchAlongRow(displayRng, i, j)
                        Dim newWidth As Single = width * rowEqual
                        Dim label As New System.Windows.Forms.Label
                        label.Text = displayRng.Cells(i, j).Value
                        label.Location = New System.Drawing.Point((j - 1) * width, (i - 1) * height)
                        label.Height = height
                        label.Width = newWidth
                        label.BorderStyle = BorderStyle.FixedSingle
                        label.TextAlign = ContentAlignment.MiddleCenter


                        If CheckBox1.Checked = True Then
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

                        j = j + rowEqual - 1

                        CustomPanel2.Controls.Add(label)
                    Next
                Next

            ElseIf RadioButton2.Checked = True = True Then

                For j = 1 To C
                    For i = 1 To r
                        Dim columnEqual As Integer = SearchAlongColumn(displayRng, i, j)
                        Dim newHeight As Single = height * columnEqual
                        Dim label As New System.Windows.Forms.Label
                        label.Text = displayRng.Cells(i, j).Value
                        label.Location = New System.Drawing.Point((j - 1) * width, (i - 1) * height)
                        label.Height = newHeight
                        label.Width = width
                        label.BorderStyle = BorderStyle.FixedSingle
                        label.TextAlign = ContentAlignment.MiddleCenter


                        If CheckBox1.Checked = True Then
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
                        i = i + columnEqual - 1
                        CustomPanel2.Controls.Add(label)
                    Next
                Next

            ElseIf RadioButton3.Checked = True Then

                Dim Arr(r * C - 1, 1) As Object
                Dim count As Integer = 0

                For i = 1 To r
                    For j = 1 To C

                        If FindInArray(i, j, Arr) = False Then
                            Dim rowEqual As Integer = SearchDiagonally(displayRng, i, j, Arr)(0)
                            Dim columnEqual As Integer = SearchDiagonally(displayRng, i, j, Arr)(1)

                            If rowEqual > 1 Or columnEqual > 1 Then
                                For m = 1 To rowEqual
                                    For n = 1 To columnEqual

                                        Arr(count, 0) = i + m - 1
                                        Arr(count, 1) = j + n - 1
                                        count = count + 1

                                    Next
                                Next
                            End If

                            Dim newWidth As Single = width * columnEqual
                            Dim newHeight = height * rowEqual

                            Dim label As New System.Windows.Forms.Label
                            label.Text = displayRng.Cells(i, j).Value
                            label.Location = New System.Drawing.Point((j - 1) * width, (i - 1) * height)
                            label.Height = newHeight
                            label.Width = newWidth
                            label.BorderStyle = BorderStyle.FixedSingle
                            label.TextAlign = ContentAlignment.MiddleCenter

                            If CheckBox1.Checked = True Then
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
                            CustomPanel2.Controls.Add(label)
                        End If
                    Next
                Next

            End If

            CustomPanel2.AutoScroll = True

        End If

    End Sub

    Private Sub Form8_Load(sender As Object, e As EventArgs) Handles Me.Load

        Try
            excelApp = Globals.ThisAddIn.Application
            workBook = excelApp.ActiveWorkbook
            workbook2 = excelApp.ActiveWorkbook
            workSheet = workBook.ActiveSheet
            workSheet2 = workbook2.ActiveSheet

            AddHandler excelApp.SheetSelectionChange, AddressOf excelApp_SheetSelectionChange

            opened = opened + 1

            Me.Label2.Enabled = False
            Me.TextBox3.Enabled = False
            Me.PictureBox6.Enabled = False

        Catch ex As Exception

        End Try

    End Sub

    Private Sub excelApp_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Excel.Range)

        Try

            Dim selectedRange As Excel.Range
            selectedRange = excelApp.Selection

            If FocusedTextBox = 1 Then
                TextBox1.Text = selectedRange.Address
                workSheet = workBook.ActiveSheet
                rng = selectedRange
                TextBox1.Focus()

            ElseIf FocusedTextBox = 3 Then
                TextBox3.Text = selectedRange.Address
                workSheet2 = workbook2.ActiveSheet
                rng2 = selectedRange
                TextBox3.Focus()
            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        Try
            excelApp = Globals.ThisAddIn.Application
            workBook = excelApp.ActiveWorkbook
            workSheet = workBook.ActiveSheet

            rng = workSheet.Range(TextBox1.Text)
            rng.Select()

        Catch ex As Exception

        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Call Display()

    End Sub

    Private Sub PictureBox6_Click(sender As Object, e As EventArgs) Handles PictureBox6.Click

    End Sub
End Class