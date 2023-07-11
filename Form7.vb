﻿Imports System.Drawing
Imports System.Windows.Forms
Imports System.Reflection.Emit
Imports Microsoft.Office.Interop.Excel
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Threading
Imports System.Diagnostics
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Button
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports System.CodeDom
Imports Microsoft.Office.Core
Imports System.Data

Public Class Form7

    Dim excelApp As Excel.Application
    Dim workbook As Excel.Workbook
    Dim workbook2 As Excel.Workbook
    Dim worksheet As Excel.Worksheet
    Dim worksheet1 As Excel.Worksheet
    Dim worksheet2 As Excel.Worksheet

    Private Sub Setup()

        excelApp = Globals.ThisAddIn.Application
        workbook = excelApp.ActiveWorkbook
        worksheet = workbook.ActiveSheet

        Dim rng As Excel.Range
        rng = worksheet.Range(TextBox1.Text)

        Dim r As Integer
        Dim c As Integer

        r = rng.Rows.Count
        c = rng.Columns.Count

        If r <> 1 And c <> 1 Then
            RadioButton1.Enabled = True
            RadioButton2.Enabled = True
            RadioButton3.Enabled = False
            RadioButton4.Enabled = False
        ElseIf r <> 1 And c = 1 Then
            RadioButton1.Enabled = False
            RadioButton2.Enabled = True
            RadioButton3.Enabled = True
            RadioButton4.Enabled = False
        ElseIf r = 1 And c <> 1 Then
            RadioButton1.Enabled = True
            RadioButton2.Enabled = False
            RadioButton3.Enabled = False
            RadioButton4.Enabled = True

        End If

    End Sub

    Private Function MaxValue(Arr)

        Dim max As Integer
        max = Arr(LBound(Arr))

        For i = LBound(Arr) + 1 To UBound(Arr)
            If Arr(i) > max Then
                max = Arr(i)
            End If
        Next

        MaxValue = max

    End Function

    Public Function GetBreakPoints(rng As Excel.Range, trace As Integer)

        Dim Arr() As Integer
        Dim Index As Integer
        Index = -1

        If trace = 1 Then
            For j = 1 To rng.Columns.Count
                If IsNothing(rng.Cells(1, j).Value) OrElse IsDBNull(rng.Cells(1, j).Value) OrElse String.IsNullOrEmpty(rng.Cells(1, j).Value.ToString()) Then
                    Index = Index + 1
                    ReDim Preserve Arr(Index)
                    Arr(Index) = j
                End If
            Next

            Index = Index + 1
            ReDim Preserve Arr(Index)
            Arr(Index) = rng.Columns.Count + 1

        Else
            For i = 1 To rng.Rows.Count
                If IsNothing(rng.Cells(i, 1).Value) OrElse IsDBNull(rng.Cells(i, 1).Value) OrElse String.IsNullOrEmpty(rng.Cells(i, 1).Value.ToString()) Then
                    Index = Index + 1
                    ReDim Preserve Arr(Index)
                    Arr(Index) = i
                End If
            Next
            Index = Index + 1
            ReDim Preserve Arr(Index)
            Arr(Index) = rng.Rows.Count + 1
        End If

        GetBreakPoints = Arr

    End Function
    Public Function GetLengths(Arr)
        Dim Arr2() As Integer
        Dim Index As Integer
        Index = -1
        Dim position As Integer
        position = 0
        Dim length As Integer

        For i = LBound(Arr) To UBound(Arr)
            length = Arr(i) - position - 1
            position = Arr(i)
            Index = Index + 1
            ReDim Preserve Arr2(Index)
            Arr2(Index) = length
        Next

        GetLengths = Arr2

    End Function

    Private Sub Display()


        CustomPanel1.Controls.Clear()
        CustomPanel2.Controls.Clear()

        excelApp = Globals.ThisAddIn.Application
        workbook = excelApp.ActiveWorkbook
        worksheet = workbook.ActiveSheet

        Dim rng As Excel.Range
        rng = worksheet.Range(TextBox1.Text)

        If rng.Rows.Count > 50 Then
            rng = worksheet.Range(rng.Cells(1, 1), rng.Cells(50, rng.Columns.Count))
        End If

        If rng.Columns.Count > 50 Then
            rng = worksheet.Range(rng.Cells(1, 1), rng.Cells(rng.Rows.Count, 50))
        End If

        Dim r As Integer
        Dim c As Integer

        r = rng.Rows.Count
        c = rng.Columns.Count


        Dim height As Integer
        Dim width As Integer

        If r > 1 And r <= 6 Then
            height = CustomPanel1.Height / r
        Else
            height = CustomPanel1.Height / 6
        End If

        If c > 1 And c <= 4 Then
            width = CustomPanel1.Width / c
        Else
            width = CustomPanel1.Width / 4
        End If

        For i = 1 To r
            For j = 1 To c
                Dim label As New System.Windows.Forms.Label
                label.Text = rng.Cells(i, j).Value
                If r <> 1 And c = 1 Then
                    label.Location = New System.Drawing.Point((2.5 - 1) * width, (i - 1) * height)
                ElseIf r = 1 And c <> 1 Then
                    label.Location = New System.Drawing.Point((j - 1) * width, (3.5 - 1) * height)
                Else
                    label.Location = New System.Drawing.Point((j - 1) * width, (i - 1) * height)
                End If
                label.Height = height
                label.Width = width
                label.BorderStyle = BorderStyle.FixedSingle
                label.TextAlign = ContentAlignment.MiddleCenter

                If CheckBox1.Checked = True Then

                    Dim cell As Excel.Range = rng.Cells(i, j)
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

        Dim X1 As Boolean
        X1 = RadioButton1.Checked

        Dim X2 As Boolean
        X2 = RadioButton2.Checked

        Dim X3 As Boolean
        X3 = RadioButton3.Checked

        Dim X4 As Boolean
        X4 = RadioButton4.Checked

        Dim X5 As Boolean
        X5 = RadioButton5.Checked

        Dim X6 As Boolean
        X6 = RadioButton6.Checked

        Dim X7 As Boolean
        X7 = RadioButton7.Checked

        Dim X8 As Boolean
        X8 = RadioButton8.Checked


        If X1 Then

            If (r * c) <= 6 Then
                height = CustomPanel2.Height / (r * c)
            Else
                height = CustomPanel2.Height / 6
            End If

            width = CustomPanel2.Width / 4

            Dim count As Integer
            count = 1

            If X5 Then

                For i = 1 To r
                    For j = 1 To c
                        Dim label As New System.Windows.Forms.Label
                        label.Text = rng.Cells(i, j).Value
                        label.Location = New System.Drawing.Point((2.5 - 1) * width, (count - 1) * height)
                        count = count + 1
                        label.Height = height
                        label.Width = width
                        label.BorderStyle = BorderStyle.FixedSingle
                        label.TextAlign = ContentAlignment.MiddleCenter

                        If CheckBox1.Checked = True Then
                            Dim cell As Excel.Range = rng.Cells(i, j)
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
                    Next
                Next

            ElseIf X6 Then

                For j = 1 To c
                    For i = 1 To r
                        Dim label As New System.Windows.Forms.Label
                        label.Text = rng.Cells(i, j).Value
                        label.Location = New System.Drawing.Point((2.5 - 1) * width, (count - 1) * height)
                        count = count + 1
                        label.Height = height
                        label.Width = width
                        label.BorderStyle = BorderStyle.FixedSingle
                        label.TextAlign = ContentAlignment.MiddleCenter

                        If CheckBox1.Checked = True Then
                            Dim cell As Excel.Range = rng.Cells(i, j)
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
                    Next
                Next

            End If

            CustomPanel2.AutoScroll = True

        End If

        If X2 Then

            If (r * c) <= 4 Then
                width = CustomPanel2.Width / (r * c)
            Else
                width = CustomPanel2.Width / 4
            End If

            height = CustomPanel2.Height / 6

            Dim count As Integer
            count = 1

            If X5 Then

                For i = 1 To r
                    For j = 1 To c
                        Dim label As New System.Windows.Forms.Label
                        label.Text = rng.Cells(i, j).Value
                        label.Location = New System.Drawing.Point((count - 1) * width, (3.5 - 1) * height)
                        count = count + 1
                        label.Height = height
                        label.Width = width
                        label.BorderStyle = BorderStyle.FixedSingle
                        label.TextAlign = ContentAlignment.MiddleCenter

                        If CheckBox1.Checked = True Then
                            Dim cell As Excel.Range = rng.Cells(i, j)
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
                    Next
                Next

            ElseIf X6 Then

                For j = 1 To c
                    For i = 1 To r
                        Dim label As New System.Windows.Forms.Label
                        label.Text = rng.Cells(i, j).Value
                        label.Location = New System.Drawing.Point((count - 1) * width, (3.5 - 1) * height)
                        count = count + 1
                        label.Height = height
                        label.Width = width
                        label.BorderStyle = BorderStyle.FixedSingle
                        label.TextAlign = ContentAlignment.MiddleCenter

                        If CheckBox1.Checked = True Then
                            Dim cell As Excel.Range = rng.Cells(i, j)
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
                    Next
                Next
            End If

            CustomPanel2.AutoScroll = True

        End If

        If X3 Then

            If X7 And (X5 Or X6) Then

                Dim BreakPoints() As Integer
                BreakPoints = GetBreakPoints(rng, 2)

                Dim lengths() As Integer
                lengths = GetLengths(BreakPoints)

                If X5 Then
                    r = UBound(BreakPoints) + 1
                    c = MaxValue(lengths)
                ElseIf X6 Then
                    c = UBound(BreakPoints) + 1
                    r = MaxValue(lengths)
                End If

                If r > 1 And r <= 6 Then
                    height = CustomPanel2.Height / r
                Else
                    height = CustomPanel2.Height / 6
                End If

                If c > 1 And c <= 4 Then
                    width = CustomPanel2.Width / c
                Else
                    width = CustomPanel2.Width / 4
                End If

                If X5 Then
                    Dim iRow As Integer
                    iRow = 0
                    For i = 1 To r
                        For j = 1 To c
                            Dim x As Integer
                            Dim y As Integer
                            x = iRow + j
                            y = 1
                            Dim label As New System.Windows.Forms.Label
                            If x <= BreakPoints(i - 1) Then
                                label.Text = rng.Cells(x, y).Value
                            Else
                                label.Text = ""
                            End If
                            label.Location = New System.Drawing.Point((j - 1) * width, (i - 1) * height)
                            label.Height = height
                            label.Width = width
                            label.BorderStyle = BorderStyle.FixedSingle
                            label.TextAlign = ContentAlignment.MiddleCenter

                            If CheckBox1.Checked = True Then
                                Dim cell As Excel.Range = rng.Cells(x, y)
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
                        Next
                        iRow = BreakPoints(i - 1)
                    Next

                ElseIf X6 Then
                    Dim iRow As Integer
                    iRow = 0
                    For j = 1 To c
                        For i = 1 To c
                            Dim x As Integer
                            Dim y As Integer
                            x = iRow + i
                            y = 1
                            Dim label As New System.Windows.Forms.Label
                            If x <= BreakPoints(j - 1) Then
                                label.Text = rng.Cells(x, y).Value
                            Else
                                label.Text = ""
                            End If
                            label.Location = New System.Drawing.Point((j - 1) * width, (i - 1) * height)
                            label.Height = height
                            label.Width = width
                            label.BorderStyle = BorderStyle.FixedSingle
                            label.TextAlign = ContentAlignment.MiddleCenter

                            If CheckBox1.Checked = True Then
                                Dim cell As Excel.Range = rng.Cells(x, y)
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
                        Next
                        iRow = BreakPoints(j - 1)
                    Next
                End If

            ElseIf X8 And (X5 Or X6) Then
                If X5 Then
                    r = Int(r / Int(TextBox2.Text)) + 1

                End If
            End If

            CustomPanel2.AutoScroll = True

        End If


    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Call Display()

    End Sub
End Class