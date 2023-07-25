Imports System.Drawing
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
    Dim rng As Excel.Range
    Dim rng2 As Excel.Range
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

    Private Sub Setup()

        excelApp = Globals.ThisAddIn.Application
        workbook = excelApp.ActiveWorkbook
        worksheet = workbook.ActiveSheet

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

        If RadioButton1.Checked = True Or RadioButton2.Checked = True Then
            TextBox2.Text = ""
            CustomGroupBox3.Enabled = False
            RadioButton7.Checked = False
            RadioButton8.Checked = False
        Else
            CustomGroupBox3.Enabled = True
        End If

        If RadioButton8.Checked = True Then
            TextBox2.Enabled = True
            TextBox2.Focus()
        Else
            TextBox2.Text = ""
            TextBox2.Enabled = False
        End If

        If RadioButton3.Checked = True Then
            RadioButton8.Text = "After number of rows:"
        ElseIf RadioButton4.Checked = True Then
            RadioButton8.Text = "After number of columns:"
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

    Private Function GetBreakPoints(rng As Excel.Range, trace As Integer)

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
    Private Function GetLengths(Arr)
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
                        For i = 1 To r
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

            ElseIf (X8 And TextBox2.Text <> "") And (X5 Or X6) Then

                If X5 Then
                    If r Mod Int(TextBox2.Text) = 0 Then
                        r = Int(r / Int(TextBox2.Text))
                    Else
                        r = Int(r / Int(TextBox2.Text)) + 1
                    End If
                    c = Int(TextBox2.Text)

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

                    For i = 1 To r
                        For j = 1 To c
                            Dim x As Integer
                            Dim y As Integer
                            x = (c * (i - 1)) + j
                            y = 1
                            Dim label As New System.Windows.Forms.Label
                            If x <= rng.Rows.Count Then
                                label.Text = rng.Cells(x, y).Value
                            Else
                                label.Text = ""
                            End If
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
                    Next

                ElseIf X6 Then
                    If r Mod Int(TextBox2.Text) = 0 Then
                        c = Int(r / Int(TextBox2.Text))
                    Else
                        c = Int(r / Int(TextBox2.Text)) + 1
                    End If
                    r = Int(TextBox2.Text)

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

                    For j = 1 To c
                        For i = 1 To r
                            Dim x As Integer
                            Dim y As Integer
                            x = (r * (j - 1)) + i
                            y = 1
                            Dim label As New System.Windows.Forms.Label
                            If x <= rng.Rows.Count Then
                                label.Text = rng.Cells(x, y).Value
                            Else
                                label.Text = ""
                            End If
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
                    Next

                End If
            End If

            CustomPanel2.AutoScroll = True

        End If

        If X4 Then

            If X7 And (X5 Or X6) Then

                Dim BreakPoints() As Integer
                BreakPoints = GetBreakPoints(rng, 1)

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
                    Dim iColumn As Integer
                    iColumn = 0
                    For i = 1 To r
                        For j = 1 To c
                            Dim x As Integer
                            Dim y As Integer
                            x = 1
                            y = iColumn + j
                            Dim label As New System.Windows.Forms.Label
                            If y <= BreakPoints(i - 1) Then
                                label.Text = rng.Cells(x, y).Value
                            Else
                                label.Text = ""
                            End If
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
                        iColumn = BreakPoints(i - 1)
                    Next

                ElseIf X6 Then
                    Dim iColumn As Integer
                    iColumn = 0
                    For j = 1 To c
                        For i = 1 To r
                            Dim x As Integer
                            Dim y As Integer
                            x = 1
                            y = iColumn + i
                            Dim label As New System.Windows.Forms.Label
                            If y <= BreakPoints(j - 1) Then
                                label.Text = rng.Cells(x, y).Value
                            Else
                                label.Text = ""
                            End If
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
                        iColumn = BreakPoints(j - 1)
                    Next
                End If

            ElseIf (X8 And TextBox2.Text <> "") And (X5 Or X6) Then

                If X5 Then
                    If c Mod Int(TextBox2.Text) = 0 Then
                        r = Int(c / Int(TextBox2.Text))
                    Else
                        r = Int(c / Int(TextBox2.Text)) + 1
                    End If
                    c = Int(TextBox2.Text)

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

                    For i = 1 To r
                        For j = 1 To c
                            Dim x As Integer
                            Dim y As Integer
                            x = 1
                            y = (c * (i - 1)) + j
                            Dim label As New System.Windows.Forms.Label
                            If x <= rng.Rows.Count Then
                                label.Text = rng.Cells(x, y).Value
                            Else
                                label.Text = ""
                            End If
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
                    Next

                ElseIf X6 Then
                    If c Mod Int(TextBox2.Text) = 0 Then
                        c = Int(c / Int(TextBox2.Text))
                    Else
                        c = Int(c / Int(TextBox2.Text)) + 1
                    End If
                    r = Int(TextBox2.Text)

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

                    For j = 1 To c
                        For i = 1 To r
                            Dim x As Integer
                            Dim y As Integer
                            x = 1
                            y = (r * (j - 1)) + i
                            Dim label As New System.Windows.Forms.Label
                            If x <= rng.Rows.Count Then
                                label.Text = rng.Cells(x, y).Value
                            Else
                                label.Text = ""
                            End If
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
                    Next

                End If
            End If

            CustomPanel2.AutoScroll = True

        End If


    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        excelApp = Globals.ThisAddIn.Application
        workbook = excelApp.ActiveWorkbook
        worksheet = workbook.ActiveSheet

        Dim worksheet2 As Excel.Worksheet = worksheet

        rng = worksheet.Range(TextBox1.Text)

        If RadioButton9.Checked = True Then
            rng2 = rng
        ElseIf RadioButton10.Checked = True And TextBox3.Text <> "" Then
            rng2 = worksheet.Range(TextBox3.Text)
        Else
            MessageBox.Show("Select a Destination Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        If CheckBox1.Checked = True Then
            worksheet.Copy(After:=workbook.Sheets(worksheet.Name))
        End If

        workbook.Sheets(worksheet.Name).Activate

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

        Dim r As Integer
        Dim c As Integer

        r = rng.Rows.Count
        c = rng.Columns.Count

        Dim i As Integer
        Dim j As Integer

        Dim Arr(r - 1, c - 1) As Object
        Dim Bolds(r - 1, c - 1) As Boolean
        Dim Italics(r - 1, c - 1) As Boolean
        Dim fontNames(r - 1, c - 1) As String
        Dim fontSizes(r - 1, c - 1) As Single
        Dim reds1(r - 1, c - 1) As Integer
        Dim reds2(r - 1, c - 1) As Integer
        Dim greens1(r - 1, c - 1) As Integer
        Dim greens2(r - 1, c - 1) As Integer
        Dim blues1(r - 1, c - 1) As Integer
        Dim blues2(r - 1, c - 1) As Integer

        For i = 1 To r
            For j = 1 To c
                Arr(i - 1, j - 1) = rng.Cells(i, j).Value

                If CheckBox1.Checked = True Then

                    Dim cell As Excel.Range = rng.Cells(i, j)
                    Dim font As Excel.Font = cell.Font

                    Bolds(i - 1, j - 1) = cell.Font.Bold
                    Italics(i - 1, j - 1) = cell.Font.Italic

                    fontSizes(i - 1, j - 1) = Convert.ToSingle(font.Size)
                    fontNames(i - 1, j - 1) = font.Name

                    If Not cell.Interior.ColorIndex = Excel.XlColorIndex.xlColorIndexNone Then
                        Dim colorValue1 As Long = CLng(cell.Interior.Color)
                        reds1(i - 1, j - 1) = colorValue1 Mod 256
                        greens1(i - 1, j - 1) = (colorValue1 \ 256) Mod 256
                        blues1(i - 1, j - 1) = (colorValue1 \ 256 \ 256) Mod 256
                    Else
                        reds1(i - 1, j - 1) = 255
                        greens1(i - 1, j - 1) = 255
                        blues1(i - 1, j - 1) = 255
                    End If

                    If Not cell.Font.ColorIndex = Excel.XlColorIndex.xlColorIndexNone Then
                        Dim colorValue2 As Long = CLng(cell.Font.Color)
                        reds2(i - 1, j - 1) = colorValue2 Mod 256
                        greens2(i - 1, j - 1) = (colorValue2 \ 256) Mod 256
                        blues2(i - 1, j - 1) = (colorValue2 \ 256 \ 256) Mod 256
                    Else
                        reds2(i - 1, j - 1) = 255
                        greens2(i - 1, j - 1) = 255
                        blues2(i - 1, j - 1) = 255
                    End If
                End If

            Next
        Next


        If X1 Then

            rng2 = worksheet2.Range(rng2.Cells(1, 1), rng2.Cells((r * c), 1))
            Dim rng2Address As String = rng2.Address
            worksheet2.Activate()
            rng2.Select()

            If Overlap(excelApp, worksheet, worksheet2, rng, rng2) = False Then
                Dim count As Integer
                count = 1

                If X5 Then
                    For i = 1 To r
                        For j = 1 To c
                            Dim x As Integer = count
                            Dim y As Integer = 1

                            If CheckBox1.Checked = False Then
                                rng2.Cells(x, y).Value = rng.Cells(i, j).Value
                                count = count + 1

                            ElseIf CheckBox1.Checked = True Then
                                rng.Cells(i, j).Copy()
                                rng2.Cells(x, y).PasteSpecial(Excel.XlPasteType.xlPasteValues)
                                rng2 = worksheet2.Range(rng2Address)
                                rng2.Cells(x, y).PasteSpecial(Excel.XlPasteType.xlPasteFormats)
                                rng2 = worksheet2.Range(rng2Address)
                                count = count + 1

                            End If

                        Next
                    Next

                    excelApp.CutCopyMode = Excel.XlCutCopyMode.xlCopy

                ElseIf X6 Then

                    For j = 1 To c
                        For i = 1 To r

                            Dim x As Integer = count
                            Dim y As Integer = 1

                            If CheckBox1.Checked = False Then
                                rng2.Cells(x, y).Value = rng.Cells(i, j).Value
                                count = count + 1

                            ElseIf CheckBox1.Checked = True Then
                                rng.Cells(i, j).Copy()
                                rng2.Cells(x, y).PasteSpecial(Excel.XlPasteType.xlPasteValues)
                                rng2 = worksheet2.Range(rng2Address)
                                rng2.Cells(x, y).PasteSpecial(Excel.XlPasteType.xlPasteFormats)
                                rng2 = worksheet2.Range(rng2Address)
                                count = count + 1

                            End If

                        Next
                    Next

                Else
                    MessageBox.Show("Choose One Transformation Option. ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

            Else
                Dim count As Integer
                count = 1
                If X5 Then
                    For i = 1 To r
                        For j = 1 To c
                            Dim x As Integer = count
                            Dim y As Integer = 1

                            rng2.Cells(x, y).Value = Arr(i - 1, j - 1)
                            count = count + 1

                            If CheckBox1.Checked = True Then

                                Dim cell2 As Excel.Range = rng2.Cells(x, y)
                                Dim font2 As Excel.Font = cell2.Font

                                Dim fontSize As Single = fontSizes(i - 1, j - 1)

                                rng2.Cells(x, y).Font.Name = fontNames(i - 1, j - 1)
                                rng2.Cells(x, y).Font.Size = fontSizes(i - 1, j - 1)

                                If Bolds(i - 1, j - 1) Then rng2.Cells(x, y).Font.Bold = True
                                If Italics(i - 1, j - 1) Then rng2.Cells(x, y).Font.Italic = True

                                If reds1(i - 1, j - 1) = 255 Then
                                    Dim red1 As Integer = reds1(i - 1, j - 1)
                                    Dim green1 As Integer = greens1(i - 1, j - 1)
                                    Dim blue1 As Integer = blues1(i - 1, j - 1)
                                    rng2.Cells(x, y).Interior.Color = System.Drawing.Color.FromArgb(red1, green1, blue1)
                                End If

                                If reds2(i - 1, j - 1) = 255 Then
                                    Dim red2 As Integer = reds2(i - 1, j - 1)
                                    Dim green2 As Integer = greens2(i - 1, j - 1)
                                    Dim blue2 As Integer = blues2(i - 1, j - 1)
                                    rng2.Cells(x, y).Font.Color = System.Drawing.Color.FromArgb(red2, green2, blue2)
                                End If
                            End If

                        Next
                    Next

                ElseIf X6 Then

                    For j = 1 To c
                        For i = 1 To r

                            Dim x As Integer = count
                            Dim y As Integer = 1

                            rng2.Cells(x, y).Value = Arr(i - 1, j - 1)
                            count = count + 1

                            If CheckBox1.Checked = True Then

                                Dim cell2 As Excel.Range = rng2.Cells(x, y)
                                Dim font2 As Excel.Font = cell2.Font

                                Dim fontSize As Single = fontSizes(i - 1, j - 1)

                                rng2.Cells(x, y).Font.Name = fontNames(i - 1, j - 1)
                                rng2.Cells(x, y).Font.Size = fontSizes(i - 1, j - 1)

                                If Bolds(i - 1, j - 1) Then rng2.Cells(x, y).Font.Bold = True
                                If Italics(i - 1, j - 1) Then rng2.Cells(x, y).Font.Italic = True

                                If reds1(i - 1, j - 1) = 255 Then
                                    Dim red1 As Integer = reds1(i - 1, j - 1)
                                    Dim green1 As Integer = greens1(i - 1, j - 1)
                                    Dim blue1 As Integer = blues1(i - 1, j - 1)
                                    rng2.Cells(x, y).Interior.Color = System.Drawing.Color.FromArgb(red1, green1, blue1)
                                End If

                                If reds2(i - 1, j - 1) = 255 Then
                                    Dim red2 As Integer = reds2(i - 1, j - 1)
                                    Dim green2 As Integer = greens2(i - 1, j - 1)
                                    Dim blue2 As Integer = blues2(i - 1, j - 1)
                                    rng2.Cells(x, y).Font.Color = System.Drawing.Color.FromArgb(red2, green2, blue2)
                                End If
                            End If

                        Next
                    Next

                Else
                    MessageBox.Show("Choose One Transformation Option. ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If
            End If

        ElseIf X2 Then
            rng2 = worksheet2.Range(rng2.Cells(1, 1), rng2.Cells(1, (r * c)))
            Dim rng2Address As String = rng2.Address
            worksheet2.Activate()
            rng2.Select()

            If Overlap(excelApp, worksheet, worksheet2, rng, rng2) = False Then

                Dim count As Integer
                count = 1

                If X5 Then
                    For i = 1 To r
                        For j = 1 To c

                            Dim x As Integer = 1
                            Dim y As Integer = count

                            If CheckBox1.Checked = False Then
                                rng2.Cells(x, y).Value = rng(i, j)
                                count = count + 1

                            ElseIf CheckBox1.Checked = True Then
                                rng.Cells(i, j).Copy()
                                rng2.Cells(x, y).PasteSpecial(Excel.XlPasteType.xlPasteValues)
                                rng2 = worksheet2.Range(rng2Address)
                                rng2.Cells(x, y).PasteSpecial(Excel.XlPasteType.xlPasteFormats)
                                rng2 = worksheet2.Range(rng2Address)
                                count = count + 1
                            End If

                        Next
                    Next

                ElseIf X6 Then

                    For j = 1 To c
                        For i = 1 To r

                            Dim x As Integer = 1
                            Dim y As Integer = count

                            If CheckBox1.Checked = False Then
                                rng2.Cells(x, y).Value = rng.Cells(i, j).Value
                                count = count + 1

                            ElseIf CheckBox1.Checked = True Then
                                rng.Cells(i, j).Copy()
                                rng2.Cells(x, y).PasteSpecial(Excel.XlPasteType.xlPasteValues)
                                rng2 = worksheet2.Range(rng2Address)
                                rng2.Cells(x, y).PasteSpecial(Excel.XlPasteType.xlPasteFormats)
                                rng2 = worksheet2.Range(rng2Address)
                                count = count + 1
                            End If

                        Next
                    Next

                Else
                    MessageBox.Show("Choose One Transformation Option. ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub

                End If
            Else
                Dim count As Integer
                count = 1

                If X5 Then

                    For i = 1 To r
                        For j = 1 To c
                            Dim x As Integer = 1
                            Dim y As Integer = count

                            rng2.Cells(x, y).Value = Arr(i - 1, j - 1)
                            count = count + 1

                            If CheckBox1.Checked = True Then

                                Dim cell2 As Excel.Range = rng2.Cells(x, y)
                                Dim font2 As Excel.Font = cell2.Font

                                Dim fontSize As Single = fontSizes(i - 1, j - 1)

                                rng2.Cells(x, y).Font.Name = fontNames(i - 1, j - 1)
                                rng2.Cells(x, y).Font.Size = fontSizes(i - 1, j - 1)

                                If Bolds(i - 1, j - 1) Then rng2.Cells(x, y).Font.Bold = True
                                If Italics(i - 1, j - 1) Then rng2.Cells(x, y).Font.Italic = True

                                If reds1(i - 1, j - 1) = 255 Then
                                    Dim red1 As Integer = reds1(i - 1, j - 1)
                                    Dim green1 As Integer = greens1(i - 1, j - 1)
                                    Dim blue1 As Integer = blues1(i - 1, j - 1)
                                    rng2.Cells(x, y).Interior.Color = System.Drawing.Color.FromArgb(red1, green1, blue1)
                                End If

                                If reds2(i - 1, j - 1) = 255 Then
                                    Dim red2 As Integer = reds2(i - 1, j - 1)
                                    Dim green2 As Integer = greens2(i - 1, j - 1)
                                    Dim blue2 As Integer = blues2(i - 1, j - 1)
                                    rng2.Cells(x, y).Font.Color = System.Drawing.Color.FromArgb(red2, green2, blue2)
                                End If
                            End If

                        Next
                    Next

                ElseIf X6 Then

                    For j = 1 To c
                        For i = 1 To r

                            Dim x As Integer = 1
                            Dim y As Integer = count

                            rng2.Cells(x, y).Value = Arr(i - 1, j - 1)
                            count = count + 1

                            If CheckBox1.Checked = True Then

                                Dim cell2 As Excel.Range = rng2.Cells(x, y)
                                Dim font2 As Excel.Font = cell2.Font

                                Dim fontSize As Single = fontSizes(i - 1, j - 1)

                                rng2.Cells(x, y).Font.Name = fontNames(i - 1, j - 1)
                                rng2.Cells(x, y).Font.Size = fontSizes(i - 1, j - 1)

                                If Bolds(i - 1, j - 1) Then rng2.Cells(x, y).Font.Bold = True
                                If Italics(i - 1, j - 1) Then rng2.Cells(x, y).Font.Italic = True

                                If reds1(i - 1, j - 1) = 255 Then
                                    Dim red1 As Integer = reds1(i - 1, j - 1)
                                    Dim green1 As Integer = greens1(i - 1, j - 1)
                                    Dim blue1 As Integer = blues1(i - 1, j - 1)
                                    rng2.Cells(x, y).Interior.Color = System.Drawing.Color.FromArgb(red1, green1, blue1)
                                End If

                                If reds2(i - 1, j - 1) = 255 Then
                                    Dim red2 As Integer = reds2(i - 1, j - 1)
                                    Dim green2 As Integer = greens2(i - 1, j - 1)
                                    Dim blue2 As Integer = blues2(i - 1, j - 1)
                                    rng2.Cells(x, y).Font.Color = System.Drawing.Color.FromArgb(red2, green2, blue2)
                                End If
                            End If

                        Next
                    Next

                Else
                    MessageBox.Show("Choose One Transformation Option. ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub

                End If
            End If

        ElseIf X3 Then

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

                rng2 = worksheet2.Range(rng2.Cells(1, 1), rng2.Cells(r, c))
                Dim rng2Address As String = rng2.Address
                worksheet2.Activate()
                rng2.Select()

                If Overlap(excelApp, worksheet, worksheet2, rng, rng2) = False Then
                    If X5 Then
                        Dim iRow As Integer
                        iRow = 0
                        For i = 1 To r
                            For j = 1 To c
                                Dim x As Integer
                                Dim y As Integer
                                x = iRow + j
                                y = 1
                                If x <= BreakPoints(i - 1) Then
                                    If CheckBox1.Checked = False Then
                                        rng2.Cells(i, j).Value = rng.Cells(x, y).Value

                                    ElseIf CheckBox1.Checked = True Then

                                        rng.Cells(x, y).Copy()
                                        rng2.Cells(i, j).PasteSpecial(Excel.XlPasteType.xlPasteValues)
                                        rng2 = worksheet2.Range(rng2Address)
                                        rng2.Cells(i, j).PasteSpecial(Excel.XlPasteType.xlPasteFormats)
                                        rng2 = worksheet2.Range(rng2Address)

                                    End If
                                End If
                            Next
                            iRow = BreakPoints(i - 1)
                        Next

                    ElseIf X6 Then
                        Dim iRow As Integer
                        iRow = 0
                        For j = 1 To c
                            For i = 1 To r
                                Dim x As Integer
                                Dim y As Integer
                                x = iRow + i
                                y = 1

                                If x <= BreakPoints(j - 1) Then
                                    If CheckBox1.Checked = False Then
                                        rng2.Cells(i, j).Value = rng.Cells(x, y).Value

                                    ElseIf CheckBox1.Checked = True Then

                                        rng.Cells(x, y).Copy()
                                        rng2.Cells(i, j).PasteSpecial(Excel.XlPasteType.xlPasteValues)
                                        rng2 = worksheet2.Range(rng2Address)
                                        rng2.Cells(i, j).PasteSpecial(Excel.XlPasteType.xlPasteFormats)
                                        rng2 = worksheet2.Range(rng2Address)

                                    End If
                                End If
                            Next
                            iRow = BreakPoints(j - 1)
                        Next
                    End If
                Else
                    If X5 Then
                        Dim iRow As Integer
                        iRow = 0
                        For i = 1 To r
                            For j = 1 To c
                                Dim x As Integer
                                Dim y As Integer
                                x = iRow + j
                                y = 1
                                If x <= BreakPoints(i - 1) Then
                                    rng2.Cells(i, j).Value = Arr(x - 1, y - 1)

                                    If CheckBox1.Checked = True Then

                                        Dim cell2 As Excel.Range = rng2.Cells(i, j)
                                        Dim font2 As Excel.Font = cell2.Font

                                        Dim fontSize As Single = fontSizes(x - 1, y - 1)

                                        rng2.Cells(i, j).Font.Name = fontNames(x - 1, y - 1)
                                        rng2.Cells(i, j).Font.Size = fontSizes(x - 1, y - 1)

                                        If Bolds(x - 1, y - 1) Then rng2.Cells(i, j).Font.Bold = True
                                        If Italics(x - 1, y - 1) Then rng2.Cells(i, j).Font.Italic = True

                                        If reds1(x - 1, y - 1) = 255 Then
                                            Dim red1 As Integer = reds1(x - 1, y - 1)
                                            Dim green1 As Integer = greens1(x - 1, y - 1)
                                            Dim blue1 As Integer = blues1(x - 1, y - 1)
                                            rng2.Cells(i, j).Interior.Color = System.Drawing.Color.FromArgb(red1, green1, blue1)
                                        End If

                                        If reds2(x - 1, y - 1) = 255 Then
                                            Dim red2 As Integer = reds2(x - 1, y - 1)
                                            Dim green2 As Integer = greens2(x - 1, y - 1)
                                            Dim blue2 As Integer = blues2(x - 1, y - 1)
                                            rng2.Cells(i, j).Font.Color = System.Drawing.Color.FromArgb(red2, green2, blue2)
                                        End If
                                    End If
                                End If
                            Next
                            iRow = BreakPoints(i - 1)
                        Next

                    ElseIf X6 Then
                        Dim iRow As Integer
                        iRow = 0
                        For j = 1 To c
                            For i = 1 To r
                                Dim x As Integer
                                Dim y As Integer
                                x = iRow + i
                                y = 1
                                If x <= BreakPoints(j - 1) Then
                                    rng2.Cells(i, j).Value = Arr(x - 1, y - 1)

                                    If CheckBox1.Checked = True Then

                                        Dim cell2 As Excel.Range = rng2.Cells(i, j)
                                        Dim font2 As Excel.Font = cell2.Font

                                        Dim fontSize As Single = fontSizes(x - 1, y - 1)

                                        rng2.Cells(i, j).Font.Name = fontNames(x - 1, y - 1)
                                        rng2.Cells(i, j).Font.Size = fontSizes(x - 1, y - 1)

                                        If Bolds(x - 1, y - 1) Then rng2.Cells(i, j).Font.Bold = True
                                        If Italics(x - 1, y - 1) Then rng2.Cells(i, j).Font.Italic = True

                                        If reds1(x - 1, y - 1) = 255 Then
                                            Dim red1 As Integer = reds1(x - 1, y - 1)
                                            Dim green1 As Integer = greens1(x - 1, y - 1)
                                            Dim blue1 As Integer = blues1(x - 1, y - 1)
                                            rng2.Cells(i, j).Interior.Color = System.Drawing.Color.FromArgb(red1, green1, blue1)
                                        End If

                                        If reds2(x - 1, y - 1) = 255 Then
                                            Dim red2 As Integer = reds2(x - 1, y - 1)
                                            Dim green2 As Integer = greens2(x - 1, y - 1)
                                            Dim blue2 As Integer = blues2(x - 1, y - 1)
                                            rng2.Cells(i, j).Font.Color = System.Drawing.Color.FromArgb(red2, green2, blue2)
                                        End If
                                    End If
                                End If
                            Next
                            iRow = BreakPoints(j - 1)
                        Next
                    End If
                End If

            ElseIf (X8 And TextBox2.Text <> "") And (X5 Or X6) Then

                If X5 Then
                    If r Mod Int(TextBox2.Text) = 0 Then
                        r = Int(r / Int(TextBox2.Text))
                    Else
                        r = Int(r / Int(TextBox2.Text)) + 1
                    End If
                    c = Int(TextBox2.Text)

                    rng2 = worksheet2.Range(rng2.Cells(1, 1), rng2.Cells(r, c))
                    Dim rng2Address As String = rng2.Address
                    worksheet2.Activate()
                    rng2.Select()

                    If Overlap(excelApp, worksheet, worksheet2, rng, rng2) = False Then
                        For i = 1 To r
                            For j = 1 To c
                                Dim x As Integer
                                Dim y As Integer
                                x = (c * (i - 1)) + j
                                y = 1
                                If x <= rng.Rows.Count Then
                                    If CheckBox1.Checked = False Then
                                        rng2.Cells(i, j).Value = rng.Cells(x, y).Value

                                    ElseIf CheckBox1.Checked = True Then

                                        rng.Cells(x, y).Copy()
                                        rng2.Cells(i, j).PasteSpecial(Excel.XlPasteType.xlPasteValues)
                                        rng2 = worksheet2.Range(rng2Address)
                                        rng2.Cells(i, j).PasteSpecial(Excel.XlPasteType.xlPasteFormats)
                                        rng2 = worksheet2.Range(rng2Address)

                                    End If
                                End If
                            Next
                        Next
                    Else
                        For i = 1 To r
                            For j = 1 To c
                                Dim x As Integer
                                Dim y As Integer
                                x = (c * (i - 1)) + j
                                y = 1
                                If x <= UBound(Arr, 1) + 1 Then
                                    rng2.Cells(i, j).Value = Arr(x - 1, y - 1)

                                    If CheckBox1.Checked = True Then

                                        Dim cell2 As Excel.Range = rng2.Cells(i, j)
                                        Dim font2 As Excel.Font = cell2.Font

                                        Dim fontSize As Single = fontSizes(x - 1, y - 1)

                                        rng2.Cells(i, j).Font.Name = fontNames(x - 1, y - 1)
                                        rng2.Cells(i, j).Font.Size = fontSizes(x - 1, y - 1)

                                        If Bolds(x - 1, y - 1) Then rng2.Cells(i, j).Font.Bold = True
                                        If Italics(x - 1, y - 1) Then rng2.Cells(i, j).Font.Italic = True

                                        If reds1(x - 1, y - 1) = 255 Then
                                            Dim red1 As Integer = reds1(x - 1, y - 1)
                                            Dim green1 As Integer = greens1(x - 1, y - 1)
                                            Dim blue1 As Integer = blues1(x - 1, y - 1)
                                            rng2.Cells(i, j).Interior.Color = System.Drawing.Color.FromArgb(red1, green1, blue1)
                                        End If

                                        If reds2(x - 1, y - 1) = 255 Then
                                            Dim red2 As Integer = reds2(x - 1, y - 1)
                                            Dim green2 As Integer = greens2(x - 1, y - 1)
                                            Dim blue2 As Integer = blues2(x - 1, y - 1)
                                            rng2.Cells(i, j).Font.Color = System.Drawing.Color.FromArgb(red2, green2, blue2)
                                        End If
                                    End If
                                End If
                            Next
                        Next
                    End If

                ElseIf X6 Then
                    If r Mod Int(TextBox2.Text) = 0 Then
                        c = Int(r / Int(TextBox2.Text))
                    Else
                        c = Int(r / Int(TextBox2.Text)) + 1
                    End If
                    r = Int(TextBox2.Text)

                    rng2 = worksheet2.Range(rng2.Cells(1, 1), rng2.Cells(r, c))
                    Dim rng2Address As String = rng2.Address
                    worksheet2.Activate()
                    rng2.Select()

                    If Overlap(excelApp, worksheet, worksheet2, rng, rng2) = False Then

                        For j = 1 To c
                            For i = 1 To r
                                Dim x As Integer
                                Dim y As Integer
                                x = (r * (j - 1)) + i
                                y = 1
                                If x <= rng.Rows.Count Then

                                    If CheckBox1.Checked = False Then
                                        rng2.Cells(i, j).Value = rng.Cells(x, y).Value

                                    ElseIf CheckBox1.Checked = True Then

                                        rng.Cells(x, y).Copy()
                                        rng2.Cells(i, j).PasteSpecial(Excel.XlPasteType.xlPasteValues)
                                        rng2 = worksheet2.Range(rng2Address)
                                        rng2.Cells(i, j).PasteSpecial(Excel.XlPasteType.xlPasteFormats)
                                        rng2 = worksheet2.Range(rng2Address)

                                    End If
                                End If
                            Next
                        Next

                    Else
                        For j = 1 To c
                            For i = 1 To r
                                Dim x As Integer
                                Dim y As Integer
                                x = (r * (j - 1)) + i
                                y = 1
                                If x <= UBound(Arr, 1) + 1 Then
                                    rng2.Cells(i, j).Value = Arr(x - 1, y - 1)

                                    If CheckBox1.Checked = True Then

                                        Dim cell2 As Excel.Range = rng2.Cells(i, j)
                                        Dim font2 As Excel.Font = cell2.Font

                                        Dim fontSize As Single = fontSizes(x - 1, y - 1)

                                        rng2.Cells(i, j).Font.Name = fontNames(x - 1, y - 1)
                                        rng2.Cells(i, j).Font.Size = fontSizes(x - 1, y - 1)

                                        If Bolds(x - 1, y - 1) Then rng2.Cells(i, j).Font.Bold = True
                                        If Italics(x - 1, y - 1) Then rng2.Cells(i, j).Font.Italic = True

                                        If reds1(x - 1, y - 1) = 255 Then
                                            Dim red1 As Integer = reds1(x - 1, y - 1)
                                            Dim green1 As Integer = greens1(x - 1, y - 1)
                                            Dim blue1 As Integer = blues1(x - 1, y - 1)
                                            rng2.Cells(i, j).Interior.Color = System.Drawing.Color.FromArgb(red1, green1, blue1)
                                        End If

                                        If reds2(x - 1, y - 1) = 255 Then
                                            Dim red2 As Integer = reds2(x - 1, y - 1)
                                            Dim green2 As Integer = greens2(x - 1, y - 1)
                                            Dim blue2 As Integer = blues2(x - 1, y - 1)
                                            rng2.Cells(i, j).Font.Color = System.Drawing.Color.FromArgb(red2, green2, blue2)
                                        End If
                                    End If
                                End If
                            Next
                        Next
                    End If

                End If

            Else
                MessageBox.Show("Select One Separator.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub

            End If

        ElseIf X4 Then

            If X7 And (X5 Or X6) Then

                Dim BreakPoints() As Integer
                BreakPoints = GetBreakPoints(rng, 1)

                Dim lengths() As Integer
                lengths = GetLengths(BreakPoints)

                If X5 Then
                    r = UBound(BreakPoints) + 1
                    c = MaxValue(lengths)
                ElseIf X6 Then
                    c = UBound(BreakPoints) + 1
                    r = MaxValue(lengths)
                End If

                rng2 = worksheet2.Range(rng2.Cells(1, 1), rng2.Cells(r, c))
                Dim rng2Address As String = rng2.Address
                worksheet2.Activate()
                rng2.Select()

                If Overlap(excelApp, worksheet, worksheet2, rng, rng2) = False Then
                    If X5 Then
                        Dim iColumn As Integer
                        iColumn = 0
                        For i = 1 To r
                            For j = 1 To c
                                Dim x As Integer
                                Dim y As Integer
                                x = 1
                                y = iColumn + j
                                If y <= BreakPoints(i - 1) Then
                                    If CheckBox1.Checked = False Then
                                        rng2.Cells(i, j).Value = rng.Cells(x, y).Value

                                    ElseIf CheckBox1.Checked = True Then

                                        rng.Cells(x, y).Copy()
                                        rng2.Cells(i, j).PasteSpecial(Excel.XlPasteType.xlPasteValues)
                                        rng2 = worksheet2.Range(rng2Address)
                                        rng2.Cells(i, j).PasteSpecial(Excel.XlPasteType.xlPasteFormats)
                                        rng2 = worksheet2.Range(rng2Address)

                                    End If
                                End If
                            Next
                            iColumn = BreakPoints(i - 1)
                        Next
                    ElseIf X6 Then
                        Dim iColumn As Integer
                        iColumn = 0
                        For j = 1 To c
                            For i = 1 To r
                                Dim x As Integer
                                Dim y As Integer
                                x = 1
                                y = iColumn + i
                                If y <= BreakPoints(j - 1) Then

                                    If CheckBox1.Checked = False Then
                                        rng2.Cells(i, j).Value = rng.Cells(x, y).Value

                                    ElseIf CheckBox1.Checked = True Then

                                        rng.Cells(x, y).Copy()
                                        rng2.Cells(i, j).PasteSpecial(Excel.XlPasteType.xlPasteValues)
                                        rng2 = worksheet2.Range(rng2Address)
                                        rng2.Cells(i, j).PasteSpecial(Excel.XlPasteType.xlPasteFormats)
                                        rng2 = worksheet2.Range(rng2Address)
                                    End If

                                End If
                            Next
                            iColumn = BreakPoints(j - 1)
                        Next
                    End If

                Else
                    If X5 Then

                        Dim iColumn As Integer
                        iColumn = 0
                        For i = 1 To r
                            For j = 1 To c
                                Dim x As Integer
                                Dim y As Integer
                                x = 1
                                y = iColumn + j
                                If y <= BreakPoints(i - 1) Then
                                    rng2.Cells(i, j).Value = Arr(x - 1, y - 1)

                                    If CheckBox1.Checked = True Then

                                        Dim cell2 As Excel.Range = rng2.Cells(i, j)
                                        Dim font2 As Excel.Font = cell2.Font

                                        Dim fontSize As Single = fontSizes(x - 1, y - 1)

                                        rng2.Cells(i, j).Font.Name = fontNames(x - 1, y - 1)
                                        rng2.Cells(i, j).Font.Size = fontSizes(x - 1, y - 1)

                                        If Bolds(x - 1, y - 1) Then rng2.Cells(i, j).Font.Bold = True
                                        If Italics(x - 1, y - 1) Then rng2.Cells(i, j).Font.Italic = True

                                        If reds1(x - 1, y - 1) = 255 Then
                                            Dim red1 As Integer = reds1(x - 1, y - 1)
                                            Dim green1 As Integer = greens1(x - 1, y - 1)
                                            Dim blue1 As Integer = blues1(x - 1, y - 1)
                                            rng2.Cells(i, j).Interior.Color = System.Drawing.Color.FromArgb(red1, green1, blue1)
                                        End If

                                        If reds2(x - 1, y - 1) = 255 Then
                                            Dim red2 As Integer = reds2(x - 1, y - 1)
                                            Dim green2 As Integer = greens2(x - 1, y - 1)
                                            Dim blue2 As Integer = blues2(x - 1, y - 1)
                                            rng2.Cells(i, j).Font.Color = System.Drawing.Color.FromArgb(red2, green2, blue2)
                                        End If
                                    End If
                                End If
                            Next
                            iColumn = BreakPoints(i - 1)
                        Next

                    ElseIf X6 Then
                        Dim iColumn As Integer
                        iColumn = 0
                        For j = 1 To c
                            For i = 1 To r
                                Dim x As Integer
                                Dim y As Integer
                                x = 1
                                y = iColumn + i
                                If y <= BreakPoints(j - 1) Then
                                    rng2.Cells(i, j).Value = Arr(x - 1, y - 1)

                                    If CheckBox1.Checked = True Then

                                        Dim cell2 As Excel.Range = rng2.Cells(i, j)
                                        Dim font2 As Excel.Font = cell2.Font

                                        Dim fontSize As Single = fontSizes(x - 1, y - 1)

                                        rng2.Cells(i, j).Font.Name = fontNames(x - 1, y - 1)
                                        rng2.Cells(i, j).Font.Size = fontSizes(x - 1, y - 1)

                                        If Bolds(x - 1, y - 1) Then rng2.Cells(i, j).Font.Bold = True
                                        If Italics(x - 1, y - 1) Then rng2.Cells(i, j).Font.Italic = True

                                        If reds1(x - 1, y - 1) = 255 Then
                                            Dim red1 As Integer = reds1(x - 1, y - 1)
                                            Dim green1 As Integer = greens1(x - 1, y - 1)
                                            Dim blue1 As Integer = blues1(x - 1, y - 1)
                                            rng2.Cells(i, j).Interior.Color = System.Drawing.Color.FromArgb(red1, green1, blue1)
                                        End If

                                        If reds2(x - 1, y - 1) = 255 Then
                                            Dim red2 As Integer = reds2(x - 1, y - 1)
                                            Dim green2 As Integer = greens2(x - 1, y - 1)
                                            Dim blue2 As Integer = blues2(x - 1, y - 1)
                                            rng2.Cells(i, j).Font.Color = System.Drawing.Color.FromArgb(red2, green2, blue2)
                                        End If
                                    End If
                                End If
                            Next
                            iColumn = BreakPoints(j - 1)
                        Next
                    End If
                End If

            ElseIf (X8 And TextBox2.Text <> "") And (X5 Or X6) Then

                If X5 Then
                    If c Mod Int(TextBox2.Text) = 0 Then
                        r = Int(c / Int(TextBox2.Text))
                    Else
                        r = Int(c / Int(TextBox2.Text)) + 1
                    End If
                    c = Int(TextBox2.Text)

                    rng2 = worksheet2.Range(rng2.Cells(1, 1), rng2.Cells(r, c))
                    Dim rng2Address As String = rng2.Address
                    worksheet2.Activate()
                    rng2.Select()

                    If Overlap(excelApp, worksheet, worksheet2, rng, rng2) = False Then
                        For i = 1 To r
                            For j = 1 To c
                                Dim x As Integer
                                Dim y As Integer
                                x = 1
                                y = (c * (i - 1)) + j

                                If y <= rng.Columns.Count Then
                                    If CheckBox1.Checked = False Then
                                        rng2.Cells(i, j).Value = rng.Cells(x, y).Value

                                    ElseIf CheckBox1.Checked = True Then

                                        rng.Cells(x, y).Copy()
                                        rng2.Cells(i, j).PasteSpecial(Excel.XlPasteType.xlPasteValues)
                                        rng2 = worksheet2.Range(rng2Address)
                                        rng2.Cells(i, j).PasteSpecial(Excel.XlPasteType.xlPasteFormats)
                                        rng2 = worksheet2.Range(rng2Address)
                                    End If

                                End If
                            Next
                        Next
                    Else
                        For i = 1 To r
                            For j = 1 To c
                                Dim x As Integer
                                Dim y As Integer
                                x = 1
                                y = (c * (i - 1)) + j
                                If y <= UBound(Arr, 2) + 1 Then
                                    rng2.Cells(i, j).Value = Arr(x - 1, y - 1)

                                    If CheckBox1.Checked = True Then

                                        Dim cell2 As Excel.Range = rng2.Cells(i, j)
                                        Dim font2 As Excel.Font = cell2.Font

                                        Dim fontSize As Single = fontSizes(x - 1, y - 1)

                                        rng2.Cells(i, j).Font.Name = fontNames(x - 1, y - 1)
                                        rng2.Cells(i, j).Font.Size = fontSizes(x - 1, y - 1)

                                        If Bolds(x - 1, y - 1) Then rng2.Cells(i, j).Font.Bold = True
                                        If Italics(x - 1, y - 1) Then rng2.Cells(i, j).Font.Italic = True

                                        If reds1(x - 1, y - 1) = 255 Then
                                            Dim red1 As Integer = reds1(x - 1, y - 1)
                                            Dim green1 As Integer = greens1(x - 1, y - 1)
                                            Dim blue1 As Integer = blues1(x - 1, y - 1)
                                            rng2.Cells(i, j).Interior.Color = System.Drawing.Color.FromArgb(red1, green1, blue1)
                                        End If

                                        If reds2(x - 1, y - 1) = 255 Then
                                            Dim red2 As Integer = reds2(x - 1, y - 1)
                                            Dim green2 As Integer = greens2(x - 1, y - 1)
                                            Dim blue2 As Integer = blues2(x - 1, y - 1)
                                            rng2.Cells(i, j).Font.Color = System.Drawing.Color.FromArgb(red2, green2, blue2)
                                        End If
                                    End If
                                End If
                            Next
                        Next
                    End If

                ElseIf X6 Then
                    If c Mod Int(TextBox2.Text) = 0 Then
                        c = Int(c / Int(TextBox2.Text))
                    Else
                        c = Int(c / Int(TextBox2.Text)) + 1
                    End If
                    r = Int(TextBox2.Text)

                    rng2 = worksheet2.Range(rng2.Cells(1, 1), rng2.Cells(r, c))
                    Dim rng2Address As String = rng2.Address
                    worksheet2.Activate()
                    rng2.Select()

                    If Overlap(excelApp, worksheet, worksheet2, rng, rng2) = False Then

                        For j = 1 To c
                            For i = 1 To r
                                Dim x As Integer
                                Dim y As Integer
                                x = 1
                                y = (r * (j - 1)) + i
                                If y <= rng.Columns.Count Then
                                    If CheckBox1.Checked = False Then
                                        rng2.Cells(i, j).Value = rng.Cells(x, y).Value

                                    ElseIf CheckBox1.Checked = True Then

                                        rng.Cells(x, y).Copy()
                                        rng2.Cells(i, j).PasteSpecial(Excel.XlPasteType.xlPasteValues)
                                        rng2 = worksheet2.Range(rng2Address)
                                        rng2.Cells(i, j).PasteSpecial(Excel.XlPasteType.xlPasteFormats)
                                        rng2 = worksheet2.Range(rng2Address)
                                    End If
                                End If
                            Next
                        Next
                    Else
                        For j = 1 To c
                            For i = 1 To r
                                Dim x As Integer
                                Dim y As Integer
                                x = 1
                                y = (r * (j - 1)) + i
                                If y <= UBound(Arr, 2) + 1 Then
                                    rng2.Cells(i, j).Value = Arr(x - 1, y - 1)

                                    If CheckBox1.Checked = True Then

                                        Dim cell2 As Excel.Range = rng2.Cells(i, j)
                                        Dim font2 As Excel.Font = cell2.Font

                                        Dim fontSize As Single = fontSizes(x - 1, y - 1)

                                        rng2.Cells(i, j).Font.Name = fontNames(x - 1, y - 1)
                                        rng2.Cells(i, j).Font.Size = fontSizes(x - 1, y - 1)

                                        If Bolds(x - 1, y - 1) Then rng2.Cells(i, j).Font.Bold = True
                                        If Italics(x - 1, y - 1) Then rng2.Cells(i, j).Font.Italic = True

                                        If reds1(x - 1, y - 1) = 255 Then
                                            Dim red1 As Integer = reds1(x - 1, y - 1)
                                            Dim green1 As Integer = greens1(x - 1, y - 1)
                                            Dim blue1 As Integer = blues1(x - 1, y - 1)
                                            rng2.Cells(i, j).Interior.Color = System.Drawing.Color.FromArgb(red1, green1, blue1)
                                        End If

                                        If reds2(x - 1, y - 1) = 255 Then
                                            Dim red2 As Integer = reds2(x - 1, y - 1)
                                            Dim green2 As Integer = greens2(x - 1, y - 1)
                                            Dim blue2 As Integer = blues2(x - 1, y - 1)
                                            rng2.Cells(i, j).Font.Color = System.Drawing.Color.FromArgb(red2, green2, blue2)
                                        End If
                                    End If
                                End If
                            Next
                        Next

                    End If
                End If

            Else
                MessageBox.Show("Select One Separator.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
                End If

            Else
                MessageBox.Show("Select One Transformation Type.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

        Try
            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            worksheet = workbook.ActiveSheet

            rng = worksheet.Range(TextBox1.Text)
            rng.Select()

            Call Display()

            Call Setup()

        Catch ex As Exception

        End Try

    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged

        Call Display()
        Call Setup()

    End Sub

    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged

        Call Display()
        Call Setup()

    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged

        Call Display()
        Call Setup()

    End Sub

    Private Sub RadioButton4_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton4.CheckedChanged

        Call Display()
        Call Setup()

    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged

        Call Display()
        Call Setup()

    End Sub

    Private Sub RadioButton5_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton5.CheckedChanged

        Call Display()
        Call Setup()

    End Sub

    Private Sub RadioButton6_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton6.CheckedChanged

        Call Display()
        Call Setup()

    End Sub

    Private Sub RadioButton7_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton7.CheckedChanged

        Call Display()
        Call Setup()

    End Sub

    Private Sub RadioButton8_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton8.CheckedChanged

        Call Display()
        Call Setup()

    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged


        Call Display()
        Call Setup()

    End Sub

    Private Sub PictureBox4_Click(sender As Object, e As EventArgs) Handles PictureBox4.Click

        Try

            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook

            Dim worksheet2 As Excel.Worksheet

            Dim userInput As Excel.Range = excelApp.InputBox("Select a range", Type:=8)
            Dim rng As Microsoft.Office.Interop.Excel.Range = userInput

            Try
                Dim sheetName As String
                sheetName = Split(rng.Address(True, True, Excel.XlReferenceStyle.xlA1, True), "]")(1)
                sheetName = Split(sheetName, "!")(0)
                worksheet2 = workbook.Worksheets(sheetName)
                worksheet2.Activate()

            Catch ex As Exception

            End Try

            rng.Select()

            rng = excelApp.Range(rng, rng.End(Microsoft.Office.Interop.Excel.XlDirection.xlDown))
            rng = excelApp.Range(rng, rng.End(Microsoft.Office.Interop.Excel.XlDirection.xlToRight))

            rng.Select()
            Me.TextBox1.Text = rng.Address
            Me.TextBox1.Focus()

        Catch ex As Exception

        End Try

    End Sub

    Private Sub PictureBox8_Click(sender As Object, e As EventArgs) Handles PictureBox8.Click

        Try

            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook

            Dim worksheet2 As Excel.Worksheet

            Dim userInput As Excel.Range = excelApp.InputBox("Select a range", Type:=8)
            Dim rng As Microsoft.Office.Interop.Excel.Range = userInput

            Try
                Dim sheetName As String
                sheetName = Split(rng.Address(True, True, Excel.XlReferenceStyle.xlA1, True), "]")(1)
                sheetName = Split(sheetName, "!")(0)
                worksheet2 = workbook.Worksheets(sheetName)
                worksheet2.Activate()
            Catch ex As Exception

            End Try

            rng.Select()

            TextBox1.Text = rng.Address
            TextBox1.Focus()

        Catch ex As Exception

        End Try

    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Me.Close()

    End Sub

    Private Sub PictureBox6_Click(sender As Object, e As EventArgs) Handles PictureBox6.Click

        Try

            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook

            Dim userInput As Excel.Range = excelApp.InputBox("Select a range", Type:=8)

            TextBox3.Text = userInput.Address
            TextBox3.Focus()

        Catch ex As Exception

        End Try

    End Sub

    Private Sub RadioButton10_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton10.CheckedChanged

        If RadioButton10.Checked = True Then
            TextBox3.Enabled = True
            TextBox3.Focus()
        Else
            TextBox3.Text = ""
            TextBox3.Enabled = False
        End If
    End Sub

    Private Sub Form7_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class