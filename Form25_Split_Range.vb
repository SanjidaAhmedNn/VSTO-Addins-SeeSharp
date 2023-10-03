Imports System.ComponentModel
Imports System.Diagnostics
Imports System.Drawing
Imports System.Reflection
Imports System.Reflection.Emit
Imports System.Security.Policy
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Button
Imports Microsoft.Office.Interop.Excel

Public Class Form25_Split_Range

    Dim WithEvents excelApp As Excel.Application
    Dim workBook As Excel.Workbook
    Dim workSheet As Excel.Worksheet
    Dim workSheet2 As Excel.Worksheet
    Dim rng As Excel.Range
    Dim rng2 As Excel.Range
    Dim selectedRange As Excel.Range

    Dim opened As Integer
    Dim FocusedTextBox As Integer
    Dim TextBoxChanged As Boolean

    Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As IntPtr, ByVal hWndInsertAfter As IntPtr, ByVal X As Integer, ByVal Y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal uFlags As UInteger) As Boolean
    Private Const SWP_NOMOVE As UInteger = &H2
    Private Const SWP_NOSIZE As UInteger = &H1
    Private Const SWP_NOACTIVATE As UInteger = &H10
    Private Const HWND_TOPMOST As Integer = -1

    Private Function MaxOfColumn(cRng As Excel.Range)

        Dim max As Integer
        max = Len(cRng.Cells(1, 1).value)

        For i = 2 To cRng.Rows.Count
            If Len(cRng.Cells(i, 1).value) > max Then
                max = Len(cRng.Cells(i, 1).value)
            End If
        Next

        If max < 7 Then
            max = 7
        End If

        MaxOfColumn = max

    End Function
    Private Function IsValidExcelCellReference(cellReference As String) As Boolean

        Dim cellPattern As String = "(\$?[A-Z]+\$?[0-9]+)"

        Dim referencePattern As String = "^" + cellPattern + "(:" + cellPattern + ")?$"

        Dim regex As New Regex(referencePattern)

        If regex.IsMatch(cellReference) Then
            Return True
        Else
            Return False
        End If

    End Function
    Private Function MaxOfArray(Arr)

        Dim max As Integer
        max = Len(Arr(LBound(Arr)))

        For i = LBound(Arr) + 1 To UBound(Arr)
            If Len(Arr(i)) > max Then
                max = Len(Arr(i))
            End If
        Next

        If max < 7 Then
            max = 7
        End If

        MaxOfArray = max

    End Function
    Private Function SeparateNumberText(Str As String)

        Dim Output(1) As String
        Output(0) = ""
        Output(1) = ""

        For i = 1 To Len(Str)
            If IsNumeric(Mid(Str, i, 1)) Then
                Output(0) = Output(0) & Mid(Str, i, 1)
            Else
                Output(1) = Output(1) & Mid(Str, i, 1)
            End If
        Next

        SeparateNumberText = Output

    End Function
    Public Function CountSeparator(source As String, separator As String) As Integer

        Dim count As Integer = 0
        Dim Position As Integer = 1

        For i = 1 To Len(source)
            If Mid(source, i, Len(separator)) = separator Then
                If i - Position > 0 Then
                    count = count + 1
                End If
                Position = i + Len(separator)
            End If
        Next

        If Position <= Len(source) Then
            count = count + 1
        End If

        CountSeparator = count

    End Function
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

    Private Sub Display()

        Try
            CustomPanel1.Controls.Clear()
            CustomPanel2.Controls.Clear()

            Dim displayRng As Excel.Range

            If rng.Rows.Count > 50 Then
                displayRng = rng.Rows("1:50")
            Else
                displayRng = rng
            End If

            Dim r As Integer = displayRng.Rows.Count
            Dim c As Integer = displayRng.Columns.Count

            Dim Height As Double
            Dim BaseWidth As Double
            Dim Width As Double

            If r <= 4 Then
                Height = CustomPanel1.Height / displayRng.Rows.Count
            Else
                Height = (119 / 4)
            End If

            BaseWidth = (260 / 3)

            Dim ordinate As Double = 0

            For j = 1 To c
                Dim cRng As Excel.Range = displayRng.Columns(j)
                Width = (MaxOfColumn(cRng) * BaseWidth) / 10
                For i = 1 To r
                    Dim label As New System.Windows.Forms.Label
                    label.Text = displayRng.Cells(i, j).Value
                    label.Location = New System.Drawing.Point(ordinate, (i - 1) * Height)
                    label.Height = Height
                    label.Width = Width
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

                        If IsDBNull(cell.Font.Color) Then
                            label.ForeColor = System.Drawing.Color.FromArgb(0, 0, 0)

                        ElseIf Not cell.Font.ColorIndex = Excel.XlColorIndex.xlColorIndexNone Then
                            Dim colorValue2 As Long = CLng(cell.Font.Color)
                            Dim red2 As Integer = colorValue2 Mod 256
                            Dim green2 As Integer = (colorValue2 \ 256) Mod 256
                            Dim blue2 As Integer = (colorValue2 \ 256 \ 256) Mod 256
                            label.ForeColor = System.Drawing.Color.FromArgb(red2, green2, blue2)
                        End If
                    End If
                    CustomPanel1.Controls.Add(label)
                Next
                ordinate = ordinate + Width
            Next

            CustomPanel1.AutoScroll = True

            Dim X1 As Boolean = RadioButton1.Checked
            Dim X2 As Boolean = RadioButton2.Checked
            Dim X3 As Boolean = RadioButton3.Checked
            Dim X7 As Boolean = RadioButton7.Checked
            Dim X8 As Boolean = RadioButton8.Checked
            Dim X9 As Boolean = RadioButton9.Checked
            Dim X10 As Boolean = RadioButton10.Checked
            Dim X11 As Boolean = RadioButton11.Checked
            Dim X12 As Boolean = ComboBox3.SelectedIndex <> -1

            If (X1 Or X2) And X12 And (X3 Or X7 Or X8 Or X9 Or X10 Or X11) Then

                Dim SplitColumn As Integer = ComboBox3.SelectedIndex + 1

                If X7 Or X8 Or X9 Or X10 Then

                    Dim Separator As String = ""
                    If X7 Then
                        Separator = ";"
                    ElseIf X8 Then
                        Separator = vbNewLine
                    ElseIf X9 Then
                        Separator = " "
                    ElseIf X10 Then
                        Separator = ComboBox2.Text
                    End If

                    If X1 Then
                        Dim widths(c) As Double
                        For j = 1 To c
                            widths(j - 1) = (MaxOfColumn(displayRng.Columns(j)) * BaseWidth) / 10
                        Next

                        Dim Values(0) As String
                        Dim ForFormats(0) As Integer

                        Dim Index As Integer = -1
                        Dim position As Integer

                        For i = 1 To r
                            Dim source As String = displayRng.Cells(i, SplitColumn).value
                            position = 1
                            For k = 1 To Len(source)
                                If Mid(source, k, Len(Separator)) = Separator Then
                                    If k - position > 0 Then
                                        Index = Index + 1
                                        ordinate = 0
                                        For j = 1 To SplitColumn - 1
                                            Dim label1 As New System.Windows.Forms.Label
                                            label1.Text = displayRng.Cells(i, j).Value
                                            label1.Location = New System.Drawing.Point(ordinate, Index * Height)
                                            label1.Height = Height
                                            label1.Width = widths(j - 1)
                                            label1.BorderStyle = BorderStyle.FixedSingle
                                            label1.TextAlign = ContentAlignment.MiddleCenter

                                            If CheckBox1.Checked = True Then

                                                Dim cell As Excel.Range = displayRng.Cells(i, j)
                                                Dim font As Excel.Font = cell.Font

                                                Dim fontStyle As FontStyle = FontStyle.Regular
                                                If cell.Font.Bold Then fontStyle = fontStyle Or FontStyle.Bold
                                                If cell.Font.Italic Then fontStyle = fontStyle Or FontStyle.Italic

                                                Dim fontSize As Single = Convert.ToSingle(font.Size)

                                                label1.Font = New System.Drawing.Font(font.ToString, fontSize, fontStyle)
                                                If Not cell.Interior.ColorIndex = Excel.XlColorIndex.xlColorIndexNone Then
                                                    Dim colorValue1 As Long = CLng(cell.Interior.Color)
                                                    Dim red1 As Integer = colorValue1 Mod 256
                                                    Dim green1 As Integer = (colorValue1 \ 256) Mod 256
                                                    Dim blue1 As Integer = (colorValue1 \ 256 \ 256) Mod 256
                                                    label1.BackColor = System.Drawing.Color.FromArgb(red1, green1, blue1)
                                                End If

                                                If IsDBNull(cell.Font.Color) Then
                                                    label1.ForeColor = System.Drawing.Color.FromArgb(0, 0, 0)

                                                ElseIf Not cell.Font.ColorIndex = Excel.XlColorIndex.xlColorIndexNone Then
                                                    Dim colorValue2 As Long = CLng(cell.Font.Color)
                                                    Dim red2 As Integer = colorValue2 Mod 256
                                                    Dim green2 As Integer = (colorValue2 \ 256) Mod 256
                                                    Dim blue2 As Integer = (colorValue2 \ 256 \ 256) Mod 256
                                                    label1.ForeColor = System.Drawing.Color.FromArgb(red2, green2, blue2)
                                                End If
                                            End If

                                            CustomPanel2.Controls.Add(label1)

                                            ordinate = ordinate + widths(j - 1)
                                        Next j
                                        ReDim Preserve Values(Index)
                                        ReDim Preserve ForFormats(Index)
                                        Values(Index) = Mid(source, position, k - position)
                                        ForFormats(Index) = i
                                    End If
                                    position = k + Len(Separator)
                                End If
                            Next
                            If position <= Len(source) Then
                                Index = Index + 1
                                ordinate = 0
                                For j = 1 To SplitColumn - 1
                                    Dim label1 As New System.Windows.Forms.Label
                                    label1.Text = displayRng.Cells(i, j).Value
                                    label1.Location = New System.Drawing.Point(ordinate, Index * Height)
                                    label1.Height = Height
                                    label1.Width = widths(j - 1)
                                    label1.BorderStyle = BorderStyle.FixedSingle
                                    label1.TextAlign = ContentAlignment.MiddleCenter


                                    If CheckBox1.Checked = True Then

                                        Dim cell As Excel.Range = displayRng.Cells(i, j)
                                        Dim font As Excel.Font = cell.Font

                                        Dim fontStyle As FontStyle = FontStyle.Regular
                                        If cell.Font.Bold Then fontStyle = fontStyle Or FontStyle.Bold
                                        If cell.Font.Italic Then fontStyle = fontStyle Or FontStyle.Italic

                                        Dim fontSize As Single = Convert.ToSingle(font.Size)

                                        label1.Font = New System.Drawing.Font(font.ToString, fontSize, fontStyle)
                                        If Not cell.Interior.ColorIndex = Excel.XlColorIndex.xlColorIndexNone Then
                                            Dim colorValue1 As Long = CLng(cell.Interior.Color)
                                            Dim red1 As Integer = colorValue1 Mod 256
                                            Dim green1 As Integer = (colorValue1 \ 256) Mod 256
                                            Dim blue1 As Integer = (colorValue1 \ 256 \ 256) Mod 256
                                            label1.BackColor = System.Drawing.Color.FromArgb(red1, green1, blue1)
                                        End If

                                        If IsDBNull(cell.Font.Color) Then
                                            label1.ForeColor = System.Drawing.Color.FromArgb(0, 0, 0)

                                        ElseIf Not cell.Font.ColorIndex = Excel.XlColorIndex.xlColorIndexNone Then
                                            Dim colorValue2 As Long = CLng(cell.Font.Color)
                                            Dim red2 As Integer = colorValue2 Mod 256
                                            Dim green2 As Integer = (colorValue2 \ 256) Mod 256
                                            Dim blue2 As Integer = (colorValue2 \ 256 \ 256) Mod 256
                                            label1.ForeColor = System.Drawing.Color.FromArgb(red2, green2, blue2)
                                        End If
                                    End If

                                    CustomPanel2.Controls.Add(label1)
                                    ordinate = ordinate + widths(j - 1)
                                Next j
                                ReDim Preserve Values(Index)
                                ReDim Preserve ForFormats(Index)
                                Values(Index) = Mid(source, position, Len(source) - position + 1)
                                ForFormats(Index) = i
                            End If
                        Next

                        Dim SplitOrdinate As Integer
                        SplitOrdinate = ordinate
                        Width = (MaxOfArray(Values) * BaseWidth) / 10

                        For m = LBound(Values) To UBound(Values)
                            ordinate = SplitOrdinate
                            Dim label1 As New System.Windows.Forms.Label
                            label1.Text = Values(m)
                            label1.Location = New System.Drawing.Point(ordinate, m * Height)
                            label1.Height = Height
                            label1.Width = Width
                            label1.BorderStyle = BorderStyle.FixedSingle
                            label1.TextAlign = ContentAlignment.MiddleCenter


                            If CheckBox1.Checked = True Then

                                Dim cell As Excel.Range = displayRng.Cells(ForFormats(m), SplitColumn)
                                Dim font As Excel.Font = cell.Font

                                Dim fontStyle As FontStyle = FontStyle.Regular
                                If cell.Font.Bold Then fontStyle = fontStyle Or FontStyle.Bold
                                If cell.Font.Italic Then fontStyle = fontStyle Or FontStyle.Italic

                                Dim fontSize As Single = Convert.ToSingle(font.Size)

                                label1.Font = New System.Drawing.Font(font.ToString, fontSize, fontStyle)
                                If Not cell.Interior.ColorIndex = Excel.XlColorIndex.xlColorIndexNone Then
                                    Dim colorValue1 As Long = CLng(cell.Interior.Color)
                                    Dim red1 As Integer = colorValue1 Mod 256
                                    Dim green1 As Integer = (colorValue1 \ 256) Mod 256
                                    Dim blue1 As Integer = (colorValue1 \ 256 \ 256) Mod 256
                                    label1.BackColor = System.Drawing.Color.FromArgb(red1, green1, blue1)
                                End If

                                If IsDBNull(cell.Font.Color) Then
                                    label1.ForeColor = System.Drawing.Color.FromArgb(0, 0, 0)

                                ElseIf Not cell.Font.ColorIndex = Excel.XlColorIndex.xlColorIndexNone Then
                                    Dim colorValue2 As Long = CLng(cell.Font.Color)
                                    Dim red2 As Integer = colorValue2 Mod 256
                                    Dim green2 As Integer = (colorValue2 \ 256) Mod 256
                                    Dim blue2 As Integer = (colorValue2 \ 256 \ 256) Mod 256
                                    label1.ForeColor = System.Drawing.Color.FromArgb(red2, green2, blue2)
                                End If
                            End If
                            CustomPanel2.Controls.Add(label1)
                            ordinate = ordinate + Width

                            For j = SplitColumn + 1 To c
                                Dim label2 As New System.Windows.Forms.Label
                                label2.Text = displayRng.Cells(ForFormats(m), j).value
                                label2.Location = New System.Drawing.Point(ordinate, m * Height)
                                label2.Height = Height
                                label2.Width = widths(j - 1)
                                label2.BorderStyle = BorderStyle.FixedSingle
                                label2.TextAlign = ContentAlignment.MiddleCenter

                                If CheckBox1.Checked = True Then

                                    Dim cell As Excel.Range = displayRng.Cells(ForFormats(m), j)
                                    Dim font As Excel.Font = cell.Font

                                    Dim fontStyle As FontStyle = FontStyle.Regular
                                    If cell.Font.Bold Then fontStyle = fontStyle Or FontStyle.Bold
                                    If cell.Font.Italic Then fontStyle = fontStyle Or FontStyle.Italic

                                    Dim fontSize As Single = Convert.ToSingle(font.Size)

                                    label2.Font = New System.Drawing.Font(font.ToString, fontSize, fontStyle)
                                    If Not cell.Interior.ColorIndex = Excel.XlColorIndex.xlColorIndexNone Then
                                        Dim colorValue1 As Long = CLng(cell.Interior.Color)
                                        Dim red1 As Integer = colorValue1 Mod 256
                                        Dim green1 As Integer = (colorValue1 \ 256) Mod 256
                                        Dim blue1 As Integer = (colorValue1 \ 256 \ 256) Mod 256
                                        label2.BackColor = System.Drawing.Color.FromArgb(red1, green1, blue1)
                                    End If

                                    If IsDBNull(cell.Font.Color) Then
                                        label2.ForeColor = System.Drawing.Color.FromArgb(0, 0, 0)

                                    ElseIf Not cell.Font.ColorIndex = Excel.XlColorIndex.xlColorIndexNone Then
                                        Dim colorValue2 As Long = CLng(cell.Font.Color)
                                        Dim red2 As Integer = colorValue2 Mod 256
                                        Dim green2 As Integer = (colorValue2 \ 256) Mod 256
                                        Dim blue2 As Integer = (colorValue2 \ 256 \ 256) Mod 256
                                        label2.ForeColor = System.Drawing.Color.FromArgb(red2, green2, blue2)
                                    End If
                                End If
                                CustomPanel2.Controls.Add(label2)
                                ordinate = ordinate + widths(j - 1)
                            Next
                        Next


                    ElseIf X2 Then

                        If c <= 4 Then
                            Height = CustomPanel2.Height / c
                        Else
                            Height = (119 / 4)
                        End If

                        Dim position As Integer = 1
                        Dim Index As Integer
                        ordinate = 0
                        For i = 1 To r
                            Dim source As String = displayRng.Cells(i, SplitColumn).value
                            Dim values(c - 1) As String
                            Index = -1
                            For j = 1 To c
                                Index = Index + 1
                                values(j - 1) = displayRng.Cells(i, j).value
                            Next
                            position = 1
                            For k = 1 To Len(source)
                                If Mid(source, k, Len(Separator)) = Separator Then
                                    If k - position > 0 Then
                                        values(SplitColumn - 1) = Mid(source, position, k - position)
                                        Width = (MaxOfArray(values) * BaseWidth) / 10
                                        For m = LBound(values) To UBound(values)
                                            Dim label1 As New System.Windows.Forms.Label
                                            label1.Text = values(m)
                                            label1.Location = New System.Drawing.Point(ordinate, m * Height)
                                            label1.Height = Height
                                            label1.Width = Width
                                            label1.BorderStyle = BorderStyle.FixedSingle
                                            label1.TextAlign = ContentAlignment.MiddleCenter

                                            If CheckBox1.Checked = True Then

                                                Dim cell As Excel.Range = displayRng.Cells(i, m + 1)
                                                Dim font As Excel.Font = cell.Font

                                                Dim fontStyle As FontStyle = FontStyle.Regular
                                                If cell.Font.Bold Then fontStyle = fontStyle Or FontStyle.Bold
                                                If cell.Font.Italic Then fontStyle = fontStyle Or FontStyle.Italic

                                                Dim fontSize As Single = Convert.ToSingle(font.Size)

                                                label1.Font = New System.Drawing.Font(font.ToString, fontSize, fontStyle)
                                                If Not cell.Interior.ColorIndex = Excel.XlColorIndex.xlColorIndexNone Then
                                                    Dim colorValue1 As Long = CLng(cell.Interior.Color)
                                                    Dim red1 As Integer = colorValue1 Mod 256
                                                    Dim green1 As Integer = (colorValue1 \ 256) Mod 256
                                                    Dim blue1 As Integer = (colorValue1 \ 256 \ 256) Mod 256
                                                    label1.BackColor = System.Drawing.Color.FromArgb(red1, green1, blue1)
                                                End If

                                                If IsDBNull(cell.Font.Color) Then
                                                    label1.ForeColor = System.Drawing.Color.FromArgb(0, 0, 0)

                                                ElseIf Not cell.Font.ColorIndex = Excel.XlColorIndex.xlColorIndexNone Then
                                                    Dim colorValue2 As Long = CLng(cell.Font.Color)
                                                    Dim red2 As Integer = colorValue2 Mod 256
                                                    Dim green2 As Integer = (colorValue2 \ 256) Mod 256
                                                    Dim blue2 As Integer = (colorValue2 \ 256 \ 256) Mod 256
                                                    label1.ForeColor = System.Drawing.Color.FromArgb(red2, green2, blue2)
                                                End If
                                            End If

                                            CustomPanel2.Controls.Add(label1)
                                        Next
                                        ordinate = ordinate + Width
                                    End If
                                    position = k + Len(Separator)
                                End If
                            Next
                            If position <= Len(source) Then
                                values(SplitColumn - 1) = Mid(source, position, Len(source) - position + 1)
                                Width = (MaxOfArray(values) * BaseWidth) / 10
                                For m = LBound(values) To UBound(values)
                                    Dim label1 As New System.Windows.Forms.Label
                                    label1.Text = values(m)
                                    label1.Location = New System.Drawing.Point(ordinate, m * Height)
                                    label1.Height = Height
                                    label1.Width = Width
                                    label1.BorderStyle = BorderStyle.FixedSingle
                                    label1.TextAlign = ContentAlignment.MiddleCenter

                                    If CheckBox1.Checked = True Then

                                        Dim cell As Excel.Range = displayRng.Cells(i, m + 1)
                                        Dim font As Excel.Font = cell.Font

                                        Dim fontStyle As FontStyle = FontStyle.Regular
                                        If cell.Font.Bold Then fontStyle = fontStyle Or FontStyle.Bold
                                        If cell.Font.Italic Then fontStyle = fontStyle Or FontStyle.Italic

                                        Dim fontSize As Single = Convert.ToSingle(font.Size)

                                        label1.Font = New System.Drawing.Font(font.ToString, fontSize, fontStyle)
                                        If Not cell.Interior.ColorIndex = Excel.XlColorIndex.xlColorIndexNone Then
                                            Dim colorValue1 As Long = CLng(cell.Interior.Color)
                                            Dim red1 As Integer = colorValue1 Mod 256
                                            Dim green1 As Integer = (colorValue1 \ 256) Mod 256
                                            Dim blue1 As Integer = (colorValue1 \ 256 \ 256) Mod 256
                                            label1.BackColor = System.Drawing.Color.FromArgb(red1, green1, blue1)
                                        End If

                                        If IsDBNull(cell.Font.Color) Then
                                            label1.ForeColor = System.Drawing.Color.FromArgb(0, 0, 0)

                                        ElseIf Not cell.Font.ColorIndex = Excel.XlColorIndex.xlColorIndexNone Then
                                            Dim colorValue2 As Long = CLng(cell.Font.Color)
                                            Dim red2 As Integer = colorValue2 Mod 256
                                            Dim green2 As Integer = (colorValue2 \ 256) Mod 256
                                            Dim blue2 As Integer = (colorValue2 \ 256 \ 256) Mod 256
                                            label1.ForeColor = System.Drawing.Color.FromArgb(red2, green2, blue2)
                                        End If
                                    End If

                                    CustomPanel2.Controls.Add(label1)
                                Next
                                ordinate = ordinate + Width
                            End If
                        Next
                    End If

                ElseIf X3 Then
                    If X1 Then
                        Dim widths(c) As Double
                        For j = 1 To c
                            widths(j - 1) = (MaxOfColumn(displayRng.Columns(j)) * BaseWidth) / 10
                        Next

                        Dim Values(0) As String
                        Dim Index As Integer = -1

                        For i = 1 To r

                            Dim source As String = displayRng.Cells(i, SplitColumn).value
                            Dim NumberText(1) As String
                            NumberText = SeparateNumberText(source)
                            Dim Number As String = NumberText(0)
                            Dim Text As String = NumberText(1)

                            ordinate = 0
                            Index = Index + 1
                            For j = 1 To SplitColumn - 1
                                Dim label1 As New System.Windows.Forms.Label
                                label1.Text = displayRng.Cells(i, j).Value
                                label1.Location = New System.Drawing.Point(ordinate, Index * Height)
                                label1.Height = Height
                                label1.Width = widths(j - 1)
                                label1.BorderStyle = BorderStyle.FixedSingle
                                label1.TextAlign = ContentAlignment.MiddleCenter

                                If CheckBox1.Checked = True Then

                                    Dim cell As Excel.Range = displayRng.Cells(i, j)
                                    Dim font As Excel.Font = cell.Font

                                    Dim fontStyle As FontStyle = FontStyle.Regular
                                    If cell.Font.Bold Then fontStyle = fontStyle Or FontStyle.Bold
                                    If cell.Font.Italic Then fontStyle = fontStyle Or FontStyle.Italic

                                    Dim fontSize As Single = Convert.ToSingle(font.Size)

                                    label1.Font = New System.Drawing.Font(font.ToString, fontSize, fontStyle)
                                    If Not cell.Interior.ColorIndex = Excel.XlColorIndex.xlColorIndexNone Then
                                        Dim colorValue1 As Long = CLng(cell.Interior.Color)
                                        Dim red1 As Integer = colorValue1 Mod 256
                                        Dim green1 As Integer = (colorValue1 \ 256) Mod 256
                                        Dim blue1 As Integer = (colorValue1 \ 256 \ 256) Mod 256
                                        label1.BackColor = System.Drawing.Color.FromArgb(red1, green1, blue1)
                                    End If

                                    If IsDBNull(cell.Font.Color) Then
                                        label1.ForeColor = System.Drawing.Color.FromArgb(0, 0, 0)

                                    ElseIf Not cell.Font.ColorIndex = Excel.XlColorIndex.xlColorIndexNone Then
                                        Dim colorValue2 As Long = CLng(cell.Font.Color)
                                        Dim red2 As Integer = colorValue2 Mod 256
                                        Dim green2 As Integer = (colorValue2 \ 256) Mod 256
                                        Dim blue2 As Integer = (colorValue2 \ 256 \ 256) Mod 256
                                        label1.ForeColor = System.Drawing.Color.FromArgb(red2, green2, blue2)
                                    End If
                                End If

                                CustomPanel2.Controls.Add(label1)
                                ordinate = ordinate + widths(j - 1)
                            Next j

                            ReDim Preserve Values(Index)
                            Values(Index) = Number

                            ordinate = 0
                            Index = Index + 1
                            For j = 1 To SplitColumn - 1
                                Dim label1 As New System.Windows.Forms.Label
                                label1.Text = displayRng.Cells(i, j).Value
                                label1.Location = New System.Drawing.Point(ordinate, Index * Height)
                                label1.Height = Height
                                label1.Width = widths(j - 1)
                                label1.BorderStyle = BorderStyle.FixedSingle
                                label1.TextAlign = ContentAlignment.MiddleCenter

                                If CheckBox1.Checked = True Then

                                    Dim cell As Excel.Range = displayRng.Cells(i, j)
                                    Dim font As Excel.Font = cell.Font

                                    Dim fontStyle As FontStyle = FontStyle.Regular
                                    If cell.Font.Bold Then fontStyle = fontStyle Or FontStyle.Bold
                                    If cell.Font.Italic Then fontStyle = fontStyle Or FontStyle.Italic

                                    Dim fontSize As Single = Convert.ToSingle(font.Size)

                                    label1.Font = New System.Drawing.Font(font.ToString, fontSize, fontStyle)
                                    If Not cell.Interior.ColorIndex = Excel.XlColorIndex.xlColorIndexNone Then
                                        Dim colorValue1 As Long = CLng(cell.Interior.Color)
                                        Dim red1 As Integer = colorValue1 Mod 256
                                        Dim green1 As Integer = (colorValue1 \ 256) Mod 256
                                        Dim blue1 As Integer = (colorValue1 \ 256 \ 256) Mod 256
                                        label1.BackColor = System.Drawing.Color.FromArgb(red1, green1, blue1)
                                    End If

                                    If IsDBNull(cell.Font.Color) Then
                                        label1.ForeColor = System.Drawing.Color.FromArgb(0, 0, 0)

                                    ElseIf Not cell.Font.ColorIndex = Excel.XlColorIndex.xlColorIndexNone Then
                                        Dim colorValue2 As Long = CLng(cell.Font.Color)
                                        Dim red2 As Integer = colorValue2 Mod 256
                                        Dim green2 As Integer = (colorValue2 \ 256) Mod 256
                                        Dim blue2 As Integer = (colorValue2 \ 256 \ 256) Mod 256
                                        label1.ForeColor = System.Drawing.Color.FromArgb(red2, green2, blue2)
                                    End If
                                End If
                                CustomPanel2.Controls.Add(label1)
                                ordinate = ordinate + widths(j - 1)
                            Next j
                            ReDim Preserve Values(Index)
                            Values(Index) = Text
                        Next

                        Width = (MaxOfArray(Values) * BaseWidth) / 10
                        Dim SplitOrdinate As Double
                        SplitOrdinate = ordinate

                        For i = LBound(Values) To UBound(Values)
                            ordinate = SplitOrdinate
                            Dim label1 As New System.Windows.Forms.Label
                            label1.Text = Values(i)
                            label1.Location = New System.Drawing.Point(ordinate, i * Height)
                            label1.Height = Height
                            label1.Width = Width
                            label1.BorderStyle = BorderStyle.FixedSingle
                            label1.TextAlign = ContentAlignment.MiddleCenter

                            If CheckBox1.Checked = True Then

                                Dim cell As Excel.Range = displayRng.Cells(Int(i / 2) + 1, SplitColumn)
                                Dim font As Excel.Font = cell.Font

                                Dim fontStyle As FontStyle = FontStyle.Regular
                                If cell.Font.Bold Then fontStyle = fontStyle Or FontStyle.Bold
                                If cell.Font.Italic Then fontStyle = fontStyle Or FontStyle.Italic

                                Dim fontSize As Single = Convert.ToSingle(font.Size)

                                label1.Font = New System.Drawing.Font(font.ToString, fontSize, fontStyle)
                                If Not cell.Interior.ColorIndex = Excel.XlColorIndex.xlColorIndexNone Then
                                    Dim colorValue1 As Long = CLng(cell.Interior.Color)
                                    Dim red1 As Integer = colorValue1 Mod 256
                                    Dim green1 As Integer = (colorValue1 \ 256) Mod 256
                                    Dim blue1 As Integer = (colorValue1 \ 256 \ 256) Mod 256
                                    label1.BackColor = System.Drawing.Color.FromArgb(red1, green1, blue1)
                                End If

                                If IsDBNull(cell.Font.Color) Then
                                    label1.ForeColor = System.Drawing.Color.FromArgb(0, 0, 0)

                                ElseIf Not cell.Font.ColorIndex = Excel.XlColorIndex.xlColorIndexNone Then
                                    Dim colorValue2 As Long = CLng(cell.Font.Color)
                                    Dim red2 As Integer = colorValue2 Mod 256
                                    Dim green2 As Integer = (colorValue2 \ 256) Mod 256
                                    Dim blue2 As Integer = (colorValue2 \ 256 \ 256) Mod 256
                                    label1.ForeColor = System.Drawing.Color.FromArgb(red2, green2, blue2)
                                End If
                            End If

                            CustomPanel2.Controls.Add(label1)
                            ordinate = ordinate + Width

                            For j = SplitColumn + 1 To c
                                Dim label2 As New System.Windows.Forms.Label
                                label2.Text = displayRng.Cells(Int(i / 2) + 1, j).value
                                label2.Location = New System.Drawing.Point(ordinate, i * Height)
                                label2.Height = Height
                                label2.Width = widths(j - 1)
                                label2.BorderStyle = BorderStyle.FixedSingle
                                label2.TextAlign = ContentAlignment.MiddleCenter

                                If CheckBox1.Checked = True Then

                                    Dim cell As Excel.Range = displayRng.Cells(Int(i / 2) + 1, c)
                                    Dim font As Excel.Font = cell.Font

                                    Dim fontStyle As FontStyle = FontStyle.Regular
                                    If cell.Font.Bold Then fontStyle = fontStyle Or FontStyle.Bold
                                    If cell.Font.Italic Then fontStyle = fontStyle Or FontStyle.Italic

                                    Dim fontSize As Single = Convert.ToSingle(font.Size)

                                    label2.Font = New System.Drawing.Font(font.ToString, fontSize, fontStyle)
                                    If Not cell.Interior.ColorIndex = Excel.XlColorIndex.xlColorIndexNone Then
                                        Dim colorValue1 As Long = CLng(cell.Interior.Color)
                                        Dim red1 As Integer = colorValue1 Mod 256
                                        Dim green1 As Integer = (colorValue1 \ 256) Mod 256
                                        Dim blue1 As Integer = (colorValue1 \ 256 \ 256) Mod 256
                                        label2.BackColor = System.Drawing.Color.FromArgb(red1, green1, blue1)
                                    End If

                                    If IsDBNull(cell.Font.Color) Then
                                        label2.ForeColor = System.Drawing.Color.FromArgb(0, 0, 0)

                                    ElseIf Not cell.Font.ColorIndex = Excel.XlColorIndex.xlColorIndexNone Then
                                        Dim colorValue2 As Long = CLng(cell.Font.Color)
                                        Dim red2 As Integer = colorValue2 Mod 256
                                        Dim green2 As Integer = (colorValue2 \ 256) Mod 256
                                        Dim blue2 As Integer = (colorValue2 \ 256 \ 256) Mod 256
                                        label2.ForeColor = System.Drawing.Color.FromArgb(red2, green2, blue2)
                                    End If
                                End If
                                CustomPanel2.Controls.Add(label2)
                                ordinate = ordinate + widths(j - 1)
                            Next
                        Next

                    ElseIf X2 Then

                        If c <= 4 Then
                            Height = CustomPanel2.Height / c
                        Else
                            Height = (119 / 4)
                        End If

                        Dim position As Integer = 1
                        Dim Index As Integer
                        ordinate = 0
                        For i = 1 To r
                            Dim source As String = displayRng.Cells(i, SplitColumn).value
                            Dim NumberText(1) As String
                            NumberText = SeparateNumberText(source)
                            Dim Number As String = NumberText(0)
                            Dim Text As String = NumberText(1)

                            Dim values(c - 1) As String
                            Index = -1
                            For j = 1 To c - 1
                                Index = Index + 1
                                values(j - 1) = displayRng.Cells(i, j).value
                            Next
                            values(SplitColumn - 1) = Number
                            Width = (MaxOfArray(values) * BaseWidth) / 10
                            For m = LBound(values) To UBound(values)
                                Dim label1 As New System.Windows.Forms.Label
                                label1.Text = values(m)
                                label1.Location = New System.Drawing.Point(ordinate, m * Height)
                                label1.Height = Height
                                label1.Width = Width
                                label1.BorderStyle = BorderStyle.FixedSingle
                                label1.TextAlign = ContentAlignment.MiddleCenter

                                If CheckBox1.Checked = True Then

                                    Dim cell As Excel.Range = displayRng.Cells(i, m + 1)
                                    Dim font As Excel.Font = cell.Font

                                    Dim fontStyle As FontStyle = FontStyle.Regular
                                    If cell.Font.Bold Then fontStyle = fontStyle Or FontStyle.Bold
                                    If cell.Font.Italic Then fontStyle = fontStyle Or FontStyle.Italic

                                    Dim fontSize As Single = Convert.ToSingle(font.Size)

                                    label1.Font = New System.Drawing.Font(font.ToString, fontSize, fontStyle)
                                    If Not cell.Interior.ColorIndex = Excel.XlColorIndex.xlColorIndexNone Then
                                        Dim colorValue1 As Long = CLng(cell.Interior.Color)
                                        Dim red1 As Integer = colorValue1 Mod 256
                                        Dim green1 As Integer = (colorValue1 \ 256) Mod 256
                                        Dim blue1 As Integer = (colorValue1 \ 256 \ 256) Mod 256
                                        label1.BackColor = System.Drawing.Color.FromArgb(red1, green1, blue1)
                                    End If

                                    If IsDBNull(cell.Font.Color) Then
                                        label1.ForeColor = System.Drawing.Color.FromArgb(0, 0, 0)

                                    ElseIf Not cell.Font.ColorIndex = Excel.XlColorIndex.xlColorIndexNone Then
                                        Dim colorValue2 As Long = CLng(cell.Font.Color)
                                        Dim red2 As Integer = colorValue2 Mod 256
                                        Dim green2 As Integer = (colorValue2 \ 256) Mod 256
                                        Dim blue2 As Integer = (colorValue2 \ 256 \ 256) Mod 256
                                        label1.ForeColor = System.Drawing.Color.FromArgb(red2, green2, blue2)
                                    End If
                                End If
                                CustomPanel2.Controls.Add(label1)
                            Next
                            ordinate = ordinate + Width

                            values(SplitColumn - 1) = Text
                            Width = (MaxOfArray(values) * BaseWidth) / 10
                            For m = LBound(values) To UBound(values)
                                Dim label1 As New System.Windows.Forms.Label
                                label1.Text = values(m)
                                label1.Location = New System.Drawing.Point(ordinate, m * Height)
                                label1.Height = Height
                                label1.Width = Width
                                label1.BorderStyle = BorderStyle.FixedSingle
                                label1.TextAlign = ContentAlignment.MiddleCenter

                                If CheckBox1.Checked = True Then

                                    Dim cell As Excel.Range = displayRng.Cells(i, m + 1)
                                    Dim font As Excel.Font = cell.Font

                                    Dim fontStyle As FontStyle = FontStyle.Regular
                                    If cell.Font.Bold Then fontStyle = fontStyle Or FontStyle.Bold
                                    If cell.Font.Italic Then fontStyle = fontStyle Or FontStyle.Italic

                                    Dim fontSize As Single = Convert.ToSingle(font.Size)

                                    label1.Font = New System.Drawing.Font(font.ToString, fontSize, fontStyle)
                                    If Not cell.Interior.ColorIndex = Excel.XlColorIndex.xlColorIndexNone Then
                                        Dim colorValue1 As Long = CLng(cell.Interior.Color)
                                        Dim red1 As Integer = colorValue1 Mod 256
                                        Dim green1 As Integer = (colorValue1 \ 256) Mod 256
                                        Dim blue1 As Integer = (colorValue1 \ 256 \ 256) Mod 256
                                        label1.BackColor = System.Drawing.Color.FromArgb(red1, green1, blue1)
                                    End If

                                    If IsDBNull(cell.Font.Color) Then
                                        label1.ForeColor = System.Drawing.Color.FromArgb(0, 0, 0)

                                    ElseIf Not cell.Font.ColorIndex = Excel.XlColorIndex.xlColorIndexNone Then
                                        Dim colorValue2 As Long = CLng(cell.Font.Color)
                                        Dim red2 As Integer = colorValue2 Mod 256
                                        Dim green2 As Integer = (colorValue2 \ 256) Mod 256
                                        Dim blue2 As Integer = (colorValue2 \ 256 \ 256) Mod 256
                                        label1.ForeColor = System.Drawing.Color.FromArgb(red2, green2, blue2)
                                    End If
                                End If
                                CustomPanel2.Controls.Add(label1)
                            Next
                            ordinate = ordinate + Width

                        Next
                    End If

                ElseIf X11 Then

                    Dim W As Integer

                    If TextBox3.Text = "" Then
                        W = 1
                    Else
                        W = Int(TextBox3.Text)
                    End If

                    If X1 Then
                        Dim widths(c) As Double
                        For j = 1 To c
                            widths(j - 1) = (MaxOfColumn(displayRng.Columns(j)) * BaseWidth) / 10
                        Next

                        Dim Values(0) As String
                        Dim ForFormats(0) As String
                        Dim Index As Integer = -1

                        For i = 1 To r
                            Dim source As String = displayRng.Cells(i, SplitColumn).value
                            For k = 1 To Int(Len(source) / W)
                                Index = Index + 1
                                ordinate = 0
                                For j = 1 To SplitColumn - 1
                                    Dim label1 As New System.Windows.Forms.Label
                                    label1.Text = displayRng.Cells(i, j).Value
                                    label1.Location = New System.Drawing.Point(ordinate, Index * Height)
                                    label1.Height = Height
                                    label1.Width = widths(j - 1)
                                    label1.BorderStyle = BorderStyle.FixedSingle
                                    label1.TextAlign = ContentAlignment.MiddleCenter

                                    If CheckBox1.Checked = True Then

                                        Dim cell As Excel.Range = displayRng.Cells(i, j)
                                        Dim font As Excel.Font = cell.Font

                                        Dim fontStyle As FontStyle = FontStyle.Regular
                                        If cell.Font.Bold Then fontStyle = fontStyle Or FontStyle.Bold
                                        If cell.Font.Italic Then fontStyle = fontStyle Or FontStyle.Italic

                                        Dim fontSize As Single = Convert.ToSingle(font.Size)

                                        label1.Font = New System.Drawing.Font(font.ToString, fontSize, fontStyle)
                                        If Not cell.Interior.ColorIndex = Excel.XlColorIndex.xlColorIndexNone Then
                                            Dim colorValue1 As Long = CLng(cell.Interior.Color)
                                            Dim red1 As Integer = colorValue1 Mod 256
                                            Dim green1 As Integer = (colorValue1 \ 256) Mod 256
                                            Dim blue1 As Integer = (colorValue1 \ 256 \ 256) Mod 256
                                            label1.BackColor = System.Drawing.Color.FromArgb(red1, green1, blue1)
                                        End If

                                        If IsDBNull(cell.Font.Color) Then
                                            label1.ForeColor = System.Drawing.Color.FromArgb(0, 0, 0)

                                        ElseIf Not cell.Font.ColorIndex = Excel.XlColorIndex.xlColorIndexNone Then
                                            Dim colorValue2 As Long = CLng(cell.Font.Color)
                                            Dim red2 As Integer = colorValue2 Mod 256
                                            Dim green2 As Integer = (colorValue2 \ 256) Mod 256
                                            Dim blue2 As Integer = (colorValue2 \ 256 \ 256) Mod 256
                                            label1.ForeColor = System.Drawing.Color.FromArgb(red2, green2, blue2)
                                        End If
                                    End If
                                    CustomPanel2.Controls.Add(label1)
                                    ordinate = ordinate + widths(j - 1)
                                Next j
                                ReDim Preserve Values(Index)
                                ReDim Preserve ForFormats(Index)
                                Values(Index) = Mid(source, (W * (k - 1)) + 1, W)
                                ForFormats(Index) = i
                            Next
                            If Len(source) Mod W <> 0 Then
                                Index = Index + 1
                                ordinate = 0
                                For j = 1 To SplitColumn - 1
                                    Dim label1 As New System.Windows.Forms.Label
                                    label1.Text = displayRng.Cells(i, j).Value
                                    label1.Location = New System.Drawing.Point(ordinate, Index * Height)
                                    label1.Height = Height
                                    label1.Width = widths(j - 1)
                                    label1.BorderStyle = BorderStyle.FixedSingle
                                    label1.TextAlign = ContentAlignment.MiddleCenter

                                    If CheckBox2.Checked = True Then

                                        Dim cell As Excel.Range = displayRng.Cells(i, j)
                                        Dim font As Excel.Font = cell.Font

                                        Dim fontStyle As FontStyle = FontStyle.Regular
                                        If cell.Font.Bold Then fontStyle = fontStyle Or FontStyle.Bold
                                        If cell.Font.Italic Then fontStyle = fontStyle Or FontStyle.Italic

                                        Dim fontSize As Single = Convert.ToSingle(font.Size)

                                        label1.Font = New System.Drawing.Font(font.ToString, fontSize, fontStyle)
                                        If Not cell.Interior.ColorIndex = Excel.XlColorIndex.xlColorIndexNone Then
                                            Dim colorValue1 As Long = CLng(cell.Interior.Color)
                                            Dim red1 As Integer = colorValue1 Mod 256
                                            Dim green1 As Integer = (colorValue1 \ 256) Mod 256
                                            Dim blue1 As Integer = (colorValue1 \ 256 \ 256) Mod 256
                                            label1.BackColor = System.Drawing.Color.FromArgb(red1, green1, blue1)
                                        End If

                                        If IsDBNull(cell.Font.Color) Then
                                            label1.ForeColor = System.Drawing.Color.FromArgb(0, 0, 0)

                                        ElseIf Not cell.Font.ColorIndex = Excel.XlColorIndex.xlColorIndexNone Then
                                            Dim colorValue2 As Long = CLng(cell.Font.Color)
                                            Dim red2 As Integer = colorValue2 Mod 256
                                            Dim green2 As Integer = (colorValue2 \ 256) Mod 256
                                            Dim blue2 As Integer = (colorValue2 \ 256 \ 256) Mod 256
                                            label1.ForeColor = System.Drawing.Color.FromArgb(red2, green2, blue2)
                                        End If
                                    End If
                                    CustomPanel2.Controls.Add(label1)
                                    ordinate = ordinate + widths(j - 1)
                                Next j
                                ReDim Preserve Values(Index)
                                ReDim Preserve ForFormats(Index)
                                ForFormats(Index) = i
                                Values(Index) = Mid(source, Len(source) - (Len(source) Mod W) + 1, Len(source) Mod W)
                            End If
                        Next

                        Width = (MaxOfArray(Values) * BaseWidth) / 10
                        Dim SplitOrdinate As Double
                        SplitOrdinate = ordinate

                        For i = LBound(Values) To UBound(Values)
                            ordinate = SplitOrdinate
                            Dim label1 As New System.Windows.Forms.Label
                            label1.Text = Values(i)
                            label1.Location = New System.Drawing.Point(ordinate, i * Height)
                            label1.Height = Height
                            label1.Width = Width
                            label1.BorderStyle = BorderStyle.FixedSingle
                            label1.TextAlign = ContentAlignment.MiddleCenter

                            If CheckBox1.Checked = True Then

                                Dim cell As Excel.Range = displayRng.Cells(ForFormats(i), SplitColumn)
                                Dim font As Excel.Font = cell.Font

                                Dim fontStyle As FontStyle = FontStyle.Regular
                                If cell.Font.Bold Then fontStyle = fontStyle Or FontStyle.Bold
                                If cell.Font.Italic Then fontStyle = fontStyle Or FontStyle.Italic

                                Dim fontSize As Single = Convert.ToSingle(font.Size)

                                label1.Font = New System.Drawing.Font(font.ToString, fontSize, fontStyle)
                                If Not cell.Interior.ColorIndex = Excel.XlColorIndex.xlColorIndexNone Then
                                    Dim colorValue1 As Long = CLng(cell.Interior.Color)
                                    Dim red1 As Integer = colorValue1 Mod 256
                                    Dim green1 As Integer = (colorValue1 \ 256) Mod 256
                                    Dim blue1 As Integer = (colorValue1 \ 256 \ 256) Mod 256
                                    label1.BackColor = System.Drawing.Color.FromArgb(red1, green1, blue1)
                                End If

                                If IsDBNull(cell.Font.Color) Then
                                    label1.ForeColor = System.Drawing.Color.FromArgb(0, 0, 0)

                                ElseIf Not cell.Font.ColorIndex = Excel.XlColorIndex.xlColorIndexNone Then
                                    Dim colorValue2 As Long = CLng(cell.Font.Color)
                                    Dim red2 As Integer = colorValue2 Mod 256
                                    Dim green2 As Integer = (colorValue2 \ 256) Mod 256
                                    Dim blue2 As Integer = (colorValue2 \ 256 \ 256) Mod 256
                                    label1.ForeColor = System.Drawing.Color.FromArgb(red2, green2, blue2)
                                End If
                            End If
                            CustomPanel2.Controls.Add(label1)
                            ordinate = ordinate + Width

                            For j = SplitColumn + 1 To c
                                Dim label2 As New System.Windows.Forms.Label
                                label2.Text = displayRng.Cells(Int(i / 2) + 1, j).value
                                label2.Location = New System.Drawing.Point(ordinate, i * Height)
                                label2.Height = Height
                                label2.Width = widths(j - 1)
                                label2.BorderStyle = BorderStyle.FixedSingle
                                label2.TextAlign = ContentAlignment.MiddleCenter

                                If CheckBox1.Checked = True Then

                                    Dim cell As Excel.Range = displayRng.Cells(Int(i / 2) + 1, c)
                                    Dim font As Excel.Font = cell.Font

                                    Dim fontStyle As FontStyle = FontStyle.Regular
                                    If cell.Font.Bold Then fontStyle = fontStyle Or FontStyle.Bold
                                    If cell.Font.Italic Then fontStyle = fontStyle Or FontStyle.Italic

                                    Dim fontSize As Single = Convert.ToSingle(font.Size)

                                    label2.Font = New System.Drawing.Font(font.ToString, fontSize, fontStyle)
                                    If Not cell.Interior.ColorIndex = Excel.XlColorIndex.xlColorIndexNone Then
                                        Dim colorValue1 As Long = CLng(cell.Interior.Color)
                                        Dim red1 As Integer = colorValue1 Mod 256
                                        Dim green1 As Integer = (colorValue1 \ 256) Mod 256
                                        Dim blue1 As Integer = (colorValue1 \ 256 \ 256) Mod 256
                                        label2.BackColor = System.Drawing.Color.FromArgb(red1, green1, blue1)
                                    End If

                                    If IsDBNull(cell.Font.Color) Then
                                        label2.ForeColor = System.Drawing.Color.FromArgb(0, 0, 0)

                                    ElseIf Not cell.Font.ColorIndex = Excel.XlColorIndex.xlColorIndexNone Then
                                        Dim colorValue2 As Long = CLng(cell.Font.Color)
                                        Dim red2 As Integer = colorValue2 Mod 256
                                        Dim green2 As Integer = (colorValue2 \ 256) Mod 256
                                        Dim blue2 As Integer = (colorValue2 \ 256 \ 256) Mod 256
                                        label2.ForeColor = System.Drawing.Color.FromArgb(red2, green2, blue2)
                                    End If
                                End If
                                CustomPanel2.Controls.Add(label2)
                                ordinate = ordinate + widths(j - 1)
                            Next

                        Next

                    ElseIf X2 Then

                        If c <= 4 Then
                            Height = CustomPanel2.Height / c
                        Else
                            Height = (119 / 4)
                        End If

                        Dim Index As Integer
                        ordinate = 0
                        For i = 1 To r
                            Dim source As String = displayRng.Cells(i, SplitColumn).value
                            Dim values(c - 1) As String
                            Index = -1
                            For j = 1 To c - 1
                                Index = Index + 1
                                values(j - 1) = displayRng.Cells(i, j).value
                            Next
                            For k = 1 To Int(Len(source) / W)
                                values(SplitColumn - 1) = Mid(source, (W * (k - 1)) + 1, W)
                                Width = (MaxOfArray(values) * BaseWidth) / 10
                                For m = LBound(values) To UBound(values)
                                    Dim label1 As New System.Windows.Forms.Label
                                    label1.Text = values(m)
                                    label1.Location = New System.Drawing.Point(ordinate, m * Height)
                                    label1.Height = Height
                                    label1.Width = Width
                                    label1.BorderStyle = BorderStyle.FixedSingle
                                    label1.TextAlign = ContentAlignment.MiddleCenter

                                    If CheckBox1.Checked = True Then

                                        Dim cell As Excel.Range = displayRng.Cells(i, m + 1)
                                        Dim font As Excel.Font = cell.Font

                                        Dim fontStyle As FontStyle = FontStyle.Regular
                                        If cell.Font.Bold Then fontStyle = fontStyle Or FontStyle.Bold
                                        If cell.Font.Italic Then fontStyle = fontStyle Or FontStyle.Italic

                                        Dim fontSize As Single = Convert.ToSingle(font.Size)

                                        label1.Font = New System.Drawing.Font(font.ToString, fontSize, fontStyle)
                                        If Not cell.Interior.ColorIndex = Excel.XlColorIndex.xlColorIndexNone Then
                                            Dim colorValue1 As Long = CLng(cell.Interior.Color)
                                            Dim red1 As Integer = colorValue1 Mod 256
                                            Dim green1 As Integer = (colorValue1 \ 256) Mod 256
                                            Dim blue1 As Integer = (colorValue1 \ 256 \ 256) Mod 256
                                            label1.BackColor = System.Drawing.Color.FromArgb(red1, green1, blue1)
                                        End If

                                        If IsDBNull(cell.Font.Color) Then
                                            label1.ForeColor = System.Drawing.Color.FromArgb(0, 0, 0)

                                        ElseIf Not cell.Font.ColorIndex = Excel.XlColorIndex.xlColorIndexNone Then
                                            Dim colorValue2 As Long = CLng(cell.Font.Color)
                                            Dim red2 As Integer = colorValue2 Mod 256
                                            Dim green2 As Integer = (colorValue2 \ 256) Mod 256
                                            Dim blue2 As Integer = (colorValue2 \ 256 \ 256) Mod 256
                                            label1.ForeColor = System.Drawing.Color.FromArgb(red2, green2, blue2)
                                        End If
                                    End If
                                    CustomPanel2.Controls.Add(label1)
                                Next
                                ordinate = ordinate + Width
                            Next
                            If Len(source) Mod W <> 0 Then
                                values(SplitColumn - 1) = Mid(source, Len(source) - (Len(source) Mod W) + 1, Len(source) Mod W)
                                Width = (MaxOfArray(values) * BaseWidth) / 10
                                For m = LBound(values) To UBound(values)
                                    Dim label1 As New System.Windows.Forms.Label
                                    label1.Text = values(m)
                                    label1.Location = New System.Drawing.Point(ordinate, m * Height)
                                    label1.Height = Height
                                    label1.Width = Width
                                    label1.BorderStyle = BorderStyle.FixedSingle
                                    label1.TextAlign = ContentAlignment.MiddleCenter

                                    If CheckBox2.Checked = True Then

                                        Dim cell As Excel.Range = displayRng.Cells(i, m + 1)
                                        Dim font As Excel.Font = cell.Font

                                        Dim fontStyle As FontStyle = FontStyle.Regular
                                        If cell.Font.Bold Then fontStyle = fontStyle Or FontStyle.Bold
                                        If cell.Font.Italic Then fontStyle = fontStyle Or FontStyle.Italic

                                        Dim fontSize As Single = Convert.ToSingle(font.Size)

                                        label1.Font = New System.Drawing.Font(font.ToString, fontSize, fontStyle)
                                        If Not cell.Interior.ColorIndex = Excel.XlColorIndex.xlColorIndexNone Then
                                            Dim colorValue1 As Long = CLng(cell.Interior.Color)
                                            Dim red1 As Integer = colorValue1 Mod 256
                                            Dim green1 As Integer = (colorValue1 \ 256) Mod 256
                                            Dim blue1 As Integer = (colorValue1 \ 256 \ 256) Mod 256
                                            label1.BackColor = System.Drawing.Color.FromArgb(red1, green1, blue1)
                                        End If

                                        If IsDBNull(cell.Font.Color) Then
                                            label1.ForeColor = System.Drawing.Color.FromArgb(0, 0, 0)

                                        ElseIf Not cell.Font.ColorIndex = Excel.XlColorIndex.xlColorIndexNone Then
                                            Dim colorValue2 As Long = CLng(cell.Font.Color)
                                            Dim red2 As Integer = colorValue2 Mod 256
                                            Dim green2 As Integer = (colorValue2 \ 256) Mod 256
                                            Dim blue2 As Integer = (colorValue2 \ 256 \ 256) Mod 256
                                            label1.ForeColor = System.Drawing.Color.FromArgb(red2, green2, blue2)
                                        End If
                                    End If
                                    CustomPanel2.Controls.Add(label1)
                                Next
                                ordinate = ordinate + Width
                            End If
                        Next
                    End If

                End If

                CustomPanel2.AutoScroll = True

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Try

            If TextBox1.Text = "" Then
                MessageBox.Show("Select a Source Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                TextBox1.Focus()
                workSheet.Activate()
                rng.Select()
                Exit Sub
            End If

            If IsValidExcelCellReference(TextBox1.Text) = False Then
                MessageBox.Show("Select a Valid Source Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                TextBox1.Focus()
                workSheet.Activate()
                rng.Select()
                Exit Sub
            End If

            If RadioButton4.Checked = False And RadioButton5.Checked = False Then
                MessageBox.Show("Enter a Destination Cell.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                workSheet.Activate()
                rng.Select()
                Exit Sub
            End If

            If RadioButton4.Checked = True And (TextBox4.Text = "" Or IsValidExcelCellReference(TextBox4.Text) = False) Then
                MessageBox.Show("Enter a valid Destination Cell.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                workSheet.Activate()
                rng.Select()
                Exit Sub
            End If

            If CheckBox2.Checked = True Then
                workSheet.Copy(After:=workBook.Sheets(workSheet.Name))
            End If


            Dim X1 As Boolean = RadioButton1.Checked
            Dim X2 As Boolean = RadioButton2.Checked
            Dim X3 As Boolean = RadioButton3.Checked
            Dim X7 As Boolean = RadioButton7.Checked
            Dim X8 As Boolean = RadioButton8.Checked
            Dim X9 As Boolean = RadioButton9.Checked
            Dim X10 As Boolean = RadioButton10.Checked
            Dim X11 As Boolean = RadioButton11.Checked
            Dim X12 As Boolean = ComboBox3.SelectedIndex <> -1

            If X12 = False Then
                MessageBox.Show("Select a Column by Which You Want to Split the Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                workSheet.Activate()
                rng.Select()
                Exit Sub
            End If

            If X1 = False And X2 = False Then
                MessageBox.Show("Select a Split Option.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                workSheet.Activate()
                rng.Select()
                Exit Sub
            End If

            If X3 = False And X7 = False And X8 = False And X9 = False And X10 = False And X11 = False Then
                MessageBox.Show("Select a Separator to Split the Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                workSheet.Activate()
                rng.Select()
                Exit Sub
            End If

            Dim r As Integer = rng.Rows.Count
            Dim c As Integer = rng.Columns.Count

            If (X1 Or X2) And X12 And (X3 Or X7 Or X8 Or X9 Or X10 Or X11) Then

                Dim TotalRows As Integer = 0
                Dim SplitColumn As Integer = ComboBox3.SelectedIndex + 1
                Dim Separator As String = ""
                If X7 Then
                    Separator = ";"
                ElseIf X8 Then
                    Separator = vbNewLine
                ElseIf X9 Then
                    Separator = " "
                ElseIf X10 Then
                    Separator = ComboBox2.Text
                End If

                For i = 1 To r
                    TotalRows = TotalRows + CountSeparator(rng.Cells(i, SplitColumn).value, Separator)
                Next

                If X1 Then
                    rng2 = workSheet2.Range(rng2.Cells(1, 1), rng2.Cells(TotalRows, c))
                Else
                    rng2 = workSheet2.Range(rng2.Cells(1, 1), rng2.Cells(c, TotalRows))
                End If

                If Overlap(excelApp, workSheet, workSheet2, rng, rng2) = False Then

                    Dim rng2Address As String = rng2.Address


                    If X7 Or X8 Or X9 Or X10 Then

                        If X1 Then

                            Dim Index As Integer = 0
                            Dim position As Integer

                            For i = 1 To r
                                Dim source As String = rng.Cells(i, SplitColumn).value
                                position = 1
                                For k = 1 To Len(source)
                                    If Mid(source, k, Len(Separator)) = Separator Then
                                        If k - position > 0 Then
                                            Index = Index + 1
                                            For j = 1 To SplitColumn - 1
                                                rng2.Cells(Index, j).value = rng.Cells(i, j).value
                                                If CheckBox1.Checked = True Then
                                                    rng.Cells(i, j).copy
                                                    rng2.Cells(Index, j).PasteSpecial(Excel.XlPasteType.xlPasteFormats)
                                                    rng2 = workSheet2.Range(rng2Address)
                                                    workSheet2.Activate()
                                                Else
                                                    rng2.Cells(Index, j).ClearFormats()
                                                End If
                                            Next j

                                            rng2.Cells(Index, SplitColumn) = Mid(source, position, k - position)
                                            If CheckBox1.Checked = True Then
                                                rng.Cells(i, SplitColumn).copy
                                                rng2.Cells(Index, SplitColumn).PasteSpecial(Excel.XlPasteType.xlPasteFormats)
                                                rng2 = workSheet2.Range(rng2Address)
                                                workSheet2.Activate()
                                            Else
                                                rng2.Cells(Index, SplitColumn).ClearFormats()
                                            End If

                                            For j = SplitColumn + 1 To c
                                                rng2.Cells(Index, j).value = rng.Cells(i, j).value
                                                If CheckBox1.Checked = True Then
                                                    rng.Cells(i, j).copy
                                                    rng2.Cells(Index, j).PasteSpecial(Excel.XlPasteType.xlPasteFormats)
                                                    rng2 = workSheet2.Range(rng2Address)
                                                    workSheet2.Activate()
                                                Else
                                                    rng2.Cells(Index, j).ClearFormats()
                                                End If
                                            Next j

                                        End If
                                        position = k + Len(Separator)
                                    End If
                                Next
                                If position <= Len(source) Then
                                    Index = Index + 1

                                    For j = 1 To SplitColumn - 1
                                        rng2.Cells(Index, j).Value = rng.Cells(i, j).value
                                        If CheckBox1.Checked = True Then
                                            rng.Cells(i, j).copy
                                            rng2.Cells(Index, j).PasteSpecial(Excel.XlPasteType.xlPasteFormats)
                                            rng2 = workSheet2.Range(rng2Address)
                                            workSheet2.Activate()
                                        Else
                                            rng2.Cells(Index, j).ClearFormats()
                                        End If
                                    Next j

                                    rng2.Cells(Index, SplitColumn) = Mid(source, position, Len(source) - position + 1)
                                    If CheckBox1.Checked = True Then
                                        rng.Cells(i, SplitColumn).copy
                                        rng2.Cells(Index, SplitColumn).PasteSpecial(Excel.XlPasteType.xlPasteFormats)
                                        rng2 = workSheet2.Range(rng2Address)
                                        workSheet2.Activate()
                                    Else
                                        rng2.Cells(Index, SplitColumn).ClearFormats()
                                    End If

                                    For j = SplitColumn + 1 To c
                                        rng2.Cells(Index, j).value = rng.Cells(i, j).value
                                        If CheckBox1.Checked = True Then
                                            rng.Cells(i, j).copy
                                            rng2.Cells(Index, j).PasteSpecial(Excel.XlPasteType.xlPasteFormats)
                                            rng2 = workSheet2.Range(rng2Address)
                                            workSheet2.Activate()
                                        Else
                                            rng2.Cells(Index, j).ClearFormats()
                                        End If
                                    Next j

                                End If
                            Next
                            excelApp.CutCopyMode = False

                        ElseIf X2 Then

                            Dim Index As Integer = 0
                            Dim position As Integer

                            For i = 1 To r
                                Dim source As String = rng.Cells(i, SplitColumn).value
                                position = 1
                                For k = 1 To Len(source)
                                    If Mid(source, k, Len(Separator)) = Separator Then
                                        If k - position > 0 Then
                                            Index = Index + 1
                                            For j = 1 To SplitColumn - 1
                                                rng2.Cells(j, Index).value = rng.Cells(i, j).value
                                                If CheckBox1.Checked = True Then
                                                    rng.Cells(i, j).copy
                                                    rng2.Cells(j, Index).PasteSpecial(Excel.XlPasteType.xlPasteFormats)
                                                    rng2 = workSheet2.Range(rng2Address)
                                                    workSheet2.Activate()
                                                Else
                                                    rng2.Cells(j, Index).ClearFormats()
                                                End If
                                            Next j

                                            rng2.Cells(SplitColumn, Index) = Mid(source, position, k - position)
                                            If CheckBox1.Checked = True Then
                                                rng.Cells(i, SplitColumn).copy
                                                rng2.Cells(SplitColumn, Index).PasteSpecial(Excel.XlPasteType.xlPasteFormats)
                                                rng2 = workSheet2.Range(rng2Address)
                                                workSheet2.Activate()
                                            Else
                                                rng2.Cells(SplitColumn, Index).ClearFormats()
                                            End If

                                            For j = SplitColumn + 1 To c
                                                rng2.Cells(j, Index).value = rng.Cells(i, j).value
                                                If CheckBox1.Checked = True Then
                                                    rng.Cells(i, j).copy
                                                    rng2.Cells(j, Index).PasteSpecial(Excel.XlPasteType.xlPasteFormats)
                                                    rng2 = workSheet2.Range(rng2Address)
                                                    workSheet2.Activate()
                                                Else
                                                    rng2.Cells(j, Index).ClearFormats()
                                                End If
                                            Next j
                                        End If
                                        position = k + Len(Separator)
                                    End If
                                Next
                                If position <= Len(source) Then
                                    Index = Index + 1
                                    For j = 1 To SplitColumn - 1
                                        rng2.Cells(j, Index).Value = rng.Cells(i, j).value
                                        If CheckBox1.Checked = True Then
                                            rng.Cells(i, j).copy
                                            rng2.Cells(j, Index).PasteSpecial(Excel.XlPasteType.xlPasteFormats)
                                            rng2 = workSheet2.Range(rng2Address)
                                            workSheet2.Activate()
                                        Else
                                            rng2.Cells(j, Index).ClearFormats()
                                        End If
                                    Next j

                                    rng2.Cells(SplitColumn, Index) = Mid(source, position, Len(source) - position + 1)
                                    If CheckBox1.Checked = True Then
                                        rng.Cells(i, SplitColumn).copy
                                        rng2.Cells(SplitColumn, Index).PasteSpecial(Excel.XlPasteType.xlPasteFormats)
                                        rng2 = workSheet2.Range(rng2Address)
                                        workSheet2.Activate()
                                    Else
                                        rng2.Cells(SplitColumn, Index).ClearFormats()
                                    End If

                                    For j = SplitColumn + 1 To c
                                        rng2.Cells(j, Index).value = rng.Cells(i, j).value
                                        If CheckBox1.Checked = True Then
                                            rng.Cells(i, j).copy
                                            rng2.Cells(j, Index).PasteSpecial(Excel.XlPasteType.xlPasteFormats)
                                            rng2 = workSheet2.Range(rng2Address)
                                            workSheet2.Activate()
                                        Else
                                            rng2.Cells(j, Index).ClearFormats()
                                        End If
                                    Next j
                                End If
                            Next
                            excelApp.CutCopyMode = False
                        End If

                    ElseIf X3 Then

                        If X1 Then

                            Dim Index As Integer = 0

                            For i = 1 To r

                                Dim source As String = rng.Cells(i, SplitColumn).value
                                Dim NumberText(1) As String
                                NumberText = SeparateNumberText(source)
                                Dim Number As String = NumberText(0)
                                Dim Text As String = NumberText(1)

                                Index = Index + 1
                                For j = 1 To SplitColumn - 1
                                    rng2.Cells(i, j).value = rng.Cells(i, j).value
                                    If CheckBox1.Checked = True Then
                                        rng.Cells(i, j).copy
                                        rng2.Cells(Index, j).PasteSpecial(Excel.XlPasteType.xlPasteFormats)
                                        rng2 = workSheet2.Range(rng2Address)
                                        workSheet2.Activate()
                                    Else
                                        rng2.Cells(Index, j).ClearFormats()
                                    End If
                                Next j

                                rng2.Cells(Index, SplitColumn).value = Number
                                If CheckBox1.Checked = True Then
                                    rng.Cells(i, SplitColumn).copy
                                    rng2.Cells(Index, SplitColumn).PasteSpecial(Excel.XlPasteType.xlPasteFormats)
                                    rng2 = workSheet2.Range(rng2Address)
                                    workSheet2.Activate()
                                Else
                                    rng2.Cells(Index, SplitColumn).ClearFormats()
                                End If

                                For j = SplitColumn + 1 To c
                                    rng2.Cells(i, j).value = rng.Cells(i, j).value
                                    If CheckBox1.Checked = True Then
                                        rng.Cells(i, j).copy
                                        rng2.Cells(Index, j).PasteSpecial(Excel.XlPasteType.xlPasteFormats)
                                        rng2 = workSheet2.Range(rng2Address)
                                        workSheet2.Activate()
                                    Else
                                        rng2.Cells(Index, j).ClearFormats()
                                    End If
                                Next j

                                Index = Index + 1
                                For j = 1 To SplitColumn - 1
                                    rng2.Cells(i, j).value = rng.Cells(i, j).value
                                    If CheckBox1.Checked = True Then
                                        rng.Cells(i, j).copy
                                        rng2.Cells(Index, j).PasteSpecial(Excel.XlPasteType.xlPasteFormats)
                                        rng2 = workSheet2.Range(rng2Address)
                                        workSheet2.Activate()
                                    Else
                                        rng2.Cells(Index, j).ClearFormats()
                                    End If
                                Next j

                                rng2.Cells(Index, SplitColumn).value = Text
                                If CheckBox1.Checked = True Then
                                    rng.Cells(i, SplitColumn).copy
                                    rng2.Cells(Index, SplitColumn).PasteSpecial(Excel.XlPasteType.xlPasteFormats)
                                    rng2 = workSheet2.Range(rng2Address)
                                    workSheet2.Activate()
                                Else
                                    rng2.Cells(Index, SplitColumn).ClearFormats()
                                End If

                                For j = SplitColumn + 1 To c
                                    rng2.Cells(i, j).value = rng.Cells(i, j).value
                                    If CheckBox1.Checked = True Then
                                        rng.Cells(i, j).copy
                                        rng2.Cells(Index, j).PasteSpecial(Excel.XlPasteType.xlPasteFormats)
                                        rng2 = workSheet2.Range(rng2Address)
                                        workSheet2.Activate()
                                    Else
                                        rng2.Cells(Index, j).ClearFormats()
                                    End If
                                Next j

                            Next
                            excelApp.CutCopyMode = False

                        ElseIf X2 Then

                            Dim Index As Integer = 0

                            For i = 1 To r

                                Dim source As String = rng.Cells(i, SplitColumn).value
                                Dim NumberText(1) As String
                                NumberText = SeparateNumberText(source)
                                Dim Number As String = NumberText(0)
                                Dim Text As String = NumberText(1)

                                Index = Index + 1
                                For j = 1 To c - 1
                                    rng2.Cells(j, Index).value = rng.Cells(i, j).value
                                    If CheckBox1.Checked = True Then
                                        rng.Cells(i, j).copy
                                        rng2.Cells(j, Index).PasteSpecial(Excel.XlPasteType.xlPasteFormats)
                                        rng2 = workSheet2.Range(rng2Address)
                                        workSheet2.Activate()
                                    Else
                                        rng2.Cells(j, Index).ClearFormats()
                                    End If
                                Next j

                                rng2.Cells(SplitColumn, Index).value = Number
                                If CheckBox1.Checked = True Then
                                    rng.Cells(i, SplitColumn).copy
                                    rng2.Cells(SplitColumn, Index).PasteSpecial(Excel.XlPasteType.xlPasteFormats)
                                    rng2 = workSheet2.Range(rng2Address)
                                    workSheet2.Activate()
                                Else
                                    rng2.Cells(SplitColumn, Index).ClearFormats()
                                End If

                                For j = SplitColumn + 1 To c
                                    rng2.Cells(j, Index).value = rng.Cells(i, j).value
                                    If CheckBox1.Checked = True Then
                                        rng.Cells(i, j).copy
                                        rng2.Cells(j, Index).PasteSpecial(Excel.XlPasteType.xlPasteFormats)
                                        rng2 = workSheet2.Range(rng2Address)
                                        workSheet2.Activate()
                                    Else
                                        rng2.Cells(j, Index).ClearFormats()
                                    End If
                                Next j

                                Index = Index + 1
                                For j = 1 To SplitColumn - 1
                                    rng2.Cells(j, Index).value = rng.Cells(i, j).value
                                    If CheckBox1.Checked = True Then
                                        rng.Cells(i, j).copy
                                        rng2.Cells(j, Index).PasteSpecial(Excel.XlPasteType.xlPasteFormats)
                                        rng2 = workSheet2.Range(rng2Address)
                                        workSheet2.Activate()
                                    Else
                                        rng2.Cells(j, Index).ClearFormats()
                                    End If
                                Next j

                                rng2.Cells(SplitColumn, Index).value = Text
                                If CheckBox1.Checked = True Then
                                    rng.Cells(i, SplitColumn).copy
                                    rng2.Cells(SplitColumn, Index).PasteSpecial(Excel.XlPasteType.xlPasteFormats)
                                    rng2 = workSheet2.Range(rng2Address)
                                    workSheet2.Activate()
                                Else
                                    rng2.Cells(SplitColumn, Index).ClearFormats()
                                End If

                                For j = SplitColumn + 1 To c
                                    rng2.Cells(j, Index).value = rng.Cells(i, j).value
                                    If CheckBox1.Checked = True Then
                                        rng.Cells(i, j).copy
                                        rng2.Cells(j, Index).PasteSpecial(Excel.XlPasteType.xlPasteFormats)
                                        rng2 = workSheet2.Range(rng2Address)
                                        workSheet2.Activate()
                                    Else
                                        rng2.Cells(j, Index).ClearFormats()
                                    End If
                                Next j
                            Next

                            excelApp.CutCopyMode = False

                        End If

                    ElseIf X11 Then

                        Dim W As Integer

                        If TextBox3.Text = "" Then
                            W = 1
                        Else
                            W = Int(TextBox3.Text)
                        End If

                        If X1 Then

                            Dim Index As Integer = 0

                            For i = 1 To r
                                Dim source As String = rng.Cells(i, SplitColumn).value
                                For k = 1 To Int(Len(source) / W)
                                    Index = Index + 1
                                    For j = 1 To SplitColumn - 1
                                        rng2.Cells(Index, j).value = rng.Cells(i, j).value
                                        If CheckBox1.Checked = True Then
                                            rng.Cells(i, j).copy
                                            rng2.Cells(Index, j).PasteSpecial(Excel.XlPasteType.xlPasteFormats)
                                            rng2 = workSheet2.Range(rng2Address)
                                            workSheet2.Activate()
                                        Else
                                            rng2.Cells(Index, j).ClearFormats()
                                        End If
                                    Next j
                                    rng2.Cells(Index, SplitColumn).value = Mid(source, (W * (k - 1)) + 1, W)
                                    If CheckBox1.Checked = True Then
                                        rng.Cells(i, SplitColumn).copy
                                        rng2.Cells(Index, SplitColumn).PasteSpecial(Excel.XlPasteType.xlPasteFormats)
                                        rng2 = workSheet2.Range(rng2Address)
                                        workSheet2.Activate()
                                    Else
                                        rng2.Cells(Index, SplitColumn).ClearFormats()
                                    End If
                                    For j = SplitColumn + 1 To c
                                        rng2.Cells(Index, j).value = rng.Cells(i, j).value
                                        If CheckBox1.Checked = True Then
                                            rng.Cells(i, j).copy
                                            rng2.Cells(Index, j).PasteSpecial(Excel.XlPasteType.xlPasteFormats)
                                            rng2 = workSheet2.Range(rng2Address)
                                            workSheet2.Activate()
                                        Else
                                            rng2.Cells(Index, j).ClearFormats()
                                        End If
                                    Next j
                                Next
                                If Len(source) Mod W <> 0 Then
                                    Index = Index + 1
                                    For j = 1 To SplitColumn - 1
                                        rng2.Cells(Index, j).value = rng.Cells(i, j).value
                                        If CheckBox1.Checked = True Then
                                            rng.Cells(i, j).copy
                                            rng2.Cells(Index, j).PasteSpecial(Excel.XlPasteType.xlPasteFormats)
                                            rng2 = workSheet2.Range(rng2Address)
                                            workSheet2.Activate()
                                        Else
                                            rng2.Cells(Index, j).ClearFormats()
                                        End If
                                    Next j
                                    rng2.Cells(Index, SplitColumn).value = Mid(source, Len(source) - (Len(source) Mod W) + 1, Len(source) Mod W)
                                    If CheckBox1.Checked = True Then
                                        rng.Cells(i, SplitColumn).copy
                                        rng2.Cells(Index, SplitColumn).PasteSpecial(Excel.XlPasteType.xlPasteFormats)
                                        rng2 = workSheet2.Range(rng2Address)
                                        workSheet2.Activate()
                                    Else
                                        rng2.Cells(Index, SplitColumn).ClearFormats()
                                    End If
                                    For j = SplitColumn + 1 To c
                                        rng2.Cells(Index, j).value = rng.Cells(i, j).value
                                        If CheckBox1.Checked = True Then
                                            rng.Cells(i, j).copy
                                            rng2.Cells(Index, j).PasteSpecial(Excel.XlPasteType.xlPasteFormats)
                                            rng2 = workSheet2.Range(rng2Address)
                                            workSheet2.Activate()
                                        Else
                                            rng2.Cells(Index, j).ClearFormats()
                                        End If
                                    Next j
                                End If
                            Next

                            excelApp.CutCopyMode = False

                        ElseIf X2 Then

                            Dim Index As Integer = 0

                            For i = 1 To r
                                Dim source As String = rng.Cells(i, SplitColumn).value
                                For k = 1 To Int(Len(source) / W)
                                    Index = Index + 1
                                    For j = 1 To SplitColumn - 1
                                        rng2.Cells(j, Index).value = rng.Cells(i, j).value
                                        If CheckBox1.Checked = True Then
                                            rng.Cells(i, j).copy
                                            rng2.Cells(j, Index).PasteSpecial(Excel.XlPasteType.xlPasteFormats)
                                            rng2 = workSheet2.Range(rng2Address)
                                            workSheet2.Activate()
                                        Else
                                            rng2.Cells(j, Index).ClearFormats()
                                        End If
                                    Next j
                                    rng2.Cells(SplitColumn, Index).value = Mid(source, (W * (k - 1)) + 1, W)
                                    If CheckBox1.Checked = True Then
                                        rng.Cells(i, SplitColumn).copy
                                        rng2.Cells(SplitColumn, Index).PasteSpecial(Excel.XlPasteType.xlPasteFormats)
                                        rng2 = workSheet2.Range(rng2Address)
                                        workSheet2.Activate()
                                    Else
                                        rng2.Cells(SplitColumn, Index).ClearFormats()
                                    End If
                                    For j = SplitColumn + 1 To c
                                        rng2.Cells(j, Index).value = rng.Cells(i, j).value
                                        If CheckBox1.Checked = True Then
                                            rng.Cells(i, j).copy
                                            rng2.Cells(j, Index).PasteSpecial(Excel.XlPasteType.xlPasteFormats)
                                            rng2 = workSheet2.Range(rng2Address)
                                            workSheet2.Activate()
                                        Else
                                            rng2.Cells(j, Index).ClearFormats()
                                        End If
                                    Next j
                                Next
                                If Len(source) Mod W <> 0 Then
                                    Index = Index + 1
                                    For j = 1 To SplitColumn - 1
                                        rng2.Cells(j, Index).value = rng.Cells(i, j).value
                                        If CheckBox1.Checked = True Then
                                            rng.Cells(i, j).copy
                                            rng2.Cells(j, Index).PasteSpecial(Excel.XlPasteType.xlPasteFormats)
                                            rng2 = workSheet2.Range(rng2Address)
                                            workSheet2.Activate()
                                        Else
                                            rng2.Cells(j, Index).ClearFormats()
                                        End If
                                    Next j
                                    rng2.Cells(SplitColumn, Index).value = Mid(source, Len(source) - (Len(source) Mod W) + 1, Len(source) Mod W)
                                    If CheckBox1.Checked = True Then
                                        rng.Cells(i, SplitColumn).copy
                                        rng2.Cells(SplitColumn, Index).PasteSpecial(Excel.XlPasteType.xlPasteFormats)
                                        rng2 = workSheet2.Range(rng2Address)
                                        workSheet2.Activate()
                                    Else
                                        rng2.Cells(SplitColumn, Index).ClearFormats()
                                    End If
                                    For j = SplitColumn + 1 To c
                                        rng2.Cells(j, Index).value = rng.Cells(i, j).value
                                        If CheckBox1.Checked = True Then
                                            rng.Cells(i, j).copy
                                            rng2.Cells(j, Index).PasteSpecial(Excel.XlPasteType.xlPasteFormats)
                                            rng2 = workSheet2.Range(rng2Address)
                                            workSheet2.Activate()
                                        Else
                                            rng2.Cells(j, Index).ClearFormats()
                                        End If
                                    Next j
                                End If
                            Next

                            excelApp.CutCopyMode = False

                        End If

                    End If

                Else

                    Dim rng2Address As String = rng2.Address

                    Dim Arr(rng.Rows.Count - 1, rng.Columns.Count - 1) As Object

                    For i = LBound(Arr, 1) To UBound(Arr, 1)
                        For j = LBound(Arr, 2) To UBound(Arr, 2)
                            Arr(i, j) = rng.Cells(i + 1, j + 1).Value
                        Next
                    Next

                    Dim FontNames(rng.Rows.Count - 1, rng.Columns.Count - 1) As String
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

                    If X7 Or X8 Or X9 Or X10 Then

                        If X1 Then

                            Dim Index As Integer = 0
                            Dim position As Integer

                            For i = 1 To r
                                Dim source As String = Arr(i - 1, SplitColumn - 1)
                                position = 1
                                For k = 1 To Len(source)
                                    If Mid(source, k, Len(Separator)) = Separator Then
                                        If k - position > 0 Then
                                            Index = Index + 1
                                            For j = 1 To SplitColumn - 1
                                                rng2.Cells(Index, j).value = Arr(i - 1, j - 1)
                                                If CheckBox1.Checked = True Then
                                                    Dim x As Integer = i - 1
                                                    Dim y As Integer = j - 1

                                                    rng2.Cells(Index, j).Font.Name = FontNames(x, y)
                                                    rng2.Cells(Index, j).Font.Size = FontSizes(x, y)

                                                    If FontBolds(x, y) Then rng2.Cells(Index, j).Font.Bold = True
                                                    If Fontitalics(x, y) Then rng2.Cells(Index, j).Font.Italic = True


                                                    rng2.Cells(Index, j).Interior.Color = System.Drawing.Color.FromArgb(Red1s(x, y), Green1s(x, y), Blue1s(x, y))

                                                    rng2.Cells(Index, j).Font.Color = System.Drawing.Color.FromArgb(Red2s(x, y), Green2s(x, y), Blue2s(x, y))
                                                    workSheet2.Activate()
                                                Else
                                                    rng2.Cells(Index, j).ClearFormats()
                                                End If
                                            Next j

                                            rng2.Cells(Index, SplitColumn) = Mid(source, position, k - position)
                                            If CheckBox1.Checked = True Then
                                                Dim x As Integer = i - 1
                                                Dim y As Integer = SplitColumn - 1

                                                rng2.Cells(Index, SplitColumn).Font.Name = FontNames(x, y)
                                                rng2.Cells(Index, SplitColumn).Font.Size = FontSizes(x, y)

                                                If FontBolds(x, y) Then rng2.Cells(Index, SplitColumn).Font.Bold = True
                                                If Fontitalics(x, y) Then rng2.Cells(Index, SplitColumn).Font.Italic = True


                                                rng2.Cells(Index, SplitColumn).Interior.Color = System.Drawing.Color.FromArgb(Red1s(x, y), Green1s(x, y), Blue1s(x, y))

                                                rng2.Cells(Index, SplitColumn).Font.Color = System.Drawing.Color.FromArgb(Red2s(x, y), Green2s(x, y), Blue2s(x, y))
                                                workSheet2.Activate()
                                            Else
                                                rng2.Cells(Index, SplitColumn).ClearFormats()
                                            End If

                                            For j = SplitColumn + 1 To c
                                                rng2.Cells(Index, j).value = Arr(i - 1, j - 1)
                                                If CheckBox1.Checked = True Then
                                                    Dim x As Integer = i - 1
                                                    Dim y As Integer = j - 1

                                                    rng2.Cells(Index, j).Font.Name = FontNames(x, y)
                                                    rng2.Cells(Index, j).Font.Size = FontSizes(x, y)

                                                    If FontBolds(x, y) Then rng2.Cells(Index, j).Font.Bold = True
                                                    If Fontitalics(x, y) Then rng2.Cells(Index, j).Font.Italic = True


                                                    rng2.Cells(Index, j).Interior.Color = System.Drawing.Color.FromArgb(Red1s(x, y), Green1s(x, y), Blue1s(x, y))

                                                    rng2.Cells(Index, j).Font.Color = System.Drawing.Color.FromArgb(Red2s(x, y), Green2s(x, y), Blue2s(x, y))
                                                    workSheet2.Activate()
                                                Else
                                                    rng2.Cells(Index, j).ClearFormats()
                                                End If
                                            Next j

                                        End If
                                        position = k + Len(Separator)
                                    End If
                                Next
                                If position <= Len(source) Then
                                    Index = Index + 1

                                    For j = 1 To SplitColumn - 1
                                        rng2.Cells(Index, j).Value = Arr(i - 1, j - 1)
                                        If CheckBox1.Checked = True Then
                                            Dim x As Integer = i - 1
                                            Dim y As Integer = j - 1

                                            rng2.Cells(Index, j).Font.Name = FontNames(x, y)
                                            rng2.Cells(Index, j).Font.Size = FontSizes(x, y)

                                            If FontBolds(x, y) Then rng2.Cells(Index, j).Font.Bold = True
                                            If Fontitalics(x, y) Then rng2.Cells(Index, j).Font.Italic = True


                                            rng2.Cells(Index, j).Interior.Color = System.Drawing.Color.FromArgb(Red1s(x, y), Green1s(x, y), Blue1s(x, y))

                                            rng2.Cells(Index, j).Font.Color = System.Drawing.Color.FromArgb(Red2s(x, y), Green2s(x, y), Blue2s(x, y))
                                            workSheet2.Activate()
                                        Else
                                            rng2.Cells(Index, j).ClearFormats()
                                        End If
                                    Next j

                                    rng2.Cells(Index, SplitColumn) = Mid(source, position, Len(source) - position + 1)
                                    If CheckBox1.Checked = True Then
                                        Dim x As Integer = i - 1
                                        Dim y As Integer = SplitColumn - 1

                                        rng2.Cells(Index, SplitColumn).Font.Name = FontNames(x, y)
                                        rng2.Cells(Index, SplitColumn).Font.Size = FontSizes(x, y)

                                        If FontBolds(x, y) Then rng2.Cells(Index, SplitColumn).Font.Bold = True
                                        If Fontitalics(x, y) Then rng2.Cells(Index, SplitColumn).Font.Italic = True


                                        rng2.Cells(Index, SplitColumn).Interior.Color = System.Drawing.Color.FromArgb(Red1s(x, y), Green1s(x, y), Blue1s(x, y))

                                        rng2.Cells(Index, SplitColumn).Font.Color = System.Drawing.Color.FromArgb(Red2s(x, y), Green2s(x, y), Blue2s(x, y))
                                        workSheet2.Activate()
                                    Else
                                        rng2.Cells(Index, SplitColumn).ClearFormats()
                                    End If

                                    For j = SplitColumn + 1 To c
                                        rng2.Cells(Index, j).value = Arr(i - 1, j - 1)
                                        If CheckBox1.Checked = True Then
                                            Dim x As Integer = i - 1
                                            Dim y As Integer = j - 1

                                            rng2.Cells(Index, j).Font.Name = FontNames(x, y)
                                            rng2.Cells(Index, j).Font.Size = FontSizes(x, y)

                                            If FontBolds(x, y) Then rng2.Cells(Index, j).Font.Bold = True
                                            If Fontitalics(x, y) Then rng2.Cells(Index, j).Font.Italic = True


                                            rng2.Cells(Index, j).Interior.Color = System.Drawing.Color.FromArgb(Red1s(x, y), Green1s(x, y), Blue1s(x, y))

                                            rng2.Cells(Index, j).Font.Color = System.Drawing.Color.FromArgb(Red2s(x, y), Green2s(x, y), Blue2s(x, y))
                                            workSheet2.Activate()
                                        Else
                                            rng2.Cells(Index, j).ClearFormats()
                                        End If
                                    Next j

                                End If
                            Next

                        ElseIf X2 Then

                            Dim Index As Integer = 0
                            Dim position As Integer

                            For i = 1 To r
                                Dim source As String = Arr(i - 1, SplitColumn - 1)
                                position = 1
                                For k = 1 To Len(source)
                                    If Mid(source, k, Len(Separator)) = Separator Then
                                        If k - position > 0 Then
                                            Index = Index + 1
                                            For j = 1 To SplitColumn - 1
                                                rng2.Cells(j, Index).value = Arr(i - 1, j - 1)
                                                If CheckBox1.Checked = True Then
                                                    Dim x As Integer = i - 1
                                                    Dim y As Integer = j - 1

                                                    rng2.Cells(j, Index).Font.Name = FontNames(x, y)
                                                    rng2.Cells(j, Index).Font.Size = FontSizes(x, y)

                                                    If FontBolds(x, y) Then rng2.Cells(j, Index).Font.Bold = True
                                                    If Fontitalics(x, y) Then rng2.Cells(j, Index).Font.Italic = True


                                                    rng2.Cells(j, Index).Interior.Color = System.Drawing.Color.FromArgb(Red1s(x, y), Green1s(x, y), Blue1s(x, y))

                                                    rng2.Cells(j, Index).Font.Color = System.Drawing.Color.FromArgb(Red2s(x, y), Green2s(x, y), Blue2s(x, y))

                                                    workSheet2.Activate()
                                                Else
                                                    rng2.Cells(j, Index).ClearFormats()
                                                End If
                                            Next j

                                            rng2.Cells(SplitColumn, Index) = Mid(source, position, k - position)
                                            If CheckBox1.Checked = True Then
                                                Dim x As Integer = i - 1
                                                Dim y As Integer = SplitColumn - 1

                                                rng2.Cells(SplitColumn, Index).Font.Name = FontNames(x, y)
                                                rng2.Cells(SplitColumn, Index).Font.Size = FontSizes(x, y)

                                                If FontBolds(x, y) Then rng2.Cells(SplitColumn, Index).Font.Bold = True
                                                If Fontitalics(x, y) Then rng2.Cells(SplitColumn, Index).Font.Italic = True


                                                rng2.Cells(SplitColumn, Index).Interior.Color = System.Drawing.Color.FromArgb(Red1s(x, y), Green1s(x, y), Blue1s(x, y))

                                                rng2.Cells(SplitColumn, Index).Font.Color = System.Drawing.Color.FromArgb(Red2s(x, y), Green2s(x, y), Blue2s(x, y))
                                                workSheet2.Activate()
                                            Else
                                                rng2.Cells(SplitColumn, Index).ClearFormats()
                                            End If

                                            For j = SplitColumn + 1 To c
                                                rng2.Cells(j, Index).value = Arr(i - 1, j - 1)
                                                If CheckBox1.Checked = True Then
                                                    Dim x As Integer = i - 1
                                                    Dim y As Integer = j - 1

                                                    rng2.Cells(j, Index).Font.Name = FontNames(x, y)
                                                    rng2.Cells(j, Index).Font.Size = FontSizes(x, y)

                                                    If FontBolds(x, y) Then rng2.Cells(j, Index).Font.Bold = True
                                                    If Fontitalics(x, y) Then rng2.Cells(j, Index).Font.Italic = True


                                                    rng2.Cells(j, Index).Interior.Color = System.Drawing.Color.FromArgb(Red1s(x, y), Green1s(x, y), Blue1s(x, y))

                                                    rng2.Cells(j, Index).Font.Color = System.Drawing.Color.FromArgb(Red2s(x, y), Green2s(x, y), Blue2s(x, y))
                                                    workSheet2.Activate()
                                                Else
                                                    rng2.Cells(j, Index).ClearFormats()
                                                End If
                                            Next j
                                        End If
                                        position = k + Len(Separator)
                                    End If
                                Next
                                If position <= Len(source) Then
                                    Index = Index + 1
                                    For j = 1 To SplitColumn - 1
                                        rng2.Cells(j, Index).Value = Arr(i - 1, j - 1)
                                        If CheckBox1.Checked = True Then
                                            Dim x As Integer = i - 1
                                            Dim y As Integer = j - 1

                                            rng2.Cells(j, Index).Font.Name = FontNames(x, y)
                                            rng2.Cells(j, Index).Font.Size = FontSizes(x, y)

                                            If FontBolds(x, y) Then rng2.Cells(j, Index).Font.Bold = True
                                            If Fontitalics(x, y) Then rng2.Cells(j, Index).Font.Italic = True


                                            rng2.Cells(j, Index).Interior.Color = System.Drawing.Color.FromArgb(Red1s(x, y), Green1s(x, y), Blue1s(x, y))

                                            rng2.Cells(j, Index).Font.Color = System.Drawing.Color.FromArgb(Red2s(x, y), Green2s(x, y), Blue2s(x, y))
                                            workSheet2.Activate()
                                        Else
                                            rng2.Cells(j, Index).ClearFormats()
                                        End If
                                    Next j

                                    rng2.Cells(SplitColumn, Index) = Mid(source, position, Len(source) - position + 1)
                                    If CheckBox1.Checked = True Then
                                        Dim x As Integer = i - 1
                                        Dim y As Integer = SplitColumn - 1

                                        rng2.Cells(SplitColumn, Index).Font.Name = FontNames(x, y)
                                        rng2.Cells(SplitColumn, Index).Font.Size = FontSizes(x, y)

                                        If FontBolds(x, y) Then rng2.Cells(SplitColumn, Index).Font.Bold = True
                                        If Fontitalics(x, y) Then rng2.Cells(SplitColumn, Index).Font.Italic = True


                                        rng2.Cells(SplitColumn, Index).Interior.Color = System.Drawing.Color.FromArgb(Red1s(x, y), Green1s(x, y), Blue1s(x, y))

                                        rng2.Cells(SplitColumn, Index).Font.Color = System.Drawing.Color.FromArgb(Red2s(x, y), Green2s(x, y), Blue2s(x, y))
                                        workSheet2.Activate()
                                    Else
                                        rng2.Cells(SplitColumn, Index).ClearFormats()
                                    End If

                                    For j = SplitColumn + 1 To c
                                        rng2.Cells(j, Index).value = Arr(i - 1, j - 1)
                                        If CheckBox1.Checked = True Then
                                            Dim x As Integer = i - 1
                                            Dim y As Integer = j - 1

                                            rng2.Cells(j, Index).Font.Name = FontNames(x, y)
                                            rng2.Cells(j, Index).Font.Size = FontSizes(x, y)

                                            If FontBolds(x, y) Then rng2.Cells(j, Index).Font.Bold = True
                                            If Fontitalics(x, y) Then rng2.Cells(j, Index).Font.Italic = True


                                            rng2.Cells(j, Index).Interior.Color = System.Drawing.Color.FromArgb(Red1s(x, y), Green1s(x, y), Blue1s(x, y))

                                            rng2.Cells(j, Index).Font.Color = System.Drawing.Color.FromArgb(Red2s(x, y), Green2s(x, y), Blue2s(x, y))
                                            workSheet2.Activate()
                                        Else
                                            rng2.Cells(j, Index).ClearFormats()
                                        End If
                                    Next j
                                End If
                            Next
                            excelApp.CutCopyMode = False
                        End If

                    ElseIf X3 Then

                        If X1 Then

                            Dim Index As Integer = 0

                            For i = 1 To r

                                Dim source As String = Arr(i - 1, SplitColumn - 1)
                                Dim NumberText(1) As String
                                NumberText = SeparateNumberText(source)
                                Dim Number As String = NumberText(0)
                                Dim Text As String = NumberText(1)

                                Index = Index + 1
                                For j = 1 To SplitColumn - 1
                                    rng2.Cells(i, j).value = Arr(i - 1, j - 1)
                                    If CheckBox1.Checked = True Then
                                        Dim x As Integer = i - 1
                                        Dim y As Integer = j - 1

                                        rng2.Cells(Index, j).Font.Name = FontNames(x, y)
                                        rng2.Cells(Index, j).Font.Size = FontSizes(x, y)

                                        If FontBolds(x, y) Then rng2.Cells(Index, j).Font.Bold = True
                                        If Fontitalics(x, y) Then rng2.Cells(Index, j).Font.Italic = True


                                        rng2.Cells(Index, j).Interior.Color = System.Drawing.Color.FromArgb(Red1s(x, y), Green1s(x, y), Blue1s(x, y))

                                        rng2.Cells(Index, j).Font.Color = System.Drawing.Color.FromArgb(Red2s(x, y), Green2s(x, y), Blue2s(x, y))
                                        workSheet2.Activate()
                                    Else
                                        rng2.Cells(Index, j).ClearFormats()
                                    End If
                                Next j

                                rng2.Cells(Index, SplitColumn).value = Number
                                If CheckBox1.Checked = True Then
                                    Dim x As Integer = i - 1
                                    Dim y As Integer = SplitColumn - 1

                                    rng2.Cells(Index, SplitColumn).Font.Name = FontNames(x, y)
                                    rng2.Cells(Index, SplitColumn).Font.Size = FontSizes(x, y)

                                    If FontBolds(x, y) Then rng2.Cells(Index, SplitColumn).Font.Bold = True
                                    If Fontitalics(x, y) Then rng2.Cells(Index, SplitColumn).Font.Italic = True


                                    rng2.Cells(Index, SplitColumn).Interior.Color = System.Drawing.Color.FromArgb(Red1s(x, y), Green1s(x, y), Blue1s(x, y))

                                    rng2.Cells(Index, SplitColumn).Font.Color = System.Drawing.Color.FromArgb(Red2s(x, y), Green2s(x, y), Blue2s(x, y))
                                    workSheet2.Activate()
                                Else
                                    rng2.Cells(Index, SplitColumn).ClearFormats()
                                End If

                                For j = SplitColumn + 1 To c
                                    rng2.Cells(i, j).value = Arr(i - 1, j - 1)
                                    If CheckBox1.Checked = True Then
                                        Dim x As Integer = i - 1
                                        Dim y As Integer = j - 1

                                        rng2.Cells(Index, j).Font.Name = FontNames(x, y)
                                        rng2.Cells(Index, j).Font.Size = FontSizes(x, y)

                                        If FontBolds(x, y) Then rng2.Cells(Index, j).Font.Bold = True
                                        If Fontitalics(x, y) Then rng2.Cells(Index, j).Font.Italic = True


                                        rng2.Cells(Index, j).Interior.Color = System.Drawing.Color.FromArgb(Red1s(x, y), Green1s(x, y), Blue1s(x, y))

                                        rng2.Cells(Index, j).Font.Color = System.Drawing.Color.FromArgb(Red2s(x, y), Green2s(x, y), Blue2s(x, y))
                                        workSheet2.Activate()
                                    Else
                                        rng2.Cells(Index, j).ClearFormats()
                                    End If
                                Next j

                                Index = Index + 1
                                For j = 1 To SplitColumn - 1
                                    rng2.Cells(i, j).value = Arr(i - 1, j - 1)
                                    If CheckBox1.Checked = True Then
                                        Dim x As Integer = i - 1
                                        Dim y As Integer = j - 1

                                        rng2.Cells(Index, j).Font.Name = FontNames(x, y)
                                        rng2.Cells(Index, j).Font.Size = FontSizes(x, y)

                                        If FontBolds(x, y) Then rng2.Cells(Index, j).Font.Bold = True
                                        If Fontitalics(x, y) Then rng2.Cells(Index, j).Font.Italic = True


                                        rng2.Cells(Index, j).Interior.Color = System.Drawing.Color.FromArgb(Red1s(x, y), Green1s(x, y), Blue1s(x, y))

                                        rng2.Cells(Index, j).Font.Color = System.Drawing.Color.FromArgb(Red2s(x, y), Green2s(x, y), Blue2s(x, y))
                                        workSheet2.Activate()
                                    Else
                                        rng2.Cells(Index, j).ClearFormats()
                                    End If
                                Next j

                                rng2.Cells(Index, SplitColumn).value = Text
                                If CheckBox1.Checked = True Then
                                    Dim x As Integer = i - 1
                                    Dim y As Integer = SplitColumn - 1

                                    rng2.Cells(Index, SplitColumn).Font.Name = FontNames(x, y)
                                    rng2.Cells(Index, SplitColumn).Font.Size = FontSizes(x, y)

                                    If FontBolds(x, y) Then rng2.Cells(Index, SplitColumn).Font.Bold = True
                                    If Fontitalics(x, y) Then rng2.Cells(Index, SplitColumn).Font.Italic = True


                                    rng2.Cells(Index, SplitColumn).Interior.Color = System.Drawing.Color.FromArgb(Red1s(x, y), Green1s(x, y), Blue1s(x, y))

                                    rng2.Cells(Index, SplitColumn).Font.Color = System.Drawing.Color.FromArgb(Red2s(x, y), Green2s(x, y), Blue2s(x, y))
                                    workSheet2.Activate()
                                Else
                                    rng2.Cells(Index, SplitColumn).ClearFormats()
                                End If

                                For j = SplitColumn + 1 To c
                                    rng2.Cells(i, j).value = Arr(i - 1, j - 1)
                                    If CheckBox1.Checked = True Then
                                        Dim x As Integer = i - 1
                                        Dim y As Integer = j - 1

                                        rng2.Cells(Index, j).Font.Name = FontNames(x, y)
                                        rng2.Cells(Index, j).Font.Size = FontSizes(x, y)

                                        If FontBolds(x, y) Then rng2.Cells(Index, j).Font.Bold = True
                                        If Fontitalics(x, y) Then rng2.Cells(Index, j).Font.Italic = True


                                        rng2.Cells(Index, j).Interior.Color = System.Drawing.Color.FromArgb(Red1s(x, y), Green1s(x, y), Blue1s(x, y))

                                        rng2.Cells(Index, j).Font.Color = System.Drawing.Color.FromArgb(Red2s(x, y), Green2s(x, y), Blue2s(x, y))
                                        workSheet2.Activate()
                                    Else
                                        rng2.Cells(Index, j).ClearFormats()
                                    End If
                                Next j

                            Next
                            excelApp.CutCopyMode = False

                        ElseIf X2 Then

                            Dim Index As Integer = 0

                            For i = 1 To r

                                Dim source As String = Arr(i - 1, SplitColumn - 1)
                                Dim NumberText(1) As String
                                NumberText = SeparateNumberText(source)
                                Dim Number As String = NumberText(0)
                                Dim Text As String = NumberText(1)

                                Index = Index + 1
                                For j = 1 To c - 1
                                    rng2.Cells(j, Index).value = Arr(i - 1, j - 1)
                                    If CheckBox1.Checked = True Then
                                        Dim x As Integer = i - 1
                                        Dim y As Integer = j - 1

                                        rng2.Cells(j, Index).Font.Name = FontNames(x, y)
                                        rng2.Cells(j, Index).Font.Size = FontSizes(x, y)

                                        If FontBolds(x, y) Then rng2.Cells(j, Index).Font.Bold = True
                                        If Fontitalics(x, y) Then rng2.Cells(j, Index).Font.Italic = True


                                        rng2.Cells(j, Index).Interior.Color = System.Drawing.Color.FromArgb(Red1s(x, y), Green1s(x, y), Blue1s(x, y))

                                        rng2.Cells(j, Index).Font.Color = System.Drawing.Color.FromArgb(Red2s(x, y), Green2s(x, y), Blue2s(x, y))
                                        workSheet2.Activate()
                                    Else
                                        rng2.Cells(j, Index).ClearFormats()
                                    End If
                                Next j

                                rng2.Cells(SplitColumn, Index).value = Number
                                If CheckBox1.Checked = True Then
                                    Dim x As Integer = i - 1
                                    Dim y As Integer = SplitColumn - 1

                                    rng2.Cells(SplitColumn, Index).Font.Name = FontNames(x, y)
                                    rng2.Cells(SplitColumn, Index).Font.Size = FontSizes(x, y)

                                    If FontBolds(x, y) Then rng2.Cells(SplitColumn, Index).Font.Bold = True
                                    If Fontitalics(x, y) Then rng2.Cells(SplitColumn, Index).Font.Italic = True


                                    rng2.Cells(SplitColumn, Index).Interior.Color = System.Drawing.Color.FromArgb(Red1s(x, y), Green1s(x, y), Blue1s(x, y))

                                    rng2.Cells(SplitColumn, Index).Font.Color = System.Drawing.Color.FromArgb(Red2s(x, y), Green2s(x, y), Blue2s(x, y))
                                    workSheet2.Activate()
                                Else
                                    rng2.Cells(SplitColumn, Index).ClearFormats()
                                End If

                                For j = SplitColumn + 1 To c
                                    rng2.Cells(j, Index).value = Arr(i - 1, j - 1)
                                    If CheckBox1.Checked = True Then
                                        Dim x As Integer = i - 1
                                        Dim y As Integer = j - 1

                                        rng2.Cells(j, Index).Font.Name = FontNames(x, y)
                                        rng2.Cells(j, Index).Font.Size = FontSizes(x, y)

                                        If FontBolds(x, y) Then rng2.Cells(j, Index).Font.Bold = True
                                        If Fontitalics(x, y) Then rng2.Cells(j, Index).Font.Italic = True


                                        rng2.Cells(j, Index).Interior.Color = System.Drawing.Color.FromArgb(Red1s(x, y), Green1s(x, y), Blue1s(x, y))

                                        rng2.Cells(j, Index).Font.Color = System.Drawing.Color.FromArgb(Red2s(x, y), Green2s(x, y), Blue2s(x, y))
                                        workSheet2.Activate()
                                    Else
                                        rng2.Cells(j, Index).ClearFormats()
                                    End If
                                Next j

                                Index = Index + 1
                                For j = 1 To SplitColumn - 1
                                    rng2.Cells(j, Index).value = Arr(i - 1, j - 1)
                                    If CheckBox1.Checked = True Then
                                        Dim x As Integer = i - 1
                                        Dim y As Integer = j - 1

                                        rng2.Cells(j, Index).Font.Name = FontNames(x, y)
                                        rng2.Cells(j, Index).Font.Size = FontSizes(x, y)

                                        If FontBolds(x, y) Then rng2.Cells(j, Index).Font.Bold = True
                                        If Fontitalics(x, y) Then rng2.Cells(j, Index).Font.Italic = True


                                        rng2.Cells(j, Index).Interior.Color = System.Drawing.Color.FromArgb(Red1s(x, y), Green1s(x, y), Blue1s(x, y))

                                        rng2.Cells(j, Index).Font.Color = System.Drawing.Color.FromArgb(Red2s(x, y), Green2s(x, y), Blue2s(x, y))
                                        workSheet2.Activate()
                                    Else
                                        rng2.Cells(j, Index).ClearFormats()
                                    End If
                                Next j

                                rng2.Cells(SplitColumn, Index).value = Text
                                If CheckBox1.Checked = True Then
                                    Dim x As Integer = i - 1
                                    Dim y As Integer = SplitColumn - 1

                                    rng2.Cells(SplitColumn, Index).Font.Name = FontNames(x, y)
                                    rng2.Cells(SplitColumn, Index).Font.Size = FontSizes(x, y)

                                    If FontBolds(x, y) Then rng2.Cells(SplitColumn, Index).Font.Bold = True
                                    If Fontitalics(x, y) Then rng2.Cells(SplitColumn, Index).Font.Italic = True


                                    rng2.Cells(SplitColumn, Index).Interior.Color = System.Drawing.Color.FromArgb(Red1s(x, y), Green1s(x, y), Blue1s(x, y))

                                    rng2.Cells(SplitColumn, Index).Font.Color = System.Drawing.Color.FromArgb(Red2s(x, y), Green2s(x, y), Blue2s(x, y))
                                    workSheet2.Activate()
                                Else
                                    rng2.Cells(SplitColumn, Index).ClearFormats()
                                End If

                                For j = SplitColumn + 1 To c
                                    rng2.Cells(j, Index).value = Arr(i - 1, j - 1)
                                    If CheckBox1.Checked = True Then
                                        Dim x As Integer = i - 1
                                        Dim y As Integer = j - 1

                                        rng2.Cells(j, Index).Font.Name = FontNames(x, y)
                                        rng2.Cells(j, Index).Font.Size = FontSizes(x, y)

                                        If FontBolds(x, y) Then rng2.Cells(j, Index).Font.Bold = True
                                        If Fontitalics(x, y) Then rng2.Cells(j, Index).Font.Italic = True


                                        rng2.Cells(j, Index).Interior.Color = System.Drawing.Color.FromArgb(Red1s(x, y), Green1s(x, y), Blue1s(x, y))

                                        rng2.Cells(j, Index).Font.Color = System.Drawing.Color.FromArgb(Red2s(x, y), Green2s(x, y), Blue2s(x, y))
                                        workSheet2.Activate()
                                    Else
                                        rng2.Cells(j, Index).ClearFormats()
                                    End If
                                Next j
                            Next

                        End If

                    ElseIf X11 Then

                        Dim W As Integer

                        If TextBox3.Text = "" Then
                            W = 1
                        Else
                            W = Int(TextBox3.Text)
                        End If

                        If X1 Then

                            Dim Index As Integer = 0

                            For i = 1 To r
                                Dim source As String = Arr(i - 1, SplitColumn - 1)
                                For k = 1 To Int(Len(source) / W)
                                    Index = Index + 1
                                    For j = 1 To SplitColumn - 1
                                        rng2.Cells(Index, j).value = Arr(i - 1, j - 1)
                                        If CheckBox1.Checked = True Then
                                            Dim x As Integer = i - 1
                                            Dim y As Integer = j - 1

                                            rng2.Cells(Index, j).Font.Name = FontNames(x, y)
                                            rng2.Cells(Index, j).Font.Size = FontSizes(x, y)

                                            If FontBolds(x, y) Then rng2.Cells(Index, j).Font.Bold = True
                                            If Fontitalics(x, y) Then rng2.Cells(Index, j).Font.Italic = True


                                            rng2.Cells(Index, j).Interior.Color = System.Drawing.Color.FromArgb(Red1s(x, y), Green1s(x, y), Blue1s(x, y))

                                            rng2.Cells(Index, j).Font.Color = System.Drawing.Color.FromArgb(Red2s(x, y), Green2s(x, y), Blue2s(x, y))
                                            workSheet2.Activate()
                                        Else
                                            rng2.Cells(Index, j).ClearFormats()
                                        End If
                                    Next j
                                    rng2.Cells(Index, SplitColumn).value = Mid(source, (W * (k - 1)) + 1, W)
                                    If CheckBox1.Checked = True Then
                                        Dim x As Integer = i - 1
                                        Dim y As Integer = SplitColumn - 1

                                        rng2.Cells(Index, SplitColumn).Font.Name = FontNames(x, y)
                                        rng2.Cells(Index, SplitColumn).Font.Size = FontSizes(x, y)

                                        If FontBolds(x, y) Then rng2.Cells(Index, SplitColumn).Font.Bold = True
                                        If Fontitalics(x, y) Then rng2.Cells(Index, SplitColumn).Font.Italic = True


                                        rng2.Cells(Index, SplitColumn).Interior.Color = System.Drawing.Color.FromArgb(Red1s(x, y), Green1s(x, y), Blue1s(x, y))

                                        rng2.Cells(Index, SplitColumn).Font.Color = System.Drawing.Color.FromArgb(Red2s(x, y), Green2s(x, y), Blue2s(x, y))
                                        workSheet2.Activate()
                                    Else
                                        rng2.Cells(Index, SplitColumn).ClearFormats()
                                    End If
                                    For j = SplitColumn + 1 To c
                                        rng2.Cells(Index, j).value = Arr(i - 1, j - 1)
                                        If CheckBox1.Checked = True Then
                                            Dim x As Integer = i - 1
                                            Dim y As Integer = j - 1

                                            rng2.Cells(Index, j).Font.Name = FontNames(x, y)
                                            rng2.Cells(Index, j).Font.Size = FontSizes(x, y)

                                            If FontBolds(x, y) Then rng2.Cells(Index, j).Font.Bold = True
                                            If Fontitalics(x, y) Then rng2.Cells(Index, j).Font.Italic = True


                                            rng2.Cells(Index, j).Interior.Color = System.Drawing.Color.FromArgb(Red1s(x, y), Green1s(x, y), Blue1s(x, y))

                                            rng2.Cells(Index, j).Font.Color = System.Drawing.Color.FromArgb(Red2s(x, y), Green2s(x, y), Blue2s(x, y))
                                            workSheet2.Activate()
                                        Else
                                            rng2.Cells(Index, j).ClearFormats()
                                        End If
                                    Next j
                                Next
                                If Len(source) Mod W <> 0 Then
                                    Index = Index + 1
                                    For j = 1 To SplitColumn - 1
                                        rng2.Cells(Index, j).value = Arr(i - 1, j - 1)
                                        If CheckBox1.Checked = True Then
                                            Dim x As Integer = i - 1
                                            Dim y As Integer = j - 1

                                            rng2.Cells(Index, j).Font.Name = FontNames(x, y)
                                            rng2.Cells(Index, j).Font.Size = FontSizes(x, y)

                                            If FontBolds(x, y) Then rng2.Cells(Index, j).Font.Bold = True
                                            If Fontitalics(x, y) Then rng2.Cells(Index, j).Font.Italic = True


                                            rng2.Cells(Index, j).Interior.Color = System.Drawing.Color.FromArgb(Red1s(x, y), Green1s(x, y), Blue1s(x, y))

                                            rng2.Cells(Index, j).Font.Color = System.Drawing.Color.FromArgb(Red2s(x, y), Green2s(x, y), Blue2s(x, y))
                                            workSheet2.Activate()
                                        Else
                                            rng2.Cells(Index, j).ClearFormats()
                                        End If
                                    Next j
                                    rng2.Cells(Index, SplitColumn).value = Mid(source, Len(source) - (Len(source) Mod W) + 1, Len(source) Mod W)
                                    If CheckBox1.Checked = True Then
                                        Dim x As Integer = i - 1
                                        Dim y As Integer = SplitColumn - 1

                                        rng2.Cells(Index, SplitColumn).Font.Name = FontNames(x, y)
                                        rng2.Cells(Index, SplitColumn).Font.Size = FontSizes(x, y)

                                        If FontBolds(x, y) Then rng2.Cells(Index, SplitColumn).Font.Bold = True
                                        If Fontitalics(x, y) Then rng2.Cells(Index, SplitColumn).Font.Italic = True


                                        rng2.Cells(Index, SplitColumn).Interior.Color = System.Drawing.Color.FromArgb(Red1s(x, y), Green1s(x, y), Blue1s(x, y))

                                        rng2.Cells(Index, SplitColumn).Font.Color = System.Drawing.Color.FromArgb(Red2s(x, y), Green2s(x, y), Blue2s(x, y))
                                        workSheet2.Activate()
                                    Else
                                        rng2.Cells(Index, SplitColumn).ClearFormats()
                                    End If
                                    For j = SplitColumn + 1 To c
                                        rng2.Cells(Index, j).value = Arr(i - 1, j - 1)
                                        If CheckBox1.Checked = True Then
                                            Dim x As Integer = i - 1
                                            Dim y As Integer = j - 1

                                            rng2.Cells(Index, j).Font.Name = FontNames(x, y)
                                            rng2.Cells(Index, j).Font.Size = FontSizes(x, y)

                                            If FontBolds(x, y) Then rng2.Cells(Index, j).Font.Bold = True
                                            If Fontitalics(x, y) Then rng2.Cells(Index, j).Font.Italic = True


                                            rng2.Cells(Index, j).Interior.Color = System.Drawing.Color.FromArgb(Red1s(x, y), Green1s(x, y), Blue1s(x, y))

                                            rng2.Cells(Index, j).Font.Color = System.Drawing.Color.FromArgb(Red2s(x, y), Green2s(x, y), Blue2s(x, y))
                                            workSheet2.Activate()
                                        Else
                                            rng2.Cells(Index, j).ClearFormats()
                                        End If
                                    Next j
                                End If
                            Next

                        ElseIf X2 Then

                            Dim Index As Integer = 0

                            For i = 1 To r
                                Dim source As String = Arr(i - 1, SplitColumn - 1)
                                For k = 1 To Int(Len(source) / W)
                                    Index = Index + 1
                                    For j = 1 To SplitColumn - 1
                                        rng2.Cells(j, Index).value = Arr(i - 1, j - 1)
                                        If CheckBox1.Checked = True Then
                                            Dim x As Integer = i - 1
                                            Dim y As Integer = j - 1

                                            rng2.Cells(j, Index).Font.Name = FontNames(x, y)
                                            rng2.Cells(j, Index).Font.Size = FontSizes(x, y)

                                            If FontBolds(x, y) Then rng2.Cells(j, Index).Font.Bold = True
                                            If Fontitalics(x, y) Then rng2.Cells(j, Index).Font.Italic = True


                                            rng2.Cells(j, Index).Interior.Color = System.Drawing.Color.FromArgb(Red1s(x, y), Green1s(x, y), Blue1s(x, y))

                                            rng2.Cells(j, Index).Font.Color = System.Drawing.Color.FromArgb(Red2s(x, y), Green2s(x, y), Blue2s(x, y))
                                            workSheet2.Activate()
                                        Else
                                            rng2.Cells(j, Index).ClearFormats()
                                        End If
                                    Next j
                                    rng2.Cells(SplitColumn, Index).value = Mid(source, (W * (k - 1)) + 1, W)
                                    If CheckBox1.Checked = True Then
                                        Dim x As Integer = i - 1
                                        Dim y As Integer = SplitColumn - 1

                                        rng2.Cells(SplitColumn, Index).Font.Name = FontNames(x, y)
                                        rng2.Cells(SplitColumn, Index).Font.Size = FontSizes(x, y)

                                        If FontBolds(x, y) Then rng2.Cells(SplitColumn, Index).Font.Bold = True
                                        If Fontitalics(x, y) Then rng2.Cells(SplitColumn, Index).Font.Italic = True


                                        rng2.Cells(SplitColumn, Index).Interior.Color = System.Drawing.Color.FromArgb(Red1s(x, y), Green1s(x, y), Blue1s(x, y))

                                        rng2.Cells(SplitColumn, Index).Font.Color = System.Drawing.Color.FromArgb(Red2s(x, y), Green2s(x, y), Blue2s(x, y))
                                        workSheet2.Activate()
                                    Else
                                        rng2.Cells(SplitColumn, Index).ClearFormats()
                                    End If
                                    For j = SplitColumn + 1 To c
                                        rng2.Cells(j, Index).value = Arr(i - 1, j - 1)
                                        If CheckBox1.Checked = True Then
                                            Dim x As Integer = i - 1
                                            Dim y As Integer = j - 1

                                            rng2.Cells(j, Index).Font.Name = FontNames(x, y)
                                            rng2.Cells(j, Index).Font.Size = FontSizes(x, y)

                                            If FontBolds(x, y) Then rng2.Cells(j, Index).Font.Bold = True
                                            If Fontitalics(x, y) Then rng2.Cells(j, Index).Font.Italic = True


                                            rng2.Cells(j, Index).Interior.Color = System.Drawing.Color.FromArgb(Red1s(x, y), Green1s(x, y), Blue1s(x, y))

                                            rng2.Cells(j, Index).Font.Color = System.Drawing.Color.FromArgb(Red2s(x, y), Green2s(x, y), Blue2s(x, y))
                                            workSheet2.Activate()
                                        Else
                                            rng2.Cells(j, Index).ClearFormats()
                                        End If
                                    Next j
                                Next
                                If Len(source) Mod W <> 0 Then
                                    Index = Index + 1
                                    For j = 1 To SplitColumn - 1
                                        rng2.Cells(j, Index).value = Arr(i - 1, j - 1)
                                        If CheckBox1.Checked = True Then
                                            Dim x As Integer = i - 1
                                            Dim y As Integer = j - 1

                                            rng2.Cells(j, Index).Font.Name = FontNames(x, y)
                                            rng2.Cells(j, Index).Font.Size = FontSizes(x, y)

                                            If FontBolds(x, y) Then rng2.Cells(j, Index).Font.Bold = True
                                            If Fontitalics(x, y) Then rng2.Cells(j, Index).Font.Italic = True


                                            rng2.Cells(j, Index).Interior.Color = System.Drawing.Color.FromArgb(Red1s(x, y), Green1s(x, y), Blue1s(x, y))

                                            rng2.Cells(j, Index).Font.Color = System.Drawing.Color.FromArgb(Red2s(x, y), Green2s(x, y), Blue2s(x, y))
                                            workSheet2.Activate()
                                        Else
                                            rng2.Cells(j, Index).ClearFormats()
                                        End If
                                    Next j
                                    rng2.Cells(SplitColumn, Index).value = Mid(source, Len(source) - (Len(source) Mod W) + 1, Len(source) Mod W)
                                    If CheckBox1.Checked = True Then
                                        Dim x As Integer = i - 1
                                        Dim y As Integer = SplitColumn - 1

                                        rng2.Cells(SplitColumn, Index).Font.Name = FontNames(x, y)
                                        rng2.Cells(SplitColumn, Index).Font.Size = FontSizes(x, y)

                                        If FontBolds(x, y) Then rng2.Cells(SplitColumn, Index).Font.Bold = True
                                        If Fontitalics(x, y) Then rng2.Cells(SplitColumn, Index).Font.Italic = True


                                        rng2.Cells(SplitColumn, Index).Interior.Color = System.Drawing.Color.FromArgb(Red1s(x, y), Green1s(x, y), Blue1s(x, y))

                                        rng2.Cells(SplitColumn, Index).Font.Color = System.Drawing.Color.FromArgb(Red2s(x, y), Green2s(x, y), Blue2s(x, y))
                                        workSheet2.Activate()
                                    Else
                                        rng2.Cells(SplitColumn, Index).ClearFormats()
                                    End If
                                    For j = SplitColumn + 1 To c
                                        rng2.Cells(j, Index).value = Arr(i - 1, j - 1)
                                        If CheckBox1.Checked = True Then
                                            Dim x As Integer = i - 1
                                            Dim y As Integer = j - 1

                                            rng2.Cells(j, Index).Font.Name = FontNames(x, y)
                                            rng2.Cells(j, Index).Font.Size = FontSizes(x, y)

                                            If FontBolds(x, y) Then rng2.Cells(j, Index).Font.Bold = True
                                            If Fontitalics(x, y) Then rng2.Cells(j, Index).Font.Italic = True


                                            rng2.Cells(j, Index).Interior.Color = System.Drawing.Color.FromArgb(Red1s(x, y), Green1s(x, y), Blue1s(x, y))

                                            rng2.Cells(j, Index).Font.Color = System.Drawing.Color.FromArgb(Red2s(x, y), Green2s(x, y), Blue2s(x, y))
                                            workSheet2.Activate()
                                        Else
                                            rng2.Cells(j, Index).ClearFormats()
                                        End If
                                    Next j
                                End If
                            Next

                        End If

                    End If

                End If

                Me.Close()
                workSheet2.Activate()
                rng2.Select()

                Dim columnNum As Integer
                For j = 1 To rng2.Columns.Count
                    columnNum = rng2.Cells(1, j).column
                    workSheet2.Columns(columnNum).Autofit
                Next

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

        Try
            excelApp = Globals.ThisAddIn.Application
            workBook = excelApp.ActiveWorkbook
            workSheet = workBook.ActiveSheet

            TextBox1.SelectionStart = TextBox1.Text.Length
            TextBox1.ScrollToCaret()

            rng = workSheet.Range(TextBox1.Text)
            TextBoxChanged = True
            rng.Select()

            ComboBox3.Items.Clear()

            For j = 1 To rng.Columns.Count
                Dim ItemName As String
                Dim CName As String = Split(rng.Cells(1, j).Address, "$")(1)
                If rng.Cells(1, 1).Row > 1 Then
                    ItemName = "Column " & CName & " (" & rng.Cells(0, j).value & ") "
                Else
                    ItemName = "Column " & CName
                End If
                ComboBox3.Items.Add(ItemName)
            Next

            Call Display()

            TextBoxChanged = False

        Catch ex As Exception

        End Try

    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged

        Try
            If RadioButton1.Checked Then
                Call Display()
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged

        Try
            If RadioButton2.Checked Then
                Call Display()
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub RadioButton9_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton9.CheckedChanged

        Try
            If RadioButton9.Checked Then
                Call Display()
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub RadioButton8_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton8.CheckedChanged

        Try
            If RadioButton8.Checked Then
                Call Display()
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged

        Try
            If RadioButton3.Checked Then
                Call Display()
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub RadioButton7_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton7.CheckedChanged

        Try
            If RadioButton7.Checked Then
                Call Display()
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub RadioButton10_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton10.CheckedChanged

        Try
            If RadioButton10.Checked Then
                ComboBox2.Enabled = True
                ComboBox2.Focus()
                Call Display()
            Else
                ComboBox2.Text = ""
                ComboBox2.Enabled = False
            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub RadioButton11_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton11.CheckedChanged

        Try
            If RadioButton11.Checked Then
                PictureBox11.Enabled = True
                TextBox3.Enabled = True
                TextBox3.Focus()
                Call Display()
            Else
                TextBox3.Clear()
                PictureBox11.Enabled = False
                TextBox3.Enabled = False
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub ComboBox2_TextChanged(sender As Object, e As EventArgs) Handles ComboBox2.TextChanged

        Try
            Call Display()
        Catch ex As Exception

        End Try

    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged

        Try
            If IsNumeric(TextBox3.Text) Or TextBox3.Text = "" Then
                Call Display()
            Else
                MessageBox.Show("Enter a Numerical Value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        Try
            Call Display()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub AutoSelection_Click(sender As Object, e As EventArgs) Handles AutoSelection.Click
        Try

            FocusedTextBox = 1

            Dim userInput As Excel.Range = excelApp.InputBox("Select a range", Type:=8)
            rng = userInput

            Try
                Dim sheetName As String
                sheetName = Split(rng.Address(True, True, Excel.XlReferenceStyle.xlA1, True), "]")(1)
                sheetName = Split(sheetName, "!")(0)

                If Mid(sheetName, Len(sheetName), 1) = "'" Then
                    sheetName = Mid(sheetName, 1, Len(sheetName) - 1)
                End If

                workSheet = workBook.Worksheets(sheetName)
                workSheet.Activate()

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

    Private Sub Selection_Click(sender As Object, e As EventArgs) Handles Selection.Click
        Try
            FocusedTextBox = 1

            Dim userInput As Excel.Range = excelApp.InputBox("Select a range", Type:=8)
            Dim rng As Microsoft.Office.Interop.Excel.Range = userInput

            Try
                Dim sheetName As String
                sheetName = Split(rng.Address(True, True, Excel.XlReferenceStyle.xlA1, True), "]")(1)
                sheetName = Split(sheetName, "!")(0)

                If Mid(sheetName, Len(sheetName), 1) = "'" Then
                    sheetName = Mid(sheetName, 1, Len(sheetName) - 1)
                End If

                workSheet = workBook.Worksheets(sheetName)
                workSheet.Activate()
            Catch ex As Exception

            End Try

            rng.Select()

            TextBox1.Text = rng.Address
            TextBox1.Focus()

        Catch ex As Exception

        End Try
    End Sub

    Private Sub RadioButton5_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton5.CheckedChanged
        Try
            If RadioButton5.Checked = True Then
                workSheet2 = workSheet
                rng2 = rng
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub RadioButton4_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton4.CheckedChanged
        Try
            If RadioButton4.Checked = True Then
                Label3.Enabled = True
                PictureBox2.Enabled = True
                PictureBox3.Enabled = True
                TextBox4.Enabled = True
                TextBox4.Focus()
            Else
                TextBox4.Clear()
                Label3.Enabled = False
                PictureBox2.Enabled = False
                PictureBox3.Enabled = False
                TextBox4.Enabled = False
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        Try
            workSheet2 = workBook.ActiveSheet

            TextBox4.SelectionStart = TextBox4.Text.Length
            TextBox4.ScrollToCaret()

            rng2 = workSheet2.Range(TextBox3.Text)

            TextBoxChanged = True
            rng2.Select()
            TextBoxChanged = False

        Catch ex As Exception

        End Try
    End Sub

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click
        Try
            FocusedTextBox = 4
            Me.Hide()

            excelApp = Globals.ThisAddIn.Application
            workBook = excelApp.ActiveWorkbook

            Dim userInput As Excel.Range = excelApp.InputBox("Select a range", Type:=8)
            rng2 = userInput


            Dim sheetName As String
            sheetName = Split(rng2.Address(True, True, Excel.XlReferenceStyle.xlA1, True), "]")(1)
            sheetName = Split(sheetName, "!")(0)

            If Mid(sheetName, Len(sheetName), 1) = "'" Then
                sheetName = Mid(sheetName, 1, Len(sheetName) - 1)
            End If

            workSheet2 = workBook.Worksheets(sheetName)
            workSheet2.Activate()

            rng2.Select()

            TextBox4.Text = rng2.Address

            Me.Show()
            TextBox4.Focus()

        Catch ex As Exception

            Me.Show()
            TextBox4.Focus()

        End Try
    End Sub

    Private Sub Form25_Split_Range_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try

            excelApp = Globals.ThisAddIn.Application
            workBook = excelApp.ActiveWorkbook
            workSheet = workBook.ActiveSheet
            workSheet2 = workBook.ActiveSheet

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

            If TextBoxChanged = False Then
                If FocusedTextBox = 1 Then
                    TextBox1.Text = selectedRange.Address
                    workSheet = workBook.ActiveSheet
                    rng = selectedRange
                    TextBox1.Focus()

                ElseIf FocusedTextBox = 4 Then
                    TextBox4.Text = selectedRange.Address
                    workSheet2 = workBook.ActiveSheet
                    rng2 = selectedRange
                    TextBox4.Focus()
                End If
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

    Private Sub TextBox4_GotFocus(sender As Object, e As EventArgs) Handles TextBox4.GotFocus
        Try
            FocusedTextBox = 4
        Catch ex As Exception

        End Try
    End Sub

    Private Sub AutoSelection_GotFocus(sender As Object, e As EventArgs) Handles AutoSelection.GotFocus
        Try
            FocusedTextBox = 1
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Selection_GotFocus(sender As Object, e As EventArgs) Handles Selection.GotFocus
        Try
            FocusedTextBox = 1
        Catch ex As Exception

        End Try
    End Sub

    Private Sub PictureBox3_GotFocus(sender As Object, e As EventArgs) Handles PictureBox3.GotFocus
        Try
            FocusedTextBox = 4
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Me.Close()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        Try
            Call Display()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Button2_MouseEnter(sender As Object, e As EventArgs) Handles Button2.MouseEnter
        Try
            Button2.BackColor = Color.FromArgb(65, 105, 225)
            Button2.ForeColor = Color.FromArgb(255, 255, 255)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Button1_MouseEnter(sender As Object, e As EventArgs) Handles Button1.MouseEnter

        Try
            Button1.BackColor = Color.FromArgb(65, 105, 225)
            Button1.ForeColor = Color.FromArgb(255, 255, 255)
        Catch ex As Exception

        End Try

    End Sub

    Private Sub Button2_MouseLeave(sender As Object, e As EventArgs) Handles Button2.MouseLeave
        Try

            Button2.BackColor = Color.FromArgb(255, 255, 255)
            Button2.ForeColor = Color.FromArgb(70, 70, 70)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Button1_MouseLeave(sender As Object, e As EventArgs) Handles Button1.MouseLeave
        Try

            Button1.BackColor = Color.FromArgb(255, 255, 255)
            Button1.ForeColor = Color.FromArgb(70, 70, 70)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub AutoSelection_KeyDown(sender As Object, e As KeyEventArgs) Handles AutoSelection.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub Button1_KeyDown(sender As Object, e As KeyEventArgs) Handles Button1.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub Button2_KeyDown(sender As Object, e As KeyEventArgs) Handles Button2.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub CheckBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles CheckBox1.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub CheckBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles CheckBox2.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub ComboBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox1.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub ComboBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox2.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub ComboBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox3.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub CustomGroupBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles CustomGroupBox1.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub CustomGroupBox10_KeyDown(sender As Object, e As KeyEventArgs) Handles CustomGroupBox10.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub CustomGroupBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles CustomGroupBox2.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub CustomGroupBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles CustomGroupBox4.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub CustomGroupBox5_KeyDown(sender As Object, e As KeyEventArgs) Handles CustomGroupBox5.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub CustomGroupBox6_KeyDown(sender As Object, e As KeyEventArgs) Handles CustomGroupBox6.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub CustomGroupBox7_KeyDown(sender As Object, e As KeyEventArgs) Handles CustomGroupBox7.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub CustomGroupBox8_KeyDown(sender As Object, e As KeyEventArgs) Handles CustomGroupBox8.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub CustomPanel1_KeyDown(sender As Object, e As KeyEventArgs) Handles CustomPanel1.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub CustomPanel2_KeyDown(sender As Object, e As KeyEventArgs) Handles CustomPanel2.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub Label1_KeyDown(sender As Object, e As KeyEventArgs) Handles Label1.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub Label2_KeyDown(sender As Object, e As KeyEventArgs) Handles Label2.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub Label3_KeyDown(sender As Object, e As KeyEventArgs) Handles Label3.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub PictureBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles PictureBox1.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub PictureBox10_KeyDown(sender As Object, e As KeyEventArgs) Handles PictureBox10.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub PictureBox11_KeyDown(sender As Object, e As KeyEventArgs) Handles PictureBox11.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub PictureBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles PictureBox2.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub PictureBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles PictureBox3.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub PictureBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles PictureBox4.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub PictureBox5_KeyDown(sender As Object, e As KeyEventArgs) Handles PictureBox5.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub PictureBox6_KeyDown(sender As Object, e As KeyEventArgs) Handles PictureBox6.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub PictureBox7_KeyDown(sender As Object, e As KeyEventArgs) Handles PictureBox7.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub PictureBox8_KeyDown(sender As Object, e As KeyEventArgs) Handles PictureBox8.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub RadioButton1_KeyDown(sender As Object, e As KeyEventArgs) Handles RadioButton1.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub RadioButton10_KeyDown(sender As Object, e As KeyEventArgs) Handles RadioButton10.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub RadioButton11_KeyDown(sender As Object, e As KeyEventArgs) Handles RadioButton11.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub RadioButton2_KeyDown(sender As Object, e As KeyEventArgs) Handles RadioButton2.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub RadioButton3_KeyDown(sender As Object, e As KeyEventArgs) Handles RadioButton3.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub RadioButton4_KeyDown(sender As Object, e As KeyEventArgs) Handles RadioButton4.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub RadioButton5_KeyDown(sender As Object, e As KeyEventArgs) Handles RadioButton5.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub RadioButton7_KeyDown(sender As Object, e As KeyEventArgs) Handles RadioButton7.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub RadioButton8_KeyDown(sender As Object, e As KeyEventArgs) Handles RadioButton8.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub RadioButton9_KeyDown(sender As Object, e As KeyEventArgs) Handles RadioButton9.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub TextBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox3.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub TextBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox4.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub VScrollBar1_KeyDown(sender As Object, e As KeyEventArgs) Handles VScrollBar1.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub Form25_Split_Range_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        form_flag = False
    End Sub

    Private Sub Form25_Split_Range_Disposed(sender As Object, e As EventArgs) Handles Me.Disposed
        form_flag = False
    End Sub

    Private Sub Form25_Split_Range_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        Me.Focus()
        Me.BringToFront()
        Me.Activate()
        Me.BeginInvoke(New System.Action(Sub()
                                             TextBox1.Text = rng.Address
                                             SetWindowPos(Me.Handle, New IntPtr(HWND_TOPMOST), 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOSIZE)
                                         End Sub))
    End Sub
End Class