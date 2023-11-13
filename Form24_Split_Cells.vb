Imports System.ComponentModel
Imports System.Diagnostics
Imports System.Drawing
Imports System.Reflection
Imports System.Reflection.Emit
Imports System.Security.Cryptography
Imports System.Security.Policy
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Button
Imports Microsoft.Office.Interop.Excel

Public Class Form24_Split_Cells

    Dim WithEvents excelApp As Excel.Application
    Dim workBook As Excel.Workbook
    Dim workSheet As Excel.Worksheet
    Dim workSheet2 As Excel.Worksheet
    Public OpenSheet As Excel.Worksheet
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
    Private Function FindMax(Arr)

        Dim Max As Integer = Arr(LBound(Arr))

        For i = LBound(Arr) + 1 To UBound(Arr)
            If Arr(i) > Max Then
                Max = Arr(i)
            End If
        Next
        FindMax = Max

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
            TextBoxChanged = True

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

            If ((MaxOfColumn(displayRng) * BaseWidth) / 10) > CustomPanel1.Width Then
                Width = ((MaxOfColumn(displayRng) * BaseWidth) / 10)
            Else
                Width = CustomPanel1.Width
            End If

            For i = 1 To r
                Dim label As New System.Windows.Forms.Label
                label.Text = displayRng.Cells(i, 1).Value
                label.Location = New System.Drawing.Point(0, (i - 1) * Height)
                label.Height = Height
                label.Width = Width
                label.BorderStyle = BorderStyle.FixedSingle
                label.TextAlign = ContentAlignment.MiddleCenter

                If CheckBox1.Checked = True Then

                    Dim cell As Excel.Range = displayRng.Cells(i, 1)
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

            CustomPanel1.AutoScroll = True

            Dim X1 As Boolean = RadioButton1.Checked
            Dim X2 As Boolean = RadioButton2.Checked
            Dim X3 As Boolean = RadioButton3.Checked
            Dim X7 As Boolean = RadioButton7.Checked
            Dim X8 As Boolean = RadioButton8.Checked
            Dim X9 As Boolean = RadioButton9.Checked
            Dim X10 As Boolean = RadioButton10.Checked
            Dim X11 As Boolean = RadioButton11.Checked


            If (X1 Or X2) And (X3 Or X7 Or X8 Or X9 Or X10 Or X11) Then

                Dim SplitColumn As Integer = 1
                Dim Ordinate As Double

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

                    If X2 Then

                        Dim Lengths(r - 1) As Integer

                        Dim Index As Integer
                        Dim position As Integer

                        For i = 1 To r
                            Dim source As String = displayRng.Cells(i, SplitColumn).value
                            Lengths(i - 1) = CountSeparator(source, Separator)
                        Next i

                        Dim TotalColumn As Integer = FindMax(Lengths)
                        Dim SplitValues(r - 1, TotalColumn - 1) As String

                        For i = 1 To r
                            Dim source As String = displayRng.Cells(i, SplitColumn).value
                            position = 1
                            Index = -1
                            For k = 1 To Len(source)
                                If Mid(source, k, Len(Separator)) = Separator Then
                                    If k - position > 0 Then
                                        Index = Index + 1
                                        SplitValues(i - 1, Index) = Mid(source, position, k - position)
                                    End If
                                    position = k + Len(Separator)
                                End If
                            Next
                            If position <= Len(source) Then
                                Index = Index + 1
                                SplitValues(i - 1, Index) = Mid(source, position, Len(source) - position + 1)
                            End If
                        Next

                        Ordinate = 0

                        For j = LBound(SplitValues, 2) To UBound(SplitValues, 2)
                            Dim NewColumn(r - 1) As String
                            For i = LBound(SplitValues, 1) To UBound(SplitValues, 1)
                                NewColumn(i) = SplitValues(i, j)
                            Next
                            If TotalColumn = 1 Then
                                Width = CustomPanel2.Width
                            Else
                                Width = (MaxOfArray(NewColumn) * BaseWidth) / 10
                            End If
                            For i = LBound(SplitValues, 1) To UBound(SplitValues, 1)
                                Dim label As New System.Windows.Forms.Label
                                label.Text = SplitValues(i, j)
                                label.Location = New System.Drawing.Point(Ordinate, i * Height)
                                label.Height = Height
                                label.Width = Width
                                label.BorderStyle = BorderStyle.FixedSingle
                                label.TextAlign = ContentAlignment.MiddleCenter

                                If CheckBox1.Checked = True Then

                                    Dim cell As Excel.Range = displayRng.Cells(i + 1, 1)
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
                                CustomPanel2.Controls.Add(label)
                            Next
                            Ordinate = Ordinate + Width
                        Next

                    ElseIf X1 Then

                        Dim Lengths(r - 1) As Integer

                        Dim Index As Integer
                        Dim position As Integer

                        For i = 1 To r
                            Dim source As String = displayRng.Cells(i, SplitColumn).value
                            Lengths(i - 1) = CountSeparator(source, Separator)
                        Next i

                        Dim TotalColumn As Integer = FindMax(Lengths)
                        Dim SplitValues(r - 1, TotalColumn - 1) As String

                        For i = 1 To r
                            Dim source As String = displayRng.Cells(i, SplitColumn).value
                            position = 1
                            Index = -1
                            For k = 1 To Len(source)
                                If Mid(source, k, Len(Separator)) = Separator Then
                                    If k - position > 0 Then
                                        Index = Index + 1
                                        SplitValues(i - 1, Index) = Mid(source, position, k - position)
                                    End If
                                    position = k + Len(Separator)
                                End If
                            Next
                            If position <= Len(source) Then
                                Index = Index + 1
                                SplitValues(i - 1, Index) = Mid(source, position, Len(source) - position + 1)
                            End If
                        Next

                        Ordinate = 0

                        For i = LBound(SplitValues, 1) To UBound(SplitValues, 1)
                            Dim NewColumn(TotalColumn - 1) As String
                            For j = LBound(SplitValues, 2) To UBound(SplitValues, 2)
                                NewColumn(j) = SplitValues(i, j)
                            Next
                            If TotalColumn * Height < CustomPanel2.Height Then
                                Height = CustomPanel2.Height / TotalColumn
                            End If
                            If r = 1 Then
                                Width = CustomPanel2.Width
                            Else
                                Width = (MaxOfArray(NewColumn) * BaseWidth) / 10
                            End If
                            For j = LBound(SplitValues, 2) To UBound(SplitValues, 2)
                                Dim label As New System.Windows.Forms.Label
                                label.Text = SplitValues(i, j)
                                label.Location = New System.Drawing.Point(Ordinate, j * Height)
                                label.Height = Height
                                label.Width = Width
                                label.BorderStyle = BorderStyle.FixedSingle
                                label.TextAlign = ContentAlignment.MiddleCenter

                                If CheckBox1.Checked = True Then

                                    Dim cell As Excel.Range = displayRng.Cells(i + 1, 1)
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
                                CustomPanel2.Controls.Add(label)
                            Next
                            Ordinate = Ordinate + Width
                        Next
                    End If

                ElseIf X3 Then

                    If X2 Then

                        Dim Numbers(r - 1) As String
                        Dim Texts(r - 1) As String

                        For i = 1 To r
                            Dim source As String = displayRng.Cells(i, SplitColumn).value
                            Dim NumberText(1) As String
                            NumberText = SeparateNumberText(source)
                            Numbers(i - 1) = NumberText(0)
                            Texts(i - 1) = NumberText(1)
                        Next

                        Dim NumbersWidth As Double = (MaxOfArray(Numbers) * BaseWidth) / 10
                        Dim TextsWidth As Double = (MaxOfArray(Texts) * BaseWidth) / 10

                        If (NumbersWidth + TextsWidth) < CustomPanel2.Width Then
                            NumbersWidth = CustomPanel2.Width / 2
                            TextsWidth = CustomPanel2.Width / 2
                        End If

                        Ordinate = 0

                        For i = LBound(Numbers) To UBound(Numbers)
                            Dim label As New System.Windows.Forms.Label
                            label.Text = Numbers(i)
                            label.Location = New System.Drawing.Point(Ordinate, i * Height)
                            label.Height = Height
                            label.Width = NumbersWidth
                            label.BorderStyle = BorderStyle.FixedSingle
                            label.TextAlign = ContentAlignment.MiddleCenter

                            If CheckBox1.Checked = True Then

                                Dim cell As Excel.Range = displayRng.Cells(i + 1, 1)
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
                            CustomPanel2.Controls.Add(label)
                        Next
                        Ordinate = Ordinate + NumbersWidth

                        For i = LBound(Texts) To UBound(Texts)
                            Dim label As New System.Windows.Forms.Label
                            label.Text = Texts(i)
                            label.Location = New System.Drawing.Point(Ordinate, i * Height)
                            label.Height = Height
                            label.Width = TextsWidth
                            label.BorderStyle = BorderStyle.FixedSingle
                            label.TextAlign = ContentAlignment.MiddleCenter

                            If CheckBox1.Checked = True Then

                                Dim cell As Excel.Range = displayRng.Cells(i + 1, 1)
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
                            CustomPanel2.Controls.Add(label)
                        Next

                    ElseIf X1 Then

                        Dim TotalColumn As Integer = 2
                        Dim SplitValues(r - 1, TotalColumn - 1) As String

                        For i = 1 To r
                            Dim source As String = displayRng.Cells(i, SplitColumn).value
                            Dim NumberText(1) As String
                            NumberText = SeparateNumberText(source)
                            SplitValues(i - 1, 0) = NumberText(0)
                            SplitValues(i - 1, 1) = NumberText(1)
                        Next

                        Ordinate = 0

                        For i = LBound(SplitValues, 1) To UBound(SplitValues, 1)
                            Dim NewColumn(TotalColumn - 1) As String
                            For j = LBound(SplitValues, 2) To UBound(SplitValues, 2)
                                NewColumn(j) = SplitValues(i, j)
                            Next
                            Height = CustomPanel2.Height / 2
                            If r = 1 Then
                                Width = CustomPanel2.Width
                            Else
                                Width = (MaxOfArray(NewColumn) * BaseWidth) / 10
                            End If
                            For j = LBound(SplitValues, 2) To UBound(SplitValues, 2)
                                Dim label As New System.Windows.Forms.Label
                                label.Text = SplitValues(i, j)
                                label.Location = New System.Drawing.Point(Ordinate, j * Height)
                                label.Height = Height
                                label.Width = Width
                                label.BorderStyle = BorderStyle.FixedSingle
                                label.TextAlign = ContentAlignment.MiddleCenter

                                If CheckBox1.Checked = True Then

                                    Dim cell As Excel.Range = displayRng.Cells(i + 1, 1)
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
                                CustomPanel2.Controls.Add(label)
                            Next
                            Ordinate = Ordinate + Width
                        Next
                    End If

                ElseIf X11 Then

                    Dim W As Integer

                    If TextBox3.Text = "" Then
                        W = 1
                    Else
                        W = Int(TextBox3.Text)
                    End If

                    If X2 Then

                        Dim Lengths(r - 1) As Integer

                        Dim Index As Integer

                        For i = 1 To r
                            Dim source As String = displayRng.Cells(i, SplitColumn).value
                            If Len(source) Mod W = 0 Then
                                Lengths(i - 1) = Int(Len(source) / W)
                            Else
                                Lengths(i - 1) = Int(Len(source) / W) + 1
                            End If
                        Next i

                        Dim TotalColumn As Integer = FindMax(Lengths)
                        Dim SplitValues(r - 1, TotalColumn - 1) As String

                        For i = 1 To r
                            Dim source As String = displayRng.Cells(i, SplitColumn).value
                            Index = -1
                            For k = 1 To Int(Len(source) / W)
                                Index = Index + 1
                                SplitValues(i - 1, Index) = Mid(source, (W * (k - 1)) + 1, W)
                            Next
                            If Len(source) Mod W <> 0 Then
                                Index = Index + 1
                                SplitValues(i - 1, Index) = Mid(source, Len(source) - (Len(source) Mod W) + 1, Len(source) Mod W)
                            End If
                        Next

                        Ordinate = 0

                        For j = LBound(SplitValues, 2) To UBound(SplitValues, 2)
                            Dim NewColumn(r - 1) As String
                            For i = LBound(SplitValues, 1) To UBound(SplitValues, 1)
                                NewColumn(i) = SplitValues(i, j)
                            Next
                            If TotalColumn = 1 Then
                                Width = CustomPanel2.Width
                            Else
                                Width = (MaxOfArray(NewColumn) * BaseWidth) / 10
                            End If
                            For i = LBound(SplitValues, 1) To UBound(SplitValues, 1)
                                Dim label As New System.Windows.Forms.Label
                                label.Text = SplitValues(i, j)
                                label.Location = New System.Drawing.Point(Ordinate, i * Height)
                                label.Height = Height
                                label.Width = Width
                                label.BorderStyle = BorderStyle.FixedSingle
                                label.TextAlign = ContentAlignment.MiddleCenter

                                If CheckBox1.Checked = True Then

                                    Dim cell As Excel.Range = displayRng.Cells(i + 1, 1)
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
                                CustomPanel2.Controls.Add(label)
                            Next
                            Ordinate = Ordinate + Width
                        Next


                    ElseIf X1 Then

                        Dim Lengths(r - 1) As Integer

                        Dim Index As Integer

                        For i = 1 To r
                            Dim source As String = displayRng.Cells(i, SplitColumn).value
                            If Len(source) Mod W = 0 Then
                                Lengths(i - 1) = Int(Len(source) / W)
                            Else
                                Lengths(i - 1) = Int(Len(source) / W) + 1
                            End If
                        Next i

                        Dim TotalColumn As Integer = FindMax(Lengths)
                        Dim SplitValues(r - 1, TotalColumn - 1) As String

                        For i = 1 To r
                            Dim source As String = displayRng.Cells(i, SplitColumn).value
                            Index = -1
                            For k = 1 To Int(Len(source) / W)
                                Index = Index + 1
                                SplitValues(i - 1, Index) = Mid(source, (W * (k - 1)) + 1, W)
                            Next
                            If Len(source) Mod W <> 0 Then
                                Index = Index + 1
                                SplitValues(i - 1, Index) = Mid(source, Len(source) - (Len(source) Mod W) + 1, Len(source) Mod W)
                            End If
                        Next

                        Ordinate = 0

                        For i = LBound(SplitValues, 1) To UBound(SplitValues, 1)
                            Dim NewColumn(TotalColumn - 1) As String
                            For j = LBound(SplitValues, 2) To UBound(SplitValues, 2)
                                NewColumn(j) = SplitValues(i, j)
                            Next
                            If TotalColumn * Height < CustomPanel2.Height Then
                                Height = CustomPanel2.Height / TotalColumn
                            End If
                            If r = 1 Then
                                Width = CustomPanel2.Width
                            Else
                                Width = (MaxOfArray(NewColumn) * BaseWidth) / 10
                            End If
                            For j = LBound(SplitValues, 2) To UBound(SplitValues, 2)
                                Dim label As New System.Windows.Forms.Label
                                label.Text = SplitValues(i, j)
                                label.Location = New System.Drawing.Point(Ordinate, j * Height)
                                label.Height = Height
                                label.Width = Width
                                label.BorderStyle = BorderStyle.FixedSingle
                                label.TextAlign = ContentAlignment.MiddleCenter

                                If CheckBox1.Checked = True Then

                                    Dim cell As Excel.Range = displayRng.Cells(i + 1, 1)
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
                                CustomPanel2.Controls.Add(label)
                            Next
                            Ordinate = Ordinate + Width
                        Next

                    End If

                End If

                CustomPanel2.AutoScroll = True

            End If

            TextBoxChanged = False

        Catch ex As Exception

        End Try

    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Try
            TextBoxChanged = True

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

            Dim X1 As Boolean = RadioButton1.Checked
            Dim X2 As Boolean = RadioButton2.Checked
            Dim X3 As Boolean = RadioButton3.Checked
            Dim X7 As Boolean = RadioButton7.Checked
            Dim X8 As Boolean = RadioButton8.Checked
            Dim X9 As Boolean = RadioButton9.Checked
            Dim X10 As Boolean = RadioButton10.Checked
            Dim X11 As Boolean = RadioButton11.Checked

            If X1 = False And X2 = False Then
                MessageBox.Show("Select a Split Option.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                workSheet.Activate()
                rng.Select()
                Exit Sub
            End If

            If X3 = False And X7 = False And X8 = False And X9 = False And X10 = False And X11 = False Then
                MessageBox.Show("Select a Separator to Split the Cells.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                workSheet.Activate()
                rng.Select()
                Exit Sub
            End If

            If CheckBox2.Checked = True Then
                workSheet.Copy(After:=workBook.Sheets(workSheet.Name))
            End If

            Dim r As Integer = rng.Rows.Count
            Dim c As Integer = rng.Columns.Count
            Dim rng2Address As String

            Dim TotalColumns As Integer

            If X7 Or X8 Or X9 Or X10 Then
                Dim Separator As String = ""
                Dim Columns(r - 1) As Integer
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
                    Columns(i - 1) = CountSeparator(rng.Cells(i, 1).value, Separator)
                Next
                TotalColumns = FindMax(Columns)
                If X2 Then
                    rng2 = workSheet2.Range(rng2.Cells(1, 1), rng2.Cells(r, TotalColumns))
                ElseIf X1 Then
                    rng2 = workSheet2.Range(rng2.Cells(1, 1), rng2.Cells(TotalColumns, r))
                End If
            ElseIf X3 Then
                TotalColumns = 2
                If X2 Then
                    rng2 = workSheet2.Range(rng2.Cells(1, 1), rng2.Cells(r, TotalColumns))
                ElseIf X1 Then
                    rng2 = workSheet2.Range(rng2.Cells(1, 1), rng2.Cells(TotalColumns, r))
                End If
            ElseIf X11 Then
                Dim W As Integer
                If TextBox3.Text = "" Then
                    W = 1
                Else
                    W = Int(TextBox3.Text)
                End If
                Dim Columns(r - 1) As Integer
                For i = 1 To r
                    If Len(rng.Cells(i, 1).value) Mod W = 0 Then
                        Columns(i - 1) = Int(Len(rng.Cells(i, 1).value) / W)
                    Else
                        Columns(i - 1) = Int(Len(rng.Cells(i, 1).value) / W) + 1
                    End If
                Next
                TotalColumns = FindMax(Columns)
                If X2 Then
                    rng2 = workSheet2.Range(rng2.Cells(1, 1), rng2.Cells(r, TotalColumns))
                ElseIf X1 Then
                    rng2 = workSheet2.Range(rng2.Cells(1, 1), rng2.Cells(TotalColumns, r))
                End If
            End If

            rng2Address = rng2.Address

            If Overlap(excelApp, workSheet, workSheet2, rng, rng2) = False Then

                If (X1 Or X2) And (X3 Or X7 Or X8 Or X9 Or X10 Or X11) Then

                    Dim SplitColumn As Integer = 1

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

                        If X2 Then

                            Dim Index As Integer
                            Dim position As Integer

                            For i = 1 To r
                                Dim source As String = rng.Cells(i, SplitColumn).value
                                position = 1
                                Index = 0
                                For k = 1 To Len(source)
                                    If Mid(source, k, Len(Separator)) = Separator Then
                                        If k - position > 0 Then
                                            Index = Index + 1
                                            rng2.Cells(i, Index).value = Mid(source, position, k - position)
                                        End If
                                        position = k + Len(Separator)
                                    End If
                                Next
                                If position <= Len(source) Then
                                    Index = Index + 1
                                    rng2.Cells(i, Index).value = Mid(source, position, Len(source) - position + 1)
                                End If

                                If CheckBox1.Checked = True Then
                                    For m = 1 To TotalColumns
                                        rng.Cells(i, SplitColumn).copy
                                        rng2.Cells(i, m).PasteSpecial(Excel.XlPasteType.xlPasteFormats)
                                        rng2 = workSheet2.Range(rng2Address)
                                        workSheet2.Activate()
                                    Next
                                Else
                                    For m = 1 To TotalColumns
                                        rng2.Cells(i, m).ClearFormats()
                                    Next
                                End If

                            Next
                            excelApp.CutCopyMode = False

                        ElseIf X1 Then

                            Dim Index As Integer
                            Dim position As Integer

                            For i = 1 To r
                                Dim source As String = rng.Cells(i, SplitColumn).value
                                position = 1
                                Index = 0
                                For k = 1 To Len(source)
                                    If Mid(source, k, Len(Separator)) = Separator Then
                                        If k - position > 0 Then
                                            Index = Index + 1
                                            rng2.Cells(Index, i).value = Mid(source, position, k - position)
                                        End If
                                        position = k + Len(Separator)
                                    End If
                                Next
                                If position <= Len(source) Then
                                    Index = Index + 1
                                    rng2.Cells(Index, i).value = Mid(source, position, Len(source) - position + 1)
                                End If
                                If CheckBox1.Checked = True Then
                                    For m = 1 To TotalColumns
                                        rng.Cells(i, SplitColumn).copy
                                        rng2.Cells(m, i).PasteSpecial(Excel.XlPasteType.xlPasteFormats)
                                        rng2 = workSheet2.Range(rng2Address)
                                        workSheet2.Activate()
                                    Next
                                Else
                                    For m = 1 To TotalColumns
                                        rng2.Cells(m, i).ClearFormats()
                                    Next
                                End If
                            Next
                            excelApp.CutCopyMode = False
                        End If

                    ElseIf X3 Then

                        If X2 Then

                            For i = 1 To r
                                Dim source As String = rng.Cells(i, SplitColumn).value
                                Dim NumberText(1) As String
                                NumberText = SeparateNumberText(source)
                                rng2.Cells(i, 1).value = NumberText(0)
                                rng2.Cells(i, 2).value = NumberText(1)
                                If CheckBox1.Checked = True Then
                                    For m = 1 To TotalColumns
                                        rng.Cells(i, SplitColumn).copy
                                        rng2.Cells(i, m).PasteSpecial(Excel.XlPasteType.xlPasteFormats)
                                        rng2 = workSheet2.Range(rng2Address)
                                        workSheet2.Activate()
                                    Next
                                Else
                                    For m = 1 To TotalColumns
                                        rng2.Cells(i, m).ClearFormats()
                                    Next
                                End If
                            Next
                            excelApp.CutCopyMode = False

                        ElseIf X1 Then

                            For i = 1 To r
                                Dim source As String = rng.Cells(i, SplitColumn).value
                                Dim NumberText(1) As String
                                NumberText = SeparateNumberText(source)
                                rng2.Cells(1, i).value = NumberText(0)
                                rng2.Cells(2, i).value = NumberText(1)
                                If CheckBox1.Checked = True Then
                                    For m = 1 To TotalColumns
                                        rng.Cells(i, SplitColumn).copy
                                        rng2.Cells(m, i).PasteSpecial(Excel.XlPasteType.xlPasteFormats)
                                        rng2 = workSheet2.Range(rng2Address)
                                        workSheet2.Activate()
                                    Next
                                Else
                                    For m = 1 To TotalColumns
                                        rng2.Cells(m, i).ClearFormats()
                                    Next
                                End If
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

                        If X2 Then

                            Dim Index As Integer

                            For i = 1 To r
                                Dim source As String = rng.Cells(i, SplitColumn).value
                                Index = 0
                                For k = 1 To Int(Len(source) / W)
                                    Index = Index + 1
                                    rng.Cells(i, Index).value = Mid(source, (W * (k - 1)) + 1, W)
                                Next
                                If Len(source) Mod W <> 0 Then
                                    Index = Index + 1
                                    rng.Cells(i, Index).value = Mid(source, Len(source) - (Len(source) Mod W) + 1, Len(source) Mod W)
                                End If
                                If CheckBox1.Checked = True Then
                                    For m = 1 To TotalColumns
                                        rng.Cells(i, SplitColumn).copy
                                        rng2.Cells(i, m).PasteSpecial(Excel.XlPasteType.xlPasteFormats)
                                        rng2 = workSheet2.Range(rng2Address)
                                        workSheet2.Activate()
                                    Next
                                Else
                                    For m = 1 To TotalColumns
                                        rng2.Cells(i, m).ClearFormats()
                                    Next
                                End If
                            Next
                            excelApp.CutCopyMode = False

                        ElseIf X1 Then

                            Dim Index As Integer

                            For i = 1 To r
                                Dim source As String = rng.Cells(i, SplitColumn).value
                                Index = 0
                                For k = 1 To Int(Len(source) / W)
                                    Index = Index + 1
                                    rng.Cells(Index, i).value = Mid(source, (W * (k - 1)) + 1, W)
                                Next
                                If Len(source) Mod W <> 0 Then
                                    Index = Index + 1
                                    rng.Cells(Index, i).value = Mid(source, Len(source) - (Len(source) Mod W) + 1, Len(source) Mod W)
                                End If
                                If CheckBox1.Checked = True Then
                                    For m = 1 To TotalColumns
                                        rng.Cells(i, SplitColumn).copy
                                        rng2.Cells(m, i).PasteSpecial(Excel.XlPasteType.xlPasteFormats)
                                        rng2 = workSheet2.Range(rng2Address)
                                        workSheet2.Activate()
                                    Next
                                Else
                                    For m = 1 To TotalColumns
                                        rng2.Cells(m, i).ClearFormats()
                                    Next
                                End If
                            Next
                            excelApp.CutCopyMode = False
                        End If

                    End If

                End If

                rng2.Select()
                For j = 1 To rng2.Columns.Count
                    rng2.Columns(j).Autofit
                Next

                Me.Close()

            Else
                If (X1 Or X2) And (X3 Or X7 Or X8 Or X9 Or X10 Or X11) Then

                    Dim SplitColumn As Integer = 1

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

                        If X2 Then

                            Dim Index As Integer
                            Dim position As Integer

                            For i = 1 To r
                                Dim source As String = rng.Cells(i, SplitColumn).value
                                position = 1
                                Index = 0
                                For k = 1 To Len(source)
                                    If Mid(source, k, Len(Separator)) = Separator Then
                                        If k - position > 0 Then
                                            Index = Index + 1
                                            rng2.Cells(i, Index).value = Mid(source, position, k - position)
                                        End If
                                        position = k + Len(Separator)
                                    End If
                                Next
                                If position <= Len(source) Then
                                    Index = Index + 1
                                    rng2.Cells(i, Index).value = Mid(source, position, Len(source) - position + 1)
                                End If
                                If CheckBox1.Checked = True Then

                                    Dim x As Integer = i - 1
                                    Dim y As Integer = SplitColumn - 1

                                    workSheet2.Range(rng2.Cells(i, 1), rng2.Cells(i, TotalColumns)).Font.Name = FontNames(x, y)
                                    workSheet2.Range(rng2.Cells(i, 1), rng2.Cells(i, TotalColumns)).Font.Size = FontSizes(x, y)

                                    If FontBolds(x, y) Then workSheet2.Range(rng2.Cells(i, 1), rng2.Cells(i, TotalColumns)).Font.Bold = True
                                    If Fontitalics(x, y) Then workSheet2.Range(rng2.Cells(i, 1), rng2.Cells(i, TotalColumns)).Font.Italic = True

                                    workSheet2.Range(rng2.Cells(i, 1), rng2.Cells(i, TotalColumns)).Interior.Color = System.Drawing.Color.FromArgb(Red1s(x, y), Green1s(x, y), Blue1s(x, y))

                                    workSheet2.Range(rng2.Cells(i, 1), rng2.Cells(i, TotalColumns)).Font.Color = System.Drawing.Color.FromArgb(Red2s(x, y), Green2s(x, y), Blue2s(x, y))
                                    workSheet2.Activate()

                                Else
                                    workSheet2.Range(rng2.Cells(i, 1), rng2.Cells(i, TotalColumns)).ClearFormats()

                                End If
                            Next

                        ElseIf X1 Then

                            Dim Index As Integer
                            Dim position As Integer

                            For i = 1 To r
                                Dim source As String = rng.Cells(i, SplitColumn).value
                                position = 1
                                Index = 0
                                For k = 1 To Len(source)
                                    If Mid(source, k, Len(Separator)) = Separator Then
                                        If k - position > 0 Then
                                            Index = Index + 1
                                            rng2.Cells(Index, i).value = Mid(source, position, k - position)
                                        End If
                                        position = k + Len(Separator)
                                    End If
                                Next
                                If position <= Len(source) Then
                                    Index = Index + 1
                                    rng2.Cells(Index, i).value = Mid(source, position, Len(source) - position + 1)
                                End If
                                If CheckBox1.Checked = True Then

                                    Dim x As Integer = i - 1
                                    Dim y As Integer = SplitColumn - 1

                                    workSheet2.Range(rng2.Cells(1, i), rng2.Cells(TotalColumns, i)).Font.Name = FontNames(x, y)
                                    workSheet2.Range(rng2.Cells(1, i), rng2.Cells(TotalColumns, i)).Font.Size = FontSizes(x, y)

                                    If FontBolds(x, y) Then workSheet2.Range(rng2.Cells(1, i), rng2.Cells(TotalColumns, i)).Font.Bold = True
                                    If Fontitalics(x, y) Then workSheet2.Range(rng2.Cells(1, i), rng2.Cells(TotalColumns, i)).Font.Italic = True

                                    workSheet2.Range(rng2.Cells(1, i), rng2.Cells(TotalColumns, i)).Interior.Color = System.Drawing.Color.FromArgb(Red1s(x, y), Green1s(x, y), Blue1s(x, y))

                                    workSheet2.Range(rng2.Cells(1, i), rng2.Cells(TotalColumns, i)).Font.Color = System.Drawing.Color.FromArgb(Red2s(x, y), Green2s(x, y), Blue2s(x, y))
                                    workSheet2.Activate()
                                Else
                                    workSheet2.Range(rng2.Cells(1, i), rng2.Cells(TotalColumns, i)).ClearFormats()

                                End If
                            Next
                        End If

                    ElseIf X3 Then

                        If X2 Then

                            For i = 1 To r
                                Dim source As String = rng.Cells(i, SplitColumn).value
                                Dim NumberText(1) As String
                                NumberText = SeparateNumberText(source)
                                rng2.Cells(i, 1).value = NumberText(0)
                                rng2.Cells(i, 2).value = NumberText(1)
                                If CheckBox1.Checked = True Then
                                    Dim x As Integer = i - 1
                                    Dim y As Integer = SplitColumn - 1

                                    workSheet2.Range(rng2.Cells(i, 1), rng2.Cells(i, 2)).Font.Name = FontNames(x, y)
                                    workSheet2.Range(rng2.Cells(i, 1), rng2.Cells(i, 2)).Font.Size = FontSizes(x, y)

                                    If FontBolds(x, y) Then workSheet2.Range(rng2.Cells(i, 1), rng2.Cells(i, 2)).Font.Bold = True
                                    If Fontitalics(x, y) Then workSheet2.Range(rng2.Cells(i, 1), rng2.Cells(i, 2)).Font.Italic = True

                                    workSheet2.Range(rng2.Cells(i, 1), rng2.Cells(i, 2)).Interior.Color = System.Drawing.Color.FromArgb(Red1s(x, y), Green1s(x, y), Blue1s(x, y))

                                    workSheet2.Range(rng2.Cells(i, 1), rng2.Cells(i, 2)).Font.Color = System.Drawing.Color.FromArgb(Red2s(x, y), Green2s(x, y), Blue2s(x, y))
                                    workSheet2.Activate()
                                Else
                                    workSheet2.Range(rng2.Cells(i, 1), rng2.Cells(i, 2)).ClearFormats()

                                End If
                            Next

                        ElseIf X1 Then

                            For i = 1 To r
                                Dim source As String = rng.Cells(i, SplitColumn).value
                                Dim NumberText(1) As String
                                NumberText = SeparateNumberText(source)
                                rng2.Cells(1, i).value = NumberText(0)
                                rng2.Cells(2, i).value = NumberText(1)
                                If CheckBox1.Checked = True Then
                                    Dim x As Integer = i - 1
                                    Dim y As Integer = SplitColumn - 1

                                    workSheet2.Range(rng2.Cells(1, i), rng2.Cells(2, i)).Font.Name = FontNames(x, y)
                                    workSheet2.Range(rng2.Cells(1, i), rng2.Cells(2, i)).Font.Size = FontSizes(x, y)

                                    If FontBolds(x, y) Then workSheet2.Range(rng2.Cells(1, i), rng2.Cells(2, i)).Font.Bold = True
                                    If Fontitalics(x, y) Then workSheet2.Range(rng2.Cells(1, i), rng2.Cells(2, i)).Font.Italic = True

                                    workSheet2.Range(rng2.Cells(1, i), rng2.Cells(2, i)).Interior.Color = System.Drawing.Color.FromArgb(Red1s(x, y), Green1s(x, y), Blue1s(x, y))

                                    workSheet2.Range(rng2.Cells(1, i), rng2.Cells(2, i)).Font.Color = System.Drawing.Color.FromArgb(Red2s(x, y), Green2s(x, y), Blue2s(x, y))
                                    workSheet2.Activate()
                                Else
                                    workSheet2.Range(rng2.Cells(1, i), rng2.Cells(2, i)).ClearFormats()
                                End If
                            Next
                        End If

                    ElseIf X11 Then

                        Dim W As Integer

                        If TextBox3.Text = "" Then
                            W = 1
                        Else
                            W = Int(TextBox3.Text)
                        End If

                        If X2 Then

                            Dim Index As Integer

                            For i = 1 To r
                                Dim source As String = rng.Cells(i, SplitColumn).value
                                Index = 0
                                For k = 1 To Int(Len(source) / W)
                                    Index = Index + 1
                                    rng.Cells(i, Index).value = Mid(source, (W * (k - 1)) + 1, W)
                                Next
                                If Len(source) Mod W <> 0 Then
                                    Index = Index + 1
                                    rng.Cells(i, Index).value = Mid(source, Len(source) - (Len(source) Mod W) + 1, Len(source) Mod W)
                                End If
                                If CheckBox1.Checked = True Then
                                    Dim x As Integer = i - 1
                                    Dim y As Integer = SplitColumn - 1

                                    workSheet2.Range(rng2.Cells(i, 1), rng2.Cells(i, TotalColumns)).Font.Name = FontNames(x, y)
                                    workSheet2.Range(rng2.Cells(i, 1), rng2.Cells(i, TotalColumns)).Font.Size = FontSizes(x, y)

                                    If FontBolds(x, y) Then workSheet2.Range(rng2.Cells(i, 1), rng2.Cells(i, TotalColumns)).Font.Bold = True
                                    If Fontitalics(x, y) Then workSheet2.Range(rng2.Cells(i, 1), rng2.Cells(i, TotalColumns)).Font.Italic = True

                                    workSheet2.Range(rng2.Cells(i, 1), rng2.Cells(i, TotalColumns)).Interior.Color = System.Drawing.Color.FromArgb(Red1s(x, y), Green1s(x, y), Blue1s(x, y))

                                    workSheet2.Range(rng2.Cells(i, 1), rng2.Cells(i, TotalColumns)).Font.Color = System.Drawing.Color.FromArgb(Red2s(x, y), Green2s(x, y), Blue2s(x, y))
                                    workSheet2.Activate()
                                Else
                                    workSheet2.Range(rng2.Cells(i, 1), rng2.Cells(i, TotalColumns)).ClearFormats()
                                End If
                            Next

                        ElseIf X1 Then

                            Dim Index As Integer

                            For i = 1 To r
                                Dim source As String = rng.Cells(i, SplitColumn).value
                                Index = 0
                                For k = 1 To Int(Len(source) / W)
                                    Index = Index + 1
                                    rng.Cells(Index, i).value = Mid(source, (W * (k - 1)) + 1, W)
                                Next
                                If Len(source) Mod W <> 0 Then
                                    Index = Index + 1
                                    rng.Cells(Index, i).value = Mid(source, Len(source) - (Len(source) Mod W) + 1, Len(source) Mod W)
                                End If
                                If CheckBox1.Checked = True Then
                                    Dim x As Integer = i - 1
                                    Dim y As Integer = SplitColumn - 1

                                    workSheet2.Range(rng2.Cells(1, i), rng2.Cells(TotalColumns, i)).Font.Name = FontNames(x, y)
                                    workSheet2.Range(rng2.Cells(1, i), rng2.Cells(TotalColumns, i)).Font.Size = FontSizes(x, y)

                                    If FontBolds(x, y) Then workSheet2.Range(rng2.Cells(1, i), rng2.Cells(TotalColumns, i)).Font.Bold = True
                                    If Fontitalics(x, y) Then workSheet2.Range(rng2.Cells(1, i), rng2.Cells(TotalColumns, i)).Font.Italic = True

                                    workSheet2.Range(rng2.Cells(1, i), rng2.Cells(TotalColumns, i)).Interior.Color = System.Drawing.Color.FromArgb(Red1s(x, y), Green1s(x, y), Blue1s(x, y))

                                    workSheet2.Range(rng2.Cells(1, i), rng2.Cells(TotalColumns, i)).Font.Color = System.Drawing.Color.FromArgb(Red2s(x, y), Green2s(x, y), Blue2s(x, y))
                                    workSheet2.Activate()
                                Else
                                    workSheet2.Range(rng2.Cells(1, i), rng2.Cells(TotalColumns, i)).ClearFormats()
                                End If
                            Next
                        End If

                    End If

                End If

                rng2.Select()

                For j = 1 To rng2.Columns.Count
                    rng2.Columns(j).Autofit
                Next

                Me.Close()

            End If

            TextBoxChanged = False

        Catch ex As Exception

        End Try

    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged

        Try
            If RadioButton1.Checked = True Then
                Call Display()
            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        Try
            If RadioButton2.Checked = True Then
                Call Display()
            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

        Try

            excelApp = Globals.ThisAddIn.Application
            workBook = excelApp.ActiveWorkbook
            workSheet = workBook.ActiveSheet

            Dim rngArray() As String = Split(TextBox1.Text, "!")
            Dim rngAddress As String = rngArray(UBound(rngArray))
            rng = workSheet.Range(rngAddress)
            TextBoxChanged = True
            rng.Select()
            Call Display()
            TextBoxChanged = False

        Catch ex As Exception

        End Try

    End Sub

    Private Sub AutoSelection_Click(sender As Object, e As EventArgs) Handles AutoSelection.Click

        Try
            FocusedTextBox = 1

            Dim activeRange As Excel.Range = excelApp.ActiveCell

            Dim startRow As Integer = activeRange.Row
            Dim startColumn As Integer = activeRange.Column
            Dim endRow As Integer = activeRange.Row
            Dim endColumn As Integer = activeRange.Column

            'Find the upper boundary
            Do While startRow > 1 AndAlso Not IsNothing(workSheet.Cells(startRow - 1, startColumn).Value)
                startRow -= 1
            Loop

            'Find the lower boundary
            Do While Not IsNothing(workSheet.Cells(endRow + 1, endColumn).Value)
                endRow += 1
            Loop

            'Select the determined range
            rng = workSheet.Range(workSheet.Cells(startRow, startColumn), workSheet.Cells(endRow, endColumn))

            rng.Select()

            Dim sheetName As String

            sheetName = Split(rng.Address(True, True, Excel.XlReferenceStyle.xlA1, True), "]")(1)
            sheetName = Split(sheetName, "!")(0)

            If Mid(sheetName, Len(sheetName), 1) = "'" Then
                sheetName = Mid(sheetName, 1, Len(sheetName) - 1)
            End If

            workSheet = workBook.Worksheets(sheetName)
            workSheet.Activate()

            If workSheet.Name <> OpenSheet.Name Then
                TextBox1.Text = workSheet.Name & "!" & rng.Address
            Else
                TextBox1.Text = rng.Address
            End If

            Me.TextBox1.Focus()

        Catch ex As Exception

            Me.Show()
            TextBox1.Focus()

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

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged

        Try
            excelApp = Globals.ThisAddIn.Application
            workBook = excelApp.ActiveWorkbook
            workSheet2 = workBook.ActiveSheet

            Dim rng2Array() As String = Split(TextBox4.Text, "!")
            Dim rng2Address As String = rng2Array(UBound(rng2Array))
            rng2 = workSheet2.Range(rng2Address)

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

            Dim userInput As Excel.Range = excelApp.InputBox("Select a Cell.", Type:=8)
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

            If workSheet2.Name <> OpenSheet.Name Then
                TextBox4.Text = workSheet2.Name & "!" & rng2.Address
            Else
                TextBox4.Text = rng2.Address
            End If

            Me.Show()
            TextBox4.Focus()

        Catch ex As Exception

            Me.Show()
            TextBox4.Focus()

        End Try

    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        Try
            Call Display()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
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

    Private Sub Form24_Split_Cells_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try

            excelApp = Globals.ThisAddIn.Application
            workBook = excelApp.ActiveWorkbook
            workSheet = workBook.ActiveSheet
            workSheet2 = workBook.ActiveSheet
            Me.KeyPreview = True

            AddHandler excelApp.SheetSelectionChange, AddressOf excelApp_SheetSelectionChange

            opened = opened + 1

        Catch ex As Exception

        End Try
    End Sub

    Private Sub excelApp_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Excel.Range)

        Try

            Dim selectedRange As Excel.Range
            selectedRange = excelApp.Selection

            If TextBoxChanged = False Then
                If FocusedTextBox = 1 Then
                    workSheet = workBook.ActiveSheet
                    If workSheet.Name <> OpenSheet.Name Then
                        TextBox1.Text = workSheet.Name & "!" & selectedRange.Address
                    Else
                        TextBox1.Text = selectedRange.Address
                    End If
                    rng = selectedRange
                    TextBox1.Focus()

                ElseIf FocusedTextBox = 4 Then
                    workSheet2 = workBook.ActiveSheet
                    If workSheet2.Name <> OpenSheet.Name Then
                        TextBox4.Text = workSheet2.Name & "!" & selectedRange.Address
                    Else
                        TextBox4.Text = selectedRange.Address
                    End If
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

    Private Sub TextBox4_GotFocus(sender As Object, e As EventArgs) Handles TextBox4.GotFocus
        Try
            FocusedTextBox = 4
        Catch ex As Exception

        End Try
    End Sub

    Private Sub PictureBox3_GotFocus(sender As Object, e As EventArgs) Handles PictureBox3.GotFocus
        Try
            FocusedTextBox = 4
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Form24_Split_Cells_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Me.Close()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub ComboBox2_TextChanged(sender As Object, e As EventArgs) Handles ComboBox2.TextChanged
        Try
            Call Display()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Form24_Split_Cells_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        form_flag = False
    End Sub

    Private Sub Form24_Split_Cells_Disposed(sender As Object, e As EventArgs) Handles Me.Disposed
        form_flag = False
    End Sub

    Private Sub Form24_Split_Cells_Shown(sender As Object, e As EventArgs) Handles Me.Shown

        Try

            Me.Focus()
        Me.BringToFront()
        Me.Activate()

        Dim TextBoxText As String

        If workSheet.Name <> OpenSheet.Name Then
            TextBoxText = workSheet.Name & "!" & rng.Address
        Else
            TextBoxText = rng.Address
        End If

        Me.BeginInvoke(New System.Action(Sub()
                                             TextBox1.Text = TextBoxText
                                             SetWindowPos(Me.Handle, New IntPtr(HWND_TOPMOST), 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOSIZE)
                                         End Sub))

        Catch ex As Exception

        End Try

    End Sub

End Class