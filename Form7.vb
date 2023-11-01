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
Imports System.Text.RegularExpressions
Imports System.ComponentModel

Public Class Form7

    Dim WithEvents excelApp As Excel.Application
    Dim workbook As Excel.Workbook
    Dim workbook2 As Excel.Workbook
    Dim worksheet As Excel.Worksheet
    Dim worksheet2 As Excel.Worksheet
    Public OpenSheet As Excel.Worksheet
    Dim rng As Excel.Range
    Dim rng2 As Excel.Range
    Dim FocusedTextBox As Integer
    Dim opened As Integer
    Dim TextBoxChanged As Boolean


    Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As IntPtr, ByVal hWndInsertAfter As IntPtr, ByVal X As Integer, ByVal Y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal uFlags As UInteger) As Boolean
    Private Const SWP_NOMOVE As UInteger = &H2
    Private Const SWP_NOSIZE As UInteger = &H1
    Private Const SWP_NOACTIVATE As UInteger = &H10
    Private Const HWND_TOPMOST As Integer = -1

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
    Private Function CanConvertToInt(input As String) As Boolean

        Dim number As Integer

        ' TryParse returns True if conversion is possible, False otherwise
        If Integer.TryParse(input, number) Then
            Return True
        Else
            Return False
        End If

    End Function
    Private Function IsValidExcelCellReference(cellReference As String) As Boolean

        Dim cellPattern As String = "(\$?[A-Z]+\$?[0-9]+)"
        Dim referencePattern As String = "^" + cellPattern + "(:" + cellPattern + ")?$"

        Dim regex As New Regex(referencePattern)

        Dim refArr() As String = Split(cellReference, "!")

        Dim reference As String = refArr(UBound(refArr))

        If regex.IsMatch(reference) Then
            Return True
        Else
            Return False
        End If

    End Function

    Private Sub Setup()

        Try
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

        Catch ex As Exception

        End Try

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

        Dim Arr2(0) As Integer
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
    Private Function MaxOfColumn(cRng As Excel.Range)

        Dim max As Integer
        Dim CharNumbers As Integer

        If IsNumeric(cRng.Cells(1, 1).value) Then
            max = Len(Str(cRng.Cells(1, 1).value))
        Else
            max = Len(cRng.Cells(1, 1).value)
        End If

        For i = 2 To cRng.Rows.Count
            If IsNumeric(cRng.Cells(i, 1).value) Then
                CharNumbers = Len(Str(cRng.Cells(i, 1).value))
            Else
                CharNumbers = Len(cRng.Cells(i, 1).value)
            End If
            If CharNumbers > max Then
                max = CharNumbers
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
    Private Function AdjustWidth(Widths, CWidth)

        Dim SumWidth As Double = 0

        For i = LBound(Widths) To UBound(Widths)
            SumWidth = SumWidth + Widths(i)
        Next

        If SumWidth < CWidth Then
            Dim Extra As Double = CWidth - SumWidth
            Extra = Extra / (UBound(Widths) + 1)
            For i = LBound(Widths) To UBound(Widths)
                Widths(i) = Widths(i) + Extra
            Next
        End If

        AdjustWidth = Widths

    End Function

    Private Sub Display()

        Try

            TextBoxChanged = True

            CustomPanel1.Controls.Clear()
            CustomPanel2.Controls.Clear()

            Dim displayRng As Excel.Range

            displayRng = rng

            If rng.Rows.Count > 50 Then
                displayRng = worksheet.Range(rng.Cells(1, 1), rng.Cells(50, rng.Columns.Count))
            End If

            If rng.Columns.Count > 50 Then
                displayRng = worksheet.Range(rng.Cells(1, 1), rng.Cells(rng.Rows.Count, 50))
            End If

            Dim r As Integer
            Dim c As Integer

            r = displayRng.Rows.Count
            c = displayRng.Columns.Count

            Dim height As Double
            Dim BaseWidth As Double
            Dim width As Double

            If r > 1 And r <= 6 Then
                height = CustomPanel1.Height / r
            Else
                height = CustomPanel1.Height / 6
            End If

            BaseWidth = 260 / 3

            Dim Ordinate As Double = 0
            Dim Widths(c - 1) As Double

            For j = 1 To c
                Dim CRng As Excel.Range = worksheet.Range(displayRng.Cells(1, j), displayRng.Cells(r, j))
                Widths(j - 1) = (MaxOfColumn(CRng) * BaseWidth) / 10
            Next

            Widths = AdjustWidth(Widths, CustomPanel2.Width)

            For j = 1 To c
                For i = 1 To r
                    Dim label As New System.Windows.Forms.Label
                    label.Text = displayRng.Cells(i, j).Value
                    label.Location = New System.Drawing.Point(Ordinate, (i - 1) * height)
                    label.Height = height
                    label.Width = Widths(j - 1)
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
                Ordinate = Ordinate + Widths(j - 1)
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

                Dim values(displayRng.Cells.Count - 1) As Object
                For k = 1 To displayRng.Cells.Count
                    values(k - 1) = displayRng.Cells(k).value
                Next

                Dim Widths2(0) As Double
                Widths2(0) = (MaxOfArray(values) * BaseWidth) / 10

                Widths2 = AdjustWidth(Widths2, CustomPanel2.Width)

                Dim count As Integer
                count = 1

                If X5 Then

                    For i = 1 To r
                        For j = 1 To c
                            Dim label As New System.Windows.Forms.Label
                            label.Text = displayRng.Cells(i, j).Value
                            label.Location = New System.Drawing.Point(0, (count - 1) * height)
                            count = count + 1
                            label.Height = height
                            label.Width = Widths2(0)
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

                            CustomPanel2.Controls.Add(label)
                        Next
                    Next

                ElseIf X6 Then

                    For j = 1 To c
                        For i = 1 To r
                            Dim label As New System.Windows.Forms.Label
                            label.Text = displayRng.Cells(i, j).Value
                            label.Location = New System.Drawing.Point(0, (count - 1) * height)
                            count = count + 1
                            label.Height = height
                            label.Width = Widths2(0)
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
                    Ordinate = 0
                    Dim Length As Integer
                    For i = 1 To r
                        For j = 1 To c
                            Length = Len(displayRng.Cells(i, j).Value)
                            If Length < 7 Then
                                Length = 7
                            End If
                            width = (Length * BaseWidth) / 10
                            Dim label As New System.Windows.Forms.Label
                            label.Text = displayRng.Cells(i, j).Value
                            label.Location = New System.Drawing.Point(Ordinate, (3.5 - 1) * height)
                            count = count + 1
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
                            Ordinate = Ordinate + width
                        Next
                    Next

                ElseIf X6 Then
                    Ordinate = 0
                    Dim Length As Integer
                    For j = 1 To c
                        For i = 1 To r
                            Length = Len(displayRng.Cells(i, j).Value)
                            If Length < 7 Then
                                Length = 7
                            End If
                            width = (Length * BaseWidth) / 10
                            Dim label As New System.Windows.Forms.Label
                            label.Text = displayRng.Cells(i, j).Value
                            label.Location = New System.Drawing.Point(Ordinate, (3.5 - 1) * height)
                            count = count + 1
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
                            Ordinate = Ordinate + width
                        Next
                    Next
                End If

                CustomPanel2.AutoScroll = True

            End If

            If X3 Then

                If X7 And (X5 Or X6) Then

                    Dim BreakPoints() As Integer
                    BreakPoints = GetBreakPoints(displayRng, 2)

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

                    Dim Values(r - 1, c - 1) As Object
                    Dim References(r - 1, c - 1) As Integer

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
                                    Values(i - 1, j - 1) = displayRng.Cells(x, y).Value
                                Else
                                    Values(i - 1, j - 1) = ""
                                End If
                                References(i - 1, j - 1) = x
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
                                    Values(i - 1, j - 1) = displayRng.Cells(x, y).Value
                                Else
                                    Values(i - 1, j - 1) = ""
                                End If
                                References(i - 1, j - 1) = x
                            Next
                            iRow = BreakPoints(j - 1)
                        Next
                    End If

                    Dim ColumnValues(r - 1) As Object
                    Dim widths2(c - 1) As Double

                    For j = 0 To c - 1
                        For i = 0 To r - 1
                            ColumnValues(i) = Values(i, j)
                        Next
                        widths2(j) = (MaxOfArray(ColumnValues) * BaseWidth) / 10
                    Next
                    widths2 = AdjustWidth(widths2, CustomPanel2.Width)

                    Ordinate = 0

                    For j = 1 To c
                        For i = 1 To r
                            Dim label As New System.Windows.Forms.Label
                            label.Text = Values(i - 1, j - 1)
                            label.Location = New System.Drawing.Point(Ordinate, (i - 1) * height)
                            label.Height = height
                            label.Width = widths2(j - 1)
                            label.BorderStyle = BorderStyle.FixedSingle
                            label.TextAlign = ContentAlignment.MiddleCenter

                            If CheckBox1.Checked = True Then
                                Dim x As Integer
                                Dim y As Integer
                                x = References(i - 1, j - 1)
                                y = 1
                                Dim cell As Excel.Range = displayRng.Cells(x, y)
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
                        Ordinate = Ordinate + widths2(j - 1)
                    Next

                ElseIf (X8 And TextBox2.Text <> "" And CanConvertToInt(TextBox2.Text) = True) And (X5 Or X6) Then

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

                        Dim Values(r - 1, c - 1) As Object
                        Dim References(r - 1, c - 1) As Integer

                        For i = 1 To r
                            For j = 1 To c
                                Dim x As Integer
                                Dim y As Integer
                                x = (c * (i - 1)) + j
                                y = 1
                                If x <= displayRng.Rows.Count Then
                                    Values(i - 1, j - 1) = displayRng.Cells(x, y).Value
                                Else
                                    Values(i - 1, j - 1) = ""
                                End If
                                References(i - 1, j - 1) = x
                            Next
                        Next

                        Dim ColumnValues(r - 1) As Object
                        Dim widths2(c - 1) As Double

                        For j = 0 To c - 1
                            For i = 0 To r - 1
                                ColumnValues(i) = Values(i, j)
                            Next
                            widths2(j) = (MaxOfArray(ColumnValues) * BaseWidth) / 10
                        Next
                        widths2 = AdjustWidth(widths2, CustomPanel2.Width)

                        Ordinate = 0

                        For j = 1 To c
                            For i = 1 To r
                                Dim label As New System.Windows.Forms.Label
                                label.Text = Values(i - 1, j - 1)
                                label.Location = New System.Drawing.Point(Ordinate, (i - 1) * height)
                                label.Height = height
                                label.Width = widths2(j - 1)
                                label.BorderStyle = BorderStyle.FixedSingle
                                label.TextAlign = ContentAlignment.MiddleCenter

                                If CheckBox1.Checked = True Then
                                    Dim x As Integer
                                    Dim y As Integer
                                    x = References(i - 1, j - 1)
                                    y = 1
                                    Dim cell As Excel.Range = displayRng.Cells(x, y)
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
                            Ordinate = Ordinate + widths2(j - 1)
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

                        Dim Values(r - 1, c - 1) As Object
                        Dim References(r - 1, c - 1) As Integer

                        For i = 1 To r
                            For j = 1 To c
                                Dim x As Integer
                                Dim y As Integer
                                x = (r * (j - 1)) + i
                                y = 1
                                If x <= displayRng.Rows.Count Then
                                    Values(i - 1, j - 1) = displayRng.Cells(x, y).Value
                                Else
                                    Values(i - 1, j - 1) = ""
                                End If
                                References(i - 1, j - 1) = x
                            Next
                        Next

                        Dim ColumnValues(r - 1) As Object
                        Dim widths2(c - 1) As Double

                        For j = 0 To c - 1
                            For i = 0 To r - 1
                                ColumnValues(i) = Values(i, j)
                            Next
                            widths2(j) = (MaxOfArray(ColumnValues) * BaseWidth) / 10
                        Next
                        widths2 = AdjustWidth(widths2, CustomPanel2.Width)

                        Ordinate = 0

                        For j = 1 To c
                            For i = 1 To r
                                Dim label As New System.Windows.Forms.Label
                                label.Text = Values(i - 1, j - 1)
                                label.Location = New System.Drawing.Point(Ordinate, (i - 1) * height)
                                label.Height = height
                                label.Width = widths2(j - 1)
                                label.BorderStyle = BorderStyle.FixedSingle
                                label.TextAlign = ContentAlignment.MiddleCenter

                                If CheckBox1.Checked = True Then
                                    Dim x As Integer
                                    Dim y As Integer
                                    x = References(i - 1, j - 1)
                                    y = 1
                                    Dim cell As Excel.Range = displayRng.Cells(x, y)
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
                            Ordinate = Ordinate + widths2(j - 1)
                        Next

                    End If
                End If

                CustomPanel2.AutoScroll = True

            End If

            If X4 Then

                If X7 And (X5 Or X6) Then

                    Dim BreakPoints() As Integer
                    BreakPoints = GetBreakPoints(displayRng, 1)

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

                    Dim Values(r - 1, c - 1) As Object
                    Dim References(r - 1, c - 1) As Integer

                    If X5 Then
                        Dim iColumn As Integer
                        iColumn = 0
                        For i = 1 To r
                            For j = 1 To c
                                Dim x As Integer
                                Dim y As Integer
                                x = 1
                                y = iColumn + j
                                If x <= BreakPoints(i - 1) Then
                                    Values(i - 1, j - 1) = displayRng.Cells(x, y).Value
                                Else
                                    Values(i - 1, j - 1) = ""
                                End If
                                References(i - 1, j - 1) = y
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
                                If x <= BreakPoints(j - 1) Then
                                    Values(i - 1, j - 1) = displayRng.Cells(x, y).Value
                                Else
                                    Values(i - 1, j - 1) = ""
                                End If
                                References(i - 1, j - 1) = y
                            Next
                            iColumn = BreakPoints(j - 1)
                        Next
                    End If

                    Dim ColumnValues(r - 1) As Object
                    Dim widths2(c - 1) As Double

                    For j = 0 To c - 1
                        For i = 0 To r - 1
                            ColumnValues(i) = Values(i, j)
                        Next
                        widths2(j) = (MaxOfArray(ColumnValues) * BaseWidth) / 10
                    Next
                    widths2 = AdjustWidth(widths2, CustomPanel2.Width)

                    Ordinate = 0

                    For j = 1 To c
                        For i = 1 To r
                            Dim label As New System.Windows.Forms.Label
                            label.Text = Values(i - 1, j - 1)
                            label.Location = New System.Drawing.Point(Ordinate, (i - 1) * height)
                            label.Height = height
                            label.Width = widths2(j - 1)
                            label.BorderStyle = BorderStyle.FixedSingle
                            label.TextAlign = ContentAlignment.MiddleCenter

                            If CheckBox1.Checked = True Then
                                Dim x As Integer
                                Dim y As Integer
                                x = 1
                                y = References(i - 1, j - 1)
                                Dim cell As Excel.Range = displayRng.Cells(x, y)
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
                        Ordinate = Ordinate + widths2(j - 1)
                    Next


                ElseIf (X8 And TextBox2.Text <> "" And CanConvertToInt(TextBox2.Text) = True) And (X5 Or X6) Then

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

                        Dim Values(r - 1, c - 1) As Object
                        Dim References(r - 1, c - 1) As Integer

                        For i = 1 To r
                            For j = 1 To c
                                Dim x As Integer
                                Dim y As Integer
                                x = 1
                                y = (c * (i - 1)) + j
                                If x <= displayRng.Rows.Count Then
                                    Values(i - 1, j - 1) = displayRng.Cells(x, y).Value
                                Else
                                    Values(i - 1, j - 1) = ""
                                End If
                                References(i - 1, j - 1) = y
                            Next
                        Next

                        Dim ColumnValues(r - 1) As Object
                        Dim widths2(c - 1) As Double

                        For j = 0 To c - 1
                            For i = 0 To r - 1
                                ColumnValues(i) = Values(i, j)
                            Next
                            widths2(j) = (MaxOfArray(ColumnValues) * BaseWidth) / 10
                        Next
                        widths2 = AdjustWidth(widths2, CustomPanel2.Width)

                        Ordinate = 0

                        For j = 1 To c
                            For i = 1 To r
                                Dim label As New System.Windows.Forms.Label
                                label.Text = Values(i - 1, j - 1)
                                label.Location = New System.Drawing.Point(Ordinate, (i - 1) * height)
                                label.Height = height
                                label.Width = widths2(j - 1)
                                label.BorderStyle = BorderStyle.FixedSingle
                                label.TextAlign = ContentAlignment.MiddleCenter

                                If CheckBox1.Checked = True Then
                                    Dim x As Integer
                                    Dim y As Integer
                                    x = 1
                                    y = References(i - 1, j - 1)
                                    Dim cell As Excel.Range = displayRng.Cells(x, y)
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
                            Ordinate = Ordinate + widths2(j - 1)
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

                        Dim Values(r - 1, c - 1) As Object
                        Dim References(r - 1, c - 1) As Integer

                        For i = 1 To r
                            For j = 1 To c
                                Dim x As Integer
                                Dim y As Integer
                                x = 1
                                y = (r * (j - 1)) + i
                                If x <= displayRng.Rows.Count Then
                                    Values(i - 1, j - 1) = displayRng.Cells(x, y).Value
                                Else
                                    Values(i - 1, j - 1) = ""
                                End If
                                References(i - 1, j - 1) = y
                            Next
                        Next

                        Dim ColumnValues(r - 1) As Object
                        Dim widths2(c - 1) As Double

                        For j = 0 To c - 1
                            For i = 0 To r - 1
                                ColumnValues(i) = Values(i, j)
                            Next
                            widths2(j) = (MaxOfArray(ColumnValues) * BaseWidth) / 10
                        Next
                        widths2 = AdjustWidth(widths2, CustomPanel2.Width)

                        Ordinate = 0

                        For j = 1 To c
                            For i = 1 To r
                                Dim label As New System.Windows.Forms.Label
                                label.Text = Values(i - 1, j - 1)
                                label.Location = New System.Drawing.Point(Ordinate, (i - 1) * height)
                                label.Height = height
                                label.Width = widths2(j - 1)
                                label.BorderStyle = BorderStyle.FixedSingle
                                label.TextAlign = ContentAlignment.MiddleCenter

                                If CheckBox1.Checked = True Then
                                    Dim x As Integer
                                    Dim y As Integer
                                    x = 1
                                    y = References(i - 1, j - 1)
                                    Dim cell As Excel.Range = displayRng.Cells(x, y)
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
                            Ordinate = Ordinate + widths2(j - 1)
                        Next

                    End If
                End If

                CustomPanel2.AutoScroll = True

            End If

            TextBoxChanged = False

        Catch ex As Exception

        End Try

    End Sub
    Private Sub Form7_Load(sender As Object, e As EventArgs) Handles Me.Load

        Try

            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            workbook2 = excelApp.ActiveWorkbook
            worksheet = workbook.ActiveSheet
            worksheet2 = workbook2.ActiveSheet
            Me.KeyPreview = True

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
                    worksheet = workbook.ActiveSheet
                    If worksheet.Name <> OpenSheet.Name Then
                        TextBox1.Text = worksheet.Name & "!" & selectedRange.Address
                    Else
                        TextBox1.Text = selectedRange.Address
                    End If
                    rng = selectedRange
                    TextBox1.Focus()

                ElseIf FocusedTextBox = 3 Then
                    worksheet2 = workbook2.ActiveSheet
                    If worksheet2.Name <> OpenSheet.Name Then
                        TextBox3.Text = worksheet2.Name & "!" & selectedRange.Address
                    Else
                        TextBox3.Text = selectedRange.Address
                    End If
                    rng2 = selectedRange
                    TextBox3.Focus()
                End If
            End If

        Catch ex As Exception

        End Try

    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        TextBoxChanged = True
        If TextBox1.Text = "" Then
            MessageBox.Show("Select a Source Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TextBox1.Focus()
            worksheet.Activate()
            rng.Select()
            Exit Sub
        End If

        If IsValidExcelCellReference(TextBox1.Text) = False Then
            MessageBox.Show("Select a Valid Source Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            TextBox1.Focus()
            worksheet.Activate()
            rng.Select()
            Exit Sub
        End If

        If (RadioButton1.Checked = False And RadioButton2.Checked = False And RadioButton3.Checked = False And RadioButton4.Checked = False) Then
            MessageBox.Show("Select a Transformation Type.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            worksheet.Activate()
            rng.Select()
            Exit Sub
        End If

        If (RadioButton5.Checked = False And RadioButton6.Checked = False) Then
            MessageBox.Show("Select a Transformation Option.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            worksheet.Activate()
            rng.Select()
            Exit Sub
        End If

        If (RadioButton3.Checked = True Or RadioButton4.Checked = True) And (RadioButton7.Checked = False And RadioButton8.Checked = False) Then
            MessageBox.Show("Select a Separator.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            worksheet.Activate()
            rng.Select()
            Exit Sub
        End If

        If RadioButton8.Checked = True And (TextBox2.Text = "" Or CanConvertToInt(TextBox2.Text) = False) Then
            Dim Texts() As String
            Texts = Split(RadioButton8.Text, " ")
            Dim iText As String = Texts(UBound(Texts))
            MessageBox.Show("Enter a valid Number of " & iText & ".", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            worksheet.Activate()
            rng.Select()
            TextBox2.Focus()
            Exit Sub
        End If

        If (RadioButton9.Checked = False And RadioButton10.Checked = False) Then
            MessageBox.Show("Select a Destination Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            worksheet.Activate()
            rng.Select()
            Exit Sub
        End If

        If RadioButton10.Checked = True And (TextBox3.Text = "" Or IsValidExcelCellReference(TextBox3.Text) = False) Then
            MessageBox.Show("Enter a valid Destination Cell.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            worksheet.Activate()
            rng.Select()
            Exit Sub
        End If

        If CheckBox2.Checked = True Then
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
                    excelApp.CutCopyMode = Excel.XlCutCopyMode.xlCopy
                Else
                    MessageBox.Show("Choose One Transformation Option. ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

            Else

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

                            If IsDBNull(font.Name) = False Then
                                fontNames(i - 1, j - 1) = font.Name
                            Else
                                fontNames(i - 1, j - 1) = "Calibri"
                            End If

                            If IsDBNull(font.Size) = False Then
                                Dim fontSize As Single = Convert.ToSingle(font.Size)
                                fontSizes(i - 1, j - 1) = fontSize
                            Else
                                fontSizes(i - 1, j - 1) = 11
                            End If

                            If IsDBNull(cell.Interior.Color) Then
                                reds1(i - 1, j - 1) = 255
                                greens1(i - 1, j - 1) = 255
                                blues1(i - 1, j - 1) = 255
                            Else
                                Dim colorValue1 As Long = CLng(cell.Interior.Color)
                                Dim red1 As Integer = colorValue1 Mod 256
                                Dim green1 As Integer = (colorValue1 \ 256) Mod 256
                                Dim blue1 As Integer = (colorValue1 \ 256 \ 256) Mod 256
                                reds1(i - 1, j - 1) = red1
                                greens1(i - 1, j - 1) = green1
                                blues1(i - 1, j - 1) = blue1
                            End If

                            If IsDBNull(cell.Font.Color) Then
                                reds2(i - 1, j - 1) = 0
                                greens2(i - 1, j - 1) = 0
                                blues2(i - 1, j - 1) = 0
                            Else
                                Dim colorValue2 As Long = CLng(cell.Font.Color)
                                Dim red2 As Integer = colorValue2 Mod 256
                                Dim green2 As Integer = (colorValue2 \ 256) Mod 256
                                Dim blue2 As Integer = (colorValue2 \ 256 \ 256) Mod 256
                                reds2(i - 1, j - 1) = red2
                                greens2(i - 1, j - 1) = green2
                                blues2(i - 1, j - 1) = blue2
                            End If
                        End If

                    Next
                Next

                rng.ClearContents()
                rng.ClearFormats()

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


                                Dim red1 As Integer = reds1(i - 1, j - 1)
                                Dim green1 As Integer = greens1(i - 1, j - 1)
                                Dim blue1 As Integer = blues1(i - 1, j - 1)

                                rng2.Cells(x, y).Interior.Color = System.Drawing.Color.FromArgb(red1, green1, blue1)

                                Dim red2 As Integer = reds2(i - 1, j - 1)
                                Dim green2 As Integer = greens2(i - 1, j - 1)
                                Dim blue2 As Integer = blues2(i - 1, j - 1)
                                rng2.Cells(x, y).Font.Color = System.Drawing.Color.FromArgb(red2, green2, blue2)
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

                                Dim red1 As Integer = reds1(i - 1, j - 1)
                                Dim green1 As Integer = greens1(i - 1, j - 1)
                                Dim blue1 As Integer = blues1(i - 1, j - 1)
                                rng2.Cells(x, y).Interior.Color = System.Drawing.Color.FromArgb(red1, green1, blue1)

                                Dim red2 As Integer = reds2(i - 1, j - 1)
                                Dim green2 As Integer = greens2(i - 1, j - 1)
                                Dim blue2 As Integer = blues2(i - 1, j - 1)
                                rng2.Cells(x, y).Font.Color = System.Drawing.Color.FromArgb(red2, green2, blue2)
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
                    excelApp.CutCopyMode = Excel.XlCutCopyMode.xlCopy
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
                    excelApp.CutCopyMode = Excel.XlCutCopyMode.xlCopy
                Else
                    MessageBox.Show("Choose One Transformation Option. ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub

                End If
            Else

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


                            If IsDBNull(font.Name) = False Then
                                fontNames(i - 1, j - 1) = font.Name
                            Else
                                fontNames(i - 1, j - 1) = "Calibri"
                            End If

                            If IsDBNull(font.Size) = False Then
                                Dim fontSize As Single = Convert.ToSingle(font.Size)
                                fontSizes(i - 1, j - 1) = fontSize
                            Else
                                fontSizes(i - 1, j - 1) = 11
                            End If

                            If IsDBNull(cell.Interior.Color) Then
                                reds1(i - 1, j - 1) = 255
                                greens1(i - 1, j - 1) = 255
                                blues1(i - 1, j - 1) = 255
                            Else
                                Dim colorValue1 As Long = CLng(cell.Interior.Color)
                                Dim red1 As Integer = colorValue1 Mod 256
                                Dim green1 As Integer = (colorValue1 \ 256) Mod 256
                                Dim blue1 As Integer = (colorValue1 \ 256 \ 256) Mod 256
                                reds1(i - 1, j - 1) = red1
                                greens1(i - 1, j - 1) = green1
                                blues1(i - 1, j - 1) = blue1
                            End If

                            If IsDBNull(cell.Font.Color) Then
                                reds2(i - 1, j - 1) = 0
                                greens2(i - 1, j - 1) = 0
                                blues2(i - 1, j - 1) = 0
                            Else
                                Dim colorValue2 As Long = CLng(cell.Font.Color)
                                Dim red2 As Integer = colorValue2 Mod 256
                                Dim green2 As Integer = (colorValue2 \ 256) Mod 256
                                Dim blue2 As Integer = (colorValue2 \ 256 \ 256) Mod 256
                                reds2(i - 1, j - 1) = red2
                                greens2(i - 1, j - 1) = green2
                                blues2(i - 1, j - 1) = blue2
                            End If
                        End If

                    Next
                Next

                rng.ClearContents()
                rng.ClearFormats()

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

                                Dim red1 As Integer = reds1(i - 1, j - 1)
                                Dim green1 As Integer = greens1(i - 1, j - 1)
                                Dim blue1 As Integer = blues1(i - 1, j - 1)
                                rng2.Cells(x, y).Interior.Color = System.Drawing.Color.FromArgb(red1, green1, blue1)

                                Dim red2 As Integer = reds2(i - 1, j - 1)
                                Dim green2 As Integer = greens2(i - 1, j - 1)
                                Dim blue2 As Integer = blues2(i - 1, j - 1)
                                rng2.Cells(x, y).Font.Color = System.Drawing.Color.FromArgb(red2, green2, blue2)
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

                                Dim red1 As Integer = reds1(i - 1, j - 1)
                                Dim green1 As Integer = greens1(i - 1, j - 1)
                                Dim blue1 As Integer = blues1(i - 1, j - 1)
                                rng2.Cells(x, y).Interior.Color = System.Drawing.Color.FromArgb(red1, green1, blue1)

                                Dim red2 As Integer = reds2(i - 1, j - 1)
                                Dim green2 As Integer = greens2(i - 1, j - 1)
                                Dim blue2 As Integer = blues2(i - 1, j - 1)
                                rng2.Cells(x, y).Font.Color = System.Drawing.Color.FromArgb(red2, green2, blue2)
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

                Dim r2 As Integer
                Dim c2 As Integer

                If X5 Then
                    r2 = UBound(BreakPoints) + 1
                    c2 = MaxValue(lengths)
                ElseIf X6 Then
                    c2 = UBound(BreakPoints) + 1
                    r2 = MaxValue(lengths)
                End If

                rng2 = worksheet2.Range(rng2.Cells(1, 1), rng2.Cells(r2, c2))
                Dim rng2Address As String = rng2.Address
                worksheet2.Activate()
                rng2.Select()

                If Overlap(excelApp, worksheet, worksheet2, rng, rng2) = False Then
                    If X5 Then
                        Dim iRow As Integer
                        iRow = 0
                        For i = 1 To r2
                            For j = 1 To c2
                                Dim x As Integer
                                Dim y As Integer
                                x = iRow + j
                                y = 1
                                If x < BreakPoints(i - 1) Then
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
                        excelApp.CutCopyMode = Excel.XlCutCopyMode.xlCopy
                    ElseIf X6 Then
                        Dim iRow As Integer
                        iRow = 0
                        For j = 1 To c2
                            For i = 1 To r2
                                Dim x As Integer
                                Dim y As Integer
                                x = iRow + i
                                y = 1

                                If x < BreakPoints(j - 1) Then
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
                        excelApp.CutCopyMode = Excel.XlCutCopyMode.xlCopy
                    End If
                Else

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


                                If IsDBNull(font.Name) = False Then
                                    fontNames(i - 1, j - 1) = font.Name
                                Else
                                    fontNames(i - 1, j - 1) = "Calibri"
                                End If

                                If IsDBNull(font.Size) = False Then
                                    Dim fontSize As Single = Convert.ToSingle(font.Size)
                                    fontSizes(i - 1, j - 1) = fontSize
                                Else
                                    fontSizes(i - 1, j - 1) = 11
                                End If

                                If IsDBNull(cell.Interior.Color) Then
                                    reds1(i - 1, j - 1) = 255
                                    greens1(i - 1, j - 1) = 255
                                    blues1(i - 1, j - 1) = 255
                                Else
                                    Dim colorValue1 As Long = CLng(cell.Interior.Color)
                                    Dim red1 As Integer = colorValue1 Mod 256
                                    Dim green1 As Integer = (colorValue1 \ 256) Mod 256
                                    Dim blue1 As Integer = (colorValue1 \ 256 \ 256) Mod 256
                                    reds1(i - 1, j - 1) = red1
                                    greens1(i - 1, j - 1) = green1
                                    blues1(i - 1, j - 1) = blue1
                                End If

                                If IsDBNull(cell.Font.Color) Then
                                    reds2(i - 1, j - 1) = 0
                                    greens2(i - 1, j - 1) = 0
                                    blues2(i - 1, j - 1) = 0
                                Else
                                    Dim colorValue2 As Long = CLng(cell.Font.Color)
                                    Dim red2 As Integer = colorValue2 Mod 256
                                    Dim green2 As Integer = (colorValue2 \ 256) Mod 256
                                    Dim blue2 As Integer = (colorValue2 \ 256 \ 256) Mod 256
                                    reds2(i - 1, j - 1) = red2
                                    greens2(i - 1, j - 1) = green2
                                    blues2(i - 1, j - 1) = blue2
                                End If
                            End If

                        Next
                    Next

                    rng.ClearContents()
                    rng.ClearFormats()

                    If X5 Then
                        Dim iRow As Integer
                        iRow = 0
                        For i = 1 To r2
                            For j = 1 To c2
                                Dim x As Integer
                                Dim y As Integer
                                x = iRow + j
                                y = 1
                                If x < BreakPoints(i - 1) And x <= r Then
                                    rng2.Cells(i, j).Value = Arr(x - 1, y - 1)

                                    If CheckBox1.Checked = True Then

                                        Dim cell2 As Excel.Range = rng2.Cells(i, j)
                                        Dim font2 As Excel.Font = cell2.Font

                                        Dim fontSize As Single = fontSizes(x - 1, y - 1)

                                        rng2.Cells(i, j).Font.Name = fontNames(x - 1, y - 1)
                                        rng2.Cells(i, j).Font.Size = fontSizes(x - 1, y - 1)

                                        If Bolds(x - 1, y - 1) Then rng2.Cells(i, j).Font.Bold = True
                                        If Italics(x - 1, y - 1) Then rng2.Cells(i, j).Font.Italic = True

                                        Dim red1 As Integer = reds1(x - 1, y - 1)
                                        Dim green1 As Integer = greens1(x - 1, y - 1)
                                        Dim blue1 As Integer = blues1(x - 1, y - 1)
                                        rng2.Cells(i, j).Interior.Color = System.Drawing.Color.FromArgb(red1, green1, blue1)

                                        Dim red2 As Integer = reds2(x - 1, y - 1)
                                        Dim green2 As Integer = greens2(x - 1, y - 1)
                                        Dim blue2 As Integer = blues2(x - 1, y - 1)
                                        rng2.Cells(i, j).Font.Color = System.Drawing.Color.FromArgb(red2, green2, blue2)
                                    End If
                                End If
                            Next
                            iRow = BreakPoints(i - 1)
                        Next

                    ElseIf X6 Then
                        Dim iRow As Integer
                        iRow = 0
                        For j = 1 To c2
                            For i = 1 To r2
                                Dim x As Integer
                                Dim y As Integer
                                x = iRow + i
                                y = 1
                                If x < BreakPoints(j - 1) Then
                                    rng2.Cells(i, j).Value = Arr(x - 1, y - 1)

                                    If CheckBox1.Checked = True Then

                                        Dim cell2 As Excel.Range = rng2.Cells(i, j)
                                        Dim font2 As Excel.Font = cell2.Font

                                        Dim fontSize As Single = fontSizes(x - 1, y - 1)

                                        rng2.Cells(i, j).Font.Name = fontNames(x - 1, y - 1)
                                        rng2.Cells(i, j).Font.Size = fontSizes(x - 1, y - 1)

                                        If Bolds(x - 1, y - 1) Then rng2.Cells(i, j).Font.Bold = True
                                        If Italics(x - 1, y - 1) Then rng2.Cells(i, j).Font.Italic = True

                                        Dim red1 As Integer = reds1(x - 1, y - 1)
                                        Dim green1 As Integer = greens1(x - 1, y - 1)
                                        Dim blue1 As Integer = blues1(x - 1, y - 1)
                                        rng2.Cells(i, j).Interior.Color = System.Drawing.Color.FromArgb(red1, green1, blue1)

                                        Dim red2 As Integer = reds2(x - 1, y - 1)
                                        Dim green2 As Integer = greens2(x - 1, y - 1)
                                        Dim blue2 As Integer = blues2(x - 1, y - 1)
                                        rng2.Cells(i, j).Font.Color = System.Drawing.Color.FromArgb(red2, green2, blue2)
                                    End If
                                End If
                            Next
                            iRow = BreakPoints(j - 1)
                        Next
                    End If
                End If

            ElseIf (X8 And TextBox2.Text <> "" And CanConvertToInt(TextBox2.Text) = True) And (X5 Or X6) Then

                If X5 Then

                    Dim r2 As Integer
                    Dim c2 As Integer

                    If r Mod Int(TextBox2.Text) = 0 Then
                        r2 = Int(r / Int(TextBox2.Text))
                    Else
                        r2 = Int(r / Int(TextBox2.Text)) + 1
                    End If
                    c2 = Int(TextBox2.Text)

                    rng2 = worksheet2.Range(rng2.Cells(1, 1), rng2.Cells(r2, c2))
                    Dim rng2Address As String = rng2.Address
                    worksheet2.Activate()
                    rng2.Select()

                    If Overlap(excelApp, worksheet, worksheet2, rng, rng2) = False Then
                        For i = 1 To r2
                            For j = 1 To c2
                                Dim x As Integer
                                Dim y As Integer
                                x = (c2 * (i - 1)) + j
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
                        excelApp.CutCopyMode = Excel.XlCutCopyMode.xlCopy
                    Else

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


                                    If IsDBNull(font.Name) = False Then
                                        fontNames(i - 1, j - 1) = font.Name
                                    Else
                                        fontNames(i - 1, j - 1) = "Calibri"
                                    End If

                                    If IsDBNull(font.Size) = False Then
                                        Dim fontSize As Single = Convert.ToSingle(font.Size)
                                        fontSizes(i - 1, j - 1) = fontSize
                                    Else
                                        fontSizes(i - 1, j - 1) = 11
                                    End If

                                    If IsDBNull(cell.Interior.Color) Then
                                        reds1(i - 1, j - 1) = 255
                                        greens1(i - 1, j - 1) = 255
                                        blues1(i - 1, j - 1) = 255
                                    Else
                                        Dim colorValue1 As Long = CLng(cell.Interior.Color)
                                        Dim red1 As Integer = colorValue1 Mod 256
                                        Dim green1 As Integer = (colorValue1 \ 256) Mod 256
                                        Dim blue1 As Integer = (colorValue1 \ 256 \ 256) Mod 256
                                        reds1(i - 1, j - 1) = red1
                                        greens1(i - 1, j - 1) = green1
                                        blues1(i - 1, j - 1) = blue1
                                    End If

                                    If IsDBNull(cell.Font.Color) Then
                                        reds2(i - 1, j - 1) = 0
                                        greens2(i - 1, j - 1) = 0
                                        blues2(i - 1, j - 1) = 0
                                    Else
                                        Dim colorValue2 As Long = CLng(cell.Font.Color)
                                        Dim red2 As Integer = colorValue2 Mod 256
                                        Dim green2 As Integer = (colorValue2 \ 256) Mod 256
                                        Dim blue2 As Integer = (colorValue2 \ 256 \ 256) Mod 256
                                        reds2(i - 1, j - 1) = red2
                                        greens2(i - 1, j - 1) = green2
                                        blues2(i - 1, j - 1) = blue2
                                    End If
                                End If

                            Next
                        Next

                        rng.ClearContents()
                        rng.ClearFormats()

                        For i = 1 To r2
                            For j = 1 To c2
                                Dim x As Integer
                                Dim y As Integer
                                x = (c2 * (i - 1)) + j
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

                                        Dim red1 As Integer = reds1(x - 1, y - 1)
                                        Dim green1 As Integer = greens1(x - 1, y - 1)
                                        Dim blue1 As Integer = blues1(x - 1, y - 1)
                                        rng2.Cells(i, j).Interior.Color = System.Drawing.Color.FromArgb(red1, green1, blue1)

                                        Dim red2 As Integer = reds2(x - 1, y - 1)
                                        Dim green2 As Integer = greens2(x - 1, y - 1)
                                        Dim blue2 As Integer = blues2(x - 1, y - 1)
                                        rng2.Cells(i, j).Font.Color = System.Drawing.Color.FromArgb(red2, green2, blue2)
                                    End If
                                End If
                            Next
                        Next
                    End If

                ElseIf X6 Then

                    Dim r2 As Integer
                    Dim c2 As Integer

                    If r Mod Int(TextBox2.Text) = 0 Then
                        c2 = Int(r / Int(TextBox2.Text))
                    Else
                        c2 = Int(r / Int(TextBox2.Text)) + 1
                    End If
                    r2 = Int(TextBox2.Text)

                    rng2 = worksheet2.Range(rng2.Cells(1, 1), rng2.Cells(r2, c2))
                    Dim rng2Address As String = rng2.Address
                    worksheet2.Activate()
                    rng2.Select()

                    If Overlap(excelApp, worksheet, worksheet2, rng, rng2) = False Then

                        For j = 1 To c2
                            For i = 1 To r2
                                Dim x As Integer
                                Dim y As Integer
                                x = (r2 * (j - 1)) + i
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
                        excelApp.CutCopyMode = Excel.XlCutCopyMode.xlCopy

                    Else

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


                                    If IsDBNull(font.Name) = False Then
                                        fontNames(i - 1, j - 1) = font.Name
                                    Else
                                        fontNames(i - 1, j - 1) = "Calibri"
                                    End If

                                    If IsDBNull(font.Size) = False Then
                                        Dim fontSize As Single = Convert.ToSingle(font.Size)
                                        fontSizes(i - 1, j - 1) = fontSize
                                    Else
                                        fontSizes(i - 1, j - 1) = 11
                                    End If

                                    If IsDBNull(cell.Interior.Color) Then
                                        reds1(i - 1, j - 1) = 255
                                        greens1(i - 1, j - 1) = 255
                                        blues1(i - 1, j - 1) = 255
                                    Else
                                        Dim colorValue1 As Long = CLng(cell.Interior.Color)
                                        Dim red1 As Integer = colorValue1 Mod 256
                                        Dim green1 As Integer = (colorValue1 \ 256) Mod 256
                                        Dim blue1 As Integer = (colorValue1 \ 256 \ 256) Mod 256
                                        reds1(i - 1, j - 1) = red1
                                        greens1(i - 1, j - 1) = green1
                                        blues1(i - 1, j - 1) = blue1
                                    End If

                                    If IsDBNull(cell.Font.Color) Then
                                        reds2(i - 1, j - 1) = 0
                                        greens2(i - 1, j - 1) = 0
                                        blues2(i - 1, j - 1) = 0
                                    Else
                                        Dim colorValue2 As Long = CLng(cell.Font.Color)
                                        Dim red2 As Integer = colorValue2 Mod 256
                                        Dim green2 As Integer = (colorValue2 \ 256) Mod 256
                                        Dim blue2 As Integer = (colorValue2 \ 256 \ 256) Mod 256
                                        reds2(i - 1, j - 1) = red2
                                        greens2(i - 1, j - 1) = green2
                                        blues2(i - 1, j - 1) = blue2
                                    End If
                                End If

                            Next
                        Next

                        rng.ClearContents()
                        rng.ClearFormats()

                        For j = 1 To c2
                            For i = 1 To r2
                                Dim x As Integer
                                Dim y As Integer
                                x = (r2 * (j - 1)) + i
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

                                        Dim red1 As Integer = reds1(x - 1, y - 1)
                                        Dim green1 As Integer = greens1(x - 1, y - 1)
                                        Dim blue1 As Integer = blues1(x - 1, y - 1)
                                        rng2.Cells(i, j).Interior.Color = System.Drawing.Color.FromArgb(red1, green1, blue1)

                                        Dim red2 As Integer = reds2(x - 1, y - 1)
                                        Dim green2 As Integer = greens2(x - 1, y - 1)
                                        Dim blue2 As Integer = blues2(x - 1, y - 1)
                                        rng2.Cells(i, j).Font.Color = System.Drawing.Color.FromArgb(red2, green2, blue2)
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

                Dim r2 As Integer
                Dim c2 As Integer

                Dim BreakPoints() As Integer
                BreakPoints = GetBreakPoints(rng, 1)

                Dim lengths() As Integer
                lengths = GetLengths(BreakPoints)

                If X5 Then
                    r2 = UBound(BreakPoints) + 1
                    c2 = MaxValue(lengths)
                ElseIf X6 Then
                    c2 = UBound(BreakPoints) + 1
                    r2 = MaxValue(lengths)
                End If

                rng2 = worksheet2.Range(rng2.Cells(1, 1), rng2.Cells(r2, c2))
                Dim rng2Address As String = rng2.Address
                worksheet2.Activate()
                rng2.Select()

                If Overlap(excelApp, worksheet, worksheet2, rng, rng2) = False Then
                    If X5 Then
                        Dim iColumn As Integer
                        iColumn = 0
                        For i = 1 To r2
                            For j = 1 To c2
                                Dim x As Integer
                                Dim y As Integer
                                x = 1
                                y = iColumn + j
                                If y < BreakPoints(i - 1) Then
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
                        excelApp.CutCopyMode = Excel.XlCutCopyMode.xlCopy
                    ElseIf X6 Then
                        Dim iColumn As Integer
                        iColumn = 0
                        For j = 1 To c2
                            For i = 1 To r2
                                Dim x As Integer
                                Dim y As Integer
                                x = 1
                                y = iColumn + i
                                If y < BreakPoints(j - 1) Then

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
                        excelApp.CutCopyMode = Excel.XlCutCopyMode.xlCopy
                    End If

                Else

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


                                If IsDBNull(font.Name) = False Then
                                    fontNames(i - 1, j - 1) = font.Name
                                Else
                                    fontNames(i - 1, j - 1) = "Calibri"
                                End If

                                If IsDBNull(font.Size) = False Then
                                    Dim fontSize As Single = Convert.ToSingle(font.Size)
                                    fontSizes(i - 1, j - 1) = fontSize
                                Else
                                    fontSizes(i - 1, j - 1) = 11
                                End If

                                If IsDBNull(cell.Interior.Color) Then
                                    reds1(i - 1, j - 1) = 255
                                    greens1(i - 1, j - 1) = 255
                                    blues1(i - 1, j - 1) = 255
                                Else
                                    Dim colorValue1 As Long = CLng(cell.Interior.Color)
                                    Dim red1 As Integer = colorValue1 Mod 256
                                    Dim green1 As Integer = (colorValue1 \ 256) Mod 256
                                    Dim blue1 As Integer = (colorValue1 \ 256 \ 256) Mod 256
                                    reds1(i - 1, j - 1) = red1
                                    greens1(i - 1, j - 1) = green1
                                    blues1(i - 1, j - 1) = blue1
                                End If

                                If IsDBNull(cell.Font.Color) Then
                                    reds2(i - 1, j - 1) = 0
                                    greens2(i - 1, j - 1) = 0
                                    blues2(i - 1, j - 1) = 0
                                Else
                                    Dim colorValue2 As Long = CLng(cell.Font.Color)
                                    Dim red2 As Integer = colorValue2 Mod 256
                                    Dim green2 As Integer = (colorValue2 \ 256) Mod 256
                                    Dim blue2 As Integer = (colorValue2 \ 256 \ 256) Mod 256
                                    reds2(i - 1, j - 1) = red2
                                    greens2(i - 1, j - 1) = green2
                                    blues2(i - 1, j - 1) = blue2
                                End If
                            End If

                        Next
                    Next

                    rng.ClearContents()
                    rng.ClearFormats()

                    If X5 Then

                        Dim iColumn As Integer
                        iColumn = 0
                        For i = 1 To r2
                            For j = 1 To c2
                                Dim x As Integer
                                Dim y As Integer
                                x = 1
                                y = iColumn + j
                                If y < BreakPoints(i - 1) Then
                                    rng2.Cells(i, j).Value = Arr(x - 1, y - 1)

                                    If CheckBox1.Checked = True Then

                                        Dim cell2 As Excel.Range = rng2.Cells(i, j)
                                        Dim font2 As Excel.Font = cell2.Font

                                        Dim fontSize As Single = fontSizes(x - 1, y - 1)

                                        rng2.Cells(i, j).Font.Name = fontNames(x - 1, y - 1)
                                        rng2.Cells(i, j).Font.Size = fontSizes(x - 1, y - 1)

                                        If Bolds(x - 1, y - 1) Then rng2.Cells(i, j).Font.Bold = True
                                        If Italics(x - 1, y - 1) Then rng2.Cells(i, j).Font.Italic = True

                                        Dim red1 As Integer = reds1(x - 1, y - 1)
                                        Dim green1 As Integer = greens1(x - 1, y - 1)
                                        Dim blue1 As Integer = blues1(x - 1, y - 1)
                                        rng2.Cells(i, j).Interior.Color = System.Drawing.Color.FromArgb(red1, green1, blue1)

                                        Dim red2 As Integer = reds2(x - 1, y - 1)
                                        Dim green2 As Integer = greens2(x - 1, y - 1)
                                        Dim blue2 As Integer = blues2(x - 1, y - 1)
                                        rng2.Cells(i, j).Font.Color = System.Drawing.Color.FromArgb(red2, green2, blue2)
                                    End If
                                End If
                            Next
                            iColumn = BreakPoints(i - 1)
                        Next

                    ElseIf X6 Then
                        Dim iColumn As Integer
                        iColumn = 0
                        For j = 1 To c2
                            For i = 1 To r2
                                Dim x As Integer
                                Dim y As Integer
                                x = 1
                                y = iColumn + i
                                If y < BreakPoints(j - 1) Then
                                    rng2.Cells(i, j).Value = Arr(x - 1, y - 1)

                                    If CheckBox1.Checked = True Then

                                        Dim cell2 As Excel.Range = rng2.Cells(i, j)
                                        Dim font2 As Excel.Font = cell2.Font

                                        Dim fontSize As Single = fontSizes(x - 1, y - 1)

                                        rng2.Cells(i, j).Font.Name = fontNames(x - 1, y - 1)
                                        rng2.Cells(i, j).Font.Size = fontSizes(x - 1, y - 1)

                                        If Bolds(x - 1, y - 1) Then rng2.Cells(i, j).Font.Bold = True
                                        If Italics(x - 1, y - 1) Then rng2.Cells(i, j).Font.Italic = True

                                        Dim red1 As Integer = reds1(x - 1, y - 1)
                                        Dim green1 As Integer = greens1(x - 1, y - 1)
                                        Dim blue1 As Integer = blues1(x - 1, y - 1)
                                        rng2.Cells(i, j).Interior.Color = System.Drawing.Color.FromArgb(red1, green1, blue1)

                                        Dim red2 As Integer = reds2(x - 1, y - 1)
                                        Dim green2 As Integer = greens2(x - 1, y - 1)
                                        Dim blue2 As Integer = blues2(x - 1, y - 1)
                                        rng2.Cells(i, j).Font.Color = System.Drawing.Color.FromArgb(red2, green2, blue2)

                                    End If
                                End If
                            Next
                            iColumn = BreakPoints(j - 1)
                        Next
                    End If
                End If

            ElseIf (X8 And TextBox2.Text <> "" And CanConvertToInt(TextBox2.Text) = True) And (X5 Or X6) Then

                If X5 Then
                    Dim r2 As Integer
                    Dim c2 As Integer
                    If c Mod Int(TextBox2.Text) = 0 Then
                        r2 = Int(c / Int(TextBox2.Text))
                    Else
                        r2 = Int(c / Int(TextBox2.Text)) + 1
                    End If
                    c2 = Int(TextBox2.Text)

                    rng2 = worksheet2.Range(rng2.Cells(1, 1), rng2.Cells(r2, c2))
                    Dim rng2Address As String = rng2.Address
                    worksheet2.Activate()
                    rng2.Select()

                    If Overlap(excelApp, worksheet, worksheet2, rng, rng2) = False Then
                        For i = 1 To r2
                            For j = 1 To c2
                                Dim x As Integer
                                Dim y As Integer
                                x = 1
                                y = (c2 * (i - 1)) + j

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
                        excelApp.CutCopyMode = Excel.XlCutCopyMode.xlCopy
                    Else

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


                                    If IsDBNull(font.Name) = False Then
                                        fontNames(i - 1, j - 1) = font.Name
                                    Else
                                        fontNames(i - 1, j - 1) = "Calibri"
                                    End If

                                    If IsDBNull(font.Size) = False Then
                                        Dim fontSize As Single = Convert.ToSingle(font.Size)
                                        fontSizes(i - 1, j - 1) = fontSize
                                    Else
                                        fontSizes(i - 1, j - 1) = 11
                                    End If

                                    If IsDBNull(cell.Interior.Color) Then
                                        reds1(i - 1, j - 1) = 255
                                        greens1(i - 1, j - 1) = 255
                                        blues1(i - 1, j - 1) = 255
                                    Else
                                        Dim colorValue1 As Long = CLng(cell.Interior.Color)
                                        Dim red1 As Integer = colorValue1 Mod 256
                                        Dim green1 As Integer = (colorValue1 \ 256) Mod 256
                                        Dim blue1 As Integer = (colorValue1 \ 256 \ 256) Mod 256
                                        reds1(i - 1, j - 1) = red1
                                        greens1(i - 1, j - 1) = green1
                                        blues1(i - 1, j - 1) = blue1
                                    End If

                                    If IsDBNull(cell.Font.Color) Then
                                        reds2(i - 1, j - 1) = 0
                                        greens2(i - 1, j - 1) = 0
                                        blues2(i - 1, j - 1) = 0
                                    Else
                                        Dim colorValue2 As Long = CLng(cell.Font.Color)
                                        Dim red2 As Integer = colorValue2 Mod 256
                                        Dim green2 As Integer = (colorValue2 \ 256) Mod 256
                                        Dim blue2 As Integer = (colorValue2 \ 256 \ 256) Mod 256
                                        reds2(i - 1, j - 1) = red2
                                        greens2(i - 1, j - 1) = green2
                                        blues2(i - 1, j - 1) = blue2
                                    End If
                                End If

                            Next
                        Next

                        rng.ClearContents()
                        rng.ClearFormats()

                        For i = 1 To r2
                            For j = 1 To c2
                                Dim x As Integer
                                Dim y As Integer
                                x = 1
                                y = (c2 * (i - 1)) + j
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

                                        Dim red1 As Integer = reds1(x - 1, y - 1)
                                        Dim green1 As Integer = greens1(x - 1, y - 1)
                                        Dim blue1 As Integer = blues1(x - 1, y - 1)
                                        rng2.Cells(i, j).Interior.Color = System.Drawing.Color.FromArgb(red1, green1, blue1)

                                        Dim red2 As Integer = reds2(x - 1, y - 1)
                                        Dim green2 As Integer = greens2(x - 1, y - 1)
                                        Dim blue2 As Integer = blues2(x - 1, y - 1)
                                        rng2.Cells(i, j).Font.Color = System.Drawing.Color.FromArgb(red2, green2, blue2)
                                    End If
                                End If
                            Next
                        Next
                    End If

                ElseIf X6 Then
                    Dim r2 As Integer
                    Dim c2 As Integer
                    If c Mod Int(TextBox2.Text) = 0 Then
                        c2 = Int(c / Int(TextBox2.Text))
                    Else
                        c2 = Int(c / Int(TextBox2.Text)) + 1
                    End If
                    r2 = Int(TextBox2.Text)

                    rng2 = worksheet2.Range(rng2.Cells(1, 1), rng2.Cells(r2, c2))
                    Dim rng2Address As String = rng2.Address
                    worksheet2.Activate()
                    rng2.Select()

                    If Overlap(excelApp, worksheet, worksheet2, rng, rng2) = False Then

                        For j = 1 To c2
                            For i = 1 To r2
                                Dim x As Integer
                                Dim y As Integer
                                x = 1
                                y = (r2 * (j - 1)) + i
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
                        excelApp.CutCopyMode = Excel.XlCutCopyMode.xlCopy
                    Else

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


                                    If IsDBNull(font.Name) = False Then
                                        fontNames(i - 1, j - 1) = font.Name
                                    Else
                                        fontNames(i - 1, j - 1) = "Calibri"
                                    End If

                                    If IsDBNull(font.Size) = False Then
                                        Dim fontSize As Single = Convert.ToSingle(font.Size)
                                        fontSizes(i - 1, j - 1) = fontSize
                                    Else
                                        fontSizes(i - 1, j - 1) = 11
                                    End If

                                    If IsDBNull(cell.Interior.Color) Then
                                        reds1(i - 1, j - 1) = 255
                                        greens1(i - 1, j - 1) = 255
                                        blues1(i - 1, j - 1) = 255
                                    Else
                                        Dim colorValue1 As Long = CLng(cell.Interior.Color)
                                        Dim red1 As Integer = colorValue1 Mod 256
                                        Dim green1 As Integer = (colorValue1 \ 256) Mod 256
                                        Dim blue1 As Integer = (colorValue1 \ 256 \ 256) Mod 256
                                        reds1(i - 1, j - 1) = red1
                                        greens1(i - 1, j - 1) = green1
                                        blues1(i - 1, j - 1) = blue1
                                    End If

                                    If IsDBNull(cell.Font.Color) Then
                                        reds2(i - 1, j - 1) = 0
                                        greens2(i - 1, j - 1) = 0
                                        blues2(i - 1, j - 1) = 0
                                    Else
                                        Dim colorValue2 As Long = CLng(cell.Font.Color)
                                        Dim red2 As Integer = colorValue2 Mod 256
                                        Dim green2 As Integer = (colorValue2 \ 256) Mod 256
                                        Dim blue2 As Integer = (colorValue2 \ 256 \ 256) Mod 256
                                        reds2(i - 1, j - 1) = red2
                                        greens2(i - 1, j - 1) = green2
                                        blues2(i - 1, j - 1) = blue2
                                    End If
                                End If

                            Next
                        Next

                        rng.ClearContents()
                        rng.ClearFormats()

                        For j = 1 To c2
                            For i = 1 To r2
                                Dim x As Integer
                                Dim y As Integer
                                x = 1
                                y = (r2 * (j - 1)) + i
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

                                        Dim red1 As Integer = reds1(x - 1, y - 1)
                                        Dim green1 As Integer = greens1(x - 1, y - 1)
                                        Dim blue1 As Integer = blues1(x - 1, y - 1)
                                        rng2.Cells(i, j).Interior.Color = System.Drawing.Color.FromArgb(red1, green1, blue1)

                                        Dim red2 As Integer = reds2(x - 1, y - 1)
                                        Dim green2 As Integer = greens2(x - 1, y - 1)
                                        Dim blue2 As Integer = blues2(x - 1, y - 1)
                                        rng2.Cells(i, j).Font.Color = System.Drawing.Color.FromArgb(red2, green2, blue2)
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

        For j = 1 To rng2.Columns.Count
            rng2.Columns(j).Autofit
        Next

        Me.Close()

        TextBoxChanged = False

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

        Try
            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            worksheet = workbook.ActiveSheet

            Dim rngArray() As String = Split(TextBox1.Text, "!")
            Dim rngAddress As String = rngArray(UBound(rngArray))
            rng = worksheet.Range(rngAddress)
            TextBoxChanged = True
            rng.Select()
            Call Display()
            Call Setup()
            TextBoxChanged = False
        Catch ex As Exception

        End Try

    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged

        Try
            If RadioButton1.Checked = True Then
                Call Display()
                Call Setup()
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged

        Try
            If RadioButton3.Checked = True Then
                Call Display()
                Call Setup()
            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged

        Try
            If RadioButton2.Checked = True Then
                Call Display()
                Call Setup()
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub RadioButton4_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton4.CheckedChanged

        Try
            If RadioButton4.Checked = True Then
                Call Display()
                Call Setup()
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged

        Try
            Call Display()
            Call Setup()
        Catch ex As Exception

        End Try

    End Sub

    Private Sub RadioButton5_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton5.CheckedChanged

        Try
            If RadioButton5.Checked = True Then
                Call Display()
                Call Setup()
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub RadioButton6_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton6.CheckedChanged

        Try
            If RadioButton6.Checked = True Then
                Call Display()
                Call Setup()
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub RadioButton7_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton7.CheckedChanged

        Try
            If RadioButton7.Checked = True Then
                Call Display()
                Call Setup()
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub RadioButton8_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton8.CheckedChanged

        Try
            If RadioButton8.Checked = True Then
                Call Display()
                Call Setup()
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

        Try
            Call Display()
            Call Setup()
        Catch ex As Exception

        End Try

    End Sub

    Private Sub PictureBox4_Click(sender As Object, e As EventArgs) Handles PictureBox4.Click

        Try
            FocusedTextBox = 1

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
            rng = worksheet.Range(worksheet.Cells(startRow, startColumn), worksheet.Cells(endRow, endColumn))

            rng.Select()

            Dim sheetName As String

            sheetName = Split(rng.Address(True, True, Excel.XlReferenceStyle.xlA1, True), "]")(1)
            sheetName = Split(sheetName, "!")(0)

            If Mid(sheetName, Len(sheetName), 1) = "'" Then
                sheetName = Mid(sheetName, 1, Len(sheetName) - 1)
            End If

            worksheet = workbook.Worksheets(sheetName)
            worksheet.Activate()

            If worksheet.Name <> OpenSheet.Name Then
                TextBox1.Text = worksheet.Name & "!" & rng.Address
            Else
                TextBox1.Text = rng.Address
            End If

            Me.TextBox1.Focus()

        Catch ex As Exception

            Me.Show()
            TextBox1.Focus()

        End Try

    End Sub

    Private Sub PictureBox8_Click(sender As Object, e As EventArgs) Handles PictureBox8.Click

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

                worksheet = workbook.Worksheets(sheetName)
                worksheet.Activate()
            Catch ex As Exception

            End Try

            rng.Select()

            If worksheet.Name <> OpenSheet.Name Then
                TextBox1.Text = worksheet.Name & "!" & rng.Address
            Else
                TextBox1.Text = rng.Address
            End If

            TextBox1.Focus()

        Catch ex As Exception

        End Try

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Me.Close()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub PictureBox6_Click(sender As Object, e As EventArgs) Handles PictureBox6.Click

        Try
            FocusedTextBox = 3
            Me.Hide()

            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook

            Dim userInput As Excel.Range = excelApp.InputBox("Select a range", Type:=8)
            rng2 = userInput


            Dim sheetName As String
            sheetName = Split(rng2.Address(True, True, Excel.XlReferenceStyle.xlA1, True), "]")(1)
            sheetName = Split(sheetName, "!")(0)

            If Mid(sheetName, Len(sheetName), 1) = "'" Then
                sheetName = Mid(sheetName, 1, Len(sheetName) - 1)
            End If

            worksheet2 = workbook.Worksheets(sheetName)
            worksheet2.Activate()

            rng2.Select()

            TextBox3.Text = rng2.Address

            Me.Show()
            TextBox3.Focus()

        Catch ex As Exception

            Me.Show()
            TextBox3.Focus()

        End Try

    End Sub

    Private Sub RadioButton10_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton10.CheckedChanged

        Try
            If RadioButton10.Checked = True Then
                TextBox3.Enabled = True
                TextBox3.Focus()
            Else
                TextBox3.Clear()
                TextBox3.Enabled = False
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

    Private Sub TextBox3_GotFocus(sender As Object, e As EventArgs) Handles TextBox3.GotFocus

        Try
            FocusedTextBox = 3
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

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged

        Try
            excelApp = Globals.ThisAddIn.Application
            workbook2 = excelApp.ActiveWorkbook
            worksheet2 = workbook2.ActiveSheet

            Dim rng2Array() As String = Split(TextBox3.Text, "!")
            Dim rng2Address As String = rng2Array(UBound(rng2Array))
            rng2 = worksheet2.Range(rng2Address)

            TextBoxChanged = True

            rng2.Select()

            TextBoxChanged = False

        Catch ex As Exception

        End Try

    End Sub

    Private Sub RadioButton9_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton9.CheckedChanged

        Try
            If RadioButton9.Checked = True Then
                worksheet2 = worksheet
                rng2 = rng
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub Button1_GotFocus(sender As Object, e As EventArgs) Handles Button1.GotFocus

        Try
            FocusedTextBox = 0

        Catch ex As Exception

        End Try

    End Sub

    Private Sub Button2_GotFocus(sender As Object, e As EventArgs) Handles Button2.GotFocus

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

    Private Sub CheckBox2_GotFocus(sender As Object, e As EventArgs) Handles CheckBox2.GotFocus

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

    Private Sub CustomGroupBox10_GotFocus(sender As Object, e As EventArgs) Handles CustomGroupBox10.GotFocus

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

    Private Sub CustomGroupBox3_GotFocus(sender As Object, e As EventArgs) Handles CustomGroupBox3.GotFocus

        Try
            FocusedTextBox = 0

        Catch ex As Exception

        End Try

    End Sub

    Private Sub CustomGroupBox4_GotFocus(sender As Object, e As EventArgs) Handles CustomGroupBox4.GotFocus

        Try
            FocusedTextBox = 0

        Catch ex As Exception

        End Try

    End Sub

    Private Sub CustomGroupBox5_GotFocus(sender As Object, e As EventArgs) Handles CustomGroupBox5.GotFocus

        Try
            FocusedTextBox = 0

        Catch ex As Exception

        End Try

    End Sub

    Private Sub CustomGroupBox6_GotFocus(sender As Object, e As EventArgs) Handles CustomGroupBox6.GotFocus

        Try
            FocusedTextBox = 0

        Catch ex As Exception

        End Try

    End Sub

    Private Sub CustomGroupBox7_GotFocus(sender As Object, e As EventArgs) Handles CustomGroupBox7.GotFocus

        Try
            FocusedTextBox = 0

        Catch ex As Exception

        End Try

    End Sub

    Private Sub CustomGroupBox8_GotFocus(sender As Object, e As EventArgs) Handles CustomGroupBox8.GotFocus

        Try
            FocusedTextBox = 0

        Catch ex As Exception

        End Try

    End Sub

    Private Sub CustomGroupBox9_GotFocus(sender As Object, e As EventArgs) Handles CustomGroupBox9.GotFocus

        Try
            FocusedTextBox = 0

        Catch ex As Exception

        End Try

    End Sub

    Private Sub CustomPanel1_GotFocus(sender As Object, e As EventArgs) Handles CustomPanel1.GotFocus

        Try
            FocusedTextBox = 0

        Catch ex As Exception

        End Try

    End Sub

    Private Sub CustomPanel2_GotFocus(sender As Object, e As EventArgs) Handles CustomPanel2.GotFocus

        Try
            FocusedTextBox = 0

        Catch ex As Exception

        End Try

    End Sub

    Private Sub Label1_GotFocus(sender As Object, e As EventArgs) Handles Label1.GotFocus

        Try
            FocusedTextBox = 0

        Catch ex As Exception

        End Try

    End Sub

    Private Sub Label2_GotFocus(sender As Object, e As EventArgs) Handles Label2.GotFocus

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

    Private Sub PictureBox2_GotFocus(sender As Object, e As EventArgs) Handles PictureBox2.GotFocus

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

    Private Sub PictureBox4_GotFocus(sender As Object, e As EventArgs) Handles PictureBox4.GotFocus

        Try
            FocusedTextBox = 1

        Catch ex As Exception

        End Try

    End Sub

    Private Sub PictureBox5_GotFocus(sender As Object, e As EventArgs) Handles PictureBox5.GotFocus

        Try
            FocusedTextBox = 0

        Catch ex As Exception

        End Try

    End Sub

    Private Sub PictureBox6_GotFocus(sender As Object, e As EventArgs) Handles PictureBox6.GotFocus

        Try
            FocusedTextBox = 3

        Catch ex As Exception

        End Try

    End Sub

    Private Sub PictureBox7_GotFocus(sender As Object, e As EventArgs)

        Try
            FocusedTextBox = 0

        Catch ex As Exception

        End Try

    End Sub

    Private Sub PictureBox8_GotFocus(sender As Object, e As EventArgs) Handles PictureBox8.GotFocus

        Try
            FocusedTextBox = 1

        Catch ex As Exception

        End Try

    End Sub

    Private Sub RadioButton1_GotFocus(sender As Object, e As EventArgs) Handles RadioButton1.GotFocus

        Try
            FocusedTextBox = 0

        Catch ex As Exception

        End Try

    End Sub

    Private Sub RadioButton10_GotFocus(sender As Object, e As EventArgs) Handles RadioButton10.GotFocus

        Try
            FocusedTextBox = 0

        Catch ex As Exception

        End Try

    End Sub

    Private Sub RadioButton2_GotFocus(sender As Object, e As EventArgs) Handles RadioButton2.GotFocus

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

    Private Sub RadioButton6_GotFocus(sender As Object, e As EventArgs) Handles RadioButton6.GotFocus

        Try
            FocusedTextBox = 0

        Catch ex As Exception

        End Try

    End Sub

    Private Sub RadioButton7_GotFocus(sender As Object, e As EventArgs) Handles RadioButton7.GotFocus

        Try
            FocusedTextBox = 0

        Catch ex As Exception

        End Try

    End Sub

    Private Sub RadioButton8_GotFocus(sender As Object, e As EventArgs) Handles RadioButton8.GotFocus

        Try
            FocusedTextBox = 0

        Catch ex As Exception

        End Try

    End Sub

    Private Sub RadioButton9_GotFocus(sender As Object, e As EventArgs) Handles RadioButton9.GotFocus

        Try
            FocusedTextBox = 0

        Catch ex As Exception

        End Try

    End Sub

    Private Sub TextBox2_GotFocus(sender As Object, e As EventArgs) Handles TextBox2.GotFocus

        Try
            FocusedTextBox = 0

        Catch ex As Exception

        End Try

    End Sub

    Private Sub VScrollBar1_GotFocus(sender As Object, e As EventArgs) Handles VScrollBar1.GotFocus

        Try
            FocusedTextBox = 0

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

    Private Sub Form7_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        Try
            form_flag = False

        Catch ex As Exception

        End Try

    End Sub

    Private Sub Form7_Shown(sender As Object, e As EventArgs) Handles Me.Shown

        Try
            Me.Focus()
            Me.BringToFront()
            Me.Activate()
            Me.BeginInvoke(New System.Action(Sub()
                                                 TextBox1.Text = rng.Address
                                                 SetWindowPos(Me.Handle, New IntPtr(HWND_TOPMOST), 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOSIZE)
                                             End Sub))

        Catch ex As Exception

        End Try

    End Sub

    Private Sub Form7_Disposed(sender As Object, e As EventArgs) Handles Me.Disposed

        Try
            form_flag = False

        Catch ex As Exception

        End Try

    End Sub

    Private Sub Form7_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown

        Try
            If e.KeyCode = Keys.Enter Then

                Button2.Focus()
                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try

    End Sub

End Class