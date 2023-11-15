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

Public Class Form18_CombineRanges

    Dim WithEvents excelApp As Excel.Application
    Dim workBook As Excel.Workbook
    Dim workSheet As Excel.Worksheet
    Dim rng As Excel.Range
    Dim selectedRange As Excel.Range

    Dim opened As Integer
    Dim FocusedTextBox As Integer
    Dim TextBoxChanged As Boolean


    Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As IntPtr, ByVal hWndInsertAfter As IntPtr, ByVal X As Integer, ByVal Y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal uFlags As UInteger) As Boolean
    Private Const SWP_NOMOVE As UInteger = &H2
    Private Const SWP_NOSIZE As UInteger = &H1
    Private Const SWP_NOACTIVATE As UInteger = &H10
    Private Const HWND_TOPMOST As Integer = -1

    Private Function IsValidExcelCellReference(cellReference As String) As Boolean

        'Checks whether a string is a valid cell reference or not.

        Dim cellPattern As String = "(\$?[A-Z]+\$?[0-9]+)"

        Dim referencePattern As String = "^" + cellPattern + "(:" + cellPattern + ")?$"

        Dim regex As New Regex(referencePattern)

        If regex.IsMatch(cellReference) Then
            Return True
        Else
            Return False
        End If

    End Function
    Private Function Operation(Arr, Flag)

        'Takes an array of numbers and conduct mathematical operations. The operation name is input as flag in the format "=...".

        If Flag = "=SUM()" Then
            Dim Output As Double = 0
            For i = LBound(Arr) To UBound(Arr)
                If IsNumeric(Arr(i)) = True Then
                    Output = Output + Arr(i)
                End If
            Next
            Operation = Output

        ElseIf Flag = "=COUNT()" Then
            Dim Output As Integer = 0
            For i = LBound(Arr) To UBound(Arr)
                Output = Output + 1
            Next
            Operation = Output

        ElseIf Flag = "=COUNTA()" Then
            Dim Output As Integer = 0
            For i = LBound(Arr) To UBound(Arr)
                If Arr(i) IsNot Nothing Then
                    Output = Output + 1
                End If
            Next
            Operation = Output

        ElseIf Flag = "=AVERAGE()" Then
            Dim Output As Double = 0
            For i = LBound(Arr) To UBound(Arr)
                If IsNumeric(Arr(i)) = True Then
                    Output = Output + Arr(i)
                End If
            Next
            Output = Output / (UBound(Arr) + 1)
            Operation = Output

        ElseIf Flag = "=MAX()" Then
            Dim Output As Object
            Dim i As Integer = LBound(Arr)
            While IsNumeric(Arr(i)) = False And i <= UBound(Arr) - 1
                i = i + 1
            End While
            Output = Arr(i)
            For i = LBound(Arr) To UBound(Arr)
                If IsNumeric(Arr(i)) = True Then
                    If Arr(i) > Output Then
                        Output = Arr(i)
                    End If
                End If
            Next
            Operation = Output

        ElseIf Flag = "=MIN()" Then
            Dim Output As Object
            Dim i As Integer = LBound(Arr)
            While IsNumeric(Arr(i)) = False And i <= UBound(Arr) - 1
                i = i + 1
            End While

            Output = Arr(i)
            For i = LBound(Arr) To UBound(Arr)
                If IsNumeric(Arr(i)) = True Then
                    If Arr(i) < Output Then
                        Output = Arr(i)
                    End If
                End If
            Next
            Operation = Output

        ElseIf Flag = "=PRODUCT()" Then
            Dim Output As Double = 1
            Dim count As Integer = 0
            For i = LBound(Arr) To UBound(Arr)
                If IsNumeric(Arr(i)) = True Then
                    Output = Output * Arr(i)
                    count = count + 1
                End If
            Next
            If count = 0 Then
                Operation = 0
            Else
                Operation = Output
            End If
        Else
            Operation = 0
        End If

    End Function

    Private Sub Display()

        Try

            CustomPanel1.Controls.Clear()
            CustomPanel2.Controls.Clear()

            Dim displayRng As Excel.Range

            'Takes the first 50 rows of the input to display.
            If rng.Rows.Count > 50 Then
                displayRng = rng.Rows("1:50")
            Else
                displayRng = rng
            End If


            Dim height As Double
            Dim width As Double

            'Default number of rows in the display box is 4.
            If displayRng.Rows.Count <= 4 Then
                height = CustomPanel1.Height / displayRng.Rows.Count
            Else
                height = (119 / 4)
            End If

            'Default number of columns in the display box is 4.
            If displayRng.Columns.Count <= 3 Then
                width = CustomPanel1.Width / displayRng.Columns.Count
            Else
                width = (260 / 3)
            End If

            'Copies the input range to the display box.
            For i = 1 To displayRng.Rows.Count
                For j = 1 To displayRng.Columns.Count

                    Dim label As New System.Windows.Forms.Label
                    label.Text = displayRng.Cells(i, j).Value
                    label.Location = New System.Drawing.Point((j - 1) * width, (i - 1) * height)
                    label.Height = height
                    label.Width = width
                    label.BorderStyle = BorderStyle.FixedSingle
                    label.TextAlign = ContentAlignment.MiddleCenter

                    'Copies the format of the input range.
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
            Next

            CustomPanel1.AutoScroll = True

            Dim X4 As Boolean

            If RadioButton4.Checked Then
                X4 = True
            Else
                X4 = ComboBox3.SelectedIndex <> -1
            End If

            If (RadioButton1.Checked Or RadioButton2.Checked Or RadioButton3.Checked) And (RadioButton4.Checked Or RadioButton5.Checked Or RadioButton6.Checked) And X4 Then

                'Works for Merging Into Single Column.
                If RadioButton1.Checked Then

                    Dim newWidth As Double
                    Dim newHeight As Double
                    Dim combinedColumn As Integer

                    If RadioButton6.Checked Or RadioButton5.Checked Then
                        If ComboBox3.SelectedItem = "Into Left Column" Then
                            combinedColumn = 1
                        ElseIf ComboBox3.SelectedItem = "Into Right Column" Then
                            combinedColumn = displayRng.Columns.Count
                        End If

                        For i = 1 To displayRng.Rows.Count

                            Dim label As New System.Windows.Forms.Label
                            If ComboBox2.SelectedIndex <= 3 Then
                                'Finds the combined string.
                                Dim combinedValue As String = ""
                                Dim Separator As String
                                Dim HFactor As Integer
                                Dim WFactor As Integer

                                If ComboBox2.SelectedIndex = 3 Then
                                    Separator = vbNewLine
                                    HFactor = displayRng.Columns.Count / 1.75
                                    WFactor = 1
                                Else
                                    If CheckBox4.Checked Then
                                        Separator = ComboBox2.Text & vbNewLine
                                        HFactor = displayRng.Columns.Count / 1.75
                                        WFactor = 1
                                    Else
                                        Separator = ComboBox2.Text
                                        HFactor = 1
                                        WFactor = displayRng.Columns.Count
                                    End If
                                End If

                                For j = 1 To displayRng.Columns.Count - 1
                                    If CheckBox3.Checked Then
                                        If displayRng.Cells(i, j).value IsNot Nothing Then
                                            combinedValue = combinedValue & displayRng.Cells(i, j).Value & Separator
                                        End If
                                    Else
                                        combinedValue = combinedValue & displayRng.Cells(i, j).Value & Separator
                                    End If
                                Next

                                If CheckBox3.Checked Then
                                    If displayRng.Cells(i, displayRng.Columns.Count).value IsNot Nothing Then
                                        combinedValue = combinedValue & displayRng.Cells(i, displayRng.Columns.Count).Value
                                    Else
                                        If Len(combinedValue) >= Len(Separator) Then
                                            combinedValue = Mid(combinedValue, 1, Len(combinedValue) - Len(Separator))
                                        End If
                                    End If
                                Else
                                    combinedValue = combinedValue & displayRng.Cells(i, displayRng.Columns.Count).Value
                                End If
                                newWidth = width * WFactor
                                newHeight = height * HFactor

                                label.Text = combinedValue
                            Else
                                'Finds the mathematical operated value (sum, max, min, count, etc...)
                                Dim OperatedValue As Double
                                Dim Values(0) As Double
                                Dim Index As Integer = -1
                                For j = 1 To displayRng.Columns.Count
                                    If IsNumeric(displayRng.Cells(i, j).Value) Then
                                        If CheckBox3.Checked Then
                                            If displayRng.Cells(i, j).value IsNot Nothing Then
                                                Index = Index + 1
                                                ReDim Preserve Values(Index)
                                                Values(Index) = displayRng.Cells(i, j).value
                                            End If
                                        Else
                                            Index = Index + 1
                                            ReDim Preserve Values(Index)
                                            Values(Index) = displayRng.Cells(i, j).value
                                        End If
                                    End If
                                Next
                                OperatedValue = Operation(Values, ComboBox2.SelectedItem)
                                label.Text = OperatedValue
                                newWidth = width
                                newHeight = height
                            End If

                            'Puts the output value in the display box.
                            label.Location = New System.Drawing.Point((combinedColumn - 1) * width, (i - 1) * newHeight)
                            label.Height = newHeight
                            label.Width = newWidth
                            label.BorderStyle = BorderStyle.FixedSingle
                            label.TextAlign = ContentAlignment.MiddleCenter

                            'Copies the format of the output cell.
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

                            CustomPanel2.Controls.Add(label)
                        Next

                        'Copies the rest columns other than the merged column.
                        If ComboBox3.SelectedItem = "Into Left Column" Then

                            For i = 1 To displayRng.Rows.Count
                                For j = 2 To displayRng.Columns.Count
                                    Dim label As New System.Windows.Forms.Label
                                    If RadioButton6.Checked Then
                                        label.Text = displayRng.Cells(i, j).value
                                    ElseIf RadioButton5.Checked Then
                                        label.Text = ""
                                    End If
                                    label.Location = New System.Drawing.Point(newWidth + (j - 2) * width, (i - 1) * newHeight)
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

                        ElseIf ComboBox3.SelectedItem = "Into Right Column" Then
                            For i = 1 To displayRng.Rows.Count
                                For j = 1 To displayRng.Columns.Count - 1
                                    Dim label As New System.Windows.Forms.Label
                                    If RadioButton6.Checked Then
                                        label.Text = displayRng.Cells(i, j).value
                                    ElseIf RadioButton5.Checked Then
                                        label.Text = ""
                                    End If
                                    label.Location = New System.Drawing.Point((j - 1) * width, (i - 1) * newHeight)
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

                        'Works to Merge the Whole Row.       
                    ElseIf RadioButton4.Checked Then
                        For i = 1 To displayRng.Rows.Count
                            Dim label As New System.Windows.Forms.Label
                            Dim HFactor As Integer
                            Dim WFactor As Integer
                            If ComboBox2.SelectedIndex <= 3 Then
                                Dim combinedValue As String = ""
                                Dim Separator As String

                                If ComboBox2.SelectedIndex = 3 Then
                                    Separator = vbNewLine
                                    HFactor = displayRng.Columns.Count / 1.75
                                    WFactor = 1
                                Else
                                    If CheckBox4.Checked Then
                                        Separator = ComboBox2.Text & vbNewLine
                                        HFactor = displayRng.Columns.Count / 1.75
                                        WFactor = 1
                                    Else
                                        Separator = ComboBox2.Text
                                        HFactor = 1
                                        WFactor = displayRng.Columns.Count
                                    End If
                                End If

                                For j = 1 To displayRng.Columns.Count - 1
                                    If CheckBox3.Checked Then
                                        If displayRng.Cells(i, j).value IsNot Nothing Then
                                            combinedValue = combinedValue & displayRng.Cells(i, j).Value & Separator
                                        End If
                                    Else
                                        combinedValue = combinedValue & displayRng.Cells(i, j).Value & Separator
                                    End If
                                Next

                                If CheckBox3.Checked Then
                                    If displayRng.Cells(i, displayRng.Columns.Count).value IsNot Nothing Then
                                        combinedValue = combinedValue & displayRng.Cells(i, displayRng.Columns.Count).Value
                                    Else
                                        If Len(combinedValue) >= Len(Separator) Then
                                            combinedValue = Mid(combinedValue, 1, Len(combinedValue) - Len(Separator))
                                        End If
                                    End If
                                Else
                                    combinedValue = combinedValue & displayRng.Cells(i, displayRng.Columns.Count).Value
                                End If
                                newWidth = width * WFactor
                                newHeight = height * HFactor

                                label.Text = combinedValue
                            Else
                                Dim OperatedValue As Double
                                Dim Values(0) As Double
                                Dim Index As Integer = -1
                                For j = 1 To displayRng.Columns.Count
                                    If IsNumeric(displayRng.Cells(i, j).Value) Then
                                        If CheckBox3.Checked Then
                                            If displayRng.Cells(i, j).value IsNot Nothing Then
                                                Index = Index + 1
                                                ReDim Preserve Values(Index)
                                                Values(Index) = displayRng.Cells(i, j).value
                                            End If
                                        Else
                                            Index = Index + 1
                                            ReDim Preserve Values(Index)
                                            Values(Index) = displayRng.Cells(i, j).value
                                        End If
                                    End If
                                Next
                                OperatedValue = Operation(Values, ComboBox2.SelectedItem)
                                newHeight = height
                                newWidth = width

                                label.Text = OperatedValue
                            End If
                            label.Location = New System.Drawing.Point(0, (i - 1) * newHeight)
                            label.Height = newHeight
                            If WFactor <> 1 Then
                                label.Width = newWidth
                            Else
                                label.Width = CustomPanel2.Width
                            End If
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

                            CustomPanel2.Controls.Add(label)
                        Next

                    End If

                    CustomPanel2.AutoScroll = True

                ElseIf RadioButton2.Checked Then

                    Dim newHeight As Double
                    Dim newWidth As Double
                    Dim combinedRow As Integer

                    If RadioButton6.Checked Or RadioButton5.Checked Then
                        If ComboBox3.SelectedItem = "Into Top Row" Then
                            combinedRow = 1
                        ElseIf ComboBox3.SelectedItem = "Into Bottom Row" Then
                            combinedRow = displayRng.Rows.Count
                        End If
                        For j = 1 To displayRng.Columns.Count
                            Dim label As New System.Windows.Forms.Label
                            If ComboBox2.SelectedIndex <= 3 Then
                                Dim combinedValue As String = ""
                                Dim Separator As String
                                Dim HFactor As Integer
                                Dim WFactor As Integer

                                If ComboBox2.SelectedIndex = 3 Then
                                    Separator = vbNewLine
                                    HFactor = displayRng.Rows.Count / 1.75
                                    WFactor = 1
                                Else
                                    If CheckBox4.Checked Then
                                        Separator = ComboBox2.Text & vbNewLine
                                        HFactor = displayRng.Rows.Count / 1.75
                                        WFactor = 1
                                    Else
                                        Separator = ComboBox2.Text
                                        HFactor = 1
                                        WFactor = displayRng.Rows.Count
                                    End If
                                End If

                                For i = 1 To displayRng.Rows.Count - 1
                                    If CheckBox3.Checked Then
                                        If displayRng.Cells(i, j).value IsNot Nothing Then
                                            combinedValue = combinedValue & displayRng.Cells(i, j).Value & Separator
                                        End If
                                    Else
                                        combinedValue = combinedValue & displayRng.Cells(i, j).Value & Separator
                                    End If
                                Next

                                If CheckBox3.Checked Then
                                    If displayRng.Cells(displayRng.Rows.Count, j).value IsNot Nothing Then
                                        combinedValue = combinedValue & displayRng.Cells(displayRng.Rows.Count, j).Value
                                    Else
                                        If Len(combinedValue) >= Len(Separator) Then
                                            combinedValue = Mid(combinedValue, 1, Len(combinedValue) - Len(Separator))
                                        End If
                                    End If
                                Else
                                    combinedValue = combinedValue & displayRng.Cells(displayRng.Rows.Count, j).Value
                                End If
                                newWidth = width * WFactor
                                newHeight = height * HFactor

                                label.Text = combinedValue

                            Else
                                Dim OperatedValue As Double
                                Dim Values(0) As Double
                                Dim Index As Integer = -1
                                For i = 1 To displayRng.Rows.Count
                                    If IsNumeric(displayRng.Cells(i, j).Value) Then
                                        If CheckBox3.Checked Then
                                            If displayRng.Cells(i, j).value IsNot Nothing Then
                                                Index = Index + 1
                                                ReDim Preserve Values(Index)
                                                Values(Index) = displayRng.Cells(i, j).value
                                            End If
                                        Else
                                            Index = Index + 1
                                            ReDim Preserve Values(Index)
                                            Values(Index) = displayRng.Cells(i, j).value
                                        End If
                                    End If
                                Next
                                OperatedValue = Operation(Values, ComboBox2.SelectedItem)
                                label.Text = OperatedValue
                                newHeight = height
                                newWidth = width
                            End If
                            label.Location = New System.Drawing.Point((j - 1) * newWidth, (combinedRow - 1) * height)
                            label.Height = newHeight
                            label.Width = newWidth
                            label.BorderStyle = BorderStyle.FixedSingle
                            label.TextAlign = ContentAlignment.MiddleCenter

                            If CheckBox1.Checked = True Then
                                Dim cell As Excel.Range = displayRng.Cells(1, j)
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

                        If ComboBox3.SelectedItem = "Into Top Row" Then

                            For j = 1 To displayRng.Columns.Count
                                For i = 2 To displayRng.Rows.Count
                                    Dim label As New System.Windows.Forms.Label
                                    If RadioButton6.Checked Then
                                        label.Text = displayRng.Cells(i, j).value
                                    ElseIf RadioButton5.Checked Then
                                        label.Text = ""
                                    End If
                                    label.Location = New System.Drawing.Point((j - 1) * newWidth, newHeight + (i - 2) * height)
                                    label.Height = height
                                    label.Width = newWidth
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

                                    CustomPanel2.Controls.Add(label)
                                Next
                            Next

                        ElseIf ComboBox3.SelectedItem = "Into Bottom Row" Then
                            For j = 1 To displayRng.Columns.Count
                                For i = 1 To displayRng.Rows.Count - 1
                                    Dim label As New System.Windows.Forms.Label
                                    If RadioButton6.Checked Then
                                        label.Text = displayRng.Cells(i, j).value
                                    ElseIf RadioButton5.Checked Then
                                        label.Text = ""
                                    End If
                                    label.Location = New System.Drawing.Point((j - 1) * newWidth, (i - 1) * height)
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

                    ElseIf RadioButton4.Checked Then
                        For j = 1 To displayRng.Columns.Count
                            Dim label As New System.Windows.Forms.Label
                            Dim HFactor As Integer
                            Dim WFactor As Integer

                            If ComboBox2.SelectedIndex <= 3 Then
                                Dim combinedValue As String = ""
                                Dim Separator As String

                                If ComboBox2.SelectedIndex = 3 Then
                                    Separator = vbNewLine
                                    HFactor = displayRng.Rows.Count / 1.75
                                    WFactor = 1
                                Else
                                    If CheckBox4.Checked Then
                                        Separator = ComboBox2.Text & vbNewLine
                                        HFactor = displayRng.Rows.Count / 1.75
                                        WFactor = 1
                                    Else
                                        Separator = ComboBox2.Text
                                        HFactor = 1
                                        WFactor = displayRng.Rows.Count
                                    End If
                                End If

                                For i = 1 To displayRng.Rows.Count - 1
                                    If CheckBox3.Checked Then
                                        If displayRng.Cells(i, j).value IsNot Nothing Then
                                            combinedValue = combinedValue & displayRng.Cells(i, j).Value & Separator
                                        End If
                                    Else
                                        combinedValue = combinedValue & displayRng.Cells(i, j).Value & Separator
                                    End If
                                Next

                                If CheckBox3.Checked Then
                                    If displayRng.Cells(displayRng.Rows.Count, j).value IsNot Nothing Then
                                        combinedValue = combinedValue & displayRng.Cells(displayRng.Rows.Count, j).Value
                                    Else
                                        If Len(combinedValue) >= Len(Separator) Then
                                            combinedValue = Mid(combinedValue, 1, Len(combinedValue) - Len(Separator))
                                        End If
                                    End If
                                Else
                                    combinedValue = combinedValue & displayRng.Cells(displayRng.Rows.Count, j).Value
                                End If
                                newWidth = width * WFactor
                                newHeight = height * HFactor
                                label.Text = combinedValue
                            Else
                                Dim OperatedValue As Double
                                Dim Values(0) As Double
                                Dim Index As Integer = -1
                                For i = 1 To displayRng.Rows.Count
                                    If IsNumeric(displayRng.Cells(i, j).Value) Then
                                        If CheckBox3.Checked Then
                                            If displayRng.Cells(i, j).value IsNot Nothing Then
                                                Index = Index + 1
                                                ReDim Preserve Values(Index)
                                                Values(Index) = displayRng.Cells(i, j).value
                                            End If
                                        Else
                                            Index = Index + 1
                                            ReDim Preserve Values(Index)
                                            Values(Index) = displayRng.Cells(i, j).value
                                        End If
                                    End If
                                Next
                                OperatedValue = Operation(Values, ComboBox2.SelectedItem)
                                newWidth = newWidth
                                newHeight = newHeight

                                label.Text = OperatedValue
                            End If
                            label.Location = New System.Drawing.Point((j - 1) * newWidth, 0)
                            If HFactor <> 1 Then
                                label.Height = newHeight
                            Else
                                label.Height = CustomPanel2.Height
                            End If
                            label.Width = newWidth
                            label.BorderStyle = BorderStyle.FixedSingle
                            label.TextAlign = ContentAlignment.MiddleCenter

                            If CheckBox1.Checked = True Then
                                Dim cell As Excel.Range = displayRng.Cells(1, j)
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

                    End If
                    CustomPanel2.AutoScroll = True

                ElseIf RadioButton3.Checked Then

                    Dim combinedRow As Integer
                    Dim combinedColumn As Integer
                    Dim combinedValue As String = ""
                    Dim OperatedValue As Double
                    Dim Values(0) As Double
                    Dim Index As Integer = -1

                    Dim Separator As String
                    Dim RowColumn As String

                    If ComboBox2.SelectedIndex = 3 Then
                        Separator = vbNewLine
                        RowColumn = "Row"
                    Else
                        If CheckBox4.Checked Then
                            Separator = ComboBox2.Text & vbNewLine
                            RowColumn = "Row"
                        Else
                            Separator = ComboBox2.Text
                            RowColumn = "Column"
                        End If
                    End If

                    For i = 1 To displayRng.Rows.Count - 1
                        For j = 1 To displayRng.Columns.Count
                            If ComboBox2.SelectedIndex <= 3 Then
                                If CheckBox3.Checked Then
                                    If displayRng.Cells(i, j).value IsNot Nothing Then
                                        combinedValue = combinedValue & displayRng.Cells(i, j).Value & Separator
                                    End If
                                Else
                                    combinedValue = combinedValue & displayRng.Cells(i, j).Value & Separator
                                End If
                            Else
                                If IsNumeric(displayRng.Cells(i, j).Value) Then
                                    If CheckBox3.Checked Then
                                        If displayRng.Cells(i, j).value IsNot Nothing Then
                                            Index = Index + 1
                                            ReDim Preserve Values(Index)
                                            Values(Index) = displayRng.Cells(i, j).value
                                        End If
                                    Else
                                        Index = Index + 1
                                        ReDim Preserve Values(Index)
                                        Values(Index) = displayRng.Cells(i, j).value
                                    End If
                                End If
                            End If
                        Next
                    Next

                    For j = 1 To displayRng.Columns.Count - 1
                        If ComboBox2.SelectedIndex <= 3 Then
                            If CheckBox3.Checked Then
                                If displayRng.Cells(displayRng.Rows.Count, j).value IsNot Nothing Then
                                    combinedValue = combinedValue & displayRng.Cells(displayRng.Rows.Count, j).Value & Separator
                                End If

                            Else
                                combinedValue = combinedValue & displayRng.Cells(rng.Rows.Count, j).Value & Separator
                            End If
                        Else
                            If IsNumeric(displayRng.Cells(displayRng.Rows.Count, j).Value) Then
                                If CheckBox3.Checked Then
                                    If displayRng.Cells(displayRng.Rows.Count, j).value IsNot Nothing Then
                                        Index = Index + 1
                                        ReDim Preserve Values(Index)
                                        Values(Index) = displayRng.Cells(displayRng.Rows.Count, j).value
                                    End If
                                Else
                                    Index = Index + 1
                                    ReDim Preserve Values(Index)
                                    Values(Index) = displayRng.Cells(displayRng.Rows.Count, j).value
                                End If
                            End If
                        End If
                    Next

                    If ComboBox2.SelectedIndex <= 3 Then
                        If CheckBox3.Checked Then
                            If displayRng.Cells(displayRng.Rows.Count, displayRng.Columns.Count).value IsNot Nothing Then
                                combinedValue = combinedValue & displayRng.Cells(displayRng.Rows.Count, displayRng.Columns.Count).Value
                            Else
                                If Len(combinedValue) >= Len(Separator) Then
                                    combinedValue = Mid(combinedValue, 1, Len(combinedValue) - Len(Separator))
                                End If
                            End If

                        Else
                            combinedValue = combinedValue & rng.Cells(rng.Rows.Count, rng.Columns.Count).Value
                        End If
                    Else
                        If IsNumeric(displayRng.Cells(displayRng.Rows.Count, displayRng.Columns.Count).Value) Then
                            If CheckBox3.Checked Then
                                If displayRng.Cells(displayRng.Rows.Count, displayRng.Columns.Count).value IsNot Nothing Then
                                    Index = Index + 1
                                    ReDim Preserve Values(Index)
                                    Values(Index) = displayRng.Cells(displayRng.Rows.Count, displayRng.Columns.Count).value
                                End If
                            Else
                                Index = Index + 1
                                ReDim Preserve Values(Index)
                                Values(Index) = displayRng.Cells(displayRng.Rows.Count, displayRng.Columns.Count).value
                            End If
                        End If
                    End If

                    OperatedValue = Operation(Values, ComboBox2.SelectedItem)

                    If RadioButton6.Checked Or RadioButton5.Checked Then
                        If ComboBox3.SelectedItem = "Into Top-Left Cell" Then
                            combinedRow = 1
                            combinedColumn = 1
                        ElseIf ComboBox3.SelectedItem = "Into Top-Right Cell" Then
                            combinedRow = 1
                            combinedColumn = displayRng.Columns.Count
                        ElseIf ComboBox3.SelectedItem = "Into Bottom-Left Cell" Then
                            combinedRow = displayRng.Rows.Count
                            combinedColumn = 1
                        ElseIf ComboBox3.SelectedItem = "Into Bottom-Right Cell" Then
                            combinedRow = displayRng.Rows.Count
                            combinedColumn = displayRng.Columns.Count
                        End If

                        For i = 1 To displayRng.Rows.Count
                            For j = 1 To displayRng.Columns.Count
                                If i = combinedRow And j = combinedColumn Then
                                    Dim label As New System.Windows.Forms.Label
                                    If ComboBox2.SelectedIndex <= 3 Then
                                        label.Text = combinedValue
                                    Else
                                        label.Text = OperatedValue
                                    End If

                                    If RowColumn = "Row" Then
                                        If i > combinedRow Then
                                            label.Location = New System.Drawing.Point((j - 1) * width, height * (displayRng.Cells.Count / 1.75) + (i - 2) * height)
                                        Else
                                            label.Location = New System.Drawing.Point((j - 1) * width, (i - 1) * height)
                                        End If
                                    Else
                                        If j > combinedColumn Then
                                            label.Location = New System.Drawing.Point(width * displayRng.Cells.Count + (j - 2) * width, (i - 1) * height)
                                        Else
                                            label.Location = New System.Drawing.Point((j - 1) * width, (i - 1) * height)
                                        End If
                                    End If

                                    If RowColumn = "Row" Then
                                        label.Height = height * displayRng.Cells.Count / 1.75
                                        label.Width = width
                                    Else
                                        label.Height = height
                                        label.Width = width * displayRng.Cells.Count
                                    End If

                                    label.BorderStyle = BorderStyle.FixedSingle
                                    label.TextAlign = ContentAlignment.MiddleCenter

                                    If CheckBox1.Checked = True Then
                                        Dim cell As Excel.Range = displayRng.Cells(combinedRow, combinedColumn)
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
                                Else

                                    Dim label As New System.Windows.Forms.Label

                                    If RadioButton6.Checked Then
                                        label.Text = displayRng.Cells(i, j).Value
                                    Else
                                        label.Text = ""
                                    End If
                                    If RowColumn = "Row" Then
                                        If i > combinedRow Then
                                            label.Location = New System.Drawing.Point((j - 1) * width, height * (displayRng.Cells.Count / 1.75) + (i - 2) * height)
                                        Else
                                            label.Location = New System.Drawing.Point((j - 1) * width, (i - 1) * height)
                                        End If
                                    Else
                                        If j > combinedColumn Then
                                            label.Location = New System.Drawing.Point((width * displayRng.Cells.Count) + (j - 2) * width, (i - 1) * height)
                                        Else
                                            label.Location = New System.Drawing.Point((j - 1) * width, (i - 1) * height)
                                        End If
                                    End If
                                    If RowColumn = "Row" Then
                                        If i = combinedRow Then
                                            label.Height = height * displayRng.Cells.Count / 1.75
                                            label.Width = width
                                        Else
                                            label.Height = height
                                            label.Width = width
                                        End If
                                    Else
                                        If j = combinedColumn Then
                                            label.Height = height
                                            label.Width = width * displayRng.Cells.Count
                                        Else
                                            label.Height = height
                                            label.Width = width
                                        End If
                                    End If
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
                                End If
                            Next
                        Next
                        CustomPanel2.AutoScroll = True

                    ElseIf RadioButton4.Checked Then

                        Dim label As New System.Windows.Forms.Label
                        If ComboBox2.SelectedIndex <= 3 Then
                            label.Text = combinedValue
                        Else
                            label.Text = OperatedValue
                        End If

                        label.Location = New System.Drawing.Point(0, 0)
                        If RowColumn = "Row" Then
                            label.Height = height * displayRng.Cells.Count
                            label.Width = CustomPanel2.Width
                        Else
                            label.Height = CustomPanel2.Height
                            label.Width = width * displayRng.Cells.Count
                        End If

                        label.BorderStyle = BorderStyle.FixedSingle
                        label.TextAlign = ContentAlignment.MiddleCenter

                        If CheckBox1.Checked = True Then
                            Dim cell As Excel.Range = displayRng.Cells(1, 1)
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

                        CustomPanel2.AutoScroll = True
                    End If
                End If

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Try

            Dim X4 As Boolean
            If RadioButton4.Checked Then
                X4 = True
            Else
                X4 = ComboBox3.SelectedIndex <> -1
            End If

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

            If RadioButton1.Checked = False And RadioButton2.Checked = False And RadioButton3.Checked = False Then
                MessageBox.Show("Select Where to Combine the Selected Data.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                workSheet.Activate()
                rng.Select()
                Exit Sub
            ElseIf X4 = False Then
                MessageBox.Show("Select Where to Store the Selected Data.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                workSheet.Activate()
                rng.Select()
                Exit Sub
            ElseIf RadioButton6.Checked = False And RadioButton5.Checked = False And RadioButton4.Checked = False Then
                MessageBox.Show("Select One Combination Option.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                ComboBox3.Focus()
                workSheet.Activate()
                rng.Select()
                Exit Sub
            End If

            If CheckBox2.Checked = True Then
                workSheet.Copy(After:=workBook.Sheets(workSheet.Name))
                workSheet.Activate()
            End If

            If (RadioButton1.Checked Or RadioButton2.Checked Or RadioButton3.Checked) And (RadioButton4.Checked Or RadioButton5.Checked Or RadioButton6.Checked) And X4 Then

                If RadioButton1.Checked Then

                    If RadioButton6.Checked Or RadioButton5.Checked Then
                        Dim combinedColumn As Integer
                        If ComboBox3.SelectedItem = "Into Left Column" Then
                            combinedColumn = 1
                        ElseIf ComboBox3.SelectedItem = "Into Right Column" Then
                            combinedColumn = rng.Columns.Count
                        End If
                        For i = 1 To rng.Rows.Count

                            If ComboBox2.SelectedIndex <= 3 Then
                                Dim combinedValue As String = ""
                                Dim Separator As String

                                If ComboBox2.SelectedIndex = 3 Then
                                    Separator = vbNewLine
                                Else
                                    If CheckBox4.Checked Then
                                        Separator = ComboBox2.Text & vbNewLine
                                    Else
                                        Separator = ComboBox2.Text
                                    End If
                                End If

                                For j = 1 To rng.Columns.Count - 1
                                    If CheckBox3.Checked Then
                                        If rng.Cells(i, j).value IsNot Nothing Then
                                            combinedValue = combinedValue & rng.Cells(i, j).Value & Separator
                                        End If
                                    Else
                                        combinedValue = combinedValue & rng.Cells(i, j).Value & Separator
                                    End If
                                Next

                                If CheckBox3.Checked Then
                                    If rng.Cells(i, rng.Columns.Count).value IsNot Nothing Then
                                        combinedValue = combinedValue & rng.Cells(i, rng.Columns.Count).Value
                                    Else
                                        If Len(combinedValue) >= Len(Separator) Then
                                            combinedValue = Mid(combinedValue, 1, Len(combinedValue) - Len(Separator))
                                        End If
                                    End If

                                Else
                                    combinedValue = combinedValue & rng.Cells(i, rng.Columns.Count).Value
                                End If

                                rng.Cells(i, combinedColumn).value = combinedValue

                            Else
                                Dim OperatedValue As Double
                                Dim Values(0) As Double
                                Dim Index As Integer = -1
                                For j = 1 To rng.Columns.Count
                                    If IsNumeric(rng.Cells(i, j).Value) Then
                                        If CheckBox3.Checked Then
                                            If rng.Cells(i, j).value IsNot Nothing Then
                                                Index = Index + 1
                                                ReDim Preserve Values(Index)
                                                Values(Index) = rng.Cells(i, j).value
                                            End If
                                        Else
                                            Index = Index + 1
                                            ReDim Preserve Values(Index)
                                            Values(Index) = rng.Cells(i, j).value
                                        End If
                                    End If
                                Next
                                OperatedValue = Operation(Values, ComboBox2.SelectedItem)
                                rng.Cells(i, combinedColumn).value = OperatedValue
                            End If

                        Next

                        If ComboBox3.SelectedItem = "Into Left Column" Then

                            For i = 1 To rng.Rows.Count
                                For j = 2 To rng.Columns.Count
                                    If RadioButton5.Checked Then
                                        rng.Cells(i, j).Clear()
                                    End If
                                Next
                            Next
                        ElseIf ComboBox3.SelectedItem = "Into Right Column" Then
                            For i = 1 To rng.Rows.Count
                                For j = 1 To rng.Columns.Count - 1
                                    If RadioButton5.Checked Then
                                        rng.Cells(i, j).Clear()
                                    End If
                                Next
                            Next
                            excelApp.DisplayAlerts = True
                        End If

                    ElseIf RadioButton4.Checked Then
                        excelApp.DisplayAlerts = False
                        For i = 1 To rng.Rows.Count
                            If ComboBox2.SelectedIndex <= 3 Then
                                Dim combinedValue As String = ""
                                Dim Separator As String

                                If ComboBox2.SelectedIndex = 3 Then
                                    Separator = vbNewLine
                                Else
                                    If CheckBox4.Checked Then
                                        Separator = ComboBox2.Text & vbNewLine
                                    Else
                                        Separator = ComboBox2.Text
                                    End If
                                End If
                                For j = 1 To rng.Columns.Count - 1
                                    If CheckBox3.Checked Then
                                        If rng.Cells(i, j).value IsNot Nothing Then
                                            combinedValue = combinedValue & rng.Cells(i, j).Value & Separator
                                        End If
                                    Else
                                        combinedValue = combinedValue & rng.Cells(i, j).Value & Separator
                                    End If
                                Next
                                If CheckBox3.Checked Then
                                    If rng.Cells(i, rng.Columns.Count).value IsNot Nothing Then
                                        combinedValue = combinedValue & rng.Cells(i, rng.Columns.Count).Value
                                    Else
                                        If Len(combinedValue) >= Len(Separator) Then
                                            combinedValue = Mid(combinedValue, 1, Len(combinedValue) - Len(Separator))
                                        End If
                                    End If
                                Else
                                    combinedValue = combinedValue & rng.Cells(i, rng.Columns.Count).Value
                                End If
                                rng.Cells(i, 1).value = combinedValue
                            Else
                                Dim OperatedValue As Double
                                Dim Values(0) As Double
                                Dim Index As Integer = -1
                                For j = 1 To rng.Columns.Count
                                    If IsNumeric(rng.Cells(i, j).Value) Then
                                        If CheckBox3.Checked Then
                                            If rng.Cells(i, j).value IsNot Nothing Then
                                                Index = Index + 1
                                                ReDim Preserve Values(Index)
                                                Values(Index) = rng.Cells(i, j).value
                                            End If
                                        Else
                                            Index = Index + 1
                                            ReDim Preserve Values(Index)
                                            Values(Index) = rng.Cells(i, j).value
                                        End If
                                    End If
                                Next
                                OperatedValue = Operation(Values, ComboBox2.SelectedItem)
                                rng.Cells(i, 1).value = OperatedValue
                            End If
                            rng.Rows(i).Merge
                        Next
                        excelApp.DisplayAlerts = True
                    End If
                ElseIf RadioButton2.Checked Then
                    If RadioButton6.Checked Or RadioButton5.Checked Then
                        Dim combinedRow As Integer
                        If ComboBox3.SelectedItem = "Into Top Row" Then
                            combinedRow = 1
                        ElseIf ComboBox3.SelectedItem = "Into Bottom Row" Then
                            combinedRow = rng.Rows.Count
                        End If
                        For j = 1 To rng.Columns.Count
                            If ComboBox2.SelectedIndex <= 3 Then
                                Dim combinedValue As String = ""
                                Dim Separator As String

                                If ComboBox2.SelectedIndex = 3 Then
                                    Separator = vbNewLine
                                Else
                                    If CheckBox4.Checked Then
                                        Separator = ComboBox2.Text & vbNewLine
                                    Else
                                        Separator = ComboBox2.Text
                                    End If
                                End If

                                For i = 1 To rng.Rows.Count - 1
                                    If CheckBox3.Checked Then
                                        If rng.Cells(i, j).value IsNot Nothing Then
                                            combinedValue = combinedValue & rng.Cells(i, j).Value & Separator
                                        End If
                                    Else
                                        combinedValue = combinedValue & rng.Cells(i, j).Value & Separator
                                    End If
                                Next
                                If CheckBox3.Checked Then
                                    If rng.Cells(rng.Rows.Count, j).value IsNot Nothing Then
                                        combinedValue = combinedValue & rng.Cells(rng.Rows.Count, j).Value
                                    Else
                                        If Len(combinedValue) >= Len(Separator) Then
                                            combinedValue = Mid(combinedValue, 1, Len(combinedValue) - Len(Separator))
                                        End If
                                    End If
                                Else
                                    combinedValue = combinedValue & rng.Cells(rng.Rows.Count, j).Value
                                End If
                                rng.Cells(combinedRow, j).Value = combinedValue
                            Else
                                Dim OperatedValue As Double
                                Dim Values(0) As Double
                                Dim Index As Integer = -1
                                For i = 1 To rng.Rows.Count
                                    If IsNumeric(rng.Cells(i, j).Value) Then
                                        If CheckBox3.Checked Then
                                            If rng.Cells(i, j).value IsNot Nothing Then
                                                Index = Index + 1
                                                ReDim Preserve Values(Index)
                                                Values(Index) = rng.Cells(i, j).value
                                            End If
                                        Else
                                            Index = Index + 1
                                            ReDim Preserve Values(Index)
                                            Values(Index) = rng.Cells(i, j).value
                                        End If
                                    End If
                                Next
                                OperatedValue = Operation(Values, ComboBox2.SelectedItem)
                                rng.Cells(combinedRow, j).Value = OperatedValue
                            End If
                        Next
                        If ComboBox3.SelectedItem = "Into Top Row" Then

                            For j = 1 To rng.Columns.Count
                                For i = 2 To rng.Rows.Count
                                    If RadioButton5.Checked Then
                                        rng.Cells(i, j).Clear()
                                    End If
                                Next
                            Next

                        ElseIf ComboBox3.SelectedItem = "Into Bottom Row" Then
                            For j = 1 To rng.Columns.Count
                                For i = 1 To rng.Rows.Count - 1
                                    If RadioButton5.Checked Then
                                        rng.Cells(i, j).Clear()
                                    End If
                                Next
                            Next
                            excelApp.DisplayAlerts = True
                        End If

                    ElseIf RadioButton4.Checked Then
                        excelApp.DisplayAlerts = False
                        For j = 1 To rng.Columns.Count
                            If ComboBox2.SelectedIndex <= 3 Then
                                Dim combinedValue As String = ""
                                Dim Separator As String

                                If ComboBox2.SelectedIndex = 3 Then
                                    Separator = vbNewLine
                                Else
                                    If CheckBox4.Checked Then
                                        Separator = ComboBox2.Text & vbNewLine
                                    Else
                                        Separator = ComboBox2.Text
                                    End If
                                End If

                                For i = 1 To rng.Rows.Count - 1
                                    If CheckBox3.Checked Then
                                        If rng.Cells(i, j).value IsNot Nothing Then
                                            combinedValue = combinedValue & rng.Cells(i, j).Value & ComboBox2.SelectedItem
                                        End If

                                    Else
                                        combinedValue = combinedValue & rng.Cells(i, j).Value & ComboBox2.SelectedItem
                                    End If
                                Next
                                If CheckBox3.Checked Then
                                    If rng.Cells(rng.Rows.Count, j).value IsNot Nothing Then
                                        combinedValue = combinedValue & rng.Cells(rng.Rows.Count, j).Value
                                    Else
                                        If Len(combinedValue) >= Len(Separator) Then
                                            combinedValue = Mid(combinedValue, 1, Len(combinedValue) - Len(Separator))
                                        End If
                                    End If
                                Else
                                    combinedValue = combinedValue & rng.Cells(rng.Rows.Count, j).Value
                                End If
                                rng.Cells(1, j).value = combinedValue

                            Else
                                Dim OperatedValue As Double
                                Dim Values(0) As Double
                                Dim Index As Integer = -1
                                For i = 1 To rng.Rows.Count
                                    If IsNumeric(rng.Cells(i, j).Value) Then
                                        If CheckBox3.Checked Then
                                            If rng.Cells(i, j).value IsNot Nothing Then
                                                Index = Index + 1
                                                ReDim Preserve Values(Index)
                                                Values(Index) = rng.Cells(i, j).value
                                            End If
                                        Else
                                            Index = Index + 1
                                            ReDim Preserve Values(Index)
                                            Values(Index) = rng.Cells(i, j).value
                                        End If
                                    End If
                                Next
                                OperatedValue = Operation(Values, ComboBox2.SelectedItem)
                                rng.Cells(1, j).value = OperatedValue
                            End If
                            rng.Columns(j).Merge
                        Next
                        excelApp.DisplayAlerts = True
                    End If

                ElseIf RadioButton3.Checked Then

                    Dim combinedRow As Integer
                    Dim combinedColumn As Integer
                    Dim combinedValue As String = ""
                    Dim OperatedValue As Double
                    Dim Values(0) As Double
                    Dim Index As Integer = -1
                    Dim Separator As String

                    If ComboBox2.SelectedIndex = 3 Then
                        Separator = vbNewLine
                    Else
                        If CheckBox4.Checked Then
                            Separator = ComboBox2.Text & vbNewLine
                        Else
                            Separator = ComboBox2.Text
                        End If
                    End If

                    For i = 1 To rng.Rows.Count - 1
                        For j = 1 To rng.Columns.Count
                            If ComboBox2.SelectedIndex <= 3 Then
                                If CheckBox3.Checked Then
                                    If rng.Cells(i, j).value IsNot Nothing Then
                                        combinedValue = combinedValue & rng.Cells(i, j).Value & Separator
                                    End If

                                Else
                                    combinedValue = combinedValue & rng.Cells(i, j).Value & Separator
                                End If
                            Else
                                If IsNumeric(rng.Cells(i, j).Value) Then
                                    If CheckBox3.Checked Then
                                        If rng.Cells(i, j).value IsNot Nothing Then
                                            Index = Index + 1
                                            ReDim Preserve Values(Index)
                                            Values(Index) = rng.Cells(i, j).value
                                        End If
                                    Else
                                        Index = Index + 1
                                        ReDim Preserve Values(Index)
                                        Values(Index) = rng.Cells(i, j).value
                                    End If
                                End If
                            End If
                        Next
                    Next

                    For j = 1 To rng.Columns.Count - 1
                        If ComboBox2.SelectedIndex <= 3 Then
                            If CheckBox3.Checked Then
                                If rng.Cells(rng.Rows.Count, j).value IsNot Nothing Then
                                    combinedValue = combinedValue & rng.Cells(rng.Rows.Count, j).Value & Separator
                                End If

                            Else
                                combinedValue = combinedValue & rng.Cells(rng.Rows.Count, j).Value & Separator
                            End If
                        Else
                            If IsNumeric(rng.Cells(rng.Rows.Count, j).Value) Then
                                If CheckBox3.Checked Then
                                    If rng.Cells(rng.Rows.Count, j).value IsNot Nothing Then
                                        Index = Index + 1
                                        ReDim Preserve Values(Index)
                                        Values(Index) = rng.Cells(rng.Rows.Count, j).value
                                    End If
                                Else
                                    Index = Index + 1
                                    ReDim Preserve Values(Index)
                                    Values(Index) = rng.Cells(rng.Rows.Count, j).value
                                End If
                            End If
                        End If
                    Next

                    If ComboBox2.SelectedIndex <= 3 Then
                        If CheckBox3.Checked Then
                            If rng.Cells(rng.Rows.Count, rng.Columns.Count).value IsNot Nothing Then
                                combinedValue = combinedValue & rng.Cells(rng.Rows.Count, rng.Columns.Count).Value
                            Else
                                If Len(combinedValue) >= Len(Separator) Then
                                    combinedValue = Mid(combinedValue, 1, Len(combinedValue) - Len(Separator))
                                End If
                            End If

                        Else
                            combinedValue = combinedValue & rng.Cells(rng.Rows.Count, rng.Columns.Count).Value
                        End If
                    Else
                        If IsNumeric(rng.Cells(rng.Rows.Count, rng.Columns.Count).Value) Then
                            If CheckBox3.Checked Then
                                If rng.Cells(rng.Rows.Count, rng.Columns.Count).value IsNot Nothing Then
                                    Index = Index + 1
                                    ReDim Preserve Values(Index)
                                    Values(Index) = rng.Cells(rng.Rows.Count, rng.Columns.Count).value
                                End If
                            Else
                                Index = Index + 1
                                ReDim Preserve Values(Index)
                                Values(Index) = rng.Cells(rng.Rows.Count, rng.Columns.Count).value
                            End If
                        End If
                    End If

                    OperatedValue = Operation(Values, ComboBox2.SelectedItem)
                    If RadioButton6.Checked Or RadioButton5.Checked Then
                        If ComboBox3.SelectedItem = "Into Top-Left Cell" Then
                            combinedRow = 1
                            combinedColumn = 1
                        ElseIf ComboBox3.SelectedItem = "Into Top-Right Cell" Then
                            combinedRow = 1
                            combinedColumn = rng.Columns.Count
                        ElseIf ComboBox3.SelectedItem = "Into Bottom-Left Cell" Then
                            combinedRow = rng.Rows.Count
                            combinedColumn = 1
                        ElseIf ComboBox3.SelectedItem = "Into Bottom-Right Cell" Then
                            combinedRow = rng.Rows.Count
                            combinedColumn = rng.Columns.Count
                        End If

                        For i = 1 To rng.Rows.Count
                            For j = 1 To rng.Columns.Count
                                If i = combinedRow And j = combinedColumn Then
                                    If ComboBox2.SelectedIndex <= 3 Then
                                        rng.Cells(i, j).value = combinedValue
                                    Else
                                        rng.Cells(i, j).value = OperatedValue
                                    End If

                                Else

                                    If RadioButton5.Checked Then
                                        rng.Cells(i, j).Clear
                                    End If
                                End If
                            Next
                        Next

                    ElseIf RadioButton4.Checked Then
                        If ComboBox2.SelectedIndex <= 3 Then
                            rng.Cells(1, 1).value = combinedValue
                        Else
                            rng.Cells(1, 1).value = OperatedValue
                        End If
                        excelApp.DisplayAlerts = False
                        rng.Merge()
                        excelApp.DisplayAlerts = True
                    End If
                End If
                For j = 1 To rng.Columns.Count
                    rng.Columns(j).Autofit
                Next

                If CheckBox1.Checked = False Then
                    rng.ClearFormats()
                End If

                If CheckBox4.Checked = False Then
                    For i = 1 To rng.Rows.Count
                        rng.Rows(i).Autofit
                    Next
                End If

                Me.Close()

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
            Call Display()
            TextBoxChanged = False
        Catch ex As Exception
        End Try
    End Sub
    Private Sub TextBox1_GotFocus(sender As Object, e As EventArgs) Handles TextBox1.GotFocus

        Try
            FocusedTextBox = 1

        Catch ex As Exception

        End Try

    End Sub
    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        Try
            Call Display()
        Catch ex As Exception
        End Try
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged

        Try
            If RadioButton1.Checked Then

                ComboBox3.Text = ""
                ComboBox3.Items.Clear()
                ComboBox3.Items.Add("Into Left Column")
                ComboBox3.Items.Add("Into Right Column")

                Call Display()

            End If

        Catch ex As Exception

        End Try

    End Sub
    Private Sub RadioButton6_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton6.CheckedChanged
        Try

            If RadioButton6.Checked Then
                If ComboBox3.Enabled = False Then
                    ComboBox3.Enabled = True
                    Label3.Enabled = True
                    ComboBox3.Items.Clear()
                    If RadioButton1.Checked Then
                        ComboBox3.Items.Add("Into Left Column")
                        ComboBox3.Items.Add("Into Right Column")
                    ElseIf RadioButton2.Checked Then
                        ComboBox3.Items.Add("Into Top Row")
                        ComboBox3.Items.Add("Into Bottom Row")
                    ElseIf RadioButton3.Checked Then
                        ComboBox3.Items.Add("Into Top-Left Cell")
                        ComboBox3.Items.Add("Into Top-Right Cell")
                        ComboBox3.Items.Add("Into Bottom-Left Cell")
                        ComboBox3.Items.Add("Into Bottom-Right Cell")
                    End If

                End If

                Call Display()

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub RadioButton5_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton5.CheckedChanged

        Try
            If RadioButton5.Checked Then

                If ComboBox3.Enabled = False Then

                    ComboBox3.Enabled = True
                    Label3.Enabled = True
                    ComboBox3.Items.Clear()

                    If RadioButton1.Checked Then
                        ComboBox3.Items.Add("Into Left Column")
                        ComboBox3.Items.Add("Into Right Column")
                    ElseIf RadioButton2.Checked Then
                        ComboBox3.Items.Add("Into Top Row")
                        ComboBox3.Items.Add("Into Bottom Row")
                    ElseIf RadioButton3.Checked Then
                        ComboBox3.Items.Add("Into Top-Left Cell")
                        ComboBox3.Items.Add("Into Top-Right Cell")
                        ComboBox3.Items.Add("Into Bottom-Left Cell")
                        ComboBox3.Items.Add("Into Bottom-Right Cell")
                    End If
                End If
                Call Display()
            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub RadioButton4_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton4.CheckedChanged
        Try
            If RadioButton4.Checked Then

                ComboBox3.SelectedText = ""
                ComboBox3.Items.Clear()
                ComboBox3.Enabled = False
                Label3.Enabled = False

                Call Display()
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged

        Try
            If RadioButton2.Checked Then

                ComboBox3.Text = ""
                ComboBox3.Items.Clear()
                ComboBox3.Items.Add("Into Top Row")
                ComboBox3.Items.Add("Into Bottom Row")

                Call Display()

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged

        Try
            If RadioButton3.Checked Then

                ComboBox3.Text = ""
                ComboBox3.Items.Clear()
                ComboBox3.Items.Add("Into Top-Left Cell")
                ComboBox3.Items.Add("Into Top-Right Cell")
                ComboBox3.Items.Add("Into Bottom-Left Cell")
                ComboBox3.Items.Add("Into Bottom-Right Cell")

                Call Display()

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub AutoSelection_Click(sender As Object, e As EventArgs) Handles AutoSelection.Click

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

            If Mid(sheetName, Len(sheetName), 1) = "'" Then
                sheetName = Mid(sheetName, 1, Len(sheetName) - 1)
            End If

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

    Private Sub Selection_Click(sender As Object, e As EventArgs) Handles Selection.Click
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

            If Mid(sheetName, Len(sheetName), 1) = "'" Then
                sheetName = Mid(sheetName, 1, Len(sheetName) - 1)
            End If

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

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Me.Close()
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
    Private Sub Button2_MouseLeave(sender As Object, e As EventArgs) Handles Button2.MouseLeave
        Try

            Button2.BackColor = Color.FromArgb(255, 255, 255)
            Button2.ForeColor = Color.FromArgb(70, 70, 70)
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
    Private Sub Button1_MouseLeave(sender As Object, e As EventArgs) Handles Button1.MouseLeave
        Try

            Button1.BackColor = Color.FromArgb(255, 255, 255)
            Button1.ForeColor = Color.FromArgb(70, 70, 70)
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

    Private Sub CheckBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles CheckBox3.KeyDown

        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub CheckBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles CheckBox4.KeyDown

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

    Private Sub CustomGroupBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles CustomGroupBox2.KeyDown

        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub CustomGroupBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles CustomGroupBox3.KeyDown

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

    Private Sub PictureBox1_KeyDown(sender As Object, e As KeyEventArgs) 

        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub PictureBox2_KeyDown(sender As Object, e As KeyEventArgs) 

        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub PictureBox3_KeyDown(sender As Object, e As KeyEventArgs) 

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

    Private Sub RadioButton1_KeyDown(sender As Object, e As KeyEventArgs) Handles RadioButton1.KeyDown

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

    Private Sub RadioButton6_KeyDown(sender As Object, e As KeyEventArgs) Handles RadioButton6.KeyDown

        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub Selection_KeyDown(sender As Object, e As KeyEventArgs) Handles Selection.KeyDown

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

    Private Sub Form18_CombineRanges_Load(sender As Object, e As EventArgs) Handles Me.Load
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

    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        Try
            Call Display()
        Catch ex As Exception
        End Try
    End Sub

    Private Sub CheckBox4_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox4.CheckedChanged
        Try
            Call Display()
        Catch ex As Exception
        End Try
    End Sub

    Private Sub ComboBox2_TextChanged(sender As Object, e As EventArgs) Handles ComboBox2.TextChanged
        Try
            Call Display()
        Catch ex As Exception
        End Try
    End Sub

    Private Sub Form18_CombineRanges_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        form_flag = False
    End Sub

    Private Sub Form18_CombineRanges_Disposed(sender As Object, e As EventArgs) Handles Me.Disposed
        form_flag = False
    End Sub

    Private Sub Form18_CombineRanges_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        Me.Focus()
        Me.BringToFront()
        Me.Activate()
        Me.BeginInvoke(New System.Action(Sub()
                                             TextBox1.Text = rng.Address
                                             SetWindowPos(Me.Handle, New IntPtr(HWND_TOPMOST), 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOSIZE)
                                         End Sub))
    End Sub
End Class