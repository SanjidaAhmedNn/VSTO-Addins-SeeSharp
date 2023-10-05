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
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.ToolBar
Imports Microsoft.Office.Interop.Excel

Public Class Form28_Split_text_bypattern

    Dim WithEvents excelApp As Excel.Application
    Dim workBook As Excel.Workbook
    Dim workSheet As Excel.Worksheet
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
    Private Function MatchArr(Arr, source, index)

        Dim Matched(1) As Object
        Matched(0) = False
        Dim value As String

        For i = LBound(Arr) To UBound(Arr)
            value = Mid(source, index, Len(Arr(i)))
            If Arr(i) = value Then
                Matched(0) = True
                Matched(1) = Arr(i)
                Exit For
            End If
        Next

        MatchArr = Matched

    End Function
    Private Function SplitText(Source As String, Pattern As String, Consecutive As Boolean, KeepSeparator As Boolean, Before As Boolean)

        Dim SplitValues(0) As String
        Dim Index As Integer = -1
        Dim Start As Integer = 1
        Dim SearchStart As Integer = 1
        Dim Ending As Integer
        Dim separator As String = ""

        For i = 1 To Len(Pattern)

            If Mid(Pattern, i, 1) <> "*" Then
                Dim SeparatorLength As Integer = 1
                While Mid(Pattern, i + SeparatorLength, 1) <> "*" And (i + SeparatorLength) <= Len(Pattern)
                    SeparatorLength = SeparatorLength + 1
                End While

                separator = Mid(Pattern, i, SeparatorLength)

                Ending = InStr(SearchStart, Source, separator)

                If Ending <> 0 Then
                    Index = Index + 1
                    ReDim Preserve SplitValues(Index)
                    If KeepSeparator = True Then
                        If Before = True Then
                            SplitValues(Index) = Mid(Source, Start, Ending - Start)
                        Else
                            SplitValues(Index) = Mid(Source, Start, Ending - Start + Len(separator))
                        End If
                    Else
                        SplitValues(Index) = Mid(Source, Start, Ending - Start)
                    End If
                    SearchStart = Ending + Len(separator)
                    Start = Ending + Len(separator)
                    If Consecutive = True Then
                        While Mid(Source, SearchStart, Len(separator)) = separator
                            SearchStart = SearchStart + Len(separator)
                            Start = Start + Len(separator)
                        End While
                    End If
                    If KeepSeparator = True Then
                        If Before = True Then
                            Start = Start - Len(separator)
                        End If
                    End If
                End If

            End If

        Next

        Ending = Len(Source) + 1
        Index = Index + 1
        ReDim Preserve SplitValues(Index)

        If KeepSeparator = True Then
            If Before = True Then
                SplitValues(Index) = Mid(Source, Start, Ending - Start)
            Else
                SplitValues(Index) = Mid(Source, Start, Ending - Start + Len(separator))
            End If
        Else
            SplitValues(Index) = Mid(Source, Start, Ending - Start)
        End If

        SplitText = SplitValues

    End Function
    Private Function SplitCount(Source As String, Pattern As String, Consecutive As Boolean)

        Dim Index As Integer = 0
        Dim SearchStart As Integer = 1
        Dim Ending As Integer
        Dim separator As String = ""

        For i = 1 To Len(Pattern)

            If Mid(Pattern, i, 1) <> "*" Then
                Dim SeparatorLength As Integer = 1
                While Mid(Pattern, i + SeparatorLength, 1) <> "*" And (i + SeparatorLength) <= Len(Pattern)
                    SeparatorLength = SeparatorLength + 1
                End While

                separator = Mid(Pattern, i, SeparatorLength)

                Ending = InStr(SearchStart, Source, separator)

                If Ending <> 0 Then
                    Index = Index + 1
                    SearchStart = Ending + Len(separator)
                    If Consecutive = True Then
                        While Mid(Source, SearchStart, Len(separator)) = separator
                            SearchStart = SearchStart + Len(separator)
                        End While
                    End If
                End If

            End If

        Next

        Ending = Len(Source) + 1
        Index = Index + 1

        SplitCount = Index

    End Function

    Private Sub Display()

        Panel_InputRange.Controls.Clear()
        Panel_ExpectedOutput.Controls.Clear()

        excelApp = Globals.ThisAddIn.Application
        workBook = excelApp.ActiveWorkbook
        workSheet = workBook.ActiveSheet

        rng = workSheet.Range(TB_source_range.Text)

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
            Height = Panel_InputRange.Height / displayRng.Rows.Count
        Else
            Height = (119 / 4)
        End If

        BaseWidth = (260 / 3)
        Width = ((MaxOfColumn(displayRng) * BaseWidth) / 10)

        Dim Width1 As Double

        If Width > Panel_InputRange.Width Then
            Width1 = Width
        Else
            Width1 = Panel_InputRange.Width
        End If
        Dim ordinate As Double = 0

        For i = 1 To r
            Dim label As New System.Windows.Forms.Label
            label.Text = displayRng.Cells(i, 1).Value
            label.Location = New System.Drawing.Point(ordinate, (i - 1) * Height)
            label.Height = Height
            label.Width = Width1
            label.BorderStyle = BorderStyle.FixedSingle
            label.TextAlign = ContentAlignment.MiddleCenter

            If CB_formatting.Checked = True Then

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
            Panel_InputRange.Controls.Add(label)
        Next

        Panel_InputRange.AutoScroll = True

        Dim X1 As Boolean = RB_rows.Checked
        Dim X2 As Boolean = RB_columns.Checked


        If (X1 Or X2) And ComboBox2.Text <> "" Then

            Dim pattern As String = ComboBox2.Text

            Dim Consecutive As Boolean
            If CB_consecutive_separators.Checked Then
                Consecutive = True
            Else
                Consecutive = False
            End If

            Dim KeepSeparator As Boolean
            If CB_separators_finaloutput.Checked Then
                KeepSeparator = True
            Else
                KeepSeparator = False
            End If

            Dim Before As Boolean
            If RB_starting_point.Checked Then
                Before = True
            Else
                Before = False
            End If

            If X1 Then
                Dim values(0) As String
                Dim Index As Integer = -1
                For i = 1 To r
                    Dim source As String = displayRng.Cells(i, 1).value
                    Dim SplitValues() As String
                    SplitValues = SplitText(source, pattern, Consecutive, KeepSeparator, Before)
                    For m = LBound(SplitValues) To UBound(SplitValues)
                        Index = Index + 1
                        ReDim Preserve values(Index)
                        values(Index) = SplitValues(m)
                    Next
                Next

                Dim Width2 As Double = (MaxOfArray(values) * BaseWidth) / 10
                If Width + Width2 < Panel_ExpectedOutput.Width Then
                    Width2 = Panel_ExpectedOutput.Width - Width
                End If
                Dim abscissa1 As Double = 0
                Dim abscissa2 As Double = 0
                For i = 1 To r
                    Dim source As String = displayRng.Cells(i, 1).value
                    Dim SplitValues() As String
                    SplitValues = SplitText(source, pattern, Consecutive, KeepSeparator, Before)

                    Dim label As New System.Windows.Forms.Label
                    label.Text = displayRng.Cells(i, 1).Value
                    label.Location = New System.Drawing.Point(0, abscissa1)
                    label.Height = Height
                    label.Width = Width
                    label.BorderStyle = BorderStyle.FixedSingle
                    label.TextAlign = ContentAlignment.MiddleCenter

                    If CB_formatting.Checked = True Then
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
                    Panel_ExpectedOutput.Controls.Add(label)
                    abscissa1 = abscissa1 + Height
                    For m = LBound(SplitValues) + 1 To UBound(SplitValues)
                        Dim label1 As New System.Windows.Forms.Label
                        label1.Text = ""
                        label1.Location = New System.Drawing.Point(0, abscissa1)
                        label1.Height = Height
                        label1.Width = Width
                        label1.BorderStyle = BorderStyle.FixedSingle
                        label1.TextAlign = ContentAlignment.MiddleCenter

                        If CB_formatting.Checked = True Then
                            Dim cell As Excel.Range = displayRng.Cells(i, 1)
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
                        Panel_ExpectedOutput.Controls.Add(label1)
                        abscissa1 = abscissa1 + Height
                    Next

                    For m = LBound(SplitValues) To UBound(SplitValues)
                        Dim label1 As New System.Windows.Forms.Label
                        label1.Text = SplitValues(m)
                        label1.Location = New System.Drawing.Point(Width, abscissa2)
                        label1.Height = Height
                        label1.Width = Width2
                        label1.BorderStyle = BorderStyle.FixedSingle
                        label1.TextAlign = ContentAlignment.MiddleCenter

                        If CB_formatting.Checked = True Then
                            Dim cell As Excel.Range = displayRng.Cells(i, 1)
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
                        Panel_ExpectedOutput.Controls.Add(label1)
                        abscissa2 = abscissa2 + Height
                    Next
                Next


            ElseIf X2 Then
                ordinate = 0

                For i = 1 To displayRng.Rows.Count
                    Dim label As New System.Windows.Forms.Label
                    label.Text = displayRng.Cells(i, 1).Value
                    label.Location = New System.Drawing.Point(ordinate, (i - 1) * Height)
                    label.Height = Height
                    label.Width = Width
                    label.BorderStyle = BorderStyle.FixedSingle
                    label.TextAlign = ContentAlignment.MiddleCenter

                    If CB_formatting.Checked = True Then
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
                    Panel_ExpectedOutput.Controls.Add(label)
                Next
                ordinate = ordinate + Width
                Dim lengths(r - 1) As Integer
                For i = 1 To displayRng.Rows.Count
                    Dim source As String = displayRng.Cells(i, 1).value
                    lengths(i - 1) = SplitCount(source, pattern, Consecutive)
                Next
                Dim TotalWidth As Integer = FindMax(lengths)

                Dim Values(r - 1, TotalWidth - 1) As String

                For i = 1 To displayRng.Rows.Count
                    Dim source As String = displayRng.Cells(i, 1).value
                    Dim SplitValues() As String
                    SplitValues = SplitText(source, pattern, Consecutive, KeepSeparator, Before)
                    For j = LBound(SplitValues) To UBound(SplitValues)
                        Values(i - 1, j) = SplitValues(j)
                    Next
                Next
                For j = 0 To TotalWidth - 1
                    Dim ColumnValues(r - 1) As String
                    For i = 0 To r - 1
                        ColumnValues(i) = Values(i, j)
                    Next
                    Width1 = (MaxOfArray(ColumnValues) * BaseWidth) / 10
                    For i = 0 To r - 1
                        Dim label1 As New System.Windows.Forms.Label
                        label1.Text = ColumnValues(i)
                        label1.Location = New System.Drawing.Point(ordinate, i * Height)
                        label1.Height = Height
                        label1.Width = Width1
                        label1.BorderStyle = BorderStyle.FixedSingle
                        label1.TextAlign = ContentAlignment.MiddleCenter

                        If CB_formatting.Checked = True Then
                            Dim cell As Excel.Range = displayRng.Cells(i + 1, 1)
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
                        Panel_ExpectedOutput.Controls.Add(label1)
                    Next
                    ordinate = ordinate + Width1
                Next
            End If
            Panel_ExpectedOutput.AutoScroll = True
        End If

    End Sub
    Private Sub CB_separators_finaloutput_CheckedChanged(sender As Object, e As EventArgs) Handles CB_separators_finaloutput.CheckedChanged
        Try
            If CB_separators_finaloutput.Checked = True Then
                RB_starting_point.Enabled = True
                RB_ending_point.Enabled = True
                PictureBox4.Enabled = True
                PictureBox3.Enabled = True

            ElseIf CB_separators_finaloutput.Checked = False Then
                RB_starting_point.Enabled = False
                RB_ending_point.Enabled = False
                PictureBox4.Enabled = False
                PictureBox3.Enabled = False
            End If
            Call Display()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub CB_consecutive_separators_CheckedChanged(sender As Object, e As EventArgs) Handles CB_consecutive_separators.CheckedChanged
        Try
            Call Display()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Btn_OK_Click(sender As Object, e As EventArgs) Handles Btn_OK.Click

        Try
            If TB_source_range.Text = "" Then
                MessageBox.Show("Select a Source Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                TB_source_range.Focus()
                workSheet.Activate()
                rng.Select()
                Exit Sub
            End If

            If IsValidExcelCellReference(TB_source_range.Text) = False Then
                MessageBox.Show("Select a Valid Source Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                TB_source_range.Focus()
                workSheet.Activate()
                rng.Select()
                Exit Sub
            End If

            Dim r As Integer = rng.Rows.Count
            Dim c As Integer = rng.Columns.Count

            Dim X1 As Boolean = RB_rows.Checked
            Dim X2 As Boolean = RB_columns.Checked

            If X1 = False And X2 = False Then
                MessageBox.Show("Select a Split Option.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                workSheet.Activate()
                rng.Select()
                Exit Sub
            End If

            If CB_backup.Checked = True Then
                workSheet.Copy(After:=workBook.Sheets(workSheet.Name))
            End If

            workSheet.Activate()

            If (X1 Or X2) And ComboBox2.Text <> "" Then


                Dim pattern As String = ComboBox2.Text

                Dim Consecutive As Boolean
                If CB_consecutive_separators.Checked Then
                    Consecutive = True
                Else
                    Consecutive = False
                End If

                Dim KeepSeparator As Boolean
                If CB_separators_finaloutput.Checked Then
                    KeepSeparator = True
                Else
                    KeepSeparator = False
                End If

                Dim Before As Boolean
                If RB_starting_point.Checked Then
                    Before = True
                Else
                    Before = False
                End If

                If X1 Then

                    Dim Arr(r - 1) As String
                    Dim Lengths(r - 1) As String
                    Dim RowNumber As Integer = 0
                    For i = 1 To r
                        Dim source As String = rng.Cells(i, 1).Value
                        Arr(i - 1) = source
                        Dim SplitValues() As String
                        SplitValues = SplitText(source, pattern, Consecutive, KeepSeparator, Before)
                        Lengths(i - 1) = UBound(SplitValues) + 1
                        For m = LBound(SplitValues) To UBound(SplitValues)
                            RowNumber = RowNumber + 1
                            rng.Cells(RowNumber, 2) = SplitValues(m)
                            If CB_formatting.Checked Then
                                rng.Cells(i, 1).Copy()
                                rng.Cells(RowNumber, 2).PasteSpecial(Excel.XlPasteType.xlPasteFormats)
                            Else
                                rng.Cells(RowNumber, 2).ClearFormats()
                            End If
                        Next
                    Next

                    RowNumber = 0
                    For i = 1 To r
                        RowNumber = RowNumber + 1
                        rng.Cells(RowNumber, 1) = Arr(i - 1)
                        If CB_formatting.Checked Then
                            rng.Cells(RowNumber, 2).Copy()
                            rng.Cells(RowNumber, 1).PasteSpecial(Excel.XlPasteType.xlPasteFormats)
                        Else
                            rng.Cells(RowNumber, 1).ClearFormats()
                        End If
                        For m = 1 To Lengths(i - 1) - 1
                            RowNumber = RowNumber + 1
                            rng.Cells(RowNumber, 1) = ""
                            If CB_formatting.Checked Then
                                rng.Cells(RowNumber, 2).Copy()
                                rng.Cells(RowNumber, 1).PasteSpecial(Excel.XlPasteType.xlPasteFormats)
                            Else
                                rng.Cells(RowNumber, 1).ClearFormats()
                            End If
                        Next
                    Next

                    excelApp.CutCopyMode = False

                    rng2 = workSheet.Range(rng.Cells(1, 1), rng.Cells(RowNumber, 2))
                    rng2.Select()
                    For j = 1 To rng2.Columns.Count
                        rng2.Columns(j).AutoFit()
                    Next

                ElseIf X2 Then

                    Dim MaxColumns As Integer = 1
                    For i = 1 To r
                        Dim source As String = rng.Cells(i, 1).Value
                        Dim SplitValues() As String
                        SplitValues = SplitText(source, pattern, Consecutive, KeepSeparator, Before)
                        If UBound(SplitValues) + 1 > MaxColumns Then
                            MaxColumns = UBound(SplitValues) + 1
                        End If
                        If CB_formatting.Checked = False Then
                            rng.Cells(i, 1).ClearFormats()
                        End If
                        For m = LBound(SplitValues) To UBound(SplitValues)
                            rng.Cells(i, m + 2) = SplitValues(m)
                        Next
                    Next
                    For i = 1 To r
                        If CB_formatting.Checked Then
                            rng.Cells(i, 1).Copy()

                            For m = 1 To MaxColumns
                                rng.Cells(i, m + 1).PasteSpecial(Excel.XlPasteType.xlPasteFormats)
                            Next m
                        Else
                            For m = 1 To MaxColumns
                                rng.Cells(i, m + 1).ClearFormats()
                            Next m
                        End If
                    Next

                    excelApp.CutCopyMode = False

                    rng2 = workSheet.Range(rng.Cells(1, 1), rng.Cells(r, MaxColumns + 1))
                    rng2.Select()
                    For j = 1 To rng2.Columns.Count
                        rng2.Columns(j).AutoFit()
                    Next

                End If

                Me.Close()

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub TB_source_range_TextChanged(sender As Object, e As EventArgs) Handles TB_source_range.TextChanged

        Try
            excelApp = Globals.ThisAddIn.Application
            workBook = excelApp.ActiveWorkbook
            workSheet = workBook.ActiveSheet

            TB_source_range.SelectionStart = TB_source_range.Text.Length
            TB_source_range.ScrollToCaret()

            rng = workSheet.Range(TB_source_range.Text)
            TextBoxChanged = True
            rng.Select()

            Call Display()

            TextBoxChanged = False

        Catch ex As Exception

        End Try

    End Sub

    Private Sub RB_rows_CheckedChanged(sender As Object, e As EventArgs) Handles RB_rows.CheckedChanged

        Try
            If RB_rows.Checked Then
                Call Display()
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub RB_columns_CheckedChanged(sender As Object, e As EventArgs) Handles RB_columns.CheckedChanged

        Try
            If RB_columns.Checked Then
                Call Display()
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

    Private Sub RB_starting_point_CheckedChanged(sender As Object, e As EventArgs) Handles RB_starting_point.CheckedChanged

        Try
            Call Display()
        Catch ex As Exception

        End Try

    End Sub

    Private Sub RB_ending_point_CheckedChanged(sender As Object, e As EventArgs) Handles RB_ending_point.CheckedChanged

        Try
            Call Display()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub CB_formatting_CheckedChanged(sender As Object, e As EventArgs) Handles CB_formatting.CheckedChanged
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
            Me.TB_source_range.Text = rng.Address
            Me.TB_source_range.Focus()

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

            TB_source_range.Text = rng.Address
            TB_source_range.Focus()

        Catch ex As Exception

        End Try

    End Sub

    Private Sub TB_source_range_GotFocus(sender As Object, e As EventArgs) Handles TB_source_range.GotFocus
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

    Private Sub Btn_OK_MouseEnter(sender As Object, e As EventArgs) Handles Btn_OK.MouseEnter
        Try
            Btn_OK.BackColor = Color.FromArgb(65, 105, 225)
            Btn_OK.ForeColor = Color.FromArgb(255, 255, 255)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Btn_Cancel_MouseEnter(sender As Object, e As EventArgs) Handles Btn_Cancel.MouseEnter

        Try
            Btn_Cancel.BackColor = Color.FromArgb(65, 105, 225)
            Btn_Cancel.ForeColor = Color.FromArgb(255, 255, 255)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub Btn_OK_MouseLeave(sender As Object, e As EventArgs) Handles Btn_OK.MouseLeave
        Try

            Btn_OK.BackColor = Color.FromArgb(255, 255, 255)
            Btn_OK.ForeColor = Color.FromArgb(70, 70, 70)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Btn_Cancel_MouseLeave(sender As Object, e As EventArgs) Handles Btn_Cancel.MouseLeave
        Try

            Btn_Cancel.BackColor = Color.FromArgb(255, 255, 255)
            Btn_Cancel.ForeColor = Color.FromArgb(70, 70, 70)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub Btn_Cancel_Click(sender As Object, e As EventArgs) Handles Btn_Cancel.Click
        Try
            Me.Close()
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

    Private Sub Form27_Split_text_bystrings_Load(sender As Object, e As EventArgs) Handles Me.Load

        Try

            excelApp = Globals.ThisAddIn.Application
            workBook = excelApp.ActiveWorkbook
            workSheet = workBook.ActiveSheet
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
                    TB_source_range.Text = selectedRange.Address
                    workSheet = workBook.ActiveSheet
                    rng = selectedRange
                    TB_source_range.Focus()
                End If
            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub Form27_Split_text_bystrings_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown

        Try
            If e.KeyCode = Keys.Enter Then

                Call Btn_OK_Click(sender, e)

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub Form28_Split_text_bypattern_Disposed(sender As Object, e As EventArgs) Handles Me.Disposed
        form_flag = False
    End Sub

    Private Sub Form28_Split_text_bypattern_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        form_flag = False
    End Sub

    Private Sub Form28_Split_text_bypattern_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        Me.Focus()
        Me.BringToFront()
        Me.Activate()
        Me.BeginInvoke(New System.Action(Sub()
                                             TB_source_range.Text = rng.Address
                                             SetWindowPos(Me.Handle, New IntPtr(HWND_TOPMOST), 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOSIZE)
                                         End Sub))
    End Sub
End Class