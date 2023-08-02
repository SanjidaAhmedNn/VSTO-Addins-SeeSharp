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

        Dim cellPattern As String = "(\$?[A-Z]+\$?[0-9]+)"

        Dim referencePattern As String = "^" + cellPattern + "(:" + cellPattern + ")?$"

        Dim regex As New Regex(referencePattern)

        If regex.IsMatch(cellReference) Then
            Return True
        Else
            Return False
        End If

    End Function
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

    Private Function SearchDiagonally(Rng, r, c)

        Dim rowEqual As Integer = SearchAlongRow(Rng, r, c)

        Dim activesheet As Excel.Worksheet = CType(excelApp.ActiveSheet, Excel.Worksheet)

        Dim Rng2 As Excel.Range
        Rng2 = activesheet.Range(Rng.Cells(1, 1), Rng.Cells(1, rowEqual))

        Dim Output(1) As Integer
        Output(0) = 1
        Output(1) = rowEqual

        Dim TotalCells As Integer = Rng2.Cells.Count

        Dim j As Integer

        j = 0

        Dim min As Integer = Rng.Rows.Count

        While SearchAlongColumn(Rng, r, c + j) > 1 And j + 1 <= rowEqual
            If SearchAlongColumn(Rng, r, c + j) <= min Then
                min = SearchAlongColumn(Rng, r, c + j)
            End If
            If activesheet.Range(Rng.Cells(1, 1), Rng.Cells(min, j + 1)).Cells.Count >= TotalCells Then
                Output(0) = min
                Output(1) = j + 1
                TotalCells = Rng2.Cells.Count
            End If
            j = j + 1

        End While

        SearchDiagonally = Output

    End Function
    Private Function CrossCheck(excelApp As Excel.Application, rng1 As Excel.Range, rng2 As Excel.Range)

        Dim intersectRange As Range = excelApp.Intersect(rng1, rng2)

        If intersectRange Is Nothing Then
            Return False
        Else
            Return True
        End If

    End Function

    Private Function RemoveCrossings(excelApp, Arr)

        Dim activesheet As Excel.Worksheet = CType(excelApp.ActiveSheet, Excel.Worksheet)
        Dim Rng1 As Excel.Range
        Dim Rng2 As Excel.Range
        Dim Count1 As Integer
        Dim Count2 As Integer
        For i = LBound(Arr, 1) To UBound(Arr, 1)
            If Arr(i, 0) > 0 Then
                Rng1 = activesheet.Range("A1")
                Rng1 = activesheet.Range(Rng1.Cells(Arr(i, 0), Arr(i, 1)), Rng1.Cells(Arr(i, 2), Arr(i, 3)))

                For j = LBound(Arr, 1) To UBound(Arr, 1)
                    If i <> j Then
                        Rng2 = activesheet.Range("A1")
                        If Arr(j, 0) > 0 Then
                            Rng2 = activesheet.Range(Rng2.Cells(Arr(j, 0), Arr(j, 1)), Rng2.Cells(Arr(j, 2), Arr(j, 3)))

                            If CrossCheck(excelApp, Rng1, Rng2) = True Then

                                Count1 = Rng1.Cells.Count
                                Count2 = Rng2.Cells.Count

                                If Count1 < Count2 Then
                                    Arr(i, 0) = 0
                                    Exit For
                                ElseIf Count1 = Count2 Then
                                    If (Rng1.Rows.Count = 1 Or Rng1.Columns.Count = 1) Then
                                        Arr(i, 0) = 0
                                        Exit For
                                    Else
                                        Arr(j, 0) = 0
                                    End If
                                Else
                                    Arr(j, 0) = 0
                                End If
                            End If
                        End If
                    End If
                Next
            End If
        Next

        RemoveCrossings = Arr

    End Function
    Private Function IsWithinRange(r As Integer, c As Integer, Rng As Excel.Range)

        If r >= Rng.Cells(1, 1).Row And r <= Rng.Cells(Rng.Rows.Count, 1).Row And c >= Rng.Cells(1, 1).Column And r <= Rng.Cells(1, Rng.Columns.Count).Column Then

            IsWithinRange = True
        Else
            IsWithinRange = False
        End If

    End Function
    Private Sub Display()

        Try

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

                    Dim activesheet As Excel.Worksheet = CType(excelApp.ActiveSheet, Excel.Worksheet)

                    Dim Arr(r * C - 1, 3) As Integer
                    For i = LBound(Arr, 1) To UBound(Arr, 1)
                        Arr(i, 0) = 0
                    Next

                    Dim Count As Integer = 0

                    For i = 1 To r
                        For j = 1 To C

                            Dim rowEqual As Integer = SearchDiagonally(displayRng, i, j)(0)
                            Dim columnEqual As Integer = SearchDiagonally(displayRng, i, j)(1)

                            Arr(Count, 0) = i
                            Arr(Count, 1) = j
                            Arr(Count, 2) = i + rowEqual - 1
                            Arr(Count, 3) = j + columnEqual - 1

                            Count = Count + 1
                        Next j
                    Next i

                    Arr = RemoveCrossings(excelApp, Arr)

                    For i = 1 To r
                        For j = 1 To C

                            Dim MRng As Excel.Range
                            MRng = activesheet.Range(displayRng.Cells(i, j).Address)

                            For m = LBound(Arr, 1) To UBound(Arr, 1)

                                If Arr(m, 0) > 0 Then

                                    Dim Rng1 As Excel.Range = activesheet.Range("A1")
                                    Rng1 = activesheet.Range(Rng1.Cells(Arr(m, 0), Arr(m, 1)), Rng1.Cells(Arr(m, 2), Arr(m, 3)))

                                    If IsWithinRange(i, j, Rng1) = True Then
                                        MRng = Rng1
                                        Arr(m, 0) = 0
                                        Exit For
                                    End If

                                End If

                            Next

                            Dim newWidth As Single = width * MRng.Columns.Count
                            Dim newHeight = height * MRng.Rows.Count

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
                        Next
                    Next

                End If

                CustomPanel2.AutoScroll = True

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub Form8_Load(sender As Object, e As EventArgs) Handles Me.Load

        Try

            excelApp = Globals.ThisAddIn.Application

            Me.Label2.Enabled = False
            Me.TextBox3.Enabled = False
            Me.PictureBox6.Enabled = False

            AddHandler excelApp.SheetSelectionChange, AddressOf excelApp_SheetSelectionChange

            opened = opened + 1

        Catch ex As Exception

        End Try

    End Sub

    Private Sub excelApp_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Excel.Range)

        Try

            Dim selectedRange As Excel.Range
            selectedRange = excelApp.Selection

            If FocusedTextBox = 1 Then
                TextBox1.Text = selectedRange.Address
                workBook = excelApp.ActiveWorkbook
                workSheet = workBook.ActiveSheet
                rng = selectedRange
                TextBox1.Focus()

            ElseIf FocusedTextBox = 3 Then
                TextBox3.Text = selectedRange.Address
                workbook2 = excelApp.ActiveWorkbook
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

            Call Display()

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
                MessageBox.Show("Enter a Valid Source Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                TextBox1.Focus()
                workSheet.Activate()
                rng.Select()
                Exit Sub
            End If

            If (RadioButton1.Checked = False And RadioButton2.Checked = False And RadioButton3.Checked = False) Then
                MessageBox.Show("Select a Merge Type.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                workSheet.Activate()
                rng.Select()
                Exit Sub
            End If

            If (RadioButton9.Checked = False And RadioButton10.Checked = False) Then
                MessageBox.Show("Select a Destination Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                workSheet.Activate()
                rng.Select()
                Exit Sub
            End If

            If (RadioButton10.Checked = True And (TextBox3.Text = "" Or IsValidExcelCellReference(TextBox3.Text) = False)) Then
                MessageBox.Show("Enter a Valid Destination Cell.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                TextBox3.Focus()
                workSheet.Activate()
                rng.Select()
                Exit Sub
            End If

            If CheckBox2.Checked = True Then
                workSheet.Copy(After:=workBook.Sheets(workSheet.Name))
                workSheet2.Activate()
            End If

            rng2 = workSheet2.Range(rng2.Cells(1, 1), rng2.Cells(rng.Rows.Count, rng.Columns.Count))
            workSheet2.Activate()

            If Overlap(excelApp, workSheet, workSheet2, rng, rng2) = True Then
                rng2 = rng
                If CheckBox1.Checked = False Then
                    rng2.ClearFormats()
                End If
            Else
                rng.Copy()
                rng2.PasteSpecial(Excel.XlPasteType.xlPasteValues)
                If CheckBox1.Checked = True Then
                    rng2.PasteSpecial(Excel.XlPasteType.xlPasteFormats)
                End If
                excelApp.CutCopyMode = Excel.XlCutCopyMode.xlCopy
            End If

            rng2.Select()

            Dim r As Integer = rng2.Rows.Count
            Dim c As Integer = rng.Columns.Count

            Dim i As Integer
            Dim j As Integer

            If RadioButton1.Checked = True Or RadioButton2.Checked = True Or RadioButton3.Checked = True Then

                excelApp.DisplayAlerts = False

                If RadioButton1.Checked = True Then
                    For i = 1 To r
                        For j = 1 To c
                            Dim rowEqual As Integer = SearchAlongRow(rng2, i, j)
                            workSheet2.Range(rng2.Cells(i, j), rng2.Cells(i, j + rowEqual - 1)).Merge()
                            If rowEqual > 1 Then
                                rng2.Cells(i, j).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                            End If
                            j = j + rowEqual - 1
                        Next
                    Next

                ElseIf RadioButton2.Checked = True = True Then

                    For j = 1 To c
                        For i = 1 To r
                            Dim columnEqual As Integer = SearchAlongColumn(rng2, i, j)
                            workSheet2.Range(rng2.Cells(i, j), rng2.Cells(i + columnEqual - 1, j)).Merge()
                            If columnEqual > 1 Then
                                rng2.Cells(i, j).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                            End If
                            i = i + columnEqual - 1
                        Next
                    Next

                ElseIf RadioButton3.Checked = True Then

                    Dim activesheet As Excel.Worksheet = CType(excelApp.ActiveSheet, Excel.Worksheet)

                    Dim Arr(r * c - 1, 3) As Integer
                    For i = LBound(Arr, 1) To UBound(Arr, 1)
                        Arr(i, 0) = 0
                    Next

                    Dim Count As Integer = 0

                    For i = 1 To r
                        For j = 1 To c
                            Dim rowEqual As Integer = SearchDiagonally(rng2, i, j)(0)
                            Dim columnEqual As Integer = SearchDiagonally(rng2, i, j)(1)
                            Arr(Count, 0) = i
                            Arr(Count, 1) = j
                            Arr(Count, 2) = i + rowEqual - 1
                            Arr(Count, 3) = j + columnEqual - 1
                            Count = Count + 1
                        Next j
                    Next i

                    Arr = RemoveCrossings(excelApp, Arr)

                    For i = 1 To r
                        For j = 1 To c

                            Dim MRng As Excel.Range
                            MRng = activesheet.Range(rng2.Cells(i, j).Address)

                            For m = LBound(Arr, 1) To UBound(Arr, 1)

                                If Arr(m, 0) > 0 Then

                                    Dim Rng1 As Excel.Range = activesheet.Range("A1")
                                    Rng1 = activesheet.Range(Rng1.Cells(Arr(m, 0), Arr(m, 1)), Rng1.Cells(Arr(m, 2), Arr(m, 3)))

                                    If IsWithinRange(i, j, Rng1) = True Then
                                        MRng = Rng1
                                        Arr(m, 0) = 0
                                        Exit For
                                    End If

                                End If

                            Next

                            workSheet2.Range(rng2.Cells(i, j), rng2.Cells(i + MRng.Rows.Count - 1, j + MRng.Columns.Count - 1)).Merge()
                            If MRng.Columns.Count > 1 Then
                                rng2.Cells(i, j).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                            End If
                            If MRng.Rows.Count > 1 Then
                                rng2.Cells(i, j).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                            End If
                        Next
                    Next

                End If
                excelApp.DisplayAlerts = False

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub PictureBox6_Click(sender As Object, e As EventArgs) Handles PictureBox6.Click

        Try
            FocusedTextBox = 3
            Me.Hide()

            excelApp = Globals.ThisAddIn.Application
            workbook2 = excelApp.ActiveWorkbook

            Dim userInput As Excel.Range = excelApp.InputBox("Select a range", Type:=8)
            rng2 = userInput


            Dim sheetName As String
            sheetName = Split(rng2.Address(True, True, Excel.XlReferenceStyle.xlA1, True), "]")(1)
            sheetName = Split(sheetName, "!")(0)

            If Mid(sheetName, Len(sheetName), 1) = "'" Then
                sheetName = Mid(sheetName, 1, Len(sheetName) - 1)
            End If

            workSheet2 = workbook2.Worksheets(sheetName)
            workSheet2.Activate()

            rng2.Select()

            TextBox3.Text = rng2.Address

            Me.Show()
            TextBox3.Focus()

        Catch ex As Exception

            Me.Show()
            TextBox3.Focus()

        End Try

    End Sub

    Private Sub PictureBox9_Click(sender As Object, e As EventArgs) Handles PictureBox9.Click

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

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged

        Try
            excelApp = Globals.ThisAddIn.Application
            workbook2 = excelApp.ActiveWorkbook
            workSheet2 = workbook2.ActiveSheet

            rng2 = workSheet2.Range(TextBox3.Text)
            rng2.Select()

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

    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged

        Try
            If RadioButton3.Checked = True Then
                Call Display()
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

    Private Sub RadioButton9_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton9.CheckedChanged

        Try
            If RadioButton9.Checked = True Then
                workbook2 = workBook
                workSheet2 = workSheet
                rng2 = rng
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub RadioButton10_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton10.CheckedChanged

        Try
            If RadioButton10.Checked = True Then
                Label2.Enabled = True
                TextBox3.Enabled = True
                PictureBox6.Enabled = True
                TextBox3.Focus()
            Else
                TextBox3.Clear()
                Label2.Enabled = False
                TextBox3.Enabled = False
                PictureBox6.Enabled = False
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

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Try
            Me.Close()

        Catch ex As Exception

        End Try
    End Sub

    Private Sub PictureBox9_GotFocus(sender As Object, e As EventArgs) Handles PictureBox9.GotFocus
        Try
            FocusedTextBox = 1

        Catch ex As Exception

        End Try
    End Sub

    Private Sub PictureBox4_GotFocus(sender As Object, e As EventArgs) Handles PictureBox4.GotFocus
        Try
            FocusedTextBox = 1

        Catch ex As Exception

        End Try
    End Sub

    Private Sub PictureBox6_GotFocus(sender As Object, e As EventArgs) Handles PictureBox6.GotFocus
        Try
            FocusedTextBox = 3

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

    Private Sub PictureBox5_GotFocus(sender As Object, e As EventArgs) Handles PictureBox5.GotFocus
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

    Private Sub RadioButton9_GotFocus(sender As Object, e As EventArgs) Handles RadioButton9.GotFocus
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

            Button2.BackColor = Color.FromArgb(255, 255, 255)
            Button2.ForeColor = Color.FromArgb(70, 70, 70)
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

    Private Sub PictureBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles PictureBox1.KeyDown

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

    Private Sub PictureBox9_KeyDown(sender As Object, e As KeyEventArgs) Handles PictureBox9.KeyDown

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

End Class