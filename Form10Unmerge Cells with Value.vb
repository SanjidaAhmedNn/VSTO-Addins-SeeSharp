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
Public Class Form10

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
    Private Function SearchInArray(i, j, Arr)

        Dim Result As Object = 0

        For k = LBound(Arr, 1) To UBound(Arr, 1)
            If Arr(k, 0) = i And Arr(k, 1) = j Then
                Result = Arr(k, 2)
                Exit For
            End If
        Next k

        SearchInArray = Result

    End Function
    Private Function Available(i, j, Arr)

        Dim Result As Boolean = False

        For k = LBound(Arr, 1) To UBound(Arr, 1)
            If Arr(k, 0) = i And Arr(k, 1) = j Then
                Result = True
                Exit For
            End If
        Next k

        Available = Result

    End Function
    Private Sub Display()

        Try
            CustomPanel1.Controls.Clear()
            CustomPanel2.Controls.Clear()


            Dim displayRng As Excel.Range

            If Rng.Rows.Count > 50 Then
                displayRng = workSheet.Range(Rng.Cells(1, 1), Rng.Cells(50, Rng.Columns.Count))
            Else
                displayRng = workSheet.Range(Rng.Cells(1, 1), Rng.Cells(Rng.Rows.Count, Rng.Columns.Count))
            End If

            Dim r As Integer
            Dim c As Integer

            r = displayRng.Rows.Count
            c = displayRng.Columns.Count

            Dim height As Single
            Dim width As Single

            If r <= 6 Then
                height = CustomPanel1.Height / r
            Else
                height = CustomPanel1.Height / 6
            End If

            If c <= 4 Then
                width = CustomPanel1.Width / c
            Else
                width = CustomPanel1.Width / 4
            End If

            Dim Arr((r * c) - 1, 1) As Object

            Dim Count As Integer = 0

            For i = 1 To r
                For j = 1 To C
                    If Available(i, j, Arr) = False Then
                        If displayRng.Cells(i, j).MergeCells = True Then
                            For k = 2 To displayRng.Cells(i, j).MergeArea.Columns.Count
                                If Available(i, j + k - 1, Arr) = False Then
                                    Arr(Count, 0) = i
                                    Arr(Count, 1) = j + k - 1
                                    Count = Count + 1
                                End If
                            Next k
                            For m = 2 To displayRng.Cells(i, j).MergeArea.Rows.Count
                                For n = 1 To displayRng.Cells(i, j).MergeArea.Columns.Count
                                    If Available(i + m - 1, j + n - 1, Arr) = False Then
                                        Arr(Count, 0) = i + m - 1
                                        Arr(Count, 1) = j + n - 1
                                        Count = Count + 1
                                    End If
                                Next n
                            Next m
                        End If
                    End If
                Next j
            Next i

            For i = 1 To r
                For j = 1 To C
                    If Available(i, j, Arr) = False Then
                        Dim height2 As Single = height * displayRng.Cells(i, j).MergeArea.Rows.Count
                        Dim width2 As Single = width * displayRng.Cells(i, j).MergeArea.Columns.Count
                        Dim label As New System.Windows.Forms.Label
                        label.Text = displayRng.Cells(i, j).Value
                        label.Location = New System.Drawing.Point((j - 1) * width, (i - 1) * height)
                        label.Height = height2
                        label.Width = width2
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
                    End If
                Next j
            Next i
            CustomPanel1.AutoScroll = True

            Dim Arr2((r * c) - 1, 2) As Object

            Count = 0

            For i = 1 To r
                For j = 1 To C
                    If rng.Cells(i, j).MergeCells = True And Available(i, j, Arr2) = False Then
                        For m = 0 To rng.Cells(i, j).MergeArea.Rows.Count - 1
                            For n = 0 To rng.Cells(i, j).MergeArea.Columns.Count - 1
                                Arr2(Count, 0) = i + m
                                Arr2(Count, 1) = j + n
                                Arr2(Count, 2) = displayRng.Cells(i, j).Value
                                Count = Count + 1
                            Next n
                        Next m
                    End If
                Next j
            Next i

            For i = 1 To r
                For j = 1 To c
                    Dim label As New System.Windows.Forms.Label
                    If Rng.Cells(i, j).MergeCells = True Then
                        label.Text = SearchInArray(i, j, Arr2)
                    Else
                        label.Text = displayRng.Cells(i, j).Value
                    End If
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
                Next j
            Next i

            CustomPanel2.AutoScroll = True

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
            Else
                rng.Copy()
                rng2.PasteSpecial(Excel.XlPasteType.xlPasteValues)
                rng2.PasteSpecial(Excel.XlPasteType.xlPasteFormats)
                excelApp.CutCopyMode = Excel.XlCutCopyMode.xlCopy
            End If

            rng2.Select()

            Dim r As Integer = rng2.Rows.Count
            Dim C As Integer = rng2.Columns.Count

            For i = 1 To r
                For j = 1 To C
                    If rng2.Cells(i, j).MergeCells = True Then
                        Dim Merged_Rows As Integer = rng2.Cells(i, j).MergeArea.Rows.Count
                        Dim Merged_Columns As Integer = rng2.Cells(i, j).MergeArea.Columns.Count
                        rng2.Cells(i, j).UnMerge
                        For m = 0 To Merged_Rows - 1
                            For n = 0 To Merged_Columns - 1
                                rng2.Cells(i + m, j + n).value = rng2.Cells(i, j).value
                            Next n
                        Next m
                    End If
                Next j
            Next i

            If CheckBox1.Checked = False Then
                rng2.ClearFormats()
            End If

            Me.Close()

        Catch ex As Exception

        End Try

    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click

        Try
            FocusedTextBox = 1
            Me.Hide()

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

    Private Sub PictureBox9_Click(sender As Object, e As EventArgs) Handles PictureBox9.Click

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

    Private Sub Form10_Load(sender As Object, e As EventArgs) Handles Me.Load

        Try

            excelApp = Globals.ThisAddIn.Application
            workBook = excelApp.ActiveWorkbook
            workbook2 = excelApp.ActiveWorkbook
            workSheet = workBook.ActiveSheet
            workSheet2 = workbook2.ActiveSheet

            AddHandler excelApp.SheetSelectionChange, AddressOf excelApp_SheetSelectionChange

            opened = opened + 1

            Me.Label3.Enabled = False
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

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged

        Try
            Call Display()
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
            rng.Select()

            Call Display()

        Catch ex As Exception

        End Try

    End Sub

    Private Sub TextBox1_GotFocus(sender As Object, e As EventArgs) Handles TextBox1.GotFocus

        Try
            FocusedTextBox = 1
        Catch ex As Exception

        End Try

    End Sub

    Private Sub RadioButton9_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton9.CheckedChanged

        Try
            If RadioButton9.Checked = True Then
                workSheet2 = workSheet
                rng2 = rng
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub RadioButton10_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton10.CheckedChanged

        Try
            If RadioButton10.Checked = True Then
                Label3.Enabled = True
                TextBox3.Enabled = True
                TextBox3.Focus()
                PictureBox6.Enabled = True
            Else
                Label3.Enabled = False
                TextBox3.Clear()
                TextBox3.Enabled = False
                PictureBox6.Enabled = False
            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged

        Try
            workSheet2 = workbook2.ActiveSheet

            TextBox3.SelectionStart = TextBox3.Text.Length
            TextBox3.ScrollToCaret()

            rng2 = workSheet2.Range(TextBox3.Text)
            rng2.Select()

        Catch ex As Exception

        End Try

    End Sub

    Private Sub PictureBox6_Click(sender As Object, e As EventArgs) Handles PictureBox6.Click

        Try
            FocusedTextBox = 3
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

            TextBox3.Text = rng2.Address

            Me.Show()
            TextBox3.Focus()

        Catch ex As Exception

            Me.Show()
            TextBox3.Focus()

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

    Private Sub PictureBox1_GotFocus(sender As Object, e As EventArgs) Handles PictureBox1.GotFocus
        Try
            FocusedTextBox = 1
        Catch ex As Exception

        End Try
    End Sub

    Private Sub PictureBox9_GotFocus(sender As Object, e As EventArgs) Handles PictureBox9.GotFocus
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

    Private Sub Label3_GotFocus(sender As Object, e As EventArgs) Handles Label3.GotFocus
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

    Private Sub RadioButton9_GotFocus(sender As Object, e As EventArgs) Handles RadioButton9.GotFocus
        Try
            FocusedTextBox = 0
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

    Private Sub PictureBox6_KeyDown(sender As Object, e As KeyEventArgs) Handles PictureBox6.KeyDown
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

    Private Sub RadioButton10_KeyDown(sender As Object, e As KeyEventArgs) Handles RadioButton10.KeyDown
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

    Private Sub Button1_MouseEnter(sender As Object, e As EventArgs) Handles Button1.MouseEnter
        Try
            Button1.BackColor = Color.FromArgb(65, 105, 225)
            Button1.ForeColor = Color.FromArgb(255, 255, 255)
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

    Private Sub Button1_MouseLeave(sender As Object, e As EventArgs) Handles Button1.MouseLeave
        Try

            Button1.BackColor = Color.FromArgb(255, 255, 255)
            Button1.ForeColor = Color.FromArgb(70, 70, 70)
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

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Me.Close()
        Catch ex As Exception

        End Try
    End Sub
End Class