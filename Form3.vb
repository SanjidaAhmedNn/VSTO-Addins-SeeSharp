Imports System.Drawing
Imports System.Windows.Forms
Imports System.Reflection.Emit
Imports Microsoft.Office.Interop.Excel
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Threading
Imports System.Diagnostics
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Button
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports System.Windows.Forms.Application
Imports System.Text.RegularExpressions

Public Class Form3

    Public WithEvents excelApp As Excel.Application

    Public workbook As Excel.Workbook
    Public workbook2 As Excel.Workbook

    Public worksheet As Excel.Worksheet
    Public worksheet2 As Excel.Worksheet

    Public rng As Excel.Range
    Public rng2 As Excel.Range
    Public FocusedTextBox As Integer
    Public Opened As Integer

    Public Form4Open As Integer
    Public Workbook2Opened As Boolean


    Private Function IsValidExcelCellReference(cellReference As String) As Boolean

        ' Regular expression pattern for a cell reference.
        ' This pattern will match references like A1, $A$1, etc.
        Dim cellPattern As String = "(\$?[A-Z]+\$?[0-9]+)"

        ' Regular expression pattern for an Excel reference.
        ' This pattern will match references like A1:B13, $A$1:$B$13, A1, $B$1, etc.
        Dim referencePattern As String = "^" + cellPattern + "(:" + cellPattern + ")?$"

        ' Create a regex object with the pattern.
        Dim regex As New Regex(referencePattern)

        ' Test the input string against the regex pattern.
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

            panel1.Controls.Clear()
            panel2.Controls.Clear()

            Dim displayRng As Excel.Range

            If rng.Rows.Count > 50 Then
                displayRng = rng.Rows("1:50")
            Else
                displayRng = rng
            End If

            Dim r As Integer
            Dim c As Integer

            r = displayRng.Rows.Count
            c = displayRng.Columns.Count

            Dim height As Integer
            Dim width As Integer

            If r <= 6 Then
                height = panel1.Height / r
            Else
                height = panel1.Height / 6
            End If

            If c <= 4 Then
                width = panel1.Width / c
            Else
                width = panel1.Width / 4
            End If

            For i = 1 To displayRng.Rows.Count
                For j = 1 To displayRng.Columns.Count
                    Dim label As New System.Windows.Forms.Label
                    label.Text = displayRng.Cells(i, j).Value
                    label.Location = New System.Drawing.Point((j - 1) * width, (i - 1) * height)
                    label.Height = height
                    label.Width = width
                    label.BorderStyle = BorderStyle.FixedSingle
                    label.TextAlign = ContentAlignment.MiddleCenter

                    If CheckBox2.Checked = True Then

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
                    panel1.Controls.Add(label)
                Next
            Next

            panel1.AutoScroll = True

            If (RadioButton2.Checked = True Or RadioButton3.Checked = True) Then

                If c <= 6 Then
                    height = panel2.Height / c
                Else
                    height = panel2.Height / 6
                End If

                If r <= 4 Then
                    width = panel2.Width / r
                Else
                    width = panel2.Width / 4
                End If

                For i = 1 To displayRng.Rows.Count
                    For j = 1 To displayRng.Columns.Count
                        Dim label As New System.Windows.Forms.Label
                        label.Text = displayRng.Cells(i, j).Value
                        label.Location = New System.Drawing.Point((i - 1) * width, (j - 1) * height)
                        label.Height = height
                        label.Width = width
                        label.BorderStyle = BorderStyle.FixedSingle
                        label.TextAlign = ContentAlignment.MiddleCenter

                        If CheckBox2.Checked = True Then
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

                        panel2.Controls.Add(label)
                    Next
                Next

                panel2.AutoScroll = True

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub DestinationChange()

        Try
            If RadioButton1.Checked = True Then
                If Form4Open = 1 Then
                    If Me.Workbook2Opened = True Then
                        workbook2.Close()
                        workbook.Activate()
                    End If
                    workbook2 = workbook
                    Form4Open = 0
                End If
                TextBox2.Visible = True
                PictureBox2.Visible = True
                TextBox2.Location = New System.Drawing.Point(121, 7)
                PictureBox2.Location = New System.Drawing.Point(226, 7)
                TextBox2.Focus()
            Else
                TextBox2.Clear()
            End If

            If RadioButton4.Checked = True Then

                If Me.Form4Open = 1 Then
                    If Me.Workbook2Opened = True Then
                        workbook2.Close()
                        workbook.Activate()
                    End If
                    workbook2 = workbook
                    Me.Form4Open = 0
                End If
                TextBox2.Visible = True
                PictureBox2.Visible = True
                TextBox2.Location = New System.Drawing.Point(121, 30)
                PictureBox2.Location = New System.Drawing.Point(226, 30)

                Dim ws As Excel.Worksheet = CType(workbook.Worksheets.Add(), Excel.Worksheet)
                TextBox2.Focus()
            Else
                TextBox2.Clear()
            End If

            If RadioButton5.Checked = True And Form4Open = 0 Then
                TextBox2.Visible = False
                PictureBox2.Visible = False
                Dim MyForm4 As New Form4
                MyForm4.excelApp = Me.excelApp
                MyForm4.workbook = Me.workbook
                MyForm4.worksheet = Me.worksheet
                MyForm4.rng = Me.rng
                MyForm4.Opened = Me.Opened
                MyForm4.FocusedTextBox = Me.FocusedTextBox
                MyForm4.Form4Open = Me.Form4Open
                MyForm4.Workbook2Opened = False
                If Me.RadioButton3.Checked = True Then
                    MyForm4.GB6 = 3
                ElseIf Me.RadioButton2.Checked = True Then
                    MyForm4.GB6 = 2
                Else
                    MyForm4.GB6 = 0
                End If
                If Me.CheckBox1.Checked = True Then
                    MyForm4.CB1 = 1
                Else
                    MyForm4.CB1 = 0
                End If
                If Me.CheckBox2.Checked = True Then
                    MyForm4.CB2 = 1
                Else
                    MyForm4.CB2 = 0
                End If
                Me.Close()
                MyForm4.Show()

            End If

        Catch ex As Exception

        End Try


    End Sub


    Private Sub btn_OK_MouseLeave(sender As Object, e As EventArgs) Handles btn_OK.MouseLeave

        Try

            btn_OK.ForeColor = Color.FromArgb(70, 70, 70)
            btn_OK.BackColor = Color.White

        Catch ex As Exception

        End Try

    End Sub

    Private Sub btn_cancel_MouseLeave(sender As Object, e As EventArgs) Handles btn_cancel.MouseLeave

        Try

            btn_cancel.ForeColor = Color.FromArgb(70, 70, 70)
            btn_cancel.BackColor = Color.White

        Catch ex As Exception

        End Try

    End Sub

    Private Sub PictureBox8_Click(sender As Object, e As EventArgs) Handles PictureBox8.Click

        Try
            FocusedTextBox = 1
            Me.Hide()


            Dim userInput As Excel.Range = excelApp.InputBox("Select a range", Type:=8)
            rng = userInput

            Dim sheetName As String
            sheetName = Split(rng.Address(True, True, Excel.XlReferenceStyle.xlA1, True), "]")(1)
            sheetName = Split(sheetName, "!")(0)
            worksheet = workbook.Worksheets(sheetName)
            worksheet.Activate()

            rng.Select()

            TextBox1.Text = rng.Address

            Me.Show()
            TextBox1.Focus()

        Catch ex As Exception

            Me.Show()
            TextBox1.Focus()

        End Try

    End Sub

    Private Sub PictureBox4_Click(sender As Object, e As EventArgs) Handles PictureBox4.Click

        Try
            FocusedTextBox = 1
            Me.Hide()


            Dim userInput As Excel.Range = excelApp.InputBox("Select a range", Type:=8)
            rng = userInput

            Dim sheetName As String
            sheetName = Split(rng.Address(True, True, Excel.XlReferenceStyle.xlA1, True), "]")(1)
            sheetName = Split(sheetName, "!")(0)
            worksheet = workbook.Worksheets(sheetName)
            worksheet.Activate()

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


    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs)

        Try

            TextBox2.Location = New System.Drawing.Point(121, 7)
            PictureBox2.Location = New System.Drawing.Point(226, 7)

        Catch ex As Exception

        End Try

    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click

        Try
            FocusedTextBox = 2
            Me.Hide()


            Dim userInput As Excel.Range = excelApp.InputBox("Select a Cell.", Type:=8)
            rng2 = userInput

            Dim sheetName As String
            sheetName = Split(rng2.Address(True, True, Excel.XlReferenceStyle.xlA1, True), "]")(1)
            sheetName = Split(sheetName, "!")(0)
            worksheet2 = workbook.Worksheets(sheetName)
            worksheet2.Activate()

            rng2.Select()

            TextBox2.Text = rng2.Address

            Me.Show()
            TextBox2.Focus()

        Catch ex As Exception

            Me.Show()
            TextBox2.Focus()

        End Try

    End Sub

    Private Sub btn_OK_Click(sender As Object, e As EventArgs) Handles btn_OK.Click

        Try
            If TextBox1.Text = "" Or IsValidExcelCellReference(TextBox1.Text) = False Then
                MessageBox.Show("Enter a Valid Source Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                worksheet.Activate()
                rng.Select()
                Exit Sub
            End If

            If RadioButton2.Checked = False And RadioButton3.Checked = False Then
                MessageBox.Show("Select a Paste Option.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                worksheet.Activate()
                rng.Select()
                Exit Sub
            End If

            If RadioButton1.Checked = False And RadioButton4.Checked = False And RadioButton5.Checked = False Then
                MessageBox.Show("Select a Destination Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                worksheet.Activate()
                rng.Select()
                Exit Sub
            End If

            If (RadioButton1.Checked = True Or RadioButton4.Checked = True) Then
                If TextBox2.Text = "" Or IsValidExcelCellReference(TextBox2.Text) = False Then
                    MessageBox.Show("Select a Valid Destination Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    worksheet.Activate()
                    rng.Select()
                    Exit Sub
                End If
            End If


            If (RadioButton2.Checked = True Or RadioButton3.Checked = True) Then

                rng2 = worksheet2.Range(rng2.Cells(1, 1), rng2.Cells(rng.Columns.Count, rng.Rows.Count))
                Dim rng2Address As String = rng2.Address

                If CheckBox1.Checked = True Then
                    worksheet.Copy(After:=workbook.Sheets(worksheet.Name))
                End If

                If (Overlap(excelApp, worksheet, worksheet2, rng, rng2)) = False Then
                    If RadioButton3.Checked = True Then
                        If CheckBox2.Checked = True Then
                            For i = 1 To rng.Rows.Count
                                For j = 1 To rng.Columns.Count
                                    rng.Cells(i, j).Copy
                                    rng2.Cells(j, i).PasteSpecial(Excel.XlPasteType.xlPasteAll)
                                    excelApp.CutCopyMode = Excel.XlCutCopyMode.xlCopy
                                Next
                            Next
                        Else
                            For i = 1 To rng.Rows.Count
                                For j = 1 To rng.Columns.Count
                                    rng.Cells(i, j).Copy
                                    rng2.Cells(j, i).PasteSpecial(Excel.XlPasteType.xlPasteValues)
                                    excelApp.CutCopyMode = Excel.XlCutCopyMode.xlCopy
                                Next
                            Next
                        End If
                    ElseIf RadioButton2.Checked = True Then
                        If CheckBox2.Checked = True Then
                            For i = 1 To rng.Rows.Count
                                For j = 1 To rng.Columns.Count
                                    rng2.Cells(j, i).Value = "=" & rng.Cells(i, j).Address(True, True, Excel.XlReferenceStyle.xlA1, True)
                                    rng.Cells(i, j).Copy
                                    rng2.Cells(j, i).PasteSpecial(Excel.XlPasteType.xlPasteFormats)
                                Next
                            Next
                        Else
                            For i = 1 To rng.Rows.Count
                                For j = 1 To rng.Columns.Count
                                    rng2.Cells(j, i).Value = "=" & rng.Cells(i, j).Address(True, True, Excel.XlReferenceStyle.xlA1, True)
                                Next
                            Next
                        End If
                    End If
                    excelApp.CutCopyMode = Excel.XlCutCopyMode.xlCopy
                Else

                    Dim Arr(rng.Rows.Count - 1, rng.Columns.Count - 1) As Object

                    For i = LBound(Arr, 1) To UBound(Arr, 1)
                        For j = LBound(Arr, 2) To UBound(Arr, 2)
                            If RadioButton3.Checked = True Then
                                Arr(i, j) = rng.Cells(i + 1, j + 1)
                            ElseIf RadioButton2.Checked = True Then
                                Arr(i, j) = "=" & rng.Cells(i + 1, j + 1).Address(True, True, Excel.XlReferenceStyle.xlA1, True)
                            End If
                        Next
                    Next


                    For i = 1 To rng.Rows.Count
                        For j = 1 To rng.Columns.Count
                            rng2.Cells(j, i) = Arr(i - 1, j - 1)
                        Next
                    Next

                    If CheckBox2.Checked = True Then

                        Dim FontNames(rng.Rows.Count - 1, rng.Columns.Count - 1) As String
                        Dim FontSizes(rng.Rows.Count - 1, rng.Columns.Count - 1) As Single

                        Dim Bolds(rng.Rows.Count - 1, rng.Columns.Count - 1) As Boolean
                        Dim Italics(rng.Rows.Count - 1, rng.Columns.Count - 1) As Boolean

                        Dim Reds1(rng.Rows.Count - 1, rng.Columns.Count - 1) As Integer
                        Dim Reds2(rng.Rows.Count - 1, rng.Columns.Count - 1) As Integer

                        Dim Greens1(rng.Rows.Count - 1, rng.Columns.Count - 1) As Integer
                        Dim Greens2(rng.Rows.Count - 1, rng.Columns.Count - 1) As Integer

                        Dim Blues1(rng.Rows.Count - 1, rng.Columns.Count - 1) As Integer
                        Dim Blues2(rng.Rows.Count - 1, rng.Columns.Count - 1) As Integer


                        For i = LBound(Arr, 1) To UBound(Arr, 1)
                            For j = LBound(Arr, 2) To UBound(Arr, 2)
                                Dim cell As Excel.Range = rng.Cells(i + 1, j + 1)
                                Dim font As Excel.Font = cell.Font
                                FontNames(i, j) = CStr(cell.Font.Name)
                                FontSizes(i, j) = Convert.ToSingle(font.Size)
                                Bolds(i, j) = cell.Font.Bold
                                Italics(i, j) = cell.Font.Italic
                                Dim colorValue1 As Long = CLng(cell.Interior.Color)
                                Reds1(i, j) = colorValue1 Mod 256
                                Greens1(i, j) = (colorValue1 \ 256) Mod 256
                                Blues1(i, j) = (colorValue1 \ 256 \ 256) Mod 256
                                Dim colorValue2 As Long = CLng(cell.Font.Color)
                                Reds2(i, j) = colorValue2 Mod 256
                                Greens2(i, j) = (colorValue2 \ 256) Mod 256
                                Blues2(i, j) = (colorValue2 \ 256 \ 256) Mod 256
                            Next
                        Next

                        For i = 1 To rng.Rows.Count
                            For j = 1 To rng.Columns.Count
                                With rng2.Cells(j, i).Font
                                    .Name = FontNames(i - 1, j - 1)
                                    .Size = FontSizes(i - 1, j - 1)
                                    .Bold = Bolds(i - 1, j - 1)
                                    .Italic = Italics(i - 1, j - 1)
                                End With

                                Dim red1 As Integer = Reds1(i - 1, j - 1)
                                Dim green1 As Integer = Greens1(i - 1, j - 1)
                                Dim blue1 As Integer = Blues1(i - 1, j - 1)
                                rng2.Cells(j, i).Interior.Color = System.Drawing.Color.FromArgb(red1, green1, blue1)

                                Dim red2 As Integer = Reds2(i - 1, j - 1)
                                Dim green2 As Integer = Greens2(i - 1, j - 1)
                                Dim blue2 As Integer = Blues2(i - 1, j - 1)
                                rng2.Cells(j, i).Font.Color = System.Drawing.Color.FromArgb(red2, green2, blue2)
                            Next
                        Next
                    End If
                End If

                rng2.Select()

                Me.Close()

            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub btn_OK_MouseEnter(sender As Object, e As EventArgs) Handles btn_OK.MouseEnter

        Try

            btn_OK.ForeColor = Color.White
            btn_OK.BackColor = Color.FromArgb(76, 111, 174)

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

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged

        Try
            If RadioButton2.Checked = True Then
                Call Display()
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged

        Try
            Call Display()
        Catch ex As Exception

        End Try

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

        Try
            If TextBox1.Text <> "" And Form4Open = 0 Then
                worksheet = workbook.ActiveSheet
                rng = worksheet.Range(TextBox1.Text)
                rng.Select()
                Call Display()
            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

        Try
            If TextBox2.Text <> "" Then
                worksheet2 = workbook.ActiveSheet
                worksheet2.Range(TextBox2.Text).Select()
            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown

        Try

            If e.KeyCode = Keys.Enter Then
                Call btn_OK_Click(sender, e)
            End If

        Catch ex As Exception

        End Try

        Try

            If e.KeyCode = Keys.Enter Then

                MessageBox.Show("You pressed the Enter key.")
            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub Form3_Loaded(sender As Object, e As EventArgs) Handles Me.Load

        Try

            AddHandler excelApp.SheetSelectionChange, AddressOf excelApp_SheetSelectionChange

            Opened = Opened + 1

        Catch ex As Exception

        End Try

    End Sub

    Private Sub excelApp_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Excel.Range)

        Try

            Dim selectedRange As Excel.Range
            selectedRange = excelApp.Selection

            If FocusedTextBox = 1 Then
                TextBox1.Text = selectedRange.Address
                worksheet = workbook.ActiveSheet
                rng = selectedRange
                TextBox1.Focus()

            ElseIf FocusedTextBox = 2 Then
                TextBox2.Text = selectedRange.Address
                worksheet2 = workbook.ActiveSheet
                rng2 = selectedRange
                TextBox2.Focus()
            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

        Try
            If ComboBox1.SelectedItem = "SOFTEKO" And Opened >= 1 Then

                Dim url As String = "https://www.softeko.co"
                Process.Start(url)

            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub RadioButton4_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton4.CheckedChanged

        Try

            If RadioButton4.Checked = True Then
                Call DestinationChange()
            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub RadioButton1_CheckedChanged_1(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged

        Try
            If RadioButton1.Checked = True Then
                Call DestinationChange()
            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub RadioButton5_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton5.CheckedChanged

        Try

            If RadioButton5.Checked = True Then
                Call DestinationChange()
            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub PictureBox8_GotFocus(sender As Object, e As EventArgs) Handles PictureBox8.GotFocus

        Try
            FocusedTextBox = 1
        Catch ex As Exception

        End Try

    End Sub

    Private Sub TextBox1_GotFocus(sender As Object, e As EventArgs) Handles TextBox1.GotFocus

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

    Private Sub TextBox2_GotFocus(sender As Object, e As EventArgs) Handles TextBox2.GotFocus

        Try
            FocusedTextBox = 2
        Catch ex As Exception

        End Try

    End Sub

    Private Sub PictureBox2_GotFocus(sender As Object, e As EventArgs) Handles PictureBox2.GotFocus

        Try
            FocusedTextBox = 2
        Catch ex As Exception

        End Try

    End Sub

    Private Sub RadioButton3_GotFocus(sender As Object, e As EventArgs) Handles RadioButton3.GotFocus

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

    Private Sub PictureBox5_GotFocus(sender As Object, e As EventArgs) Handles PictureBox5.GotFocus
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

    Private Sub CustomGroupBox3_GotFocus(sender As Object, e As EventArgs) Handles CustomGroupBox3.GotFocus
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

    Private Sub CustomGroupBox5_GotFocus(sender As Object, e As EventArgs) Handles CustomGroupBox5.GotFocus
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

    Private Sub CheckBox2_GotFocus(sender As Object, e As EventArgs) Handles CheckBox2.GotFocus
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

    Private Sub CustomGroupBox2_GotFocus(sender As Object, e As EventArgs) Handles CustomGroupBox2.GotFocus
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

    Private Sub panel1_GotFocus(sender As Object, e As EventArgs) Handles panel1.GotFocus
        Try
            FocusedTextBox = 0
        Catch ex As Exception

        End Try
    End Sub

    Private Sub panel2_GotFocus(sender As Object, e As EventArgs) Handles panel2.GotFocus
        Try
            FocusedTextBox = 0
        Catch ex As Exception

        End Try
    End Sub

    Private Sub PictureBox7_Click(sender As Object, e As EventArgs) Handles PictureBox7.Click

    End Sub

    Private Sub PictureBox7_GotFocus(sender As Object, e As EventArgs) Handles PictureBox7.GotFocus
        Try
            FocusedTextBox = 0
        Catch ex As Exception

        End Try
    End Sub

    Private Sub btn_OK_GotFocus(sender As Object, e As EventArgs) Handles btn_OK.GotFocus
        Try
            FocusedTextBox = 0
        Catch ex As Exception

        End Try
    End Sub

    Private Sub btn_cancel_GotFocus(sender As Object, e As EventArgs) Handles btn_cancel.GotFocus

        Try
            FocusedTextBox = 0
        Catch ex As Exception

        End Try

    End Sub

    Private Sub btn_cancel_Click(sender As Object, e As EventArgs) Handles btn_cancel.Click

        Me.Close()

    End Sub

    Private Sub btn_cancel_MouseEnter(sender As Object, e As EventArgs) Handles btn_cancel.MouseEnter

        Try

            btn_cancel.ForeColor = Color.White
            btn_cancel.BackColor = Color.FromArgb(76, 111, 174)

        Catch ex As Exception

        End Try

    End Sub

    Private Sub btn_cancel_KeyDown(sender As Object, e As KeyEventArgs) Handles btn_cancel.KeyDown

        Try

            If e.KeyCode = Keys.Enter Then
                Call btn_OK_Click(sender, e)
            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub btn_OK_KeyDown(sender As Object, e As KeyEventArgs) Handles btn_OK.KeyDown

        Try

            If e.KeyCode = Keys.Enter Then
                Call btn_OK_Click(sender, e)
            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub CheckBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles CheckBox1.KeyDown

        Try

            If e.KeyCode = Keys.Enter Then
                Call btn_OK_Click(sender, e)
            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub CheckBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles CheckBox2.KeyDown

        Try

            If e.KeyCode = Keys.Enter Then
                Call btn_OK_Click(sender, e)
            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub ComboBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox1.KeyDown

        Try

            If e.KeyCode = Keys.Enter Then
                Call btn_OK_Click(sender, e)
            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub CustomGroupBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles CustomGroupBox1.KeyDown

        Try

            If e.KeyCode = Keys.Enter Then
                Call btn_OK_Click(sender, e)
            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub CustomGroupBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles CustomGroupBox2.KeyDown

        Try

            If e.KeyCode = Keys.Enter Then
                Call btn_OK_Click(sender, e)
            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub CustomGroupBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles CustomGroupBox3.KeyDown

        Try

            If e.KeyCode = Keys.Enter Then
                Call btn_OK_Click(sender, e)
            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub CustomGroupBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles CustomGroupBox4.KeyDown

        Try

            If e.KeyCode = Keys.Enter Then
                Call btn_OK_Click(sender, e)
            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub CustomGroupBox5_KeyDown(sender As Object, e As KeyEventArgs) Handles CustomGroupBox5.KeyDown

        Try

            If e.KeyCode = Keys.Enter Then
                Call btn_OK_Click(sender, e)
            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub CustomGroupBox6_KeyDown(sender As Object, e As KeyEventArgs) Handles CustomGroupBox6.KeyDown

        Try

            If e.KeyCode = Keys.Enter Then
                Call btn_OK_Click(sender, e)
            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub Label1_KeyDown(sender As Object, e As KeyEventArgs) Handles Label1.KeyDown

        Try

            If e.KeyCode = Keys.Enter Then
                Call btn_OK_Click(sender, e)
            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub panel1_KeyDown(sender As Object, e As KeyEventArgs) Handles panel1.KeyDown

        Try

            If e.KeyCode = Keys.Enter Then
                Call btn_OK_Click(sender, e)
            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub panel2_KeyDown(sender As Object, e As KeyEventArgs) Handles panel2.KeyDown

        Try

            If e.KeyCode = Keys.Enter Then
                Call btn_OK_Click(sender, e)
            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub PictureBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles PictureBox1.KeyDown

        Try

            If e.KeyCode = Keys.Enter Then
                Call btn_OK_Click(sender, e)
            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub PictureBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles PictureBox2.KeyDown

        Try

            If e.KeyCode = Keys.Enter Then
                Call btn_OK_Click(sender, e)
            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub PictureBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles PictureBox4.KeyDown

        Try

            If e.KeyCode = Keys.Enter Then
                Call btn_OK_Click(sender, e)
            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub PictureBox5_KeyDown(sender As Object, e As KeyEventArgs) Handles PictureBox5.KeyDown

        Try

            If e.KeyCode = Keys.Enter Then
                Call btn_OK_Click(sender, e)
            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub PictureBox7_KeyDown(sender As Object, e As KeyEventArgs) Handles PictureBox7.KeyDown

        Try

            If e.KeyCode = Keys.Enter Then
                Call btn_OK_Click(sender, e)
            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub PictureBox8_KeyDown(sender As Object, e As KeyEventArgs) Handles PictureBox8.KeyDown

        Try

            If e.KeyCode = Keys.Enter Then
                Call btn_OK_Click(sender, e)
            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub RadioButton1_KeyDown(sender As Object, e As KeyEventArgs) Handles RadioButton1.KeyDown

        Try

            If e.KeyCode = Keys.Enter Then
                Call btn_OK_Click(sender, e)
            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub RadioButton2_KeyDown(sender As Object, e As KeyEventArgs) Handles RadioButton2.KeyDown

        Try

            If e.KeyCode = Keys.Enter Then
                Call btn_OK_Click(sender, e)
            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub RadioButton3_KeyDown(sender As Object, e As KeyEventArgs) Handles RadioButton3.KeyDown

        Try

            If e.KeyCode = Keys.Enter Then
                Call btn_OK_Click(sender, e)
            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub RadioButton4_KeyDown(sender As Object, e As KeyEventArgs) Handles RadioButton4.KeyDown

        Try

            If e.KeyCode = Keys.Enter Then
                Call btn_OK_Click(sender, e)
            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub RadioButton5_KeyDown(sender As Object, e As KeyEventArgs) Handles RadioButton5.KeyDown

        Try

            If e.KeyCode = Keys.Enter Then
                Call btn_OK_Click(sender, e)
            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown

        Try

            If e.KeyCode = Keys.Enter Then
                Call btn_OK_Click(sender, e)
            End If

        Catch ex As Exception

        End Try

    End Sub

End Class