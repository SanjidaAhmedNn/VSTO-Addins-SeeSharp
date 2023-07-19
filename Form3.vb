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

Public Class Form3
    Public WithEvents excelApp As Excel.Application
    Public workbook As Excel.Workbook
    Public workbook2 As Excel.Workbook
    Public worksheet As Excel.Worksheet
    Public worksheet1 As Excel.Worksheet
    Public worksheet2 As Excel.Worksheet
    Public rng As Excel.Range
    Public rng2 As Excel.Range
    Public FocuesdTextBox As Integer
    Public Opened As Integer
    Public GB5 As Integer
    Public GB6 As Integer
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

        If RadioButton1.Checked = True Then
            TextBox2.Enabled = True
            PictureBox2.Enabled = True
        Else
            TextBox2.Clear()
            TextBox2.Enabled = False
            PictureBox2.Enabled = False
        End If

        If RadioButton4.Checked = True Then

            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            Dim ws As Excel.Worksheet = CType(workbook.Worksheets.Add(), Excel.Worksheet)
            ws.Name = "Transpose Sheet"
            worksheet2 = ws
            rng2 = worksheet2.Range("A1")

        End If

        If RadioButton5.Checked = True Then
            Me.Visible = False
            Dim MyForm4 As New Form4
            MyForm4.excelApp = Me.excelApp
            MyForm4.workbook = Me.workbook
            MyForm4.worksheet = Me.worksheet
            MyForm4.rng = Me.rng
            MyForm4.Opened = Me.Opened
            MyForm4.FocuesdTextBox = Me.FocuesdTextBox
            If Me.RadioButton3.Checked = True Then
                MyForm4.GB6 = 3
            ElseIf Me.RadioButton2.Checked = True Then
                MyForm4.GB6 = 2
            End If
            MyForm4.GB5 = 3
            MyForm4.Show()

        End If

    End Sub


    ' Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    'Try

    '       excelApp = Globals.ThisAddIn.Application
    '       workbook = excelApp.ActiveWorkbook

    'Dim myPanel As New Panel()
    '       ComboBox1.Text = "Softeko"
    'Me.TextBox2.Text = MyVar

    'Me.KeyPreview = True

    '     ComboBox2.Items.Clear()

    'For Each sheet As Excel.Worksheet In workbook.Sheets
    '           ComboBox2.Items.Add(sheet.Name)
    'Next
    '       ComboBox2.Items.Add("Add New")
    '       ComboBox2.SelectedItem = workbook.ActiveSheet.Name

    '       ComboBox3.Items.Clear()
    '      ComboBox3.Items.Add("This Workbook")
    '      ComboBox3.Items.Add("Existing Workbook")
    '       ComboBox3.Items.Add("New Workbook")
    '       ComboBox3.SelectedItem = "This Workbook"

    'Catch ex As Exception

    'End Try
    ' End Sub


    Private Sub btn_OK_MouseEnter(sender As Object, e As EventArgs) Handles btn_OK.MouseEnter

        Try

            btn_OK.ForeColor = Color.White
            btn_OK.BackColor = Color.FromArgb(76, 111, 174)

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

    Private Sub btn_cancel_MouseEnter(sender As Object, e As EventArgs) Handles btn_cancel.MouseEnter

        Try

            btn_cancel.ForeColor = Color.White
            btn_cancel.BackColor = Color.FromArgb(76, 111, 174)

        Catch ex As Exception

        End Try

    End Sub

    Private Sub PictureBox8_Click(sender As Object, e As EventArgs) Handles PictureBox8.Click

        Try
            FocuesdTextBox = 1
            Me.Hide()

            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook

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
            FocuesdTextBox = 1
            Me.Hide()

            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook

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
            FocuesdTextBox = 2
            Me.Hide()

            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook

            Dim userInput As Excel.Range = excelApp.InputBox("Select a range", Type:=8)
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

    Private Sub panel1_Paint(sender As Object, e As PaintEventArgs) Handles panel1.Paint

    End Sub

    Private Sub btn_OK_Click(sender As Object, e As EventArgs) Handles btn_OK.Click

        Try

            If (RadioButton2.Checked = True Or RadioButton3.Checked = True) Then


                Dim rng As Excel.Range
                rng = worksheet1.Range(TextBox1.Text)

                Dim rng2 As Excel.Range
                rng2 = worksheet2.Range(TextBox2.Text)

                If CheckBox1.Checked = True Then
                    worksheet1.Copy(After:=workbook.Sheets(worksheet1.Name))
                End If

                worksheet2.Activate()

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

        Catch ex As Exception

        End Try

    End Sub

    Private Sub btn_OK_MouseHover(sender As Object, e As EventArgs) Handles btn_OK.MouseHover

    End Sub

    Private Sub CustomGroupBox2_Enter(sender As Object, e As EventArgs) Handles CustomGroupBox2.Enter

    End Sub

    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged

        Try
            Call Display()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged

        Try
            Call Display()
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

            If TextBox1.Text <> "" Then
                excelApp = Globals.ThisAddIn.Application
                workbook = excelApp.ActiveWorkbook
                worksheet1 = workbook.ActiveSheet
                rng = worksheet1.Range(TextBox1.Text)
                rng.Select()
                Call Display()
            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

        Try
            If TextBox2.Text <> "" Then
                excelApp = Globals.ThisAddIn.Application
                workbook = excelApp.ActiveWorkbook
                worksheet = workbook.ActiveSheet
                worksheet.Range(TextBox2.Text).Select()
            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown

        Try

            If e.KeyCode = Keys.Enter Then

                MessageBox.Show("You pressed the Enter key.")
            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub Form3_Activated(sender As Object, e As EventArgs) Handles Me.Activated

        Try

            excelApp = Globals.ThisAddIn.Application

            AddHandler excelApp.SheetSelectionChange, AddressOf excelApp_SheetSelectionChange

            Opened = Opened + 1

            Call DestinationChange()

        Catch ex As Exception

        End Try
    End Sub

    Private Sub excelApp_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Excel.Range)

        Try

            excelApp = Globals.ThisAddIn.Application
            Dim selectedRange As Excel.Range
            selectedRange = excelApp.Selection
            If FocuesdTextBox = 1 Then
                TextBox1.Text = selectedRange.Address
                worksheet = workbook.ActiveSheet
                rng = selectedRange
                TextBox1.Focus()
            ElseIf FocuesdTextBox = 2 Then
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

            Call DestinationChange()
        Catch ex As Exception

        End Try

    End Sub

    Private Sub RadioButton1_CheckedChanged_1(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged

        Try

            Call DestinationChange()

        Catch ex As Exception

        End Try

    End Sub

    Private Sub RadioButton5_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton5.CheckedChanged

        Try

            Call DestinationChange()

        Catch ex As Exception

        End Try

    End Sub

    ' Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged

    'Try

    'If ComboBox3.SelectedItem = "Existing Workbook" Then
    'Dim openFileDialog1 As New OpenFileDialog()

    '         openFileDialog1.Filter = "All Files (*.*)|*.*"
    '        openFileDialog1.FilterIndex = 1

    'Dim userClickedOK As DialogResult = openFileDialog1.ShowDialog()

    'If userClickedOK = DialogResult.OK Then

    'Dim filePath As String = openFileDialog1.FileName
    'Dim workbookName As String = System.IO.Path.GetFileName(filePath)

    '              workbook = excelApp.Workbooks.Open(filePath)

    '             excelApp.Visible = True

    '             ComboBox3.Focus()
    ' Call Form3_Load(sender, e)
    'End If

    'End If

    'Catch ex As Exception

    'End Try
    ' End Sub

End Class