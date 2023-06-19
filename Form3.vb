Imports System.Drawing
Imports System.Windows.Forms
Imports System.Reflection.Emit
Imports Microsoft.Office.Interop.Excel
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Threading
Imports System.Diagnostics
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Button

Public Class Form3
    Dim excelApp As Excel.Application
    Dim workbook As Excel.Workbook
    Dim worksheet As Excel.Worksheet

    Private Sub Display()
        panel1.Controls.Clear()
        panel2.Controls.Clear()

        excelApp = Globals.ThisAddIn.Application
        workbook = excelApp.ActiveWorkbook
        worksheet = workbook.ActiveSheet

        Dim rng As Excel.Range
        rng = worksheet.Range(TextBox1.Text)

        Dim r As Integer
        Dim c As Integer

        r = rng.Rows.Count
        c = rng.Columns.Count

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

        For i = 1 To rng.Rows.Count
            For j = 1 To rng.Columns.Count
                Dim label As New System.Windows.Forms.Label
                label.Text = rng.Cells(i, j).Value
                label.Location = New System.Drawing.Point((j - 1) * width, (i - 1) * height)
                label.Height = height
                label.Width = width
                label.BorderStyle = BorderStyle.FixedSingle
                label.TextAlign = ContentAlignment.MiddleCenter

                If CheckBox2.Checked = True Then

                    Dim cell As Excel.Range = rng.Cells(i, j)
                    Dim font As Excel.Font = cell.Font

                    Dim fontStyle As FontStyle = fontStyle.Regular
                    If cell.Font.Bold Then fontStyle = fontStyle Or fontStyle.Bold
                    If cell.Font.Italic Then fontStyle = fontStyle Or fontStyle.Italic


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

            For i = 1 To rng.Rows.Count
                For j = 1 To rng.Columns.Count
                    Dim label As New System.Windows.Forms.Label
                    label.Text = rng.Cells(i, j).Value
                    label.Location = New System.Drawing.Point((i - 1) * width, (j - 1) * height)
                    label.Height = height
                    label.Width = width
                    label.BorderStyle = BorderStyle.FixedSingle
                    label.TextAlign = ContentAlignment.MiddleCenter

                    If CheckBox2.Checked = True Then
                        Dim cell As Excel.Range = rng.Cells(i, j)
                        Dim font As Excel.Font = cell.Font

                        Dim fontStyle As FontStyle = fontStyle.Regular
                        If cell.Font.Bold Then fontStyle = fontStyle Or fontStyle.Bold
                        If cell.Font.Italic Then fontStyle = fontStyle Or fontStyle.Italic

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


    End Sub


    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        excelApp = Globals.ThisAddIn.Application
        workbook = excelApp.ActiveWorkbook
        'formLoaded = True
        Dim myPanel As New Panel()
        ComboBox1.Text = "Softeko"
        Me.TextBox2.Text = MyVar

        ComboBox2.Items.Clear()

        For Each sheet As Excel.Worksheet In workbook.Sheets
            ComboBox2.Items.Add(sheet.Name)
        Next

    End Sub


    Private Sub btn_OK_MouseEnter(sender As Object, e As EventArgs) Handles btn_OK.MouseEnter
        btn_OK.ForeColor = Color.White
        btn_OK.BackColor = Color.FromArgb(76, 111, 174)
    End Sub

    Private Sub btn_OK_MouseLeave(sender As Object, e As EventArgs) Handles btn_OK.MouseLeave
        btn_OK.ForeColor = Color.FromArgb(70, 70, 70)
        btn_OK.BackColor = Color.White
    End Sub

    Private Sub btn_cancel_MouseLeave(sender As Object, e As EventArgs) Handles btn_cancel.MouseLeave
        btn_cancel.ForeColor = Color.FromArgb(70, 70, 70)
        btn_cancel.BackColor = Color.White
    End Sub

    Private Sub btn_cancel_MouseEnter(sender As Object, e As EventArgs) Handles btn_cancel.MouseEnter
        btn_cancel.ForeColor = Color.White
        btn_cancel.BackColor = Color.FromArgb(76, 111, 174)
    End Sub

    Private Sub RadioButton5_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton5.CheckedChanged
        If RadioButton5.Checked = True Then

            TextBox2.Location = New System.Drawing.Point(121, 47)
            PictureBox2.Location = New System.Drawing.Point(226, 47)

            Dim form As New Form4()
            form.Show()
            Me.Hide()

        End If
    End Sub

    Private Sub PictureBox8_Click(sender As Object, e As EventArgs) Handles PictureBox8.Click
        Me.Visible = False
        'TextBox1.Text = selectedRange.Address

        Dim selectedRange As Excel.Range = excelApp.InputBox("Select a range", Type:=8)
        selectedRange.Select()
        Me.Visible = True

        ' Put the selected range's address into the TextBox.
        TextBox1.Text = selectedRange.Address
    End Sub

    Private Sub PictureBox4_Click(sender As Object, e As EventArgs) Handles PictureBox4.Click
        excelApp = Globals.ThisAddIn.Application
        Me.Hide()
        Dim userInput As Excel.Range = excelApp.InputBox("Select a range", Type:=8)
        Me.Show()

        Dim rng As Microsoft.Office.Interop.Excel.Range = userInput

        ' Select the range
        rng.Select()

        ' Expand the selection downwards
        rng = excelApp.Range(rng, rng.End(Microsoft.Office.Interop.Excel.XlDirection.xlDown))
        'rng = Range(rng, rng.End(Microsoft.Office.Interop.Excel.XlDirection.xlDown))
        rng.Select()

        ' Expand the selection to the right
        rng = excelApp.Range(rng, rng.End(Microsoft.Office.Interop.Excel.XlDirection.xlToRight))
        rng.Select()

        ' Get the address of the selected range
        Dim selectedRangeAddress As String = excelApp.Selection.Address
        Me.TextBox1.Text = selectedRangeAddress
        Me.TextBox1.Focus()
    End Sub

    Private Sub RadioButton4_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton4.CheckedChanged
        If RadioButton4.Checked = True Then


            excelApp = Globals.ThisAddIn.Application

            TextBox2.Location = New System.Drawing.Point(121, 27)
            PictureBox2.Location = New System.Drawing.Point(226, 27)


            Dim copiedWorksheet As Excel.Worksheet

            Dim activeSheet As Excel.Worksheet = CType(excelApp.ActiveWorkbook.ActiveSheet, Excel.Worksheet)

            ' worksheet = CType(workbook.ActiveSheet, Excel.Worksheet)

            ' Copy the active sheet. In this case, it's copied to the end.
            'activeSheet.Copy(After:=activeSheet)
            copiedWorksheet = excelApp.ActiveWorkbook.Worksheets.Add(After:=activeSheet)

            ' Get the newly copied worksheet (which is the last one) and rename it
            copiedWorksheet = CType(excelApp.ActiveWorkbook.Sheets(excelApp.ActiveWorkbook.Sheets.Count), Excel.Worksheet)
            'copiedWorksheet.Name = "CopiedSheet" ' Your desired name

            Me.Visible = False

            Dim selectedRange As Excel.Range = excelApp.InputBox("Select a range", Type:=8)
            selectedRange.Select()
            Me.Visible = True

            ' Put the selected range's address into the TextBox.
            TextBox2.Text = selectedRange.Address
        End If
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        TextBox2.Location = New System.Drawing.Point(121, 7)
        PictureBox2.Location = New System.Drawing.Point(226, 7)

    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click
        Me.Visible = False

        Dim selectedRange As Excel.Range = excelApp.InputBox("Select a range", Type:=8)
        'selectedRange.Select()
        Me.Visible = True

        ' Put the selected range's address into the TextBox.
        TextBox2.Text = selectedRange.Address
    End Sub

    Private Sub panel1_Paint(sender As Object, e As PaintEventArgs) Handles panel1.Paint

    End Sub

    Private Sub btn_OK_Click(sender As Object, e As EventArgs) Handles btn_OK.Click

        If (RadioButton2.Checked = True Or RadioButton3.Checked = True) And (RadioButton1.Checked = True Or RadioButton4.Checked = True Or RadioButton5.Checked = True) Then

            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            worksheet = workbook.ActiveSheet

            Dim worksheet2 As Excel.Worksheet

            Dim rng As Excel.Range
            rng = worksheet.Range(TextBox1.Text)

            Dim rng2 As Excel.Range

            Dim Arr(rng.Rows.Count - 1, rng.Columns.Count - 1) As Object

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

            If RadioButton1.Checked = True Then
                rng2 = worksheet.Range(TextBox2.Text)
            ElseIf RadioButton4.Checked = True Then
                worksheet2 = workbook.Worksheets.Add
                rng2 = worksheet2.Range(TextBox1.Text)
            End If


            For i = LBound(Arr, 1) To UBound(Arr, 1)
                For j = LBound(Arr, 2) To UBound(Arr, 2)
                    Arr(i, j) = rng.Cells(i + 1, j + 1)
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
            'Test Comment
            For i = 1 To rng.Rows.Count
                For j = 1 To rng.Columns.Count
                    rng2.Cells(j, i) = Arr(i - 1, j - 1)

                    If CheckBox2.Checked = True Then

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

                    End If
                Next
            Next
        End If

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
            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            worksheet = workbook.ActiveSheet
            worksheet.Range(TextBox1.Text).Select()
            Call Display()
        Catch ex As Exception

        End Try

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged

        Try

            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            worksheet = workbook.Sheets(ComboBox2.SelectedItem)
            worksheet.Activate()

            Dim userInput As Excel.Range = excelApp.InputBox("Select a cell", Type:=8)

            TextBox2.Text = userInput.Address
        Catch ex As Exception

        End Try
    End Sub

End Class