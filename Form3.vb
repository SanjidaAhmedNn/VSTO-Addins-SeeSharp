Imports System.Drawing
Imports System.Windows.Forms
Imports System.Reflection.Emit
Imports Microsoft.Office.Interop.Excel
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Threading
Imports System.Diagnostics
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Button
Imports System.Windows.Forms.VisualStyles.VisualStyleElement

Public Class Form3
    Dim excelApp As Excel.Application
    Dim workbook As Excel.Workbook
    Dim workbook2 As Excel.Workbook
    Dim worksheet As Excel.Worksheet
    Dim worksheet1 As Excel.Worksheet
    Dim worksheet2 As Excel.Worksheet
    Private Function Overlap(rng1 As Excel.Range, rng2 As Excel.Range, Sheet1 As Excel.Worksheet, Sheet2 As Excel.Worksheet)

        Dim Result As Boolean
        If Sheet1.Name <> Sheet2.Name Then
            Result = False
        Else
            Dim X1 As Boolean
            Dim X2 As Boolean
            X1 = (rng2.Cells(1, 1).Row >= rng1.Cells(1, 1).Row) And (rng2.Cells(1, 1).Row <= rng1.Cells(rng1.Rows.Count, rng1.Columns.Count).Row)
            X2 = (rng2.Cells(1, 1).Column >= rng1.Cells(1, 1).Column And rng2.Cells(1, 1).Column <= rng1.Cells(rng1.Rows.Count, rng1.Columns.Count).Column)

            If X1 And X2 Then
                Result = True
            Else
                Result = False
            End If

        End If

        Overlap = Result

    End Function

    Private Sub Display()

        Try

            panel1.Controls.Clear()
            panel2.Controls.Clear()

            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            worksheet = workbook.ActiveSheet

            Dim rng As Excel.Range
            rng = worksheet.Range(TextBox1.Text)

            If rng.Rows.Count > 50 Then
                rng = worksheet.Range(rng.Cells(1, 1), rng.Cells(50, rng.Columns.Count))
            End If

            If rng.Columns.Count > 50 Then
                rng = worksheet.Range(rng.Cells(1, 1), rng.Cells(rng.Rows.Count, 50))
            End If

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

            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook

            Dim worksheet2 As Excel.Worksheet

            Dim userInput As Excel.Range = excelApp.InputBox("Select a range", Type:=8)
            Dim rng As Microsoft.Office.Interop.Excel.Range = userInput

            Try
                Dim sheetName As String
                sheetName = Split(rng.Address(True, True, Excel.XlReferenceStyle.xlA1, True), "]")(1)
                sheetName = Split(sheetName, "!")(0)
                worksheet2 = workbook.Worksheets(sheetName)
                worksheet2.Activate()
            Catch ex As Exception

            End Try

            rng.Select()

            TextBox1.Text = rng.Address
            TextBox1.Focus()

        Catch ex As Exception

        End Try

    End Sub

    Private Sub PictureBox4_Click(sender As Object, e As EventArgs) Handles PictureBox4.Click

        Try

            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook

            Dim worksheet2 As Excel.Worksheet

            Dim userInput As Excel.Range = excelApp.InputBox("Select a range", Type:=8)
            Dim rng As Microsoft.Office.Interop.Excel.Range = userInput

            Try
                Dim sheetName As String
                sheetName = Split(rng.Address(True, True, Excel.XlReferenceStyle.xlA1, True), "]")(1)
                sheetName = Split(sheetName, "!")(0)
                worksheet2 = workbook.Worksheets(sheetName)
                worksheet2.Activate()
                '   ComboBox2.SelectedItem = sheetName
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


    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs)

        Try

            TextBox2.Location = New System.Drawing.Point(121, 7)
            PictureBox2.Location = New System.Drawing.Point(226, 7)

        Catch ex As Exception

        End Try

    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click

        Try

            Dim selectedRange As Excel.Range = excelApp.InputBox("Select a range", Type:=8)
            TextBox2.Text = selectedRange.Address
            TextBox2.Focus()

        Catch ex As Exception

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

            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            worksheet1 = workbook.ActiveSheet
            worksheet1.Range(TextBox1.Text).Select()
            Call Display()

        Catch ex As Exception

        End Try

    End Sub

    ' Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged

    'Try

    '      excelApp = Globals.ThisAddIn.Application
    '     workbook = excelApp.ActiveWorkbook
    'If ComboBox2.SelectedItem = "Add New" Then
    '             worksheet2 = workbook.Sheets.Add(After:=workbook.Sheets(workbook.Sheets.Count))
    '            ComboBox2.Items(ComboBox2.FindStringExact("Add New")) = worksheet2.Name
    '           ComboBox2.Items.Add("Add New")
    ' Else
    '            worksheet2 = workbook.Sheets(ComboBox2.SelectedItem)
    '           worksheet2.Activate()
    'End If
    '        ComboBox2.Focus()
    'Catch ex As Exception

    'End Try
    'End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

        Try
            excelApp = Globals.ThisAddIn.Application
            workbook = excelApp.ActiveWorkbook
            worksheet = workbook.ActiveSheet
            worksheet.Range(TextBox2.Text).Select()

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

    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load

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