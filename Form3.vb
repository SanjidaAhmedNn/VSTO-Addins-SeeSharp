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
    'Private form As Form4

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


    End Sub

    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        excelApp = Globals.ThisAddIn.Application
        'formLoaded = True
        Dim myPanel As New Panel()
        ComboBox1.Text = "Softeko"
        Me.TextBox2.Text = MyVar
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

    Private Sub PictureBox8_Click(sender As Object, e As EventArgs) Handles PictureBox8.Click
        Try

            excelApp = Globals.ThisAddIn.Application
            Dim selectedRange As Excel.Range
            selectedRange = excelApp.InputBox("Select a range", Type:=8)
            TextBox1.Text = selectedRange.Address

        Catch ex As Exception

        End Try
    End Sub

    Private Sub PictureBox4_Click(sender As Object, e As EventArgs) Handles PictureBox4.Click
        Try

            excelApp = Globals.ThisAddIn.Application
            Dim selectedRange As Excel.Range
            selectedRange = excelApp.InputBox("Select a range", Type:=8)
            TextBox1.Text = selectedRange.Address

        Catch ex As Exception

        End Try

    End Sub


    Private Sub btn_OK_MouseHover(sender As Object, e As EventArgs) Handles btn_OK.MouseHover

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

        excelApp = Globals.ThisAddIn.Application
        workbook = excelApp.ActiveWorkbook
        worksheet = workbook.ActiveSheet

        Try
            ' Code that may cause an exception.
            Dim rng As Excel.Range
            rng = worksheet.Range(TextBox1.Text)
            rng.Select()
            Call Display()
        Catch ex As Exception
            ' Do nothing.
        End Try


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
End Class