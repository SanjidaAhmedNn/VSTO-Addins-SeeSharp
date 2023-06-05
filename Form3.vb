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
            copiedWorksheet.Name = "CopiedSheet" ' Your desired name

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
        selectedRange.Select()
        Me.Visible = True

        ' Put the selected range's address into the TextBox.
        TextBox2.Text = selectedRange.Address
    End Sub


End Class