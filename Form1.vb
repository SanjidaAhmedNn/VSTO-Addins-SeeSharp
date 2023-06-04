Imports System.Drawing
Imports System.Reflection.Emit
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Excel
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Threading
Imports System.Diagnostics
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Button

Public Class Form1
    Dim excelApp As Excel.Application
    Dim workbook As Excel.Workbook
    Dim worksheet As Excel.Worksheet

    Dim myPanel As New Panel()
    ' Private newForm As New NewForm()
    Private isFormOpen As Boolean = False
    Private formLoaded As Boolean = False

    ' Set panel properties


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        excelApp = Globals.ThisAddIn.Application
        formLoaded = True
        Dim myPanel As New Panel()
        ComboBox1.Text = "Softeko"

    End Sub

    Private Sub btn_OK_Click(sender As Object, e As EventArgs) Handles btn_OK.Click
        Dim selectedRange As Range = excelApp.Selection

        Dim rowCount As Integer = selectedRange.Rows.Count
        Dim columnCount As Integer = selectedRange.Columns.Count
        Dim temp As Object
        For i = 1 To rowCount
            For j = 1 To columnCount / 2
                temp = selectedRange(i, j).Value
                selectedRange(i, j).Value = selectedRange(i, columnCount - j + 1).Value
                selectedRange(i, columnCount - j + 1).Value = temp
            Next
        Next

        Me.Close()
        Me.Dispose()
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

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        Panel1.Controls.Clear()
        Panel2.Controls.Clear()

        Dim selectedRange As Range = excelApp.Selection
        Dim rowCount As Integer = selectedRange.Rows.Count
        Dim columnCount As Integer = selectedRange.Columns.Count



        Dim w As Integer = selectedRange(1, 1).width
        Dim l As Integer = selectedRange(1, 1).height

        If w * columnCount < 331 Then
            w = 331 / columnCount
        End If

        If l * rowCount < 154 Then
            l = 154 / rowCount
        End If

        If w * columnCount <= 331 And l * rowCount <= 154 Then
            Panel1.AutoScroll = False
            Panel2.AutoScroll = False
        End If
        'Dim w As Integer = 60  ' Width of each label.
        'Dim l As Integer = 20  ' Height of each label.


        For i = 1 To rowCount
            For j = 1 To columnCount
                Dim lbl As New System.Windows.Forms.Label()
                With lbl
                    .Top = ((i - 1) * l)
                    .Left = (((j - 1) * w))
                    .Text = selectedRange.Cells(i, j).value
                    .Width = w
                    .Height = l
                    'ForeColor = selectedRange(i, j).Font.Color
                    .ForeColor = ColorTranslator.FromOle(CLng(selectedRange.Cells(i, j).Font.Color))
                    .BackColor = ColorTranslator.FromOle(CLng(selectedRange.Cells(i, j).interior.Color))
                    .BorderStyle = BorderStyle.FixedSingle
                    '.Font = CType(selectedRange.Cells(i, j), Excel.Range).Font.Name
                    '.Font = New System.Drawing.Font(lbl.Font.FontFamily, CSng(selectedRange.Cells(i, j).Font.Size))
                    .Font = New System.Drawing.Font(selectedRange.Cells(i, j).Font.Name.ToString(), CSng(selectedRange.Cells(i, j).Font.Size))
                    .TextAlign = ContentAlignment.MiddleCenter
                End With
                Panel2.Controls.Add(lbl)

            Next
        Next

        For i = 1 To rowCount
            For j = 1 To columnCount
                Dim lbl As New System.Windows.Forms.Label()
                With lbl
                    .Top = ((i - 1) * l)
                    .Left = (((j - 1) * w))
                    .Text = selectedRange.Cells(i, columnCount - j + 1).value
                    .Width = w
                    .Height = l
                    'ForeColor = selectedRange(i, j).Font.Color
                    .ForeColor = ColorTranslator.FromOle(CLng(selectedRange.Cells(i, columnCount - j + 1).Font.Color))
                    .BackColor = ColorTranslator.FromOle(CLng(selectedRange.Cells(i, columnCount - j + 1).interior.Color))
                    .BorderStyle = BorderStyle.FixedSingle
                    .Font = New System.Drawing.Font(selectedRange.Cells(i, j).Font.Name.ToString(), CSng(selectedRange.Cells(i, j).Font.Size))
                    .TextAlign = ContentAlignment.MiddleCenter

                End With
                Panel1.Controls.Add(lbl)

            Next
        Next
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        excelApp = Globals.ThisAddIn.Application
        Dim copiedWorksheet As Excel.Worksheet

        Dim activeSheet As Excel.Worksheet = CType(excelApp.ActiveWorkbook.ActiveSheet, Excel.Worksheet)

        ' worksheet = CType(workbook.ActiveSheet, Excel.Worksheet)

        ' Copy the active sheet. In this case, it's copied to the end.
        activeSheet.Copy(After:=activeSheet)

        ' Get the newly copied worksheet (which is the last one) and rename it
        copiedWorksheet = CType(excelApp.ActiveWorkbook.Sheets(excelApp.ActiveWorkbook.Sheets.Count), Excel.Worksheet)
        copiedWorksheet.Name = "CopiedSheet" ' Your desired name
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If formLoaded Then
            Dim selectedOption As String = ComboBox1.SelectedItem.ToString()

            If selectedOption = "SOFTEKO" Then
                Process.Start("https://www.softeko.co/")

            ElseIf selectedOption = "About Us" Then
                Process.Start("https://www.softeko.co/")

            ElseIf selectedOption = "Help" Then
                Process.Start("https://www.softeko.co/")
            ElseIf selectedOption = "Feedback" Then
                Process.Start("https://www.softeko.co/")
                ' Add more ElseIf statements for more options if necessary
            End If
        End If
    End Sub

    Private Sub btn_cancel_Click(sender As Object, e As EventArgs) Handles btn_cancel.Click
        Me.Close()
        Me.Dispose()
    End Sub

    Private Sub panel1_Paint(sender As Object, e As PaintEventArgs) Handles panel1.Paint

    End Sub
End Class


