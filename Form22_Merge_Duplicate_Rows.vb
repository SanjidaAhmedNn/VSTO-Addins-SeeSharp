Imports System.Drawing
Imports System.Reflection.Emit
Imports System.Text.RegularExpressions
Imports System.Windows.Forms

Public Class Form22_Merge_Duplicate_Rows

    Dim WithEvents excelApp As Excel.Application
    Dim workBook As Excel.Workbook
    Dim workSheet As Excel.Worksheet
    Dim workSheet2 As Excel.Worksheet
    Dim rng As Excel.Range
    Dim rng2 As Excel.Range
    Dim selectedRange As Excel.Range

    Dim opened As Integer
    Dim FocusedTextBox As Integer

    Dim variables As New Dictionary(Of String, System.Windows.Forms.Label)
    Dim labels As New List(Of System.Windows.Forms.Label)()
    Dim labels2 As New List(Of System.Windows.Forms.Label)()
    Dim clickedLabelNumber As Integer
    Dim EnteredLabelNumber As Integer


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

    Private Sub Setup()

        excelApp = Globals.ThisAddIn.Application
        workBook = excelApp.ActiveWorkbook
        workSheet = workBook.ActiveSheet
        rng = workSheet.Range(TextBox1.Text)

        Dim height As Single = Label3.Height

        Dim i As Integer


        For i = 1 To rng.Columns.Count

            Dim lbl As New System.Windows.Forms.Label()
            lbl.Text = rng.Cells(1, i).Value
            lbl.Location = New System.Drawing.Point(1, (i - 1) * height)
            lbl.Height = height
            lbl.Width = Label2.Width - 3.5
            lbl.Font = New Font("Segoe UI", 9.75F)
            lbl.TextAlign = ContentAlignment.MiddleCenter
            lbl.TextAlign = ContentAlignment.MiddleLeft
            CustomGroupBox7.Controls.Add(lbl)
            labels.Add(lbl)

            AddHandler lbl.Click, AddressOf Me.lbl_Click
            AddHandler lbl.MouseEnter, AddressOf Me.lbl_MouseEnter

            Dim lbl2 As New System.Windows.Forms.Label
            lbl2.Text = rng.Cells(2, i).Value
            lbl2.Location = New System.Drawing.Point(Label2.Width - 3.5, (i - 1) * height)
            lbl2.Height = height
            lbl2.Width = Label4.Width - 3.5
            lbl2.Font = New Font("Segoe UI", 9.75F)
            lbl2.TextAlign = ContentAlignment.MiddleCenter
            lbl2.TextAlign = ContentAlignment.MiddleLeft
            CustomGroupBox7.Controls.Add(lbl2)
            labels2.Add(lbl2)

        Next


    End Sub

    Private Sub lbl_Click(sender As Object, e As EventArgs)

        If clickedLabelNumber <> -1 Then
            labels(clickedLabelNumber).BackColor = Color.FromArgb(255, 255, 255)
            labels(clickedLabelNumber).ForeColor = Color.FromArgb(70, 70, 70)
            labels2(clickedLabelNumber).BackColor = Color.FromArgb(255, 255, 255)
            labels2(clickedLabelNumber).ForeColor = Color.FromArgb(70, 70, 70)
        End If

        Dim clickedLabel As System.Windows.Forms.Label
        clickedLabel = DirectCast(sender, System.Windows.Forms.Label)

        clickedLabelNumber = labels.IndexOf(clickedLabel)

        clickedLabel.BackColor = Color.FromArgb(0, 120, 215)
        clickedLabel.ForeColor = Color.FromArgb(255, 255, 255)
        labels2(clickedLabelNumber).BackColor = Color.FromArgb(0, 120, 215)
        labels2(clickedLabelNumber).ForeColor = Color.FromArgb(255, 255, 255)


    End Sub
    Private Sub lbl_MouseEnter(sender As Object, e As EventArgs)

        If EnteredLabelNumber <> -1 And clickedLabelNumber <> EnteredLabelNumber Then
            labels(EnteredLabelNumber).BackColor = Color.FromArgb(255, 255, 255)
            labels(EnteredLabelNumber).ForeColor = Color.FromArgb(70, 70, 70)
            labels2(EnteredLabelNumber).BackColor = Color.FromArgb(255, 255, 255)
            labels2(EnteredLabelNumber).ForeColor = Color.FromArgb(70, 70, 70)
        End If

        Dim clickedLabel As System.Windows.Forms.Label
        clickedLabel = DirectCast(sender, System.Windows.Forms.Label)

        EnteredLabelNumber = labels.IndexOf(clickedLabel)

        clickedLabel.BackColor = Color.FromArgb(229, 243, 255)
        clickedLabel.ForeColor = Color.FromArgb(70, 70, 70)
        labels2(EnteredLabelNumber).BackColor = Color.FromArgb(229, 243, 255)
        labels2(EnteredLabelNumber).ForeColor = Color.FromArgb(70, 70, 70)

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

        If TextBox1.Text <> "" And IsValidExcelCellReference(TextBox1.Text) = True Then
            Call Setup()
        End If

    End Sub

    Private Sub Form22_Merge_Duplicate_Rows_Load(sender As Object, e As EventArgs) Handles Me.Load

        clickedLabelNumber = -1
        EnteredLabelNumber = -1

    End Sub
End Class
