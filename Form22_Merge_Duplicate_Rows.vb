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
    Dim labels3 As New List(Of System.Windows.Forms.Label)()
    Dim comboBoxes As New List(Of System.Windows.Forms.ComboBox)()
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

        CustomGroupBox7.Controls.Clear()

        excelApp = Globals.ThisAddIn.Application
        workBook = excelApp.ActiveWorkbook
        workSheet = workBook.ActiveSheet
        rng = workSheet.Range(TextBox1.Text)

        If CheckBox5.Checked = True Then
            rng = workSheet.Range(rng.Cells(2, 1), rng.Cells(rng.Rows.Count, rng.Columns.Count))
        End If

        rng.Select()

        Dim height As Single = Label3.Height

        Dim i As Integer

        For i = 1 To rng.Columns.Count

            Dim lbl As New System.Windows.Forms.Label()
            If CheckBox5.Checked = True Then
                lbl.Text = rng.Cells(0, i).Value
            Else
                Dim columnLetter As String = Split(rng.Cells(1, i).Address(True, True), "$")(1)
                lbl.Text = "Column " & columnLetter
            End If
            lbl.Location = New System.Drawing.Point(1, (i - 1) * height)
            lbl.Height = height
            lbl.Width = Label2.Width - 4
            lbl.Font = New Font("Segoe UI", 9.75F)
            lbl.TextAlign = ContentAlignment.MiddleCenter
            lbl.TextAlign = ContentAlignment.MiddleLeft
            lbl.BorderStyle = BorderStyle.FixedSingle
            CustomGroupBox7.Controls.Add(lbl)
            labels.Add(lbl)

            AddHandler lbl.Click, AddressOf Me.lbl_Click
            AddHandler lbl.MouseEnter, AddressOf Me.lbl_MouseEnter
            AddHandler lbl.Paint, AddressOf lbl_Paint

            Dim lbl2 As New System.Windows.Forms.Label
            lbl2.Text = rng.Cells(1, i).Value
            lbl2.Location = New System.Drawing.Point(Label2.Width - 4, (i - 1) * height)
            lbl2.Height = height
            lbl2.Width = Label4.Width - 4.25
            lbl2.Font = New Font("Segoe UI", 9.75F)
            lbl2.TextAlign = ContentAlignment.MiddleCenter
            lbl2.TextAlign = ContentAlignment.MiddleLeft
            lbl2.BorderStyle = BorderStyle.FixedSingle
            CustomGroupBox7.Controls.Add(lbl2)
            labels2.Add(lbl2)

            AddHandler lbl2.Click, AddressOf Me.lbl2_Click
            AddHandler lbl2.MouseEnter, AddressOf Me.lbl2_MouseEnter
            AddHandler lbl2.Paint, AddressOf lbl2_Paint

            Dim lbl3 As New System.Windows.Forms.Label
            lbl3.Text = ""
            lbl3.Location = New System.Drawing.Point((Label2.Width + Label4.Width) - 8.75, (i - 1) * height)
            lbl3.Height = height
            lbl3.Width = Label5.Width
            lbl3.Font = New Font("Segoe UI", 9.75F)
            lbl3.TextAlign = ContentAlignment.MiddleCenter
            lbl3.TextAlign = ContentAlignment.MiddleLeft
            lbl3.BorderStyle = BorderStyle.FixedSingle
            CustomGroupBox7.Controls.Add(lbl3)
            labels3.Add(lbl3)

            AddHandler lbl3.Click, AddressOf Me.lbl3_Click
            AddHandler lbl3.MouseEnter, AddressOf Me.lbl3_MouseEnter
            AddHandler lbl3.Paint, AddressOf lbl3_Paint

            Dim comboBox As New System.Windows.Forms.ComboBox()

            comboBox.DrawMode = DrawMode.OwnerDrawFixed
            AddHandler comboBox.DrawItem, AddressOf ComboBox_DrawItem
            AddHandler comboBox.MeasureItem, AddressOf ComboBox_MeasureItem
            AddHandler comboBox.SelectedIndexChanged, AddressOf ComboBox_SelectedIndexChanged

            comboBox.Items.Add("Primary Key")
            comboBox.Items.Add("    Primary Key")
            comboBox.Items.Add("Separator")
            comboBox.Items.Add("    Comma")
            comboBox.Items.Add("    Colon")
            comboBox.Items.Add("    Semicolon")
            comboBox.Items.Add("    Space")
            comboBox.Items.Add("    Nothing")
            comboBox.Items.Add("    New Line")
            comboBox.Items.Add("Function")
            comboBox.Items.Add("    Sum")
            comboBox.Items.Add("    Count")
            comboBox.Items.Add("    Average")
            comboBox.Items.Add("    Max")
            comboBox.Items.Add("    Min")
            comboBox.Items.Add("    Product")

            comboBox.Location = New System.Drawing.Point((Label2.Width + Label4.Width) - 8 + 0.5, (i - 1) * height + 0.5)
            comboBox.Height = height - 5
            comboBox.Font = New Font("Segoe UI", 9.75F)
            comboBox.Width = Label5.Width - 0.5
            comboBox.Visible = False

            CustomGroupBox7.Controls.Add(comboBox)
            comboBoxes.Add(comboBox)

        Next

    End Sub

    Private Sub ComboBox_DrawItem(ByVal sender As Object, ByVal e As DrawItemEventArgs)

        Dim comboBox As System.Windows.Forms.ComboBox
        comboBox = DirectCast(sender, System.Windows.Forms.ComboBox)

        If e.Index = -1 Then
            Return
        End If

        If e.Index >= 0 Then
            Dim isHeader As Boolean = comboBox.Items(e.Index).StartsWith("  ")
            If isHeader = False Then
                e.Graphics.FillRectangle(Brushes.LightGray, e.Bounds)
                e.Graphics.DrawString(comboBox.Items(e.Index).ToString(), e.Font, Brushes.Black, e.Bounds)
            Else
                e.DrawBackground()
                e.Graphics.DrawString(comboBox.Items(e.Index).ToString(), e.Font, Brushes.Black, e.Bounds)
            End If
        End If

    End Sub

    Private Sub ComboBox_MeasureItem(ByVal sender As Object, ByVal e As MeasureItemEventArgs)

        Dim comboBox As System.Windows.Forms.ComboBox
        comboBox = DirectCast(sender, System.Windows.Forms.ComboBox)

        If e.Index >= 0 Then
            Dim isHeader As Boolean = comboBox.Items(e.Index).StartsWith("  ")
            If isHeader = False Then
                e.ItemHeight = 20
            Else
                e.ItemHeight = 15
            End If
        End If

    End Sub

    Private Sub ComboBox_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs)

        Dim comboBox As System.Windows.Forms.ComboBox
        comboBox = DirectCast(sender, System.Windows.Forms.ComboBox)

        If comboBox.SelectedIndex >= 0 Then
            Dim isHeader As Boolean = comboBox.SelectedItem.StartsWith("    ")
            If isHeader = False Then
                comboBox.SelectedIndex = -1
            Else
                Dim clickedBoxNumber As Integer = comboBoxes.IndexOf(comboBox)
                labels3(clickedBoxNumber).Text = comboBox.SelectedItem
                labels3(clickedBoxNumber).Visible = True
                comboBox.Visible = False
            End If
        End If

    End Sub
    Private Sub lbl_Paint(sender As Object, e As PaintEventArgs)

        Dim lbl = DirectCast(sender, System.Windows.Forms.Label)
        Dim borderColor As Color = Color.FromArgb(245, 245, 245)
        Dim borderWidth As Integer = 0.2

        Dim borderPen As New Pen(borderColor, borderWidth)

        borderPen.DashStyle = Drawing2D.DashStyle.Dash

        e.Graphics.DrawRectangle(borderPen, 0, 0, lbl.Width - 1, lbl.Height - 1)

        borderPen.Dispose()

    End Sub
    Private Sub lbl2_Paint(sender As Object, e As PaintEventArgs)

        Dim lbl = DirectCast(sender, System.Windows.Forms.Label)
        Dim borderColor As Color = Color.FromArgb(245, 245, 245)
        Dim borderWidth As Integer = 0.2

        Dim borderPen As New Pen(borderColor, borderWidth)

        borderPen.DashStyle = Drawing2D.DashStyle.Dash

        e.Graphics.DrawRectangle(borderPen, 0, 0, lbl.Width - 1, lbl.Height - 1)

        borderPen.Dispose()

    End Sub
    Private Sub lbl3_Paint(sender As Object, e As PaintEventArgs)

        Dim lbl = DirectCast(sender, System.Windows.Forms.Label)
        Dim borderColor As Color = Color.FromArgb(245, 245, 245)
        Dim borderWidth As Integer = 0.2

        Dim borderPen As New Pen(borderColor, borderWidth)

        borderPen.DashStyle = Drawing2D.DashStyle.Dash

        e.Graphics.DrawRectangle(borderPen, 0, 0, lbl.Width - 1, lbl.Height - 1)

        borderPen.Dispose()

    End Sub

    Private Sub lbl_Click(sender As Object, e As EventArgs)

        Dim clickedLabel As System.Windows.Forms.Label
        clickedLabel = DirectCast(sender, System.Windows.Forms.Label)

        clickedLabelNumber = labels.IndexOf(clickedLabel)

        clickedLabel.BackColor = Color.FromArgb(217, 217, 217)
        labels2(clickedLabelNumber).BackColor = Color.FromArgb(217, 217, 217)
        labels3(clickedLabelNumber).BackColor = Color.FromArgb(217, 217, 217)

        For Each label As System.Windows.Forms.Label In labels
            Dim lNumber As Integer = labels.IndexOf(label)
            If lNumber <> clickedLabelNumber Then
                labels(lNumber).BackColor = Color.FromArgb(255, 255, 255)
                labels2(lNumber).BackColor = Color.FromArgb(255, 255, 255)
                labels3(lNumber).BackColor = Color.FromArgb(255, 255, 255)
                comboBoxes(lNumber).Visible = False
                labels3(lNumber).Visible = True
            End If
        Next

        comboBoxes(clickedLabelNumber).Visible = True
        labels3(clickedLabelNumber).Visible = False

    End Sub
    Private Sub lbl_MouseEnter(sender As Object, e As EventArgs)

        Dim clickedLabel As System.Windows.Forms.Label
        clickedLabel = DirectCast(sender, System.Windows.Forms.Label)

        EnteredLabelNumber = labels.IndexOf(clickedLabel)

        If (EnteredLabelNumber <> clickedLabelNumber) Then
            clickedLabel.BackColor = Color.FromArgb(229, 243, 255)
            labels2(EnteredLabelNumber).BackColor = Color.FromArgb(229, 243, 255)
            labels3(EnteredLabelNumber).BackColor = Color.FromArgb(229, 243, 255)
        End If

        For Each label As System.Windows.Forms.Label In labels
            Dim lNumber As Integer = labels.IndexOf(label)
            If lNumber <> EnteredLabelNumber And lNumber <> clickedLabelNumber Then
                labels(lNumber).BackColor = Color.FromArgb(255, 255, 255)
                labels2(lNumber).BackColor = Color.FromArgb(255, 255, 255)
                labels3(lNumber).BackColor = Color.FromArgb(255, 255, 255)
            End If
        Next

    End Sub
    Private Sub lbl2_Click(sender As Object, e As EventArgs)

        Dim clickedLabel As System.Windows.Forms.Label
        clickedLabel = DirectCast(sender, System.Windows.Forms.Label)

        clickedLabelNumber = labels2.IndexOf(clickedLabel)

        clickedLabel.BackColor = Color.FromArgb(217, 217, 217)
        labels(clickedLabelNumber).BackColor = Color.FromArgb(217, 217, 217)
        labels3(clickedLabelNumber).BackColor = Color.FromArgb(217, 217, 217)

        For Each label As System.Windows.Forms.Label In labels
            Dim lNumber As Integer = labels.IndexOf(label)
            If lNumber <> clickedLabelNumber Then
                labels(lNumber).BackColor = Color.FromArgb(255, 255, 255)
                labels2(lNumber).BackColor = Color.FromArgb(255, 255, 255)
                labels3(lNumber).BackColor = Color.FromArgb(255, 255, 255)
                comboBoxes(lNumber).Visible = False
                labels3(lNumber).Visible = True
            End If
        Next

        comboBoxes(clickedLabelNumber).Visible = True
        labels3(clickedLabelNumber).Visible = False

    End Sub
    Private Sub lbl2_MouseEnter(sender As Object, e As EventArgs)

        Dim clickedLabel As System.Windows.Forms.Label
        clickedLabel = DirectCast(sender, System.Windows.Forms.Label)

        EnteredLabelNumber = labels2.IndexOf(clickedLabel)

        If (EnteredLabelNumber <> clickedLabelNumber) Then
            clickedLabel.BackColor = Color.FromArgb(229, 243, 255)
            labels(EnteredLabelNumber).BackColor = Color.FromArgb(229, 243, 255)
            labels3(EnteredLabelNumber).BackColor = Color.FromArgb(229, 243, 255)
        End If

        For Each label As System.Windows.Forms.Label In labels
            Dim lNumber As Integer = labels.IndexOf(label)
            If lNumber <> EnteredLabelNumber And lNumber <> clickedLabelNumber Then
                labels(lNumber).BackColor = Color.FromArgb(255, 255, 255)
                labels2(lNumber).BackColor = Color.FromArgb(255, 255, 255)
                labels3(lNumber).BackColor = Color.FromArgb(255, 255, 255)
            End If
        Next

    End Sub
    Private Sub lbl3_Click(sender As Object, e As EventArgs)

        Dim clickedLabel As System.Windows.Forms.Label
        clickedLabel = DirectCast(sender, System.Windows.Forms.Label)

        clickedLabelNumber = labels3.IndexOf(clickedLabel)

        clickedLabel.BackColor = Color.FromArgb(217, 217, 217)
        labels(clickedLabelNumber).BackColor = Color.FromArgb(217, 217, 217)
        labels2(clickedLabelNumber).BackColor = Color.FromArgb(217, 217, 217)

        For Each label As System.Windows.Forms.Label In labels
            Dim lNumber As Integer = labels.IndexOf(label)
            If lNumber <> clickedLabelNumber Then
                labels(lNumber).BackColor = Color.FromArgb(255, 255, 255)
                labels2(lNumber).BackColor = Color.FromArgb(255, 255, 255)
                labels3(lNumber).BackColor = Color.FromArgb(255, 255, 255)
                comboBoxes(lNumber).Visible = False
                labels3(lNumber).Visible = True
            End If
        Next

        comboBoxes(clickedLabelNumber).Visible = True
        labels3(clickedLabelNumber).Visible = False

    End Sub
    Private Sub lbl3_MouseEnter(sender As Object, e As EventArgs)

        Dim clickedLabel As System.Windows.Forms.Label
        clickedLabel = DirectCast(sender, System.Windows.Forms.Label)

        EnteredLabelNumber = labels3.IndexOf(clickedLabel)

        If (EnteredLabelNumber <> clickedLabelNumber) Then
            clickedLabel.BackColor = Color.FromArgb(229, 243, 255)
            labels(EnteredLabelNumber).BackColor = Color.FromArgb(229, 243, 255)
            labels2(EnteredLabelNumber).BackColor = Color.FromArgb(229, 243, 255)
        End If

        For Each label As System.Windows.Forms.Label In labels
            Dim lNumber As Integer = labels.IndexOf(label)
            If lNumber <> EnteredLabelNumber And lNumber <> clickedLabelNumber Then
                labels(lNumber).BackColor = Color.FromArgb(255, 255, 255)
                labels2(lNumber).BackColor = Color.FromArgb(255, 255, 255)
                labels3(lNumber).BackColor = Color.FromArgb(255, 255, 255)
            End If
        Next

    End Sub

    Private Sub Display()

        CustomPanel1.Controls.Clear()
        CustomPanel2.Controls.Clear()

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

        For i = 1 To displayRng.Rows.Count
            For j = 1 To displayRng.Columns.Count
                Dim label As New System.Windows.Forms.Label
                label.Text = displayRng.Cells(i, j).Value
                label.Location = New System.Drawing.Point((j - 1) * width, (i - 1) * height)
                label.Height = height
                label.Width = width
                label.BorderStyle = BorderStyle.FixedSingle
                label.TextAlign = ContentAlignment.MiddleCenter

                If CheckBox4.Checked = True Then

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
            Next
        Next

        CustomPanel1.AutoScroll = True

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

        Try
            If TextBox1.Text <> "" And IsValidExcelCellReference(TextBox1.Text) = True Then
                Call Setup()
                Call Display()
            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub Form22_Merge_Duplicate_Rows_Load(sender As Object, e As EventArgs) Handles Me.Load

        clickedLabelNumber = -1
        EnteredLabelNumber = -1

    End Sub

    Private Sub CheckBox5_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox5.CheckedChanged

        Try
            Call Setup()

        Catch ex As Exception

        End Try

    End Sub

    Private Sub CheckBox4_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox4.CheckedChanged

        Call Display()

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

    End Sub
End Class
