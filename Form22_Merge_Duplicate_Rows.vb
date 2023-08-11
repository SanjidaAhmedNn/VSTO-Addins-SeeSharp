Imports System.Drawing
Imports System.Reflection
Imports System.Reflection.Emit
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Excel

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
    Private Function Search(Arr, value)

        Dim Result As Boolean
        Result = False

        For i = LBound(Arr) To UBound(Arr)

            Dim Type1 As Type = Arr(i).GetType
            Dim Type2 As Type = value.GetType

            If Type1.Equals(Type2) Then
                If Arr(i) = value Then
                    Result = True
                    Exit For
                End If
            End If
        Next

        Search = Result

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

    Private Sub Setup()

        CustomGroupBox7.Controls.Clear()

        labels.Clear()
        labels2.Clear()
        labels3.Clear()
        comboBoxes.Clear()

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
            lbl.Font = New System.Drawing.Font("Segoe UI", 9.75F)
            lbl.TextAlign = ContentAlignment.MiddleCenter
            lbl.TextAlign = ContentAlignment.MiddleLeft
            lbl.BorderStyle = BorderStyle.None
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
            lbl2.Font = New System.Drawing.Font("Segoe UI", 9.75F)
            lbl2.TextAlign = ContentAlignment.MiddleCenter
            lbl2.TextAlign = ContentAlignment.MiddleLeft
            lbl2.BorderStyle = BorderStyle.None
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
            lbl3.Font = New System.Drawing.Font("Segoe UI", 9.75F)
            lbl3.TextAlign = ContentAlignment.MiddleCenter
            lbl3.TextAlign = ContentAlignment.MiddleLeft
            lbl3.BorderStyle = BorderStyle.None
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
            comboBox.Font = New System.Drawing.Font("Segoe UI", 9.75F)
            comboBox.Width = Label5.Width - 0.5
            comboBox.Visible = False

            CustomGroupBox7.Controls.Add(comboBox)
            comboBoxes.Add(comboBox)

        Next
        clickedLabelNumber = 0
        labels(0).BackColor = Color.FromArgb(217, 217, 217)
        labels2(0).BackColor = Color.FromArgb(217, 217, 217)
        labels3(0).BackColor = Color.FromArgb(217, 217, 217)
        labels3(0).Text = "    Primary Key"

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

        Call Display()

    End Sub
    Private Sub lbl_Paint(sender As Object, e As PaintEventArgs)

        Dim lbl = DirectCast(sender, System.Windows.Forms.Label)
        Dim borderColor As Color = Color.FromArgb(245, 245, 245)
        Dim borderWidth As Double = 0.4

        Dim borderPen As New Pen(borderColor, borderWidth)

        borderPen.DashStyle = Drawing2D.DashStyle.Dash

        e.Graphics.DrawRectangle(borderPen, 0, 0, lbl.Width - 1, lbl.Height - 1)

        borderPen.Dispose()

    End Sub
    Private Sub lbl2_Paint(sender As Object, e As PaintEventArgs)

        Dim lbl = DirectCast(sender, System.Windows.Forms.Label)
        Dim borderColor As Color = Color.FromArgb(245, 245, 245)
        Dim borderWidth As Double = 0.4

        Dim borderPen As New Pen(borderColor, borderWidth)

        borderPen.DashStyle = Drawing2D.DashStyle.Dash

        e.Graphics.DrawRectangle(borderPen, 0, 0, lbl.Width - 1, lbl.Height - 1)

        borderPen.Dispose()

    End Sub
    Private Sub lbl3_Paint(sender As Object, e As PaintEventArgs)

        Dim lbl = DirectCast(sender, System.Windows.Forms.Label)
        Dim borderColor As Color = Color.FromArgb(245, 245, 245)
        Dim borderWidth As Double = 0.4

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

        excelApp = Globals.ThisAddIn.Application
        workBook = excelApp.ActiveWorkbook
        workSheet = workBook.ActiveSheet
        rng = workSheet.Range(TextBox1.Text)

        If CheckBox5.Checked = True Then
            rng = workSheet.Range(rng.Cells(2, 1), rng.Cells(rng.Rows.Count, rng.Columns.Count))
        End If

        rng.Select()

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

        Dim Active As Boolean = True

        For Each lbl In labels3
            If lbl.Text = "" Then
                Active = False
                Exit For
            End If
        Next

        Dim IsPrimary As Boolean
        IsPrimary = False

        Dim PrimaryColumn As Integer = 0

        For Each lbl In labels3
            If lbl.Text = "    Primary Key" Then
                IsPrimary = True
                PrimaryColumn = labels3.IndexOf(lbl) + 1
                Exit For
            End If
        Next

        Active = Active And IsPrimary

        If Active = True Then

            Dim cRng As Excel.Range
            cRng = workSheet.Range(displayRng.Cells(1, PrimaryColumn), displayRng.Cells(displayRng.Rows.Count, PrimaryColumn))

            Dim Arr1(0) As Object
            Dim Arr2(0) As Integer

            Dim Index1 As Integer = 0
            Dim Index2 As Integer = 0

            Arr1(0) = cRng.Cells(1, 1).Value
            Arr2(0) = 1

            For i = 1 To cRng.Rows.Count
                If Search(Arr1, cRng.Cells(i, 1).Value) = False Then
                    Index1 = Index1 + 1
                    Index2 = Index2 + 1
                    ReDim Preserve Arr1(Index1)
                    ReDim Preserve Arr2(Index2)
                    Arr1(Index1) = cRng.Cells(i, 1).Value
                    Arr2(Index2) = i
                End If
            Next

            If (UBound(Arr1) + 1) <= 6 Then
                height = CustomPanel1.Height / (UBound(Arr1) + 1)
            Else
                height = CustomPanel1.Height / 6
            End If

            Dim ordinate As Single = 0
            For j = 1 To displayRng.Columns.Count
                If j <> PrimaryColumn Then
                    Dim max As Integer = 1
                    For k = LBound(Arr1) To UBound(Arr1)
                        Dim count As Integer = 0
                        For i = 1 To displayRng.Rows.Count
                            If displayRng.Cells(i, PrimaryColumn).value = Arr1(k) Then
                                count = count + 1
                            End If
                        Next
                        If count > max Then
                            max = count
                        End If
                    Next

                    For k = LBound(Arr1) To UBound(Arr1)

                        Dim separator As String = " "
                        Dim concatenatedValue As String = ""

                        For i = 1 To displayRng.Rows.Count
                            If displayRng.Cells(i, PrimaryColumn).value = Arr1(k) Then
                                concatenatedValue = concatenatedValue & displayRng.Cells(i, j).Value & separator
                            End If
                        Next

                        Dim label As New System.Windows.Forms.Label
                        label.Text = concatenatedValue
                        label.Location = New System.Drawing.Point(ordinate, (k + 1 - 1) * height)
                        label.Height = height
                        label.Width = max * width
                        label.BorderStyle = BorderStyle.FixedSingle
                        label.TextAlign = ContentAlignment.MiddleCenter
                        CustomPanel2.Controls.Add(label)

                        If CheckBox4.Checked = True Then

                            Dim cell As Excel.Range = displayRng.Cells(Arr2(k), j)
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
                    Next

                    ordinate = ordinate + max * width
                Else
                    For k = LBound(Arr1) To UBound(Arr1)
                        Dim label As New System.Windows.Forms.Label
                        label.Text = Arr1(k)
                        label.Location = New System.Drawing.Point(ordinate, (k + 1 - 1) * height)
                        label.Height = height
                        label.Width = width
                        label.BorderStyle = BorderStyle.FixedSingle
                        label.TextAlign = ContentAlignment.MiddleCenter
                        CustomPanel2.Controls.Add(label)

                        If CheckBox4.Checked = True Then

                            Dim cell As Excel.Range = displayRng.Cells(Arr2(k), j)
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
                    Next
                    ordinate = ordinate + width

                End If
            Next
            CustomPanel2.AutoScroll = True
        End If
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

        EnteredLabelNumber = -1

    End Sub

    Private Sub CheckBox5_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox5.CheckedChanged

        Try
            Call Setup()
            Call Display()

        Catch ex As Exception

        End Try

    End Sub

    Private Sub CheckBox4_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox4.CheckedChanged

        Try
            Call Display()

        Catch ex As Exception

        End Try

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        excelApp = Globals.ThisAddIn.Application
        workBook = excelApp.ActiveWorkbook
        workSheet = workBook.ActiveSheet
        workSheet2 = workBook.ActiveSheet
        rng = workSheet.Range(TextBox1.Text)
        rng2 = workSheet2.Range(TextBox2.Text)
        Dim rng2Address As String = rng2.Address

        If CheckBox5.Checked = True Then
            rng = workSheet.Range(rng.Cells(2, 1), rng.Cells(rng.Rows.Count, rng.Columns.Count))
        End If

        rng.Select()


        Dim r As Integer
        Dim c As Integer

        r = rng.Rows.Count
        c = rng.Columns.Count

        Dim Active As Boolean = True

        For Each lbl In labels3
            If lbl.Text = "" Then
                Active = False
                Exit For
            End If
        Next

        Dim IsPrimary As Boolean
        IsPrimary = False

        Dim PrimaryColumn As Integer = 0

        For Each lbl In labels3
            If lbl.Text = "    Primary Key" Then
                IsPrimary = True
                PrimaryColumn = labels3.IndexOf(lbl) + 1
                Exit For
            End If
        Next

        Active = Active And IsPrimary

        If Active = True Then

            Dim cRng As Excel.Range
            cRng = workSheet.Range(rng.Cells(1, PrimaryColumn), rng.Cells(rng.Rows.Count, PrimaryColumn))

            Dim Arr1(0) As Object
            Dim Arr2(0) As Integer

            Dim Index1 As Integer = 0
            Dim Index2 As Integer = 0

            Arr1(0) = cRng.Cells(1, 1).Value
            Arr2(0) = 1

            For i = 1 To cRng.Rows.Count
                If Search(Arr1, cRng.Cells(i, 1).Value) = False Then
                    Index1 = Index1 + 1
                    Index2 = Index2 + 1
                    ReDim Preserve Arr1(Index1)
                    ReDim Preserve Arr2(Index2)
                    Arr1(Index1) = cRng.Cells(i, 1).Value
                    Arr2(Index2) = i
                End If
            Next

            For j = 1 To rng.Columns.Count
                If j <> PrimaryColumn Then
                    For k = LBound(Arr1) To UBound(Arr1)

                        Dim separator As String = " "
                        Dim concatenatedValue As String = ""

                        Dim index As Integer = 0
                        For i = 1 To rng.Rows.Count
                            If rng.Cells(i, PrimaryColumn).value = Arr1(k) Then
                                concatenatedValue = concatenatedValue & rng.Cells(i, j).Value & separator
                                If index = 0 Then
                                    index = i
                                End If
                            End If
                        Next

                        rng2.Cells(k + 1, j).value = concatenatedValue

                        If Overlap(excelApp, workSheet, workSheet2, rng, rng2) = False Then
                            If CheckBox4.Checked = True Then
                                rng.Cells(index, j).Copy()
                                rng2.Cells(k + 1, j).PasteSpecial(Excel.XlPasteType.xlPasteFormats)
                                rng2 = workSheet2.Range(rng2Address)
                            End If
                            excelApp.CutCopyMode = Excel.XlCutCopyMode.xlCopy
                        End If

                    Next
                Else
                    For k = LBound(Arr1) To UBound(Arr1)
                        rng2.Cells(k + 1, j).value = Arr1(k)

                        If Overlap(excelApp, workSheet, workSheet2, rng, rng2) = False Then
                            If CheckBox4.Checked = True Then
                                rng.Cells(Arr2(k), j).Copy()
                                rng2.Cells(k + 1, j).PasteSpecial(Excel.XlPasteType.xlPasteFormats)
                                rng2 = workSheet2.Range(rng2Address)
                            End If
                            excelApp.CutCopyMode = Excel.XlCutCopyMode.xlCopy
                        End If
                    Next

                End If
            Next
        End If

    End Sub
End Class
