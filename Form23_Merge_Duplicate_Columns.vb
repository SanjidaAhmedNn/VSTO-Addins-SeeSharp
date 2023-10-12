Imports System.ComponentModel
Imports System.Drawing
Imports System.Reflection
Imports System.Reflection.Emit
Imports System.Security.Policy
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Button
Imports Microsoft.Office.Interop.Excel

Public Class Form23_Merge_Duplicate_Columns

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


    Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As IntPtr, ByVal hWndInsertAfter As IntPtr, ByVal X As Integer, ByVal Y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal uFlags As UInteger) As Boolean
    Private Const SWP_NOMOVE As UInteger = &H2
    Private Const SWP_NOSIZE As UInteger = &H1
    Private Const SWP_NOACTIVATE As UInteger = &H10
    Private Const HWND_TOPMOST As Integer = -1


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
    Private Function GetUniques(rng As Excel.Range, CaseSensitive As Boolean, RowWise As Boolean)

        Dim Uniques(0) As Integer
        Uniques(0) = 1
        Dim Index As Integer = 0

        If RowWise = True Then
            Dim Matched As Boolean
            For i = 2 To rng.Rows.Count
                Matched = False
                For l = LBound(Uniques) To UBound(Uniques)
                    Dim count As Integer = 0
                    For j = 1 To rng.Columns.Count
                        Dim Type1 As Type
                        Dim Type2 As Type

                        If rng.Cells(i, j).value Is Nothing Then
                            Type1 = GetType(String)
                        Else
                            Type1 = rng.Cells(i, j).value.GetType()
                        End If

                        If rng.Cells(Uniques(l), j).value Is Nothing Then
                            Type2 = GetType(String)
                        Else
                            Type2 = rng.Cells(Uniques(l), j).value.GetType()
                        End If

                        If Type1.Equals(Type2) Then
                            If CaseSensitive = True Then
                                If rng.Cells(i, j).value = rng.Cells(Uniques(l), j).value Then
                                    count = count + 1
                                End If
                            Else
                                If LCase(rng.Cells(i, j).value) = LCase(rng.Cells(Uniques(l), j).value) Then
                                    count = count + 1
                                End If
                            End If
                        End If
                    Next
                    If count = rng.Columns.Count Then
                        Matched = True
                        Exit For
                    End If
                Next
                If Matched = False Then
                    Index = Index + 1
                    ReDim Preserve Uniques(Index)
                    Uniques(Index) = i
                End If
            Next

            GetUniques = Uniques

        Else

            Dim Matched As Boolean
            For j = 2 To rng.Columns.Count
                Matched = False
                For l = LBound(Uniques) To UBound(Uniques)
                    Dim count As Integer = 0
                    For i = 1 To rng.Rows.Count
                        Dim Type1 As Type
                        Dim Type2 As Type

                        If rng.Cells(i, j).value Is Nothing Then
                            Type1 = GetType(String)
                        Else
                            Type1 = rng.Cells(i, j).value.GetType()
                        End If

                        If rng.Cells(i, Uniques(l)).value Is Nothing Then
                            Type2 = GetType(String)
                        Else
                            Type2 = rng.Cells(i, Uniques(l)).value.GetType()
                        End If

                        If Type1.Equals(Type2) Then
                            If CaseSensitive = True Then
                                If rng.Cells(i, j).value = rng.Cells(i, Uniques(l)).value Then
                                    count = count + 1
                                End If
                            Else
                                If LCase(rng.Cells(i, j).value) = LCase(rng.Cells(i, Uniques(l)).value) Then
                                    count = count + 1
                                End If
                            End If
                        End If
                    Next
                    If count = rng.Rows.Count Then
                        Matched = True
                        Exit For
                    End If
                Next
                If Matched = False Then
                    Index = Index + 1
                    ReDim Preserve Uniques(Index)
                    Uniques(Index) = j
                End If
            Next

            GetUniques = Uniques

        End If

    End Function
    Private Function SearchInArray(Arr, value)

        Dim Result As Boolean = False

        For i = LBound(Arr) To UBound(Arr)
            If Arr(i) = value Then
                Result = True
                Exit For
            End If
        Next

        SearchInArray = Result

    End Function

    Private Function Operation(Arr, Flag)

        If Flag = "    Sum" Then
            Dim Output As Double = 0
            For i = LBound(Arr) To UBound(Arr)
                If IsNumeric(Arr(i)) = True Then
                    Output = Output + Arr(i)
                End If
            Next
            Operation = Output

        ElseIf Flag = "    Count" Then
            Dim Output As Integer = 0
            For i = LBound(Arr) To UBound(Arr)
                If Arr(i) IsNot Nothing Then
                    Output = Output + 1
                End If
            Next
            Operation = Output

        ElseIf Flag = "    Average" Then
            Dim Output As Double = 0
            For i = LBound(Arr) To UBound(Arr)
                If IsNumeric(Arr(i)) = True Then
                    Output = Output + Arr(i)
                End If
            Next
            Output = Output / (UBound(Arr) + 1)
            Operation = Output

        ElseIf Flag = "    Max" Then
            Dim Output As Object
            Dim i As Integer = LBound(Arr)
            While IsNumeric(Arr(i)) = False And i <= UBound(Arr) - 1
                i = i + 1
            End While
            Output = Arr(i)
            For i = LBound(Arr) To UBound(Arr)
                If IsNumeric(Arr(i)) = True Then
                    If Arr(i) > Output Then
                        Output = Arr(i)
                    End If
                End If
            Next
            Operation = Output

        ElseIf Flag = "    Min" Then
            Dim Output As Object
            Dim i As Integer = LBound(Arr)
            While IsNumeric(Arr(i)) = False And i <= UBound(Arr) - 1
                i = i + 1
            End While

            Output = Arr(i)
            For i = LBound(Arr) To UBound(Arr)
                If IsNumeric(Arr(i)) = True Then
                    If Arr(i) < Output Then
                        Output = Arr(i)
                    End If
                End If
            Next
            Operation = Output

        ElseIf Flag = "    Product" Then
            Dim Output As Double = 1
            Dim count As Integer = 0
            For i = LBound(Arr) To UBound(Arr)
                If IsNumeric(Arr(i)) = True Then
                    Output = Output * Arr(i)
                    count = count + 1
                End If
            Next
            If count = 0 Then
                Operation = 0
            Else
                Operation = Output
            End If
        Else
            Operation = 0
        End If

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
    Private Function Search(Arr, value, CaseSensitive)

        Dim Result As Boolean
        Result = False

        For i = LBound(Arr) To UBound(Arr)

            Dim Type1 As Type
            Dim Type2 As Type

            If Arr(i) Is Nothing Then
                Type1 = GetType(String)
            Else
                Type1 = Arr(i).GetType()
            End If

            If value Is Nothing Then
                Type2 = GetType(String)
            Else
                Type2 = value.GetType()
            End If

            If Type1.Equals(Type2) Then
                If CaseSensitive = True Then
                    If Arr(i) = value Then
                        Result = True
                        Exit For
                    End If
                Else
                    If IsNumeric(Arr(i)) = True Then
                        If Arr(i) = value Then
                            Result = True
                            Exit For
                        End If
                    Else
                        If LCase(Arr(i)) = LCase(value) Then
                            Result = True
                            Exit For
                        End If
                    End If
                End If
            End If
        Next

        Search = Result

    End Function
    Private Sub Setup()

        Try
            CustomGroupBox7.Controls.Clear()

            labels.Clear()
            labels2.Clear()
            labels3.Clear()
            comboBoxes.Clear()

            rng.Select()

            Dim height As Single = Label3.Height

            Dim i As Integer

            For i = 1 To rng.Rows.Count

                Dim lbl As New System.Windows.Forms.Label()
                If CheckBox5.Checked = True Then
                    lbl.Text = rng.Cells(i, 0).Value
                Else
                    Dim rowLetter As String = Split(rng.Cells(i, 1).Address(True, True), "$")(2)
                    lbl.Text = "Row " & rowLetter
                End If
                lbl.Location = New System.Drawing.Point(1, (i - 1) * height)
                lbl.Height = height
                lbl.Width = Label2.Width - 1
                lbl.Font = New System.Drawing.Font("Segoe UI", 9.75F)
                lbl.TextAlign = ContentAlignment.MiddleCenter
                lbl.TextAlign = ContentAlignment.MiddleLeft
                lbl.BorderStyle = BorderStyle.None
                CustomGroupBox7.Controls.Add(lbl)
                labels.Add(lbl)

                AddHandler lbl.Click, AddressOf Me.lbl_Click
                AddHandler lbl.MouseEnter, AddressOf Me.lbl_MouseEnter
                AddHandler lbl.Paint, AddressOf lbl_Paint
                AddHandler lbl.KeyDown, AddressOf lbl_KeyDown

                Dim lbl2 As New System.Windows.Forms.Label
                lbl2.Text = rng.Cells(1, i).Value
                lbl2.Location = New System.Drawing.Point(Label2.Width - 1, (i - 1) * height)
                lbl2.Height = height
                lbl2.Width = Label4.Width + 0.5
                lbl2.Font = New System.Drawing.Font("Segoe UI", 9.75F)
                lbl2.TextAlign = ContentAlignment.MiddleCenter
                lbl2.TextAlign = ContentAlignment.MiddleLeft
                lbl2.BorderStyle = BorderStyle.None
                CustomGroupBox7.Controls.Add(lbl2)
                labels2.Add(lbl2)

                AddHandler lbl2.Click, AddressOf Me.lbl2_Click
                AddHandler lbl2.MouseEnter, AddressOf Me.lbl2_MouseEnter
                AddHandler lbl2.Paint, AddressOf lbl2_Paint
                AddHandler lbl2.KeyDown, AddressOf lbl2_KeyDown

                Dim lbl3 As New System.Windows.Forms.Label
                lbl3.Text = ""
                lbl3.Location = New System.Drawing.Point((Label2.Width + Label4.Width - 0.5), (i - 1) * height)
                lbl3.Height = height
                lbl3.Width = Label5.Width - 1
                lbl3.Font = New System.Drawing.Font("Segoe UI", 9.75F)
                lbl3.TextAlign = ContentAlignment.MiddleCenter
                lbl3.TextAlign = ContentAlignment.MiddleLeft
                lbl3.BorderStyle = BorderStyle.None
                CustomGroupBox7.Controls.Add(lbl3)
                labels3.Add(lbl3)

                AddHandler lbl3.Click, AddressOf Me.lbl3_Click
                AddHandler lbl3.MouseEnter, AddressOf Me.lbl3_MouseEnter
                AddHandler lbl3.Paint, AddressOf lbl3_Paint
                AddHandler lbl3.KeyDown, AddressOf lbl3_KeyDown

                Dim comboBox As New System.Windows.Forms.ComboBox()

                comboBox.DrawMode = DrawMode.OwnerDrawFixed
                AddHandler comboBox.DrawItem, AddressOf ComboBox_DrawItem
                AddHandler comboBox.MeasureItem, AddressOf ComboBox_MeasureItem
                AddHandler comboBox.SelectedIndexChanged, AddressOf ComboBox_SelectedIndexChanged
                AddHandler comboBox.KeyDown, AddressOf comboBox_KeyDown

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

                comboBox.Location = New System.Drawing.Point((Label2.Width + Label4.Width), (i - 1) * height + 0.5)
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

        Catch ex As Exception

        End Try

    End Sub
    Private Sub lbl_KeyDown(sender As Object, e As KeyEventArgs)

        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try
    End Sub
    Private Sub lbl2_KeyDown(sender As Object, e As KeyEventArgs)
        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try
    End Sub
    Private Sub lbl3_KeyDown(sender As Object, e As KeyEventArgs)

        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try

    End Sub
    Private Sub comboBox_KeyDown(sender As Object, e As KeyEventArgs)

        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub ComboBox_DrawItem(ByVal sender As Object, ByVal e As DrawItemEventArgs)

        Try

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

        Catch ex As Exception

        End Try

    End Sub

    Private Sub ComboBox_MeasureItem(ByVal sender As Object, ByVal e As MeasureItemEventArgs)

        Try
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

        Catch ex As Exception

        End Try

    End Sub

    Private Sub ComboBox_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs)

        Try
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

            Dim count As Integer = 0
            For Each label As System.Windows.Forms.Label In labels3
                If label.Text = "    Primary Key" Then
                    count = count + 1
                    If count > 1 Then
                        MsgBox("There can't be more than one primary key.")
                        label.Text = ""
                        Exit Sub
                    End If
                End If
            Next

            Call Display()

        Catch ex As Exception

        End Try
    End Sub
    Private Sub lbl_Paint(sender As Object, e As PaintEventArgs)

        Try

            Dim lbl = DirectCast(sender, System.Windows.Forms.Label)
            Dim borderColor As Color = Color.FromArgb(245, 245, 245)
            Dim borderWidth As Double = 0.4

            Dim borderPen As New Pen(borderColor, borderWidth)

            borderPen.DashStyle = Drawing2D.DashStyle.Dash

            e.Graphics.DrawRectangle(borderPen, 0, 0, lbl.Width - 1, lbl.Height - 1)

            borderPen.Dispose()

        Catch ex As Exception

        End Try

    End Sub
    Private Sub lbl2_Paint(sender As Object, e As PaintEventArgs)

        Try
            Dim lbl = DirectCast(sender, System.Windows.Forms.Label)
            Dim borderColor As Color = Color.FromArgb(245, 245, 245)
            Dim borderWidth As Double = 0.4

            Dim borderPen As New Pen(borderColor, borderWidth)

            borderPen.DashStyle = Drawing2D.DashStyle.Dash

            e.Graphics.DrawRectangle(borderPen, 0, 0, lbl.Width - 1, lbl.Height - 1)

            borderPen.Dispose()

        Catch ex As Exception

        End Try

    End Sub
    Private Sub lbl3_Paint(sender As Object, e As PaintEventArgs)

        Try

            Dim lbl = DirectCast(sender, System.Windows.Forms.Label)
            Dim borderColor As Color = Color.FromArgb(245, 245, 245)
            Dim borderWidth As Double = 0.4

            Dim borderPen As New Pen(borderColor, borderWidth)

            borderPen.DashStyle = Drawing2D.DashStyle.Dash

            e.Graphics.DrawRectangle(borderPen, 0, 0, lbl.Width - 1, lbl.Height - 1)

            borderPen.Dispose()

        Catch ex As Exception

        End Try

    End Sub

    Private Sub lbl_Click(sender As Object, e As EventArgs)

        Try

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

        Catch ex As Exception

        End Try

    End Sub
    Private Sub lbl_MouseEnter(sender As Object, e As EventArgs)

        Try

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
        Catch ex As Exception

        End Try

    End Sub
    Private Sub lbl2_Click(sender As Object, e As EventArgs)

        Try
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

        Catch ex As Exception

        End Try
    End Sub
    Private Sub lbl2_MouseEnter(sender As Object, e As EventArgs)

        Try
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
        Catch ex As Exception

        End Try

    End Sub
    Private Sub lbl3_Click(sender As Object, e As EventArgs)

        Try
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

        Catch ex As Exception

        End Try

    End Sub
    Private Sub lbl3_MouseEnter(sender As Object, e As EventArgs)

        Try
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
        Catch ex As Exception

        End Try

    End Sub

    Private Sub Display()

        Try
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

            height = CustomPanel2.Height / 6

            CustomPanel2.AutoScroll = True

            Dim Active As Boolean = True

            For Each lbl In labels3
                If lbl.Text = "" Then
                    Active = False
                    Exit For
                End If
            Next

            Dim IsPrimary As Boolean
            IsPrimary = False

            Dim PrimaryRow As Integer = 0

            For Each lbl In labels3
                If lbl.Text = "    Primary Key" Then
                    IsPrimary = True
                    PrimaryRow = labels3.IndexOf(lbl) + 1
                    Exit For
                End If
            Next

            Active = Active And IsPrimary

            If Active = True Then

                Dim cRng As Excel.Range
                cRng = workSheet.Range(displayRng.Cells(PrimaryRow, 1), displayRng.Cells(PrimaryRow, displayRng.Columns.Count))

                Dim Arr1(0) As Object
                Dim Arr2(0) As Integer

                Dim Index1 As Integer = 0
                Dim Index2 As Integer = 0

                Arr1(0) = cRng.Cells(1, 1).Value
                Arr2(0) = 1

                Dim CaseSensitive As Boolean

                If CheckBox1.Checked = True Then
                    CaseSensitive = True
                Else
                    CaseSensitive = False
                End If

                If CheckBox3.Checked = True Then
                    For j = 1 To cRng.Columns.Count
                        If cRng.Cells(1, j).Value IsNot Nothing Then
                            If Search(Arr1, cRng.Cells(1, j).Value, CaseSensitive) = False Then
                                Index1 = Index1 + 1
                                Index2 = Index2 + 1
                                ReDim Preserve Arr1(Index1)
                                ReDim Preserve Arr2(Index2)
                                Arr1(Index1) = cRng.Cells(1, j).Value
                                Arr2(Index2) = j
                            End If
                        End If
                    Next
                Else
                    For j = 1 To cRng.Columns.Count
                        If Search(Arr1, cRng.Cells(1, j).Value, CaseSensitive) = False Then
                            Index1 = Index1 + 1
                            Index2 = Index2 + 1
                            ReDim Preserve Arr1(Index1)
                            ReDim Preserve Arr2(Index2)
                            Arr1(Index1) = cRng.Cells(1, j).Value
                            Arr2(Index2) = j
                        End If
                    Next
                End If

                If (UBound(Arr1) + 1) <= 6 Then
                    width = CustomPanel2.Width / (UBound(Arr1) + 1)
                Else
                    width = CustomPanel2.Width / 6
                End If

                Dim abscissa As Single = 0

                Dim UniQueArr() As Integer = GetUniques(displayRng, CaseSensitive, False)

                For i = 1 To displayRng.Rows.Count
                    If i <> PrimaryRow Then

                        Dim max As Integer = 1
                        For k = LBound(Arr1) To UBound(Arr1)
                            Dim count As Integer = 0

                            For j = 1 To displayRng.Columns.Count

                                Dim DuplicateCondition As Boolean
                                If CheckBox6.Checked = True Then
                                    DuplicateCondition = SearchInArray(UniQueArr, j)
                                Else
                                    DuplicateCondition = True
                                End If

                                If displayRng.Cells(i, j).value IsNot Nothing And DuplicateCondition Then
                                    Dim Matched As Boolean

                                    If CheckBox1.Checked = True Then
                                        Matched = displayRng.Cells(PrimaryRow, j).value = Arr1(k)
                                    Else
                                        Matched = LCase(displayRng.Cells(PrimaryRow, j).value) = LCase(Arr1(k))
                                    End If

                                    If Matched = True Then
                                        count = count + 1
                                    End If
                                End If
                            Next
                            If count > max Then
                                max = count
                            End If
                        Next

                        Dim widthFlag As Boolean
                        Dim separator As String = " "
                        Dim Flag As String = ""

                        If labels3(i - 1).Text = "    Comma" Then
                            separator = ", "
                            widthFlag = True
                            Flag = "a"
                        ElseIf labels3(i - 1).Text = "    Colon" Then
                            separator = ": "
                            widthFlag = True
                            Flag = "a"
                        ElseIf labels3(i - 1).Text = "    Semicolon" Then
                            separator = "; "
                            widthFlag = True
                            Flag = "a"
                        ElseIf labels3(i - 1).Text = "    Space" Then
                            separator = " "
                            widthFlag = True
                            Flag = "b"
                        ElseIf labels3(i - 1).Text = "    Nothing" Then
                            separator = ""
                            widthFlag = True
                            Flag = "c"
                        ElseIf labels3(i - 1).Text = "    New Line" Then
                            separator = vbNewLine
                            widthFlag = True
                            Flag = "b"
                        Else
                            widthFlag = False
                            Flag = labels3(i - 1).Text
                        End If

                        For k = LBound(Arr1) To UBound(Arr1)

                            Dim concatenatedValue As String = ""
                            Dim OperatedValue As Object
                            Dim Valuess(0) As Object
                            Dim indx As Integer = -1

                            For j = 1 To displayRng.Columns.Count

                                Dim DuplicateCondition As Boolean
                                If CheckBox6.Checked = True Then
                                    DuplicateCondition = SearchInArray(UniQueArr, j)
                                Else
                                    DuplicateCondition = True
                                End If

                                If widthFlag = True Then
                                    If displayRng.Cells(i, j).Value IsNot Nothing And DuplicateCondition Then
                                        Dim Matched As Boolean
                                        If CheckBox1.Checked = True Then
                                            Matched = displayRng.Cells(PrimaryRow, j).value = Arr1(k)
                                        Else
                                            Matched = LCase(displayRng.Cells(PrimaryRow, j).value) = LCase(Arr1(k))
                                        End If

                                        If Matched = True Then
                                            concatenatedValue = concatenatedValue & displayRng.Cells(i, j).Value & separator
                                        End If
                                    End If
                                Else

                                    Dim Matched As Boolean
                                    If displayRng.Cells(i, j).Value IsNot Nothing And DuplicateCondition Then
                                        If CheckBox1.Checked = True Then
                                            Matched = displayRng.Cells(PrimaryRow, j).value = Arr1(k)
                                        Else
                                            Matched = LCase(displayRng.Cells(PrimaryRow, j).value) = LCase(Arr1(k))
                                        End If
                                        If Matched = True Then
                                            indx = indx + 1
                                            ReDim Preserve Valuess(indx)
                                            Valuess(indx) = displayRng.Cells(i, j).Value
                                        End If
                                    End If
                                End If
                            Next
                            OperatedValue = Operation(Valuess, Flag)

                            If Flag = "a" Then
                                If concatenatedValue <> "" Then
                                    concatenatedValue = Mid(concatenatedValue, 1, Len(concatenatedValue) - 2)
                                End If
                            ElseIf Flag = "b" Then
                                If concatenatedValue <> "" Then
                                    concatenatedValue = Mid(concatenatedValue, 1, Len(concatenatedValue) - 1)
                                End If
                            End If

                            Dim label As New System.Windows.Forms.Label

                            label.Font = New System.Drawing.Font("Segoe UI", 9.75F)
                            label.Location = New System.Drawing.Point((k + 1 - 1) * width, abscissa)
                            label.Width = width

                            If widthFlag = True Then
                                label.Height = (Int(max / 2) + 1) * height
                                label.Text = concatenatedValue
                            Else
                                label.Height = height
                                label.Text = OperatedValue
                            End If

                            label.TextAlign = ContentAlignment.MiddleCenter
                            CustomPanel2.Controls.Add(label)

                            If CheckBox4.Checked = True Then

                                Dim cell As Excel.Range = displayRng.Cells(i, Arr2(k))
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

                            AddHandler label.Paint, AddressOf label_Paint

                        Next

                        If widthFlag = True Then
                            abscissa = abscissa + (Int(max / 2) + 1) * height
                        Else
                            abscissa = abscissa + height
                        End If

                    Else
                        For k = LBound(Arr1) To UBound(Arr1)
                            Dim label As New System.Windows.Forms.Label
                            label.Text = Arr1(k)
                            label.Font = New System.Drawing.Font("Segoe UI", 9.75F)
                            label.Location = New System.Drawing.Point((k + 1 - 1) * width, abscissa)
                            label.Height = height
                            label.Width = width
                            label.TextAlign = ContentAlignment.MiddleCenter
                            CustomPanel2.Controls.Add(label)

                            If CheckBox4.Checked = True Then

                                Dim cell As Excel.Range = displayRng.Cells(i, Arr2(k))
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

                            AddHandler label.Paint, AddressOf label_Paint

                        Next

                        abscissa = abscissa + height


                    End If
                Next
                CustomPanel2.AutoScroll = True
            End If

        Catch ex As Exception

        End Try

    End Sub
    Private Sub label_Paint(sender As Object, e As PaintEventArgs)

        Try

            Dim lbl = DirectCast(sender, System.Windows.Forms.Label)
            Dim borderColor As Color = Color.FromArgb(245, 245, 245)
            Dim borderWidth As Double = 0.4

            Dim borderPen As New Pen(borderColor, borderWidth)

            borderPen.DashStyle = Drawing2D.DashStyle.Dash

            e.Graphics.DrawRectangle(borderPen, 0, 0, lbl.Width - 1, lbl.Height - 1)

            borderPen.Dispose()

        Catch ex As Exception

        End Try

    End Sub

    Private Sub Form23_Merge_Duplicate_Columns_Load(sender As Object, e As EventArgs) Handles Me.Load

        Try

            excelApp = Globals.ThisAddIn.Application

            AddHandler excelApp.SheetSelectionChange, AddressOf excelApp_SheetSelectionChange

            opened = opened + 1

            EnteredLabelNumber = -1

        Catch ex As Exception

        End Try

    End Sub

    Private Sub excelApp_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Excel.Range)

        Try

            excelApp = Globals.ThisAddIn.Application
            Dim selectedRange As Excel.Range
            selectedRange = excelApp.Selection
            If FocusedTextBox = 1 Then
                TextBox1.Text = selectedRange.Address
                workSheet = workBook.ActiveSheet
                rng = selectedRange
                TextBox1.Focus()
            ElseIf FocusedTextBox = 2 Then
                TextBox2.Text = selectedRange.Address
                workSheet2 = workBook.ActiveSheet
                rng2 = selectedRange
                TextBox2.Focus()
            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

        Try
            If TextBox1.Text <> "" And IsValidExcelCellReference(TextBox1.Text) = True Then

                excelApp = Globals.ThisAddIn.Application
                workBook = excelApp.ActiveWorkbook
                workSheet = workBook.ActiveSheet

                TextBox1.SelectionStart = TextBox1.Text.Length
                TextBox1.ScrollToCaret()

                rng = workSheet.Range(TextBox1.Text)
                rng.Select()

                Call Setup()
                Call Display()

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub CheckBox5_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox5.CheckedChanged

        Try

            If CheckBox5.Checked = True Then
                rng = workSheet.Range(rng.Cells(1, 2), rng.Cells(rng.Rows.Count, rng.Columns.Count))
            Else
                rng = workSheet.Range(rng.Cells(1, 0), rng.Cells(rng.Rows.Count, rng.Columns.Count))
            End If

            TextBox1.Text = rng.Address

        Catch ex As Exception

        End Try

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        If TextBox1.Text = "" Then
            MessageBox.Show("Select a Source Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            TextBox1.Focus()
            Exit Sub
        End If

        If IsValidExcelCellReference(TextBox1.Text) = False Then
            MessageBox.Show("Select a Valid Source Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            TextBox1.Focus()
            Exit Sub
        End If

        If RadioButton10.Checked = False And RadioButton3.Checked = False Then
            MessageBox.Show("Select a Destination Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        If RadioButton10.Checked And TextBox2.Text = "" Then
            MessageBox.Show("Select a Destination Cell.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            TextBox2.Focus()
            Exit Sub
        End If

        If RadioButton10.Checked And IsValidExcelCellReference(TextBox2.Text) = False Then
            MessageBox.Show("Select a Valid Destination Cell.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            TextBox2.Focus()
            Exit Sub
        End If

        If CheckBox1.Checked = True Then
            workSheet.Copy(After:=workBook.Sheets(workSheet.Name))
            workSheet2.Activate()
        End If

        Dim rng2Address As String

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

        Dim PrimaryRow As Integer = 0

        For Each lbl In labels3
            If lbl.Text = "    Primary Key" Then
                IsPrimary = True
                PrimaryRow = labels3.IndexOf(lbl) + 1
                Exit For
            End If
        Next

        Active = Active And IsPrimary

        If Active = True Then

            Dim cRng As Excel.Range
            cRng = workSheet.Range(rng.Cells(PrimaryRow, 1), rng.Cells(PrimaryRow, rng.Columns.Count))

            Dim Arr1(0) As Object
            Dim Arr2(0) As Integer

            Dim Index1 As Integer = 0
            Dim Index2 As Integer = 0

            Arr1(0) = cRng.Cells(1, 1).Value
            Arr2(0) = 1

            Dim CaseSensitive As Boolean

            If CheckBox1.Checked = True Then
                CaseSensitive = True
            Else
                CaseSensitive = False
            End If

            If CheckBox3.Checked = True Then
                For j = 1 To cRng.Columns.Count
                    If cRng.Cells(1, j).Value IsNot Nothing Then
                        If Search(Arr1, cRng.Cells(1, j).Value, CaseSensitive) = False Then
                            Index1 = Index1 + 1
                            Index2 = Index2 + 1
                            ReDim Preserve Arr1(Index1)
                            ReDim Preserve Arr2(Index2)
                            Arr1(Index1) = cRng.Cells(1, j).Value
                            Arr2(Index2) = j
                        End If
                    End If
                Next
            Else
                For j = 1 To cRng.Columns.Count
                    If Search(Arr1, cRng.Cells(1, j).Value, CaseSensitive) = False Then
                        Index1 = Index1 + 1
                        Index2 = Index2 + 1
                        ReDim Preserve Arr1(Index1)
                        ReDim Preserve Arr2(Index2)
                        Arr1(Index1) = cRng.Cells(1, j).Value
                        Arr2(Index2) = j
                    End If
                Next
            End If

            rng2 = workSheet2.Range(rng2.Cells(1, 1), rng2.Cells(rng.Rows.Count, UBound(Arr1) + 1))
            rng2Address = rng2.Address

            Dim UniQueArr() As Integer = GetUniques(rng, CaseSensitive, False)

            If Overlap(excelApp, workSheet, workSheet2, rng, rng2) = True Then

                Dim ValueArr(rng.Rows.Count - 1, rng.Columns.Count - 1) As Object

                For i = 1 To rng.Rows.Count
                    For j = 1 To rng.Columns.Count
                        ValueArr(i - 1, j - 1) = rng.Cells(i, j).Value
                    Next
                Next

                Dim FontNames(rng.Rows.Count - 1, rng.Columns.Count - 1) As String
                Dim FontSizes(rng.Rows.Count - 1, rng.Columns.Count - 1) As Single

                Dim FontBolds(rng.Rows.Count - 1, rng.Columns.Count - 1) As Boolean
                Dim Fontitalics(rng.Rows.Count - 1, rng.Columns.Count - 1) As Boolean
                Dim Red1s(rng.Rows.Count - 1, rng.Columns.Count - 1) As Integer
                Dim Green1s(rng.Rows.Count - 1, rng.Columns.Count - 1) As Integer
                Dim Blue1s(rng.Rows.Count - 1, rng.Columns.Count - 1) As Integer
                Dim Red2s(rng.Rows.Count - 1, rng.Columns.Count - 1) As Integer
                Dim Green2s(rng.Rows.Count - 1, rng.Columns.Count - 1) As Integer
                Dim Blue2s(rng.Rows.Count - 1, rng.Columns.Count - 1) As Integer

                If CheckBox4.Checked = True Then

                    For i = LBound(FontSizes, 1) To UBound(FontSizes, 1)
                        For j = LBound(FontSizes, 2) To UBound(FontSizes, 2)

                            Dim cell As Excel.Range = rng.Cells(i + 1, j + 1)
                            Dim font As Excel.Font = cell.Font

                            If IsDBNull(font.Name) = False Then
                                FontNames(i, j) = font.Name
                            Else
                                FontNames(i, j) = "Calibri"
                            End If

                            FontBolds(i, j) = cell.Font.Bold
                            Fontitalics(i, j) = cell.Font.Italic

                            If IsDBNull(font.Size) = False Then
                                Dim fontSize As Single = Convert.ToSingle(font.Size)
                                FontSizes(i, j) = fontSize
                            Else
                                FontSizes(i, j) = 11
                            End If

                            If IsDBNull(cell.Interior.Color) Then
                                Red1s(i, j) = 0
                                Green1s(i, j) = 0
                                Blue1s(i, j) = 0
                            Else
                                Dim colorValue1 As Long = CLng(cell.Interior.Color)
                                Dim red1 As Integer = colorValue1 Mod 256
                                Dim green1 As Integer = (colorValue1 \ 256) Mod 256
                                Dim blue1 As Integer = (colorValue1 \ 256 \ 256) Mod 256
                                Red1s(i, j) = red1
                                Green1s(i, j) = green1
                                Blue1s(i, j) = blue1
                            End If

                            If IsDBNull(cell.Font.Color) Then
                                Red2s(i, j) = 0
                                Green2s(i, j) = 0
                                Blue2s(i, j) = 0
                            Else
                                Dim colorValue2 As Long = CLng(cell.Font.Color)
                                Dim red2 As Integer = colorValue2 Mod 256
                                Dim green2 As Integer = (colorValue2 \ 256) Mod 256
                                Dim blue2 As Integer = (colorValue2 \ 256 \ 256) Mod 256
                                Red2s(i, j) = red2
                                Green2s(i, j) = green2
                                Blue2s(i, j) = blue2
                            End If

                        Next
                    Next
                End If

                rng.ClearContents()
                rng.ClearFormats()

                If CheckBox4.Checked = True Then

                    For i = 1 To rng2.Rows.Count
                        For j = 1 To rng2.Columns.Count
                            Dim x As Integer = i - 1
                            Dim y As Integer = Arr2(j - 1) - 1

                            rng2.Cells(i, j).Font.Name = FontNames(x, y)
                            rng2.Cells(i, j).Font.Size = FontSizes(x, y)

                            If FontBolds(x, y) Then rng2.Cells(i, j).Font.Bold = True
                            If Fontitalics(x, y) Then rng2.Cells(i, j).Font.Italic = True

                            rng2.Cells(i, j).Interior.Color = System.Drawing.Color.FromArgb(Red1s(x, y), Green1s(x, y), Blue1s(x, y))

                            rng2.Cells(i, j).Font.Color = System.Drawing.Color.FromArgb(Red2s(x, y), Green2s(x, y), Blue2s(x, y))

                            Dim targetCell As Excel.Range = rng2.Cells(i, j)

                            For k As Integer = 7 To 11
                                targetCell.Borders(k).LineStyle = Excel.XlLineStyle.xlContinuous
                                targetCell.Borders(k).Color = System.Drawing.Color.Black.ToArgb()
                            Next

                        Next
                    Next

                End If

                For i = 1 To rng.Rows.Count
                    If i <> PrimaryRow Then
                        Dim widthFlag As Boolean
                        Dim separator As String = " "
                        Dim Flag As String = ""

                        If labels3(i - 1).Text = "    Comma" Then
                            separator = ", "
                            widthFlag = True
                            Flag = "a"
                        ElseIf labels3(i - 1).Text = "    Colon" Then
                            separator = ": "
                            widthFlag = True
                            Flag = "a"
                        ElseIf labels3(i - 1).Text = "    Semicolon" Then
                            separator = "; "
                            widthFlag = True
                            Flag = "a"
                        ElseIf labels3(i - 1).Text = "    Space" Then
                            separator = " "
                            widthFlag = True
                            Flag = "b"
                        ElseIf labels3(i - 1).Text = "    Nothing" Then
                            separator = ""
                            widthFlag = True
                            Flag = c
                        ElseIf labels3(i - 1).Text = "    New Line" Then
                            separator = vbNewLine
                            widthFlag = True
                            Flag = "b"
                        Else
                            widthFlag = False
                            Flag = labels3(i - 1).Text
                        End If

                        For k = LBound(Arr1) To UBound(Arr1)

                            Dim concatenatedValue As String = ""
                            Dim OperatedValue As Object
                            Dim Valuess(0) As Object
                            Dim indx As Integer = -1

                            For j = 1 To rng.Columns.Count

                                Dim DuplicateCondition As Boolean
                                If CheckBox6.Checked = True Then
                                    DuplicateCondition = SearchInArray(UniQueArr, j)
                                Else
                                    DuplicateCondition = True
                                End If

                                If widthFlag = True Then

                                    If ValueArr(i - 1, j - 1) IsNot Nothing And DuplicateCondition Then
                                        Dim Matched As Boolean
                                        If CheckBox1.Checked = True Then
                                            Matched = ValueArr(PrimaryRow - 1, j - 1) = Arr1(k)
                                        Else
                                            Matched = LCase(ValueArr(PrimaryRow - 1, j - 1)) = LCase(Arr1(k))
                                        End If

                                        If Matched = True Then
                                            concatenatedValue = concatenatedValue & ValueArr(i - 1, j - 1) & separator
                                        End If
                                    End If

                                Else
                                    If ValueArr(i - 1, j - 1) IsNot Nothing And DuplicateCondition Then
                                        Dim Matched As Boolean
                                        If CheckBox1.Checked = True Then
                                            Matched = ValueArr(PrimaryRow - 1, j - 1) = Arr1(k)
                                        Else
                                            Matched = LCase(ValueArr(PrimaryRow - 1, j - 1)) = LCase(Arr1(k))
                                        End If
                                        If Matched = True Then
                                            indx = indx + 1
                                            ReDim Preserve Valuess(indx)
                                            Valuess(indx) = ValueArr(i - 1, j - 1)
                                        End If

                                    End If
                                End If
                            Next

                            OperatedValue = Operation(Valuess, Flag)

                            If Flag = "a" Then
                                If concatenatedValue <> "" Then
                                    concatenatedValue = Mid(concatenatedValue, 1, Len(concatenatedValue) - 2)
                                End If
                            ElseIf Flag = "b" Then
                                If concatenatedValue <> "" Then
                                    concatenatedValue = Mid(concatenatedValue, 1, Len(concatenatedValue) - 1)
                                End If
                            End If

                            If widthFlag = True Then
                                rng2.Cells(i, k + 1).value = concatenatedValue
                            Else
                                rng2.Cells(i, k + 1).value = OperatedValue
                            End If

                        Next
                    Else
                        For k = LBound(Arr1) To UBound(Arr1)
                            rng2.Cells(i, k + 1).value = Arr1(k)
                        Next

                    End If

                Next

            Else

                For i = 1 To rng.Rows.Count
                    If i <> PrimaryRow Then
                        Dim widthFlag As Boolean
                        Dim separator As String = " "
                        Dim Flag As String = ""

                        If labels3(i - 1).Text = "    Comma" Then
                            separator = ", "
                            widthFlag = True
                            Flag = "a"
                        ElseIf labels3(i - 1).Text = "    Colon" Then
                            separator = ": "
                            widthFlag = True
                            Flag = "a"
                        ElseIf labels3(i - 1).Text = "    Semicolon" Then
                            separator = "; "
                            widthFlag = True
                            Flag = "a"
                        ElseIf labels3(i - 1).Text = "    Space" Then
                            separator = " "
                            widthFlag = True
                            Flag = "b"
                        ElseIf labels3(i - 1).Text = "    Nothing" Then
                            separator = ""
                            widthFlag = True
                            Flag = c
                        ElseIf labels3(i - 1).Text = "    New Line" Then
                            separator = vbNewLine
                            widthFlag = True
                            Flag = "b"
                        Else
                            widthFlag = False
                            Flag = labels3(i - 1).Text
                        End If

                        For k = LBound(Arr1) To UBound(Arr1)

                            Dim concatenatedValue As String = ""
                            Dim OperatedValue As Object
                            Dim Valuess(0) As Object
                            Dim indx As Integer = -1

                            For j = 1 To rng.Columns.Count

                                Dim DuplicateCondition As Boolean
                                If CheckBox6.Checked = True Then
                                    DuplicateCondition = SearchInArray(UniQueArr, j)
                                Else
                                    DuplicateCondition = True
                                End If

                                If widthFlag = True Then
                                    If rng.Cells(i, j).Value IsNot Nothing And DuplicateCondition Then
                                        Dim Matched As Boolean
                                        If CheckBox1.Checked = True Then
                                            Matched = rng.Cells(PrimaryRow, j).Value = Arr1(k)
                                        Else
                                            Matched = LCase(rng.Cells(PrimaryRow, j).Value) = LCase(Arr1(k))
                                        End If

                                        If Matched = True Then
                                            concatenatedValue = concatenatedValue & rng.Cells(i, j).Value & separator
                                        End If
                                    End If

                                Else
                                    If rng.Cells(i, j).Value IsNot Nothing And DuplicateCondition Then
                                        Dim Matched As Boolean
                                        If CheckBox1.Checked = True Then
                                            Matched = rng.Cells(PrimaryRow, j).Value = Arr1(k)
                                        Else
                                            Matched = LCase(rng.Cells(PrimaryRow, j).Value) = LCase(Arr1(k))
                                        End If
                                        If Matched = True Then
                                            indx = indx + 1
                                            ReDim Preserve Valuess(indx)
                                            Valuess(indx) = rng.Cells(i, j).Value
                                        End If
                                    End If

                                End If
                            Next

                            OperatedValue = Operation(Valuess, Flag)

                            If Flag = "a" Then
                                If concatenatedValue <> "" Then
                                    concatenatedValue = Mid(concatenatedValue, 1, Len(concatenatedValue) - 2)
                                End If
                            ElseIf Flag = "b" Then
                                If concatenatedValue <> "" Then
                                    concatenatedValue = Mid(concatenatedValue, 1, Len(concatenatedValue) - 1)
                                End If
                            End If

                            If widthFlag = True Then
                                rng2.Cells(i, k + 1).value = concatenatedValue

                            Else
                                rng2.Cells(i, k + 1).value = OperatedValue
                            End If
                            If CheckBox4.Checked = True Then
                                rng.Cells(i, Arr2(k)).Copy
                                rng2.Cells(i, k + 1).PasteSpecial(Excel.XlPasteType.xlPasteFormats)
                                rng2 = workSheet2.Range(rng2Address)
                            End If
                        Next
                        excelApp.CutCopyMode = Excel.XlCutCopyMode.xlCopy
                    Else
                        For k = LBound(Arr1) To UBound(Arr1)
                            rng2.Cells(i, k + 1).value = Arr1(k)
                            If CheckBox4.Checked = True Then
                                rng.Cells(i, Arr2(k)).Copy
                                rng2.Cells(i, k + 1).PasteSpecial(Excel.XlPasteType.xlPasteFormats)
                                rng2 = workSheet2.Range(rng2Address)
                            End If
                        Next
                        excelApp.CutCopyMode = Excel.XlCutCopyMode.xlCopy

                    End If

                Next

                For j = 1 To rng2.Columns.Count
                    rng2.Cells(rng2.Rows.Count, j).Borders(9).LineStyle = rng.Cells(rng.Rows.Count, j).Borders(9).LineStyle
                    rng2.Cells(rng2.Rows.Count, j).Borders(9).Color = rng.Cells(rng.Rows.Count, j).Borders(9).Color
                    rng2.Cells(rng2.Rows.Count, j).Borders(9).weight = rng.Cells(rng.Rows.Count, j).Borders(9).weight
                Next

            End If

            Dim columnNum As Integer
            For j = 1 To rng2.Columns.Count
                columnNum = rng2.Cells(1, j).column
                workSheet2.Columns(columnNum).Autofit
            Next

            Me.Close()

            rng2.Select()

        End If


    End Sub

    Private Sub CheckBox4_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox4.CheckedChanged

        Try
            Call Display()

        Catch ex As Exception

        End Try

    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged

        Try
            Call Display()

        Catch ex As Exception

        End Try

    End Sub

    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged

        Try
            Call Display()

        Catch ex As Exception

        End Try

    End Sub

    Private Sub CheckBox6_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox6.CheckedChanged

        Try
            Call Display()

        Catch ex As Exception

        End Try

    End Sub

    Private Sub AutoSelection_Click(sender As Object, e As EventArgs) Handles AutoSelection.Click
        Try
            FocusedTextBox = 1
            Me.Hide()

            excelApp = Globals.ThisAddIn.Application
            workBook = excelApp.ActiveWorkbook

            Dim userInput As Excel.Range = excelApp.InputBox("Select a range", Type:=8)
            rng = userInput


            Dim sheetName As String
            sheetName = Split(rng.Address(True, True, Excel.XlReferenceStyle.xlA1, True), "]")(1)
            sheetName = Split(sheetName, "!")(0)

            If Mid(sheetName, Len(sheetName), 1) = "'" Then
                sheetName = Mid(sheetName, 1, Len(sheetName) - 1)
            End If

            workSheet = workBook.Worksheets(sheetName)
            workSheet.Activate()

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

    Private Sub Selection_Click(sender As Object, e As EventArgs) Handles Selection.Click
        Try
            FocusedTextBox = 1
            Me.Hide()

            excelApp = Globals.ThisAddIn.Application
            workBook = excelApp.ActiveWorkbook

            Dim userInput As Excel.Range = excelApp.InputBox("Select a range", Type:=8)
            rng = userInput

            Dim sheetName As String
            sheetName = Split(rng.Address(True, True, Excel.XlReferenceStyle.xlA1, True), "]")(1)
            sheetName = Split(sheetName, "!")(0)

            If Mid(sheetName, Len(sheetName), 1) = "'" Then
                sheetName = Mid(sheetName, 1, Len(sheetName) - 1)
            End If

            workSheet = workBook.Worksheets(sheetName)
            workSheet.Activate()

            rng.Select()

            TextBox1.Text = rng.Address

            Me.Show()
            TextBox1.Focus()

        Catch ex As Exception

            Me.Show()
            TextBox1.Focus()

        End Try
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        Try
            excelApp = Globals.ThisAddIn.Application
            workBook = excelApp.ActiveWorkbook
            workSheet2 = workBook.ActiveSheet

            TextBox2.SelectionStart = TextBox2.Text.Length
            TextBox2.ScrollToCaret()

            rng2 = workSheet2.Range(TextBox2.Text)
            rng2.Select()

        Catch ex As Exception

        End Try
    End Sub

    Private Sub TextBox1_GotFocus(sender As Object, e As EventArgs) Handles TextBox1.GotFocus
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

    Private Sub AutoSelection_GotFocus(sender As Object, e As EventArgs) Handles AutoSelection.GotFocus
        Try
            FocusedTextBox = 1

        Catch ex As Exception

        End Try
    End Sub

    Private Sub Selection_GotFocus(sender As Object, e As EventArgs) Handles Selection.GotFocus
        Try
            FocusedTextBox = 1

        Catch ex As Exception

        End Try
    End Sub

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click
        Try
            FocusedTextBox = 2
            Me.Hide()

            excelApp = Globals.ThisAddIn.Application
            workBook = excelApp.ActiveWorkbook

            Dim userInput As Excel.Range = excelApp.InputBox("Select a range", Type:=8)
            rng2 = userInput


            Dim sheetName As String
            sheetName = Split(rng2.Address(True, True, Excel.XlReferenceStyle.xlA1, True), "]")(1)
            sheetName = Split(sheetName, "!")(0)

            If Mid(sheetName, Len(sheetName), 1) = "'" Then
                sheetName = Mid(sheetName, 1, Len(sheetName) - 1)
            End If

            workSheet2 = workBook.Worksheets(sheetName)
            workSheet2.Activate()

            rng2.Select()

            TextBox2.Text = rng2.Address

            Me.Show()
            TextBox2.Focus()

        Catch ex As Exception

            Me.Show()
            TextBox2.Focus()

        End Try
    End Sub

    Private Sub PictureBox3_GotFocus(sender As Object, e As EventArgs) Handles PictureBox3.GotFocus
        Try
            FocusedTextBox = 2

        Catch ex As Exception

        End Try
    End Sub

    Private Sub Button2_MouseEnter(sender As Object, e As EventArgs) Handles Button2.MouseEnter
        Try

            Button2.BackColor = Color.FromArgb(65, 105, 225)
            Button2.ForeColor = Color.FromArgb(255, 255, 255)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Try
            Me.Close()
        Catch ex As Exception

        End Try

    End Sub

    Private Sub Button1_MouseEnter(sender As Object, e As EventArgs) Handles Button1.MouseEnter

        Try

            Button1.BackColor = Color.FromArgb(65, 105, 225)
            Button1.ForeColor = Color.FromArgb(255, 255, 255)
        Catch ex As Exception

        End Try

    End Sub

    Private Sub Button2_MouseLeave(sender As Object, e As EventArgs) Handles Button2.MouseLeave
        Try

            Button2.BackColor = Color.FromArgb(255, 255, 255)
            Button2.ForeColor = Color.FromArgb(70, 70, 70)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Button1_MouseLeave(sender As Object, e As EventArgs) Handles Button1.MouseLeave
        Try

            Button1.BackColor = Color.FromArgb(255, 255, 255)
            Button1.ForeColor = Color.FromArgb(70, 70, 70)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub AutoSelection_KeyDown(sender As Object, e As KeyEventArgs) Handles AutoSelection.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub Button1_KeyDown(sender As Object, e As KeyEventArgs) Handles Button1.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub Button2_KeyDown(sender As Object, e As KeyEventArgs) Handles Button2.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub CheckBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles CheckBox1.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub CheckBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles CheckBox2.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub CheckBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles CheckBox3.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub CheckBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles CheckBox4.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub CheckBox5_KeyDown(sender As Object, e As KeyEventArgs) Handles CheckBox5.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub CheckBox6_KeyDown(sender As Object, e As KeyEventArgs) Handles CheckBox6.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub ComboBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox1.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub CustomGroupBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles CustomGroupBox1.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub CustomGroupBox10_KeyDown(sender As Object, e As KeyEventArgs) Handles CustomGroupBox10.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub CustomGroupBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles CustomGroupBox2.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub CustomGroupBox6_KeyDown(sender As Object, e As KeyEventArgs) Handles CustomGroupBox6.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub CustomGroupBox7_KeyDown(sender As Object, e As KeyEventArgs) Handles CustomGroupBox7.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub CustomPanel2_KeyDown(sender As Object, e As KeyEventArgs) Handles CustomPanel2.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub Info_KeyDown(sender As Object, e As KeyEventArgs)
        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub Label1_KeyDown(sender As Object, e As KeyEventArgs) Handles Label1.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub Label2_KeyDown(sender As Object, e As KeyEventArgs) Handles Label2.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub Label3_KeyDown(sender As Object, e As KeyEventArgs) Handles Label3.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub Label4_KeyDown(sender As Object, e As KeyEventArgs) Handles Label4.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub Label5_KeyDown(sender As Object, e As KeyEventArgs) Handles Label5.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub PictureBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles PictureBox2.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub PictureBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles PictureBox3.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub RadioButton10_KeyDown(sender As Object, e As KeyEventArgs) Handles RadioButton10.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub RadioButton3_KeyDown(sender As Object, e As KeyEventArgs) Handles RadioButton3.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub Selection_KeyDown(sender As Object, e As KeyEventArgs) Handles Selection.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        Try
            If e.KeyCode = Keys.Enter Then

                Call Button2_Click(sender, e)

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged


    End Sub

    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged
        Try

            If RadioButton3.Checked = True Then

                excelApp = Globals.ThisAddIn.Application
                workBook = excelApp.ActiveWorkbook
                workSheet2 = workSheet

                rng2 = rng

                rng2.Select()

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub RadioButton10_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton10.CheckedChanged

        If RadioButton10.Checked = True Then

            Label3.Visible = True
            TextBox2.Visible = True
            PictureBox2.Visible = True
            PictureBox3.Visible = True
            TextBox2.Focus()
        Else
            TextBox2.Clear()
            Label3.Visible = False
            TextBox2.Visible = False
            PictureBox2.Visible = False
            PictureBox3.Visible = False

        End If
    End Sub

    Private Sub Form23_Merge_Duplicate_Columns_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        form_flag = False
    End Sub

    Private Sub Form23_Merge_Duplicate_Columns_Disposed(sender As Object, e As EventArgs) Handles Me.Disposed
        form_flag = False
    End Sub

    Private Sub Form23_Merge_Duplicate_Columns_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        Me.Focus()
        Me.BringToFront()
        Me.Activate()
        Me.BeginInvoke(New System.Action(Sub()
                                             TextBox1.Text = rng.Address
                                             SetWindowPos(Me.Handle, New IntPtr(HWND_TOPMOST), 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOSIZE)
                                         End Sub))
    End Sub

    Private Sub Info_Click(sender As Object, e As EventArgs)

    End Sub
End Class