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
Imports System.Text.RegularExpressions
Imports System.ComponentModel

Public Class Form3

    Public WithEvents excelApp As Excel.Application

    Public workbook As Excel.Workbook
    Public workbook2 As Excel.Workbook

    Public worksheet As Excel.Worksheet
    Public worksheet2 As Excel.Worksheet
    Public OpenSheet As Excel.Worksheet

    Public rng As Excel.Range
    Public rng2 As Excel.Range
    Public FocusedTextBox As Integer
    Public Opened As Integer

    Public Form4Open As Integer
    Public Workbook2Opened As Boolean

    Public TextBoxChanged As Boolean


    Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As IntPtr, ByVal hWndInsertAfter As IntPtr, ByVal X As Integer, ByVal Y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal uFlags As UInteger) As Boolean
    Private Const SWP_NOMOVE As UInteger = &H2
    Private Const SWP_NOSIZE As UInteger = &H1
    Private Const SWP_NOACTIVATE As UInteger = &H10
    Private Const HWND_TOPMOST As Integer = -1


    Private Function IsValidExcelCellReference(cellReference As String) As Boolean

        Dim cellPattern As String = "(\$?[A-Z]+\$?[0-9]+)"
        Dim referencePattern As String = "^" + cellPattern + "(:" + cellPattern + ")?$"

        Dim regex As New Regex(referencePattern)

        Dim refArr() As String = Split(cellReference, "!")

        Dim reference As String = refArr(UBound(refArr))

        If regex.IsMatch(reference) Then
            Return True
        Else
            Return False
        End If

    End Function
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

            Dim height As Single
            Dim width As Single

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

                        panel2.Controls.Add(label)
                    Next
                Next

                panel2.AutoScroll = True

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub DestinationChange()

        Try
            If RadioButton1.Checked = True Then
                If Form4Open = 1 Then
                    If Me.Workbook2Opened = True Then
                        workbook2.Close()
                        workbook.Activate()
                    End If
                    workbook2 = workbook
                    Form4Open = 0
                End If
                TextBox2.Visible = True
                PictureBox2.Visible = True
                TextBox2.Location = New System.Drawing.Point(121, 7)
                PictureBox2.Location = New System.Drawing.Point(226, 7)
                TextBox2.Focus()
            Else
                TextBox2.Clear()
            End If

            If RadioButton4.Checked = True Then

                If Me.Form4Open = 1 Then
                    If Me.Workbook2Opened = True Then
                        workbook2.Close()
                        workbook.Activate()
                    End If
                    workbook2 = workbook
                    Me.Form4Open = 0
                End If
                TextBox2.Visible = True
                PictureBox2.Visible = True
                TextBox2.Location = New System.Drawing.Point(121, 30)
                PictureBox2.Location = New System.Drawing.Point(226, 30)

                Dim ws As Excel.Worksheet = CType(workbook.Worksheets.Add(), Excel.Worksheet)
                TextBox2.Focus()
            Else
                TextBox2.Clear()
            End If

            If RadioButton5.Checked = True And Form4Open = 0 Then
                TextBox2.Visible = False
                PictureBox2.Visible = False
                Dim MyForm4 As New Form4
                MyForm4.excelApp = Me.excelApp
                MyForm4.workbook = Me.workbook
                MyForm4.worksheet = Me.worksheet
                MyForm4.OpenSheet = Me.OpenSheet
                MyForm4.rng = Me.rng
                MyForm4.Opened = Me.Opened
                MyForm4.FocusedTextBox = Me.FocusedTextBox
                MyForm4.TextBoxChanged = Me.TextBoxChanged
                MyForm4.Form4Open = Me.Form4Open
                MyForm4.Workbook2Opened = False
                If Me.RadioButton3.Checked = True Then
                    MyForm4.GB6 = 3
                ElseIf Me.RadioButton2.Checked = True Then
                    MyForm4.GB6 = 2
                Else
                    MyForm4.GB6 = 0
                End If
                If Me.CheckBox1.Checked = True Then
                    MyForm4.CB1 = 1
                Else
                    MyForm4.CB1 = 0
                End If
                If Me.CheckBox2.Checked = True Then
                    MyForm4.CB2 = 1
                Else
                    MyForm4.CB2 = 0
                End If
                Me.Close()
                MyForm4.Show()

            End If

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

    Private Sub PictureBox8_Click(sender As Object, e As EventArgs) Handles PictureBox8.Click

        Try
            FocusedTextBox = 1
            Me.Hide()


            Dim userInput As Excel.Range = excelApp.InputBox("Select a range", Type:=8)
            rng = userInput

            Dim sheetName As String
            sheetName = Split(rng.Address(True, True, Excel.XlReferenceStyle.xlA1, True), "]")(1)
            sheetName = Split(sheetName, "!")(0)

            If Mid(sheetName, Len(sheetName), 1) = "'" Then
                sheetName = Mid(sheetName, 1, Len(sheetName) - 1)
            End If

            worksheet = workbook.Worksheets(sheetName)
            worksheet.Activate()

            rng.Select()

            If worksheet.Name <> OpenSheet.Name Then
                TextBox1.Text = worksheet.Name & "!" & rng.Address
            Else
                TextBox1.Text = rng.Address
            End If

            Me.Show()
            TextBox1.Focus()

        Catch ex As Exception

            Me.Show()
            TextBox1.Focus()

        End Try

    End Sub

    Private Sub PictureBox4_Click(sender As Object, e As EventArgs) Handles PictureBox4.Click

        Try
            FocusedTextBox = 1
            Me.Hide()


            Dim userInput As Excel.Range = excelApp.InputBox("Select a range", Type:=8)
            rng = userInput

            Dim sheetName As String

            sheetName = Split(rng.Address(True, True, Excel.XlReferenceStyle.xlA1, True), "]")(1)
            sheetName = Split(sheetName, "!")(0)

            If Mid(sheetName, Len(sheetName), 1) = "'" Then
                sheetName = Mid(sheetName, 1, Len(sheetName) - 1)
            End If

            worksheet = workbook.Worksheets(sheetName)
            worksheet.Activate()

            rng.Select()

            Dim activeRange As Excel.Range = excelApp.ActiveCell

            Dim startRow As Integer = activeRange.Row
            Dim startColumn As Integer = activeRange.Column
            Dim endRow As Integer = activeRange.Row
            Dim endColumn As Integer = activeRange.Column

            'Find the upper boundary
            Do While startRow > 1 AndAlso Not IsNothing(worksheet.Cells(startRow - 1, startColumn).Value)
                startRow -= 1
            Loop

            'Find the lower boundary
            Do While Not IsNothing(worksheet.Cells(endRow + 1, endColumn).Value)
                endRow += 1
            Loop

            'Find the left boundary
            Do While startColumn > 1 AndAlso Not IsNothing(worksheet.Cells(startRow, startColumn - 1).Value)
                startColumn -= 1
            Loop

            'Find the right boundary
            Do While Not IsNothing(worksheet.Cells(endRow, endColumn + 1).Value)
                endColumn += 1
            Loop

            'Select the determined range
            rng = worksheet.Range(worksheet.Cells(startRow, startColumn), worksheet.Cells(endRow, endColumn))

            rng.Select()

            If worksheet.Name <> OpenSheet.Name Then
                TextBox1.Text = worksheet.Name & "!" & rng.Address
            Else
                TextBox1.Text = rng.Address
            End If

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
            FocusedTextBox = 2
            Me.Hide()

            Dim userInput As Excel.Range = excelApp.InputBox("Select a Cell.", Type:=8)
            rng2 = userInput

            Dim sheetName As String
            sheetName = Split(rng2.Address(True, True, Excel.XlReferenceStyle.xlA1, True), "]")(1)
            sheetName = Split(sheetName, "!")(0)

            If Mid(sheetName, Len(sheetName), 1) = "'" Then
                sheetName = Mid(sheetName, 1, Len(sheetName) - 1)
            End If

            worksheet2 = workbook.Worksheets(sheetName)
            worksheet2.Activate()

            rng2.Select()

            If worksheet2.Name <> OpenSheet.Name Then
                TextBox2.Text = worksheet2.Name & "!" & rng2.Address
            Else
                TextBox2.Text = rng2.Address
            End If

            Me.Show()
            TextBox2.Focus()

        Catch ex As Exception

            Me.Show()
            TextBox2.Focus()

        End Try

    End Sub

    Private Sub btn_OK_Click(sender As Object, e As EventArgs) Handles btn_OK.Click

        Try
            If TextBox1.Text = "" Or IsValidExcelCellReference(TextBox1.Text) = False Then
                MessageBox.Show("Enter a Valid Source Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                worksheet.Activate()
                rng.Select()
                Exit Sub
            End If

            If RadioButton2.Checked = False And RadioButton3.Checked = False Then
                MessageBox.Show("Select a Paste Option.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                worksheet.Activate()
                rng.Select()
                Exit Sub
            End If

            If RadioButton1.Checked = False And RadioButton4.Checked = False And RadioButton5.Checked = False Then
                MessageBox.Show("Select a Destination Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                worksheet.Activate()
                rng.Select()
                Exit Sub
            End If

            If (RadioButton1.Checked = True Or RadioButton4.Checked = True) Then
                If TextBox2.Text = "" Or IsValidExcelCellReference(TextBox2.Text) = False Then
                    MessageBox.Show("Select a Valid Destination Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    worksheet.Activate()
                    rng.Select()
                    Exit Sub
                End If
            End If


            If (RadioButton2.Checked = True Or RadioButton3.Checked = True) Then

                rng2 = worksheet2.Range(rng2.Cells(1, 1), rng2.Cells(rng.Columns.Count, rng.Rows.Count))
                Dim rng2Address As String = rng2.Address

                If CheckBox1.Checked = True Then
                    worksheet.Copy(After:=workbook.Sheets(worksheet.Name))
                End If

                worksheet2.Activate()

                If (Overlap(excelApp, worksheet, worksheet2, rng, rng2)) = False Then

                    If RadioButton3.Checked = True Then
                        For i = 1 To rng.Rows.Count
                            For j = 1 To rng.Columns.Count
                                rng2.Cells(j, i).Value = rng.Cells(i, j).Value
                                rng2 = worksheet2.Range(rng2Address)
                                If CheckBox2.Checked = True Then
                                    MsgBox(rng.Cells(i, j).Address)
                                    rng.Cells(i, j).Copy
                                    rng2.Cells(j, i).PasteSpecial(Excel.XlPasteType.xlPasteFormats)
                                    rng2 = worksheet2.Range(rng2Address)
                                    Dim sourceCell As Excel.Range = rng.Cells(i, j)
                                    Dim targetCell As Excel.Range = rng2.Cells(j, i)

                                    If sourceCell.Borders(7).LineStyle <> Excel.XlLineStyle.xlLineStyleNone Then
                                        targetCell.Borders(8).LineStyle = sourceCell.Borders(7).LineStyle
                                        targetCell.Borders(8).Color = sourceCell.Borders(7).Color
                                        targetCell.Borders(8).Weight = sourceCell.Borders(7).Weight
                                    Else
                                        targetCell.Borders(8).LineStyle = Excel.XlLineStyle.xlLineStyleNone
                                    End If

                                    If sourceCell.Borders(8).LineStyle <> Excel.XlLineStyle.xlLineStyleNone Then
                                        targetCell.Borders(7).LineStyle = sourceCell.Borders(8).LineStyle
                                        targetCell.Borders(7).Color = sourceCell.Borders(8).Color
                                        targetCell.Borders(7).Weight = sourceCell.Borders(8).Weight
                                    Else
                                        targetCell.Borders(7).LineStyle = Excel.XlLineStyle.xlLineStyleNone
                                    End If

                                    If sourceCell.Borders(9).LineStyle <> Excel.XlLineStyle.xlLineStyleNone Then
                                        targetCell.Borders(10).LineStyle = sourceCell.Borders(9).LineStyle
                                        targetCell.Borders(10).Color = sourceCell.Borders(9).Color
                                        targetCell.Borders(10).Weight = sourceCell.Borders(9).Weight
                                    Else
                                        targetCell.Borders(10).LineStyle = Excel.XlLineStyle.xlLineStyleNone
                                    End If

                                    If sourceCell.Borders(10).LineStyle <> Excel.XlLineStyle.xlLineStyleNone Then
                                        targetCell.Borders(9).LineStyle = sourceCell.Borders(10).LineStyle
                                        targetCell.Borders(9).Color = sourceCell.Borders(10).Color
                                        targetCell.Borders(9).Weight = sourceCell.Borders(10).Weight
                                    Else
                                        targetCell.Borders(9).LineStyle = Excel.XlLineStyle.xlLineStyleNone
                                    End If

                                End If
                            Next
                        Next

                    ElseIf RadioButton2.Checked = True Then
                        For i = 1 To rng.Rows.Count
                            For j = 1 To rng.Columns.Count
                                rng2.Cells(j, i).Value = "=" & rng.Cells(i, j).Address(True, True, Excel.XlReferenceStyle.xlA1, True)
                                If CheckBox2.Checked = True Then
                                    rng.Cells(i, j).Copy
                                    rng2.Cells(j, i).PasteSpecial(Excel.XlPasteType.xlPasteFormats)
                                    rng2 = worksheet2.Range(rng2Address)

                                    Dim sourceCell As Excel.Range = rng.Cells(i, j)
                                    Dim targetCell As Excel.Range = rng2.Cells(j, i)

                                    If sourceCell.Borders(7).LineStyle <> Excel.XlLineStyle.xlLineStyleNone Then
                                        targetCell.Borders(8).LineStyle = sourceCell.Borders(7).LineStyle
                                        targetCell.Borders(8).Color = sourceCell.Borders(7).Color
                                        targetCell.Borders(8).Weight = sourceCell.Borders(7).Weight
                                    Else
                                        targetCell.Borders(8).LineStyle = Excel.XlLineStyle.xlLineStyleNone
                                    End If

                                    If sourceCell.Borders(8).LineStyle <> Excel.XlLineStyle.xlLineStyleNone Then
                                        targetCell.Borders(7).LineStyle = sourceCell.Borders(8).LineStyle
                                        targetCell.Borders(7).Color = sourceCell.Borders(8).Color
                                        targetCell.Borders(7).Weight = sourceCell.Borders(8).Weight
                                    Else
                                        targetCell.Borders(7).LineStyle = Excel.XlLineStyle.xlLineStyleNone
                                    End If

                                    If sourceCell.Borders(9).LineStyle <> Excel.XlLineStyle.xlLineStyleNone Then
                                        targetCell.Borders(10).LineStyle = sourceCell.Borders(9).LineStyle
                                        targetCell.Borders(10).Color = sourceCell.Borders(9).Color
                                        targetCell.Borders(10).Weight = sourceCell.Borders(9).Weight
                                    Else
                                        targetCell.Borders(10).LineStyle = Excel.XlLineStyle.xlLineStyleNone
                                    End If

                                    If sourceCell.Borders(10).LineStyle <> Excel.XlLineStyle.xlLineStyleNone Then
                                        targetCell.Borders(9).LineStyle = sourceCell.Borders(10).LineStyle
                                        targetCell.Borders(9).Color = sourceCell.Borders(10).Color
                                        targetCell.Borders(9).Weight = sourceCell.Borders(10).Weight
                                    Else
                                        targetCell.Borders(9).LineStyle = Excel.XlLineStyle.xlLineStyleNone
                                    End If

                                End If
                            Next
                        Next
                    End If

                    excelApp.CutCopyMode = Excel.XlCutCopyMode.xlCopy

                Else

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

                                Dim targetCell As Excel.Range = rng2.Cells(j, i)

                                For k As Integer = 7 To 11
                                    targetCell.Borders(k).LineStyle = Excel.XlLineStyle.xlContinuous
                                    targetCell.Borders(k).Color = System.Drawing.Color.Black.ToArgb()
                                Next

                            Next
                        Next
                    End If
                End If

                rng2.Select()

                For j = 1 To rng2.Columns.Count
                    rng2.Columns(j).Autofit
                Next

                Me.Close()

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub btn_OK_MouseEnter(sender As Object, e As EventArgs) Handles btn_OK.MouseEnter

        Try

            btn_OK.ForeColor = Color.White
            btn_OK.BackColor = Color.FromArgb(76, 111, 174)

        Catch ex As Exception

        End Try

    End Sub

    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged

        Try
            If RadioButton3.Checked = True Then
                Call Display()
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged

        Try
            If RadioButton2.Checked = True Then
                Call Display()
            End If
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
            If TextBox1.Text <> "" And Form4Open = 0 Then
                worksheet = workbook.ActiveSheet
                TextBox1.SelectionStart = TextBox1.Text.Length
                TextBox1.ScrollToCaret()
                Dim rngArray() As String = Split(TextBox1.Text, "!")
                Dim rngAddress As String = rngArray(UBound(rngArray))
                rng = worksheet.Range(rngAddress)
                TextBoxChanged = True
                rng.Select()
                Call Display()
                TextBoxChanged = False
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

        Try
            If TextBox2.Text <> "" Then
                worksheet2 = workbook.ActiveSheet
                TextBox2.SelectionStart = TextBox2.Text.Length
                TextBox2.ScrollToCaret()
                Dim rng2Array() As String = Split(TextBox2.Text, "!")
                Dim rng2Address As String = rng2Array(UBound(rng2Array))
                rng2 = worksheet2.Range(rng2Address)
                TextBoxChanged = True
                rng2.Select()
                TextBoxChanged = False
            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles Me.Load

        Try

            AddHandler excelApp.SheetSelectionChange, AddressOf excelApp_SheetSelectionChange

            Opened = Opened + 1
            Me.KeyPreview = True

        Catch ex As Exception

        End Try

    End Sub

    Private Sub excelApp_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Excel.Range)

        Try

            Dim selectedRange As Excel.Range
            selectedRange = excelApp.Selection

            If TextBoxChanged = False Then
                If FocusedTextBox = 1 Then
                    worksheet = workbook.ActiveSheet
                    If worksheet.Name <> OpenSheet.Name Then
                        TextBox1.Text = worksheet.Name & "!" & selectedRange.Address
                    Else
                        TextBox1.Text = selectedRange.Address
                    End If
                    rng = selectedRange
                    TextBox1.Focus()

                ElseIf FocusedTextBox = 2 Then
                    worksheet2 = workbook.ActiveSheet
                    If worksheet2.Name <> OpenSheet.Name Then
                        TextBox2.Text = worksheet2.Name & "!" & selectedRange.Address
                    Else
                        TextBox2.Text = selectedRange.Address
                    End If
                    rng2 = selectedRange
                    TextBox2.Focus()
                End If
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

            If RadioButton4.Checked = True Then
                Call DestinationChange()
            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub RadioButton1_CheckedChanged_1(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged

        Try
            If RadioButton1.Checked = True Then
                Call DestinationChange()
            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub RadioButton5_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton5.CheckedChanged

        Try

            If RadioButton5.Checked = True Then
                Call DestinationChange()
            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub PictureBox8_GotFocus(sender As Object, e As EventArgs) Handles PictureBox8.GotFocus

        Try
            FocusedTextBox = 1
        Catch ex As Exception

        End Try

    End Sub

    Private Sub TextBox1_GotFocus(sender As Object, e As EventArgs) Handles TextBox1.GotFocus

        Try
            FocusedTextBox = 1
        Catch ex As Exception

        End Try

    End Sub

    Private Sub PictureBox4_GotFocus(sender As Object, e As EventArgs) Handles PictureBox4.GotFocus


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

    Private Sub PictureBox2_GotFocus(sender As Object, e As EventArgs) Handles PictureBox2.GotFocus

        Try
            FocusedTextBox = 2
        Catch ex As Exception

        End Try

    End Sub

    Private Sub RadioButton3_GotFocus(sender As Object, e As EventArgs) Handles RadioButton3.GotFocus

        Try
            FocusedTextBox = 0
        Catch ex As Exception

        End Try

    End Sub

    Private Sub RadioButton2_GotFocus(sender As Object, e As EventArgs) Handles RadioButton2.GotFocus

        Try
            FocusedTextBox = 0
        Catch ex As Exception

        End Try

    End Sub

    Private Sub PictureBox5_GotFocus(sender As Object, e As EventArgs) Handles PictureBox5.GotFocus
        Try
            FocusedTextBox = 0
        Catch ex As Exception

        End Try

    End Sub

    Private Sub PictureBox1_GotFocus(sender As Object, e As EventArgs) Handles PictureBox1.GotFocus

        Try
            FocusedTextBox = 0
        Catch ex As Exception

        End Try

    End Sub

    Private Sub CustomGroupBox3_GotFocus(sender As Object, e As EventArgs) Handles CustomGroupBox3.GotFocus
        Try
            FocusedTextBox = 0
        Catch ex As Exception

        End Try

    End Sub


    Private Sub CustomGroupBox6_GotFocus(sender As Object, e As EventArgs) Handles CustomGroupBox6.GotFocus
        Try
            FocusedTextBox = 0
        Catch ex As Exception

        End Try
    End Sub

    Private Sub CustomGroupBox5_GotFocus(sender As Object, e As EventArgs) Handles CustomGroupBox5.GotFocus
        Try
            FocusedTextBox = 0
        Catch ex As Exception

        End Try
    End Sub

    Private Sub RadioButton1_GotFocus(sender As Object, e As EventArgs) Handles RadioButton1.GotFocus
        Try
            FocusedTextBox = 0
        Catch ex As Exception

        End Try
    End Sub

    Private Sub RadioButton4_GotFocus(sender As Object, e As EventArgs) Handles RadioButton4.GotFocus
        Try
            FocusedTextBox = 0

        Catch ex As Exception

        End Try

    End Sub

    Private Sub RadioButton5_GotFocus(sender As Object, e As EventArgs) Handles RadioButton5.GotFocus

        Try
            FocusedTextBox = 0
        Catch ex As Exception

        End Try

    End Sub

    Private Sub CheckBox2_GotFocus(sender As Object, e As EventArgs) Handles CheckBox2.GotFocus
        Try
            FocusedTextBox = 0
        Catch ex As Exception

        End Try
    End Sub

    Private Sub CheckBox1_GotFocus(sender As Object, e As EventArgs) Handles CheckBox1.GotFocus
        Try
            FocusedTextBox = 0
        Catch ex As Exception

        End Try

    End Sub

    Private Sub ComboBox1_GotFocus(sender As Object, e As EventArgs) Handles ComboBox1.GotFocus
        Try
            FocusedTextBox = 0
        Catch ex As Exception

        End Try

    End Sub

    Private Sub CustomGroupBox1_GotFocus(sender As Object, e As EventArgs) Handles CustomGroupBox1.GotFocus
        Try
            FocusedTextBox = 0
        Catch ex As Exception

        End Try
    End Sub

    Private Sub CustomGroupBox2_GotFocus(sender As Object, e As EventArgs) Handles CustomGroupBox2.GotFocus
        Try
            FocusedTextBox = 0
        Catch ex As Exception

        End Try

    End Sub

    Private Sub CustomGroupBox4_GotFocus(sender As Object, e As EventArgs) Handles CustomGroupBox4.GotFocus
        Try
            FocusedTextBox = 0
        Catch ex As Exception

        End Try
    End Sub

    Private Sub panel1_GotFocus(sender As Object, e As EventArgs) Handles panel1.GotFocus
        Try
            FocusedTextBox = 0
        Catch ex As Exception

        End Try
    End Sub

    Private Sub panel2_GotFocus(sender As Object, e As EventArgs) Handles panel2.GotFocus
        Try
            FocusedTextBox = 0
        Catch ex As Exception

        End Try
    End Sub

    Private Sub PictureBox7_Click(sender As Object, e As EventArgs) Handles PictureBox7.Click

    End Sub

    Private Sub PictureBox7_GotFocus(sender As Object, e As EventArgs) Handles PictureBox7.GotFocus
        Try
            FocusedTextBox = 0
        Catch ex As Exception

        End Try
    End Sub

    Private Sub btn_OK_GotFocus(sender As Object, e As EventArgs) Handles btn_OK.GotFocus
        Try
            FocusedTextBox = 0
        Catch ex As Exception

        End Try
    End Sub

    Private Sub btn_cancel_GotFocus(sender As Object, e As EventArgs) Handles btn_cancel.GotFocus

        Try
            FocusedTextBox = 0
        Catch ex As Exception

        End Try

    End Sub

    Private Sub btn_cancel_Click(sender As Object, e As EventArgs) Handles btn_cancel.Click

        Try
            Me.Close()
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

    'Private Sub Form3_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
    '    form_flag = False
    'End Sub

    ''Private Sub Form3_Shown(sender As Object, e As EventArgs) Handles Me.Shown
    ''    Me.Focus()
    ''    Me.BringToFront()
    ''    Me.Activate()
    ''    Me.BeginInvoke(New System.Action(Sub()
    ''                                         TextBox1.Text = rng.Address
    ''                                         SetWindowPos(Me.Handle, New IntPtr(HWND_TOPMOST), 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOSIZE)
    ''                                     End Sub))
    ''End Sub

    'Private Sub Form3_Disposed(sender As Object, e As EventArgs) Handles Me.Disposed
    '    form_flag = False
    'End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown

        Try

            If e.KeyCode = Keys.Enter Then
                Call btn_OK_Click(sender, e)
            End If

        Catch ex As Exception

        End Try

    End Sub
    Private Sub Form3_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        form_flag = False
    End Sub

    Private Sub Form3_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        Me.Focus()
        Me.BringToFront()
        Me.Activate()
        Me.BeginInvoke(New System.Action(Sub()
                                             TextBox1.Text = rng.Address
                                             SetWindowPos(Me.Handle, New IntPtr(HWND_TOPMOST), 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOSIZE)
                                         End Sub))
    End Sub

    Private Sub Form3_Disposed(sender As Object, e As EventArgs) Handles Me.Disposed
        form_flag = False
    End Sub
End Class