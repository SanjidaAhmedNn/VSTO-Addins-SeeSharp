Imports System.Drawing
Imports System.Windows.Forms
Imports System.Reflection.Emit
Imports Microsoft.Office.Interop.Excel
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Threading
Imports System.Diagnostics
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Button
Imports System.IO
Imports Microsoft.Office.Interop


Public Class Form4
    Public WithEvents excelApp As Excel.Application
    Public workbook As Excel.Workbook
    Public workbook2 As Excel.Workbook
    Public worksheet As Excel.Worksheet
    Public worksheet1 As Excel.Worksheet
    Public worksheet2 As Excel.Worksheet
    Public rng As Excel.Range
    Public rng2 As Excel.Range
    Public FocusedTextBox As Integer
    Public Opened As Integer
    Public GB6 As Integer
    Dim ThisFocusedTextBox As Integer
    Public Form4Open As Integer

    Private Function IsValidExcelFile(filePath As String) As Boolean
        ' Check if the file exists.
        If Not File.Exists(filePath) Then
            Return False

        Else

            ' Get the file extension.
            Dim extension As String = Path.GetExtension(filePath)

            ' Check if the extension is a valid Excel extension.
            If extension = ".xls" OrElse extension = ".xlsx" OrElse extension = ".xlsm" Then
                Return True
            Else
                Return False
            End If
        End If

    End Function
    Private Sub Setup()

        If RadioButton1.Checked = True Then
            TextBox1.Enabled = True
            PictureBox8.Enabled = True
        Else
            TextBox1.Clear()
            TextBox1.Enabled = False
            PictureBox8.Enabled = False
        End If

        If RadioButton2.Checked = True Then
            TextBox2.Enabled = True
            PictureBox1.Enabled = True
            TextBox3.Enabled = True
            PictureBox2.Enabled = True
            Label1.Enabled = True
            PictureBox3.Enabled = True
        Else
            TextBox2.Clear()
            TextBox3.Clear()
            TextBox2.Enabled = False
            PictureBox1.Enabled = False
            TextBox3.Enabled = False
            PictureBox2.Enabled = False
            Label1.Enabled = False
            PictureBox3.Enabled = False
        End If
    End Sub

    'Worksheet.Name = "New Worksheet"
    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged

        If RadioButton1.Checked = True Then
            workbook2 = excelApp.Workbooks.Add()
            Me.Show()
            Me.TextBox1.Focus()
        End If

        Call Setup()

    End Sub



    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click

        ThisFocusedTextBox = 2

        Me.Hide()
        Dim openFileDialog As New OpenFileDialog()
        openFileDialog.Title = "Open Your File"
        openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm"

        If openFileDialog.ShowDialog() = DialogResult.OK Then
            Dim filePath As String = openFileDialog.FileName
            workbook2 = excelApp.Workbooks.Open(filePath)
            TextBox2.Text = filePath
            excelApp.Visible = True
        End If

        Me.Show()
        TextBox2.Focus()

    End Sub

    Private Sub PictureBox8_Click(sender As Object, e As EventArgs) Handles PictureBox8.Click

        Try
            ThisFocusedTextBox = 1
            Me.Hide()

            Dim userInput As Excel.Range = excelApp.InputBox("Select a Cell.", Type:=8)
            rng2 = userInput

            Dim sheetName As String
            sheetName = Split(rng2.Address(True, True, Excel.XlReferenceStyle.xlA1, True), "]")(1)
            sheetName = Split(sheetName, "!")(0)
            worksheet2 = workbook2.Worksheets(sheetName)
            worksheet2.Activate()

            rng2.Select()

            TextBox1.Text = rng2.Address

            Me.Show()
            TextBox1.Focus()

        Catch ex As Exception

            Me.Show()
            TextBox1.Focus()

        End Try


    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim MyForm3 As New Form3
        MyForm3.rng = Me.rng
        MyForm3.workbook = Me.workbook
        MyForm3.workbook2 = Me.workbook2
        MyForm3.worksheet = Me.worksheet
        MyForm3.worksheet2 = Me.worksheet2
        MyForm3.rng2 = Me.rng2
        MyForm3.TextBox1.Text = Me.rng.Address
        MyForm3.Form4Open = Me.Form4Open

        If Me.GB6 = 3 Then
            MyForm3.RadioButton3.Checked = True
        ElseIf Me.GB6 = 2 Then
            MyForm3.RadioButton2.Checked = True
        End If
        MyForm3.RadioButton5.Checked = True
        MyForm3.Opened = Me.Opened
        MyForm3.Show()
        Me.Close()

    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click

        Try
            ThisFocusedTextBox = 3
            Me.Hide()

            Dim userInput As Excel.Range = excelApp.InputBox("Select a Cell", Type:=8)
            rng2 = userInput


            Dim sheetName As String
            sheetName = Split(rng2.Address(True, True, Excel.XlReferenceStyle.xlA1, True), "]")(1)
            sheetName = Split(sheetName, "!")(0)
            worksheet2 = workbook2.Worksheets(sheetName)
            worksheet2.Activate()

            rng2.Select()

            TextBox3.Text = rng2.Address

            Me.Show()
            TextBox3.Focus()

        Catch ex As Exception

            Me.Show()
            TextBox3.Focus()

        End Try

    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged

        Call Setup()

    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

        If TextBox2.Text <> "" Then
            If IsValidExcelFile(TextBox2.Text) = True Then
                Dim filePath As String = TextBox2.Text
                workbook2 = excelApp.Workbooks.Open(filePath)
                excelApp.Visible = True
            End If
        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged

        Try

            If TextBox3.Text <> "" Then
                worksheet2 = workbook2.ActiveSheet
                rng2 = worksheet2.Range(TextBox3.Text)
                rng2.Select()
            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub Form4_Activated(sender As Object, e As EventArgs) Handles Me.Activated

        Try


            AddHandler excelApp.SheetSelectionChange, AddressOf excelApp_SheetSelectionChange

            Call Setup()

        Catch ex As Exception

        End Try

    End Sub

    Private Sub excelApp_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Excel.Range)

        Try

            Dim selectedRange As Excel.Range
            selectedRange = excelApp.Selection

            If ThisFocusedTextBox = 1 Then
                TextBox1.Text = selectedRange.Address
                worksheet2 = workbook2.ActiveSheet
                rng2 = selectedRange
                TextBox1.Focus()

            ElseIf ThisFocusedTextBox = 3 Then
                TextBox3.Text = selectedRange.Address
                worksheet2 = workbook2.ActiveSheet
                rng2 = selectedRange
                TextBox3.Focus()
            End If

        Catch ex As Exception

        End Try

    End Sub
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

        Try

            If TextBox1.Text <> "" Then
                worksheet2 = workbook2.ActiveSheet
                rng2 = worksheet2.Range(TextBox1.Text)
                rng2.Select()
            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub TextBox1_GotFocus(sender As Object, e As EventArgs) Handles TextBox1.GotFocus

        ThisFocusedTextBox = 1

    End Sub

    Private Sub TextBox2_GotFocus(sender As Object, e As EventArgs) Handles TextBox2.GotFocus
        Try
            ThisFocusedTextBox = 0
        Catch ex As Exception

        End Try

    End Sub

    Private Sub TextBox3_GotFocus(sender As Object, e As EventArgs) Handles TextBox3.GotFocus

        ThisFocusedTextBox = 3

    End Sub

    Private Sub PictureBox8_GotFocus(sender As Object, e As EventArgs) Handles PictureBox8.GotFocus

        Try
            ThisFocusedTextBox = 1
        Catch ex As Exception

        End Try

    End Sub

    Private Sub PictureBox2_GotFocus(sender As Object, e As EventArgs) Handles PictureBox2.GotFocus

        Try
            ThisFocusedTextBox = 3
        Catch ex As Exception

        End Try

    End Sub

    Private Sub RadioButton1_GotFocus(sender As Object, e As EventArgs) Handles RadioButton1.GotFocus
        Try
            ThisFocusedTextBox = 0
        Catch ex As Exception

        End Try
    End Sub

    Private Sub RadioButton2_GotFocus(sender As Object, e As EventArgs) Handles RadioButton2.GotFocus

        Try
            ThisFocusedTextBox = 0
        Catch ex As Exception

        End Try

    End Sub

    Private Sub PictureBox1_GotFocus(sender As Object, e As EventArgs) Handles PictureBox1.GotFocus

        Try
            ThisFocusedTextBox = 0
        Catch ex As Exception

        End Try

    End Sub

    Private Sub PictureBox3_GotFocus(sender As Object, e As EventArgs) Handles PictureBox3.GotFocus

        Try
            ThisFocusedTextBox = 0
        Catch ex As Exception

        End Try

    End Sub

    Private Sub Button1_GotFocus(sender As Object, e As EventArgs) Handles Button1.GotFocus

        Try
            ThisFocusedTextBox = 0
        Catch ex As Exception

        End Try

    End Sub

    Private Sub Button2_GotFocus(sender As Object, e As EventArgs) Handles Button2.GotFocus

        Try
            ThisFocusedTextBox = 0
        Catch ex As Exception

        End Try

    End Sub

    Private Sub Button3_GotFocus(sender As Object, e As EventArgs) Handles Button3.GotFocus

        Try
            ThisFocusedTextBox = 0
        Catch ex As Exception

        End Try

    End Sub
End Class