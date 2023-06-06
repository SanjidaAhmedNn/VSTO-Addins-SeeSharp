Imports System.Drawing
Imports System.Windows.Forms
Imports System.Reflection.Emit
Imports Microsoft.Office.Interop.Excel
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Threading
Imports System.Diagnostics
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Button

Public Module MyModule
    Public MyVar As String
End Module

Public Class Form4
    Dim excelApp As Excel.Application
    Dim form As New Form3



    Private Sub Form4_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        RadioButton1.Checked = False
        RadioButton2.Checked = False
        TextBox1.Focus()

    End Sub


    'Worksheet.Name = "New Worksheet"
    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged

        If RadioButton1.Checked = True Then

            TextBox1.Enabled = True
            PictureBox8.Enabled = True

            TextBox2.Enabled = False
            PictureBox1.Enabled = False
            TextBox3.Enabled = False
            PictureBox2.Enabled = False
            TextBox3.Text = ""


            'For hiding the visibilty of the selection range of existing workbook
            Me.PictureBox3.Visible = True
            Me.Label1.Visible = True
            Me.TextBox3.Visible = True
            Me.PictureBox2.Visible = True

            excelApp = Globals.ThisAddIn.Application
            Dim workbook As Excel.Workbook = excelApp.Workbooks.Add
            Dim Worksheet As Excel.Worksheet = CType(workbook.Sheets.Add(workbook.Sheets(1), Type.Missing, 1, Excel.XlSheetType.xlWorksheet), Excel.Worksheet)

            'Worksheet.Name = "New Worksheet"

        End If
        TextBox1.Focus()
    End Sub



    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        ' Create a new instance of the OpenFileDialog.
        Dim openFileDialog As New OpenFileDialog()

        ' Set some properties of the OpenFileDialog.
        openFileDialog.InitialDirectory = "c:\" ' Initial directory to be shown in the dialog
        openFileDialog.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*" ' Set the filter to txt files
        openFileDialog.FilterIndex = 2
        openFileDialog.RestoreDirectory = True

        ' Show the OpenFileDialog and get the result.
        If openFileDialog.ShowDialog() = DialogResult.OK Then
            ' Open the selected file.
            Try
                Process.Start(openFileDialog.FileName)
                Dim fileName As String = openFileDialog.FileName
                'TextBox2.Text = fileName
                TextBox2.Text = System.IO.Path.GetFileName(openFileDialog.FileName)
            Catch ex As Exception
                MessageBox.Show("An error occurred while trying to open the file: " & ex.Message)
            End Try
        End If
    End Sub

    Private Sub PictureBox8_Click(sender As Object, e As EventArgs) Handles PictureBox8.Click

        Me.Visible = False

        Dim selectedRange As Excel.Range = excelApp.InputBox("Select a range", Type:=8)
        selectedRange.Select()
        Me.Visible = True

        ' Put the selected range's address into the TextBox.
        TextBox1.Text = selectedRange.Address
        MyVar = TextBox1.Text
        'form.TextBox2.Text = TextBox1.Text


    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim form As New Form3()
        form.Show()
        Me.Hide()
    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click
        Me.Visible = False

        Dim selectedRange As Excel.Range = excelApp.InputBox("Select a range", Type:=8)
        selectedRange.Select()
        Me.Visible = True

        ' Put the selected range's address into the TextBox.
        TextBox3.Text = selectedRange.Address
        MyVar = TextBox3.Text
        ' form.TextBox2.Text = TextBox3.Text
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        If RadioButton2.Checked = True Then

            TextBox2.Enabled = True
            PictureBox1.Enabled = True

            TextBox1.Enabled = False
            PictureBox8.Enabled = False
            TextBox1.Text = ""

            'For the visibilty of the selection range
            Me.PictureBox3.Visible = True
            Me.Label1.Visible = True
            Me.TextBox3.Visible = True
            Me.PictureBox2.Visible = True



        End If

    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        TextBox3.Enabled = True
        PictureBox2.Enabled = True
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

    End Sub
End Class