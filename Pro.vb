Imports System.Threading
Imports System.Windows.Forms
Imports System.Windows.Forms.VisualStyles.VisualStyleElement

Public Class FormProgressBar
    Public Shared proBar As Windows.Forms.ProgressBar
    Private Sub Pro_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'For i As Integer = 0 To 100
        '    ProgressBar1.Value = i



        '    ' Include your task logic here
        '    ' Use Application.DoEvents() if needed to refresh the UI
        'Next

        'Label1.Text = Ribbon1.captiontxt


        proBar = ProgressBar1
        proBar.Minimum = 0
        proBar.Maximum = 100

        Me.Label1.Text = "Progress: " & proBar.Value.ToString() & "%"

    End Sub


    Private Sub HandleProgressBarValueChanged()
        'Me.Label1.Text = "Progress: " & proBar.Value.ToString() & "%"
    End Sub
End Class