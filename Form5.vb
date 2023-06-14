Imports System.Drawing
Imports System.Windows.Forms

Public Class Form5
    Private Sub CustomButton1_Click(sender As Object, e As EventArgs) Handles CustomButton1.Click
        'CustomButton1.Font = New Font("Segoe UI Semibold", 10.75, FontStyle.Regular)
    End Sub

    Private Sub CustomButton1_MouseHover(sender As Object, e As EventArgs) Handles CustomButton1.MouseHover
        ' CustomButton1.Font = New Font("Segoe UI", 9.75, FontStyle.Bold)
        CustomButton1.Font = New Font("Segoe UI", 9.75, FontStyle.Bold)
    End Sub

    Private Sub CustomButton1_MouseLeave(sender As Object, e As EventArgs) Handles CustomButton1.MouseLeave
        CustomButton1.Font = New Font("Segoe UI semibold", 9.75, FontStyle.Bold)
    End Sub

    Private Sub CustomButton1_LostFocus(sender As Object, e As EventArgs) Handles CustomButton1.LostFocus
        CustomButton1.Font = New Font("Segoe UI semibold", 9.75, FontStyle.Bold)
    End Sub

End Class