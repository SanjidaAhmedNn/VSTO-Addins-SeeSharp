Imports System.Drawing

Public Class Form3
    Private Sub btn_OK_Click(sender As Object, e As EventArgs) Handles btn_OK.Click

    End Sub

    Private Sub btn_OK_MouseEnter(sender As Object, e As EventArgs) Handles btn_OK.MouseEnter
        btn_OK.ForeColor = Color.White
        btn_OK.BackColor = Color.FromArgb(76, 111, 174)
    End Sub

    Private Sub btn_OK_MouseLeave(sender As Object, e As EventArgs) Handles btn_OK.MouseLeave
        btn_OK.ForeColor = Color.FromArgb(70, 70, 70)
        btn_OK.BackColor = Color.White
    End Sub

    Private Sub btn_cancel_MouseLeave(sender As Object, e As EventArgs) Handles btn_cancel.MouseLeave
        btn_cancel.ForeColor = Color.FromArgb(70, 70, 70)
        btn_cancel.BackColor = Color.White
    End Sub

    Private Sub btn_cancel_MouseEnter(sender As Object, e As EventArgs) Handles btn_cancel.MouseEnter
        btn_cancel.ForeColor = Color.White
        btn_cancel.BackColor = Color.FromArgb(76, 111, 174)
    End Sub

    Private Sub RadioButton5_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton5.CheckedChanged
        Dim form As New Form4()
        form.Show()

    End Sub
End Class