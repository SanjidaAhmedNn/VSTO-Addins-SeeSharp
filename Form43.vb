Imports System.ComponentModel

Public Class Form43
    Dim form As Form33_ColorBasedDropDownList = Nothing

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        form = New Form33_ColorBasedDropDownList
        form.Show()
        Me.Dispose()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        form_flag = False
        Me.Dispose()
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            sessionflag1 = False
        Else
            sessionflag1 = True
        End If
    End Sub

    Private Sub Form43_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Form43_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        form_flag = False
    End Sub

    Private Sub Form43_Disposed(sender As Object, e As EventArgs) Handles Me.Disposed
        form_flag = False
    End Sub
End Class