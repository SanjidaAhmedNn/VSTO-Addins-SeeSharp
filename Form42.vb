Imports System.Windows.Forms

Public Class Form42
    Dim form1 As Form29_Simple_Drop_down_List = Nothing
    Dim form2 As Form30_Create_Dynamic_Drop_down_List = Nothing
    Private Sub RadioButton5_CheckedChanged(sender As Object, e As EventArgs) Handles RB_No.CheckedChanged
        If RB_No.Checked = True Then
            CGB.Enabled = False

        End If
    End Sub

    Private Sub RB_Yes_CheckedChanged(sender As Object, e As EventArgs) Handles RB_Yes.CheckedChanged
        If RB_Yes.Checked = True Then
            CGB.Enabled = True

        End If
    End Sub

    Private Sub Form42_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Btn_Cancel_Click(sender As Object, e As EventArgs) Handles Btn_Cancel.Click
        Me.Close()
    End Sub

    Private Sub Btn_OK_Click(sender As Object, e As EventArgs) Handles Btn_OK.Click
        If RB_Simple.Checked = True Then
            form1 = New Form29_Simple_Drop_down_List()
            Me.Hide()
            form1.Show()
        ElseIf RB_Dynamic.Checked = True Then
            form2 = New Form30_Create_Dynamic_Drop_down_List()
            Me.Hide()
            form2.Show()

        End If
        Me.Close()

    End Sub
End Class