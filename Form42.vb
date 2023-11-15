Imports System.ComponentModel
Imports System.Windows.Forms

Public Class Form42
    Dim form1 As Form29_Simple_Drop_down_List = Nothing
    Dim form2 As Form30_Create_Dynamic_Drop_down_List = Nothing
    Private Sub RadioButton5_CheckedChanged(sender As Object, e As EventArgs) Handles RB_No.CheckedChanged
        If RB_No.Checked = True Then
            CGB.Enabled = False
            RB_Simple.Enabled = False
            RB_Dynamic.Enabled = False
        End If
    End Sub

    Private Sub RB_Yes_CheckedChanged(sender As Object, e As EventArgs) Handles RB_Yes.CheckedChanged
        If RB_Yes.Checked = True Then
            CGB.Enabled = True
            RB_Simple.Enabled = True
            RB_Dynamic.Enabled = True

        End If
    End Sub

    Private Sub Form42_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Btn_Cancel_Click(sender As Object, e As EventArgs) Handles Btn_Cancel.Click
        Me.Close()
    End Sub

    Private Sub Btn_OK_Click(sender As Object, e As EventArgs) Handles Btn_OK.Click
        If RB_Simple.Checked = True And RB_Simple.Enabled = True Then
            form1 = New Form29_Simple_Drop_down_List()
            Me.Hide()
            form1.Show()
        ElseIf RB_Dynamic.Checked = True And RB_Dynamic.Enabled = True Then
            form2 = New Form30_Create_Dynamic_Drop_down_List()
            Me.Hide()
            form2.Show()
        Else
            Me.Dispose()
        End If
        Me.Close()

    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            sessionflag2 = False
        Else
            sessionflag2 = True
        End If
    End Sub

    Private Sub Form42_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        form_flag = False
    End Sub

    Private Sub Form42_Disposed(sender As Object, e As EventArgs) Handles Me.Disposed
        form_flag = False
    End Sub

    Private Sub RB_Simple_CheckedChanged(sender As Object, e As EventArgs) Handles RB_Simple.CheckedChanged

    End Sub
End Class