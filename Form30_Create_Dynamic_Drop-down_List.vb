Imports System.ComponentModel

Public Class Form30_Create_Dynamic_Drop_down_List
    Private Sub CB_ascending_CheckedChanged(sender As Object, e As EventArgs) Handles CB_ascending.CheckedChanged
        If CB_ascending.Checked = True Then
            CB_descending.Checked = False
        End If
    End Sub

    Private Sub CB_descending_CheckedChanged(sender As Object, e As EventArgs) Handles CB_descending.CheckedChanged
        If CB_descending.Checked = True Then
            CB_ascending.Checked = False
        End If
    End Sub

    Private Sub RB_columns_CheckedChanged(sender As Object, e As EventArgs) Handles RB_columns.CheckedChanged
        If RB_columns.Checked = True Then

            CB_header.Enabled = True
            CB_ascending.Enabled = True
            CB_descending.Enabled = True
            CB_text.Enabled = True
            GB_list_option.Enabled = False

        End If
    End Sub

    Private Sub RB_rows_CheckedChanged(sender As Object, e As EventArgs) Handles RB_rows.CheckedChanged
        If RB_rows.Checked = True Then
            GB_list_option.Enabled = True
            CB_header.Enabled = False
            CB_ascending.Enabled = False
            CB_descending.Enabled = False
            CB_text.Enabled = False

        End If
    End Sub

    Private Sub Form30_Create_Dynamic_Drop_down_List_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Btn_OK_Click(sender As Object, e As EventArgs) Handles Btn_OK.Click

    End Sub

    Private Sub Form30_Create_Dynamic_Drop_down_List_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        form_flag = False
    End Sub

    Private Sub Form30_Create_Dynamic_Drop_down_List_Disposed(sender As Object, e As EventArgs) Handles Me.Disposed
        form_flag = False
    End Sub

    Private Sub Form30_Create_Dynamic_Drop_down_List_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        Me.Focus()
        Me.BringToFront()
        Me.Activate()
        Me.BeginInvoke(New System.Action(Sub()
                                             TB_src_range.Text = src_range.Address
                                             SetWindowPos(Me.Handle, New IntPtr(HWND_TOPMOST), 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOSIZE)
                                         End Sub))
    End Sub

    Private Sub TB_src_range_TextChanged(sender As Object, e As EventArgs) Handles TB_src_range.TextChanged

    End Sub
End Class