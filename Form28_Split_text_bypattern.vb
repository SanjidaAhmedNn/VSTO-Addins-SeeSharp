Public Class Form28_Split_text_bypattern
    Private Sub CB_separators_final_output_CheckedChanged(sender As Object, e As EventArgs) Handles CB_separators_final_output.CheckedChanged
        If CB_separators_final_output.Checked = True Then
            RB_starting_point.Enabled = True
            RB_ending_point.Enabled = True

        ElseIf CB_separators_final_output.Checked = False Then
            RB_starting_point.Enabled = False
            RB_ending_point.Enabled = False
        End If
    End Sub
End Class