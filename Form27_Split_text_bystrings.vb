Public Class Form27_Split_text_bystrings
    Private Sub CB_separators_finaloutput_CheckedChanged(sender As Object, e As EventArgs) Handles CB_separators_finaloutput.CheckedChanged
        If CB_separators_finaloutput.Checked = True Then
            RB_starting_point.Enabled = True
            RB_ending_point.Enabled = True

        ElseIf CB_separators_finaloutput.Checked = False Then
            RB_starting_point.Enabled = False
            RB_ending_point.Enabled = False
        End If
    End Sub

    Private Sub CB_consecutive_separators_CheckedChanged(sender As Object, e As EventArgs) Handles CB_consecutive_separators.CheckedChanged

    End Sub
End Class