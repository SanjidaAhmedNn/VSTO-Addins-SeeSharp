Public Class Form26_split_text_bycharacters
    Private Sub CB_separators_finaloutput_CheckedChanged(sender As Object, e As EventArgs) Handles CB_separators_finaloutput.CheckedChanged
        If CB_separators_finaloutput.Checked = True Then
            RB_starting_point.Enabled = True
            RB_ending_point.Enabled = True

        ElseIf CB_separators_finaloutput.Checked = False Then
            RB_starting_point.Enabled = False
            RB_ending_point.Enabled = False
        End If

    End Sub
End Class