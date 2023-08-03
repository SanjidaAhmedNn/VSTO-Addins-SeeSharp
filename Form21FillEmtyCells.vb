Public Class Form21FillEmtyCells
    Private Sub RB_Linear_values_CheckedChanged(sender As Object, e As EventArgs) Handles RB_Linear_values.CheckedChanged
        If RB_Linear_values.Checked = True Then
            ComboBox_Options.Items.Clear()
            ComboBox_Options.Items.Add("Top to Buttom")
            ComboBox_Options.Items.Add("Left to Right")
            ComboBox_Options.SelectedIndex = 0
            TextBox_Value.Enabled = False
            L_Fill_Value.Enabled = False
            ComboBox_Options.Enabled = True
            L_Fill_Options.Enabled = True

        End If

    End Sub

    Private Sub RB_Values_fromselected_range_CheckedChanged(sender As Object, e As EventArgs) Handles RB_Values_fromselected_range.CheckedChanged
        If RB_Values_fromselected_range.Checked = True Then
            ComboBox_Options.Items.Clear()
            ComboBox_Options.Items.Add("Downwards")
            ComboBox_Options.Items.Add("Upwards")
            ComboBox_Options.Items.Add("Towards the Right")
            ComboBox_Options.Items.Add("Towards the Left")
            ComboBox_Options.SelectedIndex = 0
            TextBox_Value.Enabled = False
            L_Fill_Value.Enabled = False
            ComboBox_Options.Enabled = True
            L_Fill_Options.Enabled = True

        End If
    End Sub

    Private Sub RB_Certain_value_CheckedChanged(sender As Object, e As EventArgs) Handles RB_Certain_value.CheckedChanged
        If RB_Certain_value.Checked = True Then
            ComboBox_Options.Items.Clear()
            ComboBox_Options.SelectedItem = ""
            TextBox_Value.Enabled = True
            L_Fill_Value.Enabled = True
            ComboBox_Options.Enabled = False
            L_Fill_Options.Enabled = False
        End If
    End Sub


End Class