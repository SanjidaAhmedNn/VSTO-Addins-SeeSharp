<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form21FillEmtyCells
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form21FillEmtyCells))
        Me.btn_OK = New System.Windows.Forms.Button()
        Me.btn_Cancel = New System.Windows.Forms.Button()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.Backup_sheet = New System.Windows.Forms.CheckBox()
        Me.AutoSelection = New System.Windows.Forms.PictureBox()
        Me.Textbox1 = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Selection = New System.Windows.Forms.PictureBox()
        Me.Checkbox_Keepformatting = New System.Windows.Forms.CheckBox()
        Me.PictureBox7 = New System.Windows.Forms.PictureBox()
        Me.L_Fill_Options = New System.Windows.Forms.Label()
        Me.ComboBox_Options = New System.Windows.Forms.ComboBox()
        Me.L_Fill_Value = New System.Windows.Forms.Label()
        Me.CustomGroupBox3 = New VSTO_Addins.CustomGroupBox()
        Me.CustomGroupBox6 = New VSTO_Addins.CustomGroupBox()
        Me.RB_Certain_value = New System.Windows.Forms.RadioButton()
        Me.RB_Values_fromselected_range = New System.Windows.Forms.RadioButton()
        Me.RB_Linear_values = New System.Windows.Forms.RadioButton()
        Me.GB_Expected_Output = New VSTO_Addins.CustomGroupBox()
        Me.CP_Output_Range = New VSTO_Addins.CustomPanel()
        Me.GB_Input_Range = New VSTO_Addins.CustomGroupBox()
        Me.CP_Input_Range = New VSTO_Addins.CustomPanel()
        Me.TextBox_Value = New System.Windows.Forms.TextBox()
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.PictureBox3 = New System.Windows.Forms.PictureBox()
        CType(Me.AutoSelection, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Selection, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox7, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.CustomGroupBox3.SuspendLayout()
        Me.CustomGroupBox6.SuspendLayout()
        Me.GB_Expected_Output.SuspendLayout()
        Me.GB_Input_Range.SuspendLayout()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btn_OK
        '
        Me.btn_OK.BackColor = System.Drawing.Color.White
        Me.btn_OK.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        Me.btn_OK.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btn_OK.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_OK.ForeColor = System.Drawing.Color.FromArgb(CType(CType(70, Byte), Integer), CType(CType(70, Byte), Integer), CType(CType(70, Byte), Integer))
        Me.btn_OK.Location = New System.Drawing.Point(397, 351)
        Me.btn_OK.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.btn_OK.Name = "btn_OK"
        Me.btn_OK.Size = New System.Drawing.Size(62, 26)
        Me.btn_OK.TabIndex = 167
        Me.btn_OK.Text = "OK"
        Me.btn_OK.UseVisualStyleBackColor = False
        '
        'btn_Cancel
        '
        Me.btn_Cancel.BackColor = System.Drawing.Color.White
        Me.btn_Cancel.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        Me.btn_Cancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btn_Cancel.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btn_Cancel.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_Cancel.ForeColor = System.Drawing.Color.FromArgb(CType(CType(70, Byte), Integer), CType(CType(70, Byte), Integer), CType(CType(70, Byte), Integer))
        Me.btn_Cancel.Location = New System.Drawing.Point(475, 351)
        Me.btn_Cancel.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.btn_Cancel.Name = "btn_Cancel"
        Me.btn_Cancel.Size = New System.Drawing.Size(62, 26)
        Me.btn_Cancel.TabIndex = 166
        Me.btn_Cancel.Text = "Cancel"
        Me.btn_Cancel.UseVisualStyleBackColor = False
        '
        'ComboBox1
        '
        Me.ComboBox1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Items.AddRange(New Object() {"SOFTEKO", "About Us", "Help", "Feedback"})
        Me.ComboBox1.Location = New System.Drawing.Point(15, 353)
        Me.ComboBox1.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(100, 25)
        Me.ComboBox1.TabIndex = 165
        Me.ComboBox1.Text = "SOFTEKO"
        '
        'Backup_sheet
        '
        Me.Backup_sheet.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Backup_sheet.Location = New System.Drawing.Point(15, 312)
        Me.Backup_sheet.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Backup_sheet.Name = "Backup_sheet"
        Me.Backup_sheet.Size = New System.Drawing.Size(258, 29)
        Me.Backup_sheet.TabIndex = 164
        Me.Backup_sheet.Text = "Create a copy of the original worksheet"
        Me.Backup_sheet.UseVisualStyleBackColor = True
        '
        'AutoSelection
        '
        Me.AutoSelection.BackColor = System.Drawing.Color.White
        Me.AutoSelection.Image = CType(resources.GetObject("AutoSelection.Image"), System.Drawing.Image)
        Me.AutoSelection.Location = New System.Drawing.Point(214, 43)
        Me.AutoSelection.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.AutoSelection.Name = "AutoSelection"
        Me.AutoSelection.Size = New System.Drawing.Size(24, 23)
        Me.AutoSelection.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.AutoSelection.TabIndex = 163
        Me.AutoSelection.TabStop = False
        '
        'Textbox1
        '
        Me.Textbox1.BackColor = System.Drawing.Color.White
        Me.Textbox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Textbox1.Font = New System.Drawing.Font("Segoe UI", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Textbox1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(70, Byte), Integer), CType(CType(70, Byte), Integer), CType(CType(70, Byte), Integer))
        Me.Textbox1.Location = New System.Drawing.Point(15, 42)
        Me.Textbox1.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Textbox1.Name = "Textbox1"
        Me.Textbox1.Size = New System.Drawing.Size(248, 25)
        Me.Textbox1.TabIndex = 162
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Black
        Me.Label1.Location = New System.Drawing.Point(15, 15)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(175, 20)
        Me.Label1.TabIndex = 161
        Me.Label1.Text = "Source Range:"
        '
        'Selection
        '
        Me.Selection.BackColor = System.Drawing.Color.White
        Me.Selection.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Selection.Image = CType(resources.GetObject("Selection.Image"), System.Drawing.Image)
        Me.Selection.Location = New System.Drawing.Point(239, 42)
        Me.Selection.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Selection.Name = "Selection"
        Me.Selection.Size = New System.Drawing.Size(24, 25)
        Me.Selection.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.Selection.TabIndex = 168
        Me.Selection.TabStop = False
        '
        'Checkbox_Keepformatting
        '
        Me.Checkbox_Keepformatting.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Checkbox_Keepformatting.Location = New System.Drawing.Point(15, 193)
        Me.Checkbox_Keepformatting.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Checkbox_Keepformatting.Name = "Checkbox_Keepformatting"
        Me.Checkbox_Keepformatting.Size = New System.Drawing.Size(136, 29)
        Me.Checkbox_Keepformatting.TabIndex = 165
        Me.Checkbox_Keepformatting.Text = "Keep formatting"
        Me.Checkbox_Keepformatting.UseVisualStyleBackColor = True
        '
        'PictureBox7
        '
        Me.PictureBox7.Image = CType(resources.GetObject("PictureBox7.Image"), System.Drawing.Image)
        Me.PictureBox7.Location = New System.Drawing.Point(397, 159)
        Me.PictureBox7.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.PictureBox7.Name = "PictureBox7"
        Me.PictureBox7.Size = New System.Drawing.Size(40, 40)
        Me.PictureBox7.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox7.TabIndex = 173
        Me.PictureBox7.TabStop = False
        '
        'L_Fill_Options
        '
        Me.L_Fill_Options.AutoSize = True
        Me.L_Fill_Options.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.L_Fill_Options.Location = New System.Drawing.Point(15, 234)
        Me.L_Fill_Options.Name = "L_Fill_Options"
        Me.L_Fill_Options.Size = New System.Drawing.Size(76, 17)
        Me.L_Fill_Options.TabIndex = 174
        Me.L_Fill_Options.Text = "Fill Options"
        '
        'ComboBox_Options
        '
        Me.ComboBox_Options.BackColor = System.Drawing.Color.White
        Me.ComboBox_Options.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_Options.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBox_Options.ForeColor = System.Drawing.Color.Black
        Me.ComboBox_Options.FormattingEnabled = True
        Me.ComboBox_Options.Location = New System.Drawing.Point(101, 230)
        Me.ComboBox_Options.Name = "ComboBox_Options"
        Me.ComboBox_Options.Size = New System.Drawing.Size(163, 25)
        Me.ComboBox_Options.TabIndex = 175
        '
        'L_Fill_Value
        '
        Me.L_Fill_Value.AutoSize = True
        Me.L_Fill_Value.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.L_Fill_Value.Location = New System.Drawing.Point(15, 276)
        Me.L_Fill_Value.Name = "L_Fill_Value"
        Me.L_Fill_Value.Size = New System.Drawing.Size(60, 17)
        Me.L_Fill_Value.TabIndex = 176
        Me.L_Fill_Value.Text = "Fill Value"
        '
        'CustomGroupBox3
        '
        Me.CustomGroupBox3.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.CustomGroupBox3.Controls.Add(Me.CustomGroupBox6)
        Me.CustomGroupBox3.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CustomGroupBox3.Location = New System.Drawing.Point(15, 80)
        Me.CustomGroupBox3.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.CustomGroupBox3.Name = "CustomGroupBox3"
        Me.CustomGroupBox3.Padding = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.CustomGroupBox3.Size = New System.Drawing.Size(249, 107)
        Me.CustomGroupBox3.TabIndex = 169
        Me.CustomGroupBox3.TabStop = False
        Me.CustomGroupBox3.Text = "Fill Cells"
        '
        'CustomGroupBox6
        '
        Me.CustomGroupBox6.BackColor = System.Drawing.Color.White
        Me.CustomGroupBox6.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.CustomGroupBox6.Controls.Add(Me.PictureBox3)
        Me.CustomGroupBox6.Controls.Add(Me.PictureBox1)
        Me.CustomGroupBox6.Controls.Add(Me.PictureBox2)
        Me.CustomGroupBox6.Controls.Add(Me.RB_Certain_value)
        Me.CustomGroupBox6.Controls.Add(Me.RB_Values_fromselected_range)
        Me.CustomGroupBox6.Controls.Add(Me.RB_Linear_values)
        Me.CustomGroupBox6.Location = New System.Drawing.Point(1, 24)
        Me.CustomGroupBox6.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.CustomGroupBox6.Name = "CustomGroupBox6"
        Me.CustomGroupBox6.Padding = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.CustomGroupBox6.Size = New System.Drawing.Size(248, 82)
        Me.CustomGroupBox6.TabIndex = 0
        Me.CustomGroupBox6.TabStop = False
        '
        'RB_Certain_value
        '
        Me.RB_Certain_value.AutoSize = True
        Me.RB_Certain_value.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RB_Certain_value.Location = New System.Drawing.Point(8, 55)
        Me.RB_Certain_value.Name = "RB_Certain_value"
        Me.RB_Certain_value.Size = New System.Drawing.Size(129, 21)
        Me.RB_Certain_value.TabIndex = 94
        Me.RB_Certain_value.Text = "With certain value"
        Me.RB_Certain_value.UseVisualStyleBackColor = True
        '
        'RB_Values_fromselected_range
        '
        Me.RB_Values_fromselected_range.AutoSize = True
        Me.RB_Values_fromselected_range.Checked = True
        Me.RB_Values_fromselected_range.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RB_Values_fromselected_range.Location = New System.Drawing.Point(8, 8)
        Me.RB_Values_fromselected_range.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.RB_Values_fromselected_range.Name = "RB_Values_fromselected_range"
        Me.RB_Values_fromselected_range.Size = New System.Drawing.Size(214, 21)
        Me.RB_Values_fromselected_range.TabIndex = 93
        Me.RB_Values_fromselected_range.TabStop = True
        Me.RB_Values_fromselected_range.Text = "With values from selected range"
        Me.RB_Values_fromselected_range.UseVisualStyleBackColor = True
        '
        'RB_Linear_values
        '
        Me.RB_Linear_values.AutoSize = True
        Me.RB_Linear_values.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RB_Linear_values.Location = New System.Drawing.Point(8, 31)
        Me.RB_Linear_values.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.RB_Linear_values.Name = "RB_Linear_values"
        Me.RB_Linear_values.Size = New System.Drawing.Size(128, 21)
        Me.RB_Linear_values.TabIndex = 92
        Me.RB_Linear_values.Text = "With linear values"
        Me.RB_Linear_values.UseVisualStyleBackColor = True
        '
        'GB_Expected_Output
        '
        Me.GB_Expected_Output.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.GB_Expected_Output.Controls.Add(Me.CP_Output_Range)
        Me.GB_Expected_Output.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GB_Expected_Output.Location = New System.Drawing.Point(291, 195)
        Me.GB_Expected_Output.Name = "GB_Expected_Output"
        Me.GB_Expected_Output.Size = New System.Drawing.Size(247, 140)
        Me.GB_Expected_Output.TabIndex = 172
        Me.GB_Expected_Output.TabStop = False
        Me.GB_Expected_Output.Text = "Expected Output"
        '
        'CP_Output_Range
        '
        Me.CP_Output_Range.BackColor = System.Drawing.Color.White
        Me.CP_Output_Range.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.CP_Output_Range.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.CP_Output_Range.BorderWidth = 1
        Me.CP_Output_Range.Location = New System.Drawing.Point(1, 30)
        Me.CP_Output_Range.Name = "CP_Output_Range"
        Me.CP_Output_Range.Size = New System.Drawing.Size(245, 110)
        Me.CP_Output_Range.TabIndex = 11
        '
        'GB_Input_Range
        '
        Me.GB_Input_Range.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.GB_Input_Range.Controls.Add(Me.CP_Input_Range)
        Me.GB_Input_Range.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GB_Input_Range.Location = New System.Drawing.Point(293, 14)
        Me.GB_Input_Range.Name = "GB_Input_Range"
        Me.GB_Input_Range.Size = New System.Drawing.Size(247, 140)
        Me.GB_Input_Range.TabIndex = 171
        Me.GB_Input_Range.TabStop = False
        Me.GB_Input_Range.Text = "Input Range"
        '
        'CP_Input_Range
        '
        Me.CP_Input_Range.BackColor = System.Drawing.Color.White
        Me.CP_Input_Range.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.CP_Input_Range.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.CP_Input_Range.BorderWidth = 1
        Me.CP_Input_Range.Location = New System.Drawing.Point(1, 30)
        Me.CP_Input_Range.Name = "CP_Input_Range"
        Me.CP_Input_Range.Size = New System.Drawing.Size(245, 110)
        Me.CP_Input_Range.TabIndex = 0
        '
        'TextBox_Value
        '
        Me.TextBox_Value.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_Value.Location = New System.Drawing.Point(101, 273)
        Me.TextBox_Value.Name = "TextBox_Value"
        Me.TextBox_Value.Size = New System.Drawing.Size(163, 25)
        Me.TextBox_Value.TabIndex = 178
        '
        'PictureBox2
        '
        Me.PictureBox2.Image = CType(resources.GetObject("PictureBox2.Image"), System.Drawing.Image)
        Me.PictureBox2.Location = New System.Drawing.Point(220, 9)
        Me.PictureBox2.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(20, 20)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox2.TabIndex = 230
        Me.PictureBox2.TabStop = False
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(220, 32)
        Me.PictureBox1.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(20, 20)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 231
        Me.PictureBox1.TabStop = False
        '
        'PictureBox3
        '
        Me.PictureBox3.Image = CType(resources.GetObject("PictureBox3.Image"), System.Drawing.Image)
        Me.PictureBox3.Location = New System.Drawing.Point(220, 56)
        Me.PictureBox3.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.PictureBox3.Name = "PictureBox3"
        Me.PictureBox3.Size = New System.Drawing.Size(20, 20)
        Me.PictureBox3.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox3.TabIndex = 232
        Me.PictureBox3.TabStop = False
        '
        'Form21FillEmtyCells
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(560, 399)
        Me.Controls.Add(Me.TextBox_Value)
        Me.Controls.Add(Me.L_Fill_Value)
        Me.Controls.Add(Me.ComboBox_Options)
        Me.Controls.Add(Me.L_Fill_Options)
        Me.Controls.Add(Me.PictureBox7)
        Me.Controls.Add(Me.Checkbox_Keepformatting)
        Me.Controls.Add(Me.CustomGroupBox3)
        Me.Controls.Add(Me.GB_Expected_Output)
        Me.Controls.Add(Me.GB_Input_Range)
        Me.Controls.Add(Me.btn_OK)
        Me.Controls.Add(Me.btn_Cancel)
        Me.Controls.Add(Me.ComboBox1)
        Me.Controls.Add(Me.Backup_sheet)
        Me.Controls.Add(Me.AutoSelection)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Selection)
        Me.Controls.Add(Me.Textbox1)
        Me.Name = "Form21FillEmtyCells"
        Me.Text = "Fill Emty Cells"
        CType(Me.AutoSelection, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Selection, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox7, System.ComponentModel.ISupportInitialize).EndInit()
        Me.CustomGroupBox3.ResumeLayout(False)
        Me.CustomGroupBox6.ResumeLayout(False)
        Me.CustomGroupBox6.PerformLayout()
        Me.GB_Expected_Output.ResumeLayout(False)
        Me.GB_Input_Range.ResumeLayout(False)
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents RB_Values_fromselected_range As Windows.Forms.RadioButton
    Friend WithEvents RB_Linear_values As Windows.Forms.RadioButton
    Friend WithEvents CustomGroupBox6 As CustomGroupBox
    Friend WithEvents RB_Certain_value As Windows.Forms.RadioButton
    Friend WithEvents CustomGroupBox3 As CustomGroupBox
    Friend WithEvents CP_Output_Range As CustomPanel
    Friend WithEvents GB_Expected_Output As CustomGroupBox
    Friend WithEvents CP_Input_Range As CustomPanel
    Friend WithEvents GB_Input_Range As CustomGroupBox
    Friend WithEvents btn_OK As Windows.Forms.Button
    Friend WithEvents btn_Cancel As Windows.Forms.Button
    Friend WithEvents ComboBox1 As Windows.Forms.ComboBox
    Friend WithEvents Backup_sheet As Windows.Forms.CheckBox
    Friend WithEvents AutoSelection As Windows.Forms.PictureBox
    Friend WithEvents Textbox1 As Windows.Forms.TextBox
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents Selection As Windows.Forms.PictureBox
    Friend WithEvents Checkbox_Keepformatting As Windows.Forms.CheckBox
    Friend WithEvents PictureBox7 As Windows.Forms.PictureBox
    Friend WithEvents L_Fill_Options As Windows.Forms.Label
    Friend WithEvents ComboBox_Options As Windows.Forms.ComboBox
    Friend WithEvents L_Fill_Value As Windows.Forms.Label
    Friend WithEvents TextBox_Value As Windows.Forms.TextBox
    Friend WithEvents PictureBox3 As Windows.Forms.PictureBox
    Friend WithEvents PictureBox1 As Windows.Forms.PictureBox
    Friend WithEvents PictureBox2 As Windows.Forms.PictureBox
End Class
