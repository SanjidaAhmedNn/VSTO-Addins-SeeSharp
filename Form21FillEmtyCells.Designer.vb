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
        Me.CB_Backup_Sheet = New System.Windows.Forms.CheckBox()
        Me.txtSourceRange = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Selection = New System.Windows.Forms.PictureBox()
        Me.CB_Keepformatting = New System.Windows.Forms.CheckBox()
        Me.L_Fill_Options = New System.Windows.Forms.Label()
        Me.ComboBox_Options = New System.Windows.Forms.ComboBox()
        Me.L_Fill_Value = New System.Windows.Forms.Label()
        Me.txtFillValue = New System.Windows.Forms.TextBox()
        Me.CustomGroupBox3 = New VSTO_Addins.CustomGroupBox()
        Me.CustomGroupBox6 = New VSTO_Addins.CustomGroupBox()
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.PictureBox11 = New System.Windows.Forms.PictureBox()
        Me.RB_Certain_value = New System.Windows.Forms.RadioButton()
        Me.RB_Values_fromselected_range = New System.Windows.Forms.RadioButton()
        Me.RB_Linear_values = New System.Windows.Forms.RadioButton()
        Me.GB_sample = New VSTO_Addins.CustomGroupBox()
        CType(Me.Selection, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.CustomGroupBox3.SuspendLayout()
        Me.CustomGroupBox6.SuspendLayout()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox11, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btn_OK
        '
        Me.btn_OK.BackColor = System.Drawing.Color.White
        Me.btn_OK.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        Me.btn_OK.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btn_OK.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
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
        Me.btn_Cancel.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
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
        'CB_Backup_Sheet
        '
        Me.CB_Backup_Sheet.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CB_Backup_Sheet.Location = New System.Drawing.Point(15, 312)
        Me.CB_Backup_Sheet.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.CB_Backup_Sheet.Name = "CB_Backup_Sheet"
        Me.CB_Backup_Sheet.Size = New System.Drawing.Size(258, 29)
        Me.CB_Backup_Sheet.TabIndex = 164
        Me.CB_Backup_Sheet.Text = "Create a copy of the original worksheet"
        Me.CB_Backup_Sheet.UseVisualStyleBackColor = True
        '
        'txtSourceRange
        '
        Me.txtSourceRange.BackColor = System.Drawing.Color.White
        Me.txtSourceRange.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtSourceRange.Font = New System.Drawing.Font("Segoe UI", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSourceRange.ForeColor = System.Drawing.Color.FromArgb(CType(CType(70, Byte), Integer), CType(CType(70, Byte), Integer), CType(CType(70, Byte), Integer))
        Me.txtSourceRange.Location = New System.Drawing.Point(15, 42)
        Me.txtSourceRange.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.txtSourceRange.Name = "txtSourceRange"
        Me.txtSourceRange.Size = New System.Drawing.Size(248, 25)
        Me.txtSourceRange.TabIndex = 162
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
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
        'CB_Keepformatting
        '
        Me.CB_Keepformatting.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CB_Keepformatting.Location = New System.Drawing.Point(15, 193)
        Me.CB_Keepformatting.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.CB_Keepformatting.Name = "CB_Keepformatting"
        Me.CB_Keepformatting.Size = New System.Drawing.Size(136, 29)
        Me.CB_Keepformatting.TabIndex = 165
        Me.CB_Keepformatting.Text = "Keep formatting"
        Me.CB_Keepformatting.UseVisualStyleBackColor = True
        '
        'L_Fill_Options
        '
        Me.L_Fill_Options.AutoSize = True
        Me.L_Fill_Options.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
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
        Me.L_Fill_Value.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.L_Fill_Value.Location = New System.Drawing.Point(15, 276)
        Me.L_Fill_Value.Name = "L_Fill_Value"
        Me.L_Fill_Value.Size = New System.Drawing.Size(60, 17)
        Me.L_Fill_Value.TabIndex = 176
        Me.L_Fill_Value.Text = "Fill Value"
        '
        'txtFillValue
        '
        Me.txtFillValue.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtFillValue.Location = New System.Drawing.Point(101, 273)
        Me.txtFillValue.Name = "txtFillValue"
        Me.txtFillValue.Size = New System.Drawing.Size(163, 25)
        Me.txtFillValue.TabIndex = 178
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
        Me.CustomGroupBox6.Controls.Add(Me.PictureBox2)
        Me.CustomGroupBox6.Controls.Add(Me.PictureBox1)
        Me.CustomGroupBox6.Controls.Add(Me.PictureBox11)
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
        'PictureBox2
        '
        Me.PictureBox2.Image = CType(resources.GetObject("PictureBox2.Image"), System.Drawing.Image)
        Me.PictureBox2.Location = New System.Drawing.Point(220, 57)
        Me.PictureBox2.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(20, 20)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox2.TabIndex = 235
        Me.PictureBox2.TabStop = False
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(220, 33)
        Me.PictureBox1.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(20, 20)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 234
        Me.PictureBox1.TabStop = False
        '
        'PictureBox11
        '
        Me.PictureBox11.Image = CType(resources.GetObject("PictureBox11.Image"), System.Drawing.Image)
        Me.PictureBox11.Location = New System.Drawing.Point(220, 9)
        Me.PictureBox11.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.PictureBox11.Name = "PictureBox11"
        Me.PictureBox11.Size = New System.Drawing.Size(20, 20)
        Me.PictureBox11.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox11.TabIndex = 233
        Me.PictureBox11.TabStop = False
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
        'GB_sample
        '
        Me.GB_sample.BackColor = System.Drawing.Color.White
        Me.GB_sample.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.GB_sample.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GB_sample.Location = New System.Drawing.Point(286, 15)
        Me.GB_sample.Name = "GB_sample"
        Me.GB_sample.Size = New System.Drawing.Size(251, 315)
        Me.GB_sample.TabIndex = 401
        Me.GB_sample.TabStop = False
        Me.GB_sample.Text = "Sample Image"
        '
        'Form21FillEmtyCells
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.ClientSize = New System.Drawing.Size(560, 399)
        Me.Controls.Add(Me.GB_sample)
        Me.Controls.Add(Me.txtFillValue)
        Me.Controls.Add(Me.L_Fill_Value)
        Me.Controls.Add(Me.ComboBox_Options)
        Me.Controls.Add(Me.L_Fill_Options)
        Me.Controls.Add(Me.CB_Keepformatting)
        Me.Controls.Add(Me.CustomGroupBox3)
        Me.Controls.Add(Me.btn_OK)
        Me.Controls.Add(Me.btn_Cancel)
        Me.Controls.Add(Me.ComboBox1)
        Me.Controls.Add(Me.CB_Backup_Sheet)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Selection)
        Me.Controls.Add(Me.txtSourceRange)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Form21FillEmtyCells"
        Me.Text = "Fill Emty Cells"
        CType(Me.Selection, System.ComponentModel.ISupportInitialize).EndInit()
        Me.CustomGroupBox3.ResumeLayout(False)
        Me.CustomGroupBox6.ResumeLayout(False)
        Me.CustomGroupBox6.PerformLayout()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox11, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents RB_Values_fromselected_range As Windows.Forms.RadioButton
    Friend WithEvents RB_Linear_values As Windows.Forms.RadioButton
    Friend WithEvents CustomGroupBox6 As CustomGroupBox
    Friend WithEvents RB_Certain_value As Windows.Forms.RadioButton
    Friend WithEvents CustomGroupBox3 As CustomGroupBox
    Friend WithEvents btn_OK As Windows.Forms.Button
    Friend WithEvents btn_Cancel As Windows.Forms.Button
    Friend WithEvents ComboBox1 As Windows.Forms.ComboBox
    Friend WithEvents CB_Backup_Sheet As Windows.Forms.CheckBox
    Friend WithEvents txtSourceRange As Windows.Forms.TextBox
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents Selection As Windows.Forms.PictureBox
    Friend WithEvents CB_Keepformatting As Windows.Forms.CheckBox
    Friend WithEvents L_Fill_Options As Windows.Forms.Label
    Friend WithEvents ComboBox_Options As Windows.Forms.ComboBox
    Friend WithEvents L_Fill_Value As Windows.Forms.Label
    Friend WithEvents txtFillValue As Windows.Forms.TextBox
    Friend WithEvents PictureBox2 As Windows.Forms.PictureBox
    Friend WithEvents PictureBox1 As Windows.Forms.PictureBox
    Friend WithEvents PictureBox11 As Windows.Forms.PictureBox
    Friend WithEvents GB_sample As CustomGroupBox
End Class
