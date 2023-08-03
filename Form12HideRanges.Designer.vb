<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form12HideRanges
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form12HideRanges))
        Me.Selection = New System.Windows.Forms.PictureBox()
        Me.btn_OK = New System.Windows.Forms.Button()
        Me.btn_Cancel = New System.Windows.Forms.Button()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        Me.AutoSelection = New System.Windows.Forms.PictureBox()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.PictureBox7 = New System.Windows.Forms.PictureBox()
        Me.GB_Input_Range = New VSTO_Addins.CustomGroupBox()
        Me.CP_Input_Range = New VSTO_Addins.CustomPanel()
        Me.GB_Expected_Output = New VSTO_Addins.CustomGroupBox()
        Me.CP_Output_Range = New VSTO_Addins.CustomPanel()
        Me.CustomGroupBox3 = New VSTO_Addins.CustomGroupBox()
        Me.CustomGroupBox6 = New VSTO_Addins.CustomGroupBox()
        Me.RB_Single_Range = New System.Windows.Forms.RadioButton()
        Me.RB_Multiple_Range = New System.Windows.Forms.RadioButton()
        Me.CustomGroupBox4 = New VSTO_Addins.CustomGroupBox()
        Me.CustomGroupBox5 = New VSTO_Addins.CustomGroupBox()
        Me.RB_bidirection = New System.Windows.Forms.RadioButton()
        Me.RB_Row = New System.Windows.Forms.RadioButton()
        Me.RB_Column = New System.Windows.Forms.RadioButton()
        CType(Me.Selection, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AutoSelection, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox7, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GB_Input_Range.SuspendLayout()
        Me.GB_Expected_Output.SuspendLayout()
        Me.CustomGroupBox3.SuspendLayout()
        Me.CustomGroupBox6.SuspendLayout()
        Me.CustomGroupBox4.SuspendLayout()
        Me.CustomGroupBox5.SuspendLayout()
        Me.SuspendLayout()
        '
        'Selection
        '
        Me.Selection.BackColor = System.Drawing.Color.White
        Me.Selection.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Selection.Image = CType(resources.GetObject("Selection.Image"), System.Drawing.Image)
        Me.Selection.Location = New System.Drawing.Point(240, 42)
        Me.Selection.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Selection.Name = "Selection"
        Me.Selection.Size = New System.Drawing.Size(24, 25)
        Me.Selection.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.Selection.TabIndex = 126
        Me.Selection.TabStop = False
        '
        'btn_OK
        '
        Me.btn_OK.BackColor = System.Drawing.Color.White
        Me.btn_OK.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        Me.btn_OK.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btn_OK.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_OK.ForeColor = System.Drawing.Color.FromArgb(CType(CType(70, Byte), Integer), CType(CType(70, Byte), Integer), CType(CType(70, Byte), Integer))
        Me.btn_OK.Location = New System.Drawing.Point(407, 338)
        Me.btn_OK.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.btn_OK.Name = "btn_OK"
        Me.btn_OK.Size = New System.Drawing.Size(62, 26)
        Me.btn_OK.TabIndex = 124
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
        Me.btn_Cancel.Location = New System.Drawing.Point(484, 338)
        Me.btn_Cancel.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.btn_Cancel.Name = "btn_Cancel"
        Me.btn_Cancel.Size = New System.Drawing.Size(62, 26)
        Me.btn_Cancel.TabIndex = 123
        Me.btn_Cancel.Text = "Cancel"
        Me.btn_Cancel.UseVisualStyleBackColor = False
        '
        'ComboBox1
        '
        Me.ComboBox1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Items.AddRange(New Object() {"SOFTEKO", "About Us", "Help", "Feedback"})
        Me.ComboBox1.Location = New System.Drawing.Point(13, 340)
        Me.ComboBox1.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(84, 25)
        Me.ComboBox1.TabIndex = 122
        Me.ComboBox1.Text = "SOFTEKO"
        '
        'CheckBox1
        '
        Me.CheckBox1.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckBox1.Location = New System.Drawing.Point(13, 302)
        Me.CheckBox1.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(258, 29)
        Me.CheckBox1.TabIndex = 121
        Me.CheckBox1.Text = "Create a copy of the original worksheet"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'AutoSelection
        '
        Me.AutoSelection.BackColor = System.Drawing.Color.White
        Me.AutoSelection.Image = CType(resources.GetObject("AutoSelection.Image"), System.Drawing.Image)
        Me.AutoSelection.Location = New System.Drawing.Point(215, 43)
        Me.AutoSelection.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.AutoSelection.Name = "AutoSelection"
        Me.AutoSelection.Size = New System.Drawing.Size(24, 23)
        Me.AutoSelection.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.AutoSelection.TabIndex = 119
        Me.AutoSelection.TabStop = False
        '
        'TextBox1
        '
        Me.TextBox1.BackColor = System.Drawing.Color.White
        Me.TextBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TextBox1.Font = New System.Drawing.Font("Segoe UI", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(70, Byte), Integer), CType(CType(70, Byte), Integer), CType(CType(70, Byte), Integer))
        Me.TextBox1.Location = New System.Drawing.Point(15, 42)
        Me.TextBox1.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(248, 25)
        Me.TextBox1.TabIndex = 118
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Black
        Me.Label1.Location = New System.Drawing.Point(15, 15)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(175, 20)
        Me.Label1.TabIndex = 117
        Me.Label1.Text = "Source Range:"
        '
        'PictureBox7
        '
        Me.PictureBox7.Image = CType(resources.GetObject("PictureBox7.Image"), System.Drawing.Image)
        Me.PictureBox7.Location = New System.Drawing.Point(407, 150)
        Me.PictureBox7.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.PictureBox7.Name = "PictureBox7"
        Me.PictureBox7.Size = New System.Drawing.Size(33, 37)
        Me.PictureBox7.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox7.TabIndex = 160
        Me.PictureBox7.TabStop = False
        '
        'GB_Input_Range
        '
        Me.GB_Input_Range.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.GB_Input_Range.Controls.Add(Me.CP_Input_Range)
        Me.GB_Input_Range.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GB_Input_Range.Location = New System.Drawing.Point(295, 12)
        Me.GB_Input_Range.Name = "GB_Input_Range"
        Me.GB_Input_Range.Size = New System.Drawing.Size(252, 135)
        Me.GB_Input_Range.TabIndex = 158
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
        Me.CP_Input_Range.Size = New System.Drawing.Size(250, 105)
        Me.CP_Input_Range.TabIndex = 0
        '
        'GB_Expected_Output
        '
        Me.GB_Expected_Output.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.GB_Expected_Output.Controls.Add(Me.CP_Output_Range)
        Me.GB_Expected_Output.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GB_Expected_Output.Location = New System.Drawing.Point(294, 186)
        Me.GB_Expected_Output.Name = "GB_Expected_Output"
        Me.GB_Expected_Output.Size = New System.Drawing.Size(252, 135)
        Me.GB_Expected_Output.TabIndex = 159
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
        Me.CP_Output_Range.Size = New System.Drawing.Size(250, 105)
        Me.CP_Output_Range.TabIndex = 11
        '
        'CustomGroupBox3
        '
        Me.CustomGroupBox3.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.CustomGroupBox3.Controls.Add(Me.CustomGroupBox6)
        Me.CustomGroupBox3.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CustomGroupBox3.Location = New System.Drawing.Point(15, 81)
        Me.CustomGroupBox3.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.CustomGroupBox3.Name = "CustomGroupBox3"
        Me.CustomGroupBox3.Padding = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.CustomGroupBox3.Size = New System.Drawing.Size(249, 88)
        Me.CustomGroupBox3.TabIndex = 132
        Me.CustomGroupBox3.TabStop = False
        Me.CustomGroupBox3.Text = "Range Type"
        '
        'CustomGroupBox6
        '
        Me.CustomGroupBox6.BackColor = System.Drawing.Color.White
        Me.CustomGroupBox6.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.CustomGroupBox6.Controls.Add(Me.RB_Single_Range)
        Me.CustomGroupBox6.Controls.Add(Me.RB_Multiple_Range)
        Me.CustomGroupBox6.Location = New System.Drawing.Point(1, 24)
        Me.CustomGroupBox6.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.CustomGroupBox6.Name = "CustomGroupBox6"
        Me.CustomGroupBox6.Padding = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.CustomGroupBox6.Size = New System.Drawing.Size(248, 64)
        Me.CustomGroupBox6.TabIndex = 0
        Me.CustomGroupBox6.TabStop = False
        '
        'RB_Single_Range
        '
        Me.RB_Single_Range.AutoSize = True
        Me.RB_Single_Range.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RB_Single_Range.Location = New System.Drawing.Point(8, 9)
        Me.RB_Single_Range.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.RB_Single_Range.Name = "RB_Single_Range"
        Me.RB_Single_Range.Size = New System.Drawing.Size(102, 21)
        Me.RB_Single_Range.TabIndex = 93
        Me.RB_Single_Range.TabStop = True
        Me.RB_Single_Range.Text = "Single Range"
        Me.RB_Single_Range.UseVisualStyleBackColor = True
        '
        'RB_Multiple_Range
        '
        Me.RB_Multiple_Range.AutoSize = True
        Me.RB_Multiple_Range.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RB_Multiple_Range.Location = New System.Drawing.Point(8, 33)
        Me.RB_Multiple_Range.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.RB_Multiple_Range.Name = "RB_Multiple_Range"
        Me.RB_Multiple_Range.Size = New System.Drawing.Size(197, 21)
        Me.RB_Multiple_Range.TabIndex = 92
        Me.RB_Multiple_Range.TabStop = True
        Me.RB_Multiple_Range.Text = "Multiple Non-adjacent Range"
        Me.RB_Multiple_Range.UseVisualStyleBackColor = True
        '
        'CustomGroupBox4
        '
        Me.CustomGroupBox4.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.CustomGroupBox4.Controls.Add(Me.CustomGroupBox5)
        Me.CustomGroupBox4.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CustomGroupBox4.Location = New System.Drawing.Point(15, 186)
        Me.CustomGroupBox4.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.CustomGroupBox4.Name = "CustomGroupBox4"
        Me.CustomGroupBox4.Padding = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.CustomGroupBox4.Size = New System.Drawing.Size(249, 110)
        Me.CustomGroupBox4.TabIndex = 133
        Me.CustomGroupBox4.TabStop = False
        Me.CustomGroupBox4.Text = "Hide Option"
        '
        'CustomGroupBox5
        '
        Me.CustomGroupBox5.BackColor = System.Drawing.Color.White
        Me.CustomGroupBox5.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.CustomGroupBox5.Controls.Add(Me.RB_bidirection)
        Me.CustomGroupBox5.Controls.Add(Me.RB_Row)
        Me.CustomGroupBox5.Controls.Add(Me.RB_Column)
        Me.CustomGroupBox5.Location = New System.Drawing.Point(1, 22)
        Me.CustomGroupBox5.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.CustomGroupBox5.Name = "CustomGroupBox5"
        Me.CustomGroupBox5.Padding = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.CustomGroupBox5.Size = New System.Drawing.Size(248, 87)
        Me.CustomGroupBox5.TabIndex = 0
        Me.CustomGroupBox5.TabStop = False
        '
        'RB_bidirection
        '
        Me.RB_bidirection.AutoSize = True
        Me.RB_bidirection.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RB_bidirection.Location = New System.Drawing.Point(8, 58)
        Me.RB_bidirection.Name = "RB_bidirection"
        Me.RB_bidirection.Size = New System.Drawing.Size(117, 21)
        Me.RB_bidirection.TabIndex = 136
        Me.RB_bidirection.TabStop = True
        Me.RB_bidirection.Text = "Both directional"
        Me.RB_bidirection.UseVisualStyleBackColor = True
        '
        'RB_Row
        '
        Me.RB_Row.AutoSize = True
        Me.RB_Row.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RB_Row.Location = New System.Drawing.Point(8, 8)
        Me.RB_Row.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.RB_Row.Name = "RB_Row"
        Me.RB_Row.Size = New System.Drawing.Size(109, 21)
        Me.RB_Row.TabIndex = 117
        Me.RB_Row.TabStop = True
        Me.RB_Row.Text = "Row-wise only"
        Me.RB_Row.UseVisualStyleBackColor = True
        '
        'RB_Column
        '
        Me.RB_Column.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RB_Column.Location = New System.Drawing.Point(8, 32)
        Me.RB_Column.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.RB_Column.Name = "RB_Column"
        Me.RB_Column.Size = New System.Drawing.Size(151, 24)
        Me.RB_Column.TabIndex = 94
        Me.RB_Column.TabStop = True
        Me.RB_Column.Text = "Column-wise only"
        Me.RB_Column.UseVisualStyleBackColor = True
        '
        'Form12HideRanges
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.ClientSize = New System.Drawing.Size(565, 382)
        Me.Controls.Add(Me.GB_Input_Range)
        Me.Controls.Add(Me.GB_Expected_Output)
        Me.Controls.Add(Me.PictureBox7)
        Me.Controls.Add(Me.CustomGroupBox3)
        Me.Controls.Add(Me.CustomGroupBox4)
        Me.Controls.Add(Me.Selection)
        Me.Controls.Add(Me.btn_OK)
        Me.Controls.Add(Me.btn_Cancel)
        Me.Controls.Add(Me.ComboBox1)
        Me.Controls.Add(Me.CheckBox1)
        Me.Controls.Add(Me.AutoSelection)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.Label1)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Form12HideRanges"
        Me.Text = "Hide Only the Selected Range"
        CType(Me.Selection, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AutoSelection, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox7, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GB_Input_Range.ResumeLayout(False)
        Me.GB_Expected_Output.ResumeLayout(False)
        Me.CustomGroupBox3.ResumeLayout(False)
        Me.CustomGroupBox6.ResumeLayout(False)
        Me.CustomGroupBox6.PerformLayout()
        Me.CustomGroupBox4.ResumeLayout(False)
        Me.CustomGroupBox5.ResumeLayout(False)
        Me.CustomGroupBox5.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents RB_Single_Range As Windows.Forms.RadioButton
    Friend WithEvents RB_Multiple_Range As Windows.Forms.RadioButton
    Friend WithEvents CustomGroupBox6 As CustomGroupBox
    Friend WithEvents CustomGroupBox3 As CustomGroupBox
    Friend WithEvents CustomGroupBox5 As CustomGroupBox
    Friend WithEvents RB_Row As Windows.Forms.RadioButton
    Friend WithEvents RB_Column As Windows.Forms.RadioButton
    Friend WithEvents CustomGroupBox4 As CustomGroupBox
    Friend WithEvents Selection As Windows.Forms.PictureBox
    Friend WithEvents btn_OK As Windows.Forms.Button
    Friend WithEvents btn_Cancel As Windows.Forms.Button
    Friend WithEvents ComboBox1 As Windows.Forms.ComboBox
    Friend WithEvents CheckBox1 As Windows.Forms.CheckBox
    Friend WithEvents AutoSelection As Windows.Forms.PictureBox
    Friend WithEvents TextBox1 As Windows.Forms.TextBox
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents RB_bidirection As Windows.Forms.RadioButton
    Friend WithEvents GB_Input_Range As CustomGroupBox
    Friend WithEvents CP_Input_Range As CustomPanel
    Friend WithEvents GB_Expected_Output As CustomGroupBox
    Friend WithEvents CP_Output_Range As CustomPanel
    Friend WithEvents PictureBox7 As Windows.Forms.PictureBox
End Class
