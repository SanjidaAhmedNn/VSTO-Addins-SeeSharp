<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form16PasteintoVisibleRange
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form16PasteintoVisibleRange))
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.CheckBox2 = New System.Windows.Forms.CheckBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Selection = New System.Windows.Forms.PictureBox()
        Me.AutoSelection = New System.Windows.Forms.PictureBox()
        Me.CustomGroupBox4 = New VSTO_Addins.CustomGroupBox()
        Me.CustomGroupBox10 = New VSTO_Addins.CustomGroupBox()
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.RadioButton10 = New System.Windows.Forms.RadioButton()
        Me.RadioButton9 = New System.Windows.Forms.RadioButton()
        Me.CustomGroupBox5 = New VSTO_Addins.CustomGroupBox()
        Me.CustomPanel1 = New VSTO_Addins.CustomPanel()
        CType(Me.Selection, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AutoSelection, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.CustomGroupBox4.SuspendLayout()
        Me.CustomGroupBox10.SuspendLayout()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.CustomGroupBox5.SuspendLayout()
        Me.SuspendLayout()
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckBox1.Location = New System.Drawing.Point(14, 82)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(122, 21)
        Me.CheckBox1.TabIndex = 184
        Me.CheckBox1.Text = "Keep formatting"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.BackColor = System.Drawing.Color.White
        Me.Button2.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Button2.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button2.Location = New System.Drawing.Point(447, 312)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(62, 26)
        Me.Button2.TabIndex = 190
        Me.Button2.Text = "OK"
        Me.Button2.UseVisualStyleBackColor = False
        '
        'ComboBox1
        '
        Me.ComboBox1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Location = New System.Drawing.Point(15, 312)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(90, 25)
        Me.ComboBox1.TabIndex = 186
        Me.ComboBox1.Text = "SOFTEKO"
        '
        'CheckBox2
        '
        Me.CheckBox2.AutoSize = True
        Me.CheckBox2.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckBox2.Location = New System.Drawing.Point(16, 278)
        Me.CheckBox2.Name = "CheckBox2"
        Me.CheckBox2.Size = New System.Drawing.Size(257, 21)
        Me.CheckBox2.TabIndex = 185
        Me.CheckBox2.Text = "Create a copy of the original worksheet"
        Me.CheckBox2.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(15, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(98, 17)
        Me.Label1.TabIndex = 182
        Me.Label1.Text = "Source Range :"
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.White
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Button1.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.Location = New System.Drawing.Point(528, 312)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(62, 26)
        Me.Button1.TabIndex = 189
        Me.Button1.Text = "Cancel"
        Me.Button1.UseVisualStyleBackColor = False
        '
        'TextBox1
        '
        Me.TextBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TextBox1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.TextBox1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox1.Location = New System.Drawing.Point(15, 44)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(287, 25)
        Me.TextBox1.TabIndex = 204
        '
        'Selection
        '
        Me.Selection.BackColor = System.Drawing.Color.White
        Me.Selection.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Selection.Image = CType(resources.GetObject("Selection.Image"), System.Drawing.Image)
        Me.Selection.Location = New System.Drawing.Point(277, 44)
        Me.Selection.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Selection.Name = "Selection"
        Me.Selection.Size = New System.Drawing.Size(24, 25)
        Me.Selection.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.Selection.TabIndex = 206
        Me.Selection.TabStop = False
        '
        'AutoSelection
        '
        Me.AutoSelection.BackColor = System.Drawing.Color.White
        Me.AutoSelection.Image = CType(resources.GetObject("AutoSelection.Image"), System.Drawing.Image)
        Me.AutoSelection.Location = New System.Drawing.Point(251, 45)
        Me.AutoSelection.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.AutoSelection.Name = "AutoSelection"
        Me.AutoSelection.Size = New System.Drawing.Size(24, 23)
        Me.AutoSelection.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.AutoSelection.TabIndex = 205
        Me.AutoSelection.TabStop = False
        '
        'CustomGroupBox4
        '
        Me.CustomGroupBox4.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.CustomGroupBox4.Controls.Add(Me.CustomGroupBox10)
        Me.CustomGroupBox4.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CustomGroupBox4.Location = New System.Drawing.Point(14, 113)
        Me.CustomGroupBox4.Name = "CustomGroupBox4"
        Me.CustomGroupBox4.Size = New System.Drawing.Size(287, 148)
        Me.CustomGroupBox4.TabIndex = 192
        Me.CustomGroupBox4.TabStop = False
        Me.CustomGroupBox4.Text = "Destination Range"
        '
        'CustomGroupBox10
        '
        Me.CustomGroupBox10.BackColor = System.Drawing.Color.White
        Me.CustomGroupBox10.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.CustomGroupBox10.Controls.Add(Me.PictureBox2)
        Me.CustomGroupBox10.Controls.Add(Me.PictureBox1)
        Me.CustomGroupBox10.Controls.Add(Me.TextBox2)
        Me.CustomGroupBox10.Controls.Add(Me.Label3)
        Me.CustomGroupBox10.Controls.Add(Me.RadioButton10)
        Me.CustomGroupBox10.Controls.Add(Me.RadioButton9)
        Me.CustomGroupBox10.Location = New System.Drawing.Point(1, 22)
        Me.CustomGroupBox10.Name = "CustomGroupBox10"
        Me.CustomGroupBox10.Size = New System.Drawing.Size(285, 126)
        Me.CustomGroupBox10.TabIndex = 0
        Me.CustomGroupBox10.TabStop = False
        '
        'PictureBox2
        '
        Me.PictureBox2.Image = CType(resources.GetObject("PictureBox2.Image"), System.Drawing.Image)
        Me.PictureBox2.Location = New System.Drawing.Point(25, 59)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(14, 14)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox2.TabIndex = 208
        Me.PictureBox2.TabStop = False
        '
        'PictureBox1
        '
        Me.PictureBox1.BackColor = System.Drawing.Color.White
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(245, 85)
        Me.PictureBox1.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(24, 23)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 207
        Me.PictureBox1.TabStop = False
        '
        'TextBox2
        '
        Me.TextBox2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TextBox2.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.TextBox2.Font = New System.Drawing.Font("Segoe UI", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox2.Location = New System.Drawing.Point(26, 84)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(244, 25)
        Me.TextBox2.TabIndex = 206
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(42, 57)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(109, 17)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "Select the range :"
        '
        'RadioButton10
        '
        Me.RadioButton10.AutoSize = True
        Me.RadioButton10.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadioButton10.Location = New System.Drawing.Point(8, 31)
        Me.RadioButton10.Name = "RadioButton10"
        Me.RadioButton10.Size = New System.Drawing.Size(185, 21)
        Me.RadioButton10.TabIndex = 1
        Me.RadioButton10.TabStop = True
        Me.RadioButton10.Text = "Store into a different range"
        Me.RadioButton10.UseVisualStyleBackColor = True
        '
        'RadioButton9
        '
        Me.RadioButton9.AutoSize = True
        Me.RadioButton9.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadioButton9.Location = New System.Drawing.Point(8, 6)
        Me.RadioButton9.Name = "RadioButton9"
        Me.RadioButton9.Size = New System.Drawing.Size(178, 21)
        Me.RadioButton9.TabIndex = 0
        Me.RadioButton9.TabStop = True
        Me.RadioButton9.Text = "Same as the source range"
        Me.RadioButton9.UseVisualStyleBackColor = True
        '
        'CustomGroupBox5
        '
        Me.CustomGroupBox5.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.CustomGroupBox5.Controls.Add(Me.CustomPanel1)
        Me.CustomGroupBox5.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CustomGroupBox5.Location = New System.Drawing.Point(338, 12)
        Me.CustomGroupBox5.Name = "CustomGroupBox5"
        Me.CustomGroupBox5.Size = New System.Drawing.Size(252, 283)
        Me.CustomGroupBox5.TabIndex = 187
        Me.CustomGroupBox5.TabStop = False
        Me.CustomGroupBox5.Text = "Sample Image"
        '
        'CustomPanel1
        '
        Me.CustomPanel1.BackColor = System.Drawing.Color.White
        Me.CustomPanel1.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.CustomPanel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.CustomPanel1.BorderWidth = 1
        Me.CustomPanel1.Location = New System.Drawing.Point(1, 30)
        Me.CustomPanel1.Name = "CustomPanel1"
        Me.CustomPanel1.Size = New System.Drawing.Size(251, 253)
        Me.CustomPanel1.TabIndex = 0
        '
        'Form16PasteintoVisibleRange
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(614, 355)
        Me.Controls.Add(Me.Selection)
        Me.Controls.Add(Me.AutoSelection)
        Me.Controls.Add(Me.CustomGroupBox4)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.CheckBox1)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.ComboBox1)
        Me.Controls.Add(Me.CheckBox2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.CustomGroupBox5)
        Me.Controls.Add(Me.Button1)
        Me.Name = "Form16PasteintoVisibleRange"
        Me.Text = "Paste into Visible Range"
        CType(Me.Selection, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AutoSelection, System.ComponentModel.ISupportInitialize).EndInit()
        Me.CustomGroupBox4.ResumeLayout(False)
        Me.CustomGroupBox10.ResumeLayout(False)
        Me.CustomGroupBox10.PerformLayout()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.CustomGroupBox5.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents CustomPanel1 As CustomPanel
    Friend WithEvents Label3 As Windows.Forms.Label
    Friend WithEvents RadioButton10 As Windows.Forms.RadioButton
    Friend WithEvents RadioButton9 As Windows.Forms.RadioButton
    Friend WithEvents CustomGroupBox10 As CustomGroupBox
    Friend WithEvents CustomGroupBox4 As CustomGroupBox
    Friend WithEvents CheckBox1 As Windows.Forms.CheckBox
    Friend WithEvents Button2 As Windows.Forms.Button
    Friend WithEvents ComboBox1 As Windows.Forms.ComboBox
    Friend WithEvents CheckBox2 As Windows.Forms.CheckBox
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents CustomGroupBox5 As CustomGroupBox
    Friend WithEvents Button1 As Windows.Forms.Button
    Friend WithEvents TextBox1 As Windows.Forms.TextBox
    Friend WithEvents PictureBox1 As Windows.Forms.PictureBox
    Friend WithEvents TextBox2 As Windows.Forms.TextBox
    Friend WithEvents Selection As Windows.Forms.PictureBox
    Friend WithEvents AutoSelection As Windows.Forms.PictureBox
    Friend WithEvents PictureBox2 As Windows.Forms.PictureBox
End Class
