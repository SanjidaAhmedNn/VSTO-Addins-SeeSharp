<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form11SwapRanges
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form11SwapRanges))
        Me.CB_CopyWs = New System.Windows.Forms.CheckBox()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.btnOK = New System.Windows.Forms.Button()
        Me.PictureBox7 = New System.Windows.Forms.PictureBox()
        Me.CB_KeepFormatting = New System.Windows.Forms.CheckBox()
        Me.CustomGroupBox2 = New VSTO_Addins.CustomGroupBox()
        Me.AutoSelection2 = New System.Windows.Forms.PictureBox()
        Me.rngSelection2 = New System.Windows.Forms.PictureBox()
        Me.txtSourceRange2 = New System.Windows.Forms.TextBox()
        Me.lblSourceRng2 = New System.Windows.Forms.Label()
        Me.AutoSelection1 = New System.Windows.Forms.PictureBox()
        Me.rngSelection1 = New System.Windows.Forms.PictureBox()
        Me.txtSourceRange1 = New System.Windows.Forms.TextBox()
        Me.lblSourceRng1 = New System.Windows.Forms.Label()
        Me.CustomGroupBox6 = New VSTO_Addins.CustomGroupBox()
        Me.CP_OutputRng = New VSTO_Addins.CustomPanel()
        Me.CustomGroupBox5 = New VSTO_Addins.CustomGroupBox()
        Me.CP_InputRng = New VSTO_Addins.CustomPanel()
        Me.CustomGroupBox1 = New VSTO_Addins.CustomGroupBox()
        Me.CustomGroupBox7 = New VSTO_Addins.CustomGroupBox()
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        Me.radBtnValues = New System.Windows.Forms.RadioButton()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.PictureBox5 = New System.Windows.Forms.PictureBox()
        Me.radBtnKeepRef = New System.Windows.Forms.RadioButton()
        Me.radBtnAdjustRef = New System.Windows.Forms.RadioButton()
        CType(Me.PictureBox7, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.CustomGroupBox2.SuspendLayout()
        CType(Me.AutoSelection2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.rngSelection2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AutoSelection1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.rngSelection1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.CustomGroupBox6.SuspendLayout()
        Me.CustomGroupBox5.SuspendLayout()
        Me.CustomGroupBox1.SuspendLayout()
        Me.CustomGroupBox7.SuspendLayout()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox5, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'CB_CopyWs
        '
        Me.CB_CopyWs.AutoSize = True
        Me.CB_CopyWs.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CB_CopyWs.Location = New System.Drawing.Point(15, 309)
        Me.CB_CopyWs.Name = "CB_CopyWs"
        Me.CB_CopyWs.Size = New System.Drawing.Size(257, 21)
        Me.CB_CopyWs.TabIndex = 151
        Me.CB_CopyWs.Text = "Create a copy of the original worksheet"
        Me.CB_CopyWs.UseVisualStyleBackColor = True
        '
        'ComboBox1
        '
        Me.ComboBox1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Location = New System.Drawing.Point(15, 346)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(154, 25)
        Me.ComboBox1.TabIndex = 152
        Me.ComboBox1.Text = "SOFTEKO"
        '
        'btnCancel
        '
        Me.btnCancel.BackColor = System.Drawing.Color.White
        Me.btnCancel.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnCancel.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCancel.Location = New System.Drawing.Point(518, 346)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(62, 26)
        Me.btnCancel.TabIndex = 155
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = False
        '
        'btnOK
        '
        Me.btnOK.BackColor = System.Drawing.Color.White
        Me.btnOK.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnOK.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOK.Location = New System.Drawing.Point(446, 346)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(62, 26)
        Me.btnOK.TabIndex = 156
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = False
        '
        'PictureBox7
        '
        Me.PictureBox7.Image = CType(resources.GetObject("PictureBox7.Image"), System.Drawing.Image)
        Me.PictureBox7.Location = New System.Drawing.Point(437, 145)
        Me.PictureBox7.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.PictureBox7.Name = "PictureBox7"
        Me.PictureBox7.Size = New System.Drawing.Size(43, 49)
        Me.PictureBox7.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox7.TabIndex = 157
        Me.PictureBox7.TabStop = False
        '
        'CB_KeepFormatting
        '
        Me.CB_KeepFormatting.AutoSize = True
        Me.CB_KeepFormatting.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CB_KeepFormatting.Location = New System.Drawing.Point(15, 280)
        Me.CB_KeepFormatting.Name = "CB_KeepFormatting"
        Me.CB_KeepFormatting.Size = New System.Drawing.Size(122, 21)
        Me.CB_KeepFormatting.TabIndex = 150
        Me.CB_KeepFormatting.Text = "Keep formatting"
        Me.CB_KeepFormatting.UseVisualStyleBackColor = True
        '
        'CustomGroupBox2
        '
        Me.CustomGroupBox2.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.CustomGroupBox2.Controls.Add(Me.AutoSelection2)
        Me.CustomGroupBox2.Controls.Add(Me.rngSelection2)
        Me.CustomGroupBox2.Controls.Add(Me.txtSourceRange2)
        Me.CustomGroupBox2.Controls.Add(Me.lblSourceRng2)
        Me.CustomGroupBox2.Controls.Add(Me.AutoSelection1)
        Me.CustomGroupBox2.Controls.Add(Me.rngSelection1)
        Me.CustomGroupBox2.Controls.Add(Me.txtSourceRange1)
        Me.CustomGroupBox2.Controls.Add(Me.lblSourceRng1)
        Me.CustomGroupBox2.Location = New System.Drawing.Point(15, 15)
        Me.CustomGroupBox2.Name = "CustomGroupBox2"
        Me.CustomGroupBox2.Size = New System.Drawing.Size(278, 136)
        Me.CustomGroupBox2.TabIndex = 164
        Me.CustomGroupBox2.TabStop = False
        '
        'AutoSelection2
        '
        Me.AutoSelection2.Image = CType(resources.GetObject("AutoSelection2.Image"), System.Drawing.Image)
        Me.AutoSelection2.Location = New System.Drawing.Point(206, 96)
        Me.AutoSelection2.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.AutoSelection2.Name = "AutoSelection2"
        Me.AutoSelection2.Size = New System.Drawing.Size(27, 23)
        Me.AutoSelection2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.AutoSelection2.TabIndex = 171
        Me.AutoSelection2.TabStop = False
        '
        'rngSelection2
        '
        Me.rngSelection2.Image = CType(resources.GetObject("rngSelection2.Image"), System.Drawing.Image)
        Me.rngSelection2.Location = New System.Drawing.Point(241, 95)
        Me.rngSelection2.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.rngSelection2.Name = "rngSelection2"
        Me.rngSelection2.Size = New System.Drawing.Size(27, 23)
        Me.rngSelection2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.rngSelection2.TabIndex = 170
        Me.rngSelection2.TabStop = False
        '
        'txtSourceRange2
        '
        Me.txtSourceRange2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtSourceRange2.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtSourceRange2.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSourceRange2.Location = New System.Drawing.Point(8, 93)
        Me.txtSourceRange2.Name = "txtSourceRange2"
        Me.txtSourceRange2.Size = New System.Drawing.Size(260, 25)
        Me.txtSourceRange2.TabIndex = 169
        '
        'lblSourceRng2
        '
        Me.lblSourceRng2.AutoSize = True
        Me.lblSourceRng2.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSourceRng2.Location = New System.Drawing.Point(8, 65)
        Me.lblSourceRng2.Name = "lblSourceRng2"
        Me.lblSourceRng2.Size = New System.Drawing.Size(256, 17)
        Me.lblSourceRng2.TabIndex = 168
        Me.lblSourceRng2.Text = "2nd Source Range (X rows x Y columns) :"
        '
        'AutoSelection1
        '
        Me.AutoSelection1.Image = CType(resources.GetObject("AutoSelection1.Image"), System.Drawing.Image)
        Me.AutoSelection1.Location = New System.Drawing.Point(206, 31)
        Me.AutoSelection1.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.AutoSelection1.Name = "AutoSelection1"
        Me.AutoSelection1.Size = New System.Drawing.Size(27, 23)
        Me.AutoSelection1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.AutoSelection1.TabIndex = 167
        Me.AutoSelection1.TabStop = False
        '
        'rngSelection1
        '
        Me.rngSelection1.Image = CType(resources.GetObject("rngSelection1.Image"), System.Drawing.Image)
        Me.rngSelection1.Location = New System.Drawing.Point(241, 31)
        Me.rngSelection1.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.rngSelection1.Name = "rngSelection1"
        Me.rngSelection1.Size = New System.Drawing.Size(27, 23)
        Me.rngSelection1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.rngSelection1.TabIndex = 166
        Me.rngSelection1.TabStop = False
        '
        'txtSourceRange1
        '
        Me.txtSourceRange1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtSourceRange1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtSourceRange1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSourceRange1.Location = New System.Drawing.Point(8, 30)
        Me.txtSourceRange1.Name = "txtSourceRange1"
        Me.txtSourceRange1.Size = New System.Drawing.Size(260, 25)
        Me.txtSourceRange1.TabIndex = 165
        '
        'lblSourceRng1
        '
        Me.lblSourceRng1.AutoSize = True
        Me.lblSourceRng1.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSourceRng1.Location = New System.Drawing.Point(8, 6)
        Me.lblSourceRng1.Name = "lblSourceRng1"
        Me.lblSourceRng1.Size = New System.Drawing.Size(249, 17)
        Me.lblSourceRng1.TabIndex = 164
        Me.lblSourceRng1.Text = "1st Source Range (X rows x Y columns) :"
        '
        'CustomGroupBox6
        '
        Me.CustomGroupBox6.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.CustomGroupBox6.Controls.Add(Me.CP_OutputRng)
        Me.CustomGroupBox6.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CustomGroupBox6.Location = New System.Drawing.Point(330, 195)
        Me.CustomGroupBox6.Name = "CustomGroupBox6"
        Me.CustomGroupBox6.Size = New System.Drawing.Size(252, 135)
        Me.CustomGroupBox6.TabIndex = 154
        Me.CustomGroupBox6.TabStop = False
        Me.CustomGroupBox6.Text = "Expected Output"
        '
        'CP_OutputRng
        '
        Me.CP_OutputRng.BackColor = System.Drawing.Color.White
        Me.CP_OutputRng.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.CP_OutputRng.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.CP_OutputRng.BorderWidth = 1
        Me.CP_OutputRng.Location = New System.Drawing.Point(1, 30)
        Me.CP_OutputRng.Name = "CP_OutputRng"
        Me.CP_OutputRng.Size = New System.Drawing.Size(250, 105)
        Me.CP_OutputRng.TabIndex = 11
        '
        'CustomGroupBox5
        '
        Me.CustomGroupBox5.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.CustomGroupBox5.Controls.Add(Me.CP_InputRng)
        Me.CustomGroupBox5.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CustomGroupBox5.Location = New System.Drawing.Point(329, 4)
        Me.CustomGroupBox5.Name = "CustomGroupBox5"
        Me.CustomGroupBox5.Size = New System.Drawing.Size(252, 135)
        Me.CustomGroupBox5.TabIndex = 153
        Me.CustomGroupBox5.TabStop = False
        Me.CustomGroupBox5.Text = "Input Range"
        '
        'CP_InputRng
        '
        Me.CP_InputRng.BackColor = System.Drawing.Color.White
        Me.CP_InputRng.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.CP_InputRng.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.CP_InputRng.BorderWidth = 1
        Me.CP_InputRng.Location = New System.Drawing.Point(1, 30)
        Me.CP_InputRng.Name = "CP_InputRng"
        Me.CP_InputRng.Size = New System.Drawing.Size(250, 105)
        Me.CP_InputRng.TabIndex = 0
        '
        'CustomGroupBox1
        '
        Me.CustomGroupBox1.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.CustomGroupBox1.Controls.Add(Me.CustomGroupBox7)
        Me.CustomGroupBox1.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CustomGroupBox1.Location = New System.Drawing.Point(15, 162)
        Me.CustomGroupBox1.Name = "CustomGroupBox1"
        Me.CustomGroupBox1.Size = New System.Drawing.Size(281, 110)
        Me.CustomGroupBox1.TabIndex = 148
        Me.CustomGroupBox1.TabStop = False
        Me.CustomGroupBox1.Text = "Swap Type"
        '
        'CustomGroupBox7
        '
        Me.CustomGroupBox7.BackColor = System.Drawing.Color.White
        Me.CustomGroupBox7.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.CustomGroupBox7.Controls.Add(Me.PictureBox2)
        Me.CustomGroupBox7.Controls.Add(Me.radBtnValues)
        Me.CustomGroupBox7.Controls.Add(Me.PictureBox1)
        Me.CustomGroupBox7.Controls.Add(Me.PictureBox5)
        Me.CustomGroupBox7.Controls.Add(Me.radBtnKeepRef)
        Me.CustomGroupBox7.Controls.Add(Me.radBtnAdjustRef)
        Me.CustomGroupBox7.Location = New System.Drawing.Point(1, 22)
        Me.CustomGroupBox7.Name = "CustomGroupBox7"
        Me.CustomGroupBox7.Size = New System.Drawing.Size(280, 88)
        Me.CustomGroupBox7.TabIndex = 0
        Me.CustomGroupBox7.TabStop = False
        '
        'PictureBox2
        '
        Me.PictureBox2.Image = CType(resources.GetObject("PictureBox2.Image"), System.Drawing.Image)
        Me.PictureBox2.Location = New System.Drawing.Point(245, 61)
        Me.PictureBox2.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(20, 20)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox2.TabIndex = 130
        Me.PictureBox2.TabStop = False
        '
        'radBtnValues
        '
        Me.radBtnValues.AutoSize = True
        Me.radBtnValues.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.radBtnValues.Location = New System.Drawing.Point(8, 8)
        Me.radBtnValues.Name = "radBtnValues"
        Me.radBtnValues.Size = New System.Drawing.Size(93, 21)
        Me.radBtnValues.TabIndex = 129
        Me.radBtnValues.TabStop = True
        Me.radBtnValues.Text = "Values Only"
        Me.radBtnValues.UseVisualStyleBackColor = True
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(245, 34)
        Me.PictureBox1.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(20, 20)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 128
        Me.PictureBox1.TabStop = False
        '
        'PictureBox5
        '
        Me.PictureBox5.Image = CType(resources.GetObject("PictureBox5.Image"), System.Drawing.Image)
        Me.PictureBox5.Location = New System.Drawing.Point(245, 7)
        Me.PictureBox5.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.PictureBox5.Name = "PictureBox5"
        Me.PictureBox5.Size = New System.Drawing.Size(20, 20)
        Me.PictureBox5.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox5.TabIndex = 127
        Me.PictureBox5.TabStop = False
        '
        'radBtnKeepRef
        '
        Me.radBtnKeepRef.AutoSize = True
        Me.radBtnKeepRef.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.radBtnKeepRef.Location = New System.Drawing.Point(8, 60)
        Me.radBtnKeepRef.Name = "radBtnKeepRef"
        Me.radBtnKeepRef.Size = New System.Drawing.Size(143, 21)
        Me.radBtnKeepRef.TabIndex = 1
        Me.radBtnKeepRef.TabStop = True
        Me.radBtnKeepRef.Text = "Keep Cell Reference"
        Me.radBtnKeepRef.UseVisualStyleBackColor = True
        '
        'radBtnAdjustRef
        '
        Me.radBtnAdjustRef.AutoSize = True
        Me.radBtnAdjustRef.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.radBtnAdjustRef.Location = New System.Drawing.Point(8, 34)
        Me.radBtnAdjustRef.Name = "radBtnAdjustRef"
        Me.radBtnAdjustRef.Size = New System.Drawing.Size(149, 21)
        Me.radBtnAdjustRef.TabIndex = 0
        Me.radBtnAdjustRef.TabStop = True
        Me.radBtnAdjustRef.Text = "Adjust Cell Reference"
        Me.radBtnAdjustRef.UseVisualStyleBackColor = True
        '
        'Form11SwapRanges
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.ClientSize = New System.Drawing.Size(596, 387)
        Me.Controls.Add(Me.CustomGroupBox2)
        Me.Controls.Add(Me.CustomGroupBox6)
        Me.Controls.Add(Me.CustomGroupBox5)
        Me.Controls.Add(Me.CB_KeepFormatting)
        Me.Controls.Add(Me.PictureBox7)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.ComboBox1)
        Me.Controls.Add(Me.CB_CopyWs)
        Me.Controls.Add(Me.CustomGroupBox1)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Form11SwapRanges"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Swap"
        CType(Me.PictureBox7, System.ComponentModel.ISupportInitialize).EndInit()
        Me.CustomGroupBox2.ResumeLayout(False)
        Me.CustomGroupBox2.PerformLayout()
        CType(Me.AutoSelection2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.rngSelection2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AutoSelection1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.rngSelection1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.CustomGroupBox6.ResumeLayout(False)
        Me.CustomGroupBox5.ResumeLayout(False)
        Me.CustomGroupBox1.ResumeLayout(False)
        Me.CustomGroupBox7.ResumeLayout(False)
        Me.CustomGroupBox7.PerformLayout()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox5, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents lblSourceRng1 As Windows.Forms.Label
    Friend WithEvents txtSourceRange1 As Windows.Forms.TextBox
    Friend WithEvents rngSelection1 As Windows.Forms.PictureBox
    Friend WithEvents AutoSelection1 As Windows.Forms.PictureBox
    Friend WithEvents lblSourceRng2 As Windows.Forms.Label
    Friend WithEvents txtSourceRange2 As Windows.Forms.TextBox
    Friend WithEvents rngSelection2 As Windows.Forms.PictureBox
    Friend WithEvents AutoSelection2 As Windows.Forms.PictureBox
    Friend WithEvents CustomGroupBox2 As CustomGroupBox
    Friend WithEvents CustomGroupBox1 As CustomGroupBox
    Friend WithEvents CustomGroupBox7 As CustomGroupBox
    Friend WithEvents PictureBox2 As Windows.Forms.PictureBox
    Friend WithEvents radBtnValues As Windows.Forms.RadioButton
    Friend WithEvents PictureBox1 As Windows.Forms.PictureBox
    Friend WithEvents PictureBox5 As Windows.Forms.PictureBox
    Friend WithEvents radBtnKeepRef As Windows.Forms.RadioButton
    Friend WithEvents radBtnAdjustRef As Windows.Forms.RadioButton
    Friend WithEvents CB_CopyWs As Windows.Forms.CheckBox
    Friend WithEvents ComboBox1 As Windows.Forms.ComboBox
    Friend WithEvents btnCancel As Windows.Forms.Button
    Friend WithEvents btnOK As Windows.Forms.Button
    Friend WithEvents PictureBox7 As Windows.Forms.PictureBox
    Friend WithEvents CB_KeepFormatting As Windows.Forms.CheckBox
    Friend WithEvents CustomGroupBox5 As CustomGroupBox
    Friend WithEvents CP_InputRng As CustomPanel
    Friend WithEvents CP_OutputRng As CustomPanel
    Friend WithEvents CustomGroupBox6 As CustomGroupBox
End Class
