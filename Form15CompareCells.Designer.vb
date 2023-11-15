<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form15CompareCells
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form15CompareCells))
        Me.checkBoxFormatting = New System.Windows.Forms.CheckBox()
        Me.btnOK = New System.Windows.Forms.Button()
        Me.btnCanecl = New System.Windows.Forms.Button()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.checkBoxCopyWs = New System.Windows.Forms.CheckBox()
        Me.checkBoxCase = New System.Windows.Forms.CheckBox()
        Me.CD_Fill_Background = New System.Windows.Forms.ColorDialog()
        Me.CD_Fill_Font = New System.Windows.Forms.ColorDialog()
        Me.CustomPanel1 = New VSTO_Addins.CustomPanel()
        Me.CustomGroupBox5 = New VSTO_Addins.CustomGroupBox()
        Me.CP_Input_Range2 = New VSTO_Addins.CustomPanel()
        Me.GB_Input_Range = New VSTO_Addins.CustomGroupBox()
        Me.CP_Input_Range1 = New VSTO_Addins.CustomPanel()
        Me.GB_Expected_Output = New VSTO_Addins.CustomGroupBox()
        Me.CP_Output_Range = New VSTO_Addins.CustomPanel()
        Me.PictureBox7 = New System.Windows.Forms.PictureBox()
        Me.GB_Display_Result = New VSTO_Addins.CustomGroupBox()
        Me.CustomGroupBox4 = New VSTO_Addins.CustomGroupBox()
        Me.CbFillFont = New System.Windows.Forms.ComboBox()
        Me.CBFillBackground = New System.Windows.Forms.ComboBox()
        Me.checkBoxFillFont = New System.Windows.Forms.CheckBox()
        Me.checkBoxFillBack = New System.Windows.Forms.CheckBox()
        Me.CustomGroupBox1 = New VSTO_Addins.CustomGroupBox()
        Me.CustomGroupBox7 = New VSTO_Addins.CustomGroupBox()
        Me.radBtnSameValues = New System.Windows.Forms.RadioButton()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.PictureBox5 = New System.Windows.Forms.PictureBox()
        Me.radBtnDifferentValues = New System.Windows.Forms.RadioButton()
        Me.CustomGroupBox2 = New VSTO_Addins.CustomGroupBox()
        Me.rngSelection2 = New System.Windows.Forms.PictureBox()
        Me.rngSelection1 = New System.Windows.Forms.PictureBox()
        Me.AutoSelection2 = New System.Windows.Forms.PictureBox()
        Me.AutoSelection1 = New System.Windows.Forms.PictureBox()
        Me.txtSourceRange2 = New System.Windows.Forms.TextBox()
        Me.txtSourceRange1 = New System.Windows.Forms.TextBox()
        Me.lblSourceRng2 = New System.Windows.Forms.Label()
        Me.lblSourceRng1 = New System.Windows.Forms.Label()
        Me.CustomPanel1.SuspendLayout()
        Me.CustomGroupBox5.SuspendLayout()
        Me.GB_Input_Range.SuspendLayout()
        Me.GB_Expected_Output.SuspendLayout()
        CType(Me.PictureBox7, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GB_Display_Result.SuspendLayout()
        Me.CustomGroupBox4.SuspendLayout()
        Me.CustomGroupBox1.SuspendLayout()
        Me.CustomGroupBox7.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox5, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.CustomGroupBox2.SuspendLayout()
        CType(Me.rngSelection2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.rngSelection1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AutoSelection2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AutoSelection1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'checkBoxFormatting
        '
        Me.checkBoxFormatting.AutoSize = True
        Me.checkBoxFormatting.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.checkBoxFormatting.Location = New System.Drawing.Point(15, 259)
        Me.checkBoxFormatting.Name = "checkBoxFormatting"
        Me.checkBoxFormatting.Size = New System.Drawing.Size(122, 21)
        Me.checkBoxFormatting.TabIndex = 166
        Me.checkBoxFormatting.Text = "Keep formatting"
        Me.checkBoxFormatting.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.BackColor = System.Drawing.Color.White
        Me.btnOK.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnOK.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOK.Location = New System.Drawing.Point(611, 423)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(62, 26)
        Me.btnOK.TabIndex = 172
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = False
        '
        'btnCanecl
        '
        Me.btnCanecl.BackColor = System.Drawing.Color.White
        Me.btnCanecl.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnCanecl.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCanecl.Location = New System.Drawing.Point(692, 423)
        Me.btnCanecl.Name = "btnCanecl"
        Me.btnCanecl.Size = New System.Drawing.Size(62, 26)
        Me.btnCanecl.TabIndex = 171
        Me.btnCanecl.Text = "Cancel"
        Me.btnCanecl.UseVisualStyleBackColor = False
        '
        'ComboBox1
        '
        Me.ComboBox1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Location = New System.Drawing.Point(15, 423)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(154, 25)
        Me.ComboBox1.TabIndex = 168
        Me.ComboBox1.Text = "SOFTEKO"
        '
        'checkBoxCopyWs
        '
        Me.checkBoxCopyWs.AutoSize = True
        Me.checkBoxCopyWs.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.checkBoxCopyWs.Location = New System.Drawing.Point(15, 392)
        Me.checkBoxCopyWs.Name = "checkBoxCopyWs"
        Me.checkBoxCopyWs.Size = New System.Drawing.Size(257, 21)
        Me.checkBoxCopyWs.TabIndex = 167
        Me.checkBoxCopyWs.Text = "Create a copy of the original worksheet"
        Me.checkBoxCopyWs.UseVisualStyleBackColor = True
        '
        'checkBoxCase
        '
        Me.checkBoxCase.AutoSize = True
        Me.checkBoxCase.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.checkBoxCase.Location = New System.Drawing.Point(179, 259)
        Me.checkBoxCase.Name = "checkBoxCase"
        Me.checkBoxCase.Size = New System.Drawing.Size(108, 21)
        Me.checkBoxCase.TabIndex = 176
        Me.checkBoxCase.Text = "Case sensitive"
        Me.checkBoxCase.UseVisualStyleBackColor = True
        '
        'CustomPanel1
        '
        Me.CustomPanel1.BorderColor = System.Drawing.Color.Empty
        Me.CustomPanel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.CustomPanel1.BorderWidth = 0
        Me.CustomPanel1.Controls.Add(Me.CustomGroupBox5)
        Me.CustomPanel1.Controls.Add(Me.GB_Input_Range)
        Me.CustomPanel1.Controls.Add(Me.GB_Expected_Output)
        Me.CustomPanel1.Controls.Add(Me.PictureBox7)
        Me.CustomPanel1.Location = New System.Drawing.Point(322, 15)
        Me.CustomPanel1.Name = "CustomPanel1"
        Me.CustomPanel1.Size = New System.Drawing.Size(432, 390)
        Me.CustomPanel1.TabIndex = 177
        '
        'CustomGroupBox5
        '
        Me.CustomGroupBox5.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.CustomGroupBox5.Controls.Add(Me.CP_Input_Range2)
        Me.CustomGroupBox5.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CustomGroupBox5.Location = New System.Drawing.Point(221, 15)
        Me.CustomGroupBox5.Name = "CustomGroupBox5"
        Me.CustomGroupBox5.Size = New System.Drawing.Size(192, 142)
        Me.CustomGroupBox5.TabIndex = 161
        Me.CustomGroupBox5.TabStop = False
        Me.CustomGroupBox5.Text = "2nd Input Range"
        '
        'CP_Input_Range2
        '
        Me.CP_Input_Range2.BackColor = System.Drawing.Color.White
        Me.CP_Input_Range2.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.CP_Input_Range2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.CP_Input_Range2.BorderWidth = 1
        Me.CP_Input_Range2.Location = New System.Drawing.Point(1, 30)
        Me.CP_Input_Range2.Name = "CP_Input_Range2"
        Me.CP_Input_Range2.Size = New System.Drawing.Size(190, 112)
        Me.CP_Input_Range2.TabIndex = 0
        '
        'GB_Input_Range
        '
        Me.GB_Input_Range.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.GB_Input_Range.Controls.Add(Me.CP_Input_Range1)
        Me.GB_Input_Range.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GB_Input_Range.Location = New System.Drawing.Point(15, 15)
        Me.GB_Input_Range.Name = "GB_Input_Range"
        Me.GB_Input_Range.Size = New System.Drawing.Size(192, 142)
        Me.GB_Input_Range.TabIndex = 160
        Me.GB_Input_Range.TabStop = False
        Me.GB_Input_Range.Text = "1st Input Range"
        '
        'CP_Input_Range1
        '
        Me.CP_Input_Range1.BackColor = System.Drawing.Color.White
        Me.CP_Input_Range1.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.CP_Input_Range1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.CP_Input_Range1.BorderWidth = 1
        Me.CP_Input_Range1.Location = New System.Drawing.Point(1, 30)
        Me.CP_Input_Range1.Name = "CP_Input_Range1"
        Me.CP_Input_Range1.Size = New System.Drawing.Size(190, 112)
        Me.CP_Input_Range1.TabIndex = 0
        '
        'GB_Expected_Output
        '
        Me.GB_Expected_Output.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.GB_Expected_Output.Controls.Add(Me.CP_Output_Range)
        Me.GB_Expected_Output.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GB_Expected_Output.Location = New System.Drawing.Point(129, 229)
        Me.GB_Expected_Output.Name = "GB_Expected_Output"
        Me.GB_Expected_Output.Size = New System.Drawing.Size(192, 142)
        Me.GB_Expected_Output.TabIndex = 161
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
        Me.CP_Output_Range.Size = New System.Drawing.Size(190, 112)
        Me.CP_Output_Range.TabIndex = 11
        '
        'PictureBox7
        '
        Me.PictureBox7.Image = CType(resources.GetObject("PictureBox7.Image"), System.Drawing.Image)
        Me.PictureBox7.Location = New System.Drawing.Point(194, 168)
        Me.PictureBox7.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.PictureBox7.Name = "PictureBox7"
        Me.PictureBox7.Size = New System.Drawing.Size(50, 60)
        Me.PictureBox7.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.PictureBox7.TabIndex = 173
        Me.PictureBox7.TabStop = False
        '
        'GB_Display_Result
        '
        Me.GB_Display_Result.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.GB_Display_Result.Controls.Add(Me.CustomGroupBox4)
        Me.GB_Display_Result.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GB_Display_Result.Location = New System.Drawing.Point(15, 290)
        Me.GB_Display_Result.Name = "GB_Display_Result"
        Me.GB_Display_Result.Size = New System.Drawing.Size(281, 94)
        Me.GB_Display_Result.TabIndex = 175
        Me.GB_Display_Result.TabStop = False
        Me.GB_Display_Result.Text = "Display Result"
        '
        'CustomGroupBox4
        '
        Me.CustomGroupBox4.BackColor = System.Drawing.Color.White
        Me.CustomGroupBox4.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.CustomGroupBox4.Controls.Add(Me.CbFillFont)
        Me.CustomGroupBox4.Controls.Add(Me.CBFillBackground)
        Me.CustomGroupBox4.Controls.Add(Me.checkBoxFillFont)
        Me.CustomGroupBox4.Controls.Add(Me.checkBoxFillBack)
        Me.CustomGroupBox4.Location = New System.Drawing.Point(1, 22)
        Me.CustomGroupBox4.Name = "CustomGroupBox4"
        Me.CustomGroupBox4.Size = New System.Drawing.Size(279, 72)
        Me.CustomGroupBox4.TabIndex = 0
        Me.CustomGroupBox4.TabStop = False
        '
        'CbFillFont
        '
        Me.CbFillFont.BackColor = System.Drawing.Color.MidnightBlue
        Me.CbFillFont.DropDownHeight = 1
        Me.CbFillFont.DropDownWidth = 1
        Me.CbFillFont.ForeColor = System.Drawing.Color.Navy
        Me.CbFillFont.FormattingEnabled = True
        Me.CbFillFont.IntegralHeight = False
        Me.CbFillFont.Location = New System.Drawing.Point(162, 35)
        Me.CbFillFont.Name = "CbFillFont"
        Me.CbFillFont.Size = New System.Drawing.Size(110, 25)
        Me.CbFillFont.TabIndex = 134
        '
        'CBFillBackground
        '
        Me.CBFillBackground.BackColor = System.Drawing.Color.LightSteelBlue
        Me.CBFillBackground.DropDownHeight = 1
        Me.CBFillBackground.DropDownWidth = 1
        Me.CBFillBackground.FormattingEnabled = True
        Me.CBFillBackground.IntegralHeight = False
        Me.CBFillBackground.Location = New System.Drawing.Point(8, 35)
        Me.CBFillBackground.Name = "CBFillBackground"
        Me.CBFillBackground.Size = New System.Drawing.Size(110, 25)
        Me.CBFillBackground.TabIndex = 133
        '
        'checkBoxFillFont
        '
        Me.checkBoxFillFont.AutoSize = True
        Me.checkBoxFillFont.Location = New System.Drawing.Point(163, 8)
        Me.checkBoxFillFont.Name = "checkBoxFillFont"
        Me.checkBoxFillFont.Size = New System.Drawing.Size(106, 21)
        Me.checkBoxFillFont.TabIndex = 132
        Me.checkBoxFillFont.Text = "Fill font color"
        Me.checkBoxFillFont.UseVisualStyleBackColor = True
        '
        'checkBoxFillBack
        '
        Me.checkBoxFillBack.AutoSize = True
        Me.checkBoxFillBack.Location = New System.Drawing.Point(8, 8)
        Me.checkBoxFillBack.Name = "checkBoxFillBack"
        Me.checkBoxFillBack.Size = New System.Drawing.Size(120, 21)
        Me.checkBoxFillBack.TabIndex = 131
        Me.checkBoxFillBack.Text = "Fill background"
        Me.checkBoxFillBack.UseVisualStyleBackColor = True
        '
        'CustomGroupBox1
        '
        Me.CustomGroupBox1.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.CustomGroupBox1.Controls.Add(Me.CustomGroupBox7)
        Me.CustomGroupBox1.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CustomGroupBox1.Location = New System.Drawing.Point(15, 163)
        Me.CustomGroupBox1.Name = "CustomGroupBox1"
        Me.CustomGroupBox1.Size = New System.Drawing.Size(281, 86)
        Me.CustomGroupBox1.TabIndex = 165
        Me.CustomGroupBox1.TabStop = False
        Me.CustomGroupBox1.Text = "Compare Type"
        '
        'CustomGroupBox7
        '
        Me.CustomGroupBox7.BackColor = System.Drawing.Color.White
        Me.CustomGroupBox7.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.CustomGroupBox7.Controls.Add(Me.radBtnSameValues)
        Me.CustomGroupBox7.Controls.Add(Me.PictureBox1)
        Me.CustomGroupBox7.Controls.Add(Me.PictureBox5)
        Me.CustomGroupBox7.Controls.Add(Me.radBtnDifferentValues)
        Me.CustomGroupBox7.Location = New System.Drawing.Point(1, 22)
        Me.CustomGroupBox7.Name = "CustomGroupBox7"
        Me.CustomGroupBox7.Size = New System.Drawing.Size(279, 64)
        Me.CustomGroupBox7.TabIndex = 0
        Me.CustomGroupBox7.TabStop = False
        '
        'radBtnSameValues
        '
        Me.radBtnSameValues.AutoSize = True
        Me.radBtnSameValues.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.radBtnSameValues.Location = New System.Drawing.Point(8, 8)
        Me.radBtnSameValues.Name = "radBtnSameValues"
        Me.radBtnSameValues.Size = New System.Drawing.Size(157, 21)
        Me.radBtnSameValues.TabIndex = 129
        Me.radBtnSameValues.TabStop = True
        Me.radBtnSameValues.Text = "Cells with Same Values"
        Me.radBtnSameValues.UseVisualStyleBackColor = True
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(245, 36)
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
        Me.PictureBox5.Location = New System.Drawing.Point(245, 8)
        Me.PictureBox5.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.PictureBox5.Name = "PictureBox5"
        Me.PictureBox5.Size = New System.Drawing.Size(20, 20)
        Me.PictureBox5.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox5.TabIndex = 127
        Me.PictureBox5.TabStop = False
        '
        'radBtnDifferentValues
        '
        Me.radBtnDifferentValues.AutoSize = True
        Me.radBtnDifferentValues.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.radBtnDifferentValues.Location = New System.Drawing.Point(8, 34)
        Me.radBtnDifferentValues.Name = "radBtnDifferentValues"
        Me.radBtnDifferentValues.Size = New System.Drawing.Size(175, 21)
        Me.radBtnDifferentValues.TabIndex = 0
        Me.radBtnDifferentValues.TabStop = True
        Me.radBtnDifferentValues.Text = "Cells with Different Values"
        Me.radBtnDifferentValues.UseVisualStyleBackColor = True
        '
        'CustomGroupBox2
        '
        Me.CustomGroupBox2.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.CustomGroupBox2.Controls.Add(Me.rngSelection2)
        Me.CustomGroupBox2.Controls.Add(Me.rngSelection1)
        Me.CustomGroupBox2.Controls.Add(Me.AutoSelection2)
        Me.CustomGroupBox2.Controls.Add(Me.AutoSelection1)
        Me.CustomGroupBox2.Controls.Add(Me.txtSourceRange2)
        Me.CustomGroupBox2.Controls.Add(Me.txtSourceRange1)
        Me.CustomGroupBox2.Controls.Add(Me.lblSourceRng2)
        Me.CustomGroupBox2.Controls.Add(Me.lblSourceRng1)
        Me.CustomGroupBox2.Location = New System.Drawing.Point(15, 15)
        Me.CustomGroupBox2.Name = "CustomGroupBox2"
        Me.CustomGroupBox2.Size = New System.Drawing.Size(278, 134)
        Me.CustomGroupBox2.TabIndex = 174
        Me.CustomGroupBox2.TabStop = False
        '
        'rngSelection2
        '
        Me.rngSelection2.BackColor = System.Drawing.Color.White
        Me.rngSelection2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.rngSelection2.Image = CType(resources.GetObject("rngSelection2.Image"), System.Drawing.Image)
        Me.rngSelection2.Location = New System.Drawing.Point(244, 95)
        Me.rngSelection2.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.rngSelection2.Name = "rngSelection2"
        Me.rngSelection2.Size = New System.Drawing.Size(24, 25)
        Me.rngSelection2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.rngSelection2.TabIndex = 180
        Me.rngSelection2.TabStop = False
        '
        'rngSelection1
        '
        Me.rngSelection1.BackColor = System.Drawing.Color.White
        Me.rngSelection1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.rngSelection1.Image = CType(resources.GetObject("rngSelection1.Image"), System.Drawing.Image)
        Me.rngSelection1.Location = New System.Drawing.Point(242, 33)
        Me.rngSelection1.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.rngSelection1.Name = "rngSelection1"
        Me.rngSelection1.Size = New System.Drawing.Size(24, 25)
        Me.rngSelection1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.rngSelection1.TabIndex = 171
        Me.rngSelection1.TabStop = False
        '
        'AutoSelection2
        '
        Me.AutoSelection2.BackColor = System.Drawing.Color.White
        Me.AutoSelection2.Image = CType(resources.GetObject("AutoSelection2.Image"), System.Drawing.Image)
        Me.AutoSelection2.Location = New System.Drawing.Point(219, 96)
        Me.AutoSelection2.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.AutoSelection2.Name = "AutoSelection2"
        Me.AutoSelection2.Size = New System.Drawing.Size(24, 23)
        Me.AutoSelection2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.AutoSelection2.TabIndex = 179
        Me.AutoSelection2.TabStop = False
        '
        'AutoSelection1
        '
        Me.AutoSelection1.BackColor = System.Drawing.Color.White
        Me.AutoSelection1.Image = CType(resources.GetObject("AutoSelection1.Image"), System.Drawing.Image)
        Me.AutoSelection1.Location = New System.Drawing.Point(217, 34)
        Me.AutoSelection1.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.AutoSelection1.Name = "AutoSelection1"
        Me.AutoSelection1.Size = New System.Drawing.Size(24, 23)
        Me.AutoSelection1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.AutoSelection1.TabIndex = 170
        Me.AutoSelection1.TabStop = False
        '
        'txtSourceRange2
        '
        Me.txtSourceRange2.BackColor = System.Drawing.Color.White
        Me.txtSourceRange2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtSourceRange2.Font = New System.Drawing.Font("Segoe UI", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSourceRange2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(70, Byte), Integer), CType(CType(70, Byte), Integer), CType(CType(70, Byte), Integer))
        Me.txtSourceRange2.Location = New System.Drawing.Point(11, 95)
        Me.txtSourceRange2.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.txtSourceRange2.Name = "txtSourceRange2"
        Me.txtSourceRange2.Size = New System.Drawing.Size(256, 25)
        Me.txtSourceRange2.TabIndex = 178
        '
        'txtSourceRange1
        '
        Me.txtSourceRange1.BackColor = System.Drawing.Color.White
        Me.txtSourceRange1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtSourceRange1.Font = New System.Drawing.Font("Segoe UI", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSourceRange1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(70, Byte), Integer), CType(CType(70, Byte), Integer), CType(CType(70, Byte), Integer))
        Me.txtSourceRange1.Location = New System.Drawing.Point(9, 33)
        Me.txtSourceRange1.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.txtSourceRange1.Name = "txtSourceRange1"
        Me.txtSourceRange1.Size = New System.Drawing.Size(256, 25)
        Me.txtSourceRange1.TabIndex = 169
        '
        'lblSourceRng2
        '
        Me.lblSourceRng2.AutoSize = True
        Me.lblSourceRng2.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSourceRng2.Location = New System.Drawing.Point(8, 68)
        Me.lblSourceRng2.Name = "lblSourceRng2"
        Me.lblSourceRng2.Size = New System.Drawing.Size(256, 17)
        Me.lblSourceRng2.TabIndex = 168
        Me.lblSourceRng2.Text = "2nd Source Range (X rows x Y columns) :"
        '
        'lblSourceRng1
        '
        Me.lblSourceRng1.AutoSize = True
        Me.lblSourceRng1.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSourceRng1.Location = New System.Drawing.Point(8, 8)
        Me.lblSourceRng1.Name = "lblSourceRng1"
        Me.lblSourceRng1.Size = New System.Drawing.Size(249, 17)
        Me.lblSourceRng1.TabIndex = 164
        Me.lblSourceRng1.Text = "1st Source Range (X rows x Y columns) :"
        '
        'Form15CompareCells
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.ClientSize = New System.Drawing.Size(775, 465)
        Me.Controls.Add(Me.CustomPanel1)
        Me.Controls.Add(Me.checkBoxCase)
        Me.Controls.Add(Me.GB_Display_Result)
        Me.Controls.Add(Me.checkBoxFormatting)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.btnCanecl)
        Me.Controls.Add(Me.ComboBox1)
        Me.Controls.Add(Me.checkBoxCopyWs)
        Me.Controls.Add(Me.CustomGroupBox1)
        Me.Controls.Add(Me.CustomGroupBox2)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Form15CompareCells"
        Me.Text = "Compare Cells"
        Me.CustomPanel1.ResumeLayout(False)
        Me.CustomGroupBox5.ResumeLayout(False)
        Me.GB_Input_Range.ResumeLayout(False)
        Me.GB_Expected_Output.ResumeLayout(False)
        CType(Me.PictureBox7, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GB_Display_Result.ResumeLayout(False)
        Me.CustomGroupBox4.ResumeLayout(False)
        Me.CustomGroupBox4.PerformLayout()
        Me.CustomGroupBox1.ResumeLayout(False)
        Me.CustomGroupBox7.ResumeLayout(False)
        Me.CustomGroupBox7.PerformLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox5, System.ComponentModel.ISupportInitialize).EndInit()
        Me.CustomGroupBox2.ResumeLayout(False)
        Me.CustomGroupBox2.PerformLayout()
        CType(Me.rngSelection2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.rngSelection1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AutoSelection2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AutoSelection1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents checkBoxFormatting As Windows.Forms.CheckBox
    Friend WithEvents PictureBox7 As Windows.Forms.PictureBox
    Friend WithEvents btnOK As Windows.Forms.Button
    Friend WithEvents btnCanecl As Windows.Forms.Button
    Friend WithEvents ComboBox1 As Windows.Forms.ComboBox
    Friend WithEvents checkBoxCopyWs As Windows.Forms.CheckBox
    Friend WithEvents radBtnSameValues As Windows.Forms.RadioButton
    Friend WithEvents PictureBox1 As Windows.Forms.PictureBox
    Friend WithEvents CustomGroupBox1 As CustomGroupBox
    Friend WithEvents CustomGroupBox7 As CustomGroupBox
    Friend WithEvents PictureBox5 As Windows.Forms.PictureBox
    Friend WithEvents radBtnDifferentValues As Windows.Forms.RadioButton
    Friend WithEvents CustomGroupBox2 As CustomGroupBox
    Friend WithEvents lblSourceRng2 As Windows.Forms.Label
    Friend WithEvents lblSourceRng1 As Windows.Forms.Label
    Friend WithEvents GB_Display_Result As CustomGroupBox
    Friend WithEvents CustomGroupBox4 As CustomGroupBox
    Friend WithEvents checkBoxCase As Windows.Forms.CheckBox
    Friend WithEvents CBFillBackground As Windows.Forms.ComboBox
    Friend WithEvents checkBoxFillFont As Windows.Forms.CheckBox
    Friend WithEvents checkBoxFillBack As Windows.Forms.CheckBox
    Friend WithEvents CustomPanel1 As CustomPanel
    Friend WithEvents CustomGroupBox5 As CustomGroupBox
    Friend WithEvents CP_Input_Range2 As CustomPanel
    Friend WithEvents GB_Input_Range As CustomGroupBox
    Friend WithEvents CP_Input_Range1 As CustomPanel
    Friend WithEvents GB_Expected_Output As CustomGroupBox
    Friend WithEvents CP_Output_Range As CustomPanel
    Friend WithEvents rngSelection2 As Windows.Forms.PictureBox
    Friend WithEvents rngSelection1 As Windows.Forms.PictureBox
    Friend WithEvents AutoSelection2 As Windows.Forms.PictureBox
    Friend WithEvents AutoSelection1 As Windows.Forms.PictureBox
    Friend WithEvents txtSourceRange2 As Windows.Forms.TextBox
    Friend WithEvents txtSourceRange1 As Windows.Forms.TextBox
    Friend WithEvents CD_Fill_Background As Windows.Forms.ColorDialog
    Friend WithEvents CD_Fill_Font As Windows.Forms.ColorDialog
    Friend WithEvents CbFillFont As Windows.Forms.ComboBox
End Class
