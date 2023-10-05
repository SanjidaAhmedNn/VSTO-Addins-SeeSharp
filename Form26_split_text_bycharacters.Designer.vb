<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form26_split_text_bycharacters
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form26_split_text_bycharacters))
        Me.Label1 = New System.Windows.Forms.Label()
        Me.AutoSelection = New System.Windows.Forms.PictureBox()
        Me.Info = New System.Windows.Forms.PictureBox()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.TB_source_range = New System.Windows.Forms.TextBox()
        Me.Selection = New System.Windows.Forms.PictureBox()
        Me.CheckBox2 = New System.Windows.Forms.CheckBox()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.Btn_OK = New System.Windows.Forms.Button()
        Me.Btn_Cancel = New System.Windows.Forms.Button()
        Me.CB_formatting = New System.Windows.Forms.CheckBox()
        Me.PictureBox7 = New System.Windows.Forms.PictureBox()
        Me.CB_consecute_separators = New System.Windows.Forms.CheckBox()
        Me.CustomGroupBox2 = New VSTO_Addins.CustomGroupBox()
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        Me.PictureBox3 = New System.Windows.Forms.PictureBox()
        Me.RB_ending_point = New System.Windows.Forms.RadioButton()
        Me.RB_starting_point = New System.Windows.Forms.RadioButton()
        Me.CB_separators_finaloutput = New System.Windows.Forms.CheckBox()
        Me.CustomGroupBox4 = New VSTO_Addins.CustomGroupBox()
        Me.CustomGroupBox8 = New VSTO_Addins.CustomGroupBox()
        Me.ComboBox2 = New System.Windows.Forms.ComboBox()
        Me.PictureBox10 = New System.Windows.Forms.PictureBox()
        Me.PictureBox6 = New System.Windows.Forms.PictureBox()
        Me.PictureBox5 = New System.Windows.Forms.PictureBox()
        Me.PictureBox4 = New System.Windows.Forms.PictureBox()
        Me.PictureBox11 = New System.Windows.Forms.PictureBox()
        Me.VScrollBar1 = New System.Windows.Forms.VScrollBar()
        Me.RB_others = New System.Windows.Forms.RadioButton()
        Me.RB_semicolon = New System.Windows.Forms.RadioButton()
        Me.RB_numbertext = New System.Windows.Forms.RadioButton()
        Me.RB_newline = New System.Windows.Forms.RadioButton()
        Me.RB_space = New System.Windows.Forms.RadioButton()
        Me.TextBox3 = New System.Windows.Forms.TextBox()
        Me.RB_width = New System.Windows.Forms.RadioButton()
        Me.CustomGroupBox5 = New VSTO_Addins.CustomGroupBox()
        Me.Panel_InputRange = New VSTO_Addins.CustomPanel()
        Me.CustomGroupBox1 = New VSTO_Addins.CustomGroupBox()
        Me.CustomGroupBox7 = New VSTO_Addins.CustomGroupBox()
        Me.PictureBox8 = New System.Windows.Forms.PictureBox()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.RB_columns = New System.Windows.Forms.RadioButton()
        Me.RB_rows = New System.Windows.Forms.RadioButton()
        Me.CustomGroupBox6 = New VSTO_Addins.CustomGroupBox()
        Me.Panel_ExpectedOutput = New VSTO_Addins.CustomPanel()
        CType(Me.AutoSelection, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Info, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Selection, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox7, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.CustomGroupBox2.SuspendLayout()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.CustomGroupBox4.SuspendLayout()
        Me.CustomGroupBox8.SuspendLayout()
        CType(Me.PictureBox10, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox6, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox5, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox4, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox11, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.CustomGroupBox5.SuspendLayout()
        Me.CustomGroupBox1.SuspendLayout()
        Me.CustomGroupBox7.SuspendLayout()
        CType(Me.PictureBox8, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.CustomGroupBox6.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(15, 15)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(98, 17)
        Me.Label1.TabIndex = 290
        Me.Label1.Text = "Source Range :"
        '
        'AutoSelection
        '
        Me.AutoSelection.BackColor = System.Drawing.Color.White
        Me.AutoSelection.Image = CType(resources.GetObject("AutoSelection.Image"), System.Drawing.Image)
        Me.AutoSelection.Location = New System.Drawing.Point(228, 43)
        Me.AutoSelection.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.AutoSelection.Name = "AutoSelection"
        Me.AutoSelection.Size = New System.Drawing.Size(24, 23)
        Me.AutoSelection.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.AutoSelection.TabIndex = 301
        Me.AutoSelection.TabStop = False
        '
        'Info
        '
        Me.Info.Image = CType(resources.GetObject("Info.Image"), System.Drawing.Image)
        Me.Info.Location = New System.Drawing.Point(118, 15)
        Me.Info.Name = "Info"
        Me.Info.Size = New System.Drawing.Size(20, 20)
        Me.Info.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.Info.TabIndex = 305
        Me.Info.TabStop = False
        Me.ToolTip1.SetToolTip(Me.Info, "Please, select single column")
        '
        'TB_source_range
        '
        Me.TB_source_range.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TB_source_range.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.TB_source_range.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TB_source_range.Location = New System.Drawing.Point(15, 42)
        Me.TB_source_range.Name = "TB_source_range"
        Me.TB_source_range.Size = New System.Drawing.Size(262, 25)
        Me.TB_source_range.TabIndex = 300
        '
        'Selection
        '
        Me.Selection.BackColor = System.Drawing.Color.White
        Me.Selection.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Selection.Image = CType(resources.GetObject("Selection.Image"), System.Drawing.Image)
        Me.Selection.Location = New System.Drawing.Point(253, 42)
        Me.Selection.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Selection.Name = "Selection"
        Me.Selection.Size = New System.Drawing.Size(24, 25)
        Me.Selection.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.Selection.TabIndex = 302
        Me.Selection.TabStop = False
        '
        'CheckBox2
        '
        Me.CheckBox2.AutoSize = True
        Me.CheckBox2.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckBox2.Location = New System.Drawing.Point(15, 528)
        Me.CheckBox2.Name = "CheckBox2"
        Me.CheckBox2.Size = New System.Drawing.Size(257, 21)
        Me.CheckBox2.TabIndex = 293
        Me.CheckBox2.Text = "Create a copy of the original worksheet"
        Me.CheckBox2.UseVisualStyleBackColor = True
        '
        'ComboBox1
        '
        Me.ComboBox1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Items.AddRange(New Object() {"SOFTEKO", "About Us", "Help", "Feedback"})
        Me.ComboBox1.Location = New System.Drawing.Point(15, 560)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(154, 25)
        Me.ComboBox1.TabIndex = 294
        Me.ComboBox1.Text = "SOFTEKO"
        '
        'Btn_OK
        '
        Me.Btn_OK.BackColor = System.Drawing.Color.White
        Me.Btn_OK.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Btn_OK.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Btn_OK.Location = New System.Drawing.Point(481, 558)
        Me.Btn_OK.Name = "Btn_OK"
        Me.Btn_OK.Size = New System.Drawing.Size(62, 26)
        Me.Btn_OK.TabIndex = 298
        Me.Btn_OK.Text = "OK"
        Me.Btn_OK.UseVisualStyleBackColor = False
        '
        'Btn_Cancel
        '
        Me.Btn_Cancel.BackColor = System.Drawing.Color.White
        Me.Btn_Cancel.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Btn_Cancel.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Btn_Cancel.Location = New System.Drawing.Point(559, 558)
        Me.Btn_Cancel.Name = "Btn_Cancel"
        Me.Btn_Cancel.Size = New System.Drawing.Size(62, 26)
        Me.Btn_Cancel.TabIndex = 297
        Me.Btn_Cancel.Text = "Cancel"
        Me.Btn_Cancel.UseVisualStyleBackColor = False
        '
        'CB_formatting
        '
        Me.CB_formatting.AutoSize = True
        Me.CB_formatting.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CB_formatting.Location = New System.Drawing.Point(15, 168)
        Me.CB_formatting.Name = "CB_formatting"
        Me.CB_formatting.Size = New System.Drawing.Size(122, 21)
        Me.CB_formatting.TabIndex = 292
        Me.CB_formatting.Text = "Keep formatting"
        Me.CB_formatting.UseVisualStyleBackColor = True
        '
        'PictureBox7
        '
        Me.PictureBox7.Image = CType(resources.GetObject("PictureBox7.Image"), System.Drawing.Image)
        Me.PictureBox7.Location = New System.Drawing.Point(435, 240)
        Me.PictureBox7.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.PictureBox7.Name = "PictureBox7"
        Me.PictureBox7.Size = New System.Drawing.Size(65, 65)
        Me.PictureBox7.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox7.TabIndex = 299
        Me.PictureBox7.TabStop = False
        '
        'CB_consecute_separators
        '
        Me.CB_consecute_separators.AutoSize = True
        Me.CB_consecute_separators.Enabled = False
        Me.CB_consecute_separators.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CB_consecute_separators.Location = New System.Drawing.Point(15, 396)
        Me.CB_consecute_separators.Name = "CB_consecute_separators"
        Me.CB_consecute_separators.Size = New System.Drawing.Size(237, 21)
        Me.CB_consecute_separators.TabIndex = 306
        Me.CB_consecute_separators.Text = "Treat consecutive separators as one"
        Me.CB_consecute_separators.UseVisualStyleBackColor = True
        '
        'CustomGroupBox2
        '
        Me.CustomGroupBox2.BackColor = System.Drawing.Color.White
        Me.CustomGroupBox2.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.CustomGroupBox2.Controls.Add(Me.PictureBox2)
        Me.CustomGroupBox2.Controls.Add(Me.PictureBox3)
        Me.CustomGroupBox2.Controls.Add(Me.RB_ending_point)
        Me.CustomGroupBox2.Controls.Add(Me.RB_starting_point)
        Me.CustomGroupBox2.Controls.Add(Me.CB_separators_finaloutput)
        Me.CustomGroupBox2.Enabled = False
        Me.CustomGroupBox2.Location = New System.Drawing.Point(15, 425)
        Me.CustomGroupBox2.Name = "CustomGroupBox2"
        Me.CustomGroupBox2.Size = New System.Drawing.Size(257, 92)
        Me.CustomGroupBox2.TabIndex = 307
        Me.CustomGroupBox2.TabStop = False
        '
        'PictureBox2
        '
        Me.PictureBox2.Enabled = False
        Me.PictureBox2.Image = CType(resources.GetObject("PictureBox2.Image"), System.Drawing.Image)
        Me.PictureBox2.Location = New System.Drawing.Point(226, 34)
        Me.PictureBox2.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(20, 20)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox2.TabIndex = 308
        Me.PictureBox2.TabStop = False
        '
        'PictureBox3
        '
        Me.PictureBox3.Enabled = False
        Me.PictureBox3.Image = CType(resources.GetObject("PictureBox3.Image"), System.Drawing.Image)
        Me.PictureBox3.Location = New System.Drawing.Point(226, 60)
        Me.PictureBox3.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.PictureBox3.Name = "PictureBox3"
        Me.PictureBox3.Size = New System.Drawing.Size(20, 20)
        Me.PictureBox3.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox3.TabIndex = 309
        Me.PictureBox3.TabStop = False
        '
        'RB_ending_point
        '
        Me.RB_ending_point.AutoSize = True
        Me.RB_ending_point.Enabled = False
        Me.RB_ending_point.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RB_ending_point.Location = New System.Drawing.Point(25, 60)
        Me.RB_ending_point.Name = "RB_ending_point"
        Me.RB_ending_point.Size = New System.Drawing.Size(138, 21)
        Me.RB_ending_point.TabIndex = 310
        Me.RB_ending_point.TabStop = True
        Me.RB_ending_point.Text = "At the ending point"
        Me.RB_ending_point.UseVisualStyleBackColor = True
        '
        'RB_starting_point
        '
        Me.RB_starting_point.AutoSize = True
        Me.RB_starting_point.Enabled = False
        Me.RB_starting_point.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RB_starting_point.Location = New System.Drawing.Point(25, 34)
        Me.RB_starting_point.Name = "RB_starting_point"
        Me.RB_starting_point.Size = New System.Drawing.Size(142, 21)
        Me.RB_starting_point.TabIndex = 309
        Me.RB_starting_point.TabStop = True
        Me.RB_starting_point.Text = "At the starting point"
        Me.RB_starting_point.UseVisualStyleBackColor = True
        '
        'CB_separators_finaloutput
        '
        Me.CB_separators_finaloutput.AutoSize = True
        Me.CB_separators_finaloutput.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CB_separators_finaloutput.Location = New System.Drawing.Point(8, 9)
        Me.CB_separators_finaloutput.Name = "CB_separators_finaloutput"
        Me.CB_separators_finaloutput.Size = New System.Drawing.Size(230, 21)
        Me.CB_separators_finaloutput.TabIndex = 308
        Me.CB_separators_finaloutput.Text = "Keep separators in the final output"
        Me.CB_separators_finaloutput.UseVisualStyleBackColor = True
        '
        'CustomGroupBox4
        '
        Me.CustomGroupBox4.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.CustomGroupBox4.Controls.Add(Me.CustomGroupBox8)
        Me.CustomGroupBox4.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CustomGroupBox4.Location = New System.Drawing.Point(15, 195)
        Me.CustomGroupBox4.Name = "CustomGroupBox4"
        Me.CustomGroupBox4.Size = New System.Drawing.Size(260, 195)
        Me.CustomGroupBox4.TabIndex = 303
        Me.CustomGroupBox4.TabStop = False
        Me.CustomGroupBox4.Text = "Select Separator"
        '
        'CustomGroupBox8
        '
        Me.CustomGroupBox8.BackColor = System.Drawing.Color.White
        Me.CustomGroupBox8.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.CustomGroupBox8.Controls.Add(Me.ComboBox2)
        Me.CustomGroupBox8.Controls.Add(Me.PictureBox10)
        Me.CustomGroupBox8.Controls.Add(Me.PictureBox6)
        Me.CustomGroupBox8.Controls.Add(Me.PictureBox5)
        Me.CustomGroupBox8.Controls.Add(Me.PictureBox4)
        Me.CustomGroupBox8.Controls.Add(Me.PictureBox11)
        Me.CustomGroupBox8.Controls.Add(Me.VScrollBar1)
        Me.CustomGroupBox8.Controls.Add(Me.RB_others)
        Me.CustomGroupBox8.Controls.Add(Me.RB_semicolon)
        Me.CustomGroupBox8.Controls.Add(Me.RB_numbertext)
        Me.CustomGroupBox8.Controls.Add(Me.RB_newline)
        Me.CustomGroupBox8.Controls.Add(Me.RB_space)
        Me.CustomGroupBox8.Controls.Add(Me.TextBox3)
        Me.CustomGroupBox8.Controls.Add(Me.RB_width)
        Me.CustomGroupBox8.Location = New System.Drawing.Point(1, 22)
        Me.CustomGroupBox8.Name = "CustomGroupBox8"
        Me.CustomGroupBox8.Size = New System.Drawing.Size(259, 173)
        Me.CustomGroupBox8.TabIndex = 0
        Me.CustomGroupBox8.TabStop = False
        '
        'ComboBox2
        '
        Me.ComboBox2.FormattingEnabled = True
        Me.ComboBox2.Items.AddRange(New Object() {",", "/", "."})
        Me.ComboBox2.Location = New System.Drawing.Point(112, 104)
        Me.ComboBox2.Name = "ComboBox2"
        Me.ComboBox2.Size = New System.Drawing.Size(133, 25)
        Me.ComboBox2.TabIndex = 275
        '
        'PictureBox10
        '
        Me.PictureBox10.Image = CType(resources.GetObject("PictureBox10.Image"), System.Drawing.Image)
        Me.PictureBox10.Location = New System.Drawing.Point(225, 6)
        Me.PictureBox10.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.PictureBox10.Name = "PictureBox10"
        Me.PictureBox10.Size = New System.Drawing.Size(20, 20)
        Me.PictureBox10.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox10.TabIndex = 272
        Me.PictureBox10.TabStop = False
        '
        'PictureBox6
        '
        Me.PictureBox6.Image = CType(resources.GetObject("PictureBox6.Image"), System.Drawing.Image)
        Me.PictureBox6.Location = New System.Drawing.Point(225, 30)
        Me.PictureBox6.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.PictureBox6.Name = "PictureBox6"
        Me.PictureBox6.Size = New System.Drawing.Size(20, 20)
        Me.PictureBox6.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox6.TabIndex = 272
        Me.PictureBox6.TabStop = False
        '
        'PictureBox5
        '
        Me.PictureBox5.Image = CType(resources.GetObject("PictureBox5.Image"), System.Drawing.Image)
        Me.PictureBox5.Location = New System.Drawing.Point(225, 54)
        Me.PictureBox5.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.PictureBox5.Name = "PictureBox5"
        Me.PictureBox5.Size = New System.Drawing.Size(20, 20)
        Me.PictureBox5.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox5.TabIndex = 274
        Me.PictureBox5.TabStop = False
        '
        'PictureBox4
        '
        Me.PictureBox4.Image = CType(resources.GetObject("PictureBox4.Image"), System.Drawing.Image)
        Me.PictureBox4.Location = New System.Drawing.Point(225, 78)
        Me.PictureBox4.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.PictureBox4.Name = "PictureBox4"
        Me.PictureBox4.Size = New System.Drawing.Size(20, 20)
        Me.PictureBox4.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox4.TabIndex = 273
        Me.PictureBox4.TabStop = False
        '
        'PictureBox11
        '
        Me.PictureBox11.Image = CType(resources.GetObject("PictureBox11.Image"), System.Drawing.Image)
        Me.PictureBox11.Location = New System.Drawing.Point(225, 143)
        Me.PictureBox11.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.PictureBox11.Name = "PictureBox11"
        Me.PictureBox11.Size = New System.Drawing.Size(20, 20)
        Me.PictureBox11.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox11.TabIndex = 271
        Me.PictureBox11.TabStop = False
        '
        'VScrollBar1
        '
        Me.VScrollBar1.Location = New System.Drawing.Point(190, 141)
        Me.VScrollBar1.Name = "VScrollBar1"
        Me.VScrollBar1.Size = New System.Drawing.Size(21, 21)
        Me.VScrollBar1.TabIndex = 236
        '
        'RB_others
        '
        Me.RB_others.AutoSize = True
        Me.RB_others.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RB_others.Location = New System.Drawing.Point(8, 108)
        Me.RB_others.Name = "RB_others"
        Me.RB_others.Size = New System.Drawing.Size(72, 21)
        Me.RB_others.TabIndex = 233
        Me.RB_others.TabStop = True
        Me.RB_others.Text = "Others :"
        Me.RB_others.UseVisualStyleBackColor = True
        '
        'RB_semicolon
        '
        Me.RB_semicolon.AutoSize = True
        Me.RB_semicolon.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RB_semicolon.Location = New System.Drawing.Point(8, 78)
        Me.RB_semicolon.Name = "RB_semicolon"
        Me.RB_semicolon.Size = New System.Drawing.Size(86, 21)
        Me.RB_semicolon.TabIndex = 232
        Me.RB_semicolon.TabStop = True
        Me.RB_semicolon.Text = "Semicolon"
        Me.RB_semicolon.UseVisualStyleBackColor = True
        '
        'RB_numbertext
        '
        Me.RB_numbertext.AutoSize = True
        Me.RB_numbertext.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RB_numbertext.Location = New System.Drawing.Point(8, 54)
        Me.RB_numbertext.Name = "RB_numbertext"
        Me.RB_numbertext.Size = New System.Drawing.Size(125, 21)
        Me.RB_numbertext.TabIndex = 231
        Me.RB_numbertext.TabStop = True
        Me.RB_numbertext.Text = "Number and text"
        Me.RB_numbertext.UseVisualStyleBackColor = True
        '
        'RB_newline
        '
        Me.RB_newline.AutoSize = True
        Me.RB_newline.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RB_newline.Location = New System.Drawing.Point(8, 30)
        Me.RB_newline.Name = "RB_newline"
        Me.RB_newline.Size = New System.Drawing.Size(76, 21)
        Me.RB_newline.TabIndex = 1
        Me.RB_newline.TabStop = True
        Me.RB_newline.Text = "New line"
        Me.RB_newline.UseVisualStyleBackColor = True
        '
        'RB_space
        '
        Me.RB_space.AutoSize = True
        Me.RB_space.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RB_space.Location = New System.Drawing.Point(8, 6)
        Me.RB_space.Name = "RB_space"
        Me.RB_space.Size = New System.Drawing.Size(61, 21)
        Me.RB_space.TabIndex = 0
        Me.RB_space.TabStop = True
        Me.RB_space.Text = "Space"
        Me.RB_space.UseVisualStyleBackColor = True
        '
        'TextBox3
        '
        Me.TextBox3.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox3.Location = New System.Drawing.Point(112, 140)
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.Size = New System.Drawing.Size(100, 25)
        Me.TextBox3.TabIndex = 237
        '
        'RB_width
        '
        Me.RB_width.AutoSize = True
        Me.RB_width.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RB_width.Location = New System.Drawing.Point(8, 141)
        Me.RB_width.Name = "RB_width"
        Me.RB_width.Size = New System.Drawing.Size(105, 21)
        Me.RB_width.TabIndex = 234
        Me.RB_width.TabStop = True
        Me.RB_width.Text = "Define width :"
        Me.RB_width.UseVisualStyleBackColor = True
        '
        'CustomGroupBox5
        '
        Me.CustomGroupBox5.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.CustomGroupBox5.Controls.Add(Me.Panel_InputRange)
        Me.CustomGroupBox5.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CustomGroupBox5.Location = New System.Drawing.Point(315, 17)
        Me.CustomGroupBox5.Name = "CustomGroupBox5"
        Me.CustomGroupBox5.Size = New System.Drawing.Size(302, 200)
        Me.CustomGroupBox5.TabIndex = 295
        Me.CustomGroupBox5.TabStop = False
        Me.CustomGroupBox5.Text = "Input Range"
        '
        'Panel_InputRange
        '
        Me.Panel_InputRange.BackColor = System.Drawing.Color.White
        Me.Panel_InputRange.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.Panel_InputRange.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel_InputRange.BorderWidth = 1
        Me.Panel_InputRange.Location = New System.Drawing.Point(1, 30)
        Me.Panel_InputRange.Name = "Panel_InputRange"
        Me.Panel_InputRange.Size = New System.Drawing.Size(300, 170)
        Me.Panel_InputRange.TabIndex = 0
        '
        'CustomGroupBox1
        '
        Me.CustomGroupBox1.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.CustomGroupBox1.Controls.Add(Me.CustomGroupBox7)
        Me.CustomGroupBox1.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CustomGroupBox1.Location = New System.Drawing.Point(15, 76)
        Me.CustomGroupBox1.Name = "CustomGroupBox1"
        Me.CustomGroupBox1.Size = New System.Drawing.Size(260, 84)
        Me.CustomGroupBox1.TabIndex = 291
        Me.CustomGroupBox1.TabStop = False
        Me.CustomGroupBox1.Text = "Split Option"
        '
        'CustomGroupBox7
        '
        Me.CustomGroupBox7.BackColor = System.Drawing.Color.White
        Me.CustomGroupBox7.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.CustomGroupBox7.Controls.Add(Me.PictureBox8)
        Me.CustomGroupBox7.Controls.Add(Me.PictureBox1)
        Me.CustomGroupBox7.Controls.Add(Me.RB_columns)
        Me.CustomGroupBox7.Controls.Add(Me.RB_rows)
        Me.CustomGroupBox7.Location = New System.Drawing.Point(1, 22)
        Me.CustomGroupBox7.Name = "CustomGroupBox7"
        Me.CustomGroupBox7.Size = New System.Drawing.Size(259, 62)
        Me.CustomGroupBox7.TabIndex = 0
        Me.CustomGroupBox7.TabStop = False
        '
        'PictureBox8
        '
        Me.PictureBox8.Image = CType(resources.GetObject("PictureBox8.Image"), System.Drawing.Image)
        Me.PictureBox8.Location = New System.Drawing.Point(225, 33)
        Me.PictureBox8.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.PictureBox8.Name = "PictureBox8"
        Me.PictureBox8.Size = New System.Drawing.Size(20, 20)
        Me.PictureBox8.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox8.TabIndex = 275
        Me.PictureBox8.TabStop = False
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(225, 7)
        Me.PictureBox1.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(20, 20)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 274
        Me.PictureBox1.TabStop = False
        '
        'RB_columns
        '
        Me.RB_columns.AutoSize = True
        Me.RB_columns.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RB_columns.Location = New System.Drawing.Point(8, 32)
        Me.RB_columns.Name = "RB_columns"
        Me.RB_columns.Size = New System.Drawing.Size(167, 21)
        Me.RB_columns.TabIndex = 1
        Me.RB_columns.Text = "Split range into columns"
        Me.RB_columns.UseVisualStyleBackColor = True
        '
        'RB_rows
        '
        Me.RB_rows.AutoSize = True
        Me.RB_rows.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RB_rows.Location = New System.Drawing.Point(8, 6)
        Me.RB_rows.Name = "RB_rows"
        Me.RB_rows.Size = New System.Drawing.Size(147, 21)
        Me.RB_rows.TabIndex = 0
        Me.RB_rows.Text = "Split range into rows"
        Me.RB_rows.UseVisualStyleBackColor = True
        '
        'CustomGroupBox6
        '
        Me.CustomGroupBox6.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.CustomGroupBox6.Controls.Add(Me.Panel_ExpectedOutput)
        Me.CustomGroupBox6.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CustomGroupBox6.Location = New System.Drawing.Point(315, 322)
        Me.CustomGroupBox6.Name = "CustomGroupBox6"
        Me.CustomGroupBox6.Size = New System.Drawing.Size(302, 200)
        Me.CustomGroupBox6.TabIndex = 296
        Me.CustomGroupBox6.TabStop = False
        Me.CustomGroupBox6.Text = "Expected Output"
        '
        'Panel_ExpectedOutput
        '
        Me.Panel_ExpectedOutput.BackColor = System.Drawing.Color.White
        Me.Panel_ExpectedOutput.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.Panel_ExpectedOutput.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel_ExpectedOutput.BorderWidth = 1
        Me.Panel_ExpectedOutput.Location = New System.Drawing.Point(1, 30)
        Me.Panel_ExpectedOutput.Name = "Panel_ExpectedOutput"
        Me.Panel_ExpectedOutput.Size = New System.Drawing.Size(300, 170)
        Me.Panel_ExpectedOutput.TabIndex = 11
        '
        'Form26_split_text_bycharacters
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(643, 607)
        Me.Controls.Add(Me.CustomGroupBox2)
        Me.Controls.Add(Me.CB_consecute_separators)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.AutoSelection)
        Me.Controls.Add(Me.CustomGroupBox4)
        Me.Controls.Add(Me.Info)
        Me.Controls.Add(Me.Selection)
        Me.Controls.Add(Me.CustomGroupBox5)
        Me.Controls.Add(Me.CheckBox2)
        Me.Controls.Add(Me.CustomGroupBox1)
        Me.Controls.Add(Me.ComboBox1)
        Me.Controls.Add(Me.Btn_OK)
        Me.Controls.Add(Me.Btn_Cancel)
        Me.Controls.Add(Me.CB_formatting)
        Me.Controls.Add(Me.CustomGroupBox6)
        Me.Controls.Add(Me.PictureBox7)
        Me.Controls.Add(Me.TB_source_range)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Form26_split_text_bycharacters"
        Me.Text = "Split Text by Characters"
        CType(Me.AutoSelection, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Info, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Selection, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox7, System.ComponentModel.ISupportInitialize).EndInit()
        Me.CustomGroupBox2.ResumeLayout(False)
        Me.CustomGroupBox2.PerformLayout()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.CustomGroupBox4.ResumeLayout(False)
        Me.CustomGroupBox8.ResumeLayout(False)
        Me.CustomGroupBox8.PerformLayout()
        CType(Me.PictureBox10, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox6, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox5, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox11, System.ComponentModel.ISupportInitialize).EndInit()
        Me.CustomGroupBox5.ResumeLayout(False)
        Me.CustomGroupBox1.ResumeLayout(False)
        Me.CustomGroupBox7.ResumeLayout(False)
        Me.CustomGroupBox7.PerformLayout()
        CType(Me.PictureBox8, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.CustomGroupBox6.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents AutoSelection As Windows.Forms.PictureBox
    Friend WithEvents CustomGroupBox4 As CustomGroupBox
    Friend WithEvents CustomGroupBox8 As CustomGroupBox
    Friend WithEvents ComboBox2 As Windows.Forms.ComboBox
    Friend WithEvents PictureBox10 As Windows.Forms.PictureBox
    Friend WithEvents PictureBox6 As Windows.Forms.PictureBox
    Friend WithEvents PictureBox5 As Windows.Forms.PictureBox
    Friend WithEvents PictureBox4 As Windows.Forms.PictureBox
    Friend WithEvents PictureBox11 As Windows.Forms.PictureBox
    Friend WithEvents VScrollBar1 As Windows.Forms.VScrollBar
    Friend WithEvents RB_others As Windows.Forms.RadioButton
    Friend WithEvents RB_semicolon As Windows.Forms.RadioButton
    Friend WithEvents RB_numbertext As Windows.Forms.RadioButton
    Friend WithEvents RB_newline As Windows.Forms.RadioButton
    Friend WithEvents RB_space As Windows.Forms.RadioButton
    Friend WithEvents TextBox3 As Windows.Forms.TextBox
    Friend WithEvents RB_width As Windows.Forms.RadioButton
    Friend WithEvents Info As Windows.Forms.PictureBox
    Friend WithEvents ToolTip1 As Windows.Forms.ToolTip
    Friend WithEvents TB_source_range As Windows.Forms.TextBox
    Friend WithEvents Selection As Windows.Forms.PictureBox
    Friend WithEvents CustomGroupBox5 As CustomGroupBox
    Friend WithEvents Panel_InputRange As CustomPanel
    Friend WithEvents CheckBox2 As Windows.Forms.CheckBox
    Friend WithEvents CustomGroupBox1 As CustomGroupBox
    Friend WithEvents CustomGroupBox7 As CustomGroupBox
    Friend WithEvents PictureBox8 As Windows.Forms.PictureBox
    Friend WithEvents PictureBox1 As Windows.Forms.PictureBox
    Friend WithEvents RB_columns As Windows.Forms.RadioButton
    Friend WithEvents RB_rows As Windows.Forms.RadioButton
    Friend WithEvents ComboBox1 As Windows.Forms.ComboBox
    Friend WithEvents Btn_OK As Windows.Forms.Button
    Friend WithEvents Btn_Cancel As Windows.Forms.Button
    Friend WithEvents CB_formatting As Windows.Forms.CheckBox
    Friend WithEvents CustomGroupBox6 As CustomGroupBox
    Friend WithEvents Panel_ExpectedOutput As CustomPanel
    Friend WithEvents PictureBox7 As Windows.Forms.PictureBox
    Friend WithEvents CB_consecute_separators As Windows.Forms.CheckBox
    Friend WithEvents CustomGroupBox2 As CustomGroupBox
    Friend WithEvents RB_ending_point As Windows.Forms.RadioButton
    Friend WithEvents RB_starting_point As Windows.Forms.RadioButton
    Friend WithEvents CB_separators_finaloutput As Windows.Forms.CheckBox
    Friend WithEvents PictureBox2 As Windows.Forms.PictureBox
    Friend WithEvents PictureBox3 As Windows.Forms.PictureBox
End Class
