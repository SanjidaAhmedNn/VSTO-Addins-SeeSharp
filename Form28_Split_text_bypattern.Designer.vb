<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form28_Split_text_bypattern
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form28_Split_text_bypattern))
        Me.ComboBox2 = New System.Windows.Forms.ComboBox()
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.CustomGroupBox2 = New VSTO_Addins.CustomGroupBox()
        Me.PictureBox3 = New System.Windows.Forms.PictureBox()
        Me.PictureBox4 = New System.Windows.Forms.PictureBox()
        Me.RB_ending_point = New System.Windows.Forms.RadioButton()
        Me.RB_starting_point = New System.Windows.Forms.RadioButton()
        Me.CB_separators_finaloutput = New System.Windows.Forms.CheckBox()
        Me.CB_consecutive_separators = New System.Windows.Forms.CheckBox()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Info = New System.Windows.Forms.PictureBox()
        Me.Panel_InputRange = New VSTO_Addins.CustomPanel()
        Me.CustomGroupBox7 = New VSTO_Addins.CustomGroupBox()
        Me.PictureBox8 = New System.Windows.Forms.PictureBox()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.RB_columns = New System.Windows.Forms.RadioButton()
        Me.RB_rows = New System.Windows.Forms.RadioButton()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.AutoSelection = New System.Windows.Forms.PictureBox()
        Me.Selection = New System.Windows.Forms.PictureBox()
        Me.PictureBox7 = New System.Windows.Forms.PictureBox()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.Btn_OK = New System.Windows.Forms.Button()
        Me.Panel_ExpectedOutput = New VSTO_Addins.CustomPanel()
        Me.Btn_Cancel = New System.Windows.Forms.Button()
        Me.CB_formatting = New System.Windows.Forms.CheckBox()
        Me.CustomGroupBox6 = New VSTO_Addins.CustomGroupBox()
        Me.CB_backup = New System.Windows.Forms.CheckBox()
        Me.CustomGroupBox1 = New VSTO_Addins.CustomGroupBox()
        Me.CustomGroupBox5 = New VSTO_Addins.CustomGroupBox()
        Me.TB_source_range = New System.Windows.Forms.TextBox()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.CustomGroupBox2.SuspendLayout()
        CType(Me.PictureBox3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox4, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Info, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.CustomGroupBox7.SuspendLayout()
        CType(Me.PictureBox8, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AutoSelection, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Selection, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox7, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.CustomGroupBox6.SuspendLayout()
        Me.CustomGroupBox1.SuspendLayout()
        Me.CustomGroupBox5.SuspendLayout()
        Me.SuspendLayout()
        '
        'ComboBox2
        '
        Me.ComboBox2.Font = New System.Drawing.Font("Segoe UI", 11.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBox2.FormattingEnabled = True
        Me.ComboBox2.Location = New System.Drawing.Point(15, 242)
        Me.ComboBox2.Name = "ComboBox2"
        Me.ComboBox2.Size = New System.Drawing.Size(254, 28)
        Me.ComboBox2.TabIndex = 346
        '
        'PictureBox2
        '
        Me.PictureBox2.Image = CType(resources.GetObject("PictureBox2.Image"), System.Drawing.Image)
        Me.PictureBox2.Location = New System.Drawing.Point(112, 209)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(20, 20)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox2.TabIndex = 345
        Me.PictureBox2.TabStop = False
        Me.ToolTip1.SetToolTip(Me.PictureBox2, "Please, select single column")
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(15, 210)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(96, 17)
        Me.Label2.TabIndex = 344
        Me.Label2.Text = "Enter Pattern :"
        '
        'CustomGroupBox2
        '
        Me.CustomGroupBox2.BackColor = System.Drawing.Color.White
        Me.CustomGroupBox2.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.CustomGroupBox2.Controls.Add(Me.PictureBox3)
        Me.CustomGroupBox2.Controls.Add(Me.PictureBox4)
        Me.CustomGroupBox2.Controls.Add(Me.RB_ending_point)
        Me.CustomGroupBox2.Controls.Add(Me.RB_starting_point)
        Me.CustomGroupBox2.Controls.Add(Me.CB_separators_finaloutput)
        Me.CustomGroupBox2.Location = New System.Drawing.Point(15, 315)
        Me.CustomGroupBox2.Name = "CustomGroupBox2"
        Me.CustomGroupBox2.Size = New System.Drawing.Size(257, 92)
        Me.CustomGroupBox2.TabIndex = 343
        Me.CustomGroupBox2.TabStop = False
        '
        'PictureBox3
        '
        Me.PictureBox3.Image = CType(resources.GetObject("PictureBox3.Image"), System.Drawing.Image)
        Me.PictureBox3.Location = New System.Drawing.Point(226, 36)
        Me.PictureBox3.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.PictureBox3.Name = "PictureBox3"
        Me.PictureBox3.Size = New System.Drawing.Size(20, 20)
        Me.PictureBox3.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox3.TabIndex = 347
        Me.PictureBox3.TabStop = False
        '
        'PictureBox4
        '
        Me.PictureBox4.Image = CType(resources.GetObject("PictureBox4.Image"), System.Drawing.Image)
        Me.PictureBox4.Location = New System.Drawing.Point(226, 61)
        Me.PictureBox4.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.PictureBox4.Name = "PictureBox4"
        Me.PictureBox4.Size = New System.Drawing.Size(20, 20)
        Me.PictureBox4.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox4.TabIndex = 348
        Me.PictureBox4.TabStop = False
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
        'CB_consecutive_separators
        '
        Me.CB_consecutive_separators.AutoSize = True
        Me.CB_consecutive_separators.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CB_consecutive_separators.Location = New System.Drawing.Point(15, 284)
        Me.CB_consecutive_separators.Name = "CB_consecutive_separators"
        Me.CB_consecutive_separators.Size = New System.Drawing.Size(237, 21)
        Me.CB_consecutive_separators.TabIndex = 342
        Me.CB_consecutive_separators.Text = "Treat consecutive separators as one"
        Me.CB_consecutive_separators.UseVisualStyleBackColor = True
        '
        'Info
        '
        Me.Info.Image = CType(resources.GetObject("Info.Image"), System.Drawing.Image)
        Me.Info.Location = New System.Drawing.Point(115, 15)
        Me.Info.Name = "Info"
        Me.Info.Size = New System.Drawing.Size(20, 20)
        Me.Info.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.Info.TabIndex = 341
        Me.Info.TabStop = False
        Me.ToolTip1.SetToolTip(Me.Info, "Please, select single column")
        '
        'Panel_InputRange
        '
        Me.Panel_InputRange.BackColor = System.Drawing.Color.White
        Me.Panel_InputRange.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.Panel_InputRange.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel_InputRange.BorderWidth = 1
        Me.Panel_InputRange.Location = New System.Drawing.Point(1, 30)
        Me.Panel_InputRange.Name = "Panel_InputRange"
        Me.Panel_InputRange.Size = New System.Drawing.Size(280, 150)
        Me.Panel_InputRange.TabIndex = 0
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
        Me.RB_columns.Size = New System.Drawing.Size(154, 21)
        Me.RB_columns.TabIndex = 1
        Me.RB_columns.Text = "Split text into columns"
        Me.RB_columns.UseVisualStyleBackColor = True
        '
        'RB_rows
        '
        Me.RB_rows.AutoSize = True
        Me.RB_rows.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RB_rows.Location = New System.Drawing.Point(8, 6)
        Me.RB_rows.Name = "RB_rows"
        Me.RB_rows.Size = New System.Drawing.Size(134, 21)
        Me.RB_rows.TabIndex = 0
        Me.RB_rows.Text = "Split text into rows"
        Me.RB_rows.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(15, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(98, 17)
        Me.Label1.TabIndex = 328
        Me.Label1.Text = "Source Range :"
        '
        'AutoSelection
        '
        Me.AutoSelection.BackColor = System.Drawing.Color.White
        Me.AutoSelection.Image = CType(resources.GetObject("AutoSelection.Image"), System.Drawing.Image)
        Me.AutoSelection.Location = New System.Drawing.Point(226, 45)
        Me.AutoSelection.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.AutoSelection.Name = "AutoSelection"
        Me.AutoSelection.Size = New System.Drawing.Size(24, 23)
        Me.AutoSelection.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.AutoSelection.TabIndex = 339
        Me.AutoSelection.TabStop = False
        '
        'Selection
        '
        Me.Selection.BackColor = System.Drawing.Color.White
        Me.Selection.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Selection.Image = CType(resources.GetObject("Selection.Image"), System.Drawing.Image)
        Me.Selection.Location = New System.Drawing.Point(252, 44)
        Me.Selection.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Selection.Name = "Selection"
        Me.Selection.Size = New System.Drawing.Size(24, 25)
        Me.Selection.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.Selection.TabIndex = 340
        Me.Selection.TabStop = False
        '
        'PictureBox7
        '
        Me.PictureBox7.Image = CType(resources.GetObject("PictureBox7.Image"), System.Drawing.Image)
        Me.PictureBox7.Location = New System.Drawing.Point(425, 204)
        Me.PictureBox7.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.PictureBox7.Name = "PictureBox7"
        Me.PictureBox7.Size = New System.Drawing.Size(52, 52)
        Me.PictureBox7.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox7.TabIndex = 337
        Me.PictureBox7.TabStop = False
        '
        'ComboBox1
        '
        Me.ComboBox1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Location = New System.Drawing.Point(16, 457)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(154, 25)
        Me.ComboBox1.TabIndex = 332
        Me.ComboBox1.Text = "SOFTEKO"
        '
        'Btn_OK
        '
        Me.Btn_OK.BackColor = System.Drawing.Color.White
        Me.Btn_OK.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Btn_OK.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Btn_OK.Location = New System.Drawing.Point(449, 457)
        Me.Btn_OK.Name = "Btn_OK"
        Me.Btn_OK.Size = New System.Drawing.Size(62, 26)
        Me.Btn_OK.TabIndex = 336
        Me.Btn_OK.Text = "OK"
        Me.Btn_OK.UseVisualStyleBackColor = False
        '
        'Panel_ExpectedOutput
        '
        Me.Panel_ExpectedOutput.BackColor = System.Drawing.Color.White
        Me.Panel_ExpectedOutput.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.Panel_ExpectedOutput.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel_ExpectedOutput.BorderWidth = 1
        Me.Panel_ExpectedOutput.Location = New System.Drawing.Point(1, 30)
        Me.Panel_ExpectedOutput.Name = "Panel_ExpectedOutput"
        Me.Panel_ExpectedOutput.Size = New System.Drawing.Size(280, 150)
        Me.Panel_ExpectedOutput.TabIndex = 11
        '
        'Btn_Cancel
        '
        Me.Btn_Cancel.BackColor = System.Drawing.Color.White
        Me.Btn_Cancel.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Btn_Cancel.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Btn_Cancel.Location = New System.Drawing.Point(530, 455)
        Me.Btn_Cancel.Name = "Btn_Cancel"
        Me.Btn_Cancel.Size = New System.Drawing.Size(62, 26)
        Me.Btn_Cancel.TabIndex = 335
        Me.Btn_Cancel.Text = "Cancel"
        Me.Btn_Cancel.UseVisualStyleBackColor = False
        '
        'CB_formatting
        '
        Me.CB_formatting.AutoSize = True
        Me.CB_formatting.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CB_formatting.Location = New System.Drawing.Point(15, 177)
        Me.CB_formatting.Name = "CB_formatting"
        Me.CB_formatting.Size = New System.Drawing.Size(122, 21)
        Me.CB_formatting.TabIndex = 330
        Me.CB_formatting.Text = "Keep formatting"
        Me.CB_formatting.UseVisualStyleBackColor = True
        '
        'CustomGroupBox6
        '
        Me.CustomGroupBox6.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.CustomGroupBox6.Controls.Add(Me.Panel_ExpectedOutput)
        Me.CustomGroupBox6.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CustomGroupBox6.Location = New System.Drawing.Point(310, 252)
        Me.CustomGroupBox6.Name = "CustomGroupBox6"
        Me.CustomGroupBox6.Size = New System.Drawing.Size(282, 180)
        Me.CustomGroupBox6.TabIndex = 334
        Me.CustomGroupBox6.TabStop = False
        Me.CustomGroupBox6.Text = "Expected Output"
        '
        'CB_backup
        '
        Me.CB_backup.AutoSize = True
        Me.CB_backup.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CB_backup.Location = New System.Drawing.Point(15, 420)
        Me.CB_backup.Name = "CB_backup"
        Me.CB_backup.Size = New System.Drawing.Size(257, 21)
        Me.CB_backup.TabIndex = 331
        Me.CB_backup.Text = "Create a copy of the original worksheet"
        Me.CB_backup.UseVisualStyleBackColor = True
        '
        'CustomGroupBox1
        '
        Me.CustomGroupBox1.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.CustomGroupBox1.Controls.Add(Me.CustomGroupBox7)
        Me.CustomGroupBox1.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CustomGroupBox1.Location = New System.Drawing.Point(15, 81)
        Me.CustomGroupBox1.Name = "CustomGroupBox1"
        Me.CustomGroupBox1.Size = New System.Drawing.Size(260, 84)
        Me.CustomGroupBox1.TabIndex = 329
        Me.CustomGroupBox1.TabStop = False
        Me.CustomGroupBox1.Text = "Split Option"
        '
        'CustomGroupBox5
        '
        Me.CustomGroupBox5.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.CustomGroupBox5.Controls.Add(Me.Panel_InputRange)
        Me.CustomGroupBox5.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CustomGroupBox5.Location = New System.Drawing.Point(310, 15)
        Me.CustomGroupBox5.Name = "CustomGroupBox5"
        Me.CustomGroupBox5.Size = New System.Drawing.Size(282, 180)
        Me.CustomGroupBox5.TabIndex = 333
        Me.CustomGroupBox5.TabStop = False
        Me.CustomGroupBox5.Text = "Input Range"
        '
        'TB_source_range
        '
        Me.TB_source_range.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TB_source_range.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.TB_source_range.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TB_source_range.Location = New System.Drawing.Point(15, 44)
        Me.TB_source_range.Name = "TB_source_range"
        Me.TB_source_range.Size = New System.Drawing.Size(262, 25)
        Me.TB_source_range.TabIndex = 338
        '
        'Form28_Split_text_bypattern
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.ClientSize = New System.Drawing.Size(616, 503)
        Me.Controls.Add(Me.ComboBox2)
        Me.Controls.Add(Me.PictureBox2)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.CustomGroupBox2)
        Me.Controls.Add(Me.CB_consecutive_separators)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.AutoSelection)
        Me.Controls.Add(Me.Selection)
        Me.Controls.Add(Me.PictureBox7)
        Me.Controls.Add(Me.ComboBox1)
        Me.Controls.Add(Me.Btn_OK)
        Me.Controls.Add(Me.Info)
        Me.Controls.Add(Me.Btn_Cancel)
        Me.Controls.Add(Me.CB_formatting)
        Me.Controls.Add(Me.CustomGroupBox6)
        Me.Controls.Add(Me.CB_backup)
        Me.Controls.Add(Me.CustomGroupBox1)
        Me.Controls.Add(Me.CustomGroupBox5)
        Me.Controls.Add(Me.TB_source_range)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Form28_Split_text_bypattern"
        Me.Text = "Split Text by Pattern"
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.CustomGroupBox2.ResumeLayout(False)
        Me.CustomGroupBox2.PerformLayout()
        CType(Me.PictureBox3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Info, System.ComponentModel.ISupportInitialize).EndInit()
        Me.CustomGroupBox7.ResumeLayout(False)
        Me.CustomGroupBox7.PerformLayout()
        CType(Me.PictureBox8, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AutoSelection, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Selection, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox7, System.ComponentModel.ISupportInitialize).EndInit()
        Me.CustomGroupBox6.ResumeLayout(False)
        Me.CustomGroupBox1.ResumeLayout(False)
        Me.CustomGroupBox5.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents ComboBox2 As Windows.Forms.ComboBox
    Friend WithEvents PictureBox2 As Windows.Forms.PictureBox
    Friend WithEvents ToolTip1 As Windows.Forms.ToolTip
    Friend WithEvents Label2 As Windows.Forms.Label
    Friend WithEvents CustomGroupBox2 As CustomGroupBox
    Friend WithEvents RB_ending_point As Windows.Forms.RadioButton
    Friend WithEvents RB_starting_point As Windows.Forms.RadioButton
    Friend WithEvents CB_separators_finaloutput As Windows.Forms.CheckBox
    Friend WithEvents CB_consecutive_separators As Windows.Forms.CheckBox
    Friend WithEvents Info As Windows.Forms.PictureBox
    Friend WithEvents Panel_InputRange As CustomPanel
    Friend WithEvents CustomGroupBox7 As CustomGroupBox
    Friend WithEvents PictureBox8 As Windows.Forms.PictureBox
    Friend WithEvents PictureBox1 As Windows.Forms.PictureBox
    Friend WithEvents RB_columns As Windows.Forms.RadioButton
    Friend WithEvents RB_rows As Windows.Forms.RadioButton
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents AutoSelection As Windows.Forms.PictureBox
    Friend WithEvents Selection As Windows.Forms.PictureBox
    Friend WithEvents PictureBox7 As Windows.Forms.PictureBox
    Friend WithEvents ComboBox1 As Windows.Forms.ComboBox
    Friend WithEvents Btn_OK As Windows.Forms.Button
    Friend WithEvents Panel_ExpectedOutput As CustomPanel
    Friend WithEvents Btn_Cancel As Windows.Forms.Button
    Friend WithEvents CB_formatting As Windows.Forms.CheckBox
    Friend WithEvents CustomGroupBox6 As CustomGroupBox
    Friend WithEvents CB_backup As Windows.Forms.CheckBox
    Friend WithEvents CustomGroupBox1 As CustomGroupBox
    Friend WithEvents CustomGroupBox5 As CustomGroupBox
    Friend WithEvents TB_source_range As Windows.Forms.TextBox
    Friend WithEvents PictureBox3 As Windows.Forms.PictureBox
    Friend WithEvents PictureBox4 As Windows.Forms.PictureBox
End Class
