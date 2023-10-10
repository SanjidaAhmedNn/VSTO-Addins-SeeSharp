<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form30_Create_Dynamic_Drop_down_List
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form30_Create_Dynamic_Drop_down_List))
        Me.Btn_OK = New System.Windows.Forms.Button()
        Me.Btn_Cancel = New System.Windows.Forms.Button()
        Me.ComboBox2 = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Selection_source = New System.Windows.Forms.PictureBox()
        Me.TB_src_range = New System.Windows.Forms.TextBox()
        Me.CB_header = New System.Windows.Forms.CheckBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Selection_destination = New System.Windows.Forms.PictureBox()
        Me.TB_dest_range = New System.Windows.Forms.TextBox()
        Me.CB_ascending = New System.Windows.Forms.CheckBox()
        Me.CB_descending = New System.Windows.Forms.CheckBox()
        Me.CB_text = New System.Windows.Forms.CheckBox()
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        Me.PictureBox3 = New System.Windows.Forms.PictureBox()
        Me.RB_vertical = New System.Windows.Forms.RadioButton()
        Me.RB_Horizontal = New System.Windows.Forms.RadioButton()
        Me.PictureBox8 = New System.Windows.Forms.PictureBox()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.RB_2_5_levels = New System.Windows.Forms.RadioButton()
        Me.RB_2_levels = New System.Windows.Forms.RadioButton()
        Me.CustomGroupBox1 = New VSTO_Addins.CustomGroupBox()
        Me.CustomGroupBox3 = New VSTO_Addins.CustomGroupBox()
        Me.PictureBox4 = New System.Windows.Forms.PictureBox()
        Me.RB_Dropdown_35_Labels = New System.Windows.Forms.RadioButton()
        Me.PictureBox5 = New System.Windows.Forms.PictureBox()
        Me.RB_Dropdown_2_Labels = New System.Windows.Forms.RadioButton()
        Me.CustomGroupBox2 = New VSTO_Addins.CustomGroupBox()
        Me.GB_list_option = New VSTO_Addins.CustomGroupBox()
        Me.CustomGroupBox5 = New VSTO_Addins.CustomGroupBox()
        Me.PictureBox6 = New System.Windows.Forms.PictureBox()
        Me.RB_Verti = New System.Windows.Forms.RadioButton()
        Me.PictureBox7 = New System.Windows.Forms.PictureBox()
        Me.RB_Horizon = New System.Windows.Forms.RadioButton()
        Me.CustomGroupBox4 = New VSTO_Addins.CustomGroupBox()
        Me.CustomGroupBox7 = New VSTO_Addins.CustomGroupBox()
        CType(Me.Selection_source, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Selection_destination, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox8, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.CustomGroupBox1.SuspendLayout()
        Me.CustomGroupBox3.SuspendLayout()
        CType(Me.PictureBox4, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox5, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GB_list_option.SuspendLayout()
        Me.CustomGroupBox5.SuspendLayout()
        CType(Me.PictureBox6, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox7, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Btn_OK
        '
        Me.Btn_OK.BackColor = System.Drawing.Color.White
        Me.Btn_OK.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Btn_OK.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Btn_OK.Location = New System.Drawing.Point(519, 436)
        Me.Btn_OK.Name = "Btn_OK"
        Me.Btn_OK.Size = New System.Drawing.Size(62, 26)
        Me.Btn_OK.TabIndex = 366
        Me.Btn_OK.Text = "OK"
        Me.Btn_OK.UseVisualStyleBackColor = False
        '
        'Btn_Cancel
        '
        Me.Btn_Cancel.BackColor = System.Drawing.Color.White
        Me.Btn_Cancel.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Btn_Cancel.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Btn_Cancel.Location = New System.Drawing.Point(597, 436)
        Me.Btn_Cancel.Name = "Btn_Cancel"
        Me.Btn_Cancel.Size = New System.Drawing.Size(62, 26)
        Me.Btn_Cancel.TabIndex = 365
        Me.Btn_Cancel.Text = "Cancel"
        Me.Btn_Cancel.UseVisualStyleBackColor = False
        '
        'ComboBox2
        '
        Me.ComboBox2.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBox2.FormattingEnabled = True
        Me.ComboBox2.Location = New System.Drawing.Point(15, 436)
        Me.ComboBox2.Name = "ComboBox2"
        Me.ComboBox2.Size = New System.Drawing.Size(154, 25)
        Me.ComboBox2.TabIndex = 362
        Me.ComboBox2.Text = "SOFTEKO"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(15, 15)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(98, 17)
        Me.Label1.TabIndex = 355
        Me.Label1.Text = "Source Range :"
        '
        'Selection_source
        '
        Me.Selection_source.BackColor = System.Drawing.Color.White
        Me.Selection_source.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Selection_source.Image = CType(resources.GetObject("Selection_source.Image"), System.Drawing.Image)
        Me.Selection_source.Location = New System.Drawing.Point(298, 40)
        Me.Selection_source.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Selection_source.Name = "Selection_source"
        Me.Selection_source.Size = New System.Drawing.Size(24, 25)
        Me.Selection_source.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.Selection_source.TabIndex = 357
        Me.Selection_source.TabStop = False
        '
        'TB_src_range
        '
        Me.TB_src_range.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TB_src_range.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.TB_src_range.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TB_src_range.Location = New System.Drawing.Point(15, 40)
        Me.TB_src_range.Name = "TB_src_range"
        Me.TB_src_range.Size = New System.Drawing.Size(307, 25)
        Me.TB_src_range.TabIndex = 356
        '
        'CB_header
        '
        Me.CB_header.AutoSize = True
        Me.CB_header.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CB_header.Location = New System.Drawing.Point(15, 74)
        Me.CB_header.Name = "CB_header"
        Me.CB_header.Size = New System.Drawing.Size(180, 21)
        Me.CB_header.TabIndex = 367
        Me.CB_header.Text = "My range contains header"
        Me.CB_header.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(15, 286)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(126, 17)
        Me.Label2.TabIndex = 370
        Me.Label2.Text = "Destination Range :"
        '
        'Selection_destination
        '
        Me.Selection_destination.BackColor = System.Drawing.Color.White
        Me.Selection_destination.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Selection_destination.Image = CType(resources.GetObject("Selection_destination.Image"), System.Drawing.Image)
        Me.Selection_destination.Location = New System.Drawing.Point(298, 312)
        Me.Selection_destination.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Selection_destination.Name = "Selection_destination"
        Me.Selection_destination.Size = New System.Drawing.Size(24, 25)
        Me.Selection_destination.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.Selection_destination.TabIndex = 372
        Me.Selection_destination.TabStop = False
        '
        'TB_dest_range
        '
        Me.TB_dest_range.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TB_dest_range.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.TB_dest_range.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TB_dest_range.Location = New System.Drawing.Point(15, 312)
        Me.TB_dest_range.Name = "TB_dest_range"
        Me.TB_dest_range.Size = New System.Drawing.Size(307, 25)
        Me.TB_dest_range.TabIndex = 371
        '
        'CB_ascending
        '
        Me.CB_ascending.AutoSize = True
        Me.CB_ascending.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CB_ascending.Location = New System.Drawing.Point(15, 345)
        Me.CB_ascending.Name = "CB_ascending"
        Me.CB_ascending.Size = New System.Drawing.Size(165, 21)
        Me.CB_ascending.TabIndex = 373
        Me.CB_ascending.Text = "Sort in ascending order"
        Me.CB_ascending.UseVisualStyleBackColor = True
        '
        'CB_descending
        '
        Me.CB_descending.AutoSize = True
        Me.CB_descending.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CB_descending.Location = New System.Drawing.Point(15, 372)
        Me.CB_descending.Name = "CB_descending"
        Me.CB_descending.Size = New System.Drawing.Size(173, 21)
        Me.CB_descending.TabIndex = 374
        Me.CB_descending.Text = "Sort in descending order"
        Me.CB_descending.UseVisualStyleBackColor = True
        '
        'CB_text
        '
        Me.CB_text.AutoSize = True
        Me.CB_text.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CB_text.Location = New System.Drawing.Point(15, 400)
        Me.CB_text.Name = "CB_text"
        Me.CB_text.Size = New System.Drawing.Size(147, 21)
        Me.CB_text.TabIndex = 375
        Me.CB_text.Text = "Store all data as text"
        Me.CB_text.UseVisualStyleBackColor = True
        '
        'PictureBox2
        '
        Me.PictureBox2.Image = CType(resources.GetObject("PictureBox2.Image"), System.Drawing.Image)
        Me.PictureBox2.Location = New System.Drawing.Point(277, 33)
        Me.PictureBox2.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(20, 20)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox2.TabIndex = 275
        Me.PictureBox2.TabStop = False
        '
        'PictureBox3
        '
        Me.PictureBox3.Image = CType(resources.GetObject("PictureBox3.Image"), System.Drawing.Image)
        Me.PictureBox3.Location = New System.Drawing.Point(277, 7)
        Me.PictureBox3.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.PictureBox3.Name = "PictureBox3"
        Me.PictureBox3.Size = New System.Drawing.Size(20, 20)
        Me.PictureBox3.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox3.TabIndex = 274
        Me.PictureBox3.TabStop = False
        '
        'RB_vertical
        '
        Me.RB_vertical.AutoSize = True
        Me.RB_vertical.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RB_vertical.Location = New System.Drawing.Point(8, 32)
        Me.RB_vertical.Name = "RB_vertical"
        Me.RB_vertical.Size = New System.Drawing.Size(158, 21)
        Me.RB_vertical.TabIndex = 1
        Me.RB_vertical.Text = "Vertical drop-down list"
        Me.RB_vertical.UseVisualStyleBackColor = True
        '
        'RB_Horizontal
        '
        Me.RB_Horizontal.AutoSize = True
        Me.RB_Horizontal.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RB_Horizontal.Location = New System.Drawing.Point(8, 6)
        Me.RB_Horizontal.Name = "RB_Horizontal"
        Me.RB_Horizontal.Size = New System.Drawing.Size(176, 21)
        Me.RB_Horizontal.TabIndex = 0
        Me.RB_Horizontal.Text = "Horizontal drop-down list"
        Me.RB_Horizontal.UseVisualStyleBackColor = True
        Me.RB_Horizontal.Visible = False
        '
        'PictureBox8
        '
        Me.PictureBox8.Image = CType(resources.GetObject("PictureBox8.Image"), System.Drawing.Image)
        Me.PictureBox8.Location = New System.Drawing.Point(277, 33)
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
        Me.PictureBox1.Location = New System.Drawing.Point(277, 7)
        Me.PictureBox1.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(20, 20)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 274
        Me.PictureBox1.TabStop = False
        '
        'RB_2_5_levels
        '
        Me.RB_2_5_levels.AutoSize = True
        Me.RB_2_5_levels.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RB_2_5_levels.Location = New System.Drawing.Point(8, 32)
        Me.RB_2_5_levels.Name = "RB_2_5_levels"
        Me.RB_2_5_levels.Size = New System.Drawing.Size(266, 21)
        Me.RB_2_5_levels.TabIndex = 1
        Me.RB_2_5_levels.Text = "Dynamic drop-down list with 2 to 5 levels"
        Me.RB_2_5_levels.UseVisualStyleBackColor = True
        '
        'RB_2_levels
        '
        Me.RB_2_levels.AutoSize = True
        Me.RB_2_levels.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RB_2_levels.Location = New System.Drawing.Point(8, 6)
        Me.RB_2_levels.Name = "RB_2_levels"
        Me.RB_2_levels.Size = New System.Drawing.Size(239, 21)
        Me.RB_2_levels.TabIndex = 0
        Me.RB_2_levels.Text = "Dynamic drop-down list with 2 levels"
        Me.RB_2_levels.UseVisualStyleBackColor = True
        '
        'CustomGroupBox1
        '
        Me.CustomGroupBox1.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.CustomGroupBox1.Controls.Add(Me.CustomGroupBox3)
        Me.CustomGroupBox1.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CustomGroupBox1.Location = New System.Drawing.Point(15, 102)
        Me.CustomGroupBox1.Name = "CustomGroupBox1"
        Me.CustomGroupBox1.Size = New System.Drawing.Size(307, 84)
        Me.CustomGroupBox1.TabIndex = 368
        Me.CustomGroupBox1.TabStop = False
        Me.CustomGroupBox1.Text = "List Type"
        '
        'CustomGroupBox3
        '
        Me.CustomGroupBox3.BackColor = System.Drawing.Color.White
        Me.CustomGroupBox3.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.CustomGroupBox3.Controls.Add(Me.PictureBox4)
        Me.CustomGroupBox3.Controls.Add(Me.RB_Dropdown_35_Labels)
        Me.CustomGroupBox3.Controls.Add(Me.PictureBox5)
        Me.CustomGroupBox3.Controls.Add(Me.RB_Dropdown_2_Labels)
        Me.CustomGroupBox3.Location = New System.Drawing.Point(1, 22)
        Me.CustomGroupBox3.Name = "CustomGroupBox3"
        Me.CustomGroupBox3.Size = New System.Drawing.Size(305, 62)
        Me.CustomGroupBox3.TabIndex = 0
        Me.CustomGroupBox3.TabStop = False
        '
        'PictureBox4
        '
        Me.PictureBox4.Image = CType(resources.GetObject("PictureBox4.Image"), System.Drawing.Image)
        Me.PictureBox4.Location = New System.Drawing.Point(277, 34)
        Me.PictureBox4.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.PictureBox4.Name = "PictureBox4"
        Me.PictureBox4.Size = New System.Drawing.Size(20, 20)
        Me.PictureBox4.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox4.TabIndex = 377
        Me.PictureBox4.TabStop = False
        '
        'RB_Dropdown_35_Labels
        '
        Me.RB_Dropdown_35_Labels.AutoSize = True
        Me.RB_Dropdown_35_Labels.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RB_Dropdown_35_Labels.Location = New System.Drawing.Point(8, 33)
        Me.RB_Dropdown_35_Labels.Name = "RB_Dropdown_35_Labels"
        Me.RB_Dropdown_35_Labels.Size = New System.Drawing.Size(258, 21)
        Me.RB_Dropdown_35_Labels.TabIndex = 1
        Me.RB_Dropdown_35_Labels.Text = "Dynamic Drop-down List with 3-5 Levels"
        Me.RB_Dropdown_35_Labels.UseVisualStyleBackColor = True
        '
        'PictureBox5
        '
        Me.PictureBox5.Image = CType(resources.GetObject("PictureBox5.Image"), System.Drawing.Image)
        Me.PictureBox5.Location = New System.Drawing.Point(277, 8)
        Me.PictureBox5.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.PictureBox5.Name = "PictureBox5"
        Me.PictureBox5.Size = New System.Drawing.Size(20, 20)
        Me.PictureBox5.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox5.TabIndex = 376
        Me.PictureBox5.TabStop = False
        '
        'RB_Dropdown_2_Labels
        '
        Me.RB_Dropdown_2_Labels.AutoSize = True
        Me.RB_Dropdown_2_Labels.Checked = True
        Me.RB_Dropdown_2_Labels.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RB_Dropdown_2_Labels.Location = New System.Drawing.Point(8, 7)
        Me.RB_Dropdown_2_Labels.Name = "RB_Dropdown_2_Labels"
        Me.RB_Dropdown_2_Labels.Size = New System.Drawing.Size(246, 21)
        Me.RB_Dropdown_2_Labels.TabIndex = 0
        Me.RB_Dropdown_2_Labels.TabStop = True
        Me.RB_Dropdown_2_Labels.Text = "Dynamic Drop-down List with 2 Levels"
        Me.RB_Dropdown_2_Labels.UseVisualStyleBackColor = True
        '
        'CustomGroupBox2
        '
        Me.CustomGroupBox2.BackColor = System.Drawing.Color.White
        Me.CustomGroupBox2.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.CustomGroupBox2.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CustomGroupBox2.Location = New System.Drawing.Point(352, 15)
        Me.CustomGroupBox2.Name = "CustomGroupBox2"
        Me.CustomGroupBox2.Size = New System.Drawing.Size(307, 395)
        Me.CustomGroupBox2.TabIndex = 363
        Me.CustomGroupBox2.TabStop = False
        Me.CustomGroupBox2.Text = "Sample Image"
        '
        'GB_list_option
        '
        Me.GB_list_option.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.GB_list_option.Controls.Add(Me.CustomGroupBox5)
        Me.GB_list_option.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GB_list_option.Location = New System.Drawing.Point(15, 195)
        Me.GB_list_option.Name = "GB_list_option"
        Me.GB_list_option.Size = New System.Drawing.Size(307, 84)
        Me.GB_list_option.TabIndex = 369
        Me.GB_list_option.TabStop = False
        Me.GB_list_option.Text = "List Option"
        '
        'CustomGroupBox5
        '
        Me.CustomGroupBox5.BackColor = System.Drawing.Color.White
        Me.CustomGroupBox5.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.CustomGroupBox5.Controls.Add(Me.PictureBox6)
        Me.CustomGroupBox5.Controls.Add(Me.RB_Verti)
        Me.CustomGroupBox5.Controls.Add(Me.PictureBox7)
        Me.CustomGroupBox5.Controls.Add(Me.RB_Horizon)
        Me.CustomGroupBox5.Location = New System.Drawing.Point(1, 22)
        Me.CustomGroupBox5.Name = "CustomGroupBox5"
        Me.CustomGroupBox5.Size = New System.Drawing.Size(305, 62)
        Me.CustomGroupBox5.TabIndex = 1
        Me.CustomGroupBox5.TabStop = False
        '
        'PictureBox6
        '
        Me.PictureBox6.Image = CType(resources.GetObject("PictureBox6.Image"), System.Drawing.Image)
        Me.PictureBox6.Location = New System.Drawing.Point(277, 34)
        Me.PictureBox6.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.PictureBox6.Name = "PictureBox6"
        Me.PictureBox6.Size = New System.Drawing.Size(20, 20)
        Me.PictureBox6.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox6.TabIndex = 377
        Me.PictureBox6.TabStop = False
        '
        'RB_Verti
        '
        Me.RB_Verti.AutoSize = True
        Me.RB_Verti.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RB_Verti.Location = New System.Drawing.Point(8, 33)
        Me.RB_Verti.Name = "RB_Verti"
        Me.RB_Verti.Size = New System.Drawing.Size(162, 21)
        Me.RB_Verti.TabIndex = 1
        Me.RB_Verti.Text = "Vertical Drop-down List"
        Me.RB_Verti.UseVisualStyleBackColor = True
        '
        'PictureBox7
        '
        Me.PictureBox7.Image = CType(resources.GetObject("PictureBox7.Image"), System.Drawing.Image)
        Me.PictureBox7.Location = New System.Drawing.Point(277, 8)
        Me.PictureBox7.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.PictureBox7.Name = "PictureBox7"
        Me.PictureBox7.Size = New System.Drawing.Size(20, 20)
        Me.PictureBox7.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox7.TabIndex = 376
        Me.PictureBox7.TabStop = False
        '
        'RB_Horizon
        '
        Me.RB_Horizon.AutoSize = True
        Me.RB_Horizon.Checked = True
        Me.RB_Horizon.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RB_Horizon.Location = New System.Drawing.Point(8, 7)
        Me.RB_Horizon.Name = "RB_Horizon"
        Me.RB_Horizon.Size = New System.Drawing.Size(180, 21)
        Me.RB_Horizon.TabIndex = 0
        Me.RB_Horizon.TabStop = True
        Me.RB_Horizon.Text = "Horizontal Drop-down List"
        Me.RB_Horizon.UseVisualStyleBackColor = True
        '
        'CustomGroupBox4
        '
        Me.CustomGroupBox4.BackColor = System.Drawing.Color.White
        Me.CustomGroupBox4.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.CustomGroupBox4.Location = New System.Drawing.Point(1, 22)
        Me.CustomGroupBox4.Name = "CustomGroupBox4"
        Me.CustomGroupBox4.Size = New System.Drawing.Size(306, 62)
        Me.CustomGroupBox4.TabIndex = 0
        Me.CustomGroupBox4.TabStop = False
        Me.CustomGroupBox4.Text = "hhk"
        '
        'CustomGroupBox7
        '
        Me.CustomGroupBox7.BackColor = System.Drawing.Color.Black
        Me.CustomGroupBox7.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.CustomGroupBox7.Location = New System.Drawing.Point(1, 22)
        Me.CustomGroupBox7.Name = "CustomGroupBox7"
        Me.CustomGroupBox7.Size = New System.Drawing.Size(306, 62)
        Me.CustomGroupBox7.TabIndex = 0
        Me.CustomGroupBox7.TabStop = False
        '
        'Form30_Create_Dynamic_Drop_down_List
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(685, 486)
        Me.Controls.Add(Me.CB_text)
        Me.Controls.Add(Me.CB_descending)
        Me.Controls.Add(Me.CB_ascending)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Selection_destination)
        Me.Controls.Add(Me.TB_dest_range)
        Me.Controls.Add(Me.CustomGroupBox1)
        Me.Controls.Add(Me.CB_header)
        Me.Controls.Add(Me.Btn_OK)
        Me.Controls.Add(Me.Btn_Cancel)
        Me.Controls.Add(Me.CustomGroupBox2)
        Me.Controls.Add(Me.ComboBox2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Selection_source)
        Me.Controls.Add(Me.TB_src_range)
        Me.Controls.Add(Me.GB_list_option)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Form30_Create_Dynamic_Drop_down_List"
        Me.Text = "Create Dynamic Drop-down List"
        CType(Me.Selection_source, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Selection_destination, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox8, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.CustomGroupBox1.ResumeLayout(False)
        Me.CustomGroupBox3.ResumeLayout(False)
        Me.CustomGroupBox3.PerformLayout()
        CType(Me.PictureBox4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox5, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GB_list_option.ResumeLayout(False)
        Me.CustomGroupBox5.ResumeLayout(False)
        Me.CustomGroupBox5.PerformLayout()
        CType(Me.PictureBox6, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox7, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Btn_OK As Windows.Forms.Button
    Friend WithEvents Btn_Cancel As Windows.Forms.Button
    Friend WithEvents CustomGroupBox2 As CustomGroupBox
    Friend WithEvents ComboBox2 As Windows.Forms.ComboBox
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents Selection_source As Windows.Forms.PictureBox
    Friend WithEvents TB_src_range As Windows.Forms.TextBox
    Friend WithEvents CB_header As Windows.Forms.CheckBox
    Friend WithEvents CustomGroupBox1 As CustomGroupBox
    Friend WithEvents CustomGroupBox7 As CustomGroupBox
    Friend WithEvents PictureBox8 As Windows.Forms.PictureBox
    Friend WithEvents PictureBox1 As Windows.Forms.PictureBox
    Friend WithEvents RB_2_5_levels As Windows.Forms.RadioButton
    Friend WithEvents RB_2_levels As Windows.Forms.RadioButton
    Friend WithEvents GB_list_option As CustomGroupBox
    Friend WithEvents CustomGroupBox4 As CustomGroupBox
    Friend WithEvents PictureBox2 As Windows.Forms.PictureBox
    Friend WithEvents PictureBox3 As Windows.Forms.PictureBox
    Friend WithEvents RB_vertical As Windows.Forms.RadioButton
    Friend WithEvents RB_Horizontal As Windows.Forms.RadioButton
    Friend WithEvents Label2 As Windows.Forms.Label
    Friend WithEvents Selection_destination As Windows.Forms.PictureBox
    Friend WithEvents TB_dest_range As Windows.Forms.TextBox
    Friend WithEvents CB_ascending As Windows.Forms.CheckBox
    Friend WithEvents CB_descending As Windows.Forms.CheckBox
    Friend WithEvents CB_text As Windows.Forms.CheckBox
    Friend WithEvents CustomGroupBox3 As CustomGroupBox
    Friend WithEvents CustomGroupBox5 As CustomGroupBox
    Friend WithEvents RB_Verti As Windows.Forms.RadioButton
    Friend WithEvents RB_Horizon As Windows.Forms.RadioButton
    Friend WithEvents RB_Dropdown_35_Labels As Windows.Forms.RadioButton
    Friend WithEvents RB_Dropdown_2_Labels As Windows.Forms.RadioButton
    Friend WithEvents PictureBox6 As Windows.Forms.PictureBox
    Friend WithEvents PictureBox7 As Windows.Forms.PictureBox
    Friend WithEvents PictureBox4 As Windows.Forms.PictureBox
    Friend WithEvents PictureBox5 As Windows.Forms.PictureBox
End Class
