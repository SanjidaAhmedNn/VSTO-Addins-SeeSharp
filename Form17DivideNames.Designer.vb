<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form17DivideNames
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form17DivideNames))
        Me.Selection = New System.Windows.Forms.PictureBox()
        Me.AutoSelection = New System.Windows.Forms.PictureBox()
        Me.txtSourceRange = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.PictureBox7 = New System.Windows.Forms.PictureBox()
        Me.btnOK = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.CB_Backup_Sheet = New System.Windows.Forms.CheckBox()
        Me.CB_Keep_Formatting = New System.Windows.Forms.CheckBox()
        Me.CB_Add_Header = New System.Windows.Forms.CheckBox()
        Me.CustomGroupBox4 = New VSTO_Addins.CustomGroupBox()
        Me.CustomGroupBox10 = New VSTO_Addins.CustomGroupBox()
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        Me.destinationSelection = New System.Windows.Forms.PictureBox()
        Me.txtDestRange = New System.Windows.Forms.TextBox()
        Me.lbl_destRange_Selection = New System.Windows.Forms.Label()
        Me.RB_Different_Range = New System.Windows.Forms.RadioButton()
        Me.RB_Same_As_Source_Range = New System.Windows.Forms.RadioButton()
        Me.CustomGroupBox5 = New VSTO_Addins.CustomGroupBox()
        Me.CustomPanel1 = New VSTO_Addins.CustomPanel()
        Me.CustomGroupBox6 = New VSTO_Addins.CustomGroupBox()
        Me.CustomPanel2 = New VSTO_Addins.CustomPanel()
        Me.CustomGroupBox1 = New VSTO_Addins.CustomGroupBox()
        Me.CustomGroupBox7 = New VSTO_Addins.CustomGroupBox()
        Me.PictureBox11 = New System.Windows.Forms.PictureBox()
        Me.CB_Select_All = New System.Windows.Forms.CheckBox()
        Me.PictureBox8 = New System.Windows.Forms.PictureBox()
        Me.PictureBox9 = New System.Windows.Forms.PictureBox()
        Me.PictureBox10 = New System.Windows.Forms.PictureBox()
        Me.PictureBox6 = New System.Windows.Forms.PictureBox()
        Me.PictureBox4 = New System.Windows.Forms.PictureBox()
        Me.PictureBox3 = New System.Windows.Forms.PictureBox()
        Me.PictureBox5 = New System.Windows.Forms.PictureBox()
        Me.CB_Name_Suffix = New System.Windows.Forms.CheckBox()
        Me.CB_Title = New System.Windows.Forms.CheckBox()
        Me.CB_Name_Abbreviations = New System.Windows.Forms.CheckBox()
        Me.CB_Last_Name = New System.Windows.Forms.CheckBox()
        Me.CB_Last_Name_Prefix = New System.Windows.Forms.CheckBox()
        Me.CB_Middle_Name = New System.Windows.Forms.CheckBox()
        Me.CB_First_Name = New System.Windows.Forms.CheckBox()
        CType(Me.Selection, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AutoSelection, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox7, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.CustomGroupBox4.SuspendLayout()
        Me.CustomGroupBox10.SuspendLayout()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.destinationSelection, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.CustomGroupBox5.SuspendLayout()
        Me.CustomGroupBox6.SuspendLayout()
        Me.CustomGroupBox1.SuspendLayout()
        Me.CustomGroupBox7.SuspendLayout()
        CType(Me.PictureBox11, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox8, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox9, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox10, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox6, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox4, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox5, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Selection
        '
        Me.Selection.BackColor = System.Drawing.Color.White
        Me.Selection.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Selection.Image = CType(resources.GetObject("Selection.Image"), System.Drawing.Image)
        Me.Selection.Location = New System.Drawing.Point(227, 40)
        Me.Selection.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Selection.Name = "Selection"
        Me.Selection.Size = New System.Drawing.Size(24, 25)
        Me.Selection.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.Selection.TabIndex = 224
        Me.Selection.TabStop = False
        '
        'AutoSelection
        '
        Me.AutoSelection.BackColor = System.Drawing.Color.White
        Me.AutoSelection.Image = CType(resources.GetObject("AutoSelection.Image"), System.Drawing.Image)
        Me.AutoSelection.Location = New System.Drawing.Point(198, 41)
        Me.AutoSelection.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.AutoSelection.Name = "AutoSelection"
        Me.AutoSelection.Size = New System.Drawing.Size(24, 23)
        Me.AutoSelection.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.AutoSelection.TabIndex = 223
        Me.AutoSelection.TabStop = False
        '
        'txtSourceRange
        '
        Me.txtSourceRange.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtSourceRange.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtSourceRange.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSourceRange.Location = New System.Drawing.Point(15, 40)
        Me.txtSourceRange.Name = "txtSourceRange"
        Me.txtSourceRange.Size = New System.Drawing.Size(236, 25)
        Me.txtSourceRange.TabIndex = 222
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(15, 14)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(98, 17)
        Me.Label1.TabIndex = 221
        Me.Label1.Text = "Source Range :"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'PictureBox7
        '
        Me.PictureBox7.Image = CType(resources.GetObject("PictureBox7.Image"), System.Drawing.Image)
        Me.PictureBox7.Location = New System.Drawing.Point(380, 209)
        Me.PictureBox7.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.PictureBox7.Name = "PictureBox7"
        Me.PictureBox7.Size = New System.Drawing.Size(52, 60)
        Me.PictureBox7.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox7.TabIndex = 220
        Me.PictureBox7.TabStop = False
        '
        'btnOK
        '
        Me.btnOK.BackColor = System.Drawing.Color.White
        Me.btnOK.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnOK.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOK.Location = New System.Drawing.Point(394, 506)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(62, 26)
        Me.btnOK.TabIndex = 219
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = False
        '
        'btnCancel
        '
        Me.btnCancel.BackColor = System.Drawing.Color.White
        Me.btnCancel.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnCancel.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCancel.Location = New System.Drawing.Point(472, 506)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(62, 26)
        Me.btnCancel.TabIndex = 218
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = False
        '
        'ComboBox1
        '
        Me.ComboBox1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Location = New System.Drawing.Point(15, 506)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(154, 25)
        Me.ComboBox1.TabIndex = 215
        Me.ComboBox1.Text = "SOFTEKO"
        '
        'CB_Backup_Sheet
        '
        Me.CB_Backup_Sheet.AutoSize = True
        Me.CB_Backup_Sheet.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CB_Backup_Sheet.Location = New System.Drawing.Point(15, 475)
        Me.CB_Backup_Sheet.Name = "CB_Backup_Sheet"
        Me.CB_Backup_Sheet.Size = New System.Drawing.Size(257, 21)
        Me.CB_Backup_Sheet.TabIndex = 214
        Me.CB_Backup_Sheet.Text = "Create a copy of the original worksheet"
        Me.CB_Backup_Sheet.UseVisualStyleBackColor = True
        '
        'CB_Keep_Formatting
        '
        Me.CB_Keep_Formatting.AutoSize = True
        Me.CB_Keep_Formatting.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CB_Keep_Formatting.Location = New System.Drawing.Point(15, 304)
        Me.CB_Keep_Formatting.Name = "CB_Keep_Formatting"
        Me.CB_Keep_Formatting.Size = New System.Drawing.Size(122, 21)
        Me.CB_Keep_Formatting.TabIndex = 226
        Me.CB_Keep_Formatting.Text = "Keep formatting"
        Me.CB_Keep_Formatting.UseVisualStyleBackColor = True
        '
        'CB_Add_Header
        '
        Me.CB_Add_Header.AutoSize = True
        Me.CB_Add_Header.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CB_Add_Header.Location = New System.Drawing.Point(150, 304)
        Me.CB_Add_Header.Name = "CB_Add_Header"
        Me.CB_Add_Header.Size = New System.Drawing.Size(98, 21)
        Me.CB_Add_Header.TabIndex = 227
        Me.CB_Add_Header.Text = "Add Header"
        Me.CB_Add_Header.UseVisualStyleBackColor = True
        '
        'CustomGroupBox4
        '
        Me.CustomGroupBox4.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.CustomGroupBox4.Controls.Add(Me.CustomGroupBox10)
        Me.CustomGroupBox4.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CustomGroupBox4.Location = New System.Drawing.Point(15, 331)
        Me.CustomGroupBox4.Name = "CustomGroupBox4"
        Me.CustomGroupBox4.Size = New System.Drawing.Size(236, 137)
        Me.CustomGroupBox4.TabIndex = 225
        Me.CustomGroupBox4.TabStop = False
        Me.CustomGroupBox4.Text = "Destination Range"
        '
        'CustomGroupBox10
        '
        Me.CustomGroupBox10.BackColor = System.Drawing.Color.White
        Me.CustomGroupBox10.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.CustomGroupBox10.Controls.Add(Me.PictureBox2)
        Me.CustomGroupBox10.Controls.Add(Me.destinationSelection)
        Me.CustomGroupBox10.Controls.Add(Me.txtDestRange)
        Me.CustomGroupBox10.Controls.Add(Me.lbl_destRange_Selection)
        Me.CustomGroupBox10.Controls.Add(Me.RB_Different_Range)
        Me.CustomGroupBox10.Controls.Add(Me.RB_Same_As_Source_Range)
        Me.CustomGroupBox10.Location = New System.Drawing.Point(1, 22)
        Me.CustomGroupBox10.Name = "CustomGroupBox10"
        Me.CustomGroupBox10.Size = New System.Drawing.Size(235, 115)
        Me.CustomGroupBox10.TabIndex = 0
        Me.CustomGroupBox10.TabStop = False
        '
        'PictureBox2
        '
        Me.PictureBox2.Image = CType(resources.GetObject("PictureBox2.Image"), System.Drawing.Image)
        Me.PictureBox2.Location = New System.Drawing.Point(25, 58)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(14, 14)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox2.TabIndex = 208
        Me.PictureBox2.TabStop = False
        '
        'destinationSelection
        '
        Me.destinationSelection.BackColor = System.Drawing.Color.White
        Me.destinationSelection.Image = CType(resources.GetObject("destinationSelection.Image"), System.Drawing.Image)
        Me.destinationSelection.Location = New System.Drawing.Point(193, 81)
        Me.destinationSelection.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.destinationSelection.Name = "destinationSelection"
        Me.destinationSelection.Size = New System.Drawing.Size(24, 23)
        Me.destinationSelection.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.destinationSelection.TabIndex = 207
        Me.destinationSelection.TabStop = False
        '
        'txtDestRange
        '
        Me.txtDestRange.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtDestRange.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDestRange.Font = New System.Drawing.Font("Segoe UI", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDestRange.Location = New System.Drawing.Point(25, 80)
        Me.txtDestRange.Name = "txtDestRange"
        Me.txtDestRange.Size = New System.Drawing.Size(193, 25)
        Me.txtDestRange.TabIndex = 206
        '
        'lbl_destRange_Selection
        '
        Me.lbl_destRange_Selection.AutoSize = True
        Me.lbl_destRange_Selection.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_destRange_Selection.Location = New System.Drawing.Point(42, 56)
        Me.lbl_destRange_Selection.Name = "lbl_destRange_Selection"
        Me.lbl_destRange_Selection.Size = New System.Drawing.Size(109, 17)
        Me.lbl_destRange_Selection.TabIndex = 2
        Me.lbl_destRange_Selection.Text = "Select the range :"
        '
        'RB_Different_Range
        '
        Me.RB_Different_Range.AutoSize = True
        Me.RB_Different_Range.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RB_Different_Range.Location = New System.Drawing.Point(8, 31)
        Me.RB_Different_Range.Name = "RB_Different_Range"
        Me.RB_Different_Range.Size = New System.Drawing.Size(185, 21)
        Me.RB_Different_Range.TabIndex = 1
        Me.RB_Different_Range.TabStop = True
        Me.RB_Different_Range.Text = "Store into a different range"
        Me.RB_Different_Range.UseVisualStyleBackColor = True
        '
        'RB_Same_As_Source_Range
        '
        Me.RB_Same_As_Source_Range.AutoSize = True
        Me.RB_Same_As_Source_Range.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RB_Same_As_Source_Range.Location = New System.Drawing.Point(8, 6)
        Me.RB_Same_As_Source_Range.Name = "RB_Same_As_Source_Range"
        Me.RB_Same_As_Source_Range.Size = New System.Drawing.Size(178, 21)
        Me.RB_Same_As_Source_Range.TabIndex = 0
        Me.RB_Same_As_Source_Range.TabStop = True
        Me.RB_Same_As_Source_Range.Text = "Same as the source range"
        Me.RB_Same_As_Source_Range.UseVisualStyleBackColor = True
        '
        'CustomGroupBox5
        '
        Me.CustomGroupBox5.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.CustomGroupBox5.Controls.Add(Me.CustomPanel1)
        Me.CustomGroupBox5.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CustomGroupBox5.Location = New System.Drawing.Point(283, 22)
        Me.CustomGroupBox5.Name = "CustomGroupBox5"
        Me.CustomGroupBox5.Size = New System.Drawing.Size(252, 180)
        Me.CustomGroupBox5.TabIndex = 216
        Me.CustomGroupBox5.TabStop = False
        Me.CustomGroupBox5.Text = "Input Range"
        '
        'CustomPanel1
        '
        Me.CustomPanel1.BackColor = System.Drawing.Color.White
        Me.CustomPanel1.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.CustomPanel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.CustomPanel1.BorderWidth = 1
        Me.CustomPanel1.Location = New System.Drawing.Point(1, 30)
        Me.CustomPanel1.Name = "CustomPanel1"
        Me.CustomPanel1.Size = New System.Drawing.Size(250, 150)
        Me.CustomPanel1.TabIndex = 0
        '
        'CustomGroupBox6
        '
        Me.CustomGroupBox6.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.CustomGroupBox6.Controls.Add(Me.CustomPanel2)
        Me.CustomGroupBox6.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CustomGroupBox6.Location = New System.Drawing.Point(284, 280)
        Me.CustomGroupBox6.Name = "CustomGroupBox6"
        Me.CustomGroupBox6.Size = New System.Drawing.Size(251, 180)
        Me.CustomGroupBox6.TabIndex = 217
        Me.CustomGroupBox6.TabStop = False
        Me.CustomGroupBox6.Text = "Expected Output"
        '
        'CustomPanel2
        '
        Me.CustomPanel2.BackColor = System.Drawing.Color.White
        Me.CustomPanel2.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.CustomPanel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.CustomPanel2.BorderWidth = 1
        Me.CustomPanel2.Location = New System.Drawing.Point(1, 30)
        Me.CustomPanel2.Name = "CustomPanel2"
        Me.CustomPanel2.Size = New System.Drawing.Size(250, 150)
        Me.CustomPanel2.TabIndex = 11
        '
        'CustomGroupBox1
        '
        Me.CustomGroupBox1.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.CustomGroupBox1.Controls.Add(Me.CustomGroupBox7)
        Me.CustomGroupBox1.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CustomGroupBox1.Location = New System.Drawing.Point(15, 72)
        Me.CustomGroupBox1.Name = "CustomGroupBox1"
        Me.CustomGroupBox1.Size = New System.Drawing.Size(236, 224)
        Me.CustomGroupBox1.TabIndex = 212
        Me.CustomGroupBox1.TabStop = False
        Me.CustomGroupBox1.Text = "Divide by"
        '
        'CustomGroupBox7
        '
        Me.CustomGroupBox7.BackColor = System.Drawing.Color.White
        Me.CustomGroupBox7.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.CustomGroupBox7.Controls.Add(Me.PictureBox11)
        Me.CustomGroupBox7.Controls.Add(Me.CB_Select_All)
        Me.CustomGroupBox7.Controls.Add(Me.PictureBox8)
        Me.CustomGroupBox7.Controls.Add(Me.PictureBox9)
        Me.CustomGroupBox7.Controls.Add(Me.PictureBox10)
        Me.CustomGroupBox7.Controls.Add(Me.PictureBox6)
        Me.CustomGroupBox7.Controls.Add(Me.PictureBox4)
        Me.CustomGroupBox7.Controls.Add(Me.PictureBox3)
        Me.CustomGroupBox7.Controls.Add(Me.PictureBox5)
        Me.CustomGroupBox7.Controls.Add(Me.CB_Name_Suffix)
        Me.CustomGroupBox7.Controls.Add(Me.CB_Title)
        Me.CustomGroupBox7.Controls.Add(Me.CB_Name_Abbreviations)
        Me.CustomGroupBox7.Controls.Add(Me.CB_Last_Name)
        Me.CustomGroupBox7.Controls.Add(Me.CB_Last_Name_Prefix)
        Me.CustomGroupBox7.Controls.Add(Me.CB_Middle_Name)
        Me.CustomGroupBox7.Controls.Add(Me.CB_First_Name)
        Me.CustomGroupBox7.Location = New System.Drawing.Point(1, 22)
        Me.CustomGroupBox7.Name = "CustomGroupBox7"
        Me.CustomGroupBox7.Size = New System.Drawing.Size(235, 202)
        Me.CustomGroupBox7.TabIndex = 0
        Me.CustomGroupBox7.TabStop = False
        '
        'PictureBox11
        '
        Me.PictureBox11.Image = CType(resources.GetObject("PictureBox11.Image"), System.Drawing.Image)
        Me.PictureBox11.Location = New System.Drawing.Point(194, 175)
        Me.PictureBox11.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.PictureBox11.Name = "PictureBox11"
        Me.PictureBox11.Size = New System.Drawing.Size(20, 20)
        Me.PictureBox11.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox11.TabIndex = 217
        Me.PictureBox11.TabStop = False
        '
        'CB_Select_All
        '
        Me.CB_Select_All.AutoSize = True
        Me.CB_Select_All.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CB_Select_All.Location = New System.Drawing.Point(9, 7)
        Me.CB_Select_All.Name = "CB_Select_All"
        Me.CB_Select_All.Size = New System.Drawing.Size(79, 21)
        Me.CB_Select_All.TabIndex = 216
        Me.CB_Select_All.Text = "Select All"
        Me.CB_Select_All.UseVisualStyleBackColor = True
        '
        'PictureBox8
        '
        Me.PictureBox8.Image = CType(resources.GetObject("PictureBox8.Image"), System.Drawing.Image)
        Me.PictureBox8.Location = New System.Drawing.Point(194, 151)
        Me.PictureBox8.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.PictureBox8.Name = "PictureBox8"
        Me.PictureBox8.Size = New System.Drawing.Size(20, 20)
        Me.PictureBox8.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox8.TabIndex = 212
        Me.PictureBox8.TabStop = False
        '
        'PictureBox9
        '
        Me.PictureBox9.Image = CType(resources.GetObject("PictureBox9.Image"), System.Drawing.Image)
        Me.PictureBox9.Location = New System.Drawing.Point(194, 127)
        Me.PictureBox9.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.PictureBox9.Name = "PictureBox9"
        Me.PictureBox9.Size = New System.Drawing.Size(20, 20)
        Me.PictureBox9.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox9.TabIndex = 213
        Me.PictureBox9.TabStop = False
        '
        'PictureBox10
        '
        Me.PictureBox10.Image = CType(resources.GetObject("PictureBox10.Image"), System.Drawing.Image)
        Me.PictureBox10.Location = New System.Drawing.Point(194, 103)
        Me.PictureBox10.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.PictureBox10.Name = "PictureBox10"
        Me.PictureBox10.Size = New System.Drawing.Size(20, 20)
        Me.PictureBox10.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox10.TabIndex = 214
        Me.PictureBox10.TabStop = False
        '
        'PictureBox6
        '
        Me.PictureBox6.Image = CType(resources.GetObject("PictureBox6.Image"), System.Drawing.Image)
        Me.PictureBox6.Location = New System.Drawing.Point(194, 31)
        Me.PictureBox6.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.PictureBox6.Name = "PictureBox6"
        Me.PictureBox6.Size = New System.Drawing.Size(20, 20)
        Me.PictureBox6.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox6.TabIndex = 215
        Me.PictureBox6.TabStop = False
        '
        'PictureBox4
        '
        Me.PictureBox4.Image = CType(resources.GetObject("PictureBox4.Image"), System.Drawing.Image)
        Me.PictureBox4.Location = New System.Drawing.Point(194, 55)
        Me.PictureBox4.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.PictureBox4.Name = "PictureBox4"
        Me.PictureBox4.Size = New System.Drawing.Size(20, 20)
        Me.PictureBox4.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox4.TabIndex = 214
        Me.PictureBox4.TabStop = False
        '
        'PictureBox3
        '
        Me.PictureBox3.Image = CType(resources.GetObject("PictureBox3.Image"), System.Drawing.Image)
        Me.PictureBox3.Location = New System.Drawing.Point(194, 79)
        Me.PictureBox3.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.PictureBox3.Name = "PictureBox3"
        Me.PictureBox3.Size = New System.Drawing.Size(20, 20)
        Me.PictureBox3.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox3.TabIndex = 213
        Me.PictureBox3.TabStop = False
        '
        'PictureBox5
        '
        Me.PictureBox5.BackColor = System.Drawing.Color.White
        Me.PictureBox5.Image = CType(resources.GetObject("PictureBox5.Image"), System.Drawing.Image)
        Me.PictureBox5.Location = New System.Drawing.Point(194, 7)
        Me.PictureBox5.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.PictureBox5.Name = "PictureBox5"
        Me.PictureBox5.Size = New System.Drawing.Size(20, 20)
        Me.PictureBox5.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox5.TabIndex = 212
        Me.PictureBox5.TabStop = False
        '
        'CB_Name_Suffix
        '
        Me.CB_Name_Suffix.AutoSize = True
        Me.CB_Name_Suffix.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CB_Name_Suffix.Location = New System.Drawing.Point(9, 30)
        Me.CB_Name_Suffix.Name = "CB_Name_Suffix"
        Me.CB_Name_Suffix.Size = New System.Drawing.Size(51, 21)
        Me.CB_Name_Suffix.TabIndex = 6
        Me.CB_Name_Suffix.Text = "Title"
        Me.CB_Name_Suffix.UseVisualStyleBackColor = True
        '
        'CB_Title
        '
        Me.CB_Title.AutoSize = True
        Me.CB_Title.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CB_Title.Location = New System.Drawing.Point(9, 54)
        Me.CB_Title.Name = "CB_Title"
        Me.CB_Title.Size = New System.Drawing.Size(90, 21)
        Me.CB_Title.TabIndex = 5
        Me.CB_Title.Text = "First Name"
        Me.CB_Title.UseVisualStyleBackColor = True
        '
        'CB_Name_Abbreviations
        '
        Me.CB_Name_Abbreviations.AutoSize = True
        Me.CB_Name_Abbreviations.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CB_Name_Abbreviations.Location = New System.Drawing.Point(9, 78)
        Me.CB_Name_Abbreviations.Name = "CB_Name_Abbreviations"
        Me.CB_Name_Abbreviations.Size = New System.Drawing.Size(107, 21)
        Me.CB_Name_Abbreviations.TabIndex = 4
        Me.CB_Name_Abbreviations.Text = "Middle Name"
        Me.CB_Name_Abbreviations.UseVisualStyleBackColor = True
        '
        'CB_Last_Name
        '
        Me.CB_Last_Name.AutoSize = True
        Me.CB_Last_Name.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CB_Last_Name.Location = New System.Drawing.Point(9, 102)
        Me.CB_Last_Name.Name = "CB_Last_Name"
        Me.CB_Last_Name.Size = New System.Drawing.Size(125, 21)
        Me.CB_Last_Name.TabIndex = 3
        Me.CB_Last_Name.Text = "Last Name Prefix"
        Me.CB_Last_Name.UseVisualStyleBackColor = True
        '
        'CB_Last_Name_Prefix
        '
        Me.CB_Last_Name_Prefix.AutoSize = True
        Me.CB_Last_Name_Prefix.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CB_Last_Name_Prefix.Location = New System.Drawing.Point(9, 126)
        Me.CB_Last_Name_Prefix.Name = "CB_Last_Name_Prefix"
        Me.CB_Last_Name_Prefix.Size = New System.Drawing.Size(89, 21)
        Me.CB_Last_Name_Prefix.TabIndex = 2
        Me.CB_Last_Name_Prefix.Text = "Last Name"
        Me.CB_Last_Name_Prefix.UseVisualStyleBackColor = True
        '
        'CB_Middle_Name
        '
        Me.CB_Middle_Name.AutoSize = True
        Me.CB_Middle_Name.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CB_Middle_Name.Location = New System.Drawing.Point(9, 150)
        Me.CB_Middle_Name.Name = "CB_Middle_Name"
        Me.CB_Middle_Name.Size = New System.Drawing.Size(97, 21)
        Me.CB_Middle_Name.TabIndex = 1
        Me.CB_Middle_Name.Text = "Name Suffix"
        Me.CB_Middle_Name.UseVisualStyleBackColor = True
        '
        'CB_First_Name
        '
        Me.CB_First_Name.AutoSize = True
        Me.CB_First_Name.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CB_First_Name.Location = New System.Drawing.Point(9, 174)
        Me.CB_First_Name.Name = "CB_First_Name"
        Me.CB_First_Name.Size = New System.Drawing.Size(146, 21)
        Me.CB_First_Name.TabIndex = 0
        Me.CB_First_Name.Text = "Name Abbreviations"
        Me.CB_First_Name.UseVisualStyleBackColor = True
        '
        'Form17DivideNames
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.ClientSize = New System.Drawing.Size(559, 549)
        Me.Controls.Add(Me.CB_Add_Header)
        Me.Controls.Add(Me.PictureBox7)
        Me.Controls.Add(Me.CB_Keep_Formatting)
        Me.Controls.Add(Me.CustomGroupBox4)
        Me.Controls.Add(Me.Selection)
        Me.Controls.Add(Me.AutoSelection)
        Me.Controls.Add(Me.txtSourceRange)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.CustomGroupBox5)
        Me.Controls.Add(Me.CustomGroupBox6)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.ComboBox1)
        Me.Controls.Add(Me.CustomGroupBox1)
        Me.Controls.Add(Me.CB_Backup_Sheet)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Form17DivideNames"
        Me.Text = "Divide Names"
        CType(Me.Selection, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AutoSelection, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox7, System.ComponentModel.ISupportInitialize).EndInit()
        Me.CustomGroupBox4.ResumeLayout(False)
        Me.CustomGroupBox10.ResumeLayout(False)
        Me.CustomGroupBox10.PerformLayout()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.destinationSelection, System.ComponentModel.ISupportInitialize).EndInit()
        Me.CustomGroupBox5.ResumeLayout(False)
        Me.CustomGroupBox6.ResumeLayout(False)
        Me.CustomGroupBox1.ResumeLayout(False)
        Me.CustomGroupBox7.ResumeLayout(False)
        Me.CustomGroupBox7.PerformLayout()
        CType(Me.PictureBox11, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox8, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox9, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox10, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox6, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox5, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents PictureBox8 As Windows.Forms.PictureBox
    Friend WithEvents PictureBox9 As Windows.Forms.PictureBox
    Friend WithEvents PictureBox10 As Windows.Forms.PictureBox
    Friend WithEvents PictureBox6 As Windows.Forms.PictureBox
    Friend WithEvents PictureBox3 As Windows.Forms.PictureBox
    Friend WithEvents Selection As Windows.Forms.PictureBox
    Friend WithEvents AutoSelection As Windows.Forms.PictureBox
    Friend WithEvents txtSourceRange As Windows.Forms.TextBox
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents PictureBox4 As Windows.Forms.PictureBox
    Friend WithEvents PictureBox5 As Windows.Forms.PictureBox
    Friend WithEvents CustomGroupBox5 As CustomGroupBox
    Friend WithEvents CustomPanel1 As CustomPanel
    Friend WithEvents CustomGroupBox6 As CustomGroupBox
    Friend WithEvents CustomPanel2 As CustomPanel
    Friend WithEvents PictureBox7 As Windows.Forms.PictureBox
    Friend WithEvents btnOK As Windows.Forms.Button
    Friend WithEvents btnCancel As Windows.Forms.Button
    Friend WithEvents CustomGroupBox7 As CustomGroupBox
    Friend WithEvents CB_Name_Suffix As Windows.Forms.CheckBox
    Friend WithEvents CB_Title As Windows.Forms.CheckBox
    Friend WithEvents CB_Name_Abbreviations As Windows.Forms.CheckBox
    Friend WithEvents CB_Last_Name As Windows.Forms.CheckBox
    Friend WithEvents CB_Last_Name_Prefix As Windows.Forms.CheckBox
    Friend WithEvents CB_Middle_Name As Windows.Forms.CheckBox
    Friend WithEvents CB_First_Name As Windows.Forms.CheckBox
    Friend WithEvents ComboBox1 As Windows.Forms.ComboBox
    Friend WithEvents CustomGroupBox1 As CustomGroupBox
    Friend WithEvents CB_Backup_Sheet As Windows.Forms.CheckBox
    Friend WithEvents CustomGroupBox4 As CustomGroupBox
    Friend WithEvents CustomGroupBox10 As CustomGroupBox
    Friend WithEvents PictureBox2 As Windows.Forms.PictureBox
    Friend WithEvents destinationSelection As Windows.Forms.PictureBox
    Friend WithEvents txtDestRange As Windows.Forms.TextBox
    Friend WithEvents lbl_destRange_Selection As Windows.Forms.Label
    Friend WithEvents RB_Different_Range As Windows.Forms.RadioButton
    Friend WithEvents RB_Same_As_Source_Range As Windows.Forms.RadioButton
    Friend WithEvents CB_Keep_Formatting As Windows.Forms.CheckBox
    Friend WithEvents PictureBox11 As Windows.Forms.PictureBox
    Friend WithEvents CB_Select_All As Windows.Forms.CheckBox
    Friend WithEvents CB_Add_Header As Windows.Forms.CheckBox
End Class
