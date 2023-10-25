<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form33_ColorBasedDropDownList
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form33_ColorBasedDropDownList))
        Me.ColorDialog1 = New System.Windows.Forms.ColorDialog()
        Me.FlowLayoutPanel1 = New System.Windows.Forms.FlowLayoutPanel()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.PictureBox3 = New System.Windows.Forms.PictureBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Selection_source = New System.Windows.Forms.PictureBox()
        Me.TB_src_rng = New System.Windows.Forms.TextBox()
        Me.Selection_destination = New System.Windows.Forms.PictureBox()
        Me.TB_des_rng = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.List_Preview = New System.Windows.Forms.ListBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.btn_OK = New System.Windows.Forms.Button()
        Me.btn_Cancel = New System.Windows.Forms.Button()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.Backup_sheet = New System.Windows.Forms.CheckBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Btn_NC = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Btn_color = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.GB_sample = New VSTO_Addins.CustomGroupBox()
        Me.CustomGroupBox1 = New VSTO_Addins.CustomGroupBox()
        Me.CustomGroupBox7 = New VSTO_Addins.CustomGroupBox()
        Me.RB_Row = New System.Windows.Forms.RadioButton()
        Me.RB_cell = New System.Windows.Forms.RadioButton()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Selection_source, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Selection_destination, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.CustomGroupBox1.SuspendLayout()
        Me.CustomGroupBox7.SuspendLayout()
        Me.SuspendLayout()
        '
        'FlowLayoutPanel1
        '
        Me.FlowLayoutPanel1.BackColor = System.Drawing.Color.White
        Me.FlowLayoutPanel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.FlowLayoutPanel1.Location = New System.Drawing.Point(154, 262)
        Me.FlowLayoutPanel1.Name = "FlowLayoutPanel1"
        Me.FlowLayoutPanel1.Size = New System.Drawing.Size(186, 224)
        Me.FlowLayoutPanel1.TabIndex = 0
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(269, 15)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(20, 20)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 407
        Me.PictureBox1.TabStop = False
        Me.ToolTip1.SetToolTip(Me.PictureBox1, "Please select a range that contains a data validation list")
        '
        'PictureBox3
        '
        Me.PictureBox3.Image = CType(resources.GetObject("PictureBox3.Image"), System.Drawing.Image)
        Me.PictureBox3.Location = New System.Drawing.Point(212, 170)
        Me.PictureBox3.Name = "PictureBox3"
        Me.PictureBox3.Size = New System.Drawing.Size(20, 20)
        Me.PictureBox3.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox3.TabIndex = 408
        Me.PictureBox3.TabStop = False
        Me.ToolTip1.SetToolTip(Me.PictureBox3, "Please select a range intersecting data validation list")
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(15, 15)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(253, 17)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Drop-down List (Data Validation) Range :"
        '
        'Selection_source
        '
        Me.Selection_source.BackColor = System.Drawing.Color.White
        Me.Selection_source.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Selection_source.Image = CType(resources.GetObject("Selection_source.Image"), System.Drawing.Image)
        Me.Selection_source.Location = New System.Drawing.Point(319, 41)
        Me.Selection_source.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Selection_source.Name = "Selection_source"
        Me.Selection_source.Size = New System.Drawing.Size(24, 25)
        Me.Selection_source.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.Selection_source.TabIndex = 391
        Me.Selection_source.TabStop = False
        '
        'TB_src_rng
        '
        Me.TB_src_rng.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TB_src_rng.Location = New System.Drawing.Point(15, 41)
        Me.TB_src_rng.Name = "TB_src_rng"
        Me.TB_src_rng.Size = New System.Drawing.Size(328, 25)
        Me.TB_src_rng.TabIndex = 390
        '
        'Selection_destination
        '
        Me.Selection_destination.BackColor = System.Drawing.Color.White
        Me.Selection_destination.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Selection_destination.Image = CType(resources.GetObject("Selection_destination.Image"), System.Drawing.Image)
        Me.Selection_destination.Location = New System.Drawing.Point(319, 197)
        Me.Selection_destination.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Selection_destination.Name = "Selection_destination"
        Me.Selection_destination.Size = New System.Drawing.Size(24, 25)
        Me.Selection_destination.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.Selection_destination.TabIndex = 395
        Me.Selection_destination.TabStop = False
        '
        'TB_des_rng
        '
        Me.TB_des_rng.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TB_des_rng.Location = New System.Drawing.Point(15, 197)
        Me.TB_des_rng.Name = "TB_des_rng"
        Me.TB_des_rng.Size = New System.Drawing.Size(328, 25)
        Me.TB_des_rng.TabIndex = 394
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(15, 170)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(197, 17)
        Me.Label2.TabIndex = 393
        Me.Label2.Text = "Select range to highlight rows :"
        '
        'List_Preview
        '
        Me.List_Preview.BackColor = System.Drawing.SystemColors.Window
        Me.List_Preview.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawVariable
        Me.List_Preview.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.List_Preview.FormattingEnabled = True
        Me.List_Preview.ItemHeight = 17
        Me.List_Preview.Location = New System.Drawing.Point(15, 262)
        Me.List_Preview.Name = "List_Preview"
        Me.List_Preview.Size = New System.Drawing.Size(125, 259)
        Me.List_Preview.TabIndex = 397
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(15, 236)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(86, 17)
        Me.Label6.TabIndex = 396
        Me.Label6.Text = "List Preview :"
        '
        'btn_OK
        '
        Me.btn_OK.BackColor = System.Drawing.Color.White
        Me.btn_OK.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        Me.btn_OK.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btn_OK.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_OK.ForeColor = System.Drawing.Color.FromArgb(CType(CType(70, Byte), Integer), CType(CType(70, Byte), Integer), CType(CType(70, Byte), Integer))
        Me.btn_OK.Location = New System.Drawing.Point(554, 560)
        Me.btn_OK.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.btn_OK.Name = "btn_OK"
        Me.btn_OK.Size = New System.Drawing.Size(62, 26)
        Me.btn_OK.TabIndex = 402
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
        Me.btn_Cancel.Location = New System.Drawing.Point(627, 560)
        Me.btn_Cancel.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.btn_Cancel.Name = "btn_Cancel"
        Me.btn_Cancel.Size = New System.Drawing.Size(62, 26)
        Me.btn_Cancel.TabIndex = 401
        Me.btn_Cancel.Text = "Cancel"
        Me.btn_Cancel.UseVisualStyleBackColor = False
        '
        'ComboBox1
        '
        Me.ComboBox1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Items.AddRange(New Object() {"SOFTEKO", "About Us", "Help", "Feedback"})
        Me.ComboBox1.Location = New System.Drawing.Point(13, 562)
        Me.ComboBox1.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(100, 25)
        Me.ComboBox1.TabIndex = 400
        Me.ComboBox1.Text = "SOFTEKO"
        '
        'Backup_sheet
        '
        Me.Backup_sheet.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Backup_sheet.Location = New System.Drawing.Point(15, 526)
        Me.Backup_sheet.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Backup_sheet.Name = "Backup_sheet"
        Me.Backup_sheet.Size = New System.Drawing.Size(258, 29)
        Me.Backup_sheet.TabIndex = 399
        Me.Backup_sheet.Text = "Create a copy of the original worksheet"
        Me.Backup_sheet.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(154, 236)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(93, 17)
        Me.Label3.TabIndex = 403
        Me.Label3.Text = "Color Palette :"
        '
        'Btn_NC
        '
        Me.Btn_NC.BackColor = System.Drawing.Color.White
        Me.Btn_NC.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Btn_NC.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Btn_NC.Location = New System.Drawing.Point(159, 422)
        Me.Btn_NC.Name = "Btn_NC"
        Me.Btn_NC.Size = New System.Drawing.Size(176, 22)
        Me.Btn_NC.TabIndex = 404
        Me.Btn_NC.Text = "No Color"
        Me.Btn_NC.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.BackColor = System.Drawing.Color.White
        Me.Button2.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Button2.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button2.Location = New System.Drawing.Point(190, 454)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(144, 22)
        Me.Button2.TabIndex = 405
        Me.Button2.Text = "More Colors"
        Me.Button2.UseVisualStyleBackColor = False
        '
        'Btn_color
        '
        Me.Btn_color.BackColor = System.Drawing.Color.White
        Me.Btn_color.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Btn_color.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Btn_color.Location = New System.Drawing.Point(159, 454)
        Me.Btn_color.Name = "Btn_color"
        Me.Btn_color.Size = New System.Drawing.Size(24, 22)
        Me.Btn_color.TabIndex = 406
        Me.Btn_color.UseVisualStyleBackColor = False
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.White
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Button1.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.Location = New System.Drawing.Point(154, 494)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(186, 24)
        Me.Button1.TabIndex = 409
        Me.Button1.Text = "Clear all selected colors"
        Me.Button1.UseVisualStyleBackColor = False
        '
        'GB_sample
        '
        Me.GB_sample.BackColor = System.Drawing.Color.White
        Me.GB_sample.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.GB_sample.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GB_sample.Location = New System.Drawing.Point(382, 15)
        Me.GB_sample.Name = "GB_sample"
        Me.GB_sample.Size = New System.Drawing.Size(307, 503)
        Me.GB_sample.TabIndex = 398
        Me.GB_sample.TabStop = False
        Me.GB_sample.Text = "Sample Image"
        '
        'CustomGroupBox1
        '
        Me.CustomGroupBox1.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.CustomGroupBox1.Controls.Add(Me.CustomGroupBox7)
        Me.CustomGroupBox1.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CustomGroupBox1.Location = New System.Drawing.Point(15, 75)
        Me.CustomGroupBox1.Name = "CustomGroupBox1"
        Me.CustomGroupBox1.Size = New System.Drawing.Size(328, 84)
        Me.CustomGroupBox1.TabIndex = 392
        Me.CustomGroupBox1.TabStop = False
        Me.CustomGroupBox1.Text = "Apply Color to"
        '
        'CustomGroupBox7
        '
        Me.CustomGroupBox7.BackColor = System.Drawing.Color.White
        Me.CustomGroupBox7.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.CustomGroupBox7.Controls.Add(Me.RB_Row)
        Me.CustomGroupBox7.Controls.Add(Me.RB_cell)
        Me.CustomGroupBox7.Location = New System.Drawing.Point(1, 22)
        Me.CustomGroupBox7.Name = "CustomGroupBox7"
        Me.CustomGroupBox7.Size = New System.Drawing.Size(327, 62)
        Me.CustomGroupBox7.TabIndex = 0
        Me.CustomGroupBox7.TabStop = False
        '
        'RB_Row
        '
        Me.RB_Row.AutoSize = True
        Me.RB_Row.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RB_Row.Location = New System.Drawing.Point(8, 32)
        Me.RB_Row.Name = "RB_Row"
        Me.RB_Row.Size = New System.Drawing.Size(237, 21)
        Me.RB_Row.TabIndex = 1
        Me.RB_Row.Text = "Full row of the drop-down list range"
        Me.RB_Row.UseVisualStyleBackColor = True
        '
        'RB_cell
        '
        Me.RB_cell.AutoSize = True
        Me.RB_cell.Checked = True
        Me.RB_cell.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RB_cell.Location = New System.Drawing.Point(8, 6)
        Me.RB_cell.Name = "RB_cell"
        Me.RB_cell.Size = New System.Drawing.Size(265, 21)
        Me.RB_cell.TabIndex = 0
        Me.RB_cell.TabStop = True
        Me.RB_cell.Text = "Only the cell that contains data validation"
        Me.RB_cell.UseVisualStyleBackColor = True
        '
        'Form33_ColorBasedDropDownList
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.ClientSize = New System.Drawing.Size(713, 604)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.PictureBox3)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.Btn_color)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Btn_NC)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.btn_OK)
        Me.Controls.Add(Me.btn_Cancel)
        Me.Controls.Add(Me.ComboBox1)
        Me.Controls.Add(Me.Backup_sheet)
        Me.Controls.Add(Me.GB_sample)
        Me.Controls.Add(Me.List_Preview)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Selection_destination)
        Me.Controls.Add(Me.TB_des_rng)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.CustomGroupBox1)
        Me.Controls.Add(Me.Selection_source)
        Me.Controls.Add(Me.TB_src_rng)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.FlowLayoutPanel1)
        Me.ForeColor = System.Drawing.SystemColors.ControlText
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Form33_ColorBasedDropDownList"
        Me.Text = "Color Based Drop-down List"
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Selection_source, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Selection_destination, System.ComponentModel.ISupportInitialize).EndInit()
        Me.CustomGroupBox1.ResumeLayout(False)
        Me.CustomGroupBox7.ResumeLayout(False)
        Me.CustomGroupBox7.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ColorDialog1 As Windows.Forms.ColorDialog
    Friend WithEvents FlowLayoutPanel1 As Windows.Forms.FlowLayoutPanel
    Friend WithEvents ToolTip1 As Windows.Forms.ToolTip
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents Selection_source As Windows.Forms.PictureBox
    Friend WithEvents TB_src_rng As Windows.Forms.TextBox
    Friend WithEvents CustomGroupBox1 As CustomGroupBox
    Friend WithEvents CustomGroupBox7 As CustomGroupBox
    Friend WithEvents RB_Row As Windows.Forms.RadioButton
    Friend WithEvents RB_cell As Windows.Forms.RadioButton
    Friend WithEvents Selection_destination As Windows.Forms.PictureBox
    Friend WithEvents TB_des_rng As Windows.Forms.TextBox
    Friend WithEvents Label2 As Windows.Forms.Label
    Friend WithEvents List_Preview As Windows.Forms.ListBox
    Friend WithEvents Label6 As Windows.Forms.Label
    Friend WithEvents GB_sample As CustomGroupBox
    Friend WithEvents btn_OK As Windows.Forms.Button
    Friend WithEvents btn_Cancel As Windows.Forms.Button
    Friend WithEvents ComboBox1 As Windows.Forms.ComboBox
    Friend WithEvents Backup_sheet As Windows.Forms.CheckBox
    Friend WithEvents Label3 As Windows.Forms.Label
    Friend WithEvents Btn_NC As Windows.Forms.Button
    Friend WithEvents Button2 As Windows.Forms.Button
    Friend WithEvents Btn_color As Windows.Forms.Button
    Friend WithEvents PictureBox1 As Windows.Forms.PictureBox
    Friend WithEvents PictureBox3 As Windows.Forms.PictureBox
    Friend WithEvents Button1 As Windows.Forms.Button
End Class
