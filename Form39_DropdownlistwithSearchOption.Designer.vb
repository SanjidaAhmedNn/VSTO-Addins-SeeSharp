<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form39_DropdownlistwithSearchOption
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form39_DropdownlistwithSearchOption))
        Me.Btn_OK = New System.Windows.Forms.Button()
        Me.Btn_Cancel = New System.Windows.Forms.Button()
        Me.CB_About = New System.Windows.Forms.ComboBox()
        Me.GB_Sample = New VSTO_Addins.CustomGroupBox()
        Me.CustomGroupBox4 = New VSTO_Addins.CustomGroupBox()
        Me.CustomGroupBox6 = New VSTO_Addins.CustomGroupBox()
        Me.Selection_source = New System.Windows.Forms.PictureBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TB_src_rng = New System.Windows.Forms.TextBox()
        Me.CB_Source = New System.Windows.Forms.ComboBox()
        Me.CustomGroupBox4.SuspendLayout()
        Me.CustomGroupBox6.SuspendLayout()
        CType(Me.Selection_source, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Btn_OK
        '
        Me.Btn_OK.BackColor = System.Drawing.Color.White
        Me.Btn_OK.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Btn_OK.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Btn_OK.Location = New System.Drawing.Point(166, 485)
        Me.Btn_OK.Name = "Btn_OK"
        Me.Btn_OK.Size = New System.Drawing.Size(62, 26)
        Me.Btn_OK.TabIndex = 416
        Me.Btn_OK.Text = "OK"
        Me.Btn_OK.UseVisualStyleBackColor = False
        '
        'Btn_Cancel
        '
        Me.Btn_Cancel.BackColor = System.Drawing.Color.White
        Me.Btn_Cancel.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Btn_Cancel.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Btn_Cancel.Location = New System.Drawing.Point(241, 485)
        Me.Btn_Cancel.Name = "Btn_Cancel"
        Me.Btn_Cancel.Size = New System.Drawing.Size(62, 26)
        Me.Btn_Cancel.TabIndex = 415
        Me.Btn_Cancel.Text = "Cancel"
        Me.Btn_Cancel.UseVisualStyleBackColor = False
        '
        'CB_About
        '
        Me.CB_About.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CB_About.FormattingEnabled = True
        Me.CB_About.Location = New System.Drawing.Point(16, 487)
        Me.CB_About.Name = "CB_About"
        Me.CB_About.Size = New System.Drawing.Size(98, 25)
        Me.CB_About.TabIndex = 413
        Me.CB_About.Text = "SOFTEKO"
        '
        'GB_Sample
        '
        Me.GB_Sample.BackColor = System.Drawing.Color.White
        Me.GB_Sample.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.GB_Sample.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GB_Sample.Location = New System.Drawing.Point(15, 160)
        Me.GB_Sample.Name = "GB_Sample"
        Me.GB_Sample.Size = New System.Drawing.Size(288, 299)
        Me.GB_Sample.TabIndex = 414
        Me.GB_Sample.TabStop = False
        Me.GB_Sample.Text = "Sample Image"
        '
        'CustomGroupBox4
        '
        Me.CustomGroupBox4.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.CustomGroupBox4.Controls.Add(Me.CustomGroupBox6)
        Me.CustomGroupBox4.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CustomGroupBox4.Location = New System.Drawing.Point(15, 15)
        Me.CustomGroupBox4.Name = "CustomGroupBox4"
        Me.CustomGroupBox4.Size = New System.Drawing.Size(288, 132)
        Me.CustomGroupBox4.TabIndex = 412
        Me.CustomGroupBox4.TabStop = False
        Me.CustomGroupBox4.Text = "Data Validation Range"
        '
        'CustomGroupBox6
        '
        Me.CustomGroupBox6.BackColor = System.Drawing.Color.White
        Me.CustomGroupBox6.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.CustomGroupBox6.Controls.Add(Me.Selection_source)
        Me.CustomGroupBox6.Controls.Add(Me.Label1)
        Me.CustomGroupBox6.Controls.Add(Me.TB_src_rng)
        Me.CustomGroupBox6.Controls.Add(Me.CB_Source)
        Me.CustomGroupBox6.Location = New System.Drawing.Point(1, 22)
        Me.CustomGroupBox6.Name = "CustomGroupBox6"
        Me.CustomGroupBox6.Size = New System.Drawing.Size(287, 110)
        Me.CustomGroupBox6.TabIndex = 0
        Me.CustomGroupBox6.TabStop = False
        '
        'Selection_source
        '
        Me.Selection_source.BackColor = System.Drawing.Color.White
        Me.Selection_source.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Selection_source.Image = CType(resources.GetObject("Selection_source.Image"), System.Drawing.Image)
        Me.Selection_source.Location = New System.Drawing.Point(243, 72)
        Me.Selection_source.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Selection_source.Name = "Selection_source"
        Me.Selection_source.Size = New System.Drawing.Size(24, 25)
        Me.Selection_source.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.Selection_source.TabIndex = 404
        Me.Selection_source.TabStop = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(12, 46)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(116, 17)
        Me.Label1.TabIndex = 379
        Me.Label1.Text = "Define the range :"
        '
        'TB_src_rng
        '
        Me.TB_src_rng.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TB_src_rng.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.TB_src_rng.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TB_src_rng.Location = New System.Drawing.Point(12, 72)
        Me.TB_src_rng.Name = "TB_src_rng"
        Me.TB_src_rng.Size = New System.Drawing.Size(255, 25)
        Me.TB_src_rng.TabIndex = 403
        '
        'CB_Source
        '
        Me.CB_Source.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CB_Source.FormattingEnabled = True
        Me.CB_Source.Location = New System.Drawing.Point(12, 13)
        Me.CB_Source.Name = "CB_Source"
        Me.CB_Source.Size = New System.Drawing.Size(255, 25)
        Me.CB_Source.TabIndex = 378
        '
        'Form39_DropdownlistwithSearchOption
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.ClientSize = New System.Drawing.Size(326, 527)
        Me.Controls.Add(Me.GB_Sample)
        Me.Controls.Add(Me.Btn_OK)
        Me.Controls.Add(Me.Btn_Cancel)
        Me.Controls.Add(Me.CB_About)
        Me.Controls.Add(Me.CustomGroupBox4)
        Me.HelpButton = True
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Form39_DropdownlistwithSearchOption"
        Me.Text = "Drop-down List with Search Option"
        Me.CustomGroupBox4.ResumeLayout(False)
        Me.CustomGroupBox6.ResumeLayout(False)
        Me.CustomGroupBox6.PerformLayout()
        CType(Me.Selection_source, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents CustomGroupBox4 As CustomGroupBox
    Friend WithEvents CustomGroupBox6 As CustomGroupBox
    Friend WithEvents Selection_source As Windows.Forms.PictureBox
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents TB_src_rng As Windows.Forms.TextBox
    Friend WithEvents CB_Source As Windows.Forms.ComboBox
    Friend WithEvents GB_Sample As CustomGroupBox
    Friend WithEvents Btn_OK As Windows.Forms.Button
    Friend WithEvents Btn_Cancel As Windows.Forms.Button
    Friend WithEvents CB_About As Windows.Forms.ComboBox
End Class
