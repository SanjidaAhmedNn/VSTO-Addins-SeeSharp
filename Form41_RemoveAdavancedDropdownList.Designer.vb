<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form41_RemoveAdavancedDropdownList
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form41_RemoveAdavancedDropdownList))
        Me.Btn_OK = New System.Windows.Forms.Button()
        Me.Btn_Cancel = New System.Windows.Forms.Button()
        Me.CB_About = New System.Windows.Forms.ComboBox()
        Me.CustomGroupBox2 = New VSTO_Addins.CustomGroupBox()
        Me.CustomGroupBox1 = New VSTO_Addins.CustomGroupBox()
        Me.CB_search = New System.Windows.Forms.CheckBox()
        Me.CB_checkbox = New System.Windows.Forms.CheckBox()
        Me.CB_multiselect = New System.Windows.Forms.CheckBox()
        Me.CustomGroupBox4 = New VSTO_Addins.CustomGroupBox()
        Me.CustomGroupBox6 = New VSTO_Addins.CustomGroupBox()
        Me.Selection_source = New System.Windows.Forms.PictureBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TB_src_rng = New System.Windows.Forms.TextBox()
        Me.CB_Source = New System.Windows.Forms.ComboBox()
        Me.CustomGroupBox2.SuspendLayout()
        Me.CustomGroupBox1.SuspendLayout()
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
        Me.Btn_OK.Location = New System.Drawing.Point(164, 310)
        Me.Btn_OK.Name = "Btn_OK"
        Me.Btn_OK.Size = New System.Drawing.Size(62, 26)
        Me.Btn_OK.TabIndex = 421
        Me.Btn_OK.Text = "OK"
        Me.Btn_OK.UseVisualStyleBackColor = False
        '
        'Btn_Cancel
        '
        Me.Btn_Cancel.BackColor = System.Drawing.Color.White
        Me.Btn_Cancel.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Btn_Cancel.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Btn_Cancel.Location = New System.Drawing.Point(239, 310)
        Me.Btn_Cancel.Name = "Btn_Cancel"
        Me.Btn_Cancel.Size = New System.Drawing.Size(62, 26)
        Me.Btn_Cancel.TabIndex = 420
        Me.Btn_Cancel.Text = "Cancel"
        Me.Btn_Cancel.UseVisualStyleBackColor = False
        '
        'CB_About
        '
        Me.CB_About.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CB_About.FormattingEnabled = True
        Me.CB_About.Location = New System.Drawing.Point(14, 312)
        Me.CB_About.Name = "CB_About"
        Me.CB_About.Size = New System.Drawing.Size(98, 25)
        Me.CB_About.TabIndex = 418
        Me.CB_About.Text = "SOFTEKO"
        '
        'CustomGroupBox2
        '
        Me.CustomGroupBox2.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.CustomGroupBox2.Controls.Add(Me.CustomGroupBox1)
        Me.CustomGroupBox2.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CustomGroupBox2.Location = New System.Drawing.Point(13, 164)
        Me.CustomGroupBox2.Name = "CustomGroupBox2"
        Me.CustomGroupBox2.Size = New System.Drawing.Size(288, 128)
        Me.CustomGroupBox2.TabIndex = 422
        Me.CustomGroupBox2.TabStop = False
        Me.CustomGroupBox2.Text = "Data Validation List Type"
        '
        'CustomGroupBox1
        '
        Me.CustomGroupBox1.BackColor = System.Drawing.Color.White
        Me.CustomGroupBox1.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.CustomGroupBox1.Controls.Add(Me.CB_search)
        Me.CustomGroupBox1.Controls.Add(Me.CB_checkbox)
        Me.CustomGroupBox1.Controls.Add(Me.CB_multiselect)
        Me.CustomGroupBox1.Location = New System.Drawing.Point(1, 22)
        Me.CustomGroupBox1.Name = "CustomGroupBox1"
        Me.CustomGroupBox1.Size = New System.Drawing.Size(287, 105)
        Me.CustomGroupBox1.TabIndex = 0
        Me.CustomGroupBox1.TabStop = False
        '
        'CB_search
        '
        Me.CB_search.AutoSize = True
        Me.CB_search.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CB_search.Location = New System.Drawing.Point(10, 70)
        Me.CB_search.Name = "CB_search"
        Me.CB_search.Size = New System.Drawing.Size(234, 21)
        Me.CB_search.TabIndex = 2
        Me.CB_search.Text = "Drop-down list with search option"
        Me.CB_search.UseVisualStyleBackColor = True
        '
        'CB_checkbox
        '
        Me.CB_checkbox.AutoSize = True
        Me.CB_checkbox.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CB_checkbox.Location = New System.Drawing.Point(10, 41)
        Me.CB_checkbox.Name = "CB_checkbox"
        Me.CB_checkbox.Size = New System.Drawing.Size(250, 21)
        Me.CB_checkbox.TabIndex = 1
        Me.CB_checkbox.Text = "Drop-down list containing check box"
        Me.CB_checkbox.UseVisualStyleBackColor = True
        '
        'CB_multiselect
        '
        Me.CB_multiselect.AutoSize = True
        Me.CB_multiselect.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CB_multiselect.Location = New System.Drawing.Point(10, 12)
        Me.CB_multiselect.Name = "CB_multiselect"
        Me.CB_multiselect.Size = New System.Drawing.Size(249, 21)
        Me.CB_multiselect.TabIndex = 0
        Me.CB_multiselect.Text = "Multi selection-based drop-down list"
        Me.CB_multiselect.UseVisualStyleBackColor = True
        '
        'CustomGroupBox4
        '
        Me.CustomGroupBox4.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.CustomGroupBox4.Controls.Add(Me.CustomGroupBox6)
        Me.CustomGroupBox4.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CustomGroupBox4.Location = New System.Drawing.Point(12, 12)
        Me.CustomGroupBox4.Name = "CustomGroupBox4"
        Me.CustomGroupBox4.Size = New System.Drawing.Size(288, 132)
        Me.CustomGroupBox4.TabIndex = 417
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
        'Form41_RemoveAdavancedDropdownList
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.ClientSize = New System.Drawing.Size(321, 352)
        Me.Controls.Add(Me.CustomGroupBox2)
        Me.Controls.Add(Me.Btn_OK)
        Me.Controls.Add(Me.Btn_Cancel)
        Me.Controls.Add(Me.CB_About)
        Me.Controls.Add(Me.CustomGroupBox4)
        Me.HelpButton = True
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Form41_RemoveAdavancedDropdownList"
        Me.Text = "Remove Adavanced Dropdown List"
        Me.CustomGroupBox2.ResumeLayout(False)
        Me.CustomGroupBox1.ResumeLayout(False)
        Me.CustomGroupBox1.PerformLayout()
        Me.CustomGroupBox4.ResumeLayout(False)
        Me.CustomGroupBox6.ResumeLayout(False)
        Me.CustomGroupBox6.PerformLayout()
        CType(Me.Selection_source, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Btn_OK As Windows.Forms.Button
    Friend WithEvents Btn_Cancel As Windows.Forms.Button
    Friend WithEvents CB_About As Windows.Forms.ComboBox
    Friend WithEvents CustomGroupBox4 As CustomGroupBox
    Friend WithEvents CustomGroupBox6 As CustomGroupBox
    Friend WithEvents Selection_source As Windows.Forms.PictureBox
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents TB_src_rng As Windows.Forms.TextBox
    Friend WithEvents CB_Source As Windows.Forms.ComboBox
    Friend WithEvents CustomGroupBox1 As CustomGroupBox
    Friend WithEvents CB_search As Windows.Forms.CheckBox
    Friend WithEvents CB_checkbox As Windows.Forms.CheckBox
    Friend WithEvents CB_multiselect As Windows.Forms.CheckBox
    Friend WithEvents CustomGroupBox2 As CustomGroupBox
End Class
