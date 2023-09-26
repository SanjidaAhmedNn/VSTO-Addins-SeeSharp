<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form32_ExtendDropDownList
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form32_ExtendDropDownList))
        Me.Info = New System.Windows.Forms.PictureBox()
        Me.Source_selection = New System.Windows.Forms.PictureBox()
        Me.Btn_OK = New System.Windows.Forms.Button()
        Me.Btn_Cancel = New System.Windows.Forms.Button()
        Me.ComboBox2 = New System.Windows.Forms.ComboBox()
        Me.TB_src_rng = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Dest_selection = New System.Windows.Forms.PictureBox()
        Me.TB_des_rng = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.L_warning = New VSTO_Addins.CustomLabel()
        CType(Me.Info, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Source_selection, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Dest_selection, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Info
        '
        Me.Info.Image = CType(resources.GetObject("Info.Image"), System.Drawing.Image)
        Me.Info.Location = New System.Drawing.Point(220, 14)
        Me.Info.Name = "Info"
        Me.Info.Size = New System.Drawing.Size(20, 20)
        Me.Info.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.Info.TabIndex = 390
        Me.Info.TabStop = False
        '
        'Source_selection
        '
        Me.Source_selection.BackColor = System.Drawing.Color.White
        Me.Source_selection.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Source_selection.Image = CType(resources.GetObject("Source_selection.Image"), System.Drawing.Image)
        Me.Source_selection.Location = New System.Drawing.Point(309, 42)
        Me.Source_selection.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Source_selection.Name = "Source_selection"
        Me.Source_selection.Size = New System.Drawing.Size(24, 25)
        Me.Source_selection.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.Source_selection.TabIndex = 389
        Me.Source_selection.TabStop = False
        '
        'Btn_OK
        '
        Me.Btn_OK.BackColor = System.Drawing.Color.White
        Me.Btn_OK.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Btn_OK.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Btn_OK.Location = New System.Drawing.Point(195, 188)
        Me.Btn_OK.Name = "Btn_OK"
        Me.Btn_OK.Size = New System.Drawing.Size(62, 26)
        Me.Btn_OK.TabIndex = 388
        Me.Btn_OK.Text = "OK"
        Me.Btn_OK.UseVisualStyleBackColor = False
        '
        'Btn_Cancel
        '
        Me.Btn_Cancel.BackColor = System.Drawing.Color.White
        Me.Btn_Cancel.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Btn_Cancel.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Btn_Cancel.Location = New System.Drawing.Point(271, 188)
        Me.Btn_Cancel.Name = "Btn_Cancel"
        Me.Btn_Cancel.Size = New System.Drawing.Size(62, 26)
        Me.Btn_Cancel.TabIndex = 387
        Me.Btn_Cancel.Text = "Cancel"
        Me.Btn_Cancel.UseVisualStyleBackColor = False
        '
        'ComboBox2
        '
        Me.ComboBox2.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBox2.FormattingEnabled = True
        Me.ComboBox2.Location = New System.Drawing.Point(15, 189)
        Me.ComboBox2.Name = "ComboBox2"
        Me.ComboBox2.Size = New System.Drawing.Size(96, 25)
        Me.ComboBox2.TabIndex = 386
        Me.ComboBox2.Text = "SOFTEKO"
        '
        'TB_src_rng
        '
        Me.TB_src_rng.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TB_src_rng.Location = New System.Drawing.Point(15, 42)
        Me.TB_src_rng.Name = "TB_src_rng"
        Me.TB_src_rng.Size = New System.Drawing.Size(318, 25)
        Me.TB_src_rng.TabIndex = 384
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(15, 15)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(203, 17)
        Me.Label1.TabIndex = 383
        Me.Label1.Text = "Select Dynamic Drop-down List :"
        '
        'Dest_selection
        '
        Me.Dest_selection.BackColor = System.Drawing.Color.White
        Me.Dest_selection.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Dest_selection.Image = CType(resources.GetObject("Dest_selection.Image"), System.Drawing.Image)
        Me.Dest_selection.Location = New System.Drawing.Point(309, 107)
        Me.Dest_selection.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Dest_selection.Name = "Dest_selection"
        Me.Dest_selection.Size = New System.Drawing.Size(24, 25)
        Me.Dest_selection.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.Dest_selection.TabIndex = 393
        Me.Dest_selection.TabStop = False
        '
        'TB_des_rng
        '
        Me.TB_des_rng.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TB_des_rng.Location = New System.Drawing.Point(15, 107)
        Me.TB_des_rng.Name = "TB_des_rng"
        Me.TB_des_rng.Size = New System.Drawing.Size(318, 25)
        Me.TB_des_rng.TabIndex = 392
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(15, 78)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(326, 17)
        Me.Label2.TabIndex = 391
        Me.Label2.Text = "Select the expanded Dynamic drop-down list range :"
        '
        'L_warning
        '
        Me.L_warning.BorderColor = System.Drawing.Color.DimGray
        Me.L_warning.BorderWidth = 1
        Me.L_warning.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.L_warning.Location = New System.Drawing.Point(29, 147)
        Me.L_warning.Name = "L_warning"
        Me.L_warning.Size = New System.Drawing.Size(286, 23)
        Me.L_warning.TabIndex = 396
        Me.L_warning.Text = " These two ranges must intersect each other"
        Me.L_warning.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Form32_ExtendDropDownList
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(353, 234)
        Me.Controls.Add(Me.L_warning)
        Me.Controls.Add(Me.Dest_selection)
        Me.Controls.Add(Me.TB_des_rng)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Info)
        Me.Controls.Add(Me.Source_selection)
        Me.Controls.Add(Me.Btn_OK)
        Me.Controls.Add(Me.Btn_Cancel)
        Me.Controls.Add(Me.ComboBox2)
        Me.Controls.Add(Me.TB_src_rng)
        Me.Controls.Add(Me.Label1)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Form32_ExtendDropDownList"
        Me.Text = "Extend Drop-down List"
        CType(Me.Info, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Source_selection, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Dest_selection, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Info As Windows.Forms.PictureBox
    Friend WithEvents Source_selection As Windows.Forms.PictureBox
    Friend WithEvents Btn_OK As Windows.Forms.Button
    Friend WithEvents Btn_Cancel As Windows.Forms.Button
    Friend WithEvents ComboBox2 As Windows.Forms.ComboBox
    Friend WithEvents TB_src_rng As Windows.Forms.TextBox
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents Dest_selection As Windows.Forms.PictureBox
    Friend WithEvents TB_des_rng As Windows.Forms.TextBox
    Friend WithEvents Label2 As Windows.Forms.Label
    Friend WithEvents L_warning As CustomLabel
End Class
