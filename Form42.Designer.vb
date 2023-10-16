<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form42
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
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Btn_OK = New System.Windows.Forms.Button()
        Me.Btn_Cancel = New System.Windows.Forms.Button()
        Me.CB_About = New System.Windows.Forms.ComboBox()
        Me.CustomGroupBox3 = New VSTO_Addins.CustomGroupBox()
        Me.CGB = New VSTO_Addins.CustomGroupBox()
        Me.RB_Simple = New System.Windows.Forms.RadioButton()
        Me.RB_Dynamic = New System.Windows.Forms.RadioButton()
        Me.RB_No = New System.Windows.Forms.RadioButton()
        Me.RB_Yes = New System.Windows.Forms.RadioButton()
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        Me.CustomGroupBox3.SuspendLayout()
        Me.CGB.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label1.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Label1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(12, 13)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(285, 63)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Your current selection does not contain any Data Validation List.         " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Do yo" &
    "u want to create a Data Validation List?"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Btn_OK
        '
        Me.Btn_OK.BackColor = System.Drawing.Color.White
        Me.Btn_OK.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Btn_OK.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Btn_OK.Location = New System.Drawing.Point(167, 248)
        Me.Btn_OK.Name = "Btn_OK"
        Me.Btn_OK.Size = New System.Drawing.Size(62, 26)
        Me.Btn_OK.TabIndex = 424
        Me.Btn_OK.Text = "OK"
        Me.Btn_OK.UseVisualStyleBackColor = False
        '
        'Btn_Cancel
        '
        Me.Btn_Cancel.BackColor = System.Drawing.Color.White
        Me.Btn_Cancel.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Btn_Cancel.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Btn_Cancel.Location = New System.Drawing.Point(235, 248)
        Me.Btn_Cancel.Name = "Btn_Cancel"
        Me.Btn_Cancel.Size = New System.Drawing.Size(62, 26)
        Me.Btn_Cancel.TabIndex = 423
        Me.Btn_Cancel.Text = "Cancel"
        Me.Btn_Cancel.UseVisualStyleBackColor = False
        '
        'CB_About
        '
        Me.CB_About.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CB_About.FormattingEnabled = True
        Me.CB_About.Location = New System.Drawing.Point(12, 248)
        Me.CB_About.Name = "CB_About"
        Me.CB_About.Size = New System.Drawing.Size(98, 25)
        Me.CB_About.TabIndex = 422
        Me.CB_About.Text = "SOFTEKO"
        '
        'CustomGroupBox3
        '
        Me.CustomGroupBox3.BackColor = System.Drawing.Color.White
        Me.CustomGroupBox3.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.CustomGroupBox3.Controls.Add(Me.CGB)
        Me.CustomGroupBox3.Controls.Add(Me.RB_No)
        Me.CustomGroupBox3.Controls.Add(Me.RB_Yes)
        Me.CustomGroupBox3.Location = New System.Drawing.Point(12, 90)
        Me.CustomGroupBox3.Name = "CustomGroupBox3"
        Me.CustomGroupBox3.Size = New System.Drawing.Size(285, 119)
        Me.CustomGroupBox3.TabIndex = 427
        Me.CustomGroupBox3.TabStop = False
        '
        'CGB
        '
        Me.CGB.BorderColor = System.Drawing.Color.White
        Me.CGB.Controls.Add(Me.RB_Simple)
        Me.CGB.Controls.Add(Me.RB_Dynamic)
        Me.CGB.Location = New System.Drawing.Point(25, 32)
        Me.CGB.Name = "CGB"
        Me.CGB.Size = New System.Drawing.Size(194, 53)
        Me.CGB.TabIndex = 431
        Me.CGB.TabStop = False
        '
        'RB_Simple
        '
        Me.RB_Simple.AutoSize = True
        Me.RB_Simple.Checked = True
        Me.RB_Simple.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RB_Simple.Location = New System.Drawing.Point(6, 3)
        Me.RB_Simple.Name = "RB_Simple"
        Me.RB_Simple.Size = New System.Drawing.Size(155, 21)
        Me.RB_Simple.TabIndex = 429
        Me.RB_Simple.TabStop = True
        Me.RB_Simple.Text = "Simple drop-down list"
        Me.RB_Simple.UseVisualStyleBackColor = True
        '
        'RB_Dynamic
        '
        Me.RB_Dynamic.AutoSize = True
        Me.RB_Dynamic.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RB_Dynamic.Location = New System.Drawing.Point(6, 29)
        Me.RB_Dynamic.Name = "RB_Dynamic"
        Me.RB_Dynamic.Size = New System.Drawing.Size(165, 21)
        Me.RB_Dynamic.TabIndex = 430
        Me.RB_Dynamic.Text = "Dynamic drop-down list"
        Me.RB_Dynamic.UseVisualStyleBackColor = True
        '
        'RB_No
        '
        Me.RB_No.AutoSize = True
        Me.RB_No.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RB_No.Location = New System.Drawing.Point(11, 87)
        Me.RB_No.Name = "RB_No"
        Me.RB_No.Size = New System.Drawing.Size(211, 21)
        Me.RB_No.TabIndex = 428
        Me.RB_No.Text = "No, I have a Data Validation List"
        Me.RB_No.UseVisualStyleBackColor = True
        '
        'RB_Yes
        '
        Me.RB_Yes.AutoSize = True
        Me.RB_Yes.Checked = True
        Me.RB_Yes.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RB_Yes.Location = New System.Drawing.Point(11, 11)
        Me.RB_Yes.Name = "RB_Yes"
        Me.RB_Yes.Size = New System.Drawing.Size(268, 21)
        Me.RB_Yes.TabIndex = 427
        Me.RB_Yes.TabStop = True
        Me.RB_Yes.Text = "Yes, I want to create a Data Validation List"
        Me.RB_Yes.UseVisualStyleBackColor = True
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckBox1.Location = New System.Drawing.Point(12, 215)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(253, 21)
        Me.CheckBox1.TabIndex = 428
        Me.CheckBox1.Text = "Don't show this for this current session"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'Form42
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.ClientSize = New System.Drawing.Size(318, 286)
        Me.Controls.Add(Me.CheckBox1)
        Me.Controls.Add(Me.CustomGroupBox3)
        Me.Controls.Add(Me.Btn_OK)
        Me.Controls.Add(Me.Btn_Cancel)
        Me.Controls.Add(Me.CB_About)
        Me.Controls.Add(Me.Label1)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Form42"
        Me.Text = "Softeko for Excel"
        Me.CustomGroupBox3.ResumeLayout(False)
        Me.CustomGroupBox3.PerformLayout()
        Me.CGB.ResumeLayout(False)
        Me.CGB.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents Btn_OK As Windows.Forms.Button
    Friend WithEvents Btn_Cancel As Windows.Forms.Button
    Friend WithEvents CB_About As Windows.Forms.ComboBox
    Friend WithEvents CustomGroupBox3 As CustomGroupBox
    Friend WithEvents RB_Dynamic As Windows.Forms.RadioButton
    Friend WithEvents RB_Simple As Windows.Forms.RadioButton
    Friend WithEvents RB_No As Windows.Forms.RadioButton
    Friend WithEvents RB_Yes As Windows.Forms.RadioButton
    Friend WithEvents CGB As CustomGroupBox
    Friend WithEvents CheckBox1 As Windows.Forms.CheckBox
End Class
