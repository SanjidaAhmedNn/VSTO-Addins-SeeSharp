<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form13HideAllExceptSelectedRange
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form13HideAllExceptSelectedRange))
        Me.pctBoxSelectRange = New System.Windows.Forms.PictureBox()
        Me.btnOK = New System.Windows.Forms.Button()
        Me.comboDropDown = New System.Windows.Forms.ComboBox()
        Me.checkBoxCopyWorksheet = New System.Windows.Forms.CheckBox()
        Me.txtSourceRange = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        Me.CustomGroupBox6 = New VSTO_Addins.CustomGroupBox()
        Me.custPanExcpectedOutput = New VSTO_Addins.CustomPanel()
        Me.CustomGroupBox5 = New VSTO_Addins.CustomGroupBox()
        Me.custPanInputRange = New VSTO_Addins.CustomPanel()
        CType(Me.pctBoxSelectRange, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.CustomGroupBox6.SuspendLayout()
        Me.CustomGroupBox5.SuspendLayout()
        Me.SuspendLayout()
        '
        'pctBoxSelectRange
        '
        Me.pctBoxSelectRange.BackColor = System.Drawing.Color.White
        Me.pctBoxSelectRange.Image = CType(resources.GetObject("pctBoxSelectRange.Image"), System.Drawing.Image)
        Me.pctBoxSelectRange.Location = New System.Drawing.Point(482, 13)
        Me.pctBoxSelectRange.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.pctBoxSelectRange.Name = "pctBoxSelectRange"
        Me.pctBoxSelectRange.Size = New System.Drawing.Size(24, 23)
        Me.pctBoxSelectRange.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.pctBoxSelectRange.TabIndex = 191
        Me.pctBoxSelectRange.TabStop = False
        '
        'btnOK
        '
        Me.btnOK.BackColor = System.Drawing.Color.White
        Me.btnOK.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnOK.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOK.Location = New System.Drawing.Point(370, 234)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(62, 26)
        Me.btnOK.TabIndex = 190
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = False
        '
        'comboDropDown
        '
        Me.comboDropDown.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.comboDropDown.FormattingEnabled = True
        Me.comboDropDown.Location = New System.Drawing.Point(15, 236)
        Me.comboDropDown.Name = "comboDropDown"
        Me.comboDropDown.Size = New System.Drawing.Size(90, 25)
        Me.comboDropDown.TabIndex = 186
        Me.comboDropDown.Text = "SOFTEKO"
        '
        'checkBoxCopyWorksheet
        '
        Me.checkBoxCopyWorksheet.AutoSize = True
        Me.checkBoxCopyWorksheet.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.checkBoxCopyWorksheet.Location = New System.Drawing.Point(15, 206)
        Me.checkBoxCopyWorksheet.Name = "checkBoxCopyWorksheet"
        Me.checkBoxCopyWorksheet.Size = New System.Drawing.Size(257, 21)
        Me.checkBoxCopyWorksheet.TabIndex = 185
        Me.checkBoxCopyWorksheet.Text = "Create a copy of the original worksheet"
        Me.checkBoxCopyWorksheet.UseVisualStyleBackColor = True
        '
        'txtSourceRange
        '
        Me.txtSourceRange.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtSourceRange.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtSourceRange.Font = New System.Drawing.Font("Segoe UI", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSourceRange.Location = New System.Drawing.Point(119, 12)
        Me.txtSourceRange.Name = "txtSourceRange"
        Me.txtSourceRange.Size = New System.Drawing.Size(388, 25)
        Me.txtSourceRange.TabIndex = 183
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(15, 15)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(98, 17)
        Me.Label1.TabIndex = 182
        Me.Label1.Text = "Source Range :"
        '
        'btnCancel
        '
        Me.btnCancel.BackColor = System.Drawing.Color.White
        Me.btnCancel.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnCancel.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCancel.Location = New System.Drawing.Point(447, 234)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(62, 26)
        Me.btnCancel.TabIndex = 189
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = False
        '
        'PictureBox2
        '
        Me.PictureBox2.Image = CType(resources.GetObject("PictureBox2.Image"), System.Drawing.Image)
        Me.PictureBox2.Location = New System.Drawing.Point(237, 113)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(50, 49)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox2.TabIndex = 194
        Me.PictureBox2.TabStop = False
        '
        'CustomGroupBox6
        '
        Me.CustomGroupBox6.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.CustomGroupBox6.Controls.Add(Me.custPanExcpectedOutput)
        Me.CustomGroupBox6.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CustomGroupBox6.Location = New System.Drawing.Point(287, 52)
        Me.CustomGroupBox6.Name = "CustomGroupBox6"
        Me.CustomGroupBox6.Size = New System.Drawing.Size(222, 146)
        Me.CustomGroupBox6.TabIndex = 188
        Me.CustomGroupBox6.TabStop = False
        Me.CustomGroupBox6.Text = "Expected Output"
        '
        'custPanExcpectedOutput
        '
        Me.custPanExcpectedOutput.BackColor = System.Drawing.Color.White
        Me.custPanExcpectedOutput.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.custPanExcpectedOutput.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.custPanExcpectedOutput.BorderWidth = 1
        Me.custPanExcpectedOutput.Location = New System.Drawing.Point(1, 30)
        Me.custPanExcpectedOutput.Name = "custPanExcpectedOutput"
        Me.custPanExcpectedOutput.Size = New System.Drawing.Size(220, 115)
        Me.custPanExcpectedOutput.TabIndex = 11
        '
        'CustomGroupBox5
        '
        Me.CustomGroupBox5.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.CustomGroupBox5.Controls.Add(Me.custPanInputRange)
        Me.CustomGroupBox5.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CustomGroupBox5.Location = New System.Drawing.Point(15, 52)
        Me.CustomGroupBox5.Name = "CustomGroupBox5"
        Me.CustomGroupBox5.Size = New System.Drawing.Size(222, 146)
        Me.CustomGroupBox5.TabIndex = 187
        Me.CustomGroupBox5.TabStop = False
        Me.CustomGroupBox5.Text = "Input Range"
        '
        'custPanInputRange
        '
        Me.custPanInputRange.BackColor = System.Drawing.Color.White
        Me.custPanInputRange.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.custPanInputRange.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.custPanInputRange.BorderWidth = 1
        Me.custPanInputRange.Location = New System.Drawing.Point(1, 30)
        Me.custPanInputRange.Name = "custPanInputRange"
        Me.custPanInputRange.Size = New System.Drawing.Size(220, 115)
        Me.custPanInputRange.TabIndex = 0
        '
        'Form13HideAllExceptSelectedRange
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.ClientSize = New System.Drawing.Size(526, 279)
        Me.Controls.Add(Me.PictureBox2)
        Me.Controls.Add(Me.pctBoxSelectRange)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.comboDropDown)
        Me.Controls.Add(Me.checkBoxCopyWorksheet)
        Me.Controls.Add(Me.txtSourceRange)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.CustomGroupBox6)
        Me.Controls.Add(Me.CustomGroupBox5)
        Me.Controls.Add(Me.btnCancel)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Form13HideAllExceptSelectedRange"
        Me.Text = " "
        CType(Me.pctBoxSelectRange, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.CustomGroupBox6.ResumeLayout(False)
        Me.CustomGroupBox5.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents custPanInputRange As CustomPanel
    Friend WithEvents custPanExcpectedOutput As CustomPanel
    Friend WithEvents pctBoxSelectRange As Windows.Forms.PictureBox
    Friend WithEvents btnOK As Windows.Forms.Button
    Friend WithEvents comboDropDown As Windows.Forms.ComboBox
    Friend WithEvents checkBoxCopyWorksheet As Windows.Forms.CheckBox
    Friend WithEvents txtSourceRange As Windows.Forms.TextBox
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents CustomGroupBox6 As CustomGroupBox
    Friend WithEvents CustomGroupBox5 As CustomGroupBox
    Friend WithEvents btnCancel As Windows.Forms.Button
    Friend WithEvents PictureBox2 As Windows.Forms.PictureBox
End Class
