﻿<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form31_Update_Dynamic_Drop_down_List
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form31_Update_Dynamic_Drop_down_List))
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.CustomGroupBox2 = New VSTO_Addins.CustomGroupBox()
        Me.CustomGroupBox10 = New VSTO_Addins.CustomGroupBox()
        Me.TextBox3 = New System.Windows.Forms.TextBox()
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        Me.PictureBox3 = New System.Windows.Forms.PictureBox()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.RadioButton10 = New System.Windows.Forms.RadioButton()
        Me.RadioButton3 = New System.Windows.Forms.RadioButton()
        Me.Btn_OK = New System.Windows.Forms.Button()
        Me.Btn_Cancel = New System.Windows.Forms.Button()
        Me.ComboBox2 = New System.Windows.Forms.ComboBox()
        Me.Selection_source = New System.Windows.Forms.PictureBox()
        Me.Info = New System.Windows.Forms.PictureBox()
        Me.CustomGroupBox2.SuspendLayout()
        Me.CustomGroupBox10.SuspendLayout()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Selection_source, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Info, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(15, 15)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(154, 17)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Updated Source Range :"
        '
        'TextBox1
        '
        Me.TextBox1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox1.Location = New System.Drawing.Point(15, 44)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(289, 25)
        Me.TextBox1.TabIndex = 1
        '
        'CustomGroupBox2
        '
        Me.CustomGroupBox2.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.CustomGroupBox2.Controls.Add(Me.CustomGroupBox10)
        Me.CustomGroupBox2.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CustomGroupBox2.Location = New System.Drawing.Point(15, 83)
        Me.CustomGroupBox2.Name = "CustomGroupBox2"
        Me.CustomGroupBox2.Size = New System.Drawing.Size(310, 170)
        Me.CustomGroupBox2.TabIndex = 270
        Me.CustomGroupBox2.TabStop = False
        Me.CustomGroupBox2.Text = "Destination Range"
        '
        'CustomGroupBox10
        '
        Me.CustomGroupBox10.BackColor = System.Drawing.Color.White
        Me.CustomGroupBox10.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.CustomGroupBox10.Controls.Add(Me.TextBox3)
        Me.CustomGroupBox10.Controls.Add(Me.PictureBox2)
        Me.CustomGroupBox10.Controls.Add(Me.PictureBox3)
        Me.CustomGroupBox10.Controls.Add(Me.TextBox2)
        Me.CustomGroupBox10.Controls.Add(Me.Label3)
        Me.CustomGroupBox10.Controls.Add(Me.RadioButton10)
        Me.CustomGroupBox10.Controls.Add(Me.RadioButton3)
        Me.CustomGroupBox10.Location = New System.Drawing.Point(1, 22)
        Me.CustomGroupBox10.Name = "CustomGroupBox10"
        Me.CustomGroupBox10.Size = New System.Drawing.Size(308, 148)
        Me.CustomGroupBox10.TabIndex = 0
        Me.CustomGroupBox10.TabStop = False
        '
        'TextBox3
        '
        Me.TextBox3.Location = New System.Drawing.Point(25, 33)
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.Size = New System.Drawing.Size(263, 25)
        Me.TextBox3.TabIndex = 209
        '
        'PictureBox2
        '
        Me.PictureBox2.Image = CType(resources.GetObject("PictureBox2.Image"), System.Drawing.Image)
        Me.PictureBox2.Location = New System.Drawing.Point(25, 88)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(14, 14)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox2.TabIndex = 208
        Me.PictureBox2.TabStop = False
        '
        'PictureBox3
        '
        Me.PictureBox3.BackColor = System.Drawing.Color.White
        Me.PictureBox3.Image = CType(resources.GetObject("PictureBox3.Image"), System.Drawing.Image)
        Me.PictureBox3.Location = New System.Drawing.Point(263, 115)
        Me.PictureBox3.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.PictureBox3.Name = "PictureBox3"
        Me.PictureBox3.Size = New System.Drawing.Size(24, 23)
        Me.PictureBox3.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox3.TabIndex = 207
        Me.PictureBox3.TabStop = False
        '
        'TextBox2
        '
        Me.TextBox2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TextBox2.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.TextBox2.Font = New System.Drawing.Font("Segoe UI", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox2.Location = New System.Drawing.Point(24, 114)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(263, 25)
        Me.TextBox2.TabIndex = 206
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(42, 85)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(109, 17)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "Select the range :"
        '
        'RadioButton10
        '
        Me.RadioButton10.AutoSize = True
        Me.RadioButton10.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadioButton10.Location = New System.Drawing.Point(8, 62)
        Me.RadioButton10.Name = "RadioButton10"
        Me.RadioButton10.Size = New System.Drawing.Size(185, 21)
        Me.RadioButton10.TabIndex = 1
        Me.RadioButton10.TabStop = True
        Me.RadioButton10.Text = "Store into a different range"
        Me.RadioButton10.UseVisualStyleBackColor = True
        '
        'RadioButton3
        '
        Me.RadioButton3.AutoSize = True
        Me.RadioButton3.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RadioButton3.Location = New System.Drawing.Point(8, 6)
        Me.RadioButton3.Name = "RadioButton3"
        Me.RadioButton3.Size = New System.Drawing.Size(178, 21)
        Me.RadioButton3.TabIndex = 0
        Me.RadioButton3.TabStop = True
        Me.RadioButton3.Text = "Same as the source range"
        Me.RadioButton3.UseVisualStyleBackColor = True
        '
        'Btn_OK
        '
        Me.Btn_OK.BackColor = System.Drawing.Color.White
        Me.Btn_OK.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Btn_OK.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Btn_OK.Location = New System.Drawing.Point(192, 274)
        Me.Btn_OK.Name = "Btn_OK"
        Me.Btn_OK.Size = New System.Drawing.Size(62, 26)
        Me.Btn_OK.TabIndex = 379
        Me.Btn_OK.Text = "OK"
        Me.Btn_OK.UseVisualStyleBackColor = False
        '
        'Btn_Cancel
        '
        Me.Btn_Cancel.BackColor = System.Drawing.Color.White
        Me.Btn_Cancel.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Btn_Cancel.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Btn_Cancel.Location = New System.Drawing.Point(263, 274)
        Me.Btn_Cancel.Name = "Btn_Cancel"
        Me.Btn_Cancel.Size = New System.Drawing.Size(62, 26)
        Me.Btn_Cancel.TabIndex = 378
        Me.Btn_Cancel.Text = "Cancel"
        Me.Btn_Cancel.UseVisualStyleBackColor = False
        '
        'ComboBox2
        '
        Me.ComboBox2.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBox2.FormattingEnabled = True
        Me.ComboBox2.Location = New System.Drawing.Point(16, 274)
        Me.ComboBox2.Name = "ComboBox2"
        Me.ComboBox2.Size = New System.Drawing.Size(96, 25)
        Me.ComboBox2.TabIndex = 377
        Me.ComboBox2.Text = "SOFTEKO"
        '
        'Selection_source
        '
        Me.Selection_source.BackColor = System.Drawing.Color.White
        Me.Selection_source.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Selection_source.Image = CType(resources.GetObject("Selection_source.Image"), System.Drawing.Image)
        Me.Selection_source.Location = New System.Drawing.Point(280, 44)
        Me.Selection_source.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Selection_source.Name = "Selection_source"
        Me.Selection_source.Size = New System.Drawing.Size(24, 25)
        Me.Selection_source.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.Selection_source.TabIndex = 381
        Me.Selection_source.TabStop = False
        '
        'Info
        '
        Me.Info.Image = CType(resources.GetObject("Info.Image"), System.Drawing.Image)
        Me.Info.Location = New System.Drawing.Point(308, 46)
        Me.Info.Name = "Info"
        Me.Info.Size = New System.Drawing.Size(20, 20)
        Me.Info.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.Info.TabIndex = 382
        Me.Info.TabStop = False
        '
        'Form31_Update_Dynamic_Drop_down_List
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(349, 322)
        Me.Controls.Add(Me.Info)
        Me.Controls.Add(Me.Selection_source)
        Me.Controls.Add(Me.Btn_OK)
        Me.Controls.Add(Me.Btn_Cancel)
        Me.Controls.Add(Me.ComboBox2)
        Me.Controls.Add(Me.CustomGroupBox2)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.Label1)
        Me.Name = "Form31_Update_Dynamic_Drop_down_List"
        Me.Text = "Update Dynamic Drop-down List"
        Me.CustomGroupBox2.ResumeLayout(False)
        Me.CustomGroupBox10.ResumeLayout(False)
        Me.CustomGroupBox10.PerformLayout()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Selection_source, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Info, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents TextBox1 As Windows.Forms.TextBox
    Friend WithEvents CustomGroupBox2 As CustomGroupBox
    Friend WithEvents CustomGroupBox10 As CustomGroupBox
    Friend WithEvents TextBox3 As Windows.Forms.TextBox
    Friend WithEvents PictureBox2 As Windows.Forms.PictureBox
    Friend WithEvents PictureBox3 As Windows.Forms.PictureBox
    Friend WithEvents TextBox2 As Windows.Forms.TextBox
    Friend WithEvents Label3 As Windows.Forms.Label
    Friend WithEvents RadioButton10 As Windows.Forms.RadioButton
    Friend WithEvents RadioButton3 As Windows.Forms.RadioButton
    Friend WithEvents Btn_OK As Windows.Forms.Button
    Friend WithEvents Btn_Cancel As Windows.Forms.Button
    Friend WithEvents ComboBox2 As Windows.Forms.ComboBox
    Friend WithEvents Selection_source As Windows.Forms.PictureBox
    Friend WithEvents Info As Windows.Forms.PictureBox
End Class