<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form5
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
        Me.CustomButton1 = New VSTO_Addins.CustomButton()
        Me.CustomGroupBox1 = New VSTO_Addins.CustomGroupBox()
        Me.SuspendLayout()
        '
        'CustomButton1
        '
        Me.CustomButton1.BackColor = System.Drawing.Color.White
        Me.CustomButton1.BorderColor = System.Drawing.Color.FromArgb(CType(CType(166, Byte), Integer), CType(CType(166, Byte), Integer), CType(CType(166, Byte), Integer))
        Me.CustomButton1.BorderThickness = 1
        Me.CustomButton1.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CustomButton1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(70, Byte), Integer), CType(CType(70, Byte), Integer), CType(CType(70, Byte), Integer))
        Me.CustomButton1.HoverColor = System.Drawing.Color.CornflowerBlue
        Me.CustomButton1.HoverTextColor = System.Drawing.Color.White
        Me.CustomButton1.Location = New System.Drawing.Point(39, 41)
        Me.CustomButton1.Name = "CustomButton1"
        Me.CustomButton1.PressedColor = System.Drawing.Color.RoyalBlue
        Me.CustomButton1.PressedTextColor = System.Drawing.Color.White
        Me.CustomButton1.Size = New System.Drawing.Size(155, 41)
        Me.CustomButton1.TabIndex = 0
        Me.CustomButton1.Text = "Check"
        Me.CustomButton1.UseVisualStyleBackColor = False
        '
        'CustomGroupBox1
        '
        Me.CustomGroupBox1.BorderColor = System.Drawing.SystemColors.Window
        Me.CustomGroupBox1.Location = New System.Drawing.Point(249, 110)
        Me.CustomGroupBox1.Name = "CustomGroupBox1"
        Me.CustomGroupBox1.Size = New System.Drawing.Size(8, 8)
        Me.CustomGroupBox1.TabIndex = 1
        Me.CustomGroupBox1.TabStop = False
        Me.CustomGroupBox1.Text = "CustomGroupBox1"
        '
        'Form5
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(245, 149)
        Me.Controls.Add(Me.CustomGroupBox1)
        Me.Controls.Add(Me.CustomButton1)
        Me.Name = "Form5"
        Me.Text = "Form5"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents CustomButton1 As CustomButton
    Friend WithEvents CustomGroupBox1 As CustomGroupBox
End Class
