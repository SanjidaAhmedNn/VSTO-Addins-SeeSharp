<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form14SpecifyScrollArea
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form14SpecifyScrollArea))
        Me.CP_Input_Range = New VSTO_Addins.CustomPanel()
        Me.CP_Output_Range = New VSTO_Addins.CustomPanel()
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        Me.Selection = New System.Windows.Forms.PictureBox()
        Me.Btn_OK = New System.Windows.Forms.Button()
        Me.ComboBox = New System.Windows.Forms.ComboBox()
        Me.CheckBox = New System.Windows.Forms.CheckBox()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.GB_ExpectedOutput = New VSTO_Addins.CustomGroupBox()
        Me.GB_InputRange = New VSTO_Addins.CustomGroupBox()
        Me.Btn_Cancel = New System.Windows.Forms.Button()
        Me.Info = New System.Windows.Forms.PictureBox()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Selection, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GB_ExpectedOutput.SuspendLayout()
        Me.GB_InputRange.SuspendLayout()
        CType(Me.Info, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'CP_Input_Range
        '
        Me.CP_Input_Range.BackColor = System.Drawing.Color.White
        Me.CP_Input_Range.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.CP_Input_Range.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.CP_Input_Range.BorderWidth = 1
        Me.CP_Input_Range.Location = New System.Drawing.Point(1, 30)
        Me.CP_Input_Range.Name = "CP_Input_Range"
        Me.CP_Input_Range.Size = New System.Drawing.Size(220, 115)
        Me.CP_Input_Range.TabIndex = 0
        '
        'CP_Output_Range
        '
        Me.CP_Output_Range.BackColor = System.Drawing.Color.White
        Me.CP_Output_Range.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.CP_Output_Range.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.CP_Output_Range.BorderWidth = 1
        Me.CP_Output_Range.Location = New System.Drawing.Point(1, 30)
        Me.CP_Output_Range.Name = "CP_Output_Range"
        Me.CP_Output_Range.Size = New System.Drawing.Size(220, 115)
        Me.CP_Output_Range.TabIndex = 11
        '
        'PictureBox2
        '
        Me.PictureBox2.Image = CType(resources.GetObject("PictureBox2.Image"), System.Drawing.Image)
        Me.PictureBox2.Location = New System.Drawing.Point(238, 116)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(50, 49)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox2.TabIndex = 204
        Me.PictureBox2.TabStop = False
        '
        'Selection
        '
        Me.Selection.BackColor = System.Drawing.Color.White
        Me.Selection.Image = CType(resources.GetObject("Selection.Image"), System.Drawing.Image)
        Me.Selection.Location = New System.Drawing.Point(447, 16)
        Me.Selection.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Selection.Name = "Selection"
        Me.Selection.Size = New System.Drawing.Size(24, 23)
        Me.Selection.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.Selection.TabIndex = 203
        Me.Selection.TabStop = False
        '
        'Btn_OK
        '
        Me.Btn_OK.BackColor = System.Drawing.Color.White
        Me.Btn_OK.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Btn_OK.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Btn_OK.Location = New System.Drawing.Point(371, 237)
        Me.Btn_OK.Name = "Btn_OK"
        Me.Btn_OK.Size = New System.Drawing.Size(62, 26)
        Me.Btn_OK.TabIndex = 202
        Me.Btn_OK.Text = "OK"
        Me.Btn_OK.UseVisualStyleBackColor = False
        '
        'ComboBox
        '
        Me.ComboBox.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBox.FormattingEnabled = True
        Me.ComboBox.Location = New System.Drawing.Point(15, 239)
        Me.ComboBox.Name = "ComboBox"
        Me.ComboBox.Size = New System.Drawing.Size(90, 25)
        Me.ComboBox.TabIndex = 198
        Me.ComboBox.Text = "SOFTEKO"
        '
        'CheckBox
        '
        Me.CheckBox.AutoSize = True
        Me.CheckBox.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CheckBox.Location = New System.Drawing.Point(15, 209)
        Me.CheckBox.Name = "CheckBox"
        Me.CheckBox.Size = New System.Drawing.Size(257, 21)
        Me.CheckBox.TabIndex = 197
        Me.CheckBox.Text = "Create a copy of the original worksheet"
        Me.CheckBox.UseVisualStyleBackColor = True
        '
        'TextBox1
        '
        Me.TextBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TextBox1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.TextBox1.Font = New System.Drawing.Font("Segoe UI", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox1.Location = New System.Drawing.Point(120, 15)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(352, 25)
        Me.TextBox1.TabIndex = 196
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(15, 18)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(98, 17)
        Me.Label1.TabIndex = 195
        Me.Label1.Text = "Source Range :"
        '
        'GB_ExpectedOutput
        '
        Me.GB_ExpectedOutput.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.GB_ExpectedOutput.Controls.Add(Me.CP_Output_Range)
        Me.GB_ExpectedOutput.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GB_ExpectedOutput.Location = New System.Drawing.Point(288, 55)
        Me.GB_ExpectedOutput.Name = "GB_ExpectedOutput"
        Me.GB_ExpectedOutput.Size = New System.Drawing.Size(222, 146)
        Me.GB_ExpectedOutput.TabIndex = 200
        Me.GB_ExpectedOutput.TabStop = False
        Me.GB_ExpectedOutput.Text = "Expected Output"
        '
        'GB_InputRange
        '
        Me.GB_InputRange.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.GB_InputRange.Controls.Add(Me.CP_Input_Range)
        Me.GB_InputRange.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GB_InputRange.Location = New System.Drawing.Point(15, 55)
        Me.GB_InputRange.Name = "GB_InputRange"
        Me.GB_InputRange.Size = New System.Drawing.Size(222, 146)
        Me.GB_InputRange.TabIndex = 199
        Me.GB_InputRange.TabStop = False
        Me.GB_InputRange.Text = "Input Range"
        '
        'Btn_Cancel
        '
        Me.Btn_Cancel.BackColor = System.Drawing.Color.White
        Me.Btn_Cancel.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Btn_Cancel.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Btn_Cancel.Location = New System.Drawing.Point(448, 237)
        Me.Btn_Cancel.Name = "Btn_Cancel"
        Me.Btn_Cancel.Size = New System.Drawing.Size(62, 26)
        Me.Btn_Cancel.TabIndex = 201
        Me.Btn_Cancel.Text = "Cancel"
        Me.Btn_Cancel.UseVisualStyleBackColor = False
        '
        'Info
        '
        Me.Info.Image = CType(resources.GetObject("Info.Image"), System.Drawing.Image)
        Me.Info.Location = New System.Drawing.Point(483, 14)
        Me.Info.Name = "Info"
        Me.Info.Size = New System.Drawing.Size(26, 26)
        Me.Info.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.Info.TabIndex = 205
        Me.Info.TabStop = False
        '
        'Form14SpecifyScrollArea
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(526, 279)
        Me.Controls.Add(Me.Info)
        Me.Controls.Add(Me.PictureBox2)
        Me.Controls.Add(Me.Selection)
        Me.Controls.Add(Me.Btn_OK)
        Me.Controls.Add(Me.ComboBox)
        Me.Controls.Add(Me.CheckBox)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.GB_ExpectedOutput)
        Me.Controls.Add(Me.GB_InputRange)
        Me.Controls.Add(Me.Btn_Cancel)
        Me.Name = "Form14SpecifyScrollArea"
        Me.Text = "Specify Scroll Area"
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Selection, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GB_ExpectedOutput.ResumeLayout(False)
        Me.GB_InputRange.ResumeLayout(False)
        CType(Me.Info, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents CP_Input_Range As CustomPanel
    Friend WithEvents CP_Output_Range As CustomPanel
    Friend WithEvents PictureBox2 As Windows.Forms.PictureBox
    Friend WithEvents Selection As Windows.Forms.PictureBox
    Friend WithEvents Btn_OK As Windows.Forms.Button
    Friend WithEvents ComboBox As Windows.Forms.ComboBox
    Friend WithEvents CheckBox As Windows.Forms.CheckBox
    Friend WithEvents TextBox1 As Windows.Forms.TextBox
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents GB_ExpectedOutput As CustomGroupBox
    Friend WithEvents GB_InputRange As CustomGroupBox
    Friend WithEvents Btn_Cancel As Windows.Forms.Button
    Friend WithEvents Info As Windows.Forms.PictureBox
End Class
