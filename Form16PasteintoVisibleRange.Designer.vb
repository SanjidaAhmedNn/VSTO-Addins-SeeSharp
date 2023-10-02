<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form16PasteintoVisibleRange
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form16PasteintoVisibleRange))
        Me.CB_keepFormat = New System.Windows.Forms.CheckBox()
        Me.btnOK = New System.Windows.Forms.Button()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.CB_copyWs = New System.Windows.Forms.CheckBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.txtSourceRange = New System.Windows.Forms.TextBox()
        Me.Selection = New System.Windows.Forms.PictureBox()
        Me.AutoSelection = New System.Windows.Forms.PictureBox()
        Me.CustomGroupBox5 = New VSTO_Addins.CustomGroupBox()
        Me.CustomPanel1 = New VSTO_Addins.CustomPanel()
        Me.destinationSelection = New System.Windows.Forms.PictureBox()
        Me.txtDestRange = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        CType(Me.Selection, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.AutoSelection, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.CustomGroupBox5.SuspendLayout()
        CType(Me.destinationSelection, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'CB_keepFormat
        '
        Me.CB_keepFormat.AutoSize = True
        Me.CB_keepFormat.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CB_keepFormat.Location = New System.Drawing.Point(19, 77)
        Me.CB_keepFormat.Name = "CB_keepFormat"
        Me.CB_keepFormat.Size = New System.Drawing.Size(122, 21)
        Me.CB_keepFormat.TabIndex = 184
        Me.CB_keepFormat.Text = "Keep formatting"
        Me.CB_keepFormat.UseVisualStyleBackColor = True
        '
        'btnOK
        '
        Me.btnOK.BackColor = System.Drawing.Color.White
        Me.btnOK.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnOK.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOK.Location = New System.Drawing.Point(300, 316)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(62, 26)
        Me.btnOK.TabIndex = 190
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = False
        '
        'ComboBox1
        '
        Me.ComboBox1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Location = New System.Drawing.Point(12, 316)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(90, 25)
        Me.ComboBox1.TabIndex = 186
        Me.ComboBox1.Text = "SOFTEKO"
        '
        'CB_copyWs
        '
        Me.CB_copyWs.AutoSize = True
        Me.CB_copyWs.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CB_copyWs.Location = New System.Drawing.Point(13, 286)
        Me.CB_copyWs.Name = "CB_copyWs"
        Me.CB_copyWs.Size = New System.Drawing.Size(257, 21)
        Me.CB_copyWs.TabIndex = 185
        Me.CB_copyWs.Text = "Create a copy of the original worksheet"
        Me.CB_copyWs.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(15, 17)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(123, 17)
        Me.Label1.TabIndex = 182
        Me.Label1.Text = "Data to be copied :"
        '
        'btnCancel
        '
        Me.btnCancel.BackColor = System.Drawing.Color.White
        Me.btnCancel.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.btnCancel.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCancel.Location = New System.Drawing.Point(374, 316)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(62, 26)
        Me.btnCancel.TabIndex = 189
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = False
        '
        'txtSourceRange
        '
        Me.txtSourceRange.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtSourceRange.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtSourceRange.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSourceRange.Location = New System.Drawing.Point(144, 15)
        Me.txtSourceRange.Name = "txtSourceRange"
        Me.txtSourceRange.Size = New System.Drawing.Size(292, 25)
        Me.txtSourceRange.TabIndex = 204
        '
        'Selection
        '
        Me.Selection.BackColor = System.Drawing.Color.White
        Me.Selection.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Selection.Image = CType(resources.GetObject("Selection.Image"), System.Drawing.Image)
        Me.Selection.Location = New System.Drawing.Point(412, 15)
        Me.Selection.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Selection.Name = "Selection"
        Me.Selection.Size = New System.Drawing.Size(24, 25)
        Me.Selection.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.Selection.TabIndex = 206
        Me.Selection.TabStop = False
        '
        'AutoSelection
        '
        Me.AutoSelection.BackColor = System.Drawing.Color.White
        Me.AutoSelection.Image = CType(resources.GetObject("AutoSelection.Image"), System.Drawing.Image)
        Me.AutoSelection.Location = New System.Drawing.Point(385, 16)
        Me.AutoSelection.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.AutoSelection.Name = "AutoSelection"
        Me.AutoSelection.Size = New System.Drawing.Size(24, 23)
        Me.AutoSelection.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.AutoSelection.TabIndex = 205
        Me.AutoSelection.TabStop = False
        '
        'CustomGroupBox5
        '
        Me.CustomGroupBox5.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.CustomGroupBox5.Controls.Add(Me.CustomPanel1)
        Me.CustomGroupBox5.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CustomGroupBox5.Location = New System.Drawing.Point(15, 110)
        Me.CustomGroupBox5.Name = "CustomGroupBox5"
        Me.CustomGroupBox5.Size = New System.Drawing.Size(421, 164)
        Me.CustomGroupBox5.TabIndex = 187
        Me.CustomGroupBox5.TabStop = False
        Me.CustomGroupBox5.Text = "Sample Image"
        '
        'CustomPanel1
        '
        Me.CustomPanel1.BackColor = System.Drawing.Color.White
        Me.CustomPanel1.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.CustomPanel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.CustomPanel1.BorderWidth = 1
        Me.CustomPanel1.Location = New System.Drawing.Point(1, 30)
        Me.CustomPanel1.Name = "CustomPanel1"
        Me.CustomPanel1.Size = New System.Drawing.Size(420, 134)
        Me.CustomPanel1.TabIndex = 0
        '
        'destinationSelection
        '
        Me.destinationSelection.BackColor = System.Drawing.Color.White
        Me.destinationSelection.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.destinationSelection.Image = CType(resources.GetObject("destinationSelection.Image"), System.Drawing.Image)
        Me.destinationSelection.Location = New System.Drawing.Point(412, 46)
        Me.destinationSelection.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.destinationSelection.Name = "destinationSelection"
        Me.destinationSelection.Size = New System.Drawing.Size(24, 25)
        Me.destinationSelection.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.destinationSelection.TabIndex = 209
        Me.destinationSelection.TabStop = False
        '
        'txtDestRange
        '
        Me.txtDestRange.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtDestRange.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDestRange.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDestRange.Location = New System.Drawing.Point(144, 46)
        Me.txtDestRange.Name = "txtDestRange"
        Me.txtDestRange.Size = New System.Drawing.Size(292, 25)
        Me.txtDestRange.TabIndex = 208
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(15, 49)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(126, 17)
        Me.Label2.TabIndex = 207
        Me.Label2.Text = "Destination Range :"
        '
        'Form16PasteintoVisibleRange
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.ClientSize = New System.Drawing.Size(457, 361)
        Me.Controls.Add(Me.destinationSelection)
        Me.Controls.Add(Me.Selection)
        Me.Controls.Add(Me.AutoSelection)
        Me.Controls.Add(Me.txtDestRange)
        Me.Controls.Add(Me.txtSourceRange)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.CB_keepFormat)
        Me.Controls.Add(Me.btnOK)
        Me.Controls.Add(Me.ComboBox1)
        Me.Controls.Add(Me.CB_copyWs)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.CustomGroupBox5)
        Me.Controls.Add(Me.btnCancel)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Form16PasteintoVisibleRange"
        Me.Text = "Paste into Visible Range"
        CType(Me.Selection, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.AutoSelection, System.ComponentModel.ISupportInitialize).EndInit()
        Me.CustomGroupBox5.ResumeLayout(False)
        CType(Me.destinationSelection, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents CustomPanel1 As CustomPanel
    Friend WithEvents CB_keepFormat As Windows.Forms.CheckBox
    Friend WithEvents btnOK As Windows.Forms.Button
    Friend WithEvents ComboBox1 As Windows.Forms.ComboBox
    Friend WithEvents CB_copyWs As Windows.Forms.CheckBox
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents CustomGroupBox5 As CustomGroupBox
    Friend WithEvents btnCancel As Windows.Forms.Button
    Friend WithEvents txtSourceRange As Windows.Forms.TextBox
    Friend WithEvents Selection As Windows.Forms.PictureBox
    Friend WithEvents AutoSelection As Windows.Forms.PictureBox
    Friend WithEvents destinationSelection As Windows.Forms.PictureBox
    Friend WithEvents txtDestRange As Windows.Forms.TextBox
    Friend WithEvents Label2 As Windows.Forms.Label
End Class
