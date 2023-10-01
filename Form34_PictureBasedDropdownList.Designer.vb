<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form34_PictureBasedDropdownList
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form34_PictureBasedDropdownList))
        Me.Src_selection = New System.Windows.Forms.PictureBox()
        Me.Btn_OK = New System.Windows.Forms.Button()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.TB_src_rng = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Btn_Cancel = New System.Windows.Forms.Button()
        Me.Des_selection = New System.Windows.Forms.PictureBox()
        Me.TB_des_rng = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Info = New System.Windows.Forms.PictureBox()
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.CustomGroupBox2 = New VSTO_Addins.CustomGroupBox()
        CType(Me.Src_selection, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Des_selection, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Info, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Src_selection
        '
        Me.Src_selection.BackColor = System.Drawing.Color.White
        Me.Src_selection.Image = CType(resources.GetObject("Src_selection.Image"), System.Drawing.Image)
        Me.Src_selection.Location = New System.Drawing.Point(249, 44)
        Me.Src_selection.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Src_selection.Name = "Src_selection"
        Me.Src_selection.Size = New System.Drawing.Size(24, 23)
        Me.Src_selection.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.Src_selection.TabIndex = 203
        Me.Src_selection.TabStop = False
        '
        'Btn_OK
        '
        Me.Btn_OK.BackColor = System.Drawing.Color.White
        Me.Btn_OK.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Btn_OK.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Btn_OK.Location = New System.Drawing.Point(465, 165)
        Me.Btn_OK.Name = "Btn_OK"
        Me.Btn_OK.Size = New System.Drawing.Size(62, 26)
        Me.Btn_OK.TabIndex = 202
        Me.Btn_OK.Text = "OK"
        Me.Btn_OK.UseVisualStyleBackColor = False
        '
        'ComboBox1
        '
        Me.ComboBox1.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Location = New System.Drawing.Point(18, 165)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(90, 25)
        Me.ComboBox1.TabIndex = 198
        Me.ComboBox1.Text = "SOFTEKO"
        '
        'TB_src_rng
        '
        Me.TB_src_rng.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TB_src_rng.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.TB_src_rng.Font = New System.Drawing.Font("Segoe UI", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TB_src_rng.Location = New System.Drawing.Point(15, 43)
        Me.TB_src_rng.Name = "TB_src_rng"
        Me.TB_src_rng.Size = New System.Drawing.Size(259, 25)
        Me.TB_src_rng.TabIndex = 196
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(15, 15)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(98, 17)
        Me.Label1.TabIndex = 195
        Me.Label1.Text = "Source Range :"
        '
        'Btn_Cancel
        '
        Me.Btn_Cancel.BackColor = System.Drawing.Color.White
        Me.Btn_Cancel.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Btn_Cancel.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Btn_Cancel.Location = New System.Drawing.Point(543, 165)
        Me.Btn_Cancel.Name = "Btn_Cancel"
        Me.Btn_Cancel.Size = New System.Drawing.Size(62, 26)
        Me.Btn_Cancel.TabIndex = 201
        Me.Btn_Cancel.Text = "Cancel"
        Me.Btn_Cancel.UseVisualStyleBackColor = False
        '
        'Des_selection
        '
        Me.Des_selection.BackColor = System.Drawing.Color.White
        Me.Des_selection.Image = CType(resources.GetObject("Des_selection.Image"), System.Drawing.Image)
        Me.Des_selection.Location = New System.Drawing.Point(249, 114)
        Me.Des_selection.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Des_selection.Name = "Des_selection"
        Me.Des_selection.Size = New System.Drawing.Size(24, 23)
        Me.Des_selection.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.Des_selection.TabIndex = 206
        Me.Des_selection.TabStop = False
        '
        'TB_des_rng
        '
        Me.TB_des_rng.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TB_des_rng.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.TB_des_rng.Font = New System.Drawing.Font("Segoe UI", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TB_des_rng.Location = New System.Drawing.Point(15, 113)
        Me.TB_des_rng.Name = "TB_des_rng"
        Me.TB_des_rng.Size = New System.Drawing.Size(259, 25)
        Me.TB_des_rng.TabIndex = 205
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(15, 81)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(126, 17)
        Me.Label2.TabIndex = 204
        Me.Label2.Text = "Destination Range :"
        '
        'Info
        '
        Me.Info.Image = CType(resources.GetObject("Info.Image"), System.Drawing.Image)
        Me.Info.Location = New System.Drawing.Point(119, 15)
        Me.Info.Name = "Info"
        Me.Info.Size = New System.Drawing.Size(20, 20)
        Me.Info.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.Info.TabIndex = 400
        Me.Info.TabStop = False
        Me.ToolTip1.SetToolTip(Me.Info, "Please, select both of the columns that contain the data and the relevant images")
        '
        'PictureBox2
        '
        Me.PictureBox2.Image = CType(resources.GetObject("PictureBox2.Image"), System.Drawing.Image)
        Me.PictureBox2.Location = New System.Drawing.Point(147, 81)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(20, 20)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox2.TabIndex = 401
        Me.PictureBox2.TabStop = False
        Me.ToolTip1.SetToolTip(Me.PictureBox2, "Please, select 2 columns")
        '
        'CustomGroupBox2
        '
        Me.CustomGroupBox2.BackColor = System.Drawing.Color.White
        Me.CustomGroupBox2.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.CustomGroupBox2.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CustomGroupBox2.Location = New System.Drawing.Point(301, 15)
        Me.CustomGroupBox2.Name = "CustomGroupBox2"
        Me.CustomGroupBox2.Size = New System.Drawing.Size(304, 126)
        Me.CustomGroupBox2.TabIndex = 399
        Me.CustomGroupBox2.TabStop = False
        Me.CustomGroupBox2.Text = "Sample Image"
        '
        'Form34_PictureBasedDropdownList
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.ClientSize = New System.Drawing.Size(629, 213)
        Me.Controls.Add(Me.PictureBox2)
        Me.Controls.Add(Me.Info)
        Me.Controls.Add(Me.CustomGroupBox2)
        Me.Controls.Add(Me.Des_selection)
        Me.Controls.Add(Me.TB_des_rng)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Src_selection)
        Me.Controls.Add(Me.Btn_OK)
        Me.Controls.Add(Me.ComboBox1)
        Me.Controls.Add(Me.TB_src_rng)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Btn_Cancel)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Form34_PictureBasedDropdownList"
        Me.Text = "Picture Based Drop-down List"
        CType(Me.Src_selection, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Des_selection, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Info, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Src_selection As Windows.Forms.PictureBox
    Friend WithEvents Btn_OK As Windows.Forms.Button
    Friend WithEvents ComboBox1 As Windows.Forms.ComboBox
    Friend WithEvents TB_src_rng As Windows.Forms.TextBox
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents Btn_Cancel As Windows.Forms.Button
    Friend WithEvents Des_selection As Windows.Forms.PictureBox
    Friend WithEvents TB_des_rng As Windows.Forms.TextBox
    Friend WithEvents Label2 As Windows.Forms.Label
    Friend WithEvents CustomGroupBox2 As CustomGroupBox
    Friend WithEvents Info As Windows.Forms.PictureBox
    Friend WithEvents PictureBox2 As Windows.Forms.PictureBox
    Friend WithEvents ToolTip1 As Windows.Forms.ToolTip
End Class
