<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form29_Simple_Drop_down_List
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form29_Simple_Drop_down_List))
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Selection = New System.Windows.Forms.PictureBox()
        Me.TB_dest_range = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.ListBox2 = New System.Windows.Forms.ListBox()
        Me.ComboBox2 = New System.Windows.Forms.ComboBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Btn_OK = New System.Windows.Forms.Button()
        Me.Btn_Cancel = New System.Windows.Forms.Button()
        Me.CustomGroupBox2 = New VSTO_Addins.CustomGroupBox()
        Me.CustomGroupBox1 = New VSTO_Addins.CustomGroupBox()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.ListBox1 = New System.Windows.Forms.ListBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Selection_Source = New System.Windows.Forms.PictureBox()
        Me.Info = New System.Windows.Forms.PictureBox()
        Me.TB_src_range = New System.Windows.Forms.TextBox()
        CType(Me.Selection, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.CustomGroupBox1.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Selection_Source, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Info, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(15, 15)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(126, 17)
        Me.Label1.TabIndex = 342
        Me.Label1.Text = "Destination Range :"
        '
        'Selection
        '
        Me.Selection.BackColor = System.Drawing.Color.White
        Me.Selection.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Selection.Image = CType(resources.GetObject("Selection.Image"), System.Drawing.Image)
        Me.Selection.Location = New System.Drawing.Point(253, 43)
        Me.Selection.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Selection.Name = "Selection"
        Me.Selection.Size = New System.Drawing.Size(24, 25)
        Me.Selection.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.Selection.TabIndex = 344
        Me.Selection.TabStop = False
        '
        'TB_dest_range
        '
        Me.TB_dest_range.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TB_dest_range.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.TB_dest_range.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TB_dest_range.Location = New System.Drawing.Point(15, 43)
        Me.TB_dest_range.Name = "TB_dest_range"
        Me.TB_dest_range.Size = New System.Drawing.Size(262, 25)
        Me.TB_dest_range.TabIndex = 343
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(15, 348)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(141, 17)
        Me.Label5.TabIndex = 347
        Me.Label5.Text = "Total items in the list :"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(15, 375)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(86, 17)
        Me.Label6.TabIndex = 348
        Me.Label6.Text = "List Preview :"
        '
        'ListBox2
        '
        Me.ListBox2.BackColor = System.Drawing.SystemColors.Window
        Me.ListBox2.Font = New System.Drawing.Font("Segoe UI", 9.38!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ListBox2.FormattingEnabled = True
        Me.ListBox2.ItemHeight = 17
        Me.ListBox2.Location = New System.Drawing.Point(15, 401)
        Me.ListBox2.Name = "ListBox2"
        Me.ListBox2.Size = New System.Drawing.Size(260, 89)
        Me.ListBox2.TabIndex = 349
        '
        'ComboBox2
        '
        Me.ComboBox2.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBox2.FormattingEnabled = True
        Me.ComboBox2.Location = New System.Drawing.Point(15, 509)
        Me.ComboBox2.Name = "ComboBox2"
        Me.ComboBox2.Size = New System.Drawing.Size(154, 25)
        Me.ComboBox2.TabIndex = 350
        Me.ComboBox2.Text = "SOFTEKO"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.BackColor = System.Drawing.SystemColors.Control
        Me.Label7.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Label7.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(158, 346)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(46, 17)
        Me.Label7.TabIndex = 352
        Me.Label7.Text = "Label7"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.Label7.Visible = False
        '
        'Btn_OK
        '
        Me.Btn_OK.BackColor = System.Drawing.Color.White
        Me.Btn_OK.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Btn_OK.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Btn_OK.Location = New System.Drawing.Point(485, 507)
        Me.Btn_OK.Name = "Btn_OK"
        Me.Btn_OK.Size = New System.Drawing.Size(62, 26)
        Me.Btn_OK.TabIndex = 354
        Me.Btn_OK.Text = "OK"
        Me.Btn_OK.UseVisualStyleBackColor = False
        '
        'Btn_Cancel
        '
        Me.Btn_Cancel.BackColor = System.Drawing.Color.White
        Me.Btn_Cancel.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Btn_Cancel.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Btn_Cancel.Location = New System.Drawing.Point(562, 507)
        Me.Btn_Cancel.Name = "Btn_Cancel"
        Me.Btn_Cancel.Size = New System.Drawing.Size(62, 26)
        Me.Btn_Cancel.TabIndex = 353
        Me.Btn_Cancel.Text = "Cancel"
        Me.Btn_Cancel.UseVisualStyleBackColor = False
        '
        'CustomGroupBox2
        '
        Me.CustomGroupBox2.BackColor = System.Drawing.Color.White
        Me.CustomGroupBox2.BorderColor = System.Drawing.Color.FromArgb(CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer), CType(CType(191, Byte), Integer))
        Me.CustomGroupBox2.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CustomGroupBox2.Location = New System.Drawing.Point(317, 15)
        Me.CustomGroupBox2.Name = "CustomGroupBox2"
        Me.CustomGroupBox2.Size = New System.Drawing.Size(307, 455)
        Me.CustomGroupBox2.TabIndex = 351
        Me.CustomGroupBox2.TabStop = False
        Me.CustomGroupBox2.Text = "Sample Image"
        '
        'CustomGroupBox1
        '
        Me.CustomGroupBox1.BackColor = System.Drawing.Color.White
        Me.CustomGroupBox1.BorderColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.CustomGroupBox1.Controls.Add(Me.PictureBox1)
        Me.CustomGroupBox1.Controls.Add(Me.ComboBox1)
        Me.CustomGroupBox1.Controls.Add(Me.Label4)
        Me.CustomGroupBox1.Controls.Add(Me.ListBox1)
        Me.CustomGroupBox1.Controls.Add(Me.Label3)
        Me.CustomGroupBox1.Controls.Add(Me.Label2)
        Me.CustomGroupBox1.Controls.Add(Me.Selection_Source)
        Me.CustomGroupBox1.Controls.Add(Me.Info)
        Me.CustomGroupBox1.Controls.Add(Me.TB_src_range)
        Me.CustomGroupBox1.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CustomGroupBox1.Location = New System.Drawing.Point(15, 88)
        Me.CustomGroupBox1.Name = "CustomGroupBox1"
        Me.CustomGroupBox1.Size = New System.Drawing.Size(260, 253)
        Me.CustomGroupBox1.TabIndex = 346
        Me.CustomGroupBox1.TabStop = False
        Me.CustomGroupBox1.Text = "Source Range"
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(94, 186)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(20, 20)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 352
        Me.PictureBox1.TabStop = False
        '
        'ComboBox1
        '
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Location = New System.Drawing.Point(14, 213)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(230, 25)
        Me.ComboBox1.TabIndex = 351
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(14, 188)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(77, 17)
        Me.Label4.TabIndex = 350
        Me.Label4.Text = "Other lists :"
        '
        'ListBox1
        '
        Me.ListBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.ListBox1.ColumnWidth = 10
        Me.ListBox1.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.ImeMode = System.Windows.Forms.ImeMode.Off
        Me.ListBox1.ItemHeight = 15
        Me.ListBox1.Items.AddRange(New Object() {"Sunday,Monday,Tuesday,Wednesday,Thursday,Friday,Saturday", "Sun,Mon,Tue,Wed,Thu,Fri,Sat", "January,February,March,April,May,June,July,August,September,October,November,Dece" &
                "mber", "Jan,Feb,Mar,Apr,May,Jun,July,Aug,Sep,Oct,Nov,Dec", "1,2,3,4,5,6,7,8,9,10", "I,II,III,IV,V,VI,VII,VIII,IX,X", "One,Two,Three,Four,Five,Six,Seven,Eight,Nine,Ten", "a,b,c ,d,e,f,g,h,i,j"})
        Me.ListBox1.Location = New System.Drawing.Point(15, 102)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(230, 77)
        Me.ListBox1.TabIndex = 349
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(11, 80)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(107, 17)
        Me.Label3.TabIndex = 348
        Me.Label3.Text = "Predefined lists :"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(11, 21)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(89, 17)
        Me.Label2.TabIndex = 345
        Me.Label2.Text = "Enter Range :"
        '
        'Selection_Source
        '
        Me.Selection_Source.BackColor = System.Drawing.Color.White
        Me.Selection_Source.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Selection_Source.Image = CType(resources.GetObject("Selection_Source.Image"), System.Drawing.Image)
        Me.Selection_Source.Location = New System.Drawing.Point(220, 47)
        Me.Selection_Source.Margin = New System.Windows.Forms.Padding(4, 3, 4, 3)
        Me.Selection_Source.Name = "Selection_Source"
        Me.Selection_Source.Size = New System.Drawing.Size(24, 25)
        Me.Selection_Source.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.Selection_Source.TabIndex = 347
        Me.Selection_Source.TabStop = False
        '
        'Info
        '
        Me.Info.Image = CType(resources.GetObject("Info.Image"), System.Drawing.Image)
        Me.Info.Location = New System.Drawing.Point(103, 20)
        Me.Info.Name = "Info"
        Me.Info.Size = New System.Drawing.Size(20, 20)
        Me.Info.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.Info.TabIndex = 345
        Me.Info.TabStop = False
        '
        'TB_src_range
        '
        Me.TB_src_range.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TB_src_range.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.TB_src_range.Font = New System.Drawing.Font("Segoe UI", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TB_src_range.Location = New System.Drawing.Point(14, 47)
        Me.TB_src_range.Name = "TB_src_range"
        Me.TB_src_range.Size = New System.Drawing.Size(230, 25)
        Me.TB_src_range.TabIndex = 346
        '
        'Form29_Simple_Drop_down_List
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(650, 553)
        Me.Controls.Add(Me.Btn_OK)
        Me.Controls.Add(Me.Btn_Cancel)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.CustomGroupBox2)
        Me.Controls.Add(Me.ComboBox2)
        Me.Controls.Add(Me.ListBox2)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.CustomGroupBox1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Selection)
        Me.Controls.Add(Me.TB_dest_range)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Form29_Simple_Drop_down_List"
        Me.Text = "Simple Drop-down List"
        CType(Me.Selection, System.ComponentModel.ISupportInitialize).EndInit()
        Me.CustomGroupBox1.ResumeLayout(False)
        Me.CustomGroupBox1.PerformLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Selection_Source, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Info, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents Selection As Windows.Forms.PictureBox
    Friend WithEvents TB_dest_range As Windows.Forms.TextBox
    Friend WithEvents CustomGroupBox1 As CustomGroupBox
    Friend WithEvents ComboBox1 As Windows.Forms.ComboBox
    Friend WithEvents Label4 As Windows.Forms.Label
    Friend WithEvents Label3 As Windows.Forms.Label
    Friend WithEvents Label2 As Windows.Forms.Label
    Friend WithEvents Selection_Source As Windows.Forms.PictureBox
    Friend WithEvents Info As Windows.Forms.PictureBox
    Friend WithEvents TB_src_range As Windows.Forms.TextBox
    Friend WithEvents Label5 As Windows.Forms.Label
    Friend WithEvents Label6 As Windows.Forms.Label
    Friend WithEvents ListBox2 As Windows.Forms.ListBox
    Friend WithEvents ListBox1 As Windows.Forms.ListBox
    Friend WithEvents ComboBox2 As Windows.Forms.ComboBox
    Friend WithEvents CustomGroupBox2 As CustomGroupBox
    Friend WithEvents Label7 As Windows.Forms.Label
    Friend WithEvents Btn_OK As Windows.Forms.Button
    Friend WithEvents Btn_Cancel As Windows.Forms.Button
    Friend WithEvents PictureBox1 As Windows.Forms.PictureBox
End Class
