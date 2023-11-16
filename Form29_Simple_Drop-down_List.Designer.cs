using System;

namespace VSTO_Addins
{
    [Microsoft.VisualBasic.CompilerServices.DesignerGenerated()]
    public partial class Form29_Simple_Drop_down_List : System.Windows.Forms.Form
    {

        // Form overrides dispose to clean up the component list.
        [System.Diagnostics.DebuggerNonUserCode()]
        protected override void Dispose(bool disposing)
        {
            try
            {
                if (disposing && components is not null)
                {
                    components.Dispose();
                }
            }
            finally
            {
                base.Dispose(disposing);
            }
        }

        // Required by the Windows Form Designer
        private System.ComponentModel.IContainer components;

        // NOTE: The following procedure is required by the Windows Form Designer
        // It can be modified using the Windows Form Designer.  
        // Do not modify it using the code editor.
        [System.Diagnostics.DebuggerStepThrough()]
        private void InitializeComponent()
        {
            components = new System.ComponentModel.Container();
            var resources = new System.ComponentModel.ComponentResourceManager(typeof(Form29_Simple_Drop_down_List));
            Label1 = new System.Windows.Forms.Label();
            Selection_destination = new System.Windows.Forms.PictureBox();
            Selection_destination.Click += new EventHandler(Selection_Click);
            Selection_destination.KeyDown += new System.Windows.Forms.KeyEventHandler(destination);
            TB_dest_range = new System.Windows.Forms.TextBox();
            TB_dest_range.TextChanged += new EventHandler(TB_dest_rane_TextChanged);
            TB_dest_range.KeyDown += new System.Windows.Forms.KeyEventHandler(TB_dest);
            TB_dest_range.KeyDown += new System.Windows.Forms.KeyEventHandler(TB_dest_range_Enter);
            Label5 = new System.Windows.Forms.Label();
            Label6 = new System.Windows.Forms.Label();
            List_Preview = new System.Windows.Forms.ListBox();
            List_Preview.KeyDown += new System.Windows.Forms.KeyEventHandler(Listboxx2);
            List_Preview.SelectedIndexChanged += new EventHandler(List_Preview_SelectedIndexChanged);
            ComboBox2 = new System.Windows.Forms.ComboBox();
            ComboBox2.SelectedIndexChanged += new EventHandler(ComboBox2_SelectedIndexChanged);
            ComboBox2.MouseLeave += new EventHandler(ComboBox2_MouseLeave);
            ComboBox2.KeyDown += new System.Windows.Forms.KeyEventHandler(ComboBox2_KeyDown);
            Label7 = new System.Windows.Forms.Label();
            Btn_OK = new System.Windows.Forms.Button();
            Btn_OK.Click += new EventHandler(Btn_OK_Click);
            Btn_Cancel = new System.Windows.Forms.Button();
            Btn_Cancel.Click += new EventHandler(Btn_Cancel_Click);
            ToolTip1 = new System.Windows.Forms.ToolTip(components);
            CustomGroupBox2 = new CustomGroupBox();
            CustomGroupBox1 = new CustomGroupBox();
            RadioButton3 = new System.Windows.Forms.RadioButton();
            RadioButton3.KeyDown += new System.Windows.Forms.KeyEventHandler(RB_3);
            RadioButton3.CheckedChanged += new EventHandler(RadioButton3_CheckedChanged);
            RadioButton2 = new System.Windows.Forms.RadioButton();
            RadioButton2.KeyDown += new System.Windows.Forms.KeyEventHandler(RB_2);
            RadioButton2.CheckedChanged += new EventHandler(RadioButton2_CheckedChanged);
            RadioButton1 = new System.Windows.Forms.RadioButton();
            RadioButton1.KeyDown += new System.Windows.Forms.KeyEventHandler(RB_1);
            RadioButton1.CheckedChanged += new EventHandler(RadioButton1_CheckedChanged);
            PictureBox1 = new System.Windows.Forms.PictureBox();
            ComboBox1 = new System.Windows.Forms.ComboBox();
            ComboBox1.MouseClick += new System.Windows.Forms.MouseEventHandler(ComboBox1_MouseClick);
            ComboBox1.Enter += new EventHandler(ComboBox1_Enter);
            ComboBox1.KeyPress += new System.Windows.Forms.KeyPressEventHandler(ComboBox1_KeyPress);
            ComboBox1.KeyDown += new System.Windows.Forms.KeyEventHandler(ComboBox1_KeyDown);
            ComboBox1.Leave += new EventHandler(ComboBox1_Leave);
            ComboBox1.SelectedValueChanged += new EventHandler(Selection);
            ComboBox1.TextUpdate += new EventHandler(ComboBox1_TextUpdate);
            ComboBox1.SelectedIndexChanged += new EventHandler(ComboBox1_SelectedIndexChanged);
            ComboBox1.SelectedValueChanged += new EventHandler(Selection);
            ListBox1 = new System.Windows.Forms.ListBox();
            ListBox1.SelectedIndexChanged += new EventHandler(ListBox1_SelectedIndexChanged);
            ListBox1.DrawItem += new System.Windows.Forms.DrawItemEventHandler(ListBox1_DrawItem);
            ListBox1.KeyDown += new System.Windows.Forms.KeyEventHandler(Listbox);
            Selection_Source = new System.Windows.Forms.PictureBox();
            Selection_Source.Click += new EventHandler(Selection_Source_Click);
            Selection_Source.KeyDown += new System.Windows.Forms.KeyEventHandler(source);
            Info = new System.Windows.Forms.PictureBox();
            Info.Click += new EventHandler(Info_Click);
            TB_src_range = new System.Windows.Forms.TextBox();
            TB_src_range.TextChanged += new EventHandler(TB_src_range_TextChanged);
            TB_src_range.KeyDown += new System.Windows.Forms.KeyEventHandler(TB_src_range_Enter);
            ((System.ComponentModel.ISupportInitialize)Selection_destination).BeginInit();
            CustomGroupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)PictureBox1).BeginInit();
            ((System.ComponentModel.ISupportInitialize)Selection_Source).BeginInit();
            ((System.ComponentModel.ISupportInitialize)Info).BeginInit();
            SuspendLayout();
            // 
            // Label1
            // 
            Label1.AutoSize = true;
            Label1.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            Label1.Location = new System.Drawing.Point(15, 15);
            Label1.Name = "Label1";
            Label1.Size = new System.Drawing.Size(122, 17);
            Label1.TabIndex = 342;
            Label1.Text = "Destination Range:";
            // 
            // Selection_destination
            // 
            Selection_destination.BackColor = System.Drawing.Color.White;
            Selection_destination.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            Selection_destination.Image = (System.Drawing.Image)resources.GetObject("Selection_destination.Image");
            Selection_destination.Location = new System.Drawing.Point(253, 43);
            Selection_destination.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            Selection_destination.Name = "Selection_destination";
            Selection_destination.Size = new System.Drawing.Size(24, 25);
            Selection_destination.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            Selection_destination.TabIndex = 344;
            Selection_destination.TabStop = false;
            // 
            // TB_dest_range
            // 
            TB_dest_range.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            TB_dest_range.Cursor = System.Windows.Forms.Cursors.IBeam;
            TB_dest_range.Font = new System.Drawing.Font("Segoe UI", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            TB_dest_range.Location = new System.Drawing.Point(15, 43);
            TB_dest_range.Name = "TB_dest_range";
            TB_dest_range.Size = new System.Drawing.Size(262, 25);
            TB_dest_range.TabIndex = 343;
            // 
            // Label5
            // 
            Label5.AutoSize = true;
            Label5.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            Label5.Location = new System.Drawing.Point(15, 348);
            Label5.Name = "Label5";
            Label5.Size = new System.Drawing.Size(137, 17);
            Label5.TabIndex = 347;
            Label5.Text = "Total items in the list:";
            // 
            // Label6
            // 
            Label6.AutoSize = true;
            Label6.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            Label6.Location = new System.Drawing.Point(15, 375);
            Label6.Name = "Label6";
            Label6.Size = new System.Drawing.Size(82, 17);
            Label6.TabIndex = 348;
            Label6.Text = "List Preview:";
            // 
            // List_Preview
            // 
            List_Preview.BackColor = System.Drawing.SystemColors.Window;
            List_Preview.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            List_Preview.FormattingEnabled = true;
            List_Preview.ItemHeight = 17;
            List_Preview.Location = new System.Drawing.Point(15, 401);
            List_Preview.Name = "List_Preview";
            List_Preview.Size = new System.Drawing.Size(260, 89);
            List_Preview.TabIndex = 349;
            // 
            // ComboBox2
            // 
            ComboBox2.Font = new System.Drawing.Font("Segoe UI", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            ComboBox2.FormattingEnabled = true;
            ComboBox2.Location = new System.Drawing.Point(15, 509);
            ComboBox2.Name = "ComboBox2";
            ComboBox2.Size = new System.Drawing.Size(154, 25);
            ComboBox2.TabIndex = 350;
            ComboBox2.Text = "SOFTEKO";
            // 
            // Label7
            // 
            Label7.AutoSize = true;
            Label7.BackColor = System.Drawing.SystemColors.Control;
            Label7.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            Label7.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 0);
            Label7.Location = new System.Drawing.Point(155, 348);
            Label7.Name = "Label7";
            Label7.Size = new System.Drawing.Size(46, 17);
            Label7.TabIndex = 352;
            Label7.Text = "Label7";
            Label7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            Label7.Visible = false;
            // 
            // Btn_OK
            // 
            Btn_OK.BackColor = System.Drawing.Color.White;
            Btn_OK.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            Btn_OK.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            Btn_OK.Location = new System.Drawing.Point(485, 507);
            Btn_OK.Name = "Btn_OK";
            Btn_OK.Size = new System.Drawing.Size(62, 26);
            Btn_OK.TabIndex = 354;
            Btn_OK.Text = "OK";
            Btn_OK.UseVisualStyleBackColor = false;
            // 
            // Btn_Cancel
            // 
            Btn_Cancel.BackColor = System.Drawing.Color.White;
            Btn_Cancel.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            Btn_Cancel.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            Btn_Cancel.Location = new System.Drawing.Point(562, 507);
            Btn_Cancel.Name = "Btn_Cancel";
            Btn_Cancel.Size = new System.Drawing.Size(62, 26);
            Btn_Cancel.TabIndex = 353;
            Btn_Cancel.Text = "Cancel";
            Btn_Cancel.UseVisualStyleBackColor = false;
            // 
            // CustomGroupBox2
            // 
            CustomGroupBox2.BackColor = System.Drawing.Color.White;
            CustomGroupBox2.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            CustomGroupBox2.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 0);
            CustomGroupBox2.Location = new System.Drawing.Point(317, 15);
            CustomGroupBox2.Name = "CustomGroupBox2";
            CustomGroupBox2.Size = new System.Drawing.Size(307, 455);
            CustomGroupBox2.TabIndex = 351;
            CustomGroupBox2.TabStop = false;
            CustomGroupBox2.Text = "Sample Image";
            // 
            // CustomGroupBox1
            // 
            CustomGroupBox1.BackColor = System.Drawing.Color.White;
            CustomGroupBox1.BorderColor = System.Drawing.Color.FromArgb(64, 64, 64);
            CustomGroupBox1.Controls.Add(RadioButton3);
            CustomGroupBox1.Controls.Add(RadioButton2);
            CustomGroupBox1.Controls.Add(RadioButton1);
            CustomGroupBox1.Controls.Add(PictureBox1);
            CustomGroupBox1.Controls.Add(ComboBox1);
            CustomGroupBox1.Controls.Add(ListBox1);
            CustomGroupBox1.Controls.Add(Selection_Source);
            CustomGroupBox1.Controls.Add(Info);
            CustomGroupBox1.Controls.Add(TB_src_range);
            CustomGroupBox1.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 0);
            CustomGroupBox1.Location = new System.Drawing.Point(15, 88);
            CustomGroupBox1.Name = "CustomGroupBox1";
            CustomGroupBox1.Size = new System.Drawing.Size(260, 253);
            CustomGroupBox1.TabIndex = 346;
            CustomGroupBox1.TabStop = false;
            CustomGroupBox1.Text = "Source Range";
            // 
            // RadioButton3
            // 
            RadioButton3.AutoSize = true;
            RadioButton3.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            RadioButton3.Location = new System.Drawing.Point(15, 190);
            RadioButton3.Name = "RadioButton3";
            RadioButton3.Size = new System.Drawing.Size(94, 21);
            RadioButton3.TabIndex = 355;
            RadioButton3.Text = "Other Lists:";
            RadioButton3.UseVisualStyleBackColor = true;
            // 
            // RadioButton2
            // 
            RadioButton2.AutoSize = true;
            RadioButton2.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            RadioButton2.Location = new System.Drawing.Point(14, 76);
            RadioButton2.Name = "RadioButton2";
            RadioButton2.Size = new System.Drawing.Size(124, 21);
            RadioButton2.TabIndex = 354;
            RadioButton2.Text = "Predefined Lists:";
            RadioButton2.UseVisualStyleBackColor = true;
            // 
            // RadioButton1
            // 
            RadioButton1.AutoSize = true;
            RadioButton1.Checked = true;
            RadioButton1.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            RadioButton1.Location = new System.Drawing.Point(14, 20);
            RadioButton1.Name = "RadioButton1";
            RadioButton1.Size = new System.Drawing.Size(103, 21);
            RadioButton1.TabIndex = 353;
            RadioButton1.TabStop = true;
            RadioButton1.Text = "Enter Range:";
            RadioButton1.UseVisualStyleBackColor = true;
            // 
            // PictureBox1
            // 
            PictureBox1.Image = (System.Drawing.Image)resources.GetObject("PictureBox1.Image");
            PictureBox1.Location = new System.Drawing.Point(116, 191);
            PictureBox1.Name = "PictureBox1";
            PictureBox1.Size = new System.Drawing.Size(20, 20);
            PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            PictureBox1.TabIndex = 352;
            PictureBox1.TabStop = false;
            // 
            // ComboBox1
            // 
            ComboBox1.FormattingEnabled = true;
            ComboBox1.Location = new System.Drawing.Point(14, 217);
            ComboBox1.Name = "ComboBox1";
            ComboBox1.Size = new System.Drawing.Size(230, 25);
            ComboBox1.TabIndex = 351;
            ToolTip1.SetToolTip(ComboBox1, "Please, Enter the items separated by Comma (,)");
            // 
            // ListBox1
            // 
            ListBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            ListBox1.ColumnWidth = 10;
            ListBox1.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            ListBox1.Font = new System.Drawing.Font("Segoe UI", 9.0f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            ListBox1.ImeMode = System.Windows.Forms.ImeMode.Off;
            ListBox1.ItemHeight = 20;
            ListBox1.Items.AddRange(new object[] { "Sunday,Monday,Tuesday,Wednesday,Thursday,Friday,Saturday", "Sun,Mon,Tue,Wed,Thu,Fri,Sat", "January,February,March,April,May,June,July,August,September,October,November,Dece" + "mber", "Jan,Feb,Mar,Apr,May,Jun,July,Aug,Sep,Oct,Nov,Dec", "1,2,3,4,5,6,7,8,9,10", "I,II,III,IV,V,VI,VII,VIII,IX,X", "One,Two,Three,Four,Five,Six,Seven,Eight,Nine,Ten", "a,b,c ,d,e,f,g,h,i,j" });
            ListBox1.Location = new System.Drawing.Point(15, 102);
            ListBox1.Name = "ListBox1";
            ListBox1.Size = new System.Drawing.Size(230, 82);
            ListBox1.TabIndex = 349;
            // 
            // Selection_Source
            // 
            Selection_Source.BackColor = System.Drawing.Color.White;
            Selection_Source.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            Selection_Source.Image = (System.Drawing.Image)resources.GetObject("Selection_Source.Image");
            Selection_Source.Location = new System.Drawing.Point(220, 47);
            Selection_Source.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            Selection_Source.Name = "Selection_Source";
            Selection_Source.Size = new System.Drawing.Size(24, 25);
            Selection_Source.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            Selection_Source.TabIndex = 347;
            Selection_Source.TabStop = false;
            // 
            // Info
            // 
            Info.Image = (System.Drawing.Image)resources.GetObject("Info.Image");
            Info.Location = new System.Drawing.Point(121, 21);
            Info.Name = "Info";
            Info.Size = new System.Drawing.Size(20, 20);
            Info.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            Info.TabIndex = 345;
            Info.TabStop = false;
            // 
            // TB_src_range
            // 
            TB_src_range.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            TB_src_range.Cursor = System.Windows.Forms.Cursors.IBeam;
            TB_src_range.Font = new System.Drawing.Font("Segoe UI", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            TB_src_range.Location = new System.Drawing.Point(14, 47);
            TB_src_range.Name = "TB_src_range";
            TB_src_range.Size = new System.Drawing.Size(230, 25);
            TB_src_range.TabIndex = 346;
            // 
            // Form29_Simple_Drop_down_List
            // 
            AutoScaleDimensions = new System.Drawing.SizeF(6.0f, 13.0f);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            ClientSize = new System.Drawing.Size(650, 553);
            Controls.Add(Btn_OK);
            Controls.Add(Btn_Cancel);
            Controls.Add(Label7);
            Controls.Add(CustomGroupBox2);
            Controls.Add(ComboBox2);
            Controls.Add(List_Preview);
            Controls.Add(Label6);
            Controls.Add(Label5);
            Controls.Add(CustomGroupBox1);
            Controls.Add(Label1);
            Controls.Add(Selection_destination);
            Controls.Add(TB_dest_range);
            FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "Form29_Simple_Drop_down_List";
            Text = "Simple Drop-down List";
            ((System.ComponentModel.ISupportInitialize)Selection_destination).EndInit();
            CustomGroupBox1.ResumeLayout(false);
            CustomGroupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)PictureBox1).EndInit();
            ((System.ComponentModel.ISupportInitialize)Selection_Source).EndInit();
            ((System.ComponentModel.ISupportInitialize)Info).EndInit();
            Load += new EventHandler(Form1_Load);
            KeyDown += new System.Windows.Forms.KeyEventHandler(form);
            Closing += new System.ComponentModel.CancelEventHandler(Form29_Simple_Drop_down_List_Closing);
            Disposed += new EventHandler(Form29_Simple_Drop_down_List_Disposed);
            Shown += new EventHandler(Form29_Simple_Drop_down_List_Shown);
            ResumeLayout(false);
            PerformLayout();

        }

        internal System.Windows.Forms.Label Label1;
        internal System.Windows.Forms.PictureBox Selection_destination;
        internal System.Windows.Forms.TextBox TB_dest_range;
        internal CustomGroupBox CustomGroupBox1;
        internal System.Windows.Forms.ComboBox ComboBox1;
        internal System.Windows.Forms.PictureBox Selection_Source;
        internal System.Windows.Forms.PictureBox Info;
        internal System.Windows.Forms.TextBox TB_src_range;
        internal System.Windows.Forms.Label Label5;
        internal System.Windows.Forms.Label Label6;
        internal System.Windows.Forms.ListBox List_Preview;
        internal System.Windows.Forms.ListBox ListBox1;
        internal System.Windows.Forms.ComboBox ComboBox2;
        internal CustomGroupBox CustomGroupBox2;
        internal System.Windows.Forms.Label Label7;
        internal System.Windows.Forms.Button Btn_OK;
        internal System.Windows.Forms.Button Btn_Cancel;
        internal System.Windows.Forms.PictureBox PictureBox1;
        internal System.Windows.Forms.RadioButton RadioButton3;
        internal System.Windows.Forms.RadioButton RadioButton2;
        internal System.Windows.Forms.RadioButton RadioButton1;
        internal System.Windows.Forms.ToolTip ToolTip1;
    }
}