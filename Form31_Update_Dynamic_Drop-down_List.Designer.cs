using System;

namespace VSTO_Addins
{
    [Microsoft.VisualBasic.CompilerServices.DesignerGenerated()]
    public partial class Form31_UpdateDynamicDropdownList : System.Windows.Forms.Form
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
            var resources = new System.ComponentModel.ComponentResourceManager(typeof(Form31_UpdateDynamicDropdownList));
            Label1 = new System.Windows.Forms.Label();
            TB_src_rng = new System.Windows.Forms.TextBox();
            TB_src_rng.KeyDown += new System.Windows.Forms.KeyEventHandler(TB_src_range_Enter);
            TB_src_rng.TextChanged += new EventHandler(TB_src_rng_TextChanged);
            Btn_OK = new System.Windows.Forms.Button();
            Btn_OK.Click += new EventHandler(Btn_OK_Click);
            Btn_OK.KeyDown += new System.Windows.Forms.KeyEventHandler(OK);
            Btn_Cancel = new System.Windows.Forms.Button();
            Btn_Cancel.Click += new EventHandler(Btn_Cancel_Click);
            Btn_Cancel.KeyDown += new System.Windows.Forms.KeyEventHandler(Cancel);
            ComboBox2 = new System.Windows.Forms.ComboBox();
            Selection_source = new System.Windows.Forms.PictureBox();
            Selection_source.Click += new EventHandler(Selection_source_Click);
            Info = new System.Windows.Forms.PictureBox();
            CustomGroupBox2 = new CustomGroupBox();
            CustomGroupBox2.Enter += new EventHandler(CustomGroupBox2_Enter);
            CustomGroupBox10 = new CustomGroupBox();
            TB_des_rng1 = new System.Windows.Forms.TextBox();
            PictureBox2 = new System.Windows.Forms.PictureBox();
            PictureBox3 = new System.Windows.Forms.PictureBox();
            PictureBox3.Click += new EventHandler(PictureBox3_Click);
            TB_des_rng2 = new System.Windows.Forms.TextBox();
            TB_des_rng2.KeyDown += new System.Windows.Forms.KeyEventHandler(TB_dest_range_Enter);
            TB_des_rng2.TextChanged += new EventHandler(TB_des_rng2_TextChanged);
            L_select = new System.Windows.Forms.Label();
            RB_diff_rng = new System.Windows.Forms.RadioButton();
            RB_diff_rng.CheckedChanged += new EventHandler(RB_diff_rng_CheckedChanged);
            RB_diff_rng.KeyDown += new System.Windows.Forms.KeyEventHandler(RB_Different);
            RB_same_source = new System.Windows.Forms.RadioButton();
            RB_same_source.CheckedChanged += new EventHandler(RB_same_source_CheckedChanged);
            RB_same_source.KeyDown += new System.Windows.Forms.KeyEventHandler(RB_same);
            TextBox1 = new System.Windows.Forms.TextBox();
            ((System.ComponentModel.ISupportInitialize)Selection_source).BeginInit();
            ((System.ComponentModel.ISupportInitialize)Info).BeginInit();
            CustomGroupBox2.SuspendLayout();
            CustomGroupBox10.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)PictureBox2).BeginInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox3).BeginInit();
            SuspendLayout();
            // 
            // Label1
            // 
            Label1.AutoSize = true;
            Label1.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            Label1.Location = new System.Drawing.Point(15, 15);
            Label1.Name = "Label1";
            Label1.Size = new System.Drawing.Size(154, 17);
            Label1.TabIndex = 0;
            Label1.Text = "Updated Source Range :";
            // 
            // TB_src_rng
            // 
            TB_src_rng.Font = new System.Drawing.Font("Segoe UI", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            TB_src_rng.Location = new System.Drawing.Point(15, 44);
            TB_src_rng.Name = "TB_src_rng";
            TB_src_rng.Size = new System.Drawing.Size(289, 25);
            TB_src_rng.TabIndex = 1;
            // 
            // Btn_OK
            // 
            Btn_OK.BackColor = System.Drawing.Color.White;
            Btn_OK.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            Btn_OK.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            Btn_OK.Location = new System.Drawing.Point(192, 274);
            Btn_OK.Name = "Btn_OK";
            Btn_OK.Size = new System.Drawing.Size(62, 26);
            Btn_OK.TabIndex = 379;
            Btn_OK.Text = "OK";
            Btn_OK.UseVisualStyleBackColor = false;
            // 
            // Btn_Cancel
            // 
            Btn_Cancel.BackColor = System.Drawing.Color.White;
            Btn_Cancel.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            Btn_Cancel.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            Btn_Cancel.Location = new System.Drawing.Point(263, 274);
            Btn_Cancel.Name = "Btn_Cancel";
            Btn_Cancel.Size = new System.Drawing.Size(62, 26);
            Btn_Cancel.TabIndex = 378;
            Btn_Cancel.Text = "Cancel";
            Btn_Cancel.UseVisualStyleBackColor = false;
            // 
            // ComboBox2
            // 
            ComboBox2.Font = new System.Drawing.Font("Segoe UI", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            ComboBox2.FormattingEnabled = true;
            ComboBox2.Location = new System.Drawing.Point(16, 274);
            ComboBox2.Name = "ComboBox2";
            ComboBox2.Size = new System.Drawing.Size(96, 25);
            ComboBox2.TabIndex = 377;
            ComboBox2.Text = "Softeko";
            // 
            // Selection_source
            // 
            Selection_source.BackColor = System.Drawing.Color.White;
            Selection_source.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            Selection_source.Image = (System.Drawing.Image)resources.GetObject("Selection_source.Image");
            Selection_source.Location = new System.Drawing.Point(280, 44);
            Selection_source.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            Selection_source.Name = "Selection_source";
            Selection_source.Size = new System.Drawing.Size(24, 25);
            Selection_source.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            Selection_source.TabIndex = 381;
            Selection_source.TabStop = false;
            // 
            // Info
            // 
            Info.Image = (System.Drawing.Image)resources.GetObject("Info.Image");
            Info.Location = new System.Drawing.Point(308, 46);
            Info.Name = "Info";
            Info.Size = new System.Drawing.Size(20, 20);
            Info.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            Info.TabIndex = 382;
            Info.TabStop = false;
            // 
            // CustomGroupBox2
            // 
            CustomGroupBox2.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            CustomGroupBox2.Controls.Add(CustomGroupBox10);
            CustomGroupBox2.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 0);
            CustomGroupBox2.Location = new System.Drawing.Point(15, 83);
            CustomGroupBox2.Name = "CustomGroupBox2";
            CustomGroupBox2.Size = new System.Drawing.Size(310, 170);
            CustomGroupBox2.TabIndex = 270;
            CustomGroupBox2.TabStop = false;
            CustomGroupBox2.Text = "Destination Range";
            // 
            // CustomGroupBox10
            // 
            CustomGroupBox10.BackColor = System.Drawing.Color.White;
            CustomGroupBox10.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            CustomGroupBox10.Controls.Add(TB_des_rng1);
            CustomGroupBox10.Controls.Add(PictureBox2);
            CustomGroupBox10.Controls.Add(PictureBox3);
            CustomGroupBox10.Controls.Add(TB_des_rng2);
            CustomGroupBox10.Controls.Add(L_select);
            CustomGroupBox10.Controls.Add(RB_diff_rng);
            CustomGroupBox10.Controls.Add(RB_same_source);
            CustomGroupBox10.Location = new System.Drawing.Point(1, 22);
            CustomGroupBox10.Name = "CustomGroupBox10";
            CustomGroupBox10.Size = new System.Drawing.Size(308, 148);
            CustomGroupBox10.TabIndex = 0;
            CustomGroupBox10.TabStop = false;
            // 
            // TB_des_rng1
            // 
            TB_des_rng1.Font = new System.Drawing.Font("Segoe UI", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            TB_des_rng1.Location = new System.Drawing.Point(25, 33);
            TB_des_rng1.Name = "TB_des_rng1";
            TB_des_rng1.Size = new System.Drawing.Size(263, 25);
            TB_des_rng1.TabIndex = 209;
            // 
            // PictureBox2
            // 
            PictureBox2.Image = (System.Drawing.Image)resources.GetObject("PictureBox2.Image");
            PictureBox2.Location = new System.Drawing.Point(25, 88);
            PictureBox2.Name = "PictureBox2";
            PictureBox2.Size = new System.Drawing.Size(14, 14);
            PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            PictureBox2.TabIndex = 208;
            PictureBox2.TabStop = false;
            // 
            // PictureBox3
            // 
            PictureBox3.BackColor = System.Drawing.Color.White;
            PictureBox3.Image = (System.Drawing.Image)resources.GetObject("PictureBox3.Image");
            PictureBox3.Location = new System.Drawing.Point(262, 115);
            PictureBox3.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            PictureBox3.Name = "PictureBox3";
            PictureBox3.Size = new System.Drawing.Size(24, 23);
            PictureBox3.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            PictureBox3.TabIndex = 207;
            PictureBox3.TabStop = false;
            // 
            // TB_des_rng2
            // 
            TB_des_rng2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            TB_des_rng2.Cursor = System.Windows.Forms.Cursors.IBeam;
            TB_des_rng2.Font = new System.Drawing.Font("Segoe UI", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            TB_des_rng2.Location = new System.Drawing.Point(24, 114);
            TB_des_rng2.Name = "TB_des_rng2";
            TB_des_rng2.Size = new System.Drawing.Size(263, 25);
            TB_des_rng2.TabIndex = 206;
            // 
            // L_select
            // 
            L_select.AutoSize = true;
            L_select.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            L_select.Location = new System.Drawing.Point(42, 85);
            L_select.Name = "L_select";
            L_select.Size = new System.Drawing.Size(109, 17);
            L_select.TabIndex = 2;
            L_select.Text = "Select the range :";
            // 
            // RB_diff_rng
            // 
            RB_diff_rng.AutoSize = true;
            RB_diff_rng.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            RB_diff_rng.Location = new System.Drawing.Point(8, 62);
            RB_diff_rng.Name = "RB_diff_rng";
            RB_diff_rng.Size = new System.Drawing.Size(185, 21);
            RB_diff_rng.TabIndex = 1;
            RB_diff_rng.Text = "Store into a different range";
            RB_diff_rng.UseVisualStyleBackColor = true;
            // 
            // RB_same_source
            // 
            RB_same_source.AutoSize = true;
            RB_same_source.Checked = true;
            RB_same_source.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            RB_same_source.Location = new System.Drawing.Point(8, 6);
            RB_same_source.Name = "RB_same_source";
            RB_same_source.Size = new System.Drawing.Size(225, 21);
            RB_same_source.TabIndex = 0;
            RB_same_source.TabStop = true;
            RB_same_source.Text = "Same as the original output range";
            RB_same_source.UseVisualStyleBackColor = true;
            // 
            // TextBox1
            // 
            TextBox1.Location = new System.Drawing.Point(373, 83);
            TextBox1.Name = "TextBox1";
            TextBox1.Size = new System.Drawing.Size(100, 20);
            TextBox1.TabIndex = 383;
            TextBox1.Visible = false;
            // 
            // Form31_UpdateDynamicDropdownList
            // 
            AutoScaleDimensions = new System.Drawing.SizeF(6.0f, 13.0f);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            ClientSize = new System.Drawing.Size(345, 322);
            Controls.Add(TextBox1);
            Controls.Add(Info);
            Controls.Add(Selection_source);
            Controls.Add(Btn_OK);
            Controls.Add(Btn_Cancel);
            Controls.Add(ComboBox2);
            Controls.Add(CustomGroupBox2);
            Controls.Add(TB_src_rng);
            Controls.Add(Label1);
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "Form31_UpdateDynamicDropdownList";
            Text = "Update Dynamic Drop-down List";
            ((System.ComponentModel.ISupportInitialize)Selection_source).EndInit();
            ((System.ComponentModel.ISupportInitialize)Info).EndInit();
            CustomGroupBox2.ResumeLayout(false);
            CustomGroupBox10.ResumeLayout(false);
            CustomGroupBox10.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)PictureBox2).EndInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox3).EndInit();
            Load += new EventHandler(Form1_Load);
            KeyDown += new System.Windows.Forms.KeyEventHandler(Form_load);
            Closing += new System.ComponentModel.CancelEventHandler(Form31_UpdateDynamicDropdownList_Closing);
            Disposed += new EventHandler(Form31_UpdateDynamicDropdownList_Disposed);
            Shown += new EventHandler(Form31_UpdateDynamicDropdownList_Shown);
            ResumeLayout(false);
            PerformLayout();

        }

        internal System.Windows.Forms.Label Label1;
        internal System.Windows.Forms.TextBox TB_src_rng;
        internal CustomGroupBox CustomGroupBox2;
        internal CustomGroupBox CustomGroupBox10;
        internal System.Windows.Forms.TextBox TB_des_rng1;
        internal System.Windows.Forms.PictureBox PictureBox2;
        internal System.Windows.Forms.PictureBox PictureBox3;
        internal System.Windows.Forms.TextBox TB_des_rng2;
        internal System.Windows.Forms.Label L_select;
        internal System.Windows.Forms.RadioButton RB_diff_rng;
        internal System.Windows.Forms.RadioButton RB_same_source;
        internal System.Windows.Forms.Button Btn_OK;
        internal System.Windows.Forms.Button Btn_Cancel;
        internal System.Windows.Forms.ComboBox ComboBox2;
        internal System.Windows.Forms.PictureBox Selection_source;
        internal System.Windows.Forms.PictureBox Info;
        internal System.Windows.Forms.TextBox TextBox1;
    }
}