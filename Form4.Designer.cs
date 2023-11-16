using System;

namespace VSTO_Addins
{
    [Microsoft.VisualBasic.CompilerServices.DesignerGenerated()]
    public partial class Form4 : System.Windows.Forms.Form
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
            var resources = new System.ComponentModel.ComponentResourceManager(typeof(Form4));
            RadioButton1 = new System.Windows.Forms.RadioButton();
            RadioButton1.CheckedChanged += new EventHandler(RadioButton1_CheckedChanged);
            RadioButton1.GotFocus += new EventHandler(RadioButton1_GotFocus);
            RadioButton2 = new System.Windows.Forms.RadioButton();
            RadioButton2.CheckedChanged += new EventHandler(RadioButton2_CheckedChanged);
            RadioButton2.GotFocus += new EventHandler(RadioButton2_GotFocus);
            Label1 = new System.Windows.Forms.Label();
            PictureBox8 = new System.Windows.Forms.PictureBox();
            PictureBox8.Click += new EventHandler(PictureBox8_Click);
            PictureBox8.GotFocus += new EventHandler(PictureBox8_GotFocus);
            TextBox1 = new System.Windows.Forms.TextBox();
            TextBox1.TextChanged += new EventHandler(TextBox1_TextChanged);
            TextBox1.GotFocus += new EventHandler(TextBox1_GotFocus);
            PictureBox1 = new System.Windows.Forms.PictureBox();
            PictureBox1.Click += new EventHandler(PictureBox1_Click);
            PictureBox1.GotFocus += new EventHandler(PictureBox1_GotFocus);
            TextBox2 = new System.Windows.Forms.TextBox();
            TextBox2.TextChanged += new EventHandler(TextBox2_TextChanged);
            TextBox2.GotFocus += new EventHandler(TextBox2_GotFocus);
            Button1 = new System.Windows.Forms.Button();
            Button1.Click += new EventHandler(Button1_Click);
            Button1.GotFocus += new EventHandler(Button1_GotFocus);
            Button1.MouseEnter += new EventHandler(Button1_MouseEnter);
            Button1.MouseLeave += new EventHandler(Button1_MouseLeave);
            Button2 = new System.Windows.Forms.Button();
            Button2.Click += new EventHandler(Button2_Click);
            Button2.GotFocus += new EventHandler(Button2_GotFocus);
            Button2.MouseEnter += new EventHandler(Button2_MouseEnter);
            Button2.MouseLeave += new EventHandler(Button2_MouseLeave);
            Button3 = new System.Windows.Forms.Button();
            Button3.GotFocus += new EventHandler(Button3_GotFocus);
            Button3.Click += new EventHandler(Button3_Click);
            Button3.MouseEnter += new EventHandler(Button3_MouseEnter);
            Button3.MouseLeave += new EventHandler(Button3_MouseLeave);
            PictureBox2 = new System.Windows.Forms.PictureBox();
            PictureBox2.Click += new EventHandler(PictureBox2_Click);
            PictureBox2.GotFocus += new EventHandler(PictureBox2_GotFocus);
            TextBox3 = new System.Windows.Forms.TextBox();
            TextBox3.TextChanged += new EventHandler(TextBox3_TextChanged);
            TextBox3.GotFocus += new EventHandler(TextBox3_GotFocus);
            PictureBox3 = new System.Windows.Forms.PictureBox();
            PictureBox3.GotFocus += new EventHandler(PictureBox3_GotFocus);
            GroupBox1 = new System.Windows.Forms.GroupBox();
            ((System.ComponentModel.ISupportInitialize)PictureBox8).BeginInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox1).BeginInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox2).BeginInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox3).BeginInit();
            GroupBox1.SuspendLayout();
            SuspendLayout();
            // 
            // RadioButton1
            // 
            RadioButton1.AutoSize = true;
            RadioButton1.Font = new System.Drawing.Font("Segoe UI", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            RadioButton1.Location = new System.Drawing.Point(8, 19);
            RadioButton1.Name = "RadioButton1";
            RadioButton1.Size = new System.Drawing.Size(152, 21);
            RadioButton1.TabIndex = 0;
            RadioButton1.Text = "Open New Workbook";
            RadioButton1.UseVisualStyleBackColor = true;
            // 
            // RadioButton2
            // 
            RadioButton2.AutoSize = true;
            RadioButton2.Font = new System.Drawing.Font("Segoe UI", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            RadioButton2.Location = new System.Drawing.Point(8, 53);
            RadioButton2.Name = "RadioButton2";
            RadioButton2.Size = new System.Drawing.Size(170, 21);
            RadioButton2.TabIndex = 1;
            RadioButton2.Text = "Open Existing Workbook";
            RadioButton2.UseVisualStyleBackColor = true;
            // 
            // Label1
            // 
            Label1.AutoSize = true;
            Label1.Font = new System.Drawing.Font("Segoe UI", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            Label1.Location = new System.Drawing.Point(77, 84);
            Label1.Name = "Label1";
            Label1.Size = new System.Drawing.Size(83, 17);
            Label1.TabIndex = 2;
            Label1.Text = "Select Range";
            // 
            // PictureBox8
            // 
            PictureBox8.Enabled = false;
            PictureBox8.Image = (System.Drawing.Image)resources.GetObject("PictureBox8.Image");
            PictureBox8.Location = new System.Drawing.Point(333, 17);
            PictureBox8.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            PictureBox8.Name = "PictureBox8";
            PictureBox8.Size = new System.Drawing.Size(26, 24);
            PictureBox8.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            PictureBox8.TabIndex = 128;
            PictureBox8.TabStop = false;
            // 
            // TextBox1
            // 
            TextBox1.BackColor = System.Drawing.Color.White;
            TextBox1.Enabled = false;
            TextBox1.ForeColor = System.Drawing.Color.FromArgb(70, 70, 70);
            TextBox1.Location = new System.Drawing.Point(180, 16);
            TextBox1.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            TextBox1.Multiline = true;
            TextBox1.Name = "TextBox1";
            TextBox1.Size = new System.Drawing.Size(179, 26);
            TextBox1.TabIndex = 127;
            // 
            // PictureBox1
            // 
            PictureBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            PictureBox1.Enabled = false;
            PictureBox1.Image = (System.Drawing.Image)resources.GetObject("PictureBox1.Image");
            PictureBox1.Location = new System.Drawing.Point(333, 50);
            PictureBox1.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            PictureBox1.Name = "PictureBox1";
            PictureBox1.Size = new System.Drawing.Size(26, 24);
            PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            PictureBox1.TabIndex = 130;
            PictureBox1.TabStop = false;
            // 
            // TextBox2
            // 
            TextBox2.BackColor = System.Drawing.Color.White;
            TextBox2.Enabled = false;
            TextBox2.ForeColor = System.Drawing.Color.FromArgb(70, 70, 70);
            TextBox2.Location = new System.Drawing.Point(180, 49);
            TextBox2.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            TextBox2.Multiline = true;
            TextBox2.Name = "TextBox2";
            TextBox2.Size = new System.Drawing.Size(179, 26);
            TextBox2.TabIndex = 129;
            // 
            // Button1
            // 
            Button1.BackColor = System.Drawing.Color.White;
            Button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            Button1.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 0);
            Button1.Location = new System.Drawing.Point(162, 122);
            Button1.Name = "Button1";
            Button1.Size = new System.Drawing.Size(62, 26);
            Button1.TabIndex = 131;
            Button1.Text = "OK";
            Button1.UseVisualStyleBackColor = false;
            // 
            // Button2
            // 
            Button2.BackColor = System.Drawing.Color.White;
            Button2.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            Button2.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 0);
            Button2.Location = new System.Drawing.Point(230, 122);
            Button2.Name = "Button2";
            Button2.Size = new System.Drawing.Size(62, 26);
            Button2.TabIndex = 132;
            Button2.Text = "Back";
            Button2.UseVisualStyleBackColor = false;
            // 
            // Button3
            // 
            Button3.BackColor = System.Drawing.Color.White;
            Button3.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            Button3.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 0);
            Button3.Location = new System.Drawing.Point(298, 122);
            Button3.Name = "Button3";
            Button3.Size = new System.Drawing.Size(62, 26);
            Button3.TabIndex = 133;
            Button3.Text = "Cancel";
            Button3.UseVisualStyleBackColor = false;
            // 
            // PictureBox2
            // 
            PictureBox2.Enabled = false;
            PictureBox2.Image = (System.Drawing.Image)resources.GetObject("PictureBox2.Image");
            PictureBox2.Location = new System.Drawing.Point(333, 81);
            PictureBox2.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            PictureBox2.Name = "PictureBox2";
            PictureBox2.Size = new System.Drawing.Size(26, 24);
            PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            PictureBox2.TabIndex = 135;
            PictureBox2.TabStop = false;
            // 
            // TextBox3
            // 
            TextBox3.BackColor = System.Drawing.Color.White;
            TextBox3.Enabled = false;
            TextBox3.ForeColor = System.Drawing.Color.FromArgb(70, 70, 70);
            TextBox3.Location = new System.Drawing.Point(180, 80);
            TextBox3.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            TextBox3.Multiline = true;
            TextBox3.Name = "TextBox3";
            TextBox3.Size = new System.Drawing.Size(179, 26);
            TextBox3.TabIndex = 134;
            // 
            // PictureBox3
            // 
            PictureBox3.Image = (System.Drawing.Image)resources.GetObject("PictureBox3.Image");
            PictureBox3.Location = new System.Drawing.Point(52, 84);
            PictureBox3.Name = "PictureBox3";
            PictureBox3.Size = new System.Drawing.Size(19, 19);
            PictureBox3.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            PictureBox3.TabIndex = 136;
            PictureBox3.TabStop = false;
            // 
            // GroupBox1
            // 
            GroupBox1.Controls.Add(PictureBox3);
            GroupBox1.Controls.Add(PictureBox2);
            GroupBox1.Controls.Add(TextBox3);
            GroupBox1.Controls.Add(Button3);
            GroupBox1.Controls.Add(Button2);
            GroupBox1.Controls.Add(Button1);
            GroupBox1.Controls.Add(PictureBox1);
            GroupBox1.Controls.Add(TextBox2);
            GroupBox1.Controls.Add(PictureBox8);
            GroupBox1.Controls.Add(TextBox1);
            GroupBox1.Controls.Add(Label1);
            GroupBox1.Controls.Add(RadioButton2);
            GroupBox1.Controls.Add(RadioButton1);
            GroupBox1.Location = new System.Drawing.Point(12, 12);
            GroupBox1.Name = "GroupBox1";
            GroupBox1.Size = new System.Drawing.Size(373, 160);
            GroupBox1.TabIndex = 137;
            GroupBox1.TabStop = false;
            // 
            // Form4
            // 
            AutoScaleDimensions = new System.Drawing.SizeF(6.0f, 13.0f);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            ClientSize = new System.Drawing.Size(398, 187);
            Controls.Add(GroupBox1);
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "Form4";
            Text = " New workbook";
            ((System.ComponentModel.ISupportInitialize)PictureBox8).EndInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox1).EndInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox2).EndInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox3).EndInit();
            GroupBox1.ResumeLayout(false);
            GroupBox1.PerformLayout();
            Load += new EventHandler(Form4_Loaded);
            KeyDown += new System.Windows.Forms.KeyEventHandler(Form4_KeyDown);
            ResumeLayout(false);

        }

        internal System.Windows.Forms.RadioButton RadioButton1;
        internal System.Windows.Forms.RadioButton RadioButton2;
        internal System.Windows.Forms.Label Label1;
        internal System.Windows.Forms.PictureBox PictureBox8;
        internal System.Windows.Forms.TextBox TextBox1;
        internal System.Windows.Forms.PictureBox PictureBox1;
        internal System.Windows.Forms.TextBox TextBox2;
        internal System.Windows.Forms.Button Button1;
        internal System.Windows.Forms.Button Button2;
        internal System.Windows.Forms.Button Button3;
        internal System.Windows.Forms.PictureBox PictureBox2;
        internal System.Windows.Forms.TextBox TextBox3;
        internal System.Windows.Forms.PictureBox PictureBox3;
        internal System.Windows.Forms.GroupBox GroupBox1;
    }
}