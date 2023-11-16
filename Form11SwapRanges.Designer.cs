using System;

namespace VSTO_Addins
{
    [Microsoft.VisualBasic.CompilerServices.DesignerGenerated()]
    public partial class Form11SwapRanges : System.Windows.Forms.Form
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
            var resources = new System.ComponentModel.ComponentResourceManager(typeof(Form11SwapRanges));
            CB_CopyWs = new System.Windows.Forms.CheckBox();
            ComboBox1 = new System.Windows.Forms.ComboBox();
            btnCancel = new System.Windows.Forms.Button();
            btnCancel.Click += new EventHandler(btnCancel_Click);
            btnOK = new System.Windows.Forms.Button();
            btnOK.Click += new EventHandler(btnOK_Click);
            PictureBox7 = new System.Windows.Forms.PictureBox();
            CB_KeepFormatting = new System.Windows.Forms.CheckBox();
            CustomGroupBox2 = new CustomGroupBox();
            AutoSelection2 = new System.Windows.Forms.PictureBox();
            AutoSelection2.Click += new EventHandler(AutoSelection2_Click);
            rngSelection2 = new System.Windows.Forms.PictureBox();
            rngSelection2.Click += new EventHandler(rngSelection2_Click);
            txtSourceRange2 = new System.Windows.Forms.TextBox();
            txtSourceRange2.TextChanged += new EventHandler(txtSourceRange2_TextChanged);
            txtSourceRange2.GotFocus += new EventHandler(txtSourceRange2_GotFocus);
            lblSourceRng2 = new System.Windows.Forms.Label();
            AutoSelection1 = new System.Windows.Forms.PictureBox();
            AutoSelection1.Click += new EventHandler(AutoSelection1_Click);
            rngSelection1 = new System.Windows.Forms.PictureBox();
            rngSelection1.Click += new EventHandler(rngSelection1_Click);
            txtSourceRange1 = new System.Windows.Forms.TextBox();
            txtSourceRange1.TextChanged += new EventHandler(txtSourceRange1_TextChanged);
            txtSourceRange1.GotFocus += new EventHandler(txtSourceRange1_GotFocus);
            lblSourceRng1 = new System.Windows.Forms.Label();
            CustomGroupBox6 = new CustomGroupBox();
            CP_OutputRng = new CustomPanel();
            CustomGroupBox5 = new CustomGroupBox();
            CP_InputRng = new CustomPanel();
            CustomGroupBox1 = new CustomGroupBox();
            CustomGroupBox7 = new CustomGroupBox();
            PictureBox2 = new System.Windows.Forms.PictureBox();
            radBtnValues = new System.Windows.Forms.RadioButton();
            PictureBox1 = new System.Windows.Forms.PictureBox();
            PictureBox5 = new System.Windows.Forms.PictureBox();
            radBtnKeepRef = new System.Windows.Forms.RadioButton();
            radBtnAdjustRef = new System.Windows.Forms.RadioButton();
            ((System.ComponentModel.ISupportInitialize)PictureBox7).BeginInit();
            CustomGroupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)AutoSelection2).BeginInit();
            ((System.ComponentModel.ISupportInitialize)rngSelection2).BeginInit();
            ((System.ComponentModel.ISupportInitialize)AutoSelection1).BeginInit();
            ((System.ComponentModel.ISupportInitialize)rngSelection1).BeginInit();
            CustomGroupBox6.SuspendLayout();
            CustomGroupBox5.SuspendLayout();
            CustomGroupBox1.SuspendLayout();
            CustomGroupBox7.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)PictureBox2).BeginInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox1).BeginInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox5).BeginInit();
            SuspendLayout();
            // 
            // CB_CopyWs
            // 
            CB_CopyWs.AutoSize = true;
            CB_CopyWs.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            CB_CopyWs.Location = new System.Drawing.Point(15, 309);
            CB_CopyWs.Name = "CB_CopyWs";
            CB_CopyWs.Size = new System.Drawing.Size(257, 21);
            CB_CopyWs.TabIndex = 151;
            CB_CopyWs.Text = "Create a copy of the original worksheet";
            CB_CopyWs.UseVisualStyleBackColor = true;
            // 
            // ComboBox1
            // 
            ComboBox1.Font = new System.Drawing.Font("Segoe UI", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            ComboBox1.FormattingEnabled = true;
            ComboBox1.Location = new System.Drawing.Point(15, 346);
            ComboBox1.Name = "ComboBox1";
            ComboBox1.Size = new System.Drawing.Size(154, 25);
            ComboBox1.TabIndex = 152;
            ComboBox1.Text = "SOFTEKO";
            // 
            // btnCancel
            // 
            btnCancel.BackColor = System.Drawing.Color.White;
            btnCancel.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            btnCancel.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            btnCancel.Location = new System.Drawing.Point(520, 349);
            btnCancel.Name = "btnCancel";
            btnCancel.Size = new System.Drawing.Size(62, 26);
            btnCancel.TabIndex = 155;
            btnCancel.Text = "Cancel";
            btnCancel.UseVisualStyleBackColor = false;
            // 
            // btnOK
            // 
            btnOK.BackColor = System.Drawing.Color.White;
            btnOK.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            btnOK.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            btnOK.Location = new System.Drawing.Point(448, 349);
            btnOK.Name = "btnOK";
            btnOK.Size = new System.Drawing.Size(62, 26);
            btnOK.TabIndex = 156;
            btnOK.Text = "OK";
            btnOK.UseVisualStyleBackColor = false;
            // 
            // PictureBox7
            // 
            PictureBox7.Image = (System.Drawing.Image)resources.GetObject("PictureBox7.Image");
            PictureBox7.Location = new System.Drawing.Point(437, 145);
            PictureBox7.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            PictureBox7.Name = "PictureBox7";
            PictureBox7.Size = new System.Drawing.Size(43, 49);
            PictureBox7.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            PictureBox7.TabIndex = 157;
            PictureBox7.TabStop = false;
            // 
            // CB_KeepFormatting
            // 
            CB_KeepFormatting.AutoSize = true;
            CB_KeepFormatting.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            CB_KeepFormatting.Location = new System.Drawing.Point(15, 280);
            CB_KeepFormatting.Name = "CB_KeepFormatting";
            CB_KeepFormatting.Size = new System.Drawing.Size(122, 21);
            CB_KeepFormatting.TabIndex = 150;
            CB_KeepFormatting.Text = "Keep formatting";
            CB_KeepFormatting.UseVisualStyleBackColor = true;
            // 
            // CustomGroupBox2
            // 
            CustomGroupBox2.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            CustomGroupBox2.Controls.Add(AutoSelection2);
            CustomGroupBox2.Controls.Add(rngSelection2);
            CustomGroupBox2.Controls.Add(txtSourceRange2);
            CustomGroupBox2.Controls.Add(lblSourceRng2);
            CustomGroupBox2.Controls.Add(AutoSelection1);
            CustomGroupBox2.Controls.Add(rngSelection1);
            CustomGroupBox2.Controls.Add(txtSourceRange1);
            CustomGroupBox2.Controls.Add(lblSourceRng1);
            CustomGroupBox2.Location = new System.Drawing.Point(15, 15);
            CustomGroupBox2.Name = "CustomGroupBox2";
            CustomGroupBox2.Size = new System.Drawing.Size(278, 136);
            CustomGroupBox2.TabIndex = 164;
            CustomGroupBox2.TabStop = false;
            // 
            // AutoSelection2
            // 
            AutoSelection2.Image = (System.Drawing.Image)resources.GetObject("AutoSelection2.Image");
            AutoSelection2.Location = new System.Drawing.Point(212, 94);
            AutoSelection2.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            AutoSelection2.Name = "AutoSelection2";
            AutoSelection2.Size = new System.Drawing.Size(25, 23);
            AutoSelection2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            AutoSelection2.TabIndex = 171;
            AutoSelection2.TabStop = false;
            // 
            // rngSelection2
            // 
            rngSelection2.Image = (System.Drawing.Image)resources.GetObject("rngSelection2.Image");
            rngSelection2.Location = new System.Drawing.Point(242, 94);
            rngSelection2.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            rngSelection2.Name = "rngSelection2";
            rngSelection2.Size = new System.Drawing.Size(25, 23);
            rngSelection2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            rngSelection2.TabIndex = 170;
            rngSelection2.TabStop = false;
            // 
            // txtSourceRange2
            // 
            txtSourceRange2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            txtSourceRange2.Cursor = System.Windows.Forms.Cursors.IBeam;
            txtSourceRange2.Font = new System.Drawing.Font("Segoe UI", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            txtSourceRange2.Location = new System.Drawing.Point(8, 93);
            txtSourceRange2.Name = "txtSourceRange2";
            txtSourceRange2.Size = new System.Drawing.Size(260, 25);
            txtSourceRange2.TabIndex = 169;
            // 
            // lblSourceRng2
            // 
            lblSourceRng2.AutoSize = true;
            lblSourceRng2.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            lblSourceRng2.Location = new System.Drawing.Point(8, 65);
            lblSourceRng2.Name = "lblSourceRng2";
            lblSourceRng2.Size = new System.Drawing.Size(256, 17);
            lblSourceRng2.TabIndex = 168;
            lblSourceRng2.Text = "2nd Source Range (X rows x Y columns) :";
            // 
            // AutoSelection1
            // 
            AutoSelection1.Image = (System.Drawing.Image)resources.GetObject("AutoSelection1.Image");
            AutoSelection1.Location = new System.Drawing.Point(212, 31);
            AutoSelection1.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            AutoSelection1.Name = "AutoSelection1";
            AutoSelection1.Size = new System.Drawing.Size(25, 23);
            AutoSelection1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            AutoSelection1.TabIndex = 167;
            AutoSelection1.TabStop = false;
            // 
            // rngSelection1
            // 
            rngSelection1.Image = (System.Drawing.Image)resources.GetObject("rngSelection1.Image");
            rngSelection1.Location = new System.Drawing.Point(242, 31);
            rngSelection1.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            rngSelection1.Name = "rngSelection1";
            rngSelection1.Size = new System.Drawing.Size(25, 23);
            rngSelection1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            rngSelection1.TabIndex = 166;
            rngSelection1.TabStop = false;
            // 
            // txtSourceRange1
            // 
            txtSourceRange1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            txtSourceRange1.Cursor = System.Windows.Forms.Cursors.IBeam;
            txtSourceRange1.Font = new System.Drawing.Font("Segoe UI", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            txtSourceRange1.Location = new System.Drawing.Point(8, 30);
            txtSourceRange1.Name = "txtSourceRange1";
            txtSourceRange1.Size = new System.Drawing.Size(260, 25);
            txtSourceRange1.TabIndex = 165;
            // 
            // lblSourceRng1
            // 
            lblSourceRng1.AutoSize = true;
            lblSourceRng1.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            lblSourceRng1.Location = new System.Drawing.Point(8, 6);
            lblSourceRng1.Name = "lblSourceRng1";
            lblSourceRng1.Size = new System.Drawing.Size(249, 17);
            lblSourceRng1.TabIndex = 164;
            lblSourceRng1.Text = "1st Source Range (X rows x Y columns) :";
            // 
            // CustomGroupBox6
            // 
            CustomGroupBox6.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            CustomGroupBox6.Controls.Add(CP_OutputRng);
            CustomGroupBox6.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 0);
            CustomGroupBox6.Location = new System.Drawing.Point(330, 195);
            CustomGroupBox6.Name = "CustomGroupBox6";
            CustomGroupBox6.Size = new System.Drawing.Size(252, 135);
            CustomGroupBox6.TabIndex = 154;
            CustomGroupBox6.TabStop = false;
            CustomGroupBox6.Text = "Expected Output";
            // 
            // CP_OutputRng
            // 
            CP_OutputRng.BackColor = System.Drawing.Color.White;
            CP_OutputRng.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            CP_OutputRng.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            CP_OutputRng.BorderWidth = 1;
            CP_OutputRng.Location = new System.Drawing.Point(1, 30);
            CP_OutputRng.Name = "CP_OutputRng";
            CP_OutputRng.Size = new System.Drawing.Size(250, 105);
            CP_OutputRng.TabIndex = 11;
            // 
            // CustomGroupBox5
            // 
            CustomGroupBox5.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            CustomGroupBox5.Controls.Add(CP_InputRng);
            CustomGroupBox5.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 0);
            CustomGroupBox5.Location = new System.Drawing.Point(329, 4);
            CustomGroupBox5.Name = "CustomGroupBox5";
            CustomGroupBox5.Size = new System.Drawing.Size(252, 135);
            CustomGroupBox5.TabIndex = 153;
            CustomGroupBox5.TabStop = false;
            CustomGroupBox5.Text = "Input Range";
            // 
            // CP_InputRng
            // 
            CP_InputRng.BackColor = System.Drawing.Color.White;
            CP_InputRng.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            CP_InputRng.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            CP_InputRng.BorderWidth = 1;
            CP_InputRng.Location = new System.Drawing.Point(1, 30);
            CP_InputRng.Name = "CP_InputRng";
            CP_InputRng.Size = new System.Drawing.Size(250, 105);
            CP_InputRng.TabIndex = 0;
            // 
            // CustomGroupBox1
            // 
            CustomGroupBox1.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            CustomGroupBox1.Controls.Add(CustomGroupBox7);
            CustomGroupBox1.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 0);
            CustomGroupBox1.Location = new System.Drawing.Point(15, 162);
            CustomGroupBox1.Name = "CustomGroupBox1";
            CustomGroupBox1.Size = new System.Drawing.Size(281, 110);
            CustomGroupBox1.TabIndex = 148;
            CustomGroupBox1.TabStop = false;
            CustomGroupBox1.Text = "Swap Type";
            // 
            // CustomGroupBox7
            // 
            CustomGroupBox7.BackColor = System.Drawing.Color.White;
            CustomGroupBox7.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            CustomGroupBox7.Controls.Add(PictureBox2);
            CustomGroupBox7.Controls.Add(radBtnValues);
            CustomGroupBox7.Controls.Add(PictureBox1);
            CustomGroupBox7.Controls.Add(PictureBox5);
            CustomGroupBox7.Controls.Add(radBtnKeepRef);
            CustomGroupBox7.Controls.Add(radBtnAdjustRef);
            CustomGroupBox7.Location = new System.Drawing.Point(1, 22);
            CustomGroupBox7.Name = "CustomGroupBox7";
            CustomGroupBox7.Size = new System.Drawing.Size(280, 88);
            CustomGroupBox7.TabIndex = 0;
            CustomGroupBox7.TabStop = false;
            // 
            // PictureBox2
            // 
            PictureBox2.Image = (System.Drawing.Image)resources.GetObject("PictureBox2.Image");
            PictureBox2.Location = new System.Drawing.Point(245, 61);
            PictureBox2.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            PictureBox2.Name = "PictureBox2";
            PictureBox2.Size = new System.Drawing.Size(20, 20);
            PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            PictureBox2.TabIndex = 130;
            PictureBox2.TabStop = false;
            // 
            // radBtnValues
            // 
            radBtnValues.AutoSize = true;
            radBtnValues.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            radBtnValues.Location = new System.Drawing.Point(8, 8);
            radBtnValues.Name = "radBtnValues";
            radBtnValues.Size = new System.Drawing.Size(93, 21);
            radBtnValues.TabIndex = 129;
            radBtnValues.TabStop = true;
            radBtnValues.Text = "Values Only";
            radBtnValues.UseVisualStyleBackColor = true;
            // 
            // PictureBox1
            // 
            PictureBox1.Image = (System.Drawing.Image)resources.GetObject("PictureBox1.Image");
            PictureBox1.Location = new System.Drawing.Point(245, 34);
            PictureBox1.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            PictureBox1.Name = "PictureBox1";
            PictureBox1.Size = new System.Drawing.Size(20, 20);
            PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            PictureBox1.TabIndex = 128;
            PictureBox1.TabStop = false;
            // 
            // PictureBox5
            // 
            PictureBox5.Image = (System.Drawing.Image)resources.GetObject("PictureBox5.Image");
            PictureBox5.Location = new System.Drawing.Point(245, 7);
            PictureBox5.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            PictureBox5.Name = "PictureBox5";
            PictureBox5.Size = new System.Drawing.Size(20, 20);
            PictureBox5.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            PictureBox5.TabIndex = 127;
            PictureBox5.TabStop = false;
            // 
            // radBtnKeepRef
            // 
            radBtnKeepRef.AutoSize = true;
            radBtnKeepRef.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            radBtnKeepRef.Location = new System.Drawing.Point(8, 60);
            radBtnKeepRef.Name = "radBtnKeepRef";
            radBtnKeepRef.Size = new System.Drawing.Size(143, 21);
            radBtnKeepRef.TabIndex = 1;
            radBtnKeepRef.TabStop = true;
            radBtnKeepRef.Text = "Keep Cell Reference";
            radBtnKeepRef.UseVisualStyleBackColor = true;
            // 
            // radBtnAdjustRef
            // 
            radBtnAdjustRef.AutoSize = true;
            radBtnAdjustRef.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            radBtnAdjustRef.Location = new System.Drawing.Point(8, 34);
            radBtnAdjustRef.Name = "radBtnAdjustRef";
            radBtnAdjustRef.Size = new System.Drawing.Size(149, 21);
            radBtnAdjustRef.TabIndex = 0;
            radBtnAdjustRef.TabStop = true;
            radBtnAdjustRef.Text = "Adjust Cell Reference";
            radBtnAdjustRef.UseVisualStyleBackColor = true;
            // 
            // Form11SwapRanges
            // 
            AutoScaleDimensions = new System.Drawing.SizeF(6.0f, 13.0f);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            ClientSize = new System.Drawing.Size(602, 393);
            Controls.Add(CustomGroupBox2);
            Controls.Add(CustomGroupBox6);
            Controls.Add(CustomGroupBox5);
            Controls.Add(CB_KeepFormatting);
            Controls.Add(PictureBox7);
            Controls.Add(btnOK);
            Controls.Add(btnCancel);
            Controls.Add(ComboBox1);
            Controls.Add(CB_CopyWs);
            Controls.Add(CustomGroupBox1);
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "Form11SwapRanges";
            StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            Text = "Swap";
            ((System.ComponentModel.ISupportInitialize)PictureBox7).EndInit();
            CustomGroupBox2.ResumeLayout(false);
            CustomGroupBox2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)AutoSelection2).EndInit();
            ((System.ComponentModel.ISupportInitialize)rngSelection2).EndInit();
            ((System.ComponentModel.ISupportInitialize)AutoSelection1).EndInit();
            ((System.ComponentModel.ISupportInitialize)rngSelection1).EndInit();
            CustomGroupBox6.ResumeLayout(false);
            CustomGroupBox5.ResumeLayout(false);
            CustomGroupBox1.ResumeLayout(false);
            CustomGroupBox7.ResumeLayout(false);
            CustomGroupBox7.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)PictureBox2).EndInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox1).EndInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox5).EndInit();
            KeyDown += new System.Windows.Forms.KeyEventHandler(Form1_KeyDown);
            Load += new EventHandler(Form11SwapRanges_Load);
            Activated += new EventHandler(Form1_Activated);
            Closing += new System.ComponentModel.CancelEventHandler(Form11SwapRanges_Closing);
            Shown += new EventHandler(Form11SwapRanges_Shown);
            Disposed += new EventHandler(Form11SwapRanges_Disposed);
            ResumeLayout(false);
            PerformLayout();

        }

        internal System.Windows.Forms.Label lblSourceRng1;
        internal System.Windows.Forms.TextBox txtSourceRange1;
        internal System.Windows.Forms.PictureBox rngSelection1;
        internal System.Windows.Forms.PictureBox AutoSelection1;
        internal System.Windows.Forms.Label lblSourceRng2;
        internal System.Windows.Forms.TextBox txtSourceRange2;
        internal System.Windows.Forms.PictureBox rngSelection2;
        internal System.Windows.Forms.PictureBox AutoSelection2;
        internal CustomGroupBox CustomGroupBox2;
        internal CustomGroupBox CustomGroupBox1;
        internal CustomGroupBox CustomGroupBox7;
        internal System.Windows.Forms.PictureBox PictureBox2;
        internal System.Windows.Forms.RadioButton radBtnValues;
        internal System.Windows.Forms.PictureBox PictureBox1;
        internal System.Windows.Forms.PictureBox PictureBox5;
        internal System.Windows.Forms.RadioButton radBtnKeepRef;
        internal System.Windows.Forms.RadioButton radBtnAdjustRef;
        internal System.Windows.Forms.CheckBox CB_CopyWs;
        internal System.Windows.Forms.ComboBox ComboBox1;
        internal System.Windows.Forms.Button btnCancel;
        internal System.Windows.Forms.Button btnOK;
        internal System.Windows.Forms.PictureBox PictureBox7;
        internal System.Windows.Forms.CheckBox CB_KeepFormatting;
        internal CustomGroupBox CustomGroupBox5;
        internal CustomPanel CP_InputRng;
        internal CustomPanel CP_OutputRng;
        internal CustomGroupBox CustomGroupBox6;
    }
}