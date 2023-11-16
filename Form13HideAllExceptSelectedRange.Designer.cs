using System;

namespace VSTO_Addins
{
    [Microsoft.VisualBasic.CompilerServices.DesignerGenerated()]
    public partial class Form13HideAllExceptSelectedRange : System.Windows.Forms.Form
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
            var resources = new System.ComponentModel.ComponentResourceManager(typeof(Form13HideAllExceptSelectedRange));
            pctBoxSelectRange = new System.Windows.Forms.PictureBox();
            pctBoxSelectRange.Click += new EventHandler(pctBoxSelectRange_Click);
            btnOK = new System.Windows.Forms.Button();
            btnOK.Click += new EventHandler(btnOK_Click);
            ComboBox1 = new System.Windows.Forms.ComboBox();
            checkBoxCopyWorksheet = new System.Windows.Forms.CheckBox();
            txtSourceRange = new System.Windows.Forms.TextBox();
            txtSourceRange.TextChanged += new EventHandler(txtSourceRange_TextChanged);
            txtSourceRange.GotFocus += new EventHandler(txtSourceRange_GotFocus);
            Label1 = new System.Windows.Forms.Label();
            btnCancel = new System.Windows.Forms.Button();
            btnCancel.Click += new EventHandler(btnCancel_Click);
            checkBox_Header = new System.Windows.Forms.CheckBox();
            GB_sample = new CustomGroupBox();
            ((System.ComponentModel.ISupportInitialize)pctBoxSelectRange).BeginInit();
            SuspendLayout();
            // 
            // pctBoxSelectRange
            // 
            pctBoxSelectRange.BackColor = System.Drawing.Color.White;
            pctBoxSelectRange.Image = (System.Drawing.Image)resources.GetObject("pctBoxSelectRange.Image");
            pctBoxSelectRange.Location = new System.Drawing.Point(482, 13);
            pctBoxSelectRange.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            pctBoxSelectRange.Name = "pctBoxSelectRange";
            pctBoxSelectRange.Size = new System.Drawing.Size(24, 23);
            pctBoxSelectRange.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            pctBoxSelectRange.TabIndex = 191;
            pctBoxSelectRange.TabStop = false;
            // 
            // btnOK
            // 
            btnOK.BackColor = System.Drawing.Color.White;
            btnOK.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            btnOK.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            btnOK.Location = new System.Drawing.Point(367, 255);
            btnOK.Name = "btnOK";
            btnOK.Size = new System.Drawing.Size(62, 26);
            btnOK.TabIndex = 190;
            btnOK.Text = "OK";
            btnOK.UseVisualStyleBackColor = false;
            // 
            // ComboBox1
            // 
            ComboBox1.Font = new System.Drawing.Font("Segoe UI", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            ComboBox1.FormattingEnabled = true;
            ComboBox1.Location = new System.Drawing.Point(15, 257);
            ComboBox1.Name = "ComboBox1";
            ComboBox1.Size = new System.Drawing.Size(90, 25);
            ComboBox1.TabIndex = 186;
            ComboBox1.Text = "SOFTEKO";
            // 
            // checkBoxCopyWorksheet
            // 
            checkBoxCopyWorksheet.AutoSize = true;
            checkBoxCopyWorksheet.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            checkBoxCopyWorksheet.Location = new System.Drawing.Point(15, 227);
            checkBoxCopyWorksheet.Name = "checkBoxCopyWorksheet";
            checkBoxCopyWorksheet.Size = new System.Drawing.Size(257, 21);
            checkBoxCopyWorksheet.TabIndex = 185;
            checkBoxCopyWorksheet.Text = "Create a copy of the original worksheet";
            checkBoxCopyWorksheet.UseVisualStyleBackColor = true;
            // 
            // txtSourceRange
            // 
            txtSourceRange.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            txtSourceRange.Cursor = System.Windows.Forms.Cursors.IBeam;
            txtSourceRange.Font = new System.Drawing.Font("Segoe UI", 10.0f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            txtSourceRange.Location = new System.Drawing.Point(119, 12);
            txtSourceRange.Name = "txtSourceRange";
            txtSourceRange.Size = new System.Drawing.Size(388, 25);
            txtSourceRange.TabIndex = 183;
            // 
            // Label1
            // 
            Label1.AutoSize = true;
            Label1.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            Label1.Location = new System.Drawing.Point(15, 15);
            Label1.Name = "Label1";
            Label1.Size = new System.Drawing.Size(98, 17);
            Label1.TabIndex = 182;
            Label1.Text = "Source Range :";
            // 
            // btnCancel
            // 
            btnCancel.BackColor = System.Drawing.Color.White;
            btnCancel.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            btnCancel.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            btnCancel.Location = new System.Drawing.Point(444, 255);
            btnCancel.Name = "btnCancel";
            btnCancel.Size = new System.Drawing.Size(62, 26);
            btnCancel.TabIndex = 189;
            btnCancel.Text = "Cancel";
            btnCancel.UseVisualStyleBackColor = false;
            // 
            // checkBox_Header
            // 
            checkBox_Header.AutoSize = true;
            checkBox_Header.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            checkBox_Header.Location = new System.Drawing.Point(15, 46);
            checkBox_Header.Name = "checkBox_Header";
            checkBox_Header.Size = new System.Drawing.Size(194, 21);
            checkBox_Header.TabIndex = 195;
            checkBox_Header.Text = "I have headers in my dataset";
            checkBox_Header.UseVisualStyleBackColor = true;
            // 
            // GB_sample
            // 
            GB_sample.BackColor = System.Drawing.Color.White;
            GB_sample.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            GB_sample.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 0);
            GB_sample.Location = new System.Drawing.Point(15, 82);
            GB_sample.Name = "GB_sample";
            GB_sample.Size = new System.Drawing.Size(492, 130);
            GB_sample.TabIndex = 400;
            GB_sample.TabStop = false;
            GB_sample.Text = "Sample Image";
            // 
            // Form13HideAllExceptSelectedRange
            // 
            AutoScaleDimensions = new System.Drawing.SizeF(6.0f, 13.0f);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            ClientSize = new System.Drawing.Size(526, 296);
            Controls.Add(GB_sample);
            Controls.Add(checkBox_Header);
            Controls.Add(pctBoxSelectRange);
            Controls.Add(btnOK);
            Controls.Add(ComboBox1);
            Controls.Add(checkBoxCopyWorksheet);
            Controls.Add(txtSourceRange);
            Controls.Add(Label1);
            Controls.Add(btnCancel);
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "Form13HideAllExceptSelectedRange";
            Text = "Hide All Except the Selected Range";
            ((System.ComponentModel.ISupportInitialize)pctBoxSelectRange).EndInit();
            KeyDown += new System.Windows.Forms.KeyEventHandler(Form1_KeyDown);
            Load += new EventHandler(Form13HideAllExceptSelectedRange_Load);
            Activated += new EventHandler(Form1_Activated);
            Closing += new System.ComponentModel.CancelEventHandler(Form13HideAllExceptSelectedRange_Closing);
            Disposed += new EventHandler(Form13HideAllExceptSelectedRange_Disposed);
            Shown += new EventHandler(Form13HideAllExceptSelectedRange_Shown);
            ResumeLayout(false);
            PerformLayout();

        }
        internal System.Windows.Forms.PictureBox pctBoxSelectRange;
        internal System.Windows.Forms.Button btnOK;
        internal System.Windows.Forms.ComboBox ComboBox1;
        internal System.Windows.Forms.CheckBox checkBoxCopyWorksheet;
        internal System.Windows.Forms.TextBox txtSourceRange;
        internal System.Windows.Forms.Label Label1;
        internal System.Windows.Forms.Button btnCancel;
        internal System.Windows.Forms.CheckBox checkBox_Header;
        internal CustomGroupBox GB_sample;
    }
}