using System;

namespace VSTO_Addins
{
    [Microsoft.VisualBasic.CompilerServices.DesignerGenerated()]
    public partial class Form39_DropdownlistwithSearchOption : System.Windows.Forms.Form
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
            var resources = new System.ComponentModel.ComponentResourceManager(typeof(Form39_DropdownlistwithSearchOption));
            Btn_OK = new System.Windows.Forms.Button();
            Btn_OK.Click += new EventHandler(Btn_OK_Click);
            Btn_Cancel = new System.Windows.Forms.Button();
            Btn_Cancel.Click += new EventHandler(Btn_Cancel_Click);
            CB_About = new System.Windows.Forms.ComboBox();
            GB_Sample = new CustomGroupBox();
            GB_Sample.Enter += new EventHandler(GB_Sample_Enter);
            CustomGroupBox4 = new CustomGroupBox();
            CustomGroupBox6 = new CustomGroupBox();
            Selection_source = new System.Windows.Forms.PictureBox();
            Selection_source.Click += new EventHandler(Selection_source_Click);
            Label1 = new System.Windows.Forms.Label();
            TB_src_rng = new System.Windows.Forms.TextBox();
            TB_src_rng.TextChanged += new EventHandler(TB_src_rng_TextChanged);
            CB_Source = new System.Windows.Forms.ComboBox();
            CB_Source.SelectedIndexChanged += new EventHandler(CB_Source_SelectedIndexChanged);
            CustomGroupBox4.SuspendLayout();
            CustomGroupBox6.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)Selection_source).BeginInit();
            SuspendLayout();
            // 
            // Btn_OK
            // 
            Btn_OK.BackColor = System.Drawing.Color.White;
            Btn_OK.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            Btn_OK.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            Btn_OK.Location = new System.Drawing.Point(166, 485);
            Btn_OK.Name = "Btn_OK";
            Btn_OK.Size = new System.Drawing.Size(62, 26);
            Btn_OK.TabIndex = 416;
            Btn_OK.Text = "OK";
            Btn_OK.UseVisualStyleBackColor = false;
            // 
            // Btn_Cancel
            // 
            Btn_Cancel.BackColor = System.Drawing.Color.White;
            Btn_Cancel.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            Btn_Cancel.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            Btn_Cancel.Location = new System.Drawing.Point(241, 485);
            Btn_Cancel.Name = "Btn_Cancel";
            Btn_Cancel.Size = new System.Drawing.Size(62, 26);
            Btn_Cancel.TabIndex = 415;
            Btn_Cancel.Text = "Cancel";
            Btn_Cancel.UseVisualStyleBackColor = false;
            // 
            // CB_About
            // 
            CB_About.Font = new System.Drawing.Font("Segoe UI", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            CB_About.FormattingEnabled = true;
            CB_About.Location = new System.Drawing.Point(16, 487);
            CB_About.Name = "CB_About";
            CB_About.Size = new System.Drawing.Size(98, 25);
            CB_About.TabIndex = 413;
            CB_About.Text = "SOFTEKO";
            // 
            // GB_Sample
            // 
            GB_Sample.BackColor = System.Drawing.Color.White;
            GB_Sample.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            GB_Sample.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 0);
            GB_Sample.Location = new System.Drawing.Point(15, 160);
            GB_Sample.Name = "GB_Sample";
            GB_Sample.Size = new System.Drawing.Size(288, 299);
            GB_Sample.TabIndex = 414;
            GB_Sample.TabStop = false;
            GB_Sample.Text = "Sample Image";
            // 
            // CustomGroupBox4
            // 
            CustomGroupBox4.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            CustomGroupBox4.Controls.Add(CustomGroupBox6);
            CustomGroupBox4.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 0);
            CustomGroupBox4.Location = new System.Drawing.Point(15, 15);
            CustomGroupBox4.Name = "CustomGroupBox4";
            CustomGroupBox4.Size = new System.Drawing.Size(288, 132);
            CustomGroupBox4.TabIndex = 412;
            CustomGroupBox4.TabStop = false;
            CustomGroupBox4.Text = "Data Validation Range";
            // 
            // CustomGroupBox6
            // 
            CustomGroupBox6.BackColor = System.Drawing.Color.White;
            CustomGroupBox6.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            CustomGroupBox6.Controls.Add(Selection_source);
            CustomGroupBox6.Controls.Add(Label1);
            CustomGroupBox6.Controls.Add(TB_src_rng);
            CustomGroupBox6.Controls.Add(CB_Source);
            CustomGroupBox6.Location = new System.Drawing.Point(1, 22);
            CustomGroupBox6.Name = "CustomGroupBox6";
            CustomGroupBox6.Size = new System.Drawing.Size(287, 110);
            CustomGroupBox6.TabIndex = 0;
            CustomGroupBox6.TabStop = false;
            // 
            // Selection_source
            // 
            Selection_source.BackColor = System.Drawing.Color.White;
            Selection_source.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            Selection_source.Image = (System.Drawing.Image)resources.GetObject("Selection_source.Image");
            Selection_source.Location = new System.Drawing.Point(243, 72);
            Selection_source.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            Selection_source.Name = "Selection_source";
            Selection_source.Size = new System.Drawing.Size(24, 25);
            Selection_source.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            Selection_source.TabIndex = 404;
            Selection_source.TabStop = false;
            // 
            // Label1
            // 
            Label1.AutoSize = true;
            Label1.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            Label1.Location = new System.Drawing.Point(12, 46);
            Label1.Name = "Label1";
            Label1.Size = new System.Drawing.Size(116, 17);
            Label1.TabIndex = 379;
            Label1.Text = "Define the range :";
            // 
            // TB_src_rng
            // 
            TB_src_rng.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            TB_src_rng.Cursor = System.Windows.Forms.Cursors.IBeam;
            TB_src_rng.Font = new System.Drawing.Font("Segoe UI", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            TB_src_rng.Location = new System.Drawing.Point(12, 72);
            TB_src_rng.Name = "TB_src_rng";
            TB_src_rng.Size = new System.Drawing.Size(255, 25);
            TB_src_rng.TabIndex = 403;
            // 
            // CB_Source
            // 
            CB_Source.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            CB_Source.FormattingEnabled = true;
            CB_Source.Location = new System.Drawing.Point(12, 13);
            CB_Source.Name = "CB_Source";
            CB_Source.Size = new System.Drawing.Size(255, 25);
            CB_Source.TabIndex = 378;
            // 
            // Form39_DropdownlistwithSearchOption
            // 
            AutoScaleDimensions = new System.Drawing.SizeF(6.0f, 13.0f);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            ClientSize = new System.Drawing.Size(326, 527);
            Controls.Add(GB_Sample);
            Controls.Add(Btn_OK);
            Controls.Add(Btn_Cancel);
            Controls.Add(CB_About);
            Controls.Add(CustomGroupBox4);
            HelpButton = true;
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "Form39_DropdownlistwithSearchOption";
            Text = "Drop-down List with Search Option";
            CustomGroupBox4.ResumeLayout(false);
            CustomGroupBox6.ResumeLayout(false);
            CustomGroupBox6.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)Selection_source).EndInit();
            HelpButtonClicked += new System.ComponentModel.CancelEventHandler(Form1_HelpButtonClicked);
            Load += new EventHandler(Form39_DropdownlistwithSearchOption_Load);
            KeyDown += new System.Windows.Forms.KeyEventHandler(Form39_DropdownlistwithSearchOption_KeyDown);
            Shown += new EventHandler(Form39_Shown);
            Closing += new System.ComponentModel.CancelEventHandler(Form39_DropdownlistwithSearchOption_Closing);
            Disposed += new EventHandler(Form39_DropdownlistwithSearchOption_Disposed);
            ResumeLayout(false);

        }

        internal CustomGroupBox CustomGroupBox4;
        internal CustomGroupBox CustomGroupBox6;
        internal System.Windows.Forms.PictureBox Selection_source;
        internal System.Windows.Forms.Label Label1;
        internal System.Windows.Forms.TextBox TB_src_rng;
        internal System.Windows.Forms.ComboBox CB_Source;
        internal CustomGroupBox GB_Sample;
        internal System.Windows.Forms.Button Btn_OK;
        internal System.Windows.Forms.Button Btn_Cancel;
        internal System.Windows.Forms.ComboBox CB_About;
    }
}