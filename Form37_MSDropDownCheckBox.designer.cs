using System;

namespace VSTO_Addins
{
    [Microsoft.VisualBasic.CompilerServices.DesignerGenerated()]
    public partial class Form37_MSDropDownCheckBox : System.Windows.Forms.Form
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
            var resources = new System.ComponentModel.ComponentResourceManager(typeof(Form37_MSDropDownCheckBox));
            CB_Separator = new System.Windows.Forms.ComboBox();
            CB_Separator.SelectedIndexChanged += new EventHandler(CB_Separator_SelectedIndexChanged);
            CB_Separator.KeyUp += new System.Windows.Forms.KeyEventHandler(CB_Separator_KeyUp);
            Label3 = new System.Windows.Forms.Label();
            Btn_OK = new System.Windows.Forms.Button();
            Btn_OK.Click += new EventHandler(Btn_OK_Click);
            Btn_Cancel = new System.Windows.Forms.Button();
            Btn_Cancel.Click += new EventHandler(Btn_Cancel_Click);
            CB_About = new System.Windows.Forms.ComboBox();
            CB_Search = new System.Windows.Forms.CheckBox();
            GB_Sample = new CustomGroupBox();
            CustomGroupBox1 = new CustomGroupBox();
            CustomGroupBox3 = new CustomGroupBox();
            PictureBox4 = new System.Windows.Forms.PictureBox();
            RB_Vertical = new System.Windows.Forms.RadioButton();
            PictureBox5 = new System.Windows.Forms.PictureBox();
            RB_Horizontal = new System.Windows.Forms.RadioButton();
            CustomGroupBox4 = new CustomGroupBox();
            CustomGroupBox6 = new CustomGroupBox();
            Selection_source = new System.Windows.Forms.PictureBox();
            Selection_source.Click += new EventHandler(Selection_source_Click);
            Label1 = new System.Windows.Forms.Label();
            TB_src_rng = new System.Windows.Forms.TextBox();
            TB_src_rng.TextChanged += new EventHandler(TB_src_rng_TextChanged);
            CB_Source = new System.Windows.Forms.ComboBox();
            CB_Source.SelectedIndexChanged += new EventHandler(CB_Source_SelectedIndexChanged);
            CustomGroupBox1.SuspendLayout();
            CustomGroupBox3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)PictureBox4).BeginInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox5).BeginInit();
            CustomGroupBox4.SuspendLayout();
            CustomGroupBox6.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)Selection_source).BeginInit();
            SuspendLayout();
            // 
            // CB_Separator
            // 
            CB_Separator.Font = new System.Drawing.Font("Segoe UI", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            CB_Separator.FormattingEnabled = true;
            CB_Separator.Location = new System.Drawing.Point(15, 275);
            CB_Separator.Name = "CB_Separator";
            CB_Separator.Size = new System.Drawing.Size(273, 25);
            CB_Separator.TabIndex = 413;
            CB_Separator.Text = ",";
            // 
            // Label3
            // 
            Label3.AutoSize = true;
            Label3.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            Label3.Location = new System.Drawing.Point(15, 248);
            Label3.Name = "Label3";
            Label3.Size = new System.Drawing.Size(113, 17);
            Label3.TabIndex = 412;
            Label3.Text = "Select Separator :";
            // 
            // Btn_OK
            // 
            Btn_OK.BackColor = System.Drawing.Color.White;
            Btn_OK.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            Btn_OK.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            Btn_OK.Location = new System.Drawing.Point(448, 339);
            Btn_OK.Name = "Btn_OK";
            Btn_OK.Size = new System.Drawing.Size(62, 26);
            Btn_OK.TabIndex = 408;
            Btn_OK.Text = "OK";
            Btn_OK.UseVisualStyleBackColor = false;
            // 
            // Btn_Cancel
            // 
            Btn_Cancel.BackColor = System.Drawing.Color.White;
            Btn_Cancel.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            Btn_Cancel.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            Btn_Cancel.Location = new System.Drawing.Point(525, 337);
            Btn_Cancel.Name = "Btn_Cancel";
            Btn_Cancel.Size = new System.Drawing.Size(62, 26);
            Btn_Cancel.TabIndex = 407;
            Btn_Cancel.Text = "Cancel";
            Btn_Cancel.UseVisualStyleBackColor = false;
            // 
            // CB_About
            // 
            CB_About.Font = new System.Drawing.Font("Segoe UI", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            CB_About.FormattingEnabled = true;
            CB_About.Location = new System.Drawing.Point(15, 342);
            CB_About.Name = "CB_About";
            CB_About.Size = new System.Drawing.Size(154, 25);
            CB_About.TabIndex = 405;
            CB_About.Text = "SOFTEKO";
            // 
            // CB_Search
            // 
            CB_Search.AutoSize = true;
            CB_Search.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            CB_Search.Location = new System.Drawing.Point(15, 310);
            CB_Search.Name = "CB_Search";
            CB_Search.Size = new System.Drawing.Size(144, 21);
            CB_Search.TabIndex = 410;
            CB_Search.Text = "Keep Search Option";
            CB_Search.UseVisualStyleBackColor = true;
            // 
            // GB_Sample
            // 
            GB_Sample.BackColor = System.Drawing.Color.White;
            GB_Sample.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            GB_Sample.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 0);
            GB_Sample.Location = new System.Drawing.Point(329, 15);
            GB_Sample.Name = "GB_Sample";
            GB_Sample.Size = new System.Drawing.Size(261, 300);
            GB_Sample.TabIndex = 406;
            GB_Sample.TabStop = false;
            GB_Sample.Text = "Sample Image";
            // 
            // CustomGroupBox1
            // 
            CustomGroupBox1.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            CustomGroupBox1.Controls.Add(CustomGroupBox3);
            CustomGroupBox1.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 0);
            CustomGroupBox1.Location = new System.Drawing.Point(15, 156);
            CustomGroupBox1.Name = "CustomGroupBox1";
            CustomGroupBox1.Size = new System.Drawing.Size(274, 84);
            CustomGroupBox1.TabIndex = 409;
            CustomGroupBox1.TabStop = false;
            CustomGroupBox1.Text = "List Type";
            // 
            // CustomGroupBox3
            // 
            CustomGroupBox3.BackColor = System.Drawing.Color.White;
            CustomGroupBox3.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            CustomGroupBox3.Controls.Add(PictureBox4);
            CustomGroupBox3.Controls.Add(RB_Vertical);
            CustomGroupBox3.Controls.Add(PictureBox5);
            CustomGroupBox3.Controls.Add(RB_Horizontal);
            CustomGroupBox3.Location = new System.Drawing.Point(1, 22);
            CustomGroupBox3.Name = "CustomGroupBox3";
            CustomGroupBox3.Size = new System.Drawing.Size(272, 62);
            CustomGroupBox3.TabIndex = 0;
            CustomGroupBox3.TabStop = false;
            // 
            // PictureBox4
            // 
            PictureBox4.Image = (System.Drawing.Image)resources.GetObject("PictureBox4.Image");
            PictureBox4.Location = new System.Drawing.Point(233, 34);
            PictureBox4.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            PictureBox4.Name = "PictureBox4";
            PictureBox4.Size = new System.Drawing.Size(20, 20);
            PictureBox4.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            PictureBox4.TabIndex = 377;
            PictureBox4.TabStop = false;
            // 
            // RB_Vertical
            // 
            RB_Vertical.AutoSize = true;
            RB_Vertical.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            RB_Vertical.Location = new System.Drawing.Point(8, 33);
            RB_Vertical.Name = "RB_Vertical";
            RB_Vertical.Size = new System.Drawing.Size(77, 21);
            RB_Vertical.TabIndex = 1;
            RB_Vertical.Text = "Vertically";
            RB_Vertical.UseVisualStyleBackColor = true;
            // 
            // PictureBox5
            // 
            PictureBox5.Image = (System.Drawing.Image)resources.GetObject("PictureBox5.Image");
            PictureBox5.Location = new System.Drawing.Point(233, 8);
            PictureBox5.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            PictureBox5.Name = "PictureBox5";
            PictureBox5.Size = new System.Drawing.Size(20, 20);
            PictureBox5.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            PictureBox5.TabIndex = 376;
            PictureBox5.TabStop = false;
            // 
            // RB_Horizontal
            // 
            RB_Horizontal.AutoSize = true;
            RB_Horizontal.Checked = true;
            RB_Horizontal.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            RB_Horizontal.Location = new System.Drawing.Point(8, 7);
            RB_Horizontal.Name = "RB_Horizontal";
            RB_Horizontal.Size = new System.Drawing.Size(95, 21);
            RB_Horizontal.TabIndex = 0;
            RB_Horizontal.TabStop = true;
            RB_Horizontal.Text = "Horizontally";
            RB_Horizontal.UseVisualStyleBackColor = true;
            // 
            // CustomGroupBox4
            // 
            CustomGroupBox4.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            CustomGroupBox4.Controls.Add(CustomGroupBox6);
            CustomGroupBox4.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 0);
            CustomGroupBox4.Location = new System.Drawing.Point(15, 15);
            CustomGroupBox4.Name = "CustomGroupBox4";
            CustomGroupBox4.Size = new System.Drawing.Size(274, 132);
            CustomGroupBox4.TabIndex = 411;
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
            CustomGroupBox6.Size = new System.Drawing.Size(272, 110);
            CustomGroupBox6.TabIndex = 0;
            CustomGroupBox6.TabStop = false;
            // 
            // Selection_source
            // 
            Selection_source.BackColor = System.Drawing.Color.White;
            Selection_source.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            Selection_source.Image = (System.Drawing.Image)resources.GetObject("Selection_source.Image");
            Selection_source.Location = new System.Drawing.Point(233, 72);
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
            TB_src_rng.Size = new System.Drawing.Size(245, 25);
            TB_src_rng.TabIndex = 403;
            // 
            // CB_Source
            // 
            CB_Source.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            CB_Source.FormattingEnabled = true;
            CB_Source.Location = new System.Drawing.Point(12, 13);
            CB_Source.Name = "CB_Source";
            CB_Source.Size = new System.Drawing.Size(245, 25);
            CB_Source.TabIndex = 378;
            // 
            // Form37_MSDropDownCheckBox
            // 
            AutoScaleDimensions = new System.Drawing.SizeF(6.0f, 13.0f);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            ClientSize = new System.Drawing.Size(613, 393);
            Controls.Add(GB_Sample);
            Controls.Add(CustomGroupBox1);
            Controls.Add(CustomGroupBox4);
            Controls.Add(CB_Separator);
            Controls.Add(Label3);
            Controls.Add(Btn_OK);
            Controls.Add(Btn_Cancel);
            Controls.Add(CB_About);
            Controls.Add(CB_Search);
            HelpButton = true;
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "Form37_MSDropDownCheckBox";
            Text = "Drop down List with Checkbox";
            CustomGroupBox1.ResumeLayout(false);
            CustomGroupBox3.ResumeLayout(false);
            CustomGroupBox3.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)PictureBox4).EndInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox5).EndInit();
            CustomGroupBox4.ResumeLayout(false);
            CustomGroupBox6.ResumeLayout(false);
            CustomGroupBox6.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)Selection_source).EndInit();
            HelpButtonClicked += new System.ComponentModel.CancelEventHandler(Form1_HelpButtonClicked);
            Load += new EventHandler(YourForm_Load);
            KeyDown += new System.Windows.Forms.KeyEventHandler(Form37_MSDropDownCheckBox_KeyDown);
            Activated += new EventHandler(Form37_MSDropDownCheckBox_Activated);
            Shown += new EventHandler(Form37_MSDropDownCheckBox_Shown);
            Closing += new System.ComponentModel.CancelEventHandler(Form37_MSDropDownCheckBox_Closing);
            Disposed += new EventHandler(Form37_MSDropDownCheckBox_Disposed);
            ResumeLayout(false);
            PerformLayout();

        }

        internal CustomGroupBox GB_Sample;
        internal System.Windows.Forms.PictureBox PictureBox4;
        internal System.Windows.Forms.RadioButton RB_Vertical;
        internal System.Windows.Forms.PictureBox PictureBox5;
        internal CustomGroupBox CustomGroupBox3;
        internal System.Windows.Forms.RadioButton RB_Horizontal;
        internal CustomGroupBox CustomGroupBox1;
        internal System.Windows.Forms.PictureBox Selection_source;
        internal System.Windows.Forms.Label Label1;
        internal CustomGroupBox CustomGroupBox4;
        internal CustomGroupBox CustomGroupBox6;
        internal System.Windows.Forms.TextBox TB_src_rng;
        internal System.Windows.Forms.ComboBox CB_Source;
        internal System.Windows.Forms.ComboBox CB_Separator;
        internal System.Windows.Forms.Label Label3;
        internal System.Windows.Forms.Button Btn_OK;
        internal System.Windows.Forms.Button Btn_Cancel;
        internal System.Windows.Forms.ComboBox CB_About;
        internal System.Windows.Forms.CheckBox CB_Search;
    }
}