using System;

namespace VSTO_Addins
{
    [Microsoft.VisualBasic.CompilerServices.DesignerGenerated()]
    public partial class Form41_RemoveAdavancedDropdownList : System.Windows.Forms.Form
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
            var resources = new System.ComponentModel.ComponentResourceManager(typeof(Form41_RemoveAdavancedDropdownList));
            Btn_OK = new System.Windows.Forms.Button();
            Btn_OK.Click += new EventHandler(Btn_OK_Click);
            Btn_Cancel = new System.Windows.Forms.Button();
            Btn_Cancel.Click += new EventHandler(Btn_Cancel_Click);
            CB_About = new System.Windows.Forms.ComboBox();
            CustomGroupBox2 = new CustomGroupBox();
            CustomGroupBox1 = new CustomGroupBox();
            CB_search = new System.Windows.Forms.CheckBox();
            CB_search.CheckedChanged += new EventHandler(CheckBox3_CheckedChanged);
            CB_checkbox = new System.Windows.Forms.CheckBox();
            CB_multiselect = new System.Windows.Forms.CheckBox();
            CB_multiselect.CheckedChanged += new EventHandler(CB_multiselect_CheckedChanged);
            CustomGroupBox4 = new CustomGroupBox();
            CustomGroupBox6 = new CustomGroupBox();
            Selection_source = new System.Windows.Forms.PictureBox();
            Selection_source.Click += new EventHandler(Selection_source_Click);
            Label1 = new System.Windows.Forms.Label();
            TB_src_rng = new System.Windows.Forms.TextBox();
            TB_src_rng.TextChanged += new EventHandler(TB_src_rng_TextChanged);
            CB_Source = new System.Windows.Forms.ComboBox();
            CustomGroupBox2.SuspendLayout();
            CustomGroupBox1.SuspendLayout();
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
            Btn_OK.Location = new System.Drawing.Point(164, 310);
            Btn_OK.Name = "Btn_OK";
            Btn_OK.Size = new System.Drawing.Size(62, 26);
            Btn_OK.TabIndex = 421;
            Btn_OK.Text = "OK";
            Btn_OK.UseVisualStyleBackColor = false;
            // 
            // Btn_Cancel
            // 
            Btn_Cancel.BackColor = System.Drawing.Color.White;
            Btn_Cancel.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            Btn_Cancel.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            Btn_Cancel.Location = new System.Drawing.Point(239, 310);
            Btn_Cancel.Name = "Btn_Cancel";
            Btn_Cancel.Size = new System.Drawing.Size(62, 26);
            Btn_Cancel.TabIndex = 420;
            Btn_Cancel.Text = "Cancel";
            Btn_Cancel.UseVisualStyleBackColor = false;
            // 
            // CB_About
            // 
            CB_About.Font = new System.Drawing.Font("Segoe UI", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            CB_About.FormattingEnabled = true;
            CB_About.Location = new System.Drawing.Point(14, 312);
            CB_About.Name = "CB_About";
            CB_About.Size = new System.Drawing.Size(98, 25);
            CB_About.TabIndex = 418;
            CB_About.Text = "SOFTEKO";
            // 
            // CustomGroupBox2
            // 
            CustomGroupBox2.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            CustomGroupBox2.Controls.Add(CustomGroupBox1);
            CustomGroupBox2.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 0);
            CustomGroupBox2.Location = new System.Drawing.Point(13, 164);
            CustomGroupBox2.Name = "CustomGroupBox2";
            CustomGroupBox2.Size = new System.Drawing.Size(288, 128);
            CustomGroupBox2.TabIndex = 422;
            CustomGroupBox2.TabStop = false;
            CustomGroupBox2.Text = "Data Validation List Type";
            // 
            // CustomGroupBox1
            // 
            CustomGroupBox1.BackColor = System.Drawing.Color.White;
            CustomGroupBox1.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            CustomGroupBox1.Controls.Add(CB_search);
            CustomGroupBox1.Controls.Add(CB_checkbox);
            CustomGroupBox1.Controls.Add(CB_multiselect);
            CustomGroupBox1.Location = new System.Drawing.Point(1, 22);
            CustomGroupBox1.Name = "CustomGroupBox1";
            CustomGroupBox1.Size = new System.Drawing.Size(287, 105);
            CustomGroupBox1.TabIndex = 0;
            CustomGroupBox1.TabStop = false;
            // 
            // CB_search
            // 
            CB_search.AutoSize = true;
            CB_search.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            CB_search.Location = new System.Drawing.Point(10, 70);
            CB_search.Name = "CB_search";
            CB_search.Size = new System.Drawing.Size(234, 21);
            CB_search.TabIndex = 2;
            CB_search.Text = "Drop-down list with search option";
            CB_search.UseVisualStyleBackColor = true;
            // 
            // CB_checkbox
            // 
            CB_checkbox.AutoSize = true;
            CB_checkbox.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            CB_checkbox.Location = new System.Drawing.Point(10, 41);
            CB_checkbox.Name = "CB_checkbox";
            CB_checkbox.Size = new System.Drawing.Size(250, 21);
            CB_checkbox.TabIndex = 1;
            CB_checkbox.Text = "Drop-down list containing check box";
            CB_checkbox.UseVisualStyleBackColor = true;
            // 
            // CB_multiselect
            // 
            CB_multiselect.AutoSize = true;
            CB_multiselect.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            CB_multiselect.Location = new System.Drawing.Point(10, 12);
            CB_multiselect.Name = "CB_multiselect";
            CB_multiselect.Size = new System.Drawing.Size(249, 21);
            CB_multiselect.TabIndex = 0;
            CB_multiselect.Text = "Multi selection-based drop-down list";
            CB_multiselect.UseVisualStyleBackColor = true;
            // 
            // CustomGroupBox4
            // 
            CustomGroupBox4.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            CustomGroupBox4.Controls.Add(CustomGroupBox6);
            CustomGroupBox4.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 0);
            CustomGroupBox4.Location = new System.Drawing.Point(12, 12);
            CustomGroupBox4.Name = "CustomGroupBox4";
            CustomGroupBox4.Size = new System.Drawing.Size(288, 132);
            CustomGroupBox4.TabIndex = 417;
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
            // Form41_RemoveAdavancedDropdownList
            // 
            AutoScaleDimensions = new System.Drawing.SizeF(6.0f, 13.0f);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            ClientSize = new System.Drawing.Size(321, 352);
            Controls.Add(CustomGroupBox2);
            Controls.Add(Btn_OK);
            Controls.Add(Btn_Cancel);
            Controls.Add(CB_About);
            Controls.Add(CustomGroupBox4);
            HelpButton = true;
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "Form41_RemoveAdavancedDropdownList";
            Text = "Remove Adavanced Dropdown List";
            CustomGroupBox2.ResumeLayout(false);
            CustomGroupBox1.ResumeLayout(false);
            CustomGroupBox1.PerformLayout();
            CustomGroupBox4.ResumeLayout(false);
            CustomGroupBox6.ResumeLayout(false);
            CustomGroupBox6.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)Selection_source).EndInit();
            Load += new EventHandler(Form41_RemoveAdavancedDropdownList_Load);
            KeyDown += new System.Windows.Forms.KeyEventHandler(Form41_RemoveAdavancedDropdownList_KeyDown);
            Closing += new System.ComponentModel.CancelEventHandler(Form41_RemoveAdavancedDropdownList_Closing);
            Disposed += new EventHandler(Form41_RemoveAdavancedDropdownList_Disposed);
            Shown += new EventHandler(Form41_RemoveAdavancedDropdownList_Shown);
            ResumeLayout(false);

        }
        internal System.Windows.Forms.Button Btn_OK;
        internal System.Windows.Forms.Button Btn_Cancel;
        internal System.Windows.Forms.ComboBox CB_About;
        internal CustomGroupBox CustomGroupBox4;
        internal CustomGroupBox CustomGroupBox6;
        internal System.Windows.Forms.PictureBox Selection_source;
        internal System.Windows.Forms.Label Label1;
        internal System.Windows.Forms.TextBox TB_src_rng;
        internal System.Windows.Forms.ComboBox CB_Source;
        internal CustomGroupBox CustomGroupBox1;
        internal System.Windows.Forms.CheckBox CB_search;
        internal System.Windows.Forms.CheckBox CB_checkbox;
        internal System.Windows.Forms.CheckBox CB_multiselect;
        internal CustomGroupBox CustomGroupBox2;
    }
}