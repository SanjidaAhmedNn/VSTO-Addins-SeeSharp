using System;

namespace VSTO_Addins
{
    [Microsoft.VisualBasic.CompilerServices.DesignerGenerated()]
    public partial class Form42 : System.Windows.Forms.Form
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
            Label1 = new System.Windows.Forms.Label();
            Btn_OK = new System.Windows.Forms.Button();
            Btn_OK.Click += new EventHandler(Btn_OK_Click);
            Btn_Cancel = new System.Windows.Forms.Button();
            Btn_Cancel.Click += new EventHandler(Btn_Cancel_Click);
            CB_About = new System.Windows.Forms.ComboBox();
            CustomGroupBox3 = new CustomGroupBox();
            CGB = new CustomGroupBox();
            RB_Simple = new System.Windows.Forms.RadioButton();
            RB_Simple.CheckedChanged += new EventHandler(RB_Simple_CheckedChanged);
            RB_Dynamic = new System.Windows.Forms.RadioButton();
            RB_No = new System.Windows.Forms.RadioButton();
            RB_No.CheckedChanged += new EventHandler(RadioButton5_CheckedChanged);
            RB_Yes = new System.Windows.Forms.RadioButton();
            RB_Yes.CheckedChanged += new EventHandler(RB_Yes_CheckedChanged);
            CheckBox1 = new System.Windows.Forms.CheckBox();
            CheckBox1.CheckedChanged += new EventHandler(CheckBox1_CheckedChanged);
            CustomGroupBox3.SuspendLayout();
            CGB.SuspendLayout();
            SuspendLayout();
            // 
            // Label1
            // 
            Label1.BackColor = System.Drawing.Color.FromArgb(224, 224, 224);
            Label1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            Label1.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            Label1.Font = new System.Drawing.Font("Segoe UI", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            Label1.Location = new System.Drawing.Point(12, 13);
            Label1.Name = "Label1";
            Label1.Size = new System.Drawing.Size(285, 63);
            Label1.TabIndex = 0;
            Label1.Text = "Your current selection does not contain any Data Validation List.         " + '\r' + '\n' + "Do yo" + "u want to create a Data Validation List?";
            Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // Btn_OK
            // 
            Btn_OK.BackColor = System.Drawing.Color.White;
            Btn_OK.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            Btn_OK.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            Btn_OK.Location = new System.Drawing.Point(167, 248);
            Btn_OK.Name = "Btn_OK";
            Btn_OK.Size = new System.Drawing.Size(62, 26);
            Btn_OK.TabIndex = 424;
            Btn_OK.Text = "OK";
            Btn_OK.UseVisualStyleBackColor = false;
            // 
            // Btn_Cancel
            // 
            Btn_Cancel.BackColor = System.Drawing.Color.White;
            Btn_Cancel.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            Btn_Cancel.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            Btn_Cancel.Location = new System.Drawing.Point(235, 248);
            Btn_Cancel.Name = "Btn_Cancel";
            Btn_Cancel.Size = new System.Drawing.Size(62, 26);
            Btn_Cancel.TabIndex = 423;
            Btn_Cancel.Text = "Cancel";
            Btn_Cancel.UseVisualStyleBackColor = false;
            // 
            // CB_About
            // 
            CB_About.Font = new System.Drawing.Font("Segoe UI", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            CB_About.FormattingEnabled = true;
            CB_About.Location = new System.Drawing.Point(12, 248);
            CB_About.Name = "CB_About";
            CB_About.Size = new System.Drawing.Size(98, 25);
            CB_About.TabIndex = 422;
            CB_About.Text = "SOFTEKO";
            // 
            // CustomGroupBox3
            // 
            CustomGroupBox3.BackColor = System.Drawing.Color.White;
            CustomGroupBox3.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            CustomGroupBox3.Controls.Add(CGB);
            CustomGroupBox3.Controls.Add(RB_No);
            CustomGroupBox3.Controls.Add(RB_Yes);
            CustomGroupBox3.Location = new System.Drawing.Point(12, 90);
            CustomGroupBox3.Name = "CustomGroupBox3";
            CustomGroupBox3.Size = new System.Drawing.Size(285, 119);
            CustomGroupBox3.TabIndex = 427;
            CustomGroupBox3.TabStop = false;
            // 
            // CGB
            // 
            CGB.BorderColor = System.Drawing.Color.White;
            CGB.Controls.Add(RB_Simple);
            CGB.Controls.Add(RB_Dynamic);
            CGB.Location = new System.Drawing.Point(25, 32);
            CGB.Name = "CGB";
            CGB.Size = new System.Drawing.Size(194, 53);
            CGB.TabIndex = 431;
            CGB.TabStop = false;
            // 
            // RB_Simple
            // 
            RB_Simple.AutoSize = true;
            RB_Simple.Checked = true;
            RB_Simple.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            RB_Simple.Location = new System.Drawing.Point(6, 3);
            RB_Simple.Name = "RB_Simple";
            RB_Simple.Size = new System.Drawing.Size(155, 21);
            RB_Simple.TabIndex = 429;
            RB_Simple.TabStop = true;
            RB_Simple.Text = "Simple drop-down list";
            RB_Simple.UseVisualStyleBackColor = true;
            // 
            // RB_Dynamic
            // 
            RB_Dynamic.AutoSize = true;
            RB_Dynamic.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            RB_Dynamic.Location = new System.Drawing.Point(6, 29);
            RB_Dynamic.Name = "RB_Dynamic";
            RB_Dynamic.Size = new System.Drawing.Size(165, 21);
            RB_Dynamic.TabIndex = 430;
            RB_Dynamic.Text = "Dynamic drop-down list";
            RB_Dynamic.UseVisualStyleBackColor = true;
            // 
            // RB_No
            // 
            RB_No.AutoSize = true;
            RB_No.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            RB_No.Location = new System.Drawing.Point(11, 87);
            RB_No.Name = "RB_No";
            RB_No.Size = new System.Drawing.Size(211, 21);
            RB_No.TabIndex = 428;
            RB_No.Text = "No, I have a Data Validation List";
            RB_No.UseVisualStyleBackColor = true;
            // 
            // RB_Yes
            // 
            RB_Yes.AutoSize = true;
            RB_Yes.Checked = true;
            RB_Yes.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            RB_Yes.Location = new System.Drawing.Point(11, 11);
            RB_Yes.Name = "RB_Yes";
            RB_Yes.Size = new System.Drawing.Size(268, 21);
            RB_Yes.TabIndex = 427;
            RB_Yes.TabStop = true;
            RB_Yes.Text = "Yes, I want to create a Data Validation List";
            RB_Yes.UseVisualStyleBackColor = true;
            // 
            // CheckBox1
            // 
            CheckBox1.AutoSize = true;
            CheckBox1.Font = new System.Drawing.Font("Segoe UI", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            CheckBox1.Location = new System.Drawing.Point(12, 215);
            CheckBox1.Name = "CheckBox1";
            CheckBox1.Size = new System.Drawing.Size(253, 21);
            CheckBox1.TabIndex = 428;
            CheckBox1.Text = "Don't show this for this current session";
            CheckBox1.UseVisualStyleBackColor = true;
            // 
            // Form42
            // 
            AutoScaleDimensions = new System.Drawing.SizeF(6.0f, 13.0f);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            ClientSize = new System.Drawing.Size(318, 286);
            Controls.Add(CheckBox1);
            Controls.Add(CustomGroupBox3);
            Controls.Add(Btn_OK);
            Controls.Add(Btn_Cancel);
            Controls.Add(CB_About);
            Controls.Add(Label1);
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "Form42";
            Text = "Softeko for Excel";
            CustomGroupBox3.ResumeLayout(false);
            CustomGroupBox3.PerformLayout();
            CGB.ResumeLayout(false);
            CGB.PerformLayout();
            Load += new EventHandler(Form42_Load);
            Closing += new System.ComponentModel.CancelEventHandler(Form42_Closing);
            Disposed += new EventHandler(Form42_Disposed);
            ResumeLayout(false);
            PerformLayout();

        }

        internal System.Windows.Forms.Label Label1;
        internal System.Windows.Forms.Button Btn_OK;
        internal System.Windows.Forms.Button Btn_Cancel;
        internal System.Windows.Forms.ComboBox CB_About;
        internal CustomGroupBox CustomGroupBox3;
        internal System.Windows.Forms.RadioButton RB_Dynamic;
        internal System.Windows.Forms.RadioButton RB_Simple;
        internal System.Windows.Forms.RadioButton RB_No;
        internal System.Windows.Forms.RadioButton RB_Yes;
        internal CustomGroupBox CGB;
        internal System.Windows.Forms.CheckBox CheckBox1;
    }
}