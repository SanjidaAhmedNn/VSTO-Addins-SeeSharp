using System;

namespace VSTO_Addins
{
    [Microsoft.VisualBasic.CompilerServices.DesignerGenerated()]
    public partial class Form14SpecifyScrollArea : System.Windows.Forms.Form
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
            var resources = new System.ComponentModel.ComponentResourceManager(typeof(Form14SpecifyScrollArea));
            Selection = new System.Windows.Forms.PictureBox();
            Selection.Click += new EventHandler(Selection_Click);
            Btn_OK = new System.Windows.Forms.Button();
            Btn_OK.Click += new EventHandler(Btn_OK_Click);
            ComboBox = new System.Windows.Forms.ComboBox();
            CheckBox = new System.Windows.Forms.CheckBox();
            txtSourceRange = new System.Windows.Forms.TextBox();
            txtSourceRange.TextChanged += new EventHandler(txtSourceRange_TextChanged);
            txtSourceRange.GotFocus += new EventHandler(txtSourceRange_GotFocus);
            Label1 = new System.Windows.Forms.Label();
            Btn_Cancel = new System.Windows.Forms.Button();
            Btn_Cancel.Click += new EventHandler(Btn_Cancel_Click);
            Info = new System.Windows.Forms.PictureBox();
            GB_sample = new CustomGroupBox();
            ((System.ComponentModel.ISupportInitialize)Selection).BeginInit();
            ((System.ComponentModel.ISupportInitialize)Info).BeginInit();
            SuspendLayout();
            // 
            // Selection
            // 
            Selection.BackColor = System.Drawing.Color.White;
            Selection.Image = (System.Drawing.Image)resources.GetObject("Selection.Image");
            Selection.Location = new System.Drawing.Point(447, 16);
            Selection.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            Selection.Name = "Selection";
            Selection.Size = new System.Drawing.Size(24, 23);
            Selection.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            Selection.TabIndex = 203;
            Selection.TabStop = false;
            // 
            // Btn_OK
            // 
            Btn_OK.BackColor = System.Drawing.Color.White;
            Btn_OK.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            Btn_OK.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            Btn_OK.Location = new System.Drawing.Point(371, 237);
            Btn_OK.Name = "Btn_OK";
            Btn_OK.Size = new System.Drawing.Size(62, 26);
            Btn_OK.TabIndex = 202;
            Btn_OK.Text = "OK";
            Btn_OK.UseVisualStyleBackColor = false;
            // 
            // ComboBox
            // 
            ComboBox.Font = new System.Drawing.Font("Segoe UI", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            ComboBox.FormattingEnabled = true;
            ComboBox.Location = new System.Drawing.Point(15, 239);
            ComboBox.Name = "ComboBox";
            ComboBox.Size = new System.Drawing.Size(90, 25);
            ComboBox.TabIndex = 198;
            ComboBox.Text = "SOFTEKO";
            // 
            // CheckBox
            // 
            CheckBox.AutoSize = true;
            CheckBox.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            CheckBox.Location = new System.Drawing.Point(15, 209);
            CheckBox.Name = "CheckBox";
            CheckBox.Size = new System.Drawing.Size(257, 21);
            CheckBox.TabIndex = 197;
            CheckBox.Text = "Create a copy of the original worksheet";
            CheckBox.UseVisualStyleBackColor = true;
            // 
            // txtSourceRange
            // 
            txtSourceRange.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            txtSourceRange.Cursor = System.Windows.Forms.Cursors.IBeam;
            txtSourceRange.Font = new System.Drawing.Font("Segoe UI", 10.0f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            txtSourceRange.Location = new System.Drawing.Point(120, 15);
            txtSourceRange.Name = "txtSourceRange";
            txtSourceRange.Size = new System.Drawing.Size(352, 25);
            txtSourceRange.TabIndex = 196;
            // 
            // Label1
            // 
            Label1.AutoSize = true;
            Label1.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            Label1.Location = new System.Drawing.Point(15, 18);
            Label1.Name = "Label1";
            Label1.Size = new System.Drawing.Size(98, 17);
            Label1.TabIndex = 195;
            Label1.Text = "Source Range :";
            // 
            // Btn_Cancel
            // 
            Btn_Cancel.BackColor = System.Drawing.Color.White;
            Btn_Cancel.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            Btn_Cancel.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            Btn_Cancel.Location = new System.Drawing.Point(448, 237);
            Btn_Cancel.Name = "Btn_Cancel";
            Btn_Cancel.Size = new System.Drawing.Size(62, 26);
            Btn_Cancel.TabIndex = 201;
            Btn_Cancel.Text = "Cancel";
            Btn_Cancel.UseVisualStyleBackColor = false;
            // 
            // Info
            // 
            Info.Image = (System.Drawing.Image)resources.GetObject("Info.Image");
            Info.Location = new System.Drawing.Point(483, 14);
            Info.Name = "Info";
            Info.Size = new System.Drawing.Size(26, 26);
            Info.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            Info.TabIndex = 205;
            Info.TabStop = false;
            // 
            // GB_sample
            // 
            GB_sample.BackColor = System.Drawing.Color.White;
            GB_sample.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            GB_sample.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 0);
            GB_sample.Location = new System.Drawing.Point(15, 60);
            GB_sample.Name = "GB_sample";
            GB_sample.Size = new System.Drawing.Size(494, 140);
            GB_sample.TabIndex = 400;
            GB_sample.TabStop = false;
            GB_sample.Text = "Sample Image";
            // 
            // Form14SpecifyScrollArea
            // 
            AutoScaleDimensions = new System.Drawing.SizeF(6.0f, 13.0f);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            ClientSize = new System.Drawing.Size(526, 279);
            Controls.Add(GB_sample);
            Controls.Add(Info);
            Controls.Add(Selection);
            Controls.Add(Btn_OK);
            Controls.Add(ComboBox);
            Controls.Add(CheckBox);
            Controls.Add(txtSourceRange);
            Controls.Add(Label1);
            Controls.Add(Btn_Cancel);
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "Form14SpecifyScrollArea";
            Text = "Specify Scroll Area";
            ((System.ComponentModel.ISupportInitialize)Selection).EndInit();
            ((System.ComponentModel.ISupportInitialize)Info).EndInit();
            KeyDown += new System.Windows.Forms.KeyEventHandler(Form1_KeyDown);
            Load += new EventHandler(Form14SpecifyScrollArea_Load);
            Activated += new EventHandler(Form1_Activated);
            Disposed += new EventHandler(Form14SpecifyScrollArea_Disposed);
            Closing += new System.ComponentModel.CancelEventHandler(Form14SpecifyScrollArea_Closing);
            Shown += new EventHandler(Form14SpecifyScrollArea_Shown);
            ResumeLayout(false);
            PerformLayout();

        }
        internal System.Windows.Forms.PictureBox Selection;
        internal System.Windows.Forms.Button Btn_OK;
        internal System.Windows.Forms.ComboBox ComboBox;
        internal System.Windows.Forms.CheckBox CheckBox;
        internal System.Windows.Forms.TextBox txtSourceRange;
        internal System.Windows.Forms.Label Label1;
        internal System.Windows.Forms.Button Btn_Cancel;
        internal System.Windows.Forms.PictureBox Info;
        internal CustomGroupBox GB_sample;
    }
}