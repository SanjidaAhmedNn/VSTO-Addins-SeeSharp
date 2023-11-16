using System;

namespace VSTO_Addins
{
    [Microsoft.VisualBasic.CompilerServices.DesignerGenerated()]
    public partial class Form12HideRanges : System.Windows.Forms.Form
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
            var resources = new System.ComponentModel.ComponentResourceManager(typeof(Form12HideRanges));
            Selection = new System.Windows.Forms.PictureBox();
            Selection.Click += new EventHandler(Selection_Click);
            btn_OK = new System.Windows.Forms.Button();
            btn_OK.Click += new EventHandler(btn_OK_Click);
            btn_Cancel = new System.Windows.Forms.Button();
            btn_Cancel.Click += new EventHandler(btn_Cancel_Click);
            ComboBox1 = new System.Windows.Forms.ComboBox();
            CheckBox1 = new System.Windows.Forms.CheckBox();
            AutoSelection = new System.Windows.Forms.PictureBox();
            AutoSelection.Click += new EventHandler(AutoSelection_Click);
            txtSourceRange = new System.Windows.Forms.TextBox();
            txtSourceRange.TextChanged += new EventHandler(txtSourceRange_TextChanged);
            txtSourceRange.GotFocus += new EventHandler(txtSourceRange_GotFocus);
            Label1 = new System.Windows.Forms.Label();
            GB_sample = new CustomGroupBox();
            CustomGroupBox3 = new CustomGroupBox();
            CustomGroupBox6 = new CustomGroupBox();
            RB_Single_Range = new System.Windows.Forms.RadioButton();
            RB_Multiple_Range = new System.Windows.Forms.RadioButton();
            CustomGroupBox4 = new CustomGroupBox();
            CustomGroupBox5 = new CustomGroupBox();
            RB_bidirection = new System.Windows.Forms.RadioButton();
            RB_Row = new System.Windows.Forms.RadioButton();
            RB_Column = new System.Windows.Forms.RadioButton();
            ((System.ComponentModel.ISupportInitialize)Selection).BeginInit();
            ((System.ComponentModel.ISupportInitialize)AutoSelection).BeginInit();
            CustomGroupBox3.SuspendLayout();
            CustomGroupBox6.SuspendLayout();
            CustomGroupBox4.SuspendLayout();
            CustomGroupBox5.SuspendLayout();
            SuspendLayout();
            // 
            // Selection
            // 
            Selection.BackColor = System.Drawing.Color.White;
            Selection.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            Selection.Image = (System.Drawing.Image)resources.GetObject("Selection.Image");
            Selection.Location = new System.Drawing.Point(240, 42);
            Selection.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            Selection.Name = "Selection";
            Selection.Size = new System.Drawing.Size(24, 25);
            Selection.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            Selection.TabIndex = 126;
            Selection.TabStop = false;
            // 
            // btn_OK
            // 
            btn_OK.BackColor = System.Drawing.Color.White;
            btn_OK.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            btn_OK.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            btn_OK.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            btn_OK.ForeColor = System.Drawing.Color.FromArgb(70, 70, 70);
            btn_OK.Location = new System.Drawing.Point(407, 338);
            btn_OK.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            btn_OK.Name = "btn_OK";
            btn_OK.Size = new System.Drawing.Size(62, 26);
            btn_OK.TabIndex = 124;
            btn_OK.Text = "OK";
            btn_OK.UseVisualStyleBackColor = false;
            // 
            // btn_Cancel
            // 
            btn_Cancel.BackColor = System.Drawing.Color.White;
            btn_Cancel.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            btn_Cancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            btn_Cancel.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            btn_Cancel.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            btn_Cancel.ForeColor = System.Drawing.Color.FromArgb(70, 70, 70);
            btn_Cancel.Location = new System.Drawing.Point(483, 338);
            btn_Cancel.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            btn_Cancel.Name = "btn_Cancel";
            btn_Cancel.Size = new System.Drawing.Size(62, 26);
            btn_Cancel.TabIndex = 123;
            btn_Cancel.Text = "Cancel";
            btn_Cancel.UseVisualStyleBackColor = false;
            // 
            // ComboBox1
            // 
            ComboBox1.Font = new System.Drawing.Font("Segoe UI", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            ComboBox1.FormattingEnabled = true;
            ComboBox1.Items.AddRange(new object[] { "SOFTEKO", "About Us", "Help", "Feedback" });
            ComboBox1.Location = new System.Drawing.Point(13, 340);
            ComboBox1.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            ComboBox1.Name = "ComboBox1";
            ComboBox1.Size = new System.Drawing.Size(84, 25);
            ComboBox1.TabIndex = 122;
            ComboBox1.Text = "SOFTEKO";
            // 
            // CheckBox1
            // 
            CheckBox1.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            CheckBox1.Location = new System.Drawing.Point(13, 302);
            CheckBox1.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            CheckBox1.Name = "CheckBox1";
            CheckBox1.Size = new System.Drawing.Size(258, 29);
            CheckBox1.TabIndex = 121;
            CheckBox1.Text = "Create a copy of the original worksheet";
            CheckBox1.UseVisualStyleBackColor = true;
            // 
            // AutoSelection
            // 
            AutoSelection.BackColor = System.Drawing.Color.White;
            AutoSelection.Image = (System.Drawing.Image)resources.GetObject("AutoSelection.Image");
            AutoSelection.Location = new System.Drawing.Point(215, 43);
            AutoSelection.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            AutoSelection.Name = "AutoSelection";
            AutoSelection.Size = new System.Drawing.Size(24, 23);
            AutoSelection.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            AutoSelection.TabIndex = 119;
            AutoSelection.TabStop = false;
            // 
            // txtSourceRange
            // 
            txtSourceRange.BackColor = System.Drawing.Color.White;
            txtSourceRange.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            txtSourceRange.Font = new System.Drawing.Font("Segoe UI", 10.0f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            txtSourceRange.ForeColor = System.Drawing.Color.FromArgb(70, 70, 70);
            txtSourceRange.Location = new System.Drawing.Point(15, 42);
            txtSourceRange.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            txtSourceRange.Name = "txtSourceRange";
            txtSourceRange.Size = new System.Drawing.Size(248, 25);
            txtSourceRange.TabIndex = 118;
            // 
            // Label1
            // 
            Label1.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 0);
            Label1.ForeColor = System.Drawing.Color.Black;
            Label1.Location = new System.Drawing.Point(15, 15);
            Label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            Label1.Name = "Label1";
            Label1.Size = new System.Drawing.Size(175, 20);
            Label1.TabIndex = 117;
            Label1.Text = "Source Range:";
            // 
            // GB_sample
            // 
            GB_sample.BackColor = System.Drawing.Color.White;
            GB_sample.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            GB_sample.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 0);
            GB_sample.Location = new System.Drawing.Point(296, 15);
            GB_sample.Name = "GB_sample";
            GB_sample.Size = new System.Drawing.Size(249, 296);
            GB_sample.TabIndex = 399;
            GB_sample.TabStop = false;
            GB_sample.Text = "Sample Image";
            // 
            // CustomGroupBox3
            // 
            CustomGroupBox3.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            CustomGroupBox3.Controls.Add(CustomGroupBox6);
            CustomGroupBox3.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 0);
            CustomGroupBox3.Location = new System.Drawing.Point(15, 81);
            CustomGroupBox3.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            CustomGroupBox3.Name = "CustomGroupBox3";
            CustomGroupBox3.Padding = new System.Windows.Forms.Padding(4, 3, 4, 3);
            CustomGroupBox3.Size = new System.Drawing.Size(249, 88);
            CustomGroupBox3.TabIndex = 132;
            CustomGroupBox3.TabStop = false;
            CustomGroupBox3.Text = "Range Type";
            // 
            // CustomGroupBox6
            // 
            CustomGroupBox6.BackColor = System.Drawing.Color.White;
            CustomGroupBox6.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            CustomGroupBox6.Controls.Add(RB_Single_Range);
            CustomGroupBox6.Controls.Add(RB_Multiple_Range);
            CustomGroupBox6.Location = new System.Drawing.Point(1, 24);
            CustomGroupBox6.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            CustomGroupBox6.Name = "CustomGroupBox6";
            CustomGroupBox6.Padding = new System.Windows.Forms.Padding(4, 3, 4, 3);
            CustomGroupBox6.Size = new System.Drawing.Size(248, 64);
            CustomGroupBox6.TabIndex = 0;
            CustomGroupBox6.TabStop = false;
            // 
            // RB_Single_Range
            // 
            RB_Single_Range.AutoSize = true;
            RB_Single_Range.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            RB_Single_Range.Location = new System.Drawing.Point(8, 9);
            RB_Single_Range.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            RB_Single_Range.Name = "RB_Single_Range";
            RB_Single_Range.Size = new System.Drawing.Size(102, 21);
            RB_Single_Range.TabIndex = 93;
            RB_Single_Range.TabStop = true;
            RB_Single_Range.Text = "Single Range";
            RB_Single_Range.UseVisualStyleBackColor = true;
            // 
            // RB_Multiple_Range
            // 
            RB_Multiple_Range.AutoSize = true;
            RB_Multiple_Range.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            RB_Multiple_Range.Location = new System.Drawing.Point(8, 33);
            RB_Multiple_Range.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            RB_Multiple_Range.Name = "RB_Multiple_Range";
            RB_Multiple_Range.Size = new System.Drawing.Size(197, 21);
            RB_Multiple_Range.TabIndex = 92;
            RB_Multiple_Range.TabStop = true;
            RB_Multiple_Range.Text = "Multiple Non-adjacent Range";
            RB_Multiple_Range.UseVisualStyleBackColor = true;
            // 
            // CustomGroupBox4
            // 
            CustomGroupBox4.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            CustomGroupBox4.Controls.Add(CustomGroupBox5);
            CustomGroupBox4.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 0);
            CustomGroupBox4.Location = new System.Drawing.Point(15, 186);
            CustomGroupBox4.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            CustomGroupBox4.Name = "CustomGroupBox4";
            CustomGroupBox4.Padding = new System.Windows.Forms.Padding(4, 3, 4, 3);
            CustomGroupBox4.Size = new System.Drawing.Size(249, 110);
            CustomGroupBox4.TabIndex = 133;
            CustomGroupBox4.TabStop = false;
            CustomGroupBox4.Text = "Hide Option";
            // 
            // CustomGroupBox5
            // 
            CustomGroupBox5.BackColor = System.Drawing.Color.White;
            CustomGroupBox5.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            CustomGroupBox5.Controls.Add(RB_bidirection);
            CustomGroupBox5.Controls.Add(RB_Row);
            CustomGroupBox5.Controls.Add(RB_Column);
            CustomGroupBox5.Location = new System.Drawing.Point(1, 22);
            CustomGroupBox5.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            CustomGroupBox5.Name = "CustomGroupBox5";
            CustomGroupBox5.Padding = new System.Windows.Forms.Padding(4, 3, 4, 3);
            CustomGroupBox5.Size = new System.Drawing.Size(248, 87);
            CustomGroupBox5.TabIndex = 0;
            CustomGroupBox5.TabStop = false;
            // 
            // RB_bidirection
            // 
            RB_bidirection.AutoSize = true;
            RB_bidirection.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            RB_bidirection.Location = new System.Drawing.Point(8, 58);
            RB_bidirection.Name = "RB_bidirection";
            RB_bidirection.Size = new System.Drawing.Size(117, 21);
            RB_bidirection.TabIndex = 136;
            RB_bidirection.TabStop = true;
            RB_bidirection.Text = "Both directional";
            RB_bidirection.UseVisualStyleBackColor = true;
            // 
            // RB_Row
            // 
            RB_Row.AutoSize = true;
            RB_Row.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            RB_Row.Location = new System.Drawing.Point(8, 8);
            RB_Row.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            RB_Row.Name = "RB_Row";
            RB_Row.Size = new System.Drawing.Size(109, 21);
            RB_Row.TabIndex = 117;
            RB_Row.TabStop = true;
            RB_Row.Text = "Row-wise only";
            RB_Row.UseVisualStyleBackColor = true;
            // 
            // RB_Column
            // 
            RB_Column.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            RB_Column.Location = new System.Drawing.Point(8, 32);
            RB_Column.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            RB_Column.Name = "RB_Column";
            RB_Column.Size = new System.Drawing.Size(151, 24);
            RB_Column.TabIndex = 94;
            RB_Column.TabStop = true;
            RB_Column.Text = "Column-wise only";
            RB_Column.UseVisualStyleBackColor = true;
            // 
            // Form12HideRanges
            // 
            AutoScaleDimensions = new System.Drawing.SizeF(6.0f, 13.0f);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            ClientSize = new System.Drawing.Size(565, 385);
            Controls.Add(GB_sample);
            Controls.Add(CustomGroupBox3);
            Controls.Add(CustomGroupBox4);
            Controls.Add(Selection);
            Controls.Add(btn_OK);
            Controls.Add(btn_Cancel);
            Controls.Add(ComboBox1);
            Controls.Add(CheckBox1);
            Controls.Add(AutoSelection);
            Controls.Add(txtSourceRange);
            Controls.Add(Label1);
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "Form12HideRanges";
            Text = "Hide Only the Selected Range";
            ((System.ComponentModel.ISupportInitialize)Selection).EndInit();
            ((System.ComponentModel.ISupportInitialize)AutoSelection).EndInit();
            CustomGroupBox3.ResumeLayout(false);
            CustomGroupBox6.ResumeLayout(false);
            CustomGroupBox6.PerformLayout();
            CustomGroupBox4.ResumeLayout(false);
            CustomGroupBox5.ResumeLayout(false);
            CustomGroupBox5.PerformLayout();
            KeyDown += new System.Windows.Forms.KeyEventHandler(Form1_KeyDown);
            Load += new EventHandler(Form12HideRanges_Load);
            Activated += new EventHandler(Form1_Activated);
            Closing += new System.ComponentModel.CancelEventHandler(Form12HideRanges_Closing);
            Disposed += new EventHandler(Form12HideRanges_Disposed);
            Shown += new EventHandler(Form12HideRanges_Shown);
            ResumeLayout(false);
            PerformLayout();

        }
        internal System.Windows.Forms.RadioButton RB_Single_Range;
        internal System.Windows.Forms.RadioButton RB_Multiple_Range;
        internal CustomGroupBox CustomGroupBox6;
        internal CustomGroupBox CustomGroupBox3;
        internal CustomGroupBox CustomGroupBox5;
        internal System.Windows.Forms.RadioButton RB_Row;
        internal System.Windows.Forms.RadioButton RB_Column;
        internal CustomGroupBox CustomGroupBox4;
        internal System.Windows.Forms.PictureBox Selection;
        internal System.Windows.Forms.Button btn_OK;
        internal System.Windows.Forms.Button btn_Cancel;
        internal System.Windows.Forms.ComboBox ComboBox1;
        internal System.Windows.Forms.CheckBox CheckBox1;
        internal System.Windows.Forms.PictureBox AutoSelection;
        internal System.Windows.Forms.TextBox txtSourceRange;
        internal System.Windows.Forms.Label Label1;
        internal System.Windows.Forms.RadioButton RB_bidirection;
        internal CustomGroupBox GB_sample;
    }
}