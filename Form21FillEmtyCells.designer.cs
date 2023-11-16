using System;

namespace VSTO_Addins
{
    [Microsoft.VisualBasic.CompilerServices.DesignerGenerated()]
    public partial class Form21FillEmtyCells : System.Windows.Forms.Form
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
            var resources = new System.ComponentModel.ComponentResourceManager(typeof(Form21FillEmtyCells));
            btn_OK = new System.Windows.Forms.Button();
            btn_OK.Click += new EventHandler(btn_OK_Click);
            btn_Cancel = new System.Windows.Forms.Button();
            btn_Cancel.Click += new EventHandler(btn_Cancel_Click);
            ComboBox1 = new System.Windows.Forms.ComboBox();
            CB_Backup_Sheet = new System.Windows.Forms.CheckBox();
            txtSourceRange = new System.Windows.Forms.TextBox();
            txtSourceRange.TextChanged += new EventHandler(Textbox1_TextChanged);
            txtSourceRange.GotFocus += new EventHandler(txtSourceRange_GotFocus);
            Label1 = new System.Windows.Forms.Label();
            Selection = new System.Windows.Forms.PictureBox();
            Selection.Click += new EventHandler(Selection_Click);
            CB_Keepformatting = new System.Windows.Forms.CheckBox();
            L_Fill_Options = new System.Windows.Forms.Label();
            ComboBox_Options = new System.Windows.Forms.ComboBox();
            L_Fill_Value = new System.Windows.Forms.Label();
            txtFillValue = new System.Windows.Forms.TextBox();
            CustomGroupBox3 = new CustomGroupBox();
            CustomGroupBox6 = new CustomGroupBox();
            PictureBox2 = new System.Windows.Forms.PictureBox();
            PictureBox1 = new System.Windows.Forms.PictureBox();
            PictureBox11 = new System.Windows.Forms.PictureBox();
            RB_Certain_value = new System.Windows.Forms.RadioButton();
            RB_Certain_value.CheckedChanged += new EventHandler(RB_Certain_value_CheckedChanged);
            RB_Values_fromselected_range = new System.Windows.Forms.RadioButton();
            RB_Values_fromselected_range.CheckedChanged += new EventHandler(RB_Values_fromselected_range_CheckedChanged);
            RB_Linear_values = new System.Windows.Forms.RadioButton();
            RB_Linear_values.CheckedChanged += new EventHandler(RB_Linear_values_CheckedChanged);
            GB_sample = new CustomGroupBox();
            ((System.ComponentModel.ISupportInitialize)Selection).BeginInit();
            CustomGroupBox3.SuspendLayout();
            CustomGroupBox6.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)PictureBox2).BeginInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox1).BeginInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox11).BeginInit();
            SuspendLayout();
            // 
            // btn_OK
            // 
            btn_OK.BackColor = System.Drawing.Color.White;
            btn_OK.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            btn_OK.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            btn_OK.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            btn_OK.ForeColor = System.Drawing.Color.FromArgb(70, 70, 70);
            btn_OK.Location = new System.Drawing.Point(397, 351);
            btn_OK.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            btn_OK.Name = "btn_OK";
            btn_OK.Size = new System.Drawing.Size(62, 26);
            btn_OK.TabIndex = 167;
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
            btn_Cancel.Location = new System.Drawing.Point(475, 351);
            btn_Cancel.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            btn_Cancel.Name = "btn_Cancel";
            btn_Cancel.Size = new System.Drawing.Size(62, 26);
            btn_Cancel.TabIndex = 166;
            btn_Cancel.Text = "Cancel";
            btn_Cancel.UseVisualStyleBackColor = false;
            // 
            // ComboBox1
            // 
            ComboBox1.Font = new System.Drawing.Font("Segoe UI", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            ComboBox1.FormattingEnabled = true;
            ComboBox1.Items.AddRange(new object[] { "SOFTEKO", "About Us", "Help", "Feedback" });
            ComboBox1.Location = new System.Drawing.Point(15, 353);
            ComboBox1.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            ComboBox1.Name = "ComboBox1";
            ComboBox1.Size = new System.Drawing.Size(100, 25);
            ComboBox1.TabIndex = 165;
            ComboBox1.Text = "SOFTEKO";
            // 
            // CB_Backup_Sheet
            // 
            CB_Backup_Sheet.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            CB_Backup_Sheet.Location = new System.Drawing.Point(15, 312);
            CB_Backup_Sheet.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            CB_Backup_Sheet.Name = "CB_Backup_Sheet";
            CB_Backup_Sheet.Size = new System.Drawing.Size(258, 29);
            CB_Backup_Sheet.TabIndex = 164;
            CB_Backup_Sheet.Text = "Create a copy of the original worksheet";
            CB_Backup_Sheet.UseVisualStyleBackColor = true;
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
            txtSourceRange.TabIndex = 162;
            // 
            // Label1
            // 
            Label1.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            Label1.ForeColor = System.Drawing.Color.Black;
            Label1.Location = new System.Drawing.Point(15, 15);
            Label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            Label1.Name = "Label1";
            Label1.Size = new System.Drawing.Size(175, 20);
            Label1.TabIndex = 161;
            Label1.Text = "Source Range:";
            // 
            // Selection
            // 
            Selection.BackColor = System.Drawing.Color.White;
            Selection.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            Selection.Image = (System.Drawing.Image)resources.GetObject("Selection.Image");
            Selection.Location = new System.Drawing.Point(239, 42);
            Selection.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            Selection.Name = "Selection";
            Selection.Size = new System.Drawing.Size(24, 25);
            Selection.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            Selection.TabIndex = 168;
            Selection.TabStop = false;
            // 
            // CB_Keepformatting
            // 
            CB_Keepformatting.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            CB_Keepformatting.Location = new System.Drawing.Point(15, 193);
            CB_Keepformatting.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            CB_Keepformatting.Name = "CB_Keepformatting";
            CB_Keepformatting.Size = new System.Drawing.Size(136, 29);
            CB_Keepformatting.TabIndex = 165;
            CB_Keepformatting.Text = "Keep formatting";
            CB_Keepformatting.UseVisualStyleBackColor = true;
            // 
            // L_Fill_Options
            // 
            L_Fill_Options.AutoSize = true;
            L_Fill_Options.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            L_Fill_Options.Location = new System.Drawing.Point(15, 234);
            L_Fill_Options.Name = "L_Fill_Options";
            L_Fill_Options.Size = new System.Drawing.Size(76, 17);
            L_Fill_Options.TabIndex = 174;
            L_Fill_Options.Text = "Fill Options";
            // 
            // ComboBox_Options
            // 
            ComboBox_Options.BackColor = System.Drawing.Color.White;
            ComboBox_Options.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            ComboBox_Options.Font = new System.Drawing.Font("Segoe UI", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            ComboBox_Options.ForeColor = System.Drawing.Color.Black;
            ComboBox_Options.FormattingEnabled = true;
            ComboBox_Options.Location = new System.Drawing.Point(101, 230);
            ComboBox_Options.Name = "ComboBox_Options";
            ComboBox_Options.Size = new System.Drawing.Size(163, 25);
            ComboBox_Options.TabIndex = 175;
            // 
            // L_Fill_Value
            // 
            L_Fill_Value.AutoSize = true;
            L_Fill_Value.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            L_Fill_Value.Location = new System.Drawing.Point(15, 276);
            L_Fill_Value.Name = "L_Fill_Value";
            L_Fill_Value.Size = new System.Drawing.Size(60, 17);
            L_Fill_Value.TabIndex = 176;
            L_Fill_Value.Text = "Fill Value";
            // 
            // txtFillValue
            // 
            txtFillValue.Font = new System.Drawing.Font("Segoe UI", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            txtFillValue.Location = new System.Drawing.Point(101, 273);
            txtFillValue.Name = "txtFillValue";
            txtFillValue.Size = new System.Drawing.Size(163, 25);
            txtFillValue.TabIndex = 178;
            // 
            // CustomGroupBox3
            // 
            CustomGroupBox3.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            CustomGroupBox3.Controls.Add(CustomGroupBox6);
            CustomGroupBox3.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 0);
            CustomGroupBox3.Location = new System.Drawing.Point(15, 80);
            CustomGroupBox3.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            CustomGroupBox3.Name = "CustomGroupBox3";
            CustomGroupBox3.Padding = new System.Windows.Forms.Padding(4, 3, 4, 3);
            CustomGroupBox3.Size = new System.Drawing.Size(249, 107);
            CustomGroupBox3.TabIndex = 169;
            CustomGroupBox3.TabStop = false;
            CustomGroupBox3.Text = "Fill Cells";
            // 
            // CustomGroupBox6
            // 
            CustomGroupBox6.BackColor = System.Drawing.Color.White;
            CustomGroupBox6.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            CustomGroupBox6.Controls.Add(PictureBox2);
            CustomGroupBox6.Controls.Add(PictureBox1);
            CustomGroupBox6.Controls.Add(PictureBox11);
            CustomGroupBox6.Controls.Add(RB_Certain_value);
            CustomGroupBox6.Controls.Add(RB_Values_fromselected_range);
            CustomGroupBox6.Controls.Add(RB_Linear_values);
            CustomGroupBox6.Location = new System.Drawing.Point(1, 24);
            CustomGroupBox6.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            CustomGroupBox6.Name = "CustomGroupBox6";
            CustomGroupBox6.Padding = new System.Windows.Forms.Padding(4, 3, 4, 3);
            CustomGroupBox6.Size = new System.Drawing.Size(248, 82);
            CustomGroupBox6.TabIndex = 0;
            CustomGroupBox6.TabStop = false;
            // 
            // PictureBox2
            // 
            PictureBox2.Image = (System.Drawing.Image)resources.GetObject("PictureBox2.Image");
            PictureBox2.Location = new System.Drawing.Point(220, 57);
            PictureBox2.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            PictureBox2.Name = "PictureBox2";
            PictureBox2.Size = new System.Drawing.Size(20, 20);
            PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            PictureBox2.TabIndex = 235;
            PictureBox2.TabStop = false;
            // 
            // PictureBox1
            // 
            PictureBox1.Image = (System.Drawing.Image)resources.GetObject("PictureBox1.Image");
            PictureBox1.Location = new System.Drawing.Point(220, 33);
            PictureBox1.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            PictureBox1.Name = "PictureBox1";
            PictureBox1.Size = new System.Drawing.Size(20, 20);
            PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            PictureBox1.TabIndex = 234;
            PictureBox1.TabStop = false;
            // 
            // PictureBox11
            // 
            PictureBox11.Image = (System.Drawing.Image)resources.GetObject("PictureBox11.Image");
            PictureBox11.Location = new System.Drawing.Point(220, 9);
            PictureBox11.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            PictureBox11.Name = "PictureBox11";
            PictureBox11.Size = new System.Drawing.Size(20, 20);
            PictureBox11.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            PictureBox11.TabIndex = 233;
            PictureBox11.TabStop = false;
            // 
            // RB_Certain_value
            // 
            RB_Certain_value.AutoSize = true;
            RB_Certain_value.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            RB_Certain_value.Location = new System.Drawing.Point(8, 55);
            RB_Certain_value.Name = "RB_Certain_value";
            RB_Certain_value.Size = new System.Drawing.Size(129, 21);
            RB_Certain_value.TabIndex = 94;
            RB_Certain_value.Text = "With certain value";
            RB_Certain_value.UseVisualStyleBackColor = true;
            // 
            // RB_Values_fromselected_range
            // 
            RB_Values_fromselected_range.AutoSize = true;
            RB_Values_fromselected_range.Checked = true;
            RB_Values_fromselected_range.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            RB_Values_fromselected_range.Location = new System.Drawing.Point(8, 8);
            RB_Values_fromselected_range.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            RB_Values_fromselected_range.Name = "RB_Values_fromselected_range";
            RB_Values_fromselected_range.Size = new System.Drawing.Size(214, 21);
            RB_Values_fromselected_range.TabIndex = 93;
            RB_Values_fromselected_range.TabStop = true;
            RB_Values_fromselected_range.Text = "With values from selected range";
            RB_Values_fromselected_range.UseVisualStyleBackColor = true;
            // 
            // RB_Linear_values
            // 
            RB_Linear_values.AutoSize = true;
            RB_Linear_values.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            RB_Linear_values.Location = new System.Drawing.Point(8, 31);
            RB_Linear_values.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            RB_Linear_values.Name = "RB_Linear_values";
            RB_Linear_values.Size = new System.Drawing.Size(128, 21);
            RB_Linear_values.TabIndex = 92;
            RB_Linear_values.Text = "With linear values";
            RB_Linear_values.UseVisualStyleBackColor = true;
            // 
            // GB_sample
            // 
            GB_sample.BackColor = System.Drawing.Color.White;
            GB_sample.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            GB_sample.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 0);
            GB_sample.Location = new System.Drawing.Point(286, 15);
            GB_sample.Name = "GB_sample";
            GB_sample.Size = new System.Drawing.Size(251, 315);
            GB_sample.TabIndex = 401;
            GB_sample.TabStop = false;
            GB_sample.Text = "Sample Image";
            // 
            // Form21FillEmtyCells
            // 
            AutoScaleDimensions = new System.Drawing.SizeF(6.0f, 13.0f);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            ClientSize = new System.Drawing.Size(560, 399);
            Controls.Add(GB_sample);
            Controls.Add(txtFillValue);
            Controls.Add(L_Fill_Value);
            Controls.Add(ComboBox_Options);
            Controls.Add(L_Fill_Options);
            Controls.Add(CB_Keepformatting);
            Controls.Add(CustomGroupBox3);
            Controls.Add(btn_OK);
            Controls.Add(btn_Cancel);
            Controls.Add(ComboBox1);
            Controls.Add(CB_Backup_Sheet);
            Controls.Add(Label1);
            Controls.Add(Selection);
            Controls.Add(txtSourceRange);
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "Form21FillEmtyCells";
            Text = "Fill Emty Cells";
            ((System.ComponentModel.ISupportInitialize)Selection).EndInit();
            CustomGroupBox3.ResumeLayout(false);
            CustomGroupBox6.ResumeLayout(false);
            CustomGroupBox6.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)PictureBox2).EndInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox1).EndInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox11).EndInit();
            KeyDown += new System.Windows.Forms.KeyEventHandler(Form1_KeyDown);
            Load += new EventHandler(Form21FillEmtyCells_Load);
            Activated += new EventHandler(Form1_Activated);
            Closing += new System.ComponentModel.CancelEventHandler(Form21FillEmtyCells_Closing);
            Disposed += new EventHandler(Form21FillEmtyCells_Disposed);
            Shown += new EventHandler(Form21FillEmtyCells_Shown);
            ResumeLayout(false);
            PerformLayout();

        }
        internal System.Windows.Forms.RadioButton RB_Values_fromselected_range;
        internal System.Windows.Forms.RadioButton RB_Linear_values;
        internal CustomGroupBox CustomGroupBox6;
        internal System.Windows.Forms.RadioButton RB_Certain_value;
        internal CustomGroupBox CustomGroupBox3;
        internal System.Windows.Forms.Button btn_OK;
        internal System.Windows.Forms.Button btn_Cancel;
        internal System.Windows.Forms.ComboBox ComboBox1;
        internal System.Windows.Forms.CheckBox CB_Backup_Sheet;
        internal System.Windows.Forms.TextBox txtSourceRange;
        internal System.Windows.Forms.Label Label1;
        internal System.Windows.Forms.PictureBox Selection;
        internal System.Windows.Forms.CheckBox CB_Keepformatting;
        internal System.Windows.Forms.Label L_Fill_Options;
        internal System.Windows.Forms.ComboBox ComboBox_Options;
        internal System.Windows.Forms.Label L_Fill_Value;
        internal System.Windows.Forms.TextBox txtFillValue;
        internal System.Windows.Forms.PictureBox PictureBox2;
        internal System.Windows.Forms.PictureBox PictureBox1;
        internal System.Windows.Forms.PictureBox PictureBox11;
        internal CustomGroupBox GB_sample;
    }
}