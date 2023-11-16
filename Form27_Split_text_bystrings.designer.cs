using System;

namespace VSTO_Addins
{
    [Microsoft.VisualBasic.CompilerServices.DesignerGenerated()]
    public partial class Form27_Split_text_bystrings : System.Windows.Forms.Form
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
            components = new System.ComponentModel.Container();
            var resources = new System.ComponentModel.ComponentResourceManager(typeof(Form27_Split_text_bystrings));
            CB_consecutive_separators = new System.Windows.Forms.CheckBox();
            CB_consecutive_separators.CheckedChanged += new EventHandler(CB_consecutive_separators_CheckedChanged);
            ToolTip1 = new System.Windows.Forms.ToolTip(components);
            Info = new System.Windows.Forms.PictureBox();
            PictureBox2 = new System.Windows.Forms.PictureBox();
            Label1 = new System.Windows.Forms.Label();
            AutoSelection = new System.Windows.Forms.PictureBox();
            AutoSelection.Click += new EventHandler(AutoSelection_Click);
            AutoSelection.GotFocus += new EventHandler(AutoSelection_GotFocus);
            Selection = new System.Windows.Forms.PictureBox();
            Selection.Click += new EventHandler(Selection_Click);
            Selection.GotFocus += new EventHandler(Selection_GotFocus);
            PictureBox7 = new System.Windows.Forms.PictureBox();
            ComboBox1 = new System.Windows.Forms.ComboBox();
            ComboBox1.SelectedIndexChanged += new EventHandler(ComboBox1_SelectedIndexChanged);
            Btn_OK = new System.Windows.Forms.Button();
            Btn_OK.Click += new EventHandler(Btn_OK_Click);
            Btn_OK.MouseEnter += new EventHandler(Btn_OK_MouseEnter);
            Btn_OK.MouseLeave += new EventHandler(Btn_OK_MouseLeave);
            Btn_Cancel = new System.Windows.Forms.Button();
            Btn_Cancel.MouseEnter += new EventHandler(Btn_Cancel_MouseEnter);
            Btn_Cancel.MouseLeave += new EventHandler(Btn_Cancel_MouseLeave);
            Btn_Cancel.Click += new EventHandler(Btn_Cancel_Click);
            CB_formatting = new System.Windows.Forms.CheckBox();
            CB_formatting.CheckedChanged += new EventHandler(CB_formatting_CheckedChanged);
            CB_backup = new System.Windows.Forms.CheckBox();
            TB_source_range = new System.Windows.Forms.TextBox();
            TB_source_range.TextChanged += new EventHandler(TB_source_range_TextChanged);
            TB_source_range.GotFocus += new EventHandler(TB_source_range_GotFocus);
            CustomGroupBox2 = new CustomGroupBox();
            PictureBox3 = new System.Windows.Forms.PictureBox();
            PictureBox4 = new System.Windows.Forms.PictureBox();
            RB_ending_point = new System.Windows.Forms.RadioButton();
            RB_ending_point.CheckedChanged += new EventHandler(RB_ending_point_CheckedChanged);
            RB_starting_point = new System.Windows.Forms.RadioButton();
            RB_starting_point.CheckedChanged += new EventHandler(RB_starting_point_CheckedChanged);
            CB_separators_finaloutput = new System.Windows.Forms.CheckBox();
            CB_separators_finaloutput.CheckedChanged += new EventHandler(CB_separators_finaloutput_CheckedChanged);
            CustomGroupBox4 = new CustomGroupBox();
            TB_TypeString = new System.Windows.Forms.TextBox();
            TB_TypeString.TextChanged += new EventHandler(TB_TypeString_TextChanged);
            Label2 = new System.Windows.Forms.Label();
            CustomGroupBox6 = new CustomGroupBox();
            Panel_ExpectedOutput = new CustomPanel();
            CustomGroupBox1 = new CustomGroupBox();
            CustomGroupBox7 = new CustomGroupBox();
            PictureBox8 = new System.Windows.Forms.PictureBox();
            PictureBox1 = new System.Windows.Forms.PictureBox();
            RB_columns = new System.Windows.Forms.RadioButton();
            RB_columns.CheckedChanged += new EventHandler(RB_columns_CheckedChanged);
            RB_rows = new System.Windows.Forms.RadioButton();
            RB_rows.CheckedChanged += new EventHandler(RB_rows_CheckedChanged);
            CustomGroupBox5 = new CustomGroupBox();
            Panel_InputRange = new CustomPanel();
            ((System.ComponentModel.ISupportInitialize)Info).BeginInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox2).BeginInit();
            ((System.ComponentModel.ISupportInitialize)AutoSelection).BeginInit();
            ((System.ComponentModel.ISupportInitialize)Selection).BeginInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox7).BeginInit();
            CustomGroupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)PictureBox3).BeginInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox4).BeginInit();
            CustomGroupBox4.SuspendLayout();
            CustomGroupBox6.SuspendLayout();
            CustomGroupBox1.SuspendLayout();
            CustomGroupBox7.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)PictureBox8).BeginInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox1).BeginInit();
            CustomGroupBox5.SuspendLayout();
            SuspendLayout();
            // 
            // CB_consecutive_separators
            // 
            CB_consecutive_separators.AutoSize = true;
            CB_consecutive_separators.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            CB_consecutive_separators.Location = new System.Drawing.Point(15, 330);
            CB_consecutive_separators.Name = "CB_consecutive_separators";
            CB_consecutive_separators.Size = new System.Drawing.Size(237, 21);
            CB_consecutive_separators.TabIndex = 323;
            CB_consecutive_separators.Text = "Treat consecutive separators as one";
            CB_consecutive_separators.UseVisualStyleBackColor = true;
            // 
            // Info
            // 
            Info.Image = (System.Drawing.Image)resources.GetObject("Info.Image");
            Info.Location = new System.Drawing.Point(118, 15);
            Info.Name = "Info";
            Info.Size = new System.Drawing.Size(20, 20);
            Info.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            Info.TabIndex = 322;
            Info.TabStop = false;
            ToolTip1.SetToolTip(Info, "Please, select single column");
            // 
            // PictureBox2
            // 
            PictureBox2.Image = (System.Drawing.Image)resources.GetObject("PictureBox2.Image");
            PictureBox2.Location = new System.Drawing.Point(144, 27);
            PictureBox2.Name = "PictureBox2";
            PictureBox2.Size = new System.Drawing.Size(20, 20);
            PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            PictureBox2.TabIndex = 323;
            PictureBox2.TabStop = false;
            ToolTip1.SetToolTip(PictureBox2, "Please, select single column");
            // 
            // Label1
            // 
            Label1.AutoSize = true;
            Label1.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            Label1.Location = new System.Drawing.Point(15, 15);
            Label1.Name = "Label1";
            Label1.Size = new System.Drawing.Size(98, 17);
            Label1.TabIndex = 308;
            Label1.Text = "Source Range :";
            // 
            // AutoSelection
            // 
            AutoSelection.BackColor = System.Drawing.Color.White;
            AutoSelection.Image = (System.Drawing.Image)resources.GetObject("AutoSelection.Image");
            AutoSelection.Location = new System.Drawing.Point(228, 43);
            AutoSelection.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            AutoSelection.Name = "AutoSelection";
            AutoSelection.Size = new System.Drawing.Size(24, 23);
            AutoSelection.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            AutoSelection.TabIndex = 319;
            AutoSelection.TabStop = false;
            // 
            // Selection
            // 
            Selection.BackColor = System.Drawing.Color.White;
            Selection.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            Selection.Image = (System.Drawing.Image)resources.GetObject("Selection.Image");
            Selection.Location = new System.Drawing.Point(253, 42);
            Selection.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            Selection.Name = "Selection";
            Selection.Size = new System.Drawing.Size(24, 25);
            Selection.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            Selection.TabIndex = 320;
            Selection.TabStop = false;
            // 
            // PictureBox7
            // 
            PictureBox7.Image = (System.Drawing.Image)resources.GetObject("PictureBox7.Image");
            PictureBox7.Location = new System.Drawing.Point(432, 221);
            PictureBox7.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            PictureBox7.Name = "PictureBox7";
            PictureBox7.Size = new System.Drawing.Size(60, 60);
            PictureBox7.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            PictureBox7.TabIndex = 317;
            PictureBox7.TabStop = false;
            // 
            // ComboBox1
            // 
            ComboBox1.Font = new System.Drawing.Font("Segoe UI", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            ComboBox1.FormattingEnabled = true;
            ComboBox1.Items.AddRange(new object[] { "SOFTEKO", "About Us", "Help", "Feedback" });
            ComboBox1.Location = new System.Drawing.Point(15, 500);
            ComboBox1.Name = "ComboBox1";
            ComboBox1.Size = new System.Drawing.Size(154, 25);
            ComboBox1.TabIndex = 312;
            ComboBox1.Text = "SOFTEKO";
            // 
            // Btn_OK
            // 
            Btn_OK.BackColor = System.Drawing.Color.White;
            Btn_OK.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            Btn_OK.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            Btn_OK.Location = new System.Drawing.Point(466, 498);
            Btn_OK.Name = "Btn_OK";
            Btn_OK.Size = new System.Drawing.Size(62, 26);
            Btn_OK.TabIndex = 316;
            Btn_OK.Text = "OK";
            Btn_OK.UseVisualStyleBackColor = false;
            // 
            // Btn_Cancel
            // 
            Btn_Cancel.BackColor = System.Drawing.Color.White;
            Btn_Cancel.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            Btn_Cancel.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            Btn_Cancel.Location = new System.Drawing.Point(545, 498);
            Btn_Cancel.Name = "Btn_Cancel";
            Btn_Cancel.Size = new System.Drawing.Size(62, 26);
            Btn_Cancel.TabIndex = 315;
            Btn_Cancel.Text = "Cancel";
            Btn_Cancel.UseVisualStyleBackColor = false;
            // 
            // CB_formatting
            // 
            CB_formatting.AutoSize = true;
            CB_formatting.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            CB_formatting.Location = new System.Drawing.Point(15, 171);
            CB_formatting.Name = "CB_formatting";
            CB_formatting.Size = new System.Drawing.Size(122, 21);
            CB_formatting.TabIndex = 310;
            CB_formatting.Text = "Keep formatting";
            CB_formatting.UseVisualStyleBackColor = true;
            // 
            // CB_backup
            // 
            CB_backup.AutoSize = true;
            CB_backup.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            CB_backup.Location = new System.Drawing.Point(15, 465);
            CB_backup.Name = "CB_backup";
            CB_backup.Size = new System.Drawing.Size(257, 21);
            CB_backup.TabIndex = 311;
            CB_backup.Text = "Create a copy of the original worksheet";
            CB_backup.UseVisualStyleBackColor = true;
            // 
            // TB_source_range
            // 
            TB_source_range.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            TB_source_range.Cursor = System.Windows.Forms.Cursors.IBeam;
            TB_source_range.Font = new System.Drawing.Font("Segoe UI", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            TB_source_range.Location = new System.Drawing.Point(15, 42);
            TB_source_range.Name = "TB_source_range";
            TB_source_range.Size = new System.Drawing.Size(262, 25);
            TB_source_range.TabIndex = 318;
            // 
            // CustomGroupBox2
            // 
            CustomGroupBox2.BackColor = System.Drawing.Color.White;
            CustomGroupBox2.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            CustomGroupBox2.Controls.Add(PictureBox3);
            CustomGroupBox2.Controls.Add(PictureBox4);
            CustomGroupBox2.Controls.Add(RB_ending_point);
            CustomGroupBox2.Controls.Add(RB_starting_point);
            CustomGroupBox2.Controls.Add(CB_separators_finaloutput);
            CustomGroupBox2.Location = new System.Drawing.Point(15, 360);
            CustomGroupBox2.Name = "CustomGroupBox2";
            CustomGroupBox2.Size = new System.Drawing.Size(257, 92);
            CustomGroupBox2.TabIndex = 324;
            CustomGroupBox2.TabStop = false;
            // 
            // PictureBox3
            // 
            PictureBox3.Image = (System.Drawing.Image)resources.GetObject("PictureBox3.Image");
            PictureBox3.Location = new System.Drawing.Point(226, 36);
            PictureBox3.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            PictureBox3.Name = "PictureBox3";
            PictureBox3.Size = new System.Drawing.Size(20, 20);
            PictureBox3.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            PictureBox3.TabIndex = 311;
            PictureBox3.TabStop = false;
            // 
            // PictureBox4
            // 
            PictureBox4.Image = (System.Drawing.Image)resources.GetObject("PictureBox4.Image");
            PictureBox4.Location = new System.Drawing.Point(225, 61);
            PictureBox4.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            PictureBox4.Name = "PictureBox4";
            PictureBox4.Size = new System.Drawing.Size(20, 20);
            PictureBox4.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            PictureBox4.TabIndex = 312;
            PictureBox4.TabStop = false;
            // 
            // RB_ending_point
            // 
            RB_ending_point.AutoSize = true;
            RB_ending_point.Enabled = false;
            RB_ending_point.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            RB_ending_point.Location = new System.Drawing.Point(25, 60);
            RB_ending_point.Name = "RB_ending_point";
            RB_ending_point.Size = new System.Drawing.Size(138, 21);
            RB_ending_point.TabIndex = 310;
            RB_ending_point.TabStop = true;
            RB_ending_point.Text = "At the ending point";
            RB_ending_point.UseVisualStyleBackColor = true;
            // 
            // RB_starting_point
            // 
            RB_starting_point.AutoSize = true;
            RB_starting_point.Enabled = false;
            RB_starting_point.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            RB_starting_point.Location = new System.Drawing.Point(25, 34);
            RB_starting_point.Name = "RB_starting_point";
            RB_starting_point.Size = new System.Drawing.Size(142, 21);
            RB_starting_point.TabIndex = 309;
            RB_starting_point.TabStop = true;
            RB_starting_point.Text = "At the starting point";
            RB_starting_point.UseVisualStyleBackColor = true;
            // 
            // CB_separators_finaloutput
            // 
            CB_separators_finaloutput.AutoSize = true;
            CB_separators_finaloutput.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            CB_separators_finaloutput.Location = new System.Drawing.Point(8, 9);
            CB_separators_finaloutput.Name = "CB_separators_finaloutput";
            CB_separators_finaloutput.Size = new System.Drawing.Size(230, 21);
            CB_separators_finaloutput.TabIndex = 308;
            CB_separators_finaloutput.Text = "Keep separators in the final output";
            CB_separators_finaloutput.UseVisualStyleBackColor = true;
            // 
            // CustomGroupBox4
            // 
            CustomGroupBox4.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            CustomGroupBox4.Controls.Add(TB_TypeString);
            CustomGroupBox4.Controls.Add(PictureBox2);
            CustomGroupBox4.Controls.Add(Label2);
            CustomGroupBox4.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 0);
            CustomGroupBox4.Location = new System.Drawing.Point(15, 198);
            CustomGroupBox4.Name = "CustomGroupBox4";
            CustomGroupBox4.Size = new System.Drawing.Size(260, 120);
            CustomGroupBox4.TabIndex = 321;
            CustomGroupBox4.TabStop = false;
            CustomGroupBox4.Text = "Select Separator";
            // 
            // TB_TypeString
            // 
            TB_TypeString.Location = new System.Drawing.Point(15, 55);
            TB_TypeString.Multiline = true;
            TB_TypeString.Name = "TB_TypeString";
            TB_TypeString.Size = new System.Drawing.Size(230, 49);
            TB_TypeString.TabIndex = 324;
            // 
            // Label2
            // 
            Label2.AutoSize = true;
            Label2.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            Label2.Location = new System.Drawing.Point(12, 27);
            Label2.Name = "Label2";
            Label2.Size = new System.Drawing.Size(127, 17);
            Label2.TabIndex = 0;
            Label2.Text = "Type in the strings :";
            // 
            // CustomGroupBox6
            // 
            CustomGroupBox6.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            CustomGroupBox6.Controls.Add(Panel_ExpectedOutput);
            CustomGroupBox6.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 0);
            CustomGroupBox6.Location = new System.Drawing.Point(315, 287);
            CustomGroupBox6.Name = "CustomGroupBox6";
            CustomGroupBox6.Size = new System.Drawing.Size(292, 190);
            CustomGroupBox6.TabIndex = 314;
            CustomGroupBox6.TabStop = false;
            CustomGroupBox6.Text = "Expected Output";
            // 
            // Panel_ExpectedOutput
            // 
            Panel_ExpectedOutput.BackColor = System.Drawing.Color.White;
            Panel_ExpectedOutput.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            Panel_ExpectedOutput.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            Panel_ExpectedOutput.BorderWidth = 1;
            Panel_ExpectedOutput.Location = new System.Drawing.Point(1, 30);
            Panel_ExpectedOutput.Name = "Panel_ExpectedOutput";
            Panel_ExpectedOutput.Size = new System.Drawing.Size(290, 160);
            Panel_ExpectedOutput.TabIndex = 11;
            // 
            // CustomGroupBox1
            // 
            CustomGroupBox1.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            CustomGroupBox1.Controls.Add(CustomGroupBox7);
            CustomGroupBox1.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 0);
            CustomGroupBox1.Location = new System.Drawing.Point(15, 79);
            CustomGroupBox1.Name = "CustomGroupBox1";
            CustomGroupBox1.Size = new System.Drawing.Size(260, 84);
            CustomGroupBox1.TabIndex = 309;
            CustomGroupBox1.TabStop = false;
            CustomGroupBox1.Text = "Split Option";
            // 
            // CustomGroupBox7
            // 
            CustomGroupBox7.BackColor = System.Drawing.Color.White;
            CustomGroupBox7.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            CustomGroupBox7.Controls.Add(PictureBox8);
            CustomGroupBox7.Controls.Add(PictureBox1);
            CustomGroupBox7.Controls.Add(RB_columns);
            CustomGroupBox7.Controls.Add(RB_rows);
            CustomGroupBox7.Location = new System.Drawing.Point(1, 22);
            CustomGroupBox7.Name = "CustomGroupBox7";
            CustomGroupBox7.Size = new System.Drawing.Size(259, 62);
            CustomGroupBox7.TabIndex = 0;
            CustomGroupBox7.TabStop = false;
            // 
            // PictureBox8
            // 
            PictureBox8.Image = (System.Drawing.Image)resources.GetObject("PictureBox8.Image");
            PictureBox8.Location = new System.Drawing.Point(225, 33);
            PictureBox8.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            PictureBox8.Name = "PictureBox8";
            PictureBox8.Size = new System.Drawing.Size(20, 20);
            PictureBox8.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            PictureBox8.TabIndex = 275;
            PictureBox8.TabStop = false;
            // 
            // PictureBox1
            // 
            PictureBox1.Image = (System.Drawing.Image)resources.GetObject("PictureBox1.Image");
            PictureBox1.Location = new System.Drawing.Point(225, 7);
            PictureBox1.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            PictureBox1.Name = "PictureBox1";
            PictureBox1.Size = new System.Drawing.Size(20, 20);
            PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            PictureBox1.TabIndex = 274;
            PictureBox1.TabStop = false;
            // 
            // RB_columns
            // 
            RB_columns.AutoSize = true;
            RB_columns.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            RB_columns.Location = new System.Drawing.Point(8, 32);
            RB_columns.Name = "RB_columns";
            RB_columns.Size = new System.Drawing.Size(154, 21);
            RB_columns.TabIndex = 1;
            RB_columns.Text = "Split text into columns";
            RB_columns.UseVisualStyleBackColor = true;
            // 
            // RB_rows
            // 
            RB_rows.AutoSize = true;
            RB_rows.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            RB_rows.Location = new System.Drawing.Point(8, 6);
            RB_rows.Name = "RB_rows";
            RB_rows.Size = new System.Drawing.Size(134, 21);
            RB_rows.TabIndex = 0;
            RB_rows.Text = "Split text into rows";
            RB_rows.UseVisualStyleBackColor = true;
            // 
            // CustomGroupBox5
            // 
            CustomGroupBox5.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            CustomGroupBox5.Controls.Add(Panel_InputRange);
            CustomGroupBox5.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 0);
            CustomGroupBox5.Location = new System.Drawing.Point(315, 20);
            CustomGroupBox5.Name = "CustomGroupBox5";
            CustomGroupBox5.Size = new System.Drawing.Size(292, 190);
            CustomGroupBox5.TabIndex = 313;
            CustomGroupBox5.TabStop = false;
            CustomGroupBox5.Text = "Input Range";
            // 
            // Panel_InputRange
            // 
            Panel_InputRange.BackColor = System.Drawing.Color.White;
            Panel_InputRange.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            Panel_InputRange.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            Panel_InputRange.BorderWidth = 1;
            Panel_InputRange.Location = new System.Drawing.Point(1, 30);
            Panel_InputRange.Name = "Panel_InputRange";
            Panel_InputRange.Size = new System.Drawing.Size(290, 160);
            Panel_InputRange.TabIndex = 0;
            // 
            // Form27_Split_text_bystrings
            // 
            AutoScaleDimensions = new System.Drawing.SizeF(6.0f, 13.0f);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            ClientSize = new System.Drawing.Size(635, 546);
            Controls.Add(CustomGroupBox2);
            Controls.Add(CB_consecutive_separators);
            Controls.Add(Label1);
            Controls.Add(AutoSelection);
            Controls.Add(CustomGroupBox4);
            Controls.Add(Selection);
            Controls.Add(PictureBox7);
            Controls.Add(ComboBox1);
            Controls.Add(Btn_OK);
            Controls.Add(Btn_Cancel);
            Controls.Add(CB_formatting);
            Controls.Add(CustomGroupBox6);
            Controls.Add(CB_backup);
            Controls.Add(CustomGroupBox1);
            Controls.Add(CustomGroupBox5);
            Controls.Add(TB_source_range);
            Controls.Add(Info);
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "Form27_Split_text_bystrings";
            Text = "Split Text by Strings";
            ((System.ComponentModel.ISupportInitialize)Info).EndInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox2).EndInit();
            ((System.ComponentModel.ISupportInitialize)AutoSelection).EndInit();
            ((System.ComponentModel.ISupportInitialize)Selection).EndInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox7).EndInit();
            CustomGroupBox2.ResumeLayout(false);
            CustomGroupBox2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)PictureBox3).EndInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox4).EndInit();
            CustomGroupBox4.ResumeLayout(false);
            CustomGroupBox4.PerformLayout();
            CustomGroupBox6.ResumeLayout(false);
            CustomGroupBox1.ResumeLayout(false);
            CustomGroupBox7.ResumeLayout(false);
            CustomGroupBox7.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)PictureBox8).EndInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox1).EndInit();
            CustomGroupBox5.ResumeLayout(false);
            Load += new EventHandler(Form27_Split_text_bystrings_Load);
            KeyDown += new System.Windows.Forms.KeyEventHandler(Form27_Split_text_bystrings_KeyDown);
            Closing += new System.ComponentModel.CancelEventHandler(Form27_Split_text_bystrings_Closing);
            Disposed += new EventHandler(Form27_Split_text_bystrings_Disposed);
            Shown += new EventHandler(Form27_Split_text_bystrings_Shown);
            ResumeLayout(false);
            PerformLayout();

        }

        internal System.Windows.Forms.RadioButton RB_ending_point;
        internal System.Windows.Forms.CheckBox CB_separators_finaloutput;
        internal CustomGroupBox CustomGroupBox2;
        internal System.Windows.Forms.RadioButton RB_starting_point;
        internal System.Windows.Forms.CheckBox CB_consecutive_separators;
        internal CustomPanel Panel_ExpectedOutput;
        internal System.Windows.Forms.PictureBox PictureBox8;
        internal System.Windows.Forms.PictureBox PictureBox1;
        internal System.Windows.Forms.RadioButton RB_columns;
        internal System.Windows.Forms.RadioButton RB_rows;
        internal CustomGroupBox CustomGroupBox7;
        internal CustomPanel Panel_InputRange;
        internal System.Windows.Forms.ToolTip ToolTip1;
        internal System.Windows.Forms.PictureBox Info;
        internal System.Windows.Forms.Label Label1;
        internal System.Windows.Forms.PictureBox AutoSelection;
        internal System.Windows.Forms.PictureBox Selection;
        internal System.Windows.Forms.PictureBox PictureBox7;
        internal System.Windows.Forms.ComboBox ComboBox1;
        internal System.Windows.Forms.Button Btn_OK;
        internal System.Windows.Forms.Button Btn_Cancel;
        internal System.Windows.Forms.CheckBox CB_formatting;
        internal CustomGroupBox CustomGroupBox6;
        internal System.Windows.Forms.CheckBox CB_backup;
        internal CustomGroupBox CustomGroupBox1;
        internal CustomGroupBox CustomGroupBox5;
        internal System.Windows.Forms.TextBox TB_source_range;
        internal System.Windows.Forms.Label Label2;
        internal System.Windows.Forms.PictureBox PictureBox2;
        internal System.Windows.Forms.TextBox TB_TypeString;
        internal CustomGroupBox CustomGroupBox4;
        internal System.Windows.Forms.PictureBox PictureBox3;
        internal System.Windows.Forms.PictureBox PictureBox4;
    }
}