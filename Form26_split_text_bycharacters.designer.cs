using System;

namespace VSTO_Addins
{
    [Microsoft.VisualBasic.CompilerServices.DesignerGenerated()]
    public partial class Form26_split_text_bycharacters : System.Windows.Forms.Form
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
            var resources = new System.ComponentModel.ComponentResourceManager(typeof(Form26_split_text_bycharacters));
            Label1 = new System.Windows.Forms.Label();
            AutoSelection = new System.Windows.Forms.PictureBox();
            AutoSelection.Click += new EventHandler(AutoSelection_Click);
            AutoSelection.GotFocus += new EventHandler(AutoSelection_GotFocus);
            Info = new System.Windows.Forms.PictureBox();
            ToolTip1 = new System.Windows.Forms.ToolTip(components);
            TB_source_range = new System.Windows.Forms.TextBox();
            TB_source_range.TextChanged += new EventHandler(TB_source_range_TextChanged);
            TB_source_range.GotFocus += new EventHandler(TB_source_range_GotFocus);
            Selection = new System.Windows.Forms.PictureBox();
            Selection.Click += new EventHandler(Selection_Click);
            Selection.GotFocus += new EventHandler(Selection_GotFocus);
            CheckBox2 = new System.Windows.Forms.CheckBox();
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
            PictureBox7 = new System.Windows.Forms.PictureBox();
            CB_consecute_separators = new System.Windows.Forms.CheckBox();
            CB_consecute_separators.CheckedChanged += new EventHandler(CB_consecute_separators_CheckedChanged);
            CustomGroupBox2 = new CustomGroupBox();
            PictureBox2 = new System.Windows.Forms.PictureBox();
            PictureBox3 = new System.Windows.Forms.PictureBox();
            RB_ending_point = new System.Windows.Forms.RadioButton();
            RB_ending_point.CheckedChanged += new EventHandler(RB_ending_point_CheckedChanged);
            RB_starting_point = new System.Windows.Forms.RadioButton();
            RB_starting_point.CheckedChanged += new EventHandler(RB_starting_point_CheckedChanged);
            CB_separators_finaloutput = new System.Windows.Forms.CheckBox();
            CB_separators_finaloutput.CheckedChanged += new EventHandler(CB_separators_finaloutput_CheckedChanged);
            CustomGroupBox4 = new CustomGroupBox();
            CustomGroupBox8 = new CustomGroupBox();
            ComboBox2 = new System.Windows.Forms.ComboBox();
            ComboBox2.TextChanged += new EventHandler(ComboBox2_TextChanged);
            PictureBox10 = new System.Windows.Forms.PictureBox();
            PictureBox6 = new System.Windows.Forms.PictureBox();
            PictureBox5 = new System.Windows.Forms.PictureBox();
            PictureBox4 = new System.Windows.Forms.PictureBox();
            PictureBox11 = new System.Windows.Forms.PictureBox();
            VScrollBar1 = new System.Windows.Forms.VScrollBar();
            RB_others = new System.Windows.Forms.RadioButton();
            RB_others.CheckedChanged += new EventHandler(RB_others_CheckedChanged);
            RB_semicolon = new System.Windows.Forms.RadioButton();
            RB_semicolon.CheckedChanged += new EventHandler(RB_semicolon_CheckedChanged);
            RB_numbertext = new System.Windows.Forms.RadioButton();
            RB_numbertext.CheckedChanged += new EventHandler(RB_numbertext_CheckedChanged);
            RB_newline = new System.Windows.Forms.RadioButton();
            RB_newline.CheckedChanged += new EventHandler(RB_newline_CheckedChanged);
            RB_space = new System.Windows.Forms.RadioButton();
            RB_space.CheckedChanged += new EventHandler(RB_space_CheckedChanged);
            TextBox3 = new System.Windows.Forms.TextBox();
            TextBox3.TextChanged += new EventHandler(TextBox3_TextChanged);
            RB_width = new System.Windows.Forms.RadioButton();
            RB_width.CheckedChanged += new EventHandler(RB_width_CheckedChanged);
            CustomGroupBox5 = new CustomGroupBox();
            Panel_InputRange = new CustomPanel();
            CustomGroupBox1 = new CustomGroupBox();
            CustomGroupBox7 = new CustomGroupBox();
            PictureBox8 = new System.Windows.Forms.PictureBox();
            PictureBox1 = new System.Windows.Forms.PictureBox();
            RB_columns = new System.Windows.Forms.RadioButton();
            RB_columns.CheckedChanged += new EventHandler(RB_columns_CheckedChanged);
            RB_rows = new System.Windows.Forms.RadioButton();
            RB_rows.CheckedChanged += new EventHandler(RB_rows_CheckedChanged);
            CustomGroupBox6 = new CustomGroupBox();
            Panel_ExpectedOutput = new CustomPanel();
            ((System.ComponentModel.ISupportInitialize)AutoSelection).BeginInit();
            ((System.ComponentModel.ISupportInitialize)Info).BeginInit();
            ((System.ComponentModel.ISupportInitialize)Selection).BeginInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox7).BeginInit();
            CustomGroupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)PictureBox2).BeginInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox3).BeginInit();
            CustomGroupBox4.SuspendLayout();
            CustomGroupBox8.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)PictureBox10).BeginInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox6).BeginInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox5).BeginInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox4).BeginInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox11).BeginInit();
            CustomGroupBox5.SuspendLayout();
            CustomGroupBox1.SuspendLayout();
            CustomGroupBox7.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)PictureBox8).BeginInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox1).BeginInit();
            CustomGroupBox6.SuspendLayout();
            SuspendLayout();
            // 
            // Label1
            // 
            Label1.AutoSize = true;
            Label1.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            Label1.Location = new System.Drawing.Point(15, 15);
            Label1.Name = "Label1";
            Label1.Size = new System.Drawing.Size(98, 17);
            Label1.TabIndex = 290;
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
            AutoSelection.TabIndex = 301;
            AutoSelection.TabStop = false;
            // 
            // Info
            // 
            Info.Image = (System.Drawing.Image)resources.GetObject("Info.Image");
            Info.Location = new System.Drawing.Point(118, 15);
            Info.Name = "Info";
            Info.Size = new System.Drawing.Size(20, 20);
            Info.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            Info.TabIndex = 305;
            Info.TabStop = false;
            ToolTip1.SetToolTip(Info, "Please, select single column");
            // 
            // TB_source_range
            // 
            TB_source_range.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            TB_source_range.Cursor = System.Windows.Forms.Cursors.IBeam;
            TB_source_range.Font = new System.Drawing.Font("Segoe UI", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            TB_source_range.Location = new System.Drawing.Point(15, 42);
            TB_source_range.Name = "TB_source_range";
            TB_source_range.Size = new System.Drawing.Size(262, 25);
            TB_source_range.TabIndex = 300;
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
            Selection.TabIndex = 302;
            Selection.TabStop = false;
            // 
            // CheckBox2
            // 
            CheckBox2.AutoSize = true;
            CheckBox2.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            CheckBox2.Location = new System.Drawing.Point(15, 528);
            CheckBox2.Name = "CheckBox2";
            CheckBox2.Size = new System.Drawing.Size(257, 21);
            CheckBox2.TabIndex = 293;
            CheckBox2.Text = "Create a copy of the original worksheet";
            CheckBox2.UseVisualStyleBackColor = true;
            // 
            // ComboBox1
            // 
            ComboBox1.Font = new System.Drawing.Font("Segoe UI", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            ComboBox1.FormattingEnabled = true;
            ComboBox1.Items.AddRange(new object[] { "SOFTEKO", "About Us", "Help", "Feedback" });
            ComboBox1.Location = new System.Drawing.Point(15, 560);
            ComboBox1.Name = "ComboBox1";
            ComboBox1.Size = new System.Drawing.Size(154, 25);
            ComboBox1.TabIndex = 294;
            ComboBox1.Text = "SOFTEKO";
            // 
            // Btn_OK
            // 
            Btn_OK.BackColor = System.Drawing.Color.White;
            Btn_OK.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            Btn_OK.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            Btn_OK.Location = new System.Drawing.Point(481, 558);
            Btn_OK.Name = "Btn_OK";
            Btn_OK.Size = new System.Drawing.Size(62, 26);
            Btn_OK.TabIndex = 298;
            Btn_OK.Text = "OK";
            Btn_OK.UseVisualStyleBackColor = false;
            // 
            // Btn_Cancel
            // 
            Btn_Cancel.BackColor = System.Drawing.Color.White;
            Btn_Cancel.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            Btn_Cancel.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            Btn_Cancel.Location = new System.Drawing.Point(559, 558);
            Btn_Cancel.Name = "Btn_Cancel";
            Btn_Cancel.Size = new System.Drawing.Size(62, 26);
            Btn_Cancel.TabIndex = 297;
            Btn_Cancel.Text = "Cancel";
            Btn_Cancel.UseVisualStyleBackColor = false;
            // 
            // CB_formatting
            // 
            CB_formatting.AutoSize = true;
            CB_formatting.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            CB_formatting.Location = new System.Drawing.Point(15, 168);
            CB_formatting.Name = "CB_formatting";
            CB_formatting.Size = new System.Drawing.Size(122, 21);
            CB_formatting.TabIndex = 292;
            CB_formatting.Text = "Keep formatting";
            CB_formatting.UseVisualStyleBackColor = true;
            // 
            // PictureBox7
            // 
            PictureBox7.Image = (System.Drawing.Image)resources.GetObject("PictureBox7.Image");
            PictureBox7.Location = new System.Drawing.Point(435, 240);
            PictureBox7.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            PictureBox7.Name = "PictureBox7";
            PictureBox7.Size = new System.Drawing.Size(65, 65);
            PictureBox7.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            PictureBox7.TabIndex = 299;
            PictureBox7.TabStop = false;
            // 
            // CB_consecute_separators
            // 
            CB_consecute_separators.AutoSize = true;
            CB_consecute_separators.Enabled = false;
            CB_consecute_separators.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            CB_consecute_separators.Location = new System.Drawing.Point(15, 396);
            CB_consecute_separators.Name = "CB_consecute_separators";
            CB_consecute_separators.Size = new System.Drawing.Size(237, 21);
            CB_consecute_separators.TabIndex = 306;
            CB_consecute_separators.Text = "Treat consecutive separators as one";
            CB_consecute_separators.UseVisualStyleBackColor = true;
            // 
            // CustomGroupBox2
            // 
            CustomGroupBox2.BackColor = System.Drawing.Color.White;
            CustomGroupBox2.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            CustomGroupBox2.Controls.Add(PictureBox2);
            CustomGroupBox2.Controls.Add(PictureBox3);
            CustomGroupBox2.Controls.Add(RB_ending_point);
            CustomGroupBox2.Controls.Add(RB_starting_point);
            CustomGroupBox2.Controls.Add(CB_separators_finaloutput);
            CustomGroupBox2.Enabled = false;
            CustomGroupBox2.Location = new System.Drawing.Point(15, 425);
            CustomGroupBox2.Name = "CustomGroupBox2";
            CustomGroupBox2.Size = new System.Drawing.Size(257, 92);
            CustomGroupBox2.TabIndex = 307;
            CustomGroupBox2.TabStop = false;
            // 
            // PictureBox2
            // 
            PictureBox2.Enabled = false;
            PictureBox2.Image = (System.Drawing.Image)resources.GetObject("PictureBox2.Image");
            PictureBox2.Location = new System.Drawing.Point(226, 34);
            PictureBox2.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            PictureBox2.Name = "PictureBox2";
            PictureBox2.Size = new System.Drawing.Size(20, 20);
            PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            PictureBox2.TabIndex = 308;
            PictureBox2.TabStop = false;
            // 
            // PictureBox3
            // 
            PictureBox3.Enabled = false;
            PictureBox3.Image = (System.Drawing.Image)resources.GetObject("PictureBox3.Image");
            PictureBox3.Location = new System.Drawing.Point(226, 60);
            PictureBox3.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            PictureBox3.Name = "PictureBox3";
            PictureBox3.Size = new System.Drawing.Size(20, 20);
            PictureBox3.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            PictureBox3.TabIndex = 309;
            PictureBox3.TabStop = false;
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
            CustomGroupBox4.Controls.Add(CustomGroupBox8);
            CustomGroupBox4.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 0);
            CustomGroupBox4.Location = new System.Drawing.Point(15, 195);
            CustomGroupBox4.Name = "CustomGroupBox4";
            CustomGroupBox4.Size = new System.Drawing.Size(260, 195);
            CustomGroupBox4.TabIndex = 303;
            CustomGroupBox4.TabStop = false;
            CustomGroupBox4.Text = "Select Separator";
            // 
            // CustomGroupBox8
            // 
            CustomGroupBox8.BackColor = System.Drawing.Color.White;
            CustomGroupBox8.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            CustomGroupBox8.Controls.Add(ComboBox2);
            CustomGroupBox8.Controls.Add(PictureBox10);
            CustomGroupBox8.Controls.Add(PictureBox6);
            CustomGroupBox8.Controls.Add(PictureBox5);
            CustomGroupBox8.Controls.Add(PictureBox4);
            CustomGroupBox8.Controls.Add(PictureBox11);
            CustomGroupBox8.Controls.Add(VScrollBar1);
            CustomGroupBox8.Controls.Add(RB_others);
            CustomGroupBox8.Controls.Add(RB_semicolon);
            CustomGroupBox8.Controls.Add(RB_numbertext);
            CustomGroupBox8.Controls.Add(RB_newline);
            CustomGroupBox8.Controls.Add(RB_space);
            CustomGroupBox8.Controls.Add(TextBox3);
            CustomGroupBox8.Controls.Add(RB_width);
            CustomGroupBox8.Location = new System.Drawing.Point(1, 22);
            CustomGroupBox8.Name = "CustomGroupBox8";
            CustomGroupBox8.Size = new System.Drawing.Size(259, 173);
            CustomGroupBox8.TabIndex = 0;
            CustomGroupBox8.TabStop = false;
            // 
            // ComboBox2
            // 
            ComboBox2.FormattingEnabled = true;
            ComboBox2.Items.AddRange(new object[] { ",", "/", "." });
            ComboBox2.Location = new System.Drawing.Point(112, 104);
            ComboBox2.Name = "ComboBox2";
            ComboBox2.Size = new System.Drawing.Size(133, 25);
            ComboBox2.TabIndex = 275;
            // 
            // PictureBox10
            // 
            PictureBox10.Image = (System.Drawing.Image)resources.GetObject("PictureBox10.Image");
            PictureBox10.Location = new System.Drawing.Point(225, 6);
            PictureBox10.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            PictureBox10.Name = "PictureBox10";
            PictureBox10.Size = new System.Drawing.Size(20, 20);
            PictureBox10.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            PictureBox10.TabIndex = 272;
            PictureBox10.TabStop = false;
            // 
            // PictureBox6
            // 
            PictureBox6.Image = (System.Drawing.Image)resources.GetObject("PictureBox6.Image");
            PictureBox6.Location = new System.Drawing.Point(225, 30);
            PictureBox6.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            PictureBox6.Name = "PictureBox6";
            PictureBox6.Size = new System.Drawing.Size(20, 20);
            PictureBox6.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            PictureBox6.TabIndex = 272;
            PictureBox6.TabStop = false;
            // 
            // PictureBox5
            // 
            PictureBox5.Image = (System.Drawing.Image)resources.GetObject("PictureBox5.Image");
            PictureBox5.Location = new System.Drawing.Point(225, 54);
            PictureBox5.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            PictureBox5.Name = "PictureBox5";
            PictureBox5.Size = new System.Drawing.Size(20, 20);
            PictureBox5.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            PictureBox5.TabIndex = 274;
            PictureBox5.TabStop = false;
            // 
            // PictureBox4
            // 
            PictureBox4.Image = (System.Drawing.Image)resources.GetObject("PictureBox4.Image");
            PictureBox4.Location = new System.Drawing.Point(225, 78);
            PictureBox4.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            PictureBox4.Name = "PictureBox4";
            PictureBox4.Size = new System.Drawing.Size(20, 20);
            PictureBox4.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            PictureBox4.TabIndex = 273;
            PictureBox4.TabStop = false;
            // 
            // PictureBox11
            // 
            PictureBox11.Image = (System.Drawing.Image)resources.GetObject("PictureBox11.Image");
            PictureBox11.Location = new System.Drawing.Point(225, 143);
            PictureBox11.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            PictureBox11.Name = "PictureBox11";
            PictureBox11.Size = new System.Drawing.Size(20, 20);
            PictureBox11.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            PictureBox11.TabIndex = 271;
            PictureBox11.TabStop = false;
            // 
            // VScrollBar1
            // 
            VScrollBar1.Location = new System.Drawing.Point(190, 141);
            VScrollBar1.Name = "VScrollBar1";
            VScrollBar1.Size = new System.Drawing.Size(21, 21);
            VScrollBar1.TabIndex = 236;
            // 
            // RB_others
            // 
            RB_others.AutoSize = true;
            RB_others.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            RB_others.Location = new System.Drawing.Point(8, 108);
            RB_others.Name = "RB_others";
            RB_others.Size = new System.Drawing.Size(72, 21);
            RB_others.TabIndex = 233;
            RB_others.TabStop = true;
            RB_others.Text = "Others :";
            RB_others.UseVisualStyleBackColor = true;
            // 
            // RB_semicolon
            // 
            RB_semicolon.AutoSize = true;
            RB_semicolon.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            RB_semicolon.Location = new System.Drawing.Point(8, 78);
            RB_semicolon.Name = "RB_semicolon";
            RB_semicolon.Size = new System.Drawing.Size(86, 21);
            RB_semicolon.TabIndex = 232;
            RB_semicolon.TabStop = true;
            RB_semicolon.Text = "Semicolon";
            RB_semicolon.UseVisualStyleBackColor = true;
            // 
            // RB_numbertext
            // 
            RB_numbertext.AutoSize = true;
            RB_numbertext.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            RB_numbertext.Location = new System.Drawing.Point(8, 54);
            RB_numbertext.Name = "RB_numbertext";
            RB_numbertext.Size = new System.Drawing.Size(125, 21);
            RB_numbertext.TabIndex = 231;
            RB_numbertext.TabStop = true;
            RB_numbertext.Text = "Number and text";
            RB_numbertext.UseVisualStyleBackColor = true;
            // 
            // RB_newline
            // 
            RB_newline.AutoSize = true;
            RB_newline.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            RB_newline.Location = new System.Drawing.Point(8, 30);
            RB_newline.Name = "RB_newline";
            RB_newline.Size = new System.Drawing.Size(76, 21);
            RB_newline.TabIndex = 1;
            RB_newline.TabStop = true;
            RB_newline.Text = "New line";
            RB_newline.UseVisualStyleBackColor = true;
            // 
            // RB_space
            // 
            RB_space.AutoSize = true;
            RB_space.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            RB_space.Location = new System.Drawing.Point(8, 6);
            RB_space.Name = "RB_space";
            RB_space.Size = new System.Drawing.Size(61, 21);
            RB_space.TabIndex = 0;
            RB_space.TabStop = true;
            RB_space.Text = "Space";
            RB_space.UseVisualStyleBackColor = true;
            // 
            // TextBox3
            // 
            TextBox3.Font = new System.Drawing.Font("Segoe UI", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            TextBox3.Location = new System.Drawing.Point(112, 140);
            TextBox3.Name = "TextBox3";
            TextBox3.Size = new System.Drawing.Size(100, 25);
            TextBox3.TabIndex = 237;
            // 
            // RB_width
            // 
            RB_width.AutoSize = true;
            RB_width.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            RB_width.Location = new System.Drawing.Point(8, 141);
            RB_width.Name = "RB_width";
            RB_width.Size = new System.Drawing.Size(105, 21);
            RB_width.TabIndex = 234;
            RB_width.TabStop = true;
            RB_width.Text = "Define width :";
            RB_width.UseVisualStyleBackColor = true;
            // 
            // CustomGroupBox5
            // 
            CustomGroupBox5.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            CustomGroupBox5.Controls.Add(Panel_InputRange);
            CustomGroupBox5.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 0);
            CustomGroupBox5.Location = new System.Drawing.Point(315, 17);
            CustomGroupBox5.Name = "CustomGroupBox5";
            CustomGroupBox5.Size = new System.Drawing.Size(302, 200);
            CustomGroupBox5.TabIndex = 295;
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
            Panel_InputRange.Size = new System.Drawing.Size(300, 170);
            Panel_InputRange.TabIndex = 0;
            // 
            // CustomGroupBox1
            // 
            CustomGroupBox1.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            CustomGroupBox1.Controls.Add(CustomGroupBox7);
            CustomGroupBox1.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 0);
            CustomGroupBox1.Location = new System.Drawing.Point(15, 76);
            CustomGroupBox1.Name = "CustomGroupBox1";
            CustomGroupBox1.Size = new System.Drawing.Size(260, 84);
            CustomGroupBox1.TabIndex = 291;
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
            RB_columns.Size = new System.Drawing.Size(167, 21);
            RB_columns.TabIndex = 1;
            RB_columns.Text = "Split range into columns";
            RB_columns.UseVisualStyleBackColor = true;
            // 
            // RB_rows
            // 
            RB_rows.AutoSize = true;
            RB_rows.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            RB_rows.Location = new System.Drawing.Point(8, 6);
            RB_rows.Name = "RB_rows";
            RB_rows.Size = new System.Drawing.Size(147, 21);
            RB_rows.TabIndex = 0;
            RB_rows.Text = "Split range into rows";
            RB_rows.UseVisualStyleBackColor = true;
            // 
            // CustomGroupBox6
            // 
            CustomGroupBox6.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            CustomGroupBox6.Controls.Add(Panel_ExpectedOutput);
            CustomGroupBox6.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 0);
            CustomGroupBox6.Location = new System.Drawing.Point(315, 322);
            CustomGroupBox6.Name = "CustomGroupBox6";
            CustomGroupBox6.Size = new System.Drawing.Size(302, 200);
            CustomGroupBox6.TabIndex = 296;
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
            Panel_ExpectedOutput.Size = new System.Drawing.Size(300, 170);
            Panel_ExpectedOutput.TabIndex = 11;
            // 
            // Form26_split_text_bycharacters
            // 
            AutoScaleDimensions = new System.Drawing.SizeF(6.0f, 13.0f);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            BackColor = System.Drawing.SystemColors.Control;
            ClientSize = new System.Drawing.Size(643, 607);
            Controls.Add(CustomGroupBox2);
            Controls.Add(CB_consecute_separators);
            Controls.Add(Label1);
            Controls.Add(AutoSelection);
            Controls.Add(CustomGroupBox4);
            Controls.Add(Info);
            Controls.Add(Selection);
            Controls.Add(CustomGroupBox5);
            Controls.Add(CheckBox2);
            Controls.Add(CustomGroupBox1);
            Controls.Add(ComboBox1);
            Controls.Add(Btn_OK);
            Controls.Add(Btn_Cancel);
            Controls.Add(CB_formatting);
            Controls.Add(CustomGroupBox6);
            Controls.Add(PictureBox7);
            Controls.Add(TB_source_range);
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "Form26_split_text_bycharacters";
            Text = "Split Text by Characters";
            ((System.ComponentModel.ISupportInitialize)AutoSelection).EndInit();
            ((System.ComponentModel.ISupportInitialize)Info).EndInit();
            ((System.ComponentModel.ISupportInitialize)Selection).EndInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox7).EndInit();
            CustomGroupBox2.ResumeLayout(false);
            CustomGroupBox2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)PictureBox2).EndInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox3).EndInit();
            CustomGroupBox4.ResumeLayout(false);
            CustomGroupBox8.ResumeLayout(false);
            CustomGroupBox8.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)PictureBox10).EndInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox6).EndInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox5).EndInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox4).EndInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox11).EndInit();
            CustomGroupBox5.ResumeLayout(false);
            CustomGroupBox1.ResumeLayout(false);
            CustomGroupBox7.ResumeLayout(false);
            CustomGroupBox7.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)PictureBox8).EndInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox1).EndInit();
            CustomGroupBox6.ResumeLayout(false);
            Load += new EventHandler(Form26_split_text_bycharacters_Load);
            KeyDown += new System.Windows.Forms.KeyEventHandler(Form26_split_text_bycharacters_KeyDown);
            Closing += new System.ComponentModel.CancelEventHandler(Form26_split_text_bycharacters_Closing);
            Disposed += new EventHandler(Form26_split_text_bycharacters_Disposed);
            Shown += new EventHandler(Form26_split_text_bycharacters_Shown);
            ResumeLayout(false);
            PerformLayout();

        }

        internal System.Windows.Forms.Label Label1;
        internal System.Windows.Forms.PictureBox AutoSelection;
        internal CustomGroupBox CustomGroupBox4;
        internal CustomGroupBox CustomGroupBox8;
        internal System.Windows.Forms.ComboBox ComboBox2;
        internal System.Windows.Forms.PictureBox PictureBox10;
        internal System.Windows.Forms.PictureBox PictureBox6;
        internal System.Windows.Forms.PictureBox PictureBox5;
        internal System.Windows.Forms.PictureBox PictureBox4;
        internal System.Windows.Forms.PictureBox PictureBox11;
        internal System.Windows.Forms.VScrollBar VScrollBar1;
        internal System.Windows.Forms.RadioButton RB_others;
        internal System.Windows.Forms.RadioButton RB_semicolon;
        internal System.Windows.Forms.RadioButton RB_numbertext;
        internal System.Windows.Forms.RadioButton RB_newline;
        internal System.Windows.Forms.RadioButton RB_space;
        internal System.Windows.Forms.TextBox TextBox3;
        internal System.Windows.Forms.RadioButton RB_width;
        internal System.Windows.Forms.PictureBox Info;
        internal System.Windows.Forms.ToolTip ToolTip1;
        internal System.Windows.Forms.TextBox TB_source_range;
        internal System.Windows.Forms.PictureBox Selection;
        internal CustomGroupBox CustomGroupBox5;
        internal CustomPanel Panel_InputRange;
        internal System.Windows.Forms.CheckBox CheckBox2;
        internal CustomGroupBox CustomGroupBox1;
        internal CustomGroupBox CustomGroupBox7;
        internal System.Windows.Forms.PictureBox PictureBox8;
        internal System.Windows.Forms.PictureBox PictureBox1;
        internal System.Windows.Forms.RadioButton RB_columns;
        internal System.Windows.Forms.RadioButton RB_rows;
        internal System.Windows.Forms.ComboBox ComboBox1;
        internal System.Windows.Forms.Button Btn_OK;
        internal System.Windows.Forms.Button Btn_Cancel;
        internal System.Windows.Forms.CheckBox CB_formatting;
        internal CustomGroupBox CustomGroupBox6;
        internal CustomPanel Panel_ExpectedOutput;
        internal System.Windows.Forms.PictureBox PictureBox7;
        internal System.Windows.Forms.CheckBox CB_consecute_separators;
        internal CustomGroupBox CustomGroupBox2;
        internal System.Windows.Forms.RadioButton RB_ending_point;
        internal System.Windows.Forms.RadioButton RB_starting_point;
        internal System.Windows.Forms.CheckBox CB_separators_finaloutput;
        internal System.Windows.Forms.PictureBox PictureBox2;
        internal System.Windows.Forms.PictureBox PictureBox3;
    }
}