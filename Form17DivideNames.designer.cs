using System;

namespace VSTO_Addins
{
    [Microsoft.VisualBasic.CompilerServices.DesignerGenerated()]
    public partial class Form17DivideNames : System.Windows.Forms.Form
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
            var resources = new System.ComponentModel.ComponentResourceManager(typeof(Form17DivideNames));
            Selection = new System.Windows.Forms.PictureBox();
            Selection.Click += new EventHandler(Selection_Click);
            AutoSelection = new System.Windows.Forms.PictureBox();
            AutoSelection.Click += new EventHandler(AutoSelection_Click);
            txtSourceRange = new System.Windows.Forms.TextBox();
            txtSourceRange.TextChanged += new EventHandler(txtSourceRange_TextChanged);
            txtSourceRange.GotFocus += new EventHandler(txtSourceRange_GotFocus);
            Label1 = new System.Windows.Forms.Label();
            PictureBox7 = new System.Windows.Forms.PictureBox();
            btnOK = new System.Windows.Forms.Button();
            btnOK.Click += new EventHandler(btnOK_Click);
            btnCancel = new System.Windows.Forms.Button();
            btnCancel.Click += new EventHandler(btnCancel_Click);
            ComboBox1 = new System.Windows.Forms.ComboBox();
            CB_Backup_Sheet = new System.Windows.Forms.CheckBox();
            CB_Keep_Formatting = new System.Windows.Forms.CheckBox();
            CB_Keep_Formatting.CheckedChanged += new EventHandler(CB_Keep_Formatting_CheckedChanged);
            CB_Add_Header = new System.Windows.Forms.CheckBox();
            CB_Add_Header.CheckedChanged += new EventHandler(CB_Add_Header_CheckedChanged);
            CustomGroupBox4 = new CustomGroupBox();
            CustomGroupBox10 = new CustomGroupBox();
            PictureBox2 = new System.Windows.Forms.PictureBox();
            destinationSelection = new System.Windows.Forms.PictureBox();
            destinationSelection.Click += new EventHandler(destinationSelection_Click);
            txtDestRange = new System.Windows.Forms.TextBox();
            txtDestRange.TextChanged += new EventHandler(txtDestRange_TextChanged);
            txtDestRange.GotFocus += new EventHandler(txtDestRange_GotFocus);
            lbl_destRange_Selection = new System.Windows.Forms.Label();
            RB_Different_Range = new System.Windows.Forms.RadioButton();
            RB_Same_As_Source_Range = new System.Windows.Forms.RadioButton();
            RB_Same_As_Source_Range.CheckedChanged += new EventHandler(RB_Same_As_Source_Range_CheckedChanged);
            CustomGroupBox5 = new CustomGroupBox();
            CustomPanel1 = new CustomPanel();
            CustomGroupBox6 = new CustomGroupBox();
            CustomPanel2 = new CustomPanel();
            CustomGroupBox1 = new CustomGroupBox();
            CustomGroupBox7 = new CustomGroupBox();
            PictureBox11 = new System.Windows.Forms.PictureBox();
            CB_Select_All = new System.Windows.Forms.CheckBox();
            CB_Select_All.CheckedChanged += new EventHandler(CB_Select_All_CheckedChanged);
            PictureBox8 = new System.Windows.Forms.PictureBox();
            PictureBox9 = new System.Windows.Forms.PictureBox();
            PictureBox10 = new System.Windows.Forms.PictureBox();
            PictureBox6 = new System.Windows.Forms.PictureBox();
            PictureBox4 = new System.Windows.Forms.PictureBox();
            PictureBox3 = new System.Windows.Forms.PictureBox();
            PictureBox5 = new System.Windows.Forms.PictureBox();
            CB_Name_Suffix = new System.Windows.Forms.CheckBox();
            CB_Name_Suffix.CheckedChanged += new EventHandler(CB_Name_Suffix_CheckedChanged);
            CB_Title = new System.Windows.Forms.CheckBox();
            CB_Title.CheckedChanged += new EventHandler(CB_Title_CheckedChanged);
            CB_Name_Abbreviations = new System.Windows.Forms.CheckBox();
            CB_Name_Abbreviations.CheckedChanged += new EventHandler(CB_Name_Abbreviations_CheckedChanged);
            CB_Last_Name = new System.Windows.Forms.CheckBox();
            CB_Last_Name.CheckedChanged += new EventHandler(CB_Last_Name_CheckedChanged);
            CB_Last_Name_Prefix = new System.Windows.Forms.CheckBox();
            CB_Last_Name_Prefix.CheckedChanged += new EventHandler(CB_Last_Name_Prefix_CheckedChanged);
            CB_Middle_Name = new System.Windows.Forms.CheckBox();
            CB_Middle_Name.CheckedChanged += new EventHandler(CB_Middle_Name_CheckedChanged);
            CB_First_Name = new System.Windows.Forms.CheckBox();
            CB_First_Name.CheckedChanged += new EventHandler(CB_First_Name_CheckedChanged);
            ((System.ComponentModel.ISupportInitialize)Selection).BeginInit();
            ((System.ComponentModel.ISupportInitialize)AutoSelection).BeginInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox7).BeginInit();
            CustomGroupBox4.SuspendLayout();
            CustomGroupBox10.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)PictureBox2).BeginInit();
            ((System.ComponentModel.ISupportInitialize)destinationSelection).BeginInit();
            CustomGroupBox5.SuspendLayout();
            CustomGroupBox6.SuspendLayout();
            CustomGroupBox1.SuspendLayout();
            CustomGroupBox7.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)PictureBox11).BeginInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox8).BeginInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox9).BeginInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox10).BeginInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox6).BeginInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox4).BeginInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox3).BeginInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox5).BeginInit();
            SuspendLayout();
            // 
            // Selection
            // 
            Selection.BackColor = System.Drawing.Color.White;
            Selection.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            Selection.Image = (System.Drawing.Image)resources.GetObject("Selection.Image");
            Selection.Location = new System.Drawing.Point(227, 40);
            Selection.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            Selection.Name = "Selection";
            Selection.Size = new System.Drawing.Size(24, 25);
            Selection.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            Selection.TabIndex = 224;
            Selection.TabStop = false;
            // 
            // AutoSelection
            // 
            AutoSelection.BackColor = System.Drawing.Color.White;
            AutoSelection.Image = (System.Drawing.Image)resources.GetObject("AutoSelection.Image");
            AutoSelection.Location = new System.Drawing.Point(198, 41);
            AutoSelection.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            AutoSelection.Name = "AutoSelection";
            AutoSelection.Size = new System.Drawing.Size(24, 23);
            AutoSelection.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            AutoSelection.TabIndex = 223;
            AutoSelection.TabStop = false;
            // 
            // txtSourceRange
            // 
            txtSourceRange.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            txtSourceRange.Cursor = System.Windows.Forms.Cursors.IBeam;
            txtSourceRange.Font = new System.Drawing.Font("Segoe UI", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            txtSourceRange.Location = new System.Drawing.Point(15, 40);
            txtSourceRange.Name = "txtSourceRange";
            txtSourceRange.Size = new System.Drawing.Size(236, 25);
            txtSourceRange.TabIndex = 222;
            // 
            // Label1
            // 
            Label1.AutoSize = true;
            Label1.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 0);
            Label1.Location = new System.Drawing.Point(15, 14);
            Label1.Name = "Label1";
            Label1.Size = new System.Drawing.Size(98, 17);
            Label1.TabIndex = 221;
            Label1.Text = "Source Range :";
            Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // PictureBox7
            // 
            PictureBox7.Image = (System.Drawing.Image)resources.GetObject("PictureBox7.Image");
            PictureBox7.Location = new System.Drawing.Point(380, 209);
            PictureBox7.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            PictureBox7.Name = "PictureBox7";
            PictureBox7.Size = new System.Drawing.Size(52, 60);
            PictureBox7.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            PictureBox7.TabIndex = 220;
            PictureBox7.TabStop = false;
            // 
            // btnOK
            // 
            btnOK.BackColor = System.Drawing.Color.White;
            btnOK.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            btnOK.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            btnOK.Location = new System.Drawing.Point(394, 506);
            btnOK.Name = "btnOK";
            btnOK.Size = new System.Drawing.Size(62, 26);
            btnOK.TabIndex = 219;
            btnOK.Text = "OK";
            btnOK.UseVisualStyleBackColor = false;
            // 
            // btnCancel
            // 
            btnCancel.BackColor = System.Drawing.Color.White;
            btnCancel.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            btnCancel.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            btnCancel.Location = new System.Drawing.Point(472, 506);
            btnCancel.Name = "btnCancel";
            btnCancel.Size = new System.Drawing.Size(62, 26);
            btnCancel.TabIndex = 218;
            btnCancel.Text = "Cancel";
            btnCancel.UseVisualStyleBackColor = false;
            // 
            // ComboBox1
            // 
            ComboBox1.Font = new System.Drawing.Font("Segoe UI", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            ComboBox1.FormattingEnabled = true;
            ComboBox1.Location = new System.Drawing.Point(15, 506);
            ComboBox1.Name = "ComboBox1";
            ComboBox1.Size = new System.Drawing.Size(154, 25);
            ComboBox1.TabIndex = 215;
            ComboBox1.Text = "SOFTEKO";
            // 
            // CB_Backup_Sheet
            // 
            CB_Backup_Sheet.AutoSize = true;
            CB_Backup_Sheet.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            CB_Backup_Sheet.Location = new System.Drawing.Point(15, 475);
            CB_Backup_Sheet.Name = "CB_Backup_Sheet";
            CB_Backup_Sheet.Size = new System.Drawing.Size(257, 21);
            CB_Backup_Sheet.TabIndex = 214;
            CB_Backup_Sheet.Text = "Create a copy of the original worksheet";
            CB_Backup_Sheet.UseVisualStyleBackColor = true;
            // 
            // CB_Keep_Formatting
            // 
            CB_Keep_Formatting.AutoSize = true;
            CB_Keep_Formatting.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            CB_Keep_Formatting.Location = new System.Drawing.Point(15, 304);
            CB_Keep_Formatting.Name = "CB_Keep_Formatting";
            CB_Keep_Formatting.Size = new System.Drawing.Size(122, 21);
            CB_Keep_Formatting.TabIndex = 226;
            CB_Keep_Formatting.Text = "Keep formatting";
            CB_Keep_Formatting.UseVisualStyleBackColor = true;
            // 
            // CB_Add_Header
            // 
            CB_Add_Header.AutoSize = true;
            CB_Add_Header.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            CB_Add_Header.Location = new System.Drawing.Point(150, 304);
            CB_Add_Header.Name = "CB_Add_Header";
            CB_Add_Header.Size = new System.Drawing.Size(98, 21);
            CB_Add_Header.TabIndex = 227;
            CB_Add_Header.Text = "Add Header";
            CB_Add_Header.UseVisualStyleBackColor = true;
            // 
            // CustomGroupBox4
            // 
            CustomGroupBox4.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            CustomGroupBox4.Controls.Add(CustomGroupBox10);
            CustomGroupBox4.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 0);
            CustomGroupBox4.Location = new System.Drawing.Point(15, 331);
            CustomGroupBox4.Name = "CustomGroupBox4";
            CustomGroupBox4.Size = new System.Drawing.Size(236, 137);
            CustomGroupBox4.TabIndex = 225;
            CustomGroupBox4.TabStop = false;
            CustomGroupBox4.Text = "Destination Range";
            // 
            // CustomGroupBox10
            // 
            CustomGroupBox10.BackColor = System.Drawing.Color.White;
            CustomGroupBox10.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            CustomGroupBox10.Controls.Add(PictureBox2);
            CustomGroupBox10.Controls.Add(destinationSelection);
            CustomGroupBox10.Controls.Add(txtDestRange);
            CustomGroupBox10.Controls.Add(lbl_destRange_Selection);
            CustomGroupBox10.Controls.Add(RB_Different_Range);
            CustomGroupBox10.Controls.Add(RB_Same_As_Source_Range);
            CustomGroupBox10.Location = new System.Drawing.Point(1, 22);
            CustomGroupBox10.Name = "CustomGroupBox10";
            CustomGroupBox10.Size = new System.Drawing.Size(235, 115);
            CustomGroupBox10.TabIndex = 0;
            CustomGroupBox10.TabStop = false;
            // 
            // PictureBox2
            // 
            PictureBox2.Image = (System.Drawing.Image)resources.GetObject("PictureBox2.Image");
            PictureBox2.Location = new System.Drawing.Point(25, 58);
            PictureBox2.Name = "PictureBox2";
            PictureBox2.Size = new System.Drawing.Size(14, 14);
            PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            PictureBox2.TabIndex = 208;
            PictureBox2.TabStop = false;
            // 
            // destinationSelection
            // 
            destinationSelection.BackColor = System.Drawing.Color.White;
            destinationSelection.Image = (System.Drawing.Image)resources.GetObject("destinationSelection.Image");
            destinationSelection.Location = new System.Drawing.Point(193, 81);
            destinationSelection.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            destinationSelection.Name = "destinationSelection";
            destinationSelection.Size = new System.Drawing.Size(24, 23);
            destinationSelection.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            destinationSelection.TabIndex = 207;
            destinationSelection.TabStop = false;
            // 
            // txtDestRange
            // 
            txtDestRange.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            txtDestRange.Cursor = System.Windows.Forms.Cursors.IBeam;
            txtDestRange.Font = new System.Drawing.Font("Segoe UI", 10.0f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            txtDestRange.Location = new System.Drawing.Point(25, 80);
            txtDestRange.Name = "txtDestRange";
            txtDestRange.Size = new System.Drawing.Size(193, 25);
            txtDestRange.TabIndex = 206;
            // 
            // lbl_destRange_Selection
            // 
            lbl_destRange_Selection.AutoSize = true;
            lbl_destRange_Selection.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            lbl_destRange_Selection.Location = new System.Drawing.Point(42, 56);
            lbl_destRange_Selection.Name = "lbl_destRange_Selection";
            lbl_destRange_Selection.Size = new System.Drawing.Size(109, 17);
            lbl_destRange_Selection.TabIndex = 2;
            lbl_destRange_Selection.Text = "Select the range :";
            // 
            // RB_Different_Range
            // 
            RB_Different_Range.AutoSize = true;
            RB_Different_Range.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            RB_Different_Range.Location = new System.Drawing.Point(8, 31);
            RB_Different_Range.Name = "RB_Different_Range";
            RB_Different_Range.Size = new System.Drawing.Size(185, 21);
            RB_Different_Range.TabIndex = 1;
            RB_Different_Range.TabStop = true;
            RB_Different_Range.Text = "Store into a different range";
            RB_Different_Range.UseVisualStyleBackColor = true;
            // 
            // RB_Same_As_Source_Range
            // 
            RB_Same_As_Source_Range.AutoSize = true;
            RB_Same_As_Source_Range.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            RB_Same_As_Source_Range.Location = new System.Drawing.Point(8, 6);
            RB_Same_As_Source_Range.Name = "RB_Same_As_Source_Range";
            RB_Same_As_Source_Range.Size = new System.Drawing.Size(178, 21);
            RB_Same_As_Source_Range.TabIndex = 0;
            RB_Same_As_Source_Range.TabStop = true;
            RB_Same_As_Source_Range.Text = "Same as the source range";
            RB_Same_As_Source_Range.UseVisualStyleBackColor = true;
            // 
            // CustomGroupBox5
            // 
            CustomGroupBox5.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            CustomGroupBox5.Controls.Add(CustomPanel1);
            CustomGroupBox5.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 0);
            CustomGroupBox5.Location = new System.Drawing.Point(283, 22);
            CustomGroupBox5.Name = "CustomGroupBox5";
            CustomGroupBox5.Size = new System.Drawing.Size(252, 180);
            CustomGroupBox5.TabIndex = 216;
            CustomGroupBox5.TabStop = false;
            CustomGroupBox5.Text = "Input Range";
            // 
            // CustomPanel1
            // 
            CustomPanel1.BackColor = System.Drawing.Color.White;
            CustomPanel1.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            CustomPanel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            CustomPanel1.BorderWidth = 1;
            CustomPanel1.Location = new System.Drawing.Point(1, 30);
            CustomPanel1.Name = "CustomPanel1";
            CustomPanel1.Size = new System.Drawing.Size(250, 150);
            CustomPanel1.TabIndex = 0;
            // 
            // CustomGroupBox6
            // 
            CustomGroupBox6.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            CustomGroupBox6.Controls.Add(CustomPanel2);
            CustomGroupBox6.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 0);
            CustomGroupBox6.Location = new System.Drawing.Point(284, 280);
            CustomGroupBox6.Name = "CustomGroupBox6";
            CustomGroupBox6.Size = new System.Drawing.Size(251, 180);
            CustomGroupBox6.TabIndex = 217;
            CustomGroupBox6.TabStop = false;
            CustomGroupBox6.Text = "Expected Output";
            // 
            // CustomPanel2
            // 
            CustomPanel2.BackColor = System.Drawing.Color.White;
            CustomPanel2.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            CustomPanel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            CustomPanel2.BorderWidth = 1;
            CustomPanel2.Location = new System.Drawing.Point(1, 30);
            CustomPanel2.Name = "CustomPanel2";
            CustomPanel2.Size = new System.Drawing.Size(250, 150);
            CustomPanel2.TabIndex = 11;
            // 
            // CustomGroupBox1
            // 
            CustomGroupBox1.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            CustomGroupBox1.Controls.Add(CustomGroupBox7);
            CustomGroupBox1.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 0);
            CustomGroupBox1.Location = new System.Drawing.Point(15, 72);
            CustomGroupBox1.Name = "CustomGroupBox1";
            CustomGroupBox1.Size = new System.Drawing.Size(236, 224);
            CustomGroupBox1.TabIndex = 212;
            CustomGroupBox1.TabStop = false;
            CustomGroupBox1.Text = "Divide by";
            // 
            // CustomGroupBox7
            // 
            CustomGroupBox7.BackColor = System.Drawing.Color.White;
            CustomGroupBox7.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            CustomGroupBox7.Controls.Add(PictureBox11);
            CustomGroupBox7.Controls.Add(CB_Select_All);
            CustomGroupBox7.Controls.Add(PictureBox8);
            CustomGroupBox7.Controls.Add(PictureBox9);
            CustomGroupBox7.Controls.Add(PictureBox10);
            CustomGroupBox7.Controls.Add(PictureBox6);
            CustomGroupBox7.Controls.Add(PictureBox4);
            CustomGroupBox7.Controls.Add(PictureBox3);
            CustomGroupBox7.Controls.Add(PictureBox5);
            CustomGroupBox7.Controls.Add(CB_Name_Suffix);
            CustomGroupBox7.Controls.Add(CB_Title);
            CustomGroupBox7.Controls.Add(CB_Name_Abbreviations);
            CustomGroupBox7.Controls.Add(CB_Last_Name);
            CustomGroupBox7.Controls.Add(CB_Last_Name_Prefix);
            CustomGroupBox7.Controls.Add(CB_Middle_Name);
            CustomGroupBox7.Controls.Add(CB_First_Name);
            CustomGroupBox7.Location = new System.Drawing.Point(1, 22);
            CustomGroupBox7.Name = "CustomGroupBox7";
            CustomGroupBox7.Size = new System.Drawing.Size(235, 202);
            CustomGroupBox7.TabIndex = 0;
            CustomGroupBox7.TabStop = false;
            // 
            // PictureBox11
            // 
            PictureBox11.Image = (System.Drawing.Image)resources.GetObject("PictureBox11.Image");
            PictureBox11.Location = new System.Drawing.Point(194, 175);
            PictureBox11.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            PictureBox11.Name = "PictureBox11";
            PictureBox11.Size = new System.Drawing.Size(20, 20);
            PictureBox11.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            PictureBox11.TabIndex = 217;
            PictureBox11.TabStop = false;
            // 
            // CB_Select_All
            // 
            CB_Select_All.AutoSize = true;
            CB_Select_All.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            CB_Select_All.Location = new System.Drawing.Point(9, 7);
            CB_Select_All.Name = "CB_Select_All";
            CB_Select_All.Size = new System.Drawing.Size(79, 21);
            CB_Select_All.TabIndex = 216;
            CB_Select_All.Text = "Select All";
            CB_Select_All.UseVisualStyleBackColor = true;
            // 
            // PictureBox8
            // 
            PictureBox8.Image = (System.Drawing.Image)resources.GetObject("PictureBox8.Image");
            PictureBox8.Location = new System.Drawing.Point(194, 151);
            PictureBox8.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            PictureBox8.Name = "PictureBox8";
            PictureBox8.Size = new System.Drawing.Size(20, 20);
            PictureBox8.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            PictureBox8.TabIndex = 212;
            PictureBox8.TabStop = false;
            // 
            // PictureBox9
            // 
            PictureBox9.Image = (System.Drawing.Image)resources.GetObject("PictureBox9.Image");
            PictureBox9.Location = new System.Drawing.Point(194, 127);
            PictureBox9.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            PictureBox9.Name = "PictureBox9";
            PictureBox9.Size = new System.Drawing.Size(20, 20);
            PictureBox9.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            PictureBox9.TabIndex = 213;
            PictureBox9.TabStop = false;
            // 
            // PictureBox10
            // 
            PictureBox10.Image = (System.Drawing.Image)resources.GetObject("PictureBox10.Image");
            PictureBox10.Location = new System.Drawing.Point(194, 103);
            PictureBox10.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            PictureBox10.Name = "PictureBox10";
            PictureBox10.Size = new System.Drawing.Size(20, 20);
            PictureBox10.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            PictureBox10.TabIndex = 214;
            PictureBox10.TabStop = false;
            // 
            // PictureBox6
            // 
            PictureBox6.Image = (System.Drawing.Image)resources.GetObject("PictureBox6.Image");
            PictureBox6.Location = new System.Drawing.Point(194, 31);
            PictureBox6.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            PictureBox6.Name = "PictureBox6";
            PictureBox6.Size = new System.Drawing.Size(20, 20);
            PictureBox6.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            PictureBox6.TabIndex = 215;
            PictureBox6.TabStop = false;
            // 
            // PictureBox4
            // 
            PictureBox4.Image = (System.Drawing.Image)resources.GetObject("PictureBox4.Image");
            PictureBox4.Location = new System.Drawing.Point(194, 55);
            PictureBox4.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            PictureBox4.Name = "PictureBox4";
            PictureBox4.Size = new System.Drawing.Size(20, 20);
            PictureBox4.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            PictureBox4.TabIndex = 214;
            PictureBox4.TabStop = false;
            // 
            // PictureBox3
            // 
            PictureBox3.Image = (System.Drawing.Image)resources.GetObject("PictureBox3.Image");
            PictureBox3.Location = new System.Drawing.Point(194, 79);
            PictureBox3.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            PictureBox3.Name = "PictureBox3";
            PictureBox3.Size = new System.Drawing.Size(20, 20);
            PictureBox3.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            PictureBox3.TabIndex = 213;
            PictureBox3.TabStop = false;
            // 
            // PictureBox5
            // 
            PictureBox5.BackColor = System.Drawing.Color.White;
            PictureBox5.Image = (System.Drawing.Image)resources.GetObject("PictureBox5.Image");
            PictureBox5.Location = new System.Drawing.Point(194, 7);
            PictureBox5.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            PictureBox5.Name = "PictureBox5";
            PictureBox5.Size = new System.Drawing.Size(20, 20);
            PictureBox5.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            PictureBox5.TabIndex = 212;
            PictureBox5.TabStop = false;
            // 
            // CB_Name_Suffix
            // 
            CB_Name_Suffix.AutoSize = true;
            CB_Name_Suffix.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            CB_Name_Suffix.Location = new System.Drawing.Point(9, 30);
            CB_Name_Suffix.Name = "CB_Name_Suffix";
            CB_Name_Suffix.Size = new System.Drawing.Size(51, 21);
            CB_Name_Suffix.TabIndex = 6;
            CB_Name_Suffix.Text = "Title";
            CB_Name_Suffix.UseVisualStyleBackColor = true;
            // 
            // CB_Title
            // 
            CB_Title.AutoSize = true;
            CB_Title.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            CB_Title.Location = new System.Drawing.Point(9, 54);
            CB_Title.Name = "CB_Title";
            CB_Title.Size = new System.Drawing.Size(90, 21);
            CB_Title.TabIndex = 5;
            CB_Title.Text = "First Name";
            CB_Title.UseVisualStyleBackColor = true;
            // 
            // CB_Name_Abbreviations
            // 
            CB_Name_Abbreviations.AutoSize = true;
            CB_Name_Abbreviations.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            CB_Name_Abbreviations.Location = new System.Drawing.Point(9, 78);
            CB_Name_Abbreviations.Name = "CB_Name_Abbreviations";
            CB_Name_Abbreviations.Size = new System.Drawing.Size(107, 21);
            CB_Name_Abbreviations.TabIndex = 4;
            CB_Name_Abbreviations.Text = "Middle Name";
            CB_Name_Abbreviations.UseVisualStyleBackColor = true;
            // 
            // CB_Last_Name
            // 
            CB_Last_Name.AutoSize = true;
            CB_Last_Name.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            CB_Last_Name.Location = new System.Drawing.Point(9, 102);
            CB_Last_Name.Name = "CB_Last_Name";
            CB_Last_Name.Size = new System.Drawing.Size(125, 21);
            CB_Last_Name.TabIndex = 3;
            CB_Last_Name.Text = "Last Name Prefix";
            CB_Last_Name.UseVisualStyleBackColor = true;
            // 
            // CB_Last_Name_Prefix
            // 
            CB_Last_Name_Prefix.AutoSize = true;
            CB_Last_Name_Prefix.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            CB_Last_Name_Prefix.Location = new System.Drawing.Point(9, 126);
            CB_Last_Name_Prefix.Name = "CB_Last_Name_Prefix";
            CB_Last_Name_Prefix.Size = new System.Drawing.Size(89, 21);
            CB_Last_Name_Prefix.TabIndex = 2;
            CB_Last_Name_Prefix.Text = "Last Name";
            CB_Last_Name_Prefix.UseVisualStyleBackColor = true;
            // 
            // CB_Middle_Name
            // 
            CB_Middle_Name.AutoSize = true;
            CB_Middle_Name.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            CB_Middle_Name.Location = new System.Drawing.Point(9, 150);
            CB_Middle_Name.Name = "CB_Middle_Name";
            CB_Middle_Name.Size = new System.Drawing.Size(97, 21);
            CB_Middle_Name.TabIndex = 1;
            CB_Middle_Name.Text = "Name Suffix";
            CB_Middle_Name.UseVisualStyleBackColor = true;
            // 
            // CB_First_Name
            // 
            CB_First_Name.AutoSize = true;
            CB_First_Name.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            CB_First_Name.Location = new System.Drawing.Point(9, 174);
            CB_First_Name.Name = "CB_First_Name";
            CB_First_Name.Size = new System.Drawing.Size(146, 21);
            CB_First_Name.TabIndex = 0;
            CB_First_Name.Text = "Name Abbreviations";
            CB_First_Name.UseVisualStyleBackColor = true;
            // 
            // Form17DivideNames
            // 
            AutoScaleDimensions = new System.Drawing.SizeF(6.0f, 13.0f);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            ClientSize = new System.Drawing.Size(559, 549);
            Controls.Add(CB_Add_Header);
            Controls.Add(PictureBox7);
            Controls.Add(CB_Keep_Formatting);
            Controls.Add(CustomGroupBox4);
            Controls.Add(Selection);
            Controls.Add(AutoSelection);
            Controls.Add(txtSourceRange);
            Controls.Add(Label1);
            Controls.Add(CustomGroupBox5);
            Controls.Add(CustomGroupBox6);
            Controls.Add(btnOK);
            Controls.Add(btnCancel);
            Controls.Add(ComboBox1);
            Controls.Add(CustomGroupBox1);
            Controls.Add(CB_Backup_Sheet);
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "Form17DivideNames";
            Text = "Divide Names";
            ((System.ComponentModel.ISupportInitialize)Selection).EndInit();
            ((System.ComponentModel.ISupportInitialize)AutoSelection).EndInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox7).EndInit();
            CustomGroupBox4.ResumeLayout(false);
            CustomGroupBox10.ResumeLayout(false);
            CustomGroupBox10.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)PictureBox2).EndInit();
            ((System.ComponentModel.ISupportInitialize)destinationSelection).EndInit();
            CustomGroupBox5.ResumeLayout(false);
            CustomGroupBox6.ResumeLayout(false);
            CustomGroupBox1.ResumeLayout(false);
            CustomGroupBox7.ResumeLayout(false);
            CustomGroupBox7.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)PictureBox11).EndInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox8).EndInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox9).EndInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox10).EndInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox6).EndInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox4).EndInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox3).EndInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox5).EndInit();
            KeyDown += new System.Windows.Forms.KeyEventHandler(Form1_KeyDown);
            Load += new EventHandler(Form17DivideNames_Load);
            Activated += new EventHandler(Form1_Activated);
            Closing += new System.ComponentModel.CancelEventHandler(Form17DivideNames_Closing);
            Disposed += new EventHandler(Form17DivideNames_Disposed);
            Shown += new EventHandler(Form17DivideNames_Shown);
            ResumeLayout(false);
            PerformLayout();

        }

        internal System.Windows.Forms.PictureBox PictureBox8;
        internal System.Windows.Forms.PictureBox PictureBox9;
        internal System.Windows.Forms.PictureBox PictureBox10;
        internal System.Windows.Forms.PictureBox PictureBox6;
        internal System.Windows.Forms.PictureBox PictureBox3;
        internal System.Windows.Forms.PictureBox Selection;
        internal System.Windows.Forms.PictureBox AutoSelection;
        internal System.Windows.Forms.TextBox txtSourceRange;
        internal System.Windows.Forms.Label Label1;
        internal System.Windows.Forms.PictureBox PictureBox4;
        internal System.Windows.Forms.PictureBox PictureBox5;
        internal CustomGroupBox CustomGroupBox5;
        internal CustomPanel CustomPanel1;
        internal CustomGroupBox CustomGroupBox6;
        internal CustomPanel CustomPanel2;
        internal System.Windows.Forms.PictureBox PictureBox7;
        internal System.Windows.Forms.Button btnOK;
        internal System.Windows.Forms.Button btnCancel;
        internal CustomGroupBox CustomGroupBox7;
        internal System.Windows.Forms.CheckBox CB_Name_Suffix;
        internal System.Windows.Forms.CheckBox CB_Title;
        internal System.Windows.Forms.CheckBox CB_Name_Abbreviations;
        internal System.Windows.Forms.CheckBox CB_Last_Name;
        internal System.Windows.Forms.CheckBox CB_Last_Name_Prefix;
        internal System.Windows.Forms.CheckBox CB_Middle_Name;
        internal System.Windows.Forms.CheckBox CB_First_Name;
        internal System.Windows.Forms.ComboBox ComboBox1;
        internal CustomGroupBox CustomGroupBox1;
        internal System.Windows.Forms.CheckBox CB_Backup_Sheet;
        internal CustomGroupBox CustomGroupBox4;
        internal CustomGroupBox CustomGroupBox10;
        internal System.Windows.Forms.PictureBox PictureBox2;
        internal System.Windows.Forms.PictureBox destinationSelection;
        internal System.Windows.Forms.TextBox txtDestRange;
        internal System.Windows.Forms.Label lbl_destRange_Selection;
        internal System.Windows.Forms.RadioButton RB_Different_Range;
        internal System.Windows.Forms.RadioButton RB_Same_As_Source_Range;
        internal System.Windows.Forms.CheckBox CB_Keep_Formatting;
        internal System.Windows.Forms.PictureBox PictureBox11;
        internal System.Windows.Forms.CheckBox CB_Select_All;
        internal System.Windows.Forms.CheckBox CB_Add_Header;
    }
}