using System;

namespace VSTO_Addins
{
    [Microsoft.VisualBasic.CompilerServices.DesignerGenerated()]
    public partial class Form30_Create_Dynamic_Drop_down_List : System.Windows.Forms.Form
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
            var resources = new System.ComponentModel.ComponentResourceManager(typeof(Form30_Create_Dynamic_Drop_down_List));
            Btn_OK = new System.Windows.Forms.Button();
            Btn_OK.Click += new EventHandler(Btn_OK_Click);
            Btn_Cancel = new System.Windows.Forms.Button();
            Btn_Cancel.Click += new EventHandler(Btn_Cancel_Click);
            ComboBox2 = new System.Windows.Forms.ComboBox();
            Label1 = new System.Windows.Forms.Label();
            Selection_source = new System.Windows.Forms.PictureBox();
            Selection_source.Click += new EventHandler(Selection_source_Click);
            TB_src_range = new System.Windows.Forms.TextBox();
            TB_src_range.TextChanged += new EventHandler(TB_src_range_TextChanged);
            CB_header = new System.Windows.Forms.CheckBox();
            Label2 = new System.Windows.Forms.Label();
            Selection_destination = new System.Windows.Forms.PictureBox();
            Selection_destination.Click += new EventHandler(Selection_destination_Click);
            TB_dest_range = new System.Windows.Forms.TextBox();
            TB_dest_range.TextChanged += new EventHandler(TB_dest_range_TextChanged);
            CB_ascending = new System.Windows.Forms.CheckBox();
            CB_ascending.CheckedChanged += new EventHandler(CB_ascending_CheckedChanged);
            CB_descending = new System.Windows.Forms.CheckBox();
            CB_descending.CheckedChanged += new EventHandler(CB_descending_CheckedChanged);
            CB_text = new System.Windows.Forms.CheckBox();
            PictureBox2 = new System.Windows.Forms.PictureBox();
            PictureBox3 = new System.Windows.Forms.PictureBox();
            RB_vertical = new System.Windows.Forms.RadioButton();
            RB_Horizontal = new System.Windows.Forms.RadioButton();
            PictureBox8 = new System.Windows.Forms.PictureBox();
            PictureBox1 = new System.Windows.Forms.PictureBox();
            RB_2_5_levels = new System.Windows.Forms.RadioButton();
            RB_2_levels = new System.Windows.Forms.RadioButton();
            CustomGroupBox1 = new CustomGroupBox();
            CustomGroupBox3 = new CustomGroupBox();
            PictureBox4 = new System.Windows.Forms.PictureBox();
            RB_Dropdown_35_Labels = new System.Windows.Forms.RadioButton();
            RB_Dropdown_35_Labels.CheckedChanged += new EventHandler(RB_columns_CheckedChanged);
            PictureBox5 = new System.Windows.Forms.PictureBox();
            RB_Dropdown_2_Labels = new System.Windows.Forms.RadioButton();
            RB_Dropdown_2_Labels.CheckedChanged += new EventHandler(RB_rows_CheckedChanged);
            CustomGroupBox2 = new CustomGroupBox();
            GB_list_option = new CustomGroupBox();
            CustomGroupBox5 = new CustomGroupBox();
            PictureBox6 = new System.Windows.Forms.PictureBox();
            RB_Verti = new System.Windows.Forms.RadioButton();
            PictureBox7 = new System.Windows.Forms.PictureBox();
            RB_Horizon = new System.Windows.Forms.RadioButton();
            CustomGroupBox4 = new CustomGroupBox();
            CustomGroupBox7 = new CustomGroupBox();
            Label_ext = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)Selection_source).BeginInit();
            ((System.ComponentModel.ISupportInitialize)Selection_destination).BeginInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox2).BeginInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox3).BeginInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox8).BeginInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox1).BeginInit();
            CustomGroupBox1.SuspendLayout();
            CustomGroupBox3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)PictureBox4).BeginInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox5).BeginInit();
            GB_list_option.SuspendLayout();
            CustomGroupBox5.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)PictureBox6).BeginInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox7).BeginInit();
            SuspendLayout();
            // 
            // Btn_OK
            // 
            Btn_OK.BackColor = System.Drawing.Color.White;
            Btn_OK.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            Btn_OK.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            Btn_OK.Location = new System.Drawing.Point(519, 436);
            Btn_OK.Name = "Btn_OK";
            Btn_OK.Size = new System.Drawing.Size(62, 26);
            Btn_OK.TabIndex = 366;
            Btn_OK.Text = "OK";
            Btn_OK.UseVisualStyleBackColor = false;
            // 
            // Btn_Cancel
            // 
            Btn_Cancel.BackColor = System.Drawing.Color.White;
            Btn_Cancel.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            Btn_Cancel.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            Btn_Cancel.Location = new System.Drawing.Point(597, 436);
            Btn_Cancel.Name = "Btn_Cancel";
            Btn_Cancel.Size = new System.Drawing.Size(62, 26);
            Btn_Cancel.TabIndex = 365;
            Btn_Cancel.Text = "Cancel";
            Btn_Cancel.UseVisualStyleBackColor = false;
            // 
            // ComboBox2
            // 
            ComboBox2.Font = new System.Drawing.Font("Segoe UI", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            ComboBox2.FormattingEnabled = true;
            ComboBox2.Location = new System.Drawing.Point(15, 436);
            ComboBox2.Name = "ComboBox2";
            ComboBox2.Size = new System.Drawing.Size(154, 25);
            ComboBox2.TabIndex = 362;
            ComboBox2.Text = "Softeko";
            // 
            // Label1
            // 
            Label1.AutoSize = true;
            Label1.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            Label1.Location = new System.Drawing.Point(15, 15);
            Label1.Name = "Label1";
            Label1.Size = new System.Drawing.Size(98, 17);
            Label1.TabIndex = 355;
            Label1.Text = "Source Range :";
            // 
            // Selection_source
            // 
            Selection_source.BackColor = System.Drawing.Color.White;
            Selection_source.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            Selection_source.Image = (System.Drawing.Image)resources.GetObject("Selection_source.Image");
            Selection_source.Location = new System.Drawing.Point(298, 40);
            Selection_source.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            Selection_source.Name = "Selection_source";
            Selection_source.Size = new System.Drawing.Size(24, 25);
            Selection_source.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            Selection_source.TabIndex = 357;
            Selection_source.TabStop = false;
            // 
            // TB_src_range
            // 
            TB_src_range.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            TB_src_range.Cursor = System.Windows.Forms.Cursors.IBeam;
            TB_src_range.Font = new System.Drawing.Font("Segoe UI", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            TB_src_range.Location = new System.Drawing.Point(15, 40);
            TB_src_range.Name = "TB_src_range";
            TB_src_range.Size = new System.Drawing.Size(307, 25);
            TB_src_range.TabIndex = 356;
            // 
            // CB_header
            // 
            CB_header.AutoSize = true;
            CB_header.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            CB_header.Location = new System.Drawing.Point(15, 74);
            CB_header.Name = "CB_header";
            CB_header.Size = new System.Drawing.Size(180, 21);
            CB_header.TabIndex = 367;
            CB_header.Text = "My range contains header";
            CB_header.UseVisualStyleBackColor = true;
            // 
            // Label2
            // 
            Label2.AutoSize = true;
            Label2.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            Label2.Location = new System.Drawing.Point(15, 286);
            Label2.Name = "Label2";
            Label2.Size = new System.Drawing.Size(126, 17);
            Label2.TabIndex = 370;
            Label2.Text = "Destination Range :";
            // 
            // Selection_destination
            // 
            Selection_destination.BackColor = System.Drawing.Color.White;
            Selection_destination.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            Selection_destination.Image = (System.Drawing.Image)resources.GetObject("Selection_destination.Image");
            Selection_destination.Location = new System.Drawing.Point(298, 312);
            Selection_destination.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            Selection_destination.Name = "Selection_destination";
            Selection_destination.Size = new System.Drawing.Size(24, 25);
            Selection_destination.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            Selection_destination.TabIndex = 372;
            Selection_destination.TabStop = false;
            // 
            // TB_dest_range
            // 
            TB_dest_range.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            TB_dest_range.Cursor = System.Windows.Forms.Cursors.IBeam;
            TB_dest_range.Font = new System.Drawing.Font("Segoe UI", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            TB_dest_range.Location = new System.Drawing.Point(15, 312);
            TB_dest_range.Name = "TB_dest_range";
            TB_dest_range.Size = new System.Drawing.Size(307, 25);
            TB_dest_range.TabIndex = 371;
            // 
            // CB_ascending
            // 
            CB_ascending.AutoSize = true;
            CB_ascending.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            CB_ascending.Location = new System.Drawing.Point(15, 345);
            CB_ascending.Name = "CB_ascending";
            CB_ascending.Size = new System.Drawing.Size(165, 21);
            CB_ascending.TabIndex = 373;
            CB_ascending.Text = "Sort in ascending order";
            CB_ascending.UseVisualStyleBackColor = true;
            // 
            // CB_descending
            // 
            CB_descending.AutoSize = true;
            CB_descending.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            CB_descending.Location = new System.Drawing.Point(15, 372);
            CB_descending.Name = "CB_descending";
            CB_descending.Size = new System.Drawing.Size(173, 21);
            CB_descending.TabIndex = 374;
            CB_descending.Text = "Sort in descending order";
            CB_descending.UseVisualStyleBackColor = true;
            // 
            // CB_text
            // 
            CB_text.AutoSize = true;
            CB_text.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            CB_text.Location = new System.Drawing.Point(15, 400);
            CB_text.Name = "CB_text";
            CB_text.Size = new System.Drawing.Size(147, 21);
            CB_text.TabIndex = 375;
            CB_text.Text = "Store all data as text";
            CB_text.UseVisualStyleBackColor = true;
            // 
            // PictureBox2
            // 
            PictureBox2.Image = (System.Drawing.Image)resources.GetObject("PictureBox2.Image");
            PictureBox2.Location = new System.Drawing.Point(277, 33);
            PictureBox2.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            PictureBox2.Name = "PictureBox2";
            PictureBox2.Size = new System.Drawing.Size(20, 20);
            PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            PictureBox2.TabIndex = 275;
            PictureBox2.TabStop = false;
            // 
            // PictureBox3
            // 
            PictureBox3.Image = (System.Drawing.Image)resources.GetObject("PictureBox3.Image");
            PictureBox3.Location = new System.Drawing.Point(277, 7);
            PictureBox3.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            PictureBox3.Name = "PictureBox3";
            PictureBox3.Size = new System.Drawing.Size(20, 20);
            PictureBox3.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            PictureBox3.TabIndex = 274;
            PictureBox3.TabStop = false;
            // 
            // RB_vertical
            // 
            RB_vertical.AutoSize = true;
            RB_vertical.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            RB_vertical.Location = new System.Drawing.Point(8, 32);
            RB_vertical.Name = "RB_vertical";
            RB_vertical.Size = new System.Drawing.Size(158, 21);
            RB_vertical.TabIndex = 1;
            RB_vertical.Text = "Vertical drop-down list";
            RB_vertical.UseVisualStyleBackColor = true;
            // 
            // RB_Horizontal
            // 
            RB_Horizontal.AutoSize = true;
            RB_Horizontal.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            RB_Horizontal.Location = new System.Drawing.Point(8, 6);
            RB_Horizontal.Name = "RB_Horizontal";
            RB_Horizontal.Size = new System.Drawing.Size(176, 21);
            RB_Horizontal.TabIndex = 0;
            RB_Horizontal.Text = "Horizontal drop-down list";
            RB_Horizontal.UseVisualStyleBackColor = true;
            RB_Horizontal.Visible = false;
            // 
            // PictureBox8
            // 
            PictureBox8.Image = (System.Drawing.Image)resources.GetObject("PictureBox8.Image");
            PictureBox8.Location = new System.Drawing.Point(277, 33);
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
            PictureBox1.Location = new System.Drawing.Point(277, 7);
            PictureBox1.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            PictureBox1.Name = "PictureBox1";
            PictureBox1.Size = new System.Drawing.Size(20, 20);
            PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            PictureBox1.TabIndex = 274;
            PictureBox1.TabStop = false;
            // 
            // RB_2_5_levels
            // 
            RB_2_5_levels.AutoSize = true;
            RB_2_5_levels.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            RB_2_5_levels.Location = new System.Drawing.Point(8, 32);
            RB_2_5_levels.Name = "RB_2_5_levels";
            RB_2_5_levels.Size = new System.Drawing.Size(266, 21);
            RB_2_5_levels.TabIndex = 1;
            RB_2_5_levels.Text = "Dynamic drop-down list with 2 to 5 levels";
            RB_2_5_levels.UseVisualStyleBackColor = true;
            // 
            // RB_2_levels
            // 
            RB_2_levels.AutoSize = true;
            RB_2_levels.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            RB_2_levels.Location = new System.Drawing.Point(8, 6);
            RB_2_levels.Name = "RB_2_levels";
            RB_2_levels.Size = new System.Drawing.Size(239, 21);
            RB_2_levels.TabIndex = 0;
            RB_2_levels.Text = "Dynamic drop-down list with 2 levels";
            RB_2_levels.UseVisualStyleBackColor = true;
            // 
            // CustomGroupBox1
            // 
            CustomGroupBox1.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            CustomGroupBox1.Controls.Add(CustomGroupBox3);
            CustomGroupBox1.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 0);
            CustomGroupBox1.Location = new System.Drawing.Point(15, 102);
            CustomGroupBox1.Name = "CustomGroupBox1";
            CustomGroupBox1.Size = new System.Drawing.Size(307, 84);
            CustomGroupBox1.TabIndex = 368;
            CustomGroupBox1.TabStop = false;
            CustomGroupBox1.Text = "List Type";
            // 
            // CustomGroupBox3
            // 
            CustomGroupBox3.BackColor = System.Drawing.Color.White;
            CustomGroupBox3.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            CustomGroupBox3.Controls.Add(PictureBox4);
            CustomGroupBox3.Controls.Add(RB_Dropdown_35_Labels);
            CustomGroupBox3.Controls.Add(PictureBox5);
            CustomGroupBox3.Controls.Add(RB_Dropdown_2_Labels);
            CustomGroupBox3.Location = new System.Drawing.Point(1, 22);
            CustomGroupBox3.Name = "CustomGroupBox3";
            CustomGroupBox3.Size = new System.Drawing.Size(305, 62);
            CustomGroupBox3.TabIndex = 0;
            CustomGroupBox3.TabStop = false;
            // 
            // PictureBox4
            // 
            PictureBox4.Image = (System.Drawing.Image)resources.GetObject("PictureBox4.Image");
            PictureBox4.Location = new System.Drawing.Point(277, 34);
            PictureBox4.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            PictureBox4.Name = "PictureBox4";
            PictureBox4.Size = new System.Drawing.Size(20, 20);
            PictureBox4.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            PictureBox4.TabIndex = 377;
            PictureBox4.TabStop = false;
            // 
            // RB_Dropdown_35_Labels
            // 
            RB_Dropdown_35_Labels.AutoSize = true;
            RB_Dropdown_35_Labels.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            RB_Dropdown_35_Labels.Location = new System.Drawing.Point(8, 33);
            RB_Dropdown_35_Labels.Name = "RB_Dropdown_35_Labels";
            RB_Dropdown_35_Labels.Size = new System.Drawing.Size(266, 21);
            RB_Dropdown_35_Labels.TabIndex = 1;
            RB_Dropdown_35_Labels.Text = "Dynamic drop-down list with 2 to 5 levels";
            RB_Dropdown_35_Labels.UseVisualStyleBackColor = true;
            // 
            // PictureBox5
            // 
            PictureBox5.Image = (System.Drawing.Image)resources.GetObject("PictureBox5.Image");
            PictureBox5.Location = new System.Drawing.Point(277, 8);
            PictureBox5.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            PictureBox5.Name = "PictureBox5";
            PictureBox5.Size = new System.Drawing.Size(20, 20);
            PictureBox5.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            PictureBox5.TabIndex = 376;
            PictureBox5.TabStop = false;
            // 
            // RB_Dropdown_2_Labels
            // 
            RB_Dropdown_2_Labels.AutoSize = true;
            RB_Dropdown_2_Labels.Checked = true;
            RB_Dropdown_2_Labels.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            RB_Dropdown_2_Labels.Location = new System.Drawing.Point(8, 7);
            RB_Dropdown_2_Labels.Name = "RB_Dropdown_2_Labels";
            RB_Dropdown_2_Labels.Size = new System.Drawing.Size(239, 21);
            RB_Dropdown_2_Labels.TabIndex = 0;
            RB_Dropdown_2_Labels.TabStop = true;
            RB_Dropdown_2_Labels.Text = "Dynamic drop-down list with 2 levels";
            RB_Dropdown_2_Labels.UseVisualStyleBackColor = true;
            // 
            // CustomGroupBox2
            // 
            CustomGroupBox2.BackColor = System.Drawing.Color.White;
            CustomGroupBox2.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            CustomGroupBox2.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 0);
            CustomGroupBox2.Location = new System.Drawing.Point(352, 15);
            CustomGroupBox2.Name = "CustomGroupBox2";
            CustomGroupBox2.Size = new System.Drawing.Size(307, 395);
            CustomGroupBox2.TabIndex = 363;
            CustomGroupBox2.TabStop = false;
            CustomGroupBox2.Text = "Sample Image";
            // 
            // GB_list_option
            // 
            GB_list_option.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            GB_list_option.Controls.Add(CustomGroupBox5);
            GB_list_option.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 0);
            GB_list_option.Location = new System.Drawing.Point(15, 195);
            GB_list_option.Name = "GB_list_option";
            GB_list_option.Size = new System.Drawing.Size(307, 84);
            GB_list_option.TabIndex = 369;
            GB_list_option.TabStop = false;
            GB_list_option.Text = "List Options";
            // 
            // CustomGroupBox5
            // 
            CustomGroupBox5.BackColor = System.Drawing.Color.White;
            CustomGroupBox5.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            CustomGroupBox5.Controls.Add(PictureBox6);
            CustomGroupBox5.Controls.Add(RB_Verti);
            CustomGroupBox5.Controls.Add(PictureBox7);
            CustomGroupBox5.Controls.Add(RB_Horizon);
            CustomGroupBox5.Location = new System.Drawing.Point(1, 22);
            CustomGroupBox5.Name = "CustomGroupBox5";
            CustomGroupBox5.Size = new System.Drawing.Size(305, 62);
            CustomGroupBox5.TabIndex = 1;
            CustomGroupBox5.TabStop = false;
            // 
            // PictureBox6
            // 
            PictureBox6.Image = (System.Drawing.Image)resources.GetObject("PictureBox6.Image");
            PictureBox6.Location = new System.Drawing.Point(277, 34);
            PictureBox6.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            PictureBox6.Name = "PictureBox6";
            PictureBox6.Size = new System.Drawing.Size(20, 20);
            PictureBox6.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            PictureBox6.TabIndex = 377;
            PictureBox6.TabStop = false;
            // 
            // RB_Verti
            // 
            RB_Verti.AutoSize = true;
            RB_Verti.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            RB_Verti.Location = new System.Drawing.Point(8, 33);
            RB_Verti.Name = "RB_Verti";
            RB_Verti.Size = new System.Drawing.Size(158, 21);
            RB_Verti.TabIndex = 1;
            RB_Verti.Text = "Vertical drop-down list";
            RB_Verti.UseVisualStyleBackColor = true;
            // 
            // PictureBox7
            // 
            PictureBox7.Image = (System.Drawing.Image)resources.GetObject("PictureBox7.Image");
            PictureBox7.Location = new System.Drawing.Point(277, 8);
            PictureBox7.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            PictureBox7.Name = "PictureBox7";
            PictureBox7.Size = new System.Drawing.Size(20, 20);
            PictureBox7.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            PictureBox7.TabIndex = 376;
            PictureBox7.TabStop = false;
            // 
            // RB_Horizon
            // 
            RB_Horizon.AutoSize = true;
            RB_Horizon.Checked = true;
            RB_Horizon.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            RB_Horizon.Location = new System.Drawing.Point(8, 7);
            RB_Horizon.Name = "RB_Horizon";
            RB_Horizon.Size = new System.Drawing.Size(176, 21);
            RB_Horizon.TabIndex = 0;
            RB_Horizon.TabStop = true;
            RB_Horizon.Text = "Horizontal drop-down list";
            RB_Horizon.UseVisualStyleBackColor = true;
            // 
            // CustomGroupBox4
            // 
            CustomGroupBox4.BackColor = System.Drawing.Color.White;
            CustomGroupBox4.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            CustomGroupBox4.Location = new System.Drawing.Point(1, 22);
            CustomGroupBox4.Name = "CustomGroupBox4";
            CustomGroupBox4.Size = new System.Drawing.Size(306, 62);
            CustomGroupBox4.TabIndex = 0;
            CustomGroupBox4.TabStop = false;
            CustomGroupBox4.Text = "hhk";
            // 
            // CustomGroupBox7
            // 
            CustomGroupBox7.BackColor = System.Drawing.Color.Black;
            CustomGroupBox7.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            CustomGroupBox7.Location = new System.Drawing.Point(1, 22);
            CustomGroupBox7.Name = "CustomGroupBox7";
            CustomGroupBox7.Size = new System.Drawing.Size(306, 62);
            CustomGroupBox7.TabIndex = 0;
            CustomGroupBox7.TabStop = false;
            // 
            // Label_ext
            // 
            Label_ext.AutoSize = true;
            Label_ext.Location = new System.Drawing.Point(812, 132);
            Label_ext.Name = "Label_ext";
            Label_ext.Size = new System.Drawing.Size(0, 13);
            Label_ext.TabIndex = 376;
            // 
            // Form30_Create_Dynamic_Drop_down_List
            // 
            AutoScaleDimensions = new System.Drawing.SizeF(6.0f, 13.0f);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            ClientSize = new System.Drawing.Size(684, 486);
            Controls.Add(Label_ext);
            Controls.Add(CB_text);
            Controls.Add(CB_descending);
            Controls.Add(CB_ascending);
            Controls.Add(Label2);
            Controls.Add(Selection_destination);
            Controls.Add(TB_dest_range);
            Controls.Add(CustomGroupBox1);
            Controls.Add(CB_header);
            Controls.Add(Btn_OK);
            Controls.Add(Btn_Cancel);
            Controls.Add(CustomGroupBox2);
            Controls.Add(ComboBox2);
            Controls.Add(Label1);
            Controls.Add(Selection_source);
            Controls.Add(TB_src_range);
            Controls.Add(GB_list_option);
            KeyPreview = true;
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "Form30_Create_Dynamic_Drop_down_List";
            Text = "Create Dynamic Drop-down List";
            ((System.ComponentModel.ISupportInitialize)Selection_source).EndInit();
            ((System.ComponentModel.ISupportInitialize)Selection_destination).EndInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox2).EndInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox3).EndInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox8).EndInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox1).EndInit();
            CustomGroupBox1.ResumeLayout(false);
            CustomGroupBox3.ResumeLayout(false);
            CustomGroupBox3.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)PictureBox4).EndInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox5).EndInit();
            GB_list_option.ResumeLayout(false);
            CustomGroupBox5.ResumeLayout(false);
            CustomGroupBox5.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)PictureBox6).EndInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox7).EndInit();
            Load += new EventHandler(Form1_Load);
            KeyDown += new System.Windows.Forms.KeyEventHandler(form);
            Closing += new System.ComponentModel.CancelEventHandler(Form30_Create_Dynamic_Drop_down_List_Closing);
            Disposed += new EventHandler(Form30_Create_Dynamic_Drop_down_List_Disposed);
            Shown += new EventHandler(Form30_Create_Dynamic_Drop_down_List_Shown);
            ResumeLayout(false);
            PerformLayout();

        }
        internal System.Windows.Forms.Button Btn_OK;
        internal System.Windows.Forms.Button Btn_Cancel;
        internal CustomGroupBox CustomGroupBox2;
        internal System.Windows.Forms.ComboBox ComboBox2;
        internal System.Windows.Forms.Label Label1;
        internal System.Windows.Forms.PictureBox Selection_source;
        internal System.Windows.Forms.TextBox TB_src_range;
        internal System.Windows.Forms.CheckBox CB_header;
        internal CustomGroupBox CustomGroupBox1;
        internal CustomGroupBox CustomGroupBox7;
        internal System.Windows.Forms.PictureBox PictureBox8;
        internal System.Windows.Forms.PictureBox PictureBox1;
        internal System.Windows.Forms.RadioButton RB_2_5_levels;
        internal System.Windows.Forms.RadioButton RB_2_levels;
        internal CustomGroupBox GB_list_option;
        internal CustomGroupBox CustomGroupBox4;
        internal System.Windows.Forms.PictureBox PictureBox2;
        internal System.Windows.Forms.PictureBox PictureBox3;
        internal System.Windows.Forms.RadioButton RB_vertical;
        internal System.Windows.Forms.RadioButton RB_Horizontal;
        internal System.Windows.Forms.Label Label2;
        internal System.Windows.Forms.PictureBox Selection_destination;
        internal System.Windows.Forms.TextBox TB_dest_range;
        internal System.Windows.Forms.CheckBox CB_ascending;
        internal System.Windows.Forms.CheckBox CB_descending;
        internal System.Windows.Forms.CheckBox CB_text;
        internal CustomGroupBox CustomGroupBox3;
        internal CustomGroupBox CustomGroupBox5;
        internal System.Windows.Forms.RadioButton RB_Verti;
        internal System.Windows.Forms.RadioButton RB_Horizon;
        internal System.Windows.Forms.RadioButton RB_Dropdown_35_Labels;
        internal System.Windows.Forms.RadioButton RB_Dropdown_2_Labels;
        internal System.Windows.Forms.PictureBox PictureBox6;
        internal System.Windows.Forms.PictureBox PictureBox7;
        internal System.Windows.Forms.PictureBox PictureBox4;
        internal System.Windows.Forms.PictureBox PictureBox5;
        internal System.Windows.Forms.Label Label_ext;
    }
}