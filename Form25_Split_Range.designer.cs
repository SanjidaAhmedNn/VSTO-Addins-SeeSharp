using System;

namespace VSTO_Addins
{
    [Microsoft.VisualBasic.CompilerServices.DesignerGenerated()]
    public partial class Form25_Split_Range : System.Windows.Forms.Form
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
            var resources = new System.ComponentModel.ComponentResourceManager(typeof(Form25_Split_Range));
            Selection = new System.Windows.Forms.PictureBox();
            Selection.Click += new EventHandler(Selection_Click);
            Selection.GotFocus += new EventHandler(Selection_GotFocus);
            CheckBox2 = new System.Windows.Forms.CheckBox();
            CheckBox2.KeyDown += new System.Windows.Forms.KeyEventHandler(CheckBox2_KeyDown);
            PictureBox7 = new System.Windows.Forms.PictureBox();
            PictureBox7.KeyDown += new System.Windows.Forms.KeyEventHandler(PictureBox7_KeyDown);
            Button2 = new System.Windows.Forms.Button();
            Button2.Click += new EventHandler(Button2_Click);
            Button2.MouseEnter += new EventHandler(Button2_MouseEnter);
            Button2.MouseLeave += new EventHandler(Button2_MouseLeave);
            Button2.KeyDown += new System.Windows.Forms.KeyEventHandler(Button2_KeyDown);
            Button1 = new System.Windows.Forms.Button();
            Button1.Click += new EventHandler(Button1_Click);
            Button1.MouseEnter += new EventHandler(Button1_MouseEnter);
            Button1.MouseLeave += new EventHandler(Button1_MouseLeave);
            Button1.KeyDown += new System.Windows.Forms.KeyEventHandler(Button1_KeyDown);
            CheckBox1 = new System.Windows.Forms.CheckBox();
            CheckBox1.CheckedChanged += new EventHandler(CheckBox1_CheckedChanged);
            CheckBox1.KeyDown += new System.Windows.Forms.KeyEventHandler(CheckBox1_KeyDown);
            ComboBox1 = new System.Windows.Forms.ComboBox();
            ComboBox1.KeyDown += new System.Windows.Forms.KeyEventHandler(ComboBox1_KeyDown);
            Label1 = new System.Windows.Forms.Label();
            Label1.KeyDown += new System.Windows.Forms.KeyEventHandler(Label1_KeyDown);
            ToolTip1 = new System.Windows.Forms.ToolTip(components);
            AutoSelection = new System.Windows.Forms.PictureBox();
            AutoSelection.Click += new EventHandler(AutoSelection_Click);
            AutoSelection.GotFocus += new EventHandler(AutoSelection_GotFocus);
            AutoSelection.KeyDown += new System.Windows.Forms.KeyEventHandler(AutoSelection_KeyDown);
            TextBox1 = new System.Windows.Forms.TextBox();
            TextBox1.TextChanged += new EventHandler(TextBox1_TextChanged);
            TextBox1.GotFocus += new EventHandler(TextBox1_GotFocus);
            TextBox1.KeyDown += new System.Windows.Forms.KeyEventHandler(TextBox1_KeyDown);
            ComboBox3 = new System.Windows.Forms.ComboBox();
            ComboBox3.SelectedIndexChanged += new EventHandler(ComboBox3_SelectedIndexChanged);
            ComboBox3.KeyDown += new System.Windows.Forms.KeyEventHandler(ComboBox3_KeyDown);
            Label2 = new System.Windows.Forms.Label();
            Label2.KeyDown += new System.Windows.Forms.KeyEventHandler(Label2_KeyDown);
            CustomGroupBox5 = new CustomGroupBox();
            CustomGroupBox5.KeyDown += new System.Windows.Forms.KeyEventHandler(CustomGroupBox5_KeyDown);
            CustomPanel1 = new CustomPanel();
            CustomPanel1.KeyDown += new System.Windows.Forms.KeyEventHandler(CustomPanel1_KeyDown);
            CustomGroupBox1 = new CustomGroupBox();
            CustomGroupBox1.KeyDown += new System.Windows.Forms.KeyEventHandler(CustomGroupBox1_KeyDown);
            CustomGroupBox7 = new CustomGroupBox();
            CustomGroupBox7.KeyDown += new System.Windows.Forms.KeyEventHandler(CustomGroupBox7_KeyDown);
            PictureBox8 = new System.Windows.Forms.PictureBox();
            PictureBox8.KeyDown += new System.Windows.Forms.KeyEventHandler(PictureBox8_KeyDown);
            PictureBox1 = new System.Windows.Forms.PictureBox();
            PictureBox1.KeyDown += new System.Windows.Forms.KeyEventHandler(PictureBox1_KeyDown);
            RadioButton2 = new System.Windows.Forms.RadioButton();
            RadioButton2.CheckedChanged += new EventHandler(RadioButton2_CheckedChanged);
            RadioButton2.KeyDown += new System.Windows.Forms.KeyEventHandler(RadioButton2_KeyDown);
            RadioButton1 = new System.Windows.Forms.RadioButton();
            RadioButton1.CheckedChanged += new EventHandler(RadioButton1_CheckedChanged);
            RadioButton1.KeyDown += new System.Windows.Forms.KeyEventHandler(RadioButton1_KeyDown);
            CustomGroupBox6 = new CustomGroupBox();
            CustomGroupBox6.KeyDown += new System.Windows.Forms.KeyEventHandler(CustomGroupBox6_KeyDown);
            CustomPanel2 = new CustomPanel();
            CustomPanel2.KeyDown += new System.Windows.Forms.KeyEventHandler(CustomPanel2_KeyDown);
            CustomGroupBox2 = new CustomGroupBox();
            CustomGroupBox2.KeyDown += new System.Windows.Forms.KeyEventHandler(CustomGroupBox2_KeyDown);
            CustomGroupBox10 = new CustomGroupBox();
            CustomGroupBox10.KeyDown += new System.Windows.Forms.KeyEventHandler(CustomGroupBox10_KeyDown);
            PictureBox2 = new System.Windows.Forms.PictureBox();
            PictureBox2.KeyDown += new System.Windows.Forms.KeyEventHandler(PictureBox2_KeyDown);
            PictureBox3 = new System.Windows.Forms.PictureBox();
            PictureBox3.Click += new EventHandler(PictureBox3_Click);
            PictureBox3.GotFocus += new EventHandler(PictureBox3_GotFocus);
            PictureBox3.KeyDown += new System.Windows.Forms.KeyEventHandler(PictureBox3_KeyDown);
            TextBox4 = new System.Windows.Forms.TextBox();
            TextBox4.TextChanged += new EventHandler(TextBox4_TextChanged);
            TextBox4.GotFocus += new EventHandler(TextBox4_GotFocus);
            TextBox4.KeyDown += new System.Windows.Forms.KeyEventHandler(TextBox4_KeyDown);
            Label3 = new System.Windows.Forms.Label();
            Label3.KeyDown += new System.Windows.Forms.KeyEventHandler(Label3_KeyDown);
            RadioButton4 = new System.Windows.Forms.RadioButton();
            RadioButton4.CheckedChanged += new EventHandler(RadioButton4_CheckedChanged);
            RadioButton4.KeyDown += new System.Windows.Forms.KeyEventHandler(RadioButton4_KeyDown);
            RadioButton5 = new System.Windows.Forms.RadioButton();
            RadioButton5.CheckedChanged += new EventHandler(RadioButton5_CheckedChanged);
            RadioButton5.KeyDown += new System.Windows.Forms.KeyEventHandler(RadioButton5_KeyDown);
            CustomGroupBox4 = new CustomGroupBox();
            CustomGroupBox4.KeyDown += new System.Windows.Forms.KeyEventHandler(CustomGroupBox4_KeyDown);
            CustomGroupBox8 = new CustomGroupBox();
            CustomGroupBox8.KeyDown += new System.Windows.Forms.KeyEventHandler(CustomGroupBox8_KeyDown);
            ComboBox2 = new System.Windows.Forms.ComboBox();
            ComboBox2.TextChanged += new EventHandler(ComboBox2_TextChanged);
            ComboBox2.KeyDown += new System.Windows.Forms.KeyEventHandler(ComboBox2_KeyDown);
            PictureBox10 = new System.Windows.Forms.PictureBox();
            PictureBox10.KeyDown += new System.Windows.Forms.KeyEventHandler(PictureBox10_KeyDown);
            PictureBox6 = new System.Windows.Forms.PictureBox();
            PictureBox6.KeyDown += new System.Windows.Forms.KeyEventHandler(PictureBox6_KeyDown);
            PictureBox5 = new System.Windows.Forms.PictureBox();
            PictureBox5.KeyDown += new System.Windows.Forms.KeyEventHandler(PictureBox5_KeyDown);
            PictureBox4 = new System.Windows.Forms.PictureBox();
            PictureBox4.KeyDown += new System.Windows.Forms.KeyEventHandler(PictureBox4_KeyDown);
            PictureBox11 = new System.Windows.Forms.PictureBox();
            PictureBox11.KeyDown += new System.Windows.Forms.KeyEventHandler(PictureBox11_KeyDown);
            VScrollBar1 = new System.Windows.Forms.VScrollBar();
            VScrollBar1.KeyDown += new System.Windows.Forms.KeyEventHandler(VScrollBar1_KeyDown);
            RadioButton10 = new System.Windows.Forms.RadioButton();
            RadioButton10.CheckedChanged += new EventHandler(RadioButton10_CheckedChanged);
            RadioButton10.KeyDown += new System.Windows.Forms.KeyEventHandler(RadioButton10_KeyDown);
            RadioButton7 = new System.Windows.Forms.RadioButton();
            RadioButton7.CheckedChanged += new EventHandler(RadioButton7_CheckedChanged);
            RadioButton7.KeyDown += new System.Windows.Forms.KeyEventHandler(RadioButton7_KeyDown);
            RadioButton3 = new System.Windows.Forms.RadioButton();
            RadioButton3.CheckedChanged += new EventHandler(RadioButton3_CheckedChanged);
            RadioButton3.KeyDown += new System.Windows.Forms.KeyEventHandler(RadioButton3_KeyDown);
            RadioButton8 = new System.Windows.Forms.RadioButton();
            RadioButton8.CheckedChanged += new EventHandler(RadioButton8_CheckedChanged);
            RadioButton8.KeyDown += new System.Windows.Forms.KeyEventHandler(RadioButton8_KeyDown);
            RadioButton9 = new System.Windows.Forms.RadioButton();
            RadioButton9.CheckedChanged += new EventHandler(RadioButton9_CheckedChanged);
            RadioButton9.KeyDown += new System.Windows.Forms.KeyEventHandler(RadioButton9_KeyDown);
            TextBox3 = new System.Windows.Forms.TextBox();
            TextBox3.TextChanged += new EventHandler(TextBox3_TextChanged);
            TextBox3.KeyDown += new System.Windows.Forms.KeyEventHandler(TextBox3_KeyDown);
            RadioButton11 = new System.Windows.Forms.RadioButton();
            RadioButton11.CheckedChanged += new EventHandler(RadioButton11_CheckedChanged);
            RadioButton11.KeyDown += new System.Windows.Forms.KeyEventHandler(RadioButton11_KeyDown);
            ((System.ComponentModel.ISupportInitialize)Selection).BeginInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox7).BeginInit();
            ((System.ComponentModel.ISupportInitialize)AutoSelection).BeginInit();
            CustomGroupBox5.SuspendLayout();
            CustomGroupBox1.SuspendLayout();
            CustomGroupBox7.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)PictureBox8).BeginInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox1).BeginInit();
            CustomGroupBox6.SuspendLayout();
            CustomGroupBox2.SuspendLayout();
            CustomGroupBox10.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)PictureBox2).BeginInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox3).BeginInit();
            CustomGroupBox4.SuspendLayout();
            CustomGroupBox8.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)PictureBox10).BeginInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox6).BeginInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox5).BeginInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox4).BeginInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox11).BeginInit();
            SuspendLayout();
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
            Selection.TabIndex = 286;
            Selection.TabStop = false;
            // 
            // CheckBox2
            // 
            CheckBox2.AutoSize = true;
            CheckBox2.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            CheckBox2.Location = new System.Drawing.Point(18, 568);
            CheckBox2.Name = "CheckBox2";
            CheckBox2.Size = new System.Drawing.Size(257, 21);
            CheckBox2.TabIndex = 277;
            CheckBox2.Text = "Create a copy of the original worksheet";
            CheckBox2.UseVisualStyleBackColor = true;
            // 
            // PictureBox7
            // 
            PictureBox7.Image = (System.Drawing.Image)resources.GetObject("PictureBox7.Image");
            PictureBox7.Location = new System.Drawing.Point(440, 240);
            PictureBox7.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            PictureBox7.Name = "PictureBox7";
            PictureBox7.Size = new System.Drawing.Size(65, 65);
            PictureBox7.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            PictureBox7.TabIndex = 283;
            PictureBox7.TabStop = false;
            // 
            // Button2
            // 
            Button2.BackColor = System.Drawing.Color.White;
            Button2.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            Button2.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            Button2.Location = new System.Drawing.Point(483, 567);
            Button2.Name = "Button2";
            Button2.Size = new System.Drawing.Size(62, 26);
            Button2.TabIndex = 282;
            Button2.Text = "OK";
            Button2.UseVisualStyleBackColor = false;
            // 
            // Button1
            // 
            Button1.BackColor = System.Drawing.Color.White;
            Button1.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            Button1.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            Button1.Location = new System.Drawing.Point(560, 567);
            Button1.Name = "Button1";
            Button1.Size = new System.Drawing.Size(62, 26);
            Button1.TabIndex = 281;
            Button1.Text = "Cancel";
            Button1.UseVisualStyleBackColor = false;
            // 
            // CheckBox1
            // 
            CheckBox1.AutoSize = true;
            CheckBox1.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            CheckBox1.Location = new System.Drawing.Point(18, 196);
            CheckBox1.Name = "CheckBox1";
            CheckBox1.Size = new System.Drawing.Size(122, 21);
            CheckBox1.TabIndex = 276;
            CheckBox1.Text = "Keep formatting";
            CheckBox1.UseVisualStyleBackColor = true;
            // 
            // ComboBox1
            // 
            ComboBox1.Font = new System.Drawing.Font("Segoe UI", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            ComboBox1.FormattingEnabled = true;
            ComboBox1.Location = new System.Drawing.Point(18, 597);
            ComboBox1.Name = "ComboBox1";
            ComboBox1.Size = new System.Drawing.Size(154, 25);
            ComboBox1.TabIndex = 278;
            ComboBox1.Text = "SOFTEKO";
            // 
            // Label1
            // 
            Label1.AutoSize = true;
            Label1.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 0);
            Label1.Location = new System.Drawing.Point(15, 17);
            Label1.Name = "Label1";
            Label1.Size = new System.Drawing.Size(98, 17);
            Label1.TabIndex = 274;
            Label1.Text = "Source Range :";
            // 
            // AutoSelection
            // 
            AutoSelection.BackColor = System.Drawing.Color.White;
            AutoSelection.Image = (System.Drawing.Image)resources.GetObject("AutoSelection.Image");
            AutoSelection.Location = new System.Drawing.Point(224, 43);
            AutoSelection.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            AutoSelection.Name = "AutoSelection";
            AutoSelection.Size = new System.Drawing.Size(24, 23);
            AutoSelection.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            AutoSelection.TabIndex = 285;
            AutoSelection.TabStop = false;
            // 
            // TextBox1
            // 
            TextBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            TextBox1.Cursor = System.Windows.Forms.Cursors.IBeam;
            TextBox1.Font = new System.Drawing.Font("Segoe UI", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            TextBox1.Location = new System.Drawing.Point(15, 42);
            TextBox1.Name = "TextBox1";
            TextBox1.Size = new System.Drawing.Size(262, 25);
            TextBox1.TabIndex = 284;
            // 
            // ComboBox3
            // 
            ComboBox3.Font = new System.Drawing.Font("Segoe UI", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            ComboBox3.FormattingEnabled = true;
            ComboBox3.Location = new System.Drawing.Point(95, 77);
            ComboBox3.Name = "ComboBox3";
            ComboBox3.Size = new System.Drawing.Size(183, 25);
            ComboBox3.TabIndex = 289;
            // 
            // Label2
            // 
            Label2.AutoSize = true;
            Label2.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            Label2.Location = new System.Drawing.Point(16, 78);
            Label2.Name = "Label2";
            Label2.Size = new System.Drawing.Size(56, 17);
            Label2.TabIndex = 290;
            Label2.Text = "Split by:";
            // 
            // CustomGroupBox5
            // 
            CustomGroupBox5.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            CustomGroupBox5.Controls.Add(CustomPanel1);
            CustomGroupBox5.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 0);
            CustomGroupBox5.Location = new System.Drawing.Point(320, 17);
            CustomGroupBox5.Name = "CustomGroupBox5";
            CustomGroupBox5.Size = new System.Drawing.Size(302, 200);
            CustomGroupBox5.TabIndex = 279;
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
            CustomPanel1.Size = new System.Drawing.Size(300, 170);
            CustomPanel1.TabIndex = 0;
            // 
            // CustomGroupBox1
            // 
            CustomGroupBox1.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            CustomGroupBox1.Controls.Add(CustomGroupBox7);
            CustomGroupBox1.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 0);
            CustomGroupBox1.Location = new System.Drawing.Point(17, 106);
            CustomGroupBox1.Name = "CustomGroupBox1";
            CustomGroupBox1.Size = new System.Drawing.Size(260, 84);
            CustomGroupBox1.TabIndex = 275;
            CustomGroupBox1.TabStop = false;
            CustomGroupBox1.Text = "Split Option";
            // 
            // CustomGroupBox7
            // 
            CustomGroupBox7.BackColor = System.Drawing.Color.White;
            CustomGroupBox7.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            CustomGroupBox7.Controls.Add(PictureBox8);
            CustomGroupBox7.Controls.Add(PictureBox1);
            CustomGroupBox7.Controls.Add(RadioButton2);
            CustomGroupBox7.Controls.Add(RadioButton1);
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
            // RadioButton2
            // 
            RadioButton2.AutoSize = true;
            RadioButton2.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            RadioButton2.Location = new System.Drawing.Point(8, 32);
            RadioButton2.Name = "RadioButton2";
            RadioButton2.Size = new System.Drawing.Size(167, 21);
            RadioButton2.TabIndex = 1;
            RadioButton2.Text = "Split range into columns";
            RadioButton2.UseVisualStyleBackColor = true;
            // 
            // RadioButton1
            // 
            RadioButton1.AutoSize = true;
            RadioButton1.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            RadioButton1.Location = new System.Drawing.Point(8, 6);
            RadioButton1.Name = "RadioButton1";
            RadioButton1.Size = new System.Drawing.Size(147, 21);
            RadioButton1.TabIndex = 0;
            RadioButton1.Text = "Split range into rows";
            RadioButton1.UseVisualStyleBackColor = true;
            // 
            // CustomGroupBox6
            // 
            CustomGroupBox6.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            CustomGroupBox6.Controls.Add(CustomPanel2);
            CustomGroupBox6.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 0);
            CustomGroupBox6.Location = new System.Drawing.Point(320, 322);
            CustomGroupBox6.Name = "CustomGroupBox6";
            CustomGroupBox6.Size = new System.Drawing.Size(302, 200);
            CustomGroupBox6.TabIndex = 280;
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
            CustomPanel2.Size = new System.Drawing.Size(300, 170);
            CustomPanel2.TabIndex = 11;
            // 
            // CustomGroupBox2
            // 
            CustomGroupBox2.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            CustomGroupBox2.Controls.Add(CustomGroupBox10);
            CustomGroupBox2.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 0);
            CustomGroupBox2.Location = new System.Drawing.Point(18, 424);
            CustomGroupBox2.Name = "CustomGroupBox2";
            CustomGroupBox2.Size = new System.Drawing.Size(260, 137);
            CustomGroupBox2.TabIndex = 288;
            CustomGroupBox2.TabStop = false;
            CustomGroupBox2.Text = "Destination Range";
            // 
            // CustomGroupBox10
            // 
            CustomGroupBox10.BackColor = System.Drawing.Color.White;
            CustomGroupBox10.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            CustomGroupBox10.Controls.Add(PictureBox2);
            CustomGroupBox10.Controls.Add(PictureBox3);
            CustomGroupBox10.Controls.Add(TextBox4);
            CustomGroupBox10.Controls.Add(Label3);
            CustomGroupBox10.Controls.Add(RadioButton4);
            CustomGroupBox10.Controls.Add(RadioButton5);
            CustomGroupBox10.Location = new System.Drawing.Point(1, 22);
            CustomGroupBox10.Name = "CustomGroupBox10";
            CustomGroupBox10.Size = new System.Drawing.Size(259, 115);
            CustomGroupBox10.TabIndex = 0;
            CustomGroupBox10.TabStop = false;
            // 
            // PictureBox2
            // 
            PictureBox2.Enabled = false;
            PictureBox2.Image = (System.Drawing.Image)resources.GetObject("PictureBox2.Image");
            PictureBox2.Location = new System.Drawing.Point(25, 58);
            PictureBox2.Name = "PictureBox2";
            PictureBox2.Size = new System.Drawing.Size(14, 14);
            PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            PictureBox2.TabIndex = 208;
            PictureBox2.TabStop = false;
            // 
            // PictureBox3
            // 
            PictureBox3.BackColor = System.Drawing.Color.White;
            PictureBox3.Enabled = false;
            PictureBox3.Image = (System.Drawing.Image)resources.GetObject("PictureBox3.Image");
            PictureBox3.Location = new System.Drawing.Point(224, 81);
            PictureBox3.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            PictureBox3.Name = "PictureBox3";
            PictureBox3.Size = new System.Drawing.Size(24, 23);
            PictureBox3.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            PictureBox3.TabIndex = 207;
            PictureBox3.TabStop = false;
            // 
            // TextBox4
            // 
            TextBox4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            TextBox4.Cursor = System.Windows.Forms.Cursors.IBeam;
            TextBox4.Enabled = false;
            TextBox4.Font = new System.Drawing.Font("Segoe UI", 10.0f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            TextBox4.Location = new System.Drawing.Point(25, 80);
            TextBox4.Name = "TextBox4";
            TextBox4.Size = new System.Drawing.Size(224, 25);
            TextBox4.TabIndex = 206;
            // 
            // Label3
            // 
            Label3.AutoSize = true;
            Label3.Enabled = false;
            Label3.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            Label3.Location = new System.Drawing.Point(42, 56);
            Label3.Name = "Label3";
            Label3.Size = new System.Drawing.Size(109, 17);
            Label3.TabIndex = 2;
            Label3.Text = "Select the range :";
            // 
            // RadioButton4
            // 
            RadioButton4.AutoSize = true;
            RadioButton4.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            RadioButton4.Location = new System.Drawing.Point(8, 31);
            RadioButton4.Name = "RadioButton4";
            RadioButton4.Size = new System.Drawing.Size(185, 21);
            RadioButton4.TabIndex = 1;
            RadioButton4.TabStop = true;
            RadioButton4.Text = "Store into a different range";
            RadioButton4.UseVisualStyleBackColor = true;
            // 
            // RadioButton5
            // 
            RadioButton5.AutoSize = true;
            RadioButton5.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            RadioButton5.Location = new System.Drawing.Point(8, 6);
            RadioButton5.Name = "RadioButton5";
            RadioButton5.Size = new System.Drawing.Size(178, 21);
            RadioButton5.TabIndex = 0;
            RadioButton5.TabStop = true;
            RadioButton5.Text = "Same as the source range";
            RadioButton5.UseVisualStyleBackColor = true;
            // 
            // CustomGroupBox4
            // 
            CustomGroupBox4.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            CustomGroupBox4.Controls.Add(CustomGroupBox8);
            CustomGroupBox4.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 0);
            CustomGroupBox4.Location = new System.Drawing.Point(18, 223);
            CustomGroupBox4.Name = "CustomGroupBox4";
            CustomGroupBox4.Size = new System.Drawing.Size(260, 195);
            CustomGroupBox4.TabIndex = 287;
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
            CustomGroupBox8.Controls.Add(RadioButton10);
            CustomGroupBox8.Controls.Add(RadioButton7);
            CustomGroupBox8.Controls.Add(RadioButton3);
            CustomGroupBox8.Controls.Add(RadioButton8);
            CustomGroupBox8.Controls.Add(RadioButton9);
            CustomGroupBox8.Controls.Add(TextBox3);
            CustomGroupBox8.Controls.Add(RadioButton11);
            CustomGroupBox8.Location = new System.Drawing.Point(1, 22);
            CustomGroupBox8.Name = "CustomGroupBox8";
            CustomGroupBox8.Size = new System.Drawing.Size(259, 173);
            CustomGroupBox8.TabIndex = 0;
            CustomGroupBox8.TabStop = false;
            // 
            // ComboBox2
            // 
            ComboBox2.Enabled = false;
            ComboBox2.FormattingEnabled = true;
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
            PictureBox11.Enabled = false;
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
            // RadioButton10
            // 
            RadioButton10.AutoSize = true;
            RadioButton10.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            RadioButton10.Location = new System.Drawing.Point(8, 108);
            RadioButton10.Name = "RadioButton10";
            RadioButton10.Size = new System.Drawing.Size(72, 21);
            RadioButton10.TabIndex = 233;
            RadioButton10.TabStop = true;
            RadioButton10.Text = "Others :";
            RadioButton10.UseVisualStyleBackColor = true;
            // 
            // RadioButton7
            // 
            RadioButton7.AutoSize = true;
            RadioButton7.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            RadioButton7.Location = new System.Drawing.Point(8, 78);
            RadioButton7.Name = "RadioButton7";
            RadioButton7.Size = new System.Drawing.Size(86, 21);
            RadioButton7.TabIndex = 232;
            RadioButton7.TabStop = true;
            RadioButton7.Text = "Semicolon";
            RadioButton7.UseVisualStyleBackColor = true;
            // 
            // RadioButton3
            // 
            RadioButton3.AutoSize = true;
            RadioButton3.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            RadioButton3.Location = new System.Drawing.Point(8, 54);
            RadioButton3.Name = "RadioButton3";
            RadioButton3.Size = new System.Drawing.Size(125, 21);
            RadioButton3.TabIndex = 231;
            RadioButton3.TabStop = true;
            RadioButton3.Text = "Number and text";
            RadioButton3.UseVisualStyleBackColor = true;
            // 
            // RadioButton8
            // 
            RadioButton8.AutoSize = true;
            RadioButton8.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            RadioButton8.Location = new System.Drawing.Point(8, 30);
            RadioButton8.Name = "RadioButton8";
            RadioButton8.Size = new System.Drawing.Size(76, 21);
            RadioButton8.TabIndex = 1;
            RadioButton8.TabStop = true;
            RadioButton8.Text = "New line";
            RadioButton8.UseVisualStyleBackColor = true;
            // 
            // RadioButton9
            // 
            RadioButton9.AutoSize = true;
            RadioButton9.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            RadioButton9.Location = new System.Drawing.Point(8, 6);
            RadioButton9.Name = "RadioButton9";
            RadioButton9.Size = new System.Drawing.Size(61, 21);
            RadioButton9.TabIndex = 0;
            RadioButton9.TabStop = true;
            RadioButton9.Text = "Space";
            RadioButton9.UseVisualStyleBackColor = true;
            // 
            // TextBox3
            // 
            TextBox3.Enabled = false;
            TextBox3.Font = new System.Drawing.Font("Segoe UI", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            TextBox3.Location = new System.Drawing.Point(112, 140);
            TextBox3.Name = "TextBox3";
            TextBox3.Size = new System.Drawing.Size(100, 25);
            TextBox3.TabIndex = 237;
            // 
            // RadioButton11
            // 
            RadioButton11.AutoSize = true;
            RadioButton11.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            RadioButton11.Location = new System.Drawing.Point(8, 141);
            RadioButton11.Name = "RadioButton11";
            RadioButton11.Size = new System.Drawing.Size(105, 21);
            RadioButton11.TabIndex = 234;
            RadioButton11.TabStop = true;
            RadioButton11.Text = "Define width :";
            RadioButton11.UseVisualStyleBackColor = true;
            // 
            // Form25_Split_Range
            // 
            AutoScaleDimensions = new System.Drawing.SizeF(6.0f, 13.0f);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            ClientSize = new System.Drawing.Size(643, 640);
            Controls.Add(Label2);
            Controls.Add(ComboBox3);
            Controls.Add(Selection);
            Controls.Add(CustomGroupBox5);
            Controls.Add(CheckBox2);
            Controls.Add(CustomGroupBox1);
            Controls.Add(PictureBox7);
            Controls.Add(Button2);
            Controls.Add(Button1);
            Controls.Add(CheckBox1);
            Controls.Add(CustomGroupBox6);
            Controls.Add(CustomGroupBox2);
            Controls.Add(ComboBox1);
            Controls.Add(Label1);
            Controls.Add(AutoSelection);
            Controls.Add(CustomGroupBox4);
            Controls.Add(TextBox1);
            Icon = (System.Drawing.Icon)resources.GetObject("$this.Icon");
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "Form25_Split_Range";
            Text = "Split Range";
            ((System.ComponentModel.ISupportInitialize)Selection).EndInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox7).EndInit();
            ((System.ComponentModel.ISupportInitialize)AutoSelection).EndInit();
            CustomGroupBox5.ResumeLayout(false);
            CustomGroupBox1.ResumeLayout(false);
            CustomGroupBox7.ResumeLayout(false);
            CustomGroupBox7.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)PictureBox8).EndInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox1).EndInit();
            CustomGroupBox6.ResumeLayout(false);
            CustomGroupBox2.ResumeLayout(false);
            CustomGroupBox10.ResumeLayout(false);
            CustomGroupBox10.PerformLayout();
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
            Load += new EventHandler(Form25_Split_Range_Load);
            Closing += new System.ComponentModel.CancelEventHandler(Form25_Split_Range_Closing);
            Disposed += new EventHandler(Form25_Split_Range_Disposed);
            Shown += new EventHandler(Form25_Split_Range_Shown);
            ResumeLayout(false);
            PerformLayout();

        }

        internal System.Windows.Forms.PictureBox Selection;
        internal CustomGroupBox CustomGroupBox5;
        internal CustomPanel CustomPanel1;
        internal System.Windows.Forms.CheckBox CheckBox2;
        internal CustomGroupBox CustomGroupBox1;
        internal CustomGroupBox CustomGroupBox7;
        internal System.Windows.Forms.PictureBox PictureBox8;
        internal System.Windows.Forms.PictureBox PictureBox1;
        internal System.Windows.Forms.RadioButton RadioButton2;
        internal System.Windows.Forms.RadioButton RadioButton1;
        internal System.Windows.Forms.PictureBox PictureBox7;
        internal System.Windows.Forms.Button Button2;
        internal System.Windows.Forms.Button Button1;
        internal System.Windows.Forms.CheckBox CheckBox1;
        internal CustomGroupBox CustomGroupBox6;
        internal CustomPanel CustomPanel2;
        internal CustomGroupBox CustomGroupBox2;
        internal CustomGroupBox CustomGroupBox10;
        internal System.Windows.Forms.PictureBox PictureBox2;
        internal System.Windows.Forms.PictureBox PictureBox3;
        internal System.Windows.Forms.TextBox TextBox4;
        internal System.Windows.Forms.Label Label3;
        internal System.Windows.Forms.RadioButton RadioButton4;
        internal System.Windows.Forms.RadioButton RadioButton5;
        internal System.Windows.Forms.ComboBox ComboBox1;
        internal System.Windows.Forms.Label Label1;
        internal System.Windows.Forms.ToolTip ToolTip1;
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
        internal System.Windows.Forms.RadioButton RadioButton10;
        internal System.Windows.Forms.RadioButton RadioButton7;
        internal System.Windows.Forms.RadioButton RadioButton3;
        internal System.Windows.Forms.RadioButton RadioButton8;
        internal System.Windows.Forms.RadioButton RadioButton9;
        internal System.Windows.Forms.TextBox TextBox3;
        internal System.Windows.Forms.RadioButton RadioButton11;
        internal System.Windows.Forms.TextBox TextBox1;
        internal System.Windows.Forms.ComboBox ComboBox3;
        internal System.Windows.Forms.Label Label2;
    }
}