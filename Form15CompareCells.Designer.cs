using System;

namespace VSTO_Addins
{
    [Microsoft.VisualBasic.CompilerServices.DesignerGenerated()]
    public partial class Form15CompareCells : System.Windows.Forms.Form
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
            var resources = new System.ComponentModel.ComponentResourceManager(typeof(Form15CompareCells));
            checkBoxFormatting = new System.Windows.Forms.CheckBox();
            checkBoxFormatting.CheckedChanged += new EventHandler(checkBoxFormatting_CheckedChanged);
            btnOK = new System.Windows.Forms.Button();
            btnOK.Click += new EventHandler(btnOK_Click);
            btnCanecl = new System.Windows.Forms.Button();
            btnCanecl.Click += new EventHandler(btnCanecl_Click);
            ComboBox1 = new System.Windows.Forms.ComboBox();
            checkBoxCopyWs = new System.Windows.Forms.CheckBox();
            checkBoxCase = new System.Windows.Forms.CheckBox();
            checkBoxCase.CheckedChanged += new EventHandler(checkBoxCase_CheckedChanged);
            CD_Fill_Background = new System.Windows.Forms.ColorDialog();
            CD_Fill_Font = new System.Windows.Forms.ColorDialog();
            CustomPanel1 = new CustomPanel();
            CustomPanel1.Paint += new System.Windows.Forms.PaintEventHandler(CustomPanel1_Paint);
            CustomGroupBox5 = new CustomGroupBox();
            CP_Input_Range2 = new CustomPanel();
            GB_Input_Range = new CustomGroupBox();
            CP_Input_Range1 = new CustomPanel();
            GB_Expected_Output = new CustomGroupBox();
            CP_Output_Range = new CustomPanel();
            PictureBox7 = new System.Windows.Forms.PictureBox();
            GB_Display_Result = new CustomGroupBox();
            CustomGroupBox4 = new CustomGroupBox();
            CbFillFont = new System.Windows.Forms.ComboBox();
            CbFillFont.Click += new EventHandler(CbFillFont_Click);
            CbFillFont.BackColorChanged += new EventHandler(CbFillFont_BackColorChanged);
            CBFillBackground = new System.Windows.Forms.ComboBox();
            CBFillBackground.Click += new EventHandler(CBFillBackground_Click);
            CBFillBackground.SelectedIndexChanged += new EventHandler(CBFillBackground_SelectedIndexChanged);
            CBFillBackground.BackColorChanged += new EventHandler(CBFillBackground_BackColorChanged);
            checkBoxFillFont = new System.Windows.Forms.CheckBox();
            checkBoxFillFont.CheckedChanged += new EventHandler(checkBoxFillFont_CheckedChanged);
            checkBoxFillBack = new System.Windows.Forms.CheckBox();
            checkBoxFillBack.CheckedChanged += new EventHandler(checkBoxFillBack_CheckedChanged);
            CustomGroupBox1 = new CustomGroupBox();
            CustomGroupBox7 = new CustomGroupBox();
            radBtnSameValues = new System.Windows.Forms.RadioButton();
            radBtnSameValues.CheckedChanged += new EventHandler(radBtnSameValues_CheckedChanged);
            PictureBox1 = new System.Windows.Forms.PictureBox();
            PictureBox5 = new System.Windows.Forms.PictureBox();
            radBtnDifferentValues = new System.Windows.Forms.RadioButton();
            radBtnDifferentValues.CheckedChanged += new EventHandler(radBtnDifferentValues_CheckedChanged);
            CustomGroupBox2 = new CustomGroupBox();
            rngSelection2 = new System.Windows.Forms.PictureBox();
            rngSelection2.Click += new EventHandler(rngSelection2_Click);
            rngSelection1 = new System.Windows.Forms.PictureBox();
            rngSelection1.Click += new EventHandler(rngSelection1_Click);
            AutoSelection2 = new System.Windows.Forms.PictureBox();
            AutoSelection2.Click += new EventHandler(AutoSelection2_Click);
            AutoSelection1 = new System.Windows.Forms.PictureBox();
            AutoSelection1.Click += new EventHandler(AutoSelection1_Click);
            txtSourceRange2 = new System.Windows.Forms.TextBox();
            txtSourceRange2.TextChanged += new EventHandler(txtSourceRange2_TextChanged);
            txtSourceRange2.GotFocus += new EventHandler(txtSourceRange2_GotFocus);
            txtSourceRange2.Click += new EventHandler(txtSourceRange2_Click);
            txtSourceRange1 = new System.Windows.Forms.TextBox();
            txtSourceRange1.TextChanged += new EventHandler(txtSourceRange1_TextChanged);
            txtSourceRange1.GotFocus += new EventHandler(txtSourceRange1_GotFocus);
            txtSourceRange1.Click += new EventHandler(txtSourceRange1_Click);
            lblSourceRng2 = new System.Windows.Forms.Label();
            lblSourceRng1 = new System.Windows.Forms.Label();
            CustomPanel1.SuspendLayout();
            CustomGroupBox5.SuspendLayout();
            GB_Input_Range.SuspendLayout();
            GB_Expected_Output.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)PictureBox7).BeginInit();
            GB_Display_Result.SuspendLayout();
            CustomGroupBox4.SuspendLayout();
            CustomGroupBox1.SuspendLayout();
            CustomGroupBox7.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)PictureBox1).BeginInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox5).BeginInit();
            CustomGroupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)rngSelection2).BeginInit();
            ((System.ComponentModel.ISupportInitialize)rngSelection1).BeginInit();
            ((System.ComponentModel.ISupportInitialize)AutoSelection2).BeginInit();
            ((System.ComponentModel.ISupportInitialize)AutoSelection1).BeginInit();
            SuspendLayout();
            // 
            // checkBoxFormatting
            // 
            checkBoxFormatting.AutoSize = true;
            checkBoxFormatting.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            checkBoxFormatting.Location = new System.Drawing.Point(15, 259);
            checkBoxFormatting.Name = "checkBoxFormatting";
            checkBoxFormatting.Size = new System.Drawing.Size(122, 21);
            checkBoxFormatting.TabIndex = 166;
            checkBoxFormatting.Text = "Keep formatting";
            checkBoxFormatting.UseVisualStyleBackColor = true;
            // 
            // btnOK
            // 
            btnOK.BackColor = System.Drawing.Color.White;
            btnOK.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            btnOK.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            btnOK.Location = new System.Drawing.Point(611, 423);
            btnOK.Name = "btnOK";
            btnOK.Size = new System.Drawing.Size(62, 26);
            btnOK.TabIndex = 172;
            btnOK.Text = "OK";
            btnOK.UseVisualStyleBackColor = false;
            // 
            // btnCanecl
            // 
            btnCanecl.BackColor = System.Drawing.Color.White;
            btnCanecl.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            btnCanecl.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            btnCanecl.Location = new System.Drawing.Point(692, 423);
            btnCanecl.Name = "btnCanecl";
            btnCanecl.Size = new System.Drawing.Size(62, 26);
            btnCanecl.TabIndex = 171;
            btnCanecl.Text = "Cancel";
            btnCanecl.UseVisualStyleBackColor = false;
            // 
            // ComboBox1
            // 
            ComboBox1.Font = new System.Drawing.Font("Segoe UI", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            ComboBox1.FormattingEnabled = true;
            ComboBox1.Location = new System.Drawing.Point(15, 423);
            ComboBox1.Name = "ComboBox1";
            ComboBox1.Size = new System.Drawing.Size(154, 25);
            ComboBox1.TabIndex = 168;
            ComboBox1.Text = "SOFTEKO";
            // 
            // checkBoxCopyWs
            // 
            checkBoxCopyWs.AutoSize = true;
            checkBoxCopyWs.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            checkBoxCopyWs.Location = new System.Drawing.Point(15, 392);
            checkBoxCopyWs.Name = "checkBoxCopyWs";
            checkBoxCopyWs.Size = new System.Drawing.Size(257, 21);
            checkBoxCopyWs.TabIndex = 167;
            checkBoxCopyWs.Text = "Create a copy of the original worksheet";
            checkBoxCopyWs.UseVisualStyleBackColor = true;
            // 
            // checkBoxCase
            // 
            checkBoxCase.AutoSize = true;
            checkBoxCase.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            checkBoxCase.Location = new System.Drawing.Point(179, 259);
            checkBoxCase.Name = "checkBoxCase";
            checkBoxCase.Size = new System.Drawing.Size(108, 21);
            checkBoxCase.TabIndex = 176;
            checkBoxCase.Text = "Case sensitive";
            checkBoxCase.UseVisualStyleBackColor = true;
            // 
            // CustomPanel1
            // 
            CustomPanel1.BorderColor = System.Drawing.Color.Empty;
            CustomPanel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            CustomPanel1.BorderWidth = 0;
            CustomPanel1.Controls.Add(CustomGroupBox5);
            CustomPanel1.Controls.Add(GB_Input_Range);
            CustomPanel1.Controls.Add(GB_Expected_Output);
            CustomPanel1.Controls.Add(PictureBox7);
            CustomPanel1.Location = new System.Drawing.Point(322, 15);
            CustomPanel1.Name = "CustomPanel1";
            CustomPanel1.Size = new System.Drawing.Size(432, 390);
            CustomPanel1.TabIndex = 177;
            // 
            // CustomGroupBox5
            // 
            CustomGroupBox5.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            CustomGroupBox5.Controls.Add(CP_Input_Range2);
            CustomGroupBox5.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 0);
            CustomGroupBox5.Location = new System.Drawing.Point(221, 15);
            CustomGroupBox5.Name = "CustomGroupBox5";
            CustomGroupBox5.Size = new System.Drawing.Size(192, 142);
            CustomGroupBox5.TabIndex = 161;
            CustomGroupBox5.TabStop = false;
            CustomGroupBox5.Text = "2nd Input Range";
            // 
            // CP_Input_Range2
            // 
            CP_Input_Range2.BackColor = System.Drawing.Color.White;
            CP_Input_Range2.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            CP_Input_Range2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            CP_Input_Range2.BorderWidth = 1;
            CP_Input_Range2.Location = new System.Drawing.Point(1, 30);
            CP_Input_Range2.Name = "CP_Input_Range2";
            CP_Input_Range2.Size = new System.Drawing.Size(190, 112);
            CP_Input_Range2.TabIndex = 0;
            // 
            // GB_Input_Range
            // 
            GB_Input_Range.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            GB_Input_Range.Controls.Add(CP_Input_Range1);
            GB_Input_Range.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 0);
            GB_Input_Range.Location = new System.Drawing.Point(15, 15);
            GB_Input_Range.Name = "GB_Input_Range";
            GB_Input_Range.Size = new System.Drawing.Size(192, 142);
            GB_Input_Range.TabIndex = 160;
            GB_Input_Range.TabStop = false;
            GB_Input_Range.Text = "1st Input Range";
            // 
            // CP_Input_Range1
            // 
            CP_Input_Range1.BackColor = System.Drawing.Color.White;
            CP_Input_Range1.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            CP_Input_Range1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            CP_Input_Range1.BorderWidth = 1;
            CP_Input_Range1.Location = new System.Drawing.Point(1, 30);
            CP_Input_Range1.Name = "CP_Input_Range1";
            CP_Input_Range1.Size = new System.Drawing.Size(190, 112);
            CP_Input_Range1.TabIndex = 0;
            // 
            // GB_Expected_Output
            // 
            GB_Expected_Output.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            GB_Expected_Output.Controls.Add(CP_Output_Range);
            GB_Expected_Output.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 0);
            GB_Expected_Output.Location = new System.Drawing.Point(129, 229);
            GB_Expected_Output.Name = "GB_Expected_Output";
            GB_Expected_Output.Size = new System.Drawing.Size(192, 142);
            GB_Expected_Output.TabIndex = 161;
            GB_Expected_Output.TabStop = false;
            GB_Expected_Output.Text = "Expected Output";
            // 
            // CP_Output_Range
            // 
            CP_Output_Range.BackColor = System.Drawing.Color.White;
            CP_Output_Range.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            CP_Output_Range.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            CP_Output_Range.BorderWidth = 1;
            CP_Output_Range.Location = new System.Drawing.Point(1, 30);
            CP_Output_Range.Name = "CP_Output_Range";
            CP_Output_Range.Size = new System.Drawing.Size(190, 112);
            CP_Output_Range.TabIndex = 11;
            // 
            // PictureBox7
            // 
            PictureBox7.Image = (System.Drawing.Image)resources.GetObject("PictureBox7.Image");
            PictureBox7.Location = new System.Drawing.Point(194, 168);
            PictureBox7.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            PictureBox7.Name = "PictureBox7";
            PictureBox7.Size = new System.Drawing.Size(50, 60);
            PictureBox7.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            PictureBox7.TabIndex = 173;
            PictureBox7.TabStop = false;
            // 
            // GB_Display_Result
            // 
            GB_Display_Result.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            GB_Display_Result.Controls.Add(CustomGroupBox4);
            GB_Display_Result.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 0);
            GB_Display_Result.Location = new System.Drawing.Point(15, 290);
            GB_Display_Result.Name = "GB_Display_Result";
            GB_Display_Result.Size = new System.Drawing.Size(281, 94);
            GB_Display_Result.TabIndex = 175;
            GB_Display_Result.TabStop = false;
            GB_Display_Result.Text = "Display Result";
            // 
            // CustomGroupBox4
            // 
            CustomGroupBox4.BackColor = System.Drawing.Color.White;
            CustomGroupBox4.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            CustomGroupBox4.Controls.Add(CbFillFont);
            CustomGroupBox4.Controls.Add(CBFillBackground);
            CustomGroupBox4.Controls.Add(checkBoxFillFont);
            CustomGroupBox4.Controls.Add(checkBoxFillBack);
            CustomGroupBox4.Location = new System.Drawing.Point(1, 22);
            CustomGroupBox4.Name = "CustomGroupBox4";
            CustomGroupBox4.Size = new System.Drawing.Size(279, 72);
            CustomGroupBox4.TabIndex = 0;
            CustomGroupBox4.TabStop = false;
            // 
            // CbFillFont
            // 
            CbFillFont.BackColor = System.Drawing.Color.MidnightBlue;
            CbFillFont.DropDownHeight = 1;
            CbFillFont.DropDownWidth = 1;
            CbFillFont.ForeColor = System.Drawing.Color.Navy;
            CbFillFont.FormattingEnabled = true;
            CbFillFont.IntegralHeight = false;
            CbFillFont.Location = new System.Drawing.Point(162, 35);
            CbFillFont.Name = "CbFillFont";
            CbFillFont.Size = new System.Drawing.Size(110, 25);
            CbFillFont.TabIndex = 134;
            // 
            // CBFillBackground
            // 
            CBFillBackground.BackColor = System.Drawing.Color.LightSteelBlue;
            CBFillBackground.DropDownHeight = 1;
            CBFillBackground.DropDownWidth = 1;
            CBFillBackground.FormattingEnabled = true;
            CBFillBackground.IntegralHeight = false;
            CBFillBackground.Location = new System.Drawing.Point(8, 35);
            CBFillBackground.Name = "CBFillBackground";
            CBFillBackground.Size = new System.Drawing.Size(110, 25);
            CBFillBackground.TabIndex = 133;
            // 
            // checkBoxFillFont
            // 
            checkBoxFillFont.AutoSize = true;
            checkBoxFillFont.Location = new System.Drawing.Point(163, 8);
            checkBoxFillFont.Name = "checkBoxFillFont";
            checkBoxFillFont.Size = new System.Drawing.Size(106, 21);
            checkBoxFillFont.TabIndex = 132;
            checkBoxFillFont.Text = "Fill font color";
            checkBoxFillFont.UseVisualStyleBackColor = true;
            // 
            // checkBoxFillBack
            // 
            checkBoxFillBack.AutoSize = true;
            checkBoxFillBack.Location = new System.Drawing.Point(8, 8);
            checkBoxFillBack.Name = "checkBoxFillBack";
            checkBoxFillBack.Size = new System.Drawing.Size(120, 21);
            checkBoxFillBack.TabIndex = 131;
            checkBoxFillBack.Text = "Fill background";
            checkBoxFillBack.UseVisualStyleBackColor = true;
            // 
            // CustomGroupBox1
            // 
            CustomGroupBox1.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            CustomGroupBox1.Controls.Add(CustomGroupBox7);
            CustomGroupBox1.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 0);
            CustomGroupBox1.Location = new System.Drawing.Point(15, 163);
            CustomGroupBox1.Name = "CustomGroupBox1";
            CustomGroupBox1.Size = new System.Drawing.Size(281, 86);
            CustomGroupBox1.TabIndex = 165;
            CustomGroupBox1.TabStop = false;
            CustomGroupBox1.Text = "Compare Type";
            // 
            // CustomGroupBox7
            // 
            CustomGroupBox7.BackColor = System.Drawing.Color.White;
            CustomGroupBox7.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            CustomGroupBox7.Controls.Add(radBtnSameValues);
            CustomGroupBox7.Controls.Add(PictureBox1);
            CustomGroupBox7.Controls.Add(PictureBox5);
            CustomGroupBox7.Controls.Add(radBtnDifferentValues);
            CustomGroupBox7.Location = new System.Drawing.Point(1, 22);
            CustomGroupBox7.Name = "CustomGroupBox7";
            CustomGroupBox7.Size = new System.Drawing.Size(279, 64);
            CustomGroupBox7.TabIndex = 0;
            CustomGroupBox7.TabStop = false;
            // 
            // radBtnSameValues
            // 
            radBtnSameValues.AutoSize = true;
            radBtnSameValues.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            radBtnSameValues.Location = new System.Drawing.Point(8, 8);
            radBtnSameValues.Name = "radBtnSameValues";
            radBtnSameValues.Size = new System.Drawing.Size(157, 21);
            radBtnSameValues.TabIndex = 129;
            radBtnSameValues.TabStop = true;
            radBtnSameValues.Text = "Cells with Same Values";
            radBtnSameValues.UseVisualStyleBackColor = true;
            // 
            // PictureBox1
            // 
            PictureBox1.Image = (System.Drawing.Image)resources.GetObject("PictureBox1.Image");
            PictureBox1.Location = new System.Drawing.Point(245, 36);
            PictureBox1.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            PictureBox1.Name = "PictureBox1";
            PictureBox1.Size = new System.Drawing.Size(20, 20);
            PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            PictureBox1.TabIndex = 128;
            PictureBox1.TabStop = false;
            // 
            // PictureBox5
            // 
            PictureBox5.Image = (System.Drawing.Image)resources.GetObject("PictureBox5.Image");
            PictureBox5.Location = new System.Drawing.Point(245, 8);
            PictureBox5.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            PictureBox5.Name = "PictureBox5";
            PictureBox5.Size = new System.Drawing.Size(20, 20);
            PictureBox5.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            PictureBox5.TabIndex = 127;
            PictureBox5.TabStop = false;
            // 
            // radBtnDifferentValues
            // 
            radBtnDifferentValues.AutoSize = true;
            radBtnDifferentValues.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            radBtnDifferentValues.Location = new System.Drawing.Point(8, 34);
            radBtnDifferentValues.Name = "radBtnDifferentValues";
            radBtnDifferentValues.Size = new System.Drawing.Size(175, 21);
            radBtnDifferentValues.TabIndex = 0;
            radBtnDifferentValues.TabStop = true;
            radBtnDifferentValues.Text = "Cells with Different Values";
            radBtnDifferentValues.UseVisualStyleBackColor = true;
            // 
            // CustomGroupBox2
            // 
            CustomGroupBox2.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            CustomGroupBox2.Controls.Add(rngSelection2);
            CustomGroupBox2.Controls.Add(rngSelection1);
            CustomGroupBox2.Controls.Add(AutoSelection2);
            CustomGroupBox2.Controls.Add(AutoSelection1);
            CustomGroupBox2.Controls.Add(txtSourceRange2);
            CustomGroupBox2.Controls.Add(txtSourceRange1);
            CustomGroupBox2.Controls.Add(lblSourceRng2);
            CustomGroupBox2.Controls.Add(lblSourceRng1);
            CustomGroupBox2.Location = new System.Drawing.Point(15, 15);
            CustomGroupBox2.Name = "CustomGroupBox2";
            CustomGroupBox2.Size = new System.Drawing.Size(278, 134);
            CustomGroupBox2.TabIndex = 174;
            CustomGroupBox2.TabStop = false;
            // 
            // rngSelection2
            // 
            rngSelection2.BackColor = System.Drawing.Color.White;
            rngSelection2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            rngSelection2.Image = (System.Drawing.Image)resources.GetObject("rngSelection2.Image");
            rngSelection2.Location = new System.Drawing.Point(244, 95);
            rngSelection2.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            rngSelection2.Name = "rngSelection2";
            rngSelection2.Size = new System.Drawing.Size(24, 25);
            rngSelection2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            rngSelection2.TabIndex = 180;
            rngSelection2.TabStop = false;
            // 
            // rngSelection1
            // 
            rngSelection1.BackColor = System.Drawing.Color.White;
            rngSelection1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            rngSelection1.Image = (System.Drawing.Image)resources.GetObject("rngSelection1.Image");
            rngSelection1.Location = new System.Drawing.Point(242, 33);
            rngSelection1.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            rngSelection1.Name = "rngSelection1";
            rngSelection1.Size = new System.Drawing.Size(24, 25);
            rngSelection1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            rngSelection1.TabIndex = 171;
            rngSelection1.TabStop = false;
            // 
            // AutoSelection2
            // 
            AutoSelection2.BackColor = System.Drawing.Color.White;
            AutoSelection2.Image = (System.Drawing.Image)resources.GetObject("AutoSelection2.Image");
            AutoSelection2.Location = new System.Drawing.Point(219, 96);
            AutoSelection2.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            AutoSelection2.Name = "AutoSelection2";
            AutoSelection2.Size = new System.Drawing.Size(24, 23);
            AutoSelection2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            AutoSelection2.TabIndex = 179;
            AutoSelection2.TabStop = false;
            // 
            // AutoSelection1
            // 
            AutoSelection1.BackColor = System.Drawing.Color.White;
            AutoSelection1.Image = (System.Drawing.Image)resources.GetObject("AutoSelection1.Image");
            AutoSelection1.Location = new System.Drawing.Point(217, 34);
            AutoSelection1.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            AutoSelection1.Name = "AutoSelection1";
            AutoSelection1.Size = new System.Drawing.Size(24, 23);
            AutoSelection1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            AutoSelection1.TabIndex = 170;
            AutoSelection1.TabStop = false;
            // 
            // txtSourceRange2
            // 
            txtSourceRange2.BackColor = System.Drawing.Color.White;
            txtSourceRange2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            txtSourceRange2.Font = new System.Drawing.Font("Segoe UI", 10.0f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            txtSourceRange2.ForeColor = System.Drawing.Color.FromArgb(70, 70, 70);
            txtSourceRange2.Location = new System.Drawing.Point(11, 95);
            txtSourceRange2.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            txtSourceRange2.Name = "txtSourceRange2";
            txtSourceRange2.Size = new System.Drawing.Size(256, 25);
            txtSourceRange2.TabIndex = 178;
            // 
            // txtSourceRange1
            // 
            txtSourceRange1.BackColor = System.Drawing.Color.White;
            txtSourceRange1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            txtSourceRange1.Font = new System.Drawing.Font("Segoe UI", 10.0f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            txtSourceRange1.ForeColor = System.Drawing.Color.FromArgb(70, 70, 70);
            txtSourceRange1.Location = new System.Drawing.Point(9, 33);
            txtSourceRange1.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            txtSourceRange1.Name = "txtSourceRange1";
            txtSourceRange1.Size = new System.Drawing.Size(256, 25);
            txtSourceRange1.TabIndex = 169;
            // 
            // lblSourceRng2
            // 
            lblSourceRng2.AutoSize = true;
            lblSourceRng2.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            lblSourceRng2.Location = new System.Drawing.Point(8, 68);
            lblSourceRng2.Name = "lblSourceRng2";
            lblSourceRng2.Size = new System.Drawing.Size(256, 17);
            lblSourceRng2.TabIndex = 168;
            lblSourceRng2.Text = "2nd Source Range (X rows x Y columns) :";
            // 
            // lblSourceRng1
            // 
            lblSourceRng1.AutoSize = true;
            lblSourceRng1.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            lblSourceRng1.Location = new System.Drawing.Point(8, 8);
            lblSourceRng1.Name = "lblSourceRng1";
            lblSourceRng1.Size = new System.Drawing.Size(249, 17);
            lblSourceRng1.TabIndex = 164;
            lblSourceRng1.Text = "1st Source Range (X rows x Y columns) :";
            // 
            // Form15CompareCells
            // 
            AutoScaleDimensions = new System.Drawing.SizeF(6.0f, 13.0f);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            ClientSize = new System.Drawing.Size(775, 465);
            Controls.Add(CustomPanel1);
            Controls.Add(checkBoxCase);
            Controls.Add(GB_Display_Result);
            Controls.Add(checkBoxFormatting);
            Controls.Add(btnOK);
            Controls.Add(btnCanecl);
            Controls.Add(ComboBox1);
            Controls.Add(checkBoxCopyWs);
            Controls.Add(CustomGroupBox1);
            Controls.Add(CustomGroupBox2);
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "Form15CompareCells";
            Text = "Compare Cells";
            CustomPanel1.ResumeLayout(false);
            CustomGroupBox5.ResumeLayout(false);
            GB_Input_Range.ResumeLayout(false);
            GB_Expected_Output.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)PictureBox7).EndInit();
            GB_Display_Result.ResumeLayout(false);
            CustomGroupBox4.ResumeLayout(false);
            CustomGroupBox4.PerformLayout();
            CustomGroupBox1.ResumeLayout(false);
            CustomGroupBox7.ResumeLayout(false);
            CustomGroupBox7.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)PictureBox1).EndInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox5).EndInit();
            CustomGroupBox2.ResumeLayout(false);
            CustomGroupBox2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)rngSelection2).EndInit();
            ((System.ComponentModel.ISupportInitialize)rngSelection1).EndInit();
            ((System.ComponentModel.ISupportInitialize)AutoSelection2).EndInit();
            ((System.ComponentModel.ISupportInitialize)AutoSelection1).EndInit();
            KeyDown += new System.Windows.Forms.KeyEventHandler(Form1_KeyDown);
            Load += new EventHandler(Form15CompareCells_Load);
            Activated += new EventHandler(Form1_Activated);
            Closing += new System.ComponentModel.CancelEventHandler(Form15CompareCells_Closing);
            Shown += new EventHandler(Form15CompareCells_Shown);
            Disposed += new EventHandler(Form15CompareCells_Disposed);
            ResumeLayout(false);
            PerformLayout();

        }
        internal System.Windows.Forms.CheckBox checkBoxFormatting;
        internal System.Windows.Forms.PictureBox PictureBox7;
        internal System.Windows.Forms.Button btnOK;
        internal System.Windows.Forms.Button btnCanecl;
        internal System.Windows.Forms.ComboBox ComboBox1;
        internal System.Windows.Forms.CheckBox checkBoxCopyWs;
        internal System.Windows.Forms.RadioButton radBtnSameValues;
        internal System.Windows.Forms.PictureBox PictureBox1;
        internal CustomGroupBox CustomGroupBox1;
        internal CustomGroupBox CustomGroupBox7;
        internal System.Windows.Forms.PictureBox PictureBox5;
        internal System.Windows.Forms.RadioButton radBtnDifferentValues;
        internal CustomGroupBox CustomGroupBox2;
        internal System.Windows.Forms.Label lblSourceRng2;
        internal System.Windows.Forms.Label lblSourceRng1;
        internal CustomGroupBox GB_Display_Result;
        internal CustomGroupBox CustomGroupBox4;
        internal System.Windows.Forms.CheckBox checkBoxCase;
        internal System.Windows.Forms.ComboBox CBFillBackground;
        internal System.Windows.Forms.CheckBox checkBoxFillFont;
        internal System.Windows.Forms.CheckBox checkBoxFillBack;
        internal CustomPanel CustomPanel1;
        internal CustomGroupBox CustomGroupBox5;
        internal CustomPanel CP_Input_Range2;
        internal CustomGroupBox GB_Input_Range;
        internal CustomPanel CP_Input_Range1;
        internal CustomGroupBox GB_Expected_Output;
        internal CustomPanel CP_Output_Range;
        internal System.Windows.Forms.PictureBox rngSelection2;
        internal System.Windows.Forms.PictureBox rngSelection1;
        internal System.Windows.Forms.PictureBox AutoSelection2;
        internal System.Windows.Forms.PictureBox AutoSelection1;
        internal System.Windows.Forms.TextBox txtSourceRange2;
        internal System.Windows.Forms.TextBox txtSourceRange1;
        internal System.Windows.Forms.ColorDialog CD_Fill_Background;
        internal System.Windows.Forms.ColorDialog CD_Fill_Font;
        internal System.Windows.Forms.ComboBox CbFillFont;
    }
}