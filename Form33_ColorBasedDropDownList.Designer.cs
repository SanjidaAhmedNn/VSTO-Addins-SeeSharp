using System;

namespace VSTO_Addins
{
    [Microsoft.VisualBasic.CompilerServices.DesignerGenerated()]
    public partial class Form33_ColorBasedDropDownList : System.Windows.Forms.Form
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
            var resources = new System.ComponentModel.ComponentResourceManager(typeof(Form33_ColorBasedDropDownList));
            ColorDialog1 = new System.Windows.Forms.ColorDialog();
            FlowLayoutPanel1 = new System.Windows.Forms.FlowLayoutPanel();
            ToolTip1 = new System.Windows.Forms.ToolTip(components);
            PictureBox1 = new System.Windows.Forms.PictureBox();
            PictureBox3 = new System.Windows.Forms.PictureBox();
            Label1 = new System.Windows.Forms.Label();
            Selection_source = new System.Windows.Forms.PictureBox();
            Selection_source.Click += new EventHandler(Selection_source_Click);
            TB_src_rng = new System.Windows.Forms.TextBox();
            TB_src_rng.TextChanged += new EventHandler(TB_src_rng_TextChanged);
            Selection_destination = new System.Windows.Forms.PictureBox();
            Selection_destination.Click += new EventHandler(Selection_destination_Click);
            TB_des_rng = new System.Windows.Forms.TextBox();
            TB_des_rng.TextChanged += new EventHandler(TB_des_rng_TextChanged);
            Label2 = new System.Windows.Forms.Label();
            List_Preview = new System.Windows.Forms.ListBox();
            List_Preview.SelectedIndexChanged += new EventHandler(List_Box_IndexChanged);
            List_Preview.DrawItem += new System.Windows.Forms.DrawItemEventHandler(ListBox1_DrawItem);
            Label6 = new System.Windows.Forms.Label();
            btn_OK = new System.Windows.Forms.Button();
            btn_OK.Click += new EventHandler(btn_OK_Click);
            btn_Cancel = new System.Windows.Forms.Button();
            btn_Cancel.Click += new EventHandler(btn_Cancel_Click);
            ComboBox1 = new System.Windows.Forms.ComboBox();
            Backup_sheet = new System.Windows.Forms.CheckBox();
            Label3 = new System.Windows.Forms.Label();
            Btn_NC = new System.Windows.Forms.Button();
            Btn_NC.Click += new EventHandler(Btn_NC_Click);
            Button2 = new System.Windows.Forms.Button();
            Button2.Click += new EventHandler(Button2_Click);
            Btn_color = new System.Windows.Forms.Button();
            Btn_color.Click += new EventHandler(Btn_color_Click);
            Button1 = new System.Windows.Forms.Button();
            Button1.Click += new EventHandler(Button1_Click);
            GB_sample = new CustomGroupBox();
            CustomGroupBox1 = new CustomGroupBox();
            CustomGroupBox7 = new CustomGroupBox();
            RB_Row = new System.Windows.Forms.RadioButton();
            RB_Row.CheckedChanged += new EventHandler(RB_Row_CheckedChanged);
            RB_cell = new System.Windows.Forms.RadioButton();
            RB_cell.CheckedChanged += new EventHandler(RB_cell_CheckedChanged);
            ((System.ComponentModel.ISupportInitialize)PictureBox1).BeginInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox3).BeginInit();
            ((System.ComponentModel.ISupportInitialize)Selection_source).BeginInit();
            ((System.ComponentModel.ISupportInitialize)Selection_destination).BeginInit();
            CustomGroupBox1.SuspendLayout();
            CustomGroupBox7.SuspendLayout();
            SuspendLayout();
            // 
            // FlowLayoutPanel1
            // 
            FlowLayoutPanel1.BackColor = System.Drawing.Color.White;
            FlowLayoutPanel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            FlowLayoutPanel1.Location = new System.Drawing.Point(154, 262);
            FlowLayoutPanel1.Name = "FlowLayoutPanel1";
            FlowLayoutPanel1.Size = new System.Drawing.Size(186, 224);
            FlowLayoutPanel1.TabIndex = 0;
            // 
            // PictureBox1
            // 
            PictureBox1.Image = (System.Drawing.Image)resources.GetObject("PictureBox1.Image");
            PictureBox1.Location = new System.Drawing.Point(269, 15);
            PictureBox1.Name = "PictureBox1";
            PictureBox1.Size = new System.Drawing.Size(20, 20);
            PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            PictureBox1.TabIndex = 407;
            PictureBox1.TabStop = false;
            ToolTip1.SetToolTip(PictureBox1, "Please select a range that contains a data validation list");
            // 
            // PictureBox3
            // 
            PictureBox3.Image = (System.Drawing.Image)resources.GetObject("PictureBox3.Image");
            PictureBox3.Location = new System.Drawing.Point(212, 170);
            PictureBox3.Name = "PictureBox3";
            PictureBox3.Size = new System.Drawing.Size(20, 20);
            PictureBox3.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            PictureBox3.TabIndex = 408;
            PictureBox3.TabStop = false;
            ToolTip1.SetToolTip(PictureBox3, "Please select a range intersecting data validation list");
            // 
            // Label1
            // 
            Label1.AutoSize = true;
            Label1.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            Label1.Location = new System.Drawing.Point(15, 15);
            Label1.Name = "Label1";
            Label1.Size = new System.Drawing.Size(253, 17);
            Label1.TabIndex = 2;
            Label1.Text = "Drop-down List (Data Validation) Range :";
            // 
            // Selection_source
            // 
            Selection_source.BackColor = System.Drawing.Color.White;
            Selection_source.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            Selection_source.Image = (System.Drawing.Image)resources.GetObject("Selection_source.Image");
            Selection_source.Location = new System.Drawing.Point(319, 41);
            Selection_source.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            Selection_source.Name = "Selection_source";
            Selection_source.Size = new System.Drawing.Size(24, 25);
            Selection_source.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            Selection_source.TabIndex = 391;
            Selection_source.TabStop = false;
            // 
            // TB_src_rng
            // 
            TB_src_rng.Font = new System.Drawing.Font("Segoe UI", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            TB_src_rng.Location = new System.Drawing.Point(15, 41);
            TB_src_rng.Name = "TB_src_rng";
            TB_src_rng.Size = new System.Drawing.Size(328, 25);
            TB_src_rng.TabIndex = 390;
            // 
            // Selection_destination
            // 
            Selection_destination.BackColor = System.Drawing.Color.White;
            Selection_destination.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            Selection_destination.Image = (System.Drawing.Image)resources.GetObject("Selection_destination.Image");
            Selection_destination.Location = new System.Drawing.Point(319, 197);
            Selection_destination.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            Selection_destination.Name = "Selection_destination";
            Selection_destination.Size = new System.Drawing.Size(24, 25);
            Selection_destination.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            Selection_destination.TabIndex = 395;
            Selection_destination.TabStop = false;
            // 
            // TB_des_rng
            // 
            TB_des_rng.Font = new System.Drawing.Font("Segoe UI", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            TB_des_rng.Location = new System.Drawing.Point(15, 197);
            TB_des_rng.Name = "TB_des_rng";
            TB_des_rng.Size = new System.Drawing.Size(328, 25);
            TB_des_rng.TabIndex = 394;
            // 
            // Label2
            // 
            Label2.AutoSize = true;
            Label2.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            Label2.Location = new System.Drawing.Point(15, 170);
            Label2.Name = "Label2";
            Label2.Size = new System.Drawing.Size(197, 17);
            Label2.TabIndex = 393;
            Label2.Text = "Select range to highlight rows :";
            // 
            // List_Preview
            // 
            List_Preview.BackColor = System.Drawing.SystemColors.Window;
            List_Preview.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawVariable;
            List_Preview.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            List_Preview.FormattingEnabled = true;
            List_Preview.ItemHeight = 17;
            List_Preview.Location = new System.Drawing.Point(15, 262);
            List_Preview.Name = "List_Preview";
            List_Preview.Size = new System.Drawing.Size(125, 259);
            List_Preview.TabIndex = 397;
            // 
            // Label6
            // 
            Label6.AutoSize = true;
            Label6.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            Label6.Location = new System.Drawing.Point(15, 236);
            Label6.Name = "Label6";
            Label6.Size = new System.Drawing.Size(86, 17);
            Label6.TabIndex = 396;
            Label6.Text = "List Preview :";
            // 
            // btn_OK
            // 
            btn_OK.BackColor = System.Drawing.Color.White;
            btn_OK.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            btn_OK.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            btn_OK.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            btn_OK.ForeColor = System.Drawing.Color.FromArgb(70, 70, 70);
            btn_OK.Location = new System.Drawing.Point(554, 560);
            btn_OK.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            btn_OK.Name = "btn_OK";
            btn_OK.Size = new System.Drawing.Size(62, 26);
            btn_OK.TabIndex = 402;
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
            btn_Cancel.Location = new System.Drawing.Point(627, 560);
            btn_Cancel.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            btn_Cancel.Name = "btn_Cancel";
            btn_Cancel.Size = new System.Drawing.Size(62, 26);
            btn_Cancel.TabIndex = 401;
            btn_Cancel.Text = "Cancel";
            btn_Cancel.UseVisualStyleBackColor = false;
            // 
            // ComboBox1
            // 
            ComboBox1.Font = new System.Drawing.Font("Segoe UI", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            ComboBox1.FormattingEnabled = true;
            ComboBox1.Items.AddRange(new object[] { "SOFTEKO", "About Us", "Help", "Feedback" });
            ComboBox1.Location = new System.Drawing.Point(13, 562);
            ComboBox1.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            ComboBox1.Name = "ComboBox1";
            ComboBox1.Size = new System.Drawing.Size(100, 25);
            ComboBox1.TabIndex = 400;
            ComboBox1.Text = "SOFTEKO";
            // 
            // Backup_sheet
            // 
            Backup_sheet.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            Backup_sheet.Location = new System.Drawing.Point(15, 526);
            Backup_sheet.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            Backup_sheet.Name = "Backup_sheet";
            Backup_sheet.Size = new System.Drawing.Size(258, 29);
            Backup_sheet.TabIndex = 399;
            Backup_sheet.Text = "Create a copy of the original worksheet";
            Backup_sheet.UseVisualStyleBackColor = true;
            // 
            // Label3
            // 
            Label3.AutoSize = true;
            Label3.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            Label3.Location = new System.Drawing.Point(154, 236);
            Label3.Name = "Label3";
            Label3.Size = new System.Drawing.Size(93, 17);
            Label3.TabIndex = 403;
            Label3.Text = "Color Palette :";
            // 
            // Btn_NC
            // 
            Btn_NC.BackColor = System.Drawing.Color.White;
            Btn_NC.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            Btn_NC.Font = new System.Drawing.Font("Segoe UI", 9.0f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            Btn_NC.Location = new System.Drawing.Point(159, 422);
            Btn_NC.Name = "Btn_NC";
            Btn_NC.Size = new System.Drawing.Size(176, 22);
            Btn_NC.TabIndex = 404;
            Btn_NC.Text = "No Color";
            Btn_NC.UseVisualStyleBackColor = true;
            // 
            // Button2
            // 
            Button2.BackColor = System.Drawing.Color.White;
            Button2.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            Button2.Font = new System.Drawing.Font("Segoe UI", 9.0f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            Button2.Location = new System.Drawing.Point(190, 454);
            Button2.Name = "Button2";
            Button2.Size = new System.Drawing.Size(144, 22);
            Button2.TabIndex = 405;
            Button2.Text = "More Colors";
            Button2.UseVisualStyleBackColor = false;
            // 
            // Btn_color
            // 
            Btn_color.BackColor = System.Drawing.Color.White;
            Btn_color.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            Btn_color.Font = new System.Drawing.Font("Segoe UI", 9.0f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            Btn_color.Location = new System.Drawing.Point(159, 454);
            Btn_color.Name = "Btn_color";
            Btn_color.Size = new System.Drawing.Size(24, 22);
            Btn_color.TabIndex = 406;
            Btn_color.UseVisualStyleBackColor = false;
            // 
            // Button1
            // 
            Button1.BackColor = System.Drawing.Color.White;
            Button1.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            Button1.Font = new System.Drawing.Font("Segoe UI", 9.0f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            Button1.Location = new System.Drawing.Point(154, 494);
            Button1.Name = "Button1";
            Button1.Size = new System.Drawing.Size(186, 24);
            Button1.TabIndex = 409;
            Button1.Text = "Clear all selected colors";
            Button1.UseVisualStyleBackColor = false;
            // 
            // GB_sample
            // 
            GB_sample.BackColor = System.Drawing.Color.White;
            GB_sample.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            GB_sample.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 0);
            GB_sample.Location = new System.Drawing.Point(382, 15);
            GB_sample.Name = "GB_sample";
            GB_sample.Size = new System.Drawing.Size(307, 503);
            GB_sample.TabIndex = 398;
            GB_sample.TabStop = false;
            GB_sample.Text = "Sample Image";
            // 
            // CustomGroupBox1
            // 
            CustomGroupBox1.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            CustomGroupBox1.Controls.Add(CustomGroupBox7);
            CustomGroupBox1.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            CustomGroupBox1.Location = new System.Drawing.Point(15, 75);
            CustomGroupBox1.Name = "CustomGroupBox1";
            CustomGroupBox1.Size = new System.Drawing.Size(328, 84);
            CustomGroupBox1.TabIndex = 392;
            CustomGroupBox1.TabStop = false;
            CustomGroupBox1.Text = "Apply Color to";
            // 
            // CustomGroupBox7
            // 
            CustomGroupBox7.BackColor = System.Drawing.Color.White;
            CustomGroupBox7.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            CustomGroupBox7.Controls.Add(RB_Row);
            CustomGroupBox7.Controls.Add(RB_cell);
            CustomGroupBox7.Location = new System.Drawing.Point(1, 22);
            CustomGroupBox7.Name = "CustomGroupBox7";
            CustomGroupBox7.Size = new System.Drawing.Size(327, 62);
            CustomGroupBox7.TabIndex = 0;
            CustomGroupBox7.TabStop = false;
            // 
            // RB_Row
            // 
            RB_Row.AutoSize = true;
            RB_Row.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            RB_Row.Location = new System.Drawing.Point(8, 32);
            RB_Row.Name = "RB_Row";
            RB_Row.Size = new System.Drawing.Size(237, 21);
            RB_Row.TabIndex = 1;
            RB_Row.Text = "Full row of the drop-down list range";
            RB_Row.UseVisualStyleBackColor = true;
            // 
            // RB_cell
            // 
            RB_cell.AutoSize = true;
            RB_cell.Checked = true;
            RB_cell.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            RB_cell.Location = new System.Drawing.Point(8, 6);
            RB_cell.Name = "RB_cell";
            RB_cell.Size = new System.Drawing.Size(265, 21);
            RB_cell.TabIndex = 0;
            RB_cell.TabStop = true;
            RB_cell.Text = "Only the cell that contains data validation";
            RB_cell.UseVisualStyleBackColor = true;
            // 
            // Form33_ColorBasedDropDownList
            // 
            AutoScaleDimensions = new System.Drawing.SizeF(6.0f, 13.0f);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            ClientSize = new System.Drawing.Size(713, 604);
            Controls.Add(Button1);
            Controls.Add(PictureBox3);
            Controls.Add(PictureBox1);
            Controls.Add(Btn_color);
            Controls.Add(Button2);
            Controls.Add(Btn_NC);
            Controls.Add(Label3);
            Controls.Add(btn_OK);
            Controls.Add(btn_Cancel);
            Controls.Add(ComboBox1);
            Controls.Add(Backup_sheet);
            Controls.Add(GB_sample);
            Controls.Add(List_Preview);
            Controls.Add(Label6);
            Controls.Add(Selection_destination);
            Controls.Add(TB_des_rng);
            Controls.Add(Label2);
            Controls.Add(CustomGroupBox1);
            Controls.Add(Selection_source);
            Controls.Add(TB_src_rng);
            Controls.Add(Label1);
            Controls.Add(FlowLayoutPanel1);
            ForeColor = System.Drawing.SystemColors.ControlText;
            KeyPreview = true;
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "Form33_ColorBasedDropDownList";
            Text = "Color Based Drop-down List";
            ((System.ComponentModel.ISupportInitialize)PictureBox1).EndInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox3).EndInit();
            ((System.ComponentModel.ISupportInitialize)Selection_source).EndInit();
            ((System.ComponentModel.ISupportInitialize)Selection_destination).EndInit();
            CustomGroupBox1.ResumeLayout(false);
            CustomGroupBox7.ResumeLayout(false);
            CustomGroupBox7.PerformLayout();
            KeyDown += new System.Windows.Forms.KeyEventHandler(Form1_KeyDown);
            Load += new EventHandler(Form1_Load);
            Closing += new System.ComponentModel.CancelEventHandler(Form33_ColorBasedDropDownList_Closing);
            Disposed += new EventHandler(Form33_ColorBasedDropDownList_Disposed);
            Shown += new EventHandler(Form33_ColorBasedDropDownList_Shown);
            ResumeLayout(false);
            PerformLayout();

        }
        internal System.Windows.Forms.ColorDialog ColorDialog1;
        internal System.Windows.Forms.FlowLayoutPanel FlowLayoutPanel1;
        internal System.Windows.Forms.ToolTip ToolTip1;
        internal System.Windows.Forms.Label Label1;
        internal System.Windows.Forms.PictureBox Selection_source;
        internal System.Windows.Forms.TextBox TB_src_rng;
        internal CustomGroupBox CustomGroupBox1;
        internal CustomGroupBox CustomGroupBox7;
        internal System.Windows.Forms.RadioButton RB_Row;
        internal System.Windows.Forms.RadioButton RB_cell;
        internal System.Windows.Forms.PictureBox Selection_destination;
        internal System.Windows.Forms.TextBox TB_des_rng;
        internal System.Windows.Forms.Label Label2;
        internal System.Windows.Forms.ListBox List_Preview;
        internal System.Windows.Forms.Label Label6;
        internal CustomGroupBox GB_sample;
        internal System.Windows.Forms.Button btn_OK;
        internal System.Windows.Forms.Button btn_Cancel;
        internal System.Windows.Forms.ComboBox ComboBox1;
        internal System.Windows.Forms.CheckBox Backup_sheet;
        internal System.Windows.Forms.Label Label3;
        internal System.Windows.Forms.Button Btn_NC;
        internal System.Windows.Forms.Button Button2;
        internal System.Windows.Forms.Button Btn_color;
        internal System.Windows.Forms.PictureBox PictureBox1;
        internal System.Windows.Forms.PictureBox PictureBox3;
        internal System.Windows.Forms.Button Button1;
    }
}