using System;

namespace VSTO_Addins
{
    [Microsoft.VisualBasic.CompilerServices.DesignerGenerated()]
    public partial class Form34_PictureBasedDropdownList : System.Windows.Forms.Form
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
            var resources = new System.ComponentModel.ComponentResourceManager(typeof(Form34_PictureBasedDropdownList));
            Src_selection = new System.Windows.Forms.PictureBox();
            Src_selection.Click += new EventHandler(PictureBox9_Click);
            Src_selection.KeyDown += new System.Windows.Forms.KeyEventHandler(source);
            Btn_OK = new System.Windows.Forms.Button();
            Btn_OK.Click += new EventHandler(Btn_OK_Click);
            ComboBox1 = new System.Windows.Forms.ComboBox();
            ComboBox1.KeyDown += new System.Windows.Forms.KeyEventHandler(Combobox1_enter);
            TB_src_rng = new System.Windows.Forms.TextBox();
            TB_src_rng.TextChanged += new EventHandler(TB_src_rng_TextChanged);
            TB_src_rng.KeyDown += new System.Windows.Forms.KeyEventHandler(source_TextBox);
            Label1 = new System.Windows.Forms.Label();
            Btn_Cancel = new System.Windows.Forms.Button();
            Btn_Cancel.Click += new EventHandler(Btn_Cancel_Click);
            Des_selection = new System.Windows.Forms.PictureBox();
            Des_selection.Click += new EventHandler(PictureBox1_Click);
            Des_selection.KeyDown += new System.Windows.Forms.KeyEventHandler(Destination);
            TB_des_rng = new System.Windows.Forms.TextBox();
            TB_des_rng.TextChanged += new EventHandler(TB_des_rng_TextChanged);
            TB_des_rng.KeyDown += new System.Windows.Forms.KeyEventHandler(destination_TextBox);
            Label2 = new System.Windows.Forms.Label();
            Info = new System.Windows.Forms.PictureBox();
            PictureBox2 = new System.Windows.Forms.PictureBox();
            ToolTip1 = new System.Windows.Forms.ToolTip(components);
            CustomGroupBox2 = new CustomGroupBox();
            ((System.ComponentModel.ISupportInitialize)Src_selection).BeginInit();
            ((System.ComponentModel.ISupportInitialize)Des_selection).BeginInit();
            ((System.ComponentModel.ISupportInitialize)Info).BeginInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox2).BeginInit();
            SuspendLayout();
            // 
            // Src_selection
            // 
            Src_selection.BackColor = System.Drawing.Color.White;
            Src_selection.Image = (System.Drawing.Image)resources.GetObject("Src_selection.Image");
            Src_selection.Location = new System.Drawing.Point(249, 44);
            Src_selection.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            Src_selection.Name = "Src_selection";
            Src_selection.Size = new System.Drawing.Size(24, 23);
            Src_selection.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            Src_selection.TabIndex = 203;
            Src_selection.TabStop = false;
            // 
            // Btn_OK
            // 
            Btn_OK.BackColor = System.Drawing.Color.White;
            Btn_OK.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            Btn_OK.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            Btn_OK.Location = new System.Drawing.Point(465, 165);
            Btn_OK.Name = "Btn_OK";
            Btn_OK.Size = new System.Drawing.Size(62, 26);
            Btn_OK.TabIndex = 202;
            Btn_OK.Text = "OK";
            Btn_OK.UseVisualStyleBackColor = false;
            // 
            // ComboBox1
            // 
            ComboBox1.Font = new System.Drawing.Font("Segoe UI", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            ComboBox1.FormattingEnabled = true;
            ComboBox1.Location = new System.Drawing.Point(18, 165);
            ComboBox1.Name = "ComboBox1";
            ComboBox1.Size = new System.Drawing.Size(90, 25);
            ComboBox1.TabIndex = 198;
            ComboBox1.Text = "SOFTEKO";
            // 
            // TB_src_rng
            // 
            TB_src_rng.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            TB_src_rng.Cursor = System.Windows.Forms.Cursors.IBeam;
            TB_src_rng.Font = new System.Drawing.Font("Segoe UI", 10.0f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            TB_src_rng.Location = new System.Drawing.Point(15, 43);
            TB_src_rng.Name = "TB_src_rng";
            TB_src_rng.Size = new System.Drawing.Size(259, 25);
            TB_src_rng.TabIndex = 196;
            // 
            // Label1
            // 
            Label1.AutoSize = true;
            Label1.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 0);
            Label1.Location = new System.Drawing.Point(15, 15);
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
            Btn_Cancel.Location = new System.Drawing.Point(543, 165);
            Btn_Cancel.Name = "Btn_Cancel";
            Btn_Cancel.Size = new System.Drawing.Size(62, 26);
            Btn_Cancel.TabIndex = 201;
            Btn_Cancel.Text = "Cancel";
            Btn_Cancel.UseVisualStyleBackColor = false;
            // 
            // Des_selection
            // 
            Des_selection.BackColor = System.Drawing.Color.White;
            Des_selection.Image = (System.Drawing.Image)resources.GetObject("Des_selection.Image");
            Des_selection.Location = new System.Drawing.Point(249, 114);
            Des_selection.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            Des_selection.Name = "Des_selection";
            Des_selection.Size = new System.Drawing.Size(24, 23);
            Des_selection.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            Des_selection.TabIndex = 206;
            Des_selection.TabStop = false;
            // 
            // TB_des_rng
            // 
            TB_des_rng.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            TB_des_rng.Cursor = System.Windows.Forms.Cursors.IBeam;
            TB_des_rng.Font = new System.Drawing.Font("Segoe UI", 10.0f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            TB_des_rng.Location = new System.Drawing.Point(15, 113);
            TB_des_rng.Name = "TB_des_rng";
            TB_des_rng.Size = new System.Drawing.Size(259, 25);
            TB_des_rng.TabIndex = 205;
            // 
            // Label2
            // 
            Label2.AutoSize = true;
            Label2.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 0);
            Label2.Location = new System.Drawing.Point(15, 81);
            Label2.Name = "Label2";
            Label2.Size = new System.Drawing.Size(126, 17);
            Label2.TabIndex = 204;
            Label2.Text = "Destination Range :";
            // 
            // Info
            // 
            Info.Image = (System.Drawing.Image)resources.GetObject("Info.Image");
            Info.Location = new System.Drawing.Point(119, 15);
            Info.Name = "Info";
            Info.Size = new System.Drawing.Size(20, 20);
            Info.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            Info.TabIndex = 400;
            Info.TabStop = false;
            ToolTip1.SetToolTip(Info, "Please, select both of the columns that contain the data and the relevant images");
            // 
            // PictureBox2
            // 
            PictureBox2.Image = (System.Drawing.Image)resources.GetObject("PictureBox2.Image");
            PictureBox2.Location = new System.Drawing.Point(147, 81);
            PictureBox2.Name = "PictureBox2";
            PictureBox2.Size = new System.Drawing.Size(20, 20);
            PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            PictureBox2.TabIndex = 401;
            PictureBox2.TabStop = false;
            ToolTip1.SetToolTip(PictureBox2, "Please, select 2 columns");
            // 
            // CustomGroupBox2
            // 
            CustomGroupBox2.BackColor = System.Drawing.Color.White;
            CustomGroupBox2.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            CustomGroupBox2.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 0);
            CustomGroupBox2.Location = new System.Drawing.Point(301, 15);
            CustomGroupBox2.Name = "CustomGroupBox2";
            CustomGroupBox2.Size = new System.Drawing.Size(304, 126);
            CustomGroupBox2.TabIndex = 399;
            CustomGroupBox2.TabStop = false;
            CustomGroupBox2.Text = "Sample Image";
            // 
            // Form34_PictureBasedDropdownList
            // 
            AutoScaleDimensions = new System.Drawing.SizeF(6.0f, 13.0f);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            ClientSize = new System.Drawing.Size(629, 213);
            Controls.Add(PictureBox2);
            Controls.Add(Info);
            Controls.Add(CustomGroupBox2);
            Controls.Add(Des_selection);
            Controls.Add(TB_des_rng);
            Controls.Add(Label2);
            Controls.Add(Src_selection);
            Controls.Add(Btn_OK);
            Controls.Add(ComboBox1);
            Controls.Add(TB_src_rng);
            Controls.Add(Label1);
            Controls.Add(Btn_Cancel);
            FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "Form34_PictureBasedDropdownList";
            Text = "Picture Based Drop-down List";
            ((System.ComponentModel.ISupportInitialize)Src_selection).EndInit();
            ((System.ComponentModel.ISupportInitialize)Des_selection).EndInit();
            ((System.ComponentModel.ISupportInitialize)Info).EndInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox2).EndInit();
            Load += new EventHandler(Form34_PictureBasedDropdownList_Load);
            KeyDown += new System.Windows.Forms.KeyEventHandler(form_enter);
            Closing += new System.ComponentModel.CancelEventHandler(Form34_PictureBasedDropdownList_Closing);
            Disposed += new EventHandler(Form34_PictureBasedDropdownList_Disposed);
            Shown += new EventHandler(Form34_PictureBasedDropdownList_Shown);
            ResumeLayout(false);
            PerformLayout();

        }
        internal System.Windows.Forms.PictureBox Src_selection;
        internal System.Windows.Forms.Button Btn_OK;
        internal System.Windows.Forms.ComboBox ComboBox1;
        internal System.Windows.Forms.TextBox TB_src_rng;
        internal System.Windows.Forms.Label Label1;
        internal System.Windows.Forms.Button Btn_Cancel;
        internal System.Windows.Forms.PictureBox Des_selection;
        internal System.Windows.Forms.TextBox TB_des_rng;
        internal System.Windows.Forms.Label Label2;
        internal CustomGroupBox CustomGroupBox2;
        internal System.Windows.Forms.PictureBox Info;
        internal System.Windows.Forms.PictureBox PictureBox2;
        internal System.Windows.Forms.ToolTip ToolTip1;
    }
}