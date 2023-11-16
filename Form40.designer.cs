using System;

namespace VSTO_Addins
{
    [Microsoft.VisualBasic.CompilerServices.DesignerGenerated()]
    public partial class Form40 : System.Windows.Forms.Form
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
            var resources = new System.ComponentModel.ComponentResourceManager(typeof(Form40));
            PictureBox1 = new System.Windows.Forms.PictureBox();
            txtSearch = new System.Windows.Forms.TextBox();
            txtSearch.TextChanged += new EventHandler(txtSearch_TextChanged);
            Panel1 = new System.Windows.Forms.Panel();
            Panel1.Paint += new System.Windows.Forms.PaintEventHandler(Panel1_Paint);
            PictureBox3 = new System.Windows.Forms.PictureBox();
            PictureBox3.Click += new EventHandler(PictureBox3_Click);
            PictureBox2 = new System.Windows.Forms.PictureBox();
            ListBox1 = new System.Windows.Forms.ListBox();
            ListBox1.SelectedIndexChanged += new EventHandler(ListBox1_SelectedIndexChanged);
            ListBox1.KeyDown += new System.Windows.Forms.KeyEventHandler(listbox_enter);
            ((System.ComponentModel.ISupportInitialize)PictureBox1).BeginInit();
            Panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)PictureBox3).BeginInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox2).BeginInit();
            SuspendLayout();
            // 
            // PictureBox1
            // 
            PictureBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            PictureBox1.Image = (System.Drawing.Image)resources.GetObject("PictureBox1.Image");
            PictureBox1.Location = new System.Drawing.Point(0, 33);
            PictureBox1.Name = "PictureBox1";
            PictureBox1.Size = new System.Drawing.Size(28, 24);
            PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            PictureBox1.TabIndex = 412;
            PictureBox1.TabStop = false;
            // 
            // txtSearch
            // 
            txtSearch.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            txtSearch.Font = new System.Drawing.Font("Segoe UI", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            txtSearch.Location = new System.Drawing.Point(30, 33);
            txtSearch.Name = "txtSearch";
            txtSearch.Size = new System.Drawing.Size(209, 25);
            txtSearch.TabIndex = 410;
            // 
            // Panel1
            // 
            Panel1.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            Panel1.BackColor = System.Drawing.Color.WhiteSmoke;
            Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            Panel1.Controls.Add(PictureBox3);
            Panel1.Controls.Add(PictureBox2);
            Panel1.Location = new System.Drawing.Point(1, 2);
            Panel1.Name = "Panel1";
            Panel1.Size = new System.Drawing.Size(238, 34);
            Panel1.TabIndex = 413;
            // 
            // PictureBox3
            // 
            PictureBox3.BackColor = System.Drawing.Color.WhiteSmoke;
            PictureBox3.Image = (System.Drawing.Image)resources.GetObject("PictureBox3.Image");
            PictureBox3.Location = new System.Drawing.Point(208, 4);
            PictureBox3.Name = "PictureBox3";
            PictureBox3.Size = new System.Drawing.Size(25, 25);
            PictureBox3.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            PictureBox3.TabIndex = 403;
            PictureBox3.TabStop = false;
            // 
            // PictureBox2
            // 
            PictureBox2.BackColor = System.Drawing.Color.WhiteSmoke;
            PictureBox2.Image = (System.Drawing.Image)resources.GetObject("PictureBox2.Image");
            PictureBox2.Location = new System.Drawing.Point(6, 6);
            PictureBox2.Name = "PictureBox2";
            PictureBox2.Size = new System.Drawing.Size(20, 20);
            PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            PictureBox2.TabIndex = 402;
            PictureBox2.TabStop = false;
            // 
            // ListBox1
            // 
            ListBox1.Font = new System.Drawing.Font("Segoe UI", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            ListBox1.FormattingEnabled = true;
            ListBox1.ItemHeight = 17;
            ListBox1.Location = new System.Drawing.Point(0, 58);
            ListBox1.Name = "ListBox1";
            ListBox1.Size = new System.Drawing.Size(239, 242);
            ListBox1.TabIndex = 414;
            // 
            // Form40
            // 
            AutoScaleDimensions = new System.Drawing.SizeF(6.0f, 13.0f);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            ClientSize = new System.Drawing.Size(241, 301);
            Controls.Add(Panel1);
            Controls.Add(ListBox1);
            Controls.Add(PictureBox1);
            Controls.Add(txtSearch);
            FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "Form40";
            ShowInTaskbar = false;
            Text = "Form40";
            ((System.ComponentModel.ISupportInitialize)PictureBox1).EndInit();
            Panel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)PictureBox3).EndInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox2).EndInit();
            Load += new EventHandler(Form40_Load);
            Activated += new EventHandler(Form40_Activated);
            KeyDown += new System.Windows.Forms.KeyEventHandler(form_enter);
            Shown += new EventHandler(Form40_Shown);
            ResumeLayout(false);
            PerformLayout();

        }

        internal System.Windows.Forms.PictureBox PictureBox1;
        internal System.Windows.Forms.TextBox txtSearch;
        internal System.Windows.Forms.Panel Panel1;
        internal System.Windows.Forms.PictureBox PictureBox3;
        internal System.Windows.Forms.PictureBox PictureBox2;
        internal System.Windows.Forms.ListBox ListBox1;
    }
}