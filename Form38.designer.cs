using System;

namespace VSTO_Addins
{
    [Microsoft.VisualBasic.CompilerServices.DesignerGenerated()]
    public partial class Form38 : System.Windows.Forms.Form
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
            var resources = new System.ComponentModel.ComponentResourceManager(typeof(Form38));
            PictureBox1 = new System.Windows.Forms.PictureBox();
            DataGridView1 = new System.Windows.Forms.DataGridView();
            DataGridView1.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(DataGridView1_CellClick);
            DataGridView1.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(DataGridView1_CellValueChanged);
            DataGridView1.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(DataGridView1_CellContentClick);
            txtSearch = new System.Windows.Forms.TextBox();
            txtSearch.TextChanged += new EventHandler(txtSearch_TextChanged);
            PictureBox3 = new System.Windows.Forms.PictureBox();
            PictureBox3.Click += new EventHandler(PictureBox3_Click);
            PictureBox4 = new System.Windows.Forms.PictureBox();
            PictureBox4.Click += new EventHandler(PictureBox4_Click);
            PictureBox2 = new System.Windows.Forms.PictureBox();
            Panel1 = new System.Windows.Forms.Panel();
            Panel1.Paint += new System.Windows.Forms.PaintEventHandler(Panel1_Paint);
            ((System.ComponentModel.ISupportInitialize)PictureBox1).BeginInit();
            ((System.ComponentModel.ISupportInitialize)DataGridView1).BeginInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox3).BeginInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox4).BeginInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox2).BeginInit();
            Panel1.SuspendLayout();
            SuspendLayout();
            // 
            // PictureBox1
            // 
            PictureBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            PictureBox1.Image = (System.Drawing.Image)resources.GetObject("PictureBox1.Image");
            PictureBox1.Location = new System.Drawing.Point(1, 34);
            PictureBox1.Name = "PictureBox1";
            PictureBox1.Size = new System.Drawing.Size(28, 24);
            PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            PictureBox1.TabIndex = 408;
            PictureBox1.TabStop = false;
            // 
            // DataGridView1
            // 
            DataGridView1.AllowUserToAddRows = false;
            DataGridView1.AllowUserToDeleteRows = false;
            DataGridView1.AllowUserToResizeColumns = false;
            DataGridView1.AllowUserToResizeRows = false;
            DataGridView1.BackgroundColor = System.Drawing.SystemColors.ButtonHighlight;
            DataGridView1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            DataGridView1.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter;
            DataGridView1.Location = new System.Drawing.Point(1, 59);
            DataGridView1.Name = "DataGridView1";
            DataGridView1.ReadOnly = true;
            DataGridView1.RowHeadersVisible = false;
            DataGridView1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            DataGridView1.Size = new System.Drawing.Size(238, 240);
            DataGridView1.TabIndex = 407;
            // 
            // txtSearch
            // 
            txtSearch.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            txtSearch.Font = new System.Drawing.Font("Segoe UI", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            txtSearch.Location = new System.Drawing.Point(30, 34);
            txtSearch.Name = "txtSearch";
            txtSearch.Size = new System.Drawing.Size(209, 25);
            txtSearch.TabIndex = 406;
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
            // PictureBox4
            // 
            PictureBox4.BackColor = System.Drawing.Color.WhiteSmoke;
            PictureBox4.Image = (System.Drawing.Image)resources.GetObject("PictureBox4.Image");
            PictureBox4.Location = new System.Drawing.Point(5, 4);
            PictureBox4.Name = "PictureBox4";
            PictureBox4.Size = new System.Drawing.Size(25, 25);
            PictureBox4.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            PictureBox4.TabIndex = 404;
            PictureBox4.TabStop = false;
            // 
            // PictureBox2
            // 
            PictureBox2.BackColor = System.Drawing.Color.WhiteSmoke;
            PictureBox2.Image = (System.Drawing.Image)resources.GetObject("PictureBox2.Image");
            PictureBox2.Location = new System.Drawing.Point(36, 7);
            PictureBox2.Name = "PictureBox2";
            PictureBox2.Size = new System.Drawing.Size(20, 20);
            PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            PictureBox2.TabIndex = 402;
            PictureBox2.TabStop = false;
            // 
            // Panel1
            // 
            Panel1.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            Panel1.BackColor = System.Drawing.Color.WhiteSmoke;
            Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            Panel1.Controls.Add(PictureBox3);
            Panel1.Controls.Add(PictureBox4);
            Panel1.Controls.Add(PictureBox2);
            Panel1.Location = new System.Drawing.Point(1, 0);
            Panel1.Name = "Panel1";
            Panel1.Size = new System.Drawing.Size(238, 34);
            Panel1.TabIndex = 409;
            // 
            // Form38
            // 
            AutoScaleDimensions = new System.Drawing.SizeF(6.0f, 13.0f);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            ClientSize = new System.Drawing.Size(239, 300);
            Controls.Add(Panel1);
            Controls.Add(PictureBox1);
            Controls.Add(DataGridView1);
            Controls.Add(txtSearch);
            FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            Name = "Form38";
            ShowInTaskbar = false;
            Text = "Form38";
            ((System.ComponentModel.ISupportInitialize)PictureBox1).EndInit();
            ((System.ComponentModel.ISupportInitialize)DataGridView1).EndInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox3).EndInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox4).EndInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox2).EndInit();
            Panel1.ResumeLayout(false);
            Load += new EventHandler(Form38_Load);
            Activated += new EventHandler(Form38_Activated);
            Shown += new EventHandler(Form38_Shown);
            ResumeLayout(false);
            PerformLayout();

        }

        internal System.Windows.Forms.PictureBox PictureBox1;
        internal System.Windows.Forms.DataGridView DataGridView1;
        internal System.Windows.Forms.TextBox txtSearch;
        internal System.Windows.Forms.PictureBox PictureBox3;
        internal System.Windows.Forms.PictureBox PictureBox4;
        internal System.Windows.Forms.PictureBox PictureBox2;
        internal System.Windows.Forms.Panel Panel1;
    }
}