using System;

namespace VSTO_Addins
{
    [Microsoft.VisualBasic.CompilerServices.DesignerGenerated()]
    public partial class Form36 : System.Windows.Forms.Form
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
            var resources = new System.ComponentModel.ComponentResourceManager(typeof(Form36));
            txtSearch = new System.Windows.Forms.TextBox();
            txtSearch.TextChanged += new EventHandler(txtSearch_TextChanged);
            DataGridView1 = new System.Windows.Forms.DataGridView();
            DataGridView1.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(DataGridView1_CellClick);
            DataGridView1.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(DataGridView1_CellContentClick);
            PB_Search = new System.Windows.Forms.PictureBox();
            PictureBox2 = new System.Windows.Forms.PictureBox();
            PictureBox3 = new System.Windows.Forms.PictureBox();
            PictureBox3.Click += new EventHandler(PictureBox3_Click);
            PictureBox4 = new System.Windows.Forms.PictureBox();
            PictureBox4.Click += new EventHandler(PictureBox4_Click);
            Panel1 = new System.Windows.Forms.Panel();
            Panel1.Paint += new System.Windows.Forms.PaintEventHandler(Panel1_Paint);
            ((System.ComponentModel.ISupportInitialize)DataGridView1).BeginInit();
            ((System.ComponentModel.ISupportInitialize)PB_Search).BeginInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox2).BeginInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox3).BeginInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox4).BeginInit();
            Panel1.SuspendLayout();
            SuspendLayout();
            // 
            // txtSearch
            // 
            resources.ApplyResources(txtSearch, "txtSearch");
            txtSearch.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            txtSearch.Name = "txtSearch";
            // 
            // DataGridView1
            // 
            resources.ApplyResources(DataGridView1, "DataGridView1");
            DataGridView1.AllowUserToAddRows = false;
            DataGridView1.AllowUserToDeleteRows = false;
            DataGridView1.AllowUserToResizeColumns = false;
            DataGridView1.AllowUserToResizeRows = false;
            DataGridView1.BackgroundColor = System.Drawing.SystemColors.ButtonHighlight;
            DataGridView1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            DataGridView1.Name = "DataGridView1";
            DataGridView1.ReadOnly = true;
            DataGridView1.RowHeadersVisible = false;
            DataGridView1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            // 
            // PB_Search
            // 
            resources.ApplyResources(PB_Search, "PB_Search");
            PB_Search.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            PB_Search.Name = "PB_Search";
            PB_Search.TabStop = false;
            // 
            // PictureBox2
            // 
            resources.ApplyResources(PictureBox2, "PictureBox2");
            PictureBox2.BackColor = System.Drawing.Color.WhiteSmoke;
            PictureBox2.Name = "PictureBox2";
            PictureBox2.TabStop = false;
            // 
            // PictureBox3
            // 
            resources.ApplyResources(PictureBox3, "PictureBox3");
            PictureBox3.BackColor = System.Drawing.Color.WhiteSmoke;
            PictureBox3.Name = "PictureBox3";
            PictureBox3.TabStop = false;
            // 
            // PictureBox4
            // 
            resources.ApplyResources(PictureBox4, "PictureBox4");
            PictureBox4.BackColor = System.Drawing.Color.WhiteSmoke;
            PictureBox4.Name = "PictureBox4";
            PictureBox4.TabStop = false;
            // 
            // Panel1
            // 
            resources.ApplyResources(Panel1, "Panel1");
            Panel1.BackColor = System.Drawing.Color.WhiteSmoke;
            Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            Panel1.Controls.Add(PictureBox3);
            Panel1.Controls.Add(PictureBox4);
            Panel1.Controls.Add(PictureBox2);
            Panel1.Name = "Panel1";
            // 
            // Form36
            // 
            resources.ApplyResources(this, "$this");
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            Controls.Add(Panel1);
            Controls.Add(PB_Search);
            Controls.Add(DataGridView1);
            Controls.Add(txtSearch);
            FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "Form36";
            ShowInTaskbar = false;
            ((System.ComponentModel.ISupportInitialize)DataGridView1).EndInit();
            ((System.ComponentModel.ISupportInitialize)PB_Search).EndInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox2).EndInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox3).EndInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox4).EndInit();
            Panel1.ResumeLayout(false);
            Load += new EventHandler(Form36_Load);
            Activated += new EventHandler(Form36_Activated);
            Shown += new EventHandler(Form36_Shown);
            ResumeLayout(false);
            PerformLayout();

        }
        internal System.Windows.Forms.TextBox txtSearch;
        internal System.Windows.Forms.DataGridView DataGridView1;
        internal System.Windows.Forms.PictureBox PB_Search;
        internal System.Windows.Forms.PictureBox PictureBox2;
        internal System.Windows.Forms.PictureBox PictureBox3;
        internal System.Windows.Forms.PictureBox PictureBox4;
        internal System.Windows.Forms.Panel Panel1;
    }
}