using System;

namespace VSTO_Addins
{
    [Microsoft.VisualBasic.CompilerServices.DesignerGenerated()]
    public partial class Form31_2_updated_selection : System.Windows.Forms.Form
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
            DataGridView1 = new System.Windows.Forms.DataGridView();
            DataGridView1.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(DataGridView1_CellContentClick);
            DataGridView1.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(dataGridView1_CellValueChanged);
            DataGridView1.DataBindingComplete += new System.Windows.Forms.DataGridViewBindingCompleteEventHandler(dataGridView1_DataBindingComplete);
            DataGridView1.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(dataGridView1_CellClick);
            Btn_OK = new System.Windows.Forms.Button();
            Btn_OK.Click += new EventHandler(Btn_OK_Click);
            Btn_Cancel = new System.Windows.Forms.Button();
            Btn_Cancel.Click += new EventHandler(Btn_Cancel_Click);
            ComboBox2 = new System.Windows.Forms.ComboBox();
            ((System.ComponentModel.ISupportInitialize)DataGridView1).BeginInit();
            SuspendLayout();
            // 
            // DataGridView1
            // 
            DataGridView1.AllowUserToAddRows = false;
            DataGridView1.AllowUserToDeleteRows = false;
            DataGridView1.AllowUserToResizeColumns = false;
            DataGridView1.AllowUserToResizeRows = false;
            DataGridView1.BackgroundColor = System.Drawing.Color.White;
            DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            DataGridView1.Location = new System.Drawing.Point(12, 12);
            DataGridView1.Name = "DataGridView1";
            DataGridView1.RowHeadersVisible = false;
            DataGridView1.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing;
            DataGridView1.Size = new System.Drawing.Size(379, 155);
            DataGridView1.TabIndex = 0;
            // 
            // Btn_OK
            // 
            Btn_OK.BackColor = System.Drawing.Color.White;
            Btn_OK.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            Btn_OK.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            Btn_OK.Location = new System.Drawing.Point(251, 184);
            Btn_OK.Name = "Btn_OK";
            Btn_OK.Size = new System.Drawing.Size(62, 26);
            Btn_OK.TabIndex = 368;
            Btn_OK.Text = "OK";
            Btn_OK.UseVisualStyleBackColor = false;
            // 
            // Btn_Cancel
            // 
            Btn_Cancel.BackColor = System.Drawing.Color.White;
            Btn_Cancel.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            Btn_Cancel.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            Btn_Cancel.Location = new System.Drawing.Point(329, 184);
            Btn_Cancel.Name = "Btn_Cancel";
            Btn_Cancel.Size = new System.Drawing.Size(62, 26);
            Btn_Cancel.TabIndex = 367;
            Btn_Cancel.Text = "Cancel";
            Btn_Cancel.UseVisualStyleBackColor = false;
            // 
            // ComboBox2
            // 
            ComboBox2.Font = new System.Drawing.Font("Segoe UI", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            ComboBox2.FormattingEnabled = true;
            ComboBox2.Location = new System.Drawing.Point(12, 184);
            ComboBox2.Name = "ComboBox2";
            ComboBox2.Size = new System.Drawing.Size(111, 25);
            ComboBox2.TabIndex = 369;
            ComboBox2.Text = "Softeko";
            // 
            // Form31_2_updated_selection
            // 
            AutoScaleDimensions = new System.Drawing.SizeF(6.0f, 13.0f);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            ClientSize = new System.Drawing.Size(403, 227);
            Controls.Add(ComboBox2);
            Controls.Add(Btn_OK);
            Controls.Add(Btn_Cancel);
            Controls.Add(DataGridView1);
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "Form31_2_updated_selection";
            Text = "Updated Dynamic Drop-down List";
            ((System.ComponentModel.ISupportInitialize)DataGridView1).EndInit();
            Load += new EventHandler(Form31_2_updated_selection_Load);
            ResumeLayout(false);

        }

        internal System.Windows.Forms.DataGridView DataGridView1;
        internal System.Windows.Forms.Button Btn_OK;
        internal System.Windows.Forms.Button Btn_Cancel;
        internal System.Windows.Forms.ComboBox ComboBox2;
    }
}