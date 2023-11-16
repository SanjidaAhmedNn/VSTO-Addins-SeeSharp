using System;

namespace VSTO_Addins
{
    [Microsoft.VisualBasic.CompilerServices.DesignerGenerated()]
    public partial class FormProgressBar : System.Windows.Forms.Form
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
            Label1 = new System.Windows.Forms.Label();
            ProgressBar1 = new System.Windows.Forms.ProgressBar();
            SuspendLayout();
            // 
            // Label1
            // 
            Label1.AutoSize = true;
            Label1.Location = new System.Drawing.Point(12, 20);
            Label1.Name = "Label1";
            Label1.Size = new System.Drawing.Size(39, 13);
            Label1.TabIndex = 1;
            Label1.Text = "Label1";
            Label1.Visible = false;
            // 
            // ProgressBar1
            // 
            ProgressBar1.Location = new System.Drawing.Point(12, 52);
            ProgressBar1.Name = "ProgressBar1";
            ProgressBar1.Size = new System.Drawing.Size(459, 23);
            ProgressBar1.TabIndex = 0;
            // 
            // FormProgressBar
            // 
            AutoScaleDimensions = new System.Drawing.SizeF(6.0f, 13.0f);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            ClientSize = new System.Drawing.Size(483, 87);
            Controls.Add(Label1);
            Controls.Add(ProgressBar1);
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "FormProgressBar";
            Text = "SOFTEKO";
            Load += new EventHandler(Pro_Load);
            ResumeLayout(false);
            PerformLayout();

        }
        internal System.Windows.Forms.Label Label1;
        internal System.Windows.Forms.ProgressBar ProgressBar1;
    }
}