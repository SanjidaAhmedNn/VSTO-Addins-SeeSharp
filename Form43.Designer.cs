using System;

namespace VSTO_Addins
{
    [Microsoft.VisualBasic.CompilerServices.DesignerGenerated()]
    public partial class Form43 : System.Windows.Forms.Form
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
            CheckBox1 = new System.Windows.Forms.CheckBox();
            CheckBox1.CheckedChanged += new EventHandler(CheckBox1_CheckedChanged);
            Button1 = new System.Windows.Forms.Button();
            Button1.Click += new EventHandler(Button1_Click);
            Button2 = new System.Windows.Forms.Button();
            Button2.Click += new EventHandler(Button2_Click);
            SuspendLayout();
            // 
            // Label1
            // 
            Label1.BackColor = System.Drawing.Color.FromArgb(224, 224, 224);
            Label1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            Label1.FlatStyle = System.Windows.Forms.FlatStyle.System;
            Label1.Font = new System.Drawing.Font("Segoe UI", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            Label1.Location = new System.Drawing.Point(15, 15);
            Label1.Name = "Label1";
            Label1.Size = new System.Drawing.Size(317, 42);
            Label1.TabIndex = 0;
            Label1.Text = "You have not selected any colors. Do you want to change the colors?";
            Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // CheckBox1
            // 
            CheckBox1.AutoSize = true;
            CheckBox1.Font = new System.Drawing.Font("Segoe UI", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            CheckBox1.Location = new System.Drawing.Point(18, 167);
            CheckBox1.Name = "CheckBox1";
            CheckBox1.Size = new System.Drawing.Size(253, 21);
            CheckBox1.TabIndex = 3;
            CheckBox1.Text = "Don't show this for this current session";
            CheckBox1.UseVisualStyleBackColor = true;
            // 
            // Button1
            // 
            Button1.FlatStyle = System.Windows.Forms.FlatStyle.System;
            Button1.Font = new System.Drawing.Font("Segoe UI", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            Button1.Location = new System.Drawing.Point(18, 67);
            Button1.Name = "Button1";
            Button1.Size = new System.Drawing.Size(313, 39);
            Button1.TabIndex = 4;
            Button1.Text = "Yes, help me to change the colors";
            Button1.UseVisualStyleBackColor = true;
            // 
            // Button2
            // 
            Button2.FlatStyle = System.Windows.Forms.FlatStyle.System;
            Button2.Font = new System.Drawing.Font("Segoe UI", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            Button2.Location = new System.Drawing.Point(18, 117);
            Button2.Name = "Button2";
            Button2.Size = new System.Drawing.Size(313, 40);
            Button2.TabIndex = 5;
            Button2.Text = "No, I don't want any formatting";
            Button2.UseVisualStyleBackColor = true;
            // 
            // Form43
            // 
            AutoScaleDimensions = new System.Drawing.SizeF(6.0f, 13.0f);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            ClientSize = new System.Drawing.Size(350, 200);
            Controls.Add(Button2);
            Controls.Add(Button1);
            Controls.Add(CheckBox1);
            Controls.Add(Label1);
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "Form43";
            Text = "Softeko for Excel";
            Load += new EventHandler(Form43_Load);
            Closing += new System.ComponentModel.CancelEventHandler(Form43_Closing);
            Disposed += new EventHandler(Form43_Disposed);
            ResumeLayout(false);
            PerformLayout();

        }

        internal System.Windows.Forms.Label Label1;
        internal System.Windows.Forms.CheckBox CheckBox1;
        internal System.Windows.Forms.Button Button1;
        internal System.Windows.Forms.Button Button2;
    }
}