using System;

namespace VSTO_Addins
{
    [Microsoft.VisualBasic.CompilerServices.DesignerGenerated()]
    public partial class Form32_ExtendDropDownList : System.Windows.Forms.Form
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
            var resources = new System.ComponentModel.ComponentResourceManager(typeof(Form32_ExtendDropDownList));
            Info = new System.Windows.Forms.PictureBox();
            Source_selection = new System.Windows.Forms.PictureBox();
            Source_selection.Click += new EventHandler(Selection_source_Click);
            Btn_OK = new System.Windows.Forms.Button();
            Btn_OK.Click += new EventHandler(Btn_OK_Click);
            Btn_Cancel = new System.Windows.Forms.Button();
            Btn_Cancel.Click += new EventHandler(Btn_Cancel_Click);
            ComboBox2 = new System.Windows.Forms.ComboBox();
            TB_src_rng = new System.Windows.Forms.TextBox();
            TB_src_rng.TextChanged += new EventHandler(TB_src_rng_TextChanged);
            Label1 = new System.Windows.Forms.Label();
            Dest_selection = new System.Windows.Forms.PictureBox();
            Dest_selection.Click += new EventHandler(Dest_selection_Click);
            TB_des_rng = new System.Windows.Forms.TextBox();
            TB_des_rng.TextChanged += new EventHandler(TB_des_rng_TextChanged);
            Label2 = new System.Windows.Forms.Label();
            ToolTip1 = new System.Windows.Forms.ToolTip(components);
            L_warning = new CustomLabel();
            ((System.ComponentModel.ISupportInitialize)Info).BeginInit();
            ((System.ComponentModel.ISupportInitialize)Source_selection).BeginInit();
            ((System.ComponentModel.ISupportInitialize)Dest_selection).BeginInit();
            SuspendLayout();
            // 
            // Info
            // 
            Info.Image = (System.Drawing.Image)resources.GetObject("Info.Image");
            Info.Location = new System.Drawing.Point(249, 16);
            Info.Name = "Info";
            Info.Size = new System.Drawing.Size(20, 20);
            Info.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            Info.TabIndex = 390;
            Info.TabStop = false;
            // 
            // Source_selection
            // 
            Source_selection.BackColor = System.Drawing.Color.White;
            Source_selection.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            Source_selection.Image = (System.Drawing.Image)resources.GetObject("Source_selection.Image");
            Source_selection.Location = new System.Drawing.Point(309, 42);
            Source_selection.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            Source_selection.Name = "Source_selection";
            Source_selection.Size = new System.Drawing.Size(24, 25);
            Source_selection.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            Source_selection.TabIndex = 389;
            Source_selection.TabStop = false;
            // 
            // Btn_OK
            // 
            Btn_OK.BackColor = System.Drawing.Color.White;
            Btn_OK.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            Btn_OK.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            Btn_OK.Location = new System.Drawing.Point(195, 188);
            Btn_OK.Name = "Btn_OK";
            Btn_OK.Size = new System.Drawing.Size(62, 26);
            Btn_OK.TabIndex = 388;
            Btn_OK.Text = "OK";
            Btn_OK.UseVisualStyleBackColor = false;
            // 
            // Btn_Cancel
            // 
            Btn_Cancel.BackColor = System.Drawing.Color.White;
            Btn_Cancel.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            Btn_Cancel.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            Btn_Cancel.Location = new System.Drawing.Point(271, 188);
            Btn_Cancel.Name = "Btn_Cancel";
            Btn_Cancel.Size = new System.Drawing.Size(62, 26);
            Btn_Cancel.TabIndex = 387;
            Btn_Cancel.Text = "Cancel";
            Btn_Cancel.UseVisualStyleBackColor = false;
            // 
            // ComboBox2
            // 
            ComboBox2.Font = new System.Drawing.Font("Segoe UI", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            ComboBox2.FormattingEnabled = true;
            ComboBox2.Location = new System.Drawing.Point(15, 189);
            ComboBox2.Name = "ComboBox2";
            ComboBox2.Size = new System.Drawing.Size(96, 25);
            ComboBox2.TabIndex = 386;
            ComboBox2.Text = "Softeko";
            // 
            // TB_src_rng
            // 
            TB_src_rng.Font = new System.Drawing.Font("Segoe UI", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            TB_src_rng.Location = new System.Drawing.Point(15, 42);
            TB_src_rng.Name = "TB_src_rng";
            TB_src_rng.Size = new System.Drawing.Size(318, 25);
            TB_src_rng.TabIndex = 384;
            // 
            // Label1
            // 
            Label1.AutoSize = true;
            Label1.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            Label1.Location = new System.Drawing.Point(15, 15);
            Label1.Name = "Label1";
            Label1.Size = new System.Drawing.Size(237, 17);
            Label1.TabIndex = 383;
            Label1.Text = "Select dynamic drop-down list range :";
            // 
            // Dest_selection
            // 
            Dest_selection.BackColor = System.Drawing.Color.White;
            Dest_selection.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            Dest_selection.Image = (System.Drawing.Image)resources.GetObject("Dest_selection.Image");
            Dest_selection.Location = new System.Drawing.Point(309, 107);
            Dest_selection.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            Dest_selection.Name = "Dest_selection";
            Dest_selection.Size = new System.Drawing.Size(24, 25);
            Dest_selection.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            Dest_selection.TabIndex = 393;
            Dest_selection.TabStop = false;
            // 
            // TB_des_rng
            // 
            TB_des_rng.Font = new System.Drawing.Font("Segoe UI", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            TB_des_rng.Location = new System.Drawing.Point(15, 107);
            TB_des_rng.Name = "TB_des_rng";
            TB_des_rng.Size = new System.Drawing.Size(318, 25);
            TB_des_rng.TabIndex = 392;
            // 
            // Label2
            // 
            Label2.AutoSize = true;
            Label2.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            Label2.Location = new System.Drawing.Point(15, 78);
            Label2.Name = "Label2";
            Label2.Size = new System.Drawing.Size(325, 17);
            Label2.TabIndex = 391;
            Label2.Text = "Select the expanded dynamic drop-down list range :";
            // 
            // L_warning
            // 
            L_warning.BorderColor = System.Drawing.Color.DimGray;
            L_warning.BorderWidth = 1;
            L_warning.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            L_warning.Location = new System.Drawing.Point(29, 147);
            L_warning.Name = "L_warning";
            L_warning.Size = new System.Drawing.Size(286, 23);
            L_warning.TabIndex = 396;
            L_warning.Text = " These two ranges must intersect each other";
            L_warning.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // Form32_ExtendDropDownList
            // 
            AutoScaleDimensions = new System.Drawing.SizeF(6.0f, 13.0f);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            ClientSize = new System.Drawing.Size(353, 234);
            Controls.Add(L_warning);
            Controls.Add(Dest_selection);
            Controls.Add(TB_des_rng);
            Controls.Add(Label2);
            Controls.Add(Info);
            Controls.Add(Source_selection);
            Controls.Add(Btn_OK);
            Controls.Add(Btn_Cancel);
            Controls.Add(ComboBox2);
            Controls.Add(TB_src_rng);
            Controls.Add(Label1);
            KeyPreview = true;
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "Form32_ExtendDropDownList";
            Text = "Extend Dynamic Drop-down List";
            ((System.ComponentModel.ISupportInitialize)Info).EndInit();
            ((System.ComponentModel.ISupportInitialize)Source_selection).EndInit();
            ((System.ComponentModel.ISupportInitialize)Dest_selection).EndInit();
            KeyDown += new System.Windows.Forms.KeyEventHandler(Form1_KeyDown);
            Load += new EventHandler(Form32_ExtendDropDownList_Load);
            Disposed += new EventHandler(Form32_ExtendDropDownList_Disposed);
            Closing += new System.ComponentModel.CancelEventHandler(Form32_ExtendDropDownList_Closing);
            Shown += new EventHandler(Form32_ExtendDropDownList_Shown);
            ResumeLayout(false);
            PerformLayout();

        }

        internal System.Windows.Forms.PictureBox Info;
        internal System.Windows.Forms.PictureBox Source_selection;
        internal System.Windows.Forms.Button Btn_OK;
        internal System.Windows.Forms.Button Btn_Cancel;
        internal System.Windows.Forms.ComboBox ComboBox2;
        internal System.Windows.Forms.TextBox TB_src_rng;
        internal System.Windows.Forms.Label Label1;
        internal System.Windows.Forms.PictureBox Dest_selection;
        internal System.Windows.Forms.TextBox TB_des_rng;
        internal System.Windows.Forms.Label Label2;
        internal CustomLabel L_warning;
        internal System.Windows.Forms.ToolTip ToolTip1;
    }
}