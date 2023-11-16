using System;

namespace VSTO_Addins
{
    [Microsoft.VisualBasic.CompilerServices.DesignerGenerated()]
    public partial class Form16PasteintoVisibleRange : System.Windows.Forms.Form
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
            var resources = new System.ComponentModel.ComponentResourceManager(typeof(Form16PasteintoVisibleRange));
            CB_keepFormat = new System.Windows.Forms.CheckBox();
            btnOK = new System.Windows.Forms.Button();
            btnOK.Click += new EventHandler(btnOK_Click);
            ComboBox1 = new System.Windows.Forms.ComboBox();
            CB_copyWs = new System.Windows.Forms.CheckBox();
            Label1 = new System.Windows.Forms.Label();
            btnCancel = new System.Windows.Forms.Button();
            btnCancel.Click += new EventHandler(btnCancel_Click);
            txtSourceRange = new System.Windows.Forms.TextBox();
            txtSourceRange.TextChanged += new EventHandler(txtSourceRange_TextChanged);
            txtSourceRange.GotFocus += new EventHandler(txtSourceRange_GotFocus);
            Selection = new System.Windows.Forms.PictureBox();
            Selection.Click += new EventHandler(Selection_Click);
            AutoSelection = new System.Windows.Forms.PictureBox();
            AutoSelection.Click += new EventHandler(AutoSelection_Click);
            CustomGroupBox5 = new CustomGroupBox();
            CustomPanel1 = new CustomPanel();
            destinationSelection = new System.Windows.Forms.PictureBox();
            destinationSelection.Click += new EventHandler(destinationSelection_Click);
            txtDestRange = new System.Windows.Forms.TextBox();
            txtDestRange.TextChanged += new EventHandler(txtDestRange_TextChanged);
            txtDestRange.GotFocus += new EventHandler(txtDestRange_GotFocus);
            Label2 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)Selection).BeginInit();
            ((System.ComponentModel.ISupportInitialize)AutoSelection).BeginInit();
            CustomGroupBox5.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)destinationSelection).BeginInit();
            SuspendLayout();
            // 
            // CB_keepFormat
            // 
            CB_keepFormat.AutoSize = true;
            CB_keepFormat.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            CB_keepFormat.Location = new System.Drawing.Point(19, 77);
            CB_keepFormat.Name = "CB_keepFormat";
            CB_keepFormat.Size = new System.Drawing.Size(122, 21);
            CB_keepFormat.TabIndex = 184;
            CB_keepFormat.Text = "Keep formatting";
            CB_keepFormat.UseVisualStyleBackColor = true;
            // 
            // btnOK
            // 
            btnOK.BackColor = System.Drawing.Color.White;
            btnOK.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            btnOK.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            btnOK.Location = new System.Drawing.Point(300, 316);
            btnOK.Name = "btnOK";
            btnOK.Size = new System.Drawing.Size(62, 26);
            btnOK.TabIndex = 190;
            btnOK.Text = "OK";
            btnOK.UseVisualStyleBackColor = false;
            // 
            // ComboBox1
            // 
            ComboBox1.Font = new System.Drawing.Font("Segoe UI", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            ComboBox1.FormattingEnabled = true;
            ComboBox1.Location = new System.Drawing.Point(12, 316);
            ComboBox1.Name = "ComboBox1";
            ComboBox1.Size = new System.Drawing.Size(90, 25);
            ComboBox1.TabIndex = 186;
            ComboBox1.Text = "SOFTEKO";
            // 
            // CB_copyWs
            // 
            CB_copyWs.AutoSize = true;
            CB_copyWs.Font = new System.Drawing.Font("Segoe UI", 9.38f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            CB_copyWs.Location = new System.Drawing.Point(13, 286);
            CB_copyWs.Name = "CB_copyWs";
            CB_copyWs.Size = new System.Drawing.Size(257, 21);
            CB_copyWs.TabIndex = 185;
            CB_copyWs.Text = "Create a copy of the original worksheet";
            CB_copyWs.UseVisualStyleBackColor = true;
            // 
            // Label1
            // 
            Label1.AutoSize = true;
            Label1.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            Label1.Location = new System.Drawing.Point(15, 17);
            Label1.Name = "Label1";
            Label1.Size = new System.Drawing.Size(123, 17);
            Label1.TabIndex = 182;
            Label1.Text = "Data to be copied :";
            // 
            // btnCancel
            // 
            btnCancel.BackColor = System.Drawing.Color.White;
            btnCancel.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            btnCancel.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            btnCancel.Location = new System.Drawing.Point(374, 316);
            btnCancel.Name = "btnCancel";
            btnCancel.Size = new System.Drawing.Size(62, 26);
            btnCancel.TabIndex = 189;
            btnCancel.Text = "Cancel";
            btnCancel.UseVisualStyleBackColor = false;
            // 
            // txtSourceRange
            // 
            txtSourceRange.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            txtSourceRange.Cursor = System.Windows.Forms.Cursors.IBeam;
            txtSourceRange.Font = new System.Drawing.Font("Segoe UI", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            txtSourceRange.Location = new System.Drawing.Point(144, 15);
            txtSourceRange.Name = "txtSourceRange";
            txtSourceRange.Size = new System.Drawing.Size(292, 25);
            txtSourceRange.TabIndex = 204;
            // 
            // Selection
            // 
            Selection.BackColor = System.Drawing.Color.White;
            Selection.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            Selection.Image = (System.Drawing.Image)resources.GetObject("Selection.Image");
            Selection.Location = new System.Drawing.Point(412, 15);
            Selection.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            Selection.Name = "Selection";
            Selection.Size = new System.Drawing.Size(24, 25);
            Selection.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            Selection.TabIndex = 206;
            Selection.TabStop = false;
            // 
            // AutoSelection
            // 
            AutoSelection.BackColor = System.Drawing.Color.White;
            AutoSelection.Image = (System.Drawing.Image)resources.GetObject("AutoSelection.Image");
            AutoSelection.Location = new System.Drawing.Point(385, 16);
            AutoSelection.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            AutoSelection.Name = "AutoSelection";
            AutoSelection.Size = new System.Drawing.Size(24, 23);
            AutoSelection.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            AutoSelection.TabIndex = 205;
            AutoSelection.TabStop = false;
            // 
            // CustomGroupBox5
            // 
            CustomGroupBox5.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            CustomGroupBox5.Controls.Add(CustomPanel1);
            CustomGroupBox5.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 0);
            CustomGroupBox5.Location = new System.Drawing.Point(15, 110);
            CustomGroupBox5.Name = "CustomGroupBox5";
            CustomGroupBox5.Size = new System.Drawing.Size(421, 164);
            CustomGroupBox5.TabIndex = 187;
            CustomGroupBox5.TabStop = false;
            CustomGroupBox5.Text = "Sample Image";
            // 
            // CustomPanel1
            // 
            CustomPanel1.BackColor = System.Drawing.Color.White;
            CustomPanel1.BorderColor = System.Drawing.Color.FromArgb(191, 191, 191);
            CustomPanel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            CustomPanel1.BorderWidth = 1;
            CustomPanel1.Location = new System.Drawing.Point(1, 30);
            CustomPanel1.Name = "CustomPanel1";
            CustomPanel1.Size = new System.Drawing.Size(420, 134);
            CustomPanel1.TabIndex = 0;
            // 
            // destinationSelection
            // 
            destinationSelection.BackColor = System.Drawing.Color.White;
            destinationSelection.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            destinationSelection.Image = (System.Drawing.Image)resources.GetObject("destinationSelection.Image");
            destinationSelection.Location = new System.Drawing.Point(412, 46);
            destinationSelection.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            destinationSelection.Name = "destinationSelection";
            destinationSelection.Size = new System.Drawing.Size(24, 25);
            destinationSelection.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            destinationSelection.TabIndex = 209;
            destinationSelection.TabStop = false;
            // 
            // txtDestRange
            // 
            txtDestRange.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            txtDestRange.Cursor = System.Windows.Forms.Cursors.IBeam;
            txtDestRange.Font = new System.Drawing.Font("Segoe UI", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            txtDestRange.Location = new System.Drawing.Point(144, 46);
            txtDestRange.Name = "txtDestRange";
            txtDestRange.Size = new System.Drawing.Size(292, 25);
            txtDestRange.TabIndex = 208;
            // 
            // Label2
            // 
            Label2.AutoSize = true;
            Label2.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            Label2.Location = new System.Drawing.Point(15, 49);
            Label2.Name = "Label2";
            Label2.Size = new System.Drawing.Size(126, 17);
            Label2.TabIndex = 207;
            Label2.Text = "Destination Range :";
            // 
            // Form16PasteintoVisibleRange
            // 
            AutoScaleDimensions = new System.Drawing.SizeF(6.0f, 13.0f);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            ClientSize = new System.Drawing.Size(457, 361);
            Controls.Add(destinationSelection);
            Controls.Add(Selection);
            Controls.Add(AutoSelection);
            Controls.Add(txtDestRange);
            Controls.Add(txtSourceRange);
            Controls.Add(Label2);
            Controls.Add(CB_keepFormat);
            Controls.Add(btnOK);
            Controls.Add(ComboBox1);
            Controls.Add(CB_copyWs);
            Controls.Add(Label1);
            Controls.Add(CustomGroupBox5);
            Controls.Add(btnCancel);
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "Form16PasteintoVisibleRange";
            Text = "Paste into Visible Range";
            ((System.ComponentModel.ISupportInitialize)Selection).EndInit();
            ((System.ComponentModel.ISupportInitialize)AutoSelection).EndInit();
            CustomGroupBox5.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)destinationSelection).EndInit();
            KeyDown += new System.Windows.Forms.KeyEventHandler(Form1_KeyDown);
            Load += new EventHandler(Form16PasteintoVisibleRange_Load);
            Activated += new EventHandler(Form1_Activated);
            Closing += new System.ComponentModel.CancelEventHandler(Form16PasteintoVisibleRange_Closing);
            Disposed += new EventHandler(Form16PasteintoVisibleRange_Disposed);
            Shown += new EventHandler(Form16PasteintoVisibleRange_Shown);
            ResumeLayout(false);
            PerformLayout();

        }

        internal CustomPanel CustomPanel1;
        internal System.Windows.Forms.CheckBox CB_keepFormat;
        internal System.Windows.Forms.Button btnOK;
        internal System.Windows.Forms.ComboBox ComboBox1;
        internal System.Windows.Forms.CheckBox CB_copyWs;
        internal System.Windows.Forms.Label Label1;
        internal CustomGroupBox CustomGroupBox5;
        internal System.Windows.Forms.Button btnCancel;
        internal System.Windows.Forms.TextBox txtSourceRange;
        internal System.Windows.Forms.PictureBox Selection;
        internal System.Windows.Forms.PictureBox AutoSelection;
        internal System.Windows.Forms.PictureBox destinationSelection;
        internal System.Windows.Forms.TextBox txtDestRange;
        internal System.Windows.Forms.Label Label2;
    }
}