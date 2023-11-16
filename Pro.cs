using System;

namespace VSTO_Addins
{

    public partial class FormProgressBar
    {
        public static System.Windows.Forms.ProgressBar proBar;

        public FormProgressBar()
        {
            InitializeComponent();
        }
        private void Pro_Load(object sender, EventArgs e)
        {
            // For i As Integer = 0 To 100
            // ProgressBar1.Value = i



            // ' Include your task logic here
            // ' Use Application.DoEvents() if needed to refresh the UI
            // Next

            // Label1.Text = Ribbon1.captiontxt


            proBar = ProgressBar1;
            proBar.Minimum = 0;
            proBar.Maximum = 100;

            Label1.Text = "Progress: " + proBar.Value.ToString() + "%";

        }


        private void HandleProgressBarValueChanged()
        {
            // Me.Label1.Text = "Progress: " & proBar.Value.ToString() & "%"
        }
    }
}