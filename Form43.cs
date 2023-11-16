using System;
using System.ComponentModel;

namespace VSTO_Addins
{

    public partial class Form43
    {
        private Form33_ColorBasedDropDownList form = null;

        public Form43()
        {
            InitializeComponent();
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            form = new Form33_ColorBasedDropDownList();
            form.Show();
            Dispose();
        }

        private void Button2_Click(object sender, EventArgs e)
        {

            GlobalModule.form_flag = false;
            Dispose();
        }

        private void CheckBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (CheckBox1.Checked == true)
            {
                GlobalModule.sessionflag1 = false;
            }
            else
            {
                GlobalModule.sessionflag1 = true;
            }
        }

        private void Form43_Load(object sender, EventArgs e)
        {

        }

        private void Form43_Closing(object sender, CancelEventArgs e)
        {
            GlobalModule.form_flag = false;
        }

        private void Form43_Disposed(object sender, EventArgs e)
        {
            GlobalModule.form_flag = false;
        }
    }
}