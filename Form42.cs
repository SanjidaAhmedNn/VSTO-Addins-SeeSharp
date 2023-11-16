using System;
using System.ComponentModel;

namespace VSTO_Addins
{

    public partial class Form42
    {
        private Form29_Simple_Drop_down_List form1 = null;
        private Form30_Create_Dynamic_Drop_down_List form2 = null;

        public Form42()
        {
            InitializeComponent();
        }
        private void RadioButton5_CheckedChanged(object sender, EventArgs e)
        {
            if (RB_No.Checked == true)
            {
                CGB.Enabled = false;
                RB_Simple.Enabled = false;
                RB_Dynamic.Enabled = false;
            }
        }

        private void RB_Yes_CheckedChanged(object sender, EventArgs e)
        {
            if (RB_Yes.Checked == true)
            {
                CGB.Enabled = true;
                RB_Simple.Enabled = true;
                RB_Dynamic.Enabled = true;

            }
        }

        private void Form42_Load(object sender, EventArgs e)
        {

        }

        private void Btn_Cancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void Btn_OK_Click(object sender, EventArgs e)
        {
            if (RB_Simple.Checked == true & RB_Simple.Enabled == true)
            {
                form1 = new Form29_Simple_Drop_down_List();
                Hide();
                form1.Show();
            }
            else if (RB_Dynamic.Checked == true & RB_Dynamic.Enabled == true)
            {
                form2 = new Form30_Create_Dynamic_Drop_down_List();
                Hide();
                form2.Show();
            }
            else
            {
                Dispose();
            }
            Close();

        }

        private void CheckBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (CheckBox1.Checked == true)
            {
                GlobalModule.sessionflag2 = false;
            }
            else
            {
                GlobalModule.sessionflag2 = true;
            }
        }

        private void Form42_Closing(object sender, CancelEventArgs e)
        {
            GlobalModule.form_flag = false;
        }

        private void Form42_Disposed(object sender, EventArgs e)
        {
            GlobalModule.form_flag = false;
        }

        private void RB_Simple_CheckedChanged(object sender, EventArgs e)
        {

        }
    }
}