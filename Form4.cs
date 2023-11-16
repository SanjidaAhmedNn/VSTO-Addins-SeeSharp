using System;
using System.Drawing;
using System.IO;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;

namespace VSTO_Addins
{


    public partial class Form4
    {
        private Excel.Application _excelApp;

        public virtual Excel.Application excelApp
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _excelApp;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                _excelApp = value;
            }
        }
        public Excel.Workbook workbook;
        public Excel.Workbook workbook2;
        public Excel.Worksheet worksheet;
        public Excel.Worksheet worksheet1;
        public Excel.Worksheet worksheet2;
        public Excel.Worksheet OpenSheet;
        public Range rng;
        public Range rng2;
        public int FocusedTextBox;
        public int Opened;
        public int GB6;
        private int ThisFocusedTextBox;
        public int Form4Open;
        public bool Workbook2Opened;
        public int CB1;
        public int CB2;
        public bool TextBoxChanged;

        public Form4()
        {
            InitializeComponent();
        }

        private bool IsValidExcelFile(string filePath)
        {
            // Check if the file exists.
            if (!File.Exists(filePath))
            {
                return false;
            }

            else
            {

                // Get the file extension.
                string extension = Path.GetExtension(filePath);

                // Check if the extension is a valid Excel extension.
                if (extension == ".xls" || extension == ".xlsx" || extension == ".xlsm")
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }

        }
        private void Setup()
        {

            try
            {
                if (RadioButton1.Checked == true)
                {
                    TextBox1.Enabled = true;
                    PictureBox8.Enabled = true;
                }
                else
                {
                    TextBox1.Clear();
                    TextBox1.Enabled = false;
                    PictureBox8.Enabled = false;
                }

                if (RadioButton2.Checked == true)
                {
                    TextBox2.Enabled = true;
                    PictureBox1.Enabled = true;
                    TextBox3.Enabled = true;
                    PictureBox2.Enabled = true;
                    Label1.Enabled = true;
                    PictureBox3.Enabled = true;
                }
                else
                {
                    TextBox2.Clear();
                    TextBox3.Clear();
                    TextBox2.Enabled = false;
                    PictureBox1.Enabled = false;
                    TextBox3.Enabled = false;
                    PictureBox2.Enabled = false;
                    Label1.Enabled = false;
                    PictureBox3.Enabled = false;
                }
            }

            catch (Exception ex)
            {

            }

        }

        // Worksheet.Name = "New Worksheet"
        private void RadioButton1_CheckedChanged(object sender, EventArgs e)
        {

            try
            {
                if (RadioButton1.Checked == true)
                {
                    workbook2 = excelApp.Workbooks.Add();
                    Show();
                    TextBox1.Focus();
                    Workbook2Opened = true;
                    Setup();
                }
            }

            catch (Exception ex)
            {

            }

        }



        private void PictureBox1_Click(object sender, EventArgs e)
        {

            try
            {
                ThisFocusedTextBox = 2;

                Hide();
                var openFileDialog = new OpenFileDialog();
                openFileDialog.Title = "Open Your File";
                openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string filePath = openFileDialog.FileName;
                    workbook2 = excelApp.Workbooks.Open(filePath);
                    TextBox2.Text = filePath;
                    excelApp.Visible = true;
                    Workbook2Opened = true;
                }

                Show();
                TextBox2.Focus();
            }

            catch (Exception ex)
            {

            }

        }

        private void PictureBox8_Click(object sender, EventArgs e)
        {

            try
            {
                ThisFocusedTextBox = 1;
                Hide();

                Range userInput = (Range)excelApp.InputBox("Select a Cell.", Type: 8);
                rng2 = userInput;

                string sheetName;
                sheetName = Strings.Split(rng2.get_Address(true, true, XlReferenceStyle.xlA1, true), "]")[1];
                sheetName = Strings.Split(sheetName, "!")[0];
                worksheet2 = (Excel.Worksheet)workbook2.Worksheets[sheetName];
                worksheet2.Activate();

                rng2.Select();

                TextBox1.Text = rng2.get_Address();

                Show();
                TextBox1.Focus();
            }

            catch (Exception ex)
            {

                Show();
                TextBox1.Focus();

            }


        }

        private void Button1_Click(object sender, EventArgs e)
        {

            try
            {
                var MyForm3 = new Form3();
                MyForm3.excelApp = excelApp;
                Form4Open = 1;
                MyForm3.Form4Open = Form4Open;
                MyForm3.rng = rng;
                MyForm3.workbook = workbook;
                MyForm3.workbook2 = workbook2;
                MyForm3.worksheet = worksheet;
                MyForm3.worksheet2 = worksheet2;
                MyForm3.OpenSheet = OpenSheet;
                MyForm3.rng2 = rng2;
                MyForm3.TextBoxChanged = TextBoxChanged;
                if ((worksheet.Name ?? "") != (OpenSheet.Name ?? ""))
                {
                    MyForm3.TextBox1.Text = worksheet.Name + "!" + rng.get_Address();
                }
                else
                {
                    MyForm3.TextBox1.Text = rng.get_Address();
                }
                MyForm3.Workbook2Opened = Workbook2Opened;

                if (GB6 == 3)
                {
                    MyForm3.RadioButton3.Checked = true;
                }
                else if (GB6 == 2)
                {
                    MyForm3.RadioButton2.Checked = true;
                }

                if (CB1 == 1)
                {
                    MyForm3.CheckBox1.Checked = true;
                }
                if (CB2 == 1)
                {
                    MyForm3.CheckBox2.Checked = true;
                }

                MyForm3.RadioButton5.Checked = true;
                MyForm3.Opened = Opened;
                MyForm3.Show();
                Close();
            }

            catch (Exception ex)
            {

            }

        }

        private void PictureBox2_Click(object sender, EventArgs e)
        {

            try
            {
                ThisFocusedTextBox = 3;
                Hide();

                Range userInput = (Range)excelApp.InputBox("Select a Cell", Type: 8);
                rng2 = userInput;


                string sheetName;
                sheetName = Strings.Split(rng2.get_Address(true, true, XlReferenceStyle.xlA1, true), "]")[1];
                sheetName = Strings.Split(sheetName, "!")[0];
                worksheet2 = (Excel.Worksheet)workbook2.Worksheets[sheetName];
                worksheet2.Activate();

                rng2.Select();

                TextBox3.Text = rng2.get_Address();

                Show();
                TextBox3.Focus();
            }

            catch (Exception ex)
            {

                Show();
                TextBox3.Focus();

            }

        }

        private void RadioButton2_CheckedChanged(object sender, EventArgs e)
        {

            try
            {
                if (RadioButton2.Checked == true)
                {
                    Setup();
                }
            }
            catch (Exception ex)
            {

            }
        }

        private void TextBox2_TextChanged(object sender, EventArgs e)
        {
            try
            {
                if (!string.IsNullOrEmpty(TextBox2.Text))
                {
                    if (IsValidExcelFile(TextBox2.Text) == true)
                    {
                        string filePath = TextBox2.Text;
                        workbook2 = excelApp.Workbooks.Open(filePath);
                        Workbook2Opened = true;
                        excelApp.Visible = true;
                    }
                }
            }
            catch (Exception ex)
            {

            }
        }

        private void Button2_Click(object sender, EventArgs e)
        {

            try
            {
                var MyForm3 = new Form3();
                MyForm3.excelApp = excelApp;
                MyForm3.Form4Open = Form4Open;
                MyForm3.rng = rng;
                MyForm3.workbook = workbook;
                MyForm3.worksheet = worksheet;
                MyForm3.OpenSheet = OpenSheet;
                MyForm3.TextBoxChanged = TextBoxChanged;
                if ((worksheet.Name ?? "") != (OpenSheet.Name ?? ""))
                {
                    MyForm3.TextBox1.Text = worksheet.Name + "!" + rng.get_Address();
                }
                else
                {
                    MyForm3.TextBox1.Text = rng.get_Address();
                }
                MyForm3.Workbook2Opened = Workbook2Opened;

                if (GB6 == 3)
                {
                    MyForm3.RadioButton3.Checked = true;
                }
                else if (GB6 == 2)
                {
                    MyForm3.RadioButton2.Checked = true;
                }

                if (CB1 == 1)
                {
                    MyForm3.CheckBox1.Checked = true;
                }
                if (CB2 == 1)
                {
                    MyForm3.CheckBox2.Checked = true;
                }

                MyForm3.Opened = Opened;
                MyForm3.Show();
                Close();
                if (Workbook2Opened == true)
                {
                    workbook2.Close();
                    workbook.Activate();
                }
            }
            catch (Exception ex)
            {

            }

        }

        private void TextBox3_TextChanged(object sender, EventArgs e)
        {

            try
            {

                if (!string.IsNullOrEmpty(TextBox3.Text))
                {
                    worksheet2 = (Excel.Worksheet)workbook2.ActiveSheet;
                    rng2 = worksheet2.get_Range(TextBox3.Text);
                    rng2.Select();
                }
            }

            catch (Exception ex)
            {

            }
        }

        private void Form4_Loaded(object sender, EventArgs e)
        {

            try
            {

                excelApp.SheetSelectionChange += excelApp_SheetSelectionChange;
                KeyPreview = true;

                Setup();
            }

            catch (Exception ex)
            {

            }

        }

        private void excelApp_SheetSelectionChange(object Sh, Range Target)
        {

            try
            {

                Range selectedRange;
                selectedRange = (Range)excelApp.Selection;

                if (ThisFocusedTextBox == 1)
                {
                    TextBox1.Text = selectedRange.get_Address();
                    worksheet2 = (Excel.Worksheet)workbook2.ActiveSheet;
                    rng2 = selectedRange;
                    TextBox1.Focus();
                }

                else if (ThisFocusedTextBox == 3)
                {
                    TextBox3.Text = selectedRange.get_Address();
                    worksheet2 = (Excel.Worksheet)workbook2.ActiveSheet;
                    rng2 = selectedRange;
                    TextBox3.Focus();
                }
            }

            catch (Exception ex)
            {

            }

        }
        private void TextBox1_TextChanged(object sender, EventArgs e)
        {

            try
            {

                if (!string.IsNullOrEmpty(TextBox1.Text))
                {
                    worksheet2 = (Excel.Worksheet)workbook2.ActiveSheet;
                    rng2 = worksheet2.get_Range(TextBox1.Text);
                    rng2.Select();
                }
            }

            catch (Exception ex)
            {

            }
        }

        private void TextBox1_GotFocus(object sender, EventArgs e)
        {

            try
            {
                ThisFocusedTextBox = 1;
            }
            catch (Exception ex)
            {

            }
        }

        private void TextBox2_GotFocus(object sender, EventArgs e)
        {
            try
            {
                ThisFocusedTextBox = 0;
            }

            catch (Exception ex)
            {

            }

        }

        private void TextBox3_GotFocus(object sender, EventArgs e)
        {
            try
            {
                ThisFocusedTextBox = 3;
            }
            catch (Exception ex)
            {

            }
        }

        private void PictureBox8_GotFocus(object sender, EventArgs e)
        {

            try
            {
                ThisFocusedTextBox = 1;
            }
            catch (Exception ex)
            {

            }

        }

        private void PictureBox2_GotFocus(object sender, EventArgs e)
        {

            try
            {
                ThisFocusedTextBox = 3;
            }
            catch (Exception ex)
            {

            }

        }

        private void RadioButton1_GotFocus(object sender, EventArgs e)
        {
            try
            {
                ThisFocusedTextBox = 0;
            }
            catch (Exception ex)
            {

            }
        }

        private void RadioButton2_GotFocus(object sender, EventArgs e)
        {

            try
            {
                ThisFocusedTextBox = 0;
            }
            catch (Exception ex)
            {

            }

        }

        private void PictureBox1_GotFocus(object sender, EventArgs e)
        {

            try
            {
                ThisFocusedTextBox = 0;
            }
            catch (Exception ex)
            {

            }

        }

        private void PictureBox3_GotFocus(object sender, EventArgs e)
        {

            try
            {
                ThisFocusedTextBox = 0;
            }
            catch (Exception ex)
            {

            }

        }

        private void Button1_GotFocus(object sender, EventArgs e)
        {

            try
            {
                ThisFocusedTextBox = 0;
            }
            catch (Exception ex)
            {

            }

        }

        private void Button2_GotFocus(object sender, EventArgs e)
        {

            try
            {
                ThisFocusedTextBox = 0;
            }
            catch (Exception ex)
            {

            }

        }

        private void Button3_GotFocus(object sender, EventArgs e)
        {

            try
            {
                ThisFocusedTextBox = 0;
            }
            catch (Exception ex)
            {

            }

        }

        private void Button3_Click(object sender, EventArgs e)
        {

            try
            {
                var MyForm3 = new Form3();
                MyForm3.excelApp = excelApp;
                MyForm3.Form4Open = Form4Open;
                MyForm3.rng = rng;
                MyForm3.workbook = workbook;
                MyForm3.worksheet = worksheet;
                MyForm3.OpenSheet = OpenSheet;
                MyForm3.TextBoxChanged = TextBoxChanged;

                if ((worksheet.Name ?? "") != (OpenSheet.Name ?? ""))
                {
                    MyForm3.TextBox1.Text = worksheet.Name + "!" + rng.get_Address();
                }
                else
                {
                    MyForm3.TextBox1.Text = rng.get_Address();
                }

                if (GB6 == 3)
                {
                    MyForm3.RadioButton3.Checked = true;
                }
                else if (GB6 == 2)
                {
                    MyForm3.RadioButton2.Checked = true;
                }

                if (CB1 == 1)
                {
                    MyForm3.CheckBox1.Checked = true;
                }
                if (CB2 == 1)
                {
                    MyForm3.CheckBox2.Checked = true;
                }

                MyForm3.Opened = Opened;
                MyForm3.Show();
                Close();

                if (Workbook2Opened == true)
                {
                    workbook2.Close();
                    workbook.Activate();
                }
            }

            catch (Exception ex)
            {

            }

        }

        private void Button1_MouseEnter(object sender, EventArgs e)
        {


            try
            {

                Button1.ForeColor = Color.White;
                Button1.BackColor = Color.FromArgb(76, 111, 174);
            }

            catch (Exception ex)
            {

            }

        }

        private void Button2_MouseEnter(object sender, EventArgs e)
        {


            try
            {

                Button2.ForeColor = Color.White;
                Button2.BackColor = Color.FromArgb(76, 111, 174);
            }

            catch (Exception ex)
            {

            }

        }

        private void Button3_MouseEnter(object sender, EventArgs e)
        {


            try
            {

                Button3.ForeColor = Color.White;
                Button3.BackColor = Color.FromArgb(76, 111, 174);
            }

            catch (Exception ex)
            {

            }

        }

        private void Button1_MouseLeave(object sender, EventArgs e)
        {

            try
            {

                Button1.ForeColor = Color.FromArgb(70, 70, 70);
                Button1.BackColor = Color.White;
            }

            catch (Exception ex)
            {

            }

        }

        private void Button2_MouseLeave(object sender, EventArgs e)
        {

            try
            {

                Button2.ForeColor = Color.FromArgb(70, 70, 70);
                Button2.BackColor = Color.White;
            }

            catch (Exception ex)
            {

            }

        }

        private void Button3_MouseLeave(object sender, EventArgs e)
        {

            try
            {

                Button3.ForeColor = Color.FromArgb(70, 70, 70);
                Button3.BackColor = Color.White;
            }

            catch (Exception ex)
            {

            }

        }

        private void Form4_KeyDown(object sender, KeyEventArgs e)
        {

            try
            {

                if (e.KeyCode == Keys.Enter)
                {
                    Button1_Click(sender, e);
                }
            }

            catch (Exception ex)
            {

            }

        }

    }
}