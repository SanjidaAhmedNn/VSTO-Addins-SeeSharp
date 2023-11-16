using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace VSTO_Addins
{

    public partial class Form3
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
        public Excel.Worksheet worksheet2;
        public Excel.Worksheet OpenSheet;

        public Range rng;
        public Range rng2;
        public int FocusedTextBox;
        public int Opened;

        public int Form4Open;
        public bool Workbook2Opened;

        public bool TextBoxChanged;

        public Form3()
        {
            InitializeComponent();
        }


        [DllImport("user32")]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
        private const uint SWP_NOMOVE = 0x2U;
        private const uint SWP_NOSIZE = 0x1U;
        private const uint SWP_NOACTIVATE = 0x10U;
        private const int HWND_TOPMOST = -1;


        private bool IsValidExcelCellReference(string cellReference)
        {

            string cellPattern = @"(\$?[A-Z]+\$?[0-9]+)";
            string referencePattern = "^" + cellPattern + "(:" + cellPattern + ")?$";

            var regex = new Regex(referencePattern);

            string[] refArr = Strings.Split(cellReference, "!");

            string reference = refArr[Information.UBound(refArr)];

            if (regex.IsMatch(reference))
            {
                return true;
            }
            else
            {
                return false;
            }

        }
        private bool Overlap(Excel.Application excelApp, Excel.Worksheet sheet1, Excel.Worksheet sheet2, Range rng1, Range rng2)
        {

            if ((sheet1.Name ?? "") != (sheet2.Name ?? ""))
            {
                return false;
            }

            else
            {
                Excel.Worksheet activesheet = (Excel.Worksheet)excelApp.ActiveSheet;

                var rng3 = activesheet.get_Range(rng1.get_Address());
                var rng4 = activesheet.get_Range(rng2.get_Address());

                var intersectRange = excelApp.Intersect(rng3, rng4);

                if (intersectRange is null)
                {
                    return false;
                }
                else
                {
                    return true;
                }
            }

        }
        private object MaxOfColumn(Range cRng)
        {
            object MaxOfColumnRet = default;

            int max;
            int CharNumbers;

            if (Information.IsNumeric(cRng.Cells[1, 1].value))
            {
                max = Strings.Len(Conversion.Str(cRng.Cells[1, 1].value));
            }
            else
            {
                max = Strings.Len(cRng.Cells[1, 1].value);
            }

            for (int i = 2, loopTo = cRng.Rows.Count; i <= loopTo; i++)
            {
                if (Information.IsNumeric(cRng.Cells[i, 1].value))
                {
                    CharNumbers = Strings.Len(Conversion.Str(cRng.Cells[i, 1].value));
                }
                else
                {
                    CharNumbers = Strings.Len(cRng.Cells[i, 1].value);
                }
                if (CharNumbers > max)
                {
                    max = CharNumbers;
                }
            }

            if (max < 7)
            {
                max = 7;
            }

            MaxOfColumnRet = max;
            return MaxOfColumnRet;

        }
        private object MaxOfArray(object Arr)
        {
            object MaxOfArrayRet = default;

            int max;
            max = Strings.Len(Arr((object)Information.LBound((Array)Arr)));

            for (int i = Information.LBound((Array)Arr) + 1, loopTo = Information.UBound((Array)Arr); i <= loopTo; i++)
            {
                if (Strings.Len(Arr((object)i)) > max)
                {
                    max = Strings.Len(Arr((object)i));
                }
            }

            if (max < 7)
            {
                max = 7;
            }

            MaxOfArrayRet = max;
            return MaxOfArrayRet;

        }

        private void Display()
        {

            try
            {

                panel1.Controls.Clear();
                panel2.Controls.Clear();

                Range displayRng;

                if (rng.Rows.Count > 50)
                {
                    displayRng = (Range)rng.Rows["1:50"];
                }
                else
                {
                    displayRng = rng;
                }

                int r;
                int c;

                r = displayRng.Rows.Count;
                c = displayRng.Columns.Count;

                float height;
                double Basewidth;
                double width;
                Basewidth = 260d / 3d;

                if (displayRng.Rows.Count <= 4)
                {
                    height = (float)(panel1.Height / (double)displayRng.Rows.Count);
                }
                else
                {
                    height = (float)(119d / 4d);
                }

                Range CRng;
                double Ordinate = 0d;

                for (int j = 1, loopTo = c; j <= loopTo; j++)
                {

                    CRng = worksheet.get_Range(displayRng.Cells[1, j], displayRng.Cells[r, j]);
                    width = Conversions.ToDouble(Operators.DivideObject(Operators.MultiplyObject(MaxOfColumn(CRng), Basewidth), 10));

                    for (int i = 1, loopTo1 = r; i <= loopTo1; i++)
                    {
                        var label = new System.Windows.Forms.Label();
                        label.Text = Conversions.ToString(displayRng.Cells[i, j].Value);
                        label.Location = new System.Drawing.Point((int)Math.Round(Ordinate), (int)Math.Round((i - 1) * height));
                        label.Height = (int)Math.Round(height);
                        label.Width = (int)Math.Round(width);
                        label.BorderStyle = BorderStyle.FixedSingle;
                        label.TextAlign = ContentAlignment.MiddleCenter;

                        if (CheckBox2.Checked == true)
                        {

                            Range cell = (Range)displayRng.Cells[i, j];
                            var font = cell.Font;
                            var fontStyle = FontStyle.Regular;
                            if (Conversions.ToBoolean(cell.Font.Bold))
                                fontStyle = fontStyle | FontStyle.Bold;
                            if (Conversions.ToBoolean(cell.Font.Italic))
                                fontStyle = fontStyle | FontStyle.Italic;

                            float fontSize = Convert.ToSingle(font.Size);

                            label.Font = new System.Drawing.Font(font.ToString(), fontSize, fontStyle);
                            if (Conversions.ToBoolean(!Operators.ConditionalCompareObjectEqual(cell.Interior.ColorIndex, XlColorIndex.xlColorIndexNone, false)))
                            {
                                long colorValue1 = Conversions.ToLong(cell.Interior.Color);
                                int red1 = (int)(colorValue1 % 256L);
                                int green1 = (int)(colorValue1 / 256L % 256L);
                                int blue1 = (int)(colorValue1 / 256L / 256L % 256L);
                                label.BackColor = Color.FromArgb(red1, green1, blue1);
                            }

                            if (cell.Font.Color is DBNull)
                            {
                                label.ForeColor = Color.FromArgb(0, 0, 0);
                            }

                            else if (Conversions.ToBoolean(!Operators.ConditionalCompareObjectEqual(cell.Font.ColorIndex, XlColorIndex.xlColorIndexNone, false)))
                            {
                                long colorValue2 = Conversions.ToLong(cell.Font.Color);
                                int red2 = (int)(colorValue2 % 256L);
                                int green2 = (int)(colorValue2 / 256L % 256L);
                                int blue2 = (int)(colorValue2 / 256L / 256L % 256L);
                                label.ForeColor = Color.FromArgb(red2, green2, blue2);
                            }
                        }
                        panel1.Controls.Add(label);
                    }
                    Ordinate = Ordinate + width;
                }

                panel1.AutoScroll = true;

                if (RadioButton2.Checked == true | RadioButton3.Checked == true)
                {

                    var Values = new object[c];
                    var Widths = new object[r];

                    if (c <= 4)
                    {
                        height = (float)(panel2.Height / (double)c);
                    }
                    else
                    {
                        height = (float)(panel2.Height / 4d);
                    }

                    for (int i = 1, loopTo2 = r; i <= loopTo2; i++)
                    {
                        for (int j = 1, loopTo3 = c; j <= loopTo3; j++)
                            Values[j - 1] = displayRng.Cells[i, j].value;
                        Widths[i - 1] = Operators.DivideObject(Operators.MultiplyObject(MaxOfArray(Values), Basewidth), 10);
                    }

                    Ordinate = 0d;

                    for (int i = 1, loopTo4 = displayRng.Rows.Count; i <= loopTo4; i++)
                    {
                        for (int j = 1, loopTo5 = displayRng.Columns.Count; j <= loopTo5; j++)
                        {
                            var label = new System.Windows.Forms.Label();
                            label.Text = Conversions.ToString(displayRng.Cells[i, j].Value);
                            label.Location = new System.Drawing.Point((int)Math.Round(Ordinate), (int)Math.Round((j - 1) * height));
                            label.Height = (int)Math.Round(height);
                            label.Width = Conversions.ToInteger(Widths[i - 1]);
                            label.BorderStyle = BorderStyle.FixedSingle;
                            label.TextAlign = ContentAlignment.MiddleCenter;

                            if (CheckBox2.Checked == true)
                            {
                                Range cell = (Range)displayRng.Cells[i, j];
                                var font = cell.Font;
                                var fontStyle = FontStyle.Regular;
                                if (Conversions.ToBoolean(cell.Font.Bold))
                                    fontStyle = fontStyle | FontStyle.Bold;
                                if (Conversions.ToBoolean(cell.Font.Italic))
                                    fontStyle = fontStyle | FontStyle.Italic;

                                float fontSize = Convert.ToSingle(font.Size);

                                label.Font = new System.Drawing.Font(font.ToString(), fontSize, fontStyle);
                                if (Conversions.ToBoolean(!Operators.ConditionalCompareObjectEqual(cell.Interior.ColorIndex, XlColorIndex.xlColorIndexNone, false)))
                                {
                                    long colorValue1 = Conversions.ToLong(cell.Interior.Color);
                                    int red1 = (int)(colorValue1 % 256L);
                                    int green1 = (int)(colorValue1 / 256L % 256L);
                                    int blue1 = (int)(colorValue1 / 256L / 256L % 256L);
                                    label.BackColor = Color.FromArgb(red1, green1, blue1);
                                }

                                if (cell.Font.Color is DBNull)
                                {
                                    label.ForeColor = Color.FromArgb(0, 0, 0);
                                }

                                else if (Conversions.ToBoolean(!Operators.ConditionalCompareObjectEqual(cell.Font.ColorIndex, XlColorIndex.xlColorIndexNone, false)))
                                {
                                    long colorValue2 = Conversions.ToLong(cell.Font.Color);
                                    int red2 = (int)(colorValue2 % 256L);
                                    int green2 = (int)(colorValue2 / 256L % 256L);
                                    int blue2 = (int)(colorValue2 / 256L / 256L % 256L);
                                    label.ForeColor = Color.FromArgb(red2, green2, blue2);
                                }
                            }
                            panel2.Controls.Add(label);
                        }
                        Ordinate = Conversions.ToDouble(Operators.AddObject(Ordinate, Widths[i - 1]));
                    }

                    panel2.AutoScroll = true;

                }
            }

            catch (Exception ex)
            {

            }

        }

        private void DestinationChange()
        {

            try
            {
                if (RadioButton1.Checked == true)
                {
                    if (Form4Open == 1)
                    {
                        if (Workbook2Opened == true)
                        {
                            workbook2.Close();
                            workbook.Activate();
                        }
                        workbook2 = workbook;
                        Form4Open = 0;
                    }
                    TextBox2.Visible = true;
                    PictureBox2.Visible = true;
                    TextBox2.Location = new System.Drawing.Point(121, 7);
                    PictureBox2.Location = new System.Drawing.Point(226, 7);
                    TextBox2.Focus();
                }
                else
                {
                    TextBox2.Clear();
                }

                if (RadioButton4.Checked == true)
                {

                    if (Form4Open == 1)
                    {
                        if (Workbook2Opened == true)
                        {
                            workbook2.Close();
                            workbook.Activate();
                        }
                        workbook2 = workbook;
                        Form4Open = 0;
                    }
                    TextBox2.Visible = true;
                    PictureBox2.Visible = true;
                    TextBox2.Location = new System.Drawing.Point(121, 30);
                    PictureBox2.Location = new System.Drawing.Point(226, 30);

                    Excel.Worksheet ws = (Excel.Worksheet)workbook.Worksheets.Add();
                    TextBox2.Focus();
                }
                else
                {
                    TextBox2.Clear();
                }

                if (RadioButton5.Checked == true & Form4Open == 0)
                {
                    TextBox2.Visible = false;
                    PictureBox2.Visible = false;
                    var MyForm4 = new Form4();
                    MyForm4.excelApp = excelApp;
                    MyForm4.workbook = workbook;
                    MyForm4.worksheet = worksheet;
                    MyForm4.OpenSheet = OpenSheet;
                    MyForm4.rng = rng;
                    MyForm4.Opened = Opened;
                    MyForm4.FocusedTextBox = FocusedTextBox;
                    MyForm4.TextBoxChanged = TextBoxChanged;
                    MyForm4.Form4Open = Form4Open;
                    MyForm4.Workbook2Opened = false;
                    if (RadioButton3.Checked == true)
                    {
                        MyForm4.GB6 = 3;
                    }
                    else if (RadioButton2.Checked == true)
                    {
                        MyForm4.GB6 = 2;
                    }
                    else
                    {
                        MyForm4.GB6 = 0;
                    }
                    if (CheckBox1.Checked == true)
                    {
                        MyForm4.CB1 = 1;
                    }
                    else
                    {
                        MyForm4.CB1 = 0;
                    }
                    if (CheckBox2.Checked == true)
                    {
                        MyForm4.CB2 = 1;
                    }
                    else
                    {
                        MyForm4.CB2 = 0;
                    }
                    Close();
                    MyForm4.Show();

                }
            }

            catch (Exception ex)
            {

            }


        }


        private void btn_OK_MouseLeave(object sender, EventArgs e)
        {

            try
            {

                btn_OK.ForeColor = Color.FromArgb(70, 70, 70);
                btn_OK.BackColor = Color.White;
            }

            catch (Exception ex)
            {

            }

        }

        private void btn_cancel_MouseLeave(object sender, EventArgs e)
        {

            try
            {

                btn_cancel.ForeColor = Color.FromArgb(70, 70, 70);
                btn_cancel.BackColor = Color.White;
            }

            catch (Exception ex)
            {

            }

        }

        private void PictureBox8_Click(object sender, EventArgs e)
        {

            try
            {
                FocusedTextBox = 1;
                Hide();


                Range userInput = (Range)excelApp.InputBox("Select a range", Type: 8);
                rng = userInput;

                string sheetName;
                sheetName = Strings.Split(rng.get_Address(true, true, XlReferenceStyle.xlA1, true), "]")[1];
                sheetName = Strings.Split(sheetName, "!")[0];

                if (Strings.Mid(sheetName, Strings.Len(sheetName), 1) == "'")
                {
                    sheetName = Strings.Mid(sheetName, 1, Strings.Len(sheetName) - 1);
                }

                worksheet = (Excel.Worksheet)workbook.Worksheets[sheetName];
                worksheet.Activate();

                rng.Select();

                if ((worksheet.Name ?? "") != (OpenSheet.Name ?? ""))
                {
                    TextBox1.Text = worksheet.Name + "!" + rng.get_Address();
                }
                else
                {
                    TextBox1.Text = rng.get_Address();
                }

                Show();
                TextBox1.Focus();
            }

            catch (Exception ex)
            {

                Show();
                TextBox1.Focus();

            }

        }

        private void PictureBox4_Click(object sender, EventArgs e)
        {

            try
            {
                FocusedTextBox = 1;

                var activeRange = excelApp.ActiveCell;

                int startRow = activeRange.Row;
                int startColumn = activeRange.Column;
                int endRow = activeRange.Row;
                int endColumn = activeRange.Column;

                // Find the upper boundary
                while (startRow > 1 && !(worksheet.Cells[startRow - 1, startColumn].Value == null))
                    startRow -= 1;

                // Find the lower boundary
                while (!(worksheet.Cells[endRow + 1, endColumn].Value == null))
                    endRow += 1;

                // Find the left boundary
                while (startColumn > 1 && !(worksheet.Cells[startRow, startColumn - 1].Value == null))
                    startColumn -= 1;

                // Find the right boundary
                while (!(worksheet.Cells[endRow, endColumn + 1].Value == null))
                    endColumn += 1;

                // Select the determined range
                rng = worksheet.get_Range(worksheet.Cells[startRow, startColumn], worksheet.Cells[endRow, endColumn]);

                rng.Select();

                string sheetName;

                sheetName = Strings.Split(rng.get_Address(true, true, XlReferenceStyle.xlA1, true), "]")[1];
                sheetName = Strings.Split(sheetName, "!")[0];

                if (Strings.Mid(sheetName, Strings.Len(sheetName), 1) == "'")
                {
                    sheetName = Strings.Mid(sheetName, 1, Strings.Len(sheetName) - 1);
                }

                worksheet = (Excel.Worksheet)workbook.Worksheets[sheetName];
                worksheet.Activate();

                if ((worksheet.Name ?? "") != (OpenSheet.Name ?? ""))
                {
                    TextBox1.Text = worksheet.Name + "!" + rng.get_Address();
                }
                else
                {
                    TextBox1.Text = rng.get_Address();
                }

                TextBox1.Focus();
            }

            catch (Exception ex)
            {

                Show();
                TextBox1.Focus();

            }

        }


        private void RadioButton1_CheckedChanged(object sender, EventArgs e)
        {

            try
            {

                TextBox2.Location = new System.Drawing.Point(121, 7);
                PictureBox2.Location = new System.Drawing.Point(226, 7);
            }

            catch (Exception ex)
            {

            }

        }

        private void PictureBox2_Click(object sender, EventArgs e)
        {

            try
            {
                FocusedTextBox = 2;
                Hide();

                Range userInput = (Range)excelApp.InputBox("Select a Cell.", Type: 8);
                rng2 = userInput;

                string sheetName;
                sheetName = Strings.Split(rng2.get_Address(true, true, XlReferenceStyle.xlA1, true), "]")[1];
                sheetName = Strings.Split(sheetName, "!")[0];

                if (Strings.Mid(sheetName, Strings.Len(sheetName), 1) == "'")
                {
                    sheetName = Strings.Mid(sheetName, 1, Strings.Len(sheetName) - 1);
                }

                worksheet2 = (Excel.Worksheet)workbook.Worksheets[sheetName];
                worksheet2.Activate();

                rng2.Select();

                if ((worksheet2.Name ?? "") != (OpenSheet.Name ?? ""))
                {
                    TextBox2.Text = worksheet2.Name + "!" + rng2.get_Address();
                }
                else
                {
                    TextBox2.Text = rng2.get_Address();
                }

                Show();
                TextBox2.Focus();
            }

            catch (Exception ex)
            {

                Show();
                TextBox2.Focus();

            }

        }

        private void btn_OK_Click(object sender, EventArgs e)
        {

            try
            {

                TextBoxChanged = true;
                if (string.IsNullOrEmpty(TextBox1.Text) | IsValidExcelCellReference(TextBox1.Text) == false)
                {
                    MessageBox.Show("Enter a Valid Source Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    worksheet.Activate();
                    rng.Select();
                    return;
                }

                if (RadioButton2.Checked == false & RadioButton3.Checked == false)
                {
                    MessageBox.Show("Select a Paste Option.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    worksheet.Activate();
                    rng.Select();
                    return;
                }

                if (RadioButton1.Checked == false & RadioButton4.Checked == false & RadioButton5.Checked == false)
                {
                    MessageBox.Show("Select a Destination Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    worksheet.Activate();
                    rng.Select();
                    return;
                }

                if (RadioButton1.Checked == true | RadioButton4.Checked == true)
                {
                    if (string.IsNullOrEmpty(TextBox2.Text) | IsValidExcelCellReference(TextBox2.Text) == false)
                    {
                        MessageBox.Show("Select a Valid Destination Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        worksheet.Activate();
                        rng.Select();
                        return;
                    }
                }


                if (RadioButton2.Checked == true | RadioButton3.Checked == true)
                {

                    rng2 = worksheet2.get_Range(rng2.Cells[1, 1], rng2.Cells[rng.Columns.Count, rng.Rows.Count]);
                    string rng2Address = rng2.get_Address();

                    if (CheckBox1.Checked == true)
                    {
                        worksheet.Copy(After: workbook.Sheets[worksheet.Name]);
                    }

                    worksheet2.Activate();

                    if (Overlap(excelApp, worksheet, worksheet2, rng, rng2) == false)
                    {

                        rng2.ClearFormats();

                        if (RadioButton3.Checked == true)
                        {
                            for (int i = 1, loopTo = rng.Rows.Count; i <= loopTo; i++)
                            {
                                for (int j = 1, loopTo1 = rng.Columns.Count; j <= loopTo1; j++)
                                {
                                    rng2.Cells[j, i].Value = rng.Cells[i, j].Value;
                                    rng2 = worksheet2.get_Range(rng2Address);
                                    if (CheckBox2.Checked == true)
                                    {
                                        rng.Cells[i, j].Copy();
                                        rng2.Cells[j, i].PasteSpecial(XlPasteType.xlPasteFormats);
                                        rng2 = worksheet2.get_Range(rng2Address);
                                        Range sourceCell = (Range)rng.Cells[i, j];
                                        Range targetCell = (Range)rng2.Cells[j, i];

                                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(sourceCell.Borders[(XlBordersIndex)7].LineStyle, XlLineStyle.xlLineStyleNone, false)))
                                        {
                                            targetCell.Borders[(XlBordersIndex)8].LineStyle = sourceCell.Borders[(XlBordersIndex)7].LineStyle;
                                            targetCell.Borders[(XlBordersIndex)8].Color = sourceCell.Borders[(XlBordersIndex)7].Color;
                                            targetCell.Borders[(XlBordersIndex)8].Weight = sourceCell.Borders[(XlBordersIndex)7].Weight;
                                        }
                                        else
                                        {
                                            targetCell.Borders[(XlBordersIndex)8].LineStyle = XlLineStyle.xlLineStyleNone;
                                        }

                                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(sourceCell.Borders[(XlBordersIndex)8].LineStyle, XlLineStyle.xlLineStyleNone, false)))
                                        {
                                            targetCell.Borders[(XlBordersIndex)7].LineStyle = sourceCell.Borders[(XlBordersIndex)8].LineStyle;
                                            targetCell.Borders[(XlBordersIndex)7].Color = sourceCell.Borders[(XlBordersIndex)8].Color;
                                            targetCell.Borders[(XlBordersIndex)7].Weight = sourceCell.Borders[(XlBordersIndex)8].Weight;
                                        }
                                        else
                                        {
                                            targetCell.Borders[(XlBordersIndex)7].LineStyle = XlLineStyle.xlLineStyleNone;
                                        }

                                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(sourceCell.Borders[(XlBordersIndex)9].LineStyle, XlLineStyle.xlLineStyleNone, false)))
                                        {
                                            targetCell.Borders[(XlBordersIndex)10].LineStyle = sourceCell.Borders[(XlBordersIndex)9].LineStyle;
                                            targetCell.Borders[(XlBordersIndex)10].Color = sourceCell.Borders[(XlBordersIndex)9].Color;
                                            targetCell.Borders[(XlBordersIndex)10].Weight = sourceCell.Borders[(XlBordersIndex)9].Weight;
                                        }
                                        else
                                        {
                                            targetCell.Borders[(XlBordersIndex)10].LineStyle = XlLineStyle.xlLineStyleNone;
                                        }

                                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(sourceCell.Borders[(XlBordersIndex)10].LineStyle, XlLineStyle.xlLineStyleNone, false)))
                                        {
                                            targetCell.Borders[(XlBordersIndex)9].LineStyle = sourceCell.Borders[(XlBordersIndex)10].LineStyle;
                                            targetCell.Borders[(XlBordersIndex)9].Color = sourceCell.Borders[(XlBordersIndex)10].Color;
                                            targetCell.Borders[(XlBordersIndex)9].Weight = sourceCell.Borders[(XlBordersIndex)10].Weight;
                                        }
                                        else
                                        {
                                            targetCell.Borders[(XlBordersIndex)9].LineStyle = XlLineStyle.xlLineStyleNone;
                                        }

                                    }
                                }
                            }
                        }

                        else if (RadioButton2.Checked == true)
                        {
                            for (int i = 1, loopTo2 = rng.Rows.Count; i <= loopTo2; i++)
                            {
                                for (int j = 1, loopTo3 = rng.Columns.Count; j <= loopTo3; j++)
                                {
                                    rng2.Cells[j, i].Value = Operators.ConcatenateObject("=", rng.Cells[i, j].Address((object)true, (object)true, XlReferenceStyle.xlA1, (object)true));
                                    if (CheckBox2.Checked == true)
                                    {
                                        rng.Cells[i, j].Copy();
                                        rng2.Cells[j, i].PasteSpecial(XlPasteType.xlPasteFormats);
                                        rng2 = worksheet2.get_Range(rng2Address);

                                        Range sourceCell = (Range)rng.Cells[i, j];
                                        Range targetCell = (Range)rng2.Cells[j, i];

                                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(sourceCell.Borders[(XlBordersIndex)7].LineStyle, XlLineStyle.xlLineStyleNone, false)))
                                        {
                                            targetCell.Borders[(XlBordersIndex)8].LineStyle = sourceCell.Borders[(XlBordersIndex)7].LineStyle;
                                            targetCell.Borders[(XlBordersIndex)8].Color = sourceCell.Borders[(XlBordersIndex)7].Color;
                                            targetCell.Borders[(XlBordersIndex)8].Weight = sourceCell.Borders[(XlBordersIndex)7].Weight;
                                        }
                                        else
                                        {
                                            targetCell.Borders[(XlBordersIndex)8].LineStyle = XlLineStyle.xlLineStyleNone;
                                        }

                                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(sourceCell.Borders[(XlBordersIndex)8].LineStyle, XlLineStyle.xlLineStyleNone, false)))
                                        {
                                            targetCell.Borders[(XlBordersIndex)7].LineStyle = sourceCell.Borders[(XlBordersIndex)8].LineStyle;
                                            targetCell.Borders[(XlBordersIndex)7].Color = sourceCell.Borders[(XlBordersIndex)8].Color;
                                            targetCell.Borders[(XlBordersIndex)7].Weight = sourceCell.Borders[(XlBordersIndex)8].Weight;
                                        }
                                        else
                                        {
                                            targetCell.Borders[(XlBordersIndex)7].LineStyle = XlLineStyle.xlLineStyleNone;
                                        }

                                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(sourceCell.Borders[(XlBordersIndex)9].LineStyle, XlLineStyle.xlLineStyleNone, false)))
                                        {
                                            targetCell.Borders[(XlBordersIndex)10].LineStyle = sourceCell.Borders[(XlBordersIndex)9].LineStyle;
                                            targetCell.Borders[(XlBordersIndex)10].Color = sourceCell.Borders[(XlBordersIndex)9].Color;
                                            targetCell.Borders[(XlBordersIndex)10].Weight = sourceCell.Borders[(XlBordersIndex)9].Weight;
                                        }
                                        else
                                        {
                                            targetCell.Borders[(XlBordersIndex)10].LineStyle = XlLineStyle.xlLineStyleNone;
                                        }

                                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(sourceCell.Borders[(XlBordersIndex)10].LineStyle, XlLineStyle.xlLineStyleNone, false)))
                                        {
                                            targetCell.Borders[(XlBordersIndex)9].LineStyle = sourceCell.Borders[(XlBordersIndex)10].LineStyle;
                                            targetCell.Borders[(XlBordersIndex)9].Color = sourceCell.Borders[(XlBordersIndex)10].Color;
                                            targetCell.Borders[(XlBordersIndex)9].Weight = sourceCell.Borders[(XlBordersIndex)10].Weight;
                                        }
                                        else
                                        {
                                            targetCell.Borders[(XlBordersIndex)9].LineStyle = XlLineStyle.xlLineStyleNone;
                                        }

                                    }
                                }
                            }
                        }

                        excelApp.CutCopyMode = XlCutCopyMode.xlCopy;
                    }

                    else
                    {

                        var Arr = new object[rng.Rows.Count, rng.Columns.Count];

                        for (int i = Information.LBound(Arr, 1), loopTo4 = Information.UBound(Arr, 1); i <= loopTo4; i++)
                        {
                            for (int j = Information.LBound(Arr, 2), loopTo5 = Information.UBound(Arr, 2); j <= loopTo5; j++)
                            {
                                if (RadioButton3.Checked == true)
                                {
                                    Arr[i, j] = rng.Cells[i + 1, j + 1].value;
                                }
                                else if (RadioButton2.Checked == true)
                                {
                                    Arr[i, j] = Operators.ConcatenateObject("=", rng.Cells[i + 1, j + 1].Address((object)true, (object)true, XlReferenceStyle.xlA1, (object)true));
                                }
                            }
                        }

                        var FontNames = new string[rng.Rows.Count, rng.Columns.Count];
                        var FontSizes = new float[rng.Rows.Count, rng.Columns.Count];

                        var Bolds = new bool[rng.Rows.Count, rng.Columns.Count];
                        var Italics = new bool[rng.Rows.Count, rng.Columns.Count];

                        var Reds1 = new int[rng.Rows.Count, rng.Columns.Count];
                        var Reds2 = new int[rng.Rows.Count, rng.Columns.Count];

                        var Greens1 = new int[rng.Rows.Count, rng.Columns.Count];
                        var Greens2 = new int[rng.Rows.Count, rng.Columns.Count];

                        var Blues1 = new int[rng.Rows.Count, rng.Columns.Count];
                        var Blues2 = new int[rng.Rows.Count, rng.Columns.Count];

                        var Borders7 = new bool[rng.Rows.Count, rng.Columns.Count];
                        var Borders8 = new bool[rng.Rows.Count, rng.Columns.Count];
                        var Borders9 = new bool[rng.Rows.Count, rng.Columns.Count];
                        var Borders10 = new bool[rng.Rows.Count, rng.Columns.Count];

                        var Borders7L = new object[rng.Rows.Count, rng.Columns.Count];
                        var Borders8L = new object[rng.Rows.Count, rng.Columns.Count];
                        var Borders9L = new object[rng.Rows.Count, rng.Columns.Count];
                        var Borders10L = new object[rng.Rows.Count, rng.Columns.Count];

                        var Borders7W = new object[rng.Rows.Count, rng.Columns.Count];
                        var Borders8W = new object[rng.Rows.Count, rng.Columns.Count];
                        var Borders9W = new object[rng.Rows.Count, rng.Columns.Count];
                        var Borders10W = new object[rng.Rows.Count, rng.Columns.Count];

                        var Borders7C = new object[rng.Rows.Count, rng.Columns.Count];
                        var Borders8C = new object[rng.Rows.Count, rng.Columns.Count];
                        var Borders9C = new object[rng.Rows.Count, rng.Columns.Count];
                        var Borders10C = new object[rng.Rows.Count, rng.Columns.Count];

                        if (CheckBox2.Checked == true)
                        {

                            for (int i = Information.LBound(Arr, 1), loopTo6 = Information.UBound(Arr, 1); i <= loopTo6; i++)
                            {
                                for (int j = Information.LBound(Arr, 2), loopTo7 = Information.UBound(Arr, 2); j <= loopTo7; j++)
                                {
                                    Range cell = (Range)rng.Cells[i + 1, j + 1];
                                    var font = cell.Font;

                                    if (font.Name is DBNull == false)
                                    {
                                        FontNames[i, j] = Conversions.ToString(font.Name);
                                    }
                                    else
                                    {
                                        FontNames[i, j] = "Calibri";
                                    }

                                    if (font.Size is DBNull == false)
                                    {
                                        float fontSize = Convert.ToSingle(font.Size);
                                        FontSizes[i, j] = fontSize;
                                    }
                                    else
                                    {
                                        FontSizes[i, j] = 11f;
                                    }

                                    Bolds[i, j] = Conversions.ToBoolean(cell.Font.Bold);
                                    Italics[i, j] = Conversions.ToBoolean(cell.Font.Italic);

                                    if (cell.Interior.Color is DBNull)
                                    {
                                        Reds1[i, j] = 255;
                                        Greens1[i, j] = 255;
                                        Blues1[i, j] = 255;
                                    }
                                    else
                                    {
                                        long colorValue1 = Conversions.ToLong(cell.Interior.Color);
                                        int red1 = (int)(colorValue1 % 256L);
                                        int green1 = (int)(colorValue1 / 256L % 256L);
                                        int blue1 = (int)(colorValue1 / 256L / 256L % 256L);
                                        Reds1[i, j] = red1;
                                        Greens1[i, j] = green1;
                                        Blues1[i, j] = blue1;
                                    }

                                    if (cell.Font.Color is DBNull)
                                    {
                                        Reds2[i, j] = 0;
                                        Greens2[i, j] = 0;
                                        Blues2[i, j] = 0;
                                    }
                                    else
                                    {
                                        long colorValue2 = Conversions.ToLong(cell.Font.Color);
                                        int red2 = (int)(colorValue2 % 256L);
                                        int green2 = (int)(colorValue2 / 256L % 256L);
                                        int blue2 = (int)(colorValue2 / 256L / 256L % 256L);
                                        Reds2[i, j] = red2;
                                        Greens2[i, j] = green2;
                                        Blues2[i, j] = blue2;
                                    }

                                    if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(cell.Borders[(XlBordersIndex)7].LineStyle, XlLineStyle.xlLineStyleNone, false)))
                                    {
                                        Borders7[i, j] = true;
                                        Borders7L[i, j] = cell.Borders[(XlBordersIndex)7].LineStyle;
                                        Borders7C[i, j] = cell.Borders[(XlBordersIndex)7].Color;
                                        Borders7W[i, j] = cell.Borders[(XlBordersIndex)7].Weight;
                                    }
                                    else
                                    {
                                        Borders7[i, j] = false;
                                    }

                                    if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(cell.Borders[(XlBordersIndex)8].LineStyle, XlLineStyle.xlLineStyleNone, false)))
                                    {
                                        Borders8[i, j] = true;
                                        Borders8L[i, j] = cell.Borders[(XlBordersIndex)8].LineStyle;
                                        Borders8C[i, j] = cell.Borders[(XlBordersIndex)8].Color;
                                        Borders8W[i, j] = cell.Borders[(XlBordersIndex)8].Weight;
                                    }
                                    else
                                    {
                                        Borders8[i, j] = false;
                                    }

                                    if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(cell.Borders[(XlBordersIndex)9].LineStyle, XlLineStyle.xlLineStyleNone, false)))
                                    {
                                        Borders9[i, j] = true;
                                        Borders9L[i, j] = cell.Borders[(XlBordersIndex)9].LineStyle;
                                        Borders9C[i, j] = cell.Borders[(XlBordersIndex)9].Color;
                                        Borders9W[i, j] = cell.Borders[(XlBordersIndex)9].Weight;
                                    }
                                    else
                                    {
                                        Borders9[i, j] = false;
                                    }

                                    if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(cell.Borders[(XlBordersIndex)10].LineStyle, XlLineStyle.xlLineStyleNone, false)))
                                    {
                                        Borders10[i, j] = true;
                                        Borders10L[i, j] = cell.Borders[(XlBordersIndex)10].LineStyle;
                                        Borders10C[i, j] = cell.Borders[(XlBordersIndex)10].Color;
                                        Borders10W[i, j] = cell.Borders[(XlBordersIndex)10].Weight;
                                    }
                                    else
                                    {
                                        Borders10[i, j] = false;
                                    }

                                }
                            }

                        }

                        rng.ClearContents();
                        rng.ClearFormats();

                        rng2.ClearFormats();

                        for (int i = 1, loopTo8 = rng.Rows.Count; i <= loopTo8; i++)
                        {
                            for (int j = 1, loopTo9 = rng.Columns.Count; j <= loopTo9; j++)
                                rng2.Cells[j, i] = Arr[i - 1, j - 1];
                        }

                        if (CheckBox2.Checked == true)
                        {
                            for (int i = 1, loopTo10 = rng.Rows.Count; i <= loopTo10; i++)
                            {
                                for (int j = 1, loopTo11 = rng.Columns.Count; j <= loopTo11; j++)
                                {
                                    {
                                        ref var withBlock = ref rng2.Cells[j, i].Font;
                                        withBlock.Name = FontNames[i - 1, j - 1];
                                        withBlock.Size = (object)FontSizes[i - 1, j - 1];
                                        withBlock.Bold = (object)Bolds[i - 1, j - 1];
                                        withBlock.Italic = (object)Italics[i - 1, j - 1];
                                    }

                                    int red1 = Reds1[i - 1, j - 1];
                                    int green1 = Greens1[i - 1, j - 1];
                                    int blue1 = Blues1[i - 1, j - 1];
                                    rng2.Cells[j, i].Interior.Color = (object)Color.FromArgb(red1, green1, blue1);

                                    int red2 = Reds2[i - 1, j - 1];
                                    int green2 = Greens2[i - 1, j - 1];
                                    int blue2 = Blues2[i - 1, j - 1];
                                    rng2.Cells[j, i].Font.Color = (object)Color.FromArgb(red2, green2, blue2);

                                    Range targetCell = (Range)rng2.Cells[j, i];
                                    int x = i - 1;
                                    int y = j - 1;

                                    if (Borders7[x, y] == true)
                                    {
                                        targetCell.Borders[(XlBordersIndex)8].LineStyle = Borders7L[x, y];
                                        targetCell.Borders[(XlBordersIndex)8].Color = Borders7C[x, y];
                                        targetCell.Borders[(XlBordersIndex)8].Weight = Borders7W[x, y];
                                    }
                                    else
                                    {
                                        targetCell.Borders[(XlBordersIndex)8].LineStyle = XlLineStyle.xlLineStyleNone;
                                    }

                                    if (Borders8[x, y] == true)
                                    {
                                        targetCell.Borders[(XlBordersIndex)7].LineStyle = Borders8L[x, y];
                                        targetCell.Borders[(XlBordersIndex)7].Color = Borders8C[x, y];
                                        targetCell.Borders[(XlBordersIndex)7].Weight = Borders8W[x, y];
                                    }
                                    else
                                    {
                                        targetCell.Borders[(XlBordersIndex)7].LineStyle = XlLineStyle.xlLineStyleNone;
                                    }

                                    if (Borders9[x, y] == true)
                                    {
                                        targetCell.Borders[(XlBordersIndex)10].LineStyle = Borders9L[x, y];
                                        targetCell.Borders[(XlBordersIndex)10].Color = Borders9C[x, y];
                                        targetCell.Borders[(XlBordersIndex)10].Weight = Borders9W[x, y];
                                    }
                                    else
                                    {
                                        targetCell.Borders[(XlBordersIndex)10].LineStyle = XlLineStyle.xlLineStyleNone;
                                    }

                                    if (Borders10[x, y] == true)
                                    {
                                        targetCell.Borders[(XlBordersIndex)9].LineStyle = Borders10L[x, y];
                                        targetCell.Borders[(XlBordersIndex)9].Color = Borders10C[x, y];
                                        targetCell.Borders[(XlBordersIndex)9].Weight = Borders10W[x, y];
                                    }
                                    else
                                    {
                                        targetCell.Borders[(XlBordersIndex)9].LineStyle = XlLineStyle.xlLineStyleNone;
                                    }

                                }
                            }

                        }

                    }

                    rng2.Select();

                    for (int j = 1, loopTo12 = rng2.Columns.Count; j <= loopTo12; j++)
                        rng2.Columns[j].Autofit();

                    TextBoxChanged = false;

                    Close();

                }
            }

            catch (Exception ex)
            {

            }

        }

        private void btn_OK_MouseEnter(object sender, EventArgs e)
        {

            try
            {

                btn_OK.ForeColor = Color.White;
                btn_OK.BackColor = Color.FromArgb(76, 111, 174);
            }

            catch (Exception ex)
            {

            }

        }

        private void RadioButton3_CheckedChanged(object sender, EventArgs e)
        {

            try
            {
                if (RadioButton3.Checked == true)
                {
                    Display();
                }
            }
            catch (Exception ex)
            {

            }

        }

        private void RadioButton2_CheckedChanged(object sender, EventArgs e)
        {

            try
            {
                if (RadioButton2.Checked == true)
                {
                    Display();
                }
            }
            catch (Exception ex)
            {

            }

        }

        private void CheckBox2_CheckedChanged(object sender, EventArgs e)
        {

            try
            {
                Display();
            }
            catch (Exception ex)
            {

            }

        }

        private void TextBox1_TextChanged(object sender, EventArgs e)
        {

            try
            {
                if (!string.IsNullOrEmpty(TextBox1.Text) & Form4Open == 0)
                {
                    worksheet = (Excel.Worksheet)workbook.ActiveSheet;
                    string[] rngArray = Strings.Split(TextBox1.Text, "!");
                    string rngAddress = rngArray[Information.UBound(rngArray)];
                    rng = worksheet.get_Range(rngAddress);
                    TextBoxChanged = true;
                    rng.Select();
                    Display();
                    TextBoxChanged = false;
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
                    worksheet2 = (Excel.Worksheet)workbook.ActiveSheet;
                    string[] rng2Array = Strings.Split(TextBox2.Text, "!");
                    string rng2Address = rng2Array[Information.UBound(rng2Array)];
                    rng2 = worksheet2.get_Range(rng2Address);
                    TextBoxChanged = true;
                    rng2.Select();
                    TextBoxChanged = false;
                }
            }

            catch (Exception ex)
            {

            }

        }

        private void Form3_Load(object sender, EventArgs e)
        {

            try
            {

                excelApp.SheetSelectionChange += excelApp_SheetSelectionChange;

                Opened = Opened + 1;
                KeyPreview = true;
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

                if (TextBoxChanged == false)
                {
                    if (FocusedTextBox == 1)
                    {
                        worksheet = (Excel.Worksheet)workbook.ActiveSheet;
                        if ((worksheet.Name ?? "") != (OpenSheet.Name ?? ""))
                        {
                            TextBox1.Text = worksheet.Name + "!" + selectedRange.get_Address();
                        }
                        else
                        {
                            TextBox1.Text = selectedRange.get_Address();
                        }
                        rng = selectedRange;
                        TextBox1.Focus();
                    }

                    else if (FocusedTextBox == 2)
                    {
                        worksheet2 = (Excel.Worksheet)workbook.ActiveSheet;
                        if ((worksheet2.Name ?? "") != (OpenSheet.Name ?? ""))
                        {
                            TextBox2.Text = worksheet2.Name + "!" + selectedRange.get_Address();
                        }
                        else
                        {
                            TextBox2.Text = selectedRange.get_Address();
                        }
                        rng2 = selectedRange;
                        TextBox2.Focus();
                    }
                }
            }

            catch (Exception ex)
            {

            }

        }

        private void ComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

            try
            {
                if (Conversions.ToBoolean(Operators.AndObject(Operators.ConditionalCompareObjectEqual(ComboBox1.SelectedItem, "SOFTEKO", false), Opened >= 1)))
                {

                    string url = "https://www.softeko.co";
                    Process.Start(url);

                }
            }
            catch (Exception ex)
            {

            }

        }

        private void RadioButton4_CheckedChanged(object sender, EventArgs e)
        {

            try
            {

                if (RadioButton4.Checked == true)
                {
                    DestinationChange();
                }
            }

            catch (Exception ex)
            {

            }

        }

        private void RadioButton1_CheckedChanged_1(object sender, EventArgs e)
        {

            try
            {
                if (RadioButton1.Checked == true)
                {
                    DestinationChange();
                }
            }

            catch (Exception ex)
            {

            }

        }

        private void RadioButton5_CheckedChanged(object sender, EventArgs e)
        {

            try
            {

                if (RadioButton5.Checked == true)
                {
                    DestinationChange();
                }
            }

            catch (Exception ex)
            {

            }

        }

        private void PictureBox8_GotFocus(object sender, EventArgs e)
        {

            try
            {
                FocusedTextBox = 1;
            }
            catch (Exception ex)
            {

            }

        }

        private void TextBox1_GotFocus(object sender, EventArgs e)
        {

            try
            {
                FocusedTextBox = 1;
            }
            catch (Exception ex)
            {

            }

        }

        private void PictureBox4_GotFocus(object sender, EventArgs e)
        {

            try
            {
                FocusedTextBox = 1;
            }
            catch (Exception ex)
            {

            }

        }

        private void TextBox2_GotFocus(object sender, EventArgs e)
        {

            try
            {
                FocusedTextBox = 2;
            }
            catch (Exception ex)
            {

            }

        }

        private void PictureBox2_GotFocus(object sender, EventArgs e)
        {

            try
            {
                FocusedTextBox = 2;
            }
            catch (Exception ex)
            {

            }

        }

        private void RadioButton3_GotFocus(object sender, EventArgs e)
        {

            try
            {
                FocusedTextBox = 0;
            }
            catch (Exception ex)
            {

            }

        }

        private void RadioButton2_GotFocus(object sender, EventArgs e)
        {

            try
            {
                FocusedTextBox = 0;
            }
            catch (Exception ex)
            {

            }

        }

        private void PictureBox5_GotFocus(object sender, EventArgs e)
        {
            try
            {
                FocusedTextBox = 0;
            }
            catch (Exception ex)
            {

            }

        }

        private void PictureBox1_GotFocus(object sender, EventArgs e)
        {

            try
            {
                FocusedTextBox = 0;
            }
            catch (Exception ex)
            {

            }

        }

        private void CustomGroupBox3_GotFocus(object sender, EventArgs e)
        {

            try
            {
                FocusedTextBox = 0;
            }
            catch (Exception ex)
            {

            }

        }


        private void CustomGroupBox6_GotFocus(object sender, EventArgs e)
        {

            try
            {
                FocusedTextBox = 0;
            }
            catch (Exception ex)
            {

            }

        }

        private void CustomGroupBox5_GotFocus(object sender, EventArgs e)
        {

            try
            {
                FocusedTextBox = 0;
            }
            catch (Exception ex)
            {

            }

        }

        private void RadioButton1_GotFocus(object sender, EventArgs e)
        {
            try
            {
                FocusedTextBox = 0;
            }
            catch (Exception ex)
            {

            }
        }

        private void RadioButton4_GotFocus(object sender, EventArgs e)
        {
            try
            {
                FocusedTextBox = 0;
            }

            catch (Exception ex)
            {

            }

        }

        private void RadioButton5_GotFocus(object sender, EventArgs e)
        {

            try
            {
                FocusedTextBox = 0;
            }
            catch (Exception ex)
            {

            }

        }

        private void CheckBox2_GotFocus(object sender, EventArgs e)
        {

            try
            {
                FocusedTextBox = 0;
            }
            catch (Exception ex)
            {

            }

        }

        private void CheckBox1_GotFocus(object sender, EventArgs e)
        {
            try
            {
                FocusedTextBox = 0;
            }
            catch (Exception ex)
            {

            }

        }

        private void ComboBox1_GotFocus(object sender, EventArgs e)
        {
            try
            {
                FocusedTextBox = 0;
            }
            catch (Exception ex)
            {

            }

        }

        private void CustomGroupBox1_GotFocus(object sender, EventArgs e)
        {
            try
            {
                FocusedTextBox = 0;
            }
            catch (Exception ex)
            {

            }
        }

        private void CustomGroupBox2_GotFocus(object sender, EventArgs e)
        {
            try
            {
                FocusedTextBox = 0;
            }
            catch (Exception ex)
            {

            }

        }

        private void CustomGroupBox4_GotFocus(object sender, EventArgs e)
        {
            try
            {
                FocusedTextBox = 0;
            }
            catch (Exception ex)
            {

            }
        }

        private void panel1_GotFocus(object sender, EventArgs e)
        {
            try
            {
                FocusedTextBox = 0;
            }
            catch (Exception ex)
            {

            }
        }

        private void panel2_GotFocus(object sender, EventArgs e)
        {
            try
            {
                FocusedTextBox = 0;
            }
            catch (Exception ex)
            {

            }
        }

        private void PictureBox7_GotFocus(object sender, EventArgs e)
        {
            try
            {
                FocusedTextBox = 0;
            }
            catch (Exception ex)
            {

            }
        }

        private void btn_OK_GotFocus(object sender, EventArgs e)
        {
            try
            {
                FocusedTextBox = 0;
            }
            catch (Exception ex)
            {

            }
        }

        private void btn_cancel_GotFocus(object sender, EventArgs e)
        {

            try
            {
                FocusedTextBox = 0;
            }
            catch (Exception ex)
            {

            }

        }

        private void btn_cancel_Click(object sender, EventArgs e)
        {

            try
            {
                Close();
            }
            catch (Exception ex)
            {

            }
        }

        private void btn_cancel_MouseEnter(object sender, EventArgs e)
        {

            try
            {

                btn_cancel.ForeColor = Color.White;
                btn_cancel.BackColor = Color.FromArgb(76, 111, 174);
            }

            catch (Exception ex)
            {

            }

        }

        private void Form3_Closing(object sender, CancelEventArgs e)
        {

            try
            {
                GlobalModule.form_flag = false;
            }

            catch (Exception ex)
            {

            }

        }

        private void Form3_Shown(object sender, EventArgs e)
        {

            try
            {
                Focus();
                BringToFront();
                Activate();

                string TextBoxText;

                if ((worksheet.Name ?? "") != (OpenSheet.Name ?? ""))
                {
                    TextBoxText = worksheet.Name + "!" + rng.get_Address();
                }
                else
                {
                    TextBoxText = rng.get_Address();
                }

                BeginInvoke(new System.Action(() =>
                    {
                        TextBox1.Text = TextBoxText;
                        SetWindowPos(Handle, new IntPtr(HWND_TOPMOST), 0, 0, 0, 0, SWP_NOACTIVATE | SWP_NOMOVE | SWP_NOSIZE);
                    }));
            }
            catch (Exception ex)
            {

            }

        }

        private void Form3_Disposed(object sender, EventArgs e)
        {
            try
            {
                GlobalModule.form_flag = false;
            }
            catch (Exception ex)
            {

            }
        }

        private void Form3_KeyDown(object sender, KeyEventArgs e)
        {


            try
            {

                if (e.KeyCode == Keys.Enter)
                {
                    btn_OK.Focus();
                    btn_OK_Click(sender, e);
                }
            }

            catch (Exception ex)
            {

            }

        }

    }
}