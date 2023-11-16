using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace VSTO_Addins
{

    public partial class Form10
    {

        private Excel.Application _excelApp;

        private Excel.Application excelApp
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

        private Excel.Workbook workBook;
        private Excel.Workbook workbook2;

        private Excel.Worksheet workSheet;
        private Excel.Worksheet workSheet2;

        private Range rng;
        private Range rng2;

        private int opened;
        private int FocusedTextBox;

        public Form10()
        {
            InitializeComponent();
        }


        [DllImport("user32")]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
        private const uint SWP_NOMOVE = 0x2U;
        private const uint SWP_NOSIZE = 0x1U;
        private const uint SWP_NOACTIVATE = 0x10U;
        private const int HWND_TOPMOST = -1;


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
        private bool IsValidExcelCellReference(string cellReference)
        {

            string cellPattern = @"(\$?[A-Z]+\$?[0-9]+)";

            string referencePattern = "^" + cellPattern + "(:" + cellPattern + ")?$";

            var regex = new Regex(referencePattern);

            if (regex.IsMatch(cellReference))
            {
                return true;
            }
            else
            {
                return false;
            }

        }
        private object SearchInArray(object i, object j, object Arr)
        {
            object SearchInArrayRet = default;

            object Result = 0;

            for (int k = Information.LBound((Array)Arr, 1), loopTo = Information.UBound((Array)Arr, 1); k <= loopTo; k++)
            {
                if (Conversions.ToBoolean(Operators.AndObject(Operators.ConditionalCompareObjectEqual(Arr((object)k, (object)0), i, false), Operators.ConditionalCompareObjectEqual(Arr((object)k, (object)1), j, false))))
                {
                    Result = Arr((object)k, (object)2);
                    break;
                }
            }

            SearchInArrayRet = Result;
            return SearchInArrayRet;

        }
        private object Available(object i, object j, object Arr)
        {
            object AvailableRet = default;

            bool Result = false;

            for (int k = Information.LBound((Array)Arr, 1), loopTo = Information.UBound((Array)Arr, 1); k <= loopTo; k++)
            {
                if (Conversions.ToBoolean(Operators.AndObject(Operators.ConditionalCompareObjectEqual(Arr((object)k, (object)0), i, false), Operators.ConditionalCompareObjectEqual(Arr((object)k, (object)1), j, false))))
                {
                    Result = true;
                    break;
                }
            }

            AvailableRet = Result;
            return AvailableRet;

        }
        private void Display()
        {

            try
            {
                CustomPanel1.Controls.Clear();
                CustomPanel2.Controls.Clear();


                Range displayRng;

                if (rng.Rows.Count > 50)
                {
                    displayRng = workSheet.get_Range(rng.Cells[1, 1], rng.Cells[50, rng.Columns.Count]);
                }
                else
                {
                    displayRng = workSheet.get_Range(rng.Cells[1, 1], rng.Cells[rng.Rows.Count, rng.Columns.Count]);
                }

                int r;
                int c;

                r = displayRng.Rows.Count;
                c = displayRng.Columns.Count;

                float height;
                float width;

                if (r <= 6)
                {
                    height = (float)(CustomPanel1.Height / (double)r);
                }
                else
                {
                    height = (float)(CustomPanel1.Height / 6d);
                }

                if (c <= 4)
                {
                    width = (float)(CustomPanel1.Width / (double)c);
                }
                else
                {
                    width = (float)(CustomPanel1.Width / 4d);
                }

                var Arr = new object[(r * c), 2];

                int Count = 0;

                for (int i = 1, loopTo = r; i <= loopTo; i++)
                {
                    for (int j = 1, loopTo1 = c; j <= loopTo1; j++)
                    {
                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(Available(i, j, Arr), false, false)))
                        {
                            if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(displayRng.Cells[i, j].MergeCells, true, false)))
                            {
                                for (int k = 2, loopTo2 = Conversions.ToInteger(displayRng.Cells[i, j].MergeArea.Columns.Count); k <= loopTo2; k++)
                                {
                                    if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(Available(i, Operators.SubtractObject(Operators.AddObject(j, k), 1), Arr), false, false)))
                                    {
                                        Arr[Count, 0] = i;
                                        Arr[Count, 1] = Operators.SubtractObject(Operators.AddObject(j, k), 1);
                                        Count = Count + 1;
                                    }
                                }
                                for (int m = 2, loopTo3 = Conversions.ToInteger(displayRng.Cells[i, j].MergeArea.Rows.Count); m <= loopTo3; m++)
                                {
                                    for (int n = 1, loopTo4 = Conversions.ToInteger(displayRng.Cells[i, j].MergeArea.Columns.Count); n <= loopTo4; n++)
                                    {
                                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(Available(Operators.SubtractObject(Operators.AddObject(i, m), 1), Operators.SubtractObject(Operators.AddObject(j, n), 1), Arr), false, false)))
                                        {
                                            Arr[Count, 0] = Operators.SubtractObject(Operators.AddObject(i, m), 1);
                                            Arr[Count, 1] = Operators.SubtractObject(Operators.AddObject(j, n), 1);
                                            Count = Count + 1;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                for (int i = 1, loopTo5 = r; i <= loopTo5; i++)
                {
                    for (int j = 1, loopTo6 = c; j <= loopTo6; j++)
                    {
                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(Available(i, j, Arr), false, false)))
                        {
                            float height2 = Conversions.ToSingle(Operators.MultiplyObject(height, displayRng.Cells[i, j].MergeArea.Rows.Count));
                            float width2 = Conversions.ToSingle(Operators.MultiplyObject(width, displayRng.Cells[i, j].MergeArea.Columns.Count));
                            var label = new System.Windows.Forms.Label();
                            label.Text = Conversions.ToString(displayRng.Cells[i, j].Value);
                            label.Location = new System.Drawing.Point((int)Math.Round((j - 1) * width), (int)Math.Round((i - 1) * height));
                            label.Height = (int)Math.Round(height2);
                            label.Width = (int)Math.Round(width2);
                            label.BorderStyle = BorderStyle.FixedSingle;
                            label.TextAlign = ContentAlignment.MiddleCenter;

                            if (CheckBox1.Checked == true)
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

                            CustomPanel1.Controls.Add(label);
                        }
                    }
                }
                CustomPanel1.AutoScroll = true;

                var Arr2 = new object[(r * c), 3];

                Count = 0;

                for (int i = 1, loopTo7 = r; i <= loopTo7; i++)
                {
                    for (int j = 1, loopTo8 = c; j <= loopTo8; j++)
                    {
                        if (Conversions.ToBoolean(Operators.AndObject(Operators.ConditionalCompareObjectEqual(rng.Cells[i, j].MergeCells, true, false), Operators.ConditionalCompareObjectEqual(Available(i, j, Arr2), false, false))))
                        {
                            for (int m = 0, loopTo9 = Conversions.ToInteger(Operators.SubtractObject(rng.Cells[i, j].MergeArea.Rows.Count, 1)); m <= loopTo9; m++)
                            {
                                for (int n = 0, loopTo10 = Conversions.ToInteger(Operators.SubtractObject(rng.Cells[i, j].MergeArea.Columns.Count, 1)); n <= loopTo10; n++)
                                {
                                    Arr2[Count, 0] = Operators.AddObject(i, m);
                                    Arr2[Count, 1] = Operators.AddObject(j, n);
                                    Arr2[Count, 2] = displayRng.Cells[i, j].Value;
                                    Count = Count + 1;
                                }
                            }
                        }
                    }
                }

                for (int i = 1, loopTo11 = r; i <= loopTo11; i++)
                {
                    for (int j = 1, loopTo12 = c; j <= loopTo12; j++)
                    {
                        var label = new System.Windows.Forms.Label();
                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(rng.Cells[i, j].MergeCells, true, false)))
                        {
                            label.Text = Conversions.ToString(SearchInArray(i, j, Arr2));
                        }
                        else
                        {
                            label.Text = Conversions.ToString(displayRng.Cells[i, j].Value);
                        }
                        label.Location = new System.Drawing.Point((int)Math.Round((j - 1) * width), (int)Math.Round((i - 1) * height));
                        label.Height = (int)Math.Round(height);
                        label.Width = (int)Math.Round(width);
                        label.BorderStyle = BorderStyle.FixedSingle;
                        label.TextAlign = ContentAlignment.MiddleCenter;

                        if (CheckBox1.Checked == true)
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
                        CustomPanel2.Controls.Add(label);
                    }
                }

                CustomPanel2.AutoScroll = true;
            }

            catch (Exception ex)
            {

            }

        }

        private void Button2_Click(object sender, EventArgs e)
        {

            try
            {

                if (string.IsNullOrEmpty(TextBox1.Text))
                {
                    MessageBox.Show("Select a Source Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    TextBox1.Focus();
                    workSheet.Activate();
                    rng.Select();
                    return;
                }

                if (IsValidExcelCellReference(TextBox1.Text) == false)
                {
                    MessageBox.Show("Enter a Valid Source Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    TextBox1.Focus();
                    workSheet.Activate();
                    rng.Select();
                    return;
                }

                if (RadioButton9.Checked == false & RadioButton10.Checked == false)
                {
                    MessageBox.Show("Select a Destination Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    workSheet.Activate();
                    rng.Select();
                    return;
                }

                if (RadioButton10.Checked == true & (string.IsNullOrEmpty(TextBox3.Text) | IsValidExcelCellReference(TextBox3.Text) == false))
                {
                    MessageBox.Show("Enter a Valid Destination Cell.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    TextBox3.Focus();
                    workSheet.Activate();
                    rng.Select();
                    return;
                }

                if (CheckBox2.Checked == true)
                {
                    workSheet.Copy(After: workBook.Sheets[workSheet.Name]);
                    workSheet2.Activate();
                }

                rng2 = workSheet2.get_Range(rng2.Cells[1, 1], rng2.Cells[rng.Rows.Count, rng.Columns.Count]);
                workSheet2.Activate();

                if (Overlap(excelApp, workSheet, workSheet2, rng, rng2) == true)
                {
                    rng2 = rng;
                }
                else
                {
                    rng.Copy();
                    rng2.PasteSpecial(XlPasteType.xlPasteValues);
                    rng2.PasteSpecial(XlPasteType.xlPasteFormats);
                    excelApp.CutCopyMode = XlCutCopyMode.xlCopy;
                }

                rng2.Select();

                int r = rng2.Rows.Count;
                int C = rng2.Columns.Count;

                for (int i = 1, loopTo = r; i <= loopTo; i++)
                {
                    for (int j = 1, loopTo1 = C; j <= loopTo1; j++)
                    {
                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(rng2.Cells[i, j].MergeCells, true, false)))
                        {
                            int Merged_Rows = Conversions.ToInteger(rng2.Cells[i, j].MergeArea.Rows.Count);
                            int Merged_Columns = Conversions.ToInteger(rng2.Cells[i, j].MergeArea.Columns.Count);
                            rng2.Cells[i, j].UnMerge();
                            for (int m = 0, loopTo2 = Merged_Rows - 1; m <= loopTo2; m++)
                            {
                                for (int n = 0, loopTo3 = Merged_Columns - 1; n <= loopTo3; n++)
                                    rng2.Cells[i + m, j + n].value = rng2.Cells[i, j].value;
                            }
                        }
                    }
                }

                if (CheckBox1.Checked == false)
                {
                    rng2.ClearFormats();
                }

                Close();
            }

            catch (Exception ex)
            {

            }

        }

        private void PictureBox1_Click(object sender, EventArgs e)
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

                workSheet = (Excel.Worksheet)workBook.Worksheets[sheetName];
                workSheet.Activate();

                rng.Select();

                rng = excelApp.get_Range(rng, rng.get_End(XlDirection.xlDown));
                rng = excelApp.get_Range(rng, rng.get_End(XlDirection.xlToRight));

                rng.Select();
                TextBox1.Text = rng.get_Address();

                Show();
                TextBox1.Focus();
            }

            catch (Exception ex)
            {

                Show();
                TextBox1.Focus();

            }

        }

        private void PictureBox9_Click(object sender, EventArgs e)
        {

            try
            {
                FocusedTextBox = 1;

                Range userInput = (Range)excelApp.InputBox("Select a range", Type: 8);
                var rng = userInput;

                try
                {
                    string sheetName;
                    sheetName = Strings.Split(rng.get_Address(true, true, XlReferenceStyle.xlA1, true), "]")[1];
                    sheetName = Strings.Split(sheetName, "!")[0];

                    if (Strings.Mid(sheetName, Strings.Len(sheetName), 1) == "'")
                    {
                        sheetName = Strings.Mid(sheetName, 1, Strings.Len(sheetName) - 1);
                    }

                    workSheet = (Excel.Worksheet)workBook.Worksheets[sheetName];
                    workSheet.Activate();
                }
                catch (Exception ex)
                {

                }

                rng.Select();

                TextBox1.Text = rng.get_Address();
                TextBox1.Focus();
            }

            catch (Exception ex)
            {

            }

        }

        private void Form10_Load(object sender, EventArgs e)
        {

            try
            {

                excelApp = Globals.ThisAddIn.Application;
                workBook = excelApp.ActiveWorkbook;
                workbook2 = excelApp.ActiveWorkbook;
                workSheet = (Excel.Worksheet)workBook.ActiveSheet;
                workSheet2 = (Excel.Worksheet)workbook2.ActiveSheet;

                excelApp.SheetSelectionChange += excelApp_SheetSelectionChange;

                opened = opened + 1;

                Label3.Enabled = false;
                TextBox3.Enabled = false;
                PictureBox6.Enabled = false;
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

                if (FocusedTextBox == 1)
                {
                    TextBox1.Text = selectedRange.get_Address();
                    workSheet = (Excel.Worksheet)workBook.ActiveSheet;
                    rng = selectedRange;
                    TextBox1.Focus();
                }

                else if (FocusedTextBox == 3)
                {
                    TextBox3.Text = selectedRange.get_Address();
                    workSheet2 = (Excel.Worksheet)workbook2.ActiveSheet;
                    rng2 = selectedRange;
                    TextBox3.Focus();
                }
            }

            catch (Exception ex)
            {

            }

        }

        private void CheckBox1_CheckedChanged(object sender, EventArgs e)
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
                excelApp = Globals.ThisAddIn.Application;
                workBook = excelApp.ActiveWorkbook;
                workSheet = (Excel.Worksheet)workBook.ActiveSheet;

                TextBox1.SelectionStart = TextBox1.Text.Length;
                TextBox1.ScrollToCaret();

                rng = workSheet.get_Range(TextBox1.Text);
                rng.Select();

                Display();
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

        private void RadioButton9_CheckedChanged(object sender, EventArgs e)
        {

            try
            {
                if (RadioButton9.Checked == true)
                {
                    workSheet2 = workSheet;
                    rng2 = rng;
                }
            }
            catch (Exception ex)
            {

            }

        }

        private void RadioButton10_CheckedChanged(object sender, EventArgs e)
        {

            try
            {
                if (RadioButton10.Checked == true)
                {
                    Label3.Enabled = true;
                    TextBox3.Enabled = true;
                    TextBox3.Focus();
                    PictureBox6.Enabled = true;
                }
                else
                {
                    Label3.Enabled = false;
                    TextBox3.Clear();
                    TextBox3.Enabled = false;
                    PictureBox6.Enabled = false;
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
                workSheet2 = (Excel.Worksheet)workbook2.ActiveSheet;

                TextBox3.SelectionStart = TextBox3.Text.Length;
                TextBox3.ScrollToCaret();

                rng2 = workSheet2.get_Range(TextBox3.Text);
                rng2.Select();
            }

            catch (Exception ex)
            {

            }

        }

        private void PictureBox6_Click(object sender, EventArgs e)
        {

            try
            {
                FocusedTextBox = 3;
                Hide();

                excelApp = Globals.ThisAddIn.Application;
                workBook = excelApp.ActiveWorkbook;

                Range userInput = (Range)excelApp.InputBox("Select a range", Type: 8);
                rng2 = userInput;


                string sheetName;
                sheetName = Strings.Split(rng2.get_Address(true, true, XlReferenceStyle.xlA1, true), "]")[1];
                sheetName = Strings.Split(sheetName, "!")[0];

                if (Strings.Mid(sheetName, Strings.Len(sheetName), 1) == "'")
                {
                    sheetName = Strings.Mid(sheetName, 1, Strings.Len(sheetName) - 1);
                }

                workSheet2 = (Excel.Worksheet)workBook.Worksheets[sheetName];
                workSheet2.Activate();

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

        private void TextBox3_GotFocus(object sender, EventArgs e)
        {

            try
            {
                FocusedTextBox = 3;
            }
            catch (Exception ex)
            {

            }

        }

        private void ComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {


            try
            {
                if (Conversions.ToBoolean(Operators.AndObject(Operators.ConditionalCompareObjectEqual(ComboBox1.SelectedItem, "SOFTEKO", false), opened >= 1)))
                {

                    string url = "https://www.softeko.co";
                    Process.Start(url);

                }
            }
            catch (Exception ex)
            {

            }

        }

        private void PictureBox1_GotFocus(object sender, EventArgs e)
        {
            try
            {
                FocusedTextBox = 1;
            }
            catch (Exception ex)
            {

            }
        }

        private void PictureBox9_GotFocus(object sender, EventArgs e)
        {
            try
            {
                FocusedTextBox = 1;
            }
            catch (Exception ex)
            {

            }
        }

        private void PictureBox6_GotFocus(object sender, EventArgs e)
        {
            try
            {
                FocusedTextBox = 3;
            }
            catch (Exception ex)
            {

            }
        }

        private void Button1_GotFocus(object sender, EventArgs e)
        {
            try
            {
                FocusedTextBox = 0;
            }
            catch (Exception ex)
            {

            }
        }

        private void Button2_GotFocus(object sender, EventArgs e)
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

        private void CustomGroupBox10_GotFocus(object sender, EventArgs e)
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

        private void CustomPanel1_GotFocus(object sender, EventArgs e)
        {
            try
            {
                FocusedTextBox = 0;
            }
            catch (Exception ex)
            {

            }
        }

        private void CustomPanel2_GotFocus(object sender, EventArgs e)
        {
            try
            {
                FocusedTextBox = 0;
            }
            catch (Exception ex)
            {

            }
        }

        private void Label1_GotFocus(object sender, EventArgs e)
        {
            try
            {
                FocusedTextBox = 0;
            }
            catch (Exception ex)
            {

            }
        }

        private void Label3_GotFocus(object sender, EventArgs e)
        {
            try
            {
                FocusedTextBox = 0;
            }
            catch (Exception ex)
            {

            }
        }

        private void RadioButton10_GotFocus(object sender, EventArgs e)
        {
            try
            {
                FocusedTextBox = 0;
            }
            catch (Exception ex)
            {

            }
        }

        private void RadioButton9_GotFocus(object sender, EventArgs e)
        {
            try
            {
                FocusedTextBox = 0;
            }
            catch (Exception ex)
            {

            }
        }

        private void Button1_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.KeyCode == Keys.Enter)
                {

                    Button2_Click(sender, e);

                }
            }

            catch (Exception ex)
            {

            }
        }

        private void Button2_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.KeyCode == Keys.Enter)
                {

                    Button2_Click(sender, e);

                }
            }

            catch (Exception ex)
            {

            }
        }

        private void CheckBox1_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.KeyCode == Keys.Enter)
                {

                    Button2_Click(sender, e);

                }
            }

            catch (Exception ex)
            {

            }
        }

        private void CheckBox2_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.KeyCode == Keys.Enter)
                {

                    Button2_Click(sender, e);

                }
            }

            catch (Exception ex)
            {

            }
        }

        private void ComboBox1_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.KeyCode == Keys.Enter)
                {

                    Button2_Click(sender, e);

                }
            }

            catch (Exception ex)
            {

            }
        }

        private void CustomGroupBox10_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.KeyCode == Keys.Enter)
                {

                    Button2_Click(sender, e);

                }
            }

            catch (Exception ex)
            {

            }
        }

        private void CustomGroupBox4_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.KeyCode == Keys.Enter)
                {

                    Button2_Click(sender, e);

                }
            }

            catch (Exception ex)
            {

            }
        }

        private void CustomGroupBox5_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.KeyCode == Keys.Enter)
                {

                    Button2_Click(sender, e);

                }
            }

            catch (Exception ex)
            {

            }
        }

        private void CustomGroupBox6_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.KeyCode == Keys.Enter)
                {

                    Button2_Click(sender, e);

                }
            }

            catch (Exception ex)
            {

            }
        }

        private void CustomPanel1_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.KeyCode == Keys.Enter)
                {

                    Button2_Click(sender, e);

                }
            }

            catch (Exception ex)
            {

            }
        }

        private void CustomPanel2_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.KeyCode == Keys.Enter)
                {

                    Button2_Click(sender, e);

                }
            }

            catch (Exception ex)
            {

            }
        }

        private void Label1_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.KeyCode == Keys.Enter)
                {

                    Button2_Click(sender, e);

                }
            }

            catch (Exception ex)
            {

            }
        }

        private void Label3_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.KeyCode == Keys.Enter)
                {

                    Button2_Click(sender, e);

                }
            }

            catch (Exception ex)
            {

            }
        }

        private void PictureBox1_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.KeyCode == Keys.Enter)
                {

                    Button2_Click(sender, e);

                }
            }

            catch (Exception ex)
            {

            }
        }

        private void PictureBox6_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.KeyCode == Keys.Enter)
                {

                    Button2_Click(sender, e);

                }
            }

            catch (Exception ex)
            {

            }
        }

        private void PictureBox9_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.KeyCode == Keys.Enter)
                {

                    Button2_Click(sender, e);

                }
            }

            catch (Exception ex)
            {

            }
        }

        private void RadioButton10_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.KeyCode == Keys.Enter)
                {

                    Button2_Click(sender, e);

                }
            }

            catch (Exception ex)
            {

            }
        }

        private void RadioButton9_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.KeyCode == Keys.Enter)
                {

                    Button2_Click(sender, e);

                }
            }

            catch (Exception ex)
            {

            }
        }

        private void TextBox1_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.KeyCode == Keys.Enter)
                {

                    Button2_Click(sender, e);

                }
            }

            catch (Exception ex)
            {

            }
        }

        private void TextBox3_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.KeyCode == Keys.Enter)
                {

                    Button2_Click(sender, e);

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
                Button1.BackColor = Color.FromArgb(65, 105, 225);
                Button1.ForeColor = Color.FromArgb(255, 255, 255);
            }
            catch (Exception ex)
            {

            }
        }

        private void Button2_MouseEnter(object sender, EventArgs e)
        {
            try
            {
                Button2.BackColor = Color.FromArgb(65, 105, 225);
                Button2.ForeColor = Color.FromArgb(255, 255, 255);
            }
            catch (Exception ex)
            {

            }
        }

        private void Button1_MouseLeave(object sender, EventArgs e)
        {
            try
            {

                Button1.BackColor = Color.FromArgb(255, 255, 255);
                Button1.ForeColor = Color.FromArgb(70, 70, 70);
            }
            catch (Exception ex)
            {

            }
        }

        private void Button2_MouseLeave(object sender, EventArgs e)
        {
            try
            {

                Button2.BackColor = Color.FromArgb(255, 255, 255);
                Button2.ForeColor = Color.FromArgb(70, 70, 70);
            }
            catch (Exception ex)
            {

            }
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            try
            {
                Close();
            }
            catch (Exception ex)
            {

            }
        }

        private void Form10_Closing(object sender, CancelEventArgs e)
        {
            GlobalModule.form_flag = false;
        }

        private void Form10_Disposed(object sender, EventArgs e)
        {
            GlobalModule.form_flag = false;
        }

        private void Form10_Shown(object sender, EventArgs e)
        {
            Focus();
            BringToFront();
            Activate();
            BeginInvoke(new System.Action(() =>
                {
                    TextBox1.Text = rng.get_Address();
                    SetWindowPos(Handle, new IntPtr(HWND_TOPMOST), 0, 0, 0, 0, SWP_NOACTIVATE | SWP_NOMOVE | SWP_NOSIZE);
                }));
        }
    }
}