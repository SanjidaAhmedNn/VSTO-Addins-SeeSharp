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

    public partial class Form18_CombineRanges
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
        private Excel.Worksheet workSheet;
        private Range rng;
        private Range selectedRange;

        private int opened;
        private int FocusedTextBox;
        private bool TextBoxChanged;

        public Form18_CombineRanges()
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

            // Checks whether a string is a valid cell reference or not.

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
        private object Operation(object Arr, object Flag)
        {
            object OperationRet = default;

            // Takes an array of numbers and conduct mathematical operations. The operation name is input as flag in the format "=...".

            if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(Flag, "=SUM()", false)))
            {
                double Output = 0d;
                for (int i = Information.LBound((Array)Arr), loopTo = Information.UBound((Array)Arr); i <= loopTo; i++)
                {
                    if (Information.IsNumeric(Arr((object)i)) == true)
                    {
                        Output = Conversions.ToDouble(Operators.AddObject(Output, Arr((object)i)));
                    }
                }
                OperationRet = Output;
            }

            else if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(Flag, "=COUNT()", false)))
            {
                int Output = 0;
                for (int i = Information.LBound((Array)Arr), loopTo1 = Information.UBound((Array)Arr); i <= loopTo1; i++)
                    Output = Output + 1;
                OperationRet = Output;
            }

            else if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(Flag, "=COUNTA()", false)))
            {
                int Output = 0;
                for (int i = Information.LBound((Array)Arr), loopTo2 = Information.UBound((Array)Arr); i <= loopTo2; i++)
                {
                    if (Arr((object)i) is not null)
                    {
                        Output = Output + 1;
                    }
                }
                OperationRet = Output;
            }

            else if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(Flag, "=AVERAGE()", false)))
            {
                double Output = 0d;
                for (int i = Information.LBound((Array)Arr), loopTo3 = Information.UBound((Array)Arr); i <= loopTo3; i++)
                {
                    if (Information.IsNumeric(Arr((object)i)) == true)
                    {
                        Output = Conversions.ToDouble(Operators.AddObject(Output, Arr((object)i)));
                    }
                }
                Output = Output / (Information.UBound((Array)Arr) + 1);
                OperationRet = Output;
            }

            else if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(Flag, "=MAX()", false)))
            {
                object Output;
                int i = Information.LBound((Array)Arr);
                while (Information.IsNumeric(Arr((object)i)) == false & i <= Information.UBound((Array)Arr) - 1)
                    i = i + 1;
                Output = Arr((object)i);
                var loopTo4 = Information.UBound((Array)Arr);
                for (i = Information.LBound((Array)Arr); i <= loopTo4; i++)
                {
                    if (Information.IsNumeric(Arr((object)i)) == true)
                    {
                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectGreater(Arr((object)i), Output, false)))
                        {
                            Output = Arr((object)i);
                        }
                    }
                }
                OperationRet = Output;
            }

            else if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(Flag, "=MIN()", false)))
            {
                object Output;
                int i = Information.LBound((Array)Arr);
                while (Information.IsNumeric(Arr((object)i)) == false & i <= Information.UBound((Array)Arr) - 1)
                    i = i + 1;

                Output = Arr((object)i);
                var loopTo5 = Information.UBound((Array)Arr);
                for (i = Information.LBound((Array)Arr); i <= loopTo5; i++)
                {
                    if (Information.IsNumeric(Arr((object)i)) == true)
                    {
                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectLess(Arr((object)i), Output, false)))
                        {
                            Output = Arr((object)i);
                        }
                    }
                }
                OperationRet = Output;
            }

            else if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(Flag, "=PRODUCT()", false)))
            {
                double Output = 1d;
                int count = 0;
                for (int i = Information.LBound((Array)Arr), loopTo6 = Information.UBound((Array)Arr); i <= loopTo6; i++)
                {
                    if (Information.IsNumeric(Arr((object)i)) == true)
                    {
                        Output = Conversions.ToDouble(Operators.MultiplyObject(Output, Arr((object)i)));
                        count = count + 1;
                    }
                }
                if (count == 0)
                {
                    OperationRet = 0;
                }
                else
                {
                    OperationRet = Output;
                }
            }
            else
            {
                OperationRet = 0;
            }

            return OperationRet;

        }

        private void Display()
        {

            try
            {

                CustomPanel1.Controls.Clear();
                CustomPanel2.Controls.Clear();

                Range displayRng;

                // Takes the first 50 rows of the input to display.
                if (rng.Rows.Count > 50)
                {
                    displayRng = (Range)rng.Rows["1:50"];
                }
                else
                {
                    displayRng = rng;
                }


                double height;
                double width;

                // Default number of rows in the display box is 4.
                if (displayRng.Rows.Count <= 4)
                {
                    height = CustomPanel1.Height / (double)displayRng.Rows.Count;
                }
                else
                {
                    height = 119d / 4d;
                }

                // Default number of columns in the display box is 4.
                if (displayRng.Columns.Count <= 3)
                {
                    width = CustomPanel1.Width / (double)displayRng.Columns.Count;
                }
                else
                {
                    width = 260d / 3d;
                }

                // Copies the input range to the display box.
                for (int i = 1, loopTo = displayRng.Rows.Count; i <= loopTo; i++)
                {
                    for (int j = 1, loopTo1 = displayRng.Columns.Count; j <= loopTo1; j++)
                    {

                        var label = new System.Windows.Forms.Label();
                        label.Text = Conversions.ToString(displayRng.Cells[i, j].Value);
                        label.Location = new System.Drawing.Point((int)Math.Round((j - 1) * width), (int)Math.Round((i - 1) * height));
                        label.Height = (int)Math.Round(height);
                        label.Width = (int)Math.Round(width);
                        label.BorderStyle = BorderStyle.FixedSingle;
                        label.TextAlign = ContentAlignment.MiddleCenter;

                        // Copies the format of the input range.
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

                CustomPanel1.AutoScroll = true;

                bool X4;

                if (RadioButton4.Checked)
                {
                    X4 = true;
                }
                else
                {
                    X4 = ComboBox3.SelectedIndex != -1;
                }

                if ((RadioButton1.Checked | RadioButton2.Checked | RadioButton3.Checked) & (RadioButton4.Checked | RadioButton5.Checked | RadioButton6.Checked) & X4)
                {

                    // Works for Merging Into Single Column.
                    if (RadioButton1.Checked)
                    {

                        var newWidth = default(double);
                        var newHeight = default(double);
                        var combinedColumn = default(int);

                        if (RadioButton6.Checked | RadioButton5.Checked)
                        {
                            if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(ComboBox3.SelectedItem, "Into Left Column", false)))
                            {
                                combinedColumn = 1;
                            }
                            else if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(ComboBox3.SelectedItem, "Into Right Column", false)))
                            {
                                combinedColumn = displayRng.Columns.Count;
                            }

                            for (int i = 1, loopTo2 = displayRng.Rows.Count; i <= loopTo2; i++)
                            {

                                var label = new System.Windows.Forms.Label();
                                if (ComboBox2.SelectedIndex <= 3)
                                {
                                    // Finds the combined string.
                                    string combinedValue = "";
                                    string Separator;
                                    int HFactor;
                                    int WFactor;

                                    if (ComboBox2.SelectedIndex == 3)
                                    {
                                        Separator = Microsoft.VisualBasic.Constants.vbNewLine;
                                        HFactor = (int)Math.Round(displayRng.Columns.Count / 1.75d);
                                        WFactor = 1;
                                    }
                                    else if (CheckBox4.Checked)
                                    {
                                        Separator = ComboBox2.Text + Microsoft.VisualBasic.Constants.vbNewLine;
                                        HFactor = (int)Math.Round(displayRng.Columns.Count / 1.75d);
                                        WFactor = 1;
                                    }
                                    else
                                    {
                                        Separator = ComboBox2.Text;
                                        HFactor = 1;
                                        WFactor = displayRng.Columns.Count;
                                    }

                                    for (int j = 1, loopTo3 = displayRng.Columns.Count - 1; j <= loopTo3; j++)
                                    {
                                        if (CheckBox3.Checked)
                                        {
                                            if (displayRng.Cells[i, j].value is not null)
                                            {
                                                combinedValue = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject(combinedValue, displayRng.Cells[i, j].Value), Separator));
                                            }
                                        }
                                        else
                                        {
                                            combinedValue = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject(combinedValue, displayRng.Cells[i, j].Value), Separator));
                                        }
                                    }

                                    if (CheckBox3.Checked)
                                    {
                                        if (displayRng.Cells[i, displayRng.Columns.Count].value is not null)
                                        {
                                            combinedValue = Conversions.ToString(Operators.ConcatenateObject(combinedValue, displayRng.Cells[i, displayRng.Columns.Count].Value));
                                        }
                                        else if (Strings.Len(combinedValue) >= Strings.Len(Separator))
                                        {
                                            combinedValue = Strings.Mid(combinedValue, 1, Strings.Len(combinedValue) - Strings.Len(Separator));
                                        }
                                    }
                                    else
                                    {
                                        combinedValue = Conversions.ToString(Operators.ConcatenateObject(combinedValue, displayRng.Cells[i, displayRng.Columns.Count].Value));
                                    }
                                    newWidth = width * WFactor;
                                    newHeight = height * HFactor;

                                    label.Text = combinedValue;
                                }
                                else
                                {
                                    // Finds the mathematical operated value (sum, max, min, count, etc...)
                                    double OperatedValue;
                                    var Values = new double[1];
                                    int Index = -1;
                                    for (int j = 1, loopTo4 = displayRng.Columns.Count; j <= loopTo4; j++)
                                    {
                                        if (Information.IsNumeric(displayRng.Cells[i, j].Value))
                                        {
                                            if (CheckBox3.Checked)
                                            {
                                                if (displayRng.Cells[i, j].value is not null)
                                                {
                                                    Index = Index + 1;
                                                    Array.Resize(ref Values, Index + 1);
                                                    Values[Index] = Conversions.ToDouble(displayRng.Cells[i, j].value);
                                                }
                                            }
                                            else
                                            {
                                                Index = Index + 1;
                                                Array.Resize(ref Values, Index + 1);
                                                Values[Index] = Conversions.ToDouble(displayRng.Cells[i, j].value);
                                            }
                                        }
                                    }
                                    OperatedValue = Conversions.ToDouble(Operation(Values, ComboBox2.SelectedItem));
                                    label.Text = OperatedValue.ToString();
                                    newWidth = width;
                                    newHeight = height;
                                }

                                // Puts the output value in the display box.
                                label.Location = new System.Drawing.Point((int)Math.Round((combinedColumn - 1) * width), (int)Math.Round((i - 1) * newHeight));
                                label.Height = (int)Math.Round(newHeight);
                                label.Width = (int)Math.Round(newWidth);
                                label.BorderStyle = BorderStyle.FixedSingle;
                                label.TextAlign = ContentAlignment.MiddleCenter;

                                // Copies the format of the output cell.
                                if (CheckBox1.Checked == true)
                                {
                                    Range cell = (Range)displayRng.Cells[i, 1];
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

                            // Copies the rest columns other than the merged column.
                            if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(ComboBox3.SelectedItem, "Into Left Column", false)))
                            {

                                for (int i = 1, loopTo5 = displayRng.Rows.Count; i <= loopTo5; i++)
                                {
                                    for (int j = 2, loopTo6 = displayRng.Columns.Count; j <= loopTo6; j++)
                                    {
                                        var label = new System.Windows.Forms.Label();
                                        if (RadioButton6.Checked)
                                        {
                                            label.Text = Conversions.ToString(displayRng.Cells[i, j].value);
                                        }
                                        else if (RadioButton5.Checked)
                                        {
                                            label.Text = "";
                                        }
                                        label.Location = new System.Drawing.Point((int)Math.Round(newWidth + (j - 2) * width), (int)Math.Round((i - 1) * newHeight));
                                        label.Height = (int)Math.Round(newHeight);
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
                            }

                            else if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(ComboBox3.SelectedItem, "Into Right Column", false)))
                            {
                                for (int i = 1, loopTo7 = displayRng.Rows.Count; i <= loopTo7; i++)
                                {
                                    for (int j = 1, loopTo8 = displayRng.Columns.Count - 1; j <= loopTo8; j++)
                                    {
                                        var label = new System.Windows.Forms.Label();
                                        if (RadioButton6.Checked)
                                        {
                                            label.Text = Conversions.ToString(displayRng.Cells[i, j].value);
                                        }
                                        else if (RadioButton5.Checked)
                                        {
                                            label.Text = "";
                                        }
                                        label.Location = new System.Drawing.Point((int)Math.Round((j - 1) * width), (int)Math.Round((i - 1) * newHeight));
                                        label.Height = (int)Math.Round(newHeight);
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
                            }
                        }

                        // Works to Merge the Whole Row.       
                        else if (RadioButton4.Checked)
                        {
                            var WFactor = default(int);
                            for (int i = 1, loopTo9 = displayRng.Rows.Count; i <= loopTo9; i++)
                            {
                                var label = new System.Windows.Forms.Label();
                                int HFactor;
                                if (ComboBox2.SelectedIndex <= 3)
                                {
                                    string combinedValue = "";
                                    string Separator;

                                    if (ComboBox2.SelectedIndex == 3)
                                    {
                                        Separator = Microsoft.VisualBasic.Constants.vbNewLine;
                                        HFactor = (int)Math.Round(displayRng.Columns.Count / 1.75d);
                                        WFactor = 1;
                                    }
                                    else if (CheckBox4.Checked)
                                    {
                                        Separator = ComboBox2.Text + Microsoft.VisualBasic.Constants.vbNewLine;
                                        HFactor = (int)Math.Round(displayRng.Columns.Count / 1.75d);
                                        WFactor = 1;
                                    }
                                    else
                                    {
                                        Separator = ComboBox2.Text;
                                        HFactor = 1;
                                        WFactor = displayRng.Columns.Count;
                                    }

                                    for (int j = 1, loopTo10 = displayRng.Columns.Count - 1; j <= loopTo10; j++)
                                    {
                                        if (CheckBox3.Checked)
                                        {
                                            if (displayRng.Cells[i, j].value is not null)
                                            {
                                                combinedValue = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject(combinedValue, displayRng.Cells[i, j].Value), Separator));
                                            }
                                        }
                                        else
                                        {
                                            combinedValue = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject(combinedValue, displayRng.Cells[i, j].Value), Separator));
                                        }
                                    }

                                    if (CheckBox3.Checked)
                                    {
                                        if (displayRng.Cells[i, displayRng.Columns.Count].value is not null)
                                        {
                                            combinedValue = Conversions.ToString(Operators.ConcatenateObject(combinedValue, displayRng.Cells[i, displayRng.Columns.Count].Value));
                                        }
                                        else if (Strings.Len(combinedValue) >= Strings.Len(Separator))
                                        {
                                            combinedValue = Strings.Mid(combinedValue, 1, Strings.Len(combinedValue) - Strings.Len(Separator));
                                        }
                                    }
                                    else
                                    {
                                        combinedValue = Conversions.ToString(Operators.ConcatenateObject(combinedValue, displayRng.Cells[i, displayRng.Columns.Count].Value));
                                    }
                                    newWidth = width * WFactor;
                                    newHeight = height * HFactor;

                                    label.Text = combinedValue;
                                }
                                else
                                {
                                    double OperatedValue;
                                    var Values = new double[1];
                                    int Index = -1;
                                    for (int j = 1, loopTo11 = displayRng.Columns.Count; j <= loopTo11; j++)
                                    {
                                        if (Information.IsNumeric(displayRng.Cells[i, j].Value))
                                        {
                                            if (CheckBox3.Checked)
                                            {
                                                if (displayRng.Cells[i, j].value is not null)
                                                {
                                                    Index = Index + 1;
                                                    Array.Resize(ref Values, Index + 1);
                                                    Values[Index] = Conversions.ToDouble(displayRng.Cells[i, j].value);
                                                }
                                            }
                                            else
                                            {
                                                Index = Index + 1;
                                                Array.Resize(ref Values, Index + 1);
                                                Values[Index] = Conversions.ToDouble(displayRng.Cells[i, j].value);
                                            }
                                        }
                                    }
                                    OperatedValue = Conversions.ToDouble(Operation(Values, ComboBox2.SelectedItem));
                                    newHeight = height;
                                    newWidth = width;

                                    label.Text = OperatedValue.ToString();
                                }
                                label.Location = new System.Drawing.Point(0, (int)Math.Round((i - 1) * newHeight));
                                label.Height = (int)Math.Round(newHeight);
                                if (WFactor != 1)
                                {
                                    label.Width = (int)Math.Round(newWidth);
                                }
                                else
                                {
                                    label.Width = CustomPanel2.Width;
                                }
                                label.BorderStyle = BorderStyle.FixedSingle;
                                label.TextAlign = ContentAlignment.MiddleCenter;

                                if (CheckBox1.Checked == true)
                                {
                                    Range cell = (Range)displayRng.Cells[i, 1];
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

                    else if (RadioButton2.Checked)
                    {

                        var newHeight = default(double);
                        var newWidth = default(double);
                        var combinedRow = default(int);

                        if (RadioButton6.Checked | RadioButton5.Checked)
                        {
                            if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(ComboBox3.SelectedItem, "Into Top Row", false)))
                            {
                                combinedRow = 1;
                            }
                            else if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(ComboBox3.SelectedItem, "Into Bottom Row", false)))
                            {
                                combinedRow = displayRng.Rows.Count;
                            }
                            for (int j = 1, loopTo12 = displayRng.Columns.Count; j <= loopTo12; j++)
                            {
                                var label = new System.Windows.Forms.Label();
                                if (ComboBox2.SelectedIndex <= 3)
                                {
                                    string combinedValue = "";
                                    string Separator;
                                    int HFactor;
                                    int WFactor;

                                    if (ComboBox2.SelectedIndex == 3)
                                    {
                                        Separator = Microsoft.VisualBasic.Constants.vbNewLine;
                                        HFactor = (int)Math.Round(displayRng.Rows.Count / 1.75d);
                                        WFactor = 1;
                                    }
                                    else if (CheckBox4.Checked)
                                    {
                                        Separator = ComboBox2.Text + Microsoft.VisualBasic.Constants.vbNewLine;
                                        HFactor = (int)Math.Round(displayRng.Rows.Count / 1.75d);
                                        WFactor = 1;
                                    }
                                    else
                                    {
                                        Separator = ComboBox2.Text;
                                        HFactor = 1;
                                        WFactor = displayRng.Rows.Count;
                                    }

                                    for (int i = 1, loopTo13 = displayRng.Rows.Count - 1; i <= loopTo13; i++)
                                    {
                                        if (CheckBox3.Checked)
                                        {
                                            if (displayRng.Cells[i, j].value is not null)
                                            {
                                                combinedValue = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject(combinedValue, displayRng.Cells[i, j].Value), Separator));
                                            }
                                        }
                                        else
                                        {
                                            combinedValue = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject(combinedValue, displayRng.Cells[i, j].Value), Separator));
                                        }
                                    }

                                    if (CheckBox3.Checked)
                                    {
                                        if (displayRng.Cells[displayRng.Rows.Count, j].value is not null)
                                        {
                                            combinedValue = Conversions.ToString(Operators.ConcatenateObject(combinedValue, displayRng.Cells[displayRng.Rows.Count, j].Value));
                                        }
                                        else if (Strings.Len(combinedValue) >= Strings.Len(Separator))
                                        {
                                            combinedValue = Strings.Mid(combinedValue, 1, Strings.Len(combinedValue) - Strings.Len(Separator));
                                        }
                                    }
                                    else
                                    {
                                        combinedValue = Conversions.ToString(Operators.ConcatenateObject(combinedValue, displayRng.Cells[displayRng.Rows.Count, j].Value));
                                    }
                                    newWidth = width * WFactor;
                                    newHeight = height * HFactor;

                                    label.Text = combinedValue;
                                }

                                else
                                {
                                    double OperatedValue;
                                    var Values = new double[1];
                                    int Index = -1;
                                    for (int i = 1, loopTo14 = displayRng.Rows.Count; i <= loopTo14; i++)
                                    {
                                        if (Information.IsNumeric(displayRng.Cells[i, j].Value))
                                        {
                                            if (CheckBox3.Checked)
                                            {
                                                if (displayRng.Cells[i, j].value is not null)
                                                {
                                                    Index = Index + 1;
                                                    Array.Resize(ref Values, Index + 1);
                                                    Values[Index] = Conversions.ToDouble(displayRng.Cells[i, j].value);
                                                }
                                            }
                                            else
                                            {
                                                Index = Index + 1;
                                                Array.Resize(ref Values, Index + 1);
                                                Values[Index] = Conversions.ToDouble(displayRng.Cells[i, j].value);
                                            }
                                        }
                                    }
                                    OperatedValue = Conversions.ToDouble(Operation(Values, ComboBox2.SelectedItem));
                                    label.Text = OperatedValue.ToString();
                                    newHeight = height;
                                    newWidth = width;
                                }
                                label.Location = new System.Drawing.Point((int)Math.Round((j - 1) * newWidth), (int)Math.Round((combinedRow - 1) * height));
                                label.Height = (int)Math.Round(newHeight);
                                label.Width = (int)Math.Round(newWidth);
                                label.BorderStyle = BorderStyle.FixedSingle;
                                label.TextAlign = ContentAlignment.MiddleCenter;

                                if (CheckBox1.Checked == true)
                                {
                                    Range cell = (Range)displayRng.Cells[1, j];
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

                            if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(ComboBox3.SelectedItem, "Into Top Row", false)))
                            {

                                for (int j = 1, loopTo15 = displayRng.Columns.Count; j <= loopTo15; j++)
                                {
                                    for (int i = 2, loopTo16 = displayRng.Rows.Count; i <= loopTo16; i++)
                                    {
                                        var label = new System.Windows.Forms.Label();
                                        if (RadioButton6.Checked)
                                        {
                                            label.Text = Conversions.ToString(displayRng.Cells[i, j].value);
                                        }
                                        else if (RadioButton5.Checked)
                                        {
                                            label.Text = "";
                                        }
                                        label.Location = new System.Drawing.Point((int)Math.Round((j - 1) * newWidth), (int)Math.Round(newHeight + (i - 2) * height));
                                        label.Height = (int)Math.Round(height);
                                        label.Width = (int)Math.Round(newWidth);
                                        label.BorderStyle = BorderStyle.FixedSingle;
                                        label.TextAlign = ContentAlignment.MiddleCenter;

                                        if (CheckBox1.Checked == true)
                                        {
                                            Range cell = (Range)displayRng.Cells[i, 1];
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
                            }

                            else if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(ComboBox3.SelectedItem, "Into Bottom Row", false)))
                            {
                                for (int j = 1, loopTo17 = displayRng.Columns.Count; j <= loopTo17; j++)
                                {
                                    for (int i = 1, loopTo18 = displayRng.Rows.Count - 1; i <= loopTo18; i++)
                                    {
                                        var label = new System.Windows.Forms.Label();
                                        if (RadioButton6.Checked)
                                        {
                                            label.Text = Conversions.ToString(displayRng.Cells[i, j].value);
                                        }
                                        else if (RadioButton5.Checked)
                                        {
                                            label.Text = "";
                                        }
                                        label.Location = new System.Drawing.Point((int)Math.Round((j - 1) * newWidth), (int)Math.Round((i - 1) * height));
                                        label.Height = (int)Math.Round(height);
                                        label.Width = (int)Math.Round(newWidth);
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
                            }
                        }

                        else if (RadioButton4.Checked)
                        {
                            var HFactor = default(int);
                            for (int j = 1, loopTo19 = displayRng.Columns.Count; j <= loopTo19; j++)
                            {
                                var label = new System.Windows.Forms.Label();
                                int WFactor;

                                if (ComboBox2.SelectedIndex <= 3)
                                {
                                    string combinedValue = "";
                                    string Separator;

                                    if (ComboBox2.SelectedIndex == 3)
                                    {
                                        Separator = Microsoft.VisualBasic.Constants.vbNewLine;
                                        HFactor = (int)Math.Round(displayRng.Rows.Count / 1.75d);
                                        WFactor = 1;
                                    }
                                    else if (CheckBox4.Checked)
                                    {
                                        Separator = ComboBox2.Text + Microsoft.VisualBasic.Constants.vbNewLine;
                                        HFactor = (int)Math.Round(displayRng.Rows.Count / 1.75d);
                                        WFactor = 1;
                                    }
                                    else
                                    {
                                        Separator = ComboBox2.Text;
                                        HFactor = 1;
                                        WFactor = displayRng.Rows.Count;
                                    }

                                    for (int i = 1, loopTo20 = displayRng.Rows.Count - 1; i <= loopTo20; i++)
                                    {
                                        if (CheckBox3.Checked)
                                        {
                                            if (displayRng.Cells[i, j].value is not null)
                                            {
                                                combinedValue = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject(combinedValue, displayRng.Cells[i, j].Value), Separator));
                                            }
                                        }
                                        else
                                        {
                                            combinedValue = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject(combinedValue, displayRng.Cells[i, j].Value), Separator));
                                        }
                                    }

                                    if (CheckBox3.Checked)
                                    {
                                        if (displayRng.Cells[displayRng.Rows.Count, j].value is not null)
                                        {
                                            combinedValue = Conversions.ToString(Operators.ConcatenateObject(combinedValue, displayRng.Cells[displayRng.Rows.Count, j].Value));
                                        }
                                        else if (Strings.Len(combinedValue) >= Strings.Len(Separator))
                                        {
                                            combinedValue = Strings.Mid(combinedValue, 1, Strings.Len(combinedValue) - Strings.Len(Separator));
                                        }
                                    }
                                    else
                                    {
                                        combinedValue = Conversions.ToString(Operators.ConcatenateObject(combinedValue, displayRng.Cells[displayRng.Rows.Count, j].Value));
                                    }
                                    newWidth = width * WFactor;
                                    newHeight = height * HFactor;
                                    label.Text = combinedValue;
                                }
                                else
                                {
                                    double OperatedValue;
                                    var Values = new double[1];
                                    int Index = -1;
                                    for (int i = 1, loopTo21 = displayRng.Rows.Count; i <= loopTo21; i++)
                                    {
                                        if (Information.IsNumeric(displayRng.Cells[i, j].Value))
                                        {
                                            if (CheckBox3.Checked)
                                            {
                                                if (displayRng.Cells[i, j].value is not null)
                                                {
                                                    Index = Index + 1;
                                                    Array.Resize(ref Values, Index + 1);
                                                    Values[Index] = Conversions.ToDouble(displayRng.Cells[i, j].value);
                                                }
                                            }
                                            else
                                            {
                                                Index = Index + 1;
                                                Array.Resize(ref Values, Index + 1);
                                                Values[Index] = Conversions.ToDouble(displayRng.Cells[i, j].value);
                                            }
                                        }
                                    }
                                    OperatedValue = Conversions.ToDouble(Operation(Values, ComboBox2.SelectedItem));
                                    newWidth = newWidth;
                                    newHeight = newHeight;

                                    label.Text = OperatedValue.ToString();
                                }
                                label.Location = new System.Drawing.Point((int)Math.Round((j - 1) * newWidth), 0);
                                if (HFactor != 1)
                                {
                                    label.Height = (int)Math.Round(newHeight);
                                }
                                else
                                {
                                    label.Height = CustomPanel2.Height;
                                }
                                label.Width = (int)Math.Round(newWidth);
                                label.BorderStyle = BorderStyle.FixedSingle;
                                label.TextAlign = ContentAlignment.MiddleCenter;

                                if (CheckBox1.Checked == true)
                                {
                                    Range cell = (Range)displayRng.Cells[1, j];
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

                    else if (RadioButton3.Checked)
                    {

                        var combinedRow = default(int);
                        var combinedColumn = default(int);
                        string combinedValue = "";
                        double OperatedValue;
                        var Values = new double[1];
                        int Index = -1;

                        string Separator;
                        string RowColumn;

                        if (ComboBox2.SelectedIndex == 3)
                        {
                            Separator = Microsoft.VisualBasic.Constants.vbNewLine;
                            RowColumn = "Row";
                        }
                        else if (CheckBox4.Checked)
                        {
                            Separator = ComboBox2.Text + Microsoft.VisualBasic.Constants.vbNewLine;
                            RowColumn = "Row";
                        }
                        else
                        {
                            Separator = ComboBox2.Text;
                            RowColumn = "Column";
                        }

                        for (int i = 1, loopTo22 = displayRng.Rows.Count - 1; i <= loopTo22; i++)
                        {
                            for (int j = 1, loopTo23 = displayRng.Columns.Count; j <= loopTo23; j++)
                            {
                                if (ComboBox2.SelectedIndex <= 3)
                                {
                                    if (CheckBox3.Checked)
                                    {
                                        if (displayRng.Cells[i, j].value is not null)
                                        {
                                            combinedValue = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject(combinedValue, displayRng.Cells[i, j].Value), Separator));
                                        }
                                    }
                                    else
                                    {
                                        combinedValue = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject(combinedValue, displayRng.Cells[i, j].Value), Separator));
                                    }
                                }
                                else if (Information.IsNumeric(displayRng.Cells[i, j].Value))
                                {
                                    if (CheckBox3.Checked)
                                    {
                                        if (displayRng.Cells[i, j].value is not null)
                                        {
                                            Index = Index + 1;
                                            Array.Resize(ref Values, Index + 1);
                                            Values[Index] = Conversions.ToDouble(displayRng.Cells[i, j].value);
                                        }
                                    }
                                    else
                                    {
                                        Index = Index + 1;
                                        Array.Resize(ref Values, Index + 1);
                                        Values[Index] = Conversions.ToDouble(displayRng.Cells[i, j].value);
                                    }
                                }
                            }
                        }

                        for (int j = 1, loopTo24 = displayRng.Columns.Count - 1; j <= loopTo24; j++)
                        {
                            if (ComboBox2.SelectedIndex <= 3)
                            {
                                if (CheckBox3.Checked)
                                {
                                    if (displayRng.Cells[displayRng.Rows.Count, j].value is not null)
                                    {
                                        combinedValue = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject(combinedValue, displayRng.Cells[displayRng.Rows.Count, j].Value), Separator));
                                    }
                                }

                                else
                                {
                                    combinedValue = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject(combinedValue, displayRng.Cells[rng.Rows.Count, j].Value), Separator));
                                }
                            }
                            else if (Information.IsNumeric(displayRng.Cells[displayRng.Rows.Count, j].Value))
                            {
                                if (CheckBox3.Checked)
                                {
                                    if (displayRng.Cells[displayRng.Rows.Count, j].value is not null)
                                    {
                                        Index = Index + 1;
                                        Array.Resize(ref Values, Index + 1);
                                        Values[Index] = Conversions.ToDouble(displayRng.Cells[displayRng.Rows.Count, j].value);
                                    }
                                }
                                else
                                {
                                    Index = Index + 1;
                                    Array.Resize(ref Values, Index + 1);
                                    Values[Index] = Conversions.ToDouble(displayRng.Cells[displayRng.Rows.Count, j].value);
                                }
                            }
                        }

                        if (ComboBox2.SelectedIndex <= 3)
                        {
                            if (CheckBox3.Checked)
                            {
                                if (displayRng.Cells[displayRng.Rows.Count, displayRng.Columns.Count].value is not null)
                                {
                                    combinedValue = Conversions.ToString(Operators.ConcatenateObject(combinedValue, displayRng.Cells[displayRng.Rows.Count, displayRng.Columns.Count].Value));
                                }
                                else if (Strings.Len(combinedValue) >= Strings.Len(Separator))
                                {
                                    combinedValue = Strings.Mid(combinedValue, 1, Strings.Len(combinedValue) - Strings.Len(Separator));
                                }
                            }

                            else
                            {
                                combinedValue = Conversions.ToString(Operators.ConcatenateObject(combinedValue, rng.Cells[rng.Rows.Count, rng.Columns.Count].Value));
                            }
                        }
                        else if (Information.IsNumeric(displayRng.Cells[displayRng.Rows.Count, displayRng.Columns.Count].Value))
                        {
                            if (CheckBox3.Checked)
                            {
                                if (displayRng.Cells[displayRng.Rows.Count, displayRng.Columns.Count].value is not null)
                                {
                                    Index = Index + 1;
                                    Array.Resize(ref Values, Index + 1);
                                    Values[Index] = Conversions.ToDouble(displayRng.Cells[displayRng.Rows.Count, displayRng.Columns.Count].value);
                                }
                            }
                            else
                            {
                                Index = Index + 1;
                                Array.Resize(ref Values, Index + 1);
                                Values[Index] = Conversions.ToDouble(displayRng.Cells[displayRng.Rows.Count, displayRng.Columns.Count].value);
                            }
                        }

                        OperatedValue = Conversions.ToDouble(Operation(Values, ComboBox2.SelectedItem));

                        if (RadioButton6.Checked | RadioButton5.Checked)
                        {
                            if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(ComboBox3.SelectedItem, "Into Top-Left Cell", false)))
                            {
                                combinedRow = 1;
                                combinedColumn = 1;
                            }
                            else if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(ComboBox3.SelectedItem, "Into Top-Right Cell", false)))
                            {
                                combinedRow = 1;
                                combinedColumn = displayRng.Columns.Count;
                            }
                            else if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(ComboBox3.SelectedItem, "Into Bottom-Left Cell", false)))
                            {
                                combinedRow = displayRng.Rows.Count;
                                combinedColumn = 1;
                            }
                            else if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(ComboBox3.SelectedItem, "Into Bottom-Right Cell", false)))
                            {
                                combinedRow = displayRng.Rows.Count;
                                combinedColumn = displayRng.Columns.Count;
                            }

                            for (int i = 1, loopTo25 = displayRng.Rows.Count; i <= loopTo25; i++)
                            {
                                for (int j = 1, loopTo26 = displayRng.Columns.Count; j <= loopTo26; j++)
                                {
                                    if (i == combinedRow & j == combinedColumn)
                                    {
                                        var label = new System.Windows.Forms.Label();
                                        if (ComboBox2.SelectedIndex <= 3)
                                        {
                                            label.Text = combinedValue;
                                        }
                                        else
                                        {
                                            label.Text = OperatedValue.ToString();
                                        }

                                        if (RowColumn == "Row")
                                        {
                                            if (i > combinedRow)
                                            {
                                                label.Location = new System.Drawing.Point((int)Math.Round((j - 1) * width), (int)Math.Round(height * (displayRng.Cells.Count / 1.75d) + (i - 2) * height));
                                            }
                                            else
                                            {
                                                label.Location = new System.Drawing.Point((int)Math.Round((j - 1) * width), (int)Math.Round((i - 1) * height));
                                            }
                                        }
                                        else if (j > combinedColumn)
                                        {
                                            label.Location = new System.Drawing.Point((int)Math.Round(width * displayRng.Cells.Count + (j - 2) * width), (int)Math.Round((i - 1) * height));
                                        }
                                        else
                                        {
                                            label.Location = new System.Drawing.Point((int)Math.Round((j - 1) * width), (int)Math.Round((i - 1) * height));
                                        }

                                        if (RowColumn == "Row")
                                        {
                                            label.Height = (int)Math.Round(height * displayRng.Cells.Count / 1.75d);
                                            label.Width = (int)Math.Round(width);
                                        }
                                        else
                                        {
                                            label.Height = (int)Math.Round(height);
                                            label.Width = (int)Math.Round(width * displayRng.Cells.Count);
                                        }

                                        label.BorderStyle = BorderStyle.FixedSingle;
                                        label.TextAlign = ContentAlignment.MiddleCenter;

                                        if (CheckBox1.Checked == true)
                                        {
                                            Range cell = (Range)displayRng.Cells[combinedRow, combinedColumn];
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
                                    else
                                    {

                                        var label = new System.Windows.Forms.Label();

                                        if (RadioButton6.Checked)
                                        {
                                            label.Text = Conversions.ToString(displayRng.Cells[i, j].Value);
                                        }
                                        else
                                        {
                                            label.Text = "";
                                        }
                                        if (RowColumn == "Row")
                                        {
                                            if (i > combinedRow)
                                            {
                                                label.Location = new System.Drawing.Point((int)Math.Round((j - 1) * width), (int)Math.Round(height * (displayRng.Cells.Count / 1.75d) + (i - 2) * height));
                                            }
                                            else
                                            {
                                                label.Location = new System.Drawing.Point((int)Math.Round((j - 1) * width), (int)Math.Round((i - 1) * height));
                                            }
                                        }
                                        else if (j > combinedColumn)
                                        {
                                            label.Location = new System.Drawing.Point((int)Math.Round(width * displayRng.Cells.Count + (j - 2) * width), (int)Math.Round((i - 1) * height));
                                        }
                                        else
                                        {
                                            label.Location = new System.Drawing.Point((int)Math.Round((j - 1) * width), (int)Math.Round((i - 1) * height));
                                        }
                                        if (RowColumn == "Row")
                                        {
                                            if (i == combinedRow)
                                            {
                                                label.Height = (int)Math.Round(height * displayRng.Cells.Count / 1.75d);
                                                label.Width = (int)Math.Round(width);
                                            }
                                            else
                                            {
                                                label.Height = (int)Math.Round(height);
                                                label.Width = (int)Math.Round(width);
                                            }
                                        }
                                        else if (j == combinedColumn)
                                        {
                                            label.Height = (int)Math.Round(height);
                                            label.Width = (int)Math.Round(width * displayRng.Cells.Count);
                                        }
                                        else
                                        {
                                            label.Height = (int)Math.Round(height);
                                            label.Width = (int)Math.Round(width);
                                        }
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
                            }
                            CustomPanel2.AutoScroll = true;
                        }

                        else if (RadioButton4.Checked)
                        {

                            var label = new System.Windows.Forms.Label();
                            if (ComboBox2.SelectedIndex <= 3)
                            {
                                label.Text = combinedValue;
                            }
                            else
                            {
                                label.Text = OperatedValue.ToString();
                            }

                            label.Location = new System.Drawing.Point(0, 0);
                            if (RowColumn == "Row")
                            {
                                label.Height = (int)Math.Round(height * displayRng.Cells.Count);
                                label.Width = CustomPanel2.Width;
                            }
                            else
                            {
                                label.Height = CustomPanel2.Height;
                                label.Width = (int)Math.Round(width * displayRng.Cells.Count);
                            }

                            label.BorderStyle = BorderStyle.FixedSingle;
                            label.TextAlign = ContentAlignment.MiddleCenter;

                            if (CheckBox1.Checked == true)
                            {
                                Range cell = (Range)displayRng.Cells[1, 1];
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

                            CustomPanel2.AutoScroll = true;
                        }
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

                bool X4;
                if (RadioButton4.Checked)
                {
                    X4 = true;
                }
                else
                {
                    X4 = ComboBox3.SelectedIndex != -1;
                }

                if (string.IsNullOrEmpty(TextBox1.Text))
                {
                    MessageBox.Show("Select a Source Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    TextBox1.Focus();
                    return;
                }

                if (IsValidExcelCellReference(TextBox1.Text) == false)
                {
                    MessageBox.Show("Select a Valid Source Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    TextBox1.Focus();
                    return;
                }

                if (RadioButton1.Checked == false & RadioButton2.Checked == false & RadioButton3.Checked == false)
                {
                    MessageBox.Show("Select Where to Combine the Selected Data.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    workSheet.Activate();
                    rng.Select();
                    return;
                }
                else if (X4 == false)
                {
                    MessageBox.Show("Select Where to Store the Selected Data.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    workSheet.Activate();
                    rng.Select();
                    return;
                }
                else if (RadioButton6.Checked == false & RadioButton5.Checked == false & RadioButton4.Checked == false)
                {
                    MessageBox.Show("Select One Combination Option.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    ComboBox3.Focus();
                    workSheet.Activate();
                    rng.Select();
                    return;
                }

                if (CheckBox2.Checked == true)
                {
                    workSheet.Copy(After: workBook.Sheets[workSheet.Name]);
                    workSheet.Activate();
                }

                if ((RadioButton1.Checked | RadioButton2.Checked | RadioButton3.Checked) & (RadioButton4.Checked | RadioButton5.Checked | RadioButton6.Checked) & X4)
                {

                    if (RadioButton1.Checked)
                    {

                        if (RadioButton6.Checked | RadioButton5.Checked)
                        {
                            var combinedColumn = default(int);
                            if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(ComboBox3.SelectedItem, "Into Left Column", false)))
                            {
                                combinedColumn = 1;
                            }
                            else if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(ComboBox3.SelectedItem, "Into Right Column", false)))
                            {
                                combinedColumn = rng.Columns.Count;
                            }
                            for (int i = 1, loopTo = rng.Rows.Count; i <= loopTo; i++)
                            {

                                if (ComboBox2.SelectedIndex <= 3)
                                {
                                    string combinedValue = "";
                                    string Separator;

                                    if (ComboBox2.SelectedIndex == 3)
                                    {
                                        Separator = Microsoft.VisualBasic.Constants.vbNewLine;
                                    }
                                    else if (CheckBox4.Checked)
                                    {
                                        Separator = ComboBox2.Text + Microsoft.VisualBasic.Constants.vbNewLine;
                                    }
                                    else
                                    {
                                        Separator = ComboBox2.Text;
                                    }

                                    for (int j = 1, loopTo1 = rng.Columns.Count - 1; j <= loopTo1; j++)
                                    {
                                        if (CheckBox3.Checked)
                                        {
                                            if (rng.Cells[i, j].value is not null)
                                            {
                                                combinedValue = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject(combinedValue, rng.Cells[i, j].Value), Separator));
                                            }
                                        }
                                        else
                                        {
                                            combinedValue = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject(combinedValue, rng.Cells[i, j].Value), Separator));
                                        }
                                    }

                                    if (CheckBox3.Checked)
                                    {
                                        if (rng.Cells[i, rng.Columns.Count].value is not null)
                                        {
                                            combinedValue = Conversions.ToString(Operators.ConcatenateObject(combinedValue, rng.Cells[i, rng.Columns.Count].Value));
                                        }
                                        else if (Strings.Len(combinedValue) >= Strings.Len(Separator))
                                        {
                                            combinedValue = Strings.Mid(combinedValue, 1, Strings.Len(combinedValue) - Strings.Len(Separator));
                                        }
                                    }

                                    else
                                    {
                                        combinedValue = Conversions.ToString(Operators.ConcatenateObject(combinedValue, rng.Cells[i, rng.Columns.Count].Value));
                                    }

                                    rng.Cells[i, combinedColumn].value = combinedValue;
                                }

                                else
                                {
                                    double OperatedValue;
                                    var Values = new double[1];
                                    int Index = -1;
                                    for (int j = 1, loopTo2 = rng.Columns.Count; j <= loopTo2; j++)
                                    {
                                        if (Information.IsNumeric(rng.Cells[i, j].Value))
                                        {
                                            if (CheckBox3.Checked)
                                            {
                                                if (rng.Cells[i, j].value is not null)
                                                {
                                                    Index = Index + 1;
                                                    Array.Resize(ref Values, Index + 1);
                                                    Values[Index] = Conversions.ToDouble(rng.Cells[i, j].value);
                                                }
                                            }
                                            else
                                            {
                                                Index = Index + 1;
                                                Array.Resize(ref Values, Index + 1);
                                                Values[Index] = Conversions.ToDouble(rng.Cells[i, j].value);
                                            }
                                        }
                                    }
                                    OperatedValue = Conversions.ToDouble(Operation(Values, ComboBox2.SelectedItem));
                                    rng.Cells[i, combinedColumn].value = (object)OperatedValue;
                                }

                            }

                            if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(ComboBox3.SelectedItem, "Into Left Column", false)))
                            {

                                for (int i = 1, loopTo3 = rng.Rows.Count; i <= loopTo3; i++)
                                {
                                    for (int j = 2, loopTo4 = rng.Columns.Count; j <= loopTo4; j++)
                                    {
                                        if (RadioButton5.Checked)
                                        {
                                            rng.Cells[i, j].Clear();
                                        }
                                    }
                                }
                            }
                            else if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(ComboBox3.SelectedItem, "Into Right Column", false)))
                            {
                                for (int i = 1, loopTo5 = rng.Rows.Count; i <= loopTo5; i++)
                                {
                                    for (int j = 1, loopTo6 = rng.Columns.Count - 1; j <= loopTo6; j++)
                                    {
                                        if (RadioButton5.Checked)
                                        {
                                            rng.Cells[i, j].Clear();
                                        }
                                    }
                                }
                                excelApp.DisplayAlerts = true;
                            }
                        }

                        else if (RadioButton4.Checked)
                        {
                            excelApp.DisplayAlerts = false;
                            for (int i = 1, loopTo7 = rng.Rows.Count; i <= loopTo7; i++)
                            {
                                if (ComboBox2.SelectedIndex <= 3)
                                {
                                    string combinedValue = "";
                                    string Separator;

                                    if (ComboBox2.SelectedIndex == 3)
                                    {
                                        Separator = Microsoft.VisualBasic.Constants.vbNewLine;
                                    }
                                    else if (CheckBox4.Checked)
                                    {
                                        Separator = ComboBox2.Text + Microsoft.VisualBasic.Constants.vbNewLine;
                                    }
                                    else
                                    {
                                        Separator = ComboBox2.Text;
                                    }
                                    for (int j = 1, loopTo8 = rng.Columns.Count - 1; j <= loopTo8; j++)
                                    {
                                        if (CheckBox3.Checked)
                                        {
                                            if (rng.Cells[i, j].value is not null)
                                            {
                                                combinedValue = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject(combinedValue, rng.Cells[i, j].Value), Separator));
                                            }
                                        }
                                        else
                                        {
                                            combinedValue = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject(combinedValue, rng.Cells[i, j].Value), Separator));
                                        }
                                    }
                                    if (CheckBox3.Checked)
                                    {
                                        if (rng.Cells[i, rng.Columns.Count].value is not null)
                                        {
                                            combinedValue = Conversions.ToString(Operators.ConcatenateObject(combinedValue, rng.Cells[i, rng.Columns.Count].Value));
                                        }
                                        else if (Strings.Len(combinedValue) >= Strings.Len(Separator))
                                        {
                                            combinedValue = Strings.Mid(combinedValue, 1, Strings.Len(combinedValue) - Strings.Len(Separator));
                                        }
                                    }
                                    else
                                    {
                                        combinedValue = Conversions.ToString(Operators.ConcatenateObject(combinedValue, rng.Cells[i, rng.Columns.Count].Value));
                                    }
                                    rng.Cells[i, 1].value = combinedValue;
                                }
                                else
                                {
                                    double OperatedValue;
                                    var Values = new double[1];
                                    int Index = -1;
                                    for (int j = 1, loopTo9 = rng.Columns.Count; j <= loopTo9; j++)
                                    {
                                        if (Information.IsNumeric(rng.Cells[i, j].Value))
                                        {
                                            if (CheckBox3.Checked)
                                            {
                                                if (rng.Cells[i, j].value is not null)
                                                {
                                                    Index = Index + 1;
                                                    Array.Resize(ref Values, Index + 1);
                                                    Values[Index] = Conversions.ToDouble(rng.Cells[i, j].value);
                                                }
                                            }
                                            else
                                            {
                                                Index = Index + 1;
                                                Array.Resize(ref Values, Index + 1);
                                                Values[Index] = Conversions.ToDouble(rng.Cells[i, j].value);
                                            }
                                        }
                                    }
                                    OperatedValue = Conversions.ToDouble(Operation(Values, ComboBox2.SelectedItem));
                                    rng.Cells[i, 1].value = (object)OperatedValue;
                                }
                                rng.Rows[i].Merge();
                            }
                            excelApp.DisplayAlerts = true;
                        }
                    }
                    else if (RadioButton2.Checked)
                    {
                        if (RadioButton6.Checked | RadioButton5.Checked)
                        {
                            var combinedRow = default(int);
                            if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(ComboBox3.SelectedItem, "Into Top Row", false)))
                            {
                                combinedRow = 1;
                            }
                            else if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(ComboBox3.SelectedItem, "Into Bottom Row", false)))
                            {
                                combinedRow = rng.Rows.Count;
                            }
                            for (int j = 1, loopTo10 = rng.Columns.Count; j <= loopTo10; j++)
                            {
                                if (ComboBox2.SelectedIndex <= 3)
                                {
                                    string combinedValue = "";
                                    string Separator;

                                    if (ComboBox2.SelectedIndex == 3)
                                    {
                                        Separator = Microsoft.VisualBasic.Constants.vbNewLine;
                                    }
                                    else if (CheckBox4.Checked)
                                    {
                                        Separator = ComboBox2.Text + Microsoft.VisualBasic.Constants.vbNewLine;
                                    }
                                    else
                                    {
                                        Separator = ComboBox2.Text;
                                    }

                                    for (int i = 1, loopTo11 = rng.Rows.Count - 1; i <= loopTo11; i++)
                                    {
                                        if (CheckBox3.Checked)
                                        {
                                            if (rng.Cells[i, j].value is not null)
                                            {
                                                combinedValue = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject(combinedValue, rng.Cells[i, j].Value), Separator));
                                            }
                                        }
                                        else
                                        {
                                            combinedValue = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject(combinedValue, rng.Cells[i, j].Value), Separator));
                                        }
                                    }
                                    if (CheckBox3.Checked)
                                    {
                                        if (rng.Cells[rng.Rows.Count, j].value is not null)
                                        {
                                            combinedValue = Conversions.ToString(Operators.ConcatenateObject(combinedValue, rng.Cells[rng.Rows.Count, j].Value));
                                        }
                                        else if (Strings.Len(combinedValue) >= Strings.Len(Separator))
                                        {
                                            combinedValue = Strings.Mid(combinedValue, 1, Strings.Len(combinedValue) - Strings.Len(Separator));
                                        }
                                    }
                                    else
                                    {
                                        combinedValue = Conversions.ToString(Operators.ConcatenateObject(combinedValue, rng.Cells[rng.Rows.Count, j].Value));
                                    }
                                    rng.Cells[combinedRow, j].Value = combinedValue;
                                }
                                else
                                {
                                    double OperatedValue;
                                    var Values = new double[1];
                                    int Index = -1;
                                    for (int i = 1, loopTo12 = rng.Rows.Count; i <= loopTo12; i++)
                                    {
                                        if (Information.IsNumeric(rng.Cells[i, j].Value))
                                        {
                                            if (CheckBox3.Checked)
                                            {
                                                if (rng.Cells[i, j].value is not null)
                                                {
                                                    Index = Index + 1;
                                                    Array.Resize(ref Values, Index + 1);
                                                    Values[Index] = Conversions.ToDouble(rng.Cells[i, j].value);
                                                }
                                            }
                                            else
                                            {
                                                Index = Index + 1;
                                                Array.Resize(ref Values, Index + 1);
                                                Values[Index] = Conversions.ToDouble(rng.Cells[i, j].value);
                                            }
                                        }
                                    }
                                    OperatedValue = Conversions.ToDouble(Operation(Values, ComboBox2.SelectedItem));
                                    rng.Cells[combinedRow, j].Value = (object)OperatedValue;
                                }
                            }
                            if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(ComboBox3.SelectedItem, "Into Top Row", false)))
                            {

                                for (int j = 1, loopTo13 = rng.Columns.Count; j <= loopTo13; j++)
                                {
                                    for (int i = 2, loopTo14 = rng.Rows.Count; i <= loopTo14; i++)
                                    {
                                        if (RadioButton5.Checked)
                                        {
                                            rng.Cells[i, j].Clear();
                                        }
                                    }
                                }
                            }

                            else if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(ComboBox3.SelectedItem, "Into Bottom Row", false)))
                            {
                                for (int j = 1, loopTo15 = rng.Columns.Count; j <= loopTo15; j++)
                                {
                                    for (int i = 1, loopTo16 = rng.Rows.Count - 1; i <= loopTo16; i++)
                                    {
                                        if (RadioButton5.Checked)
                                        {
                                            rng.Cells[i, j].Clear();
                                        }
                                    }
                                }
                                excelApp.DisplayAlerts = true;
                            }
                        }

                        else if (RadioButton4.Checked)
                        {
                            excelApp.DisplayAlerts = false;
                            for (int j = 1, loopTo17 = rng.Columns.Count; j <= loopTo17; j++)
                            {
                                if (ComboBox2.SelectedIndex <= 3)
                                {
                                    string combinedValue = "";
                                    string Separator;

                                    if (ComboBox2.SelectedIndex == 3)
                                    {
                                        Separator = Microsoft.VisualBasic.Constants.vbNewLine;
                                    }
                                    else if (CheckBox4.Checked)
                                    {
                                        Separator = ComboBox2.Text + Microsoft.VisualBasic.Constants.vbNewLine;
                                    }
                                    else
                                    {
                                        Separator = ComboBox2.Text;
                                    }

                                    for (int i = 1, loopTo18 = rng.Rows.Count - 1; i <= loopTo18; i++)
                                    {
                                        if (CheckBox3.Checked)
                                        {
                                            if (rng.Cells[i, j].value is not null)
                                            {
                                                combinedValue = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject(combinedValue, rng.Cells[i, j].Value), ComboBox2.SelectedItem));
                                            }
                                        }

                                        else
                                        {
                                            combinedValue = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject(combinedValue, rng.Cells[i, j].Value), ComboBox2.SelectedItem));
                                        }
                                    }
                                    if (CheckBox3.Checked)
                                    {
                                        if (rng.Cells[rng.Rows.Count, j].value is not null)
                                        {
                                            combinedValue = Conversions.ToString(Operators.ConcatenateObject(combinedValue, rng.Cells[rng.Rows.Count, j].Value));
                                        }
                                        else if (Strings.Len(combinedValue) >= Strings.Len(Separator))
                                        {
                                            combinedValue = Strings.Mid(combinedValue, 1, Strings.Len(combinedValue) - Strings.Len(Separator));
                                        }
                                    }
                                    else
                                    {
                                        combinedValue = Conversions.ToString(Operators.ConcatenateObject(combinedValue, rng.Cells[rng.Rows.Count, j].Value));
                                    }
                                    rng.Cells[1, j].value = combinedValue;
                                }

                                else
                                {
                                    double OperatedValue;
                                    var Values = new double[1];
                                    int Index = -1;
                                    for (int i = 1, loopTo19 = rng.Rows.Count; i <= loopTo19; i++)
                                    {
                                        if (Information.IsNumeric(rng.Cells[i, j].Value))
                                        {
                                            if (CheckBox3.Checked)
                                            {
                                                if (rng.Cells[i, j].value is not null)
                                                {
                                                    Index = Index + 1;
                                                    Array.Resize(ref Values, Index + 1);
                                                    Values[Index] = Conversions.ToDouble(rng.Cells[i, j].value);
                                                }
                                            }
                                            else
                                            {
                                                Index = Index + 1;
                                                Array.Resize(ref Values, Index + 1);
                                                Values[Index] = Conversions.ToDouble(rng.Cells[i, j].value);
                                            }
                                        }
                                    }
                                    OperatedValue = Conversions.ToDouble(Operation(Values, ComboBox2.SelectedItem));
                                    rng.Cells[1, j].value = (object)OperatedValue;
                                }
                                rng.Columns[j].Merge();
                            }
                            excelApp.DisplayAlerts = true;
                        }
                    }

                    else if (RadioButton3.Checked)
                    {

                        var combinedRow = default(int);
                        var combinedColumn = default(int);
                        string combinedValue = "";
                        double OperatedValue;
                        var Values = new double[1];
                        int Index = -1;
                        string Separator;

                        if (ComboBox2.SelectedIndex == 3)
                        {
                            Separator = Microsoft.VisualBasic.Constants.vbNewLine;
                        }
                        else if (CheckBox4.Checked)
                        {
                            Separator = ComboBox2.Text + Microsoft.VisualBasic.Constants.vbNewLine;
                        }
                        else
                        {
                            Separator = ComboBox2.Text;
                        }

                        for (int i = 1, loopTo20 = rng.Rows.Count - 1; i <= loopTo20; i++)
                        {
                            for (int j = 1, loopTo21 = rng.Columns.Count; j <= loopTo21; j++)
                            {
                                if (ComboBox2.SelectedIndex <= 3)
                                {
                                    if (CheckBox3.Checked)
                                    {
                                        if (rng.Cells[i, j].value is not null)
                                        {
                                            combinedValue = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject(combinedValue, rng.Cells[i, j].Value), Separator));
                                        }
                                    }

                                    else
                                    {
                                        combinedValue = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject(combinedValue, rng.Cells[i, j].Value), Separator));
                                    }
                                }
                                else if (Information.IsNumeric(rng.Cells[i, j].Value))
                                {
                                    if (CheckBox3.Checked)
                                    {
                                        if (rng.Cells[i, j].value is not null)
                                        {
                                            Index = Index + 1;
                                            Array.Resize(ref Values, Index + 1);
                                            Values[Index] = Conversions.ToDouble(rng.Cells[i, j].value);
                                        }
                                    }
                                    else
                                    {
                                        Index = Index + 1;
                                        Array.Resize(ref Values, Index + 1);
                                        Values[Index] = Conversions.ToDouble(rng.Cells[i, j].value);
                                    }
                                }
                            }
                        }

                        for (int j = 1, loopTo22 = rng.Columns.Count - 1; j <= loopTo22; j++)
                        {
                            if (ComboBox2.SelectedIndex <= 3)
                            {
                                if (CheckBox3.Checked)
                                {
                                    if (rng.Cells[rng.Rows.Count, j].value is not null)
                                    {
                                        combinedValue = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject(combinedValue, rng.Cells[rng.Rows.Count, j].Value), Separator));
                                    }
                                }

                                else
                                {
                                    combinedValue = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject(combinedValue, rng.Cells[rng.Rows.Count, j].Value), Separator));
                                }
                            }
                            else if (Information.IsNumeric(rng.Cells[rng.Rows.Count, j].Value))
                            {
                                if (CheckBox3.Checked)
                                {
                                    if (rng.Cells[rng.Rows.Count, j].value is not null)
                                    {
                                        Index = Index + 1;
                                        Array.Resize(ref Values, Index + 1);
                                        Values[Index] = Conversions.ToDouble(rng.Cells[rng.Rows.Count, j].value);
                                    }
                                }
                                else
                                {
                                    Index = Index + 1;
                                    Array.Resize(ref Values, Index + 1);
                                    Values[Index] = Conversions.ToDouble(rng.Cells[rng.Rows.Count, j].value);
                                }
                            }
                        }

                        if (ComboBox2.SelectedIndex <= 3)
                        {
                            if (CheckBox3.Checked)
                            {
                                if (rng.Cells[rng.Rows.Count, rng.Columns.Count].value is not null)
                                {
                                    combinedValue = Conversions.ToString(Operators.ConcatenateObject(combinedValue, rng.Cells[rng.Rows.Count, rng.Columns.Count].Value));
                                }
                                else if (Strings.Len(combinedValue) >= Strings.Len(Separator))
                                {
                                    combinedValue = Strings.Mid(combinedValue, 1, Strings.Len(combinedValue) - Strings.Len(Separator));
                                }
                            }

                            else
                            {
                                combinedValue = Conversions.ToString(Operators.ConcatenateObject(combinedValue, rng.Cells[rng.Rows.Count, rng.Columns.Count].Value));
                            }
                        }
                        else if (Information.IsNumeric(rng.Cells[rng.Rows.Count, rng.Columns.Count].Value))
                        {
                            if (CheckBox3.Checked)
                            {
                                if (rng.Cells[rng.Rows.Count, rng.Columns.Count].value is not null)
                                {
                                    Index = Index + 1;
                                    Array.Resize(ref Values, Index + 1);
                                    Values[Index] = Conversions.ToDouble(rng.Cells[rng.Rows.Count, rng.Columns.Count].value);
                                }
                            }
                            else
                            {
                                Index = Index + 1;
                                Array.Resize(ref Values, Index + 1);
                                Values[Index] = Conversions.ToDouble(rng.Cells[rng.Rows.Count, rng.Columns.Count].value);
                            }
                        }

                        OperatedValue = Conversions.ToDouble(Operation(Values, ComboBox2.SelectedItem));
                        if (RadioButton6.Checked | RadioButton5.Checked)
                        {
                            if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(ComboBox3.SelectedItem, "Into Top-Left Cell", false)))
                            {
                                combinedRow = 1;
                                combinedColumn = 1;
                            }
                            else if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(ComboBox3.SelectedItem, "Into Top-Right Cell", false)))
                            {
                                combinedRow = 1;
                                combinedColumn = rng.Columns.Count;
                            }
                            else if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(ComboBox3.SelectedItem, "Into Bottom-Left Cell", false)))
                            {
                                combinedRow = rng.Rows.Count;
                                combinedColumn = 1;
                            }
                            else if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(ComboBox3.SelectedItem, "Into Bottom-Right Cell", false)))
                            {
                                combinedRow = rng.Rows.Count;
                                combinedColumn = rng.Columns.Count;
                            }

                            for (int i = 1, loopTo23 = rng.Rows.Count; i <= loopTo23; i++)
                            {
                                for (int j = 1, loopTo24 = rng.Columns.Count; j <= loopTo24; j++)
                                {
                                    if (i == combinedRow & j == combinedColumn)
                                    {
                                        if (ComboBox2.SelectedIndex <= 3)
                                        {
                                            rng.Cells[i, j].value = combinedValue;
                                        }
                                        else
                                        {
                                            rng.Cells[i, j].value = (object)OperatedValue;
                                        }
                                    }


                                    else if (RadioButton5.Checked)
                                    {
                                        rng.Cells[i, j].Clear();
                                    }
                                }
                            }
                        }

                        else if (RadioButton4.Checked)
                        {
                            if (ComboBox2.SelectedIndex <= 3)
                            {
                                rng.Cells[1, 1].value = combinedValue;
                            }
                            else
                            {
                                rng.Cells[1, 1].value = (object)OperatedValue;
                            }
                            excelApp.DisplayAlerts = false;
                            rng.Merge();
                            excelApp.DisplayAlerts = true;
                        }
                    }
                    for (int j = 1, loopTo25 = rng.Columns.Count; j <= loopTo25; j++)
                        rng.Columns[j].Autofit();

                    if (CheckBox1.Checked == false)
                    {
                        rng.ClearFormats();
                    }

                    if (CheckBox4.Checked == false)
                    {
                        for (int i = 1, loopTo26 = rng.Rows.Count; i <= loopTo26; i++)
                            rng.Rows[i].Autofit();
                    }

                    Close();

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
                excelApp = Globals.ThisAddIn.Application;
                workBook = excelApp.ActiveWorkbook;
                workSheet = (Excel.Worksheet)workBook.ActiveSheet;
                TextBox1.SelectionStart = TextBox1.Text.Length;
                TextBox1.ScrollToCaret();
                rng = workSheet.get_Range(TextBox1.Text);
                TextBoxChanged = true;
                rng.Select();
                Display();
                TextBoxChanged = false;
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
        private void ComboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                Display();
            }
            catch (Exception ex)
            {
            }
        }

        private void RadioButton1_CheckedChanged(object sender, EventArgs e)
        {

            try
            {
                if (RadioButton1.Checked)
                {

                    ComboBox3.Text = "";
                    ComboBox3.Items.Clear();
                    ComboBox3.Items.Add("Into Left Column");
                    ComboBox3.Items.Add("Into Right Column");

                    Display();

                }
            }

            catch (Exception ex)
            {

            }

        }
        private void RadioButton6_CheckedChanged(object sender, EventArgs e)
        {
            try
            {

                if (RadioButton6.Checked)
                {
                    if (ComboBox3.Enabled == false)
                    {
                        ComboBox3.Enabled = true;
                        Label3.Enabled = true;
                        ComboBox3.Items.Clear();
                        if (RadioButton1.Checked)
                        {
                            ComboBox3.Items.Add("Into Left Column");
                            ComboBox3.Items.Add("Into Right Column");
                        }
                        else if (RadioButton2.Checked)
                        {
                            ComboBox3.Items.Add("Into Top Row");
                            ComboBox3.Items.Add("Into Bottom Row");
                        }
                        else if (RadioButton3.Checked)
                        {
                            ComboBox3.Items.Add("Into Top-Left Cell");
                            ComboBox3.Items.Add("Into Top-Right Cell");
                            ComboBox3.Items.Add("Into Bottom-Left Cell");
                            ComboBox3.Items.Add("Into Bottom-Right Cell");
                        }

                    }

                    Display();

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
                if (RadioButton5.Checked)
                {

                    if (ComboBox3.Enabled == false)
                    {

                        ComboBox3.Enabled = true;
                        Label3.Enabled = true;
                        ComboBox3.Items.Clear();

                        if (RadioButton1.Checked)
                        {
                            ComboBox3.Items.Add("Into Left Column");
                            ComboBox3.Items.Add("Into Right Column");
                        }
                        else if (RadioButton2.Checked)
                        {
                            ComboBox3.Items.Add("Into Top Row");
                            ComboBox3.Items.Add("Into Bottom Row");
                        }
                        else if (RadioButton3.Checked)
                        {
                            ComboBox3.Items.Add("Into Top-Left Cell");
                            ComboBox3.Items.Add("Into Top-Right Cell");
                            ComboBox3.Items.Add("Into Bottom-Left Cell");
                            ComboBox3.Items.Add("Into Bottom-Right Cell");
                        }
                    }
                    Display();
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
                if (RadioButton4.Checked)
                {

                    ComboBox3.SelectedText = "";
                    ComboBox3.Items.Clear();
                    ComboBox3.Enabled = false;
                    Label3.Enabled = false;

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
                if (RadioButton2.Checked)
                {

                    ComboBox3.Text = "";
                    ComboBox3.Items.Clear();
                    ComboBox3.Items.Add("Into Top Row");
                    ComboBox3.Items.Add("Into Bottom Row");

                    Display();

                }
            }

            catch (Exception ex)
            {

            }
        }

        private void RadioButton3_CheckedChanged(object sender, EventArgs e)
        {

            try
            {
                if (RadioButton3.Checked)
                {

                    ComboBox3.Text = "";
                    ComboBox3.Items.Clear();
                    ComboBox3.Items.Add("Into Top-Left Cell");
                    ComboBox3.Items.Add("Into Top-Right Cell");
                    ComboBox3.Items.Add("Into Bottom-Left Cell");
                    ComboBox3.Items.Add("Into Bottom-Right Cell");

                    Display();

                }
            }

            catch (Exception ex)
            {

            }

        }

        private void AutoSelection_Click(object sender, EventArgs e)
        {

            try
            {
                FocusedTextBox = 1;
                Hide();

                excelApp = Globals.ThisAddIn.Application;
                workBook = excelApp.ActiveWorkbook;

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

        private void Selection_Click(object sender, EventArgs e)
        {
            try
            {
                FocusedTextBox = 1;
                Hide();

                excelApp = Globals.ThisAddIn.Application;
                workBook = excelApp.ActiveWorkbook;

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

        private void AutoSelection_GotFocus(object sender, EventArgs e)
        {
            try
            {
                FocusedTextBox = 1;
            }
            catch (Exception ex)
            {

            }
        }

        private void Selection_GotFocus(object sender, EventArgs e)
        {
            try
            {
                FocusedTextBox = 1;
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

        private void AutoSelection_KeyDown(object sender, KeyEventArgs e)
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

        private void CheckBox3_KeyDown(object sender, KeyEventArgs e)
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

        private void CheckBox4_KeyDown(object sender, KeyEventArgs e)
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

        private void ComboBox2_KeyDown(object sender, KeyEventArgs e)
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

        private void ComboBox3_KeyDown(object sender, KeyEventArgs e)
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

        private void CustomGroupBox1_KeyDown(object sender, KeyEventArgs e)
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

        private void CustomGroupBox2_KeyDown(object sender, KeyEventArgs e)
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

        private void CustomGroupBox3_KeyDown(object sender, KeyEventArgs e)
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

        private void CustomGroupBox7_KeyDown(object sender, KeyEventArgs e)
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

        private void Label2_KeyDown(object sender, KeyEventArgs e)
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

        private void PictureBox2_KeyDown(object sender, KeyEventArgs e)
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

        private void PictureBox3_KeyDown(object sender, KeyEventArgs e)
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

        private void PictureBox7_KeyDown(object sender, KeyEventArgs e)
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

        private void RadioButton1_KeyDown(object sender, KeyEventArgs e)
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

        private void RadioButton2_KeyDown(object sender, KeyEventArgs e)
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

        private void RadioButton3_KeyDown(object sender, KeyEventArgs e)
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

        private void RadioButton4_KeyDown(object sender, KeyEventArgs e)
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

        private void RadioButton5_KeyDown(object sender, KeyEventArgs e)
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

        private void RadioButton6_KeyDown(object sender, KeyEventArgs e)
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

        private void Selection_KeyDown(object sender, KeyEventArgs e)
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

        private void Form18_CombineRanges_Load(object sender, EventArgs e)
        {
            try
            {

                excelApp = Globals.ThisAddIn.Application;

                excelApp.SheetSelectionChange += excelApp_SheetSelectionChange;

                opened = opened + 1;
            }

            catch (Exception ex)
            {

            }
        }

        private void excelApp_SheetSelectionChange(object Sh, Range Target)
        {

            try
            {

                excelApp = Globals.ThisAddIn.Application;
                Range selectedRange;
                selectedRange = (Range)excelApp.Selection;
                if (FocusedTextBox == 1)
                {
                    TextBox1.Text = selectedRange.get_Address();
                    workSheet = (Excel.Worksheet)workBook.ActiveSheet;
                    rng = selectedRange;
                    TextBox1.Focus();
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

        private void CheckBox3_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                Display();
            }
            catch (Exception ex)
            {
            }
        }

        private void CheckBox4_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                Display();
            }
            catch (Exception ex)
            {
            }
        }

        private void ComboBox2_TextChanged(object sender, EventArgs e)
        {
            try
            {
                Display();
            }
            catch (Exception ex)
            {
            }
        }

        private void Form18_CombineRanges_Closing(object sender, CancelEventArgs e)
        {
            GlobalModule.form_flag = false;
        }

        private void Form18_CombineRanges_Disposed(object sender, EventArgs e)
        {
            GlobalModule.form_flag = false;
        }

        private void Form18_CombineRanges_Shown(object sender, EventArgs e)
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