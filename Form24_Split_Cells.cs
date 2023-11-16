using System;
using System.ComponentModel;
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

    public partial class Form24_Split_Cells
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
        private Excel.Worksheet workSheet2;
        public Excel.Worksheet OpenSheet;
        private Range rng;
        private Range rng2;
        private Range selectedRange;

        private int opened;
        private int FocusedTextBox;
        private bool TextBoxChanged;

        public Form24_Split_Cells()
        {
            InitializeComponent();
        }

        [DllImport("user32")]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
        private const uint SWP_NOMOVE = 0x2U;
        private const uint SWP_NOSIZE = 0x1U;
        private const uint SWP_NOACTIVATE = 0x10U;
        private const int HWND_TOPMOST = -1;

        private object MaxOfColumn(Range cRng)
        {
            object MaxOfColumnRet = default;

            int max;
            max = Strings.Len(cRng.Cells[1, 1].value);

            for (int i = 2, loopTo = cRng.Rows.Count; i <= loopTo; i++)
            {
                if (Strings.Len(cRng.Cells[i, 1].value) > max)
                {
                    max = Strings.Len(cRng.Cells[i, 1].value);
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
        private object FindMax(object Arr)
        {
            object FindMaxRet = default;

            int Max = Conversions.ToInteger(Arr((object)Information.LBound((Array)Arr)));

            for (int i = Information.LBound((Array)Arr) + 1, loopTo = Information.UBound((Array)Arr); i <= loopTo; i++)
            {
                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectGreater(Arr((object)i), Max, false)))
                {
                    Max = Conversions.ToInteger(Arr((object)i));
                }
            }
            FindMaxRet = Max;
            return FindMaxRet;

        }

        private object SeparateNumberText(string Str)
        {
            object SeparateNumberTextRet = default;

            var Output = new string[2];
            Output[0] = "";
            Output[1] = "";

            for (int i = 1, loopTo = Strings.Len(Str); i <= loopTo; i++)
            {
                if (Information.IsNumeric(Strings.Mid(Str, i, 1)))
                {
                    Output[0] = Output[0] + Strings.Mid(Str, i, 1);
                }
                else
                {
                    Output[1] = Output[1] + Strings.Mid(Str, i, 1);
                }
            }

            SeparateNumberTextRet = Output;
            return SeparateNumberTextRet;

        }
        public int CountSeparator(string source, string separator)
        {
            int CountSeparatorRet = default;

            int count = 0;
            int Position = 1;

            for (int i = 1, loopTo = Strings.Len(source); i <= loopTo; i++)
            {
                if ((Strings.Mid(source, i, Strings.Len(separator)) ?? "") == (separator ?? ""))
                {
                    if (i - Position > 0)
                    {
                        count = count + 1;
                    }
                    Position = i + Strings.Len(separator);
                }
            }

            if (Position <= Strings.Len(source))
            {
                count = count + 1;
            }

            CountSeparatorRet = count;
            return CountSeparatorRet;

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

        private void Display()
        {

            try
            {
                TextBoxChanged = true;

                CustomPanel1.Controls.Clear();
                CustomPanel2.Controls.Clear();

                Range displayRng;

                if (rng.Rows.Count > 50)
                {
                    displayRng = (Range)rng.Rows["1:50"];
                }
                else
                {
                    displayRng = rng;
                }

                int r = displayRng.Rows.Count;
                int c = displayRng.Columns.Count;

                double Height;
                double BaseWidth;
                double Width;

                if (r <= 4)
                {
                    Height = CustomPanel1.Height / (double)displayRng.Rows.Count;
                }
                else
                {
                    Height = 119d / 4d;
                }

                BaseWidth = 260d / 3d;

                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectGreater(Operators.DivideObject(Operators.MultiplyObject(MaxOfColumn(displayRng), BaseWidth), 10), CustomPanel1.Width, false)))
                {
                    Width = Conversions.ToDouble(Operators.DivideObject(Operators.MultiplyObject(MaxOfColumn(displayRng), BaseWidth), 10));
                }
                else
                {
                    Width = CustomPanel1.Width;
                }

                for (int i = 1, loopTo = r; i <= loopTo; i++)
                {
                    var label = new System.Windows.Forms.Label();
                    label.Text = Conversions.ToString(displayRng.Cells[i, 1].Value);
                    label.Location = new System.Drawing.Point(0, (int)Math.Round((i - 1) * Height));
                    label.Height = (int)Math.Round(Height);
                    label.Width = (int)Math.Round(Width);
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
                    CustomPanel1.Controls.Add(label);
                }

                CustomPanel1.AutoScroll = true;

                bool X1 = RadioButton1.Checked;
                bool X2 = RadioButton2.Checked;
                bool X3 = RadioButton3.Checked;
                bool X7 = RadioButton7.Checked;
                bool X8 = RadioButton8.Checked;
                bool X9 = RadioButton9.Checked;
                bool X10 = RadioButton10.Checked;
                bool X11 = RadioButton11.Checked;


                if ((X1 | X2) & (X3 | X7 | X8 | X9 | X10 | X11))
                {

                    int SplitColumn = 1;
                    double Ordinate;

                    if (X7 | X8 | X9 | X10)
                    {

                        string Separator = "";
                        if (X7)
                        {
                            Separator = ";";
                        }
                        else if (X8)
                        {
                            Separator = Microsoft.VisualBasic.Constants.vbNewLine;
                        }
                        else if (X9)
                        {
                            Separator = " ";
                        }
                        else if (X10)
                        {
                            Separator = ComboBox2.Text;
                        }

                        if (X2)
                        {

                            var Lengths = new int[r];

                            int Index;
                            int position;

                            for (int i = 1, loopTo1 = r; i <= loopTo1; i++)
                            {
                                string source = Conversions.ToString(displayRng.Cells[i, SplitColumn].value);
                                Lengths[i - 1] = CountSeparator(source, Separator);
                            }

                            int TotalColumn = Conversions.ToInteger(FindMax(Lengths));
                            var SplitValues = new string[r, TotalColumn];

                            for (int i = 1, loopTo2 = r; i <= loopTo2; i++)
                            {
                                string source = Conversions.ToString(displayRng.Cells[i, SplitColumn].value);
                                position = 1;
                                Index = -1;
                                for (int k = 1, loopTo3 = Strings.Len(source); k <= loopTo3; k++)
                                {
                                    if ((Strings.Mid(source, k, Strings.Len(Separator)) ?? "") == (Separator ?? ""))
                                    {
                                        if (k - position > 0)
                                        {
                                            Index = Index + 1;
                                            SplitValues[i - 1, Index] = Strings.Mid(source, position, k - position);
                                        }
                                        position = k + Strings.Len(Separator);
                                    }
                                }
                                if (position <= Strings.Len(source))
                                {
                                    Index = Index + 1;
                                    SplitValues[i - 1, Index] = Strings.Mid(source, position, Strings.Len(source) - position + 1);
                                }
                            }

                            Ordinate = 0d;

                            for (int j = Information.LBound(SplitValues, 2), loopTo4 = Information.UBound(SplitValues, 2); j <= loopTo4; j++)
                            {
                                var NewColumn = new string[r];
                                for (int i = Information.LBound(SplitValues, 1), loopTo5 = Information.UBound(SplitValues, 1); i <= loopTo5; i++)
                                    NewColumn[i] = SplitValues[i, j];
                                if (TotalColumn == 1)
                                {
                                    Width = CustomPanel2.Width;
                                }
                                else
                                {
                                    Width = Conversions.ToDouble(Operators.DivideObject(Operators.MultiplyObject(MaxOfArray(NewColumn), BaseWidth), 10));
                                }
                                for (int i = Information.LBound(SplitValues, 1), loopTo6 = Information.UBound(SplitValues, 1); i <= loopTo6; i++)
                                {
                                    var label = new System.Windows.Forms.Label();
                                    label.Text = SplitValues[i, j];
                                    label.Location = new System.Drawing.Point((int)Math.Round(Ordinate), (int)Math.Round(i * Height));
                                    label.Height = (int)Math.Round(Height);
                                    label.Width = (int)Math.Round(Width);
                                    label.BorderStyle = BorderStyle.FixedSingle;
                                    label.TextAlign = ContentAlignment.MiddleCenter;

                                    if (CheckBox1.Checked == true)
                                    {

                                        Range cell = (Range)displayRng.Cells[i + 1, 1];
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
                                Ordinate = Ordinate + Width;
                            }
                        }

                        else if (X1)
                        {

                            var Lengths = new int[r];

                            int Index;
                            int position;

                            for (int i = 1, loopTo7 = r; i <= loopTo7; i++)
                            {
                                string source = Conversions.ToString(displayRng.Cells[i, SplitColumn].value);
                                Lengths[i - 1] = CountSeparator(source, Separator);
                            }

                            int TotalColumn = Conversions.ToInteger(FindMax(Lengths));
                            var SplitValues = new string[r, TotalColumn];

                            for (int i = 1, loopTo8 = r; i <= loopTo8; i++)
                            {
                                string source = Conversions.ToString(displayRng.Cells[i, SplitColumn].value);
                                position = 1;
                                Index = -1;
                                for (int k = 1, loopTo9 = Strings.Len(source); k <= loopTo9; k++)
                                {
                                    if ((Strings.Mid(source, k, Strings.Len(Separator)) ?? "") == (Separator ?? ""))
                                    {
                                        if (k - position > 0)
                                        {
                                            Index = Index + 1;
                                            SplitValues[i - 1, Index] = Strings.Mid(source, position, k - position);
                                        }
                                        position = k + Strings.Len(Separator);
                                    }
                                }
                                if (position <= Strings.Len(source))
                                {
                                    Index = Index + 1;
                                    SplitValues[i - 1, Index] = Strings.Mid(source, position, Strings.Len(source) - position + 1);
                                }
                            }

                            Ordinate = 0d;

                            for (int i = Information.LBound(SplitValues, 1), loopTo10 = Information.UBound(SplitValues, 1); i <= loopTo10; i++)
                            {
                                var NewColumn = new string[TotalColumn];
                                for (int j = Information.LBound(SplitValues, 2), loopTo11 = Information.UBound(SplitValues, 2); j <= loopTo11; j++)
                                    NewColumn[j] = SplitValues[i, j];
                                if (TotalColumn * Height < CustomPanel2.Height)
                                {
                                    Height = CustomPanel2.Height / (double)TotalColumn;
                                }
                                if (r == 1)
                                {
                                    Width = CustomPanel2.Width;
                                }
                                else
                                {
                                    Width = Conversions.ToDouble(Operators.DivideObject(Operators.MultiplyObject(MaxOfArray(NewColumn), BaseWidth), 10));
                                }
                                for (int j = Information.LBound(SplitValues, 2), loopTo12 = Information.UBound(SplitValues, 2); j <= loopTo12; j++)
                                {
                                    var label = new System.Windows.Forms.Label();
                                    label.Text = SplitValues[i, j];
                                    label.Location = new System.Drawing.Point((int)Math.Round(Ordinate), (int)Math.Round(j * Height));
                                    label.Height = (int)Math.Round(Height);
                                    label.Width = (int)Math.Round(Width);
                                    label.BorderStyle = BorderStyle.FixedSingle;
                                    label.TextAlign = ContentAlignment.MiddleCenter;

                                    if (CheckBox1.Checked == true)
                                    {

                                        Range cell = (Range)displayRng.Cells[i + 1, 1];
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
                                Ordinate = Ordinate + Width;
                            }
                        }
                    }

                    else if (X3)
                    {

                        if (X2)
                        {

                            var Numbers = new string[r];
                            var Texts = new string[r];

                            for (int i = 1, loopTo13 = r; i <= loopTo13; i++)
                            {
                                string source = Conversions.ToString(displayRng.Cells[i, SplitColumn].value);
                                var NumberText = new string[2];
                                NumberText = (string[])SeparateNumberText(source);
                                Numbers[i - 1] = NumberText[0];
                                Texts[i - 1] = NumberText[1];
                            }

                            double NumbersWidth = Conversions.ToDouble(Operators.DivideObject(Operators.MultiplyObject(MaxOfArray(Numbers), BaseWidth), 10));
                            double TextsWidth = Conversions.ToDouble(Operators.DivideObject(Operators.MultiplyObject(MaxOfArray(Texts), BaseWidth), 10));

                            if (NumbersWidth + TextsWidth < CustomPanel2.Width)
                            {
                                NumbersWidth = CustomPanel2.Width / 2d;
                                TextsWidth = CustomPanel2.Width / 2d;
                            }

                            Ordinate = 0d;

                            for (int i = Information.LBound(Numbers), loopTo14 = Information.UBound(Numbers); i <= loopTo14; i++)
                            {
                                var label = new System.Windows.Forms.Label();
                                label.Text = Numbers[i];
                                label.Location = new System.Drawing.Point((int)Math.Round(Ordinate), (int)Math.Round(i * Height));
                                label.Height = (int)Math.Round(Height);
                                label.Width = (int)Math.Round(NumbersWidth);
                                label.BorderStyle = BorderStyle.FixedSingle;
                                label.TextAlign = ContentAlignment.MiddleCenter;

                                if (CheckBox1.Checked == true)
                                {

                                    Range cell = (Range)displayRng.Cells[i + 1, 1];
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
                            Ordinate = Ordinate + NumbersWidth;

                            for (int i = Information.LBound(Texts), loopTo15 = Information.UBound(Texts); i <= loopTo15; i++)
                            {
                                var label = new System.Windows.Forms.Label();
                                label.Text = Texts[i];
                                label.Location = new System.Drawing.Point((int)Math.Round(Ordinate), (int)Math.Round(i * Height));
                                label.Height = (int)Math.Round(Height);
                                label.Width = (int)Math.Round(TextsWidth);
                                label.BorderStyle = BorderStyle.FixedSingle;
                                label.TextAlign = ContentAlignment.MiddleCenter;

                                if (CheckBox1.Checked == true)
                                {

                                    Range cell = (Range)displayRng.Cells[i + 1, 1];
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

                        else if (X1)
                        {

                            int TotalColumn = 2;
                            var SplitValues = new string[r, TotalColumn];

                            for (int i = 1, loopTo16 = r; i <= loopTo16; i++)
                            {
                                string source = Conversions.ToString(displayRng.Cells[i, SplitColumn].value);
                                var NumberText = new string[2];
                                NumberText = (string[])SeparateNumberText(source);
                                SplitValues[i - 1, 0] = NumberText[0];
                                SplitValues[i - 1, 1] = NumberText[1];
                            }

                            Ordinate = 0d;

                            for (int i = Information.LBound(SplitValues, 1), loopTo17 = Information.UBound(SplitValues, 1); i <= loopTo17; i++)
                            {
                                var NewColumn = new string[TotalColumn];
                                for (int j = Information.LBound(SplitValues, 2), loopTo18 = Information.UBound(SplitValues, 2); j <= loopTo18; j++)
                                    NewColumn[j] = SplitValues[i, j];
                                Height = CustomPanel2.Height / 2d;
                                if (r == 1)
                                {
                                    Width = CustomPanel2.Width;
                                }
                                else
                                {
                                    Width = Conversions.ToDouble(Operators.DivideObject(Operators.MultiplyObject(MaxOfArray(NewColumn), BaseWidth), 10));
                                }
                                for (int j = Information.LBound(SplitValues, 2), loopTo19 = Information.UBound(SplitValues, 2); j <= loopTo19; j++)
                                {
                                    var label = new System.Windows.Forms.Label();
                                    label.Text = SplitValues[i, j];
                                    label.Location = new System.Drawing.Point((int)Math.Round(Ordinate), (int)Math.Round(j * Height));
                                    label.Height = (int)Math.Round(Height);
                                    label.Width = (int)Math.Round(Width);
                                    label.BorderStyle = BorderStyle.FixedSingle;
                                    label.TextAlign = ContentAlignment.MiddleCenter;

                                    if (CheckBox1.Checked == true)
                                    {

                                        Range cell = (Range)displayRng.Cells[i + 1, 1];
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
                                Ordinate = Ordinate + Width;
                            }
                        }
                    }

                    else if (X11)
                    {

                        int W;

                        if (string.IsNullOrEmpty(TextBox3.Text))
                        {
                            W = 1;
                        }
                        else
                        {
                            W = Conversions.ToInteger(Conversion.Int(TextBox3.Text));
                        }

                        if (X2)
                        {

                            var Lengths = new int[r];

                            int Index;

                            for (int i = 1, loopTo20 = r; i <= loopTo20; i++)
                            {
                                string source = Conversions.ToString(displayRng.Cells[i, SplitColumn].value);
                                if (Strings.Len(source) % W == 0)
                                {
                                    Lengths[i - 1] = (int)Math.Round(Conversion.Int(Strings.Len(source) / (double)W));
                                }
                                else
                                {
                                    Lengths[i - 1] = (int)Math.Round(Conversion.Int(Strings.Len(source) / (double)W) + 1d);
                                }
                            }

                            int TotalColumn = Conversions.ToInteger(FindMax(Lengths));
                            var SplitValues = new string[r, TotalColumn];

                            for (int i = 1, loopTo21 = r; i <= loopTo21; i++)
                            {
                                string source = Conversions.ToString(displayRng.Cells[i, SplitColumn].value);
                                Index = -1;
                                for (double k = 1d, loopTo22 = Conversion.Int(Strings.Len(source) / (double)W); k <= loopTo22; k++)
                                {
                                    Index = Index + 1;
                                    SplitValues[i - 1, Index] = Strings.Mid(source, (int)Math.Round(W * (k - 1d) + 1d), W);
                                }
                                if (Strings.Len(source) % W != 0)
                                {
                                    Index = Index + 1;
                                    SplitValues[i - 1, Index] = Strings.Mid(source, Strings.Len(source) - Strings.Len(source) % W + 1, Strings.Len(source) % W);
                                }
                            }

                            Ordinate = 0d;

                            for (int j = Information.LBound(SplitValues, 2), loopTo23 = Information.UBound(SplitValues, 2); j <= loopTo23; j++)
                            {
                                var NewColumn = new string[r];
                                for (int i = Information.LBound(SplitValues, 1), loopTo24 = Information.UBound(SplitValues, 1); i <= loopTo24; i++)
                                    NewColumn[i] = SplitValues[i, j];
                                if (TotalColumn == 1)
                                {
                                    Width = CustomPanel2.Width;
                                }
                                else
                                {
                                    Width = Conversions.ToDouble(Operators.DivideObject(Operators.MultiplyObject(MaxOfArray(NewColumn), BaseWidth), 10));
                                }
                                for (int i = Information.LBound(SplitValues, 1), loopTo25 = Information.UBound(SplitValues, 1); i <= loopTo25; i++)
                                {
                                    var label = new System.Windows.Forms.Label();
                                    label.Text = SplitValues[i, j];
                                    label.Location = new System.Drawing.Point((int)Math.Round(Ordinate), (int)Math.Round(i * Height));
                                    label.Height = (int)Math.Round(Height);
                                    label.Width = (int)Math.Round(Width);
                                    label.BorderStyle = BorderStyle.FixedSingle;
                                    label.TextAlign = ContentAlignment.MiddleCenter;

                                    if (CheckBox1.Checked == true)
                                    {

                                        Range cell = (Range)displayRng.Cells[i + 1, 1];
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
                                Ordinate = Ordinate + Width;
                            }
                        }


                        else if (X1)
                        {

                            var Lengths = new int[r];

                            int Index;

                            for (int i = 1, loopTo26 = r; i <= loopTo26; i++)
                            {
                                string source = Conversions.ToString(displayRng.Cells[i, SplitColumn].value);
                                if (Strings.Len(source) % W == 0)
                                {
                                    Lengths[i - 1] = (int)Math.Round(Conversion.Int(Strings.Len(source) / (double)W));
                                }
                                else
                                {
                                    Lengths[i - 1] = (int)Math.Round(Conversion.Int(Strings.Len(source) / (double)W) + 1d);
                                }
                            }

                            int TotalColumn = Conversions.ToInteger(FindMax(Lengths));
                            var SplitValues = new string[r, TotalColumn];

                            for (int i = 1, loopTo27 = r; i <= loopTo27; i++)
                            {
                                string source = Conversions.ToString(displayRng.Cells[i, SplitColumn].value);
                                Index = -1;
                                for (double k = 1d, loopTo28 = Conversion.Int(Strings.Len(source) / (double)W); k <= loopTo28; k++)
                                {
                                    Index = Index + 1;
                                    SplitValues[i - 1, Index] = Strings.Mid(source, (int)Math.Round(W * (k - 1d) + 1d), W);
                                }
                                if (Strings.Len(source) % W != 0)
                                {
                                    Index = Index + 1;
                                    SplitValues[i - 1, Index] = Strings.Mid(source, Strings.Len(source) - Strings.Len(source) % W + 1, Strings.Len(source) % W);
                                }
                            }

                            Ordinate = 0d;

                            for (int i = Information.LBound(SplitValues, 1), loopTo29 = Information.UBound(SplitValues, 1); i <= loopTo29; i++)
                            {
                                var NewColumn = new string[TotalColumn];
                                for (int j = Information.LBound(SplitValues, 2), loopTo30 = Information.UBound(SplitValues, 2); j <= loopTo30; j++)
                                    NewColumn[j] = SplitValues[i, j];
                                if (TotalColumn * Height < CustomPanel2.Height)
                                {
                                    Height = CustomPanel2.Height / (double)TotalColumn;
                                }
                                if (r == 1)
                                {
                                    Width = CustomPanel2.Width;
                                }
                                else
                                {
                                    Width = Conversions.ToDouble(Operators.DivideObject(Operators.MultiplyObject(MaxOfArray(NewColumn), BaseWidth), 10));
                                }
                                for (int j = Information.LBound(SplitValues, 2), loopTo31 = Information.UBound(SplitValues, 2); j <= loopTo31; j++)
                                {
                                    var label = new System.Windows.Forms.Label();
                                    label.Text = SplitValues[i, j];
                                    label.Location = new System.Drawing.Point((int)Math.Round(Ordinate), (int)Math.Round(j * Height));
                                    label.Height = (int)Math.Round(Height);
                                    label.Width = (int)Math.Round(Width);
                                    label.BorderStyle = BorderStyle.FixedSingle;
                                    label.TextAlign = ContentAlignment.MiddleCenter;

                                    if (CheckBox1.Checked == true)
                                    {

                                        Range cell = (Range)displayRng.Cells[i + 1, 1];
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
                                Ordinate = Ordinate + Width;
                            }

                        }

                    }

                    CustomPanel2.AutoScroll = true;

                }

                TextBoxChanged = false;
            }

            catch (Exception ex)
            {

            }

        }
        private void Button2_Click(object sender, EventArgs e)
        {

            try
            {
                TextBoxChanged = true;

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
                    MessageBox.Show("Select a Valid Source Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    TextBox1.Focus();
                    workSheet.Activate();
                    rng.Select();
                    return;
                }

                if (RadioButton4.Checked == false & RadioButton5.Checked == false)
                {
                    MessageBox.Show("Enter a Destination Cell.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    workSheet.Activate();
                    rng.Select();
                    return;
                }

                if (RadioButton4.Checked == true & (string.IsNullOrEmpty(TextBox4.Text) | IsValidExcelCellReference(TextBox4.Text) == false))
                {
                    MessageBox.Show("Enter a valid Destination Cell.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    workSheet.Activate();
                    rng.Select();
                    return;
                }

                bool X1 = RadioButton1.Checked;
                bool X2 = RadioButton2.Checked;
                bool X3 = RadioButton3.Checked;
                bool X7 = RadioButton7.Checked;
                bool X8 = RadioButton8.Checked;
                bool X9 = RadioButton9.Checked;
                bool X10 = RadioButton10.Checked;
                bool X11 = RadioButton11.Checked;

                if (X1 == false & X2 == false)
                {
                    MessageBox.Show("Select a Split Option.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    workSheet.Activate();
                    rng.Select();
                    return;
                }

                if (X3 == false & X7 == false & X8 == false & X9 == false & X10 == false & X11 == false)
                {
                    MessageBox.Show("Select a Separator to Split the Cells.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    workSheet.Activate();
                    rng.Select();
                    return;
                }

                if (CheckBox2.Checked == true)
                {
                    workSheet.Copy(After: workBook.Sheets[workSheet.Name]);
                }

                int r = rng.Rows.Count;
                int c = rng.Columns.Count;
                string rng2Address;

                var TotalColumns = default(int);

                if (X7 | X8 | X9 | X10)
                {
                    string Separator = "";
                    var Columns = new int[r];
                    if (X7)
                    {
                        Separator = ";";
                    }
                    else if (X8)
                    {
                        Separator = Microsoft.VisualBasic.Constants.vbNewLine;
                    }
                    else if (X9)
                    {
                        Separator = " ";
                    }
                    else if (X10)
                    {
                        Separator = ComboBox2.Text;
                    }
                    for (int i = 1, loopTo = r; i <= loopTo; i++)
                        Columns[i - 1] = CountSeparator(Conversions.ToString(rng.Cells[i, 1].value), Separator);
                    TotalColumns = Conversions.ToInteger(FindMax(Columns));
                    if (X2)
                    {
                        rng2 = workSheet2.get_Range(rng2.Cells[1, 1], rng2.Cells[r, TotalColumns]);
                    }
                    else if (X1)
                    {
                        rng2 = workSheet2.get_Range(rng2.Cells[1, 1], rng2.Cells[TotalColumns, r]);
                    }
                }
                else if (X3)
                {
                    TotalColumns = 2;
                    if (X2)
                    {
                        rng2 = workSheet2.get_Range(rng2.Cells[1, 1], rng2.Cells[r, TotalColumns]);
                    }
                    else if (X1)
                    {
                        rng2 = workSheet2.get_Range(rng2.Cells[1, 1], rng2.Cells[TotalColumns, r]);
                    }
                }
                else if (X11)
                {
                    int W;
                    if (string.IsNullOrEmpty(TextBox3.Text))
                    {
                        W = 1;
                    }
                    else
                    {
                        W = Conversions.ToInteger(Conversion.Int(TextBox3.Text));
                    }
                    var Columns = new int[r];
                    for (int i = 1, loopTo1 = r; i <= loopTo1; i++)
                    {
                        if (Strings.Len(rng.Cells[i, 1].value) % W == 0)
                        {
                            Columns[i - 1] = (int)Math.Round(Conversion.Int((double)Strings.Len(rng.Cells[i, 1].value) / W));
                        }
                        else
                        {
                            Columns[i - 1] = (int)Math.Round(Conversion.Int((double)Strings.Len(rng.Cells[i, 1].value) / W) + 1d);
                        }
                    }
                    TotalColumns = Conversions.ToInteger(FindMax(Columns));
                    if (X2)
                    {
                        rng2 = workSheet2.get_Range(rng2.Cells[1, 1], rng2.Cells[r, TotalColumns]);
                    }
                    else if (X1)
                    {
                        rng2 = workSheet2.get_Range(rng2.Cells[1, 1], rng2.Cells[TotalColumns, r]);
                    }
                }

                rng2Address = rng2.get_Address();

                if (Overlap(excelApp, workSheet, workSheet2, rng, rng2) == false)
                {

                    if ((X1 | X2) & (X3 | X7 | X8 | X9 | X10 | X11))
                    {

                        int SplitColumn = 1;

                        if (X7 | X8 | X9 | X10)
                        {

                            string Separator = "";
                            if (X7)
                            {
                                Separator = ";";
                            }
                            else if (X8)
                            {
                                Separator = Microsoft.VisualBasic.Constants.vbNewLine;
                            }
                            else if (X9)
                            {
                                Separator = " ";
                            }
                            else if (X10)
                            {
                                Separator = ComboBox2.Text;
                            }

                            if (X2)
                            {

                                int Index;
                                int position;

                                for (int i = 1, loopTo2 = r; i <= loopTo2; i++)
                                {
                                    string source = Conversions.ToString(rng.Cells[i, SplitColumn].value);
                                    position = 1;
                                    Index = 0;
                                    for (int k = 1, loopTo3 = Strings.Len(source); k <= loopTo3; k++)
                                    {
                                        if ((Strings.Mid(source, k, Strings.Len(Separator)) ?? "") == (Separator ?? ""))
                                        {
                                            if (k - position > 0)
                                            {
                                                Index = Index + 1;
                                                rng2.Cells[i, Index].value = Strings.Mid(source, position, k - position);
                                            }
                                            position = k + Strings.Len(Separator);
                                        }
                                    }
                                    if (position <= Strings.Len(source))
                                    {
                                        Index = Index + 1;
                                        rng2.Cells[i, Index].value = Strings.Mid(source, position, Strings.Len(source) - position + 1);
                                    }

                                    if (CheckBox1.Checked == true)
                                    {
                                        for (int m = 1, loopTo4 = TotalColumns; m <= loopTo4; m++)
                                        {
                                            rng.Cells[i, SplitColumn].copy();
                                            rng2.Cells[i, m].PasteSpecial(XlPasteType.xlPasteFormats);
                                            rng2 = workSheet2.get_Range(rng2Address);
                                            workSheet2.Activate();
                                        }
                                    }
                                    else
                                    {
                                        for (int m = 1, loopTo5 = TotalColumns; m <= loopTo5; m++)
                                            rng2.Cells[i, m].ClearFormats();
                                    }

                                }
                                excelApp.CutCopyMode = (XlCutCopyMode)Conversions.ToInteger(false);
                            }

                            else if (X1)
                            {

                                int Index;
                                int position;

                                for (int i = 1, loopTo6 = r; i <= loopTo6; i++)
                                {
                                    string source = Conversions.ToString(rng.Cells[i, SplitColumn].value);
                                    position = 1;
                                    Index = 0;
                                    for (int k = 1, loopTo7 = Strings.Len(source); k <= loopTo7; k++)
                                    {
                                        if ((Strings.Mid(source, k, Strings.Len(Separator)) ?? "") == (Separator ?? ""))
                                        {
                                            if (k - position > 0)
                                            {
                                                Index = Index + 1;
                                                rng2.Cells[Index, i].value = Strings.Mid(source, position, k - position);
                                            }
                                            position = k + Strings.Len(Separator);
                                        }
                                    }
                                    if (position <= Strings.Len(source))
                                    {
                                        Index = Index + 1;
                                        rng2.Cells[Index, i].value = Strings.Mid(source, position, Strings.Len(source) - position + 1);
                                    }
                                    if (CheckBox1.Checked == true)
                                    {
                                        for (int m = 1, loopTo8 = TotalColumns; m <= loopTo8; m++)
                                        {
                                            rng.Cells[i, SplitColumn].copy();
                                            rng2.Cells[m, i].PasteSpecial(XlPasteType.xlPasteFormats);
                                            rng2 = workSheet2.get_Range(rng2Address);
                                            workSheet2.Activate();
                                        }
                                    }
                                    else
                                    {
                                        for (int m = 1, loopTo9 = TotalColumns; m <= loopTo9; m++)
                                            rng2.Cells[m, i].ClearFormats();
                                    }
                                }
                                excelApp.CutCopyMode = (XlCutCopyMode)Conversions.ToInteger(false);
                            }
                        }

                        else if (X3)
                        {

                            if (X2)
                            {

                                for (int i = 1, loopTo10 = r; i <= loopTo10; i++)
                                {
                                    string source = Conversions.ToString(rng.Cells[i, SplitColumn].value);
                                    var NumberText = new string[2];
                                    NumberText = (string[])SeparateNumberText(source);
                                    rng2.Cells[i, 1].value = NumberText[0];
                                    rng2.Cells[i, 2].value = NumberText[1];
                                    if (CheckBox1.Checked == true)
                                    {
                                        for (int m = 1, loopTo11 = TotalColumns; m <= loopTo11; m++)
                                        {
                                            rng.Cells[i, SplitColumn].copy();
                                            rng2.Cells[i, m].PasteSpecial(XlPasteType.xlPasteFormats);
                                            rng2 = workSheet2.get_Range(rng2Address);
                                            workSheet2.Activate();
                                        }
                                    }
                                    else
                                    {
                                        for (int m = 1, loopTo12 = TotalColumns; m <= loopTo12; m++)
                                            rng2.Cells[i, m].ClearFormats();
                                    }
                                }
                                excelApp.CutCopyMode = (XlCutCopyMode)Conversions.ToInteger(false);
                            }

                            else if (X1)
                            {

                                for (int i = 1, loopTo13 = r; i <= loopTo13; i++)
                                {
                                    string source = Conversions.ToString(rng.Cells[i, SplitColumn].value);
                                    var NumberText = new string[2];
                                    NumberText = (string[])SeparateNumberText(source);
                                    rng2.Cells[1, i].value = NumberText[0];
                                    rng2.Cells[2, i].value = NumberText[1];
                                    if (CheckBox1.Checked == true)
                                    {
                                        for (int m = 1, loopTo14 = TotalColumns; m <= loopTo14; m++)
                                        {
                                            rng.Cells[i, SplitColumn].copy();
                                            rng2.Cells[m, i].PasteSpecial(XlPasteType.xlPasteFormats);
                                            rng2 = workSheet2.get_Range(rng2Address);
                                            workSheet2.Activate();
                                        }
                                    }
                                    else
                                    {
                                        for (int m = 1, loopTo15 = TotalColumns; m <= loopTo15; m++)
                                            rng2.Cells[m, i].ClearFormats();
                                    }
                                }
                                excelApp.CutCopyMode = (XlCutCopyMode)Conversions.ToInteger(false);
                            }
                        }

                        else if (X11)
                        {

                            int W;

                            if (string.IsNullOrEmpty(TextBox3.Text))
                            {
                                W = 1;
                            }
                            else
                            {
                                W = Conversions.ToInteger(Conversion.Int(TextBox3.Text));
                            }

                            if (X2)
                            {

                                int Index;

                                for (int i = 1, loopTo16 = r; i <= loopTo16; i++)
                                {
                                    string source = Conversions.ToString(rng.Cells[i, SplitColumn].value);
                                    Index = 0;
                                    for (double k = 1d, loopTo17 = Conversion.Int(Strings.Len(source) / (double)W); k <= loopTo17; k++)
                                    {
                                        Index = Index + 1;
                                        rng.Cells[i, Index].value = Strings.Mid(source, (int)Math.Round(W * (k - 1d) + 1d), W);
                                    }
                                    if (Strings.Len(source) % W != 0)
                                    {
                                        Index = Index + 1;
                                        rng.Cells[i, Index].value = Strings.Mid(source, Strings.Len(source) - Strings.Len(source) % W + 1, Strings.Len(source) % W);
                                    }
                                    if (CheckBox1.Checked == true)
                                    {
                                        for (int m = 1, loopTo18 = TotalColumns; m <= loopTo18; m++)
                                        {
                                            rng.Cells[i, SplitColumn].copy();
                                            rng2.Cells[i, m].PasteSpecial(XlPasteType.xlPasteFormats);
                                            rng2 = workSheet2.get_Range(rng2Address);
                                            workSheet2.Activate();
                                        }
                                    }
                                    else
                                    {
                                        for (int m = 1, loopTo19 = TotalColumns; m <= loopTo19; m++)
                                            rng2.Cells[i, m].ClearFormats();
                                    }
                                }
                                excelApp.CutCopyMode = (XlCutCopyMode)Conversions.ToInteger(false);
                            }

                            else if (X1)
                            {

                                int Index;

                                for (int i = 1, loopTo20 = r; i <= loopTo20; i++)
                                {
                                    string source = Conversions.ToString(rng.Cells[i, SplitColumn].value);
                                    Index = 0;
                                    for (double k = 1d, loopTo21 = Conversion.Int(Strings.Len(source) / (double)W); k <= loopTo21; k++)
                                    {
                                        Index = Index + 1;
                                        rng.Cells[Index, i].value = Strings.Mid(source, (int)Math.Round(W * (k - 1d) + 1d), W);
                                    }
                                    if (Strings.Len(source) % W != 0)
                                    {
                                        Index = Index + 1;
                                        rng.Cells[Index, i].value = Strings.Mid(source, Strings.Len(source) - Strings.Len(source) % W + 1, Strings.Len(source) % W);
                                    }
                                    if (CheckBox1.Checked == true)
                                    {
                                        for (int m = 1, loopTo22 = TotalColumns; m <= loopTo22; m++)
                                        {
                                            rng.Cells[i, SplitColumn].copy();
                                            rng2.Cells[m, i].PasteSpecial(XlPasteType.xlPasteFormats);
                                            rng2 = workSheet2.get_Range(rng2Address);
                                            workSheet2.Activate();
                                        }
                                    }
                                    else
                                    {
                                        for (int m = 1, loopTo23 = TotalColumns; m <= loopTo23; m++)
                                            rng2.Cells[m, i].ClearFormats();
                                    }
                                }
                                excelApp.CutCopyMode = (XlCutCopyMode)Conversions.ToInteger(false);
                            }

                        }

                    }

                    rng2.Select();
                    for (int j = 1, loopTo24 = rng2.Columns.Count; j <= loopTo24; j++)
                        rng2.Columns[j].Autofit();

                    Close();
                }

                else
                {
                    if ((X1 | X2) & (X3 | X7 | X8 | X9 | X10 | X11))
                    {

                        int SplitColumn = 1;

                        var Arr = new object[rng.Rows.Count, rng.Columns.Count];

                        for (int i = Information.LBound(Arr, 1), loopTo25 = Information.UBound(Arr, 1); i <= loopTo25; i++)
                        {
                            for (int j = Information.LBound(Arr, 2), loopTo26 = Information.UBound(Arr, 2); j <= loopTo26; j++)
                                Arr[i, j] = rng.Cells[i + 1, j + 1].Value;
                        }

                        var FontNames = new string[rng.Rows.Count, rng.Columns.Count];
                        var FontSizes = new float[rng.Rows.Count, rng.Columns.Count];
                        var FontBolds = new bool[rng.Rows.Count, rng.Columns.Count];
                        var Fontitalics = new bool[rng.Rows.Count, rng.Columns.Count];
                        var Red1s = new int[rng.Rows.Count, rng.Columns.Count];
                        var Green1s = new int[rng.Rows.Count, rng.Columns.Count];
                        var Blue1s = new int[rng.Rows.Count, rng.Columns.Count];
                        var Red2s = new int[rng.Rows.Count, rng.Columns.Count];
                        var Green2s = new int[rng.Rows.Count, rng.Columns.Count];
                        var Blue2s = new int[rng.Rows.Count, rng.Columns.Count];

                        for (int i = Information.LBound(FontSizes, 1), loopTo27 = Information.UBound(FontSizes, 1); i <= loopTo27; i++)
                        {
                            for (int j = Information.LBound(FontSizes, 2), loopTo28 = Information.UBound(FontSizes, 2); j <= loopTo28; j++)
                            {

                                Range cell = (Range)rng.Cells[i + 1, j + 1];

                                var font = cell.Font;
                                FontNames[i, j] = Conversions.ToString(font.Name);
                                FontBolds[i, j] = Conversions.ToBoolean(cell.Font.Bold);
                                Fontitalics[i, j] = Conversions.ToBoolean(cell.Font.Italic);


                                float fontSize = Convert.ToSingle(font.Size);
                                FontSizes[i, j] = fontSize;

                                long colorValue1 = Conversions.ToLong(cell.Interior.Color);
                                int red1 = (int)(colorValue1 % 256L);
                                int green1 = (int)(colorValue1 / 256L % 256L);
                                int blue1 = (int)(colorValue1 / 256L / 256L % 256L);
                                Red1s[i, j] = red1;
                                Green1s[i, j] = green1;
                                Blue1s[i, j] = blue1;
                                long colorValue2 = Conversions.ToLong(cell.Font.Color);
                                int red2 = (int)(colorValue2 % 256L);
                                int green2 = (int)(colorValue2 / 256L % 256L);
                                int blue2 = (int)(colorValue2 / 256L / 256L % 256L);
                                Red2s[i, j] = red2;
                                Green2s[i, j] = green2;
                                Blue2s[i, j] = blue2;

                            }
                        }

                        if (X7 | X8 | X9 | X10)
                        {

                            string Separator = "";
                            if (X7)
                            {
                                Separator = ";";
                            }
                            else if (X8)
                            {
                                Separator = Microsoft.VisualBasic.Constants.vbNewLine;
                            }
                            else if (X9)
                            {
                                Separator = " ";
                            }
                            else if (X10)
                            {
                                Separator = ComboBox2.Text;
                            }

                            if (X2)
                            {

                                int Index;
                                int position;

                                for (int i = 1, loopTo29 = r; i <= loopTo29; i++)
                                {
                                    string source = Conversions.ToString(rng.Cells[i, SplitColumn].value);
                                    position = 1;
                                    Index = 0;
                                    for (int k = 1, loopTo30 = Strings.Len(source); k <= loopTo30; k++)
                                    {
                                        if ((Strings.Mid(source, k, Strings.Len(Separator)) ?? "") == (Separator ?? ""))
                                        {
                                            if (k - position > 0)
                                            {
                                                Index = Index + 1;
                                                rng2.Cells[i, Index].value = Strings.Mid(source, position, k - position);
                                            }
                                            position = k + Strings.Len(Separator);
                                        }
                                    }
                                    if (position <= Strings.Len(source))
                                    {
                                        Index = Index + 1;
                                        rng2.Cells[i, Index].value = Strings.Mid(source, position, Strings.Len(source) - position + 1);
                                    }
                                    if (CheckBox1.Checked == true)
                                    {

                                        int x = i - 1;
                                        int y = SplitColumn - 1;

                                        workSheet2.get_Range(rng2.Cells[i, 1], rng2.Cells[i, TotalColumns]).Font.Name = FontNames[x, y];
                                        workSheet2.get_Range(rng2.Cells[i, 1], rng2.Cells[i, TotalColumns]).Font.Size = (object)FontSizes[x, y];

                                        if (FontBolds[x, y])
                                            workSheet2.get_Range(rng2.Cells[i, 1], rng2.Cells[i, TotalColumns]).Font.Bold = (object)true;
                                        if (Fontitalics[x, y])
                                            workSheet2.get_Range(rng2.Cells[i, 1], rng2.Cells[i, TotalColumns]).Font.Italic = (object)true;

                                        workSheet2.get_Range(rng2.Cells[i, 1], rng2.Cells[i, TotalColumns]).Interior.Color = (object)Color.FromArgb(Red1s[x, y], Green1s[x, y], Blue1s[x, y]);

                                        workSheet2.get_Range(rng2.Cells[i, 1], rng2.Cells[i, TotalColumns]).Font.Color = (object)Color.FromArgb(Red2s[x, y], Green2s[x, y], Blue2s[x, y]);
                                        workSheet2.Activate();
                                    }

                                    else
                                    {
                                        workSheet2.get_Range(rng2.Cells[i, 1], rng2.Cells[i, TotalColumns]).ClearFormats();

                                    }
                                }
                            }

                            else if (X1)
                            {

                                int Index;
                                int position;

                                for (int i = 1, loopTo31 = r; i <= loopTo31; i++)
                                {
                                    string source = Conversions.ToString(rng.Cells[i, SplitColumn].value);
                                    position = 1;
                                    Index = 0;
                                    for (int k = 1, loopTo32 = Strings.Len(source); k <= loopTo32; k++)
                                    {
                                        if ((Strings.Mid(source, k, Strings.Len(Separator)) ?? "") == (Separator ?? ""))
                                        {
                                            if (k - position > 0)
                                            {
                                                Index = Index + 1;
                                                rng2.Cells[Index, i].value = Strings.Mid(source, position, k - position);
                                            }
                                            position = k + Strings.Len(Separator);
                                        }
                                    }
                                    if (position <= Strings.Len(source))
                                    {
                                        Index = Index + 1;
                                        rng2.Cells[Index, i].value = Strings.Mid(source, position, Strings.Len(source) - position + 1);
                                    }
                                    if (CheckBox1.Checked == true)
                                    {

                                        int x = i - 1;
                                        int y = SplitColumn - 1;

                                        workSheet2.get_Range(rng2.Cells[1, i], rng2.Cells[TotalColumns, i]).Font.Name = FontNames[x, y];
                                        workSheet2.get_Range(rng2.Cells[1, i], rng2.Cells[TotalColumns, i]).Font.Size = (object)FontSizes[x, y];

                                        if (FontBolds[x, y])
                                            workSheet2.get_Range(rng2.Cells[1, i], rng2.Cells[TotalColumns, i]).Font.Bold = (object)true;
                                        if (Fontitalics[x, y])
                                            workSheet2.get_Range(rng2.Cells[1, i], rng2.Cells[TotalColumns, i]).Font.Italic = (object)true;

                                        workSheet2.get_Range(rng2.Cells[1, i], rng2.Cells[TotalColumns, i]).Interior.Color = (object)Color.FromArgb(Red1s[x, y], Green1s[x, y], Blue1s[x, y]);

                                        workSheet2.get_Range(rng2.Cells[1, i], rng2.Cells[TotalColumns, i]).Font.Color = (object)Color.FromArgb(Red2s[x, y], Green2s[x, y], Blue2s[x, y]);
                                        workSheet2.Activate();
                                    }
                                    else
                                    {
                                        workSheet2.get_Range(rng2.Cells[1, i], rng2.Cells[TotalColumns, i]).ClearFormats();

                                    }
                                }
                            }
                        }

                        else if (X3)
                        {

                            if (X2)
                            {

                                for (int i = 1, loopTo33 = r; i <= loopTo33; i++)
                                {
                                    string source = Conversions.ToString(rng.Cells[i, SplitColumn].value);
                                    var NumberText = new string[2];
                                    NumberText = (string[])SeparateNumberText(source);
                                    rng2.Cells[i, 1].value = NumberText[0];
                                    rng2.Cells[i, 2].value = NumberText[1];
                                    if (CheckBox1.Checked == true)
                                    {
                                        int x = i - 1;
                                        int y = SplitColumn - 1;

                                        workSheet2.get_Range(rng2.Cells[i, 1], rng2.Cells[i, 2]).Font.Name = FontNames[x, y];
                                        workSheet2.get_Range(rng2.Cells[i, 1], rng2.Cells[i, 2]).Font.Size = (object)FontSizes[x, y];

                                        if (FontBolds[x, y])
                                            workSheet2.get_Range(rng2.Cells[i, 1], rng2.Cells[i, 2]).Font.Bold = (object)true;
                                        if (Fontitalics[x, y])
                                            workSheet2.get_Range(rng2.Cells[i, 1], rng2.Cells[i, 2]).Font.Italic = (object)true;

                                        workSheet2.get_Range(rng2.Cells[i, 1], rng2.Cells[i, 2]).Interior.Color = (object)Color.FromArgb(Red1s[x, y], Green1s[x, y], Blue1s[x, y]);

                                        workSheet2.get_Range(rng2.Cells[i, 1], rng2.Cells[i, 2]).Font.Color = (object)Color.FromArgb(Red2s[x, y], Green2s[x, y], Blue2s[x, y]);
                                        workSheet2.Activate();
                                    }
                                    else
                                    {
                                        workSheet2.get_Range(rng2.Cells[i, 1], rng2.Cells[i, 2]).ClearFormats();

                                    }
                                }
                            }

                            else if (X1)
                            {

                                for (int i = 1, loopTo34 = r; i <= loopTo34; i++)
                                {
                                    string source = Conversions.ToString(rng.Cells[i, SplitColumn].value);
                                    var NumberText = new string[2];
                                    NumberText = (string[])SeparateNumberText(source);
                                    rng2.Cells[1, i].value = NumberText[0];
                                    rng2.Cells[2, i].value = NumberText[1];
                                    if (CheckBox1.Checked == true)
                                    {
                                        int x = i - 1;
                                        int y = SplitColumn - 1;

                                        workSheet2.get_Range(rng2.Cells[1, i], rng2.Cells[2, i]).Font.Name = FontNames[x, y];
                                        workSheet2.get_Range(rng2.Cells[1, i], rng2.Cells[2, i]).Font.Size = (object)FontSizes[x, y];

                                        if (FontBolds[x, y])
                                            workSheet2.get_Range(rng2.Cells[1, i], rng2.Cells[2, i]).Font.Bold = (object)true;
                                        if (Fontitalics[x, y])
                                            workSheet2.get_Range(rng2.Cells[1, i], rng2.Cells[2, i]).Font.Italic = (object)true;

                                        workSheet2.get_Range(rng2.Cells[1, i], rng2.Cells[2, i]).Interior.Color = (object)Color.FromArgb(Red1s[x, y], Green1s[x, y], Blue1s[x, y]);

                                        workSheet2.get_Range(rng2.Cells[1, i], rng2.Cells[2, i]).Font.Color = (object)Color.FromArgb(Red2s[x, y], Green2s[x, y], Blue2s[x, y]);
                                        workSheet2.Activate();
                                    }
                                    else
                                    {
                                        workSheet2.get_Range(rng2.Cells[1, i], rng2.Cells[2, i]).ClearFormats();
                                    }
                                }
                            }
                        }

                        else if (X11)
                        {

                            int W;

                            if (string.IsNullOrEmpty(TextBox3.Text))
                            {
                                W = 1;
                            }
                            else
                            {
                                W = Conversions.ToInteger(Conversion.Int(TextBox3.Text));
                            }

                            if (X2)
                            {

                                int Index;

                                for (int i = 1, loopTo35 = r; i <= loopTo35; i++)
                                {
                                    string source = Conversions.ToString(rng.Cells[i, SplitColumn].value);
                                    Index = 0;
                                    for (double k = 1d, loopTo36 = Conversion.Int(Strings.Len(source) / (double)W); k <= loopTo36; k++)
                                    {
                                        Index = Index + 1;
                                        rng.Cells[i, Index].value = Strings.Mid(source, (int)Math.Round(W * (k - 1d) + 1d), W);
                                    }
                                    if (Strings.Len(source) % W != 0)
                                    {
                                        Index = Index + 1;
                                        rng.Cells[i, Index].value = Strings.Mid(source, Strings.Len(source) - Strings.Len(source) % W + 1, Strings.Len(source) % W);
                                    }
                                    if (CheckBox1.Checked == true)
                                    {
                                        int x = i - 1;
                                        int y = SplitColumn - 1;

                                        workSheet2.get_Range(rng2.Cells[i, 1], rng2.Cells[i, TotalColumns]).Font.Name = FontNames[x, y];
                                        workSheet2.get_Range(rng2.Cells[i, 1], rng2.Cells[i, TotalColumns]).Font.Size = (object)FontSizes[x, y];

                                        if (FontBolds[x, y])
                                            workSheet2.get_Range(rng2.Cells[i, 1], rng2.Cells[i, TotalColumns]).Font.Bold = (object)true;
                                        if (Fontitalics[x, y])
                                            workSheet2.get_Range(rng2.Cells[i, 1], rng2.Cells[i, TotalColumns]).Font.Italic = (object)true;

                                        workSheet2.get_Range(rng2.Cells[i, 1], rng2.Cells[i, TotalColumns]).Interior.Color = (object)Color.FromArgb(Red1s[x, y], Green1s[x, y], Blue1s[x, y]);

                                        workSheet2.get_Range(rng2.Cells[i, 1], rng2.Cells[i, TotalColumns]).Font.Color = (object)Color.FromArgb(Red2s[x, y], Green2s[x, y], Blue2s[x, y]);
                                        workSheet2.Activate();
                                    }
                                    else
                                    {
                                        workSheet2.get_Range(rng2.Cells[i, 1], rng2.Cells[i, TotalColumns]).ClearFormats();
                                    }
                                }
                            }

                            else if (X1)
                            {

                                int Index;

                                for (int i = 1, loopTo37 = r; i <= loopTo37; i++)
                                {
                                    string source = Conversions.ToString(rng.Cells[i, SplitColumn].value);
                                    Index = 0;
                                    for (double k = 1d, loopTo38 = Conversion.Int(Strings.Len(source) / (double)W); k <= loopTo38; k++)
                                    {
                                        Index = Index + 1;
                                        rng.Cells[Index, i].value = Strings.Mid(source, (int)Math.Round(W * (k - 1d) + 1d), W);
                                    }
                                    if (Strings.Len(source) % W != 0)
                                    {
                                        Index = Index + 1;
                                        rng.Cells[Index, i].value = Strings.Mid(source, Strings.Len(source) - Strings.Len(source) % W + 1, Strings.Len(source) % W);
                                    }
                                    if (CheckBox1.Checked == true)
                                    {
                                        int x = i - 1;
                                        int y = SplitColumn - 1;

                                        workSheet2.get_Range(rng2.Cells[1, i], rng2.Cells[TotalColumns, i]).Font.Name = FontNames[x, y];
                                        workSheet2.get_Range(rng2.Cells[1, i], rng2.Cells[TotalColumns, i]).Font.Size = (object)FontSizes[x, y];

                                        if (FontBolds[x, y])
                                            workSheet2.get_Range(rng2.Cells[1, i], rng2.Cells[TotalColumns, i]).Font.Bold = (object)true;
                                        if (Fontitalics[x, y])
                                            workSheet2.get_Range(rng2.Cells[1, i], rng2.Cells[TotalColumns, i]).Font.Italic = (object)true;

                                        workSheet2.get_Range(rng2.Cells[1, i], rng2.Cells[TotalColumns, i]).Interior.Color = (object)Color.FromArgb(Red1s[x, y], Green1s[x, y], Blue1s[x, y]);

                                        workSheet2.get_Range(rng2.Cells[1, i], rng2.Cells[TotalColumns, i]).Font.Color = (object)Color.FromArgb(Red2s[x, y], Green2s[x, y], Blue2s[x, y]);
                                        workSheet2.Activate();
                                    }
                                    else
                                    {
                                        workSheet2.get_Range(rng2.Cells[1, i], rng2.Cells[TotalColumns, i]).ClearFormats();
                                    }
                                }
                            }

                        }

                    }

                    rng2.Select();

                    for (int j = 1, loopTo39 = rng2.Columns.Count; j <= loopTo39; j++)
                        rng2.Columns[j].Autofit();

                    Close();

                }

                TextBoxChanged = false;
            }

            catch (Exception ex)
            {

            }

        }

        private void RadioButton1_CheckedChanged(object sender, EventArgs e)
        {

            try
            {
                if (RadioButton1.Checked == true)
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

        private void TextBox1_TextChanged(object sender, EventArgs e)
        {

            try
            {

                excelApp = Globals.ThisAddIn.Application;
                workBook = excelApp.ActiveWorkbook;
                workSheet = (Excel.Worksheet)workBook.ActiveSheet;

                string[] rngArray = Strings.Split(TextBox1.Text, "!");
                string rngAddress = rngArray[Information.UBound(rngArray)];
                rng = workSheet.get_Range(rngAddress);
                TextBoxChanged = true;
                rng.Select();
                Display();
                TextBoxChanged = false;
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

                var activeRange = excelApp.ActiveCell;

                int startRow = activeRange.Row;
                int startColumn = activeRange.Column;
                int endRow = activeRange.Row;
                int endColumn = activeRange.Column;

                // Find the upper boundary
                while (startRow > 1 && !(workSheet.Cells[startRow - 1, startColumn].Value == null))
                    startRow -= 1;

                // Find the lower boundary
                while (!(workSheet.Cells[endRow + 1, endColumn].Value == null))
                    endRow += 1;

                // Select the determined range
                rng = workSheet.get_Range(workSheet.Cells[startRow, startColumn], workSheet.Cells[endRow, endColumn]);

                rng.Select();

                string sheetName;

                sheetName = Strings.Split(rng.get_Address(true, true, XlReferenceStyle.xlA1, true), "]")[1];
                sheetName = Strings.Split(sheetName, "!")[0];

                if (Strings.Mid(sheetName, Strings.Len(sheetName), 1) == "'")
                {
                    sheetName = Strings.Mid(sheetName, 1, Strings.Len(sheetName) - 1);
                }

                workSheet = (Excel.Worksheet)workBook.Worksheets[sheetName];
                workSheet.Activate();

                if ((workSheet.Name ?? "") != (OpenSheet.Name ?? ""))
                {
                    TextBox1.Text = workSheet.Name + "!" + rng.get_Address();
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

        private void Selection_Click(object sender, EventArgs e)
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

        private void TextBox4_TextChanged(object sender, EventArgs e)
        {

            try
            {
                excelApp = Globals.ThisAddIn.Application;
                workBook = excelApp.ActiveWorkbook;
                workSheet2 = (Excel.Worksheet)workBook.ActiveSheet;

                string[] rng2Array = Strings.Split(TextBox4.Text, "!");
                string rng2Address = rng2Array[Information.UBound(rng2Array)];
                rng2 = workSheet2.get_Range(rng2Address);

                TextBoxChanged = true;

                rng2.Select();

                TextBoxChanged = false;
            }

            catch (Exception ex)
            {

            }

        }

        private void PictureBox3_Click(object sender, EventArgs e)
        {

            try
            {
                FocusedTextBox = 4;
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

                workSheet2 = (Excel.Worksheet)workBook.Worksheets[sheetName];
                workSheet2.Activate();

                rng2.Select();

                if ((workSheet2.Name ?? "") != (OpenSheet.Name ?? ""))
                {
                    TextBox4.Text = workSheet2.Name + "!" + rng2.get_Address();
                }
                else
                {
                    TextBox4.Text = rng2.get_Address();
                }

                Show();
                TextBox4.Focus();
            }

            catch (Exception ex)
            {

                Show();
                TextBox4.Focus();

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

        private void ComboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                Display();
            }
            catch (Exception ex)
            {

            }
        }

        private void TextBox3_TextChanged(object sender, EventArgs e)
        {
            try
            {
                if (Information.IsNumeric(TextBox3.Text) | string.IsNullOrEmpty(TextBox3.Text))
                {
                    Display();
                }
                else
                {
                    MessageBox.Show("Enter a Numerical Value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {

            }
        }

        private void RadioButton9_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                if (RadioButton9.Checked)
                {
                    Display();
                }
            }
            catch (Exception ex)
            {

            }
        }

        private void RadioButton8_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                if (RadioButton8.Checked)
                {
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
                    Display();
                }
            }
            catch (Exception ex)
            {

            }
        }

        private void RadioButton7_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                if (RadioButton7.Checked)
                {
                    Display();
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
                if (RadioButton10.Checked)
                {
                    ComboBox2.Enabled = true;
                    ComboBox2.Focus();
                    Display();
                }
                else
                {
                    ComboBox2.Text = "";
                    ComboBox2.Enabled = false;
                }
            }

            catch (Exception ex)
            {

            }
        }

        private void RadioButton11_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                if (RadioButton11.Checked)
                {
                    PictureBox11.Enabled = true;
                    TextBox3.Enabled = true;
                    TextBox3.Focus();
                    Display();
                }
                else
                {
                    TextBox3.Clear();
                    PictureBox11.Enabled = false;
                    TextBox3.Enabled = false;
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
                    workSheet2 = workSheet;
                    rng2 = rng;
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
                    Label3.Enabled = true;
                    PictureBox2.Enabled = true;
                    PictureBox3.Enabled = true;
                    TextBox4.Enabled = true;
                    TextBox4.Focus();
                }
                else
                {
                    TextBox4.Clear();
                    Label3.Enabled = false;
                    PictureBox2.Enabled = false;
                    PictureBox3.Enabled = false;
                    TextBox4.Enabled = false;
                }
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

        private void Form24_Split_Cells_Load(object sender, EventArgs e)
        {
            try
            {

                excelApp = Globals.ThisAddIn.Application;
                workBook = excelApp.ActiveWorkbook;
                workSheet = (Excel.Worksheet)workBook.ActiveSheet;
                workSheet2 = (Excel.Worksheet)workBook.ActiveSheet;
                KeyPreview = true;

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

                Range selectedRange;
                selectedRange = (Range)excelApp.Selection;

                if (TextBoxChanged == false)
                {
                    if (FocusedTextBox == 1)
                    {
                        workSheet = (Excel.Worksheet)workBook.ActiveSheet;
                        if ((workSheet.Name ?? "") != (OpenSheet.Name ?? ""))
                        {
                            TextBox1.Text = workSheet.Name + "!" + selectedRange.get_Address();
                        }
                        else
                        {
                            TextBox1.Text = selectedRange.get_Address();
                        }
                        rng = selectedRange;
                        TextBox1.Focus();
                    }

                    else if (FocusedTextBox == 4)
                    {
                        workSheet2 = (Excel.Worksheet)workBook.ActiveSheet;
                        if ((workSheet2.Name ?? "") != (OpenSheet.Name ?? ""))
                        {
                            TextBox4.Text = workSheet2.Name + "!" + selectedRange.get_Address();
                        }
                        else
                        {
                            TextBox4.Text = selectedRange.get_Address();
                        }
                        rng2 = selectedRange;
                        TextBox4.Focus();
                    }
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
                FocusedTextBox = 1;
            }
            catch (Exception ex)
            {

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

        private void TextBox4_GotFocus(object sender, EventArgs e)
        {
            try
            {
                FocusedTextBox = 4;
            }
            catch (Exception ex)
            {

            }
        }

        private void PictureBox3_GotFocus(object sender, EventArgs e)
        {
            try
            {
                FocusedTextBox = 4;
            }
            catch (Exception ex)
            {

            }
        }

        private void Form24_Split_Cells_KeyDown(object sender, KeyEventArgs e)
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

        private void Form24_Split_Cells_Closing(object sender, CancelEventArgs e)
        {
            GlobalModule.form_flag = false;
        }

        private void Form24_Split_Cells_Disposed(object sender, EventArgs e)
        {
            GlobalModule.form_flag = false;
        }

        private void Form24_Split_Cells_Shown(object sender, EventArgs e)
        {

            try
            {

                Focus();
                BringToFront();
                Activate();

                string TextBoxText;

                if ((workSheet.Name ?? "") != (OpenSheet.Name ?? ""))
                {
                    TextBoxText = workSheet.Name + "!" + rng.get_Address();
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

    }
}