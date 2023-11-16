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

    public partial class Form26_split_text_bycharacters
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
        private Range rng2;
        private Range selectedRange;

        private int opened;
        private int FocusedTextBox;
        private bool TextBoxChanged;

        public Form26_split_text_bycharacters()
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

        private object SplitText(string Source, string Separator, bool Consecutive, bool KeepSeparator, bool Before)
        {
            object SplitTextRet = default;

            int position = 1;
            int Index = -1;
            var Splitvalues = new string[1];

            for (int k = 1, loopTo = Strings.Len(Source); k <= loopTo; k++)
            {
                if ((Strings.Mid(Source, k, Strings.Len(Separator)) ?? "") == (Separator ?? ""))
                {
                    bool Condition;
                    if (Consecutive == true)
                    {
                        Condition = k - position > 0;
                    }
                    else
                    {
                        Condition = true;
                    }
                    if (k != 1 & Condition)
                    {
                        Index = Index + 1;
                        Array.Resize(ref Splitvalues, Index + 1);
                        string Value;
                        if (KeepSeparator == true)
                        {
                            if (Before == true)
                            {
                                Value = Separator + Strings.Mid(Source, position, k - position);
                            }
                            else
                            {
                                Value = Strings.Mid(Source, position, k - position) + Separator;
                            }
                        }
                        else
                        {
                            Value = Strings.Mid(Source, position, k - position);
                        }
                        Splitvalues[Index] = Value;
                    }
                    position = k + Strings.Len(Separator);
                }
            }

            if (position <= Strings.Len(Source))
            {
                Index = Index + 1;
                Array.Resize(ref Splitvalues, Index + 1);
                string Value;
                if (KeepSeparator == true)
                {
                    if (Before == true)
                    {
                        Value = Separator + Strings.Mid(Source, position, Strings.Len(Source) - position + 1);
                    }
                    else
                    {
                        Value = Strings.Mid(Source, position, Strings.Len(Source) - position + 1);
                    }
                }
                else
                {
                    Value = Strings.Mid(Source, position, Strings.Len(Source) - position + 1);
                }
                Splitvalues[Index] = Value;
            }

            SplitTextRet = Splitvalues;
            return SplitTextRet;

        }
        private object SplitCount(string Source, string Separator, bool Consecutive)
        {
            object SplitCountRet = default;

            int position = 1;
            int Index = 0;

            for (int k = 1, loopTo = Strings.Len(Source); k <= loopTo; k++)
            {
                if ((Strings.Mid(Source, k, Strings.Len(Separator)) ?? "") == (Separator ?? ""))
                {
                    bool Condition;
                    if (Consecutive == true)
                    {
                        Condition = k - position > 0;
                    }
                    else
                    {
                        Condition = true;
                    }
                    if (k != 1 & Condition)
                    {
                        Index = Index + 1;
                    }
                    position = k + Strings.Len(Separator);
                }
            }

            if (position <= Strings.Len(Source))
            {
                Index = Index + 1;
            }

            SplitCountRet = Index;
            return SplitCountRet;

        }
        private object SplitByWidth(object source, object W)
        {
            object SplitByWidthRet = default;

            int Index = -1;
            var SplitValues = new string[1];
            for (int k = 1, loopTo = Conversions.ToInteger(Conversion.Int(Operators.DivideObject(Strings.Len(source), W))); k <= loopTo; k++)
            {
                Index = Index + 1;
                Array.Resize(ref SplitValues, Index + 1);
                SplitValues[Index] = Strings.Mid(Conversions.ToString(source), Conversions.ToInteger(Operators.AddObject(Operators.MultiplyObject(W, Operators.SubtractObject(k, 1)), 1)), Conversions.ToInteger(W));
            }
            if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(Operators.ModObject(Strings.Len(source), W), 0, false)))
            {
                Index = Index + 1;
                Array.Resize(ref SplitValues, Index + 1);
                SplitValues[Index] = Strings.Mid(Conversions.ToString(source), Conversions.ToInteger(Operators.AddObject(Operators.SubtractObject(Strings.Len(source), Operators.ModObject(Strings.Len(source), W)), 1)), Conversions.ToInteger(Operators.ModObject(Strings.Len(source), W)));
            }

            SplitByWidthRet = SplitValues;
            return SplitByWidthRet;

        }

        private void Display()
        {

            Panel_InputRange.Controls.Clear();
            Panel_ExpectedOutput.Controls.Clear();

            excelApp = Globals.ThisAddIn.Application;
            workBook = excelApp.ActiveWorkbook;
            workSheet = (Excel.Worksheet)workBook.ActiveSheet;

            rng = workSheet.get_Range(TB_source_range.Text);

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
                Height = Panel_InputRange.Height / (double)displayRng.Rows.Count;
            }
            else
            {
                Height = 119d / 4d;
            }

            BaseWidth = 260d / 3d;
            Width = Conversions.ToDouble(Operators.DivideObject(Operators.MultiplyObject(MaxOfColumn(displayRng), BaseWidth), 10));

            double Width1;

            if (Width > Panel_InputRange.Width)
            {
                Width1 = Width;
            }
            else
            {
                Width1 = Panel_InputRange.Width;
            }
            double ordinate = 0d;

            for (int i = 1, loopTo = r; i <= loopTo; i++)
            {
                var label = new System.Windows.Forms.Label();
                label.Text = Conversions.ToString(displayRng.Cells[i, 1].Value);
                label.Location = new System.Drawing.Point((int)Math.Round(ordinate), (int)Math.Round((i - 1) * Height));
                label.Height = (int)Math.Round(Height);
                label.Width = (int)Math.Round(Width1);
                label.BorderStyle = BorderStyle.FixedSingle;
                label.TextAlign = ContentAlignment.MiddleCenter;

                if (CB_formatting.Checked == true)
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
                Panel_InputRange.Controls.Add(label);
            }
            Panel_InputRange.AutoScroll = true;

            bool X1 = RB_rows.Checked;
            bool X2 = RB_columns.Checked;
            bool X3 = RB_space.Checked;
            bool X7 = RB_newline.Checked;
            bool X8 = RB_numbertext.Checked;
            bool X9 = RB_semicolon.Checked;
            bool X10 = RB_others.Checked;
            bool X11 = RB_width.Checked;


            if ((X1 | X2) & (X3 | X7 | X8 | X9 | X10 | X11))
            {

                if (X3 | X7 | X9 | X10)
                {

                    string Separator = ",";
                    if (X7)
                    {
                        Separator = Microsoft.VisualBasic.Constants.vbNewLine;
                    }
                    else if (X3)
                    {
                        Separator = " ";
                    }
                    else if (X9)
                    {
                        Separator = ";";
                    }
                    else if (X10)
                    {
                        Separator = ComboBox2.Text;
                    }

                    bool Consecutive;
                    if (CB_consecute_separators.Checked)
                    {
                        Consecutive = true;
                    }
                    else
                    {
                        Consecutive = false;
                    }

                    bool KeepSeparator;
                    if (CB_separators_finaloutput.Checked)
                    {
                        KeepSeparator = true;
                    }
                    else
                    {
                        KeepSeparator = false;
                    }

                    bool Before;
                    if (RB_starting_point.Checked)
                    {
                        Before = true;
                    }
                    else
                    {
                        Before = false;
                    }

                    if (X1)
                    {
                        var values = new string[1];
                        int Index = -1;
                        for (int i = 1, loopTo1 = r; i <= loopTo1; i++)
                        {
                            string source = Conversions.ToString(displayRng.Cells[i, 1].value);
                            string[] SplitValues;
                            SplitValues = (string[])SplitText(source, Separator, Consecutive, KeepSeparator, Before);
                            for (int m = Information.LBound(SplitValues), loopTo2 = Information.UBound(SplitValues); m <= loopTo2; m++)
                            {
                                Index = Index + 1;
                                Array.Resize(ref values, Index + 1);
                                values[Index] = SplitValues[m];
                            }
                        }

                        double Width2 = Conversions.ToDouble(Operators.DivideObject(Operators.MultiplyObject(MaxOfArray(values), BaseWidth), 10));
                        if (Width + Width2 < Panel_ExpectedOutput.Width)
                        {
                            Width2 = Panel_ExpectedOutput.Width - Width;
                        }
                        double abscissa1 = 0d;
                        double abscissa2 = 0d;
                        for (int i = 1, loopTo3 = r; i <= loopTo3; i++)
                        {
                            string source = Conversions.ToString(displayRng.Cells[i, 1].value);
                            string[] SplitValues;
                            SplitValues = (string[])SplitText(source, Separator, Consecutive, KeepSeparator, Before);

                            var label = new System.Windows.Forms.Label();
                            label.Text = Conversions.ToString(displayRng.Cells[i, 1].Value);
                            label.Location = new System.Drawing.Point(0, (int)Math.Round(abscissa1));
                            label.Height = (int)Math.Round(Height);
                            label.Width = (int)Math.Round(Width);
                            label.BorderStyle = BorderStyle.FixedSingle;
                            label.TextAlign = ContentAlignment.MiddleCenter;

                            if (CB_formatting.Checked == true)
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
                            Panel_ExpectedOutput.Controls.Add(label);
                            abscissa1 = abscissa1 + Height;
                            for (int m = Information.LBound(SplitValues) + 1, loopTo4 = Information.UBound(SplitValues); m <= loopTo4; m++)
                            {
                                var label1 = new System.Windows.Forms.Label();
                                label1.Text = "";
                                label1.Location = new System.Drawing.Point(0, (int)Math.Round(abscissa1));
                                label1.Height = (int)Math.Round(Height);
                                label1.Width = (int)Math.Round(Width);
                                label1.BorderStyle = BorderStyle.FixedSingle;
                                label1.TextAlign = ContentAlignment.MiddleCenter;

                                if (CB_formatting.Checked == true)
                                {
                                    Range cell = (Range)displayRng.Cells[i, 1];
                                    var font = cell.Font;
                                    var fontStyle = FontStyle.Regular;
                                    if (Conversions.ToBoolean(cell.Font.Bold))
                                        fontStyle = fontStyle | FontStyle.Bold;
                                    if (Conversions.ToBoolean(cell.Font.Italic))
                                        fontStyle = fontStyle | FontStyle.Italic;

                                    float fontSize = Convert.ToSingle(font.Size);

                                    label1.Font = new System.Drawing.Font(font.ToString(), fontSize, fontStyle);
                                    if (Conversions.ToBoolean(!Operators.ConditionalCompareObjectEqual(cell.Interior.ColorIndex, XlColorIndex.xlColorIndexNone, false)))
                                    {
                                        long colorValue1 = Conversions.ToLong(cell.Interior.Color);
                                        int red1 = (int)(colorValue1 % 256L);
                                        int green1 = (int)(colorValue1 / 256L % 256L);
                                        int blue1 = (int)(colorValue1 / 256L / 256L % 256L);
                                        label1.BackColor = Color.FromArgb(red1, green1, blue1);
                                    }

                                    if (cell.Font.Color is DBNull)
                                    {
                                        label1.ForeColor = Color.FromArgb(0, 0, 0);
                                    }

                                    else if (Conversions.ToBoolean(!Operators.ConditionalCompareObjectEqual(cell.Font.ColorIndex, XlColorIndex.xlColorIndexNone, false)))
                                    {
                                        long colorValue2 = Conversions.ToLong(cell.Font.Color);
                                        int red2 = (int)(colorValue2 % 256L);
                                        int green2 = (int)(colorValue2 / 256L % 256L);
                                        int blue2 = (int)(colorValue2 / 256L / 256L % 256L);
                                        label1.ForeColor = Color.FromArgb(red2, green2, blue2);
                                    }
                                }
                                Panel_ExpectedOutput.Controls.Add(label1);
                                abscissa1 = abscissa1 + Height;
                            }

                            for (int m = Information.LBound(SplitValues), loopTo5 = Information.UBound(SplitValues); m <= loopTo5; m++)
                            {
                                var label1 = new System.Windows.Forms.Label();
                                label1.Text = SplitValues[m];
                                label1.Location = new System.Drawing.Point((int)Math.Round(Width), (int)Math.Round(abscissa2));
                                label1.Height = (int)Math.Round(Height);
                                label1.Width = (int)Math.Round(Width2);
                                label1.BorderStyle = BorderStyle.FixedSingle;
                                label1.TextAlign = ContentAlignment.MiddleCenter;

                                if (CB_formatting.Checked == true)
                                {
                                    Range cell = (Range)displayRng.Cells[i, 1];
                                    var font = cell.Font;
                                    var fontStyle = FontStyle.Regular;
                                    if (Conversions.ToBoolean(cell.Font.Bold))
                                        fontStyle = fontStyle | FontStyle.Bold;
                                    if (Conversions.ToBoolean(cell.Font.Italic))
                                        fontStyle = fontStyle | FontStyle.Italic;

                                    float fontSize = Convert.ToSingle(font.Size);

                                    label1.Font = new System.Drawing.Font(font.ToString(), fontSize, fontStyle);
                                    if (Conversions.ToBoolean(!Operators.ConditionalCompareObjectEqual(cell.Interior.ColorIndex, XlColorIndex.xlColorIndexNone, false)))
                                    {
                                        long colorValue1 = Conversions.ToLong(cell.Interior.Color);
                                        int red1 = (int)(colorValue1 % 256L);
                                        int green1 = (int)(colorValue1 / 256L % 256L);
                                        int blue1 = (int)(colorValue1 / 256L / 256L % 256L);
                                        label1.BackColor = Color.FromArgb(red1, green1, blue1);
                                    }

                                    if (cell.Font.Color is DBNull)
                                    {
                                        label1.ForeColor = Color.FromArgb(0, 0, 0);
                                    }

                                    else if (Conversions.ToBoolean(!Operators.ConditionalCompareObjectEqual(cell.Font.ColorIndex, XlColorIndex.xlColorIndexNone, false)))
                                    {
                                        long colorValue2 = Conversions.ToLong(cell.Font.Color);
                                        int red2 = (int)(colorValue2 % 256L);
                                        int green2 = (int)(colorValue2 / 256L % 256L);
                                        int blue2 = (int)(colorValue2 / 256L / 256L % 256L);
                                        label1.ForeColor = Color.FromArgb(red2, green2, blue2);
                                    }
                                }
                                Panel_ExpectedOutput.Controls.Add(label1);
                                abscissa2 = abscissa2 + Height;
                            }
                        }
                    }

                    else if (X2)
                    {
                        ordinate = 0d;

                        for (int i = 1, loopTo6 = displayRng.Rows.Count; i <= loopTo6; i++)
                        {
                            var label = new System.Windows.Forms.Label();
                            label.Text = Conversions.ToString(displayRng.Cells[i, 1].Value);
                            label.Location = new System.Drawing.Point((int)Math.Round(ordinate), (int)Math.Round((i - 1) * Height));
                            label.Height = (int)Math.Round(Height);
                            label.Width = (int)Math.Round(Width);
                            label.BorderStyle = BorderStyle.FixedSingle;
                            label.TextAlign = ContentAlignment.MiddleCenter;

                            if (CB_formatting.Checked == true)
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
                            Panel_ExpectedOutput.Controls.Add(label);
                        }
                        ordinate = ordinate + Width;
                        var lengths = new int[r];
                        for (int i = 1, loopTo7 = displayRng.Rows.Count; i <= loopTo7; i++)
                        {
                            string source = Conversions.ToString(displayRng.Cells[i, 1].value);
                            lengths[i - 1] = Conversions.ToInteger(SplitCount(source, Separator, Consecutive));
                        }
                        int TotalWidth = Conversions.ToInteger(FindMax(lengths));

                        var Values = new string[r, TotalWidth];

                        for (int i = 1, loopTo8 = displayRng.Rows.Count; i <= loopTo8; i++)
                        {
                            string source = Conversions.ToString(displayRng.Cells[i, 1].value);
                            string[] SplitValues;
                            SplitValues = (string[])SplitText(source, Separator, Consecutive, KeepSeparator, Before);
                            for (int j = Information.LBound(SplitValues), loopTo9 = Information.UBound(SplitValues); j <= loopTo9; j++)
                                Values[i - 1, j] = SplitValues[j];
                        }
                        for (int j = 0, loopTo10 = TotalWidth - 1; j <= loopTo10; j++)
                        {
                            var ColumnValues = new string[r];
                            for (int i = 0, loopTo11 = r - 1; i <= loopTo11; i++)
                                ColumnValues[i] = Values[i, j];
                            Width1 = Conversions.ToDouble(Operators.DivideObject(Operators.MultiplyObject(MaxOfArray(ColumnValues), BaseWidth), 10));
                            for (int i = 0, loopTo12 = r - 1; i <= loopTo12; i++)
                            {
                                var label1 = new System.Windows.Forms.Label();
                                label1.Text = ColumnValues[i];
                                label1.Location = new System.Drawing.Point((int)Math.Round(ordinate), (int)Math.Round(i * Height));
                                label1.Height = (int)Math.Round(Height);
                                label1.Width = (int)Math.Round(Width1);
                                label1.BorderStyle = BorderStyle.FixedSingle;
                                label1.TextAlign = ContentAlignment.MiddleCenter;

                                if (CB_formatting.Checked == true)
                                {
                                    Range cell = (Range)displayRng.Cells[i + 1, 1];
                                    var font = cell.Font;
                                    var fontStyle = FontStyle.Regular;
                                    if (Conversions.ToBoolean(cell.Font.Bold))
                                        fontStyle = fontStyle | FontStyle.Bold;
                                    if (Conversions.ToBoolean(cell.Font.Italic))
                                        fontStyle = fontStyle | FontStyle.Italic;

                                    float fontSize = Convert.ToSingle(font.Size);

                                    label1.Font = new System.Drawing.Font(font.ToString(), fontSize, fontStyle);
                                    if (Conversions.ToBoolean(!Operators.ConditionalCompareObjectEqual(cell.Interior.ColorIndex, XlColorIndex.xlColorIndexNone, false)))
                                    {
                                        long colorValue1 = Conversions.ToLong(cell.Interior.Color);
                                        int red1 = (int)(colorValue1 % 256L);
                                        int green1 = (int)(colorValue1 / 256L % 256L);
                                        int blue1 = (int)(colorValue1 / 256L / 256L % 256L);
                                        label1.BackColor = Color.FromArgb(red1, green1, blue1);
                                    }

                                    if (cell.Font.Color is DBNull)
                                    {
                                        label1.ForeColor = Color.FromArgb(0, 0, 0);
                                    }

                                    else if (Conversions.ToBoolean(!Operators.ConditionalCompareObjectEqual(cell.Font.ColorIndex, XlColorIndex.xlColorIndexNone, false)))
                                    {
                                        long colorValue2 = Conversions.ToLong(cell.Font.Color);
                                        int red2 = (int)(colorValue2 % 256L);
                                        int green2 = (int)(colorValue2 / 256L % 256L);
                                        int blue2 = (int)(colorValue2 / 256L / 256L % 256L);
                                        label1.ForeColor = Color.FromArgb(red2, green2, blue2);
                                    }
                                }
                                Panel_ExpectedOutput.Controls.Add(label1);
                            }
                            ordinate = ordinate + Width1;
                        }
                    }
                    Panel_ExpectedOutput.AutoScroll = true;
                }

                else if (X8)
                {

                    if (X1)
                    {
                        var Values = new string[(r * 2)];
                        int Index = -1;
                        for (int i = 1, loopTo13 = r; i <= loopTo13; i++)
                        {
                            string source = Conversions.ToString(displayRng.Cells[i, 1].value);
                            var NumberText = new string[2];
                            NumberText = (string[])SeparateNumberText(source);
                            Index = Index + 1;
                            Array.Resize(ref Values, Index + 1);
                            Values[Index] = NumberText[0];
                            Index = Index + 1;
                            Array.Resize(ref Values, Index + 1);
                            Values[Index] = NumberText[1];
                        }

                        double Width2 = Conversions.ToDouble(Operators.DivideObject(Operators.MultiplyObject(MaxOfArray(Values), BaseWidth), 10));

                        if (Width + Width2 < Panel_ExpectedOutput.Width)
                        {
                            Width2 = Panel_ExpectedOutput.Width - Width;
                        }

                        Index = 0;

                        for (int i = 1, loopTo14 = r; i <= loopTo14; i++)
                        {
                            var label = new System.Windows.Forms.Label();
                            label.Text = Conversions.ToString(displayRng.Cells[i, 1].Value);
                            label.Location = new System.Drawing.Point(0, (int)Math.Round(Index * Height));
                            label.Height = (int)Math.Round(Height);
                            label.Width = (int)Math.Round(Width);
                            label.BorderStyle = BorderStyle.FixedSingle;
                            label.TextAlign = ContentAlignment.MiddleCenter;

                            if (CB_formatting.Checked == true)
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
                            Panel_ExpectedOutput.Controls.Add(label);

                            var label1 = new System.Windows.Forms.Label();
                            label1.Text = Values[Index];
                            label1.Location = new System.Drawing.Point((int)Math.Round(Width), (int)Math.Round(Index * Height));
                            label1.Height = (int)Math.Round(Height);
                            label1.Width = (int)Math.Round(Width2);
                            label1.BorderStyle = BorderStyle.FixedSingle;
                            label1.TextAlign = ContentAlignment.MiddleCenter;

                            if (CB_formatting.Checked == true)
                            {
                                Range cell = (Range)displayRng.Cells[i, 1];
                                var font = cell.Font;
                                var fontStyle = FontStyle.Regular;
                                if (Conversions.ToBoolean(cell.Font.Bold))
                                    fontStyle = fontStyle | FontStyle.Bold;
                                if (Conversions.ToBoolean(cell.Font.Italic))
                                    fontStyle = fontStyle | FontStyle.Italic;

                                float fontSize = Convert.ToSingle(font.Size);

                                label1.Font = new System.Drawing.Font(font.ToString(), fontSize, fontStyle);
                                if (Conversions.ToBoolean(!Operators.ConditionalCompareObjectEqual(cell.Interior.ColorIndex, XlColorIndex.xlColorIndexNone, false)))
                                {
                                    long colorValue1 = Conversions.ToLong(cell.Interior.Color);
                                    int red1 = (int)(colorValue1 % 256L);
                                    int green1 = (int)(colorValue1 / 256L % 256L);
                                    int blue1 = (int)(colorValue1 / 256L / 256L % 256L);
                                    label1.BackColor = Color.FromArgb(red1, green1, blue1);
                                }

                                if (cell.Font.Color is DBNull)
                                {
                                    label1.ForeColor = Color.FromArgb(0, 0, 0);
                                }

                                else if (Conversions.ToBoolean(!Operators.ConditionalCompareObjectEqual(cell.Font.ColorIndex, XlColorIndex.xlColorIndexNone, false)))
                                {
                                    long colorValue2 = Conversions.ToLong(cell.Font.Color);
                                    int red2 = (int)(colorValue2 % 256L);
                                    int green2 = (int)(colorValue2 / 256L % 256L);
                                    int blue2 = (int)(colorValue2 / 256L / 256L % 256L);
                                    label1.ForeColor = Color.FromArgb(red2, green2, blue2);
                                }
                            }
                            Panel_ExpectedOutput.Controls.Add(label1);

                            Index = Index + 1;

                            var label2 = new System.Windows.Forms.Label();
                            label2.Text = "";
                            label2.Location = new System.Drawing.Point(0, (int)Math.Round(Index * Height));
                            label2.Height = (int)Math.Round(Height);
                            label2.Width = (int)Math.Round(Width);
                            label2.BorderStyle = BorderStyle.FixedSingle;
                            label2.TextAlign = ContentAlignment.MiddleCenter;

                            if (CB_formatting.Checked == true)
                            {
                                Range cell = (Range)displayRng.Cells[i, 1];
                                var font = cell.Font;
                                var fontStyle = FontStyle.Regular;
                                if (Conversions.ToBoolean(cell.Font.Bold))
                                    fontStyle = fontStyle | FontStyle.Bold;
                                if (Conversions.ToBoolean(cell.Font.Italic))
                                    fontStyle = fontStyle | FontStyle.Italic;

                                float fontSize = Convert.ToSingle(font.Size);

                                label2.Font = new System.Drawing.Font(font.ToString(), fontSize, fontStyle);
                                if (Conversions.ToBoolean(!Operators.ConditionalCompareObjectEqual(cell.Interior.ColorIndex, XlColorIndex.xlColorIndexNone, false)))
                                {
                                    long colorValue1 = Conversions.ToLong(cell.Interior.Color);
                                    int red1 = (int)(colorValue1 % 256L);
                                    int green1 = (int)(colorValue1 / 256L % 256L);
                                    int blue1 = (int)(colorValue1 / 256L / 256L % 256L);
                                    label2.BackColor = Color.FromArgb(red1, green1, blue1);
                                }

                                if (cell.Font.Color is DBNull)
                                {
                                    label2.ForeColor = Color.FromArgb(0, 0, 0);
                                }

                                else if (Conversions.ToBoolean(!Operators.ConditionalCompareObjectEqual(cell.Font.ColorIndex, XlColorIndex.xlColorIndexNone, false)))
                                {
                                    long colorValue2 = Conversions.ToLong(cell.Font.Color);
                                    int red2 = (int)(colorValue2 % 256L);
                                    int green2 = (int)(colorValue2 / 256L % 256L);
                                    int blue2 = (int)(colorValue2 / 256L / 256L % 256L);
                                    label2.ForeColor = Color.FromArgb(red2, green2, blue2);
                                }
                            }
                            Panel_ExpectedOutput.Controls.Add(label2);

                            var label3 = new System.Windows.Forms.Label();
                            label3.Text = Values[Index];
                            label3.Location = new System.Drawing.Point((int)Math.Round(Width), (int)Math.Round(Index * Height));
                            label3.Height = (int)Math.Round(Height);
                            label3.Width = (int)Math.Round(Width2);
                            label3.BorderStyle = BorderStyle.FixedSingle;
                            label3.TextAlign = ContentAlignment.MiddleCenter;

                            if (CB_formatting.Checked == true)
                            {
                                Range cell = (Range)displayRng.Cells[i, 1];
                                var font = cell.Font;
                                var fontStyle = FontStyle.Regular;
                                if (Conversions.ToBoolean(cell.Font.Bold))
                                    fontStyle = fontStyle | FontStyle.Bold;
                                if (Conversions.ToBoolean(cell.Font.Italic))
                                    fontStyle = fontStyle | FontStyle.Italic;

                                float fontSize = Convert.ToSingle(font.Size);

                                label3.Font = new System.Drawing.Font(font.ToString(), fontSize, fontStyle);
                                if (Conversions.ToBoolean(!Operators.ConditionalCompareObjectEqual(cell.Interior.ColorIndex, XlColorIndex.xlColorIndexNone, false)))
                                {
                                    long colorValue1 = Conversions.ToLong(cell.Interior.Color);
                                    int red1 = (int)(colorValue1 % 256L);
                                    int green1 = (int)(colorValue1 / 256L % 256L);
                                    int blue1 = (int)(colorValue1 / 256L / 256L % 256L);
                                    label3.BackColor = Color.FromArgb(red1, green1, blue1);
                                }

                                if (cell.Font.Color is DBNull)
                                {
                                    label3.ForeColor = Color.FromArgb(0, 0, 0);
                                }

                                else if (Conversions.ToBoolean(!Operators.ConditionalCompareObjectEqual(cell.Font.ColorIndex, XlColorIndex.xlColorIndexNone, false)))
                                {
                                    long colorValue2 = Conversions.ToLong(cell.Font.Color);
                                    int red2 = (int)(colorValue2 % 256L);
                                    int green2 = (int)(colorValue2 / 256L % 256L);
                                    int blue2 = (int)(colorValue2 / 256L / 256L % 256L);
                                    label3.ForeColor = Color.FromArgb(red2, green2, blue2);
                                }
                            }
                            Index = Index + 1;
                            Panel_ExpectedOutput.Controls.Add(label3);
                        }
                    }

                    else if (X2)
                    {
                        var Numbers = new string[r];
                        var Texts = new string[r];

                        for (int i = 1, loopTo15 = r; i <= loopTo15; i++)
                        {
                            string source = Conversions.ToString(displayRng.Cells[i, 1].value);
                            var NumberText = new string[2];
                            NumberText = (string[])SeparateNumberText(source);
                            Numbers[i - 1] = NumberText[0];
                            Texts[i - 1] = NumberText[1];
                        }

                        double Width2 = Conversions.ToDouble(Operators.DivideObject(Operators.MultiplyObject(MaxOfArray(Numbers), BaseWidth), 10));
                        double Width3 = Conversions.ToDouble(Operators.DivideObject(Operators.MultiplyObject(MaxOfArray(Texts), BaseWidth), 10));

                        if (Width + Width2 + Width3 < Panel_ExpectedOutput.Width)
                        {
                            Width3 = Panel_ExpectedOutput.Width - (Width + Width2);
                        }

                        for (int i = 1, loopTo16 = r; i <= loopTo16; i++)
                        {
                            var label = new System.Windows.Forms.Label();
                            label.Text = Conversions.ToString(displayRng.Cells[i, 1].Value);
                            label.Location = new System.Drawing.Point(0, (int)Math.Round((i - 1) * Height));
                            label.Height = (int)Math.Round(Height);
                            label.Width = (int)Math.Round(Width);
                            label.BorderStyle = BorderStyle.FixedSingle;
                            label.TextAlign = ContentAlignment.MiddleCenter;

                            if (CB_formatting.Checked == true)
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
                            Panel_ExpectedOutput.Controls.Add(label);

                            var label1 = new System.Windows.Forms.Label();
                            label1.Text = Numbers[i - 1];
                            label1.Location = new System.Drawing.Point((int)Math.Round(Width), (int)Math.Round((i - 1) * Height));
                            label1.Height = (int)Math.Round(Height);
                            label1.Width = (int)Math.Round(Width2);
                            label1.BorderStyle = BorderStyle.FixedSingle;
                            label1.TextAlign = ContentAlignment.MiddleCenter;

                            if (CB_formatting.Checked == true)
                            {
                                Range cell = (Range)displayRng.Cells[i, 1];
                                var font = cell.Font;
                                var fontStyle = FontStyle.Regular;
                                if (Conversions.ToBoolean(cell.Font.Bold))
                                    fontStyle = fontStyle | FontStyle.Bold;
                                if (Conversions.ToBoolean(cell.Font.Italic))
                                    fontStyle = fontStyle | FontStyle.Italic;

                                float fontSize = Convert.ToSingle(font.Size);

                                label1.Font = new System.Drawing.Font(font.ToString(), fontSize, fontStyle);
                                if (Conversions.ToBoolean(!Operators.ConditionalCompareObjectEqual(cell.Interior.ColorIndex, XlColorIndex.xlColorIndexNone, false)))
                                {
                                    long colorValue1 = Conversions.ToLong(cell.Interior.Color);
                                    int red1 = (int)(colorValue1 % 256L);
                                    int green1 = (int)(colorValue1 / 256L % 256L);
                                    int blue1 = (int)(colorValue1 / 256L / 256L % 256L);
                                    label1.BackColor = Color.FromArgb(red1, green1, blue1);
                                }

                                if (cell.Font.Color is DBNull)
                                {
                                    label1.ForeColor = Color.FromArgb(0, 0, 0);
                                }

                                else if (Conversions.ToBoolean(!Operators.ConditionalCompareObjectEqual(cell.Font.ColorIndex, XlColorIndex.xlColorIndexNone, false)))
                                {
                                    long colorValue2 = Conversions.ToLong(cell.Font.Color);
                                    int red2 = (int)(colorValue2 % 256L);
                                    int green2 = (int)(colorValue2 / 256L % 256L);
                                    int blue2 = (int)(colorValue2 / 256L / 256L % 256L);
                                    label1.ForeColor = Color.FromArgb(red2, green2, blue2);
                                }
                            }
                            Panel_ExpectedOutput.Controls.Add(label1);

                            var label2 = new System.Windows.Forms.Label();
                            label2.Text = Texts[i - 1];
                            label2.Location = new System.Drawing.Point((int)Math.Round(Width + Width2), (int)Math.Round((i - 1) * Height));
                            label2.Height = (int)Math.Round(Height);
                            label2.Width = (int)Math.Round(Width3);
                            label2.BorderStyle = BorderStyle.FixedSingle;
                            label2.TextAlign = ContentAlignment.MiddleCenter;

                            if (CB_formatting.Checked == true)
                            {
                                Range cell = (Range)displayRng.Cells[i, 1];
                                var font = cell.Font;
                                var fontStyle = FontStyle.Regular;
                                if (Conversions.ToBoolean(cell.Font.Bold))
                                    fontStyle = fontStyle | FontStyle.Bold;
                                if (Conversions.ToBoolean(cell.Font.Italic))
                                    fontStyle = fontStyle | FontStyle.Italic;

                                float fontSize = Convert.ToSingle(font.Size);

                                label2.Font = new System.Drawing.Font(font.ToString(), fontSize, fontStyle);
                                if (Conversions.ToBoolean(!Operators.ConditionalCompareObjectEqual(cell.Interior.ColorIndex, XlColorIndex.xlColorIndexNone, false)))
                                {
                                    long colorValue1 = Conversions.ToLong(cell.Interior.Color);
                                    int red1 = (int)(colorValue1 % 256L);
                                    int green1 = (int)(colorValue1 / 256L % 256L);
                                    int blue1 = (int)(colorValue1 / 256L / 256L % 256L);
                                    label2.BackColor = Color.FromArgb(red1, green1, blue1);
                                }

                                if (cell.Font.Color is DBNull)
                                {
                                    label2.ForeColor = Color.FromArgb(0, 0, 0);
                                }

                                else if (Conversions.ToBoolean(!Operators.ConditionalCompareObjectEqual(cell.Font.ColorIndex, XlColorIndex.xlColorIndexNone, false)))
                                {
                                    long colorValue2 = Conversions.ToLong(cell.Font.Color);
                                    int red2 = (int)(colorValue2 % 256L);
                                    int green2 = (int)(colorValue2 / 256L % 256L);
                                    int blue2 = (int)(colorValue2 / 256L / 256L % 256L);
                                    label2.ForeColor = Color.FromArgb(red2, green2, blue2);
                                }
                            }
                            Panel_ExpectedOutput.Controls.Add(label2);
                        }
                    }
                    Panel_ExpectedOutput.AutoScroll = true;
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


                    if (X1)
                    {

                        var Values = new string[1];
                        int Index = -1;

                        for (int i = 1, loopTo17 = r; i <= loopTo17; i++)
                        {
                            string source = Conversions.ToString(displayRng.Cells[i, 1].value);
                            for (double k = 1d, loopTo18 = Conversion.Int(Strings.Len(source) / (double)W); k <= loopTo18; k++)
                            {
                                Index = Index + 1;
                                Array.Resize(ref Values, Index + 1);
                                Values[Index] = Strings.Mid(source, (int)Math.Round(W * (k - 1d) + 1d), W);
                            }
                            if (Strings.Len(source) % W != 0)
                            {
                                Index = Index + 1;
                                Array.Resize(ref Values, Index + 1);
                                Values[Index] = Strings.Mid(source, Strings.Len(source) - Strings.Len(source) % W + 1, Strings.Len(source) % W);
                            }
                        }

                        double Width2 = Conversions.ToDouble(Operators.DivideObject(Operators.MultiplyObject(MaxOfArray(Values), BaseWidth), 10));

                        if (Width + Width2 < Panel_ExpectedOutput.Width)
                        {
                            Width2 = Panel_ExpectedOutput.Width - Width;
                        }


                        double abscissa1 = 0d;
                        double abscissa2 = 0d;

                        for (int i = 1, loopTo19 = r; i <= loopTo19; i++)
                        {
                            string source = Conversions.ToString(displayRng.Cells[i, 1].value);
                            var SplitValues = new string[1];
                            Index = -1;
                            for (double k = 1d, loopTo20 = Conversion.Int(Strings.Len(source) / (double)W); k <= loopTo20; k++)
                            {
                                Index = Index + 1;
                                Array.Resize(ref SplitValues, Index + 1);
                                SplitValues[Index] = Strings.Mid(source, (int)Math.Round(W * (k - 1d) + 1d), W);
                            }
                            if (Strings.Len(source) % W != 0)
                            {
                                Index = Index + 1;
                                Array.Resize(ref SplitValues, Index + 1);
                                SplitValues[Index] = Strings.Mid(source, Strings.Len(source) - Strings.Len(source) % W + 1, Strings.Len(source) % W);
                            }

                            var label = new System.Windows.Forms.Label();
                            label.Text = Conversions.ToString(displayRng.Cells[i, 1].Value);
                            label.Location = new System.Drawing.Point(0, (int)Math.Round(abscissa1));
                            label.Height = (int)Math.Round(Height);
                            label.Width = (int)Math.Round(Width);
                            label.BorderStyle = BorderStyle.FixedSingle;
                            label.TextAlign = ContentAlignment.MiddleCenter;

                            if (CB_formatting.Checked == true)
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

                            Panel_ExpectedOutput.Controls.Add(label);
                            abscissa1 = abscissa1 + Height;
                            for (int m = Information.LBound(SplitValues) + 1, loopTo21 = Information.UBound(SplitValues); m <= loopTo21; m++)
                            {
                                var label1 = new System.Windows.Forms.Label();
                                label1.Text = "";
                                label1.Location = new System.Drawing.Point(0, (int)Math.Round(abscissa1));
                                label1.Height = (int)Math.Round(Height);
                                label1.Width = (int)Math.Round(Width);
                                label1.BorderStyle = BorderStyle.FixedSingle;
                                label1.TextAlign = ContentAlignment.MiddleCenter;

                                if (CB_formatting.Checked == true)
                                {
                                    Range cell = (Range)displayRng.Cells[i, 1];
                                    var font = cell.Font;
                                    var fontStyle = FontStyle.Regular;
                                    if (Conversions.ToBoolean(cell.Font.Bold))
                                        fontStyle = fontStyle | FontStyle.Bold;
                                    if (Conversions.ToBoolean(cell.Font.Italic))
                                        fontStyle = fontStyle | FontStyle.Italic;

                                    float fontSize = Convert.ToSingle(font.Size);

                                    label1.Font = new System.Drawing.Font(font.ToString(), fontSize, fontStyle);
                                    if (Conversions.ToBoolean(!Operators.ConditionalCompareObjectEqual(cell.Interior.ColorIndex, XlColorIndex.xlColorIndexNone, false)))
                                    {
                                        long colorValue1 = Conversions.ToLong(cell.Interior.Color);
                                        int red1 = (int)(colorValue1 % 256L);
                                        int green1 = (int)(colorValue1 / 256L % 256L);
                                        int blue1 = (int)(colorValue1 / 256L / 256L % 256L);
                                        label1.BackColor = Color.FromArgb(red1, green1, blue1);
                                    }

                                    if (cell.Font.Color is DBNull)
                                    {
                                        label1.ForeColor = Color.FromArgb(0, 0, 0);
                                    }

                                    else if (Conversions.ToBoolean(!Operators.ConditionalCompareObjectEqual(cell.Font.ColorIndex, XlColorIndex.xlColorIndexNone, false)))
                                    {
                                        long colorValue2 = Conversions.ToLong(cell.Font.Color);
                                        int red2 = (int)(colorValue2 % 256L);
                                        int green2 = (int)(colorValue2 / 256L % 256L);
                                        int blue2 = (int)(colorValue2 / 256L / 256L % 256L);
                                        label1.ForeColor = Color.FromArgb(red2, green2, blue2);
                                    }
                                }
                                Panel_ExpectedOutput.Controls.Add(label1);
                                abscissa1 = abscissa1 + Height;
                            }

                            for (int m = Information.LBound(SplitValues), loopTo22 = Information.UBound(SplitValues); m <= loopTo22; m++)
                            {
                                var label1 = new System.Windows.Forms.Label();
                                label1.Text = SplitValues[m];
                                label1.Location = new System.Drawing.Point((int)Math.Round(Width), (int)Math.Round(abscissa2));
                                label1.Height = (int)Math.Round(Height);
                                label1.Width = (int)Math.Round(Width2);
                                label1.BorderStyle = BorderStyle.FixedSingle;
                                label1.TextAlign = ContentAlignment.MiddleCenter;

                                if (CB_formatting.Checked == true)
                                {
                                    Range cell = (Range)displayRng.Cells[i, 1];
                                    var font = cell.Font;
                                    var fontStyle = FontStyle.Regular;
                                    if (Conversions.ToBoolean(cell.Font.Bold))
                                        fontStyle = fontStyle | FontStyle.Bold;
                                    if (Conversions.ToBoolean(cell.Font.Italic))
                                        fontStyle = fontStyle | FontStyle.Italic;

                                    float fontSize = Convert.ToSingle(font.Size);

                                    label1.Font = new System.Drawing.Font(font.ToString(), fontSize, fontStyle);
                                    if (Conversions.ToBoolean(!Operators.ConditionalCompareObjectEqual(cell.Interior.ColorIndex, XlColorIndex.xlColorIndexNone, false)))
                                    {
                                        long colorValue1 = Conversions.ToLong(cell.Interior.Color);
                                        int red1 = (int)(colorValue1 % 256L);
                                        int green1 = (int)(colorValue1 / 256L % 256L);
                                        int blue1 = (int)(colorValue1 / 256L / 256L % 256L);
                                        label1.BackColor = Color.FromArgb(red1, green1, blue1);
                                    }

                                    if (cell.Font.Color is DBNull)
                                    {
                                        label1.ForeColor = Color.FromArgb(0, 0, 0);
                                    }

                                    else if (Conversions.ToBoolean(!Operators.ConditionalCompareObjectEqual(cell.Font.ColorIndex, XlColorIndex.xlColorIndexNone, false)))
                                    {
                                        long colorValue2 = Conversions.ToLong(cell.Font.Color);
                                        int red2 = (int)(colorValue2 % 256L);
                                        int green2 = (int)(colorValue2 / 256L % 256L);
                                        int blue2 = (int)(colorValue2 / 256L / 256L % 256L);
                                        label1.ForeColor = Color.FromArgb(red2, green2, blue2);
                                    }
                                }
                                Panel_ExpectedOutput.Controls.Add(label1);
                                abscissa2 = abscissa2 + Height;
                            }
                        }
                    }

                    else if (X2)
                    {

                        var Lengths = new int[r];

                        for (int i = 1, loopTo23 = r; i <= loopTo23; i++)
                        {
                            string source = Conversions.ToString(displayRng.Cells[i, 1].value);
                            int Index = 0;
                            for (double k = 1d, loopTo24 = Conversion.Int(Strings.Len(source) / (double)W); k <= loopTo24; k++)
                                Index = Index + 1;
                            if (Strings.Len(source) % W != 0)
                            {
                                Index = Index + 1;
                            }
                            Lengths[i - 1] = Index;
                        }

                        int TotalColumns = Conversions.ToInteger(FindMax(Lengths));
                        var SplitValues = new string[r, TotalColumns];

                        int Index2 = -1;
                        for (int i = 1, loopTo25 = r; i <= loopTo25; i++)
                        {
                            string source = Conversions.ToString(displayRng.Cells[i, 1].value);
                            Index2 = Index2 + 1;
                            for (double k = 1d, loopTo26 = Conversion.Int(Strings.Len(source) / (double)W); k <= loopTo26; k++)
                                SplitValues[Index2, (int)Math.Round(k - 1d)] = Strings.Mid(source, (int)Math.Round(W * (k - 1d) + 1d), W);
                            if (Strings.Len(source) % W != 0)
                            {
                                SplitValues[Index2, (int)Math.Round(Conversion.Int(Strings.Len(source) / (double)W))] = Strings.Mid(source, Strings.Len(source) - Strings.Len(source) % W + 1, Strings.Len(source) % W);
                            }
                        }

                        ordinate = 0d;
                        Width = Conversions.ToDouble(Operators.DivideObject(Operators.MultiplyObject(MaxOfColumn(displayRng), BaseWidth), 10));

                        for (int i = 1, loopTo27 = r; i <= loopTo27; i++)
                        {
                            var label1 = new System.Windows.Forms.Label();
                            label1.Text = Conversions.ToString(displayRng.Cells[i, 1].value);
                            label1.Location = new System.Drawing.Point((int)Math.Round(ordinate), (int)Math.Round((i - 1) * Height));
                            label1.Height = (int)Math.Round(Height);
                            label1.Width = (int)Math.Round(Width);
                            label1.BorderStyle = BorderStyle.FixedSingle;
                            label1.TextAlign = ContentAlignment.MiddleCenter;

                            if (CB_formatting.Checked == true)
                            {
                                Range cell = (Range)displayRng.Cells[i, 1];
                                var font = cell.Font;
                                var fontStyle = FontStyle.Regular;
                                if (Conversions.ToBoolean(cell.Font.Bold))
                                    fontStyle = fontStyle | FontStyle.Bold;
                                if (Conversions.ToBoolean(cell.Font.Italic))
                                    fontStyle = fontStyle | FontStyle.Italic;

                                float fontSize = Convert.ToSingle(font.Size);

                                label1.Font = new System.Drawing.Font(font.ToString(), fontSize, fontStyle);
                                if (Conversions.ToBoolean(!Operators.ConditionalCompareObjectEqual(cell.Interior.ColorIndex, XlColorIndex.xlColorIndexNone, false)))
                                {
                                    long colorValue1 = Conversions.ToLong(cell.Interior.Color);
                                    int red1 = (int)(colorValue1 % 256L);
                                    int green1 = (int)(colorValue1 / 256L % 256L);
                                    int blue1 = (int)(colorValue1 / 256L / 256L % 256L);
                                    label1.BackColor = Color.FromArgb(red1, green1, blue1);
                                }

                                if (cell.Font.Color is DBNull)
                                {
                                    label1.ForeColor = Color.FromArgb(0, 0, 0);
                                }

                                else if (Conversions.ToBoolean(!Operators.ConditionalCompareObjectEqual(cell.Font.ColorIndex, XlColorIndex.xlColorIndexNone, false)))
                                {
                                    long colorValue2 = Conversions.ToLong(cell.Font.Color);
                                    int red2 = (int)(colorValue2 % 256L);
                                    int green2 = (int)(colorValue2 / 256L % 256L);
                                    int blue2 = (int)(colorValue2 / 256L / 256L % 256L);
                                    label1.ForeColor = Color.FromArgb(red2, green2, blue2);
                                }
                            }
                            Panel_ExpectedOutput.Controls.Add(label1);
                        }
                        ordinate = ordinate + Width;

                        for (int j = 1, loopTo28 = TotalColumns; j <= loopTo28; j++)
                        {
                            var Columns = new string[r];
                            for (int i = 1, loopTo29 = r; i <= loopTo29; i++)
                                Columns[i - 1] = SplitValues[i - 1, j - 1];
                            Width = Conversions.ToDouble(Operators.DivideObject(Operators.MultiplyObject(MaxOfArray(Columns), BaseWidth), 10));
                            for (int i = 1, loopTo30 = r; i <= loopTo30; i++)
                            {
                                var label1 = new System.Windows.Forms.Label();
                                label1.Text = Columns[i - 1];
                                label1.Location = new System.Drawing.Point((int)Math.Round(ordinate), (int)Math.Round((i - 1) * Height));
                                label1.Height = (int)Math.Round(Height);
                                label1.Width = (int)Math.Round(Width);
                                label1.BorderStyle = BorderStyle.FixedSingle;
                                label1.TextAlign = ContentAlignment.MiddleCenter;

                                if (CB_formatting.Checked == true)
                                {
                                    Range cell = (Range)displayRng.Cells[i, 1];
                                    var font = cell.Font;
                                    var fontStyle = FontStyle.Regular;
                                    if (Conversions.ToBoolean(cell.Font.Bold))
                                        fontStyle = fontStyle | FontStyle.Bold;
                                    if (Conversions.ToBoolean(cell.Font.Italic))
                                        fontStyle = fontStyle | FontStyle.Italic;

                                    float fontSize = Convert.ToSingle(font.Size);

                                    label1.Font = new System.Drawing.Font(font.ToString(), fontSize, fontStyle);
                                    if (Conversions.ToBoolean(!Operators.ConditionalCompareObjectEqual(cell.Interior.ColorIndex, XlColorIndex.xlColorIndexNone, false)))
                                    {
                                        long colorValue1 = Conversions.ToLong(cell.Interior.Color);
                                        int red1 = (int)(colorValue1 % 256L);
                                        int green1 = (int)(colorValue1 / 256L % 256L);
                                        int blue1 = (int)(colorValue1 / 256L / 256L % 256L);
                                        label1.BackColor = Color.FromArgb(red1, green1, blue1);
                                    }

                                    if (cell.Font.Color is DBNull)
                                    {
                                        label1.ForeColor = Color.FromArgb(0, 0, 0);
                                    }

                                    else if (Conversions.ToBoolean(!Operators.ConditionalCompareObjectEqual(cell.Font.ColorIndex, XlColorIndex.xlColorIndexNone, false)))
                                    {
                                        long colorValue2 = Conversions.ToLong(cell.Font.Color);
                                        int red2 = (int)(colorValue2 % 256L);
                                        int green2 = (int)(colorValue2 / 256L % 256L);
                                        int blue2 = (int)(colorValue2 / 256L / 256L % 256L);
                                        label1.ForeColor = Color.FromArgb(red2, green2, blue2);
                                    }
                                }
                                Panel_ExpectedOutput.Controls.Add(label1);
                            }
                            ordinate = ordinate + Width;
                        }
                    }

                    Panel_ExpectedOutput.AutoScroll = true;

                }
            }
        }
        private void CB_separators_finaloutput_CheckedChanged(object sender, EventArgs e)
        {

            try
            {
                if (CB_separators_finaloutput.Checked == true)
                {
                    RB_starting_point.Enabled = true;
                    RB_ending_point.Enabled = true;
                    PictureBox2.Enabled = true;
                    PictureBox3.Enabled = true;
                }

                else if (CB_separators_finaloutput.Checked == false)
                {
                    RB_starting_point.Enabled = false;
                    RB_ending_point.Enabled = false;
                    PictureBox2.Enabled = false;
                    PictureBox3.Enabled = false;
                }
                Display();
            }
            catch (Exception ex)
            {

            }

        }

        private void Btn_OK_Click(object sender, EventArgs e)
        {

            if (string.IsNullOrEmpty(TB_source_range.Text))
            {
                MessageBox.Show("Select a Source Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                TB_source_range.Focus();
                workSheet.Activate();
                rng.Select();
                return;
            }

            if (IsValidExcelCellReference(TB_source_range.Text) == false)
            {
                MessageBox.Show("Select a Valid Source Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                TB_source_range.Focus();
                workSheet.Activate();
                rng.Select();
                return;
            }

            int r = rng.Rows.Count;
            int c = rng.Columns.Count;

            bool X1 = RB_rows.Checked;
            bool X2 = RB_columns.Checked;
            bool X3 = RB_space.Checked;
            bool X7 = RB_newline.Checked;
            bool X8 = RB_numbertext.Checked;
            bool X9 = RB_semicolon.Checked;
            bool X10 = RB_others.Checked;
            bool X11 = RB_width.Checked;

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

            workSheet.Activate();

            if ((X1 | X2) & (X3 | X7 | X8 | X9 | X10 | X11))
            {

                if (X3 | X7 | X9 | X10)
                {

                    string Separator = ",";
                    if (X7)
                    {
                        Separator = Microsoft.VisualBasic.Constants.vbNewLine;
                    }
                    else if (X3)
                    {
                        Separator = " ";
                    }
                    else if (X9)
                    {
                        Separator = ";";
                    }
                    else if (X10)
                    {
                        Separator = ComboBox2.Text;
                    }

                    bool Consecutive;
                    if (CB_consecute_separators.Checked)
                    {
                        Consecutive = true;
                    }
                    else
                    {
                        Consecutive = false;
                    }

                    bool KeepSeparator;
                    if (CB_separators_finaloutput.Checked)
                    {
                        KeepSeparator = true;
                    }
                    else
                    {
                        KeepSeparator = false;
                    }

                    bool Before;
                    if (RB_starting_point.Checked)
                    {
                        Before = true;
                    }
                    else
                    {
                        Before = false;
                    }

                    if (X1)
                    {

                        var Arr = new string[r];
                        var Lengths = new string[r];
                        int RowNumber = 0;
                        for (int i = 1, loopTo = r; i <= loopTo; i++)
                        {
                            string source = Conversions.ToString(rng.Cells[i, 1].value);
                            Arr[i - 1] = source;
                            string[] SplitValues;
                            SplitValues = (string[])SplitText(source, Separator, Consecutive, KeepSeparator, Before);
                            Lengths[i - 1] = (Information.UBound(SplitValues) + 1).ToString();
                            for (int m = Information.LBound(SplitValues), loopTo1 = Information.UBound(SplitValues); m <= loopTo1; m++)
                            {
                                RowNumber = RowNumber + 1;
                                rng.Cells[RowNumber, 2] = SplitValues[m];
                                if (CB_formatting.Checked)
                                {
                                    rng.Cells[i, 1].Copy();
                                    rng.Cells[RowNumber, 2].PasteSpecial(XlPasteType.xlPasteFormats);
                                }
                                else
                                {
                                    rng.Cells[RowNumber, 2].ClearFormats();
                                }
                            }
                        }

                        RowNumber = 0;
                        for (int i = 1, loopTo2 = r; i <= loopTo2; i++)
                        {
                            RowNumber = RowNumber + 1;
                            rng.Cells[RowNumber, 1] = Arr[i - 1];
                            if (CB_formatting.Checked)
                            {
                                rng.Cells[RowNumber, 2].Copy();
                                rng.Cells[RowNumber, 1].PasteSpecial(XlPasteType.xlPasteFormats);
                            }
                            else
                            {
                                rng.Cells[RowNumber, 1].ClearFormats();
                            }
                            for (double m = 1d, loopTo3 = Conversions.ToDouble(Lengths[i - 1]) - 1d; m <= loopTo3; m++)
                            {
                                RowNumber = RowNumber + 1;
                                rng.Cells[RowNumber, 1] = "";
                                if (CB_formatting.Checked)
                                {
                                    rng.Cells[RowNumber, 2].Copy();
                                    rng.Cells[RowNumber, 1].PasteSpecial(XlPasteType.xlPasteFormats);
                                }
                                else
                                {
                                    rng.Cells[RowNumber, 1].ClearFormats();
                                }
                            }
                        }

                        excelApp.CutCopyMode = (XlCutCopyMode)Conversions.ToInteger(false);

                        rng2 = workSheet.get_Range(rng.Cells[1, 1], rng.Cells[RowNumber, 2]);
                        rng2.Select();
                        for (int j = 1, loopTo4 = rng2.Columns.Count; j <= loopTo4; j++)
                            rng2.Columns[j].Autofit();
                    }

                    else if (X2)
                    {

                        int MaxColumns = 1;
                        for (int i = 1, loopTo5 = r; i <= loopTo5; i++)
                        {
                            string source = Conversions.ToString(rng.Cells[i, 1].value);
                            string[] SplitValues;
                            SplitValues = (string[])SplitText(source, Separator, Consecutive, KeepSeparator, Before);
                            if (Information.UBound(SplitValues) + 1 > MaxColumns)
                            {
                                MaxColumns = Information.UBound(SplitValues) + 1;
                            }
                            if (CB_formatting.Checked == false)
                            {
                                rng.Cells[i, 1].ClearFormats();
                            }
                            for (int m = Information.LBound(SplitValues), loopTo6 = Information.UBound(SplitValues); m <= loopTo6; m++)
                                rng.Cells[i, m + 2] = SplitValues[m];
                        }
                        for (int i = 1, loopTo7 = r; i <= loopTo7; i++)
                        {
                            if (CB_formatting.Checked)
                            {
                                rng.Cells[i, 1].Copy();
                                for (int m = 1, loopTo8 = MaxColumns; m <= loopTo8; m++)
                                    rng.Cells[i, m + 1].PasteSpecial(XlPasteType.xlPasteFormats);
                            }
                            else
                            {
                                for (int m = 1, loopTo9 = MaxColumns; m <= loopTo9; m++)
                                    rng.Cells[i, m + 1].ClearFormats();
                            }
                        }

                        excelApp.CutCopyMode = (XlCutCopyMode)Conversions.ToInteger(false);

                        rng2 = workSheet.get_Range(rng.Cells[1, 1], rng.Cells[r, MaxColumns + 1]);
                        rng2.Select();
                        for (int j = 1, loopTo10 = rng2.Columns.Count; j <= loopTo10; j++)
                            rng2.Columns[j].Autofit();

                    }
                }

                else if (X8)
                {

                    if (X1)
                    {
                        var Arr = new string[r];
                        int RowNumber = 0;
                        for (int i = 1, loopTo11 = r; i <= loopTo11; i++)
                        {
                            string source = Conversions.ToString(rng.Cells[i, 1].value);
                            Arr[i - 1] = source;
                            string[] SplitValues;
                            SplitValues = (string[])SeparateNumberText(source);
                            for (int m = 0; m <= 1; m++)
                            {
                                RowNumber = RowNumber + 1;
                                rng.Cells[RowNumber, 2] = SplitValues[m];
                                if (CB_formatting.Checked)
                                {
                                    rng.Cells[i, 1].Copy();
                                    rng.Cells[RowNumber, 2].PasteSpecial(XlPasteType.xlPasteFormats);
                                }
                                else
                                {
                                    rng.Cells[RowNumber, 2].ClearFormats();
                                }
                            }
                        }

                        RowNumber = 0;

                        for (int i = 1, loopTo12 = r; i <= loopTo12; i++)
                        {
                            RowNumber = RowNumber + 1;
                            rng.Cells[RowNumber, 1] = Arr[i - 1];
                            if (CB_formatting.Checked)
                            {
                                rng.Cells[RowNumber, 2].Copy();
                                rng.Cells[RowNumber, 1].PasteSpecial(XlPasteType.xlPasteFormats);
                            }
                            else
                            {
                                rng.Cells[RowNumber, 1].ClearFormats();
                            }
                            RowNumber = RowNumber + 1;
                            rng.Cells[RowNumber, 1] = "";
                            if (CB_formatting.Checked)
                            {
                                rng.Cells[RowNumber, 2].Copy();
                                rng.Cells[RowNumber, 1].PasteSpecial(XlPasteType.xlPasteFormats);
                            }
                            else
                            {
                                rng.Cells[RowNumber, 1].ClearFormats();
                            }
                        }

                        excelApp.CutCopyMode = (XlCutCopyMode)Conversions.ToInteger(false);

                        rng2 = workSheet.get_Range(rng.Cells[1, 1], rng.Cells[RowNumber, 2]);
                        rng2.Select();
                        for (int j = 1, loopTo13 = rng2.Columns.Count; j <= loopTo13; j++)
                            rng2.Columns[j].Autofit();
                    }

                    else if (X2)
                    {

                        int MaxColumns = 2;
                        for (int i = 1, loopTo14 = r; i <= loopTo14; i++)
                        {
                            string source = Conversions.ToString(rng.Cells[i, 1].value);
                            string[] SplitValues;
                            SplitValues = (string[])SeparateNumberText(source);
                            if (CB_formatting.Checked == false)
                            {
                                rng.Cells[i, 1].ClearFormats();
                            }
                            for (int m = 0; m <= 1; m++)
                                rng.Cells[i, m + 2] = SplitValues[m];
                            if (CB_formatting.Checked)
                            {
                                rng.Cells[i, 1].Copy();
                                for (int m = 1, loopTo15 = MaxColumns; m <= loopTo15; m++)
                                    rng.Cells[i, m + 1].PasteSpecial(XlPasteType.xlPasteFormats);
                            }
                            else
                            {
                                for (int m = 1, loopTo16 = MaxColumns; m <= loopTo16; m++)
                                    rng.Cells[i, m + 1].ClearFormats();
                            }
                        }

                        excelApp.CutCopyMode = (XlCutCopyMode)Conversions.ToInteger(false);

                        rng2 = workSheet.get_Range(rng.Cells[1, 1], rng.Cells[r, MaxColumns + 1]);
                        rng2.Select();
                        for (int j = 1, loopTo17 = rng2.Columns.Count; j <= loopTo17; j++)
                            rng2.Columns[j].Autofit();

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


                    if (X1)
                    {

                        var Arr = new string[r];
                        var Lengths = new string[r];
                        int RowNumber = 0;
                        for (int i = 1, loopTo18 = r; i <= loopTo18; i++)
                        {
                            string source = Conversions.ToString(rng.Cells[i, 1].value);
                            Arr[i - 1] = source;
                            string[] SplitValues;
                            SplitValues = (string[])SplitByWidth(source, W);
                            Lengths[i - 1] = (Information.UBound(SplitValues) + 1).ToString();
                            for (int m = Information.LBound(SplitValues), loopTo19 = Information.UBound(SplitValues); m <= loopTo19; m++)
                            {
                                RowNumber = RowNumber + 1;
                                rng.Cells[RowNumber, 2] = SplitValues[m];
                                if (CB_formatting.Checked)
                                {
                                    rng.Cells[i, 1].Copy();
                                    rng.Cells[RowNumber, 2].PasteSpecial(XlPasteType.xlPasteFormats);
                                }
                                else
                                {
                                    rng.Cells[RowNumber, 2].ClearFormats();
                                }
                            }
                        }

                        RowNumber = 0;
                        for (int i = 1, loopTo20 = r; i <= loopTo20; i++)
                        {
                            RowNumber = RowNumber + 1;
                            rng.Cells[RowNumber, 1] = Arr[i - 1];
                            if (CB_formatting.Checked)
                            {
                                rng.Cells[RowNumber, 2].Copy();
                                rng.Cells[RowNumber, 1].PasteSpecial(XlPasteType.xlPasteFormats);
                            }
                            else
                            {
                                rng.Cells[RowNumber, 1].ClearFormats();
                            }
                            for (double m = 1d, loopTo21 = Conversions.ToDouble(Lengths[i - 1]) - 1d; m <= loopTo21; m++)
                            {
                                RowNumber = RowNumber + 1;
                                rng.Cells[RowNumber, 1] = "";
                                if (CB_formatting.Checked)
                                {
                                    rng.Cells[RowNumber, 2].Copy();
                                    rng.Cells[RowNumber, 1].PasteSpecial(XlPasteType.xlPasteFormats);
                                }
                                else
                                {
                                    rng.Cells[RowNumber, 1].ClearFormats();
                                }
                            }
                        }

                        excelApp.CutCopyMode = (XlCutCopyMode)Conversions.ToInteger(false);

                        rng2 = workSheet.get_Range(rng.Cells[1, 1], rng.Cells[RowNumber, 2]);
                        rng2.Select();
                        for (int j = 1, loopTo22 = rng2.Columns.Count; j <= loopTo22; j++)
                            rng2.Columns[j].Autofit();
                    }

                    else if (X2)
                    {

                        int MaxColumns = 1;
                        for (int i = 1, loopTo23 = r; i <= loopTo23; i++)
                        {
                            string source = Conversions.ToString(rng.Cells[i, 1].value);
                            string[] SplitValues;
                            SplitValues = (string[])SplitByWidth(source, W);
                            if (Information.UBound(SplitValues) + 1 > MaxColumns)
                            {
                                MaxColumns = Information.UBound(SplitValues) + 1;
                            }
                            if (CB_formatting.Checked == false)
                            {
                                rng.Cells[i, 1].ClearFormats();
                            }
                            for (int m = Information.LBound(SplitValues), loopTo24 = Information.UBound(SplitValues); m <= loopTo24; m++)
                                rng.Cells[i, m + 2] = SplitValues[m];
                        }
                        for (int i = 1, loopTo25 = r; i <= loopTo25; i++)
                        {
                            if (CB_formatting.Checked)
                            {
                                rng.Cells[i, 1].Copy();
                                for (int m = 1, loopTo26 = MaxColumns; m <= loopTo26; m++)
                                    rng.Cells[i, m + 1].PasteSpecial(XlPasteType.xlPasteFormats);
                            }
                            else
                            {
                                for (int m = 1, loopTo27 = MaxColumns; m <= loopTo27; m++)
                                    rng.Cells[i, m + 1].ClearFormats();
                            }
                        }

                        excelApp.CutCopyMode = (XlCutCopyMode)Conversions.ToInteger(false);

                        rng2 = workSheet.get_Range(rng.Cells[1, 1], rng.Cells[r, MaxColumns + 1]);
                        rng2.Select();
                        for (int j = 1, loopTo28 = rng2.Columns.Count; j <= loopTo28; j++)
                            rng2.Columns[j].Autofit();

                    }

                }

                Close();

            }
        }
        private void Btn_OK_MouseEnter(object sender, EventArgs e)
        {
            try
            {
                Btn_OK.BackColor = Color.FromArgb(65, 105, 225);
                Btn_OK.ForeColor = Color.FromArgb(255, 255, 255);
            }
            catch (Exception ex)
            {

            }
        }

        private void Btn_Cancel_MouseEnter(object sender, EventArgs e)
        {

            try
            {
                Btn_Cancel.BackColor = Color.FromArgb(65, 105, 225);
                Btn_Cancel.ForeColor = Color.FromArgb(255, 255, 255);
            }
            catch (Exception ex)
            {
            }
        }

        private void Btn_OK_MouseLeave(object sender, EventArgs e)
        {
            try
            {

                Btn_OK.BackColor = Color.FromArgb(255, 255, 255);
                Btn_OK.ForeColor = Color.FromArgb(70, 70, 70);
            }
            catch (Exception ex)
            {

            }
        }

        private void Btn_Cancel_MouseLeave(object sender, EventArgs e)
        {
            try
            {

                Btn_Cancel.BackColor = Color.FromArgb(255, 255, 255);
                Btn_Cancel.ForeColor = Color.FromArgb(70, 70, 70);
            }
            catch (Exception ex)
            {
            }
        }

        private void RB_rows_CheckedChanged(object sender, EventArgs e)
        {

            try
            {
                if (RB_rows.Checked)
                {
                    Display();
                }
            }
            catch (Exception ex)
            {
            }
        }

        private void RB_columns_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                if (RB_columns.Checked)
                {
                    Display();
                }
            }
            catch (Exception ex)
            {

            }
        }

        private void CB_formatting_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                Display();
            }
            catch (Exception ex)
            {

            }
        }

        private void RB_space_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                if (RB_space.Checked)
                {
                    CustomGroupBox2.Enabled = true;
                    CB_consecute_separators.Enabled = true;
                    Display();
                }
            }
            catch (Exception ex)
            {

            }
        }

        private void RB_newline_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                if (RB_newline.Checked)
                {
                    CustomGroupBox2.Enabled = true;
                    CB_consecute_separators.Enabled = true;
                    Display();
                }
            }
            catch (Exception ex)
            {

            }
        }

        private void RB_numbertext_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                if (RB_numbertext.Checked)
                {
                    CustomGroupBox2.Enabled = false;
                    CB_consecute_separators.Enabled = false;
                    Display();
                }
            }
            catch (Exception ex)
            {

            }
        }

        private void RB_semicolon_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                if (RB_semicolon.Checked)
                {
                    CustomGroupBox2.Enabled = true;
                    CB_consecute_separators.Enabled = true;
                    Display();
                }
            }
            catch (Exception ex)
            {

            }
        }

        private void RB_others_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                if (RB_others.Checked)
                {
                    CustomGroupBox2.Enabled = true;
                    CB_consecute_separators.Enabled = true;
                    Display();
                }
            }
            catch (Exception ex)
            {

            }
        }

        private void RB_width_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                if (RB_width.Checked)
                {
                    CustomGroupBox2.Enabled = false;
                    CB_consecute_separators.Enabled = false;
                    Display();
                }
            }
            catch (Exception ex)
            {

            }
        }

        private void CB_consecute_separators_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                Display();
            }
            catch (Exception ex)
            {

            }
        }

        private void RB_starting_point_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                Display();
            }
            catch (Exception ex)
            {

            }
        }

        private void RB_ending_point_CheckedChanged(object sender, EventArgs e)
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

        private void TB_source_range_TextChanged(object sender, EventArgs e)
        {
            try
            {
                excelApp = Globals.ThisAddIn.Application;
                workBook = excelApp.ActiveWorkbook;
                workSheet = (Excel.Worksheet)workBook.ActiveSheet;

                TB_source_range.SelectionStart = TB_source_range.Text.Length;
                TB_source_range.ScrollToCaret();

                rng = workSheet.get_Range(TB_source_range.Text);
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

                Range userInput = (Range)excelApp.InputBox("Select a range", Type: 8);
                rng = userInput;

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

                rng = excelApp.get_Range(rng, rng.get_End(XlDirection.xlDown));
                rng = excelApp.get_Range(rng, rng.get_End(XlDirection.xlToRight));

                rng.Select();
                TB_source_range.Text = rng.get_Address();
                TB_source_range.Focus();
            }

            catch (Exception ex)
            {

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

                TB_source_range.Text = rng.get_Address();
                TB_source_range.Focus();
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

        private void Form26_split_text_bycharacters_Load(object sender, EventArgs e)
        {
            try
            {

                excelApp = Globals.ThisAddIn.Application;
                workBook = excelApp.ActiveWorkbook;
                workSheet = (Excel.Worksheet)workBook.ActiveSheet;
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

                excelApp = Globals.ThisAddIn.Application;
                Range selectedRange;
                selectedRange = (Range)excelApp.Selection;

                if (TextBoxChanged == false)
                {
                    if (FocusedTextBox == 1)
                    {
                        TB_source_range.Text = selectedRange.get_Address();
                        workSheet = (Excel.Worksheet)workBook.ActiveSheet;
                        rng = selectedRange;
                        TB_source_range.Focus();
                    }
                }
            }

            catch (Exception ex)
            {

            }

        }

        private void Form26_split_text_bycharacters_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.KeyCode == Keys.Enter)
                {

                    Btn_OK_Click(sender, e);

                }
            }

            catch (Exception ex)
            {

            }
        }

        private void Btn_Cancel_Click(object sender, EventArgs e)
        {
            try
            {
                Close();
            }
            catch (Exception ex)
            {

            }
        }

        private void TB_source_range_GotFocus(object sender, EventArgs e)
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

        private void Form26_split_text_bycharacters_Closing(object sender, CancelEventArgs e)
        {
            GlobalModule.form_flag = false;
        }

        private void Form26_split_text_bycharacters_Disposed(object sender, EventArgs e)
        {
            GlobalModule.form_flag = false;
        }

        private void Form26_split_text_bycharacters_Shown(object sender, EventArgs e)
        {
            Focus();
            BringToFront();
            Activate();
            BeginInvoke(new System.Action(() =>
                {
                    TB_source_range.Text = rng.get_Address();
                    SetWindowPos(Handle, new IntPtr(HWND_TOPMOST), 0, 0, 0, 0, SWP_NOACTIVATE | SWP_NOMOVE | SWP_NOSIZE);
                }));
        }
    }
}