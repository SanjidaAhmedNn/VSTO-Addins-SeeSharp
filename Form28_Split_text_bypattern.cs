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

    public partial class Form28_Split_text_bypattern
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

        public Form28_Split_text_bypattern()
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

            if (regex.IsMatch(cellReference))
            {
                return true;
            }
            else
            {
                return false;
            }

        }
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
        private object MatchArr(object Arr, object source, object index)
        {
            object MatchArrRet = default;

            var Matched = new object[2];
            Matched[0] = false;
            string value;

            for (int i = Information.LBound((Array)Arr), loopTo = Information.UBound((Array)Arr); i <= loopTo; i++)
            {
                value = Strings.Mid(Conversions.ToString(source), Conversions.ToInteger(index), Strings.Len(Arr((object)i)));
                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(Arr((object)i), value, false)))
                {
                    Matched[0] = true;
                    Matched[1] = Arr((object)i);
                    break;
                }
            }

            MatchArrRet = Matched;
            return MatchArrRet;

        }
        private object SplitText(string Source, string Pattern, bool Consecutive, bool KeepSeparator, bool Before)
        {
            object SplitTextRet = default;

            var SplitValues = new string[1];
            int Index = -1;
            int Start = 1;
            int SearchStart = 1;
            int Ending;
            string separator = "";

            for (int i = 1, loopTo = Strings.Len(Pattern); i <= loopTo; i++)
            {

                if (Strings.Mid(Pattern, i, 1) != "*")
                {
                    int SeparatorLength = 1;
                    while (Strings.Mid(Pattern, i + SeparatorLength, 1) != "*" & i + SeparatorLength <= Strings.Len(Pattern))
                        SeparatorLength = SeparatorLength + 1;

                    separator = Strings.Mid(Pattern, i, SeparatorLength);

                    Ending = Strings.InStr(SearchStart, Source, separator);

                    if (Ending != 0)
                    {
                        Index = Index + 1;
                        Array.Resize(ref SplitValues, Index + 1);
                        if (KeepSeparator == true)
                        {
                            if (Before == true)
                            {
                                SplitValues[Index] = Strings.Mid(Source, Start, Ending - Start);
                            }
                            else
                            {
                                SplitValues[Index] = Strings.Mid(Source, Start, Ending - Start + Strings.Len(separator));
                            }
                        }
                        else
                        {
                            SplitValues[Index] = Strings.Mid(Source, Start, Ending - Start);
                        }
                        SearchStart = Ending + Strings.Len(separator);
                        Start = Ending + Strings.Len(separator);
                        if (Consecutive == true)
                        {
                            while ((Strings.Mid(Source, SearchStart, Strings.Len(separator)) ?? "") == (separator ?? ""))
                            {
                                SearchStart = SearchStart + Strings.Len(separator);
                                Start = Start + Strings.Len(separator);
                            }
                        }
                        if (KeepSeparator == true)
                        {
                            if (Before == true)
                            {
                                Start = Start - Strings.Len(separator);
                            }
                        }
                    }

                }

            }

            Ending = Strings.Len(Source) + 1;
            Index = Index + 1;
            Array.Resize(ref SplitValues, Index + 1);

            if (KeepSeparator == true)
            {
                if (Before == true)
                {
                    SplitValues[Index] = Strings.Mid(Source, Start, Ending - Start);
                }
                else
                {
                    SplitValues[Index] = Strings.Mid(Source, Start, Ending - Start + Strings.Len(separator));
                }
            }
            else
            {
                SplitValues[Index] = Strings.Mid(Source, Start, Ending - Start);
            }

            SplitTextRet = SplitValues;
            return SplitTextRet;

        }
        private object SplitCount(string Source, string Pattern, bool Consecutive)
        {
            object SplitCountRet = default;

            int Index = 0;
            int SearchStart = 1;
            int Ending;
            string separator = "";

            for (int i = 1, loopTo = Strings.Len(Pattern); i <= loopTo; i++)
            {

                if (Strings.Mid(Pattern, i, 1) != "*")
                {
                    int SeparatorLength = 1;
                    while (Strings.Mid(Pattern, i + SeparatorLength, 1) != "*" & i + SeparatorLength <= Strings.Len(Pattern))
                        SeparatorLength = SeparatorLength + 1;

                    separator = Strings.Mid(Pattern, i, SeparatorLength);

                    Ending = Strings.InStr(SearchStart, Source, separator);

                    if (Ending != 0)
                    {
                        Index = Index + 1;
                        SearchStart = Ending + Strings.Len(separator);
                        if (Consecutive == true)
                        {
                            while ((Strings.Mid(Source, SearchStart, Strings.Len(separator)) ?? "") == (separator ?? ""))
                                SearchStart = SearchStart + Strings.Len(separator);
                        }
                    }

                }

            }

            Ending = Strings.Len(Source) + 1;
            Index = Index + 1;

            SplitCountRet = Index;
            return SplitCountRet;

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


            if ((X1 | X2) & !string.IsNullOrEmpty(ComboBox2.Text))
            {

                string pattern = ComboBox2.Text;

                bool Consecutive;
                if (CB_consecutive_separators.Checked)
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
                        SplitValues = (string[])SplitText(source, pattern, Consecutive, KeepSeparator, Before);
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
                        SplitValues = (string[])SplitText(source, pattern, Consecutive, KeepSeparator, Before);

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
                        lengths[i - 1] = Conversions.ToInteger(SplitCount(source, pattern, Consecutive));
                    }
                    int TotalWidth = Conversions.ToInteger(FindMax(lengths));

                    var Values = new string[r, TotalWidth];

                    for (int i = 1, loopTo8 = displayRng.Rows.Count; i <= loopTo8; i++)
                    {
                        string source = Conversions.ToString(displayRng.Cells[i, 1].value);
                        string[] SplitValues;
                        SplitValues = (string[])SplitText(source, pattern, Consecutive, KeepSeparator, Before);
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

        }
        private void CB_separators_finaloutput_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                if (CB_separators_finaloutput.Checked == true)
                {
                    RB_starting_point.Enabled = true;
                    RB_ending_point.Enabled = true;
                    PictureBox4.Enabled = true;
                    PictureBox3.Enabled = true;
                }

                else if (CB_separators_finaloutput.Checked == false)
                {
                    RB_starting_point.Enabled = false;
                    RB_ending_point.Enabled = false;
                    PictureBox4.Enabled = false;
                    PictureBox3.Enabled = false;
                }
                Display();
            }
            catch (Exception ex)
            {

            }
        }

        private void CB_consecutive_separators_CheckedChanged(object sender, EventArgs e)
        {
            // Try
            // Call Display()
            // Catch ex As Exception

            // End Try
        }

        private void Btn_OK_Click(object sender, EventArgs e)
        {

            try
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

                if (X1 == false & X2 == false)
                {
                    MessageBox.Show("Select a Split Option.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    workSheet.Activate();
                    rng.Select();
                    return;
                }

                if (CB_backup.Checked == true)
                {
                    workSheet.Copy(After: workBook.Sheets[workSheet.Name]);
                }

                workSheet.Activate();

                if ((X1 | X2) & !string.IsNullOrEmpty(ComboBox2.Text))
                {


                    string pattern = ComboBox2.Text;

                    bool Consecutive;
                    if (CB_consecutive_separators.Checked)
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
                            string source = Conversions.ToString(rng.Cells[i, 1].Value);
                            Arr[i - 1] = source;
                            string[] SplitValues;
                            SplitValues = (string[])SplitText(source, pattern, Consecutive, KeepSeparator, Before);
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
                            rng2.Columns[j].AutoFit();
                    }

                    else if (X2)
                    {

                        int MaxColumns = 1;
                        for (int i = 1, loopTo5 = r; i <= loopTo5; i++)
                        {
                            string source = Conversions.ToString(rng.Cells[i, 1].Value);
                            string[] SplitValues;
                            SplitValues = (string[])SplitText(source, pattern, Consecutive, KeepSeparator, Before);
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
                            rng2.Columns[j].AutoFit();

                    }

                    Close();

                }
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

        private void Form27_Split_text_bystrings_Load(object sender, EventArgs e)
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

        private void Form27_Split_text_bystrings_KeyDown(object sender, KeyEventArgs e)
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

        private void Form28_Split_text_bypattern_Disposed(object sender, EventArgs e)
        {
            GlobalModule.form_flag = false;
        }

        private void Form28_Split_text_bypattern_Closing(object sender, CancelEventArgs e)
        {
            GlobalModule.form_flag = false;
        }

        private void Form28_Split_text_bypattern_Shown(object sender, EventArgs e)
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