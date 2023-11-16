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

    public partial class Form1
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

        public Form1()
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
        private object IsWithin(Range rng1, Range rng2)
        {

            var excelApp = Globals.ThisAddIn.Application;

            // Use Intersect method to check if range1 is within range2
            var intersectRange = excelApp.Intersect(rng2, rng1);

            // If intersectRange is nothing, then range1 is not within range2
            if (intersectRange is null)
            {
                return false;
            }
            // If the address of intersectRange is same as range1, then range1 is within range2
            else if ((intersectRange.get_Address() ?? "") == (rng2.get_Address() ?? ""))
            {
                return true;
            }

            // Default return value
            return false;

        }
        private string ReplaceNotInRange(string input, string find, string replaceWith)
        {

            // Build the regex pattern to exclude range notation and exclamation mark
            string pattern = string.Format("(?<![!{0}:]){0}(?![:{0}])", Regex.Escape(find));

            // Create a Regex object.
            var reg = new Regex(pattern);

            // Call the Regex.Replace method to replace matching text.
            return reg.Replace(input, replaceWith);
        }
        private object ReplaceReference(string Ref, Range rng, Range rng2, int type)
        {
            object ReplaceReferenceRet = default;

            if (Strings.InStr(1, Ref, "!") > 0)
            {
                ReplaceReferenceRet = Ref;
            }
            else
            {

                Excel.Worksheet activesheet = (Excel.Worksheet)excelApp.ActiveSheet;

                int colNum;
                int rowNum;
                int colNum2;
                int rowNum2;
                string colName;
                string rowName;
                string colName2;
                string rowName2;
                Range expRange;
                int Ext;
                int Ext2;
                string Ref2;
                string Ref3;
                int distance1;
                int distance2;

                distance1 = Conversions.ToInteger(Operators.SubtractObject(rng2.Cells[1, 1].Row, rng.Cells[1, 1].Row));
                distance2 = Conversions.ToInteger(Operators.SubtractObject(rng2.Cells[1, 1].Column, rng.Cells[1, 1].Column));

                expRange = activesheet.get_Range(Ref);

                if (type == 1)
                {
                    colNum = expRange.Column;
                    colName = Strings.Split(Conversions.ToString(activesheet.Cells[1, colNum].Address), "$")[1];
                    Ext = Conversions.ToInteger(Operators.AddObject(Operators.SubtractObject(colNum, rng.Cells[1, 1].Column), 1));
                    Ext2 = rng.Columns.Count - Ext + 1;
                    colNum2 = Conversions.ToInteger(Operators.AddObject(Operators.SubtractObject(rng.Cells[1, 1].Column, 1), Ext2));
                    colName2 = Strings.Split(Conversions.ToString(activesheet.Cells[1, colNum2].Address), "$")[1];
                    Ref2 = Strings.Replace(Ref, colName, colName2);
                    expRange = activesheet.get_Range(Ref2);
                    rowNum = expRange.Row;
                    colNum = expRange.Column;
                    rowNum2 = rowNum + distance1;
                    colNum2 = colNum + distance2;
                    rowName = Strings.Split(Conversions.ToString(activesheet.Cells[rowNum, 1].Address), "$")[2];
                    rowName2 = Strings.Split(Conversions.ToString(activesheet.Cells[rowNum2, 1].Address), "$")[2];
                    colName = Strings.Split(Conversions.ToString(activesheet.Cells[1, colNum].Address), "$")[1];
                    colName2 = Strings.Split(Conversions.ToString(activesheet.Cells[1, colNum2].Address), "$")[1];
                    Ref3 = Strings.Replace(Ref2, rowName, rowName2);
                    Ref3 = Strings.Replace(Ref3, colName, colName2);
                }
                else if (type == 2)
                {
                    rowNum = expRange.Row;
                    rowName = Strings.Split(Conversions.ToString(activesheet.Cells[rowNum, 1].Address), "$")[2];
                    Ext = Conversions.ToInteger(Operators.AddObject(Operators.SubtractObject(rowNum, rng.Cells[1, 1].Row), 1));
                    Ext2 = rng.Rows.Count - Ext + 1;
                    rowNum2 = Conversions.ToInteger(Operators.AddObject(Operators.SubtractObject(rng.Cells[1, 1].Row, 1), Ext2));
                    rowName2 = Strings.Split(Conversions.ToString(activesheet.Cells[rowNum2, 1].Address), "$")[2];
                    Ref2 = Strings.Replace(Ref, rowName, rowName2);
                    expRange = activesheet.get_Range(Ref2);
                    rowNum = expRange.Row;
                    colNum = expRange.Column;
                    rowNum2 = rowNum + distance1;
                    colNum2 = colNum + distance2;
                    rowName = Strings.Split(Conversions.ToString(activesheet.Cells[rowNum, 1].Address), "$")[2];
                    rowName2 = Strings.Split(Conversions.ToString(activesheet.Cells[rowNum2, 1].Address), "$")[2];
                    colName = Strings.Split(Conversions.ToString(activesheet.Cells[1, colNum].Address), "$")[1];
                    colName2 = Strings.Split(Conversions.ToString(activesheet.Cells[1, colNum2].Address), "$")[1];
                    Ref3 = Strings.Replace(Ref2, rowName, rowName2);
                    Ref3 = Strings.Replace(Ref3, colName, colName2);
                }
                else
                {
                    Ref3 = Ref;
                }

                ReplaceReferenceRet = Ref3;
            }

            return ReplaceReferenceRet;
        }
        private object ReplaceRange(string Ref, Range rng, Range rng2, int Type)
        {
            object ReplaceRangeRet = default;

            if (Strings.InStr(1, Ref, "!") > 0)
            {
                ReplaceRangeRet = Ref;
            }
            else
            {
                string Ref1;
                string Ref2;

                string[] R1;
                R1 = Strings.Split(Ref, ":");
                Ref1 = R1[0];
                Ref2 = R1[1];

                Ref1 = Conversions.ToString(ReplaceReference(Ref1, rng, rng2, Type));
                Ref2 = Conversions.ToString(ReplaceReference(Ref2, rng, rng2, Type));

                string NewRef;
                NewRef = Ref1 + ":" + Ref2;

                ReplaceRangeRet = NewRef;
            }

            return ReplaceRangeRet;

        }
        private object ReplaceFormula(string Formula, Range Rng, Range rng2, int Type, Excel.Worksheet sheet1, Excel.Worksheet sheet2)
        {
            object ReplaceFormulaRet = default;

            Excel.Worksheet activesheet = (Excel.Worksheet)excelApp.ActiveSheet;

            string[] Starters = new string[] { "--", "=", "(", ",", " ", "+", "-", "*", "/", "^", ")" };

            var Arr = new string[1];

            int Index;
            Index = -1;

            var Arr1 = default(int[]);

            int Index1;
            Index1 = -1;

            var Refs = new string[1];

            int i;
            int j;

            var loopTo = Strings.Len(Formula);
            for (i = 1; i <= loopTo; i++)
            {
                var loopTo1 = Information.UBound(Starters);
                for (j = Information.LBound(Starters); j <= loopTo1; j++)
                {
                    if ((Strings.Mid(Formula, i, 1) ?? "") == (Starters[j] ?? ""))
                    {
                        Index1 = Index1 + 1;
                        Array.Resize(ref Arr1, Index1 + 1);
                        Arr1[Index1] = i;
                        break;
                    }
                }
            }

            Index1 = Index1 + 1;
            Array.Resize(ref Arr1, Index1 + 1);
            Arr1[Index1] = Strings.Len(Formula) + 1;

            int Start;
            int Ending;
            string Ref;

            var loopTo2 = Information.UBound(Arr1) - 1;
            for (i = Information.LBound(Arr1); i <= loopTo2; i++)
            {
                Index = Index + 1;
                Start = Arr1[i];
                Ending = Arr1[i + 1];
                Ref = Strings.Mid(Formula, Start + 1, Ending - Start - 1);
                Array.Resize(ref Arr, Index + 1);
                Arr[Index] = Ref;
            }

            Index = -1;

            bool C1;
            bool C2;
            bool C3;
            bool C4;
            bool C5;

            var loopTo3 = Information.UBound(Arr);
            for (i = Information.LBound(Arr); i <= loopTo3; i++)
            {

                if (!string.IsNullOrEmpty(Arr[i]))
                {
                    C1 = Strings.Asc(Strings.Mid(Arr[i], Strings.Len(Arr[i]), 1)) >= 48 & Strings.Asc(Strings.Mid(Arr[i], Strings.Len(Arr[i]), 1)) <= 57;
                    C2 = Strings.Asc(Strings.Mid(Arr[i], 1, 1)) >= 65 & Strings.Asc(Strings.Mid(Arr[i], 1, 1)) <= 90;
                    C3 = Strings.Asc(Strings.Mid(Arr[i], 1, 1)) >= 97 & Strings.Asc(Strings.Mid(Arr[i], 1, 1)) <= 122;
                    C4 = Strings.Mid(Arr[i], 1, 1) == "$";
                    C5 = Strings.InStr(1, Arr[i], "!") == 0;

                    if (C1 & (C2 | C3 | C4) & C5)
                    {
                        Index = Index + 1;
                        Array.Resize(ref Refs, Index + 1);
                        Refs[Index] = Arr[i];
                    }
                }
            }

            Range expRange;

            foreach (var currentRef in Refs)
            {
                Ref = currentRef;
                if (Strings.InStr(1, Ref, ":") == 0)
                {
                    expRange = activesheet.get_Range(Ref);
                    if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(IsWithin(Rng, expRange), false, false)))
                    {
                        if ((sheet1.Name ?? "") != (sheet2.Name ?? ""))
                        {
                            string Ref2 = "'" + sheet1.Name + "'" + "!" + Ref;
                            Formula = ReplaceNotInRange(Formula, Ref, Ref2);
                        }
                    }
                    else
                    {
                        string Ref2 = Conversions.ToString(ReplaceReference(Ref, Rng, rng2, Type));
                        Formula = ReplaceNotInRange(Formula, Ref, Ref2);
                    }
                }
                else
                {
                    expRange = activesheet.get_Range(Ref);

                    if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(IsWithin(Rng, expRange), false, false)))
                    {
                        if ((sheet1.Name ?? "") != (sheet2.Name ?? ""))
                        {
                            string[] ex;
                            ex = Strings.Split(Ref, ":");
                            string ex1 = ex[0];
                            string ex2 = ex[1];
                            string Ref2 = "'" + sheet1.Name + "'" + "!" + ex1 + ":" + "'" + sheet1.Name + "'" + "!" + ex2;
                            Formula = ReplaceNotInRange(Formula, Ref, Ref2);
                        }
                    }
                    else
                    {
                        string[] ex;
                        ex = Strings.Split(Ref, ":");
                        string ex1 = ex[0];
                        string ex2 = ex[1];
                        string ex3 = Conversions.ToString(ReplaceReference(ex1, Rng, rng2, Type));
                        string ex4 = Conversions.ToString(ReplaceReference(ex2, Rng, rng2, Type));
                        string Ref2 = ex3 + ":" + ex4;
                        Formula = ReplaceNotInRange(Formula, Ref, Ref2);
                    }

                }
            }

            ReplaceFormulaRet = Formula;
            return ReplaceFormulaRet;

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

                double height;
                double BaseWidth;

                if (displayRng.Rows.Count <= 4)
                {
                    height = panel1.Height / (double)displayRng.Rows.Count;
                }
                else
                {
                    height = 119d / 4d;
                }

                BaseWidth = 260d / 3d;

                double Ordinate = 0d;
                Range CRng;

                var widths = new object[displayRng.Columns.Count];
                for (int j = 1, loopTo = displayRng.Columns.Count; j <= loopTo; j++)
                {

                    CRng = workSheet.get_Range(displayRng.Cells[1, j], displayRng.Cells[displayRng.Rows.Count, j]);
                    widths[j - 1] = Operators.DivideObject(Operators.MultiplyObject(MaxOfColumn(CRng), BaseWidth), 10);

                    for (int i = 1, loopTo1 = displayRng.Rows.Count; i <= loopTo1; i++)
                    {
                        var label = new System.Windows.Forms.Label();
                        label.Text = Conversions.ToString(displayRng.Cells[i, j].Value);
                        label.Location = new System.Drawing.Point((int)Math.Round(Ordinate), (int)Math.Round((i - 1) * height));
                        label.Height = (int)Math.Round(height);
                        label.Width = Conversions.ToInteger(widths[j - 1]);
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
                    Ordinate = Conversions.ToDouble(Operators.AddObject(Ordinate, widths[j - 1]));
                }

                panel1.AutoScroll = true;

                if ((RadioButton1.Checked == true | RadioButton4.Checked == true | RadioButton5.Checked == true) & (RadioButton3.Checked == true | RadioButton2.Checked == true))
                {

                    if (RadioButton3.Checked == true)
                    {
                        Ordinate = 0d;
                        for (int j = 1, loopTo2 = displayRng.Columns.Count; j <= loopTo2; j++)
                        {
                            for (int i = 1, loopTo3 = displayRng.Rows.Count; i <= loopTo3; i++)
                            {
                                var label = new System.Windows.Forms.Label();
                                label.Text = Conversions.ToString(displayRng.Cells[i, displayRng.Columns.Count - j + 1].Value);
                                label.Location = new System.Drawing.Point((int)Math.Round(Ordinate), (int)Math.Round((i - 1) * height));
                                label.Height = (int)Math.Round(height);
                                label.Width = Conversions.ToInteger(widths[displayRng.Columns.Count - j + 1 - 1]);
                                label.BorderStyle = BorderStyle.FixedSingle;
                                label.TextAlign = ContentAlignment.MiddleCenter;

                                if (CheckBox2.Checked == true)
                                {
                                    Range cell = (Range)displayRng.Cells[i, displayRng.Columns.Count - j + 1];
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
                            Ordinate = Conversions.ToDouble(Operators.AddObject(Ordinate, widths[displayRng.Columns.Count - j + 1 - 1]));
                        }

                    }

                    if (RadioButton2.Checked == true)
                    {
                        Ordinate = 0d;
                        for (int j = 1, loopTo4 = displayRng.Columns.Count; j <= loopTo4; j++)
                        {
                            for (int i = 1, loopTo5 = displayRng.Rows.Count; i <= loopTo5; i++)
                            {
                                var label = new System.Windows.Forms.Label();
                                label.Text = Conversions.ToString(displayRng.Cells[displayRng.Rows.Count - i + 1, j].Value);
                                label.Location = new System.Drawing.Point((int)Math.Round(Ordinate), (int)Math.Round((i - 1) * height));
                                label.Height = (int)Math.Round(height);
                                label.Width = Conversions.ToInteger(widths[j - 1]);
                                label.BorderStyle = BorderStyle.FixedSingle;
                                label.TextAlign = ContentAlignment.MiddleCenter;

                                if (CheckBox2.Checked == true)
                                {
                                    Range cell = (Range)displayRng.Cells[i, rng.Columns.Count - j + 1];
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
                            Ordinate = Conversions.ToDouble(Operators.AddObject(Ordinate, widths[j - 1]));
                        }

                    }

                    panel2.AutoScroll = true;

                }
            }


            catch (Exception ex)
            {

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
                while (startRow > 1 && !(workSheet.Cells[startRow - 1, startColumn].Value == null))
                    startRow -= 1;

                // Find the lower boundary
                while (!(workSheet.Cells[endRow + 1, endColumn].Value == null))
                    endRow += 1;

                // Find the left boundary
                while (startColumn > 1 && !(workSheet.Cells[startRow, startColumn - 1].Value == null))
                    startColumn -= 1;

                // Find the right boundary
                while (!(workSheet.Cells[endRow, endColumn + 1].Value == null))
                    endColumn += 1;

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

        private void PictureBox8_Click(object sender, EventArgs e)
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

                if ((workSheet.Name ?? "") != (OpenSheet.Name ?? ""))
                {
                    TextBox1.Text = workSheet.Name + "!" + rng.get_Address();
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

        private void RadioButton1_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                Display();
            }
            catch (Exception ex)
            {

            }
        }

        private void RadioButton2_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                Display();
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

        private void RadioButton4_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                Display();
            }
            catch (Exception ex)
            {

            }
        }

        private void RadioButton3_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                Display();
            }
            catch (Exception ex)
            {

            }
        }

        private void btn_OK_Click(object sender, EventArgs e)
        {

            try
            {

                TextBoxChanged = true;
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

                if (RadioButton9.Checked == false & RadioButton10.Checked == false)
                {
                    MessageBox.Show("Select a Destination Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    workSheet.Activate();
                    rng.Select();
                    return;
                }

                if (RadioButton10.Checked == true & string.IsNullOrEmpty(TextBox2.Text))
                {
                    MessageBox.Show("Select a Destination Cell.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    TextBox2.Focus();
                    return;
                }

                if (RadioButton10.Checked == true & IsValidExcelCellReference(TextBox2.Text) == false)
                {
                    MessageBox.Show("Select a Valid Destination Cell.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    TextBox2.Focus();
                    return;
                }

                if (RadioButton2.Checked == false & RadioButton3.Checked == false)
                {
                    MessageBox.Show("Select a Flip Type.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    workSheet.Activate();
                    rng.Select();
                    return;
                }

                else if (RadioButton1.Checked == false & RadioButton4.Checked == false & RadioButton5.Checked == false)
                {
                    MessageBox.Show("Select a Flip Option.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    workSheet.Activate();
                    rng.Select();
                    return;
                }

                if (CheckBox1.Checked == true)
                {
                    workSheet.Copy(After: workBook.Sheets[workSheet.Name]);
                    workSheet2.Activate();
                }

                if (RadioButton9.Checked == true)
                {
                    rng2 = rng;
                }
                else
                {
                    rng2 = workSheet2.get_Range(rng2.Cells[1, 1], rng2.Cells[rng.Rows.Count, rng.Columns.Count]);
                }

                string rng2Address = rng2.get_Address();

                workSheet2.Activate();

                int i;
                int j;

                if ((RadioButton1.Checked == true | RadioButton4.Checked == true | RadioButton5.Checked == true) & (RadioButton3.Checked == true | RadioButton2.Checked == true))
                {


                    if (Overlap(excelApp, workSheet, workSheet2, rng, rng2) == false)
                    {

                        rng2.ClearFormats();

                        if (RadioButton3.Checked == true)
                        {
                            var loopTo = rng.Rows.Count;
                            for (i = 1; i <= loopTo; i++)
                            {
                                var loopTo1 = rng.Columns.Count;
                                for (j = 1; j <= loopTo1; j++)
                                {
                                    if (RadioButton1.Checked == true)
                                    {
                                        rng2.Cells[i, j].Value = rng.Cells[i, rng.Columns.Count - j + 1].Value;
                                    }

                                    if (RadioButton4.Checked == true)
                                    {
                                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(rng.Cells[i, rng.Columns.Count - j + 1].HasFormula, true, false)))
                                        {
                                            rng2.Cells[i, j].Formula = ReplaceFormula(Conversions.ToString(rng.Cells[i, rng.Columns.Count - j + 1].Formula), rng, rng2, 1, workSheet, workSheet2);
                                        }
                                        else
                                        {
                                            rng2.Cells[i, j].Value = rng.Cells[i, rng.Columns.Count - j + 1].Value;
                                        }
                                    }

                                    if (RadioButton5.Checked == true)
                                    {
                                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(rng.Cells[i, rng.Columns.Count - j + 1].HasFormula, true, false)))
                                        {
                                            rng2.Cells[i, j].Formula = rng.Cells[i, rng.Columns.Count - j + 1].Formula;
                                        }
                                        else
                                        {
                                            rng2.Cells[i, j].Value = rng.Cells[i, rng.Columns.Count - j + 1].Value;
                                        }
                                    }
                                    if (CheckBox2.Checked == true)
                                    {
                                        rng.Cells[i, rng.Columns.Count - j + 1].Copy();
                                        rng2.Cells[i, j].PasteSpecial(XlPasteType.xlPasteFormats);
                                        rng2 = workSheet2.get_Range(rng2Address);

                                        Range sourceCell = (Range)rng.Cells[i, rng.Columns.Count - j + 1];
                                        Range targetCell = (Range)rng2.Cells[i, j];

                                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(sourceCell.Borders[(XlBordersIndex)7].LineStyle, XlLineStyle.xlLineStyleNone, false)))
                                        {
                                            targetCell.Borders[(XlBordersIndex)10].LineStyle = sourceCell.Borders[(XlBordersIndex)7].LineStyle;
                                            targetCell.Borders[(XlBordersIndex)10].Color = sourceCell.Borders[(XlBordersIndex)7].Color;
                                            targetCell.Borders[(XlBordersIndex)10].Weight = sourceCell.Borders[(XlBordersIndex)7].Weight;
                                        }
                                        else
                                        {
                                            targetCell.Borders[(XlBordersIndex)10].LineStyle = XlLineStyle.xlLineStyleNone;
                                        }

                                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(sourceCell.Borders[(XlBordersIndex)10].LineStyle, XlLineStyle.xlLineStyleNone, false)))
                                        {
                                            targetCell.Borders[(XlBordersIndex)7].LineStyle = sourceCell.Borders[(XlBordersIndex)10].LineStyle;
                                            targetCell.Borders[(XlBordersIndex)7].Color = sourceCell.Borders[(XlBordersIndex)10].Color;
                                            targetCell.Borders[(XlBordersIndex)7].Weight = sourceCell.Borders[(XlBordersIndex)10].Weight;
                                        }
                                        else
                                        {
                                            targetCell.Borders[(XlBordersIndex)7].LineStyle = XlLineStyle.xlLineStyleNone;
                                        }

                                    }
                                    excelApp.CutCopyMode = XlCutCopyMode.xlCopy;
                                }
                            }

                        }

                        if (RadioButton2.Checked == true)
                        {

                            var loopTo2 = rng.Rows.Count;
                            for (i = 1; i <= loopTo2; i++)
                            {
                                var loopTo3 = rng.Columns.Count;
                                for (j = 1; j <= loopTo3; j++)
                                {

                                    if (RadioButton1.Checked == true)
                                    {
                                        rng2.Cells[i, j].Value = rng.Cells[rng.Rows.Count - i + 1, j].Value;

                                    }

                                    if (RadioButton4.Checked == true)
                                    {
                                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(rng.Cells[rng.Rows.Count - i + 1, j].HasFormula, true, false)))
                                        {
                                            rng2.Cells[i, j].Formula = ReplaceFormula(Conversions.ToString(rng.Cells[rng.Rows.Count - i + 1, j].Formula), rng, rng2, 2, workSheet, workSheet2);
                                        }
                                        else
                                        {
                                            rng2.Cells[i, j].Value = rng.Cells[rng.Rows.Count - i + 1, j].Value;
                                        }
                                    }

                                    if (RadioButton5.Checked == true)
                                    {
                                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(rng.Cells[rng.Rows.Count - i + 1, j].HasFormula, true, false)))
                                        {
                                            rng2.Cells[i, j].Formula = rng.Cells[rng.Rows.Count - i + 1, j].Formula;
                                        }
                                        else
                                        {
                                            rng2.Cells[i, j].Value = rng.Cells[rng.Rows.Count - i + 1, j].Value;
                                        }
                                    }

                                    if (CheckBox2.Checked == true)
                                    {
                                        rng.Cells[rng.Rows.Count - i + 1, j].Copy();
                                        rng2.Cells[i, j].PasteSpecial(XlPasteType.xlPasteFormats);
                                        rng2 = workSheet2.get_Range(rng2Address);

                                        Range sourceCell = (Range)rng.Cells[rng.Rows.Count - i + 1, j];
                                        Range targetCell = (Range)rng2.Cells[i, j];

                                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(sourceCell.Borders[(XlBordersIndex)8].LineStyle, XlLineStyle.xlLineStyleNone, false)))
                                        {
                                            targetCell.Borders[(XlBordersIndex)9].LineStyle = sourceCell.Borders[(XlBordersIndex)8].LineStyle;
                                            targetCell.Borders[(XlBordersIndex)9].Color = sourceCell.Borders[(XlBordersIndex)8].Color;
                                            targetCell.Borders[(XlBordersIndex)9].Weight = sourceCell.Borders[(XlBordersIndex)8].Weight;
                                        }
                                        else
                                        {
                                            targetCell.Borders[(XlBordersIndex)9].LineStyle = XlLineStyle.xlLineStyleNone;
                                        }

                                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(sourceCell.Borders[(XlBordersIndex)9].LineStyle, XlLineStyle.xlLineStyleNone, false)))
                                        {
                                            targetCell.Borders[(XlBordersIndex)8].LineStyle = sourceCell.Borders[(XlBordersIndex)9].LineStyle;
                                            targetCell.Borders[(XlBordersIndex)8].Color = sourceCell.Borders[(XlBordersIndex)9].Color;
                                            targetCell.Borders[(XlBordersIndex)8].Weight = sourceCell.Borders[(XlBordersIndex)9].Weight;
                                        }
                                        else
                                        {
                                            targetCell.Borders[(XlBordersIndex)8].LineStyle = XlLineStyle.xlLineStyleNone;
                                        }
                                    }
                                    excelApp.CutCopyMode = XlCutCopyMode.xlCopy;
                                }
                            }

                        }
                    }

                    else
                    {

                        var Arr = new object[rng.Rows.Count, rng.Columns.Count];

                        var loopTo4 = Information.UBound(Arr, 1);
                        for (i = Information.LBound(Arr, 1); i <= loopTo4; i++)
                        {
                            var loopTo5 = Information.UBound(Arr, 2);
                            for (j = Information.LBound(Arr, 2); j <= loopTo5; j++)
                                Arr[i, j] = rng.Cells[i + 1, j + 1].Value;
                        }

                        var FontNames = new string[rng.Rows.Count, rng.Columns.Count];
                        var HasFormulas = new bool[rng.Rows.Count, rng.Columns.Count];
                        var Formulas = new string[rng.Rows.Count, rng.Columns.Count];
                        var FontSizes = new float[rng.Rows.Count, rng.Columns.Count];

                        var FontBolds = new bool[rng.Rows.Count, rng.Columns.Count];
                        var Fontitalics = new bool[rng.Rows.Count, rng.Columns.Count];
                        var Red1s = new int[rng.Rows.Count, rng.Columns.Count];
                        var Green1s = new int[rng.Rows.Count, rng.Columns.Count];
                        var Blue1s = new int[rng.Rows.Count, rng.Columns.Count];
                        var Red2s = new int[rng.Rows.Count, rng.Columns.Count];
                        var Green2s = new int[rng.Rows.Count, rng.Columns.Count];
                        var Blue2s = new int[rng.Rows.Count, rng.Columns.Count];

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

                            var loopTo6 = Information.UBound(FontSizes, 1);
                            for (i = Information.LBound(FontSizes, 1); i <= loopTo6; i++)
                            {
                                var loopTo7 = Information.UBound(FontSizes, 2);
                                for (j = Information.LBound(FontSizes, 2); j <= loopTo7; j++)
                                {

                                    Range cell = (Range)rng.Cells[i + 1, j + 1];
                                    if (Conversions.ToBoolean(cell.HasFormula))
                                    {
                                        HasFormulas[i, j] = true;
                                    }
                                    else
                                    {
                                        HasFormulas[i, j] = false;
                                    }

                                    Formulas[i, j] = Conversions.ToString(cell.Formula);
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

                                    FontBolds[i, j] = Conversions.ToBoolean(cell.Font.Bold);
                                    Fontitalics[i, j] = Conversions.ToBoolean(cell.Font.Italic);

                                    if (cell.Interior.Color is DBNull)
                                    {
                                        Red1s[i, j] = 255;
                                        Green1s[i, j] = 255;
                                        Blue1s[i, j] = 255;
                                    }
                                    else
                                    {
                                        long colorValue1 = Conversions.ToLong(cell.Interior.Color);
                                        int red1 = (int)(colorValue1 % 256L);
                                        int green1 = (int)(colorValue1 / 256L % 256L);
                                        int blue1 = (int)(colorValue1 / 256L / 256L % 256L);
                                        Red1s[i, j] = red1;
                                        Green1s[i, j] = green1;
                                        Blue1s[i, j] = blue1;
                                    }

                                    if (cell.Font.Color is DBNull)
                                    {
                                        Red2s[i, j] = 0;
                                        Green2s[i, j] = 0;
                                        Blue2s[i, j] = 0;
                                    }
                                    else
                                    {
                                        long colorValue2 = Conversions.ToLong(cell.Font.Color);
                                        int red2 = (int)(colorValue2 % 256L);
                                        int green2 = (int)(colorValue2 / 256L % 256L);
                                        int blue2 = (int)(colorValue2 / 256L / 256L % 256L);
                                        Red2s[i, j] = red2;
                                        Green2s[i, j] = green2;
                                        Blue2s[i, j] = blue2;
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

                        if (RadioButton3.Checked == true)
                        {

                            var loopTo8 = rng.Rows.Count;
                            for (i = 1; i <= loopTo8; i++)
                            {
                                var loopTo9 = rng.Columns.Count;
                                for (j = 1; j <= loopTo9; j++)
                                {

                                    if (RadioButton1.Checked == true)
                                    {
                                        rng2.Cells[i, j].Value = Arr[i - 1, rng.Columns.Count - j + 1 - 1];
                                    }

                                    if (RadioButton4.Checked == true)
                                    {
                                        if (HasFormulas[i - 1, rng.Columns.Count - j + 1 - 1] == true)
                                        {
                                            rng2.Cells[i, j].Formula = ReplaceFormula(Formulas[i - 1, rng.Columns.Count - j + 1 - 1], rng, rng2, 1, workSheet, workSheet2);
                                        }
                                        else
                                        {
                                            rng2.Cells[i, j] = Arr[i - 1, rng.Columns.Count - j + 1 - 1];
                                        }
                                    }

                                    if (RadioButton5.Checked == true)
                                    {
                                        if (HasFormulas[i - 1, rng.Columns.Count - j + 1 - 1] == true)
                                        {
                                            rng2.Cells[i, j].Formula = Formulas[i - 1, rng.Columns.Count - j + 1 - 1];
                                        }
                                        else
                                        {
                                            rng2.Cells[i, j] = Arr[i - 1, rng.Columns.Count - j + 1 - 1];
                                        }
                                    }

                                    if (CheckBox2.Checked == true)
                                    {
                                        int x = i - 1;
                                        int y = rng.Columns.Count - j + 1 - 1;

                                        rng2.Cells[i, j].Font.Name = FontNames[x, y];
                                        rng2.Cells[i, j].Font.Size = (object)FontSizes[x, y];

                                        if (FontBolds[x, y])
                                            rng2.Cells[i, j].Font.Bold = (object)true;
                                        if (Fontitalics[x, y])
                                            rng2.Cells[i, j].Font.Italic = (object)true;


                                        rng2.Cells[i, j].Interior.Color = (object)Color.FromArgb(Red1s[x, y], Green1s[x, y], Blue1s[x, y]);

                                        rng2.Cells[i, j].Font.Color = (object)Color.FromArgb(Red2s[x, y], Green2s[x, y], Blue2s[x, y]);

                                        Range targetCell = (Range)rng2.Cells[i, j];

                                        if (Borders7[x, y] == true)
                                        {
                                            targetCell.Borders[(XlBordersIndex)10].LineStyle = Borders7L[x, y];
                                            targetCell.Borders[(XlBordersIndex)10].Color = Borders7C[x, y];
                                            targetCell.Borders[(XlBordersIndex)10].Weight = Borders7W[x, y];
                                        }
                                        else
                                        {
                                            targetCell.Borders[(XlBordersIndex)10].LineStyle = XlLineStyle.xlLineStyleNone;
                                        }

                                        if (Borders8[x, y] == true)
                                        {
                                            targetCell.Borders[(XlBordersIndex)8].LineStyle = Borders8L[x, y];
                                            targetCell.Borders[(XlBordersIndex)8].Color = Borders8C[x, y];
                                            targetCell.Borders[(XlBordersIndex)8].Weight = Borders8W[x, y];
                                        }
                                        else
                                        {
                                            targetCell.Borders[(XlBordersIndex)8].LineStyle = XlLineStyle.xlLineStyleNone;
                                        }

                                        if (Borders9[x, y] == true)
                                        {
                                            targetCell.Borders[(XlBordersIndex)9].LineStyle = Borders9L[x, y];
                                            targetCell.Borders[(XlBordersIndex)9].Color = Borders9C[x, y];
                                            targetCell.Borders[(XlBordersIndex)9].Weight = Borders9W[x, y];
                                        }
                                        else
                                        {
                                            targetCell.Borders[(XlBordersIndex)9].LineStyle = XlLineStyle.xlLineStyleNone;
                                        }

                                        if (Borders10[x, y] == true)
                                        {
                                            targetCell.Borders[(XlBordersIndex)7].LineStyle = Borders10L[x, y];
                                            targetCell.Borders[(XlBordersIndex)7].Color = Borders10C[x, y];
                                            targetCell.Borders[(XlBordersIndex)7].Weight = Borders10W[x, y];
                                        }
                                        else
                                        {
                                            targetCell.Borders[(XlBordersIndex)7].LineStyle = XlLineStyle.xlLineStyleNone;
                                        }

                                    }

                                }
                            }

                        }

                        if (RadioButton2.Checked == true)
                        {

                            var loopTo10 = rng.Rows.Count;
                            for (i = 1; i <= loopTo10; i++)
                            {
                                var loopTo11 = rng.Columns.Count;
                                for (j = 1; j <= loopTo11; j++)
                                {

                                    if (RadioButton1.Checked == true)
                                    {
                                        rng2.Cells[i, j].Value = Arr[rng.Rows.Count - i + 1 - 1, j - 1];
                                    }

                                    if (RadioButton4.Checked == true)
                                    {
                                        if (HasFormulas[rng.Rows.Count - i + 1 - 1, j - 1] == true)
                                        {
                                            rng2.Cells[i, j].Formula = ReplaceFormula(Formulas[rng.Rows.Count - i + 1 - 1, j - 1], rng, rng2, 2, workSheet, workSheet2);
                                        }
                                        else
                                        {
                                            rng2.Cells[i, j] = Arr[rng.Rows.Count - i + 1 - 1, j - 1];
                                        }
                                    }

                                    if (RadioButton5.Checked == true)
                                    {
                                        if (HasFormulas[rng.Rows.Count - i + 1 - 1, j - 1] == true)
                                        {
                                            rng2.Cells[i, j].Formula = Formulas[rng.Rows.Count - i + 1 - 1, j - 1];
                                        }
                                        else
                                        {
                                            rng2.Cells[i, j] = Arr[rng.Rows.Count - i + 1 - 1, j - 1];
                                        }
                                    }

                                    if (CheckBox2.Checked == true)
                                    {
                                        int x = rng.Rows.Count - i + 1 - 1;
                                        int y = j - 1;

                                        var fontStyle = FontStyle.Regular;

                                        if (FontBolds[x, y])
                                            fontStyle = fontStyle | FontStyle.Bold;
                                        if (Fontitalics[x, y])
                                            fontStyle = fontStyle | FontStyle.Italic;

                                        rng2.Cells[i, j].Font.Name = FontNames[x, y];
                                        rng2.Cells[i, j].Font.Size = (object)FontSizes[x, y];

                                        if (FontBolds[x, y])
                                            rng2.Cells[i, j].Font.Bold = (object)true;
                                        if (Fontitalics[x, y])
                                            rng2.Cells[i, j].Font.Italic = (object)true;

                                        rng2.Cells[i, j].Interior.Color = (object)Color.FromArgb(Red1s[x, y], Green1s[x, y], Blue1s[x, y]);
                                        rng2.Cells[i, j].Font.Color = (object)Color.FromArgb(Red2s[x, y], Green2s[x, y], Blue2s[x, y]);

                                        Range targetCell = (Range)rng2.Cells[i, j];

                                        if (Borders7[x, y] == true)
                                        {
                                            targetCell.Borders[(XlBordersIndex)7].LineStyle = Borders7L[x, y];
                                            targetCell.Borders[(XlBordersIndex)7].Color = Borders7C[x, y];
                                            targetCell.Borders[(XlBordersIndex)7].Weight = Borders7W[x, y];
                                        }
                                        else
                                        {
                                            targetCell.Borders[(XlBordersIndex)7].LineStyle = XlLineStyle.xlLineStyleNone;
                                        }

                                        if (Borders9[x, y] == true)
                                        {
                                            targetCell.Borders[(XlBordersIndex)8].LineStyle = Borders9L[x, y];
                                            targetCell.Borders[(XlBordersIndex)8].Color = Borders9C[x, y];
                                            targetCell.Borders[(XlBordersIndex)8].Weight = Borders9W[x, y];
                                        }
                                        else
                                        {
                                            targetCell.Borders[(XlBordersIndex)8].LineStyle = XlLineStyle.xlLineStyleNone;
                                        }

                                        if (Borders8[x, y] == true)
                                        {
                                            targetCell.Borders[(XlBordersIndex)9].LineStyle = Borders8L[x, y];
                                            targetCell.Borders[(XlBordersIndex)9].Color = Borders8C[x, y];
                                            targetCell.Borders[(XlBordersIndex)9].Weight = Borders8W[x, y];
                                        }
                                        else
                                        {
                                            targetCell.Borders[(XlBordersIndex)9].LineStyle = XlLineStyle.xlLineStyleNone;
                                        }

                                        if (Borders10[x, y] == true)
                                        {
                                            targetCell.Borders[(XlBordersIndex)10].LineStyle = Borders10L[x, y];
                                            targetCell.Borders[(XlBordersIndex)10].Color = Borders10C[x, y];
                                            targetCell.Borders[(XlBordersIndex)10].Weight = Borders10W[x, y];
                                        }
                                        else
                                        {
                                            targetCell.Borders[(XlBordersIndex)10].LineStyle = XlLineStyle.xlLineStyleNone;
                                        }

                                    }

                                }
                            }

                        }
                    }

                    rng2.Select();

                    var loopTo12 = rng2.Columns.Count;
                    for (j = 1; j <= loopTo12; j++)
                        rng2.Columns[j].Autofit();

                    TextBoxChanged = false;

                    Close();

                }
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

        private void PictureBox9_Click(object sender, EventArgs e)
        {

            try
            {
                FocusedTextBox = 2;
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

                if ((workSheet2.Name ?? "") != (OpenSheet.Name ?? ""))
                {
                    TextBox2.Text = workSheet2.Name + "!" + rng2.get_Address();
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

        private void TextBox2_TextChanged(object sender, EventArgs e)
        {

            try
            {
                excelApp = Globals.ThisAddIn.Application;
                workBook = excelApp.ActiveWorkbook;
                workSheet2 = (Excel.Worksheet)workBook.ActiveSheet;

                string[] rng2Array = Strings.Split(TextBox2.Text, "!");
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

        private void Form1_Load(object sender, EventArgs e)
        {

            try
            {

                excelApp = Globals.ThisAddIn.Application;

                excelApp.SheetSelectionChange += excelApp_SheetSelectionChange;

                KeyPreview = true;

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
                    else if (FocusedTextBox == 2)
                    {
                        workSheet2 = (Excel.Worksheet)workBook.ActiveSheet;
                        if ((workSheet2.Name ?? "") != (OpenSheet.Name ?? ""))
                        {
                            TextBox2.Text = workSheet2.Name + "!" + selectedRange.get_Address();
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

        private void PictureBox10_GotFocus(object sender, EventArgs e)
        {
            try
            {
                FocusedTextBox = 2;
            }
            catch (Exception ex)
            {

            }
        }

        private void PictureBox9_GotFocus(object sender, EventArgs e)
        {
            try
            {
                FocusedTextBox = 2;
            }

            catch (Exception ex)
            {

            }
        }


        private void btn_OK_Enter(object sender, EventArgs e)
        {

            try
            {

                btn_OK.BackColor = Color.FromArgb(65, 105, 225);
                btn_OK.ForeColor = Color.FromArgb(255, 255, 255);
            }
            catch (Exception ex)
            {

            }
        }

        private void btn_OK_MouseLeave(object sender, EventArgs e)
        {

            try
            {

                btn_OK.BackColor = Color.FromArgb(255, 255, 255);
                btn_OK.ForeColor = Color.FromArgb(70, 70, 70);
            }
            catch (Exception ex)
            {

            }

        }

        private void btn_cancel_MouseEnter(object sender, EventArgs e)
        {

            try
            {
                btn_cancel.BackColor = Color.FromArgb(65, 105, 225);
                btn_cancel.ForeColor = Color.FromArgb(255, 255, 255);
            }
            catch (Exception ex)
            {

            }
        }

        private void btn_cancel_MouseLeave(object sender, EventArgs e)
        {

            try
            {
                btn_cancel.BackColor = Color.FromArgb(255, 255, 255);
                btn_cancel.ForeColor = Color.FromArgb(70, 70, 70);
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


        private void PictureBox3_GotFocus(object sender, EventArgs e)
        {
            try
            {
                FocusedTextBox = 0;
            }
            catch (Exception ex)
            {

            }
        }

        private void PictureBox6_GotFocus(object sender, EventArgs e)
        {
            try
            {
                FocusedTextBox = 0;
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


        private void PictureBox2_GotFocus(object sender, EventArgs e)
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

        private void Form1_Closing(object sender, CancelEventArgs e)
        {
            try
            {

                GlobalModule.form_flag = false;
            }

            catch (Exception ex)
            {

            }
        }

        private void Form1_Shown(object sender, EventArgs e)
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

        private void Form1_Disposed(object sender, EventArgs e)
        {

            try
            {

                GlobalModule.form_flag = false;
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
                    TextBox2.Enabled = true;
                    TextBox2.Focus();
                    PictureBox9.Enabled = true;
                }
                else
                {
                    Label3.Enabled = false;
                    TextBox2.Clear();
                    TextBox2.Enabled = false;
                    PictureBox9.Enabled = false;
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

        private void Form1_KeyDown(object sender, KeyEventArgs e)
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