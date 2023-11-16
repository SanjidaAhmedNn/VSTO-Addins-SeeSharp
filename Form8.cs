using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
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

    public partial class Form8
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

        public Form8()
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

            // This function takes two ranges as the inputs and checks whether they intersect or not.
            // It will be used to check whether the input range and the output range intersect or not. If they don't intersect, we can directly copy the values and formats from the input range to the output range.
            // But if they do intersect, then the process becomes a bit complex. First we have to copy everything from the input range to a number of arrays, then we will copy them from the arrays to the output range.

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

            // This function takes a string as the input and checks whether it's a valid cell reference or not.
            // It will be used when the user presses the OK button. First it'll check whether all the cell references put in the corresponding text boxes are valid cell references or not.
            // If is, then we will continue the next procedures. Otherwise, we'll exit with an error message asking the user to input a valid cell reference in the required text box.

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
        public int FindMinValue(int[] arr)
        {

            // This function finds the minimum value of an array.

            int min = arr[0];

            foreach (var num in arr)
            {
                if (num < min)
                {
                    min = num;
                }
            }

            return min;

        }
        private object SearchAlongRow(Range Rng, int r, int C)
        {
            object SearchAlongRowRet = default;

            // This is a very important function for this class.
            // It takes a specific co-ordinate of a cell within a range as the input, and finds out the number of adjacent cells with the same value, along the row.
            // It will be used while we attempt to merge similar values row-wise.

            int i = 1;

            bool search = true;

            Type Type1;
            Type Type2;

            while (search == true)
            {

                if (Rng.Cells[r, C + i].Value is null)
                {
                    Type1 = typeof(string);
                }
                else
                {
                    Type1 = Rng.Cells[r, C + i].Value.GetType();
                }

                if (Rng.Cells[r, C].Value is null)
                {
                    Type2 = typeof(string);
                }
                else
                {
                    Type2 = Rng.Cells[r, C].Value.GetType();
                }

                if (Type1.Equals(Type2))
                {
                    if (Conversions.ToBoolean(Operators.AndObject(Operators.AndObject(Operators.AndObject(Operators.ConditionalCompareObjectEqual(Rng.Cells[r, C + i].Value, Rng.Cells[r, C].value, false), C + i <= Rng.Columns.Count), Operators.ConditionalCompareObjectEqual(Rng.Cells[r, C].MergeCells, false, false)), Operators.ConditionalCompareObjectEqual(Rng.Cells[r, C + i].MergeCells, false, false))))
                    {
                        i = i + 1;
                        search = true;
                    }
                    else
                    {
                        search = false;
                    }
                }
                else
                {
                    search = false;
                }

            }

            SearchAlongRowRet = i;
            return SearchAlongRowRet;

        }
        private object SearchAlongColumn(Range Rng, int r, int C)
        {
            object SearchAlongColumnRet = default;

            // This is also a very important function for this class.
            // It takes a specific co-ordinate of a cell within a range as the input, and finds out the number of adjacent cells with the same value, along the column.
            // It will be used while we attempt to merge similar values column-wise.

            int i = 1;

            bool search = true;

            Type Type1;
            Type Type2;

            while (search == true)
            {

                if (Rng.Cells[r + i, C].Value is null)
                {
                    Type1 = typeof(string);
                }
                else
                {
                    Type1 = Rng.Cells[r + i, C].Value.GetType();
                }

                if (Rng.Cells[r, C].Value is null)
                {
                    Type2 = typeof(string);
                }
                else
                {
                    Type2 = Rng.Cells[r, C].Value.GetType();
                }

                if (Type1.Equals(Type2))
                {
                    if (Conversions.ToBoolean(Operators.AndObject(Operators.AndObject(Operators.AndObject(Operators.ConditionalCompareObjectEqual(Rng.Cells[r + i, C].value, Rng.Cells[r, C].value, false), r + i <= Rng.Rows.Count), Operators.ConditionalCompareObjectEqual(Rng.Cells[r, C].MergeCells, false, false)), Operators.ConditionalCompareObjectEqual(Rng.Cells[r + i, C].MergeCells, false, false))))
                    {
                        i = i + 1;
                        search = true;
                    }
                    else
                    {
                        search = false;
                    }
                }
                else
                {
                    search = false;
                }

            }

            SearchAlongColumnRet = i;
            return SearchAlongColumnRet;

        }

        private object SearchDiagonally(object Rng, object r, object c)
        {
            object SearchDiagonallyRet = default;

            // This is another very important function for this class.
            // It takes a specific co-ordinate of a cell within a range as the input, and finds out the highest number of adjacent cells with the same value, both row-wise and column-wise.
            // It will be used while we attempt to merge similar values both row-wise and column-wise.

            int rowEqual = Conversions.ToInteger(SearchAlongRow((Range)Rng, Conversions.ToInteger(r), Conversions.ToInteger(c)));

            Excel.Worksheet activesheet = (Excel.Worksheet)excelApp.ActiveSheet;

            Range Rng2;
            Rng2 = activesheet.get_Range(Rng.Cells((object)1, (object)1), Rng.Cells((object)1, rowEqual));

            var Output = new int[2];
            Output[0] = 1;
            Output[1] = rowEqual;

            int TotalCells = Rng2.Cells.Count;

            int j;

            j = 0;

            int min = Conversions.ToInteger(Rng.Rows.Count);

            while (Operators.AndObject(Operators.ConditionalCompareObjectGreater(SearchAlongColumn((Range)Rng, Conversions.ToInteger(r), Conversions.ToInteger(Operators.AddObject(c, j))), 1, false), j + 1 <= rowEqual))
            {
                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectLessEqual(SearchAlongColumn((Range)Rng, Conversions.ToInteger(r), Conversions.ToInteger(Operators.AddObject(c, j))), min, false)))
                {
                    min = Conversions.ToInteger(SearchAlongColumn((Range)Rng, Conversions.ToInteger(r), Conversions.ToInteger(Operators.AddObject(c, j))));
                }
                if (activesheet.get_Range(Rng.Cells((object)1, (object)1), Rng.Cells(min, (object)(j + 1))).Cells.Count >= TotalCells)
                {
                    Output[0] = min;
                    Output[1] = j + 1;
                    TotalCells = Rng2.Cells.Count;
                }
                j = j + 1;

            }

            SearchDiagonallyRet = Output;
            return SearchDiagonallyRet;

        }
        private object CrossCheck(Excel.Application excelApp, Range rng1, Range rng2)
        {

            // This function takes two ranges (within the same worksheet) as the inputs and checkes whether they overlap or not.
            // It will be used while we attempt to merge smiliar values both row-wise and column-wise.
            // There may be cases where there are multiple ranges within the input range with similar values that intersect, and this function will be used to sort this out.

            var intersectRange = excelApp.Intersect(rng1, rng2);

            if (intersectRange is null)
            {
                return false;
            }
            else
            {
                return true;
            }

        }

        private object RemoveCrossings(object excelApp, object Arr)
        {
            object RemoveCrossingsRet = default;

            // This function takes all the possible ranges within the input range that contain similar values.
            // Then from all the ranges that overlap, it will keep only the largest range and remove all the other ranges.

            Excel.Worksheet activesheet = (Excel.Worksheet)excelApp.ActiveSheet;
            Range Rng1;
            Range Rng2;
            int Count1;
            int Count2;
            for (int i = Information.LBound((Array)Arr, 1), loopTo = Information.UBound((Array)Arr, 1); i <= loopTo; i++)
            {
                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectGreater(Arr((object)i, (object)0), 0, false)))
                {
                    Rng1 = activesheet.get_Range("A1");
                    Rng1 = activesheet.get_Range(Rng1.Cells[Arr((object)i, (object)0), Arr((object)i, (object)1)], Rng1.Cells[Arr((object)i, (object)2), Arr((object)i, (object)3)]);

                    for (int j = Information.LBound((Array)Arr, 1), loopTo1 = Information.UBound((Array)Arr, 1); j <= loopTo1; j++)
                    {
                        if (i != j)
                        {
                            Rng2 = activesheet.get_Range("A1");
                            if (Conversions.ToBoolean(Operators.ConditionalCompareObjectGreater(Arr((object)j, (object)0), 0, false)))
                            {
                                Rng2 = activesheet.get_Range(Rng2.Cells[Arr((object)j, (object)0), Arr((object)j, (object)1)], Rng2.Cells[Arr((object)j, (object)2), Arr((object)j, (object)3)]);

                                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(CrossCheck((Excel.Application)excelApp, Rng1, Rng2), true, false)))
                                {

                                    Count1 = Rng1.Cells.Count;
                                    Count2 = Rng2.Cells.Count;

                                    if (Count1 < Count2)
                                    {
                                        Arr((object)i, (object)0) = (object)0;
                                        break;
                                    }
                                    else if (Count1 == Count2)
                                    {
                                        if (Rng1.Rows.Count == 1 | Rng1.Columns.Count == 1)
                                        {
                                            Arr((object)i, (object)0) = (object)0;
                                            break;
                                        }
                                        else
                                        {
                                            Arr((object)j, (object)0) = (object)0;
                                        }
                                    }
                                    else
                                    {
                                        Arr((object)j, (object)0) = (object)0;
                                    }
                                }
                            }
                        }
                    }
                }
            }

            RemoveCrossingsRet = Arr;
            return RemoveCrossingsRet;

        }
        private object IsWithinRange(int r, int c, Range Rng)
        {
            object IsWithinRangeRet = default;

            // This function takes a co-ordinate of a cell as the input, and checks whether it's located within a given range or not.
            // This will be used while we attempt to merge same values both row-wise and column-wise.

            if (Conversions.ToBoolean(Operators.AndObject(Operators.AndObject(Operators.AndObject(Operators.ConditionalCompareObjectGreaterEqual(r, Rng.Cells[1, 1].Row, false), Operators.ConditionalCompareObjectLessEqual(r, Rng.Cells[Rng.Rows.Count, 1].Row, false)), Operators.ConditionalCompareObjectGreaterEqual(c, Rng.Cells[1, 1].Column, false)), Operators.ConditionalCompareObjectLessEqual(r, Rng.Cells[1, Rng.Columns.Count].Column, false))))
            {

                IsWithinRangeRet = true;
            }
            else
            {
                IsWithinRangeRet = false;
            }

            return IsWithinRangeRet;

        }
        private void Display()
        {

            try
            {

                CustomPanel1.Controls.Clear();
                CustomPanel2.Controls.Clear();

                excelApp = Globals.ThisAddIn.Application;
                workBook = excelApp.ActiveWorkbook;
                workSheet = (Excel.Worksheet)workBook.ActiveSheet;


                var Rng = workSheet.get_Range(TextBox1.Text);
                Range displayRng;
                Rng.Select();

                if (Rng.Rows.Count > 50)
                {
                    displayRng = workSheet.get_Range(Rng.Cells[1, 1], Rng.Cells[50, Rng.Columns.Count]);
                }
                else
                {
                    displayRng = workSheet.get_Range(Rng.Cells[1, 1], Rng.Cells[Rng.Rows.Count, Rng.Columns.Count]);
                }

                int r = displayRng.Rows.Count;
                int C = displayRng.Columns.Count;

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

                if (C <= 4)
                {
                    width = (float)(CustomPanel1.Width / (double)C);
                }
                else
                {
                    width = (float)(CustomPanel1.Width / 4d);
                }

                int i;
                int j;

                var loopTo = r;
                for (i = 1; i <= loopTo; i++)
                {
                    var loopTo1 = C;
                    for (j = 1; j <= loopTo1; j++)
                    {
                        var label = new System.Windows.Forms.Label();
                        label.Text = Conversions.ToString(displayRng.Cells[i, j].Value);
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
                        CustomPanel1.Controls.Add(label);
                    }
                }

                CustomPanel1.AutoScroll = true;

                if (RadioButton1.Checked == true | RadioButton2.Checked == true | RadioButton3.Checked == true)
                {

                    if (RadioButton1.Checked == true)
                    {
                        var loopTo2 = r;
                        for (i = 1; i <= loopTo2; i++)
                        {
                            var loopTo3 = C;
                            for (j = 1; j <= loopTo3; j++)
                            {
                                int rowEqual = Conversions.ToInteger(SearchAlongRow(displayRng, i, j));
                                float newWidth = width * rowEqual;
                                var label = new System.Windows.Forms.Label();
                                label.Text = Conversions.ToString(displayRng.Cells[i, j].Value);
                                label.Location = new System.Drawing.Point((int)Math.Round((j - 1) * width), (int)Math.Round((i - 1) * height));
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

                                j = j + rowEqual - 1;

                                CustomPanel2.Controls.Add(label);
                            }
                        }
                    }

                    else if (RadioButton2.Checked == true == true)
                    {

                        var loopTo4 = C;
                        for (j = 1; j <= loopTo4; j++)
                        {
                            var loopTo5 = r;
                            for (i = 1; i <= loopTo5; i++)
                            {
                                int columnEqual = Conversions.ToInteger(SearchAlongColumn(displayRng, i, j));
                                float newHeight = height * columnEqual;
                                var label = new System.Windows.Forms.Label();
                                label.Text = Conversions.ToString(displayRng.Cells[i, j].Value);
                                label.Location = new System.Drawing.Point((int)Math.Round((j - 1) * width), (int)Math.Round((i - 1) * height));
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
                                i = i + columnEqual - 1;
                                CustomPanel2.Controls.Add(label);
                            }
                        }
                    }

                    else if (RadioButton3.Checked == true)
                    {

                        Excel.Worksheet activesheet = (Excel.Worksheet)excelApp.ActiveSheet;

                        var Arr = new int[(r * C), 4];
                        var loopTo6 = Information.UBound(Arr, 1);
                        for (i = Information.LBound(Arr, 1); i <= loopTo6; i++)
                            Arr[i, 0] = 0;

                        int Count = 0;

                        var loopTo7 = r;
                        for (i = 1; i <= loopTo7; i++)
                        {
                            var loopTo8 = C;
                            for (j = 1; j <= loopTo8; j++)
                            {

                                int rowEqual = Conversions.ToInteger(SearchDiagonally(displayRng, i, j)((object)0));
                                int columnEqual = Conversions.ToInteger(SearchDiagonally(displayRng, i, j)((object)1));

                                Arr[Count, 0] = i;
                                Arr[Count, 1] = j;
                                Arr[Count, 2] = i + rowEqual - 1;
                                Arr[Count, 3] = j + columnEqual - 1;

                                Count = Count + 1;
                            }
                        }

                        Arr = (int[,])RemoveCrossings(excelApp, Arr);

                        var loopTo9 = r;
                        for (i = 1; i <= loopTo9; i++)
                        {
                            var loopTo10 = C;
                            for (j = 1; j <= loopTo10; j++)
                            {

                                Range MRng;
                                MRng = activesheet.get_Range(displayRng.Cells[i, j].Address);

                                for (int m = Information.LBound(Arr, 1), loopTo11 = Information.UBound(Arr, 1); m <= loopTo11; m++)
                                {

                                    if (Arr[m, 0] > 0)
                                    {

                                        var Rng1 = activesheet.get_Range("A1");
                                        Rng1 = activesheet.get_Range(Rng1.Cells[Arr[m, 0], Arr[m, 1]], Rng1.Cells[Arr[m, 2], Arr[m, 3]]);

                                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(IsWithinRange(i, j, Rng1), true, false)))
                                        {
                                            MRng = Rng1;
                                            Arr[m, 0] = 0;
                                            break;
                                        }

                                    }

                                }

                                float newWidth = width * MRng.Columns.Count;
                                float newHeight = height * MRng.Rows.Count;

                                var label = new System.Windows.Forms.Label();
                                label.Text = Conversions.ToString(displayRng.Cells[i, j].Value);
                                label.Location = new System.Drawing.Point((int)Math.Round((j - 1) * width), (int)Math.Round((i - 1) * height));
                                label.Height = (int)Math.Round(newHeight);
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

                    CustomPanel2.AutoScroll = true;

                }
            }

            catch (Exception ex)
            {

            }

        }

        private void Form8_Load(object sender, EventArgs e)
        {

            try
            {

                excelApp = Globals.ThisAddIn.Application;

                Label2.Enabled = false;
                TextBox3.Enabled = false;
                PictureBox6.Enabled = false;

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

                if (FocusedTextBox == 1)
                {
                    TextBox1.Text = selectedRange.get_Address();
                    workBook = excelApp.ActiveWorkbook;
                    workSheet = (Excel.Worksheet)workBook.ActiveSheet;
                    rng = selectedRange;
                    TextBox1.Focus();
                }

                else if (FocusedTextBox == 3)
                {
                    TextBox3.Text = selectedRange.get_Address();
                    workbook2 = excelApp.ActiveWorkbook;
                    workSheet2 = (Excel.Worksheet)workbook2.ActiveSheet;
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

                if (RadioButton1.Checked == false & RadioButton2.Checked == false & RadioButton3.Checked == false)
                {
                    MessageBox.Show("Select a Merge Type.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                    if (CheckBox1.Checked == false)
                    {
                        rng2.ClearFormats();
                    }
                }
                else
                {
                    rng.Copy();
                    rng2.PasteSpecial(XlPasteType.xlPasteValues);
                    if (CheckBox1.Checked == true)
                    {
                        rng2.PasteSpecial(XlPasteType.xlPasteFormats);
                    }
                    excelApp.CutCopyMode = XlCutCopyMode.xlCopy;
                }

                rng2.Select();

                int r = rng2.Rows.Count;
                int c = rng.Columns.Count;

                int i;
                int j;

                if (RadioButton1.Checked == true | RadioButton2.Checked == true | RadioButton3.Checked == true)
                {

                    excelApp.DisplayAlerts = false;

                    int mergeCount = 0;

                    if (RadioButton1.Checked == true)
                    {
                        var loopTo = r;
                        for (i = 1; i <= loopTo; i++)
                        {
                            var loopTo1 = c;
                            for (j = 1; j <= loopTo1; j++)
                            {
                                int rowEqual = Conversions.ToInteger(SearchAlongRow(rng2, i, j));
                                if (rowEqual > 1)
                                {
                                    workSheet2.get_Range(rng2.Cells[i, j], rng2.Cells[i, j + rowEqual - 1]).Merge();
                                    rng2.Cells[i, j].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                                    mergeCount = mergeCount + 1;
                                }
                                j = j + rowEqual - 1;
                            }
                        }
                    }

                    else if (RadioButton2.Checked == true == true)
                    {

                        var loopTo2 = c;
                        for (j = 1; j <= loopTo2; j++)
                        {
                            var loopTo3 = r;
                            for (i = 1; i <= loopTo3; i++)
                            {
                                int columnEqual = Conversions.ToInteger(SearchAlongColumn(rng2, i, j));
                                if (columnEqual > 1)
                                {
                                    workSheet2.get_Range(rng2.Cells[i, j], rng2.Cells[i + columnEqual - 1, j]).Merge();
                                    rng2.Cells[i, j].VerticalAlignment = XlVAlign.xlVAlignCenter;
                                    mergeCount = mergeCount + 1;
                                }
                                i = i + columnEqual - 1;
                            }
                        }
                    }

                    else if (RadioButton3.Checked == true)
                    {

                        Excel.Worksheet activesheet = (Excel.Worksheet)excelApp.ActiveSheet;

                        var Arr = new int[(r * c), 4];
                        var loopTo4 = Information.UBound(Arr, 1);
                        for (i = Information.LBound(Arr, 1); i <= loopTo4; i++)
                            Arr[i, 0] = 0;

                        int Count = 0;

                        var loopTo5 = r;
                        for (i = 1; i <= loopTo5; i++)
                        {
                            var loopTo6 = c;
                            for (j = 1; j <= loopTo6; j++)
                            {
                                int rowEqual = Conversions.ToInteger(SearchDiagonally(rng2, i, j)((object)0));
                                int columnEqual = Conversions.ToInteger(SearchDiagonally(rng2, i, j)((object)1));
                                Arr[Count, 0] = i;
                                Arr[Count, 1] = j;
                                Arr[Count, 2] = i + rowEqual - 1;
                                Arr[Count, 3] = j + columnEqual - 1;
                                Count = Count + 1;
                            }
                        }

                        Arr = (int[,])RemoveCrossings(excelApp, Arr);

                        var loopTo7 = r;
                        for (i = 1; i <= loopTo7; i++)
                        {
                            var loopTo8 = c;
                            for (j = 1; j <= loopTo8; j++)
                            {

                                Range MRng;
                                MRng = activesheet.get_Range(rng2.Cells[i, j].Address);

                                for (int m = Information.LBound(Arr, 1), loopTo9 = Information.UBound(Arr, 1); m <= loopTo9; m++)
                                {

                                    if (Arr[m, 0] > 0)
                                    {

                                        var Rng1 = activesheet.get_Range("A1");
                                        Rng1 = activesheet.get_Range(Rng1.Cells[Arr[m, 0], Arr[m, 1]], Rng1.Cells[Arr[m, 2], Arr[m, 3]]);

                                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(IsWithinRange(i, j, Rng1), true, false)))
                                        {
                                            MRng = Rng1;
                                            Arr[m, 0] = 0;
                                            break;
                                        }

                                    }

                                }

                                if (MRng.Rows.Count > 1 | MRng.Columns.Count > 1)
                                {
                                    workSheet2.get_Range(rng2.Cells[i, j], rng2.Cells[i + MRng.Rows.Count - 1, j + MRng.Columns.Count - 1]).Merge();
                                    mergeCount = mergeCount + 1;
                                }

                                if (MRng.Columns.Count > 1)
                                {
                                    rng2.Cells[i, j].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                                }

                                if (MRng.Rows.Count > 1)
                                {
                                    rng2.Cells[i, j].VerticalAlignment = XlVAlign.xlVAlignCenter;
                                }

                            }
                        }

                    }
                    excelApp.DisplayAlerts = false;

                    string msg;
                    if (mergeCount > 1)
                    {
                        msg = " cells have been merged.";
                    }
                    else
                    {
                        msg = " cell has been merged.";
                    }
                    Interaction.MsgBox(Conversion.Str(mergeCount) + msg, Title: "Merge Cells");
                }
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
                workbook2 = excelApp.ActiveWorkbook;

                Range userInput = (Range)excelApp.InputBox("Select a range", Type: 8);
                rng2 = userInput;


                string sheetName;
                sheetName = Strings.Split(rng2.get_Address(true, true, XlReferenceStyle.xlA1, true), "]")[1];
                sheetName = Strings.Split(sheetName, "!")[0];

                if (Strings.Mid(sheetName, Strings.Len(sheetName), 1) == "'")
                {
                    sheetName = Strings.Mid(sheetName, 1, Strings.Len(sheetName) - 1);
                }

                workSheet2 = (Excel.Worksheet)workbook2.Worksheets[sheetName];
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

        private void PictureBox9_Click(object sender, EventArgs e)
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

        private void PictureBox4_Click(object sender, EventArgs e)
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

        private void TextBox3_TextChanged(object sender, EventArgs e)
        {

            try
            {
                excelApp = Globals.ThisAddIn.Application;
                workbook2 = excelApp.ActiveWorkbook;
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

        private void RadioButton9_CheckedChanged(object sender, EventArgs e)
        {

            try
            {
                if (RadioButton9.Checked == true)
                {
                    workbook2 = workBook;
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
                    Label2.Enabled = true;
                    TextBox3.Enabled = true;
                    PictureBox6.Enabled = true;
                    TextBox3.Focus();
                }
                else
                {
                    TextBox3.Clear();
                    Label2.Enabled = false;
                    TextBox3.Enabled = false;
                    PictureBox6.Enabled = false;
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

        private void CustomGroupBox7_GotFocus(object sender, EventArgs e)
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

        private void Label2_GotFocus(object sender, EventArgs e)
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

        private void PictureBox4_KeyDown(object sender, KeyEventArgs e)
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

        private void PictureBox5_KeyDown(object sender, KeyEventArgs e)
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

        private void Form8_Closing(object sender, CancelEventArgs e)
        {
            GlobalModule.form_flag = false;
        }

        private void Form8_Disposed(object sender, EventArgs e)
        {
            GlobalModule.form_flag = false;
        }

        private void Form8_Shown(object sender, EventArgs e)
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