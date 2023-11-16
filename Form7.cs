using System;
using global::System.CodeDom;
using System.Collections.Generic;
using global::System.ComponentModel;
using global::System.Data;
using global::System.Diagnostics;
using global::System.Drawing;
using System.Linq;
using global::System.Reflection.Emit;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using global::System.Text.RegularExpressions;
using global::System.Threading;
using global::System.Windows.Forms;
using static global::System.Windows.Forms.VisualStyles.VisualStyleElement;
using static global::System.Windows.Forms.VisualStyles.VisualStyleElement.Button;
using System.Xml.Linq;
using global::Microsoft.Office.Core;
using Office = Microsoft.Office.Core;
using global::Microsoft.Office.Interop.Excel;
using Excel = global::Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace VSTO_Addins
{

    public partial class Form7
    {

        private global::Microsoft.Office.Interop.Excel.Application _excelApp;

        private global::Microsoft.Office.Interop.Excel.Application excelApp
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
        private global::Microsoft.Office.Interop.Excel.Workbook workbook;
        private global::Microsoft.Office.Interop.Excel.Workbook workbook2;
        private global::Microsoft.Office.Interop.Excel.Worksheet worksheet;
        private global::Microsoft.Office.Interop.Excel.Worksheet worksheet2;
        public global::Microsoft.Office.Interop.Excel.Worksheet OpenSheet;
        private global::Microsoft.Office.Interop.Excel.Range rng;
        private global::Microsoft.Office.Interop.Excel.Range rng2;
        private global::System.Int32 FocusedTextBox;
        private global::System.Int32 opened;
        private global::System.Boolean TextBoxChanged;

        public Form7()
        {
            InitializeComponent();
        }


        [DllImport("user32")]
        private static extern bool SetWindowPos(global::System.IntPtr hWnd, global::System.IntPtr hWndInsertAfter, global::System.Int32 X, global::System.Int32 Y, global::System.Int32 cx, global::System.Int32 cy, global::System.UInt32 uFlags);
        private const global::System.UInt32 SWP_NOMOVE = 0x2U;
        private const global::System.UInt32 SWP_NOSIZE = 0x1U;
        private const global::System.UInt32 SWP_NOACTIVATE = 0x10U;
        private const global::System.Int32 HWND_TOPMOST = -(1);

        private global::System.Boolean Overlap(global::Microsoft.Office.Interop.Excel.Application excelApp, global::Microsoft.Office.Interop.Excel.Worksheet sheet1, global::Microsoft.Office.Interop.Excel.Worksheet sheet2, global::Microsoft.Office.Interop.Excel.Range rng1, global::Microsoft.Office.Interop.Excel.Range rng2)
        {

            if (((sheet1.Name ?? "") != (sheet2.Name ?? "")))
            {
                return false;
            }

            else
            {
                global::Microsoft.Office.Interop.Excel.Worksheet activesheet = ((global::Microsoft.Office.Interop.Excel.Worksheet)(excelApp.ActiveSheet));

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
        private global::System.Boolean CanConvertToInt(global::System.String input)
        {

            global::System.Int32 number;

            // TryParse returns True if conversion is possible, False otherwise
            if (global::System.Int32.TryParse(input, out number))
            {
                return true;
            }
            else
            {
                return false;
            }

        }
        private global::System.Boolean IsValidExcelCellReference(global::System.String cellReference)
        {

            global::System.String cellPattern = @"(\$?[A-Z]+\$?[0-9]+)";
            global::System.String referencePattern = ((((("^") + (cellPattern)) + ("(:")) + (cellPattern)) + (")?$"));

            var regex = new global::System.Text.RegularExpressions.Regex(referencePattern);

            global::System.String[] refArr = global::Microsoft.VisualBasic.Strings.Split(cellReference, "!");

            global::System.String reference = refArr[global::Microsoft.VisualBasic.Information.UBound(refArr)];

            if (regex.IsMatch(reference))
            {
                return true;
            }
            else
            {
                return false;
            }

        }

        private void Setup()
        {

            try
            {
                this.excelApp = global::VSTO_Addins.Globals.ThisAddIn.Application;
                this.workbook = excelApp.ActiveWorkbook;
                this.worksheet = (global::Microsoft.Office.Interop.Excel.Worksheet)workbook.ActiveSheet;

                this.rng = this.worksheet.get_Range(this.TextBox1.Text);

                global::System.Int32 r;
                global::System.Int32 c;

                r = this.rng.Rows.Count;
                c = this.rng.Columns.Count;

                if ((((r) != (1)) & ((c) != (1))))
                {
                    this.RadioButton1.Enabled = true;
                    this.RadioButton2.Enabled = true;
                    this.RadioButton3.Enabled = false;
                    this.RadioButton4.Enabled = false;
                }
                else if ((((r) != (1)) & ((c) == (1))))
                {
                    this.RadioButton1.Enabled = false;
                    this.RadioButton2.Enabled = true;
                    this.RadioButton3.Enabled = true;
                    this.RadioButton4.Enabled = false;
                }
                else if ((((r) == (1)) & ((c) != (1))))
                {
                    this.RadioButton1.Enabled = true;
                    this.RadioButton2.Enabled = false;
                    this.RadioButton3.Enabled = false;
                    this.RadioButton4.Enabled = true;

                }

                if ((((this.RadioButton1.Checked) == (true)) | ((this.RadioButton2.Checked) == (true))))
                {
                    this.TextBox2.Text = "";
                    this.CustomGroupBox3.Enabled = false;
                    this.RadioButton7.Checked = false;
                    this.RadioButton8.Checked = false;
                }
                else
                {
                    this.CustomGroupBox3.Enabled = true;
                }

                if (((this.RadioButton8.Checked) == (true)))
                {
                    this.TextBox2.Enabled = true;
                    this.TextBox2.Focus();
                }
                else
                {
                    this.TextBox2.Text = "";
                    this.TextBox2.Enabled = false;
                }

                if (((this.RadioButton3.Checked) == (true)))
                {
                    this.RadioButton8.Text = "After number of rows:";
                }
                else if (((this.RadioButton4.Checked) == (true)))
                {
                    this.RadioButton8.Text = "After number of columns:";
                }
            }

            catch (global::System.Exception ex)
            {

            }

        }

        private global::System.Object MaxValue(global::System.Object Arr)
        {
            global::System.Object MaxValueRet = default(global::System.Object);

            global::System.Int32 max;
            max = Conversions.ToInteger(Arr((global::System.Object)global::Microsoft.VisualBasic.Information.LBound((global::System.Array)Arr)));

            for (global::System.Int32 i = (global::Microsoft.VisualBasic.Information.LBound((global::System.Array)Arr)) + (1), loopTo = global::Microsoft.VisualBasic.Information.UBound((global::System.Array)Arr); i <= loopTo; i++)
            {
                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectGreater(Arr((global::System.Object)i), max, false)))
                {
                    max = Conversions.ToInteger(Arr((global::System.Object)i));
                }
            }

            MaxValueRet = (global::System.Object)max;
            return MaxValueRet;

        }

        private global::System.Object GetBreakPoints(global::Microsoft.Office.Interop.Excel.Range rng, global::System.Int32 trace)
        {
            global::System.Object GetBreakPointsRet = default(global::System.Object);

            var Arr = default(global::System.Int32[]);
            global::System.Int32 Index;
            Index = -(1);

            if (((trace) == (1)))
            {
                for (global::System.Int32 j = 1, loopTo = rng.Columns.Count; j <= loopTo; j++)
                {
                    if (((((rng.Cells[(global::System.Object)1, (global::System.Object)j].Value == null)) || ((rng.Cells[(global::System.Object)1, (global::System.Object)j].Value is System.DBNull))) || (global::System.String.IsNullOrEmpty(rng.Cells[(global::System.Object)1, (global::System.Object)j].Value.ToString()))))
                    {
                        Index = ((Index) + (1));
                        Array.Resize(ref Arr, Index + 1);
                        Arr[Index] = j;
                    }
                }

                Index = ((Index) + (1));
                Array.Resize(ref Arr, Index + 1);
                Arr[Index] = ((rng.Columns.Count) + (1));
            }

            else
            {
                for (global::System.Int32 i = 1, loopTo1 = rng.Rows.Count; i <= loopTo1; i++)
                {
                    if (((((rng.Cells[(global::System.Object)i, (global::System.Object)1].Value == null)) || ((rng.Cells[(global::System.Object)i, (global::System.Object)1].Value is System.DBNull))) || (global::System.String.IsNullOrEmpty(rng.Cells[(global::System.Object)i, (global::System.Object)1].Value.ToString()))))
                    {
                        Index = ((Index) + (1));
                        Array.Resize(ref Arr, Index + 1);
                        Arr[Index] = i;
                    }
                }
                Index = ((Index) + (1));
                Array.Resize(ref Arr, Index + 1);
                Arr[Index] = ((rng.Rows.Count) + (1));
            }

            GetBreakPointsRet = Arr;
            return GetBreakPointsRet;

        }
        private global::System.Object GetLengths(global::System.Object Arr)
        {
            global::System.Object GetLengthsRet = default(global::System.Object);

            var Arr2 = new global::System.Int32[1];
            global::System.Int32 Index;
            Index = -(1);
            global::System.Int32 position;
            position = 0;
            global::System.Int32 length;

            for (global::System.Int32 i = global::Microsoft.VisualBasic.Information.LBound((global::System.Array)Arr), loopTo = global::Microsoft.VisualBasic.Information.UBound((global::System.Array)Arr); i <= loopTo; i++)
            {
                length = Conversions.ToInteger(Operators.SubtractObject(Operators.SubtractObject(Arr((global::System.Object)i), position), 1));
                position = Conversions.ToInteger(Arr((global::System.Object)i));
                Index = ((Index) + (1));
                Array.Resize(ref Arr2, Index + 1);
                Arr2[Index] = length;
            }

            GetLengthsRet = Arr2;
            return GetLengthsRet;

        }
        private global::System.Object MaxOfColumn(global::Microsoft.Office.Interop.Excel.Range cRng)
        {
            global::System.Object MaxOfColumnRet = default(global::System.Object);

            global::System.Int32 max;
            global::System.Int32 CharNumbers;

            if (global::Microsoft.VisualBasic.Information.IsNumeric(cRng.Cells[(global::System.Object)1, (global::System.Object)1].value))
            {
                max = global::Microsoft.VisualBasic.Strings.Len(global::Microsoft.VisualBasic.Conversion.Str(cRng.Cells[(global::System.Object)1, (global::System.Object)1].value));
            }
            else
            {
                max = global::Microsoft.VisualBasic.Strings.Len(cRng.Cells[(global::System.Object)1, (global::System.Object)1].value);
            }

            for (global::System.Int32 i = 2, loopTo = cRng.Rows.Count; i <= loopTo; i++)
            {
                if (global::Microsoft.VisualBasic.Information.IsNumeric(cRng.Cells[(global::System.Object)i, (global::System.Object)1].value))
                {
                    CharNumbers = global::Microsoft.VisualBasic.Strings.Len(global::Microsoft.VisualBasic.Conversion.Str(cRng.Cells[(global::System.Object)i, (global::System.Object)1].value));
                }
                else
                {
                    CharNumbers = global::Microsoft.VisualBasic.Strings.Len(cRng.Cells[(global::System.Object)i, (global::System.Object)1].value);
                }
                if (((CharNumbers) > (max)))
                {
                    max = CharNumbers;
                }
            }

            if (((max) < (7)))
            {
                max = 7;
            }

            MaxOfColumnRet = (global::System.Object)max;
            return MaxOfColumnRet;

        }
        private global::System.Object MaxOfArray(global::System.Object Arr)
        {
            global::System.Object MaxOfArrayRet = default(global::System.Object);

            global::System.Int32 max;
            max = global::Microsoft.VisualBasic.Strings.Len(Arr((global::System.Object)global::Microsoft.VisualBasic.Information.LBound((global::System.Array)Arr)));

            for (global::System.Int32 i = (global::Microsoft.VisualBasic.Information.LBound((global::System.Array)Arr)) + (1), loopTo = global::Microsoft.VisualBasic.Information.UBound((global::System.Array)Arr); i <= loopTo; i++)
            {
                if (((global::Microsoft.VisualBasic.Strings.Len(Arr((global::System.Object)i))) > (max)))
                {
                    max = global::Microsoft.VisualBasic.Strings.Len(Arr((global::System.Object)i));
                }
            }

            if (((max) < (7)))
            {
                max = 7;
            }

            MaxOfArrayRet = (global::System.Object)max;
            return MaxOfArrayRet;

        }
        private global::System.Object AdjustWidth(global::System.Object Widths, global::System.Object CWidth)
        {
            global::System.Object AdjustWidthRet = default(global::System.Object);

            global::System.Double SumWidth = 0d;

            for (global::System.Int32 i = global::Microsoft.VisualBasic.Information.LBound((global::System.Array)Widths), loopTo = global::Microsoft.VisualBasic.Information.UBound((global::System.Array)Widths); i <= loopTo; i++)
                SumWidth = Conversions.ToDouble(Operators.AddObject(SumWidth, Widths((global::System.Object)i)));

            if (Conversions.ToBoolean(Operators.ConditionalCompareObjectLess(SumWidth, CWidth, false)))
            {
                global::System.Double Extra = Conversions.ToDouble(Operators.SubtractObject(CWidth, SumWidth));
                Extra = ((Extra) / (global::System.Double)((((global::Microsoft.VisualBasic.Information.UBound((global::System.Array)Widths)) + (1)))));
                for (global::System.Int32 i = global::Microsoft.VisualBasic.Information.LBound((global::System.Array)Widths), loopTo1 = global::Microsoft.VisualBasic.Information.UBound((global::System.Array)Widths); i <= loopTo1; i++)
                    Widths((global::System.Object)i) = Operators.AddObject(Widths((global::System.Object)i), Extra);
            }

            AdjustWidthRet = Widths;
            return AdjustWidthRet;

        }

        private void Display()
        {

            try
            {

                this.TextBoxChanged = true;

                this.CustomPanel1.Controls.Clear();
                this.CustomPanel2.Controls.Clear();

                global::Microsoft.Office.Interop.Excel.Range displayRng;

                displayRng = this.rng;

                if (((this.rng.Rows.Count) > (50)))
                {
                    displayRng = this.worksheet.get_Range(this.rng.Cells[(global::System.Object)1, (global::System.Object)1], this.rng.Cells[(global::System.Object)50, (global::System.Object)this.rng.Columns.Count]);
                }

                if (((this.rng.Columns.Count) > (50)))
                {
                    displayRng = this.worksheet.get_Range(this.rng.Cells[(global::System.Object)1, (global::System.Object)1], this.rng.Cells[(global::System.Object)this.rng.Rows.Count, (global::System.Object)50]);
                }

                global::System.Int32 r;
                global::System.Int32 c;

                r = displayRng.Rows.Count;
                c = displayRng.Columns.Count;

                global::System.Double height;
                global::System.Double BaseWidth;
                global::System.Double width;

                if ((((r) > (1)) & ((r) <= (6))))
                {
                    height = ((global::System.Double)(this.CustomPanel1.Height) / (global::System.Double)(r));
                }
                else
                {
                    height = ((global::System.Double)(this.CustomPanel1.Height) / (6d));
                }

                BaseWidth = ((260d) / (3d));

                global::System.Double Ordinate = 0d;
                var Widths = new global::System.Double[(c)];

                for (global::System.Int32 j = 1, loopTo = c; j <= loopTo; j++)
                {
                    var CRng = this.worksheet.get_Range(displayRng.Cells[(global::System.Object)1, (global::System.Object)j], displayRng.Cells[(global::System.Object)r, (global::System.Object)j]);
                    Widths[(j) - (1)] = Conversions.ToDouble(Operators.DivideObject((Operators.MultiplyObject(this.MaxOfColumn(CRng), BaseWidth)), 10));
                }

                Widths = (global::System.Double[])this.AdjustWidth(Widths, (global::System.Object)this.CustomPanel2.Width);

                for (global::System.Int32 j = 1, loopTo1 = c; j <= loopTo1; j++)
                {
                    for (global::System.Int32 i = 1, loopTo2 = r; i <= loopTo2; i++)
                    {
                        var label = new global::System.Windows.Forms.Label();
                        label.Text = Conversions.ToString(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Value);
                        label.Location = new global::System.Drawing.Point((global::System.Int32)Math.Round(Ordinate), (global::System.Int32)Math.Round((global::System.Double)((((i) - (1)))) * (height)));
                        label.Height = (global::System.Int32)Math.Round(height);
                        label.Width = (global::System.Int32)Math.Round(Widths[(j) - (1)]);
                        label.BorderStyle = global::System.Windows.Forms.BorderStyle.FixedSingle;
                        label.TextAlign = global::System.Drawing.ContentAlignment.MiddleCenter;

                        if (((this.CheckBox1.Checked) == (true)))
                        {

                            global::Microsoft.Office.Interop.Excel.Range cell = (global::Microsoft.Office.Interop.Excel.Range)displayRng.Cells[(global::System.Object)i, (global::System.Object)j];
                            var font = cell.Font;
                            var fontStyle = global::System.Drawing.FontStyle.Regular;
                            if (Conversions.ToBoolean(cell.Font.Bold))
                                fontStyle = (fontStyle | global::System.Drawing.FontStyle.Bold);
                            if (Conversions.ToBoolean(cell.Font.Italic))
                                fontStyle = (fontStyle | global::System.Drawing.FontStyle.Italic);


                            global::System.Single fontSize = global::System.Convert.ToSingle(font.Size);

                            label.Font = new global::System.Drawing.Font(font.ToString(), fontSize, fontStyle);
                            if (Conversions.ToBoolean(!(Operators.ConditionalCompareObjectEqual(cell.Interior.ColorIndex, global::Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexNone, false))))
                            {
                                global::System.Int64 colorValue1 = Conversions.ToLong(cell.Interior.Color);
                                global::System.Int32 red1 = (global::System.Int32)((colorValue1) % (256L));
                                global::System.Int32 green1 = (global::System.Int32)(((((colorValue1) / (256L)))) % (256L));
                                global::System.Int32 blue1 = (global::System.Int32)((((((colorValue1) / (256L)) / (256L)))) % (256L));
                                label.BackColor = global::System.Drawing.Color.FromArgb(red1, green1, blue1);
                            }

                            if ((cell.Font.Color is System.DBNull))
                            {
                                label.ForeColor = global::System.Drawing.Color.FromArgb(0, 0, 0);
                            }

                            else if (Conversions.ToBoolean(!(Operators.ConditionalCompareObjectEqual(cell.Font.ColorIndex, global::Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexNone, false))))
                            {
                                global::System.Int64 colorValue2 = Conversions.ToLong(cell.Font.Color);
                                global::System.Int32 red2 = (global::System.Int32)((colorValue2) % (256L));
                                global::System.Int32 green2 = (global::System.Int32)(((((colorValue2) / (256L)))) % (256L));
                                global::System.Int32 blue2 = (global::System.Int32)((((((colorValue2) / (256L)) / (256L)))) % (256L));
                                label.ForeColor = global::System.Drawing.Color.FromArgb(red2, green2, blue2);
                            }
                        }
                        this.CustomPanel1.Controls.Add(label);
                    }
                    Ordinate = ((Ordinate) + (Widths[(j) - (1)]));
                }

                this.CustomPanel1.AutoScroll = true;

                global::System.Boolean X1;
                X1 = this.RadioButton1.Checked;

                global::System.Boolean X2;
                X2 = this.RadioButton2.Checked;

                global::System.Boolean X3;
                X3 = this.RadioButton3.Checked;

                global::System.Boolean X4;
                X4 = this.RadioButton4.Checked;

                global::System.Boolean X5;
                X5 = this.RadioButton5.Checked;

                global::System.Boolean X6;
                X6 = this.RadioButton6.Checked;

                global::System.Boolean X7;
                X7 = this.RadioButton7.Checked;

                global::System.Boolean X8;
                X8 = this.RadioButton8.Checked;


                if (X1)
                {

                    if ((((((r) * (c)))) <= (6)))
                    {
                        height = ((global::System.Double)(this.CustomPanel2.Height) / (global::System.Double)((((r) * (c)))));
                    }
                    else
                    {
                        height = ((global::System.Double)(this.CustomPanel2.Height) / (6d));
                    }

                    var values = new global::System.Object[(displayRng.Cells.Count)];
                    for (global::System.Int32 k = 1, loopTo3 = displayRng.Cells.Count; k <= loopTo3; k++)
                        values[(k) - (1)] = displayRng.Cells[(global::System.Object)k].value;

                    var Widths2 = new global::System.Double[1];
                    Widths2[0] = Conversions.ToDouble(Operators.DivideObject((Operators.MultiplyObject(this.MaxOfArray(values), BaseWidth)), 10));

                    Widths2 = (global::System.Double[])this.AdjustWidth(Widths2, (global::System.Object)this.CustomPanel2.Width);

                    global::System.Int32 count;
                    count = 1;

                    if (X5)
                    {

                        for (global::System.Int32 i = 1, loopTo4 = r; i <= loopTo4; i++)
                        {
                            for (global::System.Int32 j = 1, loopTo5 = c; j <= loopTo5; j++)
                            {
                                var label = new global::System.Windows.Forms.Label();
                                label.Text = Conversions.ToString(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Value);
                                label.Location = new global::System.Drawing.Point(0, (global::System.Int32)Math.Round((global::System.Double)((((count) - (1)))) * (height)));
                                count = ((count) + (1));
                                label.Height = (global::System.Int32)Math.Round(height);
                                label.Width = (global::System.Int32)Math.Round(Widths2[0]);
                                label.BorderStyle = global::System.Windows.Forms.BorderStyle.FixedSingle;
                                label.TextAlign = global::System.Drawing.ContentAlignment.MiddleCenter;

                                if (((this.CheckBox1.Checked) == (true)))
                                {
                                    global::Microsoft.Office.Interop.Excel.Range cell = (global::Microsoft.Office.Interop.Excel.Range)displayRng.Cells[(global::System.Object)i, (global::System.Object)j];
                                    var font = cell.Font;
                                    var fontStyle = global::System.Drawing.FontStyle.Regular;
                                    if (Conversions.ToBoolean(cell.Font.Bold))
                                        fontStyle = (fontStyle | global::System.Drawing.FontStyle.Bold);
                                    if (Conversions.ToBoolean(cell.Font.Italic))
                                        fontStyle = (fontStyle | global::System.Drawing.FontStyle.Italic);

                                    global::System.Single fontSize = global::System.Convert.ToSingle(font.Size);

                                    label.Font = new global::System.Drawing.Font(font.ToString(), fontSize, fontStyle);
                                    if (Conversions.ToBoolean(!(Operators.ConditionalCompareObjectEqual(cell.Interior.ColorIndex, global::Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexNone, false))))
                                    {
                                        global::System.Int64 colorValue1 = Conversions.ToLong(cell.Interior.Color);
                                        global::System.Int32 red1 = (global::System.Int32)((colorValue1) % (256L));
                                        global::System.Int32 green1 = (global::System.Int32)(((((colorValue1) / (256L)))) % (256L));
                                        global::System.Int32 blue1 = (global::System.Int32)((((((colorValue1) / (256L)) / (256L)))) % (256L));
                                        label.BackColor = global::System.Drawing.Color.FromArgb(red1, green1, blue1);
                                    }

                                    if ((cell.Font.Color is System.DBNull))
                                    {
                                        label.ForeColor = global::System.Drawing.Color.FromArgb(0, 0, 0);
                                    }

                                    else if (Conversions.ToBoolean(!(Operators.ConditionalCompareObjectEqual(cell.Font.ColorIndex, global::Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexNone, false))))
                                    {
                                        global::System.Int64 colorValue2 = Conversions.ToLong(cell.Font.Color);
                                        global::System.Int32 red2 = (global::System.Int32)((colorValue2) % (256L));
                                        global::System.Int32 green2 = (global::System.Int32)(((((colorValue2) / (256L)))) % (256L));
                                        global::System.Int32 blue2 = (global::System.Int32)((((((colorValue2) / (256L)) / (256L)))) % (256L));
                                        label.ForeColor = global::System.Drawing.Color.FromArgb(red2, green2, blue2);
                                    }
                                }

                                this.CustomPanel2.Controls.Add(label);
                            }
                        }
                    }

                    else if (X6)
                    {

                        for (global::System.Int32 j = 1, loopTo6 = c; j <= loopTo6; j++)
                        {
                            for (global::System.Int32 i = 1, loopTo7 = r; i <= loopTo7; i++)
                            {
                                var label = new global::System.Windows.Forms.Label();
                                label.Text = Conversions.ToString(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Value);
                                label.Location = new global::System.Drawing.Point(0, (global::System.Int32)Math.Round((global::System.Double)((((count) - (1)))) * (height)));
                                count = ((count) + (1));
                                label.Height = (global::System.Int32)Math.Round(height);
                                label.Width = (global::System.Int32)Math.Round(Widths2[0]);
                                label.BorderStyle = global::System.Windows.Forms.BorderStyle.FixedSingle;
                                label.TextAlign = global::System.Drawing.ContentAlignment.MiddleCenter;

                                if (((this.CheckBox1.Checked) == (true)))
                                {
                                    global::Microsoft.Office.Interop.Excel.Range cell = (global::Microsoft.Office.Interop.Excel.Range)displayRng.Cells[(global::System.Object)i, (global::System.Object)j];
                                    var font = cell.Font;
                                    var fontStyle = global::System.Drawing.FontStyle.Regular;
                                    if (Conversions.ToBoolean(cell.Font.Bold))
                                        fontStyle = (fontStyle | global::System.Drawing.FontStyle.Bold);
                                    if (Conversions.ToBoolean(cell.Font.Italic))
                                        fontStyle = (fontStyle | global::System.Drawing.FontStyle.Italic);

                                    global::System.Single fontSize = global::System.Convert.ToSingle(font.Size);

                                    label.Font = new global::System.Drawing.Font(font.ToString(), fontSize, fontStyle);
                                    if (Conversions.ToBoolean(!(Operators.ConditionalCompareObjectEqual(cell.Interior.ColorIndex, global::Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexNone, false))))
                                    {
                                        global::System.Int64 colorValue1 = Conversions.ToLong(cell.Interior.Color);
                                        global::System.Int32 red1 = (global::System.Int32)((colorValue1) % (256L));
                                        global::System.Int32 green1 = (global::System.Int32)(((((colorValue1) / (256L)))) % (256L));
                                        global::System.Int32 blue1 = (global::System.Int32)((((((colorValue1) / (256L)) / (256L)))) % (256L));
                                        label.BackColor = global::System.Drawing.Color.FromArgb(red1, green1, blue1);
                                    }

                                    if ((cell.Font.Color is System.DBNull))
                                    {
                                        label.ForeColor = global::System.Drawing.Color.FromArgb(0, 0, 0);
                                    }

                                    else if (Conversions.ToBoolean(!(Operators.ConditionalCompareObjectEqual(cell.Font.ColorIndex, global::Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexNone, false))))
                                    {
                                        global::System.Int64 colorValue2 = Conversions.ToLong(cell.Font.Color);
                                        global::System.Int32 red2 = (global::System.Int32)((colorValue2) % (256L));
                                        global::System.Int32 green2 = (global::System.Int32)(((((colorValue2) / (256L)))) % (256L));
                                        global::System.Int32 blue2 = (global::System.Int32)((((((colorValue2) / (256L)) / (256L)))) % (256L));
                                        label.ForeColor = global::System.Drawing.Color.FromArgb(red2, green2, blue2);
                                    }
                                }

                                this.CustomPanel2.Controls.Add(label);
                            }
                        }

                    }

                    this.CustomPanel2.AutoScroll = true;

                }

                if (X2)
                {

                    if ((((((r) * (c)))) <= (4)))
                    {
                        width = ((global::System.Double)(this.CustomPanel2.Width) / (global::System.Double)((((r) * (c)))));
                    }
                    else
                    {
                        width = ((global::System.Double)(this.CustomPanel2.Width) / (4d));
                    }

                    height = ((global::System.Double)(this.CustomPanel2.Height) / (6d));

                    global::System.Int32 count;
                    count = 1;

                    if (X5)
                    {
                        Ordinate = 0d;
                        global::System.Int32 Length;
                        for (global::System.Int32 i = 1, loopTo8 = r; i <= loopTo8; i++)
                        {
                            for (global::System.Int32 j = 1, loopTo9 = c; j <= loopTo9; j++)
                            {
                                Length = global::Microsoft.VisualBasic.Strings.Len(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Value);
                                if (((Length) < (7)))
                                {
                                    Length = 7;
                                }
                                width = (((((global::System.Double)(Length) * (BaseWidth)))) / (10d));
                                var label = new global::System.Windows.Forms.Label();
                                label.Text = Conversions.ToString(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Value);
                                label.Location = new global::System.Drawing.Point((global::System.Int32)Math.Round(Ordinate), (global::System.Int32)Math.Round(((((3.5d) - (1d)))) * (height)));
                                count = ((count) + (1));
                                label.Height = (global::System.Int32)Math.Round(height);
                                label.Width = (global::System.Int32)Math.Round(width);
                                label.BorderStyle = global::System.Windows.Forms.BorderStyle.FixedSingle;
                                label.TextAlign = global::System.Drawing.ContentAlignment.MiddleCenter;

                                if (((this.CheckBox1.Checked) == (true)))
                                {
                                    global::Microsoft.Office.Interop.Excel.Range cell = (global::Microsoft.Office.Interop.Excel.Range)displayRng.Cells[(global::System.Object)i, (global::System.Object)j];
                                    var font = cell.Font;
                                    var fontStyle = global::System.Drawing.FontStyle.Regular;
                                    if (Conversions.ToBoolean(cell.Font.Bold))
                                        fontStyle = (fontStyle | global::System.Drawing.FontStyle.Bold);
                                    if (Conversions.ToBoolean(cell.Font.Italic))
                                        fontStyle = (fontStyle | global::System.Drawing.FontStyle.Italic);

                                    global::System.Single fontSize = global::System.Convert.ToSingle(font.Size);

                                    label.Font = new global::System.Drawing.Font(font.ToString(), fontSize, fontStyle);
                                    if (Conversions.ToBoolean(!(Operators.ConditionalCompareObjectEqual(cell.Interior.ColorIndex, global::Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexNone, false))))
                                    {
                                        global::System.Int64 colorValue1 = Conversions.ToLong(cell.Interior.Color);
                                        global::System.Int32 red1 = (global::System.Int32)((colorValue1) % (256L));
                                        global::System.Int32 green1 = (global::System.Int32)(((((colorValue1) / (256L)))) % (256L));
                                        global::System.Int32 blue1 = (global::System.Int32)((((((colorValue1) / (256L)) / (256L)))) % (256L));
                                        label.BackColor = global::System.Drawing.Color.FromArgb(red1, green1, blue1);
                                    }

                                    if ((cell.Font.Color is System.DBNull))
                                    {
                                        label.ForeColor = global::System.Drawing.Color.FromArgb(0, 0, 0);
                                    }

                                    else if (Conversions.ToBoolean(!(Operators.ConditionalCompareObjectEqual(cell.Font.ColorIndex, global::Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexNone, false))))
                                    {
                                        global::System.Int64 colorValue2 = Conversions.ToLong(cell.Font.Color);
                                        global::System.Int32 red2 = (global::System.Int32)((colorValue2) % (256L));
                                        global::System.Int32 green2 = (global::System.Int32)(((((colorValue2) / (256L)))) % (256L));
                                        global::System.Int32 blue2 = (global::System.Int32)((((((colorValue2) / (256L)) / (256L)))) % (256L));
                                        label.ForeColor = global::System.Drawing.Color.FromArgb(red2, green2, blue2);
                                    }
                                }

                                this.CustomPanel2.Controls.Add(label);
                                Ordinate = ((Ordinate) + (width));
                            }
                        }
                    }

                    else if (X6)
                    {
                        Ordinate = 0d;
                        global::System.Int32 Length;
                        for (global::System.Int32 j = 1, loopTo10 = c; j <= loopTo10; j++)
                        {
                            for (global::System.Int32 i = 1, loopTo11 = r; i <= loopTo11; i++)
                            {
                                Length = global::Microsoft.VisualBasic.Strings.Len(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Value);
                                if (((Length) < (7)))
                                {
                                    Length = 7;
                                }
                                width = (((((global::System.Double)(Length) * (BaseWidth)))) / (10d));
                                var label = new global::System.Windows.Forms.Label();
                                label.Text = Conversions.ToString(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Value);
                                label.Location = new global::System.Drawing.Point((global::System.Int32)Math.Round(Ordinate), (global::System.Int32)Math.Round(((((3.5d) - (1d)))) * (height)));
                                count = ((count) + (1));
                                label.Height = (global::System.Int32)Math.Round(height);
                                label.Width = (global::System.Int32)Math.Round(width);
                                label.BorderStyle = global::System.Windows.Forms.BorderStyle.FixedSingle;
                                label.TextAlign = global::System.Drawing.ContentAlignment.MiddleCenter;

                                if (((this.CheckBox1.Checked) == (true)))
                                {
                                    global::Microsoft.Office.Interop.Excel.Range cell = (global::Microsoft.Office.Interop.Excel.Range)displayRng.Cells[(global::System.Object)i, (global::System.Object)j];
                                    var font = cell.Font;
                                    var fontStyle = global::System.Drawing.FontStyle.Regular;
                                    if (Conversions.ToBoolean(cell.Font.Bold))
                                        fontStyle = (fontStyle | global::System.Drawing.FontStyle.Bold);
                                    if (Conversions.ToBoolean(cell.Font.Italic))
                                        fontStyle = (fontStyle | global::System.Drawing.FontStyle.Italic);

                                    global::System.Single fontSize = global::System.Convert.ToSingle(font.Size);

                                    label.Font = new global::System.Drawing.Font(font.ToString(), fontSize, fontStyle);
                                    if (Conversions.ToBoolean(!(Operators.ConditionalCompareObjectEqual(cell.Interior.ColorIndex, global::Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexNone, false))))
                                    {
                                        global::System.Int64 colorValue1 = Conversions.ToLong(cell.Interior.Color);
                                        global::System.Int32 red1 = (global::System.Int32)((colorValue1) % (256L));
                                        global::System.Int32 green1 = (global::System.Int32)(((((colorValue1) / (256L)))) % (256L));
                                        global::System.Int32 blue1 = (global::System.Int32)((((((colorValue1) / (256L)) / (256L)))) % (256L));
                                        label.BackColor = global::System.Drawing.Color.FromArgb(red1, green1, blue1);
                                    }

                                    if ((cell.Font.Color is System.DBNull))
                                    {
                                        label.ForeColor = global::System.Drawing.Color.FromArgb(0, 0, 0);
                                    }

                                    else if (Conversions.ToBoolean(!(Operators.ConditionalCompareObjectEqual(cell.Font.ColorIndex, global::Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexNone, false))))
                                    {
                                        global::System.Int64 colorValue2 = Conversions.ToLong(cell.Font.Color);
                                        global::System.Int32 red2 = (global::System.Int32)((colorValue2) % (256L));
                                        global::System.Int32 green2 = (global::System.Int32)(((((colorValue2) / (256L)))) % (256L));
                                        global::System.Int32 blue2 = (global::System.Int32)((((((colorValue2) / (256L)) / (256L)))) % (256L));
                                        label.ForeColor = global::System.Drawing.Color.FromArgb(red2, green2, blue2);
                                    }
                                }
                                this.CustomPanel2.Controls.Add(label);
                                Ordinate = ((Ordinate) + (width));
                            }
                        }
                    }

                    this.CustomPanel2.AutoScroll = true;

                }

                if (X3)
                {

                    if (((X7) & ((((X5) | (X6))))))
                    {

                        global::System.Int32[] BreakPoints;
                        BreakPoints = (global::System.Int32[])this.GetBreakPoints(displayRng, 2);

                        global::System.Int32[] lengths;
                        lengths = (global::System.Int32[])this.GetLengths(BreakPoints);

                        if (X5)
                        {
                            r = ((global::Microsoft.VisualBasic.Information.UBound(BreakPoints)) + (1));
                            c = Conversions.ToInteger(this.MaxValue(lengths));
                        }
                        else if (X6)
                        {
                            c = ((global::Microsoft.VisualBasic.Information.UBound(BreakPoints)) + (1));
                            r = Conversions.ToInteger(this.MaxValue(lengths));
                        }

                        if ((((r) > (1)) & ((r) <= (6))))
                        {
                            height = ((global::System.Double)(this.CustomPanel2.Height) / (global::System.Double)(r));
                        }
                        else
                        {
                            height = ((global::System.Double)(this.CustomPanel2.Height) / (6d));
                        }

                        var Values = new global::System.Object[(r), (c)];
                        var References = new global::System.Int32[(r), (c)];

                        if (X5)
                        {
                            global::System.Int32 iRow;
                            iRow = 0;
                            for (global::System.Int32 i = 1, loopTo12 = r; i <= loopTo12; i++)
                            {
                                for (global::System.Int32 j = 1, loopTo13 = c; j <= loopTo13; j++)
                                {
                                    global::System.Int32 x;
                                    global::System.Int32 y;
                                    x = ((iRow) + (j));
                                    y = 1;
                                    if (((x) <= (BreakPoints[(i) - (1)])))
                                    {
                                        Values[(i) - (1), (j) - (1)] = displayRng.Cells[(global::System.Object)x, (global::System.Object)y].Value;
                                    }
                                    else
                                    {
                                        Values[(i) - (1), (j) - (1)] = "";
                                    }
                                    References[(i) - (1), (j) - (1)] = x;
                                }
                                iRow = BreakPoints[(i) - (1)];
                            }
                        }

                        else if (X6)
                        {
                            global::System.Int32 iRow;
                            iRow = 0;
                            for (global::System.Int32 j = 1, loopTo14 = c; j <= loopTo14; j++)
                            {
                                for (global::System.Int32 i = 1, loopTo15 = r; i <= loopTo15; i++)
                                {
                                    global::System.Int32 x;
                                    global::System.Int32 y;
                                    x = ((iRow) + (i));
                                    y = 1;
                                    if (((x) <= (BreakPoints[(j) - (1)])))
                                    {
                                        Values[(i) - (1), (j) - (1)] = displayRng.Cells[(global::System.Object)x, (global::System.Object)y].Value;
                                    }
                                    else
                                    {
                                        Values[(i) - (1), (j) - (1)] = "";
                                    }
                                    References[(i) - (1), (j) - (1)] = x;
                                }
                                iRow = BreakPoints[(j) - (1)];
                            }
                        }

                        var ColumnValues = new global::System.Object[(r)];
                        var widths2 = new global::System.Double[(c)];

                        for (global::System.Int32 j = 0, loopTo16 = (c) - (1); j <= loopTo16; j++)
                        {
                            for (global::System.Int32 i = 0, loopTo17 = (r) - (1); i <= loopTo17; i++)
                                ColumnValues[i] = Values[i, j];
                            widths2[j] = Conversions.ToDouble(Operators.DivideObject((Operators.MultiplyObject(this.MaxOfArray(ColumnValues), BaseWidth)), 10));
                        }
                        widths2 = (global::System.Double[])this.AdjustWidth(widths2, (global::System.Object)this.CustomPanel2.Width);

                        Ordinate = 0d;

                        for (global::System.Int32 j = 1, loopTo18 = c; j <= loopTo18; j++)
                        {
                            for (global::System.Int32 i = 1, loopTo19 = r; i <= loopTo19; i++)
                            {
                                var label = new global::System.Windows.Forms.Label();
                                label.Text = Conversions.ToString(Values[(i) - (1), (j) - (1)]);
                                label.Location = new global::System.Drawing.Point((global::System.Int32)Math.Round(Ordinate), (global::System.Int32)Math.Round((global::System.Double)((((i) - (1)))) * (height)));
                                label.Height = (global::System.Int32)Math.Round(height);
                                label.Width = (global::System.Int32)Math.Round(widths2[(j) - (1)]);
                                label.BorderStyle = global::System.Windows.Forms.BorderStyle.FixedSingle;
                                label.TextAlign = global::System.Drawing.ContentAlignment.MiddleCenter;

                                if (((this.CheckBox1.Checked) == (true)))
                                {
                                    global::System.Int32 x;
                                    global::System.Int32 y;
                                    x = References[(i) - (1), (j) - (1)];
                                    y = 1;
                                    global::Microsoft.Office.Interop.Excel.Range cell = (global::Microsoft.Office.Interop.Excel.Range)displayRng.Cells[(global::System.Object)x, (global::System.Object)y];
                                    var font = cell.Font;
                                    var fontStyle = global::System.Drawing.FontStyle.Regular;
                                    if (Conversions.ToBoolean(cell.Font.Bold))
                                        fontStyle = (fontStyle | global::System.Drawing.FontStyle.Bold);
                                    if (Conversions.ToBoolean(cell.Font.Italic))
                                        fontStyle = (fontStyle | global::System.Drawing.FontStyle.Italic);

                                    global::System.Single fontSize = global::System.Convert.ToSingle(font.Size);

                                    label.Font = new global::System.Drawing.Font(font.ToString(), fontSize, fontStyle);
                                    if (Conversions.ToBoolean(!(Operators.ConditionalCompareObjectEqual(cell.Interior.ColorIndex, global::Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexNone, false))))
                                    {
                                        global::System.Int64 colorValue1 = Conversions.ToLong(cell.Interior.Color);
                                        global::System.Int32 red1 = (global::System.Int32)((colorValue1) % (256L));
                                        global::System.Int32 green1 = (global::System.Int32)(((((colorValue1) / (256L)))) % (256L));
                                        global::System.Int32 blue1 = (global::System.Int32)((((((colorValue1) / (256L)) / (256L)))) % (256L));
                                        label.BackColor = global::System.Drawing.Color.FromArgb(red1, green1, blue1);
                                    }

                                    if ((cell.Font.Color is System.DBNull))
                                    {
                                        label.ForeColor = global::System.Drawing.Color.FromArgb(0, 0, 0);
                                    }

                                    else if (Conversions.ToBoolean(!(Operators.ConditionalCompareObjectEqual(cell.Font.ColorIndex, global::Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexNone, false))))
                                    {
                                        global::System.Int64 colorValue2 = Conversions.ToLong(cell.Font.Color);
                                        global::System.Int32 red2 = (global::System.Int32)((colorValue2) % (256L));
                                        global::System.Int32 green2 = (global::System.Int32)(((((colorValue2) / (256L)))) % (256L));
                                        global::System.Int32 blue2 = (global::System.Int32)((((((colorValue2) / (256L)) / (256L)))) % (256L));
                                        label.ForeColor = global::System.Drawing.Color.FromArgb(red2, green2, blue2);
                                    }
                                }

                                this.CustomPanel2.Controls.Add(label);
                            }
                            Ordinate = ((Ordinate) + (widths2[(j) - (1)]));
                        }
                    }

                    else if (((((((X8) & !string.IsNullOrEmpty(this.TextBox2.Text)) & ((this.CanConvertToInt(this.TextBox2.Text)) == (true))))) & ((((X5) | (X6))))))
                    {

                        if (X5)
                        {
                            if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(Operators.ModObject(r, global::Microsoft.VisualBasic.Conversion.Int(this.TextBox2.Text)), 0, false)))
                            {
                                r = Conversions.ToInteger(global::Microsoft.VisualBasic.Conversion.Int(Operators.DivideObject(r, global::Microsoft.VisualBasic.Conversion.Int(this.TextBox2.Text))));
                            }
                            else
                            {
                                r = Conversions.ToInteger(Operators.AddObject(global::Microsoft.VisualBasic.Conversion.Int(Operators.DivideObject(r, global::Microsoft.VisualBasic.Conversion.Int(this.TextBox2.Text))), 1));
                            }
                            c = Conversions.ToInteger(global::Microsoft.VisualBasic.Conversion.Int(this.TextBox2.Text));

                            if ((((r) > (1)) & ((r) <= (6))))
                            {
                                height = ((global::System.Double)(this.CustomPanel2.Height) / (global::System.Double)(r));
                            }
                            else
                            {
                                height = ((global::System.Double)(this.CustomPanel2.Height) / (6d));
                            }

                            var Values = new global::System.Object[(r), (c)];
                            var References = new global::System.Int32[(r), (c)];

                            for (global::System.Int32 i = 1, loopTo20 = r; i <= loopTo20; i++)
                            {
                                for (global::System.Int32 j = 1, loopTo21 = c; j <= loopTo21; j++)
                                {
                                    global::System.Int32 x;
                                    global::System.Int32 y;
                                    x = (((((c) * ((((i) - (1))))))) + (j));
                                    y = 1;
                                    if (((x) <= (displayRng.Rows.Count)))
                                    {
                                        Values[(i) - (1), (j) - (1)] = displayRng.Cells[(global::System.Object)x, (global::System.Object)y].Value;
                                    }
                                    else
                                    {
                                        Values[(i) - (1), (j) - (1)] = "";
                                    }
                                    References[(i) - (1), (j) - (1)] = x;
                                }
                            }

                            var ColumnValues = new global::System.Object[(r)];
                            var widths2 = new global::System.Double[(c)];

                            for (global::System.Int32 j = 0, loopTo22 = (c) - (1); j <= loopTo22; j++)
                            {
                                for (global::System.Int32 i = 0, loopTo23 = (r) - (1); i <= loopTo23; i++)
                                    ColumnValues[i] = Values[i, j];
                                widths2[j] = Conversions.ToDouble(Operators.DivideObject((Operators.MultiplyObject(this.MaxOfArray(ColumnValues), BaseWidth)), 10));
                            }
                            widths2 = (global::System.Double[])this.AdjustWidth(widths2, (global::System.Object)this.CustomPanel2.Width);

                            Ordinate = 0d;

                            for (global::System.Int32 j = 1, loopTo24 = c; j <= loopTo24; j++)
                            {
                                for (global::System.Int32 i = 1, loopTo25 = r; i <= loopTo25; i++)
                                {
                                    var label = new global::System.Windows.Forms.Label();
                                    label.Text = Conversions.ToString(Values[(i) - (1), (j) - (1)]);
                                    label.Location = new global::System.Drawing.Point((global::System.Int32)Math.Round(Ordinate), (global::System.Int32)Math.Round((global::System.Double)((((i) - (1)))) * (height)));
                                    label.Height = (global::System.Int32)Math.Round(height);
                                    label.Width = (global::System.Int32)Math.Round(widths2[(j) - (1)]);
                                    label.BorderStyle = global::System.Windows.Forms.BorderStyle.FixedSingle;
                                    label.TextAlign = global::System.Drawing.ContentAlignment.MiddleCenter;

                                    if (((this.CheckBox1.Checked) == (true)))
                                    {
                                        global::System.Int32 x;
                                        global::System.Int32 y;
                                        x = References[(i) - (1), (j) - (1)];
                                        y = 1;
                                        global::Microsoft.Office.Interop.Excel.Range cell = (global::Microsoft.Office.Interop.Excel.Range)displayRng.Cells[(global::System.Object)x, (global::System.Object)y];
                                        var font = cell.Font;
                                        var fontStyle = global::System.Drawing.FontStyle.Regular;
                                        if (Conversions.ToBoolean(cell.Font.Bold))
                                            fontStyle = (fontStyle | global::System.Drawing.FontStyle.Bold);
                                        if (Conversions.ToBoolean(cell.Font.Italic))
                                            fontStyle = (fontStyle | global::System.Drawing.FontStyle.Italic);

                                        global::System.Single fontSize = global::System.Convert.ToSingle(font.Size);

                                        label.Font = new global::System.Drawing.Font(font.ToString(), fontSize, fontStyle);
                                        if (Conversions.ToBoolean(!(Operators.ConditionalCompareObjectEqual(cell.Interior.ColorIndex, global::Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexNone, false))))
                                        {
                                            global::System.Int64 colorValue1 = Conversions.ToLong(cell.Interior.Color);
                                            global::System.Int32 red1 = (global::System.Int32)((colorValue1) % (256L));
                                            global::System.Int32 green1 = (global::System.Int32)(((((colorValue1) / (256L)))) % (256L));
                                            global::System.Int32 blue1 = (global::System.Int32)((((((colorValue1) / (256L)) / (256L)))) % (256L));
                                            label.BackColor = global::System.Drawing.Color.FromArgb(red1, green1, blue1);
                                        }

                                        if ((cell.Font.Color is System.DBNull))
                                        {
                                            label.ForeColor = global::System.Drawing.Color.FromArgb(0, 0, 0);
                                        }

                                        else if (Conversions.ToBoolean(!(Operators.ConditionalCompareObjectEqual(cell.Font.ColorIndex, global::Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexNone, false))))
                                        {
                                            global::System.Int64 colorValue2 = Conversions.ToLong(cell.Font.Color);
                                            global::System.Int32 red2 = (global::System.Int32)((colorValue2) % (256L));
                                            global::System.Int32 green2 = (global::System.Int32)(((((colorValue2) / (256L)))) % (256L));
                                            global::System.Int32 blue2 = (global::System.Int32)((((((colorValue2) / (256L)) / (256L)))) % (256L));
                                            label.ForeColor = global::System.Drawing.Color.FromArgb(red2, green2, blue2);
                                        }
                                    }

                                    this.CustomPanel2.Controls.Add(label);
                                }
                                Ordinate = ((Ordinate) + (widths2[(j) - (1)]));
                            }
                        }

                        else if (X6)
                        {
                            if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(Operators.ModObject(r, global::Microsoft.VisualBasic.Conversion.Int(this.TextBox2.Text)), 0, false)))
                            {
                                c = Conversions.ToInteger(global::Microsoft.VisualBasic.Conversion.Int(Operators.DivideObject(r, global::Microsoft.VisualBasic.Conversion.Int(this.TextBox2.Text))));
                            }
                            else
                            {
                                c = Conversions.ToInteger(Operators.AddObject(global::Microsoft.VisualBasic.Conversion.Int(Operators.DivideObject(r, global::Microsoft.VisualBasic.Conversion.Int(this.TextBox2.Text))), 1));
                            }
                            r = Conversions.ToInteger(global::Microsoft.VisualBasic.Conversion.Int(this.TextBox2.Text));

                            if ((((r) > (1)) & ((r) <= (6))))
                            {
                                height = ((global::System.Double)(this.CustomPanel2.Height) / (global::System.Double)(r));
                            }
                            else
                            {
                                height = ((global::System.Double)(this.CustomPanel2.Height) / (6d));
                            }

                            var Values = new global::System.Object[(r), (c)];
                            var References = new global::System.Int32[(r), (c)];

                            for (global::System.Int32 i = 1, loopTo26 = r; i <= loopTo26; i++)
                            {
                                for (global::System.Int32 j = 1, loopTo27 = c; j <= loopTo27; j++)
                                {
                                    global::System.Int32 x;
                                    global::System.Int32 y;
                                    x = (((((r) * ((((j) - (1))))))) + (i));
                                    y = 1;
                                    if (((x) <= (displayRng.Rows.Count)))
                                    {
                                        Values[(i) - (1), (j) - (1)] = displayRng.Cells[(global::System.Object)x, (global::System.Object)y].Value;
                                    }
                                    else
                                    {
                                        Values[(i) - (1), (j) - (1)] = "";
                                    }
                                    References[(i) - (1), (j) - (1)] = x;
                                }
                            }

                            var ColumnValues = new global::System.Object[(r)];
                            var widths2 = new global::System.Double[(c)];

                            for (global::System.Int32 j = 0, loopTo28 = (c) - (1); j <= loopTo28; j++)
                            {
                                for (global::System.Int32 i = 0, loopTo29 = (r) - (1); i <= loopTo29; i++)
                                    ColumnValues[i] = Values[i, j];
                                widths2[j] = Conversions.ToDouble(Operators.DivideObject((Operators.MultiplyObject(this.MaxOfArray(ColumnValues), BaseWidth)), 10));
                            }
                            widths2 = (global::System.Double[])this.AdjustWidth(widths2, (global::System.Object)this.CustomPanel2.Width);

                            Ordinate = 0d;

                            for (global::System.Int32 j = 1, loopTo30 = c; j <= loopTo30; j++)
                            {
                                for (global::System.Int32 i = 1, loopTo31 = r; i <= loopTo31; i++)
                                {
                                    var label = new global::System.Windows.Forms.Label();
                                    label.Text = Conversions.ToString(Values[(i) - (1), (j) - (1)]);
                                    label.Location = new global::System.Drawing.Point((global::System.Int32)Math.Round(Ordinate), (global::System.Int32)Math.Round((global::System.Double)((((i) - (1)))) * (height)));
                                    label.Height = (global::System.Int32)Math.Round(height);
                                    label.Width = (global::System.Int32)Math.Round(widths2[(j) - (1)]);
                                    label.BorderStyle = global::System.Windows.Forms.BorderStyle.FixedSingle;
                                    label.TextAlign = global::System.Drawing.ContentAlignment.MiddleCenter;

                                    if (((this.CheckBox1.Checked) == (true)))
                                    {
                                        global::System.Int32 x;
                                        global::System.Int32 y;
                                        x = References[(i) - (1), (j) - (1)];
                                        y = 1;
                                        global::Microsoft.Office.Interop.Excel.Range cell = (global::Microsoft.Office.Interop.Excel.Range)displayRng.Cells[(global::System.Object)x, (global::System.Object)y];
                                        var font = cell.Font;
                                        var fontStyle = global::System.Drawing.FontStyle.Regular;
                                        if (Conversions.ToBoolean(cell.Font.Bold))
                                            fontStyle = (fontStyle | global::System.Drawing.FontStyle.Bold);
                                        if (Conversions.ToBoolean(cell.Font.Italic))
                                            fontStyle = (fontStyle | global::System.Drawing.FontStyle.Italic);

                                        global::System.Single fontSize = global::System.Convert.ToSingle(font.Size);

                                        label.Font = new global::System.Drawing.Font(font.ToString(), fontSize, fontStyle);
                                        if (Conversions.ToBoolean(!(Operators.ConditionalCompareObjectEqual(cell.Interior.ColorIndex, global::Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexNone, false))))
                                        {
                                            global::System.Int64 colorValue1 = Conversions.ToLong(cell.Interior.Color);
                                            global::System.Int32 red1 = (global::System.Int32)((colorValue1) % (256L));
                                            global::System.Int32 green1 = (global::System.Int32)(((((colorValue1) / (256L)))) % (256L));
                                            global::System.Int32 blue1 = (global::System.Int32)((((((colorValue1) / (256L)) / (256L)))) % (256L));
                                            label.BackColor = global::System.Drawing.Color.FromArgb(red1, green1, blue1);
                                        }

                                        if ((cell.Font.Color is System.DBNull))
                                        {
                                            label.ForeColor = global::System.Drawing.Color.FromArgb(0, 0, 0);
                                        }

                                        else if (Conversions.ToBoolean(!(Operators.ConditionalCompareObjectEqual(cell.Font.ColorIndex, global::Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexNone, false))))
                                        {
                                            global::System.Int64 colorValue2 = Conversions.ToLong(cell.Font.Color);
                                            global::System.Int32 red2 = (global::System.Int32)((colorValue2) % (256L));
                                            global::System.Int32 green2 = (global::System.Int32)(((((colorValue2) / (256L)))) % (256L));
                                            global::System.Int32 blue2 = (global::System.Int32)((((((colorValue2) / (256L)) / (256L)))) % (256L));
                                            label.ForeColor = global::System.Drawing.Color.FromArgb(red2, green2, blue2);
                                        }
                                    }

                                    this.CustomPanel2.Controls.Add(label);
                                }
                                Ordinate = ((Ordinate) + (widths2[(j) - (1)]));
                            }

                        }
                    }

                    this.CustomPanel2.AutoScroll = true;

                }

                if (X4)
                {

                    if (((X7) & ((((X5) | (X6))))))
                    {

                        global::System.Int32[] BreakPoints;
                        BreakPoints = (global::System.Int32[])this.GetBreakPoints(displayRng, 1);

                        global::System.Int32[] lengths;
                        lengths = (global::System.Int32[])this.GetLengths(BreakPoints);

                        if (X5)
                        {
                            r = ((global::Microsoft.VisualBasic.Information.UBound(BreakPoints)) + (1));
                            c = Conversions.ToInteger(this.MaxValue(lengths));
                        }
                        else if (X6)
                        {
                            c = ((global::Microsoft.VisualBasic.Information.UBound(BreakPoints)) + (1));
                            r = Conversions.ToInteger(this.MaxValue(lengths));
                        }

                        if ((((r) > (1)) & ((r) <= (6))))
                        {
                            height = ((global::System.Double)(this.CustomPanel2.Height) / (global::System.Double)(r));
                        }
                        else
                        {
                            height = ((global::System.Double)(this.CustomPanel2.Height) / (6d));
                        }

                        var Values = new global::System.Object[(r), (c)];
                        var References = new global::System.Int32[(r), (c)];

                        if (X5)
                        {
                            global::System.Int32 iColumn;
                            iColumn = 0;
                            for (global::System.Int32 i = 1, loopTo32 = r; i <= loopTo32; i++)
                            {
                                for (global::System.Int32 j = 1, loopTo33 = c; j <= loopTo33; j++)
                                {
                                    global::System.Int32 x;
                                    global::System.Int32 y;
                                    x = 1;
                                    y = ((iColumn) + (j));
                                    if (((x) <= (BreakPoints[(i) - (1)])))
                                    {
                                        Values[(i) - (1), (j) - (1)] = displayRng.Cells[(global::System.Object)x, (global::System.Object)y].Value;
                                    }
                                    else
                                    {
                                        Values[(i) - (1), (j) - (1)] = "";
                                    }
                                    References[(i) - (1), (j) - (1)] = y;
                                }
                                iColumn = BreakPoints[(i) - (1)];
                            }
                        }

                        else if (X6)
                        {
                            global::System.Int32 iColumn;
                            iColumn = 0;
                            for (global::System.Int32 j = 1, loopTo34 = c; j <= loopTo34; j++)
                            {
                                for (global::System.Int32 i = 1, loopTo35 = r; i <= loopTo35; i++)
                                {
                                    global::System.Int32 x;
                                    global::System.Int32 y;
                                    x = 1;
                                    y = ((iColumn) + (i));
                                    if (((x) <= (BreakPoints[(j) - (1)])))
                                    {
                                        Values[(i) - (1), (j) - (1)] = displayRng.Cells[(global::System.Object)x, (global::System.Object)y].Value;
                                    }
                                    else
                                    {
                                        Values[(i) - (1), (j) - (1)] = "";
                                    }
                                    References[(i) - (1), (j) - (1)] = y;
                                }
                                iColumn = BreakPoints[(j) - (1)];
                            }
                        }

                        var ColumnValues = new global::System.Object[(r)];
                        var widths2 = new global::System.Double[(c)];

                        for (global::System.Int32 j = 0, loopTo36 = (c) - (1); j <= loopTo36; j++)
                        {
                            for (global::System.Int32 i = 0, loopTo37 = (r) - (1); i <= loopTo37; i++)
                                ColumnValues[i] = Values[i, j];
                            widths2[j] = Conversions.ToDouble(Operators.DivideObject((Operators.MultiplyObject(this.MaxOfArray(ColumnValues), BaseWidth)), 10));
                        }
                        widths2 = (global::System.Double[])this.AdjustWidth(widths2, (global::System.Object)this.CustomPanel2.Width);

                        Ordinate = 0d;

                        for (global::System.Int32 j = 1, loopTo38 = c; j <= loopTo38; j++)
                        {
                            for (global::System.Int32 i = 1, loopTo39 = r; i <= loopTo39; i++)
                            {
                                var label = new global::System.Windows.Forms.Label();
                                label.Text = Conversions.ToString(Values[(i) - (1), (j) - (1)]);
                                label.Location = new global::System.Drawing.Point((global::System.Int32)Math.Round(Ordinate), (global::System.Int32)Math.Round((global::System.Double)((((i) - (1)))) * (height)));
                                label.Height = (global::System.Int32)Math.Round(height);
                                label.Width = (global::System.Int32)Math.Round(widths2[(j) - (1)]);
                                label.BorderStyle = global::System.Windows.Forms.BorderStyle.FixedSingle;
                                label.TextAlign = global::System.Drawing.ContentAlignment.MiddleCenter;

                                if (((this.CheckBox1.Checked) == (true)))
                                {
                                    global::System.Int32 x;
                                    global::System.Int32 y;
                                    x = 1;
                                    y = References[(i) - (1), (j) - (1)];
                                    global::Microsoft.Office.Interop.Excel.Range cell = (global::Microsoft.Office.Interop.Excel.Range)displayRng.Cells[(global::System.Object)x, (global::System.Object)y];
                                    var font = cell.Font;
                                    var fontStyle = global::System.Drawing.FontStyle.Regular;
                                    if (Conversions.ToBoolean(cell.Font.Bold))
                                        fontStyle = (fontStyle | global::System.Drawing.FontStyle.Bold);
                                    if (Conversions.ToBoolean(cell.Font.Italic))
                                        fontStyle = (fontStyle | global::System.Drawing.FontStyle.Italic);

                                    global::System.Single fontSize = global::System.Convert.ToSingle(font.Size);

                                    label.Font = new global::System.Drawing.Font(font.ToString(), fontSize, fontStyle);
                                    if (Conversions.ToBoolean(!(Operators.ConditionalCompareObjectEqual(cell.Interior.ColorIndex, global::Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexNone, false))))
                                    {
                                        global::System.Int64 colorValue1 = Conversions.ToLong(cell.Interior.Color);
                                        global::System.Int32 red1 = (global::System.Int32)((colorValue1) % (256L));
                                        global::System.Int32 green1 = (global::System.Int32)(((((colorValue1) / (256L)))) % (256L));
                                        global::System.Int32 blue1 = (global::System.Int32)((((((colorValue1) / (256L)) / (256L)))) % (256L));
                                        label.BackColor = global::System.Drawing.Color.FromArgb(red1, green1, blue1);
                                    }

                                    if ((cell.Font.Color is System.DBNull))
                                    {
                                        label.ForeColor = global::System.Drawing.Color.FromArgb(0, 0, 0);
                                    }

                                    else if (Conversions.ToBoolean(!(Operators.ConditionalCompareObjectEqual(cell.Font.ColorIndex, global::Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexNone, false))))
                                    {
                                        global::System.Int64 colorValue2 = Conversions.ToLong(cell.Font.Color);
                                        global::System.Int32 red2 = (global::System.Int32)((colorValue2) % (256L));
                                        global::System.Int32 green2 = (global::System.Int32)(((((colorValue2) / (256L)))) % (256L));
                                        global::System.Int32 blue2 = (global::System.Int32)((((((colorValue2) / (256L)) / (256L)))) % (256L));
                                        label.ForeColor = global::System.Drawing.Color.FromArgb(red2, green2, blue2);
                                    }
                                }

                                this.CustomPanel2.Controls.Add(label);
                            }
                            Ordinate = ((Ordinate) + (widths2[(j) - (1)]));
                        }
                    }


                    else if (((((((X8) & !string.IsNullOrEmpty(this.TextBox2.Text)) & ((this.CanConvertToInt(this.TextBox2.Text)) == (true))))) & ((((X5) | (X6))))))
                    {

                        if (X5)
                        {
                            if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(Operators.ModObject(c, global::Microsoft.VisualBasic.Conversion.Int(this.TextBox2.Text)), 0, false)))
                            {
                                r = Conversions.ToInteger(global::Microsoft.VisualBasic.Conversion.Int(Operators.DivideObject(c, global::Microsoft.VisualBasic.Conversion.Int(this.TextBox2.Text))));
                            }
                            else
                            {
                                r = Conversions.ToInteger(Operators.AddObject(global::Microsoft.VisualBasic.Conversion.Int(Operators.DivideObject(c, global::Microsoft.VisualBasic.Conversion.Int(this.TextBox2.Text))), 1));
                            }
                            c = Conversions.ToInteger(global::Microsoft.VisualBasic.Conversion.Int(this.TextBox2.Text));

                            if ((((r) > (1)) & ((r) <= (6))))
                            {
                                height = ((global::System.Double)(this.CustomPanel2.Height) / (global::System.Double)(r));
                            }
                            else
                            {
                                height = ((global::System.Double)(this.CustomPanel2.Height) / (6d));
                            }

                            var Values = new global::System.Object[(r), (c)];
                            var References = new global::System.Int32[(r), (c)];

                            for (global::System.Int32 i = 1, loopTo40 = r; i <= loopTo40; i++)
                            {
                                for (global::System.Int32 j = 1, loopTo41 = c; j <= loopTo41; j++)
                                {
                                    global::System.Int32 x;
                                    global::System.Int32 y;
                                    x = 1;
                                    y = (((((c) * ((((i) - (1))))))) + (j));
                                    if (((x) <= (displayRng.Rows.Count)))
                                    {
                                        Values[(i) - (1), (j) - (1)] = displayRng.Cells[(global::System.Object)x, (global::System.Object)y].Value;
                                    }
                                    else
                                    {
                                        Values[(i) - (1), (j) - (1)] = "";
                                    }
                                    References[(i) - (1), (j) - (1)] = y;
                                }
                            }

                            var ColumnValues = new global::System.Object[(r)];
                            var widths2 = new global::System.Double[(c)];

                            for (global::System.Int32 j = 0, loopTo42 = (c) - (1); j <= loopTo42; j++)
                            {
                                for (global::System.Int32 i = 0, loopTo43 = (r) - (1); i <= loopTo43; i++)
                                    ColumnValues[i] = Values[i, j];
                                widths2[j] = Conversions.ToDouble(Operators.DivideObject((Operators.MultiplyObject(this.MaxOfArray(ColumnValues), BaseWidth)), 10));
                            }
                            widths2 = (global::System.Double[])this.AdjustWidth(widths2, (global::System.Object)this.CustomPanel2.Width);

                            Ordinate = 0d;

                            for (global::System.Int32 j = 1, loopTo44 = c; j <= loopTo44; j++)
                            {
                                for (global::System.Int32 i = 1, loopTo45 = r; i <= loopTo45; i++)
                                {
                                    var label = new global::System.Windows.Forms.Label();
                                    label.Text = Conversions.ToString(Values[(i) - (1), (j) - (1)]);
                                    label.Location = new global::System.Drawing.Point((global::System.Int32)Math.Round(Ordinate), (global::System.Int32)Math.Round((global::System.Double)((((i) - (1)))) * (height)));
                                    label.Height = (global::System.Int32)Math.Round(height);
                                    label.Width = (global::System.Int32)Math.Round(widths2[(j) - (1)]);
                                    label.BorderStyle = global::System.Windows.Forms.BorderStyle.FixedSingle;
                                    label.TextAlign = global::System.Drawing.ContentAlignment.MiddleCenter;

                                    if (((this.CheckBox1.Checked) == (true)))
                                    {
                                        global::System.Int32 x;
                                        global::System.Int32 y;
                                        x = 1;
                                        y = References[(i) - (1), (j) - (1)];
                                        global::Microsoft.Office.Interop.Excel.Range cell = (global::Microsoft.Office.Interop.Excel.Range)displayRng.Cells[(global::System.Object)x, (global::System.Object)y];
                                        var font = cell.Font;
                                        var fontStyle = global::System.Drawing.FontStyle.Regular;
                                        if (Conversions.ToBoolean(cell.Font.Bold))
                                            fontStyle = (fontStyle | global::System.Drawing.FontStyle.Bold);
                                        if (Conversions.ToBoolean(cell.Font.Italic))
                                            fontStyle = (fontStyle | global::System.Drawing.FontStyle.Italic);

                                        global::System.Single fontSize = global::System.Convert.ToSingle(font.Size);

                                        label.Font = new global::System.Drawing.Font(font.ToString(), fontSize, fontStyle);
                                        if (Conversions.ToBoolean(!(Operators.ConditionalCompareObjectEqual(cell.Interior.ColorIndex, global::Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexNone, false))))
                                        {
                                            global::System.Int64 colorValue1 = Conversions.ToLong(cell.Interior.Color);
                                            global::System.Int32 red1 = (global::System.Int32)((colorValue1) % (256L));
                                            global::System.Int32 green1 = (global::System.Int32)(((((colorValue1) / (256L)))) % (256L));
                                            global::System.Int32 blue1 = (global::System.Int32)((((((colorValue1) / (256L)) / (256L)))) % (256L));
                                            label.BackColor = global::System.Drawing.Color.FromArgb(red1, green1, blue1);
                                        }

                                        if ((cell.Font.Color is System.DBNull))
                                        {
                                            label.ForeColor = global::System.Drawing.Color.FromArgb(0, 0, 0);
                                        }

                                        else if (Conversions.ToBoolean(!(Operators.ConditionalCompareObjectEqual(cell.Font.ColorIndex, global::Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexNone, false))))
                                        {
                                            global::System.Int64 colorValue2 = Conversions.ToLong(cell.Font.Color);
                                            global::System.Int32 red2 = (global::System.Int32)((colorValue2) % (256L));
                                            global::System.Int32 green2 = (global::System.Int32)(((((colorValue2) / (256L)))) % (256L));
                                            global::System.Int32 blue2 = (global::System.Int32)((((((colorValue2) / (256L)) / (256L)))) % (256L));
                                            label.ForeColor = global::System.Drawing.Color.FromArgb(red2, green2, blue2);
                                        }
                                    }

                                    this.CustomPanel2.Controls.Add(label);
                                }
                                Ordinate = ((Ordinate) + (widths2[(j) - (1)]));
                            }
                        }

                        else if (X6)
                        {
                            if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(Operators.ModObject(c, global::Microsoft.VisualBasic.Conversion.Int(this.TextBox2.Text)), 0, false)))
                            {
                                c = Conversions.ToInteger(global::Microsoft.VisualBasic.Conversion.Int(Operators.DivideObject(c, global::Microsoft.VisualBasic.Conversion.Int(this.TextBox2.Text))));
                            }
                            else
                            {
                                c = Conversions.ToInteger(Operators.AddObject(global::Microsoft.VisualBasic.Conversion.Int(Operators.DivideObject(c, global::Microsoft.VisualBasic.Conversion.Int(this.TextBox2.Text))), 1));
                            }
                            r = Conversions.ToInteger(global::Microsoft.VisualBasic.Conversion.Int(this.TextBox2.Text));

                            if ((((r) > (1)) & ((r) <= (6))))
                            {
                                height = ((global::System.Double)(this.CustomPanel2.Height) / (global::System.Double)(r));
                            }
                            else
                            {
                                height = ((global::System.Double)(this.CustomPanel2.Height) / (6d));
                            }

                            var Values = new global::System.Object[(r), (c)];
                            var References = new global::System.Int32[(r), (c)];

                            for (global::System.Int32 i = 1, loopTo46 = r; i <= loopTo46; i++)
                            {
                                for (global::System.Int32 j = 1, loopTo47 = c; j <= loopTo47; j++)
                                {
                                    global::System.Int32 x;
                                    global::System.Int32 y;
                                    x = 1;
                                    y = (((((r) * ((((j) - (1))))))) + (i));
                                    if (((x) <= (displayRng.Rows.Count)))
                                    {
                                        Values[(i) - (1), (j) - (1)] = displayRng.Cells[(global::System.Object)x, (global::System.Object)y].Value;
                                    }
                                    else
                                    {
                                        Values[(i) - (1), (j) - (1)] = "";
                                    }
                                    References[(i) - (1), (j) - (1)] = y;
                                }
                            }

                            var ColumnValues = new global::System.Object[(r)];
                            var widths2 = new global::System.Double[(c)];

                            for (global::System.Int32 j = 0, loopTo48 = (c) - (1); j <= loopTo48; j++)
                            {
                                for (global::System.Int32 i = 0, loopTo49 = (r) - (1); i <= loopTo49; i++)
                                    ColumnValues[i] = Values[i, j];
                                widths2[j] = Conversions.ToDouble(Operators.DivideObject((Operators.MultiplyObject(this.MaxOfArray(ColumnValues), BaseWidth)), 10));
                            }
                            widths2 = (global::System.Double[])this.AdjustWidth(widths2, (global::System.Object)this.CustomPanel2.Width);

                            Ordinate = 0d;

                            for (global::System.Int32 j = 1, loopTo50 = c; j <= loopTo50; j++)
                            {
                                for (global::System.Int32 i = 1, loopTo51 = r; i <= loopTo51; i++)
                                {
                                    var label = new global::System.Windows.Forms.Label();
                                    label.Text = Conversions.ToString(Values[(i) - (1), (j) - (1)]);
                                    label.Location = new global::System.Drawing.Point((global::System.Int32)Math.Round(Ordinate), (global::System.Int32)Math.Round((global::System.Double)((((i) - (1)))) * (height)));
                                    label.Height = (global::System.Int32)Math.Round(height);
                                    label.Width = (global::System.Int32)Math.Round(widths2[(j) - (1)]);
                                    label.BorderStyle = global::System.Windows.Forms.BorderStyle.FixedSingle;
                                    label.TextAlign = global::System.Drawing.ContentAlignment.MiddleCenter;

                                    if (((this.CheckBox1.Checked) == (true)))
                                    {
                                        global::System.Int32 x;
                                        global::System.Int32 y;
                                        x = 1;
                                        y = References[(i) - (1), (j) - (1)];
                                        global::Microsoft.Office.Interop.Excel.Range cell = (global::Microsoft.Office.Interop.Excel.Range)displayRng.Cells[(global::System.Object)x, (global::System.Object)y];
                                        var font = cell.Font;
                                        var fontStyle = global::System.Drawing.FontStyle.Regular;
                                        if (Conversions.ToBoolean(cell.Font.Bold))
                                            fontStyle = (fontStyle | global::System.Drawing.FontStyle.Bold);
                                        if (Conversions.ToBoolean(cell.Font.Italic))
                                            fontStyle = (fontStyle | global::System.Drawing.FontStyle.Italic);

                                        global::System.Single fontSize = global::System.Convert.ToSingle(font.Size);

                                        label.Font = new global::System.Drawing.Font(font.ToString(), fontSize, fontStyle);
                                        if (Conversions.ToBoolean(!(Operators.ConditionalCompareObjectEqual(cell.Interior.ColorIndex, global::Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexNone, false))))
                                        {
                                            global::System.Int64 colorValue1 = Conversions.ToLong(cell.Interior.Color);
                                            global::System.Int32 red1 = (global::System.Int32)((colorValue1) % (256L));
                                            global::System.Int32 green1 = (global::System.Int32)(((((colorValue1) / (256L)))) % (256L));
                                            global::System.Int32 blue1 = (global::System.Int32)((((((colorValue1) / (256L)) / (256L)))) % (256L));
                                            label.BackColor = global::System.Drawing.Color.FromArgb(red1, green1, blue1);
                                        }

                                        if ((cell.Font.Color is System.DBNull))
                                        {
                                            label.ForeColor = global::System.Drawing.Color.FromArgb(0, 0, 0);
                                        }

                                        else if (Conversions.ToBoolean(!(Operators.ConditionalCompareObjectEqual(cell.Font.ColorIndex, global::Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexNone, false))))
                                        {
                                            global::System.Int64 colorValue2 = Conversions.ToLong(cell.Font.Color);
                                            global::System.Int32 red2 = (global::System.Int32)((colorValue2) % (256L));
                                            global::System.Int32 green2 = (global::System.Int32)(((((colorValue2) / (256L)))) % (256L));
                                            global::System.Int32 blue2 = (global::System.Int32)((((((colorValue2) / (256L)) / (256L)))) % (256L));
                                            label.ForeColor = global::System.Drawing.Color.FromArgb(red2, green2, blue2);
                                        }
                                    }

                                    this.CustomPanel2.Controls.Add(label);
                                }
                                Ordinate = ((Ordinate) + (widths2[(j) - (1)]));
                            }

                        }
                    }

                    this.CustomPanel2.AutoScroll = true;

                }

                this.TextBoxChanged = false;
            }

            catch (global::System.Exception ex)
            {

            }

        }
        private void Form7_Load(global::System.Object sender, global::System.EventArgs e)
        {

            try
            {

                this.excelApp = global::VSTO_Addins.Globals.ThisAddIn.Application;
                this.workbook = excelApp.ActiveWorkbook;
                this.workbook2 = excelApp.ActiveWorkbook;
                this.worksheet = (global::Microsoft.Office.Interop.Excel.Worksheet)workbook.ActiveSheet;
                this.worksheet2 = (global::Microsoft.Office.Interop.Excel.Worksheet)workbook2.ActiveSheet;
                this.KeyPreview = true;

                this.excelApp.SheetSelectionChange += this.excelApp_SheetSelectionChange;

                this.opened = ((this.opened) + (1));
            }

            catch (global::System.Exception ex)
            {

            }

        }

        private void excelApp_SheetSelectionChange(global::System.Object Sh, global::Microsoft.Office.Interop.Excel.Range Target)
        {

            try
            {

                this.excelApp = global::VSTO_Addins.Globals.ThisAddIn.Application;
                global::Microsoft.Office.Interop.Excel.Range selectedRange;
                selectedRange = (global::Microsoft.Office.Interop.Excel.Range)excelApp.Selection;

                if (((this.TextBoxChanged) == (false)))
                {
                    if (((this.FocusedTextBox) == (1)))
                    {
                        this.worksheet = (global::Microsoft.Office.Interop.Excel.Worksheet)workbook.ActiveSheet;
                        if (((worksheet.Name ?? "") != (OpenSheet.Name ?? "")))
                        {
                            this.TextBox1.Text = ((worksheet.Name + "!") + selectedRange.get_Address());
                        }
                        else
                        {
                            this.TextBox1.Text = selectedRange.get_Address();
                        }
                        this.rng = selectedRange;
                        this.TextBox1.Focus();
                    }

                    else if (((this.FocusedTextBox) == (3)))
                    {
                        this.worksheet2 = (global::Microsoft.Office.Interop.Excel.Worksheet)workbook2.ActiveSheet;
                        if (((worksheet2.Name ?? "") != (OpenSheet.Name ?? "")))
                        {
                            this.TextBox3.Text = ((worksheet2.Name + "!") + selectedRange.get_Address());
                        }
                        else
                        {
                            this.TextBox3.Text = selectedRange.get_Address();
                        }
                        this.rng2 = selectedRange;
                        this.TextBox3.Focus();
                    }
                }
            }

            catch (global::System.Exception ex)
            {

            }

        }
        private void Button2_Click(global::System.Object sender, global::System.EventArgs e)
        {

            try
            {

                this.TextBoxChanged = true;
                if (string.IsNullOrEmpty(this.TextBox1.Text))
                {
                    global::System.Windows.Forms.MessageBox.Show("Select a Source Range.", "Error", global::System.Windows.Forms.MessageBoxButtons.OK, global::System.Windows.Forms.MessageBoxIcon.Error);
                    this.TextBox1.Focus();
                    this.worksheet.Activate();
                    this.rng.Select();
                    return;
                }

                if (((this.IsValidExcelCellReference(this.TextBox1.Text)) == (false)))
                {
                    global::System.Windows.Forms.MessageBox.Show("Select a Valid Source Range.", "Error", global::System.Windows.Forms.MessageBoxButtons.OK, global::System.Windows.Forms.MessageBoxIcon.Error);
                    this.TextBox1.Focus();
                    this.worksheet.Activate();
                    this.rng.Select();
                    return;
                }

                if (((((((this.RadioButton1.Checked) == (false)) & ((this.RadioButton2.Checked) == (false))) & ((this.RadioButton3.Checked) == (false))) & ((this.RadioButton4.Checked) == (false)))))
                {
                    global::System.Windows.Forms.MessageBox.Show("Select a Transformation Type.", "Error", global::System.Windows.Forms.MessageBoxButtons.OK, global::System.Windows.Forms.MessageBoxIcon.Error);
                    this.worksheet.Activate();
                    this.rng.Select();
                    return;
                }

                if (((((this.RadioButton5.Checked) == (false)) & ((this.RadioButton6.Checked) == (false)))))
                {
                    global::System.Windows.Forms.MessageBox.Show("Select a Transformation Option.", "Error", global::System.Windows.Forms.MessageBoxButtons.OK, global::System.Windows.Forms.MessageBoxIcon.Error);
                    this.worksheet.Activate();
                    this.rng.Select();
                    return;
                }

                if (((((((this.RadioButton3.Checked) == (true)) | ((this.RadioButton4.Checked) == (true))))) & (((((this.RadioButton7.Checked) == (false)) & ((this.RadioButton8.Checked) == (false)))))))
                {
                    global::System.Windows.Forms.MessageBox.Show("Select a Separator.", "Error", global::System.Windows.Forms.MessageBoxButtons.OK, global::System.Windows.Forms.MessageBoxIcon.Error);
                    this.worksheet.Activate();
                    this.rng.Select();
                    return;
                }

                if ((((this.RadioButton8.Checked) == (true)) & (((string.IsNullOrEmpty(this.TextBox2.Text) | ((this.CanConvertToInt(this.TextBox2.Text)) == (false)))))))
                {
                    global::System.String[] Texts;
                    Texts = global::Microsoft.VisualBasic.Strings.Split(this.RadioButton8.Text, " ");
                    global::System.String iText = Texts[global::Microsoft.VisualBasic.Information.UBound(Texts)];
                    global::System.Windows.Forms.MessageBox.Show(("Enter a valid Number of " + iText) + ".", "Error", global::System.Windows.Forms.MessageBoxButtons.OK, global::System.Windows.Forms.MessageBoxIcon.Error);
                    this.worksheet.Activate();
                    this.rng.Select();
                    this.TextBox2.Focus();
                    return;
                }

                if (((((this.RadioButton9.Checked) == (false)) & ((this.RadioButton10.Checked) == (false)))))
                {
                    global::System.Windows.Forms.MessageBox.Show("Select a Destination Range.", "Error", global::System.Windows.Forms.MessageBoxButtons.OK, global::System.Windows.Forms.MessageBoxIcon.Error);
                    this.worksheet.Activate();
                    this.rng.Select();
                    return;
                }

                if ((((this.RadioButton10.Checked) == (true)) & (((string.IsNullOrEmpty(this.TextBox3.Text) | ((this.IsValidExcelCellReference(this.TextBox3.Text)) == (false)))))))
                {
                    global::System.Windows.Forms.MessageBox.Show("Enter a valid Destination Cell.", "Error", global::System.Windows.Forms.MessageBoxButtons.OK, global::System.Windows.Forms.MessageBoxIcon.Error);
                    this.worksheet.Activate();
                    this.rng.Select();
                    return;
                }

                if (((this.CheckBox2.Checked) == (true)))
                {
                    this.worksheet.Copy(After: workbook.Sheets[worksheet.Name]);
                }

                workbook.Sheets[worksheet.Name].Activate();

                global::System.Boolean X1;
                X1 = this.RadioButton1.Checked;

                global::System.Boolean X2;
                X2 = this.RadioButton2.Checked;

                global::System.Boolean X3;
                X3 = this.RadioButton3.Checked;

                global::System.Boolean X4;
                X4 = this.RadioButton4.Checked;

                global::System.Boolean X5;
                X5 = this.RadioButton5.Checked;

                global::System.Boolean X6;
                X6 = this.RadioButton6.Checked;

                global::System.Boolean X7;
                X7 = this.RadioButton7.Checked;

                global::System.Boolean X8;
                X8 = this.RadioButton8.Checked;

                global::System.Int32 r;
                global::System.Int32 c;

                r = this.rng.Rows.Count;
                c = this.rng.Columns.Count;

                global::System.Int32 i;
                global::System.Int32 j;

                if (X1)
                {

                    this.rng2 = this.worksheet2.get_Range(this.rng2.Cells[(global::System.Object)1, (global::System.Object)1], this.rng2.Cells[(global::System.Object)((r) * (c)), (global::System.Object)1]);
                    global::System.String rng2Address = this.rng2.get_Address();
                    this.worksheet2.Activate();
                    this.rng2.Select();

                    if (((this.Overlap(this.excelApp, this.worksheet, this.worksheet2, this.rng, this.rng2)) == (false)))
                    {

                        this.rng2.ClearFormats();

                        global::System.Int32 count;
                        count = 1;

                        if (X5)
                        {
                            var loopTo = r;
                            for (i = 1; i <= loopTo; i++)
                            {
                                var loopTo1 = c;
                                for (j = 1; j <= loopTo1; j++)
                                {
                                    global::System.Int32 x = count;
                                    global::System.Int32 y = 1;

                                    if (((this.CheckBox1.Checked) == (false)))
                                    {
                                        this.rng2.Cells[(global::System.Object)x, (global::System.Object)y].Value = this.rng.Cells[(global::System.Object)i, (global::System.Object)j].Value;
                                        count = ((count) + (1));
                                    }

                                    else if (((this.CheckBox1.Checked) == (true)))
                                    {
                                        this.rng.Cells[(global::System.Object)i, (global::System.Object)j].Copy();
                                        this.rng2.Cells[(global::System.Object)x, (global::System.Object)y].PasteSpecial(global::Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues);
                                        this.rng2 = this.worksheet2.get_Range(rng2Address);
                                        this.rng2.Cells[(global::System.Object)x, (global::System.Object)y].PasteSpecial(global::Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats);
                                        this.rng2 = this.worksheet2.get_Range(rng2Address);
                                        count = ((count) + (1));

                                    }

                                }
                            }

                            excelApp.CutCopyMode = global::Microsoft.Office.Interop.Excel.XlCutCopyMode.xlCopy;
                        }

                        else if (X6)
                        {

                            var loopTo2 = c;
                            for (j = 1; j <= loopTo2; j++)
                            {
                                var loopTo3 = r;
                                for (i = 1; i <= loopTo3; i++)
                                {

                                    global::System.Int32 x = count;
                                    global::System.Int32 y = 1;

                                    if (((this.CheckBox1.Checked) == (false)))
                                    {
                                        this.rng2.Cells[(global::System.Object)x, (global::System.Object)y].Value = this.rng.Cells[(global::System.Object)i, (global::System.Object)j].Value;
                                        count = ((count) + (1));
                                    }

                                    else if (((this.CheckBox1.Checked) == (true)))
                                    {
                                        this.rng.Cells[(global::System.Object)i, (global::System.Object)j].Copy();
                                        this.rng2.Cells[(global::System.Object)x, (global::System.Object)y].PasteSpecial(global::Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues);
                                        this.rng2 = this.worksheet2.get_Range(rng2Address);
                                        this.rng2.Cells[(global::System.Object)x, (global::System.Object)y].PasteSpecial(global::Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats);
                                        this.rng2 = this.worksheet2.get_Range(rng2Address);
                                        count = ((count) + (1));

                                    }

                                }
                            }
                            excelApp.CutCopyMode = global::Microsoft.Office.Interop.Excel.XlCutCopyMode.xlCopy;
                        }
                        else
                        {
                            global::System.Windows.Forms.MessageBox.Show("Choose One Transformation Option. ", "Error", global::System.Windows.Forms.MessageBoxButtons.OK, global::System.Windows.Forms.MessageBoxIcon.Error);
                            return;
                        }

                        if (((this.CheckBox1.Checked) == (true)))
                        {
                            global::Microsoft.Office.Interop.Excel.Range TopCell = (global::Microsoft.Office.Interop.Excel.Range)this.rng.Cells[(global::System.Object)1, (global::System.Object)1];

                            if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)8].LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                            {
                                var loopTo4 = this.rng2.Columns.Count;
                                for (j = 1; j <= loopTo4; j++)
                                {
                                    this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)8).LineStyle = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)8].LineStyle;
                                    this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)8).Color = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)8].Color;
                                    this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)8).Weight = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)8].Weight;
                                }
                            }
                            else
                            {
                                var loopTo5 = this.rng2.Columns.Count;
                                for (j = 1; j <= loopTo5; j++)
                                    this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)8).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                            }

                            if (((this.rng2.Rows.Count) > (1)))
                            {
                                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                                {
                                    var loopTo6 = this.rng2.Columns.Count;
                                    for (j = 1; j <= loopTo6; j++)
                                    {
                                        this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)9).LineStyle = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].LineStyle;
                                        this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)9).Color = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].Color;
                                        this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)9).Weight = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].Weight;
                                    }
                                }
                                else
                                {
                                    var loopTo7 = this.rng2.Columns.Count;
                                    for (j = 1; j <= loopTo7; j++)
                                        this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)9).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                }
                            }

                            if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)7].LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                            {
                                var loopTo8 = this.rng2.Rows.Count;
                                for (i = 1; i <= loopTo8; i++)
                                {
                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)1].Borders((global::System.Object)7).LineStyle = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)7].LineStyle;
                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)1].Borders((global::System.Object)7).Color = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)7].Color;
                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)1].Borders((global::System.Object)7).Weight = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)7].Weight;
                                }
                            }
                            else
                            {
                                var loopTo9 = this.rng2.Rows.Count;
                                for (i = 1; i <= loopTo9; i++)
                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)1].Borders((global::System.Object)7).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                            }

                            global::Microsoft.Office.Interop.Excel.Range BottomCell = (global::Microsoft.Office.Interop.Excel.Range)this.rng.Cells[(global::System.Object)r, (global::System.Object)c];

                            if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(BottomCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                            {
                                var loopTo10 = this.rng2.Columns.Count;
                                for (j = 1; j <= loopTo10; j++)
                                {
                                    this.rng2.Cells[(global::System.Object)this.rng2.Rows.Count, (global::System.Object)j].Borders((global::System.Object)9).LineStyle = BottomCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].LineStyle;
                                    this.rng2.Cells[(global::System.Object)this.rng2.Rows.Count, (global::System.Object)j].Borders((global::System.Object)9).Color = BottomCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].Color;
                                    this.rng2.Cells[(global::System.Object)this.rng2.Rows.Count, (global::System.Object)j].Borders((global::System.Object)9).Weight = BottomCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].Weight;
                                }
                            }
                            else
                            {
                                var loopTo11 = this.rng2.Columns.Count;
                                for (j = 1; j <= loopTo11; j++)
                                    this.rng2.Cells[(global::System.Object)this.rng2.Rows.Count, (global::System.Object)this.rng2.Columns.Count].Borders((global::System.Object)9).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                            }

                            if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(BottomCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)10].LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                            {
                                var loopTo12 = this.rng2.Rows.Count;
                                for (i = 1; i <= loopTo12; i++)
                                {
                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)this.rng2.Columns.Count].Borders((global::System.Object)10).LineStyle = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)10].LineStyle;
                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)this.rng2.Columns.Count].Borders((global::System.Object)10).Color = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)10].Color;
                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)this.rng2.Columns.Count].Borders((global::System.Object)10).Weight = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)10].Weight;
                                }
                            }
                            else
                            {
                                var loopTo13 = this.rng2.Rows.Count;
                                for (i = 1; i <= loopTo13; i++)
                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)this.rng2.Columns.Count].Borders((global::System.Object)10).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                            }

                            if (((this.rng.Rows.Count) > (1)))
                            {
                                global::Microsoft.Office.Interop.Excel.Range MiddleCell = (global::Microsoft.Office.Interop.Excel.Range)this.rng.Cells[(global::System.Object)2, (global::System.Object)1];
                                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(MiddleCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                                {
                                    var loopTo14 = (this.rng2.Rows.Count) - (1);
                                    for (i = 2; i <= loopTo14; i++)
                                    {
                                        var loopTo15 = this.rng2.Columns.Count;
                                        for (j = 1; j <= loopTo15; j++)
                                        {
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)9).LineStyle = MiddleCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].LineStyle;
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)9).Color = MiddleCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].Color;
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)9).Weight = MiddleCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].Weight;
                                        }
                                    }
                                }
                                else
                                {
                                    var loopTo16 = (this.rng2.Rows.Count) - (1);
                                    for (i = 2; i <= loopTo16; i++)
                                    {
                                        var loopTo17 = this.rng2.Columns.Count;
                                        for (j = 1; j <= loopTo17; j++)
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)9).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                    }
                                }
                            }

                            if (((this.rng.Columns.Count) > (1)))
                            {
                                global::Microsoft.Office.Interop.Excel.Range MiddleCell = (global::Microsoft.Office.Interop.Excel.Range)this.rng.Cells[(global::System.Object)1, (global::System.Object)2];
                                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(MiddleCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)10].LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                                {
                                    var loopTo18 = (this.rng2.Columns.Count) - (1);
                                    for (j = 1; j <= loopTo18; j++)
                                    {
                                        var loopTo19 = this.rng2.Rows.Count;
                                        for (i = 1; i <= loopTo19; i++)
                                        {
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)10).LineStyle = MiddleCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)10].LineStyle;
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)10).Color = MiddleCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)10].Color;
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)10).Weight = MiddleCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)10].Weight;
                                        }
                                    }
                                }
                                else
                                {
                                    var loopTo20 = (this.rng2.Columns.Count) - (1);
                                    for (j = 1; j <= loopTo20; j++)
                                    {
                                        var loopTo21 = this.rng2.Rows.Count;
                                        for (i = 1; i <= loopTo21; i++)
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)10).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                    }
                                }
                            }
                        }
                    }

                    else
                    {

                        var Arr = new global::System.Object[(r), (c)];
                        var Bolds = new global::System.Boolean[(r), (c)];
                        var Italics = new global::System.Boolean[(r), (c)];
                        var fontNames = new global::System.String[(r), (c)];
                        var fontSizes = new global::System.Single[(r), (c)];
                        var reds1 = new global::System.Int32[(r), (c)];
                        var reds2 = new global::System.Int32[(r), (c)];
                        var greens1 = new global::System.Int32[(r), (c)];
                        var greens2 = new global::System.Int32[(r), (c)];
                        var blues1 = new global::System.Int32[(r), (c)];
                        var blues2 = new global::System.Int32[(r), (c)];

                        global::System.Boolean TopBorder7;
                        global::System.Object TopBorder7L;
                        global::System.Object TopBorder7C;
                        global::System.Object TopBorder7W;

                        global::System.Boolean TopBorder8;
                        global::System.Object TopBorder8L;
                        global::System.Object TopBorder8C;
                        global::System.Object TopBorder8W;

                        global::System.Boolean TopBorder9;
                        global::System.Object TopBorder9L;
                        global::System.Object TopBorder9C;
                        global::System.Object TopBorder9W;

                        global::System.Boolean BottomBorder9;
                        global::System.Object BottomBorder9L;
                        global::System.Object BottomBorder9C;
                        global::System.Object BottomBorder9W;

                        global::System.Boolean BottomBorder10;
                        global::System.Object BottomBorder10L;
                        global::System.Object BottomBorder10C;
                        global::System.Object BottomBorder10W;

                        global::System.Boolean MiddleBorder9;
                        global::System.Object MiddleBorder9L;
                        global::System.Object MiddleBorder9C;
                        global::System.Object MiddleBorder9W;

                        global::System.Boolean MiddleBorder10;
                        global::System.Object MiddleBorder10L;
                        global::System.Object MiddleBorder10C;
                        global::System.Object MiddleBorder10W;

                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)7).LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                        {
                            TopBorder7 = true;
                            TopBorder7L = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)7).LineStyle;
                            TopBorder7C = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)7).Color;
                            TopBorder7W = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)7).Weight;
                        }
                        else
                        {
                            TopBorder7 = false;
                            TopBorder7L = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)7).LineStyle;
                            TopBorder7C = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)7).Color;
                            TopBorder7W = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)7).Weight;
                        }

                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)8).LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                        {
                            TopBorder8 = true;
                            TopBorder8L = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)8).LineStyle;
                            TopBorder8C = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)8).Color;
                            TopBorder8W = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)8).Weight;
                        }
                        else
                        {
                            TopBorder8 = false;
                            TopBorder8L = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)8).LineStyle;
                            TopBorder8C = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)8).Color;
                            TopBorder8W = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)8).Weight;
                        }

                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)9).LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                        {
                            TopBorder9 = true;
                            TopBorder9L = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)9).LineStyle;
                            TopBorder9C = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)9).Color;
                            TopBorder9W = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)9).Weight;
                        }
                        else
                        {
                            TopBorder9 = false;
                            TopBorder9L = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)9).LineStyle;
                            TopBorder9C = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)9).Color;
                            TopBorder9W = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)9).Weight;
                        }

                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)9).LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                        {
                            BottomBorder9 = true;
                            BottomBorder9L = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)9).LineStyle;
                            BottomBorder9C = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)9).Color;
                            BottomBorder9W = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)9).Weight;
                        }
                        else
                        {
                            BottomBorder9 = false;
                            BottomBorder9L = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)9).LineStyle;
                            BottomBorder9C = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)9).Color;
                            BottomBorder9W = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)9).Weight;
                        }

                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)10).LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                        {
                            BottomBorder10 = true;
                            BottomBorder10L = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)10).LineStyle;
                            BottomBorder10C = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)10).Color;
                            BottomBorder10W = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)10).Weight;
                        }
                        else
                        {
                            BottomBorder10 = false;
                            BottomBorder10L = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)10).LineStyle;
                            BottomBorder10C = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)10).Color;
                            BottomBorder10W = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)10).Weight;
                        }

                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(this.rng.Cells[(global::System.Object)2, (global::System.Object)1].Borders((global::System.Object)9).LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                        {
                            MiddleBorder9 = true;
                            MiddleBorder9L = this.rng.Cells[(global::System.Object)2, (global::System.Object)1].Borders((global::System.Object)9).LineStyle;
                            MiddleBorder9C = this.rng.Cells[(global::System.Object)2, (global::System.Object)1].Borders((global::System.Object)9).Color;
                            MiddleBorder9W = this.rng.Cells[(global::System.Object)2, (global::System.Object)1].Borders((global::System.Object)9).Weight;
                        }
                        else
                        {
                            MiddleBorder9 = false;
                            MiddleBorder9L = this.rng.Cells[(global::System.Object)2, (global::System.Object)1].Borders((global::System.Object)9).LineStyle;
                            MiddleBorder9C = this.rng.Cells[(global::System.Object)2, (global::System.Object)1].Borders((global::System.Object)9).Color;
                            MiddleBorder9W = this.rng.Cells[(global::System.Object)2, (global::System.Object)1].Borders((global::System.Object)9).Weight;
                        }

                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(this.rng.Cells[(global::System.Object)1, (global::System.Object)2].Borders((global::System.Object)10).LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                        {
                            MiddleBorder10 = true;
                            MiddleBorder10L = this.rng.Cells[(global::System.Object)1, (global::System.Object)2].Borders((global::System.Object)10).LineStyle;
                            MiddleBorder10C = this.rng.Cells[(global::System.Object)1, (global::System.Object)2].Borders((global::System.Object)10).Color;
                            MiddleBorder10W = this.rng.Cells[(global::System.Object)1, (global::System.Object)2].Borders((global::System.Object)10).Weight;
                        }
                        else
                        {
                            MiddleBorder10 = false;
                            MiddleBorder10L = this.rng.Cells[(global::System.Object)1, (global::System.Object)2].Borders((global::System.Object)10).LineStyle;
                            MiddleBorder10C = this.rng.Cells[(global::System.Object)1, (global::System.Object)2].Borders((global::System.Object)10).Color;
                            MiddleBorder10W = this.rng.Cells[(global::System.Object)1, (global::System.Object)2].Borders((global::System.Object)10).Weight;
                        }

                        var loopTo22 = r;
                        for (i = 1; i <= loopTo22; i++)
                        {
                            var loopTo23 = c;
                            for (j = 1; j <= loopTo23; j++)
                            {
                                Arr[(i) - (1), (j) - (1)] = this.rng.Cells[(global::System.Object)i, (global::System.Object)j].Value;

                                if (((this.CheckBox1.Checked) == (true)))
                                {

                                    global::Microsoft.Office.Interop.Excel.Range cell = (global::Microsoft.Office.Interop.Excel.Range)this.rng.Cells[(global::System.Object)i, (global::System.Object)j];
                                    var font = cell.Font;

                                    Bolds[(i) - (1), (j) - (1)] = Conversions.ToBoolean(cell.Font.Bold);
                                    Italics[(i) - (1), (j) - (1)] = Conversions.ToBoolean(cell.Font.Italic);

                                    if ((((font.Name is System.DBNull)) == (false)))
                                    {
                                        fontNames[(i) - (1), (j) - (1)] = Conversions.ToString(font.Name);
                                    }
                                    else
                                    {
                                        fontNames[(i) - (1), (j) - (1)] = "Calibri";
                                    }

                                    if ((((font.Size is System.DBNull)) == (false)))
                                    {
                                        global::System.Single fontSize = global::System.Convert.ToSingle(font.Size);
                                        fontSizes[(i) - (1), (j) - (1)] = fontSize;
                                    }
                                    else
                                    {
                                        fontSizes[(i) - (1), (j) - (1)] = 11f;
                                    }

                                    if ((cell.Interior.Color is System.DBNull))
                                    {
                                        reds1[(i) - (1), (j) - (1)] = 255;
                                        greens1[(i) - (1), (j) - (1)] = 255;
                                        blues1[(i) - (1), (j) - (1)] = 255;
                                    }
                                    else
                                    {
                                        global::System.Int64 colorValue1 = Conversions.ToLong(cell.Interior.Color);
                                        global::System.Int32 red1 = (global::System.Int32)((colorValue1) % (256L));
                                        global::System.Int32 green1 = (global::System.Int32)(((((colorValue1) / (256L)))) % (256L));
                                        global::System.Int32 blue1 = (global::System.Int32)((((((colorValue1) / (256L)) / (256L)))) % (256L));
                                        reds1[(i) - (1), (j) - (1)] = red1;
                                        greens1[(i) - (1), (j) - (1)] = green1;
                                        blues1[(i) - (1), (j) - (1)] = blue1;
                                    }

                                    if ((cell.Font.Color is System.DBNull))
                                    {
                                        reds2[(i) - (1), (j) - (1)] = 0;
                                        greens2[(i) - (1), (j) - (1)] = 0;
                                        blues2[(i) - (1), (j) - (1)] = 0;
                                    }
                                    else
                                    {
                                        global::System.Int64 colorValue2 = Conversions.ToLong(cell.Font.Color);
                                        global::System.Int32 red2 = (global::System.Int32)((colorValue2) % (256L));
                                        global::System.Int32 green2 = (global::System.Int32)(((((colorValue2) / (256L)))) % (256L));
                                        global::System.Int32 blue2 = (global::System.Int32)((((((colorValue2) / (256L)) / (256L)))) % (256L));
                                        reds2[(i) - (1), (j) - (1)] = red2;
                                        greens2[(i) - (1), (j) - (1)] = green2;
                                        blues2[(i) - (1), (j) - (1)] = blue2;
                                    }
                                }

                            }
                        }

                        this.rng.ClearContents();
                        this.rng.ClearFormats();
                        this.rng2.ClearFormats();

                        global::System.Int32 count;
                        count = 1;

                        if (X5)
                        {
                            var loopTo24 = r;
                            for (i = 1; i <= loopTo24; i++)
                            {
                                var loopTo25 = c;
                                for (j = 1; j <= loopTo25; j++)
                                {
                                    global::System.Int32 x = count;
                                    global::System.Int32 y = 1;

                                    this.rng2.Cells[(global::System.Object)x, (global::System.Object)y].Value = Arr[(i) - (1), (j) - (1)];
                                    count = ((count) + (1));

                                    if (((this.CheckBox1.Checked) == (true)))
                                    {

                                        global::Microsoft.Office.Interop.Excel.Range cell2 = (global::Microsoft.Office.Interop.Excel.Range)this.rng2.Cells[(global::System.Object)x, (global::System.Object)y];
                                        var font2 = cell2.Font;

                                        global::System.Single fontSize = fontSizes[(i) - (1), (j) - (1)];

                                        this.rng2.Cells[(global::System.Object)x, (global::System.Object)y].Font.Name = fontNames[(i) - (1), (j) - (1)];
                                        this.rng2.Cells[(global::System.Object)x, (global::System.Object)y].Font.Size = (global::System.Object)fontSizes[(i) - (1), (j) - (1)];

                                        if (Bolds[(i) - (1), (j) - (1)])
                                            this.rng2.Cells[(global::System.Object)x, (global::System.Object)y].Font.Bold = (global::System.Object)true;
                                        if (Italics[(i) - (1), (j) - (1)])
                                            this.rng2.Cells[(global::System.Object)x, (global::System.Object)y].Font.Italic = (global::System.Object)true;


                                        global::System.Int32 red1 = reds1[(i) - (1), (j) - (1)];
                                        global::System.Int32 green1 = greens1[(i) - (1), (j) - (1)];
                                        global::System.Int32 blue1 = blues1[(i) - (1), (j) - (1)];

                                        this.rng2.Cells[(global::System.Object)x, (global::System.Object)y].Interior.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(red1, green1, blue1);

                                        global::System.Int32 red2 = reds2[(i) - (1), (j) - (1)];
                                        global::System.Int32 green2 = greens2[(i) - (1), (j) - (1)];
                                        global::System.Int32 blue2 = blues2[(i) - (1), (j) - (1)];
                                        this.rng2.Cells[(global::System.Object)x, (global::System.Object)y].Font.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(red2, green2, blue2);
                                    }

                                }
                            }
                        }

                        else if (X6)
                        {

                            var loopTo26 = c;
                            for (j = 1; j <= loopTo26; j++)
                            {
                                var loopTo27 = r;
                                for (i = 1; i <= loopTo27; i++)
                                {

                                    global::System.Int32 x = count;
                                    global::System.Int32 y = 1;

                                    this.rng2.Cells[(global::System.Object)x, (global::System.Object)y].Value = Arr[(i) - (1), (j) - (1)];
                                    count = ((count) + (1));

                                    if (((this.CheckBox1.Checked) == (true)))
                                    {

                                        global::Microsoft.Office.Interop.Excel.Range cell2 = (global::Microsoft.Office.Interop.Excel.Range)this.rng2.Cells[(global::System.Object)x, (global::System.Object)y];
                                        var font2 = cell2.Font;

                                        global::System.Single fontSize = fontSizes[(i) - (1), (j) - (1)];

                                        this.rng2.Cells[(global::System.Object)x, (global::System.Object)y].Font.Name = fontNames[(i) - (1), (j) - (1)];
                                        this.rng2.Cells[(global::System.Object)x, (global::System.Object)y].Font.Size = (global::System.Object)fontSizes[(i) - (1), (j) - (1)];

                                        if (Bolds[(i) - (1), (j) - (1)])
                                            this.rng2.Cells[(global::System.Object)x, (global::System.Object)y].Font.Bold = (global::System.Object)true;
                                        if (Italics[(i) - (1), (j) - (1)])
                                            this.rng2.Cells[(global::System.Object)x, (global::System.Object)y].Font.Italic = (global::System.Object)true;

                                        global::System.Int32 red1 = reds1[(i) - (1), (j) - (1)];
                                        global::System.Int32 green1 = greens1[(i) - (1), (j) - (1)];
                                        global::System.Int32 blue1 = blues1[(i) - (1), (j) - (1)];
                                        this.rng2.Cells[(global::System.Object)x, (global::System.Object)y].Interior.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(red1, green1, blue1);

                                        global::System.Int32 red2 = reds2[(i) - (1), (j) - (1)];
                                        global::System.Int32 green2 = greens2[(i) - (1), (j) - (1)];
                                        global::System.Int32 blue2 = blues2[(i) - (1), (j) - (1)];
                                        this.rng2.Cells[(global::System.Object)x, (global::System.Object)y].Font.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(red2, green2, blue2);
                                    }

                                }
                            }
                        }

                        else
                        {
                            global::System.Windows.Forms.MessageBox.Show("Choose One Transformation Option. ", "Error", global::System.Windows.Forms.MessageBoxButtons.OK, global::System.Windows.Forms.MessageBoxIcon.Error);
                            return;
                        }

                        if (((this.CheckBox1.Checked) == (true)))
                        {

                            if (((TopBorder8) == (true)))
                            {
                                var loopTo28 = this.rng2.Columns.Count;
                                for (j = 1; j <= loopTo28; j++)
                                {
                                    this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)8).LineStyle = TopBorder8L;
                                    this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)8).Color = TopBorder8C;
                                    this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)8).Weight = TopBorder8W;
                                }
                            }
                            else
                            {
                                var loopTo29 = this.rng2.Columns.Count;
                                for (j = 1; j <= loopTo29; j++)
                                    this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)8).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                            }

                            if (((this.rng2.Rows.Count) > (1)))
                            {
                                if (((TopBorder9) == (true)))
                                {
                                    var loopTo30 = this.rng2.Columns.Count;
                                    for (j = 1; j <= loopTo30; j++)
                                    {
                                        this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)9).LineStyle = TopBorder9L;
                                        this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)9).Color = TopBorder9C;
                                        this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)9).Weight = TopBorder9W;
                                    }
                                }
                                else
                                {
                                    var loopTo31 = this.rng2.Columns.Count;
                                    for (j = 1; j <= loopTo31; j++)
                                        this.rng2.Cells[(global::System.Object)this.rng2.Rows.Count, (global::System.Object)j].Borders((global::System.Object)9).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                }
                            }

                            if (((TopBorder7) == (true)))
                            {
                                var loopTo32 = this.rng2.Rows.Count;
                                for (i = 1; i <= loopTo32; i++)
                                {
                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)1].Borders((global::System.Object)7).LineStyle = TopBorder7L;
                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)1].Borders((global::System.Object)7).Color = TopBorder7C;
                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)1].Borders((global::System.Object)7).Weight = TopBorder7W;
                                }
                            }
                            else
                            {
                                var loopTo33 = this.rng2.Rows.Count;
                                for (i = 1; i <= loopTo33; i++)
                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)1].Borders((global::System.Object)7).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                            }

                            if (((BottomBorder9) == (true)))
                            {
                                var loopTo34 = this.rng2.Columns.Count;
                                for (j = 1; j <= loopTo34; j++)
                                {
                                    this.rng2.Cells[(global::System.Object)this.rng2.Rows.Count, (global::System.Object)j].Borders((global::System.Object)9).LineStyle = BottomBorder9L;
                                    this.rng2.Cells[(global::System.Object)this.rng2.Rows.Count, (global::System.Object)j].Borders((global::System.Object)9).Color = BottomBorder9C;
                                    this.rng2.Cells[(global::System.Object)this.rng2.Rows.Count, (global::System.Object)j].Borders((global::System.Object)9).Weight = BottomBorder9W;
                                }
                            }
                            else
                            {
                                var loopTo35 = this.rng2.Columns.Count;
                                for (j = 1; j <= loopTo35; j++)
                                    this.rng2.Cells[(global::System.Object)this.rng2.Rows.Count, (global::System.Object)this.rng2.Columns.Count].Borders((global::System.Object)9).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                            }

                            if (((BottomBorder10) == (true)))
                            {
                                var loopTo36 = this.rng2.Rows.Count;
                                for (i = 1; i <= loopTo36; i++)
                                {
                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)this.rng2.Columns.Count].Borders((global::System.Object)10).LineStyle = BottomBorder10L;
                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)this.rng2.Columns.Count].Borders((global::System.Object)10).Color = BottomBorder10C;
                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)this.rng2.Columns.Count].Borders((global::System.Object)10).Weight = BottomBorder10W;
                                }
                            }
                            else
                            {
                                var loopTo37 = this.rng2.Rows.Count;
                                for (i = 1; i <= loopTo37; i++)
                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)this.rng2.Columns.Count].Borders((global::System.Object)10).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                            }

                            if (((this.rng.Rows.Count) > (1)))
                            {

                                if (((MiddleBorder9) == (true)))
                                {
                                    var loopTo38 = (this.rng2.Rows.Count) - (1);
                                    for (i = 2; i <= loopTo38; i++)
                                    {
                                        var loopTo39 = this.rng2.Columns.Count;
                                        for (j = 1; j <= loopTo39; j++)
                                        {
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)9).LineStyle = MiddleBorder9L;
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)9).Color = MiddleBorder9C;
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)9).Weight = MiddleBorder9W;
                                        }
                                    }
                                }
                                else
                                {
                                    var loopTo40 = (this.rng2.Rows.Count) - (1);
                                    for (i = 2; i <= loopTo40; i++)
                                    {
                                        var loopTo41 = this.rng2.Columns.Count;
                                        for (j = 1; j <= loopTo41; j++)
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)9).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                    }
                                }

                            }

                            if (((this.rng.Columns.Count) > (1)))
                            {

                                if (((MiddleBorder10) == (true)))
                                {
                                    var loopTo42 = (this.rng2.Columns.Count) - (1);
                                    for (j = 1; j <= loopTo42; j++)
                                    {
                                        var loopTo43 = this.rng2.Rows.Count;
                                        for (i = 1; i <= loopTo43; i++)
                                        {
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)10).LineStyle = MiddleBorder10L;
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)10).Color = MiddleBorder10C;
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)10).Weight = MiddleBorder10W;
                                        }
                                    }
                                }
                                else
                                {
                                    var loopTo44 = (this.rng2.Columns.Count) - (1);
                                    for (j = 1; j <= loopTo44; j++)
                                    {
                                        var loopTo45 = this.rng2.Rows.Count;
                                        for (i = 1; i <= loopTo45; i++)
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)10).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                    }
                                }

                            }

                        }

                    }
                }

                else if (X2)
                {

                    this.rng2 = this.worksheet2.get_Range(this.rng2.Cells[(global::System.Object)1, (global::System.Object)1], this.rng2.Cells[(global::System.Object)1, (global::System.Object)((r) * (c))]);
                    global::System.String rng2Address = this.rng2.get_Address();
                    this.worksheet2.Activate();
                    this.rng2.Select();

                    if (((this.Overlap(this.excelApp, this.worksheet, this.worksheet2, this.rng, this.rng2)) == (false)))
                    {

                        this.rng2.ClearFormats();
                        global::System.Int32 count;
                        count = 1;

                        if (X5)
                        {
                            var loopTo46 = r;
                            for (i = 1; i <= loopTo46; i++)
                            {
                                var loopTo47 = c;
                                for (j = 1; j <= loopTo47; j++)
                                {

                                    global::System.Int32 x = 1;
                                    global::System.Int32 y = count;

                                    if (((this.CheckBox1.Checked) == (false)))
                                    {
                                        this.rng2.Cells[(global::System.Object)x, (global::System.Object)y].Value = this.rng[(global::System.Object)i, (global::System.Object)j];
                                        count = ((count) + (1));
                                    }

                                    else if (((this.CheckBox1.Checked) == (true)))
                                    {
                                        this.rng.Cells[(global::System.Object)i, (global::System.Object)j].Copy();
                                        this.rng2.Cells[(global::System.Object)x, (global::System.Object)y].PasteSpecial(global::Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues);
                                        this.rng2 = this.worksheet2.get_Range(rng2Address);
                                        this.rng2.Cells[(global::System.Object)x, (global::System.Object)y].PasteSpecial(global::Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats);
                                        this.rng2 = this.worksheet2.get_Range(rng2Address);
                                        count = ((count) + (1));
                                    }

                                }
                            }
                            excelApp.CutCopyMode = global::Microsoft.Office.Interop.Excel.XlCutCopyMode.xlCopy;
                        }
                        else if (X6)
                        {

                            var loopTo48 = c;
                            for (j = 1; j <= loopTo48; j++)
                            {
                                var loopTo49 = r;
                                for (i = 1; i <= loopTo49; i++)
                                {

                                    global::System.Int32 x = 1;
                                    global::System.Int32 y = count;

                                    if (((this.CheckBox1.Checked) == (false)))
                                    {
                                        this.rng2.Cells[(global::System.Object)x, (global::System.Object)y].Value = this.rng.Cells[(global::System.Object)i, (global::System.Object)j].Value;
                                        count = ((count) + (1));
                                    }

                                    else if (((this.CheckBox1.Checked) == (true)))
                                    {
                                        this.rng.Cells[(global::System.Object)i, (global::System.Object)j].Copy();
                                        this.rng2.Cells[(global::System.Object)x, (global::System.Object)y].PasteSpecial(global::Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues);
                                        this.rng2 = this.worksheet2.get_Range(rng2Address);
                                        this.rng2.Cells[(global::System.Object)x, (global::System.Object)y].PasteSpecial(global::Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats);
                                        this.rng2 = this.worksheet2.get_Range(rng2Address);
                                        count = ((count) + (1));
                                    }

                                }
                            }
                            excelApp.CutCopyMode = global::Microsoft.Office.Interop.Excel.XlCutCopyMode.xlCopy;
                        }
                        else
                        {
                            global::System.Windows.Forms.MessageBox.Show("Choose One Transformation Option. ", "Error", global::System.Windows.Forms.MessageBoxButtons.OK, global::System.Windows.Forms.MessageBoxIcon.Error);
                            return;

                        }
                    }



                    else
                    {

                        var Arr = new global::System.Object[(r), (c)];
                        var Bolds = new global::System.Boolean[(r), (c)];
                        var Italics = new global::System.Boolean[(r), (c)];
                        var fontNames = new global::System.String[(r), (c)];
                        var fontSizes = new global::System.Single[(r), (c)];
                        var reds1 = new global::System.Int32[(r), (c)];
                        var reds2 = new global::System.Int32[(r), (c)];
                        var greens1 = new global::System.Int32[(r), (c)];
                        var greens2 = new global::System.Int32[(r), (c)];
                        var blues1 = new global::System.Int32[(r), (c)];
                        var blues2 = new global::System.Int32[(r), (c)];

                        global::System.Boolean TopBorder7;
                        global::System.Object TopBorder7L;
                        global::System.Object TopBorder7C;
                        global::System.Object TopBorder7W;

                        global::System.Boolean TopBorder8;
                        global::System.Object TopBorder8L;
                        global::System.Object TopBorder8C;
                        global::System.Object TopBorder8W;

                        global::System.Boolean TopBorder9;
                        global::System.Object TopBorder9L;
                        global::System.Object TopBorder9C;
                        global::System.Object TopBorder9W;

                        global::System.Boolean BottomBorder9;
                        global::System.Object BottomBorder9L;
                        global::System.Object BottomBorder9C;
                        global::System.Object BottomBorder9W;

                        global::System.Boolean BottomBorder10;
                        global::System.Object BottomBorder10L;
                        global::System.Object BottomBorder10C;
                        global::System.Object BottomBorder10W;

                        global::System.Boolean MiddleBorder9;
                        global::System.Object MiddleBorder9L;
                        global::System.Object MiddleBorder9C;
                        global::System.Object MiddleBorder9W;

                        global::System.Boolean MiddleBorder10;
                        global::System.Object MiddleBorder10L;
                        global::System.Object MiddleBorder10C;
                        global::System.Object MiddleBorder10W;

                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)7).LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                        {
                            TopBorder7 = true;
                            TopBorder7L = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)7).LineStyle;
                            TopBorder7C = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)7).Color;
                            TopBorder7W = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)7).Weight;
                        }
                        else
                        {
                            TopBorder7 = false;
                            TopBorder7L = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)7).LineStyle;
                            TopBorder7C = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)7).Color;
                            TopBorder7W = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)7).Weight;
                        }

                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)8).LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                        {
                            TopBorder8 = true;
                            TopBorder8L = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)8).LineStyle;
                            TopBorder8C = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)8).Color;
                            TopBorder8W = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)8).Weight;
                        }
                        else
                        {
                            TopBorder8 = false;
                            TopBorder8L = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)8).LineStyle;
                            TopBorder8C = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)8).Color;
                            TopBorder8W = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)8).Weight;
                        }

                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)9).LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                        {
                            TopBorder9 = true;
                            TopBorder9L = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)9).LineStyle;
                            TopBorder9C = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)9).Color;
                            TopBorder9W = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)9).Weight;
                        }
                        else
                        {
                            TopBorder9 = false;
                            TopBorder9L = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)9).LineStyle;
                            TopBorder9C = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)9).Color;
                            TopBorder9W = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)9).Weight;
                        }

                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)9).LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                        {
                            BottomBorder9 = true;
                            BottomBorder9L = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)9).LineStyle;
                            BottomBorder9C = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)9).Color;
                            BottomBorder9W = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)9).Weight;
                        }
                        else
                        {
                            BottomBorder9 = false;
                            BottomBorder9L = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)9).LineStyle;
                            BottomBorder9C = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)9).Color;
                            BottomBorder9W = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)9).Weight;
                        }

                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)10).LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                        {
                            BottomBorder10 = true;
                            BottomBorder10L = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)10).LineStyle;
                            BottomBorder10C = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)10).Color;
                            BottomBorder10W = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)10).Weight;
                        }
                        else
                        {
                            BottomBorder10 = false;
                            BottomBorder10L = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)10).LineStyle;
                            BottomBorder10C = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)10).Color;
                            BottomBorder10W = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)10).Weight;
                        }

                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(this.rng.Cells[(global::System.Object)2, (global::System.Object)1].Borders((global::System.Object)9).LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                        {
                            MiddleBorder9 = true;
                            MiddleBorder9L = this.rng.Cells[(global::System.Object)2, (global::System.Object)1].Borders((global::System.Object)9).LineStyle;
                            MiddleBorder9C = this.rng.Cells[(global::System.Object)2, (global::System.Object)1].Borders((global::System.Object)9).Color;
                            MiddleBorder9W = this.rng.Cells[(global::System.Object)2, (global::System.Object)1].Borders((global::System.Object)9).Weight;
                        }
                        else
                        {
                            MiddleBorder9 = false;
                            MiddleBorder9L = this.rng.Cells[(global::System.Object)2, (global::System.Object)1].Borders((global::System.Object)9).LineStyle;
                            MiddleBorder9C = this.rng.Cells[(global::System.Object)2, (global::System.Object)1].Borders((global::System.Object)9).Color;
                            MiddleBorder9W = this.rng.Cells[(global::System.Object)2, (global::System.Object)1].Borders((global::System.Object)9).Weight;
                        }

                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(this.rng.Cells[(global::System.Object)1, (global::System.Object)2].Borders((global::System.Object)10).LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                        {
                            MiddleBorder10 = true;
                            MiddleBorder10L = this.rng.Cells[(global::System.Object)1, (global::System.Object)2].Borders((global::System.Object)10).LineStyle;
                            MiddleBorder10C = this.rng.Cells[(global::System.Object)1, (global::System.Object)2].Borders((global::System.Object)10).Color;
                            MiddleBorder10W = this.rng.Cells[(global::System.Object)1, (global::System.Object)2].Borders((global::System.Object)10).Weight;
                        }
                        else
                        {
                            MiddleBorder10 = false;
                            MiddleBorder10L = this.rng.Cells[(global::System.Object)1, (global::System.Object)2].Borders((global::System.Object)10).LineStyle;
                            MiddleBorder10C = this.rng.Cells[(global::System.Object)1, (global::System.Object)2].Borders((global::System.Object)10).Color;
                            MiddleBorder10W = this.rng.Cells[(global::System.Object)1, (global::System.Object)2].Borders((global::System.Object)10).Weight;
                        }

                        var loopTo50 = r;
                        for (i = 1; i <= loopTo50; i++)
                        {
                            var loopTo51 = c;
                            for (j = 1; j <= loopTo51; j++)
                            {
                                Arr[(i) - (1), (j) - (1)] = this.rng.Cells[(global::System.Object)i, (global::System.Object)j].Value;

                                if (((this.CheckBox1.Checked) == (true)))
                                {

                                    global::Microsoft.Office.Interop.Excel.Range cell = (global::Microsoft.Office.Interop.Excel.Range)this.rng.Cells[(global::System.Object)i, (global::System.Object)j];
                                    var font = cell.Font;

                                    Bolds[(i) - (1), (j) - (1)] = Conversions.ToBoolean(cell.Font.Bold);
                                    Italics[(i) - (1), (j) - (1)] = Conversions.ToBoolean(cell.Font.Italic);


                                    if ((((font.Name is System.DBNull)) == (false)))
                                    {
                                        fontNames[(i) - (1), (j) - (1)] = Conversions.ToString(font.Name);
                                    }
                                    else
                                    {
                                        fontNames[(i) - (1), (j) - (1)] = "Calibri";
                                    }

                                    if ((((font.Size is System.DBNull)) == (false)))
                                    {
                                        global::System.Single fontSize = global::System.Convert.ToSingle(font.Size);
                                        fontSizes[(i) - (1), (j) - (1)] = fontSize;
                                    }
                                    else
                                    {
                                        fontSizes[(i) - (1), (j) - (1)] = 11f;
                                    }

                                    if ((cell.Interior.Color is System.DBNull))
                                    {
                                        reds1[(i) - (1), (j) - (1)] = 255;
                                        greens1[(i) - (1), (j) - (1)] = 255;
                                        blues1[(i) - (1), (j) - (1)] = 255;
                                    }
                                    else
                                    {
                                        global::System.Int64 colorValue1 = Conversions.ToLong(cell.Interior.Color);
                                        global::System.Int32 red1 = (global::System.Int32)((colorValue1) % (256L));
                                        global::System.Int32 green1 = (global::System.Int32)(((((colorValue1) / (256L)))) % (256L));
                                        global::System.Int32 blue1 = (global::System.Int32)((((((colorValue1) / (256L)) / (256L)))) % (256L));
                                        reds1[(i) - (1), (j) - (1)] = red1;
                                        greens1[(i) - (1), (j) - (1)] = green1;
                                        blues1[(i) - (1), (j) - (1)] = blue1;
                                    }

                                    if ((cell.Font.Color is System.DBNull))
                                    {
                                        reds2[(i) - (1), (j) - (1)] = 0;
                                        greens2[(i) - (1), (j) - (1)] = 0;
                                        blues2[(i) - (1), (j) - (1)] = 0;
                                    }
                                    else
                                    {
                                        global::System.Int64 colorValue2 = Conversions.ToLong(cell.Font.Color);
                                        global::System.Int32 red2 = (global::System.Int32)((colorValue2) % (256L));
                                        global::System.Int32 green2 = (global::System.Int32)(((((colorValue2) / (256L)))) % (256L));
                                        global::System.Int32 blue2 = (global::System.Int32)((((((colorValue2) / (256L)) / (256L)))) % (256L));
                                        reds2[(i) - (1), (j) - (1)] = red2;
                                        greens2[(i) - (1), (j) - (1)] = green2;
                                        blues2[(i) - (1), (j) - (1)] = blue2;
                                    }
                                }

                            }
                        }

                        this.rng.ClearContents();
                        this.rng.ClearFormats();

                        this.rng2.ClearFormats();

                        global::System.Int32 count;
                        count = 1;

                        if (X5)
                        {

                            var loopTo52 = r;
                            for (i = 1; i <= loopTo52; i++)
                            {
                                var loopTo53 = c;
                                for (j = 1; j <= loopTo53; j++)
                                {
                                    global::System.Int32 x = 1;
                                    global::System.Int32 y = count;

                                    this.rng2.Cells[(global::System.Object)x, (global::System.Object)y].Value = Arr[(i) - (1), (j) - (1)];
                                    count = ((count) + (1));

                                    if (((this.CheckBox1.Checked) == (true)))
                                    {

                                        global::Microsoft.Office.Interop.Excel.Range cell2 = (global::Microsoft.Office.Interop.Excel.Range)this.rng2.Cells[(global::System.Object)x, (global::System.Object)y];
                                        var font2 = cell2.Font;

                                        global::System.Single fontSize = fontSizes[(i) - (1), (j) - (1)];

                                        this.rng2.Cells[(global::System.Object)x, (global::System.Object)y].Font.Name = fontNames[(i) - (1), (j) - (1)];
                                        this.rng2.Cells[(global::System.Object)x, (global::System.Object)y].Font.Size = (global::System.Object)fontSizes[(i) - (1), (j) - (1)];

                                        if (Bolds[(i) - (1), (j) - (1)])
                                            this.rng2.Cells[(global::System.Object)x, (global::System.Object)y].Font.Bold = (global::System.Object)true;
                                        if (Italics[(i) - (1), (j) - (1)])
                                            this.rng2.Cells[(global::System.Object)x, (global::System.Object)y].Font.Italic = (global::System.Object)true;

                                        global::System.Int32 red1 = reds1[(i) - (1), (j) - (1)];
                                        global::System.Int32 green1 = greens1[(i) - (1), (j) - (1)];
                                        global::System.Int32 blue1 = blues1[(i) - (1), (j) - (1)];
                                        this.rng2.Cells[(global::System.Object)x, (global::System.Object)y].Interior.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(red1, green1, blue1);

                                        global::System.Int32 red2 = reds2[(i) - (1), (j) - (1)];
                                        global::System.Int32 green2 = greens2[(i) - (1), (j) - (1)];
                                        global::System.Int32 blue2 = blues2[(i) - (1), (j) - (1)];
                                        this.rng2.Cells[(global::System.Object)x, (global::System.Object)y].Font.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(red2, green2, blue2);
                                    }

                                }
                            }
                        }

                        else if (X6)
                        {

                            var loopTo54 = c;
                            for (j = 1; j <= loopTo54; j++)
                            {
                                var loopTo55 = r;
                                for (i = 1; i <= loopTo55; i++)
                                {

                                    global::System.Int32 x = 1;
                                    global::System.Int32 y = count;

                                    this.rng2.Cells[(global::System.Object)x, (global::System.Object)y].Value = Arr[(i) - (1), (j) - (1)];
                                    count = ((count) + (1));

                                    if (((this.CheckBox1.Checked) == (true)))
                                    {

                                        global::Microsoft.Office.Interop.Excel.Range cell2 = (global::Microsoft.Office.Interop.Excel.Range)this.rng2.Cells[(global::System.Object)x, (global::System.Object)y];
                                        var font2 = cell2.Font;

                                        global::System.Single fontSize = fontSizes[(i) - (1), (j) - (1)];

                                        this.rng2.Cells[(global::System.Object)x, (global::System.Object)y].Font.Name = fontNames[(i) - (1), (j) - (1)];
                                        this.rng2.Cells[(global::System.Object)x, (global::System.Object)y].Font.Size = (global::System.Object)fontSizes[(i) - (1), (j) - (1)];

                                        if (Bolds[(i) - (1), (j) - (1)])
                                            this.rng2.Cells[(global::System.Object)x, (global::System.Object)y].Font.Bold = (global::System.Object)true;
                                        if (Italics[(i) - (1), (j) - (1)])
                                            this.rng2.Cells[(global::System.Object)x, (global::System.Object)y].Font.Italic = (global::System.Object)true;

                                        global::System.Int32 red1 = reds1[(i) - (1), (j) - (1)];
                                        global::System.Int32 green1 = greens1[(i) - (1), (j) - (1)];
                                        global::System.Int32 blue1 = blues1[(i) - (1), (j) - (1)];
                                        this.rng2.Cells[(global::System.Object)x, (global::System.Object)y].Interior.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(red1, green1, blue1);

                                        global::System.Int32 red2 = reds2[(i) - (1), (j) - (1)];
                                        global::System.Int32 green2 = greens2[(i) - (1), (j) - (1)];
                                        global::System.Int32 blue2 = blues2[(i) - (1), (j) - (1)];
                                        this.rng2.Cells[(global::System.Object)x, (global::System.Object)y].Font.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(red2, green2, blue2);
                                    }

                                }
                            }
                        }

                        else
                        {
                            global::System.Windows.Forms.MessageBox.Show("Choose One Transformation Option. ", "Error", global::System.Windows.Forms.MessageBoxButtons.OK, global::System.Windows.Forms.MessageBoxIcon.Error);
                            return;

                        }

                        if (((this.CheckBox1.Checked) == (true)))
                        {

                            if (((TopBorder8) == (true)))
                            {
                                var loopTo56 = this.rng2.Columns.Count;
                                for (j = 1; j <= loopTo56; j++)
                                {
                                    this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)8).LineStyle = TopBorder8L;
                                    this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)8).Color = TopBorder8C;
                                    this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)8).Weight = TopBorder8W;
                                }
                            }
                            else
                            {
                                var loopTo57 = this.rng2.Columns.Count;
                                for (j = 1; j <= loopTo57; j++)
                                    this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)8).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                            }

                            if (((this.rng2.Rows.Count) > (1)))
                            {
                                if (((TopBorder9) == (true)))
                                {
                                    var loopTo58 = this.rng2.Columns.Count;
                                    for (j = 1; j <= loopTo58; j++)
                                    {
                                        this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)9).LineStyle = TopBorder9L;
                                        this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)9).Color = TopBorder9C;
                                        this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)9).Weight = TopBorder9W;
                                    }
                                }
                                else
                                {
                                    var loopTo59 = this.rng2.Columns.Count;
                                    for (j = 1; j <= loopTo59; j++)
                                        this.rng2.Cells[(global::System.Object)this.rng2.Rows.Count, (global::System.Object)j].Borders((global::System.Object)9).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                }
                            }

                            if (((TopBorder7) == (true)))
                            {
                                var loopTo60 = this.rng2.Rows.Count;
                                for (i = 1; i <= loopTo60; i++)
                                {
                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)1].Borders((global::System.Object)7).LineStyle = TopBorder7L;
                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)1].Borders((global::System.Object)7).Color = TopBorder7C;
                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)1].Borders((global::System.Object)7).Weight = TopBorder7W;
                                }
                            }
                            else
                            {
                                var loopTo61 = this.rng2.Rows.Count;
                                for (i = 1; i <= loopTo61; i++)
                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)1].Borders((global::System.Object)7).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                            }

                            if (((BottomBorder9) == (true)))
                            {
                                var loopTo62 = this.rng2.Columns.Count;
                                for (j = 1; j <= loopTo62; j++)
                                {
                                    this.rng2.Cells[(global::System.Object)this.rng2.Rows.Count, (global::System.Object)j].Borders((global::System.Object)9).LineStyle = BottomBorder9L;
                                    this.rng2.Cells[(global::System.Object)this.rng2.Rows.Count, (global::System.Object)j].Borders((global::System.Object)9).Color = BottomBorder9C;
                                    this.rng2.Cells[(global::System.Object)this.rng2.Rows.Count, (global::System.Object)j].Borders((global::System.Object)9).Weight = BottomBorder9W;
                                }
                            }
                            else
                            {
                                var loopTo63 = this.rng2.Columns.Count;
                                for (j = 1; j <= loopTo63; j++)
                                    this.rng2.Cells[(global::System.Object)this.rng2.Rows.Count, (global::System.Object)this.rng2.Columns.Count].Borders((global::System.Object)9).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                            }

                            if (((BottomBorder10) == (true)))
                            {
                                var loopTo64 = this.rng2.Rows.Count;
                                for (i = 1; i <= loopTo64; i++)
                                {
                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)this.rng2.Columns.Count].Borders((global::System.Object)10).LineStyle = BottomBorder10L;
                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)this.rng2.Columns.Count].Borders((global::System.Object)10).Color = BottomBorder10C;
                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)this.rng2.Columns.Count].Borders((global::System.Object)10).Weight = BottomBorder10W;
                                }
                            }
                            else
                            {
                                var loopTo65 = this.rng2.Rows.Count;
                                for (i = 1; i <= loopTo65; i++)
                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)this.rng2.Columns.Count].Borders((global::System.Object)10).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                            }

                            if (((this.rng.Rows.Count) > (1)))
                            {

                                if (((MiddleBorder9) == (true)))
                                {
                                    var loopTo66 = (this.rng2.Rows.Count) - (1);
                                    for (i = 2; i <= loopTo66; i++)
                                    {
                                        var loopTo67 = this.rng2.Columns.Count;
                                        for (j = 1; j <= loopTo67; j++)
                                        {
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)9).LineStyle = MiddleBorder9L;
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)9).Color = MiddleBorder9C;
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)9).Weight = MiddleBorder9W;
                                        }
                                    }
                                }
                                else
                                {
                                    var loopTo68 = (this.rng2.Rows.Count) - (1);
                                    for (i = 2; i <= loopTo68; i++)
                                    {
                                        var loopTo69 = this.rng2.Columns.Count;
                                        for (j = 1; j <= loopTo69; j++)
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)9).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                    }
                                }

                            }

                            if (((this.rng.Columns.Count) > (1)))
                            {

                                if (((MiddleBorder10) == (true)))
                                {
                                    var loopTo70 = (this.rng2.Columns.Count) - (1);
                                    for (j = 1; j <= loopTo70; j++)
                                    {
                                        var loopTo71 = this.rng2.Rows.Count;
                                        for (i = 1; i <= loopTo71; i++)
                                        {
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)10).LineStyle = MiddleBorder10L;
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)10).Color = MiddleBorder10C;
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)10).Weight = MiddleBorder10W;
                                        }
                                    }
                                }
                                else
                                {
                                    var loopTo72 = (this.rng2.Columns.Count) - (1);
                                    for (j = 1; j <= loopTo72; j++)
                                    {
                                        var loopTo73 = this.rng2.Rows.Count;
                                        for (i = 1; i <= loopTo73; i++)
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)10).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                    }
                                }

                            }

                        }

                    }
                }

                else if (X3)
                {

                    if (((X7) & ((((X5) | (X6))))))
                    {

                        global::System.Int32[] BreakPoints;
                        BreakPoints = (global::System.Int32[])this.GetBreakPoints(this.rng, 2);

                        global::System.Int32[] lengths;
                        lengths = (global::System.Int32[])this.GetLengths(BreakPoints);

                        var r2 = default(global::System.Int32);
                        var c2 = default(global::System.Int32);

                        if (X5)
                        {
                            r2 = ((global::Microsoft.VisualBasic.Information.UBound(BreakPoints)) + (1));
                            c2 = Conversions.ToInteger(this.MaxValue(lengths));
                        }
                        else if (X6)
                        {
                            c2 = ((global::Microsoft.VisualBasic.Information.UBound(BreakPoints)) + (1));
                            r2 = Conversions.ToInteger(this.MaxValue(lengths));
                        }

                        this.rng2 = this.worksheet2.get_Range(this.rng2.Cells[(global::System.Object)1, (global::System.Object)1], this.rng2.Cells[(global::System.Object)r2, (global::System.Object)c2]);
                        global::System.String rng2Address = this.rng2.get_Address();
                        this.worksheet2.Activate();
                        this.rng2.Select();

                        if (((this.Overlap(this.excelApp, this.worksheet, this.worksheet2, this.rng, this.rng2)) == (false)))
                        {

                            this.rng2.ClearFormats();

                            if (X5)
                            {
                                global::System.Int32 iRow;
                                iRow = 0;
                                var loopTo74 = r2;
                                for (i = 1; i <= loopTo74; i++)
                                {
                                    var loopTo75 = c2;
                                    for (j = 1; j <= loopTo75; j++)
                                    {
                                        global::System.Int32 x;
                                        global::System.Int32 y;
                                        x = ((iRow) + (j));
                                        y = 1;
                                        if (((x) < (BreakPoints[(i) - (1)])))
                                        {
                                            if (((this.CheckBox1.Checked) == (false)))
                                            {
                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Value = this.rng.Cells[(global::System.Object)x, (global::System.Object)y].Value;
                                            }

                                            else if (((this.CheckBox1.Checked) == (true)))
                                            {

                                                this.rng.Cells[(global::System.Object)x, (global::System.Object)y].Copy();
                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].PasteSpecial(global::Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues);
                                                this.rng2 = this.worksheet2.get_Range(rng2Address);
                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].PasteSpecial(global::Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats);
                                                this.rng2 = this.worksheet2.get_Range(rng2Address);

                                            }
                                        }
                                    }
                                    iRow = BreakPoints[(i) - (1)];
                                }
                                excelApp.CutCopyMode = global::Microsoft.Office.Interop.Excel.XlCutCopyMode.xlCopy;
                            }
                            else if (X6)
                            {
                                global::System.Int32 iRow;
                                iRow = 0;
                                var loopTo76 = c2;
                                for (j = 1; j <= loopTo76; j++)
                                {
                                    var loopTo77 = r2;
                                    for (i = 1; i <= loopTo77; i++)
                                    {
                                        global::System.Int32 x;
                                        global::System.Int32 y;
                                        x = ((iRow) + (i));
                                        y = 1;

                                        if (((x) < (BreakPoints[(j) - (1)])))
                                        {
                                            if (((this.CheckBox1.Checked) == (false)))
                                            {
                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Value = this.rng.Cells[(global::System.Object)x, (global::System.Object)y].Value;
                                            }

                                            else if (((this.CheckBox1.Checked) == (true)))
                                            {

                                                this.rng.Cells[(global::System.Object)x, (global::System.Object)y].Copy();
                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].PasteSpecial(global::Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues);
                                                this.rng2 = this.worksheet2.get_Range(rng2Address);
                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].PasteSpecial(global::Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats);
                                                this.rng2 = this.worksheet2.get_Range(rng2Address);

                                            }
                                        }
                                    }
                                    iRow = BreakPoints[(j) - (1)];
                                }
                                excelApp.CutCopyMode = global::Microsoft.Office.Interop.Excel.XlCutCopyMode.xlCopy;
                            }

                            if (((this.CheckBox1.Checked) == (true)))
                            {
                                global::Microsoft.Office.Interop.Excel.Range TopCell = (global::Microsoft.Office.Interop.Excel.Range)this.rng.Cells[(global::System.Object)1, (global::System.Object)1];

                                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)8].LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                                {
                                    var loopTo78 = this.rng2.Columns.Count;
                                    for (j = 1; j <= loopTo78; j++)
                                    {
                                        this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)8).LineStyle = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)8].LineStyle;
                                        this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)8).Color = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)8].Color;
                                        this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)8).Weight = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)8].Weight;
                                    }
                                }
                                else
                                {
                                    var loopTo79 = this.rng2.Columns.Count;
                                    for (j = 1; j <= loopTo79; j++)
                                        this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)8).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                }

                                if (((this.rng2.Rows.Count) > (1)))
                                {
                                    if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                                    {
                                        var loopTo80 = this.rng2.Columns.Count;
                                        for (j = 1; j <= loopTo80; j++)
                                        {
                                            this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)9).LineStyle = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].LineStyle;
                                            this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)9).Color = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].Color;
                                            this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)9).Weight = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].Weight;
                                        }
                                    }
                                    else
                                    {
                                        var loopTo81 = this.rng2.Columns.Count;
                                        for (j = 1; j <= loopTo81; j++)
                                            this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)9).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                    }
                                }

                                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)7].LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                                {
                                    var loopTo82 = this.rng2.Rows.Count;
                                    for (i = 1; i <= loopTo82; i++)
                                    {
                                        this.rng2.Cells[(global::System.Object)i, (global::System.Object)1].Borders((global::System.Object)7).LineStyle = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)7].LineStyle;
                                        this.rng2.Cells[(global::System.Object)i, (global::System.Object)1].Borders((global::System.Object)7).Color = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)7].Color;
                                        this.rng2.Cells[(global::System.Object)i, (global::System.Object)1].Borders((global::System.Object)7).Weight = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)7].Weight;
                                    }
                                }
                                else
                                {
                                    var loopTo83 = this.rng2.Rows.Count;
                                    for (i = 1; i <= loopTo83; i++)
                                        this.rng2.Cells[(global::System.Object)i, (global::System.Object)1].Borders((global::System.Object)7).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                }

                                global::Microsoft.Office.Interop.Excel.Range BottomCell = (global::Microsoft.Office.Interop.Excel.Range)this.rng.Cells[(global::System.Object)r, (global::System.Object)c];

                                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(BottomCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                                {
                                    var loopTo84 = this.rng2.Columns.Count;
                                    for (j = 1; j <= loopTo84; j++)
                                    {
                                        this.rng2.Cells[(global::System.Object)this.rng2.Rows.Count, (global::System.Object)j].Borders((global::System.Object)9).LineStyle = BottomCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].LineStyle;
                                        this.rng2.Cells[(global::System.Object)this.rng2.Rows.Count, (global::System.Object)j].Borders((global::System.Object)9).Color = BottomCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].Color;
                                        this.rng2.Cells[(global::System.Object)this.rng2.Rows.Count, (global::System.Object)j].Borders((global::System.Object)9).Weight = BottomCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].Weight;
                                    }
                                }
                                else
                                {
                                    var loopTo85 = this.rng2.Columns.Count;
                                    for (j = 1; j <= loopTo85; j++)
                                        this.rng2.Cells[(global::System.Object)this.rng2.Rows.Count, (global::System.Object)this.rng2.Columns.Count].Borders((global::System.Object)9).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                }

                                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(BottomCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)10].LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                                {
                                    var loopTo86 = this.rng2.Rows.Count;
                                    for (i = 1; i <= loopTo86; i++)
                                    {
                                        this.rng2.Cells[(global::System.Object)i, (global::System.Object)this.rng2.Columns.Count].Borders((global::System.Object)10).LineStyle = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)10].LineStyle;
                                        this.rng2.Cells[(global::System.Object)i, (global::System.Object)this.rng2.Columns.Count].Borders((global::System.Object)10).Color = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)10].Color;
                                        this.rng2.Cells[(global::System.Object)i, (global::System.Object)this.rng2.Columns.Count].Borders((global::System.Object)10).Weight = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)10].Weight;
                                    }
                                }
                                else
                                {
                                    var loopTo87 = this.rng2.Rows.Count;
                                    for (i = 1; i <= loopTo87; i++)
                                        this.rng2.Cells[(global::System.Object)i, (global::System.Object)this.rng2.Columns.Count].Borders((global::System.Object)10).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                }

                                if (((this.rng.Rows.Count) > (1)))
                                {
                                    global::Microsoft.Office.Interop.Excel.Range MiddleCell = (global::Microsoft.Office.Interop.Excel.Range)this.rng.Cells[(global::System.Object)2, (global::System.Object)1];
                                    if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(MiddleCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                                    {
                                        var loopTo88 = (this.rng2.Rows.Count) - (1);
                                        for (i = 2; i <= loopTo88; i++)
                                        {
                                            var loopTo89 = this.rng2.Columns.Count;
                                            for (j = 1; j <= loopTo89; j++)
                                            {
                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)9).LineStyle = MiddleCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].LineStyle;
                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)9).Color = MiddleCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].Color;
                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)9).Weight = MiddleCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].Weight;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        var loopTo90 = (this.rng2.Rows.Count) - (1);
                                        for (i = 2; i <= loopTo90; i++)
                                        {
                                            var loopTo91 = this.rng2.Columns.Count;
                                            for (j = 1; j <= loopTo91; j++)
                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)9).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                        }
                                    }
                                }

                                if (((this.rng.Columns.Count) > (1)))
                                {
                                    global::Microsoft.Office.Interop.Excel.Range MiddleCell = (global::Microsoft.Office.Interop.Excel.Range)this.rng.Cells[(global::System.Object)1, (global::System.Object)2];
                                    if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(MiddleCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)10].LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                                    {
                                        var loopTo92 = (this.rng2.Columns.Count) - (1);
                                        for (j = 1; j <= loopTo92; j++)
                                        {
                                            var loopTo93 = this.rng2.Rows.Count;
                                            for (i = 1; i <= loopTo93; i++)
                                            {
                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)10).LineStyle = MiddleCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)10].LineStyle;
                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)10).Color = MiddleCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)10].Color;
                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)10).Weight = MiddleCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)10].Weight;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        var loopTo94 = (this.rng2.Columns.Count) - (1);
                                        for (j = 1; j <= loopTo94; j++)
                                        {
                                            var loopTo95 = this.rng2.Rows.Count;
                                            for (i = 1; i <= loopTo95; i++)
                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)10).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                        }
                                    }
                                }
                            }
                        }

                        else
                        {

                            var Arr = new global::System.Object[(r), (c)];
                            var Bolds = new global::System.Boolean[(r), (c)];
                            var Italics = new global::System.Boolean[(r), (c)];
                            var fontNames = new global::System.String[(r), (c)];
                            var fontSizes = new global::System.Single[(r), (c)];
                            var reds1 = new global::System.Int32[(r), (c)];
                            var reds2 = new global::System.Int32[(r), (c)];
                            var greens1 = new global::System.Int32[(r), (c)];
                            var greens2 = new global::System.Int32[(r), (c)];
                            var blues1 = new global::System.Int32[(r), (c)];
                            var blues2 = new global::System.Int32[(r), (c)];

                            global::System.Boolean TopBorder7;
                            global::System.Object TopBorder7L;
                            global::System.Object TopBorder7C;
                            global::System.Object TopBorder7W;

                            global::System.Boolean TopBorder8;
                            global::System.Object TopBorder8L;
                            global::System.Object TopBorder8C;
                            global::System.Object TopBorder8W;

                            global::System.Boolean TopBorder9;
                            global::System.Object TopBorder9L;
                            global::System.Object TopBorder9C;
                            global::System.Object TopBorder9W;

                            global::System.Boolean BottomBorder9;
                            global::System.Object BottomBorder9L;
                            global::System.Object BottomBorder9C;
                            global::System.Object BottomBorder9W;

                            global::System.Boolean BottomBorder10;
                            global::System.Object BottomBorder10L;
                            global::System.Object BottomBorder10C;
                            global::System.Object BottomBorder10W;

                            global::System.Boolean MiddleBorder9;
                            global::System.Object MiddleBorder9L;
                            global::System.Object MiddleBorder9C;
                            global::System.Object MiddleBorder9W;

                            global::System.Boolean MiddleBorder10;
                            global::System.Object MiddleBorder10L;
                            global::System.Object MiddleBorder10C;
                            global::System.Object MiddleBorder10W;

                            if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)7).LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                            {
                                TopBorder7 = true;
                                TopBorder7L = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)7).LineStyle;
                                TopBorder7C = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)7).Color;
                                TopBorder7W = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)7).Weight;
                            }
                            else
                            {
                                TopBorder7 = false;
                                TopBorder7L = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)7).LineStyle;
                                TopBorder7C = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)7).Color;
                                TopBorder7W = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)7).Weight;
                            }

                            if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)8).LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                            {
                                TopBorder8 = true;
                                TopBorder8L = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)8).LineStyle;
                                TopBorder8C = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)8).Color;
                                TopBorder8W = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)8).Weight;
                            }
                            else
                            {
                                TopBorder8 = false;
                                TopBorder8L = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)8).LineStyle;
                                TopBorder8C = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)8).Color;
                                TopBorder8W = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)8).Weight;
                            }

                            if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)9).LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                            {
                                TopBorder9 = true;
                                TopBorder9L = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)9).LineStyle;
                                TopBorder9C = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)9).Color;
                                TopBorder9W = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)9).Weight;
                            }
                            else
                            {
                                TopBorder9 = false;
                                TopBorder9L = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)9).LineStyle;
                                TopBorder9C = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)9).Color;
                                TopBorder9W = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)9).Weight;
                            }

                            if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)9).LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                            {
                                BottomBorder9 = true;
                                BottomBorder9L = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)9).LineStyle;
                                BottomBorder9C = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)9).Color;
                                BottomBorder9W = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)9).Weight;
                            }
                            else
                            {
                                BottomBorder9 = false;
                                BottomBorder9L = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)9).LineStyle;
                                BottomBorder9C = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)9).Color;
                                BottomBorder9W = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)9).Weight;
                            }

                            if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)10).LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                            {
                                BottomBorder10 = true;
                                BottomBorder10L = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)10).LineStyle;
                                BottomBorder10C = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)10).Color;
                                BottomBorder10W = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)10).Weight;
                            }
                            else
                            {
                                BottomBorder10 = false;
                                BottomBorder10L = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)10).LineStyle;
                                BottomBorder10C = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)10).Color;
                                BottomBorder10W = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)10).Weight;
                            }

                            if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(this.rng.Cells[(global::System.Object)2, (global::System.Object)1].Borders((global::System.Object)9).LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                            {
                                MiddleBorder9 = true;
                                MiddleBorder9L = this.rng.Cells[(global::System.Object)2, (global::System.Object)1].Borders((global::System.Object)9).LineStyle;
                                MiddleBorder9C = this.rng.Cells[(global::System.Object)2, (global::System.Object)1].Borders((global::System.Object)9).Color;
                                MiddleBorder9W = this.rng.Cells[(global::System.Object)2, (global::System.Object)1].Borders((global::System.Object)9).Weight;
                            }
                            else
                            {
                                MiddleBorder9 = false;
                                MiddleBorder9L = this.rng.Cells[(global::System.Object)2, (global::System.Object)1].Borders((global::System.Object)9).LineStyle;
                                MiddleBorder9C = this.rng.Cells[(global::System.Object)2, (global::System.Object)1].Borders((global::System.Object)9).Color;
                                MiddleBorder9W = this.rng.Cells[(global::System.Object)2, (global::System.Object)1].Borders((global::System.Object)9).Weight;
                            }

                            if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(this.rng.Cells[(global::System.Object)1, (global::System.Object)2].Borders((global::System.Object)10).LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                            {
                                MiddleBorder10 = true;
                                MiddleBorder10L = this.rng.Cells[(global::System.Object)1, (global::System.Object)2].Borders((global::System.Object)10).LineStyle;
                                MiddleBorder10C = this.rng.Cells[(global::System.Object)1, (global::System.Object)2].Borders((global::System.Object)10).Color;
                                MiddleBorder10W = this.rng.Cells[(global::System.Object)1, (global::System.Object)2].Borders((global::System.Object)10).Weight;
                            }
                            else
                            {
                                MiddleBorder10 = false;
                                MiddleBorder10L = this.rng.Cells[(global::System.Object)1, (global::System.Object)2].Borders((global::System.Object)10).LineStyle;
                                MiddleBorder10C = this.rng.Cells[(global::System.Object)1, (global::System.Object)2].Borders((global::System.Object)10).Color;
                                MiddleBorder10W = this.rng.Cells[(global::System.Object)1, (global::System.Object)2].Borders((global::System.Object)10).Weight;
                            }

                            var loopTo96 = r;
                            for (i = 1; i <= loopTo96; i++)
                            {
                                var loopTo97 = c;
                                for (j = 1; j <= loopTo97; j++)
                                {
                                    Arr[(i) - (1), (j) - (1)] = this.rng.Cells[(global::System.Object)i, (global::System.Object)j].Value;

                                    if (((this.CheckBox1.Checked) == (true)))
                                    {

                                        global::Microsoft.Office.Interop.Excel.Range cell = (global::Microsoft.Office.Interop.Excel.Range)this.rng.Cells[(global::System.Object)i, (global::System.Object)j];
                                        var font = cell.Font;

                                        Bolds[(i) - (1), (j) - (1)] = Conversions.ToBoolean(cell.Font.Bold);
                                        Italics[(i) - (1), (j) - (1)] = Conversions.ToBoolean(cell.Font.Italic);


                                        if ((((font.Name is System.DBNull)) == (false)))
                                        {
                                            fontNames[(i) - (1), (j) - (1)] = Conversions.ToString(font.Name);
                                        }
                                        else
                                        {
                                            fontNames[(i) - (1), (j) - (1)] = "Calibri";
                                        }

                                        if ((((font.Size is System.DBNull)) == (false)))
                                        {
                                            global::System.Single fontSize = global::System.Convert.ToSingle(font.Size);
                                            fontSizes[(i) - (1), (j) - (1)] = fontSize;
                                        }
                                        else
                                        {
                                            fontSizes[(i) - (1), (j) - (1)] = 11f;
                                        }

                                        if ((cell.Interior.Color is System.DBNull))
                                        {
                                            reds1[(i) - (1), (j) - (1)] = 255;
                                            greens1[(i) - (1), (j) - (1)] = 255;
                                            blues1[(i) - (1), (j) - (1)] = 255;
                                        }
                                        else
                                        {
                                            global::System.Int64 colorValue1 = Conversions.ToLong(cell.Interior.Color);
                                            global::System.Int32 red1 = (global::System.Int32)((colorValue1) % (256L));
                                            global::System.Int32 green1 = (global::System.Int32)(((((colorValue1) / (256L)))) % (256L));
                                            global::System.Int32 blue1 = (global::System.Int32)((((((colorValue1) / (256L)) / (256L)))) % (256L));
                                            reds1[(i) - (1), (j) - (1)] = red1;
                                            greens1[(i) - (1), (j) - (1)] = green1;
                                            blues1[(i) - (1), (j) - (1)] = blue1;
                                        }

                                        if ((cell.Font.Color is System.DBNull))
                                        {
                                            reds2[(i) - (1), (j) - (1)] = 0;
                                            greens2[(i) - (1), (j) - (1)] = 0;
                                            blues2[(i) - (1), (j) - (1)] = 0;
                                        }
                                        else
                                        {
                                            global::System.Int64 colorValue2 = Conversions.ToLong(cell.Font.Color);
                                            global::System.Int32 red2 = (global::System.Int32)((colorValue2) % (256L));
                                            global::System.Int32 green2 = (global::System.Int32)(((((colorValue2) / (256L)))) % (256L));
                                            global::System.Int32 blue2 = (global::System.Int32)((((((colorValue2) / (256L)) / (256L)))) % (256L));
                                            reds2[(i) - (1), (j) - (1)] = red2;
                                            greens2[(i) - (1), (j) - (1)] = green2;
                                            blues2[(i) - (1), (j) - (1)] = blue2;
                                        }
                                    }

                                }
                            }

                            this.rng.ClearContents();
                            this.rng.ClearFormats();

                            this.rng2.ClearFormats();

                            if (X5)
                            {
                                global::System.Int32 iRow;
                                iRow = 0;
                                var loopTo98 = r2;
                                for (i = 1; i <= loopTo98; i++)
                                {
                                    var loopTo99 = c2;
                                    for (j = 1; j <= loopTo99; j++)
                                    {
                                        global::System.Int32 x;
                                        global::System.Int32 y;
                                        x = ((iRow) + (j));
                                        y = 1;
                                        if ((((x) < (BreakPoints[(i) - (1)])) & ((x) <= (r))))
                                        {
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Value = Arr[(x) - (1), (y) - (1)];

                                            if (((this.CheckBox1.Checked) == (true)))
                                            {

                                                global::Microsoft.Office.Interop.Excel.Range cell2 = (global::Microsoft.Office.Interop.Excel.Range)this.rng2.Cells[(global::System.Object)i, (global::System.Object)j];
                                                var font2 = cell2.Font;

                                                global::System.Single fontSize = fontSizes[(x) - (1), (y) - (1)];

                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Font.Name = fontNames[(x) - (1), (y) - (1)];
                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Font.Size = (global::System.Object)fontSizes[(x) - (1), (y) - (1)];

                                                if (Bolds[(x) - (1), (y) - (1)])
                                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Font.Bold = (global::System.Object)true;
                                                if (Italics[(x) - (1), (y) - (1)])
                                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Font.Italic = (global::System.Object)true;

                                                global::System.Int32 red1 = reds1[(x) - (1), (y) - (1)];
                                                global::System.Int32 green1 = greens1[(x) - (1), (y) - (1)];
                                                global::System.Int32 blue1 = blues1[(x) - (1), (y) - (1)];
                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Interior.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(red1, green1, blue1);

                                                global::System.Int32 red2 = reds2[(x) - (1), (y) - (1)];
                                                global::System.Int32 green2 = greens2[(x) - (1), (y) - (1)];
                                                global::System.Int32 blue2 = blues2[(x) - (1), (y) - (1)];
                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Font.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(red2, green2, blue2);
                                            }
                                        }
                                    }
                                    iRow = BreakPoints[(i) - (1)];
                                }
                            }

                            else if (X6)
                            {
                                global::System.Int32 iRow;
                                iRow = 0;
                                var loopTo100 = c2;
                                for (j = 1; j <= loopTo100; j++)
                                {
                                    var loopTo101 = r2;
                                    for (i = 1; i <= loopTo101; i++)
                                    {
                                        global::System.Int32 x;
                                        global::System.Int32 y;
                                        x = ((iRow) + (i));
                                        y = 1;
                                        if (((x) < (BreakPoints[(j) - (1)])))
                                        {
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Value = Arr[(x) - (1), (y) - (1)];

                                            if (((this.CheckBox1.Checked) == (true)))
                                            {

                                                global::Microsoft.Office.Interop.Excel.Range cell2 = (global::Microsoft.Office.Interop.Excel.Range)this.rng2.Cells[(global::System.Object)i, (global::System.Object)j];
                                                var font2 = cell2.Font;

                                                global::System.Single fontSize = fontSizes[(x) - (1), (y) - (1)];

                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Font.Name = fontNames[(x) - (1), (y) - (1)];
                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Font.Size = (global::System.Object)fontSizes[(x) - (1), (y) - (1)];

                                                if (Bolds[(x) - (1), (y) - (1)])
                                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Font.Bold = (global::System.Object)true;
                                                if (Italics[(x) - (1), (y) - (1)])
                                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Font.Italic = (global::System.Object)true;

                                                global::System.Int32 red1 = reds1[(x) - (1), (y) - (1)];
                                                global::System.Int32 green1 = greens1[(x) - (1), (y) - (1)];
                                                global::System.Int32 blue1 = blues1[(x) - (1), (y) - (1)];
                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Interior.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(red1, green1, blue1);

                                                global::System.Int32 red2 = reds2[(x) - (1), (y) - (1)];
                                                global::System.Int32 green2 = greens2[(x) - (1), (y) - (1)];
                                                global::System.Int32 blue2 = blues2[(x) - (1), (y) - (1)];
                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Font.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(red2, green2, blue2);
                                            }
                                        }
                                    }
                                    iRow = BreakPoints[(j) - (1)];
                                }
                            }

                            if (((this.CheckBox1.Checked) == (true)))
                            {

                                if (((TopBorder8) == (true)))
                                {
                                    var loopTo102 = this.rng2.Columns.Count;
                                    for (j = 1; j <= loopTo102; j++)
                                    {
                                        this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)8).LineStyle = TopBorder8L;
                                        this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)8).Color = TopBorder8C;
                                        this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)8).Weight = TopBorder8W;
                                    }
                                }
                                else
                                {
                                    var loopTo103 = this.rng2.Columns.Count;
                                    for (j = 1; j <= loopTo103; j++)
                                        this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)8).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                }

                                if (((this.rng2.Rows.Count) > (1)))
                                {
                                    if (((TopBorder9) == (true)))
                                    {
                                        var loopTo104 = this.rng2.Columns.Count;
                                        for (j = 1; j <= loopTo104; j++)
                                        {
                                            this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)9).LineStyle = TopBorder9L;
                                            this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)9).Color = TopBorder9C;
                                            this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)9).Weight = TopBorder9W;
                                        }
                                    }
                                    else
                                    {
                                        var loopTo105 = this.rng2.Columns.Count;
                                        for (j = 1; j <= loopTo105; j++)
                                            this.rng2.Cells[(global::System.Object)this.rng2.Rows.Count, (global::System.Object)j].Borders((global::System.Object)9).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                    }
                                }

                                if (((TopBorder7) == (true)))
                                {
                                    var loopTo106 = this.rng2.Rows.Count;
                                    for (i = 1; i <= loopTo106; i++)
                                    {
                                        this.rng2.Cells[(global::System.Object)i, (global::System.Object)1].Borders((global::System.Object)7).LineStyle = TopBorder7L;
                                        this.rng2.Cells[(global::System.Object)i, (global::System.Object)1].Borders((global::System.Object)7).Color = TopBorder7C;
                                        this.rng2.Cells[(global::System.Object)i, (global::System.Object)1].Borders((global::System.Object)7).Weight = TopBorder7W;
                                    }
                                }
                                else
                                {
                                    var loopTo107 = this.rng2.Rows.Count;
                                    for (i = 1; i <= loopTo107; i++)
                                        this.rng2.Cells[(global::System.Object)i, (global::System.Object)1].Borders((global::System.Object)7).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                }

                                if (((BottomBorder9) == (true)))
                                {
                                    var loopTo108 = this.rng2.Columns.Count;
                                    for (j = 1; j <= loopTo108; j++)
                                    {
                                        this.rng2.Cells[(global::System.Object)this.rng2.Rows.Count, (global::System.Object)j].Borders((global::System.Object)9).LineStyle = BottomBorder9L;
                                        this.rng2.Cells[(global::System.Object)this.rng2.Rows.Count, (global::System.Object)j].Borders((global::System.Object)9).Color = BottomBorder9C;
                                        this.rng2.Cells[(global::System.Object)this.rng2.Rows.Count, (global::System.Object)j].Borders((global::System.Object)9).Weight = BottomBorder9W;
                                    }
                                }
                                else
                                {
                                    var loopTo109 = this.rng2.Columns.Count;
                                    for (j = 1; j <= loopTo109; j++)
                                        this.rng2.Cells[(global::System.Object)this.rng2.Rows.Count, (global::System.Object)this.rng2.Columns.Count].Borders((global::System.Object)9).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                }

                                if (((BottomBorder10) == (true)))
                                {
                                    var loopTo110 = this.rng2.Rows.Count;
                                    for (i = 1; i <= loopTo110; i++)
                                    {
                                        this.rng2.Cells[(global::System.Object)i, (global::System.Object)this.rng2.Columns.Count].Borders((global::System.Object)10).LineStyle = BottomBorder10L;
                                        this.rng2.Cells[(global::System.Object)i, (global::System.Object)this.rng2.Columns.Count].Borders((global::System.Object)10).Color = BottomBorder10C;
                                        this.rng2.Cells[(global::System.Object)i, (global::System.Object)this.rng2.Columns.Count].Borders((global::System.Object)10).Weight = BottomBorder10W;
                                    }
                                }
                                else
                                {
                                    var loopTo111 = this.rng2.Rows.Count;
                                    for (i = 1; i <= loopTo111; i++)
                                        this.rng2.Cells[(global::System.Object)i, (global::System.Object)this.rng2.Columns.Count].Borders((global::System.Object)10).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                }

                                if (((this.rng.Rows.Count) > (1)))
                                {

                                    if (((MiddleBorder9) == (true)))
                                    {
                                        var loopTo112 = (this.rng2.Rows.Count) - (1);
                                        for (i = 2; i <= loopTo112; i++)
                                        {
                                            var loopTo113 = this.rng2.Columns.Count;
                                            for (j = 1; j <= loopTo113; j++)
                                            {
                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)9).LineStyle = MiddleBorder9L;
                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)9).Color = MiddleBorder9C;
                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)9).Weight = MiddleBorder9W;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        var loopTo114 = (this.rng2.Rows.Count) - (1);
                                        for (i = 2; i <= loopTo114; i++)
                                        {
                                            var loopTo115 = this.rng2.Columns.Count;
                                            for (j = 1; j <= loopTo115; j++)
                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)9).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                        }
                                    }

                                }

                                if (((this.rng.Columns.Count) > (1)))
                                {

                                    if (((MiddleBorder10) == (true)))
                                    {
                                        var loopTo116 = (this.rng2.Columns.Count) - (1);
                                        for (j = 1; j <= loopTo116; j++)
                                        {
                                            var loopTo117 = this.rng2.Rows.Count;
                                            for (i = 1; i <= loopTo117; i++)
                                            {
                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)10).LineStyle = MiddleBorder10L;
                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)10).Color = MiddleBorder10C;
                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)10).Weight = MiddleBorder10W;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        var loopTo118 = (this.rng2.Columns.Count) - (1);
                                        for (j = 1; j <= loopTo118; j++)
                                        {
                                            var loopTo119 = this.rng2.Rows.Count;
                                            for (i = 1; i <= loopTo119; i++)
                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)10).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                        }
                                    }

                                }

                            }

                        }
                    }

                    else if (((((((X8) & !string.IsNullOrEmpty(this.TextBox2.Text)) & ((this.CanConvertToInt(this.TextBox2.Text)) == (true))))) & ((((X5) | (X6))))))
                    {

                        if (X5)
                        {

                            global::System.Int32 r2;
                            global::System.Int32 c2;

                            if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(Operators.ModObject(r, global::Microsoft.VisualBasic.Conversion.Int(this.TextBox2.Text)), 0, false)))
                            {
                                r2 = Conversions.ToInteger(global::Microsoft.VisualBasic.Conversion.Int(Operators.DivideObject(r, global::Microsoft.VisualBasic.Conversion.Int(this.TextBox2.Text))));
                            }
                            else
                            {
                                r2 = Conversions.ToInteger(Operators.AddObject(global::Microsoft.VisualBasic.Conversion.Int(Operators.DivideObject(r, global::Microsoft.VisualBasic.Conversion.Int(this.TextBox2.Text))), 1));
                            }
                            c2 = Conversions.ToInteger(global::Microsoft.VisualBasic.Conversion.Int(this.TextBox2.Text));

                            this.rng2 = this.worksheet2.get_Range(this.rng2.Cells[(global::System.Object)1, (global::System.Object)1], this.rng2.Cells[(global::System.Object)r2, (global::System.Object)c2]);
                            global::System.String rng2Address = this.rng2.get_Address();
                            this.worksheet2.Activate();
                            this.rng2.Select();

                            if (((this.Overlap(this.excelApp, this.worksheet, this.worksheet2, this.rng, this.rng2)) == (false)))
                            {

                                this.rng2.ClearFormats();

                                var loopTo120 = r2;
                                for (i = 1; i <= loopTo120; i++)
                                {
                                    var loopTo121 = c2;
                                    for (j = 1; j <= loopTo121; j++)
                                    {
                                        global::System.Int32 x;
                                        global::System.Int32 y;
                                        x = (((((c2) * ((((i) - (1))))))) + (j));
                                        y = 1;
                                        if (((x) <= (this.rng.Rows.Count)))
                                        {
                                            if (((this.CheckBox1.Checked) == (false)))
                                            {
                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Value = this.rng.Cells[(global::System.Object)x, (global::System.Object)y].Value;
                                            }

                                            else if (((this.CheckBox1.Checked) == (true)))
                                            {

                                                this.rng.Cells[(global::System.Object)x, (global::System.Object)y].Copy();
                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].PasteSpecial(global::Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues);
                                                this.rng2 = this.worksheet2.get_Range(rng2Address);
                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].PasteSpecial(global::Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats);
                                                this.rng2 = this.worksheet2.get_Range(rng2Address);

                                            }
                                        }
                                    }
                                }
                                excelApp.CutCopyMode = global::Microsoft.Office.Interop.Excel.XlCutCopyMode.xlCopy;

                                if (((this.CheckBox1.Checked) == (true)))
                                {
                                    global::Microsoft.Office.Interop.Excel.Range TopCell = (global::Microsoft.Office.Interop.Excel.Range)this.rng.Cells[(global::System.Object)1, (global::System.Object)1];

                                    if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)8].LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                                    {
                                        var loopTo122 = this.rng2.Columns.Count;
                                        for (j = 1; j <= loopTo122; j++)
                                        {
                                            this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)8).LineStyle = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)8].LineStyle;
                                            this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)8).Color = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)8].Color;
                                            this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)8).Weight = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)8].Weight;
                                        }
                                    }
                                    else
                                    {
                                        var loopTo123 = this.rng2.Columns.Count;
                                        for (j = 1; j <= loopTo123; j++)
                                            this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)8).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                    }

                                    if (((this.rng2.Rows.Count) > (1)))
                                    {
                                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                                        {
                                            var loopTo124 = this.rng2.Columns.Count;
                                            for (j = 1; j <= loopTo124; j++)
                                            {
                                                this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)9).LineStyle = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].LineStyle;
                                                this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)9).Color = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].Color;
                                                this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)9).Weight = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].Weight;
                                            }
                                        }
                                        else
                                        {
                                            var loopTo125 = this.rng2.Columns.Count;
                                            for (j = 1; j <= loopTo125; j++)
                                                this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)9).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                        }
                                    }

                                    if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)7].LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                                    {
                                        var loopTo126 = this.rng2.Rows.Count;
                                        for (i = 1; i <= loopTo126; i++)
                                        {
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)1].Borders((global::System.Object)7).LineStyle = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)7].LineStyle;
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)1].Borders((global::System.Object)7).Color = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)7].Color;
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)1].Borders((global::System.Object)7).Weight = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)7].Weight;
                                        }
                                    }
                                    else
                                    {
                                        var loopTo127 = this.rng2.Rows.Count;
                                        for (i = 1; i <= loopTo127; i++)
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)1].Borders((global::System.Object)7).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                    }

                                    global::Microsoft.Office.Interop.Excel.Range BottomCell = (global::Microsoft.Office.Interop.Excel.Range)this.rng.Cells[(global::System.Object)r, (global::System.Object)c];

                                    if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(BottomCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                                    {
                                        var loopTo128 = this.rng2.Columns.Count;
                                        for (j = 1; j <= loopTo128; j++)
                                        {
                                            this.rng2.Cells[(global::System.Object)this.rng2.Rows.Count, (global::System.Object)j].Borders((global::System.Object)9).LineStyle = BottomCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].LineStyle;
                                            this.rng2.Cells[(global::System.Object)this.rng2.Rows.Count, (global::System.Object)j].Borders((global::System.Object)9).Color = BottomCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].Color;
                                            this.rng2.Cells[(global::System.Object)this.rng2.Rows.Count, (global::System.Object)j].Borders((global::System.Object)9).Weight = BottomCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].Weight;
                                        }
                                    }
                                    else
                                    {
                                        var loopTo129 = this.rng2.Columns.Count;
                                        for (j = 1; j <= loopTo129; j++)
                                            this.rng2.Cells[(global::System.Object)this.rng2.Rows.Count, (global::System.Object)this.rng2.Columns.Count].Borders((global::System.Object)9).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                    }

                                    if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(BottomCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)10].LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                                    {
                                        var loopTo130 = this.rng2.Rows.Count;
                                        for (i = 1; i <= loopTo130; i++)
                                        {
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)this.rng2.Columns.Count].Borders((global::System.Object)10).LineStyle = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)10].LineStyle;
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)this.rng2.Columns.Count].Borders((global::System.Object)10).Color = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)10].Color;
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)this.rng2.Columns.Count].Borders((global::System.Object)10).Weight = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)10].Weight;
                                        }
                                    }
                                    else
                                    {
                                        var loopTo131 = this.rng2.Rows.Count;
                                        for (i = 1; i <= loopTo131; i++)
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)this.rng2.Columns.Count].Borders((global::System.Object)10).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                    }

                                    if (((this.rng.Rows.Count) > (1)))
                                    {
                                        global::Microsoft.Office.Interop.Excel.Range MiddleCell = (global::Microsoft.Office.Interop.Excel.Range)this.rng.Cells[(global::System.Object)2, (global::System.Object)1];
                                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(MiddleCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                                        {
                                            var loopTo132 = (this.rng2.Rows.Count) - (1);
                                            for (i = 2; i <= loopTo132; i++)
                                            {
                                                var loopTo133 = this.rng2.Columns.Count;
                                                for (j = 1; j <= loopTo133; j++)
                                                {
                                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)9).LineStyle = MiddleCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].LineStyle;
                                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)9).Color = MiddleCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].Color;
                                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)9).Weight = MiddleCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].Weight;
                                                }
                                            }
                                        }
                                        else
                                        {
                                            var loopTo134 = (this.rng2.Rows.Count) - (1);
                                            for (i = 2; i <= loopTo134; i++)
                                            {
                                                var loopTo135 = this.rng2.Columns.Count;
                                                for (j = 1; j <= loopTo135; j++)
                                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)9).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                            }
                                        }
                                    }

                                    if (((this.rng.Columns.Count) > (1)))
                                    {
                                        global::Microsoft.Office.Interop.Excel.Range MiddleCell = (global::Microsoft.Office.Interop.Excel.Range)this.rng.Cells[(global::System.Object)1, (global::System.Object)2];
                                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(MiddleCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)10].LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                                        {
                                            var loopTo136 = (this.rng2.Columns.Count) - (1);
                                            for (j = 1; j <= loopTo136; j++)
                                            {
                                                var loopTo137 = this.rng2.Rows.Count;
                                                for (i = 1; i <= loopTo137; i++)
                                                {
                                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)10).LineStyle = MiddleCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)10].LineStyle;
                                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)10).Color = MiddleCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)10].Color;
                                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)10).Weight = MiddleCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)10].Weight;
                                                }
                                            }
                                        }
                                        else
                                        {
                                            var loopTo138 = (this.rng2.Columns.Count) - (1);
                                            for (j = 1; j <= loopTo138; j++)
                                            {
                                                var loopTo139 = this.rng2.Rows.Count;
                                                for (i = 1; i <= loopTo139; i++)
                                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)10).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                            }
                                        }
                                    }
                                }
                            }

                            else
                            {

                                var Arr = new global::System.Object[(r), (c)];
                                var Bolds = new global::System.Boolean[(r), (c)];
                                var Italics = new global::System.Boolean[(r), (c)];
                                var fontNames = new global::System.String[(r), (c)];
                                var fontSizes = new global::System.Single[(r), (c)];
                                var reds1 = new global::System.Int32[(r), (c)];
                                var reds2 = new global::System.Int32[(r), (c)];
                                var greens1 = new global::System.Int32[(r), (c)];
                                var greens2 = new global::System.Int32[(r), (c)];
                                var blues1 = new global::System.Int32[(r), (c)];
                                var blues2 = new global::System.Int32[(r), (c)];

                                global::System.Boolean TopBorder7;
                                global::System.Object TopBorder7L;
                                global::System.Object TopBorder7C;
                                global::System.Object TopBorder7W;

                                global::System.Boolean TopBorder8;
                                global::System.Object TopBorder8L;
                                global::System.Object TopBorder8C;
                                global::System.Object TopBorder8W;

                                global::System.Boolean TopBorder9;
                                global::System.Object TopBorder9L;
                                global::System.Object TopBorder9C;
                                global::System.Object TopBorder9W;

                                global::System.Boolean BottomBorder9;
                                global::System.Object BottomBorder9L;
                                global::System.Object BottomBorder9C;
                                global::System.Object BottomBorder9W;

                                global::System.Boolean BottomBorder10;
                                global::System.Object BottomBorder10L;
                                global::System.Object BottomBorder10C;
                                global::System.Object BottomBorder10W;

                                global::System.Boolean MiddleBorder9;
                                global::System.Object MiddleBorder9L;
                                global::System.Object MiddleBorder9C;
                                global::System.Object MiddleBorder9W;

                                global::System.Boolean MiddleBorder10;
                                global::System.Object MiddleBorder10L;
                                global::System.Object MiddleBorder10C;
                                global::System.Object MiddleBorder10W;

                                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)7).LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                                {
                                    TopBorder7 = true;
                                    TopBorder7L = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)7).LineStyle;
                                    TopBorder7C = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)7).Color;
                                    TopBorder7W = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)7).Weight;
                                }
                                else
                                {
                                    TopBorder7 = false;
                                    TopBorder7L = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)7).LineStyle;
                                    TopBorder7C = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)7).Color;
                                    TopBorder7W = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)7).Weight;
                                }

                                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)8).LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                                {
                                    TopBorder8 = true;
                                    TopBorder8L = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)8).LineStyle;
                                    TopBorder8C = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)8).Color;
                                    TopBorder8W = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)8).Weight;
                                }
                                else
                                {
                                    TopBorder8 = false;
                                    TopBorder8L = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)8).LineStyle;
                                    TopBorder8C = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)8).Color;
                                    TopBorder8W = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)8).Weight;
                                }

                                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)9).LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                                {
                                    TopBorder9 = true;
                                    TopBorder9L = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)9).LineStyle;
                                    TopBorder9C = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)9).Color;
                                    TopBorder9W = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)9).Weight;
                                }
                                else
                                {
                                    TopBorder9 = false;
                                    TopBorder9L = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)9).LineStyle;
                                    TopBorder9C = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)9).Color;
                                    TopBorder9W = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)9).Weight;
                                }

                                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)9).LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                                {
                                    BottomBorder9 = true;
                                    BottomBorder9L = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)9).LineStyle;
                                    BottomBorder9C = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)9).Color;
                                    BottomBorder9W = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)9).Weight;
                                }
                                else
                                {
                                    BottomBorder9 = false;
                                    BottomBorder9L = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)9).LineStyle;
                                    BottomBorder9C = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)9).Color;
                                    BottomBorder9W = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)9).Weight;
                                }

                                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)10).LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                                {
                                    BottomBorder10 = true;
                                    BottomBorder10L = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)10).LineStyle;
                                    BottomBorder10C = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)10).Color;
                                    BottomBorder10W = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)10).Weight;
                                }
                                else
                                {
                                    BottomBorder10 = false;
                                    BottomBorder10L = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)10).LineStyle;
                                    BottomBorder10C = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)10).Color;
                                    BottomBorder10W = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)10).Weight;
                                }

                                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(this.rng.Cells[(global::System.Object)2, (global::System.Object)1].Borders((global::System.Object)9).LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                                {
                                    MiddleBorder9 = true;
                                    MiddleBorder9L = this.rng.Cells[(global::System.Object)2, (global::System.Object)1].Borders((global::System.Object)9).LineStyle;
                                    MiddleBorder9C = this.rng.Cells[(global::System.Object)2, (global::System.Object)1].Borders((global::System.Object)9).Color;
                                    MiddleBorder9W = this.rng.Cells[(global::System.Object)2, (global::System.Object)1].Borders((global::System.Object)9).Weight;
                                }
                                else
                                {
                                    MiddleBorder9 = false;
                                    MiddleBorder9L = this.rng.Cells[(global::System.Object)2, (global::System.Object)1].Borders((global::System.Object)9).LineStyle;
                                    MiddleBorder9C = this.rng.Cells[(global::System.Object)2, (global::System.Object)1].Borders((global::System.Object)9).Color;
                                    MiddleBorder9W = this.rng.Cells[(global::System.Object)2, (global::System.Object)1].Borders((global::System.Object)9).Weight;
                                }

                                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(this.rng.Cells[(global::System.Object)1, (global::System.Object)2].Borders((global::System.Object)10).LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                                {
                                    MiddleBorder10 = true;
                                    MiddleBorder10L = this.rng.Cells[(global::System.Object)1, (global::System.Object)2].Borders((global::System.Object)10).LineStyle;
                                    MiddleBorder10C = this.rng.Cells[(global::System.Object)1, (global::System.Object)2].Borders((global::System.Object)10).Color;
                                    MiddleBorder10W = this.rng.Cells[(global::System.Object)1, (global::System.Object)2].Borders((global::System.Object)10).Weight;
                                }
                                else
                                {
                                    MiddleBorder10 = false;
                                    MiddleBorder10L = this.rng.Cells[(global::System.Object)1, (global::System.Object)2].Borders((global::System.Object)10).LineStyle;
                                    MiddleBorder10C = this.rng.Cells[(global::System.Object)1, (global::System.Object)2].Borders((global::System.Object)10).Color;
                                    MiddleBorder10W = this.rng.Cells[(global::System.Object)1, (global::System.Object)2].Borders((global::System.Object)10).Weight;
                                }

                                var loopTo140 = r;
                                for (i = 1; i <= loopTo140; i++)
                                {
                                    var loopTo141 = c;
                                    for (j = 1; j <= loopTo141; j++)
                                    {
                                        Arr[(i) - (1), (j) - (1)] = this.rng.Cells[(global::System.Object)i, (global::System.Object)j].Value;

                                        if (((this.CheckBox1.Checked) == (true)))
                                        {

                                            global::Microsoft.Office.Interop.Excel.Range cell = (global::Microsoft.Office.Interop.Excel.Range)this.rng.Cells[(global::System.Object)i, (global::System.Object)j];
                                            var font = cell.Font;

                                            Bolds[(i) - (1), (j) - (1)] = Conversions.ToBoolean(cell.Font.Bold);
                                            Italics[(i) - (1), (j) - (1)] = Conversions.ToBoolean(cell.Font.Italic);


                                            if ((((font.Name is System.DBNull)) == (false)))
                                            {
                                                fontNames[(i) - (1), (j) - (1)] = Conversions.ToString(font.Name);
                                            }
                                            else
                                            {
                                                fontNames[(i) - (1), (j) - (1)] = "Calibri";
                                            }

                                            if ((((font.Size is System.DBNull)) == (false)))
                                            {
                                                global::System.Single fontSize = global::System.Convert.ToSingle(font.Size);
                                                fontSizes[(i) - (1), (j) - (1)] = fontSize;
                                            }
                                            else
                                            {
                                                fontSizes[(i) - (1), (j) - (1)] = 11f;
                                            }

                                            if ((cell.Interior.Color is System.DBNull))
                                            {
                                                reds1[(i) - (1), (j) - (1)] = 255;
                                                greens1[(i) - (1), (j) - (1)] = 255;
                                                blues1[(i) - (1), (j) - (1)] = 255;
                                            }
                                            else
                                            {
                                                global::System.Int64 colorValue1 = Conversions.ToLong(cell.Interior.Color);
                                                global::System.Int32 red1 = (global::System.Int32)((colorValue1) % (256L));
                                                global::System.Int32 green1 = (global::System.Int32)(((((colorValue1) / (256L)))) % (256L));
                                                global::System.Int32 blue1 = (global::System.Int32)((((((colorValue1) / (256L)) / (256L)))) % (256L));
                                                reds1[(i) - (1), (j) - (1)] = red1;
                                                greens1[(i) - (1), (j) - (1)] = green1;
                                                blues1[(i) - (1), (j) - (1)] = blue1;
                                            }

                                            if ((cell.Font.Color is System.DBNull))
                                            {
                                                reds2[(i) - (1), (j) - (1)] = 0;
                                                greens2[(i) - (1), (j) - (1)] = 0;
                                                blues2[(i) - (1), (j) - (1)] = 0;
                                            }
                                            else
                                            {
                                                global::System.Int64 colorValue2 = Conversions.ToLong(cell.Font.Color);
                                                global::System.Int32 red2 = (global::System.Int32)((colorValue2) % (256L));
                                                global::System.Int32 green2 = (global::System.Int32)(((((colorValue2) / (256L)))) % (256L));
                                                global::System.Int32 blue2 = (global::System.Int32)((((((colorValue2) / (256L)) / (256L)))) % (256L));
                                                reds2[(i) - (1), (j) - (1)] = red2;
                                                greens2[(i) - (1), (j) - (1)] = green2;
                                                blues2[(i) - (1), (j) - (1)] = blue2;
                                            }
                                        }

                                    }
                                }

                                this.rng.ClearContents();
                                this.rng.ClearFormats();

                                this.rng2.ClearFormats();

                                var loopTo142 = r2;
                                for (i = 1; i <= loopTo142; i++)
                                {
                                    var loopTo143 = c2;
                                    for (j = 1; j <= loopTo143; j++)
                                    {
                                        global::System.Int32 x;
                                        global::System.Int32 y;
                                        x = (((((c2) * ((((i) - (1))))))) + (j));
                                        y = 1;
                                        if (((x) <= ((global::Microsoft.VisualBasic.Information.UBound(Arr, 1)) + (1))))
                                        {
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Value = Arr[(x) - (1), (y) - (1)];

                                            if (((this.CheckBox1.Checked) == (true)))
                                            {

                                                global::Microsoft.Office.Interop.Excel.Range cell2 = (global::Microsoft.Office.Interop.Excel.Range)this.rng2.Cells[(global::System.Object)i, (global::System.Object)j];
                                                var font2 = cell2.Font;

                                                global::System.Single fontSize = fontSizes[(x) - (1), (y) - (1)];

                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Font.Name = fontNames[(x) - (1), (y) - (1)];
                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Font.Size = (global::System.Object)fontSizes[(x) - (1), (y) - (1)];

                                                if (Bolds[(x) - (1), (y) - (1)])
                                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Font.Bold = (global::System.Object)true;
                                                if (Italics[(x) - (1), (y) - (1)])
                                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Font.Italic = (global::System.Object)true;

                                                global::System.Int32 red1 = reds1[(x) - (1), (y) - (1)];
                                                global::System.Int32 green1 = greens1[(x) - (1), (y) - (1)];
                                                global::System.Int32 blue1 = blues1[(x) - (1), (y) - (1)];
                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Interior.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(red1, green1, blue1);

                                                global::System.Int32 red2 = reds2[(x) - (1), (y) - (1)];
                                                global::System.Int32 green2 = greens2[(x) - (1), (y) - (1)];
                                                global::System.Int32 blue2 = blues2[(x) - (1), (y) - (1)];
                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Font.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(red2, green2, blue2);
                                            }

                                        }
                                    }
                                }

                                if (((this.CheckBox1.Checked) == (true)))
                                {

                                    if (((TopBorder8) == (true)))
                                    {
                                        var loopTo144 = this.rng2.Columns.Count;
                                        for (j = 1; j <= loopTo144; j++)
                                        {
                                            this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)8).LineStyle = TopBorder8L;
                                            this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)8).Color = TopBorder8C;
                                            this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)8).Weight = TopBorder8W;
                                        }
                                    }
                                    else
                                    {
                                        var loopTo145 = this.rng2.Columns.Count;
                                        for (j = 1; j <= loopTo145; j++)
                                            this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)8).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                    }

                                    if (((this.rng2.Rows.Count) > (1)))
                                    {
                                        if (((TopBorder9) == (true)))
                                        {
                                            var loopTo146 = this.rng2.Columns.Count;
                                            for (j = 1; j <= loopTo146; j++)
                                            {
                                                this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)9).LineStyle = TopBorder9L;
                                                this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)9).Color = TopBorder9C;
                                                this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)9).Weight = TopBorder9W;
                                            }
                                        }
                                        else
                                        {
                                            var loopTo147 = this.rng2.Columns.Count;
                                            for (j = 1; j <= loopTo147; j++)
                                                this.rng2.Cells[(global::System.Object)this.rng2.Rows.Count, (global::System.Object)j].Borders((global::System.Object)9).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                        }
                                    }

                                    if (((TopBorder7) == (true)))
                                    {
                                        var loopTo148 = this.rng2.Rows.Count;
                                        for (i = 1; i <= loopTo148; i++)
                                        {
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)1].Borders((global::System.Object)7).LineStyle = TopBorder7L;
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)1].Borders((global::System.Object)7).Color = TopBorder7C;
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)1].Borders((global::System.Object)7).Weight = TopBorder7W;
                                        }
                                    }
                                    else
                                    {
                                        var loopTo149 = this.rng2.Rows.Count;
                                        for (i = 1; i <= loopTo149; i++)
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)1].Borders((global::System.Object)7).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                    }

                                    if (((BottomBorder9) == (true)))
                                    {
                                        var loopTo150 = this.rng2.Columns.Count;
                                        for (j = 1; j <= loopTo150; j++)
                                        {
                                            this.rng2.Cells[(global::System.Object)this.rng2.Rows.Count, (global::System.Object)j].Borders((global::System.Object)9).LineStyle = BottomBorder9L;
                                            this.rng2.Cells[(global::System.Object)this.rng2.Rows.Count, (global::System.Object)j].Borders((global::System.Object)9).Color = BottomBorder9C;
                                            this.rng2.Cells[(global::System.Object)this.rng2.Rows.Count, (global::System.Object)j].Borders((global::System.Object)9).Weight = BottomBorder9W;
                                        }
                                    }
                                    else
                                    {
                                        var loopTo151 = this.rng2.Columns.Count;
                                        for (j = 1; j <= loopTo151; j++)
                                            this.rng2.Cells[(global::System.Object)this.rng2.Rows.Count, (global::System.Object)this.rng2.Columns.Count].Borders((global::System.Object)9).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                    }

                                    if (((BottomBorder10) == (true)))
                                    {
                                        var loopTo152 = this.rng2.Rows.Count;
                                        for (i = 1; i <= loopTo152; i++)
                                        {
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)this.rng2.Columns.Count].Borders((global::System.Object)10).LineStyle = BottomBorder10L;
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)this.rng2.Columns.Count].Borders((global::System.Object)10).Color = BottomBorder10C;
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)this.rng2.Columns.Count].Borders((global::System.Object)10).Weight = BottomBorder10W;
                                        }
                                    }
                                    else
                                    {
                                        var loopTo153 = this.rng2.Rows.Count;
                                        for (i = 1; i <= loopTo153; i++)
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)this.rng2.Columns.Count].Borders((global::System.Object)10).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                    }

                                    if (((this.rng.Rows.Count) > (1)))
                                    {

                                        if (((MiddleBorder9) == (true)))
                                        {
                                            var loopTo154 = (this.rng2.Rows.Count) - (1);
                                            for (i = 2; i <= loopTo154; i++)
                                            {
                                                var loopTo155 = this.rng2.Columns.Count;
                                                for (j = 1; j <= loopTo155; j++)
                                                {
                                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)9).LineStyle = MiddleBorder9L;
                                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)9).Color = MiddleBorder9C;
                                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)9).Weight = MiddleBorder9W;
                                                }
                                            }
                                        }
                                        else
                                        {
                                            var loopTo156 = (this.rng2.Rows.Count) - (1);
                                            for (i = 2; i <= loopTo156; i++)
                                            {
                                                var loopTo157 = this.rng2.Columns.Count;
                                                for (j = 1; j <= loopTo157; j++)
                                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)9).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                            }
                                        }

                                    }

                                    if (((this.rng.Columns.Count) > (1)))
                                    {

                                        if (((MiddleBorder10) == (true)))
                                        {
                                            var loopTo158 = (this.rng2.Columns.Count) - (1);
                                            for (j = 1; j <= loopTo158; j++)
                                            {
                                                var loopTo159 = this.rng2.Rows.Count;
                                                for (i = 1; i <= loopTo159; i++)
                                                {
                                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)10).LineStyle = MiddleBorder10L;
                                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)10).Color = MiddleBorder10C;
                                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)10).Weight = MiddleBorder10W;
                                                }
                                            }
                                        }
                                        else
                                        {
                                            var loopTo160 = (this.rng2.Columns.Count) - (1);
                                            for (j = 1; j <= loopTo160; j++)
                                            {
                                                var loopTo161 = this.rng2.Rows.Count;
                                                for (i = 1; i <= loopTo161; i++)
                                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)10).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                            }
                                        }

                                    }

                                }

                            }
                        }

                        else if (X6)
                        {

                            global::System.Int32 r2;
                            global::System.Int32 c2;

                            if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(Operators.ModObject(r, global::Microsoft.VisualBasic.Conversion.Int(this.TextBox2.Text)), 0, false)))
                            {
                                c2 = Conversions.ToInteger(global::Microsoft.VisualBasic.Conversion.Int(Operators.DivideObject(r, global::Microsoft.VisualBasic.Conversion.Int(this.TextBox2.Text))));
                            }
                            else
                            {
                                c2 = Conversions.ToInteger(Operators.AddObject(global::Microsoft.VisualBasic.Conversion.Int(Operators.DivideObject(r, global::Microsoft.VisualBasic.Conversion.Int(this.TextBox2.Text))), 1));
                            }
                            r2 = Conversions.ToInteger(global::Microsoft.VisualBasic.Conversion.Int(this.TextBox2.Text));

                            this.rng2 = this.worksheet2.get_Range(this.rng2.Cells[(global::System.Object)1, (global::System.Object)1], this.rng2.Cells[(global::System.Object)r2, (global::System.Object)c2]);
                            global::System.String rng2Address = this.rng2.get_Address();
                            this.worksheet2.Activate();
                            this.rng2.Select();

                            if (((this.Overlap(this.excelApp, this.worksheet, this.worksheet2, this.rng, this.rng2)) == (false)))
                            {

                                this.rng2.ClearFormats();

                                var loopTo162 = c2;
                                for (j = 1; j <= loopTo162; j++)
                                {
                                    var loopTo163 = r2;
                                    for (i = 1; i <= loopTo163; i++)
                                    {
                                        global::System.Int32 x;
                                        global::System.Int32 y;
                                        x = (((((r2) * ((((j) - (1))))))) + (i));
                                        y = 1;
                                        if (((x) <= (this.rng.Rows.Count)))
                                        {

                                            if (((this.CheckBox1.Checked) == (false)))
                                            {
                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Value = this.rng.Cells[(global::System.Object)x, (global::System.Object)y].Value;
                                            }

                                            else if (((this.CheckBox1.Checked) == (true)))
                                            {

                                                this.rng.Cells[(global::System.Object)x, (global::System.Object)y].Copy();
                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].PasteSpecial(global::Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues);
                                                this.rng2 = this.worksheet2.get_Range(rng2Address);
                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].PasteSpecial(global::Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats);
                                                this.rng2 = this.worksheet2.get_Range(rng2Address);

                                            }
                                        }
                                    }
                                }
                                excelApp.CutCopyMode = global::Microsoft.Office.Interop.Excel.XlCutCopyMode.xlCopy;

                                if (((this.CheckBox1.Checked) == (true)))
                                {
                                    global::Microsoft.Office.Interop.Excel.Range TopCell = (global::Microsoft.Office.Interop.Excel.Range)this.rng.Cells[(global::System.Object)1, (global::System.Object)1];

                                    if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)8].LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                                    {
                                        var loopTo164 = this.rng2.Columns.Count;
                                        for (j = 1; j <= loopTo164; j++)
                                        {
                                            this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)8).LineStyle = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)8].LineStyle;
                                            this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)8).Color = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)8].Color;
                                            this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)8).Weight = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)8].Weight;
                                        }
                                    }
                                    else
                                    {
                                        var loopTo165 = this.rng2.Columns.Count;
                                        for (j = 1; j <= loopTo165; j++)
                                            this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)8).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                    }

                                    if (((this.rng2.Rows.Count) > (1)))
                                    {
                                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                                        {
                                            var loopTo166 = this.rng2.Columns.Count;
                                            for (j = 1; j <= loopTo166; j++)
                                            {
                                                this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)9).LineStyle = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].LineStyle;
                                                this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)9).Color = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].Color;
                                                this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)9).Weight = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].Weight;
                                            }
                                        }
                                        else
                                        {
                                            var loopTo167 = this.rng2.Columns.Count;
                                            for (j = 1; j <= loopTo167; j++)
                                                this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)9).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                        }
                                    }

                                    if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)7].LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                                    {
                                        var loopTo168 = this.rng2.Rows.Count;
                                        for (i = 1; i <= loopTo168; i++)
                                        {
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)1].Borders((global::System.Object)7).LineStyle = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)7].LineStyle;
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)1].Borders((global::System.Object)7).Color = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)7].Color;
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)1].Borders((global::System.Object)7).Weight = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)7].Weight;
                                        }
                                    }
                                    else
                                    {
                                        var loopTo169 = this.rng2.Rows.Count;
                                        for (i = 1; i <= loopTo169; i++)
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)1].Borders((global::System.Object)7).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                    }

                                    global::Microsoft.Office.Interop.Excel.Range BottomCell = (global::Microsoft.Office.Interop.Excel.Range)this.rng.Cells[(global::System.Object)r, (global::System.Object)c];

                                    if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(BottomCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                                    {
                                        var loopTo170 = this.rng2.Columns.Count;
                                        for (j = 1; j <= loopTo170; j++)
                                        {
                                            this.rng2.Cells[(global::System.Object)this.rng2.Rows.Count, (global::System.Object)j].Borders((global::System.Object)9).LineStyle = BottomCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].LineStyle;
                                            this.rng2.Cells[(global::System.Object)this.rng2.Rows.Count, (global::System.Object)j].Borders((global::System.Object)9).Color = BottomCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].Color;
                                            this.rng2.Cells[(global::System.Object)this.rng2.Rows.Count, (global::System.Object)j].Borders((global::System.Object)9).Weight = BottomCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].Weight;
                                        }
                                    }
                                    else
                                    {
                                        var loopTo171 = this.rng2.Columns.Count;
                                        for (j = 1; j <= loopTo171; j++)
                                            this.rng2.Cells[(global::System.Object)this.rng2.Rows.Count, (global::System.Object)this.rng2.Columns.Count].Borders((global::System.Object)9).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                    }

                                    if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(BottomCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)10].LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                                    {
                                        var loopTo172 = this.rng2.Rows.Count;
                                        for (i = 1; i <= loopTo172; i++)
                                        {
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)this.rng2.Columns.Count].Borders((global::System.Object)10).LineStyle = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)10].LineStyle;
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)this.rng2.Columns.Count].Borders((global::System.Object)10).Color = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)10].Color;
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)this.rng2.Columns.Count].Borders((global::System.Object)10).Weight = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)10].Weight;
                                        }
                                    }
                                    else
                                    {
                                        var loopTo173 = this.rng2.Rows.Count;
                                        for (i = 1; i <= loopTo173; i++)
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)this.rng2.Columns.Count].Borders((global::System.Object)10).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                    }

                                    if (((this.rng.Rows.Count) > (1)))
                                    {
                                        global::Microsoft.Office.Interop.Excel.Range MiddleCell = (global::Microsoft.Office.Interop.Excel.Range)this.rng.Cells[(global::System.Object)2, (global::System.Object)1];
                                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(MiddleCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                                        {
                                            var loopTo174 = (this.rng2.Rows.Count) - (1);
                                            for (i = 2; i <= loopTo174; i++)
                                            {
                                                var loopTo175 = this.rng2.Columns.Count;
                                                for (j = 1; j <= loopTo175; j++)
                                                {
                                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)9).LineStyle = MiddleCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].LineStyle;
                                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)9).Color = MiddleCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].Color;
                                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)9).Weight = MiddleCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].Weight;
                                                }
                                            }
                                        }
                                        else
                                        {
                                            var loopTo176 = (this.rng2.Rows.Count) - (1);
                                            for (i = 2; i <= loopTo176; i++)
                                            {
                                                var loopTo177 = this.rng2.Columns.Count;
                                                for (j = 1; j <= loopTo177; j++)
                                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)9).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                            }
                                        }
                                    }

                                    if (((this.rng.Columns.Count) > (1)))
                                    {
                                        global::Microsoft.Office.Interop.Excel.Range MiddleCell = (global::Microsoft.Office.Interop.Excel.Range)this.rng.Cells[(global::System.Object)1, (global::System.Object)2];
                                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(MiddleCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)10].LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                                        {
                                            var loopTo178 = (this.rng2.Columns.Count) - (1);
                                            for (j = 1; j <= loopTo178; j++)
                                            {
                                                var loopTo179 = this.rng2.Rows.Count;
                                                for (i = 1; i <= loopTo179; i++)
                                                {
                                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)10).LineStyle = MiddleCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)10].LineStyle;
                                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)10).Color = MiddleCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)10].Color;
                                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)10).Weight = MiddleCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)10].Weight;
                                                }
                                            }
                                        }
                                        else
                                        {
                                            var loopTo180 = (this.rng2.Columns.Count) - (1);
                                            for (j = 1; j <= loopTo180; j++)
                                            {
                                                var loopTo181 = this.rng2.Rows.Count;
                                                for (i = 1; i <= loopTo181; i++)
                                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)10).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                            }
                                        }
                                    }
                                }
                            }

                            else
                            {

                                var Arr = new global::System.Object[(r), (c)];
                                var Bolds = new global::System.Boolean[(r), (c)];
                                var Italics = new global::System.Boolean[(r), (c)];
                                var fontNames = new global::System.String[(r), (c)];
                                var fontSizes = new global::System.Single[(r), (c)];
                                var reds1 = new global::System.Int32[(r), (c)];
                                var reds2 = new global::System.Int32[(r), (c)];
                                var greens1 = new global::System.Int32[(r), (c)];
                                var greens2 = new global::System.Int32[(r), (c)];
                                var blues1 = new global::System.Int32[(r), (c)];
                                var blues2 = new global::System.Int32[(r), (c)];

                                global::System.Boolean TopBorder7;
                                global::System.Object TopBorder7L;
                                global::System.Object TopBorder7C;
                                global::System.Object TopBorder7W;

                                global::System.Boolean TopBorder8;
                                global::System.Object TopBorder8L;
                                global::System.Object TopBorder8C;
                                global::System.Object TopBorder8W;

                                global::System.Boolean TopBorder9;
                                global::System.Object TopBorder9L;
                                global::System.Object TopBorder9C;
                                global::System.Object TopBorder9W;

                                global::System.Boolean BottomBorder9;
                                global::System.Object BottomBorder9L;
                                global::System.Object BottomBorder9C;
                                global::System.Object BottomBorder9W;

                                global::System.Boolean BottomBorder10;
                                global::System.Object BottomBorder10L;
                                global::System.Object BottomBorder10C;
                                global::System.Object BottomBorder10W;

                                global::System.Boolean MiddleBorder9;
                                global::System.Object MiddleBorder9L;
                                global::System.Object MiddleBorder9C;
                                global::System.Object MiddleBorder9W;

                                global::System.Boolean MiddleBorder10;
                                global::System.Object MiddleBorder10L;
                                global::System.Object MiddleBorder10C;
                                global::System.Object MiddleBorder10W;

                                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)7).LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                                {
                                    TopBorder7 = true;
                                    TopBorder7L = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)7).LineStyle;
                                    TopBorder7C = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)7).Color;
                                    TopBorder7W = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)7).Weight;
                                }
                                else
                                {
                                    TopBorder7 = false;
                                    TopBorder7L = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)7).LineStyle;
                                    TopBorder7C = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)7).Color;
                                    TopBorder7W = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)7).Weight;
                                }

                                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)8).LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                                {
                                    TopBorder8 = true;
                                    TopBorder8L = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)8).LineStyle;
                                    TopBorder8C = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)8).Color;
                                    TopBorder8W = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)8).Weight;
                                }
                                else
                                {
                                    TopBorder8 = false;
                                    TopBorder8L = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)8).LineStyle;
                                    TopBorder8C = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)8).Color;
                                    TopBorder8W = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)8).Weight;
                                }

                                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)9).LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                                {
                                    TopBorder9 = true;
                                    TopBorder9L = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)9).LineStyle;
                                    TopBorder9C = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)9).Color;
                                    TopBorder9W = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)9).Weight;
                                }
                                else
                                {
                                    TopBorder9 = false;
                                    TopBorder9L = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)9).LineStyle;
                                    TopBorder9C = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)9).Color;
                                    TopBorder9W = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)9).Weight;
                                }

                                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)9).LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                                {
                                    BottomBorder9 = true;
                                    BottomBorder9L = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)9).LineStyle;
                                    BottomBorder9C = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)9).Color;
                                    BottomBorder9W = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)9).Weight;
                                }
                                else
                                {
                                    BottomBorder9 = false;
                                    BottomBorder9L = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)9).LineStyle;
                                    BottomBorder9C = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)9).Color;
                                    BottomBorder9W = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)9).Weight;
                                }

                                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)10).LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                                {
                                    BottomBorder10 = true;
                                    BottomBorder10L = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)10).LineStyle;
                                    BottomBorder10C = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)10).Color;
                                    BottomBorder10W = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)10).Weight;
                                }
                                else
                                {
                                    BottomBorder10 = false;
                                    BottomBorder10L = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)10).LineStyle;
                                    BottomBorder10C = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)10).Color;
                                    BottomBorder10W = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)10).Weight;
                                }

                                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(this.rng.Cells[(global::System.Object)2, (global::System.Object)1].Borders((global::System.Object)9).LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                                {
                                    MiddleBorder9 = true;
                                    MiddleBorder9L = this.rng.Cells[(global::System.Object)2, (global::System.Object)1].Borders((global::System.Object)9).LineStyle;
                                    MiddleBorder9C = this.rng.Cells[(global::System.Object)2, (global::System.Object)1].Borders((global::System.Object)9).Color;
                                    MiddleBorder9W = this.rng.Cells[(global::System.Object)2, (global::System.Object)1].Borders((global::System.Object)9).Weight;
                                }
                                else
                                {
                                    MiddleBorder9 = false;
                                    MiddleBorder9L = this.rng.Cells[(global::System.Object)2, (global::System.Object)1].Borders((global::System.Object)9).LineStyle;
                                    MiddleBorder9C = this.rng.Cells[(global::System.Object)2, (global::System.Object)1].Borders((global::System.Object)9).Color;
                                    MiddleBorder9W = this.rng.Cells[(global::System.Object)2, (global::System.Object)1].Borders((global::System.Object)9).Weight;
                                }

                                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(this.rng.Cells[(global::System.Object)1, (global::System.Object)2].Borders((global::System.Object)10).LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                                {
                                    MiddleBorder10 = true;
                                    MiddleBorder10L = this.rng.Cells[(global::System.Object)1, (global::System.Object)2].Borders((global::System.Object)10).LineStyle;
                                    MiddleBorder10C = this.rng.Cells[(global::System.Object)1, (global::System.Object)2].Borders((global::System.Object)10).Color;
                                    MiddleBorder10W = this.rng.Cells[(global::System.Object)1, (global::System.Object)2].Borders((global::System.Object)10).Weight;
                                }
                                else
                                {
                                    MiddleBorder10 = false;
                                    MiddleBorder10L = this.rng.Cells[(global::System.Object)1, (global::System.Object)2].Borders((global::System.Object)10).LineStyle;
                                    MiddleBorder10C = this.rng.Cells[(global::System.Object)1, (global::System.Object)2].Borders((global::System.Object)10).Color;
                                    MiddleBorder10W = this.rng.Cells[(global::System.Object)1, (global::System.Object)2].Borders((global::System.Object)10).Weight;
                                }

                                var loopTo182 = r;
                                for (i = 1; i <= loopTo182; i++)
                                {
                                    var loopTo183 = c;
                                    for (j = 1; j <= loopTo183; j++)
                                    {
                                        Arr[(i) - (1), (j) - (1)] = this.rng.Cells[(global::System.Object)i, (global::System.Object)j].Value;

                                        if (((this.CheckBox1.Checked) == (true)))
                                        {

                                            global::Microsoft.Office.Interop.Excel.Range cell = (global::Microsoft.Office.Interop.Excel.Range)this.rng.Cells[(global::System.Object)i, (global::System.Object)j];
                                            var font = cell.Font;

                                            Bolds[(i) - (1), (j) - (1)] = Conversions.ToBoolean(cell.Font.Bold);
                                            Italics[(i) - (1), (j) - (1)] = Conversions.ToBoolean(cell.Font.Italic);


                                            if ((((font.Name is System.DBNull)) == (false)))
                                            {
                                                fontNames[(i) - (1), (j) - (1)] = Conversions.ToString(font.Name);
                                            }
                                            else
                                            {
                                                fontNames[(i) - (1), (j) - (1)] = "Calibri";
                                            }

                                            if ((((font.Size is System.DBNull)) == (false)))
                                            {
                                                global::System.Single fontSize = global::System.Convert.ToSingle(font.Size);
                                                fontSizes[(i) - (1), (j) - (1)] = fontSize;
                                            }
                                            else
                                            {
                                                fontSizes[(i) - (1), (j) - (1)] = 11f;
                                            }

                                            if ((cell.Interior.Color is System.DBNull))
                                            {
                                                reds1[(i) - (1), (j) - (1)] = 255;
                                                greens1[(i) - (1), (j) - (1)] = 255;
                                                blues1[(i) - (1), (j) - (1)] = 255;
                                            }
                                            else
                                            {
                                                global::System.Int64 colorValue1 = Conversions.ToLong(cell.Interior.Color);
                                                global::System.Int32 red1 = (global::System.Int32)((colorValue1) % (256L));
                                                global::System.Int32 green1 = (global::System.Int32)(((((colorValue1) / (256L)))) % (256L));
                                                global::System.Int32 blue1 = (global::System.Int32)((((((colorValue1) / (256L)) / (256L)))) % (256L));
                                                reds1[(i) - (1), (j) - (1)] = red1;
                                                greens1[(i) - (1), (j) - (1)] = green1;
                                                blues1[(i) - (1), (j) - (1)] = blue1;
                                            }

                                            if ((cell.Font.Color is System.DBNull))
                                            {
                                                reds2[(i) - (1), (j) - (1)] = 0;
                                                greens2[(i) - (1), (j) - (1)] = 0;
                                                blues2[(i) - (1), (j) - (1)] = 0;
                                            }
                                            else
                                            {
                                                global::System.Int64 colorValue2 = Conversions.ToLong(cell.Font.Color);
                                                global::System.Int32 red2 = (global::System.Int32)((colorValue2) % (256L));
                                                global::System.Int32 green2 = (global::System.Int32)(((((colorValue2) / (256L)))) % (256L));
                                                global::System.Int32 blue2 = (global::System.Int32)((((((colorValue2) / (256L)) / (256L)))) % (256L));
                                                reds2[(i) - (1), (j) - (1)] = red2;
                                                greens2[(i) - (1), (j) - (1)] = green2;
                                                blues2[(i) - (1), (j) - (1)] = blue2;
                                            }
                                        }

                                    }
                                }

                                this.rng.ClearContents();
                                this.rng.ClearFormats();

                                this.rng2.ClearFormats();

                                var loopTo184 = c2;
                                for (j = 1; j <= loopTo184; j++)
                                {
                                    var loopTo185 = r2;
                                    for (i = 1; i <= loopTo185; i++)
                                    {
                                        global::System.Int32 x;
                                        global::System.Int32 y;
                                        x = (((((r2) * ((((j) - (1))))))) + (i));
                                        y = 1;
                                        if (((x) <= ((global::Microsoft.VisualBasic.Information.UBound(Arr, 1)) + (1))))
                                        {
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Value = Arr[(x) - (1), (y) - (1)];

                                            if (((this.CheckBox1.Checked) == (true)))
                                            {

                                                global::Microsoft.Office.Interop.Excel.Range cell2 = (global::Microsoft.Office.Interop.Excel.Range)this.rng2.Cells[(global::System.Object)i, (global::System.Object)j];
                                                var font2 = cell2.Font;

                                                global::System.Single fontSize = fontSizes[(x) - (1), (y) - (1)];

                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Font.Name = fontNames[(x) - (1), (y) - (1)];
                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Font.Size = (global::System.Object)fontSizes[(x) - (1), (y) - (1)];

                                                if (Bolds[(x) - (1), (y) - (1)])
                                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Font.Bold = (global::System.Object)true;
                                                if (Italics[(x) - (1), (y) - (1)])
                                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Font.Italic = (global::System.Object)true;

                                                global::System.Int32 red1 = reds1[(x) - (1), (y) - (1)];
                                                global::System.Int32 green1 = greens1[(x) - (1), (y) - (1)];
                                                global::System.Int32 blue1 = blues1[(x) - (1), (y) - (1)];
                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Interior.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(red1, green1, blue1);

                                                global::System.Int32 red2 = reds2[(x) - (1), (y) - (1)];
                                                global::System.Int32 green2 = greens2[(x) - (1), (y) - (1)];
                                                global::System.Int32 blue2 = blues2[(x) - (1), (y) - (1)];
                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Font.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(red2, green2, blue2);
                                            }
                                        }
                                    }
                                }

                                if (((this.CheckBox1.Checked) == (true)))
                                {

                                    if (((TopBorder8) == (true)))
                                    {
                                        var loopTo186 = this.rng2.Columns.Count;
                                        for (j = 1; j <= loopTo186; j++)
                                        {
                                            this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)8).LineStyle = TopBorder8L;
                                            this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)8).Color = TopBorder8C;
                                            this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)8).Weight = TopBorder8W;
                                        }
                                    }
                                    else
                                    {
                                        var loopTo187 = this.rng2.Columns.Count;
                                        for (j = 1; j <= loopTo187; j++)
                                            this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)8).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                    }

                                    if (((this.rng2.Rows.Count) > (1)))
                                    {
                                        if (((TopBorder9) == (true)))
                                        {
                                            var loopTo188 = this.rng2.Columns.Count;
                                            for (j = 1; j <= loopTo188; j++)
                                            {
                                                this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)9).LineStyle = TopBorder9L;
                                                this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)9).Color = TopBorder9C;
                                                this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)9).Weight = TopBorder9W;
                                            }
                                        }
                                        else
                                        {
                                            var loopTo189 = this.rng2.Columns.Count;
                                            for (j = 1; j <= loopTo189; j++)
                                                this.rng2.Cells[(global::System.Object)this.rng2.Rows.Count, (global::System.Object)j].Borders((global::System.Object)9).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                        }
                                    }

                                    if (((TopBorder7) == (true)))
                                    {
                                        var loopTo190 = this.rng2.Rows.Count;
                                        for (i = 1; i <= loopTo190; i++)
                                        {
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)1].Borders((global::System.Object)7).LineStyle = TopBorder7L;
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)1].Borders((global::System.Object)7).Color = TopBorder7C;
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)1].Borders((global::System.Object)7).Weight = TopBorder7W;
                                        }
                                    }
                                    else
                                    {
                                        var loopTo191 = this.rng2.Rows.Count;
                                        for (i = 1; i <= loopTo191; i++)
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)1].Borders((global::System.Object)7).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                    }

                                    if (((BottomBorder9) == (true)))
                                    {
                                        var loopTo192 = this.rng2.Columns.Count;
                                        for (j = 1; j <= loopTo192; j++)
                                        {
                                            this.rng2.Cells[(global::System.Object)this.rng2.Rows.Count, (global::System.Object)j].Borders((global::System.Object)9).LineStyle = BottomBorder9L;
                                            this.rng2.Cells[(global::System.Object)this.rng2.Rows.Count, (global::System.Object)j].Borders((global::System.Object)9).Color = BottomBorder9C;
                                            this.rng2.Cells[(global::System.Object)this.rng2.Rows.Count, (global::System.Object)j].Borders((global::System.Object)9).Weight = BottomBorder9W;
                                        }
                                    }
                                    else
                                    {
                                        var loopTo193 = this.rng2.Columns.Count;
                                        for (j = 1; j <= loopTo193; j++)
                                            this.rng2.Cells[(global::System.Object)this.rng2.Rows.Count, (global::System.Object)this.rng2.Columns.Count].Borders((global::System.Object)9).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                    }

                                    if (((BottomBorder10) == (true)))
                                    {
                                        var loopTo194 = this.rng2.Rows.Count;
                                        for (i = 1; i <= loopTo194; i++)
                                        {
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)this.rng2.Columns.Count].Borders((global::System.Object)10).LineStyle = BottomBorder10L;
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)this.rng2.Columns.Count].Borders((global::System.Object)10).Color = BottomBorder10C;
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)this.rng2.Columns.Count].Borders((global::System.Object)10).Weight = BottomBorder10W;
                                        }
                                    }
                                    else
                                    {
                                        var loopTo195 = this.rng2.Rows.Count;
                                        for (i = 1; i <= loopTo195; i++)
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)this.rng2.Columns.Count].Borders((global::System.Object)10).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                    }

                                    if (((this.rng.Rows.Count) > (1)))
                                    {

                                        if (((MiddleBorder9) == (true)))
                                        {
                                            var loopTo196 = (this.rng2.Rows.Count) - (1);
                                            for (i = 2; i <= loopTo196; i++)
                                            {
                                                var loopTo197 = this.rng2.Columns.Count;
                                                for (j = 1; j <= loopTo197; j++)
                                                {
                                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)9).LineStyle = MiddleBorder9L;
                                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)9).Color = MiddleBorder9C;
                                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)9).Weight = MiddleBorder9W;
                                                }
                                            }
                                        }
                                        else
                                        {
                                            var loopTo198 = (this.rng2.Rows.Count) - (1);
                                            for (i = 2; i <= loopTo198; i++)
                                            {
                                                var loopTo199 = this.rng2.Columns.Count;
                                                for (j = 1; j <= loopTo199; j++)
                                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)9).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                            }
                                        }

                                    }

                                    if (((this.rng.Columns.Count) > (1)))
                                    {

                                        if (((MiddleBorder10) == (true)))
                                        {
                                            var loopTo200 = (this.rng2.Columns.Count) - (1);
                                            for (j = 1; j <= loopTo200; j++)
                                            {
                                                var loopTo201 = this.rng2.Rows.Count;
                                                for (i = 1; i <= loopTo201; i++)
                                                {
                                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)10).LineStyle = MiddleBorder10L;
                                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)10).Color = MiddleBorder10C;
                                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)10).Weight = MiddleBorder10W;
                                                }
                                            }
                                        }
                                        else
                                        {
                                            var loopTo202 = (this.rng2.Columns.Count) - (1);
                                            for (j = 1; j <= loopTo202; j++)
                                            {
                                                var loopTo203 = this.rng2.Rows.Count;
                                                for (i = 1; i <= loopTo203; i++)
                                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)10).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                            }
                                        }

                                    }

                                }

                            }

                        }
                    }

                    else
                    {
                        global::System.Windows.Forms.MessageBox.Show("Select One Separator.", "Error", global::System.Windows.Forms.MessageBoxButtons.OK, global::System.Windows.Forms.MessageBoxIcon.Error);
                        return;

                    }
                }

                else if (X4)
                {

                    if (((X7) & ((((X5) | (X6))))))
                    {

                        var r2 = default(global::System.Int32);
                        var c2 = default(global::System.Int32);

                        global::System.Int32[] BreakPoints;
                        BreakPoints = (global::System.Int32[])this.GetBreakPoints(this.rng, 1);

                        global::System.Int32[] lengths;
                        lengths = (global::System.Int32[])this.GetLengths(BreakPoints);

                        if (X5)
                        {
                            r2 = ((global::Microsoft.VisualBasic.Information.UBound(BreakPoints)) + (1));
                            c2 = Conversions.ToInteger(this.MaxValue(lengths));
                        }
                        else if (X6)
                        {
                            c2 = ((global::Microsoft.VisualBasic.Information.UBound(BreakPoints)) + (1));
                            r2 = Conversions.ToInteger(this.MaxValue(lengths));
                        }

                        this.rng2 = this.worksheet2.get_Range(this.rng2.Cells[(global::System.Object)1, (global::System.Object)1], this.rng2.Cells[(global::System.Object)r2, (global::System.Object)c2]);
                        global::System.String rng2Address = this.rng2.get_Address();
                        this.worksheet2.Activate();
                        this.rng2.Select();

                        if (((this.Overlap(this.excelApp, this.worksheet, this.worksheet2, this.rng, this.rng2)) == (false)))
                        {

                            this.rng2.ClearFormats();

                            if (X5)
                            {
                                global::System.Int32 iColumn;
                                iColumn = 0;
                                var loopTo204 = r2;
                                for (i = 1; i <= loopTo204; i++)
                                {
                                    var loopTo205 = c2;
                                    for (j = 1; j <= loopTo205; j++)
                                    {
                                        global::System.Int32 x;
                                        global::System.Int32 y;
                                        x = 1;
                                        y = ((iColumn) + (j));
                                        if (((y) < (BreakPoints[(i) - (1)])))
                                        {
                                            if (((this.CheckBox1.Checked) == (false)))
                                            {
                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Value = this.rng.Cells[(global::System.Object)x, (global::System.Object)y].Value;
                                            }

                                            else if (((this.CheckBox1.Checked) == (true)))
                                            {

                                                this.rng.Cells[(global::System.Object)x, (global::System.Object)y].Copy();
                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].PasteSpecial(global::Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues);
                                                this.rng2 = this.worksheet2.get_Range(rng2Address);
                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].PasteSpecial(global::Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats);
                                                this.rng2 = this.worksheet2.get_Range(rng2Address);

                                            }
                                        }
                                    }
                                    iColumn = BreakPoints[(i) - (1)];
                                }
                                excelApp.CutCopyMode = global::Microsoft.Office.Interop.Excel.XlCutCopyMode.xlCopy;
                            }
                            else if (X6)
                            {
                                global::System.Int32 iColumn;
                                iColumn = 0;
                                var loopTo206 = c2;
                                for (j = 1; j <= loopTo206; j++)
                                {
                                    var loopTo207 = r2;
                                    for (i = 1; i <= loopTo207; i++)
                                    {
                                        global::System.Int32 x;
                                        global::System.Int32 y;
                                        x = 1;
                                        y = ((iColumn) + (i));
                                        if (((y) < (BreakPoints[(j) - (1)])))
                                        {

                                            if (((this.CheckBox1.Checked) == (false)))
                                            {
                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Value = this.rng.Cells[(global::System.Object)x, (global::System.Object)y].Value;
                                            }

                                            else if (((this.CheckBox1.Checked) == (true)))
                                            {

                                                this.rng.Cells[(global::System.Object)x, (global::System.Object)y].Copy();
                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].PasteSpecial(global::Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues);
                                                this.rng2 = this.worksheet2.get_Range(rng2Address);
                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].PasteSpecial(global::Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats);
                                                this.rng2 = this.worksheet2.get_Range(rng2Address);
                                            }

                                        }
                                    }
                                    iColumn = BreakPoints[(j) - (1)];
                                }
                                excelApp.CutCopyMode = global::Microsoft.Office.Interop.Excel.XlCutCopyMode.xlCopy;
                            }

                            if (((this.CheckBox1.Checked) == (true)))
                            {
                                global::Microsoft.Office.Interop.Excel.Range TopCell = (global::Microsoft.Office.Interop.Excel.Range)this.rng.Cells[(global::System.Object)1, (global::System.Object)1];

                                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)8].LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                                {
                                    var loopTo208 = this.rng2.Columns.Count;
                                    for (j = 1; j <= loopTo208; j++)
                                    {
                                        this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)8).LineStyle = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)8].LineStyle;
                                        this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)8).Color = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)8].Color;
                                        this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)8).Weight = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)8].Weight;
                                    }
                                }
                                else
                                {
                                    var loopTo209 = this.rng2.Columns.Count;
                                    for (j = 1; j <= loopTo209; j++)
                                        this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)8).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                }

                                if (((this.rng2.Rows.Count) > (1)))
                                {
                                    if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                                    {
                                        var loopTo210 = this.rng2.Columns.Count;
                                        for (j = 1; j <= loopTo210; j++)
                                        {
                                            this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)9).LineStyle = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].LineStyle;
                                            this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)9).Color = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].Color;
                                            this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)9).Weight = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].Weight;
                                        }
                                    }
                                    else
                                    {
                                        var loopTo211 = this.rng2.Columns.Count;
                                        for (j = 1; j <= loopTo211; j++)
                                            this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)9).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                    }
                                }

                                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)7].LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                                {
                                    var loopTo212 = this.rng2.Rows.Count;
                                    for (i = 1; i <= loopTo212; i++)
                                    {
                                        this.rng2.Cells[(global::System.Object)i, (global::System.Object)1].Borders((global::System.Object)7).LineStyle = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)7].LineStyle;
                                        this.rng2.Cells[(global::System.Object)i, (global::System.Object)1].Borders((global::System.Object)7).Color = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)7].Color;
                                        this.rng2.Cells[(global::System.Object)i, (global::System.Object)1].Borders((global::System.Object)7).Weight = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)7].Weight;
                                    }
                                }
                                else
                                {
                                    var loopTo213 = this.rng2.Rows.Count;
                                    for (i = 1; i <= loopTo213; i++)
                                        this.rng2.Cells[(global::System.Object)i, (global::System.Object)1].Borders((global::System.Object)7).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                }

                                global::Microsoft.Office.Interop.Excel.Range BottomCell = (global::Microsoft.Office.Interop.Excel.Range)this.rng.Cells[(global::System.Object)r, (global::System.Object)c];

                                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(BottomCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                                {
                                    var loopTo214 = this.rng2.Columns.Count;
                                    for (j = 1; j <= loopTo214; j++)
                                    {
                                        this.rng2.Cells[(global::System.Object)this.rng2.Rows.Count, (global::System.Object)j].Borders((global::System.Object)9).LineStyle = BottomCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].LineStyle;
                                        this.rng2.Cells[(global::System.Object)this.rng2.Rows.Count, (global::System.Object)j].Borders((global::System.Object)9).Color = BottomCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].Color;
                                        this.rng2.Cells[(global::System.Object)this.rng2.Rows.Count, (global::System.Object)j].Borders((global::System.Object)9).Weight = BottomCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].Weight;
                                    }
                                }
                                else
                                {
                                    var loopTo215 = this.rng2.Columns.Count;
                                    for (j = 1; j <= loopTo215; j++)
                                        this.rng2.Cells[(global::System.Object)this.rng2.Rows.Count, (global::System.Object)this.rng2.Columns.Count].Borders((global::System.Object)9).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                }

                                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(BottomCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)10].LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                                {
                                    var loopTo216 = this.rng2.Rows.Count;
                                    for (i = 1; i <= loopTo216; i++)
                                    {
                                        this.rng2.Cells[(global::System.Object)i, (global::System.Object)this.rng2.Columns.Count].Borders((global::System.Object)10).LineStyle = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)10].LineStyle;
                                        this.rng2.Cells[(global::System.Object)i, (global::System.Object)this.rng2.Columns.Count].Borders((global::System.Object)10).Color = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)10].Color;
                                        this.rng2.Cells[(global::System.Object)i, (global::System.Object)this.rng2.Columns.Count].Borders((global::System.Object)10).Weight = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)10].Weight;
                                    }
                                }
                                else
                                {
                                    var loopTo217 = this.rng2.Rows.Count;
                                    for (i = 1; i <= loopTo217; i++)
                                        this.rng2.Cells[(global::System.Object)i, (global::System.Object)this.rng2.Columns.Count].Borders((global::System.Object)10).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                }

                                if (((this.rng.Rows.Count) > (1)))
                                {
                                    global::Microsoft.Office.Interop.Excel.Range MiddleCell = (global::Microsoft.Office.Interop.Excel.Range)this.rng.Cells[(global::System.Object)2, (global::System.Object)1];
                                    if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(MiddleCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                                    {
                                        var loopTo218 = (this.rng2.Rows.Count) - (1);
                                        for (i = 2; i <= loopTo218; i++)
                                        {
                                            var loopTo219 = this.rng2.Columns.Count;
                                            for (j = 1; j <= loopTo219; j++)
                                            {
                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)9).LineStyle = MiddleCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].LineStyle;
                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)9).Color = MiddleCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].Color;
                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)9).Weight = MiddleCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].Weight;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        var loopTo220 = (this.rng2.Rows.Count) - (1);
                                        for (i = 2; i <= loopTo220; i++)
                                        {
                                            var loopTo221 = this.rng2.Columns.Count;
                                            for (j = 1; j <= loopTo221; j++)
                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)9).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                        }
                                    }
                                }

                                if (((this.rng.Columns.Count) > (1)))
                                {
                                    global::Microsoft.Office.Interop.Excel.Range MiddleCell = (global::Microsoft.Office.Interop.Excel.Range)this.rng.Cells[(global::System.Object)1, (global::System.Object)2];
                                    if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(MiddleCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)10].LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                                    {
                                        var loopTo222 = (this.rng2.Columns.Count) - (1);
                                        for (j = 1; j <= loopTo222; j++)
                                        {
                                            var loopTo223 = this.rng2.Rows.Count;
                                            for (i = 1; i <= loopTo223; i++)
                                            {
                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)10).LineStyle = MiddleCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)10].LineStyle;
                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)10).Color = MiddleCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)10].Color;
                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)10).Weight = MiddleCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)10].Weight;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        var loopTo224 = (this.rng2.Columns.Count) - (1);
                                        for (j = 1; j <= loopTo224; j++)
                                        {
                                            var loopTo225 = this.rng2.Rows.Count;
                                            for (i = 1; i <= loopTo225; i++)
                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)10).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                        }
                                    }
                                }
                            }
                        }

                        else
                        {

                            var Arr = new global::System.Object[(r), (c)];
                            var Bolds = new global::System.Boolean[(r), (c)];
                            var Italics = new global::System.Boolean[(r), (c)];
                            var fontNames = new global::System.String[(r), (c)];
                            var fontSizes = new global::System.Single[(r), (c)];
                            var reds1 = new global::System.Int32[(r), (c)];
                            var reds2 = new global::System.Int32[(r), (c)];
                            var greens1 = new global::System.Int32[(r), (c)];
                            var greens2 = new global::System.Int32[(r), (c)];
                            var blues1 = new global::System.Int32[(r), (c)];
                            var blues2 = new global::System.Int32[(r), (c)];

                            global::System.Boolean TopBorder7;
                            global::System.Object TopBorder7L;
                            global::System.Object TopBorder7C;
                            global::System.Object TopBorder7W;

                            global::System.Boolean TopBorder8;
                            global::System.Object TopBorder8L;
                            global::System.Object TopBorder8C;
                            global::System.Object TopBorder8W;

                            global::System.Boolean TopBorder9;
                            global::System.Object TopBorder9L;
                            global::System.Object TopBorder9C;
                            global::System.Object TopBorder9W;

                            global::System.Boolean BottomBorder9;
                            global::System.Object BottomBorder9L;
                            global::System.Object BottomBorder9C;
                            global::System.Object BottomBorder9W;

                            global::System.Boolean BottomBorder10;
                            global::System.Object BottomBorder10L;
                            global::System.Object BottomBorder10C;
                            global::System.Object BottomBorder10W;

                            global::System.Boolean MiddleBorder9;
                            global::System.Object MiddleBorder9L;
                            global::System.Object MiddleBorder9C;
                            global::System.Object MiddleBorder9W;

                            global::System.Boolean MiddleBorder10;
                            global::System.Object MiddleBorder10L;
                            global::System.Object MiddleBorder10C;
                            global::System.Object MiddleBorder10W;

                            if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)7).LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                            {
                                TopBorder7 = true;
                                TopBorder7L = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)7).LineStyle;
                                TopBorder7C = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)7).Color;
                                TopBorder7W = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)7).Weight;
                            }
                            else
                            {
                                TopBorder7 = false;
                                TopBorder7L = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)7).LineStyle;
                                TopBorder7C = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)7).Color;
                                TopBorder7W = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)7).Weight;
                            }

                            if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)8).LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                            {
                                TopBorder8 = true;
                                TopBorder8L = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)8).LineStyle;
                                TopBorder8C = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)8).Color;
                                TopBorder8W = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)8).Weight;
                            }
                            else
                            {
                                TopBorder8 = false;
                                TopBorder8L = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)8).LineStyle;
                                TopBorder8C = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)8).Color;
                                TopBorder8W = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)8).Weight;
                            }

                            if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)9).LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                            {
                                TopBorder9 = true;
                                TopBorder9L = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)9).LineStyle;
                                TopBorder9C = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)9).Color;
                                TopBorder9W = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)9).Weight;
                            }
                            else
                            {
                                TopBorder9 = false;
                                TopBorder9L = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)9).LineStyle;
                                TopBorder9C = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)9).Color;
                                TopBorder9W = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)9).Weight;
                            }

                            if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)9).LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                            {
                                BottomBorder9 = true;
                                BottomBorder9L = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)9).LineStyle;
                                BottomBorder9C = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)9).Color;
                                BottomBorder9W = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)9).Weight;
                            }
                            else
                            {
                                BottomBorder9 = false;
                                BottomBorder9L = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)9).LineStyle;
                                BottomBorder9C = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)9).Color;
                                BottomBorder9W = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)9).Weight;
                            }

                            if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)10).LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                            {
                                BottomBorder10 = true;
                                BottomBorder10L = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)10).LineStyle;
                                BottomBorder10C = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)10).Color;
                                BottomBorder10W = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)10).Weight;
                            }
                            else
                            {
                                BottomBorder10 = false;
                                BottomBorder10L = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)10).LineStyle;
                                BottomBorder10C = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)10).Color;
                                BottomBorder10W = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)10).Weight;
                            }

                            if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(this.rng.Cells[(global::System.Object)2, (global::System.Object)1].Borders((global::System.Object)9).LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                            {
                                MiddleBorder9 = true;
                                MiddleBorder9L = this.rng.Cells[(global::System.Object)2, (global::System.Object)1].Borders((global::System.Object)9).LineStyle;
                                MiddleBorder9C = this.rng.Cells[(global::System.Object)2, (global::System.Object)1].Borders((global::System.Object)9).Color;
                                MiddleBorder9W = this.rng.Cells[(global::System.Object)2, (global::System.Object)1].Borders((global::System.Object)9).Weight;
                            }
                            else
                            {
                                MiddleBorder9 = false;
                                MiddleBorder9L = this.rng.Cells[(global::System.Object)2, (global::System.Object)1].Borders((global::System.Object)9).LineStyle;
                                MiddleBorder9C = this.rng.Cells[(global::System.Object)2, (global::System.Object)1].Borders((global::System.Object)9).Color;
                                MiddleBorder9W = this.rng.Cells[(global::System.Object)2, (global::System.Object)1].Borders((global::System.Object)9).Weight;
                            }

                            if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(this.rng.Cells[(global::System.Object)1, (global::System.Object)2].Borders((global::System.Object)10).LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                            {
                                MiddleBorder10 = true;
                                MiddleBorder10L = this.rng.Cells[(global::System.Object)1, (global::System.Object)2].Borders((global::System.Object)10).LineStyle;
                                MiddleBorder10C = this.rng.Cells[(global::System.Object)1, (global::System.Object)2].Borders((global::System.Object)10).Color;
                                MiddleBorder10W = this.rng.Cells[(global::System.Object)1, (global::System.Object)2].Borders((global::System.Object)10).Weight;
                            }
                            else
                            {
                                MiddleBorder10 = false;
                                MiddleBorder10L = this.rng.Cells[(global::System.Object)1, (global::System.Object)2].Borders((global::System.Object)10).LineStyle;
                                MiddleBorder10C = this.rng.Cells[(global::System.Object)1, (global::System.Object)2].Borders((global::System.Object)10).Color;
                                MiddleBorder10W = this.rng.Cells[(global::System.Object)1, (global::System.Object)2].Borders((global::System.Object)10).Weight;
                            }

                            var loopTo226 = r;
                            for (i = 1; i <= loopTo226; i++)
                            {
                                var loopTo227 = c;
                                for (j = 1; j <= loopTo227; j++)
                                {
                                    Arr[(i) - (1), (j) - (1)] = this.rng.Cells[(global::System.Object)i, (global::System.Object)j].Value;

                                    if (((this.CheckBox1.Checked) == (true)))
                                    {

                                        global::Microsoft.Office.Interop.Excel.Range cell = (global::Microsoft.Office.Interop.Excel.Range)this.rng.Cells[(global::System.Object)i, (global::System.Object)j];
                                        var font = cell.Font;

                                        Bolds[(i) - (1), (j) - (1)] = Conversions.ToBoolean(cell.Font.Bold);
                                        Italics[(i) - (1), (j) - (1)] = Conversions.ToBoolean(cell.Font.Italic);


                                        if ((((font.Name is System.DBNull)) == (false)))
                                        {
                                            fontNames[(i) - (1), (j) - (1)] = Conversions.ToString(font.Name);
                                        }
                                        else
                                        {
                                            fontNames[(i) - (1), (j) - (1)] = "Calibri";
                                        }

                                        if ((((font.Size is System.DBNull)) == (false)))
                                        {
                                            global::System.Single fontSize = global::System.Convert.ToSingle(font.Size);
                                            fontSizes[(i) - (1), (j) - (1)] = fontSize;
                                        }
                                        else
                                        {
                                            fontSizes[(i) - (1), (j) - (1)] = 11f;
                                        }

                                        if ((cell.Interior.Color is System.DBNull))
                                        {
                                            reds1[(i) - (1), (j) - (1)] = 255;
                                            greens1[(i) - (1), (j) - (1)] = 255;
                                            blues1[(i) - (1), (j) - (1)] = 255;
                                        }
                                        else
                                        {
                                            global::System.Int64 colorValue1 = Conversions.ToLong(cell.Interior.Color);
                                            global::System.Int32 red1 = (global::System.Int32)((colorValue1) % (256L));
                                            global::System.Int32 green1 = (global::System.Int32)(((((colorValue1) / (256L)))) % (256L));
                                            global::System.Int32 blue1 = (global::System.Int32)((((((colorValue1) / (256L)) / (256L)))) % (256L));
                                            reds1[(i) - (1), (j) - (1)] = red1;
                                            greens1[(i) - (1), (j) - (1)] = green1;
                                            blues1[(i) - (1), (j) - (1)] = blue1;
                                        }

                                        if ((cell.Font.Color is System.DBNull))
                                        {
                                            reds2[(i) - (1), (j) - (1)] = 0;
                                            greens2[(i) - (1), (j) - (1)] = 0;
                                            blues2[(i) - (1), (j) - (1)] = 0;
                                        }
                                        else
                                        {
                                            global::System.Int64 colorValue2 = Conversions.ToLong(cell.Font.Color);
                                            global::System.Int32 red2 = (global::System.Int32)((colorValue2) % (256L));
                                            global::System.Int32 green2 = (global::System.Int32)(((((colorValue2) / (256L)))) % (256L));
                                            global::System.Int32 blue2 = (global::System.Int32)((((((colorValue2) / (256L)) / (256L)))) % (256L));
                                            reds2[(i) - (1), (j) - (1)] = red2;
                                            greens2[(i) - (1), (j) - (1)] = green2;
                                            blues2[(i) - (1), (j) - (1)] = blue2;
                                        }
                                    }

                                }
                            }

                            this.rng.ClearContents();
                            this.rng.ClearFormats();

                            this.rng2.ClearFormats();

                            if (X5)
                            {

                                global::System.Int32 iColumn;
                                iColumn = 0;
                                var loopTo228 = r2;
                                for (i = 1; i <= loopTo228; i++)
                                {
                                    var loopTo229 = c2;
                                    for (j = 1; j <= loopTo229; j++)
                                    {
                                        global::System.Int32 x;
                                        global::System.Int32 y;
                                        x = 1;
                                        y = ((iColumn) + (j));
                                        if (((y) < (BreakPoints[(i) - (1)])))
                                        {
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Value = Arr[(x) - (1), (y) - (1)];

                                            if (((this.CheckBox1.Checked) == (true)))
                                            {

                                                global::Microsoft.Office.Interop.Excel.Range cell2 = (global::Microsoft.Office.Interop.Excel.Range)this.rng2.Cells[(global::System.Object)i, (global::System.Object)j];
                                                var font2 = cell2.Font;

                                                global::System.Single fontSize = fontSizes[(x) - (1), (y) - (1)];

                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Font.Name = fontNames[(x) - (1), (y) - (1)];
                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Font.Size = (global::System.Object)fontSizes[(x) - (1), (y) - (1)];

                                                if (Bolds[(x) - (1), (y) - (1)])
                                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Font.Bold = (global::System.Object)true;
                                                if (Italics[(x) - (1), (y) - (1)])
                                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Font.Italic = (global::System.Object)true;

                                                global::System.Int32 red1 = reds1[(x) - (1), (y) - (1)];
                                                global::System.Int32 green1 = greens1[(x) - (1), (y) - (1)];
                                                global::System.Int32 blue1 = blues1[(x) - (1), (y) - (1)];
                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Interior.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(red1, green1, blue1);

                                                global::System.Int32 red2 = reds2[(x) - (1), (y) - (1)];
                                                global::System.Int32 green2 = greens2[(x) - (1), (y) - (1)];
                                                global::System.Int32 blue2 = blues2[(x) - (1), (y) - (1)];
                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Font.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(red2, green2, blue2);
                                            }
                                        }
                                    }
                                    iColumn = BreakPoints[(i) - (1)];
                                }
                            }

                            else if (X6)
                            {
                                global::System.Int32 iColumn;
                                iColumn = 0;
                                var loopTo230 = c2;
                                for (j = 1; j <= loopTo230; j++)
                                {
                                    var loopTo231 = r2;
                                    for (i = 1; i <= loopTo231; i++)
                                    {
                                        global::System.Int32 x;
                                        global::System.Int32 y;
                                        x = 1;
                                        y = ((iColumn) + (i));
                                        if (((y) < (BreakPoints[(j) - (1)])))
                                        {
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Value = Arr[(x) - (1), (y) - (1)];

                                            if (((this.CheckBox1.Checked) == (true)))
                                            {

                                                global::Microsoft.Office.Interop.Excel.Range cell2 = (global::Microsoft.Office.Interop.Excel.Range)this.rng2.Cells[(global::System.Object)i, (global::System.Object)j];
                                                var font2 = cell2.Font;

                                                global::System.Single fontSize = fontSizes[(x) - (1), (y) - (1)];

                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Font.Name = fontNames[(x) - (1), (y) - (1)];
                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Font.Size = (global::System.Object)fontSizes[(x) - (1), (y) - (1)];

                                                if (Bolds[(x) - (1), (y) - (1)])
                                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Font.Bold = (global::System.Object)true;
                                                if (Italics[(x) - (1), (y) - (1)])
                                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Font.Italic = (global::System.Object)true;

                                                global::System.Int32 red1 = reds1[(x) - (1), (y) - (1)];
                                                global::System.Int32 green1 = greens1[(x) - (1), (y) - (1)];
                                                global::System.Int32 blue1 = blues1[(x) - (1), (y) - (1)];
                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Interior.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(red1, green1, blue1);

                                                global::System.Int32 red2 = reds2[(x) - (1), (y) - (1)];
                                                global::System.Int32 green2 = greens2[(x) - (1), (y) - (1)];
                                                global::System.Int32 blue2 = blues2[(x) - (1), (y) - (1)];
                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Font.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(red2, green2, blue2);

                                            }
                                        }
                                    }
                                    iColumn = BreakPoints[(j) - (1)];
                                }
                            }

                            if (((this.CheckBox1.Checked) == (true)))
                            {

                                if (((TopBorder8) == (true)))
                                {
                                    var loopTo232 = this.rng2.Columns.Count;
                                    for (j = 1; j <= loopTo232; j++)
                                    {
                                        this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)8).LineStyle = TopBorder8L;
                                        this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)8).Color = TopBorder8C;
                                        this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)8).Weight = TopBorder8W;
                                    }
                                }
                                else
                                {
                                    var loopTo233 = this.rng2.Columns.Count;
                                    for (j = 1; j <= loopTo233; j++)
                                        this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)8).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                }

                                if (((this.rng2.Rows.Count) > (1)))
                                {
                                    if (((TopBorder9) == (true)))
                                    {
                                        var loopTo234 = this.rng2.Columns.Count;
                                        for (j = 1; j <= loopTo234; j++)
                                        {
                                            this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)9).LineStyle = TopBorder9L;
                                            this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)9).Color = TopBorder9C;
                                            this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)9).Weight = TopBorder9W;
                                        }
                                    }
                                    else
                                    {
                                        var loopTo235 = this.rng2.Columns.Count;
                                        for (j = 1; j <= loopTo235; j++)
                                            this.rng2.Cells[(global::System.Object)this.rng2.Rows.Count, (global::System.Object)j].Borders((global::System.Object)9).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                    }
                                }

                                if (((TopBorder7) == (true)))
                                {
                                    var loopTo236 = this.rng2.Rows.Count;
                                    for (i = 1; i <= loopTo236; i++)
                                    {
                                        this.rng2.Cells[(global::System.Object)i, (global::System.Object)1].Borders((global::System.Object)7).LineStyle = TopBorder7L;
                                        this.rng2.Cells[(global::System.Object)i, (global::System.Object)1].Borders((global::System.Object)7).Color = TopBorder7C;
                                        this.rng2.Cells[(global::System.Object)i, (global::System.Object)1].Borders((global::System.Object)7).Weight = TopBorder7W;
                                    }
                                }
                                else
                                {
                                    var loopTo237 = this.rng2.Rows.Count;
                                    for (i = 1; i <= loopTo237; i++)
                                        this.rng2.Cells[(global::System.Object)i, (global::System.Object)1].Borders((global::System.Object)7).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                }

                                if (((BottomBorder9) == (true)))
                                {
                                    var loopTo238 = this.rng2.Columns.Count;
                                    for (j = 1; j <= loopTo238; j++)
                                    {
                                        this.rng2.Cells[(global::System.Object)this.rng2.Rows.Count, (global::System.Object)j].Borders((global::System.Object)9).LineStyle = BottomBorder9L;
                                        this.rng2.Cells[(global::System.Object)this.rng2.Rows.Count, (global::System.Object)j].Borders((global::System.Object)9).Color = BottomBorder9C;
                                        this.rng2.Cells[(global::System.Object)this.rng2.Rows.Count, (global::System.Object)j].Borders((global::System.Object)9).Weight = BottomBorder9W;
                                    }
                                }
                                else
                                {
                                    var loopTo239 = this.rng2.Columns.Count;
                                    for (j = 1; j <= loopTo239; j++)
                                        this.rng2.Cells[(global::System.Object)this.rng2.Rows.Count, (global::System.Object)this.rng2.Columns.Count].Borders((global::System.Object)9).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                }

                                if (((BottomBorder10) == (true)))
                                {
                                    var loopTo240 = this.rng2.Rows.Count;
                                    for (i = 1; i <= loopTo240; i++)
                                    {
                                        this.rng2.Cells[(global::System.Object)i, (global::System.Object)this.rng2.Columns.Count].Borders((global::System.Object)10).LineStyle = BottomBorder10L;
                                        this.rng2.Cells[(global::System.Object)i, (global::System.Object)this.rng2.Columns.Count].Borders((global::System.Object)10).Color = BottomBorder10C;
                                        this.rng2.Cells[(global::System.Object)i, (global::System.Object)this.rng2.Columns.Count].Borders((global::System.Object)10).Weight = BottomBorder10W;
                                    }
                                }
                                else
                                {
                                    var loopTo241 = this.rng2.Rows.Count;
                                    for (i = 1; i <= loopTo241; i++)
                                        this.rng2.Cells[(global::System.Object)i, (global::System.Object)this.rng2.Columns.Count].Borders((global::System.Object)10).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                }

                                if (((this.rng.Rows.Count) > (1)))
                                {

                                    if (((MiddleBorder9) == (true)))
                                    {
                                        var loopTo242 = (this.rng2.Rows.Count) - (1);
                                        for (i = 2; i <= loopTo242; i++)
                                        {
                                            var loopTo243 = this.rng2.Columns.Count;
                                            for (j = 1; j <= loopTo243; j++)
                                            {
                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)9).LineStyle = MiddleBorder9L;
                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)9).Color = MiddleBorder9C;
                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)9).Weight = MiddleBorder9W;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        var loopTo244 = (this.rng2.Rows.Count) - (1);
                                        for (i = 2; i <= loopTo244; i++)
                                        {
                                            var loopTo245 = this.rng2.Columns.Count;
                                            for (j = 1; j <= loopTo245; j++)
                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)9).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                        }
                                    }

                                }

                                if (((this.rng.Columns.Count) > (1)))
                                {

                                    if (((MiddleBorder10) == (true)))
                                    {
                                        var loopTo246 = (this.rng2.Columns.Count) - (1);
                                        for (j = 1; j <= loopTo246; j++)
                                        {
                                            var loopTo247 = this.rng2.Rows.Count;
                                            for (i = 1; i <= loopTo247; i++)
                                            {
                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)10).LineStyle = MiddleBorder10L;
                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)10).Color = MiddleBorder10C;
                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)10).Weight = MiddleBorder10W;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        var loopTo248 = (this.rng2.Columns.Count) - (1);
                                        for (j = 1; j <= loopTo248; j++)
                                        {
                                            var loopTo249 = this.rng2.Rows.Count;
                                            for (i = 1; i <= loopTo249; i++)
                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)10).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                        }
                                    }

                                }

                            }

                        }
                    }

                    else if (((((((X8) & !string.IsNullOrEmpty(this.TextBox2.Text)) & ((this.CanConvertToInt(this.TextBox2.Text)) == (true))))) & ((((X5) | (X6))))))
                    {

                        if (X5)
                        {
                            global::System.Int32 r2;
                            global::System.Int32 c2;
                            if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(Operators.ModObject(c, global::Microsoft.VisualBasic.Conversion.Int(this.TextBox2.Text)), 0, false)))
                            {
                                r2 = Conversions.ToInteger(global::Microsoft.VisualBasic.Conversion.Int(Operators.DivideObject(c, global::Microsoft.VisualBasic.Conversion.Int(this.TextBox2.Text))));
                            }
                            else
                            {
                                r2 = Conversions.ToInteger(Operators.AddObject(global::Microsoft.VisualBasic.Conversion.Int(Operators.DivideObject(c, global::Microsoft.VisualBasic.Conversion.Int(this.TextBox2.Text))), 1));
                            }
                            c2 = Conversions.ToInteger(global::Microsoft.VisualBasic.Conversion.Int(this.TextBox2.Text));

                            this.rng2 = this.worksheet2.get_Range(this.rng2.Cells[(global::System.Object)1, (global::System.Object)1], this.rng2.Cells[(global::System.Object)r2, (global::System.Object)c2]);
                            global::System.String rng2Address = this.rng2.get_Address();
                            this.worksheet2.Activate();
                            this.rng2.Select();

                            if (((this.Overlap(this.excelApp, this.worksheet, this.worksheet2, this.rng, this.rng2)) == (false)))
                            {

                                this.rng2.ClearFormats();

                                var loopTo250 = r2;
                                for (i = 1; i <= loopTo250; i++)
                                {
                                    var loopTo251 = c2;
                                    for (j = 1; j <= loopTo251; j++)
                                    {
                                        global::System.Int32 x;
                                        global::System.Int32 y;
                                        x = 1;
                                        y = (((((c2) * ((((i) - (1))))))) + (j));

                                        if (((y) <= (this.rng.Columns.Count)))
                                        {
                                            if (((this.CheckBox1.Checked) == (false)))
                                            {
                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Value = this.rng.Cells[(global::System.Object)x, (global::System.Object)y].Value;
                                            }

                                            else if (((this.CheckBox1.Checked) == (true)))
                                            {

                                                this.rng.Cells[(global::System.Object)x, (global::System.Object)y].Copy();
                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].PasteSpecial(global::Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues);
                                                this.rng2 = this.worksheet2.get_Range(rng2Address);
                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].PasteSpecial(global::Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats);
                                                this.rng2 = this.worksheet2.get_Range(rng2Address);
                                            }

                                        }
                                    }
                                }
                                excelApp.CutCopyMode = global::Microsoft.Office.Interop.Excel.XlCutCopyMode.xlCopy;

                                if (((this.CheckBox1.Checked) == (true)))
                                {
                                    global::Microsoft.Office.Interop.Excel.Range TopCell = (global::Microsoft.Office.Interop.Excel.Range)this.rng.Cells[(global::System.Object)1, (global::System.Object)1];

                                    if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)8].LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                                    {
                                        var loopTo252 = this.rng2.Columns.Count;
                                        for (j = 1; j <= loopTo252; j++)
                                        {
                                            this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)8).LineStyle = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)8].LineStyle;
                                            this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)8).Color = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)8].Color;
                                            this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)8).Weight = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)8].Weight;
                                        }
                                    }
                                    else
                                    {
                                        var loopTo253 = this.rng2.Columns.Count;
                                        for (j = 1; j <= loopTo253; j++)
                                            this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)8).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                    }

                                    if (((this.rng2.Rows.Count) > (1)))
                                    {
                                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                                        {
                                            var loopTo254 = this.rng2.Columns.Count;
                                            for (j = 1; j <= loopTo254; j++)
                                            {
                                                this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)9).LineStyle = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].LineStyle;
                                                this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)9).Color = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].Color;
                                                this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)9).Weight = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].Weight;
                                            }
                                        }
                                        else
                                        {
                                            var loopTo255 = this.rng2.Columns.Count;
                                            for (j = 1; j <= loopTo255; j++)
                                                this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)9).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                        }
                                    }

                                    if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)7].LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                                    {
                                        var loopTo256 = this.rng2.Rows.Count;
                                        for (i = 1; i <= loopTo256; i++)
                                        {
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)1].Borders((global::System.Object)7).LineStyle = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)7].LineStyle;
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)1].Borders((global::System.Object)7).Color = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)7].Color;
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)1].Borders((global::System.Object)7).Weight = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)7].Weight;
                                        }
                                    }
                                    else
                                    {
                                        var loopTo257 = this.rng2.Rows.Count;
                                        for (i = 1; i <= loopTo257; i++)
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)1].Borders((global::System.Object)7).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                    }

                                    global::Microsoft.Office.Interop.Excel.Range BottomCell = (global::Microsoft.Office.Interop.Excel.Range)this.rng.Cells[(global::System.Object)r, (global::System.Object)c];

                                    if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(BottomCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                                    {
                                        var loopTo258 = this.rng2.Columns.Count;
                                        for (j = 1; j <= loopTo258; j++)
                                        {
                                            this.rng2.Cells[(global::System.Object)this.rng2.Rows.Count, (global::System.Object)j].Borders((global::System.Object)9).LineStyle = BottomCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].LineStyle;
                                            this.rng2.Cells[(global::System.Object)this.rng2.Rows.Count, (global::System.Object)j].Borders((global::System.Object)9).Color = BottomCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].Color;
                                            this.rng2.Cells[(global::System.Object)this.rng2.Rows.Count, (global::System.Object)j].Borders((global::System.Object)9).Weight = BottomCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].Weight;
                                        }
                                    }
                                    else
                                    {
                                        var loopTo259 = this.rng2.Columns.Count;
                                        for (j = 1; j <= loopTo259; j++)
                                            this.rng2.Cells[(global::System.Object)this.rng2.Rows.Count, (global::System.Object)this.rng2.Columns.Count].Borders((global::System.Object)9).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                    }

                                    if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(BottomCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)10].LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                                    {
                                        var loopTo260 = this.rng2.Rows.Count;
                                        for (i = 1; i <= loopTo260; i++)
                                        {
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)this.rng2.Columns.Count].Borders((global::System.Object)10).LineStyle = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)10].LineStyle;
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)this.rng2.Columns.Count].Borders((global::System.Object)10).Color = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)10].Color;
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)this.rng2.Columns.Count].Borders((global::System.Object)10).Weight = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)10].Weight;
                                        }
                                    }
                                    else
                                    {
                                        var loopTo261 = this.rng2.Rows.Count;
                                        for (i = 1; i <= loopTo261; i++)
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)this.rng2.Columns.Count].Borders((global::System.Object)10).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                    }

                                    if (((this.rng.Rows.Count) > (1)))
                                    {
                                        global::Microsoft.Office.Interop.Excel.Range MiddleCell = (global::Microsoft.Office.Interop.Excel.Range)this.rng.Cells[(global::System.Object)2, (global::System.Object)1];
                                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(MiddleCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                                        {
                                            var loopTo262 = (this.rng2.Rows.Count) - (1);
                                            for (i = 2; i <= loopTo262; i++)
                                            {
                                                var loopTo263 = this.rng2.Columns.Count;
                                                for (j = 1; j <= loopTo263; j++)
                                                {
                                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)9).LineStyle = MiddleCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].LineStyle;
                                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)9).Color = MiddleCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].Color;
                                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)9).Weight = MiddleCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].Weight;
                                                }
                                            }
                                        }
                                        else
                                        {
                                            var loopTo264 = (this.rng2.Rows.Count) - (1);
                                            for (i = 2; i <= loopTo264; i++)
                                            {
                                                var loopTo265 = this.rng2.Columns.Count;
                                                for (j = 1; j <= loopTo265; j++)
                                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)9).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                            }
                                        }
                                    }

                                    if (((this.rng.Columns.Count) > (1)))
                                    {
                                        global::Microsoft.Office.Interop.Excel.Range MiddleCell = (global::Microsoft.Office.Interop.Excel.Range)this.rng.Cells[(global::System.Object)1, (global::System.Object)2];
                                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(MiddleCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)10].LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                                        {
                                            var loopTo266 = (this.rng2.Columns.Count) - (1);
                                            for (j = 1; j <= loopTo266; j++)
                                            {
                                                var loopTo267 = this.rng2.Rows.Count;
                                                for (i = 1; i <= loopTo267; i++)
                                                {
                                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)10).LineStyle = MiddleCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)10].LineStyle;
                                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)10).Color = MiddleCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)10].Color;
                                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)10).Weight = MiddleCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)10].Weight;
                                                }
                                            }
                                        }
                                        else
                                        {
                                            var loopTo268 = (this.rng2.Columns.Count) - (1);
                                            for (j = 1; j <= loopTo268; j++)
                                            {
                                                var loopTo269 = this.rng2.Rows.Count;
                                                for (i = 1; i <= loopTo269; i++)
                                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)10).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                            }
                                        }
                                    }
                                }
                            }

                            else
                            {

                                var Arr = new global::System.Object[(r), (c)];
                                var Bolds = new global::System.Boolean[(r), (c)];
                                var Italics = new global::System.Boolean[(r), (c)];
                                var fontNames = new global::System.String[(r), (c)];
                                var fontSizes = new global::System.Single[(r), (c)];
                                var reds1 = new global::System.Int32[(r), (c)];
                                var reds2 = new global::System.Int32[(r), (c)];
                                var greens1 = new global::System.Int32[(r), (c)];
                                var greens2 = new global::System.Int32[(r), (c)];
                                var blues1 = new global::System.Int32[(r), (c)];
                                var blues2 = new global::System.Int32[(r), (c)];

                                global::System.Boolean TopBorder7;
                                global::System.Object TopBorder7L;
                                global::System.Object TopBorder7C;
                                global::System.Object TopBorder7W;

                                global::System.Boolean TopBorder8;
                                global::System.Object TopBorder8L;
                                global::System.Object TopBorder8C;
                                global::System.Object TopBorder8W;

                                global::System.Boolean TopBorder9;
                                global::System.Object TopBorder9L;
                                global::System.Object TopBorder9C;
                                global::System.Object TopBorder9W;

                                global::System.Boolean BottomBorder9;
                                global::System.Object BottomBorder9L;
                                global::System.Object BottomBorder9C;
                                global::System.Object BottomBorder9W;

                                global::System.Boolean BottomBorder10;
                                global::System.Object BottomBorder10L;
                                global::System.Object BottomBorder10C;
                                global::System.Object BottomBorder10W;

                                global::System.Boolean MiddleBorder9;
                                global::System.Object MiddleBorder9L;
                                global::System.Object MiddleBorder9C;
                                global::System.Object MiddleBorder9W;

                                global::System.Boolean MiddleBorder10;
                                global::System.Object MiddleBorder10L;
                                global::System.Object MiddleBorder10C;
                                global::System.Object MiddleBorder10W;

                                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)7).LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                                {
                                    TopBorder7 = true;
                                    TopBorder7L = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)7).LineStyle;
                                    TopBorder7C = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)7).Color;
                                    TopBorder7W = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)7).Weight;
                                }
                                else
                                {
                                    TopBorder7 = false;
                                    TopBorder7L = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)7).LineStyle;
                                    TopBorder7C = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)7).Color;
                                    TopBorder7W = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)7).Weight;
                                }

                                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)8).LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                                {
                                    TopBorder8 = true;
                                    TopBorder8L = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)8).LineStyle;
                                    TopBorder8C = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)8).Color;
                                    TopBorder8W = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)8).Weight;
                                }
                                else
                                {
                                    TopBorder8 = false;
                                    TopBorder8L = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)8).LineStyle;
                                    TopBorder8C = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)8).Color;
                                    TopBorder8W = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)8).Weight;
                                }

                                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)9).LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                                {
                                    TopBorder9 = true;
                                    TopBorder9L = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)9).LineStyle;
                                    TopBorder9C = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)9).Color;
                                    TopBorder9W = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)9).Weight;
                                }
                                else
                                {
                                    TopBorder9 = false;
                                    TopBorder9L = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)9).LineStyle;
                                    TopBorder9C = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)9).Color;
                                    TopBorder9W = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)9).Weight;
                                }

                                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)9).LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                                {
                                    BottomBorder9 = true;
                                    BottomBorder9L = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)9).LineStyle;
                                    BottomBorder9C = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)9).Color;
                                    BottomBorder9W = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)9).Weight;
                                }
                                else
                                {
                                    BottomBorder9 = false;
                                    BottomBorder9L = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)9).LineStyle;
                                    BottomBorder9C = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)9).Color;
                                    BottomBorder9W = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)9).Weight;
                                }

                                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)10).LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                                {
                                    BottomBorder10 = true;
                                    BottomBorder10L = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)10).LineStyle;
                                    BottomBorder10C = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)10).Color;
                                    BottomBorder10W = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)10).Weight;
                                }
                                else
                                {
                                    BottomBorder10 = false;
                                    BottomBorder10L = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)10).LineStyle;
                                    BottomBorder10C = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)10).Color;
                                    BottomBorder10W = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)10).Weight;
                                }

                                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(this.rng.Cells[(global::System.Object)2, (global::System.Object)1].Borders((global::System.Object)9).LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                                {
                                    MiddleBorder9 = true;
                                    MiddleBorder9L = this.rng.Cells[(global::System.Object)2, (global::System.Object)1].Borders((global::System.Object)9).LineStyle;
                                    MiddleBorder9C = this.rng.Cells[(global::System.Object)2, (global::System.Object)1].Borders((global::System.Object)9).Color;
                                    MiddleBorder9W = this.rng.Cells[(global::System.Object)2, (global::System.Object)1].Borders((global::System.Object)9).Weight;
                                }
                                else
                                {
                                    MiddleBorder9 = false;
                                    MiddleBorder9L = this.rng.Cells[(global::System.Object)2, (global::System.Object)1].Borders((global::System.Object)9).LineStyle;
                                    MiddleBorder9C = this.rng.Cells[(global::System.Object)2, (global::System.Object)1].Borders((global::System.Object)9).Color;
                                    MiddleBorder9W = this.rng.Cells[(global::System.Object)2, (global::System.Object)1].Borders((global::System.Object)9).Weight;
                                }

                                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(this.rng.Cells[(global::System.Object)1, (global::System.Object)2].Borders((global::System.Object)10).LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                                {
                                    MiddleBorder10 = true;
                                    MiddleBorder10L = this.rng.Cells[(global::System.Object)1, (global::System.Object)2].Borders((global::System.Object)10).LineStyle;
                                    MiddleBorder10C = this.rng.Cells[(global::System.Object)1, (global::System.Object)2].Borders((global::System.Object)10).Color;
                                    MiddleBorder10W = this.rng.Cells[(global::System.Object)1, (global::System.Object)2].Borders((global::System.Object)10).Weight;
                                }
                                else
                                {
                                    MiddleBorder10 = false;
                                    MiddleBorder10L = this.rng.Cells[(global::System.Object)1, (global::System.Object)2].Borders((global::System.Object)10).LineStyle;
                                    MiddleBorder10C = this.rng.Cells[(global::System.Object)1, (global::System.Object)2].Borders((global::System.Object)10).Color;
                                    MiddleBorder10W = this.rng.Cells[(global::System.Object)1, (global::System.Object)2].Borders((global::System.Object)10).Weight;
                                }

                                var loopTo270 = r;
                                for (i = 1; i <= loopTo270; i++)
                                {
                                    var loopTo271 = c;
                                    for (j = 1; j <= loopTo271; j++)
                                    {
                                        Arr[(i) - (1), (j) - (1)] = this.rng.Cells[(global::System.Object)i, (global::System.Object)j].Value;

                                        if (((this.CheckBox1.Checked) == (true)))
                                        {

                                            global::Microsoft.Office.Interop.Excel.Range cell = (global::Microsoft.Office.Interop.Excel.Range)this.rng.Cells[(global::System.Object)i, (global::System.Object)j];
                                            var font = cell.Font;

                                            Bolds[(i) - (1), (j) - (1)] = Conversions.ToBoolean(cell.Font.Bold);
                                            Italics[(i) - (1), (j) - (1)] = Conversions.ToBoolean(cell.Font.Italic);


                                            if ((((font.Name is System.DBNull)) == (false)))
                                            {
                                                fontNames[(i) - (1), (j) - (1)] = Conversions.ToString(font.Name);
                                            }
                                            else
                                            {
                                                fontNames[(i) - (1), (j) - (1)] = "Calibri";
                                            }

                                            if ((((font.Size is System.DBNull)) == (false)))
                                            {
                                                global::System.Single fontSize = global::System.Convert.ToSingle(font.Size);
                                                fontSizes[(i) - (1), (j) - (1)] = fontSize;
                                            }
                                            else
                                            {
                                                fontSizes[(i) - (1), (j) - (1)] = 11f;
                                            }

                                            if ((cell.Interior.Color is System.DBNull))
                                            {
                                                reds1[(i) - (1), (j) - (1)] = 255;
                                                greens1[(i) - (1), (j) - (1)] = 255;
                                                blues1[(i) - (1), (j) - (1)] = 255;
                                            }
                                            else
                                            {
                                                global::System.Int64 colorValue1 = Conversions.ToLong(cell.Interior.Color);
                                                global::System.Int32 red1 = (global::System.Int32)((colorValue1) % (256L));
                                                global::System.Int32 green1 = (global::System.Int32)(((((colorValue1) / (256L)))) % (256L));
                                                global::System.Int32 blue1 = (global::System.Int32)((((((colorValue1) / (256L)) / (256L)))) % (256L));
                                                reds1[(i) - (1), (j) - (1)] = red1;
                                                greens1[(i) - (1), (j) - (1)] = green1;
                                                blues1[(i) - (1), (j) - (1)] = blue1;
                                            }

                                            if ((cell.Font.Color is System.DBNull))
                                            {
                                                reds2[(i) - (1), (j) - (1)] = 0;
                                                greens2[(i) - (1), (j) - (1)] = 0;
                                                blues2[(i) - (1), (j) - (1)] = 0;
                                            }
                                            else
                                            {
                                                global::System.Int64 colorValue2 = Conversions.ToLong(cell.Font.Color);
                                                global::System.Int32 red2 = (global::System.Int32)((colorValue2) % (256L));
                                                global::System.Int32 green2 = (global::System.Int32)(((((colorValue2) / (256L)))) % (256L));
                                                global::System.Int32 blue2 = (global::System.Int32)((((((colorValue2) / (256L)) / (256L)))) % (256L));
                                                reds2[(i) - (1), (j) - (1)] = red2;
                                                greens2[(i) - (1), (j) - (1)] = green2;
                                                blues2[(i) - (1), (j) - (1)] = blue2;
                                            }
                                        }

                                    }
                                }

                                this.rng.ClearContents();
                                this.rng.ClearFormats();

                                this.rng2.ClearFormats();

                                var loopTo272 = r2;
                                for (i = 1; i <= loopTo272; i++)
                                {
                                    var loopTo273 = c2;
                                    for (j = 1; j <= loopTo273; j++)
                                    {
                                        global::System.Int32 x;
                                        global::System.Int32 y;
                                        x = 1;
                                        y = (((((c2) * ((((i) - (1))))))) + (j));
                                        if (((y) <= ((global::Microsoft.VisualBasic.Information.UBound(Arr, 2)) + (1))))
                                        {
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Value = Arr[(x) - (1), (y) - (1)];

                                            if (((this.CheckBox1.Checked) == (true)))
                                            {

                                                global::Microsoft.Office.Interop.Excel.Range cell2 = (global::Microsoft.Office.Interop.Excel.Range)this.rng2.Cells[(global::System.Object)i, (global::System.Object)j];
                                                var font2 = cell2.Font;

                                                global::System.Single fontSize = fontSizes[(x) - (1), (y) - (1)];

                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Font.Name = fontNames[(x) - (1), (y) - (1)];
                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Font.Size = (global::System.Object)fontSizes[(x) - (1), (y) - (1)];

                                                if (Bolds[(x) - (1), (y) - (1)])
                                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Font.Bold = (global::System.Object)true;
                                                if (Italics[(x) - (1), (y) - (1)])
                                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Font.Italic = (global::System.Object)true;

                                                global::System.Int32 red1 = reds1[(x) - (1), (y) - (1)];
                                                global::System.Int32 green1 = greens1[(x) - (1), (y) - (1)];
                                                global::System.Int32 blue1 = blues1[(x) - (1), (y) - (1)];
                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Interior.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(red1, green1, blue1);

                                                global::System.Int32 red2 = reds2[(x) - (1), (y) - (1)];
                                                global::System.Int32 green2 = greens2[(x) - (1), (y) - (1)];
                                                global::System.Int32 blue2 = blues2[(x) - (1), (y) - (1)];
                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Font.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(red2, green2, blue2);
                                            }
                                        }
                                    }
                                }

                                if (((this.CheckBox1.Checked) == (true)))
                                {

                                    if (((TopBorder8) == (true)))
                                    {
                                        var loopTo274 = this.rng2.Columns.Count;
                                        for (j = 1; j <= loopTo274; j++)
                                        {
                                            this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)8).LineStyle = TopBorder8L;
                                            this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)8).Color = TopBorder8C;
                                            this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)8).Weight = TopBorder8W;
                                        }
                                    }
                                    else
                                    {
                                        var loopTo275 = this.rng2.Columns.Count;
                                        for (j = 1; j <= loopTo275; j++)
                                            this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)8).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                    }

                                    if (((this.rng2.Rows.Count) > (1)))
                                    {
                                        if (((TopBorder9) == (true)))
                                        {
                                            var loopTo276 = this.rng2.Columns.Count;
                                            for (j = 1; j <= loopTo276; j++)
                                            {
                                                this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)9).LineStyle = TopBorder9L;
                                                this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)9).Color = TopBorder9C;
                                                this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)9).Weight = TopBorder9W;
                                            }
                                        }
                                        else
                                        {
                                            var loopTo277 = this.rng2.Columns.Count;
                                            for (j = 1; j <= loopTo277; j++)
                                                this.rng2.Cells[(global::System.Object)this.rng2.Rows.Count, (global::System.Object)j].Borders((global::System.Object)9).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                        }
                                    }

                                    if (((TopBorder7) == (true)))
                                    {
                                        var loopTo278 = this.rng2.Rows.Count;
                                        for (i = 1; i <= loopTo278; i++)
                                        {
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)1].Borders((global::System.Object)7).LineStyle = TopBorder7L;
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)1].Borders((global::System.Object)7).Color = TopBorder7C;
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)1].Borders((global::System.Object)7).Weight = TopBorder7W;
                                        }
                                    }
                                    else
                                    {
                                        var loopTo279 = this.rng2.Rows.Count;
                                        for (i = 1; i <= loopTo279; i++)
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)1].Borders((global::System.Object)7).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                    }

                                    if (((BottomBorder9) == (true)))
                                    {
                                        var loopTo280 = this.rng2.Columns.Count;
                                        for (j = 1; j <= loopTo280; j++)
                                        {
                                            this.rng2.Cells[(global::System.Object)this.rng2.Rows.Count, (global::System.Object)j].Borders((global::System.Object)9).LineStyle = BottomBorder9L;
                                            this.rng2.Cells[(global::System.Object)this.rng2.Rows.Count, (global::System.Object)j].Borders((global::System.Object)9).Color = BottomBorder9C;
                                            this.rng2.Cells[(global::System.Object)this.rng2.Rows.Count, (global::System.Object)j].Borders((global::System.Object)9).Weight = BottomBorder9W;
                                        }
                                    }
                                    else
                                    {
                                        var loopTo281 = this.rng2.Columns.Count;
                                        for (j = 1; j <= loopTo281; j++)
                                            this.rng2.Cells[(global::System.Object)this.rng2.Rows.Count, (global::System.Object)this.rng2.Columns.Count].Borders((global::System.Object)9).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                    }

                                    if (((BottomBorder10) == (true)))
                                    {
                                        var loopTo282 = this.rng2.Rows.Count;
                                        for (i = 1; i <= loopTo282; i++)
                                        {
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)this.rng2.Columns.Count].Borders((global::System.Object)10).LineStyle = BottomBorder10L;
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)this.rng2.Columns.Count].Borders((global::System.Object)10).Color = BottomBorder10C;
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)this.rng2.Columns.Count].Borders((global::System.Object)10).Weight = BottomBorder10W;
                                        }
                                    }
                                    else
                                    {
                                        var loopTo283 = this.rng2.Rows.Count;
                                        for (i = 1; i <= loopTo283; i++)
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)this.rng2.Columns.Count].Borders((global::System.Object)10).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                    }

                                    if (((this.rng.Rows.Count) > (1)))
                                    {

                                        if (((MiddleBorder9) == (true)))
                                        {
                                            var loopTo284 = (this.rng2.Rows.Count) - (1);
                                            for (i = 2; i <= loopTo284; i++)
                                            {
                                                var loopTo285 = this.rng2.Columns.Count;
                                                for (j = 1; j <= loopTo285; j++)
                                                {
                                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)9).LineStyle = MiddleBorder9L;
                                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)9).Color = MiddleBorder9C;
                                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)9).Weight = MiddleBorder9W;
                                                }
                                            }
                                        }
                                        else
                                        {
                                            var loopTo286 = (this.rng2.Rows.Count) - (1);
                                            for (i = 2; i <= loopTo286; i++)
                                            {
                                                var loopTo287 = this.rng2.Columns.Count;
                                                for (j = 1; j <= loopTo287; j++)
                                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)9).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                            }
                                        }

                                    }

                                    if (((this.rng.Columns.Count) > (1)))
                                    {

                                        if (((MiddleBorder10) == (true)))
                                        {
                                            var loopTo288 = (this.rng2.Columns.Count) - (1);
                                            for (j = 1; j <= loopTo288; j++)
                                            {
                                                var loopTo289 = this.rng2.Rows.Count;
                                                for (i = 1; i <= loopTo289; i++)
                                                {
                                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)10).LineStyle = MiddleBorder10L;
                                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)10).Color = MiddleBorder10C;
                                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)10).Weight = MiddleBorder10W;
                                                }
                                            }
                                        }
                                        else
                                        {
                                            var loopTo290 = (this.rng2.Columns.Count) - (1);
                                            for (j = 1; j <= loopTo290; j++)
                                            {
                                                var loopTo291 = this.rng2.Rows.Count;
                                                for (i = 1; i <= loopTo291; i++)
                                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)10).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                            }
                                        }

                                    }

                                }

                            }
                        }

                        else if (X6)
                        {
                            global::System.Int32 r2;
                            global::System.Int32 c2;
                            if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(Operators.ModObject(c, global::Microsoft.VisualBasic.Conversion.Int(this.TextBox2.Text)), 0, false)))
                            {
                                c2 = Conversions.ToInteger(global::Microsoft.VisualBasic.Conversion.Int(Operators.DivideObject(c, global::Microsoft.VisualBasic.Conversion.Int(this.TextBox2.Text))));
                            }
                            else
                            {
                                c2 = Conversions.ToInteger(Operators.AddObject(global::Microsoft.VisualBasic.Conversion.Int(Operators.DivideObject(c, global::Microsoft.VisualBasic.Conversion.Int(this.TextBox2.Text))), 1));
                            }
                            r2 = Conversions.ToInteger(global::Microsoft.VisualBasic.Conversion.Int(this.TextBox2.Text));

                            this.rng2 = this.worksheet2.get_Range(this.rng2.Cells[(global::System.Object)1, (global::System.Object)1], this.rng2.Cells[(global::System.Object)r2, (global::System.Object)c2]);
                            global::System.String rng2Address = this.rng2.get_Address();
                            this.worksheet2.Activate();
                            this.rng2.Select();

                            if (((this.Overlap(this.excelApp, this.worksheet, this.worksheet2, this.rng, this.rng2)) == (false)))
                            {

                                this.rng2.ClearFormats();

                                var loopTo292 = c2;
                                for (j = 1; j <= loopTo292; j++)
                                {
                                    var loopTo293 = r2;
                                    for (i = 1; i <= loopTo293; i++)
                                    {
                                        global::System.Int32 x;
                                        global::System.Int32 y;
                                        x = 1;
                                        y = (((((r2) * ((((j) - (1))))))) + (i));
                                        if (((y) <= (this.rng.Columns.Count)))
                                        {
                                            if (((this.CheckBox1.Checked) == (false)))
                                            {
                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Value = this.rng.Cells[(global::System.Object)x, (global::System.Object)y].Value;
                                            }

                                            else if (((this.CheckBox1.Checked) == (true)))
                                            {

                                                this.rng.Cells[(global::System.Object)x, (global::System.Object)y].Copy();
                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].PasteSpecial(global::Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues);
                                                this.rng2 = this.worksheet2.get_Range(rng2Address);
                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].PasteSpecial(global::Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats);
                                                this.rng2 = this.worksheet2.get_Range(rng2Address);
                                            }
                                        }
                                    }
                                }
                                excelApp.CutCopyMode = global::Microsoft.Office.Interop.Excel.XlCutCopyMode.xlCopy;

                                if (((this.CheckBox1.Checked) == (true)))
                                {
                                    global::Microsoft.Office.Interop.Excel.Range TopCell = (global::Microsoft.Office.Interop.Excel.Range)this.rng.Cells[(global::System.Object)1, (global::System.Object)1];

                                    if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)8].LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                                    {
                                        var loopTo294 = this.rng2.Columns.Count;
                                        for (j = 1; j <= loopTo294; j++)
                                        {
                                            this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)8).LineStyle = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)8].LineStyle;
                                            this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)8).Color = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)8].Color;
                                            this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)8).Weight = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)8].Weight;
                                        }
                                    }
                                    else
                                    {
                                        var loopTo295 = this.rng2.Columns.Count;
                                        for (j = 1; j <= loopTo295; j++)
                                            this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)8).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                    }

                                    if (((this.rng2.Rows.Count) > (1)))
                                    {
                                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                                        {
                                            var loopTo296 = this.rng2.Columns.Count;
                                            for (j = 1; j <= loopTo296; j++)
                                            {
                                                this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)9).LineStyle = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].LineStyle;
                                                this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)9).Color = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].Color;
                                                this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)9).Weight = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].Weight;
                                            }
                                        }
                                        else
                                        {
                                            var loopTo297 = this.rng2.Columns.Count;
                                            for (j = 1; j <= loopTo297; j++)
                                                this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)9).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                        }
                                    }

                                    if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)7].LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                                    {
                                        var loopTo298 = this.rng2.Rows.Count;
                                        for (i = 1; i <= loopTo298; i++)
                                        {
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)1].Borders((global::System.Object)7).LineStyle = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)7].LineStyle;
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)1].Borders((global::System.Object)7).Color = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)7].Color;
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)1].Borders((global::System.Object)7).Weight = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)7].Weight;
                                        }
                                    }
                                    else
                                    {
                                        var loopTo299 = this.rng2.Rows.Count;
                                        for (i = 1; i <= loopTo299; i++)
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)1].Borders((global::System.Object)7).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                    }

                                    global::Microsoft.Office.Interop.Excel.Range BottomCell = (global::Microsoft.Office.Interop.Excel.Range)this.rng.Cells[(global::System.Object)r, (global::System.Object)c];

                                    if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(BottomCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                                    {
                                        var loopTo300 = this.rng2.Columns.Count;
                                        for (j = 1; j <= loopTo300; j++)
                                        {
                                            this.rng2.Cells[(global::System.Object)this.rng2.Rows.Count, (global::System.Object)j].Borders((global::System.Object)9).LineStyle = BottomCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].LineStyle;
                                            this.rng2.Cells[(global::System.Object)this.rng2.Rows.Count, (global::System.Object)j].Borders((global::System.Object)9).Color = BottomCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].Color;
                                            this.rng2.Cells[(global::System.Object)this.rng2.Rows.Count, (global::System.Object)j].Borders((global::System.Object)9).Weight = BottomCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].Weight;
                                        }
                                    }
                                    else
                                    {
                                        var loopTo301 = this.rng2.Columns.Count;
                                        for (j = 1; j <= loopTo301; j++)
                                            this.rng2.Cells[(global::System.Object)this.rng2.Rows.Count, (global::System.Object)this.rng2.Columns.Count].Borders((global::System.Object)9).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                    }

                                    if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(BottomCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)10].LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                                    {
                                        var loopTo302 = this.rng2.Rows.Count;
                                        for (i = 1; i <= loopTo302; i++)
                                        {
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)this.rng2.Columns.Count].Borders((global::System.Object)10).LineStyle = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)10].LineStyle;
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)this.rng2.Columns.Count].Borders((global::System.Object)10).Color = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)10].Color;
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)this.rng2.Columns.Count].Borders((global::System.Object)10).Weight = TopCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)10].Weight;
                                        }
                                    }
                                    else
                                    {
                                        var loopTo303 = this.rng2.Rows.Count;
                                        for (i = 1; i <= loopTo303; i++)
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)this.rng2.Columns.Count].Borders((global::System.Object)10).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                    }

                                    if (((this.rng.Rows.Count) > (1)))
                                    {
                                        global::Microsoft.Office.Interop.Excel.Range MiddleCell = (global::Microsoft.Office.Interop.Excel.Range)this.rng.Cells[(global::System.Object)2, (global::System.Object)1];
                                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(MiddleCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                                        {
                                            var loopTo304 = (this.rng2.Rows.Count) - (1);
                                            for (i = 2; i <= loopTo304; i++)
                                            {
                                                var loopTo305 = this.rng2.Columns.Count;
                                                for (j = 1; j <= loopTo305; j++)
                                                {
                                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)9).LineStyle = MiddleCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].LineStyle;
                                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)9).Color = MiddleCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].Color;
                                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)9).Weight = MiddleCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)9].Weight;
                                                }
                                            }
                                        }
                                        else
                                        {
                                            var loopTo306 = (this.rng2.Rows.Count) - (1);
                                            for (i = 2; i <= loopTo306; i++)
                                            {
                                                var loopTo307 = this.rng2.Columns.Count;
                                                for (j = 1; j <= loopTo307; j++)
                                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)9).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                            }
                                        }
                                    }

                                    if (((this.rng.Columns.Count) > (1)))
                                    {
                                        global::Microsoft.Office.Interop.Excel.Range MiddleCell = (global::Microsoft.Office.Interop.Excel.Range)this.rng.Cells[(global::System.Object)1, (global::System.Object)2];
                                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(MiddleCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)10].LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                                        {
                                            var loopTo308 = (this.rng2.Columns.Count) - (1);
                                            for (j = 1; j <= loopTo308; j++)
                                            {
                                                var loopTo309 = this.rng2.Rows.Count;
                                                for (i = 1; i <= loopTo309; i++)
                                                {
                                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)10).LineStyle = MiddleCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)10].LineStyle;
                                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)10).Color = MiddleCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)10].Color;
                                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)10).Weight = MiddleCell.Borders[(global::Microsoft.Office.Interop.Excel.XlBordersIndex)10].Weight;
                                                }
                                            }
                                        }
                                        else
                                        {
                                            var loopTo310 = (this.rng2.Columns.Count) - (1);
                                            for (j = 1; j <= loopTo310; j++)
                                            {
                                                var loopTo311 = this.rng2.Rows.Count;
                                                for (i = 1; i <= loopTo311; i++)
                                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)10).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                            }
                                        }
                                    }
                                }
                            }
                            else
                            {

                                var Arr = new global::System.Object[(r), (c)];
                                var Bolds = new global::System.Boolean[(r), (c)];
                                var Italics = new global::System.Boolean[(r), (c)];
                                var fontNames = new global::System.String[(r), (c)];
                                var fontSizes = new global::System.Single[(r), (c)];
                                var reds1 = new global::System.Int32[(r), (c)];
                                var reds2 = new global::System.Int32[(r), (c)];
                                var greens1 = new global::System.Int32[(r), (c)];
                                var greens2 = new global::System.Int32[(r), (c)];
                                var blues1 = new global::System.Int32[(r), (c)];
                                var blues2 = new global::System.Int32[(r), (c)];

                                global::System.Boolean TopBorder7;
                                global::System.Object TopBorder7L;
                                global::System.Object TopBorder7C;
                                global::System.Object TopBorder7W;

                                global::System.Boolean TopBorder8;
                                global::System.Object TopBorder8L;
                                global::System.Object TopBorder8C;
                                global::System.Object TopBorder8W;

                                global::System.Boolean TopBorder9;
                                global::System.Object TopBorder9L;
                                global::System.Object TopBorder9C;
                                global::System.Object TopBorder9W;

                                global::System.Boolean BottomBorder9;
                                global::System.Object BottomBorder9L;
                                global::System.Object BottomBorder9C;
                                global::System.Object BottomBorder9W;

                                global::System.Boolean BottomBorder10;
                                global::System.Object BottomBorder10L;
                                global::System.Object BottomBorder10C;
                                global::System.Object BottomBorder10W;

                                global::System.Boolean MiddleBorder9;
                                global::System.Object MiddleBorder9L;
                                global::System.Object MiddleBorder9C;
                                global::System.Object MiddleBorder9W;

                                global::System.Boolean MiddleBorder10;
                                global::System.Object MiddleBorder10L;
                                global::System.Object MiddleBorder10C;
                                global::System.Object MiddleBorder10W;

                                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)7).LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                                {
                                    TopBorder7 = true;
                                    TopBorder7L = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)7).LineStyle;
                                    TopBorder7C = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)7).Color;
                                    TopBorder7W = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)7).Weight;
                                }
                                else
                                {
                                    TopBorder7 = false;
                                    TopBorder7L = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)7).LineStyle;
                                    TopBorder7C = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)7).Color;
                                    TopBorder7W = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)7).Weight;
                                }

                                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)8).LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                                {
                                    TopBorder8 = true;
                                    TopBorder8L = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)8).LineStyle;
                                    TopBorder8C = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)8).Color;
                                    TopBorder8W = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)8).Weight;
                                }
                                else
                                {
                                    TopBorder8 = false;
                                    TopBorder8L = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)8).LineStyle;
                                    TopBorder8C = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)8).Color;
                                    TopBorder8W = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)8).Weight;
                                }

                                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)9).LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                                {
                                    TopBorder9 = true;
                                    TopBorder9L = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)9).LineStyle;
                                    TopBorder9C = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)9).Color;
                                    TopBorder9W = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)9).Weight;
                                }
                                else
                                {
                                    TopBorder9 = false;
                                    TopBorder9L = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)9).LineStyle;
                                    TopBorder9C = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)9).Color;
                                    TopBorder9W = this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Borders((global::System.Object)9).Weight;
                                }

                                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)9).LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                                {
                                    BottomBorder9 = true;
                                    BottomBorder9L = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)9).LineStyle;
                                    BottomBorder9C = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)9).Color;
                                    BottomBorder9W = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)9).Weight;
                                }
                                else
                                {
                                    BottomBorder9 = false;
                                    BottomBorder9L = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)9).LineStyle;
                                    BottomBorder9C = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)9).Color;
                                    BottomBorder9W = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)9).Weight;
                                }

                                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)10).LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                                {
                                    BottomBorder10 = true;
                                    BottomBorder10L = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)10).LineStyle;
                                    BottomBorder10C = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)10).Color;
                                    BottomBorder10W = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)10).Weight;
                                }
                                else
                                {
                                    BottomBorder10 = false;
                                    BottomBorder10L = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)10).LineStyle;
                                    BottomBorder10C = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)10).Color;
                                    BottomBorder10W = this.rng.Cells[(global::System.Object)r, (global::System.Object)c].Borders((global::System.Object)10).Weight;
                                }

                                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(this.rng.Cells[(global::System.Object)2, (global::System.Object)1].Borders((global::System.Object)9).LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                                {
                                    MiddleBorder9 = true;
                                    MiddleBorder9L = this.rng.Cells[(global::System.Object)2, (global::System.Object)1].Borders((global::System.Object)9).LineStyle;
                                    MiddleBorder9C = this.rng.Cells[(global::System.Object)2, (global::System.Object)1].Borders((global::System.Object)9).Color;
                                    MiddleBorder9W = this.rng.Cells[(global::System.Object)2, (global::System.Object)1].Borders((global::System.Object)9).Weight;
                                }
                                else
                                {
                                    MiddleBorder9 = false;
                                    MiddleBorder9L = this.rng.Cells[(global::System.Object)2, (global::System.Object)1].Borders((global::System.Object)9).LineStyle;
                                    MiddleBorder9C = this.rng.Cells[(global::System.Object)2, (global::System.Object)1].Borders((global::System.Object)9).Color;
                                    MiddleBorder9W = this.rng.Cells[(global::System.Object)2, (global::System.Object)1].Borders((global::System.Object)9).Weight;
                                }

                                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(this.rng.Cells[(global::System.Object)1, (global::System.Object)2].Borders((global::System.Object)10).LineStyle, global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone, false)))
                                {
                                    MiddleBorder10 = true;
                                    MiddleBorder10L = this.rng.Cells[(global::System.Object)1, (global::System.Object)2].Borders((global::System.Object)10).LineStyle;
                                    MiddleBorder10C = this.rng.Cells[(global::System.Object)1, (global::System.Object)2].Borders((global::System.Object)10).Color;
                                    MiddleBorder10W = this.rng.Cells[(global::System.Object)1, (global::System.Object)2].Borders((global::System.Object)10).Weight;
                                }
                                else
                                {
                                    MiddleBorder10 = false;
                                    MiddleBorder10L = this.rng.Cells[(global::System.Object)1, (global::System.Object)2].Borders((global::System.Object)10).LineStyle;
                                    MiddleBorder10C = this.rng.Cells[(global::System.Object)1, (global::System.Object)2].Borders((global::System.Object)10).Color;
                                    MiddleBorder10W = this.rng.Cells[(global::System.Object)1, (global::System.Object)2].Borders((global::System.Object)10).Weight;
                                }

                                var loopTo312 = r;
                                for (i = 1; i <= loopTo312; i++)
                                {
                                    var loopTo313 = c;
                                    for (j = 1; j <= loopTo313; j++)
                                    {
                                        Arr[(i) - (1), (j) - (1)] = this.rng.Cells[(global::System.Object)i, (global::System.Object)j].Value;

                                        if (((this.CheckBox1.Checked) == (true)))
                                        {

                                            global::Microsoft.Office.Interop.Excel.Range cell = (global::Microsoft.Office.Interop.Excel.Range)this.rng.Cells[(global::System.Object)i, (global::System.Object)j];
                                            var font = cell.Font;

                                            Bolds[(i) - (1), (j) - (1)] = Conversions.ToBoolean(cell.Font.Bold);
                                            Italics[(i) - (1), (j) - (1)] = Conversions.ToBoolean(cell.Font.Italic);


                                            if ((((font.Name is System.DBNull)) == (false)))
                                            {
                                                fontNames[(i) - (1), (j) - (1)] = Conversions.ToString(font.Name);
                                            }
                                            else
                                            {
                                                fontNames[(i) - (1), (j) - (1)] = "Calibri";
                                            }

                                            if ((((font.Size is System.DBNull)) == (false)))
                                            {
                                                global::System.Single fontSize = global::System.Convert.ToSingle(font.Size);
                                                fontSizes[(i) - (1), (j) - (1)] = fontSize;
                                            }
                                            else
                                            {
                                                fontSizes[(i) - (1), (j) - (1)] = 11f;
                                            }

                                            if ((cell.Interior.Color is System.DBNull))
                                            {
                                                reds1[(i) - (1), (j) - (1)] = 255;
                                                greens1[(i) - (1), (j) - (1)] = 255;
                                                blues1[(i) - (1), (j) - (1)] = 255;
                                            }
                                            else
                                            {
                                                global::System.Int64 colorValue1 = Conversions.ToLong(cell.Interior.Color);
                                                global::System.Int32 red1 = (global::System.Int32)((colorValue1) % (256L));
                                                global::System.Int32 green1 = (global::System.Int32)(((((colorValue1) / (256L)))) % (256L));
                                                global::System.Int32 blue1 = (global::System.Int32)((((((colorValue1) / (256L)) / (256L)))) % (256L));
                                                reds1[(i) - (1), (j) - (1)] = red1;
                                                greens1[(i) - (1), (j) - (1)] = green1;
                                                blues1[(i) - (1), (j) - (1)] = blue1;
                                            }

                                            if ((cell.Font.Color is System.DBNull))
                                            {
                                                reds2[(i) - (1), (j) - (1)] = 0;
                                                greens2[(i) - (1), (j) - (1)] = 0;
                                                blues2[(i) - (1), (j) - (1)] = 0;
                                            }
                                            else
                                            {
                                                global::System.Int64 colorValue2 = Conversions.ToLong(cell.Font.Color);
                                                global::System.Int32 red2 = (global::System.Int32)((colorValue2) % (256L));
                                                global::System.Int32 green2 = (global::System.Int32)(((((colorValue2) / (256L)))) % (256L));
                                                global::System.Int32 blue2 = (global::System.Int32)((((((colorValue2) / (256L)) / (256L)))) % (256L));
                                                reds2[(i) - (1), (j) - (1)] = red2;
                                                greens2[(i) - (1), (j) - (1)] = green2;
                                                blues2[(i) - (1), (j) - (1)] = blue2;
                                            }
                                        }

                                    }
                                }

                                this.rng.ClearContents();
                                this.rng.ClearFormats();

                                this.rng2.ClearFormats();

                                var loopTo314 = c2;
                                for (j = 1; j <= loopTo314; j++)
                                {
                                    var loopTo315 = r2;
                                    for (i = 1; i <= loopTo315; i++)
                                    {
                                        global::System.Int32 x;
                                        global::System.Int32 y;
                                        x = 1;
                                        y = (((((r2) * ((((j) - (1))))))) + (i));
                                        if (((y) <= ((global::Microsoft.VisualBasic.Information.UBound(Arr, 2)) + (1))))
                                        {
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Value = Arr[(x) - (1), (y) - (1)];

                                            if (((this.CheckBox1.Checked) == (true)))
                                            {

                                                global::Microsoft.Office.Interop.Excel.Range cell2 = (global::Microsoft.Office.Interop.Excel.Range)this.rng2.Cells[(global::System.Object)i, (global::System.Object)j];
                                                var font2 = cell2.Font;

                                                global::System.Single fontSize = fontSizes[(x) - (1), (y) - (1)];

                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Font.Name = fontNames[(x) - (1), (y) - (1)];
                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Font.Size = (global::System.Object)fontSizes[(x) - (1), (y) - (1)];

                                                if (Bolds[(x) - (1), (y) - (1)])
                                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Font.Bold = (global::System.Object)true;
                                                if (Italics[(x) - (1), (y) - (1)])
                                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Font.Italic = (global::System.Object)true;

                                                global::System.Int32 red1 = reds1[(x) - (1), (y) - (1)];
                                                global::System.Int32 green1 = greens1[(x) - (1), (y) - (1)];
                                                global::System.Int32 blue1 = blues1[(x) - (1), (y) - (1)];
                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Interior.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(red1, green1, blue1);

                                                global::System.Int32 red2 = reds2[(x) - (1), (y) - (1)];
                                                global::System.Int32 green2 = greens2[(x) - (1), (y) - (1)];
                                                global::System.Int32 blue2 = blues2[(x) - (1), (y) - (1)];
                                                this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Font.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(red2, green2, blue2);
                                            }
                                        }
                                    }
                                }

                                if (((this.CheckBox1.Checked) == (true)))
                                {

                                    if (((TopBorder8) == (true)))
                                    {
                                        var loopTo316 = this.rng2.Columns.Count;
                                        for (j = 1; j <= loopTo316; j++)
                                        {
                                            this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)8).LineStyle = TopBorder8L;
                                            this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)8).Color = TopBorder8C;
                                            this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)8).Weight = TopBorder8W;
                                        }
                                    }
                                    else
                                    {
                                        var loopTo317 = this.rng2.Columns.Count;
                                        for (j = 1; j <= loopTo317; j++)
                                            this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)8).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                    }

                                    if (((this.rng2.Rows.Count) > (1)))
                                    {
                                        if (((TopBorder9) == (true)))
                                        {
                                            var loopTo318 = this.rng2.Columns.Count;
                                            for (j = 1; j <= loopTo318; j++)
                                            {
                                                this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)9).LineStyle = TopBorder9L;
                                                this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)9).Color = TopBorder9C;
                                                this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].Borders((global::System.Object)9).Weight = TopBorder9W;
                                            }
                                        }
                                        else
                                        {
                                            var loopTo319 = this.rng2.Columns.Count;
                                            for (j = 1; j <= loopTo319; j++)
                                                this.rng2.Cells[(global::System.Object)this.rng2.Rows.Count, (global::System.Object)j].Borders((global::System.Object)9).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                        }
                                    }

                                    if (((TopBorder7) == (true)))
                                    {
                                        var loopTo320 = this.rng2.Rows.Count;
                                        for (i = 1; i <= loopTo320; i++)
                                        {
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)1].Borders((global::System.Object)7).LineStyle = TopBorder7L;
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)1].Borders((global::System.Object)7).Color = TopBorder7C;
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)1].Borders((global::System.Object)7).Weight = TopBorder7W;
                                        }
                                    }
                                    else
                                    {
                                        var loopTo321 = this.rng2.Rows.Count;
                                        for (i = 1; i <= loopTo321; i++)
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)1].Borders((global::System.Object)7).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                    }

                                    if (((BottomBorder9) == (true)))
                                    {
                                        var loopTo322 = this.rng2.Columns.Count;
                                        for (j = 1; j <= loopTo322; j++)
                                        {
                                            this.rng2.Cells[(global::System.Object)this.rng2.Rows.Count, (global::System.Object)j].Borders((global::System.Object)9).LineStyle = BottomBorder9L;
                                            this.rng2.Cells[(global::System.Object)this.rng2.Rows.Count, (global::System.Object)j].Borders((global::System.Object)9).Color = BottomBorder9C;
                                            this.rng2.Cells[(global::System.Object)this.rng2.Rows.Count, (global::System.Object)j].Borders((global::System.Object)9).Weight = BottomBorder9W;
                                        }
                                    }
                                    else
                                    {
                                        var loopTo323 = this.rng2.Columns.Count;
                                        for (j = 1; j <= loopTo323; j++)
                                            this.rng2.Cells[(global::System.Object)this.rng2.Rows.Count, (global::System.Object)this.rng2.Columns.Count].Borders((global::System.Object)9).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                    }

                                    if (((BottomBorder10) == (true)))
                                    {
                                        var loopTo324 = this.rng2.Rows.Count;
                                        for (i = 1; i <= loopTo324; i++)
                                        {
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)this.rng2.Columns.Count].Borders((global::System.Object)10).LineStyle = BottomBorder10L;
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)this.rng2.Columns.Count].Borders((global::System.Object)10).Color = BottomBorder10C;
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)this.rng2.Columns.Count].Borders((global::System.Object)10).Weight = BottomBorder10W;
                                        }
                                    }
                                    else
                                    {
                                        var loopTo325 = this.rng2.Rows.Count;
                                        for (i = 1; i <= loopTo325; i++)
                                            this.rng2.Cells[(global::System.Object)i, (global::System.Object)this.rng2.Columns.Count].Borders((global::System.Object)10).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                    }

                                    if (((this.rng.Rows.Count) > (1)))
                                    {

                                        if (((MiddleBorder9) == (true)))
                                        {
                                            var loopTo326 = (this.rng2.Rows.Count) - (1);
                                            for (i = 2; i <= loopTo326; i++)
                                            {
                                                var loopTo327 = this.rng2.Columns.Count;
                                                for (j = 1; j <= loopTo327; j++)
                                                {
                                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)9).LineStyle = MiddleBorder9L;
                                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)9).Color = MiddleBorder9C;
                                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)9).Weight = MiddleBorder9W;
                                                }
                                            }
                                        }
                                        else
                                        {
                                            var loopTo328 = (this.rng2.Rows.Count) - (1);
                                            for (i = 2; i <= loopTo328; i++)
                                            {
                                                var loopTo329 = this.rng2.Columns.Count;
                                                for (j = 1; j <= loopTo329; j++)
                                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)9).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                            }
                                        }

                                    }

                                    if (((this.rng.Columns.Count) > (1)))
                                    {

                                        if (((MiddleBorder10) == (true)))
                                        {
                                            var loopTo330 = (this.rng2.Columns.Count) - (1);
                                            for (j = 1; j <= loopTo330; j++)
                                            {
                                                var loopTo331 = this.rng2.Rows.Count;
                                                for (i = 1; i <= loopTo331; i++)
                                                {
                                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)10).LineStyle = MiddleBorder10L;
                                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)10).Color = MiddleBorder10C;
                                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)10).Weight = MiddleBorder10W;
                                                }
                                            }
                                        }
                                        else
                                        {
                                            var loopTo332 = (this.rng2.Columns.Count) - (1);
                                            for (j = 1; j <= loopTo332; j++)
                                            {
                                                var loopTo333 = this.rng2.Rows.Count;
                                                for (i = 1; i <= loopTo333; i++)
                                                    this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].Borders((global::System.Object)10).LineStyle = global::Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone;
                                            }
                                        }

                                    }

                                }

                            }
                        }
                    }

                    else
                    {
                        global::System.Windows.Forms.MessageBox.Show("Select One Separator.", "Error", global::System.Windows.Forms.MessageBoxButtons.OK, global::System.Windows.Forms.MessageBoxIcon.Error);
                        return;
                    }
                }

                else
                {
                    global::System.Windows.Forms.MessageBox.Show("Select One Transformation Type.", "Error", global::System.Windows.Forms.MessageBoxButtons.OK, global::System.Windows.Forms.MessageBoxIcon.Error);
                    return;
                }

                var loopTo334 = this.rng2.Columns.Count;
                for (j = 1; j <= loopTo334; j++)
                    this.rng2.Columns[(global::System.Object)j].Autofit();

                this.Close();

                this.TextBoxChanged = false;
            }

            catch (global::System.Exception ex)
            {

            }

        }

        private void TextBox1_TextChanged(global::System.Object sender, global::System.EventArgs e)
        {

            try
            {
                this.excelApp = global::VSTO_Addins.Globals.ThisAddIn.Application;
                this.workbook = excelApp.ActiveWorkbook;
                this.worksheet = (global::Microsoft.Office.Interop.Excel.Worksheet)workbook.ActiveSheet;

                global::System.String[] rngArray = global::Microsoft.VisualBasic.Strings.Split(this.TextBox1.Text, "!");
                global::System.String rngAddress = rngArray[global::Microsoft.VisualBasic.Information.UBound(rngArray)];
                this.rng = this.worksheet.get_Range(rngAddress);
                this.TextBoxChanged = true;
                this.rng.Select();
                this.Display();
                this.Setup();
                this.TextBoxChanged = false;
            }
            catch (global::System.Exception ex)
            {

            }

        }

        private void RadioButton1_CheckedChanged(global::System.Object sender, global::System.EventArgs e)
        {

            try
            {
                if (((this.RadioButton1.Checked) == (true)))
                {
                    this.Display();
                    this.Setup();
                }
            }
            catch (global::System.Exception ex)
            {

            }

        }

        private void RadioButton3_CheckedChanged(global::System.Object sender, global::System.EventArgs e)
        {

            try
            {
                if (((this.RadioButton3.Checked) == (true)))
                {
                    this.Display();
                    this.Setup();
                }
            }

            catch (global::System.Exception ex)
            {

            }

        }

        private void RadioButton2_CheckedChanged(global::System.Object sender, global::System.EventArgs e)
        {

            try
            {
                if (((this.RadioButton2.Checked) == (true)))
                {
                    this.Display();
                    this.Setup();
                }
            }
            catch (global::System.Exception ex)
            {

            }

        }

        private void RadioButton4_CheckedChanged(global::System.Object sender, global::System.EventArgs e)
        {

            try
            {
                if (((this.RadioButton4.Checked) == (true)))
                {
                    this.Display();
                    this.Setup();
                }
            }
            catch (global::System.Exception ex)
            {

            }

        }

        private void CheckBox1_CheckedChanged(global::System.Object sender, global::System.EventArgs e)
        {

            try
            {
                this.Display();
                this.Setup();
            }
            catch (global::System.Exception ex)
            {

            }

        }

        private void RadioButton5_CheckedChanged(global::System.Object sender, global::System.EventArgs e)
        {

            try
            {
                if (((this.RadioButton5.Checked) == (true)))
                {
                    this.Display();
                    this.Setup();
                }
            }
            catch (global::System.Exception ex)
            {

            }

        }

        private void RadioButton6_CheckedChanged(global::System.Object sender, global::System.EventArgs e)
        {

            try
            {
                if (((this.RadioButton6.Checked) == (true)))
                {
                    this.Display();
                    this.Setup();
                }
            }
            catch (global::System.Exception ex)
            {

            }

        }

        private void RadioButton7_CheckedChanged(global::System.Object sender, global::System.EventArgs e)
        {

            try
            {
                if (((this.RadioButton7.Checked) == (true)))
                {
                    this.Display();
                    this.Setup();
                }
            }
            catch (global::System.Exception ex)
            {

            }

        }

        private void RadioButton8_CheckedChanged(global::System.Object sender, global::System.EventArgs e)
        {

            try
            {
                if (((this.RadioButton8.Checked) == (true)))
                {
                    this.Display();
                    this.Setup();
                }
            }
            catch (global::System.Exception ex)
            {

            }

        }

        private void TextBox2_TextChanged(global::System.Object sender, global::System.EventArgs e)
        {

            try
            {
                this.Display();
                this.Setup();
            }
            catch (global::System.Exception ex)
            {

            }

        }

        private void PictureBox4_Click(global::System.Object sender, global::System.EventArgs e)
        {

            try
            {
                this.FocusedTextBox = 1;

                var activeRange = excelApp.ActiveCell;

                global::System.Int32 startRow = activeRange.Row;
                global::System.Int32 startColumn = activeRange.Column;
                global::System.Int32 endRow = activeRange.Row;
                global::System.Int32 endColumn = activeRange.Column;

                // Find the upper boundary
                while ((((startRow) > (1)) && (!((worksheet.Cells[(global::System.Object)((startRow) - (1)), (global::System.Object)startColumn].Value == null)))))
                    startRow -= 1;

                // Find the lower boundary
                while (!((worksheet.Cells[(global::System.Object)((endRow) + (1)), (global::System.Object)endColumn].Value == null)))
                    endRow += 1;

                // Find the left boundary
                while ((((startColumn) > (1)) && (!((worksheet.Cells[(global::System.Object)startRow, (global::System.Object)((startColumn) - (1))].Value == null)))))
                    startColumn -= 1;

                // Find the right boundary
                while (!((worksheet.Cells[(global::System.Object)endRow, (global::System.Object)((endColumn) + (1))].Value == null)))
                    endColumn += 1;

                // Select the determined range
                this.rng = this.worksheet.get_Range(worksheet.Cells[(global::System.Object)startRow, (global::System.Object)startColumn], worksheet.Cells[(global::System.Object)endRow, (global::System.Object)endColumn]);

                this.rng.Select();

                global::System.String sheetName;

                sheetName = global::Microsoft.VisualBasic.Strings.Split(this.rng.get_Address((global::System.Object)true, (global::System.Object)true, global::Microsoft.Office.Interop.Excel.XlReferenceStyle.xlA1, (global::System.Object)true), "]")[1];
                sheetName = global::Microsoft.VisualBasic.Strings.Split(sheetName, "!")[0];

                if ((global::Microsoft.VisualBasic.Strings.Mid(sheetName, global::Microsoft.VisualBasic.Strings.Len(sheetName), 1) == "'"))
                {
                    sheetName = global::Microsoft.VisualBasic.Strings.Mid(sheetName, 1, (global::Microsoft.VisualBasic.Strings.Len(sheetName)) - (1));
                }

                this.worksheet = (global::Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets[sheetName];
                this.worksheet.Activate();

                if (((worksheet.Name ?? "") != (OpenSheet.Name ?? "")))
                {
                    this.TextBox1.Text = ((worksheet.Name + "!") + this.rng.get_Address());
                }
                else
                {
                    this.TextBox1.Text = this.rng.get_Address();
                }

                this.TextBox1.Focus();
            }

            catch (global::System.Exception ex)
            {

                this.Show();
                this.TextBox1.Focus();

            }

        }

        private void PictureBox8_Click(global::System.Object sender, global::System.EventArgs e)
        {

            try
            {
                this.FocusedTextBox = 1;

                global::Microsoft.Office.Interop.Excel.Range userInput = (global::Microsoft.Office.Interop.Excel.Range)this.excelApp.InputBox("Select a range", Type: (global::System.Object)8);
                var rng = userInput;

                try
                {
                    global::System.String sheetName;
                    sheetName = global::Microsoft.VisualBasic.Strings.Split(rng.get_Address((global::System.Object)true, (global::System.Object)true, global::Microsoft.Office.Interop.Excel.XlReferenceStyle.xlA1, (global::System.Object)true), "]")[1];
                    sheetName = global::Microsoft.VisualBasic.Strings.Split(sheetName, "!")[0];

                    if ((global::Microsoft.VisualBasic.Strings.Mid(sheetName, global::Microsoft.VisualBasic.Strings.Len(sheetName), 1) == "'"))
                    {
                        sheetName = global::Microsoft.VisualBasic.Strings.Mid(sheetName, 1, (global::Microsoft.VisualBasic.Strings.Len(sheetName)) - (1));
                    }

                    this.worksheet = (global::Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets[sheetName];
                    this.worksheet.Activate();
                }
                catch (global::System.Exception ex)
                {

                }

                rng.Select();

                if (((worksheet.Name ?? "") != (OpenSheet.Name ?? "")))
                {
                    this.TextBox1.Text = ((worksheet.Name + "!") + rng.get_Address());
                }
                else
                {
                    this.TextBox1.Text = rng.get_Address();
                }

                this.TextBox1.Focus();
            }

            catch (global::System.Exception ex)
            {

            }

        }

        private void Button1_Click(global::System.Object sender, global::System.EventArgs e)
        {
            try
            {
                this.Close();
            }
            catch (global::System.Exception ex)
            {

            }
        }

        private void PictureBox6_Click(global::System.Object sender, global::System.EventArgs e)
        {

            try
            {
                this.FocusedTextBox = 3;
                this.Hide();

                this.excelApp = global::VSTO_Addins.Globals.ThisAddIn.Application;
                this.workbook = excelApp.ActiveWorkbook;

                global::Microsoft.Office.Interop.Excel.Range userInput = (global::Microsoft.Office.Interop.Excel.Range)this.excelApp.InputBox("Select a range", Type: (global::System.Object)8);
                this.rng2 = userInput;


                global::System.String sheetName;
                sheetName = global::Microsoft.VisualBasic.Strings.Split(this.rng2.get_Address((global::System.Object)true, (global::System.Object)true, global::Microsoft.Office.Interop.Excel.XlReferenceStyle.xlA1, (global::System.Object)true), "]")[1];
                sheetName = global::Microsoft.VisualBasic.Strings.Split(sheetName, "!")[0];

                if ((global::Microsoft.VisualBasic.Strings.Mid(sheetName, global::Microsoft.VisualBasic.Strings.Len(sheetName), 1) == "'"))
                {
                    sheetName = global::Microsoft.VisualBasic.Strings.Mid(sheetName, 1, (global::Microsoft.VisualBasic.Strings.Len(sheetName)) - (1));
                }

                this.worksheet2 = (global::Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets[sheetName];
                this.worksheet2.Activate();

                this.rng2.Select();

                this.TextBox3.Text = this.rng2.get_Address();

                this.Show();
                this.TextBox3.Focus();
            }

            catch (global::System.Exception ex)
            {

                this.Show();
                this.TextBox3.Focus();

            }

        }

        private void RadioButton10_CheckedChanged(global::System.Object sender, global::System.EventArgs e)
        {

            try
            {
                if (((this.RadioButton10.Checked) == (true)))
                {
                    this.TextBox3.Enabled = true;
                    this.TextBox3.Focus();
                }
                else
                {
                    this.TextBox3.Clear();
                    this.TextBox3.Enabled = false;
                }
            }
            catch (global::System.Exception ex)
            {

            }

        }

        private void TextBox1_GotFocus(global::System.Object sender, global::System.EventArgs e)
        {

            try
            {
                this.FocusedTextBox = 1;
            }
            catch (global::System.Exception ex)
            {

            }

        }

        private void TextBox3_GotFocus(global::System.Object sender, global::System.EventArgs e)
        {

            try
            {
                this.FocusedTextBox = 3;
            }
            catch (global::System.Exception ex)
            {

            }

        }

        private void ComboBox1_SelectedIndexChanged(global::System.Object sender, global::System.EventArgs e)
        {

            try
            {
                if (Conversions.ToBoolean(Operators.AndObject(Operators.ConditionalCompareObjectEqual(this.ComboBox1.SelectedItem, "SOFTEKO", false), ((this.opened) >= (1)))))
                {

                    global::System.String url = "https://www.softeko.co";
                    global::System.Diagnostics.Process.Start(url);

                }
            }
            catch (global::System.Exception ex)
            {

            }

        }

        private void TextBox3_TextChanged(global::System.Object sender, global::System.EventArgs e)
        {

            try
            {
                this.excelApp = global::VSTO_Addins.Globals.ThisAddIn.Application;
                this.workbook2 = excelApp.ActiveWorkbook;
                this.worksheet2 = (global::Microsoft.Office.Interop.Excel.Worksheet)workbook2.ActiveSheet;

                global::System.String[] rng2Array = global::Microsoft.VisualBasic.Strings.Split(this.TextBox3.Text, "!");
                global::System.String rng2Address = rng2Array[global::Microsoft.VisualBasic.Information.UBound(rng2Array)];
                this.rng2 = this.worksheet2.get_Range(rng2Address);

                this.TextBoxChanged = true;

                this.rng2.Select();

                this.TextBoxChanged = false;
            }

            catch (global::System.Exception ex)
            {

            }

        }

        private void RadioButton9_CheckedChanged(global::System.Object sender, global::System.EventArgs e)
        {

            try
            {
                if (((this.RadioButton9.Checked) == (true)))
                {
                    this.worksheet2 = this.worksheet;
                    this.rng2 = this.rng;
                }
            }
            catch (global::System.Exception ex)
            {

            }

        }

        private void Button1_GotFocus(global::System.Object sender, global::System.EventArgs e)
        {

            try
            {
                this.FocusedTextBox = 0;
            }

            catch (global::System.Exception ex)
            {

            }

        }

        private void Button2_GotFocus(global::System.Object sender, global::System.EventArgs e)
        {

            try
            {
                this.FocusedTextBox = 0;
            }

            catch (global::System.Exception ex)
            {

            }

        }

        private void CheckBox1_GotFocus(global::System.Object sender, global::System.EventArgs e)
        {

            try
            {
                this.FocusedTextBox = 0;
            }

            catch (global::System.Exception ex)
            {

            }

        }

        private void CheckBox2_GotFocus(global::System.Object sender, global::System.EventArgs e)
        {

            try
            {
                this.FocusedTextBox = 0;
            }

            catch (global::System.Exception ex)
            {

            }

        }

        private void ComboBox1_GotFocus(global::System.Object sender, global::System.EventArgs e)
        {

            try
            {
                this.FocusedTextBox = 0;
            }

            catch (global::System.Exception ex)
            {

            }

        }

        private void CustomGroupBox1_GotFocus(global::System.Object sender, global::System.EventArgs e)
        {

            try
            {
                this.FocusedTextBox = 0;
            }

            catch (global::System.Exception ex)
            {

            }

        }

        private void CustomGroupBox10_GotFocus(global::System.Object sender, global::System.EventArgs e)
        {

            try
            {
                this.FocusedTextBox = 0;
            }

            catch (global::System.Exception ex)
            {

            }

        }

        private void CustomGroupBox2_GotFocus(global::System.Object sender, global::System.EventArgs e)
        {

            try
            {
                this.FocusedTextBox = 0;
            }

            catch (global::System.Exception ex)
            {

            }

        }

        private void CustomGroupBox3_GotFocus(global::System.Object sender, global::System.EventArgs e)
        {

            try
            {
                this.FocusedTextBox = 0;
            }

            catch (global::System.Exception ex)
            {

            }

        }

        private void CustomGroupBox4_GotFocus(global::System.Object sender, global::System.EventArgs e)
        {

            try
            {
                this.FocusedTextBox = 0;
            }

            catch (global::System.Exception ex)
            {

            }

        }

        private void CustomGroupBox5_GotFocus(global::System.Object sender, global::System.EventArgs e)
        {

            try
            {
                this.FocusedTextBox = 0;
            }

            catch (global::System.Exception ex)
            {

            }

        }

        private void CustomGroupBox6_GotFocus(global::System.Object sender, global::System.EventArgs e)
        {

            try
            {
                this.FocusedTextBox = 0;
            }

            catch (global::System.Exception ex)
            {

            }

        }

        private void CustomGroupBox7_GotFocus(global::System.Object sender, global::System.EventArgs e)
        {

            try
            {
                this.FocusedTextBox = 0;
            }

            catch (global::System.Exception ex)
            {

            }

        }

        private void CustomGroupBox8_GotFocus(global::System.Object sender, global::System.EventArgs e)
        {

            try
            {
                this.FocusedTextBox = 0;
            }

            catch (global::System.Exception ex)
            {

            }

        }

        private void CustomGroupBox9_GotFocus(global::System.Object sender, global::System.EventArgs e)
        {

            try
            {
                this.FocusedTextBox = 0;
            }

            catch (global::System.Exception ex)
            {

            }

        }

        private void CustomPanel1_GotFocus(global::System.Object sender, global::System.EventArgs e)
        {

            try
            {
                this.FocusedTextBox = 0;
            }

            catch (global::System.Exception ex)
            {

            }

        }

        private void CustomPanel2_GotFocus(global::System.Object sender, global::System.EventArgs e)
        {

            try
            {
                this.FocusedTextBox = 0;
            }

            catch (global::System.Exception ex)
            {

            }

        }

        private void Label1_GotFocus(global::System.Object sender, global::System.EventArgs e)
        {

            try
            {
                this.FocusedTextBox = 0;
            }

            catch (global::System.Exception ex)
            {

            }

        }

        private void Label2_GotFocus(global::System.Object sender, global::System.EventArgs e)
        {

            try
            {
                this.FocusedTextBox = 0;
            }

            catch (global::System.Exception ex)
            {

            }

        }

        private void PictureBox1_GotFocus(global::System.Object sender, global::System.EventArgs e)
        {

            try
            {
                this.FocusedTextBox = 0;
            }

            catch (global::System.Exception ex)
            {

            }

        }

        private void PictureBox2_GotFocus(global::System.Object sender, global::System.EventArgs e)
        {

            try
            {
                this.FocusedTextBox = 0;
            }

            catch (global::System.Exception ex)
            {

            }

        }

        private void PictureBox3_GotFocus(global::System.Object sender, global::System.EventArgs e)
        {

            try
            {
                this.FocusedTextBox = 0;
            }

            catch (global::System.Exception ex)
            {

            }

        }

        private void PictureBox4_GotFocus(global::System.Object sender, global::System.EventArgs e)
        {

            try
            {
                this.FocusedTextBox = 1;
            }

            catch (global::System.Exception ex)
            {

            }

        }

        private void PictureBox5_GotFocus(global::System.Object sender, global::System.EventArgs e)
        {

            try
            {
                this.FocusedTextBox = 0;
            }

            catch (global::System.Exception ex)
            {

            }

        }

        private void PictureBox6_GotFocus(global::System.Object sender, global::System.EventArgs e)
        {

            try
            {
                this.FocusedTextBox = 3;
            }

            catch (global::System.Exception ex)
            {

            }

        }

        private void PictureBox7_GotFocus(global::System.Object sender, global::System.EventArgs e)
        {

            try
            {
                this.FocusedTextBox = 0;
            }

            catch (global::System.Exception ex)
            {

            }

        }

        private void PictureBox8_GotFocus(global::System.Object sender, global::System.EventArgs e)
        {

            try
            {
                this.FocusedTextBox = 1;
            }

            catch (global::System.Exception ex)
            {

            }

        }

        private void RadioButton1_GotFocus(global::System.Object sender, global::System.EventArgs e)
        {

            try
            {
                this.FocusedTextBox = 0;
            }

            catch (global::System.Exception ex)
            {

            }

        }

        private void RadioButton10_GotFocus(global::System.Object sender, global::System.EventArgs e)
        {

            try
            {
                this.FocusedTextBox = 0;
            }

            catch (global::System.Exception ex)
            {

            }

        }

        private void RadioButton2_GotFocus(global::System.Object sender, global::System.EventArgs e)
        {

            try
            {
                this.FocusedTextBox = 0;
            }

            catch (global::System.Exception ex)
            {

            }

        }

        private void RadioButton3_GotFocus(global::System.Object sender, global::System.EventArgs e)
        {

            try
            {
                this.FocusedTextBox = 0;
            }

            catch (global::System.Exception ex)
            {

            }

        }

        private void RadioButton4_GotFocus(global::System.Object sender, global::System.EventArgs e)
        {

            try
            {
                this.FocusedTextBox = 0;
            }

            catch (global::System.Exception ex)
            {

            }

        }

        private void RadioButton5_GotFocus(global::System.Object sender, global::System.EventArgs e)
        {

            try
            {
                this.FocusedTextBox = 0;
            }

            catch (global::System.Exception ex)
            {

            }

        }

        private void RadioButton6_GotFocus(global::System.Object sender, global::System.EventArgs e)
        {

            try
            {
                this.FocusedTextBox = 0;
            }

            catch (global::System.Exception ex)
            {

            }

        }

        private void RadioButton7_GotFocus(global::System.Object sender, global::System.EventArgs e)
        {

            try
            {
                this.FocusedTextBox = 0;
            }

            catch (global::System.Exception ex)
            {

            }

        }

        private void RadioButton8_GotFocus(global::System.Object sender, global::System.EventArgs e)
        {

            try
            {
                this.FocusedTextBox = 0;
            }

            catch (global::System.Exception ex)
            {

            }

        }

        private void RadioButton9_GotFocus(global::System.Object sender, global::System.EventArgs e)
        {

            try
            {
                this.FocusedTextBox = 0;
            }

            catch (global::System.Exception ex)
            {

            }

        }

        private void TextBox2_GotFocus(global::System.Object sender, global::System.EventArgs e)
        {

            try
            {
                this.FocusedTextBox = 0;
            }

            catch (global::System.Exception ex)
            {

            }

        }

        private void VScrollBar1_GotFocus(global::System.Object sender, global::System.EventArgs e)
        {

            try
            {
                this.FocusedTextBox = 0;
            }

            catch (global::System.Exception ex)
            {

            }

        }

        private void Button2_MouseEnter(global::System.Object sender, global::System.EventArgs e)
        {
            try
            {
                this.Button2.BackColor = global::System.Drawing.Color.FromArgb(65, 105, 225);
                this.Button2.ForeColor = global::System.Drawing.Color.FromArgb(255, 255, 255);
            }
            catch (global::System.Exception ex)
            {

            }
        }

        private void Button1_MouseEnter(global::System.Object sender, global::System.EventArgs e)
        {

            try
            {
                this.Button1.BackColor = global::System.Drawing.Color.FromArgb(65, 105, 225);
                this.Button1.ForeColor = global::System.Drawing.Color.FromArgb(255, 255, 255);
            }
            catch (global::System.Exception ex)
            {

            }

        }

        private void Button2_MouseLeave(global::System.Object sender, global::System.EventArgs e)
        {
            try
            {

                this.Button2.BackColor = global::System.Drawing.Color.FromArgb(255, 255, 255);
                this.Button2.ForeColor = global::System.Drawing.Color.FromArgb(70, 70, 70);
            }
            catch (global::System.Exception ex)
            {

            }
        }

        private void Button1_MouseLeave(global::System.Object sender, global::System.EventArgs e)
        {
            try
            {

                this.Button1.BackColor = global::System.Drawing.Color.FromArgb(255, 255, 255);
                this.Button1.ForeColor = global::System.Drawing.Color.FromArgb(70, 70, 70);
            }
            catch (global::System.Exception ex)
            {

            }
        }

        private void Form7_Closing(global::System.Object sender, global::System.ComponentModel.CancelEventArgs e)
        {
            try
            {
                global::VSTO_Addins.GlobalModule.form_flag = false;
            }

            catch (global::System.Exception ex)
            {

            }

        }

        private void Form7_Shown(global::System.Object sender, global::System.EventArgs e)
        {

            try
            {
                this.Focus();
                this.BringToFront();
                this.Activate();

                global::System.String TextBoxText;

                if (((worksheet.Name ?? "") != (OpenSheet.Name ?? "")))
                {
                    TextBoxText = ((worksheet.Name + "!") + this.rng.get_Address());
                }
                else
                {
                    TextBoxText = this.rng.get_Address();
                }

                this.BeginInvoke(new global::System.Action(() =>
                    {
                        this.TextBox1.Text = TextBoxText;
                        global::VSTO_Addins.Form7.SetWindowPos(this.Handle, new global::System.IntPtr(global::VSTO_Addins.Form7.HWND_TOPMOST), 0, 0, 0, 0, ((global::VSTO_Addins.Form7.SWP_NOACTIVATE) | (global::VSTO_Addins.Form7.SWP_NOMOVE)) | (global::VSTO_Addins.Form7.SWP_NOSIZE));
                    }));
            }

            catch (global::System.Exception ex)
            {

            }

        }

        private void Form7_Disposed(global::System.Object sender, global::System.EventArgs e)
        {

            try
            {
                global::VSTO_Addins.GlobalModule.form_flag = false;
            }

            catch (global::System.Exception ex)
            {

            }

        }

        private void Form7_KeyDown(global::System.Object sender, global::System.Windows.Forms.KeyEventArgs e)
        {

            try
            {
                if ((e.KeyCode == global::System.Windows.Forms.Keys.Enter))
                {

                    this.Button2.Focus();
                    this.Button2_Click(sender, e);

                }
            }

            catch (global::System.Exception ex)
            {

            }

        }

    }
}