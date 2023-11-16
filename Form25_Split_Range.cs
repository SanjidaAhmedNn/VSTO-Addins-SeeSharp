using System;
using System.Collections.Generic;
using global::System.ComponentModel;
using global::System.Diagnostics;
using global::System.Drawing;
using System.Linq;
using global::System.Reflection;
using global::System.Reflection.Emit;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using global::System.Security.Policy;
using System.Text;
using global::System.Text.RegularExpressions;
using global::System.Windows.Forms;
using static global::System.Windows.Forms.VisualStyles.VisualStyleElement;
using static global::System.Windows.Forms.VisualStyles.VisualStyleElement.Button;
using System.Xml.Linq;
using Office = Microsoft.Office.Core;
using global::Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace VSTO_Addins
{

    public partial class Form25_Split_Range
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
        private global::Microsoft.Office.Interop.Excel.Workbook workBook;
        private global::Microsoft.Office.Interop.Excel.Worksheet workSheet;
        private global::Microsoft.Office.Interop.Excel.Worksheet workSheet2;
        private global::Microsoft.Office.Interop.Excel.Range rng;
        private global::Microsoft.Office.Interop.Excel.Range rng2;
        private global::Microsoft.Office.Interop.Excel.Range selectedRange;

        private global::System.Int32 opened;
        private global::System.Int32 FocusedTextBox;
        private global::System.Boolean TextBoxChanged;

        public Form25_Split_Range()
        {
            InitializeComponent();
        }

        [DllImport("user32")]
        private static extern bool SetWindowPos(global::System.IntPtr hWnd, global::System.IntPtr hWndInsertAfter, global::System.Int32 X, global::System.Int32 Y, global::System.Int32 cx, global::System.Int32 cy, global::System.UInt32 uFlags);
        private const global::System.UInt32 SWP_NOMOVE = 0x2U;
        private const global::System.UInt32 SWP_NOSIZE = 0x1U;
        private const global::System.UInt32 SWP_NOACTIVATE = 0x10U;
        private const global::System.Int32 HWND_TOPMOST = -(1);

        private global::System.Object MaxOfColumn(global::Microsoft.Office.Interop.Excel.Range cRng)
        {
            global::System.Object MaxOfColumnRet = default(global::System.Object);

            global::System.Int32 max;
            max = global::Microsoft.VisualBasic.Strings.Len(cRng.Cells[(global::System.Object)1, (global::System.Object)1].value);

            for (global::System.Int32 i = 2, loopTo = cRng.Rows.Count; i <= loopTo; i++)
            {
                if (((global::Microsoft.VisualBasic.Strings.Len(cRng.Cells[(global::System.Object)i, (global::System.Object)1].value)) > (max)))
                {
                    max = global::Microsoft.VisualBasic.Strings.Len(cRng.Cells[(global::System.Object)i, (global::System.Object)1].value);
                }
            }

            if (((max) < (7)))
            {
                max = 7;
            }

            MaxOfColumnRet = (global::System.Object)max;
            return MaxOfColumnRet;

        }
        private global::System.Boolean IsValidExcelCellReference(global::System.String cellReference)
        {

            global::System.String cellPattern = @"(\$?[A-Z]+\$?[0-9]+)";

            global::System.String referencePattern = ((((("^") + (cellPattern)) + ("(:")) + (cellPattern)) + (")?$"));

            var regex = new global::System.Text.RegularExpressions.Regex(referencePattern);

            if (regex.IsMatch(cellReference))
            {
                return true;
            }
            else
            {
                return false;
            }

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
        private global::System.Object SeparateNumberText(global::System.String Str)
        {
            global::System.Object SeparateNumberTextRet = default(global::System.Object);

            var Output = new global::System.String[2];
            Output[0] = "";
            Output[1] = "";

            for (global::System.Int32 i = 1, loopTo = global::Microsoft.VisualBasic.Strings.Len(Str); i <= loopTo; i++)
            {
                if (global::Microsoft.VisualBasic.Information.IsNumeric(global::Microsoft.VisualBasic.Strings.Mid(Str, i, 1)))
                {
                    Output[0] = (Output[0] + global::Microsoft.VisualBasic.Strings.Mid(Str, i, 1));
                }
                else
                {
                    Output[1] = (Output[1] + global::Microsoft.VisualBasic.Strings.Mid(Str, i, 1));
                }
            }

            SeparateNumberTextRet = Output;
            return SeparateNumberTextRet;

        }
        public global::System.Int32 CountSeparator(global::System.String source, global::System.String separator)
        {
            global::System.Int32 CountSeparatorRet = default(global::System.Int32);

            global::System.Int32 count = 0;
            global::System.Int32 Position = 1;

            for (global::System.Int32 i = 1, loopTo = global::Microsoft.VisualBasic.Strings.Len(source); i <= loopTo; i++)
            {
                if (((global::Microsoft.VisualBasic.Strings.Mid(source, i, global::Microsoft.VisualBasic.Strings.Len(separator)) ?? "") == (separator ?? "")))
                {
                    if ((((i) - (Position)) > (0)))
                    {
                        count = ((count) + (1));
                    }
                    Position = ((i) + (global::Microsoft.VisualBasic.Strings.Len(separator)));
                }
            }

            if (((Position) <= (global::Microsoft.VisualBasic.Strings.Len(source))))
            {
                count = ((count) + (1));
            }

            CountSeparatorRet = count;
            return CountSeparatorRet;

        }
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

        private void Display()
        {

            try
            {
                this.CustomPanel1.Controls.Clear();
                this.CustomPanel2.Controls.Clear();

                global::Microsoft.Office.Interop.Excel.Range displayRng;

                if (((this.rng.Rows.Count) > (50)))
                {
                    displayRng = (global::Microsoft.Office.Interop.Excel.Range)this.rng.Rows["1:50"];
                }
                else
                {
                    displayRng = this.rng;
                }

                global::System.Int32 r = displayRng.Rows.Count;
                global::System.Int32 c = displayRng.Columns.Count;

                global::System.Double Height;
                global::System.Double BaseWidth;
                global::System.Double Width;

                if (((r) <= (4)))
                {
                    Height = ((global::System.Double)(this.CustomPanel1.Height) / (global::System.Double)(displayRng.Rows.Count));
                }
                else
                {
                    Height = (((119d) / (4d)));
                }

                BaseWidth = (((260d) / (3d)));

                global::System.Double ordinate = 0d;

                for (global::System.Int32 j = 1, loopTo = c; j <= loopTo; j++)
                {
                    global::Microsoft.Office.Interop.Excel.Range cRng = (global::Microsoft.Office.Interop.Excel.Range)displayRng.Columns[(global::System.Object)j];
                    Width = Conversions.ToDouble(Operators.DivideObject((Operators.MultiplyObject(this.MaxOfColumn(cRng), BaseWidth)), 10));
                    for (global::System.Int32 i = 1, loopTo1 = r; i <= loopTo1; i++)
                    {
                        var label = new global::System.Windows.Forms.Label();
                        label.Text = Conversions.ToString(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Value);
                        label.Location = new global::System.Drawing.Point((global::System.Int32)Math.Round(ordinate), (global::System.Int32)Math.Round((global::System.Double)((((i) - (1)))) * (Height)));
                        label.Height = (global::System.Int32)Math.Round(Height);
                        label.Width = (global::System.Int32)Math.Round(Width);
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
                    ordinate = ((ordinate) + (Width));
                }

                this.CustomPanel1.AutoScroll = true;

                global::System.Boolean X1 = this.RadioButton1.Checked;
                global::System.Boolean X2 = this.RadioButton2.Checked;
                global::System.Boolean X3 = this.RadioButton3.Checked;
                global::System.Boolean X7 = this.RadioButton7.Checked;
                global::System.Boolean X8 = this.RadioButton8.Checked;
                global::System.Boolean X9 = this.RadioButton9.Checked;
                global::System.Boolean X10 = this.RadioButton10.Checked;
                global::System.Boolean X11 = this.RadioButton11.Checked;
                global::System.Boolean X12 = ((this.ComboBox3.SelectedIndex) != (-(1)));

                if (((((((X1) | (X2)))) & (X12)) & ((((((((X3) | (X7)) | (X8)) | (X9)) | (X10)) | (X11))))))
                {

                    global::System.Int32 SplitColumn = ((this.ComboBox3.SelectedIndex) + (1));

                    if (((((X7) | (X8)) | (X9)) | (X10)))
                    {

                        global::System.String Separator = "";
                        if (X7)
                        {
                            Separator = ";";
                        }
                        else if (X8)
                        {
                            Separator = global::Microsoft.VisualBasic.Constants.vbNewLine;
                        }
                        else if (X9)
                        {
                            Separator = " ";
                        }
                        else if (X10)
                        {
                            Separator = this.ComboBox2.Text;
                        }

                        if (X1)
                        {
                            var widths = new global::System.Double[c + 1];
                            for (global::System.Int32 j = 1, loopTo2 = c; j <= loopTo2; j++)
                                widths[(j) - (1)] = Conversions.ToDouble(Operators.DivideObject((Operators.MultiplyObject(this.MaxOfColumn((global::Microsoft.Office.Interop.Excel.Range)displayRng.Columns[(global::System.Object)j]), BaseWidth)), 10));

                            var Values = new global::System.String[1];
                            var ForFormats = new global::System.Int32[1];

                            global::System.Int32 Index = -(1);
                            global::System.Int32 position;

                            for (global::System.Int32 i = 1, loopTo3 = r; i <= loopTo3; i++)
                            {
                                global::System.String source = Conversions.ToString(displayRng.Cells[(global::System.Object)i, (global::System.Object)SplitColumn].value);
                                position = 1;
                                for (global::System.Int32 k = 1, loopTo4 = global::Microsoft.VisualBasic.Strings.Len(source); k <= loopTo4; k++)
                                {
                                    if (((global::Microsoft.VisualBasic.Strings.Mid(source, k, global::Microsoft.VisualBasic.Strings.Len(Separator)) ?? "") == (Separator ?? "")))
                                    {
                                        if ((((k) - (position)) > (0)))
                                        {
                                            Index = ((Index) + (1));
                                            ordinate = 0d;
                                            for (global::System.Int32 j = 1, loopTo5 = (SplitColumn) - (1); j <= loopTo5; j++)
                                            {
                                                var label1 = new global::System.Windows.Forms.Label();
                                                label1.Text = Conversions.ToString(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Value);
                                                label1.Location = new global::System.Drawing.Point((global::System.Int32)Math.Round(ordinate), (global::System.Int32)Math.Round((global::System.Double)(Index) * (Height)));
                                                label1.Height = (global::System.Int32)Math.Round(Height);
                                                label1.Width = (global::System.Int32)Math.Round(widths[(j) - (1)]);
                                                label1.BorderStyle = global::System.Windows.Forms.BorderStyle.FixedSingle;
                                                label1.TextAlign = global::System.Drawing.ContentAlignment.MiddleCenter;

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

                                                    label1.Font = new global::System.Drawing.Font(font.ToString(), fontSize, fontStyle);
                                                    if (Conversions.ToBoolean(!(Operators.ConditionalCompareObjectEqual(cell.Interior.ColorIndex, global::Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexNone, false))))
                                                    {
                                                        global::System.Int64 colorValue1 = Conversions.ToLong(cell.Interior.Color);
                                                        global::System.Int32 red1 = (global::System.Int32)((colorValue1) % (256L));
                                                        global::System.Int32 green1 = (global::System.Int32)(((((colorValue1) / (256L)))) % (256L));
                                                        global::System.Int32 blue1 = (global::System.Int32)((((((colorValue1) / (256L)) / (256L)))) % (256L));
                                                        label1.BackColor = global::System.Drawing.Color.FromArgb(red1, green1, blue1);
                                                    }

                                                    if ((cell.Font.Color is System.DBNull))
                                                    {
                                                        label1.ForeColor = global::System.Drawing.Color.FromArgb(0, 0, 0);
                                                    }

                                                    else if (Conversions.ToBoolean(!(Operators.ConditionalCompareObjectEqual(cell.Font.ColorIndex, global::Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexNone, false))))
                                                    {
                                                        global::System.Int64 colorValue2 = Conversions.ToLong(cell.Font.Color);
                                                        global::System.Int32 red2 = (global::System.Int32)((colorValue2) % (256L));
                                                        global::System.Int32 green2 = (global::System.Int32)(((((colorValue2) / (256L)))) % (256L));
                                                        global::System.Int32 blue2 = (global::System.Int32)((((((colorValue2) / (256L)) / (256L)))) % (256L));
                                                        label1.ForeColor = global::System.Drawing.Color.FromArgb(red2, green2, blue2);
                                                    }
                                                }

                                                this.CustomPanel2.Controls.Add(label1);

                                                ordinate = ((ordinate) + (widths[(j) - (1)]));
                                            }
                                            Array.Resize(ref Values, Index + 1);
                                            Array.Resize(ref ForFormats, Index + 1);
                                            Values[Index] = global::Microsoft.VisualBasic.Strings.Mid(source, position, (k) - (position));
                                            ForFormats[Index] = i;
                                        }
                                        position = ((k) + (global::Microsoft.VisualBasic.Strings.Len(Separator)));
                                    }
                                }
                                if (((position) <= (global::Microsoft.VisualBasic.Strings.Len(source))))
                                {
                                    Index = ((Index) + (1));
                                    ordinate = 0d;
                                    for (global::System.Int32 j = 1, loopTo6 = (SplitColumn) - (1); j <= loopTo6; j++)
                                    {
                                        var label1 = new global::System.Windows.Forms.Label();
                                        label1.Text = Conversions.ToString(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Value);
                                        label1.Location = new global::System.Drawing.Point((global::System.Int32)Math.Round(ordinate), (global::System.Int32)Math.Round((global::System.Double)(Index) * (Height)));
                                        label1.Height = (global::System.Int32)Math.Round(Height);
                                        label1.Width = (global::System.Int32)Math.Round(widths[(j) - (1)]);
                                        label1.BorderStyle = global::System.Windows.Forms.BorderStyle.FixedSingle;
                                        label1.TextAlign = global::System.Drawing.ContentAlignment.MiddleCenter;


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

                                            label1.Font = new global::System.Drawing.Font(font.ToString(), fontSize, fontStyle);
                                            if (Conversions.ToBoolean(!(Operators.ConditionalCompareObjectEqual(cell.Interior.ColorIndex, global::Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexNone, false))))
                                            {
                                                global::System.Int64 colorValue1 = Conversions.ToLong(cell.Interior.Color);
                                                global::System.Int32 red1 = (global::System.Int32)((colorValue1) % (256L));
                                                global::System.Int32 green1 = (global::System.Int32)(((((colorValue1) / (256L)))) % (256L));
                                                global::System.Int32 blue1 = (global::System.Int32)((((((colorValue1) / (256L)) / (256L)))) % (256L));
                                                label1.BackColor = global::System.Drawing.Color.FromArgb(red1, green1, blue1);
                                            }

                                            if ((cell.Font.Color is System.DBNull))
                                            {
                                                label1.ForeColor = global::System.Drawing.Color.FromArgb(0, 0, 0);
                                            }

                                            else if (Conversions.ToBoolean(!(Operators.ConditionalCompareObjectEqual(cell.Font.ColorIndex, global::Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexNone, false))))
                                            {
                                                global::System.Int64 colorValue2 = Conversions.ToLong(cell.Font.Color);
                                                global::System.Int32 red2 = (global::System.Int32)((colorValue2) % (256L));
                                                global::System.Int32 green2 = (global::System.Int32)(((((colorValue2) / (256L)))) % (256L));
                                                global::System.Int32 blue2 = (global::System.Int32)((((((colorValue2) / (256L)) / (256L)))) % (256L));
                                                label1.ForeColor = global::System.Drawing.Color.FromArgb(red2, green2, blue2);
                                            }
                                        }

                                        this.CustomPanel2.Controls.Add(label1);
                                        ordinate = ((ordinate) + (widths[(j) - (1)]));
                                    }
                                    Array.Resize(ref Values, Index + 1);
                                    Array.Resize(ref ForFormats, Index + 1);
                                    Values[Index] = global::Microsoft.VisualBasic.Strings.Mid(source, position, ((global::Microsoft.VisualBasic.Strings.Len(source)) - (position)) + (1));
                                    ForFormats[Index] = i;
                                }
                            }

                            global::System.Int32 SplitOrdinate;
                            SplitOrdinate = (global::System.Int32)Math.Round(ordinate);
                            Width = Conversions.ToDouble(Operators.DivideObject((Operators.MultiplyObject(this.MaxOfArray(Values), BaseWidth)), 10));

                            for (global::System.Int32 m = global::Microsoft.VisualBasic.Information.LBound(Values), loopTo7 = global::Microsoft.VisualBasic.Information.UBound(Values); m <= loopTo7; m++)
                            {
                                ordinate = (global::System.Double)SplitOrdinate;
                                var label1 = new global::System.Windows.Forms.Label();
                                label1.Text = Values[m];
                                label1.Location = new global::System.Drawing.Point((global::System.Int32)Math.Round(ordinate), (global::System.Int32)Math.Round((global::System.Double)(m) * (Height)));
                                label1.Height = (global::System.Int32)Math.Round(Height);
                                label1.Width = (global::System.Int32)Math.Round(Width);
                                label1.BorderStyle = global::System.Windows.Forms.BorderStyle.FixedSingle;
                                label1.TextAlign = global::System.Drawing.ContentAlignment.MiddleCenter;


                                if (((this.CheckBox1.Checked) == (true)))
                                {

                                    global::Microsoft.Office.Interop.Excel.Range cell = (global::Microsoft.Office.Interop.Excel.Range)displayRng.Cells[(global::System.Object)ForFormats[m], (global::System.Object)SplitColumn];
                                    var font = cell.Font;

                                    var fontStyle = global::System.Drawing.FontStyle.Regular;
                                    if (Conversions.ToBoolean(cell.Font.Bold))
                                        fontStyle = (fontStyle | global::System.Drawing.FontStyle.Bold);
                                    if (Conversions.ToBoolean(cell.Font.Italic))
                                        fontStyle = (fontStyle | global::System.Drawing.FontStyle.Italic);

                                    global::System.Single fontSize = global::System.Convert.ToSingle(font.Size);

                                    label1.Font = new global::System.Drawing.Font(font.ToString(), fontSize, fontStyle);
                                    if (Conversions.ToBoolean(!(Operators.ConditionalCompareObjectEqual(cell.Interior.ColorIndex, global::Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexNone, false))))
                                    {
                                        global::System.Int64 colorValue1 = Conversions.ToLong(cell.Interior.Color);
                                        global::System.Int32 red1 = (global::System.Int32)((colorValue1) % (256L));
                                        global::System.Int32 green1 = (global::System.Int32)(((((colorValue1) / (256L)))) % (256L));
                                        global::System.Int32 blue1 = (global::System.Int32)((((((colorValue1) / (256L)) / (256L)))) % (256L));
                                        label1.BackColor = global::System.Drawing.Color.FromArgb(red1, green1, blue1);
                                    }

                                    if ((cell.Font.Color is System.DBNull))
                                    {
                                        label1.ForeColor = global::System.Drawing.Color.FromArgb(0, 0, 0);
                                    }

                                    else if (Conversions.ToBoolean(!(Operators.ConditionalCompareObjectEqual(cell.Font.ColorIndex, global::Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexNone, false))))
                                    {
                                        global::System.Int64 colorValue2 = Conversions.ToLong(cell.Font.Color);
                                        global::System.Int32 red2 = (global::System.Int32)((colorValue2) % (256L));
                                        global::System.Int32 green2 = (global::System.Int32)(((((colorValue2) / (256L)))) % (256L));
                                        global::System.Int32 blue2 = (global::System.Int32)((((((colorValue2) / (256L)) / (256L)))) % (256L));
                                        label1.ForeColor = global::System.Drawing.Color.FromArgb(red2, green2, blue2);
                                    }
                                }
                                this.CustomPanel2.Controls.Add(label1);
                                ordinate = ((ordinate) + (Width));

                                for (global::System.Int32 j = (SplitColumn) + (1), loopTo8 = c; j <= loopTo8; j++)
                                {
                                    var label2 = new global::System.Windows.Forms.Label();
                                    label2.Text = Conversions.ToString(displayRng.Cells[(global::System.Object)ForFormats[m], (global::System.Object)j].value);
                                    label2.Location = new global::System.Drawing.Point((global::System.Int32)Math.Round(ordinate), (global::System.Int32)Math.Round((global::System.Double)(m) * (Height)));
                                    label2.Height = (global::System.Int32)Math.Round(Height);
                                    label2.Width = (global::System.Int32)Math.Round(widths[(j) - (1)]);
                                    label2.BorderStyle = global::System.Windows.Forms.BorderStyle.FixedSingle;
                                    label2.TextAlign = global::System.Drawing.ContentAlignment.MiddleCenter;

                                    if (((this.CheckBox1.Checked) == (true)))
                                    {

                                        global::Microsoft.Office.Interop.Excel.Range cell = (global::Microsoft.Office.Interop.Excel.Range)displayRng.Cells[(global::System.Object)ForFormats[m], (global::System.Object)j];
                                        var font = cell.Font;

                                        var fontStyle = global::System.Drawing.FontStyle.Regular;
                                        if (Conversions.ToBoolean(cell.Font.Bold))
                                            fontStyle = (fontStyle | global::System.Drawing.FontStyle.Bold);
                                        if (Conversions.ToBoolean(cell.Font.Italic))
                                            fontStyle = (fontStyle | global::System.Drawing.FontStyle.Italic);

                                        global::System.Single fontSize = global::System.Convert.ToSingle(font.Size);

                                        label2.Font = new global::System.Drawing.Font(font.ToString(), fontSize, fontStyle);
                                        if (Conversions.ToBoolean(!(Operators.ConditionalCompareObjectEqual(cell.Interior.ColorIndex, global::Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexNone, false))))
                                        {
                                            global::System.Int64 colorValue1 = Conversions.ToLong(cell.Interior.Color);
                                            global::System.Int32 red1 = (global::System.Int32)((colorValue1) % (256L));
                                            global::System.Int32 green1 = (global::System.Int32)(((((colorValue1) / (256L)))) % (256L));
                                            global::System.Int32 blue1 = (global::System.Int32)((((((colorValue1) / (256L)) / (256L)))) % (256L));
                                            label2.BackColor = global::System.Drawing.Color.FromArgb(red1, green1, blue1);
                                        }

                                        if ((cell.Font.Color is System.DBNull))
                                        {
                                            label2.ForeColor = global::System.Drawing.Color.FromArgb(0, 0, 0);
                                        }

                                        else if (Conversions.ToBoolean(!(Operators.ConditionalCompareObjectEqual(cell.Font.ColorIndex, global::Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexNone, false))))
                                        {
                                            global::System.Int64 colorValue2 = Conversions.ToLong(cell.Font.Color);
                                            global::System.Int32 red2 = (global::System.Int32)((colorValue2) % (256L));
                                            global::System.Int32 green2 = (global::System.Int32)(((((colorValue2) / (256L)))) % (256L));
                                            global::System.Int32 blue2 = (global::System.Int32)((((((colorValue2) / (256L)) / (256L)))) % (256L));
                                            label2.ForeColor = global::System.Drawing.Color.FromArgb(red2, green2, blue2);
                                        }
                                    }
                                    this.CustomPanel2.Controls.Add(label2);
                                    ordinate = ((ordinate) + (widths[(j) - (1)]));
                                }
                            }
                        }


                        else if (X2)
                        {

                            if (((c) <= (4)))
                            {
                                Height = ((global::System.Double)(this.CustomPanel2.Height) / (global::System.Double)(c));
                            }
                            else
                            {
                                Height = (((119d) / (4d)));
                            }

                            global::System.Int32 position = 1;
                            global::System.Int32 Index;
                            ordinate = 0d;
                            for (global::System.Int32 i = 1, loopTo9 = r; i <= loopTo9; i++)
                            {
                                global::System.String source = Conversions.ToString(displayRng.Cells[(global::System.Object)i, (global::System.Object)SplitColumn].value);
                                var values = new global::System.String[(c)];
                                Index = -(1);
                                for (global::System.Int32 j = 1, loopTo10 = c; j <= loopTo10; j++)
                                {
                                    Index = ((Index) + (1));
                                    values[(j) - (1)] = Conversions.ToString(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].value);
                                }
                                position = 1;
                                for (global::System.Int32 k = 1, loopTo11 = global::Microsoft.VisualBasic.Strings.Len(source); k <= loopTo11; k++)
                                {
                                    if (((global::Microsoft.VisualBasic.Strings.Mid(source, k, global::Microsoft.VisualBasic.Strings.Len(Separator)) ?? "") == (Separator ?? "")))
                                    {
                                        if ((((k) - (position)) > (0)))
                                        {
                                            values[(SplitColumn) - (1)] = global::Microsoft.VisualBasic.Strings.Mid(source, position, (k) - (position));
                                            Width = Conversions.ToDouble(Operators.DivideObject((Operators.MultiplyObject(this.MaxOfArray(values), BaseWidth)), 10));
                                            for (global::System.Int32 m = global::Microsoft.VisualBasic.Information.LBound(values), loopTo12 = global::Microsoft.VisualBasic.Information.UBound(values); m <= loopTo12; m++)
                                            {
                                                var label1 = new global::System.Windows.Forms.Label();
                                                label1.Text = values[m];
                                                label1.Location = new global::System.Drawing.Point((global::System.Int32)Math.Round(ordinate), (global::System.Int32)Math.Round((global::System.Double)(m) * (Height)));
                                                label1.Height = (global::System.Int32)Math.Round(Height);
                                                label1.Width = (global::System.Int32)Math.Round(Width);
                                                label1.BorderStyle = global::System.Windows.Forms.BorderStyle.FixedSingle;
                                                label1.TextAlign = global::System.Drawing.ContentAlignment.MiddleCenter;

                                                if (((this.CheckBox1.Checked) == (true)))
                                                {

                                                    global::Microsoft.Office.Interop.Excel.Range cell = (global::Microsoft.Office.Interop.Excel.Range)displayRng.Cells[(global::System.Object)i, (global::System.Object)((m) + (1))];
                                                    var font = cell.Font;

                                                    var fontStyle = global::System.Drawing.FontStyle.Regular;
                                                    if (Conversions.ToBoolean(cell.Font.Bold))
                                                        fontStyle = (fontStyle | global::System.Drawing.FontStyle.Bold);
                                                    if (Conversions.ToBoolean(cell.Font.Italic))
                                                        fontStyle = (fontStyle | global::System.Drawing.FontStyle.Italic);

                                                    global::System.Single fontSize = global::System.Convert.ToSingle(font.Size);

                                                    label1.Font = new global::System.Drawing.Font(font.ToString(), fontSize, fontStyle);
                                                    if (Conversions.ToBoolean(!(Operators.ConditionalCompareObjectEqual(cell.Interior.ColorIndex, global::Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexNone, false))))
                                                    {
                                                        global::System.Int64 colorValue1 = Conversions.ToLong(cell.Interior.Color);
                                                        global::System.Int32 red1 = (global::System.Int32)((colorValue1) % (256L));
                                                        global::System.Int32 green1 = (global::System.Int32)(((((colorValue1) / (256L)))) % (256L));
                                                        global::System.Int32 blue1 = (global::System.Int32)((((((colorValue1) / (256L)) / (256L)))) % (256L));
                                                        label1.BackColor = global::System.Drawing.Color.FromArgb(red1, green1, blue1);
                                                    }

                                                    if ((cell.Font.Color is System.DBNull))
                                                    {
                                                        label1.ForeColor = global::System.Drawing.Color.FromArgb(0, 0, 0);
                                                    }

                                                    else if (Conversions.ToBoolean(!(Operators.ConditionalCompareObjectEqual(cell.Font.ColorIndex, global::Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexNone, false))))
                                                    {
                                                        global::System.Int64 colorValue2 = Conversions.ToLong(cell.Font.Color);
                                                        global::System.Int32 red2 = (global::System.Int32)((colorValue2) % (256L));
                                                        global::System.Int32 green2 = (global::System.Int32)(((((colorValue2) / (256L)))) % (256L));
                                                        global::System.Int32 blue2 = (global::System.Int32)((((((colorValue2) / (256L)) / (256L)))) % (256L));
                                                        label1.ForeColor = global::System.Drawing.Color.FromArgb(red2, green2, blue2);
                                                    }
                                                }

                                                this.CustomPanel2.Controls.Add(label1);
                                            }
                                            ordinate = ((ordinate) + (Width));
                                        }
                                        position = ((k) + (global::Microsoft.VisualBasic.Strings.Len(Separator)));
                                    }
                                }
                                if (((position) <= (global::Microsoft.VisualBasic.Strings.Len(source))))
                                {
                                    values[(SplitColumn) - (1)] = global::Microsoft.VisualBasic.Strings.Mid(source, position, ((global::Microsoft.VisualBasic.Strings.Len(source)) - (position)) + (1));
                                    Width = Conversions.ToDouble(Operators.DivideObject((Operators.MultiplyObject(this.MaxOfArray(values), BaseWidth)), 10));
                                    for (global::System.Int32 m = global::Microsoft.VisualBasic.Information.LBound(values), loopTo13 = global::Microsoft.VisualBasic.Information.UBound(values); m <= loopTo13; m++)
                                    {
                                        var label1 = new global::System.Windows.Forms.Label();
                                        label1.Text = values[m];
                                        label1.Location = new global::System.Drawing.Point((global::System.Int32)Math.Round(ordinate), (global::System.Int32)Math.Round((global::System.Double)(m) * (Height)));
                                        label1.Height = (global::System.Int32)Math.Round(Height);
                                        label1.Width = (global::System.Int32)Math.Round(Width);
                                        label1.BorderStyle = global::System.Windows.Forms.BorderStyle.FixedSingle;
                                        label1.TextAlign = global::System.Drawing.ContentAlignment.MiddleCenter;

                                        if (((this.CheckBox1.Checked) == (true)))
                                        {

                                            global::Microsoft.Office.Interop.Excel.Range cell = (global::Microsoft.Office.Interop.Excel.Range)displayRng.Cells[(global::System.Object)i, (global::System.Object)((m) + (1))];
                                            var font = cell.Font;

                                            var fontStyle = global::System.Drawing.FontStyle.Regular;
                                            if (Conversions.ToBoolean(cell.Font.Bold))
                                                fontStyle = (fontStyle | global::System.Drawing.FontStyle.Bold);
                                            if (Conversions.ToBoolean(cell.Font.Italic))
                                                fontStyle = (fontStyle | global::System.Drawing.FontStyle.Italic);

                                            global::System.Single fontSize = global::System.Convert.ToSingle(font.Size);

                                            label1.Font = new global::System.Drawing.Font(font.ToString(), fontSize, fontStyle);
                                            if (Conversions.ToBoolean(!(Operators.ConditionalCompareObjectEqual(cell.Interior.ColorIndex, global::Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexNone, false))))
                                            {
                                                global::System.Int64 colorValue1 = Conversions.ToLong(cell.Interior.Color);
                                                global::System.Int32 red1 = (global::System.Int32)((colorValue1) % (256L));
                                                global::System.Int32 green1 = (global::System.Int32)(((((colorValue1) / (256L)))) % (256L));
                                                global::System.Int32 blue1 = (global::System.Int32)((((((colorValue1) / (256L)) / (256L)))) % (256L));
                                                label1.BackColor = global::System.Drawing.Color.FromArgb(red1, green1, blue1);
                                            }

                                            if ((cell.Font.Color is System.DBNull))
                                            {
                                                label1.ForeColor = global::System.Drawing.Color.FromArgb(0, 0, 0);
                                            }

                                            else if (Conversions.ToBoolean(!(Operators.ConditionalCompareObjectEqual(cell.Font.ColorIndex, global::Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexNone, false))))
                                            {
                                                global::System.Int64 colorValue2 = Conversions.ToLong(cell.Font.Color);
                                                global::System.Int32 red2 = (global::System.Int32)((colorValue2) % (256L));
                                                global::System.Int32 green2 = (global::System.Int32)(((((colorValue2) / (256L)))) % (256L));
                                                global::System.Int32 blue2 = (global::System.Int32)((((((colorValue2) / (256L)) / (256L)))) % (256L));
                                                label1.ForeColor = global::System.Drawing.Color.FromArgb(red2, green2, blue2);
                                            }
                                        }

                                        this.CustomPanel2.Controls.Add(label1);
                                    }
                                    ordinate = ((ordinate) + (Width));
                                }
                            }
                        }
                    }

                    else if (X3)
                    {
                        if (X1)
                        {
                            var widths = new global::System.Double[c + 1];
                            for (global::System.Int32 j = 1, loopTo14 = c; j <= loopTo14; j++)
                                widths[(j) - (1)] = Conversions.ToDouble(Operators.DivideObject((Operators.MultiplyObject(this.MaxOfColumn((global::Microsoft.Office.Interop.Excel.Range)displayRng.Columns[(global::System.Object)j]), BaseWidth)), 10));

                            var Values = new global::System.String[1];
                            global::System.Int32 Index = -(1);

                            for (global::System.Int32 i = 1, loopTo15 = r; i <= loopTo15; i++)
                            {

                                global::System.String source = Conversions.ToString(displayRng.Cells[(global::System.Object)i, (global::System.Object)SplitColumn].value);
                                var NumberText = new global::System.String[2];
                                NumberText = (global::System.String[])this.SeparateNumberText(source);
                                global::System.String Number = NumberText[0];
                                global::System.String Text = NumberText[1];

                                ordinate = 0d;
                                Index = ((Index) + (1));
                                for (global::System.Int32 j = 1, loopTo16 = (SplitColumn) - (1); j <= loopTo16; j++)
                                {
                                    var label1 = new global::System.Windows.Forms.Label();
                                    label1.Text = Conversions.ToString(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Value);
                                    label1.Location = new global::System.Drawing.Point((global::System.Int32)Math.Round(ordinate), (global::System.Int32)Math.Round((global::System.Double)(Index) * (Height)));
                                    label1.Height = (global::System.Int32)Math.Round(Height);
                                    label1.Width = (global::System.Int32)Math.Round(widths[(j) - (1)]);
                                    label1.BorderStyle = global::System.Windows.Forms.BorderStyle.FixedSingle;
                                    label1.TextAlign = global::System.Drawing.ContentAlignment.MiddleCenter;

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

                                        label1.Font = new global::System.Drawing.Font(font.ToString(), fontSize, fontStyle);
                                        if (Conversions.ToBoolean(!(Operators.ConditionalCompareObjectEqual(cell.Interior.ColorIndex, global::Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexNone, false))))
                                        {
                                            global::System.Int64 colorValue1 = Conversions.ToLong(cell.Interior.Color);
                                            global::System.Int32 red1 = (global::System.Int32)((colorValue1) % (256L));
                                            global::System.Int32 green1 = (global::System.Int32)(((((colorValue1) / (256L)))) % (256L));
                                            global::System.Int32 blue1 = (global::System.Int32)((((((colorValue1) / (256L)) / (256L)))) % (256L));
                                            label1.BackColor = global::System.Drawing.Color.FromArgb(red1, green1, blue1);
                                        }

                                        if ((cell.Font.Color is System.DBNull))
                                        {
                                            label1.ForeColor = global::System.Drawing.Color.FromArgb(0, 0, 0);
                                        }

                                        else if (Conversions.ToBoolean(!(Operators.ConditionalCompareObjectEqual(cell.Font.ColorIndex, global::Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexNone, false))))
                                        {
                                            global::System.Int64 colorValue2 = Conversions.ToLong(cell.Font.Color);
                                            global::System.Int32 red2 = (global::System.Int32)((colorValue2) % (256L));
                                            global::System.Int32 green2 = (global::System.Int32)(((((colorValue2) / (256L)))) % (256L));
                                            global::System.Int32 blue2 = (global::System.Int32)((((((colorValue2) / (256L)) / (256L)))) % (256L));
                                            label1.ForeColor = global::System.Drawing.Color.FromArgb(red2, green2, blue2);
                                        }
                                    }

                                    this.CustomPanel2.Controls.Add(label1);
                                    ordinate = ((ordinate) + (widths[(j) - (1)]));
                                }

                                Array.Resize(ref Values, Index + 1);
                                Values[Index] = Number;

                                ordinate = 0d;
                                Index = ((Index) + (1));
                                for (global::System.Int32 j = 1, loopTo17 = (SplitColumn) - (1); j <= loopTo17; j++)
                                {
                                    var label1 = new global::System.Windows.Forms.Label();
                                    label1.Text = Conversions.ToString(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Value);
                                    label1.Location = new global::System.Drawing.Point((global::System.Int32)Math.Round(ordinate), (global::System.Int32)Math.Round((global::System.Double)(Index) * (Height)));
                                    label1.Height = (global::System.Int32)Math.Round(Height);
                                    label1.Width = (global::System.Int32)Math.Round(widths[(j) - (1)]);
                                    label1.BorderStyle = global::System.Windows.Forms.BorderStyle.FixedSingle;
                                    label1.TextAlign = global::System.Drawing.ContentAlignment.MiddleCenter;

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

                                        label1.Font = new global::System.Drawing.Font(font.ToString(), fontSize, fontStyle);
                                        if (Conversions.ToBoolean(!(Operators.ConditionalCompareObjectEqual(cell.Interior.ColorIndex, global::Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexNone, false))))
                                        {
                                            global::System.Int64 colorValue1 = Conversions.ToLong(cell.Interior.Color);
                                            global::System.Int32 red1 = (global::System.Int32)((colorValue1) % (256L));
                                            global::System.Int32 green1 = (global::System.Int32)(((((colorValue1) / (256L)))) % (256L));
                                            global::System.Int32 blue1 = (global::System.Int32)((((((colorValue1) / (256L)) / (256L)))) % (256L));
                                            label1.BackColor = global::System.Drawing.Color.FromArgb(red1, green1, blue1);
                                        }

                                        if ((cell.Font.Color is System.DBNull))
                                        {
                                            label1.ForeColor = global::System.Drawing.Color.FromArgb(0, 0, 0);
                                        }

                                        else if (Conversions.ToBoolean(!(Operators.ConditionalCompareObjectEqual(cell.Font.ColorIndex, global::Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexNone, false))))
                                        {
                                            global::System.Int64 colorValue2 = Conversions.ToLong(cell.Font.Color);
                                            global::System.Int32 red2 = (global::System.Int32)((colorValue2) % (256L));
                                            global::System.Int32 green2 = (global::System.Int32)(((((colorValue2) / (256L)))) % (256L));
                                            global::System.Int32 blue2 = (global::System.Int32)((((((colorValue2) / (256L)) / (256L)))) % (256L));
                                            label1.ForeColor = global::System.Drawing.Color.FromArgb(red2, green2, blue2);
                                        }
                                    }
                                    this.CustomPanel2.Controls.Add(label1);
                                    ordinate = ((ordinate) + (widths[(j) - (1)]));
                                }
                                Array.Resize(ref Values, Index + 1);
                                Values[Index] = Text;
                            }

                            Width = Conversions.ToDouble(Operators.DivideObject((Operators.MultiplyObject(this.MaxOfArray(Values), BaseWidth)), 10));
                            global::System.Double SplitOrdinate;
                            SplitOrdinate = ordinate;

                            for (global::System.Int32 i = global::Microsoft.VisualBasic.Information.LBound(Values), loopTo18 = global::Microsoft.VisualBasic.Information.UBound(Values); i <= loopTo18; i++)
                            {
                                ordinate = SplitOrdinate;
                                var label1 = new global::System.Windows.Forms.Label();
                                label1.Text = Values[i];
                                label1.Location = new global::System.Drawing.Point((global::System.Int32)Math.Round(ordinate), (global::System.Int32)Math.Round((global::System.Double)(i) * (Height)));
                                label1.Height = (global::System.Int32)Math.Round(Height);
                                label1.Width = (global::System.Int32)Math.Round(Width);
                                label1.BorderStyle = global::System.Windows.Forms.BorderStyle.FixedSingle;
                                label1.TextAlign = global::System.Drawing.ContentAlignment.MiddleCenter;

                                if (((this.CheckBox1.Checked) == (true)))
                                {

                                    global::Microsoft.Office.Interop.Excel.Range cell = (global::Microsoft.Office.Interop.Excel.Range)displayRng.Cells[(global::System.Object)((global::Microsoft.VisualBasic.Conversion.Int((global::System.Double)(i) / (2d))) + (1d)), (global::System.Object)SplitColumn];
                                    var font = cell.Font;

                                    var fontStyle = global::System.Drawing.FontStyle.Regular;
                                    if (Conversions.ToBoolean(cell.Font.Bold))
                                        fontStyle = (fontStyle | global::System.Drawing.FontStyle.Bold);
                                    if (Conversions.ToBoolean(cell.Font.Italic))
                                        fontStyle = (fontStyle | global::System.Drawing.FontStyle.Italic);

                                    global::System.Single fontSize = global::System.Convert.ToSingle(font.Size);

                                    label1.Font = new global::System.Drawing.Font(font.ToString(), fontSize, fontStyle);
                                    if (Conversions.ToBoolean(!(Operators.ConditionalCompareObjectEqual(cell.Interior.ColorIndex, global::Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexNone, false))))
                                    {
                                        global::System.Int64 colorValue1 = Conversions.ToLong(cell.Interior.Color);
                                        global::System.Int32 red1 = (global::System.Int32)((colorValue1) % (256L));
                                        global::System.Int32 green1 = (global::System.Int32)(((((colorValue1) / (256L)))) % (256L));
                                        global::System.Int32 blue1 = (global::System.Int32)((((((colorValue1) / (256L)) / (256L)))) % (256L));
                                        label1.BackColor = global::System.Drawing.Color.FromArgb(red1, green1, blue1);
                                    }

                                    if ((cell.Font.Color is System.DBNull))
                                    {
                                        label1.ForeColor = global::System.Drawing.Color.FromArgb(0, 0, 0);
                                    }

                                    else if (Conversions.ToBoolean(!(Operators.ConditionalCompareObjectEqual(cell.Font.ColorIndex, global::Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexNone, false))))
                                    {
                                        global::System.Int64 colorValue2 = Conversions.ToLong(cell.Font.Color);
                                        global::System.Int32 red2 = (global::System.Int32)((colorValue2) % (256L));
                                        global::System.Int32 green2 = (global::System.Int32)(((((colorValue2) / (256L)))) % (256L));
                                        global::System.Int32 blue2 = (global::System.Int32)((((((colorValue2) / (256L)) / (256L)))) % (256L));
                                        label1.ForeColor = global::System.Drawing.Color.FromArgb(red2, green2, blue2);
                                    }
                                }

                                this.CustomPanel2.Controls.Add(label1);
                                ordinate = ((ordinate) + (Width));

                                for (global::System.Int32 j = (SplitColumn) + (1), loopTo19 = c; j <= loopTo19; j++)
                                {
                                    var label2 = new global::System.Windows.Forms.Label();
                                    label2.Text = Conversions.ToString(displayRng.Cells[(global::System.Object)((global::Microsoft.VisualBasic.Conversion.Int((global::System.Double)(i) / (2d))) + (1d)), (global::System.Object)j].value);
                                    label2.Location = new global::System.Drawing.Point((global::System.Int32)Math.Round(ordinate), (global::System.Int32)Math.Round((global::System.Double)(i) * (Height)));
                                    label2.Height = (global::System.Int32)Math.Round(Height);
                                    label2.Width = (global::System.Int32)Math.Round(widths[(j) - (1)]);
                                    label2.BorderStyle = global::System.Windows.Forms.BorderStyle.FixedSingle;
                                    label2.TextAlign = global::System.Drawing.ContentAlignment.MiddleCenter;

                                    if (((this.CheckBox1.Checked) == (true)))
                                    {

                                        global::Microsoft.Office.Interop.Excel.Range cell = (global::Microsoft.Office.Interop.Excel.Range)displayRng.Cells[(global::System.Object)((global::Microsoft.VisualBasic.Conversion.Int((global::System.Double)(i) / (2d))) + (1d)), (global::System.Object)c];
                                        var font = cell.Font;

                                        var fontStyle = global::System.Drawing.FontStyle.Regular;
                                        if (Conversions.ToBoolean(cell.Font.Bold))
                                            fontStyle = (fontStyle | global::System.Drawing.FontStyle.Bold);
                                        if (Conversions.ToBoolean(cell.Font.Italic))
                                            fontStyle = (fontStyle | global::System.Drawing.FontStyle.Italic);

                                        global::System.Single fontSize = global::System.Convert.ToSingle(font.Size);

                                        label2.Font = new global::System.Drawing.Font(font.ToString(), fontSize, fontStyle);
                                        if (Conversions.ToBoolean(!(Operators.ConditionalCompareObjectEqual(cell.Interior.ColorIndex, global::Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexNone, false))))
                                        {
                                            global::System.Int64 colorValue1 = Conversions.ToLong(cell.Interior.Color);
                                            global::System.Int32 red1 = (global::System.Int32)((colorValue1) % (256L));
                                            global::System.Int32 green1 = (global::System.Int32)(((((colorValue1) / (256L)))) % (256L));
                                            global::System.Int32 blue1 = (global::System.Int32)((((((colorValue1) / (256L)) / (256L)))) % (256L));
                                            label2.BackColor = global::System.Drawing.Color.FromArgb(red1, green1, blue1);
                                        }

                                        if ((cell.Font.Color is System.DBNull))
                                        {
                                            label2.ForeColor = global::System.Drawing.Color.FromArgb(0, 0, 0);
                                        }

                                        else if (Conversions.ToBoolean(!(Operators.ConditionalCompareObjectEqual(cell.Font.ColorIndex, global::Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexNone, false))))
                                        {
                                            global::System.Int64 colorValue2 = Conversions.ToLong(cell.Font.Color);
                                            global::System.Int32 red2 = (global::System.Int32)((colorValue2) % (256L));
                                            global::System.Int32 green2 = (global::System.Int32)(((((colorValue2) / (256L)))) % (256L));
                                            global::System.Int32 blue2 = (global::System.Int32)((((((colorValue2) / (256L)) / (256L)))) % (256L));
                                            label2.ForeColor = global::System.Drawing.Color.FromArgb(red2, green2, blue2);
                                        }
                                    }
                                    this.CustomPanel2.Controls.Add(label2);
                                    ordinate = ((ordinate) + (widths[(j) - (1)]));
                                }
                            }
                        }

                        else if (X2)
                        {

                            if (((c) <= (4)))
                            {
                                Height = ((global::System.Double)(this.CustomPanel2.Height) / (global::System.Double)(c));
                            }
                            else
                            {
                                Height = (((119d) / (4d)));
                            }

                            global::System.Int32 position = 1;
                            global::System.Int32 Index;
                            ordinate = 0d;
                            for (global::System.Int32 i = 1, loopTo20 = r; i <= loopTo20; i++)
                            {
                                global::System.String source = Conversions.ToString(displayRng.Cells[(global::System.Object)i, (global::System.Object)SplitColumn].value);
                                var NumberText = new global::System.String[2];
                                NumberText = (global::System.String[])this.SeparateNumberText(source);
                                global::System.String Number = NumberText[0];
                                global::System.String Text = NumberText[1];

                                var values = new global::System.String[(c)];
                                Index = -(1);
                                for (global::System.Int32 j = 1, loopTo21 = (c) - (1); j <= loopTo21; j++)
                                {
                                    Index = ((Index) + (1));
                                    values[(j) - (1)] = Conversions.ToString(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].value);
                                }
                                values[(SplitColumn) - (1)] = Number;
                                Width = Conversions.ToDouble(Operators.DivideObject((Operators.MultiplyObject(this.MaxOfArray(values), BaseWidth)), 10));
                                for (global::System.Int32 m = global::Microsoft.VisualBasic.Information.LBound(values), loopTo22 = global::Microsoft.VisualBasic.Information.UBound(values); m <= loopTo22; m++)
                                {
                                    var label1 = new global::System.Windows.Forms.Label();
                                    label1.Text = values[m];
                                    label1.Location = new global::System.Drawing.Point((global::System.Int32)Math.Round(ordinate), (global::System.Int32)Math.Round((global::System.Double)(m) * (Height)));
                                    label1.Height = (global::System.Int32)Math.Round(Height);
                                    label1.Width = (global::System.Int32)Math.Round(Width);
                                    label1.BorderStyle = global::System.Windows.Forms.BorderStyle.FixedSingle;
                                    label1.TextAlign = global::System.Drawing.ContentAlignment.MiddleCenter;

                                    if (((this.CheckBox1.Checked) == (true)))
                                    {

                                        global::Microsoft.Office.Interop.Excel.Range cell = (global::Microsoft.Office.Interop.Excel.Range)displayRng.Cells[(global::System.Object)i, (global::System.Object)((m) + (1))];
                                        var font = cell.Font;

                                        var fontStyle = global::System.Drawing.FontStyle.Regular;
                                        if (Conversions.ToBoolean(cell.Font.Bold))
                                            fontStyle = (fontStyle | global::System.Drawing.FontStyle.Bold);
                                        if (Conversions.ToBoolean(cell.Font.Italic))
                                            fontStyle = (fontStyle | global::System.Drawing.FontStyle.Italic);

                                        global::System.Single fontSize = global::System.Convert.ToSingle(font.Size);

                                        label1.Font = new global::System.Drawing.Font(font.ToString(), fontSize, fontStyle);
                                        if (Conversions.ToBoolean(!(Operators.ConditionalCompareObjectEqual(cell.Interior.ColorIndex, global::Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexNone, false))))
                                        {
                                            global::System.Int64 colorValue1 = Conversions.ToLong(cell.Interior.Color);
                                            global::System.Int32 red1 = (global::System.Int32)((colorValue1) % (256L));
                                            global::System.Int32 green1 = (global::System.Int32)(((((colorValue1) / (256L)))) % (256L));
                                            global::System.Int32 blue1 = (global::System.Int32)((((((colorValue1) / (256L)) / (256L)))) % (256L));
                                            label1.BackColor = global::System.Drawing.Color.FromArgb(red1, green1, blue1);
                                        }

                                        if ((cell.Font.Color is System.DBNull))
                                        {
                                            label1.ForeColor = global::System.Drawing.Color.FromArgb(0, 0, 0);
                                        }

                                        else if (Conversions.ToBoolean(!(Operators.ConditionalCompareObjectEqual(cell.Font.ColorIndex, global::Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexNone, false))))
                                        {
                                            global::System.Int64 colorValue2 = Conversions.ToLong(cell.Font.Color);
                                            global::System.Int32 red2 = (global::System.Int32)((colorValue2) % (256L));
                                            global::System.Int32 green2 = (global::System.Int32)(((((colorValue2) / (256L)))) % (256L));
                                            global::System.Int32 blue2 = (global::System.Int32)((((((colorValue2) / (256L)) / (256L)))) % (256L));
                                            label1.ForeColor = global::System.Drawing.Color.FromArgb(red2, green2, blue2);
                                        }
                                    }
                                    this.CustomPanel2.Controls.Add(label1);
                                }
                                ordinate = ((ordinate) + (Width));

                                values[(SplitColumn) - (1)] = Text;
                                Width = Conversions.ToDouble(Operators.DivideObject((Operators.MultiplyObject(this.MaxOfArray(values), BaseWidth)), 10));
                                for (global::System.Int32 m = global::Microsoft.VisualBasic.Information.LBound(values), loopTo23 = global::Microsoft.VisualBasic.Information.UBound(values); m <= loopTo23; m++)
                                {
                                    var label1 = new global::System.Windows.Forms.Label();
                                    label1.Text = values[m];
                                    label1.Location = new global::System.Drawing.Point((global::System.Int32)Math.Round(ordinate), (global::System.Int32)Math.Round((global::System.Double)(m) * (Height)));
                                    label1.Height = (global::System.Int32)Math.Round(Height);
                                    label1.Width = (global::System.Int32)Math.Round(Width);
                                    label1.BorderStyle = global::System.Windows.Forms.BorderStyle.FixedSingle;
                                    label1.TextAlign = global::System.Drawing.ContentAlignment.MiddleCenter;

                                    if (((this.CheckBox1.Checked) == (true)))
                                    {

                                        global::Microsoft.Office.Interop.Excel.Range cell = (global::Microsoft.Office.Interop.Excel.Range)displayRng.Cells[(global::System.Object)i, (global::System.Object)((m) + (1))];
                                        var font = cell.Font;

                                        var fontStyle = global::System.Drawing.FontStyle.Regular;
                                        if (Conversions.ToBoolean(cell.Font.Bold))
                                            fontStyle = (fontStyle | global::System.Drawing.FontStyle.Bold);
                                        if (Conversions.ToBoolean(cell.Font.Italic))
                                            fontStyle = (fontStyle | global::System.Drawing.FontStyle.Italic);

                                        global::System.Single fontSize = global::System.Convert.ToSingle(font.Size);

                                        label1.Font = new global::System.Drawing.Font(font.ToString(), fontSize, fontStyle);
                                        if (Conversions.ToBoolean(!(Operators.ConditionalCompareObjectEqual(cell.Interior.ColorIndex, global::Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexNone, false))))
                                        {
                                            global::System.Int64 colorValue1 = Conversions.ToLong(cell.Interior.Color);
                                            global::System.Int32 red1 = (global::System.Int32)((colorValue1) % (256L));
                                            global::System.Int32 green1 = (global::System.Int32)(((((colorValue1) / (256L)))) % (256L));
                                            global::System.Int32 blue1 = (global::System.Int32)((((((colorValue1) / (256L)) / (256L)))) % (256L));
                                            label1.BackColor = global::System.Drawing.Color.FromArgb(red1, green1, blue1);
                                        }

                                        if ((cell.Font.Color is System.DBNull))
                                        {
                                            label1.ForeColor = global::System.Drawing.Color.FromArgb(0, 0, 0);
                                        }

                                        else if (Conversions.ToBoolean(!(Operators.ConditionalCompareObjectEqual(cell.Font.ColorIndex, global::Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexNone, false))))
                                        {
                                            global::System.Int64 colorValue2 = Conversions.ToLong(cell.Font.Color);
                                            global::System.Int32 red2 = (global::System.Int32)((colorValue2) % (256L));
                                            global::System.Int32 green2 = (global::System.Int32)(((((colorValue2) / (256L)))) % (256L));
                                            global::System.Int32 blue2 = (global::System.Int32)((((((colorValue2) / (256L)) / (256L)))) % (256L));
                                            label1.ForeColor = global::System.Drawing.Color.FromArgb(red2, green2, blue2);
                                        }
                                    }
                                    this.CustomPanel2.Controls.Add(label1);
                                }
                                ordinate = ((ordinate) + (Width));

                            }
                        }
                    }

                    else if (X11)
                    {

                        global::System.Int32 W;

                        if (string.IsNullOrEmpty(this.TextBox3.Text))
                        {
                            W = 1;
                        }
                        else
                        {
                            W = Conversions.ToInteger(global::Microsoft.VisualBasic.Conversion.Int(this.TextBox3.Text));
                        }

                        if (X1)
                        {
                            var widths = new global::System.Double[c + 1];
                            for (global::System.Int32 j = 1, loopTo24 = c; j <= loopTo24; j++)
                                widths[(j) - (1)] = Conversions.ToDouble(Operators.DivideObject((Operators.MultiplyObject(this.MaxOfColumn((global::Microsoft.Office.Interop.Excel.Range)displayRng.Columns[(global::System.Object)j]), BaseWidth)), 10));

                            var Values = new global::System.String[1];
                            var ForFormats = new global::System.String[1];
                            global::System.Int32 Index = -(1);

                            for (global::System.Int32 i = 1, loopTo25 = r; i <= loopTo25; i++)
                            {
                                global::System.String source = Conversions.ToString(displayRng.Cells[(global::System.Object)i, (global::System.Object)SplitColumn].value);
                                for (global::System.Double k = 1d, loopTo26 = global::Microsoft.VisualBasic.Conversion.Int((global::System.Double)(global::Microsoft.VisualBasic.Strings.Len(source)) / (global::System.Double)(W)); k <= loopTo26; k++)
                                {
                                    Index = ((Index) + (1));
                                    ordinate = 0d;
                                    for (global::System.Int32 j = 1, loopTo27 = (SplitColumn) - (1); j <= loopTo27; j++)
                                    {
                                        var label1 = new global::System.Windows.Forms.Label();
                                        label1.Text = Conversions.ToString(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Value);
                                        label1.Location = new global::System.Drawing.Point((global::System.Int32)Math.Round(ordinate), (global::System.Int32)Math.Round((global::System.Double)(Index) * (Height)));
                                        label1.Height = (global::System.Int32)Math.Round(Height);
                                        label1.Width = (global::System.Int32)Math.Round(widths[(j) - (1)]);
                                        label1.BorderStyle = global::System.Windows.Forms.BorderStyle.FixedSingle;
                                        label1.TextAlign = global::System.Drawing.ContentAlignment.MiddleCenter;

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

                                            label1.Font = new global::System.Drawing.Font(font.ToString(), fontSize, fontStyle);
                                            if (Conversions.ToBoolean(!(Operators.ConditionalCompareObjectEqual(cell.Interior.ColorIndex, global::Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexNone, false))))
                                            {
                                                global::System.Int64 colorValue1 = Conversions.ToLong(cell.Interior.Color);
                                                global::System.Int32 red1 = (global::System.Int32)((colorValue1) % (256L));
                                                global::System.Int32 green1 = (global::System.Int32)(((((colorValue1) / (256L)))) % (256L));
                                                global::System.Int32 blue1 = (global::System.Int32)((((((colorValue1) / (256L)) / (256L)))) % (256L));
                                                label1.BackColor = global::System.Drawing.Color.FromArgb(red1, green1, blue1);
                                            }

                                            if ((cell.Font.Color is System.DBNull))
                                            {
                                                label1.ForeColor = global::System.Drawing.Color.FromArgb(0, 0, 0);
                                            }

                                            else if (Conversions.ToBoolean(!(Operators.ConditionalCompareObjectEqual(cell.Font.ColorIndex, global::Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexNone, false))))
                                            {
                                                global::System.Int64 colorValue2 = Conversions.ToLong(cell.Font.Color);
                                                global::System.Int32 red2 = (global::System.Int32)((colorValue2) % (256L));
                                                global::System.Int32 green2 = (global::System.Int32)(((((colorValue2) / (256L)))) % (256L));
                                                global::System.Int32 blue2 = (global::System.Int32)((((((colorValue2) / (256L)) / (256L)))) % (256L));
                                                label1.ForeColor = global::System.Drawing.Color.FromArgb(red2, green2, blue2);
                                            }
                                        }
                                        this.CustomPanel2.Controls.Add(label1);
                                        ordinate = ((ordinate) + (widths[(j) - (1)]));
                                    }
                                    Array.Resize(ref Values, Index + 1);
                                    Array.Resize(ref ForFormats, Index + 1);
                                    Values[Index] = global::Microsoft.VisualBasic.Strings.Mid(source, (global::System.Int32)Math.Round(((((global::System.Double)(W) * ((((k) - (1d))))))) + (1d)), W);
                                    ForFormats[Index] = (i).ToString();
                                }
                                if ((((global::Microsoft.VisualBasic.Strings.Len(source)) % (W)) != (0)))
                                {
                                    Index = ((Index) + (1));
                                    ordinate = 0d;
                                    for (global::System.Int32 j = 1, loopTo28 = (SplitColumn) - (1); j <= loopTo28; j++)
                                    {
                                        var label1 = new global::System.Windows.Forms.Label();
                                        label1.Text = Conversions.ToString(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].Value);
                                        label1.Location = new global::System.Drawing.Point((global::System.Int32)Math.Round(ordinate), (global::System.Int32)Math.Round((global::System.Double)(Index) * (Height)));
                                        label1.Height = (global::System.Int32)Math.Round(Height);
                                        label1.Width = (global::System.Int32)Math.Round(widths[(j) - (1)]);
                                        label1.BorderStyle = global::System.Windows.Forms.BorderStyle.FixedSingle;
                                        label1.TextAlign = global::System.Drawing.ContentAlignment.MiddleCenter;

                                        if (((this.CheckBox2.Checked) == (true)))
                                        {

                                            global::Microsoft.Office.Interop.Excel.Range cell = (global::Microsoft.Office.Interop.Excel.Range)displayRng.Cells[(global::System.Object)i, (global::System.Object)j];
                                            var font = cell.Font;

                                            var fontStyle = global::System.Drawing.FontStyle.Regular;
                                            if (Conversions.ToBoolean(cell.Font.Bold))
                                                fontStyle = (fontStyle | global::System.Drawing.FontStyle.Bold);
                                            if (Conversions.ToBoolean(cell.Font.Italic))
                                                fontStyle = (fontStyle | global::System.Drawing.FontStyle.Italic);

                                            global::System.Single fontSize = global::System.Convert.ToSingle(font.Size);

                                            label1.Font = new global::System.Drawing.Font(font.ToString(), fontSize, fontStyle);
                                            if (Conversions.ToBoolean(!(Operators.ConditionalCompareObjectEqual(cell.Interior.ColorIndex, global::Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexNone, false))))
                                            {
                                                global::System.Int64 colorValue1 = Conversions.ToLong(cell.Interior.Color);
                                                global::System.Int32 red1 = (global::System.Int32)((colorValue1) % (256L));
                                                global::System.Int32 green1 = (global::System.Int32)(((((colorValue1) / (256L)))) % (256L));
                                                global::System.Int32 blue1 = (global::System.Int32)((((((colorValue1) / (256L)) / (256L)))) % (256L));
                                                label1.BackColor = global::System.Drawing.Color.FromArgb(red1, green1, blue1);
                                            }

                                            if ((cell.Font.Color is System.DBNull))
                                            {
                                                label1.ForeColor = global::System.Drawing.Color.FromArgb(0, 0, 0);
                                            }

                                            else if (Conversions.ToBoolean(!(Operators.ConditionalCompareObjectEqual(cell.Font.ColorIndex, global::Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexNone, false))))
                                            {
                                                global::System.Int64 colorValue2 = Conversions.ToLong(cell.Font.Color);
                                                global::System.Int32 red2 = (global::System.Int32)((colorValue2) % (256L));
                                                global::System.Int32 green2 = (global::System.Int32)(((((colorValue2) / (256L)))) % (256L));
                                                global::System.Int32 blue2 = (global::System.Int32)((((((colorValue2) / (256L)) / (256L)))) % (256L));
                                                label1.ForeColor = global::System.Drawing.Color.FromArgb(red2, green2, blue2);
                                            }
                                        }
                                        this.CustomPanel2.Controls.Add(label1);
                                        ordinate = ((ordinate) + (widths[(j) - (1)]));
                                    }
                                    Array.Resize(ref Values, Index + 1);
                                    Array.Resize(ref ForFormats, Index + 1);
                                    ForFormats[Index] = (i).ToString();
                                    Values[Index] = global::Microsoft.VisualBasic.Strings.Mid(source, ((global::Microsoft.VisualBasic.Strings.Len(source)) - ((((global::Microsoft.VisualBasic.Strings.Len(source)) % (W))))) + (1), (global::Microsoft.VisualBasic.Strings.Len(source)) % (W));
                                }
                            }

                            Width = Conversions.ToDouble(Operators.DivideObject((Operators.MultiplyObject(this.MaxOfArray(Values), BaseWidth)), 10));
                            global::System.Double SplitOrdinate;
                            SplitOrdinate = ordinate;

                            for (global::System.Int32 i = global::Microsoft.VisualBasic.Information.LBound(Values), loopTo29 = global::Microsoft.VisualBasic.Information.UBound(Values); i <= loopTo29; i++)
                            {
                                ordinate = SplitOrdinate;
                                var label1 = new global::System.Windows.Forms.Label();
                                label1.Text = Values[i];
                                label1.Location = new global::System.Drawing.Point((global::System.Int32)Math.Round(ordinate), (global::System.Int32)Math.Round((global::System.Double)(i) * (Height)));
                                label1.Height = (global::System.Int32)Math.Round(Height);
                                label1.Width = (global::System.Int32)Math.Round(Width);
                                label1.BorderStyle = global::System.Windows.Forms.BorderStyle.FixedSingle;
                                label1.TextAlign = global::System.Drawing.ContentAlignment.MiddleCenter;

                                if (((this.CheckBox1.Checked) == (true)))
                                {

                                    global::Microsoft.Office.Interop.Excel.Range cell = (global::Microsoft.Office.Interop.Excel.Range)displayRng.Cells[ForFormats[i], (global::System.Object)SplitColumn];
                                    var font = cell.Font;

                                    var fontStyle = global::System.Drawing.FontStyle.Regular;
                                    if (Conversions.ToBoolean(cell.Font.Bold))
                                        fontStyle = (fontStyle | global::System.Drawing.FontStyle.Bold);
                                    if (Conversions.ToBoolean(cell.Font.Italic))
                                        fontStyle = (fontStyle | global::System.Drawing.FontStyle.Italic);

                                    global::System.Single fontSize = global::System.Convert.ToSingle(font.Size);

                                    label1.Font = new global::System.Drawing.Font(font.ToString(), fontSize, fontStyle);
                                    if (Conversions.ToBoolean(!(Operators.ConditionalCompareObjectEqual(cell.Interior.ColorIndex, global::Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexNone, false))))
                                    {
                                        global::System.Int64 colorValue1 = Conversions.ToLong(cell.Interior.Color);
                                        global::System.Int32 red1 = (global::System.Int32)((colorValue1) % (256L));
                                        global::System.Int32 green1 = (global::System.Int32)(((((colorValue1) / (256L)))) % (256L));
                                        global::System.Int32 blue1 = (global::System.Int32)((((((colorValue1) / (256L)) / (256L)))) % (256L));
                                        label1.BackColor = global::System.Drawing.Color.FromArgb(red1, green1, blue1);
                                    }

                                    if ((cell.Font.Color is System.DBNull))
                                    {
                                        label1.ForeColor = global::System.Drawing.Color.FromArgb(0, 0, 0);
                                    }

                                    else if (Conversions.ToBoolean(!(Operators.ConditionalCompareObjectEqual(cell.Font.ColorIndex, global::Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexNone, false))))
                                    {
                                        global::System.Int64 colorValue2 = Conversions.ToLong(cell.Font.Color);
                                        global::System.Int32 red2 = (global::System.Int32)((colorValue2) % (256L));
                                        global::System.Int32 green2 = (global::System.Int32)(((((colorValue2) / (256L)))) % (256L));
                                        global::System.Int32 blue2 = (global::System.Int32)((((((colorValue2) / (256L)) / (256L)))) % (256L));
                                        label1.ForeColor = global::System.Drawing.Color.FromArgb(red2, green2, blue2);
                                    }
                                }
                                this.CustomPanel2.Controls.Add(label1);
                                ordinate = ((ordinate) + (Width));

                                for (global::System.Int32 j = (SplitColumn) + (1), loopTo30 = c; j <= loopTo30; j++)
                                {
                                    var label2 = new global::System.Windows.Forms.Label();
                                    label2.Text = Conversions.ToString(displayRng.Cells[(global::System.Object)((global::Microsoft.VisualBasic.Conversion.Int((global::System.Double)(i) / (2d))) + (1d)), (global::System.Object)j].value);
                                    label2.Location = new global::System.Drawing.Point((global::System.Int32)Math.Round(ordinate), (global::System.Int32)Math.Round((global::System.Double)(i) * (Height)));
                                    label2.Height = (global::System.Int32)Math.Round(Height);
                                    label2.Width = (global::System.Int32)Math.Round(widths[(j) - (1)]);
                                    label2.BorderStyle = global::System.Windows.Forms.BorderStyle.FixedSingle;
                                    label2.TextAlign = global::System.Drawing.ContentAlignment.MiddleCenter;

                                    if (((this.CheckBox1.Checked) == (true)))
                                    {

                                        global::Microsoft.Office.Interop.Excel.Range cell = (global::Microsoft.Office.Interop.Excel.Range)displayRng.Cells[(global::System.Object)((global::Microsoft.VisualBasic.Conversion.Int((global::System.Double)(i) / (2d))) + (1d)), (global::System.Object)c];
                                        var font = cell.Font;

                                        var fontStyle = global::System.Drawing.FontStyle.Regular;
                                        if (Conversions.ToBoolean(cell.Font.Bold))
                                            fontStyle = (fontStyle | global::System.Drawing.FontStyle.Bold);
                                        if (Conversions.ToBoolean(cell.Font.Italic))
                                            fontStyle = (fontStyle | global::System.Drawing.FontStyle.Italic);

                                        global::System.Single fontSize = global::System.Convert.ToSingle(font.Size);

                                        label2.Font = new global::System.Drawing.Font(font.ToString(), fontSize, fontStyle);
                                        if (Conversions.ToBoolean(!(Operators.ConditionalCompareObjectEqual(cell.Interior.ColorIndex, global::Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexNone, false))))
                                        {
                                            global::System.Int64 colorValue1 = Conversions.ToLong(cell.Interior.Color);
                                            global::System.Int32 red1 = (global::System.Int32)((colorValue1) % (256L));
                                            global::System.Int32 green1 = (global::System.Int32)(((((colorValue1) / (256L)))) % (256L));
                                            global::System.Int32 blue1 = (global::System.Int32)((((((colorValue1) / (256L)) / (256L)))) % (256L));
                                            label2.BackColor = global::System.Drawing.Color.FromArgb(red1, green1, blue1);
                                        }

                                        if ((cell.Font.Color is System.DBNull))
                                        {
                                            label2.ForeColor = global::System.Drawing.Color.FromArgb(0, 0, 0);
                                        }

                                        else if (Conversions.ToBoolean(!(Operators.ConditionalCompareObjectEqual(cell.Font.ColorIndex, global::Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexNone, false))))
                                        {
                                            global::System.Int64 colorValue2 = Conversions.ToLong(cell.Font.Color);
                                            global::System.Int32 red2 = (global::System.Int32)((colorValue2) % (256L));
                                            global::System.Int32 green2 = (global::System.Int32)(((((colorValue2) / (256L)))) % (256L));
                                            global::System.Int32 blue2 = (global::System.Int32)((((((colorValue2) / (256L)) / (256L)))) % (256L));
                                            label2.ForeColor = global::System.Drawing.Color.FromArgb(red2, green2, blue2);
                                        }
                                    }
                                    this.CustomPanel2.Controls.Add(label2);
                                    ordinate = ((ordinate) + (widths[(j) - (1)]));
                                }

                            }
                        }

                        else if (X2)
                        {

                            if (((c) <= (4)))
                            {
                                Height = ((global::System.Double)(this.CustomPanel2.Height) / (global::System.Double)(c));
                            }
                            else
                            {
                                Height = (((119d) / (4d)));
                            }

                            global::System.Int32 Index;
                            ordinate = 0d;
                            for (global::System.Int32 i = 1, loopTo31 = r; i <= loopTo31; i++)
                            {
                                global::System.String source = Conversions.ToString(displayRng.Cells[(global::System.Object)i, (global::System.Object)SplitColumn].value);
                                var values = new global::System.String[(c)];
                                Index = -(1);
                                for (global::System.Int32 j = 1, loopTo32 = (c) - (1); j <= loopTo32; j++)
                                {
                                    Index = ((Index) + (1));
                                    values[(j) - (1)] = Conversions.ToString(displayRng.Cells[(global::System.Object)i, (global::System.Object)j].value);
                                }
                                for (global::System.Double k = 1d, loopTo33 = global::Microsoft.VisualBasic.Conversion.Int((global::System.Double)(global::Microsoft.VisualBasic.Strings.Len(source)) / (global::System.Double)(W)); k <= loopTo33; k++)
                                {
                                    values[(SplitColumn) - (1)] = global::Microsoft.VisualBasic.Strings.Mid(source, (global::System.Int32)Math.Round(((((global::System.Double)(W) * ((((k) - (1d))))))) + (1d)), W);
                                    Width = Conversions.ToDouble(Operators.DivideObject((Operators.MultiplyObject(this.MaxOfArray(values), BaseWidth)), 10));
                                    for (global::System.Int32 m = global::Microsoft.VisualBasic.Information.LBound(values), loopTo34 = global::Microsoft.VisualBasic.Information.UBound(values); m <= loopTo34; m++)
                                    {
                                        var label1 = new global::System.Windows.Forms.Label();
                                        label1.Text = values[m];
                                        label1.Location = new global::System.Drawing.Point((global::System.Int32)Math.Round(ordinate), (global::System.Int32)Math.Round((global::System.Double)(m) * (Height)));
                                        label1.Height = (global::System.Int32)Math.Round(Height);
                                        label1.Width = (global::System.Int32)Math.Round(Width);
                                        label1.BorderStyle = global::System.Windows.Forms.BorderStyle.FixedSingle;
                                        label1.TextAlign = global::System.Drawing.ContentAlignment.MiddleCenter;

                                        if (((this.CheckBox1.Checked) == (true)))
                                        {

                                            global::Microsoft.Office.Interop.Excel.Range cell = (global::Microsoft.Office.Interop.Excel.Range)displayRng.Cells[(global::System.Object)i, (global::System.Object)((m) + (1))];
                                            var font = cell.Font;

                                            var fontStyle = global::System.Drawing.FontStyle.Regular;
                                            if (Conversions.ToBoolean(cell.Font.Bold))
                                                fontStyle = (fontStyle | global::System.Drawing.FontStyle.Bold);
                                            if (Conversions.ToBoolean(cell.Font.Italic))
                                                fontStyle = (fontStyle | global::System.Drawing.FontStyle.Italic);

                                            global::System.Single fontSize = global::System.Convert.ToSingle(font.Size);

                                            label1.Font = new global::System.Drawing.Font(font.ToString(), fontSize, fontStyle);
                                            if (Conversions.ToBoolean(!(Operators.ConditionalCompareObjectEqual(cell.Interior.ColorIndex, global::Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexNone, false))))
                                            {
                                                global::System.Int64 colorValue1 = Conversions.ToLong(cell.Interior.Color);
                                                global::System.Int32 red1 = (global::System.Int32)((colorValue1) % (256L));
                                                global::System.Int32 green1 = (global::System.Int32)(((((colorValue1) / (256L)))) % (256L));
                                                global::System.Int32 blue1 = (global::System.Int32)((((((colorValue1) / (256L)) / (256L)))) % (256L));
                                                label1.BackColor = global::System.Drawing.Color.FromArgb(red1, green1, blue1);
                                            }

                                            if ((cell.Font.Color is System.DBNull))
                                            {
                                                label1.ForeColor = global::System.Drawing.Color.FromArgb(0, 0, 0);
                                            }

                                            else if (Conversions.ToBoolean(!(Operators.ConditionalCompareObjectEqual(cell.Font.ColorIndex, global::Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexNone, false))))
                                            {
                                                global::System.Int64 colorValue2 = Conversions.ToLong(cell.Font.Color);
                                                global::System.Int32 red2 = (global::System.Int32)((colorValue2) % (256L));
                                                global::System.Int32 green2 = (global::System.Int32)(((((colorValue2) / (256L)))) % (256L));
                                                global::System.Int32 blue2 = (global::System.Int32)((((((colorValue2) / (256L)) / (256L)))) % (256L));
                                                label1.ForeColor = global::System.Drawing.Color.FromArgb(red2, green2, blue2);
                                            }
                                        }
                                        this.CustomPanel2.Controls.Add(label1);
                                    }
                                    ordinate = ((ordinate) + (Width));
                                }
                                if ((((global::Microsoft.VisualBasic.Strings.Len(source)) % (W)) != (0)))
                                {
                                    values[(SplitColumn) - (1)] = global::Microsoft.VisualBasic.Strings.Mid(source, ((global::Microsoft.VisualBasic.Strings.Len(source)) - ((((global::Microsoft.VisualBasic.Strings.Len(source)) % (W))))) + (1), (global::Microsoft.VisualBasic.Strings.Len(source)) % (W));
                                    Width = Conversions.ToDouble(Operators.DivideObject((Operators.MultiplyObject(this.MaxOfArray(values), BaseWidth)), 10));
                                    for (global::System.Int32 m = global::Microsoft.VisualBasic.Information.LBound(values), loopTo35 = global::Microsoft.VisualBasic.Information.UBound(values); m <= loopTo35; m++)
                                    {
                                        var label1 = new global::System.Windows.Forms.Label();
                                        label1.Text = values[m];
                                        label1.Location = new global::System.Drawing.Point((global::System.Int32)Math.Round(ordinate), (global::System.Int32)Math.Round((global::System.Double)(m) * (Height)));
                                        label1.Height = (global::System.Int32)Math.Round(Height);
                                        label1.Width = (global::System.Int32)Math.Round(Width);
                                        label1.BorderStyle = global::System.Windows.Forms.BorderStyle.FixedSingle;
                                        label1.TextAlign = global::System.Drawing.ContentAlignment.MiddleCenter;

                                        if (((this.CheckBox2.Checked) == (true)))
                                        {

                                            global::Microsoft.Office.Interop.Excel.Range cell = (global::Microsoft.Office.Interop.Excel.Range)displayRng.Cells[(global::System.Object)i, (global::System.Object)((m) + (1))];
                                            var font = cell.Font;

                                            var fontStyle = global::System.Drawing.FontStyle.Regular;
                                            if (Conversions.ToBoolean(cell.Font.Bold))
                                                fontStyle = (fontStyle | global::System.Drawing.FontStyle.Bold);
                                            if (Conversions.ToBoolean(cell.Font.Italic))
                                                fontStyle = (fontStyle | global::System.Drawing.FontStyle.Italic);

                                            global::System.Single fontSize = global::System.Convert.ToSingle(font.Size);

                                            label1.Font = new global::System.Drawing.Font(font.ToString(), fontSize, fontStyle);
                                            if (Conversions.ToBoolean(!(Operators.ConditionalCompareObjectEqual(cell.Interior.ColorIndex, global::Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexNone, false))))
                                            {
                                                global::System.Int64 colorValue1 = Conversions.ToLong(cell.Interior.Color);
                                                global::System.Int32 red1 = (global::System.Int32)((colorValue1) % (256L));
                                                global::System.Int32 green1 = (global::System.Int32)(((((colorValue1) / (256L)))) % (256L));
                                                global::System.Int32 blue1 = (global::System.Int32)((((((colorValue1) / (256L)) / (256L)))) % (256L));
                                                label1.BackColor = global::System.Drawing.Color.FromArgb(red1, green1, blue1);
                                            }

                                            if ((cell.Font.Color is System.DBNull))
                                            {
                                                label1.ForeColor = global::System.Drawing.Color.FromArgb(0, 0, 0);
                                            }

                                            else if (Conversions.ToBoolean(!(Operators.ConditionalCompareObjectEqual(cell.Font.ColorIndex, global::Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexNone, false))))
                                            {
                                                global::System.Int64 colorValue2 = Conversions.ToLong(cell.Font.Color);
                                                global::System.Int32 red2 = (global::System.Int32)((colorValue2) % (256L));
                                                global::System.Int32 green2 = (global::System.Int32)(((((colorValue2) / (256L)))) % (256L));
                                                global::System.Int32 blue2 = (global::System.Int32)((((((colorValue2) / (256L)) / (256L)))) % (256L));
                                                label1.ForeColor = global::System.Drawing.Color.FromArgb(red2, green2, blue2);
                                            }
                                        }
                                        this.CustomPanel2.Controls.Add(label1);
                                    }
                                    ordinate = ((ordinate) + (Width));
                                }
                            }
                        }

                    }

                    this.CustomPanel2.AutoScroll = true;

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

                if (string.IsNullOrEmpty(this.TextBox1.Text))
                {
                    global::System.Windows.Forms.MessageBox.Show("Select a Source Range.", "Error", global::System.Windows.Forms.MessageBoxButtons.OK, global::System.Windows.Forms.MessageBoxIcon.Error);
                    this.TextBox1.Focus();
                    this.workSheet.Activate();
                    this.rng.Select();
                    return;
                }

                if (((this.IsValidExcelCellReference(this.TextBox1.Text)) == (false)))
                {
                    global::System.Windows.Forms.MessageBox.Show("Select a Valid Source Range.", "Error", global::System.Windows.Forms.MessageBoxButtons.OK, global::System.Windows.Forms.MessageBoxIcon.Error);
                    this.TextBox1.Focus();
                    this.workSheet.Activate();
                    this.rng.Select();
                    return;
                }

                if ((((this.RadioButton4.Checked) == (false)) & ((this.RadioButton5.Checked) == (false))))
                {
                    global::System.Windows.Forms.MessageBox.Show("Enter a Destination Cell.", "Error", global::System.Windows.Forms.MessageBoxButtons.OK, global::System.Windows.Forms.MessageBoxIcon.Error);
                    this.workSheet.Activate();
                    this.rng.Select();
                    return;
                }

                if ((((this.RadioButton4.Checked) == (true)) & (((string.IsNullOrEmpty(this.TextBox4.Text) | ((this.IsValidExcelCellReference(this.TextBox4.Text)) == (false)))))))
                {
                    global::System.Windows.Forms.MessageBox.Show("Enter a valid Destination Cell.", "Error", global::System.Windows.Forms.MessageBoxButtons.OK, global::System.Windows.Forms.MessageBoxIcon.Error);
                    this.workSheet.Activate();
                    this.rng.Select();
                    return;
                }

                if (((this.CheckBox2.Checked) == (true)))
                {
                    this.workSheet.Copy(After: workBook.Sheets[workSheet.Name]);
                }


                global::System.Boolean X1 = this.RadioButton1.Checked;
                global::System.Boolean X2 = this.RadioButton2.Checked;
                global::System.Boolean X3 = this.RadioButton3.Checked;
                global::System.Boolean X7 = this.RadioButton7.Checked;
                global::System.Boolean X8 = this.RadioButton8.Checked;
                global::System.Boolean X9 = this.RadioButton9.Checked;
                global::System.Boolean X10 = this.RadioButton10.Checked;
                global::System.Boolean X11 = this.RadioButton11.Checked;
                global::System.Boolean X12 = ((this.ComboBox3.SelectedIndex) != (-(1)));

                if (((X12) == (false)))
                {
                    global::System.Windows.Forms.MessageBox.Show("Select a Column by Which You Want to Split the Range.", "Error", global::System.Windows.Forms.MessageBoxButtons.OK, global::System.Windows.Forms.MessageBoxIcon.Error);
                    this.workSheet.Activate();
                    this.rng.Select();
                    return;
                }

                if ((((X1) == (false)) & ((X2) == (false))))
                {
                    global::System.Windows.Forms.MessageBox.Show("Select a Split Option.", "Error", global::System.Windows.Forms.MessageBoxButtons.OK, global::System.Windows.Forms.MessageBoxIcon.Error);
                    this.workSheet.Activate();
                    this.rng.Select();
                    return;
                }

                if ((((((((X3) == (false)) & ((X7) == (false))) & ((X8) == (false))) & ((X9) == (false))) & ((X10) == (false))) & ((X11) == (false))))
                {
                    global::System.Windows.Forms.MessageBox.Show("Select a Separator to Split the Range.", "Error", global::System.Windows.Forms.MessageBoxButtons.OK, global::System.Windows.Forms.MessageBoxIcon.Error);
                    this.workSheet.Activate();
                    this.rng.Select();
                    return;
                }

                global::System.Int32 r = this.rng.Rows.Count;
                global::System.Int32 c = this.rng.Columns.Count;

                if (((((((X1) | (X2)))) & (X12)) & ((((((((X3) | (X7)) | (X8)) | (X9)) | (X10)) | (X11))))))
                {

                    global::System.Int32 TotalRows = 0;
                    global::System.Int32 SplitColumn = ((this.ComboBox3.SelectedIndex) + (1));
                    global::System.String Separator = "";
                    if (X7)
                    {
                        Separator = ";";
                    }
                    else if (X8)
                    {
                        Separator = global::Microsoft.VisualBasic.Constants.vbNewLine;
                    }
                    else if (X9)
                    {
                        Separator = " ";
                    }
                    else if (X10)
                    {
                        Separator = this.ComboBox2.Text;
                    }

                    for (global::System.Int32 i = 1, loopTo = r; i <= loopTo; i++)
                        TotalRows = ((TotalRows) + (this.CountSeparator(Conversions.ToString(this.rng.Cells[(global::System.Object)i, (global::System.Object)SplitColumn].value), Separator)));

                    if (X1)
                    {
                        this.rng2 = this.workSheet2.get_Range(this.rng2.Cells[(global::System.Object)1, (global::System.Object)1], this.rng2.Cells[(global::System.Object)TotalRows, (global::System.Object)c]);
                    }
                    else
                    {
                        this.rng2 = this.workSheet2.get_Range(this.rng2.Cells[(global::System.Object)1, (global::System.Object)1], this.rng2.Cells[(global::System.Object)c, (global::System.Object)TotalRows]);
                    }

                    if (((this.Overlap(this.excelApp, this.workSheet, this.workSheet2, this.rng, this.rng2)) == (false)))
                    {

                        global::System.String rng2Address = this.rng2.get_Address();


                        if (((((X7) | (X8)) | (X9)) | (X10)))
                        {

                            if (X1)
                            {

                                global::System.Int32 Index = 0;
                                global::System.Int32 position;

                                for (global::System.Int32 i = 1, loopTo1 = r; i <= loopTo1; i++)
                                {
                                    global::System.String source = Conversions.ToString(this.rng.Cells[(global::System.Object)i, (global::System.Object)SplitColumn].value);
                                    position = 1;
                                    for (global::System.Int32 k = 1, loopTo2 = global::Microsoft.VisualBasic.Strings.Len(source); k <= loopTo2; k++)
                                    {
                                        if (((global::Microsoft.VisualBasic.Strings.Mid(source, k, global::Microsoft.VisualBasic.Strings.Len(Separator)) ?? "") == (Separator ?? "")))
                                        {
                                            if ((((k) - (position)) > (0)))
                                            {
                                                Index = ((Index) + (1));
                                                for (global::System.Int32 j = 1, loopTo3 = (SplitColumn) - (1); j <= loopTo3; j++)
                                                {
                                                    this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].value = this.rng.Cells[(global::System.Object)i, (global::System.Object)j].value;
                                                    if (((this.CheckBox1.Checked) == (true)))
                                                    {
                                                        this.rng.Cells[(global::System.Object)i, (global::System.Object)j].copy();
                                                        this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].PasteSpecial(global::Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats);
                                                        this.rng2 = this.workSheet2.get_Range(rng2Address);
                                                        this.workSheet2.Activate();
                                                    }
                                                    else
                                                    {
                                                        this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].ClearFormats();
                                                    }
                                                }

                                                this.rng2.Cells[(global::System.Object)Index, (global::System.Object)SplitColumn] = global::Microsoft.VisualBasic.Strings.Mid(source, position, (k) - (position));
                                                if (((this.CheckBox1.Checked) == (true)))
                                                {
                                                    this.rng.Cells[(global::System.Object)i, (global::System.Object)SplitColumn].copy();
                                                    this.rng2.Cells[(global::System.Object)Index, (global::System.Object)SplitColumn].PasteSpecial(global::Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats);
                                                    this.rng2 = this.workSheet2.get_Range(rng2Address);
                                                    this.workSheet2.Activate();
                                                }
                                                else
                                                {
                                                    this.rng2.Cells[(global::System.Object)Index, (global::System.Object)SplitColumn].ClearFormats();
                                                }

                                                for (global::System.Int32 j = (SplitColumn) + (1), loopTo4 = c; j <= loopTo4; j++)
                                                {
                                                    this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].value = this.rng.Cells[(global::System.Object)i, (global::System.Object)j].value;
                                                    if (((this.CheckBox1.Checked) == (true)))
                                                    {
                                                        this.rng.Cells[(global::System.Object)i, (global::System.Object)j].copy();
                                                        this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].PasteSpecial(global::Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats);
                                                        this.rng2 = this.workSheet2.get_Range(rng2Address);
                                                        this.workSheet2.Activate();
                                                    }
                                                    else
                                                    {
                                                        this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].ClearFormats();
                                                    }
                                                }

                                            }
                                            position = ((k) + (global::Microsoft.VisualBasic.Strings.Len(Separator)));
                                        }
                                    }
                                    if (((position) <= (global::Microsoft.VisualBasic.Strings.Len(source))))
                                    {
                                        Index = ((Index) + (1));

                                        for (global::System.Int32 j = 1, loopTo5 = (SplitColumn) - (1); j <= loopTo5; j++)
                                        {
                                            this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].Value = this.rng.Cells[(global::System.Object)i, (global::System.Object)j].value;
                                            if (((this.CheckBox1.Checked) == (true)))
                                            {
                                                this.rng.Cells[(global::System.Object)i, (global::System.Object)j].copy();
                                                this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].PasteSpecial(global::Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats);
                                                this.rng2 = this.workSheet2.get_Range(rng2Address);
                                                this.workSheet2.Activate();
                                            }
                                            else
                                            {
                                                this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].ClearFormats();
                                            }
                                        }

                                        this.rng2.Cells[(global::System.Object)Index, (global::System.Object)SplitColumn] = global::Microsoft.VisualBasic.Strings.Mid(source, position, ((global::Microsoft.VisualBasic.Strings.Len(source)) - (position)) + (1));
                                        if (((this.CheckBox1.Checked) == (true)))
                                        {
                                            this.rng.Cells[(global::System.Object)i, (global::System.Object)SplitColumn].copy();
                                            this.rng2.Cells[(global::System.Object)Index, (global::System.Object)SplitColumn].PasteSpecial(global::Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats);
                                            this.rng2 = this.workSheet2.get_Range(rng2Address);
                                            this.workSheet2.Activate();
                                        }
                                        else
                                        {
                                            this.rng2.Cells[(global::System.Object)Index, (global::System.Object)SplitColumn].ClearFormats();
                                        }

                                        for (global::System.Int32 j = (SplitColumn) + (1), loopTo6 = c; j <= loopTo6; j++)
                                        {
                                            this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].value = this.rng.Cells[(global::System.Object)i, (global::System.Object)j].value;
                                            if (((this.CheckBox1.Checked) == (true)))
                                            {
                                                this.rng.Cells[(global::System.Object)i, (global::System.Object)j].copy();
                                                this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].PasteSpecial(global::Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats);
                                                this.rng2 = this.workSheet2.get_Range(rng2Address);
                                                this.workSheet2.Activate();
                                            }
                                            else
                                            {
                                                this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].ClearFormats();
                                            }
                                        }

                                    }
                                }
                                excelApp.CutCopyMode = (global::Microsoft.Office.Interop.Excel.XlCutCopyMode)Conversions.ToInteger(false);
                            }

                            else if (X2)
                            {

                                global::System.Int32 Index = 0;
                                global::System.Int32 position;

                                for (global::System.Int32 i = 1, loopTo7 = r; i <= loopTo7; i++)
                                {
                                    global::System.String source = Conversions.ToString(this.rng.Cells[(global::System.Object)i, (global::System.Object)SplitColumn].value);
                                    position = 1;
                                    for (global::System.Int32 k = 1, loopTo8 = global::Microsoft.VisualBasic.Strings.Len(source); k <= loopTo8; k++)
                                    {
                                        if (((global::Microsoft.VisualBasic.Strings.Mid(source, k, global::Microsoft.VisualBasic.Strings.Len(Separator)) ?? "") == (Separator ?? "")))
                                        {
                                            if ((((k) - (position)) > (0)))
                                            {
                                                Index = ((Index) + (1));
                                                for (global::System.Int32 j = 1, loopTo9 = (SplitColumn) - (1); j <= loopTo9; j++)
                                                {
                                                    this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].value = this.rng.Cells[(global::System.Object)i, (global::System.Object)j].value;
                                                    if (((this.CheckBox1.Checked) == (true)))
                                                    {
                                                        this.rng.Cells[(global::System.Object)i, (global::System.Object)j].copy();
                                                        this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].PasteSpecial(global::Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats);
                                                        this.rng2 = this.workSheet2.get_Range(rng2Address);
                                                        this.workSheet2.Activate();
                                                    }
                                                    else
                                                    {
                                                        this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].ClearFormats();
                                                    }
                                                }

                                                this.rng2.Cells[(global::System.Object)SplitColumn, (global::System.Object)Index] = global::Microsoft.VisualBasic.Strings.Mid(source, position, (k) - (position));
                                                if (((this.CheckBox1.Checked) == (true)))
                                                {
                                                    this.rng.Cells[(global::System.Object)i, (global::System.Object)SplitColumn].copy();
                                                    this.rng2.Cells[(global::System.Object)SplitColumn, (global::System.Object)Index].PasteSpecial(global::Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats);
                                                    this.rng2 = this.workSheet2.get_Range(rng2Address);
                                                    this.workSheet2.Activate();
                                                }
                                                else
                                                {
                                                    this.rng2.Cells[(global::System.Object)SplitColumn, (global::System.Object)Index].ClearFormats();
                                                }

                                                for (global::System.Int32 j = (SplitColumn) + (1), loopTo10 = c; j <= loopTo10; j++)
                                                {
                                                    this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].value = this.rng.Cells[(global::System.Object)i, (global::System.Object)j].value;
                                                    if (((this.CheckBox1.Checked) == (true)))
                                                    {
                                                        this.rng.Cells[(global::System.Object)i, (global::System.Object)j].copy();
                                                        this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].PasteSpecial(global::Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats);
                                                        this.rng2 = this.workSheet2.get_Range(rng2Address);
                                                        this.workSheet2.Activate();
                                                    }
                                                    else
                                                    {
                                                        this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].ClearFormats();
                                                    }
                                                }
                                            }
                                            position = ((k) + (global::Microsoft.VisualBasic.Strings.Len(Separator)));
                                        }
                                    }
                                    if (((position) <= (global::Microsoft.VisualBasic.Strings.Len(source))))
                                    {
                                        Index = ((Index) + (1));
                                        for (global::System.Int32 j = 1, loopTo11 = (SplitColumn) - (1); j <= loopTo11; j++)
                                        {
                                            this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].Value = this.rng.Cells[(global::System.Object)i, (global::System.Object)j].value;
                                            if (((this.CheckBox1.Checked) == (true)))
                                            {
                                                this.rng.Cells[(global::System.Object)i, (global::System.Object)j].copy();
                                                this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].PasteSpecial(global::Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats);
                                                this.rng2 = this.workSheet2.get_Range(rng2Address);
                                                this.workSheet2.Activate();
                                            }
                                            else
                                            {
                                                this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].ClearFormats();
                                            }
                                        }

                                        this.rng2.Cells[(global::System.Object)SplitColumn, (global::System.Object)Index] = global::Microsoft.VisualBasic.Strings.Mid(source, position, ((global::Microsoft.VisualBasic.Strings.Len(source)) - (position)) + (1));
                                        if (((this.CheckBox1.Checked) == (true)))
                                        {
                                            this.rng.Cells[(global::System.Object)i, (global::System.Object)SplitColumn].copy();
                                            this.rng2.Cells[(global::System.Object)SplitColumn, (global::System.Object)Index].PasteSpecial(global::Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats);
                                            this.rng2 = this.workSheet2.get_Range(rng2Address);
                                            this.workSheet2.Activate();
                                        }
                                        else
                                        {
                                            this.rng2.Cells[(global::System.Object)SplitColumn, (global::System.Object)Index].ClearFormats();
                                        }

                                        for (global::System.Int32 j = (SplitColumn) + (1), loopTo12 = c; j <= loopTo12; j++)
                                        {
                                            this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].value = this.rng.Cells[(global::System.Object)i, (global::System.Object)j].value;
                                            if (((this.CheckBox1.Checked) == (true)))
                                            {
                                                this.rng.Cells[(global::System.Object)i, (global::System.Object)j].copy();
                                                this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].PasteSpecial(global::Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats);
                                                this.rng2 = this.workSheet2.get_Range(rng2Address);
                                                this.workSheet2.Activate();
                                            }
                                            else
                                            {
                                                this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].ClearFormats();
                                            }
                                        }
                                    }
                                }
                                excelApp.CutCopyMode = (global::Microsoft.Office.Interop.Excel.XlCutCopyMode)Conversions.ToInteger(false);
                            }
                        }

                        else if (X3)
                        {

                            if (X1)
                            {

                                global::System.Int32 Index = 0;

                                for (global::System.Int32 i = 1, loopTo13 = r; i <= loopTo13; i++)
                                {

                                    global::System.String source = Conversions.ToString(this.rng.Cells[(global::System.Object)i, (global::System.Object)SplitColumn].value);
                                    var NumberText = new global::System.String[2];
                                    NumberText = (global::System.String[])this.SeparateNumberText(source);
                                    global::System.String Number = NumberText[0];
                                    global::System.String Text = NumberText[1];

                                    Index = ((Index) + (1));
                                    for (global::System.Int32 j = 1, loopTo14 = (SplitColumn) - (1); j <= loopTo14; j++)
                                    {
                                        this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].value = this.rng.Cells[(global::System.Object)i, (global::System.Object)j].value;
                                        if (((this.CheckBox1.Checked) == (true)))
                                        {
                                            this.rng.Cells[(global::System.Object)i, (global::System.Object)j].copy();
                                            this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].PasteSpecial(global::Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats);
                                            this.rng2 = this.workSheet2.get_Range(rng2Address);
                                            this.workSheet2.Activate();
                                        }
                                        else
                                        {
                                            this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].ClearFormats();
                                        }
                                    }

                                    this.rng2.Cells[(global::System.Object)Index, (global::System.Object)SplitColumn].value = Number;
                                    if (((this.CheckBox1.Checked) == (true)))
                                    {
                                        this.rng.Cells[(global::System.Object)i, (global::System.Object)SplitColumn].copy();
                                        this.rng2.Cells[(global::System.Object)Index, (global::System.Object)SplitColumn].PasteSpecial(global::Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats);
                                        this.rng2 = this.workSheet2.get_Range(rng2Address);
                                        this.workSheet2.Activate();
                                    }
                                    else
                                    {
                                        this.rng2.Cells[(global::System.Object)Index, (global::System.Object)SplitColumn].ClearFormats();
                                    }

                                    for (global::System.Int32 j = (SplitColumn) + (1), loopTo15 = c; j <= loopTo15; j++)
                                    {
                                        this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].value = this.rng.Cells[(global::System.Object)i, (global::System.Object)j].value;
                                        if (((this.CheckBox1.Checked) == (true)))
                                        {
                                            this.rng.Cells[(global::System.Object)i, (global::System.Object)j].copy();
                                            this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].PasteSpecial(global::Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats);
                                            this.rng2 = this.workSheet2.get_Range(rng2Address);
                                            this.workSheet2.Activate();
                                        }
                                        else
                                        {
                                            this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].ClearFormats();
                                        }
                                    }

                                    Index = ((Index) + (1));
                                    for (global::System.Int32 j = 1, loopTo16 = (SplitColumn) - (1); j <= loopTo16; j++)
                                    {
                                        this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].value = this.rng.Cells[(global::System.Object)i, (global::System.Object)j].value;
                                        if (((this.CheckBox1.Checked) == (true)))
                                        {
                                            this.rng.Cells[(global::System.Object)i, (global::System.Object)j].copy();
                                            this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].PasteSpecial(global::Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats);
                                            this.rng2 = this.workSheet2.get_Range(rng2Address);
                                            this.workSheet2.Activate();
                                        }
                                        else
                                        {
                                            this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].ClearFormats();
                                        }
                                    }

                                    this.rng2.Cells[(global::System.Object)Index, (global::System.Object)SplitColumn].value = Text;
                                    if (((this.CheckBox1.Checked) == (true)))
                                    {
                                        this.rng.Cells[(global::System.Object)i, (global::System.Object)SplitColumn].copy();
                                        this.rng2.Cells[(global::System.Object)Index, (global::System.Object)SplitColumn].PasteSpecial(global::Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats);
                                        this.rng2 = this.workSheet2.get_Range(rng2Address);
                                        this.workSheet2.Activate();
                                    }
                                    else
                                    {
                                        this.rng2.Cells[(global::System.Object)Index, (global::System.Object)SplitColumn].ClearFormats();
                                    }

                                    for (global::System.Int32 j = (SplitColumn) + (1), loopTo17 = c; j <= loopTo17; j++)
                                    {
                                        this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].value = this.rng.Cells[(global::System.Object)i, (global::System.Object)j].value;
                                        if (((this.CheckBox1.Checked) == (true)))
                                        {
                                            this.rng.Cells[(global::System.Object)i, (global::System.Object)j].copy();
                                            this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].PasteSpecial(global::Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats);
                                            this.rng2 = this.workSheet2.get_Range(rng2Address);
                                            this.workSheet2.Activate();
                                        }
                                        else
                                        {
                                            this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].ClearFormats();
                                        }
                                    }

                                }
                                excelApp.CutCopyMode = (global::Microsoft.Office.Interop.Excel.XlCutCopyMode)Conversions.ToInteger(false);
                            }

                            else if (X2)
                            {

                                global::System.Int32 Index = 0;

                                for (global::System.Int32 i = 1, loopTo18 = r; i <= loopTo18; i++)
                                {

                                    global::System.String source = Conversions.ToString(this.rng.Cells[(global::System.Object)i, (global::System.Object)SplitColumn].value);
                                    var NumberText = new global::System.String[2];
                                    NumberText = (global::System.String[])this.SeparateNumberText(source);
                                    global::System.String Number = NumberText[0];
                                    global::System.String Text = NumberText[1];

                                    Index = ((Index) + (1));
                                    for (global::System.Int32 j = 1, loopTo19 = (c) - (1); j <= loopTo19; j++)
                                    {
                                        this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].value = this.rng.Cells[(global::System.Object)i, (global::System.Object)j].value;
                                        if (((this.CheckBox1.Checked) == (true)))
                                        {
                                            this.rng.Cells[(global::System.Object)i, (global::System.Object)j].copy();
                                            this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].PasteSpecial(global::Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats);
                                            this.rng2 = this.workSheet2.get_Range(rng2Address);
                                            this.workSheet2.Activate();
                                        }
                                        else
                                        {
                                            this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].ClearFormats();
                                        }
                                    }

                                    this.rng2.Cells[(global::System.Object)SplitColumn, (global::System.Object)Index].value = Number;
                                    if (((this.CheckBox1.Checked) == (true)))
                                    {
                                        this.rng.Cells[(global::System.Object)i, (global::System.Object)SplitColumn].copy();
                                        this.rng2.Cells[(global::System.Object)SplitColumn, (global::System.Object)Index].PasteSpecial(global::Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats);
                                        this.rng2 = this.workSheet2.get_Range(rng2Address);
                                        this.workSheet2.Activate();
                                    }
                                    else
                                    {
                                        this.rng2.Cells[(global::System.Object)SplitColumn, (global::System.Object)Index].ClearFormats();
                                    }

                                    for (global::System.Int32 j = (SplitColumn) + (1), loopTo20 = c; j <= loopTo20; j++)
                                    {
                                        this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].value = this.rng.Cells[(global::System.Object)i, (global::System.Object)j].value;
                                        if (((this.CheckBox1.Checked) == (true)))
                                        {
                                            this.rng.Cells[(global::System.Object)i, (global::System.Object)j].copy();
                                            this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].PasteSpecial(global::Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats);
                                            this.rng2 = this.workSheet2.get_Range(rng2Address);
                                            this.workSheet2.Activate();
                                        }
                                        else
                                        {
                                            this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].ClearFormats();
                                        }
                                    }

                                    Index = ((Index) + (1));
                                    for (global::System.Int32 j = 1, loopTo21 = (SplitColumn) - (1); j <= loopTo21; j++)
                                    {
                                        this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].value = this.rng.Cells[(global::System.Object)i, (global::System.Object)j].value;
                                        if (((this.CheckBox1.Checked) == (true)))
                                        {
                                            this.rng.Cells[(global::System.Object)i, (global::System.Object)j].copy();
                                            this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].PasteSpecial(global::Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats);
                                            this.rng2 = this.workSheet2.get_Range(rng2Address);
                                            this.workSheet2.Activate();
                                        }
                                        else
                                        {
                                            this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].ClearFormats();
                                        }
                                    }

                                    this.rng2.Cells[(global::System.Object)SplitColumn, (global::System.Object)Index].value = Text;
                                    if (((this.CheckBox1.Checked) == (true)))
                                    {
                                        this.rng.Cells[(global::System.Object)i, (global::System.Object)SplitColumn].copy();
                                        this.rng2.Cells[(global::System.Object)SplitColumn, (global::System.Object)Index].PasteSpecial(global::Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats);
                                        this.rng2 = this.workSheet2.get_Range(rng2Address);
                                        this.workSheet2.Activate();
                                    }
                                    else
                                    {
                                        this.rng2.Cells[(global::System.Object)SplitColumn, (global::System.Object)Index].ClearFormats();
                                    }

                                    for (global::System.Int32 j = (SplitColumn) + (1), loopTo22 = c; j <= loopTo22; j++)
                                    {
                                        this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].value = this.rng.Cells[(global::System.Object)i, (global::System.Object)j].value;
                                        if (((this.CheckBox1.Checked) == (true)))
                                        {
                                            this.rng.Cells[(global::System.Object)i, (global::System.Object)j].copy();
                                            this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].PasteSpecial(global::Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats);
                                            this.rng2 = this.workSheet2.get_Range(rng2Address);
                                            this.workSheet2.Activate();
                                        }
                                        else
                                        {
                                            this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].ClearFormats();
                                        }
                                    }
                                }

                                excelApp.CutCopyMode = (global::Microsoft.Office.Interop.Excel.XlCutCopyMode)Conversions.ToInteger(false);

                            }
                        }

                        else if (X11)
                        {

                            global::System.Int32 W;

                            if (string.IsNullOrEmpty(this.TextBox3.Text))
                            {
                                W = 1;
                            }
                            else
                            {
                                W = Conversions.ToInteger(global::Microsoft.VisualBasic.Conversion.Int(this.TextBox3.Text));
                            }

                            if (X1)
                            {

                                global::System.Int32 Index = 0;

                                for (global::System.Int32 i = 1, loopTo23 = r; i <= loopTo23; i++)
                                {
                                    global::System.String source = Conversions.ToString(this.rng.Cells[(global::System.Object)i, (global::System.Object)SplitColumn].value);
                                    for (global::System.Double k = 1d, loopTo24 = global::Microsoft.VisualBasic.Conversion.Int((global::System.Double)(global::Microsoft.VisualBasic.Strings.Len(source)) / (global::System.Double)(W)); k <= loopTo24; k++)
                                    {
                                        Index = ((Index) + (1));
                                        for (global::System.Int32 j = 1, loopTo25 = (SplitColumn) - (1); j <= loopTo25; j++)
                                        {
                                            this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].value = this.rng.Cells[(global::System.Object)i, (global::System.Object)j].value;
                                            if (((this.CheckBox1.Checked) == (true)))
                                            {
                                                this.rng.Cells[(global::System.Object)i, (global::System.Object)j].copy();
                                                this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].PasteSpecial(global::Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats);
                                                this.rng2 = this.workSheet2.get_Range(rng2Address);
                                                this.workSheet2.Activate();
                                            }
                                            else
                                            {
                                                this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].ClearFormats();
                                            }
                                        }
                                        this.rng2.Cells[(global::System.Object)Index, (global::System.Object)SplitColumn].value = global::Microsoft.VisualBasic.Strings.Mid(source, (global::System.Int32)Math.Round(((((global::System.Double)(W) * ((((k) - (1d))))))) + (1d)), W);
                                        if (((this.CheckBox1.Checked) == (true)))
                                        {
                                            this.rng.Cells[(global::System.Object)i, (global::System.Object)SplitColumn].copy();
                                            this.rng2.Cells[(global::System.Object)Index, (global::System.Object)SplitColumn].PasteSpecial(global::Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats);
                                            this.rng2 = this.workSheet2.get_Range(rng2Address);
                                            this.workSheet2.Activate();
                                        }
                                        else
                                        {
                                            this.rng2.Cells[(global::System.Object)Index, (global::System.Object)SplitColumn].ClearFormats();
                                        }
                                        for (global::System.Int32 j = (SplitColumn) + (1), loopTo26 = c; j <= loopTo26; j++)
                                        {
                                            this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].value = this.rng.Cells[(global::System.Object)i, (global::System.Object)j].value;
                                            if (((this.CheckBox1.Checked) == (true)))
                                            {
                                                this.rng.Cells[(global::System.Object)i, (global::System.Object)j].copy();
                                                this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].PasteSpecial(global::Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats);
                                                this.rng2 = this.workSheet2.get_Range(rng2Address);
                                                this.workSheet2.Activate();
                                            }
                                            else
                                            {
                                                this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].ClearFormats();
                                            }
                                        }
                                    }
                                    if ((((global::Microsoft.VisualBasic.Strings.Len(source)) % (W)) != (0)))
                                    {
                                        Index = ((Index) + (1));
                                        for (global::System.Int32 j = 1, loopTo27 = (SplitColumn) - (1); j <= loopTo27; j++)
                                        {
                                            this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].value = this.rng.Cells[(global::System.Object)i, (global::System.Object)j].value;
                                            if (((this.CheckBox1.Checked) == (true)))
                                            {
                                                this.rng.Cells[(global::System.Object)i, (global::System.Object)j].copy();
                                                this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].PasteSpecial(global::Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats);
                                                this.rng2 = this.workSheet2.get_Range(rng2Address);
                                                this.workSheet2.Activate();
                                            }
                                            else
                                            {
                                                this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].ClearFormats();
                                            }
                                        }
                                        this.rng2.Cells[(global::System.Object)Index, (global::System.Object)SplitColumn].value = global::Microsoft.VisualBasic.Strings.Mid(source, ((global::Microsoft.VisualBasic.Strings.Len(source)) - ((((global::Microsoft.VisualBasic.Strings.Len(source)) % (W))))) + (1), (global::Microsoft.VisualBasic.Strings.Len(source)) % (W));
                                        if (((this.CheckBox1.Checked) == (true)))
                                        {
                                            this.rng.Cells[(global::System.Object)i, (global::System.Object)SplitColumn].copy();
                                            this.rng2.Cells[(global::System.Object)Index, (global::System.Object)SplitColumn].PasteSpecial(global::Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats);
                                            this.rng2 = this.workSheet2.get_Range(rng2Address);
                                            this.workSheet2.Activate();
                                        }
                                        else
                                        {
                                            this.rng2.Cells[(global::System.Object)Index, (global::System.Object)SplitColumn].ClearFormats();
                                        }
                                        for (global::System.Int32 j = (SplitColumn) + (1), loopTo28 = c; j <= loopTo28; j++)
                                        {
                                            this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].value = this.rng.Cells[(global::System.Object)i, (global::System.Object)j].value;
                                            if (((this.CheckBox1.Checked) == (true)))
                                            {
                                                this.rng.Cells[(global::System.Object)i, (global::System.Object)j].copy();
                                                this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].PasteSpecial(global::Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats);
                                                this.rng2 = this.workSheet2.get_Range(rng2Address);
                                                this.workSheet2.Activate();
                                            }
                                            else
                                            {
                                                this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].ClearFormats();
                                            }
                                        }
                                    }
                                }

                                excelApp.CutCopyMode = (global::Microsoft.Office.Interop.Excel.XlCutCopyMode)Conversions.ToInteger(false);
                            }

                            else if (X2)
                            {

                                global::System.Int32 Index = 0;

                                for (global::System.Int32 i = 1, loopTo29 = r; i <= loopTo29; i++)
                                {
                                    global::System.String source = Conversions.ToString(this.rng.Cells[(global::System.Object)i, (global::System.Object)SplitColumn].value);
                                    for (global::System.Double k = 1d, loopTo30 = global::Microsoft.VisualBasic.Conversion.Int((global::System.Double)(global::Microsoft.VisualBasic.Strings.Len(source)) / (global::System.Double)(W)); k <= loopTo30; k++)
                                    {
                                        Index = ((Index) + (1));
                                        for (global::System.Int32 j = 1, loopTo31 = (SplitColumn) - (1); j <= loopTo31; j++)
                                        {
                                            this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].value = this.rng.Cells[(global::System.Object)i, (global::System.Object)j].value;
                                            if (((this.CheckBox1.Checked) == (true)))
                                            {
                                                this.rng.Cells[(global::System.Object)i, (global::System.Object)j].copy();
                                                this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].PasteSpecial(global::Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats);
                                                this.rng2 = this.workSheet2.get_Range(rng2Address);
                                                this.workSheet2.Activate();
                                            }
                                            else
                                            {
                                                this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].ClearFormats();
                                            }
                                        }
                                        this.rng2.Cells[(global::System.Object)SplitColumn, (global::System.Object)Index].value = global::Microsoft.VisualBasic.Strings.Mid(source, (global::System.Int32)Math.Round(((((global::System.Double)(W) * ((((k) - (1d))))))) + (1d)), W);
                                        if (((this.CheckBox1.Checked) == (true)))
                                        {
                                            this.rng.Cells[(global::System.Object)i, (global::System.Object)SplitColumn].copy();
                                            this.rng2.Cells[(global::System.Object)SplitColumn, (global::System.Object)Index].PasteSpecial(global::Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats);
                                            this.rng2 = this.workSheet2.get_Range(rng2Address);
                                            this.workSheet2.Activate();
                                        }
                                        else
                                        {
                                            this.rng2.Cells[(global::System.Object)SplitColumn, (global::System.Object)Index].ClearFormats();
                                        }
                                        for (global::System.Int32 j = (SplitColumn) + (1), loopTo32 = c; j <= loopTo32; j++)
                                        {
                                            this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].value = this.rng.Cells[(global::System.Object)i, (global::System.Object)j].value;
                                            if (((this.CheckBox1.Checked) == (true)))
                                            {
                                                this.rng.Cells[(global::System.Object)i, (global::System.Object)j].copy();
                                                this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].PasteSpecial(global::Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats);
                                                this.rng2 = this.workSheet2.get_Range(rng2Address);
                                                this.workSheet2.Activate();
                                            }
                                            else
                                            {
                                                this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].ClearFormats();
                                            }
                                        }
                                    }
                                    if ((((global::Microsoft.VisualBasic.Strings.Len(source)) % (W)) != (0)))
                                    {
                                        Index = ((Index) + (1));
                                        for (global::System.Int32 j = 1, loopTo33 = (SplitColumn) - (1); j <= loopTo33; j++)
                                        {
                                            this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].value = this.rng.Cells[(global::System.Object)i, (global::System.Object)j].value;
                                            if (((this.CheckBox1.Checked) == (true)))
                                            {
                                                this.rng.Cells[(global::System.Object)i, (global::System.Object)j].copy();
                                                this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].PasteSpecial(global::Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats);
                                                this.rng2 = this.workSheet2.get_Range(rng2Address);
                                                this.workSheet2.Activate();
                                            }
                                            else
                                            {
                                                this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].ClearFormats();
                                            }
                                        }
                                        this.rng2.Cells[(global::System.Object)SplitColumn, (global::System.Object)Index].value = global::Microsoft.VisualBasic.Strings.Mid(source, ((global::Microsoft.VisualBasic.Strings.Len(source)) - ((((global::Microsoft.VisualBasic.Strings.Len(source)) % (W))))) + (1), (global::Microsoft.VisualBasic.Strings.Len(source)) % (W));
                                        if (((this.CheckBox1.Checked) == (true)))
                                        {
                                            this.rng.Cells[(global::System.Object)i, (global::System.Object)SplitColumn].copy();
                                            this.rng2.Cells[(global::System.Object)SplitColumn, (global::System.Object)Index].PasteSpecial(global::Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats);
                                            this.rng2 = this.workSheet2.get_Range(rng2Address);
                                            this.workSheet2.Activate();
                                        }
                                        else
                                        {
                                            this.rng2.Cells[(global::System.Object)SplitColumn, (global::System.Object)Index].ClearFormats();
                                        }
                                        for (global::System.Int32 j = (SplitColumn) + (1), loopTo34 = c; j <= loopTo34; j++)
                                        {
                                            this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].value = this.rng.Cells[(global::System.Object)i, (global::System.Object)j].value;
                                            if (((this.CheckBox1.Checked) == (true)))
                                            {
                                                this.rng.Cells[(global::System.Object)i, (global::System.Object)j].copy();
                                                this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].PasteSpecial(global::Microsoft.Office.Interop.Excel.XlPasteType.xlPasteFormats);
                                                this.rng2 = this.workSheet2.get_Range(rng2Address);
                                                this.workSheet2.Activate();
                                            }
                                            else
                                            {
                                                this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].ClearFormats();
                                            }
                                        }
                                    }
                                }

                                excelApp.CutCopyMode = (global::Microsoft.Office.Interop.Excel.XlCutCopyMode)Conversions.ToInteger(false);

                            }

                        }
                    }

                    else
                    {

                        global::System.String rng2Address = this.rng2.get_Address();

                        var Arr = new global::System.Object[(this.rng.Rows.Count), (this.rng.Columns.Count)];

                        for (global::System.Int32 i = global::Microsoft.VisualBasic.Information.LBound(Arr, 1), loopTo35 = global::Microsoft.VisualBasic.Information.UBound(Arr, 1); i <= loopTo35; i++)
                        {
                            for (global::System.Int32 j = global::Microsoft.VisualBasic.Information.LBound(Arr, 2), loopTo36 = global::Microsoft.VisualBasic.Information.UBound(Arr, 2); j <= loopTo36; j++)
                                Arr[i, j] = this.rng.Cells[(global::System.Object)((i) + (1)), (global::System.Object)((j) + (1))].Value;
                        }

                        var FontNames = new global::System.String[(this.rng.Rows.Count), (this.rng.Columns.Count)];
                        var FontSizes = new global::System.Single[(this.rng.Rows.Count), (this.rng.Columns.Count)];
                        var FontBolds = new global::System.Boolean[(this.rng.Rows.Count), (this.rng.Columns.Count)];
                        var Fontitalics = new global::System.Boolean[(this.rng.Rows.Count), (this.rng.Columns.Count)];
                        var Red1s = new global::System.Int32[(this.rng.Rows.Count), (this.rng.Columns.Count)];
                        var Green1s = new global::System.Int32[(this.rng.Rows.Count), (this.rng.Columns.Count)];
                        var Blue1s = new global::System.Int32[(this.rng.Rows.Count), (this.rng.Columns.Count)];
                        var Red2s = new global::System.Int32[(this.rng.Rows.Count), (this.rng.Columns.Count)];
                        var Green2s = new global::System.Int32[(this.rng.Rows.Count), (this.rng.Columns.Count)];
                        var Blue2s = new global::System.Int32[(this.rng.Rows.Count), (this.rng.Columns.Count)];

                        for (global::System.Int32 i = global::Microsoft.VisualBasic.Information.LBound(FontSizes, 1), loopTo37 = global::Microsoft.VisualBasic.Information.UBound(FontSizes, 1); i <= loopTo37; i++)
                        {
                            for (global::System.Int32 j = global::Microsoft.VisualBasic.Information.LBound(FontSizes, 2), loopTo38 = global::Microsoft.VisualBasic.Information.UBound(FontSizes, 2); j <= loopTo38; j++)
                            {

                                global::Microsoft.Office.Interop.Excel.Range cell = (global::Microsoft.Office.Interop.Excel.Range)this.rng.Cells[(global::System.Object)((i) + (1)), (global::System.Object)((j) + (1))];

                                var font = cell.Font;
                                FontNames[i, j] = Conversions.ToString(font.Name);
                                FontBolds[i, j] = Conversions.ToBoolean(cell.Font.Bold);
                                Fontitalics[i, j] = Conversions.ToBoolean(cell.Font.Italic);


                                global::System.Single fontSize = global::System.Convert.ToSingle(font.Size);
                                FontSizes[i, j] = fontSize;

                                global::System.Int64 colorValue1 = Conversions.ToLong(cell.Interior.Color);
                                global::System.Int32 red1 = (global::System.Int32)((colorValue1) % (256L));
                                global::System.Int32 green1 = (global::System.Int32)(((((colorValue1) / (256L)))) % (256L));
                                global::System.Int32 blue1 = (global::System.Int32)((((((colorValue1) / (256L)) / (256L)))) % (256L));
                                Red1s[i, j] = red1;
                                Green1s[i, j] = green1;
                                Blue1s[i, j] = blue1;
                                global::System.Int64 colorValue2 = Conversions.ToLong(cell.Font.Color);
                                global::System.Int32 red2 = (global::System.Int32)((colorValue2) % (256L));
                                global::System.Int32 green2 = (global::System.Int32)(((((colorValue2) / (256L)))) % (256L));
                                global::System.Int32 blue2 = (global::System.Int32)((((((colorValue2) / (256L)) / (256L)))) % (256L));
                                Red2s[i, j] = red2;
                                Green2s[i, j] = green2;
                                Blue2s[i, j] = blue2;

                            }
                        }

                        if (((((X7) | (X8)) | (X9)) | (X10)))
                        {

                            if (X1)
                            {

                                global::System.Int32 Index = 0;
                                global::System.Int32 position;

                                for (global::System.Int32 i = 1, loopTo39 = r; i <= loopTo39; i++)
                                {
                                    global::System.String source = Conversions.ToString(Arr[(i) - (1), (SplitColumn) - (1)]);
                                    position = 1;
                                    for (global::System.Int32 k = 1, loopTo40 = global::Microsoft.VisualBasic.Strings.Len(source); k <= loopTo40; k++)
                                    {
                                        if (((global::Microsoft.VisualBasic.Strings.Mid(source, k, global::Microsoft.VisualBasic.Strings.Len(Separator)) ?? "") == (Separator ?? "")))
                                        {
                                            if ((((k) - (position)) > (0)))
                                            {
                                                Index = ((Index) + (1));
                                                for (global::System.Int32 j = 1, loopTo41 = (SplitColumn) - (1); j <= loopTo41; j++)
                                                {
                                                    this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].value = Arr[(i) - (1), (j) - (1)];
                                                    if (((this.CheckBox1.Checked) == (true)))
                                                    {
                                                        global::System.Int32 x = ((i) - (1));
                                                        global::System.Int32 y = ((j) - (1));

                                                        this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].Font.Name = FontNames[x, y];
                                                        this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].Font.Size = (global::System.Object)FontSizes[x, y];

                                                        if (FontBolds[x, y])
                                                            this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].Font.Bold = (global::System.Object)true;
                                                        if (Fontitalics[x, y])
                                                            this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].Font.Italic = (global::System.Object)true;


                                                        this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].Interior.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(Red1s[x, y], Green1s[x, y], Blue1s[x, y]);

                                                        this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].Font.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(Red2s[x, y], Green2s[x, y], Blue2s[x, y]);
                                                        this.workSheet2.Activate();
                                                    }
                                                    else
                                                    {
                                                        this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].ClearFormats();
                                                    }
                                                }

                                                this.rng2.Cells[(global::System.Object)Index, (global::System.Object)SplitColumn] = global::Microsoft.VisualBasic.Strings.Mid(source, position, (k) - (position));
                                                if (((this.CheckBox1.Checked) == (true)))
                                                {
                                                    global::System.Int32 x = ((i) - (1));
                                                    global::System.Int32 y = ((SplitColumn) - (1));

                                                    this.rng2.Cells[(global::System.Object)Index, (global::System.Object)SplitColumn].Font.Name = FontNames[x, y];
                                                    this.rng2.Cells[(global::System.Object)Index, (global::System.Object)SplitColumn].Font.Size = (global::System.Object)FontSizes[x, y];

                                                    if (FontBolds[x, y])
                                                        this.rng2.Cells[(global::System.Object)Index, (global::System.Object)SplitColumn].Font.Bold = (global::System.Object)true;
                                                    if (Fontitalics[x, y])
                                                        this.rng2.Cells[(global::System.Object)Index, (global::System.Object)SplitColumn].Font.Italic = (global::System.Object)true;


                                                    this.rng2.Cells[(global::System.Object)Index, (global::System.Object)SplitColumn].Interior.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(Red1s[x, y], Green1s[x, y], Blue1s[x, y]);

                                                    this.rng2.Cells[(global::System.Object)Index, (global::System.Object)SplitColumn].Font.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(Red2s[x, y], Green2s[x, y], Blue2s[x, y]);
                                                    this.workSheet2.Activate();
                                                }
                                                else
                                                {
                                                    this.rng2.Cells[(global::System.Object)Index, (global::System.Object)SplitColumn].ClearFormats();
                                                }

                                                for (global::System.Int32 j = (SplitColumn) + (1), loopTo42 = c; j <= loopTo42; j++)
                                                {
                                                    this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].value = Arr[(i) - (1), (j) - (1)];
                                                    if (((this.CheckBox1.Checked) == (true)))
                                                    {
                                                        global::System.Int32 x = ((i) - (1));
                                                        global::System.Int32 y = ((j) - (1));

                                                        this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].Font.Name = FontNames[x, y];
                                                        this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].Font.Size = (global::System.Object)FontSizes[x, y];

                                                        if (FontBolds[x, y])
                                                            this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].Font.Bold = (global::System.Object)true;
                                                        if (Fontitalics[x, y])
                                                            this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].Font.Italic = (global::System.Object)true;


                                                        this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].Interior.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(Red1s[x, y], Green1s[x, y], Blue1s[x, y]);

                                                        this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].Font.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(Red2s[x, y], Green2s[x, y], Blue2s[x, y]);
                                                        this.workSheet2.Activate();
                                                    }
                                                    else
                                                    {
                                                        this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].ClearFormats();
                                                    }
                                                }

                                            }
                                            position = ((k) + (global::Microsoft.VisualBasic.Strings.Len(Separator)));
                                        }
                                    }
                                    if (((position) <= (global::Microsoft.VisualBasic.Strings.Len(source))))
                                    {
                                        Index = ((Index) + (1));

                                        for (global::System.Int32 j = 1, loopTo43 = (SplitColumn) - (1); j <= loopTo43; j++)
                                        {
                                            this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].Value = Arr[(i) - (1), (j) - (1)];
                                            if (((this.CheckBox1.Checked) == (true)))
                                            {
                                                global::System.Int32 x = ((i) - (1));
                                                global::System.Int32 y = ((j) - (1));

                                                this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].Font.Name = FontNames[x, y];
                                                this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].Font.Size = (global::System.Object)FontSizes[x, y];

                                                if (FontBolds[x, y])
                                                    this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].Font.Bold = (global::System.Object)true;
                                                if (Fontitalics[x, y])
                                                    this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].Font.Italic = (global::System.Object)true;


                                                this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].Interior.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(Red1s[x, y], Green1s[x, y], Blue1s[x, y]);

                                                this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].Font.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(Red2s[x, y], Green2s[x, y], Blue2s[x, y]);
                                                this.workSheet2.Activate();
                                            }
                                            else
                                            {
                                                this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].ClearFormats();
                                            }
                                        }

                                        this.rng2.Cells[(global::System.Object)Index, (global::System.Object)SplitColumn] = global::Microsoft.VisualBasic.Strings.Mid(source, position, ((global::Microsoft.VisualBasic.Strings.Len(source)) - (position)) + (1));
                                        if (((this.CheckBox1.Checked) == (true)))
                                        {
                                            global::System.Int32 x = ((i) - (1));
                                            global::System.Int32 y = ((SplitColumn) - (1));

                                            this.rng2.Cells[(global::System.Object)Index, (global::System.Object)SplitColumn].Font.Name = FontNames[x, y];
                                            this.rng2.Cells[(global::System.Object)Index, (global::System.Object)SplitColumn].Font.Size = (global::System.Object)FontSizes[x, y];

                                            if (FontBolds[x, y])
                                                this.rng2.Cells[(global::System.Object)Index, (global::System.Object)SplitColumn].Font.Bold = (global::System.Object)true;
                                            if (Fontitalics[x, y])
                                                this.rng2.Cells[(global::System.Object)Index, (global::System.Object)SplitColumn].Font.Italic = (global::System.Object)true;


                                            this.rng2.Cells[(global::System.Object)Index, (global::System.Object)SplitColumn].Interior.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(Red1s[x, y], Green1s[x, y], Blue1s[x, y]);

                                            this.rng2.Cells[(global::System.Object)Index, (global::System.Object)SplitColumn].Font.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(Red2s[x, y], Green2s[x, y], Blue2s[x, y]);
                                            this.workSheet2.Activate();
                                        }
                                        else
                                        {
                                            this.rng2.Cells[(global::System.Object)Index, (global::System.Object)SplitColumn].ClearFormats();
                                        }

                                        for (global::System.Int32 j = (SplitColumn) + (1), loopTo44 = c; j <= loopTo44; j++)
                                        {
                                            this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].value = Arr[(i) - (1), (j) - (1)];
                                            if (((this.CheckBox1.Checked) == (true)))
                                            {
                                                global::System.Int32 x = ((i) - (1));
                                                global::System.Int32 y = ((j) - (1));

                                                this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].Font.Name = FontNames[x, y];
                                                this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].Font.Size = (global::System.Object)FontSizes[x, y];

                                                if (FontBolds[x, y])
                                                    this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].Font.Bold = (global::System.Object)true;
                                                if (Fontitalics[x, y])
                                                    this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].Font.Italic = (global::System.Object)true;


                                                this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].Interior.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(Red1s[x, y], Green1s[x, y], Blue1s[x, y]);

                                                this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].Font.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(Red2s[x, y], Green2s[x, y], Blue2s[x, y]);
                                                this.workSheet2.Activate();
                                            }
                                            else
                                            {
                                                this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].ClearFormats();
                                            }
                                        }

                                    }
                                }
                            }

                            else if (X2)
                            {

                                global::System.Int32 Index = 0;
                                global::System.Int32 position;

                                for (global::System.Int32 i = 1, loopTo45 = r; i <= loopTo45; i++)
                                {
                                    global::System.String source = Conversions.ToString(Arr[(i) - (1), (SplitColumn) - (1)]);
                                    position = 1;
                                    for (global::System.Int32 k = 1, loopTo46 = global::Microsoft.VisualBasic.Strings.Len(source); k <= loopTo46; k++)
                                    {
                                        if (((global::Microsoft.VisualBasic.Strings.Mid(source, k, global::Microsoft.VisualBasic.Strings.Len(Separator)) ?? "") == (Separator ?? "")))
                                        {
                                            if ((((k) - (position)) > (0)))
                                            {
                                                Index = ((Index) + (1));
                                                for (global::System.Int32 j = 1, loopTo47 = (SplitColumn) - (1); j <= loopTo47; j++)
                                                {
                                                    this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].value = Arr[(i) - (1), (j) - (1)];
                                                    if (((this.CheckBox1.Checked) == (true)))
                                                    {
                                                        global::System.Int32 x = ((i) - (1));
                                                        global::System.Int32 y = ((j) - (1));

                                                        this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].Font.Name = FontNames[x, y];
                                                        this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].Font.Size = (global::System.Object)FontSizes[x, y];

                                                        if (FontBolds[x, y])
                                                            this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].Font.Bold = (global::System.Object)true;
                                                        if (Fontitalics[x, y])
                                                            this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].Font.Italic = (global::System.Object)true;


                                                        this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].Interior.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(Red1s[x, y], Green1s[x, y], Blue1s[x, y]);

                                                        this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].Font.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(Red2s[x, y], Green2s[x, y], Blue2s[x, y]);

                                                        this.workSheet2.Activate();
                                                    }
                                                    else
                                                    {
                                                        this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].ClearFormats();
                                                    }
                                                }

                                                this.rng2.Cells[(global::System.Object)SplitColumn, (global::System.Object)Index] = global::Microsoft.VisualBasic.Strings.Mid(source, position, (k) - (position));
                                                if (((this.CheckBox1.Checked) == (true)))
                                                {
                                                    global::System.Int32 x = ((i) - (1));
                                                    global::System.Int32 y = ((SplitColumn) - (1));

                                                    this.rng2.Cells[(global::System.Object)SplitColumn, (global::System.Object)Index].Font.Name = FontNames[x, y];
                                                    this.rng2.Cells[(global::System.Object)SplitColumn, (global::System.Object)Index].Font.Size = (global::System.Object)FontSizes[x, y];

                                                    if (FontBolds[x, y])
                                                        this.rng2.Cells[(global::System.Object)SplitColumn, (global::System.Object)Index].Font.Bold = (global::System.Object)true;
                                                    if (Fontitalics[x, y])
                                                        this.rng2.Cells[(global::System.Object)SplitColumn, (global::System.Object)Index].Font.Italic = (global::System.Object)true;


                                                    this.rng2.Cells[(global::System.Object)SplitColumn, (global::System.Object)Index].Interior.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(Red1s[x, y], Green1s[x, y], Blue1s[x, y]);

                                                    this.rng2.Cells[(global::System.Object)SplitColumn, (global::System.Object)Index].Font.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(Red2s[x, y], Green2s[x, y], Blue2s[x, y]);
                                                    this.workSheet2.Activate();
                                                }
                                                else
                                                {
                                                    this.rng2.Cells[(global::System.Object)SplitColumn, (global::System.Object)Index].ClearFormats();
                                                }

                                                for (global::System.Int32 j = (SplitColumn) + (1), loopTo48 = c; j <= loopTo48; j++)
                                                {
                                                    this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].value = Arr[(i) - (1), (j) - (1)];
                                                    if (((this.CheckBox1.Checked) == (true)))
                                                    {
                                                        global::System.Int32 x = ((i) - (1));
                                                        global::System.Int32 y = ((j) - (1));

                                                        this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].Font.Name = FontNames[x, y];
                                                        this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].Font.Size = (global::System.Object)FontSizes[x, y];

                                                        if (FontBolds[x, y])
                                                            this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].Font.Bold = (global::System.Object)true;
                                                        if (Fontitalics[x, y])
                                                            this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].Font.Italic = (global::System.Object)true;


                                                        this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].Interior.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(Red1s[x, y], Green1s[x, y], Blue1s[x, y]);

                                                        this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].Font.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(Red2s[x, y], Green2s[x, y], Blue2s[x, y]);
                                                        this.workSheet2.Activate();
                                                    }
                                                    else
                                                    {
                                                        this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].ClearFormats();
                                                    }
                                                }
                                            }
                                            position = ((k) + (global::Microsoft.VisualBasic.Strings.Len(Separator)));
                                        }
                                    }
                                    if (((position) <= (global::Microsoft.VisualBasic.Strings.Len(source))))
                                    {
                                        Index = ((Index) + (1));
                                        for (global::System.Int32 j = 1, loopTo49 = (SplitColumn) - (1); j <= loopTo49; j++)
                                        {
                                            this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].Value = Arr[(i) - (1), (j) - (1)];
                                            if (((this.CheckBox1.Checked) == (true)))
                                            {
                                                global::System.Int32 x = ((i) - (1));
                                                global::System.Int32 y = ((j) - (1));

                                                this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].Font.Name = FontNames[x, y];
                                                this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].Font.Size = (global::System.Object)FontSizes[x, y];

                                                if (FontBolds[x, y])
                                                    this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].Font.Bold = (global::System.Object)true;
                                                if (Fontitalics[x, y])
                                                    this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].Font.Italic = (global::System.Object)true;


                                                this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].Interior.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(Red1s[x, y], Green1s[x, y], Blue1s[x, y]);

                                                this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].Font.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(Red2s[x, y], Green2s[x, y], Blue2s[x, y]);
                                                this.workSheet2.Activate();
                                            }
                                            else
                                            {
                                                this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].ClearFormats();
                                            }
                                        }

                                        this.rng2.Cells[(global::System.Object)SplitColumn, (global::System.Object)Index] = global::Microsoft.VisualBasic.Strings.Mid(source, position, ((global::Microsoft.VisualBasic.Strings.Len(source)) - (position)) + (1));
                                        if (((this.CheckBox1.Checked) == (true)))
                                        {
                                            global::System.Int32 x = ((i) - (1));
                                            global::System.Int32 y = ((SplitColumn) - (1));

                                            this.rng2.Cells[(global::System.Object)SplitColumn, (global::System.Object)Index].Font.Name = FontNames[x, y];
                                            this.rng2.Cells[(global::System.Object)SplitColumn, (global::System.Object)Index].Font.Size = (global::System.Object)FontSizes[x, y];

                                            if (FontBolds[x, y])
                                                this.rng2.Cells[(global::System.Object)SplitColumn, (global::System.Object)Index].Font.Bold = (global::System.Object)true;
                                            if (Fontitalics[x, y])
                                                this.rng2.Cells[(global::System.Object)SplitColumn, (global::System.Object)Index].Font.Italic = (global::System.Object)true;


                                            this.rng2.Cells[(global::System.Object)SplitColumn, (global::System.Object)Index].Interior.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(Red1s[x, y], Green1s[x, y], Blue1s[x, y]);

                                            this.rng2.Cells[(global::System.Object)SplitColumn, (global::System.Object)Index].Font.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(Red2s[x, y], Green2s[x, y], Blue2s[x, y]);
                                            this.workSheet2.Activate();
                                        }
                                        else
                                        {
                                            this.rng2.Cells[(global::System.Object)SplitColumn, (global::System.Object)Index].ClearFormats();
                                        }

                                        for (global::System.Int32 j = (SplitColumn) + (1), loopTo50 = c; j <= loopTo50; j++)
                                        {
                                            this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].value = Arr[(i) - (1), (j) - (1)];
                                            if (((this.CheckBox1.Checked) == (true)))
                                            {
                                                global::System.Int32 x = ((i) - (1));
                                                global::System.Int32 y = ((j) - (1));

                                                this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].Font.Name = FontNames[x, y];
                                                this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].Font.Size = (global::System.Object)FontSizes[x, y];

                                                if (FontBolds[x, y])
                                                    this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].Font.Bold = (global::System.Object)true;
                                                if (Fontitalics[x, y])
                                                    this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].Font.Italic = (global::System.Object)true;


                                                this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].Interior.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(Red1s[x, y], Green1s[x, y], Blue1s[x, y]);

                                                this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].Font.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(Red2s[x, y], Green2s[x, y], Blue2s[x, y]);
                                                this.workSheet2.Activate();
                                            }
                                            else
                                            {
                                                this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].ClearFormats();
                                            }
                                        }
                                    }
                                }
                                excelApp.CutCopyMode = (global::Microsoft.Office.Interop.Excel.XlCutCopyMode)Conversions.ToInteger(false);
                            }
                        }

                        else if (X3)
                        {

                            if (X1)
                            {

                                global::System.Int32 Index = 0;

                                for (global::System.Int32 i = 1, loopTo51 = r; i <= loopTo51; i++)
                                {

                                    global::System.String source = Conversions.ToString(Arr[(i) - (1), (SplitColumn) - (1)]);
                                    var NumberText = new global::System.String[2];
                                    NumberText = (global::System.String[])this.SeparateNumberText(source);
                                    global::System.String Number = NumberText[0];
                                    global::System.String Text = NumberText[1];

                                    Index = ((Index) + (1));
                                    for (global::System.Int32 j = 1, loopTo52 = (SplitColumn) - (1); j <= loopTo52; j++)
                                    {
                                        this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].value = Arr[(i) - (1), (j) - (1)];
                                        if (((this.CheckBox1.Checked) == (true)))
                                        {
                                            global::System.Int32 x = ((i) - (1));
                                            global::System.Int32 y = ((j) - (1));

                                            this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].Font.Name = FontNames[x, y];
                                            this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].Font.Size = (global::System.Object)FontSizes[x, y];

                                            if (FontBolds[x, y])
                                                this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].Font.Bold = (global::System.Object)true;
                                            if (Fontitalics[x, y])
                                                this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].Font.Italic = (global::System.Object)true;


                                            this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].Interior.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(Red1s[x, y], Green1s[x, y], Blue1s[x, y]);

                                            this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].Font.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(Red2s[x, y], Green2s[x, y], Blue2s[x, y]);
                                            this.workSheet2.Activate();
                                        }
                                        else
                                        {
                                            this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].ClearFormats();
                                        }
                                    }

                                    this.rng2.Cells[(global::System.Object)Index, (global::System.Object)SplitColumn].value = Number;
                                    if (((this.CheckBox1.Checked) == (true)))
                                    {
                                        global::System.Int32 x = ((i) - (1));
                                        global::System.Int32 y = ((SplitColumn) - (1));

                                        this.rng2.Cells[(global::System.Object)Index, (global::System.Object)SplitColumn].Font.Name = FontNames[x, y];
                                        this.rng2.Cells[(global::System.Object)Index, (global::System.Object)SplitColumn].Font.Size = (global::System.Object)FontSizes[x, y];

                                        if (FontBolds[x, y])
                                            this.rng2.Cells[(global::System.Object)Index, (global::System.Object)SplitColumn].Font.Bold = (global::System.Object)true;
                                        if (Fontitalics[x, y])
                                            this.rng2.Cells[(global::System.Object)Index, (global::System.Object)SplitColumn].Font.Italic = (global::System.Object)true;


                                        this.rng2.Cells[(global::System.Object)Index, (global::System.Object)SplitColumn].Interior.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(Red1s[x, y], Green1s[x, y], Blue1s[x, y]);

                                        this.rng2.Cells[(global::System.Object)Index, (global::System.Object)SplitColumn].Font.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(Red2s[x, y], Green2s[x, y], Blue2s[x, y]);
                                        this.workSheet2.Activate();
                                    }
                                    else
                                    {
                                        this.rng2.Cells[(global::System.Object)Index, (global::System.Object)SplitColumn].ClearFormats();
                                    }

                                    for (global::System.Int32 j = (SplitColumn) + (1), loopTo53 = c; j <= loopTo53; j++)
                                    {
                                        this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].value = Arr[(i) - (1), (j) - (1)];
                                        if (((this.CheckBox1.Checked) == (true)))
                                        {
                                            global::System.Int32 x = ((i) - (1));
                                            global::System.Int32 y = ((j) - (1));

                                            this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].Font.Name = FontNames[x, y];
                                            this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].Font.Size = (global::System.Object)FontSizes[x, y];

                                            if (FontBolds[x, y])
                                                this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].Font.Bold = (global::System.Object)true;
                                            if (Fontitalics[x, y])
                                                this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].Font.Italic = (global::System.Object)true;


                                            this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].Interior.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(Red1s[x, y], Green1s[x, y], Blue1s[x, y]);

                                            this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].Font.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(Red2s[x, y], Green2s[x, y], Blue2s[x, y]);
                                            this.workSheet2.Activate();
                                        }
                                        else
                                        {
                                            this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].ClearFormats();
                                        }
                                    }

                                    Index = ((Index) + (1));
                                    for (global::System.Int32 j = 1, loopTo54 = (SplitColumn) - (1); j <= loopTo54; j++)
                                    {
                                        this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].value = Arr[(i) - (1), (j) - (1)];
                                        if (((this.CheckBox1.Checked) == (true)))
                                        {
                                            global::System.Int32 x = ((i) - (1));
                                            global::System.Int32 y = ((j) - (1));

                                            this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].Font.Name = FontNames[x, y];
                                            this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].Font.Size = (global::System.Object)FontSizes[x, y];

                                            if (FontBolds[x, y])
                                                this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].Font.Bold = (global::System.Object)true;
                                            if (Fontitalics[x, y])
                                                this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].Font.Italic = (global::System.Object)true;


                                            this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].Interior.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(Red1s[x, y], Green1s[x, y], Blue1s[x, y]);

                                            this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].Font.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(Red2s[x, y], Green2s[x, y], Blue2s[x, y]);
                                            this.workSheet2.Activate();
                                        }
                                        else
                                        {
                                            this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].ClearFormats();
                                        }
                                    }

                                    this.rng2.Cells[(global::System.Object)Index, (global::System.Object)SplitColumn].value = Text;
                                    if (((this.CheckBox1.Checked) == (true)))
                                    {
                                        global::System.Int32 x = ((i) - (1));
                                        global::System.Int32 y = ((SplitColumn) - (1));

                                        this.rng2.Cells[(global::System.Object)Index, (global::System.Object)SplitColumn].Font.Name = FontNames[x, y];
                                        this.rng2.Cells[(global::System.Object)Index, (global::System.Object)SplitColumn].Font.Size = (global::System.Object)FontSizes[x, y];

                                        if (FontBolds[x, y])
                                            this.rng2.Cells[(global::System.Object)Index, (global::System.Object)SplitColumn].Font.Bold = (global::System.Object)true;
                                        if (Fontitalics[x, y])
                                            this.rng2.Cells[(global::System.Object)Index, (global::System.Object)SplitColumn].Font.Italic = (global::System.Object)true;


                                        this.rng2.Cells[(global::System.Object)Index, (global::System.Object)SplitColumn].Interior.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(Red1s[x, y], Green1s[x, y], Blue1s[x, y]);

                                        this.rng2.Cells[(global::System.Object)Index, (global::System.Object)SplitColumn].Font.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(Red2s[x, y], Green2s[x, y], Blue2s[x, y]);
                                        this.workSheet2.Activate();
                                    }
                                    else
                                    {
                                        this.rng2.Cells[(global::System.Object)Index, (global::System.Object)SplitColumn].ClearFormats();
                                    }

                                    for (global::System.Int32 j = (SplitColumn) + (1), loopTo55 = c; j <= loopTo55; j++)
                                    {
                                        this.rng2.Cells[(global::System.Object)i, (global::System.Object)j].value = Arr[(i) - (1), (j) - (1)];
                                        if (((this.CheckBox1.Checked) == (true)))
                                        {
                                            global::System.Int32 x = ((i) - (1));
                                            global::System.Int32 y = ((j) - (1));

                                            this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].Font.Name = FontNames[x, y];
                                            this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].Font.Size = (global::System.Object)FontSizes[x, y];

                                            if (FontBolds[x, y])
                                                this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].Font.Bold = (global::System.Object)true;
                                            if (Fontitalics[x, y])
                                                this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].Font.Italic = (global::System.Object)true;


                                            this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].Interior.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(Red1s[x, y], Green1s[x, y], Blue1s[x, y]);

                                            this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].Font.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(Red2s[x, y], Green2s[x, y], Blue2s[x, y]);
                                            this.workSheet2.Activate();
                                        }
                                        else
                                        {
                                            this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].ClearFormats();
                                        }
                                    }

                                }
                                excelApp.CutCopyMode = (global::Microsoft.Office.Interop.Excel.XlCutCopyMode)Conversions.ToInteger(false);
                            }

                            else if (X2)
                            {

                                global::System.Int32 Index = 0;

                                for (global::System.Int32 i = 1, loopTo56 = r; i <= loopTo56; i++)
                                {

                                    global::System.String source = Conversions.ToString(Arr[(i) - (1), (SplitColumn) - (1)]);
                                    var NumberText = new global::System.String[2];
                                    NumberText = (global::System.String[])this.SeparateNumberText(source);
                                    global::System.String Number = NumberText[0];
                                    global::System.String Text = NumberText[1];

                                    Index = ((Index) + (1));
                                    for (global::System.Int32 j = 1, loopTo57 = (c) - (1); j <= loopTo57; j++)
                                    {
                                        this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].value = Arr[(i) - (1), (j) - (1)];
                                        if (((this.CheckBox1.Checked) == (true)))
                                        {
                                            global::System.Int32 x = ((i) - (1));
                                            global::System.Int32 y = ((j) - (1));

                                            this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].Font.Name = FontNames[x, y];
                                            this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].Font.Size = (global::System.Object)FontSizes[x, y];

                                            if (FontBolds[x, y])
                                                this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].Font.Bold = (global::System.Object)true;
                                            if (Fontitalics[x, y])
                                                this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].Font.Italic = (global::System.Object)true;


                                            this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].Interior.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(Red1s[x, y], Green1s[x, y], Blue1s[x, y]);

                                            this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].Font.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(Red2s[x, y], Green2s[x, y], Blue2s[x, y]);
                                            this.workSheet2.Activate();
                                        }
                                        else
                                        {
                                            this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].ClearFormats();
                                        }
                                    }

                                    this.rng2.Cells[(global::System.Object)SplitColumn, (global::System.Object)Index].value = Number;
                                    if (((this.CheckBox1.Checked) == (true)))
                                    {
                                        global::System.Int32 x = ((i) - (1));
                                        global::System.Int32 y = ((SplitColumn) - (1));

                                        this.rng2.Cells[(global::System.Object)SplitColumn, (global::System.Object)Index].Font.Name = FontNames[x, y];
                                        this.rng2.Cells[(global::System.Object)SplitColumn, (global::System.Object)Index].Font.Size = (global::System.Object)FontSizes[x, y];

                                        if (FontBolds[x, y])
                                            this.rng2.Cells[(global::System.Object)SplitColumn, (global::System.Object)Index].Font.Bold = (global::System.Object)true;
                                        if (Fontitalics[x, y])
                                            this.rng2.Cells[(global::System.Object)SplitColumn, (global::System.Object)Index].Font.Italic = (global::System.Object)true;


                                        this.rng2.Cells[(global::System.Object)SplitColumn, (global::System.Object)Index].Interior.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(Red1s[x, y], Green1s[x, y], Blue1s[x, y]);

                                        this.rng2.Cells[(global::System.Object)SplitColumn, (global::System.Object)Index].Font.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(Red2s[x, y], Green2s[x, y], Blue2s[x, y]);
                                        this.workSheet2.Activate();
                                    }
                                    else
                                    {
                                        this.rng2.Cells[(global::System.Object)SplitColumn, (global::System.Object)Index].ClearFormats();
                                    }

                                    for (global::System.Int32 j = (SplitColumn) + (1), loopTo58 = c; j <= loopTo58; j++)
                                    {
                                        this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].value = Arr[(i) - (1), (j) - (1)];
                                        if (((this.CheckBox1.Checked) == (true)))
                                        {
                                            global::System.Int32 x = ((i) - (1));
                                            global::System.Int32 y = ((j) - (1));

                                            this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].Font.Name = FontNames[x, y];
                                            this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].Font.Size = (global::System.Object)FontSizes[x, y];

                                            if (FontBolds[x, y])
                                                this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].Font.Bold = (global::System.Object)true;
                                            if (Fontitalics[x, y])
                                                this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].Font.Italic = (global::System.Object)true;


                                            this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].Interior.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(Red1s[x, y], Green1s[x, y], Blue1s[x, y]);

                                            this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].Font.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(Red2s[x, y], Green2s[x, y], Blue2s[x, y]);
                                            this.workSheet2.Activate();
                                        }
                                        else
                                        {
                                            this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].ClearFormats();
                                        }
                                    }

                                    Index = ((Index) + (1));
                                    for (global::System.Int32 j = 1, loopTo59 = (SplitColumn) - (1); j <= loopTo59; j++)
                                    {
                                        this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].value = Arr[(i) - (1), (j) - (1)];
                                        if (((this.CheckBox1.Checked) == (true)))
                                        {
                                            global::System.Int32 x = ((i) - (1));
                                            global::System.Int32 y = ((j) - (1));

                                            this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].Font.Name = FontNames[x, y];
                                            this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].Font.Size = (global::System.Object)FontSizes[x, y];

                                            if (FontBolds[x, y])
                                                this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].Font.Bold = (global::System.Object)true;
                                            if (Fontitalics[x, y])
                                                this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].Font.Italic = (global::System.Object)true;


                                            this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].Interior.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(Red1s[x, y], Green1s[x, y], Blue1s[x, y]);

                                            this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].Font.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(Red2s[x, y], Green2s[x, y], Blue2s[x, y]);
                                            this.workSheet2.Activate();
                                        }
                                        else
                                        {
                                            this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].ClearFormats();
                                        }
                                    }

                                    this.rng2.Cells[(global::System.Object)SplitColumn, (global::System.Object)Index].value = Text;
                                    if (((this.CheckBox1.Checked) == (true)))
                                    {
                                        global::System.Int32 x = ((i) - (1));
                                        global::System.Int32 y = ((SplitColumn) - (1));

                                        this.rng2.Cells[(global::System.Object)SplitColumn, (global::System.Object)Index].Font.Name = FontNames[x, y];
                                        this.rng2.Cells[(global::System.Object)SplitColumn, (global::System.Object)Index].Font.Size = (global::System.Object)FontSizes[x, y];

                                        if (FontBolds[x, y])
                                            this.rng2.Cells[(global::System.Object)SplitColumn, (global::System.Object)Index].Font.Bold = (global::System.Object)true;
                                        if (Fontitalics[x, y])
                                            this.rng2.Cells[(global::System.Object)SplitColumn, (global::System.Object)Index].Font.Italic = (global::System.Object)true;


                                        this.rng2.Cells[(global::System.Object)SplitColumn, (global::System.Object)Index].Interior.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(Red1s[x, y], Green1s[x, y], Blue1s[x, y]);

                                        this.rng2.Cells[(global::System.Object)SplitColumn, (global::System.Object)Index].Font.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(Red2s[x, y], Green2s[x, y], Blue2s[x, y]);
                                        this.workSheet2.Activate();
                                    }
                                    else
                                    {
                                        this.rng2.Cells[(global::System.Object)SplitColumn, (global::System.Object)Index].ClearFormats();
                                    }

                                    for (global::System.Int32 j = (SplitColumn) + (1), loopTo60 = c; j <= loopTo60; j++)
                                    {
                                        this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].value = Arr[(i) - (1), (j) - (1)];
                                        if (((this.CheckBox1.Checked) == (true)))
                                        {
                                            global::System.Int32 x = ((i) - (1));
                                            global::System.Int32 y = ((j) - (1));

                                            this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].Font.Name = FontNames[x, y];
                                            this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].Font.Size = (global::System.Object)FontSizes[x, y];

                                            if (FontBolds[x, y])
                                                this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].Font.Bold = (global::System.Object)true;
                                            if (Fontitalics[x, y])
                                                this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].Font.Italic = (global::System.Object)true;


                                            this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].Interior.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(Red1s[x, y], Green1s[x, y], Blue1s[x, y]);

                                            this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].Font.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(Red2s[x, y], Green2s[x, y], Blue2s[x, y]);
                                            this.workSheet2.Activate();
                                        }
                                        else
                                        {
                                            this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].ClearFormats();
                                        }
                                    }
                                }

                            }
                        }

                        else if (X11)
                        {

                            global::System.Int32 W;

                            if (string.IsNullOrEmpty(this.TextBox3.Text))
                            {
                                W = 1;
                            }
                            else
                            {
                                W = Conversions.ToInteger(global::Microsoft.VisualBasic.Conversion.Int(this.TextBox3.Text));
                            }

                            if (X1)
                            {

                                global::System.Int32 Index = 0;

                                for (global::System.Int32 i = 1, loopTo61 = r; i <= loopTo61; i++)
                                {
                                    global::System.String source = Conversions.ToString(Arr[(i) - (1), (SplitColumn) - (1)]);
                                    for (global::System.Double k = 1d, loopTo62 = global::Microsoft.VisualBasic.Conversion.Int((global::System.Double)(global::Microsoft.VisualBasic.Strings.Len(source)) / (global::System.Double)(W)); k <= loopTo62; k++)
                                    {
                                        Index = ((Index) + (1));
                                        for (global::System.Int32 j = 1, loopTo63 = (SplitColumn) - (1); j <= loopTo63; j++)
                                        {
                                            this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].value = Arr[(i) - (1), (j) - (1)];
                                            if (((this.CheckBox1.Checked) == (true)))
                                            {
                                                global::System.Int32 x = ((i) - (1));
                                                global::System.Int32 y = ((j) - (1));

                                                this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].Font.Name = FontNames[x, y];
                                                this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].Font.Size = (global::System.Object)FontSizes[x, y];

                                                if (FontBolds[x, y])
                                                    this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].Font.Bold = (global::System.Object)true;
                                                if (Fontitalics[x, y])
                                                    this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].Font.Italic = (global::System.Object)true;


                                                this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].Interior.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(Red1s[x, y], Green1s[x, y], Blue1s[x, y]);

                                                this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].Font.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(Red2s[x, y], Green2s[x, y], Blue2s[x, y]);
                                                this.workSheet2.Activate();
                                            }
                                            else
                                            {
                                                this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].ClearFormats();
                                            }
                                        }
                                        this.rng2.Cells[(global::System.Object)Index, (global::System.Object)SplitColumn].value = global::Microsoft.VisualBasic.Strings.Mid(source, (global::System.Int32)Math.Round(((((global::System.Double)(W) * ((((k) - (1d))))))) + (1d)), W);
                                        if (((this.CheckBox1.Checked) == (true)))
                                        {
                                            global::System.Int32 x = ((i) - (1));
                                            global::System.Int32 y = ((SplitColumn) - (1));

                                            this.rng2.Cells[(global::System.Object)Index, (global::System.Object)SplitColumn].Font.Name = FontNames[x, y];
                                            this.rng2.Cells[(global::System.Object)Index, (global::System.Object)SplitColumn].Font.Size = (global::System.Object)FontSizes[x, y];

                                            if (FontBolds[x, y])
                                                this.rng2.Cells[(global::System.Object)Index, (global::System.Object)SplitColumn].Font.Bold = (global::System.Object)true;
                                            if (Fontitalics[x, y])
                                                this.rng2.Cells[(global::System.Object)Index, (global::System.Object)SplitColumn].Font.Italic = (global::System.Object)true;


                                            this.rng2.Cells[(global::System.Object)Index, (global::System.Object)SplitColumn].Interior.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(Red1s[x, y], Green1s[x, y], Blue1s[x, y]);

                                            this.rng2.Cells[(global::System.Object)Index, (global::System.Object)SplitColumn].Font.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(Red2s[x, y], Green2s[x, y], Blue2s[x, y]);
                                            this.workSheet2.Activate();
                                        }
                                        else
                                        {
                                            this.rng2.Cells[(global::System.Object)Index, (global::System.Object)SplitColumn].ClearFormats();
                                        }
                                        for (global::System.Int32 j = (SplitColumn) + (1), loopTo64 = c; j <= loopTo64; j++)
                                        {
                                            this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].value = Arr[(i) - (1), (j) - (1)];
                                            if (((this.CheckBox1.Checked) == (true)))
                                            {
                                                global::System.Int32 x = ((i) - (1));
                                                global::System.Int32 y = ((j) - (1));

                                                this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].Font.Name = FontNames[x, y];
                                                this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].Font.Size = (global::System.Object)FontSizes[x, y];

                                                if (FontBolds[x, y])
                                                    this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].Font.Bold = (global::System.Object)true;
                                                if (Fontitalics[x, y])
                                                    this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].Font.Italic = (global::System.Object)true;


                                                this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].Interior.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(Red1s[x, y], Green1s[x, y], Blue1s[x, y]);

                                                this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].Font.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(Red2s[x, y], Green2s[x, y], Blue2s[x, y]);
                                                this.workSheet2.Activate();
                                            }
                                            else
                                            {
                                                this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].ClearFormats();
                                            }
                                        }
                                    }
                                    if ((((global::Microsoft.VisualBasic.Strings.Len(source)) % (W)) != (0)))
                                    {
                                        Index = ((Index) + (1));
                                        for (global::System.Int32 j = 1, loopTo65 = (SplitColumn) - (1); j <= loopTo65; j++)
                                        {
                                            this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].value = Arr[(i) - (1), (j) - (1)];
                                            if (((this.CheckBox1.Checked) == (true)))
                                            {
                                                global::System.Int32 x = ((i) - (1));
                                                global::System.Int32 y = ((j) - (1));

                                                this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].Font.Name = FontNames[x, y];
                                                this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].Font.Size = (global::System.Object)FontSizes[x, y];

                                                if (FontBolds[x, y])
                                                    this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].Font.Bold = (global::System.Object)true;
                                                if (Fontitalics[x, y])
                                                    this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].Font.Italic = (global::System.Object)true;


                                                this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].Interior.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(Red1s[x, y], Green1s[x, y], Blue1s[x, y]);

                                                this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].Font.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(Red2s[x, y], Green2s[x, y], Blue2s[x, y]);
                                                this.workSheet2.Activate();
                                            }
                                            else
                                            {
                                                this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].ClearFormats();
                                            }
                                        }
                                        this.rng2.Cells[(global::System.Object)Index, (global::System.Object)SplitColumn].value = global::Microsoft.VisualBasic.Strings.Mid(source, ((global::Microsoft.VisualBasic.Strings.Len(source)) - ((((global::Microsoft.VisualBasic.Strings.Len(source)) % (W))))) + (1), (global::Microsoft.VisualBasic.Strings.Len(source)) % (W));
                                        if (((this.CheckBox1.Checked) == (true)))
                                        {
                                            global::System.Int32 x = ((i) - (1));
                                            global::System.Int32 y = ((SplitColumn) - (1));

                                            this.rng2.Cells[(global::System.Object)Index, (global::System.Object)SplitColumn].Font.Name = FontNames[x, y];
                                            this.rng2.Cells[(global::System.Object)Index, (global::System.Object)SplitColumn].Font.Size = (global::System.Object)FontSizes[x, y];

                                            if (FontBolds[x, y])
                                                this.rng2.Cells[(global::System.Object)Index, (global::System.Object)SplitColumn].Font.Bold = (global::System.Object)true;
                                            if (Fontitalics[x, y])
                                                this.rng2.Cells[(global::System.Object)Index, (global::System.Object)SplitColumn].Font.Italic = (global::System.Object)true;


                                            this.rng2.Cells[(global::System.Object)Index, (global::System.Object)SplitColumn].Interior.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(Red1s[x, y], Green1s[x, y], Blue1s[x, y]);

                                            this.rng2.Cells[(global::System.Object)Index, (global::System.Object)SplitColumn].Font.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(Red2s[x, y], Green2s[x, y], Blue2s[x, y]);
                                            this.workSheet2.Activate();
                                        }
                                        else
                                        {
                                            this.rng2.Cells[(global::System.Object)Index, (global::System.Object)SplitColumn].ClearFormats();
                                        }
                                        for (global::System.Int32 j = (SplitColumn) + (1), loopTo66 = c; j <= loopTo66; j++)
                                        {
                                            this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].value = Arr[(i) - (1), (j) - (1)];
                                            if (((this.CheckBox1.Checked) == (true)))
                                            {
                                                global::System.Int32 x = ((i) - (1));
                                                global::System.Int32 y = ((j) - (1));

                                                this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].Font.Name = FontNames[x, y];
                                                this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].Font.Size = (global::System.Object)FontSizes[x, y];

                                                if (FontBolds[x, y])
                                                    this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].Font.Bold = (global::System.Object)true;
                                                if (Fontitalics[x, y])
                                                    this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].Font.Italic = (global::System.Object)true;


                                                this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].Interior.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(Red1s[x, y], Green1s[x, y], Blue1s[x, y]);

                                                this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].Font.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(Red2s[x, y], Green2s[x, y], Blue2s[x, y]);
                                                this.workSheet2.Activate();
                                            }
                                            else
                                            {
                                                this.rng2.Cells[(global::System.Object)Index, (global::System.Object)j].ClearFormats();
                                            }
                                        }
                                    }
                                }
                            }

                            else if (X2)
                            {

                                global::System.Int32 Index = 0;

                                for (global::System.Int32 i = 1, loopTo67 = r; i <= loopTo67; i++)
                                {
                                    global::System.String source = Conversions.ToString(Arr[(i) - (1), (SplitColumn) - (1)]);
                                    for (global::System.Double k = 1d, loopTo68 = global::Microsoft.VisualBasic.Conversion.Int((global::System.Double)(global::Microsoft.VisualBasic.Strings.Len(source)) / (global::System.Double)(W)); k <= loopTo68; k++)
                                    {
                                        Index = ((Index) + (1));
                                        for (global::System.Int32 j = 1, loopTo69 = (SplitColumn) - (1); j <= loopTo69; j++)
                                        {
                                            this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].value = Arr[(i) - (1), (j) - (1)];
                                            if (((this.CheckBox1.Checked) == (true)))
                                            {
                                                global::System.Int32 x = ((i) - (1));
                                                global::System.Int32 y = ((j) - (1));

                                                this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].Font.Name = FontNames[x, y];
                                                this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].Font.Size = (global::System.Object)FontSizes[x, y];

                                                if (FontBolds[x, y])
                                                    this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].Font.Bold = (global::System.Object)true;
                                                if (Fontitalics[x, y])
                                                    this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].Font.Italic = (global::System.Object)true;


                                                this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].Interior.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(Red1s[x, y], Green1s[x, y], Blue1s[x, y]);

                                                this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].Font.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(Red2s[x, y], Green2s[x, y], Blue2s[x, y]);
                                                this.workSheet2.Activate();
                                            }
                                            else
                                            {
                                                this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].ClearFormats();
                                            }
                                        }
                                        this.rng2.Cells[(global::System.Object)SplitColumn, (global::System.Object)Index].value = global::Microsoft.VisualBasic.Strings.Mid(source, (global::System.Int32)Math.Round(((((global::System.Double)(W) * ((((k) - (1d))))))) + (1d)), W);
                                        if (((this.CheckBox1.Checked) == (true)))
                                        {
                                            global::System.Int32 x = ((i) - (1));
                                            global::System.Int32 y = ((SplitColumn) - (1));

                                            this.rng2.Cells[(global::System.Object)SplitColumn, (global::System.Object)Index].Font.Name = FontNames[x, y];
                                            this.rng2.Cells[(global::System.Object)SplitColumn, (global::System.Object)Index].Font.Size = (global::System.Object)FontSizes[x, y];

                                            if (FontBolds[x, y])
                                                this.rng2.Cells[(global::System.Object)SplitColumn, (global::System.Object)Index].Font.Bold = (global::System.Object)true;
                                            if (Fontitalics[x, y])
                                                this.rng2.Cells[(global::System.Object)SplitColumn, (global::System.Object)Index].Font.Italic = (global::System.Object)true;


                                            this.rng2.Cells[(global::System.Object)SplitColumn, (global::System.Object)Index].Interior.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(Red1s[x, y], Green1s[x, y], Blue1s[x, y]);

                                            this.rng2.Cells[(global::System.Object)SplitColumn, (global::System.Object)Index].Font.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(Red2s[x, y], Green2s[x, y], Blue2s[x, y]);
                                            this.workSheet2.Activate();
                                        }
                                        else
                                        {
                                            this.rng2.Cells[(global::System.Object)SplitColumn, (global::System.Object)Index].ClearFormats();
                                        }
                                        for (global::System.Int32 j = (SplitColumn) + (1), loopTo70 = c; j <= loopTo70; j++)
                                        {
                                            this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].value = Arr[(i) - (1), (j) - (1)];
                                            if (((this.CheckBox1.Checked) == (true)))
                                            {
                                                global::System.Int32 x = ((i) - (1));
                                                global::System.Int32 y = ((j) - (1));

                                                this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].Font.Name = FontNames[x, y];
                                                this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].Font.Size = (global::System.Object)FontSizes[x, y];

                                                if (FontBolds[x, y])
                                                    this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].Font.Bold = (global::System.Object)true;
                                                if (Fontitalics[x, y])
                                                    this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].Font.Italic = (global::System.Object)true;


                                                this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].Interior.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(Red1s[x, y], Green1s[x, y], Blue1s[x, y]);

                                                this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].Font.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(Red2s[x, y], Green2s[x, y], Blue2s[x, y]);
                                                this.workSheet2.Activate();
                                            }
                                            else
                                            {
                                                this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].ClearFormats();
                                            }
                                        }
                                    }
                                    if ((((global::Microsoft.VisualBasic.Strings.Len(source)) % (W)) != (0)))
                                    {
                                        Index = ((Index) + (1));
                                        for (global::System.Int32 j = 1, loopTo71 = (SplitColumn) - (1); j <= loopTo71; j++)
                                        {
                                            this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].value = Arr[(i) - (1), (j) - (1)];
                                            if (((this.CheckBox1.Checked) == (true)))
                                            {
                                                global::System.Int32 x = ((i) - (1));
                                                global::System.Int32 y = ((j) - (1));

                                                this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].Font.Name = FontNames[x, y];
                                                this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].Font.Size = (global::System.Object)FontSizes[x, y];

                                                if (FontBolds[x, y])
                                                    this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].Font.Bold = (global::System.Object)true;
                                                if (Fontitalics[x, y])
                                                    this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].Font.Italic = (global::System.Object)true;


                                                this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].Interior.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(Red1s[x, y], Green1s[x, y], Blue1s[x, y]);

                                                this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].Font.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(Red2s[x, y], Green2s[x, y], Blue2s[x, y]);
                                                this.workSheet2.Activate();
                                            }
                                            else
                                            {
                                                this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].ClearFormats();
                                            }
                                        }
                                        this.rng2.Cells[(global::System.Object)SplitColumn, (global::System.Object)Index].value = global::Microsoft.VisualBasic.Strings.Mid(source, ((global::Microsoft.VisualBasic.Strings.Len(source)) - ((((global::Microsoft.VisualBasic.Strings.Len(source)) % (W))))) + (1), (global::Microsoft.VisualBasic.Strings.Len(source)) % (W));
                                        if (((this.CheckBox1.Checked) == (true)))
                                        {
                                            global::System.Int32 x = ((i) - (1));
                                            global::System.Int32 y = ((SplitColumn) - (1));

                                            this.rng2.Cells[(global::System.Object)SplitColumn, (global::System.Object)Index].Font.Name = FontNames[x, y];
                                            this.rng2.Cells[(global::System.Object)SplitColumn, (global::System.Object)Index].Font.Size = (global::System.Object)FontSizes[x, y];

                                            if (FontBolds[x, y])
                                                this.rng2.Cells[(global::System.Object)SplitColumn, (global::System.Object)Index].Font.Bold = (global::System.Object)true;
                                            if (Fontitalics[x, y])
                                                this.rng2.Cells[(global::System.Object)SplitColumn, (global::System.Object)Index].Font.Italic = (global::System.Object)true;


                                            this.rng2.Cells[(global::System.Object)SplitColumn, (global::System.Object)Index].Interior.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(Red1s[x, y], Green1s[x, y], Blue1s[x, y]);

                                            this.rng2.Cells[(global::System.Object)SplitColumn, (global::System.Object)Index].Font.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(Red2s[x, y], Green2s[x, y], Blue2s[x, y]);
                                            this.workSheet2.Activate();
                                        }
                                        else
                                        {
                                            this.rng2.Cells[(global::System.Object)SplitColumn, (global::System.Object)Index].ClearFormats();
                                        }
                                        for (global::System.Int32 j = (SplitColumn) + (1), loopTo72 = c; j <= loopTo72; j++)
                                        {
                                            this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].value = Arr[(i) - (1), (j) - (1)];
                                            if (((this.CheckBox1.Checked) == (true)))
                                            {
                                                global::System.Int32 x = ((i) - (1));
                                                global::System.Int32 y = ((j) - (1));

                                                this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].Font.Name = FontNames[x, y];
                                                this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].Font.Size = (global::System.Object)FontSizes[x, y];

                                                if (FontBolds[x, y])
                                                    this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].Font.Bold = (global::System.Object)true;
                                                if (Fontitalics[x, y])
                                                    this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].Font.Italic = (global::System.Object)true;


                                                this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].Interior.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(Red1s[x, y], Green1s[x, y], Blue1s[x, y]);

                                                this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].Font.Color = (global::System.Object)global::System.Drawing.Color.FromArgb(Red2s[x, y], Green2s[x, y], Blue2s[x, y]);
                                                this.workSheet2.Activate();
                                            }
                                            else
                                            {
                                                this.rng2.Cells[(global::System.Object)j, (global::System.Object)Index].ClearFormats();
                                            }
                                        }
                                    }
                                }

                            }

                        }

                    }

                    this.Close();
                    this.workSheet2.Activate();
                    this.rng2.Select();

                    global::System.Int32 columnNum;
                    for (global::System.Int32 j = 1, loopTo73 = this.rng2.Columns.Count; j <= loopTo73; j++)
                    {
                        columnNum = Conversions.ToInteger(this.rng2.Cells[(global::System.Object)1, (global::System.Object)j].column);
                        workSheet2.Columns[(global::System.Object)columnNum].Autofit();
                    }

                }
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
                this.workBook = excelApp.ActiveWorkbook;
                this.workSheet = (global::Microsoft.Office.Interop.Excel.Worksheet)workBook.ActiveSheet;

                this.TextBox1.SelectionStart = this.TextBox1.Text.Length;
                this.TextBox1.ScrollToCaret();

                this.rng = this.workSheet.get_Range(this.TextBox1.Text);
                this.TextBoxChanged = true;
                this.rng.Select();

                this.ComboBox3.Items.Clear();

                for (global::System.Int32 j = 1, loopTo = this.rng.Columns.Count; j <= loopTo; j++)
                {
                    global::System.String ItemName;
                    global::System.String CName = global::Microsoft.VisualBasic.Strings.Split(Conversions.ToString(this.rng.Cells[(global::System.Object)1, (global::System.Object)j].Address), "$")[1];
                    if (Conversions.ToBoolean(Operators.ConditionalCompareObjectGreater(this.rng.Cells[(global::System.Object)1, (global::System.Object)1].Row, 1, false)))
                    {
                        ItemName = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject((("Column " + CName) + " ("), this.rng.Cells[(global::System.Object)0, (global::System.Object)j].value), ") "));
                    }
                    else
                    {
                        ItemName = ("Column " + CName);
                    }
                    this.ComboBox3.Items.Add(ItemName);
                }

                this.Display();

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
                if (this.RadioButton1.Checked)
                {
                    this.Display();
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
                if (this.RadioButton2.Checked)
                {
                    this.Display();
                }
            }
            catch (global::System.Exception ex)
            {

            }

        }

        private void RadioButton9_CheckedChanged(global::System.Object sender, global::System.EventArgs e)
        {

            try
            {
                if (this.RadioButton9.Checked)
                {
                    this.Display();
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
                if (this.RadioButton8.Checked)
                {
                    this.Display();
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
                if (this.RadioButton3.Checked)
                {
                    this.Display();
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
                if (this.RadioButton7.Checked)
                {
                    this.Display();
                }
            }
            catch (global::System.Exception ex)
            {

            }

        }

        private void RadioButton10_CheckedChanged(global::System.Object sender, global::System.EventArgs e)
        {

            try
            {
                if (this.RadioButton10.Checked)
                {
                    this.ComboBox2.Enabled = true;
                    this.ComboBox2.Focus();
                    this.Display();
                }
                else
                {
                    this.ComboBox2.Text = "";
                    this.ComboBox2.Enabled = false;
                }
            }

            catch (global::System.Exception ex)
            {

            }

        }

        private void RadioButton11_CheckedChanged(global::System.Object sender, global::System.EventArgs e)
        {

            try
            {
                if (this.RadioButton11.Checked)
                {
                    this.PictureBox11.Enabled = true;
                    this.TextBox3.Enabled = true;
                    this.TextBox3.Focus();
                    this.Display();
                }
                else
                {
                    this.TextBox3.Clear();
                    this.PictureBox11.Enabled = false;
                    this.TextBox3.Enabled = false;
                }
            }
            catch (global::System.Exception ex)
            {

            }

        }

        private void ComboBox2_TextChanged(global::System.Object sender, global::System.EventArgs e)
        {

            try
            {
                this.Display();
            }
            catch (global::System.Exception ex)
            {

            }

        }

        private void TextBox3_TextChanged(global::System.Object sender, global::System.EventArgs e)
        {

            try
            {
                if (((global::Microsoft.VisualBasic.Information.IsNumeric(this.TextBox3.Text)) | string.IsNullOrEmpty(this.TextBox3.Text)))
                {
                    this.Display();
                }
                else
                {
                    global::System.Windows.Forms.MessageBox.Show("Enter a Numerical Value.", "Error", global::System.Windows.Forms.MessageBoxButtons.OK, global::System.Windows.Forms.MessageBoxIcon.Error);
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
            }
            catch (global::System.Exception ex)
            {

            }
        }

        private void AutoSelection_Click(global::System.Object sender, global::System.EventArgs e)
        {
            try
            {

                this.FocusedTextBox = 1;

                global::Microsoft.Office.Interop.Excel.Range userInput = (global::Microsoft.Office.Interop.Excel.Range)this.excelApp.InputBox("Select a range", Type: (global::System.Object)8);
                this.rng = userInput;

                try
                {
                    global::System.String sheetName;
                    sheetName = global::Microsoft.VisualBasic.Strings.Split(this.rng.get_Address((global::System.Object)true, (global::System.Object)true, global::Microsoft.Office.Interop.Excel.XlReferenceStyle.xlA1, (global::System.Object)true), "]")[1];
                    sheetName = global::Microsoft.VisualBasic.Strings.Split(sheetName, "!")[0];

                    if ((global::Microsoft.VisualBasic.Strings.Mid(sheetName, global::Microsoft.VisualBasic.Strings.Len(sheetName), 1) == "'"))
                    {
                        sheetName = global::Microsoft.VisualBasic.Strings.Mid(sheetName, 1, (global::Microsoft.VisualBasic.Strings.Len(sheetName)) - (1));
                    }

                    this.workSheet = (global::Microsoft.Office.Interop.Excel.Worksheet)workBook.Worksheets[sheetName];
                    this.workSheet.Activate();
                }

                catch (global::System.Exception ex)
                {

                }

                this.rng.Select();

                this.rng = this.excelApp.get_Range(this.rng, this.rng.get_End(global::Microsoft.Office.Interop.Excel.XlDirection.xlDown));
                this.rng = this.excelApp.get_Range(this.rng, this.rng.get_End(global::Microsoft.Office.Interop.Excel.XlDirection.xlToRight));

                this.rng.Select();
                this.TextBox1.Text = this.rng.get_Address();
                this.TextBox1.Focus();
            }

            catch (global::System.Exception ex)
            {

            }
        }

        private void Selection_Click(global::System.Object sender, global::System.EventArgs e)
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

                    this.workSheet = (global::Microsoft.Office.Interop.Excel.Worksheet)workBook.Worksheets[sheetName];
                    this.workSheet.Activate();
                }
                catch (global::System.Exception ex)
                {

                }

                rng.Select();

                this.TextBox1.Text = rng.get_Address();
                this.TextBox1.Focus();
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
                    this.workSheet2 = this.workSheet;
                    this.rng2 = this.rng;
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
                    this.Label3.Enabled = true;
                    this.PictureBox2.Enabled = true;
                    this.PictureBox3.Enabled = true;
                    this.TextBox4.Enabled = true;
                    this.TextBox4.Focus();
                }
                else
                {
                    this.TextBox4.Clear();
                    this.Label3.Enabled = false;
                    this.PictureBox2.Enabled = false;
                    this.PictureBox3.Enabled = false;
                    this.TextBox4.Enabled = false;
                }
            }
            catch (global::System.Exception ex)
            {

            }
        }

        private void TextBox4_TextChanged(global::System.Object sender, global::System.EventArgs e)
        {
            try
            {
                this.workSheet2 = (global::Microsoft.Office.Interop.Excel.Worksheet)workBook.ActiveSheet;

                this.TextBox4.SelectionStart = this.TextBox4.Text.Length;
                this.TextBox4.ScrollToCaret();

                this.rng2 = this.workSheet2.get_Range(this.TextBox3.Text);

                this.TextBoxChanged = true;
                this.rng2.Select();
                this.TextBoxChanged = false;
            }

            catch (global::System.Exception ex)
            {

            }
        }

        private void PictureBox3_Click(global::System.Object sender, global::System.EventArgs e)
        {
            try
            {
                this.FocusedTextBox = 4;
                this.Hide();

                this.excelApp = global::VSTO_Addins.Globals.ThisAddIn.Application;
                this.workBook = excelApp.ActiveWorkbook;

                global::Microsoft.Office.Interop.Excel.Range userInput = (global::Microsoft.Office.Interop.Excel.Range)this.excelApp.InputBox("Select a range", Type: (global::System.Object)8);
                this.rng2 = userInput;


                global::System.String sheetName;
                sheetName = global::Microsoft.VisualBasic.Strings.Split(this.rng2.get_Address((global::System.Object)true, (global::System.Object)true, global::Microsoft.Office.Interop.Excel.XlReferenceStyle.xlA1, (global::System.Object)true), "]")[1];
                sheetName = global::Microsoft.VisualBasic.Strings.Split(sheetName, "!")[0];

                if ((global::Microsoft.VisualBasic.Strings.Mid(sheetName, global::Microsoft.VisualBasic.Strings.Len(sheetName), 1) == "'"))
                {
                    sheetName = global::Microsoft.VisualBasic.Strings.Mid(sheetName, 1, (global::Microsoft.VisualBasic.Strings.Len(sheetName)) - (1));
                }

                this.workSheet2 = (global::Microsoft.Office.Interop.Excel.Worksheet)workBook.Worksheets[sheetName];
                this.workSheet2.Activate();

                this.rng2.Select();

                this.TextBox4.Text = this.rng2.get_Address();

                this.Show();
                this.TextBox4.Focus();
            }

            catch (global::System.Exception ex)
            {

                this.Show();
                this.TextBox4.Focus();

            }
        }

        private void Form25_Split_Range_Load(global::System.Object sender, global::System.EventArgs e)
        {
            try
            {

                this.excelApp = global::VSTO_Addins.Globals.ThisAddIn.Application;
                this.workBook = excelApp.ActiveWorkbook;
                this.workSheet = (global::Microsoft.Office.Interop.Excel.Worksheet)workBook.ActiveSheet;
                this.workSheet2 = (global::Microsoft.Office.Interop.Excel.Worksheet)workBook.ActiveSheet;

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
                        this.TextBox1.Text = selectedRange.get_Address();
                        this.workSheet = (global::Microsoft.Office.Interop.Excel.Worksheet)workBook.ActiveSheet;
                        this.rng = selectedRange;
                        this.TextBox1.Focus();
                    }

                    else if (((this.FocusedTextBox) == (4)))
                    {
                        this.TextBox4.Text = selectedRange.get_Address();
                        this.workSheet2 = (global::Microsoft.Office.Interop.Excel.Worksheet)workBook.ActiveSheet;
                        this.rng2 = selectedRange;
                        this.TextBox4.Focus();
                    }
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

        private void TextBox4_GotFocus(global::System.Object sender, global::System.EventArgs e)
        {
            try
            {
                this.FocusedTextBox = 4;
            }
            catch (global::System.Exception ex)
            {

            }
        }

        private void AutoSelection_GotFocus(global::System.Object sender, global::System.EventArgs e)
        {
            try
            {
                this.FocusedTextBox = 1;
            }
            catch (global::System.Exception ex)
            {

            }
        }

        private void Selection_GotFocus(global::System.Object sender, global::System.EventArgs e)
        {
            try
            {
                this.FocusedTextBox = 1;
            }
            catch (global::System.Exception ex)
            {

            }
        }

        private void PictureBox3_GotFocus(global::System.Object sender, global::System.EventArgs e)
        {
            try
            {
                this.FocusedTextBox = 4;
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

        private void ComboBox3_SelectedIndexChanged(global::System.Object sender, global::System.EventArgs e)
        {
            try
            {
                this.Display();
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

        private void AutoSelection_KeyDown(global::System.Object sender, global::System.Windows.Forms.KeyEventArgs e)
        {
            try
            {
                if ((e.KeyCode == global::System.Windows.Forms.Keys.Enter))
                {

                    this.Button2_Click(sender, e);

                }
            }

            catch (global::System.Exception ex)
            {

            }
        }

        private void Button1_KeyDown(global::System.Object sender, global::System.Windows.Forms.KeyEventArgs e)
        {
            try
            {
                if ((e.KeyCode == global::System.Windows.Forms.Keys.Enter))
                {

                    this.Button2_Click(sender, e);

                }
            }

            catch (global::System.Exception ex)
            {

            }
        }

        private void Button2_KeyDown(global::System.Object sender, global::System.Windows.Forms.KeyEventArgs e)
        {
            try
            {
                if ((e.KeyCode == global::System.Windows.Forms.Keys.Enter))
                {

                    this.Button2_Click(sender, e);

                }
            }

            catch (global::System.Exception ex)
            {

            }
        }

        private void CheckBox1_KeyDown(global::System.Object sender, global::System.Windows.Forms.KeyEventArgs e)
        {
            try
            {
                if ((e.KeyCode == global::System.Windows.Forms.Keys.Enter))
                {

                    this.Button2_Click(sender, e);

                }
            }

            catch (global::System.Exception ex)
            {

            }
        }

        private void CheckBox2_KeyDown(global::System.Object sender, global::System.Windows.Forms.KeyEventArgs e)
        {
            try
            {
                if ((e.KeyCode == global::System.Windows.Forms.Keys.Enter))
                {

                    this.Button2_Click(sender, e);

                }
            }

            catch (global::System.Exception ex)
            {

            }
        }

        private void ComboBox1_KeyDown(global::System.Object sender, global::System.Windows.Forms.KeyEventArgs e)
        {
            try
            {
                if ((e.KeyCode == global::System.Windows.Forms.Keys.Enter))
                {

                    this.Button2_Click(sender, e);

                }
            }

            catch (global::System.Exception ex)
            {

            }
        }

        private void ComboBox2_KeyDown(global::System.Object sender, global::System.Windows.Forms.KeyEventArgs e)
        {
            try
            {
                if ((e.KeyCode == global::System.Windows.Forms.Keys.Enter))
                {

                    this.Button2_Click(sender, e);

                }
            }

            catch (global::System.Exception ex)
            {

            }
        }

        private void ComboBox3_KeyDown(global::System.Object sender, global::System.Windows.Forms.KeyEventArgs e)
        {
            try
            {
                if ((e.KeyCode == global::System.Windows.Forms.Keys.Enter))
                {

                    this.Button2_Click(sender, e);

                }
            }

            catch (global::System.Exception ex)
            {

            }
        }

        private void CustomGroupBox1_KeyDown(global::System.Object sender, global::System.Windows.Forms.KeyEventArgs e)
        {
            try
            {
                if ((e.KeyCode == global::System.Windows.Forms.Keys.Enter))
                {

                    this.Button2_Click(sender, e);

                }
            }

            catch (global::System.Exception ex)
            {

            }
        }

        private void CustomGroupBox10_KeyDown(global::System.Object sender, global::System.Windows.Forms.KeyEventArgs e)
        {
            try
            {
                if ((e.KeyCode == global::System.Windows.Forms.Keys.Enter))
                {

                    this.Button2_Click(sender, e);

                }
            }

            catch (global::System.Exception ex)
            {

            }
        }

        private void CustomGroupBox2_KeyDown(global::System.Object sender, global::System.Windows.Forms.KeyEventArgs e)
        {
            try
            {
                if ((e.KeyCode == global::System.Windows.Forms.Keys.Enter))
                {

                    this.Button2_Click(sender, e);

                }
            }

            catch (global::System.Exception ex)
            {

            }
        }

        private void CustomGroupBox4_KeyDown(global::System.Object sender, global::System.Windows.Forms.KeyEventArgs e)
        {
            try
            {
                if ((e.KeyCode == global::System.Windows.Forms.Keys.Enter))
                {

                    this.Button2_Click(sender, e);

                }
            }

            catch (global::System.Exception ex)
            {

            }
        }

        private void CustomGroupBox5_KeyDown(global::System.Object sender, global::System.Windows.Forms.KeyEventArgs e)
        {
            try
            {
                if ((e.KeyCode == global::System.Windows.Forms.Keys.Enter))
                {

                    this.Button2_Click(sender, e);

                }
            }

            catch (global::System.Exception ex)
            {

            }
        }

        private void CustomGroupBox6_KeyDown(global::System.Object sender, global::System.Windows.Forms.KeyEventArgs e)
        {
            try
            {
                if ((e.KeyCode == global::System.Windows.Forms.Keys.Enter))
                {

                    this.Button2_Click(sender, e);

                }
            }

            catch (global::System.Exception ex)
            {

            }
        }

        private void CustomGroupBox7_KeyDown(global::System.Object sender, global::System.Windows.Forms.KeyEventArgs e)
        {
            try
            {
                if ((e.KeyCode == global::System.Windows.Forms.Keys.Enter))
                {

                    this.Button2_Click(sender, e);

                }
            }

            catch (global::System.Exception ex)
            {

            }
        }

        private void CustomGroupBox8_KeyDown(global::System.Object sender, global::System.Windows.Forms.KeyEventArgs e)
        {
            try
            {
                if ((e.KeyCode == global::System.Windows.Forms.Keys.Enter))
                {

                    this.Button2_Click(sender, e);

                }
            }

            catch (global::System.Exception ex)
            {

            }
        }

        private void CustomPanel1_KeyDown(global::System.Object sender, global::System.Windows.Forms.KeyEventArgs e)
        {
            try
            {
                if ((e.KeyCode == global::System.Windows.Forms.Keys.Enter))
                {

                    this.Button2_Click(sender, e);

                }
            }

            catch (global::System.Exception ex)
            {

            }
        }

        private void CustomPanel2_KeyDown(global::System.Object sender, global::System.Windows.Forms.KeyEventArgs e)
        {
            try
            {
                if ((e.KeyCode == global::System.Windows.Forms.Keys.Enter))
                {

                    this.Button2_Click(sender, e);

                }
            }

            catch (global::System.Exception ex)
            {

            }
        }

        private void Label1_KeyDown(global::System.Object sender, global::System.Windows.Forms.KeyEventArgs e)
        {
            try
            {
                if ((e.KeyCode == global::System.Windows.Forms.Keys.Enter))
                {

                    this.Button2_Click(sender, e);

                }
            }

            catch (global::System.Exception ex)
            {

            }
        }

        private void Label2_KeyDown(global::System.Object sender, global::System.Windows.Forms.KeyEventArgs e)
        {
            try
            {
                if ((e.KeyCode == global::System.Windows.Forms.Keys.Enter))
                {

                    this.Button2_Click(sender, e);

                }
            }

            catch (global::System.Exception ex)
            {

            }
        }

        private void Label3_KeyDown(global::System.Object sender, global::System.Windows.Forms.KeyEventArgs e)
        {
            try
            {
                if ((e.KeyCode == global::System.Windows.Forms.Keys.Enter))
                {

                    this.Button2_Click(sender, e);

                }
            }

            catch (global::System.Exception ex)
            {

            }

        }

        private void PictureBox1_KeyDown(global::System.Object sender, global::System.Windows.Forms.KeyEventArgs e)
        {
            try
            {
                if ((e.KeyCode == global::System.Windows.Forms.Keys.Enter))
                {

                    this.Button2_Click(sender, e);

                }
            }

            catch (global::System.Exception ex)
            {

            }
        }

        private void PictureBox10_KeyDown(global::System.Object sender, global::System.Windows.Forms.KeyEventArgs e)
        {
            try
            {
                if ((e.KeyCode == global::System.Windows.Forms.Keys.Enter))
                {

                    this.Button2_Click(sender, e);

                }
            }

            catch (global::System.Exception ex)
            {

            }
        }

        private void PictureBox11_KeyDown(global::System.Object sender, global::System.Windows.Forms.KeyEventArgs e)
        {
            try
            {
                if ((e.KeyCode == global::System.Windows.Forms.Keys.Enter))
                {

                    this.Button2_Click(sender, e);

                }
            }

            catch (global::System.Exception ex)
            {

            }
        }

        private void PictureBox2_KeyDown(global::System.Object sender, global::System.Windows.Forms.KeyEventArgs e)
        {
            try
            {
                if ((e.KeyCode == global::System.Windows.Forms.Keys.Enter))
                {

                    this.Button2_Click(sender, e);

                }
            }

            catch (global::System.Exception ex)
            {

            }
        }

        private void PictureBox3_KeyDown(global::System.Object sender, global::System.Windows.Forms.KeyEventArgs e)
        {
            try
            {
                if ((e.KeyCode == global::System.Windows.Forms.Keys.Enter))
                {

                    this.Button2_Click(sender, e);

                }
            }

            catch (global::System.Exception ex)
            {

            }
        }

        private void PictureBox4_KeyDown(global::System.Object sender, global::System.Windows.Forms.KeyEventArgs e)
        {
            try
            {
                if ((e.KeyCode == global::System.Windows.Forms.Keys.Enter))
                {

                    this.Button2_Click(sender, e);

                }
            }

            catch (global::System.Exception ex)
            {

            }
        }

        private void PictureBox5_KeyDown(global::System.Object sender, global::System.Windows.Forms.KeyEventArgs e)
        {
            try
            {
                if ((e.KeyCode == global::System.Windows.Forms.Keys.Enter))
                {

                    this.Button2_Click(sender, e);

                }
            }

            catch (global::System.Exception ex)
            {

            }
        }

        private void PictureBox6_KeyDown(global::System.Object sender, global::System.Windows.Forms.KeyEventArgs e)
        {
            try
            {
                if ((e.KeyCode == global::System.Windows.Forms.Keys.Enter))
                {

                    this.Button2_Click(sender, e);

                }
            }

            catch (global::System.Exception ex)
            {

            }
        }

        private void PictureBox7_KeyDown(global::System.Object sender, global::System.Windows.Forms.KeyEventArgs e)
        {
            try
            {
                if ((e.KeyCode == global::System.Windows.Forms.Keys.Enter))
                {

                    this.Button2_Click(sender, e);

                }
            }

            catch (global::System.Exception ex)
            {

            }
        }

        private void PictureBox8_KeyDown(global::System.Object sender, global::System.Windows.Forms.KeyEventArgs e)
        {
            try
            {
                if ((e.KeyCode == global::System.Windows.Forms.Keys.Enter))
                {

                    this.Button2_Click(sender, e);

                }
            }

            catch (global::System.Exception ex)
            {

            }
        }

        private void RadioButton1_KeyDown(global::System.Object sender, global::System.Windows.Forms.KeyEventArgs e)
        {
            try
            {
                if ((e.KeyCode == global::System.Windows.Forms.Keys.Enter))
                {

                    this.Button2_Click(sender, e);

                }
            }

            catch (global::System.Exception ex)
            {

            }
        }

        private void RadioButton10_KeyDown(global::System.Object sender, global::System.Windows.Forms.KeyEventArgs e)
        {
            try
            {
                if ((e.KeyCode == global::System.Windows.Forms.Keys.Enter))
                {

                    this.Button2_Click(sender, e);

                }
            }

            catch (global::System.Exception ex)
            {

            }
        }

        private void RadioButton11_KeyDown(global::System.Object sender, global::System.Windows.Forms.KeyEventArgs e)
        {
            try
            {
                if ((e.KeyCode == global::System.Windows.Forms.Keys.Enter))
                {

                    this.Button2_Click(sender, e);

                }
            }

            catch (global::System.Exception ex)
            {

            }
        }

        private void RadioButton2_KeyDown(global::System.Object sender, global::System.Windows.Forms.KeyEventArgs e)
        {
            try
            {
                if ((e.KeyCode == global::System.Windows.Forms.Keys.Enter))
                {

                    this.Button2_Click(sender, e);

                }
            }

            catch (global::System.Exception ex)
            {

            }
        }

        private void RadioButton3_KeyDown(global::System.Object sender, global::System.Windows.Forms.KeyEventArgs e)
        {
            try
            {
                if ((e.KeyCode == global::System.Windows.Forms.Keys.Enter))
                {

                    this.Button2_Click(sender, e);

                }
            }

            catch (global::System.Exception ex)
            {

            }
        }

        private void RadioButton4_KeyDown(global::System.Object sender, global::System.Windows.Forms.KeyEventArgs e)
        {
            try
            {
                if ((e.KeyCode == global::System.Windows.Forms.Keys.Enter))
                {

                    this.Button2_Click(sender, e);

                }
            }

            catch (global::System.Exception ex)
            {

            }
        }

        private void RadioButton5_KeyDown(global::System.Object sender, global::System.Windows.Forms.KeyEventArgs e)
        {
            try
            {
                if ((e.KeyCode == global::System.Windows.Forms.Keys.Enter))
                {

                    this.Button2_Click(sender, e);

                }
            }

            catch (global::System.Exception ex)
            {

            }
        }

        private void RadioButton7_KeyDown(global::System.Object sender, global::System.Windows.Forms.KeyEventArgs e)
        {
            try
            {
                if ((e.KeyCode == global::System.Windows.Forms.Keys.Enter))
                {

                    this.Button2_Click(sender, e);

                }
            }

            catch (global::System.Exception ex)
            {

            }
        }

        private void RadioButton8_KeyDown(global::System.Object sender, global::System.Windows.Forms.KeyEventArgs e)
        {
            try
            {
                if ((e.KeyCode == global::System.Windows.Forms.Keys.Enter))
                {

                    this.Button2_Click(sender, e);

                }
            }

            catch (global::System.Exception ex)
            {

            }
        }

        private void RadioButton9_KeyDown(global::System.Object sender, global::System.Windows.Forms.KeyEventArgs e)
        {
            try
            {
                if ((e.KeyCode == global::System.Windows.Forms.Keys.Enter))
                {

                    this.Button2_Click(sender, e);

                }
            }

            catch (global::System.Exception ex)
            {

            }
        }

        private void TextBox1_KeyDown(global::System.Object sender, global::System.Windows.Forms.KeyEventArgs e)
        {
            try
            {
                if ((e.KeyCode == global::System.Windows.Forms.Keys.Enter))
                {

                    this.Button2_Click(sender, e);

                }
            }

            catch (global::System.Exception ex)
            {

            }
        }

        private void TextBox3_KeyDown(global::System.Object sender, global::System.Windows.Forms.KeyEventArgs e)
        {
            try
            {
                if ((e.KeyCode == global::System.Windows.Forms.Keys.Enter))
                {

                    this.Button2_Click(sender, e);

                }
            }

            catch (global::System.Exception ex)
            {

            }
        }

        private void TextBox4_KeyDown(global::System.Object sender, global::System.Windows.Forms.KeyEventArgs e)
        {
            try
            {
                if ((e.KeyCode == global::System.Windows.Forms.Keys.Enter))
                {

                    this.Button2_Click(sender, e);

                }
            }

            catch (global::System.Exception ex)
            {

            }
        }

        private void VScrollBar1_KeyDown(global::System.Object sender, global::System.Windows.Forms.KeyEventArgs e)
        {
            try
            {
                if ((e.KeyCode == global::System.Windows.Forms.Keys.Enter))
                {

                    this.Button2_Click(sender, e);

                }
            }

            catch (global::System.Exception ex)
            {

            }
        }

        private void Form25_Split_Range_Closing(global::System.Object sender, global::System.ComponentModel.CancelEventArgs e)
        {
            global::VSTO_Addins.GlobalModule.form_flag = false;
        }

        private void Form25_Split_Range_Disposed(global::System.Object sender, global::System.EventArgs e)
        {
            global::VSTO_Addins.GlobalModule.form_flag = false;
        }

        private void Form25_Split_Range_Shown(global::System.Object sender, global::System.EventArgs e)
        {
            this.Focus();
            this.BringToFront();
            this.Activate();
            this.BeginInvoke(new global::System.Action(() =>
                {
                    this.TextBox1.Text = this.rng.get_Address();
                    global::VSTO_Addins.Form25_Split_Range.SetWindowPos(this.Handle, new global::System.IntPtr(global::VSTO_Addins.Form25_Split_Range.HWND_TOPMOST), 0, 0, 0, 0, ((global::VSTO_Addins.Form25_Split_Range.SWP_NOACTIVATE) | (global::VSTO_Addins.Form25_Split_Range.SWP_NOMOVE)) | (global::VSTO_Addins.Form25_Split_Range.SWP_NOSIZE));
                }));
        }
    }
}