using System;
using System.Collections.Generic;
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

    public partial class Form22_Merge_Duplicate_Rows
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
        private Range rng;
        private Range rng2;
        private Range selectedRange;

        private int opened;
        private int FocusedTextBox;

        private Dictionary<string, System.Windows.Forms.Label> variables = new Dictionary<string, System.Windows.Forms.Label>();
        private List<System.Windows.Forms.Label> labels = new List<System.Windows.Forms.Label>();
        private List<System.Windows.Forms.Label> labels2 = new List<System.Windows.Forms.Label>();
        private List<System.Windows.Forms.Label> labels3 = new List<System.Windows.Forms.Label>();
        private List<ComboBox> comboBoxes = new List<ComboBox>();
        private int clickedLabelNumber;
        private int EnteredLabelNumber;

        public Form22_Merge_Duplicate_Rows()
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
        private object GetUniques(Range rng, bool CaseSensitive)
        {
            object GetUniquesRet = default;

            var Uniques = new int[1];
            Uniques[0] = 1;
            int Index = 0;

            bool Matched;
            for (int i = 2, loopTo = rng.Rows.Count; i <= loopTo; i++)
            {
                Matched = false;
                for (int l = Information.LBound(Uniques), loopTo1 = Information.UBound(Uniques); l <= loopTo1; l++)
                {
                    int count = 0;
                    for (int j = 1, loopTo2 = rng.Columns.Count; j <= loopTo2; j++)
                    {
                        Type Type1;
                        Type Type2;

                        if (rng.Cells[i, j].value is null)
                        {
                            Type1 = typeof(string);
                        }
                        else
                        {
                            Type1 = rng.Cells[i, j].value.GetType();
                        }

                        if (rng.Cells[Uniques[l], j].value is null)
                        {
                            Type2 = typeof(string);
                        }
                        else
                        {
                            Type2 = rng.Cells[Uniques[l], j].value.GetType();
                        }

                        if (Type1.Equals(Type2))
                        {
                            if (CaseSensitive == true)
                            {
                                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(rng.Cells[i, j].value, rng.Cells[Uniques[l], j].value, false)))
                                {
                                    count = count + 1;
                                }
                            }
                            else if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(LCase(rng.Cells[i, j].value), LCase(rng.Cells[Uniques[l], j].value), false)))
                            {
                                count = count + 1;
                            }
                        }
                    }
                    if (count == rng.Columns.Count)
                    {
                        Matched = true;
                        break;
                    }
                }
                if (Matched == false)
                {
                    Index = Index + 1;
                    Array.Resize(ref Uniques, Index + 1);
                    Uniques[Index] = i;
                }
            }

            GetUniquesRet = Uniques;
            return GetUniquesRet;

        }
        private object SearchInArray(object Arr, object value)
        {
            object SearchInArrayRet = default;

            bool Result = false;

            for (int i = Information.LBound((Array)Arr), loopTo = Information.UBound((Array)Arr); i <= loopTo; i++)
            {
                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(Arr((object)i), value, false)))
                {
                    Result = true;
                    break;
                }
            }

            SearchInArrayRet = Result;
            return SearchInArrayRet;

        }

        private object Operation(object Arr, object Flag)
        {
            object OperationRet = default;

            if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(Flag, "    Sum", false)))
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

            else if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(Flag, "    Count", false)))
            {
                int Output = 0;
                for (int i = Information.LBound((Array)Arr), loopTo1 = Information.UBound((Array)Arr); i <= loopTo1; i++)
                {
                    if (Arr((object)i) is not null)
                    {
                        Output = Output + 1;
                    }
                }
                OperationRet = Output;
            }

            else if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(Flag, "    Average", false)))
            {
                double Output = 0d;
                for (int i = Information.LBound((Array)Arr), loopTo2 = Information.UBound((Array)Arr); i <= loopTo2; i++)
                {
                    if (Information.IsNumeric(Arr((object)i)) == true)
                    {
                        Output = Conversions.ToDouble(Operators.AddObject(Output, Arr((object)i)));
                    }
                }
                Output = Output / (Information.UBound((Array)Arr) + 1);
                OperationRet = Output;
            }

            else if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(Flag, "    Max", false)))
            {
                object Output;
                int i = Information.LBound((Array)Arr);
                while (Information.IsNumeric(Arr((object)i)) == false & i <= Information.UBound((Array)Arr) - 1)
                    i = i + 1;
                Output = Arr((object)i);
                var loopTo3 = Information.UBound((Array)Arr);
                for (i = Information.LBound((Array)Arr); i <= loopTo3; i++)
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

            else if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(Flag, "    Min", false)))
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
                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectLess(Arr((object)i), Output, false)))
                        {
                            Output = Arr((object)i);
                        }
                    }
                }
                OperationRet = Output;
            }

            else if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(Flag, "    Product", false)))
            {
                double Output = 1d;
                int count = 0;
                for (int i = Information.LBound((Array)Arr), loopTo5 = Information.UBound((Array)Arr); i <= loopTo5; i++)
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
        private object Search(object Arr, object value, object CaseSensitive)
        {
            object SearchRet = default;

            bool Result;
            Result = false;

            for (int i = Information.LBound((Array)Arr), loopTo = Information.UBound((Array)Arr); i <= loopTo; i++)
            {

                Type Type1;
                Type Type2;

                if (Arr((object)i) is null)
                {
                    Type1 = typeof(string);
                }
                else
                {
                    Type1 = Arr((object)i).GetType();
                }

                if (value is null)
                {
                    Type2 = typeof(string);
                }
                else
                {
                    Type2 = value.GetType();
                }

                if (Type1.Equals(Type2))
                {
                    if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(CaseSensitive, true, false)))
                    {
                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(Arr((object)i), value, false)))
                        {
                            Result = true;
                            break;
                        }
                    }
                    else if (Information.IsNumeric(Arr((object)i)) == true)
                    {
                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(Arr((object)i), value, false)))
                        {
                            Result = true;
                            break;
                        }
                    }
                    else if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(LCase(Arr((object)i)), LCase(value), false)))
                    {
                        Result = true;
                        break;
                    }
                }
            }

            SearchRet = Result;
            return SearchRet;

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

        private void Setup()
        {

            try
            {
                CustomGroupBox7.Controls.Clear();

                labels.Clear();
                labels2.Clear();
                labels3.Clear();
                comboBoxes.Clear();

                rng.Select();

                float height = Label3.Height;

                int i;

                var loopTo = rng.Columns.Count;
                for (i = 1; i <= loopTo; i++)
                {

                    var lbl = new System.Windows.Forms.Label();
                    if (CheckBox5.Checked == true)
                    {
                        lbl.Text = Conversions.ToString(rng.Cells[0, i].Value);
                    }
                    else
                    {
                        string columnLetter = Strings.Split(Conversions.ToString(rng.Cells[1, i].Address((object)true, (object)true)), "$")[1];
                        lbl.Text = "Column " + columnLetter;
                    }
                    lbl.Location = new System.Drawing.Point(1, (int)Math.Round((i - 1) * height));
                    lbl.Height = (int)Math.Round(height);
                    lbl.Width = Label2.Width - 1;
                    lbl.Font = new System.Drawing.Font("Segoe UI", 9.75f);
                    lbl.TextAlign = ContentAlignment.MiddleCenter;
                    lbl.TextAlign = ContentAlignment.MiddleLeft;
                    lbl.BorderStyle = BorderStyle.None;
                    CustomGroupBox7.Controls.Add(lbl);
                    labels.Add(lbl);

                    lbl.Click += lbl_Click;
                    lbl.MouseEnter += lbl_MouseEnter;
                    lbl.Paint += lbl_Paint;
                    lbl.KeyDown += lbl_KeyDown;

                    var lbl2 = new System.Windows.Forms.Label();
                    lbl2.Text = Conversions.ToString(rng.Cells[1, i].Value);
                    lbl2.Location = new System.Drawing.Point(Label2.Width - 1, (int)Math.Round((i - 1) * height));
                    lbl2.Height = (int)Math.Round(height);
                    lbl2.Width = (int)Math.Round(Label4.Width + 0.5d);
                    lbl2.Font = new System.Drawing.Font("Segoe UI", 9.75f);
                    lbl2.TextAlign = ContentAlignment.MiddleCenter;
                    lbl2.TextAlign = ContentAlignment.MiddleLeft;
                    lbl2.BorderStyle = BorderStyle.None;
                    CustomGroupBox7.Controls.Add(lbl2);
                    labels2.Add(lbl2);

                    lbl2.Click += lbl2_Click;
                    lbl2.MouseEnter += lbl2_MouseEnter;
                    lbl2.Paint += lbl2_Paint;
                    lbl2.KeyDown += lbl2_KeyDown;

                    var lbl3 = new System.Windows.Forms.Label();
                    lbl3.Text = "";
                    lbl3.Location = new System.Drawing.Point((int)Math.Round(Label2.Width + Label4.Width - 0.5d), (int)Math.Round((i - 1) * height));
                    lbl3.Height = (int)Math.Round(height);
                    lbl3.Width = Label5.Width - 1;
                    lbl3.Font = new System.Drawing.Font("Segoe UI", 9.75f);
                    lbl3.TextAlign = ContentAlignment.MiddleCenter;
                    lbl3.TextAlign = ContentAlignment.MiddleLeft;
                    lbl3.BorderStyle = BorderStyle.None;
                    CustomGroupBox7.Controls.Add(lbl3);
                    labels3.Add(lbl3);

                    lbl3.Click += lbl3_Click;
                    lbl3.MouseEnter += lbl3_MouseEnter;
                    lbl3.Paint += lbl3_Paint;
                    lbl3.KeyDown += lbl3_KeyDown;

                    var comboBox = new ComboBox();

                    comboBox.DrawMode = DrawMode.OwnerDrawFixed;
                    comboBox.DrawItem += ComboBox_DrawItem;
                    comboBox.MeasureItem += ComboBox_MeasureItem;
                    comboBox.SelectedIndexChanged += ComboBox_SelectedIndexChanged;
                    comboBox.KeyDown += comboBox_KeyDown;

                    comboBox.Items.Add("Primary Key");
                    comboBox.Items.Add("    Primary Key");
                    comboBox.Items.Add("Separator");
                    comboBox.Items.Add("    Comma");
                    comboBox.Items.Add("    Colon");
                    comboBox.Items.Add("    Semicolon");
                    comboBox.Items.Add("    Space");
                    comboBox.Items.Add("    Nothing");
                    comboBox.Items.Add("    New Line");
                    comboBox.Items.Add("Function");
                    comboBox.Items.Add("    Sum");
                    comboBox.Items.Add("    Count");
                    comboBox.Items.Add("    Average");
                    comboBox.Items.Add("    Max");
                    comboBox.Items.Add("    Min");
                    comboBox.Items.Add("    Product");

                    comboBox.Location = new System.Drawing.Point(Label2.Width + Label4.Width, (int)Math.Round((double)((i - 1) * height) + 0.5d));
                    comboBox.Height = (int)Math.Round(height - 5f);
                    comboBox.Font = new System.Drawing.Font("Segoe UI", 9.75f);
                    comboBox.Width = (int)Math.Round(Label5.Width - 0.5d);
                    comboBox.Visible = false;

                    CustomGroupBox7.Controls.Add(comboBox);

                    comboBoxes.Add(comboBox);

                }
                clickedLabelNumber = 0;
                labels[0].BackColor = Color.FromArgb(217, 217, 217);
                labels2[0].BackColor = Color.FromArgb(217, 217, 217);
                labels3[0].BackColor = Color.FromArgb(217, 217, 217);
                labels3[0].Text = "    Primary Key";
            }

            catch (Exception ex)
            {

            }

        }

        private void ComboBox_DrawItem(object sender, DrawItemEventArgs e)
        {

            try
            {

                ComboBox comboBox;
                comboBox = (ComboBox)sender;

                if (e.Index == -1)
                {
                    return;
                }

                if (e.Index >= 0)
                {
                    bool isHeader = Conversions.ToBoolean(comboBox.Items[e.Index].StartsWith("  "));
                    if (isHeader == false)
                    {
                        e.Graphics.FillRectangle(Brushes.LightGray, e.Bounds);
                        e.Graphics.DrawString(comboBox.Items[e.Index].ToString(), e.Font, Brushes.Black, e.Bounds);
                    }
                    else
                    {
                        e.DrawBackground();
                        e.Graphics.DrawString(comboBox.Items[e.Index].ToString(), e.Font, Brushes.Black, e.Bounds);
                    }
                }
            }

            catch (Exception ex)
            {

            }

        }

        private void ComboBox_MeasureItem(object sender, MeasureItemEventArgs e)
        {

            try
            {
                ComboBox comboBox;
                comboBox = (ComboBox)sender;

                if (e.Index >= 0)
                {
                    bool isHeader = Conversions.ToBoolean(comboBox.Items[e.Index].StartsWith("  "));
                    if (isHeader == false)
                    {
                        e.ItemHeight = 20;
                    }
                    else
                    {
                        e.ItemHeight = 15;
                    }
                }
            }

            catch (Exception ex)
            {

            }

        }

        private void ComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {

            try
            {
                ComboBox comboBox;
                comboBox = (ComboBox)sender;

                if (comboBox.SelectedIndex >= 0)
                {
                    bool isHeader = Conversions.ToBoolean(comboBox.SelectedItem.StartsWith("    "));
                    if (isHeader == false)
                    {
                        comboBox.SelectedIndex = -1;
                    }
                    else
                    {
                        int clickedBoxNumber = comboBoxes.IndexOf(comboBox);
                        labels3[clickedBoxNumber].Text = Conversions.ToString(comboBox.SelectedItem);
                        labels3[clickedBoxNumber].Visible = true;
                        comboBox.Visible = false;
                    }
                }

                int count = 0;
                foreach (System.Windows.Forms.Label label in labels3)
                {
                    if (label.Text == "    Primary Key")
                    {
                        count = count + 1;
                        if (count > 1)
                        {
                            Interaction.MsgBox("There can't be more than one primary key.");
                            label.Text = "";
                            return;
                        }
                    }
                }

                Display();
            }

            catch (Exception ex)
            {

            }
        }
        private void lbl_Paint(object sender, PaintEventArgs e)
        {

            try
            {

                System.Windows.Forms.Label lbl = (System.Windows.Forms.Label)sender;
                var borderColor = Color.FromArgb(245, 245, 245);
                double borderWidth = 0.4d;

                var borderPen = new Pen(borderColor, (float)borderWidth);

                borderPen.DashStyle = System.Drawing.Drawing2D.DashStyle.Dash;

                e.Graphics.DrawRectangle(borderPen, 0, 0, lbl.Width - 1, lbl.Height - 1);

                borderPen.Dispose();
            }

            catch (Exception ex)
            {

            }

        }
        private void lbl2_Paint(object sender, PaintEventArgs e)
        {

            try
            {
                System.Windows.Forms.Label lbl = (System.Windows.Forms.Label)sender;
                var borderColor = Color.FromArgb(245, 245, 245);
                double borderWidth = 0.4d;

                var borderPen = new Pen(borderColor, (float)borderWidth);

                borderPen.DashStyle = System.Drawing.Drawing2D.DashStyle.Dash;

                e.Graphics.DrawRectangle(borderPen, 0, 0, lbl.Width - 1, lbl.Height - 1);

                borderPen.Dispose();
            }

            catch (Exception ex)
            {

            }

        }
        private void lbl3_Paint(object sender, PaintEventArgs e)
        {

            try
            {

                System.Windows.Forms.Label lbl = (System.Windows.Forms.Label)sender;
                var borderColor = Color.FromArgb(245, 245, 245);
                double borderWidth = 0.4d;

                var borderPen = new Pen(borderColor, (float)borderWidth);

                borderPen.DashStyle = System.Drawing.Drawing2D.DashStyle.Dash;

                e.Graphics.DrawRectangle(borderPen, 0, 0, lbl.Width - 1, lbl.Height - 1);

                borderPen.Dispose();
            }

            catch (Exception ex)
            {

            }

        }

        private void lbl_Click(object sender, EventArgs e)
        {

            try
            {

                System.Windows.Forms.Label clickedLabel;
                clickedLabel = (System.Windows.Forms.Label)sender;

                clickedLabelNumber = labels.IndexOf(clickedLabel);

                clickedLabel.BackColor = Color.FromArgb(217, 217, 217);
                labels2[clickedLabelNumber].BackColor = Color.FromArgb(217, 217, 217);
                labels3[clickedLabelNumber].BackColor = Color.FromArgb(217, 217, 217);

                foreach (System.Windows.Forms.Label label in labels)
                {
                    int lNumber = labels.IndexOf(label);
                    if (lNumber != clickedLabelNumber)
                    {
                        labels[lNumber].BackColor = Color.FromArgb(255, 255, 255);
                        labels2[lNumber].BackColor = Color.FromArgb(255, 255, 255);
                        labels3[lNumber].BackColor = Color.FromArgb(255, 255, 255);
                        comboBoxes[lNumber].Visible = false;
                        labels3[lNumber].Visible = true;
                    }
                }

                comboBoxes[clickedLabelNumber].Visible = true;
                labels3[clickedLabelNumber].Visible = false;
            }

            catch (Exception ex)
            {

            }

        }
        private void lbl_MouseEnter(object sender, EventArgs e)
        {

            try
            {

                System.Windows.Forms.Label clickedLabel;
                clickedLabel = (System.Windows.Forms.Label)sender;

                EnteredLabelNumber = labels.IndexOf(clickedLabel);

                if (EnteredLabelNumber != clickedLabelNumber)
                {
                    clickedLabel.BackColor = Color.FromArgb(229, 243, 255);
                    labels2[EnteredLabelNumber].BackColor = Color.FromArgb(229, 243, 255);
                    labels3[EnteredLabelNumber].BackColor = Color.FromArgb(229, 243, 255);
                }

                foreach (System.Windows.Forms.Label label in labels)
                {
                    int lNumber = labels.IndexOf(label);
                    if (lNumber != EnteredLabelNumber & lNumber != clickedLabelNumber)
                    {
                        labels[lNumber].BackColor = Color.FromArgb(255, 255, 255);
                        labels2[lNumber].BackColor = Color.FromArgb(255, 255, 255);
                        labels3[lNumber].BackColor = Color.FromArgb(255, 255, 255);
                    }
                }
            }
            catch (Exception ex)
            {

            }

        }
        private void lbl2_Click(object sender, EventArgs e)
        {

            try
            {
                System.Windows.Forms.Label clickedLabel;
                clickedLabel = (System.Windows.Forms.Label)sender;

                clickedLabelNumber = labels2.IndexOf(clickedLabel);

                clickedLabel.BackColor = Color.FromArgb(217, 217, 217);
                labels[clickedLabelNumber].BackColor = Color.FromArgb(217, 217, 217);
                labels3[clickedLabelNumber].BackColor = Color.FromArgb(217, 217, 217);

                foreach (System.Windows.Forms.Label label in labels)
                {
                    int lNumber = labels.IndexOf(label);
                    if (lNumber != clickedLabelNumber)
                    {
                        labels[lNumber].BackColor = Color.FromArgb(255, 255, 255);
                        labels2[lNumber].BackColor = Color.FromArgb(255, 255, 255);
                        labels3[lNumber].BackColor = Color.FromArgb(255, 255, 255);
                        comboBoxes[lNumber].Visible = false;
                        labels3[lNumber].Visible = true;
                    }
                }

                comboBoxes[clickedLabelNumber].Visible = true;
                labels3[clickedLabelNumber].Visible = false;
            }

            catch (Exception ex)
            {

            }
        }
        private void lbl2_MouseEnter(object sender, EventArgs e)
        {

            try
            {
                System.Windows.Forms.Label clickedLabel;
                clickedLabel = (System.Windows.Forms.Label)sender;

                EnteredLabelNumber = labels2.IndexOf(clickedLabel);


                if (EnteredLabelNumber != clickedLabelNumber)
                {
                    clickedLabel.BackColor = Color.FromArgb(229, 243, 255);
                    labels[EnteredLabelNumber].BackColor = Color.FromArgb(229, 243, 255);
                    labels3[EnteredLabelNumber].BackColor = Color.FromArgb(229, 243, 255);
                }

                foreach (System.Windows.Forms.Label label in labels)
                {
                    int lNumber = labels.IndexOf(label);
                    if (lNumber != EnteredLabelNumber & lNumber != clickedLabelNumber)
                    {
                        labels[lNumber].BackColor = Color.FromArgb(255, 255, 255);
                        labels2[lNumber].BackColor = Color.FromArgb(255, 255, 255);
                        labels3[lNumber].BackColor = Color.FromArgb(255, 255, 255);
                    }
                }
            }
            catch (Exception ex)
            {

            }

        }
        private void lbl3_Click(object sender, EventArgs e)
        {

            try
            {
                System.Windows.Forms.Label clickedLabel;
                clickedLabel = (System.Windows.Forms.Label)sender;

                clickedLabelNumber = labels3.IndexOf(clickedLabel);

                clickedLabel.BackColor = Color.FromArgb(217, 217, 217);
                labels[clickedLabelNumber].BackColor = Color.FromArgb(217, 217, 217);
                labels2[clickedLabelNumber].BackColor = Color.FromArgb(217, 217, 217);

                foreach (System.Windows.Forms.Label label in labels)
                {
                    int lNumber = labels.IndexOf(label);
                    if (lNumber != clickedLabelNumber)
                    {
                        labels[lNumber].BackColor = Color.FromArgb(255, 255, 255);
                        labels2[lNumber].BackColor = Color.FromArgb(255, 255, 255);
                        labels3[lNumber].BackColor = Color.FromArgb(255, 255, 255);
                        comboBoxes[lNumber].Visible = false;
                        labels3[lNumber].Visible = true;
                    }
                }

                comboBoxes[clickedLabelNumber].Visible = true;
                labels3[clickedLabelNumber].Visible = false;
            }

            catch (Exception ex)
            {

            }

        }
        private void lbl3_MouseEnter(object sender, EventArgs e)
        {

            try
            {
                System.Windows.Forms.Label clickedLabel;
                clickedLabel = (System.Windows.Forms.Label)sender;

                EnteredLabelNumber = labels3.IndexOf(clickedLabel);

                if (EnteredLabelNumber != clickedLabelNumber)
                {
                    clickedLabel.BackColor = Color.FromArgb(229, 243, 255);
                    labels[EnteredLabelNumber].BackColor = Color.FromArgb(229, 243, 255);
                    labels2[EnteredLabelNumber].BackColor = Color.FromArgb(229, 243, 255);
                }

                foreach (System.Windows.Forms.Label label in labels)
                {
                    int lNumber = labels.IndexOf(label);
                    if (lNumber != EnteredLabelNumber & lNumber != clickedLabelNumber)
                    {
                        labels[lNumber].BackColor = Color.FromArgb(255, 255, 255);
                        labels2[lNumber].BackColor = Color.FromArgb(255, 255, 255);
                        labels3[lNumber].BackColor = Color.FromArgb(255, 255, 255);
                    }
                }
            }
            catch (Exception ex)
            {

            }

        }

        private void Display()
        {

            try
            {
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

                int r;
                int c;

                r = displayRng.Rows.Count;
                c = displayRng.Columns.Count;

                float height;
                float width;

                if (r <= 6)
                {
                    height = (float)(CustomPanel2.Height / (double)r);
                }
                else
                {
                    height = (float)(CustomPanel2.Height / 6d);
                }

                width = (float)(CustomPanel2.Width / 6d);

                CustomPanel2.AutoScroll = true;

                bool Active = true;

                foreach (var lbl in labels3)
                {
                    if (string.IsNullOrEmpty(lbl.Text))
                    {
                        Active = false;
                        break;
                    }
                }

                bool IsPrimary;
                IsPrimary = false;

                int PrimaryColumn = 0;

                foreach (var lbl in labels3)
                {
                    if (lbl.Text == "    Primary Key")
                    {
                        IsPrimary = true;
                        PrimaryColumn = labels3.IndexOf(lbl) + 1;
                        break;
                    }
                }

                Active = Active & IsPrimary;

                if (Active == true)
                {

                    Range cRng;
                    cRng = workSheet.get_Range(displayRng.Cells[1, PrimaryColumn], displayRng.Cells[displayRng.Rows.Count, PrimaryColumn]);

                    var Arr1 = new object[1];
                    var Arr2 = new int[1];

                    int Index1 = 0;
                    int Index2 = 0;

                    Arr1[0] = cRng.Cells[1, 1].Value;
                    Arr2[0] = 1;

                    bool CaseSensitive;

                    if (CheckBox1.Checked == true)
                    {
                        CaseSensitive = true;
                    }
                    else
                    {
                        CaseSensitive = false;
                    }

                    if (CheckBox3.Checked == true)
                    {
                        for (int i = 1, loopTo = cRng.Rows.Count; i <= loopTo; i++)
                        {
                            if (cRng.Cells[i, 1].Value is not null)
                            {
                                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(Search(Arr1, cRng.Cells[i, 1].Value, (object)CaseSensitive), false, false)))
                                {
                                    Index1 = Index1 + 1;
                                    Index2 = Index2 + 1;
                                    Array.Resize(ref Arr1, Index1 + 1);
                                    Array.Resize(ref Arr2, Index2 + 1);
                                    Arr1[Index1] = cRng.Cells[i, 1].Value;
                                    Arr2[Index2] = i;
                                }
                            }
                        }
                    }
                    else
                    {
                        for (int i = 1, loopTo1 = cRng.Rows.Count; i <= loopTo1; i++)
                        {
                            if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(Search(Arr1, cRng.Cells[i, 1].Value, (object)CaseSensitive), false, false)))
                            {
                                Index1 = Index1 + 1;
                                Index2 = Index2 + 1;
                                Array.Resize(ref Arr1, Index1 + 1);
                                Array.Resize(ref Arr2, Index2 + 1);
                                Arr1[Index1] = cRng.Cells[i, 1].Value;
                                Arr2[Index2] = i;
                            }
                        }
                    }

                    if (Information.UBound(Arr1) + 1 <= 6)
                    {
                        height = (float)(CustomPanel2.Height / (double)(Information.UBound(Arr1) + 1));
                    }
                    else
                    {
                        height = (float)(CustomPanel2.Height / 6d);
                    }

                    float ordinate = 0f;

                    int[] UniQueArr = (int[])GetUniques(displayRng, CaseSensitive);

                    for (int j = 1, loopTo2 = displayRng.Columns.Count; j <= loopTo2; j++)
                    {
                        if (j != PrimaryColumn)
                        {
                            int max = 1;
                            for (int k = Information.LBound(Arr1), loopTo3 = Information.UBound(Arr1); k <= loopTo3; k++)
                            {
                                int count = 0;

                                for (int i = 1, loopTo4 = displayRng.Rows.Count; i <= loopTo4; i++)
                                {

                                    bool DuplicateCondition;
                                    if (CheckBox6.Checked == true)
                                    {
                                        DuplicateCondition = Conversions.ToBoolean(SearchInArray(UniQueArr, i));
                                    }
                                    else
                                    {
                                        DuplicateCondition = true;
                                    }

                                    if (displayRng.Cells[i, j].value is not null & DuplicateCondition)
                                    {
                                        bool Matched;

                                        if (CheckBox1.Checked == true)
                                        {
                                            Matched = Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(displayRng.Cells[i, PrimaryColumn].value, Arr1[k], false));
                                        }
                                        else
                                        {
                                            Matched = Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(LCase(displayRng.Cells[i, PrimaryColumn].value), LCase(Arr1[k]), false));
                                        }

                                        if (Matched == true)
                                        {
                                            count = count + 1;
                                        }
                                    }
                                }
                                if (count > max)
                                {
                                    max = count;
                                }
                            }

                            bool widthFlag;
                            string separator = " ";
                            string Flag = "";

                            if (labels3[j - 1].Text == "    Comma")
                            {
                                separator = ", ";
                                widthFlag = true;
                                Flag = "a";
                            }
                            else if (labels3[j - 1].Text == "    Colon")
                            {
                                separator = ": ";
                                widthFlag = true;
                                Flag = "a";
                            }
                            else if (labels3[j - 1].Text == "    Semicolon")
                            {
                                separator = "; ";
                                widthFlag = true;
                                Flag = "a";
                            }
                            else if (labels3[j - 1].Text == "    Space")
                            {
                                separator = " ";
                                widthFlag = true;
                                Flag = "b";
                            }
                            else if (labels3[j - 1].Text == "    Nothing")
                            {
                                separator = "";
                                widthFlag = true;
                                Flag = "c";
                            }
                            else if (labels3[j - 1].Text == "    New Line")
                            {
                                separator = Microsoft.VisualBasic.Constants.vbNewLine;
                                widthFlag = true;
                                Flag = "b";
                            }
                            else
                            {
                                widthFlag = false;
                                Flag = labels3[j - 1].Text;
                            }

                            for (int k = Information.LBound(Arr1), loopTo5 = Information.UBound(Arr1); k <= loopTo5; k++)
                            {

                                string concatenatedValue = "";
                                object OperatedValue;
                                var Valuess = new object[1];
                                int indx = -1;

                                for (int i = 1, loopTo6 = displayRng.Rows.Count; i <= loopTo6; i++)
                                {

                                    bool DuplicateCondition;
                                    if (CheckBox6.Checked == true)
                                    {
                                        DuplicateCondition = Conversions.ToBoolean(SearchInArray(UniQueArr, i));
                                    }
                                    else
                                    {
                                        DuplicateCondition = true;
                                    }

                                    if (widthFlag == true)
                                    {
                                        if (displayRng.Cells[i, j].Value is not null & DuplicateCondition)
                                        {
                                            bool Matched;
                                            if (CheckBox1.Checked == true)
                                            {
                                                Matched = Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(displayRng.Cells[i, PrimaryColumn].value, Arr1[k], false));
                                            }
                                            else
                                            {
                                                Matched = Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(LCase(displayRng.Cells[i, PrimaryColumn].value), LCase(Arr1[k]), false));
                                            }

                                            if (Matched == true)
                                            {
                                                concatenatedValue = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject(concatenatedValue, displayRng.Cells[i, j].Value), separator));
                                            }
                                        }
                                    }
                                    else
                                    {

                                        bool Matched;
                                        if (displayRng.Cells[i, j].Value is not null & DuplicateCondition)
                                        {
                                            if (CheckBox1.Checked == true)
                                            {
                                                Matched = Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(displayRng.Cells[i, PrimaryColumn].value, Arr1[k], false));
                                            }
                                            else
                                            {
                                                Matched = Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(LCase(displayRng.Cells[i, PrimaryColumn].value), LCase(Arr1[k]), false));
                                            }
                                            if (Matched == true)
                                            {
                                                indx = indx + 1;
                                                Array.Resize(ref Valuess, indx + 1);
                                                Valuess[indx] = displayRng.Cells[i, j].Value;
                                            }
                                        }
                                    }
                                }
                                OperatedValue = Operation(Valuess, Flag);

                                if (Flag == "a")
                                {
                                    if (!string.IsNullOrEmpty(concatenatedValue))
                                    {
                                        concatenatedValue = Strings.Mid(concatenatedValue, 1, Strings.Len(concatenatedValue) - 2);
                                    }
                                }
                                else if (Flag == "b")
                                {
                                    if (!string.IsNullOrEmpty(concatenatedValue))
                                    {
                                        concatenatedValue = Strings.Mid(concatenatedValue, 1, Strings.Len(concatenatedValue) - 1);
                                    }
                                }

                                var label = new System.Windows.Forms.Label();

                                label.Font = new System.Drawing.Font("Segoe UI", 9.75f);
                                label.Location = new System.Drawing.Point((int)Math.Round(ordinate), (int)Math.Round((k + 1 - 1) * height));
                                label.Height = (int)Math.Round(height);
                                if (widthFlag == true)
                                {
                                    label.Width = (int)Math.Round(max * width);
                                    label.Text = concatenatedValue;
                                }
                                else
                                {
                                    label.Width = (int)Math.Round(width);
                                    label.Text = Conversions.ToString(OperatedValue);
                                }
                                label.TextAlign = ContentAlignment.MiddleCenter;
                                CustomPanel2.Controls.Add(label);

                                if (CheckBox4.Checked == true)
                                {

                                    Range cell = (Range)displayRng.Cells[Arr2[k], j];
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

                                label.Paint += label_Paint;
                            }
                            if (widthFlag == true)
                            {
                                ordinate = ordinate + max * width;
                            }
                            else
                            {
                                ordinate = ordinate + width;
                            }
                        }

                        else
                        {
                            for (int k = Information.LBound(Arr1), loopTo7 = Information.UBound(Arr1); k <= loopTo7; k++)
                            {
                                var label = new System.Windows.Forms.Label();
                                label.Text = Conversions.ToString(Arr1[k]);
                                label.Font = new System.Drawing.Font("Segoe UI", 9.75f);
                                label.Location = new System.Drawing.Point((int)Math.Round(ordinate), (int)Math.Round((k + 1 - 1) * height));
                                label.Height = (int)Math.Round(height);
                                label.Width = (int)Math.Round(width);
                                label.TextAlign = ContentAlignment.MiddleCenter;
                                CustomPanel2.Controls.Add(label);

                                if (CheckBox4.Checked == true)
                                {

                                    Range cell = (Range)displayRng.Cells[Arr2[k], j];
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

                                label.Paint += label_Paint;

                            }
                            ordinate = ordinate + width;

                        }
                    }
                    CustomPanel2.AutoScroll = true;
                }
            }

            catch (Exception ex)
            {

            }

        }

        private void label_Paint(object sender, PaintEventArgs e)
        {

            try
            {

                System.Windows.Forms.Label lbl = (System.Windows.Forms.Label)sender;
                var borderColor = Color.FromArgb(245, 245, 245);
                double borderWidth = 0.4d;

                var borderPen = new Pen(borderColor, (float)borderWidth);

                borderPen.DashStyle = System.Drawing.Drawing2D.DashStyle.Dash;

                e.Graphics.DrawRectangle(borderPen, 0, 0, lbl.Width - 1, lbl.Height - 1);

                borderPen.Dispose();
            }

            catch (Exception ex)
            {

            }

        }
        private void lbl_KeyDown(object sender, KeyEventArgs e)
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
        private void lbl2_KeyDown(object sender, KeyEventArgs e)
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
        private void lbl3_KeyDown(object sender, KeyEventArgs e)
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
        private void comboBox_KeyDown(object sender, KeyEventArgs e)
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

        private void TextBox1_TextChanged(object sender, EventArgs e)
        {

            try
            {
                if (!string.IsNullOrEmpty(TextBox1.Text) & IsValidExcelCellReference(TextBox1.Text) == true)
                {

                    excelApp = Globals.ThisAddIn.Application;
                    workBook = excelApp.ActiveWorkbook;
                    workSheet = (Excel.Worksheet)workBook.ActiveSheet;

                    TextBox1.SelectionStart = TextBox1.Text.Length;
                    TextBox1.ScrollToCaret();

                    rng = workSheet.get_Range(TextBox1.Text);
                    rng.Select();

                    Setup();
                    Display();

                }
            }

            catch (Exception ex)
            {

            }

        }

        private void Form22_Merge_Duplicate_Rows_Load(object sender, EventArgs e)
        {

            try
            {

                excelApp = Globals.ThisAddIn.Application;

                excelApp.SheetSelectionChange += excelApp_SheetSelectionChange;

                opened = opened + 1;

                EnteredLabelNumber = -1;
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
                else if (FocusedTextBox == 2)
                {
                    TextBox2.Text = selectedRange.get_Address();
                    workSheet2 = (Excel.Worksheet)workBook.ActiveSheet;
                    rng2 = selectedRange;
                    TextBox2.Focus();
                }
            }

            catch (Exception ex)
            {

            }

        }

        private void CheckBox5_CheckedChanged(object sender, EventArgs e)
        {

            try
            {

                if (CheckBox5.Checked == true)
                {
                    rng = workSheet.get_Range(rng.Cells[2, 1], rng.Cells[rng.Rows.Count, rng.Columns.Count]);
                }
                else
                {
                    rng = workSheet.get_Range(rng.Cells[0, 1], rng.Cells[rng.Rows.Count, rng.Columns.Count]);
                }

                TextBox1.Text = rng.get_Address();
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

        private void Button2_Click(object sender, EventArgs e)
        {

            try
            {

                if (string.IsNullOrEmpty(TextBox1.Text))
                {
                    MessageBox.Show("Select a Source Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    TextBox1.Focus();
                    return;
                }

                if (IsValidExcelCellReference(TextBox1.Text) == false)
                {
                    MessageBox.Show("Select a Valid Source Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    TextBox1.Focus();
                    return;
                }

                if (RadioButton10.Checked == false & RadioButton3.Checked == false)
                {
                    MessageBox.Show("Select a Destination Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (RadioButton10.Checked & string.IsNullOrEmpty(TextBox2.Text))
                {
                    MessageBox.Show("Select a Destination Cell.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    TextBox2.Focus();
                    return;
                }

                if (RadioButton10.Checked & IsValidExcelCellReference(TextBox2.Text) == false)
                {
                    MessageBox.Show("Select a Valid Destination Cell.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    TextBox2.Focus();
                    return;
                }

                if (CheckBox1.Checked == true)
                {
                    workSheet.Copy(After: workBook.Sheets[workSheet.Name]);
                    workSheet2.Activate();
                }

                string rng2Address = rng2.get_Address();

                int r;
                int c;

                r = rng.Rows.Count;
                c = rng.Columns.Count;

                bool Active = true;

                foreach (var lbl in labels3)
                {
                    if (string.IsNullOrEmpty(lbl.Text))
                    {
                        Active = false;
                        break;
                    }
                }

                bool IsPrimary;
                IsPrimary = false;

                int PrimaryColumn = 0;

                foreach (var lbl in labels3)
                {
                    if (lbl.Text == "    Primary Key")
                    {
                        IsPrimary = true;
                        PrimaryColumn = labels3.IndexOf(lbl) + 1;
                        break;
                    }
                }

                Active = Active & IsPrimary;

                if (Active == true)
                {

                    Range cRng;
                    cRng = workSheet.get_Range(rng.Cells[1, PrimaryColumn], rng.Cells[rng.Rows.Count, PrimaryColumn]);

                    var Arr1 = new object[1];
                    var Arr2 = new int[1];

                    int Index1 = 0;
                    int Index2 = 0;

                    Arr1[0] = cRng.Cells[1, 1].Value;
                    Arr2[0] = 1;

                    bool CaseSensitive;

                    if (CheckBox1.Checked == true)
                    {
                        CaseSensitive = true;
                    }
                    else
                    {
                        CaseSensitive = false;
                    }

                    if (CheckBox3.Checked == true)
                    {
                        for (int i = 1, loopTo = cRng.Rows.Count; i <= loopTo; i++)
                        {
                            if (cRng.Cells[i, 1].Value is not null)
                            {
                                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(Search(Arr1, cRng.Cells[i, 1].Value, (object)CaseSensitive), false, false)))
                                {
                                    Index1 = Index1 + 1;
                                    Index2 = Index2 + 1;
                                    Array.Resize(ref Arr1, Index1 + 1);
                                    Array.Resize(ref Arr2, Index2 + 1);
                                    Arr1[Index1] = cRng.Cells[i, 1].Value;
                                    Arr2[Index2] = i;
                                }
                            }
                        }
                    }
                    else
                    {
                        for (int i = 1, loopTo1 = cRng.Rows.Count; i <= loopTo1; i++)
                        {
                            if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(Search(Arr1, cRng.Cells[i, 1].Value, (object)CaseSensitive), false, false)))
                            {
                                Index1 = Index1 + 1;
                                Index2 = Index2 + 1;
                                Array.Resize(ref Arr1, Index1 + 1);
                                Array.Resize(ref Arr2, Index2 + 1);
                                Arr1[Index1] = cRng.Cells[i, 1].Value;
                                Arr2[Index2] = i;
                            }
                        }
                    }

                    rng2 = workSheet2.get_Range(rng2.Cells[1, 1], rng2.Cells[Information.UBound(Arr1) + 1, rng.Columns.Count]);
                    rng2Address = rng2.get_Address();

                    int[] UniQueArr = (int[])GetUniques(rng, CaseSensitive);

                    if (Overlap(excelApp, workSheet, workSheet2, rng, rng2) == true)
                    {

                        var ValueArr = new object[rng.Rows.Count, rng.Columns.Count];

                        for (int i = 1, loopTo2 = rng.Rows.Count; i <= loopTo2; i++)
                        {
                            for (int j = 1, loopTo3 = rng.Columns.Count; j <= loopTo3; j++)
                                ValueArr[i - 1, j - 1] = rng.Cells[i, j].Value;
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

                        if (CheckBox4.Checked == true)
                        {

                            for (int i = Information.LBound(FontSizes, 1), loopTo4 = Information.UBound(FontSizes, 1); i <= loopTo4; i++)
                            {
                                for (int j = Information.LBound(FontSizes, 2), loopTo5 = Information.UBound(FontSizes, 2); j <= loopTo5; j++)
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

                                    FontBolds[i, j] = Conversions.ToBoolean(cell.Font.Bold);
                                    Fontitalics[i, j] = Conversions.ToBoolean(cell.Font.Italic);

                                    if (font.Size is DBNull == false)
                                    {
                                        float fontSize = Convert.ToSingle(font.Size);
                                        FontSizes[i, j] = fontSize;
                                    }
                                    else
                                    {
                                        FontSizes[i, j] = 11f;
                                    }

                                    if (cell.Interior.Color is DBNull)
                                    {
                                        Red1s[i, j] = 0;
                                        Green1s[i, j] = 0;
                                        Blue1s[i, j] = 0;
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

                                }
                            }
                        }

                        rng.ClearContents();
                        rng.ClearFormats();

                        if (CheckBox4.Checked == true)
                        {

                            for (int i = 1, loopTo6 = rng2.Rows.Count; i <= loopTo6; i++)
                            {
                                for (int j = 1, loopTo7 = rng2.Columns.Count; j <= loopTo7; j++)
                                {
                                    int x = Arr2[i - 1] - 1;
                                    int y = j - 1;

                                    rng2.Cells[i, j].Font.Name = FontNames[x, y];
                                    rng2.Cells[i, j].Font.Size = (object)FontSizes[x, y];

                                    if (FontBolds[x, y])
                                        rng2.Cells[i, j].Font.Bold = (object)true;
                                    if (Fontitalics[x, y])
                                        rng2.Cells[i, j].Font.Italic = (object)true;

                                    rng2.Cells[i, j].Interior.Color = (object)Color.FromArgb(Red1s[x, y], Green1s[x, y], Blue1s[x, y]);

                                    rng2.Cells[i, j].Font.Color = (object)Color.FromArgb(Red2s[x, y], Green2s[x, y], Blue2s[x, y]);

                                    Range targetCell = (Range)rng2.Cells[i, j];

                                    for (int k = 7; k <= 11; k++)
                                    {
                                        targetCell.Borders[(XlBordersIndex)k].LineStyle = XlLineStyle.xlContinuous;
                                        targetCell.Borders[(XlBordersIndex)k].Color = Color.Black.ToArgb();
                                    }

                                }
                            }

                        }

                        for (int j = 1, loopTo8 = rng.Columns.Count; j <= loopTo8; j++)
                        {
                            if (j != PrimaryColumn)
                            {
                                bool widthFlag;
                                string separator = " ";
                                string Flag = "";

                                if (labels3[j - 1].Text == "    Comma")
                                {
                                    separator = ", ";
                                    widthFlag = true;
                                    Flag = "a";
                                }
                                else if (labels3[j - 1].Text == "    Colon")
                                {
                                    separator = ": ";
                                    widthFlag = true;
                                    Flag = "a";
                                }
                                else if (labels3[j - 1].Text == "    Semicolon")
                                {
                                    separator = "; ";
                                    widthFlag = true;
                                    Flag = "a";
                                }
                                else if (labels3[j - 1].Text == "    Space")
                                {
                                    separator = " ";
                                    widthFlag = true;
                                    Flag = "b";
                                }
                                else if (labels3[j - 1].Text == "    Nothing")
                                {
                                    separator = "";
                                    widthFlag = true;
                                    Flag = c.ToString();
                                }
                                else if (labels3[j - 1].Text == "    New Line")
                                {
                                    separator = Microsoft.VisualBasic.Constants.vbNewLine;
                                    widthFlag = true;
                                    Flag = "b";
                                }
                                else
                                {
                                    widthFlag = false;
                                    Flag = labels3[j - 1].Text;
                                }

                                for (int k = Information.LBound(Arr1), loopTo9 = Information.UBound(Arr1); k <= loopTo9; k++)
                                {

                                    string concatenatedValue = "";
                                    object OperatedValue;
                                    var Valuess = new object[1];
                                    int indx = -1;

                                    for (int i = 1, loopTo10 = rng.Rows.Count; i <= loopTo10; i++)
                                    {

                                        bool DuplicateCondition;
                                        if (CheckBox6.Checked == true)
                                        {
                                            DuplicateCondition = Conversions.ToBoolean(SearchInArray(UniQueArr, i));
                                        }
                                        else
                                        {
                                            DuplicateCondition = true;
                                        }

                                        if (widthFlag == true)
                                        {

                                            if (ValueArr[i - 1, j - 1] is not null & DuplicateCondition)
                                            {
                                                bool Matched;
                                                if (CheckBox1.Checked == true)
                                                {
                                                    Matched = Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(ValueArr[i - 1, PrimaryColumn - 1], Arr1[k], false));
                                                }
                                                else
                                                {
                                                    Matched = Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(LCase(ValueArr[i - 1, PrimaryColumn - 1]), LCase(Arr1[k]), false));
                                                }

                                                if (Matched == true)
                                                {
                                                    concatenatedValue = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject(concatenatedValue, ValueArr[i - 1, j - 1]), separator));
                                                }
                                            }
                                        }

                                        else if (ValueArr[i - 1, j - 1] is not null & DuplicateCondition)
                                        {
                                            bool Matched;
                                            if (CheckBox1.Checked == true)
                                            {
                                                Matched = Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(ValueArr[i - 1, PrimaryColumn - 1], Arr1[k], false));
                                            }
                                            else
                                            {
                                                Matched = Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(LCase(ValueArr[i - 1, PrimaryColumn - 1]), LCase(Arr1[k]), false));
                                            }
                                            if (Matched == true)
                                            {
                                                indx = indx + 1;
                                                Array.Resize(ref Valuess, indx + 1);
                                                Valuess[indx] = ValueArr[i - 1, j - 1];
                                            }

                                        }
                                    }

                                    OperatedValue = Operation(Valuess, Flag);

                                    if (Flag == "a")
                                    {
                                        if (!string.IsNullOrEmpty(concatenatedValue))
                                        {
                                            concatenatedValue = Strings.Mid(concatenatedValue, 1, Strings.Len(concatenatedValue) - 2);
                                        }
                                    }
                                    else if (Flag == "b")
                                    {
                                        if (!string.IsNullOrEmpty(concatenatedValue))
                                        {
                                            concatenatedValue = Strings.Mid(concatenatedValue, 1, Strings.Len(concatenatedValue) - 1);
                                        }
                                    }

                                    if (widthFlag == true)
                                    {
                                        rng2.Cells[k + 1, j].value = concatenatedValue;
                                    }
                                    else
                                    {
                                        rng2.Cells[k + 1, j].value = OperatedValue;
                                    }

                                }
                            }
                            else
                            {
                                for (int k = Information.LBound(Arr1), loopTo11 = Information.UBound(Arr1); k <= loopTo11; k++)
                                    rng2.Cells[k + 1, j].value = Arr1[k];

                            }

                        }
                    }

                    else
                    {

                        for (int j = 1, loopTo12 = rng.Columns.Count; j <= loopTo12; j++)
                        {
                            if (j != PrimaryColumn)
                            {
                                bool widthFlag;
                                string separator = " ";
                                string Flag = "";

                                if (labels3[j - 1].Text == "    Comma")
                                {
                                    separator = ", ";
                                    widthFlag = true;
                                    Flag = "a";
                                }
                                else if (labels3[j - 1].Text == "    Colon")
                                {
                                    separator = ": ";
                                    widthFlag = true;
                                    Flag = "a";
                                }
                                else if (labels3[j - 1].Text == "    Semicolon")
                                {
                                    separator = "; ";
                                    widthFlag = true;
                                    Flag = "a";
                                }
                                else if (labels3[j - 1].Text == "    Space")
                                {
                                    separator = " ";
                                    widthFlag = true;
                                    Flag = "b";
                                }
                                else if (labels3[j - 1].Text == "    Nothing")
                                {
                                    separator = "";
                                    widthFlag = true;
                                    Flag = c.ToString();
                                }
                                else if (labels3[j - 1].Text == "    New Line")
                                {
                                    separator = Microsoft.VisualBasic.Constants.vbNewLine;
                                    widthFlag = true;
                                    Flag = "b";
                                }
                                else
                                {
                                    widthFlag = false;
                                    Flag = labels3[j - 1].Text;
                                }

                                for (int k = Information.LBound(Arr1), loopTo13 = Information.UBound(Arr1); k <= loopTo13; k++)
                                {

                                    string concatenatedValue = "";
                                    object OperatedValue;
                                    var Valuess = new object[1];
                                    int indx = -1;

                                    for (int i = 1, loopTo14 = rng.Rows.Count; i <= loopTo14; i++)
                                    {

                                        bool DuplicateCondition;
                                        if (CheckBox6.Checked == true)
                                        {
                                            DuplicateCondition = Conversions.ToBoolean(SearchInArray(UniQueArr, i));
                                        }
                                        else
                                        {
                                            DuplicateCondition = true;
                                        }

                                        if (widthFlag == true)
                                        {
                                            if (rng.Cells[i, j].Value is not null & DuplicateCondition)
                                            {
                                                bool Matched;
                                                if (CheckBox1.Checked == true)
                                                {
                                                    Matched = Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(rng.Cells[i, PrimaryColumn].Value, Arr1[k], false));
                                                }
                                                else
                                                {
                                                    Matched = Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(LCase(rng.Cells[i, PrimaryColumn].Value), LCase(Arr1[k]), false));
                                                }

                                                if (Matched == true)
                                                {
                                                    concatenatedValue = Conversions.ToString(Operators.ConcatenateObject(Operators.ConcatenateObject(concatenatedValue, rng.Cells[i, j].Value), separator));
                                                }
                                            }
                                        }

                                        else if (rng.Cells[i, j].Value is not null & DuplicateCondition)
                                        {
                                            bool Matched;
                                            if (CheckBox1.Checked == true)
                                            {
                                                Matched = Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(rng.Cells[i, PrimaryColumn].Value, Arr1[k], false));
                                            }
                                            else
                                            {
                                                Matched = Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(LCase(rng.Cells[i, PrimaryColumn].Value), LCase(Arr1[k]), false));
                                            }
                                            if (Matched == true)
                                            {
                                                indx = indx + 1;
                                                Array.Resize(ref Valuess, indx + 1);
                                                Valuess[indx] = rng.Cells[i, j].Value;
                                            }

                                        }
                                    }

                                    OperatedValue = Operation(Valuess, Flag);

                                    if (Flag == "a")
                                    {
                                        if (!string.IsNullOrEmpty(concatenatedValue))
                                        {
                                            concatenatedValue = Strings.Mid(concatenatedValue, 1, Strings.Len(concatenatedValue) - 2);
                                        }
                                    }
                                    else if (Flag == "b")
                                    {
                                        if (!string.IsNullOrEmpty(concatenatedValue))
                                        {
                                            concatenatedValue = Strings.Mid(concatenatedValue, 1, Strings.Len(concatenatedValue) - 1);
                                        }
                                    }

                                    if (widthFlag == true)
                                    {
                                        rng2.Cells[k + 1, j].value = concatenatedValue;
                                    }
                                    else
                                    {
                                        rng2.Cells[k + 1, j].value = OperatedValue;
                                    }
                                    if (CheckBox4.Checked == true)
                                    {
                                        rng.Cells[Arr2[k], j].Copy();
                                        rng2.Cells[k + 1, j].PasteSpecial(XlPasteType.xlPasteFormats);
                                        rng2 = workSheet2.get_Range(rng2Address);
                                    }
                                }
                                excelApp.CutCopyMode = XlCutCopyMode.xlCopy;
                            }
                            else
                            {
                                for (int k = Information.LBound(Arr1), loopTo15 = Information.UBound(Arr1); k <= loopTo15; k++)
                                {
                                    rng2.Cells[k + 1, j].value = Arr1[k];
                                    if (CheckBox4.Checked == true)
                                    {
                                        rng.Cells[Arr2[k], j].Copy();
                                        rng2.Cells[k + 1, j].PasteSpecial(XlPasteType.xlPasteFormats);
                                        rng2 = workSheet2.get_Range(rng2Address);
                                    }
                                }
                                excelApp.CutCopyMode = XlCutCopyMode.xlCopy;

                            }

                        }

                        for (int j = 1, loopTo16 = rng.Columns.Count; j <= loopTo16; j++)
                        {
                            rng2.Cells[rng2.Rows.Count, j].Borders((object)9).LineStyle = rng.Cells[rng.Rows.Count, j].Borders((object)9).LineStyle;
                            rng2.Cells[rng2.Rows.Count, j].Borders((object)9).Color = rng.Cells[rng.Rows.Count, j].Borders((object)9).Color;
                            rng2.Cells[rng2.Rows.Count, j].Borders((object)9).weight = rng.Cells[rng.Rows.Count, j].Borders((object)9).weight;
                        }

                    }

                    int columnNum;
                    for (int j = 1, loopTo17 = rng2.Columns.Count; j <= loopTo17; j++)
                    {
                        columnNum = Conversions.ToInteger(rng2.Cells[1, j].column);
                        workSheet2.Columns[columnNum].Autofit();
                    }

                    Close();

                    rng2.Select();

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

        private void PictureBox3_Click(object sender, EventArgs e)
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

                TextBox2.Text = rng2.get_Address();

                Show();
                TextBox2.Focus();
            }

            catch (Exception ex)
            {

                Show();
                TextBox2.Focus();

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

        private void TextBox2_TextChanged(object sender, EventArgs e)
        {
            try
            {
                excelApp = Globals.ThisAddIn.Application;
                workBook = excelApp.ActiveWorkbook;
                workSheet2 = (Excel.Worksheet)workBook.ActiveSheet;

                TextBox2.SelectionStart = TextBox2.Text.Length;
                TextBox2.ScrollToCaret();

                rng2 = workSheet2.get_Range(TextBox2.Text);
                rng2.Select();
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

        private void PictureBox3_GotFocus(object sender, EventArgs e)
        {
            try
            {
                FocusedTextBox = 2;
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

        private void CheckBox5_KeyDown(object sender, KeyEventArgs e)
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

        private void Info_KeyDown(object sender, KeyEventArgs e)
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

        private void Label4_KeyDown(object sender, KeyEventArgs e)
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

        private void Label5_KeyDown(object sender, KeyEventArgs e)
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

        private void TextBox2_KeyDown(object sender, KeyEventArgs e)
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

        private void CheckBox6_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                Display();
            }
            catch (Exception ex)
            {

            }
        }

        private void RadioButton10_CheckedChanged(object sender, EventArgs e)
        {

            if (RadioButton10.Checked == true)
            {

                Label3.Visible = true;
                TextBox2.Visible = true;
                PictureBox2.Visible = true;
                PictureBox3.Visible = true;
                TextBox2.Focus();
            }
            else
            {
                TextBox2.Clear();
                Label3.Visible = false;
                TextBox2.Visible = false;
                PictureBox2.Visible = false;
                PictureBox3.Visible = false;

            }

        }

        private void RadioButton3_CheckedChanged(object sender, EventArgs e)
        {

            try
            {

                if (RadioButton3.Checked == true)
                {

                    excelApp = Globals.ThisAddIn.Application;
                    workBook = excelApp.ActiveWorkbook;
                    workSheet2 = workSheet;

                    rng2 = rng;

                    rng2.Select();

                }
            }

            catch (Exception ex)
            {

            }

        }

        private void CheckBox6_KeyDown(object sender, KeyEventArgs e)
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

        private void Form22_Merge_Duplicate_Rows_Closing(object sender, CancelEventArgs e)
        {
            GlobalModule.form_flag = false;
        }

        private void Form22_Merge_Duplicate_Rows_Disposed(object sender, EventArgs e)
        {
            GlobalModule.form_flag = false;
        }

        private void Form22_Merge_Duplicate_Rows_Shown(object sender, EventArgs e)
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