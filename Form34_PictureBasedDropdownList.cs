using System;
using System.ComponentModel;
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

    public partial class Form34_PictureBasedDropdownList
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
                if (_excelApp != null)
                {
                    _excelApp.SheetSelectionChange -= excelApp_SheetSelectionChange;
                }

                _excelApp = value;
                if (_excelApp != null)
                {
                    _excelApp.SheetSelectionChange += excelApp_SheetSelectionChange;
                }
            }
        }
        private Excel.Workbook workBook;
        public static Excel.Worksheet workSheet;

        private Range src_rng;
        public Range des_rng;
        private Range selectedRange;

        public Range validationRange;

        private bool processingEvent = false;
        public bool focuschange;

        public Form34_PictureBasedDropdownList()
        {
            InitializeComponent();
        }

        [DllImport("user32")]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
        private const uint SWP_NOMOVE = 0x2U;
        private const uint SWP_NOSIZE = 0x1U;
        private const uint SWP_NOACTIVATE = 0x10U;
        private const int HWND_TOPMOST = -1;



        private void Btn_OK_Click(object sender, EventArgs e)
        {
            excelApp = Globals.ThisAddIn.Application;
            var workbook = excelApp.ActiveWorkbook;
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.ActiveSheet;

            GlobalModule.sheetName2 = worksheet.Name;


            bool x = false;

            if (IsValidExcelCellReference(TB_src_rng.Text) == true)
            {

                foreach (Shape pic in worksheet.Shapes)
                {
                    for (int i = 1, loopTo = src_rng.Rows.Count; i <= loopTo; i++)
                    {
                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(pic.TopLeftCell.get_Address(), src_rng[i, 2].Address, false)))
                        {

                            x = true;
                            goto BreakAllLoops;
                        }
                        else
                        {
                            x = false;

                        }
                    }

                }

BreakAllLoops:
                ;

            }

            if (string.IsNullOrEmpty(TB_src_rng.Text) & string.IsNullOrEmpty(TB_des_rng.Text))
            {
                MessageBox.Show("Select all necessary options.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                TB_src_rng.Focus();
                // Me.Close()
                return;
            }

            else if (string.IsNullOrEmpty(TB_src_rng.Text))
            {
                MessageBox.Show("Select a Source Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                TB_src_rng.Focus();
                // Me.Close()
                return;
            }

            else if (string.IsNullOrEmpty(TB_des_rng.Text))
            {
                MessageBox.Show("Select the Destination Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                TB_des_rng.Focus();
                // Me.Close()
                return;
            }


            else if (IsValidExcelCellReference(TB_src_rng.Text) == false)
            {
                MessageBox.Show("Select a valid source cell range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                TB_src_rng.Focus();
                return;
            }

            else if (IsValidExcelCellReference(TB_des_rng.Text) == false)
            {
                MessageBox.Show("Select a valid destination cell range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                TB_des_rng.Focus();
                return;
            }
            else if (src_rng.Areas.Count > 1)
            {
                MessageBox.Show("Multiple selection is not possible in the Source Range field. Please select two columns.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                TB_src_rng.Focus();

                return;
            }

            else if (src_rng.Columns.Count == 1)
            {
                MessageBox.Show("Please select both of the columns that contain the data and the relevant images.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                TB_src_rng.Focus();
                // Me.Close()
                return;
            }

            else if (src_rng.Columns.Count > 2)
            {
                MessageBox.Show("Please, Select two columns.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                TB_src_rng.Focus();
                // Me.Close()
                return;
            }

            else if (des_rng.Columns.Count != 2)
            {
                MessageBox.Show("Please, Select two columns.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                TB_des_rng.Focus();
                // Me.Close()
                return;
            }


            else if (x == false)
            {
                MessageBox.Show("Please select both of the columns that contain the data And the relevant images.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                TB_des_rng.Focus();

                return;
            }

            else
            {
                try
                {
                    // Set up validation list for 1st Column
                    Range rangeValues = (Range)src_rng.Columns[1].cells;
                    string listString = "";
                    // MsgBox(rangeValues.Address)
                    foreach (Range cell in rangeValues)
                    {
                        if (!string.IsNullOrEmpty(listString))
                        {
                            listString += ",";
                        }
                        listString = Conversions.ToString(listString + cell.get_Value());
                    }

                    // Set data validation in C1
                    validationRange = (Range)des_rng.Columns[1].cells;
                    {
                        var withBlock = validationRange.Validation;
                        withBlock.Delete(); // Delete any previous validation
                        withBlock.Add(Type: XlDVType.xlValidateList, AlertStyle: XlDVAlertStyle.xlValidAlertStop, Operator: XlFormatConditionOperator.xlBetween, Formula1: listString);
                        withBlock.IgnoreBlank = true;
                        withBlock.ShowInput = true;
                        withBlock.ShowError = true;
                    }
                    // MsgBox(2)
                    des_rng.Columns[2].ColumnWidth = src_rng.Columns[2].ColumnWidth;
                    des_rng.Rows.RowHeight = src_rng.Rows.RowHeight;



                    worksheet.Change += worksheet1_Change;

                    // 2 ta event handler dile valo vabe kaj korena. Seijonno ektar event handler er moddhe arekta call kora hoise.

                    // AddHandler worksheet.Change, AddressOf worksheet2_Change


                    Excel.Worksheet targetWorksheet = null;
                    foreach (Excel.Worksheet ws in excelApp.Worksheets)
                    {
                        if (ws.Name == "SoftekoPictureBasedDropDown")
                        {
                            targetWorksheet = ws;
                            break;
                        }
                    }

                    // If "MySpecialSheet" does not exist, add it
                    if (targetWorksheet is null)
                    {
                        targetWorksheet = (Excel.Worksheet)excelApp.Worksheets.Add(After: excelApp.Worksheets[excelApp.Worksheets.Count]);
                        targetWorksheet.Name = "SoftekoPictureBasedDropDown";
                    }


                    GlobalModule.Flag_Picture = true;
                    GlobalModule.sheetName2 = worksheet.Name;
                    GlobalModule.Src_Rng_of_PictureDDL = TB_src_rng.Text;
                    GlobalModule.Des_Rng_of_PictureDDL = TB_des_rng.Text;

                    // Write something in cell A1 of the target worksheet
                    targetWorksheet.get_Range("A1").set_Value(value: "Do not delete the sheet!");
                    targetWorksheet.get_Range("A2").set_Value(value: GlobalModule.Flag_Picture);
                    targetWorksheet.get_Range("A3").set_Value(value: GlobalModule.sheetName2);
                    targetWorksheet.get_Range("A4").set_Value(value: GlobalModule.Src_Rng_of_PictureDDL);
                    targetWorksheet.get_Range("A5").set_Value(value: GlobalModule.Des_Rng_of_PictureDDL);
                    targetWorksheet.Visible = XlSheetVisibility.xlSheetHidden;
                }

                catch (Exception ex)
                {
                }
                Dispose();
            }

        }


        private void worksheet2_Change(Range Target)
        {
            excelApp = Globals.ThisAddIn.Application;
            var workbook = excelApp.ActiveWorkbook;
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.ActiveSheet;

            try
            {

                foreach (Shape pic in worksheet.Shapes)
                {
                    // MsgBox(pic.TopLeftCell.Address)
                    if ((pic.TopLeftCell.get_Address() ?? "") == (Target.get_Offset(0, 1).get_Address() ?? ""))
                    {

                        pic.Delete();
                        // Exit For
                    }
                }
            }
            catch (Exception ex)
            {
            }
        }


        private void worksheet1_Change(Range Target)
        {

            excelApp = Globals.ThisAddIn.Application;
            var workbook = excelApp.ActiveWorkbook;
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.ActiveSheet;
            try
            {

                for (int i = 1, loopTo = src_rng.Rows.Count; i <= loopTo; i++)
                {
                    if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(src_rng[i, 1].Value, Target.get_Value(), false)))
                    {
                        // MsgBox(3)
                        try
                        {
                            worksheet2_Change(Target);
                        }
                        // MsgBox(5)
                        catch (Exception ex)
                        {
                            // MsgBox(15)
                        }

                        // MsgBox(6)

                        // Dim imageCell As Excel.Range = src_rng(i, 2)
                        // imageCell.CopyPicture(
                        // Appearance:=Excel.XlPictureAppearance.xlScreen,
                        // Format:=Excel.XlCopyPictureFormat.xlPicture)
                        // workSheet.Paste(Target.Offset(0, 1))

                        bool x = false;
                        // Try
                        foreach (Shape pic in worksheet.Shapes)
                        {
                            // MsgBox(pic.TopLeftCell.Address)
                            if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(pic.TopLeftCell.get_Address(), src_rng[i, 2].Address, false)))
                            {
                                pic.CopyPicture();
                                worksheet.Paste(Target.get_Offset(0, 1));
                                Target.get_Offset(0, 1).RowHeight = src_rng[i, 2].RowHeight;
                                // Target.Offset(0, 1).RowHeight = src_rng(i, 2).C
                                x = true;
                                break;
                            }
                            else
                            {
                                x = false;

                            }
                            // x = x + 1
                        }

                        excelApp.CutCopyMode = (XlCutCopyMode)Conversions.ToInteger(false);


                    }
                }
            }
            catch (Exception ex)
            {
            }


        }


        private void Form34_PictureBasedDropdownList_Load(object sender, EventArgs e)
        {
            try
            {

                excelApp = Globals.ThisAddIn.Application;
                var workbook = excelApp.ActiveWorkbook;
                Excel.Worksheet worksheet = (Excel.Worksheet)workbook.ActiveSheet;

                excelApp.SheetSelectionChange += excelApp_SheetSelectionChange;

                // opened = opened + 1

                if (excelApp.Selection is not null)
                {
                    selectedRange = (Range)excelApp.Selection;
                    src_rng = selectedRange;
                    TB_src_rng.Text = selectedRange.get_Address();

                }
                TB_src_rng.Focus();
            }

            catch (Exception ex)
            {
                TB_src_rng.Focus();
            }
        }

        private void excelApp_SheetSelectionChange(object Sh, Range selectionRange1)
        {
            try
            {

                excelApp = Globals.ThisAddIn.Application;
                if (focuschange == false)
                {
                    if (focuschange == false)
                    {
                        if (TB_des_rng.Focused == true | ReferenceEquals(ActiveControl, TB_des_rng))
                        {
                            if (TB_des_rng.Focused == true)
                            {
                                des_rng = selectionRange1;
                            }
                            Activate();
                            BeginInvoke(new System.Action(() =>
                                {
                                    TB_des_rng.Text = des_rng.get_Address();
                                    SetWindowPos(Handle, new IntPtr(HWND_TOPMOST), 0, 0, 0, 0, SWP_NOACTIVATE | SWP_NOMOVE | SWP_NOSIZE);
                                }));
                        }

                        // ElseIf Me.ActiveControl Is TB_src_range Then
                        else if (TB_src_rng.Focused == true | ReferenceEquals(ActiveControl, TB_src_rng))
                        {
                            if (TB_src_rng.Focused == true)
                            {
                                src_rng = selectionRange1;
                            }
                            Activate();
                            BeginInvoke(new System.Action(() =>
                                {
                                    TB_src_rng.Text = src_rng.get_Address();
                                    SetWindowPos(Handle, new IntPtr(HWND_TOPMOST), 0, 0, 0, 0, SWP_NOACTIVATE | SWP_NOMOVE | SWP_NOSIZE);
                                }));

                        }
                    }
                }
            }



            catch (Exception ex)
            {

            }

        }

        private void PictureBox9_Click(object sender, EventArgs e)
        {
            try
            {
                if (selectedRange is null)
                {
                }
                else
                {


                    TB_src_rng.Text = selectedRange.get_Address();


                    // FocusedTextBox = 1
                    Hide();

                    excelApp = Globals.ThisAddIn.Application;
                    workBook = excelApp.ActiveWorkbook;

                    Range userInput = (Range)excelApp.InputBox("Select a range", "Select a range", "=$A$1", Type: 8);
                    src_rng = userInput;

                    string sheetName;
                    sheetName = Strings.Split(src_rng.get_Address(true, true, XlReferenceStyle.xlA1, true), "]")[1];
                    sheetName = Strings.Split(sheetName, "!")[0];

                    if (Strings.Mid(sheetName, Strings.Len(sheetName), 1) == "'")
                    {
                        sheetName = Strings.Mid(sheetName, 1, Strings.Len(sheetName) - 1);
                    }
                    workSheet = (Excel.Worksheet)workBook.Worksheets[sheetName];
                    workSheet.Activate();

                    src_rng.Select();

                    TB_src_rng.Text = src_rng.get_Address();

                    Show();
                    TB_src_rng.Focus();


                    Range ran = (Range)src_rng[1, 1];





                }
            }

            catch (Exception ex)
            {

                Show();
                TB_src_rng.Focus();

            }
        }

        private void PictureBox1_Click(object sender, EventArgs e)
        {
            try
            {
                if (selectedRange is null)
                {
                }
                else
                {
                    // TB_src_range.Text = selectedRange.Address


                    Hide();

                    excelApp = Globals.ThisAddIn.Application;
                    workBook = excelApp.ActiveWorkbook;

                    // Dim userInput As String = excelApp.InputBox("Select a range", "Select range", "=$A$1")


                    Range userInput = (Range)excelApp.InputBox("Select a range", "Select a range", "=$A$1", Type: 8);
                    des_rng = userInput;

                    string sheetName;
                    sheetName = Strings.Split(des_rng.get_Address(true, true, XlReferenceStyle.xlA1, true), "]")[1];
                    sheetName = Strings.Split(sheetName, "!")[0];

                    if (Strings.Mid(sheetName, Strings.Len(sheetName), 1) == "'")
                    {
                        sheetName = Strings.Mid(sheetName, 1, Strings.Len(sheetName) - 1);
                    }

                    workSheet = (Excel.Worksheet)workBook.Worksheets[sheetName];
                    workSheet.Activate();

                    des_rng.Select();
                    // MsgBox(src_rng.Address)

                    TB_des_rng.Text = des_rng.get_Address();

                    Show();
                    TB_des_rng.Focus();

                }
            }
            catch (Exception ex)
            {
                TB_des_rng.Focus();
            }
        }



        private void Btn_Cancel_Click(object sender, EventArgs e)
        {
            Dispose();

        }



        private void TB_src_rng_TextChanged(object sender, EventArgs e)
        {
            try
            {

                if (TB_src_rng.Text is not null & IsValidExcelCellReference(TB_src_rng.Text) == true)
                {
                    focuschange = true;

                    // Define the range of cells to read (for example, cells A1 to A10)
                    src_rng = excelApp.get_Range(TB_src_rng.Text);
                    src_rng.Select();
                    var range = src_rng;

                    Activate();
                    // TB_src_range.Focus()
                    TB_src_rng.SelectionStart = TB_src_rng.Text.Length;
                    focuschange = false;

                }
            }

            catch (Exception ex)
            {

            }
        }
        private bool IsValidExcelCellReference(string cellReference)
        {

            // Regular expression pattern for a cell reference.
            // This pattern will match references like A1, $A$1, etc.
            string cellPattern = @"(\$?[A-Z]+\$?[0-9]+)";

            // Regular expression pattern for an Excel reference.
            // This pattern will match references like A1:B13, $A$1:$B$13, A1, $B$1, etc.
            string singleReferencePattern = cellPattern + "(:" + cellPattern + ")?";

            // Regular expression pattern to allow multiple cell references separated by commas
            string referencePattern = "^(" + singleReferencePattern + ")(," + singleReferencePattern + ")*$";

            // Create a regex object with the pattern.
            var regex = new Regex(referencePattern);

            // Test the input string against the regex pattern.
            return regex.IsMatch(cellReference.ToUpper());

        }


        private void TB_des_rng_TextChanged(object sender, EventArgs e)
        {
            try
            {

                if (TB_des_rng.Text is not null & IsValidExcelCellReference(TB_des_rng.Text) == true)
                {
                    focuschange = true;

                    // Define the range of cells to read (for example, cells A1 to A10)
                    des_rng = excelApp.get_Range(TB_des_rng.Text);
                    des_rng.Select();
                    var range = des_rng;

                    Activate();
                    // TB_src_range.Focus()
                    TB_des_rng.SelectionStart = TB_des_rng.Text.Length;
                    focuschange = false;

                }
            }

            catch (Exception ex)
            {

            }
        }

        private void source(object sender, KeyEventArgs e)
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

        private void Destination(object sender, KeyEventArgs e)
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

        private void source_TextBox(object sender, KeyEventArgs e)
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

        private void destination_TextBox(object sender, KeyEventArgs e)
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

        private void form_enter(object sender, KeyEventArgs e)
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

        private void Combobox1_enter(object sender, KeyEventArgs e)
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

        private void Form34_PictureBasedDropdownList_Closing(object sender, CancelEventArgs e)
        {
            GlobalModule.form_flag = false;
        }

        private void Form34_PictureBasedDropdownList_Disposed(object sender, EventArgs e)
        {
            GlobalModule.form_flag = false;
        }

        private void Form34_PictureBasedDropdownList_Shown(object sender, EventArgs e)
        {

            Focus();
            BringToFront();
            Activate();
            try
            {
                if (!string.IsNullOrEmpty(TB_src_rng.Text))
                {

                    BeginInvoke(new System.Action(() =>
                        {
                            TB_src_rng.Text = src_rng.get_Address();
                            SetWindowPos(Handle, new IntPtr(HWND_TOPMOST), 0, 0, 0, 0, SWP_NOACTIVATE | SWP_NOMOVE | SWP_NOSIZE);
                        }));


                }
            }

            catch (Exception ex)
            {
                TB_src_rng.Focus();

            }
        }

    }
}