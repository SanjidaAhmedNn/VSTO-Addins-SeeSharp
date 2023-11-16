using System;
using System.Collections.Generic;
using System.ComponentModel;
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
    public partial class Form31_UpdateDynamicDropdownList
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
        private Excel.Worksheet workSheet2;
        private Excel.Worksheet workSheet3;
        private Range src_rng;
        public Range des_rng;
        private Range selectedRange;
        public string ax;
        public bool focuschange;
        private Form30_Create_Dynamic_Drop_down_List form;


        private int opened;

        public Form31_UpdateDynamicDropdownList()
        {
            InitializeComponent();
        }


        [DllImport("user32")]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
        private const uint SWP_NOMOVE = 0x2U;
        private const uint SWP_NOSIZE = 0x1U;
        private const uint SWP_NOACTIVATE = 0x10U;
        private const int HWND_TOPMOST = -1;


        private void Selection_source_Click(object sender, EventArgs e)
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
                    GlobalModule.Variable1 = TB_src_rng.Text;
                    // MsgBox(Variable1)
                    Show();
                    TB_src_rng.Focus();
                }
            }

            catch (Exception ex)
            {

                Show();
                TB_src_rng.Focus();

            }
        }

        private void PictureBox3_Click(object sender, EventArgs e)
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

                TB_des_rng2.Text = des_rng.get_Address();

                Show();
                TB_des_rng2.Focus();

            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // Dim workbook As Excel.Workbook = excelApp.ActiveWorkbook
            // Dim worksheet As Excel.Worksheet = workbook.ActiveSheet
            try
            {

                excelApp = Globals.ThisAddIn.Application;

                excelApp.SheetSelectionChange += excelApp_SheetSelectionChange;

                opened = opened + 1;

                if (excelApp.Selection is not null)
                {
                    selectedRange = (Range)excelApp.Selection;
                    src_rng = selectedRange;
                    TB_src_rng.Text = selectedRange.get_Address();
                }
                else
                {
                    selectedRange = excelApp.get_Range(GlobalModule.Variable1);
                    src_rng = selectedRange;
                    TB_src_rng.Text = selectedRange.get_Address();

                }


                if (RB_same_source.Checked == true)
                {
                    TB_des_rng1.Enabled = true;
                    TB_des_rng2.Enabled = false;
                    PictureBox3.Enabled = false;
                    PictureBox2.Enabled = false;
                    L_select.Enabled = false;
                    if (GlobalModule.Variable2 is not null)
                    {
                        TB_des_rng1.Text = GlobalModule.Variable2;
                        des_rng = (Range)excelApp.ActiveSheet.Range(TB_des_rng1.Text);
                    }
                }

                else if (RB_diff_rng.Checked == true)
                {
                    TB_des_rng1.Enabled = false;
                    TB_des_rng2.Enabled = true;
                    PictureBox3.Enabled = true;
                    PictureBox2.Enabled = true;
                    L_select.Enabled = true;

                }
            }

            catch (Exception ex)
            {

            }

        }


        private void excelApp_SheetSelectionChange(object Sh, Range selectionRange1)
        {
            try
            {

                excelApp = Globals.ThisAddIn.Application;
                if (focuschange == false)
                {

                    if (ReferenceEquals(ActiveControl, TB_des_rng2))
                    {
                        des_rng = selectionRange1;
                        // This will run on the Excel thread, so you need to use Invoke to update the UI
                        // Me.BeginInvoke(New System.Action(Sub() TB_dest_range.Text = selectionRange1.Address))
                        Activate();
                        BeginInvoke(new System.Action(() =>
                            {
                                TB_des_rng2.Text = des_rng.get_Address();
                                SetWindowPos(Handle, new IntPtr(HWND_TOPMOST), 0, 0, 0, 0, SWP_NOACTIVATE | SWP_NOMOVE | SWP_NOSIZE);
                            }));
                    }

                    else if (ReferenceEquals(ActiveControl, TB_src_rng))
                    {
                        src_rng = selectionRange1;
                        Activate();


                        BeginInvoke(new System.Action(() =>
                            {
                                TB_src_rng.Text = src_rng.get_Address();
                                SetWindowPos(Handle, new IntPtr(HWND_TOPMOST), 0, 0, 0, 0, SWP_NOACTIVATE | SWP_NOMOVE | SWP_NOSIZE);
                            }));
                    }

                }
            }

            catch (Exception ex)
            {

            }

        }

        private void Btn_Cancel_Click(object sender, EventArgs e)
        {
            Close();
        }


        private void Btn_OK_Click(object sender, EventArgs e)
        {
            excelApp = Globals.ThisAddIn.Application;
            var workbook = excelApp.ActiveWorkbook;
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.ActiveSheet;

            Range r1;
            r1 = workSheet2.get_Range(TB_src_rng.Text);


            if (string.IsNullOrEmpty(TB_src_rng.Text) & string.IsNullOrEmpty(TB_des_rng2.Text) & TB_des_rng2.Enabled == true)
            {
                MessageBox.Show("Please, Select updated source range and destination range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                TB_src_rng.Focus();
                // Me.Close()
                return;
            }

            else if (string.IsNullOrEmpty(TB_src_rng.Text))
            {
                // MsgBox(100)
                MessageBox.Show("Check your Updated Source Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                TB_src_rng.Focus();
                // Me.Close()
                return;
            }

            else if (IsValidExcelCellReference(TB_src_rng.Text) == false)
            {
                MessageBox.Show("Select a valid Source Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                TB_src_rng.Focus();
                // Me.Close()
                return;
            }



            else if (TB_des_rng2.Enabled == true & string.IsNullOrEmpty(TB_des_rng2.Text))
            {
                MessageBox.Show("Please, Select destination range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                TB_src_rng.Focus();
                // Me.Close()
                return;
            }
            else if (TB_des_rng2.Enabled == true & IsValidExcelCellReference(TB_des_rng2.Text) == false)
            {
                MessageBox.Show("Select a valid Destination Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                TB_src_rng.Focus();
                // Me.Close()
                return;
            }

            else if (RB_diff_rng.Checked == false & RB_same_source.Checked == false)
            {
                MessageBox.Show("Select Destination Range", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                TB_des_rng2.Focus();
                // Me.Close()
                return;
            }

            else if (src_rng.Areas.Count > 1)
            {
                MessageBox.Show("Please Select dynamic drop-down list range from same worksheet.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                TB_src_rng.Focus();
            }

            else if (r1.Columns.Count != des_rng.Columns.Count)
            {
                MessageBox.Show("Check your Updated Source Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                TB_des_rng2.Focus();
            }
            else
            {
                try
                {
                    var result = MessageBox.Show("The Original Source Range is :" + GlobalModule.Variable1 + ". AND the Drop-down list is in :" + GlobalModule.Variable2 + "Do you want to continue?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    // Check if the user clicked 'Yes'
                    if (result == DialogResult.Yes)
                    {

                        GlobalModule.Variable1 = TB_src_rng.Text;
                        if ((des_rng.Worksheet.Name ?? "") != (src_rng.Worksheet.Name ?? ""))
                        {
                            GlobalModule.Variable1 = src_rng.Worksheet.Name + "!" + TB_src_rng.Text;
                            GlobalModule.Variable2 = des_rng.Worksheet.Name + "!" + des_rng.get_Address();
                        }

                        else
                        {
                            GlobalModule.Variable1 = src_rng.Worksheet.Name + "!" + TB_src_rng.Text;
                            GlobalModule.Variable2 = des_rng.Worksheet.Name + "!" + des_rng.get_Address();
                        }

                        OutPut();                                                                        // Main Function

                        var targetWorksheet = default(Excel.Worksheet);

                        foreach (var ws in excelApp.ActiveWorkbook.Worksheets)
                        {
                            if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(ws.name, "MySpecialSheet", false)))
                            {
                                targetWorksheet = (Excel.Worksheet)ws;
                                break;
                            }
                        }
                        if (Conversions.ToDouble(TextBox1.Text) == 1d)
                        {

                            targetWorksheet.get_Range("A1").set_Value(value: GlobalModule.Variable1);
                            targetWorksheet.get_Range("A2").set_Value(value: GlobalModule.Variable2);
                            targetWorksheet.get_Range("A10").set_Value(value: GlobalModule.sheetName10);
                            targetWorksheet.get_Range("A11").set_Value(value: GlobalModule.sheetName11);
                        }

                        else if (Conversions.ToDouble(TextBox1.Text) == 2d)
                        {
                            targetWorksheet.get_Range("B1").set_Value(value: GlobalModule.Variable1);
                            targetWorksheet.get_Range("B2").set_Value(value: GlobalModule.Variable2);
                            targetWorksheet.get_Range("B10").set_Value(value: GlobalModule.sheetName10);
                            targetWorksheet.get_Range("B11").set_Value(value: GlobalModule.sheetName11);
                        }

                        else if (Conversions.ToDouble(TextBox1.Text) == 3d)
                        {
                            targetWorksheet.get_Range("C1").set_Value(value: GlobalModule.Variable1);
                            targetWorksheet.get_Range("C2").set_Value(value: GlobalModule.Variable2);
                            targetWorksheet.get_Range("C10").set_Value(value: GlobalModule.sheetName10);
                            targetWorksheet.get_Range("C11").set_Value(value: GlobalModule.sheetName11);
                        }

                        else if (Conversions.ToDouble(TextBox1.Text) == 4d)
                        {
                            targetWorksheet.get_Range("D1").set_Value(value: GlobalModule.Variable1);
                            targetWorksheet.get_Range("D2").set_Value(value: GlobalModule.Variable2);
                            targetWorksheet.get_Range("D10").set_Value(value: GlobalModule.sheetName10);
                            targetWorksheet.get_Range("D11").set_Value(value: GlobalModule.sheetName11);
                        }

                        else if (Conversions.ToDouble(TextBox1.Text) == 5d)
                        {
                            targetWorksheet.get_Range("E1").set_Value(value: GlobalModule.Variable1);
                            targetWorksheet.get_Range("E2").set_Value(value: GlobalModule.Variable2);
                            targetWorksheet.get_Range("E10").set_Value(value: GlobalModule.sheetName10);
                            targetWorksheet.get_Range("E11").set_Value(value: GlobalModule.sheetName11);
                        }

                    }
                    Close();
                }
                catch (Exception ex)
                {
                    des_rng.Select();
                    Close();
                }
            }

        }

        private void TB_dest_range_Enter(object sender, KeyEventArgs e)
        {
            // If Enter key is pressed then check if the text is a valid address
            if (IsValidExcelCellReference(TB_des_rng2.Text) == true & e.KeyCode == Keys.Enter)
            {
                des_rng = excelApp.get_Range(TB_des_rng2.Text);
                TB_des_rng2.Focus();
                des_rng.Select();

                Btn_OK_Click(sender, e);   // OK button click event called
            }

            // MsgBox(des_rng.Address)
            else if (IsValidExcelCellReference(TB_des_rng2.Text) == false & e.KeyCode == Keys.Enter)
            {
                MessageBox.Show("Select the valid Destination Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                TB_des_rng2.Text = "";
                TB_des_rng2.Focus();
                // Me.Close()
                return;
            }
        }

        private void TB_src_range_Enter(object sender, KeyEventArgs e)
        {
            // If Enter key is pressed then check if the text is a valid address

            if (IsValidExcelCellReference(TB_src_rng.Text) == true & e.KeyCode == Keys.Enter)
            {
                src_rng = excelApp.get_Range(TB_src_rng.Text);
                TB_src_rng.Focus();
                src_rng.Select();

                Btn_OK_Click(sender, e);   // OK button click event called
            }

            // MsgBox(des_rng.Address)
            else if (IsValidExcelCellReference(TB_src_rng.Text) == false & e.KeyCode == Keys.Enter)
            {
                MessageBox.Show("Select the valid Source Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                TB_src_rng.Text = "";
                TB_src_rng.Focus();
                // Me.Close()
                return;
            }
        }


        public void CreateValidationList(Range cell, string listValues)
        {
            {
                var withBlock = cell.Validation;
                withBlock.Delete();
                withBlock.Add(Type: XlDVType.xlValidateList, AlertStyle: XlDVAlertStyle.xlValidAlertStop, Operator: XlFormatConditionOperator.xlBetween, Formula1: listValues);
                withBlock.ShowInput = true;
                withBlock.ShowError = true;
            }
        }
        private void OutPut()
        {

            try
            {
                excelApp = Globals.ThisAddIn.Application;
                var workbook = excelApp.ActiveWorkbook;
                Excel.Worksheet worksheet = (Excel.Worksheet)workbook.ActiveSheet;


                Range rng;
                if (GlobalModule.Header == true)
                {
                    // Dim adjustRange As Excel.Range
                    rng = src_rng.get_Offset(1, 0).get_Resize(src_rng.Rows.Count - 1, src_rng.Columns.Count);
                }

                else
                {

                    rng = src_rng;
                } // Assuming you have a range from A1 to A100

                var uniqueValues = new List<string>();

                // Extract unique values from the range
                foreach (Range cell in (System.Collections.IEnumerable)rng.Columns[1].Cells)
                {
                    string value = Conversions.ToString(cell.get_Value());
                    if (!uniqueValues.Contains(value))
                    {
                        uniqueValues.Add(value);
                    }
                }

                if (GlobalModule.Ascending == true)
                {
                    // Sort the list in ascending order
                    uniqueValues.Sort();
                }
                else if (GlobalModule.Descending == true)
                {
                    // Sort the list in ascending order
                    uniqueValues.Sort();
                    uniqueValues.Reverse();
                }

                // Create drop-down list at B1 with the unique values
                Range dropDownRange = (Range)des_rng.Columns[1];
                var validation = dropDownRange.Validation;
                validation.Delete(); // Remove any existing validation
                validation.Add(XlDVType.xlValidateList, Formula1: string.Join(",", uniqueValues));
                var range1 = excelApp.get_Range(TB_des_rng1.Text);
                // Dim range2 As Excel.Range = range1.Rows(1)
                // MsgBox(range1.Address)
                // MsgBox(des_rng.Address)
                if (RB_diff_rng.Checked == true & (range1.get_Address(1, 1) ?? "") != (des_rng.get_Address(1, 1) ?? ""))
                {

                    // MsgBox(range1.Address)
                    // If des_rng.Rows.Count < range1.Rows.Count Then
                    // Dim difference As Integer = range1.Rows.Count - des_rng.Rows.Count
                    // Dim startRowToDelete As Integer = range1.Rows.Count - difference + 1
                    // Dim endRowToDelete As Integer = range1.Rows.Count
                    // range1.Rows(String.Format("{0}:{1}", startRowToDelete, endRowToDelete)).Delete(Excel.XlDeleteShiftDirection.xlShiftUp)
                    // range1 = range1.Resize(des_rng.Rows.Count, range1.Columns.Count)
                    // End If

                    // range1.Rows(1).Cut(des_rng)

                    // For i As Integer = 1 To des_rng.Rows.Count

                    // range1.Rows(1).Copy(des_rng.Rows(i))
                    // Next

                    // des_rng.Rows(1).cut(range1.Rows(1))


                    form = new Form30_Create_Dynamic_Drop_down_List();
                    form.TB_src_range.Text = TB_src_rng.Text;
                    form.TB_dest_range.Text = TB_des_rng2.Text;
                    if (GlobalModule.OptionType == true)
                    {
                        form.RB_Dropdown_35_Labels.Checked = true;
                    }
                    if (GlobalModule.Header == true)
                    {
                        form.CB_header.Checked = true;
                    }
                    if (GlobalModule.Ascending == true)
                    {
                        form.CB_ascending.Checked = true;
                    }
                    if (GlobalModule.Descending == true)
                    {
                        form.CB_descending.Checked = true;
                    }
                    if (GlobalModule.TextConvert == true)
                    {
                        form.CB_text.Checked = true;
                    }
                    if (GlobalModule.Horizontal_CreateDP == true)
                    {
                        form.RB_Horizon.Checked = true;
                    }

                    form.Btn_OK_Click(form.Btn_OK, new EventArgs());

                }


                GlobalModule.Variable1 = TB_src_rng.Text;
                if (RB_diff_rng.Checked == true)
                {
                    GlobalModule.Variable2 = TB_des_rng2.Text;
                }
                des_rng.Select();

                des_rng.set_Value(value: null);
                GlobalModule.sheetName10 = workSheet2.Name;
                if (RB_diff_rng.Checked == true)
                {
                    GlobalModule.sheetName11 = workSheet3.Name;
                }
            }

            catch (Exception ex)
            {

            }
        }


        private bool IsValidExcelCellReference(string cellReference)
        {

            // Regular expression pattern for a valid sheet name. This is a simplified version and might not cover all edge cases.
            // Excel sheet names cannot contain the characters \, /, *, [, ], :, ?, and cannot be 'History'.
            string sheetNamePattern = @"(?i)(?![\/*[\]:?])(?!History)[^\/\[\]*?:\\]+";

            // Regular expression pattern for a cell reference.
            // This pattern will match references like A1, $A$1, etc.
            string cellPattern = @"(\$?[A-Z]+\$?[0-9]+)";

            // Regular expression pattern for an Excel reference.
            // This pattern will match references like A1:B13, $A$1:$B$13, A1, $B$1, etc.
            string singleReferencePattern = cellPattern + "(:" + cellPattern + ")?";

            // Regular expression pattern to allow the sheet name, followed by '!', before the cell reference
            string fullPattern = "^(" + sheetNamePattern + "!)?(" + singleReferencePattern + ")(," + singleReferencePattern + ")*$";

            // Create a regex object with the pattern.
            var regex = new Regex(fullPattern);

            // Test the input string against the regex pattern.
            return regex.IsMatch(cellReference.ToUpper());

        }


        private void RB_same_source_CheckedChanged(object sender, EventArgs e)
        {
            excelApp = Globals.ThisAddIn.Application;
            if (RB_same_source.Checked == true)
            {

                TB_des_rng1.Enabled = true;

                TB_des_rng2.Enabled = false;
                PictureBox3.Enabled = false;
                PictureBox2.Enabled = false;
                L_select.Enabled = false;
                // MsgBox(L_select.Enabled)
                if (GlobalModule.Variable2 is not null)
                {
                    TB_des_rng1.Text = GlobalModule.Variable2;
                    // MsgBox(Variable2)
                    des_rng = excelApp.get_Range(GlobalModule.Variable2);
                }
            }
        }

        private void RB_diff_rng_CheckedChanged(object sender, EventArgs e)
        {
            if (RB_diff_rng.Checked == true)
            {
                TB_des_rng1.Enabled = false;
                TB_des_rng2.Enabled = true;
                PictureBox3.Enabled = true;
                PictureBox2.Enabled = true;
                L_select.Enabled = true;
                TB_des_rng2.Focus();

            }

        }

        private void OK(object sender, KeyEventArgs e)
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

        private void Cancel(object sender, KeyEventArgs e)
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

        private void Form_load(object sender, KeyEventArgs e)
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

        private void RB_Different(object sender, KeyEventArgs e)
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

        private void RB_same(object sender, KeyEventArgs e)
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

        private void CustomGroupBox2_Enter(object sender, EventArgs e)
        {

        }

        private void TB_src_rng_TextChanged(object sender, EventArgs e)
        {
            excelApp = Globals.ThisAddIn.Application;
            workBook = excelApp.ActiveWorkbook;
            workSheet = (Excel.Worksheet)workBook.ActiveSheet;

            try
            {
                if (TB_src_rng.Text is not null & IsValidExcelCellReference(TB_src_rng.Text) == true)
                {
                    focuschange = true;

                    // Define the range of cells to read (for example, cells A1 to A10)
                    TB_src_rng.Text = TB_src_rng.Text.ToUpper();
                    src_rng = excelApp.get_Range(TB_src_rng.Text);
                    src_rng.Select();
                    var range = src_rng;


                    Activate();
                    // TB_src_range.Focus()
                    TB_src_rng.SelectionStart = TB_src_rng.Text.Length;
                    focuschange = false;
                    workSheet2 = workSheet;


                }
            }
            catch (Exception ex)
            {
            }
        }

        private void TB_des_rng2_TextChanged(object sender, EventArgs e)
        {

            excelApp = Globals.ThisAddIn.Application;
            var workbook = excelApp.ActiveWorkbook;
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.ActiveSheet;
            try
            {
                if (TB_des_rng2.Text is not null & IsValidExcelCellReference(TB_des_rng2.Text) == true)
                {
                    focuschange = true;

                    // Define the range of cells to read (for example, cells A1 to A10)
                    try
                    {
                        TB_des_rng2.Text = TB_des_rng2.Text;
                        des_rng = excelApp.get_Range(TB_des_rng2.Text);
                        des_rng.Select();
                    }

                    catch (Exception ex)
                    {
                        // Split the string into sheet name and cell address
                        string[] parts = TB_des_rng2.Text.Split('!');
                        string sheetName = parts[0];
                        string cellAddress = parts[1];

                        des_rng = excelApp.get_Range(cellAddress);
                        des_rng.Select();

                    }

                    if ((workSheet2.Name ?? "") != (worksheet.Name ?? ""))
                    {
                        TB_des_rng2.Text = worksheet.Name + "!" + des_rng.get_Address();
                        // src_rng = excelApp.Range(TB_src_range.Text)


                    }
                    Activate();
                    TB_des_rng2.SelectionStart = TB_des_rng2.Text.Length;
                    focuschange = false;
                    ax = worksheet.Name;
                    workSheet3 = worksheet;
                    // MsgBox(workSheet3.Name)
                }
            }
            catch (Exception ex)
            {
                ax = "";
                workSheet3 = worksheet;
            }
        }

        private void Form31_UpdateDynamicDropdownList_Closing(object sender, CancelEventArgs e)
        {
            GlobalModule.form_flag = false;
        }

        private void Form31_UpdateDynamicDropdownList_Disposed(object sender, EventArgs e)
        {
            GlobalModule.form_flag = false;
        }

        private void Form31_UpdateDynamicDropdownList_Shown(object sender, EventArgs e)
        {
            Focus();
            BringToFront();
            Activate();
            BeginInvoke(new System.Action(() =>
                {
                    TB_src_rng.Text = src_rng.get_Address();
                    SetWindowPos(Handle, new IntPtr(HWND_TOPMOST), 0, 0, 0, 0, SWP_NOACTIVATE | SWP_NOMOVE | SWP_NOSIZE);
                }));
        }

    }
}