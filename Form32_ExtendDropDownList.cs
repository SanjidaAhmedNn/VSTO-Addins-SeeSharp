using System;
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
    public partial class Form32_ExtendDropDownList
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
        private Excel.Worksheet worksheet3;
        private Range src_rng;
        public Range des_rng;
        private Range selectedRange;
        public string ax;
        public Range firstRow;

        private int opened;
        public bool focuschange;

        public Form32_ExtendDropDownList()
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

                    Show();
                    TB_src_rng.Focus();

                    firstRow = (Range)src_rng.Rows[1];
                    // MsgBox(firstRow.Address)
                }
            }

            catch (Exception ex)
            {

                Show();
                TB_src_rng.Focus();

            }
        }

        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Btn_OK.Focus();
                Btn_OK.PerformClick();
            }
        }

        private void Dest_selection_Click(object sender, EventArgs e)
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

        private void Form32_ExtendDropDownList_Load(object sender, EventArgs e)
        {
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

                    if (ReferenceEquals(ActiveControl, TB_des_rng))
                    {
                        des_rng = selectionRange1;
                        // This will run on the Excel thread, so you need to use Invoke to update the UI
                        // Me.BeginInvoke(New System.Action(Sub() TB_dest_range.Text = selectionRange1.Address))
                        Activate();
                        BeginInvoke(new System.Action(() =>
                            {
                                TB_des_rng.Text = des_rng.get_Address();
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

        private void Btn_OK_Click(object sender, EventArgs e)
        {

            excelApp = Globals.ThisAddIn.Application;
            var workbook = excelApp.ActiveWorkbook;
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.ActiveSheet;

            if (string.IsNullOrEmpty(TB_src_rng.Text) & string.IsNullOrEmpty(TB_des_rng.Text))
            {
                MessageBox.Show("Please select all necessary options.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                TB_src_rng.Focus();
                // Me.Close()
                return;
            }

            else if (string.IsNullOrEmpty(TB_src_rng.Text))
            {
                MessageBox.Show("Please, Select updated source range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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



            else if (string.IsNullOrEmpty(TB_des_rng.Text))
            {
                MessageBox.Show("Please, Select destination range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                TB_src_rng.Focus();
                // Me.Close()
                return;
            }
            else if (IsValidExcelCellReference(TB_des_rng.Text) == false)
            {
                MessageBox.Show("Select a valid Destination Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                TB_src_rng.Focus();
                // Me.Close()
                return;
            }


            else if (src_rng.Areas.Count > 1 | des_rng.Areas.Count > 1)
            {
                MessageBox.Show("Multiple selection is not possible in the Source Range field.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                TB_src_rng.Focus();
                return;
            }


            else if ((ax ?? "") != (workSheet2.Name ?? ""))
            {
                MessageBox.Show("Please select the range of the same worksheet", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                TB_des_rng.Focus();
                return;
            }

            // ElseIf src_rng.Column <> des_rng.Column Then
            // MessageBox.Show("1st column of source range and destination range should be same.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            // TB_des_rng.Focus()
            // Exit Sub

            else if (excelApp.Intersect(src_rng, des_rng) is null)
            {
                MessageBox.Show(" Please select a valid expanded dynamic drop-down list range that intersects each other.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                TB_src_rng.Focus();
                // Me.Close()
                return;
            }

            else
            {
                try
                {
                    var targetWorksheet = default(Excel.Worksheet);
                    // Dim i As Integer = 1
                    foreach (var ws in excelApp.ActiveWorkbook.Worksheets)
                    {
                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(ws.name, "MySpecialSheet", false)))
                        {
                            targetWorksheet = (Excel.Worksheet)ws;
                            break;
                        }
                    }

                    int k = 0;
                    // For i = 1 To targetWorksheet.Columns.Count
                    if (Conversions.ToBoolean(Operators.AndObject(Operators.ConditionalCompareObjectEqual(worksheet.Name, targetWorksheet.get_Range("A11").get_Value(), false), excelApp.Intersect(src_rng, excelApp.get_Range(targetWorksheet.get_Range("A2").get_Value())) is not null)))
                    {
                        GlobalModule.Variable1 = targetWorksheet.get_Range("A1").get_Value().ToString();
                        GlobalModule.Variable2 = targetWorksheet.get_Range("A2").get_Value().ToString();
                        GlobalModule.Header = Conversions.ToBoolean(targetWorksheet.get_Range("A3").get_Value().ToString());
                        GlobalModule.Ascending = Conversions.ToBoolean(targetWorksheet.get_Range("A4").get_Value().ToString());
                        GlobalModule.Descending = Conversions.ToBoolean(targetWorksheet.get_Range("A5").get_Value().ToString());
                        GlobalModule.TextConvert = Conversions.ToBoolean(targetWorksheet.get_Range("A6").get_Value().ToString());
                        GlobalModule.OptionType = Conversions.ToBoolean(targetWorksheet.get_Range("A7").get_Value().ToString());
                        GlobalModule.Horizontal_CreateDP = Conversions.ToBoolean(targetWorksheet.get_Range("A8").get_Value().ToString());
                        // Flag_CreateDDDL = targetWorksheet.Range("A9").Value.ToString
                        GlobalModule.sheetName10 = targetWorksheet.get_Range("A10").get_Value().ToString();
                        GlobalModule.sheetName11 = targetWorksheet.get_Range("A11").get_Value().ToString();
                        k = 1;
                    }

                    else if (Conversions.ToBoolean(Operators.AndObject(Operators.ConditionalCompareObjectEqual(worksheet.Name, targetWorksheet.get_Range("B11").get_Value(), false), excelApp.Intersect(src_rng, excelApp.get_Range(targetWorksheet.get_Range("B2").get_Value())) is not null)))
                    {
                        GlobalModule.Variable1 = targetWorksheet.get_Range("B1").get_Value().ToString();
                        GlobalModule.Variable2 = targetWorksheet.get_Range("B2").get_Value().ToString();
                        GlobalModule.Header = Conversions.ToBoolean(targetWorksheet.get_Range("B3").get_Value().ToString());
                        GlobalModule.Ascending = Conversions.ToBoolean(targetWorksheet.get_Range("B4").get_Value().ToString());
                        GlobalModule.Descending = Conversions.ToBoolean(targetWorksheet.get_Range("B5").get_Value().ToString());
                        GlobalModule.TextConvert = Conversions.ToBoolean(targetWorksheet.get_Range("B6").get_Value().ToString());
                        GlobalModule.OptionType = Conversions.ToBoolean(targetWorksheet.get_Range("B7").get_Value().ToString());
                        GlobalModule.Horizontal_CreateDP = Conversions.ToBoolean(targetWorksheet.get_Range("B8").get_Value().ToString());
                        // Flag_CreateDDDL = targetWorksheet.Range("B9").Value.ToString
                        GlobalModule.sheetName10 = targetWorksheet.get_Range("B10").get_Value().ToString();
                        GlobalModule.sheetName11 = targetWorksheet.get_Range("B11").get_Value().ToString();
                        k = 2;
                    }

                    else if (Conversions.ToBoolean(Operators.AndObject(Operators.ConditionalCompareObjectEqual(worksheet.Name, targetWorksheet.get_Range("C11").get_Value(), false), excelApp.Intersect(src_rng, excelApp.get_Range(targetWorksheet.get_Range("C2").get_Value())) is not null)))
                    {
                        GlobalModule.Variable1 = targetWorksheet.get_Range("C1").get_Value().ToString();
                        GlobalModule.Variable2 = targetWorksheet.get_Range("C2").get_Value().ToString();
                        GlobalModule.Header = Conversions.ToBoolean(targetWorksheet.get_Range("C3").get_Value().ToString());
                        GlobalModule.Ascending = Conversions.ToBoolean(targetWorksheet.get_Range("C4").get_Value().ToString());
                        GlobalModule.Descending = Conversions.ToBoolean(targetWorksheet.get_Range("C5").get_Value().ToString());
                        GlobalModule.TextConvert = Conversions.ToBoolean(targetWorksheet.get_Range("C6").get_Value().ToString());
                        GlobalModule.OptionType = Conversions.ToBoolean(targetWorksheet.get_Range("C7").get_Value().ToString());
                        GlobalModule.Horizontal_CreateDP = Conversions.ToBoolean(targetWorksheet.get_Range("C8").get_Value().ToString());
                        // Flag_CreateDDDL = targetWorksheet.Range("C9").Value.ToString
                        GlobalModule.sheetName10 = targetWorksheet.get_Range("C10").get_Value().ToString();
                        GlobalModule.sheetName11 = targetWorksheet.get_Range("C11").get_Value().ToString();
                        k = 3;
                    }

                    else if (Conversions.ToBoolean(Operators.AndObject(Operators.ConditionalCompareObjectEqual(worksheet.Name, targetWorksheet.get_Range("D11").get_Value(), false), excelApp.Intersect(src_rng, excelApp.get_Range(targetWorksheet.get_Range("D2").get_Value())) is not null)))
                    {
                        GlobalModule.Variable1 = targetWorksheet.get_Range("D1").get_Value().ToString();
                        GlobalModule.Variable2 = targetWorksheet.get_Range("D2").get_Value().ToString();
                        GlobalModule.Header = Conversions.ToBoolean(targetWorksheet.get_Range("D3").get_Value().ToString());
                        GlobalModule.Ascending = Conversions.ToBoolean(targetWorksheet.get_Range("D4").get_Value().ToString());
                        GlobalModule.Descending = Conversions.ToBoolean(targetWorksheet.get_Range("D5").get_Value().ToString());
                        GlobalModule.TextConvert = Conversions.ToBoolean(targetWorksheet.get_Range("D6").get_Value().ToString());
                        GlobalModule.OptionType = Conversions.ToBoolean(targetWorksheet.get_Range("D7").get_Value().ToString());
                        GlobalModule.Horizontal_CreateDP = Conversions.ToBoolean(targetWorksheet.get_Range("D8").get_Value().ToString());
                        // Flag_CreateDDDL = targetWorksheet.Range("D9").Value.ToString
                        GlobalModule.sheetName10 = targetWorksheet.get_Range("D10").get_Value().ToString();
                        GlobalModule.sheetName11 = targetWorksheet.get_Range("D11").get_Value().ToString();
                        k = 4;
                    }

                    else if (Conversions.ToBoolean(Operators.AndObject(Operators.ConditionalCompareObjectEqual(worksheet.Name, targetWorksheet.get_Range("E11").get_Value(), false), excelApp.Intersect(src_rng, excelApp.get_Range(targetWorksheet.get_Range("E2").get_Value())) is not null)))
                    {
                        GlobalModule.Variable1 = targetWorksheet.get_Range("E1").get_Value().ToString();
                        GlobalModule.Variable2 = targetWorksheet.get_Range("E2").get_Value().ToString();
                        GlobalModule.Header = Conversions.ToBoolean(targetWorksheet.get_Range("E3").get_Value().ToString());
                        GlobalModule.Ascending = Conversions.ToBoolean(targetWorksheet.get_Range("E4").get_Value().ToString());
                        GlobalModule.Descending = Conversions.ToBoolean(targetWorksheet.get_Range("E5").get_Value().ToString());
                        GlobalModule.TextConvert = Conversions.ToBoolean(targetWorksheet.get_Range("E6").get_Value().ToString());
                        GlobalModule.OptionType = Conversions.ToBoolean(targetWorksheet.get_Range("E7").get_Value().ToString());
                        GlobalModule.Horizontal_CreateDP = Conversions.ToBoolean(targetWorksheet.get_Range("E8").get_Value().ToString());
                        // Flag_CreateDDDL = targetWorksheet.Range("E9").Value.ToString
                        GlobalModule.sheetName10 = targetWorksheet.get_Range("E10").get_Value().ToString();
                        GlobalModule.sheetName11 = targetWorksheet.get_Range("E11").get_Value().ToString();
                        k = 5;

                    }


                    // Get the validation formula from the source cell
                    string validationFormula = Conversions.ToString(des_rng[1, 1].Validation.Formula1);

                    // Apply the validation to the target range
                    {
                        ref var withBlock = ref des_rng.Columns[1].Validation;
                        withBlock.Delete(); // Clear any existing validation
                        withBlock.Add(Type: XlDVType.xlValidateList, AlertStyle: XlDVAlertStyle.xlValidAlertStop, Operator: XlFormatConditionOperator.xlBetween, Formula1: validationFormula);
                        withBlock.IgnoreBlank = (object)true;
                        withBlock.InCellDropdown = (object)true;
                        withBlock.ShowInput = (object)true;
                        withBlock.ShowError = (object)true;
                    }

                    if (k == 1)
                    {

                        targetWorksheet.get_Range("A2").set_Value(value: excelApp.Union(worksheet.get_Range(targetWorksheet.get_Range("A2").get_Value()), des_rng).get_Address());
                    }
                    // Header = targetWorksheet.Range("A3").Value.ToString()
                    // Ascending = targetWorksheet.Range("A4").Value.ToString()
                    // Descending = targetWorksheet.Range("A5").Value.ToString()
                    // TextConvert = targetWorksheet.Range("A6").Value.ToString()
                    // OptionType = targetWorksheet.Range("A7").Value.ToString()
                    // Horizontal_CreateDP = targetWorksheet.Range("A8").Value.ToString()
                    // Flag_CreateDDDL = targetWorksheet.Range("A9").Value.ToString
                    // sheetName10 = targetWorksheet.Range("A10").Value.ToString
                    // sheetName11 = targetWorksheet.Range("A11").Value.ToString

                    else if (k == 2)
                    {
                        targetWorksheet.get_Range("B2").set_Value(value: excelApp.Union(worksheet.get_Range(targetWorksheet.get_Range("B2").get_Value()), des_rng).get_Address());
                    }
                    // Header = targetWorksheet.Range("B3").Value.ToString()
                    // Ascending = targetWorksheet.Range("B4").Value.ToString()
                    // Descending = targetWorksheet.Range("B5").Value.ToString()
                    // TextConvert = targetWorksheet.Range("B6").Value.ToString()
                    // OptionType = targetWorksheet.Range("B7").Value.ToString()
                    // Horizontal_CreateDP = targetWorksheet.Range("B8").Value.ToString()
                    // Flag_CreateDDDL = targetWorksheet.Range("B9").Value.ToString
                    // sheetName10 = targetWorksheet.Range("B10").Value.ToString
                    // sheetName11 = targetWorksheet.Range("B11").Value.ToString

                    else if (k == 3)
                    {
                        targetWorksheet.get_Range("C2").set_Value(value: excelApp.Union(worksheet.get_Range(targetWorksheet.get_Range("C2").get_Value()), des_rng).get_Address());
                    }
                    // Header = targetWorksheet.Range("C3").Value.ToString()
                    // Ascending = targetWorksheet.Range("C4").Value.ToString()
                    // Descending = targetWorksheet.Range("C5").Value.ToString()
                    // TextConvert = targetWorksheet.Range("C6").Value.ToString()
                    // OptionType = targetWorksheet.Range("C7").Value.ToString()
                    // Horizontal_CreateDP = targetWorksheet.Range("C8").Value.ToString()
                    // Flag_CreateDDDL = targetWorksheet.Range("C9").Value.ToString
                    // sheetName10 = targetWorksheet.Range("C10").Value.ToString
                    // sheetName11 = targetWorksheet.Range("C11").Value.ToString

                    else if (k == 4)
                    {
                        targetWorksheet.get_Range("D2").set_Value(value: excelApp.Union(worksheet.get_Range(targetWorksheet.get_Range("D2").get_Value()), des_rng).get_Address());
                    }
                    // Header = targetWorksheet.Range("D3").Value.ToString()
                    // Ascending = targetWorksheet.Range("D4").Value.ToString()
                    // Descending = targetWorksheet.Range("D5").Value.ToString()
                    // TextConvert = targetWorksheet.Range("D6").Value.ToString()
                    // OptionType = targetWorksheet.Range("D7").Value.ToString()
                    // Horizontal_CreateDP = targetWorksheet.Range("D8").Value.ToString()
                    // Flag_CreateDDDL = targetWorksheet.Range("D9").Value.ToString
                    // sheetName10 = targetWorksheet.Range("D10").Value.ToString
                    // sheetName11 = targetWorksheet.Range("D11").Value.ToString

                    else if (k == 5)
                    {
                        targetWorksheet.get_Range("E2").set_Value(value: excelApp.Union(worksheet.get_Range(targetWorksheet.get_Range("E2").get_Value()), des_rng).get_Address());
                        // Header = targetWorksheet.Range("E3").Value.ToString()
                        // Ascending = targetWorksheet.Range("E4").Value.ToString()
                        // Descending = targetWorksheet.Range("E5").Value.ToString()
                        // TextConvert = targetWorksheet.Range("E6").Value.ToString()
                        // OptionType = targetWorksheet.Range("E7").Value.ToString()
                        // Horizontal_CreateDP = targetWorksheet.Range("E8").Value.ToString()
                        // Flag_CreateDDDL = targetWorksheet.Range("E9").Value.ToString
                        // sheetName10 = targetWorksheet.Range("E10").Value.ToString
                        // sheetName11 = targetWorksheet.Range("E11").Value.ToString
                    }
                    src_rng.Select();
                    Refresh();
                    Hide();
                    MessageBox.Show("Your Dynamic Drop-down List is extended successfully.", "Softeko", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    Close();
                }
                catch (Exception ex)
                {
                    Close();
                }
            }
        }

        private void Btn_Cancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void Form32_ExtendDropDownList_Disposed(object sender, EventArgs e)
        {
            GlobalModule.form_flag = false;
        }

        private void Form32_ExtendDropDownList_Closing(object sender, CancelEventArgs e)
        {
            GlobalModule.form_flag = false;
        }

        private void Form32_ExtendDropDownList_Shown(object sender, EventArgs e)
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

        private void TB_des_rng_TextChanged(object sender, EventArgs e)
        {
            excelApp = Globals.ThisAddIn.Application;
            var workbook = excelApp.ActiveWorkbook;
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.ActiveSheet;
            try
            {
                if (TB_des_rng.Text is not null & IsValidExcelCellReference(TB_des_rng.Text) == true)
                {
                    focuschange = true;
                    string sheetname = "";

                    try
                    {

                        des_rng = worksheet.get_Range(TB_des_rng.Text);
                        des_rng.Select();
                    }
                    catch
                    {
                        // Split the string into sheet name and cell address
                        string[] parts = TB_des_rng.Text.Split('!');
                        sheetname = parts[0];
                        string cellAddress = parts[1];
                        worksheet = (Excel.Worksheet)workbook.Worksheets[sheetname];
                        worksheet.Activate();
                        des_rng = worksheet.get_Range(cellAddress);
                        des_rng.Select();
                    }

                    if ((workSheet2.Name ?? "") != (worksheet.Name ?? "") & TB_des_rng.Text.Contains("!") == false)
                    {

                        TB_des_rng.Text = worksheet.Name + "!" + TB_des_rng.Text;

                    }

                    Activate();
                    TB_des_rng.Focus();
                    TB_des_rng.SelectionStart = TB_des_rng.Text.Length;

                    focuschange = false;
                    ax = worksheet.Name;
                    worksheet3 = worksheet;
                }
            }
            catch (Exception ex)
            {
                focuschange = false;
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

        private void TB_src_rng_TextChanged(object sender, EventArgs e)
        {
            excelApp = Globals.ThisAddIn.Application;
            workBook = excelApp.ActiveWorkbook;
            workSheet = (Excel.Worksheet)workBook.ActiveSheet;

            if (TB_src_rng.Text is not null & IsValidExcelCellReference(TB_src_rng.Text.ToUpper()) == true)
            {
                focuschange = true;

                // Define the range of cells to read (for example, cells A1 to A10)
                src_rng = excelApp.get_Range(TB_src_rng.Text);
                src_rng.Select();

                Activate();
                // TB_src_range.Focus()
                TB_src_rng.SelectionStart = TB_src_rng.Text.Length;
                focuschange = false;
                workSheet2 = workSheet;
            }
        }
    }
}