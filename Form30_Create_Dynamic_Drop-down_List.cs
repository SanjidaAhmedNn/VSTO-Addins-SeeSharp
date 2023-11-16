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


    public partial class Form30_Create_Dynamic_Drop_down_List
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

        private int opened;

        public Form30_Create_Dynamic_Drop_down_List()
        {
            InitializeComponent();
        }
        // Public WithEvents Btn_OK As System.Windows.Forms.Button


        [DllImport("user32")]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
        private const uint SWP_NOMOVE = 0x2U;
        private const uint SWP_NOSIZE = 0x1U;
        private const uint SWP_NOACTIVATE = 0x10U;
        private const int HWND_TOPMOST = -1;

        private void CB_ascending_CheckedChanged(object sender, EventArgs e)
        {
            if (CB_ascending.Checked == true)
            {
                CB_descending.Checked = false;
            }
        }

        private void CB_descending_CheckedChanged(object sender, EventArgs e)
        {
            if (CB_descending.Checked == true)
            {
                CB_ascending.Checked = false;
            }
        }

        private void RB_columns_CheckedChanged(object sender, EventArgs e)
        {
            if (RB_Dropdown_35_Labels.Checked == true)
            {

                CB_header.Enabled = true;
                CB_ascending.Enabled = true;
                CB_descending.Enabled = true;
                CB_text.Enabled = true;
                GB_list_option.Enabled = false;

            }
        }

        private void RB_rows_CheckedChanged(object sender, EventArgs e)
        {
            if (RB_Dropdown_2_Labels.Checked == true)
            {
                GB_list_option.Enabled = true;
                CB_header.Enabled = false;
                CB_ascending.Enabled = false;
                CB_descending.Enabled = false;
                CB_text.Enabled = false;

            }
        }



        private void Selection_source_Click(object sender, EventArgs e)
        {
            try
            {
                if (selectedRange is null)
                {
                }
                else
                {

                    TB_src_range.Text = selectedRange.get_Address();


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

                    TB_src_range.Text = src_rng.get_Address();

                    Show();
                    TB_src_range.Focus();
                }
            }

            catch (Exception ex)
            {

                Show();
                TB_src_range.Focus();

            }
        }

        // Event handler to detect changes in E1 and adjust dropdown in E2

        public void Btn_OK_Click(object sender, EventArgs e)
        {

            excelApp = Globals.ThisAddIn.Application;
            var workbook = excelApp.ActiveWorkbook;
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.ActiveSheet;

            if (string.IsNullOrEmpty(TB_src_range.Text) & string.IsNullOrEmpty(TB_dest_range.Text))
            {
                MessageBox.Show("Please select all necessary options.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                TB_src_range.Focus();
                // Me.Close()
                return;
            }

            else if (string.IsNullOrEmpty(TB_src_range.Text))
            {
                MessageBox.Show("Please select the Source Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                TB_src_range.Focus();
                // Me.Close()
                return;
            }
            // End If

            else if (IsValidExcelCellReference(TB_src_range.Text) == false)
            {
                MessageBox.Show("Select a Valid Source Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                TB_src_range.Focus();
                // Me.Close()
                return;
            }
            // End If

            else if (string.IsNullOrEmpty(TB_dest_range.Text))
            {
                MessageBox.Show("Select a Destination Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                TB_dest_range.Focus();
                // Me.Close()
                return;
            }
            // End If

            else if (IsValidExcelCellReference(TB_dest_range.Text) == false)
            {
                MessageBox.Show("Select a Valid Destination Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                TB_dest_range.Focus();
                // Me.Close()
                return;
            }
            // End If


            else if (RB_Dropdown_2_Labels.Checked == false & RB_Dropdown_35_Labels.Checked == false)
            {
                MessageBox.Show("Select a Drop-down List type.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                worksheet.Activate();
                src_rng.Select();
                // Me.Close()
                // Exit Sub
                return;
            }

            else if (RB_Dropdown_2_Labels.Checked == true & RB_Horizon.Checked == false & RB_Verti.Checked == false)
            {
                MessageBox.Show("Select a Flip Option.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                worksheet.Activate();
                src_rng.Select();
                // Me.Close()
                return;
            }
            // End If

            else if (RB_Dropdown_35_Labels.Checked == true & src_rng.Columns.Count > 5)
            {
                MessageBox.Show("You can maximum select 5 columns.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                worksheet.Activate();
                src_rng.Select();
                // Me.Close()
                return;
            }

            else if (src_rng.Areas.Count > 1)
            {
                MessageBox.Show("Multiple selection is not possible in the Source Range field.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                TB_src_range.Focus();
                return;
            }


            else if (RB_Dropdown_2_Labels.Checked == true & src_rng.Rows.Count < 2)
            {
                MessageBox.Show("Select a valid Source Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                TB_src_range.Focus();
                return;
            }

            else if (RB_Dropdown_2_Labels.Checked == true & RB_Horizon.Checked == true & des_rng.Columns.Count != 2)
            {
                MessageBox.Show("Select a valid Destination Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                TB_dest_range.Focus();
                return;
            }

            else if (RB_Dropdown_2_Labels.Checked == true & RB_Verti.Checked == true & des_rng.Rows.Count != 2)
            {
                MessageBox.Show("Select a valid Destination Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                TB_dest_range.Focus();
                return;
            }

            else
            {
                try
                {
                    if (RB_Dropdown_35_Labels.Checked == true)
                    {
                        Range rng;
                        if (CB_header.Checked == true)
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

                        if (CB_ascending.Checked == true)
                        {
                            // Sort the list in ascending order
                            uniqueValues.Sort();
                        }
                        else if (CB_descending.Checked == true)
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

                        workSheet3.Change += worksheet1_Change;
                    }


                    else if (RB_Dropdown_2_Labels.Checked == true)
                    {
                        // Extract headers from A1:C1

                        src_rng = workSheet2.get_Range(src_rng.get_Address());

                        Range headersRange = (Range)src_rng.Rows[1];
                        var headers = new List<string>();
                        // Dim workbook As excelapp.workbook

                        foreach (Range cell in headersRange.Cells)
                            headers.Add(cell.get_Value().ToString());

                        Range dropDownRange = (Range)des_rng[1, 1];
                        var validation = dropDownRange.Validation;
                        validation.Delete(); // Remove any existing validation
                        validation.Add(XlDVType.xlValidateList, Formula1: string.Join(",", headers));

                        // Add event handler to listen for changes in E1

                        // AddHandler worksheet.Change, AddressOf worksheet1_Change
                        workSheet3.Change += worksheet1_Change;
                    }

                    if (CB_text.Checked == true)
                    {
                        des_rng.NumberFormat = "@";
                    }

                    if ((des_rng.Worksheet.Name ?? "") != (src_rng.Worksheet.Name ?? ""))
                    {
                        GlobalModule.Variable1 = src_rng.Worksheet.Name + "!" + TB_src_range.Text;
                        GlobalModule.Variable2 = TB_dest_range.Text;
                    }
                    else
                    {
                        GlobalModule.Variable1 = src_rng.Worksheet.Name + "!" + TB_src_range.Text;
                        GlobalModule.Variable2 = des_rng.Worksheet.Name + "!" + TB_dest_range.Text;
                    }

                    GlobalModule.Header = CB_header.Checked;
                    GlobalModule.Ascending = CB_ascending.Checked;
                    GlobalModule.Descending = CB_descending.Checked;
                    GlobalModule.TextConvert = CB_text.Checked;

                    Excel.Worksheet targetWorksheet = null;
                    foreach (Excel.Worksheet ws in excelApp.Worksheets)
                    {
                        if (ws.Name == "MySpecialSheet")
                        {
                            targetWorksheet = ws;
                            break;
                        }
                    }

                    // If "MySpecialSheet" does not exist, add it
                    if (targetWorksheet is null)
                    {
                        targetWorksheet = (Excel.Worksheet)excelApp.Worksheets.Add(After: excelApp.Worksheets[excelApp.Worksheets.Count]);
                        targetWorksheet.Name = "MySpecialSheet";
                    }

                    if (RB_Dropdown_2_Labels.Checked == true)
                    {
                        GlobalModule.OptionType = false;     // 2 label=False
                    }
                    else
                    {
                        GlobalModule.OptionType = true;

                    }      // 3-5 label=true

                    if (RB_Horizon.Checked == true & CustomGroupBox5.Enabled == true)
                    {
                        GlobalModule.Horizontal_CreateDP = true;
                    }
                    else if (RB_Verti.Checked == true & CustomGroupBox5.Enabled == true)
                    {
                        GlobalModule.Horizontal_CreateDP = false;
                    }

                    GlobalModule.Flag_CreateDDDL = true;
                    // sheetName = worksheet.Name
                    GlobalModule.sheetName10 = workSheet2.Name;
                    GlobalModule.sheetName11 = workSheet3.Name;
                    string sheetName1 = src_rng.Worksheet.Name;
                    string sheetName2 = des_rng.Worksheet.Name;

                    if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(targetWorksheet.get_Range("A1").get_Value(), "", false)))
                    {
                        // Write something in cell A1 of the target worksheet
                        targetWorksheet.get_Range("A1").set_Value(value: GlobalModule.Variable1);
                        targetWorksheet.get_Range("A2").set_Value(value: GlobalModule.Variable2);
                        targetWorksheet.get_Range("A3").set_Value(value: GlobalModule.Header);
                        targetWorksheet.get_Range("A4").set_Value(value: GlobalModule.Ascending);
                        targetWorksheet.get_Range("A5").set_Value(value: GlobalModule.Descending);
                        targetWorksheet.get_Range("A6").set_Value(value: GlobalModule.TextConvert);
                        targetWorksheet.get_Range("A7").set_Value(value: GlobalModule.OptionType);
                        targetWorksheet.get_Range("A8").set_Value(value: GlobalModule.Horizontal_CreateDP);
                        targetWorksheet.get_Range("A9").set_Value(value: GlobalModule.Flag_CreateDDDL);
                        targetWorksheet.get_Range("A10").set_Value(value: sheetName1);
                        targetWorksheet.get_Range("A11").set_Value(value: sheetName2);
                    }

                    else if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(targetWorksheet.get_Range("B1").get_Value(), "", false)))
                    {

                        targetWorksheet.get_Range("B1").set_Value(value: GlobalModule.Variable1);
                        targetWorksheet.get_Range("B2").set_Value(value: GlobalModule.Variable2);
                        targetWorksheet.get_Range("B3").set_Value(value: GlobalModule.Header);
                        targetWorksheet.get_Range("B4").set_Value(value: GlobalModule.Ascending);
                        targetWorksheet.get_Range("B5").set_Value(value: GlobalModule.Descending);
                        targetWorksheet.get_Range("B6").set_Value(value: GlobalModule.TextConvert);
                        targetWorksheet.get_Range("B7").set_Value(value: GlobalModule.OptionType);
                        targetWorksheet.get_Range("B8").set_Value(value: GlobalModule.Horizontal_CreateDP);
                        targetWorksheet.get_Range("B9").set_Value(value: GlobalModule.Flag_CreateDDDL);
                        targetWorksheet.get_Range("B10").set_Value(value: sheetName1);
                        targetWorksheet.get_Range("B11").set_Value(value: sheetName2);
                    }

                    else if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(targetWorksheet.get_Range("C1").get_Value(), "", false)))
                    {

                        targetWorksheet.get_Range("C1").set_Value(value: GlobalModule.Variable1);
                        targetWorksheet.get_Range("C2").set_Value(value: GlobalModule.Variable2);
                        targetWorksheet.get_Range("C3").set_Value(value: GlobalModule.Header);
                        targetWorksheet.get_Range("C4").set_Value(value: GlobalModule.Ascending);
                        targetWorksheet.get_Range("C5").set_Value(value: GlobalModule.Descending);
                        targetWorksheet.get_Range("C6").set_Value(value: GlobalModule.TextConvert);
                        targetWorksheet.get_Range("C7").set_Value(value: GlobalModule.OptionType);
                        targetWorksheet.get_Range("C8").set_Value(value: GlobalModule.Horizontal_CreateDP);
                        targetWorksheet.get_Range("C9").set_Value(value: GlobalModule.Flag_CreateDDDL);
                        targetWorksheet.get_Range("C10").set_Value(value: sheetName1);
                        targetWorksheet.get_Range("C11").set_Value(value: sheetName2);
                    }

                    else if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(targetWorksheet.get_Range("D1").get_Value(), "", false)))
                    {

                        targetWorksheet.get_Range("D1").set_Value(value: GlobalModule.Variable1);
                        targetWorksheet.get_Range("D2").set_Value(value: GlobalModule.Variable2);
                        targetWorksheet.get_Range("D3").set_Value(value: GlobalModule.Header);
                        targetWorksheet.get_Range("D4").set_Value(value: GlobalModule.Ascending);
                        targetWorksheet.get_Range("D5").set_Value(value: GlobalModule.Descending);
                        targetWorksheet.get_Range("D6").set_Value(value: GlobalModule.TextConvert);
                        targetWorksheet.get_Range("D7").set_Value(value: GlobalModule.OptionType);
                        targetWorksheet.get_Range("D8").set_Value(value: GlobalModule.Horizontal_CreateDP);
                        targetWorksheet.get_Range("D9").set_Value(value: GlobalModule.Flag_CreateDDDL);
                        targetWorksheet.get_Range("D10").set_Value(value: sheetName1);
                        targetWorksheet.get_Range("D11").set_Value(value: sheetName2);
                    }

                    else if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(targetWorksheet.get_Range("E1").get_Value(), "", false)))
                    {

                        targetWorksheet.get_Range("E1").set_Value(value: GlobalModule.Variable1);
                        targetWorksheet.get_Range("E2").set_Value(value: GlobalModule.Variable2);
                        targetWorksheet.get_Range("E3").set_Value(value: GlobalModule.Header);
                        targetWorksheet.get_Range("E4").set_Value(value: GlobalModule.Ascending);
                        targetWorksheet.get_Range("E5").set_Value(value: GlobalModule.Descending);
                        targetWorksheet.get_Range("E6").set_Value(value: GlobalModule.TextConvert);
                        targetWorksheet.get_Range("E7").set_Value(value: GlobalModule.OptionType);
                        targetWorksheet.get_Range("E8").set_Value(value: GlobalModule.Horizontal_CreateDP);
                        targetWorksheet.get_Range("E9").set_Value(value: GlobalModule.Flag_CreateDDDL);
                        targetWorksheet.get_Range("E10").set_Value(value: sheetName1);
                        targetWorksheet.get_Range("E11").set_Value(value: sheetName2);
                    }
                    else
                    {
                        // Cut range D1:D10
                        targetWorksheet.get_Range("B1:E11").Copy();

                        // Paste to range E1
                        targetWorksheet.get_Range("A1:D11").PasteSpecial(XlPasteType.xlPasteAll);
                        excelApp.CutCopyMode = (XlCutCopyMode)Conversions.ToInteger(false);

                        targetWorksheet.get_Range("E1:E11").set_Value(value: "");
                        targetWorksheet.get_Range("E1").set_Value(value: GlobalModule.Variable1);
                        targetWorksheet.get_Range("E2").set_Value(value: GlobalModule.Variable2);
                        targetWorksheet.get_Range("E3").set_Value(value: GlobalModule.Header);
                        targetWorksheet.get_Range("E4").set_Value(value: GlobalModule.Ascending);
                        targetWorksheet.get_Range("E5").set_Value(value: GlobalModule.Descending);
                        targetWorksheet.get_Range("E6").set_Value(value: GlobalModule.TextConvert);
                        targetWorksheet.get_Range("E7").set_Value(value: GlobalModule.OptionType);
                        targetWorksheet.get_Range("E8").set_Value(value: GlobalModule.Horizontal_CreateDP);
                        targetWorksheet.get_Range("E9").set_Value(value: GlobalModule.Flag_CreateDDDL);
                        targetWorksheet.get_Range("E10").set_Value(value: sheetName1);
                        targetWorksheet.get_Range("E11").set_Value(value: sheetName2);

                    }
                    // Hide the target worksheet
                    targetWorksheet.Visible = XlSheetVisibility.xlSheetHidden;


                    des_rng.set_Value(value: null);
                    des_rng.Select();

                    Dispose();
                }
                catch (Exception ex)
                {
                    Dispose();
                }
            }

        }

        private void Selection_destination_Click(object sender, EventArgs e)
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

                TB_dest_range.Text = des_rng.get_Address();

                Show();
                TB_dest_range.Focus();

            }
        }


        private void Form1_Load(object sender, EventArgs e)
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
                    TB_src_range.Text = selectedRange.get_Address();
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
                    if (ReferenceEquals(ActiveControl, TB_dest_range))
                    {
                        des_rng = selectionRange1;
                        // This will run on the Excel thread, so you need to use Invoke to update the UI
                        // Me.BeginInvoke(New System.Action(Sub() TB_dest_range.Text = selectionRange1.Address))
                        Activate();
                        BeginInvoke(new System.Action(() =>
                            {
                                TB_dest_range.Text = des_rng.get_Address();
                                SetWindowPos(Handle, new IntPtr(HWND_TOPMOST), 0, 0, 0, 0, SWP_NOACTIVATE | SWP_NOMOVE | SWP_NOSIZE);
                            }));
                    }

                    else if (ReferenceEquals(ActiveControl, TB_src_range))
                    {
                        src_rng = selectionRange1;
                        Activate();


                        BeginInvoke(new System.Action(() =>
                            {
                                TB_src_range.Text = src_rng.get_Address();
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



        public void worksheet1_Change(Range Target)
        {

            try
            {
                excelApp = Globals.ThisAddIn.Application;
                workBook = excelApp.ActiveWorkbook;
                workSheet = (Excel.Worksheet)workBook.ActiveSheet;


                Excel.Worksheet targetWorksheet = null;
                int i = 1;
                bool j = false;
                foreach (var ws in excelApp.ActiveWorkbook.Worksheets)
                {
                    if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(ws.name, "MySpecialSheet", false)))
                    {
                        targetWorksheet = (Excel.Worksheet)ws;
                        j = true;
                        break;

                    }
                }

                if (j == true)
                {
                    Range r11 = null;
                    Range r12 = null;
                    Range r13 = null;
                    Range r14 = null;
                    Range r15 = null;
                    // MsgBox(1)
                    if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(targetWorksheet.get_Range("A1").get_Value(), "", false)))
                    {
                        r11 = excelApp.get_Range(targetWorksheet.get_Range("A2").get_Value());
                        r11 = workSheet.get_Range(r11.get_Address());
                        // MsgBox(r11.Worksheet.Name)
                    }

                    if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(targetWorksheet.get_Range("B1").get_Value(), "", false)))
                    {
                        r12 = excelApp.get_Range(targetWorksheet.get_Range("B2").get_Value());
                        r12 = workSheet.get_Range(r12.get_Address());
                    }

                    if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(targetWorksheet.get_Range("C1").get_Value(), "", false)))
                    {
                        r13 = excelApp.get_Range(targetWorksheet.get_Range("C2").get_Value());
                        r13 = workSheet.get_Range(r13.get_Address());
                    }

                    if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(targetWorksheet.get_Range("D1").get_Value(), "", false)))
                    {
                        r14 = excelApp.get_Range(targetWorksheet.get_Range("D2").get_Value());
                        r14 = workSheet.get_Range(r14.get_Address());
                    }

                    if (Conversions.ToBoolean(Operators.ConditionalCompareObjectNotEqual(targetWorksheet.get_Range("E1").get_Value(), "", false)))
                    {
                        r15 = excelApp.get_Range(targetWorksheet.get_Range("E2").get_Value());
                        r15 = workSheet.get_Range(r15.get_Address());
                        // MsgBox(r15.Address)
                        // MsgBox(excelApp.Intersect(Target, r15).Address)
                        // MsgBox(Target.Worksheet.Name)
                        // MsgBox(targetWorksheet.Range("E11").Value)
                    }

                    // MsgBox(2)
                    if (Conversions.ToBoolean(Operators.AndObject(Operators.ConditionalCompareObjectEqual(Target.Worksheet.Name, targetWorksheet.get_Range("A11").get_Value(), false), excelApp.Intersect(Target, r11) is not null)))
                    {
                        // And (excelApp.Intersect(Target, r11) IsNot Nothing)) Then

                        // If excelApp.Intersect(Target, r11) IsNot Nothing Then

                        GlobalModule.Variable1 = targetWorksheet.get_Range("A1").get_Value().ToString();
                        GlobalModule.Variable2 = targetWorksheet.get_Range("A2").get_Value().ToString();
                        GlobalModule.Header = Conversions.ToBoolean(targetWorksheet.get_Range("A3").get_Value().ToString());
                        GlobalModule.Ascending = Conversions.ToBoolean(targetWorksheet.get_Range("A4").get_Value().ToString());
                        GlobalModule.Descending = Conversions.ToBoolean(targetWorksheet.get_Range("A5").get_Value().ToString());
                        GlobalModule.TextConvert = Conversions.ToBoolean(targetWorksheet.get_Range("A6").get_Value().ToString());
                        GlobalModule.OptionType = Conversions.ToBoolean(targetWorksheet.get_Range("A7").get_Value().ToString());
                        GlobalModule.Horizontal_CreateDP = Conversions.ToBoolean(targetWorksheet.get_Range("A8").get_Value().ToString());
                        GlobalModule.Flag_CreateDDDL = Conversions.ToBoolean(targetWorksheet.get_Range("A9").get_Value().ToString());
                        GlobalModule.sheetName10 = targetWorksheet.get_Range("A10").get_Value().ToString();
                        GlobalModule.sheetName11 = targetWorksheet.get_Range("A11").get_Value().ToString();
                    }
                    // End If
                    else if (Conversions.ToBoolean(Operators.AndObject(Operators.ConditionalCompareObjectEqual(Target.Worksheet.Name, targetWorksheet.get_Range("B11").get_Value(), false), excelApp.Intersect(Target, r12) is not null)))
                    {
                        // AndAlso (excelApp.Intersect(Target, r12) IsNot Nothing) Then

                        // If excelApp.Intersect(Target, r12) IsNot Nothing Then
                        GlobalModule.Variable1 = targetWorksheet.get_Range("B1").get_Value().ToString();
                        GlobalModule.Variable2 = targetWorksheet.get_Range("B2").get_Value().ToString();
                        GlobalModule.Header = Conversions.ToBoolean(targetWorksheet.get_Range("B3").get_Value().ToString());
                        GlobalModule.Ascending = Conversions.ToBoolean(targetWorksheet.get_Range("B4").get_Value().ToString());
                        GlobalModule.Descending = Conversions.ToBoolean(targetWorksheet.get_Range("B5").get_Value().ToString());
                        GlobalModule.TextConvert = Conversions.ToBoolean(targetWorksheet.get_Range("B6").get_Value().ToString());
                        GlobalModule.OptionType = Conversions.ToBoolean(targetWorksheet.get_Range("B7").get_Value().ToString());
                        GlobalModule.Horizontal_CreateDP = Conversions.ToBoolean(targetWorksheet.get_Range("B8").get_Value().ToString());
                        GlobalModule.Flag_CreateDDDL = Conversions.ToBoolean(targetWorksheet.get_Range("B9").get_Value().ToString());
                        GlobalModule.sheetName10 = targetWorksheet.get_Range("B10").get_Value().ToString();
                        GlobalModule.sheetName11 = targetWorksheet.get_Range("B11").get_Value().ToString();
                    }
                    // End If

                    else if (Conversions.ToBoolean(Operators.AndObject(Operators.AndObject(Operators.ConditionalCompareObjectEqual(Target.Worksheet.Name, targetWorksheet.get_Range("C11").get_Value(), false), r13 is not null), excelApp.Intersect(Target, r13) is not null)))
                    {
                        // If excelApp.Intersect(Target, r13) IsNot Nothing Then

                        GlobalModule.Variable1 = targetWorksheet.get_Range("C1").get_Value().ToString();
                        GlobalModule.Variable2 = targetWorksheet.get_Range("C2").get_Value().ToString();
                        GlobalModule.Header = Conversions.ToBoolean(targetWorksheet.get_Range("C3").get_Value().ToString());
                        GlobalModule.Ascending = Conversions.ToBoolean(targetWorksheet.get_Range("C4").get_Value().ToString());
                        GlobalModule.Descending = Conversions.ToBoolean(targetWorksheet.get_Range("C5").get_Value().ToString());
                        GlobalModule.TextConvert = Conversions.ToBoolean(targetWorksheet.get_Range("C6").get_Value().ToString());
                        GlobalModule.OptionType = Conversions.ToBoolean(targetWorksheet.get_Range("C7").get_Value().ToString());
                        GlobalModule.Horizontal_CreateDP = Conversions.ToBoolean(targetWorksheet.get_Range("C8").get_Value().ToString());
                        GlobalModule.Flag_CreateDDDL = Conversions.ToBoolean(targetWorksheet.get_Range("C9").get_Value().ToString());
                        GlobalModule.sheetName10 = targetWorksheet.get_Range("C10").get_Value().ToString();
                        GlobalModule.sheetName11 = targetWorksheet.get_Range("C11").get_Value().ToString();
                    }
                    // End If
                    else if (Conversions.ToBoolean(Operators.AndObject(Operators.AndObject(Operators.ConditionalCompareObjectEqual(Target.Worksheet.Name, targetWorksheet.get_Range("D11").get_Value(), false), r14 is not null), excelApp.Intersect(Target, r14) is not null)))
                    {
                        // If excelApp.Intersect(Target, r14) IsNot Nothing Then

                        GlobalModule.Variable1 = targetWorksheet.get_Range("D1").get_Value().ToString();
                        GlobalModule.Variable2 = targetWorksheet.get_Range("D2").get_Value().ToString();
                        GlobalModule.Header = Conversions.ToBoolean(targetWorksheet.get_Range("D3").get_Value().ToString());
                        GlobalModule.Ascending = Conversions.ToBoolean(targetWorksheet.get_Range("D4").get_Value().ToString());
                        GlobalModule.Descending = Conversions.ToBoolean(targetWorksheet.get_Range("D5").get_Value().ToString());
                        GlobalModule.TextConvert = Conversions.ToBoolean(targetWorksheet.get_Range("D6").get_Value().ToString());
                        GlobalModule.OptionType = Conversions.ToBoolean(targetWorksheet.get_Range("D7").get_Value().ToString());
                        GlobalModule.Horizontal_CreateDP = Conversions.ToBoolean(targetWorksheet.get_Range("D8").get_Value().ToString());
                        GlobalModule.Flag_CreateDDDL = Conversions.ToBoolean(targetWorksheet.get_Range("D9").get_Value().ToString());
                        GlobalModule.sheetName10 = targetWorksheet.get_Range("D10").get_Value().ToString();
                        GlobalModule.sheetName11 = targetWorksheet.get_Range("D11").get_Value().ToString();
                    }
                    // End If
                    else if (Conversions.ToBoolean(Operators.AndObject(Operators.AndObject(Operators.ConditionalCompareObjectEqual(Target.Worksheet.Name, targetWorksheet.get_Range("E11").get_Value(), false), r15 is not null), excelApp.Intersect(Target, r15) is not null)))
                    {
                        // If excelApp.Intersect(Target, r15) IsNot Nothing Then

                        GlobalModule.Variable1 = targetWorksheet.get_Range("E1").get_Value().ToString();
                        GlobalModule.Variable2 = targetWorksheet.get_Range("E2").get_Value().ToString();
                        GlobalModule.Header = Conversions.ToBoolean(targetWorksheet.get_Range("E3").get_Value().ToString());
                        GlobalModule.Ascending = Conversions.ToBoolean(targetWorksheet.get_Range("E4").get_Value().ToString());
                        GlobalModule.Descending = Conversions.ToBoolean(targetWorksheet.get_Range("E5").get_Value().ToString());
                        GlobalModule.TextConvert = Conversions.ToBoolean(targetWorksheet.get_Range("E6").get_Value().ToString());
                        GlobalModule.OptionType = Conversions.ToBoolean(targetWorksheet.get_Range("E7").get_Value().ToString());
                        GlobalModule.Horizontal_CreateDP = Conversions.ToBoolean(targetWorksheet.get_Range("E8").get_Value().ToString());
                        GlobalModule.Flag_CreateDDDL = Conversions.ToBoolean(targetWorksheet.get_Range("E9").get_Value().ToString());
                        GlobalModule.sheetName10 = targetWorksheet.get_Range("E10").get_Value().ToString();
                        GlobalModule.sheetName11 = targetWorksheet.get_Range("E11").get_Value().ToString();
                        // End If
                    }
                    // MsgBox(Variable1)
                    // MsgBox(Variable2)
                    src_rng = excelApp.get_Range(GlobalModule.Variable1);
                    Excel.Worksheet src_ws = (Excel.Worksheet)workBook.Worksheets[GlobalModule.sheetName10];
                    Excel.Worksheet des_ws = (Excel.Worksheet)workBook.Worksheets[GlobalModule.sheetName11];
                    src_rng = src_ws.get_Range(GlobalModule.Variable1);

                    des_rng = des_ws.get_Range(des_rng.get_Address());

                    if (excelApp.Intersect(Target, des_rng) is not null)
                    {
                        Range rng;


                        if (RB_Dropdown_35_Labels.Checked == true)
                        {
                            if (CB_header.Checked == true)
                            {
                                // Dim adjustRange As Excel.Range
                                rng = src_rng.get_Offset(1, 0).get_Resize(src_rng.Rows.Count - 1, src_rng.Columns.Count);
                            }

                            else
                            {

                                rng = src_rng;

                            }

                            int col_dif;
                            col_dif = Target.Column - des_rng.Column + 1;

                            // For k = 1 To des_rng.Rows.Count
                            var matchedValues = new List<string>();
                            var sec_matchedValues = new List<string>();
                            var thrd_matchedValues = new List<string>();
                            var four_matchedValues = new List<string>();
                            int k = Target.Row - des_rng.Row + 1;

                            if (col_dif == 1)
                            {

                                if (des_rng[k, 1].Value is not null)
                                {
                                    var loopTo = rng.Rows.Count;
                                    for (i = 1; i <= loopTo; i++)
                                    {
                                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(rng[i, 1].Value, des_rng[k, 1].Value, false)))
                                        {
                                            if (!matchedValues.Contains(Conversions.ToString(rng[i, 2].Value)))
                                            {
                                                matchedValues.Add(Conversions.ToString(rng[i, 2].Value));
                                            }
                                        }
                                    }


                                    if (CB_ascending.Checked == true)
                                    {
                                        // Sort the list in ascending order
                                        matchedValues.Sort();
                                    }
                                    else if (CB_descending.Checked == true)
                                    {
                                        // Sort the list in ascending order
                                        matchedValues.Sort();
                                        matchedValues.Reverse();
                                    }

                                    // Dim dropDownRange As Excel.Range = des_rng(k, 2)
                                    Range dropDownRange = (Range)Target[1, 2];

                                    var Validation = dropDownRange.Validation;
                                    Validation.Delete(); // Remove any existing validation
                                    Validation.Add(XlDVType.xlValidateList, Formula1: string.Join(",", matchedValues));
                                    matchedValues.Clear();

                                }
                            }

                            // Dim sec_matchedValues As New List(Of String)
                            else if (col_dif == 2)
                            {
                                if (des_rng[k, 2].Value is not null)
                                {
                                    var loopTo1 = rng.Rows.Count;
                                    for (i = 1; i <= loopTo1; i++)
                                    {
                                        if (Conversions.ToBoolean(Operators.AndObject(Operators.ConditionalCompareObjectEqual(rng[i, 1].Value, des_rng[k, 1].Value, false), Operators.ConditionalCompareObjectEqual(rng[i, 2].Value, des_rng[k, 2].Value, false))))
                                        {
                                            if (!sec_matchedValues.Contains(Conversions.ToString(rng[i, 3].Value)))
                                            {
                                                sec_matchedValues.Add(Conversions.ToString(rng[i, 3].Value));
                                            }

                                        }
                                    }


                                    if (CB_ascending.Checked == true)
                                    {
                                        // Sort the list in ascending order
                                        sec_matchedValues.Sort();
                                    }
                                    else if (CB_descending.Checked == true)
                                    {
                                        // Sort the list in ascending order
                                        sec_matchedValues.Sort();
                                        sec_matchedValues.Reverse();
                                    }


                                    // Dim dropDownRange As Excel.Range = des_rng(k, 3)
                                    Range dropDownRange = default;
                                    var Validation = dropDownRange.Validation;
                                    Validation.Delete(); // Remove any existing validation
                                    Validation.Add(XlDVType.xlValidateList, Formula1: string.Join(",", sec_matchedValues));
                                    sec_matchedValues.Clear();
                                }
                            }
                            else if (col_dif == 3)
                            {
                                // Dim thrd_matchedValues As New List(Of String)

                                if (des_rng[k, 3].Value is not null)
                                {
                                    var loopTo2 = rng.Rows.Count;
                                    for (i = 1; i <= loopTo2; i++)
                                    {
                                        if (Conversions.ToBoolean(Operators.AndObject(Operators.AndObject(Operators.ConditionalCompareObjectEqual(rng[i, 1].Value, des_rng[k, 1].Value, false), Operators.ConditionalCompareObjectEqual(rng[i, 2].Value, des_rng[k, 2].Value, false)), Operators.ConditionalCompareObjectEqual(rng[i, 3].Value, des_rng[k, 3].Value, false))))
                                        {
                                            if (!thrd_matchedValues.Contains(Conversions.ToString(rng[i, 4].Value)))
                                            {
                                                thrd_matchedValues.Add(Conversions.ToString(rng[i, 4].Value));
                                            }

                                        }
                                    }


                                    if (CB_ascending.Checked == true)
                                    {
                                        // Sort the list in ascending order
                                        thrd_matchedValues.Sort();
                                    }
                                    else if (CB_descending.Checked == true)
                                    {
                                        // Sort the list in ascending order
                                        thrd_matchedValues.Sort();
                                        thrd_matchedValues.Reverse();
                                    }


                                    // Dim dropDownRange As Excel.Range = des_rng(k, 4)
                                    Range dropDownRange = default;
                                    var Validation = dropDownRange.Validation;
                                    Validation.Delete(); // Remove any existing validation
                                    Validation.Add(XlDVType.xlValidateList, Formula1: string.Join(",", thrd_matchedValues));
                                    thrd_matchedValues.Clear();
                                }
                            }


                            // Dim four_matchedValues As New List(Of String)
                            else if (col_dif == 4)
                            {
                                if (des_rng[k, 4].Value is not null)
                                {
                                    var loopTo3 = rng.Rows.Count;
                                    for (i = 1; i <= loopTo3; i++)
                                    {
                                        if (Conversions.ToBoolean(Operators.AndObject(Operators.AndObject(Operators.AndObject(Operators.ConditionalCompareObjectEqual(rng[i, 1].Value, des_rng[k, 1].Value, false), Operators.ConditionalCompareObjectEqual(rng[i, 2].Value, des_rng[k, 2].Value, false)), Operators.ConditionalCompareObjectEqual(rng[i, 3].Value, des_rng[k, 3].Value, false)), Operators.ConditionalCompareObjectEqual(rng[i, 4].Value, des_rng[k, 4].Value, false))))
                                        {

                                            if (!four_matchedValues.Contains(Conversions.ToString(rng[i, 5].Value)))
                                            {
                                                four_matchedValues.Add(Conversions.ToString(rng[i, 5].Value));
                                            }


                                        }
                                    }


                                    if (CB_ascending.Checked == true)
                                    {
                                        // Sort the list in ascending order
                                        four_matchedValues.Sort();
                                    }
                                    else if (CB_descending.Checked == true)
                                    {
                                        // Sort the list in ascending order
                                        four_matchedValues.Sort();
                                        four_matchedValues.Reverse();
                                    }


                                    Range dropDownRange = (Range)des_rng[k, 5];
                                    var Validation = dropDownRange.Validation;
                                    Validation.Delete(); // Remove any existing validation
                                    Validation.Add(XlDVType.xlValidateList, Formula1: string.Join(",", four_matchedValues));
                                    four_matchedValues.Clear();
                                }
                            }
                        }

                        // Next

                        else if (RB_Dropdown_2_Labels.Checked == true)
                        {
                            if (RB_Horizon.Checked == true)
                            {
                                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(Target.get_Address(), des_rng[1, 1].Address, false)))
                                {

                                    // Dim worksheet As Excel.Worksheet = CType(Target.Worksheet, Excel.Worksheet)
                                    int col = src_rng.Rows.Find(Target.get_Value()).Column - src_rng.Column + 1;

                                    Range sourceRng = (Range)src_rng.Cells[2, col].Resize(Operators.SubtractObject(src_rng[src_rng.Rows.Count, col].row, 2), (object)1);

                                    Range dropDownRange = (Range)des_rng[1, 2];
                                    var Validation = dropDownRange.Validation;
                                    Validation.Delete(); // Remove any existing validation
                                    string formula = "='" + GlobalModule.sheetName10 + "'!" + sourceRng.get_Address(External: false);

                                    Validation.Add(XlDVType.xlValidateList, Formula1: formula);

                                }
                            }

                            else if (RB_Verti.Checked == true)
                            {
                                if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(Target.get_Address(), des_rng[1, 1].Address, false)))
                                {

                                    int col = src_rng.Rows.Find(Target.get_Value()).Column - src_rng.Column + 1;

                                    Range sourceRng = (Range)src_rng.Cells[2, col].Resize(Operators.SubtractObject(src_rng[src_rng.Rows.Count, col].row, 2), (object)1);

                                    Range dropDownRange = (Range)des_rng[2, 1];
                                    var Validation = dropDownRange.Validation;
                                    Validation.Delete(); // Remove any existing validation

                                    string formula = "='" + sourceRng.Worksheet.Name + "'!" + sourceRng.get_Address(External: false);
                                    Validation.Add(XlDVType.xlValidateList, Formula1: formula);
                                }
                            }

                        }

                    }

                }
            }
            catch (Exception ex)
            {

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


        private void form(object sender, KeyEventArgs e)
        {

            try
            {
                if (e.KeyCode == Keys.Enter)
                {
                    Btn_OK.Focus();

                    Btn_OK_Click(sender, e);

                }
            }

            catch (Exception ex)
            {

            }

        }

        // Private Sub CB_asceding(sender As Object, e As KeyEventArgs) Handles CB_ascending.KeyDown

        // Try
        // If e.KeyCode = Keys.Enter Then

        // Call Btn_OK_Click(sender, e)

        // End If

        // Catch ex As Exception

        // End Try

        // End Sub

        // Private Sub CB_desceding(sender As Object, e As KeyEventArgs) Handles CB_descending.KeyDown

        // Try
        // If e.KeyCode = Keys.Enter Then

        // Call Btn_OK_Click(sender, e)

        // End If

        // Catch ex As Exception

        // End Try

        // End Sub

        // Private Sub CB_head(sender As Object, e As KeyEventArgs) Handles CB_header.KeyDown

        // Try
        // If e.KeyCode = Keys.Enter Then

        // Call Btn_OK_Click(sender, e)

        // End If

        // Catch ex As Exception

        // End Try

        // End Sub

        // Private Sub CB_texting(sender As Object, e As KeyEventArgs) Handles CB_text.KeyDown

        // Try
        // If e.KeyCode = Keys.Enter Then

        // Call Btn_OK_Click(sender, e)

        // End If

        // Catch ex As Exception

        // End Try

        // End Sub

        // Private Sub RB_Label2(sender As Object, e As KeyEventArgs) Handles RB_Dropdown_2_Labels.KeyDown

        // Try
        // If e.KeyCode = Keys.Enter Then

        // Call Btn_OK_Click(sender, e)

        // End If

        // Catch ex As Exception

        // End Try

        // End Sub

        // Private Sub RB_35(sender As Object, e As KeyEventArgs) Handles RB_Dropdown_35_Labels.KeyDown

        // Try
        // If e.KeyCode = Keys.Enter Then

        // Call Btn_OK_Click(sender, e)

        // End If

        // Catch ex As Exception

        // End Try

        // End Sub

        // Private Sub RB_horiz(sender As Object, e As KeyEventArgs) Handles RB_Horizon.KeyDown

        // Try
        // If e.KeyCode = Keys.Enter Then

        // Call Btn_OK_Click(sender, e)

        // End If

        // Catch ex As Exception

        // End Try

        // End Sub

        // Private Sub RB_verticalll(sender As Object, e As KeyEventArgs) Handles RB_Verti.KeyDown

        // Try
        // If e.KeyCode = Keys.Enter Then

        // Call Btn_OK_Click(sender, e)

        // End If

        // Catch ex As Exception

        // End Try

        // End Sub

        // Private Sub TB_dest_range_Enter(sender As Object, e As KeyEventArgs) Handles TB_dest_range.KeyDown
        // 'If Enter key is pressed then check if the text is a valid address
        // If IsValidExcelCellReference(TB_dest_range.Text) = True And e.KeyCode = Keys.Enter Then
        // des_rng = excelApp.Range(TB_dest_range.Text)
        // TB_dest_range.Focus()
        // des_rng.Select()

        // Call Btn_OK_Click(sender, e)   'OK button click event called

        // 'MsgBox(des_rng.Address)
        // ElseIf IsValidExcelCellReference(TB_dest_range.Text) = False And e.KeyCode = Keys.Enter Then
        // MessageBox.Show("Select the valid Destination Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        // TB_dest_range.Text = ""
        // TB_dest_range.Focus()
        // 'Me.Close()
        // Exit Sub
        // End If
        // End Sub

        // Private Sub TB_src_range_Enter(sender As Object, e As KeyEventArgs) Handles TB_src_range.KeyDown
        // 'If Enter key is pressed then check if the text is a valid address

        // If IsValidExcelCellReference(TB_src_range.Text) = True And e.KeyCode = Keys.Enter Then
        // src_rng = excelApp.Range(TB_src_range.Text)
        // TB_src_range.Focus()
        // src_rng.Select()

        // Call Btn_OK_Click(sender, e)   'OK button click event called

        // 'MsgBox(des_rng.Address)
        // ElseIf IsValidExcelCellReference(TB_src_range.Text) = False And e.KeyCode = Keys.Enter Then
        // MessageBox.Show("Select the valid Source Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        // TB_src_range.Text = ""
        // TB_src_range.Focus()
        // 'Me.Close()
        // Exit Sub
        // End If
        // End Sub

        private void TB_dest_range_TextChanged(object sender, EventArgs e)
        {
            excelApp = Globals.ThisAddIn.Application;
            var workbook = excelApp.ActiveWorkbook;
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.ActiveSheet;
            try
            {
                if (TB_dest_range.Text is not null & IsValidExcelCellReference(TB_dest_range.Text) == true)
                {
                    focuschange = true;
                    string sheetname = "";

                    try
                    {

                        des_rng = worksheet.get_Range(TB_dest_range.Text);
                        des_rng.Select();
                    }
                    catch
                    {
                        // Split the string into sheet name and cell address
                        string[] parts = TB_dest_range.Text.Split('!');
                        sheetname = parts[0];
                        string cellAddress = parts[1];
                        worksheet = (Excel.Worksheet)workbook.Worksheets[sheetname];
                        worksheet.Activate();
                        des_rng = worksheet.get_Range(cellAddress);
                        des_rng.Select();
                    }

                    if ((workSheet2.Name ?? "") != (worksheet.Name ?? "") & TB_dest_range.Text.Contains("!") == false)
                    {

                        TB_dest_range.Text = worksheet.Name + "!" + TB_dest_range.Text;

                    }

                    Activate();
                    TB_dest_range.Focus();
                    TB_dest_range.SelectionStart = TB_dest_range.Text.Length;

                    focuschange = false;
                    ax = worksheet.Name;
                    workSheet3 = worksheet;
                }
            }
            catch (Exception ex)
            {
                focuschange = false;
            }
        }

        private void TB_src_range_TextChanged(object sender, EventArgs e)
        {
            excelApp = Globals.ThisAddIn.Application;
            var workbook = excelApp.ActiveWorkbook;
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.ActiveSheet;
            try
            {
                if (TB_src_range.Text is not null & IsValidExcelCellReference(TB_src_range.Text) == true)
                {
                    focuschange = true;

                    // Define the range of cells to read (for example, cells A1 to A10)
                    src_rng = excelApp.get_Range(TB_src_range.Text);
                    src_rng = worksheet.get_Range(TB_src_range.Text);
                    src_rng.Select();
                    var range = src_rng;


                    Activate();
                    // TB_src_range.Focus()
                    TB_src_range.SelectionStart = TB_src_range.Text.Length;
                    focuschange = false;
                    workSheet2 = worksheet;

                }
            }
            catch (Exception ex)
            {
            }
        }

        // Private Sub Form1_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        // If e.KeyCode = Keys.Enter Then
        // Btn_OK.Focus()
        // Btn_OK.PerformClick()
        // End If
        // End Sub

        private void Form30_Create_Dynamic_Drop_down_List_Closing(object sender, CancelEventArgs e)
        {
            GlobalModule.form_flag = false;
        }

        private void Form30_Create_Dynamic_Drop_down_List_Disposed(object sender, EventArgs e)
        {
            GlobalModule.form_flag = false;
        }

        private void Form30_Create_Dynamic_Drop_down_List_Shown(object sender, EventArgs e)
        {
            Focus();
            BringToFront();
            Activate();
            BeginInvoke(new System.Action(() =>
                {
                    TB_src_range.Text = src_rng.get_Address();
                    SetWindowPos(Handle, new IntPtr(HWND_TOPMOST), 0, 0, 0, 0, SWP_NOACTIVATE | SWP_NOMOVE | SWP_NOSIZE);
                }));
        }
    }
}