using System;
using System.Collections.Generic;
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

    public partial class Form41_RemoveAdavancedDropdownList
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
        private List<WorksheetHandler> SheetHandlers = new List<WorksheetHandler>();
        private DocEvents_ChangeEventHandler EventDel_CellsChange;

        private Excel.Worksheet CurrentSheet;
        private Excel.Workbook WorkbookEvents;

        private string srcRng1;
        private string srcRng2;
        private string srcRng3;

        private Form36 form = null;
        public bool focuschange;

        private Range src_rng;
        private Form35Multi_SelectionbasedDropdown frm1 = null;
        private Range selectedRange;

        public Form41_RemoveAdavancedDropdownList()
        {
            InitializeComponent();
        }

        [DllImport("user32")]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
        private const uint SWP_NOMOVE = 0x2U;
        private const uint SWP_NOSIZE = 0x1U;
        private const uint SWP_NOACTIVATE = 0x10U;
        private const int HWND_TOPMOST = -1;

        private void CheckBox3_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void CB_multiselect_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void Btn_OK_Click(object sender, EventArgs e)
        {
            excelApp = Globals.ThisAddIn.Application;
            var workbook = excelApp.ActiveWorkbook;
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.ActiveSheet;

            if (CB_Source.Text == "Select Range" & string.IsNullOrEmpty(TB_src_rng.Text))
            {
                MessageBox.Show("Select a Source Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                TB_src_rng.Focus();
                // Me.Close()
                return;
            }
            else if (CB_multiselect.Checked == false & CB_checkbox.Checked == false & CB_search.Checked == false)
            {
                MessageBox.Show("Please, select the Data Validation List type.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                TB_src_rng.Focus();
                // Me.Close()
                return;
            }

            else if (src_rng.Areas.Count > 1)
            {
                MessageBox.Show("Multiple selection is not possible in the Source Range field.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                TB_src_rng.Focus();
            }

            else
            {

                if (CB_Source.Text != "Select Range")
                {
                    src_rng = worksheet.get_Range("A1", worksheet.Cells[excelApp.Rows.Count, excelApp.Columns.Count]);
                }

                if (CB_Source.Text.Contains("Active Workbook") & CB_multiselect.Checked == true)
                {
                    foreach (Excel.Worksheet sheet in workbook.Sheets)
                    {
                        if (sheet.Name == "Newwwwwwwwww")
                        {
                            sheet.Delete();
                            break;
                        }
                    }
                    GlobalModule.GB_CB_Source1 = "";
                    GlobalModule.SR1 = "";
                    GlobalModule.Horizontal1 = false;
                    GlobalModule.Separator1 = "";
                    GlobalModule.Search1 = false;
                    GlobalModule.Flag1 = false;
                    GlobalModule.TargetVar1 = "";
                    GlobalModule.RangeType1 = "";
                }

                else if (CB_Source.Text.Contains("Active Workbook") & CB_checkbox.Checked == true)
                {
                    foreach (Excel.Worksheet sheet in workbook.Sheets)
                    {
                        if (sheet.Name == "SofTekoSofTeko")
                        {
                            sheet.Delete();
                            break;
                        }
                    }
                    GlobalModule.GB_CB_Source2 = "";
                    GlobalModule.SR2 = "";
                    GlobalModule.Horizontal2 = false;
                    GlobalModule.Separator2 = "";
                    GlobalModule.Search2 = false;
                    GlobalModule.Flag2 = false;
                    GlobalModule.TargetVar2 = "";
                    GlobalModule.RangeType2 = "";
                }

                else if (CB_Source.Text.Contains("Active Workbook") & CB_search.Checked == true)
                {
                    foreach (Excel.Worksheet sheet in workbook.Sheets)
                    {
                        if (sheet.Name == "SofTekoSofTekoSofteko")
                        {
                            sheet.Delete();
                            break;
                        }
                    }
                    GlobalModule.GB_CB_Source3 = "";
                    GlobalModule.SR3 = "";
                    GlobalModule.Horizontal3 = false;
                    GlobalModule.Separator3 = "";
                    GlobalModule.Search3 = false;
                    GlobalModule.Flag3 = false;
                    GlobalModule.TargetVar3 = "";
                    GlobalModule.RangeType3 = "";
                }

                else if (CB_Source.Text == "Select Range" | CB_Source.Text.Contains("Active Sheet"))
                {

                    if (CB_Source.Text.Contains("Active Sheet"))
                    {
                        src_rng = worksheet.get_Range("A1", worksheet.Cells[excelApp.Rows.Count, excelApp.Columns.Count]);
                    }

                    if (CB_multiselect.Checked == true)
                    {
                        // RemoveHandler worksheet.SelectionChange, AddressOf sheet_SelectionChange
                        srcRng1 = GlobalModule.GB_CB_Source1;


                        if (IsCellInsideRange(src_rng, worksheet.get_Range(srcRng1)) == true)
                        {
                            // Dim result As Excel.Range = SubtractRanges(rng1, rng2)
                            // Dim results As List(Of Excel.Range) = SubtractRanges(worksheet.Range(srcRng1), src_rng)
                            // Dim addressList As New List(Of String)

                            // Dim combinedAddress As String = ""
                            // For Each r In results
                            // addressList.Add(r.Address)

                            // Next

                            // combinedAddress = String.Join(",", addressList)
                            // targetWorksheet.Name & "!" & targetRange.Address(External:=False)
                            // GB_CB_Dlt = combinedAddress
                            GlobalModule.GB_CB_Dlt1 = src_rng.get_Address();
                            GlobalModule.Nam1 = Conversions.ToString(workbook.ActiveSheet.name);
                            // MsgBox(GB_CB_Source1)
                        }



                    }

                    if (CB_checkbox.Checked == true)
                    {
                        srcRng2 = GlobalModule.GB_CB_Source2;


                        if (IsCellInsideRange(src_rng, worksheet.get_Range(srcRng2)) == true)
                        {
                            // Dim result As Excel.Range = SubtractRanges(rng1, rng2)
                            // Dim results As List(Of Excel.Range) = SubtractRanges(worksheet.Range(srcRng2), src_rng)
                            // Dim addressList As New List(Of String)

                            // Dim combinedAddress As String = ""
                            // For Each r In results
                            // addressList.Add(r.Address)

                            // Next

                            // combinedAddress = String.Join(",", addressList)
                            // targetWorksheet.Name & "!" & targetRange.Address(External:=False)
                            // GB_CB_Dlt = combinedAddress
                            GlobalModule.GB_CB_Dlt2 = src_rng.get_Address();
                            GlobalModule.Nam2 = Conversions.ToString(workbook.ActiveSheet.name);
                            // MsgBox(GB_CB_Source1)
                        }

                    }

                    if (CB_search.Checked == true)
                    {
                        srcRng3 = GlobalModule.GB_CB_Source3;


                        if (IsCellInsideRange(src_rng, worksheet.get_Range(srcRng3)) == true)
                        {
                            // Dim result As Excel.Range = SubtractRanges(rng1, rng2)
                            // Dim results As List(Of Excel.Range) = SubtractRanges(workSheet.Range(srcRng3), src_rng)
                            // Dim addressList As New List(Of String)

                            // Dim combinedAddress As String = ""
                            // For Each r In results
                            // addressList.Add(r.Address)

                            // Next

                            // combinedAddress = String.Join(",", addressList)
                            // GB_CB_Source3 = combinedAddress

                            GlobalModule.GB_CB_Dlt3 = src_rng.get_Address();
                            GlobalModule.Nam3 = Conversions.ToString(workbook.ActiveSheet.name);

                        }

                    }
                }

                else
                {
                    // MsgBox(1)
                    // Dim targetWorksheet As Excel.Worksheet = Nothing
                    // targetWorksheet = CType(workbook.Sheets(CB_Source.Text), Excel.Worksheet)
                    // 'MsgBox(2)
                    // 'src_rng = worksheet.Range(CB_Source.Text)
                    // 'src_rng = workbook.Sheet(CB_Source.Text).src_rng
                    // src_rng = targetWorksheet.Range(src_rng.Address) ' Change the range as needed


                    if (CB_multiselect.Checked == true)
                    {
                        // srcRng1 = GB_CB_Source1
                        // Dim srcRng1_prime As Excel.Range = workbook.ActiveSheet.Range(srcRng1)

                        // If IsCellInsideRange(src_rng, srcRng1_prime) = True Then
                        // 'Dim result As Excel.Range = SubtractRanges(rng1, rng2)
                        // Dim results As List(Of Excel.Range) = SubtractRanges(srcRng1_prime, src_rng)
                        // Dim addressList As New List(Of String)

                        // Dim combinedAddress As String = ""
                        // For Each r In results
                        // addressList.Add(r.Address)

                        // Next

                        // combinedAddress = String.Join(",", addressList)
                        // GB_CB_Source1 = combinedAddress
                        // ' MsgBox()

                        // End If


                        src_rng = worksheet.get_Range("A1", worksheet.Cells[excelApp.Rows.Count, excelApp.Columns.Count]);
                        GlobalModule.GB_CB_Dlt1 = src_rng.get_Address();
                        GlobalModule.Nam1 = CB_Source.Text;


                    }

                    if (CB_checkbox.Checked == true)
                    {
                        // srcRng2 = GB_CB_Source2
                        // Dim srcRng2_prime As Excel.Range = workbook.ActiveSheet.Range(srcRng2)


                        // If IsCellInsideRange(src_rng, srcRng2_prime) = True Then
                        // 'Dim result As Excel.Range = SubtractRanges(rng1, rng2)
                        // Dim results As List(Of Excel.Range) = SubtractRanges(srcRng2_prime, src_rng)
                        // Dim addressList As New List(Of String)

                        // Dim combinedAddress As String = ""
                        // For Each r In results
                        // addressList.Add(r.Address)

                        // Next

                        // combinedAddress = String.Join(",", addressList)
                        // GB_CB_Source2 = combinedAddress

                        // End If

                        src_rng = worksheet.get_Range("A1", worksheet.Cells[excelApp.Rows.Count, excelApp.Columns.Count]);
                        GlobalModule.GB_CB_Dlt2 = src_rng.get_Address();
                        GlobalModule.Nam2 = CB_Source.Text;

                    }

                    if (CB_search.Checked == true)
                    {
                        // srcRng3 = GB_CB_Source3
                        // Dim srcRng3_prime As Excel.Range = workbook.ActiveSheet.Range(srcRng3)


                        // If IsCellInsideRange(src_rng, srcRng3_prime) = True Then
                        // 'Dim result As Excel.Range = SubtractRanges(rng1, rng2)
                        // Dim results As List(Of Excel.Range) = SubtractRanges(srcRng3_prime, src_rng)
                        // Dim addressList As New List(Of String)

                        // Dim combinedAddress As String = ""
                        // For Each r In results
                        // addressList.Add(r.Address)

                        // Next

                        // combinedAddress = String.Join(",", addressList)
                        // GB_CB_Source3 = combinedAddress

                        // End If

                        src_rng = worksheet.get_Range("A1", worksheet.Cells[excelApp.Rows.Count, excelApp.Columns.Count]);
                        GlobalModule.GB_CB_Dlt3 = src_rng.get_Address();
                        GlobalModule.Nam3 = CB_Source.Text;

                    }
                    // MsgBox(src_rng.Address)
                }

                if (CB_multiselect.Checked == true)
                {
                    GlobalModule.TType1 = CB_Source.Text;
                }

                if (CB_checkbox.Checked == true)
                {
                    GlobalModule.TType2 = CB_Source.Text;
                }

                if (CB_search.Checked == true)
                {
                    GlobalModule.TType3 = CB_Source.Text;
                }


                Close();

            }
        }

        // Private Sub sheet_SelectionChange(ByVal Target As Excel.Range)
        // excelApp = Globals.ThisAddIn.Application
        // workBook = excelApp.ActiveWorkbook
        // workSheet = workBook.ActiveSheet
        // If GB_CB_Source1 IsNot Nothing Then

        // ' src_rng = workSheet.Range(GB_CB_Source1)
        // src_rng = workSheet.Range(GB_CB_Source1)

        // 'MsgBox(workSheet.Name)
        // 'MsgBox(src_rng.Worksheet.Name)

        // If CB_Source.Text.Contains("Active Workbook") Then
        // src_rng = workSheet.Range("A1", workSheet.Cells(excelApp.Rows.Count, excelApp.Columns.Count))
        // Else

        // End If

        // src_rng = workSheet.Range(GB_CB_Source1)

        // src_rng = workBook.ActiveSheet.range(src_rng.Address)


        // ' MsgBox(src_rng.Worksheet.Name)
        // 'Recheck: Newly added
        // If CB_Source.Text.Contains("Active Sheet") And Nam <> workSheet.Name Then
        // Exit Sub
        // End If

        // If IsCellInsideRange(Target, src_rng) And Target.Cells.Count = 1 And HasDataValidationList(Target) Then
        // 'If Target.Cells.Count = 1 Then ' Ensure only one cell is selected
        // TargetVar1 = Target.Address
        // If form Is Nothing OrElse form.IsDisposed Then
        // form = New Form36()
        // form.Show()
        // form.Refresh()
        // Else
        // ' If form is already open, bring it to the front
        // 'Form = Form36()
        // 'Form.Refresh()
        // 'Form.BringToFront()
        // 'Form.Refresh()
        // form.Dispose()
        // form = New Form36()
        // form.Show()
        // End If
        // End If

        // 'Dim form As New Form36()
        // 'form.Show()
        // 'form.Focus()
        // ''form.TopMost = True
        // ''form.Activate()
        // 'form.BringToFront()
        // 'End If
        // End If

        // End Sub


        private bool HasDataValidationList(Range cell)
        {
            bool hasValidation = false;

            try
            {
                if (cell.Validation is not null && cell.Validation.Type == (int)XlDVType.xlValidateList)
                {
                    hasValidation = true;
                }
            }
            catch (Exception ex)
            {
                // Exception will be thrown if cell doesn't have validation. No action needed.
            }

            return hasValidation;
        }


        private void Btn_Cancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            excelApp = Globals.ThisAddIn.Application;
            var workbook = excelApp.ActiveWorkbook;
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.ActiveSheet;

            // Dim results As List(Of Excel.Range)


            var rng1 = worksheet.get_Range("A1:D10");
            var rng2 = worksheet.get_Range("A1:B5");
            if (IsCellInsideRange(rng2, rng1) == true)
            {
                // Dim result As Excel.Range = SubtractRanges(rng1, rng2)
                var results = SubtractRanges(rng1, rng2);
                var addressList = new List<string>();

                string combinedAddress = "";
                foreach (var r in results)


                    addressList.Add(r.get_Address());

                combinedAddress = string.Join(",", addressList);
                worksheet.get_Range(combinedAddress).Select();

            }

            // worksheet.Select("A1:A3", "D1:D2")
            // If Not result Is Nothing Then
            // ' Do something with the result range
            // MessageBox.Show(result.Address)
            // Else
            // MessageBox.Show("Ranges are either equivalent or do not have a direct subtraction result.")
            // End If

            var rng3 = worksheet.get_Range("A1:A10");
            var rng4 = worksheet.get_Range("C1:C10");

            var combinedRange = excelApp.Union(rng3, rng4); // Assuming ExcelApp is your Excel.Application object

            // combinedRange.Select()
            // MsgBox(combinedRange.Address)
        }

        public List<Range> SubtractRanges(Range rng1, Range rng2)
        {
            var result = new List<Range>();

            // Top-left and bottom-right cells of rng1
            Range tl1 = (Range)rng1.Cells[1, 1];
            Range br1 = (Range)rng1.Cells[rng1.Rows.Count, rng1.Columns.Count];

            // Top-left and bottom-right cells of rng2
            Range tl2 = (Range)rng2.Cells[1, 1];
            Range br2 = (Range)rng2.Cells[rng2.Rows.Count, rng2.Columns.Count];

            // Check rows above rng2
            if (tl1.Row < tl2.Row)
            {
                result.Add(rng1.Worksheet.get_Range(tl1, rng1.Cells[tl2.Row - 1, br1.Column]));
            }

            // Check rows below rng2
            if (br1.Row > br2.Row)
            {
                result.Add(rng1.Worksheet.get_Range(rng1.Cells[br2.Row + 1, tl1.Column], br1));
            }

            // Check columns to the left of rng2
            if (tl1.Column < tl2.Column)
            {
                result.Add(rng1.Worksheet.get_Range(tl1, rng1.Cells[br1.Row, tl2.Column - 1]));
            }

            // Check columns to the right of rng2
            if (br1.Column > br2.Column)
            {
                result.Add(rng1.Worksheet.get_Range(rng1.Cells[tl1.Row, br2.Column + 1], br1));
            }

            return result;
        }


        private bool IsCellInsideRange(Range cell, Range targetRange)
        {
            // MsgBox(cell.Address)
            // MsgBox(targetRange.Address)
            try
            {
                var intersectRange = Globals.ThisAddIn.Application.Intersect(cell, targetRange);
                // MsgBox(intersectRange.Address)
                return intersectRange is not null;
            }
            catch (Exception ex)
            {
                // MsgBox(cell.Address)
                // MsgBox(targetRange.Address)
                return false;
            }
        }

        private void Selection_source_Click(object sender, EventArgs e)
        {
            try
            {

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
            }

            catch (Exception ex)
            {

                Show();
                TB_src_rng.Focus();

            }
        }

        private void Form41_RemoveAdavancedDropdownList_Load(object sender, EventArgs e)
        {
            KeyPreview = true;

            excelApp = Globals.ThisAddIn.Application;
            var workbook = excelApp.ActiveWorkbook;
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.ActiveSheet;

            // Timer1.Start()
            CB_Source.Items.Add("Select Range");
            CB_Source.Items.Add("Active Workbook :" + workbook.Name);
            CB_Source.Items.Add("Active Sheet :" + worksheet.Name);

            int i = 0;
            // Loop through each worksheet in the workbook.
            foreach (var WS in workbook.Sheets)
            {
                // Check if the worksheet is not hidden.
                if (Conversions.ToBoolean(Operators.AndObject(Operators.ConditionalCompareObjectEqual(WS.Visible, XlSheetVisibility.xlSheetVisible, false), Operators.ConditionalCompareObjectNotEqual(WS.name, worksheet.Name, false))))
                {
                    CB_Source.Items.Add(WS.Name);
                    i = i + 1;
                }
            }

            // Only Enable when select Range is selected in combobox
            if (CB_Source.Text == "Select Range")
            {
                TB_src_rng.Enabled = true;
                Selection_source.Enabled = true;
            }
            else
            {
                TB_src_rng.Enabled = false;
                Selection_source.Enabled = false;
            }

            if (excelApp.Selection is not null)
            {
                selectedRange = (Range)excelApp.Selection;
                src_rng = selectedRange;
                TB_src_rng.Text = selectedRange.get_Address();
                TB_src_rng.Focus();
                TB_src_rng.SelectionStart = TB_src_rng.Text.Length;
                // MsgBox(TB_src_rng.Text.Length)
            }
            CB_Source.SelectedIndex = 0;

            TB_src_rng.Focus();
        }

        private void excelApp_SheetSelectionChange(object Sh, Range selectionRange1)
        {
            try
            {

                excelApp = Globals.ThisAddIn.Application;
                if (focuschange == false)
                {


                    if (ReferenceEquals(ActiveControl, TB_src_rng))
                    {
                        src_rng = selectionRange1;
                        Activate();


                        BeginInvoke(new System.Action(() =>
                            {
                                TB_src_rng.Text = src_rng.get_Address();
                                SetWindowPos(Handle, new IntPtr(HWND_TOPMOST), 0, 0, 0, 0, SWP_NOACTIVATE | SWP_NOMOVE | SWP_NOSIZE);
                            }));
                        // TB_src_rng.Focus()

                        // TB_src_rng.Focus()
                    }



                }
            }

            catch (Exception ex)
            {

            }

        }

        private void TB_src_rng_TextChanged(object sender, EventArgs e)
        {
            try
            {
                // TB_src_rng.Focus()


                if (TB_src_rng.Text is not null & IsValidExcelCellReference(TB_src_rng.Text) == true)
                {
                    focuschange = true;

                    // Define the range of cells to read (for example, cells A1 to A10)
                    src_rng = excelApp.get_Range(TB_src_rng.Text);
                    src_rng.Select();
                    var range = src_rng;

                    Activate();
                    TB_src_rng.Focus();
                    // TB_src_rng.SelectionStart = TB_src_rng.Text.Length
                    TB_src_rng.SelectionStart = TB_src_rng.Text.Length;
                    focuschange = false;

                }
            }

            catch (Exception ex)
            {
                TB_src_rng.Focus();
            }
        }

        private void Form41_RemoveAdavancedDropdownList_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Btn_OK.Focus();
                Btn_OK.PerformClick();
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

        private void Form41_RemoveAdavancedDropdownList_Closing(object sender, CancelEventArgs e)
        {
            GlobalModule.form_flag = false;
        }

        private void Form41_RemoveAdavancedDropdownList_Disposed(object sender, EventArgs e)
        {
            GlobalModule.form_flag = false;
        }


        private void Form41_RemoveAdavancedDropdownList_Shown(object sender, EventArgs e)
        {
            TB_src_rng.Focus();
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