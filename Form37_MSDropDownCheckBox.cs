using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
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

    public partial class Form37_MSDropDownCheckBox
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

        private Excel.Worksheet _CurrentSheet;

        private Excel.Worksheet CurrentSheet
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _CurrentSheet;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                _CurrentSheet = value;
            }
        }
        private Excel.Workbook _WorkbookEvents;

        private Excel.Workbook WorkbookEvents
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _WorkbookEvents;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_WorkbookEvents != null)
                {
                    _WorkbookEvents.SheetActivate -= WorkbookEvents_SheetActivate;
                }

                _WorkbookEvents = value;
                if (_WorkbookEvents != null)
                {
                    _WorkbookEvents.SheetActivate += WorkbookEvents_SheetActivate;
                }
            }
        }


        private Form38 Form = null;

        private Range src_rng;
        public Range des_rng;
        private Range selectedRange;

        public Range validationRange;

        private bool processingEvent = false;
        public bool focuschange;

        public Form37_MSDropDownCheckBox()
        {


            Timer1 = new Timer() { Interval = 100 };
            InitializeComponent();
        }

        [DllImport("user32")]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
        private const uint SWP_NOMOVE = 0x2U;
        private const uint SWP_NOSIZE = 0x1U;
        private const uint SWP_NOACTIVATE = 0x10U;
        private const int HWND_TOPMOST = -1;

        [DllImport("user32.dll")]
        public static extern short GetAsyncKeyState(int vKey);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr GetForegroundWindow();

        private Timer Timer1;
        // xlApp = Globals.ThisAddIn.Application
        // Private xlWorkbook As Excel.Workbook
        // Private xlWorksheet As Excel.Worksheet

        private void Form1_HelpButtonClicked(object sender, CancelEventArgs e)
        {
            Process.Start("https://www.softeko.co/");
            e.Cancel = true; // This will suppress any additional event handling for the Help button
        }


        private void YourForm_Load(object sender, EventArgs e)
        {
            KeyPreview = true;
            excelApp = Globals.ThisAddIn.Application;
            var workbook = excelApp.ActiveWorkbook;
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.ActiveSheet;

            // Timer1.Start()
            CB_Source.Items.Add("Select Range");
            CB_Source.Items.Add("Active Sheet :" + worksheet.Name);
            CB_Source.Items.Add("Active Workbook :" + workbook.Name);

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

            try
            {

                excelApp = Globals.ThisAddIn.Application;

                excelApp.SheetSelectionChange += excelApp_SheetSelectionChange;

                // opened = opened + 1

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


                    if (ReferenceEquals(ActiveControl, TB_src_rng))
                    {
                        src_rng = selectionRange1;
                        Activate();


                        BeginInvoke(new System.Action(() =>
                            {
                                TB_src_rng.Text = src_rng.get_Address();
                                SetWindowPos(Handle, new IntPtr(HWND_TOPMOST), 0, 0, 0, 0, SWP_NOACTIVATE | SWP_NOMOVE | SWP_NOSIZE);
                            }));
                        TB_src_rng.Focus();
                        TB_src_rng.SelectionStart = TB_src_rng.Text.Length;
                        // TB_src_rng.Focus()
                    }



                }
            }

            catch (Exception ex)
            {

            }

        }

        private void CB_Source_SelectedIndexChanged(object sender, EventArgs e)
        {
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
        }

        private void Selection_source_Click(object sender, EventArgs e)
        {
            try
            {
                // If selectedRange Is Nothing Then
                // MsgBox(1)
                // Else

                // TB_src_rng.Text = selectedRange.Address


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
            }


            // End If

            catch (Exception ex)
            {

                Show();
                TB_src_rng.Focus();

            }
        }


        // Event handler for when any sheet in the workbook is activated
        private void WorkbookEvents_SheetActivate(object Sh)
        {
            // Detach event from previous sheet
            if (CurrentSheet is not null)
            {
                CurrentSheet.SelectionChange -= sheet_SelectionChange;
            }

            // Attach event to the new active sheet
            CurrentSheet = (Excel.Worksheet)Sh;
            CurrentSheet.SelectionChange += sheet_SelectionChange;
            // MsgBox(CurrentSheet.Name)
        }



        private void Btn_OK_Click(object sender, EventArgs e)
        {
            excelApp = Globals.ThisAddIn.Application;
            workBook = excelApp.ActiveWorkbook;
            workSheet = (Excel.Worksheet)workBook.ActiveSheet;

            if (CB_Source.Text == "Select Range" & string.IsNullOrEmpty(TB_src_rng.Text))
            {
                MessageBox.Show("Select a Source Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                TB_src_rng.Focus();
                // Me.Close()
                return;
            }

            else if (CB_Source.Text == "Select Range" & IsValidExcelCellReference(TB_src_rng.Text) == false)
            {
                MessageBox.Show("Select a Valid Source Range.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

                if (GlobalModule.settingflag2 == true)
                {

                    if (RB_Horizontal.Checked == true)
                    {
                        GlobalModule.Horizontal2 = true;
                    }
                    else
                    {
                        GlobalModule.Horizontal2 = false;
                    }

                    if (CB_Search.Checked == true)
                    {
                        GlobalModule.Search2 = true;
                    }
                    else
                    {
                        GlobalModule.Search2 = false;
                    }

                    GlobalModule.Flag2 = true;

                    GlobalModule.Separator2 = CB_Separator.Text;
                }
                else
                {
                    if (CB_Source.Text.Contains("Active Sheet"))
                    {

                        src_rng = workSheet.get_Range("A1", workSheet.Cells[excelApp.Rows.Count, excelApp.Columns.Count]);
                    }

                    else if (CB_Source.Text.Contains("Active Workbook"))
                    {

                        src_rng = workSheet.get_Range("A1", workSheet.Cells[excelApp.Rows.Count, excelApp.Columns.Count]);

                    }

                    if (RB_Horizontal.Checked == true)
                    {
                        GlobalModule.Horizontal2 = true;
                    }
                    else
                    {
                        GlobalModule.Horizontal2 = false;
                    }

                    if (CB_Search.Checked == true)
                    {
                        GlobalModule.Search2 = true;
                    }
                    else
                    {
                        GlobalModule.Search2 = false;
                    }

                    GlobalModule.Flag2 = true;
                    GlobalModule.GB_CB_Source2 = TB_src_rng.Text;
                    GlobalModule.Separator2 = CB_Separator.Text;

                    // Private EventDel_CellsChange As Excel.DocEvents_ChangeEventHandler
                    int i = 1;

                    workSheet.get_Range("B1").Select();  // Randomly select a cell. If nothing is selected, addhandler show error

                    // Define an array of type Excel.Worksheet
                    Excel.Worksheet[] sheetsArray;

                    // Resize the array based on the number of sheets
                    sheetsArray = new Excel.Worksheet[workBook.Worksheets.Count + 1];

                    if (CB_Source.Text.Contains("Active Workbook"))
                    {
                        // MsgBox(1)
                        // Assuming you're working with the active workbook:
                        workSheet.SelectionChange += sheet_SelectionChange;
                        WorkbookEvents = Globals.ThisAddIn.Application.ActiveWorkbook;
                        CurrentSheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
                    }

                    // For i = 1 To excelApp.ActiveWorkbook.Worksheets.Count
                    // MsgBox(excelApp.ActiveWorkbook.Worksheets.Count)
                    // sheetsArray(i) = CType(workBook.Worksheets(i), Excel.Worksheet)

                    // 'Needs to add handler to each worksheet
                    // If sheetsArray(i).Visible = Excel.XlSheetVisibility.xlSheetVisible Then
                    // AddHandler sheetsArray(i).SelectionChange, AddressOf sheet_SelectionChange
                    // End If
                    // Next


                    else if (CB_Source.Text == "Select Range" | CB_Source.Text.Contains("Active Sheet"))
                    {

                        workSheet.SelectionChange += sheet_SelectionChange;
                    }

                    // ElseIf CB_Source.Text.Contains("Active Sheet") Then

                    // AddHandler workSheet.SelectionChange, AddressOf sheet_SelectionChange
                    else
                    {

                        var loopTo = excelApp.ActiveWorkbook.Worksheets.Count;
                        for (i = 1; i <= loopTo; i++)
                        {
                            sheetsArray[i] = (Excel.Worksheet)workBook.Worksheets[i];
                            if ((CB_Source.Text ?? "") == (sheetsArray[i].Name ?? ""))
                            {
                                sheetsArray[i].SelectionChange += sheet_SelectionChange;
                                // src_rng = sheetsArray(i).Range("A1", workSheet.Cells(excelApp.Rows.Count, excelApp.Columns.Count))
                            }
                            // i = i + 1

                        }
                        src_rng = workSheet.get_Range("A1", workSheet.Cells[excelApp.Rows.Count, excelApp.Columns.Count]);

                    }

                    // EventDel_CellsChange = New Excel.DocEvents_SelectionChangeEventHandler(AddressOf sheet_SelectionChange)
                    // AddHandler xlSheet1.Change, EventDel_CellsChange


                    if (TB_src_rng.Enabled == true)
                    {
                        GlobalModule.GB_CB_Source2 = TB_src_rng.Text; // SR is the global variable for Source Range
                    }
                    else
                    {
                        GlobalModule.GB_CB_Source2 = src_rng.get_Address();
                    }

                    GlobalModule.GB_CB_Source2 = src_rng.get_Address();

                    GlobalModule.RangeType2 = CB_Source.Text;


                    GlobalModule.TType2 = "";
                    GlobalModule.SR2 = CB_Source.Text;
                    GlobalModule.shName2 = workSheet.Name;


                }


                Excel.Worksheet workSheet2 = (Excel.Worksheet)workBook.ActiveSheet;



                // Check if "Neww" worksheet exists and delete it if it does
                foreach (var ws in workBook.Sheets)
                {
                    if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(ws.Name, "SoftekoSofteko", false)))
                    {
                        ws.Delete();
                        break;
                    }
                }

                // Add a new worksheet named "Neww"
                workSheet2 = (Excel.Worksheet)workBook.Worksheets.Add();
                workSheet2.Name = "SoftekoSofteko";

                // Add your values (here's an example to set A1 to "Sample Value")
                workSheet2.get_Range("A1").set_Value(value: "Do not Delete thesheet!");
                // ... Add more values as required ...



                // Hide the worksheet
                workSheet2.Visible = XlSheetVisibility.xlSheetHidden;

                workSheet2.get_Range("A2").set_Value(value: GlobalModule.GB_CB_Source2);   // For Source Range
                workSheet2.get_Range("A3").set_Value(value: GlobalModule.SR2);             // Range Type
                workSheet2.get_Range("A4").set_Value(value: GlobalModule.Horizontal2);
                workSheet2.get_Range("A5").set_Value(value: GlobalModule.Separator2);
                workSheet2.get_Range("A6").set_Value(value: GlobalModule.Search2);
                workSheet2.get_Range("A7").set_Value(value: GlobalModule.Flag2);            // Activated
                workSheet2.get_Range("A8").set_Value(value: GlobalModule.TargetVar2);
                workSheet2.get_Range("A9").set_Value(value: GlobalModule.shName2);
                workSheet2.get_Range("A2").set_Value(value: GlobalModule.GB_CB_Source2);

                workSheet2.get_Range("B2").set_Value(value: CB_Source.Text);

                Close();

            }
        }

        private void Btn_Cancel_Click(object sender, EventArgs e)
        {
            Close();
        }


        private void sheet_SelectionChange(Range Target)
        {
            excelApp = Globals.ThisAddIn.Application;
            workBook = excelApp.ActiveWorkbook;
            workSheet = (Excel.Worksheet)workBook.ActiveSheet;

            Range src_rng_concate;
            // MsgBox(workSheet.Name)
            // MsgBox(src_rng.Worksheet.Name)

            // src_rng = workSheet.Range(GB_CB_Source1)

            if (GlobalModule.GB_CB_Source2 is not null)
            {
                src_rng = workSheet.get_Range(GlobalModule.GB_CB_Source2);
                // MsgBox(src_rng.Address)

                // MsgBox(src_rng.Address)
                // src_rng = workSheet.Range(GB_CB_Source1)


                if (CB_Source.Text.Contains("Active Workbook"))
                {
                    src_rng = workSheet.get_Range("A1", workSheet.Cells[excelApp.Rows.Count, excelApp.Columns.Count]);
                }
                src_rng = (Range)workBook.ActiveSheet.range(src_rng.get_Address());
                // MsgBox(src_rng.Worksheet.Name)

                // Change starts from here
                if ((GlobalModule.Nam2 ?? "") == (workSheet.Name ?? "") & GlobalModule.TType2 == "Select Range" | (GlobalModule.Nam2 ?? "") == (workSheet.Name ?? "") & GlobalModule.TType2.Contains("Active Sheet") | (GlobalModule.Nam2 ?? "") == (workSheet.Name ?? "") & (GlobalModule.TType2 ?? "") == (workSheet.Name ?? ""))
                {

                    src_rng_concate = workSheet.get_Range(GlobalModule.GB_CB_Dlt2);
                    // MsgBox(src_rng_concate.Address)

                    if (IsCellInsideRange(Target, src_rng) == true & Target.Cells.Count == 1 & HasDataValidationList(Target) & IsCellInsideRange(Target, src_rng_concate) == true)
                    {
                        // MsgBox(1)
                        // If Target.Cells.Count = 1 Then ' Ensure only one cell is selected
                        GlobalModule.TargetVar2 = Target.get_Address();
                        if (Form is not null)
                        {
                            // Form = Nothing

                            Form.Dispose();
                            // MsgBox(2)
                        }
                    }







                    else if (IsCellInsideRange(Target, src_rng) & Target.Cells.Count == 1 & HasDataValidationList(Target))
                    {
                        // MsgBox(2)
                        // If Target.Cells.Count = 1 Then ' Ensure only one cell is selected
                        GlobalModule.TargetVar2 = Target.get_Address();
                        if (Form is null || Form.IsDisposed)
                        {
                            Form = new Form38();
                            Form.Show();
                            Form.BringToFront();
                            Form.Refresh();
                        }
                        else
                        {
                            // If form is already open, bring it to the front

                            Form.Dispose();
                            Form = new Form38();
                            Form.Show();
                            Form.BringToFront();

                        }

                        // Dim form As New Form36()
                        // form.Show()
                        // form.Focus()
                        // 'form.TopMost = True
                        // 'form.Activate()
                        // form.BringToFront()
                        // End If
                    }
                }

                else if (IsCellInsideRange(Target, src_rng) & Target.Cells.Count == 1 & HasDataValidationList(Target))
                {
                    // MsgBox(2)
                    // MsgBox(10)
                    // If Target.Cells.Count = 1 Then ' Ensure only one cell is selected
                    GlobalModule.TargetVar2 = Target.get_Address();
                    if (Form is null || Form.IsDisposed)
                    {
                        Form = new Form38();
                        Form.Show();
                        Form.BringToFront();
                        Form.Refresh();
                    }
                    else
                    {
                        // If form is already open, bring it to the front

                        Form.Dispose();
                        Form = new Form38();
                        Form.Show();
                        Form.BringToFront();
                        // MsgBox(3)

                    }
                }
            }
            else
            {

                // If Form IsNot Nothing Then
                // 'Form = Nothing

                // Form.Dispose()
                // 'MsgBox(2)
                // End If


            }
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

        private void CB_Separator_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void CB_Separator_KeyUp(object sender, KeyEventArgs e)
        {
            // Check if the Enter key was pressed
            if (e.KeyCode == Keys.Enter)
            {
                // Add the current text in the ComboBox to the items collection
                if (!string.IsNullOrEmpty(CB_Separator.Text) && !CB_Separator.Items.Contains(CB_Separator.Text))
                {
                    CB_Separator.Items.Add(CB_Separator.Text);
                }


            }
        }

        private void Form37_MSDropDownCheckBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Btn_OK.Focus();
                Btn_OK.PerformClick();
            }
        }

        private void TB_src_rng_TextChanged(object sender, EventArgs e)
        {
            try
            {
                TB_src_rng.Focus();
                TB_src_rng.SelectionStart = TB_src_rng.Text.Length;

                if (TB_src_rng.Text is not null & IsValidExcelCellReference(TB_src_rng.Text) == true)
                {
                    focuschange = true;

                    // Define the range of cells to read (for example, cells A1 to A10)
                    src_rng = excelApp.get_Range(TB_src_rng.Text);
                    src_rng.Select();
                    var range = src_rng;

                    Activate();
                    TB_src_rng.Focus();
                    TB_src_rng.SelectionStart = TB_src_rng.Text.Length;
                    focuschange = false;

                }
            }

            catch (Exception ex)
            {
                TB_src_rng.Focus();

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

        private void Form37_MSDropDownCheckBox_Activated(object sender, EventArgs e)
        {
            TB_src_rng.Focus();
            TB_src_rng.SelectionStart = TB_src_rng.Text.Length;
        }

        private void Form37_MSDropDownCheckBox_Shown(object sender, EventArgs e)
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



        private void Form37_MSDropDownCheckBox_Closing(object sender, CancelEventArgs e)
        {
            GlobalModule.form_flag = false;
        }

        private void Form37_MSDropDownCheckBox_Disposed(object sender, EventArgs e)
        {
            GlobalModule.form_flag = false;
        }
    }

    public class WorksheetHandler
    {
        private Excel.Worksheet _Sheet;

        public virtual Excel.Worksheet Sheet
        {
            [MethodImpl(MethodImplOptions.Synchronized)]
            get
            {
                return _Sheet;
            }

            [MethodImpl(MethodImplOptions.Synchronized)]
            set
            {
                if (_Sheet != null)
                {
                    _Sheet.SelectionChange -= Worksheet_SelectionChange;
                }

                _Sheet = value;
                if (_Sheet != null)
                {
                    _Sheet.SelectionChange += Worksheet_SelectionChange;
                }
            }
        }

        public WorksheetHandler(ref Excel.Worksheet ws)
        {
            Sheet = ws;
        }

        private void Worksheet_SelectionChange(Range Target)
        {
            // The event code goes here
        }
    }
}