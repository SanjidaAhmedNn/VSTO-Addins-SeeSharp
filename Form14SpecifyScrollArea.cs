using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace VSTO_Addins
{

    public partial class Form14SpecifyScrollArea
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
        private Excel.Workbook workbook;
        private Excel.Worksheet worksheet, worksheet1;
        private Excel.Worksheet outWorksheet;
        private Range inputRng;
        private int FocusedTxtBox;
        private Range selectedRange;
        private bool txtChanged = false;
        public static List<int> all_hidden_Row_No = new List<int>();
        public static List<int> all_hidden_Col_No = new List<int>();
        public static bool scroll_Area_Specified = false;

        public Form14SpecifyScrollArea()
        {
            InitializeComponent();
        }

        [DllImport("user32")]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
        private const uint SWP_NOMOVE = 0x2U;
        private const uint SWP_NOSIZE = 0x1U;
        private const uint SWP_NOACTIVATE = 0x10U;
        private const int HWND_TOPMOST = -1;

        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Btn_OK.PerformClick();
            }
        }

        private void Form14SpecifyScrollArea_Load(object sender, EventArgs e)
        {
            try
            {

                excelApp = Globals.ThisAddIn.Application;
                workbook = excelApp.ActiveWorkbook;
                worksheet = (Excel.Worksheet)workbook.ActiveSheet;

                Range selectedRng = (Range)excelApp.Selection;
                txtSourceRange.Text = selectedRng.get_Address();

                KeyPreview = true;



                excelApp = Globals.ThisAddIn.Application;
                workbook = excelApp.ActiveWorkbook;
                worksheet = (Excel.Worksheet)workbook.ActiveSheet;

                // store the row numbers in a list, if a row of the used range of the worksheet is hidden

                for (int i = 1, loopTo = worksheet.UsedRange.Rows.Count; i <= loopTo; i++)
                {
                    if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(worksheet.UsedRange.Cells[i, 1].entirerow.hidden, true, false)))
                    {
                        all_hidden_Row_No.Add(Conversions.ToInteger(worksheet.UsedRange.Cells[i, 1].row));
                    }
                }

                // store the column numbers in a list, if a column of the used range of the worksheet is hidden

                for (int j = 1, loopTo1 = worksheet.UsedRange.Columns.Count; j <= loopTo1; j++)
                {
                    if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(worksheet.UsedRange.Cells[1, j].entirecolumn.hidden, true, false)))
                    {
                        all_hidden_Col_No.Add(Conversions.ToInteger(worksheet.UsedRange.Cells[1, j].column));
                    }
                }
            }


            catch (Exception ex)
            {

            }



        }

        private void txtSourceRange_TextChanged(object sender, EventArgs e)
        {

            try
            {
                excelApp = Globals.ThisAddIn.Application;
                workbook = excelApp.ActiveWorkbook;
                worksheet = (Excel.Worksheet)workbook.ActiveSheet;

                txtChanged = true;

                inputRng = worksheet.get_Range(txtSourceRange.Text);
                inputRng.Select();
            }


            catch (Exception ex)
            {

            }

            txtChanged = false;
            txtSourceRange.Focus();


        }


        private void txtSourceRange_GotFocus(object sender, EventArgs e)
        {
            try
            {

                FocusedTxtBox = 1;
            }


            catch (Exception ex)
            {

            }
        }
        private void Form1_Activated(object sender, EventArgs e)
        {
            try
            {

                excelApp = Globals.ThisAddIn.Application;

                excelApp.SheetSelectionChange += rngSelectionFromTxtBox;
            }

            catch (Exception ex)
            {

            }

        }
        private void rngSelectionFromTxtBox(object Sh, Range Target)
        {

            try
            {

                excelApp = Globals.ThisAddIn.Application;
                selectedRange = (Range)excelApp.Selection;
                selectedRange.Select();

                txtSourceRange.Focus();

                if (txtChanged == false)
                {

                    if (FocusedTxtBox == 1)
                    {

                        txtSourceRange.Text = selectedRange.get_Address();
                        worksheet = (Excel.Worksheet)workbook.ActiveSheet;
                        inputRng = selectedRange;
                        txtSourceRange.Focus();

                    }

                }
            }



            catch (Exception ex)
            {


            }

        }

        private void Selection_Click(object sender, EventArgs e)
        {

            try
            {

                excelApp = Globals.ThisAddIn.Application;
                workbook = excelApp.ActiveWorkbook;
                worksheet = (Excel.Worksheet)workbook.ActiveSheet;
                selectedRange = (Range)excelApp.Selection;
                txtSourceRange.Focus();

                Hide();
                inputRng = (Range)excelApp.InputBox("Please Select a Range", "Range Selection", selectedRange.get_Address(), Type: 8);
                Show();

                inputRng.Worksheet.Activate();
                txtSourceRange.Text = inputRng.get_Address();
                inputRng.Select();
                txtSourceRange.Focus();
            }


            catch (Exception ex)
            {

                txtSourceRange.Focus();

            }


        }

        private void Btn_Cancel_Click(object sender, EventArgs e)
        {

            Dispose();

        }

        public bool IsValidRng(string input)
        {

            string pattern = @"^(\$?[A-Z]+\$?[0-9]+(:\$?[A-Z]+\$?[0-9]+)?)(,\$?[A-Z]+\$?[0-9]+(:\$?[A-Z]+\$?[0-9]+)?)*$";
            return System.Text.RegularExpressions.Regex.IsMatch(input, pattern);

        }


        private void Btn_OK_Click(object sender, EventArgs e)
        {

            try
            {
                excelApp = Globals.ThisAddIn.Application;
                workbook = excelApp.ActiveWorkbook;
                worksheet = (Excel.Worksheet)workbook.ActiveSheet;


                // checks if the user clicked OK button with an empty sourceRange textbox
                // if it is non-empty, then checks is the used range is a valid range or not
                // if any of these are true then it will give user another chance to enter correct input
                if (string.IsNullOrEmpty(txtSourceRange.Text))
                {
                    Interaction.MsgBox("Please provide source range.", MsgBoxStyle.Exclamation, "Error!");
                    txtSourceRange.Focus();
                    return;
                }
                else if (IsValidRng(txtSourceRange.Text.ToUpper()) == false)
                {
                    Interaction.MsgBox("Please provide a valid source range.", MsgBoxStyle.Exclamation, "Error!");
                    txtSourceRange.Text = "";
                    txtSourceRange.Focus();
                    return;
                }

                // counts the number of ranges used by user
                int rngCount;
                rngCount = 0;

                foreach (char c in txtSourceRange.Text)
                {

                    if (Conversions.ToString(c) == ",")
                    {
                        rngCount = rngCount + 1;
                    }

                }

                // calls different subs based on number of ranges in users' selection 
                if (rngCount == 0)
                {
                    singleRng();
                }
                else
                {
                    multiRng();
                }
            }


            catch (Exception ex)
            {

            }



        }

        private void singleRng()
        {

            // this sub will be called when user selected a single range as input

            try
            {
                excelApp = Globals.ThisAddIn.Application;
                workbook = excelApp.ActiveWorkbook;
                worksheet = (Excel.Worksheet)workbook.ActiveSheet;

                Range selectedRng;
                selectedRng = worksheet.get_Range(txtSourceRange.Text);

                // keeps the range address from the textbox in a variable and keeps the worksheet info in another variable named "worksheet1"
                string temp;
                temp = txtSourceRange.Text;
                worksheet1 = inputRng.Worksheet;

                // checks if user opted to backup the sheet. If yes then create a copy and reactivate the original worksheet
                if (CheckBox.Checked == true)
                {

                    workbook.ActiveSheet.Copy(After: workbook.Sheets[workbook.Sheets.Count]);
                    outWorksheet = (Excel.Worksheet)workbook.Sheets[workbook.Sheets.Count];

                    worksheet1.Activate();
                    txtSourceRange.Text = temp;

                }

                // cellCount variable is used to count the number of cells in users' selection.
                // Our goal is to check whether the cellCount is <= 4 or not in the next block.
                // if the cellCount exceeds 5 then exit from the loop.
                int cellCount = 0;
                for (int i = 1, loopTo = selectedRng.Rows.Count; i <= loopTo; i++)
                {
                    for (int j = 1, loopTo1 = selectedRng.Columns.Count; j <= loopTo1; j++)
                    {
                        cellCount += 1;
                        if (cellCount > 5)
                            break;
                    }
                    if (cellCount > 5)
                        break;
                }

                // checks if the cellCount is <=6 or not. If yes then show a YesNo msgbox as warning.
                // If user select yes then continue excecuting next lines, else dispose the form
                if (cellCount <= 4)
                {
                    MsgBoxResult answer;
                    answer = Interaction.MsgBox("Do you really want to hide everything except " + cellCount + " cells." + Microsoft.VisualBasic.Constants.vbCrLf + "If yes, hide every cell except the selected cell range. If no, close the add-in.", MsgBoxStyle.YesNo, "Warning!");
                    if (answer == MsgBoxResult.Yes)
                    {
                        goto Proceed;
                    }
                    else
                    {
                        goto break;
                    }
                }

Proceed:
                ;

                // store the row numbers in a list, if a row of the selected range is hidden
                var hidden_Row_No = new List<int>();
                for (int i = 1, loopTo2 = selectedRng.Rows.Count; i <= loopTo2; i++)
                {
                    if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(selectedRng.Cells[i, 1].entirerow.hidden, true, false)))
                    {
                        hidden_Row_No.Add(Conversions.ToInteger(selectedRng.Cells[i, 1].row));
                    }
                }

                // store the column numbers in a list, if a column of the selected range is hidden
                var hidden_Col_No = new List<int>();
                for (int j = 1, loopTo3 = selectedRng.Columns.Count; j <= loopTo3; j++)
                {
                    if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(selectedRng.Cells[1, j].entirecolumn.hidden, true, false)))
                    {
                        hidden_Col_No.Add(Conversions.ToInteger(selectedRng.Cells[1, j].column));
                    }
                }

                // hide all rows and columns
                worksheet.Rows.Hidden = true;
                worksheet.Columns.Hidden = true;


                // unhide all rows and columns of the selected range
                selectedRng.EntireRow.Hidden = false;
                selectedRng.EntireColumn.Hidden = false;

                scroll_Area_Specified = true;

                // loop through each element of the hidden_Row_No list, and fetch the row numbers that were hidden in the selected range
                // hide those rows
                for (int i = 0, loopTo4 = hidden_Row_No.Count - 1; i <= loopTo4; i++)
                    worksheet.Rows[hidden_Row_No[i]].hidden = (object)true;

                // loop through each element of the hidden_Col_No list, and fetch the column numbers that were hidden in the selected range
                // hide those columns
                for (int i = 0, loopTo5 = hidden_Col_No.Count - 1; i <= loopTo5; i++)
                    worksheet.Columns[hidden_Col_No[i]].hidden = (object)true;

                selectedRng.Select();

break:
                ;


                Dispose();
            }


            catch (Exception ex)
            {

            }


        }


        private void multiRng()
        {

            // this sub will be called when user selected multiple ranges as input

            excelApp = Globals.ThisAddIn.Application;
            workbook = excelApp.ActiveWorkbook;
            worksheet = (Excel.Worksheet)workbook.ActiveSheet;

            try
            {

                excelApp = Globals.ThisAddIn.Application;
                workbook = excelApp.ActiveWorkbook;
                worksheet = (Excel.Worksheet)workbook.ActiveSheet;

                // keeps the range address from the textbox in a variable and keeps the worksheet info in another variable named "worksheet1"
                string temp;
                temp = txtSourceRange.Text;
                worksheet1 = inputRng.Worksheet;

                // checks if user opted to backup the sheet. If yes then create a copy and reactivate the original worksheet
                if (CheckBox.Checked == true)
                {

                    workbook.ActiveSheet.Copy(After: workbook.Sheets[workbook.Sheets.Count]);
                    outWorksheet = (Excel.Worksheet)workbook.Sheets[workbook.Sheets.Count];

                    worksheet1.Activate();
                    txtSourceRange.Text = temp;

                }

                // keeps each of the range addresses from users' selecion in separate array elements of the arrRng array
                string[] arrRng = Strings.Split(txtSourceRange.Text, ",");

                // finds the start and end row, column numbers and store the range in scrollArea variable as range
                int minRow = int.MaxValue;
                int maxRow = int.MinValue;
                int minCol = int.MaxValue;
                int maxCol = int.MinValue;

                foreach (var address in arrRng)
                {
                    var range = worksheet.get_Range(address);
                    minRow = Math.Min(minRow, range.Row);
                    maxRow = Math.Max(maxRow, range.Row + range.Rows.Count - 1);
                    minCol = Math.Min(minCol, range.Column);
                    maxCol = Math.Max(maxCol, range.Column + range.Columns.Count - 1);
                }
                var scrollArea = worksheet.get_Range(worksheet.Cells[minRow, minCol], worksheet.Cells[maxRow, maxCol]);

                var hidden_Row_No = new List<int>();
                var hidden_Col_No = new List<int>();

                // loop through each range that user have selected
                // store the hidden row and column numbers of the selected ranges in 2 lists that decalred above
                for (int k = 0, loopTo = Information.UBound(arrRng); k <= loopTo; k++)
                {

                    // store the row numbers in a list, if a row of the selected range is hidden
                    for (int i = 1, loopTo1 = worksheet.get_Range(arrRng[k]).Rows.Count; i <= loopTo1; i++)
                    {
                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(worksheet.get_Range(arrRng[k]).Cells[i, 1].entirerow.hidden, true, false)))
                        {
                            hidden_Row_No.Add(Conversions.ToInteger(worksheet.get_Range(arrRng[k]).Cells[i, 1].row));
                        }
                    }

                    // store the column numbers in a list, if a column of the selected range is hidden
                    for (int j = 1, loopTo2 = worksheet.get_Range(arrRng[k]).Columns.Count; j <= loopTo2; j++)
                    {
                        if (Conversions.ToBoolean(Operators.ConditionalCompareObjectEqual(worksheet.get_Range(arrRng[k]).Cells[1, j].entirecolumn.hidden, true, false)))
                        {
                            hidden_Col_No.Add(Conversions.ToInteger(worksheet.get_Range(arrRng[k]).Cells[1, j].column));
                        }
                    }
                }



                // declare a booolean variable named "flag" with Fasle value
                // if the number of rows and the row number of 1st row of each range is same then flag will be True
                // if the number of columns and the column number of 1st column of each range is same then flag will be True
                // otherwise it flag will be false
                bool flag = false;
                for (int i = 0, loopTo3 = Information.UBound(arrRng) - 1; i <= loopTo3; i++)
                {

                    if (worksheet.get_Range(arrRng[i]).Rows.Count == worksheet.get_Range(arrRng[i + 1]).Rows.Count & worksheet.get_Range(arrRng[i]).Row == worksheet.get_Range(arrRng[i + 1]).Row)
                    {

                        flag = true;
                    }

                    else if (worksheet.get_Range(arrRng[i]).Columns.Count == worksheet.get_Range(arrRng[i + 1]).Columns.Count & worksheet.get_Range(arrRng[i]).Column == worksheet.get_Range(arrRng[i + 1]).Column)
                    {

                        flag = true;
                    }

                    else
                    {

                        flag = false;

                    }

                }

                // checks if the flag is true or false
                // muiltiple ranges will be hidden only if the the flag is true
                // otherwise a msgbox will open and give user another chance to enter correct inputs
                if (flag == false)
                {
                    Interaction.MsgBox("Multiple selection is not possible with this source range.", MsgBoxStyle.Exclamation, "Error!");
                    txtSourceRange.Clear();
                    txtSourceRange.Focus();
                }
                else
                {
                    worksheet.Rows.Hidden = true;
                    worksheet.Columns.Hidden = true;

                    for (int i = 0, loopTo4 = Information.UBound(arrRng); i <= loopTo4; i++)
                    {
                        worksheet.get_Range(arrRng[i]).EntireRow.Hidden = false;
                        worksheet.get_Range(arrRng[i]).EntireColumn.Hidden = false;
                    }

                    scroll_Area_Specified = true;

                    // loop through each element of the hidden_Row_No list, and fetch the row numbers that were hidden in the selected range
                    // hide those rows
                    for (int i = 0, loopTo5 = hidden_Row_No.Count - 1; i <= loopTo5; i++)
                        worksheet.Rows[hidden_Row_No[i]].hidden = (object)true;

                    // loop through each element of the hidden_Col_No list, and fetch the column numbers that were hidden in the selected range
                    // hide those columns
                    for (int i = 0, loopTo6 = hidden_Col_No.Count - 1; i <= loopTo6; i++)
                        worksheet.Columns[hidden_Col_No[i]].hidden = (object)true;

                    scrollArea.Select();
                    Dispose();

                }
            }


            catch (Exception ex)
            {

            }

        }
        private void Form14SpecifyScrollArea_Disposed(object sender, EventArgs e)
        {
            GlobalModule.form_flag = false;
        }

        private void Form14SpecifyScrollArea_Closing(object sender, CancelEventArgs e)
        {
            GlobalModule.form_flag = false;
        }



        private void Form14SpecifyScrollArea_Shown(object sender, EventArgs e)
        {
            Focus();
            BringToFront();
            Activate();
            BeginInvoke(new System.Action(() =>
                {
                    txtSourceRange.Text = inputRng.get_Address();
                    SetWindowPos(Handle, new IntPtr(HWND_TOPMOST), 0, 0, 0, 0, SWP_NOACTIVATE | SWP_NOMOVE | SWP_NOSIZE);
                }));
        }
    }
}