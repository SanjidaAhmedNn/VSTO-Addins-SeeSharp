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

    public partial class Form12HideRanges
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
        private int rngCount;
        private string[] arrRng;

        public Form12HideRanges()
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
                btn_OK.PerformClick();
            }
        }

        private void Form12HideRanges_Load(object sender, EventArgs e)
        {

            excelApp = Globals.ThisAddIn.Application;
            workbook = excelApp.ActiveWorkbook;
            worksheet = (Excel.Worksheet)workbook.ActiveSheet;

            Range selectedRng = (Range)excelApp.Selection;
            txtSourceRange.Text = selectedRng.get_Address();


            rngCount = 0;
            foreach (char c in txtSourceRange.Text)
            {

                if (Conversions.ToString(c) == ",")
                {
                    rngCount = rngCount + 1;
                }

            }

            if (rngCount == 0)
            {
                RB_Single_Range.Checked = true;
            }
            else if (rngCount > 0)
            {
                RB_Multiple_Range.Checked = true;
            }



            RB_Row.Checked = true;

            KeyPreview = true;
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


                rngCount = 0;
                foreach (char c in txtSourceRange.Text)
                {

                    if (Conversions.ToString(c) == ",")
                    {
                        rngCount = rngCount + 1;
                    }

                }

                if (rngCount == 0)
                {
                    RB_Single_Range.Checked = true;
                }
                else if (rngCount > 0)
                {
                    RB_Multiple_Range.Checked = true;
                }
            }


            catch (Exception ex)
            {

            }




            txtChanged = false;
            txtSourceRange.Focus();



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
            }


            catch (Exception ex)
            {

                txtSourceRange.Focus();

            }



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

        private void AutoSelection_Click(object sender, EventArgs e)
        {

            try
            {

                // excelApp = Globals.ThisAddIn.Application
                // workbook = excelApp.ActiveWorkbook
                // worksheet = workbook.ActiveSheet
                // selectedRange = excelApp.Selection
                // selectedRange = selectedRange.Cells(1, 1)
                // selectedRange.Select()

                // Dim topLeft, bottomRight As String



                // If selectedRange.Offset(0, -1).Value = Nothing And selectedRange.Offset(0, 1).Value = Nothing And selectedRange.Offset(-1, 0).Value = Nothing Then
                // topLeft = selectedRange.Address
                // bottomRight = worksheet.Range(topLeft).End(XlDirection.xlDown).Address
                // selectedRange = worksheet.Range(worksheet.Range(topLeft), worksheet.Range(bottomRight))

                // ElseIf selectedRange.Offset(-1, 0).Value = Nothing And selectedRange.Offset(1, 0).Value = Nothing And selectedRange.Offset(0, -1).Value = Nothing Then

                // topLeft = selectedRange.Address
                // bottomRight = worksheet.Range(topLeft).End(XlDirection.xlToRight).Address
                // selectedRange = worksheet.Range(worksheet.Range(topLeft), worksheet.Range(bottomRight))

                // ElseIf selectedRange.Offset(0, -1).Value = Nothing And selectedRange.Offset(-1, 0).Value = Nothing Then
                // bottomRight = selectedRange.End(XlDirection.xlToRight).Address
                // bottomRight = worksheet.Range(bottomRight).End(XlDirection.xlDown).Address

                // selectedRange = worksheet.Range(selectedRange, worksheet.Range(bottomRight))

                // ElseIf selectedRange.Offset(0, -1).Value = Nothing And selectedRange.Offset(0, 1).Value = Nothing Then

                // topLeft = selectedRange.End(XlDirection.xlUp).Address
                // bottomRight = worksheet.Range(topLeft).End(XlDirection.xlDown).Address
                // selectedRange = worksheet.Range(worksheet.Range(topLeft), worksheet.Range(bottomRight))

                // ElseIf selectedRange.Offset(-1, 0).Value = Nothing And selectedRange.Offset(1, 0).Value = Nothing Then
                // topLeft = selectedRange.End(XlDirection.xlToLeft).Address
                // bottomRight = worksheet.Range(topLeft).End(XlDirection.xlToRight).Address
                // selectedRange = worksheet.Range(worksheet.Range(topLeft), worksheet.Range(bottomRight))

                // ElseIf selectedRange.Offset(0, -1).Value = Nothing Then
                // topLeft = selectedRange.End(XlDirection.xlUp).Address
                // bottomRight = worksheet.Range(topLeft).End(XlDirection.xlToRight).Address
                // bottomRight = worksheet.Range(bottomRight).End(XlDirection.xlDown).Address
                // selectedRange = worksheet.Range(worksheet.Range(topLeft), worksheet.Range(bottomRight))


                // ElseIf selectedRange.Offset(-1, 0).Value = Nothing Then

                // topLeft = selectedRange.End(XlDirection.xlToLeft).Address
                // bottomRight = worksheet.Range(topLeft).End(XlDirection.xlToRight).Address
                // bottomRight = worksheet.Range(bottomRight).End(XlDirection.xlDown).Address
                // selectedRange = worksheet.Range(worksheet.Range(topLeft), worksheet.Range(bottomRight))



                // Else
                // topLeft = selectedRange.End(XlDirection.xlToLeft).Address
                // topLeft = worksheet.Range(topLeft).End(XlDirection.xlUp).Address
                // bottomRight = worksheet.Range(topLeft).End(XlDirection.xlToRight).Address
                // bottomRight = worksheet.Range(bottomRight).End(XlDirection.xlDown).Address

                // selectedRange = worksheet.Range(worksheet.Range(topLeft), worksheet.Range(bottomRight))


                // End If

                // selectedRange.Select()




                // Try

                excelApp = Globals.ThisAddIn.Application;
                workbook = excelApp.ActiveWorkbook;
                worksheet = (Excel.Worksheet)workbook.ActiveSheet;
                selectedRange = (Range)excelApp.Selection;

                var activeRange = excelApp.ActiveCell;

                int startRow = activeRange.Row;
                int startColumn = activeRange.Column;
                int endRow = activeRange.Row;
                int endColumn = activeRange.Column;

                // Find the upper boundary
                while (startRow > 1 && !(worksheet.Cells[startRow - 1, startColumn].Value == null))
                    startRow -= 1;

                // Find the lower boundary
                while (!(worksheet.Cells[endRow + 1, endColumn].Value == null))
                    endRow += 1;

                // Find the left boundary
                while (startColumn > 1 && !(worksheet.Cells[startRow, startColumn - 1].Value == null))
                    startColumn -= 1;

                // Find the right boundary
                while (!(worksheet.Cells[endRow, endColumn + 1].Value == null))
                    endColumn += 1;

                // Select the determined range
                worksheet.get_Range(worksheet.Cells[startRow, startColumn], worksheet.Cells[endRow, endColumn]).Select();
            }

            // Catch ex As Exception

            // End Try



            catch (Exception ex)
            {

            }



        }

        public bool IsValidRng(string input)
        {

            // Dim pattern As String = "^(\$?[A-Z]+\$?[0-9]+(:\$?[A-Z]+\$?[0-9]+)?)(,\$?[A-Z]+\$?[0-9]+(:\$?[A-Z]+\$?[0-9]+)?)*$"
            string pattern = @"^((\$?[A-Z]+\$?[0-9]+(:\$?[A-Z]+\$?[0-9]+)?)|(\$?[A-Z]{1,2}:\$?[A-Z]{1,2})|(\$?[1-9][0-9]{0,6}:\$?[1-9][0-9]{0,6})|([A-Z]{1,2})|([1-9][0-9]{0,6}))(,((\$?[A-Z]+\$?[0-9]+(:\$?[A-Z]+\$?[0-9]+)?)|(\$?[A-Z]{1,2}:\$?[A-Z]{1,2})|(\$?[1-9][0-9]{0,6}:\$?[1-9][0-9]{0,6})|([A-Z]{1,2})|([1-9][0-9]{0,6})))*$";

            return Regex.IsMatch(input, pattern);

        }

        private void btn_OK_Click(object sender, EventArgs e)
        {
            try
            {

                string inputWsName;
                excelApp = Globals.ThisAddIn.Application;
                workbook = excelApp.ActiveWorkbook;
                worksheet = (Excel.Worksheet)workbook.ActiveSheet;
                inputWsName = worksheet.Name;

                if (string.IsNullOrEmpty(txtSourceRange.Text))
                {
                    Interaction.MsgBox("Please select the Source Range.", MsgBoxStyle.Exclamation, "Error!");
                    txtSourceRange.Focus();
                    return;
                }
                else if (IsValidRng(txtSourceRange.Text.ToUpper()) == false)
                {
                    Interaction.MsgBox("Please use a valid range.", MsgBoxStyle.Exclamation, "Error!");
                    txtSourceRange.Text = "";
                    txtSourceRange.Focus();
                    return;
                }


                int rngCount;
                rngCount = 0;
                foreach (char c in txtSourceRange.Text)
                {

                    if (Conversions.ToString(c) == ",")
                    {
                        rngCount = rngCount + 1;
                    }

                }

                IsEntireWsHidden();

                if (rngCount == 0 & RB_Single_Range.Checked == true)
                {
                    singleRng();
                    Dispose();
                }
                else if (rngCount == 0 & RB_Multiple_Range.Checked == true)
                {
                    Interaction.MsgBox("Please select correct Range Type.", MsgBoxStyle.Exclamation, "Error!");
                    RB_Single_Range.Focus();
                }
                else if (rngCount != 0 & RB_Multiple_Range.Checked == true)
                {
                    multiRng();
                    Dispose();
                }
                else if (rngCount != 0 & RB_Single_Range.Checked == true)
                {
                    Interaction.MsgBox("Please select correct Range Type.", MsgBoxStyle.Exclamation, "Error!");
                    RB_Multiple_Range.Focus();
                }
            }


            catch (Exception ex)
            {

            }
        }



        private void singleRng()
        {

            try
            {
                string inputWsName;
                excelApp = Globals.ThisAddIn.Application;
                workbook = excelApp.ActiveWorkbook;
                worksheet = (Excel.Worksheet)workbook.ActiveSheet;
                inputWsName = worksheet.Name;

                string temp;
                MsgBoxResult answer;
                temp = txtSourceRange.Text;
                worksheet1 = inputRng.Worksheet;




                if (CheckBox1.Checked == true)
                {

                    workbook.ActiveSheet.Copy(After: workbook.Sheets[workbook.Sheets.Count]);
                    outWorksheet = (Excel.Worksheet)workbook.Sheets[workbook.Sheets.Count];

                    worksheet1.Activate();
                    txtSourceRange.Text = temp;

                }

                int firstRow, lastRow, firstColumn, lastColumn;


                selectedRange = worksheet.get_Range(txtSourceRange.Text);
                firstRow = selectedRange.Row;
                lastRow = firstRow + selectedRange.Rows.Count - 1;
                firstColumn = selectedRange.Column;
                lastColumn = firstColumn + selectedRange.Columns.Count - 1;

                if (IsEntireWsHidden() == true)
                {
                    answer = Interaction.MsgBox("You are about to hide the entire worksheet." + Microsoft.VisualBasic.Constants.vbCrLf + "Do you want to proceed?", MsgBoxStyle.YesNo, "Warning!");
                    if (answer == MsgBoxResult.Yes)
                    {
                        goto Proceed2;
                    }
                    else
                    {
                        goto break2;
                    }
                }

                if (RB_Single_Range.Checked == true & RB_Row.Checked == true)
                {
                    if (selectedRange.Rows.Count <= 2)
                    {

                        answer = Interaction.MsgBox("You are about to hide " + selectedRange.Rows.Count + " Rows." + Microsoft.VisualBasic.Constants.vbCrLf + "Do you want to proceed?", MsgBoxStyle.YesNo, "Warning!");
                        if (answer == MsgBoxResult.Yes)
                        {
                            goto Proceed1;
                        }
                        else
                        {
                            goto break1;
                        }
                    }

Proceed1:
                    ;

                    worksheet.get_Range(worksheet.Cells[firstRow, firstColumn], worksheet.Cells[lastRow, lastColumn]).EntireRow.Hidden = (object)true;
break1:
                    ;

                    Dispose();
                }

                else if (RB_Single_Range.Checked == true & RB_Column.Checked == true)
                {
                    if (selectedRange.Columns.Count <= 2)
                    {
                        answer = Interaction.MsgBox("You are about to hide " + selectedRange.Columns.Count + " Columns." + Microsoft.VisualBasic.Constants.vbCrLf + "Do you want to proceed?", MsgBoxStyle.YesNo, "Warning!");
                        if (answer == MsgBoxResult.Yes)
                        {
                            goto Proceed2;
                        }
                        else
                        {
                            goto break2;
                        }
                    }

Proceed2:
                    ;

                    worksheet.get_Range(worksheet.Cells[firstRow, firstColumn], worksheet.Cells[lastRow, lastColumn]).EntireColumn.Hidden = (object)true;
break2:
                    ;

                    Dispose();
                }

                else if (RB_Single_Range.Checked == true & RB_bidirection.Checked == true)
                {
                    if (selectedRange.Columns.Count <= 2)
                    {
                        answer = Interaction.MsgBox("You are about to hide " + selectedRange.Rows.Count + " Rows and" + selectedRange.Columns.Count + " Columns." + Microsoft.VisualBasic.Constants.vbCrLf + "Do you want to proceed?", MsgBoxStyle.YesNo, "Warning!");
                        if (answer == MsgBoxResult.Yes)
                        {
                            goto Proceed3;
                        }
                        else
                        {
                            goto break3;
                        }
                    }

Proceed3:
                    ;

                    worksheet.get_Range(worksheet.Cells[firstRow, 1], worksheet.Cells[lastRow, 1]).EntireRow.Hidden = (object)true;
                    worksheet.get_Range(worksheet.Cells[1, firstColumn], worksheet.Cells[1, lastColumn]).EntireColumn.Hidden = (object)true;

break3:
                    ;

                    Dispose();
                }
            }



            catch (Exception ex)
            {

            }


        }

        private void multiRng()
        {

            try
            {

                string inputWsName;
                excelApp = Globals.ThisAddIn.Application;
                workbook = excelApp.ActiveWorkbook;
                worksheet = (Excel.Worksheet)workbook.ActiveSheet;
                inputWsName = worksheet.Name;

                string temp;
                MsgBoxResult answer;
                temp = txtSourceRange.Text;
                worksheet1 = inputRng.Worksheet;


                if (CheckBox1.Checked == true)
                {

                    workbook.ActiveSheet.Copy(After: workbook.Sheets[workbook.Sheets.Count]);
                    outWorksheet = (Excel.Worksheet)workbook.Sheets[workbook.Sheets.Count];

                    worksheet1.Activate();
                    txtSourceRange.Text = temp;

                }


                if (IsEntireWsHidden() == true)
                {
                    answer = Interaction.MsgBox("You are about to hide the entire worksheet." + Microsoft.VisualBasic.Constants.vbCrLf + "Do you want to proceed?", MsgBoxStyle.YesNo, "Warning!");
                    if (answer == MsgBoxResult.Yes)
                    {
                        goto proceed;
                    }
                    else
                    {
                        goto break;
                    }
                }

proceed:
                ;

                int visRows, followingRows;
                int visColumns, followingColumns;
                arrRng = Strings.Split(txtSourceRange.Text, ",");

                if (RB_Multiple_Range.Checked == true & RB_Row.Checked == true)
                {
                    for (int i = 0, loopTo = Information.UBound(arrRng); i <= loopTo; i++)
                    {
                        visRows = worksheet.get_Range(arrRng[i]).Row;
                        followingRows = visRows + worksheet.get_Range(arrRng[i]).Rows.Count - 1;
                        visColumns = worksheet.get_Range(arrRng[i]).Column;
                        followingColumns = visColumns + worksheet.get_Range(arrRng[i]).Columns.Count - 1;

                        worksheet.get_Range(worksheet.Cells[visRows, 1], worksheet.Cells[followingRows, 1]).EntireRow.Hidden = (object)true;

                    }
                }



                else if (RB_Multiple_Range.Checked == true & RB_Column.Checked == true)
                {
                    for (int i = 0, loopTo2 = Information.UBound(arrRng); i <= loopTo2; i++)
                    {
                        visRows = worksheet.get_Range(arrRng[i]).Row;
                        followingRows = visRows + worksheet.get_Range(arrRng[i]).Rows.Count - 1;
                        visColumns = worksheet.get_Range(arrRng[i]).Column;
                        followingColumns = visColumns + worksheet.get_Range(arrRng[i]).Columns.Count - 1;

                        worksheet.get_Range(worksheet.Cells[1, visColumns], worksheet.Cells[1, followingColumns]).EntireColumn.Hidden = (object)true;


                    }
                }


                else
                {
                    for (int i = 0, loopTo1 = Information.UBound(arrRng); i <= loopTo1; i++)
                    {
                        visRows = worksheet.get_Range(arrRng[i]).Row;
                        followingRows = visRows + worksheet.get_Range(arrRng[i]).Rows.Count - 1;
                        visColumns = worksheet.get_Range(arrRng[i]).Column;
                        followingColumns = visColumns + worksheet.get_Range(arrRng[i]).Columns.Count - 1;

                        worksheet.get_Range(worksheet.Cells[visRows, 1], worksheet.Cells[followingRows, 1]).EntireRow.Hidden = (object)true;
                        worksheet.get_Range(worksheet.Cells[1, visColumns], worksheet.Cells[1, followingColumns]).EntireColumn.Hidden = (object)true;


                    }


                }



break:
                ;

                Dispose();
            }


            catch (Exception ex)
            {

            }
        }

        private bool IsEntireWsHidden()
        {

            Range selectedRng = (Range)excelApp.Selection;
            bool flag = false;
            arrRng = Strings.Split(txtSourceRange.Text, ",");

            if (RB_Row.Checked == true)
            {
                if (selectedRng.get_Address(XlReferenceStyle.xlA1) == "$1:$1048576")
                {
                    flag = true;
                }

                for (int i = 0, loopTo = Information.UBound(arrRng); i <= loopTo; i++)
                {
                    if (Regex.IsMatch(arrRng[i].ToUpper(), @"^(\$?[A-Z]{1,3}):(\$?[A-Z]{1,3})$"))
                    {
                        flag = true;
                        break;
                    }
                }
            }

            else if (RB_Column.Checked == true)
            {
                if (selectedRng.get_Address(XlReferenceStyle.xlA1) == "$1:$1048576")
                {
                    flag = true;
                }

                for (int i = 0, loopTo1 = Information.UBound(arrRng); i <= loopTo1; i++)
                {
                    if (Regex.IsMatch(arrRng[i].ToUpper(), @"^(\$?[1-9][0-9]*):(\$?[1-9][0-9]*)$"))
                    {
                        flag = true;
                        break;
                    }
                }
            }

            else if (RB_bidirection.Checked == true)
            {
                if (selectedRng.get_Address(XlReferenceStyle.xlA1) == "$1:$1048576")
                {
                    flag = true;
                }

                for (int i = 0, loopTo2 = Information.UBound(arrRng); i <= loopTo2; i++)
                {
                    if (Regex.IsMatch(arrRng[i].ToUpper(), @"^(\$?[A-Z]{1,3}):(\$?[A-Z]{1,3})$") | Regex.IsMatch(arrRng[i].ToUpper(), @"^(\$?[1-9][0-9]*):(\$?[1-9][0-9]*)$"))
                    {
                        flag = true;
                        break;
                    }
                }

            }

            return flag;

        }

        private void btn_Cancel_Click(object sender, EventArgs e)
        {
            Dispose();
        }

        private void Form12HideRanges_Closing(object sender, CancelEventArgs e)
        {
            GlobalModule.form_flag = false;
        }

        private void Form12HideRanges_Disposed(object sender, EventArgs e)
        {
            GlobalModule.form_flag = false;
        }

        private void Form12HideRanges_Shown(object sender, EventArgs e)
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